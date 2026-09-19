#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメントに埋め込まれたXMPメタデータから使用フォント情報を抽出し、TXT / CSV / Markdownで書き出します。
XMPで足りない分はドキュメント内のテキストから補い、環境にないフォントだけに絞り込んで書き出すこともできます。

詳細は README を参照してください。

### Overview

Extracts font usage information from the XMP metadata embedded in the document and exports it as TXT / CSV / Markdown.
What the XMP lacks is topped up from the text in the document, and the export can be narrowed to the fonts this machine does not have.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExportFontInfoFromXMP";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExportFontInfoFromXMP.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExportFontInfoFromXMP.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n16e7e95652b6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 出力ファイル名のサフィックスと区切り線 / Output filename suffix and section divider */
    var FILENAME_SUFFIX = "_fontInfo";
    var SECTION_DIVIDER = "-----------------------------";

    /* ［見つからないフォントのみ］で書き出したときに足すサフィックス / Extra suffix for a "missing fonts only" export */
    var MISSING_FILENAME_SUFFIX = "_missing";

    /* ［書き出し後にフォルダーを開く］の初期値 / Default for "Open the folder after exporting" */
    var OPEN_FOLDER_DEFAULT = true;

    /* ［見つからないフォントのみ］の初期値 / Default for "Missing fonts only" */
    var MISSING_ONLY_DEFAULT = false;

    // =========================================
    // レイアウト / Layout
    // =========================================
    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /**
     * パネルに共通のレイアウトを適用します。
     *
     * @param {Panel} panel - 対象のパネル。
     * @param {number} [spacing] - 子要素の間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループに共通のレイアウトを適用します。
     *
     * @param {Group} group - 対象のグループ。
     * @param {string} [orientation] - "row" または "column"。省略時は "column"。
     * @param {number} [spacing] - 子要素の間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================
    /**
     * 現在のUI言語を判定します。
     *
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    /**
     * 現在のUI言語。
     *
     * @type {string}
     */
    var currentLanguage = getCurrentLang();

    /**
     * UIと出力に使う文言の定義。末端は { ja, en } の組。
     *
     * @type {object}
     */
    var LABELS = {
        dialog: {
            title: { ja: "フォント情報を書き出し", en: "Export Font Info" }
        },
        format: {
            title: { ja: "書き出し形式", en: "Export Format" },
            text: { ja: "テキストファイル（.txt）", en: "Text File (.txt)" },
            csv: { ja: "CSVファイル（.csv）", en: "CSV File (.csv)" },
            markdown: { ja: "Markdownファイル（.md）", en: "Markdown File (.md)" },
            all: { ja: "すべて（3種類書き出し）", en: "All Formats (TXT + CSV + MD)" }
        },
        destination: {
            title: { ja: "書き出し先", en: "Destination" },
            desktop: { ja: "デスクトップ", en: "Desktop" },
            sameFolder: { ja: "ファイルと同じ階層", en: "Same folder as the file" },
            openFolder: { ja: "書き出し後にフォルダーを開く", en: "Open the folder after exporting" }
        },
        tooltip: {
            formatText:     { ja: "タブ区切りのテキストファイルとして書き出します。", en: "Writes a tab-separated text file." },
            formatCsv:      { ja: "カンマ区切りのCSVとして書き出します。表計算ソフトで開けます。", en: "Writes a CSV file that opens in a spreadsheet." },
            formatMarkdown: { ja: "Markdown の表として書き出します。", en: "Writes a Markdown table." },
            formatAll:      { ja: "上の3つの形式をすべて書き出します。", en: "Writes all three formats." },
            destDesktop:    { ja: "デスクトップに保存します。", en: "Saves to the desktop." },
            destSameFolder: { ja: "ドキュメントと同じフォルダーに保存します。未保存のドキュメントでは使えません。", en: "Saves next to the document. Unavailable for unsaved documents." },
            openFolder:     { ja: "書き出したあと、保存先のフォルダーを開きます。", en: "Opens the destination folder once the files are written." },
            missingOnly:    { ja: "この環境に入っていないフォントだけを書き出します。", en: "Writes only the fonts that are missing from this machine." }
        },
        exportOptions: {
            title: { ja: "オプション", en: "Options" },
            missingOnly: { ja: "見つからないフォントのみ", en: "Missing fonts only" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noFonts: { ja: "フォント情報が見つかりませんでした。", en: "No font information found." },
            noMissingFonts: { ja: "見つからないフォントはありませんでした。", en: "No missing fonts found." },
            done: { ja: "書き出しました。", en: "Exported." },
            writeFailed: { ja: "ファイルを書き込めませんでした：\n", en: "Could not write the file:\n" },
            alreadyWritten: { ja: "書き出し済みのファイル：\n", en: "Already exported:\n" },
            error: { ja: "エラーが発生しました：\n", en: "An error occurred:\n" }
        },
        output: {
            bullet: { ja: "・", en: "- " },
            fontListHeading: { ja: "使用フォント一覧", en: "Font List" },
            fontCount: { ja: "使用フォント数", en: "Font Count" },
            missingFontListHeading: { ja: "見つからないフォント一覧", en: "Missing Font List" },
            missingFontCount: { ja: "見つからないフォント数", en: "Missing Font Count" },
            fontDetailHeading: { ja: "各フォントの情報", en: "Font Details" },
            compositeFonts: { ja: "構成フォント", en: "Composite Fonts" },
            compositeType: { ja: "合成フォント", en: "Composite Font" }
        }
    };

    /**
     * ドット区切りパスでローカライズ文字列を取得します。
     *
     * @param {string} path - "format.title" のようなドット区切りのキー。
     * @returns {string} 現在の言語の文字列。キーが見つからない場合は path をそのまま返す。
     */
    function getLabel(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var idx = 0; idx < parts.length; idx++) {
            node = node[parts[idx]];
            if (node === undefined) return path;
        }
        return (node[currentLanguage] !== undefined) ? node[currentLanguage] : node.en;
    }

    /**
     * コロンを付けたラベルを返します（日本語は全角、英語は半角）。
     *
     * @param {string} path - ドット区切りのラベルキー。
     * @returns {string} コロン付きのラベル。
     */
    function labelText(path) {
        return getLabel(path) + (currentLanguage === "ja" ? "：" : ":");
    }

    /**
     * 件数を括弧付きで添えたラベルを返します（日本語は全角括弧、英語は半角括弧）。
     *
     * @param {string} path - ドット区切りのラベルキー。
     * @param {number} count - 括弧内に表示する件数。
     * @returns {string} 件数付きのラベル。
     */
    function labelWithCount(path, count) {
        if (currentLanguage === "ja") return getLabel(path) + "（" + count + "）";
        return getLabel(path) + " (" + count + ")";
    }

    // =========================================
    // 書き出し形式 / Export Formats
    // =========================================
    /**
     * 書き出し形式ごとの設定。
     *
     * @typedef {object} FormatSpec
     * @property {string} encoding - ファイルのエンコーディング。
     * @property {string} bom - 先頭に書き込むBOM。不要な形式では空文字。
     * @property {string} newline - 行の区切り文字。
     * @property {function} build - 行を生成する関数。(FontEntry[], string, boolean) => string[]
     */

    /**
     * 形式キー（"txt" / "csv" / "md"）から設定を引くテーブル。
     *
     * @type {object}
     */
    var FORMAT_SPECS = {
        txt: { encoding: "UTF-8", bom: "", newline: "\n", build: buildTxtLines },
        /* CSVはExcel向けにBOM付きUTF-16、改行はCRLF / CSV: UTF-16 with BOM and CRLF for Excel */
        csv: { encoding: "UTF-16", bom: "\uFEFF", newline: "\r\n", build: buildCsvLines },
        md: { encoding: "UTF-8", bom: "", newline: "\n", build: buildMarkdownLines }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var dialogResult = showFormatDialog();
    if (!dialogResult) return;

    exportFontInfo(dialogResult);

    /**
     * ドキュメントの保存先フォルダーを返します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {Folder|null} 保存先フォルダー。未保存なら null。
     */
    function getDocumentFolder(doc) {
        /* path は未保存でも実在するフォルダーを返しうるので、ファイル自体の有無で判定する
           / path can point at an existing folder even when unsaved, so test the file itself */
        try {
            var file = doc.fullName;
            return (file && file.exists) ? file.parent : null;
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // UI ヘルパー / UI Helpers
    // =========================================
    /**
     * 指定インデックスのラジオボタンだけを選択し、フォーカスも移します。
     *
     * @param {Array<RadioButton>} radios - 同じグループのラジオボタン。
     * @param {number} index - 選択するラジオボタンのインデックス。
     * @returns {void}
     */
    function selectRadio(radios, index) {
        for (var idx = 0; idx < radios.length; idx++) {
            radios[idx].value = (idx === index);
        }
        /* フォーカスを移さないとキーイベントが元のボタンから発火し続ける / Without moving focus, key events keep firing from the original button */
        radios[index].active = true;
    }

    /**
     * 指定方向で次に選べるラジオボタンのインデックスを返します。
     *
     * @param {Array<RadioButton>} radios - 同じグループのラジオボタン。
     * @param {number} index - 現在のインデックス。
     * @param {number} step - -1 で前、1 で次。
     * @returns {number} 次に選べるインデックス。無ければ index。
     */
    function nextEnabledRadioIndex(radios, index, step) {
        var candidate = index;
        for (var i = 0; i < radios.length; i++) {
            candidate = (candidate + step + radios.length) % radios.length;
            if (radios[candidate].enabled) return candidate;
        }
        return index;
    }

    /**
     * 上下キーでラジオボタンを循環移動できるようにします。
     *
     * @param {Array<RadioButton>} radios - 同じグループのラジオボタン。
     * @returns {void}
     */
    function enableArrowKeyNavigation(radios) {
        for (var i = 0; i < radios.length; i++) {
            (function (index) {
                radios[index].addEventListener("keydown", function (event) {
                    var key = event.keyName;
                    /* 無効なラジオは飛ばす（未保存時の［ファイルと同じ階層］など） / Skip disabled radios, e.g. "Same folder" while unsaved */
                    if (key === "Up" || key === "ArrowUp") {
                        selectRadio(radios, nextEnabledRadioIndex(radios, index, -1));
                    } else if (key === "Down" || key === "ArrowDown") {
                        selectRadio(radios, nextEnabledRadioIndex(radios, index, 1));
                    }
                });
            })(i);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /**
     * ダイアログで選択された書き出し条件。
     *
     * @typedef {object} ExportOptions
     * @property {string[]} formats - 書き出す形式（"txt" / "csv" / "md"）の配列。
     * @property {string} destination - "desktop" または "sameFolder"。
     * @property {boolean} openFolder - 書き出し後にフォルダーを開く場合は true。
     * @property {boolean} missingOnly - 見つからないフォントだけを書き出す場合は true。
     */

    /**
     * 書き出し形式と書き出し先を選ぶダイアログを表示します。
     *
     * @returns {ExportOptions|null} 選択された条件。キャンセル時は null。
     */
    function showFormatDialog() {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["left", "top"];
        dialog.spacing = 10;
        dialog.margins = 20;

        /* 書き出し形式パネル / Export format panel */
        var formatPanel = dialog.add("panel", undefined, getLabel("format.title"));
        setupPanel(formatPanel, 6);

        var radioText = formatPanel.add("radiobutton", undefined, getLabel("format.text"));
        radioText.helpTip = getLabel("tooltip.formatText");
        var radioCsv = formatPanel.add("radiobutton", undefined, getLabel("format.csv"));
        radioCsv.helpTip = getLabel("tooltip.formatCsv");
        var radioMarkdown = formatPanel.add("radiobutton", undefined, getLabel("format.markdown"));
        radioMarkdown.helpTip = getLabel("tooltip.formatMarkdown");
        var radioAll = formatPanel.add("radiobutton", undefined, getLabel("format.all"));
        radioAll.helpTip = getLabel("tooltip.formatAll");
        radioText.value = true;
        radioText.active = true;
        enableArrowKeyNavigation([radioText, radioCsv, radioMarkdown, radioAll]);

        /* 書き出し先パネル / Destination panel */
        var destinationPanel = dialog.add("panel", undefined, getLabel("destination.title"));
        setupPanel(destinationPanel, 6);

        var radioDesktop = destinationPanel.add("radiobutton", undefined, getLabel("destination.desktop"));
        radioDesktop.helpTip = getLabel("tooltip.destDesktop");
        var radioSameFolder = destinationPanel.add("radiobutton", undefined, getLabel("destination.sameFolder"));
        radioSameFolder.helpTip = getLabel("tooltip.destSameFolder");
        radioDesktop.value = true;
        /* 未保存のドキュメントには同じ階層が無いのでデスクトップだけにする / A never-saved document has no folder of its own, so leave only Desktop */
        radioSameFolder.enabled = !!getDocumentFolder(app.activeDocument);
        enableArrowKeyNavigation([radioDesktop, radioSameFolder]);

        /* ラジオボタンと同じ間隔だと選択肢に見えるため、少し離す / Nudge it down so it doesn't read as a third radio option */
        var openFolderGroup = destinationPanel.add("group");
        setupGroup(openFolderGroup, "row");
        openFolderGroup.margins = [0, 4, 0, 0];

        var openFolderCheckbox = openFolderGroup.add("checkbox", undefined, getLabel("destination.openFolder"));
        openFolderCheckbox.helpTip = getLabel("tooltip.openFolder");
        openFolderCheckbox.value = OPEN_FOLDER_DEFAULT;

        /* オプションパネル / Options panel */
        var exportOptionsPanel = dialog.add("panel", undefined, getLabel("exportOptions.title"));
        setupPanel(exportOptionsPanel, 6);

        var missingOnlyCheckbox = exportOptionsPanel.add("checkbox", undefined, getLabel("exportOptions.missingOnly"));
        missingOnlyCheckbox.helpTip = getLabel("tooltip.missingOnly");
        missingOnlyCheckbox.value = MISSING_ONLY_DEFAULT;

        var buttonGroup = dialog.add("group");
        setupGroup(buttonGroup, "row");
        /* alignment と alignChildren を対で指定しないとボタンが伸びて天地がずれる / Set alignment and alignChildren together, or the buttons stretch and drift vertically */
        buttonGroup.alignment = ["center", "top"];
        buttonGroup.alignChildren = ["center", "center"];
        buttonGroup.margins = [0, 5, 0, 0]; // ボタンエリア上に余白 +5 / Extra top margin above buttons
        buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, "OK", { name: "ok" });

        if (dialog.show() !== 1) return null;

        var formats;
        if (radioAll.value) formats = ["txt", "csv", "md"];
        else if (radioCsv.value) formats = ["csv"];
        else if (radioMarkdown.value) formats = ["md"];
        else formats = ["txt"];

        return {
            formats: formats,
            destination: radioSameFolder.value ? "sameFolder" : "desktop",
            openFolder: openFolderCheckbox.value,
            missingOnly: missingOnlyCheckbox.value
        };
    }

    // =========================================
    // 書き出し / Export
    // =========================================
    /**
     * フォント情報を集め、指定された各形式でファイルを書き出します。
     *
     * @param {ExportOptions} exportOptions - ダイアログで選択された書き出し条件。
     * @returns {void}
     */
    function exportFontInfo(exportOptions) {
        /* 途中で失敗しても書けた分を知らせるため、catch から見える位置で宣言する
           / Declared where the catch can see it, so a failure can still report what was written */
        var writtenNames = [];

        try {
            var doc = app.activeDocument;
            var entries = collectFontEntries(doc, exportOptions.missingOnly);
            if (entries.length === 0) {
                alert(getLabel("alert.noFonts"));
                return;
            }

            if (exportOptions.missingOnly) {
                entries = filterMissingFonts(entries);
                if (entries.length === 0) {
                    alert(getLabel("alert.noMissingFonts"));
                    return;
                }
            }

            /* デスクトップ、またはドキュメントと同じフォルダー / Desktop, or the document's own folder */
            var docFolder = getDocumentFolder(doc);
            var outputFolder = (exportOptions.destination === "sameFolder" && docFolder) ? docFolder : Folder.desktop;
            writeFontInfoFiles(entries, exportOptions, outputFolder, doc.name, writtenNames);

            if (exportOptions.openFolder) {
                outputFolder.execute();
            } else {
                /* フォルダーを開かない場合だけ、書き出し結果をアラートで知らせる / Report the result only when the folder is not opened */
                alert(getLabel("alert.done") + "\n" + writtenNames.join("\n"));
            }

        } catch (e) {
            alert(exportErrorMessage(e, writtenNames));
        }
    }

    /**
     * 書き出し中の例外から、アラートに表示する文面を組み立てます。
     *
     * @param {Error} e - 捕捉した例外。
     * @param {string[]} writtenNames - そこまでに書き出せたファイル名の配列。
     * @returns {string} アラートに表示する文面。
     */
    function exportErrorMessage(e, writtenNames) {
        var message = e.message ? e.message : String(e);

        /* 書き込み失敗は自前の見出しを持つので、汎用のエラー見出しを重ねない
           / A write failure carries its own heading, so don't stack the generic one on top of it */
        if (!e.isWriteFailure) {
            message = getLabel("alert.error") + message + (e.line ? "\n(line " + e.line + ")" : "");
        }

        /* 3種類書き出しでは一部だけ書けていることがあるので、その分も知らせる
           / A multi-format export may have written part of its files, so report those too */
        if (writtenNames.length > 0) {
            message += "\n\n" + getLabel("alert.alreadyWritten") + writtenNames.join("\n");
        }

        return message;
    }

    /**
     * 指定された各形式でファイルを書き出します。
     *
     * @param {FontEntry[]} entries - 抽出済みのフォント情報。
     * @param {ExportOptions} exportOptions - ダイアログで選択された書き出し条件。
     * @param {Folder} outputFolder - 書き出し先フォルダー。
     * @param {string} docName - 拡張子を含むドキュメント名。
     * @param {string[]} writtenNames - 書き出せたファイル名を追記する配列。
     * @returns {void}
     */
    function writeFontInfoFiles(entries, exportOptions, outputFolder, docName, writtenNames) {
        var formats = exportOptions.formats;
        /* 絞り込んだ結果を全件の一覧と取り違えないよう、ファイル名を分ける
           / Give the filtered export its own name so it can't be mistaken for a full font list */
        var baseName = docName.replace(/\.[^\.]+$/, "") + FILENAME_SUFFIX +
            (exportOptions.missingOnly ? MISSING_FILENAME_SUFFIX : "");

        for (var i = 0; i < formats.length; i++) {
            var spec = FORMAT_SPECS[formats[i]];
            var file = uniqueOutputFile(outputFolder, baseName, "." + formats[i]);
            writeLines(file, spec, spec.build(entries, docName, exportOptions.missingOnly));
            writtenNames.push(decodeURI(file.name));
        }
    }

    /**
     * 指定フォルダー内で重複しないファイルオブジェクトを返します。
     *
     * @param {Folder} folder - 書き出し先フォルダー。
     * @param {string} baseName - 拡張子を除いたファイル名。
     * @param {string} ext - ドットを含む拡張子（".csv" など）。
     * @returns {File} 既存ファイルと重複しない File。
     */
    function uniqueOutputFile(folder, baseName, ext) {
        var file = new File(folder + "/" + baseName + ext);
        var counter = 2;
        while (file.exists) {
            file = new File(folder + "/" + baseName + "_" + counter + ext);
            counter++;
        }
        return file;
    }

    /**
     * 形式ごとの設定にしたがって行を書き込みます。
     *
     * @param {File} file - 書き出し先ファイル。
     * @param {FormatSpec} spec - 形式ごとの設定。
     * @param {string[]} lines - 書き込む行の配列。
     * @returns {void}
     */
    function writeLines(file, spec, lines) {
        file.encoding = spec.encoding;
        /* 権限が無い場合などは無音で失敗するため、必ず戻り値を確認 / open() fails silently (e.g. no write permission), so always check it */
        if (!file.open("w")) {
            var writeError = new Error(getLabel("alert.writeFailed") + decodeURI(file.fsName));
            /* 見出し付きのメッセージであることの目印 / Marks the message as already carrying its own heading */
            writeError.isWriteFailure = true;
            throw writeError;
        }
        file.write(spec.bom + lines.join(spec.newline));
        file.close();
    }

    // =========================================
    // フォント情報の抽出 / Font Data Extraction
    // =========================================
    /**
     * フォント1件分の情報。
     *
     * @typedef {object} FontEntry
     * @property {FontFields} fields - フォント属性。
     * @property {string[]} members - 構成フォント名の配列。合成フォントでなければ空配列。
     */

    /**
     * ドキュメントからフォント情報を集めます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {boolean} missingOnly - 見つからないフォントだけを書き出す場合は true。
     * @returns {FontEntry[]} フォント情報の配列。1件も無ければ空配列。
     */
    function collectFontEntries(doc, missingOnly) {
        var xmpEntries = extractFontEntriesFromXMP(readXMPString(doc));

        /* 保存済みならXMPが現状と一致する。未保存・編集中のXMPは前回保存時のままなので、
           テキストから拾った分を足して補う（テキスト走査は重いのでここでしか走らせない）
           / A saved document's XMP matches what's on screen; an unsaved or edited one still holds the
             last-saved XMP, so top it up from the text (the costly scan runs only in that case).
           保存済みでもXMPにフォント情報が無いことがある（旧バージョンでの保存、メタデータの除去）ので、
           そのときもテキストから拾う
           / A saved document can still carry no font block at all (saved by an older Illustrator,
             or stripped by a preflight tool), so fall back to the text scan there as well.
           見つからないフォントを探すときは、保存済みでも必ず走査する。プレビュー表示だけの
           埋め込みフォントは、XMPではなくDOMのファミリー名にしか印が残らないため
           / When hunting for missing fonts, scan even a saved document: a face shown only as an
             embedded preview is marked in the DOM's family name and nowhere in the XMP */
        if (doc.saved && xmpEntries.length > 0 && !missingOnly) return xmpEntries;

        return mergeFontEntries(xmpEntries, extractFontEntriesFromText(doc));
    }

    /**
     * XMP由来のフォント情報に、そこに無いテキスト由来の分だけを足します。
     *
     * @param {FontEntry[]} xmpEntries - XMPから抽出したフォント情報。
     * @param {FontEntry[]} textEntries - テキストから集めたフォント情報。
     * @returns {FontEntry[]} マージ後の配列。XMP由来を先に並べる。
     */
    function mergeFontEntries(xmpEntries, textEntries) {
        var seen = {};
        var merged = [];
        var i;

        /* XMP由来は version やフォントファイル名まで揃っているので、重複時はこちらを残す
           / XMP entries carry version and file name, so they win on a duplicate */
        for (i = 0; i < xmpEntries.length; i++) {
            markFontEntrySeen(seen, xmpEntries[i]);
            merged.push(xmpEntries[i]);
        }

        for (i = 0; i < textEntries.length; i++) {
            var existing = seenFontEntry(seen, textEntries[i]);
            if (existing) {
                /* 埋め込みの印はDOMにしか出ないので、残すXMP側のエントリへ引き継ぐ
                   / The embedded marker shows up only in the DOM, so carry it over to the XMP entry we keep */
                if (textEntries[i].fields.isEmbedded) existing.fields.isEmbedded = true;
                continue;
            }
            markFontEntrySeen(seen, textEntries[i]);
            merged.push(textEntries[i]);
        }

        return merged;
    }

    /**
     * フォント1件を「既出」として記録します。
     *
     * @param {object} seen - 既出フォントの索引。
     * @param {FontEntry} entry - 記録するフォント。
     * @returns {void}
     */
    function markFontEntrySeen(seen, entry) {
        var keys = fontEntryKeys(entry);
        for (var i = 0; i < keys.length; i++) {
            seen[keys[i]] = entry;
        }
    }

    /**
     * 同じフォントとして既出のエントリを返します。
     *
     * PostScript名はフォントを一意に決めるので、あるときはそれだけで判定します。
     * 「ファミリー フェイス」は別フォントどうしでも一致しうるため、PostScript名が無いときの代用にとどめます。
     *
     * @param {object} seen - 既出フォントの索引。
     * @param {FontEntry} entry - 判定するフォント。
     * @returns {FontEntry|null} 既出ならそのエントリ。無ければ null。
     */
    function seenFontEntry(seen, entry) {
        if (entry.fields.name) return seen["postScript:" + entry.fields.name] || null;

        var displayName = fontDisplayName(entry.fields);
        return displayName ? (seen["display:" + displayName] || null) : null;
    }

    /**
     * 重複判定に使うキーを返します。
     *
     * XMPとDOMで綴りが揃わないことがあるため、PostScript名と表示名の両方を登録します。
     *
     * @param {FontEntry} entry - 対象のフォント。
     * @returns {string[]} 照合に使うキーの配列。
     */
    function fontEntryKeys(entry) {
        var keys = [];
        var displayName = fontDisplayName(entry.fields);
        /* Object のプロパティ名と衝突させないため接頭辞を付ける / Prefix the keys so font names can't collide with Object members */
        if (entry.fields.name) keys.push("postScript:" + entry.fields.name);
        if (displayName) keys.push("display:" + displayName);
        return keys;
    }

    /**
     * ドキュメントのXMP文字列を取得します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {string} XMP文字列。取得できない場合は空文字。
     */
    function readXMPString(doc) {
        /* 文字列として読むだけなので AdobeXMPScript の読み込みは要らない
           / Read as a plain string, so there's no need to load AdobeXMPScript */
        try {
            return doc.XMPString || "";
        } catch (e) {
            return "";
        }
    }

    /**
     * XMP文字列から主フォントと構成フォントを抽出します。
     *
     * @param {string} xmpString - ドキュメントのXMP文字列。
     * @returns {FontEntry[]} 抽出結果。フォント情報が無い場合は空配列。
     */
    function extractFontEntriesFromXMP(xmpString) {
        var fontsMatch = xmpString.match(/<xmpTPg:Fonts>[\s\S]*?<\/xmpTPg:Fonts>/);
        if (!fontsMatch) return [];

        /* 構成フォントは主フォントの rdf:li に入れ子で並ぶため、主フォントの開始タグで分割する
           / Composite members are nested inside the primary rdf:li, so split on the primary's start tag */
        var chunks = fontsMatch[0].split(/<rdf:li[^>]*rdf:parseType="Resource"[^>]*>/);

        var entries = [];

        /* chunks[0] は最初の主フォントより前の部分なので読み飛ばす / chunks[0] precedes the first primary font */
        for (var i = 1; i < chunks.length; i++) {
            entries.push({
                fields: fontFields(chunks[i]),
                members: extractCompositeMembers(chunks[i])
            });
        }

        return entries;
    }

    /**
     * ドキュメント内のテキストから使用フォントを集めます（XMPを補うための走査）。
     *
     * バージョンやフォントファイル名はDOMから取得できないため空になります。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {FontEntry[]} フォント情報の配列。テキストが無ければ空配列。
     */
    function extractFontEntriesFromText(doc) {
        var seen = {};
        var entries = [];
        var frames = doc.textFrames;
        var frameCount = frames.length;

        for (var i = 0; i < frameCount; i++) {
            var chars;
            /* 壊れたストーリーやプラグイン所有のテキストでは参照自体が失敗するので、そのフレームだけ飛ばす
               / A broken story or plugin-owned text throws on access, so skip just that frame */
            try {
                /* フォントは文字単位で変わるため、1文字ずつ見るしかない（範囲単位では混在時に先頭の書式しか返らない）
                   / Fonts change per character, and a mixed range reports only its first character's font */
                chars = frames[i].textRange.characters;
            } catch (eFrame) {
                continue;
            }

            /* 長さはDOM参照なので、ループ条件で毎回読み直さない / chars.length hits the DOM, so read it once */
            var charCount = chars.length;

            for (var j = 0; j < charCount; j++) {
                var font;
                try {
                    font = chars[j].characterAttributes.textFont;
                } catch (e) {
                    continue;
                }
                if (!font) continue;

                /* キーの体系は mergeFontEntries と揃える（DOM由来は必ずPostScript名を持つ）
                   / Same key scheme as mergeFontEntries; a DOM font always carries a PostScript name */
                if (seen["postScript:" + font.name]) continue;

                var entry = {
                    fields: {
                        name: font.name,
                        family: font.family,
                        face: font.style,
                        type: "",
                        version: "",
                        fileName: "",
                        isComposite: false,
                        isEmbedded: isEmbeddedSubsetFamily(font.family)
                    },
                    members: []
                };
                markFontEntrySeen(seen, entry);
                entries.push(entry);
            }
        }

        return entries;
    }

    /**
     * 主フォントのXML断片から、構成フォント名（合成フォントの中身）を取り出します。
     *
     * @param {string} entryXml - 主フォント1件分のXML断片。
     * @returns {string[]} 構成フォント名の配列。合成フォントでなければ空配列。
     */
    function extractCompositeMembers(entryXml) {
        var memberEntries = entryXml.match(/<rdf:li[^>]*>[\s\S]*?<\/rdf:li>/g);
        if (!memberEntries) return [];

        var members = [];
        for (var i = 0; i < memberEntries.length; i++) {
            /* タグを除いたプレーンな構成フォント名 / Plain composite member name with tags stripped */
            members.push(decodeXmlEntities(memberEntries[i].replace(/<[^>]+>/g, "")));
        }
        return members;
    }

    // =========================================
    // 見つからないフォントの判定 / Missing Font Detection
    // =========================================
    /**
     * インストール済みフォント名の索引。初回参照時に作る。
     *
     * 代入付きの宣言は巻き上げでメイン処理より後に走り、初期化子が死ぬので付けない。
     *
     * @type {object|undefined}
     */
    var installedFontIndex;

    /**
     * インストール済みフォントの名前を引ける索引を返します。
     *
     * @returns {object} 名前をキーにした索引。
     */
    function getInstalledFontIndex() {
        if (installedFontIndex) return installedFontIndex;

        installedFontIndex = {};
        var fonts = app.textFonts;
        /* 長さはDOM参照なので、ループ条件で毎回読み直さない / fonts.length hits the DOM, so read it once */
        var fontCount = fonts.length;

        for (var i = 0; i < fontCount; i++) {
            var font = fonts[i];
            /* 環境にないフォントの仮エントリは導入済みとして数えない / A placeholder is not an installed font */
            if (isPlaceholderFont(font)) continue;

            /* XMPの値がPostScript名・表示名のどちらで入っていても引けるようにする。
               ファミリー名だけのキーは作らない（未導入のフェイスまで導入済みと判定してしまう）
               / Index the PostScript and display spellings, since XMP may carry either.
                 Never index the bare family: it would pass off an absent face as installed */
            installedFontIndex["font:" + font.name] = true;
            installedFontIndex["font:" + trimSpaces(font.family + " " + font.style)] = true;
        }
        return installedFontIndex;
    }

    /**
     * 環境にないフォントの代わりにIllustratorが置く仮エントリかを判定します。
     *
     * 環境にないフォントも app.textFonts に並ぶため、名前の照合だけでは導入済みと区別できません。
     * 仮エントリはフェイス情報を持たないため style が空になり、ファミリー名にPostScript名がそのまま入ります。
     * 合成フォントも style は空ですが、ファミリー名は合成フォント名なので当たりません。
     *
     * @param {TextFont} font - app.textFonts の1件。
     * @returns {boolean} 仮エントリなら true。
     */
    function isPlaceholderFont(font) {
        var name, family, style;
        try {
            name = font.name;
            family = font.family;
            style = font.style;
        } catch (e) {
            /* 読めないものは判定できないので、導入済みのまま扱う / Unreadable entries stay indexed */
            return false;
        }
        return style === "" && family === name;
    }

    /**
     * ドキュメントに埋め込まれたフォントのファミリー名かを判定します。
     *
     * 環境にないフォントをプレビューだけ表示している場合、ファミリー名に
     * "YAYCIH+" のようなサブセット接頭辞（英大文字6文字＋"+"）が付きます。
     * 同名のフォントが app.textFonts にあっても、実際に使われているのは埋め込みの方です。
     *
     * @param {string} family - ファミリー名。
     * @returns {boolean} 埋め込み由来なら true。
     */
    function isEmbeddedSubsetFamily(family) {
        return !!family && /^[A-Z]{6}\+/.test(family);
    }

    /**
     * 指定名のフォントがインストールされているかを判定します。
     *
     * @param {string} fontName - フォント名。
     * @returns {boolean} インストール済み、または名前が空なら true。
     */
    function isFontInstalled(fontName) {
        /* 名前が無いものは判定できないので、見つからない扱いにはしない / An empty name can't be judged, so don't call it missing */
        if (!fontName) return true;
        return getInstalledFontIndex()["font:" + fontName] === true;
    }

    /**
     * フォントファイル名から拡張子を取り除きます。
     *
     * 合成フォントの構成フォントは "RyoGothicStd-Bold.otf" のようにファイル名で記録されるため、
     * 拡張子を落とすとPostScript名や「ファミリー フェイス」と照合できます。
     *
     * @param {string} memberName - 構成フォント名。
     * @returns {string} 拡張子を除いた名前。既知の拡張子でなければそのまま返す。
     */
    function stripFontFileExtension(memberName) {
        if (!memberName) return "";
        /* フォント名自体にドットが入ることがあるので、既知のフォント拡張子だけを落とす
           / A font name can contain dots, so strip only the known font extensions */
        return memberName.replace(/\.(otf|ttf|ttc|otc|dfont|pfb|pfm|suit)$/i, "");
    }

    /**
     * フォント1件が見つからないフォントかを判定します。
     *
     * @param {FontEntry} entry - 判定対象のフォント。
     * @returns {boolean} 見つからないフォントなら true。
     */
    function isFontMissing(entry) {
        var f = entry.fields;

        /* 埋め込みのサブセットで表示しているものは、この環境には入っていない
           / A face shown from an embedded subset is not installed on this machine */
        if (f.isEmbedded) return true;

        /* 合成フォントは app.textFonts に "ATC-<合成フォント名のhex>" として並ぶ（style は空、ファミリー名は合成フォント名）。
           本体が無ければその時点で欠落、あれば構成フォントの側で判定する。
           構成フォントが取れないときは判定できないため、欠落とは見なさない
           / A composite is listed as "ATC-<hex of its name>" with an empty style and the composite's name as its family.
             Missing composite: report it outright; otherwise judge by the members.
             With no members to judge by, treat it as undetermined rather than missing */
        if (f.isComposite) {
            if (!isFontInstalled(f.name)) return true;

            for (var i = 0; i < entry.members.length; i++) {
                /* 構成フォントはファイル名で記録されるので、拡張子を落としてから照合する
                   / Members are recorded as file names, so drop the extension before the lookup */
                if (!isFontInstalled(stripFontFileExtension(entry.members[i]))) return true;
            }
            return false;
        }

        return !isFontInstalled(f.name);
    }

    /**
     * 見つからないフォントだけを抜き出します。
     *
     * @param {FontEntry[]} entries - フォント情報の配列。
     * @returns {FontEntry[]} 見つからないフォントだけの配列。
     */
    function filterMissingFonts(entries) {
        var missing = [];
        for (var i = 0; i < entries.length; i++) {
            if (isFontMissing(entries[i])) missing.push(entries[i]);
        }
        return missing;
    }

    // =========================================
    // 行の生成 / Line Builders
    // =========================================
    /**
     * 一覧見出しに使うラベルキーを返します。
     *
     * @param {boolean} missingOnly - 見つからないフォントだけの一覧なら true。
     * @returns {string} ラベルキー。
     */
    function fontListHeadingKey(missingOnly) {
        return missingOnly ? "output.missingFontListHeading" : "output.fontListHeading";
    }

    /**
     * フォント数に使うラベルキーを返します。
     *
     * @param {boolean} missingOnly - 見つからないフォントだけの一覧なら true。
     * @returns {string} ラベルキー。
     */
    function fontCountKey(missingOnly) {
        return missingOnly ? "output.missingFontCount" : "output.fontCount";
    }

    /**
     * テキスト形式（タブ区切り）の行を生成します。
     *
     * @param {FontEntry[]} entries - 抽出済みのフォント情報。
     * @param {string} fileName - 見出しに使うドキュメント名。
     * @param {boolean} missingOnly - 見つからないフォントだけの一覧なら true。
     * @returns {string[]} 書き込む行の配列。
     */
    function buildTxtLines(entries, fileName, missingOnly) {
        var bullet = getLabel("output.bullet");
        var lines = [];

        lines.push(labelText(fontListHeadingKey(missingOnly)) + " " + fileName + "\n");
        lines.push(labelWithCount(fontCountKey(missingOnly), entries.length) + "\n");

        for (var i = 0; i < entries.length; i++) {
            lines.push(bullet + fontDisplayName(entries[i].fields));
        }

        lines.push("\n" + SECTION_DIVIDER);

        for (i = 0; i < entries.length; i++) {
            lines = lines.concat(txtFontBlock(entries[i]));
        }

        return lines;
    }

    /**
     * テキスト形式の、フォント1件分の行を生成します。
     *
     * @param {FontEntry} entry - 出力するフォント。
     * @returns {string[]} フォント1件分の行。
     */
    function txtFontBlock(entry) {
        var f = entry.fields;
        var pairs = fontDetailPairs(f, true);
        var lines = [];

        for (var i = 0; i < pairs.length; i++) {
            lines.push(pairs[i][0] + ":\t" + pairs[i][1]);
        }

        if (hasCompositeMembers(f, entry.members)) {
            var bullet = getLabel("output.bullet");
            lines.push(labelText("output.compositeFonts"));
            for (var j = 0; j < entry.members.length; j++) {
                lines.push(bullet + entry.members[j]);
            }
        }

        lines.push(SECTION_DIVIDER);
        return lines;
    }

    /**
     * Markdown形式の行を生成します。
     *
     * @param {FontEntry[]} entries - 抽出済みのフォント情報。
     * @param {string} fileName - 見出しに使うドキュメント名。
     * @param {boolean} missingOnly - 見つからないフォントだけの一覧なら true。
     * @returns {string[]} 書き込む行の配列。
     */
    function buildMarkdownLines(entries, fileName, missingOnly) {
        var headingKey = fontListHeadingKey(missingOnly);
        var lines = [];

        lines.push("# " + getLabel(headingKey) + " " + escapeMarkdown(fileName) + "\n");
        // lines.push("[TOC]");
        lines.push("");
        lines.push("## " + getLabel(headingKey) + "\n");
        lines.push(labelWithCount(fontCountKey(missingOnly), entries.length) + "\n");

        for (var i = 0; i < entries.length; i++) {
            lines.push("- " + escapeMarkdown(fontDisplayName(entries[i].fields)));
        }

        lines.push("\n## " + getLabel("output.fontDetailHeading") + "\n");

        for (i = 0; i < entries.length; i++) {
            lines = lines.concat(markdownFontBlock(entries[i]));
        }

        return lines;
    }

    /**
     * Markdown形式の、フォント1件分の行を生成します。
     *
     * @param {FontEntry} entry - 出力するフォント。
     * @returns {string[]} フォント1件分の行。
     */
    function markdownFontBlock(entry) {
        var f = entry.fields;
        /* ファミリー名は見出しに出るので詳細では省く / The family is already in the heading, so skip it in the details */
        var pairs = fontDetailPairs(f, false);
        var lines = [];

        /* 同じファミリーで複数フェイスがあると見出しが重複するため、フェイスまで含める / Include the face; family alone duplicates headings across faces */
        lines.push("### " + escapeMarkdown(fontDisplayName(f)));
        lines.push("");

        for (var i = 0; i < pairs.length; i++) {
            lines.push("- " + pairs[i][0] + ": " + escapeMarkdown(pairs[i][1]));
        }

        if (hasCompositeMembers(f, entry.members)) {
            lines.push("");
            lines.push("#### " + getLabel("output.compositeFonts"));
            lines.push("");
            for (var j = 0; j < entry.members.length; j++) {
                lines.push("- " + escapeMarkdown(entry.members[j]));
            }
            lines.push("");
        }

        lines.push("");
        return lines;
    }

    /**
     * CSV形式の行を生成します。
     *
     * @param {FontEntry[]} entries - 抽出済みのフォント情報。
     * @returns {string[]} ヘッダー行を含む、書き込む行の配列。
     */
    function buildCsvLines(entries) {
        var lines = ["fontName,fontFamily,fontFace,fontType,version,fileName"];

        for (var i = 0; i < entries.length; i++) {
            var f = entries[i].fields;
            lines.push([
                escapeCsv(f.name),
                escapeCsv(f.family),
                escapeCsv(f.face),
                escapeCsv(f.type),
                escapeCsv(f.version),
                escapeCsv(f.fileName)
            ].join(","));
        }

        return lines;
    }

    /**
     * 詳細セクションに並べる「項目名・値」の組を返します。
     *
     * @param {FontFields} f - フォント属性。
     * @param {boolean} includeFamily - fontFamily を含める場合は true。
     * @returns {Array<string[]>} [項目名, 値] の配列。
     */
    function fontDetailPairs(f, includeFamily) {
        var pairs = [["fontName", f.name]];
        if (includeFamily) pairs.push(["fontFamily", f.family]);
        pairs.push(["fontFace", f.face]);
        pairs.push(["fontType", f.type]);
        /* 合成フォントにバージョンは無い / Composite fonts carry no version */
        if (!f.isComposite) pairs.push(["version", f.version]);
        pairs.push(["fileName", f.fileName]);
        return pairs;
    }

    /**
     * 構成フォントの一覧を出力すべきかを判定します。
     *
     * @param {FontFields} f - フォント属性。
     * @param {string[]} members - 構成フォント名の配列。
     * @returns {boolean} 出力すべきなら true。
     */
    function hasCompositeMembers(f, members) {
        return f.isComposite && !!members && members.length > 0;
    }

    // =========================================
    // フォント情報の取り出し / Font Field Extraction
    // =========================================
    /**
     * 1フォント分の属性。
     *
     * @typedef {object} FontFields
     * @property {string} name - フォント名（PostScript名）。
     * @property {string} family - ファミリー名。
     * @property {string} face - フェイス名。
     * @property {string} type - フォント種別。合成フォントの場合はローカライズした表記。
     * @property {string} version - バージョン文字列。合成フォントの場合は空文字。
     * @property {string} fileName - フォントファイル名。
     * @property {boolean} isComposite - 合成フォントなら true。
     * @property {boolean} isEmbedded - ドキュメント埋め込みのサブセットで表示されているなら true。
     */

    /**
     * 1エントリ分のフォント属性をまとめて取り出します。
     *
     * @param {string} entryXml - 主フォント1件分のXML断片。
     * @returns {FontFields} 取り出した属性。
     */
    function fontFields(entryXml) {
        var isComposite = getFontProp(entryXml, "composite") === "True";
        var family = getFontProp(entryXml, "fontFamily");
        return {
            name: getFontProp(entryXml, "fontName"),
            family: family,
            face: getFontProp(entryXml, "fontFace"),
            type: isComposite ? getLabel("output.compositeType") : getFontProp(entryXml, "fontType"),
            version: isComposite ? "" : getFontProp(entryXml, "versionString"),
            fileName: getFontProp(entryXml, "fontFileName"),
            isComposite: isComposite,
            isEmbedded: isEmbeddedSubsetFamily(family)
        };
    }

    /**
     * 一覧表示用の「ファミリー フェイス」名を返します。
     *
     * @param {FontFields} f - フォント属性。
     * @returns {string} ファミリー名とフェイス名を連結した名前。前後の空白は除去。
     */
    function fontDisplayName(f) {
        return trimSpaces(f.family + " " + f.face);
    }

    /**
     * 前後の空白を取り除きます。
     *
     * @param {string} value - 対象の文字列。
     * @returns {string} 前後の空白を除いた文字列。
     */
    function trimSpaces(value) {
        return value ? value.toString().replace(/^\s+|\s+$/g, "") : "";
    }

    /**
     * stFnt: 名前空間のタグ値を取得します。
     *
     * @param {string} entryXml - 主フォント1件分のXML断片。
     * @param {string} tag - 名前空間を除いたタグ名（"fontName" など）。
     * @returns {string} タグの値。見つからない場合は空文字。
     */
    function getFontProp(entryXml, tag) {
        return getTagValue(entryXml, "stFnt:" + tag);
    }

    /**
     * 指定タグの内側テキストを取得します（値に改行を含む場合も対応）。
     *
     * @param {string} str - 検索対象のXML断片。
     * @param {string} tag - 名前空間を含むタグ名（"stFnt:fontName" など）。
     * @returns {string} 実体参照を復号した値。見つからない場合は空文字。
     */
    function getTagValue(str, tag) {
        var tagPattern = new RegExp("<" + tag + ">([\\s\\S]*?)</" + tag + ">");
        var match = str.match(tagPattern);
        return (match && match[1]) ? decodeXmlEntities(match[1]) : "";
    }

    // =========================================
    // エスケープ / Escaping
    // =========================================
    /**
     * XMLの実体参照を復号します。
     *
     * @param {string} value - 復号する文字列。
     * @returns {string} 復号した文字列。
     */
    function decodeXmlEntities(value) {
        if (!value) return "";
        /* &amp; は二重復号を避けるため最後に処理 / Decode &amp; last to avoid double decoding */
        return value.toString()
            .replace(/&lt;/g, "<")
            .replace(/&gt;/g, ">")
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&#x([0-9a-fA-F]+);/g, function (whole, hex) {
                return String.fromCharCode(parseInt(hex, 16));
            })
            .replace(/&#([0-9]+);/g, function (whole, dec) {
                return String.fromCharCode(parseInt(dec, 10));
            })
            .replace(/&amp;/g, "&");
    }

    /**
     * CSVセル用にエスケープします（必要なときだけダブルクォートで囲む）。
     *
     * @param {string} value - エスケープする値。
     * @returns {string} エスケープ後の文字列。
     */
    function escapeCsv(value) {
        if (!value) return "";
        value = value.toString();
        /* 改行は CR / LF どちらでも行が壊れるため、両方を対象にする / Both CR and LF break the row, so quote either one */
        return value.match(/[",\r\n]/) ? '"' + value.replace(/"/g, '""') + '"' : value;
    }

    /**
     * Markdown用にエスケープします（アンダースコアのみ）。
     *
     * @param {string} value - エスケープする値。
     * @returns {string} エスケープ後の文字列。
     */
    function escapeMarkdown(value) {
        return value ? value.toString().replace(/_/g, "\\_") : "";
    }

})();
