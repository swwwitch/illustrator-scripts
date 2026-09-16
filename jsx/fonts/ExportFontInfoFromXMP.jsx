#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメントに埋め込まれたXMPメタデータから使用フォント情報を抽出し、TXT / CSV / Markdownで書き出します。
未保存・編集中のドキュメントでは、XMPに無いフォントをドキュメント内のテキストから補います。

詳細は README を参照してください。

### Overview

Extracts font usage information from the XMP metadata embedded in the document and exports it as TXT / CSV / Markdown.
On an unsaved or edited document, fonts missing from the XMP are topped up by scanning the text in the document.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExportFontInfoFromXMP";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-17";                   /* 更新日 / last updated */

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

    /* ［書き出し後にフォルダーを開く］の初期値 / Default for "Open the folder after exporting" */
    var OPEN_FOLDER_DEFAULT = true;

    /* ［見つからないフォントのみ］の初期値 / Default for "Missing fonts only" */
    var MISSING_ONLY_DEFAULT = false;

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

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
        options: {
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
            error: { ja: "エラーが発生しました：\n", en: "An error occurred:\n" }
        },
        output: {
            bullet: { ja: "・", en: "- " },
            fontListHeading: { ja: "使用フォント一覧", en: "Font List" },
            fontCount: { ja: "使用フォント数", en: "Font Count" },
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
     * @property {function} build - 行を生成する関数。(FontEntry[], string) => string[]
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
        var radioCsv = formatPanel.add("radiobutton", undefined, getLabel("format.csv"));
        var radioMarkdown = formatPanel.add("radiobutton", undefined, getLabel("format.markdown"));
        var radioAll = formatPanel.add("radiobutton", undefined, getLabel("format.all"));
        radioText.value = true;
        radioText.active = true;
        enableArrowKeyNavigation([radioText, radioCsv, radioMarkdown, radioAll]);

        /* 書き出し先パネル / Destination panel */
        var destinationPanel = dialog.add("panel", undefined, getLabel("destination.title"));
        setupPanel(destinationPanel, 6);

        var radioDesktop = destinationPanel.add("radiobutton", undefined, getLabel("destination.desktop"));
        var radioSameFolder = destinationPanel.add("radiobutton", undefined, getLabel("destination.sameFolder"));
        radioDesktop.value = true;
        /* 未保存のドキュメントには同じ階層が無いのでデスクトップだけにする / A never-saved document has no folder of its own, so leave only Desktop */
        radioSameFolder.enabled = !!getDocumentFolder(app.activeDocument);
        enableArrowKeyNavigation([radioDesktop, radioSameFolder]);

        /* ラジオボタンと同じ間隔だと選択肢に見えるため、少し離す / Nudge it down so it doesn't read as a third radio option */
        var openFolderGroup = destinationPanel.add("group");
        setupGroup(openFolderGroup, "row");
        openFolderGroup.margins = [0, 4, 0, 0];

        var openFolderCheckbox = openFolderGroup.add("checkbox", undefined, getLabel("destination.openFolder"));
        openFolderCheckbox.value = OPEN_FOLDER_DEFAULT;

        /* オプションパネル / Options panel */
        var optionsPanel = dialog.add("panel", undefined, getLabel("options.title"));
        setupPanel(optionsPanel, 6);

        var missingOnlyCheckbox = optionsPanel.add("checkbox", undefined, getLabel("options.missingOnly"));
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
     * @param {ExportOptions} options - ダイアログで選択された書き出し条件。
     * @returns {void}
     */
    function exportFontInfo(options) {
        try {
            var doc = app.activeDocument;
            var entries = collectFontEntries(doc);
            if (entries.length === 0) {
                alert(getLabel("alert.noFonts"));
                return;
            }

            if (options.missingOnly) {
                entries = filterMissingFonts(entries);
                if (entries.length === 0) {
                    alert(getLabel("alert.noMissingFonts"));
                    return;
                }
            }

            /* デスクトップ、またはドキュメントと同じフォルダー / Desktop, or the document's own folder */
            var docFolder = getDocumentFolder(doc);
            var outputFolder = (options.destination === "sameFolder" && docFolder) ? docFolder : Folder.desktop;
            var writtenNames = writeFontInfoFiles(entries, options.formats, outputFolder, doc.name);

            if (options.openFolder) {
                outputFolder.execute();
            } else {
                /* フォルダーを開かない場合だけ、書き出し結果をアラートで知らせる / Report the result only when the folder is not opened */
                alert(getLabel("alert.done") + "\n" + writtenNames.join("\n"));
            }

        } catch (e) {
            alert(getLabel("alert.error") + (e.message ? e.message : e) + (e.line ? "\n(line " + e.line + ")" : ""));
        }
    }

    /**
     * 指定された各形式でファイルを書き出します。
     *
     * @param {FontEntry[]} entries - 抽出済みのフォント情報。
     * @param {string[]} formats - 書き出す形式（"txt" / "csv" / "md"）の配列。
     * @param {Folder} outputFolder - 書き出し先フォルダー。
     * @param {string} docName - 拡張子を含むドキュメント名。
     * @returns {string[]} 書き出したファイル名の配列。
     */
    function writeFontInfoFiles(entries, formats, outputFolder, docName) {
        var baseName = docName.replace(/\.[^\.]+$/, "") + FILENAME_SUFFIX;
        var writtenNames = [];

        for (var i = 0; i < formats.length; i++) {
            var spec = FORMAT_SPECS[formats[i]];
            var file = uniqueOutputFile(outputFolder, baseName, "." + formats[i]);
            writeLines(file, spec, spec.build(entries, docName));
            writtenNames.push(decodeURI(file.name));
        }

        return writtenNames;
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
            throw new Error(getLabel("alert.writeFailed") + decodeURI(file.fsName));
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
     * @returns {FontEntry[]} フォント情報の配列。1件も無ければ空配列。
     */
    function collectFontEntries(doc) {
        var xmpEntries = extractFontEntriesFromXMP(readXMPString(doc));

        /* 保存済みならXMPが現状と一致する。未保存・編集中のXMPは前回保存時のままなので、
           テキストから拾った分を足して補う（テキスト走査は重いのでここでしか走らせない）
           / A saved document's XMP matches what's on screen; an unsaved or edited one still holds the
             last-saved XMP, so top it up from the text (the costly scan runs only in that case) */
        if (doc.saved) return xmpEntries;

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
            if (isFontEntrySeen(seen, textEntries[i])) continue;
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
            seen[keys[i]] = true;
        }
    }

    /**
     * フォント1件が既出かを判定します。
     *
     * @param {object} seen - 既出フォントの索引。
     * @param {FontEntry} entry - 判定するフォント。
     * @returns {boolean} 既出なら true。
     */
    function isFontEntrySeen(seen, entry) {
        var keys = fontEntryKeys(entry);
        for (var i = 0; i < keys.length; i++) {
            if (seen[keys[i]]) return true;
        }
        return false;
    }

    /**
     * 重複判定に使うキーを返します。
     *
     * XMPとDOMで綴りが揃わないことがあるため、PostScript名と表示名の両方で照合します。
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
            /* フォントは文字単位で変わるため、1文字ずつ見るしかない（範囲単位では混在時に先頭の書式しか返らない）
               / Fonts change per character, and a mixed range reports only its first character's font */
            var chars = frames[i].textRange.characters;
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

                /* Object のプロパティ名と衝突させないため接頭辞を付ける / Prefix the key so font names can't collide with Object members */
                var key = "font:" + font.name;
                if (seen[key]) continue;
                seen[key] = true;

                entries.push({
                    fields: {
                        name: font.name,
                        family: font.family,
                        face: font.style,
                        type: "",
                        version: "",
                        fileName: "",
                        isComposite: false
                    },
                    members: []
                });
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
        for (var i = 0; i < fonts.length; i++) {
            /* XMPの値がPostScript名・ファミリー名・表示名のどれで入っていても引けるようにする
               / Index all three spellings, since XMP may carry the PostScript, family, or display name */
            installedFontIndex["font:" + fonts[i].name] = true;
            installedFontIndex["font:" + fonts[i].family] = true;
            installedFontIndex["font:" + (fonts[i].family + " " + fonts[i].style)] = true;
        }
        return installedFontIndex;
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
     * フォント1件が見つからないフォントかを判定します。
     *
     * @param {FontEntry} entry - 判定対象のフォント。
     * @returns {boolean} 見つからないフォントなら true。
     */
    function isFontMissing(entry) {
        var f = entry.fields;

        /* 合成フォント自体は app.textFonts に無いので、名前で引くと必ず外れる。
           構成フォントが取れないときは判定できないため、欠落とは見なさない
           / A composite font is never listed in app.textFonts, so looking up its own name always misses.
             With no members to judge by, treat it as undetermined rather than missing */
        if (f.isComposite) {
            for (var i = 0; i < entry.members.length; i++) {
                if (!isFontInstalled(entry.members[i])) return true;
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
     * テキスト形式（タブ区切り）の行を生成します。
     *
     * @param {FontEntry[]} entries - 抽出済みのフォント情報。
     * @param {string} fileName - 見出しに使うドキュメント名。
     * @returns {string[]} 書き込む行の配列。
     */
    function buildTxtLines(entries, fileName) {
        var bullet = getLabel("output.bullet");
        var lines = [];

        lines.push(labelText("output.fontListHeading") + " " + fileName + "\n");
        lines.push(labelWithCount("output.fontCount", entries.length) + "\n");

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
     * @returns {string[]} 書き込む行の配列。
     */
    function buildMarkdownLines(entries, fileName) {
        var lines = [];

        lines.push("# " + getLabel("output.fontListHeading") + " " + escapeMarkdown(fileName) + "\n");
        // lines.push("[TOC]");
        lines.push("");
        lines.push("## " + getLabel("output.fontListHeading") + "\n");
        lines.push(labelWithCount("output.fontCount", entries.length) + "\n");

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
     */

    /**
     * 1エントリ分のフォント属性をまとめて取り出します。
     *
     * @param {string} entryXml - 主フォント1件分のXML断片。
     * @returns {FontFields} 取り出した属性。
     */
    function fontFields(entryXml) {
        var isComposite = getFontProp(entryXml, "composite") === "True";
        return {
            name: getFontProp(entryXml, "fontName"),
            family: getFontProp(entryXml, "fontFamily"),
            face: getFontProp(entryXml, "fontFace"),
            type: isComposite ? getLabel("output.compositeType") : getFontProp(entryXml, "fontType"),
            version: isComposite ? "" : getFontProp(entryXml, "versionString"),
            fileName: getFontProp(entryXml, "fontFileName"),
            isComposite: isComposite
        };
    }

    /**
     * 一覧表示用の「ファミリー フェイス」名を返します。
     *
     * @param {FontFields} f - フォント属性。
     * @returns {string} ファミリー名とフェイス名を連結した名前。前後の空白は除去。
     */
    function fontDisplayName(f) {
        return (f.family + " " + f.face).replace(/^\s+|\s+$/g, "");
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
