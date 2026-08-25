#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメントに埋め込まれたXMPメタデータから使用フォント情報を抽出し、TXT / CSV / Markdownで書き出します。
フォント情報は保存済みのXMPから取得するため、未保存のドキュメントでは実行できません。

詳細は README を参照してください。

### Overview

Extracts font usage information from the XMP metadata embedded in the document and exports it as TXT / CSV / Markdown.
The data comes from the saved XMP, so the script cannot run on an unsaved document.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExportFontInfoFromXMP";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-06";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExportFontInfoFromXMP.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExportFontInfoFromXMP.md
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
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            notSaved: {
                ja: "ドキュメントが保存されていません。\n保存してから実行してください。",
                en: "The document has not been saved.\nPlease save it before running."
            },
            noFonts: { ja: "フォント情報が見つかりませんでした。", en: "No font information found." },
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
     * @property {function} build - 行を生成する関数。(FontData, string) => string[]
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

    /* 未保存のドキュメントは XMP にフォント情報が無いため、ダイアログ前に中止 / An unsaved document has no font info in XMP, so stop before the dialog */
    if (!isDocumentSaved(app.activeDocument)) {
        alert(getLabel("alert.notSaved"));
        return;
    }

    var dialogResult = showFormatDialog();
    if (!dialogResult) return;

    exportFontInfo(dialogResult.formats, dialogResult.destination, dialogResult.openFolder);

    /**
     * ドキュメントが保存済みか（未編集で、保存先のファイルが実在するか）を判定します。
     *
     * @param {Document} doc - 判定対象のドキュメント。
     * @returns {boolean} 保存済みなら true。
     */
    function isDocumentSaved(doc) {
        /* saved が false の時点で未保存なので、fullName は保存済みのときしか参照しない / fullName is only touched once the document has been saved */
        return doc.saved && doc.fullName.exists;
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
                    if (key === "Up" || key === "ArrowUp") {
                        selectRadio(radios, (index + radios.length - 1) % radios.length);
                    } else if (key === "Down" || key === "ArrowDown") {
                        selectRadio(radios, (index + 1) % radios.length);
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
        enableArrowKeyNavigation([radioDesktop, radioSameFolder]);

        /* ラジオボタンと同じ間隔だと選択肢に見えるため、少し離す / Nudge it down so it doesn't read as a third radio option */
        var openFolderGroup = destinationPanel.add("group");
        setupGroup(openFolderGroup, "row");
        openFolderGroup.margins = [0, 4, 0, 0];

        var openFolderCheckbox = openFolderGroup.add("checkbox", undefined, getLabel("destination.openFolder"));
        openFolderCheckbox.value = OPEN_FOLDER_DEFAULT;

        var buttonGroup = dialog.add("group");
        setupGroup(buttonGroup, "row");
        buttonGroup.alignment = "center";
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
            openFolder: openFolderCheckbox.value
        };
    }

    // =========================================
    // 書き出し / Export
    // =========================================
    /**
     * XMPからフォント情報を抽出し、指定された各形式でファイルを書き出します。
     *
     * @param {string[]} formats - 書き出す形式（"txt" / "csv" / "md"）の配列。
     * @param {string} destination - "desktop" または "sameFolder"。
     * @param {boolean} openFolder - 書き出し後にフォルダーを開く場合は true。
     * @returns {void}
     */
    function exportFontInfo(formats, destination, openFolder) {
        try {
            if (ExternalObject.AdobeXMPScript === undefined)
                ExternalObject.AdobeXMPScript = new ExternalObject("lib:AdobeXMPScript");

            var doc = app.activeDocument;
            var fontData = extractFontData(doc.XMPString);
            if (!fontData) {
                alert(getLabel("alert.noFonts"));
                return;
            }

            /* デスクトップ、またはドキュメントと同じフォルダー / Desktop, or the document's own folder */
            var outputFolder = (destination === "sameFolder") ? doc.path : Folder.desktop;
            var writtenNames = writeFontInfoFiles(fontData, formats, outputFolder, doc.name);

            if (openFolder) {
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
     * @param {FontData} fontData - 抽出済みのフォント情報。
     * @param {string[]} formats - 書き出す形式（"txt" / "csv" / "md"）の配列。
     * @param {Folder} outputFolder - 書き出し先フォルダー。
     * @param {string} docName - 拡張子を含むドキュメント名。
     * @returns {string[]} 書き出したファイル名の配列。
     */
    function writeFontInfoFiles(fontData, formats, outputFolder, docName) {
        var baseName = docName.replace(/\.[^\.]+$/, "") + FILENAME_SUFFIX;
        var writtenNames = [];

        for (var i = 0; i < formats.length; i++) {
            var spec = FORMAT_SPECS[formats[i]];
            var file = uniqueOutputFile(outputFolder, baseName, "." + formats[i]);
            writeLines(file, spec, spec.build(fontData, docName));
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
     * XMPから取り出したフォント情報。
     *
     * @typedef {object} FontData
     * @property {string[]} primaryEntries - 主フォント1件分のXML断片の配列。
     * @property {Array<string[]>} compositeMembers - primaryEntries と同じ並びの、構成フォント名の配列。
     */

    /**
     * XMP文字列から主フォントと構成フォントを抽出します。
     *
     * @param {string} xmpString - ドキュメントのXMP文字列。
     * @returns {FontData|null} 抽出結果。フォント情報が無い場合は null。
     */
    function extractFontData(xmpString) {
        var fontsMatch = xmpString.match(/<xmpTPg:Fonts>[\s\S]*?<\/xmpTPg:Fonts>/);
        if (!fontsMatch) return null;

        /* 構成フォントは主フォントの rdf:li に入れ子で並ぶため、主フォントの開始タグで分割する
           / Composite members are nested inside the primary rdf:li, so split on the primary's start tag */
        var chunks = fontsMatch[0].split(/<rdf:li[^>]*rdf:parseType="Resource"[^>]*>/);

        var primaryEntries = [];
        var compositeMembers = [];

        /* chunks[0] は最初の主フォントより前の部分なので読み飛ばす / chunks[0] precedes the first primary font */
        for (var i = 1; i < chunks.length; i++) {
            primaryEntries.push(chunks[i]);
            compositeMembers.push(extractCompositeMembers(chunks[i]));
        }

        /* 主フォントが1件も無ければ空ファイルを作らない / Don't write an empty file when there is no primary font */
        if (primaryEntries.length === 0) return null;

        return { primaryEntries: primaryEntries, compositeMembers: compositeMembers };
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
    // 行の生成 / Line Builders
    // =========================================
    /**
     * テキスト形式（タブ区切り）の行を生成します。
     *
     * @param {FontData} fontData - 抽出済みのフォント情報。
     * @param {string} fileName - 見出しに使うドキュメント名。
     * @returns {string[]} 書き込む行の配列。
     */
    function buildTxtLines(fontData, fileName) {
        var entries = fontData.primaryEntries;
        var bullet = getLabel("output.bullet");
        var lines = [];

        lines.push(labelText("output.fontListHeading") + " " + fileName + "\n");
        lines.push(labelWithCount("output.fontCount", entries.length) + "\n");

        for (var i = 0; i < entries.length; i++) {
            lines.push(bullet + fontDisplayName(entries[i]));
        }

        lines.push("\n" + SECTION_DIVIDER);

        for (i = 0; i < entries.length; i++) {
            lines = lines.concat(txtFontBlock(entries[i], fontData.compositeMembers[i]));
        }

        return lines;
    }

    /**
     * テキスト形式の、フォント1件分の行を生成します。
     *
     * @param {string} entryXml - 主フォント1件分のXML断片。
     * @param {string[]} members - 構成フォント名の配列。
     * @returns {string[]} フォント1件分の行。
     */
    function txtFontBlock(entryXml, members) {
        var f = fontFields(entryXml);
        var pairs = fontDetailPairs(f, true);
        var lines = [];

        for (var i = 0; i < pairs.length; i++) {
            lines.push(pairs[i][0] + ":\t" + pairs[i][1]);
        }

        if (hasCompositeMembers(f, members)) {
            var bullet = getLabel("output.bullet");
            lines.push(labelText("output.compositeFonts"));
            for (var j = 0; j < members.length; j++) {
                lines.push(bullet + members[j]);
            }
        }

        lines.push(SECTION_DIVIDER);
        return lines;
    }

    /**
     * Markdown形式の行を生成します。
     *
     * @param {FontData} fontData - 抽出済みのフォント情報。
     * @param {string} fileName - 見出しに使うドキュメント名。
     * @returns {string[]} 書き込む行の配列。
     */
    function buildMarkdownLines(fontData, fileName) {
        var entries = fontData.primaryEntries;
        var lines = [];

        lines.push("# " + getLabel("output.fontListHeading") + " " + escapeMarkdown(fileName) + "\n");
        // lines.push("[TOC]");
        lines.push("");
        lines.push("## " + getLabel("output.fontListHeading") + "\n");
        lines.push(labelWithCount("output.fontCount", entries.length) + "\n");

        for (var i = 0; i < entries.length; i++) {
            lines.push("- " + escapeMarkdown(fontDisplayName(entries[i])));
        }

        lines.push("\n## " + getLabel("output.fontDetailHeading") + "\n");

        for (i = 0; i < entries.length; i++) {
            lines = lines.concat(markdownFontBlock(entries[i], fontData.compositeMembers[i]));
        }

        return lines;
    }

    /**
     * Markdown形式の、フォント1件分の行を生成します。
     *
     * @param {string} entryXml - 主フォント1件分のXML断片。
     * @param {string[]} members - 構成フォント名の配列。
     * @returns {string[]} フォント1件分の行。
     */
    function markdownFontBlock(entryXml, members) {
        var f = fontFields(entryXml);
        /* ファミリー名は見出しに出るので詳細では省く / The family is already in the heading, so skip it in the details */
        var pairs = fontDetailPairs(f, false);
        var lines = [];

        /* 同じファミリーで複数フェイスがあると見出しが重複するため、フェイスまで含める / Include the face; family alone duplicates headings across faces */
        lines.push("### " + escapeMarkdown(fontDisplayName(entryXml)));
        lines.push("");

        for (var i = 0; i < pairs.length; i++) {
            lines.push("- " + pairs[i][0] + ": " + escapeMarkdown(pairs[i][1]));
        }

        if (hasCompositeMembers(f, members)) {
            lines.push("");
            lines.push("#### " + getLabel("output.compositeFonts"));
            lines.push("");
            for (var j = 0; j < members.length; j++) {
                lines.push("- " + escapeMarkdown(members[j]));
            }
            lines.push("");
        }

        lines.push("");
        return lines;
    }

    /**
     * CSV形式の行を生成します。
     *
     * @param {FontData} fontData - 抽出済みのフォント情報。
     * @returns {string[]} ヘッダー行を含む、書き込む行の配列。
     */
    function buildCsvLines(fontData) {
        var entries = fontData.primaryEntries;
        var lines = ["fontName,fontFamily,fontFace,fontType,version,fileName"];

        for (var i = 0; i < entries.length; i++) {
            var f = fontFields(entries[i]);
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
     * @param {string} entryXml - 主フォント1件分のXML断片。
     * @returns {string} ファミリー名とフェイス名を連結した名前。前後の空白は除去。
     */
    function fontDisplayName(entryXml) {
        var displayName = getFontProp(entryXml, "fontFamily") + " " + getFontProp(entryXml, "fontFace");
        return displayName.replace(/^\s+|\s+$/g, "");
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
