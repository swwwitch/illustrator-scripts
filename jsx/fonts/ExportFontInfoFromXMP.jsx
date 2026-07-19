/*

### 概要 / Overview

- Illustrator ドキュメントに埋め込まれた XMP メタデータから使用フォント情報を抽出し、TXT / CSV / Markdown の形式で書き出します。
- ダイアログで書き出し形式と書き出し先を選択でき、3 種類すべてを一括出力することも可能です。
- Extracts font usage information from the XMP metadata embedded in an Illustrator document and exports it as TXT / CSV / Markdown.
- The export format and destination are chosen in a dialog, and all three formats can be exported at once.

### 主な機能 / Main Features

- TXT / CSV / Markdown の 3 形式に対応 / Supports TXT, CSV, and Markdown
- 書き出し先をデスクトップ／ファイルと同じ階層から選択 / Destination can be the desktop or the document's own folder
- CSV は UTF-16（BOM 付き）で出力 / CSV is written in UTF-16 with BOM
- Markdown はアンダースコア（_）のみエスケープ / Markdown escapes underscore (_) only
- 同名ファイルが存在する場合は自動でリネーム / Automatically renames when a duplicate filename exists
- 日本語／英語 UI 対応 / Japanese and English UI

### 処理の流れ / Process Flow

1. XMP メタデータからフォント情報を抽出 / Extract font information from XMP metadata
2. ダイアログで書き出し形式と書き出し先を選択 / Choose the export format and destination in the dialog
3. 選択した書き出し先に指定形式で保存 / Save to the chosen destination in the selected format

### 補足 / Notes

- フォント情報は保存済みの XMP から取得するため、未保存のドキュメントでは実行できません / Font info comes from the saved XMP, so the document must be saved before running

### 紹介記事 / Article

https://note.com/dtp_tranist/n/n16e7e95652b6

### 更新履歴 / Update History

- v1.0.0 (20250511) : 初期バージョン / Initial version
- v1.0.1 (20260617) : 書き出し先（デスクトップ／同じ階層）の選択、パネル化、未保存チェックを追加 / Added destination choice (desktop / same folder), panel layout, and unsaved-document check

*/

#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExportFontInfoFromXMP";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 出力ファイル名のサフィックスと区切り線 / Output filename suffix and section divider */
    var FILENAME_SUFFIX = "_fontInfo";
    var SECTION_DIVIDER = "-----------------------------";

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    // =========================================
    // ローカライズ / Localization
    // =========================================
    /* 現在の UI 言語を判定 / Detect the current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

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
            sameFolder: { ja: "ファイルと同じ階層", en: "Same folder as the file" }
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

    /* ドット区切りパスでローカライズ文字列を取得 / Resolve a localized label by dot path */
    function getLabel(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var idx = 0; idx < parts.length; idx++) {
            node = node[parts[idx]];
            if (node === undefined) return path;
        }
        return (node[currentLanguage] !== undefined) ? node[currentLanguage] : node.en;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(path) {
        return getLabel(path) + (currentLanguage === "ja" ? "：" : ":");
    }

    /* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
    function labelWithCount(path, count) {
        if (currentLanguage === "ja") return getLabel(path) + "（" + count + "）";
        return getLabel(path) + " (" + count + ")";
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    /* 未保存のドキュメントは XMP にフォント情報が無いため、ダイアログ前に中止 / An unsaved document has no font info in XMP, so stop before the dialog */
    if (!app.activeDocument.saved) {
        alert(getLabel("alert.notSaved"));
        return;
    }

    var dialogResult = showFormatDialog();
    if (!dialogResult) return;

    exportFontInfo(dialogResult.formats, dialogResult.destination);

    // =========================================
    // UI ヘルパー / UI Helpers
    // =========================================
    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /* 書き出し形式と書き出し先を選ぶダイアログを表示し、{ formats, destination } を返す（キャンセル時は null）/ Show the dialog and return { formats, destination } (null on cancel) */
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

        /* 上下キーでラジオボタンを循環移動 / Cycle radio buttons with the arrow keys */
        var formatRadios = [radioText, radioCsv, radioMarkdown, radioAll];
        for (var i = 0; i < formatRadios.length; i++) {
            (function (index) {
                formatRadios[index].addEventListener("keydown", function (k) {
                    var key = k.keyName;
                    if (key === "Up" || key === "ArrowUp") {
                        var prevIndex = (index + formatRadios.length - 1) % formatRadios.length;
                        formatRadios[prevIndex].value = true;
                        formatRadios[prevIndex].notify();
                    } else if (key === "Down" || key === "ArrowDown") {
                        var nextIndex = (index + 1) % formatRadios.length;
                        formatRadios[nextIndex].value = true;
                        formatRadios[nextIndex].notify();
                    }
                });
            })(i);
        }

        /* 書き出し先パネル / Destination panel */
        var destinationPanel = dialog.add("panel", undefined, getLabel("destination.title"));
        setupPanel(destinationPanel, 6);

        var radioDesktop = destinationPanel.add("radiobutton", undefined, getLabel("destination.desktop"));
        var radioSameFolder = destinationPanel.add("radiobutton", undefined, getLabel("destination.sameFolder"));
        radioDesktop.value = true;

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
            destination: radioSameFolder.value ? "sameFolder" : "desktop"
        };
    }

    // =========================================
    // 書き出し / Export
    // =========================================
    /* XMP からフォント情報を抽出し、指定された各形式でファイルを書き出す / Extract font info from XMP and write a file for each format */
    function exportFontInfo(formats, destination) {
        try {
            if (ExternalObject.AdobeXMPScript === undefined)
                ExternalObject.AdobeXMPScript = new ExternalObject("lib:AdobeXMPScript");

            var doc = app.activeDocument;
            var docFullName = doc.name;
            var docNameNoExt = doc.name.replace(/\.[^\.]+$/, "");
            /* デスクトップ、またはドキュメントと同じフォルダー / Desktop, or the document's own folder */
            var outputFolder = (destination === "sameFolder") ? doc.path : Folder.desktop;

            var fontData = extractFontData(doc.XMPString);
            if (!fontData) {
                alert(getLabel("alert.noFonts"));
                return;
            }

            for (var i = 0; i < formats.length; i++) {
                var format = formats[i];
                var lines = buildLines(format, fontData, docFullName);
                var file = uniqueOutputFile(outputFolder, docNameNoExt + FILENAME_SUFFIX, "." + format);
                writeLines(file, format, lines);
            }

        } catch (e) {
            alert(getLabel("alert.error") + e);
        }
    }

    /* XMP 文字列から主フォントと構成フォントを抽出 / Parse primary fonts and composite members from the XMP string */
    function extractFontData(xmpString) {
        var fontsMatch = xmpString.match(/<xmpTPg:Fonts>[\s\S]*?<\/xmpTPg:Fonts>/);
        if (!fontsMatch) return null;

        var fontEntries = fontsMatch[0].match(/<rdf:li[\s\S]*?<\/rdf:li>/g);
        if (!fontEntries || fontEntries.length === 0) return null;

        var primaryEntries = [];
        var compositeMembers = {}; // 主フォント index → 構成フォント名の配列 / primary index → composite member names
        var primaryIndex = -1;

        for (var i = 0; i < fontEntries.length; i++) {
            var entryXml = fontEntries[i];
            if (entryXml.indexOf('rdf:parseType="Resource"') >= 0) {
                primaryIndex++;
                primaryEntries.push(entryXml);
                compositeMembers[primaryIndex] = [];
            } else if (primaryIndex >= 0) {
                /* タグを除いたプレーンな構成フォント名（主フォントが先に出現した場合のみ）/ Plain composite member name with tags stripped (only after a primary font appears) */
                compositeMembers[primaryIndex].push(entryXml.replace(/<[^>]+>/g, ""));
            }
        }

        return { primaryEntries: primaryEntries, compositeMembers: compositeMembers };
    }

    /* 指定フォルダー内で重複しないファイルオブジェクトを返す / Return a non-colliding File in the given folder */
    function uniqueOutputFile(folder, baseName, ext) {
        var file = new File(folder + "/" + baseName + ext);
        var counter = 2;
        while (file.exists) {
            file = new File(folder + "/" + baseName + "_" + counter + ext);
            counter++;
        }
        return file;
    }

    /* 形式に応じたエンコーディングで行を書き込む / Write lines using the encoding for the format */
    function writeLines(file, format, lines) {
        if (format === "csv") {
            file.encoding = "UTF-16";
            file.open("w");
            file.write("\uFEFF" + lines.join("\r\n"));
        } else {
            file.encoding = "UTF-8";
            file.open("w");
            file.write(lines.join("\n"));
        }
        file.close();
    }

    // =========================================
    // 行の生成 / Line Builders
    // =========================================
    /* 形式ごとの行生成へ振り分け / Dispatch to the per-format line builder */
    function buildLines(format, fontData, fileName) {
        if (format === "md") return buildMarkdownLines(fontData, fileName);
        if (format === "csv") return buildCsvLines(fontData);
        return buildTxtLines(fontData, fileName);
    }

    /* テキスト形式の行を生成 / Build lines for the TXT format */
    function buildTxtLines(fontData, fileName) {
        var entries = fontData.primaryEntries;
        var lines = [];

        lines.push(labelText("output.fontListHeading") + " " + fileName + "\n");
        lines.push(labelWithCount("output.fontCount", entries.length) + "\n");

        for (var i = 0; i < entries.length; i++) {
            lines.push(getLabel("output.bullet") + fontDisplayName(entries[i]));
        }

        lines.push("\n" + SECTION_DIVIDER);

        for (var i = 0; i < entries.length; i++) {
            var f = fontFields(entries[i]);
            lines.push("fontName:\t" + f.name);
            lines.push("fontFamily:\t" + f.family);
            lines.push("fontFace:\t" + f.face);
            lines.push("fontType:\t" + f.type);
            if (!f.isComposite) lines.push("version:\t" + f.version);
            lines.push("fileName:\t" + f.fileName);

            var members = fontData.compositeMembers[i];
            if (f.isComposite && members && members.length > 0) {
                lines.push(labelText("output.compositeFonts"));
                for (var j = 0; j < members.length; j++) {
                    lines.push(getLabel("output.bullet") + members[j]);
                }
            }
            lines.push(SECTION_DIVIDER);
        }

        return lines;
    }

    /* Markdown 形式の行を生成 / Build lines for the Markdown format */
    function buildMarkdownLines(fontData, fileName) {
        var entries = fontData.primaryEntries;
        var lines = [];

        lines.push("# " + getLabel("output.fontListHeading") + " " + escapeMarkdown(fileName) + "\n");
        lines.push("[TOC]");
        lines.push("");
        lines.push("## " + getLabel("output.fontListHeading") + "\n");
        lines.push(labelWithCount("output.fontCount", entries.length) + "\n");

        for (var i = 0; i < entries.length; i++) {
            lines.push("- " + escapeMarkdown(fontDisplayName(entries[i])));
        }

        lines.push("\n## " + getLabel("output.fontDetailHeading") + "\n");

        for (var i = 0; i < entries.length; i++) {
            var f = fontFields(entries[i]);
            lines.push("### " + escapeMarkdown(f.family));
            lines.push("");
            lines.push("- fontName: " + escapeMarkdown(f.name));
            lines.push("- fontFace: " + escapeMarkdown(f.face));
            lines.push("- fontType: " + escapeMarkdown(f.type));
            if (!f.isComposite) lines.push("- version: " + escapeMarkdown(f.version));
            lines.push("- fileName: " + escapeMarkdown(f.fileName));

            var members = fontData.compositeMembers[i];
            if (f.isComposite && members && members.length > 0) {
                lines.push("");
                lines.push("#### " + getLabel("output.compositeFonts"));
                lines.push("");
                for (var j = 0; j < members.length; j++) {
                    lines.push("- " + escapeMarkdown(members[j]));
                }
                lines.push("");
            }
            lines.push("");
        }

        return lines;
    }

    /* CSV 形式の行を生成 / Build lines for the CSV format */
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

    // =========================================
    // フォント情報の取り出し / Font Field Extraction
    // =========================================
    /* 1 エントリ分のフォント属性をまとめて取り出す / Pull all font attributes for one entry at once */
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

    /* 一覧表示用の「ファミリー フェイス」名 / "Family Face" name for list display */
    function fontDisplayName(entryXml) {
        return getFontProp(entryXml, "fontFamily") + " " + getFontProp(entryXml, "fontFace");
    }

    /* stFnt: 名前空間のタグ値を取得 / Read a tag value in the stFnt: namespace */
    function getFontProp(entryXml, tag) {
        return getTagValue(entryXml, "stFnt:" + tag);
    }

    /* 指定タグの内側テキストを取得 / Read the inner text of the given tag */
    function getTagValue(str, tag) {
        var tagPattern = new RegExp("<" + tag + ">(.*?)</" + tag + ">");
        var match = str.match(tagPattern);
        return (match && match[1]) ? match[1] : "";
    }

    // =========================================
    // エスケープ / Escaping
    // =========================================
    /* CSV セル用のエスケープ（必要時のみダブルクォート）/ Escape a CSV cell (quote only when needed) */
    function escapeCsv(value) {
        if (!value) return "";
        value = value.toString();
        return value.match(/["\,\n]/) ? '"' + value.replace(/"/g, '""') + '"' : value;
    }

    /* Markdown 用のエスケープ（アンダースコアのみ）/ Escape for Markdown (underscore only) */
    function escapeMarkdown(value) {
        return value ? value.toString().replace(/_/g, "\\_") : "";
    }

})();
