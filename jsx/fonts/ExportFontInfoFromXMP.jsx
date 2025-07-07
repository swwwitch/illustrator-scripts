#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ExportFontInfoFromXMP.jsx

### 概要

- Illustrator ドキュメントに埋め込まれた XMP メタデータから使用フォント情報を抽出し、テキスト / CSV / Markdown の形式で書き出すスクリプトです。
- 書き出し形式をダイアログで選択でき、すべての形式を一括出力することも可能です。

### 主な機能

- TXT / CSV / Markdown の3種類のフォーマットに対応
- CSV は UTF-16（BOM付き）で出力
- Markdown はアンダースコア（_）のみエスケープ処理
- 同名ファイルが存在する場合、自動でリネーム
- 日本語／英語インターフェース対応

### 処理の流れ

1. ドキュメントの XMP メタデータからフォント情報を抽出
2. ダイアログで書き出し形式を選択
3. 指定した形式でフォント情報をデスクトップに保存

### 更新履歴

- v1.0.0 (20250511) : 初期バージョン

---

### Script Name:

ExportFontInfoFromXMP.jsx

### Overview

- A script that extracts font usage information from XMP metadata embedded in an Illustrator document and exports it as text, CSV, or Markdown.
- You can select the output format via a dialog, or export all formats at once.

### Main Features

- Supports three formats: TXT, CSV, and Markdown
- CSV is output in UTF-16 with BOM
- Markdown escapes underscore (_) only
- Automatically renames if duplicate filenames exist
- Japanese and English UI support

### Process Flow

1. Extract font information from document XMP metadata
2. Select export format in the dialog
3. Save the font information on the desktop in the specified format

### Update History

- v1.0.0 (20250511): Initial version
*/

(function () {

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();
    var LABELS = {
        dialogTitle:     { ja: "フォント情報を書き出し", en: "Export Font Info" },
        dialogPrompt:    { ja: "書き出し形式を選んでください：", en: "Select output format:" },
        optText:         { ja: "テキストファイル（.txt）",   en: "Text File (.txt)" },
        optCSV:          { ja: "CSVファイル（.csv）",         en: "CSV File (.csv)" },
        optMD:           { ja: "Markdownファイル（.md）",     en: "Markdown File (.md)" },
        optAll:          { ja: "すべて（3種類書き出し）",     en: "All Formats (TXT + CSV + MD)" },
        cancel:          { ja: "キャンセル", en: "Cancel" },
        ok:              { ja: "OK",        en: "OK" },
        alertNoDoc:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        alertNoFonts:    { ja: "フォント情報が見つかりませんでした。", en: "No font information found." },
        alertDone:       { ja: "保存完了：\n", en: "Exported:\n" },
        errorMsg:        { ja: "エラーが発生しました：\n", en: "An error occurred:\n" },
        compositeFontsLabel: { ja: "構成フォント", en: "Composite Fonts" },
        txtHeader:       { ja: "使用フォント一覧：", en: "Font List: " },
        fontCount:       { ja: "使用フォント数：", en: "Font Count: " },
        fontListItem:    { ja: "・", en: "- " },
        mdHeader:        { ja: "# 使用フォント一覧：", en: "# Font List: " },
        mdFontListTitle: { ja: "## フォント一覧", en: "## Font List" },
        mdFontDetailTitle:{ ja: "## 各フォントの情報", en: "## Font Details" }
    };

    if (app.documents.length === 0) {
        alert(LABELS.alertNoDoc[lang]);
        return;
    }

    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = ["left", "top"];
    dlg.spacing = 10;
    dlg.margins = 20;

    dlg.add("statictext", undefined, LABELS.dialogPrompt[lang]);

    var formatGroup = dlg.add("group");
    formatGroup.orientation = "column";
    formatGroup.alignChildren = ["left", "top"];
    formatGroup.spacing = 6;
    formatGroup.margins = 5;

var rbText = formatGroup.add("radiobutton", undefined, LABELS.optText[lang]);
var rbCSV  = formatGroup.add("radiobutton", undefined, LABELS.optCSV[lang]);
var rbMD   = formatGroup.add("radiobutton", undefined, LABELS.optMD[lang]);
var rbAll  = formatGroup.add("radiobutton", undefined, LABELS.optAll[lang]);
rbText.value = true;
rbText.active = true;

    var radioButtons = [rbText, rbCSV, rbMD, rbAll];
    for (var i = 0; i < radioButtons.length; i++) {
        (function(index) {
            radioButtons[index].addEventListener("keydown", function(k) {
                var key = k.keyName;
                if (key === "Up" || key === "ArrowUp") {
                    var prev = (index + radioButtons.length - 1) % radioButtons.length;
                    radioButtons[prev].value = true;
                    radioButtons[prev].notify();
                } else if (key === "Down" || key === "ArrowDown") {
                    var next = (index + 1) % radioButtons.length;
                    radioButtons[next].value = true;
                    radioButtons[next].notify();
                }
            });
        })(i);
    }

    var btns = dlg.add("group");
    btns.alignment = "right";
    btns.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    btns.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    if (dlg.show() !== 1) return;

    var exportTypes = [];
    if (rbAll.value) exportTypes = ["txt", "csv", "md"];
    else if (rbCSV.value) exportTypes = ["csv"];
    else if (rbMD.value) exportTypes = ["md"];
    else exportTypes = ["txt"];

    try {
        if (ExternalObject.AdobeXMPScript === undefined)
            ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');

        var doc = app.activeDocument;
        var docName = doc.name.replace(/\.[^\.]+$/, "");
        var docBaseName = doc.name;
        var xmpString = doc.XMPString;

        var fontsMatch = xmpString.match(/<xmpTPg:Fonts>[\s\S]*?<\/xmpTPg:Fonts>/);
        if (!fontsMatch) {
            alert(LABELS.alertNoFonts[lang]);
            return;
        }

        var fontsBlock = fontsMatch[0];
        var fontEntries = fontsBlock.match(/<rdf:li[\s\S]*?<\/rdf:li>/g);
        if (!fontEntries || fontEntries.length === 0) {
            alert(LABELS.alertNoFonts[lang]);
            return;
        }

        var mainEntries = [];
        var structMap = {};
        var lastIndex = -1;

        for (var i = 0; i < fontEntries.length; i++) {
            var li = fontEntries[i];
            if (li.indexOf('rdf:parseType="Resource"') >= 0) {
                lastIndex++;
                mainEntries.push(li);
                structMap[lastIndex] = [];
            } else {
                // プレーンな構成フォント
                structMap[lastIndex].push(li.replace(/<[^>]+>/g, ""));
            }
        }

        for (var i = 0; i < exportTypes.length; i++) {
            var type = exportTypes[i];
            var lines = buildLines(type, mainEntries, docBaseName, structMap);
            var baseName = docName + "_fontInfo";
            var ext = "." + type;
            var file = new File(Folder.desktop + "/" + baseName + ext);
            var n = 2;
            while (file.exists) {
                file = new File(Folder.desktop + "/" + baseName + "_" + n + ext);
                n++;
            }

            if (type === "csv") {
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

        // alert(LABELS.alertDone[lang] + Folder.desktop.fsName);

    } catch (e) {
        alert(LABELS.errorMsg[lang] + e);
    }

    function buildLines(type, entries, fileName, structMap) {
        var lines = [];

        if (type === "txt") {
            lines.push(LABELS.txtHeader[lang] + fileName + "\n");
            lines.push(LABELS.fontCount[lang] + entries.length + "\n");

            for (var i = 0; i < entries.length; i++) {
                var fam = get(entries[i], "fontFamily");
                var face = get(entries[i], "fontFace");
                lines.push(LABELS.fontListItem[lang] + fam + " " + face);
            }

            lines.push("\n-----------------------------");

            for (var i = 0; i < entries.length; i++) {
                var b = entries[i];
                var fontType = get(b, "fontType");
                var isComposite = get(b, "composite") === "True";
                lines.push("fontName:\t"   + get(b, "fontName"));
                lines.push("fontFamily:\t" + get(b, "fontFamily"));
                lines.push("fontFace:\t"   + get(b, "fontFace"));
                lines.push("fontType:\t" + (isComposite ? "合成フォント" : fontType));
                if (!isComposite) lines.push("version:\t" + get(b, "versionString"));
                lines.push("fileName:\t"   + get(b, "fontFileName"));
                if (isComposite && structMap[i] && structMap[i].length > 0) {
                    lines.push(LABELS.compositeFontsLabel[lang] + ":");
                    for (var j = 0; j < structMap[i].length; j++) {
                        lines.push("・" + structMap[i][j]);
                    }
                }
                lines.push("-----------------------------");
            }

        } else if (type === "md") {
            
            lines.push(LABELS.mdHeader[lang] + "：" + md(fileName) + "\n");
            lines.push("[TOC]");
            lines.push("");
            lines.push(LABELS.mdFontListTitle[lang] + "\n");
            lines.push(LABELS.fontCount[lang] + entries.length + "\n");

            for (var i = 0; i < entries.length; i++) {
                var fam = get(entries[i], "fontFamily");
                var face = get(entries[i], "fontFace");
                lines.push("- " + md(fam + " " + face));
            }

            lines.push("\n" + LABELS.mdFontDetailTitle[lang] + "\n");

            for (var i = 0; i < entries.length; i++) {
                var b = entries[i];
                var fam = md(get(b, "fontFamily"));
                lines.push("### " + fam);
                lines.push("");
                lines.push("- fontName: "   + md(get(b, "fontName")));
                lines.push("- fontFace: "   + md(get(b, "fontFace")));
                var fontType = get(b, "fontType");
                var isComposite = get(b, "composite") === "True";
                lines.push("- fontType: " + (isComposite ? "合成フォント" : md(fontType)));
                if (!isComposite) lines.push("- version: " + md(get(b, "versionString")));
                lines.push("- fileName: "   + md(get(b, "fontFileName")));
                if (isComposite && structMap[i] && structMap[i].length > 0) {
                    lines.push("");
                    lines.push("#### 構成フォント");
                    lines.push("");
                    for (var j = 0; j < structMap[i].length; j++) {
                        lines.push("- " + md(structMap[i][j]));
                    }
                    lines.push("");
                }
                lines.push("");
            }

        } else if (type === "csv") {
            lines.push("fontName,fontFamily,fontFace,fontType,version,fileName");
            for (var i = 0; i < entries.length; i++) {
                var b = entries[i];
                var fontType = get(b, "fontType");
                var isComposite = get(b, "composite") === "True";
                lines.push([
                    csv(get(b, "fontName")),
                    csv(get(b, "fontFamily")),
                    csv(get(b, "fontFace")),
                    csv(isComposite ? "合成フォント" : fontType),
                    csv(isComposite ? "" : get(b, "versionString")),
                    csv(get(b, "fontFileName"))
                ].join(","));
            }
        }

        return lines;
    }

    function get(str, tag) {
        return getTagValue(str, "stFnt:" + tag);
    }

    function getTagValue(str, tag) {
        var re = new RegExp("<" + tag + ">(.*?)</" + tag + ">");
        var match = str.match(re);
        return match && match[1] ? match[1] : "";
    }

    function csv(s) {
        if (!s) return "";
        s = s.toString();
        return s.match(/["\,\n]/) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }

    function md(s) {
        return s ? s.toString().replace(/_/g, "\\_") : "";
    }

})();