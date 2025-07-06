#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ConvertFontInfoUI.jsx

### 概要

- 選択したテキストの内容をフォント情報に変換するIllustrator用スクリプトです。
- ダイアログで形式を選択し、選択中テキストをリアルタイムに書き換えることができます。

### 主な機能

- フォントファミリー名、スタイル、PostScript名、フルネーム＋サイズ、3行詳細表示を選択可能
- 変換後のプレビューを即時表示
- フォントサイズは環境設定「文字の単位」に従い小数第2位まで換算表示
- キャンセル時には元のテキスト内容と書式を復元
- 日本語／英語インターフェース対応

### 処理の流れ

1. 選択中のテキストを解析しフォント情報を取得
2. ダイアログで変換形式を選択
3. 選択に応じて即座にプレビュー表示
4. OKで確定、キャンセルで元に戻す

### 更新履歴

- v1.0.0 (20250509) : 初期バージョン

---

### Script Name:

ConvertFontInfoUI.jsx

### Overview

- An Illustrator script to convert the contents of selected text into font information.
- Allows you to select format via dialog and rewrite text content in real time.

### Main Features

- Choose from font family name, style, PostScript name, full name + size, or detailed 3-line view
- Instantly preview the converted text
- Font size is converted based on "Type Units" preference, rounded to two decimals
- Restore original text and formatting when canceled
- Japanese and English UI support

### Process Flow

1. Analyze selected text and retrieve font information
2. Choose conversion format in the dialog
3. Preview changes instantly as you select
4. Confirm with OK, or revert with Cancel

### Update History

- v1.0.0 (20250509): Initial version
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

var LABELS = {
    dialogTitle: { ja: "フォント情報に変換", en: "Convert Font Info" },
    panelTitle:  { ja: "変換形式",         en: "Conversion Format" },
    radio1:      { ja: "フォントファミリー名（.family）", en: "Font Family (.family)" },
    radio2:      { ja: "スタイル（例：Bold / Italic）（.style）", en: "Style (e.g. Bold/Italic) (.style)" },
    radio3:      { ja: "PostScript 名（.name）", en: "PostScript Name (.name)" },
    radio4:      { ja: "フルネーム＋サイズ（.fullName + size）", en: "Full Name + Size (.fullName + size)" },
    radio5:      { ja: "項目ラベル付きの詳細（3行表示）", en: "Labeled Detail (3-line view)" },
    cancel:      { ja: "キャンセル", en: "Cancel" },
    ok:          { ja: "OK",        en: "OK" }
};

// 単位変換処理（pt → 指定単位、小数第2位）
function getFontSizeWithUnit(sizePt) {
    var unitPref = app.preferences.getIntegerPreference('text/units');
    var unitLabel = "pt";
    var convertedSize = sizePt;

    switch (unitPref) {
        case 0: unitLabel = "inch"; convertedSize = sizePt / 72; break;
        case 1: unitLabel = "mm";   convertedSize = sizePt * 25.4 / 72; break;
        case 2: unitLabel = "pt";   convertedSize = sizePt; break;
        case 3: unitLabel = "pica"; convertedSize = sizePt / 12; break;
        case 4: unitLabel = "cm";   convertedSize = sizePt * 2.54 / 72; break;
        case 5: unitLabel = "Q";    convertedSize = sizePt * (25.4 / 72) * 4; break;
        case 6: unitLabel = "px";   convertedSize = sizePt * (96 / 72); break;
    }

    return (Math.round(convertedSize * 100) / 100) + " " + unitLabel;
}

var fontMap = {};

main();

function main() {
    if (app.documents.length === 0) return;

    if (app.selection.constructor.name === "TextRange") {
        var frames = app.selection.story.textFrames;
        if (frames.length === 1) app.selection = [frames[0]];
    }

    var sel = app.selection;
    if (!sel || sel.length === 0 || sel.length >= 1000) return;

    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename !== "TextFrame") continue;
        try {
            var tf = item.textRange;
            var key = item.name || ("_tmp_" + i);
            fontMap[key] = {
                font: tf.characterAttributes.textFont,
                originalSize: tf.characterAttributes.size,
                originalText: item.contents
            };
        } catch (e) {}
    }

    showDialog();
}

function previewChange(format) {
    var sel = app.selection;
    var labelFont = null;

    try {
        labelFont = textFonts.getByName("HiraginoSans-W3");
    } catch (e) {}

    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename !== "TextFrame") continue;

        try {
            var key = item.name || ("_tmp_" + i);
            var saved = fontMap[key];
            if (!saved || !saved.font) continue;

            var baseFont = saved.font;
            var originalSize = saved.originalSize;
            var cr = String.fromCharCode(13);
            var tf = item.textRange;
            var content = "";

            if (format === "label3lines") {
                content =
                    ".family（フォント名）" + cr + baseFont.family + cr +
                    ".style（ウエイト/スタイル）" + cr + baseFont.style + cr +
                    ".name（PostScript名）" + cr + baseFont.name;
                item.contents = content;

                var lines = item.textRange.lines;
                if (lines.length >= 6 && labelFont) {
                    lines[0].characterAttributes.textFont = labelFont;
                    lines[0].characterAttributes.size = 10;
                    lines[2].characterAttributes.textFont = labelFont;
                    lines[2].characterAttributes.size = 10;
                    lines[4].characterAttributes.textFont = labelFont;
                    lines[4].characterAttributes.size = 10;

                    lines[1].characterAttributes.textFont = baseFont;
                    lines[1].characterAttributes.size = originalSize;
                    lines[3].characterAttributes.textFont = baseFont;
                    lines[3].characterAttributes.size = originalSize;
                    lines[5].characterAttributes.textFont = baseFont;
                    lines[5].characterAttributes.size = originalSize;
                }

            } else {
                tf.characterAttributes.textFont = baseFont;
                tf.characterAttributes.size = originalSize;

                switch (format) {
                    case "style":
                        content = baseFont.style;
                        break;
                    case "postscript":
                        content = baseFont.name;
                        break;
                    case "name+style":
                        var sizeText = getFontSizeWithUnit(originalSize);
                        content = (baseFont.fullName && baseFont.fullName !== "")
                            ? baseFont.fullName + "\t" + sizeText
                            : baseFont.family + " " + baseFont.style + "\t" + sizeText;
                        break;
                    default:
                        content = baseFont.family;
                }

                item.contents = content;
            }

        } catch (e) {}
    }

    app.redraw();
}

function restoreOriginalText() {
    var sel = app.selection;
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename !== "TextFrame") continue;

        var key = item.name || ("_tmp_" + i);
        var saved = fontMap[key];
        if (!saved) continue;

        try {
            item.contents = saved.originalText;
            var tf = item.textRange;
            tf.characterAttributes.textFont = saved.font;
            tf.characterAttributes.size = saved.originalSize;
        } catch (e) {}
    }

    app.redraw();
}

function showDialog() {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);

    var outer = dlg.add("group");
    outer.orientation = "column";
    outer.alignChildren = "left";
    outer.alignment = "fill";
    outer.margins = 10;

    var radioPanel = outer.add("panel", undefined, LABELS.panelTitle[lang]);
    radioPanel.orientation = "column";
    radioPanel.alignChildren = "left";
    radioPanel.alignment = "fill";
    radioPanel.margins = [10, 22, 10, 20];

    var r1 = radioPanel.add("radiobutton", undefined, LABELS.radio1[lang]);
    var r2 = radioPanel.add("radiobutton", undefined, LABELS.radio2[lang]);
    var r3 = radioPanel.add("radiobutton", undefined, LABELS.radio3[lang]);
    var r4 = radioPanel.add("radiobutton", undefined, LABELS.radio4[lang]);
    var r5 = radioPanel.add("radiobutton", undefined, LABELS.radio5[lang]);
    r1.value = true;

    r1.onClick = function () { previewChange("name"); };
    r2.onClick = function () { previewChange("style"); };
    r3.onClick = function () { previewChange("postscript"); };
    r4.onClick = function () { previewChange("name+style"); };
    r5.onClick = function () { previewChange("label3lines"); };

    previewChange("name");

    var btnGroup = outer.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignChildren = ["fill", "center"];
    btnGroup.alignment = "fill";
    btnGroup.margins = [10, 20, 10, 0];
    btnGroup.spacing = 10;

    var btnCancel = btnGroup.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var spacer = btnGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 100;
    var btnOK = btnGroup.add("button", undefined, LABELS.ok[lang], { name: "ok" });
    btnOK.alignment = ["right", "center"];

    var result = dlg.show();
    if (result !== 1) { // キャンセル時
        restoreOriginalText();
    }
}