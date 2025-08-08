#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

TextCountStats.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TextCountStats.jsx

### 概要：

- Illustratorの選択テキストや全体の文字情報を統計的に可視化
- 文字数、段落数、行数、英単語数、全角文字数、半角カナ、種別（ポイント／エリア／パス上）、フォント数などをカウント

### 主な機能：

- 選択したテキストオブジェクトの各種統計をUI上に一覧表示
- 選択がない場合はドキュメント全体のテキストを対象に集計
- UIは日本語／英語対応

### 処理の流れ：

1. ダイアログボックスで統計情報を表示
2. 各パネルに分類（文字・段落、チェック項目、種別、その他）
3. 自動的に選択／全体を切り替えてカウント

### 更新履歴：

- v1.0 (20250806) : 初期バージョン

*/

/*

### Script Name:

TextCountStats.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TextCountStats.jsx

### Description:

- Visualize statistics of selected or all text objects in Illustrator
- Count characters, paragraphs, lines, English words, full-width characters, half-width kana, type (point/area/path), and fonts

### Main Features:

- Shows a summary of various counts in a dialog
- Automatically switches between selected objects and all content
- UI supports Japanese and English

### Flow:

1. Show dialog displaying text statistics
2. Divide stats into panels (Text & Paragraphs, Checks, Types, Others)
3. Auto-select between selection or whole document

### Change Log:

- v1.0 (20250806): Initial version

*/

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    // 言語判定 / Determine language
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "文字数カウント " + SCRIPT_VERSION,
        en: "Text Count Stats " + SCRIPT_VERSION
    },
    charPara: {
        ja: "文字・段落",
        en: "Characters & Paragraphs"
    },
    chars: {
        ja: "文字数：",
        en: "Characters:"
    },
    paras: {
        ja: "段落数：",
        en: "Paragraphs:"
    },
    lines: {
        ja: "行数：",
        en: "Lines:"
    },
    words: {
        ja: "英単語数：",
        en: "English Words:"
    },
    checkTitle: {
        ja: "チェック項目",
        en: "Check Items"
    },
    fullwidth: {
        ja: "全角文字数：",
        en: "Fullwidth Chars:"
    },
    hankakuKana: {
        ja: "半角カナ数：",
        en: "Half-width Kana:"
    },
    kindTitle: {
        ja: "種別",
        en: "Type"
    },
    pointText: {
        ja: "ポイント文字：",
        en: "Point Type:"
    },
    areaText: {
        ja: "エリア内文字：",
        en: "Area Type:"
    },
    pathText: {
        ja: "パス上文字：",
        en: "Type on a Path:"
    },
    otherTitle: {
        ja: "その他",
        en: "Other"
    },
    fonts: {
        ja: "使用フォント数：",
        en: "Fonts Used:"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

function main() {
    try {
        /* 選択オブジェクトを取得 / Get selected objects */
        var sel = app.activeDocument.selection;

        /* 選択がない場合の処理 / Handle case with no selection */
        if (!sel) {
            sel = [];
        }

        /* 選択数をカウント / Count selected objects */
        var count = sel.length;

        /* ドキュメント全体のオブジェクト数をカウント / Count all objects in document */
        var allCount = 0;

        function countAllObjects(items) {
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                allCount++;
                if (it.typename === "GroupItem") {
                    countAllObjects(it.pageItems);
                } else if (it.typename === "CompoundPathItem") {
                    countAllObjects(it.pathItems);
                }
            }
        }
        countAllObjects(app.activeDocument.pageItems);

        /* 種類ごとのカウント（選択と全体） / Count by type (selected and all) */
        var textCountSel = 0,
            textCountAll = 0;
        var pathCountSel = 0,
            pathCountAll = 0;
        var anchorCountSel = 0,
            anchorCountAll = 0;

        /* 種別のカウント（選択と全体） / Count text kinds (selected and all) */
        var pointTextSel = 0, areaTextSel = 0, pathTextSel = 0;
        var pointTextAll = 0, areaTextAll = 0, pathTextAll = 0;

        /* フォントセットを作成 / Create font set */
        var fontSet = {};

        /* 選択分をカウント / Count selected items */
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "TextFrame") {
                textCountSel++;
                if (sel[i].kind === TextType.POINTTEXT) pointTextSel++;
                else if (sel[i].kind === TextType.AREATEXT) areaTextSel++;
                else if (sel[i].kind === TextType.PATHTEXT) pathTextSel++;
                try {
                    var ran = sel[i].textRange || sel[i].textRanges[0];
                    var font = ran.characterAttributes.textFont.name;
                    fontSet[font] = true;
                } catch (e) {}
            } else if (sel[i].typename === "PathItem") {
                pathCountSel++;
                anchorCountSel += sel[i].pathPoints.length;
            } else if (sel[i].typename === "CompoundPathItem") {
                for (var j = 0; j < sel[i].pathItems.length; j++) {
                    pathCountSel++;
                    anchorCountSel += sel[i].pathItems[j].pathPoints.length;
                }
            }
        }

        /* 全体をカウント / Count all items */
        var allItems = app.activeDocument.pageItems;
        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];
            if (obj.typename === "TextFrame") {
                textCountAll++;
                if (obj.kind === TextType.POINTTEXT) pointTextAll++;
                else if (obj.kind === TextType.AREATEXT) areaTextAll++;
                else if (obj.kind === TextType.PATHTEXT) pathTextAll++;
                try {
                    var ran = obj.textRange || obj.textRanges[0];
                    var font = ran.characterAttributes.textFont.name;
                    fontSet[font] = true;
                } catch (e) {}
            } else if (obj.typename === "PathItem") {
                pathCountAll++;
                anchorCountAll += obj.pathPoints.length;
            } else if (obj.typename === "CompoundPathItem") {
                for (var m = 0; m < obj.pathItems.length; m++) {
                    pathCountAll++;
                    anchorCountAll += obj.pathItems[m].pathPoints.length;
                }
            }
        }

        /* 結果を表示（ダイアログボックス） / Display results (dialog box) */
        var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
        dlg.orientation = "column";
        dlg.alignChildren = "center";

        var LABEL_WIDTH = 140; /* 全ラベルの幅を統一 / Set uniform label width */

        /* ラベルと値を2カラムで表示 / Display label and value in two columns */
        var group = dlg.add("group");
        group.orientation = "column";
        group.alignChildren = ["fill", "top"];

        // addRow: ラベル、値、カスタム幅（省略可） / label, value, customWidth (optional)
        function addRow(labelText, valueText, customWidth) {
            var g = group.add("group");
            g.orientation = "row";
            var label = g.add("statictext", undefined, labelText);
            label.justify = "right";
            label.preferredSize.width = (customWidth !== undefined) ? customWidth : LABEL_WIDTH;
            g.add("statictext", undefined, valueText);
        }

        var totalObjSel = count;
        var totalObjAll = allCount;
        var artboardCount = app.activeDocument.artboards.length;

        // 1カラムのグループ / Single-column group
        var columnGroup = dlg.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];

        // 1. 文字・段落の詳細を表示するパネル / Panel to display character and paragraph details
        var panelCharPara = columnGroup.add("panel", undefined, LABELS.charPara[lang]);
        panelCharPara.orientation = "column";
        panelCharPara.alignChildren = ["fill", "top"];
        panelCharPara.margins = [15, 20, 0, 10];
        function addCharParaRow(label, value) {
            var row = panelCharPara.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        // 2. チェック項目パネル / Check items panel
        var panelCheck = columnGroup.add("panel", undefined, LABELS.checkTitle[lang]);
        panelCheck.orientation = "column";
        panelCheck.alignChildren = ["fill", "top"];
        panelCheck.margins = [15, 20, 0, 10];
        function addCheckRow(label, value) {
            var row = panelCheck.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        // 3. 種別パネルを追加 / Add kinds panel
        var panelKinds = columnGroup.add("panel", undefined, LABELS.kindTitle[lang]);
        panelKinds.orientation = "column";
        panelKinds.alignChildren = ["fill", "top"];
        panelKinds.margins = [15, 20, 0, 10];
        function addKindRow(label, value) {
            var row = panelKinds.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        // 4. その他パネル / Other panel
        var panelOther = columnGroup.add("panel", undefined, LABELS.otherTitle[lang]);
        panelOther.orientation = "column";
        panelOther.alignChildren = ["fill", "top"];
        panelOther.margins = [15, 20, 0, 10];
        function addOtherRow(label, value) {
            var row = panelOther.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        var totalCharSel = 0,
            totalCharAll = 0;
        var paraCountSel = 0,
            paraCountAll = 0;
        var wordCountSel = 0, wordCountAll = 0;
        var fullwidthCountSel = 0, fullwidthCountAll = 0;
        var hankakuKanaCountSel = 0, hankakuKanaCountAll = 0;
        // 行数カウント用変数 / Line count variables
        var lineCountSel = 0, lineCountAll = 0;

        /* 選択オブジェクト / Selected objects */
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "TextFrame") {
                try {
                    totalCharSel += sel[i].characters.length;
                } catch (err) {}
                // 段落数カウント（空段落除外）
                try {
                    var paras = sel[i].paragraphs;
                    var validParaCount = 0;
                    for (var p = 0; p < paras.length; p++) {
                        var contents = paras[p].contents;
                        if (contents.replace(/[\s\u3000]/g, "").length > 0) {
                            validParaCount++;
                        }
                    }
                    paraCountSel += validParaCount;
                } catch (err) {}
                // 行数カウント
                try {
                    lineCountSel += sel[i].lines.length;
                } catch (err) {}
                // 英単語数カウント・全角/半角カナカウント
                try {
                    var contents = sel[i].contents;
                    if (typeof contents === "string") {
                        var matches = contents.match(/\b[a-zA-Z]+\b/g);
                        if (matches) wordCountSel += matches.length;
                        // 全角文字数
                        var fwMatches = contents.match(/[\uFF01-\uFF60\uFFE0-\uFFE6]/g);
                        if (fwMatches) fullwidthCountSel += fwMatches.length;
                        // 半角カナ数
                        var kanaMatches = contents.match(/[\uFF65-\uFF9F]/g);
                        if (kanaMatches) hankakuKanaCountSel += kanaMatches.length;
                    }
                } catch (err) {}
            }
        }

        /* 全体 / All objects */
        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];
            if (obj.typename === "TextFrame") {
                try {
                    totalCharAll += obj.characters.length;
                } catch (err) {}
                // 段落数カウント（空段落除外）
                try {
                    var paras = obj.paragraphs;
                    var validParaCount = 0;
                    for (var p = 0; p < paras.length; p++) {
                        var contents = paras[p].contents;
                        if (contents.replace(/[\s\u3000]/g, "").length > 0) {
                            validParaCount++;
                        }
                    }
                    paraCountAll += validParaCount;
                } catch (err) {}
                // 行数カウント
                try {
                    lineCountAll += obj.lines.length;
                } catch (err) {}
                // 英単語数カウント・全角/半角カナカウント
                try {
                    var contents = obj.contents;
                    if (typeof contents === "string") {
                        var matches = contents.match(/\b[a-zA-Z]+\b/g);
                        if (matches) wordCountAll += matches.length;
                        // 全角文字数
                        var fwMatches = contents.match(/[\uFF01-\uFF60\uFFE0-\uFFE6]/g);
                        if (fwMatches) fullwidthCountAll += fwMatches.length;
                        // 半角カナ数
                        var kanaMatches = contents.match(/[\uFF65-\uFF9F]/g);
                        if (kanaMatches) hankakuKanaCountAll += kanaMatches.length;
                    }
                } catch (err) {}
            }
        }

        // 1. 文字・段落パネルに行を追加
        addCharParaRow(LABELS.chars[lang], totalCharSel + " / " + totalCharAll);
        addCharParaRow(LABELS.paras[lang], paraCountSel + " / " + paraCountAll);
        // 行数の行を追加
        addCharParaRow(LABELS.lines[lang], lineCountSel + " / " + lineCountAll);
        addCharParaRow(LABELS.words[lang], wordCountSel + " / " + wordCountAll);

        // 2. チェック項目パネルに行を追加
        addCheckRow(LABELS.fullwidth[lang], fullwidthCountSel + " / " + fullwidthCountAll);
        addCheckRow(LABELS.hankakuKana[lang], hankakuKanaCountSel + " / " + hankakuKanaCountAll);

        // 3. 種別パネルに行を追加
        addKindRow(LABELS.pointText[lang], pointTextSel + " / " + pointTextAll);
        addKindRow(LABELS.areaText[lang], areaTextSel + " / " + areaTextAll);
        addKindRow(LABELS.pathText[lang], pathTextSel + " / " + pathTextAll);

        // 4. その他パネルに行を追加
        // フォント数をカウント / Count font usage
        var fontCount = 0;
        for (var key in fontSet) {
            if (fontSet.hasOwnProperty(key)) fontCount++;
        }
        addOtherRow(LABELS.fonts[lang], fontCount.toString());

        // メイングループ（横並び） / Main group (horizontal layout)
        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.margins = [10, 10, 10, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        // 左側グループ / Left-side button group
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnCancel = btnLeftGroup.add("button", undefined, LABELS.cancel[lang], {
            name: "cancel"
        });

        // スペーサー（伸縮）/ Spacer (stretchable)
        var spacer = btnRowGroup.add("statictext", undefined, "");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        // 右側グループ / Right-side button group
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnOK = btnRightGroup.add("button", undefined, LABELS.ok[lang], {
            name: "ok"
        });
        btnOK.onClick = function() {
            dlg.close();
        };

        /* ダイアログ位置と透明度を調整 / Adjust dialog position and opacity */
        var offsetX = 300;
        var dialogOpacity = 0.97;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function() {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }

        setDialogOpacity(dlg, dialogOpacity);
        shiftDialogPosition(dlg, offsetX, 0);

        dlg.center();
        dlg.show();

    } catch (e) {
        alert("エラーが発生しました: " + e);
    }
}

main();