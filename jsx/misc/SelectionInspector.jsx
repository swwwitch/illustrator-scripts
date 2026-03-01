#target illustrator

/*

### スクリプト名：

SelectionInspector.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択中および全体のオブジェクト数をカウント
- テキスト・画像・グループ・パスの詳細情報を表示
- 統計情報をダイアログで表示、書き出し可能

### 主な機能：

- オブジェクト数、アートボード数を表示
- 文字数、段落数、テキスト種類の表示
- リンク画像／埋め込み画像／リンク切れの検出
- グループ／クリップグループのカウント
- パス（オープン／クローズ／アンカーポイント）のカウント
- 「書き出し」ボタンでテキストファイルとして保存

### 処理の流れ：

- ドキュメントと選択オブジェクトを解析し、各種情報を取得
- ダイアログに2カラムで情報を整然と表示
- 書き出し機能でデスクトップにテキスト保存

### 更新履歴：

- v1.0 (20250806) : 初期バージョン
- v1.1 (20250806) : 書き出し機能を追加
- v1.2 (20250807) : UI調整
- v1.3 (20260301) : パスのハンドル数を追加
- v1.4 (20260301) : 強制改行（ソフトリターン）数を追加

---

### Script Name:

SelectionInspector.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Count selected and total objects in the document
- Display detailed stats for text, images, groups, and paths
- Show statistics in a dialog and export as a text file

### Main Features:

- Display number of objects and artboards
- Show characters, paragraphs, text type counts
- Detect linked/embedded/broken images
- Count groups and clipping groups
- Count paths, open/closed paths, and anchor points
- Export button writes stats to desktop as a text file

### Process Flow:

- Analyze document and selected objects
- Display data in a two-column dialog UI
- Export stats to desktop text file

### Changelog:

- v1.0 (20250806): Initial version
- v1.1 (20250806): Added export feature
- v1.2 (20250807): UI adjustments
- v1.3 (20260301): Added handle count for paths
- v1.4 (20260301): Added forced line break (soft return) count

*/

var SCRIPT_VERSION = "v1.4";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "選択／全体オブジェクト数 " + SCRIPT_VERSION,
        en: "Selection / All Objects Count " + SCRIPT_VERSION
    },
    other: {
        ja: "基本",
        en: "Basics"
    },
    artboards: {
        ja: "アートボード：",
        en: "Artboards:"
    },
    objects: {
        ja: "オブジェクト：",
        en: "Objects:"
    },
    charPara: {
        ja: "文字・段落",
        en: "Characters & Paragraphs"
    },
    transparency: {
        ja: "透明",
        en: "Transparency"
    },
    chars: {
        ja: "文字数：",
        en: "Characters:"
    },
    paras: {
        ja: "段落数：",
        en: "Paragraphs:"
    },
    forcedBreaks: {
        ja: "強制改行：",
        en: "Line Breaks:"
    },
    opacityLt100: {
        ja: "不透明度<100：",
        en: "Opacity < 100:"
    },
    blendNotNormal: {
        ja: "描画モード≠通常：",
        en: "Blend Mode ≠ Normal:"
    },
    texts: {
        ja: "テキスト：",
        en: "Text Frames:"
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
        en: "Path Text:"
    },
    link: {
        ja: "配置画像",
        en: "Images"
    },
    linked: {
        ja: "リンク：",
        en: "Linked Images:"
    },
    embed: {
        ja: "埋め込み：",
        en: "Embedded Images:"
    },
    broken: {
        ja: "リンク切れ：",
        en: "Broken Links:"
    },
    group: {
        ja: "グループ：",
        en: "Groups:"
    },
    clipGroup: {
        ja: "クリップグループ：",
        en: "Clipping Groups:"
    },
    path: {
        ja: "パス",
        en: "Paths"
    },
    pathCount: {
        ja: "パス：",
        en: "Paths:"
    },
    openPath: {
        ja: "オープンパス：",
        en: "Open Paths:"
    },
    closedPath: {
        ja: "クローズパス：",
        en: "Closed Paths:"
    },
    anchors: {
        ja: "アンカーポイント：",
        en: "Anchor Points:"
    },
    handles: {
        ja: "ハンドル：",
        en: "Handles:"
    },
    compoundPath: {
        ja: "複合パス：",
        en: "Compound Paths:"
    },
    compoundShape: {
        ja: "複合シェイプ：",
        en: "Compound Shapes:"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    exportPreset: {
        ja: "書き出し",
        en: "Export"
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
        var handleCountSel = 0,
            handleCountAll = 0;
        var compoundPathSel = 0,
            compoundPathAll = 0;
        var compoundShapeSel = 0,
            compoundShapeAll = 0;

        // 透明（不透明度<100／描画モード≠通常）/ Transparency stats
        var opacityLt100Sel = 0,
            opacityLt100All = 0;
        var blendNotNormalSel = 0,
            blendNotNormalAll = 0;

        function countTransparencyForItem(it) {
            var r = { opacityLt100: 0, blendNotNormal: 0 };
            if (!it) return r;

            try {
                if (typeof it.opacity === "number" && it.opacity < 100) r.opacityLt100 = 1;
            } catch (e) { }

            try {
                if (it.blendingMode !== undefined && it.blendingMode !== BlendModes.NORMAL) r.blendNotNormal = 1;
            } catch (e) { }

            return r;
        }

        function countHandlesInPathItem(pi) {
            // Count bezier handles: left/right direction differs from anchor
            // A corner point typically has both directions == anchor => 0
            // A smooth/curve point often has 2 handles
            var c = 0;
            try {
                var pts = pi.pathPoints;
                for (var i = 0; i < pts.length; i++) {
                    var p = pts[i];
                    var a = p.anchor;
                    var l = p.leftDirection;
                    var r = p.rightDirection;
                    if (l[0] !== a[0] || l[1] !== a[1]) c++;
                    if (r[0] !== a[0] || r[1] !== a[1]) c++;
                }
            } catch (e) { }
            return c;
        }

        /* 選択分をカウント / Count selected items */
        for (var i = 0; i < sel.length; i++) {
            // CompoundPath/CompoundShape
            if (sel[i].typename === "CompoundPathItem") {
                compoundPathSel++;
            }
            if (sel[i].typename === "PluginItem") {
                // 複合シェイプ（ライブパスファインダー）
                try {
                    if (sel[i].name && sel[i].name.indexOf("Compound Shape") !== -1) {
                        compoundShapeSel++;
                    }
                } catch (e) { }
            }
            // 透明（選択）/ Transparency (selected)
            try {
                var t0 = countTransparencyForItem(sel[i]);
                opacityLt100Sel += t0.opacityLt100;
                blendNotNormalSel += t0.blendNotNormal;
            } catch (e) { }
            if (sel[i].typename === "TextFrame") {
                textCountSel++;
            } else if (sel[i].typename === "PathItem") {
                pathCountSel++;
                anchorCountSel += sel[i].pathPoints.length;
                handleCountSel += countHandlesInPathItem(sel[i]);
            } else if (sel[i].typename === "CompoundPathItem") {
                for (var j = 0; j < sel[i].pathItems.length; j++) {
                    pathCountSel++;
                    anchorCountSel += sel[i].pathItems[j].pathPoints.length;
                    handleCountSel += countHandlesInPathItem(sel[i].pathItems[j]);
                }
            }
        }

        /* 全体をカウント / Count all items */
        var allItems = app.activeDocument.pageItems;
        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];

            if (obj.typename === "CompoundPathItem") {
                compoundPathAll++;
            }
            if (obj.typename === "PluginItem") {
                try {
                    if (obj.name && obj.name.indexOf("Compound Shape") !== -1) {
                        compoundShapeAll++;
                    }
                } catch (e) { }
            }

            // 透明（全体）/ Transparency (all)
            try {
                var tAll = countTransparencyForItem(obj);
                opacityLt100All += tAll.opacityLt100;
                blendNotNormalAll += tAll.blendNotNormal;
            } catch (e) { }

            if (obj.typename === "TextFrame") {
                textCountAll++;
            } else if (obj.typename === "PathItem") {
                pathCountAll++;
                anchorCountAll += obj.pathPoints.length;
                handleCountAll += countHandlesInPathItem(obj);
            } else if (obj.typename === "CompoundPathItem") {
                for (var m = 0; m < obj.pathItems.length; m++) {
                    pathCountAll++;
                    anchorCountAll += obj.pathItems[m].pathPoints.length;
                    handleCountAll += countHandlesInPathItem(obj.pathItems[m]);
                }
            }
        }

        /* 結果を表示（ダイアログボックス） / Display results (dialog box) */
        var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
        dlg.orientation = "column";
        dlg.alignChildren = "center";

        var LABEL_WIDTH = 130; /* 全ラベルの幅を統一 / Set uniform label width */
        var LABEL_WIDTH_LEFT = 100; /* 左カラムのラベル幅 / Label width for left column */
        var LABEL_WIDTH_TOP = 90; // オブジェクト・アートボード数専用ラベル幅
        var LABEL_WIDTH_RIGHT = 140; /* 右カラムのラベル幅 / Label width for right column */

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

        // 2カラムのグループ / Two-column group
        var twoColGroup = dlg.add("group");
        twoColGroup.orientation = "row";
        twoColGroup.alignChildren = ["fill", "top"];

        // 左カラム
        var leftCol = twoColGroup.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        /* その他の詳細を表示するパネル / Panel to display other details */
        var panelOther = leftCol.add("panel", undefined, LABELS.other[lang]);
        panelOther.orientation = "column";
        panelOther.alignChildren = ["fill", "top"];
        panelOther.margins = [15, 20, 0, 10];

        function addOtherRow(label, value) {
            var row = panelOther.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        addOtherRow(LABELS.artboards[lang], artboardCount.toString());
        addOtherRow(LABELS.objects[lang], totalObjSel + " / " + totalObjAll);



        /* テキストの詳細を表示するパネル / Panel to display text details */
        var panelText = leftCol.add("panel", undefined, LABELS.texts[lang]);
        panelText.orientation = "column";
        panelText.alignChildren = ["fill", "top"];
        panelText.margins = [15, 20, 0, 10];

        /* 文字・段落の詳細を表示するパネル / Panel to display character and paragraph details */
        var panelCharPara = leftCol.add("panel", undefined, LABELS.charPara[lang]);
        panelCharPara.orientation = "column";
        panelCharPara.alignChildren = ["fill", "top"];
        panelCharPara.margins = [15, 20, 0, 10];

        /* 透明の詳細を表示するパネル / Panel to display transparency details */
        var panelTransparency = leftCol.add("panel", undefined, LABELS.transparency[lang]);
        panelTransparency.orientation = "column";
        panelTransparency.alignChildren = ["fill", "top"];
        panelTransparency.margins = [15, 20, 0, 10];

        function addTransparencyRow(label, value) {
            var row = panelTransparency.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        function addCharParaRow(label, value) {
            var row = panelCharPara.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        var totalCharSel = 0,
            totalCharAll = 0;
        var paraCountSel = 0,
            paraCountAll = 0;
        var forcedBreakSel = 0,
            forcedBreakAll = 0;

        // 強制改行（ソフトリターン）をカウント / Count forced line breaks (soft returns)
        // 段落改行（\r）は除外し、\u0003 / \n / \u2028 を強制改行として数える
        function countForcedBreaksFromContents(s) {
            if (!s || !s.length) return 0;
            var m = s.match(/[\u0003\n\u2028]/g);
            return m ? m.length : 0;
        }

        var pointTextSel = 0,
            areaTextSel = 0,
            pathTextSel = 0;
        var pointTextAll = 0,
            areaTextAll = 0,
            pathTextAll = 0;

        /* 選択オブジェクト / Selected objects */
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "TextFrame") {
                try {
                    totalCharSel += sel[i].characters.length;
                } catch (err) { }
                try {
                    paraCountSel += sel[i].paragraphs.length;
                } catch (err) { }
                // 強制改行（ソフトリターン）/ Forced line breaks (soft returns)
                try {
                    forcedBreakSel += countForcedBreaksFromContents(sel[i].contents);
                } catch (err) { }
                if (sel[i].kind === TextType.POINTTEXT) {
                    pointTextSel++;
                } else if (sel[i].kind === TextType.AREATEXT) {
                    areaTextSel++;
                } else if (sel[i].kind === TextType.PATHTEXT) {
                    pathTextSel++;
                }
            }
        }

        /* 全体 / All objects */
        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];
            if (obj.typename === "TextFrame") {
                try {
                    totalCharAll += obj.characters.length;
                } catch (err) { }
                try {
                    paraCountAll += obj.paragraphs.length;
                } catch (err) { }
                // 強制改行（ソフトリターン）/ Forced line breaks (soft returns)
                try {
                    forcedBreakAll += countForcedBreaksFromContents(obj.contents);
                } catch (err) { }
                if (obj.kind === TextType.POINTTEXT) {
                    pointTextAll++;
                } else if (obj.kind === TextType.AREATEXT) {
                    areaTextAll++;
                } else if (obj.kind === TextType.PATHTEXT) {
                    pathTextAll++;
                }
            }
        }

        function addTextRow(label, value) {
            var row = panelText.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_LEFT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        addCharParaRow(LABELS.chars[lang], totalCharSel + " / " + totalCharAll);
        addCharParaRow(LABELS.paras[lang], paraCountSel + " / " + paraCountAll);
        addCharParaRow(LABELS.forcedBreaks[lang], forcedBreakSel + " / " + forcedBreakAll);
        addTextRow(LABELS.texts[lang], textCountSel + " / " + textCountAll);
        addTextRow(LABELS.pointText[lang], pointTextSel + " / " + pointTextAll);
        addTextRow(LABELS.areaText[lang], areaTextSel + " / " + areaTextAll);
        addTextRow(LABELS.pathText[lang], pathTextSel + " / " + pathTextAll);

        addTransparencyRow(LABELS.opacityLt100[lang], opacityLt100Sel + " / " + opacityLt100All);
        addTransparencyRow(LABELS.blendNotNormal[lang], blendNotNormalSel + " / " + blendNotNormalAll);

        // 右カラム
        var rightCol = twoColGroup.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["fill", "top"];

        /* 画像の詳細を表示するパネル / Panel to display image details */
        var panelImage = rightCol.add("panel", undefined, LABELS.link[lang]);
        panelImage.orientation = "column";
        panelImage.alignChildren = ["fill", "top"];
        panelImage.margins = [15, 20, 0, 10];

        var linkedSel = 0,
            linkedAll = 0;
        var embeddedSel = 0,
            embeddedAll = 0;
        var brokenLinkSel = 0,
            brokenLinkAll = 0;

        /* 選択オブジェクト / Selected objects */
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "PlacedItem") {
                if (sel[i].embedded) {
                    embeddedSel++;
                } else {
                    linkedSel++;
                    try {
                        var f = sel[i].file;
                        if (!f || !f.exists) {
                            brokenLinkSel++;
                        }
                    } catch (err) {
                        brokenLinkSel++;
                    }
                }
            }
        }

        /* 全体 / All objects */
        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];
            if (obj.typename === "PlacedItem") {
                if (obj.embedded) {
                    embeddedAll++;
                } else {
                    linkedAll++;
                    try {
                        var fAll = obj.file;
                        if (!fAll || !fAll.exists) {
                            brokenLinkAll++;
                        }
                    } catch (err) {
                        brokenLinkAll++;
                    }
                }
            }
        }

        function addImageRow(label, value) {
            var row = panelImage.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_RIGHT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        addImageRow(LABELS.linked[lang], linkedSel + " / " + linkedAll);
        addImageRow(LABELS.embed[lang], embeddedSel + " / " + embeddedAll);
        addImageRow(LABELS.broken[lang], brokenLinkSel + " / " + brokenLinkAll);

        /* グループの詳細を表示するパネル / Panel to display group details */
        var panelGroup = rightCol.add("panel", undefined, LABELS.group[lang]);
        panelGroup.orientation = "column";
        panelGroup.alignChildren = ["fill", "top"];
        panelGroup.margins = [15, 20, 0, 10];

        var groupSel = 0,
            groupAll = 0;
        var clipGroupSel = 0,
            clipGroupAll = 0;

        /* 選択オブジェクト / Selected objects */
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "GroupItem") {
                groupSel++;
                if (sel[i].clipped) {
                    clipGroupSel++;
                }
            }
        }

        /* 全体 / All objects */
        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];
            if (obj.typename === "GroupItem") {
                groupAll++;
                if (obj.clipped) {
                    clipGroupAll++;
                }
            }
        }

        function addGroupRow(label, value) {
            var row = panelGroup.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_RIGHT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }

        addGroupRow(LABELS.group[lang], groupSel + " / " + groupAll);
        addGroupRow(LABELS.clipGroup[lang], clipGroupSel + " / " + clipGroupAll);

        /* パスの詳細を表示するパネル / Panel to display path details */
        var panelPath = rightCol.add("panel", undefined, LABELS.path[lang]);
        panelPath.orientation = "column";
        panelPath.alignChildren = ["fill", "top"];
        panelPath.margins = [15, 20, 0, 10];

        /* パス総数表示行を最上部に追加 / Add total path count row at top */
        function addPathRow(label, value) {
            var row = panelPath.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", [0, 0, LABEL_WIDTH_RIGHT, 20], label);
            lbl.justify = "right";
            var val = row.add("statictext", undefined, value);
            val.characters = value.length;
        }
        addPathRow(LABELS.pathCount[lang], pathCountSel + " / " + pathCountAll);

        /* パスの種類をカウント / Count types of paths */
        var openPathSel = 0,
            openPathAll = 0;
        var closedPathSel = 0,
            closedPathAll = 0;

        /* 選択オブジェクト / Selected objects */
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "PathItem") {
                if (sel[i].closed) {
                    closedPathSel++;
                } else {
                    openPathSel++;
                }
            }
        }

        /* 全体 / All objects */
        for (var k = 0; k < allItems.length; k++) {
            var obj = allItems[k];
            if (obj.typename === "PathItem") {
                if (obj.closed) {
                    closedPathAll++;
                } else {
                    openPathAll++;
                }
            }
        }

        addPathRow(LABELS.openPath[lang], openPathSel + " / " + openPathAll);
        addPathRow(LABELS.closedPath[lang], closedPathSel + " / " + closedPathAll);
        addPathRow(LABELS.anchors[lang], anchorCountSel + " / " + anchorCountAll);
        addPathRow(LABELS.handles[lang], handleCountSel + " / " + handleCountAll);
        addPathRow(LABELS.compoundPath[lang], compoundPathSel + " / " + compoundPathAll);
        addPathRow(LABELS.compoundShape[lang], compoundShapeSel + " / " + compoundShapeAll);

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
        var btnExport = btnRightGroup.add("button", undefined, LABELS.exportPreset[lang], {
            name: "preview"
        });
        var btnOK = btnRightGroup.add("button", undefined, LABELS.ok[lang], {
            name: "ok"
        });

        btnExport.onClick = function () {
            try {
                var docName = app.activeDocument.name.replace(/\.[^\.]+$/, ""); // 拡張子を除いたファイル名
                var today = new Date();
                var yyyy = today.getFullYear();
                var mm = ("0" + (today.getMonth() + 1)).slice(-2);
                var dd = ("0" + today.getDate()).slice(-2);
                var dateStr = yyyy + mm + dd;

                var desktopPath = Folder.desktop + "/count-" + docName + "-" + dateStr + ".txt";
                var file = new File(desktopPath);

                // Helper: write a label + "sel / all" (layout aligned with dialog)
                function wPair(labelJa, labelEn, selVal, allVal) {
                    if (lang === "ja") {
                        file.writeln(labelJa + " " + selVal + " / " + allVal);
                    } else {
                        file.writeln(labelEn + " " + selVal + " / " + allVal);
                    }
                }

                // Helper: write a label + single value
                function wSingle(labelJa, labelEn, val) {
                    if (lang === "ja") {
                        file.writeln(labelJa + " " + val);
                    } else {
                        file.writeln(labelEn + " " + val);
                    }
                }

                // Helper: section header
                function wSection(titleJa, titleEn) {
                    file.writeln("");
                    file.writeln(lang === "ja" ? titleJa : titleEn);
                }

                if (file.open("w")) {
                    // Header
                    if (lang === "ja") {
                        file.writeln("Selection Inspector Report");
                        file.writeln("Document: " + app.activeDocument.name);
                        file.writeln("Date: " + yyyy + "-" + mm + "-" + dd);
                        file.writeln("");
                        file.writeln("※ 値は『選択 / 全体』の形式です（アートボードのみ全体）");
                    } else {
                        file.writeln("Selection Inspector Report");
                        file.writeln("Document: " + app.activeDocument.name);
                        file.writeln("Date: " + yyyy + "-" + mm + "-" + dd);
                        file.writeln("");
                        file.writeln("Note: values are formatted as 'Selection / All' (Artboards is All only)");
                    }

                    // Basics / Misc
                    wSection("基本", "Basics");
                    wSingle(LABELS.artboards.ja.replace(/：$/, ":"), LABELS.artboards.en.replace(/:$/, ":"), artboardCount);
                    wPair(LABELS.objects.ja.replace(/：$/, ":"), LABELS.objects.en.replace(/:$/, ":"), totalObjSel, totalObjAll);

                    // Text Frames
                    wSection("テキスト", "Text Frames");
                    wPair(LABELS.texts.ja.replace(/：$/, ":"), LABELS.texts.en.replace(/:$/, ":"), textCountSel, textCountAll);
                    wPair(LABELS.pointText.ja.replace(/：$/, ":"), LABELS.pointText.en.replace(/:$/, ":"), pointTextSel, pointTextAll);
                    wPair(LABELS.areaText.ja.replace(/：$/, ":"), LABELS.areaText.en.replace(/:$/, ":"), areaTextSel, areaTextAll);
                    wPair(LABELS.pathText.ja.replace(/：$/, ":"), LABELS.pathText.en.replace(/:$/, ":"), pathTextSel, pathTextAll);

                    // Characters & Paragraphs
                    wSection("文字・段落", "Characters & Paragraphs");
                    wPair(LABELS.chars.ja.replace(/：$/, ":"), LABELS.chars.en.replace(/:$/, ":"), totalCharSel, totalCharAll);
                    wPair(LABELS.paras.ja.replace(/：$/, ":"), LABELS.paras.en.replace(/:$/, ":"), paraCountSel, paraCountAll);
                    wPair(LABELS.forcedBreaks.ja.replace(/：$/, ":"), LABELS.forcedBreaks.en.replace(/:$/, ":"), forcedBreakSel, forcedBreakAll);

                    // Transparency
                    wSection("透明", "Transparency");
                    wPair(LABELS.opacityLt100.ja.replace(/：$/, ":"), LABELS.opacityLt100.en.replace(/:$/, ":"), opacityLt100Sel, opacityLt100All);
                    wPair(LABELS.blendNotNormal.ja.replace(/：$/, ":"), LABELS.blendNotNormal.en.replace(/:$/, ":"), blendNotNormalSel, blendNotNormalAll);

                    // Images
                    wSection("配置画像", "Images");
                    wPair(LABELS.linked.ja.replace(/：$/, ":"), LABELS.linked.en.replace(/:$/, ":"), linkedSel, linkedAll);
                    wPair(LABELS.embed.ja.replace(/：$/, ":"), LABELS.embed.en.replace(/:$/, ":"), embeddedSel, embeddedAll);
                    wPair(LABELS.broken.ja.replace(/：$/, ":"), LABELS.broken.en.replace(/:$/, ":"), brokenLinkSel, brokenLinkAll);

                    // Groups
                    wSection("グループ", "Groups");
                    wPair(LABELS.group.ja.replace(/：$/, ":"), LABELS.group.en.replace(/:$/, ":"), groupSel, groupAll);
                    wPair(LABELS.clipGroup.ja.replace(/：$/, ":"), LABELS.clipGroup.en.replace(/:$/, ":"), clipGroupSel, clipGroupAll);

                    // Paths
                    wSection("パス", "Paths");
                    wPair(LABELS.pathCount.ja.replace(/：$/, ":"), LABELS.pathCount.en.replace(/:$/, ":"), pathCountSel, pathCountAll);
                    wPair(LABELS.openPath.ja.replace(/：$/, ":"), LABELS.openPath.en.replace(/:$/, ":"), openPathSel, openPathAll);
                    wPair(LABELS.closedPath.ja.replace(/：$/, ":"), LABELS.closedPath.en.replace(/:$/, ":"), closedPathSel, closedPathAll);
                    wPair(LABELS.anchors.ja.replace(/：$/, ":"), LABELS.anchors.en.replace(/:$/, ":"), anchorCountSel, anchorCountAll);
                    wPair(LABELS.handles.ja.replace(/：$/, ":"), LABELS.handles.en.replace(/:$/, ":"), handleCountSel, handleCountAll);
                    wPair(LABELS.compoundPath.ja.replace(/：$/, ":"), LABELS.compoundPath.en.replace(/:$/, ":"), compoundPathSel, compoundPathAll);
                    wPair(LABELS.compoundShape.ja.replace(/：$/, ":"), LABELS.compoundShape.en.replace(/:$/, ":"), compoundShapeSel, compoundShapeAll);

                    file.close();

                    if (lang === "ja") {
                        alert("書き出しました: " + desktopPath);
                    } else {
                        alert("Exported: " + desktopPath);
                    }
                } else {
                    if (lang === "ja") {
                        alert("ファイルを開けませんでした。");
                    } else {
                        alert("Failed to open the file.");
                    }
                }
            } catch (err) {
                if (lang === "ja") {
                    alert("書き出し中にエラー: " + err);
                } else {
                    alert("Export error: " + err);
                }
            }
        };
        btnOK.onClick = function () {
            dlg.close();
        };

        /* ダイアログ位置と透明度を調整 / Adjust dialog position and opacity */
        var offsetX = 300;
        var dialogOpacity = 0.97;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function () {
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