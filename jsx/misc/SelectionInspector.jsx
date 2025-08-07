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

*/

var SCRIPT_VERSION = "v1.2";

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
        ja: "その他",
        en: "Other"
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
    chars: {
        ja: "文字数：",
        en: "Characters:"
    },
    paras: {
        ja: "段落数：",
        en: "Paragraphs:"
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
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    export: {
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

        /* 選択分をカウント / Count selected items */
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].typename === "TextFrame") {
                textCountSel++;
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

        // 右カラム
        var rightCol = twoColGroup.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["fill", "top"];

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
                } catch (err) {}
                try {
                    paraCountSel += sel[i].paragraphs.length;
                } catch (err) {}
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
                } catch (err) {}
                try {
                    paraCountAll += obj.paragraphs.length;
                } catch (err) {}
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
        addTextRow(LABELS.texts[lang], textCountSel + " / " + textCountAll);
        addTextRow(LABELS.pointText[lang], pointTextSel + " / " + pointTextAll);
        addTextRow(LABELS.areaText[lang], areaTextSel + " / " + areaTextAll);
        addTextRow(LABELS.pathText[lang], pathTextSel + " / " + pathTextAll);

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
        var btnExport = btnRightGroup.add("button", undefined, LABELS.export[lang], {
            name: "preview"
        });
        var btnOK = btnRightGroup.add("button", undefined, LABELS.ok[lang], {
            name: "ok"
        });

        btnExport.onClick = function() {
            try {
                var docName = app.activeDocument.name.replace(/\.[^\.]+$/, ""); // 拡張子を除いたファイル名
                var today = new Date();
                var yyyy = today.getFullYear();
                var mm = ("0" + (today.getMonth() + 1)).slice(-2);
                var dd = ("0" + today.getDate()).slice(-2);
                var dateStr = yyyy + mm + dd;

                var desktopPath = Folder.desktop + "/count-" + docName + "-" + dateStr + ".txt";
                var file = new File(desktopPath);
                if (file.open("w")) {
                    file.writeln("Selection Inspector Report (All Counts)");
                    file.writeln("-------------------------------");
                    file.writeln("");

                    if (lang === "ja") {
                        file.writeln("その他");
                        file.writeln(LABELS.artboards.ja + " " + artboardCount);
                        file.writeln(LABELS.objects.ja + " " + totalObjAll);
                        file.writeln("");

                        file.writeln("文字");
                        file.writeln(LABELS.texts.ja + " " + textCountAll);
                        file.writeln(LABELS.chars.ja + " " + totalCharAll);
                        file.writeln("");

                        file.writeln("テキストオブジェクト");
                        file.writeln(LABELS.paras.ja + " " + paraCountAll);
                        file.writeln(LABELS.pointText.ja + " " + pointTextAll);
                        file.writeln(LABELS.areaText.ja + " " + areaTextAll);
                        file.writeln(LABELS.pathText.ja + " " + pathTextAll);
                        file.writeln("");

                        file.writeln("画像");
                        file.writeln(LABELS.linked.ja + " " + linkedAll);
                        file.writeln(LABELS.embed.ja + " " + embeddedAll);
                        file.writeln(LABELS.broken.ja + " " + brokenLinkAll);
                        file.writeln("");

                        file.writeln("グループ");
                        file.writeln(LABELS.group.ja + " " + groupAll);
                        file.writeln(LABELS.clipGroup.ja + " " + clipGroupAll);
                        file.writeln("");

                        file.writeln("パス");
                        file.writeln(LABELS.pathCount.ja + " " + pathCountAll);
                        file.writeln(LABELS.openPath.ja + " " + openPathAll);
                        file.writeln(LABELS.closedPath.ja + " " + closedPathAll);
                        file.writeln(LABELS.anchors.ja + " " + anchorCountAll);
                    } else {
                        file.writeln("Misc");
                        file.writeln("Artboards: " + artboardCount);
                        file.writeln("Objects: " + totalObjAll);
                        file.writeln("");

                        file.writeln("Character");
                        file.writeln("Text Frames: " + textCountAll);
                        file.writeln("Characters: " + totalCharAll);
                        file.writeln("");

                        file.writeln("Text Objects");
                        file.writeln("Paragraphs: " + paraCountAll);
                        file.writeln("Point Text: " + pointTextAll);
                        file.writeln("Area Text: " + areaTextAll);
                        file.writeln("Path Text: " + pathTextAll);
                        file.writeln("");

                        file.writeln("Images");
                        file.writeln("Linked Images: " + linkedAll);
                        file.writeln("Embedded Images: " + embeddedAll);
                        file.writeln("Broken Links: " + brokenLinkAll);
                        file.writeln("");

                        file.writeln("Groups");
                        file.writeln("Groups: " + groupAll);
                        file.writeln("Clipping Groups: " + clipGroupAll);
                        file.writeln("");

                        file.writeln("Paths");
                        file.writeln("Paths: " + pathCountAll);
                        file.writeln("Open Paths: " + openPathAll);
                        file.writeln("Closed Paths: " + closedPathAll);
                        file.writeln("Anchor Points: " + anchorCountAll);
                    }

                    file.close();
                    alert("書き出しました: " + desktopPath);
                } else {
                    alert("ファイルを開けませんでした。");
                }
            } catch (err) {
                alert("書き出し中にエラー: " + err);
            }
        };
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