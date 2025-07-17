#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

MakeRectangleFromGuides.jsx

### 概要

- ドキュメント内のガイド（ルーラーガイド含む）の交点を基準に長方形を自動生成します。
- 必要に応じてロックされたレイヤーも含めて処理できます。
- RGBおよびCMYKカラーモードに対応しています。
- 隣接する長方形を結合するオプションがあります。

### 主な機能

- ガイドから交点を検出して長方形を作成
- ロックされたレイヤーの一時解除・復元
- 隣接長方形の結合
- 作成後にシェイプ変換

### 処理の流れ

1. ガイドを収集・分類
2. 必要に応じてロックレイヤーを一時解除
3. 交差領域に長方形を生成
4. 必要なら結合処理を実施

### 更新履歴

- v1.0 (20250713) : 初期バージョン

### Script Name:

MakeRectangleFromGuides.jsx

### Overview

- Automatically creates rectangles based on intersections of guides (including ruler guides).
- Optionally includes locked layers during processing.
- Supports RGB and CMYK color modes.
- Includes an option to merge adjacent rectangles.

### Main Features

- Detect intersections from guides and create rectangles
- Temporarily unlock and restore locked layers
- Merge adjacent rectangles
- Convert to shape after creation

### Process Flow

1. Collect and classify guides
2. Temporarily unlock locked layers if needed
3. Create rectangles in intersection areas
4. Merge rectangles if option enabled

### Change Log

- v1.0 (20250713): Initial version

*/

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    dialogTitle: { ja: "ガイドから長方形を作成 " + SCRIPT_VERSION, en: "Create Rectangles from Guides " + SCRIPT_VERSION },
    panelTitle: { ja: "ガイドの種類", en: "Guide Type" },
    allGuides: { ja: "すべて", en: "All Guides" },
    rulerGuides: { ja: "ルーラーガイドのみ", en: "Ruler Guides Only" },
    includeLocked: { ja: "ロックされたレイヤーを含む", en: "Include Locked Layers" },
    mergeAdjacent: { ja: "隣り合う長方形を結合", en: "Merge Adjacent Rectangles" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" }
};

function main() {
    /* ユーティリティ: ロック解除かつ表示されているレイヤーを取得 / Utility: Get an unlocked and visible layer */
    function getUnlockedVisibleLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (!doc.layers[i].locked && doc.layers[i].visible) {
                return doc.layers[i];
            }
        }
        return null;
    }

    if (app.documents.length === 0) {
        return;
    }

    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var guidePanel = dlg.add("panel", undefined, LABELS.panelTitle[lang]);
    guidePanel.orientation = "column";
    guidePanel.alignChildren = "left";
    guidePanel.alignment = "fill";
    guidePanel.margins = [15, 20, 15, 10];
    var rbAll = guidePanel.add("radiobutton", undefined, LABELS.allGuides[lang]);
    var rbRuler = guidePanel.add("radiobutton", undefined, LABELS.rulerGuides[lang]);
    rbAll.value = true;

    var cbLocked = dlg.add("checkbox", undefined, LABELS.includeLocked[lang]);
    cbLocked.value = true;

    var cbMerge = dlg.add("checkbox", undefined, LABELS.mergeAdjacent[lang]);
    cbMerge.value = true;

    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignChildren = ["center", "center"];
    btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });

    if (dlg.show() !== 1) {
        return;
    }

    var mergeRects = cbMerge.value;
    var useRulerOnly = rbRuler.value;
    var includeLocked = cbLocked.value;

    var doc = app.activeDocument;
    var guides = [];
    var rectOpacity = 50; /* 不透明度（0〜100） / Opacity (0-100), change here if needed */

    /* アートボードサイズ取得 / Get artboard size */
    var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var abBounds = artboard.artboardRect;
    var abWidth = abBounds[2] - abBounds[0];
    var abHeight = abBounds[1] - abBounds[3];

    var lockedLayers = [];
    if (includeLocked) {
        /* ロックされたレイヤーのロックを一時的に解除 / Temporarily unlock locked layers */
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].locked) {
                doc.layers[i].locked = false;
                lockedLayers.push(doc.layers[i]);
            }
        }
    }

    /*
    ガイドを収集 (includeLocked=falseならロックガイド除外) / Collect guides (exclude guides on locked layers if includeLocked is false)
    */
    for (var i = 0; i < doc.pathItems.length; i++) {
        var p = doc.pathItems[i];
        if (
            p.guides &&
            !p.locked &&
            !p.hidden &&
            (!p.layer.locked || includeLocked)
        ) {
            if (useRulerOnly) {
                if (p.pathPoints.length == 2 && Math.abs(p.area) < 0.01) {
                    var pt1 = p.pathPoints[0].anchor;
                    var pt2 = p.pathPoints[1].anchor;
                    var guideLength = Math.sqrt(Math.pow(pt2[0] - pt1[0], 2) + Math.pow(pt2[1] - pt1[1], 2));
                    var isVertical = Math.abs(pt1[0] - pt2[0]) < 0.01;
                    var isHorizontal = Math.abs(pt1[1] - pt2[1]) < 0.01;
                    if (
                        (isVertical && guideLength > abHeight + 100) ||
                        (isHorizontal && guideLength > abWidth + 100)
                    ) {
                        guides.push(p);
                    }
                }
            } else {
                guides.push(p);
            }
        }
    }

    if (includeLocked) {
        /* 元に戻す: ロック状態を復元 / Restore locked state */
        for (var i = 0; i < lockedLayers.length; i++) {
            lockedLayers[i].locked = true;
        }
        /* アクティブレイヤーがロックまたは非表示の場合は、ロック解除かつ表示されているレイヤーを設定 / If active layer is locked or hidden, set unlocked visible layer as active */
        var targetLayer = doc.activeLayer;
        if (targetLayer.locked || !targetLayer.visible) {
            var newLayer = getUnlockedVisibleLayer(doc);
            if (newLayer) {
                doc.activeLayer = newLayer;
            } else {
                alert("ロック解除かつ表示されているレイヤーがありません。");
                return;
            }
        }
    }

    var verticals = [];
    var horizontals = [];

    /*
    ガイドを縦横に分類 / Classify guides into vertical and horizontal
    */
    for (var j = 0; j < guides.length; j++) {
        var g = guides[j];
        var pt1 = g.pathPoints[0].anchor;
        var pt2 = g.pathPoints[1].anchor;

        if (Math.abs(pt1[0] - pt2[0]) < 0.01) {
            verticals.push(pt1[0]);
        } else if (Math.abs(pt1[1] - pt2[1]) < 0.01) {
            horizontals.push(pt1[1]);
        }
    }

    if (verticals.length < 2 || horizontals.length < 2) {
        alert("縦ガイド2本以上、横ガイド2本以上が必要です。");
        return;
    }

    /* 小さい順にソート（縦は昇順、横は降順） / Sort ascending verticals, descending horizontals */
    verticals.sort(function(a, b) {
        return a - b;
    });
    horizontals.sort(function(a, b) {
        return b - a;
    });

    /* レイヤーのロック状態と表示状態を考慮し、描画対象レイヤーを取得 / Get target layer considering lock and visibility */
    var targetLayer = doc.activeLayer;
    if (targetLayer.locked || !targetLayer.visible) {
        var newLayer = getUnlockedVisibleLayer(doc);
        if (newLayer) {
            doc.activeLayer = newLayer;
            targetLayer = newLayer;
        } else {
            alert("ロック解除かつ表示されているレイヤーがありません。");
            return;
        }
    }

    /* 新規レイヤーを作成してアクティブにする / Create a new layer and make it active */
    var newDrawLayer = doc.layers.add();
    newDrawLayer.name = "Generated Rectangles";
    doc.activeLayer = newDrawLayer;

    var createdRects = [];

    /*
    交差点を基に長方形を作成 / Create rectangles based on intersections
    */
    for (var xi = 0; xi < verticals.length - 1; xi++) {
        for (var yi = 0; yi < horizontals.length - 1; yi++) {
            var rectLeft = verticals[xi];
            var rectRight = verticals[xi + 1];
            var rectTop = horizontals[yi];
            var rectBottom = horizontals[yi + 1];

            var rectWidth = rectRight - rectLeft;
            var rectHeight = rectTop - rectBottom;

            var rect = doc.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
            rect.stroked = false;
            rect.filled = true;

            if (doc.documentColorSpace == DocumentColorSpace.RGB) {
                var rgb = new RGBColor();
                rgb.red = 0;
                rgb.green = 0;
                rgb.blue = 0;
                rect.fillColor = rgb;
            } else {
                var cmyk = new CMYKColor();
                cmyk.cyan = 0;
                cmyk.magenta = 0;
                cmyk.yellow = 0;
                cmyk.black = 100;
                rect.fillColor = cmyk;
            }
            rect.opacity = rectOpacity;
            rect.selected = true;
            app.executeMenuCommand('Convert to Shape');
            createdRects.push(rect);
        }
    }

    if (mergeRects && createdRects.length > 1) {
        for (var i = 0; i < createdRects.length; i++) {
            createdRects[i].selected = true;
        }
        app.executeMenuCommand('group');
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');
        app.executeMenuCommand('ungroup');
        app.executeMenuCommand('Convert to Shape');
    }
}

main();