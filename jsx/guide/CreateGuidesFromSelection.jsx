#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト・アートボード・カンバス（擬似）を基準にガイドを作成するスクリプト。

- 上下左右の辺＋中心、各種プリセットをダイアログで指定して描画
- 基準は「アートボード」または「カンバス（擬似）」を選択
- アートボード基準時はガイドを外側へ「延長」、対象基準時は「選択オブジェクトとのマージン」を指定可能
- グループ（選択全体の外接でまとめて1組）／オブジェクトごとを切り替え（選択1つ以下のときは無効）
- プレビュー境界／幾何境界の切り替え。プレビュー境界時はテキストを一時アウトライン化＋アピアランス展開して正確な境界を取得（計算後に元へ復元。ライブプレビューでは省略）
- 「_guide」レイヤーの既存ガイドを削除するオプション
- ライブプレビュー（専用レイヤーに色付きの仮線で表示し、表示中は既存ガイドを一時非表示。確定時に「_guide」レイヤーへ本物のガイドを作成）
- 単位は環境設定の rulerType を参照

### 処理の流れ

- ダイアログでオプションを設定（変更に追従してライブプレビュー）
- 基準（選択／アートボード／カンバス）の外接矩形を取得
- 「_guide」レイヤーにガイドを追加してロック

### 紹介記事（note）

https://note.com/dtp_tranist/n/nd1359cf41a2c

### 更新履歴

- v1.0 (20250711)：初期バージョン
- v1.8.0 (20250802)：UIの整理、文言とツールチップの調整
- v1.9.0 (20260628)：ライブプレビュー追加、ローカライズ構造化、IIFE化、プレビュー中は既存ガイドを非表示

*/

/*

### Overview

Creates guides based on selected objects, the active artboard, or the pseudo canvas.

- Specify the four edges + center and various presets in the dialog
- Choose the target: "Artboard" or "Canvas (Pseudo)"
- Extend guides outward for the artboard target, or set a margin from the selection for object targets
- Toggle group (one set for the combined bounds) vs. per-object (disabled when 0–1 objects are selected)
- Toggle preview vs. geometric bounds; with preview bounds, text is temporarily outlined and appearance-expanded for accurate bounds, then restored (skipped during live preview)
- Option to delete existing guides in the "_guide" layer
- Live preview (colored temporary lines on a dedicated layer; existing guides are hidden while open, real guides are created on the "_guide" layer on commit)
- Units follow the rulerType preference

### Workflow

- Configure options in the dialog (live preview follows changes)
- Get the bounding box of the target (selection / artboard / canvas)
- Add guides to the "_guide" layer and lock it

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CreateGuidesFromSelection";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.9.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 右側プリセットラジオの表示フラグ / Show flags for the preset radio buttons */
var showRbTopBottom = false;
var showRbLeftRight = false;
var showRbTopLeft = true;
var showRbBottomLeft = true;
var showRbTopRight = false;
var showRbBottomRight = false;

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在の言語を判定 / Detect current language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "選択オブジェクトからガイド作成", en: "Create Guides from Selection" }
    },
    panel: {
        target: { ja: "対象", en: "Target" },
        preset: { ja: "プリセット", en: "Presets" },
        options: { ja: "オプション", en: "Options" },
        axis: { ja: "ガイド位置（辺と中央）", en: "Guide Positions (Edges & Center)" }
    },
    radio: {
        artboard: { ja: "アートボード", en: "Artboard" },
        canvas: { ja: "カンバス（擬似）", en: "Canvas (Pseudo)" },
        allOn: { ja: "すべて", en: "All" },
        edges: { ja: "四辺", en: "Edges" },
        vertical: { ja: "上下", en: "Top & Bottom" },
        horizontal: { ja: "左右", en: "Left & Right" },
        topLeft: { ja: "左上", en: "Top Left" },
        bottomLeft: { ja: "左下", en: "Bottom Left" },
        topRight: { ja: "右上", en: "Top Right" },
        bottomRight: { ja: "右下", en: "Bottom Right" },
        centerBoth: { ja: "中心", en: "Center" },
        centerVertical: { ja: "中心線（垂直）", en: "Center line (vertical)" },
        centerHorizontal: { ja: "中心線（水平）", en: "Center line (horizontal)" },
        clear: { ja: "クリア", en: "Clear" }
    },
    checkbox: {
        left: { ja: "左", en: "Left" },
        top: { ja: "上", en: "Top" },
        center: { ja: "中心", en: "Center" },
        bottom: { ja: "下", en: "Bottom" },
        right: { ja: "右", en: "Right" },
        usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" },
        deleteGuides: { ja: "「_guide」レイヤーのガイドを削除", en: "Delete guides in \"_guide\"" },
        group: { ja: "グループとしてガイドを作成", en: "Create guides as a group" },
        preview: { ja: "プレビュー", en: "Preview" }
    },
    field: {
        extension: { ja: "延長", en: "Extend" },
        offset: { ja: "選択オブジェクトとのマージン", en: "Margin from selection" }
    },
    button: {
        draw: { ja: "ガイドを描画", en: "Draw Guides" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
        expandError: {
            ja: "アピアランス展開中にエラーが発生しました。",
            en: "An error occurred while expanding appearance."
        },
        noArtboard: { ja: "アートボードが存在しません。", en: "No artboard exists." },
        invalidArtboard: { ja: "有効なアートボードが選択されていません。", en: "No valid artboard selected." },
        deleteGuideError: {
            ja: "既存ガイド削除時にエラーが発生しました。",
            en: "An error occurred while deleting existing guides."
        },
        guideError: {
            ja: "ガイド作成中にエラーが発生しました。",
            en: "An error occurred while creating guides."
        },
        noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." }
    },
    tooltip: {
        canvas: {
            ja: "アートボード範囲ではなく、広いカンバス範囲にガイドを引きます。",
            en: "Draw guides across a wide pseudo-canvas area instead of the active artboard."
        },
        artboard: {
            ja: "アクティブなアートボードの範囲に合わせてガイドを引きます。",
            en: "Draw guides within the active artboard area."
        },
        extension: {
            ja: "アートボード基準時に、ガイド線をアートボード外へ伸ばす量です。カンバス基準では使用しません。",
            en: "Extends guide lines beyond the artboard when using the artboard target. Not used for the canvas target."
        },
        offset: {
            ja: "対象の外側へガイドを離す距離です。0の場合は対象の辺・中心に作成します。",
            en: "Distance to move guides away from the target. Use 0 to place them on the target edges or center."
        },
        usePreviewBounds: {
            ja: "線幅や効果など、見た目上の境界を基準にします。テキストは一時的にアウトライン化して計算します（ライブプレビューでは省略）。",
            en: "Use visual bounds including strokes and effects. Text is temporarily outlined for calculation (skipped during live preview)."
        },
        deleteGuides: {
            ja: "実行前に「_guide」レイヤー内の既存ガイドを削除します。ほかのレイヤーのガイドは対象外です。",
            en: "Before drawing, delete existing guides in the \"_guide\" layer only. Guides on other layers are not affected."
        },
        group: {
            ja: "ON：選択全体の外接でまとめて1組のガイドを作成。OFF：各オブジェクトごとに作成。",
            en: "On: one set of guides for the combined selection bounds. Off: separate guides per object."
        },
        preview: {
            ja: "確定前に、作成予定のガイド位置を色付きの仮線で表示します。",
            en: "Show colored temporary lines for the guide positions before committing."
        },
        edge: {
            left: { ja: "対象の左辺にガイドを作成", en: "Add a guide at the target's left edge" },
            top: { ja: "対象の上辺にガイドを作成", en: "Add a guide at the target's top edge" },
            center: { ja: "対象の中央（縦・横）にガイドを作成", en: "Add guides at the target's center (vertical & horizontal)" },
            bottom: { ja: "対象の下辺にガイドを作成", en: "Add a guide at the target's bottom edge" },
            right: { ja: "対象の右辺にガイドを作成", en: "Add a guide at the target's right edge" }
        },
        preset: {
            allOn: { ja: "四辺＋中央をすべて選択", en: "Select all four edges and the center" },
            edges: { ja: "上下左右の四辺を選択", en: "Select all four edges" },
            topBottom: { ja: "上下の辺のみ選択", en: "Top and bottom edges only" },
            leftRight: { ja: "左右の辺のみ選択", en: "Left and right edges only" },
            topLeft: { ja: "左上（左＋上）を選択", en: "Top-left (left + top)" },
            bottomLeft: { ja: "左下（左＋下）を選択", en: "Bottom-left (left + bottom)" },
            topRight: { ja: "右上（右＋上）を選択", en: "Top-right (right + top)" },
            bottomRight: { ja: "右下（右＋下）を選択", en: "Bottom-right (right + bottom)" },
            centerBoth: { ja: "中央（縦横）のみ選択", en: "Center only (vertical & horizontal)" },
            centerVertical: { ja: "垂直の中心線のみ作成", en: "Vertical center line only" },
            centerHorizontal: { ja: "水平の中心線のみ作成", en: "Horizontal center line only" },
            clear: { ja: "すべて解除", en: "Clear all" }
        }
    }
};

/* LABELS からドット区切りのキーで文言を取得 / Resolve a label by dotted key */
function getLocalizedText(key) {
    var parts = key.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return key;
        node = node[parts[i]];
    }
    if (node == null) return key;
    return node[currentLanguage] || node.en || key;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLocalizedText(key) + (currentLanguage === "ja" ? "：" : ":");
}

/* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
function labelWithCount(key, count) {
    if (currentLanguage === "ja") {
        return getLocalizedText(key) + "（" + count + "）";
    }
    return getLocalizedText(key) + " (" + count + ")";
}

// =========================================
// 単位 / Units
// =========================================

/* 単位コードとラベルのマップ / Map of unit code to label */
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/* 現在の単位ラベルを取得 / Get current ruler unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

// =========================================
// UI ヘルパー / UI helpers
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

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


/* 「_guide」レイヤーを取得（無ければ作成）/ Get or create the "_guide" layer */
function getOrCreateGuideLayer() {
    var doc = app.activeDocument;
    var layer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "_guide") {
            layer = doc.layers[i];
            break;
        }
    }
    if (!layer) {
        layer = doc.layers.add();
        layer.name = "_guide";
    }
    return layer;
}

/* アクティブなアートボードの矩形を取得（無効ならアラートして null）/ Active artboard rect (alerts and returns null if invalid) */
function getActiveArtboardRect(doc) {
    if (doc.artboards.length === 0) {
        alert(getLocalizedText("alert.noArtboard"));
        return null;
    }
    var index = doc.artboards.getActiveArtboardIndex();
    if (index < 0 || index >= doc.artboards.length) {
        alert(getLocalizedText("alert.invalidArtboard"));
        return null;
    }
    return doc.artboards[index].artboardRect;
}

/* 複数オブジェクトの外接矩形を統合 / Union the bounds of multiple items */
function unionBounds(items, usePreviewBounds) {
    var bounds = null;
    for (var i = 0; i < items.length; i++) {
        var b = getItemBounds(items[i], usePreviewBounds);
        if (!bounds) {
            bounds = b.concat();
        } else {
            bounds[0] = Math.min(bounds[0], b[0]);
            bounds[1] = Math.max(bounds[1], b[1]);
            bounds[2] = Math.max(bounds[2], b[2]);
            bounds[3] = Math.min(bounds[3], b[3]);
        }
    }
    return bounds;
}

/* プレビュー境界使用時、テキストを一時アウトライン化して境界計算用に差し替え（restoreで復元）/ Temporarily outline text for bounds; returns { items, restore() } */
function outlineTextForBounds(selItems, usePreviewBounds) {
    var passthrough = { items: selItems, restore: function () {} };
    if (!usePreviewBounds) return passthrough;

    var tempCopies = [];
    var originalTexts = [];
    for (var i = 0; i < selItems.length; i++) {
        var item = selItems[i];
        if (item && item.typename === "TextFrame") {
            tempCopies.push(item.duplicate());
            item.hidden = true;
            originalTexts.push(item);
        }
    }
    if (tempCopies.length === 0) return passthrough;

    /* 全選択解除は app.selection = null が確実（doc.selection = null は MRAP エラーになることがある）/ Use app.selection to deselect */
    app.selection = null;
    for (var j = 0; j < tempCopies.length; j++) tempCopies[j].selected = true;
    try {
        app.executeMenuCommand('expandStyle');
    } catch (e) {
        alert(getLocalizedText("alert.expandError") + "\n" + e.message);
    }

    var outlinedItems = [];
    for (var k = 0; k < tempCopies.length; k++) {
        var outlined = tempCopies[k].createOutline();
        outlinedItems.push(outlined ? outlined : tempCopies[k]);
    }
    return {
        items: outlinedItems,
        restore: function () {
            for (var m = 0; m < outlinedItems.length; m++) {
                outlinedItems[m].remove();
                originalTexts[m].hidden = false;
            }
        }
    };
}

/* 基準矩形からガイドを描画（描画する方向は directionsFromBounds が算出）/ Draw guides from a reference bounds rect */
function drawGuidesForBounds(layer, bounds, options, useCanvas, offsetValue, marginValue) {
    var ab = null;
    if (!useCanvas) {
        ab = getActiveArtboardRect(app.activeDocument);
        if (!ab) return;
    }
    var directions = directionsFromBounds(bounds, options, offsetValue);
    for (var i = 0; i < directions.length; i++) {
        createGuide(layer, directions[i].pos, directions[i].orientation, useCanvas, marginValue, ab);
    }
}

/* 選択オブジェクト・アートボード・カンバスからガイドを作成（本処理）/ Create guides from the selection, artboard, or canvas (commit) */
/* 引数: options=各方向フラグ, useCanvas=カンバス基準, offsetValue=距離(pt), marginValue=延長(pt) / Args as named */
function createGuidesFromSelection(options, useCanvas, offsetValue, marginValue) {
    var doc = app.activeDocument;
    var selItems = doc.selection;
    var layer, wasLocked;

    /* オブジェクトごとにガイドを作成 / Per-object guides */
    if (options.individual && selItems.length > 0) {
        layer = getOrCreateGuideLayer();
        wasLocked = layer.locked;
        if (wasLocked) layer.locked = false;
        for (var i = 0; i < selItems.length; i++) {
            drawGuidesForBounds(layer, getItemBounds(selItems[i], options.usePreviewBounds), options, useCanvas, offsetValue, marginValue);
        }
        layer.locked = true;
        return;
    }

    /* 基準となる外接矩形を決定 / Determine the reference bounds */
    var bounds;
    var textState = null;
    if (selItems.length === 0) {
        bounds = getActiveArtboardRect(doc);
        if (!bounds) return; // アラートは getActiveArtboardRect 内 / alerted inside
        bounds = bounds.concat();
    } else {
        textState = outlineTextForBounds(selItems, options.usePreviewBounds);
        bounds = unionBounds(textState.items, options.usePreviewBounds);
    }

    layer = getOrCreateGuideLayer();
    wasLocked = layer.locked;
    if (wasLocked) layer.locked = false;

    drawGuidesForBounds(layer, bounds, options, useCanvas, offsetValue, marginValue);

    if (textState) textState.restore();
    layer.locked = true;
}

/* 指定位置・向きにガイドを1本作成（線分は guideSegment が算出）/ Create a single guide (segment from guideSegment) */
function createGuide(layer, pos, orientation, useCanvas, marginValue, ab) {
    var guide = layer.pathItems.add();
    guide.setEntirePath(guideSegment(pos, orientation, useCanvas, marginValue, ab));
    guide.filled = false;
    guide.stroked = false;
    guide.guides = true;
}

/* 単位コード→pt 換算係数を取得 / Get the points-per-unit factor for a unit code */
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0; // in
        case 1:
            return 72.0 / 25.4; // mm
        case 2:
            return 1.0; // pt
        case 3:
            return 12.0; // pica
        case 4:
            return 72.0 / 2.54; // cm
        case 5:
            return 72.0 / 25.4 * 0.25; // Q or H
        case 6:
            return 1.0; // px
        case 7:
            return 72.0 * 12.0; // ft/in
        case 8:
            return 72.0 / 25.4 * 1000.0; // m
        case 9:
            return 72.0 * 36.0; // yd
        case 10:
            return 72.0 * 12.0; // ft
        default:
            return 1.0;
    }
}

/* ===== ライブプレビュー / Live preview ===== */

/* プレビュー用レイヤー名 / Preview layer name */
var PREVIEW_LAYER_NAME = "__GuidePreview__";

/* プレビュー用の色（ドキュメントのカラースペースに合わせる）/ Preview color (matches doc color space) */
function makePreviewColor(doc) {
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
        var cmyk = new CMYKColor();
        cmyk.cyan = 0; cmyk.magenta = 90; cmyk.yellow = 0; cmyk.black = 0;
        return cmyk;
    }
    var rgb = new RGBColor();
    rgb.red = 255; rgb.green = 0; rgb.blue = 255;
    return rgb;
}

/* プレビュー用レイヤーを削除 / Remove the preview layer if present */
function removePreviewLayer(doc) {
    try {
        var layer = doc.layers.getByName(PREVIEW_LAYER_NAME);
        layer.locked = false;
        layer.visible = true;
        layer.remove();
    } catch (e) {}
}

/* プレビュー用レイヤーを用意（既存は作り直し）/ Create a fresh preview layer */
function createPreviewLayer(doc) {
    removePreviewLayer(doc);
    var layer = doc.layers.add();
    layer.name = PREVIEW_LAYER_NAME;
    return layer;
}

/* 線分配列を色付きプレビュー線としてレイヤーへ描画 / Draw colored preview lines into a layer */
function addPreviewSegments(layer, segments, color) {
    for (var i = 0; i < segments.length; i++) {
        var path = layer.pathItems.add();
        path.setEntirePath([segments[i][0], segments[i][1]]);
        path.filled = false;
        path.stroked = true;
        path.strokeColor = color;
        path.strokeWidth = 1;
    }
}

/* オブジェクトの境界（クリップグループ対応）/ Bounds of an item (clip-group aware) */
function getItemBounds(item, usePreview) {
    if (item.typename === "GroupItem" && item.clipped) {
        for (var j = 0; j < item.pageItems.length; j++) {
            if (item.pageItems[j].clipping) {
                return usePreview ? item.pageItems[j].visibleBounds : item.pageItems[j].geometricBounds;
            }
        }
    }
    return usePreview ? item.visibleBounds : item.geometricBounds;
}

/* 境界とオプションから引くべき方向（位置・向き）を算出 / Directions to draw from bounds */
function directionsFromBounds(bounds, options, offsetPt) {
    var top = bounds[1] + offsetPt;
    var left = bounds[0] - offsetPt;
    var bottom = bounds[3] - offsetPt;
    var right = bounds[2] + offsetPt;
    var centerX = (left + right) / 2;
    var centerY = (top + bottom) / 2;

    var dirs = [];
    if (options.left) dirs.push({ pos: left, orientation: "vertical" });
    if (options.right) dirs.push({ pos: right, orientation: "vertical" });
    if (options.top) dirs.push({ pos: top, orientation: "horizontal" });
    if (options.bottom) dirs.push({ pos: bottom, orientation: "horizontal" });
    if (options.center) {
        dirs.push({ pos: centerX, orientation: "vertical" });
        dirs.push({ pos: centerY, orientation: "horizontal" });
    }
    if (options.centerMode === "vertical") dirs.push({ pos: centerX, orientation: "vertical" });
    if (options.centerMode === "horizontal") dirs.push({ pos: centerY, orientation: "horizontal" });
    return dirs;
}

/* ガイド1本の線分（[start, end]）/ One guide segment */
function guideSegment(pos, orientation, useCanvas, marginPt, ab) {
    if (useCanvas) {
        var big = 8000;
        return (orientation === "horizontal") ? [[-big, pos], [big, pos]] : [[pos, big], [pos, -big]];
    }
    return (orientation === "horizontal")
        ? [[ab[0] - marginPt, pos], [ab[2] + marginPt, pos]]
        : [[pos, ab[1] + marginPt], [pos, ab[3] - marginPt]];
}

/* 現在の設定からプレビュー線分を収集（本処理の幾何を簡易再現。テキストのアウトライン化は省略）/ Collect preview segments */
function collectPreviewSegments(options, useCanvas, offsetPt, marginPt) {
    var doc = app.activeDocument;
    var segments = [];

    var ab = null;
    if (doc.artboards.length > 0) {
        var abIndex = doc.artboards.getActiveArtboardIndex();
        if (abIndex >= 0 && abIndex < doc.artboards.length) ab = doc.artboards[abIndex].artboardRect;
    }
    if (!useCanvas && !ab) return segments;

    var selItems = doc.selection;

    function pushDirs(bounds) {
        var dirs = directionsFromBounds(bounds, options, offsetPt);
        for (var d = 0; d < dirs.length; d++) {
            segments.push(guideSegment(dirs[d].pos, dirs[d].orientation, useCanvas, marginPt, ab));
        }
    }

    if (options.individual && selItems.length > 0) {
        for (var i = 0; i < selItems.length; i++) {
            pushDirs(getItemBounds(selItems[i], options.usePreviewBounds));
        }
    } else if (selItems.length === 0) {
        if (ab) pushDirs([ab[0], ab[1], ab[2], ab[3]]);
    } else {
        var bounds = null;
        for (var i = 0; i < selItems.length; i++) {
            var b = getItemBounds(selItems[i], options.usePreviewBounds);
            if (!bounds) {
                bounds = b.concat();
            } else {
                bounds[0] = Math.min(bounds[0], b[0]);
                bounds[1] = Math.max(bounds[1], b[1]);
                bounds[2] = Math.max(bounds[2], b[2]);
                bounds[3] = Math.min(bounds[3], b[3]);
            }
        }
        if (bounds) pushDirs(bounds);
    }
    return segments;
}

/* プリセットのラジオボタンを生成・配線（cross=十字チェックボックス, centerLine={mode}）/ Build & wire preset radios */
function buildPresetRadios(parent, cross, centerLine) {
    /* ラジオボタン生成（tooltipKey で説明、show が false なら作らない）/ Create a radio (with tooltip; skip when show is false) */
    function createRadioButton(label, tooltipKey, show) {
        if (typeof show !== "undefined" && !show) return null;
        var radio = parent.add("radiobutton", undefined, label);
        if (tooltipKey) radio.helpTip = getLocalizedText(tooltipKey);
        return radio;
    }

    var radios = {
        allOn: createRadioButton(getLocalizedText("radio.allOn"), "tooltip.preset.allOn"),
        edges: createRadioButton(getLocalizedText("radio.edges"), "tooltip.preset.edges"),
        topBottom: createRadioButton(getLocalizedText("radio.vertical"), "tooltip.preset.topBottom", showRbTopBottom),
        leftRight: createRadioButton(getLocalizedText("radio.horizontal"), "tooltip.preset.leftRight", showRbLeftRight),
        topLeft: createRadioButton(getLocalizedText("radio.topLeft"), "tooltip.preset.topLeft", showRbTopLeft),
        bottomLeft: createRadioButton(getLocalizedText("radio.bottomLeft"), "tooltip.preset.bottomLeft", showRbBottomLeft),
        topRight: createRadioButton(getLocalizedText("radio.topRight"), "tooltip.preset.topRight", showRbTopRight),
        bottomRight: createRadioButton(getLocalizedText("radio.bottomRight"), "tooltip.preset.bottomRight", showRbBottomRight),
        centerBoth: createRadioButton(getLocalizedText("radio.centerBoth"), "tooltip.preset.centerBoth"),
        centerVertical: createRadioButton(getLocalizedText("radio.centerVertical"), "tooltip.preset.centerVertical"),
        centerHorizontal: createRadioButton(getLocalizedText("radio.centerHorizontal"), "tooltip.preset.centerHorizontal"),
        clear: createRadioButton(getLocalizedText("radio.clear"), "tooltip.preset.clear")
    };

    /* 十字チェックボックス（左/上/右/下/中央）の状態をまとめて設定 / Set the cross checkboxes at once */
    function setCrossState(showLeft, showTop, showRight, showBottom, showCenter) {
        centerLine.mode = "";
        cross.left.value = showLeft;
        cross.top.value = showTop;
        cross.right.value = showRight;
        cross.bottom.value = showBottom;
        cross.center.value = showCenter;
    }

    /* プリセット定義（cross=[左,上,右,下,中央], centerLine=中心線モード）/ Preset table */
    var presetDefs = [
        { rb: radios.allOn,            cross: [true,  true,  true,  true,  true ] },
        { rb: radios.edges,            cross: [true,  true,  true,  true,  false] },
        { rb: radios.topBottom,        cross: [false, true,  false, true,  false] },
        { rb: radios.leftRight,        cross: [true,  false, true,  false, false] },
        { rb: radios.topLeft,          cross: [true,  true,  false, false, false] },
        { rb: radios.bottomLeft,       cross: [true,  false, false, true,  false] },
        { rb: radios.topRight,         cross: [false, true,  true,  false, false] },
        { rb: radios.bottomRight,      cross: [false, false, true,  true,  false] },
        { rb: radios.clear,            cross: [false, false, false, false, false] },
        { rb: radios.centerBoth,       cross: [false, false, false, false, true ] },
        { rb: radios.centerVertical,   cross: [false, false, false, false, false], centerLine: "vertical" },
        { rb: radios.centerHorizontal, cross: [false, false, false, false, false], centerLine: "horizontal" }
    ];
    for (var i = 0; i < presetDefs.length; i++) {
        (function (def) {
            if (!def.rb) return;
            def.rb.onClick = function () {
                if (!def.rb.value) return;
                setCrossState(def.cross[0], def.cross[1], def.cross[2], def.cross[3], def.cross[4]);
                if (def.centerLine) centerLine.mode = def.centerLine;
            };
        })(presetDefs[i]);
    }

    /* デフォルト選択は四辺を優先 / Default selection (prefer "Edges") */
    if (radios.edges) {
        radios.edges.value = true;
    } else if (radios.allOn) {
        radios.allOn.value = true;
    }

    /* 手動でチェックを変えたら中心線モードを解除 / Clear center-line mode on manual checkbox toggle */
    function clearCenterLineOnToggle(checkbox) {
        checkbox.onClick = function () {
            centerLine.mode = "";
        };
    }
    clearCenterLineOnToggle(cross.left);
    clearCenterLineOnToggle(cross.top);
    clearCenterLineOnToggle(cross.right);
    clearCenterLineOnToggle(cross.bottom);
    clearCenterLineOnToggle(cross.center);

    return radios;
}

/* メインダイアログを構築して表示 / Build and show the main dialog */
function buildDialog() {
    /* 中心線モードの保持（プリセットと共有・参照渡し用に holder）/ Center-line mode holder (shared with presets) */
    var centerLine = { mode: "" };
    var dialog = new Window("dialog", getLocalizedText("dialog.title") + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = ["center", "top"];
    dialog.spacing = 10;
    dialog.margins = 15;

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];
    mainGroup.spacing = 20;
    mainGroup.margins = 0;

    var leftGroup = mainGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = ["left", "top"];
    leftGroup.spacing = 10;

    var presetPanel = mainGroup.add("panel", undefined, getLocalizedText("panel.preset"));
    setupPanel(presetPanel);

    var rightGroup = presetPanel.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = ["left", "top"];
    rightGroup.spacing = 10;
    rightGroup.margins = 0;

    var targetPanel = leftGroup.add("panel", undefined, getLocalizedText("panel.target"));
    setupPanel(targetPanel);
    /* カンバス→アートボードの順にラジオボタンを追加し、デフォルトをアートボードに / Add radio buttons in order: Canvas → Artboard, default to Artboard */
    var rbCanvas = targetPanel.add("radiobutton", undefined, getLocalizedText("radio.canvas"));
    rbCanvas.helpTip = getLocalizedText("tooltip.canvas");
    var rbArtboard = targetPanel.add("radiobutton", undefined, getLocalizedText("radio.artboard"));
    rbArtboard.helpTip = getLocalizedText("tooltip.artboard");
    rbArtboard.value = true;
    /* 「延長」グループを targetPanel 内に追加 / Add extension group to targetPanel */
    var marginGroup = targetPanel.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = ["left", "center"];
    var marginLabel = marginGroup.add("statictext", undefined, labelText("field.extension"));
    marginLabel.helpTip = getLocalizedText("tooltip.extension");
    var marginInput = marginGroup.add("edittext", undefined, "20");
    marginInput.helpTip = getLocalizedText("tooltip.extension");
    marginInput.characters = 3;
    var marginUnitLabel = marginGroup.add("statictext", undefined, getCurrentUnitLabel());
    marginUnitLabel.helpTip = getLocalizedText("tooltip.extension");
    changeValueByArrowKey(marginInput);

    var axisGroup = leftGroup.add("panel", undefined, undefined, {
        name: "axisGroup"
    });
    axisGroup.text = getLocalizedText("panel.axis");
    setupPanel(axisGroup);

    var crossGroup = axisGroup.add("group", undefined, {
        name: "crossGroup"
    });
    crossGroup.orientation = "row";
    crossGroup.alignChildren = ["left", "center"];
    crossGroup.spacing = 20;
    crossGroup.margins = 0;
    /* パネル内で左右中央に配置 / Center the cross horizontally in the panel */
    crossGroup.alignment = ["center", "top"];

    var colLeft = crossGroup.add("group", undefined, {
        name: "colLeft"
    });
    colLeft.orientation = "row";
    colLeft.alignChildren = ["left", "center"];
    colLeft.spacing = 10;
    colLeft.margins = 0;
    var cbLeft = colLeft.add("checkbox", undefined, undefined, {
        name: "cbLeft"
    });
    cbLeft.text = getLocalizedText("checkbox.left");
    cbLeft.helpTip = getLocalizedText("tooltip.edge.left");
    cbLeft.value = true;

    var colCenter = crossGroup.add("group", undefined, {
        name: "colCenter"
    });

    colCenter.orientation = "column";
    colCenter.alignChildren = ["left", "center"];
    colCenter.spacing = 10;
    colCenter.margins = 0;
    var cbTop = colCenter.add("checkbox", undefined, undefined, {
        name: "cbTop"
    });
    cbTop.text = getLocalizedText("checkbox.top");
    cbTop.helpTip = getLocalizedText("tooltip.edge.top");
    cbTop.value = true;
    var cbCenter = colCenter.add("checkbox", undefined, undefined, {
        name: "cbCenter"
    });
    cbCenter.text = getLocalizedText("checkbox.center");
    cbCenter.helpTip = getLocalizedText("tooltip.edge.center");
    cbCenter.value = false;
    /* for horizontal center line / 水平中心線用 */
    var cbBottom = colCenter.add("checkbox", undefined, undefined, {
        name: "cbBottom"
    });
    cbBottom.text = getLocalizedText("checkbox.bottom");
    cbBottom.helpTip = getLocalizedText("tooltip.edge.bottom");
    cbBottom.value = true;

    var colRight = crossGroup.add("group", undefined, {
        name: "colRight"
    });
    colRight.orientation = "row";
    colRight.alignChildren = ["left", "center"];
    colRight.spacing = 10;
    colRight.margins = 0;
    var cbRight = colRight.add("checkbox", undefined, undefined, {
        name: "cbRight"
    });
    cbRight.text = getLocalizedText("checkbox.right");
    cbRight.helpTip = getLocalizedText("tooltip.edge.right");
    cbRight.value = true;

    /* プリセットのラジオを生成・配線 / Build and wire the preset radios */
    var crossCheckboxes = { left: cbLeft, top: cbTop, right: cbRight, bottom: cbBottom, center: cbCenter };
    var presetRadios = buildPresetRadios(rightGroup, crossCheckboxes, centerLine);

    var optionsGroup = dialog.add("panel", undefined, getLocalizedText("panel.options"), {
        name: "optionsGroup"
    });
    setupPanel(optionsGroup);

    var cbUsePreview = optionsGroup.add("checkbox", undefined, getLocalizedText("checkbox.usePreviewBounds"));
    cbUsePreview.value = true;
    cbUsePreview.helpTip = getLocalizedText("tooltip.usePreviewBounds");
    var cbDeleteGuide = optionsGroup.add("checkbox", undefined, getLocalizedText("checkbox.deleteGuides"));
    cbDeleteGuide.value = true;
    cbDeleteGuide.helpTip = getLocalizedText("tooltip.deleteGuides");

    /* 個別ガイドチェックボックス追加 / Add individual guide checkbox */
    var cbGroup = optionsGroup.add("checkbox", undefined, getLocalizedText("checkbox.group"));
    cbGroup.value = true;
    cbGroup.helpTip = getLocalizedText("tooltip.group");
    /* 選択が1つ以下なら個別/グループの区別が無いのでディム / Dim when 0–1 objects are selected */
    if (app.activeDocument.selection.length <= 1) {
        cbGroup.enabled = false;
    }

    /* オフセット入力欄追加 / Add offset input field */
    var offsetGroup = optionsGroup.add("group");
    offsetGroup.orientation = "row";
    offsetGroup.alignChildren = ["left", "center"];
    var offsetLabel = offsetGroup.add("statictext", undefined, labelText("field.offset"));
    offsetLabel.helpTip = getLocalizedText("tooltip.offset");
    offsetLabel.margins = [0, 0, 8, 0];
    var offsetInput = offsetGroup.add("edittext", undefined, "0");
    offsetInput.helpTip = getLocalizedText("tooltip.offset");
    offsetInput.characters = 3;
    var offsetUnitLabel = offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
    offsetUnitLabel.helpTip = getLocalizedText("tooltip.offset");
    offsetInput.active = true;
    changeValueByArrowKey(offsetInput);

    /* 「延長」行（ラベル＋入力＋単位）を「カンバス」選択時に丸ごとディム / Dim the whole extension row when canvas is selected */
    function updateMarginEnabled() {
        marginGroup.enabled = !rbCanvas.value;
    }
    rbArtboard.onClick = updateMarginEnabled;
    rbCanvas.onClick = updateMarginEnabled;
    /* 初期状態に合わせる / Set initial state */
    updateMarginEnabled();

    /* 選択オブジェクトがアートボード外にある場合、自動的にカンバス選択 / Auto-select Canvas if selection is outside artboard */
    var doc = app.activeDocument;
    var selItems = doc.selection;
    if (selItems.length > 0 && doc.artboards.length > 0) {
        var abIndex = doc.artboards.getActiveArtboardIndex();
        var ab = doc.artboards[abIndex].artboardRect;
        var abLeft = ab[0];
        var abTop = ab[1];
        var abRight = ab[2];
        var abBottom = ab[3];

        var allOutside = true;

        for (var i = 0; i < selItems.length; i++) {
            var itemBounds = selItems[i].geometricBounds;
            if (!(itemBounds[0] > abRight || itemBounds[2] < abLeft || itemBounds[1] < abBottom || itemBounds[3] > abTop)) {
                allOutside = false;
                break;
            }
        }

        if (allOutside) {
            rbCanvas.value = true;
            rbArtboard.value = false;
            updateMarginEnabled();
        }
    }

    /* ===== ボタン領域（3カラム：左=プレビュー / 中央=spacer / 右=ボタン）/ Footer (3 columns) ===== */
    var footerGroup = dialog.add("group");
    footerGroup.orientation = "row";
    footerGroup.alignment = "fill";
    footerGroup.alignChildren = ["fill", "center"];
    footerGroup.margins = [0, 10, 0, 0];

    /* 左：プレビュー / Left: preview */
    var footerLeft = footerGroup.add("group");
    footerLeft.orientation = "row";
    footerLeft.alignment = ["left", "center"];
    var cbPreview = footerLeft.add("checkbox", undefined, getLocalizedText("checkbox.preview"));
    cbPreview.value = true;
    cbPreview.helpTip = getLocalizedText("tooltip.preview");

    /* 中央：スペーサー（余白を吸収）/ Center: spacer (absorbs free space) */
    var footerSpacer = footerGroup.add("group");
    footerSpacer.alignment = ["fill", "center"];

    /* 右：ボタン（Mac 順：Cancel → OK）/ Right: buttons (Mac order) */
    var footerRight = footerGroup.add("group");
    footerRight.orientation = "row";
    footerRight.alignment = ["right", "center"];
    var btnCancel = footerRight.add("button", undefined, getLocalizedText("button.cancel"));
    var btnCreateGuides = footerRight.add("button", undefined, getLocalizedText("button.draw"), { name: "ok" });

    /* ===== プレビュー配線 / Preview wiring ===== */
    function previewToPoints(text) {
        var value = parseFloat(text);
        if (isNaN(value)) value = 0;
        return value * getPtFactorFromUnitCode(app.preferences.getIntegerPreference("rulerType"));
    }
    function readPreviewState() {
        return {
            options: {
                left: cbLeft.value,
                right: cbRight.value,
                top: cbTop.value,
                bottom: cbBottom.value,
                center: cbCenter.value,
                centerMode: centerLine.mode,
                usePreviewBounds: cbUsePreview.value,
                individual: !cbGroup.value
            },
            useCanvas: rbCanvas.value,
            offsetPt: previewToPoints(offsetInput.text),
            marginPt: previewToPoints(marginInput.text)
        };
    }
    function clearPreview() {
        removePreviewLayer(app.activeDocument);
    }
    /* 既存ガイドの表示/非表示（showguide はトグルなので状態を自前で追跡）/ Hide/show existing guides (showguide is a toggle; track state) */
    var guidesHidden = false;
    function setGuidesHidden(hide) {
        if (hide === guidesHidden) return;
        try { app.executeMenuCommand("showguide"); } catch (e) {}
        guidesHidden = hide;
    }
    function renderPreview() {
        clearPreview();
        /* プレビュー中は既存ガイドを隠し、仮ガイド（色付き線）だけ見せる / Hide real guides during preview */
        setGuidesHidden(cbPreview.value);
        if (cbPreview.value) {
            try {
                var state = readPreviewState();
                var segments = collectPreviewSegments(state.options, state.useCanvas, state.offsetPt, state.marginPt);
                if (segments.length > 0) {
                    var previewDoc = app.activeDocument;
                    var previewLayer = createPreviewLayer(previewDoc);
                    addPreviewSegments(previewLayer, segments, makePreviewColor(previewDoc));
                    previewLayer.locked = true;
                }
            } catch (e) {
                clearPreview();
            }
        }
        app.redraw();
    }
    /* 既存 onClick を保持しつつプレビュー更新を連結 / Chain renderPreview after existing onClick */
    function chainPreview(control) {
        if (!control) return;
        var previousOnClick = control.onClick;
        control.onClick = function () {
            if (previousOnClick) previousOnClick();
            renderPreview();
        };
    }
    var previewTriggers = [
        rbCanvas, rbArtboard,
        presetRadios.allOn, presetRadios.edges, presetRadios.topBottom, presetRadios.leftRight,
        presetRadios.topLeft, presetRadios.bottomLeft, presetRadios.topRight, presetRadios.bottomRight,
        presetRadios.centerBoth, presetRadios.centerVertical, presetRadios.centerHorizontal, presetRadios.clear,
        cbLeft, cbTop, cbRight, cbBottom, cbCenter,
        cbUsePreview, cbGroup, cbPreview
    ];
    for (var i = 0; i < previewTriggers.length; i++) {
        chainPreview(previewTriggers[i]);
    }
    offsetInput.onChanging = renderPreview;
    marginInput.onChanging = renderPreview;

    btnCancel.onClick = function () {
        dialog.close();
    };

    btnCreateGuides.onClick = function () {
        clearPreview(); // プレビューを片付けてから本処理 / clean up preview before committing
        try {
            var options = {
                left: cbLeft.value,
                right: cbRight.value,
                top: cbTop.value,
                bottom: cbBottom.value,
                center: cbCenter.value,
                centerMode: centerLine.mode,
                usePreviewBounds: cbUsePreview.value,
                individual: !cbGroup.value
            };
            var useCanvas = rbCanvas.value;

            /* _guideレイヤー取得または作成 / Get or create "_guide" layer */
            var layer = getOrCreateGuideLayer();
            var wasLocked = layer.locked;
            if (wasLocked) layer.locked = false;

            /* 削除チェックON時、既存ガイド削除 / Remove existing guides if checked */
            if (cbDeleteGuide.value) {
                try {
                    for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                        if (layer.pageItems[i].guides) {
                            layer.pageItems[i].remove();
                        }
                    }
                } catch (ex) {
                    alert(getLocalizedText("alert.deleteGuideError") + "\n" + ex.message);
                }
            }

            var offsetVal = parseFloat(offsetInput.text);
            if (isNaN(offsetVal)) offsetVal = 0;
            var marginVal = parseFloat(marginInput.text);
            if (isNaN(marginVal)) marginVal = 0;
            var unitCode = app.preferences.getIntegerPreference("rulerType");
            var ptFactor = getPtFactorFromUnitCode(unitCode);
            var offsetValPt = offsetVal * ptFactor;
            var marginValPt = marginVal * ptFactor;

            createGuidesFromSelection(options, useCanvas, offsetValPt, marginValPt);
            dialog.close();
        } catch (e) {
            alert(getLocalizedText("alert.guideError") + "\n" + (e && e.message ? e.message : e) + "\n" + (e && e.stack ? e.stack : ""));
        }
    };

    /* 表示時に初期プレビュー / Initial preview on show */
    dialog.onShow = function () {
        renderPreview();
    };
    /* 閉じる時は仮ガイドを片付け、隠した既存ガイドを再表示 / On close: clear preview and restore guide visibility */
    dialog.onClose = function () {
        clearPreview();
        setGuidesHidden(false);
        app.redraw();
    };

    dialog.show();
}

/* テキストフィールドで上下矢印キーによる数値増減を可能にする / Enable value change with up/down arrow keys in an edittext */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            if (keyboard.shiftKey) {
                /* Shift押下時は10の倍数スナップ / Snap to tens if Shift is pressed */
                value = Math.round(value / 10) * 10 + (event.keyName == "Up" ? 10 : -10);
            } else {
                var delta = event.keyName == "Up" ? 1 : -1;
                value += delta;
            }

            event.preventDefault();
            editText.text = value;
            // プログラム変更は onChanging を発火しないため明示的に呼ぶ / fire onChanging manually
            if (typeof editText.onChanging === "function") editText.onChanging();
        }
    });
}

/* エントリーポイント / Entry point */
(function main() {
    if (!app.documents.length) {
        alert(getLocalizedText("alert.noDocument"));
        return;
    }
    buildDialog();
})();