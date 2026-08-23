#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト・アートボード・カンバスを基準にガイドを作成します。
詳細はREADMEを参照してください。

### Overview

Creates guides based on the selection, the artboard, or the canvas.
See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CreateGuidesFromSelection";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.9.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-24";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGuidesFromSelection.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGuidesFromSelection.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd1359cf41a2c"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* ガイドを作成するレイヤー名 / Layer that receives the guides */
var GUIDE_LAYER_NAME = "_guide";

/* ライブプレビュー用の一時レイヤー名 / Temporary layer for the live preview */
var PREVIEW_LAYER_NAME = "__GuidePreview__";

/* カンバス基準時にガイドを伸ばす長さ（pt）/ Guide reach for the canvas target (pt) */
var CANVAS_GUIDE_REACH = 8000;

/* プリセットラジオボタンの表示 / Show flags for the preset radio buttons */
var showPresetTopBottom = false;
var showPresetLeftRight = false;
var showPresetTopLeft = true;
var showPresetBottomLeft = true;
var showPresetTopRight = false;
var showPresetBottomRight = false;

// =========================================
// レイアウト / Layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var CROSS_SPACING  = 20;                 /* 十字レイアウトの間隔 / gap inside the cross layout */
var ROW_SPACING    = 4;                  /* 行内の要素間隔 / gap inside a row */

/**
 * ウィンドウの共通設定
 * @param {Window} win - 対象ウィンドウ
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定
 * @param {Panel} panel - 対象パネル
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
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
 * 行グループの共通設定（ボタン列・入力行など）
 * @param {Group} group - 対象グループ
 * @param {string} [alignment] - グループ自身の整列（省略時は "left"）
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    /* alignment と alignChildren は対で指定する（片方だけだと天地がずれ、中のボタンが横に伸びる）
       Set both: alone, either one lets the row drift vertically or stretch its buttons */
    group.alignment = [alignment || "left", "center"];
    group.alignChildren = ["left", "center"];
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// ローカライズ / Localization
// =========================================

/**
 * 現在の言語を判定
 * @returns {string} "ja" または "en"
 */
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
        destination: { ja: "描画先", en: "Destination" },
        options: { ja: "オプション", en: "Options" },
        axis: { ja: "ガイド位置（辺と中央）", en: "Guide Positions (Edges & Center)" }
    },
    radio: {
        artboard: { ja: "アートボード", en: "Artboard" },
        canvas: { ja: "カンバス（擬似）", en: "Canvas (Pseudo)" },
        selectionLayer: { ja: "選択オブジェクトと同じレイヤー", en: "Same layer as the selection" },
        guideLayer: { ja: "「_guide」レイヤー", en: "\"_guide\" layer" },
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
        individual: { ja: "オブジェクトごとに作成", en: "Create per object" },
        group: { ja: "描画するガイドをグループ化", en: "Group the guides to draw" },
        preview: { ja: "プレビュー", en: "Preview" }
    },
    fieldLabel: {
        extension: { ja: "延長", en: "Extend" },
        offset: { ja: "選択オブジェクトとのマージン", en: "Margin from selection" }
    },
    button: {
        draw: { ja: "ガイドを描画", en: "Draw Guides" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        expandError: {
            ja: "アピアランス展開中にエラーが発生しました。",
            en: "An error occurred while expanding appearance."
        },
        noArtboard: { ja: "アートボードが存在しません。", en: "No artboard exists." },
        invalidArtboard: { ja: "有効なアートボードが選択されていません。", en: "No valid artboard selected." },
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
        selectionLayer: {
            ja: "選択オブジェクトと同じレイヤーにガイドを作成します。選択がないときはアクティブレイヤーに作成します。",
            en: "Create the guides on the same layer as the selection. With no selection, the active layer is used."
        },
        guideLayer: {
            ja: "「_guide」レイヤーにまとめてガイドを作成し、作成後はレイヤーをロックします。",
            en: "Create the guides on the \"_guide\" layer and lock that layer afterwards."
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
        individual: {
            ja: "ON：選択オブジェクトごとにガイドを作成。OFF：選択全体の外接でまとめて1組作成。選択が1つ以下のときは使用しません。",
            en: "On: one set of guides per selected object. Off: one set for the combined selection bounds. Not used with 0–1 objects selected."
        },
        group: {
            ja: "作成したガイドを1つのグループにまとめます。オブジェクトごとに作成する場合は、オブジェクト単位でグループ化します。",
            en: "Put the created guides into a group. With per-object guides, each object gets its own group."
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
            right: { ja: "対象の右辺にガイドを作成", en: "Add a guide at the target's right edge" },
            soloHint: {
                ja: "option（Alt）＋クリックで、この項目だけをオンにします。",
                en: "Option/Alt-click to turn this one on and all the others off."
            }
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

/**
 * LABELS からドット区切りのキーで文言を取得
 * @param {string} labelKey - "panel.target" のようなドット区切りキー
 * @returns {string} 現在の言語の文言（見つからない場合はキーそのもの）
 */
function getLabel(labelKey) {
    var keyParts = labelKey.split(".");
    var node = LABELS;
    for (var i = 0; i < keyParts.length; i++) {
        if (node == null) return labelKey;
        node = node[keyParts[i]];
    }
    if (node == null) return labelKey;
    return node[currentLanguage] || node.en || labelKey;
}

/**
 * コロン付きラベルを取得（日本語は全角、英語は半角）
 * @param {string} labelKey - ドット区切りキー
 * @returns {string} コロンを付けた文言
 */
function labelWithColon(labelKey) {
    return getLabel(labelKey) + (currentLanguage === "ja" ? "：" : ":");
}

// =========================================
// 単位 / Units
// =========================================

/* rulerType の単位コード→ラベルと pt 換算係数 / Unit code to label and points-per-unit */
var UNIT_INFO = {
    0:  { label: "in",    points: 72.0 },
    1:  { label: "mm",    points: 72.0 / 25.4 },
    2:  { label: "pt",    points: 1.0 },
    3:  { label: "pica",  points: 12.0 },
    4:  { label: "cm",    points: 72.0 / 2.54 },
    5:  { label: "H",     points: 72.0 / 25.4 * 0.25 },
    6:  { label: "px",    points: 1.0 },
    7:  { label: "ft/in", points: 72.0 * 12.0 },
    8:  { label: "m",     points: 72.0 / 25.4 * 1000.0 },
    9:  { label: "yd",    points: 72.0 * 36.0 },
    10: { label: "ft",    points: 72.0 * 12.0 }
};

/* 単位が取れないときの既定 / Fallback when the ruler unit cannot be read */
var FALLBACK_UNIT_INFO = { label: "pt", points: 1.0 };

/**
 * 現在の定規単位の情報を取得
 * @returns {object} { label: string, points: number }（不明な場合は pt 扱い）
 */
function getRulerUnitInfo() {
    return UNIT_INFO[app.preferences.getIntegerPreference("rulerType")] || FALLBACK_UNIT_INFO;
}

/**
 * 現在の定規単位ラベルを取得
 * @returns {string} 単位ラベル（不明な場合は "pt"）
 */
function getRulerUnitLabel() {
    return getRulerUnitInfo().label;
}

/**
 * 入力文字列を定規単位として解釈し pt に変換
 * @param {string} inputText - 入力欄の文字列
 * @returns {number} pt 値（数値でない場合は 0）
 */
function rulerTextToPoints(inputText) {
    var inputValue = parseFloat(inputText);
    if (isNaN(inputValue)) inputValue = 0;
    return inputValue * getRulerUnitInfo().points;
}

// =========================================
// 境界の取得 / Bounds
// =========================================

/**
 * オブジェクトの境界を取得（クリップグループはマスクパス基準）
 * @param {PageItem} pageItem - 対象オブジェクト
 * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
 * @returns {number[]} [左, 上, 右, 下]
 */
function getItemBounds(pageItem, usePreviewBounds) {
    if (pageItem.typename === "GroupItem" && pageItem.clipped) {
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            if (pageItem.pageItems[i].clipping) {
                return usePreviewBounds ? pageItem.pageItems[i].visibleBounds : pageItem.pageItems[i].geometricBounds;
            }
        }
    }
    return usePreviewBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
}

/**
 * 複数オブジェクトの外接矩形を統合
 * @param {PageItem[]} pageItems - 対象オブジェクト
 * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
 * @returns {number[]|null} 統合した [左, 上, 右, 下]（対象が無い場合は null）
 */
function unionBounds(pageItems, usePreviewBounds) {
    var unionRect = null;
    for (var i = 0; i < pageItems.length; i++) {
        var itemBounds = getItemBounds(pageItems[i], usePreviewBounds);
        if (!unionRect) {
            unionRect = itemBounds.concat();
        } else {
            unionRect[0] = Math.min(unionRect[0], itemBounds[0]);
            unionRect[1] = Math.max(unionRect[1], itemBounds[1]);
            unionRect[2] = Math.max(unionRect[2], itemBounds[2]);
            unionRect[3] = Math.min(unionRect[3], itemBounds[3]);
        }
    }
    return unionRect;
}

/**
 * アクティブなアートボードの矩形を取得
 * @param {Document} doc - 対象ドキュメント
 * @returns {number[]|null} [左, 上, 右, 下]（取得できない場合は null）
 */
function getActiveArtboardRect(doc) {
    if (doc.artboards.length === 0) return null;
    var artboardIndex = doc.artboards.getActiveArtboardIndex();
    if (artboardIndex < 0 || artboardIndex >= doc.artboards.length) return null;
    return doc.artboards[artboardIndex].artboardRect;
}

/**
 * プレビュー境界使用時、テキストを一時的にアウトライン化して境界計算用に差し替える
 * @param {PageItem[]} selectedItems - 選択オブジェクト
 * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
 * @returns {object} { items: 境界計算用オブジェクト配列, restore: 復元関数 }
 */
function outlineTextForBounds(selectedItems, usePreviewBounds) {
    var passthrough = { items: selectedItems, restore: function () {} };
    if (!usePreviewBounds) return passthrough;

    var boundsItems = [];
    var duplicatedTexts = [];
    var originalTexts = [];
    for (var i = 0; i < selectedItems.length; i++) {
        var selectedItem = selectedItems[i];
        if (selectedItem && selectedItem.typename === "TextFrame") {
            duplicatedTexts.push(selectedItem.duplicate());
            originalTexts.push(selectedItem);
        } else {
            boundsItems.push(selectedItem);
        }
    }
    if (duplicatedTexts.length === 0) return passthrough;

    for (var j = 0; j < originalTexts.length; j++) originalTexts[j].hidden = true;

    /* 全選択解除は app.selection = null が確実（doc.selection = null は MRAP エラーになることがある）/ Use app.selection to deselect */
    app.selection = null;
    for (var k = 0; k < duplicatedTexts.length; k++) duplicatedTexts[k].selected = true;
    try {
        app.executeMenuCommand('expandStyle');
    } catch (e) {
        alert(getLabel("alert.expandError") + "\n" + e.message);
    }

    var outlinedTexts = [];
    for (var m = 0; m < duplicatedTexts.length; m++) {
        var outlined = duplicatedTexts[m].createOutline();
        outlinedTexts.push(outlined ? outlined : duplicatedTexts[m]);
    }
    return {
        items: boundsItems.concat(outlinedTexts),
        restore: function () {
            for (var n = 0; n < outlinedTexts.length; n++) {
                outlinedTexts[n].remove();
                originalTexts[n].hidden = false;
            }
        }
    };
}

/**
 * ガイドの基準となる外接矩形と、その元オブジェクトが載るレイヤーを取得
 * @param {PageItem[]} boundsItems - 境界計算に使うオブジェクト
 * @param {object} guideOptions - ガイド作成オプション
 * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
 * @returns {object[]} { bounds: number[], layer: Layer|null } の配列
 */
function collectTargetEntries(boundsItems, guideOptions, artboardRect) {
    /* 選択が無いときはアートボードを基準にする / Fall back to the artboard when nothing is selected */
    if (boundsItems.length === 0) return artboardRect ? [{ bounds: artboardRect.concat(), layer: null }] : [];

    if (guideOptions.individual) {
        var targetEntries = [];
        for (var i = 0; i < boundsItems.length; i++) {
            targetEntries.push({
                bounds: getItemBounds(boundsItems[i], guideOptions.usePreviewBounds),
                layer: boundsItems[i].layer
            });
        }
        return targetEntries;
    }
    var combinedBounds = unionBounds(boundsItems, guideOptions.usePreviewBounds);
    /* まとめて1組にするときは先頭オブジェクトのレイヤーを代表にする / Use the first item's layer for the combined set */
    return combinedBounds ? [{ bounds: combinedBounds, layer: boundsItems[0].layer }] : [];
}

// =========================================
// ガイドの線分 / Guide segments
// =========================================

/**
 * 境界とオプションから引くべきガイドの位置と向きを算出
 * @param {number[]} bounds - 基準の外接矩形 [左, 上, 右, 下]
 * @param {object} guideOptions - ガイド作成オプション
 * @param {number} offsetPt - 対象から離す距離（pt）
 * @returns {object[]} { position: number, orientation: string } の配列
 */
function directionsFromBounds(bounds, guideOptions, offsetPt) {
    var topPosition = bounds[1] + offsetPt;
    var leftPosition = bounds[0] - offsetPt;
    var bottomPosition = bounds[3] - offsetPt;
    var rightPosition = bounds[2] + offsetPt;
    var centerX = (leftPosition + rightPosition) / 2;
    var centerY = (topPosition + bottomPosition) / 2;

    var guideDirections = [];
    if (guideOptions.left) guideDirections.push({ position: leftPosition, orientation: "vertical" });
    if (guideOptions.right) guideDirections.push({ position: rightPosition, orientation: "vertical" });
    if (guideOptions.top) guideDirections.push({ position: topPosition, orientation: "horizontal" });
    if (guideOptions.bottom) guideDirections.push({ position: bottomPosition, orientation: "horizontal" });
    if (guideOptions.center) {
        guideDirections.push({ position: centerX, orientation: "vertical" });
        guideDirections.push({ position: centerY, orientation: "horizontal" });
    }
    if (guideOptions.centerMode === "vertical") guideDirections.push({ position: centerX, orientation: "vertical" });
    if (guideOptions.centerMode === "horizontal") guideDirections.push({ position: centerY, orientation: "horizontal" });
    return guideDirections;
}

/**
 * ガイド1本分の線分を算出
 * @param {number} position - ガイドの座標
 * @param {string} orientation - "vertical" または "horizontal"
 * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
 * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
 * @returns {number[][]} [始点, 終点]
 */
function guideSegment(position, orientation, drawSettings, artboardRect) {
    if (drawSettings.useCanvas) {
        return (orientation === "horizontal")
            ? [[-CANVAS_GUIDE_REACH, position], [CANVAS_GUIDE_REACH, position]]
            : [[position, CANVAS_GUIDE_REACH], [position, -CANVAS_GUIDE_REACH]];
    }
    var extensionPt = drawSettings.extensionPt;
    return (orientation === "horizontal")
        ? [[artboardRect[0] - extensionPt, position], [artboardRect[2] + extensionPt, position]]
        : [[position, artboardRect[1] + extensionPt], [position, artboardRect[3] - extensionPt]];
}

/**
 * 基準矩形1つ分のガイドの線分を算出
 * @param {number[]} bounds - 基準の外接矩形 [左, 上, 右, 下]
 * @param {object} guideOptions - ガイド作成オプション
 * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
 * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
 * @returns {number[][][]} 線分（[始点, 終点]）の配列
 */
function segmentsForBounds(bounds, guideOptions, drawSettings, artboardRect) {
    var guideDirections = directionsFromBounds(bounds, guideOptions, drawSettings.offsetPt);
    var segments = [];
    for (var i = 0; i < guideDirections.length; i++) {
        segments.push(guideSegment(guideDirections[i].position, guideDirections[i].orientation, drawSettings, artboardRect));
    }
    return segments;
}

/**
 * 基準矩形ごとに「描画先レイヤーと線分」の組を作る（本描画とプレビューで共通）
 * @param {object[]} targetEntries - { bounds: number[], layer: Layer|null } の配列
 * @param {object} guideOptions - ガイド作成オプション
 * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
 * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
 * @returns {object[]} { layer: Layer|null, segments: number[][][] } の配列
 */
function buildDrawPlan(targetEntries, guideOptions, drawSettings, artboardRect) {
    var drawPlan = [];
    for (var i = 0; i < targetEntries.length; i++) {
        drawPlan.push({
            layer: targetEntries[i].layer,
            segments: segmentsForBounds(targetEntries[i].bounds, guideOptions, drawSettings, artboardRect)
        });
    }
    return drawPlan;
}

// =========================================
// ガイドの作成 / Guide creation
// =========================================

/**
 * 名前でレイヤーを検索
 * @param {Document} doc - 対象ドキュメント
 * @param {string} layerName - レイヤー名
 * @returns {Layer|null} 見つかったレイヤー（無ければ null）
 */
function findLayerByName(doc, layerName) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) return doc.layers[i];
    }
    return null;
}

/**
 * ガイド用レイヤーを取得（無ければ作成）
 * @param {Document} doc - 対象ドキュメント
 * @returns {Layer} ガイド用レイヤー
 */
function getOrCreateGuideLayer(doc) {
    var guideLayer = findLayerByName(doc, GUIDE_LAYER_NAME);
    if (!guideLayer) {
        guideLayer = doc.layers.add();
        guideLayer.name = GUIDE_LAYER_NAME;
    }
    return guideLayer;
}

/**
 * レイヤー・グループ内の既存ガイドを削除（グループ化されたガイドも対象）
 * @param {Layer|GroupItem} guideContainer - 対象レイヤーまたはグループ
 * @returns {void}
 */
function deleteGuidesInContainer(guideContainer) {
    for (var i = guideContainer.pageItems.length - 1; i >= 0; i--) {
        var containedItem = guideContainer.pageItems[i];
        if (containedItem.typename === "GroupItem") {
            deleteGuidesInContainer(containedItem);
            /* 中身が無くなったグループは残さない / Drop groups left empty */
            if (containedItem.pageItems.length === 0) containedItem.remove();
        } else if (containedItem.guides) {
            containedItem.remove();
        }
    }
}

/**
 * 線分の配列からガイドを作成
 * @param {Layer} targetLayer - 作成先レイヤー
 * @param {number[][][]} segments - 線分（[始点, 終点]）の配列
 * @param {boolean} groupGuides - 作成したガイドをグループにまとめるかどうか
 * @returns {void}
 */
function addGuidesFromSegments(targetLayer, segments, groupGuides) {
    if (segments.length === 0) return;

    /* グループ化するときはガイドをグループ内に直接作成 / Create the guides inside a group when grouping */
    var guideContainer = groupGuides ? targetLayer.groupItems.add() : targetLayer;
    for (var i = 0; i < segments.length; i++) {
        var guidePath = guideContainer.pathItems.add();
        guidePath.setEntirePath(segments[i]);
        guidePath.filled = false;
        guidePath.stroked = false;
        guidePath.guides = true;
    }
}

/**
 * レイヤーのロック状態を記録して解除（記録済みなら何もしない）
 * @param {Layer} targetLayer - 対象レイヤー
 * @param {object[]} lockStates - { layer, wasLocked } の記録先
 * @returns {void}
 */
function unlockLayerOnce(targetLayer, lockStates) {
    for (var i = 0; i < lockStates.length; i++) {
        if (lockStates[i].layer === targetLayer) return;
    }
    lockStates.push({ layer: targetLayer, wasLocked: targetLayer.locked });
    targetLayer.locked = false;
}

/**
 * 選択オブジェクト・アートボード・カンバスを基準にガイドを作成（本処理）
 * @param {object} guideOptions - ガイド作成オプション（各方向フラグ・プレビュー境界・個別作成・グループ化・描画先・既存削除）
 * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
 * @returns {void}
 */
function createGuides(guideOptions, drawSettings) {
    var doc = app.activeDocument;
    var artboardRect = getActiveArtboardRect(doc);
    if (!drawSettings.useCanvas && !artboardRect) {
        alert(getLabel(doc.artboards.length === 0 ? "alert.noArtboard" : "alert.invalidArtboard"));
        return;
    }

    /* 境界計算のあいだだけテキストをアウトライン化 / Outline text only while measuring */
    var outlinedTextState = outlineTextForBounds(doc.selection, guideOptions.usePreviewBounds);
    var targetEntries = collectTargetEntries(outlinedTextState.items, guideOptions, artboardRect);
    /* 線分はアウトラインを戻す前に確定させる / Freeze the segments before restoring the outlined text */
    var drawPlan = buildDrawPlan(targetEntries, guideOptions, drawSettings, artboardRect);
    outlinedTextState.restore();

    /* 「_guide」レイヤーへまとめる場合は先に用意して既存ガイドを整理 / Prepare the "_guide" layer up front */
    var guideLayer = null;
    if (!guideOptions.drawOnSelectionLayer) {
        guideLayer = getOrCreateGuideLayer(doc);
        guideLayer.locked = false;
        if (guideOptions.deleteExisting) deleteGuidesInContainer(guideLayer);
    }

    var lockStates = [];
    for (var i = 0; i < drawPlan.length; i++) {
        var targetLayer = guideLayer || drawPlan[i].layer || doc.activeLayer;
        unlockLayerOnce(targetLayer, lockStates);
        addGuidesFromSegments(targetLayer, drawPlan[i].segments, guideOptions.groupGuides);
    }

    /* 「_guide」レイヤーはロックし、ほかは元の状態へ戻す / Lock "_guide", restore the others */
    for (var j = 0; j < lockStates.length; j++) {
        if (lockStates[j].layer !== guideLayer) lockStates[j].layer.locked = lockStates[j].wasLocked;
    }
    if (guideLayer) guideLayer.locked = true;
}

// =========================================
// ライブプレビュー / Live preview
// =========================================

/**
 * プレビュー用の色を作成（ドキュメントのカラースペースに合わせる）
 * @param {Document} doc - 対象ドキュメント
 * @returns {CMYKColor|RGBColor} プレビュー線の色
 */
function makePreviewColor(doc) {
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 90;
        cmykColor.yellow = 0;
        cmykColor.black = 0;
        return cmykColor;
    }
    var rgbColor = new RGBColor();
    rgbColor.red = 255;
    rgbColor.green = 0;
    rgbColor.blue = 255;
    return rgbColor;
}

/**
 * プレビュー用レイヤーを削除
 * @param {Document} doc - 対象ドキュメント
 * @returns {void}
 */
function removePreviewLayer(doc) {
    var previewLayer = findLayerByName(doc, PREVIEW_LAYER_NAME);
    if (!previewLayer) return;
    previewLayer.locked = false;
    previewLayer.visible = true;
    previewLayer.remove();
}

/**
 * プレビュー用レイヤーを用意（既存があれば作り直す）
 * @param {Document} doc - 対象ドキュメント
 * @returns {Layer} プレビュー用レイヤー
 */
function createPreviewLayer(doc) {
    removePreviewLayer(doc);
    var previewLayer = doc.layers.add();
    previewLayer.name = PREVIEW_LAYER_NAME;
    return previewLayer;
}

/**
 * 線分を色付きの仮線としてレイヤーへ描画
 * @param {Layer} previewLayer - 描画先レイヤー
 * @param {number[][][]} segments - 線分（[始点, 終点]）の配列
 * @param {CMYKColor|RGBColor} previewColor - 仮線の色
 * @returns {void}
 */
function drawPreviewSegments(previewLayer, segments, previewColor) {
    for (var i = 0; i < segments.length; i++) {
        var previewPath = previewLayer.pathItems.add();
        previewPath.setEntirePath([segments[i][0], segments[i][1]]);
        previewPath.filled = false;
        previewPath.stroked = true;
        previewPath.strokeColor = previewColor;
        previewPath.strokeWidth = 1;
    }
}

/**
 * 現在の設定からプレビュー線分を収集（テキストのアウトライン化は省略）
 * @param {object} guideOptions - ガイド作成オプション
 * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
 * @returns {number[][][]} 線分（[始点, 終点]）の配列
 */
function collectPreviewSegments(guideOptions, drawSettings) {
    var doc = app.activeDocument;
    var artboardRect = getActiveArtboardRect(doc);
    if (!drawSettings.useCanvas && !artboardRect) return [];

    var targetEntries = collectTargetEntries(doc.selection, guideOptions, artboardRect);
    var drawPlan = buildDrawPlan(targetEntries, guideOptions, drawSettings, artboardRect);
    var segments = [];
    for (var i = 0; i < drawPlan.length; i++) {
        segments = segments.concat(drawPlan[i].segments);
    }
    return segments;
}

// =========================================
// UI 部品 / UI parts
// =========================================

/**
 * ツールチップ付きチェックボックスを追加
 * @param {Group|Panel} parentContainer - 追加先コンテナ
 * @param {string} labelKey - ラベルのドット区切りキー
 * @param {string} tooltipKey - ツールチップのドット区切りキー
 * @param {boolean} initialValue - 初期状態
 * @returns {Checkbox} 追加したチェックボックス
 */
function addCheckbox(parentContainer, labelKey, tooltipKey, initialValue) {
    var checkbox = parentContainer.add("checkbox", undefined, getLabel(labelKey));
    checkbox.helpTip = getLabel(tooltipKey);
    checkbox.value = initialValue;
    return checkbox;
}

/**
 * 「ラベル＋数値入力＋単位」の行を追加
 * @param {Group|Panel} parentContainer - 追加先コンテナ
 * @param {string} labelKey - ラベルのドット区切りキー
 * @param {string} tooltipKey - ツールチップのドット区切りキー
 * @param {string} initialValue - 入力欄の初期値
 * @returns {object} { row: Group, input: EditText }
 */
function addUnitInputRow(parentContainer, labelKey, tooltipKey, initialValue) {
    var inputRow = parentContainer.add("group");
    setupRow(inputRow, "left", ROW_SPACING);
    var tooltipText = getLabel(tooltipKey);

    var rowLabel = inputRow.add("statictext", undefined, labelWithColon(labelKey));
    rowLabel.helpTip = tooltipText;
    var valueInput = inputRow.add("edittext", undefined, initialValue);
    valueInput.characters = 3;
    valueInput.helpTip = tooltipText;
    var unitLabel = inputRow.add("statictext", undefined, getRulerUnitLabel());
    unitLabel.helpTip = tooltipText;
    changeValueByArrowKey(valueInput);

    return { row: inputRow, input: valueInput };
}

/**
 * テキストフィールドで上下矢印キーによる数値増減を可能にする
 * @param {EditText} editText - 対象の入力欄
 * @returns {void}
 */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var currentValue = Number(editText.text);
        if (isNaN(currentValue)) return;
        var keyboardState = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            if (keyboardState.shiftKey) {
                /* Shift押下時は10の倍数スナップ / Snap to tens if Shift is pressed */
                currentValue = Math.round(currentValue / 10) * 10 + (event.keyName == "Up" ? 10 : -10);
            } else {
                currentValue += (event.keyName == "Up" ? 1 : -1);
            }

            event.preventDefault();
            editText.text = currentValue;
            /* プログラム変更は onChanging を発火しないため明示的に呼ぶ / fire onChanging manually */
            if (typeof editText.onChanging === "function") editText.onChanging();
        }
    });
}

/**
 * 「対象」パネル（基準ラジオ＋延長）を構築
 * @param {Group} parentContainer - 追加先コンテナ
 * @returns {object} { canvasRadio, artboardRadio, extensionInput, updateExtensionEnabled }
 */
function buildTargetPanel(parentContainer) {
    var targetPanel = parentContainer.add("panel", undefined, getLabel("panel.target"));
    setupPanel(targetPanel, 6);

    /* カンバス→アートボードの順に並べ、デフォルトはアートボード / Canvas → Artboard, default to Artboard */
    var canvasRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.canvas"));
    canvasRadio.helpTip = getLabel("tooltip.canvas");
    var artboardRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.artboard"));
    artboardRadio.helpTip = getLabel("tooltip.artboard");
    artboardRadio.value = true;

    var extensionRow = addUnitInputRow(targetPanel, "fieldLabel.extension", "tooltip.extension", "20");

    /**
     * 「延長」行のディムを更新する（カンバス基準では使わないので行ごとディムする）
     * @returns {void}
     */
    function updateExtensionEnabled() {
        extensionRow.row.enabled = !canvasRadio.value;
    }
    artboardRadio.onClick = updateExtensionEnabled;
    canvasRadio.onClick = updateExtensionEnabled;
    updateExtensionEnabled();

    return {
        canvasRadio: canvasRadio,
        artboardRadio: artboardRadio,
        extensionInput: extensionRow.input,
        updateExtensionEnabled: updateExtensionEnabled
    };
}

/**
 * 「ガイド位置」パネル（十字のチェックボックス）を構築
 * @param {Group} parentContainer - 追加先コンテナ
 * @returns {object} { left, top, right, bottom, center } のチェックボックス
 */
function buildAxisPanel(parentContainer) {
    var axisPanel = parentContainer.add("panel", undefined, getLabel("panel.axis"));
    setupPanel(axisPanel, 8);

    /**
     * 辺・中心のチェックボックスを追加する（option＋クリックの説明をツールチップに添える）
     * @param {Group|Panel} targetContainer - 追加先コンテナ
     * @param {string} labelKey - ラベルのドット区切りキー
     * @param {string} tooltipKey - ツールチップのドット区切りキー
     * @param {boolean} initialValue - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addEdgeCheckbox(targetContainer, labelKey, tooltipKey, initialValue) {
        var checkbox = addCheckbox(targetContainer, labelKey, tooltipKey, initialValue);
        checkbox.helpTip += "\n" + getLabel("tooltip.edge.soloHint");
        return checkbox;
    }

    var crossGroup = axisPanel.add("group");
    setupRow(crossGroup, undefined, CROSS_SPACING);
    /* パネル内で左右中央に配置 / Center the cross horizontally in the panel */
    crossGroup.alignment = ["center", "top"];

    var leftColumn = crossGroup.add("group");
    setupRow(leftColumn, undefined, 10);
    var leftCheckbox = addEdgeCheckbox(leftColumn, "checkbox.left", "tooltip.edge.left", true);

    var centerColumn = crossGroup.add("group");
    centerColumn.orientation = "column";
    centerColumn.alignChildren = ["left", "center"];
    centerColumn.spacing = 10;
    var topCheckbox = addEdgeCheckbox(centerColumn, "checkbox.top", "tooltip.edge.top", true);
    /* 中心（縦横）の中心線用 / Center lines (vertical & horizontal) */
    var centerCheckbox = addEdgeCheckbox(centerColumn, "checkbox.center", "tooltip.edge.center", false);
    var bottomCheckbox = addEdgeCheckbox(centerColumn, "checkbox.bottom", "tooltip.edge.bottom", true);

    var rightColumn = crossGroup.add("group");
    setupRow(rightColumn, undefined, 10);
    var rightCheckbox = addEdgeCheckbox(rightColumn, "checkbox.right", "tooltip.edge.right", true);

    return {
        left: leftCheckbox,
        top: topCheckbox,
        right: rightCheckbox,
        bottom: bottomCheckbox,
        center: centerCheckbox
    };
}

/**
 * プリセットのラジオボタンを生成して十字チェックボックスと連動させる
 * @param {Group} parentContainer - 追加先コンテナ
 * @param {object} crossCheckboxes - 十字のチェックボックス群
 * @param {object} centerLineState - 中心線モードの保持オブジェクト（{ mode: string }）
 * @returns {object} 生成したラジオボタン群
 */
function buildPresetRadios(parentContainer, crossCheckboxes, centerLineState) {
    /**
     * プリセットのラジオボタンを追加する（show が false のときは作らない）
     * @param {string} labelKey - ラベルのドット区切りキー
     * @param {string} tooltipKey - ツールチップのドット区切りキー
     * @param {boolean} [show] - 表示するかどうか（省略時は表示）
     * @returns {RadioButton|null} 追加したラジオボタン（作らなかった場合は null）
     */
    function addPresetRadio(labelKey, tooltipKey, show) {
        if (typeof show !== "undefined" && !show) return null;
        var presetRadio = parentContainer.add("radiobutton", undefined, getLabel(labelKey));
        presetRadio.helpTip = getLabel(tooltipKey);
        return presetRadio;
    }

    var presetRadios = {
        allOn: addPresetRadio("radio.allOn", "tooltip.preset.allOn"),
        edges: addPresetRadio("radio.edges", "tooltip.preset.edges"),
        topBottom: addPresetRadio("radio.vertical", "tooltip.preset.topBottom", showPresetTopBottom),
        leftRight: addPresetRadio("radio.horizontal", "tooltip.preset.leftRight", showPresetLeftRight),
        topLeft: addPresetRadio("radio.topLeft", "tooltip.preset.topLeft", showPresetTopLeft),
        bottomLeft: addPresetRadio("radio.bottomLeft", "tooltip.preset.bottomLeft", showPresetBottomLeft),
        topRight: addPresetRadio("radio.topRight", "tooltip.preset.topRight", showPresetTopRight),
        bottomRight: addPresetRadio("radio.bottomRight", "tooltip.preset.bottomRight", showPresetBottomRight),
        centerBoth: addPresetRadio("radio.centerBoth", "tooltip.preset.centerBoth"),
        centerVertical: addPresetRadio("radio.centerVertical", "tooltip.preset.centerVertical"),
        centerHorizontal: addPresetRadio("radio.centerHorizontal", "tooltip.preset.centerHorizontal"),
        clear: addPresetRadio("radio.clear", "tooltip.preset.clear")
    };

    /* プリセット定義（crossValues=[左,上,右,下,中心], centerLine=中心線モード）/ Preset table */
    var presetDefinitions = [
        { radio: presetRadios.allOn,            crossValues: [true,  true,  true,  true,  true ] },
        { radio: presetRadios.edges,            crossValues: [true,  true,  true,  true,  false] },
        { radio: presetRadios.topBottom,        crossValues: [false, true,  false, true,  false] },
        { radio: presetRadios.leftRight,        crossValues: [true,  false, true,  false, false] },
        { radio: presetRadios.topLeft,          crossValues: [true,  true,  false, false, false] },
        { radio: presetRadios.bottomLeft,       crossValues: [true,  false, false, true,  false] },
        { radio: presetRadios.topRight,         crossValues: [false, true,  true,  false, false] },
        { radio: presetRadios.bottomRight,      crossValues: [false, false, true,  true,  false] },
        { radio: presetRadios.clear,            crossValues: [false, false, false, false, false] },
        { radio: presetRadios.centerBoth,       crossValues: [false, false, false, false, true ] },
        { radio: presetRadios.centerVertical,   crossValues: [false, false, false, false, false], centerLine: "vertical" },
        { radio: presetRadios.centerHorizontal, crossValues: [false, false, false, false, false], centerLine: "horizontal" }
    ];
    for (var i = 0; i < presetDefinitions.length; i++) {
        (function (presetDefinition) {
            if (!presetDefinition.radio) return;
            presetDefinition.radio.onClick = function () {
                if (!presetDefinition.radio.value) return;
                var crossValues = presetDefinition.crossValues;
                centerLineState.mode = presetDefinition.centerLine || "";
                crossCheckboxes.left.value = crossValues[0];
                crossCheckboxes.top.value = crossValues[1];
                crossCheckboxes.right.value = crossValues[2];
                crossCheckboxes.bottom.value = crossValues[3];
                crossCheckboxes.center.value = crossValues[4];
            };
        })(presetDefinitions[i]);
    }

    /* デフォルト選択は四辺を優先 / Default selection (prefer "Edges") */
    if (presetRadios.edges) {
        presetRadios.edges.value = true;
    } else if (presetRadios.allOn) {
        presetRadios.allOn.value = true;
    }

    /* 手動でチェックを変えたら中心線モードを解除し、option＋クリックはクリックした項目だけをオンに
       / Clear center-line mode on manual toggle; Option-click keeps only the clicked item on */
    var crossKeys = ["left", "top", "right", "bottom", "center"];

    /**
     * 十字のチェックボックス1つにクリック時の処理を割り当てる
     * @param {string} crossKey - 対象のキー（"left" / "top" / "right" / "bottom" / "center"）
     * @returns {void}
     */
    function setupCrossCheckbox(crossKey) {
        crossCheckboxes[crossKey].onClick = function () {
            centerLineState.mode = "";
            if (!ScriptUI.environment.keyboardState.altKey) return;
            for (var k = 0; k < crossKeys.length; k++) {
                crossCheckboxes[crossKeys[k]].value = (crossKeys[k] === crossKey);
            }
        };
    }
    for (var j = 0; j < crossKeys.length; j++) {
        setupCrossCheckbox(crossKeys[j]);
    }

    return presetRadios;
}

/**
 * 「描画先」パネルを構築
 * @param {Window} dialog - 追加先ダイアログ
 * @returns {object} { selectionLayerRadio, guideLayerRadio }
 */
function buildDestinationPanel(dialog) {
    var destinationPanel = dialog.add("panel", undefined, getLabel("panel.destination"));
    setupPanel(destinationPanel, 6);

    var selectionLayerRadio = destinationPanel.add("radiobutton", undefined, getLabel("radio.selectionLayer"));
    selectionLayerRadio.helpTip = getLabel("tooltip.selectionLayer");
    var guideLayerRadio = destinationPanel.add("radiobutton", undefined, getLabel("radio.guideLayer"));
    guideLayerRadio.helpTip = getLabel("tooltip.guideLayer");
    guideLayerRadio.value = true;

    return { selectionLayerRadio: selectionLayerRadio, guideLayerRadio: guideLayerRadio };
}

/**
 * 「オプション」パネルを構築
 * @param {Window} dialog - 追加先ダイアログ
 * @param {number} selectionCount - 選択オブジェクト数
 * @returns {object} { usePreviewBoundsCheckbox, deleteGuidesCheckbox, individualCheckbox, groupCheckbox, offsetRow, offsetInput }
 */
function buildOptionsPanel(dialog, selectionCount) {
    var optionsPanel = dialog.add("panel", undefined, getLabel("panel.options"));
    setupPanel(optionsPanel, 6);

    var usePreviewBoundsCheckbox = addCheckbox(optionsPanel, "checkbox.usePreviewBounds", "tooltip.usePreviewBounds", true);
    var deleteGuidesCheckbox = addCheckbox(optionsPanel, "checkbox.deleteGuides", "tooltip.deleteGuides", true);
    var individualCheckbox = addCheckbox(optionsPanel, "checkbox.individual", "tooltip.individual", false);
    /* 選択が1つ以下ならまとめて作成と変わらないのでディム / Dim when 0–1 objects are selected */
    if (selectionCount <= 1) individualCheckbox.enabled = false;
    var groupCheckbox = addCheckbox(optionsPanel, "checkbox.group", "tooltip.group", true);

    var offsetRow = addUnitInputRow(optionsPanel, "fieldLabel.offset", "tooltip.offset", "0");
    offsetRow.input.active = true;

    return {
        usePreviewBoundsCheckbox: usePreviewBoundsCheckbox,
        deleteGuidesCheckbox: deleteGuidesCheckbox,
        individualCheckbox: individualCheckbox,
        groupCheckbox: groupCheckbox,
        offsetRow: offsetRow.row,
        offsetInput: offsetRow.input
    };
}

/**
 * フッター（左：プレビュー／右：ボタン）を構築
 * @param {Window} dialog - 追加先ダイアログ
 * @returns {object} { previewCheckbox, cancelButton, drawButton }
 */
function buildFooter(dialog) {
    var footerRow = dialog.add("group");
    setupRow(footerRow, "fill");
    footerRow.alignChildren = ["fill", "center"];
    footerRow.margins = [0, 10, 0, 0];

    /* 左：プレビュー / Left: preview */
    var footerLeftGroup = footerRow.add("group");
    setupRow(footerLeftGroup, undefined);
    footerLeftGroup.alignment = ["left", "center"];
    var previewCheckbox = addCheckbox(footerLeftGroup, "checkbox.preview", "tooltip.preview", true);

    /* 中央：スペーサー（余白を吸収）/ Center: spacer (absorbs free space) */
    var footerSpacerGroup = footerRow.add("group");
    footerSpacerGroup.alignment = ["fill", "center"];

    /* 右：ボタン（Mac 順：Cancel → OK）/ Right: buttons (Mac order) */
    var footerRightGroup = footerRow.add("group");
    setupRow(footerRightGroup, undefined);
    footerRightGroup.alignment = ["right", "center"];
    var cancelButton = footerRightGroup.add("button", undefined, getLabel("button.cancel"));
    var drawButton = footerRightGroup.add("button", undefined, getLabel("button.draw"), { name: "ok" });

    return { previewCheckbox: previewCheckbox, cancelButton: cancelButton, drawButton: drawButton };
}

/**
 * 選択オブジェクトがすべてアートボードの外にあるか判定
 * @param {Document} doc - 対象ドキュメント
 * @returns {boolean} すべて外にあれば true
 */
function isSelectionOutsideArtboard(doc) {
    var selectedItems = doc.selection;
    var artboardRect = getActiveArtboardRect(doc);
    if (selectedItems.length === 0 || !artboardRect) return false;

    for (var i = 0; i < selectedItems.length; i++) {
        var itemBounds = selectedItems[i].geometricBounds;
        var isOutside = itemBounds[0] > artboardRect[2] || itemBounds[2] < artboardRect[0] ||
            itemBounds[1] < artboardRect[3] || itemBounds[3] > artboardRect[1];
        if (!isOutside) return false;
    }
    return true;
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * メインダイアログを構築して表示
 * @returns {void}
 */
function buildDialog() {
    var doc = app.activeDocument;
    /* 中心線モードはプリセットと共有するため holder で保持 / Center-line mode holder shared with presets */
    var centerLineState = { mode: "" };

    var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(dialog);

    var columnsGroup = dialog.add("group");
    setupRow(columnsGroup, "fill", COLUMN_SPACING);
    columnsGroup.alignChildren = ["fill", "top"];

    var leftColumnGroup = columnsGroup.add("group");
    leftColumnGroup.orientation = "column";
    leftColumnGroup.alignChildren = ["fill", "top"];
    leftColumnGroup.spacing = WINDOW_SPACING;

    var targetControls = buildTargetPanel(leftColumnGroup);
    var crossCheckboxes = buildAxisPanel(leftColumnGroup);

    var presetPanel = columnsGroup.add("panel", undefined, getLabel("panel.preset"));
    setupPanel(presetPanel, 6);
    var presetRadios = buildPresetRadios(presetPanel, crossCheckboxes, centerLineState);

    var destinationControls = buildDestinationPanel(dialog);
    var optionControls = buildOptionsPanel(dialog, doc.selection.length);
    var footerControls = buildFooter(dialog);

    /**
     * 既存ガイドの削除チェックボックスのディムを更新する（「_guide」レイヤーに描くときだけ有効）
     * @returns {void}
     */
    function updateDeleteGuidesEnabled() {
        optionControls.deleteGuidesCheckbox.enabled = destinationControls.guideLayerRadio.value;
    }
    destinationControls.selectionLayerRadio.onClick = updateDeleteGuidesEnabled;
    destinationControls.guideLayerRadio.onClick = updateDeleteGuidesEnabled;
    updateDeleteGuidesEnabled();

    /* 選択がすべてアートボード外なら自動的にカンバス基準へ / Prefer canvas when the selection is off-artboard */
    if (isSelectionOutsideArtboard(doc)) {
        targetControls.canvasRadio.value = true;
        targetControls.artboardRadio.value = false;
        targetControls.updateExtensionEnabled();
    }

    /* ===== プレビュー配線 / Preview wiring ===== */

    /**
     * ダイアログの現在値をガイド作成オプションにまとめる
     * @returns {object} ガイド作成オプション（各方向フラグ・プレビュー境界・個別作成・グループ化・描画先・既存削除）
     */
    function readGuideOptions() {
        return {
            left: crossCheckboxes.left.value,
            right: crossCheckboxes.right.value,
            top: crossCheckboxes.top.value,
            bottom: crossCheckboxes.bottom.value,
            center: crossCheckboxes.center.value,
            centerMode: centerLineState.mode,
            usePreviewBounds: optionControls.usePreviewBoundsCheckbox.value,
            individual: optionControls.individualCheckbox.value,
            groupGuides: optionControls.groupCheckbox.value,
            drawOnSelectionLayer: destinationControls.selectionLayerRadio.value,
            deleteExisting: optionControls.deleteGuidesCheckbox.value
        };
    }

    /**
     * 対象・マージン・延長の現在値を描画設定にまとめる
     * @returns {object} { useCanvas: boolean, offsetPt: number, extensionPt: number }
     */
    function readDrawSettings() {
        return {
            useCanvas: targetControls.canvasRadio.value,
            offsetPt: rulerTextToPoints(optionControls.offsetInput.text),
            extensionPt: rulerTextToPoints(targetControls.extensionInput.text)
        };
    }

    /**
     * マージン行のディムを更新する（マージンは辺のガイドにだけ効くので、中心線だけのときはディムする）
     * @returns {void}
     */
    function updateOffsetEnabled() {
        optionControls.offsetRow.enabled = crossCheckboxes.left.value || crossCheckboxes.right.value ||
            crossCheckboxes.top.value || crossCheckboxes.bottom.value;
    }

    /* 既存ガイドの表示／非表示（showguide はトグルなので状態を自前で追跡）/ Toggle guide visibility (track state) */
    var guidesHidden = false;

    /**
     * 既存ガイドの表示・非表示を切り替える
     * @param {boolean} hide - 非表示にするなら true
     * @returns {void}
     */
    function setGuidesHidden(hide) {
        if (hide === guidesHidden) return;
        /* メニューコマンドが使えない状況でもダイアログ操作を止めない / Keep the dialog usable if the command is unavailable */
        try {
            app.executeMenuCommand("showguide");
        } catch (e) {}
        guidesHidden = hide;
    }

    /**
     * 現在の設定でライブプレビューを描き直す
     * @returns {void}
     */
    function renderPreview() {
        updateOffsetEnabled();
        removePreviewLayer(doc);
        /* プレビュー中は既存ガイドを隠し、仮ガイド（色付き線）だけ見せる / Hide real guides during preview */
        setGuidesHidden(footerControls.previewCheckbox.value);
        if (footerControls.previewCheckbox.value) {
            try {
                var segments = collectPreviewSegments(readGuideOptions(), readDrawSettings());
                if (segments.length > 0) {
                    var previewLayer = createPreviewLayer(doc);
                    drawPreviewSegments(previewLayer, segments, makePreviewColor(doc));
                    previewLayer.locked = true;
                }
            } catch (e) {
                removePreviewLayer(doc);
            }
        }
        app.redraw();
    }

    /**
     * 既存の onClick を保持したまま、後ろにプレビュー更新を連結する
     * @param {Object} control - 対象のコントロール（null のときは何もしない）
     * @returns {void}
     */
    function chainPreview(control) {
        if (!control) return;
        var previousOnClick = control.onClick;
        control.onClick = function () {
            if (previousOnClick) previousOnClick();
            renderPreview();
        };
    }
    var previewTriggers = [
        targetControls.canvasRadio, targetControls.artboardRadio,
        presetRadios.allOn, presetRadios.edges, presetRadios.topBottom, presetRadios.leftRight,
        presetRadios.topLeft, presetRadios.bottomLeft, presetRadios.topRight, presetRadios.bottomRight,
        presetRadios.centerBoth, presetRadios.centerVertical, presetRadios.centerHorizontal, presetRadios.clear,
        crossCheckboxes.left, crossCheckboxes.top, crossCheckboxes.right, crossCheckboxes.bottom, crossCheckboxes.center,
        optionControls.usePreviewBoundsCheckbox, optionControls.individualCheckbox, footerControls.previewCheckbox
    ];
    for (var i = 0; i < previewTriggers.length; i++) {
        chainPreview(previewTriggers[i]);
    }
    optionControls.offsetInput.onChanging = renderPreview;
    targetControls.extensionInput.onChanging = renderPreview;

    footerControls.cancelButton.onClick = function () {
        dialog.close();
    };

    footerControls.drawButton.onClick = function () {
        removePreviewLayer(doc); /* プレビューを片付けてから本処理 / clean up the preview before committing */
        try {
            createGuides(readGuideOptions(), readDrawSettings());
            dialog.close();
        } catch (e) {
            alert(getLabel("alert.guideError") + "\n" + (e && e.message ? e.message : e) + "\n" + (e && e.stack ? e.stack : ""));
        }
    };

    /* 表示時に初期プレビュー / Initial preview on show */
    dialog.onShow = function () {
        renderPreview();
    };
    /* 閉じる時は仮ガイドを片付け、隠した既存ガイドを再表示 / On close: clear preview and restore guide visibility */
    dialog.onClose = function () {
        removePreviewLayer(doc);
        setGuidesHidden(false);
        app.redraw();
    };

    dialog.show();
}

/* エントリーポイント / Entry point */
(function main() {
    if (!app.documents.length) {
        alert(getLabel("alert.noDocument"));
        return;
    }
    buildDialog();
})();
