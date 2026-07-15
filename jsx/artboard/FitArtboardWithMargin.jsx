#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "DialogEngine"

/*

### スクリプト名：

FitArtboardWithMargin.jsx

### 概要

- 更新日：20260715
- 選択オブジェクトまたはすべてのオブジェクトのバウンディングボックスにマージンを加え、アートボードを自動調整します。
- 定規単位に応じた初期マージン値と即時プレビュー付きダイアログを提供します。
- ピクセル整数値に丸めてアートボードを設定します。

### 主な機能

- 定規単位ごとのマージン初期値設定
- 外接バウンディングボックス計算
- 即時プレビュー付きダイアログ
- ピクセル整数丸め

### 処理の流れ

1. 対象（選択オブジェクトまたはアートボード）を選択
2. ダイアログでマージン値を設定（即時プレビュー対応）
3. 設定に基づきアートボードを自動調整

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

- ダイアログボックスを閉じずにプレビュー更新
- 単位系（mm、px など）によってデフォルト値を切り替え
- アートボードの座標・サイズをピクセルベースで整数値に
- オブジェクトを選択していない場合には、すべてのオブジェクトを対象に
- ↑↓キー、shift + ↑↓キーによる入力

### note

https://note.com/dtp_transit/n/n15d3c6c5a1e5

### 更新履歴

- v1.0 (20250420) : 初期バージョン
- v1.9.0 (20260715) : 丸めモードを3択化（ピクセルグリッド／現在の単位で整数／なし）、オプションパネル・UI共通化・プレビュー巻き戻し修正

---

### Script Name:

FitArtboardWithMargin.jsx

### Overview

- Last Updated: 20260715
- Automatically resize the artboard to fit the bounding box of selected or all objects with margin.
- Provides unit-based default margin values and an instant preview dialog.
- Sets the artboard size rounded to pixel integers.

### Main Features

- Default margin values based on ruler units
- Bounding box calculation
- Dialog with live preview
- Integer pixel rounding

### Workflow

1. Select target (selection or artboard)
2. Set margin value in dialog (with live preview)
3. Automatically adjust artboard size based on settings

### Changelog

- v1.0 (20250420): Initial version
- v1.9.0 (20260715): Three-way rounding mode (pixel grid / current unit / none), options panel, shared UI layout, preview rollback fix

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.9.0";

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================
var CONFIG = {
    supportedUnits: ['inch', 'mm', 'pt', 'pica', 'cm', 'H', 'px'], // rulerType の並びと対応 / matches rulerType order
    defaultMarginByUnit: {
        mm: '5',
        px: '20',
        pt: '10',
        _fallback: '0'
    },
    previewBoundsDefault: true,        // プレビュー境界(visibleBounds)を既定に / use visibleBounds by default
    roundModeDefault: 'pixelGrid',     // 既定の丸めモード（pixelGrid / currentUnit / none）/ default rounding mode
    linkDefault: true,                 // 上下左右の連動を既定ON / link margins by default
    dialogOpacity: 0.98,
    offsetX: 300
};

// =========================================
// ローカライズ / Localize
// =========================================

/* 実行環境の言語を判定（ja / en） / Detect UI language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "アートボードサイズを調整", en: "Adjust Artboard Size" }
    },
    panel: {
        target: { ja: "対象", en: "Target" },
        margin: { ja: "マージン", en: "Margin" },
        options: { ja: "アートボードの微調整", en: "Artboard fine-tuning" }
    },
    radio: {
        selection: { ja: "選択したオブジェクト", en: "Selected Objects" },
        artboard: { ja: "現在のアートボード", en: "Current Artboard" },
        allArtboards: { ja: "すべてのアートボード", en: "All Artboards" }
    },
    field: {
        vertical: { ja: "上下", en: "Vertical" },
        horizontal: { ja: "左右", en: "Horizontal" }
    },
    checkbox: {
        linked: { ja: "連動", en: "Linked" },
        previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
    },
    roundMode: {
        pixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to pixel grid" },
        currentUnit: { ja: "現在の単位で値を整数値に", en: "Round values in current unit" },
        none: { ja: "何もしない", en: "Do nothing" }
    },
    button: {
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        enterNumber: { ja: "数値を入力してください。", en: "Please enter a number." },
        errorOccurred: { ja: "エラーが発生しました: ", en: "An error occurred: " }
    },
    tooltip: {
        marginInput: {
            ja: "↑↓キーで±1、Shift+↑↓で10単位で増減できます",
            en: "Arrow keys: ±1, Shift+Arrow: ±10"
        },
        link: {
            ja: "上下の値を左右にも自動で適用します",
            en: "Apply the vertical value to horizontal as well"
        },
        previewBounds: {
            ja: "ON：線・効果を含む見た目の境界（プレビュー境界）で計測／OFF：パスの幾何境界で計測",
            en: "On: measure with preview (visible) bounds incl. strokes/effects; Off: geometric path bounds"
        },
        roundPixel: {
            ja: "座標とサイズを整数ピクセルに丸めます",
            en: "Round position and size to integer pixels"
        },
        roundUnit: {
            ja: "現在の定規単位で座標とサイズを整数に丸めます",
            en: "Round position and size to integers in the current ruler unit"
        },
        roundNone: {
            ja: "丸めずに計測値のまま設定します",
            en: "Apply the measured values without rounding"
        }
    }
};

// =========================================
// 単位 / Units
// =========================================

/* 単位ごとの初期マージン値を返す / Return default margin string for a unit */
function getDefaultMargin(unit) {
    return CONFIG.defaultMarginByUnit.hasOwnProperty(unit) ?
        CONFIG.defaultMarginByUnit[unit] :
        CONFIG.defaultMarginByUnit._fallback;
}

/* 数値＋単位を pt に変換（失敗時は NaN） / Convert value+unit to points */
function toPt(val, unit) {
    try {
        var n = Number(val);
        if (isNaN(n)) return NaN;
        return new UnitValue(n, unit).as('pt');
    } catch (e) {
        return NaN;
    }
}

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/* ウィンドウの共通設定 / Apply shared window layout */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ） / Trim a button's height by the given px (call after layout) */
function trimButtonHeight(button, px) {
    try {
        button.size = [button.size.width, button.size.height - px];
    } catch (e) {}
}

// =========================================
// エラー処理 / Error handling
// =========================================

/* Error を行番号・ファイル名付きで読みやすく整形 / Format an Error with line number */
function formatError(e) {
    try {
        var msg = (e && e.message) ? String(e.message) : String(e);
        var ln = (e && e.line) ? (" line " + e.line) : "";
        var fn = (e && e.fileName) ? (" (" + e.fileName + ")") : "";
        return msg + ln + fn;
    } catch (_) {
        return String(e);
    }
}

// =========================================
// プレビュー管理 / Preview manager
// =========================================

/**
 * プレビュー時の履歴管理と一括Undoを制御するクラス / Preview Undo/History manager
 *
 * - updatePreview() のたびに rollback() してから addStep() で最新状態を1回だけ適用
 * - OK/Cancel 時に rollback() してプレビュー操作を履歴から取り除く
 */
function PreviewManager() {
    this.undoDepth = 0; // プレビュー中に生成された Undo ステップの総数 / total undo steps created during preview

    /**
     * 変更操作を実行し、生成された Undo ステップ数を履歴として加算する。
     * func は「実際に生成した Undo ステップ数」を数値で返すこと。
     * 返さない（undefined）場合は 1 ステップとして扱う。
     * @param {function} func - 実行したい処理（生成 Undo ステップ数を返す無名関数）
     */
    this.addStep = function(func) {
        try {
            var undoSteps = func();
            this.undoDepth += (typeof undoSteps === "number" && undoSteps >= 0) ? undoSteps : 1;
            app.redraw();
        } catch (e) {
            $.writeln("[PreviewManager] addStep error: " + e); // 失敗時はカウントしない
        }
    };

    /* プレビューのために行った変更を全て取り消す（キャンセル時など） / Undo all preview changes */
    this.rollback = function() {
        try {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
        } catch (e) {
            $.writeln("[PreviewManager] rollback error: " + e);
            this.undoDepth = 0;
        }
        app.redraw();
    };

    /**
     * 現在の状態を確定する（OK時） / Confirm current state (on OK)
     * OK時は一度 rollback() してから main() 側で本処理を1回だけ実行するため、ここでは rollback のみ。
     */
    this.confirm = function() {
        this.rollback();
    };
}

// =========================================
// ダイアログ位置の記憶 / Dialog position persistence
// =========================================
// 共通エンジン名でセッションをまたいで位置を記憶し、key で保存先を分離する。
// Share session state across scripts; separate each dialog by key.

/* 保存済みのダイアログ位置を取得 / Get stored dialog location */
function getStoredLocation(storageKey) {
    return $.global[storageKey] && $.global[storageKey].length === 2 ? $.global[storageKey] : null;
}

/* ダイアログ位置をセッションに保存 / Store dialog location */
function storeLocation(storageKey, location) {
    $.global[storageKey] = [location[0], location[1]];
}

/* 位置を画面内に収める / Clamp a location within the screen */
function clampLocationToScreen(location) {
    try {
        var visibleBounds = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
        var clampedX = Math.max(visibleBounds[0] + 10, Math.min(location[0], visibleBounds[2] - 10));
        var clampedY = Math.max(visibleBounds[1] + 10, Math.min(location[1], visibleBounds[3] - 10));
        return [clampedX, clampedY];
    } catch (e) {
        return location;
    }
}

/* ダイアログ位置の記憶を設定し、保存関数を返す / Wire up dialog position persistence; return the saver
   保存位置があれば表示時に復元、無ければ初回はセンターからのオフセットで表示する。 */
function attachPositionPersistence(dialog, positionKey, firstRunOffsetX) {
    if ($.global[positionKey] === undefined) $.global[positionKey] = null;
    var savedLocation = getStoredLocation(positionKey);

    var persist = function() {
        storeLocation(positionKey, [dialog.location[0], dialog.location[1]]);
    };

    if (savedLocation) {
        dialog.onShow = function() {
            dialog.location = clampLocationToScreen(savedLocation);
        };
    } else {
        dialog.onShow = function() {
            dialog.layout.layout(true);
            var screenWidth = $.screens[0].right - $.screens[0].left;
            var screenHeight = $.screens[0].bottom - $.screens[0].top;
            var centerX = screenWidth / 2 - dialog.bounds.width / 2;
            var centerY = screenHeight / 2 - dialog.bounds.height / 2;
            dialog.location = [centerX + firstRunOffsetX, centerY];
        };
    }

    dialog.onMove = persist;
    return persist;
}

// =========================================
// 設定の記憶（セッション内） / Settings persistence (session only)
// =========================================
// $.global に設定を保持。#targetengine のためセッション中は保持されるが、再起動でリセット。
// Kept in $.global; persists during the session but resets when Illustrator restarts.

var SETTINGS_KEY = "__FitArtboardWithMargin_Settings";

/* 保存済みの設定を取得（無ければ null） / Get stored settings */
function getStoredSettings() {
    var stored = $.global[SETTINGS_KEY];
    return (stored && typeof stored === "object") ? stored : null;
}

/* 設定をセッションに保存 / Store settings */
function storeSettings(settings) {
    $.global[SETTINGS_KEY] = settings;
}

/* 保存済み設定と文脈から、ダイアログの初期値を解決する / Resolve dialog initial values from stored settings + context
   対象は保存値が現在の文脈で有効な場合のみ採用（"selection" は選択がある時のみ）。 */
function resolveInitialSettings(defaultMargin, artboardCount, hasSelection) {
    var saved = getStoredSettings();

    var target;
    if (!hasSelection && artboardCount === 1) {
        target = "artboard";
    } else if (!hasSelection && artboardCount > 1) {
        target = "allArtboards";
    } else {
        target = "selection";
    }
    if (saved && saved.target) {
        var savedTarget = saved.target;
        if ((savedTarget === "selection" && hasSelection) || savedTarget === "artboard" || savedTarget === "allArtboards") {
            target = savedTarget;
        }
    }

    return {
        marginV: (saved && saved.marginV != null) ? saved.marginV : defaultMargin,
        marginH: (saved && saved.marginH != null) ? saved.marginH : defaultMargin,
        link: (saved && typeof saved.link === "boolean") ? saved.link : CONFIG.linkDefault,
        previewBounds: (saved && typeof saved.previewBounds === "boolean") ? saved.previewBounds : CONFIG.previewBoundsDefault,
        roundMode: (saved && saved.roundMode) ? saved.roundMode : CONFIG.roundModeDefault,
        target: target
    };
}

// =========================================
// 矩形・境界のユーティリティ / Rect & bounds utilities
// =========================================

/* 矩形にマージンを加えた新しい矩形を返す / Return a new rect expanded by margin
   Illustrator の artboardRect は [left, top, right, bottom]（上が大・下が小）。 */
function expandRectByMargin(rect, verticalMarginPt, horizontalMarginPt) {
    return [
        rect[0] - horizontalMarginPt,
        rect[1] + verticalMarginPt,
        rect[2] + horizontalMarginPt,
        rect[3] - verticalMarginPt
    ];
}

/* アートボード矩形をピクセルグリッドに最適化（X/Y/W/H を整数化） / Snap artboard rect to pixel grid */
function snapRectToPixelGrid(rect) {
    var left = Math.round(rect[0]);
    var top = Math.round(rect[1]);
    var right = Math.round(rect[2]);
    var bottom = Math.round(rect[3]);
    var width = Math.round(right - left);
    var height = Math.round(top - bottom);
    return [left, top, left + width, top - height];
}

/* 矩形の X/Y/W/H を指定単位で整数化 / Snap rect X/Y/W/H to integers in the given unit
   例：単位がmmなら各値をmm単位の整数に丸める。単位変換に失敗した場合はピクセルグリッドにフォールバック。 */
function snapRectToUnitGrid(rect, unit) {
    var ptPerUnit = toPt(1, unit);
    if (isNaN(ptPerUnit) || ptPerUnit === 0) return snapRectToPixelGrid(rect);
    var left = Math.round(rect[0] / ptPerUnit) * ptPerUnit;
    var top = Math.round(rect[1] / ptPerUnit) * ptPerUnit;
    var right = Math.round(rect[2] / ptPerUnit) * ptPerUnit;
    var bottom = Math.round(rect[3] / ptPerUnit) * ptPerUnit;
    var width = Math.round((right - left) / ptPerUnit) * ptPerUnit;
    var height = Math.round((top - bottom) / ptPerUnit) * ptPerUnit;
    return [left, top, left + width, top - height];
}

/* 丸めモードに従って矩形を整数化 / Round rect according to the selected mode
   "pixelGrid" = ピクセル整数、"currentUnit" = 現在の単位で整数、"none" = 丸めなし。 */
function applyRounding(rect, roundMode, unit) {
    if (roundMode === "pixelGrid") return snapRectToPixelGrid(rect);
    if (roundMode === "currentUnit") return snapRectToUnitGrid(rect, unit);
    return rect; // "none"
}

/* 2つの矩形が実質的に同一か判定 / Check if two rects are effectively equal
   マージン0などで値が変わらない場合はUndoが生成されないため、プレビューのUndoカウント除外に使う。 */
function rectsEqual(rectA, rectB) {
    if (!rectA || !rectB) return false;
    for (var i = 0; i < 4; i++) {
        if (Math.abs(rectA[i] - rectB[i]) > 0.0001) return false;
    }
    return true;
}

/* オブジェクトのバウンディングボックスを取得 / Get bounding box of a single object
   usePreviewBounds=true なら visibleBounds（塗り/線を含む）、false なら geometricBounds（パス外形のみ）。 */
function getBounds(item, usePreviewBounds) {
    return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
}

/* 複数アイテムから最大の外接バウンディングボックスを取得 / Get union bounding box from multiple items */
function getMaxBounds(items, usePreviewBounds) {
    var bounds = getBounds(items[0], usePreviewBounds);
    for (var i = 1; i < items.length; i++) {
        var itemBounds = getBounds(items[i], usePreviewBounds);
        bounds[0] = Math.min(bounds[0], itemBounds[0]);
        bounds[1] = Math.max(bounds[1], itemBounds[1]);
        bounds[2] = Math.max(bounds[2], itemBounds[2]);
        bounds[3] = Math.min(bounds[3], itemBounds[3]);
    }
    return bounds;
}

/* 選択アイテムの正規化 / Normalize selected items
   クリップグループ(GroupItem.clipped=true)はクリッピングパスのみを採用、それ以外はそのまま。
   プレビューと本処理でアイテム収集を共通化する。 */
function collectEffectiveItems(items) {
    var out = [];
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it && it.typename === "GroupItem" && it.clipped) {
            // クリップグループはクリッピングパスのみ採用 / include only the clipping path
            try {
                for (var j = 0; j < it.pageItems.length; j++) {
                    var child = it.pageItems[j];
                    if (child.clipping) {
                        out.push(child);
                        break;
                    }
                }
            } catch (e) {
                /* ignore */ }
        } else {
            out.push(it);
        }
    }
    return out;
}

/* 選択オブジェクトの外接境界を取得（テキストは一時アウトライン化して正確に計測） / Measure selection bounds
   テキストを複製して隠し、原本をアウトライン化 → 計測 → アウトラインを削除し複製を復元する。 */
function measureSelectionBounds(items, usePreviewBounds) {
    var originalTextFrames = [];
    var outlinedTextFrames = [];
    var nonTextItems = [];

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === "TextFrame") {
            var hiddenTextClone = item.duplicate();
            hiddenTextClone.hidden = true;
            originalTextFrames.push(hiddenTextClone);

            var outlinedResult = item.createOutline();
            // createOutline() が複数パスを返す場合に対応 / handle multi-path outline results
            if (outlinedResult.length && outlinedResult.length > 0) {
                for (var j = 0; j < outlinedResult.length; j++) outlinedTextFrames.push(outlinedResult[j]);
            } else {
                outlinedTextFrames.push(outlinedResult);
            }
        } else {
            nonTextItems.push(item);
        }
    }

    var bounds = getMaxBounds(outlinedTextFrames.concat(nonTextItems), usePreviewBounds);

    // アウトラインを破棄し、隠していた複製を復元 / discard outlines, restore hidden clones
    for (var i = 0; i < outlinedTextFrames.length; i++) {
        try { outlinedTextFrames[i].remove(); } catch (e) {}
    }
    for (var i = 0; i < originalTextFrames.length; i++) {
        try { originalTextFrames[i].hidden = false; } catch (e) {}
    }
    return bounds;
}

/* 1枚のアートボードにマージンを適用（丸めモードに従って整数化） / Apply margin to a single artboard */
function applyMarginToArtboard(artboard, verticalMarginPt, horizontalMarginPt, roundMode, unit) {
    var expandedRect = expandRectByMargin(artboard.artboardRect, verticalMarginPt, horizontalMarginPt);
    expandedRect = applyRounding(expandedRect, roundMode, unit);
    artboard.artboardRect = expandedRect;
}

// =========================================
// プレビュー適用 / Preview application
// =========================================
// 各関数は「実際に artboardRect を書き換えて生成した Undo ステップ数」を返す。
// 矩形が変わらない場合は書き換えずに 0 を返し、rollback 時の過剰 Undo を防ぐ。

/* 全アートボードにプレビュー適用。変更した枚数（＝生成Undoステップ数）を返す / Preview all artboards */
function previewAllArtboards(originalRects, verticalMarginPt, horizontalMarginPt) {
    var artboards = app.activeDocument.artboards;
    var changedCount = 0;
    for (var i = 0; i < artboards.length; i++) {
        var expandedRect = expandRectByMargin(originalRects[i], verticalMarginPt, horizontalMarginPt);
        if (rectsEqual(artboards[i].artboardRect, expandedRect)) continue;
        artboards[i].artboardRect = expandedRect;
        changedCount++;
    }
    app.redraw();
    return changedCount;
}

/* 現在のアートボードにプレビュー適用。0 or 1（生成Undoステップ数）を返す / Preview active artboard */
function previewArtboard(originalRects, index, verticalMarginPt, horizontalMarginPt) {
    var artboards = app.activeDocument.artboards;
    var expandedRect = expandRectByMargin(originalRects[index], verticalMarginPt, horizontalMarginPt);
    if (rectsEqual(artboards[index].artboardRect, expandedRect)) {
        app.redraw();
        return 0;
    }
    artboards[index].artboardRect = expandedRect;
    app.redraw();
    return 1;
}

/* 選択オブジェクトの外接からプレビュー適用。0 or 1 を返す / Preview from selection bounds */
function previewSelection(index, usePreviewBounds, verticalMarginPt, horizontalMarginPt) {
    var doc = app.activeDocument;
    var previewItems = doc.selection.length === 0 ? doc.pageItems : doc.selection;
    var effectiveItems = collectEffectiveItems(previewItems);
    if (effectiveItems.length === 0) return 0;
    var expandedRect = expandRectByMargin(getMaxBounds(effectiveItems, usePreviewBounds), verticalMarginPt, horizontalMarginPt);
    if (rectsEqual(doc.artboards[index].artboardRect, expandedRect)) {
        app.redraw();
        return 0;
    }
    doc.artboards[index].artboardRect = expandedRect;
    app.redraw();
    return 1;
}

// =========================================
// 入力ユーティリティ / Input utilities
// =========================================

/* マージン入力の検証と単位変換を一元化 / Validate margins and convert to pt
   `h` が未入力/NaN の場合は `v` を採用（連動相当のフォールバック）。 */
function parseMarginPair(textV, textH, unit) {
    var v = parseFloat(textV);
    var h = parseFloat(textH);
    if (isNaN(v)) return { valid: false };
    if (isNaN(h)) h = v; // fallback like linked
    var vPt = toPt(v, unit);
    var hPt = toPt(h, unit);
    if (isNaN(vPt) || isNaN(hPt)) return { valid: false };
    return { valid: true, v: v, h: h, vPt: vPt, hPt: hPt };
}

/* edittext に矢印キーでの増減を付与 / Add arrow-key increment/decrement to an edittext
   ↑↓で±1、shift+↑↓で10の倍数にスナップ。 */
function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        // 修飾キーは event から読む（keyboardState は macOS で誤報あり）。取得不可時のみフォールバック。
        var keyboardState = ScriptUI.environment.keyboardState;
        var isShiftPressed = (event.shiftKey !== undefined) ? event.shiftKey : keyboardState.shiftKey;

        if (event.keyName == "Up" || event.keyName == "Down") {
            if (isShiftPressed) {
                // 10の倍数にスナップ / snap to multiples of 10
                var base = Math.round(value / 10) * 10;
                value = (event.keyName == "Up") ? base + 10 : base - 10;
            } else {
                value += (event.keyName == "Up") ? 1 : -1;
            }

            event.preventDefault();
            editText.text = value;
            if (typeof onUpdate === "function") onUpdate(editText.text);
        }
    });
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* マージン入力ダイアログを表示し設定を返す（ライブプレビュー付き） / Show margin dialog with live preview */
function showMarginDialog(defaultMargin, rulerUnit, artboardCount, hasSelection) {
    var dialog = new Window("dialog", LABELS.dialog.title[lang] + " " + SCRIPT_VERSION);

    // ダイアログ位置の記憶（保存関数を受け取る） / wire position persistence, get the saver
    var persistDialogLocation = attachPositionPersistence(dialog, "__FitArtboardWithMargin_Dialog", CONFIG.offsetX);
    dialog.opacity = CONFIG.dialogOpacity;
    setupWindow(dialog);

    // 保存済み設定（セッション内）から初期値を解決 / resolve initial values from stored settings
    var initial = resolveInitialSettings(defaultMargin, artboardCount, hasSelection);

    /* 対象選択パネル / Target selection panel */
    var targetPanel = dialog.add("panel", undefined, LABELS.panel.target[lang]);
    setupPanel(targetPanel);

    var targetRadioGroup = targetPanel.add("group");
    targetRadioGroup.orientation = "column";
    targetRadioGroup.alignChildren = "left";

    var selectionRadio = targetRadioGroup.add("radiobutton", undefined, LABELS.radio.selection[lang]);
    selectionRadio.enabled = hasSelection;
    var artboardRadio = targetRadioGroup.add("radiobutton", undefined, LABELS.radio.artboard[lang]);
    var allArtboardsRadio = targetRadioGroup.add("radiobutton", undefined, LABELS.radio.allArtboards[lang]);

    /* 対象選択の初期値を適用 / apply initial radio selection */
    selectionRadio.value = (initial.target === "selection");
    artboardRadio.value = (initial.target === "artboard");
    allArtboardsRadio.value = (initial.target === "allArtboards");

    /* マージン入力パネル（2カラム） / Margin input panel (two columns) */
    var marginPanel = dialog.add("panel", undefined, LABELS.panel.margin[lang] + " (" + rulerUnit + ")");
    setupPanel(marginPanel);
    marginPanel.orientation = "row";
    marginPanel.alignChildren = ["left", "top"];
    marginPanel.spacing = COLUMN_SPACING;

    var marginFieldsColumn = marginPanel.add("group");
    marginFieldsColumn.orientation = "column";
    marginFieldsColumn.alignChildren = ["left", "center"];

    var linkColumn = marginPanel.add("group");
    linkColumn.orientation = "column";
    linkColumn.alignChildren = ["left", "center"];
    linkColumn.alignment = ["left", "center"];

    /* 上下マージン入力欄 / Vertical margin input */
    var verticalMarginRow = marginFieldsColumn.add("group");
    verticalMarginRow.orientation = "row";
    var verticalMarginLabel = verticalMarginRow.add("statictext", undefined, LABELS.field.vertical[lang] + ":");
    var verticalMarginInput = verticalMarginRow.add("edittext", undefined, initial.marginV);
    verticalMarginInput.characters = 4;
    verticalMarginInput.helpTip = LABELS.tooltip.marginInput[lang];

    /* 左右マージン入力欄 / Horizontal margin input */
    var horizontalMarginRow = marginFieldsColumn.add("group");
    horizontalMarginRow.orientation = "row";
    var horizontalMarginLabel = horizontalMarginRow.add("statictext", undefined, LABELS.field.horizontal[lang] + ":");
    var horizontalMarginInput = horizontalMarginRow.add("edittext", undefined, initial.marginH);
    horizontalMarginInput.characters = 4;
    horizontalMarginInput.enabled = !initial.link; // 連動ONなら左右はディム
    horizontalMarginInput.helpTip = LABELS.tooltip.marginInput[lang];

    /* 連動チェックボックス / Linked checkbox */
    var linkCheckbox = linkColumn.add("checkbox", undefined, LABELS.checkbox.linked[lang]);
    linkCheckbox.value = initial.link;
    linkCheckbox.helpTip = LABELS.tooltip.link[lang];

    /* プレビュー境界（visibleBounds を採用するか） / use visibleBounds for measurement */
    var previewBoundsCheckbox = linkColumn.add("checkbox", undefined, LABELS.checkbox.previewBounds[lang]);
    previewBoundsCheckbox.value = initial.previewBounds;
    previewBoundsCheckbox.helpTip = LABELS.tooltip.previewBounds[lang];
    previewBoundsCheckbox.onClick = updatePreview;

    /* 現在／全アートボードの rect を保存（プレビュー復元用） / Save artboard rects for preview restore */
    var activeArtboardIndex = app.activeDocument.artboards.getActiveArtboardIndex();
    var originalArtboardRects = [];
    for (var i = 0; i < app.activeDocument.artboards.length; i++) {
        originalArtboardRects.push(app.activeDocument.artboards[i].artboardRect.slice());
    }

    var previewManager = new PreviewManager();

    /* プレビュー更新：直前分を rollback してから最新状態を1回だけ適用 / Refresh preview via PreviewManager
       対象モードに応じてトップレベルのプレビュー関数へ委譲する。 */
    function updatePreview() {
        previewManager.rollback();

        var parsed = parseMarginPair(verticalMarginInput.text, horizontalMarginInput.text, rulerUnit);
        if (!parsed.valid) return;
        var verticalMarginPt = parsed.vPt;
        var horizontalMarginPt = parsed.hPt;

        var targetMode = selectionRadio.value ? "selection" : (artboardRadio.value ? "artboard" : "allArtboards");

        previewManager.addStep(function() {
            if (targetMode === "allArtboards") return previewAllArtboards(originalArtboardRects, verticalMarginPt, horizontalMarginPt);
            if (targetMode === "artboard") return previewArtboard(originalArtboardRects, activeArtboardIndex, verticalMarginPt, horizontalMarginPt);
            return previewSelection(activeArtboardIndex, previewBoundsCheckbox.value, verticalMarginPt, horizontalMarginPt);
        });
    }

    /* 矢印キー・入力・ラジオ・連動のハンドラ登録 / Wire up input handlers */
    changeValueByArrowKey(verticalMarginInput, function(val) {
        if (linkCheckbox.value) horizontalMarginInput.text = val;
        updatePreview();
    });
    changeValueByArrowKey(horizontalMarginInput, function() {
        if (linkCheckbox.value) return;
        updatePreview();
    });
    verticalMarginInput.active = true;
    verticalMarginInput.onChanging = function() {
        if (linkCheckbox.value) horizontalMarginInput.text = verticalMarginInput.text;
        updatePreview();
    };
    horizontalMarginInput.onChanging = function() {
        if (linkCheckbox.value) return; // 連動中は水平の直接編集は無効
        updatePreview();
    };
    linkCheckbox.onClick = function() {
        if (linkCheckbox.value) {
            horizontalMarginInput.text = verticalMarginInput.text;
            horizontalMarginInput.enabled = false;
        } else {
            horizontalMarginInput.enabled = true;
        }
        updatePreview();
    };
    selectionRadio.onClick = updatePreview;
    artboardRadio.onClick = updatePreview;
    allArtboardsRadio.onClick = updatePreview;

    /* アートボードの微調整パネル / Artboard fine-tuning panel */
    var optionsPanel = dialog.add("panel", undefined, LABELS.panel.options[lang]);
    setupPanel(optionsPanel);

    // 丸めモード（XYWHの整数化方法） / rounding mode for artboard X/Y/W/H
    // ライブプレビューは丸め前の値で表示し、確定時（main）にのみ丸めを適用する。
    var roundModeGroup = optionsPanel.add("group");
    roundModeGroup.orientation = "column";
    roundModeGroup.alignChildren = "left";
    var roundPixelRadio = roundModeGroup.add("radiobutton", undefined, LABELS.roundMode.pixelGrid[lang]);
    var roundUnitRadio = roundModeGroup.add("radiobutton", undefined, LABELS.roundMode.currentUnit[lang]);
    var roundNoneRadio = roundModeGroup.add("radiobutton", undefined, LABELS.roundMode.none[lang]);
    roundPixelRadio.helpTip = LABELS.tooltip.roundPixel[lang];
    roundUnitRadio.helpTip = LABELS.tooltip.roundUnit[lang];
    roundNoneRadio.helpTip = LABELS.tooltip.roundNone[lang];
    roundUnitRadio.value = (initial.roundMode === "currentUnit");
    roundNoneRadio.value = (initial.roundMode === "none");
    roundPixelRadio.value = !roundUnitRadio.value && !roundNoneRadio.value; // 既定 / default

    /* ボタングループ（中央寄せ、fillしない） / Button group (centered, not filled) */
    var buttonGroup = dialog.add("group");
    setupRow(buttonGroup, "center");
    var cancelButton = buttonGroup.add("button", undefined, LABELS.button.cancel[lang], { name: "cancel" });
    var okButton = buttonGroup.add("button", undefined, LABELS.button.ok[lang], { name: "ok" });

    var dialogResult = null;
    okButton.onClick = function() {
        var parsed = parseMarginPair(verticalMarginInput.text, horizontalMarginInput.text, rulerUnit);
        if (parsed && parsed.valid) {
            dialogResult = {
                marginV: verticalMarginInput.text,
                marginH: horizontalMarginInput.text,
                target: selectionRadio.value ? "selection" : (artboardRadio.value ? "artboard" : "allArtboards"),
                previewBounds: previewBoundsCheckbox.value,
                roundMode: roundPixelRadio.value ? "pixelGrid" : (roundUnitRadio.value ? "currentUnit" : "none")
            };
            // 設定をセッションに保存（次回の初期値に） / store settings for next run
            storeSettings({
                marginV: dialogResult.marginV,
                marginH: dialogResult.marginH,
                link: linkCheckbox.value,
                previewBounds: dialogResult.previewBounds,
                roundMode: dialogResult.roundMode,
                target: dialogResult.target
            });
            persistDialogLocation();
            // プレビューによる履歴を一括Undoしてから閉じる / clean preview history, then close
            previewManager.confirm();
            dialog.close(1);
        } else {
            alert(LABELS.alert.enterNumber[lang]);
        }
    };
    cancelButton.onClick = function() {
        persistDialogLocation();
        // キャンセル時は必ずロールバックして閉じる / rollback preview and close
        previewManager.rollback();
        dialog.close(0);
    };

    updatePreview();
    dialog.show();
    return dialogResult;
}

// =========================================
// メイン処理 / Main
// =========================================

/* 対象を決定し、ダイアログの設定に従ってアートボードを調整する / Resolve target and apply margins */
function main() {
    try {
        var doc = app.activeDocument;

        var selectedItems = doc.selection;
        if (selectedItems.length === 0) {
            selectedItems = doc.pageItems;
            if (selectedItems.length === 0) return;
        }

        var artboards = doc.artboards;
        var rulerType = app.preferences.getIntegerPreference("rulerType");
        var rulerUnit = CONFIG.supportedUnits[rulerType];

        /* 単位ごとの初期マージン値 / default margin for the unit */
        var defaultMarginValue = getDefaultMargin(rulerUnit);

        /* 選択なし・複数アートボード時は allArtboards をデフォルトに（初期マージンは0） / default to allArtboards with 0 margin */
        if (doc.selection.length === 0 && artboards.length > 1) {
            defaultMarginValue = '0';
        }

        var userInput = showMarginDialog(defaultMarginValue, rulerUnit, artboards.length, doc.selection.length > 0);
        if (!userInput) return;

        var verticalMarginPt = toPt(parseFloat(userInput.marginV), rulerUnit);
        var horizontalMarginPt = toPt(parseFloat(userInput.marginH), rulerUnit);
        var roundMode = userInput.roundMode;

        if (userInput.target === "artboard") {
            applyMarginToArtboard(artboards[artboards.getActiveArtboardIndex()], verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit);
            return;
        }

        if (userInput.target === "allArtboards") {
            for (var i = 0; i < artboards.length; i++) {
                applyMarginToArtboard(artboards[i], verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit);
            }
            return;
        }

        if (userInput.target === "selection") {
            // 選択の正規化（クリップグループ→クリッピングパス） / normalize selection
            selectedItems = collectEffectiveItems(selectedItems);
            if (selectedItems.length === 0) return;

            var selectedBounds = measureSelectionBounds(selectedItems, userInput.previewBounds);
            selectedBounds = expandRectByMargin(selectedBounds, verticalMarginPt, horizontalMarginPt);
            selectedBounds = applyRounding(selectedBounds, roundMode, rulerUnit);

            artboards[artboards.getActiveArtboardIndex()].artboardRect = selectedBounds;
        }

    } catch (e) {
        $.writeln("[FitArtboardWithMargin] ERROR: " + formatError(e));
        alert(LABELS.alert.errorOccurred[lang] + formatError(e));
    }
}

main();
app.selectTool("Adobe Select Tool");

})();
