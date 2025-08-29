#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

var SCRIPT_VERSION = "v1.2";

/*
===============================================================================
スクリプト名 / Script Name
  Fit Artboard Size to Specified Width/Height

GitHub
  https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ResizeArtboardsAll.jsx

概要 / Overview
  JP: ダイアログで指定した「幅」「高さ」に、アートボードを即時プレビューしながら変形します。
      選択がない場合は、各アートボード内のオブジェクトを基準に、すべてのアートボードを個別に調整します。
  EN: Resize artboards to the specified width and height with live preview.
      When nothing is selected, each artboard is adjusted individually based on items inside it.

主な機能 / Key Features
  - ライブプレビュー（高速化のため app.redraw をデバウンス）
  - 「対象」パネル（作業ABのみ／すべてのAB）
  - 「基準」パネル（左上／中央）
  - テキストは一時的にアウトライン化して境界を精密取得（処理後に復元）
  - 入力フィールドは単位ラベル付き（値は数値のみ）／ダイアログ位置を記憶

使い方 / Flow
  1) ダイアログで幅・高さを入力（↑↓: ±1、Shift: ±10、Option: ±0.1 ※最終的に整数）
  2) プレビューで即時にアートボードを更新
  3) OKで確定、Cancelで元に戻す

注意 / Notes
  - プレビューはパフォーマンス重視。画面描画は間引き（throttle）ています。
  - 単位はドキュメントの定規単位に追従します（値はフィールド内で単位なし表示）。

更新履歴 / Changelog
  - v1.0 (2025-08-29): 初期バージョン / Initial release
===============================================================================
*/

// --- Core references & defaults ---
var doc = (app.documents.length > 0) ? app.activeDocument : null;
if (!doc) {
    alert(LABELS.alertNoDoc[lang]);
    throw new Error("No document");
}
var artboards = doc.artboards;

function _detectUnitString() {
    try {
        var ru = doc.rulerUnits;
        switch (ru) {
            case RulerUnits.Millimeters:
                return 'mm';
            case RulerUnits.Centimeters:
                return 'cm';
            case RulerUnits.Inches:
                return 'inch';
            case RulerUnits.Points:
                return 'pt';
            case RulerUnits.Picas:
                return 'pica';
            case RulerUnits.Pixels:
                return 'px';
            default:
                return 'pt';
        }
    } catch (e) {
        return 'pt';
    }
}

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese–English label definitions */


var LABELS = {
    // 1) Dialog title
    dialogTitle: {
        ja: "アートボードサイズを調整 " + SCRIPT_VERSION,
        en: "Adjust Artboard Size " + SCRIPT_VERSION
    },
    // 4) Target panel
    targetPanelTitle: {
        ja: "対象",
        en: "Target"
    },
    targetActiveArtboard: {
        ja: "作業アートボードのみ",
        en: "Active artboard only"
    },
    targetAllArtboards: {
        ja: "すべてのアートボード",
        en: "All artboards"
    },
    // 6) Buttons
    okBtn: {
        ja: "OK",
        en: "OK"
    },
    cancelBtn: {
        ja: "キャンセル",
        en: "Cancel"
    },
    // 7) Alerts / messages
    numberAlert: {
        ja: "数値を入力してください。",
        en: "Please enter a number."
    },
    errorOccurred: {
        ja: "エラーが発生しました: ",
        en: "An error occurred: "
    },
    alertNoDoc: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    // 8) Size panel
    sizePanelTitle: {
        ja: "サイズ",
        en: "Size"
    },
    widthLabel: {
        ja: "幅",
        en: "Width"
    },
    heightLabel: {
        ja: "高さ",
        en: "Height"
    },
    // 9) Anchor panel
    anchorPanelTitle: {
        ja: "基準",
        en: "Anchor"
    },
    anchorTopLeft: {
        ja: "左上",
        en: "Top-Left"
    },
    anchorCenter: {
        ja: "中央",
        en: "Center"
    }
};

/*
エラー整形ヘルパー / Error formatting helper
Illustrator/ExtendScript の Error から行番号などを含めて読みやすく整形。
*/
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


/*
共通エンジン名を使用 / Use a common engine name
複数スクリプト間で位置記憶を共有しつつ、key で保存先を分離します。
Share session state across scripts; separate each dialog by key.
*/
function _getSavedLoc(key) {
    return $.global[key] && $.global[key].length === 2 ? $.global[key] : null;
}

function _setSavedLoc(key, loc) {
    $.global[key] = [loc[0], loc[1]];
}

function _clampToScreen(loc) {
    try {
        var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
        var x = Math.max(vb[0] + 10, Math.min(loc[0], vb[2] - 10));
        var y = Math.max(vb[1] + 10, Math.min(loc[1], vb[3] - 10));
        return [x, y];
    } catch (e) {
        return loc;
    }
}



// -------------------------------
// 設定定数 / Configuration constants
// -------------------------------
var CONFIG = {
    dialogOpacity: 0.95,
    offsetX: 300
};


/*
ダイアログ表示 / Show dialog with live preview
*/
function showMarginDialog(defaultValue, unit) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    // スクリプト名＋バージョンでキーをネームスペース化 / Namespace the key by script name + version
    var dlgPositionKey = "__FitArtboardHeightToSelection_" + SCRIPT_VERSION + "__Dialog";
    if ($.global[dlgPositionKey] === undefined) $.global[dlgPositionKey] = null; // ensure slot
    var __savedLoc = _getSavedLoc(dlgPositionKey);

    // apply saved location (fallback to existing centering/offset if none)
    if (__savedLoc) {
        dlg.onShow = (function(prev) {
            return function() {
                try {
                    if (typeof prev === 'function') prev();
                } catch (_) {}
                dlg.location = _clampToScreen(__savedLoc);
            };
        })(dlg.onShow);
    }

    // save on move
    var __saveDlgLoc = function() {
        _setSavedLoc(dlgPositionKey, [dlg.location[0], dlg.location[1]]);
    };
    dlg.onMove = (function(prev) {
        return function() {
            try {
                if (typeof prev === 'function') prev();
            } catch (_) {}
            __saveDlgLoc();
        };
    })(dlg.onMove);
    /* ダイアログ位置と不透明度のカスタマイズ / Customize dialog offset and opacity */

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            dlg.layout.layout(true);
            var dialogWidth = dlg.bounds.width;
            var dialogHeight = dlg.bounds.height;

            var screenWidth = $.screens[0].right - $.screens[0].left;
            var screenHeight = $.screens[0].bottom - $.screens[0].top;

            var centerX = screenWidth / 2 - dialogWidth / 2;
            var centerY = screenHeight / 2 - dialogHeight / 2;

            dlg.location = [centerX + offsetX, centerY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    setDialogOpacity(dlg, CONFIG.dialogOpacity);
    if (!__savedLoc) {
        /* 初回のみセンターからのオフセットを適用 / Apply offset only on first run (no saved location) */
        shiftDialogPosition(dlg, CONFIG.offsetX, 0);
    }
    /* ダイアログ位置と不透明度のカスタマイズ: ここまで / Dialog offset & opacity: end */

    /* サイズ表示パネル / Size display panel */
    var sizePanel = dlg.add("panel", undefined, LABELS.sizePanelTitle[lang]);
    sizePanel.orientation = "row";
    sizePanel.alignChildren = ["left", "center"];
    sizePanel.margins = [15, 20, 15, 10];

    var sizeGroup = sizePanel.add("group");
    sizeGroup.orientation = "column"; // 縦並び
    sizeGroup.alignChildren = ["left", "center"];
    sizeGroup.spacing = 6;

    // Width row
    var wRow = sizeGroup.add("group");
    wRow.orientation = "row";
    var wLabel = wRow.add("statictext", undefined, LABELS.widthLabel[lang] + "：");
    var wValue = wRow.add("edittext", undefined, "-");
    wValue.characters = 5;
    var wUnitLabel = wRow.add("statictext", undefined, unit);

    // Height row
    var hRow = sizeGroup.add("group");
    hRow.orientation = "row";
    var hLabel = hRow.add("statictext", undefined, LABELS.heightLabel[lang] + "：");
    var hValue = hRow.add("edittext", undefined, "-");
    hValue.characters = 5;
    var hUnitLabel = hRow.add("statictext", undefined, unit);

    // --- Align label widths and right-justify ---
    try {
        wLabel.justify = "right";
        hLabel.justify = "right";
        var g = (sizePanel && sizePanel.graphics) ? sizePanel.graphics : dlg.graphics;

        function _tw(s) {
            try {
                var m = g.measureString(String(s));
                return Math.ceil(m[0]);
            } catch (e) {
                return String(s).length * 7;
            }
        }
        var maxLabelW = Math.max(_tw(wLabel.text), _tw(hLabel.text)) + 6; // small padding
        // Fix both labels to the same width
        wLabel.minimumSize = [maxLabelW, 0];
        wLabel.maximumSize = [maxLabelW, 1000];
        hLabel.minimumSize = [maxLabelW, 0];
        hLabel.maximumSize = [maxLabelW, 1000];
    } catch (_) {}

    // 初期フォーカス：幅 / Set initial focus to Width field on open (preserve any existing onShow)
    dlg.onShow = (function(prev) {
        return function() {
            try {
                if (typeof prev === 'function') prev();
            } catch (_) {}
            try {
                wValue.active = true;
            } catch (_) {}
        };
    })(dlg.onShow);

    /* 現在のアートボードrectと全アートボードrectを保存（プレビュー用に復元） / Save current and all artboard rects for preview restore */
    var abIndex = app.activeDocument.artboards.getActiveArtboardIndex();
    var originalRects = [];
    for (var i = 0; i < app.activeDocument.artboards.length; i++) {
        originalRects.push(app.activeDocument.artboards[i].artboardRect.slice());
    }

    // --- Helpers to display AB size in current unit ---
    function fromPt(valPt, unitStr) {
        try {
            var u = (unitStr || 'pt').toString().toLowerCase();
            if (u === 'h' || u === 'q') u = 'mm';
            if (_UNIT_FACTORS_PT.hasOwnProperty(u)) return valPt / _UNIT_FACTORS_PT[u];
            // fallback via UnitValue
            var alt = (u === 'inch') ? 'in' : (u === 'pica' ? 'pc' : u);
            return new UnitValue(valPt, 'pt').as(alt);
        } catch (e) {
            return valPt;
        }
    }

    function fmt(val, u) {
        if (u === 'px' || u === 'pt') return String(Math.round(val));
        return String(Math.round(val * 100) / 100);
    }

    function updateSizePanelDisplay() {
        try {
            var abIdx = app.activeDocument.artboards.getActiveArtboardIndex();
            var r = app.activeDocument.artboards[abIdx].artboardRect; // [L,T,R,B]
            var wPt = Math.abs(r[2] - r[0]);
            var hPt = Math.abs(r[1] - r[3]);
            var wU = fromPt(wPt, unit);
            var hU = fromPt(hPt, unit);
            wValue.text = fmt(wU, unit);
            hValue.text = fmt(hU, unit);
        } catch (e) {
            wValue.text = "-";
            hValue.text = "-";
        }
    }


    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 15;

    /* 基準パネル / Anchor selection panel */
    var anchorPanel = dlg.add("panel", undefined, LABELS.anchorPanelTitle[lang]);
    anchorPanel.orientation = "row";
    anchorPanel.alignChildren = ["left", "top"];
    anchorPanel.margins = [15, 20, 15, 10];

    var anchorGroup = anchorPanel.add("group");
    anchorGroup.orientation = "row"; // 横並び
    anchorGroup.alignChildren = ["left", "center"];
    anchorGroup.spacing = 12;

    var radioAnchorTopLeft = anchorGroup.add("radiobutton", undefined, LABELS.anchorTopLeft[lang]);
    var radioAnchorCenter = anchorGroup.add("radiobutton", undefined, LABELS.anchorCenter[lang]);
    radioAnchorTopLeft.alignment = "left";
    radioAnchorCenter.alignment = "left";

    // デフォルトは左上 / Default: Top-Left
    radioAnchorTopLeft.value = true;

    radioAnchorTopLeft.onClick = applyResizePreview;
    radioAnchorCenter.onClick = applyResizePreview;

    /* 対象パネル / Target selection panel */
    var targetPanel = dlg.add("panel", undefined, LABELS.targetPanelTitle[lang]);
    targetPanel.orientation = "row";
    targetPanel.alignChildren = ["left", "top"];
    targetPanel.margins = [15, 20, 15, 10];

    var targetGroup = targetPanel.add("group");
    targetGroup.orientation = "column";
    targetGroup.alignChildren = ["left", "top"];

    var radioActive = targetGroup.add("radiobutton", undefined, LABELS.targetActiveArtboard[lang]);
    var radioAll = targetGroup.add("radiobutton", undefined, LABELS.targetAllArtboards[lang]);
    radioActive.alignment = "left";
    radioAll.alignment = "left";

    // アートボードが1つなら「すべてのアートボード」をディム / Dim "All artboards" if only one artboard
    var __abCount = app.activeDocument.artboards.length;
    if (__abCount <= 1) {
        radioAll.enabled = false;
        radioActive.value = true; // ensure default stays active
    }

    // デフォルト：作業アートボードのみ / Default: active artboard only
    radioActive.value = true;


    /* 再描画をデバウンス / Throttle redraw to reduce jank */
    var __redrawLast = 0;
    var __redrawInterval = 40; // ms（約25fpsで十分）
    function throttledRedraw() {
        try {
            var now = (new Date()).getTime();
            if (now - __redrawLast >= __redrawInterval) {
                app.redraw();
                __redrawLast = now;
            }
        } catch (e) {
            // フォールバック：念のためプレビューが壊れないように
            try {
                app.redraw();
            } catch (_) {}
        }
    }

    function applyResizePreview() {
        function _parseWH() {
            var w = parseFloat(String(wValue.text).replace(/[^0-9.\-]/g, ''));
            var h = parseFloat(String(hValue.text).replace(/[^0-9.\-]/g, ''));
            if (isNaN(w) || isNaN(h)) return null;
            if (w < 0) w = 0;
            if (h < 0) h = 0;
            return {
                wPt: toPt(w, unit),
                hPt: toPt(h, unit)
            };
        }
        var wh = _parseWH();
        if (!wh) return; // ignore until valid

        var scopeAll = (typeof radioAll !== 'undefined' && radioAll && radioAll.value === true);
        var abCount = app.activeDocument.artboards.length;

        function _resizeOne(abIdx) {
            var r = app.activeDocument.artboards[abIdx].artboardRect.slice(); // [L,T,R,B]
            var wPt = wh.wPt,
                hPt = wh.hPt;
            var newRect;
            if (radioAnchorTopLeft && radioAnchorTopLeft.value === true) {
                // 左上基準：L/T固定で右・下に展開
                var L = r[0],
                    T = r[1];
                newRect = [L, T, L + wPt, T - hPt];
            } else {
                // 中央基準（デフォルト）
                var cx = (r[0] + r[2]) / 2;
                var cy = (r[1] + r[3]) / 2;
                var halfW = wPt / 2;
                var halfH = hPt / 2;
                newRect = [cx - halfW, cy + halfH, cx + halfW, cy - halfH];
            }
            app.activeDocument.artboards[abIdx].artboardRect = newRect;
        }

        if (scopeAll) {
            for (var i = 0; i < abCount; i++) _resizeOne(i);
        } else {
            var idx = app.activeDocument.artboards.getActiveArtboardIndex();
            _resizeOne(idx);
        }
        throttledRedraw();
        updateSizePanelDisplay();
    }

    function updatePreview() {
        applyResizePreview();
    }

    wValue.onChanging = applyResizePreview;
    hValue.onChanging = applyResizePreview;
    changeValueByArrowKey(wValue, applyResizePreview);
    changeValueByArrowKey(hValue, applyResizePreview);

    radioActive.onClick = applyResizePreview;
    radioAll.onClick = applyResizePreview;

    /* ボタングループ / Button group */
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    btnGroup.margins = [0, 5, 0, 0];
    var cancelBtn = btnGroup.add("button", undefined, LABELS.cancelBtn[lang], {
        name: "cancel"
    });
    var okBtn = btnGroup.add("button", undefined, LABELS.okBtn[lang], {
        name: "ok"
    });

    /* 閉じる時に位置を記憶 / Persist location on close */
    var result = null;

    okBtn.onClick = function() {
        applyResizePreview();
        result = {
            targetScope: (radioActive.value ? "active" : "all")
        };
        dlg.close();
    };
    cancelBtn.onClick = function() {
        /* プレビューで変更した全アートボードrectを元に戻す / Restore all artboard rects after preview */
        for (var i = 0; i < app.activeDocument.artboards.length; i++) {
            app.activeDocument.artboards[i].artboardRect = originalRects[i].slice();
        }
        updateSizePanelDisplay();
        throttledRedraw();
        dlg.close();
    };

    okBtn.onClick = (function(prev) {
        return function() {
            try {
                __saveDlgLoc();
            } catch (_) {}
            if (typeof prev === 'function') return prev();
            dlg.close(1);
        };
    })(okBtn.onClick);

    if (typeof cancelBtn !== 'undefined' && cancelBtn) {
        cancelBtn.onClick = (function(prev) {
            return function() {
                try {
                    __saveDlgLoc();
                } catch (_) {}
                if (typeof prev === 'function') return prev();
                dlg.close(0);
            };
        })(cancelBtn.onClick);
    }
    updateSizePanelDisplay();
    applyResizePreview();
    dlg.show();
    return result;
}

/* メイン処理 / Main process */
function main() {
    var userInput = showMarginDialog(null, _detectUnitString());
    if (!userInput) return;
    // All resizing already applied in preview; nothing to do here.
}

/*
選択アイテムの正規化 / Normalize selected items
- クリップグループ(GroupItem.clipped=true)はクリッピングパスのみを採用
- それ以外はそのまま
This unifies item collection for preview/runtime to avoid duplication.
*/
function collectEffectiveItems(items) {
    var out = [];
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it && it.typename === "GroupItem" && it.clipped) {
            // クリップグループはクリッピングパスのみ採用 / For clipped group, include only the clipping path
            try {
                for (var j = 0; j < it.pageItems.length; j++) {
                    var child = it.pageItems[j];
                    if (child.clipping) {
                        out.push(child);
                        break;
                    }
                }
            } catch (e) {
                /* ignore */
            }
        } else {
            out.push(it);
        }
    }
    return out;
}

/*
テキストの一時アウトライン化ユーティリティ / Temporary outlining utility for text frames
- outlineTextItems(items):
    与えられた items を走査し、TextFrame は duplicate+hidden → createOutline() で図形化し、
    それ以外はそのまま集約して bounds 計算用の tempItems にまとめる。
    戻り値: { tempItems, originals, outlines }
- cleanupOutlinedText(outlines, originals):
    createOutline で作成したアウトラインを削除し、hidden にした元 TextFrame を再表示。
*/
function outlineTextItems(items) {
    var originals = [];
    var outlines = [];
    var tempItems = [];
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        if (it.typename === "TextFrame") {
            try {
                var dup = it.duplicate();
                dup.hidden = true;
                originals.push(dup);
                var out = it.createOutline();
                if (out && out.length && out.length > 0) {
                    for (var q = 0; q < out.length; q++) {
                        outlines.push(out[q]);
                        tempItems.push(out[q]);
                    }
                } else if (out) {
                    outlines.push(out);
                    tempItems.push(out);
                }
            } catch (e) {
                // 失敗時は元のテキストをそのまま使う（visibleBoundsの誤差は許容）
                tempItems.push(it);
            }
        } else {
            tempItems.push(it);
        }
    }
    return {
        tempItems: tempItems,
        originals: originals,
        outlines: outlines
    };
}

function cleanupOutlinedText(outlines, originals) {
    for (var i = 0; i < outlines.length; i++) {
        try {
            outlines[i].remove();
        } catch (e) {}
    }
    for (var j = 0; j < originals.length; j++) {
        try {
            originals[j].hidden = false;
        } catch (e) {}
    }
}

/*
作業アートボード内のオブジェクトを収集 / Collect items within an artboard
- usePreviewBounds=true: visibleBounds（塗り/線を含む）で判定
- usePreviewBounds=false: geometricBounds（パス形状）で判定
*/
function getItemsInArtboard(abIndex, usePreviewBounds) {
    var abRect = artboards[abIndex].artboardRect; // [L, T, R, B]
    var result = [];
    var n = doc.pageItems.length;
    for (var i = 0; i < n; i++) {
        var it = doc.pageItems[i];
        try {
            if (!it || it.locked || it.hidden || it.guides) continue;
            var b = getBounds(it, usePreviewBounds); // [L, T, R, B]
            // 交差判定（どこか一辺でも外れていれば非交差）
            var outside = (b[2] <= abRect[0]) || (b[0] >= abRect[2]) || (b[3] >= abRect[1]) || (b[1] <= abRect[3]);
            if (!outside) result.push(it);
        } catch (e) {
            /* ignore */
        }
    }
    return result;
}

/*
単位→pt変換ユーティリティ（係数キャッシュ版） / Unit to pt conversion with cached factors
- よく使う単位は乗算のみで高速化（UnitValue生成を回避）
- 未サポート単位は従来どおり UnitValue にフォールバック
*/
var _UNIT_FACTORS_PT = {
    pt: 1,
    px: 1, // Illustrator既定：1px ≒ 1pt（72ppi基準）
    mm: 72 / 25.4, // 2.834645669...
    cm: 72 / 2.54, // 28.34645669...
    inch: 72,
    "in": 72,
    pica: 12, // 1pc = 12pt
    pc: 12
};

function toPt(val, unit) {
    try {
        var n = Number(val);
        if (isNaN(n)) return NaN;
        var u = (unit || 'pt').toString().toLowerCase();
        // 正規化 / normalize
        if (u === 'h' || u === 'q') u = 'mm';
        // 係数キャッシュがあれば乗算 / fast path
        if (_UNIT_FACTORS_PT.hasOwnProperty(u)) {
            return n * _UNIT_FACTORS_PT[u];
        }
        // フォールバック（稀な単位） / fallback to UnitValue for rare units
        if (u === 'inch') u = 'in';
        if (u === 'pica') u = 'pc';
        return new UnitValue(n, u).as('pt');
    } catch (e) {
        return NaN;
    }
}


/* 選択オブジェクト群から最大のバウンディングボックスを取得 / Get maximum bounding box from multiple items */
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

/* オブジェクトのバウンディングボックスを取得 / Get bounding box of a single object
   usePreviewBounds=true なら visibleBounds（プレビュー境界: 塗り/線を含む）
   usePreviewBounds=false なら geometricBounds（幾何境界: パス外形のみ） */
function getBounds(item, usePreviewBounds) {
    return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
}

/* edittextに矢印キーで値を増減する機能を追加 / Add arrow key increment/decrement to edittext */
function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function(event) {
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        // robust parse: keep only [0-9.-]
        var raw = String(editText.text);
        var num = parseFloat(raw.replace(/[^0-9.\-]/g, ""));
        if (isNaN(num)) num = 0;

        var kb = ScriptUI.environment.keyboardState;
        var step = 1;
        // Option(Alt) = 0.1, Shift = 10. If both held, Option優先で0.1。
        if (kb.altKey || kb.optionKey) {
            step = 0.1;
        } else if (kb.shiftKey) {
            step = 10;
        }

        var delta = (event.keyName === "Up") ? step : -step;
        var next = num + delta;
        if (next < 0) next = 0; // prevent negative

        // ↑↓操作では常に整数へ丸める
        next = Math.round(next);

        event.preventDefault();
        editText.text = String(next);
        if (typeof onUpdate === "function") onUpdate(editText.text);
    });
}

try {
    main();
} catch (e) {
    try {
        $.writeln("[FitArtboardWithMargin] ERROR: " + formatError(e));
    } catch (_) {}
    alert(LABELS.errorOccurred[lang] + formatError(e));
}

app.selectTool("Adobe Select Tool");