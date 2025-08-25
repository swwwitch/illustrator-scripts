#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

var SCRIPT_VERSION = "v1.2";

/*
### スクリプト名：

Fit Artboard Height to Selection (Same Width)

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択したオブジェクトのバウンディングボックスに **上下マージン** を加味して、作業中のアートボードの **高さのみ** を自動調整（左右＝幅は固定）。
- **選択がない場合**は、各アートボード内のオブジェクトを対象にして、**すべてのアートボード**の高さを個別に調整。

### 主な機能：

- ライブプレビュー（軽量化のため `visibleBounds` 固定）
- プレビューの再描画をデバウンス（軽量化）
- プレビュー境界の切替（確定時のみ `visible`/`geometric` を反映）
- 〈対象〉パネル：  
  ・「作業アートボードのみ」＝選択ありはアクティブABのみ／選択なしは全AB（従来挙動）  
  ・「すべてのアートボード」＝選択を無視して全AB
- 文字オブジェクトは一時アウトライン化して正確な境界を計測（処理後に復元）
- 単位ラベル表示、ダイアログ位置の記憶（バージョン別キー）

### 処理の流れ：

1) ダイアログで上下マージンを指定  
2) プレビューでアートボードの上下のみ更新（左右は保持）  
3) OKで確定、Cancelで元に戻す

### note：

- プレビューは高速化のため `visibleBounds` 固定です。確定時に設定の `visible/geometric` が反映されます。

### 更新履歴：

- v1.2 (20250825) : 〈対象〉パネル追加（全AB/作業AB 切替）。プレビューを `visibleBounds` 固定＋`app.redraw()` をデバウンス。単位ラベルとローカライズ整理。ダイアログ位置の保存キーをバージョン別に。
- v1.1 (20250825) : アートボード **高さのみ** 調整に簡素化。未定義 `doc` 参照を修正。コメント整理と説明文更新。
- v1.0 (20250825) : 初期バージョン

---

### Script name (EN):

Fit Artboard Height to Selection (Same Width)

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Adjust the **height only** of the active artboard to the selection’s bounds plus **vertical margin** (width/left/right are preserved).
- **When nothing is selected,** adjust the **height of all artboards** individually based on items inside each artboard.

### Key features:

- Live preview (uses `visibleBounds` for speed)
- Throttled `app.redraw()` during preview
- Preview-bounds toggle (apply `visible`/`geometric` only on commit)
- Target panel:  
  · “Active artboard only” = with selection → active AB, without selection → all AB (legacy behavior)  
  · “All artboards” = ignore selection and process all AB  
- Temporary outlining for text to measure exact bounds (restored after)
- Unit label next to input, dialog position remembered (namespaced per version)

### Flow:

1) Enter vertical margin in the dialog  
2) Preview updates top/bottom only (left/right preserved)  
3) Confirm to commit or cancel to restore

### note:

- Preview always uses `visibleBounds` for performance; final commit respects the chosen `visible/geometric` option.

### Changelog:

- v1.2 (2025-08-25): Added Target panel (All/Active). Preview fixed to `visibleBounds` with throttled redraw. Unit label & localization cleanup. Version-scoped dialog position key.
- v1.1 (2025-08-25): Simplified to height-only adjustment. Fixed undefined `doc`. Cleaned comments and docs.
- v1.0 (2025-08-25): Initial version.
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
    // 2) Panel title
    marginLabel: {
        ja: "マージン",
        en: "Margin"
    },
    // 3) Vertical margin label
    marginVertical: {
        ja: "上下",
        en: "Vertical"
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
    // 5) Preview-bounds checkbox
    previewBounds: {
        ja: "プレビュー境界",
        en: "Preview bounds"
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
    supportedUnits: ['inch', 'mm', 'pt', 'pica', 'cm', 'H', 'px'],
    defaultMarginByUnit: {
        mm: '5',
        px: '20',
        pt: '10',
        _fallback: '0'
    },
    previewBoundsDefault: true,
    dialogOpacity: 0.95,
    offsetX: 300
};

function getDefaultMargin(unit) {
    return CONFIG.defaultMarginByUnit.hasOwnProperty(unit) ?
        CONFIG.defaultMarginByUnit[unit] :
        CONFIG.defaultMarginByUnit._fallback;
}

// (moved here for correct initialization order)
var marginUnit = _detectUnitString();
var defaultMarginValue = getDefaultMargin(marginUnit);

/*
マージンダイアログ表示 / Show margin input dialog with live preview
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

    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 15;


    /* マージン入力パネル / Margin input panel */
    var marginPanel = dlg.add("panel", undefined, LABELS.marginLabel[lang]);
    marginPanel.orientation = "row";
    marginPanel.alignChildren = ["fill", "top"];
    marginPanel.margins = [15, 20, 15, 10];


    /* 上下マージン入力欄 / Vertical margin input */
    var verticalGroup = marginPanel.add("group");
    verticalGroup.orientation = "row";
    var labelV = verticalGroup.add("statictext", undefined, LABELS.marginVertical[lang] + ":");
    var inputV = verticalGroup.add("edittext", undefined, defaultValue);
    var unitLabelV = verticalGroup.add("statictext", undefined, unit);
    unitLabelV.justify = "left";
    inputV.characters = 4;

    /* 対象パネル / Target selection panel */
    var targetPanel = dlg.add("panel", undefined, LABELS.targetPanelTitle[lang]);
    targetPanel.orientation = "row";
    targetPanel.alignChildren = ["left", "top"];
    targetPanel.margins = [15, 15, 15, 10];

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

    // 将来のロジック拡張に備えてプレビューだけ更新（挙動は現状維持）
    radioActive.onClick = function() {
        updatePreview(inputV.text);
    };
    radioAll.onClick = function() {
        updatePreview(inputV.text);
    };



    /* 現在のアートボードrectと全アートボードrectを保存（プレビュー用に復元） / Save current and all artboard rects for preview restore */
    var abIndex = app.activeDocument.artboards.getActiveArtboardIndex();
    var originalRects = [];
    for (var i = 0; i < app.activeDocument.artboards.length; i++) {
        originalRects.push(app.activeDocument.artboards[i].artboardRect.slice());
    }


    function previewSelection(previewMarginV) {
        var sel = app.activeDocument.selection;
        var hasSel = (sel && sel.length > 0);
        var scopeAll = (typeof radioAll !== 'undefined' && radioAll && radioAll.value === true);

        // --- If "All artboards" is selected, ignore selection and adjust ALL ---
        if (scopeAll) {
            for (var i = 0; i < app.activeDocument.artboards.length; i++) {
                var itemsAll = getItemsInArtboard(i, true); // preview uses visibleBounds for speed
                if (!itemsAll || itemsAll.length === 0) {
                    app.activeDocument.artboards[i].artboardRect = originalRects[i].slice();
                    continue;
                }
                var tempAll = collectEffectiveItems(itemsAll);
                if (tempAll.length === 0) continue;

                var bAll = getMaxBounds(tempAll, true);
                var tAll = bAll[1] + previewMarginV;
                var btmAll = bAll[3] - previewMarginV;

                var rectAll = originalRects[i].slice();
                rectAll[1] = tAll;
                rectAll[3] = btmAll;
                app.activeDocument.artboards[i].artboardRect = rectAll;
            }
            throttledRedraw();
            return;
        }

        // --- "Active artboard only" (current behavior): selection → active AB, no selection → all AB ---
        if (hasSel) {
            var tempSel = collectEffectiveItems(sel);
            if (tempSel.length === 0) return;

            var pb = getMaxBounds(tempSel, true);
            var newTop = pb[1] + previewMarginV;
            var newBottom = pb[3] - previewMarginV;

            var abRect = originalRects[abIndex].slice();
            abRect[1] = newTop;
            abRect[3] = newBottom;
            app.activeDocument.artboards[abIndex].artboardRect = abRect;
            throttledRedraw();
            return;
        }

        // no selection → adjust all AB (existing behavior)
        for (var j = 0; j < app.activeDocument.artboards.length; j++) {
            var items = getItemsInArtboard(j, true);
            if (!items || items.length === 0) {
                app.activeDocument.artboards[j].artboardRect = originalRects[j].slice();
                continue;
            }
            var temp = collectEffectiveItems(items);
            if (temp.length === 0) continue;

            var b = getMaxBounds(temp, true);
            var t = b[1] + previewMarginV;
            var bt = b[3] - previewMarginV;

            var rect = originalRects[j].slice();
            rect[1] = t;
            rect[3] = bt;
            app.activeDocument.artboards[j].artboardRect = rect;
        }
        throttledRedraw();
    }

    /*
    プレビュー更新関数 / Update artboard preview for dialog
    入力値・対象に応じてアートボードを一時的に調整 / Temporarily adjust artboard for preview
    */
    function updatePreview(valueV) {
        var parsed = parseMarginSingle(valueV, unit);
        if (!parsed.valid) return;
        var previewMarginV = parsed.vPt;
        previewSelection(previewMarginV);
    }

    /* 入力欄で矢印キーによる増減を可能に / Enable arrow key increment/decrement in input */
    changeValueByArrowKey(inputV, function(val) {
        updatePreview(inputV.text);
    });
    inputV.active = true;
    inputV.onChanging = function() {
        updatePreview(inputV.text);
    };

    /* 最下部に配置するチェックボックス / Checkbox above buttons */
    var previewBoundsCheckbox = dlg.add("checkbox", undefined, LABELS.previewBounds[lang]);
    previewBoundsCheckbox.alignment = "center";
    previewBoundsCheckbox.value = CONFIG.previewBoundsDefault; // デフォルト
    previewBoundsCheckbox.margins = [0, 5, 0, 0];
    /* チェック切替でプレビュー更新 / Refresh preview on toggle */
    previewBoundsCheckbox.onClick = function() {
        updatePreview(inputV.text);
    };

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

    okBtn.onClick = function() {
        var parsed = parseMarginSingle(inputV.text, unit);
        if (parsed && parsed.valid) {
            result = {
                marginV: inputV.text,
                previewBounds: previewBoundsCheckbox.value,
                targetScope: (radioActive.value ? "active" : "all") // logic to be applied later
            };
            updatePreview(result.marginV);
            dlg.close();
        } else {
            alert(LABELS.numberAlert[lang]);
        }
    };
    cancelBtn.onClick = function() {
        /* プレビューで変更した全アートボードrectを元に戻す / Restore all artboard rects after preview */
        for (var i = 0; i < app.activeDocument.artboards.length; i++) {
            app.activeDocument.artboards[i].artboardRect = originalRects[i].slice();
        }
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
    updatePreview(inputV.text);
    dlg.show();
    return result;
}

/* メイン処理 / Main process */
function main() {
    var selectedItems = doc.selection;

    var userInput = showMarginDialog(defaultMarginValue, marginUnit);
    if (!userInput) return;

    var marginV = parseFloat(userInput.marginV);
    var marginVInPoints = toPt(marginV, marginUnit);
    var scopeAll = (userInput.targetScope === "all");
    var hasSelection = (selectedItems && selectedItems.length > 0);

    // Helper: finalize per-artboard height from items
    function _finalizeArtboardHeightByItems(abIdx, items) {
        if (!items || items.length === 0) return false;

        var outlined = outlineTextItems(items);
        var temp = outlined.tempItems;
        if (!temp || temp.length === 0) {
            cleanupOutlinedText(outlined.outlines, outlined.originals);
            return false;
        }

        var b = getMaxBounds(temp, userInput.previewBounds);
        var newTop = Math.round(b[1] + marginVInPoints);
        var newBottom = Math.round(b[3] - marginVInPoints);

        var rect = artboards[abIdx].artboardRect.slice();
        rect[1] = newTop;
        rect[3] = newBottom;

        cleanupOutlinedText(outlined.outlines, outlined.originals);
        artboards[abIdx].artboardRect = rect;
        return true;
    }

    // Case 1: target = "all" → ignore selection, process all AB
    if (scopeAll) {
        for (var i = 0; i < artboards.length; i++) {
            var itemsAll = getItemsInArtboard(i, userInput.previewBounds);
            _finalizeArtboardHeightByItems(i, itemsAll);
        }
        return;
    }

    // Case 2: target = "active" (current behavior)
    if (!hasSelection) {
        // No selection → all AB (existing behavior)
        for (var j = 0; j < artboards.length; j++) {
            var itemsNoSel = getItemsInArtboard(j, userInput.previewBounds);
            _finalizeArtboardHeightByItems(j, itemsNoSel);
        }
        return;
    }

    // Selection exists → adjust active AB only
    var abIndexFinal = artboards.getActiveArtboardIndex();
    var itemsSel = collectEffectiveItems(selectedItems);
    // outline & finalize just for active artboard
    _finalizeArtboardHeightByItems(abIndexFinal, itemsSel);
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

/*
入力検証（縦のみ） / Input validation (vertical only)
- 単位変換まで実施し pt 値も返す
*/
function parseMarginSingle(text, unit) {
    var v = parseFloat(text);
    if (isNaN(v)) return {
        valid: false
    };
    var vPt = toPt(v, unit);
    if (isNaN(vPt)) return {
        valid: false
    };
    return {
        valid: true,
        v: v,
        vPt: vPt
    };
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
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            /* Shift時は10の倍数へスナップ / Snap to multiples of 10 when holding Shift */
            if (keyboard.shiftKey) {
                if (event.keyName == "Up") {
                    // 次の上位10の倍数へ（5→10, 10→20）
                    var upBase = Math.ceil(value / 10) * 10;
                    value = (upBase === value) ? value + 10 : upBase;
                } else {
                    // 次の下位10の倍数へ（15→10, 10→0, 5→0）
                    var downBase = Math.floor(value / 10) * 10;
                    value = (downBase === value) ? value - 10 : downBase;
                    if (value < 0) value = 0; // Prevent negative
                }
            } else {
                var delta = event.keyName == "Up" ? 1 : -1;
                value += delta;
            }

            event.preventDefault();
            editText.text = value;
            if (typeof onUpdate === "function") {
                onUpdate(editText.text);
            }
        }
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