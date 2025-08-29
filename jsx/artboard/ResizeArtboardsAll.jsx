#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### スクリプト名：

アートボードサイズを調整 / Fit Artboard Size to Specified Width/Height

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- ダイアログで指定した「幅」「高さ」に、アートボードをライブプレビューしながら変形します。
- 選択がない場合は、各アートボード内のオブジェクトを基準に、すべてのアートボードを個別に調整します。

### 主な機能：

- ライブプレビュー（app.redraw のデバウンスで高速化）
- 対象のアートボード（作業のみ／すべて／指定（1始まりの範囲・カンマ列：例 1-3 / 1,3 / 2-4,7））
- 基準点の切替（左上／中央）
- 単位はドキュメントの定規単位に追従（px時は左上座標を整数にスナップ）
- ダイアログ位置・不透明度の記憶（セッション間で復元）
- 矢印キー操作：↑↓=±1、Shift+↑↓=10の倍数にスナップ、Option(Alt)+↑↓=±0.1（最終的に整数化）

### 処理の流れ：

1. ダイアログで幅・高さを入力（必要なら対象・基準点を選択）
2. プレビューで即時にアートボードを更新
3. OKで確定、Cancelで元に戻す

### note：

- 「指定」は 1 始まりで解釈し、内部で 0 始まりに変換します。
- px/pt 表示は整数、小数単位（mm/cm/inch/pica）は小数2桁で表示します。

### 更新履歴：

- v1.0 (20250829) : 初期バージョン

---

### Script Name:

Fit Artboard Size to Specified Width/Height

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Resize artboards to the specified width/height with live preview.
- When nothing is selected, each artboard is adjusted individually based on items inside it.

### Key Features:

- Live preview (debounced app.redraw for performance)
- Target artboards: Active only / All / Specify (1-based ranges & lists, e.g., 1-3 / 1,3 / 2-4,7)
- Anchor: Top-Left / Center
- Units follow the document ruler (snap top-left to integer when in px)
- Dialog position & opacity persistence across sessions
- Arrow keys: Up/Down = ±1, Shift = snap to multiples of 10, Option(Alt) = ±0.1 (rounded to integer at commit)

### Flow:

1. Enter width/height (optionally choose target & anchor)
2. See instant preview of artboard updates
3. Press OK to commit, Cancel to restore

### Changelog:

- v1.0 (2025-08-29): Initial release

*/

var SCRIPT_VERSION = "v1.0";

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
        ja: "対象のアートボード",
        en: "Target Artboard"
    },
    targetActiveArtboard: {
        ja: "作業アートボードのみ",
        en: "Active artboard only"
    },
    targetAllArtboards: {
        ja: "すべてのアートボード",
        en: "All artboards"
    },
    targetSpecify: {
        ja: "指定",
        en: "Specify"
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

function _detectUnitString() {
    try {
        var d = app.activeDocument;
        var ru = d.rulerUnits;
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
    var dlgPositionKey = "__ResizeArtboardsAll_" + SCRIPT_VERSION + "__Dialog";
    if ($.global[dlgPositionKey] === undefined) $.global[dlgPositionKey] = null; // ensure slot
    var __savedLoc = _getSavedLoc(dlgPositionKey);

    // apply saved location (fallback to existing centering/offset if none)
    if (__savedLoc) {
        dlg.onShow = (function(prev) {
            return function() {
                if (typeof prev === 'function') prev();
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
            if (typeof prev === 'function') prev();
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
    // === Two-column layout container ===
    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 15;

    var cols = dlg.add("group");
    cols.orientation = "row";
    cols.alignChildren = ["fill", "top"]; // fill width per column, align to top
    cols.spacing = 15;

    var leftCol = cols.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "fill";
    leftCol.spacing = 10;

    var rightCol = cols.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = "fill";
    rightCol.spacing = 10;

    var sizePanel = leftCol.add("panel", undefined, LABELS.sizePanelTitle[lang] + "（" + unit + "）");
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
    // var wUnitLabel = wRow.add("statictext", undefined, unit);

    // Height row
    var hRow = sizeGroup.add("group");
    hRow.orientation = "row";
    var hLabel = hRow.add("statictext", undefined, LABELS.heightLabel[lang] + "：");
    var hValue = hRow.add("edittext", undefined, "-");
    hValue.characters = 5;
    // var hUnitLabel = hRow.add("statictext", undefined, unit);

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
            if (typeof prev === 'function') prev();
            wValue.active = true;
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

    /* 基準パネル / Anchor selection panel */
    var anchorPanel = leftCol.add("panel", undefined, LABELS.anchorPanelTitle[lang]);
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
    var targetPanel = rightCol.add("panel", undefined, LABELS.targetPanelTitle[lang]);
    targetPanel.orientation = "row";
    targetPanel.alignChildren = ["left", "top"];
    targetPanel.margins = [15, 20, 15, 10];

    var targetGroup = targetPanel.add("group");
    targetGroup.orientation = "column";
    targetGroup.alignChildren = ["left", "top"];

    var radioActive = targetGroup.add("radiobutton", undefined, LABELS.targetActiveArtboard[lang]);
    var radioAll = targetGroup.add("radiobutton", undefined, LABELS.targetAllArtboards[lang]);

    // 指定（縦配置：下に入力） / Specify (vertical: input below)
    var radioSpecify = targetGroup.add("radiobutton", undefined, LABELS.targetSpecify[lang]);
    radioSpecify.alignment = "left";

    // 入力フィールド（ロジックは後で） / Input field (logic to be added later)
    var inputSpecify = targetGroup.add("edittext", undefined, "");
    inputSpecify.characters = 12; // visible width
    inputSpecify.alignment = "left";
    inputSpecify.helpTip = (lang === "ja") ? "例: 1-3 / 1,3 / 2-4,7（1始まり）" : "e.g., 1-3 / 1,3 / 2-4,7 (1-based)";

    // 指定入力はデフォルト無効（Active/All選択時）
    inputSpecify.enabled = false;

    // "1-3,5,7-9" 形式を 0-based index 配列にパース / Parse to zero-based indices
    function parseSpecifyInput(txt, abCount) {
        var s = String(txt || "").replace(/\s+/g, "");
        if (!s) return [];
        var parts = s.split(',');
        var out = {};
        for (var i = 0; i < parts.length; i++) {
            var p = parts[i];
            if (!p) continue;
            var m = p.match(/^(\d+)-(\d+)$/); // range a-b
            if (m) {
                var a = parseInt(m[1], 10);
                var b = parseInt(m[2], 10);
                if (isNaN(a) || isNaN(b)) continue;
                if (a > b) { var t = a; a = b; b = t; }
                for (var v = a; v <= b; v++) {
                    var idx = v - 1; // 1-based -> 0-based
                    if (idx >= 0 && idx < abCount) out[idx] = true;
                }
            } else {
                var n = parseInt(p, 10);
                if (!isNaN(n)) {
                    var idx2 = n - 1; // 1-based -> 0-based
                    if (idx2 >= 0 && idx2 < abCount) out[idx2] = true;
                }
            }
        }
        // return unique, sorted
        var arr = [];
        for (var k in out) if (out.hasOwnProperty(k)) arr.push(parseInt(k,10));
        arr.sort(function(a,b){return a-b;});
        return arr;
    }

    function getSpecifiedIndices() {
        return parseSpecifyInput(inputSpecify.text, app.activeDocument.artboards.length);
    }

    function updateSpecifyEnabled() {
        var useSpecify = (radioSpecify && radioSpecify.value === true);
        if (inputSpecify) inputSpecify.enabled = !!useSpecify;
    }

    radioActive.alignment = "left";
    radioAll.alignment = "left";

    // ラジオ切替で有効/無効を反映
    radioActive.onClick = function(){ updateSpecifyEnabled(); applyResizePreview(); };
    radioAll.onClick    = function(){ updateSpecifyEnabled(); applyResizePreview(); };
    radioSpecify.onClick= function(){ updateSpecifyEnabled(); applyResizePreview(); };

    // 入力中もプレビュー反映（ロジックは選択がSpecifyのときのみ効く）
    inputSpecify.onChanging = function(){ if (radioSpecify.value) applyResizePreview(); };
    inputSpecify.onChange   = function(){ if (radioSpecify.value) applyResizePreview(); };

    // 初期状態を反映
    updateSpecifyEnabled();

    // アートボードが1つなら「すべてのアートボード」「指定」をディム / Dim "All artboards" and "Specify" if only one artboard
    var __abCount = app.activeDocument.artboards.length;
    if (__abCount <= 1) {
        radioAll.enabled = false;
        if (typeof radioSpecify !== 'undefined' && radioSpecify) radioSpecify.enabled = false;
        if (typeof inputSpecify !== 'undefined' && inputSpecify) inputSpecify.enabled = false;
        radioActive.value = true; // ensure default stays active
    }

    // デフォルト：作業アートボードのみ / Default: active artboard only
    radioActive.value = true;
    // 既定選択を反映してディム状態を最終更新 / Finalize dim state after defaults
    updateSpecifyEnabled();

    /* 再描画をデバウンス / Throttle redraw to reduce jank */
    var __redrawLast = 0;
    var __redrawInterval = 40; // ms（約25fpsで十分）
    function throttledRedraw() {
        var now = (new Date()).getTime();
        if (now - __redrawLast >= __redrawInterval) {
            try { app.redraw(); } catch (e) {}
            __redrawLast = now;
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
            // ピクセル単位のときは左上を整数座標にスナップ / If unit is px, snap top-left to integer
            if (String(unit).toLowerCase() === 'px') {
                var snappedL = Math.round(newRect[0]);
                var snappedT = Math.round(newRect[1]);
                // preserve width/height (wPt/hPt are already in pt units corresponding to px)
                newRect = [snappedL, snappedT, snappedL + wPt, snappedT - hPt];
            }
            app.activeDocument.artboards[abIdx].artboardRect = newRect;
        }

        var used = false;
        if (radioSpecify && radioSpecify.value === true) {
            var list = getSpecifiedIndices();
            if (list && list.length) {
                for (var li = 0; li < list.length; li++) {
                    _resizeOne(list[li]);
                }
                used = true;
            }
        }
        if (!used) {
            if (scopeAll) {
                for (var i = 0; i < abCount; i++) _resizeOne(i);
            } else {
                var idx = app.activeDocument.artboards.getActiveArtboardIndex();
                _resizeOne(idx);
            }
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
        var mode = "active";
        var specified = null;
        if (radioAll.value) mode = "all";
        else if (radioSpecify.value) { mode = "specify"; specified = getSpecifiedIndices(); }
        result = {
            targetScope: mode,
            specifiedIndexes: specified // 0-based indices if specify-mode
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
            __saveDlgLoc();
            if (typeof prev === 'function') return prev();
            dlg.close(1);
        };
    })(okBtn.onClick);

    if (typeof cancelBtn !== 'undefined' && cancelBtn) {
        cancelBtn.onClick = (function(prev) {
            return function() {
                __saveDlgLoc();
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
    if (app.documents.length === 0) {
        alert(LABELS.alertNoDoc[lang]);
        return;
    }
    var userInput = showMarginDialog(null, _detectUnitString());
    if (!userInput) return;
    // All resizing already applied in preview; nothing to do here.
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
        var isAlt = (kb.altKey || kb.optionKey);
        var isShift = kb.shiftKey && !isAlt; // Alt優先

        var next;
        if (isAlt) {
            // Option(Alt): ±0.1（最終的に整数丸めは下で実施）
            var stepSmall = 0.1;
            var deltaSmall = (event.keyName === "Up") ? stepSmall : -stepSmall;
            next = num + deltaSmall;
        } else if (isShift) {
            // Shift: 10の倍数へスナップ
            var n = Math.max(0, num);
            var up = (event.keyName === "Up");
            var isMultiple = (n % 10 === 0);
            if (up) {
                next = isMultiple ? (n + 10) : (Math.ceil(n / 10) * 10);
            } else {
                next = isMultiple ? (Math.max(0, n - 10)) : (Math.floor(n / 10) * 10);
            }
        } else {
            // 通常: ±1
            var delta = (event.keyName === "Up") ? 1 : -1;
            next = num + delta;
        }

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