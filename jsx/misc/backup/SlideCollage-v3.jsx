#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
slide-collage

### 更新日：
20260302

### 概要：
開いているドキュメント（B）のアートボードを元に、新規ドキュメントを作成し、指定したアートボードをグリッド配置してポートフォリオ用のサムネイル一覧を作成します。

・起動時に新規ドキュメント設定ダイアログ（サイズ／単位／カラーモード）を表示
  - プリセット（A4 / フルHD）＋セッション内プリセット保存に対応

・読み込み：
  - 指定：アートボード番号（例 1-10 / 1,3,5）を指定
  - 総数：最終的に配置する総数を指定
    例）指定 1-10 / 総数 30 → 1-10 を繰り返して30個配置

・アイテム：PDFの配置範囲（アート / トリミング / 仕上がり / 裁ち落とし）を選択、角丸（pt換算）を適用
  - 各アイテムは同サイズの矩形でクリップグループ化し、角丸（ライブエフェクト）はクリップグループに適用

・グリッド：方向（横 / 縦 / ランダム）、列数、間隔を設定
・偶数列：配分モード（偶数列＋1）で偶数列に+1スロットを追加、ずらし（偶数列の上下オフセット）を個別調整
・レイアウト：スケール（自動フィット結果に対する追加倍率）、回転（全体を回転し中心をアートボード中心に合わせる）、位置調整（横 / 縦）

・マスク：OK時にマージン内側でクリッピング
  - マスク角丸を設定可能（クリップグループに適用）

・アートボード：背景色（HEX指定）を追加可能（マスク対象外）

変更操作の多くはリアルタイムプレビュー（debounce 120ms）で確認できます（読み込みのアートボード番号変更は対象外）。
数値入力欄は ↑↓ / Shift / Option で増減できます。
ダイアログ位置はセッション中のみ記憶・復元します（Illustrator終了でリセット）。

オリジナルアイデア
Slide Collage - Portfolio Layout Generator -
https://slide-collage.vercel.app/

*/

// =========================================
// Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.3.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "Slide Collage for Illustrator",
        en: "Slide Collage for Illustrator"
    },

    // Panels
    panelLoad: { ja: "アートボードの読み込み", en: "Load Artboards" },
    panelItem: { ja: "アイテム", en: "Items" },
    panelGrid: { ja: "グリッド", en: "Grid" },
    panelLayout: { ja: "レイアウト", en: "Layout" },
    panelArtboard: { ja: "アートボードとマスク", en: "Artboard & Mask" },

    // Dialog X (New document)
    dialogNewDoc: { ja: "新規ドキュメント", en: "New Document" },
    panelUnitX: { ja: "単位", en: "Unit" },
    panelSizeX: { ja: "サイズ", en: "Size" },
    panelColorModeX: { ja: "カラーモード", en: "Color Mode" },
    presetCustom: { ja: "カスタム", en: "Custom" },
    presetA4Cmyk: { ja: "A4（210mm x 297mm）, CMYK", en: "A4 (210mm x 297mm), CMYK" },
    presetFhdRgb: { ja: "フルHD（1920 px × 1080 px）, RGB", en: "Full HD (1920 px × 1080 px), RGB" },
    btnSavePreset: { ja: "保存", en: "Save" },
    presetNameTitle: { ja: "プリセット名", en: "Preset name" },
    presetNamePanel: { ja: "名前", en: "Name" },
    presetDefaultName: { ja: "プリセット", en: "Preset" },

    // Load panel
    labelArtboards: { ja: "アートボード", en: "Artboards" },
    btnLoad: { ja: "読み込み", en: "Load" },

    // Item panel
    cropArt: { ja: "アート", en: "Art" },
    cropTrim: { ja: "トリミング", en: "Trim" },
    cropCrop: { ja: "仕上がり", en: "Crop" },
    cropBleed: { ja: "裁ち落とし", en: "Bleed" },
    round: { ja: "角丸", en: "Round" },

    // Grid panel
    direction: { ja: "方向：", en: "Flow: " },
    dirH: { ja: "横", en: "Horizontal" },
    dirV: { ja: "縦", en: "Vertical" },
    dirRandom: { ja: "ランダム", en: "Random" },
    dist: { ja: "配分：", en: "Distribution: " },
    evenPlus: { ja: "＋1スロット", en: "+1 slot" },
    cols: { ja: "列数", en: "Cols" },
    spacing: { ja: "間隔", en: "Gap" },
    shift: { ja: "ずらし", en: "Shift" },

    panelEven: { ja: "偶数列", en: "Even Columns" },

    // Layout panel
    scale: { ja: "スケール", en: "Scale" },
    rotate: { ja: "回転", en: "Rotate" },
    offsetX: { ja: "横方向の位置調整", en: "Offset X" },
    offsetY: { ja: "縦方向の位置調整", en: "Offset Y" },

    // Artboard panel
    margin: { ja: "マージン", en: "Margin" },
    mask: { ja: "マスク", en: "Mask" },
    maskRound: { ja: "マスク角丸", en: "Mask Round" },
    bg: { ja: "背景色", en: "Background" },

    // Buttons
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },

    // Buttons (extra)
    reset: { ja: "リセット", en: "Reset" },
    lightPreview: { ja: "軽量プレビュー", en: "Light preview" },

    // Zoom
    zoom: { ja: "画面ズーム", en: "Zoom" },
    lightMode: { ja: "軽量モード", en: "Light mode" },

    // File / Alerts
    fileDialogTitle: { ja: "配置するファイル（.ai または .pdf）を選択してください", en: "Select a file to place (.ai or .pdf)" },
    alertNeedDoc: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running." },
    alertPlaceError: { ja: "配置中にエラーが発生しました。", en: "An error occurred while placing items." },

    // Units / Defaults
    unitPercent: { ja: "%", en: "%" },
    unitDegree: { ja: "°", en: "°" },
    gapSpace: { ja: " ", en: " " },

    defaultPages: { ja: "1-20", en: "1-20" },
    defaultHex: { ja: "#000000", en: "#000000" },
    defaultCols: { ja: "5", en: "5" },
    defaultSpacing: { ja: "0", en: "0" },
    defaultShift: { ja: "0", en: "0" },
    defaultScale: { ja: "100", en: "100" },
    defaultRotate: { ja: "-12", en: "-12" },
    defaultRound: { ja: "10", en: "10" }
};

function L(key) {
    var o = LABELS[key];
    if (!o) return key;
    return o[lang] || o.en || o.ja || key;
}


// =========================================
// Unit utilities (rulerType)
// =========================================

// --- 外部定義：共通単位マップ ---
var __SC_unitMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/**
 * 単位コードと設定キーから適切な単位ラベルを返す（Q/H分岐含む）
 */
function __SC_getUnitLabel(code, prefKey) {
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return __SC_unitMap[code] || "pt";
}

/**
 * 単位コードから pt 換算係数を返す
 */
function __SC_getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0: return 72.0;                        // in
        case 1: return 72.0 / 25.4;                 // mm
        case 2: return 1.0;                         // pt
        case 3: return 12.0;                        // pica
        case 4: return 72.0 / 2.54;                 // cm
        case 5: return 72.0 / 25.4 * 0.25;          // Q or H
        case 6: return 1.0;                         // px
        case 7: return 72.0 * 12.0;                 // ft/in
        case 8: return 72.0 / 25.4 * 1000.0;        // m
        case 9: return 72.0 * 36.0;                 // yd
        case 10: return 72.0 * 12.0;                // ft
        default: return 1.0;
    }
}

/**
 * rulerType から {code,label,factor} を返す
 */

function __SC_getRulerUnitInfo() {
    var code = 2;
    try {
        code = app.preferences.getIntegerPreference("rulerType");
    } catch (e) {
        code = 2;
    }
    var label = __SC_getUnitLabel(code, "rulerType");
    var factor = __SC_getPtFactorFromUnitCode(code);
    return { code: code, label: label, factor: factor };
}

function __SC_round(n, digits) {
    var p = Math.pow(10, digits);
    return Math.round(n * p) / p;
}

function __SC_ptToUnit(pt, factor) {
    if (!(factor > 0)) factor = 1;
    return pt / factor;
}


function __SC_unitToPt(val, factor) {
    if (!(factor > 0)) factor = 1;
    return val * factor;
}

// =========================================
// PDF Import crop (trim / bleed / art)
// =========================================

// UI mode constants
var __SC_CROP_ART = 4;
var __SC_CROP_TRIM = 3;
var __SC_CROP_BLEED = 2;
var __SC_CROP_CROP = 1;
var __SC_CROP_MEDIA = 0;

/**
 * Set PDF import crop box preference.
 * Illustrator のバージョン差異を吸収するため、複数キーに試行します。
 * 期待する値（多くの環境で）: 0=Media, 1=Crop, 2=Bleed, 3=Trim, 4=Art
 */
function __SC_setPdfCropPreference(cropVal) {
    var keys = [
        "plugin/PDFImport/CropToBox",
        "plugin/PDFImport/CropTo",
        "plugin/PDFImport/CropBox",
        "plugin/PDFImport/CropToType"
    ];
    for (var i = 0; i < keys.length; i++) {
        try {
            app.preferences.setIntegerPreference(keys[i], cropVal);
        } catch (e) { }
    }
}


// Placing .ai in Illustrator uses the PDF import pipeline as well (AI is PDF-compatible),
// so we treat .ai as "PDF-like" for page/artboard selection.
function __SC_isPdfLikeFile(f) {
    try {
        if (!f) return false;
        var n = (f.name || "").toLowerCase();
        return (n.indexOf(".pdf") > -1) || (n.indexOf(".ai") > -1);
    } catch (e) {
        return false;
    }
}

// UI-only: crop box options are meaningful for PDF; keep disabled for AI.
function __SC_isPdfFile(f) {
    try {
        if (!f) return false;
        var n = (f.name || "").toLowerCase();
        return (n.indexOf(".pdf") > -1);
    } catch (e) {
        return false;
    }
}

// =========================================
// Auto-fit measurement cache (session only)
// key: fileFsName|cropMode|page
// value: { w:Number, h:Number }
// =========================================

function __SC_getAutoFitMeasureCache() {
    if (!$.global.__SC_autoFitMeasureCache) $.global.__SC_autoFitMeasureCache = {};
    return $.global.__SC_autoFitMeasureCache;
}

function __SC_getAutoFitMeasureKey(fileA, cropMode, pageNum) {
    var p = "";
    try { p = (fileA && fileA.fsName) ? String(fileA.fsName) : String(fileA); } catch (_) { p = String(fileA); }
    return p + "|" + String(cropMode) + "|" + String(pageNum);
}

function __SC_getAutoFitMeasure(fileA, cropMode, pageNum) {
    try {
        var box = __SC_getAutoFitMeasureCache();
        var k = __SC_getAutoFitMeasureKey(fileA, cropMode, pageNum);
        var v = box[k];
        if (v && v.w > 0 && v.h > 0) return v;
    } catch (_) { }
    return null;
}

function __SC_setAutoFitMeasure(fileA, cropMode, pageNum, w, h) {
    try {
        if (!(w > 0 && h > 0)) return;
        var box = __SC_getAutoFitMeasureCache();
        var k = __SC_getAutoFitMeasureKey(fileA, cropMode, pageNum);
        box[k] = { w: w, h: h };
    } catch (_) { }
}

// 入力された文字列（例："1-20", "1,3,5"）を数字の配列に変換する関数
function parsePageNumbers(inputStr) {
    var result = [];
    var parts = inputStr.split(',');

    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].replace(/^\s+|\s+$/g, '');

        if (part.indexOf('-') > -1) {
            var bounds = part.split('-');
            var start = parseInt(bounds[0], 10);
            var end = parseInt(bounds[1], 10);

            if (!isNaN(start) && !isNaN(end)) {
                var min = Math.min(start, end);
                var max = Math.max(start, end);
                for (var j = min; j <= max; j++) {
                    result.push(j);
                }
            }
        } else {
            var num = parseInt(part, 10);
            if (!isNaN(num)) {
                result.push(num);
            }
        }
    }
    return result;
}

// =========================================
// Dialog position (session only)
// =========================================

function __SC_getSessionDialogBoundsKey() {
    return "__SC_slideCollage_dialog_bounds";
}

function __SC_loadDialogBounds() {
    try {
        var k = __SC_getSessionDialogBoundsKey();
        return $.global[k];
    } catch (e) {
        return null;
    }
}

function __SC_saveDialogBounds(b) {
    try {
        if (!b) return;
        // store plain object (avoid ScriptUI objects)
        var l = b.left, t = b.top, r = b.right, bt = b.bottom;
        if (!(typeof l === "number" && isFinite(l))) return;
        if (!(typeof t === "number" && isFinite(t))) return;
        if (!(typeof r === "number" && isFinite(r))) return;
        if (!(typeof bt === "number" && isFinite(bt))) return;
        var o = { left: l, top: t, right: r, bottom: bt };
        $.global[__SC_getSessionDialogBoundsKey()] = o;
    } catch (e) { }
}


function __SC_applyDialogBounds(win) {
    try {
        var b = __SC_loadDialogBounds();
        if (!b) return;

        function isFiniteNum(n) {
            return (typeof n === "number") && isFinite(n);
        }

        var l = b.left, t = b.top, r = b.right, bt = b.bottom;
        if (!isFiniteNum(l) || !isFiniteNum(t) || !isFiniteNum(r) || !isFiniteNum(bt)) return;

        // Size sanity
        var w = r - l;
        var h = bt - t;
        if (!(w >= 200 && h >= 120)) return; // too small/invalid -> skip

        // Screen sanity: must intersect at least one screen bounds
        var intersects = false;
        try {
            if ($.screens && $.screens.length) {
                for (var i = 0; i < $.screens.length; i++) {
                    var sb = $.screens[i].bounds; // [L,T,R,B]
                    var sl = sb[0], st = sb[1], sr = sb[2], sbt = sb[3];
                    var il = Math.max(l, sl);
                    var it = Math.max(t, st);
                    var ir = Math.min(r, sr);
                    var ib = Math.min(bt, sbt);
                    if ((ir - il) > 40 && (ib - it) > 40) { // require some visible area
                        intersects = true;
                        break;
                    }
                }
            }
        } catch (eScr) {
            // If screens API is unavailable, fall back to applying.
            intersects = true;
        }

        if (!intersects) return;

        // Apply bounds
        win.bounds = [l, t, r, bt];
    } catch (e) { }
}

// =========================================
// Global updatePreview wrapper (for scheduleTask / global callbacks)
// =========================================
// `updatePreview` itself is defined inside main(). When something calls `updatePreview()`
// from the global scope, route it to the latest function stored in $.global.
function updatePreview() {
    try {
        if ($.global.__SC_updatePreview) {
            $.global.__SC_updatePreview();
        }
    } catch (e) { }
}

// =========================================
// TMK Zoom Module (collision-safe + Light mode)
// - Light mode: apply zoom only on slider release
// =========================================

function __TMKZoom_captureViewState(doc) {
    var st = { view: null, zoom: null, center: null };
    try {
        st.view = doc.activeView;
        st.zoom = st.view.zoom;
        st.center = st.view.centerPoint;
    } catch (_) { }
    return st;
}

function __TMKZoom_restoreViewState(doc, state) {
    if (!state) return;
    try {
        var v = state.view || doc.activeView;
        if (v && state.zoom != null) v.zoom = state.zoom;
        if (v && state.center != null) v.centerPoint = state.center;
    } catch (_) { }
}

function __TMKZoom_addControls(parent, doc, labelText, initialState, options) {
    options = options || {};
    var minZoom = (typeof options.min === "number") ? options.min : 0.1;
    var maxZoom = (typeof options.max === "number") ? options.max : 16;
    var sliderWidth = (typeof options.sliderWidth === "number") ? options.sliderWidth : 360;
    var doRedraw = (options.redraw !== false);

    // Light mode options
    var showLightMode = (options.lightMode !== false);            // default: show
    var lightModeLabel = options.lightModeLabel || "Light mode";
    var lightModeDefault = (options.lightModeDefault === true);   // default: false

    // UI group
    var g = parent.add("group");
    g.orientation = "row";
    g.alignChildren = ["center", "center"];
    g.alignment = "center";
    try { if (options.margins) g.margins = options.margins; } catch (_) { }

    var stLabel = g.add("statictext", undefined, String(labelText || "Zoom"));

    // Initial zoom
    var initZoom = 1;
    try {
        if (initialState && initialState.zoom != null) initZoom = Number(initialState.zoom);
        else initZoom = Number(doc.activeView.zoom);
    } catch (_) { }
    if (!initZoom || isNaN(initZoom)) initZoom = 1;

    var sld = g.add("slider", undefined, initZoom, minZoom, maxZoom);
    try { sld.preferredSize.width = sliderWidth; } catch (_) { }

    var chkLight = null;
    if (showLightMode) {
        chkLight = g.add("checkbox", undefined, String(lightModeLabel));
        chkLight.value = lightModeDefault;
    }

    function isLightMode() {
        return !!(chkLight && chkLight.value);
    }

    function applyZoom(z) {
        try {
            var v = (initialState && initialState.view) ? initialState.view : doc.activeView;
            if (!v) return;
            v.zoom = z;
            if (doRedraw) { try { app.redraw(); } catch (_) { } }
        } catch (_) { }
    }

    function syncFromView() {
        try {
            var v = (initialState && initialState.view) ? initialState.view : doc.activeView;
            if (!v) return;
            sld.value = v.zoom;
        } catch (_) { }
    }

    // Live drag (disabled in light mode)
    sld.onChanging = function () {
        if (isLightMode()) return; // ✅ lightweight: do nothing while dragging
        applyZoom(Number(sld.value));
    };

    // Always apply once on release
    sld.onChange = function () {
        applyZoom(Number(sld.value));
    };

    if (chkLight) {
        chkLight.onClick = function () {
            // Toggle feels consistent: apply current value immediately
            try { applyZoom(Number(sld.value)); } catch (_) { }
        };
    }

    return {
        group: g,
        label: stLabel,
        slider: sld,
        lightModeCheckbox: chkLight,
        applyZoom: applyZoom,
        syncFromView: syncFromView,
        restoreInitial: function () { __TMKZoom_restoreViewState(doc, initialState); }
    };
}

function main() {
    if (app.documents.length === 0) {
        alert(L("alertNeedDoc"));
        return;
    }

    // 1. 読み込み元（B）は「いま開いているドキュメント」
    var srcDocB = app.activeDocument;

    // B が未保存なら保存を促す（必須）
    try {
        if (!srcDocB.saved || !srcDocB.fullName) {
            var msg = (lang === "ja")
                ? "読み込み元ドキュメント（B）が未保存です。保存してから続行します。"
                : "The source document (B) is not saved. Please save it to continue.";
            alert(msg);

            var saveTo = File.saveDialog((lang === "ja") ? "読み込み元ドキュメントを保存" : "Save the source document");
            if (!saveTo) return;
            try { srcDocB.saveAs(saveTo); } catch (_) { return; }
        }
    } catch (_) { }

    // B のパス（配置に使うファイル）
    var fileA = null;
    try { fileA = srcDocB.fullName; } catch (_) { fileA = null; }
    if (!fileA) return;

    // B の総アートボード数（n）を取得
    var __srcArtboardCount = 0;
    try { __srcArtboardCount = (srcDocB.artboards) ? srcDocB.artboards.length : 0; } catch (_) { __srcArtboardCount = 0; }

    // セッションキャッシュへ（open せずに確定値を入れる）
    try {
        if (!$.global.__SC_sourcePageCountCache) $.global.__SC_sourcePageCountCache = {};
        var __k = (fileA && fileA.fsName) ? String(fileA.fsName) : String(fileA);
        if (__srcArtboardCount > 0) $.global.__SC_sourcePageCountCache[__k] = __srcArtboardCount;
    } catch (_) { }

    // 現在の定規単位（rulerType）※ダイアログXで使う
    var rulerUnit = __SC_getRulerUnitInfo();

    // B のアクティブアートボードサイズをデフォルトにする
    var __wPt = 1920, __hPt = 1080;
    try {
        var ab0 = srcDocB.artboards[srcDocB.artboards.getActiveArtboardIndex()];
        var r0 = ab0.artboardRect; // [L,T,R,B]
        __wPt = Math.abs(r0[2] - r0[0]);
        __hPt = Math.abs(r0[1] - r0[3]);
        if (!(__wPt > 0)) __wPt = 1920;
        if (!(__hPt > 0)) __hPt = 1080;
    } catch (_) { }

    function __SC_showSetupDialogX(defaultWPt, defaultHPt, defaultColorSpace) {
        var dlgX = new Window('dialog', L('dialogNewDoc'));
        dlgX.alignChildren = 'fill';

        // Unit selection for Size inputs
        var __unitLabel = 'pt';
        var __unitFactor = 1;
        function __SC_setUnit(label) {
            __unitLabel = String(label || 'pt');
            if (__unitLabel === 'mm') __unitFactor = (72 / 25.4);
            else if (__unitLabel === 'px') __unitFactor = 1; // 1px = 1pt in this script
            else __unitFactor = 1; // pt

            // update unit labels
            try { stUnitW.text = __unitLabel; } catch (_) { }
            try { stUnitH.text = __unitLabel; } catch (_) { }

            // re-normalize current inputs (treat current text as old unit? -> keep numeric as-is)
            // We keep the numeric values and only change meaning. (Predictable for users.)
        }

        function __SC_unitToPtX(val) {
            var v = parseFloat(String(val));
            if (isNaN(v)) return NaN;
            return v * __unitFactor;
        }
        function __SC_ptToUnitX(pt) {
            return pt / __unitFactor;
        }

        // Preset dropdown + save button (spans both columns)
        var groupPreset = dlgX.add('group');
        groupPreset.orientation = 'row';
        groupPreset.alignChildren = ['center', 'center'];

        var ddPreset = groupPreset.add('dropdownlist', undefined, [
            L('presetCustom'),
            L('presetA4Cmyk'),
            L('presetFhdRgb')
        ]);
        ddPreset.selection = 2;

        var btnSavePreset = groupPreset.add('button', undefined, L('btnSavePreset'));
        try { btnSavePreset.preferredSize = [54, 22]; } catch (_) { }

        // Session preset store
        if (!$.global.__SC_docPresets) $.global.__SC_docPresets = [];

        btnSavePreset.onClick = function () {
            try {
                // Ask preset name
                var presetName = "";
                try {
                    var d = new Window('dialog', L('presetNameTitle'));
                    d.alignChildren = 'fill';
                    var p = d.add('panel', undefined, L('presetNamePanel'));
                    p.orientation = 'column';
                    p.alignChildren = ['left', 'top'];
                    p.margins = [15, 20, 15, 10];
                    var et = p.add('edittext', undefined, '');
                    et.characters = 20;
                    et.active = true;

                    var gb = d.add('group');
                    gb.orientation = 'row';
                    gb.alignChildren = ['right', 'center'];
                    var bc = gb.add('button', undefined, L('cancel'), { name: 'cancel' });
                    var bo = gb.add('button', undefined, L('ok'), { name: 'ok' });

                    bo.onClick = function () {
                        presetName = String(et.text || '').replace(/^\s+|\s+$/g, '');
                        d.close(1);
                    };
                    bc.onClick = function () { d.close(0); };

                    var rr = d.show();
                    if (rr !== 1) return; // cancelled
                    if (!presetName) {
                        // empty -> fallback
                        presetName = L('presetDefaultName');
                    }
                } catch (_) {
                    // Fallback without dialog
                    presetName = L('presetDefaultName');
                }

                var preset = {
                    name: presetName,
                    unit: (rbUmm.value ? 'mm' : rbUpx.value ? 'px' : 'pt'),
                    w: editW.text,
                    h: editH.text,
                    color: (rbCMYK.value ? 'CMYK' : 'RGB')
                };

                $.global.__SC_docPresets.push(preset);

                var label = preset.name + ': ' + preset.w + '×' + preset.h + ' ' + preset.unit + ' / ' + preset.color;
                ddPreset.add('item', label);
                ddPreset.selection = ddPreset.items.length - 1;
            } catch (_) { }
        };

        // 3-column layout
        var row2 = dlgX.add('group');
        row2.orientation = 'row';
        row2.alignChildren = ['left', 'top'];

        // ---- Unit panel ----
        var pnlUnit = row2.add('panel', undefined, L('panelUnitX'));
        pnlUnit.orientation = 'column';
        pnlUnit.alignChildren = ['left', 'top'];
        pnlUnit.margins = [15, 20, 15, 10];

        // Unit radios (vertical)
        var gUnit = pnlUnit.add('group');
        gUnit.orientation = 'column';
        gUnit.alignChildren = ['left', 'top'];
        var rbUmm = gUnit.add('radiobutton', undefined, 'mm');
        var rbUpx = gUnit.add('radiobutton', undefined, 'px');
        var rbUpt = gUnit.add('radiobutton', undefined, 'pt');

        // default unit follows current ruler if possible
        try {
            if (rulerUnit.label === 'mm') rbUmm.value = true;
            else if (rulerUnit.label === 'px') rbUpx.value = true;
            else rbUpt.value = true;
        } catch (_) {
            rbUpt.value = true;
        }

        // ---- Size panel ----
        var pnlSize = row2.add('panel', undefined, L('panelSizeX'));
        pnlSize.orientation = 'column';
        pnlSize.alignChildren = ['left', 'top'];
        pnlSize.margins = [15, 20, 15, 10];

        // Helper for setting W/H edit fields by pt
        function setWHByPt(wPt, hPt) {
            editW.text = String(__SC_round(__SC_ptToUnitX(wPt), 2));
            editH.text = String(__SC_round(__SC_ptToUnitX(hPt), 2));
        }

        // Width group
        var gW = pnlSize.add('group');
        gW.orientation = 'row';
        gW.alignChildren = ['left', 'center'];
        var stW = gW.add('statictext', undefined, (lang === 'ja') ? '幅' : 'Width');
        stW.justify = 'right';
        stW.preferredSize.width = 44;
        var editW = gW.add('edittext', undefined, String(__SC_round(__SC_ptToUnitX(defaultWPt), 2)));
        editW.characters = 5;
        var stUnitW = gW.add('statictext', undefined, __unitLabel);
        stUnitW.preferredSize.width = 26;

        // Height group
        var gH = pnlSize.add('group');
        gH.orientation = 'row';
        gH.alignChildren = ['left', 'center'];
        var stH = gH.add('statictext', undefined, (lang === 'ja') ? '高さ' : 'Height');
        stH.justify = 'right';
        stH.preferredSize.width = 44;
        var editH = gH.add('edittext', undefined, String(__SC_round(__SC_ptToUnitX(defaultHPt), 2)));
        editH.characters = 5;
        var stUnitH = gH.add('statictext', undefined, __unitLabel);
        stUnitH.preferredSize.width = 26;

        // Initialize unit based on default radio
        if (rbUmm.value) __SC_setUnit('mm');
        else if (rbUpx.value) __SC_setUnit('px');
        else __SC_setUnit('pt');

        function __SC_applyUnitFromRadios() {
            if (rbUmm.value) __SC_setUnit('mm');
            else if (rbUpx.value) __SC_setUnit('px');
            else __SC_setUnit('pt');
        }

        rbUmm.onClick = function () { __SC_applyUnitFromRadios(); };
        rbUpx.onClick = function () { __SC_applyUnitFromRadios(); };
        rbUpt.onClick = function () { __SC_applyUnitFromRadios(); };

        // Preset dropdown logic
        ddPreset.onChange = function () {
            // Session presets
            if ($.global.__SC_docPresets && ddPreset.selection.index >= 3) {
                var idx2 = ddPreset.selection.index - 3;
                var p = $.global.__SC_docPresets[idx2];
                if (p) {
                    try {
                        rbUmm.value = (p.unit === 'mm');
                        rbUpx.value = (p.unit === 'px');
                        rbUpt.value = (p.unit === 'pt');
                        __SC_applyUnitFromRadios();

                        editW.text = p.w;
                        editH.text = p.h;

                        rbCMYK.value = (p.color === 'CMYK');
                        rbRGB.value = (p.color === 'RGB');
                    } catch (_) { }
                    return;
                }
            }
            if (!ddPreset.selection) return;
            var idx = ddPreset.selection.index;

            if (idx === 1) {
                // A4: 210mm x 297mm  → unit = mm, color = CMYK
                try {
                    rbUmm.value = true;
                    rbUpx.value = false;
                    rbUpt.value = false;
                    __SC_applyUnitFromRadios();
                } catch (_) { }

                var wPt = 210 * (72 / 25.4);
                var hPt = 297 * (72 / 25.4);
                setWHByPt(wPt, hPt);

                try {
                    rbCMYK.value = true;
                    rbRGB.value = false;
                } catch (_) { }

            } else if (idx === 2) {
                // FullHD: 1920px x 1080px → unit = px, color = RGB
                try {
                    rbUmm.value = false;
                    rbUpx.value = true;
                    rbUpt.value = false;
                    __SC_applyUnitFromRadios();
                } catch (_) { }

                setWHByPt(1920, 1080);

                try {
                    rbRGB.value = true;
                    rbCMYK.value = false;
                } catch (_) { }

            } else {
                // カスタム: 単位・サイズは変更しない
            }
        };

        // ---- Color mode panel ----
        var pnlColor = row2.add('panel', undefined, L('panelColorModeX'));
        pnlColor.orientation = 'column';
        pnlColor.alignChildren = ['left', 'top'];
        pnlColor.margins = [15, 20, 15, 10];

        // Color mode radios (vertical)
        var gC = pnlColor.add('group');
        gC.orientation = 'column';
        gC.alignChildren = ['left', 'top'];
        var rbRGB = gC.add('radiobutton', undefined, 'RGB');
        var rbCMYK = gC.add('radiobutton', undefined, 'CMYK');

        try {
            if (defaultColorSpace === DocumentColorSpace.CMYK) rbCMYK.value = true;
            else rbRGB.value = true;
        } catch (_) {
            rbRGB.value = true;
        }

        // Buttons
        var gBtn = dlgX.add('group');
        gBtn.orientation = 'row';
        gBtn.alignChildren = ['center', 'center'];
        var btnCancelX = gBtn.add('button', undefined, L('cancel'), { name: 'cancel' });
        var btnOkX = gBtn.add('button', undefined, L('ok'), { name: 'ok' });

        function parseNum(s, fallback) {
            var v = parseFloat(String(s));
            if (isNaN(v) || !(v > 0)) return fallback;
            return v;
        }

        btnOkX.onClick = function () {
            var wU = parseNum(editW.text, __SC_ptToUnitX(defaultWPt));
            var hU = parseNum(editH.text, __SC_ptToUnitX(defaultHPt));
            var wPt = __SC_unitToPtX(wU);
            var hPt = __SC_unitToPtX(hU);
            if (!(wPt > 1)) wPt = defaultWPt;
            if (!(hPt > 1)) hPt = defaultHPt;

            var cs = DocumentColorSpace.RGB;
            try { cs = rbCMYK.value ? DocumentColorSpace.CMYK : DocumentColorSpace.RGB; } catch (_) { cs = DocumentColorSpace.RGB; }

            dlgX.__result = { wPt: wPt, hPt: hPt, colorSpace: cs };
            dlgX.close(1);
        };

        btnCancelX.onClick = function () { dlgX.close(0); };

        var r = dlgX.show();
        if (r === 1 && dlgX.__result) return dlgX.__result;
        return null;
    }

    var __defaultCS = DocumentColorSpace.RGB;
    try { __defaultCS = srcDocB.documentColorSpace; } catch (_) { __defaultCS = DocumentColorSpace.RGB; }

    // --- ダイアログXを表示 ---
    var setupX = __SC_showSetupDialogX(__wPt, __hPt, __defaultCS);
    if (!setupX) return; // user cancelled

    // 2b. 新規ドキュメント（A）を作成
    var doc = null;
    try {
        doc = app.documents.add(setupX.colorSpace, setupX.wPt, setupX.hPt);
    } catch (_) {
        try { doc = app.documents.add(DocumentColorSpace.RGB, setupX.wPt, setupX.hPt); } catch (e2) { doc = null; }
    }
    if (!doc) return;

    // 既定値は pt ベースで保持し、表示時に定規単位へ変換
    var DEFAULT_SPACING_PT = 20;
    var DEFAULT_MARGIN_PT = 20;

    // ラベル幅（列数/間隔/列ずらし/スケール/回転 を揃える）
    var LABEL_W = 60;
    // 位置調整ラベル幅（横方向の位置調整/縦方向の位置調整）
    var OFFSET_LABEL_W = 140;
    // 単位・補助ラベル幅（空白/pt/%/° を揃える）
    var UNIT_W = 24;

    // スライダー幅（列数/列ずらし/スケール/回転/横/縦 を揃える）
    var SLIDER_W = 140;

    // 自動フィット（UIは非表示。ロジックは維持して常にON）
    var AUTO_FIT_ENABLED = true;

    // 3. ダイアログ（Y）を開く段階で B を閉じる
    try {
        srcDocB.close(SaveOptions.DONOTSAVECHANGES);
    } catch (_) { }

    // -----------------------------------------
    // Source file page/artboard count (session cache)
    // - For .ai: artboards length
    // - For .pdf: pages are represented as artboards when opened
    // -----------------------------------------
    function __SC_getSourceDocKey(fileObj) {
        try { return (fileObj && fileObj.fsName) ? String(fileObj.fsName) : String(fileObj); }
        catch (_) { return String(fileObj); }
    }

    function __SC_getSourcePageCount(fileObj) {
        try {
            if (!fileObj) return 0;

            if (!$.global.__SC_sourcePageCountCache) $.global.__SC_sourcePageCountCache = {};
            var cache = $.global.__SC_sourcePageCountCache;
            var key = __SC_getSourceDocKey(fileObj);
            if (cache[key] && cache[key] > 0) return cache[key];

            var tempDoc = null;
            var n = 0;
            try {
                // Open the source file to read artboard/page count, then close without saving.
                tempDoc = app.open(fileObj);
                try { n = (tempDoc && tempDoc.artboards) ? tempDoc.artboards.length : 0; } catch (_) { n = 0; }
            } catch (_) {
                n = 0;
            } finally {
                try { if (tempDoc) tempDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (_) { }
            }

            if (n > 0) cache[key] = n;
            return n;
        } catch (_) { }
        return 0;
    }

    // Map requested page numbers to existing page range by repeating (e.g., 1-10 with 6 pages -> 1234561234)
    function __SC_repeatPagesWithinCount(pages, maxCount) {
        if (!pages || pages.length === 0) return pages;
        if (!(maxCount > 0)) return pages;
        var out = [];
        for (var i = 0; i < pages.length; i++) {
            var p = parseInt(pages[i], 10);
            if (isNaN(p) || p < 1) p = 1;
            var m = ((p - 1) % maxCount) + 1;
            out.push(m);
        }
        return out;
    }

    // Expand base sequence to desired total length by cycling.
    // Example: base=[1..10], total=13 -> 1..10,1,2,3
    function __SC_expandPagesToTotal(basePages, totalCount) {
        if (!basePages || basePages.length === 0) return basePages;
        var t = parseInt(totalCount, 10);
        if (isNaN(t) || t <= 0) return basePages;
        var out = [];
        for (var i = 0; i < t; i++) {
            out.push(basePages[i % basePages.length]);
        }
        return out;
    }

    // =========================================
    // Preview/cache declarations (moved to main() early)
    // =========================================

    // プレビュー用に配置したアイテムを保持する配列
    var previewItems = [];

    // Preview cache: placed items are created on [読み込み] and reused until next load
    var __previewCache = {
        items: [],
        baseW: [],
        baseH: [],
        pagesKey: "",
        cropMode: null,
        bgItem: null,
        group: null,
        currentRot: 0,
        randOrder: null,
        previewWrapped: false,
        previewRoundRadiusPt: 0
    };

    function __SC_removeItemSafe(it) {
        try { if (it) it.remove(); } catch (_) { }
    }

    function __SC_clearPreviewCache() {
        // remove group (removes its children automatically)
        if (__previewCache.group) {
            __SC_removeItemSafe(__previewCache.group);
        }
        __previewCache.group = null;
        __previewCache.currentRot = 0;

        // clear arrays (items already removed with group)
        __previewCache.items = [];
        __previewCache.baseW = [];
        __previewCache.baseH = [];

        // remove bg
        __SC_removeItemSafe(__previewCache.bgItem);
        __previewCache.bgItem = null;

        __previewCache.pagesKey = "";
        __previewCache.cropMode = null;
        __previewCache.previewWrapped = false;
        __previewCache.previewRoundRadiusPt = 0;
        __previewCache.randOrder = null;
    }
    function __SC_shuffleIndexArray(n) {
        var a = [];
        for (var i = 0; i < n; i++) a.push(i);
        for (var j = a.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var t = a[j];
            a[j] = a[k];
            a[k] = t;
        }
        return a;
    }

    function __SC_ensureRandomOrder() {
        var n = (__previewCache.items) ? __previewCache.items.length : 0;
        if (n <= 1) { __previewCache.randOrder = null; return; }
        if (!__previewCache.randOrder || __previewCache.randOrder.length !== n) {
            __previewCache.randOrder = __SC_shuffleIndexArray(n);
        }
    }

    function __SC_clearRandomOrder() {
        __previewCache.randOrder = null;
    }

    // 2. ダイアログボックスの作成
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.alignChildren = "fill";

    // セッション中のダイアログ位置を復元 / Restore dialog position in this session
    __SC_applyDialogBounds(win);

    // If no valid stored bounds, center as a safe fallback
    try {
        if (!__SC_loadDialogBounds()) {
            win.center();
        }
    } catch (eCenter) { }

    // 移動時に位置を記憶 / Remember position while moving
    win.onMove = function () { __SC_saveDialogBounds(win.bounds); };
    win.onMoving = function () { __SC_saveDialogBounds(win.bounds); };

    // === 2カラム構成 ===
    var mainRow = win.add("group");
    mainRow.orientation = "row";
    mainRow.alignChildren = ["fill", "fill"];

    // 左カラム
    var leftCol = mainRow.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "fill";

    // 右カラム
    var rightCol = mainRow.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = "fill";

    // --- ページ設定パネル ---
    var panelPage = leftCol.add("panel", undefined, L("panelLoad"));
    panelPage.alignChildren = "left";
    panelPage.margins = [15, 20, 15, 10];

    var groupPage = panelPage.add("group");
    groupPage.orientation = "column";
    groupPage.alignChildren = ["left", "top"];

    // Row 1: 指定
    var rowSpec = groupPage.add("group");
    rowSpec.orientation = "row";
    rowSpec.alignChildren = ["left", "center"];

    var stSpec = rowSpec.add("statictext", undefined, "指定");
    stSpec.preferredSize.width = 32;

    var editPages = rowSpec.add(
        "edittext",
        undefined,
        (__srcArtboardCount > 0)
            ? ("1-" + __srcArtboardCount)
            : L("defaultPages")
    );
    editPages.characters = 8;
    editPages.active = true;

    function __SC_getTargetPagesFromUI() {
        var base = parsePageNumbers(editPages.text);
        if (!base || base.length === 0) base = [1];

        // Clamp each requested number into existing source range (repeat within source count)
        var srcCount = 0;
        try { srcCount = __SC_getSourcePageCount(fileA); } catch (_) { srcCount = 0; }
        if (srcCount > 0) base = __SC_repeatPagesWithinCount(base, srcCount);

        // Desired total placements
        var total = parseInt(String(editTotal.text || '').replace(/^\s+|\s+$/g, ''), 10);
        if (isNaN(total) || total <= 0) {
            total = (srcCount > 0) ? srcCount : base.length;
        }

        return __SC_expandPagesToTotal(base, total);
    }

    // Row 2: 総数（読み込み元のアートボード数）
    var rowTotal = groupPage.add("group");
    rowTotal.orientation = "row";
    rowTotal.alignChildren = ["left", "center"];

    var stTotal = rowTotal.add("statictext", undefined, "総数");
    stTotal.preferredSize.width = 32;

    var editTotal = rowTotal.add(
        "edittext",
        undefined,
        (__srcArtboardCount > 0) ? String(__srcArtboardCount) : ""
    );
    editTotal.characters = 8;
    editTotal.enabled = true; // 編集可能（ディムにしない）

    // Row 3: 読み込みボタン
    var rowBtn = groupPage.add("group");
    rowBtn.orientation = "row";
    rowBtn.alignChildren = ["right", "right"];
    // rowBtn.alignment = ["fill", "top"];

    var btnPreview = rowBtn.add("button", undefined, L("btnLoad"));
    btnPreview.preferredSize = [80, 22];

    // --- トリミングパネル（配置範囲：PDF のトリミング設定） ---
    var panelCrop = leftCol.add("panel", undefined, L("panelItem"));
    panelCrop.alignChildren = "left";
    panelCrop.margins = [15, 20, 15, 10];

    // panelCrop.add("statictext", undefined, "トリミング");
    var ddCrop = panelCrop.add("dropdownlist", undefined, [L("cropArt"), L("cropTrim"), L("cropCrop"), L("cropBleed")]);
    ddCrop.minimumSize.width = 160;


    /* 角丸 / Round corners */
    var groupRound = panelCrop.add("group");
    groupRound.orientation = "row";
    groupRound.alignChildren = ["left", "center"];

    var cbRound = groupRound.add("checkbox", undefined, L("round"));
    cbRound.value = false;

    var editRound = groupRound.add("edittext", undefined, L("defaultRound"));
    editRound.characters = 3;

    groupRound.add("statictext", undefined, rulerUnit.label);

    // 初期状態
    editRound.enabled = cbRound.value;

    cbRound.onClick = function () {
        editRound.enabled = cbRound.value;
    };

    // デフォルト：仕上がり
    ddCrop.selection = 2;

    // PDF のときのみ有効（AIはPDF-likeだが、crop boxはPDFのみ有効）
    ddCrop.enabled = __SC_isPdfFile(fileA);

    function __SC_getCropModeFromUI() {
        var idx = (ddCrop.selection) ? ddCrop.selection.index : 2;
        // 0:アート / 1:トリミング / 2:仕上がり / 3:裁ち落とし
        if (idx === 0) return __SC_CROP_ART;
        if (idx === 1) return __SC_CROP_TRIM;
        if (idx === 3) return __SC_CROP_BLEED;
        // 仕上がりは CropBox を想定（環境差があるため失敗時は無視される）
        return __SC_CROP_CROP;
    }

    // -----------------------------------------
    // Import page/artboard selection
    // - PDF: uses PDFImport/PageNumber
    // - AI: uses IllustratorImport/ArtboardNumber (+ PlaceArtboards)
    // -----------------------------------------
    function __SC_setImportPageNumber(fileObj, pageNum) {
        var n = parseInt(pageNum, 10);
        if (isNaN(n) || n < 1) n = 1;
        try {
            // Use PDFImport for both PDF and AI (AI is handled by the PDF import pipeline when placing)
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", n);
        } catch (_) { }
    }

    function __SC_resetImportPageNumber(fileObj) {
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1);
        } catch (_) { }
    }

    /* グリッド / Grid */
    var panelLayout = rightCol.add("panel", undefined, L("panelGrid"));
    panelLayout.alignChildren = "left";
    panelLayout.margins = [15, 20, 15, 10];

    // 方向（配置順）
    var groupDir = panelLayout.add("group");
    groupDir.orientation = "row";
    groupDir.alignChildren = ["left", "center"];

    groupDir.add("statictext", undefined, L("direction"));

    var rbDirH = groupDir.add("radiobutton", undefined, L("dirH"));
    var rbDirV = groupDir.add("radiobutton", undefined, L("dirV"));
    var rbDirR = groupDir.add("radiobutton", undefined, L("dirRandom"));

    // デフォルト：縦方向
    rbDirV.value = true;

    // 0=横方向, 1=縦方向, 2=ランダム
    function __SC_getFlowMode() {
        if (rbDirR.value) return 2;
        if (rbDirV.value) return 1;
        return 0;
    }

    // 列数
    var groupCols = panelLayout.add("group");
    groupCols.orientation = "row";
    groupCols.alignChildren = ["left", "center"];

    var stCols = groupCols.add("statictext", undefined, L("cols"));
    stCols.preferredSize.width = LABEL_W;
    var editCols = groupCols.add("edittext", undefined, L("defaultCols"));
    editCols.characters = 4;

    var gapCols = groupCols.add("statictext", undefined, L("gapSpace"));
    gapCols.preferredSize.width = UNIT_W;

    var spacerCols = groupCols.add("group");
    spacerCols.alignment = ["fill", "fill"];
    spacerCols.minimumSize.width = 0;

    var sldCols = groupCols.add("slider", undefined, 5, 0, 10);
    sldCols.preferredSize.width = SLIDER_W;

    function __SC_syncColsFromEdit() {
        var v = parseInt(editCols.text, 10);
        if (isNaN(v) || v < 0) v = 0;
        if (v > 10) v = 10;
        editCols.text = String(v);
        sldCols.value = v;
    }

    function __SC_syncColsFromSlider() {
        var v = Math.round(sldCols.value);
        if (v < 0) v = 0;
        if (v > 10) v = 10;
        editCols.text = String(v);
    }

    __SC_syncColsFromEdit();
    editCols.onChanging = function () { __SC_syncColsFromEdit(); };
    sldCols.onChanging = function () { __SC_syncColsFromSlider(); };

    // リアルタイムプレビュー（確定時）

    var groupSpacing = panelLayout.add("group");
    var stSpacing = groupSpacing.add("statictext", undefined, L("spacing"));
    stSpacing.preferredSize.width = LABEL_W;
    var editSpacing = groupSpacing.add("edittext", undefined, String(__SC_round(__SC_ptToUnit(DEFAULT_SPACING_PT, rulerUnit.factor), 2)));
    editSpacing.characters = 4;
    var stSpacingUnit = groupSpacing.add("statictext", undefined, rulerUnit.label);
    stSpacingUnit.preferredSize.width = UNIT_W;

    // 間隔スライダー（右側）
    var spacerSpacing = groupSpacing.add("group");
    spacerSpacing.alignment = ["fill", "fill"];
    spacerSpacing.minimumSize.width = 0;

    var sldSpacing = groupSpacing.add("slider", undefined, 0, 0, 100);
    sldSpacing.preferredSize.width = SLIDER_W;

    // 間隔

    function __SC_syncSpacingFromEdit() {
        var v = parseFloat(editSpacing.text);
        if (isNaN(v) || v < 0) v = 0;
        if (v > 100) v = 100;
        v = __SC_round(v, 2);
        editSpacing.text = String(v);
        try { sldSpacing.value = v; } catch (e) { }
    }

    function __SC_syncSpacingFromSlider() {
        var v = sldSpacing.value;
        if (isNaN(v) || v < 0) v = 0;
        if (v > 100) v = 100;
        v = __SC_round(v, 2);
        editSpacing.text = String(v);
    }

    // 初期同期
    __SC_syncSpacingFromEdit();

    // 連動
    editSpacing.onChanging = function () { __SC_syncSpacingFromEdit(); };
    sldSpacing.onChanging = function () { __SC_syncSpacingFromSlider(); };

    // --- 偶数列パネル（配分/ずらし） ---
    var panelEven = rightCol.add("panel", undefined, L("panelEven"));
    panelEven.alignChildren = "left";
    panelEven.margins = [15, 20, 15, 10];

    // 配分モード
    var groupDist = panelEven.add("group");
    groupDist.orientation = "row";
    groupDist.alignChildren = ["left", "center"];

    // var stDist = groupDist.add("statictext", undefined, L("dist"));
    // stDist.preferredSize.width = LABEL_W;

    var cbEvenPlus = groupDist.add("checkbox", undefined, L("evenPlus"));

    cbEvenPlus.helpTip = (lang === "ja")
        ? "偶数列にスロットを1つ追加します。空きが出ることがあります。"
        : "Adds one extra slot to even columns. Empty spaces may appear.";

    cbEvenPlus.value = false;

    // ずらし（Yオフセット）
    var groupColShift = panelEven.add("group");
    groupColShift.orientation = "row";
    groupColShift.alignChildren = ["left", "center"];

    var cbColShift = groupColShift.add("checkbox", undefined, L("shift"));

    cbColShift.helpTip = (lang === "ja")
        ? "偶数列だけのYオフセット（上下方向のずらし）"
        : "Y-offset applied to even columns only.";

    cbColShift.preferredSize.width = LABEL_W;
    cbColShift.value = true;

    var editColShift = groupColShift.add("edittext", undefined, L("defaultShift"));
    editColShift.characters = 4;

    var stShiftUnit = groupColShift.add("statictext", undefined, rulerUnit.label);
    stShiftUnit.preferredSize.width = UNIT_W;

    var spacerShift = groupColShift.add("group");
    spacerShift.alignment = ["fill", "fill"];
    spacerShift.minimumSize.width = 0;

    var sldColShift = groupColShift.add("slider", undefined, 0, -200, 200);
    sldColShift.preferredSize.width = SLIDER_W;

    function __SC_syncShiftFromEdit() {
        var v = parseFloat(editColShift.text);
        if (isNaN(v)) v = 0;
        if (v < -200) v = -200;
        if (v > 200) v = 200;
        editColShift.text = String(v);
        sldColShift.value = v;
    }

    function __SC_syncShiftFromSlider() {
        var v = sldColShift.value;
        v = Math.round(v);
        if (v < -200) v = -200;
        if (v > 200) v = 200;
        editColShift.text = String(v);
    }

    __SC_syncShiftFromEdit();

    editColShift.onChanging = function () { __SC_syncShiftFromEdit(); };
    sldColShift.onChanging = function () { __SC_syncShiftFromSlider(); };

    // 初期状態
    editColShift.enabled = cbColShift.value;
    sldColShift.enabled = cbColShift.value;

    cbColShift.onClick = function () {
        editColShift.enabled = cbColShift.value;
        sldColShift.enabled = cbColShift.value;
    };

    /* レイアウト / Layout */
    var panelLayoutRight = rightCol.add("panel", undefined, L("panelLayout"));
    panelLayoutRight.alignChildren = "left";
    panelLayoutRight.margins = [15, 20, 15, 10];

    // --- アートボードとマスクパネル（左カラム） ---
    var panelArtboard = leftCol.add("panel", undefined, L("panelArtboard"));
    panelArtboard.alignChildren = "left";
    panelArtboard.margins = [15, 20, 15, 10];

    // 背景色（アートボードと同じ大きさの図形を作成）
    var groupBg = panelArtboard.add("group");
    groupBg.orientation = "row";
    groupBg.alignChildren = ["left", "center"];

    var cbBg = groupBg.add("checkbox", undefined, L("bg"));
    cbBg.value = true;

    // カラーチップ（表示用）
    var colorSwatch = groupBg.add("panel");
    colorSwatch.preferredSize = [24, 24];

    // HEX 入力
    var editBgHex = groupBg.add("edittext", undefined, L("defaultHex"));
    editBgHex.characters = 7;

    function __SC_makeRGBColor(r, g, b) {
        var c = new RGBColor();
        c.red = r; c.green = g; c.blue = b;
        return c;
    }

    function __SC_parseHexToRGBColor(hexStr) {
        var s = String(hexStr || "").replace(/^\s+|\s+$/g, "");
        if (s.charAt(0) !== "#") s = "#" + s;
        if (!/^#[0-9a-fA-F]{6}$/.test(s)) return null;
        var r = parseInt(s.substr(1, 2), 16);
        var g = parseInt(s.substr(3, 2), 16);
        var b = parseInt(s.substr(5, 2), 16);
        if (isNaN(r) || isNaN(g) || isNaN(b)) return null;
        return __SC_makeRGBColor(r, g, b);
    }

    function __SC_setSwatchRGB(r, g, b) {
        try {
            var gg = colorSwatch.graphics;
            gg.backgroundColor = gg.newBrush(gg.BrushType.SOLID_COLOR, [r / 255, g / 255, b / 255, 1]);
            gg.foregroundColor = gg.newPen(gg.PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
        } catch (e) { }
    }

    function __SC_getBgRGBColorOrDefault() {
        var c = __SC_parseHexToRGBColor(editBgHex.text);
        if (c) return c;
        return __SC_makeRGBColor(0, 0, 0);
    }

    function __SC_updateBgControls() {
        var en = cbBg.value;
        try { colorSwatch.enabled = en; } catch (e) { }
        try { editBgHex.enabled = en; } catch (e2) { }

        // swatch preview
        var c = __SC_getBgRGBColorOrDefault();
        __SC_setSwatchRGB(c.red, c.green, c.blue);
    }

    __SC_updateBgControls();

    cbBg.onClick = function () {
        __SC_updateBgControls();
    };

    // HEX は入力確定時だけでOK（重くしない）
    editBgHex.onChange = function () {
        __SC_updateBgControls();
    };

    // --- カラーピッカー（スウォッチクリック） ---
    function __SC_openBgColorPicker() {
        var init = __SC_getBgRGBColorOrDefault();
        var c = new RGBColor();
        c.red = init.red;
        c.green = init.green;
        c.blue = init.blue;

        if (app.showColorPicker(c)) {
            function toHex(n) {
                var s = n.toString(16);
                return (s.length === 1) ? "0" + s : s;
            }
            var hex = "#" + toHex(c.red) + toHex(c.green) + toHex(c.blue);
            editBgHex.text = hex;
            __SC_updateBgControls();
            try { __SC_updateBgOnly(); } catch (e) { }
        }
    }

    colorSwatch.onClick = function () {
        __SC_openBgColorPicker();
    };

    try {
        colorSwatch.addEventListener("mousedown", function () {
            __SC_openBgColorPicker();
        });
    } catch (e) { }

    // マスク（OK時に配置物をマージン内側でクリッピング）
    var cbMask = panelArtboard.add("checkbox", undefined, L("mask"));
    cbMask.value = true;

    // 外側余白
    var groupMargin = panelArtboard.add("group");
    var stMargin = groupMargin.add("statictext", undefined, L("margin"));
    var editMargin = groupMargin.add("edittext", undefined, String(__SC_round(__SC_ptToUnit(DEFAULT_MARGIN_PT, rulerUnit.factor), 2)));
    editMargin.characters = 5;
    groupMargin.add("statictext", undefined, rulerUnit.label);

    // editMargin.onChange = function () { updatePreview(); };

    // マスク角丸（マスクしたクリップグループに角丸を適用）
    var groupMaskRound = panelArtboard.add("group");
    groupMaskRound.orientation = "row";
    groupMaskRound.alignChildren = ["left", "center"];

    var cbMaskRound = groupMaskRound.add("checkbox", undefined, L("maskRound"));
    cbMaskRound.value = false;

    var editMaskRound = groupMaskRound.add("edittext", undefined, "20");
    editMaskRound.characters = 3;

    groupMaskRound.add("statictext", undefined, rulerUnit.label);

    // 初期状態
    editMaskRound.enabled = cbMask.value && cbMaskRound.value;

    // --- mask-round UI enable/disable helper and hook ---
    function __SC_updateMaskRoundUI() {
        try {
            var en = !!cbMask.value;
            groupMaskRound.enabled = en;
            // Ensure inner input reflects both mask and checkbox
            editMaskRound.enabled = en && cbMaskRound.value;
        } catch (e) { }
    }

    // --- mask & margin/mask-round UI helper ---
    function __SC_updateMaskUI() {
        try {
            var en = !!cbMask.value;
            groupMargin.enabled = en;
        } catch (e0) { }
        try { __SC_updateMaskRoundUI(); } catch (e1) { }
    }

    // 初期状態（マスクOFFならディム）
    __SC_updateMaskUI();

    // マスク切替で追従
    cbMask.onClick = function () {
        __SC_updateMaskUI();
    };

    cbMaskRound.onClick = function () {
        // マスクOFFのときは常にディム
        editMaskRound.enabled = cbMask.value && cbMaskRound.value;
    };


    // スケール
    var groupScale = panelLayoutRight.add("group");
    groupScale.orientation = "row";
    groupScale.alignChildren = ["left", "center"];

    var stScale = groupScale.add("statictext", undefined, L("scale"));
    stScale.preferredSize.width = LABEL_W;
    var editScale = groupScale.add("edittext", undefined, L("defaultScale"));
    editScale.characters = 4;

    var stScaleUnit = groupScale.add("statictext", undefined, L("unitPercent"));
    stScaleUnit.preferredSize.width = UNIT_W;

    var spacerScale = groupScale.add("group");
    spacerScale.alignment = ["fill", "fill"];
    spacerScale.minimumSize.width = 0;

    var sldScale = groupScale.add("slider", undefined, 100, 10, 250);
    sldScale.preferredSize.width = SLIDER_W;

    function __SC_syncScaleFromEdit() {
        var v = parseFloat(editScale.text);
        if (isNaN(v)) v = 100;
        if (v < 10) v = 10;
        if (v > 250) v = 250;
        v = Math.round(v);
        editScale.text = String(v);
        try { sldScale.value = v; } catch (e) { }
    }

    function __SC_syncScaleFromSlider() {
        var v = Math.round(sldScale.value);
        if (v < 10) v = 10;
        if (v > 250) v = 250;
        editScale.text = String(v);
    }

    // 初期同期
    __SC_syncScaleFromEdit();

    // 連動
    editScale.onChanging = function () { __SC_syncScaleFromEdit(); };
    sldScale.onChanging = function () { __SC_syncScaleFromSlider(); };

    // リアルタイムプレビュー（確定時）

    // スケールUIは常に操作可能（自動フィットのロジックは維持）
    editScale.enabled = true;
    sldScale.enabled = true;


    // 回転（全体を回転）
    var groupRotate = panelLayoutRight.add("group");
    groupRotate.orientation = "row";
    groupRotate.alignChildren = ["left", "center"];

    var cbRotate = groupRotate.add("checkbox", undefined, L("rotate"));
    cbRotate.helpTip = (lang === "ja")
        ? "配置したアイテム全体を回転します（アートボード中心基準）。"
        : "Rotate the entire placed layout (centered on the artboard).";
    cbRotate.preferredSize.width = LABEL_W;
    cbRotate.value = true;

    var editRotate = groupRotate.add("edittext", undefined, L("defaultRotate"));
    editRotate.characters = 4;

    var stRotateUnit = groupRotate.add("statictext", undefined, L("unitDegree"));
    stRotateUnit.preferredSize.width = UNIT_W;

    var spacerRotate = groupRotate.add("group");
    spacerRotate.alignment = ["fill", "fill"];
    spacerRotate.minimumSize.width = 0;

    var sldRotate = groupRotate.add("slider", undefined, -12, -30, 30);
    sldRotate.preferredSize.width = SLIDER_W;

    function __SC_syncRotateFromEdit() {
        var v = parseFloat(editRotate.text);
        if (isNaN(v)) v = 0;
        if (v < -30) v = -30;
        if (v > 30) v = 30;
        v = Math.round(v);
        editRotate.text = String(v);
        try { sldRotate.value = v; } catch (e) { }
    }

    function __SC_syncRotateFromSlider() {
        var v = sldRotate.value;
        v = Math.round(v);
        if (v < -30) v = -30;
        if (v > 30) v = 30;
        editRotate.text = String(v);
    }

    // 初期同期
    __SC_syncRotateFromEdit();

    // 連動
    editRotate.onChanging = function () { __SC_syncRotateFromEdit(); };
    sldRotate.onChanging = function () { __SC_syncRotateFromSlider(); };

    // リアルタイムプレビュー（確定時）

    // 初期状態
    editRotate.enabled = cbRotate.value;
    sldRotate.enabled = cbRotate.value;

    cbRotate.onClick = function () {
        editRotate.enabled = cbRotate.value;
        sldRotate.enabled = cbRotate.value;
    };

    // 位置調整スライダー範囲（アートボードサイズに合わせる）
    var __abSizeUnit = (function () {
        try {
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            var r = ab.artboardRect; // [left, top, right, bottom]
            var wPt = Math.abs(r[2] - r[0]);
            var hPt = Math.abs(r[1] - r[3]);
            var wU = __SC_ptToUnit(wPt, rulerUnit.factor);
            var hU = __SC_ptToUnit(hPt, rulerUnit.factor);
            if (!isFinite(wU) || wU <= 0) wU = 200;
            if (!isFinite(hU) || hU <= 0) hU = 200;
            return { x: wU, y: hU };
        } catch (e) {
            return { x: 200, y: 200 };
        }
    })();

    // 全体位置（スライダー）
    // ※数値表示は不要（スライダーのみ）

    // 横位置
    var groupOffsetX = panelLayoutRight.add("group");
    var cbOffsetX = groupOffsetX.add("checkbox", undefined, L("offsetX"));
    cbOffsetX.preferredSize.width = OFFSET_LABEL_W;
    cbOffsetX.value = false;

    var spacerOffX = groupOffsetX.add("group");
    spacerOffX.alignment = ["fill", "fill"];
    spacerOffX.minimumSize.width = 0;

    var sldOffsetX = groupOffsetX.add("slider", undefined, 0, -__abSizeUnit.x, __abSizeUnit.x);
    sldOffsetX.preferredSize.width = SLIDER_W;
    try {
        if (sldOffsetX.value < -__abSizeUnit.x) sldOffsetX.value = -__abSizeUnit.x;
        if (sldOffsetX.value > __abSizeUnit.x) sldOffsetX.value = __abSizeUnit.x;
    } catch (e) { }
    sldOffsetX.enabled = cbOffsetX.value;
    cbOffsetX.onClick = function () {
        sldOffsetX.enabled = cbOffsetX.value;
    };
    // sldOffsetX.onChange = function () { updatePreview(); };

    // 縦位置（＋で下へ）
    var groupOffsetY = panelLayoutRight.add("group");
    var cbOffsetY = groupOffsetY.add("checkbox", undefined, L("offsetY"));
    cbOffsetY.preferredSize.width = OFFSET_LABEL_W;
    cbOffsetY.value = false;

    var spacerOffY = groupOffsetY.add("group");
    spacerOffY.alignment = ["fill", "fill"];
    spacerOffY.minimumSize.width = 0;

    var sldOffsetY = groupOffsetY.add("slider", undefined, 0, -__abSizeUnit.y, __abSizeUnit.y);
    sldOffsetY.preferredSize.width = SLIDER_W;
    try {
        if (sldOffsetY.value < -__abSizeUnit.y) sldOffsetY.value = -__abSizeUnit.y;
        if (sldOffsetY.value > __abSizeUnit.y) sldOffsetY.value = __abSizeUnit.y;
    } catch (e) { }
    sldOffsetY.enabled = cbOffsetY.value;
    cbOffsetY.onClick = function () {
        sldOffsetY.enabled = cbOffsetY.value;
    };
    // sldOffsetY.onChange = function () { updatePreview(); };

    // -----------------------------------------
    // 列ずらしデフォルト計算
    // A: 各アイテムの高さ（現在のスケール/回転を反映）
    // defaultShift = (A + 間隔) / 2
    // -----------------------------------------

    function __SC_getArtboardInnerSizePt(marginPt) {
        var activeAB = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var abRect = activeAB.artboardRect; // [left, top, right, bottom]
        var abW = Math.abs(abRect[2] - abRect[0]);
        var abH = Math.abs(abRect[1] - abRect[3]);
        var m = (typeof marginPt === "number" && !isNaN(marginPt) && marginPt >= 0) ? marginPt : 0;
        abW -= m * 2;
        abH -= m * 2;
        if (abW <= 0) abW = 1;
        if (abH <= 0) abH = 1;
        return { w: abW, h: abH };
    }

    function __SC_calcAutoScalePctForUI(pages, colsNum, gapPt, marginPt, rotateEnabled, rotateDeg) {
        if (!pages || pages.length === 0) return 100;

        var inner = __SC_getArtboardInnerSizePt(marginPt);
        var abW = inner.w;
        var abH = inner.h;

        var w = 0, h = 0;
        var pageNum0 = pages[0];
        var cropMode0 = __SC_getCropModeFromUI();

        var cached = __SC_getAutoFitMeasure(fileA, cropMode0, pageNum0);
        if (cached) {
            w = cached.w;
            h = cached.h;
        } else {
            var temp = null;
            try {
                if (__SC_isPdfLikeFile(fileA)) {
                    __SC_setPdfCropPreference(cropMode0);
                }
                __SC_setImportPageNumber(fileA, pageNum0);
                temp = doc.placedItems.add();
                temp.file = fileA;
                w = temp.width;
                h = temp.height;
            } catch (e) {
                w = 0; h = 0;
            } finally {
                try { if (temp) temp.remove(); } catch (e2) { }
                try { __SC_resetImportPageNumber(fileA); } catch (e3) { }
            }
            __SC_setAutoFitMeasure(fileA, cropMode0, pageNum0, w, h);
        }

        if (!(w > 0 && h > 0)) return 100;

        var total = pages.length;
        var rowsNum = Math.ceil(total / colsNum);

        var needW = (colsNum * w) + ((colsNum - 1) * gapPt);
        var needH = (rowsNum * h) + ((rowsNum - 1) * gapPt);

        // 回転がある場合は「グリッド全体」を回転させた外接サイズで見積もる
        if (rotateEnabled && rotateDeg !== 0) {
            var rad = Math.abs(rotateDeg) * Math.PI / 180.0;
            var s = Math.sin(rad);
            var c = Math.cos(rad);
            var rW = Math.abs(needW * c) + Math.abs(needH * s);
            var rH = Math.abs(needW * s) + Math.abs(needH * c);
            needW = rW;
            needH = rH;
        }

        if (!(needW > 0 && needH > 0)) return 100;

        var sW = (abW > 0) ? (abW / needW) : 1;
        var sH = (abH > 0) ? (abH / needH) : 1;
        var sMin = Math.min(sW, sH);
        if (!(sMin > 0)) sMin = 1;

        // 整数（%）に丸め
        var pct = Math.round(sMin * 100);
        if (!(pct > 0)) pct = 100;
        return pct;
    }

    function __SC_calcDefaultColShiftPt() {
        var pages = __SC_getTargetPagesFromUI();

        if (!pages || pages.length === 0) pages = [1];

        var colsNum = parseInt(editCols.text, 10) || 5;

        var spacingUnit = parseFloat(editSpacing.text);
        if (isNaN(spacingUnit) || spacingUnit < 0) spacingUnit = 0;
        var marginUnit = parseFloat(editMargin.text);
        if (isNaN(marginUnit) || marginUnit < 0) marginUnit = 0;

        var spacingPt = __SC_unitToPt(spacingUnit, rulerUnit.factor);
        var marginPt = __SC_unitToPt(marginUnit, rulerUnit.factor);

        var rot = parseFloat(editRotate.text);
        if (isNaN(rot)) rot = 0;
        var rotateEnabled = cbRotate.value;

        // scalePct は「追加倍率（%）」として扱う（自動フィットの結果に乗算）
        var userScale = parseFloat(editScale.text);
        if (isNaN(userScale) || userScale <= 0) userScale = 100;
        if (userScale < 1) userScale = 1;

        var baseScale = 100;
        if (AUTO_FIT_ENABLED) {
            baseScale = __SC_calcAutoScalePctForUI(pages, colsNum, spacingPt, marginPt, rotateEnabled, rot);
        }

        var scalePct = baseScale * (userScale / 100.0);
        if (!(scalePct > 0)) scalePct = baseScale;

        var temp = null;
        var h = 0;
        try {
            if (__SC_isPdfLikeFile(fileA)) {
                __SC_setPdfCropPreference(__SC_getCropModeFromUI());
            }
            __SC_setImportPageNumber(fileA, pages[0]);
            temp = doc.placedItems.add();
            temp.file = fileA;

            try { temp.resize(scalePct, scalePct); } catch (eResize) { }

            h = temp.height;
        } catch (e) {
            h = 0;
        } finally {
            try { if (temp) temp.remove(); } catch (e2) { }
            try { __SC_resetImportPageNumber(fileA); } catch (e3) { }
        }

        if (!(h > 0)) return 0;
        return (h + spacingPt) / 2;
    }

    // 初期値：列ずらしのデフォルトを計算して反映（チェックはOFFのまま）
    try {
        var defShiftPt = __SC_calcDefaultColShiftPt();
        var defShiftUnit = __SC_ptToUnit(defShiftPt, rulerUnit.factor);
        editColShift.text = String(__SC_round(defShiftUnit, 2));
    } catch (e) {
        // 失敗時は従来の 0
    }

    // =========================================
    // Zoom controls (bottom)
    // =========================================

    var __zoomState = __TMKZoom_captureViewState(doc);

    var zoomCtrl = __TMKZoom_addControls(win, doc, L("zoom"), __zoomState, {
        min: 0.1,
        max: 4,
        sliderWidth: 340,
        margins: [0, 0, 0, 10],
        redraw: true,

        // ✅ lightweight mode
        lightMode: true,
        lightModeLabel: L("lightMode"),
        lightModeDefault: false
    });

    // ボタン類（左：リセット / 右：キャンセル・OK）
    var groupButtons = win.add("group");
    groupButtons.orientation = "row";
    groupButtons.alignChildren = ["fill", "center"];

    // 左
    var leftBtnGroup = groupButtons.add("group");
    leftBtnGroup.orientation = "row";
    leftBtnGroup.alignChildren = ["left", "center"];
    var btnReset = leftBtnGroup.add("button", undefined, L('reset'));
    var cbLightPreview = leftBtnGroup.add("checkbox", undefined, L('lightPreview'));
    cbLightPreview.value = true;

    // 中央スペーサー
    var spacer = groupButtons.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    // 右
    var rightBtnGroup = groupButtons.add("group");
    rightBtnGroup.orientation = "row";
    rightBtnGroup.alignChildren = ["right", "center"];
    var btnCancel = rightBtnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    var btnOk = rightBtnGroup.add("button", undefined, L("ok"), { name: "ok" });

    // (Preview/cache declarations moved earlier in main)

    function resetUIToDefaults() {
        try {
            // Load
            // editPages.text = L("defaultPages");
            cbLightPreview.value = true;

            // Item
            ddCrop.selection = 2;
            ddCrop.enabled = __SC_isPdfFile(fileA);

            cbRound.value = false;
            editRound.text = L("defaultRound");
            editRound.enabled = cbRound.value;

            // Grid
            rbDirV.value = true;

            // 列数: 4
            editCols.text = "4";
            __SC_syncColsFromEdit();

            editSpacing.text = String(__SC_round(__SC_ptToUnit(DEFAULT_SPACING_PT, rulerUnit.factor), 2));
            __SC_syncSpacingFromEdit();

            // Even columns
            cbEvenPlus.value = false;

            // ずらし: OFF
            cbColShift.value = false;
            editColShift.text = L("defaultShift");
            __SC_syncShiftFromEdit();
            editColShift.enabled = cbColShift.value;
            sldColShift.enabled = cbColShift.value;

            // Layout
            editScale.text = L("defaultScale");
            __SC_syncScaleFromEdit();

            // 回転: OFF
            cbRotate.value = false;
            editRotate.text = L("defaultRotate");
            __SC_syncRotateFromEdit();
            editRotate.enabled = cbRotate.value;
            sldRotate.enabled = cbRotate.value;

            // 横方向の位置調整：左右余白が等しくなるよう自動計算
            cbOffsetY.value = false;
            sldOffsetY.value = 0;
            sldOffsetY.enabled = false;

            // default is OFF; we turn ON only if we can compute a sensible center offset
            cbOffsetX.value = false;
            sldOffsetX.value = 0;
            sldOffsetX.enabled = false;

            try {
                // Compute in pt
                var pagesForCenter = parsePageNumbers(editPages.text);
                if (!pagesForCenter || pagesForCenter.length === 0) pagesForCenter = [1];

                var colsForCenter = parseInt(editCols.text, 10);
                if (isNaN(colsForCenter) || colsForCenter < 1) colsForCenter = 1;

                var spacingU = parseFloat(editSpacing.text);
                if (isNaN(spacingU) || spacingU < 0) spacingU = 0;
                var marginU = parseFloat(editMargin.text);
                if (isNaN(marginU) || marginU < 0) marginU = 0;

                var spacingPt = __SC_unitToPt(spacingU, rulerUnit.factor);
                var marginPt = __SC_unitToPt(marginU, rulerUnit.factor);

                var cropMode = __SC_getCropModeFromUI();

                // Base auto-fit scale (rotation OFF for reset)
                var baseScale = __SC_calcAutoScalePctForUI(pagesForCenter, colsForCenter, spacingPt, marginPt, false, 0);
                var userScale = parseFloat(editScale.text);
                if (isNaN(userScale) || userScale <= 0) userScale = 100;
                var finalScale = baseScale * (userScale / 100.0);

                // Measure item width (cache-first)
                var itemW = 0;
                try {
                    if (__previewCache.items && __previewCache.items.length > 0) {
                        // Use cached base width
                        var bw0 = __previewCache.baseW[0];
                        if (!(bw0 > 0)) bw0 = __previewCache.items[0].width;
                        itemW = bw0 * (finalScale / 100.0);
                    } else {
                        // Fallback: temp place once
                        var temp = null;
                        try {
                            if (__SC_isPdfLikeFile(fileA)) {
                                __SC_setPdfCropPreference(cropMode);
                            }
                            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", pagesForCenter[0]);
                            temp = doc.placedItems.add();
                            temp.file = fileA;
                            try { temp.resize(finalScale, finalScale); } catch (_) { }
                            itemW = temp.width;
                        } catch (eTmp) {
                            itemW = 0;
                        } finally {
                            try { if (temp) temp.remove(); } catch (_) { }
                            try { app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1); } catch (_) { }
                        }
                    }
                } catch (_) {
                    itemW = 0;
                }

                if (itemW > 0) {
                    var gridW = (colsForCenter * itemW) + ((colsForCenter - 1) * spacingPt);

                    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
                    var r = ab.artboardRect;
                    var abWpt = Math.abs(r[2] - r[0]);
                    var innerW = abWpt - (marginPt * 2);
                    if (innerW < 1) innerW = 1;

                    var oxPt = (innerW - gridW) / 2;
                    // Convert to unit for slider value
                    var oxU = __SC_ptToUnit(oxPt, rulerUnit.factor);

                    if (isFinite(oxU)) {
                        cbOffsetX.value = true;
                        sldOffsetX.enabled = true;
                        try { sldOffsetX.value = oxU; } catch (_) { }
                    }
                }
            } catch (eCenter) { }

            // Artboard & Mask
            cbBg.value = true;
            editBgHex.text = L("defaultHex");
            __SC_updateBgControls();

            cbMask.value = true;
            // margin back to default
            editMargin.text = String(__SC_round(__SC_ptToUnit(DEFAULT_MARGIN_PT, rulerUnit.factor), 2));
            // update mask UI states
            __SC_updateMaskUI();

            cbMaskRound.value = false;
            editMaskRound.text = "20";
            editMaskRound.enabled = cbMask.value && cbMaskRound.value;

        } catch (e) { }
    }

    btnReset.onClick = function () {
        // Reset UI only; keep loaded items cached
        resetUIToDefaults();

        // If already loaded, re-apply layout immediately
        if (__previewCache.items && __previewCache.items.length > 0) {
            try {
                __SC_applyLayoutToCachedItems();
                var core = (__previewCache.group ? [__previewCache.group] : __previewCache.items);
                previewItems = (__previewCache.bgItem ? [__previewCache.bgItem].concat(core) : core);
                app.redraw();
            } catch (_) { }
        }
    };

    function __SC_calcAutoFitPctFromWH(w, h, total, colsNum, gapPt, abW, abH, evenPlusEnabled, doColShift, colShiftPt, doRotate, rotDeg) {
        if (!(w > 0 && h > 0)) return 100;
        if (!(colsNum > 0)) colsNum = 1;

        var rowsNum = Math.ceil(total / colsNum);

        // 偶数列＋1（スロット追加）: 配置に使うスロット表から最大行数を算出
        if (evenPlusEnabled && colsNum >= 2) {
            var base = Math.floor(total / colsNum);
            var rem = total - (base * colsNum);
            var maxH = 0;
            for (var c = 0; c < colsNum; c++) {
                var hC = base;
                if (c < rem) hC += 1;          // 標準の余り配分
                if ((c % 2) === 1) hC += 1;    // 偶数列（2,4,6...）= 0-based 1,3,5... に +1 スロット
                if (hC > maxH) maxH = hC;
            }
            if (maxH < 1) maxH = 1;
            rowsNum = maxH;
        }

        var needW = (colsNum * w) + ((colsNum - 1) * gapPt);
        var needH = (rowsNum * h) + ((rowsNum - 1) * gapPt);

        if (doColShift && colShiftPt !== 0) {
            needH += Math.abs(colShiftPt);
        }

        // 回転がある場合は「グリッド全体」を回転させた外接サイズで見積もる
        if (doRotate && rotDeg !== 0) {
            var rad = Math.abs(rotDeg) * Math.PI / 180.0;
            var s = Math.sin(rad);
            var c = Math.cos(rad);
            var gW = needW;
            var gH = needH;
            needW = Math.abs(gW * c) + Math.abs(gH * s);
            needH = Math.abs(gW * s) + Math.abs(gH * c);
        }

        if (!(needW > 0 && needH > 0)) return 100;
        var sW = (abW > 0) ? (abW / needW) : 1;
        var sH = (abH > 0) ? (abH / needH) : 1;
        var sMin = Math.min(sW, sH);
        if (!(sMin > 0)) sMin = 1;
        var pct = Math.round(sMin * 100);
        if (!(pct > 0)) pct = 100;
        return pct;
    }

    function __SC_applyLayoutToCachedItems() {
        if (!__previewCache.items || __previewCache.items.length === 0) return;

        // Artboard inner rect (margin excluded)
        var activeAB = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var abRect = activeAB.artboardRect;
        var startX = abRect[0];
        var startY = abRect[1];
        var abW = Math.abs(abRect[2] - abRect[0]);
        var abH = Math.abs(abRect[1] - abRect[3]);

        var marginUnit = parseFloat(editMargin.text);
        if (isNaN(marginUnit) || marginUnit < 0) marginUnit = 0;
        var marginPt = __SC_unitToPt(marginUnit, rulerUnit.factor);

        startX += marginPt;
        startY -= marginPt;
        abW -= marginPt * 2;
        abH -= marginPt * 2;
        if (abW <= 0) abW = 1;
        if (abH <= 0) abH = 1;

        var colsNum = parseInt(editCols.text, 10);
        if (isNaN(colsNum) || colsNum < 1) colsNum = 1;

        var spacingUnit = parseFloat(editSpacing.text);
        if (isNaN(spacingUnit) || spacingUnit < 0) spacingUnit = 0;
        var gapPt = __SC_unitToPt(spacingUnit, rulerUnit.factor);

        var doColShift = !!cbColShift.value;
        var colShiftUnit = parseFloat(editColShift.text);
        if (isNaN(colShiftUnit)) colShiftUnit = 0;
        var colShiftPt = __SC_unitToPt(colShiftUnit, rulerUnit.factor);

        var flowMode = __SC_getFlowMode();
        // Random mode: prepare a persistent shuffle order
        if (flowMode === 2) {
            __SC_ensureRandomOrder();
        } else {
            __SC_clearRandomOrder();
        }
        var evenPlusEnabled = !!cbEvenPlus.value;

        var doRotate = !!cbRotate.value;
        var rot = parseFloat(editRotate.text);
        if (isNaN(rot)) rot = 0;

        var userScale = parseFloat(editScale.text);
        if (isNaN(userScale) || userScale <= 0) userScale = 100;
        if (userScale < 1) userScale = 1;

        // base size from the first loaded item (unscaled)
        var w0 = __previewCache.baseW[0] || __previewCache.items[0].width;
        var h0 = __previewCache.baseH[0] || __previewCache.items[0].height;

        var baseScale = __SC_calcAutoFitPctFromWH(w0, h0, __previewCache.items.length, colsNum, gapPt, abW, abH, evenPlusEnabled, doColShift, colShiftPt, doRotate, rot);
        var finalScale = baseScale * (userScale / 100.0);

        // Even+1 slot heights (slot table) to compute col/row mapping
        var colHeights = null;
        var colStart = null;
        if (evenPlusEnabled && colsNum >= 2) {
            var totalN = __previewCache.items.length;
            var base = Math.floor(totalN / colsNum);
            var rem = totalN - (base * colsNum);
            colHeights = [];
            for (var c = 0; c < colsNum; c++) {
                var hh = base;
                if (c < rem) hh += 1;
                if ((c % 2) === 1) hh += 1; // +1 slot for even columns
                colHeights.push(hh);
            }
            colStart = [0];
            for (var c2 = 0; c2 < colsNum; c2++) colStart[c2 + 1] = colStart[c2] + colHeights[c2];
        }
        function indexToColRow(iIndex) {
            if (colHeights && colStart) {
                for (var c = 0; c < colHeights.length; c++) {
                    if (iIndex >= colStart[c] && iIndex < colStart[c + 1]) {
                        return { col: c, row: iIndex - colStart[c] };
                    }
                }
                return { col: colHeights.length - 1, row: 0 };
            }
            var rowsN = Math.ceil(__previewCache.items.length / colsNum);
            if (flowMode === 1 || flowMode === 2) {
                return { col: Math.floor(iIndex / rowsN), row: (iIndex % rowsN) };
            }
            return { col: (iIndex % colsNum), row: Math.floor(iIndex / colsNum) };
        }

        // Update background fill if enabled
        try {
            if (cbBg.value && __previewCache.bgItem) {
                __previewCache.bgItem.fillColor = __SC_getBgRGBColorOrDefault();
            }
        } catch (_) { }

        // If we previously grouped for rotation, un-rotate back to 0 by delta
        if (__previewCache.group && __previewCache.currentRot !== 0) {
            try {
                __previewCache.group.rotate(-__previewCache.currentRot);
            } catch (_) { }
            __previewCache.currentRot = 0;
        }

        // Resize & position each item (absolute, using baseW/baseH)
        for (var i = 0; i < __previewCache.items.length; i++) {
            var idxItem = i;
            if (flowMode === 2 && __previewCache.randOrder) {
                idxItem = __previewCache.randOrder[i];
            }
            var it = __previewCache.items[idxItem];
            if (!it) continue;

            var bw = __previewCache.baseW[idxItem];
            var bh = __previewCache.baseH[idxItem];
            if (!(bw > 0) || !(bh > 0)) {
                bw = it.width;
                bh = it.height;
            }

            // Apply size (avoid cumulative scaling)
            try {
                it.width = bw * (finalScale / 100.0);
                it.height = bh * (finalScale / 100.0);
            } catch (_) { }

            // Map to grid position
            var cr = indexToColRow(i);
            var x = startX + (cr.col * (it.width + gapPt));
            var y = startY - (cr.row * (it.height + gapPt));

            // Even-column Y shift
            if (doColShift && colShiftPt !== 0 && (cr.col % 2 === 1)) {
                if (colShiftPt > 0) y -= colShiftPt;
                else y += Math.abs(colShiftPt);
            }

            try { it.position = [x, y]; } catch (_) { }
        }

        // -------------------------------------------------
        // Round Corners in preview (heavy mode only)
        // - 軽量プレビューOFFのときだけ、各アイテムをクリップ化して角丸を適用
        // -------------------------------------------------
        try {
            var isHeavyPreview = !cbLightPreview.value;
            if (isHeavyPreview && cbRound.value) {
                var roundUnit = parseFloat(editRound.text);
                if (isNaN(roundUnit) || roundUnit < 0) roundUnit = 0;
                var roundPt = __SC_unitToPt(roundUnit, rulerUnit.factor);

                if (roundPt > 0) {
                    // Ensure every item is a clip GroupItem (retry per-item; do NOT rely on one-shot flag)
                    var allWrapped = true;
                    for (var wi = 0; wi < __previewCache.items.length; wi++) {
                        var it0 = __previewCache.items[wi];
                        if (!it0) { allWrapped = false; continue; }

                        if (it0.typename !== "GroupItem") {
                            allWrapped = false;

                            // Try to wrap this item now
                            var clipGrp = null;
                            try { clipGrp = __SC_wrapWithClipGroup(doc, it0); } catch (_) { clipGrp = null; }

                            if (clipGrp) {
                                // Keep the index mapping stable (important for baseW/baseH and randOrder)
                                __previewCache.items[wi] = clipGrp;

                                // Move the new clip group into the persistent rotation group (if any)
                                try { if (__previewCache.group) clipGrp.moveToEnd(__previewCache.group); } catch (_) { }
                            }
                        }
                    }

                    // Mark as wrapped only if all items are groups
                    __previewCache.previewWrapped = allWrapped;

                    // Apply effect when radius changed OR when we just created groups
                    if (__previewCache.previewRoundRadiusPt !== roundPt) {
                        for (var ei = 0; ei < __previewCache.items.length; ei++) {
                            try {
                                var itG = __previewCache.items[ei];
                                if (itG && itG.typename === "GroupItem") {
                                    __SC_applyRoundCorners([itG], roundPt);
                                }
                            } catch (_) { }
                        }
                        __previewCache.previewRoundRadiusPt = roundPt;
                    }
                }
            }
        } catch (_) { }

        // Rotation (group-based) and center to artboard center
        if (doRotate && rot !== 0 && __previewCache.group) {
            try {
                __previewCache.group.rotate(rot);
                __previewCache.currentRot = rot;
                __SC_moveItemCenterToArtboardCenter(doc, __previewCache.group);

                // Apply offsets after centering
                var oxPt = (cbOffsetX.value) ? __SC_unitToPt(sldOffsetX.value, rulerUnit.factor) : 0;
                var oyPt = (cbOffsetY.value) ? __SC_unitToPt(sldOffsetY.value, rulerUnit.factor) : 0;
                if (oxPt !== 0 || oyPt !== 0) {
                    __previewCache.group.translate(oxPt, -oyPt);
                }
            } catch (_) { }
        } else {
            // Non-rotate offsets: apply by shifting start already is done in non-rotate path; here we just apply nothing.
        }

        app.redraw();
    }

    function __SC_updateBgOnly() {
        try {
            // 読み込み前ならキャンバス更新不要（UI側は __SC_updateBgControls がやる）
            if (!__previewCache.items || __previewCache.items.length === 0) {
                try { app.redraw(); } catch (_) { }
                return;
            }

            if (cbBg.value) {
                // 背景ON：bgItemが無ければ作る
                if (!__previewCache.bgItem) {
                    try {
                        __previewCache.bgItem = __SC_drawArtboardBackground(doc, __SC_getBgRGBColorOrDefault());
                    } catch (_) {
                        __previewCache.bgItem = null;
                    }
                }
                // 色だけ更新
                if (__previewCache.bgItem) {
                    try { __previewCache.bgItem.fillColor = __SC_getBgRGBColorOrDefault(); } catch (_) { }
                    try { __previewCache.bgItem.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
                }
            } else {
                // 背景OFF：bgItemを消す
                if (__previewCache.bgItem) {
                    __SC_removeItemSafe(__previewCache.bgItem);
                    __previewCache.bgItem = null;
                }
            }

            // previewItems も同期
            var core = (__previewCache.group ? [__previewCache.group] : __previewCache.items);
            previewItems = (__previewCache.bgItem ? [__previewCache.bgItem].concat(core) : core);

            try { app.redraw(); } catch (_) { }
        } catch (e) { }
    }

    function clearPreview() {
        // If cache exists, do not remove its items here
        if (__previewCache.items && __previewCache.items.length > 0) {
            return;
        }
        for (var i = previewItems.length - 1; i >= 0; i--) {
            try { previewItems[i].remove(); } catch (e) { }
        }
        previewItems = [];
        app.redraw();
    }

    function __SC_updatePreviewImpl() {
        // 既存キャッシュがある場合は再配置せず更新
        if (__previewCache.items && __previewCache.items.length > 0) {
            __SC_applyLayoutToCachedItems();
            // Keep previewItems for cancel/cleanup paths
            var core = (__previewCache.group ? [__previewCache.group] : __previewCache.items);
            previewItems = (__previewCache.bgItem ? [__previewCache.bgItem].concat(core) : core);
            return;
        }
        // 未読み込みの場合は何もしない（ユーザーが[読み込み]を押す）
    }

    // =========================================
    // Step2+Step3: Centralized realtime-preview wiring (debounced)
    // - 読み込み＞アートボード（editPages）は対象外
    // - debounce: 300ms
    // =========================================

    var PREVIEW_DEBOUNCE_MS = 300;
    var __previewTaskId = null;

    // Expose updatePreview for app.scheduleTask
    $.global.__SC_updatePreview = __SC_updatePreviewImpl;

    function requestPreview() {
        try {
            if (__previewTaskId) {
                app.cancelTask(__previewTaskId);
                __previewTaskId = null;
            }
        } catch (e) { }

        try {
            __previewTaskId = app.scheduleTask("$.global.__SC_updatePreview();", PREVIEW_DEBOUNCE_MS, false);
        } catch (e2) {
            // Fallback: run immediately
            try { __SC_updatePreviewImpl(); } catch (e3) { }
        }
    }

    function wireRealtimePreview() {
        cbLightPreview.onClick = function () { requestPreview(); };
        // ✅ 読み込み＞アートボード（editPages）は対象外

        // アイテム
        ddCrop.onChange = function () { requestPreview(); };
        editRound.onChange = function () { requestPreview(); };
        var _oldRound = cbRound.onClick;
        cbRound.onClick = function () { _oldRound(); requestPreview(); };

        // グリッド：方向
        rbDirH.onClick = function () { __SC_clearRandomOrder(); requestPreview(); };
        rbDirV.onClick = function () { __SC_clearRandomOrder(); requestPreview(); };
        rbDirR.onClick = function () { __SC_clearRandomOrder(); requestPreview(); };

        // グリッド：偶数列＋1
        cbEvenPlus.onClick = function () { requestPreview(); };

        // グリッド：列数
        editCols.onChange = function () { __SC_syncColsFromEdit(); requestPreview(); };
        sldCols.onChange = function () { __SC_syncColsFromSlider(); requestPreview(); };

        // グリッド：間隔
        editSpacing.onChange = function () { __SC_syncSpacingFromEdit(); requestPreview(); };
        sldSpacing.onChange = function () { __SC_syncSpacingFromSlider(); requestPreview(); };

        // グリッド：ズレ
        editColShift.onChange = function () { __SC_syncShiftFromEdit(); requestPreview(); };
        sldColShift.onChange = function () { __SC_syncShiftFromSlider(); requestPreview(); };
        var _oldShift = cbColShift.onClick;
        cbColShift.onClick = function () { _oldShift(); requestPreview(); };

        // アートボード
        editMargin.onChange = function () { requestPreview(); };
        var _oldBg = cbBg.onClick;
        cbBg.onClick = function () {
            if (_oldBg) _oldBg();        // __SC_updateBgControls()
            __SC_updateBgOnly();         // レイアウト再計算しない
        };

        var _oldHex = editBgHex.onChange;
        editBgHex.onChange = function () {
            if (_oldHex) _oldHex();      // __SC_updateBgControls()
            __SC_updateBgOnly();         // レイアウト再計算しない
        };

        // 右カラム：レイアウト
        editScale.onChange = function () { __SC_syncScaleFromEdit(); requestPreview(); };
        sldScale.onChange = function () { __SC_syncScaleFromSlider(); requestPreview(); };

        editRotate.onChange = function () { __SC_syncRotateFromEdit(); requestPreview(); };
        sldRotate.onChange = function () { __SC_syncRotateFromSlider(); requestPreview(); };
        var _oldRot = cbRotate.onClick;
        cbRotate.onClick = function () { _oldRot(); requestPreview(); };

        var _oldOffX = cbOffsetX.onClick;
        cbOffsetX.onClick = function () { _oldOffX(); requestPreview(); };
        sldOffsetX.onChange = function () { requestPreview(); };

        var _oldOffY = cbOffsetY.onClick;
        cbOffsetY.onClick = function () { _oldOffY(); requestPreview(); };
        sldOffsetY.onChange = function () { requestPreview(); };
    }

    // 初期配線
    wireRealtimePreview();

    // =========================================
    // Arrow-key increment for edittexts (↑↓ / Shift / Option)
    // =========================================

    function changeValueByArrowKey(editText, allowNegative, afterChange) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    try { event.preventDefault(); } catch (e0) { }
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    try { event.preventDefault(); } catch (e1) { }
                } else {
                    return;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") {
                    value += delta;
                    try { event.preventDefault(); } catch (e2) { }
                } else if (event.keyName === "Down") {
                    value -= delta;
                    try { event.preventDefault(); } catch (e3) { }
                } else {
                    return;
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    try { event.preventDefault(); } catch (e4) { }
                } else if (event.keyName === "Down") {
                    value -= delta;
                    try { event.preventDefault(); } catch (e5) { }
                } else {
                    return;
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10; // 0.1
            } else {
                value = Math.round(value); // integer
            }

            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);

            if (typeof afterChange === "function") {
                try { afterChange(); } catch (e6) { }
            }
        });
    }

    // Apply to edit fields (do NOT apply to editPages)
    changeValueByArrowKey(editCols, false, function () { __SC_syncColsFromEdit(); requestPreview(); });
    changeValueByArrowKey(editSpacing, false, function () { __SC_syncSpacingFromEdit(); requestPreview(); });
    changeValueByArrowKey(editColShift, true, function () { __SC_syncShiftFromEdit(); requestPreview(); });
    changeValueByArrowKey(editScale, false, function () { __SC_syncScaleFromEdit(); requestPreview(); });
    changeValueByArrowKey(editRotate, true, function () { __SC_syncRotateFromEdit(); requestPreview(); });

    // Left column
    changeValueByArrowKey(editMargin, false, function () { requestPreview(); });
    changeValueByArrowKey(editRound, false, function () { requestPreview(); });
    changeValueByArrowKey(editMaskRound, false, function () { /* mask is OK-only */ });

    // 角丸（ライブエフェクト）
    function __SC_createRoundCornersEffectXML(radiusPt) {
        var r = (typeof radiusPt === "number" && !isNaN(radiusPt)) ? radiusPt : 0;
        if (r < 0) r = 0;
        // NOTE: Illustrator LiveEffect XML expects pt
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + r + ' "/></LiveEffect>';
        return xml;
    }

    function __SC_applyRoundCorners(items, radiusPt) {
        if (!items || items.length === 0) return;
        var xml = __SC_createRoundCornersEffectXML(radiusPt);
        for (var i = 0; i < items.length; i++) {
            try {
                items[i].applyEffect(xml);
            } catch (e) {
                // ignore
            }
        }
    }

    // 各配置アイテムを同サイズのパスでクリップグループ化
    function __SC_wrapWithClipGroup(doc, item) {
        if (!item) return null;

        var b;
        try {
            b = item.geometricBounds;
        } catch (e) {
            try { b = item.visibleBounds; } catch (e2) { return null; }
        }
        if (!b || b.length < 4) return null;

        // bounds: [left, top, right, bottom]
        var left = b[0];
        var top = b[1];
        var w = Math.abs(b[2] - b[0]);
        var h = Math.abs(b[1] - b[3]);
        if (w <= 0) w = 1;
        if (h <= 0) h = 1;

        // マスク用パス（塗りなし・線なし）
        var maskPath = doc.activeLayer.pathItems.rectangle(top, left, w, h);
        maskPath.stroked = false;
        maskPath.filled = false;
        maskPath.clipping = true;

        var grp = doc.groupItems.add();

        // まず中身（placed item）を末尾へ
        try { item.moveToEnd(grp); } catch (e3) { }
        // マスクは最前面（グループ先頭）
        try { maskPath.moveToBeginning(grp); } catch (e4) { }

        grp.clipped = true;
        return grp;
    }

    // グリッド配置を行うメイン関数
    function __SC_makeBlackColor(doc) {
        try {
            if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK) {
                var c = new CMYKColor();
                c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100;
                return c;
            }
        } catch (e) { }

        var r = new RGBColor();
        r.red = 0; r.green = 0; r.blue = 0;
        return r;
    }

    function __SC_drawArtboardBackground(doc, fillColor) {
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var r = ab.artboardRect; // [left, top, right, bottom]
        var left = r[0];
        var top = r[1];
        var w = Math.abs(r[2] - r[0]);
        var h = Math.abs(r[1] - r[3]);

        // rectangle(top, left, width, height)
        var bg = doc.activeLayer.pathItems.rectangle(top, left, w, h);
        bg.stroked = false;
        bg.filled = true;
        bg.fillColor = fillColor || __SC_makeRGBColor(0, 0, 0);

        try { bg.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
        return bg;
    }

    function __SC_applyArtboardMask(doc, itemsToClip, marginPt) {
        if (!itemsToClip || itemsToClip.length === 0) return null;

        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var r = ab.artboardRect; // [left, top, right, bottom]
        var m = (typeof marginPt === "number" && !isNaN(marginPt) && marginPt >= 0) ? marginPt : 0;

        var left = r[0] + m;
        var top = r[1] - m;
        var w = Math.abs(r[2] - r[0]) - (m * 2);
        var h = Math.abs(r[1] - r[3]) - (m * 2);
        if (w <= 0) w = 1;
        if (h <= 0) h = 1;

        // マスク用パス（塗りなし・線なし）
        var maskPath = doc.activeLayer.pathItems.rectangle(top, left, w, h);
        maskPath.stroked = false;
        maskPath.filled = false;
        maskPath.clipping = true;

        var grp = doc.groupItems.add();

        // まずクリップ対象をグループへ（末尾へ）
        for (var i = itemsToClip.length - 1; i >= 0; i--) {
            try {
                itemsToClip[i].moveToEnd(grp);
            } catch (e1) { }
        }

        // マスクは最前面（グループ先頭）に
        try { maskPath.moveToBeginning(grp); } catch (e0) { }

        grp.clipped = true;
        return grp;
    }

    function __SC_getBoundsCenter(bounds) {
        // bounds: [left, top, right, bottom]
        var cx = (bounds[0] + bounds[2]) / 2;
        var cy = (bounds[1] + bounds[3]) / 2;
        return { x: cx, y: cy };
    }

    function __SC_getArtboardCenter(doc) {
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var r = ab.artboardRect; // [left, top, right, bottom]
        return { x: (r[0] + r[2]) / 2, y: (r[1] + r[3]) / 2 };
    }

    function __SC_moveItemCenterToArtboardCenter(doc, item) {
        if (!item) return;
        var b;
        try {
            b = item.geometricBounds;
        } catch (e) {
            try { b = item.visibleBounds; } catch (e2) { return; }
        }
        if (!b || b.length < 4) return;

        var cItem = __SC_getBoundsCenter(b);
        var cAb = __SC_getArtboardCenter(doc);

        var dx = cAb.x - cItem.x;
        var dy = cAb.y - cItem.y;

        try {
            item.translate(dx, dy);
        } catch (e3) { }
    }

    function placeArtboards(targetPages, cols, spacing, scalePct, autoFit, outerMargin, colShiftEnabled, colShiftPt, rotateEnabled, rotateDeg, bgEnabled, bgFillColor, cropMode, offsetXPt, offsetYPt, flowMode, evenPlusEnabled, roundEnabled, roundRadiusPt, lightPreview) {
        var placedItemsOnly = [];
        var bgItem = null;
        var __lightPreview = !!lightPreview;
        if (bgEnabled) {
            try {
                bgItem = __SC_drawArtboardBackground(doc, bgFillColor);
            } catch (eBg) {
                bgItem = null;
            }
        }

        // 現在のアクティブなアートボードの左上を基準点にする
        var activeAB = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var abRect = activeAB.artboardRect; // [left, top, right, bottom]
        var startX = abRect[0];
        var startY = abRect[1];

        var abW = Math.abs(abRect[2] - abRect[0]);
        var abH = Math.abs(abRect[1] - abRect[3]);

        var margin = (typeof outerMargin === "number" && !isNaN(outerMargin) && outerMargin >= 0) ? outerMargin : 0;

        var doColShift = !!colShiftEnabled;
        var colShift = (typeof colShiftPt === "number" && !isNaN(colShiftPt)) ? colShiftPt : 0;
        var doRotate = !!rotateEnabled;
        var rot = (typeof rotateDeg === "number" && !isNaN(rotateDeg)) ? rotateDeg : 0;

        // 外側余白を適用
        startX += margin;
        startY -= margin;
        abW -= margin * 2;
        abH -= margin * 2;

        if (abW <= 0) abW = 1;
        if (abH <= 0) abH = 1;

        // 全体位置オフセット（横：＋で右、縦：＋で下）
        var ox = (typeof offsetXPt === "number" && !isNaN(offsetXPt)) ? offsetXPt : 0;
        var oy = (typeof offsetYPt === "number" && !isNaN(offsetYPt)) ? offsetYPt : 0;

        // 回転で全体をアートボード中心に合わせる場合は、ここでのオフセットは二重適用になるため保留し、回転後に適用する
        var willCenterAfterRotate = (doRotate && rot !== 0);
        if (!willCenterAfterRotate) {
            startX += ox;
            startY -= oy;
        }

        // scalePct は「追加倍率（%）」として扱う（自動フィットの結果に乗算）
        function calcAutoScalePct(pages, colsNum, gapPt) {
            if (!pages || pages.length === 0) return 100;
            var w = 0, h = 0;
            var pageNum0 = pages[0];

            // cache hit
            var cached = __SC_getAutoFitMeasure(fileA, cropMode, pageNum0);
            if (cached) {
                w = cached.w;
                h = cached.h;
            } else {
                var temp = null;
                try {
                    if (__SC_isPdfLikeFile(fileA)) {
                        __SC_setPdfCropPreference(cropMode);
                    }
                    app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", pageNum0);
                    temp = doc.placedItems.add();
                    temp.file = fileA;
                    w = temp.width;
                    h = temp.height;
                } catch (e) {
                    w = 0; h = 0;
                } finally {
                    try { if (temp) temp.remove(); } catch (e2) { }
                }

                // store
                __SC_setAutoFitMeasure(fileA, cropMode, pageNum0, w, h);
            }

            if (!(w > 0 && h > 0)) return 100;

            var total = pages.length;
            var rowsNum = Math.ceil(total / colsNum);

            // 偶数列＋1（スロット追加）: 配置に使うスロット表から最大行数を算出
            if (evenPlusEnabled && colsNum >= 2) {
                var base = Math.floor(total / colsNum);
                var rem = total - (base * colsNum);
                var maxH = 0;
                for (var c = 0; c < colsNum; c++) {
                    var hC = base;
                    if (c < rem) hC += 1;
                    if ((c % 2) === 1) hC += 1;
                    if (hC > maxH) maxH = hC;
                }
                if (maxH < 1) maxH = 1;
                rowsNum = maxH;
            }

            var needW = (colsNum * w) + ((colsNum - 1) * gapPt);
            var needH = (rowsNum * h) + ((rowsNum - 1) * gapPt);

            if (doColShift && colShift !== 0) {
                needH += Math.abs(colShift);
            }

            // 回転がある場合は「グリッド全体」を回転させた外接サイズで見積もる
            if (doRotate && rot !== 0) {
                var rad = Math.abs(rot) * Math.PI / 180.0;
                var s = Math.sin(rad);
                var c = Math.cos(rad);

                var gW = needW;
                var gH = needH;

                var rW = Math.abs(gW * c) + Math.abs(gH * s);
                var rH = Math.abs(gW * s) + Math.abs(gH * c);

                needW = rW;
                needH = rH;
            }

            if (!(needW > 0 && needH > 0)) return 100;

            // Auto-fit uses artboard inner size (exclude outer margin)
            var innerW = abW;
            var innerH = abH;

            if (margin > 0) {
                innerW = abW; // already margin-subtracted above in placeArtboards
                innerH = abH; // already margin-subtracted above in placeArtboards
            }

            var sW = (innerW > 0) ? (innerW / needW) : 1;
            var sH = (innerH > 0) ? (innerH / needH) : 1;
            var s = Math.min(sW, sH);
            if (!(s > 0)) s = 1;

            // 整数（%）に丸め
            var pct = Math.round(s * 100);
            if (!(pct > 0)) pct = 100;
            return pct;
        }

        // scalePct は「追加倍率（%）」として扱う（自動フィットの結果に乗算）
        var userScale = (typeof scalePct === "number" && !isNaN(scalePct) && scalePct > 0) ? scalePct : 100;
        if (userScale < 1) userScale = 1;

        var baseScale = 100;
        if (autoFit) {
            baseScale = calcAutoScalePct(targetPages, cols, spacing);
        } else {
            baseScale = 100;
        }

        var finalScale = baseScale * (userScale / 100.0);
        if (!(finalScale > 0)) finalScale = baseScale;

        var refW = null;
        var refH = null;

        // 配置順（ランダム対応）
        function __SC_shuffle(arr) {
            for (var i = arr.length - 1; i > 0; i--) {
                var j = Math.floor(Math.random() * (i + 1));
                var t = arr[i];
                arr[i] = arr[j];
                arr[j] = t;
            }
            return arr;
        }

        var order = [];
        for (var oi = 0; oi < targetPages.length; oi++) order.push(oi);
        if (flowMode === 2) {
            __SC_shuffle(order);
        }

        // 偶数列＋1 用の列高さテーブル（配分モード）を事前計算
        var __evenColHeights = null;
        var __evenColStart = null; // cumulative start index per column
        if (evenPlusEnabled) {
            var __colsN = cols;
            if (!(__colsN > 0)) __colsN = 1;
            if (__colsN >= 2) {
                var __totalN = targetPages.length;
                var __base = Math.floor(__totalN / __colsN);
                var __rem = __totalN - (__base * __colsN);

                __evenColHeights = [];

                // 配分：標準配分（rem） + 偶数列に追加スロット(+1)
                // ※合計スロット数は totalN を超える場合があり、その分は空きになる（4/5/4/5 など）
                for (var __c = 0; __c < __colsN; __c++) {
                    var h = __base;
                    if (__c < __rem) h += 1;        // 標準の余り配分
                    if ((__c % 2) === 1) h += 1;    // 偶数列（2,4,6...）= 0-based 1,3,5... に +1 スロット
                    __evenColHeights.push(h);
                }

                __evenColStart = [0];
                for (var __c4 = 0; __c4 < __colsN; __c4++) {
                    __evenColStart[__c4 + 1] = __evenColStart[__c4] + __evenColHeights[__c4];
                }
            }
        }

        function __SC_indexToColRowEven(iIndex) {
            // returns {col,row}
            if (!__evenColHeights || !__evenColStart) return null;
            var colsN = __evenColHeights.length;
            // linear search is fine for <=20 cols
            for (var c = 0; c < colsN; c++) {
                var start = __evenColStart[c];
                var end = __evenColStart[c + 1];
                if (iIndex >= start && iIndex < end) {
                    return { col: c, row: (iIndex - start) };
                }
            }
            // fallback
            return { col: colsN - 1, row: 0 };
        }

        try {
            for (var i = 0; i < targetPages.length; i++) {
                var abNumber = targetPages[order[i]];
                if (__SC_isPdfLikeFile(fileA)) {
                    __SC_setPdfCropPreference(cropMode);
                }
                __SC_setImportPageNumber(fileA, abNumber);

                var placedItem = doc.placedItems.add();
                placedItem.file = fileA;

                // 先に縮尺（中心基準）
                try {
                    placedItem.resize(finalScale, finalScale);
                } catch (eResize) {
                    // resize 失敗時は無視して続行
                }

                // 1つ目のサイズを基準にグリッド計算（ページごとにサイズが違っても配置が崩れにくい）
                if (refW === null || refH === null) {
                    refW = placedItem.width;
                    refH = placedItem.height;
                }

                // 配置位置の計算（横方向=行優先 / 縦方向=列優先）

                var totalN = targetPages.length;
                var colsN = cols;
                if (!(colsN > 0)) colsN = 1;

                // 通常の行数
                var rowsN = Math.ceil(totalN / colsN);
                if (!(rowsN > 0)) rowsN = 1;

                var colIndex, rowIndex;

                // 偶数列＋1（配分モード）が有効な場合は、方向に関わらず列配分を優先
                if (__evenColHeights && __evenColStart) {
                    var cr = __SC_indexToColRowEven(i);
                    colIndex = cr.col;
                    rowIndex = cr.row;
                } else if (flowMode === 1 || flowMode === 2) {
                    // 縦方向（列優先）/ ランダム（縦配置）
                    rowIndex = i % rowsN;
                    colIndex = Math.floor(i / rowsN);
                } else {
                    // 横方向（行優先）
                    colIndex = i % colsN;
                    rowIndex = Math.floor(i / colsN);
                }

                var posX = startX + (colIndex * (refW + spacing));
                var posY = startY - (rowIndex * (refH + spacing));

                // 偶数列（2列目,4列目...）の上下位置を調整
                // 正の値：下へずらす
                // 負の値：上へずらす
                if (doColShift && colShift !== 0 && (colIndex % 2 === 1)) {
                    if (colShift > 0) {
                        posY -= colShift;
                    } else {
                        posY += Math.abs(colShift);
                    }
                }

                placedItem.position = [posX, posY];

                if (__lightPreview) {
                    // 軽量プレビュー：クリップ/角丸は省略
                    placedItemsOnly.push(placedItem);
                } else {
                    // 配置したアイテムを同サイズのパスでクリップグループ化
                    var clipGrp = null;
                    try {
                        clipGrp = __SC_wrapWithClipGroup(doc, placedItem);
                    } catch (eClip) {
                        clipGrp = null;
                    }

                    // 角丸（ライブエフェクト）：クリップグループに適用
                    if (roundEnabled && (typeof roundRadiusPt === "number") && !isNaN(roundRadiusPt) && roundRadiusPt > 0) {
                        if (clipGrp) {
                            __SC_applyRoundCorners([clipGrp], roundRadiusPt);
                        } else {
                            // フォールバック：クリップ化に失敗した場合はアイテムに適用
                            __SC_applyRoundCorners([placedItem], roundRadiusPt);
                        }
                    }

                    placedItemsOnly.push(clipGrp ? clipGrp : placedItem);
                }
            }

            // 列ずらし適用後に「全体」を回転（背景は回転させない）
            if (doRotate && rot !== 0 && placedItemsOnly.length > 0) {
                var grp = null;
                try {
                    grp = doc.groupItems.add();

                    // moveToBeginning は順序が反転しやすいので逆順で移動して見た目順を維持
                    for (var gi = placedItemsOnly.length - 1; gi >= 0; gi--) {
                        try { placedItemsOnly[gi].moveToBeginning(grp); } catch (eMove) { }
                    }

                    try { grp.rotate(rot); } catch (eGrpRot) { }

                    // 回転後：アイテム全体の中心をアートボード中心に合わせる
                    __SC_moveItemCenterToArtboardCenter(doc, grp);

                    // 回転適用時も位置調整（横：＋で右、縦：＋で下）
                    if (ox !== 0 || oy !== 0) {
                        try { grp.translate(ox, -oy); } catch (ePos) { }
                    }

                    placedItemsOnly = [grp];
                } catch (eGrp) {
                    // グループ化/回転に失敗しても配置済みアイテムは残す
                }
            }

        } catch (e) {
            alert(L("alertPlaceError"));
        } finally {
            try { app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1); } catch (_) { }
            try { __SC_resetImportPageNumber(fileA); } catch (_) { }
        }
        var allItems = placedItemsOnly;
        if (bgItem) {
            allItems = [bgItem].concat(placedItemsOnly);
        }
        return {
            items: allItems,
            maskItems: placedItemsOnly,
            scale: finalScale
        };
    }

    // プレビューボタン
    btnPreview.onClick = function () {
        // Rebuild cache on explicit load
        __SC_clearPreviewCache();

        var pages = __SC_getTargetPagesFromUI();
        if (!pages || pages.length === 0) return;
        // Repeat pages/artboards when requested count exceeds source count
        try {
            var srcCount = __SC_getSourcePageCount(fileA);
            if (srcCount > 0) {
                pages = __SC_repeatPagesWithinCount(pages, srcCount);
            }
        } catch (_) { }
        if (!pages || pages.length === 0) return;

        var cropMode = __SC_getCropModeFromUI();

        // Background (always created on load if enabled)
        if (cbBg.value) {
            try {
                __previewCache.bgItem = __SC_drawArtboardBackground(doc, __SC_getBgRGBColorOrDefault());
            } catch (_) {
                __previewCache.bgItem = null;
            }
        }

        // Place items once
        for (var i = 0; i < pages.length; i++) {
            try {
                if (__SC_isPdfLikeFile(fileA)) {
                    __SC_setPdfCropPreference(cropMode);
                }
                __SC_setImportPageNumber(fileA, pages[i]);
                var it = doc.placedItems.add();
                it.file = fileA;
                __previewCache.items.push(it);
                __previewCache.baseW.push(it.width);
                __previewCache.baseH.push(it.height);
            } catch (_) {
                // ignore individual failures
            }
        }
        try { __SC_resetImportPageNumber(fileA); } catch (_) { }

        __previewCache.cropMode = cropMode;
        __previewCache.pagesKey = String(editPages.text);

        // --- create persistent group for fast rotation ---
        try {
            var grp = doc.groupItems.add();
            for (var gi = __previewCache.items.length - 1; gi >= 0; gi--) {
                try { __previewCache.items[gi].moveToBeginning(grp); } catch (_) { }
            }
            __previewCache.group = grp;
            __previewCache.currentRot = 0;
        } catch (eGrp) {
            __previewCache.group = null;
            __previewCache.currentRot = 0;
        }

        // Apply layout immediately
        __SC_applyLayoutToCachedItems();
        var core = (__previewCache.group ? [__previewCache.group] : __previewCache.items);
        previewItems = (__previewCache.bgItem ? [__previewCache.bgItem].concat(core) : core);
    };

    // キャンセルボタン
    btnCancel.onClick = function () {
        // Stop any pending debounced preview task (prevents callbacks after close)
        try {
            if (__previewTaskId) {
                app.cancelTask(__previewTaskId);
                __previewTaskId = null;
            }
        } catch (_) { }

        // Preview cache removal is enough (clearPreview() would early-return when cache exists)
        __SC_clearPreviewCache();

        try { if (zoomCtrl) zoomCtrl.restoreInitial(); } catch (eZ) { }
        __SC_saveDialogBounds(win.bounds);
        win.close(2);
    };

    // OKボタン
    btnOk.onClick = function () {
        // ★ プレビュー用キャッシュを完全削除（これがないと2セットになる）
        try {
            __SC_clearPreviewCache();
        } catch (_) { }

        __SC_saveDialogBounds(win.bounds);
        win.close(1);
    };

    if (win.show() === 1) {
        var finalPages = __SC_getTargetPagesFromUI();
        var finalCols = parseInt(editCols.text, 10) || 5;
        var finalSpacingUnit = parseFloat(editSpacing.text);
        if (isNaN(finalSpacingUnit) || finalSpacingUnit < 0) finalSpacingUnit = 0;
        var finalMarginUnit = parseFloat(editMargin.text);
        if (isNaN(finalMarginUnit) || finalMarginUnit < 0) finalMarginUnit = 0;

        var finalSpacing = __SC_unitToPt(finalSpacingUnit, rulerUnit.factor);
        var finalMargin = __SC_unitToPt(finalMarginUnit, rulerUnit.factor);

        var finalScalePct = parseFloat(editScale.text);
        if (isNaN(finalScalePct) || finalScalePct <= 0) finalScalePct = 100;

        var finalColShiftUnit = parseFloat(editColShift.text);
        if (isNaN(finalColShiftUnit)) finalColShiftUnit = 0;
        var finalColShiftPt = __SC_unitToPt(finalColShiftUnit, rulerUnit.factor);

        var finalRot = parseFloat(editRotate.text);
        if (isNaN(finalRot)) finalRot = 0;

        // 全体位置（スライダー値は定規単位とみなし pt に変換）
        var finalOffsetXPt = (cbOffsetX.value) ? __SC_unitToPt(sldOffsetX.value, rulerUnit.factor) : 0;
        var finalOffsetYPt = (cbOffsetY.value) ? __SC_unitToPt(sldOffsetY.value, rulerUnit.factor) : 0;

        // 角丸設定
        var finalRoundEnabled = cbRound.value;
        var finalRoundUnit = parseFloat(editRound.text);
        if (isNaN(finalRoundUnit) || finalRoundUnit < 0) finalRoundUnit = 0;
        var finalRoundPt = __SC_unitToPt(finalRoundUnit, rulerUnit.factor);

        var cropMode = __SC_getCropModeFromUI();
        var bgFillColor = __SC_getBgRGBColorOrDefault();
        var res = placeArtboards(finalPages, finalCols, finalSpacing, finalScalePct, AUTO_FIT_ENABLED, finalMargin, cbColShift.value, finalColShiftPt, cbRotate.value, finalRot, cbBg.value, bgFillColor, cropMode, finalOffsetXPt, finalOffsetYPt, __SC_getFlowMode(), cbEvenPlus.value, finalRoundEnabled, finalRoundPt, false);

        // OK時のみ：マージン内側で配置物をマスク（背景は含めない）
        if (cbMask.value && res && res.maskItems && res.maskItems.length > 0) {
            var maskGrp = __SC_applyArtboardMask(doc, res.maskItems, finalMargin);

            // マスク角丸：マスクグループに適用
            if (maskGrp && cbMaskRound.value) {
                var mrUnit = parseFloat(editMaskRound.text);
                if (isNaN(mrUnit) || mrUnit < 0) mrUnit = 0;
                var mrPt = __SC_unitToPt(mrUnit, rulerUnit.factor);
                if (mrPt > 0) {
                    try { __SC_applyRoundCorners([maskGrp], mrPt); } catch (eMR) { }
                }
            }
        }
    }

    // ダイアログ終了後に選択解除 / Clear selection after closing dialog
    try { doc.selection = null; } catch (eSel) { }
}

main();