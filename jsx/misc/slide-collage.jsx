#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
slide-collage

### 更新日：
20260301

### 概要：
アクティブなドキュメント上で、指定した .ai / .pdf（PDFはページ指定）をグリッド配置し、ポートフォリオ用のサムネイル一覧を作成します。

・読み込み：アートボード番号（例 1-20 / 1,3,5）を指定して配置
・アイテム：PDFの配置範囲（アート/トリミング/仕上がり/裁ち落とし）を選択、角丸（pt換算）を適用
  - 各アイテムは同サイズの矩形でクリップグループ化し、角丸（ライブエフェクト）はクリップグループに適用

・グリッド：方向（横/縦/ランダム）、列数、間隔を設定
・偶数列：配分モード（偶数列＋1）で偶数列に+1スロットを追加、ずらし（偶数列の上下オフセット）を個別調整
・レイアウト：スケール（自動フィット結果に対する追加倍率）、回転（全体を回転し中心をアートボード中心に合わせる）、位置調整（横/縦）

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

var SCRIPT_VERSION = "v1.0";

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
    panelArtboard: { ja: "マスク", en: "Mask" },
    panelArtboardBg: { ja: "アートボード", en: "Artboard" },

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
    evenPlus: { ja: "偶数列＋1", en: "Even+1" },
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
    bg: { ja: "背景色", en: "Background" },

    // Buttons
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },

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

function __SC_isPdfLikeFile(f) {
    try {
        if (!f) return false;
        var n = (f.name || "").toLowerCase();
        return (n.indexOf(".pdf") > -1);
    } catch (e) {
        return false;
    }
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

function main() {
    if (app.documents.length === 0) {
        alert(L("alertNeedDoc"));
        return;
    }

    var doc = app.activeDocument;

    // 現在の定規単位（rulerType）
    var rulerUnit = __SC_getRulerUnitInfo();

    // 既定値は pt ベースで保持し、表示時に定規単位へ変換
    var DEFAULT_SPACING_PT = 20;
    var DEFAULT_MARGIN_PT = 20;

    // ラベル幅（列数/間隔/列ずらし/スケール/回転/横/縦 を揃える）
    var LABEL_W = 60;
    // 位置調整ラベル幅（横方向の位置調整/縦方向の位置調整）
    var OFFSET_LABEL_W = 140;
    // 単位・補助ラベル幅（空白/pt/%/° を揃える）
    var UNIT_W = 24;

    // スライダー幅（列数/列ずらし/スケール/回転/横/縦 を揃える）
    var SLIDER_W = 140;

    // 自動フィット（UIは非表示。ロジックは維持して常にON）
    var AUTO_FIT_ENABLED = true;


    // 1. ファイル（A）を指定
    var fileA = File.openDialog(L("fileDialogTitle"));
    if (!fileA) return;

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
    panelPage.margins = 15;

    var groupPage = panelPage.add("group");
    groupPage.orientation = "row";
    groupPage.alignChildren = ["left", "center"];

    // ラベルは省略（行内で input + button にする）
    var editPages = groupPage.add("edittext", undefined, L("defaultPages"));
    editPages.characters = 8;
    editPages.active = true;

    var btnPreview = groupPage.add("button", undefined, L("btnLoad"));
    btnPreview.preferredSize = [80, 22];

    // --- トリミングパネル（配置範囲：PDF のトリミング設定） ---
    var panelCrop = leftCol.add("panel", undefined, L("panelItem"));
    panelCrop.alignChildren = "left";
    panelCrop.margins = 15;

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
    // デフォルトでディム表示
    ddCrop.enabled = false;

    function __SC_getCropModeFromUI() {
        var idx = (ddCrop.selection) ? ddCrop.selection.index : 2;
        // 0:アート / 1:トリミング / 2:仕上がり / 3:裁ち落とし
        if (idx === 0) return __SC_CROP_ART;
        if (idx === 1) return __SC_CROP_TRIM;
        if (idx === 3) return __SC_CROP_BLEED;
        // 仕上がりは CropBox を想定（環境差があるため失敗時は無視される）
        return __SC_CROP_CROP;
    }

    /* グリッド / Grid */
    var panelLayout = rightCol.add("panel", undefined, L("panelGrid"));
    panelLayout.alignChildren = "left";
    panelLayout.margins = 15;

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
    panelEven.margins = 15;

    // 配分モード
    var groupDist = panelEven.add("group");
    groupDist.orientation = "row";
    groupDist.alignChildren = ["left", "center"];

    var stDist = groupDist.add("statictext", undefined, L("dist"));
    stDist.preferredSize.width = LABEL_W;

    var cbEvenPlus = groupDist.add("checkbox", undefined, L("evenPlus"));
    cbEvenPlus.value = false;

    // ずらし（Yオフセット）
    var groupColShift = panelEven.add("group");
    groupColShift.orientation = "row";
    groupColShift.alignChildren = ["left", "center"];

    var cbColShift = groupColShift.add("checkbox", undefined, L("shift"));
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
    panelLayoutRight.margins = 15;

    // --- アートボードパネル（左カラム） ---
    var panelArtboard = leftCol.add("panel", undefined, L("panelArtboard"));
    panelArtboard.alignChildren = "left";
    panelArtboard.margins = 15;

    // 外側余白
    var groupMargin = panelArtboard.add("group");
    var stMargin = groupMargin.add("statictext", undefined, L("margin"));
    var editMargin = groupMargin.add("edittext", undefined, String(__SC_round(__SC_ptToUnit(DEFAULT_MARGIN_PT, rulerUnit.factor), 2)));
    editMargin.characters = 5;
    groupMargin.add("statictext", undefined, rulerUnit.label);

    // editMargin.onChange = function () { updatePreview(); };

    // マスク（OK時に配置物をマージン内側でクリッピング）
    var cbMask = panelArtboard.add("checkbox", undefined, L("mask"));
    cbMask.value = true;

    // マスク角丸（マスクしたクリップグループに角丸を適用）
    var groupMaskRound = panelArtboard.add("group");
    groupMaskRound.orientation = "row";
    groupMaskRound.alignChildren = ["left", "center"];

    var cbMaskRound = groupMaskRound.add("checkbox", undefined, L("round"));
    cbMaskRound.value = false;

    var editMaskRound = groupMaskRound.add("edittext", undefined, "20");
    editMaskRound.characters = 3;

    groupMaskRound.add("statictext", undefined, rulerUnit.label);

    // 初期状態
    editMaskRound.enabled = cbMaskRound.value;

    cbMaskRound.onClick = function () {
        editMaskRound.enabled = cbMaskRound.value;
    };

    // --- アートボードパネル（背景色） ---
    var panelArtboardBg = leftCol.add("panel", undefined, L("panelArtboardBg"));
    panelArtboardBg.alignChildren = "left";
    panelArtboardBg.margins = 15;

    // 背景色（アートボードと同じ大きさの図形を作成）
    var groupBg = panelArtboardBg.add("group");
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
            try { requestPreview(); } catch (e) {}
        }
    }

    colorSwatch.onClick = function () {
        __SC_openBgColorPicker();
    };

    try {
        colorSwatch.addEventListener("mousedown", function () {
            __SC_openBgColorPicker();
        });
    } catch (e) {}


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

    var sldOffsetX = groupOffsetX.add("slider", undefined, 0, -200, 200);
    sldOffsetX.preferredSize.width = SLIDER_W;
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

    var sldOffsetY = groupOffsetY.add("slider", undefined, 0, -200, 200);
    sldOffsetY.preferredSize.width = SLIDER_W;
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

        var temp = null;
        var w = 0, h = 0;
        try {
            if (__SC_isPdfLikeFile(fileA)) {
                __SC_setPdfCropPreference(__SC_getCropModeFromUI());
            }
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", pages[0]);
            temp = doc.placedItems.add();
            temp.file = fileA;
            w = temp.width;
            h = temp.height;
        } catch (e) {
            w = 0; h = 0;
        } finally {
            try { if (temp) temp.remove(); } catch (e2) { }
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
        var pages = parsePageNumbers(editPages.text);
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
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", pages[0]);
            temp = doc.placedItems.add();
            temp.file = fileA;

            try { temp.resize(scalePct, scalePct); } catch (eResize) { }

            h = temp.height;
        } catch (e) {
            h = 0;
        } finally {
            try { if (temp) temp.remove(); } catch (e2) { }
            try { app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1); } catch (e3) { }
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

    // ボタン類（左：キャンセル・OK）
    var groupButtons = win.add("group");
    groupButtons.orientation = "row";
    groupButtons.alignChildren = ["fill", "center"];

    // 左（スペースのみ、ボタンなし）
    var leftBtnGroup = groupButtons.add("group");
    leftBtnGroup.orientation = "row";
    leftBtnGroup.alignChildren = ["left", "center"];

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

    // プレビュー用に配置したアイテムを保持する配列
    var previewItems = [];

    function clearPreview() {
        for (var i = previewItems.length - 1; i >= 0; i--) {
            try { previewItems[i].remove(); } catch (e) { }
        }
        previewItems = [];
        app.redraw();
    }

    function updatePreview() {
        // アートボード（ページ）入力の変更では呼ばない前提
        clearPreview();

        var pages = parsePageNumbers(editPages.text);
        if (!pages || pages.length === 0) return;

        var cols = parseInt(editCols.text, 10) || 5;

        var spacingUnit = parseFloat(editSpacing.text);
        if (isNaN(spacingUnit) || spacingUnit < 0) spacingUnit = 0;
        var marginUnit = parseFloat(editMargin.text);
        if (isNaN(marginUnit) || marginUnit < 0) marginUnit = 0;

        var spacing = __SC_unitToPt(spacingUnit, rulerUnit.factor);
        var margin = __SC_unitToPt(marginUnit, rulerUnit.factor);

        var scalePct = parseFloat(editScale.text);
        if (isNaN(scalePct) || scalePct <= 0) scalePct = 100;

        var colShiftUnit = parseFloat(editColShift.text);
        if (isNaN(colShiftUnit)) colShiftUnit = 0;
        var colShiftPt = __SC_unitToPt(colShiftUnit, rulerUnit.factor);

        var rot = parseFloat(editRotate.text);
        if (isNaN(rot)) rot = 0;

        // 全体位置（スライダー値は定規単位とみなし pt に変換）
        var offsetXPt = (cbOffsetX.value) ? __SC_unitToPt(sldOffsetX.value, rulerUnit.factor) : 0;
        var offsetYPt = (cbOffsetY.value) ? __SC_unitToPt(sldOffsetY.value, rulerUnit.factor) : 0;

        // 角丸設定
        var roundEnabled = cbRound.value;
        var roundUnit = parseFloat(editRound.text);
        if (isNaN(roundUnit) || roundUnit < 0) roundUnit = 0;
        var roundPt = __SC_unitToPt(roundUnit, rulerUnit.factor);

        var cropMode = __SC_getCropModeFromUI();
        var bgFillColor = __SC_getBgRGBColorOrDefault();

        var r = placeArtboards(
            pages,
            cols,
            spacing,
            scalePct,
            AUTO_FIT_ENABLED,
            margin,
            cbColShift.value,
            colShiftPt,
            cbRotate.value,
            rot,
            cbBg.value,
            bgFillColor,
            cropMode,
            offsetXPt,
            offsetYPt,
            __SC_getFlowMode(),
            cbEvenPlus.value,
            roundEnabled,
            roundPt
        );

        previewItems = r.items;
        app.redraw();
    }

    // =========================================
    // Step2+Step3: Centralized realtime-preview wiring (debounced)
    // - 読み込み＞アートボード（editPages）は対象外
    // - debounce: 120ms
    // =========================================

    var PREVIEW_DEBOUNCE_MS = 120;
    var __previewTaskId = null;

    // Expose updatePreview for app.scheduleTask
    $.global.__SC_updatePreview = updatePreview;

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
            try { updatePreview(); } catch (e3) { }
        }
    }

    function wireRealtimePreview() {
        // ✅ 読み込み＞アートボード（editPages）は対象外

        // アイテム
        ddCrop.onChange = function () { requestPreview(); };
        editRound.onChange = function () { requestPreview(); };
        var _oldRound = cbRound.onClick;
        cbRound.onClick = function () { _oldRound(); requestPreview(); };

        // グリッド：方向
        rbDirH.onClick = function () { requestPreview(); };
        rbDirV.onClick = function () { requestPreview(); };
        rbDirR.onClick = function () { requestPreview(); };

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
        cbBg.onClick = function () { _oldBg(); requestPreview(); };
        var _oldHex = editBgHex.onChange;
        editBgHex.onChange = function () {
            if (_oldHex) _oldHex();
            requestPreview();
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

    function placeArtboards(targetPages, cols, spacing, scalePct, autoFit, outerMargin, colShiftEnabled, colShiftPt, rotateEnabled, rotateDeg, bgEnabled, bgFillColor, cropMode, offsetXPt, offsetYPt, flowMode, evenPlusEnabled, roundEnabled, roundRadiusPt) {
        var placedItemsOnly = [];
        var bgItem = null;
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
            var temp = null;
            var w = 0, h = 0;
            try {
                if (__SC_isPdfLikeFile(fileA)) {
                    __SC_setPdfCropPreference(cropMode);
                }
                app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", pages[0]);
                temp = doc.placedItems.add();
                temp.file = fileA;
                w = temp.width;
                h = temp.height;
            } catch (e) {
                w = 0; h = 0;
            } finally {
                try { if (temp) temp.remove(); } catch (e2) { }
            }

            if (!(w > 0 && h > 0)) return 100;

            var total = pages.length;
            var rowsNum = Math.ceil(total / colsNum);

            // 偶数列＋1（スロット追加）: 最大行数は base + (rem>0?1:0) + 1
            if (evenPlusEnabled && colsNum >= 2) {
                var base = Math.floor(total / colsNum);
                var rem = total - (base * colsNum);
                rowsNum = base + ((rem > 0) ? 1 : 0) + 1;
                if (rowsNum < 1) rowsNum = 1;
            }

            // 偶数列＋1（配分モード）の場合、最大行数は base+1 になることがある
            if (evenPlusEnabled && colsNum >= 2) {
                var base = Math.floor(total / colsNum);
                var rem = total - (base * colsNum);
                rowsNum = base + ((rem > 0) ? 1 : 0);
                if (rowsNum < 1) rowsNum = 1;
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

            var sW = (abW > 0) ? (abW / needW) : 1;
            var sH = (abH > 0) ? (abH / needH) : 1;
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
                app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", abNumber);

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
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1);
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
        updatePreview();
    };

    // キャンセルボタン
    btnCancel.onClick = function () {
        clearPreview();
        __SC_saveDialogBounds(win.bounds);
        win.close(2);
    };

    // OKボタン
    btnOk.onClick = function () {
        clearPreview();
        __SC_saveDialogBounds(win.bounds);
        win.close(1);
    };

    if (win.show() === 1) {
        var finalPages = parsePageNumbers(editPages.text);
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
        var res = placeArtboards(finalPages, finalCols, finalSpacing, finalScalePct, AUTO_FIT_ENABLED, finalMargin, cbColShift.value, finalColShiftPt, cbRotate.value, finalRot, cbBg.value, bgFillColor, cropMode, finalOffsetXPt, finalOffsetYPt, __SC_getFlowMode(), cbEvenPlus.value, finalRoundEnabled, finalRoundPt);

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