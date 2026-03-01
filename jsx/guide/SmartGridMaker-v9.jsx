#targetengine "SmartGridMakerEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  SmartGridMaker.jsx
  囲み罫とグリッド

  更新日: 2026-02-15

  長方形またはアートボードを基準に、
  外側（辺の伸縮・線端）、タイトルエリア、
  内側のオフセット（4方向・連動）、
  列／行の分割・間隔設定、塗り・分割線・線種（実線・点線・ドット点線）、
  フレーム（裁ち落とし対応）を
  プレビュー付きで一括生成するスクリプト。

  ・アートボード基準時はマージン指定可能
  ・フレームはアートボード基準で作成（裁ち落としはフレームのみに適用）
  ・UIは日英ローカライズ対応
  ・内側エリアの［塗り］は手動OFFを優先（ガター変更時の自動ONを抑制）
  ・プレビュー時のUndo履歴を汚さないための処理（PreviewHistory）を削除
  ・スクリプト再実行時に前回ダイアログの値を復元（Illustrator再起動でリセット）
  ・外側エリアに角丸オプション（UIのみ／rulerType単位）を追加
  ・内側エリアのオフセットUIを3段組レイアウトに変更
  ・マージンを上下左右に分割し、連動（3段組レイアウト）に対応
  ・UI生成を関数分割（MarginUI / OuterUI / InnerUI）
  ・セッション復元の保存形式を構造体化（旧フラット形式もフォールバック）
  ・プレビュー処理と生成処理を分離（計算→生成→描画）
    ・プレビュー/最終生成ともに collectOptions() を経由して計算を共通化
  ・長方形選択で開始した場合はフレームpanelを無効化（アートボード基準のみ有効）
  ・長方形選択で開始した場合はマージンpanelも非表示（スペースを詰める）
  ・タイトルエリアに［辺の伸縮］を追加（タイトル帯の線の長さに反映）
    ・タイトルエリアの［辺の伸縮］は正負を反転（＋で伸ばす／−で短くする）
  ・タイトルエリアの［線］がOFFのとき［辺の伸縮］をディム表示
*/

/* バージョン / Version */
var SCRIPT_VERSION = "v1.3";

// =========================
// Session-persistent UI state (kept while Illustrator is running)
// =========================
var __ENGINE_KEY = "__SmartGridMaker__";
$.global[__ENGINE_KEY] = $.global[__ENGINE_KEY] || {};
function __loadState() {
    try { return $.global[__ENGINE_KEY] || {}; } catch (_) { return {}; }
}
function __saveState(obj) {
    try { $.global[__ENGINE_KEY] = obj || {}; } catch (_) { }
}


/* 言語判定 / Detect language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "囲み罫とグリッド", en: "Frame and Grid" },

    // Panels
    panelMargin: { ja: "マージン", en: "Margin" },
    panelOuter: { ja: "外側エリア", en: "Outer Area" },
    panelCap: { ja: "線端", en: "Line Caps" },
    panelLine: { ja: "線", en: "Line" },
    panelTitleBand: { ja: "タイトルエリア", en: "Title Area" },
    panelFrame: { ja: "フレーム", en: "Frame" },
    panelInnerArea: { ja: "内側エリア", en: "Inner Area" },
    panelOffset: { ja: "オフセット", en: "Offset" },
    panelColumns: { ja: "列", en: "Columns" },
    panelRows: { ja: "行", en: "Rows" },
    panelLineType: { ja: "線の種類", en: "Line Type" },
    panelDisplay: { ja: "画面表示", en: "Display" },
    panelZoomPan: { ja: "ズームとパン", en: "Pan & Zoom" },


    // Common
    preview: { ja: "プレビュー", en: "Preview" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "実行", en: "OK" },
    zoomLabel: { ja: "ズーム", en: "Zoom" }, // 「：」なし（UI側で整える）

    alertSelectPath: {
        ja: "パスアイテム（長方形など）を選択してください。",
        en: "Please select a path item (e.g., rectangle)."
    },

    // Outer
    chkKeepOuter: { ja: "外枠を残す", en: "Keep outer frame" },
    chkEdgeScale: { ja: "辺の伸縮", en: "Edge scale" },

    // Caps
    capNone: { ja: "なし", en: "Butt" },
    capRound: { ja: "丸型", en: "Round" },
    capProject: { ja: "突出", en: "Project" },

    // Title band
    titleTop: { ja: "上", en: "Top" },
    titleBottom: { ja: "下", en: "Bottom" },
    titleLeft: { ja: "左", en: "Left" },
    titleRight: { ja: "右", en: "Right" },
    titleSize: { ja: "幅／高さ", en: "Size" },
    chkTitleEnable: { ja: "有効", en: "Enable" },
    chkFill: { ja: "塗り", en: "Fill" },

    // Frame
    chkBleed: { ja: "裁ち落とし", en: "Bleed" },
    chkFrameRound: { ja: "角丸", en: "Round" },

    // Inner area
    offsetTop: { ja: "上", en: "Top" },
    offsetBottom: { ja: "下", en: "Bottom" },
    offsetLeft: { ja: "左", en: "Left" },
    offsetRight: { ja: "右", en: "Right" },
    chkLink: { ja: "連動", en: "Link" },

    // Grid
    colCount: { ja: "列数", en: "Count" },
    rowCount: { ja: "行数", en: "Count" },
    spacing: { ja: "間隔", en: "Spacing" },
    chkDivider: { ja: "分割線", en: "Dividers" },

    // Line types
    lineSolid: { ja: "実線", en: "Solid" },
    lineDash: { ja: "点線", en: "Dash" },
    lineDotDash: { ja: "ドット点線", en: "Dots" },


    lr: { ja: "左右", en: "Pan L/R" },
    ud: { ja: "上下", en: "Pan U/D" },
};

/* ラベル取得 / Get localized label */
function L(key) {
    try {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].ja) return LABELS[key].ja;
    } catch (_) { }
    return key;
}

/* =========================================
 * ViewControl util (extractable)
 * Zoom + Pan (L/R, U/D) for Illustrator view
 *
 * UI (slider only):
 *   ズーム <===== slider =====>
 *   左右   <===== slider =====>
 *   上下   <===== slider =====>
 *
 * - Base center: active artboard center
 * - Pan offsets: relative to base center
 * - Zoom keeps current pan offsets
 * - Hold Option(Alt) while dragging => 1/10 speed (delta*0.1)
 * - Pan range: dynamic (half of artboard W/H), safety clamped
 *
 * How to use:
 *   var vc = ViewControl.create(doc, L);     // L optional
 *   vc.buildUI(pnlView, { labelWidth:58, sliderWidth:260 });
 *   // Cancel:
 *   vc.restore();
 * ========================================= */

var ViewControl = (function () {

    function clamp(v, mn, mx) { return (v < mn) ? mn : (v > mx) ? mx : v; }

    function getAltKey() {
        try { return !!(ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey); } catch (_) { }
        return false;
    }

    function applySliderAltFine(slider, st, applyFn) {
        try {
            if (!slider || !st || typeof applyFn !== "function") return;

            var raw = Number(slider.value);
            if (isNaN(raw)) raw = 0;

            if (st.raw == null || st.eff == null) {
                st.raw = raw;
                st.eff = raw;
            }

            if (getAltKey()) {
                var d = raw - Number(st.raw);
                var eff = Number(st.eff) + d * 0.1;
                st.raw = raw;
                st.eff = eff;
                try { slider.value = eff; } catch (_) { }
                applyFn(eff);
            } else {
                st.raw = raw;
                st.eff = raw;
                applyFn(raw);
            }
        } catch (_) { }
    }

    function getActiveArtboardCenter(doc, view) {
        try {
            var idx = doc.artboards.getActiveArtboardIndex();
            var r = doc.artboards[idx].artboardRect; // [L, T, R, B]
            return [r[0] + (r[2] - r[0]) / 2, r[1] + (r[3] - r[1]) / 2];
        } catch (_) { }
        try { return (view && view.centerPoint) ? view.centerPoint : [0, 0]; } catch (_) { }
        return [0, 0];
    }

    function getPanRangePt(doc) {
        // half of artboard size, with safety clamps
        try {
            var idx = doc.artboards.getActiveArtboardIndex();
            var r = doc.artboards[idx].artboardRect;
            var w = Math.abs(r[2] - r[0]);
            var h = Math.abs(r[1] - r[3]);
            var xMax = Math.round(w / 2);
            var yMax = Math.round(h / 2);
            if (!xMax || xMax < 100) xMax = 100;
            if (!yMax || yMax < 100) yMax = 100;
            if (xMax > 50000) xMax = 50000;
            if (yMax > 50000) yMax = 50000;
            return { xMax: xMax, yMax: yMax };
        } catch (_) { }
        return { xMax: 2000, yMax: 2000 };
    }

    function clampZoomFactor(z) {
        // Illustrator zoom factor safety
        if (z < 0.0313) z = 0.0313;
        if (z > 640.0) z = 640.0;
        return z;
    }

    function create(doc, Lfn) {
        var vc = {};

        vc.doc = doc;
        vc.view = null;
        vc.L = (typeof Lfn === "function") ? Lfn : null;

        vc.orgZoom = null;
        vc.orgCenter = null;

        vc.panX = 0; // pt
        vc.panY = 0; // pt (UI positive => down)
        vc.panRange = { xMax: 2000, yMax: 2000 };

        vc.sldZoom = null;
        vc.sldPanX = null;
        vc.sldPanY = null;

        try {
            vc.view = doc.views[0];
            vc.orgZoom = vc.view.zoom;
            vc.orgCenter = vc.view.centerPoint;
        } catch (_) { }

        vc.t = function (key, fallback) {
            try { if (vc.L) return vc.L(key); } catch (_) { }
            return fallback || key;
        };

        vc.refreshPanRange = function () {
            try { vc.panRange = getPanRangePt(vc.doc); } catch (_) { }
            return vc.panRange;
        };

        vc.applyCenter = function () {
            try {
                if (!vc.view) return;
                var c = getActiveArtboardCenter(vc.doc, vc.view);
                var x = c[0] + Number(vc.panX || 0);
                // Illustrator: +Y is up. UI: positive is down => subtract
                var y = c[1] - Number(vc.panY || 0);
                vc.view.centerPoint = [x, y];
                app.redraw();
            } catch (_) { }
        };

        vc.setZoomPct = function (pct, zoomMin, zoomMax) {
            try {
                if (!vc.view) return;
                var p = Number(pct);
                if (isNaN(p)) return;
                p = clamp(Math.round(p), zoomMin, zoomMax);
                vc.view.zoom = clampZoomFactor(p / 100.0);
                vc.applyCenter(); // keep pan
            } catch (_) { }
        };

        vc.setPanX = function (v) {
            try {
                var n = Number(v); if (isNaN(n)) n = 0;
                var rg = vc.refreshPanRange();
                n = clamp(Math.round(n), -rg.xMax, rg.xMax);
                vc.panX = n;
                vc.applyCenter();
            } catch (_) { }
        };

        vc.setPanY = function (v) {
            try {
                var n = Number(v); if (isNaN(n)) n = 0;
                var rg = vc.refreshPanRange();
                n = clamp(Math.round(n), -rg.yMax, rg.yMax);
                vc.panY = n;
                vc.applyCenter();
            } catch (_) { }
        };

        vc.restore = function () {
            try {
                if (vc.view && vc.orgZoom != null && vc.orgCenter != null) {
                    vc.view.zoom = vc.orgZoom;
                    vc.view.centerPoint = vc.orgCenter;
                    app.redraw();
                }
            } catch (_) { }
            vc.panX = 0; vc.panY = 0;
        };

        vc.buildUI = function (parent, opt) {
            opt = opt || {};
            var labelW = (typeof opt.labelWidth === "number") ? opt.labelWidth : 58;
            var sliderW = (typeof opt.sliderWidth === "number") ? opt.sliderWidth : 260;

            var zoomMin = (typeof opt.zoomMin === "number") ? opt.zoomMin : 10;
            var zoomMax = (typeof opt.zoomMax === "number") ? opt.zoomMax : 1600;

            // labels (prefer L())
            var labelZoom = vc.t("zoomLabel", "ズーム");
            var labelLR = vc.t("lr", "左右");
            var labelUD = vc.t("ud", "上下");

            // initial zoom pct from current view
            var initZoomPct = 100;
            try { if (vc.orgZoom != null) initZoomPct = Math.round(Number(vc.orgZoom) * 100); } catch (_) { }
            if (!initZoomPct || initZoomPct < zoomMin) initZoomPct = 100;

            // pan ranges (UI bounds decided at build time)
            var pr = vc.refreshPanRange();

            function addRow(lblText) {
                var g = parent.add("group");
                g.orientation = "row";
                g.alignChildren = ["left", "center"];
                var st = g.add("statictext", undefined, lblText);
                try { st.preferredSize.width = labelW; } catch (_) { }
                return g;
            }

            // Zoom row
            var gZ = addRow(labelZoom);
            vc.sldZoom = gZ.add("slider", undefined, initZoomPct, zoomMin, zoomMax);
            try { vc.sldZoom.preferredSize.width = sliderW; } catch (_) { }
            var stZ = { raw: null, eff: null };
            vc.sldZoom.onChanging = function () {
                applySliderAltFine(this, stZ, function (v) { vc.setZoomPct(v, zoomMin, zoomMax); });
            };

            // Pan X row
            var gX = addRow(labelLR);
            vc.sldPanX = gX.add("slider", undefined, 0, -pr.xMax, pr.xMax);
            try { vc.sldPanX.preferredSize.width = sliderW; } catch (_) { }
            var stX = { raw: null, eff: null };
            vc.sldPanX.onChanging = function () {
                applySliderAltFine(this, stX, function (v) { vc.setPanX(v); });
            };

            // Pan Y row
            var gY = addRow(labelUD);
            vc.sldPanY = gY.add("slider", undefined, 0, -pr.yMax, pr.yMax);
            try { vc.sldPanY.preferredSize.width = sliderW; } catch (_) { }
            var stY = { raw: null, eff: null };
            vc.sldPanY.onChanging = function () {
                applySliderAltFine(this, stY, function (v) { vc.setPanY(v); });
            };

            return vc;
        };

        return vc;
    }

    return { create: create };
})();


(function () {
    // --- 準備とチェック ---
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    var sel = doc.selection;

    // ViewControl instance (Zoom/Pan)
    var viewCtl = null;
    try { viewCtl = ViewControl.create(doc, L); } catch (_) { viewCtl = null; }

    // パスアイテムのみを抽出してリスト化
    var targetItems = [];
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "PathItem") {
            targetItems.push(sel[i]);
        }
    }

    // 選択した長方形の初期状態を統一
    // ・塗り：なし
    // ・線：黒、1pt
    for (var ti = 0; ti < targetItems.length; ti++) {
        try {
            var item = targetItems[ti];
            item.filled = false;
            item.stroked = true;
            item.strokeWidth = 1;

            var black = new CMYKColor();
            black.cyan = 0;
            black.magenta = 0;
            black.yellow = 0;
            black.black = 100;
            item.strokeColor = black;
        } catch (_) { }
    }

    // 選択がない場合は、現在のアートボードを基準にする
    var _usingArtboardBase = false;
    var _artboardBaseRect = null;

    // 裁ち落とし（アートボード基準のみ） / Bleed (artboard only)
    var _bleedEnabled = false;
    var BLEED_MM = 3;

    function cleanupArtboardBaseRect() {
        if (_usingArtboardBase && _artboardBaseRect) {
            // targetItems から外す（remove後に触ると Error 45 になるため）
            try {
                for (var i = targetItems.length - 1; i >= 0; i--) {
                    if (targetItems[i] === _artboardBaseRect) {
                        targetItems.splice(i, 1);
                    }
                }
            } catch (_) { }

            try { _artboardBaseRect.remove(); } catch (_) { }
            _artboardBaseRect = null;
        }
    }

    // アートボード基準の一時矩形を「マージン」分だけ内側に作り直す
    function rebuildArtboardBaseRect(mTopPt, mRightPt, mBottomPt, mLeftPt) {
        if (!_usingArtboardBase) return;
        try {
            var abIndex = doc.artboards.getActiveArtboardIndex();
            var abRect = doc.artboards[abIndex].artboardRect; // [L, T, R, B]
            var L = abRect[0], T = abRect[1], R = abRect[2], B = abRect[3];

            // ※裁ち落とし（bleed）はここでは適用しない（フレームのみで適用）
            var mt = (mTopPt && mTopPt > 0) ? mTopPt : 0;
            var mr = (mRightPt && mRightPt > 0) ? mRightPt : 0;
            var mb = (mBottomPt && mBottomPt > 0) ? mBottomPt : 0;
            var ml = (mLeftPt && mLeftPt > 0) ? mLeftPt : 0;

            var iL = L + ml;
            var iT = T - mt;
            var iR = R - mr;
            var iB = B + mb;

            var w = iR - iL;
            var h = iT - iB;
            if (!(w > 0) || !(h > 0)) {
                // マージンが大きすぎる場合は 0 扱い
                iL = L; iT = T; iR = R; iB = B;
                w = iR - iL;
                h = iT - iB;
            }

            // 既存の一時矩形があれば安全に差し替える（targetItemsからも外す）
            if (_artboardBaseRect) {
                try {
                    for (var i = targetItems.length - 1; i >= 0; i--) {
                        if (targetItems[i] === _artboardBaseRect) targetItems.splice(i, 1);
                    }
                } catch (_) { }
                try { _artboardBaseRect.remove(); } catch (_) { }
                _artboardBaseRect = null;
            }

            _artboardBaseRect = doc.activeLayer.pathItems.rectangle(iT, iL, w, h);
            _artboardBaseRect.stroked = true;
            _artboardBaseRect.filled = false;
            // アートボード基準の外枠罫線設定：1pt / 黒
            try {
                var k = new CMYKColor();
                k.cyan = 0; k.magenta = 0; k.yellow = 0; k.black = 100;
                _artboardBaseRect.strokeColor = k;
            } catch (_) { }
            try { _artboardBaseRect.strokeWidth = 1; } catch (_) { }
            // 見た目を極力変えない（ガイドにはしない。最終的に削除する）

            targetItems.push(_artboardBaseRect);
        } catch (_) { }
    }

    if (targetItems.length === 0) {
        try {
            _usingArtboardBase = true;
            // ※実体の基準矩形は PreviewHistory.start() 後に作成する（履歴を残さないため）
        } catch (_) {
            alert(L("alertSelectPath"));
            return;
        }
    }

    // --- 互換性：StrokeCap が未定義の環境対策 ---
    // 一部環境で StrokeCap が参照できない場合があるため、最低限の定数を用意
    if (typeof StrokeCap === "undefined") {
        StrokeCap = {
            BUTTENDCAP: 0,
            ROUNDENDCAP: 1,
            PROJECTINGENDCAP: 2
        };
    }

    // --- 単位（rulerType）ユーティリティ：先にマップだけ初期化（UI作成で参照されるため） ---
    // ※関数本体は後ろにあっても宣言はホイストされるが、マップ値は先に必要
    var _unitMap = {
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

    // inner box オフセットの初期値： (幅 + 高さ) / 40
    // 表示単位は rulerType に合わせる
    function calcDefaultInnerOffset() {
        try {
            if (!targetItems || targetItems.length === 0) return 0;
            var it = targetItems[0];
            var b = it.geometricBounds; // [L, T, R, B]
            var wPt = b[2] - b[0];
            var hPt = b[1] - b[3];
            if (!(wPt > 0) || !(hPt > 0)) return 0;

            var offsetPt = (wPt + hPt) / 40;
            var factor = getCurrentRulerPtFactor(); // unit -> pt
            if (!factor || factor === 0) factor = 1;

            var v = offsetPt / factor; // rulerType unit
            if (v < 0) v = 0;
            // 10単位に丸め（例：150.3→150, 152→150, 155→160）
            v = Math.round(v / 10) * 10;
            return v;
        } catch (_) {
            return 0;
        }
    }

    // 生成した一時オブジェクト（プレビュー用）を管理する配列
    var tempPreviewItems = [];



    // アートボード基準の場合はここで基準矩形を作成（start後なので履歴は後で完全に戻せる）
    if (_usingArtboardBase) {
        rebuildArtboardBaseRect(0, 0, 0, 0);
    }

    // --- ダイアログ作成 ---
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    // --- ダイアログ位置・透明度設定 ---
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        try {
            dlg.opacity = opacityValue;
        } catch (_) { }
    }

    setDialogOpacity(win, dialogOpacity);
    shiftDialogPosition(win, offsetX, offsetY);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 20;
    win.margins = 16;

    // 3タブレイアウト
    // 左：マージン / フレーム
    // 中央：外側エリア / タイトルエリア
    // 右：内側エリア
    var tabPanel = win.add("tabbedpanel");
    tabPanel.alignChildren = ["fill", "top"];
    tabPanel.alignment = ["fill", "top"];
    tabPanel.margins = [5, 20, 0, 0];
    // tabbedpanel は内容量に応じて自動で高さが伸びないことがあるため、最低サイズを与える
    try {
        tabPanel.minimumSize = [300, 460];
        tabPanel.preferredSize = [300, 460];
    } catch (_) { }

    // Tabs
    var tabLeft = tabPanel.add("tab", undefined, L("panelMargin"));
    var tabMid = tabPanel.add("tab", undefined, L("panelOuter"));
    var tabRight = tabPanel.add("tab", undefined, L("panelInnerArea"));
    var tabDisplay = tabPanel.add("tab", undefined, L("panelDisplay"));

    // Left tab (Margin / Frame)
    tabLeft.orientation = "column";
    tabLeft.alignChildren = ["fill", "top"];
    tabLeft.spacing = 8;
    tabLeft.margins = 10;
    var leftCol = tabLeft.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["fill", "top"];
    leftCol.spacing = 8;

    // Middle tab (Outer / Title)
    tabMid.orientation = "column";
    tabMid.alignChildren = ["fill", "top"];
    tabMid.spacing = 8;
    tabMid.margins = 10;
    var midCol = tabMid.add("group");
    midCol.orientation = "column";
    midCol.alignChildren = ["fill", "top"];
    midCol.spacing = 8;

    // Right tab (Inner)
    tabRight.orientation = "column";
    tabRight.alignChildren = ["fill", "top"];
    tabRight.spacing = 8;
    tabRight.margins = 10;
    var rightCol = tabRight.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];
    rightCol.spacing = 8;

    // Display tab
    tabDisplay.orientation = "column";
    tabDisplay.alignChildren = ["fill", "top"];
    tabDisplay.spacing = 8;
    tabDisplay.margins = 10;

    var displayCol = tabDisplay.add("group");
    displayCol.orientation = "column";
    displayCol.alignChildren = ["fill", "top"];
    displayCol.spacing = 8;

    // =========================================
    // UI builders (split): MarginUI / OuterUI / InnerUI
    // =========================================

    // --- Margin UI handles ---
    var marginPanel;
    var editArtboardMarginTop, editArtboardMarginBottom, editArtboardMarginLeft, editArtboardMarginRight;
    var chkArtboardMarginLink;
    var _syncingArtboardMargins = false;
    var applyArtboardMarginLinkState;

    // --- Outer UI handles ---
    var outerPanel;
    var chkKeepOuter;
    var chkEnableLen;
    var inputGroup;
    var editVal;
    var stUnit;
    var chkOuterRound;
    var editOuterRound;
    var stOuterRoundUnit;
    var applyEdgeScaleEnabledState;
    var applyOuterLinePanelEnabledState;
    var applyOuterRoundEnabledState;
    var capPanel;
    var rbCapButt, rbCapRound, rbCapProject;

    // --- Title band variable handles ---
    var chkTitleEdgeScale, editTitleEdgeScale, stTitleEdgeScaleUnit;

    // --- Inner UI handles ---
    var innerPanel;
    var editInnerOffsetTop, editInnerOffsetBottom, editInnerOffsetLeft, editInnerOffsetRight;
    var chkInnerOffsetLink;
    var _syncingInnerOffsets = false;
    var applyInnerOffsetLinkState;

    var editInnerColumns, editInnerRows;
    var editColGutter, editRowGutter;
    var chkRowFill, chkRowDivider;
    var innerCapPanel;
    var rbInnerLineSolid, rbInnerLineDash, rbInnerLineDotDash;

    // Builder: MarginUI
    function buildMarginUI(parent) {
        // マージン（アートボード基準のときだけ有効）
        marginPanel = parent.add("panel", undefined, L("panelMargin"));
        marginPanel.orientation = "column";
        marginPanel.alignChildren = ["fill", "top"];
        marginPanel.margins = [15, 20, 15, 10];
        marginPanel.spacing = 10;

        // マージン入力：3段組（上 / 左＋連動＋右 / 下）
        var marginUnitLabel = getCurrentRulerUnitLabel();

        // 1段目：上（中央寄せ）
        var marginRowTop = marginPanel.add("group");
        marginRowTop.orientation = "row";
        marginRowTop.alignChildren = ["center", "center"];
        marginRowTop.alignment = ["fill", "top"];

        var marginTopGroup = marginRowTop.add("group");
        marginTopGroup.orientation = "row";
        marginTopGroup.alignChildren = ["left", "center"];
        marginTopGroup.add("statictext", undefined, L("offsetTop"));
        editArtboardMarginTop = marginTopGroup.add("edittext", undefined, "15");
        editArtboardMarginTop.characters = 4;
        changeValueByArrowKey(editArtboardMarginTop, false);
        marginTopGroup.add("statictext", undefined, marginUnitLabel);

        // 2段目：左 ＋ 連動（中央）＋ 右
        var marginRowMid = marginPanel.add("group");
        marginRowMid.orientation = "row";
        marginRowMid.alignChildren = ["center", "center"];
        marginRowMid.alignment = ["fill", "top"];
        marginRowMid.spacing = 12;

        var marginLeftGroup = marginRowMid.add("group");
        marginLeftGroup.orientation = "row";
        marginLeftGroup.alignChildren = ["left", "center"];
        marginLeftGroup.add("statictext", undefined, L("offsetLeft"));
        editArtboardMarginLeft = marginLeftGroup.add("edittext", undefined, "15");
        editArtboardMarginLeft.characters = 4;
        changeValueByArrowKey(editArtboardMarginLeft, false);

        chkArtboardMarginLink = marginRowMid.add("checkbox", undefined, L("chkLink"));
        chkArtboardMarginLink.value = true;

        var marginRightGroup = marginRowMid.add("group");
        marginRightGroup.orientation = "row";
        marginRightGroup.alignChildren = ["left", "center"];
        marginRightGroup.add("statictext", undefined, L("offsetRight"));
        editArtboardMarginRight = marginRightGroup.add("edittext", undefined, "15");
        editArtboardMarginRight.characters = 4;
        changeValueByArrowKey(editArtboardMarginRight, false);

        // 3段目：下（中央寄せ）
        var marginRowBottom = marginPanel.add("group");
        marginRowBottom.orientation = "row";
        marginRowBottom.alignChildren = ["center", "center"];
        marginRowBottom.alignment = ["fill", "top"];

        var marginBottomGroup = marginRowBottom.add("group");
        marginBottomGroup.orientation = "row";
        marginBottomGroup.alignChildren = ["left", "center"];
        marginBottomGroup.add("statictext", undefined, L("offsetBottom"));
        editArtboardMarginBottom = marginBottomGroup.add("edittext", undefined, "15");
        editArtboardMarginBottom.characters = 4;
        changeValueByArrowKey(editArtboardMarginBottom, false);
        marginBottomGroup.add("statictext", undefined, marginUnitLabel);

        // 連動（無限ループ防止は外側の _syncingArtboardMargins を使用）
        applyArtboardMarginLinkState = function () {
            var linked = !!chkArtboardMarginLink.value;
            // 連動ON: 下/左/右はディム表示（操作不可）
            try { marginBottomGroup.enabled = !linked; } catch (_) { }
            try { marginLeftGroup.enabled = !linked; } catch (_) { }
            try { marginRightGroup.enabled = !linked; } catch (_) { }

            if (linked) {
                var v = editArtboardMarginTop.text;
                _syncingArtboardMargins = true;
                try { editArtboardMarginBottom.text = v; } catch (_) { }
                try { editArtboardMarginLeft.text = v; } catch (_) { }
                try { editArtboardMarginRight.text = v; } catch (_) { }
                _syncingArtboardMargins = false;
            }
        };

        // アートボードが対象のときだけアクティブ
        marginPanel.enabled = !!_usingArtboardBase;

        // 長方形スタート時はパネルを隠し、レイアウト上のスペースも潰す
        try {
            marginPanel.visible = !!_usingArtboardBase;
            if (!_usingArtboardBase) {
                marginPanel.maximumSize.height = 0;
                marginPanel.minimumSize.height = 0;
            } else {
                marginPanel.maximumSize.height = 10000;
                marginPanel.minimumSize.height = 0;
            }
        } catch (_) { }

        applyArtboardMarginLinkState();
    }

    // Builder: OuterUI
    function buildOuterUI(parent) {
        // 外枠パネル
        outerPanel = parent.add("panel", undefined, L("panelOuter"));
        outerPanel.orientation = "column";
        outerPanel.alignChildren = ["fill", "top"];
        outerPanel.margins = [15, 20, 15, 10];
        outerPanel.spacing = 10;

        // 外枠を残す
        chkKeepOuter = outerPanel.add("checkbox", undefined, L("chkKeepOuter"));
        chkKeepOuter.value = true;

        // 辺の長さ調整（1行）
        var lenRow = outerPanel.add("group");
        lenRow.orientation = "row";
        lenRow.alignChildren = ["left", "center"];

        chkEnableLen = lenRow.add("checkbox", undefined, L("chkEdgeScale"));
        chkEnableLen.value = true;

        // 値入力（チェックがOFFのときだけディム表示）
        inputGroup = lenRow.add("group");
        inputGroup.orientation = "row";
        inputGroup.alignChildren = ["left", "center"];

        editVal = inputGroup.add("edittext", undefined, "-5");
        editVal.characters = 4;
        stUnit = inputGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
        editVal.active = true;
        changeValueByArrowKey(editVal, true);

        // 辺の伸縮の有効/無効
        applyEdgeScaleEnabledState = function () {
            try {
                var enabled = !!chkKeepOuter.value;
                chkEnableLen.enabled = enabled;
                inputGroup.enabled = (enabled && chkEnableLen.value);
            } catch (_) { }
        };
        applyEdgeScaleEnabledState();

        // 外側エリア：角丸（UIのみ。ロジックは後で適用）
        var outerRoundRow = outerPanel.add("group");
        outerRoundRow.orientation = "row";
        outerRoundRow.alignChildren = ["left", "center"];

        chkOuterRound = outerRoundRow.add("checkbox", undefined, "角丸");
        chkOuterRound.value = false;

        editOuterRound = outerRoundRow.add("edittext", undefined, "0");
        editOuterRound.characters = 4;
        changeValueByArrowKey(editOuterRound, false);

        stOuterRoundUnit = outerRoundRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        applyOuterRoundEnabledState = function () {
            try {
                // 「辺の伸縮」がONのときは角丸は使えない（ディム表示）
                var usable = (!!chkKeepOuter.value && !(chkEnableLen && chkEnableLen.value));

                chkOuterRound.enabled = usable;

                if (!usable) {
                    chkOuterRound.value = false;
                    editOuterRound.text = "0";
                }

                editOuterRound.enabled = (usable && !!chkOuterRound.value);
                if (usable && !chkOuterRound.value) editOuterRound.text = "0";
            } catch (_) { }
        };
        applyOuterRoundEnabledState();

        chkOuterRound.onClick = function () {
            try {
                if (chkOuterRound.value) {
                    var v = parseFloat(editOuterRound.text);
                    if (isNaN(v) || v === 0) {
                        editOuterRound.text = "2";
                    }
                }
            } catch (_) { }

            applyOuterRoundEnabledState();
            // 外側エリアの角丸が0以外なら、タイトルエリアの［塗り］をOFF
            try {
                var rv = parseFloat(editOuterRound.text);
                if (chkOuterRound.value && !isNaN(rv) && rv !== 0) {
                    if (typeof chkTitleFill !== "undefined" && chkTitleFill) {
                        chkTitleFill.value = false;
                    }
                }
            } catch (_) { }
            if (chkPreview && chkPreview.value) updatePreview(false);
        };

        editOuterRound.onChanging = function () {
            try {
                var rv = parseFloat(editOuterRound.text);
                if (chkOuterRound.value && !isNaN(rv) && rv !== 0) {
                    if (typeof chkTitleFill !== "undefined" && chkTitleFill) {
                        chkTitleFill.value = false;
                    }
                }
            } catch (_) { }
            if (chkPreview && chkPreview.value) updatePreview(false);
        };

        // 線端（ストロークキャップ）
        capPanel = outerPanel.add("panel", undefined, L("panelCap"));
        capPanel.orientation = "row";
        capPanel.alignChildren = ["left", "center"];
        capPanel.margins = [15, 20, 15, 10];

        rbCapButt = capPanel.add("radiobutton", undefined, L("capNone"));
        rbCapRound = capPanel.add("radiobutton", undefined, L("capRound"));
        rbCapProject = capPanel.add("radiobutton", undefined, L("capProject"));

        // 初期値：選択オブジェクトの線端を優先
        (function initCapUI() {
            var cap = null;
            try {
                if (targetItems.length > 0 && targetItems[0].stroked) cap = targetItems[0].strokeCap;
            } catch (_) { }

            if (cap === StrokeCap.ROUNDENDCAP) {
                rbCapRound.value = true;
            } else if (cap === StrokeCap.PROJECTINGENDCAP) {
                rbCapProject.value = true;
            } else {
                rbCapButt.value = true;
            }
        })();

        // 外側エリア：線パネルの有効/無効
        applyOuterLinePanelEnabledState = function () {
            try {
                var lenVal = getEffectiveLenValue();
                capPanel.enabled = (!!chkKeepOuter.value && !!chkEnableLen.value && lenVal !== 0);
            } catch (_) { }
        };
        applyOuterLinePanelEnabledState();
    }

    // Builder: InnerUI
    function buildInnerUI(parent) {
        // inner box パネル
        innerPanel = parent.add("panel", undefined, L("panelInnerArea"));
        innerPanel.orientation = "column";
        innerPanel.alignChildren = ["fill", "top"];
        innerPanel.margins = [15, 20, 15, 10];

        // オフセット入力：3段組（上 / 左＋連動＋右 / 下）
        var innerOffsetPanel = innerPanel.add("panel", undefined, L("panelOffset") + "（" + getCurrentRulerUnitLabel() + "）");
        innerOffsetPanel.orientation = "column";
        innerOffsetPanel.alignChildren = ["fill", "top"];
        innerOffsetPanel.margins = [15, 20, 15, 10];

        // 1段目：上（中央寄せ）
        var innerOffsetRowTop = innerOffsetPanel.add("group");
        innerOffsetRowTop.orientation = "row";
        innerOffsetRowTop.alignChildren = ["center", "center"];
        innerOffsetRowTop.alignment = ["fill", "top"];

        var innerOffsetTopGroup = innerOffsetRowTop.add("group");
        innerOffsetTopGroup.orientation = "row";
        innerOffsetTopGroup.alignChildren = ["left", "center"];
        innerOffsetTopGroup.add("statictext", undefined, L("offsetTop"));
        editInnerOffsetTop = innerOffsetTopGroup.add("edittext", undefined, String(calcDefaultInnerOffset()));
        editInnerOffsetTop.characters = 3;
        changeValueByArrowKey(editInnerOffsetTop, false);

        // 2段目：左 ＋ 連動（中央）＋ 右
        var innerOffsetRowMid = innerOffsetPanel.add("group");
        innerOffsetRowMid.orientation = "row";
        innerOffsetRowMid.alignChildren = ["center", "center"];
        innerOffsetRowMid.alignment = ["fill", "top"];
        innerOffsetRowMid.spacing = 12;

        var innerOffsetLeftGroup = innerOffsetRowMid.add("group");
        innerOffsetLeftGroup.orientation = "row";
        innerOffsetLeftGroup.alignChildren = ["left", "center"];
        innerOffsetLeftGroup.add("statictext", undefined, L("offsetLeft"));
        editInnerOffsetLeft = innerOffsetLeftGroup.add("edittext", undefined, String(calcDefaultInnerOffset()));
        editInnerOffsetLeft.characters = 3;
        changeValueByArrowKey(editInnerOffsetLeft, false);

        chkInnerOffsetLink = innerOffsetRowMid.add("checkbox", undefined, L("chkLink"));
        chkInnerOffsetLink.value = true;

        var innerOffsetRightGroup = innerOffsetRowMid.add("group");
        innerOffsetRightGroup.orientation = "row";
        innerOffsetRightGroup.alignChildren = ["left", "center"];
        innerOffsetRightGroup.add("statictext", undefined, L("offsetRight"));
        editInnerOffsetRight = innerOffsetRightGroup.add("edittext", undefined, String(calcDefaultInnerOffset()));
        editInnerOffsetRight.characters = 3;
        changeValueByArrowKey(editInnerOffsetRight, false);

        // 3段目：下（中央寄せ）
        var innerOffsetRowBottom = innerOffsetPanel.add("group");
        innerOffsetRowBottom.orientation = "row";
        innerOffsetRowBottom.alignChildren = ["center", "center"];
        innerOffsetRowBottom.alignment = ["fill", "top"];

        var innerOffsetBottomGroup = innerOffsetRowBottom.add("group");
        innerOffsetBottomGroup.orientation = "row";
        innerOffsetBottomGroup.alignChildren = ["left", "center"];
        innerOffsetBottomGroup.add("statictext", undefined, L("offsetBottom"));
        editInnerOffsetBottom = innerOffsetBottomGroup.add("edittext", undefined, String(calcDefaultInnerOffset()));
        editInnerOffsetBottom.characters = 3;
        changeValueByArrowKey(editInnerOffsetBottom, false);

        applyInnerOffsetLinkState = function () {
            var linked = !!chkInnerOffsetLink.value;
            // 連動ON: 下/左/右はディム表示（操作不可）
            try { innerOffsetBottomGroup.enabled = !linked; } catch (_) { }
            try { innerOffsetLeftGroup.enabled = !linked; } catch (_) { }
            try { innerOffsetRightGroup.enabled = !linked; } catch (_) { }

            if (linked) {
                var v = editInnerOffsetTop.text;
                _syncingInnerOffsets = true;
                try { editInnerOffsetBottom.text = v; } catch (_) { }
                try { editInnerOffsetLeft.text = v; } catch (_) { }
                try { editInnerOffsetRight.text = v; } catch (_) { }
                _syncingInnerOffsets = false;
            }
        };
        applyInnerOffsetLinkState();

        // 列・行
        var innerGridWrap = innerPanel.add("group");
        innerGridWrap.orientation = "row";
        innerGridWrap.alignChildren = ["left", "top"];
        innerGridWrap.alignment = ["fill", "top"];

        var innerGridGroup = innerGridWrap.add("group");
        innerGridGroup.orientation = "column";
        innerGridGroup.alignChildren = ["left", "top"];
        innerGridGroup.alignment = ["left", "top"];
        innerGridGroup.spacing = 12;

        // 列
        var colPanel = innerGridGroup.add("panel", undefined, L("panelColumns"));
        colPanel.orientation = "column";
        colPanel.alignChildren = ["fill", "top"];
        colPanel.margins = [15, 20, 15, 10];
        colPanel.spacing = 8;

        var gridColRow = colPanel.add("group");
        gridColRow.orientation = "row";
        gridColRow.alignChildren = ["left", "center"];

        gridColRow.add("statictext", undefined, L("colCount"));
        editInnerColumns = gridColRow.add("edittext", undefined, "1");
        editInnerColumns.characters = 3;
        changeValueByArrowKey(editInnerColumns, false);

        gridColRow.add("statictext", undefined, L("spacing"));
        editColGutter = gridColRow.add("edittext", undefined, "0");
        editColGutter.characters = 4;
        changeValueByArrowKey(editColGutter, false);
        gridColRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        try { editColGutter.enabled = (parseInt(editInnerColumns.text, 10) > 1); } catch (_) { }

        // 行
        var rowPanel = innerGridGroup.add("panel", undefined, L("panelRows"));
        rowPanel.orientation = "column";
        rowPanel.alignChildren = ["fill", "top"];
        rowPanel.margins = [15, 20, 15, 10];
        rowPanel.spacing = 8;

        var gridRowRow = rowPanel.add("group");
        gridRowRow.orientation = "row";
        gridRowRow.alignChildren = ["left", "center"];

        gridRowRow.add("statictext", undefined, L("rowCount"));
        editInnerRows = gridRowRow.add("edittext", undefined, "1");
        editInnerRows.characters = 3;
        changeValueByArrowKey(editInnerRows, false);

        gridRowRow.add("statictext", undefined, L("spacing"));
        editRowGutter = gridRowRow.add("edittext", undefined, "0");
        editRowGutter.characters = 4;
        changeValueByArrowKey(editRowGutter, false);
        gridRowRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        // 行オプション
        var gridRowOptsWrap = innerGridGroup.add("group");
        gridRowOptsWrap.orientation = "row";
        gridRowOptsWrap.alignChildren = ["center", "center"];
        gridRowOptsWrap.alignment = ["fill", "top"];

        var gridRowOpts = gridRowOptsWrap.add("group");
        gridRowOpts.orientation = "row";
        gridRowOpts.alignChildren = ["left", "center"];
        gridRowOpts.alignment = ["center", "center"];

        chkRowFill = gridRowOpts.add("checkbox", undefined, L("chkFill"));
        chkRowFill.value = false;

        chkRowDivider = gridRowOpts.add("checkbox", undefined, L("chkDivider"));
        chkRowDivider.value = false;

        // 内側の線種
        innerCapPanel = innerPanel.add("panel", undefined, L("panelLineType"));
        innerCapPanel.orientation = "row";
        innerCapPanel.alignChildren = ["left", "center"];
        innerCapPanel.margins = [15, 20, 15, 10];

        rbInnerLineSolid = innerCapPanel.add("radiobutton", undefined, L("lineSolid"));
        rbInnerLineDash = innerCapPanel.add("radiobutton", undefined, L("lineDash"));
        rbInnerLineDotDash = innerCapPanel.add("radiobutton", undefined, L("lineDotDash"));
        rbInnerLineSolid.value = true;
    }

    buildMarginUI(leftCol);

    // 長方形スタート時（アートボード対象でない場合）は左タブ全体を非表示
    try {
        if (!_usingArtboardBase) {
            tabLeft.visible = false;
            tabLeft.enabled = false;
            tabPanel.selection = tabMid;
        } else {
            tabLeft.visible = true;
            tabLeft.enabled = true;
        }
    } catch (_) { }

    buildOuterUI(midCol);

    // タイトルエリア
    var titlePanel = midCol.add("panel", undefined, L("panelTitleBand"));
    titlePanel.orientation = "column";
    titlePanel.alignChildren = ["fill", "top"];
    titlePanel.margins = [15, 20, 15, 10];
    titlePanel.spacing = 10;

    // 有効 ＋ 幅／高さ（1行）
    var titleEnableRow = titlePanel.add("group");
    titleEnableRow.orientation = "row";
    titleEnableRow.alignChildren = ["left", "center"];

    var chkTitleEnable = titleEnableRow.add("checkbox", undefined, "");
    chkTitleEnable.value = false; // デフォルトOFF

    // 幅／高さ（有効チェックの右側）
    var titleSizeGroup = titleEnableRow.add("group");
    titleSizeGroup.orientation = "row";
    titleSizeGroup.alignChildren = ["left", "center"];

    // titleSizeGroup.add("statictext", undefined, L("titleSize"));
    var editTitleSize = titleSizeGroup.add("edittext", undefined, "0");
    editTitleSize.characters = 4;
    titleSizeGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
    changeValueByArrowKey(editTitleSize, false);

    // 位置（上/下/左/右）
    var titlePosGroup = titlePanel.add("group");
    titlePosGroup.orientation = "row";
    titlePosGroup.alignChildren = ["left", "center"];

    var rbTitleTop = titlePosGroup.add("radiobutton", undefined, L("titleTop"));
    var rbTitleBottom = titlePosGroup.add("radiobutton", undefined, L("titleBottom"));
    var rbTitleLeft = titlePosGroup.add("radiobutton", undefined, L("titleLeft"));
    var rbTitleRight = titlePosGroup.add("radiobutton", undefined, L("titleRight"));

    // デフォルト：上
    rbTitleTop.value = true;

    // 塗り／線
    var titleOptionGroup = titlePanel.add("group");
    titleOptionGroup.orientation = "row";
    titleOptionGroup.alignChildren = ["left", "center"];

    var chkTitleFill = titleOptionGroup.add("checkbox", undefined, L("chkFill"));
    chkTitleFill.value = false;

    var chkTitleLine = titleOptionGroup.add("checkbox", undefined, L("panelLine"));
    chkTitleLine.value = true;

    // タイトルエリア：辺の伸縮（UIのみ。ロジックは後で適用）
    var titleEdgeScaleRow = titlePanel.add("group");
    titleEdgeScaleRow.orientation = "row";
    titleEdgeScaleRow.alignChildren = ["left", "center"];

    chkTitleEdgeScale = titleEdgeScaleRow.add("checkbox", undefined, L("chkEdgeScale"));
    chkTitleEdgeScale.value = false;

    editTitleEdgeScale = titleEdgeScaleRow.add("edittext", undefined, "0");
    editTitleEdgeScale.characters = 4;
    changeValueByArrowKey(editTitleEdgeScale, true);

    stTitleEdgeScaleUnit = titleEdgeScaleRow.add("statictext", undefined, getCurrentRulerUnitLabel());

    function applyTitleEdgeScaleEnabledState() {
        try {
            editTitleEdgeScale.enabled = !!chkTitleEdgeScale.value;
            if (!chkTitleEdgeScale.value) editTitleEdgeScale.text = "0";
        } catch (_) { }
    }
    applyTitleEdgeScaleEnabledState();

    chkTitleEdgeScale.onClick = function () {
        applyTitleEdgeScaleEnabledState();
        if (chkPreview && chkPreview.value) updatePreview(false);
    };

    editTitleEdgeScale.onChanging = function () {
        if (!chkTitleEdgeScale.value) return;
        if (chkPreview && chkPreview.value) updatePreview(false);
    };

    // 0→>0 の瞬間だけ自動ON（ユーザーは後からOFF可）
    var _prevTitleHasSize = (function () {
        var _s = parseFloat(editTitleSize.text);
        return (!isNaN(_s) && _s > 0);
    })();

    function applyTitleAreaEnabledState() {
        var areaEnabled = !!chkTitleEnable.value;

        // 無効なら 0 扱い（=タイトル生成なし）
        if (!areaEnabled) {
            if (editTitleSize.text !== "0") editTitleSize.text = "0";
            chkTitleFill.value = false;
            chkTitleLine.value = false;
        }

        // 幅／高さ・塗り/線は有効チェックに従う
        titleSizeGroup.enabled = areaEnabled;
        titleOptionGroup.enabled = areaEnabled;

        // ［辺の伸縮］は「タイトル有効」かつ「線ON」のときのみ操作可能
        var edgeRowEnabled = (areaEnabled && !!chkTitleLine.value);
        try { titleEdgeScaleRow.enabled = edgeRowEnabled; } catch (_) { }
        if (!edgeRowEnabled) {
            try { chkTitleEdgeScale.value = false; } catch (_) { }
            try { editTitleEdgeScale.text = "0"; } catch (_) { }
            try { applyTitleEdgeScaleEnabledState(); } catch (_) { }
        }

        // 位置は「有効」かつ「サイズ>0」のときのみ
        var sz = parseFloat(editTitleSize.text);
        var hasSize = (!isNaN(sz) && sz > 0);
        titlePosGroup.enabled = (areaEnabled && hasSize);

        // 塗り/線の有効（サイズ>0）
        chkTitleFill.enabled = (areaEnabled && hasSize);
        if (!(areaEnabled && hasSize)) chkTitleFill.value = false;

        // 線：サイズ>0 の間だけ有効。0→>0 で自動ON
        if (!_prevTitleHasSize && hasSize && areaEnabled) {
            chkTitleLine.value = true;
        }
        chkTitleLine.enabled = (areaEnabled && hasSize);
        if (!(areaEnabled && hasSize)) chkTitleLine.value = false;
        _prevTitleHasSize = (hasSize && areaEnabled);
    }


    // 初期反映
    applyTitleAreaEnabledState();

    // 位置変更でプレビュー更新
    function onTitlePosChanged() {
        if (chkPreview && chkPreview.value) updatePreview(false);
    }
    rbTitleTop.onClick = onTitlePosChanged;
    rbTitleBottom.onClick = onTitlePosChanged;
    rbTitleLeft.onClick = onTitlePosChanged;
    rbTitleRight.onClick = onTitlePosChanged;

    // サイズ変更
    editTitleSize.onChanging = function () {
        applyTitleAreaEnabledState();
        if (chkPreview.value) updatePreview();
    };

    // 有効切替
    chkTitleEnable.onClick = function () {
        // ONにしたとき、現在値が0ならデフォルトで「10」（現在の単位系）を入れる
        try {
            if (chkTitleEnable.value) {
                var v = parseFloat(editTitleSize.text);
                if (isNaN(v) || v === 0) {
                    editTitleSize.text = "10";
                }
            }
        } catch (_) { }

        applyTitleAreaEnabledState();
        if (chkPreview.value) updatePreview();
    };



    // フレーム
    var framePanel = leftCol.add("panel", undefined, L("panelFrame"));
    framePanel.orientation = "column";
    framePanel.alignChildren = ["fill", "top"];
    framePanel.margins = [15, 20, 15, 10];
    framePanel.spacing = 10;

    var frameRow = framePanel.add("group");
    frameRow.orientation = "row";
    frameRow.alignChildren = ["left", "center"];

    // フレーム：有効
    var chkFrameEnable = frameRow.add("checkbox", undefined, "");
    chkFrameEnable.value = false; // デフォルトOFF
    // フレーム有効チェック切替時の処理
    chkFrameEnable.onClick = function () {
        try {
            var enabled = !!chkFrameEnable.value;

            if (enabled) {
                // 幅が0ならデフォルトで「10」
                var v = parseFloat(editFrameWidth.text);
                if (isNaN(v) || v === 0) {
                    editFrameWidth.text = "10";
                }

                // 裁ち落としも自動でON（表示されている場合）
                try {
                    if (chkBleed && chkBleed.visible !== false) {
                        chkBleed.value = true;
                    }
                } catch (_) { }
            }

            // ON/OFF に応じて関連UIを有効/無効化
            applyFrameEnabledState();
        } catch (_) { }
        if (chkPreview.value) updatePreview();
    };
    // frameRow.add("statictext", undefined, L("frameWidth"));
    var editFrameWidth = frameRow.add("edittext", undefined, "0");
    editFrameWidth.characters = 4;
    frameRow.add("statictext", undefined, getCurrentRulerUnitLabel());
    changeValueByArrowKey(editFrameWidth, false);

    function applyFrameEnabledState() {
        try {
            // フレームはアートボード基準のみ有効（長方形選択で開始した場合はパネル全体をディム）
            if (!_usingArtboardBase) {
                try {
                    framePanel.visible = false;
                    framePanel.maximumSize.height = 0;
                    framePanel.minimumSize.height = 0;
                } catch (_) { }

                try { chkFrameEnable.value = false; } catch (_) { }
                try { editFrameWidth.text = "0"; } catch (_) { }

                try { chkBleed.value = false; } catch (_) { }
                try { chkFrameRound.value = false; } catch (_) { }
                try { if (editFrameRound) editFrameRound.text = "0"; } catch (_) { }

                // 依存状態の更新
                try { editFrameWidth.enabled = false; } catch (_) { }
                try { if (chkBleed) chkBleed.enabled = false; } catch (_) { }
                try { if (chkFrameRound) chkFrameRound.enabled = false; } catch (_) { }
                try { if (editFrameRound) editFrameRound.enabled = false; } catch (_) { }
                return;
            } else {
                try {
                    framePanel.visible = true;
                    framePanel.maximumSize.height = 10000;
                    framePanel.minimumSize.height = 0;
                } catch (_) { }
            }

            var enabled = !!chkFrameEnable.value;

            // 幅フィールド
            editFrameWidth.enabled = enabled;
            if (!enabled) {
                // OFF時は0扱い
                editFrameWidth.text = "0";
            }

            // 裁ち落とし
            if (chkBleed) {
                var fw = parseFloat(editFrameWidth.text);
                var valid = (!isNaN(fw) && fw > 0);
                chkBleed.enabled = (enabled && _usingArtboardBase && valid);
                if (!enabled) chkBleed.value = false;
            }

            // 角丸
            if (chkFrameRound) {
                var fw2 = parseFloat(editFrameWidth.text);
                var valid2 = (!isNaN(fw2) && fw2 > 0);

                chkFrameRound.enabled = (enabled && valid2);
                if (!enabled) chkFrameRound.value = false;

                if (editFrameRound) {
                    editFrameRound.enabled = (enabled && valid2 && chkFrameRound.value);
                }
            }
        } catch (_) { }
    }


    // 裁ち落とし（アートボード基準のみ表示）
    var bleedRow = framePanel.add("group");
    bleedRow.orientation = "row";
    bleedRow.alignChildren = ["left", "center"];

    var chkBleed = bleedRow.add("checkbox", undefined, L("chkBleed"));
    chkBleed.value = false;
    // 常時表示（長方形スタート時は enabled で制御）
    chkBleed.visible = true;

    // 角丸（UIのみ）
    var frameRoundRow = framePanel.add("group");
    frameRoundRow.orientation = "row";
    frameRoundRow.alignChildren = ["left", "center"];

    var chkFrameRound = frameRoundRow.add("checkbox", undefined, L("chkFrameRound"));
    chkFrameRound.value = false;

    var editFrameRound = frameRoundRow.add("edittext", undefined, "0");
    editFrameRound.characters = 4;
    changeValueByArrowKey(editFrameRound, false);
    frameRoundRow.add("statictext", undefined, getCurrentRulerUnitLabel());

    // 初期状態：角丸OFFなら値はディム
    try { editFrameRound.enabled = chkFrameRound.value; } catch (_) { }

    // フレーム幅が0のときは裁ち落としを無効化
    try {
        var _fw0 = parseFloat(editFrameWidth.text);
        chkBleed.enabled = (_usingArtboardBase && !isNaN(_fw0) && _fw0 > 0);
        if (!chkBleed.enabled) chkBleed.value = false;
    } catch (_) { }

    // フレーム幅が0のときは角丸も無効化
    try {
        var _fw1 = parseFloat(editFrameWidth.text);
        var _enableRound = (!isNaN(_fw1) && _fw1 > 0);
        chkFrameRound.enabled = _enableRound;
        if (!_enableRound) {
            chkFrameRound.value = false;
            editFrameRound.enabled = false;
        }
    } catch (_) { }

    // 初期反映（関連UIの有効/無効を最終調整）
    applyFrameEnabledState();
    try {
        framePanel.visible = !!_usingArtboardBase;
        if (!_usingArtboardBase) {
            framePanel.maximumSize.height = 0;
            framePanel.minimumSize.height = 0;
        } else {
            framePanel.maximumSize.height = 10000;
            framePanel.minimumSize.height = 0;
        }
    } catch (_) { }


    buildInnerUI(rightCol);
    buildDisplayUI(displayCol);

    // Builder: DisplayUI
    function buildDisplayUI(parent) {
        var displayPanel = parent.add("panel", undefined, L("panelZoomPan"));
        displayPanel.orientation = "column";
        displayPanel.alignChildren = ["fill", "top"];
        displayPanel.margins = [15, 20, 15, 10];
        displayPanel.spacing = 10;

        var pnlView = displayPanel.add("group");
        pnlView.orientation = "column";
        pnlView.alignChildren = "left";
        pnlView.spacing = 8;

        if (viewCtl && typeof viewCtl.buildUI === "function") {
            viewCtl.buildUI(pnlView, { labelWidth: 58, sliderWidth: 260 });
        } else {
            pnlView.add("statictext", undefined, "(ViewControl unavailable)");
        }
    }

    // ［塗り］の手動操作を優先するためのフラグ（ガター変更での自動ONを抑制）
    var _rowFillManuallySet = false;


    // 列/行の分割が可能になった瞬間だけ「分割線」を自動ONにするためのフラグ
    var _prevGridSplittable = false;
    try {
        var _c0 = parseInt(editInnerColumns.text, 10);
        var _r0 = parseInt(editInnerRows.text, 10);
        if (isNaN(_c0) || _c0 < 1) _c0 = 1;
        if (isNaN(_r0) || _r0 < 1) _r0 = 1;
        _prevGridSplittable = (_c0 > 1 || _r0 > 1);
    } catch (_) { _prevGridSplittable = false; }

    function applyRowDividerEnabledState(colCount, rowCount, allowAutoOn) {
        var splittable = (colCount > 1 || rowCount > 1);
        try { chkRowDivider.enabled = splittable; } catch (_) { }

        if (!splittable) {
            // 分割できないなら分割線は不要
            try { chkRowDivider.value = false; } catch (_) { }
        } else if (allowAutoOn && !_prevGridSplittable) {
            // 1/1 から分割可能になった瞬間だけ自動ON
            try { chkRowDivider.value = true; } catch (_) { }
        }

        _prevGridSplittable = splittable;

        // 線パネルも連動
        try { innerCapPanel.enabled = (chkRowDivider.enabled && chkRowDivider.value); } catch (_) { }
    }

    // 初期状態：列/行が1/1なら分割線はディム（OFF）、分割可能ならON/OFFに従う
    try {
        var _cInit = parseInt(editInnerColumns.text, 10);
        var _rInit = parseInt(editInnerRows.text, 10);
        if (isNaN(_cInit) || _cInit < 1) _cInit = 1;
        if (isNaN(_rInit) || _rInit < 1) _rInit = 1;
        applyRowDividerEnabledState(_cInit, _rInit, false);
    } catch (_) {
        try { innerCapPanel.enabled = (chkRowDivider.enabled && chkRowDivider.value); } catch (_) { }
    }


    // 下段（左：プレビュー / 右：ボタン）
    var bottomRow = win.add("group");
    bottomRow.orientation = "row";
    bottomRow.alignChildren = ["left", "center"];
    bottomRow.alignment = ["fill", "center"]; // 行自体を横いっぱいに

    // プレビューチェックボックス（左寄せ）
    var chkPreview = bottomRow.add("checkbox", undefined, L("preview"));
    chkPreview.value = true; // 最初からプレビューON
    chkPreview.alignment = "left";

    // スペーサー（右側のボタンを押し出す）
    var spacer = bottomRow.add("statictext", undefined, "");
    spacer.alignment = ["fill", "center"];
    spacer.minimumSize.width = 0;
    spacer.maximumSize.width = 10000;

    // ボタンエリア（右寄せ）
    var btnGroup = bottomRow.add("group");
    btnGroup.alignment = ["right", "center"];
    var cancelBtn = btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });

    // Ensure view is restored when user cancels/closes the dialog
    try {
        cancelBtn.onClick = function () {
            try { if (viewCtl && typeof viewCtl.restore === "function") viewCtl.restore(); } catch (_) { }
            try { clearPreview(); } catch (_) { }
            try { win.close(0); } catch (_) { }
        };
    } catch (_) { }

    try {
        win.onClose = function () {
            try { if (viewCtl && typeof viewCtl.restore === "function") viewCtl.restore(); } catch (_) { }
            return true;
        };
    } catch (_) { }

    var okBtn = btnGroup.add("button", undefined, L("ok"), { name: "ok" });

    // Ensure view is restored when user cancels/closes the dialog
    try {
        cancelBtn.onClick = function () {
            try { if (viewCtl && typeof viewCtl.restore === "function") viewCtl.restore(); } catch (_) { }
            try { clearPreview(); } catch (_) { }
            try { win.close(0); } catch (_) { }
        };
    } catch (_) { }

    try {
        win.onClose = function () {
            try { if (viewCtl && typeof viewCtl.restore === "function") viewCtl.restore(); } catch (_) { }
            return true;
        };
    } catch (_) { }

    // --- Restore last UI state (session only) ---
    (function restoreUIState() {
        var st = __loadState();
        if (!st) return;
        function safeSet(fn) { try { fn(); } catch (_) { } }

        function getNested(obj, pathArr) {
            try {
                var cur = obj;
                for (var i = 0; i < pathArr.length; i++) {
                    if (!cur) return undefined;
                    cur = cur[pathArr[i]];
                }
                return cur;
            } catch (_) { return undefined; }
        }

        function firstDefined(v1, v2, v3) {
            return (typeof v1 !== "undefined") ? v1 : ((typeof v2 !== "undefined") ? v2 : v3);
        }

        safeSet(function () { if (typeof st.preview !== "undefined") chkPreview.value = !!st.preview; });

        // Margin (4-way). New nested form: st.margin.{top,bottom,left,right,link}
        // Fallback from old flat keys: st.marginTop etc, and older single st.margin
        safeSet(function () {
            var v = firstDefined(getNested(st, ["margin", "top"]), st.marginTop, st.margin);
            if (typeof v !== "undefined") editArtboardMarginTop.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["margin", "bottom"]), st.marginBottom, st.margin);
            if (typeof v !== "undefined") editArtboardMarginBottom.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["margin", "left"]), st.marginLeft, st.margin);
            if (typeof v !== "undefined") editArtboardMarginLeft.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["margin", "right"]), st.marginRight, st.margin);
            if (typeof v !== "undefined") editArtboardMarginRight.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["margin", "link"]), st.marginLink, undefined);
            if (typeof v !== "undefined") chkArtboardMarginLink.value = !!v;
        });
        safeSet(function () { applyArtboardMarginLinkState(); });

        // Outer
        safeSet(function () {
            var v = firstDefined(getNested(st, ["outer", "keepOuter"]), st.keepOuter, undefined);
            if (typeof v !== "undefined") chkKeepOuter.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["outer", "enableLen"]), st.enableLen, undefined);
            if (typeof v !== "undefined") chkEnableLen.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["outer", "lenVal"]), st.lenVal, undefined);
            if (typeof v !== "undefined") editVal.text = String(v);
        });
        safeSet(function () {
            var cap = firstDefined(getNested(st, ["outer", "cap"]), st.cap, undefined);
            if (cap === "round") rbCapRound.value = true;
            else if (cap === "project") rbCapProject.value = true;
            else if (typeof cap !== "undefined") rbCapButt.value = true;
        });

        safeSet(function () {
            var v = firstDefined(getNested(st, ["outer", "round", "enable"]), st.outerRoundEnable, undefined);
            if (typeof v !== "undefined") chkOuterRound.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["outer", "round", "val"]), st.outerRoundVal, undefined);
            if (typeof v !== "undefined") editOuterRound.text = String(v);
        });
        safeSet(function () { applyOuterRoundEnabledState(); });

        // Title
        safeSet(function () {
            var v = firstDefined(getNested(st, ["title", "enable"]), st.titleEnable, undefined);
            if (typeof v !== "undefined") chkTitleEnable.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["title", "size"]), st.titleSize, undefined);
            if (typeof v !== "undefined") editTitleSize.text = String(v);
        });
        safeSet(function () {
            var p = firstDefined(getNested(st, ["title", "pos"]), st.titlePos, undefined);
            if (p === "bottom") rbTitleBottom.value = true;
            else if (p === "left") rbTitleLeft.value = true;
            else if (p === "right") rbTitleRight.value = true;
            else if (typeof p !== "undefined") rbTitleTop.value = true;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["title", "fill"]), st.titleFill, undefined);
            if (typeof v !== "undefined") chkTitleFill.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["title", "line"]), st.titleLine, undefined);
            if (typeof v !== "undefined") chkTitleLine.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["title", "edgeScale", "enable"]), st.titleEdgeScaleEnable, undefined);
            if (typeof v !== "undefined") chkTitleEdgeScale.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["title", "edgeScale", "val"]), st.titleEdgeScaleVal, undefined);
            if (typeof v !== "undefined") editTitleEdgeScale.text = String(v);
        });
        safeSet(function () { applyTitleEdgeScaleEnabledState(); });

        // Frame
        safeSet(function () {
            var v = firstDefined(getNested(st, ["frame", "enable"]), st.frameEnable, undefined);
            if (typeof v !== "undefined") chkFrameEnable.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["frame", "width"]), st.frameWidth, undefined);
            if (typeof v !== "undefined") editFrameWidth.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["frame", "bleed"]), st.bleed, undefined);
            if (typeof v !== "undefined") chkBleed.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["frame", "round", "enable"]), st.frameRound, undefined);
            if (typeof v !== "undefined") chkFrameRound.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["frame", "round", "val"]), st.frameRoundVal, undefined);
            if (typeof v !== "undefined") editFrameRound.text = String(v);
        });

        // Inner
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "link"]), st.innerLink, undefined);
            if (typeof v !== "undefined") chkInnerOffsetLink.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "offset", "top"]), st.offTop, undefined);
            if (typeof v !== "undefined") editInnerOffsetTop.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "offset", "bottom"]), st.offBottom, undefined);
            if (typeof v !== "undefined") editInnerOffsetBottom.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "offset", "left"]), st.offLeft, undefined);
            if (typeof v !== "undefined") editInnerOffsetLeft.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "offset", "right"]), st.offRight, undefined);
            if (typeof v !== "undefined") editInnerOffsetRight.text = String(v);
        });

        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "grid", "cols"]), st.cols, undefined);
            if (typeof v !== "undefined") editInnerColumns.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "grid", "rows"]), st.rows, undefined);
            if (typeof v !== "undefined") editInnerRows.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "grid", "gutter", "col"]), st.colGutter, undefined);
            if (typeof v !== "undefined") editColGutter.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "grid", "gutter", "row"]), st.rowGutter, undefined);
            if (typeof v !== "undefined") editRowGutter.text = String(v);
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "grid", "fill"]), st.rowFill, undefined);
            if (typeof v !== "undefined") chkRowFill.value = !!v;
        });
        safeSet(function () {
            var v = firstDefined(getNested(st, ["inner", "grid", "divider"]), st.rowDivider, undefined);
            if (typeof v !== "undefined") chkRowDivider.value = !!v;
        });

        safeSet(function () {
            var lt = firstDefined(getNested(st, ["inner", "grid", "lineType"]), st.innerLine, undefined);
            if (lt === "dash") rbInnerLineDash.value = true;
            else if (lt === "dotdash") rbInnerLineDotDash.value = true;
            else if (typeof lt !== "undefined") rbInnerLineSolid.value = true;
        });

        // Re-apply enabled/disabled states that depend on other controls
        safeSet(function () { applyEdgeScaleEnabledState(); });
        safeSet(function () { applyArtboardMarginLinkState(); });
        safeSet(function () { applyOuterLinePanelEnabledState(); });
        safeSet(function () { applyOuterRoundEnabledState(); });
        safeSet(function () { applyTitleAreaEnabledState(); });
        safeSet(function () { applyFrameEnabledState(); });
        safeSet(function () {
            var c = parseInt(editInnerColumns.text, 10);
            var r = parseInt(editInnerRows.text, 10);
            if (isNaN(c) || c < 1) c = 1;
            if (isNaN(r) || r < 1) r = 1;
            applyRowDividerEnabledState(c, r, false);
        });
        safeSet(function () { applyInnerOffsetLinkState(); });
    })();

    // --- イベント処理 ---

    // 値が変更されたらプレビュー更新
    // onChanging はキー入力のたびに発火します
    editVal.onChanging = function () {
        if (!chkEnableLen.value) return;
        applyOuterLinePanelEnabledState();
        if (chkPreview.value) {
            updatePreview();
        }
    };
    chkEnableLen.onClick = function () {
        applyEdgeScaleEnabledState();
        applyOuterLinePanelEnabledState();
        applyOuterRoundEnabledState();
        if (chkPreview.value) updatePreview();
    };

    // マージン（アートボード基準のみ）
    editArtboardMarginTop.onChanging = function () {
        if (!_usingArtboardBase) return;
        if (_syncingArtboardMargins) return;

        if (chkArtboardMarginLink && chkArtboardMarginLink.value) {
            var v = editArtboardMarginTop.text;
            _syncingArtboardMargins = true;
            try { editArtboardMarginBottom.text = v; } catch (_) { }
            try { editArtboardMarginLeft.text = v; } catch (_) { }
            try { editArtboardMarginRight.text = v; } catch (_) { }
            _syncingArtboardMargins = false;
        }
        if (chkPreview.value) updatePreview();
    };
    editArtboardMarginBottom.onChanging = function () { if (!_usingArtboardBase) return; if (_syncingArtboardMargins) return; if (chkPreview.value) updatePreview(); };
    editArtboardMarginLeft.onChanging = function () { if (!_usingArtboardBase) return; if (_syncingArtboardMargins) return; if (chkPreview.value) updatePreview(); };
    editArtboardMarginRight.onChanging = function () { if (!_usingArtboardBase) return; if (_syncingArtboardMargins) return; if (chkPreview.value) updatePreview(); };

    chkArtboardMarginLink.onClick = function () {
        applyArtboardMarginLinkState();
        if (chkPreview.value) updatePreview();
    };

    // 裁ち落とし
    chkBleed.onClick = function () {
        _bleedEnabled = !!chkBleed.value;
        if (chkPreview.value) updatePreview();
    };

    // フレーム幅
    editFrameWidth.onChanging = function () {
        if (!chkFrameEnable.value) {
            editFrameWidth.text = "0";
            applyFrameEnabledState();
            if (chkPreview.value) updatePreview();
            return;
        }
        try {
            var _fw = parseFloat(editFrameWidth.text);
            var _valid = (!isNaN(_fw) && _fw > 0);

            // 裁ち落とし
            var _enableBleed = (_usingArtboardBase && _valid);
            chkBleed.enabled = _enableBleed;

            if (_enableBleed) {
                // 幅が0→>0になった場合、自動で裁ち落としON
                if (!chkBleed.value) {
                    chkBleed.value = true;
                }
            } else {
                chkBleed.value = false;
            }

            // 角丸
            chkFrameRound.enabled = _valid;
            if (!_valid) {
                chkFrameRound.value = false;
                editFrameRound.enabled = false;
            }
        } catch (_) { }
        if (chkPreview.value) updatePreview();
    };

    // 角丸（フレーム）
    chkFrameRound.onClick = function () {
        try {
            editFrameRound.enabled = chkFrameRound.value;

            // ONにしたとき、現在値が0ならデフォルトで「2」（現在の単位系）を入れる
            if (chkFrameRound.value) {
                var v = parseFloat(editFrameRound.text);
                if (isNaN(v) || v === 0) {
                    editFrameRound.text = "2";
                }
            }
        } catch (_) { }
        if (chkPreview.value) updatePreview();
    };

    editFrameRound.onChanging = function () {
        if (chkPreview.value) updatePreview();
    };

    editInnerOffsetTop.onChanging = function () {
        if (_syncingInnerOffsets) return;
        if (chkInnerOffsetLink.value) {
            var v = editInnerOffsetTop.text;
            _syncingInnerOffsets = true;
            try { editInnerOffsetBottom.text = v; } catch (_) { }
            try { editInnerOffsetLeft.text = v; } catch (_) { }
            try { editInnerOffsetRight.text = v; } catch (_) { }
            _syncingInnerOffsets = false;
        }
        if (chkPreview.value) updatePreview();
    };
    editInnerOffsetBottom.onChanging = function () { if (_syncingInnerOffsets) return; if (chkPreview.value) updatePreview(); };
    editInnerOffsetLeft.onChanging = function () { if (_syncingInnerOffsets) return; if (chkPreview.value) updatePreview(); };
    editInnerOffsetRight.onChanging = function () { if (_syncingInnerOffsets) return; if (chkPreview.value) updatePreview(); };
    editInnerColumns.onChanging = function () {
        var c = parseInt(editInnerColumns.text, 10);
        if (isNaN(c) || c < 1) c = 1;
        if (String(c) !== editInnerColumns.text) editInnerColumns.text = String(c);
        // 列が1なら列ガターはディム表示
        try { editColGutter.enabled = (c > 1); } catch (_) { }

        var r = parseInt(editInnerRows.text, 10);
        if (isNaN(r) || r < 1) r = 1;
        if (String(r) !== editInnerRows.text) editInnerRows.text = String(r);

        try { if (typeof applyGridGutterEnabledState === "function") applyGridGutterEnabledState(c, r); } catch (_) { }
        try { applyRowDividerEnabledState(c, r, true); } catch (_) { }
        if (chkPreview.value) updatePreview();
    };

    editInnerRows.onChanging = function () {
        var c = parseInt(editInnerColumns.text, 10);
        if (isNaN(c) || c < 1) c = 1;
        if (String(c) !== editInnerColumns.text) editInnerColumns.text = String(c);

        var r = parseInt(editInnerRows.text, 10);
        if (isNaN(r) || r < 1) r = 1;
        if (String(r) !== editInnerRows.text) editInnerRows.text = String(r);

        try { if (typeof applyGridGutterEnabledState === "function") applyGridGutterEnabledState(c, r); } catch (_) { }
        try { applyRowDividerEnabledState(c, r, true); } catch (_) { }
        if (chkPreview.value) updatePreview();
    };

    // チェックボックス切り替え時
    chkPreview.onClick = function () {
        if (chkPreview.value) {
            updatePreview();
        } else {
            clearPreview(); // プレビューOFFなら元に戻す
        }
    };
    function onCapChanged() {
        if (chkPreview.value) updatePreview();
    }
    rbCapButt.onClick = onCapChanged;
    rbCapRound.onClick = onCapChanged;
    rbCapProject.onClick = onCapChanged;

    function onInnerCapChanged() {
        if (chkPreview.value) updatePreview();
    }
    rbInnerLineSolid.onClick = onInnerCapChanged;
    rbInnerLineDash.onClick = onInnerCapChanged;
    rbInnerLineDotDash.onClick = onInnerCapChanged;

    chkKeepOuter.onClick = function () {
        applyEdgeScaleEnabledState();
        applyOuterLinePanelEnabledState();
        applyOuterRoundEnabledState();

        if (chkPreview.value) updatePreview();
    };

    chkTitleFill.onClick = function () {
        if (chkPreview.value) updatePreview();
    };

    chkTitleLine.onClick = function () {
        applyTitleAreaEnabledState();
        if (chkPreview.value) updatePreview();
    };

    chkInnerOffsetLink.onClick = function () {
        applyInnerOffsetLinkState();
        if (chkPreview.value) updatePreview();
    };

    chkRowDivider.onClick = function () {
        try { innerCapPanel.enabled = (chkRowDivider.enabled && chkRowDivider.value); } catch (_) { }
        if (chkPreview.value) updatePreview();
    };

    chkRowFill.onClick = function () {
        // ユーザーが明示的に操作したら以後は自動ONしない
        _rowFillManuallySet = true;
        if (chkPreview.value) updatePreview();
    };

    // 列のガター変更時
    editColGutter.onChanging = function () {
        var g = parseFloat(editColGutter.text);
        // ガターが設定されたら塗りを自動ON（ただし手動操作があれば尊重）
        if (!_rowFillManuallySet && !isNaN(g) && g !== 0) {
            try { if (!chkRowFill.value) chkRowFill.value = true; } catch (_) { }
        }
        if (chkPreview.value) updatePreview();
    };

    // 行のガター変更時
    editRowGutter.onChanging = function () {
        var g = parseFloat(editRowGutter.text);
        // ガターが設定されたら塗りを自動ON（ただし手動操作があれば尊重）
        if (!_rowFillManuallySet && !isNaN(g) && g !== 0) {
            try { if (!chkRowFill.value) chkRowFill.value = true; } catch (_) { }
        }
        if (chkPreview.value) updatePreview();
    };


    // レイアウト確定（tabbedpanel の内容が潰れるのを防ぐ）
    try { win.layout.layout(true); } catch (_) { }
    try { win.layout.resize(); } catch (_) { }
    updatePreview(false);

    // --- ダイアログ表示 ---
    var result = win.show();

    persistUIState();

    // キャンセル時：プレビュー生成物を削除して終了
    if (result != 1) {
        try { if (viewCtl && typeof viewCtl.restore === "function") viewCtl.restore(); } catch (_) { }
        try { clearPreview(); } catch (_) { }
        return;
    }

    // ダイアログ終了時：プレビューで増えたUndoを戻す（OK時は後で最終生成し直す）

    if (result == 1) {
        // OKが押された場合
        // 長さ調整=0 かつ オフセット=0 のときは何もせず終了
        var _okVal = getEffectiveLenValue();
        var _okOffTop = parseFloat(editInnerOffsetTop.text);
        var _okOffBottom = parseFloat(editInnerOffsetBottom.text);
        var _okOffLeft = parseFloat(editInnerOffsetLeft.text);
        var _okOffRight = parseFloat(editInnerOffsetRight.text);
        if ((!isNaN(_okVal) && _okVal <= 0)
            && (!isNaN(_okOffTop) && _okOffTop <= 0)
            && (!isNaN(_okOffBottom) && _okOffBottom <= 0)
            && (!isNaN(_okOffLeft) && _okOffLeft <= 0)
            && (!isNaN(_okOffRight) && _okOffRight <= 0)) {
            // 何も変更がない場合はそのまま終了（プレビューは消さない）
            persistUIState();
            return;
        }


        // 最終生成（OK結果はヒストリーに残す）
        updatePreview(true);

        // 外枠の扱い：
        // - 辺の長さ調整 > 0 の場合：元の長方形は不要なので常に削除（外枠は4辺線で表現）
        // - 辺の長さ調整 <= 0 の場合：チェックOFFなら元の長方形も削除（外枠なし）
        var _okLen = getEffectiveLenValue();
        if (!isNaN(_okLen) && _okLen !== 0) {
            // 後でガイドを作るため、元の外形（bounds）を控える
            var _outerGuideSpecs = [];
            for (var i = 0; i < targetItems.length; i++) {
                try {
                    var _it = targetItems[i];
                    var _b = _it.geometricBounds; // [L, T, R, B]
                    _outerGuideSpecs.push({
                        layer: _it.layer,
                        left: _b[0],
                        top: _b[1],
                        right: _b[2],
                        bottom: _b[3]
                    });
                } catch (_) { }
            }

            // 元の長方形は削除（外枠は4辺線で表現）
            // ただしアートボード基準の外枠矩形は残す
            for (var i = 0; i < targetItems.length; i++) {
                try {
                    if (_usingArtboardBase && _artboardBaseRect && targetItems[i] === _artboardBaseRect) continue;
                    targetItems[i].remove();
                } catch (_) { }
            }
        } else {
            // 辺の伸縮がOFF（=0扱い）の場合でも、外枠の表示/削除は「外枠を残す」に従う
            if (!chkEnableLen.value) {
                if (!chkKeepOuter.value) {
                    for (var i = 0; i < targetItems.length; i++) {
                        try { targetItems[i].remove(); } catch (_) { }
                    }
                } else {
                    for (var i = 0; i < targetItems.length; i++) {
                        try { targetItems[i].hidden = false; } catch (_) { }
                    }
                }
            } else {
                if (!chkEnableLen.value) {
                    // 辺の長さ調整OFFでは4辺線を作らない（長方形を残す）
                } else if (!chkKeepOuter.value) {
                    for (var i = 0; i < targetItems.length; i++) {
                        try {
                            if (_usingArtboardBase && _artboardBaseRect && targetItems[i] === _artboardBaseRect) continue;
                            targetItems[i].remove();
                        } catch (_) { }
                    }
                } else {
                    for (var i = 0; i < targetItems.length; i++) {
                        try { targetItems[i].hidden = false; } catch (_) { }
                    }
                }
            }
        }

        // 外枠を残す OFF のときは、生成した4辺（外枠線）を削除
        if (!chkKeepOuter.value) {
            for (var i = tempPreviewItems.length - 1; i >= 0; i--) {
                try {
                    var it = tempPreviewItems[i];
                    var isOuterEdge = false;
                    try { if (it.note === "__OuterEdge__") isOuterEdge = true; } catch (_) { }
                    if (!isOuterEdge) {
                        try { if (it.name === "__OuterEdge__") isOuterEdge = true; } catch (_) { }
                    }
                    if (isOuterEdge) {
                        try { it.remove(); } catch (_) { }
                        tempPreviewItems.splice(i, 1);
                    }
                } catch (_) { }
            }
        }



        // 内側の長方形（__InnerBoxFill__）の扱い
        // 塗りON：通常オブジェクトとして残す（ガイド化しない）
        // 塗りOFF：内側エリアの塗りは残さない（オブジェクト自体を削除）
        for (var i = tempPreviewItems.length - 1; i >= 0; i--) {
            try {
                if (tempPreviewItems[i] && tempPreviewItems[i].typename === "PathItem") {
                    var it = tempPreviewItems[i];
                    var isInnerFill = false;
                    try { if (it.note === "__InnerBoxFill__") isInnerFill = true; } catch (_) { }
                    if (!isInnerFill) {
                        try { if (it.name === "__InnerBoxFill__") isInnerFill = true; } catch (_) { }
                    }
                    if (isInnerFill) {
                        if (chkRowFill && chkRowFill.value) {
                            // 通常オブジェクトとして保持（念のためガイド属性を解除）
                            try { it.guides = false; } catch (_) { }
                            try { it.stroked = false; } catch (_) { }
                            try { it.filled = true; } catch (_) { }
                            try { it.fillColor = makeK15Fill(); } catch (_) { }
                        } else {
                            // 塗りOFF：内側エリアの塗りは残さない
                            try { it.remove(); } catch (_) { }
                        }
                        // 選択リストから外す
                        tempPreviewItems.splice(i, 1);
                    }
                }
            } catch (_) { }
        }

        // タイトル帯の塗り（__TitleFill__）は残し、選択対象から外す
        for (var i = tempPreviewItems.length - 1; i >= 0; i--) {
            try {
                var it2 = tempPreviewItems[i];
                var isTitleFill = false;
                try { if (it2.note === "__TitleFill__") isTitleFill = true; } catch (_) { }
                if (!isTitleFill) {
                    try { if (it2.name === "__TitleFill__") isTitleFill = true; } catch (_) { }
                }
                if (isTitleFill) tempPreviewItems.splice(i, 1);
            } catch (_) { }
        }

        // フレーム（__FrameFill__）は残し、選択対象から外す
        for (var i = tempPreviewItems.length - 1; i >= 0; i--) {
            try {
                var itF = tempPreviewItems[i];
                var isFrame = false;
                try { if (itF.note === "__FrameFill__") isFrame = true; } catch (_) { }
                if (!isFrame) {
                    try { if (itF.name === "__FrameFill__") isFrame = true; } catch (_) { }
                }
                if (isFrame) tempPreviewItems.splice(i, 1);
            } catch (_) { }
        }

        // アートボード基準の外枠矩形は残す（OK後に消さない）
        // ※キャンセル時は clearPreview() 内で cleanupArtboardBaseRect() が破棄する
        // ダイアログ終了時は何も選択しない状態にする
        try { doc.selection = null; } catch (_) { }

    }


    // --- 関数定義 ---

    // -----------------------------
    // Preview / Generation split
    // - collectOptions(): read UI, compute pt values
    // - generateFromOptions(): create objects (core)
    // - renderPreview()/renderFinal(): preview-specific cleanup/redraw
    // Public entry remains: updatePreview(isFinal)
    // -----------------------------

    function collectOptions() {
        // 入力値チェック
        var factor = getCurrentRulerPtFactor();

        var val = getEffectiveLenValue();
        var distPt = val * factor; // rulerType -> pt

        // タイトルエリア：辺の伸縮（タイトル帯の線の長さにのみ反映）
        var titleLenVal = 0;
        try {
            if (chkTitleLine && chkTitleLine.value && chkTitleEdgeScale && chkTitleEdgeScale.value) {
                var t = parseFloat(editTitleEdgeScale && editTitleEdgeScale.text);
                if (!isNaN(t)) titleLenVal = t;
            }
        } catch (_) { }
        var titleDistPt = (-titleLenVal) * factor; // 正負を反転

        // タイトル領域サイズ（外枠基準）
        var titleVal = (typeof chkTitleEnable !== "undefined" && chkTitleEnable && !chkTitleEnable.value) ? 0 : parseFloat(editTitleSize && editTitleSize.text);
        var titleSizePt = (!isNaN(titleVal) ? (titleVal * factor) : 0);

        // フレーム幅（pt）
        var frameVal = (chkFrameEnable && chkFrameEnable.value) ? parseFloat(editFrameWidth.text) : 0;
        var framePt = (!isNaN(frameVal) ? (frameVal * factor) : 0);
        if (framePt < 0) framePt = 0;
        if (_bleedEnabled) {
            var bleedPtForFrame = (72.0 / 25.4) * BLEED_MM;
            framePt += bleedPtForFrame;
        }

        // 内側オフセット
        var offTopVal = parseFloat(editInnerOffsetTop.text);
        var offBottomVal = parseFloat(editInnerOffsetBottom.text);
        var offLeftVal = parseFloat(editInnerOffsetLeft.text);
        var offRightVal = parseFloat(editInnerOffsetRight.text);

        var offTopPt = (!isNaN(offTopVal) ? (offTopVal * factor) : 0);
        var offBottomPt = (!isNaN(offBottomVal) ? (offBottomVal * factor) : 0);
        var offLeftPt = (!isNaN(offLeftVal) ? (offLeftVal * factor) : 0);
        var offRightPt = (!isNaN(offRightVal) ? (offRightVal * factor) : 0);

        // 列・行
        var colVal = parseInt(editInnerColumns.text, 10);
        if (isNaN(colVal) || colVal < 1) colVal = 1;
        var rowVal = parseInt(editInnerRows.text, 10);
        if (isNaN(rowVal) || rowVal < 1) rowVal = 1;

        // ガター（列/行）：列/行が1なら無効（0扱い）
        try { editColGutter.enabled = (colVal > 1); } catch (_) { }
        try { editRowGutter.enabled = (rowVal > 1); } catch (_) { }

        var colGutterVal = parseFloat(editColGutter.text);
        var rowGutterVal = parseFloat(editRowGutter.text);
        var colGutterPt = (colVal > 1 && !isNaN(colGutterVal)) ? (colGutterVal * factor) : 0;
        var rowGutterPt = (rowVal > 1 && !isNaN(rowGutterVal)) ? (rowGutterVal * factor) : 0;
        if (colGutterPt < 0) colGutterPt = 0;
        if (rowGutterPt < 0) rowGutterPt = 0;

        // ガターが設定されたら塗りを自動ON（ただし手動操作があれば尊重）
        if (!_rowFillManuallySet && ((colGutterPt && colGutterPt !== 0) || (rowGutterPt && rowGutterPt !== 0))) {
            try { if (!chkRowFill.value) chkRowFill.value = true; } catch (_) { }
        }

        // 依存UIの更新
        try { applyRowDividerEnabledState(colVal, rowVal, false); } catch (_) { }
        try { applyOuterLinePanelEnabledState(); } catch (_) { }

        return {
            factor: factor,
            distPt: distPt,
            titleSizePt: titleSizePt,
            titleDistPt: titleDistPt,
            framePt: framePt,
            offTopPt: offTopPt,
            offBottomPt: offBottomPt,
            offLeftPt: offLeftPt,
            offRightPt: offRightPt,
            colVal: colVal,
            rowVal: rowVal,
            colGutterPt: colGutterPt,
            rowGutterPt: rowGutterPt
        };
    }

    function generateFromOptions(opt, isFinal) {
        // アートボード基準のときは、マージンを反映した一時矩形に更新（裁ち落としはフレームのみに適用）
        if (_usingArtboardBase) {
            var mtVal = parseFloat(editArtboardMarginTop.text);
            var mbVal = parseFloat(editArtboardMarginBottom.text);
            var mlVal = parseFloat(editArtboardMarginLeft.text);
            var mrVal = parseFloat(editArtboardMarginRight.text);

            if (isNaN(mtVal) || mtVal < 0) mtVal = 0;
            if (isNaN(mbVal) || mbVal < 0) mbVal = 0;
            if (isNaN(mlVal) || mlVal < 0) mlVal = 0;
            if (isNaN(mrVal) || mrVal < 0) mrVal = 0;

            var mtPt = mtVal * opt.factor;
            var mbPt = mbVal * opt.factor;
            var mlPt = mlVal * opt.factor;
            var mrPt = mrVal * opt.factor;

            try { if (typeof chkBleed !== "undefined") _bleedEnabled = !!chkBleed.value; } catch (_) { }
            rebuildArtboardBaseRect(mtPt, mrPt, mbPt, mlPt);
        }

        // distPt===0：分割しない
        if (opt.distPt === 0) {
            // フレーム（アートボードサイズ基準）
            if (opt.framePt > 0) {
                var abB = getActiveArtboardBounds();
                if (abB) {
                    createFrameFill(targetItems[0], opt.framePt, abB);
                }
            }

            // タイトル帯（塗り）
            if (opt.titleSizePt > 0) {
                for (var i = 0; i < targetItems.length; i++) {
                    createTitleFill(targetItems[i], opt.titleSizePt);
                }
            }

            // タイトル帯の分割線
            if (opt.titleSizePt > 0) {
                for (var i = 0; i < targetItems.length; i++) {
                    createTitleDivider(targetItems[i], opt.titleSizePt, opt.titleDistPt);
                }
            }

            // 線端パネルはディム表示
            try { capPanel.enabled = false; } catch (_) { }

            // 外枠の表示は「外枠を残す」に従う
            var showOuterRect = !!chkKeepOuter.value;
            for (var i = 0; i < targetItems.length; i++) {
                try { targetItems[i].hidden = !showOuterRect; } catch (_) { }
            }

            // 内側エリア（オフセットが0でも描画）
            for (var i = 0; i < targetItems.length; i++) {
                var innerB = getInnerAreaBounds(targetItems[i], opt.titleSizePt);
                if (innerB) {
                    createInnerBox(
                        targetItems[i],
                        innerB,
                        opt.offTopPt, opt.offBottomPt, opt.offLeftPt, opt.offRightPt,
                        opt.colVal, opt.rowVal,
                        opt.colGutterPt, opt.rowGutterPt
                    );
                }
            }

            return;
        }

        // distPt!=0：4辺線
        try { capPanel.enabled = true; } catch (_) { }

        // 分解した4辺を表示するため、元の長方形は常に隠す
        for (var i = 0; i < targetItems.length; i++) {
            try { targetItems[i].hidden = true; } catch (_) { }
        }

        // 外枠を残すONのときのみ4辺罫線を描画
        if (chkKeepOuter.value) {
            for (var i = 0; i < targetItems.length; i++) {
                createShortenedLines(targetItems[i], opt.distPt);
            }
        }

        // フレーム（アートボードサイズ基準）
        if (opt.framePt > 0) {
            var abB2 = getActiveArtboardBounds();
            if (abB2) {
                createFrameFill(targetItems[0], opt.framePt, abB2);
            }
        }

        // タイトル帯（塗り）
        if (opt.titleSizePt > 0) {
            for (var i = 0; i < targetItems.length; i++) {
                createTitleFill(targetItems[i], opt.titleSizePt);
            }
        }

        // タイトル帯の分割線
        if (opt.titleSizePt > 0) {
            for (var i = 0; i < targetItems.length; i++) {
                createTitleDivider(targetItems[i], opt.titleSizePt, opt.titleDistPt);
            }
        }

        // 内側エリア
        for (var i = 0; i < targetItems.length; i++) {
            var innerB2 = getInnerAreaBounds(targetItems[i], opt.titleSizePt);
            if (innerB2) {
                createInnerBox(
                    targetItems[i],
                    innerB2,
                    opt.offTopPt, opt.offBottomPt, opt.offLeftPt, opt.offRightPt,
                    opt.colVal, opt.rowVal,
                    opt.colGutterPt, opt.rowGutterPt
                );
            }
        }
    }

    function renderPreview() {
        removeTempItems();
        var opt = collectOptions();
        generateFromOptions(opt, false);
        app.redraw();
    }

    function renderFinal() {
        removeTempItems();
        var opt = collectOptions();
        generateFromOptions(opt, true);
        app.redraw();
    }

    function updatePreview(isFinal) {
        if (isFinal) renderFinal();
        else renderPreview();
    }

    function persistUIState() {
        var st = {};
        try {
            st.preview = !!chkPreview.value;

            st.margin = {
                top: editArtboardMarginTop.text,
                bottom: editArtboardMarginBottom.text,
                left: editArtboardMarginLeft.text,
                right: editArtboardMarginRight.text,
                link: !!chkArtboardMarginLink.value
            };

            st.outer = {
                keepOuter: !!chkKeepOuter.value,
                enableLen: !!chkEnableLen.value,
                lenVal: editVal.text,
                cap: (rbCapRound.value ? "round" : (rbCapProject.value ? "project" : "butt")),
                round: {
                    enable: !!chkOuterRound.value,
                    val: editOuterRound.text
                }
            };

            st.title = {
                enable: !!chkTitleEnable.value,
                size: editTitleSize.text,
                pos: (rbTitleBottom.value ? "bottom" : (rbTitleLeft.value ? "left" : (rbTitleRight.value ? "right" : "top"))),
                fill: !!chkTitleFill.value,
                line: !!chkTitleLine.value,
                edgeScale: {
                    enable: !!chkTitleEdgeScale.value,
                    val: editTitleEdgeScale.text
                }
            };

            st.frame = {
                enable: !!chkFrameEnable.value,
                width: editFrameWidth.text,
                bleed: !!chkBleed.value,
                round: {
                    enable: !!chkFrameRound.value,
                    val: editFrameRound.text
                }
            };

            st.inner = {
                link: !!chkInnerOffsetLink.value,
                offset: {
                    top: editInnerOffsetTop.text,
                    bottom: editInnerOffsetBottom.text,
                    left: editInnerOffsetLeft.text,
                    right: editInnerOffsetRight.text
                },
                grid: {
                    cols: editInnerColumns.text,
                    rows: editInnerRows.text,
                    gutter: {
                        col: editColGutter.text,
                        row: editRowGutter.text
                    },
                    fill: !!chkRowFill.value,
                    divider: !!chkRowDivider.value,
                    lineType: (rbInnerLineDash.value ? "dash" : (rbInnerLineDotDash.value ? "dotdash" : "solid"))
                }
            };
        } catch (_) { }
        __saveState(st);
    }

    // --- 単位（rulerType）ユーティリティ ---
    // 単位コード→ラベル（Q/H分岐は rulerType では H 扱い）

    function getUnitLabel(code, prefKey) {
        // code 5 は Q/H（rulerType は H 扱い）
        if (code === 5) {
            return (prefKey === "rulerType") ? "H" : "Q";
        }
        return _unitMap[code] || "pt";
    }

    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H (0.25mm)
            case 6: return 1.0;                         // px（Illustrator内部はpt基準で扱われることが多い）
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    function getCurrentRulerUnitCode() {
        try {
            return app.preferences.getIntegerPreference("rulerType");
        } catch (_) {
            return 2; // pt
        }
    }

    function getCurrentRulerUnitLabel() {
        var code = getCurrentRulerUnitCode();
        return getUnitLabel(code, "rulerType");
    }

    function getCurrentRulerPtFactor() {
        var code = getCurrentRulerUnitCode();
        return getPtFactorFromUnitCode(code);
    }

    /* アートボード矩形 / Get active artboard bounds */
    function getActiveArtboardBounds() {
        try {
            var abIndex = doc.artboards.getActiveArtboardIndex();
            var abRect = doc.artboards[abIndex].artboardRect; // [L,T,R,B]
            var L = abRect[0], T = abRect[1], R = abRect[2], B = abRect[3];
            if (_bleedEnabled) {
                var bleedPt = (72.0 / 25.4) * BLEED_MM;
                L -= bleedPt;
                T += bleedPt;
                R += bleedPt;
                B -= bleedPt;
            }
            return [L, T, R, B];
        } catch (_) {
            return null;
        }
    }

    // 辺の長さ調整の有効/無効を考慮した値を返す（無効時は0）
    function getEffectiveLenValue() {
        try {
            if (!chkEnableLen || !chkEnableLen.value) return 0;
            var v = parseFloat(editVal.text);
            return isNaN(v) ? 0 : v;
        } catch (_) {
            return 0;
        }
    }

    // ↑↓キーで値を増減（shift=±10, option=±0.1）
    // allowNegative=false の場合は 0 未満にしない
    function changeValueByArrowKey(editText, allowNegative) {
        editText.addEventListener("keydown", function (event) {
            if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") {
                    value += delta;
                } else {
                    value -= delta;
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                } else {
                    value -= delta;
                }
            }

            if (!allowNegative && value < 0) value = 0;

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;

            // keydownでtextを書き換えた場合、onChangingが発火しないことがあるため明示的に呼ぶ
            try {
                if (typeof editText.onChanging === "function") {
                    editText.onChanging();
                }
            } catch (_) { }

            // イベントのデフォルト動作を抑制
            try { event.preventDefault(); } catch (_) { }
        });
    }

    // UIで選択された線端を取得
    function getSelectedStrokeCap() {
        if (rbCapRound.value) return StrokeCap.ROUNDENDCAP;
        if (rbCapProject.value) return StrokeCap.PROJECTINGENDCAP;
        return StrokeCap.BUTTENDCAP; // 線端なし
    }



    // ドット点線を適用（線端を丸型にして strokeDashes=[0, strokeW*2]）
    function applyDotDash(pathItem) {
        try {
            if (!pathItem.stroked) {
                pathItem.stroked = true;
            }
            pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
            try { pathItem.strokeJoin = StrokeJoin.ROUNDENDJOIN; } catch (_) { }
            var strokeW = pathItem.strokeWidth;
            pathItem.strokeDashes = [0, strokeW * 2];
        } catch (e) {
            // エラー時は何もしない
        }
    }

    // 内側の区切り線に線種を適用
    function applyInnerLineStyle(pathItem) {
        try {
            // 実線
            if (rbInnerLineSolid && rbInnerLineSolid.value) {
                try { pathItem.strokeWidth = 1; } catch (_) { }
                try { pathItem.strokeDashes = []; } catch (_) { }
                // 線端は外枠の線端設定に合わせる
                try { pathItem.strokeCap = getSelectedStrokeCap(); } catch (_) { }
                return;
            }

            // 点線（ダッシュ）
            if (rbInnerLineDash && rbInnerLineDash.value) {
                try { pathItem.strokeWidth = 1; } catch (_) { }
                var w = pathItem.strokeWidth;
                try { pathItem.strokeDashes = [w * 4, w * 2]; } catch (_) { }
                try { pathItem.strokeCap = getSelectedStrokeCap(); } catch (_) { }
                return;
            }

            // ドット点線
            if (rbInnerLineDotDash && rbInnerLineDotDash.value) {
                try { pathItem.strokeWidth = 2; } catch (_) { }
                applyDotDash(pathItem);
                return;
            }
        } catch (_) { }
    }


    // K15（CMYK: K=15）塗りを返す
    function makeK15Fill() {
        var c = new CMYKColor();
        c.cyan = 0;
        c.magenta = 0;
        c.yellow = 0;
        c.black = 15;
        return c;
    }

    // K30（CMYK: K=30）塗りを返す
    function makeK30Fill() {
        var c = new CMYKColor();
        c.cyan = 0;
        c.magenta = 0;
        c.yellow = 0;
        c.black = 30;
        return c;
    }

    // K0（白）塗りを返す
    function makeK0Fill() {
        var c = new CMYKColor();
        c.cyan = 0;
        c.magenta = 0;
        c.yellow = 0;
        c.black = 0;
        return c;
    }

    // 角丸（ライブエフェクト） / Round Corners live effect
    function createRoundCornersEffectXML(radiusPt) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radiusPt);
    }

    function applyRoundCornersEffect(item, radiusPt) {
        try {
            if (!item) return;
            if (!(radiusPt > 0)) return;
            var xml = createRoundCornersEffectXML(radiusPt);
            item.applyEffect(xml);
        } catch (_) { }
    }

    // フレームを作成して tempPreviewItems に登録（外枠基準 or 指定bounds）
    // 外：K50 / 内：透明の穴あき（グループ＋Live Pathfinder Exclude）
    function createFrameFill(pathItem, framePt, baseBounds) {
        try {
            if (!framePt || framePt <= 0) return;

            var b = (baseBounds && baseBounds.length === 4) ? baseBounds : pathItem.geometricBounds; // [L,T,R,B]
            var L = b[0], T = b[1], R = b[2], B = b[3];
            var w = R - L;
            var h = T - B;
            if (!(w > 0) || !(h > 0)) return;

            var innerL = L + framePt;
            var innerT = T - framePt;
            var innerR = R - framePt;
            var innerB = B + framePt;
            var innerW = innerR - innerL;
            var innerH = innerT - innerB;
            if (!(innerW > 0) || !(innerH > 0)) return;

            var layer = pathItem.layer;

            // K50 塗り
            var c50 = new CMYKColor();
            c50.cyan = 0;
            c50.magenta = 0;
            c50.yellow = 0;
            c50.black = 50;

            // 外側・内側の矩形を作成
            var outerRect = layer.pathItems.rectangle(T, L, w, h);
            outerRect.stroked = false;
            outerRect.filled = true;
            outerRect.fillColor = c50;

            var innerRect = layer.pathItems.rectangle(innerT, innerL, innerW, innerH);
            innerRect.stroked = false;
            innerRect.filled = true;
            // 内側は一時的な塗り（色は最終的に Exclude の結果で穴になる）
            innerRect.fillColor = makeK0Fill();

            // 角丸は「内側の長方形」に適用する（穴側を丸める）
            try {
                if (typeof chkFrameRound !== "undefined" && chkFrameRound && chkFrameRound.value) {
                    var rVal0 = parseFloat(editFrameRound.text);
                    if (!isNaN(rVal0) && rVal0 > 0) {
                        // 入力値は rulerType 単位 → pt に換算
                        var rPt0 = rVal0 * getCurrentRulerPtFactor();
                        applyRoundCornersEffect(innerRect, rPt0);
                    }
                }
            } catch (_) { }

            // グループ化
            var g = layer.groupItems.add();
            outerRect.move(g, ElementPlacement.PLACEATEND);
            innerRect.move(g, ElementPlacement.PLACEATEND);

            // 既存選択を退避して、グループだけ選択
            var prevSel;
            try { prevSel = doc.selection; } catch (_) { prevSel = null; }
            try { doc.selection = null; } catch (_) { }
            try { g.selected = true; } catch (_) { }

            // Live Pathfinder Exclude を実行
            try { app.executeMenuCommand('Live Pathfinder Exclude'); } catch (_) { }

            // 結果オブジェクトを取得（selection の先頭を採用）
            var resultItem = null;
            try {
                if (doc.selection && doc.selection.length > 0) {
                    resultItem = doc.selection[0];
                }
            } catch (_) { }

            // 選択を元に戻す
            try { doc.selection = prevSel; } catch (_) { }

            // 結果が取れない場合は g を使う
            if (!resultItem) resultItem = g;

            // タグ付け
            try { resultItem.name = "__FrameFill__"; } catch (_) { }
            try { resultItem.note = "__FrameFill__"; } catch (_) { }


            // 背面へ
            try { resultItem.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }

            tempPreviewItems.push(resultItem);
        } catch (_) { }
    }

    // タイトル帯：塗り矩形を作成して tempPreviewItems に登録（外枠基準）
    function createTitleFill(pathItem, sizePt) {
        try {
            if (!chkTitleFill || !chkTitleFill.value) return;
            if (!sizePt || sizePt <= 0) return;

            var b = pathItem.geometricBounds; // [L,T,R,B]
            var L = b[0], T = b[1], R = b[2], B = b[3];
            var w = R - L;
            var h = T - B;
            if (w <= 0 || h <= 0) return;

            var rectTop, rectLeft, rectW, rectH;

            if (rbTitleTop.value) {
                if (sizePt >= h) return;
                rectTop = T; rectLeft = L; rectW = w; rectH = sizePt;
            } else if (rbTitleBottom.value) {
                if (sizePt >= h) return;
                rectTop = B + sizePt; rectLeft = L; rectW = w; rectH = sizePt;
            } else if (rbTitleLeft.value) {
                if (sizePt >= w) return;
                rectTop = T; rectLeft = L; rectW = sizePt; rectH = h;
            } else {
                // 右
                if (sizePt >= w) return;
                rectTop = T; rectLeft = R - sizePt; rectW = sizePt; rectH = h;
            }

            var rr = pathItem.layer.pathItems.rectangle(rectTop, rectLeft, rectW, rectH);
            rr.stroked = false;
            rr.filled = true;
            rr.fillColor = makeK30Fill();
            try { rr.note = "__TitleFill__"; } catch (_) { }
            try { rr.name = "__TitleFill__"; } catch (_) { }
            // 背面へ（他の罫線や要素の下に敷く）
            try { rr.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
            tempPreviewItems.push(rr);
        } catch (_) { }
    }

    function createTitleDivider(target, titleSizePt, titleDistPt) {
        try {
            // タイトル帯の境界線（帯と本文の仕切り線）
            // titleDistPt:
            //   + → 両端を短くする
            //   - → 両端を伸ばす
            if (!target || target.typename !== "PathItem") return;
            if (!titleSizePt || titleSizePt === 0) return;
            if (!(chkTitleLine && chkTitleLine.value)) return;

            var b = target.geometricBounds; // [L, T, R, B]
            var L = b[0], T = b[1], R = b[2], B = b[3];

            var x1, y1, x2, y2;

            if (rbTitleTop && rbTitleTop.value) {
                y1 = y2 = T - titleSizePt;
                x1 = L + titleDistPt;
                x2 = R - titleDistPt;
                if (x1 >= x2) { x1 = L; x2 = R; }
            } else if (rbTitleBottom && rbTitleBottom.value) {
                y1 = y2 = B + titleSizePt;
                x1 = L + titleDistPt;
                x2 = R - titleDistPt;
                if (x1 >= x2) { x1 = L; x2 = R; }
            } else if (rbTitleLeft && rbTitleLeft.value) {
                x1 = x2 = L + titleSizePt;
                y1 = T - titleDistPt;
                y2 = B + titleDistPt;
            } else if (rbTitleRight && rbTitleRight.value) {
                x1 = x2 = R - titleSizePt;
                y1 = T - titleDistPt;
                y2 = B + titleDistPt;
            } else {
                y1 = y2 = T - titleSizePt;
                x1 = L + titleDistPt;
                x2 = R - titleDistPt;
                if (x1 >= x2) { x1 = L; x2 = R; }
            }

            var p = doc.activeLayer.pathItems.add();
            p.stroked = true;
            p.filled = false;

            // 黒 1pt
            try {
                var k = new CMYKColor();
                k.cyan = 0; k.magenta = 0; k.yellow = 0; k.black = 100;
                p.strokeColor = k;
            } catch (_) { }
            try { p.strokeWidth = 1; } catch (_) { }

            p.setEntirePath([[x1, y1], [x2, y2]]);

            try { p.note = "__TitleDivider__"; } catch (_) { }
            try { p.name = "__TitleDivider__"; } catch (_) { }

            try { tempPreviewItems.push(p); } catch (_) { }
        } catch (_) { }
    }

    // タイトル領域を除外した「内側罫線」用の計算領域を返す
    // 戻り値: [L, T, R, B] / 不成立の場合は null
    function getInnerAreaBounds(pathItem, titleSizePt) {
        try {
            var b = pathItem.geometricBounds; // [L, T, R, B]
            var L = b[0], T = b[1], R = b[2], B = b[3];
            var w = R - L;
            var h = T - B;
            if (w <= 0 || h <= 0) return null;

            var s = (titleSizePt && titleSizePt > 0) ? titleSizePt : 0;
            if (s <= 0) return [L, T, R, B];

            // タイトル領域ぶんを相殺
            if (rbTitleTop.value) {
                if (s >= h) return null;
                T = T - s;
            } else if (rbTitleBottom.value) {
                if (s >= h) return null;
                B = B + s;
            } else if (rbTitleLeft.value) {
                if (s >= w) return null;
                L = L + s;
            } else {
                // 右
                if (s >= w) return null;
                R = R - s;
            }

            // 念のため再チェック
            if ((R - L) <= 0 || (T - B) <= 0) return null;
            return [L, T, R, B];
        } catch (_) {
            return null;
        }
    }




    // 選択矩形から内側の矩形（inner box）を作成して登録
    // baseBounds: [L,T,R,B] or null (if null, uses pathItem.geometricBounds)
    function createInnerBox(pathItem, baseBounds, offTopPt, offBottomPt, offLeftPt, offRightPt, columns, rows, colGutterPt, rowGutterPt) {
        // オフセットが0でも内側エリアは描画する
        if (!offTopPt || offTopPt < 0) offTopPt = 0;
        if (!offBottomPt || offBottomPt < 0) offBottomPt = 0;
        if (!offLeftPt || offLeftPt < 0) offLeftPt = 0;
        if (!offRightPt || offRightPt < 0) offRightPt = 0;

        // 前提：軸に平行な長方形（geometricBounds を使用）
        // bounds: [left, top, right, bottom]
        var b = (baseBounds && baseBounds.length === 4) ? baseBounds : pathItem.geometricBounds;
        var left = b[0];
        var top = b[1];
        var right = b[2];
        var bottom = b[3];

        var w = right - left;
        var h = top - bottom;

        var newW = w - (offLeftPt + offRightPt);
        var newH = h - (offTopPt + offBottomPt);
        if (newW <= 0 || newH <= 0) return;

        var newLeft = left + offLeftPt;
        var newTop = top - offTopPt;

        // 列数/行数とガター（pt）
        var n = parseInt(columns, 10);
        if (isNaN(n) || n < 1) n = 1;
        var m = parseInt(rows, 10);
        if (isNaN(m) || m < 1) m = 1;

        var gX = (colGutterPt && colGutterPt > 0) ? colGutterPt : 0;
        var gY = (rowGutterPt && rowGutterPt > 0) ? rowGutterPt : 0;

        // 内側の長方形（塗り）
        // 列数/行数に応じて n×m に分割し、ガター分の空きを作る（n,mが3以上でも反映）
        function _addFillRect(topY, leftX, w, h) {
            if (!(w > 0) || !(h > 0)) return;
            var rr = pathItem.layer.pathItems.rectangle(topY, leftX, w, h);
            rr.stroked = false;
            rr.filled = true;
            rr.fillColor = makeK15Fill();
            try { rr.note = "__InnerBoxFill__"; } catch (_) { }
            try { rr.name = "__InnerBoxFill__"; } catch (_) { }
            // 背面へ（罫線などのパスの下に敷く）
            try { rr.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
            tempPreviewItems.push(rr);
        }

        var totalGX = gX * (n - 1);
        var totalGY = gY * (m - 1);
        var cellW = (n > 0) ? ((newW - totalGX) / n) : newW;
        var cellH = (m > 0) ? ((newH - totalGY) / m) : newH;
        if (!(cellW > 0) || !(cellH > 0)) return;

        for (var rIdx0 = 0; rIdx0 < m; rIdx0++) {
            var topY = newTop - (cellH * rIdx0) - (gY * rIdx0);
            for (var cIdx0 = 0; cIdx0 < n; cIdx0++) {
                var leftX = newLeft + (cellW * cIdx0) + (gX * cIdx0);
                _addFillRect(topY, leftX, cellW, cellH);
            }
        }

        // 罫線の見た目：基本は元オブジェクトの線を踏襲。なければ K100 / 1pt
        var lineColor;
        var lineWidth;
        try {
            if (pathItem.stroked) {
                lineColor = pathItem.strokeColor;
                lineWidth = pathItem.strokeWidth;
            }
        } catch (_) { }
        if (!lineColor) {
            var k = new CMYKColor();
            k.cyan = 0; k.magenta = 0; k.yellow = 0; k.black = 100;
            lineColor = k;
        }
        if (!lineWidth) lineWidth = 1;

        var bottomY = newTop - newH;
        var rightX = newLeft + newW;

        // 列（垂直）分割の罫線（columns >= 2 のとき） + 列のガター
        // 列=2 のときは「中間に1本」（ガター中心）
        if (n >= 2) {
            if (n === 2) {
                if (chkRowDivider && chkRowDivider.enabled && chkRowDivider.value) {
                    var cellW2b = (newW - gX) / 2;
                    if (cellW2b > 0) {
                        var xMid = newLeft + cellW2b + (gX / 2);
                        var lnV = pathItem.layer.pathItems.add();
                        lnV.setEntirePath([[xMid, newTop], [xMid, bottomY]]);
                        lnV.stroked = true;
                        lnV.filled = false;
                        lnV.strokeColor = lineColor;
                        lnV.strokeWidth = lineWidth;
                        applyInnerLineStyle(lnV);
                        tempPreviewItems.push(lnV);
                    }
                }
            } else {
                // ガター幅（gX）を含めて等分。区切り線は各ガターの中心（gX/2）に1本だけ。
                var totalGX = gX * (n - 1);
                var cellW = (newW - totalGX) / n;
                if (!(cellW > 0)) return;

                if (chkRowDivider && chkRowDivider.enabled && chkRowDivider.value) {
                    for (var c = 1; c <= n - 1; c++) {
                        // 境界（セル右端）位置
                        var xBoundary = newLeft + (cellW * c) + (gX * (c - 1));
                        // ガター中心（gX=0なら境界そのもの）
                        var xMid = xBoundary + (gX / 2);

                        var lnV2 = pathItem.layer.pathItems.add();
                        lnV2.setEntirePath([[xMid, newTop], [xMid, bottomY]]);
                        lnV2.stroked = true;
                        lnV2.filled = false;
                        lnV2.strokeColor = lineColor;
                        lnV2.strokeWidth = lineWidth;
                        applyInnerLineStyle(lnV2);
                        tempPreviewItems.push(lnV2);
                    }
                }
            }
        }

        // 行（水平）分割の罫線（rows >= 2 のとき） + 行のガター
        // 行=2 のときは「中間に1本」（ガター中心）
        if (m >= 2) {
            if (m === 2) {
                if (chkRowDivider && chkRowDivider.enabled && chkRowDivider.value) {
                    var cellH2b = (newH - gY) / 2;
                    if (cellH2b > 0) {
                        var yMid = newTop - cellH2b - (gY / 2);
                        var lnH = pathItem.layer.pathItems.add();
                        lnH.setEntirePath([[newLeft, yMid], [rightX, yMid]]);
                        lnH.stroked = true;
                        lnH.filled = false;
                        lnH.strokeColor = lineColor;
                        lnH.strokeWidth = lineWidth;
                        applyInnerLineStyle(lnH);
                        tempPreviewItems.push(lnH);
                    }
                }
            } else {
                // ガター幅（gY）を含めて等分。区切り線は各ガターの中心（gY/2）に1本だけ。
                var totalGY = gY * (m - 1);
                var cellH = (newH - totalGY) / m;
                if (!(cellH > 0)) return;

                if (chkRowDivider && chkRowDivider.enabled && chkRowDivider.value) {
                    for (var rIdx = 1; rIdx <= m - 1; rIdx++) {
                        // 境界（セル下端）位置
                        var yBoundary = newTop - (cellH * rIdx) - (gY * (rIdx - 1));
                        // ガター中心（gY=0なら境界そのもの）
                        var yMid = yBoundary - (gY / 2);

                        var lnH2 = pathItem.layer.pathItems.add();
                        lnH2.setEntirePath([[newLeft, yMid], [rightX, yMid]]);
                        lnH2.stroked = true;
                        lnH2.filled = false;
                        lnH2.strokeColor = lineColor;
                        lnH2.strokeWidth = lineWidth;
                        applyInnerLineStyle(lnH2);
                        tempPreviewItems.push(lnH2);
                    }
                }
            }
        }
    }

    // プレビューを描画する関数
    function updatePreview(isFinal) {
        // 先に前回のプレビューを削除
        removeTempItems();

        // アートボード基準のときは、マージンを反映した一時矩形に更新（裁ち落としはフレームのみに適用）
        if (_usingArtboardBase) {

            if (_usingArtboardBase) {
                var mtVal = parseFloat(editArtboardMarginTop.text);
                var mbVal = parseFloat(editArtboardMarginBottom.text);
                var mlVal = parseFloat(editArtboardMarginLeft.text);
                var mrVal = parseFloat(editArtboardMarginRight.text);

                if (isNaN(mtVal) || mtVal < 0) mtVal = 0;
                if (isNaN(mbVal) || mbVal < 0) mbVal = 0;
                if (isNaN(mlVal) || mlVal < 0) mlVal = 0;
                if (isNaN(mrVal) || mrVal < 0) mrVal = 0;

                var factorM = getCurrentRulerPtFactor();
                var mtPt = mtVal * factorM;
                var mbPt = mbVal * factorM;
                var mlPt = mlVal * factorM;
                var mrPt = mrVal * factorM;

                try { if (typeof chkBleed !== "undefined") _bleedEnabled = !!chkBleed.value; } catch (_) { }
                rebuildArtboardBaseRect(mtPt, mrPt, mbPt, mlPt);
            }
        }

        // 入力値チェック
        var factor = getCurrentRulerPtFactor();

        var val = getEffectiveLenValue();
        var distPt = val * factor; // rulerType -> pt

        // タイトルエリア：辺の伸縮（タイトル帯の線の長さにのみ反映）
        var titleLenVal = 0;
        try {
            if (chkTitleLine && chkTitleLine.value && chkTitleEdgeScale && chkTitleEdgeScale.value) {
                var t = parseFloat(editTitleEdgeScale && editTitleEdgeScale.text);
                if (!isNaN(t)) titleLenVal = t;
            }
        } catch (_) { }
        var titleDistPt = (-titleLenVal) * factor; // 正負を反転

        // タイトル領域サイズ（外枠基準）
        var titleVal = (typeof chkTitleEnable !== "undefined" && chkTitleEnable && !chkTitleEnable.value) ? 0 : parseFloat(editTitleSize && editTitleSize.text);
        var titleSizePt = (!isNaN(titleVal) ? (titleVal * factor) : 0);

        // フレーム幅（pt）
        // 裁ち落としONのときは 3mm をフレーム幅に加算（例：4mm + 3mm = 7mm）
        var frameVal = (chkFrameEnable && chkFrameEnable.value) ? parseFloat(editFrameWidth.text) : 0;
        var framePt = (!isNaN(frameVal) ? (frameVal * factor) : 0);
        if (framePt < 0) framePt = 0;
        if (_bleedEnabled) {
            var bleedPtForFrame = (72.0 / 25.4) * BLEED_MM;
            framePt += bleedPtForFrame;
        }

        var offTopVal = parseFloat(editInnerOffsetTop.text);
        var offBottomVal = parseFloat(editInnerOffsetBottom.text);
        var offLeftVal = parseFloat(editInnerOffsetLeft.text);
        var offRightVal = parseFloat(editInnerOffsetRight.text);

        var offTopPt = (!isNaN(offTopVal) ? (offTopVal * factor) : 0);
        var offBottomPt = (!isNaN(offBottomVal) ? (offBottomVal * factor) : 0);
        var offLeftPt = (!isNaN(offLeftVal) ? (offLeftVal * factor) : 0);
        var offRightPt = (!isNaN(offRightVal) ? (offRightVal * factor) : 0);

        var colVal = parseInt(editInnerColumns.text, 10);
        if (isNaN(colVal) || colVal < 1) colVal = 1;
        var rowVal = parseInt(editInnerRows.text, 10);
        if (isNaN(rowVal) || rowVal < 1) rowVal = 1;
        try { applyRowDividerEnabledState(colVal, rowVal, false); } catch (_) { }

        // ガター（列/行）：列/行が1なら無効（0扱い）
        try { editColGutter.enabled = (colVal > 1); } catch (_) { }
        try { editRowGutter.enabled = (rowVal > 1); } catch (_) { }

        var colGutterVal = parseFloat(editColGutter.text);
        var rowGutterVal = parseFloat(editRowGutter.text);
        var colGutterPt = (colVal > 1 && !isNaN(colGutterVal)) ? (colGutterVal * factor) : 0;
        var rowGutterPt = (rowVal > 1 && !isNaN(rowGutterVal)) ? (rowGutterVal * factor) : 0;
        if (colGutterPt < 0) colGutterPt = 0;
        if (rowGutterPt < 0) rowGutterPt = 0;

        // ガターが設定されたら塗りを自動ON（ただし手動操作があれば尊重）
        if (!_rowFillManuallySet && ((colGutterPt && colGutterPt !== 0) || (rowGutterPt && rowGutterPt !== 0))) {
            try { if (!chkRowFill.value) chkRowFill.value = true; } catch (_) { }
        }

        // 0 のときは分割しない（元のオブジェクトを表示したまま）
        if (distPt === 0) {
            // フレーム（アートボードサイズ基準）
            if (framePt > 0) {
                var abB = getActiveArtboardBounds();
                if (abB) {
                    // 代表として targetItems[0] のレイヤーに作成する
                    createFrameFill(targetItems[0], framePt, abB);
                }
            }
            // タイトル帯（塗り）
            if (titleSizePt > 0) {
                for (var i = 0; i < targetItems.length; i++) {
                    createTitleFill(targetItems[i], titleSizePt);
                }
            }
            // タイトル帯の分割線
            if (titleSizePt > 0) {
                for (var i = 0; i < targetItems.length; i++) {
                    createTitleDivider(targetItems[i], titleSizePt, titleDistPt);
                }
            }
            // 線端パネルはディム表示
            try { capPanel.enabled = false; } catch (_) { }

            // 外枠の表示は「外枠を残す」に従う（辺の伸縮のON/OFFには依存しない）
            var showOuterRect = !!chkKeepOuter.value;
            for (var i = 0; i < targetItems.length; i++) {
                try { targetItems[i].hidden = !showOuterRect; } catch (_) { }
            }

            // 外側エリア：角丸（辺の伸縮OFFのときのみ）
            // フレーム角丸と同様に、プレビュー時は新規の一時オブジェクトへ適用して元を汚さない
            try {
                var showOuterRect = !!chkKeepOuter.value;

                var useOuterRound = (showOuterRect
                    && chkOuterRound && chkOuterRound.value
                    && chkEnableLen && !chkEnableLen.value);

                var rVal = parseFloat(editOuterRound.text);
                if (isNaN(rVal) || rVal <= 0) useOuterRound = false;

                if (useOuterRound) {
                    var rPt = rVal * getCurrentRulerPtFactor();
                    if (rPt > 0) {
                        if (isFinal) {
                            // OK時：元外枠に適用（ライブエフェクト）
                            for (var i = 0; i < targetItems.length; i++) {
                                try {
                                    var it = targetItems[i];
                                    if (!it || it.typename !== "PathItem") continue;
                                    applyRoundCornersEffect(it, rPt);
                                    it.hidden = false;
                                } catch (_) { }
                            }
                        } else {
                            // プレビュー時：元外枠を隠し、同じboundsの矩形を作って角丸を当てる
                            for (var i2 = 0; i2 < targetItems.length; i2++) {
                                try {
                                    var src = targetItems[i2];
                                    if (!src || src.typename !== "PathItem") continue;

                                    var b = src.geometricBounds; // [L,T,R,B]
                                    var L = b[0], T = b[1], R = b[2], B = b[3];
                                    var w = R - L, h = T - B;
                                    if (!(w > 0) || !(h > 0)) continue;

                                    try { src.hidden = true; } catch (_) { }

                                    var rr = src.layer.pathItems.rectangle(T, L, w, h);
                                    // 見た目をsrcに寄せる
                                    try { rr.stroked = src.stroked; } catch (_) { }
                                    try { rr.filled = src.filled; } catch (_) { }
                                    try { rr.strokeColor = src.strokeColor; } catch (_) { }
                                    try { rr.fillColor = src.fillColor; } catch (_) { }
                                    try { rr.strokeWidth = src.strokeWidth; } catch (_) { }

                                    try { rr.note = "__OuterRoundPreview__"; } catch (_) { }
                                    try { rr.name = "__OuterRoundPreview__"; } catch (_) { }

                                    applyRoundCornersEffect(rr, rPt);
                                    tempPreviewItems.push(rr);
                                } catch (_) { }
                            }
                        }
                    }
                }
            } catch (_) { }

            // 内側エリア（オフセットが0でも描画）
            for (var i = 0; i < targetItems.length; i++) {
                var innerB = getInnerAreaBounds(targetItems[i], titleSizePt);
                if (innerB) createInnerBox(targetItems[i], innerB, offTopPt, offBottomPt, offLeftPt, offRightPt, colVal, rowVal, colGutterPt, rowGutterPt);
            }

            app.redraw();
            return;
        }

        // 0 以外は線端パネルを有効化
        try { capPanel.enabled = true; } catch (_) { }

        // 分解した4辺を表示するため、元の長方形は常に隠す
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].hidden = true;
        }

        // 外枠を残すONのときのみ4辺罫線を描画
        if (chkKeepOuter.value) {
            for (var i = 0; i < targetItems.length; i++) {
                createShortenedLines(targetItems[i], distPt);
            }
        }

        // フレーム（アートボードサイズ基準）
        if (framePt > 0) {
            var abB2 = getActiveArtboardBounds();
            if (abB2) {
                createFrameFill(targetItems[0], framePt, abB2);
            }
        }

        // タイトル帯（塗り）
        if (titleSizePt > 0) {
            for (var i = 0; i < targetItems.length; i++) {
                createTitleFill(targetItems[i], titleSizePt);
            }
        }

        // タイトル帯の分割線
        if (titleSizePt > 0) {
            for (var i = 0; i < targetItems.length; i++) {
                createTitleDivider(targetItems[i], titleSizePt, titleDistPt);
            }
        }

        // 内側エリア（オフセットが0でも描画）
        for (var i = 0; i < targetItems.length; i++) {
            var innerB2 = getInnerAreaBounds(targetItems[i], titleSizePt);
            if (innerB2) createInnerBox(targetItems[i], innerB2, offTopPt, offBottomPt, offLeftPt, offRightPt, colVal, rowVal, colGutterPt, rowGutterPt);
        }

        app.redraw(); // 画面を強制再描画（重要）
    }

    // プレビューを消して元に戻す関数
    function clearPreview() {
        removeTempItems();

        // 外側角丸プレビューで隠した元外枠を復帰
        try {
            for (var i = 0; i < targetItems.length; i++) {
                if (targetItems[i]) targetItems[i].hidden = false;
            }
        } catch (_) { }

        // アートボード基準の一時矩形は先に破棄（targetItemsに残っていると無効参照になる）
        cleanupArtboardBaseRect();

        // 元のオブジェクトを表示に戻す
        for (var i = 0; i < targetItems.length; i++) {
            try { targetItems[i].hidden = false; } catch (_) { }
        }
        try { capPanel.enabled = true; } catch (_) { }

        app.redraw();
    }

    // 一時アイテムを削除する関数
    function removeTempItems() {
        for (var i = 0; i < tempPreviewItems.length; i++) {
            try {
                tempPreviewItems[i].remove();
            } catch (e) {
                // エラー無視（すでに消えている場合など）
            }
        }
        tempPreviewItems = [];
    }

    // 線生成ロジック
    function createShortenedLines(pathItem, dist) {
        var points = pathItem.pathPoints;
        var isClosed = pathItem.closed;
        var limit = isClosed ? points.length : points.length - 1;

        for (var j = 0; j < limit; j++) {
            var p1 = points[j].anchor;
            var p2_index = (j + 1) % points.length;
            var p2 = points[p2_index].anchor;

            var dx = p2[0] - p1[0];
            var dy = p2[1] - p1[1];
            var currentLength = Math.sqrt(dx * dx + dy * dy);

            var absDist = Math.abs(dist);

            // 短くなりすぎて消滅する場合は作らない（伸ばす場合は制限しない）
            if (dist > 0 && currentLength <= absDist * 2) continue;

            var ratio = absDist / currentLength;

            var newX1, newY1, newX2, newY2;
            if (dist >= 0) {
                // 伸ばす（正の値）
                newX1 = p1[0] - dx * ratio;
                newY1 = p1[1] - dy * ratio;
                newX2 = p2[0] + dx * ratio;
                newY2 = p2[1] + dy * ratio;
            } else {
                // 短くする（負の値）
                newX1 = p1[0] + dx * ratio;
                newY1 = p1[1] + dy * ratio;
                newX2 = p2[0] - dx * ratio;
                newY2 = p2[1] - dy * ratio;
            }

            // 線を作成して tempPreviewItems に登録
            var newLine = pathItem.layer.pathItems.add();
            newLine.setEntirePath([[newX1, newY1], [newX2, newY2]]);

            newLine.stroked = true;
            newLine.filled = false;
            newLine.strokeColor = pathItem.strokeColor;
            newLine.strokeWidth = pathItem.strokeWidth;
            try {
                newLine.strokeCap = getSelectedStrokeCap();
            } catch (_) { }

            // 外枠（4辺）として識別できるようタグ
            try { newLine.note = "__OuterEdge__"; } catch (_) { }
            try { newLine.name = "__OuterEdge__"; } catch (_) { }

            tempPreviewItems.push(newLine);
        }
    }

})();
