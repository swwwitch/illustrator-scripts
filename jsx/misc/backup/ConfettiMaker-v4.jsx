#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* =========================================
 * 紙吹雪を生成 / Generate Confetti
 *
 * 概要 / Overview
 * - バージョン / Version: v1.3.1
 * - 更新日 / Updated: 2026-02-18
 * - 選択オブジェクトの領域を基準に、紙吹雪（円/長方形/三角形/スター/キラキラ）を生成します。
 * - ダイアログ上でプレビューを表示し、OKで確定（Confetti レイヤーに出力）します。
 * - UIは日本語/英語に対応し、タイトルバーにバージョンを表示します。
 *
 * 使い方 / Usage
 * 1) マスクにしたいオブジェクトを1つ選択
 * 2) スクリプトを実行
 * 3) 設定を調整し、OK
 *
 * メモ / Notes
 * - マージンは生成エリアのみ拡張します（マスク領域は拡張しません）。
 * - TextFrame（テキスト）選択時はマスク処理を無効化します。
 * - プレビュー用レイヤー（__ConfettiPreview__）は終了時に削除されます。
 * - 既に Confetti レイヤーがある場合は新規作成せず再利用します。
 * - マージンの最大値は（幅+高さ）/6 を基準に動的に設定します。
 * ========================================= */

// コンフェティ（紙吹雪）作成スクリプト

(function () {
    var SCRIPT_VERSION = "v1.3.1";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "紙吹雪を生成", en: "Generate Confetti" },

        pnlShape: { ja: "形状", en: "Shapes" },
        circle: { ja: "円", en: "Circle" },
        rect: { ja: "長方形", en: "Rectangle" },
        triangle: { ja: "三角形", en: "Triangle" },
        star: { ja: "スター", en: "Star" },
        sparkle: { ja: "キラキラ", en: "Sparkle" },
        ribbon: { ja: "リボン", en: "Ribbon" },

        pnlDist: { ja: "配置分布", en: "Distribution" },
        distEven: { ja: "全体に均等", en: "Uniform" },
        distGrad: { ja: "垂直方向", en: "Vertical" },
        distHollow: { ja: "放射状", en: "Radial" },
        distStrength: { ja: "強度", en: "Strength" },
        distLabel: { ja: "分布", en: "Distribution" },

        pnlCount: { ja: "生成数", en: "Count" },

        // --- Base size panel labels ---
        pnlBaseSize: { ja: "基準サイズ", en: "Base size" },
        pnlBasic: { ja: "基本設定", en: "Basic" },
        baseSize: { ja: "基準", en: "Base" },
        ptUnit: { ja: "pt", en: "pt" },

        pnlOption: { ja: "オプション", en: "Options" },
        pnlSymbol: { ja: "シンボル", en: "Symbol" },
        pnlRandom: { ja: "ランダム", en: "Random" },
        mask: { ja: "マスク処理", en: "Mask" },
        margin: { ja: "マージン", en: "Margin" },
        random: { ja: "大きさ", en: "Size" },
        zoom: { ja: "画面ズーム", en: "Zoom" },
        opacity: { ja: "不透明度", en: "Opacity" },
        skew: { ja: "歪み", en: "Skew" },
        btnCancel: { ja: "キャンセル", en: "Cancel" },
        btnOK: { ja: "OK", en: "OK" },
        // Additions for Symbol dropdown
        symbol: { ja: "シンボル", en: "Symbol" },
        symbolNone: { ja: "（なし）", en: "(None)" }
    };

    function L(key) {
        try {
            if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
            if (LABELS[key] && LABELS[key].ja) return LABELS[key].ja;
        } catch (_) { }
        return key;
    }
    var doc = app.activeDocument;

    // 選択オブジェクト取得（なければアートボードを対象）
    var selectedObj = null;
    var __useArtboard = false;

    if (doc.selection.length > 0) {
        selectedObj = doc.selection[0];
    } else {
        __useArtboard = true;
    }

    // テキスト選択時フラグ / Flag for text selection
    var __isTextSelection = false;
    try { __isTextSelection = (selectedObj && selectedObj.typename === "TextFrame"); } catch (_) { __isTextSelection = false; }

    // ビュー初期値（ズーム復元用） / Initial view state
    var __initView = null;
    var __initZoom = null;
    var __initCenter = null;
    try {
        __initView = doc.activeView;
        __initZoom = __initView.zoom;
        __initCenter = __initView.centerPoint;
    } catch (_) { }

    // ダイアログ作成
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = "left";
    try { dlg.spacing = 15; } catch (_) { }

    // ダイアログの位置・透明度 / Dialog position & opacity
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, dx, dy) {
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + dx, currentY + dy];
        } catch (_) { }
    }

    function setDialogOpacity(dlg, opacityValue) {
        try { dlg.opacity = opacityValue; } catch (_) { }
    }

    // チェックボックスのラベル幅を揃える / Unify checkbox label widths
    function unifyCheckboxLabelWidth(chkList, minWidth) {
        if (!chkList || !chkList.length) return;
        var w = (typeof minWidth === "number") ? minWidth : 0;

        for (var i = 0; i < chkList.length; i++) {
            var c = chkList[i];
            if (!c) continue;
            try {
                var t = (c.text != null) ? String(c.text) : "";
                var gw = 0;
                try {
                    // measureString が使える場合は実測
                    gw = c.graphics.measureString(t).width;
                } catch (_) {
                    // フォールバック: 文字数ベース
                    gw = t.length * 7;
                }
                // 余白を足す
                var ww = Math.ceil(gw + 18);
                if (ww > w) w = ww;
            } catch (_) { }
        }

        for (var j = 0; j < chkList.length; j++) {
            var c2 = chkList[j];
            if (!c2) continue;
            try { c2.preferredSize.width = w; } catch (_) { }
        }
    }

    setDialogOpacity(dlg, dialogOpacity);

    /* 基本設定 / Basic（ダイアログ最上部・全幅） */
    var pnlBaseSize = dlg.add("panel", undefined, L("pnlBasic"));
    pnlBaseSize.orientation = "column";
    pnlBaseSize.alignChildren = "fill";
    pnlBaseSize.margins = [15, 20, 15, 10];

    // 1行構成：ラベル + スライダー
    var gBaseRow = pnlBaseSize.add("group");
    gBaseRow.orientation = "row";
    gBaseRow.alignChildren = ["left", "center"];

    var stBaseSizeLabel = gBaseRow.add("statictext", undefined, L("pnlBaseSize"));

    // baseSizePt（例：6）がスライダーの 0 に対応（相対調整）
    var baseSizePt = 6;
    var sldBaseSize = gBaseRow.add("slider", undefined, 0, -5, 45); // 0 => 6pt, -5..45 => 1..51pt
    try { sldBaseSize.preferredSize.width = 240; } catch (_) { }

    function getBaseSizePt() {
        var v = baseSizePt + Number(sldBaseSize.value);
        if (isNaN(v)) v = baseSizePt;
        if (v < 0.1) v = 0.1;
        // 見た目は 0.1pt 刻み
        v = Math.round(v * 10) / 10;
        return v;
    }

    sldBaseSize.onChanging = function () {
        requestPreviewDebounced();
    };

    /* 生成数 / Count（基本設定panel内） */
    var gCountRow = pnlBaseSize.add("group");
    gCountRow.orientation = "row";
    gCountRow.alignChildren = ["left", "center"];

    var stCountLabel = gCountRow.add("statictext", undefined, L("pnlCount"));
    // ラベル幅を揃える（基準サイズ / 生成数）
    try {
        var maxLabelWidth = Math.max(
            stBaseSizeLabel.preferredSize.width || 0,
            stCountLabel.preferredSize.width || 0
        );
        stBaseSizeLabel.preferredSize.width = maxLabelWidth;
        stCountLabel.preferredSize.width = maxLabelWidth;
    } catch (_) { }

    var confettiCount = 150;
    var sldCount = gCountRow.add("slider", undefined, 150, 10, 500);
    try { sldCount.preferredSize.width = 240; } catch (_) { }

    sldCount.onChanging = function () {
        confettiCount = Math.round(sldCount.value);
        requestPreviewDebounced();
    };


    /* マスク処理 / Mask（基本設定panel内に移動） */
    // マスク処理 + マージン をまとめる
    var gMaskMargin = pnlBaseSize.add("group");
    gMaskMargin.orientation = "column";
    gMaskMargin.alignChildren = "left";
    try { gMaskMargin.margins = [0, 10, 0, 10]; } catch (_) { }

    var gMaskRow = gMaskMargin.add("group");

    gMaskRow.orientation = "row";
    gMaskRow.alignChildren = ["left", "center"];

    var chkMask = gMaskRow.add("checkbox", undefined, L("mask"));
    chkMask.value = true;
    try {
        if (__isTextSelection || __useArtboard) {
            chkMask.value = false;
            chkMask.enabled = false;
        }
    } catch (_) { }

    /* マージン（生成エリア用） */
    var gMarginRow = gMaskMargin.add("group");
    gMarginRow.orientation = "row";
    gMarginRow.alignChildren = ["left", "center"];

    var chkMargin = gMarginRow.add("checkbox", undefined, L("margin"));
    chkMargin.value = false;

    var sldMargin = gMarginRow.add("slider", undefined, 0, 0, 50);
    try { sldMargin.preferredSize.width = 240; } catch (_) { }
    sldMargin.enabled = chkMargin.value;

    var maskMarginPt = 0;

    // マージン最大値を対象サイズから設定
    try { applyMarginMaxToUI(); } catch (_) { }


    /* 配置分布 / Distribution（基本設定panel内・group） */
    var gDistWrap = pnlBaseSize.add("group");
    gDistWrap.orientation = "column";
    gDistWrap.alignChildren = "left";

    var gDistRow = gDistWrap.add("group");
    gDistRow.orientation = "row";
    gDistRow.alignChildren = ["left", "center"];

    var stDistLabel = gDistRow.add("statictext", undefined, L("distLabel"));

    var rbEven = gDistRow.add("radiobutton", undefined, L("distEven"));
    var rbGradY = gDistRow.add("radiobutton", undefined, L("distGrad"));
    var rbHollow = gDistRow.add("radiobutton", undefined, L("distHollow"));

    var gHollow = gDistWrap.add("group");
    gHollow.orientation = "row";
    gHollow.alignChildren = ["left", "center"];

    var stHollowLabel = gHollow.add("statictext", undefined, L("distStrength"));

    var sldHollow = gHollow.add("slider", undefined, 2.0, 1.0, 6.0);
    try { sldHollow.preferredSize.width = 200; } catch (_) { }

    var hollowStrength = 2.0;

    rbEven.value = true;
    rbHollow.value = false;
    rbGradY.value = false;

    gHollow.enabled = (rbHollow.value || rbGradY.value);


    /* 形状 / Shapes */
    var pnlShape = dlg.add("panel", undefined, L("pnlShape"));
    pnlShape.orientation = "column";
    pnlShape.alignChildren = "left";
    pnlShape.margins = [15, 20, 15, 10];
    try { pnlShape.alignment = ["fill", "top"]; } catch (_) { }

    // --- 3カラム構成 ---
    var gShapeCols = pnlShape.add("group");
    gShapeCols.orientation = "row";
    gShapeCols.alignChildren = ["left", "top"];

    // 左カラム
    var colLeft = gShapeCols.add("group");
    colLeft.orientation = "column";
    colLeft.alignChildren = "left";

    var chkCircle = colLeft.add("checkbox", undefined, L("circle"));
    var chkRect = colLeft.add("checkbox", undefined, L("rect"));

    // 中央カラム
    var colCenter = gShapeCols.add("group");
    colCenter.orientation = "column";
    colCenter.alignChildren = "left";

    var chkTriangle = colCenter.add("checkbox", undefined, L("triangle"));
    var chkRibbon = colCenter.add("checkbox", undefined, L("ribbon"));

    // 右カラム
    var colRight = gShapeCols.add("group");
    colRight.orientation = "column";
    colRight.alignChildren = "left";

    var chkStar = colRight.add("checkbox", undefined, L("star"));
    var chkStar4 = colRight.add("checkbox", undefined, L("sparkle"));

    // シンボル行（チェック + ドロップダウン）
    var gSymbolRow = pnlShape.add("group");
    gSymbolRow.orientation = "row";
    gSymbolRow.alignChildren = ["left", "center"];

    var chkSymbolShape = gSymbolRow.add("checkbox", undefined, L("symbol"));

    var ddSymbol = gSymbolRow.add("dropdownlist", undefined, [L("symbolNone")]);

    // 形状ラベル幅を揃える
    try {
        unifyCheckboxLabelWidth([
            chkCircle,
            chkRect,
            chkTriangle,
            chkRibbon,
            chkStar,
            chkStar4,
            chkSymbolShape
        ]);
    } catch (_) { }
    ddSymbol.selection = 0;
    ddSymbol.preferredSize.width = 200;

    // デフォルトはすべてON
    chkCircle.value = true;
    chkRect.value = true;
    chkTriangle.value = true;
    chkStar.value = false;
    chkStar4.value = false;
    chkRibbon.value = false;
    chkSymbolShape.value = false;
    // 初期状態ではシンボル未選択のためディム
    chkSymbolShape.enabled = false;

    /* Option(Alt)+クリックで単独選択 / Option(Alt)+click to solo */
    function soloShape(activeChk) {
        try {
            chkCircle.value = (activeChk === chkCircle);
            chkRect.value = (activeChk === chkRect);
            chkTriangle.value = (activeChk === chkTriangle);
            chkStar.value = (activeChk === chkStar);
            chkStar4.value = (activeChk === chkStar4);
            chkRibbon.value = (activeChk === chkRibbon);
            chkSymbolShape.value = (activeChk === chkSymbolShape);
            activeChk.value = true;
        } catch (_) { }
    }

    // ⌘+Option(Alt)+クリックで「それ以外」を選択 / Cmd+Option(Alt)+click to select others
    function selectOthers(activeChk) {
        try {
            chkCircle.value = (activeChk !== chkCircle);
            chkRect.value = (activeChk !== chkRect);
            chkTriangle.value = (activeChk !== chkTriangle);
            chkStar.value = (activeChk !== chkStar);
            chkStar4.value = (activeChk !== chkStar4);
            chkRibbon.value = (activeChk !== chkRibbon);
            chkSymbolShape.value = (activeChk !== chkSymbolShape);
        } catch (_) { }
    }

    // --- 形状/シンボル相互排他ヘルパー ---
    function turnOffAllShapes() {
        try {
            chkCircle.value = false;
            chkRect.value = false;
            chkTriangle.value = false;
            chkStar.value = false;
            chkStar4.value = false;
            chkRibbon.value = false;
            chkSymbolShape.value = false;
        } catch (_) { }
    }

    function clearSymbolSelection() {
        try {
            if (ddSymbol) {
                ddSymbol.selection = 0;
            }
        } catch (_) { }
        selectedSymbolRef = null;
    }

    function isAltDown() {
        try {
            return !!(ScriptUI.environment && ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey);
        } catch (_) { }
        return false;
    }

    function isCmdAltDown() {
        try {
            var ks = ScriptUI.environment && ScriptUI.environment.keyboardState;
            // macOS: Command = metaKey
            return !!(ks && ks.metaKey && ks.altKey);
        } catch (_) { }
        return false;
    }

    chkCircle.onClick = function () {
        if (isCmdAltDown()) selectOthers(chkCircle);
        else if (isAltDown()) soloShape(chkCircle);
        drawPreview();
    };
    chkRect.onClick = function () {
        if (isCmdAltDown()) selectOthers(chkRect);
        else if (isAltDown()) soloShape(chkRect);
        drawPreview();
    };
    chkTriangle.onClick = function () {
        if (isCmdAltDown()) selectOthers(chkTriangle);
        else if (isAltDown()) soloShape(chkTriangle);
        drawPreview();
    };
    chkStar.onClick = function () {
        if (isCmdAltDown()) selectOthers(chkStar);
        else if (isAltDown()) soloShape(chkStar);
        drawPreview();
    };
    chkStar4.onClick = function () {
        if (isCmdAltDown()) selectOthers(chkStar4);
        else if (isAltDown()) soloShape(chkStar4);
        drawPreview();
    };
    chkRibbon.onClick = function () {
        if (isCmdAltDown()) selectOthers(chkRibbon);
        else if (isAltDown()) soloShape(chkRibbon);
        drawPreview();
    };
    chkSymbolShape.onClick = function () {
        if (isCmdAltDown()) selectOthers(chkSymbolShape);
        else if (isAltDown()) soloShape(chkSymbolShape);
        drawPreview();
    };

    // 現在選択中のシンボル名（ロジックは追って）
    var selectedSymbolName = "";

    function refreshSymbolDropdown() {
        try {
            ddSymbol.removeAll();
        } catch (_) { }

        var hasSymbols = false;
        try {
            if (doc && doc.symbols && doc.symbols.length > 0) hasSymbols = true;
        } catch (_) { hasSymbols = false; }

        // (None)
        try { ddSymbol.add("item", L("symbolNone")); } catch (_) { }

        if (hasSymbols) {
            for (var si2 = 0; si2 < doc.symbols.length; si2++) {
                try {
                    var s = doc.symbols[si2];
                    var nm = (s && s.name) ? String(s.name) : ("Symbol " + (si2 + 1));
                    var it = ddSymbol.add("item", nm);
                    try { it._sym = s; } catch (_) { }
                } catch (_) { }
            }
            ddSymbol.enabled = true;
        } else {
            ddSymbol.enabled = false;
        }

        try { ddSymbol.selection = 0; } catch (_) { }
        selectedSymbolName = "";
        selectedSymbolRef = null;
        // シンボル未選択なので形状パネルの「シンボル」はディム
        try {
            chkSymbolShape.value = false;
            chkSymbolShape.enabled = false;
        } catch (_) { }
    }

    ddSymbol.onChange = function () {
        try {
            if (!ddSymbol.selection || ddSymbol.selection.index === 0) {
                selectedSymbolName = "";
                // None → ディム + OFF
                try {
                    chkSymbolShape.value = false;
                    chkSymbolShape.enabled = false;
                } catch (_) { }
                drawPreview();
                return;
            }

            // 実在するシンボルが選択された
            selectedSymbolName = String(ddSymbol.selection.text);

            try {
                chkSymbolShape.enabled = true;
                chkSymbolShape.value = true; // 自動でON
            } catch (_) { }
        } catch (_) {
            selectedSymbolName = "";
        }

        drawPreview();
    };

    /* ランダム / Random */
    var pnlRandom = dlg.add("panel", undefined, L("pnlRandom"));
    pnlRandom.orientation = "column";
    pnlRandom.alignChildren = "left";
    pnlRandom.margins = [15, 20, 15, 10];

    // 大きさ（□大きさ <=====>）
    var gRandSize = pnlRandom.add("group");
    gRandSize.orientation = "row";
    gRandSize.alignChildren = ["left", "center"];

    var chkRandSize = gRandSize.add("checkbox", undefined, L("random"));
    chkRandSize.value = true; // デフォルトON（現状の挙動を維持）

    var sldRandSize = gRandSize.add("slider", undefined, 100, 100, 300);
    sldRandSize.preferredSize.width = 235;
    sldRandSize.enabled = chkRandSize.value;

    var randSizeStrength = 100;

    // 不透明度（□不透明度 <=====>）
    var gOpacity = pnlRandom.add("group");
    gOpacity.orientation = "row";
    gOpacity.alignChildren = ["left", "center"];

    var chkOpacity = gOpacity.add("checkbox", undefined, L("opacity"));
    chkOpacity.value = true;

    // 70..100 を指定（ON時のみランダムに適用）
    var sldOpacity = gOpacity.add("slider", undefined, 30, 0, 100); // reversed: value = (100 - opacityMin)
    sldOpacity.preferredSize.width = 235;
    sldOpacity.enabled = chkOpacity.value;

    var opacityMin = 70;
    try { sldOpacity.value = 100 - opacityMin; } catch (_) { }

    // 歪み（□歪み <=====>）
    var gSkew = pnlRandom.add("group");
    gSkew.orientation = "row";
    gSkew.alignChildren = ["left", "center"];

    var chkSkew = gSkew.add("checkbox", undefined, L("skew"));
    chkSkew.value = false; // デフォルトOFF（既存の見た目を変えない）

    // 最大歪み角度（度）: 0..45
    var sldSkew = gSkew.add("slider", undefined, 0, 0, 45);
    sldSkew.preferredSize.width = 235;
    sldSkew.enabled = chkSkew.value;

    var skewMaxDeg = 0;



    /* ズーム / Zoom */
    var gZoom = dlg.add("group");
    gZoom.orientation = "row";
    gZoom.alignChildren = ["center", "center"];
    gZoom.alignment = "center";
    // 画面ズームの下に余白を追加
    try { gZoom.margins = [0, 0, 0, 10]; } catch (_) { }

    var stZoom = gZoom.add("statictext", undefined, L("zoom"));
    var sldZoom = gZoom.add("slider", undefined, (__initZoom != null ? __initZoom : 1), 0.1, 16);
    sldZoom.preferredSize.width = 240;

    function applyZoom(z) {
        try {
            if (!__initView) __initView = doc.activeView;
            if (!__initView) return;
            __initView.zoom = z;
        } catch (_) { }
    }

    sldZoom.onChanging = function () {
        applyZoom(Number(sldZoom.value));
        try { app.redraw(); } catch (_) { }
    };

    /* OK / Cancel */
    var gBtns = dlg.add("group");
    gBtns.alignment = "right";
    gBtns.alignChildren = ["right", "center"];

    var btnCancel = gBtns.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var btnOK = gBtns.add("button", undefined, L("btnOK"), { name: "ok" });
    dlg.defaultElement = btnOK;

    // 有効なバウンディングボックスを取得する関数（マージン考慮）
    function getEffectiveBounds() {
        var b = null;

        try {
            if (__useArtboard) {
                var abIndex = doc.artboards.getActiveArtboardIndex();
                b = doc.artboards[abIndex].artboardRect; // [L, T, R, B]
            } else if (selectedObj) {
                b = selectedObj.geometricBounds; // [L, T, R, B]
            }
        } catch (e) {
            b = null;
        }

        if (!b) return null;

        var L = b[0], T = b[1], R = b[2], B = b[3];

        try {
            if (chkMargin && chkMargin.value && typeof maskMarginPt === "number" && maskMarginPt !== 0) {
                L -= maskMarginPt;
                T += maskMarginPt;
                R += maskMarginPt;
                B -= maskMarginPt;
            }
        } catch (_) { }

        return [L, T, R, B];
    }

    // max = (width + height) / 6
    function computeMaxMarginFromTarget() {
        var b = null;
        try {
            if (__useArtboard) {
                var abIndex = doc.artboards.getActiveArtboardIndex();
                b = doc.artboards[abIndex].artboardRect; // [L, T, R, B]
            } else if (selectedObj) {
                b = selectedObj.geometricBounds; // [L, T, R, B]
            }
        } catch (_) { b = null; }

        if (!b) return 50;

        var L = Number(b[0]), T = Number(b[1]), R = Number(b[2]), B = Number(b[3]);
        if (isNaN(L) || isNaN(T) || isNaN(R) || isNaN(B)) return 50;

        var w = Math.abs(R - L);
        var h = Math.abs(T - B);
        var maxM = (w + h) / 6;

        if (!maxM || isNaN(maxM) || maxM < 0) maxM = 50;
        if (maxM > 100000) maxM = 100000; // 念のため上限
        return maxM;
    }

    function applyMarginMaxToUI() {
        if (!sldMargin) return;
        var maxM = computeMaxMarginFromTarget();
        try { sldMargin.maxvalue = maxM; } catch (_) { }
        try { if (Number(sldMargin.value) > maxM) sldMargin.value = maxM; } catch (_) { }
        try { if (maskMarginPt > maxM) maskMarginPt = maxM; } catch (_) { }
    }

    // 設定
    // 旧 min/max は互換のため残すが、現在は基準サイズ（pt）を中心に計算する
    var minSize = 3; // (legacy) not used directly
    var maxSize = 8; // (legacy) not used directly

    // カラフルな色の配列
    var colors = [
        [255, 77, 77],    // 赤
        [255, 153, 51],   // オレンジ
        [255, 255, 51],   // 黄色
        [102, 255, 102],  // 緑
        [51, 153, 255],   // 青
        [153, 102, 255],  // 紫
        [255, 102, 178],  // ピンク
        [255, 204, 51]    // 金色
    ];

    // プレビュー用レイヤー（ダイアログ中のみ）
    var previewLayer = doc.layers.add();
    previewLayer.name = "__ConfettiPreview__";

    var previewItems = [];

    function clearPreview() {
        try {
            // 既存プレビューを完全に消す（グループ/マスク含む）
            if (previewLayer) {
                while (previewLayer.pageItems.length > 0) {
                    try { previewLayer.pageItems[0].remove(); } catch (e1) { break; }
                }
            }
        } catch (e2) { }
        previewItems = [];
        try { app.redraw(); } catch (_) { }
    }

    function drawPreview() {
        clearPreview();

        var mg = buildMaskGroup(previewLayer);
        var container = mg.container;

        var __b = getEffectiveBounds();
        if (!__b) return;

        for (var i = 0; i < confettiCount; i++) {
            var pt = pickPoint(__b);
            var x = pt.x;
            var y = pt.y;
            var size = getConfettiSize();
            var color = colors[Math.floor(Math.random() * colors.length)];
            var item = createConfetti(container, x, y, size, color);
            if (item) previewItems.push(item);
        }
        if (mg.group && !__useArtboard) {
            applyMaskToGroup(mg.group, selectedObj);
        }
        try { app.redraw(); } catch (_) { }
    }

    // =========================================
    // Debounced preview (間引き)
    // - Dragging sliders can fire many times; rebuild preview only once after a short pause.
    // =========================================
    var __debounceTaskId = 0;
    var __debounceDelayMs = 140; // 体感で軽くするための遅延（ms）

    // Expose drawPreview to scheduleTask (scheduleTask はグローバルスコープで動く)
    try {
        $.global.__ConfettiMaker_drawPreview = drawPreview;
        $.global.__ConfettiMaker_runDebouncedPreview = function () {
            try {
                if ($.global.__ConfettiMaker_drawPreview) $.global.__ConfettiMaker_drawPreview();
            } catch (_) { }
        };
    } catch (_) { }

    function requestPreviewDebounced() {
        // cancel previous scheduled task
        try {
            if (__debounceTaskId) {
                app.cancelTask(__debounceTaskId);
                __debounceTaskId = 0;
            }
        } catch (_) { }
        try {
            __debounceTaskId = app.scheduleTask("__ConfettiMaker_runDebouncedPreview()", __debounceDelayMs, false);
        } catch (_) {
            // fallback: run immediately
            try { drawPreview(); } catch (__) { }
        }
    }

    function requestPreviewImmediate() {
        try {
            if (__debounceTaskId) {
                try { app.cancelTask(__debounceTaskId); } catch (_) { }
                __debounceTaskId = 0;
            }
        } catch (_) { }
        try { drawPreview(); } catch (_) { }
    }

    rbEven.onClick = function () {
        gHollow.enabled = false;
        drawPreview();
    };
    rbHollow.onClick = function () {
        gHollow.enabled = true;
        drawPreview();
    };
    rbGradY.onClick = function () {
        gHollow.enabled = true;
        drawPreview();
    };
    chkMask.onClick = function () { drawPreview(); };

    chkOpacity.onClick = function () {
        sldOpacity.enabled = chkOpacity.value;
        drawPreview();
    };

    sldOpacity.onChanging = function () {
        // reversed: slider 0..100 => opacityMin 100..0
        opacityMin = 100 - Math.round(sldOpacity.value);
        if (opacityMin < 0) opacityMin = 0;
        if (opacityMin > 100) opacityMin = 100;
        if (chkOpacity.value) requestPreviewDebounced();
    };

    chkSkew.onClick = function () {
        sldSkew.enabled = chkSkew.value;
        drawPreview();
    };

    sldSkew.onChanging = function () {
        skewMaxDeg = Math.round(sldSkew.value);
        if (skewMaxDeg < 0) skewMaxDeg = 0;
        if (skewMaxDeg > 45) skewMaxDeg = 45;
        if (chkSkew.value) requestPreviewDebounced();
    };

    chkRandSize.onClick = function () {
        sldRandSize.enabled = chkRandSize.value;
        drawPreview();
    };

    sldRandSize.onChanging = function () {
        randSizeStrength = Math.round(sldRandSize.value);
        if (chkRandSize.value) requestPreviewDebounced();
    };

    chkMargin.onClick = function () {
        sldMargin.enabled = chkMargin.value;

        // ONのときだけ値を少し右（+30）に調整
        try {
            if (chkMargin.value) {
                var newVal = Number(sldMargin.value) + 30;
                if (isNaN(newVal)) newVal = 20;
                var __mx = 50;
                try { __mx = Number(sldMargin.maxvalue); } catch (_) { __mx = 50; }
                if (!__mx || isNaN(__mx)) __mx = 50;
                if (newVal > __mx) newVal = __mx; // 上限
                sldMargin.value = newVal;
                try { maskMarginPt = Math.round(Number(sldMargin.value)); } catch (_) { }
                requestPreviewDebounced();
            }
        } catch (_) { }

        // マージンON時はマスク処理をOFF（両立させない）
        try {
            if (chkMargin.value && chkMask && chkMask.enabled) {
                chkMask.value = false;
            }
        } catch (_) { }
    };

    sldMargin.onChanging = function () {
        maskMarginPt = Math.round(sldMargin.value);
        requestPreviewDebounced();
    };

    sldHollow.onChanging = function () {
        hollowStrength = Math.round(sldHollow.value * 10) / 10; // 0.1刻み表示
        if (rbHollow.value || rbGradY.value) requestPreviewDebounced();
    };

    dlg.onShow = function () {
        shiftDialogPosition(dlg, offsetX, offsetY);
        try { refreshSymbolDropdown(); } catch (_) { }
        try { applyMarginMaxToUI(); } catch (_) { }
        try {
            if (__initView && __initZoom != null) {
                sldZoom.value = __initView.zoom;
            }
        } catch (_) { }
        drawPreview();
        try { app.redraw(); } catch (_) { }
    };

    // 位置サンプリング（分布）
    function pickPoint(bounds) {
        // bounds: [L, T, R, B]
        var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
        var cx = (L + R) * 0.5;
        var cy = (T + B) * 0.5;

        var pt;

        // 全体に均等
        if (rbEven && rbEven.value) {
            pt = {
                x: random(L, R),
                y: random(B, T)
            };
            return pt;
        }

        // グラデーション（上濃→下薄）
        if (rbGradY && rbGradY.value) {
            var h = (T - B);
            var u = Math.random(); // 0..1
            // top-bias: u を上側に寄せる
            var gy = 1 + ((typeof hollowStrength === "number" ? hollowStrength : 2.0) - 1) * 0.6; // 1..4 程度にマップ
            var t2 = 1 - Math.pow(1 - u, gy); // 0(bottom) .. 1(top)
            pt = {
                x: random(L, R),
                y: B + h * t2
            };
            return pt;
        }

        // 中心を空ける（中心:薄い / 外側:濃い の滑らかなグラデーション）
        // hollowStrength: 1..6（1=ほぼ均等 / 6=外側に強く寄せる）
        var ang = random(0, Math.PI * 2);
        var c = Math.cos(ang);
        var s = Math.sin(ang);

        // その角度で矩形内に収まる最大半径を計算
        var maxRx = (c === 0) ? 1e12 : (c > 0 ? (R - cx) / c : (L - cx) / c);
        var maxRy = (s === 0) ? 1e12 : (s > 0 ? (T - cy) / s : (B - cy) / s);
        var maxR = Math.min(Math.abs(maxRx), Math.abs(maxRy));

        var k = (typeof hollowStrength === "number" && hollowStrength > 0) ? hollowStrength : 2.0;
        if (k < 1) k = 1;
        if (k > 6) k = 6;

        // 0..1 を外側寄りに変換（中心を完全に抜かず、中心ほど出にくい）
        var uu = Math.random(); // 0..1
        var r01 = 1 - Math.pow(1 - uu, k); // bias to 1 (outer)
        var r = maxR * r01;

        pt = {
            x: cx + c * r,
            y: cy + s * r
        };

        return pt;
    }

    // ランダム関数
    function random(min, max) {
        return Math.random() * (max - min) + min;
    }

    // 歪み（シアー）を適用（中心基準） / Apply shear (skew) about center
    function applyShearToItem(item, deg) {
        if (!item) return;
        if (!deg || deg === 0) return;
        var rad = deg * Math.PI / 180;
        var m = new Matrix();
        // X方向シアー: [1  tan; 0  1]
        m.mValueA = 1;
        m.mValueB = Math.tan(rad);
        m.mValueC = 0;
        m.mValueD = 1;
        m.mValueTX = 0;
        m.mValueTY = 0;

        // 位置・塗り・線幅などは維持しつつ、中心基準で変形
        try {
            item.transform(m, true, true, true, true, 1, Transformation.CENTER);
        } catch (_) {
            // フォールバック: shear API（環境によってはこちらが効く）
            try { item.shear(deg); } catch (__) { }
        }
    }


    function getConfettiSize() {
        var base = (typeof getBaseSizePt === "function") ? getBaseSizePt() : ((minSize + maxSize) * 0.5);

        try {
            if (!(chkRandSize && chkRandSize.value)) {
                return base;
            }
        } catch (_) { }

        var strength = (typeof randSizeStrength === "number") ? randSizeStrength : 0;
        if (strength < 100) strength = 100;
        if (strength > 300) strength = 300;

        // 100 → ほぼ固定
        // 300 → 最大約±150%揺れ
        var ratio = (strength - 100) / 200; // 0..1

        // 揺れ幅を base 基準で拡張
        var maxRange = base * 1.5 * ratio; // 最大±150%
        var v = base + random(-maxRange, maxRange);

        if (v < 0.1) v = 0.1;
        return v;
    }

    function setClippingFlag(item) {
        if (!item) return false;
        try {
            if (item.typename === "PathItem") {
                item.clipping = true;
                return true;
            }
            if (item.typename === "CompoundPathItem") {
                if (item.pathItems && item.pathItems.length > 0) {
                    item.pathItems[0].clipping = true;
                    return true;
                }
            }
            if (item.typename === "GroupItem") {
                // Pathfinder結果がGroupになる場合があるため、内側の最初のPathをクリッピングに
                try {
                    if (item.compoundPathItems && item.compoundPathItems.length > 0) {
                        var cp = item.compoundPathItems[0];
                        if (cp.pathItems && cp.pathItems.length > 0) {
                            cp.pathItems[0].clipping = true;
                            return true;
                        }
                    }
                } catch (_) { }
                try {
                    if (item.pathItems && item.pathItems.length > 0) {
                        item.pathItems[0].clipping = true;
                        return true;
                    }
                } catch (_) { }
            }
        } catch (e) { }
        return false;
    }

    function buildMaskGroup(layerOrGroup) {
        // 戻り値: { container: (GroupItem or Layer), group: GroupItem|null }
        if (!chkMask || !chkMask.value) {
            return { container: layerOrGroup, group: null };
        }
        var g = layerOrGroup.groupItems.add();
        g.name = "__ConfettiMasked__";
        return { container: g, group: g };
    }

    function createMaskShapeFromSelection(container, sourceItem) {
        // 選択オブジェクトに合わせたマスク図形を作成
        // - PathItem: 選択パスを複製してマスクに
        // - クリップグループ: クリッピングパス（Path/Compound）を複製
        // - それ以外: バウンディング矩形でフォールバック
        if (!container || !sourceItem) return null;

        // Path / CompoundPath はそのまま複製
        try {
            if (sourceItem.typename === "PathItem" || sourceItem.typename === "CompoundPathItem") {
                return sourceItem.duplicate(container, ElementPlacement.PLACEATEND);
            }
        } catch (_) { }

        // テキスト: 複製してマスクとして利用（アウトライン化しない）
        try {
            if (sourceItem.typename === "TextFrame") {
                return sourceItem.duplicate(container, ElementPlacement.PLACEATEND);
            }
        } catch (_) { }

        // クリップグループ: クリッピングパスを探して複製
        try {
            if (sourceItem.typename === "GroupItem") {
                var gi = sourceItem;
                var clipCandidate = null;

                // GroupItem.clipped が true の場合はほぼクリップグループ
                // クリッピング指定の付いた Path / Compound を優先して探す
                try {
                    // CompoundPath の方を先に探す（穴あき形状などを想定）
                    if (gi.compoundPathItems && gi.compoundPathItems.length > 0) {
                        for (var cpi = 0; cpi < gi.compoundPathItems.length; cpi++) {
                            var c = gi.compoundPathItems[cpi];
                            try {
                                if (c.pathItems && c.pathItems.length > 0 && c.pathItems[0].clipping) {
                                    clipCandidate = c;
                                    break;
                                }
                            } catch (_) { }
                        }
                    }

                    if (!clipCandidate && gi.pathItems && gi.pathItems.length > 0) {
                        for (var pi = 0; pi < gi.pathItems.length; pi++) {
                            var p = gi.pathItems[pi];
                            try {
                                if (p.clipping) {
                                    clipCandidate = p;
                                    break;
                                }
                            } catch (_) { }
                        }
                    }

                    // クリッピング指定が見つからない場合: 最初のCompound/Pathを使う（最後の手段）
                    if (!clipCandidate) {
                        if (gi.compoundPathItems && gi.compoundPathItems.length > 0) clipCandidate = gi.compoundPathItems[0];
                        else if (gi.pathItems && gi.pathItems.length > 0) clipCandidate = gi.pathItems[0];
                    }

                    if (clipCandidate) {
                        return clipCandidate.duplicate(container, ElementPlacement.PLACEATEND);
                    }

                    // クリップ候補が取れない（通常グループなど）の場合：合体して単一パス化
                    var united = uniteGroupToSinglePath(container, gi);
                    if (united) {
                        return united;
                    }
                } catch (_) { }
            }
        } catch (_) { }

        // フォールバック: バウンディング矩形
        var b;
        try {
            b = sourceItem.geometricBounds; // [L, T, R, B]
        } catch (e) {
            return null;
        }
        var L = b[0], T = b[1], R = b[2], B = b[3];

        var w = (R - L);
        var h = (T - B);
        if (!w || !h) return null;

        var m = container.pathItems.rectangle(T, L, w, h);
        // クリッピングパスは「存在」している必要があるため、塗りはON（None色）にして不可視化する
        try { m.stroked = false; } catch (_) { }
        try { m.filled = true; m.fillColor = new NoColor(); } catch (_) { }
        return m;
    }

    function uniteGroupToSinglePath(container, groupItem) {
        if (!container || !groupItem) return null;
        var dup = null;
        var oldSel = null;
        try {
            oldSel = doc.selection;
        } catch (_) { oldSel = null; }

        try {
            // 複製して container に入れる
            dup = groupItem.duplicate(container, ElementPlacement.PLACEATEND);
        } catch (e) {
            return null;
        }

        try {
            // 選択を複製グループに切り替え、合体→展開
            doc.selection = null;
            dup.selected = true;
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');

            // expandStyle後は selection が合体結果（Path/Compound/Group）になる想定
            if (doc.selection && doc.selection.length > 0) {
                return doc.selection[0];
            }
        } catch (e2) {
            // 失敗したら複製を残さない
            try { dup.remove(); } catch (_) { }
            return null;
        } finally {
            // 選択復元
            try {
                doc.selection = null;
                if (oldSel && oldSel.length) {
                    for (var i = 0; i < oldSel.length; i++) {
                        try { oldSel[i].selected = true; } catch (_) { }
                    }
                } else {
                    // もともと単一選択前提
                    try { selectedObj.selected = true; } catch (_) { }
                }
            } catch (_) { }
        }

        return null;
    }

    function applyMaskToGroup(group, sourceItem) {
        if (!group || !sourceItem) return;

        // 選択オブジェクトに合わせたマスク図形を作成（複製しない）
        var maskItem = createMaskShapeFromSelection(group, sourceItem);
        if (!maskItem) return;

        // クリッピングパスは最前面に
        try { maskItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }

        // グループの最前面へ（クリッピングパスは最前面が原則）
        try { maskItem.move(group, ElementPlacement.PLACEATEND); } catch (_) { }
        // move後にも最前面を保証
        try { maskItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }

        // TextFrame は .clipping を立てられないため、makeMask を使う
        if (maskItem && maskItem.typename === "TextFrame") {
            try {
                // グループ内の全アイテムを選択してクリッピングマスクを作成
                doc.selection = null;
                try { group.selected = true; } catch (_) { }
                app.executeMenuCommand('makeMask');
            } catch (_) { }
            return;
        }

        // クリッピングフラグ
        setClippingFlag(maskItem);

        try { group.clipped = true; } catch (_) { }
    }

    // 五芒星を生成
    function createStar(layer, left, top, size) {
        // left, top は他形状同様、バウンディングの左上基準
        // size は外接正方形の一辺
        var outerR = size * 0.5;
        var innerR = outerR * 0.5; // 五芒星の内側半径（見た目優先）
        var cx = left + outerR;
        var cy = top - outerR;

        var pts = [];
        // 上向きスタート
        var startDeg = -90;
        for (var i = 0; i < 10; i++) {
            var r = (i % 2 === 0) ? outerR : innerR;
            var ang = (startDeg + i * 36) * Math.PI / 180; // 360/10=36
            var px = cx + Math.cos(ang) * r;
            var py = cy + Math.sin(ang) * r;
            pts.push([px, py]);
        }

        var p = layer.pathItems.add();
        p.setEntirePath(pts);
        p.closed = true;
        return p;
    }

    // スター（4点）を生成
    function createStar4(layer, left, top, size) {
        // left, top は他形状同様、バウンディングの左上基準
        // size は外接正方形の一辺
        var outerR = size * 0.5;
        var innerR = outerR * 0.25; // 4点スターの内側半径（尖り強め）
        var cx = left + outerR;
        var cy = top - outerR;

        var pts = [];
        // 上向きスタート（外側→内側→…）
        var startDeg = -90;
        for (var i = 0; i < 8; i++) {
            var r = (i % 2 === 0) ? outerR : innerR;
            var ang = (startDeg + i * 45) * Math.PI / 180; // 360/8=45
            var px = cx + Math.cos(ang) * r;
            var py = cy + Math.sin(ang) * r;
            pts.push([px, py]);
        }

        var p = layer.pathItems.add();
        p.setEntirePath(pts);
        p.closed = true;
        return p;
    }

    // リボン（1/4円弧×2 を反転してつなげたS字帯）を生成 / Quarter-arc mirrored S ribbon band
    function createRibbon(layer, left, top, size) {
        var cx = left + size * 0.5;
        var cy = top - size * 0.5;

        // 2つの1/4円弧でS字を作る（中心線）
        var r = Math.max(0.8, size * 0.72);              // 半径（基本を約150%）
        var halfW = Math.max(0.25, size * 0.165);         // 帯の半幅（基本を約150%）
        var seg = 10;                                    // 円弧の分割（滑らかさ）

        // ローカル座標（原点付近）で作ってから中心に寄せる
        // 1つ目: center (0,0), angle 180→90 で (-r,0)→(0,r)
        // 2つ目: center (0,2r), angle -90→0 で (0,r)→(r,2r)
        var centerPts = [];

        function ptCircle(cpx, cpy, deg) {
            var a = deg * Math.PI / 180;
            return [cpx + Math.cos(a) * r, cpy + Math.sin(a) * r];
        }

        for (var i = 0; i <= seg; i++) {
            var t = i / seg;
            var deg = 180 - 90 * t;
            centerPts.push(ptCircle(0, 0, deg));
        }
        for (var j = 1; j <= seg; j++) {
            var t2 = j / seg;
            var deg2 = -90 + 90 * t2;
            centerPts.push(ptCircle(0, 2 * r, deg2));
        }

        // 全体の重心を (cx,cy) に合わせる（中心線の中点で合わせる）
        // centerPts のYレンジは 0..2r なので、その中心は r
        var offX = cx - 0;
        var offY = cy - r;

        // タンジェントから法線を作って、上下にオフセットして帯の輪郭を作る
        function norm2(vx, vy) {
            var d = Math.sqrt(vx * vx + vy * vy);
            if (d === 0) d = 1;
            return [vx / d, vy / d];
        }

        var topPts = [];
        var botPts = [];

        for (var k = 0; k < centerPts.length; k++) {
            var p0 = centerPts[(k === 0) ? 0 : (k - 1)];
            var p1 = centerPts[k];
            var p2 = centerPts[(k === centerPts.length - 1) ? (centerPts.length - 1) : (k + 1)];

            var tx = p2[0] - p0[0];
            var ty = p2[1] - p0[1];
            var tn = norm2(tx, ty);

            // 左法線
            var nx = -tn[1];
            var ny = tn[0];

            var x = p1[0] + offX;
            var y = p1[1] + offY;

            topPts.push([x + nx * halfW, y + ny * halfW]);
            botPts.push([x - nx * halfW, y - ny * halfW]);
        }

        var pts = [];
        for (var a = 0; a < topPts.length; a++) pts.push(topPts[a]);
        for (var b = botPts.length - 1; b >= 0; b--) pts.push(botPts[b]);

        var path = layer.pathItems.add();
        path.closed = true;
        path.setEntirePath(pts);

        // 角を立てず滑らかに（S字の帯として自然に見せる）
        try {
            var ppts = path.pathPoints;
            for (var m = 0; m < ppts.length; m++) {
                try { ppts[m].pointType = PointType.SMOOTH; } catch (_) { }
            }
        } catch (_) { }

        try { path.filled = true; } catch (_) { }
        try { path.stroked = false; } catch (_) { }

        return path;
    }

    // コンフェティの形状をランダムに作成
    function createConfetti(layer, x, y, size, color) {
        // シンボル（形状パネルの「シンボル」ON かつドロップダウンで選択されている場合のみ候補にする）
        var __symbolEnabled = false;
        try { __symbolEnabled = (chkSymbolShape && chkSymbolShape.value && selectedSymbolName && selectedSymbolName !== ""); } catch (_) { __symbolEnabled = false; }

        var enabledShapes = [];
        if (chkRect.value) enabledShapes.push(0);
        if (chkCircle.value) enabledShapes.push(1);
        if (chkTriangle.value) enabledShapes.push(2);
        if (chkStar.value) enabledShapes.push(3);
        if (chkStar4.value) enabledShapes.push(4);
        if (chkRibbon.value) enabledShapes.push(5);
        if (__symbolEnabled) enabledShapes.push(6);

        if (enabledShapes.length === 0) return null;

        var shape = enabledShapes[Math.floor(Math.random() * enabledShapes.length)];
        var confetti;

        switch (shape) {
            case 0: // 四角形
                confetti = layer.pathItems.rectangle(y, x, size, size * random(0.3, 0.6));
                break;
            case 1: // 円
                confetti = layer.pathItems.ellipse(y, x, size, size);
                break;
            case 2: // 三角形
                confetti = layer.pathItems.add();
                var h = size * 1.2;
                confetti.setEntirePath([
                    [x, y],
                    [x + size, y],
                    [x + size * 0.5, y - h]
                ]);
                confetti.closed = true;
                break;
            case 3: // スター（五芒星）
                confetti = createStar(layer, x, y, size);
                break;
            case 4: // スター（4点）
                confetti = createStar4(layer, x, y, size);
                break;
            case 5: // リボン
                confetti = createRibbon(layer, x, y, size);
                break;
            case 6: // シンボル
                confetti = null;
                try {
                    // シンボル生成（既存ロジックをそのまま実行）
                    if (selectedSymbolName && selectedSymbolName !== "") {
                        var sym = null;
                        try {
                            if (doc && doc.symbols && doc.symbols.length > 0) {
                                for (var ss = 0; ss < doc.symbols.length; ss++) {
                                    try {
                                        if (String(doc.symbols[ss].name) === String(selectedSymbolName)) {
                                            sym = doc.symbols[ss];
                                            break;
                                        }
                                    } catch (_) { }
                                }
                            }
                        } catch (_) { sym = null; }

                        if (sym) {
                            var si = null;
                            var wrap = null;

                            // wrap（グループ）を先に作る（Layer/Group のどちらでもOK）
                            try {
                                if (layer && layer.groupItems && layer.groupItems.add) {
                                    wrap = layer.groupItems.add();
                                    try { wrap.name = "__ConfettiSymbol__"; } catch (_) { }
                                }
                            } catch (_) { wrap = null; }

                            // SymbolItem は Layer に作るのが確実
                            try {
                                if (doc && doc.activeLayer && doc.activeLayer.symbolItems && doc.activeLayer.symbolItems.add) {
                                    si = doc.activeLayer.symbolItems.add(sym);
                                }
                            } catch (_) { si = null; }

                            if (si) {
                                // wrap が作れなかった場合はフォールバック（単体シンボル）
                                if (!wrap) {
                                    // size に合わせて等比スケール（最大辺= size）
                                    try {
                                        var w0a = Number(si.width);
                                        var h0a = Number(si.height);
                                        var m0a = Math.max(w0a, h0a);
                                        if (m0a && m0a > 0) {
                                            var pctA = (Number(size) / m0a) * 100;
                                            if (pctA < 1) pctA = 1;
                                            si.resize(pctA, pctA);
                                        }
                                    } catch (_) { }

                                    // ランダムに回転
                                    try { si.rotate(random(0, 360)); } catch (_) { }

                                    // 歪み / Skew
                                    try {
                                        if (chkSkew && chkSkew.value) {
                                            var _degSa = (typeof skewMaxDeg === "number") ? skewMaxDeg : 0;
                                            if (_degSa < 0) _degSa = 0;
                                            if (_degSa > 45) _degSa = 45;
                                            if (_degSa > 0) {
                                                applyShearToItem(si, random(-_degSa, _degSa));
                                            }
                                        }
                                    } catch (_) { }

                                    // 位置補正（bounds で左上を合わせる）
                                    try {
                                        var gba = si.geometricBounds; // [L, T, R, B]
                                        var dxa = Number(x) - Number(gba[0]);
                                        var dya = Number(y) - Number(gba[1]);
                                        if (!isNaN(dxa) && !isNaN(dya)) {
                                            try { si.left = Number(si.left) + dxa; } catch (_) { }
                                            try { si.top = Number(si.top) + dya; } catch (_) { }
                                        }
                                    } catch (_) {
                                        try { si.left = x; } catch (__) { }
                                        try { si.top = y; } catch (__) { }
                                    }

                                    // 不透明度 / Opacity
                                    try {
                                        if (chkOpacity && chkOpacity.value) {
                                            var _minOpA = (typeof opacityMin === "number") ? opacityMin : 70;
                                            if (_minOpA < 0) _minOpA = 0;
                                            if (_minOpA > 100) _minOpA = 100;
                                            si.opacity = random(_minOpA, 100);
                                        } else {
                                            si.opacity = 100;
                                        }
                                    } catch (_) { }

                                    confetti = si;
                                } else {
                                    // wrap に移動
                                    try { si.move(wrap, ElementPlacement.PLACEATEND); } catch (_) { }

                                    // size に合わせて等比スケール（最大辺= size）: wrap 全体に適用
                                    try {
                                        var wW = Number(wrap.width);
                                        var hW = Number(wrap.height);
                                        var mW = Math.max(wW, hW);
                                        if (mW && mW > 0) {
                                            var pct = (Number(size) / mW) * 100;
                                            if (pct < 1) pct = 1;
                                            wrap.resize(pct, pct);
                                        }
                                    } catch (_) { }

                                    // ランダムに回転
                                    try { wrap.rotate(random(0, 360)); } catch (_) { }

                                    // 歪み / Skew
                                    try {
                                        if (chkSkew && chkSkew.value) {
                                            var _degS = (typeof skewMaxDeg === "number") ? skewMaxDeg : 0;
                                            if (_degS < 0) _degS = 0;
                                            if (_degS > 45) _degS = 45;
                                            if (_degS > 0) {
                                                applyShearToItem(wrap, random(-_degS, _degS));
                                            }
                                        }
                                    } catch (_) { }

                                    // 位置（bounds で左上を合わせる）
                                    try {
                                        var gb = wrap.geometricBounds; // [L, T, R, B]
                                        var dx = Number(x) - Number(gb[0]);
                                        var dy = Number(y) - Number(gb[1]);
                                        if (!isNaN(dx) && !isNaN(dy)) {
                                            try { wrap.left = Number(wrap.left) + dx; } catch (_) { }
                                            try { wrap.top = Number(wrap.top) + dy; } catch (_) { }
                                        }
                                    } catch (_) {
                                        try { wrap.left = x; } catch (__) { }
                                        try { wrap.top = y; } catch (__) { }
                                    }

                                    // 不透明度 / Opacity
                                    try {
                                        if (chkOpacity && chkOpacity.value) {
                                            var _minOp = (typeof opacityMin === "number") ? opacityMin : 70;
                                            if (_minOp < 0) _minOp = 0;
                                            if (_minOp > 100) _minOp = 100;
                                            wrap.opacity = random(_minOp, 100);
                                        } else {
                                            wrap.opacity = 100;
                                        }
                                    } catch (_) { }

                                    confetti = wrap;
                                }
                            }
                        }
                    }
                } catch (_) { confetti = null; }
                break;
        }

        // 色を設定（リボンはストローク、他は塗り）
        if (shape !== 6) {
            var rgbColor = new RGBColor();
            rgbColor.red = color[0];
            rgbColor.green = color[1];
            rgbColor.blue = color[2];

            try {
                // open-path はストローク、closed-path は塗り（リボンは閉じパスに変更）
                var isOpen = false;
                try { isOpen = (confetti && confetti.closed === false); } catch (_) { isOpen = false; }

                if (isOpen) {
                    try { confetti.filled = false; } catch (_) { }
                    try { confetti.stroked = true; } catch (_) { }
                    try { confetti.strokeColor = rgbColor; } catch (_) { }
                } else {
                    try { confetti.filled = true; } catch (_) { }
                    try { confetti.stroked = false; } catch (_) { }
                    try { confetti.fillColor = rgbColor; } catch (_) { }
                }
            } catch (_) { }

            // ランダムに回転
            confetti.rotate(random(0, 360));

            // 歪み / Skew
            try {
                if (chkSkew && chkSkew.value) {
                    var _deg = (typeof skewMaxDeg === "number") ? skewMaxDeg : 0;
                    if (_deg < 0) _deg = 0;
                    if (_deg > 45) _deg = 45;
                    if (_deg > 0) {
                        applyShearToItem(confetti, random(-_deg, _deg));
                    }
                }
            } catch (_) { }

            // 不透明度 / Opacity
            try {
                if (chkOpacity && chkOpacity.value) {
                    var _minOp = (typeof opacityMin === "number") ? opacityMin : 70;
                    if (_minOp < 0) _minOp = 0;
                    if (_minOp > 100) _minOp = 100;
                    // 不透明度をランダムに設定（_minOp..100）
                    confetti.opacity = random(_minOp, 100);
                } else {
                    // OFF: 設定しない（=100%）
                    confetti.opacity = 100;
                }
            } catch (_) { }
        }

        return confetti;
    }

    var result = dlg.show();

    try {
        if (__debounceTaskId) { app.cancelTask(__debounceTaskId); __debounceTaskId = 0; }
    } catch (_) { }

    if (result !== 1) {
        clearPreview();
        try { previewLayer.remove(); } catch (e) { }
        try {
            if (__initView && __initZoom != null) {
                __initView.zoom = __initZoom;
                if (__initCenter) __initView.centerPoint = __initCenter;
            }
        } catch (_) { }
        return;
    }

    // OK: プレビューの見た目そのままで確定（再生成しない）
    // まず、プレビューが空なら念のため一度描画
    try {
        if (previewLayer && previewLayer.pageItems && previewLayer.pageItems.length === 0) {
            drawPreview();
        }
    } catch (_) { }

    // 出力先レイヤー（既存があれば再利用）
    var confettiLayer = null;
    try {
        for (var li = 0; li < doc.layers.length; li++) {
            if (doc.layers[li].name === "Confetti") {
                confettiLayer = doc.layers[li];
                break;
            }
        }
    } catch (_) { confettiLayer = null; }

    if (!confettiLayer) {
        confettiLayer = doc.layers.add();
        confettiLayer.name = "Confetti";
    }

    // プレビュー内容を丸ごと移動（マスクグループ含む）
    // マスクOFF時は、この実行で生成したものだけをグループ化
    var __confettiNewGroup = null;
    try {
        var __doGroup = true;
        try { __doGroup = !(chkMask && chkMask.value); } catch (_) { __doGroup = true; }
        if (__doGroup) {
            try {
                __confettiNewGroup = confettiLayer.groupItems.add();
                try { __confettiNewGroup.name = "__ConfettiGroup__"; } catch (_) { }
            } catch (_) { __confettiNewGroup = null; }
        }

        if (previewLayer) {
            while (previewLayer.pageItems.length > 0) {
                try {
                    if (__confettiNewGroup) {
                        previewLayer.pageItems[0].move(__confettiNewGroup, ElementPlacement.PLACEATEND);
                    } else {
                        previewLayer.pageItems[0].move(confettiLayer, ElementPlacement.PLACEATEND);
                    }
                } catch (eMove1) {
                    break;
                }
            }
        }
    } catch (_) { }

    // プレビューレイヤーを削除
    try { previewLayer.remove(); } catch (e2) { }

    // 実行後、生成したコンフェティ全体を選択状態にする
    try {
        doc.selection = null;
        if (__confettiNewGroup) {
            try { __confettiNewGroup.selected = true; } catch (_) { }
        } else {
            for (var si = 0; si < confettiLayer.pageItems.length; si++) {
                try { confettiLayer.pageItems[si].selected = true; } catch (_) { }
            }
        }
    } catch (_) { }

    // 画面ズームはプレビュー操作を尊重（復元しない）

})();