#target illustrator
#targetengine "TitleBarLineCutEngine"
try { app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); } catch (_) { }

// =========================
// セッション内の値保持 / Session-only persistence
// Illustrator再起動で消える（targetengineの寿命に依存）
// =========================
(function initSessionState() {
    if (!$.global.__tblc_state) {
        $.global.__tblc_state = {
            marginText: null,
            roundOn: false,
            roundText: null,
            fillOn: true,
            notchOn: false,
            strokeOn: true,
            widthText: null,
            capIndex: 0 // 0:none 1:round
        };
    }
})();

// =========================
// バージョン / Version
// =========================
var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "タイトル帯の罫線欠け処理",
        en: "Title Bar Line Cut"
    },

    // Alerts
    alertOpenDoc: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    alertSelectTwo: {
        ja: "テキスト1つと長方形パス1つを選択してください。",
        en: "Select one text object and one rectangle path."
    },
    alertSelectTypes: {
        ja: "テキスト（1つ）と、4点の長方形パス（1つ）を選択してください。",
        en: "Select 1 text object and 1 rectangle path (4 points)."
    },
    alertMarginNonNegative: {
        ja: "マージン(%u)は0以上の数値で指定してください。",
        en: "Margin (%u) must be 0 or greater."
    },

    // Buttons
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },

    // Panels
    panelFill: { ja: "塗り", en: "Fill" },
    panelStroke: { ja: "線", en: "Stroke" },

    // Labels / Controls
    margin: { ja: "マージン", en: "Margin" },
    roundCorners: { ja: "角丸", en: "Round" },
    fillOn: { ja: "塗り", en: "Fill" },
    notch: { ja: "ノッチ", en: "Notch" },
    strokeWidth: { ja: "線幅", en: "Stroke" },
    lineCap: { ja: "線端", en: "Cap" },
    capNone: { ja: "なし", en: "None" },
    capRound: { ja: "丸型", en: "Round" }
};

function L(key) {
    try {
        var v = LABELS[key];
        if (!v) return key;
        return v[lang] || v.en || v.ja || key;
    } catch (e) {
        return key;
    }
}

function LF(key, unitLabel) {
    // %u を単位ラベルで置換
    var s = L(key);
    return s.replace(/%u/g, String(unitLabel));
}


(function () {
    if (app.documents.length === 0) { alert(L('alertOpenDoc')); return; }
    var doc = app.activeDocument;

    if (!doc.selection || doc.selection.length !== 2) {
        alert(L('alertSelectTwo'));
        return;
    }

    function isTextItem(it) { return it && it.typename === "TextFrame"; }
    function isRectPathItem(it) {
        return it && it.typename === "PathItem" && it.closed === true && it.pathPoints.length === 4;
    }

    var a = doc.selection[0], b = doc.selection[1];
    var tf = null, rectA = null;
    if (isTextItem(a) && isRectPathItem(b)) { tf = a; rectA = b; }
    else if (isTextItem(b) && isRectPathItem(a)) { tf = b; rectA = a; }
    else { alert(L('alertSelectTypes')); return; }

    // rulerType を参照して単位ラベルと pt 変換を決める
    var unitLabelMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    function getCurrentRulerType() {
        try {
            return app.preferences.getIntegerPreference("rulerType");
        } catch (e) {
            return 2; // pt fallback
        }
    }

    function getUnitInfo() {
        var rt = getCurrentRulerType();
        var label = unitLabelMap.hasOwnProperty(rt) ? unitLabelMap[rt] : "pt";

        // value * factor = points
        var factor;
        switch (rt) {
            case 0: factor = 72; break;                     // in
            case 1: factor = 72 / 25.4; break;              // mm
            case 2: factor = 1; break;                      // pt
            case 3: factor = 12; break;                     // pica
            case 4: factor = 72 / 2.54; break;              // cm
            case 5: factor = (72 / 25.4) * 0.25; break;     // Q (0.25mm) / H (0.25mm)
            case 6: factor = 1; break;                      // px (Illustratorは基本72ppi扱い)
            case 7: factor = 72 * 12; break;                // ft/in（ここではftを想定）
            case 8: factor = (72 / 25.4) * 1000; break;     // m
            case 9: factor = 72 * 36; break;                // yd
            case 10: factor = 72 * 12; break;                // ft
            default: factor = 1; label = "pt"; break;
        }

        return { rulerType: rt, label: label, factorToPt: factor };
    }

    var __unitInfo = getUnitInfo();

    function unitToPt(v) {
        return v * __unitInfo.factorToPt;
    }

    // strokeUnits を参照して線幅の単位ラベルと pt 変換を決める
    function getCurrentStrokeUnits() {
        try {
            return app.preferences.getIntegerPreference("strokeUnits");
        } catch (e) {
            return 2; // pt fallback
        }
    }

    function getStrokeUnitInfo() {
        var su = getCurrentStrokeUnits();
        var label = unitLabelMap.hasOwnProperty(su) ? unitLabelMap[su] : "pt";

        // value * factor = points
        var factor;
        switch (su) {
            case 0: factor = 72; break;                     // in
            case 1: factor = 72 / 25.4; break;              // mm
            case 2: factor = 1; break;                      // pt
            case 3: factor = 12; break;                     // pica
            case 4: factor = 72 / 2.54; break;              // cm
            case 5: factor = (72 / 25.4) * 0.25; break;     // Q/H (0.25mm)
            case 6: factor = 1; break;                      // px
            case 7: factor = 72 * 12; break;                // ft/in（ここではftを想定）
            case 8: factor = (72 / 25.4) * 1000; break;     // m
            case 9: factor = 72 * 36; break;                // yd
            case 10: factor = 72 * 12; break;                // ft
            default: factor = 1; label = "pt"; break;
        }

        return { strokeUnits: su, label: label, factorToPt: factor };
    }

    var __strokeUnitInfo = getStrokeUnitInfo();

    function strokeUnitToPt(v) {
        return v * __strokeUnitInfo.factorToPt;
    }

    function ptToStrokeUnit(vPt) {
        return vPt / __strokeUnitInfo.factorToPt;
    }

    // ---- 状態保持（プレビュー用）----
    var rectLayer = rectA.layer;
    var rectZ = rectA.zOrderPosition; // 同レイヤー内の相対順（AIにより不安定な場合あり）

    var orig = {
        id: rectA.uuid || null,
        // boundsは再生成に使う
        rBounds: rectA.visibleBounds.slice(0), // [L,T,R,B]
        // appearance (A基準)
        stroked: rectA.stroked,
        filled: rectA.filled,
        fillColor: null,
        strokeColor: null,
        strokeWidth: rectA.strokeWidth,
        strokeDashes: rectA.strokeDashes,
        dashOffset: rectA.dashOffset,
        strokeCap: rectA.strokeCap,
        strokeJoin: rectA.strokeJoin,
        miterLimit: rectA.miterLimit,
        strokeOverprint: rectA.strokeOverprint,
        opacity: rectA.opacity,
        blendingMode: rectA.blendingMode
    };
    try { orig.strokeColor = rectA.strokeColor; } catch (e) { }
    try { orig.fillColor = rectA.fillColor; } catch (e) { }

    var previewItem = null; // 生成したプレビュー用の開いたパス

    // A（元）/ B（塗りのみ）/ C（線のみ）
    var rectB = null;


    var rectC = null;

    function makeFillOnlyFromA(src) {
        // Aの塗りが「なし」の場合はBを作らない
        try { if (!src.filled) return null; } catch (e) { }

        var b = src.duplicate();
        try { b.hidden = false; } catch (e) { }
        b.stroked = false;
        b.filled = true;
        return b;
    }

    function makeStrokeOnlyFromA(src) {
        var c = src.duplicate();
        try { c.hidden = false; } catch (e) { }
        c.filled = false;
        c.stroked = true;
        return c;
    }

    function rebuildFillOnlyB() {
        // 角丸のON/OFFや半径変更に追従させるため、Bは作り直す（effectの二重掛け防止）
        try { if (rectB) rectB.remove(); } catch (e) { }
        rectB = null;

        // Aの塗りが「なし」の場合はBを生成しない
        try {
            if (!rectA.filled) return;
        } catch (e) { }

        rectB = makeFillOnlyFromA(rectA);
        if (!rectB) return;
        try { rectB.hidden = false; } catch (e) { }
        // 可能なら C の直前に配置（A/B/C の重なりを崩しにくい）
        try {
            if (rectC) rectB.move(rectC, ElementPlacement.PLACEBEFORE);
        } catch (e) { }

        // 角丸値（pt）を先に取得
        var rrB = 0;
        try {
            if (typeof cbRound !== "undefined" && cbRound.value) {
                rrB = getRoundRadiusPt();
            }
        } catch (e) { rrB = 0; }

        // ノッチONならBに欠き取り（ライブ：パスファインダー減算）
        var didNotch = false;
        try {
            if (typeof cbNotch !== "undefined" && cbNotch.value) {
                // UIができている（et/parseMarginが使える）ときだけ
                if (typeof et !== "undefined" && typeof parseMargin === "function") {
                    var m = parseMargin();
                    if (m !== null) {
                        var mp = unitToPt(m);
                        var newB = applyNotchSubtractToFill(rectB, mp);
                        if (newB) {
                            rectB = newB; // GroupItemになる場合がある
                            didNotch = true;
                        }
                    }
                }
            }
        } catch (e) { }

        // 角丸：ノッチON時はグループ（Z）に対して適用、OFF時は従来通りBに適用
        try {
            if (rrB > 0) {
                applyRoundCornersEffect(rectB, rrB);
            }
        } catch (e) { }

        // Bは最背面へ（塗りを常に背面に）
        try {
            if (rectB) rectB.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (e) { }
    }

    rectC = makeStrokeOnlyFromA(rectA);
    rectB = makeFillOnlyFromA(rectA);
    // Aの塗りが「なし」の場合は rectB は null のまま

    // 線はデフォルトで有効
    var __strokeEnabled = true;

    // ダイアログ中はAを隠す（最終的にOK時にAは削除する）
    var __rectAHiddenForDialog = false;
    try { rectA.hidden = true; __rectAHiddenForDialog = true; } catch (e) { }

    function copyAppearance(srcObj, dst) {
        dst.stroked = srcObj.stroked;
        dst.strokeWidth = srcObj.strokeWidth;
        dst.strokeDashes = srcObj.strokeDashes;
        dst.dashOffset = srcObj.dashOffset;
        dst.strokeCap = srcObj.strokeCap;
        dst.strokeJoin = srcObj.strokeJoin;
        dst.miterLimit = srcObj.miterLimit;
        dst.strokeOverprint = srcObj.strokeOverprint;
        try { dst.strokeColor = srcObj.strokeColor; } catch (e) { }
        try { dst.opacity = srcObj.opacity; } catch (e) { }
        try { dst.blendingMode = srcObj.blendingMode; } catch (e) { }
    }

    // テキスト外接（計算用）
    // テキストは見た目上の外接とズレることがあるため、
    // 一時的に複製→アウトライン化して、その外接で計算する（1回だけ）
    var __textBoundsForCalc = null; // [L,T,R,B]

    function buildTextBoundsForCalcOnce() {
        if (__textBoundsForCalc) return;
        try {
            // 複製 → アウトライン化（元のtfは触らない）
            var dup = tf.duplicate();
            // createOutline() は複製テキストをアウトラインに変換し、戻り値はGroupItemになる
            var outlined = dup.createOutline();
            try { outlined.hidden = true; } catch (e) { }
            __textBoundsForCalc = outlined.visibleBounds.slice(0);
            try { outlined.remove(); } catch (e) { }
        } catch (e) {
            // フォールバック：通常のvisibleBounds
            try { __textBoundsForCalc = tf.visibleBounds.slice(0); } catch (_) { __textBoundsForCalc = null; }
        }
    }

    function getDefaultMarginValueInt() {
        // テキスト高の半分を、現在単位に変換して整数に丸める
        try {
            buildTextBoundsForCalcOnce();
            var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
            var hPt = Math.abs(tb[1] - tb[3]);
            var halfPt = hPt / 2;
            var v = halfPt / __unitInfo.factorToPt; // pt -> current unit
            if (isNaN(v) || !isFinite(v)) return 2;
            return Math.round(v);
        } catch (e) {
            return 2;
        }
    }

    var __defaultMarginInt = getDefaultMarginValueInt();

    function getDefaultRoundValueInt() {
        // テキスト高の1/4を、現在単位に変換して整数に丸める
        try {
            buildTextBoundsForCalcOnce();
            var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
            var hPt = Math.abs(tb[1] - tb[3]);
            var qPt = hPt / 4;
            var v = qPt / __unitInfo.factorToPt; // pt -> current unit
            if (isNaN(v) || !isFinite(v)) return 2;
            return Math.round(v);
        } catch (e) {
            return 2;
        }
    }

    var __defaultRoundInt = getDefaultRoundValueInt();

    function getDefaultStrokeWidthValue() {
        // 元の線幅（pt）を strokeUnits に変換して表示用にする
        try {
            var v = ptToStrokeUnit(orig.strokeWidth);
            if (isNaN(v) || !isFinite(v) || v <= 0) v = 1;
            // 見やすさ優先：小数1桁に丸め（整数なら整数表示になる）
            v = Math.round(v * 10) / 10;
            return String(v);
        } catch (e) {
            return "1";
        }
    }

    var __defaultStrokeWidthText = getDefaultStrokeWidthValue();

    function getCutRange(addPt) {
        buildTextBoundsForCalcOnce();
        var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
        return { cutL: tb[0] - addPt, cutR: tb[2] + addPt };
    }

    function clampCutToRect(cut, rL, rR) {
        var cutL = cut.cutL, cutR = cut.cutR;
        if (cutL < rL) cutL = rL;
        if (cutR > rR) cutR = rR;
        return { cutL: cutL, cutR: cutR };
    }

    function makeRectByBounds(bounds, layer) {
        // bounds: [L,T,R,B]
        var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
        var w = R - L;
        var h = T - B;
        var p = doc.pathItems.rectangle(T, L, w, h);
        try { p.stroked = false; } catch (e) { }
        try { p.filled = false; } catch (e) { }
        if (layer) {
            try { p.move(layer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
        }
        return p;
    }

    function applyNotchSubtractToFill(fillItem, marginPt) {
        // Returns the new container (GroupItem) that has live pathfinder subtract applied.
        // If it fails, returns the original fillItem.
        if (!fillItem || !cbNotch || !cbNotch.value) return fillItem;
        if (marginPt === undefined || marginPt === null) return fillItem;

        try {
            buildTextBoundsForCalcOnce();
            var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]

            // X: text bounds rectangle (calc-only)
            var rectX = makeRectByBounds(tb, fillItem.layer);

            // Y: expanded by margin on all sides
            var yb = [tb[0] - marginPt, tb[1] + marginPt, tb[2] + marginPt, tb[3] - marginPt];
            var rectY = makeRectByBounds(yb, fillItem.layer);

            // Yは減算用カッターなので filled=true にして面として成立させる
            try { rectY.stroked = false; } catch (e) { }
            try { rectY.filled = true; } catch (e) { }
            try { rectY.hidden = false; } catch (e) { }
            try { rectY.fillColor = fillItem.fillColor; } catch (e) { }
            try { rectY.opacity = 100; } catch (e) { }

            // Z: group (fill + Y)
            var g = doc.groupItems.add();
            try { g.move(fillItem.layer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }

            // order: fill behind, cutter(Y) in front
            try { fillItem.move(g, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
            try { rectY.move(g, ElementPlacement.PLACEATEND); } catch (e) { }

            // X is calc-only
            try { rectX.remove(); } catch (e) { }

            // Apply Live Pathfinder Subtract to Z
            try {
                doc.selection = null;
                g.selected = true;
                app.executeMenuCommand('Live Pathfinder Minus Back');
            } catch (e) { }

            try { doc.selection = null; } catch (e) { }

            return g;
        } catch (e) {
            return fillItem;
        }
    }

    // プレビュー適用：元rectは残したまま、上に“欠けた罫線”を重ねて見せる
    //（元の上辺を消さずに重ねる方式なので、プレビューは「完成形と同じ」にするため
    //  いったん元rectを非表示にして、プレビューで置き換えて表示します）
    var origVisible = rectA.hidden;

    function clearPreview() {
        if (previewItem) {
            try { previewItem.remove(); } catch (e) { }
            previewItem = null;
        }
        try { if (rectC) rectC.hidden = false; } catch (e) { }
    }

    function applyPreview(marginMM) {
        clearPreview();

        // 線が無効ならCを作らない（=欠き取りもしない）
        try {
            if (!cbStrokeOn.value || !rectC) return;
        } catch (e) { return; }

        var addPt = unitToPt(marginMM);
        var rB = orig.rBounds;
        var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];

        var cut = getCutRange(addPt);

        // 交差しないなら何もしない（元を表示）
        if (cut.cutR <= rL || cut.cutL >= rR) {
            try { if (rectC) rectC.hidden = false; } catch (e) { }
            return;
        }

        var cc = clampCutToRect(cut, rL, rR);

        // 元rectを隠して、開いたパス（欠け罫線）で見せる
        rectC.hidden = true;

        previewItem = doc.pathItems.add();
        previewItem.setEntirePath([
            [cc.cutR, rT],
            [rR, rT],
            [rR, rBot],
            [rL, rBot],
            [rL, rT],
            [cc.cutL, rT]
        ]);
        previewItem.closed = false;
        previewItem.filled = false;
        previewItem.stroked = true;

        // C（線のみ）の見た目を基準にする
        try { copyAppearance(rectC, previewItem); } catch (e) { copyAppearance(orig, previewItem); }

        // 線幅（入力が有効なら上書き）
        try {
            var sw = parseStrokeWidth();
            if (sw !== null) previewItem.strokeWidth = strokeUnitToPt(sw);
        } catch (e) { }

        var selCap = getSelectedCapOrNull();
        if (selCap !== null) {
            try { previewItem.strokeCap = selCap; } catch (e) { }
        }

        // 角丸
        try {
            var rr = getRoundRadiusPt();
            if (rr > 0) applyRoundCornersEffect(previewItem, rr);
        } catch (e) { }

        // レイヤーを合わせる
        try { previewItem.move((rectC ? rectC.layer : rectLayer), ElementPlacement.PLACEATBEGINNING); } catch (e) { }
        // 可能なら元rectの直後/直前へ近づけたいが、zOrderPositionの厳密復元は難しいため最前面寄せ
        try { previewItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e) { }
    }

    // ---- ダイアログ / Dialog ----
    var dlg = new Window("dialog", L('dialogTitle') + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    // ダイアログの位置と透明度 / Dialog position & opacity
    var offsetX = 300;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        try { dlg.opacity = opacityValue; } catch (e) { }
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, 0);

    var row = dlg.add("group");
    row.alignChildren = ["left", "center"];

    var stMarginLabel = row.add("statictext", undefined, L('margin'));
    stMarginLabel.preferredSize.width = 60;
    stMarginLabel.justify = "right";

    var et = row.add("edittext", undefined, String(__defaultMarginInt));
    et.characters = 3;
    try {
        if ($.global.__tblc_state.marginText !== null && $.global.__tblc_state.marginText !== undefined) {
            et.text = String($.global.__tblc_state.marginText);
        }
    } catch (e) { }

    var stUnitMargin = row.add("statictext", undefined, __unitInfo.label);

    // 角丸（マージンの直下）
    var rowR = dlg.add("group");
    rowR.alignChildren = ["left", "center"];

    var cbRound = rowR.add("checkbox", undefined, L('roundCorners'));
    cbRound.value = false;
    try { cbRound.value = !!$.global.__tblc_state.roundOn; } catch (e) { }

    var etRound = rowR.add("edittext", undefined, String(__defaultRoundInt));
    try {
        if ($.global.__tblc_state.roundText !== null && $.global.__tblc_state.roundText !== undefined) {
            etRound.text = String($.global.__tblc_state.roundText);
        }
    } catch (e) { }
    etRound.characters = 3;

    var stUnitRound = rowR.add("statictext", undefined, __unitInfo.label);
    etRound.enabled = false;

    // 塗り
    var fillPanel = dlg.add("panel", undefined, L('panelFill'));
    fillPanel.orientation = "column";
    fillPanel.alignChildren = ["fill", "top"];
    fillPanel.margins = [15, 20, 15, 10];

    var rowF = fillPanel.add("group");
    rowF.alignChildren = ["left", "center"];

    var cbFillOn = rowF.add("checkbox", undefined, L('fillOn'));
    cbFillOn.value = true;
    try { cbFillOn.value = !!$.global.__tblc_state.fillOn; } catch (e) { }

    var rowN = fillPanel.add("group");
    rowN.alignChildren = ["left", "center"];

    var cbNotch = rowN.add("checkbox", undefined, L('notch'));
    cbNotch.value = false;
    try { cbNotch.value = !!$.global.__tblc_state.notchOn; } catch (e) { }

    // 線（線幅・線端）
    var linePanel = dlg.add("panel", undefined, L('panelStroke'));
    linePanel.orientation = "column";
    linePanel.alignChildren = ["fill", "top"];
    linePanel.margins = [15, 20, 15, 10];

    // 線幅（strokeUnits）
    var rowW = linePanel.add("group");
    rowW.alignChildren = ["left", "center"];

    var cbStrokeOn = rowW.add("checkbox", undefined, "");
    cbStrokeOn.value = true;
    try { cbStrokeOn.value = !!$.global.__tblc_state.strokeOn; } catch (e) { }

    var stWidthLabel = rowW.add("statictext", undefined, L('strokeWidth'));

    var etWidth = rowW.add("edittext", undefined, __defaultStrokeWidthText);
    etWidth.characters = 3;
    try {
        if ($.global.__tblc_state.widthText !== null && $.global.__tblc_state.widthText !== undefined) {
            etWidth.text = String($.global.__tblc_state.widthText);
        }
    } catch (e) { }

    var stUnitWidth = rowW.add("statictext", undefined, __strokeUnitInfo.label);

    // 線端
    var capRow = linePanel.add("group");
    capRow.orientation = "row";
    capRow.alignChildren = ["left", "center"];

    var stCapLabel = capRow.add("statictext", undefined, L('lineCap'));

    var capBtns = capRow.add("group");
    capBtns.alignment = ["left", "center"];
    capBtns.orientation = "row";
    capBtns.alignChildren = ["left", "center"];

    var rbCapNone = capBtns.add("radiobutton", undefined, L('capNone'));
    var rbCapRound = capBtns.add("radiobutton", undefined, L('capRound'));

    // セッション復元（線端）
    var __capIdx = 0;
    try { __capIdx = ($.global.__tblc_state.capIndex === 1) ? 1 : 0; } catch (e) { __capIdx = 0; }
    rbCapNone.value = (__capIdx === 0);
    rbCapRound.value = (__capIdx === 1);


    function getSelectedCapOrNull() {
        // null = 変更しない（元の線端を維持）
        if (rbCapRound.value) return StrokeCap.ROUNDENDCAP;
        return null;
    }

    function parseRoundRadius() {
        if (!cbRound.value) return 0;
        var v = parseFloat(etRound.text);
        if (isNaN(v) || v <= 0) return 0;
        return v;
    }

    function getRoundRadiusPt() {
        return unitToPt(parseRoundRadius());
    }

    function createRoundCornersEffectXML(radiusPt) {
        // Reference format: <Dict data="R radius #value# "/>
        // radiusPt is in points
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>';
        return xml;
    }

    function applyRoundCornersEffect(item, radiusPt) {
        if (!item || radiusPt <= 0) return;
        // Adobe Round Corners live effect (radius in points)
        try {
            item.applyEffect(createRoundCornersEffectXML(radiusPt));
        } catch (e) { }
    }

    function syncRoundUI() {
        etRound.enabled = cbRound.value;
    }

    function syncFillUI() {
        try { cbNotch.enabled = !!cbFillOn.value; } catch (e) { }
    }

    function syncStrokeUIAndObjects() {
        var on = true;
        try { on = !!cbStrokeOn.value; } catch (e) { on = true; }

        // UI enable/disable
        try { etWidth.enabled = on; } catch (e) { }
        try { capRow.enabled = on; } catch (e) { }

        // Object create/remove
        if (!on) {
            try { clearPreview(); } catch (e) { }
            try { if (rectC) rectC.remove(); } catch (e) { }
            rectC = null;
            return;
        }

        // ON: rectCが無ければ作る
        if (!rectC) {
            try {
                rectC = makeStrokeOnlyFromA(rectA);
                // 可能ならBの後ろへ
                try { if (rectB) rectC.move(rectB, ElementPlacement.PLACEAFTER); } catch (_) { }
            } catch (e) {
                rectC = null;
            }
        }
    }

    function persistDialogValues() {
        try {
            $.global.__tblc_state.marginText = et.text;
            $.global.__tblc_state.roundOn = !!cbRound.value;
            $.global.__tblc_state.roundText = etRound.text;
            $.global.__tblc_state.fillOn = !!cbFillOn.value;
            $.global.__tblc_state.notchOn = !!cbNotch.value;
            $.global.__tblc_state.strokeOn = !!cbStrokeOn.value;
            $.global.__tblc_state.widthText = etWidth.text;
            $.global.__tblc_state.capIndex = rbCapRound.value ? 1 : 0;
        } catch (e) { }
    }

    var btns = dlg.add("group");
    btns.alignment = "right";
    var btCancel = btns.add("button", undefined, L('cancel'), { name: "cancel" });
    var btOK = btns.add("button", undefined, L('ok'), { name: "ok" });

    function parseMargin() {
        var v = parseFloat(et.text);
        if (isNaN(v) || v < 0) return null;
        return v;
    }

    function parseStrokeWidth() {
        var v = parseFloat(etWidth.text);
        if (isNaN(v) || v < 0.1) return null; // 0は不可、最低0.1
        return v;
    }

    // ↑↓キーで数値を増減（Shift: ±10 / Option: ±0.1）
    // 使い方: changeValueByArrowKey(et);
    function changeValueByArrowKey(editText, allowNegative, minValue) {
        if (allowNegative === undefined) allowNegative = false;
        if (minValue === undefined) minValue = 0;

        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;
            var handled = false;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    handled = true;
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    handled = true;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName === "Up") {
                    value += delta;
                    handled = true;
                } else if (event.keyName === "Down") {
                    value -= delta;
                    handled = true;
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    handled = true;
                } else if (event.keyName === "Down") {
                    value -= delta;
                    handled = true;
                }
            }

            if (!handled) return;

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;
            if (minValue > 0 && value < minValue) value = minValue;

            event.preventDefault();
            editText.text = value;

            // 既存のプレビュー更新フローに乗せる
            try {
                if (typeof editText.onChanging === "function") editText.onChanging();
            } catch (e) { }
        });
    }

    function updatePreview() {
        // B（塗りのみ）を角丸設定に追従（塗りOFFなら削除）
        try {
            if (typeof cbFillOn !== "undefined" && !cbFillOn.value) {
                if (rectB) { try { rectB.remove(); } catch (e) { } }
                rectB = null;
            } else {
                rebuildFillOnlyB();
            }
        } catch (e) { }

        var m = parseMargin();
        if (m === null) return; // 不正値は触らない

        // 線が有効なときだけ欠き取りプレビュー
        try {
            if (cbStrokeOn.value && rectC) applyPreview(m);
            else clearPreview();
        } catch (e) {
            clearPreview();
        }
        app.redraw();
    }

    et.onChanging = updatePreview;
    etRound.onChanging = updatePreview;
    etWidth.onChanging = updatePreview;

    cbStrokeOn.onClick = function () {
        syncStrokeUIAndObjects();
        updatePreview();
    };

    cbRound.onClick = function () {
        syncRoundUI();
        // 角丸ONのときは線端を「丸型」に寄せる
        if (cbRound.value) {
            try { rbCapRound.value = true; } catch (e) { }
        }
        updatePreview();
    };
    cbNotch.onClick = updatePreview;
    cbFillOn.onClick = function () {
        syncFillUI();
        updatePreview();
    };
    rbCapNone.onClick = updatePreview;
    rbCapRound.onClick = updatePreview;

    changeValueByArrowKey(et); // マージンは負値NG
    changeValueByArrowKey(etRound); // 角丸半径も負値NG
    changeValueByArrowKey(etWidth, false, 0.1); // 線幅は最低0.1

    // 初期状態
    syncRoundUI();
    syncFillUI();
    syncStrokeUIAndObjects();

    // テキスト外接（計算用）を先に確定（1回だけ）
    buildTextBoundsForCalcOnce();

    // 初期のBを角丸設定に合わせて作り直し
    rebuildFillOnlyB();

    // 初期プレビュー
    updatePreview();

    var result = dlg.show();
    persistDialogValues();

    if (result !== 1) {
        // キャンセル：プレビュー解除し、B/Cを破棄してAを戻す
        clearPreview();
        try { if (rectB) rectB.remove(); } catch (e) { }
        try { if (rectC) rectC.remove(); } catch (e) { }
        try { if (__rectAHiddenForDialog) rectA.hidden = false; } catch (e) { }
        app.redraw();
        return;
    }

    // OK：プレビューを確定（元rectを置き換え）
    var marginMM = parseMargin();
    if (marginMM === null) {
        clearPreview();
        alert(LF('alertMarginNonNegative', __unitInfo.label));
        return;
    }

    // 線が無効なら：Cも欠き取りも無し。Bのみ残してAを削除（Bが無い場合はAを戻して終了）
    if (!cbStrokeOn.value) {
        try { if (rectB) rectB.hidden = false; } catch (e) { }
        if (!rectB) {
            try { if (__rectAHiddenForDialog) rectA.hidden = false; } catch (e) { }
            return;
        }
        try { rectA.remove(); } catch (e) { }
        return;
    }

    // プレビューがあるならそれを採用、なければ生成
    if (!previewItem) {
        applyPreview(marginMM);
    }

    // 交差しない等でpreviewItemが作られてない場合は何もしない
    if (!previewItem) {
        // 交差しない等で欠き取りが発生しない場合：Aを残すのではなく、B/Cを破棄してAを戻す
        try { if (rectB) rectB.remove(); } catch (e) { }
        try { if (rectC) rectC.remove(); } catch (e) { }
        try { if (__rectAHiddenForDialog) rectA.hidden = false; } catch (e) { }
        return;
    }

    // Cを削除し、previewItemを本番にする（Aは削除、Bは残す）
    try { if (rectC) rectC.remove(); } catch (e) { }
    try { rectA.remove(); } catch (e) { }

    // 念のため表示
    try { previewItem.hidden = false; } catch (e) { }

    // 最終反映（線端）
    try {
        var finalCap = getSelectedCapOrNull();
        if (finalCap !== null) previewItem.strokeCap = finalCap;
    } catch (e) { }

    // 最終反映（線幅）
    try {
        var finalSW = parseStrokeWidth();
        if (finalSW !== null) previewItem.strokeWidth = strokeUnitToPt(finalSW);
    } catch (e) { }

    // 最終反映（角丸）
    try {
        var finalRR = getRoundRadiusPt();
        if (finalRR > 0) applyRoundCornersEffect(previewItem, finalRR);
    } catch (e) { }


})();