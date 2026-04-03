#target illustrator
#targetengine "TitleBarLineCutEngine"
try { app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); } catch (_) { }

var __TBLC_NOTE_KEY = "__TBLC__";

function __tblcMakeNote(bounds) {
  try { return JSON.stringify({k:__TBLC_NOTE_KEY, ver:SCRIPT_VERSION, bounds:bounds}); }
  catch(e){ return ""; }
}
function __tblcParseNote(noteText) {
  if (!noteText) return null;
  try {
    var o = JSON.parse(noteText);
    if (o && o.k===__TBLC_NOTE_KEY && o.bounds && o.bounds.length===4) return o.bounds;
  } catch(e){}
  return null;
}
function __tblcGetBoundsFromItem(it) {
  if (!it) return null;
  try { var b = __tblcParseNote(it.note); if (b) return b; } catch(e){}
  try {
    if (it.typename==="GroupItem") {
      for (var i=0;i<it.pageItems.length;i++){
        try { var bb=__tblcParseNote(it.pageItems[i].note); if (bb) return bb; } catch(_){}
      }
    }
  } catch(e2){}
  return null;
}
function __tblcSetBoundsToItem(it, bounds) {
  if (!it || !bounds || bounds.length!==4) return;
  var note = __tblcMakeNote(bounds);
  if (!note) return;
  try { it.note = note; } catch(e){}
  try {
    if (it.typename==="GroupItem") {
      for (var i=0;i<it.pageItems.length;i++){
        try { it.pageItems[i].note = note; } catch(_){}
      }
    }
  } catch(e2){}
}
function __tblcBoundsEqual(a,b,eps){
  if(!a||!b||a.length!==4||b.length!==4) return false;
  if(eps===undefined) eps=0.01;
  for(var i=0;i<4;i++) if(Math.abs(a[i]-b[i])>eps) return false;
  return true;
}
function __tblcFindFirstStrokeSample(it){
  if(!it) return null;
  try { if(it.typename==="PathItem" && it.stroked) return it; } catch(e){}
  try {
    if(it.typename==="GroupItem"){
      for(var i=0;i<it.pathItems.length;i++){
        try{ if(it.pathItems[i].stroked) return it.pathItems[i]; }catch(_){}
      }
    }
  } catch(e2){}
  return null;
}
function __tblcFindFillSampleByBounds(layer,bounds){
  if(!layer||!bounds) return null;
  try{
    for(var i=0;i<layer.pageItems.length;i++){
      var it=layer.pageItems[i];
      var b=__tblcGetBoundsFromItem(it);
      if(!b||!__tblcBoundsEqual(b,bounds,0.01)) continue;

      if(it.typename==="PathItem"){
        if(it.filled && !it.stroked) return it;
      } else if(it.typename==="GroupItem"){
        var hasFill=false, hasStroke=false;
        for(var j=0;j<it.pathItems.length;j++){
          try{
            if(it.pathItems[j].filled) hasFill=true;
            if(it.pathItems[j].stroked) hasStroke=true;
          }catch(_){}
        }
        if(hasFill && !hasStroke) return it;
      }
    }
  }catch(e){}
  return null;
}

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
// 概要 / Summary
// 更新日 / Updated: 2026-02-14
//
// テキスト（タイトル）と長方形パスを選択し、
// テキストが重なる位置の罫線（フレーム線）を欠けさせます。
// テキスト位置が中央でなくても、四辺のうち最も近い辺（最大2辺）に対して欠け処理を行います。
// 角丸（ライブ効果）、塗り（B）/線（C）の分離、ノッチ（塗りの欠き取り）に対応。
// セッション内でダイアログ値を保持（Illustrator再起動で消えます）。
// 生成済みの欠け罫線を選択して再実行しても、内部メモ（note）から長方形を復元して再計算します。
// =========================

var SCRIPT_VERSION = "v1.3";

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
        ja: "テキスト（1つ）と、長方形パス（1つ）または欠け罫線（1つ）を選択してください。",
        en: "Select 1 text object and 1 rectangle path, or 1 previously generated cut frame."
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


// Main IIFE, with updated selection parsing and rerun support
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

    function isReRunnableFrameItem(it) {
        if (!it) return false;
        if (it.typename !== "PathItem" && it.typename !== "GroupItem") return false;
        return !!__tblcGetBoundsFromItem(it);
    }

    var a = doc.selection[0], b = doc.selection[1];
    var tf = null, rectA = null;
    var __rerunSource = null;
    var __rerunBounds = null;

    if (isTextItem(a) && isRectPathItem(b)) { tf = a; rectA = b; }
    else if (isTextItem(b) && isRectPathItem(a)) { tf = b; rectA = a; }
    else if (isTextItem(a) && isReRunnableFrameItem(b)) { tf = a; __rerunSource = b; }
    else if (isTextItem(b) && isReRunnableFrameItem(a)) { tf = b; __rerunSource = a; }
    else { alert(L('alertSelectTypes')); return; }

    // --- Re-run reconstruction ---
    if (!rectA && __rerunSource) {
        __rerunBounds = __tblcGetBoundsFromItem(__rerunSource);
        if (!__rerunBounds) { alert(L('alertSelectTypes')); return; }

        rectA = doc.pathItems.rectangle(
            __rerunBounds[1],
            __rerunBounds[0],
            (__rerunBounds[2] - __rerunBounds[0]),
            (__rerunBounds[1] - __rerunBounds[3])
        );

        var strokeSample = __tblcFindFirstStrokeSample(__rerunSource);
        if (strokeSample) {
            rectA.stroked = true;
            rectA.filled = false;
            copyAppearance(strokeSample, rectA);
        } else {
            rectA.stroked = true;
        }

        var fillSample = __tblcFindFillSampleByBounds(rectA.layer, __rerunBounds);
        if (fillSample) {
            try {
                rectA.filled = true;
                if (fillSample.typename === "PathItem") {
                    rectA.fillColor = fillSample.fillColor;
                    rectA.opacity = fillSample.opacity;
                    rectA.blendingMode = fillSample.blendingMode;
                } else if (fillSample.typename === "GroupItem" && fillSample.pathItems.length) {
                    rectA.fillColor = fillSample.pathItems[0].fillColor;
                    rectA.opacity = fillSample.opacity;
                    rectA.blendingMode = fillSample.blendingMode;
                }
            } catch (_) { }
        }

        try { rectA.move(__rerunSource, ElementPlacement.PLACEATBEGINNING); } catch (_) { }
    }

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

    var __previewItems = []; // 生成したプレビュー用（PathItem または GroupItem）

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

    // ---- どの辺でも欠けさせるためのヘルパ ----
    // bounds: [L,T,R,B]
    function overlaps1D(a0, a1, b0, b1) {
        var loA = Math.min(a0, a1), hiA = Math.max(a0, a1);
        var loB = Math.min(b0, b1), hiB = Math.max(b0, b1);
        return Math.min(hiA, hiB) - Math.max(loA, loB);
    }

    function pickNearestEdge(tb, rB) {
        // returns: "top" | "bottom" | "left" | "right"
        var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];
        var tL = tb[0], tT = tb[1], tR = tb[2], tB = tb[3];

        // Prefer edges that have actual projection overlap with the text bounds.
        var ovX = overlaps1D(tL, tR, rL, rR);
        var ovY = overlaps1D(tT, tB, rT, rBot);

        var dTop = Math.abs(tT - rT);
        var dBottom = Math.abs(tB - rBot);
        var dLeft = Math.abs(tL - rL);
        var dRight = Math.abs(tR - rR);

        var best = { edge: "top", score: 1e18 };

        function consider(edge, dist, ok) {
            if (!ok) return;
            if (dist < best.score) best = { edge: edge, score: dist };
        }

        consider("top", dTop, ovX > 0);
        consider("bottom", dBottom, ovX > 0);
        consider("left", dLeft, ovY > 0);
        consider("right", dRight, ovY > 0);

        // If overlap checks failed (rare), fall back to smallest distance.
        if (best.score >= 1e17) {
            var m = dTop, e = "top";
            if (dBottom < m) { m = dBottom; e = "bottom"; }
            if (dLeft < m) { m = dLeft; e = "left"; }
            if (dRight < m) { m = dRight; e = "right"; }
            return e;
        }

        return best.edge;
    }

    function getCutRangeForEdge(edge, addPt) {
        buildTextBoundsForCalcOnce();
        var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
        if (edge === "left" || edge === "right") {
            // Y方向（上:大 / 下:小）
            return { cutTop: tb[1] + addPt, cutBot: tb[3] - addPt };
        }
        // top/bottom: X方向
        return { cutL: tb[0] - addPt, cutR: tb[2] + addPt };
    }

    function clampCutToRectForEdge(edge, cut, rB) {
        var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];
        if (edge === "left" || edge === "right") {
            var cutTop = cut.cutTop, cutBot = cut.cutBot;
            if (cutTop > rT) cutTop = rT;
            if (cutBot < rBot) cutBot = rBot;
            return { cutTop: cutTop, cutBot: cutBot };
        }
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
        if (__previewItems && __previewItems.length) {
            for (var i = __previewItems.length - 1; i >= 0; i--) {
                try { __previewItems[i].remove(); } catch (e) { }
            }
            __previewItems = [];
        }
        try { if (rectC) rectC.hidden = false; } catch (e) { }
    }
    // ---- 複数エッジ対応ヘルパ ----
    function pickNearestEdges(tb, rB, maxEdges) {
        if (maxEdges === undefined) maxEdges = 2;
        var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];
        var tL = tb[0], tT = tb[1], tR = tb[2], tB = tb[3];

        var ovX = overlaps1D(tL, tR, rL, rR);
        var ovY = overlaps1D(tT, tB, rT, rBot);

        var cands = [];
        function push(edge, dist, ok) {
            if (!ok) return;
            cands.push({ edge: edge, dist: dist });
        }

        push("top", Math.abs(tT - rT), ovX > 0);
        push("bottom", Math.abs(tB - rBot), ovX > 0);
        push("left", Math.abs(tL - rL), ovY > 0);
        push("right", Math.abs(tR - rR), ovY > 0);

        if (!cands.length) {
            // fallback: pick one by distance only
            return [pickNearestEdge(tb, rB)];
        }

        cands.sort(function (a, b) { return a.dist - b.dist; });
        var out = [];
        for (var i = 0; i < cands.length && out.length < maxEdges; i++) out.push(cands[i].edge);
        return out;
    }

    function makeGapInterval(edge, cc, rB) {
        // returns { s0, s1 } in clockwise perimeter parameter, or null
        var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];
        var topLen = rR - rL;
        var rightLen = rT - rBot;
        var bottomLen = topLen;
        var leftLen = rightLen;
        var P = 2 * (topLen + rightLen);

        function clamp01(v) { return (v < 0) ? 0 : ((v > 1) ? 1 : v); }

        var sStart = 0, sEnd = 0;
        if (edge === "top") {
            sStart = cc.cutL - rL;
            sEnd = cc.cutR - rL;
        } else if (edge === "right") {
            // y: rT -> rBot
            sStart = topLen + (rT - cc.cutTop);
            sEnd = topLen + (rT - cc.cutBot);
        } else if (edge === "bottom") {
            // direction: rR -> rL
            var s0b = topLen + rightLen;
            sStart = s0b + (rR - cc.cutR);
            sEnd = s0b + (rR - cc.cutL);
        } else if (edge === "left") {
            // direction: rBot -> rT
            var s0l = topLen + rightLen + bottomLen;
            sStart = s0l + (cc.cutBot - rBot);
            sEnd = s0l + (cc.cutTop - rBot);
        } else {
            return null;
        }

        // normalize
        if (sEnd < sStart) {
            var tmp = sStart; sStart = sEnd; sEnd = tmp;
        }
        if (sEnd <= sStart) return null;
        if (sStart < 0) sStart = 0;
        if (sEnd > P) sEnd = P;
        if (sEnd <= sStart) return null;
        return { s0: sStart, s1: sEnd, P: P, topLen: topLen, rightLen: rightLen };
    }

    function pointAtS(s, rB, topLen, rightLen) {
        var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];
        var bottomLen = topLen;
        var leftLen = rightLen;
        var P = 2 * (topLen + rightLen);
        // wrap to [0,P]
        if (s < 0) s = 0;
        if (s > P) s = P;

        if (s <= topLen) {
            return [rL + s, rT];
        }
        s -= topLen;
        if (s <= rightLen) {
            return [rR, rT - s];
        }
        s -= rightLen;
        if (s <= bottomLen) {
            return [rR - s, rBot];
        }
        s -= bottomLen;
        // left
        if (s <= leftLen) {
            return [rL, rBot + s];
        }
        return [rL, rT];
    }

    function buildPathPointsForInterval(a, b, rB, topLen, rightLen) {
        // interval [a,b] on perimeter clockwise
        var pts = [];
        var P = 2 * (topLen + rightLen);
        function addPt(p) {
            if (!pts.length) { pts.push(p); return; }
            var q = pts[pts.length - 1];
            if (Math.abs(q[0] - p[0]) < 1e-6 && Math.abs(q[1] - p[1]) < 1e-6) return;
            pts.push(p);
        }

        var bounds = [0, topLen, topLen + rightLen, topLen + rightLen + topLen, P];
        addPt(pointAtS(a, rB, topLen, rightLen));
        for (var i = 0; i < bounds.length; i++) {
            var t = bounds[i];
            if (t > a && t < b) addPt(pointAtS(t, rB, topLen, rightLen));
        }
        addPt(pointAtS(b, rB, topLen, rightLen));
        return pts;
    }

    function subtractIntervals(base, rems) {
        // base: [0,P], rems: [{s0,s1},...] sorted by s0
        var out = [];
        var cur = base[0];
        for (var i = 0; i < rems.length; i++) {
            var r = rems[i];
            if (r.s0 > cur) out.push([cur, r.s0]);
            if (r.s1 > cur) cur = r.s1;
        }
        if (cur < base[1]) out.push([cur, base[1]]);
        // filter tiny
        var res = [];
        for (var j = 0; j < out.length; j++) {
            if (out[j][1] - out[j][0] > 1e-4) res.push(out[j]);
        }
        return res;
    }

    function createPreviewStrokeWithGaps(rB, gaps, appearanceSrc, swPt, capObj, rrPt, targetLayer) {
        // gaps: array of {edge, cc}
        var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];
        var topLen = rR - rL;
        var rightLen = rT - rBot;
        if (topLen <= 0 || rightLen <= 0) return [];

        var remIntervals = [];
        for (var i = 0; i < gaps.length; i++) {
            var gi = makeGapInterval(gaps[i].edge, gaps[i].cc, rB);
            if (gi) remIntervals.push({ s0: gi.s0, s1: gi.s1 });
        }
        if (!remIntervals.length) return [];
        remIntervals.sort(function (a, b) { return a.s0 - b.s0; });

        var P = 2 * (topLen + rightLen);
        var keep = subtractIntervals([0, P], remIntervals);
        if (!keep.length) return [];

        var created = [];

        // If more than 1 segment, group them so we can move/zorder once
        var container = null;
        if (keep.length > 1) {
            try {
                container = doc.groupItems.add();
                if (targetLayer) container.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (e) { container = null; }
        }

        for (var k = 0; k < keep.length; k++) {
            var a = keep[k][0], b = keep[k][1];
            var pts = buildPathPointsForInterval(a, b, rB, topLen, rightLen);
            if (!pts || pts.length < 2) continue;

            var p = doc.pathItems.add();
            p.setEntirePath(pts);
            p.closed = false;
            p.filled = false;
            p.stroked = true;

            // appearance
            try { copyAppearance(appearanceSrc, p); } catch (e) { }
            if (swPt !== null && swPt !== undefined) {
                try { p.strokeWidth = swPt; } catch (e) { }
            }
            if (capObj !== null) {
                try { p.strokeCap = capObj; } catch (e) { }
            }
            if (rrPt && rrPt > 0) {
                try { applyRoundCornersEffect(p, rrPt); } catch (e) { }
            }

            if (container) {
                try { p.move(container, ElementPlacement.PLACEATEND); } catch (e) { }
            } else if (targetLayer) {
                try { p.move(targetLayer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
            }

            created.push(p);
        }

        if (container && created.length) {
            // return the group as the single preview item
            return [container];
        }
        return created;
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

        // テキスト外接から「最も近い辺」を最大2つ決め、その辺の区間を欠けさせる
        buildTextBoundsForCalcOnce();
        var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]

        var edges = pickNearestEdges(tb, rB, 2);
        var gaps = [];

        for (var ei = 0; ei < edges.length; ei++) {
            var e = edges[ei];
            var cut = getCutRangeForEdge(e, addPt);
            var cc = clampCutToRectForEdge(e, cut, rB);

            // 成立チェック
            if (e === "top" || e === "bottom") {
                if (cc.cutR <= rL || cc.cutL >= rR || cc.cutR <= cc.cutL) continue;
            } else {
                if (cc.cutTop <= rBot || cc.cutBot >= rT || cc.cutTop <= cc.cutBot) continue;
            }

            gaps.push({ edge: e, cc: cc });
        }

        if (!gaps.length) {
            try { if (rectC) rectC.hidden = false; } catch (e) { }
            return;
        }

        // 元rectを隠して、開いたパス（欠け罫線）で見せる
        rectC.hidden = true;

        // appearance source
        var appearanceSrc = rectC;

        // 線幅（入力が有効なら上書き）
        var swPt = null;
        try {
            var sw = parseStrokeWidth();
            if (sw !== null) swPt = strokeUnitToPt(sw);
        } catch (e) { swPt = null; }

        // 線端
        var selCap = getSelectedCapOrNull();

        // 角丸
        var rr = 0;
        try {
            rr = getRoundRadiusPt();
        } catch (e) { rr = 0; }

        // プレビュー生成（1〜複数パス。複数ならグループで返る）
        __previewItems = createPreviewStrokeWithGaps(rB, gaps, appearanceSrc, swPt, selCap, rr, (rectC ? rectC.layer : rectLayer));

        if (!__previewItems || !__previewItems.length) {
            try { if (rectC) rectC.hidden = false; } catch (e) { }
            return;
        }

        // 前面へ
        try {
            for (var pi = 0; pi < __previewItems.length; pi++) {
                try { __previewItems[pi].zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }
            }
        } catch (e) { }
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
        // re-run temporary rect cleanup
        try {
            if (typeof __rerunSource !== "undefined" && __rerunSource && rectA) {
                rectA.remove();
            }
        } catch (_) { }
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
        // Tag fill-only output for re-run (even when stroke is OFF)
        try {
            __tblcSetBoundsToItem(rectB, orig.rBounds);
        } catch (_) { }
        return;
    }

    // プレビューがあるならそれを採用、なければ生成
    if (!__previewItems || !__previewItems.length) {
        applyPreview(marginMM);
    }

    // 交差しない等でプレビューが作られてない場合は何もしない
    if (!__previewItems || !__previewItems.length) {
        // 交差しない等で欠き取りが発生しない場合：Aを残すのではなく、B/Cを破棄してAを戻す
        try { if (rectB) rectB.remove(); } catch (e) { }
        try { if (rectC) rectC.remove(); } catch (e) { }
        try { if (__rectAHiddenForDialog) rectA.hidden = false; } catch (e) { }
        return;
    }

    // Cを削除し、previewItem(s)を本番にする（Aは削除、Bは残す）
    try { if (rectC) rectC.remove(); } catch (e) { }
    try { rectA.remove(); } catch (e) { }
    // Tag outputs with original rectangle bounds for re-run
    try {
        var tagBounds = orig.rBounds;
        if (rectB) __tblcSetBoundsToItem(rectB, tagBounds);
        if (__previewItems && __previewItems.length) {
            for (var ti = 0; ti < __previewItems.length; ti++) {
                __tblcSetBoundsToItem(__previewItems[ti], tagBounds);
            }
        }
    } catch (_) { }

    // Re-run: remove the previously generated cut frame that was selected
    try {
        if (typeof __rerunSource !== "undefined" && __rerunSource) {
            // Avoid removing newly created items by only removing the originally selected object
            __rerunSource.remove();
            __rerunSource = null;
        }
    } catch (_) { }

    // 念のため表示
    try {
        for (var pi2 = 0; pi2 < __previewItems.length; pi2++) {
            try { __previewItems[pi2].hidden = false; } catch (_) { }
            // if group, ensure children visible
            try {
                if (__previewItems[pi2].typename === "GroupItem") {
                    for (var ci = 0; ci < __previewItems[pi2].pageItems.length; ci++) {
                        try { __previewItems[pi2].pageItems[ci].hidden = false; } catch (_) { }
                    }
                }
            } catch (_) { }
        }
    } catch (e) { }

    // 最終反映（線端）
    try {
        var finalCap = getSelectedCapOrNull();
        if (finalCap !== null) {
            for (var pi3 = 0; pi3 < __previewItems.length; pi3++) {
                var it = __previewItems[pi3];
                if (it.typename === "GroupItem") {
                    for (var ci2 = 0; ci2 < it.pathItems.length; ci2++) {
                        try { it.pathItems[ci2].strokeCap = finalCap; } catch (_) { }
                    }
                } else if (it.typename === "PathItem") {
                    try { it.strokeCap = finalCap; } catch (_) { }
                }
            }
        }
    } catch (e) { }

    // 最終反映（線幅）
    try {
        var finalSW = parseStrokeWidth();
        if (finalSW !== null) {
            var swPt2 = strokeUnitToPt(finalSW);
            for (var pi4 = 0; pi4 < __previewItems.length; pi4++) {
                var it2 = __previewItems[pi4];
                if (it2.typename === "GroupItem") {
                    for (var ci3 = 0; ci3 < it2.pathItems.length; ci3++) {
                        try { it2.pathItems[ci3].strokeWidth = swPt2; } catch (_) { }
                    }
                } else if (it2.typename === "PathItem") {
                    try { it2.strokeWidth = swPt2; } catch (_) { }
                }
            }
        }
    } catch (e) { }

    // 最終反映（角丸）
    try {
        var finalRR = getRoundRadiusPt();
        if (finalRR > 0) {
            for (var pi5 = 0; pi5 < __previewItems.length; pi5++) {
                var it3 = __previewItems[pi5];
                if (it3.typename === "GroupItem") {
                    for (var ci4 = 0; ci4 < it3.pathItems.length; ci4++) {
                        try { applyRoundCornersEffect(it3.pathItems[ci4], finalRR); } catch (_) { }
                    }
                } else if (it3.typename === "PathItem") {
                    try { applyRoundCornersEffect(it3, finalRR); } catch (_) { }
                }
            }
        }
    } catch (e) { }


})();