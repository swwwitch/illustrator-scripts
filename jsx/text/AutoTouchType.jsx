#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* =========================================
 * AutoTouchType.jsx
 * 概要: 選択したテキストの各文字に対して、ベースライン/比率/回転/カーニング/トラッキングをランダムに付与する「オート文字タッチ」ツール。
 *       seed付きRNGでプレビューの見た目を安定化する。
 * 作成日: 2026-02-16
 * 更新日: 2026-02-16
 * バージョン: v1.1
 * ========================================= */

(function () {
    var SCRIPT_VERSION = "v1.1";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "オート文字タッチツール", en: "Auto Touch Type Tool" },
        alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        alertSelectText: { ja: "テキストを選択してください。", en: "Please select text." },
        alertSelectTextRange: { ja: "TextFrame または TextRange を選択してください。", en: "Please select a TextFrame or TextRange." },
        alertEnterNumber: { ja: "数値を入力してください。", en: "Please enter a number." },
        alertNoJPFonts: { ja: "対象の和文フォント（Pr6 / Pr6N）が見つかりません。", en: "No target JP fonts (Pr6 / Pr6N) were found." },
        labelBaseline: { ja: "ベースライン:", en: "Baseline:" },
        labelScale: { ja: "水平/垂直比率:", en: "Scale:" },
        labelRotation: { ja: "文字回転:", en: "Rotation:" },
        labelKerning: { ja: "カーニング:", en: "Kerning:" },
        panelTouch: { ja: "文字タッチ", en: "Touch" },
        panelFont: { ja: "フォント", en: "Font" },
        chkFontRandom: { ja: "ランダム", en: "Random" },
        chkFontJPOnly: { ja: "和文フォントに限る", en: "Japanese only" },
        chkFontRansom: { ja: "「犯行声明文」風", en: "Ransom-note style" },
        panelRansom: { ja: "「犯行声明文」風", en: "Ransom-note style" },
        chkRansomEnable: { ja: "有効", en: "Enable" },
        chkRansomTrack: { ja: "トラッキング調整", en: "Adjust tracking" },
        btnRerun: { ja: "再実行", en: "Rerun" },
        btnReset: { ja: "リセット", en: "Reset" },
        btnAllOn: { ja: "すべてON", en: "All ON" },
        btnAllOff: { ja: "すべてOFF", en: "All OFF" },
        btnCancel: { ja: "キャンセル", en: "Cancel" },
        btnOK: { ja: "OK", en: "OK" }
    };

    function L(key) {
        try {
            var o = LABELS[key];
            if (!o) return key;
            return o[lang] || o.ja || o.en || key;
        } catch (_) {
            return key;
        }
    }


    if (app.documents.length === 0) { alert(L("alertNoDoc")); return; }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) { alert(L("alertSelectText")); return; }

    // --- manifesto style backgrounds (per-character rects via outline) ---
    var BG_LAYER_NAME = '__AutoTouchType_BG__';
    var BG_GROUP_NAME = '__BGRects__';

    function getSelectionTextFrames(selection) {
        var frames = [];
        if (!selection || selection.length === 0) return frames;

        function pushUnique(tf) {
            if (!tf || tf.typename !== 'TextFrame') return;
            for (var k = 0; k < frames.length; k++) {
                if (frames[k] === tf) return;
            }
            frames.push(tf);
        }

        for (var i = 0; i < selection.length; i++) {
            try {
                var it = selection[i];
                if (!it) continue;

                if (it.typename === 'TextFrame') {
                    pushUnique(it);
                } else if (it.typename === 'TextRange') {
                    // selected text inside a text frame
                    try {
                        var p = it.parent;
                        if (p && p.typename === 'TextFrame') pushUnique(p);
                        else if (p && p.parent && p.parent.typename === 'TextFrame') pushUnique(p.parent);
                    } catch (_) { }
                }
            } catch (_) { }
        }
        return frames;
    }

    function ensureBgLayer() {
        var lyr = null;
        try {
            for (var i = 0; i < doc.layers.length; i++) {
                try {
                    if (doc.layers[i].name === BG_LAYER_NAME) { lyr = doc.layers[i]; break; }
                } catch (_) { }
            }
        } catch (_) { }
        if (!lyr) {
            try { lyr = doc.layers.add(); lyr.name = BG_LAYER_NAME; } catch (_) { }
        }
        return lyr;
    }

    function removeBgGroupIfExists(bgLayer) {
        if (!bgLayer) return;
        try {
            for (var i = bgLayer.groupItems.length - 1; i >= 0; i--) {
                if (bgLayer.groupItems[i].name === BG_GROUP_NAME) bgLayer.groupItems[i].remove();
            }
        } catch (_) { }
    }

    function randInt(min, max) { return Math.floor(Math.random() * (max - min + 1)) + min; }

    function makeRandomFill() {
        var g = new GrayColor();
        g.gray = randInt(10, 50);
        return g;
    }

    function setRectStyle(rect) {
        try { rect.stroked = false; rect.filled = true; rect.fillColor = makeRandomFill(); } catch (_) { }
    }

    function jitterRectOutward(rect, maxPt) {
        if (!rect || rect.typename !== 'PathItem') return;
        if (!rect.pathPoints || rect.pathPoints.length < 4) return;
        var b;
        try { b = rect.geometricBounds; } catch (_) { return; }
        var cx = (b[0] + b[2]) / 2;
        var cy = (b[1] + b[3]) / 2;
        for (var i = 0; i < rect.pathPoints.length; i++) {
            try {
                var pp = rect.pathPoints[i];
                var ax = pp.anchor[0], ay = pp.anchor[1];
                var dx = ax - cx;
                var dy = ay - cy;
                var len = Math.sqrt(dx * dx + dy * dy);
                if (!len) continue;
                var d = Math.random() * maxPt;
                var ox = (dx / len) * d;
                var oy = (dy / len) * d;
                pp.anchor = [ax + ox, ay + oy];
                pp.leftDirection = [pp.leftDirection[0] + ox, pp.leftDirection[1] + oy];
                pp.rightDirection = [pp.rightDirection[0] + ox, pp.rightDirection[1] + oy];
            } catch (_) { }
        }
    }

    function addRectFromBounds(bgGroup, b) {
        var PAD = 1;
        var L = b[0] - PAD;
        var T = b[1] + PAD;
        var R = b[2] + PAD;
        var B = b[3] - PAD;
        var w = R - L;
        var h = T - B;
        if (w <= 0 || h <= 0) return null;
        var rect = bgGroup.pathItems.rectangle(T, L, w, h);
        setRectStyle(rect);
        jitterRectOutward(rect, 1);
        return rect;
    }

    function collectCharGroupsFromOutlined(outlined, out) {
        if (!outlined) return;

        function pushLeafUnits(container) {
            try {
                if (!container || !container.pageItems) return;
                for (var k = 0; k < container.pageItems.length; k++) {
                    var it = container.pageItems[k];
                    if (!it) continue;
                    if (it.typename === 'GroupItem') pushLeafUnits(it);
                    else if (it.typename === 'PathItem' || it.typename === 'CompoundPathItem') out.push(it);
                }
            } catch (_) { }
        }

        try {
            if (outlined.typename === 'GroupItem') {
                var hasPushed = false;
                if (outlined.groupItems && outlined.groupItems.length > 0) {
                    for (var i = 0; i < outlined.groupItems.length; i++) {
                        var g1 = outlined.groupItems[i];
                        if (!g1 || g1.typename !== 'GroupItem') continue;
                        if (g1.groupItems && g1.groupItems.length > 0) {
                            for (var j = 0; j < g1.groupItems.length; j++) { out.push(g1.groupItems[j]); hasPushed = true; }
                        } else { out.push(g1); hasPushed = true; }
                    }
                }
                if (hasPushed) return;
                pushLeafUnits(outlined);
                if (out.length > 0) return;
                out.push(outlined);
                return;
            }
            out.push(outlined);
        } catch (_) { }
    }

    function createBackgroundRectsByOutlining(frames) {
        if (!frames || frames.length === 0) return;
        var bgLayer = ensureBgLayer();
        removeBgGroupIfExists(bgLayer);
        var bgGroup = bgLayer.groupItems.add();
        bgGroup.name = BG_GROUP_NAME;

        for (var i = 0; i < frames.length; i++) {
            var tf = frames[i];
            if (!tf || tf.typename !== 'TextFrame') continue;
            var dup = null;
            var outlined = null;
            try {
                dup = tf.duplicate();
                try { dup.move(tf, ElementPlacement.PLACEAFTER); } catch (_) { }
            } catch (_) { continue; }
            try { outlined = dup.createOutline(); } catch (_) { try { if (dup) dup.remove(); } catch (_) { } continue; }
            try { if (dup) dup.remove(); } catch (_) { }

            var charGroups = [];
            collectCharGroupsFromOutlined(outlined, charGroups);
            for (var j = 0; j < charGroups.length; j++) {
                try { addRectFromBounds(bgGroup, charGroups[j].geometricBounds); } catch (_) { }
            }
            try { if (outlined) outlined.remove(); } catch (_) { }
        }

        try { bgGroup.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
        try { bgLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
    }

    function clearBackgroundRectsIfAny() {
        var lyr = null;
        try {
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i] && doc.layers[i].name === BG_LAYER_NAME) { lyr = doc.layers[i]; break; }
            }
        } catch (_) { }
        if (!lyr) return;
        removeBgGroupIfExists(lyr);
    }

    // --- font cache (for Font panel) ---
    var __allFonts = null;
    try { __allFonts = app.textFonts; } catch (_) { __allFonts = null; }

    function getJPFonts() {
        var list = [];
        if (!__allFonts) return list;

        function matchJAKeyword(s) {
            if (!s) return false;
            s = String(s);
            return (
                s.indexOf('Pr6N') !== -1 ||
                s.indexOf('Pr6') !== -1 ||
                s.indexOf('AB-') !== -1 ||
                s.indexOf('FOT') !== -1 ||
                s.indexOf('-OTF') !== -1
            );
        }

        for (var i = 0; i < __allFonts.length; i++) {
            try {
                var f = __allFonts[i];
                if (!f) continue;
                var n1 = f.fullName ? String(f.fullName) : '';
                var n2 = f.name ? String(f.name) : '';
                var n3 = (f.postScriptName) ? String(f.postScriptName) : '';
                if (matchJAKeyword(n1) || matchJAKeyword(n2) || matchJAKeyword(n3)) {
                    list.push(f);
                }
            } catch (_) { }
        }
        return list;
    }

    function collectTextRanges(selection) {
        var list = [];
        for (var i = 0; i < selection.length; i++) {
            var it = selection[i];
            if (!it) continue;
            if (it.typename === "TextRange") list.push(it);
            else if (it.typename === "TextFrame") list.push(it.textRange);
        }
        return list;
    }

    function selectionContainsNonAlnum(ranges) {
        if (!ranges || ranges.length === 0) return false;
        for (var i = 0; i < ranges.length; i++) {
            var tr = ranges[i];
            if (!tr) continue;
            try {
                var txt = tr.contents;
                for (var j = 0; j < txt.length; j++) {
                    var ch = txt.charAt(j);
                    if (!(/[A-Za-z0-9]/.test(ch))) {
                        if (ch === '\r' || ch === '\n') continue;
                        return true;
                    }
                }
            } catch (_) { }
        }
        return false;
    }

    function getKerningSafe(ch) {
        try {
            var k = ch.kerning; // may throw (e.g. Error 9551)
            return (typeof k === "number") ? k : 0;
        } catch (e) {
            return 0;
        }
    }

    var ranges = collectTextRanges(doc.selection);
    var selTextFrames = getSelectionTextFrames(doc.selection);
    if (ranges.length === 0) { alert(L("alertSelectTextRange")); return; }

    // --- snapshot originals (per character) ---
    // we store kerning and tracking as well
    var originals = []; // { ch, baselineShift, hScale, vScale, rotation, kerning, tracking, textFont }
    function snapshotOriginals() {
        originals = [];
        for (var r = 0; r < ranges.length; r++) {
            var tr = ranges[r];
            for (var c = 0; c < tr.length; c++) {
                var ch = tr.characters[c];
                var ca = ch.characterAttributes;
                // Some attributes may not exist in very old versions; assume present
                originals.push({
                    ch: ch,
                    baselineShift: ca.baselineShift,
                    hScale: ca.horizontalScale, // %
                    vScale: ca.verticalScale,   // %
                    rotation: ca.rotation,      // degrees
                    kerning: getKerningSafe(ch),
                    tracking: (typeof ca.tracking === "number") ? ca.tracking : 0,
                    textFont: ca.textFont
                });
            }
        }
    }

    function restoreOriginals() {
        for (var i = 0; i < originals.length; i++) {
            var ca = originals[i].ch.characterAttributes;
            ca.baselineShift = originals[i].baselineShift;
            ca.horizontalScale = originals[i].hScale;
            ca.verticalScale = originals[i].vScale;
            ca.rotation = originals[i].rotation;
            try {
                ca.kerningMethod = AutoKernType.NOAUTOKERN;
                originals[i].ch.kerning = originals[i].kerning;
            } catch (e) {
                // ignore if not supported
            }
            try {
                ca.tracking = originals[i].tracking;
            } catch (e) {
                // not all Illustrator versions expose tracking writable; ignore error
            }
            try {
                if (originals[i].textFont) ca.textFont = originals[i].textFont;
            } catch (e) {
                // ignore if not supported
            }
        }
    }

    // --- seeded RNG for stable preview ---
    function makeRng(seed) {
        var s = seed >>> 0;
        return function () {
            s = (1664525 * s + 1013904223) >>> 0;
            return s / 4294967296;
        };
    }
    function randSym(rng) { return rng() * 2.0 - 1.0; } // -1..1

    // maxAbsPt: baseline shift amplitude in pt
    // maxAbsHPct: scale amplitude in percent (e.g. 10 => 90..110)
    // maxAbsVPct: vertical scale amplitude
    // maxAbsDeg: rotation amplitude in degrees
    // maxAbsKern: kerning amplitude (/1000em units)
    // maxAbsTrk: tracking amplitude (/1000em units)
    function applyRandom(maxAbsPt, maxAbsHPct, maxAbsVPct, linkScale, maxAbsDeg, maxAbsKern, maxAbsTrk, seed, optFont) {
        var rng = makeRng(seed);

        optFont = optFont || {};
        var _doFontRandom = !!optFont.fontRandom;
        var _jpOnly = !!optFont.jpOnly;
        var _ransomTrack = (optFont.ransomTrack !== false); // デフォルトtrue扱い
        var _ransom = !!optFont.ransom;

        var _previewRansomTracking = !!optFont.previewRansomTracking;

        var _fontList = null;
        if (_doFontRandom) {
            _fontList = _jpOnly ? optFont.jpFonts : optFont.allFonts;
            if (!_fontList || _fontList.length === 0) _doFontRandom = false;
        }

        for (var i = 0; i < originals.length; i++) {
            var ca = originals[i].ch.characterAttributes;

            // font
            if (_doFontRandom) {
                try {
                    // skip whitespace / line breaks
                    var _ct = originals[i].ch.contents;
                    if (!(_ct === "\r" || _ct === "\n" || _ct === " ")) {
                        var idx = Math.floor(rng() * _fontList.length);
                        ca.textFont = _fontList[idx];
                    }
                } catch (_) { }
            }

            // baseline
            ca.baselineShift = randSym(rng) * maxAbsPt;

            // rotation (added to original rotation)
            ca.rotation = originals[i].rotation + (randSym(rng) * maxAbsDeg);

            // scale multipliers
            if (linkScale) {
                var m = 1.0 + (randSym(rng) * (maxAbsHPct / 100.0));
                ca.horizontalScale = originals[i].hScale * m;
                ca.verticalScale = originals[i].vScale * m;
            } else {
                var mh = 1.0 + (randSym(rng) * (maxAbsHPct / 100.0));
                var mv = 1.0 + (randSym(rng) * (maxAbsVPct / 100.0));
                ca.horizontalScale = originals[i].hScale * mh;
                ca.verticalScale = originals[i].vScale * mv;
            }

            // kerning
            try {
                ca.kerningMethod = AutoKernType.NOAUTOKERN;
                originals[i].ch.kerning = originals[i].kerning + (randSym(rng) * maxAbsKern);
            } catch (e) {
                // ignore if not supported
            }

            // tracking
            try {
                if ((_ransom && _ransomTrack) || _previewRansomTracking) {
                    ca.tracking = (typeof optFont.ransomTrackValue === "number") ? optFont.ransomTrackValue : 200;
                } else {
                    ca.tracking = originals[i].tracking + (randSym(rng) * maxAbsTrk);
                }
            } catch (e) {
                // ignore if not supported
            }
        }
        app.redraw();
    }

    function parseNum(str) {
        var v = parseFloat(str);
        if (isNaN(v)) return null;
        return v;
    }

    snapshotOriginals();
    var seed = (new Date()).getTime() & 0xffffffff;

    try { clearBackgroundRectsIfAny(); } catch (_) { }

    // --- UI ---
    var w = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    /* ダイアログの位置と透明度 / Dialog position & opacity */
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
        } catch (_) {
            // some environments may not support opacity
        }
    }
    w.orientation = "column";
    w.alignChildren = ["fill", "top"];

    // --- panel: 文字タッチ ---
    var pnlTouch = w.add("panel", undefined, L("panelTouch"));
    pnlTouch.orientation = "column";
    pnlTouch.alignChildren = ["fill", "top"];
    pnlTouch.margins = [15, 20, 15, 10];

    // --- Font area: horizontal layout ---
    var gFontRow = w.add("group");
    gFontRow.orientation = "row";
    gFontRow.alignChildren = ["fill", "top"];

    // --- panel: フォント ---
    var pnlFont = gFontRow.add("panel", undefined, L("panelFont"));
    pnlFont.orientation = "column";
    pnlFont.alignChildren = ["left", "top"];
    pnlFont.margins = [15, 20, 15, 10];
    pnlFont.alignment = ["fill", "top"];

    var chkFontRandom = pnlFont.add("checkbox", undefined, L("chkFontRandom"));
    var chkFontJPOnly = pnlFont.add("checkbox", undefined, L("chkFontJPOnly"));

    // --- panel: 「犯行声明文」風 ---
    var pnlRansom = gFontRow.add("panel", undefined, L("panelRansom"));
    pnlRansom.orientation = "column";
    pnlRansom.alignChildren = ["left", "top"];
    pnlRansom.margins = [15, 20, 15, 10];
    pnlRansom.alignment = ["fill", "top"];

    var chkFontRansom = pnlRansom.add("checkbox", undefined, L("chkRansomEnable"));

    // tracking row: checkbox + value field
    var gRansomTrk = pnlRansom.add("group");
    gRansomTrk.orientation = "row";
    gRansomTrk.alignChildren = ["left", "center"];

    var chkRansomTrack = gRansomTrk.add("checkbox", undefined, L("chkRansomTrack"));
    var edtRansomTrk = gRansomTrk.add("edittext", undefined, "200");
    edtRansomTrk.characters = 5;

    // defaults: OFF
    chkFontRandom.value = false;
    chkFontJPOnly.value = false;
    chkFontRansom.value = false;

    chkRansomTrack.value = true;      // 既存挙動（tracking固定）を維持するためデフォルトON
    chkRansomTrack.enabled = false;   // 「有効」がOFFの間はディム
    edtRansomTrk.enabled = false;     // 「有効」+「トラッキング調整」ONのときだけ有効

    function autoEnableJPOnlyIfNeeded() {
        try {
            if (selectionContainsNonAlnum(ranges)) {
                chkFontJPOnly.value = true;
            }
        } catch (_) { }
    }

    // Auto ON: if selection includes non-alnum (except CR/LF), enable JP-only
    autoEnableJPOnlyIfNeeded();

    var gBase = pnlTouch.add("group");
    var chkBase = gBase.add("checkbox", undefined, "");
    chkBase.value = true;
    chkBase.preferredSize.width = 15;
    var stBase = gBase.add("statictext", undefined, L("labelBaseline"));
    var defaultBase = 0;
    try {
        if (originals.length > 0) {
            var fs = originals[0].ch.characterAttributes.size;
            defaultBase = Math.round(fs / 12);
        }
    } catch (_) { }
    var edtBase = gBase.add("edittext", undefined, String(defaultBase));
    edtBase.characters = 4;
    var stUnitBase = gBase.add("statictext", undefined, "pt");
    var sldBase = gBase.add("slider", undefined, defaultBase, 0, 50);
    sldBase.preferredSize.width = 180;

    var gH = pnlTouch.add("group");
    var chkScale = gH.add("checkbox", undefined, "");
    chkScale.value = true;
    chkScale.preferredSize.width = 15;
    var stH = gH.add("statictext", undefined, L("labelScale"));
    var edtH = gH.add("edittext", undefined, "10");
    edtH.characters = 4;
    var stUnitH = gH.add("statictext", undefined, "%");
    var sldH = gH.add("slider", undefined, 10, 0, 200);
    sldH.preferredSize.width = 180;

    var gRot = pnlTouch.add("group");
    var chkRot = gRot.add("checkbox", undefined, "");
    chkRot.value = true;
    chkRot.preferredSize.width = 15;
    var stRot = gRot.add("statictext", undefined, L("labelRotation"));
    var edtRot = gRot.add("edittext", undefined, "5");
    edtRot.characters = 4;
    var stUnitRot = gRot.add("statictext", undefined, "°");
    var sldRot = gRot.add("slider", undefined, 5, 0, 30);
    sldRot.preferredSize.width = 180;

    var gKern = pnlTouch.add("group");
    var chkKern = gKern.add("checkbox", undefined, "");
    chkKern.value = true;
    chkKern.preferredSize.width = 15;
    var stKern = gKern.add("statictext", undefined, L("labelKerning"));
    var edtKern = gKern.add("edittext", undefined, "-50");
    edtKern.characters = 4;
    var stUnitKern = gKern.add("statictext", undefined, "/1000em");
    var sldKern = gKern.add("slider", undefined, -50, -200, 200);
    sldKern.preferredSize.width = 180;

    // --- buttons inside 文字タッチ panel (bottom) ---
    var gTouchButtons = pnlTouch.add("group");
    gTouchButtons.orientation = "row";
    gTouchButtons.alignChildren = ["right", "center"];
    gTouchButtons.alignment = "left";
    gTouchButtons.margins = [0, 10, 0, 0];

    var btnAllOn = gTouchButtons.add("button", undefined, L("btnAllOn"));
    var btnAllOff = gTouchButtons.add("button", undefined, L("btnAllOff"));
    var btnReset = gTouchButtons.add("button", undefined, L("btnReset"));

    // --- label width align (left) ---
    try {
        var labels = [stBase, stH, stRot, stKern];
        var maxW = 0;
        for (var i = 0; i < labels.length; i++) {
            try {
                var w0 = labels[i].preferredSize.width;
                if (w0 > maxW) maxW = w0;
            } catch (_) { }
        }
        for (var j = 0; j < labels.length; j++) {
            try {
                labels[j].preferredSize.width = maxW;
                // left align
                labels[j].justify = "left";
            } catch (_) { }
        }
    } catch (_) { }

    // --- unit width align (left) ---
    try {
        var units = [stUnitBase, stUnitH, stUnitRot, stUnitKern];
        var maxU = 0;
        for (var ui = 0; ui < units.length; ui++) {
            try {
                var uw = units[ui].preferredSize.width;
                if (uw > maxU) maxU = uw;
            } catch (_) { }
        }
        for (var uj = 0; uj < units.length; uj++) {
            try {
                units[uj].preferredSize.width = maxU;
                units[uj].justify = "left";
            } catch (_) { }
        }
    } catch (_) { }

    // --- slider <-> edit sync ---
    function clamp(v, minV, maxV) {
        if (v < minV) return minV;
        if (v > maxV) return maxV;
        return v;
    }
    function syncFromEdit(edt, sld) {
        var v = parseNum(edt.text);
        if (v === null) return;
        sld.value = clamp(v, sld.minvalue, sld.maxvalue);
    }
    function syncFromSlider(sld, edt) {
        edt.text = String(Math.round(sld.value));
    }

    // --- arrow key increment for edittext ---
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (!(event && (event.keyName === "Up" || event.keyName === "Down"))) return;

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
                // Optionキー押下時は0.1単位で増減
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            } else {
                delta = 1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            }

            if (editText !== edtKern && editText !== edtRansomTrk && value < 0) value = 0;

            event.preventDefault();
            editText.text = String(value);

            // keep slider in sync (if exists)
            try {
                if (editText === edtBase) syncFromEdit(edtBase, sldBase);
                else if (editText === edtH) syncFromEdit(edtH, sldH);
                else if (editText === edtRot) syncFromEdit(edtRot, sldRot);
                else if (editText === edtKern) syncFromEdit(edtKern, sldKern);
            } catch (_) { }

            // preview update
            try { updatePreview(); } catch (_) { }
        });
    }

    // --- bottom buttons (left / spacer / right) ---
    var btnRow = w.add("group");
    btnRow.orientation = "row";
    btnRow.alignChildren = ["fill", "center"];
    btnRow.margins = [0, 15, 0, 0];

    var btnLeft = btnRow.add("group");
    btnLeft.orientation = "row";
    btnLeft.alignChildren = ["left", "center"];

    var btnRerun = btnLeft.add("button", undefined, L("btnRerun"));

    function isTouchAllOff() {
        try {
            return (!chkBase.value && !chkScale.value && !chkRot.value && !chkKern.value);
        } catch (_) {
            return false;
        }
    }

    function updateRerunEnabled() {
        // 文字タッチがすべてOFF かつ フォント「ランダム」もOFF のときは再実行を無効化
        try {
            var touchOff = isTouchAllOff();
            var fontRandOff = !(chkFontRandom && chkFontRandom.value);
            btnRerun.enabled = !(touchOff && fontRandOff);
        } catch (_) { }
    }

    var btnSpacer = btnRow.add("group");
    btnSpacer.orientation = "row";
    btnSpacer.add("statictext", undefined, "");
    btnSpacer.alignment = "fill";

    var btnRight = btnRow.add("group");
    btnRight.orientation = "row";
    btnRight.alignChildren = ["right", "center"];
    var btnCancel = btnRight.add("button", undefined, L("btnCancel"));
    var btnOK = btnRight.add("button", undefined, L("btnOK"));

    function updatePreview() {
        var base = 0;
        if (chkBase.value) { base = parseNum(edtBase.text); if (base === null) return; }

        var hPct = 0;
        if (chkScale.value) { hPct = parseNum(edtH.text); if (hPct === null) return; }

        var rot = 0;
        if (chkRot.value) { rot = parseNum(edtRot.text); if (rot === null) return; }

        var kern = 0;
        if (chkKern.value) { kern = parseNum(edtKern.text); if (kern === null) return; }



        var optFont = {
            fontRandom: (chkFontRandom && chkFontRandom.value),
            jpOnly: (chkFontJPOnly && chkFontJPOnly.value),
            ransom: false, // NOTE: 背景生成などはプレビューしない
            previewRansomTracking: false,
            ransomTrackValue: null,
            allFonts: __allFonts,
            jpFonts: null
        };

        if (optFont.fontRandom && optFont.jpOnly) {
            optFont.jpFonts = getJPFonts();
            if (!optFont.jpFonts || optFont.jpFonts.length === 0) {
                alert(L('alertNoJPFonts'));
                return;
            }
        }

        // Preview: reflect only ransom tracking fixed value (no background)
        try {
            if (chkFontRansom && chkFontRansom.value && chkRansomTrack && chkRansomTrack.value) {
                optFont.previewRansomTracking = true;
                var __pv = parseNum(edtRansomTrk.text);
                optFont.ransomTrackValue = (typeof __pv === "number" && !isNaN(__pv)) ? __pv : 200;
            }
        } catch (_) { }

        restoreOriginals();

        // NOTE: 犯行声明文風（背景生成）はプレビューでは行わない

        applyRandom(base, hPct, hPct, true, rot, kern, 0, seed, optFont);
    }

    edtBase.onChanging = function () { syncFromEdit(edtBase, sldBase); updatePreview(); };
    edtH.onChanging = function () { syncFromEdit(edtH, sldH); updatePreview(); };
    edtRot.onChanging = function () { syncFromEdit(edtRot, sldRot); updatePreview(); };
    edtKern.onChanging = function () { syncFromEdit(edtKern, sldKern); updatePreview(); };
    edtRansomTrk.onChanging = function () { try { updatePreview(); } catch (_) { } };

    changeValueByArrowKey(edtBase);
    changeValueByArrowKey(edtH);
    changeValueByArrowKey(edtRot);
    changeValueByArrowKey(edtKern);
    changeValueByArrowKey(edtRansomTrk);


    sldBase.onChanging = function () { syncFromSlider(sldBase, edtBase); updatePreview(); };
    sldH.onChanging = function () { syncFromSlider(sldH, edtH); updatePreview(); };
    sldRot.onChanging = function () { syncFromSlider(sldRot, edtRot); updatePreview(); };
    sldKern.onChanging = function () { syncFromSlider(sldKern, edtKern); updatePreview(); };


    btnOK.onClick = function () {
        var base = 0;
        if (chkBase.value) base = parseNum(edtBase.text);

        var hPct = 0;
        if (chkScale.value) hPct = parseNum(edtH.text);

        var rot = 0;
        if (chkRot.value) rot = parseNum(edtRot.text);

        var kern = 0;
        if (chkKern.value) kern = parseNum(edtKern.text);


        // ransom tracking fixed value validation (OK時のみ)
        var __rtv = null;
        try {
            if (chkFontRansom && chkFontRansom.value && chkRansomTrack && chkRansomTrack.value) {
                __rtv = parseNum(edtRansomTrk.text);
                if (__rtv === null) {
                    alert(L("alertEnterNumber"));
                    return;
                }
            }
        } catch (_) { __rtv = null; }

        if ((chkBase.value && base === null) ||
            (chkScale.value && hPct === null) ||
            (chkRot.value && rot === null) ||
            (chkKern.value && kern === null)) {
            alert(L("alertEnterNumber"));
            return;
        }

        var optFont = {
            fontRandom: (chkFontRandom && chkFontRandom.value),
            jpOnly: (chkFontJPOnly && chkFontJPOnly.value),
            ransom: (chkFontRansom && chkFontRansom.value),
            ransomTrack: (chkRansomTrack && chkRansomTrack.value),
            ransomTrackValue: __rtv,
            allFonts: __allFonts,
            jpFonts: null
        };

        if (optFont.fontRandom && optFont.jpOnly) {
            optFont.jpFonts = getJPFonts();
            if (!optFont.jpFonts || optFont.jpFonts.length === 0) {
                alert(L('alertNoJPFonts'));
                return;
            }
        }

        // 1) restore originals then apply final settings
        restoreOriginals();
        applyRandom(base, hPct, hPct, true, rot, kern, 0, seed, optFont);

        // 2) manifesto style: background rects (OK時のみ)
        try {
            if (optFont.ransom) createBackgroundRectsByOutlining(selTextFrames);
            else clearBackgroundRectsIfAny();
        } catch (_) { }

        // 3) close dialog
        w.close(1);
    };

    btnRerun.onClick = function () {
        // rerun
        try { seed = (new Date()).getTime() & 0xffffffff; } catch (_) { }
        try { restoreOriginals(); } catch (_) { }
        try { updatePreview(); } catch (_) { }
        try { autoEnableJPOnlyIfNeeded(); } catch (_) { }
        try { updateRerunEnabled(); } catch (_) { }
    };

    btnReset.onClick = function () {
        // Reset selected text attributes only (do not touch any checkboxes)
        try { restoreOriginals(); } catch (_) { }
        try { clearBackgroundRectsIfAny(); } catch (_) { }

        // Reset baseline / scale / rotation / kerning / tracking
        try {
            for (var i = 0; i < originals.length; i++) {
                try {
                    var ch = originals[i].ch;
                    var ca = ch.characterAttributes;

                    // baseline
                    ca.baselineShift = 0;
                    originals[i].baselineShift = 0;

                    // scale
                    ca.horizontalScale = 100;
                    ca.verticalScale = 100;
                    originals[i].hScale = 100;
                    originals[i].vScale = 100;

                    // rotation
                    ca.rotation = 0;
                    originals[i].rotation = 0;

                    // kerning
                    try {
                        ca.kerningMethod = AutoKernType.NOAUTOKERN;
                        ch.kerning = 0;
                        originals[i].kerning = 0;
                    } catch (_) { }

                    // tracking
                    try {
                        ca.tracking = 0;
                        originals[i].tracking = 0;
                    } catch (_) { }

                } catch (_) { }
            }
        } catch (_) { }

        try { app.redraw(); } catch (_) { }

        // Keep JP-only auto rule consistent (checkbox state may remain as-is)
        try { autoEnableJPOnlyIfNeeded(); } catch (_) { }
        try { updateRerunEnabled(); } catch (_) { }
    };

    btnCancel.onClick = function () {
        try { restoreOriginals(); } catch (_) { }
        try { clearBackgroundRectsIfAny(); } catch (_) { }
        try { app.redraw(); } catch (_) { }
        w.close(0);
    };

    syncFromEdit(edtBase, sldBase);
    syncFromEdit(edtH, sldH);
    syncFromEdit(edtRot, sldRot);
    syncFromEdit(edtKern, sldKern);

    // Checkbox events: update preview on click
    chkBase.onClick = function () { updatePreview(); updateRerunEnabled(); };
    chkScale.onClick = function () { updatePreview(); updateRerunEnabled(); };
    chkRot.onClick = function () { updatePreview(); updateRerunEnabled(); };
    chkKern.onClick = function () { updatePreview(); updateRerunEnabled(); };

    // 文字タッチ: すべてON / すべてOFF
    btnAllOn.onClick = function () {
        try {
            chkBase.value = true;
            chkScale.value = true;
            chkRot.value = true;
            chkKern.value = true;

        } catch (_) { }
        try { updatePreview(); } catch (_) { }
        try { updateRerunEnabled(); } catch (_) { }
    };

    btnAllOff.onClick = function () {
        try {
            chkBase.value = false;
            chkScale.value = false;
            chkRot.value = false;
            chkKern.value = false;

        } catch (_) { }
        try { updatePreview(); } catch (_) { }
        try { updateRerunEnabled(); } catch (_) { }
    };

    chkFontRandom.onClick = function () { updatePreview(); updateRerunEnabled(); };
    chkFontJPOnly.onClick = function () { updatePreview(); updateRerunEnabled(); };
    chkFontRansom.onClick = function () {
        // NOTE: 犯行声明文風はOK時のみ実行（プレビュー不要）
        try {
            var en = (chkFontRansom && chkFontRansom.value);
            chkRansomTrack.enabled = en;
            edtRansomTrk.enabled = (en && chkRansomTrack && chkRansomTrack.value);
        } catch (_) { }
        try { updateRerunEnabled(); } catch (_) { }
        try { updatePreview(); } catch (_) { }
    };

    chkRansomTrack.onClick = function () {
        try {
            edtRansomTrk.enabled = (chkFontRansom && chkFontRansom.value) && (chkRansomTrack && chkRansomTrack.value);
        } catch (_) { }
        try { updatePreview(); } catch (_) { }
    };

    updatePreview();
    updateRerunEnabled();

    setDialogOpacity(w, dialogOpacity);
    shiftDialogPosition(w, offsetX, offsetY);
    w.show();
})();
