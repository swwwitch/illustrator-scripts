#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {
    
    var SCRIPT_VERSION = "v1.0.1";

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
        labelBaseline: { ja: "ベースライン:", en: "Baseline:" },
        labelScale: { ja: "比率:", en: "Scale:" },
        labelRotation: { ja: "文字回転:", en: "Rotation:" },
        labelKerning: { ja: "カーニング:", en: "Kerning:" },
        labelTracking: { ja: "トラッキング:", en: "Tracking:" },
        btnRerun: { ja: "再実行", en: "Rerun" },
        btnReset: { ja: "リセット", en: "Reset" },
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

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 他スクリプトでもこのブロックをコピペすれば再利用できます。
     * 使い方:
     *   PreviewHistory.start();     // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();      // プレビュー描画ごとにカウント(+1)
     *   PreviewHistory.undo();      // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */
    (function (g) {
        if (!g.PreviewHistory) {
            g.PreviewHistory = {
                start: function () { g.__previewUndoCount = 0; },
                bump: function () { g.__previewUndoCount = (g.__previewUndoCount | 0) + 1; },
                undo: function () {
                    var n = g.__previewUndoCount | 0;
                    try { for (var i = 0; i < n; i++) app.executeMenuCommand('undo'); } catch (e) { }
                    g.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

    if (app.documents.length === 0) { alert(L("alertNoDoc")); return; }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) { alert(L("alertSelectText")); return; }

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

    function getKerningSafe(ch) {
        try {
            var k = ch.kerning; // may throw (e.g. Error 9551)
            return (typeof k === "number") ? k : 0;
        } catch (e) {
            return 0;
        }
    }

    var ranges = collectTextRanges(doc.selection);
    if (ranges.length === 0) { alert(L("alertSelectTextRange")); return; }

    // --- snapshot originals (per character) ---
    // we store kerning and tracking as well
    var originals = []; // { ch, baselineShift, hScale, vScale, rotation, kerning, tracking }
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
                    tracking: (typeof ca.tracking === "number") ? ca.tracking : 0
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
    function applyRandom(maxAbsPt, maxAbsHPct, maxAbsVPct, linkScale, maxAbsDeg, maxAbsKern, maxAbsTrk, seed) {
        var rng = makeRng(seed);

        for (var i = 0; i < originals.length; i++) {
            var ca = originals[i].ch.characterAttributes;

            // baseline
            ca.baselineShift = randSym(rng) * maxAbsPt;

            // rotation (added to original rotation)
            ca.rotation = originals[i].rotation + (randSym(rng) * maxAbsDeg);

            // scale multipliers
            if (linkScale) {
                var m = 1.0 + (randSym(rng) * (maxAbsHPct / 100.0));
                ca.horizontalScale = originals[i].hScale * m;
                ca.verticalScale   = originals[i].vScale * m;
            } else {
                var mh = 1.0 + (randSym(rng) * (maxAbsHPct / 100.0));
                var mv = 1.0 + (randSym(rng) * (maxAbsVPct / 100.0));
                ca.horizontalScale = originals[i].hScale * mh;
                ca.verticalScale   = originals[i].vScale * mv;
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
                ca.tracking = originals[i].tracking + (randSym(rng) * maxAbsTrk);
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

    try { PreviewHistory.start(); } catch (_) { }

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

    var gBase = w.add("group");
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
    } catch (_) {}
    var edtBase = gBase.add("edittext", undefined, String(defaultBase));
    edtBase.characters = 4;
    var stUnitBase = gBase.add("statictext", undefined, "pt");
    var sldBase = gBase.add("slider", undefined, defaultBase, 0, 50);
    sldBase.preferredSize.width = 180;

    var gH = w.add("group");
    var chkScale = gH.add("checkbox", undefined, "");
    chkScale.value = true;
    chkScale.preferredSize.width = 15;
    var stH = gH.add("statictext", undefined, L("labelScale"));
    var edtH = gH.add("edittext", undefined, "10");
    edtH.characters = 4;
    var stUnitH = gH.add("statictext", undefined, "%");
var sldH = gH.add("slider", undefined, 10, 0, 200);
    sldH.preferredSize.width = 180;

    var gRot = w.add("group");
    var chkRot = gRot.add("checkbox", undefined, "");
    chkRot.value = true;
    chkRot.preferredSize.width = 15;
    var stRot = gRot.add("statictext", undefined, L("labelRotation"));
    var edtRot = gRot.add("edittext", undefined, "5");
    edtRot.characters = 4;
    var stUnitRot = gRot.add("statictext", undefined, "°");
    var sldRot = gRot.add("slider", undefined, 5, 0, 30);
    sldRot.preferredSize.width = 180;

    var gKern = w.add("group");
    var chkKern = gKern.add("checkbox", undefined, "");
    chkKern.value = true;
    chkKern.preferredSize.width = 15;
    var stKern = gKern.add("statictext", undefined, L("labelKerning"));
    var edtKern = gKern.add("edittext", undefined, "-50");
    edtKern.characters = 4;
    var stUnitKern = gKern.add("statictext", undefined, "/1000em");
    var sldKern = gKern.add("slider", undefined, -50, -200, 200);
    sldKern.preferredSize.width = 180;

    var gTrk = w.add("group");
    var chkTrk = gTrk.add("checkbox", undefined, "");
    chkTrk.value = true;
    chkTrk.preferredSize.width = 15;
    var stTrk = gTrk.add("statictext", undefined, L("labelTracking"));
    var edtTrk = gTrk.add("edittext", undefined, "0");
    edtTrk.characters = 4;
    var stUnitTrk = gTrk.add("statictext", undefined, "/1000em");
    var sldTrk = gTrk.add("slider", undefined, 0, -200, 200);
    sldTrk.preferredSize.width = 180;

    // --- label width align (right) ---
    try {
        var labels = [stBase, stH, stRot, stKern, stTrk];
        var maxW = 0;
        for (var i = 0; i < labels.length; i++) {
            try {
                var w0 = labels[i].preferredSize.width;
                if (w0 > maxW) maxW = w0;
            } catch (_) {}
        }
        for (var j = 0; j < labels.length; j++) {
            try {
                labels[j].preferredSize.width = maxW;
                labels[j].justify = "right";
            } catch (_) {}
        }
    } catch (_) {}

    // --- unit width align (left) ---
    try {
        var units = [stUnitBase, stUnitH, stUnitRot, stUnitKern, stUnitTrk];
        var maxU = 0;
        for (var ui = 0; ui < units.length; ui++) {
            try {
                var uw = units[ui].preferredSize.width;
                if (uw > maxU) maxU = uw;
            } catch (_) {}
        }
        for (var uj = 0; uj < units.length; uj++) {
            try {
                units[uj].preferredSize.width = maxU;
                units[uj].justify = "left";
            } catch (_) {}
        }
    } catch (_) {}

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

            if (editText !== edtKern && editText !== edtTrk && value < 0) value = 0;

            event.preventDefault();
            editText.text = String(value);

            // keep slider in sync (if exists)
            try {
                if (editText === edtBase) syncFromEdit(edtBase, sldBase);
                else if (editText === edtH) syncFromEdit(edtH, sldH);
                else if (editText === edtRot) syncFromEdit(edtRot, sldRot);
                else if (editText === edtKern) syncFromEdit(edtKern, sldKern);
                else if (editText === edtTrk) syncFromEdit(edtTrk, sldTrk);
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
    var btnReset = btnLeft.add("button", undefined, L("btnReset"));

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

        var trk = 0;
        if (chkTrk.value) { trk = parseNum(edtTrk.text); if (trk === null) return; }

        restoreOriginals();
        applyRandom(base, hPct, hPct, true, rot, kern, trk, seed);
        try { PreviewHistory.bump(); } catch (_) { }
    }

    edtBase.onChanging = function () { syncFromEdit(edtBase, sldBase); updatePreview(); };
    edtH.onChanging    = function () { syncFromEdit(edtH, sldH); updatePreview(); };
    edtRot.onChanging  = function () { syncFromEdit(edtRot, sldRot); updatePreview(); };
    edtKern.onChanging  = function () { syncFromEdit(edtKern, sldKern); updatePreview(); };
    edtTrk.onChanging   = function () { syncFromEdit(edtTrk, sldTrk);   updatePreview(); };

    changeValueByArrowKey(edtBase);
    changeValueByArrowKey(edtH);
    changeValueByArrowKey(edtRot);
    changeValueByArrowKey(edtKern);
    changeValueByArrowKey(edtTrk);

    sldBase.onChanging = function () { syncFromSlider(sldBase, edtBase); updatePreview(); };
    sldH.onChanging    = function () { syncFromSlider(sldH, edtH);       updatePreview(); };
    sldRot.onChanging  = function () { syncFromSlider(sldRot, edtRot);   updatePreview(); };
    sldKern.onChanging  = function () { syncFromSlider(sldKern, edtKern); updatePreview(); };
    sldTrk.onChanging   = function () { syncFromSlider(sldTrk, edtTrk);   updatePreview(); };


    btnOK.onClick = function () {
        var base = 0;
        if (chkBase.value) base = parseNum(edtBase.text);

        var hPct = 0;
        if (chkScale.value) hPct = parseNum(edtH.text);

        var rot = 0;
        if (chkRot.value) rot = parseNum(edtRot.text);

        var kern = 0;
        if (chkKern.value) kern = parseNum(edtKern.text);

        var trk = 0;
        if (chkTrk.value) trk = parseNum(edtTrk.text);

        if ((chkBase.value && base === null) ||
            (chkScale.value && hPct === null) ||
            (chkRot.value && rot === null) ||
            (chkKern.value && kern === null) ||
            (chkTrk.value && trk === null)) {
            alert(L("alertEnterNumber"));
            return;
        }

        try { PreviewHistory.undo(); } catch (_) { }
        restoreOriginals();
        applyRandom(base, hPct, hPct, true, rot, kern, trk, seed);
        try { PreviewHistory.start(); } catch (_) { }
        w.close(1);
    };

    btnRerun.onClick = function () {
        var allOff = false;
        try {
            allOff = (!chkBase.value && !chkScale.value && !chkRot.value && !chkKern.value && !chkTrk.value);
        } catch (_) { allOff = false; }

        if (allOff) {
            // all OFF -> restore defaults and enable all
            try {
                chkBase.value = true;
                chkScale.value = true;
                chkRot.value = true;
                chkKern.value = true;
                chkTrk.value = true;
            } catch (_) { }

            try {
                edtBase.text = String(defaultBase);
                edtH.text = "10";
                edtRot.text = "5";
                edtKern.text = "-50";
                edtTrk.text = "0";
            } catch (_) { }

            try {
                syncFromEdit(edtBase, sldBase);
                syncFromEdit(edtH, sldH);
                syncFromEdit(edtRot, sldRot);
                syncFromEdit(edtKern, sldKern);
                syncFromEdit(edtTrk, sldTrk);
            } catch (_) { }
        }

        // rerun
        try { seed = (new Date()).getTime() & 0xffffffff; } catch (_) { }
        try { restoreOriginals(); } catch (_) { }
        try { updatePreview(); } catch (_) { }
    };

    btnReset.onClick = function () {
        // clear preview history and restore original text only
        try { PreviewHistory.undo(); } catch (_) { }
        try { restoreOriginals(); } catch (_) { }

        // set scale to 100% (actual text) / 水平・垂直比率を100%に
        try {
            for (var i = 0; i < originals.length; i++) {
                try {
                    var ca = originals[i].ch.characterAttributes;
                    ca.horizontalScale = 100;
                    ca.verticalScale = 100;
                    originals[i].hScale = 100;
                    originals[i].vScale = 100;
                } catch (_) { }
            }
        } catch (_) { }

        try { app.redraw(); } catch (_) { }
        try { PreviewHistory.start(); } catch (_) { }

        // turn all checkboxes OFF (do not change edit fields)
        try {
            chkBase.value = false;
            chkScale.value = false;
            chkRot.value = false;
            chkKern.value = false;
            chkTrk.value = false;
        } catch (_) { }
    };

    btnCancel.onClick = function () {
        try { PreviewHistory.undo(); } catch (_) { }
        try { restoreOriginals(); } catch (_) { }
        try { app.redraw(); } catch (_) { }
        w.close(0);
    };

    syncFromEdit(edtBase, sldBase);
    syncFromEdit(edtH, sldH);
    syncFromEdit(edtRot, sldRot);
    syncFromEdit(edtKern, sldKern);
    syncFromEdit(edtTrk, sldTrk);

    // Checkbox events: update preview on click
    chkBase.onClick = function () { updatePreview(); };
    chkScale.onClick = function () { updatePreview(); };
    chkRot.onClick = function () { updatePreview(); };
    chkKern.onClick = function () { updatePreview(); };
    chkTrk.onClick = function () { updatePreview(); };

    updatePreview();

    setDialogOpacity(w, dialogOpacity);
    shiftDialogPosition(w, offsetX, offsetY);
    w.show();
})();
