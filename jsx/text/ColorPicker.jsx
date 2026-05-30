#target illustrator
/*
--------------------------------------------------
Reusable Illustrator Color Picker Library
--------------------------------------------------
Usage:

#include "ColorPicker.jsx"

var result = ColorPicker.show({
    value: "FF0000",      // "RRGGBB" or "cmyk:C,M,Y,K"
    title: "Color Picker"
});

if (result !== null) {
    $.writeln(result);
}

Return values:
RGB  -> "RRGGBB"
CMYK -> "cmyk:C,M,Y,K"
--------------------------------------------------
*/

var ColorPicker = (function () {

    var _dialogPos = null;

    var L = {
        white:  { ja: "ホワイト", en: "White" },
        black:  { ja: "ブラック", en: "Black" },
        custom: { ja: "カスタム", en: "Custom" },
        gray:   { ja: "グレー",   en: "Gray" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok:     { ja: "OK",       en: "OK" }
    };

    function ll(key, lng) {
        return (L[key] && L[key][lng]) ? L[key][lng] : (L[key] ? L[key].en : key);
    }

    var DEFAULT_SWATCHES = [
        "FF0000", "FFCC00", "FFFF00",
        "00CC00", "0066FF",
        "FF99CC", "996633", "666666",
        "999999"
    ];


    function isCmykString(s) {
        return String(s).indexOf("cmyk:") === 0;
    }

    function parseCmykString(s) {
        var p = String(s).replace("cmyk:", "").split(",");
        return {
            c: Number(p[0]) || 0,
            m: Number(p[1]) || 0,
            y: Number(p[2]) || 0,
            k: Number(p[3]) || 0
        };
    }

    function cmykStringFromValues(c, m, y, k) {
        return "cmyk:" + Math.round(c) + "," + Math.round(m) + "," + Math.round(y) + "," + Math.round(k);
    }

    function rgbToHex(r, g, b) {
        function h(n) {
            var s = Math.round(n).toString(16).toUpperCase();
            return s.length < 2 ? "0" + s : s;
        }
        return h(r) + h(g) + h(b);
    }

    function hexToRGB(hex) {
        hex = String(hex || "").replace(/^#/, "");
        if (hex.length !== 6) hex = "000000";
        return {
            r: parseInt(hex.substring(0, 2), 16) || 0,
            g: parseInt(hex.substring(2, 4), 16) || 0,
            b: parseInt(hex.substring(4, 6), 16) || 0
        };
    }

    function cmykToRgbApprox(c, m, y, k) {
        var r = 255 * (1 - c / 100) * (1 - k / 100);
        var g = 255 * (1 - m / 100) * (1 - k / 100);
        var b = 255 * (1 - y / 100) * (1 - k / 100);
        return {
            r: Math.round(r),
            g: Math.round(g),
            b: Math.round(b)
        };
    }

    function rgbToCmykApprox(r, g, b) {
        var rr = r / 255;
        var gg = g / 255;
        var bb = b / 255;
        var k = 1 - Math.max(rr, gg, bb);

        if (k >= 1) {
            return { c: 0, m: 0, y: 0, k: 100 };
        }

        var c = (1 - rr - k) / (1 - k) * 100;
        var m = (1 - gg - k) / (1 - k) * 100;
        var y = (1 - bb - k) / (1 - k) * 100;

        return {
            c: Math.round(c),
            m: Math.round(m),
            y: Math.round(y),
            k: Math.round(k * 100)
        };
    }

    function clamp(value, min, max) {
        var v = Number(value);
        if (isNaN(v)) v = 0;
        if (v < min) v = min;
        if (v > max) v = max;
        return v;
    }

    function createInitialState(value) {
        var state = {
            preset: "custom",   // white | black | custom
            mode: "rgb",        // rgb | cmyk | gray
            dialogTab: "rgb",   // rgb | cmyk
            rgb: { r: 0, g: 0, b: 0 },
            cmyk: { c: 0, m: 0, y: 0, k: 0 },
            original: { r: 0, g: 0, b: 0 }
        };

        if (isCmykString(value)) {
            var cv = parseCmykString(value);
            var rgb = cmykToRgbApprox(cv.c, cv.m, cv.y, cv.k);
            state.cmyk = { c: cv.c, m: cv.m, y: cv.y, k: cv.k };
            state.rgb = { r: rgb.r, g: rgb.g, b: rgb.b };
            state.original = { r: rgb.r, g: rgb.g, b: rgb.b };
            state.dialogTab = "cmyk";
            state.mode = (cv.c === 0 && cv.m === 0 && cv.y === 0 && cv.k > 0) ? "gray" : "cmyk";

            if (cv.c === 0 && cv.m === 0 && cv.y === 0 && cv.k === 0) state.preset = "white";
            else if (cv.c === 0 && cv.m === 0 && cv.y === 0 && cv.k === 100) state.preset = "black";
        } else {
            var rgb0 = hexToRGB(value || "000000");
            state.rgb = { r: rgb0.r, g: rgb0.g, b: rgb0.b };
            state.original = { r: rgb0.r, g: rgb0.g, b: rgb0.b };
            state.cmyk = rgbToCmykApprox(rgb0.r, rgb0.g, rgb0.b);
            state.dialogTab = "rgb";
            state.mode = "rgb";

            if (rgb0.r === 255 && rgb0.g === 255 && rgb0.b === 255) state.preset = "white";
            else if (rgb0.r === 0 && rgb0.g === 0 && rgb0.b === 0) state.preset = "black";
        }

        return state;
    }

    function syncCmykFromRgb(state) {
        state.cmyk = rgbToCmykApprox(
            Math.round(state.rgb.r),
            Math.round(state.rgb.g),
            Math.round(state.rgb.b)
        );
    }

    function syncRgbFromCmyk(state) {
        var rgb = cmykToRgbApprox(state.cmyk.c, state.cmyk.m, state.cmyk.y, state.cmyk.k);
        state.rgb = { r: rgb.r, g: rgb.g, b: rgb.b };
    }

    function getPreviewRgb(state) {
        if (state.preset === "white") return { r: 255, g: 255, b: 255 };
        if (state.preset === "black") return { r: 0, g: 0, b: 0 };
        if (state.mode === "cmyk" || state.mode === "gray") {
            return cmykToRgbApprox(state.cmyk.c, state.cmyk.m, state.cmyk.y, state.cmyk.k);
        }
        return { r: state.rgb.r, g: state.rgb.g, b: state.rgb.b };
    }

    function serializeState(state) {
        if (state.preset === "white") return "FFFFFF";
        if (state.preset === "black") return "000000";
        if (state.mode === "cmyk" || state.mode === "gray") {
            return cmykStringFromValues(state.cmyk.c, state.cmyk.m, state.cmyk.y, state.cmyk.k);
        }
        return rgbToHex(state.rgb.r, state.rgb.g, state.rgb.b);
    }

    function createSlider(parent, label, value, maxValue) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];

        var st = row.add("statictext", undefined, label);
        st.preferredSize = [18, -1];

        var slider = row.add("slider", undefined, value, 0, maxValue);
        slider.preferredSize = [140, 20];

        var edit = row.add("edittext", undefined, String(Math.round(value)));
        edit.characters = 3;

        slider.onChanging = function () {
            edit.text = String(Math.round(slider.value));
        };

        edit.onChange = function () {
            slider.value = clamp(edit.text, 0, maxValue);
            edit.text = String(Math.round(slider.value));
        };

        return {
            row: row,
            slider: slider,
            edit: edit
        };
    }

    function buildUI(state, title, lng) {
        var dlg = new Window("dialog", title || "Color Picker");
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = 14;

        var previewRow = dlg.add("group");
        previewRow.orientation = "row";
        previewRow.alignment = ["center", "top"];
        previewRow.spacing = 1;

        var previewBefore = previewRow.add("group");
        previewBefore.preferredSize = [90, 40];

        var previewAfter = previewRow.add("group");
        previewAfter.preferredSize = [90, 40];

        previewBefore.onDraw = function () {
            var g = this.graphics;
            var rgb = state.original;
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
            g.rectPath(0, 0, this.size[0], this.size[1]);
            g.fillPath(brush);
        };

        previewAfter.onDraw = function () {
            var g = this.graphics;
            var rgb = getPreviewRgb(state);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
            g.rectPath(0, 0, this.size[0], this.size[1]);
            g.fillPath(brush);
        };

        var presetRow = dlg.add("group");
        presetRow.orientation = "row";
        presetRow.alignment = ["center", "top"];
        presetRow.alignChildren = ["left", "center"];
        var rbWhite = presetRow.add("radiobutton", undefined, ll("white", lng));
        var rbBlack = presetRow.add("radiobutton", undefined, ll("black", lng));
        var rbCustom = presetRow.add("radiobutton", undefined, ll("custom", lng));

        var swatchPanel = dlg.add("group");
        swatchPanel.orientation = "row";
        swatchPanel.alignment = ["center", "top"];
        swatchPanel.spacing = 2;
        var swatchItems = [];
        for (var si = 0; si < DEFAULT_SWATCHES.length; si++) {
            (function (idx) {
                var hex = DEFAULT_SWATCHES[idx];
                var rgb = hexToRGB(hex);
                var sw = swatchPanel.add("group");
                sw.preferredSize = [16, 16];
                sw.onDraw = function () {
                    var gg = this.graphics;
                    var pen = gg.newPen(gg.PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
                    var brush = gg.newBrush(gg.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
                    gg.rectPath(0, 0, this.size[0], this.size[1]);
                    gg.fillPath(brush);
                    gg.rectPath(0, 0, this.size[0], this.size[1]);
                    gg.strokePath(pen);
                };
                swatchItems.push({ element: sw, hex: hex });
            })(si);
        }

        var tabs = dlg.add("tabbedpanel");
        tabs.alignChildren = ["fill", "top"];

        var tabRGB = tabs.add("tab", undefined, "RGB");
        tabRGB.orientation = "column";
        tabRGB.margins = [14, 18, 14, 10];

        var tabCMYK = tabs.add("tab", undefined, "CMYK");
        tabCMYK.orientation = "column";
        tabCMYK.margins = [14, 18, 14, 10];

        var r = createSlider(tabRGB, "R", state.rgb.r, 255);
        var g = createSlider(tabRGB, "G", state.rgb.g, 255);
        var b = createSlider(tabRGB, "B", state.rgb.b, 255);

        tabRGB.add("panel").preferredSize.height = 10; // spacer

        var hexRow = tabRGB.add("group");
        hexRow.orientation = "row";
        hexRow.add("statictext", undefined, "#");
        var etHex = hexRow.add("edittext", undefined, rgbToHex(state.rgb.r, state.rgb.g, state.rgb.b));
        etHex.characters = 6;

        var cbGray = tabCMYK.add("checkbox", undefined, ll("gray", lng));
        var c = createSlider(tabCMYK, "C", state.cmyk.c, 100);
        var m = createSlider(tabCMYK, "M", state.cmyk.m, 100);
        var y = createSlider(tabCMYK, "Y", state.cmyk.y, 100);
        var k = createSlider(tabCMYK, "K", state.cmyk.k, 100);

        var btns = dlg.add("group");
        btns.alignment = ["center", "center"];
        btns.add("button", undefined, ll("cancel", lng), { name: "cancel" });
        btns.add("button", undefined, ll("ok", lng), { name: "ok" });

        return {
            dlg: dlg,
            previewBefore: previewBefore,
            previewAfter: previewAfter,
            rbWhite: rbWhite,
            rbBlack: rbBlack,
            rbCustom: rbCustom,
            tabs: tabs,
            tabRGB: tabRGB,
            tabCMYK: tabCMYK,
            cbGray: cbGray,
            r: r,
            g: g,
            b: b,
            c: c,
            m: m,
            y: y,
            k: k,
            etHex: etHex,
            swatchItems: swatchItems
        };
    }

    function render(state, ui, options) {
        options = options || {};
        var suppressTabSelection = !!options.suppressTabSelection;

        ui.rbWhite.value = (state.preset === "white");
        ui.rbBlack.value = (state.preset === "black");
        ui.rbCustom.value = (state.preset === "custom");

        if (!suppressTabSelection) {
            var targetTab = (state.dialogTab === "cmyk") ? ui.tabCMYK : ui.tabRGB;
            try {
                if (ui.tabs.selection !== targetTab) {
                    ui.tabs.selection = targetTab;
                }
            } catch (eTab) {}
        }

        ui.cbGray.value = (state.mode === "gray");

        ui.r.slider.value = state.rgb.r;
        ui.r.edit.text = String(state.rgb.r);
        ui.g.slider.value = state.rgb.g;
        ui.g.edit.text = String(state.rgb.g);
        ui.b.slider.value = state.rgb.b;
        ui.b.edit.text = String(state.rgb.b);
        ui.etHex.text = rgbToHex(state.rgb.r, state.rgb.g, state.rgb.b);

        ui.c.slider.value = state.cmyk.c;
        ui.c.edit.text = String(state.cmyk.c);
        ui.m.slider.value = state.cmyk.m;
        ui.m.edit.text = String(state.cmyk.m);
        ui.y.slider.value = state.cmyk.y;
        ui.y.edit.text = String(state.cmyk.y);
        ui.k.slider.value = state.cmyk.k;
        ui.k.edit.text = String(state.cmyk.k);

        var customEnabled = (state.preset === "custom");
        ui.tabs.enabled = customEnabled;
        ui.c.row.enabled = customEnabled && state.mode !== "gray";
        ui.m.row.enabled = customEnabled && state.mode !== "gray";
        ui.y.row.enabled = customEnabled && state.mode !== "gray";

        try { ui.previewAfter.hide(); ui.previewAfter.show(); } catch (ePreview) {}
    }

    function bindEvents(state, ui) {
        var syncing = false;

        function safeRender(options) {
            if (syncing) return;
            syncing = true;
            try {
                render(state, ui, options);
            } finally {
                syncing = false;
            }
        }

        function setPreset(nextPreset) {
            state.preset = nextPreset;

            if (nextPreset === "white") {
                state.rgb = { r: 255, g: 255, b: 255 };
                syncCmykFromRgb(state);
                if (state.dialogTab !== "cmyk") state.mode = "rgb";
            } else if (nextPreset === "black") {
                state.rgb = { r: 0, g: 0, b: 0 };
                syncCmykFromRgb(state);
                if (state.dialogTab !== "cmyk") state.mode = "rgb";
            } else {
                if (state.dialogTab === "cmyk") {
                    if (state.mode !== "gray") state.mode = "cmyk";
                    syncRgbFromCmyk(state);
                } else {
                    state.mode = "rgb";
                    syncCmykFromRgb(state);
                }
            }
        }

        function setRgb(r, g, b) {
            state.rgb.r = Math.round(clamp(r, 0, 255));
            state.rgb.g = Math.round(clamp(g, 0, 255));
            state.rgb.b = Math.round(clamp(b, 0, 255));
            syncCmykFromRgb(state);
            state.mode = "rgb";
            state.dialogTab = "rgb";
            state.preset = "custom";
        }

        function setCmyk(c, m, y, k, keepGray) {
            state.cmyk.c = Math.round(clamp(c, 0, 100));
            state.cmyk.m = Math.round(clamp(m, 0, 100));
            state.cmyk.y = Math.round(clamp(y, 0, 100));
            state.cmyk.k = Math.round(clamp(k, 0, 100));
            syncRgbFromCmyk(state);
            state.mode = keepGray ? "gray" : "cmyk";
            state.dialogTab = "cmyk";
            state.preset = "custom";
        }

        var rgbRows = [ui.r, ui.g, ui.b];
        for (var i = 0; i < rgbRows.length; i++) {
            (function (idx) {
                var row = rgbRows[idx];
                var orig = row.slider.onChanging;
                row.slider.onChanging = function () {
                    orig.call(row.slider);
                    setRgb(ui.r.slider.value, ui.g.slider.value, ui.b.slider.value);
                    safeRender();
                };
                row.edit.onChange = function () {
                    row.slider.value = clamp(row.edit.text, 0, 255);
                    row.edit.text = String(Math.round(row.slider.value));
                    setRgb(ui.r.slider.value, ui.g.slider.value, ui.b.slider.value);
                    safeRender();
                };
            })(i);
        }

        var cmykRows = [ui.c, ui.m, ui.y, ui.k];
        for (var j = 0; j < cmykRows.length; j++) {
            (function (idx2) {
                var row2 = cmykRows[idx2];
                var orig2 = row2.slider.onChanging;
                row2.slider.onChanging = function () {
                    orig2.call(row2.slider);
                    setCmyk(ui.c.slider.value, ui.m.slider.value, ui.y.slider.value, ui.k.slider.value, ui.cbGray.value);
                    safeRender();
                };
                row2.edit.onChange = function () {
                    row2.slider.value = clamp(row2.edit.text, 0, 100);
                    row2.edit.text = String(Math.round(row2.slider.value));
                    setCmyk(ui.c.slider.value, ui.m.slider.value, ui.y.slider.value, ui.k.slider.value, ui.cbGray.value);
                    safeRender();
                };
            })(j);
        }

        ui.etHex.onChange = function () {
            var h = String(ui.etHex.text).replace(/^#/, "");
            if (h.length !== 6) return;
            var rgb = hexToRGB(h);
            setRgb(rgb.r, rgb.g, rgb.b);
            safeRender();
        };

        ui.rbWhite.onClick = function () {
            setPreset("white");
            safeRender();
        };

        ui.rbBlack.onClick = function () {
            setPreset("black");
            safeRender();
        };

        ui.rbCustom.onClick = function () {
            setPreset("custom");
            safeRender();
        };

        ui.tabs.onChange = function () {
            if (syncing) return;
            state.dialogTab = (ui.tabs.selection === ui.tabCMYK) ? "cmyk" : "rgb";
            if (state.dialogTab === "rgb") {
                syncRgbFromCmyk(state);
                state.mode = "rgb";
            } else {
                syncCmykFromRgb(state);
                state.mode = ui.cbGray.value ? "gray" : "cmyk";
            }
            safeRender({ suppressTabSelection: true });
        };

        for (var si = 0; si < ui.swatchItems.length; si++) {
            (function (idx) {
                ui.swatchItems[idx].element.addEventListener("click", function () {
                    var rgb = hexToRGB(ui.swatchItems[idx].hex);
                    setRgb(rgb.r, rgb.g, rgb.b);
                    safeRender();
                });
            })(si);
        }

        ui.cbGray.onClick = function () {
            if (ui.cbGray.value) {
                state.cmyk.c = 0;
                state.cmyk.m = 0;
                state.cmyk.y = 0;
                state.mode = "gray";
            } else {
                state.mode = "cmyk";
            }
            state.dialogTab = "cmyk";
            state.preset = "custom";
            syncRgbFromCmyk(state);
            safeRender();
        };
    }

    function show(arg) {
        var options = (typeof arg === "object" && arg !== null) ? arg : { value: arg };
        var lng = options.lang || "en";
        var state = createInitialState(options.value || "000000");
        var ui = buildUI(state, options.title || "Color Picker", lng);

        bindEvents(state, ui);
        render(state, ui);

        if (_dialogPos) {
            try { ui.dlg.location = _dialogPos; } catch (ePos) {}
        }

        var result = ui.dlg.show();

        try { _dialogPos = ui.dlg.location; } catch (ePosSave) {}

        if (result !== 1) return null;
        return serializeState(state);
    }

    return {
        show: show,
        rgbToHex: rgbToHex,
        hexToRGB: hexToRGB,
        rgbToCmykApprox: rgbToCmykApprox,
        cmykToRgbApprox: cmykToRgbApprox,
        isCmykString: isCmykString,
        parseCmykString: parseCmykString
    };

})();
