#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
選択したオブジェクトにランダム変形を適用する Illustrator スクリプトです。
塗りと線の分離、角丸、ラフ効果A / B、グループ化、プレビュー、画面ズームに対応します。
角丸と移動は rulerType の単位で入力し、角丸 → 変形 → ラフ効果B → ラフ効果A の順で適用します。
塗りと線の両方を持つオブジェクトがない場合、塗り/線パネルはディム表示になります。
クリッピンググループ内では、分離後のグループ化は行いません。
詳細の値は整数として扱います。
画面ズームには軽量モードがあり、デフォルトは OFF です。

This Illustrator script applies randomized transformations to the selected objects.
It supports fill/stroke splitting, round corners, Roughen A / B, grouping, preview, and view zoom controls.
Round Corners and Move use the current rulerType unit, and effects are applied in this order: Round Corners → Transform → Roughen B → Roughen A.
When no object has both visible fill and visible stroke, the Fill / Stroke panel is dimmed.
Inside clipping groups, split items are not regrouped.
Detail values are handled as integers.
The Zoom control includes a Light mode option, which is OFF by default.

Updated: 2026-04-14
Version policy: Do not increment the version unless explicitly instructed.
*/

(function () {

    var SCRIPT_VERSION = "v1.0.2";

    // =========================================
    // ローカライズ / localization
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "手書き・スケッチ風の調整",
            en: "Sketchy Adjustments"
        },
        alertNoDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        alertNoSelection: {
            ja: "オブジェクトを選択してください。",
            en: "Please select an object."
        },
        alertInvalidNumber: {
            ja: "数値入力が不正です。",
            en: "One or more numeric values are invalid."
        },
        panelFillStroke: {
            ja: "塗り/線",
            en: "Fill / Stroke"
        },
        splitFillStroke: {
            ja: "塗り/線を分離",
            en: "Split Fill / Stroke"
        },
        keepFillStroke: {
            ja: "分離しない",
            en: "Do Not Split"
        },
        groupSplitItems: {
            ja: "グループ化",
            en: "Group Result"
        },
        panelTransform: {
            ja: "変形",
            en: "Transform"
        },
        scale: {
            ja: "スケール",
            en: "Scale"
        },
        move: {
            ja: "移動",
            en: "Move"
        },
        rotate: {
            ja: "回転",
            en: "Rotate"
        },
        panelRoundCorners: {
            ja: "角丸",
            en: "Round Corners"
        },
        radius: {
            ja: "半径",
            en: "Radius"
        },
        panelRoughA: {
            ja: "ラフ効果A",
            en: "Roughen A"
        },
        jaggedLine: {
            ja: "ギザギザ線",
            en: "Apply Roughen"
        },
        size: {
            ja: "サイズ",
            en: "Size"
        },
        detail: {
            ja: "詳細",
            en: "Detail"
        },
        panelRoughB: {
            ja: "ラフ効果B",
            en: "Roughen B"
        },
        applyDistort: {
            ja: "歪曲を適用",
            en: "Apply Distortion"
        },
        preview: {
            ja: "再計算",
            en: "Recalculate"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        zoom: {
            ja: "画面ズーム",
            en: "Zoom"
        },
        lightMode: {
            ja: "軽量モード",
            en: "Light mode"
        }
    };

    function L(key) {
        return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
    }

    function labelText(key) {
        return L(key) + (lang === 'ja' ? '：' : ':');
    }

    // =========================================
    // 単位関連 / Units
    // =========================================

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

    function getRulerUnitCode() {
        return app.preferences.getIntegerPreference("rulerType");
    }

    function getStrokeUnitCode() {
        return app.preferences.getIntegerPreference("strokeUnits");
    }

    function getUnitLabelFromCode(code) {
        return unitLabelMap[code] || "pt";
    }

    function getCurrentRulerUnitLabel() {
        return getUnitLabelFromCode(getRulerUnitCode());
    }

    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q/H
            case 6: return 1.0;                         // px
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    function convertToPt(value, unitCode) {
        return value * getPtFactorFromUnitCode(unitCode);
    }

    if (app.documents.length === 0) {
        alert(L("alertNoDocument"));
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    var isPreviewApplied = false;
    var dlg;
    var splitFillStrokeRb, keepFillStrokeRb, groupSplitItemsCb;
    var scaleCb, moveCb, rotateCb;
    var scaleInput, moveInput, rotateInput;
    var roundRadiusCb, roundRadiusInput;
    var roughEnableCb, roughSizeInput, roughDetailInput;
    var distortEnableCb, distortSizeInput, distortDetailInput;
    var previewBtn, cancelBtn, okBtn;

    var zoomCtrl;

    /* ランダム小数 / Random decimal */
    function rand(min, max) {
        return min + Math.random() * (max - min);
    }

    /* ラフ効果を適用 / Apply roughen effect */
    function applyRoughen(obj, amount, segmentsPerInch, roundness) {
        try {
            var xml = '<LiveEffect name="Adobe Roughen">'
                + '<Dict data="'
                + 'R asiz ' + amount
                + ' R size ' + amount
                + ' R absoluteness 0'
                + ' R dtal ' + segmentsPerInch
                + ' R roundness ' + roundness
                + ' "/></LiveEffect>';
            obj.applyEffect(xml);
        } catch (_) { }
    }


    function createRoundCornersEffectXML(radiusPt) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radiusPt);
    }


    function applyRoundCorners(obj, radiusPt) {
        try {
            obj.applyEffect(createRoundCornersEffectXML(radiusPt));
        } catch (_) { }
    }

    function hasVisibleFill(obj) {
        try {
            return !!(obj && obj.filled && obj.fillColor && obj.fillColor.typename !== "NoColor");
        } catch (_) {
            return false;
        }
    }

    function hasVisibleStroke(obj) {
        try {
            return !!(obj && obj.stroked && obj.strokeColor && obj.strokeColor.typename !== "NoColor");
        } catch (_) {
            return false;
        }
    }

    function createGroupForItem(item) {
        var parent = null;
        var grp = null;

        try {
            parent = item.parent;
            if (parent && parent.groupItems) {
                grp = parent.groupItems.add();
            }
        } catch (_) { }

        if (!grp) {
            return null;
        }

        try {
            grp.move(item, ElementPlacement.PLACEBEFORE);
        } catch (_) {
            return null;
        }

        return grp;
    }

    function isInsideClippingGroup(item) {
        var parent = null;

        try {
            parent = item.parent;
        } catch (_) { }

        while (parent) {
            try {
                if (parent.typename === "GroupItem" && parent.clipped) {
                    return true;
                }
                parent = parent.parent;
            } catch (_) {
                break;
            }
        }

        return false;
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
        var sliderWidth = (typeof options.sliderWidth === "number") ? options.sliderWidth : 240;
        var doRedraw = (options.redraw !== false);

        var showLightMode = (options.lightMode !== false);
        var lightModeLabel = options.lightModeLabel || "Light mode";
        var lightModeDefault = (options.lightModeDefault === true);

        var g = parent.add("group");
        g.orientation = "row";
        g.alignChildren = ["center", "center"];
        g.alignment = "center";
        try { if (options.margins) g.margins = options.margins; } catch (_) { }

        var stLabel = g.add("statictext", undefined, String(labelText || "Zoom"));

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

                try {
                    var kb = null;
                    try { kb = ScriptUI.environment.keyboardState; } catch (_) { kb = null; }

                    var isAlt = !!(kb && kb.altKey);
                    var isShift = !!(kb && kb.shiftKey);

                    if (isAlt && isShift) {
                        try {
                            if (doc.selection && doc.selection.length > 0) {
                                app.executeMenuCommand("fitinwindow");
                                return;
                            }
                        } catch (_) { }
                    }

                    if (isAlt) {
                        try {
                            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
                            var cx = (ab[0] + ab[2]) / 2;
                            var cy = (ab[1] + ab[3]) / 2;
                            v.centerPoint = [cx, cy];
                        } catch (_) { }
                    } else if (isShift) {
                        var sel = doc.selection;
                        if (sel && sel.length > 0) {
                            var b = null;
                            for (var i = 0; i < sel.length; i++) {
                                try {
                                    var gb = sel[i].geometricBounds;
                                    if (!b) {
                                        b = [gb[0], gb[1], gb[2], gb[3]];
                                    } else {
                                        if (gb[0] < b[0]) b[0] = gb[0];
                                        if (gb[1] > b[1]) b[1] = gb[1];
                                        if (gb[2] > b[2]) b[2] = gb[2];
                                        if (gb[3] < b[3]) b[3] = gb[3];
                                    }
                                } catch (_) { }
                            }
                            if (b) {
                                var cx2 = (b[0] + b[2]) / 2;
                                var cy2 = (b[1] + b[3]) / 2;
                                v.centerPoint = [cx2, cy2];
                            }
                        }
                    }
                } catch (_) { }

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

        sld.onChanging = function () {
            if (isLightMode()) return;
            applyZoom(Number(sld.value));
        };

        sld.onChange = function () {
            applyZoom(Number(sld.value));
        };

        if (chkLight) {
            chkLight.onClick = function () {
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

    var __zoomState = __TMKZoom_captureViewState(doc);

    function hasValidPreviewInputs() {
        var scaleRange = scaleCb.value ? parseFloat(scaleInput.text) : 0;
        var moveValue = moveCb.value ? parseFloat(moveInput.text) : 0;
        var rotateRange = rotateCb.value ? parseFloat(rotateInput.text) : 0;
        var roundRadiusValue = roundRadiusCb.value ? parseFloat(roundRadiusInput.text) : 0;
        var roughSize = roughEnableCb.value ? parseFloat(roughSizeInput.text) : 0;
        var roughDetail = roughEnableCb.value ? parseFloat(roughDetailInput.text) : 0;
        var distortSize = distortEnableCb.value ? parseFloat(distortSizeInput.text) : 0;
        var distortDetail = distortEnableCb.value ? parseFloat(distortDetailInput.text) : 0;

        return !isNaN(scaleRange)
            && !isNaN(moveValue)
            && !isNaN(rotateRange)
            && !isNaN(roundRadiusValue)
            && !isNaN(roughSize)
            && !isNaN(roughDetail)
            && !isNaN(distortSize)
            && !isNaN(distortDetail);
    }

    /* メイン処理を実行 / Apply main processing */
    function applyRandomizedEffects() {
        var useScale = scaleCb.value;
        var useMove = moveCb.value;
        var useRotate = rotateCb.value;
        var scaleRange = useScale ? parseFloat(scaleInput.text) : 0;
        var moveValue = useMove ? parseFloat(moveInput.text) : 0;
        var rotateRange = useRotate ? parseFloat(rotateInput.text) : 0;
        var doGroup = groupSplitItemsCb.value;
        var roundRadiusValue = roundRadiusCb.value ? parseFloat(roundRadiusInput.text) : 0;
        var useRough = roughEnableCb.value;
        var roughSize = parseFloat(roughSizeInput.text);
        var roughDetail = Math.round(parseFloat(roughDetailInput.text));
        var roughRoundness = 1;
        var useDistort = distortEnableCb.value;
        var distortSize = parseFloat(distortSizeInput.text);
        var distortDetail = Math.round(parseFloat(distortDetailInput.text));
        var rulerUnitCode = getRulerUnitCode();
        var moveRange = convertToPt(moveValue, rulerUnitCode);
        var roundRadiusPt = convertToPt(roundRadiusValue, rulerUnitCode);

        if (isNaN(scaleRange) || isNaN(moveValue) || isNaN(rotateRange) || isNaN(roundRadiusValue)) {
            return false;
        }
        if (useRough && (isNaN(roughSize) || isNaN(roughDetail))) {
            return false;
        }
        if (useDistort && (isNaN(distortSize) || isNaN(distortDetail))) {
            return false;
        }

        var scaleMin = 100 - scaleRange;
        var scaleMax = 100 + scaleRange;
        var rotateMin = -rotateRange;
        var rotateMax = rotateRange;

        var items = [];
        for (var i = 0; i < doc.selection.length; i++) {
            items.push(doc.selection[i]);
        }

        for (var j = 0; j < items.length; j++) {
            var item = items[j];

            // 塗りと線の両方があるか判定
            var hasFill = hasVisibleFill(item);
            var hasStroke = hasVisibleStroke(item);
            var needSplit = splitFillStrokeRb.value && hasFill && hasStroke;

            if (needSplit) {
                var dup = item.duplicate(item, ElementPlacement.PLACEBEFORE);
                try { item.stroked = false; } catch (_) { }
                try { dup.filled = false; } catch (_) { }
            }

            // 角丸
            if (roundRadiusPt > 0) {
                applyRoundCorners(item, roundRadiusPt);
                if (needSplit) {
                    applyRoundCorners(dup, roundRadiusPt);
                }
            }

            var sx = rand(scaleMin, scaleMax);
            var sy = sx;
            var angle = rand(rotateMin, rotateMax);
            var dx = rand(-moveRange, moveRange);
            var dy = rand(-moveRange, moveRange);

            try {
                item.resize(sx, sy, true, true, true, true, (sx + sy) / 2);
                item.rotate(angle, true, true, true, true, Transformation.CENTER);
                item.translate(dx, dy);
            } catch (_) { }

            if (needSplit) {
                sx = rand(scaleMin, scaleMax);
                sy = sx;
                angle = rand(rotateMin, rotateMax);
                dx = rand(-moveRange, moveRange);
                dy = rand(-moveRange, moveRange);

                try {
                    dup.resize(sx, sy, true, true, true, true, (sx + sy) / 2);
                    dup.rotate(angle, true, true, true, true, Transformation.CENTER);
                    dup.translate(dx, dy);
                } catch (_) { }
            }

            // ラフ効果B
            if (useDistort && !isNaN(distortSize) && !isNaN(distortDetail)) {
                applyRoughen(item, distortSize, distortDetail, roughRoundness);
                if (needSplit) {
                    applyRoughen(dup, distortSize, distortDetail, roughRoundness);
                }
            }

            // ラフ効果A
            if (useRough && !isNaN(roughSize) && !isNaN(roughDetail)) {
                applyRoughen(item, roughSize, roughDetail, roughRoundness);
                if (needSplit) {
                    applyRoughen(dup, roughSize, roughDetail, roughRoundness);
                }
            }

            // グループ化
            if (needSplit && doGroup && !isInsideClippingGroup(item)) {
                var grp = createGroupForItem(item);
                if (grp) {
                    dup.move(grp, ElementPlacement.PLACEATEND);
                    item.move(grp, ElementPlacement.PLACEATEND);
                }
            }
        }

        app.redraw();
        return true;
    }

    /* プレビュー適用 / Apply preview */
    function applyPreview() {
        updateGroupCheckboxState();
        if (isPreviewApplied) {
            app.undo();
            isPreviewApplied = false;
        }
        isPreviewApplied = applyRandomizedEffects();
    }

    /* プレビュー解除 / Remove preview */
    function removePreview() {
        if (isPreviewApplied) {
            app.undo();
            app.redraw();
            isPreviewApplied = false;
        }
    }

    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isAlt = keyboard.altKey;
            var isShift = keyboard.shiftKey;
            var integerOnly = (editText === roughDetailInput || editText === distortDetailInput);

            if (event.keyName !== "Up" && event.keyName !== "Down") {
                return;
            }

            if (isShift) {
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / 10) * 10;
                } else {
                    value = Math.floor((value - 1) / 10) * 10;
                    if (value < 0) value = 0;
                }
            } else if (isAlt && !integerOnly) {
                if (event.keyName == "Up") {
                    value += 0.1;
                } else {
                    value -= 0.1;
                    if (value < 0) value = 0;
                }
                value = Math.round(value * 10) / 10;
            } else {
                if (event.keyName == "Up") {
                    value += 1;
                } else {
                    value -= 1;
                    if (value < 0) value = 0;
                }
                value = Math.round(value);
            }

            if (integerOnly) {
                value = Math.round(value);
            }

            event.preventDefault();
            editText.text = String(value);
            refreshPreview();
        });
    }

    function normalizeIntegerDetailInput(editText) {
        var value = Number(editText.text);
        if (isNaN(value)) {
            return false;
        }

        value = Math.round(value);
        if (value < 0) {
            value = 0;
        }

        var normalized = String(value);
        var changed = (editText.text !== normalized);
        editText.text = normalized;
        return changed;
    }

    function getCanSplit() {
        if (!doc || !doc.selection || doc.selection.length === 0) {
            return false;
        }

        for (var k = 0; k < doc.selection.length; k++) {
            var sel = doc.selection[k];
            if (hasVisibleFill(sel) && hasVisibleStroke(sel)) {
                return true;
            }
        }
        return false;
    }

    function updateGroupCheckboxState() {
        var canSplitNow = getCanSplit();
        var panel = splitFillStrokeRb ? splitFillStrokeRb.parent : null;

        if (panel) {
            panel.enabled = canSplitNow;
        }

        splitFillStrokeRb.enabled = canSplitNow;
        keepFillStrokeRb.enabled = true;

        if (!canSplitNow) {
            splitFillStrokeRb.value = false;
            keepFillStrokeRb.value = true;
        } else if (!splitFillStrokeRb.value && !keepFillStrokeRb.value) {
            splitFillStrokeRb.value = true;
        }

        groupSplitItemsCb.enabled = canSplitNow && splitFillStrokeRb.value;
        if (!groupSplitItemsCb.enabled) {
            groupSplitItemsCb.value = false;
        }
    }

    function refreshPreview() {
        updateGroupCheckboxState();
        if (!hasValidPreviewInputs()) {
            return;
        }
        removePreview();
        applyPreview();
    }

    /* ダイアログで値を設定 / Build dialog UI */
    function createDialog() {
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        /* 上部エリア（左右カラム） / Top area (left and right columns) */
        var topArea = dlg.add("group");
        topArea.orientation = "row";
        topArea.alignChildren = ["fill", "top"];

        /* 左カラム / Left column */
        var leftCol = topArea.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        var labelWidth = 55;

        /* 塗り/線パネル / Fill and stroke panel */
        var fillStrokePanel = leftCol.add("panel", undefined, L("panelFillStroke"));
        fillStrokePanel.enabled = getCanSplit();
        fillStrokePanel.orientation = "column";
        fillStrokePanel.alignChildren = ["left", "center"];
        fillStrokePanel.margins = [15, 20, 15, 10];

        splitFillStrokeRb = fillStrokePanel.add("radiobutton", undefined, L("splitFillStroke"));
        keepFillStrokeRb = fillStrokePanel.add("radiobutton", undefined, L("keepFillStroke"));
        if (getCanSplit()) {
            splitFillStrokeRb.value = true;
        } else {
            keepFillStrokeRb.value = true;
        }

        groupSplitItemsCb = fillStrokePanel.add("checkbox", undefined, L("groupSplitItems"));
        groupSplitItemsCb.value = getCanSplit();
        updateGroupCheckboxState();

        /* 変形パネル / Transform panel */
        var transformPanel = leftCol.add("panel", undefined, L("panelTransform"));
        transformPanel.orientation = "column";
        transformPanel.alignChildren = ["left", "center"];
        transformPanel.margins = [15, 20, 15, 10];

        var scaleGroup = transformPanel.add("group");
        scaleCb = scaleGroup.add("checkbox", undefined, labelText("scale"));
        scaleCb.preferredSize.width = 85;
        scaleCb.value = true;
        scaleInput = scaleGroup.add("edittext", undefined, "3");
        scaleInput.characters = 4;
        scaleGroup.add("statictext", undefined, "%");

        var moveGroup = transformPanel.add("group");
        moveCb = moveGroup.add("checkbox", undefined, labelText("move"));
        moveCb.preferredSize.width = 85;
        moveCb.value = false;
        moveInput = moveGroup.add("edittext", undefined, "0");
        moveInput.characters = 4;
        moveInput.enabled = false;
        moveGroup.add("statictext", undefined, getCurrentRulerUnitLabel());

        var rotateGroup = transformPanel.add("group");
        rotateCb = rotateGroup.add("checkbox", undefined, labelText("rotate"));
        rotateCb.preferredSize.width = 85;
        rotateCb.value = true;
        rotateInput = rotateGroup.add("edittext", undefined, "1.5");
        rotateInput.characters = 4;
        rotateGroup.add("statictext", undefined, "°");

        scaleCb.onClick = function () { scaleInput.enabled = scaleCb.value; refreshPreview(); };
        moveCb.onClick = function () { moveInput.enabled = moveCb.value; refreshPreview(); };
        rotateCb.onClick = function () { rotateInput.enabled = rotateCb.value; refreshPreview(); };

        scaleInput.onChanging = refreshPreview;
        moveInput.onChanging = refreshPreview;
        rotateInput.onChanging = refreshPreview;

        /* 右カラム / Right column */
        var rightCol = topArea.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["fill", "top"];

        /* 角丸パネル / Round corners panel */
        var roundPanel = rightCol.add("panel", undefined, L("panelRoundCorners"));
        roundPanel.orientation = "column";
        roundPanel.alignChildren = ["left", "center"];
        roundPanel.margins = [15, 20, 15, 10];

        var roundRadiusGroup = roundPanel.add("group");
        roundRadiusCb = roundRadiusGroup.add("checkbox", undefined, labelText("radius"));
        roundRadiusCb.preferredSize.width = labelWidth + 2;
        roundRadiusCb.value = true;
        roundRadiusInput = roundRadiusGroup.add("edittext", undefined, "1");
        roundRadiusInput.characters = 4;
        roundRadiusGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
        roundRadiusInput.enabled = roundRadiusCb.value;

        /* ラフ効果Aパネル / Roughen A panel */
        var roughPanel = rightCol.add("panel", undefined, L("panelRoughA"));
        roughPanel.orientation = "column";
        roughPanel.alignChildren = ["left", "center"];
        roughPanel.margins = [15, 20, 15, 10];

        roughEnableCb = roughPanel.add("checkbox", undefined, L("jaggedLine"));
        roughEnableCb.value = true;

        var roughSizeGroup = roughPanel.add("group");
        var roughSizeLabel = roughSizeGroup.add("statictext", undefined, labelText("size"), { justify: "right" });
        roughSizeLabel.preferredSize.width = labelWidth;
        roughSizeInput = roughSizeGroup.add("edittext", undefined, "0.3");
        roughSizeInput.characters = 4;
        roughSizeGroup.add("statictext", undefined, "%");

        var roughDetailGroup = roughPanel.add("group");
        var roughDetailLabel = roughDetailGroup.add("statictext", undefined, labelText("detail"), { justify: "right" });
        roughDetailLabel.preferredSize.width = labelWidth;
        roughDetailInput = roughDetailGroup.add("edittext", undefined, "20");
        roughDetailInput.characters = 4;
        roughDetailGroup.add("statictext", undefined, "/inch");

        /* ラフ効果Bパネル / Roughen B panel */
        var distortPanel = rightCol.add("panel", undefined, L("panelRoughB"));
        distortPanel.orientation = "column";
        distortPanel.alignChildren = ["left", "center"];
        distortPanel.margins = [15, 20, 15, 10];

        distortEnableCb = distortPanel.add("checkbox", undefined, L("applyDistort"));
        distortEnableCb.value = false;

        var distortSizeGroup = distortPanel.add("group");
        var distortSizeLabel = distortSizeGroup.add("statictext", undefined, labelText("size"), { justify: "right" });
        distortSizeLabel.preferredSize.width = labelWidth;
        distortSizeInput = distortSizeGroup.add("edittext", undefined, "2");
        distortSizeInput.characters = 4;
        distortSizeGroup.add("statictext", undefined, "%");

        var distortDetailGroup = distortPanel.add("group");
        var distortDetailLabel = distortDetailGroup.add("statictext", undefined, labelText("detail"), { justify: "right" });
        distortDetailLabel.preferredSize.width = labelWidth;
        distortDetailInput = distortDetailGroup.add("edittext", undefined, "0");
        distortDetailInput.characters = 4;
        distortDetailGroup.add("statictext", undefined, "/inch");

        /* チェックボックス・値変更時にプレビュー更新 / Refresh preview on checkbox or value change */
        roundRadiusCb.onClick = function () {
            roundRadiusInput.enabled = roundRadiusCb.value;
            refreshPreview();
        };
        roundRadiusInput.onChanging = refreshPreview;
        roughEnableCb.onClick = refreshPreview;
        roughSizeInput.onChanging = refreshPreview;
        roughDetailInput.onChanging = refreshPreview;
        roughDetailInput.onChange = function () {
            normalizeIntegerDetailInput(roughDetailInput);
            refreshPreview();
        };
        distortEnableCb.onClick = refreshPreview;
        distortSizeInput.onChanging = refreshPreview;
        distortDetailInput.onChanging = refreshPreview;
        distortDetailInput.onChange = function () {
            normalizeIntegerDetailInput(distortDetailInput);
            refreshPreview();
        };
        splitFillStrokeRb.onClick = function () { updateGroupCheckboxState(); refreshPreview(); };
        keepFillStrokeRb.onClick = function () { updateGroupCheckboxState(); refreshPreview(); };
        groupSplitItemsCb.onClick = refreshPreview;
        zoomCtrl = __TMKZoom_addControls(dlg, doc, L("zoom"), __zoomState, {
            min: 0.1,
            max: 8,
            sliderWidth: 240,
            margins: [0, 0, 0, 10],
            redraw: true,
            lightMode: true,
            lightModeLabel: L("lightMode"),
            lightModeDefault: false
        });

        /* ボタンエリア / Button area */
        var btnArea = dlg.add("group");
        btnArea.orientation = "row";
        btnArea.alignment = ["fill", "top"];

        var btnLeft = btnArea.add("group");
        btnLeft.alignment = ["left", "center"];
        previewBtn = btnLeft.add("button", undefined, L("preview"));
        previewBtn.onClick = refreshPreview;

        var btnSpacer = btnArea.add("group");
        btnSpacer.alignment = ["fill", "center"];

        var btnRight = btnArea.add("group");
        btnRight.alignment = ["right", "center"];
        cancelBtn = btnRight.add("button", undefined, L("cancel"), { name: "cancel" });
        okBtn = btnRight.add("button", undefined, L("ok"), { name: "ok" });

        changeValueByArrowKey(roundRadiusInput);
        changeValueByArrowKey(scaleInput);
        changeValueByArrowKey(moveInput);
        changeValueByArrowKey(rotateInput);
        changeValueByArrowKey(roughSizeInput);
        changeValueByArrowKey(roughDetailInput);
        changeValueByArrowKey(distortSizeInput);
        changeValueByArrowKey(distortDetailInput);

        dlg.addEventListener("show", function () {
            updateGroupCheckboxState();
            if (!hasValidPreviewInputs()) {
                return;
            }
            applyPreview();
        });

        return dlg;
    }

    dlg = createDialog();

    var result = dlg.show();

    if (result === 1) {
        /* OK: プレビュー済みならそのまま確定、未プレビューなら実行 / OK: keep preview if already applied, otherwise execute */
        if (!isPreviewApplied) {
            updateGroupCheckboxState();
            var executed = applyRandomizedEffects();
            if (!executed) {
                alert(L("alertInvalidNumber"));
            }
        }
    } else {
        /* キャンセル: プレビューを取り消し / Cancel: remove preview */
        if (zoomCtrl) {
            zoomCtrl.restoreInitial();
        }
        removePreview();
    }
})();