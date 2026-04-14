#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
選択したオブジェクトにランダム変形を適用する Illustrator スクリプトです。
塗りと線の分離、角丸、オフセット、ラフ効果（ギザギザ／歪曲）、グループ化、プレビュー、画面ズームに対応します。

・角丸・オフセット・移動は rulerType の単位で入力します。
・変形はライブ効果（Transform）として適用します。
・オフセットは［パスのオフセット］を負方向→正方向の順で2回、ラウンド結合で適用します。
・適用順は「角丸 → オフセット → 変形 → ラフ効果（歪曲 → ギザギザ）」です。

対象の扱い：
・グループは「個別適用」ON時に中の各オブジェクトへ展開します。
・選択オブジェクトが1つだけのとき、「グループ：個別適用」はディム表示になります。
・塗り＋線を持つオブジェクトのみ分離対象です。
・塗りのみ／線のみのオブジェクトは分離しません。
・分離の初期設定は「分離しない」です。

UI仕様：
・角丸パネルは「角を丸くする」「オフセット」のチェックボックスで各機能をON/OFFできます。
・ラフ効果（ギザギザ／歪曲）はチェックOFF時に入力欄をディム表示します。
・スケール／移動／回転／ギザギザ／歪曲がすべてOFFのときは「再計算」ボタンを無効化します。
・ズームは軽量モード付き（デフォルトOFF）です。

This Illustrator script applies randomized transformations to selected objects.
It supports fill/stroke splitting, round corners, offset path, roughen effects (jagged / distortion), grouping, preview, and zoom.

• Round corners, offset, and move use the current rulerType unit.
• Transform is applied as a live effect.
• Offset Path is applied twice (negative then positive) with round joins.
• Effect order: Round Corners → Offset Path → Transform → Roughen (Distortion → Jagged).

Target behavior:
• Groups are expanded when "Apply Individually" is enabled.
• When only one object is selected, "Apply Individually" is dimmed.
• Only objects with both fill and stroke can be split.
• Fill-only or stroke-only objects are not split.
• The default split setting is "Do Not Split."

UI behavior:
• The Round Corners panel uses separate checkboxes for Round Corners and Offset.
• Roughen controls are dimmed when disabled.
• Recalculate is disabled when Scale, Move, Rotate, Jagged, and Distortion are all OFF.
• Zoom includes a light mode (default OFF).

Updated: 2026-04-14
Version policy: Do not increment the version unless explicitly instructed.
*/

(function () {
    try {
        try {
            app.executeMenuCommand('edge');
        } catch (_) { }

        var SCRIPT_VERSION = "v1.1.0";

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
                ja: "対象",
                en: "Target"
            },
            splitFillStroke: {
                ja: "塗り/線を分離",
                en: "Split Fill / Stroke"
            },
            keepFillStroke: {
                ja: "分離しない",
                en: "Do Not Split"
            },
            applyToEachObject: {
                ja: "グループ：個別適用",
                en: "Apply Group Items Individually"
            },
            groupAtEnd: {
                ja: "最後にグループ化",
                en: "Group After Split"
            },
            panelTransform: {
                ja: "変形",
                en: "Transform"
            },
            toggleAllTransform: {
                ja: "すべてON/OFF",
                en: "Toggle All"
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
                ja: "角を丸くする",
                en: "Round Corners"
            },
            offset: {
                ja: "オフセット　",
                en: "Offset"
            },
            panelRoughA: {
                ja: "ラフ効果：ギザギザ",
                en: "Roughen: Jagged"
            },
            jaggedLine: {
                ja: "適用",
                en: "Apply"
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
                ja: "ラフ効果：歪曲",
                en: "Roughen: Distortion"
            },
            applyDistort: {
                ja: "適用",
                en: "Apply"
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
        var applyToEachObjectCb, splitFillStrokeRb, keepFillStrokeRb, groupAtEndCb;
        var scaleCb, moveCb, rotateCb;
        var scaleInput, moveInput, rotateInput;
        var roundRadiusCb, roundRadiusInput, roundOffsetCb, roundOffsetInput;
        var roughEnableCb, roughSizeLabel, roughSizeInput, roughDetailLabel, roughDetailInput;
        var distortEnableCb, distortSizeLabel, distortSizeInput, distortDetailLabel, distortDetailInput;
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

        function createOffsetPathEffectXML(offsetPt, joinType, miterLimit) {
            var xml = '<LiveEffect name="Adobe Offset Path"><Dict data="R ofst #1 I jntp #2 R mlim #3 "/></LiveEffect>';
            return xml
                .replace(/#1/, offsetPt)
                .replace(/#2/, joinType)
                .replace(/#3/, miterLimit);
        }


        function createTransformEffectXML(options) {
            var defaults = {
                scaleHorzPercent: 100,
                scaleVertPercent: 100,
                moveHorzPts: 0,
                moveVertPts: 0,
                rotateDegrees: 0,
                randomize: false,
                numberOfCopies: 0,
                transformPointIndex: 4,
                scaleStrokes: true,
                transformPatterns: true,
                transformObjects: true,
                reflectX: false,
                reflectY: false
            };

            var o = {};
            for (var key in defaults) {
                if (defaults.hasOwnProperty(key)) {
                    o[key] = defaults[key];
                }
            }
            for (var key2 in options) {
                if (options.hasOwnProperty(key2)) {
                    o[key2] = options[key2];
                }
            }

            return '<LiveEffect name="Adobe Transform"><Dict data="R scaleH_Percent #1 R scaleV_Percent #2 R scaleH_Factor #3 R scaleV_Factor #4 R moveH_Pts #5 R moveV_Pts #6 R rotate_Degrees #7 R rotate_Radians #8 I numCopies #9 I pinPoint #10 B scaleLines #11 B transformPatterns #12 B transformObjects #13 B reflectX #14 B reflectY #15 B randomize #16 "/></LiveEffect>'
                .replace(/#1/, o.scaleHorzPercent)
                .replace(/#2/, o.scaleVertPercent)
                .replace(/#3/, o.scaleHorzPercent / 100)
                .replace(/#4/, o.scaleVertPercent / 100)
                .replace(/#5/, o.moveHorzPts)
                .replace(/#6/, -o.moveVertPts)
                .replace(/#7/, o.rotateDegrees)
                .replace(/#8/, o.rotateDegrees * Math.PI / 180)
                .replace(/#9/, o.numberOfCopies)
                .replace(/#10/, o.transformPointIndex)
                .replace(/#11/, o.scaleStrokes ? 1 : 0)
                .replace(/#12/, o.transformPatterns ? 1 : 0)
                .replace(/#13/, o.transformObjects ? 1 : 0)
                .replace(/#14/, o.reflectX ? 1 : 0)
                .replace(/#15/, o.reflectY ? 1 : 0)
                .replace(/#16/, o.randomize ? 1 : 0);
        }

        function applyLiveTransform(obj, options) {
            try {
                obj.applyEffect(createTransformEffectXML(options));
            } catch (_) { }
        }


        function applyRoundCorners(obj, radiusPt) {
            try {
                obj.applyEffect(createRoundCornersEffectXML(radiusPt));
            } catch (_) { }
        }

        function applyOffsetPath(obj, offsetPt, joinType, miterLimit) {
            try {
                obj.applyEffect(createOffsetPathEffectXML(offsetPt, joinType, miterLimit));
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

        function isClippingPathItem(item) {
            try {
                return !!(item && item.typename === "PathItem" && item.clipping);
            } catch (_) {
                return false;
            }
        }

        /*
        対象の解釈 / Target interpretation
        A: 選択オブジェクトがグループの場合は、グループ内の各オブジェクトに適用します。
        B: 塗りと線の両方を持つオブジェクトは、分離 / 分離しないを選べます。
           分離した場合のみ、最後にグループ化するかどうかを選べます。
        C: 塗りのみ、または線のみのオブジェクトは分離しません。

        A: If the selected object is a group, apply to each object inside the group.
        B: If an object has both fill and stroke, split / do not split can be selected.
           Only when split is used, grouping at the end can be applied.
        C: Objects with fill only or stroke only are not split.
        */

        function shouldApplyToEachObject() {
            return applyToEachObjectCb ? applyToEachObjectCb.value : true;
        }

        function collectProcessItems(item, result) {
            if (!item) {
                return;
            }

            try {
                if (item.typename === "GroupItem") {
                    for (var i = 0; i < item.pageItems.length; i++) {
                        var child = item.pageItems[i];
                        if (child.parent !== item) {
                            continue;
                        }
                        collectProcessItems(child, result);
                    }
                    return;
                }
            } catch (_) { }

            if (isClippingPathItem(item)) {
                return;
            }

            result.push(item);
        }

        function getSelectedProcessItems() {
            var items = [];
            var applyEach = shouldApplyToEachObject();

            if (!doc || !doc.selection || doc.selection.length === 0) {
                return items;
            }

            for (var i = 0; i < doc.selection.length; i++) {
                var sel = doc.selection[i];
                if (!applyEach) {
                    if (!isClippingPathItem(sel)) {
                        items.push(sel);
                    }
                } else {
                    collectProcessItems(sel, items);
                }
            }

            return items;
        }

        function hasMultipleSelectedObjects() {
            try {
                return !!(doc && doc.selection && doc.selection.length > 1);
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
            var roundOffsetValue = roundOffsetCb.value ? parseFloat(roundOffsetInput.text) : 0;
            var roughSize = roughEnableCb.value ? parseFloat(roughSizeInput.text) : 0;
            var roughDetail = roughEnableCb.value ? parseFloat(roughDetailInput.text) : 0;
            var distortSize = distortEnableCb.value ? parseFloat(distortSizeInput.text) : 0;
            var distortDetail = distortEnableCb.value ? parseFloat(distortDetailInput.text) : 0;

            return !isNaN(scaleRange)
                && !isNaN(moveValue)
                && !isNaN(rotateRange)
                && !isNaN(roundRadiusValue)
                && !isNaN(roundOffsetValue)
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
            var doGroupAtEnd = groupAtEndCb.value;
            var roundRadiusValue = roundRadiusCb.value ? parseFloat(roundRadiusInput.text) : 0;
            var roundOffsetValue = roundOffsetCb.value ? parseFloat(roundOffsetInput.text) : 0;
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
            var roundOffsetPt = convertToPt(roundOffsetValue, rulerUnitCode);

            if (isNaN(scaleRange) || isNaN(moveValue) || isNaN(rotateRange) || isNaN(roundRadiusValue) || isNaN(roundOffsetValue)) {
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

            var items = getSelectedProcessItems();

            for (var j = 0; j < items.length; j++) {
                var item = items[j];

                /* 塗りと線の両方があるか判定 / Check whether the object has both visible fill and visible stroke */
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

                // パスのオフセット
                if (roundOffsetCb.value && roundOffsetPt !== 0) {
                    applyOffsetPath(item, -roundOffsetPt, 0, Math.abs(roundOffsetPt));
                    applyOffsetPath(item, roundOffsetPt, 0, Math.abs(roundOffsetPt));
                    if (needSplit) {
                        applyOffsetPath(dup, -roundOffsetPt, 0, Math.abs(roundOffsetPt));
                        applyOffsetPath(dup, roundOffsetPt, 0, Math.abs(roundOffsetPt));
                    }
                }

                var sx = rand(scaleMin, scaleMax);
                var sy = sx;
                var angle = rand(rotateMin, rotateMax);
                var dx = rand(-moveRange, moveRange);
                var dy = rand(-moveRange, moveRange);

                applyLiveTransform(item, {
                    scaleHorzPercent: sx,
                    scaleVertPercent: sy,
                    moveHorzPts: dx,
                    moveVertPts: dy,
                    rotateDegrees: angle,
                    transformPointIndex: 4,
                    scaleStrokes: true,
                    transformPatterns: true,
                    transformObjects: true,
                    randomize: false,
                    numberOfCopies: 0,
                    reflectX: false,
                    reflectY: false
                });

                if (needSplit) {
                    sx = rand(scaleMin, scaleMax);
                    sy = sx;
                    angle = rand(rotateMin, rotateMax);
                    dx = rand(-moveRange, moveRange);
                    dy = rand(-moveRange, moveRange);

                    applyLiveTransform(dup, {
                        scaleHorzPercent: sx,
                        scaleVertPercent: sy,
                        moveHorzPts: dx,
                        moveVertPts: dy,
                        rotateDegrees: angle,
                        transformPointIndex: 4,
                        scaleStrokes: true,
                        transformPatterns: true,
                        transformObjects: true,
                        randomize: false,
                        numberOfCopies: 0,
                        reflectX: false,
                        reflectY: false
                    });
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

                /* 最後にグループ化 / Group at end */
                if (needSplit && doGroupAtEnd && !isInsideClippingGroup(item)) {
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
            var items = getSelectedProcessItems();

            for (var k = 0; k < items.length; k++) {
                var item = items[k];
                if (hasVisibleFill(item) && hasVisibleStroke(item)) {
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

            if (applyToEachObjectCb) {
                applyToEachObjectCb.enabled = hasMultipleSelectedObjects();
            }

            splitFillStrokeRb.enabled = canSplitNow;
            keepFillStrokeRb.enabled = true;

            if (!splitFillStrokeRb.value && !keepFillStrokeRb.value) {
                splitFillStrokeRb.value = true;
            }

            groupAtEndCb.enabled = canSplitNow && splitFillStrokeRb.value;
            if (!groupAtEndCb.enabled) {
                groupAtEndCb.value = false;
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

            applyToEachObjectCb = fillStrokePanel.add("checkbox", undefined, L("applyToEachObject"));
            applyToEachObjectCb.value = true;

            splitFillStrokeRb = fillStrokePanel.add("radiobutton", undefined, L("splitFillStroke"));
            keepFillStrokeRb = fillStrokePanel.add("radiobutton", undefined, L("keepFillStroke"));
            splitFillStrokeRb.value = false;
            keepFillStrokeRb.value = true;

            groupAtEndCb = fillStrokePanel.add("checkbox", undefined, L("groupAtEnd"));
            groupAtEndCb.value = getCanSplit();

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

            var transformToggleGroup = transformPanel.add("group");
            transformToggleGroup.alignment = ["left", "center"];
            var toggleAllTransformBtn = transformToggleGroup.add("button", undefined, L("toggleAllTransform"));

            function updateRecalculateButtonState() {
                if (!previewBtn) {
                    return;
                }
                previewBtn.enabled = !!(scaleCb.value || moveCb.value || rotateCb.value || roughEnableCb.value || distortEnableCb.value);
            }

            function updateTransformInputStates() {
                scaleInput.enabled = scaleCb.value;
                moveInput.enabled = moveCb.value;
                rotateInput.enabled = rotateCb.value;
                updateRecalculateButtonState();
            }

            toggleAllTransformBtn.onClick = function () {
                var nextValue = !(scaleCb.value && moveCb.value && rotateCb.value);
                scaleCb.value = nextValue;
                moveCb.value = nextValue;
                rotateCb.value = nextValue;
                updateTransformInputStates();
                refreshPreview();
            };

            scaleCb.onClick = function () { updateTransformInputStates(); refreshPreview(); };
            moveCb.onClick = function () { updateTransformInputStates(); refreshPreview(); };
            rotateCb.onClick = function () { updateTransformInputStates(); refreshPreview(); };

            updateTransformInputStates();

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
            roundRadiusCb = roundRadiusGroup.add("checkbox", undefined, L("radius"));
            roundRadiusCb.value = true;
            roundRadiusInput = roundRadiusGroup.add("edittext", undefined, "1");
            roundRadiusInput.characters = 4;
            roundRadiusGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
            roundRadiusInput.enabled = roundRadiusCb.value;

            var roundOffsetGroup = roundPanel.add("group");
            roundOffsetCb = roundOffsetGroup.add("checkbox", undefined, L("offset"));
            roundOffsetCb.value = false;
            roundOffsetInput = roundOffsetGroup.add("edittext", undefined, "1");
            roundOffsetInput.characters = 4;
            roundOffsetGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
            roundOffsetInput.enabled = roundOffsetCb.value;

            /* ラフ効果Aパネル / Roughen A panel */
            var roughPanel = rightCol.add("panel", undefined, L("panelRoughA"));
            roughPanel.orientation = "column";
            roughPanel.alignChildren = ["left", "center"];
            roughPanel.margins = [15, 20, 15, 10];

            roughEnableCb = roughPanel.add("checkbox", undefined, L("jaggedLine"));
            roughEnableCb.value = true;

            var roughSizeGroup = roughPanel.add("group");
            roughSizeLabel = roughSizeGroup.add("statictext", undefined, labelText("size"), { justify: "right" });
            roughSizeLabel.preferredSize.width = labelWidth;
            roughSizeInput = roughSizeGroup.add("edittext", undefined, "0.3");
            roughSizeInput.characters = 4;
            roughSizeGroup.add("statictext", undefined, "%");

            var roughDetailGroup = roughPanel.add("group");
            roughDetailLabel = roughDetailGroup.add("statictext", undefined, labelText("detail"), { justify: "right" });
            roughDetailLabel.preferredSize.width = labelWidth;
            roughDetailInput = roughDetailGroup.add("edittext", undefined, "20");
            roughDetailInput.characters = 4;
            roughDetailGroup.add("statictext", undefined, "/inch");
            roughSizeLabel.enabled = roughEnableCb.value;
            roughSizeInput.enabled = roughEnableCb.value;
            roughDetailLabel.enabled = roughEnableCb.value;
            roughDetailInput.enabled = roughEnableCb.value;

            /* ラフ効果Bパネル / Roughen B panel */
            var distortPanel = rightCol.add("panel", undefined, L("panelRoughB"));
            distortPanel.orientation = "column";
            distortPanel.alignChildren = ["left", "center"];
            distortPanel.margins = [15, 20, 15, 10];

            distortEnableCb = distortPanel.add("checkbox", undefined, L("applyDistort"));
            distortEnableCb.value = false;

            var distortSizeGroup = distortPanel.add("group");
            distortSizeLabel = distortSizeGroup.add("statictext", undefined, labelText("size"), { justify: "right" });
            distortSizeLabel.preferredSize.width = labelWidth;
            distortSizeInput = distortSizeGroup.add("edittext", undefined, "2");
            distortSizeInput.characters = 4;
            distortSizeGroup.add("statictext", undefined, "%");

            var distortDetailGroup = distortPanel.add("group");
            distortDetailLabel = distortDetailGroup.add("statictext", undefined, labelText("detail"), { justify: "right" });
            distortDetailLabel.preferredSize.width = labelWidth;
            distortDetailInput = distortDetailGroup.add("edittext", undefined, "0");
            distortDetailInput.characters = 4;
            distortDetailGroup.add("statictext", undefined, "/inch");
            distortSizeLabel.enabled = distortEnableCb.value;
            distortSizeInput.enabled = distortEnableCb.value;
            distortDetailLabel.enabled = distortEnableCb.value;
            distortDetailInput.enabled = distortEnableCb.value;

            /* チェックボックス・値変更時にプレビュー更新 / Refresh preview on checkbox or value change */
            roundRadiusCb.onClick = function () {
                roundRadiusInput.enabled = roundRadiusCb.value;
                refreshPreview();
            };
            roundRadiusInput.onChanging = refreshPreview;

            roundOffsetCb.onClick = function () {
                roundOffsetInput.enabled = roundOffsetCb.value;
                refreshPreview();
            };
            roundOffsetInput.onChanging = refreshPreview;
            roughEnableCb.onClick = function () {
                roughSizeLabel.enabled = roughEnableCb.value;
                roughSizeInput.enabled = roughEnableCb.value;
                roughDetailLabel.enabled = roughEnableCb.value;
                roughDetailInput.enabled = roughEnableCb.value;
                updateRecalculateButtonState();
                refreshPreview();
            };
            roughSizeInput.onChanging = refreshPreview;
            roughDetailInput.onChanging = refreshPreview;
            roughDetailInput.onChange = function () {
                normalizeIntegerDetailInput(roughDetailInput);
                refreshPreview();
            };
            distortEnableCb.onClick = function () {
                distortSizeLabel.enabled = distortEnableCb.value;
                distortSizeInput.enabled = distortEnableCb.value;
                distortDetailLabel.enabled = distortEnableCb.value;
                distortDetailInput.enabled = distortEnableCb.value;
                updateRecalculateButtonState();
                refreshPreview();
            };
            distortSizeInput.onChanging = refreshPreview;
            distortDetailInput.onChanging = refreshPreview;
            distortDetailInput.onChange = function () {
                normalizeIntegerDetailInput(distortDetailInput);
                refreshPreview();
            };
            applyToEachObjectCb.onClick = refreshPreview;
            splitFillStrokeRb.onClick = function () { updateGroupCheckboxState(); refreshPreview(); };
            keepFillStrokeRb.onClick = function () { updateGroupCheckboxState(); refreshPreview(); };
            groupAtEndCb.onClick = refreshPreview;
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
            updateRecalculateButtonState();

            var btnSpacer = btnArea.add("group");
            btnSpacer.alignment = ["fill", "center"];

            var btnRight = btnArea.add("group");
            btnRight.alignment = ["right", "center"];
            cancelBtn = btnRight.add("button", undefined, L("cancel"), { name: "cancel" });
            okBtn = btnRight.add("button", undefined, L("ok"), { name: "ok" });

            changeValueByArrowKey(roundRadiusInput);
            changeValueByArrowKey(roundOffsetInput);
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
    }
    finally {
        try {
            app.executeMenuCommand('edge');
        } catch (_) { }
    }
})();