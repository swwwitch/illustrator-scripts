#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要:
選択オブジェクトの visibleBounds を基準に、横線・縦線・分割線を生成するスクリプトです。
横線では、横伸張率、分割、上下に追加、左方向に延長を設定できます。
縦線では、縦伸張率、分割、左右に追加、上方向に延長を設定できます。
共通設定では、作成レイヤー名、線幅、ガイド化、グループ化を指定できます。
ダイアログ上部のプリセットから、標準・基本・2x2・左にロゴマーク・上部にロゴマークをすばやく適用できます。
左方向に延長をONにした場合は、右側を固定したまま横線全体を左方向へ延長します。
罫線およびガイドは指定したレイヤーに作成され、レイヤーがない場合には自動作成し、ある場合にはそのまま利用します。
プレビュー表示に対応しており、必要に応じて生成した線をガイドに変換できます。
線幅の入力と表示は Illustrator の線単位設定（strokeUnits）を参照します。

更新日: 2026-04-06
*/

(function () {
    // =========================================
    // バージョンとローカライズ
    // =========================================

    var SCRIPT_VERSION = "v1.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "補助線の作成",
            en: "Create Guide Lines"
        },
        labelPreset: {
            ja: "プリセット:",
            en: "Preset:"
        },
        presetDefault: {
            ja: "標準",
            en: "Default"
        },
        presetBasic: {
            ja: "基本",
            en: "Basic"
        },
        preset2x2: {
            ja: "2x2",
            en: "2x2"
        },
        presetLogoLeft: {
            ja: "左にロゴマーク",
            en: "Logo Mark at Left"
        },
        presetLogoTop: {
            ja: "上部にロゴマーク",
            en: "Logo Mark at Top"
        },
        errNoDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        errNoSelection: {
            ja: "オブジェクトを選択してください。",
            en: "Please select an object."
        },
        errNoBounds: {
            ja: "選択範囲の境界を取得できませんでした。",
            en: "Could not get the bounds of the selection."
        },
        errInvalidSize: {
            ja: "選択オブジェクトのサイズを取得できませんでした。",
            en: "Could not get the size of the selected object."
        },
        errInvalidInput: {
            ja: "入力値が不正です。数値を正しく入力してください。",
            en: "Invalid input. Please enter valid numeric values."
        },
        defaultGuideLayerName: {
            ja: "// logo_guide",
            en: "// logo_guide"
        },
        labelGuideLayer: {
            ja: "作成レイヤー",
            en: "Target Layer"
        },
        panelHorizontal: {
            ja: "横線",
            en: "Horizontal"
        },
        panelVertical: {
            ja: "縦線",
            en: "Vertical"
        },
        panelCommon: {
            ja: "共通設定",
            en: "Common"
        },
        labelWidthScale: {
            ja: "横伸張率",
            en: "Horizontal Scale"
        },
        labelHeightScale: {
            ja: "縦伸張率",
            en: "Vertical Scale"
        },
        labelStrokeWidth: {
            ja: "線幅",
            en: "Stroke Width"
        },
        checkHorizontalDiv: {
            ja: "分割",
            en: "Divide"
        },
        checkHorizontalExtra: {
            ja: "上下に追加",
            en: "Add Top/Bottom Lines"
        },
        checkHorizontalLeftExtra: {
            ja: "左方向に延長",
            en: "Add Left Line"
        },
        checkVerticalOuter: {
            ja: "左右に追加",
            en: "Add Left/Right Lines"
        },
        checkVerticalDiv: {
            ja: "分割",
            en: "Divide"
        },
        checkVerticalTopSpace: {
            ja: "上方向に延長",
            en: "Extend Upward"
        },
        checkGuides: {
            ja: "ガイドに変換する",
            en: "Convert to Guides"
        },
        checkGroupItems: {
            ja: "グループ化",
            en: "Group Items"
        },
        checkPreview: {
            ja: "プレビュー",
            en: "Preview"
        },
        btnCancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        btnOk: {
            ja: "OK",
            en: "OK"
        }
    };

    function L(key) {
        var entry = LABELS[key];
        return entry ? entry[lang] : key;
    }

    /* Live Corner Annotator 実行 / Run Live Corner Annotator */
    function runLiveCornerAnnotator() {
        try {
            if (typeof app === "undefined" || !app || typeof app.executeMenuCommand !== "function") {
                return;
            }
            app.executeMenuCommand('Live Corner Annotator');
        } catch (e) { }
    }

    // =========================================
    // 単位ユーティリティ
    // =========================================

    /* 単位コードマップ / Unit code map */
    var unitMap = {
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

    /* 単位コードと設定キーからラベル取得 / Get unit label from code and preference key */
    function getUnitLabel(code, prefKey) {
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return unitMap[code] || "pt";
    }

    /* 単位コードをpt係数へ変換 / Convert unit code to pt factor */
    function getUnitFactor(code, prefKey) {
        switch (code) {
            case 0: return 72; /* in */
            case 1: return 72 / 25.4; /* mm */
            case 2: return 1; /* pt */
            case 3: return 12; /* pica */
            case 4: return 72 / 2.54; /* cm */
            case 5: return getUnitLabel(code, prefKey) === "H" ? (72 / 72) : (72 / 101.6); /* H or Q */
            case 6: return 1; /* px (Illustrator points base) */
            case 7: return 1; /* ft/in は複合入力未対応のため安全側でpt扱い / Treat as pt because compound ft/in input is unsupported */
            case 8: return 72 / 0.0254; /* m */
            case 9: return 72 * 36; /* yd */
            case 10: return 72 * 12; /* ft */
            default: return 1;
        }
    }

    /* 環境設定の単位コード取得 / Get preference unit code */
    function getPreferenceUnitCode(prefKey) {
        try {
            return app.preferences.getIntegerPreference(prefKey);
        } catch (e) {
            return 2;
        }
    }


    /* 現在の線単位ラベル取得 / Get current stroke unit label */
    function getCurrentStrokeUnitLabel() {
        var unitCode = getPreferenceUnitCode("strokeUnits");
        return getUnitLabel(unitCode, "strokeUnits");
    }

    /* 単位値をptへ変換 / Convert unit value to pt */
    function convertToPoints(value, prefKey) {
        var code = getPreferenceUnitCode(prefKey);
        return value * getUnitFactor(code, prefKey);
    }

    var doc;
    var sel;
    var previewGroup = null;
    main();

    // =========================================
    // 関数定義
    // =========================================

    function main() {
        try {
            runLiveCornerAnnotator();
        } catch (e) { }
        try {
            if (app.documents.length === 0) {
                alert(L("errNoDocument"));
                return;
            }

            doc = app.activeDocument;

            if (doc.selection.length === 0) {
                alert(L("errNoSelection"));
                return;
            }

            sel = doc.selection;

            previewGroup = null;
            var ui = createDialog();

            var context = buildContext(doc, sel, ui.guideLayerInput.text);
            if (!context) {
                return;
            }

            bindArrowKeyHandlers(ui, context);
            bindEvents(ui, context);
            updateVerticalTopSpaceEnabled(ui);
            updateGroupCheckEnabled(ui);

            ui.dialog.show();
        } finally {
            try {
                runLiveCornerAnnotator();
            } catch (e) { }
        }
    }

    function buildContext(doc, sel, layerName) {
        var bounds = getSelectionBounds(sel);
        if (!bounds) {
            alert(L("errNoBounds"));
            return null;
        }

        var bLeft = bounds[0];
        var bTop = bounds[1];
        var bRight = bounds[2];
        var bBottom = bounds[3];

        var A = bTop - bBottom;
        var B = bRight - bLeft;

        if (A <= 0 || B <= 0) {
            alert(L("errInvalidSize"));
            return null;
        }

        var centerX = (bLeft + bRight) / 2;
        var centerY = (bTop + bBottom) / 2;

        return {
            doc: doc,
            bLeft: bLeft,
            bTop: bTop,
            bRight: bRight,
            bBottom: bBottom,
            A: A,
            B: B,
            centerX: centerX,
            centerY: centerY,
            guideLayer: getOrCreateLayer(doc, layerName)
        };
    }
    function findLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                return doc.layers[i];
            }
        }
        return null;
    }

    function getOrCreateLayer(doc, name) {
        var layer = findLayerByName(doc, name);
        if (layer) {
            return layer;
        }
        layer = doc.layers.add();
        layer.name = name;
        return layer;
    }

    function setUiValueAndNotify(editText, value, ui, context, suppressPreview) {
        if (editText === ui.hScaleInput || editText === ui.vScaleInput) {
            editText.text = Number(value).toFixed(1);
        } else {
            editText.text = String(value);
        }
        if (editText === ui.vScaleInput) {
            updateVerticalTopSpaceEnabled(ui);
        }
        if (!suppressPreview && ui.previewCheck.value) {
            updatePreview(ui, context);
        }
    }

    function applyPreset(ui, context) {
        if (!ui.presetDropdown.selection) {
            return;
        }

        var index = ui.presetDropdown.selection.index;
        var suppressPreview = true;

        if (index === 0) {
            /* 標準 / Standard */
            setUiValueAndNotify(ui.hScaleInput, 1.4, ui, context, suppressPreview);
            ui.hDivCheck.value = true;
            setUiValueAndNotify(ui.divInput, 4, ui, context, suppressPreview);
            ui.hExtraCheck.value = true;
            ui.hLeftExtraCheck.value = false;

            setUiValueAndNotify(ui.vScaleInput, 2.0, ui, context, suppressPreview);
            ui.vDivCheck.value = false;
            setUiValueAndNotify(ui.vDivInput, 2, ui, context, suppressPreview);
            ui.vOuterCheck.value = true;
            ui.vTopSpaceCheck.value = false;
            updateVerticalTopSpaceEnabled(ui);
        } else if (index === 1) {
            /* 基本 / Basic */
            setUiValueAndNotify(ui.hScaleInput, 1.4, ui, context, suppressPreview);
            ui.hDivCheck.value = false;
            setUiValueAndNotify(ui.divInput, 4, ui, context, suppressPreview);
            ui.hExtraCheck.value = false;
            ui.hLeftExtraCheck.value = false;

            setUiValueAndNotify(ui.vScaleInput, 2.0, ui, context, suppressPreview);
            ui.vDivCheck.value = false;
            setUiValueAndNotify(ui.vDivInput, 2, ui, context, suppressPreview);
            ui.vOuterCheck.value = false;
            ui.vTopSpaceCheck.value = false;
            updateVerticalTopSpaceEnabled(ui);
        } else if (index === 2) {
            /* 2x2 */
            setUiValueAndNotify(ui.hScaleInput, 2.0, ui, context, suppressPreview);
            ui.hDivCheck.value = true;
            setUiValueAndNotify(ui.divInput, 2, ui, context, suppressPreview);
            ui.hExtraCheck.value = false;
            ui.hLeftExtraCheck.value = false;

            setUiValueAndNotify(ui.vScaleInput, 2.0, ui, context, suppressPreview);
            ui.vDivCheck.value = true;
            setUiValueAndNotify(ui.vDivInput, 2, ui, context, suppressPreview);
            ui.vOuterCheck.value = false;
            ui.vTopSpaceCheck.value = false;
            updateVerticalTopSpaceEnabled(ui);
        } else if (index === 3) {
            /* 左にロゴマーク / Logo Mark at Left */
            setUiValueAndNotify(ui.hScaleInput, 10.0, ui, context, suppressPreview);
            ui.hDivCheck.value = true;
            setUiValueAndNotify(ui.divInput, 4, ui, context, suppressPreview);
            ui.hExtraCheck.value = true;
            ui.hLeftExtraCheck.value = true;

            setUiValueAndNotify(ui.vScaleInput, 2.0, ui, context, suppressPreview);
            ui.vDivCheck.value = false;
            setUiValueAndNotify(ui.vDivInput, 2, ui, context, suppressPreview);
            ui.vOuterCheck.value = true;
            ui.vTopSpaceCheck.value = false;
            updateVerticalTopSpaceEnabled(ui);
        } else if (index === 4) {
            /* 上部にロゴマーク / Logo Mark at Top */
            setUiValueAndNotify(ui.hScaleInput, 1.4, ui, context, suppressPreview);
            ui.hDivCheck.value = true;
            setUiValueAndNotify(ui.divInput, 4, ui, context, suppressPreview);
            ui.hExtraCheck.value = true;
            ui.hLeftExtraCheck.value = false;

            setUiValueAndNotify(ui.vScaleInput, 3.5, ui, context, suppressPreview);
            ui.vDivCheck.value = true;
            setUiValueAndNotify(ui.vDivInput, 3, ui, context, suppressPreview);
            ui.vOuterCheck.value = true;
            ui.vTopSpaceCheck.value = true;
            updateVerticalTopSpaceEnabled(ui);
        }

        if (ui.previewCheck.value) {
            updatePreview(ui, context);
        }
    }

    function bindArrowKeyHandlers(ui, context) {
        changeValueByArrowKey(ui.hScaleInput, false, false, ui, context);
        changeValueByArrowKey(ui.divInput, false, true, ui, context);
        changeValueByArrowKey(ui.vScaleInput, false, false, ui, context);
        changeValueByArrowKey(ui.vDivInput, false, true, ui, context);
        changeValueByArrowKey(ui.swInput, false, false, ui, context);
    }
    function bindEvents(ui, context) {
        /* イベントハンドラ / Event handlers */
        ui.presetDropdown.onChange = function () {
            applyPreset(ui, context);
        };
        ui.previewCheck.onClick = function () {
            if (ui.previewCheck.value) {
                updatePreview(ui, context);
            } else {
                removePreview();
            }
        };

        ui.hScaleInput.onChanging = ui.divInput.onChanging = ui.swInput.onChanging =
            ui.vDivInput.onChanging = function () {
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

        ui.vScaleInput.onChanging = function () {
            updateVerticalTopSpaceEnabled(ui);
            if (ui.previewCheck.value) {
                updatePreview(ui, context);
            }
        };

        ui.hScaleInput.onBlur = function () {
            var value = parseFloat(ui.hScaleInput.text);
            if (!isNaN(value) && isFinite(value)) {
                ui.hScaleInput.text = value.toFixed(1);
            }
        };

        ui.vScaleInput.onBlur = function () {
            var value = parseFloat(ui.vScaleInput.text);
            if (!isNaN(value) && isFinite(value)) {
                ui.vScaleInput.text = value.toFixed(1);
                updateVerticalTopSpaceEnabled(ui);
            }
        };

        ui.hDivCheck.onClick = ui.hExtraCheck.onClick = ui.hLeftExtraCheck.onClick =
            ui.vOuterCheck.onClick = ui.vDivCheck.onClick =
            ui.vTopSpaceCheck.onClick = ui.groupCheck.onClick = function () {
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

        ui.guideCheck.onClick = function () {
            updateGroupCheckEnabled(ui);
            if (ui.previewCheck.value) {
                updatePreview(ui, context);
            }
        };

        ui.okBtn.onClick = function () {
            removePreview();
            var params = getParams(ui);
            if (!params) {
                alert(L("errInvalidInput"));
                return;
            }
            context.guideLayer = getOrCreateLayer(context.doc, params.guideLayerName);
            createLines(context, params, false);
            ui.dialog.close(1);
        };

        ui.cancelBtn.onClick = function () {
            removePreview();
            ui.dialog.close(2);
        };

        ui.dialog.onClose = function () {
            removePreview();
        };
    }

    function createDialog() {
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        var presetWrap = dlg.add("group");
        presetWrap.alignment = ["center", "top"];

        var presetRow = presetWrap.add("group");
        presetRow.add("statictext", undefined, L("labelPreset"));
        var presetDropdown = presetRow.add("dropdownlist", undefined, [
            L("presetDefault"),
            L("presetBasic"),
            L("preset2x2"),
            L("presetLogoLeft"),
            L("presetLogoTop")
        ]);
        presetDropdown.selection = 0;

        var columns = dlg.add("group");
        columns.orientation = "row";
        columns.alignChildren = ["fill", "top"];

        var leftColumn = columns.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];

        var rightColumn = columns.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];

        /* 横線設定 / Horizontal line settings */
        var hPanel = leftColumn.add("panel", undefined, L("panelHorizontal"));
        hPanel.alignChildren = ["left", "center"];
        hPanel.margins = [15, 20, 15, 10];

        var hRow1 = hPanel.add("group");
        hRow1.add("statictext", undefined, L("labelWidthScale"));
        var hScaleInput = hRow1.add("edittext", undefined, "1.4");
        hScaleInput.characters = 5;

        var hRow2 = hPanel.add("group");
        var hDivCheck = hRow2.add("checkbox", undefined, L("checkHorizontalDiv"));
        hDivCheck.value = true;
        var divInput = hRow2.add("edittext", undefined, "4");
        divInput.characters = 5;

        var hRow3 = hPanel.add("group");
        var hExtraCheck = hRow3.add("checkbox", undefined, L("checkHorizontalExtra"));
        hExtraCheck.value = true;

        var hRow4 = hPanel.add("group");
        var hLeftExtraCheck = hRow4.add("checkbox", undefined, L("checkHorizontalLeftExtra"));
        hLeftExtraCheck.value = false;

        /* 縦線設定 / Vertical line settings */
        var vPanel = rightColumn.add("panel", undefined, L("panelVertical"));
        vPanel.alignChildren = ["left", "center"];
        vPanel.margins = [15, 20, 15, 10];

        var vRow1 = vPanel.add("group");
        vRow1.add("statictext", undefined, L("labelHeightScale"));
        var vScaleInput = vRow1.add("edittext", undefined, "2.0");
        vScaleInput.characters = 5;

        var vRow2 = vPanel.add("group");
        var vDivCheck = vRow2.add("checkbox", undefined, L("checkVerticalDiv"));
        vDivCheck.value = false;
        var vDivInput = vRow2.add("edittext", undefined, "2");
        vDivInput.characters = 5;

        var vRow3 = vPanel.add("group");
        var vOuterCheck = vRow3.add("checkbox", undefined, L("checkVerticalOuter"));
        vOuterCheck.value = true;

        var vRow4 = vPanel.add("group");
        var vTopSpaceCheck = vRow4.add("checkbox", undefined, L("checkVerticalTopSpace"));
        vTopSpaceCheck.value = false;
        vTopSpaceCheck.enabled = false;

        /* 共通設定 / Common settings */
        var commonPanel = dlg.add("panel", undefined, L("panelCommon"));
        commonPanel.alignChildren = ["left", "top"];
        commonPanel.margins = [15, 20, 15, 10];

        var cRow1 = commonPanel.add("group");
        cRow1.add("statictext", undefined, L("labelGuideLayer"));
        var guideLayerInput = cRow1.add("edittext", undefined, L("defaultGuideLayerName"));
        guideLayerInput.characters = 16;

        var cRow2 = commonPanel.add("group");
        cRow2.add("statictext", undefined, L("labelStrokeWidth"));
        var swInput = cRow2.add("edittext", undefined, "0.25");
        swInput.characters = 5;
        cRow2.add("statictext", undefined, getCurrentStrokeUnitLabel());

        var commonChecksRow = commonPanel.add("group");
        commonChecksRow.orientation = "row";
        commonChecksRow.alignChildren = ["left", "center"];

        var guideCheck = commonChecksRow.add("checkbox", undefined, L("checkGuides"));
        guideCheck.value = false;

        var groupCheck = commonChecksRow.add("checkbox", undefined, L("checkGroupItems"));
        groupCheck.value = true;

        /* 下部ボタンエリア / Bottom button area */
        var bottomRow = dlg.add("group");
        bottomRow.orientation = "row";
        bottomRow.alignChildren = ["fill", "center"];

        var leftButtons = bottomRow.add("group");
        leftButtons.orientation = "row";
        leftButtons.alignChildren = ["left", "center"];
        var previewCheck = leftButtons.add("checkbox", undefined, L("checkPreview"));
        previewCheck.value = false;

        var spacer = bottomRow.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var rightButtons = bottomRow.add("group");
        rightButtons.orientation = "row";
        rightButtons.alignChildren = ["right", "center"];
        var cancelBtn = rightButtons.add("button", undefined, L("btnCancel"), { name: "cancel" });
        var okBtn = rightButtons.add("button", undefined, L("btnOk"), { name: "ok" });

        return {
            dialog: dlg,
            presetDropdown: presetDropdown,
            hScaleInput: hScaleInput,
            hDivCheck: hDivCheck,
            divInput: divInput,
            hExtraCheck: hExtraCheck,
            hLeftExtraCheck: hLeftExtraCheck,
            vScaleInput: vScaleInput,
            vOuterCheck: vOuterCheck,
            vDivCheck: vDivCheck,
            vDivInput: vDivInput,
            vTopSpaceCheck: vTopSpaceCheck,
            guideLayerInput: guideLayerInput,
            swInput: swInput,
            guideCheck: guideCheck,
            groupCheck: groupCheck,
            previewCheck: previewCheck,
            cancelBtn: cancelBtn,
            okBtn: okBtn
        };
    }

    function updateGroupCheckEnabled(ui) {
        ui.groupCheck.enabled = !ui.guideCheck.value;
        if (ui.guideCheck.value) {
            ui.groupCheck.value = false;
        }
    }

    function updateVerticalTopSpaceEnabled(ui) {
        ui.vTopSpaceCheck.enabled = true;
    }

    function isStrictPositiveNumberText(text) {
        return /^\d+(?:\.\d+)?$/.test(text);
    }

    function readParams(ui) {
        return {
            guideLayerNameText: ui.guideLayerInput.text,
            hScaleText: ui.hScaleInput.text,
            vScaleText: ui.vScaleInput.text,
            strokeWidthText: ui.swInput.text,
            hScale: parseFloat(ui.hScaleInput.text),
            vScale: parseFloat(ui.vScaleInput.text),
            strokeWidthValue: parseFloat(ui.swInput.text),
            divisionsText: ui.divInput.text,
            vDivisionsText: ui.vDivInput.text,
            divisions: parseInt(ui.divInput.text, 10),
            vDivisions: parseInt(ui.vDivInput.text, 10),
            hDiv: ui.hDivCheck.value,
            hExtra: ui.hExtraCheck.value,
            hLeftExtra: ui.hLeftExtraCheck.value,
            makeGuides: ui.guideCheck.value,
            groupItems: ui.groupCheck.value,
            vOuter: ui.vOuterCheck.value,
            vDiv: ui.vDivCheck.value,
            vTopSpace: ui.vTopSpaceCheck.value
        };
    }

    function validateParams(raw) {
        if (!raw.guideLayerNameText || /^\s*$/.test(raw.guideLayerNameText)) {
            return false;
        }
        if (!isStrictPositiveNumberText(raw.hScaleText) || isNaN(raw.hScale) || !isFinite(raw.hScale) || raw.hScale <= 0) {
            return false;
        }
        if (!isStrictPositiveNumberText(raw.vScaleText) || isNaN(raw.vScale) || !isFinite(raw.vScale) || raw.vScale <= 0) {
            return false;
        }
        if (!/^\d+$/.test(raw.divisionsText) || isNaN(raw.divisions) || !isFinite(raw.divisions) || raw.divisions < 1) {
            return false;
        }
        if (!isStrictPositiveNumberText(raw.strokeWidthText) || isNaN(raw.strokeWidthValue) || !isFinite(raw.strokeWidthValue) || raw.strokeWidthValue <= 0) {
            return false;
        }
        if (raw.vDiv) {
            if (!/^\d+$/.test(raw.vDivisionsText) || isNaN(raw.vDivisions) || !isFinite(raw.vDivisions) || raw.vDivisions < 1) {
                return false;
            }
        }
        return true;
    }

    function normalizeParams(raw) {
        var strokeWidth = convertToPoints(raw.strokeWidthValue, "strokeUnits");
        var vDivisions = raw.vDiv ? raw.vDivisions : 0;

        if (isNaN(strokeWidth) || !isFinite(strokeWidth) || strokeWidth <= 0) {
            return null;
        }

        return {
            guideLayerName: raw.guideLayerNameText,
            hScale: raw.hScale,
            vScale: raw.vScale,
            divisions: raw.divisions,
            strokeWidth: strokeWidth,
            hDiv: raw.hDiv,
            hExtra: raw.hExtra,
            hLeftExtra: raw.hLeftExtra,
            makeGuides: raw.makeGuides,
            groupItems: raw.groupItems,
            vOuter: raw.vOuter,
            vDiv: raw.vDiv,
            vDivisions: vDivisions,
            vTopSpace: raw.vTopSpace
        };
    }

    function getParams(ui) {
        var raw = readParams(ui);
        if (!validateParams(raw)) {
            return null;
        }
        return normalizeParams(raw);
    }

    function createLines(context, params, isPreview) {
        var C = context.A / params.divisions;

        var hLineLength = context.B * params.hScale;
        var hL = context.centerX - hLineLength / 2;
        var hR = context.centerX + hLineLength / 2;
        var hLeft = hL;
        var hRight = hR;

        if (params.hLeftExtra) {
            hRight = context.bRight + (context.A / 2);
            hLeft = hRight - hLineLength;
        }

        var vLineHeight = context.A * params.vScale;
        var vT, vB;
        if (params.vTopSpace && params.vScale >= 3) {
            vB = context.bBottom - (C * 2);
            vT = vB + vLineHeight;
        } else {
            vT = context.centerY + vLineHeight / 2;
            vB = context.centerY - vLineHeight / 2;
        }

        var xLeftOuter = context.bLeft - C;
        var xRightOuter = context.bRight + C;

        var strokeColor = new GrayColor();
        strokeColor.gray = 100;

        var group = context.guideLayer.groupItems.add();
        group.name = isPreview ? "GuideLines_Preview" : "GuideLines";

        /* 横線 / Horizontal lines */
        if (params.hExtra) {
            drawLine(group, hLeft, context.bTop + C, hRight, context.bTop + C, strokeColor, params.strokeWidth);
        }
        drawLine(group, hLeft, context.bTop, hRight, context.bTop, strokeColor, params.strokeWidth);
        if (params.hDiv && params.divisions > 1) {
            for (var i = 1; i < params.divisions; i++) {
                var y = context.bTop - C * i;
                drawLine(group, hLeft, y, hRight, y, strokeColor, params.strokeWidth);
            }
        }
        drawLine(group, hLeft, context.bBottom, hRight, context.bBottom, strokeColor, params.strokeWidth);
        if (params.hExtra) {
            drawLine(group, hLeft, context.bBottom - C, hRight, context.bBottom - C, strokeColor, params.strokeWidth);
        }

        /* 縦線 / Vertical lines */
        if (params.vOuter) {
            drawLine(group, xLeftOuter, vT, xLeftOuter, vB, strokeColor, params.strokeWidth);
        }
        drawLine(group, context.bLeft, vT, context.bLeft, vB, strokeColor, params.strokeWidth);
        drawLine(group, context.bRight, vT, context.bRight, vB, strokeColor, params.strokeWidth);
        if (params.vOuter) {
            drawLine(group, xRightOuter, vT, xRightOuter, vB, strokeColor, params.strokeWidth);
        }

        /* 縦分割線 / Vertical division lines */
        if (params.vDiv && params.vDivisions > 1) {
            var vDivStep = (context.bRight - context.bLeft) / params.vDivisions;
            for (var k = 1; k < params.vDivisions; k++) {
                var xDiv = context.bLeft + vDivStep * k;
                drawLine(group, xDiv, vT, xDiv, vB, strokeColor, params.strokeWidth);
            }
        }

        /* ガイド化（本番のみ） / Convert to guides (final only) */
        if (!isPreview && params.makeGuides) {
            for (var j = 0; j < group.pathItems.length; j++) {
                group.pathItems[j].guides = true;
            }
        }

        if (!params.groupItems && !params.makeGuides) {
            while (group.pageItems.length > 0) {
                group.pageItems[0].move(context.guideLayer, ElementPlacement.PLACEATBEGINNING);
            }
            group.remove();
            return null;
        }

        return group;
    }
    function changeValueByArrowKey(editText, allowNegative, integerOnly, ui, context) {
        editText.addEventListener("keydown", function (event) {
            var keyName = event.keyName;
            if (keyName !== "Up" && keyName !== "Down") {
                return;
            }

            var value = Number(editText.text);
            if (isNaN(value) || !isFinite(value)) {
                return;
            }

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else if (keyboard.altKey && !integerOnly) {
                delta = 0.1;
                if (keyName === "Up") {
                    value += delta;
                } else {
                    value -= delta;
                }
            } else {
                if (keyName === "Up") {
                    value += delta;
                } else {
                    value -= delta;
                }
            }

            if (!allowNegative && value < 0) {
                value = 0;
            }

            if (keyboard.altKey && !integerOnly) {
                value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
            } else {
                value = Math.round(value); /* 整数に丸め / Round to integer */
            }

            event.preventDefault();
            if (editText === ui.hScaleInput || editText === ui.vScaleInput) {
                editText.text = value.toFixed(1);
            } else {
                editText.text = String(value);
            }

            if (editText === ui.vScaleInput) {
                updateVerticalTopSpaceEnabled(ui);
            }

            if (ui.previewCheck.value) {
                updatePreview(ui, context);
            }
        });
    }

    function updatePreview(ui, context) {
        removePreview();
        var params = getParams(ui);
        if (!params) return;
        context.guideLayer = getOrCreateLayer(context.doc, params.guideLayerName);
        previewGroup = createLines(context, params, true);
        app.redraw();
    }

    function removePreview() {
        if (previewGroup) {
            try {
                previewGroup.remove();
            } catch (e) {
                try {
                    $.writeln("[logo-grid-maker] Failed to remove preview: " + e);
                } catch (logErr) { }
            }
            previewGroup = null;
            app.redraw();
        }
    }

    function drawLine(parent, x1, y1, x2, y2, color, sw) {
        var line = parent.pathItems.add();
        line.setEntirePath([[x1, y1], [x2, y2]]);
        line.stroked = true;
        line.filled = false;
        line.strokeWidth = sw;
        line.strokeColor = color;
    }

    function getSelectionBounds(items) {
        var first = true;
        var l, t, r, b;

        for (var i = 0; i < items.length; i++) {
            var ib = getItemBounds(items[i]);
            if (!ib) continue;

            if (first) {
                l = ib[0];
                t = ib[1];
                r = ib[2];
                b = ib[3];
                first = false;
            } else {
                if (ib[0] < l) l = ib[0];
                if (ib[1] > t) t = ib[1];
                if (ib[2] > r) r = ib[2];
                if (ib[3] < b) b = ib[3];
            }
        }

        if (first) return null;
        return [l, t, r, b];
    }

    function getItemBounds(item) {
        try {
            return item.visibleBounds;
        } catch (e) {
            return null;
        }
    }
})();