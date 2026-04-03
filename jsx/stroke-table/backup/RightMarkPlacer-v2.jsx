#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// ==================================================
// RightMarkPlacer.jsx
// スクリプトの概要：
// レイアウト間の関係性やフローを視覚的に示す右向き記号を、選択した複数オブジェクトの間に正確に配置します。
// 選択したオブジェクトを左から順に見て、隣り合うオブジェクト同士の「アキ」の中央に、右向きの記号（▶ / > / >> / → / ➡）を作成します。
//
// ・幅・間隔・位置は現在の定規単位、線幅は線の単位に合わせて調整可能
// ・位置（左右 / 上下）を微調整可能
// ・形状ごとに最適な初期値を自動適用
// ・▶ / > / >> / → / ➡ は高さ％でサイズを調整可能
// ・→ / ➡ は幅で長さを調整可能
// ・テキストは計測用に複製してアウトライン化し、見た目に近い天地中央で配置
// ・複数選択時は、左から順に見て隣り合うオブジェクト同士のアキごとに同じ形状を作成
// ・chevron1 / chevron2 では「天地を水平に」を使って、天地が水平な形状に切り替え可能
// ・プレビューで結果を確認しながら調整可能
//
// ※隣り合うオブジェクト同士の「アキ」が正の値（重なっていない）場合のみ、その箇所に作成されます
//
// 作成日：2026-03-28
// 更新日：2026-03-29
// ===================================================

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================
var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "右向き記号を配置",
        en: "Place Right-Facing Symbols"
    },
    alertOpenDocument: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    alertSelectTwoObjects: {
        ja: "オブジェクトを2つ以上選択してください。",
        en: "Please select two or more objects."
    },
    panelShape: {
        ja: "形状",
        en: "Shape"
    },
    panelCap: {
        ja: "先端",
        en: "End Style"
    },
    panelOptions: {
        ja: "オプション",
        en: "Options"
    },
    panelAdjust: {
        ja: "位置調整",
        en: "Position"
    },
    capNone: {
        ja: "なし",
        en: "None"
    },
    capRound: {
        ja: "丸型",
        en: "Round"
    },
    labelHeight: {
        ja: "高さ",
        en: "Height"
    },
    labelWidth: {
        ja: "幅",
        en: "Width"
    },
    labelGap: {
        ja: "間隔",
        en: "Gap"
    },
    labelStroke: {
        ja: "線幅",
        en: "Stroke"
    },
    alignVerticalCenter: {
        ja: "天地を水平に",
        en: "Align top and bottom horizontally"
    },
    labelAdjustX: {
        ja: "左右",
        en: "Horizontal"
    },
    labelAdjustY: {
        ja: "上下",
        en: "Vertical"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    alertPositiveNumber: {
        ja: "正の数値を入力してください。",
        en: "Please enter a positive value."
    },
    alertMaxHeight: {
        ja: "200% 以下の値を入力してください。",
        en: "Please enter a value of 200% or less."
    },
    alertNoGap: {
        ja: "2つのオブジェクトが重なっているか、アキがありません。",
        en: "The two objects overlap or there is no gap between them."
    },
    alertStrokePositive: {
        ja: "この形状では線幅に 0 より大きい値を入力してください。",
        en: "For this shape, enter a stroke width greater than 0."
    }
};

function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    return entry[lang] || entry.ja || key;
}

(function () {

    if (app.documents.length === 0) {
        alert(L("alertOpenDocument"));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    var measurementBoundsCache = [];

    function clearMeasurementBoundsCache() {
        measurementBoundsCache = [];
    }

    // ---- 選択チェック ----
    if (sel.length < 2) {
        alert(L("alertSelectTwoObjects"));
        return;
    }


    /* CMYKカラー生成ヘルパー / CMYK color helper */
    function makeCMYK(c, m, y, k) {
        var color = new CMYKColor();
        color.cyan = c;
        color.magenta = m;
        color.yellow = y;
        color.black = k;
        return color;
    }

    /* 計測用バウンズ取得 / Get measurement bounds */
    function copyBounds(bounds) {
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    }

    function getItemMeasurementBounds(item) {
        var dup = null;
        var outlined = null;
        var i;
        for (i = 0; i < measurementBoundsCache.length; i++) {
            if (measurementBoundsCache[i].item === item) {
                return copyBounds(measurementBoundsCache[i].bounds);
            }
        }

        try {
            if (item.typename === "TextFrame") {
                dup = item.duplicate();
                outlined = dup.createOutline();
                measurementBoundsCache.push({
                    item: item,
                    bounds: copyBounds(outlined.geometricBounds)
                });
                return copyBounds(outlined.geometricBounds);
            }
            measurementBoundsCache.push({
                item: item,
                bounds: copyBounds(item.geometricBounds)
            });
            return copyBounds(item.geometricBounds);
        } catch (e) {
            try {
                if (outlined) outlined.remove();
            } catch (_) { }
            try {
                if (dup) dup.remove();
            } catch (_) { }
            return copyBounds(item.geometricBounds);
        } finally {
            try {
                if (outlined) outlined.remove();
            } catch (_) { }
            try {
                if (dup) dup.remove();
            } catch (_) { }
        }
    }

    function getItemCenterX(item) {
        var b = getItemMeasurementBounds(item);
        return (b[0] + b[2]) / 2;
    }

    function getSortedSelectionItems() {
        var items = [];
        var i;
        for (i = 0; i < sel.length; i++) {
            items.push(sel[i]);
        }
        items.sort(function (a, b) {
            return getItemCenterX(a) - getItemCenterX(b);
        });
        return items;
    }

    /* 単位ユーティリティ / Unit utilities */
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

    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H
            case 6: return 1.0;                         // px
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    function getPreferenceUnitInfo(prefKey) {
        var code = app.preferences.getIntegerPreference(prefKey);
        return {
            code: code,
            label: getUnitLabel(code, prefKey),
            factor: getPtFactorFromUnitCode(code)
        };
    }

    var rulerUnitInfo = getPreferenceUnitInfo("rulerType");
    var strokeUnitInfo = getPreferenceUnitInfo("strokeUnits");

    function convertValueToPt(value, unitInfo) {
        return value * unitInfo.factor;
    }

    function convertPtToUnitValue(valuePt, unitInfo) {
        return valuePt / unitInfo.factor;
    }

    function roundDisplayValue(value) {
        return Math.round(value * 100) / 100;
    }

    function setFieldFromPt(editText, valuePt, unitInfo) {
        editText.text = String(roundDisplayValue(convertPtToUnitValue(valuePt, unitInfo)));
    }

    function parseFieldToPt(editText, unitInfo, allowNegative) {
        var value = parseFloat(editText.text);
        if (isNaN(value)) return NaN;
        if (!allowNegative && value < 0) value = 0;
        return convertValueToPt(value, unitInfo);
    }


    /* ダイアログ作成 / Create dialog */
    var dlg = createDialog();

    function createDialog() {
        var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = 20;
        dialog.spacing = 12;
        return dialog;
    }

    /* UIラベル幅定数 / UI label width constant */
    var UI_LABEL_WIDTH = 60;

    var ui = buildUI(dlg);

    function buildUI(dialog) {
        var mainRow = dialog.add("group");
        mainRow.orientation = "row";
        mainRow.alignChildren = ["fill", "top"];
        mainRow.spacing = 12;

        var leftCol = mainRow.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];
        leftCol.spacing = 10;

        var rightCol = mainRow.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["fill", "top"];
        rightCol.spacing = 10;

        var shapeUI = buildShapePanel(leftCol);
        var capUI = buildCapPanel(leftCol);
        var optionUI = buildOptionPanel(rightCol);
        var adjustUI = buildAdjustPanel(rightCol);
        var buttonUI = buildButtonRow(dialog);

        return {
            mainRow: mainRow,
            leftCol: leftCol,
            rightCol: rightCol,
            shapePanel: shapeUI.panel,
            capPanel: capUI.panel,
            sizePanel: optionUI.panel,
            adjustPanel: adjustUI.panel,
            btnGroup: buttonUI.group,
            radioTri: shapeUI.radioTri,
            radioChevron: shapeUI.radioChevron,
            radioChevronDouble: shapeUI.radioChevronDouble,
            radioArrow: shapeUI.radioArrow,
            radioArrow3: shapeUI.radioArrow3,
            radioCapNone: capUI.radioCapNone,
            radioCapRound: capUI.radioCapRound,
            inputGroup: optionUI.inputGroup,
            widthGroup: optionUI.widthGroup,
            inputField: optionUI.inputField,
            widthField: optionUI.widthField,
            gapField: optionUI.gapField,
            strokeField: optionUI.strokeField,
            chkAlignVerticalCenter: optionUI.chkAlignVerticalCenter,


            adjustField: adjustUI.adjustField,
            adjustVField: adjustUI.adjustVField,
            previewCb: buttonUI.previewCb,
            btnSpacer: buttonUI.btnSpacer,
            cancelBtn: buttonUI.cancelBtn,
            okBtn: buttonUI.okBtn
        };
    }

    function buildShapePanel(parent) {
        var panel = parent.add("panel", undefined, L("panelShape"));
        panel.orientation = "column";
        panel.alignChildren = ["left", "top"];
        panel.spacing = 6;
        panel.margins = [15, 20, 15, 10];

        var radioTri = panel.add("radiobutton", undefined, "\u25B6");
        var radioChevron = panel.add("radiobutton", undefined, ">");
        var radioChevronDouble = panel.add("radiobutton", undefined, ">>");
        var radioArrow = panel.add("radiobutton", undefined, "\u2192");
        var radioArrow3 = panel.add("radiobutton", undefined, "\u27A1");
        radioTri.value = true;

        return {
            panel: panel,
            radioTri: radioTri,
            radioChevron: radioChevron,
            radioChevronDouble: radioChevronDouble,
            radioArrow: radioArrow,
            radioArrow3: radioArrow3
        };
    }

    function buildCapPanel(parent) {
        var panel = parent.add("panel", undefined, L("panelCap"));
        panel.orientation = "column";
        panel.alignChildren = ["left", "top"];
        panel.spacing = 6;
        panel.margins = [12, 16, 12, 8];

        var radioCapNone = panel.add("radiobutton", undefined, L("capNone"));
        var radioCapRound = panel.add("radiobutton", undefined, L("capRound"));
        radioCapNone.value = true;

        return {
            panel: panel,
            radioCapNone: radioCapNone,
            radioCapRound: radioCapRound
        };
    }

    function buildOptionPanel(parent) {
        var panel = parent.add("panel", undefined, L("panelOptions"));
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.spacing = 8;
        panel.margins = [12, 16, 12, 8];

        var inputGroup = panel.add("group");
        inputGroup.orientation = "row";
        inputGroup.alignChildren = ["left", "center"];
        inputGroup.spacing = 8;

        var labelHeight = inputGroup.add("statictext", undefined, L("labelHeight"));
        labelHeight.preferredSize = [UI_LABEL_WIDTH, -1];
        labelHeight.justify = "right";

        var inputField = inputGroup.add("edittext", undefined, "20");
        inputField.characters = 4;
        inputField.active = true;

        inputGroup.add("statictext", undefined, "%");

        var widthGroup = panel.add("group");
        widthGroup.orientation = "row";
        widthGroup.alignChildren = ["left", "center"];
        widthGroup.spacing = 8;

        var labelWidth = widthGroup.add("statictext", undefined, L("labelWidth"));
        labelWidth.preferredSize = [UI_LABEL_WIDTH, -1];
        labelWidth.justify = "right";

        var widthField = widthGroup.add("edittext", undefined, "");
        widthField.characters = 4;

        widthGroup.add("statictext", undefined, rulerUnitInfo.label);

        var gapGroup = panel.add("group");
        gapGroup.orientation = "row";
        gapGroup.alignChildren = ["left", "center"];
        gapGroup.spacing = 8;

        var labelGap = gapGroup.add("statictext", undefined, L("labelGap"));
        labelGap.preferredSize = [UI_LABEL_WIDTH, -1];
        labelGap.justify = "right";

        var gapField = gapGroup.add("edittext", undefined, String(roundDisplayValue(convertPtToUnitValue(-1, rulerUnitInfo))));
        gapField.characters = 4;

        gapGroup.add("statictext", undefined, rulerUnitInfo.label);
        gapField.enabled = false;

        var strokeGroup = panel.add("group");
        strokeGroup.orientation = "row";
        strokeGroup.alignChildren = ["left", "center"];
        strokeGroup.spacing = 8;

        var labelStroke = strokeGroup.add("statictext", undefined, L("labelStroke"));
        labelStroke.preferredSize = [UI_LABEL_WIDTH, -1];
        labelStroke.justify = "right";

        var strokeField = strokeGroup.add("edittext", undefined, String(roundDisplayValue(convertPtToUnitValue(0.3, strokeUnitInfo))));
        strokeField.characters = 4;

        strokeGroup.add("statictext", undefined, strokeUnitInfo.label);

        var alignCenterGroup = panel.add("group");
        alignCenterGroup.orientation = "row";
        alignCenterGroup.alignChildren = ["center", "center"];
        alignCenterGroup.alignment = ["center", "top"];
        alignCenterGroup.margins = [0, 10, 0, 0];

        var chkAlignVerticalCenter = alignCenterGroup.add("checkbox", undefined, L("alignVerticalCenter"));
        chkAlignVerticalCenter.value = false;

        return {
            panel: panel,
            inputGroup: inputGroup,
            widthGroup: widthGroup,
            inputField: inputField,
            widthField: widthField,
            gapField: gapField,
            strokeField: strokeField,
            chkAlignVerticalCenter: chkAlignVerticalCenter
        };
    }

    function buildAdjustPanel(parent) {
        var panel = parent.add("panel", undefined, L("panelAdjust"));
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.spacing = 8;
        panel.margins = [12, 16, 12, 8];

        var adjustGroup = panel.add("group");
        adjustGroup.orientation = "row";
        adjustGroup.alignChildren = ["left", "center"];
        adjustGroup.spacing = 8;

        var labelAdjust = adjustGroup.add("statictext", undefined, L("labelAdjustX"));
        labelAdjust.preferredSize = [UI_LABEL_WIDTH, -1];
        labelAdjust.justify = "right";

        var adjustField = adjustGroup.add("edittext", undefined, "0");
        adjustField.characters = 4;

        adjustGroup.add("statictext", undefined, rulerUnitInfo.label);

        var adjustVGroup = panel.add("group");
        adjustVGroup.orientation = "row";
        adjustVGroup.alignChildren = ["left", "center"];
        adjustVGroup.spacing = 8;

        var labelAdjustV = adjustVGroup.add("statictext", undefined, L("labelAdjustY"));
        labelAdjustV.preferredSize = [UI_LABEL_WIDTH, -1];
        labelAdjustV.justify = "right";

        var adjustVField = adjustVGroup.add("edittext", undefined, "0");
        adjustVField.characters = 4;

        adjustVGroup.add("statictext", undefined, rulerUnitInfo.label);

        return {
            panel: panel,
            adjustField: adjustField,
            adjustVField: adjustVField
        };
    }



    function buildButtonRow(dialog) {
        var group = dialog.add("group");
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        group.alignment = "fill";
        group.spacing = 8;
        group.margins = [0, 10, 0, 0];

        var previewCb = group.add("checkbox", undefined, L("preview"));
        previewCb.value = false;

        var btnSpacer = group.add("group");
        btnSpacer.alignment = ["fill", "center"];

        var cancelBtn = group.add("button", undefined, L("cancel"), { name: "cancel" });
        var okBtn = group.add("button", undefined, L("ok"), { name: "ok" });

        return {
            group: group,
            previewCb: previewCb,
            btnSpacer: btnSpacer,
            cancelBtn: cancelBtn,
            okBtn: okBtn
        };
    }

    /* プレビュー用変数 / Preview variables */
    var previewPath = null;

    /* 形状定義 / Shape definitions */
    var SHAPE_CONFIG = {
        tri: {
            radioKey: "radioTri",
            enableCapPanel: false,
            forceFillOnly: true,
            enableStrokeInput: false,
            requirePositiveStroke: false,
            enableHeightInput: true,
            enableGap: false,
            defaultHeightPercent: 30,
            defaultGap: -1,
            defaultStrokePt: 0.3,
            calcDefaultWidth: function (inputVal, totalHeight) {
                var triHeight = totalHeight * (inputVal / 100);
                return Math.round(triHeight * 100) / 100;
            }
        },
        arrow: {
            radioKey: "radioArrow",
            enableCapPanel: true,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: true,
            enableGap: false,
            defaultHeightPercent: 20,
            defaultGap: -1,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight, gapWidth) {
                var arrowWidth = gapWidth * 0.75;
                return Math.round(arrowWidth * 100) / 100;
            }
        },
        arrow3: {
            radioKey: "radioArrow3",
            enableCapPanel: false,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: false,
            enableHeightInput: true,
            enableGap: false,
            defaultHeightPercent: 50,
            defaultGap: -1,
            defaultStrokePt: 1.2,
            calcDefaultWidth: function (inputVal, totalHeight, gapWidth) {
                var arrowWidth = gapWidth * 0.75;
                return Math.round(arrowWidth * 100) / 100;
            }
        },
        chevron: {
            radioKey: "radioChevron",
            enableCapPanel: true,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: true,
            enableGap: false,
            defaultHeightPercent: 50,
            defaultGap: -1,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight) {
                var chevHeight = totalHeight * (inputVal / 100);
                return Math.round(chevHeight * 100) / 100;
            }
        },
        chevron2: {
            radioKey: "radioChevronDouble",
            enableCapPanel: true,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: true,
            enableGap: true,
            defaultHeightPercent: 50,
            defaultGap: 0,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight) {
                var chevHeight = totalHeight * (inputVal / 100);
                return Math.round(chevHeight * 100) / 100;
            }
        }
    };
    var widthManuallySet = false;

    /* 幅の初期値を自動計算 / Auto-calculate default width */
    function calcDefaultWidth() {
        clearMeasurementBoundsCache();
        var shapeKey = getShapeType();
        var shapeConfig = SHAPE_CONFIG[shapeKey];
        var inputVal = parseFloat(ui.inputField.text);
        var items = getSortedSelectionItems();
        var b0_init = getItemMeasurementBounds(items[0]);
        var b1_init = getItemMeasurementBounds(items[1]);
        var totalHeight = Math.max(b0_init[1], b1_init[1]) - Math.min(b0_init[3], b1_init[3]);
        var leftObj = (b0_init[0] < b1_init[0]) ? b0_init : b1_init;
        var rightObj = (b0_init[0] < b1_init[0]) ? b1_init : b0_init;
        var gapWidth = rightObj[0] - leftObj[2];
        var defaultWidth;
        if (shapeConfig.enableHeightInput && (isNaN(inputVal) || inputVal <= 0)) return;
        defaultWidth = shapeConfig.calcDefaultWidth(inputVal, totalHeight, gapWidth);
        if (defaultWidth === null || typeof defaultWidth === "undefined") {
            ui.widthField.text = "";
        } else {
            setFieldFromPt(ui.widthField, defaultWidth, rulerUnitInfo);
        }
    }

    /* 形状ごとの初期値適用 / Apply shape defaults */
    function applyShapeDefaults() {
        var shapeKey = getShapeType();
        var shapeConfig = SHAPE_CONFIG[shapeKey];
        if (shapeConfig.enableHeightInput && shapeConfig.defaultHeightPercent !== undefined) {
            ui.inputField.text = String(shapeConfig.defaultHeightPercent);
        }
        if (!widthManuallySet) {
            calcDefaultWidth();
        }
        setFieldFromPt(ui.gapField, shapeConfig.defaultGap, rulerUnitInfo);
        ui.gapField.enabled = !!shapeConfig.enableGap;
        setFieldFromPt(ui.strokeField, shapeConfig.defaultStrokePt, strokeUnitInfo);
        ui.capPanel.enabled = !!shapeConfig.enableCapPanel;
        ui.inputGroup.enabled = !!shapeConfig.enableHeightInput;
        ui.strokeField.enabled = !!shapeConfig.enableStrokeInput;

        if (ui.chkAlignVerticalCenter) {
            var isChevronShape = (shapeKey === "chevron" || shapeKey === "chevron2");
            ui.chkAlignVerticalCenter.enabled = isChevronShape;
            if (!isChevronShape) {
                ui.chkAlignVerticalCenter.value = false;
            }
        }

        if (!shapeConfig.enableCapPanel) {
            ui.radioCapNone.value = true;
            ui.radioCapRound.value = false;
        }
        if (!shapeConfig.enableStrokeInput) {
            ui.strokeField.text = String(roundDisplayValue(convertPtToUnitValue(shapeConfig.defaultStrokePt || 0, strokeUnitInfo)));
        }
    }

    /* 矢印キーで値を増減 / Adjust values with arrow keys */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0 && editText !== ui.gapField) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0 && editText !== ui.gapField) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = String(roundDisplayValue(value));
            if (editText === ui.inputField && !widthManuallySet) {
                calcDefaultWidth();
            }
            if (editText === ui.widthField) {
                widthManuallySet = true;
            }
            updatePreview();
        });
    }

    changeValueByArrowKey(ui.inputField);
    changeValueByArrowKey(ui.widthField);
    changeValueByArrowKey(ui.strokeField);
    changeValueByArrowKey(ui.gapField);
    changeValueByArrowKey(ui.adjustField);
    changeValueByArrowKey(ui.adjustVField);

    /* 選択中の形状を取得 / Get current shape */
    function getShapeType() {
        for (var key in SHAPE_CONFIG) {
            if (!SHAPE_CONFIG.hasOwnProperty(key)) continue;
            var radioKey = SHAPE_CONFIG[key].radioKey;
            if (ui[radioKey] && ui[radioKey].value) {
                return key;
            }
        }
        return "tri";
    }
    /* 形状ラジオボタン取得ヘルパー / Get shape radio controls */
    function getShapeRadioControls() {
        var radios = [];
        for (var key in SHAPE_CONFIG) {
            if (!SHAPE_CONFIG.hasOwnProperty(key)) continue;
            radios.push(ui[SHAPE_CONFIG[key].radioKey]);
        }
        return radios;
    }

    function validateStrokeWidthForShape(shapeKey, strokeW, showAlert) {
        var shapeConfig = SHAPE_CONFIG[shapeKey];
        if (!shapeConfig || !shapeConfig.requirePositiveStroke) {
            return true;
        }
        if (isNaN(strokeW) || strokeW <= 0) {
            if (showAlert) {
                alert(L("alertStrokePositive"));
            }
            return false;
        }
        return true;
    }



    /* シェイプ描画ヘルパー / Shape drawing helpers */
    function createTriShape(cx, cy, w, h, strokeW, forceFillOnly) {
        var points = [
            [cx + w / 2, cy],
            [cx - w / 2, cy + h / 2],
            [cx - w / 2, cy - h / 2]
        ];

        var path = doc.activeLayer.pathItems.add();
        path.setEntirePath(points);
        path.closed = true;
        path.filled = true;
        path.fillColor = makeCMYK(0, 0, 0, 100);
        if (!forceFillOnly && strokeW > 0) {
            path.stroked = true;
            path.strokeWidth = strokeW;
            path.strokeColor = makeCMYK(0, 0, 0, 100);
        } else {
            path.stroked = false;
        }
        return path;
    }

    function createArrowShape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW) {
        var gapW = gapRight - gapLeft;
        var w = (widthVal > 0) ? widthVal : (gapW * 0.75);
        var lineW = strokeW;
        var headSize = h / 2;

        var shaft = doc.activeLayer.pathItems.add();
        shaft.setEntirePath([
            [cx - w / 2, cy],
            [cx + w / 2, cy]
        ]);
        shaft.closed = false;
        shaft.filled = false;
        shaft.stroked = true;
        shaft.strokeWidth = lineW;
        shaft.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            shaft.strokeCap = StrokeCap.ROUNDENDCAP;
        }

        var head = doc.activeLayer.pathItems.add();
        head.setEntirePath([
            [cx + w / 2 - headSize, cy + headSize],
            [cx + w / 2, cy],
            [cx + w / 2 - headSize, cy - headSize]
        ]);
        head.closed = false;
        head.filled = false;
        head.stroked = true;
        head.strokeWidth = lineW;
        head.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            head.strokeCap = StrokeCap.ROUNDENDCAP;
            head.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        }

        var group = doc.activeLayer.groupItems.add();
        shaft.move(group, ElementPlacement.INSIDE);
        head.move(group, ElementPlacement.INSIDE);
        return group;
    }

    function createArrow3Shape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW) {
        var gapW = gapRight - gapLeft;
        var w = (widthVal > 0) ? widthVal : (gapW * 0.75);
        var lineW = (strokeW > 0) ? strokeW * 3 : 3;
        var r = h / 2;

        var shaft = doc.activeLayer.pathItems.add();
        shaft.setEntirePath([
            [cx - w / 2, cy],
            [cx + w / 2 - r, cy]
        ]);
        shaft.closed = false;
        shaft.filled = false;
        shaft.stroked = true;
        shaft.strokeWidth = lineW;
        shaft.strokeColor = makeCMYK(0, 0, 0, 100);

        var head = doc.activeLayer.pathItems.add();
        head.setEntirePath([
            [cx + w / 2 - r, cy - r],
            [cx + w / 2, cy],
            [cx + w / 2 - r, cy + r]
        ]);
        head.closed = true;
        head.filled = true;
        head.fillColor = makeCMYK(0, 0, 0, 100);
        head.stroked = false;

        return {
            shaft: shaft,
            head: head
        };
    }

    function createFlatChevronShape(cx, cy, h, widthVal, strokeW) {
        var fullW = (widthVal > 0) ? widthVal : h;
        var thickness = (strokeW > 0) ? strokeW : 1;
        var halfH = h / 2;
        var leftX = cx - fullW / 2;
        var tipX = cx + fullW / 2;
        var topY = cy + halfH;
        var bottomY = cy - halfH;

        var group = doc.activeLayer.groupItems.add();

        var topArm = doc.activeLayer.pathItems.add();
        topArm.setEntirePath([
            [leftX, topY],
            [leftX + thickness, topY],
            [tipX, cy],
            [tipX - thickness, cy]
        ]);
        topArm.closed = true;
        topArm.filled = true;
        topArm.fillColor = makeCMYK(0, 0, 0, 100);
        topArm.stroked = false;
        topArm.move(group, ElementPlacement.INSIDE);

        var bottomArm = doc.activeLayer.pathItems.add();
        bottomArm.setEntirePath([
            [leftX, bottomY],
            [leftX + thickness, bottomY],
            [tipX, cy],
            [tipX - thickness, cy]
        ]);
        bottomArm.closed = true;
        bottomArm.filled = true;
        bottomArm.fillColor = makeCMYK(0, 0, 0, 100);
        bottomArm.stroked = false;
        bottomArm.move(group, ElementPlacement.INSIDE);

        return group;
    }

    function createFlatChevron2Shape(cx, cy, h, widthVal, strokeW) {
        var fullW = (widthVal > 0) ? widthVal : h;
        var gap = parseFieldToPt(ui.gapField, rulerUnitInfo, true);
        if (isNaN(gap)) gap = 0;

        var leftCenter = cx - (fullW + gap) / 2;
        var rightCenter = cx + (fullW + gap) / 2;

        var group = doc.activeLayer.groupItems.add();

        var leftShape = createFlatChevronShape(leftCenter, cy, h, fullW, strokeW);
        var rightShape = createFlatChevronShape(rightCenter, cy, h, fullW, strokeW);

        leftShape.move(group, ElementPlacement.INSIDE);
        rightShape.move(group, ElementPlacement.INSIDE);

        return group;
    }

    function createChevronShape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW) {
        var lineW = strokeW;
        var chevH = h / 2;
        var chevW = (widthVal > 0) ? widthVal / 2 : chevH;

        if (ui.chkAlignVerticalCenter && ui.chkAlignVerticalCenter.value) {
            return createFlatChevronShape(cx, cy, h, widthVal, strokeW);
        }

        var head = doc.activeLayer.pathItems.add();
        head.setEntirePath([
            [cx - chevW / 2, cy + chevH],
            [cx + chevW / 2, cy],
            [cx - chevW / 2, cy - chevH]
        ]);
        head.closed = false;
        head.filled = false;
        head.stroked = true;
        head.strokeWidth = lineW;
        head.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            head.strokeCap = StrokeCap.ROUNDENDCAP;
            head.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        }

        return head;
    }

    function createChevron2Shape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW) {
        var lineW = strokeW;
        var chevH = h / 2;
        var chevW = (widthVal > 0) ? widthVal / 2 : chevH;
        var gap = parseFieldToPt(ui.gapField, rulerUnitInfo, true);
        if (isNaN(gap)) gap = 0;

        if (ui.chkAlignVerticalCenter && ui.chkAlignVerticalCenter.value) {
            return createFlatChevron2Shape(cx, cy, h, widthVal, strokeW);
        }

        var totalW = chevW + gap + chevW;
        var leftStart = cx - totalW / 2;

        var headL = doc.activeLayer.pathItems.add();
        headL.setEntirePath([
            [leftStart, cy + chevH],
            [leftStart + chevW, cy],
            [leftStart, cy - chevH]
        ]);
        headL.closed = false;
        headL.filled = false;
        headL.stroked = true;
        headL.strokeWidth = lineW;
        headL.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            headL.strokeCap = StrokeCap.ROUNDENDCAP;
            headL.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        }

        var rightStart = leftStart + chevW + gap;
        var headR = doc.activeLayer.pathItems.add();
        headR.setEntirePath([
            [rightStart, cy + chevH],
            [rightStart + chevW, cy],
            [rightStart, cy - chevH]
        ]);
        headR.closed = false;
        headR.filled = false;
        headR.stroked = true;
        headR.strokeWidth = lineW;
        headR.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            headR.strokeCap = StrokeCap.ROUNDENDCAP;
            headR.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        }

        var group = doc.activeLayer.groupItems.add();
        headL.move(group, ElementPlacement.INSIDE);
        headR.move(group, ElementPlacement.INSIDE);
        return group;
    }

    function finalizeArrow3Appearance(parts) {
        var outlinedShaft;
        if (!parts || !parts.shaft || !parts.head) return parts;

        doc.selection = [parts.shaft];
        app.executeMenuCommand('Live Outline Stroke');
        if (doc.selection && doc.selection.length > 0) {
            outlinedShaft = doc.selection[0];
        } else {
            outlinedShaft = parts.shaft;
        }

        doc.selection = [outlinedShaft, parts.head];
        app.executeMenuCommand('group');
        app.executeMenuCommand('Live Pathfinder Add');
        if (doc.selection && doc.selection.length > 0) {
            return doc.selection[0];
        }
        return outlinedShaft;
    }

    function createShapeBetweenItems(leftItem, rightItem, inputVal, widthVal, offsetX, offsetY, strokeW) {
        var b0 = getItemMeasurementBounds(leftItem);
        var b1 = getItemMeasurementBounds(rightItem);

        var leftObj, rightObj;
        if (b0[0] < b1[0]) {
            leftObj = b0;
            rightObj = b1;
        } else {
            leftObj = b1;
            rightObj = b0;
        }

        var gapLeft = leftObj[2];
        var gapRight = rightObj[0];

        if (gapRight - gapLeft <= 0) {
            return null;
        }

        var cx = (gapLeft + gapRight) / 2 + (offsetX || 0);

        var allTop = Math.max(b0[1], b1[1]);
        var allBottom = Math.min(b0[3], b1[3]);
        var cy = (allTop + allBottom) / 2 + (offsetY || 0);

        var totalHeight = allTop - allBottom;
        var h = totalHeight * (inputVal / 100);
        var w = (widthVal > 0) ? widthVal : h;
        var shape = getShapeType();
        var shapeConfig = SHAPE_CONFIG[shape];

        if (shape === "tri") {
            return createTriShape(cx, cy, w, h, strokeW, !!shapeConfig.forceFillOnly);
        }
        if (shape === "arrow") {
            return createArrowShape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW);
        }
        if (shape === "arrow3") {
            var arrow3Parts = createArrow3Shape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW);
            return finalizeArrow3Appearance(arrow3Parts);
        }
        if (shape === "chevron") {
            return createChevronShape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW);
        }
        if (shape === "chevron2") {
            return createChevron2Shape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW);
        }

        return createTriShape(cx, cy, w, h, strokeW, !!shapeConfig.forceFillOnly);
    }

    /* 形状を作成する共通関数 / Shared shape creation function */
    function createShape(inputVal, widthVal, offsetX, offsetY, strokeW) {
        clearMeasurementBoundsCache();

        var items = getSortedSelectionItems();
        var createdItems = [];
        var i;
        var shapeItem;
        var group;

        for (i = 0; i < items.length - 1; i++) {
            shapeItem = createShapeBetweenItems(items[i], items[i + 1], inputVal, widthVal, offsetX, offsetY, strokeW);
            if (shapeItem !== null) {
                createdItems.push(shapeItem);
            }
        }

        if (createdItems.length === 0) {
            return null;
        }

        if (createdItems.length === 1) {
            return createdItems[0];
        }

        group = doc.activeLayer.groupItems.add();
        for (i = 0; i < createdItems.length; i++) {
            createdItems[i].move(group, ElementPlacement.INSIDE);
        }
        return group;
    }

    /* プレビューを削除 / Remove preview */
    function removePreview() {
        if (previewPath !== null) {
            try { previewPath.remove(); } catch (e) { }
            previewPath = null;
            app.redraw();
        }
    }

    /* プレビューを更新 / Update preview */
    function updatePreview() {
        removePreview();
        if (!ui.previewCb.value) return;

        var inputVal = parseFloat(ui.inputField.text);
        if (isNaN(inputVal) || inputVal <= 0 || inputVal > 200) return;

        var widthVal = parseFieldToPt(ui.widthField, rulerUnitInfo, false);
        if (isNaN(widthVal) || widthVal < 0) widthVal = 0;

        var offsetX = parseFieldToPt(ui.adjustField, rulerUnitInfo, true);
        if (isNaN(offsetX)) offsetX = 0;

        var offsetY = parseFieldToPt(ui.adjustVField, rulerUnitInfo, true);
        if (isNaN(offsetY)) offsetY = 0;

        var strokeW = parseFieldToPt(ui.strokeField, strokeUnitInfo, false);
        if (isNaN(strokeW) || strokeW < 0) strokeW = 0;

        if (!validateStrokeWidthForShape(getShapeType(), strokeW, false)) return;

        previewPath = createShape(inputVal, widthVal, offsetX, offsetY, strokeW);
        app.redraw();
    }

    /* イベント / Events */
    ui.radioCapNone.onClick = function () {
        updatePreview();
    };

    ui.radioCapRound.onClick = function () {
        updatePreview();
    };

    if (ui.chkAlignVerticalCenter) {
        ui.chkAlignVerticalCenter.onClick = function () {
            updatePreview();
        };
    }

    ui.previewCb.onClick = function () {
        updatePreview();
    };

    ui.inputField.onChanging = function () {
        if (!widthManuallySet) {
            calcDefaultWidth();
        }
        updatePreview();
    };

    ui.widthField.onChanging = function () {
        updatePreview();
    };

    ui.adjustField.onChanging = function () {
        updatePreview();
    };

    ui.adjustVField.onChanging = function () {
        updatePreview();
    };

    ui.strokeField.onChanging = function () {
        updatePreview();
    };

    ui.gapField.onChanging = function () {
        updatePreview();
    };

    /* 形状パネルのラジオボタン選択変更時に初期値適用 / Apply defaults when shape selection changes */
    var shapeRadios = getShapeRadioControls();
    for (var i = 0; i < shapeRadios.length; i++) {
        shapeRadios[i].onClick = function () {
            widthManuallySet = false;
            applyShapeDefaults();
            updatePreview();
        };
    }

    /* OKボタン処理 / OK button handler */
    ui.okBtn.onClick = function () {

        var inputVal = parseFloat(ui.inputField.text);

        if (isNaN(inputVal) || inputVal <= 0) {
            alert(L("alertPositiveNumber"));
            ui.inputField.active = true;
            return;
        }
        if (inputVal > 200) {
            alert(L("alertMaxHeight"));
            ui.inputField.active = true;
            return;
        }

        // プレビューがあればそのまま確定
        if (previewPath !== null) {
            doc.selection = [previewPath];
            previewPath = null;
            dlg.close(1);
            return;
        }

        dlg.close(1);

        var widthVal = parseFieldToPt(ui.widthField, rulerUnitInfo, false);
        if (isNaN(widthVal) || widthVal < 0) widthVal = 0;

        var offsetX = parseFieldToPt(ui.adjustField, rulerUnitInfo, true);
        if (isNaN(offsetX)) offsetX = 0;

        var offsetY = parseFieldToPt(ui.adjustVField, rulerUnitInfo, true);
        if (isNaN(offsetY)) offsetY = 0;

        var strokeW = parseFieldToPt(ui.strokeField, strokeUnitInfo, false);
        if (isNaN(strokeW) || strokeW < 0) strokeW = 0;

        if (!validateStrokeWidthForShape(getShapeType(), strokeW, true)) {
            ui.strokeField.active = true;
            return;
        }

        var path = createShape(inputVal, widthVal, offsetX, offsetY, strokeW);
        if (path === null) {
            alert(L("alertNoGap"));
            return;
        }
        doc.selection = [path];
    };

    /* キャンセルボタン処理 / Cancel button handler */
    ui.cancelBtn.onClick = function () {
        removePreview();
        dlg.close(0);
    };

    /* 初期値適用 / Apply initial defaults */
    applyShapeDefaults();

    dlg.show();

})();