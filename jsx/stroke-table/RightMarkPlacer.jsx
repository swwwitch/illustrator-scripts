#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// ==================================================
// RightMarkPlacer.jsx
// スクリプトの概要：
// 選択した複数オブジェクトを左から順に見て、隣り合うオブジェクト同士の「アキ」の中央に、右向き記号を配置します。
// レイアウトの流れ・関係性を視覚的に示すためのツールです。
//
// ・右向き記号（▶ / > / >> / ─ / → / ➡ / ＋ / ×）を作成
// ・複数選択時、各オブジェクト間すべてに同形状を自動配置
// ・幅は「最も狭いアキ」を基準に安全側で自動決定（手動入力も可能）
// ・高さ（％）・位置（左右／上下）を調整可能
// ・単位は現在の定規設定に追従
//
// ・▶ は「凹み」で形状を変形可能（幅の80%まで）
// ・▶ は「角丸」オプションで角丸形状に変更可能（凹みの有無に関係なく適用）
//
// ・→ / ➡ は幅で長さを調整
// ・➡ は入力線幅の3倍で描画
// ・＋ / × は幅でサイズを調整（× は45°回転）
//
// ・左右逆（反転）、先端形状、プレビューに対応
// ・キーボード操作：F（なし）/ R（丸型）/ V（左右逆）
//
// ※オブジェクトが重なっている場合やアキがない場合は作成されません
//
// 作成日：2026-03-28
// 更新日：2026-04-04 v1.3 機能整理・凹み／角丸／ショートカット対応
// ===================================================

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================
var SCRIPT_VERSION = "v1.2";

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
    labelInset: {
        ja: "凹み",
        en: "Inset"
    },
    labelStroke: {
        ja: "線幅",
        en: "Stroke"
    },
    roundCorners: {
        ja: "角丸",
        en: "Rounded corners"
    },
    alignVerticalCenter: {
        ja: "天地を水平に",
        en: "Keep top and bottom edges horizontal"
    },
    mirrorHorizontal: {
        ja: "左右逆",
        en: "Mirror horizontally"
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
        ja: "隣り合うオブジェクト間に作成できるアキがありません。",
        en: "There is no usable gap between adjacent objects."
    },
    alertStrokePositive: {
        ja: "線幅は 0.25 以上の値を入力してください。",
        en: "Please enter a stroke width of 0.25 or greater."
    },
    alertWidthPositive: {
        ja: "幅は 0 以上の値を入力してください。",
        en: "Please enter a width value of 0 or greater."
    },
    alertInvalidValue: {
        ja: "入力値を確認してください。",
        en: "Please check the input values."
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
    var selectionItems = [];
    var measurementBoundsCache = [];
    var si;

    for (si = 0; si < doc.selection.length; si++) {
        selectionItems.push(doc.selection[si]);
    }

    function clearMeasurementBoundsCache() {
        measurementBoundsCache = [];
    }

    // ---- 選択チェック ----
    if (selectionItems.length < 2) {
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
        var bounds;
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
                bounds = copyBounds(outlined.geometricBounds);
                measurementBoundsCache.push({
                    item: item,
                    bounds: bounds
                });
                return copyBounds(bounds);
            }
            bounds = copyBounds(item.geometricBounds);
            measurementBoundsCache.push({
                item: item,
                bounds: bounds
            });
            return copyBounds(bounds);
        } catch (e) {
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
        for (i = 0; i < selectionItems.length; i++) {
            items.push(selectionItems[i]);
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

    /* キーボードショートカット / Keyboard shortcuts */
    function getKeyboardShortcutActions() {
        return {
            F: function () {
                if (ui.radioCapNone && ui.radioCapNone.enabled) {
                    ui.radioCapNone.value = true;
                    ui.radioCapRound.value = false;
                    updatePreview();
                    return true;
                }
                return false;
            },
            R: function () {
                if (ui.radioCapRound && ui.radioCapRound.enabled) {
                    ui.radioCapRound.value = true;
                    ui.radioCapNone.value = false;
                    updatePreview();
                    return true;
                }
                return false;
            },
            V: function () {
                if (ui.chkMirror && ui.chkMirror.enabled) {
                    ui.chkMirror.value = !ui.chkMirror.value;
                    updatePreview();
                    return true;
                }
                return false;
            }
        };
    }

    function addKeyboardHandlers(dialog) {
        dialog.addEventListener("keydown", function (event) {
            var key = event.keyName;
            var actions = getKeyboardShortcutActions();
            var action = actions[key];

            if (action && action()) {
                event.preventDefault();
            }
        });
    }
    addKeyboardHandlers(dlg);

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
            radioDash: shapeUI.radioDash,
            radioPlus: shapeUI.radioPlus,
            radioMultiply: shapeUI.radioMultiply,
            chkMirror: shapeUI.chkMirror,
            radioCapNone: capUI.radioCapNone,
            radioCapRound: capUI.radioCapRound,

            inputGroup: optionUI.inputGroup,
            widthGroup: optionUI.widthGroup,
            insetGroup: optionUI.insetGroup,

            inputField: optionUI.inputField,
            widthField: optionUI.widthField,
            insetField: optionUI.insetField,
            gapField: optionUI.gapField,
            strokeLabel: optionUI.strokeLabel,
            strokeField: optionUI.strokeField,

            chkAlignVerticalCenter: optionUI.chkAlignVerticalCenter,
            chkRoundCorners: optionUI.chkRoundCorners,

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
        var radioDash = panel.add("radiobutton", undefined, "\u2500");
        var radioArrow = panel.add("radiobutton", undefined, "\u2192");
        var radioArrow3 = panel.add("radiobutton", undefined, "\u27A1");
        radioArrow3.helpTip = (lang === "ja") ? "線幅入力の3倍で描画します。" : "Draws using 3× the entered stroke width.";
        var radioPlus = panel.add("radiobutton", undefined, "\uFF0B");
        var radioMultiply = panel.add("radiobutton", undefined, "\u00D7");

        var mirrorGroup = panel.add("group");
        mirrorGroup.orientation = "row";
        mirrorGroup.alignChildren = ["left", "center"];
        mirrorGroup.alignment = ["left", "top"];
        mirrorGroup.margins = [0, 6, 0, 0];

        var chkMirror = mirrorGroup.add("checkbox", undefined, L("mirrorHorizontal"));
        chkMirror.value = false;

        radioTri.value = true;

        return {
            panel: panel,
            radioTri: radioTri,
            radioChevron: radioChevron,
            radioChevronDouble: radioChevronDouble,
            radioArrow: radioArrow,
            radioArrow3: radioArrow3,
            radioDash: radioDash,
            radioPlus: radioPlus,
            radioMultiply: radioMultiply,
            chkMirror: chkMirror
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

        var insetGroup = panel.add("group");
        insetGroup.orientation = "row";
        insetGroup.alignChildren = ["left", "center"];
        insetGroup.spacing = 8;

        var labelInset = insetGroup.add("statictext", undefined, L("labelInset"));
        labelInset.preferredSize = [UI_LABEL_WIDTH, -1];
        labelInset.justify = "right";

        var insetField = insetGroup.add("edittext", undefined, "0");
        insetField.characters = 4;
        // Tooltip for inset field limit
        insetField.helpTip = (lang === "ja") ? "幅の80%が上限です" : "Max 80% of width";

        insetGroup.add("statictext", undefined, rulerUnitInfo.label);
        insetField.enabled = false;

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
        alignCenterGroup.alignChildren = ["left", "center"];
        alignCenterGroup.alignment = ["left", "top"];
        alignCenterGroup.margins = [40, 10, 0, 0];

        var chkAlignVerticalCenter = alignCenterGroup.add("checkbox", undefined, L("alignVerticalCenter"));
        chkAlignVerticalCenter.value = false;


        var roundCornersGroup = panel.add("group");
        roundCornersGroup.orientation = "row";
        roundCornersGroup.alignChildren = ["left", "center"];
        roundCornersGroup.alignment = ["left", "top"];
        roundCornersGroup.margins = [40, 0, 0, 0];

        var chkRoundCorners = roundCornersGroup.add("checkbox", undefined, L("roundCorners"));
        chkRoundCorners.value = true;

        return {
            panel: panel,
            inputGroup: inputGroup,
            widthGroup: widthGroup,
            insetGroup: insetGroup,
            inputField: inputField,
            widthField: widthField,
            insetField: insetField,
            gapField: gapField,
            strokeLabel: labelStroke,
            strokeField: strokeField,
            chkAlignVerticalCenter: chkAlignVerticalCenter,
            chkRoundCorners: chkRoundCorners
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
    var previewPath = [];

    /* 形状定義 / Shape definitions */
    var SHAPE_CONFIG = {
        tri: {
            radioKey: "radioTri",
            enableCapPanel: false,
            forceFillOnly: true,
            enableStrokeInput: false,
            requirePositiveStroke: false,
            enableHeightInput: true,
            enableMirror: true,
            enableGap: false,

            enableInset: true,
            enableRoundCorners: true,
            defaultHeightPercent: 30,

            defaultGap: -1,
            defaultInsetPt: 0,
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
            enableMirror: true,
            enableGap: false,
            enableInset: false,
            enableRoundCorners: false,
            defaultHeightPercent: 30,
            defaultGap: -1,
            defaultInsetPt: 0,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight, gapWidth) {
                var arrowWidth = gapWidth * 0.7;
                return Math.round(arrowWidth * 100) / 100;
            }
        },
        arrow3: {
            radioKey: "radioArrow3",
            enableCapPanel: false,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: true,
            enableMirror: true,
            enableGap: false,
            enableInset: false,
            enableRoundCorners: false,
            defaultHeightPercent: 50,
            defaultGap: -1,
            defaultInsetPt: 0,
            defaultStrokePt: 1.2,
            calcDefaultWidth: function (inputVal, totalHeight, gapWidth) {
                var arrowWidth = gapWidth * 0.7;
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
            enableMirror: true,
            enableGap: false,
            enableInset: false,
            defaultHeightPercent: 50,
            defaultGap: -1,
            defaultInsetPt: 0,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight) {
                var chevHeight = totalHeight * (inputVal / 100);
                return Math.round(chevHeight * 100) / 100;
            }
        },
        dash: {
            radioKey: "radioDash",
            enableCapPanel: true,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: false,
            enableMirror: false,
            enableGap: false,
            enableInset: false,
            defaultHeightPercent: undefined,
            defaultGap: -1,
            defaultInsetPt: 0,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight, gapWidth) {
                return Math.round(gapWidth * 0.7 * 100) / 100;
            }
        },
        plus: {
            radioKey: "radioPlus",
            enableCapPanel: true,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: false,
            enableMirror: false,
            enableGap: false,
            enableInset: false,
            defaultHeightPercent: undefined,
            defaultGap: -1,
            defaultInsetPt: 0,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight, gapWidth) {
                return Math.round(gapWidth * 0.35 * 100) / 100;
            }
        },
        multiply: {
            radioKey: "radioMultiply",
            enableCapPanel: true,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: false,
            enableMirror: false,
            enableGap: false,
            enableInset: false,
            defaultHeightPercent: undefined,
            defaultGap: -1,
            defaultInsetPt: 0,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight, gapWidth) {
                return Math.round(gapWidth * 0.35 * 100) / 100;
            }
        },
        chevron2: {
            radioKey: "radioChevronDouble",
            enableCapPanel: true,
            forceFillOnly: false,
            enableStrokeInput: true,
            requirePositiveStroke: true,
            enableHeightInput: true,
            enableMirror: true,
            enableGap: true,
            enableInset: false,
            defaultHeightPercent: 50,
            defaultGap: 0,
            defaultInsetPt: 0,
            defaultStrokePt: 0.6,
            calcDefaultWidth: function (inputVal, totalHeight) {
                var chevHeight = totalHeight * (inputVal / 100);
                return Math.round(chevHeight * 100) / 100;
            }
        }
    };
    var widthManuallySet = false;

    function resetWidthToAuto() {
        widthManuallySet = false;
        calcDefaultWidth();
    }

    /* 幅の初期値を自動計算 / Auto-calculate default width */
    function calcDefaultWidth() {
        clearMeasurementBoundsCache();
        var shapeKey = getShapeType();
        var shapeConfig = SHAPE_CONFIG[shapeKey];
        var inputVal = parseFloat(ui.inputField.text);
        var items = getSortedSelectionItems();
        var i;
        var placement;
        var minGapWidth = null;
        var minTotalHeight = null;
        var defaultWidth;

        if (shapeConfig.enableHeightInput && (isNaN(inputVal) || inputVal <= 0)) return;
        if (items.length < 2) return;

        for (i = 0; i < items.length - 1; i++) {
            placement = computePlacementBetweenItems(items[i], items[i + 1], inputVal, 0, 0, 0);
            if (!placement) {
                continue;
            }

            if (minGapWidth === null || (placement.gapRight - placement.gapLeft) < minGapWidth) {
                minGapWidth = placement.gapRight - placement.gapLeft;
            }
            if (minTotalHeight === null || placement.totalHeight < minTotalHeight) {
                minTotalHeight = placement.totalHeight;
            }
        }

        if (minGapWidth === null || minTotalHeight === null) {
            return;
        }

        defaultWidth = shapeConfig.calcDefaultWidth(inputVal, minTotalHeight, minGapWidth);
        if (defaultWidth === null || typeof defaultWidth === "undefined") {
            ui.widthField.text = "";
        } else {
            setFieldFromPt(ui.widthField, defaultWidth, rulerUnitInfo);
        }
    }

    function getCurrentEffectiveWidthPt() {
        clearMeasurementBoundsCache();
        var shapeKey = getShapeType();
        var shapeConfig = SHAPE_CONFIG[shapeKey];
        var widthVal = parseFieldToPt(ui.widthField, rulerUnitInfo, false);
        var inputVal = parseFloat(ui.inputField.text);
        var items = getSortedSelectionItems();
        var i;
        var placement;
        var minGapWidth = null;
        var minTotalHeight = null;
        var defaultWidth;

        if (!isNaN(widthVal) && widthVal > 0) {
            return widthVal;
        }

        if (!shapeConfig || !shapeConfig.calcDefaultWidth) {
            return 0;
        }

        if (shapeConfig.enableHeightInput && (isNaN(inputVal) || inputVal <= 0)) {
            return 0;
        }

        if (items.length < 2) {
            return 0;
        }

        for (i = 0; i < items.length - 1; i++) {
            placement = computePlacementBetweenItems(items[i], items[i + 1], inputVal, 0, 0, 0);
            if (!placement) {
                continue;
            }

            if (minGapWidth === null || (placement.gapRight - placement.gapLeft) < minGapWidth) {
                minGapWidth = placement.gapRight - placement.gapLeft;
            }
            if (minTotalHeight === null || placement.totalHeight < minTotalHeight) {
                minTotalHeight = placement.totalHeight;
            }
        }

        if (minGapWidth === null || minTotalHeight === null) {
            return 0;
        }

        defaultWidth = shapeConfig.calcDefaultWidth(inputVal, minTotalHeight, minGapWidth);
        return (defaultWidth && defaultWidth > 0) ? defaultWidth : 0;
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
        setFieldFromPt(ui.insetField, shapeConfig.defaultInsetPt || 0, rulerUnitInfo);
        ui.insetField.enabled = !!shapeConfig.enableInset;
        if (ui.insetGroup) {
            ui.insetGroup.visible = true;
            ui.insetGroup.enabled = !!shapeConfig.enableInset;
        }
        setFieldFromPt(ui.strokeField, shapeConfig.defaultStrokePt, strokeUnitInfo);

        ui.capPanel.enabled = !!shapeConfig.enableCapPanel;
        ui.inputGroup.enabled = !!shapeConfig.enableHeightInput;
        ui.strokeField.enabled = !!shapeConfig.enableStrokeInput;

        if (ui.strokeLabel) {
            ui.strokeLabel.text = (shapeKey === "arrow3")
                ? (lang === "ja" ? "線幅×3" : "Stroke ×3")
                : L("labelStroke");
        }

        if (ui.chkAlignVerticalCenter) {
            var isChevronShape = (shapeKey === "chevron" || shapeKey === "chevron2");
            ui.chkAlignVerticalCenter.enabled = isChevronShape;
            if (!isChevronShape) {
                ui.chkAlignVerticalCenter.value = false;
            }
        }

        if (ui.chkMirror) {
            ui.chkMirror.enabled = !!shapeConfig.enableMirror;
            if (!shapeConfig.enableMirror) {
                ui.chkMirror.value = false;
            }
        }

        if (ui.chkRoundCorners) {
            ui.chkRoundCorners.enabled = !!shapeConfig.enableRoundCorners;
            if (!shapeConfig.enableRoundCorners) {
                ui.chkRoundCorners.value = false;
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
    changeValueByArrowKey(ui.insetField);
    changeValueByArrowKey(ui.strokeField);

    ui.insetField.onChanging = function () {
        var insetNum = parseFloat(ui.insetField.text);
        var effectiveWidthPt;
        var maxInsetPt;
        if (!isNaN(insetNum) && insetNum >= 0) {
            effectiveWidthPt = getCurrentEffectiveWidthPt();
            if (effectiveWidthPt > 0) {
                maxInsetPt = effectiveWidthPt * 0.8;
                if (convertValueToPt(insetNum, rulerUnitInfo) > maxInsetPt) {
                    ui.insetField.text = String(roundDisplayValue(convertPtToUnitValue(maxInsetPt, rulerUnitInfo)));
                }
            }
        }
        schedulePreviewUpdate();
    };

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
        if (isNaN(strokeW) || strokeW < 0.25) {
            if (showAlert) {
                alert(L("alertStrokePositive"));
            }
            return false;
        }
        return true;
    }

    function readUIValues(showAlert) {
        var shapeKey = getShapeType();
        var shapeConfig = SHAPE_CONFIG[shapeKey];
        var inputVal = null;
        var widthVal = parseFieldToPt(ui.widthField, rulerUnitInfo, false);
        var insetVal = (shapeConfig && shapeConfig.enableInset)
            ? parseFieldToPt(ui.insetField, rulerUnitInfo, false)
            : 0;
        var offsetX = parseFieldToPt(ui.adjustField, rulerUnitInfo, true);
        var offsetY = parseFieldToPt(ui.adjustVField, rulerUnitInfo, true);
        var strokeW = parseFieldToPt(ui.strokeField, strokeUnitInfo, false);

        if (shapeConfig && shapeConfig.enableInset) {
            if (isNaN(insetVal) || insetVal < 0) {
                if (showAlert) {
                    alert(L("alertInvalidValue"));
                    ui.insetField.active = true;
                }
                return null;
            }
            // clamp to 80% of width when width is explicitly set
            if (widthVal > 0) {
                var maxInsetPt = widthVal * 0.8;
                if (insetVal > maxInsetPt) {
                    insetVal = maxInsetPt;
                }
            }
        }

        if (shapeConfig && shapeConfig.enableHeightInput) {
            inputVal = parseFloat(ui.inputField.text);
            if (isNaN(inputVal) || inputVal <= 0) {
                if (showAlert) {
                    alert(L("alertPositiveNumber"));
                    ui.inputField.active = true;
                }
                return null;
            }
            if (inputVal > 200) {
                if (showAlert) {
                    alert(L("alertMaxHeight"));
                    ui.inputField.active = true;
                }
                return null;
            }
        }

        if (isNaN(widthVal) || widthVal < 0) {
            if (showAlert) {
                alert(L("alertWidthPositive"));
                ui.widthField.active = true;
            }
            return null;
        }

        if (isNaN(offsetX)) {
            if (showAlert) {
                alert(L("alertInvalidValue"));
                ui.adjustField.active = true;
            }
            return null;
        }

        if (isNaN(offsetY)) {
            if (showAlert) {
                alert(L("alertInvalidValue"));
                ui.adjustVField.active = true;
            }
            return null;
        }

        if (isNaN(strokeW) || strokeW < 0.25) {
            strokeW = 0.25;
        }

        if (!validateStrokeWidthForShape(shapeKey, strokeW, showAlert)) {
            if (showAlert) {
                ui.strokeField.active = true;
            }
            return null;
        }

        return {
            shapeKey: shapeKey,
            shapeConfig: shapeConfig,
            inputVal: inputVal,
            widthVal: widthVal,
            insetVal: insetVal,
            offsetX: offsetX,
            offsetY: offsetY,
            strokeW: strokeW
        };
    }



    function getTriAutoCornerRadius(w, h, insetVal) {
        var base = Math.min(w, h);
        var radius = base * 0.12;
        if (insetVal > 0) {
            radius = Math.min(radius, insetVal * 0.45);
        }
        radius = Math.max(1, radius);
        return radius;
    }

    function applyRoundCornersLive(item, radiusPt) {
        if (!item || radiusPt <= 0) return;
        try {
            item.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>');
        } catch (e) { }
    }

    function createTriShape(cx, cy, w, h, insetVal, strokeW, forceFillOnly) {
        var inset = Math.max(0, Math.min(insetVal || 0, w * 0.8));
        var leftX = cx - w / 2;
        var rightX = cx + w / 2;
        var halfH = h / 2;
        var points;
        var path;
        var roundRadius;

        if (inset > 0) {
            points = [
                [rightX, cy],
                [leftX, cy + halfH],
                [leftX + inset, cy],
                [leftX, cy - halfH]
            ];
        } else {
            points = [
                [rightX, cy],
                [leftX, cy + halfH],
                [leftX, cy - halfH]
            ];
        }

        path = doc.activeLayer.pathItems.add();
        path.setEntirePath(points);
        path.closed = true;
        path.filled = true;
        path.fillColor = makeCMYK(0, 0, 0, 100);
        if (!forceFillOnly && strokeW > 0) {
            path.stroked = true;
            path.strokeWidth = strokeW;
            path.strokeColor = makeCMYK(0, 0, 0, 100);
            path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        } else {
            path.stroked = false;
        }

        if (ui.chkRoundCorners && ui.chkRoundCorners.value) {
            roundRadius = getTriAutoCornerRadius(w, h, inset);
            applyRoundCornersLive(path, roundRadius);
        }
        return path;
    }

    function createArrowShape(cx, cy, gapLeft, gapRight, h, widthVal, strokeW) {
        var gapW = gapRight - gapLeft;
        var w = (widthVal > 0) ? widthVal : (gapW * 0.7);
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
        var w = (widthVal > 0) ? widthVal : (gapW * 0.7);
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

    function createDashShape(cx, cy, widthVal, strokeW) {
        var half = widthVal / 2;
        var line = doc.activeLayer.pathItems.add();
        line.setEntirePath([
            [cx - half, cy],
            [cx + half, cy]
        ]);
        line.closed = false;
        line.filled = false;
        line.stroked = true;
        line.strokeWidth = strokeW;
        line.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            line.strokeCap = StrokeCap.ROUNDENDCAP;
        }
        return line;
    }

    function createPlusShape(cx, cy, widthVal, strokeW) {
        var half = widthVal / 2;
        var lineW = strokeW;

        var hLine = doc.activeLayer.pathItems.add();
        hLine.setEntirePath([
            [cx - half, cy],
            [cx + half, cy]
        ]);
        hLine.closed = false;
        hLine.filled = false;
        hLine.stroked = true;
        hLine.strokeWidth = lineW;
        hLine.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            hLine.strokeCap = StrokeCap.ROUNDENDCAP;
        }

        var vLine = doc.activeLayer.pathItems.add();
        vLine.setEntirePath([
            [cx, cy + half],
            [cx, cy - half]
        ]);
        vLine.closed = false;
        vLine.filled = false;
        vLine.stroked = true;
        vLine.strokeWidth = lineW;
        vLine.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            vLine.strokeCap = StrokeCap.ROUNDENDCAP;
        }

        var group = doc.activeLayer.groupItems.add();
        hLine.move(group, ElementPlacement.INSIDE);
        vLine.move(group, ElementPlacement.INSIDE);
        return group;
    }

    function createMultiplyShape(cx, cy, widthVal, strokeW) {
        var half = widthVal / 2;
        var diag = half * Math.SQRT2 / 2;
        var lineW = strokeW;

        var line1 = doc.activeLayer.pathItems.add();
        line1.setEntirePath([
            [cx - diag, cy + diag],
            [cx + diag, cy - diag]
        ]);
        line1.closed = false;
        line1.filled = false;
        line1.stroked = true;
        line1.strokeWidth = lineW;
        line1.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            line1.strokeCap = StrokeCap.ROUNDENDCAP;
        }

        var line2 = doc.activeLayer.pathItems.add();
        line2.setEntirePath([
            [cx - diag, cy - diag],
            [cx + diag, cy + diag]
        ]);
        line2.closed = false;
        line2.filled = false;
        line2.stroked = true;
        line2.strokeWidth = lineW;
        line2.strokeColor = makeCMYK(0, 0, 0, 100);
        if (ui.radioCapRound.value) {
            line2.strokeCap = StrokeCap.ROUNDENDCAP;
        }

        var group = doc.activeLayer.groupItems.add();
        line1.move(group, ElementPlacement.INSIDE);
        line2.move(group, ElementPlacement.INSIDE);
        return group;
    }

    function finalizeArrow3Appearance(parts) {
        var outlinedShaft;
        var result;
        var prevSelection = [];
        var i;
        if (!parts || !parts.shaft || !parts.head) return parts;

        if (doc.selection && doc.selection.length) {
            for (i = 0; i < doc.selection.length; i++) {
                prevSelection.push(doc.selection[i]);
            }
        }

        try {
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
                result = doc.selection[0];
            } else {
                result = outlinedShaft;
            }
            return result;
        } finally {
            try {
                doc.selection = null;
                for (i = 0; i < prevSelection.length; i++) {
                    try {
                        prevSelection[i].selected = true;
                    } catch (e) { }
                }
            } catch (e2) { }
        }
    }

    function computePlacementBetweenItems(leftItem, rightItem, inputVal, widthVal, offsetX, offsetY) {
        var b0 = getItemMeasurementBounds(leftItem);
        var b1 = getItemMeasurementBounds(rightItem);
        var leftObj, rightObj;
        var gapLeft, gapRight;
        var allTop, allBottom;
        var totalHeight;

        if (b0[0] < b1[0]) {
            leftObj = b0;
            rightObj = b1;
        } else {
            leftObj = b1;
            rightObj = b0;
        }

        gapLeft = leftObj[2];
        gapRight = rightObj[0];

        if (gapRight - gapLeft <= 0) {
            return null;
        }

        allTop = Math.max(b0[1], b1[1]);
        allBottom = Math.min(b0[3], b1[3]);
        totalHeight = allTop - allBottom;

        return {
            b0: b0,
            b1: b1,
            gapLeft: gapLeft,
            gapRight: gapRight,
            cx: (gapLeft + gapRight) / 2 + (offsetX || 0),
            cy: (allTop + allBottom) / 2 + (offsetY || 0),
            totalHeight: totalHeight,
            h: (inputVal === null || typeof inputVal === "undefined") ? totalHeight : (totalHeight * (inputVal / 100)),
            w: (widthVal > 0) ? widthVal : ((inputVal === null || typeof inputVal === "undefined") ? totalHeight : (totalHeight * (inputVal / 100)))
        };
    }

    function createShapeAtPlacement(shape, placement, widthVal, insetVal, strokeW, shapeConfig) {
        if (shape === "tri") {
            return createTriShape(placement.cx, placement.cy, placement.w, placement.h, insetVal, strokeW, !!shapeConfig.forceFillOnly);
        }
        if (shape === "arrow") {
            return createArrowShape(placement.cx, placement.cy, placement.gapLeft, placement.gapRight, placement.h, widthVal, strokeW);
        }
        if (shape === "arrow3") {
            var arrow3Parts = createArrow3Shape(placement.cx, placement.cy, placement.gapLeft, placement.gapRight, placement.h, widthVal, strokeW);
            return finalizeArrow3Appearance(arrow3Parts);
        }
        if (shape === "chevron") {
            return createChevronShape(placement.cx, placement.cy, placement.gapLeft, placement.gapRight, placement.h, widthVal, strokeW);
        }
        if (shape === "chevron2") {
            return createChevron2Shape(placement.cx, placement.cy, placement.gapLeft, placement.gapRight, placement.h, widthVal, strokeW);
        }
        if (shape === "dash") {
            return createDashShape(placement.cx, placement.cy, placement.w, strokeW);
        }
        if (shape === "plus") {
            return createPlusShape(placement.cx, placement.cy, placement.w, strokeW);
        }
        if (shape === "multiply") {
            return createMultiplyShape(placement.cx, placement.cy, placement.w, strokeW);
        }
        return createTriShape(placement.cx, placement.cy, placement.w, placement.h, insetVal, strokeW, !!shapeConfig.forceFillOnly);
    }

    function applyMirrorIfNeeded(result) {
        if (!(result && ui.chkMirror && ui.chkMirror.value)) {
            return result;
        }

        try {
            result.resize(
                -100,
                100,
                true,
                true,
                true,
                true,
                100,
                Transformation.CENTER
            );
        } catch (e) {
            var bMirror = result.geometricBounds;
            var mirrorCenterX = (bMirror[0] + bMirror[2]) / 2;
            result.translate(mirrorCenterX * 2, 0);
            result.resize(
                -100,
                100,
                true,
                true,
                true,
                true,
                100,
                Transformation.TOPLEFT
            );
        }

        return result;
    }

    function createShapeBetweenItems(leftItem, rightItem, inputVal, widthVal, insetVal, offsetX, offsetY, strokeW) {
        var shape = getShapeType();
        var shapeConfig = SHAPE_CONFIG[shape];
        var placement = computePlacementBetweenItems(leftItem, rightItem, inputVal, widthVal, offsetX, offsetY);
        var result;

        if (!placement) {
            return null;
        }

        result = createShapeAtPlacement(shape, placement, widthVal, insetVal, strokeW, shapeConfig);
        return applyMirrorIfNeeded(result);
    }

    /* 形状を作成する共通関数 / Shared shape creation function */
    function createShape(inputVal, widthVal, insetVal, offsetX, offsetY, strokeW) {
        clearMeasurementBoundsCache();

        var items = getSortedSelectionItems();
        var createdItems = [];
        var i;
        var shapeItem;

        for (i = 0; i < items.length - 1; i++) {
            shapeItem = createShapeBetweenItems(items[i], items[i + 1], inputVal, widthVal, insetVal, offsetX, offsetY, strokeW);
            if (shapeItem !== null) {
                createdItems.push(shapeItem);
            }
        }

        return createdItems;
    }

    /* プレビューを削除 / Remove preview */
    function removePreview() {
        cancelDebouncedPreview();
        if (!previewPath || previewPath.length === 0) {
            return;
        }

        for (var pi = 0; pi < previewPath.length; pi++) {
            try {
                previewPath[pi].remove();
            } catch (e) {
                try { $.writeln("[RightMarkPlacer] Failed to remove preview item: " + e); } catch (_) { }
            }
        }
        previewPath = [];
        app.redraw();
    }

    /* プレビューデバウンス / Preview debounce */
    var previewDebounceTaskId = null;
    var PREVIEW_DEBOUNCE_DELAY = 120;

    function cancelDebouncedPreview() {
        if (previewDebounceTaskId !== null) {
            try {
                app.cancelTask(previewDebounceTaskId);
            } catch (e) { }
            previewDebounceTaskId = null;
        }
    }

    function schedulePreviewUpdate() {
        cancelDebouncedPreview();
        try {
            previewDebounceTaskId = app.scheduleTask(
                "try { updatePreview(); } catch (e) {}",
                PREVIEW_DEBOUNCE_DELAY,
                false
            );
        } catch (e) {
            updatePreview();
        }
    }

    /* プレビューを更新 / Update preview */
    function updatePreview() {
        previewDebounceTaskId = null;
        var values;
        removePreview();
        if (!ui.previewCb.value) return;

        values = readUIValues(false);
        if (!values) return;

        previewPath = createShape(values.inputVal, values.widthVal, values.insetVal, values.offsetX, values.offsetY, values.strokeW);
        if (!previewPath) {
            previewPath = [];
        }
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

    if (ui.chkMirror) {
        ui.chkMirror.onClick = function () {
            updatePreview();
        };
    }

    if (ui.chkRoundCorners) {
        ui.chkRoundCorners.onClick = function () {
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
        schedulePreviewUpdate();
    };

    ui.widthField.onChanging = function () {
        if (ui.widthField.text === "") {
            resetWidthToAuto();
        } else {
            widthManuallySet = true;
        }
        schedulePreviewUpdate();
    };

    ui.adjustField.onChanging = function () {
        schedulePreviewUpdate();
    };

    ui.adjustVField.onChanging = function () {
        schedulePreviewUpdate();
    };

    ui.strokeField.onChanging = function () {
        schedulePreviewUpdate();
    };

    ui.gapField.onChanging = function () {
        schedulePreviewUpdate();
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
        var values;

        values = readUIValues(true);
        if (!values) {
            return;
        }

        // プレビューがあればそのまま確定
        if (previewPath.length > 0) {
            doc.selection = previewPath;
            previewPath = [];
            dlg.close(1);
            return;
        }

        dlg.close(1);

        var path = createShape(values.inputVal, values.widthVal, values.insetVal, values.offsetX, values.offsetY, values.strokeW);
        if (!path || path.length === 0) {
            alert(L("alertNoGap"));
            return;
        }
        doc.selection = path;
    };

    /* キャンセルボタン処理 / Cancel button handler */
    ui.cancelBtn.onClick = function () {
        cancelDebouncedPreview();
        removePreview();
        dlg.close(0);
    };

    /* 初期値適用 / Apply initial defaults */
    applyShapeDefaults();
    try {
        dlg.layout.layout(true);
        dlg.layout.resize();
    } catch (e) { }

    dlg.show();

})();