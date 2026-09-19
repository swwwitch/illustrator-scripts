#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、指定した縦横比に合わせて拡大・縮小します。

詳細は README を参照してください。

### Overview

Scales the selected objects to a specified aspect ratio.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AspectRatioScaler";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AspectRatioScaler.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4a212e6eacf1"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        // Dialog title / ダイアログタイトル
        dialogTitle: {
            ja: "アスペクト比で調整",
            en: "Adjust by Aspect Ratio"
        },

        // Aspect panel / アスペクト比パネル
        aspectLabel: {
            ja: "アスペクト比",
            en: "Aspect Ratio"
        },
        ratio169: {
            ja: "16:9",
            en: "16:9"
        },
        ratio11: {
            ja: "1:1（スクエア）",
            en: "1:1"
        },
        ratioA4: {
            ja: "A4（1:1.414）",
            en: "1:1.414"
        },
        ratioCustom: {
            ja: "カスタム",
            en: "Custom"
        },

        // Base (orientation) panel / 基準（向き）パネル
        baseLabel: {
            ja: "向き",
            en: "Base"
        },
        baseWidth: {
            ja: "横置き",
            en: "Landscape"
        },
        baseHeight: {
            ja: "縦置き",
            en: "Portrait"
        },

        // Size panel / サイズパネル
        sizePanel: {
            ja: "サイズ",
            en: "Size"
        },
        labelWidth: {
            ja: "横幅",
            en: "Width"
        },
        labelHeight: {
            ja: "高さ",
            en: "Height"
        },

        // Basis panel (horizontal/vertical) / 基準パネル（横／縦）
        basisPanel: {
            ja: "基準",
            en: "Basis"
        },
        basisHorizontal: {
            ja: "横",
            en: "Horizontal"
        },
        basisVertical: {
            ja: "縦",
            en: "Vertical"
        },

        // Options / オプション
        alignToPixelGrid: {
            ja: "ピクセルグリッドに最適化",
            en: "Align to Pixel Grid"
        },
        convertToArtboard: {
            ja: "アートボードに変換",
            en: "Convert to Artboard"
        },

        // Tooltips / ツールチップ
        tipRatioPreset: {
            ja: "よく使う比率です。選ぶとカスタム欄は使いません。",
            en: "Common ratios. Selecting one disables the custom fields."
        },
        tipRatioCustom: {
            ja: "下の欄に好きな比率を入力します。",
            en: "Enter any ratio in the fields below."
        },
        tipCustomWidth: {
            ja: "カスタム比の左側（横）の値です。",
            en: "The left (horizontal) value of the custom ratio."
        },
        tipCustomHeight: {
            ja: "カスタム比の右側（縦）の値です。",
            en: "The right (vertical) value of the custom ratio."
        },
        tipBaseWidth: {
            ja: "長い辺を横にします。",
            en: "Puts the longer side horizontally."
        },
        tipBaseHeight: {
            ja: "長い辺を縦にします。",
            en: "Puts the longer side vertically."
        },
        tipBasisHorizontal: {
            ja: "横幅を保ったまま高さを比率に合わせます。",
            en: "Keeps the width and fits the height to the ratio."
        },
        tipBasisVertical: {
            ja: "高さを保ったまま横幅を比率に合わせます。",
            en: "Keeps the height and fits the width to the ratio."
        },
        tipSizeValue: {
            ja: "基準にする辺の長さです。空欄なら選択範囲の大きさを使います。",
            en: "Length of the side used as the basis. Leave blank to use the size of the selection."
        },
        tipAlignToPixelGrid: {
            ja: "結果の座標と大きさを整数ピクセルに丸めます。",
            en: "Rounds the resulting position and size to whole pixels."
        },
        tipConvertToArtboard: {
            ja: "作った矩形をアートボードに変換します。",
            en: "Converts the resulting rectangle into an artboard."
        },

        // Buttons / ボタン
        run: {
            ja: "実行",
            en: "Apply"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        }
    };

    // Localization helper
    function getLabel(key) {
        try {
            return LABELS[key][uiLang] || "";
        } catch (e) {
            return "";
        }
    }

    // Original sizes for preview/apply (global)
    var __origW = [];
    var __origH = [];

    // 単位コードとラベルのマップ / Unit code to label map
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する。
     * @param {string} prefKey - 環境設定キー
     * @returns {object} code / label / pointsPerUnit を持つオブジェクト
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // 単位に応じた丸め（px=整数、mm=0.1mm刻み、その他=0.01pt刻み） / Unit-aware rounding
    function roundForUnit(valPt) {
        try {
            var unit = app.preferences.getIntegerPreference("rulerType");
            if (unit === 6) { // px -> integer (1px = 1pt assumption)
                return Math.round(valPt);
            }
            if (unit === 1) { // mm -> 0.1mm steps
                var step = (72.0 / 25.4) * 0.1; // 0.1mm in pt
                return Math.round(valPt / step) * step;
            }
        } catch (e) {}
        // default: 0.01pt
        return Math.round(valPt * 100) / 100;
    }

    // 自動横幅の既定値（選択なしのとき）/ Default auto width when no selection
    function getDefaultWidthTextForCurrentUnit() {
        try {
            var unit = app.preferences.getIntegerPreference("rulerType");
            if (unit === 6) { // px
                return "1000";
            }
            if (unit === 1) { // mm
                return "100";
            }
        } catch (e) {}
        return ""; // その他の単位は未指定 / leave empty for other units
    }

    /* ダイアログ作成 / Create dialog */
    function createDialog() {
        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function() {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }

        // UI参照用のローカル変数 / Local variables for UI refs
        var baseWidthRadio, baseHeightRadio;
        var ratio169, ratio11, ratioA4, ratioCustom;
        var editTextWidth, editTextHeight;
        // NEW: Basis radios (UI only; logic to be wired later)
        var basisHorizontalRadio, basisVerticalRadio;

        var offsetX = 300;
        var dialogOpacity = 0.97;
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        setDialogOpacity(dialog, dialogOpacity);
        shiftDialogPosition(dialog, offsetX, 0);
        dialog.alignChildren = "left";

        var topGroup = dialog.add("group");
        topGroup.orientation = "row";
        topGroup.alignChildren = "left";
        topGroup.alignChildren = ["fill", "top"];

        // 2カラム構成：左=アスペクト、右=基準+サイズ
        // 2-column layout: left = Aspect, right = Base + Size
        var leftCol = topGroup.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        var rightCol = topGroup.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["fill", "top"];

        // --- Basis panel (Horizontal / Vertical) --- (UI only; no logic yet)
        var basisPanel = rightCol.add("panel", undefined, LABELS.basisPanel[uiLang]);
        basisPanel.orientation = "column";
        basisPanel.alignChildren = "left";
        basisPanel.margins = [15, 20, 15, 10];
        basisPanel.alignment = ["fill", "top"];

        var basisGroup = basisPanel.add("group");
        basisGroup.orientation = "row";
        basisGroup.alignChildren = "left";
        basisHorizontalRadio = basisGroup.add("radiobutton", undefined, LABELS.basisHorizontal[uiLang]);
        basisHorizontalRadio.helpTip = LABELS.tipBasisHorizontal[uiLang];
        basisVerticalRadio = basisGroup.add("radiobutton", undefined, LABELS.basisVertical[uiLang]);
        basisVerticalRadio.helpTip = LABELS.tipBasisVertical[uiLang];
        basisHorizontalRadio.value = true;
        basisVerticalRadio.value = false;

        var aspectPanel = leftCol.add("panel", undefined, LABELS.aspectLabel[uiLang]);
        aspectPanel.orientation = "column";
        aspectPanel.alignChildren = "left";
        aspectPanel.margins = [15, 20, 15, 10];
        aspectPanel.alignment = ["fill", "top"];

        var aspectGroup = aspectPanel.add("group");
        aspectGroup.orientation = "column";
        aspectGroup.alignChildren = "left";
        ratio169 = aspectGroup.add("radiobutton", undefined, LABELS.ratio169[uiLang]);
        ratio169.helpTip = LABELS.tipRatioPreset[uiLang];
        ratio11 = aspectGroup.add("radiobutton", undefined, LABELS.ratio11[uiLang]);
        ratio11.helpTip = LABELS.tipRatioPreset[uiLang];
        ratioA4 = aspectGroup.add("radiobutton", undefined, LABELS.ratioA4[uiLang]);
        ratioA4.helpTip = LABELS.tipRatioPreset[uiLang];
        ratioCustom = aspectGroup.add("radiobutton", undefined, LABELS.ratioCustom[uiLang]);
        ratioCustom.helpTip = LABELS.tipRatioCustom[uiLang];

        var customRatioGroup = aspectPanel.add("group");
        customRatioGroup.orientation = "row";
        customRatioGroup.alignChildren = "left";

        editTextWidth = customRatioGroup.add("edittext", undefined, "3");
        editTextWidth.helpTip = LABELS.tipCustomWidth[uiLang];
        editTextWidth.characters = 5;

        customRatioGroup.add("statictext", undefined, ":");
        editTextHeight = customRatioGroup.add("edittext", undefined, "2");
        editTextHeight.helpTip = LABELS.tipCustomHeight[uiLang];
        editTextHeight.characters = 5;

        editTextWidth.enabled = false;
        editTextHeight.enabled = false;

        changeValueByArrowKey(editTextWidth);
        changeValueByArrowKey(editTextHeight);

        ratio169.value = true;

        var basePanel = rightCol.add("panel", undefined, LABELS.baseLabel[uiLang]);
        basePanel.orientation = "column";
        basePanel.alignChildren = "left";
        basePanel.margins = [15, 20, 15, 10];
        basePanel.alignment = ["fill", "top"];

        var baseGroup = basePanel.add("group");
        baseGroup.orientation = "column";
        baseGroup.alignChildren = "left";
        baseWidthRadio = baseGroup.add("radiobutton", undefined, LABELS.baseWidth[uiLang]);
        baseWidthRadio.helpTip = LABELS.tipBaseWidth[uiLang];
        baseHeightRadio = baseGroup.add("radiobutton", undefined, LABELS.baseHeight[uiLang]);
        baseHeightRadio.helpTip = LABELS.tipBaseHeight[uiLang];
        baseWidthRadio.value = true; // default Landscape
        baseHeightRadio.value = false;

        // --- Size panel under Base ---
        var sizePanel = rightCol.add("panel", undefined, LABELS.sizePanel[uiLang]);
        sizePanel.orientation = "column";
        sizePanel.alignChildren = "left";
        sizePanel.margins = [15, 20, 15, 10];
        sizePanel.alignment = ["fill", "top"];

        var sizeRow = sizePanel.add("group");
        sizeRow.orientation = "row";
        sizeRow.alignChildren = ["left", "center"];

        var stWidthLabel = sizeRow.add("statictext", undefined, LABELS.labelWidth[uiLang]);
        var etWidthValue = sizeRow.add("edittext", undefined, "");
        etWidthValue.helpTip = LABELS.tipSizeValue[uiLang];
        etWidthValue.characters = 5; // 少し広め / slightly wider
        var stUnitLabel = sizeRow.add("statictext", undefined, getUnitInfo("rulerType").label);

        // 追加: 基準（横/縦）ラジオに応じてラベルを切替
        function updateSizeLabel() {
            // "基準": 縦=高さ固定 / 横=幅固定
            var fixByHeight = false;
            try { fixByHeight = (basisVerticalRadio && basisVerticalRadio.value) ? true : false; } catch (e) { fixByHeight = false; }
            stWidthLabel.text = fixByHeight
                ? (LABELS.labelHeight ? LABELS.labelHeight[uiLang] : "高さ")
                : (LABELS.labelWidth ? LABELS.labelWidth[uiLang] : "横幅");
        }
        updateSizeLabel();

        var pixelGroup = dialog.add("group");
        pixelGroup.orientation = "column";
        pixelGroup.alignChildren = "left";
        var alignToPixel = pixelGroup.add("checkbox", undefined, LABELS.alignToPixelGrid[uiLang]);
        alignToPixel.helpTip = LABELS.tipAlignToPixelGrid[uiLang];
        var isPixelRuler = false;
        try {
            isPixelRuler = (app.preferences.getIntegerPreference("rulerType") === 6);
        } catch (e) {}
        alignToPixel.value = isPixelRuler; // px時のみON、その他はOFF

        var convertToArtboard = pixelGroup.add("checkbox", undefined, LABELS.convertToArtboard[uiLang]);
        convertToArtboard.helpTip = LABELS.tipConvertToArtboard[uiLang];
        convertToArtboard.value = false;

        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "center";

        var btnCancel = buttonGroup.add("button", undefined, LABELS.cancel[uiLang], {
            name: "cancel"
        });
        var btnOk = buttonGroup.add("button", undefined, LABELS.run[uiLang], {
            name: "ok"
        });

        return {
            dialog: dialog,
            ratio169: ratio169,
            ratio11: ratio11,
            ratioA4: ratioA4,
            ratioCustom: ratioCustom,
            baseVertical: baseHeightRadio,
            baseHorizontal: baseWidthRadio,
            alignToPixel: alignToPixel,
            convertToArtboard: convertToArtboard,
            btnOk: btnOk,
            btnCancel: btnCancel,
            customWidthInput: editTextWidth,
            customHeightInput: editTextHeight,
            sizePanel: sizePanel,
            sizeWidthLabel: stWidthLabel,
            sizeWidthInput: etWidthValue,
            sizeUnitLabel: stUnitLabel,
            // NEW: Basis radios (UI only)
            basisHorizontal: basisHorizontalRadio,
            basisVertical: basisVerticalRadio,
            // expose label updater
            updateSizeLabel: updateSizeLabel,
        };
    }

    /* メイン処理 / Main function */
    function main() {
        /* 選択したオブジェクトを取得する / Get selected objects */
        var selectedItems = app.activeDocument.selection;
        var isNoSelection = (!selectedItems || selectedItems.length === 0);

        // プレビュー用コピーを作成し、元は非表示にする / Create preview copies and hide originals (if any)
        var previewCopies = [];
        __origW = [];
        __origH = [];
        if (!isNoSelection) {
            for (var i = 0; i < selectedItems.length; i++) {
                var dup = selectedItems[i].duplicate();
                dup.hidden = false;
                dup.zOrder(ZOrderMethod.BRINGTOFRONT);
                previewCopies.push(dup);
                selectedItems[i].hidden = true;
                __origW.push(selectedItems[i].width); // 幅を保存 / Save width
                __origH.push(selectedItems[i].height); // 高さを保存 / Save height
            }
        }

        var dialogResult = createDialog();

        // 選択がない場合は横幅に自動入力（px:1000 / mm:100）
        // Auto-fill width when no selection (px:1000, mm:100)
        if (isNoSelection && (!dialogResult.sizeWidthInput.text || dialogResult.sizeWidthInput.text === "")) {
            var autoW = getDefaultWidthTextForCurrentUnit();
            if (autoW !== "") dialogResult.sizeWidthInput.text = autoW;
        }

        function getTargetPrimaryPt() {
            var txt = dialogResult.sizeWidthInput.text;
            if (!txt) return null;
            var v = parseFloat(txt);
            if (isNaN(v) || v <= 0) return null;
            return v * getUnitInfo("rulerType").pointsPerUnit;
        }

        // サイズパネル「横幅」入力のライブプレビュー / Live preview for width field
        dialogResult.sizeWidthInput.onChanging = function() {
            applyAspect(previewCopies, getCurrentRatio(), dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };

        // 何も選択されていない場合は、プレビュー用の長方形を新規作成 / Create a preview rectangle when no selection
        if (isNoSelection) {
            // 初期比率と目標幅を取得
            var initR;
            if (dialogResult.ratio169.value) initR = 1.777777;
            else if (dialogResult.ratio11.value) initR = 1.0;
            else if (dialogResult.ratioA4.value) initR = (210 / 297);
            else {
                var _w = parseFloat(dialogResult.customWidthInput.text);
                var _h = parseFloat(dialogResult.customHeightInput.text);
                initR = (isNaN(_w) || isNaN(_h) || _h === 0) ? 1 : _w / _h;
            }
            var wantPortraitInit = dialogResult.baseVertical.value;
            var rAdj = initR;
            if (wantPortraitInit && rAdj > 1) rAdj = 1 / rAdj;
            if (!wantPortraitInit && rAdj < 1) rAdj = 1 / rAdj;

            var targetPrimary = getTargetPrimaryPt();
            if (targetPrimary == null) targetPrimary = 200; // 既定 200pt
            // --- PATCHED: Use basisVertical (高さ固定) for fixed dimension ---
            var fixByHeightInit = dialogResult.basisVertical.value; // 基準：縦=高さ固定
            var targetW, targetH;
            if (fixByHeightInit) {
                targetH = roundForUnit(targetPrimary);
                targetW = targetH * rAdj;
            } else {
                targetW = targetPrimary;
                targetH = roundForUnit(targetW / rAdj);
            }

            var doc = app.activeDocument;
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
            var cx = (ab[0] + ab[2]) / 2;
            var cy = (ab[1] + ab[3]) / 2;
            var left = cx - targetW / 2;
            var top = cy + targetH / 2;

            var rect = doc.pathItems.rectangle(top, left, targetH, targetW);
            rect.stroked = false;
            rect.filled = true;

            previewCopies = [rect];
            __origW = [targetW];
            __origH = [targetH];
        }

        function getCurrentRatio() {
            if (dialogResult.ratio169.value) return 1.777777;
            if (dialogResult.ratio11.value) return 1.0;
            if (dialogResult.ratioA4.value) return (210 / 297);
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            if (isNaN(w) || isNaN(h) || h === 0) return 1;
            return w / h;
        }

        /* アスペクト比選択時のプレビュー更新 / Preview update on aspect ratio selection */
        // 16:9
        dialogResult.ratio169.onClick = function() {
            dialogResult.customWidthInput.enabled = false;
            dialogResult.customHeightInput.enabled = false;
            // ※ 向きは変更しない（ユーザー選択を保持）
            dialogResult.updateSizeLabel();
            applyAspect(previewCopies, 1.777777, dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };

        // 1:1
        dialogResult.ratio11.onClick = function() {
            dialogResult.customWidthInput.enabled = false;
            dialogResult.customHeightInput.enabled = false;
            dialogResult.updateSizeLabel();
            applyAspect(previewCopies, 1.0, dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };

        // A4
        dialogResult.ratioA4.onClick = function() {
            dialogResult.customWidthInput.enabled = false;
            dialogResult.customHeightInput.enabled = false;
            // ※ 向きは変更しない（ユーザー選択を保持）
            dialogResult.updateSizeLabel();
            applyAspect(previewCopies, (210 / 297), dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };

        // カスタム
        dialogResult.ratioCustom.onClick = function() {
            dialogResult.customWidthInput.enabled = true;
            dialogResult.customHeightInput.enabled = true;
            if (dialogResult.ratioCustom.value) {
                var w = parseFloat(dialogResult.customWidthInput.text);
                var h = parseFloat(dialogResult.customHeightInput.text);
                var r = (h === 0) ? 1 : w / h;
                dialogResult.updateSizeLabel();
                applyAspect(previewCopies, r, dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
            }
        };

        /* カスタム比率入力時のプレビュー更新 / Preview update on custom ratio input */
        dialogResult.customWidthInput.onChanging = function() {
            if (dialogResult.ratioCustom.value) {
                var w = parseFloat(dialogResult.customWidthInput.text);
                var h = parseFloat(dialogResult.customHeightInput.text);
                var r = (h === 0) ? 1 : w / h;
                dialogResult.updateSizeLabel();
                applyAspect(previewCopies, r, dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());

            }
        };
        dialogResult.customHeightInput.onChanging = function() {
            if (dialogResult.ratioCustom.value) {
                var w = parseFloat(dialogResult.customWidthInput.text);
                var h = parseFloat(dialogResult.customHeightInput.text);
                var r = (h === 0) ? 1 : w / h;
                dialogResult.updateSizeLabel();
                applyAspect(previewCopies, r, dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
            }
        };

        // 基準ラジオ（横/縦）: ラベル更新＆プレビュー再計算（現状ロジックは向きベース）
        dialogResult.basisHorizontal.onClick = function () {
            dialogResult.updateSizeLabel();
            applyAspect(previewCopies, getCurrentRatio(), dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };
        dialogResult.basisVertical.onClick = function () {
            dialogResult.updateSizeLabel();
            applyAspect(previewCopies, getCurrentRatio(), dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };

        dialogResult.baseVertical.onClick = function() {
            dialogResult.updateSizeLabel();
            applyAspect(previewCopies, getCurrentRatio(), true, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };
        dialogResult.baseHorizontal.onClick = function() {
            dialogResult.updateSizeLabel();
            applyAspect(previewCopies, getCurrentRatio(), false, dialogResult.basisVertical.value, getTargetPrimaryPt());
        };

        /* 初期プレビュー / Initial preview */
        var initialRatio = dialogResult.ratio169.value ? 1.777777 : (dialogResult.ratio11.value ? 1.0 : (function() {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            return (h === 0) ? 1 : w / h;
        })());
        applyAspect(previewCopies, initialRatio, dialogResult.baseVertical.value, dialogResult.basisVertical.value, getTargetPrimaryPt());

        var result = dialogResult.dialog.show();

        if (result === 1) {
            if (isNoSelection) {
                // 新規作成したプレビュー矩形を最終物として扱う / Keep the preview rectangle as final
                var finalIt = previewCopies[0];
                // ピクセルグリッド整合 / Align to pixel grid (optional)
                if (dialogResult.alignToPixel.value) {
                    app.selection = [finalIt];
                    app.executeMenuCommand('Make Pixel Perfect');
                }
                // 必要に応じてアートボードを作成 / Convert to artboard if requested
                if (dialogResult.convertToArtboard.value) {
                    var vb0 = finalIt.visibleBounds;
                    var abRect0 = [vb0[0], vb0[1], vb0[2], vb0[3]];
                    app.activeDocument.artboards.add(abRect0);
                }
                app.selection = [finalIt];
                app.redraw();
            } else {
                for (var i = 0; i < selectedItems.length; i++) {
                    // 元オブジェクトを再表示 / Unhide original
                    selectedItems[i].hidden = false;

                    // プレビューの形状を反映（拡大縮小＋位置合わせ） / Apply by scaling and repositioning
                    try {
                        var origW = __origW[i];
                        var origH = __origH[i];
                        var prevW = previewCopies[i].width;
                        var prevH = previewCopies[i].height;

                        // 安全ガード / guards
                        if (origW > 0 && origH > 0) {
                            var sx = (prevW / origW) * 100.0;
                            var sy = (prevH / origH) * 100.0;
                            // 中心基準で拡大縮小 / scale from center
                            selectedItems[i].resize(sx, sy);
                        }

                        // 位置合わせ（左上座標） / align position using top-left
                        try {
                            selectedItems[i].position = previewCopies[i].position;
                        } catch (pErr) {}
                    } catch (e) {}

                    // ピクセルグリッド整合 / Align to pixel grid (optional)
                    if (dialogResult.alignToPixel.value) {
                        app.selection = [selectedItems[i]];
                        app.executeMenuCommand('Make Pixel Perfect');
                    }

                    // プレビューを削除 / Remove preview copy
                    try {
                        previewCopies[i].remove();
                    } catch (e3) {}
                }

                // 必要に応じてアートボードを作成 / Convert to artboard if requested
                if (dialogResult.convertToArtboard.value) {
                    for (var j = 0; j < selectedItems.length; j++) {
                        var it = selectedItems[j];
                        var vb = it.visibleBounds;
                        var abRect = [vb[0], vb[1], vb[2], vb[3]];
                        app.activeDocument.artboards.add(abRect);
                    }
                }

                // 最終的にオリジナルを選択状態に / Keep originals selected
                app.selection = selectedItems;
                app.redraw();
            }
        } else {
            // キャンセル時：プレビューを片付け、非表示化を解除 / On cancel, cleanup preview and unhide originals
            for (var i = 0; i < previewCopies.length; i++) {
                try {
                    previewCopies[i].remove();
                } catch (e) {}
            }
            if (!isNoSelection) {
                for (var i = 0; i < selectedItems.length; i++) {
                    selectedItems[i].hidden = false;
                }
            }
            return;
        }
    }

    /* アスペクト比適用 / Apply aspect ratio */
    function applyAspect(items, ratio, wantPortrait, fixByHeight, targetPrimaryPt) {
        // 向きのガード（比率の向きのみ補正。固定寸法は基準ラジオで判定）
        var r = ratio;
        if (wantPortrait && r > 1) r = 1 / r;
        if (!wantPortrait && r < 1) r = 1 / r;

        var useTarget = (typeof targetPrimaryPt === 'number' && isFinite(targetPrimaryPt) && targetPrimaryPt > 0);
        for (var i = 0; i < items.length; i++) {
            if (fixByHeight) {
                // 高さ固定（基準：縦） => width = height * r
                var h = useTarget ? targetPrimaryPt : __origH[i];
                items[i].height = roundForUnit(h);
                var w = h * r;
                items[i].width = w;
            } else {
                // 幅固定（基準：横） => height = width / r
                var w2 = useTarget ? targetPrimaryPt : __origW[i];
                items[i].width = w2;
                var h2 = w2 / r;
                items[i].height = roundForUnit(h2);
            }
        }

        app.redraw();
    }

    /* 上下キーで数値変更を可能にする / Enable arrow key numeric input */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta;

            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
            // 整数に丸め / Round to integer
            value = Math.round(value);

            editText.text = value;
            if (typeof editText.onChanging === 'function') {
                try {
                    editText.onChanging();
                } catch (e) {}
            }
        });
    }

    main();

})();
