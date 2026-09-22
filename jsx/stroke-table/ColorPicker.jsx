#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

他のスクリプトから `#include` して使う、カラーピッカーの再利用ライブラリです。
`ColorPicker.show()` を呼ぶとダイアログを開き、選択された色を返します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorPicker.md

### Overview

A reusable color-picker library meant to be pulled in from other scripts with `#include`.
Calling `ColorPicker.show()` opens the dialog and returns the chosen color.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorPicker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ColorPicker";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorPicker.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorPicker.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* スウォッチ行に並べる色（RRGGBB） / Colors shown in the swatch row (RRGGBB) */
    var DEFAULT_SWATCHES = [
        "FF0000", "FFCC00", "FFFF00",
        "00CC00", "0066FF",
        "FF99CC", "996633", "666666",
        "999999"
    ];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 前回閉じたときのダイアログ位置（同じセッション内で再利用） / Dialog location from the last close */
    var lastDialogLocation = null;

    /* 表示言語。show() の lang オプションで決まる / UI language, set from the lang option of show() */
    var uiLang = "en";

    var LABELS = {
        radio: {
            white: { ja: "ホワイト", en: "White" },
            black: { ja: "ブラック", en: "Black" },
            custom: { ja: "カスタム", en: "Custom" }
        },
        checkbox: {
            gray: { ja: "グレー", en: "Gray" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            slider: {
                ja: "この成分の値をドラッグで決めます。右の欄に直接入力もできます。",
                en: "Drag to set this component. You can also type into the field on the right."
            },
            preset: {
                ja: "よく使う色をすぐ選べます。「カスタム」で自由に指定できます。",
                en: "Picks a common color. Custom lets you set any value."
            },
            hex: { ja: "色を16進数で指定します（例: FF0000）。", en: "Color as a hex value, for example FF0000." },
            gray: { ja: "CMYKのK版だけで色を作ります（グレースケール）。", en: "Builds the color from the K plate only, giving a grayscale." },
            swatch: { ja: "クリックすると、この色をRGBで設定します。", en: "Click to set this color as RGB." },
            previewBefore: { ja: "元の色です。", en: "The original color." },
            previewAfter: { ja: "現在指定している色です。", en: "The color currently set." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "radio.white" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en;
    }

    // =========================================
    // 色の変換 / Color conversion
    // =========================================

    /**
     * CMYK の色文字列（"cmyk:C,M,Y,K"）かどうかを返す
     * @param {string} colorString - 色文字列
     * @returns {boolean} CMYK の色文字列なら true
     */
    function isCmykString(colorString) {
        return String(colorString).indexOf("cmyk:") === 0;
    }

    /**
     * "cmyk:C,M,Y,K" を成分ごとの数値に分ける（数値でない成分は 0）
     * @param {string} colorString - CMYK の色文字列
     * @returns {{c: number, m: number, y: number, k: number}} CMYK の各成分
     */
    function parseCmykString(colorString) {
        var cmykParts = String(colorString).replace("cmyk:", "").split(",");
        return {
            c: Number(cmykParts[0]) || 0,
            m: Number(cmykParts[1]) || 0,
            y: Number(cmykParts[2]) || 0,
            k: Number(cmykParts[3]) || 0
        };
    }

    /**
     * CMYK の各成分を丸めて "cmyk:C,M,Y,K" にする
     * @param {number} c - シアン（0〜100）
     * @param {number} m - マゼンタ（0〜100）
     * @param {number} y - イエロー（0〜100）
     * @param {number} k - ブラック（0〜100）
     * @returns {string} CMYK の色文字列
     */
    function cmykStringFromValues(c, m, y, k) {
        return "cmyk:" + Math.round(c) + "," + Math.round(m) + "," + Math.round(y) + "," + Math.round(k);
    }

    /**
     * RGB の各成分を "RRGGBB"（大文字の16進数）にする
     * @param {number} r - レッド（0〜255）
     * @param {number} g - グリーン（0〜255）
     * @param {number} b - ブルー（0〜255）
     * @returns {string} 16進数の色文字列
     */
    function rgbToHex(r, g, b) {
        /**
         * 0〜255 の値を2桁の16進数にする
         * @param {number} channelValue - 成分の値
         * @returns {string} 2桁の16進数（大文字）
         */
        function toHexByte(channelValue) {
            var hexText = Math.round(channelValue).toString(16).toUpperCase();
            return hexText.length < 2 ? "0" + hexText : hexText;
        }
        return toHexByte(r) + toHexByte(g) + toHexByte(b);
    }

    /**
     * "RRGGBB"（先頭の # は可）を RGB の各成分にする。6桁でなければ黒
     * @param {string} hex - 16進数の色文字列
     * @returns {{r: number, g: number, b: number}} RGB の各成分
     */
    function hexToRGB(hex) {
        hex = String(hex || "").replace(/^#/, "");
        if (hex.length !== 6) hex = "000000";
        return {
            r: parseInt(hex.substring(0, 2), 16) || 0,
            g: parseInt(hex.substring(2, 4), 16) || 0,
            b: parseInt(hex.substring(4, 6), 16) || 0
        };
    }

    /**
     * CMYK を RGB に簡易変換する（カラープロファイルは使わない）
     * @param {number} c - シアン（0〜100）
     * @param {number} m - マゼンタ（0〜100）
     * @param {number} y - イエロー（0〜100）
     * @param {number} k - ブラック（0〜100）
     * @returns {{r: number, g: number, b: number}} RGB の各成分（整数）
     */
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

    /**
     * RGB を CMYK に簡易変換する（カラープロファイルは使わない）
     * @param {number} r - レッド（0〜255）
     * @param {number} g - グリーン（0〜255）
     * @param {number} b - ブルー（0〜255）
     * @returns {{c: number, m: number, y: number, k: number}} CMYK の各成分（整数）
     */
    function rgbToCmykApprox(r, g, b) {
        var redRatio = r / 255;
        var greenRatio = g / 255;
        var blueRatio = b / 255;
        var k = 1 - Math.max(redRatio, greenRatio, blueRatio);

        if (k >= 1) {
            return { c: 0, m: 0, y: 0, k: 100 };
        }

        var c = (1 - redRatio - k) / (1 - k) * 100;
        var m = (1 - greenRatio - k) / (1 - k) * 100;
        var y = (1 - blueRatio - k) / (1 - k) * 100;

        return {
            c: Math.round(c),
            m: Math.round(m),
            y: Math.round(y),
            k: Math.round(k * 100)
        };
    }

    /**
     * 値を数値にして範囲内に収める（数値でなければ 0）
     * @param {*} value - 元の値（入力欄の文字列など）
     * @param {number} min - 下限
     * @param {number} max - 上限
     * @returns {number} 範囲内の数値
     */
    function clamp(value, min, max) {
        var clampedValue = Number(value);
        if (isNaN(clampedValue)) clampedValue = 0;
        if (clampedValue < min) clampedValue = min;
        if (clampedValue > max) clampedValue = max;
        return clampedValue;
    }

    // =========================================
    // ピッカーの状態 / Picker state
    // =========================================

    /**
     * 初期値の色文字列からピッカーの状態を作る
     * @param {string} initialValue - "RRGGBB" または "cmyk:C,M,Y,K"
     * @returns {Object} ピッカーの状態（preset / mode / dialogTab / rgb / cmyk / original）
     */
    function createInitialState(initialValue) {
        var state = {
            preset: "custom",   /* white | black | custom */
            mode: "rgb",        /* rgb | cmyk | gray */
            dialogTab: "rgb",   /* rgb | cmyk */
            rgb: { r: 0, g: 0, b: 0 },
            cmyk: { c: 0, m: 0, y: 0, k: 0 },
            original: { r: 0, g: 0, b: 0 }
        };

        if (isCmykString(initialValue)) {
            var cmykValues = parseCmykString(initialValue);
            var convertedRgb = cmykToRgbApprox(cmykValues.c, cmykValues.m, cmykValues.y, cmykValues.k);
            state.cmyk = { c: cmykValues.c, m: cmykValues.m, y: cmykValues.y, k: cmykValues.k };
            state.rgb = { r: convertedRgb.r, g: convertedRgb.g, b: convertedRgb.b };
            state.original = { r: convertedRgb.r, g: convertedRgb.g, b: convertedRgb.b };
            state.dialogTab = "cmyk";
            state.mode = (cmykValues.c === 0 && cmykValues.m === 0 && cmykValues.y === 0 && cmykValues.k > 0) ? "gray" : "cmyk";

            if (cmykValues.c === 0 && cmykValues.m === 0 && cmykValues.y === 0 && cmykValues.k === 0) state.preset = "white";
            else if (cmykValues.c === 0 && cmykValues.m === 0 && cmykValues.y === 0 && cmykValues.k === 100) state.preset = "black";
        } else {
            var initialRgb = hexToRGB(initialValue || "000000");
            state.rgb = { r: initialRgb.r, g: initialRgb.g, b: initialRgb.b };
            state.original = { r: initialRgb.r, g: initialRgb.g, b: initialRgb.b };
            state.cmyk = rgbToCmykApprox(initialRgb.r, initialRgb.g, initialRgb.b);
            state.dialogTab = "rgb";
            state.mode = "rgb";

            if (initialRgb.r === 255 && initialRgb.g === 255 && initialRgb.b === 255) state.preset = "white";
            else if (initialRgb.r === 0 && initialRgb.g === 0 && initialRgb.b === 0) state.preset = "black";
        }

        return state;
    }

    /**
     * RGB の値から CMYK の値を計算し直す
     * @param {Object} state - ピッカーの状態
     * @returns {void}
     */
    function syncCmykFromRgb(state) {
        state.cmyk = rgbToCmykApprox(
            Math.round(state.rgb.r),
            Math.round(state.rgb.g),
            Math.round(state.rgb.b)
        );
    }

    /**
     * CMYK の値から RGB の値を計算し直す
     * @param {Object} state - ピッカーの状態
     * @returns {void}
     */
    function syncRgbFromCmyk(state) {
        var convertedRgb = cmykToRgbApprox(state.cmyk.c, state.cmyk.m, state.cmyk.y, state.cmyk.k);
        state.rgb = { r: convertedRgb.r, g: convertedRgb.g, b: convertedRgb.b };
    }

    /**
     * プレビューに塗る RGB を返す
     * @param {Object} state - ピッカーの状態
     * @returns {{r: number, g: number, b: number}} プレビューの色
     */
    function getPreviewRgb(state) {
        if (state.preset === "white") return { r: 255, g: 255, b: 255 };
        if (state.preset === "black") return { r: 0, g: 0, b: 0 };
        if (state.mode === "cmyk" || state.mode === "gray") {
            return cmykToRgbApprox(state.cmyk.c, state.cmyk.m, state.cmyk.y, state.cmyk.k);
        }
        return { r: state.rgb.r, g: state.rgb.g, b: state.rgb.b };
    }

    /**
     * ピッカーの状態を戻り値の色文字列にする
     * @param {Object} state - ピッカーの状態
     * @returns {string} "RRGGBB" または "cmyk:C,M,Y,K"
     */
    function serializeState(state) {
        if (state.preset === "white") return "FFFFFF";
        if (state.preset === "black") return "000000";
        if (state.mode === "cmyk" || state.mode === "gray") {
            return cmykStringFromValues(state.cmyk.c, state.cmyk.m, state.cmyk.y, state.cmyk.k);
        }
        return rgbToHex(state.rgb.r, state.rgb.g, state.rgb.b);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * コントロール全体を RGB の色で塗る（onDraw の中で呼ぶ）
     * @param {Object} targetControl - 塗るコントロール
     * @param {{r: number, g: number, b: number}} rgb - 塗る色
     * @returns {void}
     */
    function fillControlWithRgb(targetControl, rgb) {
        var controlGraphics = targetControl.graphics;
        var fillBrush = controlGraphics.newBrush(controlGraphics.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
        controlGraphics.rectPath(0, 0, targetControl.size[0], targetControl.size[1]);
        controlGraphics.fillPath(fillBrush);
    }

    /**
     * ラベル＋スライダー＋数値欄の1行を追加する（連動は bindPickerEvents() で付ける）
     * @param {Group|Tab} parentContainer - 追加先
     * @param {string} channelName - 成分名（"R" など）
     * @param {number} initialValue - 初期値
     * @param {number} maxValue - 上限値
     * @returns {{row: Group, slider: Slider, valueInput: EditText}} 生成した行と部品
     */
    function addChannelRow(parentContainer, channelName, initialValue, maxValue) {
        var channelRow = parentContainer.add("group");
        channelRow.orientation = "row";
        channelRow.alignChildren = ["left", "center"];

        var channelLabel = channelRow.add("statictext", undefined, channelName);
        channelLabel.preferredSize = [18, -1];

        var channelSlider = channelRow.add("slider", undefined, initialValue, 0, maxValue);
        channelSlider.helpTip = getLabel("tooltip.slider");
        channelSlider.preferredSize = [140, 20];

        var valueInput = channelRow.add("edittext", undefined, String(Math.round(initialValue)));
        valueInput.helpTip = getLabel("tooltip.slider");
        valueInput.characters = 3;

        return {
            row: channelRow,
            slider: channelSlider,
            valueInput: valueInput
        };
    }

    /**
     * 元の色と現在の色を並べるプレビュー行を追加する
     * @param {Window} pickerDialog - ダイアログ
     * @param {Object} state - ピッカーの状態
     * @returns {Group} 現在の色のプレビュー（再描画用）
     */
    function addPreviewRow(pickerDialog, state) {
        var previewRow = pickerDialog.add("group");
        previewRow.orientation = "row";
        previewRow.alignment = ["center", "top"];
        previewRow.spacing = 1;

        var previewBefore = previewRow.add("group");
        previewBefore.preferredSize = [90, 40];
        previewBefore.helpTip = getLabel("tooltip.previewBefore");

        var previewAfter = previewRow.add("group");
        previewAfter.preferredSize = [90, 40];
        previewAfter.helpTip = getLabel("tooltip.previewAfter");

        previewBefore.onDraw = function () {
            fillControlWithRgb(this, state.original);
        };

        previewAfter.onDraw = function () {
            fillControlWithRgb(this, getPreviewRgb(state));
        };
        return previewAfter;
    }

    /**
     * よく使う色のスウォッチ行を追加する
     * @param {Window} pickerDialog - ダイアログ
     * @returns {Object[]} スウォッチ（element: Group、hex: 色）の配列
     */
    function addSwatchRow(pickerDialog) {
        var swatchRow = pickerDialog.add("group");
        swatchRow.orientation = "row";
        swatchRow.alignment = ["center", "top"];
        swatchRow.spacing = 2;
        var swatchItems = [];
        for (var i = 0; i < DEFAULT_SWATCHES.length; i++) {
            (function (swatchHex) {
                var swatchRgb = hexToRGB(swatchHex);
                var swatchGroup = swatchRow.add("group");
                swatchGroup.preferredSize = [16, 16];
                swatchGroup.helpTip = getLabel("tooltip.swatch");
                swatchGroup.onDraw = function () {
                    var swatchGraphics = this.graphics;
                    var borderPen = swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
                    fillControlWithRgb(this, swatchRgb);
                    swatchGraphics.rectPath(0, 0, this.size[0], this.size[1]);
                    swatchGraphics.strokePath(borderPen);
                };
                swatchItems.push({ element: swatchGroup, hex: swatchHex });
            })(DEFAULT_SWATCHES[i]);
        }
        return swatchItems;
    }

    /**
     * ピッカーのダイアログを組み立てる
     * @param {Object} state - ピッカーの状態
     * @param {string} dialogTitle - ダイアログのタイトル
     * @returns {Object} ダイアログと各コントロールの参照
     */
    function buildPickerDialog(state, dialogTitle) {
        var pickerDialog = new Window("dialog", dialogTitle || "Color Picker");
        pickerDialog.orientation = "column";
        pickerDialog.alignChildren = ["fill", "top"];
        pickerDialog.margins = 14;

        var previewAfter = addPreviewRow(pickerDialog, state);

        var presetRow = pickerDialog.add("group");
        presetRow.orientation = "row";
        presetRow.alignment = ["center", "top"];
        presetRow.alignChildren = ["left", "center"];
        var whiteRadio = presetRow.add("radiobutton", undefined, getLabel("radio.white"));
        whiteRadio.helpTip = getLabel("tooltip.preset");
        var blackRadio = presetRow.add("radiobutton", undefined, getLabel("radio.black"));
        blackRadio.helpTip = getLabel("tooltip.preset");
        var customRadio = presetRow.add("radiobutton", undefined, getLabel("radio.custom"));
        customRadio.helpTip = getLabel("tooltip.preset");

        var swatchItems = addSwatchRow(pickerDialog);

        var colorTabs = pickerDialog.add("tabbedpanel");
        colorTabs.alignChildren = ["fill", "top"];

        var tabRGB = colorTabs.add("tab", undefined, "RGB");
        tabRGB.orientation = "column";
        tabRGB.margins = [14, 18, 14, 10];

        var tabCMYK = colorTabs.add("tab", undefined, "CMYK");
        tabCMYK.orientation = "column";
        tabCMYK.margins = [14, 18, 14, 10];

        var redRow = addChannelRow(tabRGB, "R", state.rgb.r, 255);
        var greenRow = addChannelRow(tabRGB, "G", state.rgb.g, 255);
        var blueRow = addChannelRow(tabRGB, "B", state.rgb.b, 255);

        tabRGB.add("panel").preferredSize.height = 10; /* 区切り線 / divider */

        var hexRow = tabRGB.add("group");
        hexRow.orientation = "row";
        hexRow.add("statictext", undefined, "#");
        var hexInput = hexRow.add("edittext", undefined, rgbToHex(state.rgb.r, state.rgb.g, state.rgb.b));
        hexInput.helpTip = getLabel("tooltip.hex");
        hexInput.characters = 6;

        var grayCheckbox = tabCMYK.add("checkbox", undefined, getLabel("checkbox.gray"));
        grayCheckbox.helpTip = getLabel("tooltip.gray");
        var cyanRow = addChannelRow(tabCMYK, "C", state.cmyk.c, 100);
        var magentaRow = addChannelRow(tabCMYK, "M", state.cmyk.m, 100);
        var yellowRow = addChannelRow(tabCMYK, "Y", state.cmyk.y, 100);
        var blackRow = addChannelRow(tabCMYK, "K", state.cmyk.k, 100);

        var btnRowGroup = pickerDialog.add("group");
        btnRowGroup.alignment = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            dialog: pickerDialog,
            previewAfter: previewAfter,
            whiteRadio: whiteRadio,
            blackRadio: blackRadio,
            customRadio: customRadio,
            colorTabs: colorTabs,
            tabRGB: tabRGB,
            tabCMYK: tabCMYK,
            grayCheckbox: grayCheckbox,
            redRow: redRow,
            greenRow: greenRow,
            blueRow: blueRow,
            cyanRow: cyanRow,
            magentaRow: magentaRow,
            yellowRow: yellowRow,
            blackRow: blackRow,
            hexInput: hexInput,
            swatchItems: swatchItems
        };
    }

    /**
     * 成分の行のスライダーと数値欄に値を入れる
     * @param {Object} channelRow - addChannelRow() の戻り値
     * @param {number} channelValue - 値
     * @returns {void}
     */
    function setChannelRowValue(channelRow, channelValue) {
        channelRow.slider.value = channelValue;
        channelRow.valueInput.text = String(channelValue);
    }

    /**
     * ピッカーの状態をダイアログに反映する
     * @param {Object} state - ピッカーの状態
     * @param {Object} pickerControls - buildPickerDialog() の戻り値
     * @param {Object} [renderOptions] - suppressTabSelection: true でタブを切り替えない
     * @returns {void}
     */
    function renderPickerState(state, pickerControls, renderOptions) {
        renderOptions = renderOptions || {};
        var suppressTabSelection = !!renderOptions.suppressTabSelection;

        pickerControls.whiteRadio.value = (state.preset === "white");
        pickerControls.blackRadio.value = (state.preset === "black");
        pickerControls.customRadio.value = (state.preset === "custom");

        if (!suppressTabSelection) {
            var targetTab = (state.dialogTab === "cmyk") ? pickerControls.tabCMYK : pickerControls.tabRGB;
            /* タブの切り替えに失敗しても表示の更新は続ける / keep rendering even if the tab switch fails */
            try {
                if (pickerControls.colorTabs.selection !== targetTab) {
                    pickerControls.colorTabs.selection = targetTab;
                }
            } catch (eTab) {}
        }

        pickerControls.grayCheckbox.value = (state.mode === "gray");

        setChannelRowValue(pickerControls.redRow, state.rgb.r);
        setChannelRowValue(pickerControls.greenRow, state.rgb.g);
        setChannelRowValue(pickerControls.blueRow, state.rgb.b);
        pickerControls.hexInput.text = rgbToHex(state.rgb.r, state.rgb.g, state.rgb.b);

        setChannelRowValue(pickerControls.cyanRow, state.cmyk.c);
        setChannelRowValue(pickerControls.magentaRow, state.cmyk.m);
        setChannelRowValue(pickerControls.yellowRow, state.cmyk.y);
        setChannelRowValue(pickerControls.blackRow, state.cmyk.k);

        var customEnabled = (state.preset === "custom");
        var colorChannelsEnabled = customEnabled && state.mode !== "gray";
        pickerControls.colorTabs.enabled = customEnabled;
        pickerControls.cyanRow.row.enabled = colorChannelsEnabled;
        pickerControls.magentaRow.row.enabled = colorChannelsEnabled;
        pickerControls.yellowRow.row.enabled = colorChannelsEnabled;

        /* 隠して出し直すことでプレビューを再描画させる / hide and show to force a redraw of the preview */
        try { pickerControls.previewAfter.hide(); pickerControls.previewAfter.show(); } catch (ePreview) {}
    }

    /**
     * ダイアログの各コントロールに、状態の更新と再描画のイベントを付ける
     * @param {Object} state - ピッカーの状態
     * @param {Object} pickerControls - buildPickerDialog() の戻り値
     * @returns {void}
     */
    function bindPickerEvents(state, pickerControls) {
        var isRendering = false;

        /**
         * 表示を更新する（更新中に呼ばれた onChange などからの再入は無視する）
         * @param {Object} [renderOptions] - renderPickerState() に渡すオプション
         * @returns {void}
         */
        function safeRender(renderOptions) {
            if (isRendering) return;
            isRendering = true;
            try {
                renderPickerState(state, pickerControls, renderOptions);
            } finally {
                isRendering = false;
            }
        }

        /**
         * ホワイト／ブラック／カスタムを切り替え、色の値をそろえる
         * @param {string} nextPreset - "white" / "black" / "custom"
         * @returns {void}
         */
        function applyPreset(nextPreset) {
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

        /**
         * RGB の値を設定し、CMYK も計算し直して「カスタム」の RGB 指定にする
         * @param {number} r - レッド（0〜255）
         * @param {number} g - グリーン（0〜255）
         * @param {number} b - ブルー（0〜255）
         * @returns {void}
         */
        function setRgb(r, g, b) {
            state.rgb.r = Math.round(clamp(r, 0, 255));
            state.rgb.g = Math.round(clamp(g, 0, 255));
            state.rgb.b = Math.round(clamp(b, 0, 255));
            syncCmykFromRgb(state);
            state.mode = "rgb";
            state.dialogTab = "rgb";
            state.preset = "custom";
        }

        /**
         * CMYK の値を設定し、RGB も計算し直して「カスタム」の CMYK 指定にする
         * @param {number} c - シアン（0〜100）
         * @param {number} m - マゼンタ（0〜100）
         * @param {number} y - イエロー（0〜100）
         * @param {number} k - ブラック（0〜100）
         * @param {boolean} keepGray - グレー（K のみ）のままにするなら true
         * @returns {void}
         */
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

        /**
         * RGB のスライダーの値を状態に入れる
         * @returns {void}
         */
        function applyRgbSliders() {
            setRgb(pickerControls.redRow.slider.value, pickerControls.greenRow.slider.value, pickerControls.blueRow.slider.value);
        }

        /**
         * CMYK のスライダーの値を状態に入れる
         * @returns {void}
         */
        function applyCmykSliders() {
            setCmyk(pickerControls.cyanRow.slider.value, pickerControls.magentaRow.slider.value,
                pickerControls.yellowRow.slider.value, pickerControls.blackRow.slider.value, pickerControls.grayCheckbox.value);
        }

        /**
         * スライダーと数値欄を連動させ、変更のたびに状態を更新する
         * @param {Object} channelRow - addChannelRow() の戻り値
         * @param {number} maxValue - 上限値
         * @param {Function} applyChannelValues - スライダーの値を状態に入れる関数
         * @returns {void}
         */
        function bindChannelRow(channelRow, maxValue, applyChannelValues) {
            channelRow.slider.onChanging = function () {
                channelRow.valueInput.text = String(Math.round(channelRow.slider.value));
                applyChannelValues();
                safeRender();
            };
            channelRow.valueInput.onChange = function () {
                channelRow.slider.value = clamp(channelRow.valueInput.text, 0, maxValue);
                channelRow.valueInput.text = String(Math.round(channelRow.slider.value));
                applyChannelValues();
                safeRender();
            };
        }

        var rgbRows = [pickerControls.redRow, pickerControls.greenRow, pickerControls.blueRow];
        for (var i = 0; i < rgbRows.length; i++) {
            bindChannelRow(rgbRows[i], 255, applyRgbSliders);
        }

        var cmykRows = [pickerControls.cyanRow, pickerControls.magentaRow, pickerControls.yellowRow, pickerControls.blackRow];
        for (var j = 0; j < cmykRows.length; j++) {
            bindChannelRow(cmykRows[j], 100, applyCmykSliders);
        }

        pickerControls.hexInput.onChange = function () {
            var hexText = String(pickerControls.hexInput.text).replace(/^#/, "");
            if (hexText.length !== 6) return;
            var enteredRgb = hexToRGB(hexText);
            setRgb(enteredRgb.r, enteredRgb.g, enteredRgb.b);
            safeRender();
        };

        pickerControls.whiteRadio.onClick = function () {
            applyPreset("white");
            safeRender();
        };

        pickerControls.blackRadio.onClick = function () {
            applyPreset("black");
            safeRender();
        };

        pickerControls.customRadio.onClick = function () {
            applyPreset("custom");
            safeRender();
        };

        pickerControls.colorTabs.onChange = function () {
            if (isRendering) return;
            state.dialogTab = (pickerControls.colorTabs.selection === pickerControls.tabCMYK) ? "cmyk" : "rgb";
            if (state.dialogTab === "rgb") {
                syncRgbFromCmyk(state);
                state.mode = "rgb";
            } else {
                syncCmykFromRgb(state);
                state.mode = pickerControls.grayCheckbox.value ? "gray" : "cmyk";
            }
            safeRender({ suppressTabSelection: true });
        };

        for (var k = 0; k < pickerControls.swatchItems.length; k++) {
            (function (swatchItem) {
                swatchItem.element.addEventListener("click", function () {
                    var swatchRgb = hexToRGB(swatchItem.hex);
                    setRgb(swatchRgb.r, swatchRgb.g, swatchRgb.b);
                    safeRender();
                });
            })(pickerControls.swatchItems[k]);
        }

        pickerControls.grayCheckbox.onClick = function () {
            if (pickerControls.grayCheckbox.value) {
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

    // =========================================
    // 公開 API / Public API
    // =========================================

    /**
     * カラーピッカーを開き、選ばれた色を返す
     * @param {Object|string} showArgument - { value, title, lang }、または初期値の色文字列
     * @returns {string|null} "RRGGBB" / "cmyk:C,M,Y,K"。キャンセル時は null
     */
    function show(showArgument) {
        var pickerOptions = (typeof showArgument === "object" && showArgument !== null) ? showArgument : { value: showArgument };
        uiLang = pickerOptions.lang || "en";
        var state = createInitialState(pickerOptions.value || "000000");
        var pickerControls = buildPickerDialog(state, pickerOptions.title || "Color Picker");

        bindPickerEvents(state, pickerControls);
        renderPickerState(state, pickerControls);

        if (lastDialogLocation) {
            /* 前回の位置が画面外などで使えないときは既定の位置のまま / keep the default position if the saved one is rejected */
            try { pickerControls.dialog.location = lastDialogLocation; } catch (ePos) {}
        }

        var dialogResult = pickerControls.dialog.show();

        /* 閉じたあとの位置の読み取り。失敗しても色の返却は続ける / reading the location after close; keep going on failure */
        try { lastDialogLocation = pickerControls.dialog.location; } catch (ePosSave) {}

        if (dialogResult !== 1) return null;
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
