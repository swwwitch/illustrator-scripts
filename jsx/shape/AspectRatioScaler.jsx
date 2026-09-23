#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、指定した縦横比に合わせて拡大・縮小します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4a212e6eacf1

### Overview

Scales the selected objects to a specified aspect ratio.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AspectRatioScaler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AspectRatioScaler";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AspectRatioScaler.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4a212e6eacf1"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プリセットの比率（横 ÷ 縦）/ Preset ratios (width / height) */
    var RATIO_16_9   = 16 / 9;
    var RATIO_SQUARE = 1;
    var RATIO_A4     = 210 / 297;

    /* カスタム比率の初期値 / Initial custom ratio */
    var DEFAULT_CUSTOM_RATIO_WIDTH  = "3";
    var DEFAULT_CUSTOM_RATIO_HEIGHT = "2";

    /* 選択なしで作る長方形の、サイズ欄が空のときの大きさ（pt）/ Rectangle size when nothing is selected and the size field is empty */
    var FALLBACK_BASE_SIZE_PT = 200;

    /* 選択なしのときサイズ欄に入れる初期値（単位コード → 値）/ Size field default when nothing is selected (unit code -> value) */
    var DEFAULT_SIZE_TEXT_BY_UNIT = { 1: "100", 6: "1000" };

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS    = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] */
    var FIELD_CHARACTERS = 5;                  /* 数値欄の幅（文字数）/ Numeric field width */
    var DIALOG_OFFSET_X  = 300;                /* ダイアログを右へずらす量 / Horizontal dialog offset */
    var DIALOG_OPACITY   = 0.97;               /* ダイアログの不透明度 / Dialog opacity */

    /**
     * 見出し付きパネルを縦並びで追加する
     * @param {Group} parent - 追加先
     * @param {string} title - パネルの見出し
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, title) {
        var panel = parent.add("panel", undefined, title);
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        panel.alignment = ["fill", "top"];
        return panel;
    }

    /**
     * 子を並べるグループを追加する
     * @param {Object} parent - 追加先のパネルまたはグループ
     * @param {string} orientation - "row" / "column"
     * @returns {Group} 追加したグループ
     */
    function addStackGroup(parent, orientation) {
        var stackGroup = parent.add("group");
        stackGroup.orientation = orientation;
        stackGroup.alignChildren = "left";
        return stackGroup;
    }

    /**
     * ダイアログを表示時に横へずらす
     * @param {Window} dlg - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            dlg.location = [dlg.location[0] + offsetX, dlg.location[1] + offsetY];
        };
    }

    /**
     * 上下キーで数値を増減する（shift: ±10、option: ±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
            if (typeof editText.onChanging === "function") editText.onChanging();
        });
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 定規の単位に合わせて丸める（px は整数、mm は 0.1mm 刻み、その他は 0.01pt 刻み）
     * @param {number} valuePt - 値（pt）
     * @returns {number} 丸めた値（pt）
     */
    function roundForUnit(valuePt) {
        var unitCode = getUnitInfo().code;
        if (unitCode === 6) return Math.round(valuePt); /* 1px = 1pt */
        if (unitCode === 1) {
            var stepPt = UNITS[1].pointsPerUnit * 0.1;
            return Math.round(valuePt / stepPt) * stepPt;
        }
        return Math.round(valuePt * 100) / 100;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIの言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "縦横比で調整", en: "Adjust by Aspect Ratio" }
        },
        panel: {
            aspectRatio: { ja: "アスペクト比", en: "Aspect Ratio" },
            orientation: { ja: "向き", en: "Orientation" },
            sizeAdjust: { ja: "サイズ調整", en: "Size Adjustment" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            ratio16x9: { ja: "16:9", en: "16:9" },
            ratioSquare: { ja: "1:1（スクエア）", en: "1:1 (Square)" },
            ratioA4: { ja: "A4（1:1.414）", en: "A4 (1:1.414)" },
            ratioCustom: { ja: "カスタム", en: "Custom" },
            landscape: { ja: "横（ランドスケープ）", en: "Landscape" },
            portrait: { ja: "縦（ポートレート）", en: "Portrait" },
            basisHorizontal: { ja: "幅", en: "Width" },
            basisVertical: { ja: "高さ", en: "Height" }
        },
        fieldLabel: {
            basis: { ja: "固定する辺", en: "Keep" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        checkbox: {
            alignToPixelGrid: { ja: "ピクセルを最適化", en: "Make Pixel Perfect" },
            addArtboard: { ja: "アートボードを追加", en: "Add Artboard" }
        },
        tooltip: {
            ratioPreset: {
                ja: "よく使う比率です。選ぶとカスタム欄は使いません。",
                en: "Common ratios. Selecting one disables the custom fields."
            },
            ratioCustom: { ja: "下の欄に好きな比率を入力します。", en: "Enter any ratio in the fields below." },
            customWidth: { ja: "カスタム比の左側（横）の値です。", en: "The left (horizontal) value of the custom ratio." },
            customHeight: { ja: "カスタム比の右側（縦）の値です。", en: "The right (vertical) value of the custom ratio." },
            landscape: { ja: "長い辺を横にします。", en: "Puts the longer side horizontally." },
            portrait: { ja: "長い辺を縦にします。", en: "Puts the longer side vertically." },
            basisHorizontal: {
                ja: "幅を保ったまま高さを比率に合わせます。",
                en: "Keeps the width and fits the height to the ratio."
            },
            basisVertical: {
                ja: "高さを保ったまま幅を比率に合わせます。",
                en: "Keeps the height and fits the width to the ratio."
            },
            sizeValue: {
                ja: "固定する辺の長さです。空欄なら選択範囲の大きさを使います。",
                en: "Length of the side to keep. Leave blank to use the size of the selection."
            },
            alignToPixelGrid: {
                ja: "結果の座標と大きさを整数ピクセルに丸めます。",
                en: "Rounds the resulting position and size to whole pixels."
            },
            addArtboard: {
                ja: "結果の範囲にアートボードを追加します。オブジェクトは残ります。",
                en: "Adds an artboard that matches the result. The objects stay in place."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: "...", en: "..." }
     * @returns {string} ラベル
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * コロン付きのラベルを返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: "...", en: "..." }
     * @returns {string} コロン付きのラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを追加する
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 表示名
     * @param {Object} tipSet - ツールチップ
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parent, labelSet, tipSet) {
        var radio = parent.add("radiobutton", undefined, getLabel(labelSet));
        radio.helpTip = getLabel(tipSet);
        return radio;
    }

    /**
     * 数値入力欄を追加する
     * @param {Group} parent - 追加先
     * @param {string} initialText - 初期値
     * @param {Object} tipSet - ツールチップ
     * @returns {EditText} 追加した入力欄
     */
    function addNumberField(parent, initialText, tipSet) {
        var field = parent.add("edittext", undefined, initialText);
        field.helpTip = getLabel(tipSet);
        field.characters = FIELD_CHARACTERS;
        changeValueByArrowKey(field);
        return field;
    }

    /**
     * ダイアログを作成する
     * @returns {Object} ダイアログと各コントロールの参照
     */
    function createDialog() {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.opacity = DIALOG_OPACITY;
        shiftDialogPosition(dialog, DIALOG_OFFSET_X, 0);
        dialog.alignChildren = ["fill", "top"];

        /* アスペクト比 / Aspect ratio */
        var aspectPanel = addPanel(dialog, getLabel(LABELS.panel.aspectRatio));
        var ratioRadioGroup = addStackGroup(aspectPanel, "column");
        var ratio16x9Radio = addRadio(ratioRadioGroup, LABELS.radio.ratio16x9, LABELS.tooltip.ratioPreset);
        var ratioSquareRadio = addRadio(ratioRadioGroup, LABELS.radio.ratioSquare, LABELS.tooltip.ratioPreset);
        var ratioA4Radio = addRadio(ratioRadioGroup, LABELS.radio.ratioA4, LABELS.tooltip.ratioPreset);
        var ratioCustomRadio = addRadio(ratioRadioGroup, LABELS.radio.ratioCustom, LABELS.tooltip.ratioCustom);
        ratio16x9Radio.value = true;

        var customRatioGroup = addStackGroup(aspectPanel, "row");
        var customWidthInput = addNumberField(customRatioGroup, DEFAULT_CUSTOM_RATIO_WIDTH, LABELS.tooltip.customWidth);
        customRatioGroup.add("statictext", undefined, ":");
        var customHeightInput = addNumberField(customRatioGroup, DEFAULT_CUSTOM_RATIO_HEIGHT, LABELS.tooltip.customHeight);
        customWidthInput.enabled = false;
        customHeightInput.enabled = false;

        /* 向き / Orientation */
        var orientationPanel = addPanel(dialog, getLabel(LABELS.panel.orientation));
        var orientationRadioGroup = addStackGroup(orientationPanel, "column");
        var landscapeRadio = addRadio(orientationRadioGroup, LABELS.radio.landscape, LABELS.tooltip.landscape);
        var portraitRadio = addRadio(orientationRadioGroup, LABELS.radio.portrait, LABELS.tooltip.portrait);
        landscapeRadio.value = true;

        /* サイズ調整（固定する辺＋サイズ）/ Size adjustment (side to keep + size) */
        var sizeAdjustPanel = addPanel(dialog, getLabel(LABELS.panel.sizeAdjust));

        var basisRow = sizeAdjustPanel.add("group");
        basisRow.orientation = "row";
        basisRow.alignChildren = ["left", "center"];
        basisRow.add("statictext", undefined, labelText(LABELS.fieldLabel.basis));
        var basisHorizontalRadio = addRadio(basisRow, LABELS.radio.basisHorizontal, LABELS.tooltip.basisHorizontal);
        var basisVerticalRadio = addRadio(basisRow, LABELS.radio.basisVertical, LABELS.tooltip.basisVertical);
        basisHorizontalRadio.value = true;

        var sizeRow = sizeAdjustPanel.add("group");
        sizeRow.orientation = "row";
        sizeRow.alignChildren = ["left", "center"];

        var sizeFieldLabel = sizeRow.add("statictext", undefined, labelText(LABELS.fieldLabel.width));
        /* 「幅」「高さ」の長いほうの幅を確保 / Reserve room for the longer of width/height */
        sizeFieldLabel.preferredSize.width = Math.max(
            sizeFieldLabel.graphics.measureString(labelText(LABELS.fieldLabel.width))[0],
            sizeFieldLabel.graphics.measureString(labelText(LABELS.fieldLabel.height))[0]
        );
        var sizeInput = addNumberField(sizeRow, "", LABELS.tooltip.sizeValue);
        sizeRow.add("statictext", undefined, getUnitInfo().label);

        /* オプション / Options */
        var optionPanel = addPanel(dialog, getLabel(LABELS.panel.options));
        var alignToPixelCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.alignToPixelGrid));
        alignToPixelCheckbox.helpTip = getLabel(LABELS.tooltip.alignToPixelGrid);
        alignToPixelCheckbox.value = (getUnitInfo().code === 6); /* px のときだけ ON / on only for px */

        var addArtboardCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.addArtboard));
        addArtboardCheckbox.helpTip = getLabel(LABELS.tooltip.addArtboard);

        /* ボタン / Buttons */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return {
            dialog: dialog,
            ratio16x9Radio: ratio16x9Radio,
            ratioSquareRadio: ratioSquareRadio,
            ratioA4Radio: ratioA4Radio,
            ratioCustomRadio: ratioCustomRadio,
            customWidthInput: customWidthInput,
            customHeightInput: customHeightInput,
            landscapeRadio: landscapeRadio,
            portraitRadio: portraitRadio,
            basisHorizontalRadio: basisHorizontalRadio,
            basisVerticalRadio: basisVerticalRadio,
            sizeFieldLabel: sizeFieldLabel,
            sizeInput: sizeInput,
            alignToPixelCheckbox: alignToPixelCheckbox,
            addArtboardCheckbox: addArtboardCheckbox
        };
    }

    /**
     * ダイアログから選択中の比率（横 ÷ 縦）を読む
     * @param {Object} ui - createDialog() の戻り値
     * @returns {number} 比率。カスタムが数値でないか 0 以下なら 1
     */
    function readRatio(ui) {
        if (ui.ratio16x9Radio.value) return RATIO_16_9;
        if (ui.ratioSquareRadio.value) return RATIO_SQUARE;
        if (ui.ratioA4Radio.value) return RATIO_A4;
        var ratioWidth = parseFloat(ui.customWidthInput.text);
        var ratioHeight = parseFloat(ui.customHeightInput.text);
        if (isNaN(ratioWidth) || isNaN(ratioHeight) || ratioWidth <= 0 || ratioHeight <= 0) return 1;
        return ratioWidth / ratioHeight;
    }

    /**
     * サイズ欄の値を pt で返す
     * @param {Object} ui - createDialog() の戻り値
     * @returns {number|null} 値（pt）。空欄・不正値・0以下なら null
     */
    function readTargetSizePt(ui) {
        var sizeValue = parseFloat(ui.sizeInput.text);
        if (isNaN(sizeValue) || sizeValue <= 0) return null;
        return sizeValue * getUnitInfo().pointsPerUnit;
    }

    /**
     * ダイアログの設定をまとめて読む
     * @param {Object} ui - createDialog() の戻り値
     * @returns {{ratio: number, wantPortrait: boolean, fixByHeight: boolean, targetSizePt: (number|null)}} 設定
     */
    function readSettings(ui) {
        return {
            ratio: readRatio(ui),
            wantPortrait: ui.portraitRadio.value,
            fixByHeight: ui.basisVerticalRadio.value,
            targetSizePt: readTargetSizePt(ui)
        };
    }

    // =========================================
    // 計算とプレビュー / Calculation and preview
    // =========================================

    /**
     * 比率を向きに合わせて反転する
     * @param {number} ratio - 比率（横 ÷ 縦）
     * @param {boolean} wantPortrait - 縦置きなら true
     * @returns {number} 向きをそろえた比率
     */
    function orientRatio(ratio, wantPortrait) {
        if (wantPortrait && ratio > 1) return 1 / ratio;
        if (!wantPortrait && ratio < 1) return 1 / ratio;
        return ratio;
    }

    /**
     * 固定する辺の長さから、比率に合う幅と高さを求める
     * @param {number} orientedRatio - 向きをそろえた比率
     * @param {boolean} fixByHeight - 高さを固定するなら true
     * @param {number} baseSizePt - 固定する辺の長さ（pt）
     * @returns {{width: number, height: number}} 幅と高さ（pt）
     */
    function computeTargetSize(orientedRatio, fixByHeight, baseSizePt) {
        if (fixByHeight) {
            return { width: baseSizePt * orientedRatio, height: roundForUnit(baseSizePt) };
        }
        return { width: baseSizePt, height: roundForUnit(baseSizePt / orientedRatio) };
    }

    /**
     * プレビューのアイテムに比率を当てる
     * @param {Object} preview - { items, originalWidths, originalHeights }
     * @param {Object} settings - readSettings() の戻り値
     * @returns {void}
     */
    function applyAspect(preview, settings) {
        var orientedRatio = orientRatio(settings.ratio, settings.wantPortrait);
        for (var i = 0; i < preview.items.length; i++) {
            var baseSizePt = settings.targetSizePt;
            if (baseSizePt === null) {
                baseSizePt = settings.fixByHeight ? preview.originalHeights[i] : preview.originalWidths[i];
            }
            var targetSize = computeTargetSize(orientedRatio, settings.fixByHeight, baseSizePt);
            preview.items[i].width = targetSize.width;
            preview.items[i].height = targetSize.height;
        }
        app.redraw();
    }

    /**
     * 選択アイテムの複製をプレビュー用に作り、元は隠す
     * @param {Object[]} selectedItems - 選択アイテム
     * @returns {Object} { items, originalWidths, originalHeights }
     */
    function createPreviewFromSelection(selectedItems) {
        var preview = { items: [], originalWidths: [], originalHeights: [] };
        for (var i = 0; i < selectedItems.length; i++) {
            var previewCopy = selectedItems[i].duplicate();
            previewCopy.zOrder(ZOrderMethod.BRINGTOFRONT);
            preview.items.push(previewCopy);
            preview.originalWidths.push(selectedItems[i].width);
            preview.originalHeights.push(selectedItems[i].height);
            selectedItems[i].hidden = true;
        }
        return preview;
    }

    /**
     * 選択がないとき、アクティブなアートボードの中央に長方形を作ってプレビューにする
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} settings - readSettings() の戻り値
     * @returns {Object} { items, originalWidths, originalHeights }
     */
    function createPreviewRectangle(doc, settings) {
        var baseSizePt = (settings.targetSizePt === null) ? FALLBACK_BASE_SIZE_PT : settings.targetSizePt;
        var targetSize = computeTargetSize(orientRatio(settings.ratio, settings.wantPortrait), settings.fixByHeight, baseSizePt);

        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; /* [L,T,R,B] */
        var centerX = (artboardRect[0] + artboardRect[2]) / 2;
        var centerY = (artboardRect[1] + artboardRect[3]) / 2;
        var rect = doc.pathItems.rectangle(centerY + targetSize.height / 2, centerX - targetSize.width / 2, targetSize.width, targetSize.height);
        rect.stroked = false;
        rect.filled = true;

        return { items: [rect], originalWidths: [targetSize.width], originalHeights: [targetSize.height] };
    }

    // =========================================
    // 確定と取り消し / Commit and cancel
    // =========================================

    /**
     * 確定後の仕上げ（ピクセル最適化・アートボード追加）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} item - 仕上げるアイテム
     * @param {Object} ui - createDialog() の戻り値
     * @returns {void}
     */
    function finishItem(doc, item, ui) {
        if (ui.alignToPixelCheckbox.value) {
            doc.selection = [item];
            app.executeMenuCommand("Make Pixel Perfect");
        }
        if (ui.addArtboardCheckbox.value) {
            doc.artboards.add(item.visibleBounds);
        }
    }

    /**
     * プレビューの大きさと位置を元のアイテムに移し、プレビューを消す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} selectedItems - 元の選択アイテム
     * @param {Object} preview - { items, originalWidths, originalHeights }
     * @param {Object} ui - createDialog() の戻り値
     * @returns {void}
     */
    function commitToOriginals(doc, selectedItems, preview, ui) {
        for (var i = 0; i < selectedItems.length; i++) {
            var original = selectedItems[i];
            var previewCopy = preview.items[i];
            original.hidden = false;

            /* 中心基準で拡大縮小し、左上をそろえる / Scale from center, then align the top-left */
            var originalWidth = preview.originalWidths[i];
            var originalHeight = preview.originalHeights[i];
            if (originalWidth > 0 && originalHeight > 0) {
                original.resize(previewCopy.width / originalWidth * 100, previewCopy.height / originalHeight * 100);
            }
            original.position = previewCopy.position;
            previewCopy.remove();

            finishItem(doc, original, ui);
        }
        doc.selection = selectedItems;
    }

    /**
     * プレビューを消し、隠した元のアイテムを戻す
     * @param {Object[]} selectedItems - 元の選択アイテム
     * @param {Object} preview - { items, originalWidths, originalHeights }
     * @returns {void}
     */
    function cancelPreview(selectedItems, preview) {
        for (var i = 0; i < preview.items.length; i++) {
            preview.items[i].remove();
        }
        for (var j = 0; j < selectedItems.length; j++) {
            selectedItems[j].hidden = false;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        var hasSelection = (selectedItems && selectedItems.length > 0);
        if (!hasSelection) selectedItems = [];

        var ui = createDialog();

        var preview;
        if (hasSelection) {
            preview = createPreviewFromSelection(selectedItems);
        } else {
            /* 選択なしのときはサイズ欄に既定値（px:1000 / mm:100）/ Default size when nothing is selected */
            ui.sizeInput.text = DEFAULT_SIZE_TEXT_BY_UNIT[getUnitInfo().code] || "";
            preview = createPreviewRectangle(doc, readSettings(ui));
        }

        /* 設定が変わるたびにプレビューを更新 / Refresh the preview on every change */
        function refreshPreview() {
            ui.customWidthInput.enabled = ui.ratioCustomRadio.value;
            ui.customHeightInput.enabled = ui.ratioCustomRadio.value;
            ui.sizeFieldLabel.text = labelText(ui.basisVerticalRadio.value ? LABELS.fieldLabel.height : LABELS.fieldLabel.width);
            applyAspect(preview, readSettings(ui));
        }
        var clickControls = [
            ui.ratio16x9Radio, ui.ratioSquareRadio, ui.ratioA4Radio, ui.ratioCustomRadio,
            ui.landscapeRadio, ui.portraitRadio, ui.basisHorizontalRadio, ui.basisVerticalRadio
        ];
        for (var i = 0; i < clickControls.length; i++) {
            clickControls[i].onClick = refreshPreview;
        }
        ui.customWidthInput.onChanging = refreshPreview;
        ui.customHeightInput.onChanging = refreshPreview;
        ui.sizeInput.onChanging = refreshPreview;

        refreshPreview();

        if (ui.dialog.show() !== 1) {
            cancelPreview(selectedItems, preview);
            return;
        }

        if (hasSelection) {
            commitToOriginals(doc, selectedItems, preview, ui);
        } else {
            /* 作った長方形をそのまま残す / Keep the preview rectangle as the result */
            finishItem(doc, preview.items[0], ui);
            doc.selection = [preview.items[0]];
        }
        app.redraw();
    }

    main();

})();
