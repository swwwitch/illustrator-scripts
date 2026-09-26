#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択している配置画像（PlacedItem / RasterItem）の拡大・縮小率（%）を表示し、入力した値で再スケールします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageScaler.md

### Overview

Shows the scale (%) of the selected placed image (PlacedItem / RasterItem) and rescales it to the value you enter.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageScaler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImageScaler";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageScaler.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageScaler.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 受け付ける拡大・縮小率の上限（%） / Largest scale accepted (%) */
    var MAX_SCALE_PERCENT = 1000;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var SCALE_FIELD_CHARACTERS = 4; /* 拡大・縮小率の入力欄の幅（文字数） / width of the scale field in characters */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        dialog: {
            title: { ja: "配置画像の拡大・縮小率", en: "Placed Image Scale" }
        },
        fieldLabel: {
            scale: { ja: "スケール", en: "Scale" },
            percentUnit: { ja: "%", en: "%" }
        },
        tooltip: {
            scale: {
                ja: "配置画像の拡大・縮小率です。100 で原寸。↑↓キーで増減できます。",
                en: "Scale of the placed image. 100 is the original size. The arrow keys step the value."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} コロン付きのラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 拡大・縮小 / Scaling
    // =========================================

    /**
     * 拡大・縮小の対象（配置画像・埋め込み画像）かどうかを返す
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @returns {boolean} PlacedItem か RasterItem なら true
     */
    function isImageItem(targetItem) {
        if (!targetItem) return false;
        return targetItem.typename === "PlacedItem" || targetItem.typename === "RasterItem";
    }

    /**
     * 変換行列から現在の拡大・縮小率を求める
     * @param {PlacedItem|RasterItem} imageItem - 対象の画像
     * @returns {{x: number, y: number}} 横・縦の拡大・縮小率（%）
     */
    function getScalePercent(imageItem) {
        var imageMatrix = imageItem.matrix;
        var scaleX = Math.sqrt(imageMatrix.mValueA * imageMatrix.mValueA + imageMatrix.mValueC * imageMatrix.mValueC);
        var scaleY = Math.sqrt(imageMatrix.mValueB * imageMatrix.mValueB + imageMatrix.mValueD * imageMatrix.mValueD);
        return { x: scaleX * 100, y: scaleY * 100 };
    }

    /**
     * 小数第1位に丸める
     * @param {number} value - 丸める値
     * @returns {number} 丸めた値
     */
    function roundToOneDecimal(value) {
        return Math.round(value * 10) / 10;
    }

    /**
     * 各画像を、指定した拡大・縮小率になるよう中心基準で拡大・縮小する
     * @param {Array<PlacedItem|RasterItem>} imageItems - 対象の画像
     * @param {number} targetPercent - 目標の拡大・縮小率（%）
     * @returns {void}
     */
    function applyScaleToItems(imageItems, targetPercent) {
        for (var i = 0; i < imageItems.length; i++) {
            var currentScale = getScalePercent(imageItems[i]);
            var relativeX = (targetPercent / currentScale.x) * 100;
            var relativeY = (targetPercent / currentScale.y) * 100;
            imageItems[i].resize(relativeX, relativeY, true, true, true, true, true, Transformation.CENTER);
        }
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値入力欄に↑↓キーでの増減を付ける（shift で10刻み、option で0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
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

            if (editText.onChanging) editText.onChanging();

        });
    }

    /**
     * ダイアログを組み立てる。入力のたびに画像へ即時適用し、OK・キャンセルは閉じるだけ
     * @param {string} defaultScaleText - 入力欄の初期値
     * @param {Array<PlacedItem|RasterItem>} imageItems - 対象の画像
     * @returns {Window} 組み立てたダイアログ
     */
    function createDialog(defaultScaleText, imageItems) {
        var dlg = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        var scaleRowGroup = dlg.add("group");
        scaleRowGroup.orientation = "row";
        scaleRowGroup.alignChildren = ["left", "center"];
        scaleRowGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.scale));

        var scaleInput = scaleRowGroup.add("edittext", undefined, defaultScaleText);
        scaleInput.helpTip = getLabel(LABELS.tooltip.scale);
        scaleInput.characters = SCALE_FIELD_CHARACTERS;
        changeValueByArrowKey(scaleInput);
        scaleInput.active = true;
        scaleRowGroup.add("statictext", undefined, getLabel(LABELS.fieldLabel.percentUnit));

        /* 範囲外や数値でない入力は無視する / Ignore values that are out of range or not numbers */
        scaleInput.onChanging = function () {
            var targetPercent = parseFloat(scaleInput.text);
            if (isNaN(targetPercent) || targetPercent <= 0 || targetPercent > MAX_SCALE_PERCENT) return;
            applyScaleToItems(imageItems, targetPercent);
        };

        var btnRowGroup = dlg.add("group");
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        /* Enter / Esc で閉じる / Enter and Esc close the dialog */
        dlg.defaultElement = btnOK;
        dlg.cancelElement = btnCancel;

        btnOK.onClick = function () { dlg.close(); };
        btnCancel.onClick = function () { dlg.close(); };

        return dlg;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から画像を集め、拡大・縮小率のダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var selectedItems = app.activeDocument.selection;
        if (!selectedItems || selectedItems.length === 0) return;

        var imageItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (isImageItem(selectedItems[i])) imageItems.push(selectedItems[i]);
        }
        if (imageItems.length === 0) return;

        /* 1つだけなら現在の率を初期値に、複数なら 100 / Show the current scale for a single image, 100 for several */
        var defaultScaleText = "100";
        if (imageItems.length === 1) {
            defaultScaleText = String(roundToOneDecimal(getScalePercent(imageItems[0]).x));
        }

        createDialog(defaultScaleText, imageItems).show();
    }

    main();

})();
