#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップグループのマスクパスの大きさを変更します。

詳細は README を参照してください。

### Overview

Resizes the mask path of a clipping group.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResizeClipMask";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeClipMask.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResizeClipMask.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 現在の言語取得 / Get current language */
    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* UIラベル定義 / UI label definitions */
    var LABELS = {
        notRect: { ja: "このマスクは長方形ではありません。処理をスキップします。", en: "This mask is not a rectangle. Skipping." },
        noMask: { ja: "マスクパスが見つかりませんでした。", en: "No mask path found." },
        noSelection: { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" },
        dialogTitle: {
            ja: "マスクパスのサイズ変更 " + SCRIPT_VERSION,
            en: "Resize Mask Path " + SCRIPT_VERSION
        },
        tipMargin: {
            ja: "マスクパスを外側に広げる量です。マイナスを入れると内側に縮みます。",
            en: "How far the mask path grows outward. A negative value shrinks it instead."
        },
        margin: { ja: "マージン", en: "Margin" }
    };

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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    /* プラスボタン処理 / Handle plus button */
    function handlePlus(input) {
        var val = parseFloat(input.text);
        if (isNaN(val)) val = 0;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.altKey ? 0.1 : 1;
        val += delta;
        val = keyboard.altKey ? Math.round(val * 10) / 10 : Math.round(val);
        input.text = String(val);
        if (typeof input.onChangeValue === "function") input.onChangeValue(val);
    }

    /* マイナスボタン処理 / Handle minus button */
    function handleMinus(input) {
        var val = parseFloat(input.text);
        if (isNaN(val)) val = 0;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.altKey ? 0.1 : 1;
        val -= delta;
        val = keyboard.altKey ? Math.round(val * 10) / 10 : Math.round(val);
        input.text = String(val);
        if (typeof input.onChangeValue === "function") input.onChangeValue(val);
    }

    /* 反転ボタン処理 / Handle swap button */
    function handleSwap(input) {
        var val = parseFloat(input.text);
        if (isNaN(val)) val = 0;
        val = -val;
        input.text = String(val);
        if (typeof input.onChangeValue === "function") input.onChangeValue(val);
    }

    /* マージンダイアログ表示 / Show margin dialog */
    function showMarginDialog(defaultValue, unitLabel, previewCallback) {
        var dlg = new Window("dialog", LABELS.dialogTitle[uiLang]);
        dlg.orientation = "column";
        dlg.alignChildren = "left";
        dlg.margins = 15;

        var inputGroup = dlg.add("group");
        inputGroup.add("statictext", undefined, LABELS.margin[uiLang] + " (" + unitLabel + "):");

        var inputSubGroup = inputGroup.add("group");
        inputSubGroup.orientation = "row";

        var input = inputSubGroup.add("edittext", undefined, defaultValue);
        input.helpTip = getLabel("tipMargin");
        input.characters = 4;
        changeValueByArrowKey(input);
        input.onChangeValue = previewCallback;

        var buttonGroup = inputSubGroup.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.spacing = 1;

        var plusMinusGroup = buttonGroup.add("group");
        plusMinusGroup.orientation = "column";
        plusMinusGroup.spacing = 1;

        var plusBtn = plusMinusGroup.add("button", [0, 0, 20, 15], "+");
        var minusBtn = plusMinusGroup.add("button", [0, 0, 20, 15], "-");

        var zeroGroup = buttonGroup.add("group");
        zeroGroup.orientation = "column";
        zeroGroup.alignChildren = ["center", "center"];

        var swapBtn = zeroGroup.add("button", [0, 0, 20, 31], "±");

        plusBtn.onClick = function() { handlePlus(input); };
        minusBtn.onClick = function() { handleMinus(input); };
        swapBtn.onClick = function() { handleSwap(input); };

        input.addEventListener("changing", function() {
            var val = parseFloat(input.text);
            if (!isNaN(val) && typeof previewCallback === "function") {
                previewCallback(val);
                app.redraw();
            }
        });

        var btns = dlg.add("group");
        btns.alignment = "right";
        var cancel = btns.add("button", undefined, LABELS.cancel[uiLang], { name: "cancel" });
        var ok = btns.add("button", undefined, LABELS.ok[uiLang], { name: "ok" });

        input.active = true;
        var result = dlg.show();
        if (result != 1) {
            if (typeof previewCallback === "function") restoreOriginalMaskRects();
            return null;
        }

        var margin = parseFloat(input.text);
        if (isNaN(margin)) {
            alert(LABELS.noSelection[uiLang]);
            return null;
        }
        return margin;
    }

    /* 選択からマスクパスを抽出 / Collect mask paths from selection */
    function collectMaskPaths(selection) {
        var masks = [];
        for (var i = 0; i < selection.length; i++) {
            var group = selection[i];
            if (group.typename === "GroupItem" && group.clipped) {
                for (var j = 0; j < group.pageItems.length; j++) {
                    var item = group.pageItems[j];
                    if (item.typename === "PathItem" && item.clipping) {
                        masks.push(item);
                        break;
                    }
                }
            }
        }
        return masks;
    }

    /* 元のマスク矩形情報保存用 / Store original mask rectangles */
    var originalRects = [];

    /* マスクに一時的にマージンを適用 / Apply temporary margin to masks */
    function applyTemporaryMarginToMasks(masks, margin) {
        restoreOriginalMaskRects(); // ★追加：まず元に戻す
        originalRects = [];
        for (var i = 0; i < masks.length; i++) {
            var mask = masks[i];
            if (!mask.clipping) continue;

            originalRects.push({
                mask: mask,
                top: mask.top,
                left: mask.left,
                width: mask.width,
                height: mask.height
            });

            mask.top += margin;
            mask.left -= margin;
            mask.width += margin * 2;
            mask.height += margin * 2;
        }
    }

    /* 元のマスクサイズに戻す / Restore original mask rectangles */
    function restoreOriginalMaskRects() {
        for (var i = 0; i < originalRects.length; i++) {
            var info = originalRects[i];
            info.mask.top = info.top;
            info.mask.left = info.left;
            info.mask.width = info.width;
            info.mask.height = info.height;
        }
    }

    /* メイン処理 / Main function */
    function main() {
        if (app.documents.length === 0) {
            alert(LABELS.noSelection[uiLang]);
            return;
        }
        var currentSelection = app.activeDocument.selection;
        if (currentSelection.length === 0) {
            alert(LABELS.noSelection[uiLang]);
            return;
        }
        var newSelection = collectMaskPaths(currentSelection);
        var marginUnit = getUnitInfo().label;
        var defaultMarginValue = '0';
        var margin = showMarginDialog(defaultMarginValue, marginUnit, function(previewMargin) {
            applyTemporaryMarginToMasks(newSelection, previewMargin);
            app.redraw();
        });
        if (margin === null) return;
        app.redraw();
        if (newSelection.length === 0) {
            alert(LABELS.noMask[uiLang]);
        }
    }
    main();

    /* 矢印キーで値を増減 / Change value by arrow key */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil(value / delta) * delta + delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor(value / delta) * delta - delta;
                    event.preventDefault();
                }
            } else {
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            }

            value = Math.round(value);
            editText.text = value;
            if (typeof editText.onChangeValue === "function") {
                editText.onChangeValue(value);
                app.redraw();
            }
        });
    }

})();
