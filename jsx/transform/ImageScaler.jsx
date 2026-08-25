#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択している配置画像（PlacedItem / RasterItem）の拡大・縮小率（%）を表示し、入力した値で再スケールします。

詳細は README を参照してください。

### Overview

Shows the scale (%) of the selected placed image (PlacedItem / RasterItem) and rescales it to the value you enter.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImageScaler";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-08-16";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageScaler.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageScaler.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        // ダイアログタイトル / Dialog title
        dialogTitle: {
            ja: "配置画像の拡大・縮小率 ",
            en: "Placed Image Scale "
        },
        // 入力ラベル / Input label
        scale: { ja: "スケール", en: "Scale" },
        // 単位 / Unit
        percent: { ja: "%", en: "%" },
        // ボタン / Buttons
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    };

    /* 対象判定 / Target filter */
    function isTargetItem(it) {
        if (!it) return false;
        var t = it.typename;
        return t === 'PlacedItem' || t === 'RasterItem';
    }

    /* 行列からのスケール算出 / Scale extraction from matrix */
    function getScalePercentXY(item) {
        var m = item.matrix; // a c tx / b d ty in Illustrator nomenclature (mValueA..F)
        var sx = Math.sqrt(m.mValueA * m.mValueA + m.mValueC * m.mValueC);
        var sy = Math.sqrt(m.mValueB * m.mValueB + m.mValueD * m.mValueD);
        return {
            x: sx * 100,
            y: sy * 100
        };
    }

    function round1(v) {
        return Math.round(v * 10) / 10;
    }

    /* 矢印キーでの値変更 / Arrow-key increment logic */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
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

    /* ダイアログの組み立て / Build dialog */
    function createDialog(defaultScaleText, targetItems) {
        var dlg = new Window('dialog', LABELS.dialogTitle[lang] + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];

        var inputGroup = dlg.add('group');
        inputGroup.orientation = 'row';
        inputGroup.alignChildren = ['left', 'center'];
        inputGroup.add('statictext', undefined, LABELS.scale[lang]);

        var scaleInput = inputGroup.add('edittext', undefined, defaultScaleText);
        scaleInput.characters = 4;
        changeValueByArrowKey(scaleInput);
        scaleInput.active = true;
        inputGroup.add('statictext', undefined, LABELS.percent[lang]);

        function applyFromField() {
            var val = parseFloat(scaleInput.text);
            if (isNaN(val) || val <= 0 || val > 1000) { return; }
            for (var j = 0; j < targetItems.length; j++) {
                var item = targetItems[j];
                var current = getScalePercentXY(item);
                var relX = (val / current.x) * 100;
                var relY = (val / current.y) * 100;
                item.resize(relX, relY, true, true, true, true, true, Transformation.CENTER);
            }
            app.redraw();
        }
        scaleInput.onChanging = applyFromField;

        var btnGroup = dlg.add('group');
        btnGroup.alignment = 'center';
        var cancelBtn = btnGroup.add('button', undefined, LABELS.cancel[lang]);
        var okBtn = btnGroup.add('button', undefined, LABELS.ok[lang]);
        okBtn.name = 'ok';
        cancelBtn.name = 'cancel';
        // Enter / Esc shortcuts
        dlg.defaultElement = okBtn;
        dlg.cancelElement = cancelBtn;

        okBtn.onClick = function() { dlg.close(); };
        cancelBtn.onClick = function() { dlg.close(); };

        return dlg;
    }

    /* メイン処理 / Main entry */
    function main() {
        if (app.documents.length === 0) {
            return;
        }
        var sel = app.activeDocument.selection;
        if (!sel || sel.length === 0) {
            return;
        }

        var targetItems = [];
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            if (!isTargetItem(it)) continue;
            targetItems.push(it);
        }

        if (targetItems.length === 0) {
            return;
        }

        var defaultScaleText = '100';
        if (targetItems.length === 1) {
            var s = getScalePercentXY(targetItems[0]);
            defaultScaleText = String(round1(s.x));
        }

        // ダイアログ作成 / Create dialog
        var dlg = createDialog(defaultScaleText, targetItems);
        dlg.show();
    }

    main();

})();
