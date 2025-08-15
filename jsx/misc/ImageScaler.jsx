#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

配置画像の拡大・縮小率

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択している配置画像（PlacedItem / RasterItem）の拡大・縮小率（%）を表示し、入力値で再スケールします。
- Rotation/Skew を含む変換行列から実スケールを算出し、入力フィールドでの↑↓キー操作による値変更にも対応します。

### 主な機能：

- 実スケール（X/Y）を行列から算出
- ダイアログでスケール%を入力（単一選択時は現在値を初期表示）
- 入力値で相対倍率を計算して即時適用
- キーボード操作での数値調整：
  - ↑↓=±1
  - Shift+↑↓=±10（10の倍数にスナップ）
  - Option(Alt)+↑↓=±0.1

### 処理の流れ：

1. 選択オブジェクトを確認し、対象（配置画像／ラスタ画像）のみ抽出
2. 単一選択時は現在のスケール値を初期表示
3. ダイアログ表示、入力値の即時反映（app.redraw）
4. OKボタンで閉じる

### 更新履歴：

- v1.0 (20250816) : 初期バージョン
- v1.1 (20250816) : キーボード操作による値変更機能追加
- v1.2 (20250816) : 入力値の即時適用（OKは閉じるのみ）、ローカライズ対応

---

### Script Name:

Placed Image Scale

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Displays and rescales the scale percentage (%) of selected placed images (PlacedItem / RasterItem).
- Calculates actual scale from transformation matrix including rotation/skew, and supports value changes via arrow keys in the input field.

### Main Features:

- Calculate actual X/Y scale from transformation matrix
- Dialog with scale % input (prefilled for single selection)
- Apply relative scaling immediately on value change
- Keyboard increments:
  - ↑↓ = ±1
  - Shift+↑↓ = ±10 (snap to multiples of 10)
  - Option(Alt)+↑↓ = ±0.1

### Process Flow:

1. Check selection and filter to placed/raster items
2. Prefill scale value if only one item is selected
3. Show dialog, apply changes immediately (app.redraw)
4. OK button closes the dialog

### Changelog:

- v1.0 (20250816) : Initial version
- v1.1 (20250816) : Added arrow key increment feature
- v1.2 (20250816) : Immediate application of changes (OK closes only), localization support
*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.3";

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