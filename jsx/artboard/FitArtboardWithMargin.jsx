#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

FitArtboardWithMargin.jsx

### 概要

- 選択オブジェクトの外接範囲にマージンを加えてアートボードサイズを自動調整
- 単位に応じたマージン初期値を設定し、即時プレビュー可能なダイアログを提供

### 主な機能

- 定規単位ごとのマージン初期値設定
- 外接バウンディングボックス計算
- 即時プレビュー付きダイアログ
- ピクセル整数値への丸め

### 処理の流れ

1. 選択オブジェクトがない場合は全オブジェクトを対象
2. 単位に応じて初期マージン値を設定
3. ダイアログでマージンを入力（即時プレビュー）
4. バウンディングボックスとマージン適用
5. アートボードを更新

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

- 数字入力支援（+/-ボタン、0ボタン）
- ダイアログボックスを閉じずにプレビュー更新
- 単位系（mm、px など）によってデフォルト値を切り替え
- アートボードの座標・サイズをピクセルベースで整数値に
- オブジェクトを選択していない場合には、すべてのオブジェクトを対象に

### 更新履歴

- v1.0 (20250420) : 初期バージョン
- v1.1 (20250708) : UI改善、ポイント初期値変更
- v1.2 (20250709) : UI改善とバグ修正

---

### Script Name:

FitArtboardWithMargin.jsx

### Overview

- Automatically resizes the artboard to fit the bounding box of selected objects with margin
- Sets default margin value by ruler units and provides an instant preview dialog

### Main Features

- Margin defaults per ruler units
- Bounding box calculation
- Live preview dialog
- Rounds to integer pixel values

### Workflow

1. Use all objects if nothing is selected
2. Set default margin value by units
3. Input margin in dialog (live preview)
4. Apply bounding box and margin
5. Update artboard

### Changelog

- v1.0 (20250420) : Initial version
- v1.1 (20250708) : UI improvements, default point value updated
- v1.2 (20250709) : UI improvements and bug fixes

*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// スクリプトバージョン
var SCRIPT_VERSION = "v1.2";

// UIラベル定義 / UI Label Definitions
var LABELS = {
    dialogTitle: {
        ja: "アートボードサイズを調整 " + SCRIPT_VERSION,
        en: "Adjust Artboard Size " + SCRIPT_VERSION
    },
    marginLabel: {
        ja: "マージン",
        en: "Margin"
    },
    numberAlert: {
        ja: "数値を入力してください。",
        en: "Please enter a number."
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    }
};

/*
マージンダイアログ表示（ScriptUI）
Show margin input dialog with live preview
*/
function showMarginDialog(defaultValue, unit, lang) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = [10, 15, 10, 10];

    var inputGroup = dlg.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = ["center", "top"];
    inputGroup.spacing = 15;

    var labelText = LABELS.marginLabel[lang] + " (" + unit + ")";
    inputGroup.add("statictext", undefined, labelText);

    var inputSubGroup = inputGroup.add("group");
    inputSubGroup.orientation = "row";

    var input = inputSubGroup.add("edittext", undefined, defaultValue);
    input.characters = 4;

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

    var zeroBtn = zeroGroup.add("button", [0, 0, 20, 31], "0");

    // 保存しておくアートボード座標 / Save original artboard coordinates
    var abIndex = app.activeDocument.artboards.getActiveArtboardIndex();
    var originalRect = app.activeDocument.artboards[abIndex].artboardRect.slice();

    zeroBtn.onClick = function() {
        input.text = "0";
        updatePreview(input.text);
    };

    plusBtn.onClick = function() {
        var val = parseFloat(input.text);
        if (!isNaN(val)) {
            val += 1;
            input.text = val.toString();
            updatePreview(input.text);
        }
    };

    minusBtn.onClick = function() {
        var val = parseFloat(input.text);
        if (!isNaN(val)) {
            val = Math.max(0, val - 1);
            input.text = val.toString();
            updatePreview(input.text);
        }
    };

    // プレビュー更新 / Update preview of artboard size
    function updatePreview(value) {
        var previewValue = parseFloat(value);
        if (!isNaN(previewValue)) {
            var previewMarginInPoints = new UnitValue(previewValue, unit).as('pt');
            var previewBounds = getMaxBounds(app.activeDocument.selection.length === 0 ? app.activeDocument.pageItems : app.activeDocument.selection);
            previewBounds[0] -= previewMarginInPoints;
            previewBounds[1] += previewMarginInPoints;
            previewBounds[2] += previewMarginInPoints;
            previewBounds[3] -= previewMarginInPoints;

            app.activeDocument.artboards[abIndex].artboardRect = previewBounds;
            app.redraw();
        }
    }

    input.onChanging = function() {
        updatePreview(input.text);
    };

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });
    btnGroup.margins = [0, 5, 0, 0];

    var result = null;

    okBtn.onClick = function() {
        if (!isNaN(parseFloat(input.text))) {
            result = input.text;
            updatePreview(result);
            dlg.close();
        } else {
            alert(LABELS.numberAlert[lang]);
        }
    };

    cancelBtn.onClick = function() {
        app.activeDocument.artboards[abIndex].artboardRect = originalRect;
        app.redraw();
        dlg.close();
    };

    updatePreview(defaultValue);

    dlg.show();
    return result;
}

/*
メイン処理 / Main process
*/
function main() {
    var selectedItems, artboards, rulerType, marginUnit, marginValue;
    var marginInPoints, defaultMarginValue, artboardIndex, selectedBounds;
    var supportedUnits = ['inch', 'mm', 'pt', 'pica', 'cm', 'H', 'px'];
    var lang = getCurrentLang();

    try {
        var doc = app.activeDocument;
        selectedItems = doc.selection;
        if (selectedItems.length === 0) {
            selectedItems = doc.pageItems;
            if (selectedItems.length === 0) return;
        }

        artboards = doc.artboards;
        rulerType = app.preferences.getIntegerPreference("rulerType");
        marginUnit = supportedUnits[rulerType];

        // 単位ごとの初期マージン値設定 / Set default margin value based on unit
        defaultMarginValue = '0';
        if (marginUnit === 'mm') {
            defaultMarginValue = '5';
        } else if (marginUnit === 'px') {
            defaultMarginValue = '20';
        } else if (marginUnit === 'pt') {
            defaultMarginValue = '10';
        }

        // ユーザーにマージンを入力させる（ScriptUI ダイアログ）/ Show margin input dialog with live preview
        var userInput = showMarginDialog(defaultMarginValue, marginUnit, lang);
        if (!userInput) return;

        marginValue = parseFloat(userInput);
        marginInPoints = new UnitValue(marginValue, marginUnit).as('pt');

        // 選択範囲のバウンディングボックス計算 / Calculate bounding box of selected items
        selectedBounds = getMaxBounds(selectedItems);
        selectedBounds[0] -= marginInPoints;
        selectedBounds[1] += marginInPoints;
        selectedBounds[2] += marginInPoints;
        selectedBounds[3] -= marginInPoints;

        // 座標と幅・高さを整数に丸める / Round coordinates and size to integers
        var x0 = Math.round(selectedBounds[0]);
        var y1 = Math.round(selectedBounds[1]);
        var x2 = Math.round(selectedBounds[2]);
        var y3 = Math.round(selectedBounds[3]);

        var width = Math.round(x2 - x0);
        var height = Math.round(y1 - y3);

        x2 = x0 + width;
        y3 = y1 - height;

        selectedBounds[0] = x0;
        selectedBounds[1] = y1;
        selectedBounds[2] = x2;
        selectedBounds[3] = y3;

        // アートボードの更新 / Update artboard
        artboardIndex = artboards.getActiveArtboardIndex();
        artboards[artboardIndex].artboardRect = selectedBounds;

    } catch (e) {
        alert("エラーが発生しました: " + e.message);
    }
}

/*
選択オブジェクト群から最大のバウンディングボックスを取得
Get maximum bounding box from multiple items
*/
function getMaxBounds(items) {
    var bounds = getBounds(items[0]);
    for (var i = 1; i < items.length; i++) {
        var itemBounds = getBounds(items[i]);
        bounds[0] = Math.min(bounds[0], itemBounds[0]);
        bounds[1] = Math.max(bounds[1], itemBounds[1]);
        bounds[2] = Math.max(bounds[2], itemBounds[2]);
        bounds[3] = Math.min(bounds[3], itemBounds[3]);
    }
    return bounds;
}

/*
オブジェクトのバウンディングボックスを取得
Get bounding box of a single object
*/
function getBounds(item) {
    // visibleBounds を使用して境界を取得 / Always use visibleBounds to get bounds
    return item.visibleBounds;
}

main();
app.selectTool("Adobe Select Tool");