#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

FitArtboardWithMargin

### 概要

- 選択オブジェクトの外接範囲にマージンを加えたサイズで、アートボードを自動調整します。
- マージンは単位に応じた初期値が自動設定され、ダイアログで即時プレビュー確認が可能です。

### 主な機能

- 定規単位に基づくマージン初期値の自動設定
- 外接バウンディングボックス計算
- 即時プレビュー機能つきダイアログ
- アートボードサイズの自動更新

### 処理の流れ

1. 選択オブジェクトがない場合は全オブジェクトを対象にする
2. 定規単位に応じてマージン初期値を設定
3. ダイアログでマージンを指定（リアルタイムプレビュー）
4. バウンディングボックス計算とマージン加算
5. アートボードを更新

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

### 更新履歴

- v1.0.0 (20250420) : 初期バージョン
- v1.0.1 (20250708) : ポイント単位の初期値変更、UI改善

---

### Script Name:

FitArtboardToSelectedObjectsWithMargin.jsx

### Overview

- Automatically resizes the artboard to fit the bounding box of selected objects with an added margin.
- Margin value is auto-set based on the current ruler unit and can be previewed instantly in a dialog.

### Main Features

- Auto margin defaults based on ruler units
- Bounding box calculation
- Instant preview dialog
- Automatic artboard resizing

### Workflow

1. If no objects are selected, all objects are targeted
2. Set default margin value based on ruler units
3. Input margin in dialog (with live preview)
4. Calculate bounding box and add margin
5. Update artboard

### Changelog

- v1.0.0 (20250420) : Initial version
- v1.0.1 (20250708) : Changed default point value, UI improvements
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// ラベル定義（UI用） // UI label definitions
var LABELS = {
    marginLabel: {
        ja: "マージン",
        en: "Margin"
    },
    numberAlert: {
        ja: "数値を入力してください。",
        en: "Please enter a number."
    },
    dialogTitle: {
        ja: "マージン設定",
        en: "Margin Setup"
    }
};

/*
マージンダイアログを表示
Show margin input dialog (ScriptUI)
*/
function showMarginDialog(defaultValue, unit, lang) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var inputGroup = dlg.add("group");
    inputGroup.orientation = "row";
    inputGroup.alignChildren = ["left", "center"];

    var labelPre = LABELS.marginLabel[lang];
    var labelPost = " " + unit;

    inputGroup.add("statictext", undefined, labelPre + ":");
    var input = inputGroup.add("edittext", undefined, defaultValue);
    input.characters = 4;
    var postLabel = inputGroup.add("statictext", undefined, labelPost);
    postLabel.characters = 5;

    /*
    プレビューを更新
    Update preview
    */
    function updatePreview(value) {
        var previewValue = parseFloat(value);
        if (!isNaN(previewValue)) {
            var previewMarginInPoints = new UnitValue(previewValue, unit).as('pt');
            var previewBounds = getMaxBounds(app.activeDocument.selection.length === 0 ? app.activeDocument.pageItems : app.activeDocument.selection);
            previewBounds[0] -= previewMarginInPoints;
            previewBounds[1] += previewMarginInPoints;
            previewBounds[2] += previewMarginInPoints;
            previewBounds[3] -= previewMarginInPoints;

            // --- 座標と幅・高さを整数に丸める ---
            var x0 = Math.round(previewBounds[0]);
            var y1 = Math.round(previewBounds[1]);
            var x2 = Math.round(previewBounds[2]);
            var y3 = Math.round(previewBounds[3]);
            var width = Math.round(x2 - x0);
            var height = Math.round(y1 - y3);
            x2 = x0 + width;
            y3 = y1 - height;
            previewBounds[0] = x0;
            previewBounds[1] = y1;
            previewBounds[2] = x2;
            previewBounds[3] = y3;

            var abIndex = app.activeDocument.artboards.getActiveArtboardIndex();
            app.activeDocument.artboards[abIndex].artboardRect = previewBounds;
            app.redraw();
        }
    }

    input.onChanging = function() {
        updatePreview(input.text);
    };

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, "Cancel", {name: "cancel"});
    var okBtn = btnGroup.add("button", undefined, "OK", {name: "ok"});

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
        dlg.close();
    };

    // ダイアログが開いたときに即プレビュー
    // Apply initial preview when dialog opens
    updatePreview(defaultValue);

    dlg.show();
    return result;
}

/*
メイン処理
Main process
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

        // 単位ごとの初期マージン値設定
        // Set default margin value based on unit
        defaultMarginValue = '0';
        if (marginUnit === 'mm') {
            defaultMarginValue = '5';
        } else if (marginUnit === 'px') {
            defaultMarginValue = '20';
        } else if (marginUnit === 'pt') {
            defaultMarginValue = '10';
        }

        // ユーザーにマージンを入力させる（ScriptUI ダイアログ）
        // Show margin input dialog (ScriptUI)
        var userInput = showMarginDialog(defaultMarginValue, marginUnit, lang);
        if (!userInput) return;

        marginValue = parseFloat(userInput);
        marginInPoints = new UnitValue(marginValue, marginUnit).as('pt');

        // 選択範囲のバウンディングボックス計算
        // Calculate bounding box of selected items
        selectedBounds = getMaxBounds(selectedItems);
        selectedBounds[0] -= marginInPoints;
        selectedBounds[1] += marginInPoints;
        selectedBounds[2] += marginInPoints;
        selectedBounds[3] -= marginInPoints;

        // --- 座標と幅・高さを整数に丸める ---
        // 左上座標 (x0, y1)、右下座標 (x2, y3)
        var x0 = Math.round(selectedBounds[0]);
        var y1 = Math.round(selectedBounds[1]);
        var x2 = Math.round(selectedBounds[2]);
        var y3 = Math.round(selectedBounds[3]);
        // 幅と高さを整数に丸める
        var width = Math.round(x2 - x0);
        var height = Math.round(y1 - y3);
        // 右下座標を再計算
        x2 = x0 + width;
        y3 = y1 - height;
        // 更新後の座標
        selectedBounds[0] = x0;
        selectedBounds[1] = y1;
        selectedBounds[2] = x2;
        selectedBounds[3] = y3;

        // アートボードの更新
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
    // 常に visibleBounds を使用して境界を取得
    var bounds = item.visibleBounds;
    return bounds;
}

main();