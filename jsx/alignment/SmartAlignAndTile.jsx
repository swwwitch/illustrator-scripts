#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartAlignAndTile.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択したオブジェクトを、行数または列数に基づいて自動整列・タイル配置
- 間隔・プレビュー境界・プレビュー更新に対応

### 主な機能：

- 行数／列数指定による整列
- 間隔（Gap）の設定
- プレビュー境界のON/OFF
- 即時プレビューとキャンセル時の復元
- ランダムボタンで重ね順をシャッフル

### 処理の流れ：

1. 選択オブジェクトを取得
2. ダイアログで整列条件を設定（方向、行数／列数、間隔）
3. 即時プレビューで整列状態を確認
4. ［OK］で確定、［キャンセル］で元に戻す
5. 「ランダム」ボタンで重ね順をランダム化して再配置

### オリジナル、謝辞：

Gorolib Design

### 更新履歴：

- v1.0 (20250406) : 初期バージョン
- v1.1 (20250407) : プレビュー境界・ランダム機能を追加
- v1.2 (20250731) : UIローカライズ、矢印キーでの値増減、ダイアログ位置・透明度調整を実装

### Script Name:

SmartAlignAndTile.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Automatically align and tile selected objects by rows or columns
- Supports spacing, preview bounds, and real-time preview

### Key Features:

- Alignment based on rows/columns
- Gap (spacing) control
- Toggle preview bounds
- Instant preview and restoration on cancel
- Randomize stacking order with the "Random" button

### Workflow:

1. Get selected objects
2. Configure alignment settings (direction, rows/columns, spacing) in dialog
3. Preview the layout in real time
4. Confirm with [OK], revert with [Cancel]
5. Randomize stacking order with "Random" button

### Original Idea / Credits:

Gorolib Design

### Change Log:

- v1.0 (20250406) : Initial version
- v1.1 (20250407) : Added preview bounds and randomization
- v1.2 (20250731) : Added localization, arrow key increments, dialog position and opacity
*/

var SCRIPT_VERSION = "v1.2";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "整列とタイル配置 " + SCRIPT_VERSION,
        en: "Smart Align and Tile " + SCRIPT_VERSION
    },
    rows: {
        ja: "行数：",
        en: "Rows:"
    },
    gap: {
        ja: "間隔：",
        en: "Gap:"
    },
    bounds: {
        ja: "プレビュー境界を使用",
        en: "Use Preview Bounds"
    },
    random: {
        ja: "ランダム",
        en: "Random"
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

main();

function main() {
    try {
        var selectedObjects = app.activeDocument.selection;
        if (selectedObjects.length === 0) {
            alert("オブジェクトが選択されていません。");
            return;
        }

        app.preferences.setBooleanPreference("includeStrokeInBounds", true);

        var originalPositions = [];
        for (var i = 0; i < selectedObjects.length; i++) {
            originalPositions.push([selectedObjects[i], selectedObjects[i].position.slice()]);
        }

        showAlignmentDialog(selectedObjects, originalPositions);

    } catch (error) {
        alert("エラーが発生しました：\r" + error.message);
    }
}

function showAlignmentDialog(selectedObjects, originalPositions) {
    var currentOrder = selectedObjects.slice();

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var inputGroup = dialog.add("group");
    inputGroup.orientation = "row";

    inputGroup.add("statictext", undefined, LABELS.rows[lang]);

    var rowControl = createSpinControl(inputGroup, "", 1, 1, 99, updatePreview, true, true);

    var unitSetting = app.activeDocument.rulerUnits;
    var unitLabel = getUnitLabel(unitSetting);

    var gapGroup = dialog.add("group");
    gapGroup.orientation = "row";

    gapGroup.add("statictext", undefined, LABELS.gap[lang]);
    var gapControl = createSpinControl(gapGroup, "", 0.0, -1000, 1000, updatePreview, false, false);
    gapGroup.add("statictext", undefined, "(" + unitLabel + ")");

    var boundsGroup = dialog.add("group");
    boundsGroup.orientation = "row";
    var boundsCheckbox = boundsGroup.add("checkbox", undefined, LABELS.bounds[lang]);
    boundsCheckbox.value = true;

    boundsCheckbox.onClick = function() {
        app.preferences.setBooleanPreference("includeStrokeInBounds", boundsCheckbox.value);
        updatePreview();
    };

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "fill";
    buttonGroup.alignChildren = ["left", "center"];

    var randomButton = buttonGroup.add("button", undefined, LABELS.random[lang]);
    var spacer = buttonGroup.add("statictext", undefined, "");
    spacer.characters = 1;
    var cancelButton = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okButton = buttonGroup.add("button", undefined, LABELS.ok[lang]);

    randomButton.onClick = function() {
        shuffleArray(currentOrder);
        updatePreview();
    };

    cancelButton.onClick = function() {
        restoreOriginalPositions(originalPositions);
        app.redraw();
        dialog.close(0); // Cancel pressed
    };

    okButton.onClick = function() {
        dialog.close(1); // OK pressed
    };

    rowControl.input.active = true;

    var offsetX = 300;
    var dialogOpacity = 0.97;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    dialog.show();

    function updatePreview() {
        restoreOriginalPositions(originalPositions);

        var gapValue = convertToPoints(gapControl.getValue(), unitSetting);
        var rowCount = rowControl.getValue();
        var totalItems = currentOrder.length;
        var fixedBounds = getObjectBounds(currentOrder[0]);

        if (rowCount > 1) {
            var colsAuto = Math.ceil(totalItems / rowCount);
            alignByGrid(currentOrder, gapValue, colsAuto, rowCount, true, fixedBounds);
        } else {
            alignHorizontally(currentOrder, gapValue, fixedBounds);
        }

        app.redraw();
    }
}

function alignHorizontally(objects, gap, fixedBounds) {
    var sorted = objects.slice(); // 並び順を維持 / Keep order
    var currentX = fixedBounds[0];
    var centerY = (fixedBounds[1] + fixedBounds[3]) / 2;

    for (var i = 0; i < sorted.length; i++) {
        var bounds = getObjectBounds(sorted[i]);
        var width = bounds[2] - bounds[0];
        var dx = currentX - bounds[0];
        var dy = centerY - (bounds[1] + bounds[3]) / 2;
        sorted[i].translate(dx, dy);
        currentX += width + gap;
    }
}

function alignByGrid(objects, gap, columnCount, rowCount, isHorizontal, fixedBounds) {
    var sorted = objects.slice(); // 並び順を維持 / Keep order
    var total = sorted.length;

    if (rowCount > 1 && columnCount <= 1) columnCount = Math.ceil(total / rowCount);
    if (columnCount > 1 && rowCount <= 1) rowCount = Math.ceil(total / columnCount);

    var boundsList = [], widths = [], heights = [];
    var maxHeight = 0;

    for (var i = 0; i < total; i++) {
        var b = getObjectBounds(sorted[i]);
        boundsList.push(b);
        var w = b[2] - b[0];
        var h = b[1] - b[3];
        widths.push(w);
        heights.push(h);
        if (h > maxHeight) maxHeight = h; // 最大の高さを記録 / Record max height
    }

    var colWidths = [];
    for (var c = 0; c < columnCount; c++) {
        var maxW = 0;
        for (var r = 0; r < rowCount; r++) {
            var idx = r * columnCount + c;
            if (idx >= total) break;
            if (widths[idx] > maxW) maxW = widths[idx];
        }
        colWidths[c] = maxW;
    }

    var startX = fixedBounds[0];
    var startY = fixedBounds[1];
    var index = 0;

    for (var r = 0; r < rowCount; r++) {
        var offsetX = 0;
        for (var c = 0; c < columnCount; c++) {
            if (index >= total) break;
            var dx = startX + offsetX - boundsList[index][0];
            var dy = startY - r * (maxHeight + gap) - boundsList[index][1];
            sorted[index].translate(dx, dy);
            offsetX += colWidths[c] + gap;
            index++;
        }
    }
}

function shuffleArray(array) {
    for (var i = array.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var tmp = array[i];
        array[i] = array[j];
        array[j] = tmp;
    }
}

function restoreOriginalPositions(positionArray) {
    for (var i = 0; i < positionArray.length; i++) {
        var obj = positionArray[i][0];
        var original = positionArray[i][1];
        var dx = original[0] - obj.position[0];
        var dy = original[1] - obj.position[1];
        obj.translate(dx, dy);
    }
}

function getObjectBounds(obj) {
    return app.preferences.getBooleanPreference("includeStrokeInBounds") ?
        obj.visibleBounds : obj.geometricBounds;
}

function createSpinControl(parentGroup, labelText, defaultValue, min, max, onChangeCallback, useInteger, step1Only) {
    var controlGroup = parentGroup.add("group");
    controlGroup.orientation = "row";
    controlGroup.alignChildren = "center";

    if (labelText && labelText !== "") {
        var label = controlGroup.add("statictext", undefined, labelText);
        label.characters = 6;
    }

    var inputField = controlGroup.add("edittext", undefined, defaultValue.toString());
    inputField.characters = 4;

    function getValue() {
        var val = parseFloat(inputField.text);
        if (isNaN(val)) val = useInteger ? 1 : 0;
        return Math.max(min, Math.min(max, useInteger ? Math.round(val) : val));
    }

    function setValue(val, callCallback) {
        if (isNaN(val)) val = useInteger ? 1 : 0;
        val = Math.max(min, Math.min(max, useInteger ? Math.round(val) : val));
        inputField.text = useInteger ? val.toString() : val.toFixed(1);
        if (callCallback && typeof onChangeCallback === "function") {
            onChangeCallback();
        }
    }

    inputField.onChange = function() {
        setValue(getValue(), true);
    };
    changeValueByArrowKey(inputField, step1Only);

    return {
        group: controlGroup,
        input: inputField,
        getValue: getValue,
        setValue: setValue
    };
}

function convertToPoints(value, unit) {
    switch (unit) {
        case RulerUnits.Millimeters:
            return value * 2.834645;
        case RulerUnits.Centimeters:
            return value * 28.34645;
        case RulerUnits.Inches:
            return value * 72;
        case RulerUnits.Picas:
            return value * 12;
        case RulerUnits.Points:
            return value;
        case RulerUnits.Pixels:
            return value;
        default:
            return value;
    }
}

function getUnitLabel(unit) {
    switch (unit) {
        case RulerUnits.Millimeters:
            return "mm";
        case RulerUnits.Centimeters:
            return "cm";
        case RulerUnits.Inches:
            return "inch";
        case RulerUnits.Picas:
            return "pica";
        case RulerUnits.Points:
            return "pt";
        case RulerUnits.Pixels:
            return "px";
        default:
            return "";
    }
}

function changeValueByArrowKey(editText, step1Only) {
    editText.addEventListener("keydown", function(event) {
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
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey && !step1Only) {
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
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey && !step1Only) {
            value = Math.round(value * 10) / 10;
        } else {
            value = Math.round(value);
        }

        editText.text = value;
        if (typeof editText.onChange === "function") {
            editText.onChange();
        }
    });
}

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