
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
### スクリプト名：

AlignAndDistribute.jsx

### 概要

- Illustrator で選択したオブジェクトを「横」または「縦」に整列します。
- ダイアログでモードとマージンを設定でき、プレビューで即時確認可能です。

### 主な機能

- 横または縦にオブジェクトを整列
- マージン設定
- 即時プレビュー

### 処理の流れ

1. 選択オブジェクトを取得
2. ダイアログでモードとマージンを設定
3. 並び順にソート
4. 指定モード・マージンで配置

### 懸念事項

- 縦並びのとき、マージンの正負の値が逆です。

### 更新履歴

- v1.0 (20250711) : 初期バージョン
- v1.1 (20250711) : 日英ラベル対応、プレビュー改善

---

### Script Name:

AlignAndDistribute.jsx

### Overview

- Align selected Illustrator objects horizontally or vertically.
- Configure mode and margin via dialog, with real-time preview.

### Main Features

- Align objects horizontally or vertically
- Margin adjustment
- Live preview

### Process Flow

1. Get selected objects
2. Configure mode and margin in dialog
3. Sort objects
4. Align based on settings

### Update History

- v1.0 (20250711): Initial version
- v1.1 (20250711): Added Japanese/English labels, improved preview
*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

// 単位コードとラベルのマップ / Unit code to label map
var unitLabelMap = {
    0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm",
    5: "Q/H", 6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft"
};

// 現在の単位ラベルを取得 / Get current document unit label
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

// 言語判定関数 / Function to get current language
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// UIラベル定義 / UI Label Definitions
var LABELS = {
    dialogTitle: {
        ja: "ぴったり整列 " + SCRIPT_VERSION,
        en: "Align Precisely " + SCRIPT_VERSION
    },
    panelMode: { ja: "並び", en: "Mode" },
    none: { ja: "なし", en: "None" },
    horizontal: { ja: "横", en: "Horizontal" },
    vertical: { ja: "縦", en: "Vertical" },
    margin: { ja: "マージン", en: "Margin" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
};


// 選択オブジェクトの位置を元に戻す関数
function resetPositions(selection, originalPositions) {
    for (var i = 0; i < selection.length; i++) {
        selection[i].left = originalPositions[i][0];
        selection[i].top = originalPositions[i][1];
    }
}


// ソート関数群 / Sorting functions
function sortByX(items) {
    var copiedItems = [];
    for (var i = 0; i < items.length; i++) {
        copiedItems.push(items[i]);
    }
    copiedItems.sort(function(a, b) {
        return a.left - b.left;
    });
    return copiedItems;
}

function sortByY(items) {
    var copiedItems = [];
    for (var i = 0; i < items.length; i++) {
        copiedItems.push(items[i]);
    }
    copiedItems.sort(function(a, b) {
        return b.top - a.top;
    });
    return copiedItems;
}

// --- Alignment and move utility functions ---
function moveNextToRight(baseItem, targetItem, margin) {
    var baseBounds = getBounds(baseItem);
    var targetBounds = getBounds(targetItem);
    var moveX = baseBounds[2] - targetBounds[0] + margin;
    targetItem.translate(moveX, 0);
}



function moveBelow(baseItem, targetItem, margin) {
    var baseBounds = getBounds(baseItem);
    var targetBounds = getBounds(targetItem);
    // Arrange downward relative to the base's bottom edge
    // Positive margin moves further downward (increases separation)
    var moveY = baseBounds[3] - targetBounds[1] + margin;
    targetItem.translate(0, moveY);
}

function alignX(baseItem, targetItem) {
    var baseBounds = getBounds(baseItem);
    var targetBounds = getBounds(targetItem);
    var diffX = baseBounds[0] - targetBounds[0];
    targetItem.translate(diffX, 0);
}

function alignY(baseItem, targetItem) {
    var baseBounds = getBounds(baseItem);
    var targetBounds = getBounds(targetItem);
    var diffY = baseBounds[1] - targetBounds[1];
    targetItem.translate(0, diffY);
}

/**
 * Get the visible bounds of an item.
 * For clipping groups, return the mask path's visibleBounds if available,
 * otherwise fall back to the group's visibleBounds.
 * For other items, return their visibleBounds directly.
 */
function getBounds(item) {
    // If item is a clipping group, try to use the mask path's visibleBounds
    if (item.clipped && item.pageItems && item.pageItems.length > 0) {
        for (var i = 0; i < item.pageItems.length; i++) {
            var child = item.pageItems[i];
            if (child.clipping) {
                // Found the mask path
                return child.visibleBounds;
            }
        }
        // No mask path found, fallback to group's visibleBounds
        return item.visibleBounds;
    }
    // For regular items, return their visibleBounds
    return item.visibleBounds;
}

function main() {
    try {
        var selectedItems = activeDocument.selection;

        if (!selectedItems || selectedItems.length === 0) {
            alert("オブジェクトを選択してください。\nPlease select objects.");
            return;
        }

        var arrangeOptions = showArrangeDialog();
        if (!arrangeOptions) return;

        app.preferences.setBooleanPreference("includeStrokeInBounds", true);

        var sortedItems = (arrangeOptions.mode === "horizontal") ? sortByX(selectedItems) : sortByY(selectedItems);

        for (var i = 0; i < sortedItems.length - 1; i++) {
            if (arrangeOptions.mode === "horizontal") {
                moveNextToRight(sortedItems[i], sortedItems[i + 1], arrangeOptions.margin);
                alignY(sortedItems[i], sortedItems[i + 1]);
            } else {
                moveBelow(sortedItems[i], sortedItems[i + 1], arrangeOptions.margin);
                alignX(sortedItems[i], sortedItems[i + 1]);
            }
        }
    } catch (e) {
        alert("エラーが発生しました: " + e.message + "\nAn error has occurred.");
    }
}

// ダイアログ表示関数 / Show dialog with language support
function showArrangeDialog() {
    var lang = getCurrentLang();
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var modePanel = dlg.add("panel", undefined, LABELS.panelMode[lang]);
    modePanel.orientation = "row";
    modePanel.alignChildren = ["center", "center"]; // 中央揃えに設定
    modePanel.margins = [15, 20, 15, 10];
    // モードパネルから「なし」オプションを削除
    var rbHorizontal = modePanel.add("radiobutton", undefined, LABELS.horizontal[lang]);
    var rbVertical = modePanel.add("radiobutton", undefined, LABELS.vertical[lang]);
    rbHorizontal.value = true;

    var unit = getCurrentUnitLabel();
    var marginPanel = dlg.add("panel", undefined, LABELS.margin[lang] + " (" + unit + ")");
    marginPanel.orientation = "row";
    marginPanel.alignChildren = ["center", "top"];
    marginPanel.margins = [15, 20, 15, 10];

    var marginInput = marginPanel.add("edittext", undefined, "0");
    marginInput.characters = 5;

    var buttonGroup = marginPanel.add("group");
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

    plusBtn.onClick = function () {
        var val = parseFloat(marginInput.text);
        if (isNaN(val)) val = 0;
        marginInput.text = (val + 1).toString();
        updatePreview();
    };
    minusBtn.onClick = function () {
        var val = parseFloat(marginInput.text);
        if (isNaN(val)) val = 0;
        marginInput.text = (val - 1).toString();
        updatePreview();
    };
    zeroBtn.onClick = function () {
        marginInput.text = "0";
        updatePreview();
    };

    rbHorizontal.onClick = function () { updatePreview(); };
    rbVertical.onClick = function () { updatePreview(); };

    var buttonGroup2 = dlg.add("group");
    buttonGroup2.alignment = "right";
    buttonGroup2.alignChildren = ["right", "center"];
    var cancelButton = buttonGroup2.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var okButton = buttonGroup2.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    var originalPositions = [];
    for (var i = 0; i < activeDocument.selection.length; i++) {
        var item = activeDocument.selection[i];
        originalPositions.push([item.left, item.top]);
    }

    function updatePreview() {
        resetPositions(activeDocument.selection, originalPositions);

        var marginValue = parseFloat(marginInput.text);
        if (isNaN(marginValue)) marginValue = 0;

        var mode = rbHorizontal.value ? "horizontal" : "vertical";

        var sortedItems = (mode === "horizontal") ? sortByX(activeDocument.selection) : sortByY(activeDocument.selection);

        for (var i = 0; i < sortedItems.length - 1; i++) {
            if (mode === "horizontal") {
                moveNextToRight(sortedItems[i], sortedItems[i + 1], marginValue);
                alignY(sortedItems[i], sortedItems[i + 1]);
            } else {
                moveBelow(sortedItems[i], sortedItems[i + 1], marginValue);
                alignX(sortedItems[i], sortedItems[i + 1]);
            }
        }
        app.redraw();
    }

    marginInput.onChanging = function () {
        updatePreview();
    };

    if (dlg.show() !== 1) {
        resetPositions(activeDocument.selection, originalPositions);
        app.redraw();
        return null;
    }

    var marginValue = parseFloat(marginInput.text);
    if (isNaN(marginValue)) marginValue = 0;

    var mode = rbHorizontal.value ? "horizontal" : "vertical";

    return { mode: mode, margin: marginValue };
}

main();
