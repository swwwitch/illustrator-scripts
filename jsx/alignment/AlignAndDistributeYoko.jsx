#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

/*
### スクリプト名：

AlignAndDistributeYoko.jsx

### Readme （GitHub）：

https://github.com/your-repo/AlignAndDistributeYoko

### 概要：

- 選択したオブジェクトを横方向に整列し、指定した間隔と縦方向の数（行数）で再配置するスクリプト。
- プレビュー時に境界線を含むオプション、ランダム配置、単位の自動取得、上下キーでの数値変更に対応。

### 主な機能：

- 横方向整列と再配置
- 行数（縦方向の数）指定
- ランダム配置オプション
- プレビュー時の境界含む切替
- 単位自動対応
- キーボードで間隔・行数調整

### 処理の流れ：

- オブジェクト選択確認
- ダイアログ表示（各オプション設定）
- プレビュー更新
- 実行時に配置確定

### 更新履歴：

- v1.0 (20250716) : 初期バージョン
- v1.1 (20250717) : 安定性改善、行数ロジック修正
- v1.2 (20250718) : コメント整理、ローカライズ統一、ランダム基準位置補正改善

---

### Script Name:

AlignAndDistributeYoko.jsx

### Readme (GitHub):

https://github.com/your-repo/AlignAndDistributeYoko

### Overview:

- Arrange selected objects horizontally and re-distribute by specified spacing and number of rows.
- Supports preview bounds option, random arrangement, auto unit detection, and keyboard adjustments.

### Main Features:

- Horizontal arrangement and redistribution
- Row count (vertical count) specification
- Random arrangement option
- Toggle including preview bounds
- Automatic unit detection
- Keyboard adjustment for spacing and rows

### Workflow:

- Check object selection
- Show dialog and set options
- Preview update
- Confirm to apply

### Update History:

- v1.0 (2025-07-16): Initial version
- v1.1 (2025-07-17): Stability improvements, row logic fix
- v1.2 (2025-07-18): Comment refinement, localization update, improved random positioning correction
*/

// バージョン変数を追加 / Script version variable
var SCRIPT_VERSION = "v1.2";

// LABELS 定義 / Label definitions for localization
var LABELS = {
    dialogTitle: {
        ja: "ぴったり整列（横） " + SCRIPT_VERSION,
        en: "Align Precisely " + SCRIPT_VERSION
    },
    margin: {
        ja: "間隔:",
        en: "Spacing:"
    },
    useBounds: {
        ja: "プレビュー境界を使用",
        en: "Use preview bounds"
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
    },
    rows: {
        ja: "行数:",
        en: "Rows:"
    }
};

// 単位コードとラベルのマップ / Map of unit codes to labels
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

// 現在の単位ラベルを取得 / Get current unit label
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

// 単位コードからポイント換算係数を取得 / Get point factor from unit code
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0; // in
        case 1:
            return 72.0 / 25.4; // mm
        case 2:
            return 1.0; // pt
        case 3:
            return 12.0; // pica
        case 4:
            return 72.0 / 2.54; // cm
        case 5:
            return 72.0 / 25.4 * 0.25; // Q or H
        case 6:
            return 1.0; // px
        case 7:
            return 72.0 * 12.0; // ft/in
        case 8:
            return 72.0 / 25.4 * 1000.0; // m
        case 9:
            return 72.0 * 36.0; // yd
        case 10:
            return 72.0 * 12.0; // ft
        default:
            return 1.0;
    }
}

/* ダイアログ表示関数 / Show dialog with language support */
function showArrangeDialog() {
    var lang = getCurrentLang();
    var dlg = new Window("dialog", LABELS.dialogTitle);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    /* マージン入力UI: ラベルとテキストフィールドを横並びで配置 / Margin input UI: label and field side by side */
    var marginGroup = dlg.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = ["left", "center"];

    var unit = getCurrentUnitLabel();
    var marginLabel = marginGroup.add("statictext", undefined, LABELS.margin);
    var marginInput = marginGroup.add("edittext", undefined, "0");
    marginInput.characters = 3;
    changeValueByArrowKey(marginInput, true, updatePreview);
    var unitLabel = marginGroup.add("statictext", undefined, "(" + unit + ")");

    /* 行数入力UI: ラベルとテキストフィールドを横並びで配置 / Rows input UI: label and field side by side */
    var rowsGroup = dlg.add("group");
    rowsGroup.orientation = "row";
    rowsGroup.alignChildren = ["left", "center"];

    var rowsLabel = rowsGroup.add("statictext", undefined, LABELS.rows);
    var rowsInput = rowsGroup.add("edittext", undefined, "1");
    rowsInput.characters = 3;
    changeValueByArrowKey(rowsInput, true, updatePreview);

    /* 揃えパネル / Align panel */
    var alignPanel = dlg.add("panel", undefined, "揃え");
    alignPanel.orientation = "row";
    alignPanel.alignChildren = ["left", "center"];
    alignPanel.margins = [15, 20, 15, 10];

    var rbTop = alignPanel.add("radiobutton", undefined, "上");
    var rbMiddle = alignPanel.add("radiobutton", undefined, "中央");
    var rbBottom = alignPanel.add("radiobutton", undefined, "下");

    rbTop.value = true; // デフォルトを「上」に

    // 各ラジオボタンのクリック時にプレビューを即更新
    rbTop.onClick = function() {
        updatePreview();
    };
    rbMiddle.onClick = function() {
        updatePreview();
    };
    rbBottom.onClick = function() {
        updatePreview();
    };

    /* プレビュー境界チェックボックス / Preview bounds checkbox */
    var boundsCheckbox = dlg.add("checkbox", undefined, LABELS.useBounds);
    boundsCheckbox.value = true;
    boundsCheckbox.onClick = function() {
        updatePreview();
    };

    /* ランダム配置チェックボックス / Random arrangement checkbox */
    var randomCheckbox = dlg.add("checkbox", undefined, LABELS.random);
    randomCheckbox.value = false;

    /* ボタン配置グループ / Button group */
    var buttonGroup2 = dlg.add("group");
    buttonGroup2.alignment = "right";
    buttonGroup2.alignChildren = ["right", "center"];
    var cancelButton = buttonGroup2.add("button", undefined, LABELS.cancel, {
        name: "cancel"
    });
    var okButton = buttonGroup2.add("button", undefined, LABELS.ok, {
        name: "ok"
    });

    var originalSelection = activeDocument.selection.slice();
    var originalPositions = [];
    for (var i = 0; i < originalSelection.length; i++) {
        var item = originalSelection[i];
        originalPositions.push([item.left, item.top]);
    }

    /* プレビュー更新処理 / Update preview positioning */
    function updatePreview() {
        resetPositions(originalSelection, originalPositions);

        app.preferences.setBooleanPreference("includeStrokeInBounds", boundsCheckbox.value);

        var marginValue = parseFloat(marginInput.text);
        if (isNaN(marginValue)) marginValue = 0;

        var unitCode = app.preferences.getIntegerPreference("rulerType");
        var ptFactor = getPtFactorFromUnitCode(unitCode);
        var marginPt = marginValue * ptFactor;

        var mode = "horizontal"; // 固定 / fixed

        var sortedItems;
        if (randomCheckbox.value) {
            /* ランダム並び替え / Random sort */
            sortedItems = originalSelection.slice();
            for (var i = sortedItems.length - 1; i > 0; i--) {
                var j = Math.floor(Math.random() * (i + 1));
                var temp = sortedItems[i];
                sortedItems[i] = sortedItems[j];
                sortedItems[j] = temp;
            }
            // ランダム配置時の基準位置を固定する補正 / Fix base position for random layout
            var baseLeft = originalPositions[0][0];
            var baseTop = originalPositions[0][1];
            for (var i = 1; i < originalPositions.length; i++) {
                if (originalPositions[i][0] < baseLeft) baseLeft = originalPositions[i][0];
                if (originalPositions[i][1] > baseTop) baseTop = originalPositions[i][1];
            }
        } else {
            sortedItems = sortByX(originalSelection);
        }

        var rowsValue = parseInt(rowsInput.text, 10);
        if (isNaN(rowsValue) || rowsValue < 1) rowsValue = 1;

        var rows = rowsValue;
        if (rows < 1) rows = 1;
        var itemsPerRow = Math.ceil(sortedItems.length / rows);

        var startX = sortedItems[0].left;
        var startY = sortedItems[0].top;

        var align = "top";
        if (rbMiddle.value) align = "middle";
        else if (rbBottom.value) align = "bottom";

        var index = 0;
        for (var r = 0; r < rows; r++) {
            var currentX = startX;
            for (var c = 0; c < itemsPerRow && index < sortedItems.length; c++, index++) {
                var item = sortedItems[index];
                item.left = currentX;
                var baseY = startY;
                if (align === "middle") {
                    baseY = startY - (item.height / 2) + (sortedItems[0].height / 2);
                } else if (align === "bottom") {
                    baseY = startY + (sortedItems[0].height - item.height);
                }
                item.top = baseY - (r * (item.height + marginPt));
                if (c < itemsPerRow - 1) {
                    currentX += item.width + marginPt;
                }
            }
        }
        if (randomCheckbox.value) {
            // 並べ替え後の左上座標を取得して補正 / Adjust position after random sorting
            var minLeft = sortedItems[0].left;
            var maxTop = sortedItems[0].top;
            for (var i = 1; i < sortedItems.length; i++) {
                if (sortedItems[i].left < minLeft) minLeft = sortedItems[i].left;
                if (sortedItems[i].top > maxTop) maxTop = sortedItems[i].top;
            }
            var offsetX = baseLeft - minLeft;
            var offsetY = baseTop - maxTop;
            for (var i = 0; i < sortedItems.length; i++) {
                sortedItems[i].left += offsetX;
                sortedItems[i].top += offsetY;
            }
        }
        app.redraw();
    }

    updatePreview();

    marginInput.onChanging = function() {
        updatePreview();
    };
    randomCheckbox.onClick = function() {
        updatePreview();
    };

    /* ダイアログを開く前に marginInput をアクティブにする / Activate margin input before showing dialog */
    marginInput.active = true;
    if (dlg.show() !== 1) {
        resetPositions(originalSelection, originalPositions);
        app.redraw();
        return null;
    }

    var marginValue = parseFloat(marginInput.text);
    if (isNaN(marginValue)) marginValue = 0;
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var ptFactor = getPtFactorFromUnitCode(unitCode);
    var marginPt = marginValue * ptFactor;

    var rowsValue = parseInt(rowsInput.text, 10);
    if (isNaN(rowsValue) || rowsValue < 1) rowsValue = 1;

    var align = "top";
    if (rbMiddle.value) align = "middle";
    else if (rbBottom.value) align = "bottom";

    return {
        mode: "horizontal",
        margin: marginPt,
        rows: rowsValue,
        random: randomCheckbox.value,
        align: align
    };
}

/* 編集テキストで上下キーによる数値変更を有効化 / Enable up/down arrow key increment/decrement on edittext inputs */
function changeValueByArrowKey(editText, allowNegative, onUpdate) {
    editText.addEventListener("keydown", function(event) {
        if (editText.text.length === 0) return;
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            var isUp = event.keyName == "Up";
            var delta = 1;

            if (keyboard.shiftKey) {
                /* 10の倍数にスナップ / Snap to multiples of 10 */
                value = Math.floor(value / 10) * 10;
                delta = 10;
            }

            value += isUp ? delta : -delta;

            /* 負数許可されない場合は0未満を禁止 / Prevent negative if not allowed */
            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            if (typeof onUpdate === "function") {
                onUpdate();
            }
        }
    });
}

/* 位置を元に戻す / Reset positions to original */
function resetPositions(items, positions) {
    for (var i = 0; i < items.length; i++) {
        items[i].left = positions[i][0];
        items[i].top = positions[i][1];
    }
}

/* メイン処理 / Main function */
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

        var originalSelection = selectedItems.slice();

        var sortedItems;
        if (arrangeOptions.random) {
            sortedItems = originalSelection.slice();
            for (var i = sortedItems.length - 1; i > 0; i--) {
                var j = Math.floor(Math.random() * (i + 1));
                var temp = sortedItems[i];
                sortedItems[i] = sortedItems[j];
                sortedItems[j] = temp;
            }
            // ランダム配置時の基準位置を固定する補正 / Fix base position for random layout
            var baseLeft = originalSelection[0].left;
            var baseTop = originalSelection[0].top;
            for (var i = 1; i < originalSelection.length; i++) {
                if (originalSelection[i].left < baseLeft) baseLeft = originalSelection[i].left;
                if (originalSelection[i].top > baseTop) baseTop = originalSelection[i].top;
            }
        } else {
            sortedItems = sortByX(originalSelection);
        }

        var rows = arrangeOptions.rows;
        if (rows < 1) rows = 1;
        var itemsPerRow = Math.ceil(sortedItems.length / rows);

        var startX = sortedItems[0].left;
        var startY = sortedItems[0].top;

        var index = 0;
        for (var r = 0; r < rows; r++) {
            var currentX = startX;
            for (var c = 0; c < itemsPerRow && index < sortedItems.length; c++, index++) {
                var item = sortedItems[index];
                item.left = currentX;
                var baseY = startY;
                if (arrangeOptions.align === "middle") {
                    baseY = startY - (item.height / 2) + (sortedItems[0].height / 2);
                } else if (arrangeOptions.align === "bottom") {
                    baseY = startY + (sortedItems[0].height - item.height);
                }
                item.top = baseY - (r * (item.height + arrangeOptions.margin));
                if (c < itemsPerRow - 1) {
                    currentX += item.width + arrangeOptions.margin;
                }
            }
        }
        if (arrangeOptions.random) {
            // 並べ替え後の左上座標を取得して補正 / Adjust position after random sorting
            var minLeft = sortedItems[0].left;
            var maxTop = sortedItems[0].top;
            for (var i = 1; i < sortedItems.length; i++) {
                if (sortedItems[i].left < minLeft) minLeft = sortedItems[i].left;
                if (sortedItems[i].top > maxTop) maxTop = sortedItems[i].top;
            }
            var offsetX = baseLeft - minLeft;
            var offsetY = baseTop - maxTop;
            for (var i = 0; i < sortedItems.length; i++) {
                sortedItems[i].left += offsetX;
                sortedItems[i].top += offsetY;
            }
        }
    } catch (e) {
        alert("エラーが発生しました: " + e.message + "\nAn error has occurred.");
    }
}

main();

/* X座標でソート / Sort items by X coordinate */
function sortByX(items) {
    var copiedItems = items.slice();
    copiedItems.sort(function(a, b) {
        return a.left - b.left;
    });
    return copiedItems;
}