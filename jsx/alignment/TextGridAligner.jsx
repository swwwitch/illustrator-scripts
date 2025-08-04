#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

TextGridAligner-v2.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/TextGridAligner.jsx

### 概要：

- Illustratorでテキストフレームを行・列単位で整列またはグループ化するスクリプト
- 行方向と列方向のしきい値を独立して調整可能

### 主な機能：

- 行方向：天地中央に整列、またはグループ化
- 列方向：左右中央に整列、またはグループ化
- 行・列のアキを均等に配置するオプション

### 処理の流れ：

1. ダイアログを表示して設定を取得
2. 選択したテキストフレームを対象に処理
3. 行・列ごとに整列またはグループ化を実行

### 更新履歴：

- v1.0 (20250802) : 初期バージョン
- v1.1 (20250803) : ダイアログUI改善とコード整理
*/

/*
### Script Name:

TextGridAligner-v2.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/TextGridAligner.jsx

### Overview:

- Script for aligning or grouping text frames in Illustrator by rows or columns
- Independent threshold adjustment for row and column directions

### Key Features:

- Row: Align or group to vertical center
- Column: Align or group to horizontal center
- Option to distribute spacing evenly for rows or columns

### Process Flow:

1. Show dialog to get settings
2. Process selected text frames
3. Execute alignment or grouping by row/column

### Update History:

- v1.0 (20250802) : Initial version
- v1.1 (20250803) : Improved dialog UI and code organization
*/

var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

/* 日英ラベル定義（UI表示順）/ Japanese & English label definitions (UI display order) */
var LABELS = {
    dialogTitle: {
        ja: "テキスト整列・グループ化 " + SCRIPT_VERSION,
        en: "Text Alignment & Grouping " + SCRIPT_VERSION
    },
    threshold: {
        ja: "行",
        en: "Row"
    },
    row: {
        ja: "揃え",
        en: "Align"
    },
    rowGroup: {
        ja: "行をグループ化",
        en: "Group Rows"
    },
    equalSpacing: {
        ja: "アキを均等に",
        en: "Distribute Evenly"
    },
    colThreshold: {
        ja: "列",
        en: "Column"
    },
    column: {
        ja: "揃え",
        en: "Align"
    },
    colGroup: {
        ja: "列をグループ化",
        en: "Group Columns"
    },
    colEqualSpacing: {
        ja: "アキを均等に",
        en: "Distribute Evenly"
    },
    group: {
        ja: "実行",
        en: "Run"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    resultMessage: {
        ja: "○個のグループ化を行いました。",
        en: " groups have been created."
    }
};

var rowOverlapThreshold, colOverlapThreshold;

/* グループ間のアキを均等に配置する / Distribute spacing evenly between groups
   direction: "horizontal" or "vertical" */
function distributeSpacingBetweenGroups(groups, direction) {
    if (!groups || groups.length <= 1) return;

    if (direction === "horizontal") {
        // 左から右にソート / Sort left to right
        groups.sort(function(a, b) {
            return a.geometricBounds[0] - b.geometricBounds[0];
        });

        var allBounds = getCombinedBounds(groups);
        var totalWidth = 0;
        for (var i = 0; i < groups.length; i++) {
            var b = groups[i].geometricBounds;
            totalWidth += (b[2] - b[0]);
        }
        var availableSpace = allBounds[2] - allBounds[0];
        var spacing = (availableSpace - totalWidth) / (groups.length - 1);

        var currentLeft = groups[0].geometricBounds[0];
        for (var i = 0; i < groups.length; i++) {
            var b = groups[i].geometricBounds;
            var itemWidth = b[2] - b[0];
            groups[i].left = currentLeft;
            currentLeft += itemWidth + spacing;
        }
    } else if (direction === "vertical") {
        // 上から下にソート / Sort top to bottom
        groups.sort(function(a, b) {
            return b.geometricBounds[1] - a.geometricBounds[1];
        });

        var allBounds = getCombinedBounds(groups);
        var totalHeight = 0;
        for (var i = 0; i < groups.length; i++) {
            var b = groups[i].geometricBounds;
            totalHeight += (b[1] - b[3]);
        }
        var availableSpace = allBounds[1] - allBounds[3];
        var spacing = (availableSpace - totalHeight) / (groups.length - 1);

        var currentTop = groups[0].geometricBounds[1];
        for (var i = 0; i < groups.length; i++) {
            var b = groups[i].geometricBounds;
            var itemHeight = b[1] - b[3];
            groups[i].top = currentTop;
            currentTop -= itemHeight + spacing;
        }
    }
}

/* ダイアログUIの表示とユーザー選択取得 / Show dialog UI and get user selections */
function showDialog(prevThreshold, prevGroupMode) {
    /* ダイアログを開く前に元の位置を保存 / Save original positions before showing dialog */
    var originalStates = [];
    var selection = app.activeDocument.selection;
    var textFrames = [];
    if (selection && selection.length > 0) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") {
                textFrames.push(selection[i]);
                originalStates.push({
                    item: selection[i],
                    bounds: selection[i].geometricBounds.slice()
                });
            }
        }
    }

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    /* ダイアログの位置と透明度 / Dialog position and opacity */
    var offsetX = 300;
    var dialogOpacity = 0.95;

    /* ダイアログ用ヘルパー関数 / Helper functions for dialog */
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



    /* 行しきい値パネル / Row Threshold Panel */
    var thresholdGroup = dialog.add("panel", undefined, LABELS.threshold[lang]);
    thresholdGroup.orientation = "column";
    thresholdGroup.alignChildren = "left";
    thresholdGroup.margins = [15, 20, 15, 10];
    thresholdGroup.enabled = true;

    /* 行揃えチェックボックス / Align rows checkbox */
    var rowCheck = thresholdGroup.add("checkbox", undefined, LABELS.row[lang]);
    rowCheck.value = true;

    /* 行グループ化チェックボックス / Group rows checkbox */
    var rowGroupCheck = thresholdGroup.add("checkbox", undefined, LABELS.rowGroup[lang]);
    rowGroupCheck.value = false;

    /* 行アキ均等チェックボックス / Equal spacing checkbox for rows */
    var equalSpacingCheck = thresholdGroup.add("checkbox", undefined, LABELS.equalSpacing[lang]);
    equalSpacingCheck.value = false;
    equalSpacingCheck.enabled = false; // 初期はディム表示 / Initially dimmed
    rowGroupCheck.onClick = function() {
        if (rowGroupCheck.value) {
            colGroupCheck.value = false;
            equalSpacingCheck.enabled = true; // 有効化
        } else {
            equalSpacingCheck.enabled = false; // 無効化
            equalSpacingCheck.value = false;
        }
    };

    /* 選択全体の幅と高さを取得 / Get total width and height of selection */
    var combinedBounds = getCombinedBounds(textFrames);
    var totalWidth = combinedBounds[2] - combinedBounds[0];
    var totalHeight = combinedBounds[1] - combinedBounds[3];

    /* 行しきい値スライダー / Row threshold slider */
    var thresholdSlider = thresholdGroup.add("slider", undefined, 10, 0, totalWidth || 100);
    thresholdSlider.value = Math.min(prevThreshold, totalWidth || 100);
    thresholdSlider.preferredSize.width = 150;

    /* 行しきい値ラベル / Row threshold label */
    var thresholdLabel = thresholdGroup.add("statictext", undefined, Math.round(thresholdSlider.value) + " pt");
    thresholdLabel.alignment = "center";
    thresholdLabel.characters = 5;

    // 行揃えチェックボックスのON/OFFでスライダーとラベルをディム/有効
    rowCheck.onClick = function() {
        thresholdSlider.enabled = rowCheck.value;
        thresholdLabel.enabled = rowCheck.value;
        rowGroupCheck.enabled = rowCheck.value; // 行揃えがOFFならグループ化もディム
        if (!rowCheck.value) {
            rowGroupCheck.value = false;
            equalSpacingCheck.enabled = false;
            equalSpacingCheck.value = false;
        }
    };
    // 初期状態
    thresholdSlider.enabled = rowCheck.value;
    thresholdLabel.enabled = rowCheck.value;
    rowGroupCheck.enabled = rowCheck.value;

    /* 列しきい値パネル / Column Threshold Panel */
    var colThresholdGroup = dialog.add("panel", undefined, LABELS.colThreshold[lang]);
    colThresholdGroup.orientation = "column";
    colThresholdGroup.alignChildren = "left";
    colThresholdGroup.margins = [15, 20, 15, 10];
    colThresholdGroup.enabled = true;

    /* 列揃えチェックボックス / Align columns checkbox */
    var colCheck = colThresholdGroup.add("checkbox", undefined, LABELS.column[lang]);
    colCheck.value = true;

    /* 列グループ化チェックボックス / Group columns checkbox */
    var colGroupCheck = colThresholdGroup.add("checkbox", undefined, LABELS.colGroup[lang]);
    colGroupCheck.value = false;

    /* 列アキ均等チェックボックス / Equal spacing checkbox for columns */
    var colEqualSpacingCheck = colThresholdGroup.add("checkbox", undefined, LABELS.colEqualSpacing[lang]);
    colEqualSpacingCheck.value = false;
    colEqualSpacingCheck.enabled = false; // 初期はディム表示 / Initially dimmed
    colGroupCheck.onClick = function() {
        if (colGroupCheck.value) {
            rowGroupCheck.value = false;
            colEqualSpacingCheck.enabled = true; // 有効化
        } else {
            colEqualSpacingCheck.enabled = false; // 無効化
            colEqualSpacingCheck.value = false;
        }
    };

    /* 列しきい値スライダー / Column threshold slider */
    var colThresholdSlider = colThresholdGroup.add("slider", undefined, 10, 0, totalHeight || 100);
    colThresholdSlider.value = Math.min(prevThreshold, totalHeight || 100);
    colThresholdSlider.preferredSize.width = 150;
    /* 列しきい値ラベル / Column threshold label */
    var colThresholdLabel = colThresholdGroup.add("statictext", undefined, Math.round(colThresholdSlider.value) + " pt");
    colThresholdLabel.alignment = "center";
    colThresholdLabel.characters = 5;

    // 列揃えチェックボックスのON/OFFでスライダーとラベルをディム/有効
    colCheck.onClick = function() {
        colThresholdSlider.enabled = colCheck.value;
        colThresholdLabel.enabled = colCheck.value;
        colGroupCheck.enabled = colCheck.value; // 列揃えがOFFならグループ化もディム
        if (!colCheck.value) {
            colGroupCheck.value = false;
            colEqualSpacingCheck.enabled = false;
            colEqualSpacingCheck.value = false;
        }
    };
    // 初期状態
    colThresholdSlider.enabled = colCheck.value;
    colThresholdLabel.enabled = colCheck.value;
    colGroupCheck.enabled = colCheck.value;

    thresholdSlider.onChanging = function() {
        thresholdLabel.text = Math.round(thresholdSlider.value) + " pt";
        rowOverlapThreshold = thresholdSlider.value; // 即時更新
        if (textFrames && textFrames.length > 0 && rowCheck.value) {
            if (rowGroupCheck.value) {
                groupOverlappingObjectsByDirection("horizontal");
            } else {
                alignHorizontallyAndCenterVertically(textFrames);
            }
            app.redraw();
        }
    };

    colThresholdSlider.onChanging = function() {
        colThresholdLabel.text = Math.round(colThresholdSlider.value) + " pt";
        colOverlapThreshold = colThresholdSlider.value; // 即時更新
        if (textFrames && textFrames.length > 0 && colCheck.value) {
            if (colGroupCheck.value) {
                // 列をグループ化のみ。アキを均等にがONの場合のみ均等処理も行う / Only group columns; distribute spacing if enabled
                var newGroups = groupObjectsOnlyByDirection("vertical");
                if (colEqualSpacingCheck.value && newGroups && newGroups.length > 0) {
                    // グループごとに左右のアキを均等にする / Distribute horizontal spacing evenly
                    for (var g = 0; g < newGroups.length; g++) {
                        var groupItems = newGroups[g].pageItems;
                        if (!groupItems || groupItems.length <= 1) continue;

                        var itemsArray = [];
                        for (var i = 0; i < groupItems.length; i++) {
                            if (groupItems[i].typename === "TextFrame") {
                                itemsArray.push(groupItems[i]);
                            }
                        }
                        if (itemsArray.length <= 1) continue;

                        // 左から右にソート / Sort left to right
                        itemsArray.sort(function(a, b) {
                            return a.geometricBounds[0] - b.geometricBounds[0];
                        });

                        var groupBounds = getCombinedBounds(itemsArray);
                        var totalWidth = 0;
                        for (var i = 0; i < itemsArray.length; i++) {
                            var b = itemsArray[i].geometricBounds;
                            totalWidth += (b[2] - b[0]);
                        }
                        var availableSpace = groupBounds[2] - groupBounds[0];
                        var spacing = (availableSpace - totalWidth) / (itemsArray.length - 1);

                        var currentLeft = itemsArray[0].geometricBounds[0];
                        for (var i = 0; i < itemsArray.length; i++) {
                            var b = itemsArray[i].geometricBounds;
                            var itemWidth = b[2] - b[0];
                            itemsArray[i].left = currentLeft;
                            currentLeft += itemWidth + spacing;
                        }
                    }
                }
            } else {
                alignVerticallyAndCenterHorizontally(textFrames);
            }
            app.redraw();
        }
    };

    /* ボタングループ / Button group */
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    buttonGroup.margins = [0, 10, 0, 10];
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.group[lang], {
        name: "ok"
    });

    /* キャンセルボタン処理 / Cancel button handler */
    cancelBtn.onClick = function() {
        for (var i = 0; i < originalStates.length; i++) {
            var obj = originalStates[i];
            var b = obj.bounds;
            var item = obj.item;
            item.left = b[0];
            item.top = b[1];
        }
        dialog.close(0);
    };

    /* OKボタン処理 / OK button handler */
    okBtn.onClick = function() {
        // groupMode and threshold values are set on dialog close
        groupMode = {
            row: rowCheck.value,
            rowGroup: rowGroupCheck.value,
            equalSpacing: equalSpacingCheck.value,
            column: colCheck.value,
            colGroup: colGroupCheck.value,
            colEqualSpacing: colEqualSpacingCheck.value
        };
        dialog.close(1);
    };

    /* 値を確定するのはダイアログを閉じるとき / Commit values on dialog close */
    dialog.onClose = function() {
        rowOverlapThreshold = thresholdSlider.value;
        colOverlapThreshold = colThresholdSlider.value;
    };

    /* ダイアログの不透明度と位置を設定 / Set opacity and position before showing the dialog */
    dialog.opacity = dialogOpacity;
    dialog.onShow = function() {
        var currentX = dialog.location[0];
        var currentY = dialog.location[1];
        dialog.location = [currentX + offsetX, currentY + 0];
    };

    var result = dialog.show();
    return result;
}

/* 行方向のDFS探索 / DFS for horizontal grouping */
function dfsHorizontal(index, items, visited, group) {
    visited[index] = true;
    group.push(items[index]);
    var boundsA = items[index].visibleBounds;
    for (var j = 0; j < items.length; j++) {
        if (visited[j]) continue;
        var boundsB = items[j].visibleBounds;
        var overlapRatio = getOverlapRatio(boundsA, boundsB);
        var axisDistanceOk = getAxisDistanceHorizontal(boundsA, boundsB) <= rowOverlapThreshold;
        if (overlapRatio > 0 || axisDistanceOk) {
            dfsHorizontal(j, items, visited, group);
        }
    }
}

/* 列方向のDFS探索 / DFS for vertical grouping */
function dfsVertical(index, items, visited, group) {
    visited[index] = true;
    group.push(items[index]);
    var boundsA = items[index].visibleBounds;
    for (var j = 0; j < items.length; j++) {
        if (visited[j]) continue;
        var boundsB = items[j].visibleBounds;
        var overlapRatio = getOverlapRatio(boundsA, boundsB);
        var axisDistanceOk = getAxisDistanceVertical(boundsA, boundsB) <= colOverlapThreshold;
        if (overlapRatio > 0 || axisDistanceOk) {
            dfsVertical(j, items, visited, group);
        }
    }
}

/* 隣接度または重なり率に基づいてグループを抽出（方向別）/ Extract groups based on adjacency or overlap by direction */
function getGroupedOverlappingItems(items, direction) {
    var groups = [];
    var visited = [];
    for (var i = 0; i < items.length; i++) {
        visited[i] = false;
    }
    for (var i = 0; i < items.length; i++) {
        if (visited[i]) continue;
        var group = [];
        if (direction === "horizontal") {
            dfsHorizontal(i, items, visited, group);
        } else if (direction === "vertical") {
            dfsVertical(i, items, visited, group);
        }
        groups.push(group);
    }
    return groups;
}


/* 2つのバウンディングボックスの水平方向距離を返す（縦方向距離がある場合は無視）/ Return horizontal distance between two bounding boxes (ignore if vertical distance exists) */
function getAxisDistanceHorizontal(a, b) {
    var ax1 = a[0],
        ay1 = a[1],
        ax2 = a[2],
        ay2 = a[3];
    var bx1 = b[0],
        by1 = b[1],
        bx2 = b[2],
        by2 = b[3];

    var vertGap = Math.max(0, Math.max(by2 - ay1, ay2 - by1));
    if (vertGap > 0) return 999999; // 縦方向に離れている場合は無視

    var horzGap = Math.max(0, Math.max(bx1 - ax2, ax1 - bx2));
    return horzGap;
}

/* 2つのバウンディングボックスの垂直方向距離を返す（横方向距離がある場合は無視）/ Return vertical distance between two bounding boxes (ignore if horizontal distance exists) */
function getAxisDistanceVertical(a, b) {
    var ax1 = a[0],
        ay1 = a[1],
        ax2 = a[2],
        ay2 = a[3];
    var bx1 = b[0],
        by1 = b[1],
        bx2 = b[2],
        by2 = b[3];

    var horzGap = Math.max(0, Math.max(bx1 - ax2, ax1 - bx2));
    if (horzGap > 0) return 999999; // 横方向に離れている場合は無視

    var vertGap = Math.max(0, Math.max(by2 - ay1, ay2 - by1));
    return vertGap;
}

/* 2つのバウンディングボックスの重なり率（大きい方の面積に対する割合）を返す / Return overlap ratio of two bounding boxes (relative to larger area) */
function getOverlapRatio(a, b) {
    var ax = Math.max(0, Math.min(a[2], b[2]) - Math.max(a[0], b[0]));
    var ay = Math.max(0, Math.min(a[1], b[1]) - Math.max(a[3], b[3]));
    var overlapArea = ax * ay;
    if (overlapArea <= 0) return 0;
    var areaA = (a[2] - a[0]) * (a[1] - a[3]);
    var areaB = (b[2] - b[0]) * (b[1] - b[3]);
    var maxArea = Math.max(areaA, areaB);
    return overlapArea / maxArea;
}


/* 共通グループ化処理（direction: "horizontal"|"vertical"）/ Common grouping process (direction: "horizontal"|"vertical") */
function groupOverlappingObjectsByDirection(direction) {
    if (!app.documents.length) return;
    var items = app.activeDocument.selection;
    if (!items || items.length === 0) return;

    var groups = getGroupedOverlappingItems(items, direction);

    var doc = app.activeDocument;
    var newGroups = [];
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;

        /* グループ化前に元のレイヤーを保存 / Save the original layer before grouping */
        var originalLayer = group[0].layer;

        /* 垂直方向中央揃え / Align vertically centered */
        var groupBounds = getCombinedBounds(group);
        var centerY = (groupBounds[1] + groupBounds[3]) / 2;
        for (var j = 0; j < group.length; j++) {
            var itemBounds = group[j].geometricBounds;
            var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;
            var shiftY = centerY - itemCenterY;
            group[j].top += shiftY;
        }

        app.executeMenuCommand('deselectall');
        for (var j = 0; j < group.length; j++) {
            group[j].selected = true;
        }
        app.executeMenuCommand('group');
        var newGroup = app.activeDocument.selection[0];
        /* 新規グループのレイヤーを元に戻す / Restore the new group's layer to the original layer */
        newGroup.layer = originalLayer;
        newGroups.push(newGroup);
    }
    app.redraw();
    /* 新規グループを選択状態に設定 / Set new groups as selected */
    app.activeDocument.selection = null;
    for (var i = 0; i < newGroups.length; i++) {
        newGroups[i].selected = true;
    }
    return newGroups;
}

/* 横方向ごとに天地中央に整列（グループ化は行わない）/ Align to vertical center by horizontal group (no grouping) */
function alignHorizontallyAndCenterVertically(items) {
    if (!items || items.length === 0) return;
    var groups = getGroupedOverlappingItems(items, "horizontal");
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;
        var combinedBounds = getCombinedBounds(group);
        var centerY = (combinedBounds[1] + combinedBounds[3]) / 2;
        for (var j = 0; j < group.length; j++) {
            var itemBounds = group[j].geometricBounds;
            var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;
            var shiftY = centerY - itemCenterY;
            group[j].top += shiftY;
        }
    }
}

/* 縦方向ごとに中央揃え（左右中央揃え）/ Align to horizontal center by vertical group */
function alignVerticallyAndCenterHorizontally(items) {
    if (!items || items.length === 0) return;
    var groups = getGroupedOverlappingItems(items, "vertical");
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;
        var combinedBounds = getCombinedBounds(group);
        var centerX = (combinedBounds[0] + combinedBounds[2]) / 2;
        for (var j = 0; j < group.length; j++) {
            var itemBounds = group[j].geometricBounds;
            var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
            var shiftX = centerX - itemCenterX;
            group[j].left += shiftX;
        }
    }
}

/* メイン処理 / Main process */
function main() {
    rowOverlapThreshold = 10;
    if (!app.documents.length) return;
    var selection = app.activeDocument.selection;
    if (!selection || selection.length === 0) return;

    /* テキストフレームのみを対象にする / Only target text frames */
    var textFrames = [];
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "TextFrame") {
            textFrames.push(selection[i]);
        }
    }
    if (textFrames.length === 0) {
        /* テキストが選択されていない場合の警告 / Alert if no text frames selected */
        alert((lang === "ja") ? "テキストが選択されていません。" : "No text frames selected.");
        return;
    }

    var result = showDialog(rowOverlapThreshold, undefined);
    if (result !== 1) return;

    // rowOverlapThreshold, colOverlapThreshold は showDialog() 内で更新される / Updated inside showDialog()
    if (groupMode.row) {
        if (groupMode.rowGroup) {
            /* 行をグループ化 / Group rows */
            var newGroups = groupOverlappingObjectsByDirection("horizontal");

            if (groupMode.equalSpacing && newGroups && newGroups.length > 1) {
                distributeSpacingBetweenGroups(newGroups, "vertical");
            }
        } else {
            /* 行ごとに天地中央揃え / Align to vertical center by row */
            alignHorizontallyAndCenterVertically(textFrames);
        }
    }
    if (groupMode.column) {
        if (groupMode.colGroup) {
            /* 列をグループ化 / Group columns */
            var newGroups = groupObjectsOnlyByDirection("vertical");

            /* 各グループ内のアイテム調整は行わない / Do not adjust items inside each group */

            if (groupMode.colEqualSpacing && newGroups && newGroups.length > 1) {
                distributeSpacingBetweenGroups(newGroups, "horizontal");
            }
        } else {
            /* 列ごとに左右中央揃え / Align to horizontal center by column */
            alignVerticallyAndCenterHorizontally(textFrames);
        }
    }
}

/* 列方向グループ化のみ（整列なし）/ Only group by column direction (no alignment) */
function groupObjectsOnlyByDirection(direction) {
    if (!app.documents.length) return;
    var items = app.activeDocument.selection;
    if (!items || items.length === 0) return;

    var groups = getGroupedOverlappingItems(items, direction);

    var doc = app.activeDocument;
    var newGroups = [];
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;

        /* グループ化前に元のレイヤーを保存 / Save the original layer before grouping */
        var originalLayer = group[0].layer;

        app.executeMenuCommand('deselectall');
        for (var j = 0; j < group.length; j++) {
            group[j].selected = true;
        }
        app.executeMenuCommand('group');
        var newGroup = app.activeDocument.selection[0];
        /* 新規グループのレイヤーを元に戻す / Restore the new group's layer to the original layer */
        newGroup.layer = originalLayer;
        newGroups.push(newGroup);
    }
    app.redraw();
    /* 新規グループを選択状態に設定 / Set new groups as selected */
    app.activeDocument.selection = null;
    for (var i = 0; i < newGroups.length; i++) {
        newGroups[i].selected = true;
    }
    return newGroups;
}

main();


/* 選択オブジェクト群の結合バウンディングボックス取得 / Get combined bounding box of selected objects */
function getCombinedBounds(items) {
    var left = null,
        top = null,
        right = null,
        bottom = null;
    for (var i = 0; i < items.length; i++) {
        var b = items[i].geometricBounds;
        if (left === null || b[0] < left) left = b[0];
        if (top === null || b[1] > top) top = b[1];
        if (right === null || b[2] > right) right = b[2];
        if (bottom === null || b[3] < bottom) bottom = b[3];
    }
    return [left, top, right, bottom];
}