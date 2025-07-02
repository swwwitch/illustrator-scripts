#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  SmartAutoGroup.jsx

  選択したオブジェクトを以下のいずれかの条件で自動的にグループ化します：
  - 重なり（のみ）
  - 垂直方向（近接しているもの）
  - 水平方向（近接しているもの）
  - 近接度（距離しきい値に応じて）

  【機能】
  - UIでモード選択＋しきい値スライダー（「重なり（のみ）」以外で有効）
  - グループ化単位ごとに Illustrator 標準の group コマンドで結合
  - グループ化後のオブジェクトのみ選択状態に
  - 未グループ化がある場合、再実行を促す確認ダイアログ付き（「重なり（のみ）」除く）

  【対象】
  - 複数選択されたオブジェクト（グループ含む）
  - アクティブドキュメントがあること

  作成日：2025-06-11
  更新日：2025-06-11
  バージョン：1.0.0
*/

var groupMode = "horizontal";

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

// 日英ラベル定義（UI表示順）
var LABELS = {
    modeGroupTitle: {
        ja: "グループ化",
        en: "Group"
    },
    vertical: {
        ja: "垂直方向",
        en: "Vertical"
    },
    horizontal: {
        ja: "水平方向",
        en: "Horizontal"
    },
    overlapOnly: {
        ja: "重なり（のみ）",
        en: "Overlap Only"
    },
    threshold: {
        ja: "しきい値（px）",
        en: "Threshold (px)"
    },
    group: {
        ja: "グループ化",
        en: "Group"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    resultMessage: {
        ja: "○個のグループ化を行いました。",
        en: " groups have been created."
    },
    retryMessage: {
        ja: "グループ化されなかったオブジェクトが {0} 個あります。\n再実行しますか？",
        en: "{0} objects were not grouped.\nDo you want to retry?"
    },
    proximity: {
        ja: "近接度",
        en: "Proximity"
    }
};

// function getTopmostPath is unused and can be removed.


// ダイアログUIの表示とユーザー選択取得
function showDialog(prevThreshold, prevGroupMode) {
    var dialog = new Window("dialog", LABELS.modeGroupTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = [25, 10, 25, 10];

    var modeGroup = dialog.add("group");
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";
    modeGroup.margins = [15, 10, 15, 10];

    var radioButtons = {
        overlapOnly: modeGroup.add("radiobutton", undefined, LABELS.overlapOnly[lang]),
        vertical: modeGroup.add("radiobutton", undefined, LABELS.vertical[lang]),
        horizontal: modeGroup.add("radiobutton", undefined, LABELS.horizontal[lang]),
        proximity: modeGroup.add("radiobutton", undefined, LABELS.proximity[lang])
    };
    var thresholdGroup = dialog.add("panel", undefined, LABELS.threshold[lang]);
    thresholdGroup.orientation = "column";
    thresholdGroup.alignChildren = "left";
    thresholdGroup.margins = [15, 20, 15, 10];
    thresholdGroup.enabled = true;

    radioButtons.vertical.onClick = function() {
        thresholdGroup.enabled = true;
    };
    radioButtons.horizontal.onClick = function() {
        thresholdGroup.enabled = true;
    };
    radioButtons.overlapOnly.onClick = function() {
        thresholdGroup.enabled = false;
    };
    radioButtons.proximity.onClick = function() {
        thresholdGroup.enabled = true;
    };

    if (prevGroupMode === "vertical") {
        radioButtons.vertical.value = true;
        thresholdGroup.enabled = true;
    } else if (prevGroupMode === "overlapOnly") {
        radioButtons.overlapOnly.value = true;
        thresholdGroup.enabled = false;
    } else if (prevGroupMode === "proximity") {
        radioButtons.proximity.value = true;
        thresholdGroup.enabled = true;
    } else {
        radioButtons.horizontal.value = true;
        thresholdGroup.enabled = true;
    }

    var thresholdSlider = thresholdGroup.add("slider", undefined, 10, 0, 100);
    thresholdSlider.value = prevThreshold;
    thresholdSlider.preferredSize.width = 150;
    var thresholdLabel = thresholdGroup.add("statictext", undefined, Math.round(thresholdSlider.value) + " pt");
    thresholdLabel.alignment = "center";
    thresholdLabel.characters = 5;

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";
    buttonGroup.margins = [0, 10, 0, 10]
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.group[lang], {
        name: "ok"
    });

    dialog.text = LABELS.modeGroupTitle[lang];

    thresholdSlider.onChanging = function() {
        thresholdLabel.text = Math.round(thresholdSlider.value) + " pt";
    };

    cancelBtn.onClick = function() {
        dialog.close();
    };
    okBtn.onClick = function() {
        overlapThreshold = thresholdSlider.value;
        if (radioButtons.vertical.value) {
            groupMode = "vertical";
        } else if (radioButtons.horizontal.value) {
            groupMode = "horizontal";
        } else if (radioButtons.overlapOnly.value) {
            groupMode = "overlapOnly";
        } else if (radioButtons.proximity.value) {
            groupMode = "proximity";
        }
        dialog.close(1);
    };

    var result = dialog.show();
    return result;
}

// 汎用DFS探索でグループ化（方向指定可）
function dfsGeneric(index, items, visited, group, threshold, direction) {
    visited[index] = true;
    group.push(items[index]);
    var boundsA = items[index].visibleBounds;
    for (var j = 0; j < items.length; j++) {
        if (visited[j]) continue;
        var boundsB = items[j].visibleBounds;
        var overlapRatio = getOverlapRatio(boundsA, boundsB);
        var adjacentDistance = getAdjacentDistance(boundsA, boundsB);
        var axisDistanceOk = false;
        if (threshold === -2 && direction) {
            if (direction === "vertical") {
                axisDistanceOk = getAxisDistanceVertical(boundsA, boundsB) <= overlapThreshold;
            } else if (direction === "horizontal") {
                axisDistanceOk = getAxisDistanceHorizontal(boundsA, boundsB) <= overlapThreshold;
            }
        }
        if (overlapRatio > 0 || adjacentDistance <= threshold || axisDistanceOk) {
            dfsGeneric(j, items, visited, group, threshold, direction);
        }
    }
}

// 隣接度または重なり率に基づいてグループを抽出（DFSによる連結成分抽出）
function getGroupedOverlappingItems(items, threshold, direction) {
    var groups = [];
    var visited = [];
    for (var i = 0; i < items.length; i++) {
        visited[i] = false;
    }
    for (var i = 0; i < items.length; i++) {
        if (visited[i]) continue;
        var group = [];
        dfsGeneric(i, items, visited, group, threshold, direction);
        groups.push(group);
    }
    return groups;
}

// 2つのバウンディングボックス間の最小距離を返す
function getAdjacentDistance(a, b) {
    var ax1 = a[0],
        ay1 = a[1],
        ax2 = a[2],
        ay2 = a[3];
    var bx1 = b[0],
        by1 = b[1],
        bx2 = b[2],
        by2 = b[3];

    var horzGap = Math.max(0, Math.max(bx1 - ax2, ax1 - bx2));
    var vertGap = Math.max(0, Math.max(ay2 - by1, by2 - ay1));

    return Math.max(horzGap, vertGap);
}

// 2つのバウンディングボックスの水平方向距離を返す（縦方向距離がある場合は無視）
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

// 2つのバウンディングボックスの垂直方向距離を返す（横方向距離がある場合は無視）
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

// 2つのバウンディングボックスの重なり率（大きい方の面積に対する割合）を返す
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

// アイテムの配列を zOrderPosition に基づいて昇順ソート
function sortByZOrder(items) {
    return items.slice().sort(function(a, b) {
        var za = -1,
            zb = -1;
        try {
            za = a.zOrderPosition;
        } catch (e) {}
        try {
            zb = b.zOrderPosition;
        } catch (e) {}
        return za - zb;
    });
}

// アイテムを zOrder でソートしてグループに移動

// 共通グループ化処理（direction: "horizontal"|"vertical"）
function groupOverlappingObjectsByDirection(direction) {
    if (!app.documents.length) return;
    var items = app.activeDocument.selection;
    if (!items || items.length === 0) return;

    var threshold = -2;
    var groups = getGroupedOverlappingItems(items, threshold, direction);

    var doc = app.activeDocument;
    var newGroups = [];
    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;

        // Save the original layer before grouping
        var originalLayer = group[0].layer;

        app.executeMenuCommand('deselectall');
        for (var j = 0; j < group.length; j++) {
            group[j].selected = true;
        }
        app.executeMenuCommand('group');
        var newGroup = app.activeDocument.selection[0];
        // Restore the new group's layer to the original layer
        newGroup.layer = originalLayer;
        newGroups.push(newGroup);
    }
    // The following block is no longer needed since the group is now placed in the original layer:
    /*
    for (var i = 0; i < newGroups.length; i++) {
        try {
            newGroups[i].move(doc, ElementPlacement.PLACEATEND);
        } catch (e) {}
    }
    */

    // グループ化されなかった単独オブジェクト数をカウントし再実行確認
    var ungroupedCount = 0;
    for (var i = 0; i < groups.length; i++) {
        if (groups[i].length === 1) ungroupedCount++;
    }
    if (ungroupedCount > 0) {
        var retryMsg = LABELS.retryMessage[lang].replace("{0}", ungroupedCount);
        var retry = confirm(retryMsg);
        if (retry) {
            showDialog(overlapThreshold, groupMode);
            main(overlapThreshold, groupMode);
            return;
        }
    }
    var alertMsg = LABELS.resultMessage[lang].replace("○", newGroups.length);
    app.redraw();
    alert(alertMsg);
    // 新規グループを選択状態に設定
    app.activeDocument.selection = null;
    for (var i = 0; i < newGroups.length; i++) {
        newGroups[i].selected = true;
    }
    return newGroups;
}

function main(prevThreshold, prevGroupMode) {
    overlapThreshold = (typeof prevThreshold === "number") ? prevThreshold : 10;
    if (!app.documents.length) return;
    var selection = app.activeDocument.selection;
    if (!selection || selection.length === 0) return;
    var result = showDialog(overlapThreshold, prevGroupMode);
    if (result !== 1) return;
    // overlapThreshold は showDialog() 内で更新される
    if (groupMode === "vertical") {
        groupOverlappingObjectsByDirection("vertical");
    } else if (groupMode === "overlapOnly") {
        groupOverlappingObjectsByOverlap();
    } else if (groupMode === "proximity") {
        groupOverlappingObjectsByProximity();
    } else {
        groupOverlappingObjectsByDirection("horizontal");
    }
}

// デフォルトグループモードを判定してメイン処理を呼び出し
var defaultGroupMode = detectDefaultGroupMode();
main(10, defaultGroupMode);

// デフォルトグループモードを選択範囲から自動判定
function detectDefaultGroupMode() {
    if (!app.documents.length) return "horizontal";
    var sel = app.activeDocument.selection;
    if (!sel || sel.length === 0) return "horizontal";

    var bounds = getCombinedBounds(sel);
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];

    return (width >= height) ? "horizontal" : "vertical";
}

// 選択オブジェクト群の結合バウンディングボックス取得
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
// 重なりのみでグループ化する関数（overlapOnlyモード用）
function groupOverlappingObjectsByOverlap() {
    if (!app.documents.length) return;
    var sel = app.activeDocument.selection;
    if (!sel || sel.length === 0) return;

    var threshold = -1; // 重なりのみで判定するため
    var groups = getGroupedOverlappingItems(sel, threshold);
    var doc = app.activeDocument;
    var newGroups = [];

    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;

        app.executeMenuCommand('deselectall');
        for (var j = 0; j < group.length; j++) {
            group[j].selected = true;
        }
        app.executeMenuCommand('group');
        var newGroup = app.activeDocument.selection[0];
        newGroups.push(newGroup);
    }

    for (var i = 0; i < newGroups.length; i++) {
        try {
            newGroups[i].move(doc, ElementPlacement.PLACEATEND);
        } catch (e) {}
    }

    app.activeDocument.selection = null;
    for (var i = 0; i < newGroups.length; i++) {
        newGroups[i].selected = true;
    }
    if (newGroups.length === 0) {
        alert((lang === "ja") ? "グループ化は行われませんでした" : "No groups were created.");
    } else {
        var alertMsg = LABELS.resultMessage[lang].replace("○", newGroups.length);
        app.redraw();
        alert(alertMsg);
    }
    return newGroups;
}

// 近接度によるグループ化（proximity モード）
function groupOverlappingObjectsByProximity() {
    if (!app.documents.length) return;
    var sel = app.activeDocument.selection;
    if (!sel || sel.length === 0) return;

    var groups = getGroupedOverlappingItems(sel, overlapThreshold);
    var doc = app.activeDocument;
    var newGroups = [];

    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;

        app.executeMenuCommand('deselectall');
        for (var j = 0; j < group.length; j++) {
            group[j].selected = true;
        }
        app.executeMenuCommand('group');
        var newGroup = app.activeDocument.selection[0];
        newGroups.push(newGroup);
    }

    for (var i = 0; i < newGroups.length; i++) {
        try {
            newGroups[i].move(doc, ElementPlacement.PLACEATEND);
        } catch (e) {}
    }

    // グループ化されなかった単独オブジェクト数をカウントし再実行確認
    var ungroupedCount = 0;
    for (var i = 0; i < groups.length; i++) {
        if (groups[i].length === 1) ungroupedCount++;
    }
    if (ungroupedCount > 0) {
        var retryMsg = LABELS.retryMessage[lang].replace("{0}", ungroupedCount);
        var retry = confirm(retryMsg);
        if (retry) {
            showDialog(overlapThreshold, groupMode);
            main(overlapThreshold, groupMode);
        }
        return;
    }

    app.activeDocument.selection = null;
    for (var i = 0; i < newGroups.length; i++) {
        newGroups[i].selected = true;
    }

    if (newGroups.length === 0) {
        alert((lang === "ja") ? "グループ化は行われませんでした" : "No groups were created.");
    } else {
        var alertMsg = LABELS.resultMessage[lang].replace("○", newGroups.length);
        app.redraw();
        alert(alertMsg);
    }

    return newGroups;
}

