#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartAutoGroup.jsx

### 概要

- 選択オブジェクトを「重なり」「垂直」「水平」「近接度」などの条件に基づいて自動的にグループ化するIllustrator用スクリプトです。
- グループ化後にアートボードへの変換やマージン設定も可能です。

### 主な機能

- モード選択（重なり、垂直、水平、近接度）
- しきい値スライダーによる距離設定
- グループ後にアートボードへ変換（オプション）
- マージン・既存アートボード削除オプション
- 未グループオブジェクト再実行確認
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログでモード、しきい値、アートボード設定を選択
2. DFS探索によりグループ化対象を抽出
3. グループ化実行と再実行確認
4. 必要に応じてアートボード生成

### 更新履歴

- v1.0.0 (20250611) : 初期バージョン
- v1.0.1 (20250612) : ダイアログUI改善、アートボードオプション追加

---

### Script Name:

SmartAutoGroup.jsx

### Overview

- An Illustrator script to automatically group selected objects based on conditions like "overlap", "vertical", "horizontal", or "proximity".
- Allows converting groups into artboards with margin and deletion options after grouping.

### Main Features

- Mode selection (Overlap, Vertical, Horizontal, Proximity)
- Distance threshold slider
- Convert grouped objects to artboards (optional)
- Margin and existing artboard deletion options
- Retry prompt for ungrouped objects
- Japanese and English UI support

### Process Flow

1. Select mode, threshold, and artboard options in dialog
2. Extract groups using DFS traversal
3. Execute grouping and prompt for retry if needed
4. Generate artboards as needed

### Update History

- v1.0.0 (20250611): Initial version
- v1.0.1 (20250612): Improved dialog UI and added artboard options
*/

var groupMode = "horizontal";

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

// UI表示用ラベル（日英・使用順）
var LABELS = {
    modeGroupTitle: {
        ja: "グループ化からのアートボード変換",
        en: "Group"
    },
    overlapOnly: {
        ja: "重なり（のみ）",
        en: "Overlap Only"
    },
    vertical: {
        ja: "垂直方向",
        en: "Vertical"
    },
    horizontal: {
        ja: "水平方向",
        en: "Horizontal"
    },
    proximity: {
        ja: "近接度",
        en: "Proximity"
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
    artboardPanel: {
        ja: "アートボード化",
        en: "Artboard"
    },
    convertToArtboard: {
        ja: "アートボードに変換",
        en: "Convert to Artboard"
    },
    margin: {
        ja: "マージン",
        en: "Margin"
    }
};

// ダイアログUI表示・ユーザー選択取得
function showDialog(prevThreshold, prevGroupMode) {
    var dialog = new Window("dialog", LABELS.modeGroupTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = [15, 20, 15, 10];
    dialog.spacing = 10;


    // ドキュメントの rulerUnits を判定し、表示用の単位文字列を取得
    var unitMap = {};
    unitMap[RulerUnits.Inches] = "in";
    unitMap[RulerUnits.Millimeters] = "mm";
    unitMap[RulerUnits.Points] = "pt";
    unitMap[RulerUnits.Picas] = "pica";
    unitMap[RulerUnits.Centimeters] = "cm";
    unitMap[RulerUnits.Pixels] = "px";
    unitMap[RulerUnits.Feet] = "ft";
    unitMap[RulerUnits.Meters] = "m";
    unitMap[RulerUnits.Yards] = "yd";
    var rulerType = app.activeDocument.rulerUnits;
    var rulerUnit = unitMap[rulerType] ? unitMap[rulerType] : "pt";

    // メイン3カラムグループ
    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = "top";

    // 左カラム
    var leftColumn = mainGroup.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = "left";

    // グループ化モードラジオボタン
    var modeGroup = leftColumn.add("group");
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";
    modeGroup.margins = [15, 10, 15, 10];
    var radioButtons = {
        overlapOnly: modeGroup.add("radiobutton", undefined, LABELS.overlapOnly[lang]),
        vertical: modeGroup.add("radiobutton", undefined, LABELS.vertical[lang]),
        horizontal: modeGroup.add("radiobutton", undefined, LABELS.horizontal[lang]),
        proximity: modeGroup.add("radiobutton", undefined, LABELS.proximity[lang])
    };

    // しきい値設定パネル
    var thresholdGroup = leftColumn.add("panel", undefined, LABELS.threshold[lang]);
    thresholdGroup.orientation = "column";
    thresholdGroup.alignChildren = "left";
    thresholdGroup.margins = [15, 20, 15, 10];
    thresholdGroup.enabled = true;

    // アートボード関連パネル（中央カラムに移動）
    // 中央カラム（空）
    var centerColumn = mainGroup.add("group");
    centerColumn.orientation = "column";
    centerColumn.alignChildren = "left";

    var artboardPanel = centerColumn.add("panel", undefined, LABELS.artboardPanel[lang]);
    artboardPanel.orientation = "column";
    artboardPanel.alignChildren = "left";
    artboardPanel.margins = [15, 25, 15, 10];
    var convertCheck = artboardPanel.add("checkbox", undefined, LABELS.convertToArtboard[lang]);
    convertCheck.value = false;
    // プレビュー境界チェックボックスを追加
    var previewBoundsCheck = artboardPanel.add("checkbox", undefined, "プレビュー境界");
    previewBoundsCheck.value = false;
    // マージン入力欄
    var marginGroup = artboardPanel.add("group");
    marginGroup.orientation = "column";
    marginGroup.alignChildren = "left";
    var marginLabel = marginGroup.add("statictext", undefined, LABELS.margin[lang] + ":");
    var inputGroup = marginGroup.add("group");
    inputGroup.orientation = "row";
    var marginInput = inputGroup.add("edittext", undefined, "0");
    marginInput.characters = 5;
    var marginUnitLabel = inputGroup.add("statictext", undefined, rulerUnit);
    // 既存のアートボードを削除チェックボックスを追加（順序をマージンの後に変更）
    var deleteArtboardsCheck = artboardPanel.add("checkbox", undefined, "既存のアートボードを削除");
    deleteArtboardsCheck.value = false;

    // （中央カラムの定義はアートボードパネル追加に移動済み）

    // 右カラム（空）
    var rightColumn = mainGroup.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = "left";

    // モード選択時のしきい値有効/無効切替
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

    // 直前のモード・値を反映
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

    // しきい値スライダー
    var thresholdSlider = thresholdGroup.add("slider", undefined, 10, 0, 100);
    thresholdSlider.value = prevThreshold;
    thresholdSlider.preferredSize.width = 150;
    var thresholdLabel = thresholdGroup.add("statictext", undefined, Math.round(thresholdSlider.value) + " pt");
    thresholdLabel.alignment = "center";
    thresholdLabel.characters = 5;
    thresholdSlider.onChanging = function() {
        thresholdLabel.text = Math.round(thresholdSlider.value) + " pt";
    };

    // ボタンをダイアログ下部に横並びで配置
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";
    buttonGroup.margins = [0, 10, 0, 10];
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, "OK", {
        name: "ok"
    });

    dialog.text = LABELS.modeGroupTitle[lang];

    // ボタンクリック時処理
    var dialogResult = null;
    cancelBtn.onClick = function() {
        dialog.close();
    };
    okBtn.onClick = function() {
        overlapThreshold = thresholdSlider.value;
        if (radioButtons.vertical.value) groupMode = "vertical";
        else if (radioButtons.horizontal.value) groupMode = "horizontal";
        else if (radioButtons.overlapOnly.value) groupMode = "overlapOnly";
        else if (radioButtons.proximity.value) groupMode = "proximity";
        dialogResult = {
            overlapThreshold: thresholdSlider.value,
            groupMode: groupMode,
            convertToArtboard: convertCheck.value,
            marginValue: marginInput.text,
            deleteArtboards: deleteArtboardsCheck.value,
            usePreviewBounds: previewBoundsCheck.value
        };
        dialog.close(1);
    };
    var result = dialog.show();
    if (result !== 1) return null;
    return dialogResult;
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

// ※sortByZOrder関数は未使用のため削除

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
        // グループ化前のレイヤーを保持
        var originalLayer = group[0].layer;
        app.executeMenuCommand('deselectall');
        for (var j = 0; j < group.length; j++) {
            group[j].selected = true;
        }
        app.executeMenuCommand('group');
        var newGroup = app.activeDocument.selection[0];
        // グループを元レイヤーへ戻す
        newGroup.layer = originalLayer;
        newGroups.push(newGroup);
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

// メイン処理
function main(prevThreshold, prevGroupMode) {
    overlapThreshold = (typeof prevThreshold === "number") ? prevThreshold : 10;
    if (!app.documents.length) return;
    var selection = app.activeDocument.selection;
    if (!selection || selection.length === 0) return;
    var dialogResult = showDialog(overlapThreshold, prevGroupMode);
    if (!dialogResult) return;
    // ダイアログ結果を取得
    overlapThreshold = dialogResult.overlapThreshold;
    groupMode = dialogResult.groupMode;
    var convertToArtboard = dialogResult.convertToArtboard;
    var marginValue = dialogResult.marginValue;
    var deleteArtboards = dialogResult.deleteArtboards;
    var usePreviewBounds = dialogResult.usePreviewBounds;
    var newGroups = [];
    // グループ化本体
    if (groupMode === "vertical") {
        newGroups = groupOverlappingObjectsByDirection("vertical");
    } else if (groupMode === "overlapOnly") {
        newGroups = groupOverlappingObjectsByOverlap();
    } else if (groupMode === "proximity") {
        newGroups = groupOverlappingObjectsByProximity();
    } else {
        newGroups = groupOverlappingObjectsByDirection("horizontal");
    }
    // アートボード変換（オプション）
    if (convertToArtboard && newGroups && newGroups.length > 0) {
        var doc = app.activeDocument;
        var margin = parseFloat(marginValue);
        if (isNaN(margin)) margin = 0;
        // 既存アートボード数を保持
        var initialCount = doc.artboards.length;
        // 先にアートボードを追加
        for (var i = 0; i < newGroups.length; i++) {
            var bounds = usePreviewBounds ? newGroups[i].visibleBounds : newGroups[i].geometricBounds;
            var left = bounds[0] - margin;
            var top = bounds[1] + margin;
            var right = bounds[2] + margin;
            var bottom = bounds[3] - margin;
            // Illustratorのアートボード：[左, 上, 右, 下]
            try {
                doc.artboards.add([left, top, right, bottom]);
            } catch (e) {
                // エラー時は何もしない
            }
        }
        // 既存アートボードを削除（オプション）
        if (deleteArtboards) {
            // 新規アートボード以外を削除（先頭から initialCount 個を削除）
            for (var i = initialCount - 1; i >= 0; i--) {
                doc.artboards.remove(i);
            }
        }
    }
}

// デフォルトグループモードを判定してメイン処理を呼び出し
var defaultGroupMode = detectDefaultGroupMode();
main(10, defaultGroupMode);

// 選択範囲の縦横比からデフォルトグループモードを自動判定
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
// 重なりのみでグループ化（overlapOnlyモード）
function groupOverlappingObjectsByOverlap() {
    if (!app.documents.length) return;
    var sel = app.activeDocument.selection;
    if (!sel || sel.length === 0) return;
    var threshold = -1; // 重なりのみ
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

// 近接度によるグループ化（proximityモード）
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