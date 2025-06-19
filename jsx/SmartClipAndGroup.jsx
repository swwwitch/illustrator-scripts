#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  スクリプト名：SmartClipAndGroup.jsx
  概要：Illustratorで選択オブジェクトを条件に応じてグループ化またはクリッピングマスクを実行するスクリプト

  【機能概要】
  - 距離しきい値または重なり率に基づくグループ化（未グループ時は再実行可）
  - 最前面／最背面オブジェクトによるクリッピングマスク
  - 配置画像の矩形／正方形マスク対応
  - 日英ローカライズ対応（LABELS定義）

  【対象オブジェクト】
  - パス、配置画像、ラスター画像、グループなど
  【対象外】
  - ロック、非表示、ガイド、テンプレートのオブジェクト

  作成日：2024年6月5日
  更新日：2025年6月10日
    - 0.0.1: 初版リリース
    - 0.0.2: UI簡略化と構造整理
    - 0.0.3: 配置画像のみ処理・正方形マスク追加
    - 0.0.4: 「重なり」によるグループ化処理追加
    - 0.0.5: 重ね順保持／再実行対応／初期しきい値復元機能を追加
*/

var overlapThreshold = 10;

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();
// 日英ラベル定義 / Define label (ja/en)
// ラベル定義（UI出現順に再配置）
var LABELS = {
    dialogTitle: {
        ja: "マスクとグループ化",
        en: "Mask and Group"
    },
    clipPanel: {
        ja: "クリッピングマスク",
        en: "Clipping Mask"
    },
    clipFront: {
        ja: "最前面でクリップ",
        en: "Clip with Front"
    },
    clip: {
        ja: "最背面でクリップ",
        en: "Clip with Back"
    },
    clipPlacedOnly: {
        ja: "配置画像のみをクリップ",
        en: "Clip Placed Only"
    },
    groupPanel: {
        ja: "グループ化",
        en: "Grouping"
    },
    overlap: {
        ja: "重なり",
        en: "Overlap"
    },
    group: {
        ja: "近接度",
        en: "Threshold"
    },
    vertical: {
        ja: "上下方向",
        en: "Vertical"
    },
    horizontal: {
        ja: "左右方向",
        en: "Horizontal"
    },
    threshold: {
        ja: "しきい値（px）",
        en: "Threshold (px)"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

// グループ内で最前面のPathItemを取得
function getTopmostPath(group) {
    var topItem = null;
    for (var i = 0; i < group.length; i++) {
        if (group[i].typename === "PathItem") {
            if (!topItem || group[i].zOrderPosition > topItem.zOrderPosition) {
                topItem = group[i];
            }
        }
    }
    return topItem;
}

// グループ内で最背面のPathItemを取得
function getBottommostPath(group) {
    var bottomItem = null;
    for (var i = 0; i < group.length; i++) {
        if (group[i].typename === "PathItem") {
            if (!bottomItem || group[i].zOrderPosition < bottomItem.zOrderPosition) {
                bottomItem = group[i];
            }
        }
    }
    return bottomItem;
}

// 有効な選択オブジェクトを取得
function getValidSelection() {
    if (!app.documents.length) return null;
    var items = app.activeDocument.selection;
    if (!items || items.length === 0) return null;
    return items;
}


// ダイアログUIの表示とユーザー選択取得
function showDialog(initialThreshold) {
    var sel = getValidSelection();
    // ラジオボタンの初期選択
    var defaultKey = (sel && isAllPlacedItems(sel)) ? "clipPlacedOnly" : "group";
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";
    dialog.margins = [25, 20, 25, 20];

    // クリッピングマスク用パネル
    var clipPanel = dialog.add("panel", undefined, LABELS.clipPanel[lang]);
    clipPanel.orientation = "column";
    clipPanel.alignChildren = "left";
    clipPanel.margins = [15, 20, 15, 10];

    // グループ化用パネル
    var groupPanel = dialog.add("panel", undefined, LABELS.groupPanel[lang]);
    groupPanel.orientation = "column";
    groupPanel.alignChildren = "left";
    groupPanel.margins = [15, 20, 15, 10];

    // ラジオボタン定義
    var radioButtons = {};
    radioButtons.clipFront = clipPanel.add("radiobutton", undefined, LABELS.clipFront[lang]);
    radioButtons.clip = clipPanel.add("radiobutton", undefined, LABELS.clip[lang]);
    radioButtons.clipPlacedOnly = clipPanel.add("radiobutton", undefined, LABELS.clipPlacedOnly[lang]);
    radioButtons.overlap = groupPanel.add("radiobutton", undefined, LABELS.overlap[lang]);
    radioButtons.group = groupPanel.add("radiobutton", undefined, LABELS.group[lang]);
    // 追加: 上下方向・左右方向ラジオボタン
    radioButtons.vertical = groupPanel.add("radiobutton", undefined, LABELS.vertical[lang]);
    radioButtons.horizontal = groupPanel.add("radiobutton", undefined, LABELS.horizontal[lang]);

    // 初期選択ラジオボタン設定
    radioButtons[defaultKey].value = true;

    // しきい値スライダー
    var thresholdSlider = groupPanel.add("slider", undefined, 10, 0, 100);
    thresholdSlider.value = (typeof initialThreshold === "number") ? initialThreshold : 10;
    thresholdSlider.preferredSize.width = 150;
    var thresholdLabel = groupPanel.add("statictext", undefined, LABELS.threshold[lang]);
    thresholdLabel.alignment = "center";
    thresholdLabel.characters = 5;
    thresholdLabel.text = Math.round(thresholdSlider.value) + " pt";

    // しきい値スライダー表示制御
    var thresholdControls = [thresholdSlider, thresholdLabel];
    thresholdSlider.onChanging = function() {
        thresholdLabel.text = Math.round(thresholdSlider.value) + " pt";
    };

    // ボタングループ
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });

    // ダイアログタイトル設定
    dialog.text = LABELS.dialogTitle[lang];

    var result = null;
    // しきい値スライダーの有効制御
    for (var key in radioButtons) {
        if (radioButtons[key].value) {
            var dim = !(key === "group" || key === "vertical");
            for (var i = 0; i < thresholdControls.length; i++) {
                thresholdControls[i].enabled = !dim;
            }
        }
    }
    for (var key in radioButtons) {
        radioButtons[key].onClick = function() {
            for (var k in radioButtons) {
                radioButtons[k].value = (radioButtons[k] === this);
            }
            var dim = !(this === radioButtons.group || this === radioButtons.vertical);
            for (var i = 0; i < thresholdControls.length; i++) {
                thresholdControls[i].enabled = !dim;
            }
        };
    }
    cancelBtn.onClick = function() {
        // On cancel, just close dialog, return null
        dialog.close();
        result = null;
    };
    okBtn.onClick = function() {
        for (var key in radioButtons) {
            if (radioButtons[key].value) result = key;
        }
        overlapThreshold = thresholdSlider.value;
        dialog.close();
    };

    dialog.show();
    return result;
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
        dfs(i, items, visited, group, threshold, direction);
        groups.push(group);
    }

    return groups;
}

// DFSで隣接または重なりオブジェクトを探索
function dfs(index, items, visited, group, threshold, direction) {
    visited[index] = true;
    group.push(items[index]);
    var boundsA = items[index].geometricBounds;

    for (var j = 0; j < items.length; j++) {
        if (visited[j]) continue;
        var boundsB = items[j].geometricBounds;
        var overlapRatio = getOverlapRatio(boundsA, boundsB);
        var adjacentDistance = getAdjacentDistance(boundsA, boundsB, direction);

        if (direction === "vertical") {
            if (adjacentDistance <= threshold) {
                dfs(j, items, visited, group, threshold, direction);
            }
        } else {
            if (overlapRatio > 0 || adjacentDistance <= threshold) {
                dfs(j, items, visited, group, threshold, direction);
            }
        }
    }
}

// 2つのバウンディングボックス間の最小距離を返す
function getAdjacentDistance(a, b, direction) {
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

    if (direction === "vertical") {
        return vertGap;
    }
    // デフォルトは最大
    return Math.max(horzGap, vertGap);
}





// 2つのバウンディングボックスの重なり率（小さい方の面積に対する割合）を返す
function getOverlapRatio(a, b) {
    var ax = Math.max(0, Math.min(a[2], b[2]) - Math.max(a[0], b[0]));
    var ay = Math.max(0, Math.min(a[1], b[1]) - Math.max(a[3], b[3]));
    var overlapArea = ax * ay;
    if (overlapArea <= 0) return 0;
    var areaA = (a[2] - a[0]) * (a[1] - a[3]);
    var areaB = (b[2] - b[0]) * (b[1] - b[3]);
    var minArea = Math.min(areaA, areaB);
    return overlapArea / minArea;
}

// 指定パスをグループ内で最前面に移動する関数
function bringReferenceToFront(group, refItem) {
    for (var i = 0; i < group.length; i++) {
        if (group[i] !== refItem) {
            if (group[i].zOrderPosition > refItem.zOrderPosition) {
                refItem.move(group[i], ElementPlacement.PLACEBEFORE);
            }
        }
    }
}

// 指定関数で取得したパスを用いてマスクを作成
function clipWithReferenceItem(itemsToGroup, getReferencePath) {
    if (itemsToGroup.length <= 1) return;
    // グループ化前に最前面のオブジェクトのzOrderPositionを記録
    // 1. 元の重ね順を記録
    var originalZOrders = [];
    for (var i = 0; i < itemsToGroup.length; i++) {
        originalZOrders.push(itemsToGroup[i].zOrderPosition);
    }
    var doc = app.activeDocument;
    var group = doc.groupItems.add();
    var refItem = getReferencePath(itemsToGroup);
    if (!refItem) return;
    bringReferenceToFront(itemsToGroup, refItem);
    refItem.move(group, ElementPlacement.PLACEATBEGINNING);
    refItem.clipping = true;
    for (var i = 0; i < itemsToGroup.length; i++) {
        if (itemsToGroup[i] !== refItem) {
            itemsToGroup[i].move(group, ElementPlacement.PLACEATEND);
        }
    }
    group.clipped = true;
    // グループ内の順序を元に戻す
    for (var i = 1; i < group.pageItems.length; i++) {
        for (var j = i; j > 0; j--) {
            var idxJ = -1, idxJm1 = -1;
            for (var k = 0; k < itemsToGroup.length; k++) {
                if (group.pageItems[j] === itemsToGroup[k]) idxJ = k;
                if (group.pageItems[j-1] === itemsToGroup[k]) idxJm1 = k;
            }
            if (group.pageItems[j] === refItem) idxJ = itemsToGroup.indexOf(refItem);
            if (group.pageItems[j-1] === refItem) idxJm1 = itemsToGroup.indexOf(refItem);
            if (idxJ < 0 || idxJm1 < 0) continue;
            if (originalZOrders[idxJ] < originalZOrders[idxJm1]) {
                group.pageItems[j].zOrder(ZOrderMethod.SENDTOBACK);
            }
        }
    }
    // グループを元の参照アイテムの直前に配置
    group.move(refItem, ElementPlacement.PLACEBEFORE);
}

// 最背面のパスを使いグループごとにクリッピングマスクを作成
function clipOverlappingObjects() {
    var items = getValidSelection();
    if (!items) return;
    var groups = getGroupedOverlappingItems(items);
    for (var i = 0; i < groups.length; i++) {
        clipWithReferenceItem(groups[i], getBottommostPath);
    }
}
// 最前面のパスを使いグループごとにクリッピングマスクを作成
function clipOverlappingObjectsFront() {
    var items = getValidSelection();
    if (!items) return;
    var groups = getGroupedOverlappingItems(items);
    for (var i = 0; i < groups.length; i++) {
        clipWithReferenceItem(groups[i], getTopmostPath);
    }
}

// 配置画像のみを矩形マスク
function clipPlacedOnlyMask() {
    processPlacedItemsByType("rect");
}

// クリッピングマスクを解除
function releaseClippingMask(groupItem) {
    var clippingPath = null;
    for (var i = 0; i < groupItem.pageItems.length; i++) {
        if (groupItem.pageItems[i].clipping) {
            clippingPath = groupItem.pageItems[i];
            break;
        }
    }
    if (clippingPath) {
        groupItem.clipped = false;
        clippingPath.remove();
    }
    ungroupGroupItem(groupItem);
}

// 配置画像を矩形でマスク
function createClippingMask(imageItem) {
    var layer = imageItem.layer;
    var wasLocked = layer.locked;
    var wasVisible = layer.visible;
    var wasTemplate = layer.isTemplate;

    if (wasLocked) layer.locked = false;
    if (!wasVisible) layer.visible = true;
    if (wasTemplate) layer.isTemplate = false;

    var rect = layer.pathItems.rectangle(
        imageItem.top,
        imageItem.left,
        imageItem.width,
        imageItem.height
    );
    rect.stroked = false;
    rect.filled = false;

    var group = layer.groupItems.add();
    imageItem.moveToBeginning(group);
    rect.moveToBeginning(group);
    group.clipped = true;

    if (wasLocked) layer.locked = true;
    if (!wasVisible) layer.visible = false;
    if (wasTemplate) layer.isTemplate = true;

    return rect;
}

// 配置画像とパスでマスク
function createMaskWithPath(imageItem, pathItem) {
    var layer = imageItem.layer;
    if (pathItem.layer != layer) {
        pathItem.move(layer, ElementPlacement.PLACEATBEGINNING);
    }

    var group = layer.groupItems.add();
    imageItem.moveToBeginning(group);
    pathItem.moveToBeginning(group);
    group.clipped = true;

    return pathItem;
}

// グループ解除
function ungroupGroupItem(groupItem) {
    var parent = groupItem.parent;
    while (groupItem.pageItems.length > 0) {
        groupItem.pageItems[0].moveToBeginning(parent);
    }
    groupItem.remove();
}

// 選択を更新
function updateSelection(clippingMasks) {
    var doc = app.activeDocument;
    doc.selection = null;
    for (var i = 0; i < clippingMasks.length; i++) {
        clippingMasks[i].parent.selected = true;
    }
}

// 配列が全てPathItemか判定
function isAllPathItems(arr) {
    for (var i = 0; i < arr.length; i++) {
        if (!arr[i] || arr[i].typename !== "PathItem") return false;
    }
    return true;
}

// 最前面のPathItemを取得
function getFrontmostPath(paths) {
    var topPath = null;
    var maxZ = -1;
    for (var i = 0; i < paths.length; i++) {
        if (paths[i].typename === "PathItem" && paths[i].zOrderPosition > maxZ) {
            topPath = paths[i];
            maxZ = paths[i].zOrderPosition;
        }
    }
    return topPath;
}

// 配列が全て配置画像か判定
function isAllPlacedItems(selectionArray) {
    if (!selectionArray || selectionArray.length === 0) return false;
    for (var i = 0; i < selectionArray.length; i++) {
        if (selectionArray[i].typename !== "PlacedItem" && selectionArray[i].typename !== "RasterItem") {
            return false;
        }
    }
    return true;
}

// ユーザー選択に応じて処理を実行
function executeUserChoice(choice) {
    switch (choice) {
        case "group":
        case "vertical":
            var newGroups = groupOverlappingObjectsByThreshold(choice);
            if (newGroups && newGroups.length > 0) {
                app.activeDocument.selection = null;
                for (var i = 0; i < newGroups.length; i++) {
                    newGroups[i].selected = true;
                }
            }
            break;
        case "clip":
            clipOverlappingObjects();
            break;
        case "clipFront":
            clipOverlappingObjectsFront();
            break;
        case "clipPlacedOnly":
            clipPlacedOnlyMask();
            break;
        case "overlap":
            groupOverlappingObjectsByThreshold("overlap");
            break;
    }
}

// 配置画像のみ（正方形）でマスク
function clipPlacedOnlySquareMask() {
    processPlacedItemsByType("square");
}

// 配置画像のみ・矩形/正方形マスク処理
function processPlacedItemsByType(type) {
    var doc = app.activeDocument;
    var sel = getValidSelection();
    if (!sel) return;

    var createdGroups = [];

    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename === 'GroupItem' && item.clipped) {
            item.clipped = false;
            var itemsToProcess = [];
            for (var j = item.pageItems.length - 1; j >= 0; j--) {
                var pageItem = item.pageItems[j];
                if (pageItem.typename === 'PlacedItem' || pageItem.typename === 'RasterItem') {
                    itemsToProcess.push({
                        img: pageItem,
                        idx: i
                    });
                } else {
                    pageItem.remove();
                }
            }
            for (var k = 0; k < itemsToProcess.length; k++) {
                processImageWithShape(itemsToProcess[k].img, type, createdGroups, itemsToProcess[k].idx);
            }
        } else if (item.typename === 'PlacedItem' || item.typename === 'RasterItem') {
            var originalLayer = item.layer;
            processImageWithShape(item, type, createdGroups, i);
            var lastGroup = createdGroups[createdGroups.length - 1].group;
            if (lastGroup.layer != originalLayer) {
                lastGroup.move(originalLayer, ElementPlacement.PLACEATEND);
            }
        }
    }

    doc.selection = null;
    sortByIndex(createdGroups);
    for (var i = 0; i < createdGroups.length; i++) {
        var group = createdGroups[i].group;
        var originalLayer = group.layer;
        group.move(originalLayer, ElementPlacement.PLACEATEND);
        group.selected = true;
    }
}

// 配置画像を指定形状でクリッピング
function processImageWithShape(image, shapeType, groups, index) {
    var bounds = image.visibleBounds;
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];
    var centerX = bounds[0] + width / 2;
    var centerY = bounds[1] - height / 2;
    var sideLength = Math.min(width, height);
    var parentLayer = image.layer;
    var wasLocked = parentLayer.locked;
    var wasTemplate = parentLayer.isTemplate;
    if (wasLocked) parentLayer.locked = false;
    if (wasTemplate) parentLayer.isTemplate = false;

    // パスを画像のすぐ上に作成
    var shape;
    if (shapeType === "square") {
        shape = parentLayer.pathItems.rectangle(centerY + sideLength / 2, centerX - sideLength / 2, sideLength, sideLength);
    } else {
        shape = parentLayer.pathItems.rectangle(bounds[1], bounds[0], width, height);
    }
    shape.stroked = false;
    shape.filled = false;

    var zOrderPos = -1;
    try {
        zOrderPos = image.zOrderPosition;
    } catch (e) {}

    // 画像とパスを一時グループにまとめてマスク
    var group = parentLayer.groupItems.add();
    // 画像を先に、パスを後に move してパスが上に来るようにする
    image.moveToBeginning(group);
    shape.moveToBeginning(group);
    group.clipped = true;

    // マスク対象のパスを最前面に再配置
    try {
        var topPath = null;
        for (var j = 0; j < group.pageItems.length; j++) {
            if (group.pageItems[j].typename === "PathItem") {
                topPath = group.pageItems[j];
                break;
            }
        }
        if (topPath) topPath.moveToBeginning(group);
    } catch (e) {}

    try {
        if (zOrderPos >= 0) group.zOrder(zOrderPos);
    } catch (e) {}

    groups.push({
        group: group,
        index: index
    });

    var originalGroup = image.parent;
    if (originalGroup.typename === "GroupItem" && originalGroup.pageItems.length === 0) {
        try {
            originalGroup.remove();
        } catch (e) {}
    }

    if (wasLocked) parentLayer.locked = true;
    if (wasTemplate) parentLayer.isTemplate = true;
}
// indexでソート
function sortByIndex(arr) {
    arr.sort(function(a, b) {
        return a.index - b.index;
    });
}
// 指定アイテムがグループ化済みか判定
function isItemGrouped(item, groups) {
    for (var i = 0; i < groups.length; i++) {
        for (var j = 0; j < groups[i].length; j++) {
            if (groups[i][j] === item) return true;
        }
    }
    return false;
}

// メイン処理
function main(prevThreshold) {
    var userChoice = showDialog(prevThreshold);
    executeUserChoice(userChoice);
}

// アイテムの配列を zOrderPosition に基づいて昇順ソートする汎用関数
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

// 指定した重なりしきい値でグループ化
function groupOverlappingObjectsByThreshold(mode) {
    // threshold の定義を追加
    var threshold = (mode === "threshold" || mode === "group") ? overlapThreshold : 0.01;
    // 指定しきい値に基づき、重なり・隣接オブジェクトをグループ化（重ね順を保持）
    var sel = getValidSelection();
    if (!sel) return;

    var groups;
    if (mode === "vertical") {
        groups = getGroupedOverlappingItems(sel, threshold, "vertical");
    } else {
        groups = getGroupedOverlappingItems(sel, threshold);
    }
    var doc = app.activeDocument;
    var newGroups = [];

    for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        if (group.length <= 1) continue;

        var newGroup = doc.groupItems.add();
        // 一旦すべて移動
        for (var j = 0; j < group.length; j++) {
            group[j].move(newGroup, ElementPlacement.PLACEATEND);
        }
        // 重ね順でソートして再配置
        moveItemsToGroupSorted(group, newGroup);
        newGroups.push(newGroup);
    }

    for (var i = 0; i < newGroups.length; i++) {
        try {
            newGroups[i].move(app.activeDocument, ElementPlacement.PLACEATEND);
        } catch (e) {}
    }

    // グループ化されなかった（単独）オブジェクト数をカウント
    var ungroupedCount = 0;
    for (var i = 0; i < groups.length; i++) {
        if (groups[i].length === 1) ungroupedCount++;
    }
    // mode === "threshold" のときのみ再実行確認
    if (mode === "threshold" && ungroupedCount > 0) {
        var retry = confirm("グループ化されなかったオブジェクトが " + ungroupedCount + " 個あります。\n再実行しますか？");
        if (retry) {
            main(overlapThreshold);
        }
    }

    // mode が "overlap" または "threshold" のとき新しく作成されたグループを選択状態に設定
    if (mode === "overlap" || mode === "threshold") {
        app.activeDocument.selection = null;
        for (var i = 0; i < newGroups.length; i++) {
            newGroups[i].selected = true;
        }
    }
    return newGroups;
}

// アイテムを zOrder でソートしてグループに移動する関数
function moveItemsToGroupSorted(items, group) {
    var sorted = sortByZOrder(items);
    for (var i = 0; i < sorted.length; i++) {
        try {
            sorted[i].move(group, ElementPlacement.PLACEATEND);
        } catch (e) {}
    }
}

// 重なり率が0より大きいオブジェクト同士をグループ化（しきい値なし）
function groupOverlappingObjectsSimple() {
    // 0.01の重なりしきい値でグループ化
    var sel = getValidSelection();
    if (!sel) return;
    var groups = getGroupedOverlappingItems(sel, 0.01);
    var doc = app.activeDocument;
    var newGroups = [];
    for (var i = 0; i < groups.length; i++) {
        var groupItems = groups[i];
        if (groupItems.length <= 1) continue;
        // グループ内で最前面のオブジェクトを取得
        var topItem = getTopmostPath(groupItems);
        var newGroup = doc.groupItems.add();
        // 一旦すべて移動
        for (var j = 0; j < groupItems.length; j++) {
            groupItems[j].move(newGroup, ElementPlacement.PLACEATEND);
        }
        // 重ね順でソートして再配置
        moveItemsToGroupSorted(groupItems, newGroup);
        // グループを最前面オブジェクトの直前に再配置
        if (topItem) {
            try {
                newGroup.move(topItem, ElementPlacement.PLACEBEFORE);
            } catch (e) {}
        }
        newGroups.push(newGroup);
    }
    // 生成したグループを選択
    app.activeDocument.selection = null;
    for (var i = 0; i < newGroups.length; i++) {
        newGroups[i].selected = true;
    }
    return newGroups;
}

// UIのしきい値をもとにグループ化を呼び出す
function groupOverlappingObjects() {
    groupOverlappingObjectsByThreshold(overlapThreshold);
}

// メイン処理の呼び出し
main();