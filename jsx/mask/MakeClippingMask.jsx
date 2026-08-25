#target illustrator

/*

### 概要

選択オブジェクトの状態に応じて、クリッピングマスクの作成と解除を切り替えて実行します。
すでにマスクが設定されていれば解除し、配置画像やパスの選択からは新たにマスクを作成します。

詳細は README を参照してください。

### Overview

Creates or releases a clipping mask, depending on what is selected.
An existing mask is released, while a selection of placed images or paths produces a new one.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MakeClippingMask";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-11-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-04-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeClippingMask.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeClippingMask.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function main() {
    var activeDoc = app.activeDocument;
    var createdClippingMasks = [];

    // クリッピングマスクを解除
    for (var i = 0; i < app.selection.length; i++) {
        var selectedObject = app.selection[i];
        if (selectedObject.typename === "GroupItem" && selectedObject.clipped) {
            releaseClippingMask(selectedObject);
        }
    }

    // 配置画像 + パス
    if (app.selection.length === 2) {
        var imageItem = null;
        var pathItem = null;

        for (var i = 0; i < 2; i++) {
            if (app.selection[i].typename === "PlacedItem" || app.selection[i].typename === "RasterItem") {
                imageItem = app.selection[i];
            } else if (app.selection[i].typename === "PathItem") {
                pathItem = app.selection[i];
            }
        }

        if (imageItem !== null && pathItem !== null) {
            var customMask = createMaskWithPath(imageItem, pathItem);
            if (customMask !== null) createdClippingMasks.push(customMask);
        }

    // パスのみ選択されているとき
    } else if (isAllPathItems(app.selection) && app.selection.length >= 2) {
        var topPath = getFrontmostPath(app.selection); // 修正ポイント
        if (topPath !== null) {
            var group = activeDoc.groupItems.add();
            topPath.moveToBeginning(group);
            topPath.clipping = true;

            for (var j = 0; j < app.selection.length; j++) {
                var obj = app.selection[j];
                if (obj !== topPath) {
                    obj.moveToEnd(group);
                }
            }

            group.clipped = true;
            createdClippingMasks.push(topPath);
        }

    // 単体の画像など
    } else {
        for (var i = 0; i < app.selection.length; i++) {
            var selectedObject = app.selection[i];
            if (selectedObject.typename === "PlacedItem" || selectedObject.typename === "RasterItem") {
                var newMask = createClippingMask(selectedObject);
                createdClippingMasks.push(newMask);
            }
        }
    }

    updateSelection(createdClippingMasks);

})();

// クリッピングマスクを解除してパスを削除
function releaseClippingMask(groupItem) {
    var clippingPath = null;
    for (var i = 0; i < groupItem.pageItems.length; i++) {
        if (groupItem.pageItems[i].clipping) {
            clippingPath = groupItem.pageItems[i];
            break;
        }
    }
    if (clippingPath !== null) {
        groupItem.clipped = false;
        clippingPath.remove();
    }
    ungroupGroupItem(groupItem);
}

// 画像に矩形マスクを作成
function createClippingMask(imageItem) {
    var targetLayer = imageItem.layer;
    var wasLocked = targetLayer.locked;
    var wasVisible = targetLayer.visible;
    var wasTemplate = targetLayer.isTemplate;

    if (wasLocked) targetLayer.locked = false;
    if (!wasVisible) targetLayer.visible = true;
    if (wasTemplate) targetLayer.isTemplate = false;

    var clippingRect = targetLayer.pathItems.rectangle(
        imageItem.top,
        imageItem.left,
        imageItem.width,
        imageItem.height
    );
    clippingRect.stroked = false;
    clippingRect.filled = false;

    var clippingGroup = targetLayer.groupItems.add();
    imageItem.moveToBeginning(clippingGroup);
    clippingRect.moveToBeginning(clippingGroup);
    clippingGroup.clipped = true;

    if (wasLocked) targetLayer.locked = true;
    if (!wasVisible) targetLayer.visible = false;
    if (wasTemplate) targetLayer.isTemplate = true;

    return clippingRect;
}

// パスでマスクを作成
function createMaskWithPath(imageItem, pathItem) {
    var targetLayer = imageItem.layer;
    if (pathItem.layer != targetLayer) {
        pathItem.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
    }

    var clippingGroup = targetLayer.groupItems.add();
    imageItem.moveToBeginning(clippingGroup);
    pathItem.moveToBeginning(clippingGroup);
    clippingGroup.clipped = true;

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

// 選択更新
function updateSelection(clippingMasks) {
    var activeDoc = app.activeDocument;
    activeDoc.selection = null;
    for (var i = 0; i < clippingMasks.length; i++) {
        clippingMasks[i].parent.selected = true;
    }
}

// すべてパスか確認
function isAllPathItems(selectionArray) {
    for (var i = 0; i < selectionArray.length; i++) {
        if (selectionArray[i].typename !== "PathItem") {
            return false;
        }
    }
    return true;
}

// 最前面のパス（zOrderPosition が最大）を取得
function getFrontmostPath(pathArray) {
    var topPath = null;
    var highestZ = -1;

    for (var i = 0; i < pathArray.length; i++) {
        var item = pathArray[i];
        if (item.typename === "PathItem" && item.zOrderPosition > highestZ) {
            highestZ = item.zOrderPosition;
            topPath = item;
        }
    }

    return topPath;
}
