#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、正方形のクリッピングマスクで切り抜きます。

詳細は README を参照してください。

### Overview

Clips the selected objects with a square clipping mask.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipWithSquare";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-11-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-08-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipWithSquare.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipWithSquare.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/*
作業用レイヤーを取得/作成 / Get or create a reusable work layer
- 名称 / Name: _clip_work
- 既存がロック/テンプレでもこのレイヤーは常に編集可能に設定
*/
function getOrCreateWorkLayer(doc) {
    var name = "_clip_work";
    var lyr = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) {
            lyr = doc.layers[i];
            break;
        }
    }
    if (!lyr) {
        lyr = doc.layers.add();
        lyr.name = name;
    }
    // ensure editable
    lyr.locked = false;
    lyr.visible = true;
    lyr.isTemplate = false;
    return lyr;
}

/*
画像に最小正方形を追加し、クリッピンググループを作成 / Build a clipping group with the minimal square
- 入力 / Input: doc (Document), image (PlacedItem|RasterItem), groups (Array)
- 動作 / Behavior: visibleBounds から正方形を作成→画像と同グループに配置→グループをクリップ化
*/
function processImage(doc, image) {
    var bounds = image.visibleBounds;
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];
    var sideLength = Math.min(width, height);
    var centerX = bounds[0] + width / 2;
    var centerY = bounds[1] - height / 2;

    // 現在のレイヤーが編集可能か確認 / Check if current layer is editable
    var parentLayer = image.layer;
    var isLockedOrTemplate = parentLayer.locked || parentLayer.isTemplate;

    // 編集可能なレイヤーを使用 / Choose a writable layer
    var targetLayer = isLockedOrTemplate ? getOrCreateWorkLayer(doc) : parentLayer;

    // 四角形とグループをレイヤーに追加 / Create square and group
    var square = targetLayer.pathItems.rectangle(centerY + sideLength / 2, centerX - sideLength / 2, sideLength, sideLength);
    var group = targetLayer.groupItems.add();

    image.moveToBeginning(group);
    square.moveToBeginning(group);
    group.clipped = true;
    group.selected = true; // 生成直後に即選択 / Select immediately after creation
}

function main() {
    // 安全ガード / Safety guard
    if (!app.documents.length || !app.selection.length) {
        return;
    }
    var doc = app.activeDocument;
    var selectedItems = app.selection;
    doc.selection = null; // 先に選択をクリア / Clear selection first

    for (var i = 0; i < selectedItems.length; i++) {
        var item = selectedItems[i];

        // クリッピングマスクの処理 / Handle clipping mask groups
        if (item.typename === 'GroupItem' && item.clipped) {
            item.clipped = false;

            var itemsToProcess = [];
            for (var j = item.pageItems.length - 1; j >= 0; j--) {
                var pageItem = item.pageItems[j];
                if (pageItem.typename === 'PlacedItem' || pageItem.typename === 'RasterItem') {
                    itemsToProcess.push(pageItem);
                } else {
                    pageItem.remove();
                }
            }

            // 解除後のグループ内にある画像を処理 / Process images extracted from the released group
            for (var k = 0; k < itemsToProcess.length; k++) {
                var imageItem = itemsToProcess[k];
                processImage(doc, imageItem);
            }
        }
        // 単体の配置/埋め込み画像を処理 / Handle standalone placed/embedded images
        else if (item.typename === 'PlacedItem' || item.typename === 'RasterItem') {
            processImage(doc, item);
        }
    }
}

main();
