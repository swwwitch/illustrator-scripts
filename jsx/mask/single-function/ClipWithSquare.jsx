#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、正方形のクリッピングマスクで切り抜きます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipWithSquare.md

### Overview

Clips the selected objects with a square clipping mask.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipWithSquare.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipWithSquare";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-11-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipWithSquare.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipWithSquare.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var WORK_LAYER_NAME = "_clip_work";  /* 画像のレイヤーが編集できないときに使う作業レイヤー / work layer used when the image's layer is not editable */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 画像（配置画像／埋め込み画像）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} 画像なら true
     */
    function isImageItem(item) {
        return item.typename === "PlacedItem" || item.typename === "RasterItem";
    }

    /**
     * 作業レイヤーを取得、無ければ作成する（既存がロック・テンプレートでも常に編集可能にする）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 作業レイヤー
     */
    function getOrCreateWorkLayer(doc) {
        var workLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === WORK_LAYER_NAME) {
                workLayer = doc.layers[i];
                break;
            }
        }
        if (!workLayer) {
            workLayer = doc.layers.add();
            workLayer.name = WORK_LAYER_NAME;
        }
        workLayer.locked = false;
        workLayer.visible = true;
        workLayer.isTemplate = false;
        return workLayer;
    }

    /**
     * 画像の短辺に合わせた中央の正方形でクリッピングマスクを作り、選択する
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacedItem|RasterItem} image - 対象の画像
     * @returns {void}
     */
    function clipImageWithSquare(doc, image) {
        var bounds = image.visibleBounds;
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        var sideLength = Math.min(width, height);
        var centerX = bounds[0] + width / 2;
        var centerY = bounds[1] - height / 2;

        /* 画像のレイヤーが編集できなければ作業レイヤーを使う / Use the work layer when the image's layer is not editable */
        var parentLayer = image.layer;
        var targetLayer = (parentLayer.locked || parentLayer.isTemplate) ? getOrCreateWorkLayer(doc) : parentLayer;

        var square = targetLayer.pathItems.rectangle(centerY + sideLength / 2, centerX - sideLength / 2, sideLength, sideLength);
        var clipGroup = targetLayer.groupItems.add();

        image.moveToBeginning(clipGroup);
        square.moveToBeginning(clipGroup);
        clipGroup.clipped = true;
        clipGroup.selected = true;
    }

    /**
     * クリップグループなら中の画像だけを取り出す（マスクのパスなど画像以外は削除）
     * @param {GroupItem} clipGroup - クリップグループ
     * @returns {PageItem[]} 取り出した画像
     */
    function extractImagesFromClipGroup(clipGroup) {
        clipGroup.clipped = false;
        var images = [];
        for (var i = clipGroup.pageItems.length - 1; i >= 0; i--) {
            var pageItem = clipGroup.pageItems[i];
            if (isImageItem(pageItem)) {
                images.push(pageItem);
            } else {
                pageItem.remove();
            }
        }
        return images;
    }

    /**
     * 選択した画像（クリップグループ内の画像を含む）を正方形で切り抜く
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、length は文字数になる / With characters selected by the Type tool, selection is a TextRange whose length is the character count */
        if (selectedItems.typename === "TextRange" || !selectedItems.length) return;

        /* 作ったグループだけが選択に残るよう、先に選択を解除 / Clear the selection first so only new groups end up selected */
        doc.selection = null;

        for (var i = 0; i < selectedItems.length; i++) {
            var item = selectedItems[i];
            if (item.typename === "GroupItem" && item.clipped) {
                var images = extractImagesFromClipGroup(item);
                for (var j = 0; j < images.length; j++) {
                    clipImageWithSquare(doc, images[j]);
                }
            } else if (isImageItem(item)) {
                clipImageWithSquare(doc, item);
            }
        }
    }

    main();

})();
