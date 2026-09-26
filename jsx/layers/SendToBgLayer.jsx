#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、重ね順を維持したまま「bg」レイヤーへ移動して最背面に配置します。
「bg」レイヤーは自動的に作成され、処理後にロックされます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SendToBgLayer.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf7c1e8a0f0c7

### Overview

Moves the selected objects to a "bg" layer, preserving their stacking order, and sends that layer to the back.
The "bg" layer is created automatically and locked once the move is done.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SendToBgLayer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SendToBgLayer";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2024-06-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SendToBgLayer.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SendToBgLayer.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf7c1e8a0f0c7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var TARGET_LAYER_NAME = "bg";  /* 移動先のレイヤー名 / name of the destination layer */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 名前でレイヤーを探し、無ければ作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 見つけた、または作ったレイヤー
     */
    function getOrCreateLayer(doc, layerName) {
        try {
            /* 見つからないと例外になる / getByName throws when the layer does not exist */
            return doc.layers.getByName(layerName);
        } catch (e) {
            var newLayer = doc.layers.add();
            newLayer.name = layerName;
            return newLayer;
        }
    }

    /**
     * 選択オブジェクトを重ね順を保ったまま背景レイヤーへ移し、レイヤーを最背面にしてロックする
     * @returns {void}
     */
    function main() {
        var activeDoc = app.activeDocument;
        var originalLayer = activeDoc.activeLayer; /* 元のアクティブレイヤー / Remember the active layer */

        /* 表示状態を控えてから、表示・ロック解除 / Note the visibility, then show and unlock */
        var bgLayer = getOrCreateLayer(activeDoc, TARGET_LAYER_NAME);
        var wasHidden = !bgLayer.visible;
        if (wasHidden) bgLayer.visible = true;
        bgLayer.locked = false;

        /* 選択オブジェクトを重ね順を保ったまま移動 / Move the selection, keeping its stacking order */
        var selectedItems = activeDoc.selection;
        for (var i = 0; i < selectedItems.length; i++) {
            try {
                selectedItems[i].move(bgLayer, ElementPlacement.PLACEATEND);
            } catch (e) {
                /* 移動できないもの（ロックされた親の中など）は飛ばす / Skip items that cannot move */
            }
        }

        /* レイヤーを最背面へ、表示状態を戻してロック / Send to back, restore visibility and lock */
        bgLayer.zOrder(ZOrderMethod.SENDTOBACK);
        if (wasHidden) bgLayer.visible = false;
        bgLayer.locked = true;

        activeDoc.activeLayer = originalLayer;
    }

    main();

})();
