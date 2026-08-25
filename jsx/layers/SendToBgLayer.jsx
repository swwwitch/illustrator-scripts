#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、重ね順を維持したまま「bg」レイヤーへ移動して最背面に配置します。
「bg」レイヤーは自動的に作成され、処理後にロックされます。

詳細は README を参照してください。

### Overview

Moves the selected objects to a "bg" layer, preserving their stacking order, and sends that layer to the back.
The "bg" layer is created automatically and locked once the move is done.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SendToBgLayer";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2024-06-25";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SendToBgLayer.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SendToBgLayer.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var TARGET_LAYER_NAME = "bg"; // 利用者が変更可能なレイヤー名 / User-editable target layer name

    /* メイン処理 / Main process */
    function main() {
        var activeDoc = app.activeDocument;
        var originalLayer = activeDoc.activeLayer; // 元のアクティブレイヤーを記憶

        // 「bg」レイヤーが存在するか確認し、なければ作成し、元の可視状態を記憶
        var bgLayer;
        try {
            bgLayer = activeDoc.layers.getByName(TARGET_LAYER_NAME);
        } catch (e) {
            bgLayer = activeDoc.layers.add();
            bgLayer.name = TARGET_LAYER_NAME;
        }
        var wasHidden = !bgLayer.visible;
        if (wasHidden) bgLayer.visible = true;
        bgLayer.locked = false; // ロックを解除

        var selectedItems;
        try {
            selectedItems = activeDoc.selection; // 現在の選択オブジェクトを取得
        } catch (e) {
            selectedItems = [];
        }

        if (selectedItems && selectedItems.length > 0) {
            // 選択オブジェクトを重ね順を維持したまま「bg」レイヤーに移動
            for (var i = 0; i < selectedItems.length; i++) {
                try {
                    selectedItems[i].move(bgLayer, ElementPlacement.PLACEATEND);
                } catch (e) {
                    // エラーがあっても処理を続行
                }
            }
        }

        bgLayer.zOrder(ZOrderMethod.SENDTOBACK); // 「bg」レイヤー自体を最背面に
        if (wasHidden) bgLayer.visible = false;
        bgLayer.locked = true; // 「bg」レイヤーを再ロック

        // 処理終了後、元のレイヤーを再アクティブ化
        activeDoc.activeLayer = originalLayer;
    }

    main();

})();
