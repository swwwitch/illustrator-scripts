#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

現在のアートボードを対象に、トンボを作成します。

詳細は README を参照してください。

### Overview

Creates trim marks for the current artboard.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddTrimMarkToCurrentArtboard"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-02-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMarkToCurrentArtboard.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMarkToCurrentArtboard.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function main() {
        var doc = app.activeDocument;
        var targetObj = null;
        var trimLayer = null;

        /* 「トンボ」レイヤーを取得（なければ作成） / Get "Trim" layer (create if not exists) */
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "トンボ") {
                trimLayer = doc.layers[i];
                break;
            }
        }
        if (!trimLayer) {
            trimLayer = doc.layers.add();
            trimLayer.name = "トンボ";
        }
        var wasLocked = trimLayer.locked;
        if (wasLocked) {
            trimLayer.locked = false;
        }

        try {
            /* 日本式トンボをONに設定 / Enable Japanese-style trim marks */
            app.preferences.setBooleanPreference('cropMarkStyle', 1);

            /* 常にアクティブなアートボードを基に処理 / Always use the active artboard */
            var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            var rect = artboard.artboardRect; // [左, 上, 右, 下]

            /* 「トンボ」レイヤー上にアートボード矩形を作成 / Create artboard rectangle on "Trim" layer */
            targetObj = trimLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
            targetObj.filled = false;
            targetObj.stroked = false;

            /* 作成オブジェクトを選択状態にする / Select the created object */
            doc.selection = [targetObj];

            /* トリムマークを作成 / Create trim marks */
            app.executeMenuCommand('TrimMark v25');

            /* 作成オブジェクトをガイド化する / Convert the created object to a guide */
            targetObj.guides = true;
        } finally {
            doc.selection = null;

            if (trimLayer) {
                trimLayer.locked = wasLocked ? true : false;
            }
        }
    }

    main();

})();
