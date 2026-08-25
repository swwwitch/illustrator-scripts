
/*

### 概要

「下絵」という名前のレイヤーを探し、そのレイヤーをテンプレート化します。

詳細は README を参照してください。

### Overview

Finds the layer named "下絵" and turns it into a template layer.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardLayerOrganizer-v2";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardLayerOrganizer-v2.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardLayerOrganizer-v2.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function setLayerTemplate(doc, name, on) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                doc.layers[i].template = on;
                return true;
            }
        }
        return false;
    }

    setLayerTemplate(app.activeDocument, "下絵", true);  // テンプレート化

})();
