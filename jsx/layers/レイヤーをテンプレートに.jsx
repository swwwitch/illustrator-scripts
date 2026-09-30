#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

「下絵」という名前のレイヤーを探し、そのレイヤーをテンプレート化します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/レイヤーをテンプレートに.md

### Overview

Finds the layer named "下絵" and turns it into a template layer.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/レイヤーをテンプレートに.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "レイヤーをテンプレートに";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/レイヤーをテンプレートに.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/レイヤーをテンプレートに.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var TARGET_LAYER_NAME = "下絵";  /* テンプレートにするレイヤー名 / name of the layer to turn into a template */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 名前が一致する最初のトップレベルレイヤーのテンプレート属性を切り替える
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @param {boolean} isTemplate - テンプレートにするなら true
     * @returns {boolean} 対象のレイヤーが見つかれば true
     */
    function setLayerTemplate(doc, layerName, isTemplate) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                doc.layers[i].template = isTemplate;
                return true;
            }
        }
        return false;
    }

    setLayerTemplate(app.activeDocument, TARGET_LAYER_NAME, true);

})();
