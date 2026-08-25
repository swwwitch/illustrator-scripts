/*

### 概要

選択しているオブジェクトごとに、アピアランスを分割します。

詳細は README を参照してください。

### Overview

Expands the appearance of each selected object individually.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExpandAppearanceEachObject";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandAppearanceEachObject.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandAppearanceEachObject.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // Illustrator用のJavaScript
    // 選択しているオブジェクトごとにアピアランスを分割

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert("オブジェクトを選択してください。");
    } else {
        for (var i = sel.length - 1; i >= 0; i--) {
            doc.selection = null;
            sel[i].selected = true;
            app.executeMenuCommand('expandStyle');
            // app.executeMenuCommand('group');
        }
    }

})();
