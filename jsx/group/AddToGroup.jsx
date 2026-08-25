#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトをいったんグループ解除してから、あらためて1つのグループにまとめます。既存のグループへオブジェクトを加えたいときに使います。

詳細は README を参照してください。

### Overview

Ungroups the selection once and then groups it again as a single group. Use it to add objects to an existing group.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddToGroup";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddToGroup.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddToGroup.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var doc = app.documents.length && app.activeDocument;
    if (!doc) return;

    var sel = doc.selection;

    if (sel.length < 2) return;

    // ungroup
    app.executeMenuCommand('ungroup');

    // group
    app.executeMenuCommand('group');

})();
