#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの入れ子グループをすべて解除してから、1つのグループにまとめ直します。

詳細は README を参照してください。

### Overview

Releases every nested group in the selection and then regroups everything as a single group.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FlattenGroup";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenGroup.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenGroup.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

var doc = app.documents.length && app.activeDocument;
if (!doc) return;

var sel = doc.selection;
if (!sel.length) return;

    if (!sel || sel.length < 1) return;

    // ungroup all
    app.executeMenuCommand('ungroupAll');

    // group
    app.executeMenuCommand('group');

})();
