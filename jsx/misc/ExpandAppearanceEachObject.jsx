#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択しているオブジェクトごとに、アピアランスを分割します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandAppearanceEachObject.md

### Overview

Expands the appearance of each selected object individually.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandAppearanceEachObject.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExpandAppearanceEachObject";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandAppearanceEachObject.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandAppearanceEachObject.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 選択しているオブジェクトを1つずつ選び直し、アピアランスを分割する / Reselect each selected object in turn and expand its appearance */
    var doc = app.activeDocument;
    var selectedItems = doc.selection;

    if (selectedItems.length === 0) {
        alert("オブジェクトを選択してください。");
    } else {
        for (var i = selectedItems.length - 1; i >= 0; i--) {
            doc.selection = null;
            selectedItems[i].selected = true;
            app.executeMenuCommand('expandStyle');
        }
    }

})();
