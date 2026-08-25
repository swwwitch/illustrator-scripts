#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの自動カーニング方式を「和文等幅」に設定します。

詳細は README を参照してください。

### Overview

Sets the auto-kerning method of the selected text to Metrics (Roman Only).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoKerningWabunSimple";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerningWabunSimple.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerningWabunSimple.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var selection = app.activeDocument.selection;
    for (var i = 0; i < selection.length; i++) {
        selection[i].textRange.characterAttributes.kerningMethod = AutoKernType.OPTICAL;
    }

})();
