#target illustrator

/*

### 概要

選択したテキストの文字ツメを30%に設定します。

詳細は README を参照してください。

### Overview

Sets the tsume of the selected text to 30%.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "Tsume30simple";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/Tsume30simple.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/Tsume30simple.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 選択したテキストの文字ツメを30%に設定する */

(function () {
    if (app.documents.length === 0) return;
    var sel = app.activeDocument.selection;
    if (!sel) return;

    /* テキスト編集モードでは selection が TextRange 単体になる */
    var items = (sel.typename === "TextRange") ? [sel] : sel;

    for (var i = 0; i < items.length; i++) {
        var range = items[i].textRange || items[i];
        try {
            range.characterAttributes.Tsume = 30;
        } catch (e) {}
    }
})();
