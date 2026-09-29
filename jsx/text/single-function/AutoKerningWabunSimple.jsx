#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの自動カーニング方式を「オプティカル」に設定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerningWabunSimple.md

### Overview

Sets the auto-kerning method of the selected text to Optical.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerningWabunSimple.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoKerningWabunSimple";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerningWabunSimple.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerningWabunSimple.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var selectedItems = app.activeDocument.selection;
    /* 文字ツールで文字を選択しているときは、その文字だけを「オプティカル」に（TextRange の length は文字数）
       With characters selected by the Type tool, set only those to Optical (a TextRange's length is its character count) */
    if (selectedItems.typename === "TextRange") {
        selectedItems.characterAttributes.kerningMethod = AutoKernType.OPTICAL;
        return;
    }
    /* 選択中のテキストオブジェクトを「オプティカル」に（テキスト以外は飛ばす）/ Set the selected text objects to Optical, skipping others */
    for (var i = 0; i < selectedItems.length; i++) {
        if (selectedItems[i].typename !== "TextFrame") continue;
        selectedItems[i].textRange.characterAttributes.kerningMethod = AutoKernType.OPTICAL;
    }

})();
