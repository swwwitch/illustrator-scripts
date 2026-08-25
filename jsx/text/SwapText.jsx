#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中の2つのテキストオブジェクトの文字列（contents）を入れ替えます。

詳細は README を参照してください。

### Overview

Swaps the contents of two selected text objects.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapText";                     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapText.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapText.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n071e09af28a7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
    } else {
        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (selectedItems.length !== 2) {
            alert("テキストオブジェクトを2つ選択してください。");
        } else if (selectedItems[0].typename !== "TextFrame" || selectedItems[1].typename !== "TextFrame") {
            alert("選択した2つは両方ともテキストオブジェクトである必要があります。");
        } else {
            var firstTextFrame = selectedItems[0];
            var secondTextFrame = selectedItems[1];

            var firstContents = firstTextFrame.contents;
            var secondContents = secondTextFrame.contents;

            firstTextFrame.contents = secondContents;
            secondTextFrame.contents = firstContents;
        }
    }

})();
