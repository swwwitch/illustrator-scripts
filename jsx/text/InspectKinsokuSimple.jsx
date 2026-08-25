#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームの段落で使われている禁則処理の値を列挙して表示します。

詳細は README を参照してください。

### Overview

Lists the kinsoku (line-breaking) settings used by the paragraphs of the selected text frames.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InspectKinsokuSimple";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InspectKinsokuSimple.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InspectKinsokuSimple.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// 選択したテキストフレームの段落で使われている禁則の値を列挙して確認する
(function () {
    var selection = app.activeDocument.selection;
    var kinsokuSet = {};

    for (var i = 0; i < selection.length; i++) {
        if (selection[i].constructor.name !== "TextFrame") continue;
        var paragraphs = selection[i].textRange.paragraphs;
        for (var j = 0; j < paragraphs.length; j++) {
            try {
                kinsokuSet[String(paragraphs[j].paragraphAttributes.kinsoku)] = true;
            } catch (e) {
                // 禁則「なし」は属性が undefined 扱いで Error 9563 を投げる
                kinsokuSet[e.number === 9563 ? "なし" : "ERROR: " + e.message] = true;
            }
        }
    }

    var report = "検出された禁則値:\n";
    for (var value in kinsokuSet) {
        report += "  → \"" + value + "\"\n";
    }
    alert(report);
})();
