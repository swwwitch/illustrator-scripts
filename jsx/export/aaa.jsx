#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のレイヤーとサブレイヤーを階層順にたどり、名前・ロック状態・表示状態をアラートで一覧表示します。

詳細は README を参照してください。

### Overview

Walks the document's layers and sublayers in hierarchy order and lists each name with its locked and visible state in an alert.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "aaa";                          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/aaa.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/aaa.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    if (app.documents.length === 0) { alert("ドキュメントがありません"); return; }
    var lines = [];
    function walk(layers, depth) {
        for (var i = 0; i < layers.length; i++) {
            var L = layers[i];
            var indent = "";
            for (var d = 0; d < depth; d++) { indent += "  "; }
            lines.push(indent + "[" + L.name + "]  locked=" + L.locked + " visible=" + L.visible);
            if (L.layers && L.layers.length > 0) { walk(L.layers, depth + 1); }
        }
    }
    walk(app.activeDocument.layers, 0);
    alert(lines.join("\n"));
})();
