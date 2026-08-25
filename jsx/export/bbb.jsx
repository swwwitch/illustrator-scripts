#target illustrator

/*

### 概要

「Guides Preview for Trim View」レイヤーを探し、ロックと非表示を解除してから削除し、結果をアラートで報告します。

詳細は README を参照してください。

### Overview

Finds the "Guides Preview for Trim View" layer, unlocks and unhides it, removes it, and reports the result in an alert.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "bbb";                          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/bbb.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/bbb.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    if (app.documents.length === 0) { alert("ドキュメントがありません"); return; }
    var doc = app.activeDocument;
    var TARGET = "Guides Preview for Trim View";
    function normalize(n) { return n.replace(/^\s*\*?\s*/, "").replace(/\s+$/, ""); }

    var log = [];
    for (var i = doc.layers.length - 1; i >= 0; i--) {
        var L = doc.layers[i];
        if (normalize(L.name) !== TARGET) { continue; }
        log.push("対象発見: [" + L.name + "] locked=" + L.locked + " visible=" + L.visible);
        try {
            L.locked = false;
            L.visible = true;
            L.remove();
            log.push("  → remove() 成功");
        } catch (e) {
            log.push("  → remove() 失敗: " + e.message + " (line " + e.line + ")");
        }
    }
    if (log.length === 0) { log.push("対象が見つかりませんでした"); }
    alert(log.join("\n"));
})();
