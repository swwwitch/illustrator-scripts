
/*

### 概要

選択したパスアイテム（グループ内も含む）の線端を、丸型線端に設定します。

詳細は README を参照してください。

### Overview

Sets the stroke cap of the selected path items, including those inside groups, to a round cap.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "丸型先端にする";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-08-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-12";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/丸型先端にする.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/丸型先端にする.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* パスアイテムに丸型線端を適用する再帰関数 / Apply a round cap to path items recursively */
function applyProjectingCap(item) {
    if (item.typename === "GroupItem") {
        // グループの場合、子アイテムを再帰的に処理
        for (var i = 0; i < item.pageItems.length; i++) {
            applyProjectingCap(item.pageItems[i]);
        }
    } else if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
        if (item.stroked && item.strokeCap !== StrokeCap.ROUNDENDCAP) {
            item.strokeCap = StrokeCap.ROUNDENDCAP;
        }
    }
}

if (app.documents.length > 0) {
    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length > 0) {
        for (var i = 0; i < sel.length; i++) {
            applyProjectingCap(sel[i]);
        }
        // alert("選択されたパスアイテムの線端を突出先端に設定しました。");
    } else {
        alert("パスアイテムが選択されていません。パスアイテムを選択してください。");
    }
} else {
    alert("ドキュメントが開かれていません。ドキュメントを開いてください。");
}
