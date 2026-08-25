#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

コピー済みのオブジェクトを、開いているすべてのドキュメントへ同じ位置に貼り付けます（pasteInPlace）。

詳細は README を参照してください。

### Overview

Pastes the copied objects into every open document at the same position, using Paste in Place.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PasteInAllDocs";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PasteInAllDocs.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PasteInAllDocs.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function main() {
        try {
            // ドキュメントが1つも開いていない場合は終了
            if (app.documents.length === 0) {
                alert("ドキュメントが開かれていません。\rNo documents are open.");
                return;
            }

            // 最初のドキュメントを保持（コピー元とする）
            var sourceDoc = app.activeDocument;

            // コピーされていないとペーストに失敗するのでチェック
            // Illustratorではクリップボードの空判定が難しいため try-catch で対応
            try {
                sourceDoc.activate();
                app.executeMenuCommand("pasteInPlace");
                // ペーストされたものを削除してクリップボードを維持
                app.cut();
            } catch (e) {
                alert("コピーされたオブジェクトがありません。\rNo objects copied.");
                return;
            }

            // 各ドキュメントへペースト
            for (var i = 0; i < app.documents.length; i++) {
                var doc = app.documents[i];
                doc.activate();
                try {
                    app.executeMenuCommand("pasteInPlace");
                } catch (e) {
                    alert("ペーストに失敗しました: " + doc.name + "\rFailed to paste in: " + doc.name);
                }
            }

            // 元のドキュメントを再度アクティブに
            sourceDoc.activate();
            alert("コピーしたオブジェクトをすべてのドキュメントにペーストしました。\rObjects pasted into all documents.");

        } catch (err) {
            alert("エラーが発生しました:\r" + err);
        }
    }

    main();

})();
