#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、1つずつ個別のグループにします。

詳細は README を参照してください。

### Overview

Puts each selected object into its own group.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEachSelection";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEachSelection.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEachSelection.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/*
  選択したオブジェクトをそれぞれ個別にグループ化するスクリプト
*/

(function() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    // オブジェクトが選択されているか確認
    if (sel.length === 0) {
        alert("オブジェクトが選択されていません。");
        return;
    }

    // エラー回避のため、選択順序を逆にして処理（インデックスずれ防止）
    // ただし、move()を使う場合は元の配列参照が生きていると安全ではないため、
    // いったん配列にコピーしてから処理するのが一般的ですが、
    // 今回は単純な移動なのでループで回します。

    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];

        // すでにグループ化されているものをさらにグループ化するのか、
        // 単体のパスなどをグループに入れるのかに関わらず、
        // 新しいグループを作成してそこに移動させます。

        // 新しいグループコンテナを作成（元のオブジェクトの親階層に作るのが安全）
        // item.layer だとレイヤー直下になるため、item.parent を参照
        var parentContainer = item.parent;
        var newGroup = parentContainer.groupItems.add();
        
        // オブジェクトの重なり順（Zオーダー）を維持するための工夫
        // move(relativeObject, elementPlacement) を使用
        // 以前の場所の「直前」にグループを移動させてから、中身を入れる
        newGroup.move(item, ElementPlacement.PLACEBEFORE);
        
        // オブジェクトを新グループ内に移動
        item.move(newGroup, ElementPlacement.PLACEATEND);
    }
    
    // 処理完了後に選択状態を解除したければ以下を有効化
    // doc.selection = null;

})();
