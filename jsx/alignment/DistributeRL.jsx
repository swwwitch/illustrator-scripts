#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

横並びに選択した複数オブジェクトのうち最も右のものを固定し、以降を環境設定［一般］の「キー入力」の値ぶんずつ左方向へ等間隔に再配置します。

詳細は README を参照してください。

### Overview

Keeps the rightmost object of a horizontal selection fixed and redistributes the rest to the left at intervals of the Keyboard Increment.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DistributeRL";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeRL.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeRL.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    if (app.documents.length < 1) return

    var selectedObjects = app.activeDocument.selection
    if (selectedObjects.length < 2) return

    // 環境設定［一般］の「キー増加」増分（cursorKeyLength、pt）を移動幅に使う
    var keyboardIncrementPt = app.preferences.getRealPreference("cursorKeyLength")

    // 横並び → 最も右を固定し、以降を keyboardIncrementPt ずつ左へ等間隔配置
    var objectsRightToLeft = sortByHorizontalPosition(selectedObjects)
    for (var i = 1; i < objectsRightToLeft.length; i++) {
        objectsRightToLeft[i].translate(-i * keyboardIncrementPt, 0)
    }

    // 選択オブジェクトを左端 X（position[0]）の降順で並べ替えた新しい配列を返す
    function sortByHorizontalPosition(objects) {
        var sortedObjects = []
        for (var i = 0; i < objects.length; i++) sortedObjects.push(objects[i])
        sortedObjects.sort(function (a, b) {
            return b.position[0] - a.position[0]
        })
        return sortedObjects
    }
})()
