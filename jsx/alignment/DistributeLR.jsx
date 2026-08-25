#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

横並びに選択した複数オブジェクトのうち最も左のものを固定し、以降を環境設定［一般］の「キー入力」の値ぶんずつ右方向へ等間隔に再配置します。

詳細は README を参照してください。

### Overview

Keeps the leftmost object of a horizontal selection fixed and redistributes the rest to the right at intervals of the Keyboard Increment.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DistributeLR";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeLR.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeLR.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    if (app.documents.length < 1) return

    var selectedObjects = app.activeDocument.selection
    if (selectedObjects.length < 2) return

    // 環境設定［一般］の「キー増加」増分（cursorKeyLength、pt）を移動幅に使う
    var keyboardIncrementPt = app.preferences.getRealPreference("cursorKeyLength")

    // 横並び → 最も左を固定し、以降を keyboardIncrementPt ずつ右へ等間隔配置
    var objectsLeftToRight = sortByHorizontalPosition(selectedObjects)
    for (var i = 1; i < objectsLeftToRight.length; i++) {
        objectsLeftToRight[i].translate(i * keyboardIncrementPt, 0)
    }

    /**
     * 選択オブジェクトを左端X（position[0]）の昇順で並べ替えた新しい配列を返す
     * @param {Array<PageItem>} objects - 並べ替える対象のオブジェクト
     * @returns {Array<PageItem>} 左端Xの昇順に並べ替えた新しい配列
     */
    function sortByHorizontalPosition(objects) {
        var sortedObjects = []
        for (var i = 0; i < objects.length; i++) sortedObjects.push(objects[i])
        sortedObjects.sort(function (a, b) {
            return a.position[0] - b.position[0]
        })
        return sortedObjects
    }
})()
