#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// 概要 / Overview
// =========================================
/*
DistributeLL.jsx

複数選択（横並び）で、最も左のオブジェクトを固定し、以降を環境設定［一般］の
「キー増加」の値ぶんずつ左方向へ移動して間隔を狭める。

移動に使う値は、環境設定［一般］の「キー増加」増分（cursorKeyLength）。
この値は pt で格納されているため単位換算は不要。
*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DistributeLL";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    if (app.documents.length < 1) return

    var selectedObjects = app.activeDocument.selection
    if (selectedObjects.length < 2) return

    // 環境設定［一般］の「キー増加」増分（cursorKeyLength、pt）を移動幅に使う
    var keyboardIncrementPt = app.preferences.getRealPreference("cursorKeyLength")

    // 横並び → 最も左を固定し、以降を keyboardIncrementPt ずつ左へ寄せて間隔を狭める
    var objectsLeftToRight = sortByHorizontalPosition(selectedObjects)
    for (var i = 1; i < objectsLeftToRight.length; i++) {
        objectsLeftToRight[i].translate(-i * keyboardIncrementPt, 0)
    }

    // 選択オブジェクトを左端 X（position[0]）の昇順で並べ替えた新しい配列を返す
    function sortByHorizontalPosition(objects) {
        var sortedObjects = []
        for (var i = 0; i < objects.length; i++) sortedObjects.push(objects[i])
        sortedObjects.sort(function (a, b) {
            return a.position[0] - b.position[0]
        })
        return sortedObjects
    }
})()
