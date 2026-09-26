#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

横並びに選択した複数オブジェクトのうち最も左のものを固定し、以降を環境設定［一般］の「キー入力」の値ぶんずつ左へ動かして間隔を狭めます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeLL.md

### Overview

Keeps the leftmost object of a horizontal selection fixed and moves the rest left by the Keyboard Increment, tightening the spacing.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeLL.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DistributeLL";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeLL.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeLL.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 並べ替え / Sorting
    // =========================================

    /**
     * 左端X（position[0]）で並べ替えた新しい配列を返す
     * @param {PageItem[]} targetObjects - 並べ替える対象のオブジェクト
     * @param {boolean} rightFirst - true なら右から左（降順）、false なら左から右（昇順）
     * @returns {PageItem[]} 並べ替えた新しい配列
     */
    function sortByLeftEdge(targetObjects, rightFirst) {
        var sortedObjects = [];
        for (var i = 0; i < targetObjects.length; i++) sortedObjects.push(targetObjects[i]);
        sortedObjects.sort(function (itemA, itemB) {
            var leftDelta = itemA.position[0] - itemB.position[0];
            return rightFirst ? -leftDelta : leftDelta;
        });
        return sortedObjects;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 最も左のオブジェクトを固定し、以降をキー入力の値ずつ左へ寄せて間隔を狭める
     * @returns {void}
     */
    function main() {
        if (app.documents.length < 1) return;

        var selectedObjects = app.activeDocument.selection;
        if (selectedObjects.length < 2) return;

        /* 環境設定［一般］の「キー入力」（cursorKeyLength、pt）を移動幅に使う / Keyboard Increment in points */
        var keyboardIncrementPt = app.preferences.getRealPreference("cursorKeyLength");

        /* 最も左を固定し、以降を keyboardIncrementPt ずつ左へ寄せて間隔を狭める / Keep the leftmost one, move the rest left to tighten the spacing */
        var objectsLeftToRight = sortByLeftEdge(selectedObjects, false);
        for (var i = 1; i < objectsLeftToRight.length; i++) {
            objectsLeftToRight[i].translate(-i * keyboardIncrementPt, 0);
        }
    }

    main();

})();
