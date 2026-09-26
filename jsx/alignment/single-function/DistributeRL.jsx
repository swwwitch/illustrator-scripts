#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

横並びに選択した複数オブジェクトのうち最も右のものを固定し、以降を環境設定［一般］の「キー入力」の値ぶんずつ左方向へ等間隔に再配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeRL.md

### Overview

Keeps the rightmost object of a horizontal selection fixed and redistributes the rest to the left at intervals of the Keyboard Increment.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeRL.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DistributeRL";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeRL.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeRL.md"; /* README (English) */

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
     * 最も右のオブジェクトを固定し、以降をキー入力の値ずつ左へ等間隔に再配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length < 1) return;

        var selectedObjects = app.activeDocument.selection;
        if (selectedObjects.length < 2) return;

        /* 環境設定［一般］の「キー入力」（cursorKeyLength、pt）を移動幅に使う / Keyboard Increment in points */
        var keyboardIncrementPt = app.preferences.getRealPreference("cursorKeyLength");

        /* 最も右を固定し、以降を keyboardIncrementPt ずつ左へ等間隔配置 / Keep the rightmost one, move the rest left step by step */
        var objectsRightToLeft = sortByLeftEdge(selectedObjects, true);
        for (var i = 1; i < objectsRightToLeft.length; i++) {
            objectsRightToLeft[i].translate(-i * keyboardIncrementPt, 0);
        }
    }

    main();

})();
