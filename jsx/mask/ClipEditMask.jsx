#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択がクリップグループのときはマスク編集モードへ切り替え、そうでないときは重なり合うオブジェクトごとにクリッピングマスクを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipEditMask.md

### Overview

Enters mask-edit mode when the selection is a clipping group, and otherwise builds a clipping mask for each cluster of overlapping objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipEditMask.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipEditMask";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipEditMask.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipEditMask.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 重なりの判定 / Overlap detection
    // =========================================

    /**
     * 2つの矩形（geometricBounds）が重なっているかを判定する
     * @param {number[]} boundsA - [left, top, right, bottom]（top > bottom）
     * @param {number[]} boundsB - [left, top, right, bottom]（top > bottom）
     * @returns {boolean} 重なっていれば true
     */
    function isBoundsOverlapping(boundsA, boundsB) {
        var horizontalOverlap = (boundsA[0] < boundsB[2]) && (boundsB[0] < boundsA[2]);
        var verticalOverlap = (boundsA[3] < boundsB[1]) && (boundsB[3] < boundsA[1]);
        return horizontalOverlap && verticalOverlap;
    }

    /**
     * オブジェクトを重なりで連結成分（クラスタ）に分割する
     * 直接重ならなくても、間のオブジェクトを介して繋がっていれば同じクラスタになる
     * @param {PageItem[]} items - 対象のオブジェクト
     * @returns {PageItem[][]} クラスタごとのオブジェクト
     */
    function clusterItemsByOverlap(items) {
        var bounds = [];
        var visited = [];
        for (var i = 0; i < items.length; i++) {
            bounds.push(items[i].geometricBounds);
            visited.push(false);
        }

        var clusters = [];
        for (var start = 0; start < items.length; start++) {
            if (visited[start]) continue;

            /* start を起点に、重なりで繋がるものを幅優先で集める / Breadth-first from start over overlaps */
            var cluster = [];
            var queue = [start];
            visited[start] = true;

            while (queue.length > 0) {
                var current = queue.shift();
                cluster.push(items[current]);

                for (var other = 0; other < items.length; other++) {
                    if (!visited[other] && isBoundsOverlapping(bounds[current], bounds[other])) {
                        visited[other] = true;
                        queue.push(other);
                    }
                }
            }

            clusters.push(cluster);
        }

        return clusters;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 指定したオブジェクトだけを選択状態にする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 選択するオブジェクト
     * @returns {void}
     */
    function selectOnly(doc, items) {
        doc.selection = null;
        for (var i = 0; i < items.length; i++) {
            items[i].selected = true;
        }
    }

    /**
     * 1つ選択ならマスク編集／作成、複数なら重なりの塊ごとにクリッピングマスクを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (currentSelection.length === 0) return;

        if (currentSelection.length === 1) {
            var selectedItem = currentSelection[0];
            if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
                /* ［マスクを編集］（2回で編集モードに入る）/ Edit Mask, run twice to enter the mode */
                app.executeMenuCommand("editMask");
                app.executeMenuCommand("editMask");
            } else {
                app.executeMenuCommand("makeMask");
            }
            return;
        }

        /* 2つ以上：重なりで塊に分けてから、塊ごとにマスク / Two or more: cluster by overlap, then mask each cluster */
        var selectedItems = [];
        for (var i = 0; i < currentSelection.length; i++) {
            selectedItems.push(currentSelection[i]);
        }

        var clusters = clusterItemsByOverlap(selectedItems);
        if (clusters.length === 1) {
            /* すべてが1つの塊 → そのまま1回だけマスク / One cluster: mask once as is */
            app.executeMenuCommand("makeMask");
            return;
        }

        for (var c = 0; c < clusters.length; c++) {
            selectOnly(doc, clusters[c]);
            app.executeMenuCommand("makeMask");
        }
    }

    main();

})();
