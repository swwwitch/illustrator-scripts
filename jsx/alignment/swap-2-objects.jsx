#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

2つのオブジェクトを選択しているとき、それぞれの中心位置を入れ替えます。
通常オブジェクトは visibleBounds、クリップグループはマスクパスの geometricBounds を基準にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/swap-2-objects.md

### Overview

Swaps the center positions of two selected objects.
Ordinary objects use their visibleBounds, while clipping groups use the geometricBounds of the mask path.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/swap-2-objects.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "swap-2-objects";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/swap-2-objects.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/swap-2-objects.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /**
     * オブジェクトの中心座標を返す
     *
     * クリップグループはマスクパスの geometricBounds、それ以外は visibleBounds を基準にする。
     * マスクパスが見つからないクリップグループは visibleBounds にフォールバックする。
     * @param {PageItem} targetItem - 中心を求める対象のオブジェクト
     * @returns {number[]} 中心の [x, y] 座標
     */
    function getCenterPoint(targetItem) {
        var referenceBounds = targetItem.visibleBounds;

        if (targetItem.typename === "GroupItem" && targetItem.clipped) {
            /* クリップグループはマスクパスの範囲を基準にする / clipping groups follow the mask path */
            for (var i = 0; i < targetItem.pageItems.length; i++) {
                if (targetItem.pageItems[i].clipping) {
                    referenceBounds = targetItem.pageItems[i].geometricBounds;
                    break;
                }
            }
        }

        return [(referenceBounds[0] + referenceBounds[2]) / 2, (referenceBounds[1] + referenceBounds[3]) / 2];
    }

    /**
     * 2つのオブジェクトの中心位置を入れ替える
     * @param {PageItem} firstObject - 入れ替える一方のオブジェクト
     * @param {PageItem} secondObject - 入れ替えるもう一方のオブジェクト
     * @returns {void}
     */
    function swapObjectsByCenter(firstObject, secondObject) {
        var firstCenter = getCenterPoint(firstObject);
        var secondCenter = getCenterPoint(secondObject);

        /* 片方の移動量を求め、もう片方はその逆向きに動かす / one offset, applied in both directions */
        var dx = secondCenter[0] - firstCenter[0];
        var dy = secondCenter[1] - firstCenter[1];

        firstObject.translate(dx, dy);
        secondObject.translate(-dx, -dy);
    }

    /**
     * ドキュメントと選択を検証し、選択した2つのオブジェクトの中心位置を入れ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var selectedObjects = app.activeDocument.selection;
        if (!selectedObjects || selectedObjects.length !== 2) {
            alert("2つのオブジェクトを選択してください。");
            return;
        }

        swapObjectsByCenter(selectedObjects[0], selectedObjects[1]);
    }

    main();

})();
