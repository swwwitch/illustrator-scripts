#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

2つのオブジェクトを選択しているとき、それぞれの中心位置を入れ替えます。
通常オブジェクトは visibleBounds、クリップグループはマスクパスの geometricBounds を基準にします。

詳細は README を参照してください。

### Overview

Swaps the center positions of two selected objects.
Ordinary objects use their visibleBounds, while clipping groups use the geometricBounds of the mask path.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "swap-2-objects";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-08-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/swap-2-objects.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/swap-2-objects.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /**
     * 2つのオブジェクトの中心位置を入れ替える
     * @param {PageItem} objA - 入れ替える一方のオブジェクト
     * @param {PageItem} objB - 入れ替えるもう一方のオブジェクト
     * @returns {void}
     */
    function swapObjectsByCenter(objA, objB) {
        /**
         * オブジェクトの中心座標を返す
         *
         * クリップグループはマスクパスの geometricBounds、それ以外は visibleBounds を基準にする。
         * マスクパスが見つからないクリップグループは visibleBounds にフォールバックする。
         * @param {PageItem} obj - 中心を求める対象のオブジェクト
         * @returns {Array<number>} 中心の [x, y] 座標
         */
        function getCenter(obj) {
            var bounds;
            if (obj.typename === "GroupItem" && obj.clipped) {
                // クリップグループの場合はマスクパスを基準に
                var mask = null;
                for (var i = 0; i < obj.pageItems.length; i++) {
                    if (obj.pageItems[i].clipping) {
                        mask = obj.pageItems[i];
                        break;
                    }
                }
                if (mask) {
                    bounds = mask.geometricBounds; // マスクパスの範囲
                } else {
                    bounds = obj.visibleBounds; // 保険：マスクが見つからなければ全体
                }
            } else {
                bounds = obj.visibleBounds;
            }
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        }

        // 中心座標を取得
        var centerA = getCenter(objA);
        var centerB = getCenter(objB);

        // 移動量を計算
        var deltaA = [centerB[0] - centerA[0], centerB[1] - centerA[1]];
        var deltaB = [centerA[0] - centerB[0], centerA[1] - centerB[1]];

        // 位置を入れ替え
        objA.translate(deltaA[0], deltaA[1]);
        objB.translate(deltaB[0], deltaB[1]);
    }

    /**
     * ドキュメントと選択を検証し、選択した2つのオブジェクトの中心位置を入れ替える
     * @returns {void}
     */
    function main() {
        try {
            if (app.documents.length === 0) {
                alert("ドキュメントが開かれていません。");
                return;
            }

            var sel = app.activeDocument.selection;
            if (!sel || sel.length !== 2) {
                alert("2つのオブジェクトを選択してください。");
                return;
            }

            var objA = sel[0];
            var objB = sel[1];

            swapObjectsByCenter(objA, objB);

        } catch (e) {
            alert("エラーが発生しました: " + e);
        }
    }

    main();

})();
