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
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/swap-2-objects.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/swap-2-objects.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo: { ja: "2つのオブジェクトを選択してください。", en: "Select two objects." }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    // =========================================
    // 入れ替え / Swapping
    // =========================================

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

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を検証し、選択した2つのオブジェクトの中心位置を入れ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var selectedObjects = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は TextRange が返り、length は文字数になる / A text selection returns a TextRange whose length counts characters */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length !== 2) {
            alert(getLabel(LABELS.alert.selectTwo));
            return;
        }

        swapObjectsByCenter(selectedObjects[0], selectedObjects[1]);
    }

    main();

})();
