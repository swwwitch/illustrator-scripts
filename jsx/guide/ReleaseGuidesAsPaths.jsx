#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

「_guide」レイヤー内のガイドを通常のパス（塗りなし、線K100・1pt）に戻し、「ReleasedGuides」レイヤーへ移動します。
移動後、「_guide」レイヤーは再ロックします。

詳細は README を参照してください。

### Overview

Turns the guides on the "_guide" layer back into regular paths (no fill, 1pt K100 stroke) and moves them to a "ReleasedGuides" layer.
The "_guide" layer is locked again afterwards.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReleaseGuidesAsPaths";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseGuidesAsPaths.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseGuidesAsPaths.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ラベル定義 / Labels
// =========================================
var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noGuideLayer: { ja: "「_guide」レイヤーが見つかりません。", en: "No \"_guide\" layer found." }
    }
};

var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

/**
 * ラベル定義から現在の言語の文言を取得する
 * @param {Object} entry - ja / en を持つラベル定義
 * @returns {string} 現在の言語の文言
 */
function getLabel(entry) {
    return entry[uiLang] || entry.en;
}

(function () {

    var GUIDE_LAYER_NAME           = "_guide";         /* ガイドを収めたレイヤー / layer holding the guides */
    var RELEASED_LAYER_NAME        = "ReleasedGuides"; /* 解除したガイドの移動先 / destination for released guides */
    var LEGACY_RELEASED_LAYER_NAME = "UnlockedGuides"; /* 以前の移動先（既存ドキュメント用） / former destination kept for existing documents */

    /**
     * 最上位レイヤーから名前が一致する最初のレイヤーを探す
     * @param {Document} targetDoc - 検索するドキュメント
     * @param {string} layerName - 探すレイヤー名
     * @returns {Layer|null} 見つかったレイヤー。なければ null
     */
    function findTopLevelLayer(targetDoc, layerName) {
        for (var i = 0; i < targetDoc.layers.length; i++) {
            if (targetDoc.layers[i].name == layerName) {
                return targetDoc.layers[i];
            }
        }
        return null;
    }

    /**
     * 移動先レイヤー（なければ以前の名前のレイヤー）を取得してロックを解除する。どちらもなければ作成する
     * @param {Document} targetDoc - 対象のドキュメント
     * @returns {Layer} 移動先レイヤー
     */
    function prepareReleasedLayer(targetDoc) {
        var releasedLayer = findTopLevelLayer(targetDoc, RELEASED_LAYER_NAME) || findTopLevelLayer(targetDoc, LEGACY_RELEASED_LAYER_NAME);
        if (releasedLayer) {
            releasedLayer.locked = false; /* ロックを解除 / Unlock the layer */
            return releasedLayer;
        }
        releasedLayer = targetDoc.layers.add();
        releasedLayer.name = RELEASED_LAYER_NAME;
        return releasedLayer;
    }

    /**
     * ガイドを通常のパスに戻し、塗りなし・線1ptにして移動先レイヤーへ移す
     * @param {PageItem} guideItem - 「_guide」レイヤー内のアイテム
     * @param {GrayColor} k100Color - 線に設定するK100のカラー
     * @param {Layer} releasedLayer - 移動先レイヤー
     * @returns {void}
     */
    function releaseGuideItem(guideItem, k100Color, releasedLayer) {
        if (guideItem.guides) {
            guideItem.guides = false; /* ガイドを解除 / Remove guide flag */
        }
        /* 外観設定：塗りなし、線はK100、1pt / Set appearance: no fill, stroke K100, 1pt */
        guideItem.filled = false;
        guideItem.stroked = true;
        guideItem.strokeColor = k100Color;
        guideItem.strokeWidth = 1;
        guideItem.move(releasedLayer, ElementPlacement.PLACEATBEGINNING);
    }

    /**
     * 「_guide」レイヤーのガイドを解除して「ReleasedGuides」レイヤーへ移し、「_guide」レイヤーを再ロックする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var activeDoc = app.activeDocument;
        var guideLayer = findTopLevelLayer(activeDoc, GUIDE_LAYER_NAME);
        if (!guideLayer) {
            alert(getLabel(LABELS.alert.noGuideLayer));
            return;
        }

        guideLayer.locked = false; /* ロックを解除 / Unlock the layer */
        var releasedLayer = prepareReleasedLayer(activeDoc);

        /* 取得したカラーの書き換えは反映されないため、作ってから代入する / Build the color first; editing the getter's copy has no effect */
        var k100Color = new GrayColor();
        k100Color.gray = 100;

        var guideLayerItems = guideLayer.pageItems;
        for (var i = guideLayerItems.length - 1; i >= 0; i--) {
            releaseGuideItem(guideLayerItems[i], k100Color, releasedLayer);
        }

        guideLayer.locked = true; /* 再ロック / Relock the layer */
    }

    main();

})();
