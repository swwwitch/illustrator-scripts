#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

「_guide」レイヤー内のガイドを通常のパス（塗りなし、線K100・1pt）に戻し、「ReleasedGuides」レイヤーへ移動します。
移動後、「_guide」レイヤーは再ロックします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseGuidesAsPaths.md

### Overview

Turns the guides on the "_guide" layer back into regular paths (no fill, 1pt K100 stroke) and moves them to a "ReleasedGuides" layer.
The "_guide" layer is locked again afterwards.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseGuidesAsPaths.md

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

// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
// ローカライズ（再利用パーツ） / Localization (reusable)
// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

/**
 * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
}

var uiLang = getCurrentLang();

/**
 * LABELS から今の UI 言語の文言を取り出す。
 * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
 * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
 * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
 */
function getLabel(labelRef, placeholderValues) {
    var labelEntry = labelRef;
    if (typeof labelRef === "string") {
        var labelPathKeys = labelRef.split(".");
        labelEntry = LABELS;
        for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
            labelEntry = labelEntry[labelPathKeys[i]];
        }
    }
    var labelString;
    if (typeof labelEntry === "string") labelString = labelEntry;
    else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
    else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
    else return (typeof labelRef === "string") ? labelRef : "";
    return fillLabelPlaceholders(String(labelString), placeholderValues);
}

/**
 * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {Object|Array} [placeholderValues] - getLabel と同じ
 * @returns {string} コロン付きの文言
 */
function labelText(labelRef, placeholderValues) {
    return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
}

/**
 * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
 * @param {string|Object} labelRef - getLabel と同じ
 * @param {string|number} value - コロンのあとに続ける値
 * @returns {string} 項目名と値をつないだ文字列
 */
function labelValueText(labelRef, value) {
    return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
}

/**
 * 文言の {name} や %1 に値を差し込む
 * @param {string} labelString - 文言
 * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
 * @returns {string} 差し込んだ文言
 */
function fillLabelPlaceholders(labelString, placeholderValues) {
    if (placeholderValues == null) return labelString;
    if (placeholderValues instanceof Array) {
        /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
        for (var i = placeholderValues.length; i >= 1; i--) {
            labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
        }
        return labelString;
    }
    for (var placeholderKey in placeholderValues) {
        if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
        labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
    }
    return labelString;
}

// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
// ローカライズ（再利用パーツ）ここまで / End of the reusable localization
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
