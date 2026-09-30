#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のすべてのガイドを削除します。
ロックされたレイヤーも一時的にロックを解除して対象にし、処理後に元のロック状態へ戻します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteAllGuides.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4907511336ad

### Overview

Deletes every guide in the document.
Locked layers are unlocked temporarily so their guides are included, then their lock state is restored.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteAllGuides.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DeleteAllGuides";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteAllGuides.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteAllGuides.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4907511336ad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    // =========================================
    // ガイドの削除 / Deleting guides
    // =========================================

    /**
     * ドキュメント内のガイドをすべて削除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removeAllGuides(doc) {
        var docPathItems = doc.pathItems;
        for (var i = docPathItems.length - 1; i >= 0; i--) {
            if (!docPathItems[i].guides) continue;
            /* ロックされたサブレイヤーやオブジェクトのガイドは削除できないので飛ばす / Guides in locked sublayers or locked guides cannot be removed; skip them */
            try {
                docPathItems[i].remove();
            } catch (e) {}
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * トップレベルのレイヤーのロックを一時的に外して、すべてのガイドを削除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;

        /* トップレベルのレイヤーのロック状態を控えて一時的に解除 / Remember and unlock the top-level layers */
        var topLayers = doc.layers;
        var layerLockStates = [];
        for (var i = 0; i < topLayers.length; i++) {
            layerLockStates[i] = topLayers[i].locked;
            if (layerLockStates[i]) topLayers[i].locked = false;
        }

        /* ガイドのロックを解除 / Unlock guides */
        doc.guidesLocked = false;

        removeAllGuides(doc);

        /* ロック状態を元に戻す / Restore the lock states */
        for (var j = 0; j < topLayers.length; j++) {
            topLayers[j].locked = layerLockStates[j];
        }
    }

    main();

})();
