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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
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
