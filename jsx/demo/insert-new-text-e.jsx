#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

現在の表示領域の中心に、欧文フォントを適用したポイントテキストを作成し、選択状態にします。

詳細は README を参照してください。

### Overview

Creates a point text with a Latin font at the center of the current view and leaves it selected.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InsertNewTextE";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewTextE.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewTextE.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n509eb6aa0a19"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var TEXT_CONTENTS = "Typography"; /* 挿入する文言 / text to insert */
    var FONT_SIZE     = 12;           /* フォントサイズ（pt）/ font size (pt) */

    /* フォント候補（上から順に試す）/ Font candidates, tried in order */
    var FONT_CANDIDATES = [
        "BebasNeuePro-Bold",
        "Gilroy-Bold",
        "Avenir-Black",
        "HelveticaNeue-Bold"
    ];

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * 現在のUI言語に合わせた文言を返す
     * @param {object} labelEntry - ja / en を持つラベル定義
     * @returns {string} 表示する文言
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * フォント候補を上から順に試し、最初に見つかったものを適用する
     * @param {TextRange} textRange - 適用先のテキスト範囲
     * @param {string[]} fontCandidates - フォント名の候補
     * @returns {boolean} 適用できたら true
     */
    function applyFallbackFont(textRange, fontCandidates) {
        for (var i = 0; i < fontCandidates.length; i++) {
            try {
                textRange.characterAttributes.textFont = app.textFonts.getByName(fontCandidates[i]);
                return true;
            } catch (e) {
                /* 次の候補を試す / Try the next candidate */
            }
        }
        return false;
    }

    /**
     * 実測バウンズの中心が指定座標に来るように移動する（position はベースライン基準のため）
     * @param {TextFrame} textFrame - 移動するテキストフレーム
     * @param {number[]} center - 合わせ先の中心座標 [x, y]
     * @returns {void}
     */
    function centerOnPoint(textFrame, center) {
        var bounds = textFrame.visibleBounds; /* [left, top, right, bottom] */
        textFrame.translate(
            center[0] - (bounds[0] + bounds[2]) / 2,
            center[1] - (bounds[1] + bounds[3]) / 2
        );
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 表示中心にポイントテキストを作成し、選択状態にする
     * @returns {void}
     */
    function main() {
        /* ドキュメント確認 / Ensure a document is open */
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;

        /* 選択解除 / Clear selection */
        doc.selection = null;

        /* テキストフレーム作成 / Create text frame */
        var textFrame = doc.textFrames.add();
        var textRange = textFrame.textRange;
        textRange.contents = TEXT_CONTENTS;

        /* フォント・サイズ・行揃え / Font, size and justification */
        applyFallbackFont(textRange, FONT_CANDIDATES);
        textRange.characterAttributes.size = FONT_SIZE;
        textRange.paragraphAttributes.justification = Justification.CENTER;

        /* 表示中心へ配置 / Position to view center */
        centerOnPoint(textFrame, doc.activeView.centerPoint);

        /* 作成オブジェクトを選択 / Select created object */
        textFrame.selected = true;
    }

    main();

})();
