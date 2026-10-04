#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

現在の表示領域の中心に、欧文フォントを適用したポイントテキストを作成し、選択状態にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewTextE.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n509eb6aa0a19

### Overview

Creates a point text with a Latin font at the center of the current view and leaves it selected.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewTextE.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InsertNewTextE";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

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
     * @param {number[]} targetCenter - 合わせ先の中心座標 [x, y]
     * @returns {void}
     */
    function centerOnPoint(textFrame, targetCenter) {
        var frameBounds = textFrame.visibleBounds; /* [left, top, right, bottom] */
        textFrame.translate(
            targetCenter[0] - (frameBounds[0] + frameBounds[2]) / 2,
            targetCenter[1] - (frameBounds[1] + frameBounds[3]) / 2
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
