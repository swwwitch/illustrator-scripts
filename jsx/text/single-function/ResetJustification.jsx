#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの段落に対して、ジャスティフィケーション関連の設定（ワードスペース、文字間、グリフスケーリング）を初期値に戻します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetJustification.md

### Overview

Resets the justification settings — word spacing, letter spacing and glyph scaling — of the selected paragraphs to their defaults.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetJustification.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResetJustification";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetJustification.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetJustification.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 戻す値（段落パネル［ジャスティフィケーション］の初期値）/ Values to restore (Justification defaults) */
    var JUSTIFICATION_DEFAULTS = {
        /* ワードスペース（%）/ Word spacing (%) */
        minimumWordSpacing: 80,
        desiredWordSpacing: 100,
        maximumWordSpacing: 133,
        /* 文字間（%）/ Letter spacing (%) */
        minimumLetterSpacing: 0,
        desiredLetterSpacing: 0,
        maximumLetterSpacing: 0,
        /* グリフスケーリング（%）/ Glyph scaling (%) */
        minimumGlyphScaling: 100,
        desiredGlyphScaling: 100,
        maximumGlyphScaling: 100
    };

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

    var LABELS = {
        alert: {
            selectText: { ja: "テキストを選択してください。", en: "Select text." },
            noText: { ja: "テキストが選択されていません。", en: "No text is selected." },
            done: {
                ja: "ジャスティフィケーションの設定を初期値に戻しました。",
                en: "Justification settings have been reset to their defaults."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * テキスト範囲の各段落にジャスティフィケーションの初期値を書き込む
     * @param {TextRange} targetTextRange - 対象のテキスト範囲
     * @returns {void}
     */
    function resetJustification(targetTextRange) {
        var paragraphs = targetTextRange.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraphAttributes = paragraphs[i].paragraphAttributes;
            for (var attributeName in JUSTIFICATION_DEFAULTS) {
                paragraphAttributes[attributeName] = JUSTIFICATION_DEFAULTS[attributeName];
            }
        }
    }

    /**
     * 選択の先頭からテキスト範囲を取り出す
     * @param {Object} firstItem - 選択の先頭の要素
     * @returns {TextRange|null} テキスト範囲。テキストでなければ null
     */
    function getTargetTextRange(firstItem) {
        if (firstItem.typename === "TextFrame") return firstItem.textRange;
        if (firstItem.story) return firstItem;
        return null;
    }

    /**
     * 選択中のテキストのジャスティフィケーションを初期値に戻す
     * @returns {void}
     */
    function main() {
        var selectedItems = app.activeDocument.selection;
        if (selectedItems.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var targetTextRange = getTargetTextRange(selectedItems[0]);
        if (targetTextRange === null) {
            alert(getLabel("alert.noText"));
            return;
        }

        resetJustification(targetTextRange);
        alert(getLabel("alert.done"));
    }

    main();

})();
