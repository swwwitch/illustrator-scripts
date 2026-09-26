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
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

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

    /**
     * ドット区切りのキーから現在の UI 言語のラベルを返す
     * @param {string} labelPath - "alert.done" のようなキー
     * @returns {string} 現在の UI 言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

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
