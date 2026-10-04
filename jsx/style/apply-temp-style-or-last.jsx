#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトに、既存のグラフィックスタイル「temp_style」を適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/apply-temp-style-or-last.md

### Overview

Applies the existing graphic style named "temp_style" to the selected objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/apply-temp-style-or-last.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "apply-temp-style-or-last";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/apply-temp-style-or-last.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/apply-temp-style-or-last.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion https://qiita.com/comsk/items/87161b2b7d2336b161c4
 * @discussion https://gorolib.blog.jp/archives/73930467.html
 */

    (function () {

        var TARGET_STYLE_NAME = "temp_style";

        /* === ローカライズ / Localization === */

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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください", en: "Please select objects" },
            noStyles: { ja: "グラフィックスタイルが存在しません", en: "No graphic styles in this document" },
            noneApplied: { ja: "スタイルを適用できるオブジェクトがありません", en: "No object accepted the style" }
        };

        /* === コアロジック / Core logic === */
        /* 指定名のスタイルがなければ最後のグラフィックスタイルを返す
           Fall back to the last graphic style if the named one is missing */
        function resolveStyleWithFallback(doc, preferredName) {
            try {
                return doc.graphicStyles.getByName(preferredName);
            } catch (e) { }
            var styles = doc.graphicStyles;
            if (styles.length === 0) return null;
            return styles[styles.length - 1];
        }

        function applyStyleToSelection(doc, style) {
            var currentSelection = doc.selection;
            var appliedCount = 0;
            for (var i = 0; i < currentSelection.length; i++) {
                try {
                    style.applyTo(currentSelection[i]);
                    appliedCount++;
                } catch (e) {
                    // 適用できないアイテムはスキップ / skip items that cannot accept the style
                }
            }
            return appliedCount;
        }

        /* === エントリポイント / Entry point === */
        function main() {
            if (app.documents.length === 0) {
                alert(getLabel("noDocument"));
                return;
            }
            var doc = app.activeDocument;

            if (doc.selection.length === 0) {
                alert(getLabel("noSelection"));
                return;
            }

            var style = resolveStyleWithFallback(doc, TARGET_STYLE_NAME);
            if (!style) {
                alert(getLabel("noStyles"));
                return;
            }

            var appliedCount = applyStyleToSelection(doc, style);
            if (appliedCount === 0) {
                alert(getLabel("noneApplied"));
            }
        }

        main();

    })();
