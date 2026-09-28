#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択している段落で使われている禁則処理の値を列挙して表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InspectKinsoku.md

### Overview

Lists the kinsoku (line-breaking) settings used by the selected paragraphs.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InspectKinsoku.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InspectKinsoku";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InspectKinsoku.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InspectKinsoku.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
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

    var LABELS = {
        alert: {
            noParagraph: {
                ja: "段落が選択されていません。\nテキスト、またはテキストフレームを選択してください。",
                en: "No paragraph is selected.\nSelect text or a text frame."
            },
            reportTitle: { ja: "検出された禁則の値", en: "Kinsoku values found" }
        },
        fallbackName: {
            kinsokuNone: { ja: "なし", en: "None" }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * コレクションの段落を配列に追加する
     * @param {TextRange[]} targetParagraphs - 追加先の配列
     * @param {Paragraphs} paragraphCollection - 段落のコレクション
     * @returns {void}
     */
    function appendParagraphs(targetParagraphs, paragraphCollection) {
        for (var i = 0; i < paragraphCollection.length; i++) {
            targetParagraphs.push(paragraphCollection[i]);
        }
    }

    /**
     * 選択中のテキスト（またはテキストフレーム）の段落を集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextRange[]} 段落の配列
     */
    function collectSelectedParagraphs(doc) {
        var currentSelection = doc.selection;
        var targetParagraphs = [];

        if (!currentSelection) {
            return targetParagraphs;
        }

        /* 文字ツールでテキストを選択した場合 / Text selected with the Type tool */
        if (currentSelection.typename === "TextRange") {
            appendParagraphs(targetParagraphs, currentSelection.paragraphs);
            return targetParagraphs;
        }

        /* 選択ツールでオブジェクトを選択した場合（配列） / Objects selected with the Selection tool (array) */
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem.typename === "TextFrame") {
                appendParagraphs(targetParagraphs, selectedItem.textRange.paragraphs);
            }
        }
        return targetParagraphs;
    }

    /**
     * 段落を走査し、検出した禁則の値を集合として返す
     * @param {TextRange[]} paragraphs - 調べる段落
     * @returns {Object} 禁則の値をキーにした集合
     */
    function collectKinsokuValues(paragraphs) {
        var detectedKinsokuSet = {};

        for (var i = 0; i < paragraphs.length; i++) {
            try {
                /* 禁則「なし」の段落は kinsoku を読むだけで Error 9563 を投げるため try は必須
                   Reading kinsoku on a "None" paragraph throws Error 9563 */
                var kinsokuValue = paragraphs[i].paragraphAttributes.kinsoku;
                detectedKinsokuSet[String(kinsokuValue)] = true;
            } catch (e) {
                if (e.number === 9563) {
                    detectedKinsokuSet[getLabel("fallbackName.kinsokuNone")] = true;
                } else {
                    detectedKinsokuSet["ERROR: " + e.message] = true;
                }
            }
        }

        return detectedKinsokuSet;
    }

    /**
     * 選択段落の禁則の値を集めてアラートで一覧表示する
     * @returns {void}
     */
    function main() {
        var targetParagraphs = collectSelectedParagraphs(app.activeDocument);
        if (!targetParagraphs.length) {
            alert(getLabel("alert.noParagraph"));
            return;
        }

        var detectedKinsokuSet = collectKinsokuValues(targetParagraphs);

        var reportText = labelText("alert.reportTitle") + "\n";
        for (var detectedValue in detectedKinsokuSet) {
            reportText += "  → \"" + detectedValue + "\"\n";
        }
        alert(reportText);
    }

    main();

})();
