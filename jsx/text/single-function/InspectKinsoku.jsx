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
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

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

    /**
     * ドット区切りのキーから現在の UI 言語のラベルを返す
     * @param {string} labelPath - "alert.noParagraph" のようなキー
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

    /**
     * ラベルに言語別のコロンを付けて返す（日本語は全角、英語は半角）
     * @param {string} labelPath - LABELS のキー
     * @returns {string} コロン付きのラベル
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

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
