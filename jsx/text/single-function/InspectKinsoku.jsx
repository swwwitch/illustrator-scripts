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
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InspectKinsoku.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InspectKinsoku.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /**
     * コレクションの段落を配列に追加する
     * @param {Object[]} targetParagraphs - 追加先の配列
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
        if (currentSelection.constructor && currentSelection.constructor.name === "TextRange") {
            appendParagraphs(targetParagraphs, currentSelection.paragraphs);
            return targetParagraphs;
        }

        /* 選択ツールでオブジェクトを選択した場合（配列） / Objects selected with the Selection tool (array) */
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem.constructor && selectedItem.constructor.name === "TextFrame") {
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
                    detectedKinsokuSet["なし"] = true;
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
            alert("段落が選択されていません。\nテキスト、またはテキストフレームを選択してください。");
            return;
        }

        var detectedKinsokuSet = collectKinsokuValues(targetParagraphs);

        var reportText = "検出された禁則値:\n";
        for (var detectedValue in detectedKinsokuSet) {
            reportText += "  → \"" + detectedValue + "\"\n";
        }
        alert(reportText);
    }

    main();

})();
