#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中の2つのテキストオブジェクトの文字列（contents）を入れ替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapText.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n071e09af28a7

### Overview

Swaps the contents of two selected text objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapText";                     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapText.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapText.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n071e09af28a7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo: { ja: "テキストオブジェクトを2つ選択してください。", en: "Select two text objects." },
            notText: {
                ja: "選択した2つは両方ともテキストオブジェクトである必要があります。",
                en: "Both selected objects must be text objects."
            }
        }
    };

    /**
     * ドット区切りのキーから現在の UI 言語のラベルを返す
     * @param {string} labelPath - "alert.noDocument" のようなキー
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
     * 選択中の2つのテキストオブジェクトの文字列を入れ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var selectedItems = app.activeDocument.selection;
        if (selectedItems.length !== 2) {
            alert(getLabel("alert.selectTwo"));
            return;
        }

        var firstTextFrame = selectedItems[0];
        var secondTextFrame = selectedItems[1];
        if (firstTextFrame.typename !== "TextFrame" || secondTextFrame.typename !== "TextFrame") {
            alert(getLabel("alert.notText"));
            return;
        }

        var firstContents = firstTextFrame.contents;
        firstTextFrame.contents = secondTextFrame.contents;
        secondTextFrame.contents = firstContents;
    }

    main();

})();
