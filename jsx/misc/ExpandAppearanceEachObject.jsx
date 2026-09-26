#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択しているオブジェクトごとに、アピアランスを分割します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandAppearanceEachObject.md

### Overview

Expands the appearance of each selected object individually.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandAppearanceEachObject.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExpandAppearanceEachObject";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandAppearanceEachObject.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandAppearanceEachObject.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文字列を引く
     * @param {string} labelPath - "alert.noSelection" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択しているオブジェクトを1つずつ選び直し、アピアランスを分割する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        if (selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        for (var i = selectedItems.length - 1; i >= 0; i--) {
            doc.selection = null;
            selectedItems[i].selected = true;
            app.executeMenuCommand('expandStyle');
        }
    }

    main();

})();
