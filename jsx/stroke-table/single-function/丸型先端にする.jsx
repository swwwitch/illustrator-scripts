#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパスアイテム（グループ内も含む）の線端を、丸型線端に設定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/丸型先端にする.md

### Overview

Sets the stroke cap of the selected path items, including those inside groups, to a round cap.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/丸型先端にする.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "丸型先端にする";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-08-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/丸型先端にする.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/丸型先端にする.md"; /* README (English) */

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
            noDocument: {
                ja: "ドキュメントが開かれていません。ドキュメントを開いてください。",
                en: "No document is open. Please open a document."
            },
            noSelection: {
                ja: "パスアイテムが選択されていません。パスアイテムを選択してください。",
                en: "No path items are selected. Please select path items."
            }
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
     * パスアイテムに丸型線端を適用する（グループ内は再帰的にたどる）
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {void}
     */
    function applyRoundCap(item) {
        if (item.typename === "GroupItem") {
            /* グループの場合、子アイテムを再帰的に処理 / Recurse into group members */
            for (var i = 0; i < item.pageItems.length; i++) {
                applyRoundCap(item.pageItems[i]);
            }
        } else if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
            if (item.stroked && item.strokeCap !== StrokeCap.ROUNDENDCAP) {
                item.strokeCap = StrokeCap.ROUNDENDCAP;
            }
        }
    }

    /**
     * ドキュメントと選択を確認し、選択中のオブジェクトに丸型線端を適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var selectedItems = app.activeDocument.selection;
        if (selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        for (var i = 0; i < selectedItems.length; i++) {
            applyRoundCap(selectedItems[i]);
        }
    }

    main();

})();
