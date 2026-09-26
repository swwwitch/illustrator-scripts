#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、1つずつ個別のグループにします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEachSelection.md

### Overview

Puts each selected object into its own group.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEachSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEachSelection";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEachSelection.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEachSelection.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects are selected." }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    // =========================================
    // 1つずつグループにする / Group each object
    // =========================================

    /**
     * 選択したオブジェクトを、1つずつ別々のグループにする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @param {boolean} skipGroups - グループはグループ化せずに残す
     * @returns {void}
     */
    function groupEachItem(doc, selectedItems, skipGroups) {
        var itemsToSelect = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var targetItem = selectedItems[i];
            if (skipGroups && targetItem.typename === "GroupItem") {
                itemsToSelect.push(targetItem);
                continue;
            }
            var wrapperGroup = targetItem.parent.groupItems.add();
            /* 元の位置のすぐ前面にグループを置いてから中へ入れ、重ね順を保つ / Put the group right in front of the item, then move the item in */
            wrapperGroup.move(targetItem, ElementPlacement.PLACEBEFORE);
            targetItem.move(wrapperGroup, ElementPlacement.PLACEATEND);
            itemsToSelect.push(wrapperGroup);
        }
        doc.selection = itemsToSelect;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したオブジェクトを、1つずつ個別のグループにする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る / A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || selectedItems.typename === "TextRange" || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        /* グループも含めてすべて包む（統合版 GroupMembership の［グループ化済みのものは飛ばす］OFF と同じ）
           Wrap everything, groups included (same as GroupMembership with "skip groups" off) */
        groupEachItem(doc, selectedItems, false);
    }

    main();

})();
