#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトそれぞれに「共通 > アピアランス」を適用し、見つかったオブジェクトをまとめて選択し直します。標準機能では基準にできるオブジェクトが1つだけという制限を回避できます。

詳細は README を参照してください。

### Overview

Applies "Select > Same > Appearance" to each selected object and reselects every object found. This works around the built-in limitation of using only one object as the reference.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectSameAppearanceMulti";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-06";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameAppearanceMulti.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameAppearanceMulti.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /**
     * UIの表示言語を返す
     * @return {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* UIラベル定義 / UI label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select at least one object." }
        }
    };

    /**
     * ラベル定義から現在の表示言語の文字列を取り出す
     * @param {object} labelSet - ja / en を持つラベル定義
     * @return {string} 表示用の文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang];
    }

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    /* doc.selection は参照するたびに新しい配列を返すため、そのまま基準リストとして保持できる */
    var referencePageItems = doc.selection;

    if (!referencePageItems || referencePageItems.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    /**
     * 指定したページアイテムを基準に「共通 > アピアランス」を実行し、選択されたページアイテムを返す
     * @param {PageItem} referencePageItem - 基準にするページアイテム
     * @return {array} アピアランスが一致したページアイテムの配列
     */
    function findPageItemsWithSameAppearance(referencePageItem) {
        doc.selection = null;
        referencePageItem.selected = true;
        /* 選択状態を確定させてからメニューコマンドを実行する */
        app.redraw();
        app.executeMenuCommand("Find Appearance menu item");
        return doc.selection;
    }

    /* 重複したまま集めてよい。選択し直す際に同じページアイテムを複数回指定しても結果は変わらない */
    var pageItemsToSelect = [];
    for (var i = 0; i < referencePageItems.length; i++) {
        var samePageItems = findPageItemsWithSameAppearance(referencePageItems[i]);
        for (var j = 0; j < samePageItems.length; j++) {
            pageItemsToSelect.push(samePageItems[j]);
        }
    }

    doc.selection = null;
    for (var k = 0; k < pageItemsToSelect.length; k++) {
        pageItemsToSelect[k].selected = true;
    }
    app.redraw();
})();
