#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した配置画像と同じリンクファイルを参照している配置画像を探し、指定したファイルへ一括で差し替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkSameImage.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne38eeee5abc8?nt=_3084117

### Overview

Finds the placed images that reference the same linked file as the selected one and relinks them all to a file you choose.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkSameImage.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RelinkSameImage";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-08-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkSameImage.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkSameImage.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne38eeee5abc8?nt=_3084117"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            chooseFile: { ja: "置換するファイルを選択してください", en: "Please select a file to replace" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "アイテムが選択されていません。", en: "No item is selected." },
            selectOneOnly: { ja: "1つだけ選択してください。", en: "Please select only one item." },
            notPlacedItem: { ja: "選択されたアイテムは配置画像ではありません。", en: "The selected item is not a placed image." },
            noFileName: { ja: "リンク画像のファイル名が取得できません。", en: "Could not get the file name of the linked image." },
            cancelSelect: { ja: "ファイルの選択がキャンセルされました。", en: "File selection was cancelled." },
            replacedCount: { ja: " 件のリンク画像を置換しました。", en: " linked image(s) relinked." }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、対象の配置画像を1つ返す（条件に合わなければアラートを出して null）
     * @param {Document} doc - 対象ドキュメント
     * @returns {PlacedItem|null} 選択した配置画像
     */
    function getSelectedPlacedItem(doc) {
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return null;
        }
        if (doc.selection.length > 1) {
            alert(getLabel("alert.selectOneOnly"));
            return null;
        }
        var selectedItem = doc.selection[0];
        if (selectedItem.typename != "PlacedItem") {
            alert(getLabel("alert.notPlacedItem"));
            return null;
        }
        if (!selectedItem.file || !selectedItem.file.name) {
            alert(getLabel("alert.noFileName"));
            return null;
        }
        return selectedItem;
    }

    /**
     * 同じファイル名（大文字・小文字を区別しない）のリンク画像を集める
     * @param {Document} doc - 対象ドキュメント
     * @param {string} fileName - ファイル名
     * @returns {PlacedItem[]} 該当する配置画像
     */
    function collectSameNameItems(doc, fileName) {
        var targetName = fileName.toLowerCase();
        var placedItems = doc.placedItems;
        var matchedItems = [];
        for (var i = 0; i < placedItems.length; i++) {
            if (placedItems[i].file.name.toLowerCase() == targetName) {
                matchedItems.push(placedItems[i]);
            }
        }
        return matchedItems;
    }

    /**
     * 指定したリンクアイテム群を新しいファイルで置換する（すでに同じファイルのものは数えない）
     * @param {PlacedItem[]} items - 対象の配置画像
     * @param {File} newFile - 置換先のファイル
     * @returns {number} 置換した件数
     */
    function replaceLinkedItems(items, newFile) {
        var replacedCount = 0;
        for (var i = 0; i < items.length; i++) {
            if (items[i].file && items[i].file.fsName !== newFile.fsName) {
                items[i].file = newFile;
                replacedCount++;
            }
        }
        return replacedCount;
    }

    /**
     * 選択した配置画像と同名のリンク画像を、選んだファイルへ一括で差し替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        var selectedItem = getSelectedPlacedItem(doc);
        if (!selectedItem) return;

        var matchedItems = collectSameNameItems(doc, selectedItem.file.name);

        var replacementFile = File.openDialog(getLabel("dialog.chooseFile"));
        if (replacementFile == null) {
            alert(getLabel("alert.cancelSelect"));
            return;
        }

        var replacedCount = replaceLinkedItems(matchedItems, replacementFile);
        alert(replacedCount + getLabel("alert.replacedCount"));
        doc.selection = null;
    }

    main();

})();
