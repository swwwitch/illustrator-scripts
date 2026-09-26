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

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "alert.noDocument" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

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
