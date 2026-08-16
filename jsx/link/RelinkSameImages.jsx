#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した配置画像と同じリンクファイルを参照している配置画像をドキュメント全体から探し、
指定したファイルへ一括で差し替えます。グループの中にある配置画像も自動で解決します。

詳細は README を参照してください。

*/

/*

### Overview

Finds every placed image in the active document that references the same linked file as the
selection, then relinks them all to a file you choose. A placed image nested inside a group
is resolved automatically.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RelinkSameImages";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkSameImages.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkSameImages.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne38eeee5abc8"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIの表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            selectReplaceFile: { ja: "置換するファイルを選択してください", en: "Select a file to replace with" }
        },
        /* メッセージ / Messages */
        alert: {
            canceled: { ja: "ファイルの選択がキャンセルされました。", en: "File selection was canceled." },
            notPlacedItem: { ja: "選択されたアイテムは配置画像ではありません。", en: "The selected item is not a placed image." },
            groupHasNoPlacedItem: { ja: "選択されたグループ内に配置画像が見つかりません。", en: "No placed image was found inside the selected group." },
            nothingSelected: { ja: "アイテムが選択されていません。", en: "No item is selected." },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noLinkedFile: {
                ja: "選択した配置画像のリンク情報を取得できません。埋め込み画像またはリンク切れの可能性があります。",
                en: "Could not retrieve the link information for the selected placed image. It may be embedded or missing."
            },
            replaced: { ja: "差し替え完了: #count#件", en: "Replacement complete: #count# item(s)" }
        }
    };

    /**
     * "category.key" 形式のラベルを現在の言語で取得する
     * @param {string} key - ラベルキー（例: "alert.noDocument"）
     * @returns {string} 現在の言語のラベル文字列（未定義のときはキーをそのまま返す）
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[keyParts[i]];
        }
        if (labelNode) {
            if (typeof labelNode[currentLanguage] === "string") return labelNode[currentLanguage];
            if (typeof labelNode.en === "string") return labelNode.en;
        }
        return key;
    }

    /**
     * 件数を埋め込んだメッセージを組み立てる
     * @param {string} key - ラベルキー（`#count#` を含むもの）
     * @param {number} count - 埋め込む件数
     * @returns {string} 件数を差し替えた文字列
     */
    function formatCount(key, count) {
        return getLabel(key).replace("#count#", String(count));
    }

    // =========================================
    // 配置画像の解決・収集 / Resolve & collect placed items
    // =========================================

    /**
     * 配置画像のリンクファイルを取得する
     * @param {PlacedItem} placedItem - 対象の配置画像
     * @returns {File|null} リンクファイル。埋め込み画像・リンク切れのときは null
     */
    function getLinkedFile(placedItem) {
        try {
            /* 埋め込み画像やリンク切れでは file の参照で例外になる / Throws for embedded images and missing links */
            return placedItem.file;
        } catch (e) {
            return null;
        }
    }

    /**
     * ページアイテムから配置画像を再帰的に解決する（グループ内も掘り下げる）
     * @param {PageItem} pageItem - 探索の起点となるページアイテム
     * @returns {PlacedItem|null} 最初に見つかった配置画像。無ければ null
     */
    function resolvePlacedItem(pageItem) {
        if (!pageItem) return null;

        if (pageItem.typename === "PlacedItem") return pageItem;

        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                var resolved = resolvePlacedItem(pageItem.pageItems[i]);
                if (resolved) return resolved;
            }
        }

        return null;
    }

    /**
     * 選択の先頭から基準となる配置画像を取得する
     * @param {Document} doc - 対象ドキュメント
     * @returns {PlacedItem|null} 基準の配置画像。取得できないときは警告して null
     */
    function getReferencePlacedItem(doc) {
        if (!doc.selection.length) {
            alert(getLabel("alert.nothingSelected"));
            return null;
        }

        var selectedItem = doc.selection[0];
        var placedItem = resolvePlacedItem(selectedItem);
        if (placedItem) return placedItem;

        /* グループを選んでいたのか、配置画像以外を選んでいたのかを分けて伝える / Tell the two failure cases apart */
        alert(selectedItem.typename === "GroupItem"
            ? getLabel("alert.groupHasNoPlacedItem")
            : getLabel("alert.notPlacedItem"));
        return null;
    }

    /**
     * 基準と同じリンクファイルを参照する配置画像をドキュメント全体から集める
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacedItem} referenceItem - 基準となる配置画像
     * @returns {Array<PlacedItem>|null} 一致した配置画像。基準のリンクが取得できないときは null
     */
    function collectMatchedPlacedItems(doc, referenceItem) {
        var referenceFile = getLinkedFile(referenceItem);
        if (!referenceFile) {
            alert(getLabel("alert.noLinkedFile"));
            return null;
        }

        /* 判定はファイル名ではなく絶対パスで行う / Compare absolute paths, not file names */
        var referenceFsName = referenceFile.fsName;
        var matchedItems = [];
        var placedItems = doc.placedItems;
        for (var i = 0; i < placedItems.length; i++) {
            var linkedFile = getLinkedFile(placedItems[i]);
            if (linkedFile && linkedFile.fsName === referenceFsName) matchedItems.push(placedItems[i]);
        }
        return matchedItems;
    }

    // =========================================
    // ファイル選択・差し替え / Choose file & replace
    // =========================================

    /**
     * 差し替え先のファイルをダイアログで選択する
     * @returns {File|null} 選択したファイル。キャンセル時は警告して null
     */
    function chooseReplacementFile() {
        var replacementFile = File.openDialog(getLabel("dialog.selectReplaceFile"));
        if (!replacementFile) {
            alert(getLabel("alert.canceled"));
            return null;
        }
        return replacementFile;
    }

    /**
     * 対象の配置画像を指定したファイルへ差し替える
     * @param {Array<PlacedItem>} placedItems - 差し替える配置画像
     * @param {File} replacementFile - 差し替え先のファイル
     * @returns {number} 差し替えた件数
     */
    function replacePlacedItems(placedItems, replacementFile) {
        for (var i = 0; i < placedItems.length; i++) {
            placedItems[i].file = replacementFile;
        }
        return placedItems.length;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 基準画像と同じリンクを参照する配置画像を、選んだファイルへ一括で差し替える
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        var referenceItem = getReferencePlacedItem(doc);
        if (!referenceItem) return;

        var targetItems = collectMatchedPlacedItems(doc, referenceItem);
        if (!targetItems) return;

        var replacementFile = chooseReplacementFile();
        if (!replacementFile) return;

        var replacedCount = replacePlacedItems(targetItems, replacementFile);

        /* 差し替え済みの選択を残さない / Do not leave the replaced items selected */
        doc.selection = null;

        alert(formatCount("alert.replaced", replacedCount));
    }

    main();

})();
