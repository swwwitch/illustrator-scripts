#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した配置画像と同じリンクファイルを参照している配置画像をドキュメント全体から探し、指定したファイルへ一括で差し替えます。
グループの中にある配置画像も自動で解決します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkSameImages.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne38eeee5abc8

### Overview

Searches the whole document for placed images that reference the same linked file as the selection and relinks them all to a file you choose.
Placed images nested inside groups are resolved automatically.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkSameImages.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RelinkSameImages";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-17";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkSameImages.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkSameImages.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne38eeee5abc8"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
            replaced: { ja: "差し替え完了: {count}件", en: "Replacement complete: {count} item(s)" }
        }
    };

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

        alert(getLabel("alert.replaced", { count: replacedCount }));
    }

    main();

})();
