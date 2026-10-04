#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

最前面のドキュメントの選択範囲の左上を基準に、ほかの開いているドキュメントの選択オブジェクトを同じ座標へ移動します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SyncSelectionPosition.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1f8155daeac4

### Overview

Moves the selected objects in every other open document to the same position, using the top-left corner of the selection in the frontmost document as the reference.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SyncSelectionPosition.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SyncSelectionPosition";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SyncSelectionPosition.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SyncSelectionPosition.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1f8155daeac4"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        alert: {
            needTwoDocuments: { ja: "2つ以上のドキュメントを開いてください。", en: "Please open at least two documents." },
            noSelection: {
                ja: "最前面のドキュメントでオブジェクトを選択してください。",
                en: "Please select objects in the frontmost document."
            },
            done: {
                ja: "完了しました。\n座標 X: %1, Y: %2 に統一しました。",
                en: "Done.\nEvery selection was aligned to X: %1, Y: %2."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 最前面のドキュメントの選択範囲の左上に、ほかのドキュメントの選択オブジェクトをそろえる
     * @returns {void}
     */
    function main() {
        /* ドキュメントが2つ未満なら処理できない / Need at least two open documents */
        if (app.documents.length < 2) {
            alert(getLabel("alert.needTwoDocuments"));
            return;
        }

        var sourceDoc = app.activeDocument;
        var sourceItems = sourceDoc.selection;

        /* 基準にする選択がなければ終了 / Exit when the reference selection is empty */
        if (!sourceItems || sourceItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 基準は選択範囲全体の左上（Left / Top）/ Reference is the top-left corner of the whole selection */
        var referencePoint = getSelectionTopLeft(sourceItems);

        applyPositionToOtherDocuments(sourceDoc, referencePoint);

        /* 元のドキュメントに戻す / Restore the original active document */
        app.activeDocument = sourceDoc;

        alert(getLabel("alert.done", [referencePoint[0].toFixed(2), referencePoint[1].toFixed(2)]));
    }

    // =========================================
    // 位置合わせ / Positioning
    // =========================================

    /**
     * 選択範囲全体の左上座標を求める
     * Illustratorの座標系ではY軸は上がプラスなので、Topは最大値になる
     * @param {PageItem[]} selectedItems - 対象の選択オブジェクト
     * @returns {number[]} [x, y] 形式の左上座標
     */
    function getSelectionTopLeft(selectedItems) {
        var leftMost = selectedItems[0].position[0];
        var topMost = selectedItems[0].position[1];

        for (var i = 1; i < selectedItems.length; i++) {
            var itemLeft = selectedItems[i].position[0];
            var itemTop = selectedItems[i].position[1];
            if (itemLeft < leftMost) leftMost = itemLeft;
            if (itemTop > topMost) topMost = itemTop;
        }
        return [leftMost, topMost];
    }

    /**
     * 選択範囲全体の左上が指定座標に来るように移動する
     * @param {PageItem[]} selectedItems - 移動する選択オブジェクト
     * @param {number[]} topLeft - 移動先の左上座標 [x, y]
     * @returns {void}
     */
    function moveSelectionTopLeftTo(selectedItems, topLeft) {
        var currentTopLeft = getSelectionTopLeft(selectedItems);
        var dx = topLeft[0] - currentTopLeft[0];
        var dy = topLeft[1] - currentTopLeft[1];

        for (var i = 0; i < selectedItems.length; i++) {
            selectedItems[i].translate(dx, dy);
        }
    }

    /**
     * 基準ドキュメント以外の開いているドキュメントで、選択オブジェクトを基準座標に揃える
     * 選択がないドキュメントは何もせずスキップする
     * @param {Document} referenceDoc - 基準にするドキュメント（処理対象から除外）
     * @param {number[]} topLeft - 揃える左上座標 [x, y]
     * @returns {void}
     */
    function applyPositionToOtherDocuments(referenceDoc, topLeft) {
        for (var i = 0; i < app.documents.length; i++) {
            var otherDoc = app.documents[i];
            if (otherDoc === referenceDoc) continue;

            /* selection を読むにはアクティブにする必要がある / The document must be active to read its selection */
            app.activeDocument = otherDoc;

            var targetItems = otherDoc.selection;
            if (targetItems && targetItems.length > 0) {
                moveSelectionTopLeftTo(targetItems, topLeft);
            }
        }
    }

    main();

})();
