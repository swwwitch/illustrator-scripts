#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

最前面のドキュメントの選択範囲の左上を基準に、ほかの開いているドキュメントの選択オブジェクトを同じ座標へ移動します。

詳細は README を参照してください。

### Overview

Moves the selected objects in every other open document to the same position, using the top-left corner of the selection in the frontmost document as the reference.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SyncSelectionPosition";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SyncSelectionPosition.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SyncSelectionPosition.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1f8155daeac4"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // 日英ラベル定義 / Japanese-English labels
    // =========================================

    /**
     * UIロケールに応じた言語コードを返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        alert: {
            needTwoDocuments: { ja: "2つ以上のドキュメントを開いてください。", en: "Please open at least two documents." },
            noSelection:      { ja: "最前面のドキュメントでオブジェクトを選択してください。", en: "Please select objects in the frontmost document." },
            done: {
                ja: "完了しました。\n座標 X: {0}, Y: {1} に統一しました。",
                en: "Done.\nEvery selection was aligned to X: {0}, Y: {1}."
            }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('alert','done')）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    /**
     * ラベル内の {0} {1} … を値で置き換える
     * @param {string} template - プレースホルダーを含む文字列
     * @param {Array} values - 差し込む値
     * @returns {string} 置き換え後の文字列
     */
    function fillPlaceholders(template, values) {
        var filledText = template;
        for (var i = 0; i < values.length; i++) {
            filledText = filledText.replace("{" + i + "}", values[i]);
        }
        return filledText;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* ドキュメントが2つ未満なら処理できない / Need at least two open documents */
    if (app.documents.length < 2) {
        alert(getLabel("alert", "needTwoDocuments"));
        return;
    }

    var sourceDoc = app.activeDocument;
    var sourceItems = sourceDoc.selection;

    /* 基準にする選択がなければ終了 / Exit when the reference selection is empty */
    if (!sourceItems || sourceItems.length === 0) {
        alert(getLabel("alert", "noSelection"));
        return;
    }

    /* 基準は選択範囲全体の左上（Left / Top）/ Reference is the top-left corner of the whole selection */
    var referencePoint = getSelectionTopLeft(sourceItems);

    applyPositionToOtherDocuments(sourceDoc, referencePoint);

    /* 元のドキュメントに戻す / Restore the original active document */
    app.activeDocument = sourceDoc;

    alert(fillPlaceholders(getLabel("alert", "done"), [referencePoint[0].toFixed(2), referencePoint[1].toFixed(2)]));

    /**
     * 選択範囲全体の左上座標を求める
     * Illustratorの座標系ではY軸は上がプラスなので、Topは最大値になる
     * @param {Array} selectedItems - 対象の選択オブジェクト
     * @returns {Array<number>} [x, y] 形式の左上座標
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
     * @param {Array} selectedItems - 移動する選択オブジェクト
     * @param {Array<number>} topLeft - 移動先の左上座標 [x, y]
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
     * @param {Array<number>} topLeft - 揃える左上座標 [x, y]
     * @returns {void}
     */
    function applyPositionToOtherDocuments(referenceDoc, topLeft) {
        for (var i = 0; i < app.documents.length; i++) {
            var otherDoc = app.documents[i];
            if (otherDoc === referenceDoc) continue;

            /* selection を読むにはアクティブにする必要がある / The document must be active to read its selection */
            app.activeDocument = otherDoc;

            var otherItems = otherDoc.selection;
            if (otherItems && otherItems.length > 0) {
                moveSelectionTopLeftTo(otherItems, topLeft);
            }
        }
    }

})();
