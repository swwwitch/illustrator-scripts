#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの塗りと線を、どちらも「なし」にします。
FillStrokeSwitcher の［塗りと線を消去］をダイアログなしで実行するワンクリック版です。

note記事も参照してください。
https://note.com/dtp_tranist/n/n81ee3a9e09b4

参照（しぶやみゃむ さんの記事）：
https://note.com/shibumi/n/n5229b4357dd3

### Overview

Sets both the fill and the stroke of the selected objects to None.
A one-click version of "Erase Fill and Stroke" in FillStrokeSwitcher, without the dialog.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RemoveFillAndStroke";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_ARTICLE_URL   = "https://note.com/dtp_tranist/n/n81ee3a9e09b4"; /* 紹介記事 / article URL */
var SCRIPT_REFERENCE_URL = "https://note.com/shibumi/n/n5229b4357dd3";     /* 参照記事（しぶやみゃむ） / reference article */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
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
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください", en: "Please select at least one object." },
            pathFailures: { ja: "パス処理の失敗", en: "Path processing failures" },
            textFailures: { ja: "テキスト処理の失敗", en: "Text processing failures" },
            details: { ja: "詳細", en: "Details" }
        }
    };

    // =========================================
    // 失敗の集計 / Failure stats
    // =========================================

    /**
     * 処理結果の集計オブジェクトを作る
     * @returns {{pathFailureCount: number, textFailureCount: number, failureDetails: string[]}} 空の集計
     */
    function createProcessStats() {
        return {
            pathFailureCount: 0,
            textFailureCount: 0,
            failureDetails: []
        };
    }

    /**
     * 失敗を数えて詳細を記録する（詳細は最大 8 件）
     * @param {Object} processStats - 集計
     * @param {string} counterKey - 増やすカウンター（"pathFailureCount" など）
     * @param {string} category - 失敗の分類（"Path" など）
     * @param {PageItem} failedItem - 失敗したオブジェクト
     * @param {Error} error - 発生した例外
     * @returns {void}
     */
    function recordFailure(processStats, counterKey, category, failedItem, error) {
        processStats[counterKey]++;
        if (processStats.failureDetails.length >= 8) return;

        var itemTypeName = 'Unknown';
        var itemName = '';
        try {
            if (failedItem && failedItem.typename) itemTypeName = failedItem.typename;
            if (failedItem && failedItem.name) itemName = String(failedItem.name);
        } catch (e) { /* 削除済みのオブジェクトはプロパティを読めない / removed items cannot be read */ }

        var detail = category + ': ' + itemTypeName;
        if (itemName !== '') detail += ' [' + itemName + ']';
        detail += ' - ' + ((error && error.message) ? String(error.message) : String(error));
        processStats.failureDetails.push(detail);
    }

    /**
     * 集計から、失敗を知らせるメッセージを作る
     * @param {Object} processStats - 集計
     * @returns {string} メッセージ（失敗が無ければ空文字）
     */
    function buildFailureMessage(processStats) {
        var messageLines = [];
        if (processStats.pathFailureCount > 0) {
            messageLines.push(labelValueText('alert.pathFailures', processStats.pathFailureCount));
        }
        if (processStats.textFailureCount > 0) {
            messageLines.push(labelValueText('alert.textFailures', processStats.textFailureCount));
        }
        if (processStats.failureDetails.length > 0) {
            messageLines.push('');
            messageLines.push(labelText('alert.details'));
            for (var i = 0; i < processStats.failureDetails.length; i++) {
                messageLines.push('- ' + processStats.failureDetails[i]);
            }
        }
        return messageLines.join('\n');
    }

    // =========================================
    // 塗りと線の消去 / Removing fill and stroke
    // =========================================

    /**
     * パスの塗りと線をなしにする（失敗は集計に記録）
     * @param {PathItem} pathItem - 対象のパス
     * @param {Object} processStats - 集計
     * @returns {void}
     */
    function removePathFillAndStroke(pathItem, processStats) {
        try {
            pathItem.filled = false;
            pathItem.stroked = false;
            pathItem.fillColor = new NoColor();
            pathItem.strokeColor = new NoColor();
        } catch (e) {
            recordFailure(processStats, 'pathFailureCount', 'Path', pathItem, e);
        }
    }

    /**
     * TextRange の塗りと線をなしにする
     * @param {TextRange} textRange - 対象の文字範囲
     * @returns {void}
     */
    function removeTextRangeFillAndStroke(textRange) {
        textRange.characterAttributes.fillColor = new NoColor();
        textRange.characterAttributes.strokeColor = new NoColor();
    }

    /**
     * テキストの各文字の塗りと線をなしにする（失敗は集計に記録）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Object} processStats - 集計
     * @returns {void}
     */
    function removeTextFillAndStroke(textFrame, processStats) {
        try {
            var textCharacters = textFrame.textRange.characters;
            /* 空のテキストは文字が無いので全体に掛ける / Empty text has no characters, so apply to the whole range */
            if (textCharacters.length === 0) {
                removeTextRangeFillAndStroke(textFrame.textRange);
                return;
            }
            for (var i = 0; i < textCharacters.length; i++) {
                removeTextRangeFillAndStroke(textCharacters[i]);
            }
        } catch (e) {
            recordFailure(processStats, 'textFailureCount', 'Text', textFrame, e);
        }
    }

    /**
     * 選択オブジェクトの塗りと線をなしにする（グループ・複合パスは再帰）
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {Object} processStats - 集計
     * @returns {void}
     */
    function processItems(targetItems, processStats) {
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];

            switch (targetItem.typename) {

                case "GroupItem":
                    processItems(targetItem.pageItems, processStats);
                    break;

                case "PathItem":
                    removePathFillAndStroke(targetItem, processStats);
                    break;

                case "CompoundPathItem":
                    processItems(targetItem.pathItems, processStats);
                    break;

                case "TextFrame":
                    removeTextFillAndStroke(targetItem, processStats);
                    break;
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel('alert.noSelection'));
            return;
        }

        /* 処理中に選択が変わらないよう配列に控える / Copy the selection so it stays stable while processing */
        var originalSelection = [];
        for (var i = 0; i < doc.selection.length; i++) {
            originalSelection.push(doc.selection[i]);
        }

        var processStats = createProcessStats();
        processItems(originalSelection, processStats);

        var failureMessage = buildFailureMessage(processStats);
        if (failureMessage !== '') {
            alert(failureMessage);
        }
    }

    main();
})();
