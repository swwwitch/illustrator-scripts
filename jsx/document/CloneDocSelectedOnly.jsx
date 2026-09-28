#target indesign
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトだけを残した複製ドキュメントを作成するInDesign用スクリプトです。
元ドキュメントを一時保存してから複製を開き、非選択・非表示のオブジェクトを削除します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloneDocSelectedOnly.md

### Overview

An InDesign script that creates a duplicate document containing only the selected objects.
The original is saved to a temporary file, the copy is opened, and unselected and hidden objects are removed from it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloneDocSelectedOnly.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CloneDocSelectedOnly";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-12-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloneDocSelectedOnly.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloneDocSelectedOnly.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var REMOVE_LOCKED_ITEMS = false; /* true: ロックされたオブジェクトやレイヤーも削除 / Remove locked items and layers if true */

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

    var LABELS = {
        alert: {
            noDocument: { ja: "開いているドキュメントがありません。", en: "No documents are open." },
            notSaved: {
                ja: "ドキュメントが一度も保存されていません。先に保存してください。",
                en: "The document has never been saved. Please save it first."
            },
            noSelection: { ja: "選択されているオブジェクトがありません。", en: "No objects are selected." }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 元ドキュメントを一時ファイルに保存して開き、選択外の非表示オブジェクトを削除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var originalDoc = app.activeDocument;
        /* 一度でも保存されているか / Check if the document has been saved at least once */
        if (!originalDoc.saved) {
            alert(getLabel(LABELS.alert.notSaved));
            return;
        }

        var selectedItems = getSelectedItems(originalDoc);
        if (selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var tempFile = getUniqueTempFile(originalDoc.fullName, originalDoc.name);
        originalDoc.saveAs(tempFile);
        var duplicateDoc = app.open(tempFile);

        removeUnselectedHiddenItems(duplicateDoc.layers, selectedItems);
    }

    // =========================================
    // 選択と削除 / Selection and removal
    // =========================================

    /**
     * 選択されているオブジェクトを配列で返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択オブジェクトの配列
     */
    function getSelectedItems(doc) {
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }
        return selectedItems;
    }

    /**
     * 選択されていない非表示のオブジェクトを削除する（ロック中のものは設定に応じて削除）
     * @param {Layers} layers - 対象ドキュメントのレイヤー
     * @param {PageItem[]} selectedItems - 残す選択オブジェクト
     * @returns {void}
     */
    function removeUnselectedHiddenItems(layers, selectedItems) {
        for (var i = layers.length - 1; i >= 0; i--) {
            var currentLayer = layers[i];

            if (currentLayer.locked) {
                if (REMOVE_LOCKED_ITEMS) {
                    currentLayer.locked = false;
                    currentLayer.remove();
                }
                continue;
            }
            if (!currentLayer.visible) continue;

            for (var j = currentLayer.pageItems.length - 1; j >= 0; j--) {
                var pageItem = currentLayer.pageItems[j];
                if (pageItem.locked) {
                    if (REMOVE_LOCKED_ITEMS) {
                        pageItem.locked = false;
                        pageItem.remove();
                    }
                    continue;
                }
                if (!isItemSelected(pageItem, selectedItems) && !pageItem.visible) {
                    pageItem.remove();
                }
            }
        }
    }

    /**
     * オブジェクトが選択オブジェクトの中にあるか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @param {PageItem[]} selectedItems - 選択オブジェクトの配列
     * @returns {boolean} 含まれていれば true
     */
    function isItemSelected(pageItem, selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (pageItem === selectedItems[i]) return true;
        }
        return false;
    }

    // =========================================
    // 一時ファイル / Temporary file
    // =========================================

    /**
     * ファイル名の最初のドットより前を返す
     * @param {string} fileName - ファイル名
     * @returns {string} ベース名
     */
    function getBaseName(fileName) {
        return fileName.split('.')[0];
    }

    /**
     * ファイル名の拡張子をドット付きで返す（無ければ空文字）
     * @param {string} fileName - ファイル名
     * @returns {string} 拡張子
     */
    function getExtension(fileName) {
        var nameParts = fileName.split('.');
        return nameParts.length > 1 ? '.' + nameParts[nameParts.length - 1] : '';
    }

    /**
     * 元ファイルと同じフォルダーに、既存と重ならない「temp-<名前>」のファイルを返す
     * @param {File} originalFile - 元ドキュメントのファイル
     * @param {string} fileName - 元ドキュメントのファイル名
     * @returns {File} 一時ファイル
     */
    function getUniqueTempFile(originalFile, fileName) {
        var folderPath = originalFile.path + "/";
        var tempFileNameBase = "temp-" + getBaseName(fileName);
        var extension = getExtension(fileName);
        var tempFileName = tempFileNameBase + extension;
        var counter = 1;

        while (File(folderPath + tempFileName).exists) {
            tempFileName = tempFileNameBase + "-" + counter + extension;
            counter++;
        }
        return new File(folderPath + tempFileName);
    }

    main();

})();
