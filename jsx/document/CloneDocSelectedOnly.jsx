#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトだけを残した複製ドキュメントを作成するIllustrator用スクリプトです。
元ドキュメントを一時保存してから複製を開き、非選択・非表示のオブジェクトを削除します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloneDocSelectedOnly.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb71dac0648f4

### Overview

An Illustrator script that creates a duplicate document containing only the selected objects.
The original is saved to a temporary file, the copy is opened, and unselected and hidden objects are removed from it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloneDocSelectedOnly.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CloneDocSelectedOnly";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-12-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-18";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloneDocSelectedOnly.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloneDocSelectedOnly.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb71dac0648f4"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var REMOVE_LOCKED_ITEMS = false; /* true: ロックされたオブジェクトやレイヤーも削除 / Remove locked items and layers if true */

    /* 選択オブジェクトに付ける一時タグの名前（一時ファイルを開き直しても選択を見分けるため）
       Temporary tag name on selected items, so they can be found again after the temp file is reopened */
    var KEEP_TAG_NAME = "CloneDocSelectedOnly_keep";

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
     * 元ドキュメントを一時ファイルに保存して開き、選択していないオブジェクトを削除する
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

        /* 元のオブジェクトの参照は開き直したドキュメントでは一致しないので、タグを付けてファイルに残す
           References to the original items do not match in the reopened document, so mark them with a tag saved in the file */
        for (var i = 0; i < selectedItems.length; i++) {
            addKeepTag(selectedItems[i]);
        }

        var tempFile = getUniqueTempFile(originalDoc.fullName, originalDoc.name);
        originalDoc.saveAs(tempFile);
        var duplicateDoc = app.open(tempFile);

        removeUnselectedFromLayers(duplicateDoc.layers);
    }

    // =========================================
    // 選択と削除 / Selection and removal
    // =========================================

    /**
     * 選択されているオブジェクトを配列で返す（文字の編集中は空）
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択オブジェクトの配列
     */
    function getSelectedItems(doc) {
        /* 文字の編集中は TextRange が返る / While editing text, selection is a TextRange */
        if (!(doc.selection instanceof Array)) return [];
        return doc.selection;
    }

    /**
     * オブジェクトに残す目印のタグを付ける
     * @param {PageItem} pageItem - 選択オブジェクト
     * @returns {void}
     */
    function addKeepTag(pageItem) {
        var keepTag = pageItem.tags.add();
        keepTag.name = KEEP_TAG_NAME;
        keepTag.value = "1";
    }

    /**
     * オブジェクト自身に目印のタグがあるか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} タグがあれば true
     */
    function hasKeepTag(pageItem) {
        for (var i = 0; i < pageItem.tags.length; i++) {
            if (pageItem.tags[i].name === KEEP_TAG_NAME) return true;
        }
        return false;
    }

    /**
     * グループや複合パスの中身を返す（それ以外は空）
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {PageItems|PathItems|Array} 子オブジェクト
     */
    function getChildItems(pageItem) {
        if (pageItem.typename === "GroupItem") return pageItem.pageItems;
        if (pageItem.typename === "CompoundPathItem") return pageItem.pathItems;
        return [];
    }

    /**
     * オブジェクト自身か、その中のどれかに目印のタグがあるか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} タグがあれば true
     */
    function containsKeepTag(pageItem) {
        if (hasKeepTag(pageItem)) return true;
        var childItems = getChildItems(pageItem);
        for (var i = 0; i < childItems.length; i++) {
            if (containsKeepTag(childItems[i])) return true;
        }
        return false;
    }

    /**
     * オブジェクトとその中から目印のタグを外す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {void}
     */
    function removeKeepTags(pageItem) {
        for (var i = pageItem.tags.length - 1; i >= 0; i--) {
            if (pageItem.tags[i].name === KEEP_TAG_NAME) pageItem.tags[i].remove();
        }
        var childItems = getChildItems(pageItem);
        for (var j = 0; j < childItems.length; j++) {
            removeKeepTags(childItems[j]);
        }
    }

    /**
     * ロックを外して削除する
     * @param {Layer|PageItem} target - 削除するレイヤーまたはオブジェクト
     * @returns {void}
     */
    function unlockAndRemove(target) {
        target.locked = false;
        target.remove();
    }

    /**
     * レイヤー（サブレイヤーを含む）から選択していないオブジェクトを削除する
     * ロック中のものは REMOVE_LOCKED_ITEMS が true のときだけ削除する
     * @param {Layers} layers - 対象のレイヤー
     * @returns {void}
     */
    function removeUnselectedFromLayers(layers) {
        for (var i = layers.length - 1; i >= 0; i--) {
            var currentLayer = layers[i];

            if (currentLayer.locked) {
                if (REMOVE_LOCKED_ITEMS) unlockAndRemove(currentLayer);
                continue;
            }
            /* 非表示のレイヤーには選択オブジェクトが無い / Hidden layers hold no selected items */
            if (!currentLayer.visible) {
                currentLayer.remove();
                continue;
            }

            removeUnselectedFromLayers(currentLayer.layers);
            removeUnselectedChildren(currentLayer);
        }
    }

    /**
     * レイヤーまたはグループの直下から、選択していないオブジェクトを削除する
     * 選択した子を含むグループは中に入って続け、クリップグループと複合パスは丸ごと残す
     * @param {Layer|GroupItem} container - 対象のレイヤーまたはグループ
     * @returns {void}
     */
    function removeUnselectedChildren(container) {
        /* 削除で並びが変わらないよう、直下のものを先に控える / Collect direct children first so removals do not shift the list */
        var directItems = [];
        for (var i = 0; i < container.pageItems.length; i++) {
            if (container.pageItems[i].parent === container) directItems.push(container.pageItems[i]);
        }

        for (var j = 0; j < directItems.length; j++) {
            var pageItem = directItems[j];
            if (pageItem.locked) {
                if (REMOVE_LOCKED_ITEMS) unlockAndRemove(pageItem);
                continue;
            }
            if (!containsKeepTag(pageItem)) {
                pageItem.remove();
                continue;
            }
            if (!hasKeepTag(pageItem) && pageItem.typename === "GroupItem" && !pageItem.clipped) {
                removeUnselectedChildren(pageItem);
                continue;
            }
            removeKeepTags(pageItem);
        }
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
