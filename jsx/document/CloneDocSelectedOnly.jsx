#target indesign
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトだけを残した複製ドキュメントを作成するInDesign用スクリプトです。
元ドキュメントを一時保存してから複製を開き、非選択・非表示のオブジェクトを削除します。

詳細は README を参照してください。

### Overview

An InDesign script that creates a duplicate document containing only the selected objects.
The original is saved to a temporary file, the copy is opened, and unselected and hidden objects are removed from it.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CloneDocSelectedOnly";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-12-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CloneDocSelectedOnly.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CloneDocSelectedOnly.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var REMOVE_LOCKED_ITEMS = false; // true: ロックされたアイテムやレイヤーも削除 / Remove locked items and layers if true

    // -------------------------------
    // 日英ラベル定義 Define labels
    // -------------------------------
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();
    var LABELS = {
        noDocument: { ja: "開いているドキュメントがありません。", en: "No documents are open." },
        notSaved: { ja: "ドキュメントが一度も保存されていません。先に保存してください。", en: "The document has never been saved. Please save it first." },
        noSelection: { ja: "選択されているオブジェクトがありません。", en: "No objects are selected." }
    };

    // スクリプト開始 // Script start
    function main() {
        if (app.documents.length > 0) {
            var originalDoc = app.activeDocument;
            // Check if the document has been saved at least once
            if (!originalDoc.saved) {
                alert(LABELS.notSaved[lang]);
                return;
            }
            var originalFilePath = originalDoc.fullName;
            var originalFileName = originalDoc.name;

            var selectedItems = getSelectedItems(originalDoc);
            if (selectedItems.length === 0) {
                alert(LABELS.noSelection[lang]);
            } else {
                var tempFileName = generateTempFileName(originalFilePath, getBaseName(originalFileName), getExtension(originalFileName));
                var tempFilePath = new File(originalFilePath.path + "/" + tempFileName);

                originalDoc.saveAs(tempFilePath);
                var duplicateDoc = app.open(tempFilePath);

                removeUnselectedHiddenItems(duplicateDoc.layers, selectedItems);
            }
        } else {
            alert(LABELS.noDocument[lang]);
        }
    }
    // スクリプト終了 // Script end

    // 選択されているオブジェクトを配列で取得 // Get selected items as array
    function getSelectedItems(doc) {
        var items = [];
        for (var i = 0; i < doc.selection.length; i++) {
            items.push(doc.selection[i]);
        }
        return items;
    }

    // 選択されていない非表示のアイテムを削除 // Remove unselected and hidden items
    function removeUnselectedHiddenItems(layers, selectedItems) {
        for (var i = layers.length - 1; i >= 0; i--) {
            var layer = layers[i];

            if (layer.locked) {
                if (REMOVE_LOCKED_ITEMS) {
                    layer.locked = false;
                    layer.remove();
                    continue;
                } else {
                    continue;
                }
            }
            if (!layer.visible) {
                continue;
            }

            for (var j = layer.pageItems.length - 1; j >= 0; j--) {
                var item = layer.pageItems[j];
                if (item.locked) {
                    if (REMOVE_LOCKED_ITEMS) {
                        item.locked = false;
                        item.remove();
                    }
                    continue;
                }
                if (!isItemSelected(item, selectedItems) && !item.visible) {
                    item.remove();
                }
            }
        }
    }

    // アイテムが選択されているか判定 // Check if item is selected
    function isItemSelected(item, selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (item === selectedItems[i]) {
                return true;
            }
        }
        return false;
    }

    // ファイル名のベース部分を取得 // Get base part of file name
    function getBaseName(fileName) {
        var parts = fileName.split('.');
        return parts[0];
    }

    // ファイル名の拡張子を取得（ドット含む） // Get file extension (with dot)
    function getExtension(fileName) {
        var parts = fileName.split('.');
        return parts.length > 1 ? '.' + parts[parts.length - 1] : '';
    }

    // 一時ファイル名を生成 // Generate temporary file name
    function generateTempFileName(originalFilePath, baseName, extension) {
        var tempFileNameBase = "temp-" + baseName;
        var tempFileName = tempFileNameBase + extension;
        var counter = 1;

        while (File(originalFilePath.path + "/" + tempFileName).exists) {
            tempFileName = tempFileNameBase + "-" + counter + extension;
            counter++;
        }
        return tempFileName;
    }

    main();

})();
