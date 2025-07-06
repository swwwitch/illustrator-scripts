#target indesign
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

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

/*

スクリプト名：CloneDocSelectedOnly.jsx // Script Name: CloneDocSelectedOnly.jsx

概要: 現在のドキュメントを一時保存し、選択オブジェクトのみ残した複製ドキュメントを作成するスクリプト。 // Description: This script temporarily saves the current document, duplicates it, and keeps only the selected objects in the new document.

処理の流れ: // Process flow:
 1. アクティブドキュメントを取得 // Get active document
 2. 選択オブジェクトを取得 // Get selected objects
 3. 一時ファイル名を生成し保存 // Generate temporary file name and save
 4. 複製ドキュメントを開く // Open duplicated document
 5. 選択されていない非表示オブジェクトを削除 // Remove unselected hidden objects

限定条件: // Limitations:
 - InDesign 2025 以降推奨 // Recommended for InDesign 2025 or later
 - 選択オブジェクトがある場合のみ有効 // Only works if objects are selected

更新履歴: // Update history:
 - v1.0.0 2023-12-26 初版リリース // Initial release
 - v1.0.1 2025-07-02 微調整 // Minor adjustments

*/

// ===== 設定セクション ===== // Settings section
var REMOVE_LOCKED_ITEMS = false; // true: ロックされたアイテムやレイヤーも削除 / Remove locked items and layers if true

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