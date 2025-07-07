#target indesign
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

CloneDocSelectedOnly.jsx

### 概要

- 選択オブジェクトのみを残した複製ドキュメントを作成するInDesign用スクリプトです。
- 元ドキュメントを一時保存後、不要なオブジェクトを削除して新しいドキュメントを生成します。

### 主な機能

- 選択オブジェクトのみを保持した複製ドキュメント作成
- 非表示・非選択オブジェクトの削除
- ロックアイテム削除オプション
- 日本語／英語インターフェース対応

### 処理の流れ

1. アクティブドキュメントと選択オブジェクトを確認
2. 一時ファイル名を生成してドキュメントを保存
3. 複製ドキュメントを開く
4. 非選択オブジェクトや非表示アイテムを削除
5. 必要に応じてロックアイテムを削除

### 更新履歴

- v1.0.0 (20231226) : 初版リリース
- v1.0.1 (20250702) : 処理微調整

---

### Script Name:

CloneDocSelectedOnly.jsx

### Overview

- An InDesign script that creates a duplicate document containing only the selected objects.
- Saves the original document temporarily, then removes unnecessary objects to create a clean copy.

### Main Features

- Create a duplicate document with only selected objects
- Remove hidden and unselected objects
- Option to remove locked items
- Japanese and English UI support

### Process Flow

1. Check active document and selected objects
2. Generate temporary file name and save document
3. Open the duplicate document
4. Remove unselected and hidden items
5. Optionally remove locked items

### Update History

- v1.0.0 (20231226): Initial release
- v1.0.1 (20250702): Minor adjustments
*/



// ===== 設定セクション ===== // Settings section
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