#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

RelinkSameImage.jsx

### Readme （GitHub）：

（未指定）

### 概要：

- 選択したリンク画像と同名のすべてのリンク画像を検索
- ユーザーが指定したファイルで一括置換する

### 主な機能：

- 配置画像1つを選ぶだけで、同名リンクを全自動で置換
- 日本語と英語のローカライズに対応
- 置換対象選択のダイアログ表示あり

### 処理の流れ：

- ドキュメントと選択状態の確認
- 選択が1件かつ配置画像か判定
- 同名リンク画像をすべて取得
- 新ファイルを選択し、全対象を置換
- 最後に選択を解除

### note

https://note.com/dtp_tranist/n/ne38eeee5abc8?nt=_3084117

### 更新履歴：

- v1.0 (20240805) : 初期バージョン作成
- v1.1 (20250721) : ローカライズ

### Script Name:

RelinkSameImage.jsx

### Readme (GitHub):

(Unspecified)

### Overview:

- Find all linked images with the same name as the selected one
- Replace them in bulk using a file selected by the user

### Features:

- Automatically replaces all same-name linked images by selecting just one
- Supports Japanese and English localization
- Prompts user to select replacement file via dialog

### Workflow:

- Check if document is open and an item is selected
- Ensure only one item is selected and it's a placed image
- Gather all linked items with the same file name
- Prompt user to select a new file
- Replace all targets and clear selection

### Changelog:

- v1.0 (20240805): Initial release
- v1.1 (20250721): Localization added
*/

var SCRIPT_VERSION = "v1.2";

function getCurrentLang() {
    // 言語を取得し、"ja" または "en" で返す / Get language code as "ja" or "en"
    var lang = app.locale || $.locale || "en";
    if (lang.indexOf("ja") === 0) {
        return "ja";
    } else {
        return "en";
    }
}

var LANG = getCurrentLang();

var LABELS = {
    fileSelect: {
        ja: "置換するファイルを選択してください",
        en: "Please select a file to replace"
    },
    cancelSelect: {
        ja: "ファイルの選択がキャンセルされました。",
        en: "File selection was cancelled."
    },
    notPlacedItem: {
        ja: "選択されたアイテムは配置画像ではありません。",
        en: "The selected item is not a placed image."
    },
    noSelection: {
        ja: "アイテムが選択されていません。",
        en: "No item is selected."
    },
    noDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    selectOneOnly: {
        ja: "1つだけ選択してください。",
        en: "Please select only one item."
    }
};

function main() {
    /* ドキュメントが開かれているか確認 / Check if any document is open */
    if (app.documents.length > 0) {
        var doc = app.activeDocument;
        
        /* アイテムが選択されているか確認 / Check if any item is selected */
        if (doc.selection.length > 0) {
            if (doc.selection.length > 1) {
                alert(LABELS.selectOneOnly[LANG]);
                return;
            }
            var selectedItem = doc.selection[0];
            
            /* 選択されたアイテムが配置画像か確認 / Check if the selected item is a placed image */
            if (selectedItem.typename == "PlacedItem") {
                if (!selectedItem.file || !selectedItem.file.name) {
                    alert("リンク画像のファイル名が取得できません。");
                    return;
                }
                /* ファイル名を取得 / Get file name */
                var fileName = selectedItem.file.name;

                /* 同名のリンク画像を収集 / Collect all placed items with the same file name */
                var linkedItems = doc.placedItems;
                var selectedItems = [];
                for (var i = 0; i < linkedItems.length; i++) {
                    if (linkedItems[i].file.name.toLowerCase() == fileName.toLowerCase()) {
                        selectedItems.push(linkedItems[i]);
                    }
                }

                /* ユーザーに置換用ファイルを選択させる / Let user select replacement file */
                var fileToReplace = File.openDialog(LABELS.fileSelect[LANG]);
                
                if (fileToReplace != null) {
                    /* すべての対象画像を新しいファイルに置換 / Replace each matched image with selected file */
                    var replacedCount = replaceLinkedItems(selectedItems, fileToReplace);
                    alert(replacedCount + " 件のリンク画像を置換しました。");
                    /* 選択を解除 / Clear selection */
                    doc.selection = null;
                } else {
                    alert(LABELS.cancelSelect[LANG]);
                }
            } else {
                alert(LABELS.notPlacedItem[LANG]);
            }
        } else {
            alert(LABELS.noSelection[LANG]);
        }
    } else {
        alert(LABELS.noDocument[LANG]);
    }
}

/**
 * 指定したリンクアイテム群を新しいファイルで置換する
 * Replace each linked item with the new file
 */
function replaceLinkedItems(items, newFile) {
    var count = 0;
    for (var i = 0; i < items.length; i++) {
        if (items[i].file && items[i].file.fsName !== newFile.fsName) {
            items[i].file = newFile;
            count++;
        }
    }
    return count;
}

main();