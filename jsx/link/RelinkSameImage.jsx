#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した配置画像と同じリンクファイルを参照している配置画像を探し、指定したファイルへ一括で差し替えます。

詳細は README を参照してください。

### Overview

Finds the placed images that reference the same linked file as the selected one and relinks them all to a file you choose.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RelinkSameImage";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-08-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-21";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RelinkSameImage.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RelinkSameImage.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne38eeee5abc8?nt=_3084117"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
