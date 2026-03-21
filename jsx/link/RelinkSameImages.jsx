#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
- 選択した配置画像を基準に、同じリンクファイルを参照している配置画像をドキュメント内から検索し、指定ファイルへ一括で差し替えます。 / Uses the selected placed image as a reference, finds placed images in the document that reference the same linked file, and batch-replaces them with a specified file.
- 選択対象は先頭の1件を基準に判定し、グループ内に配置画像がある場合も自動で解決します。 / The first selected item is used as the reference, and a placed image inside a group is resolved automatically when possible.
- ローカライズ対応を維持しつつ、グループ選択時の判定を強化し、安全性と保守性を改善しています。 / Keeps localization support while improving group-selection detection, safety, and maintainability.
 
更新日 / Updated: 2026-03-09
バージョン / Version: v1.2
*/

(function () {

    var SCRIPT_VERSION = "v1.2";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        scriptName: {
            ja: "画像差し替え一括",
            en: "Batch Image Relinker"
        },
        selectReplaceFile: {
            ja: "置換するファイルを選択してください",
            en: "Select a file to replace with"
        },
        canceled: {
            ja: "ファイルの選択がキャンセルされました。",
            en: "File selection was canceled."
        },
        notPlacedItem: {
            ja: "選択されたアイテムは配置画像ではありません。",
            en: "The selected item is not a placed image."
        },
        groupHasNoPlacedItem: {
            ja: "選択されたグループ内に配置画像が見つかりません。",
            en: "No placed image was found inside the selected group."
        },
        nothingSelected: {
            ja: "アイテムが選択されていません。",
            en: "No item is selected."
        },
        noDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        noLinkedFile: {
            ja: "選択した配置画像のリンク情報を取得できません。埋め込み画像またはリンク切れの可能性があります。",
            en: "Could not retrieve the link information for the selected placed image. It may be embedded or missing."
        },
        replacedCount: {
            ja: "差し替え完了: ",
            en: "Replacement complete: "
        },
        itemsSuffix: {
            ja: "件",
            en: " item(s)"
        }
    };

    function L(key) {
        var item = LABELS[key];
        return item ? item[lang] : key;
    }

    /* 配置画像のリンクファイルを安全に取得 / Safely get the linked file of a placed item */
    function getLinkedFile(item) {
        try {
            return item.file;
        } catch (e) {
            return null;
        }
    }

    /* 選択アイテムから配置画像を解決 / Resolve a placed item from the selected item */
    function resolvePlacedItemFromSelection(item) {
        if (!item) return null;

        if (item.typename === 'PlacedItem') {
            return item;
        }

        if (item.typename === 'GroupItem') {
            try {
                if (item.placedItems && item.placedItems.length > 0) {
                    return item.placedItems[0];
                }
            } catch (eGroupPlaced) {}

            try {
                if (item.pageItems && item.pageItems.length > 0) {
                    for (var i = 0; i < item.pageItems.length; i++) {
                        var resolved = resolvePlacedItemFromSelection(item.pageItems[i]);
                        if (resolved) return resolved;
                    }
                }
            } catch (eGroupPageItems) {}
        }

        if (item.typename === 'CompoundPathItem') {
            try {
                if (item.pathItems && item.pathItems.length > 0) {
                    for (var j = 0; j < item.pathItems.length; j++) {
                        var resolvedFromPath = resolvePlacedItemFromSelection(item.pathItems[j]);
                        if (resolvedFromPath) return resolvedFromPath;
                    }
                }
            } catch (eCompound) {}
        }

        return null;
    }

    /* 基準となる配置画像を取得 / Get the reference placed item */
    function getReferencePlacedItem(doc) {
        if (!doc.selection.length) {
            alert(L('nothingSelected'));
            return null;
        }

        var selectedItem = doc.selection[0];
        var placedItem = resolvePlacedItemFromSelection(selectedItem);

        if (placedItem) {
            return placedItem;
        }

        if (selectedItem.typename === 'GroupItem') {
            alert(L('groupHasNoPlacedItem'));
            return null;
        }

        alert(L('notPlacedItem'));
        return null;
    }

    /* 同じリンクファイルを参照する配置画像を収集 / Collect placed items that reference the same linked file */
    function collectMatchedPlacedItems(doc, referenceItem) {
        var referenceFile = getLinkedFile(referenceItem);
        if (!referenceFile) {
            alert(L('noLinkedFile'));
            return null;
        }

        var targets = [];
        var placedItems = doc.placedItems;
        var referenceFsName = referenceFile.fsName;

        for (var i = 0; i < placedItems.length; i++) {
            var linkedFile = getLinkedFile(placedItems[i]);
            if (linkedFile && linkedFile.fsName === referenceFsName) {
                targets.push(placedItems[i]);
            }
        }

        return targets;
    }

    /* 置換ファイルを選択 / Choose a replacement file */
    function chooseReplacementFile() {
        var file = File.openDialog(L('selectReplaceFile'));
        if (!file) {
            alert(L('canceled'));
            return null;
        }
        return file;
    }

    /* 配置画像を差し替え / Replace placed items */
    function replacePlacedItems(items, fileToReplace) {
        for (var i = 0; i < items.length; i++) {
            items[i].file = fileToReplace;
        }
        return items.length;
    }

    /* 実行 / Run */
    if (!app.documents.length) {
        alert(L('noDocument'));
        return;
    }

    var doc = app.activeDocument;
    var referenceItem = getReferencePlacedItem(doc);
    if (!referenceItem) return;

    var targetItems = collectMatchedPlacedItems(doc, referenceItem);
    if (!targetItems) return;

    var fileToReplace = chooseReplacementFile();
    if (!fileToReplace) return;

    var replacedCount = replacePlacedItems(targetItems, fileToReplace);

    /* 既存の選択をクリア / Clear the current selection */
    doc.selection = null;

    alert(L('replacedCount') + replacedCount + L('itemsSuffix'));

})();