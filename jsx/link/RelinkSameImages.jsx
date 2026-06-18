#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択した配置画像を基準に、同じリンクファイルを参照している配置画像をドキュメント内から検索し、指定ファイルへ一括で差し替えます。
- 選択対象は先頭の1件を基準に判定し、グループ内に配置画像がある場合も自動で解決します。
- ローカライズ対応を維持しつつ、グループ選択時の判定を強化し、安全性と保守性を改善しています。

### Overview

- Uses the selected placed image as a reference, finds placed images in the document that reference the same linked file, and batch-replaces them with a specified file.
- The first selected item is used as the reference, and a placed image inside a group is resolved automatically when possible.
- Keeps localization support while improving group-selection detection, safety, and maintainability.

### バージョン履歴 / Version History

- v1.0 (2024-06-15): 初版 / Initial release
- v1.1 (2025-01-20): グループ内の配置画像解決を改善 / Improved resolution of placed images inside groups
- v1.2 (2026-03-09): 安全性と保守性の改善 / Improved safety and maintainability
- v1.2.1 (2026-06-18): 命名整理・ラベルのカテゴリ化・不要な try と分岐の削除によるリファクタ / Refactor — clearer naming, categorized labels, removed unnecessary try blocks and dead branches

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.2.1";

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のUI言語を判定 / Detect the current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義（カテゴリ分け） / Japanese-English label definitions (by category) */
    var LABELS = {
        dialog: {
            selectReplaceFile: { ja: "置換するファイルを選択してください", en: "Select a file to replace with" }
        },
        message: {
            canceled: { ja: "ファイルの選択がキャンセルされました。", en: "File selection was canceled." },
            notPlacedItem: { ja: "選択されたアイテムは配置画像ではありません。", en: "The selected item is not a placed image." },
            groupHasNoPlacedItem: { ja: "選択されたグループ内に配置画像が見つかりません。", en: "No placed image was found inside the selected group." },
            nothingSelected: { ja: "アイテムが選択されていません。", en: "No item is selected." },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noLinkedFile: {
                ja: "選択した配置画像のリンク情報を取得できません。埋め込み画像またはリンク切れの可能性があります。",
                en: "Could not retrieve the link information for the selected placed image. It may be embedded or missing."
            }
        },
        result: {
            replacedCount: { ja: "差し替え完了: ", en: "Replacement complete: " },
            itemsSuffix: { ja: "件", en: " item(s)" }
        }
    };

    /* カテゴリとキーから現在の言語のラベルを取得 / Get a label for the current language by category and key */
    function L(category, key) {
        var entry = LABELS[category] && LABELS[category][key];
        return entry ? entry[lang] : key;
    }

    // =========================================
    // 配置画像の解決・収集 / Resolve & collect placed items
    // =========================================

    /* 配置画像のリンクファイルを安全に取得（埋め込み・リンク切れは null） / Safely get the linked file of a placed item (null when embedded or missing) */
    function getLinkedFile(placedItem) {
        try {
            return placedItem.file;
        } catch (e) {
            return null;
        }
    }

    /* ページアイテムから配置画像を再帰的に解決（グループ内も掘り下げ） / Recursively resolve a placed item from a page item (digging into groups) */
    function resolvePlacedItemFromSelection(pageItem) {
        if (!pageItem) return null;

        if (pageItem.typename === 'PlacedItem') {
            return pageItem;
        }

        if (pageItem.typename === 'GroupItem') {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                var resolved = resolvePlacedItemFromSelection(pageItem.pageItems[i]);
                if (resolved) return resolved;
            }
        }

        return null;
    }

    /* 選択の先頭から基準となる配置画像を取得（取得不可は警告して null） / Get the reference placed item from the first selection (alerts and returns null on failure) */
    function getReferencePlacedItem(doc) {
        if (!doc.selection.length) {
            alert(L('message', 'nothingSelected'));
            return null;
        }

        var selectedItem = doc.selection[0];
        var placedItem = resolvePlacedItemFromSelection(selectedItem);

        if (placedItem) {
            return placedItem;
        }

        if (selectedItem.typename === 'GroupItem') {
            alert(L('message', 'groupHasNoPlacedItem'));
            return null;
        }

        alert(L('message', 'notPlacedItem'));
        return null;
    }

    /* 基準と同じリンクファイルを参照する配置画像をドキュメント全体から収集 / Collect placed items across the document that reference the same linked file as the reference */
    function collectMatchedPlacedItems(doc, referenceItem) {
        var referenceFile = getLinkedFile(referenceItem);
        if (!referenceFile) {
            alert(L('message', 'noLinkedFile'));
            return null;
        }

        var matchedItems = [];
        var placedItems = doc.placedItems;
        var referenceFsName = referenceFile.fsName;

        for (var i = 0; i < placedItems.length; i++) {
            var linkedFile = getLinkedFile(placedItems[i]);
            if (linkedFile && linkedFile.fsName === referenceFsName) {
                matchedItems.push(placedItems[i]);
            }
        }

        return matchedItems;
    }

    // =========================================
    // ファイル選択・差し替え / Choose file & replace
    // =========================================

    /* 差し替え先ファイルをダイアログで選択（キャンセルは null） / Choose the replacement file via dialog (null on cancel) */
    function chooseReplacementFile() {
        var replacementFile = File.openDialog(L('dialog', 'selectReplaceFile'));
        if (!replacementFile) {
            alert(L('message', 'canceled'));
            return null;
        }
        return replacementFile;
    }

    /* 対象の配置画像を差し替え、件数を返す / Replace the target placed items and return the count */
    function replacePlacedItems(placedItems, replacementFile) {
        for (var i = 0; i < placedItems.length; i++) {
            placedItems[i].file = replacementFile;
        }
        return placedItems.length;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* ドキュメント有無を確認 / Ensure a document is open */
    if (!app.documents.length) {
        alert(L('message', 'noDocument'));
        return;
    }

    var doc = app.activeDocument;

    /* 基準画像 → 同一リンクの収集 → 差し替え先選択 → 一括差し替え / Reference image -> collect same-link items -> choose replacement -> batch replace */
    var referenceItem = getReferencePlacedItem(doc);
    if (!referenceItem) return;

    var targetItems = collectMatchedPlacedItems(doc, referenceItem);
    if (!targetItems) return;

    var replacementFile = chooseReplacementFile();
    if (!replacementFile) return;

    var replacedCount = replacePlacedItems(targetItems, replacementFile);

    /* 既存の選択をクリア / Clear the current selection */
    doc.selection = null;

    alert(L('result', 'replacedCount') + replacedCount + L('result', 'itemsSuffix'));

})();
