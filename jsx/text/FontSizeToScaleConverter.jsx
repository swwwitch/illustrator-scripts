#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択テキスト内で文字サイズが混在しているとき、各テキストの先頭文字のサイズへ統一する
- 見た目の大きさは変えず、サイズ差を水平比率・垂直比率（horizontalScale / verticalScale）に変換して補正する
- ポイント文字・エリア内文字・パス上文字、グループ内のテキスト、複数選択に対応

### Overview

- Unifies mixed font sizes within selected text to the size of each text's first character
- Keeps the visual size unchanged by converting the size difference into horizontal / vertical scale
- Supports point / area / path text, text inside groups, and multiple selections

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ローカライズ / Localization
// =========================================
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
        noText: { ja: "処理できるテキストが選択されていません。", en: "No editable text is selected." }
    }
};

/* キーからラベルを現在の言語で取得（"alert.noDocument" のようにドット区切り）
   Resolve a label by key in the current language (dot-separated, e.g. "alert.noDocument") */
function getLocalizedText(key) {
    var parts = key.split(".");
    var label = LABELS[parts[0]][parts[1]];
    return label[currentLanguage] || label.en;
}

// =========================================
// メイン処理 / Main
// =========================================
(function () {
    /* ドキュメントの有無を確認 / Check that a document is open */
    if (app.documents.length === 0) {
        alert(getLocalizedText("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;

    /* 選択の有無を確認 / Check that something is selected */
    if (doc.selection.length === 0) {
        alert(getLocalizedText("alert.noSelection"));
        return;
    }

    var selectedItems = [];
    var targetTextCount = 0;

    for (var i = 0; i < doc.selection.length; i++) {
        selectedItems.push(doc.selection[i]);
    }

    /* 選択した各オブジェクトを処理 / Process each selected object */
    for (var i = 0; i < selectedItems.length; i++) {
        convertPageItem(selectedItems[i]);
    }

    /* 処理対象が無かった場合の通知 / Notify when nothing could be processed */
    if (targetTextCount === 0) {
        alert(getLocalizedText("alert.noText"));
    }

    /* オブジェクト種別ごとに振り分け（テキストは変換、グループは再帰）
       Dispatch by object type (convert text, recurse into groups) */
    function convertPageItem(pageItem) {
        /* ロック中・非表示のオブジェクトはスキップ（TextRange は該当プロパティが無く undefined＝対象）
           Skip locked or hidden objects (TextRange has no such property → undefined, so it is processed) */
        if (pageItem.locked || pageItem.hidden) return;

        if (pageItem.typename === "TextFrame") {
            /* ポイント文字・エリア内文字・パス上文字はすべて TextFrame として扱える
               Point, area, and path text are all TextFrame objects */
            if (convertTextRangeSizes(pageItem.textRange)) targetTextCount++;
        } else if (pageItem.typename === "TextRange") {
            if (convertTextRangeSizes(pageItem)) targetTextCount++;
        } else if (pageItem.typename === "GroupItem") {
            /* グループ内のテキストを再帰的にたどる / Recurse into group contents */
            for (var j = 0; j < pageItem.pageItems.length; j++) {
                convertPageItem(pageItem.pageItems[j]);
            }
        }
    }

    /* テキスト内の各文字サイズを先頭文字サイズへ揃え、差分を比率へ変換（空なら false を返す）
       Unify each character's size to the first character's size, converting the difference into scale (returns false when empty) */
    function convertTextRangeSizes(textRange) {
        /* 空テキストは処理対象に数えない / Do not count empty text as processed */
        if (textRange.characters.length === 0) return false;

        /* 先頭文字のサイズを基準にする / Use the first character's size as the base */
        var baseFontSize = textRange.characters[0].characterAttributes.size;

        /* 基準サイズが 0 や不正値なら除算できないのでスキップ / Skip when the base size is zero or invalid (cannot divide) */
        if (!baseFontSize || baseFontSize <= 0) return false;

        for (var i = 0; i < textRange.characters.length; i++) {
            var charAttributes = textRange.characters[i].characterAttributes;
            var originalFontSize = charAttributes.size;

            /* サイズが異なる文字だけ補正 / Adjust only characters whose size differs */
            if (originalFontSize !== baseFontSize) {
                var currentHorizontalScale = charAttributes.horizontalScale;
                var currentVerticalScale = charAttributes.verticalScale;

                /* 見た目の大きさを保つための倍率 / Ratio that preserves the visual size */
                var sizeRatio = originalFontSize / baseFontSize;

                charAttributes.size = baseFontSize;
                charAttributes.horizontalScale = currentHorizontalScale * sizeRatio;
                charAttributes.verticalScale = currentVerticalScale * sizeRatio;
            }
        }
        return true;
    }
})();
