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

/* ラベルを現在の言語で取得 / Resolve a label in the current language */
function L(labelObj) {
    return labelObj[currentLanguage] || labelObj.en;
}

// =========================================
// メイン処理 / Main
// =========================================
(function () {
    /* ドキュメントの有無を確認 / Check that a document is open */
    if (app.documents.length === 0) {
        alert(L(LABELS.alert.noDocument));
        return;
    }

    /* 選択の有無を確認 / Check that something is selected */
    if (app.selection.length === 0) {
        alert(L(LABELS.alert.noSelection));
        return;
    }

    var selectedItems = app.selection;
    var processedCount = 0;

    /* 選択した各オブジェクトを処理 / Process each selected object */
    for (var i = 0; i < selectedItems.length; i++) {
        convertPageItem(selectedItems[i]);
    }

    /* 処理対象が無かった場合の通知 / Notify when nothing could be processed */
    if (processedCount === 0) {
        alert(L(LABELS.alert.noText));
    }

    /* オブジェクト種別ごとに振り分け（テキストは変換、グループは再帰）
       Dispatch by object type (convert text, recurse into groups) */
    function convertPageItem(pageItem) {
        if (pageItem.typename === "TextFrame") {
            /* ポイント文字・エリア内文字・パス上文字はすべて TextFrame として扱える
               Point, area, and path text are all TextFrame objects */
            convertTextRangeSizes(pageItem.textRange);
            processedCount++;
        } else if (pageItem.typename === "TextRange") {
            convertTextRangeSizes(pageItem);
            processedCount++;
        } else if (pageItem.typename === "GroupItem") {
            /* グループ内のテキストを再帰的にたどる / Recurse into group contents */
            for (var j = 0; j < pageItem.pageItems.length; j++) {
                convertPageItem(pageItem.pageItems[j]);
            }
        }
    }

    /* テキスト内の各文字サイズを先頭文字サイズへ揃え、差分を比率へ変換
       Unify each character's size to the first character's size, converting the difference into scale */
    function convertTextRangeSizes(textRange) {
        if (textRange.characters.length === 0) return;

        /* 先頭文字のサイズを基準にする / Use the first character's size as the base */
        var baseFontSize = textRange.characters[0].characterAttributes.size;

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
    }
})();
