#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの段落に対して、ジャスティフィケーション関連の設定（ワードスペース、文字間、グリフスケーリング）を初期値に戻します。

詳細は README を参照してください。

### Overview

Resets the justification settings — word spacing, letter spacing and glyph scaling — of the selected paragraphs to their defaults.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResetJustification";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetJustification.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetJustification.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function resetJustification(targetTextRange) {
        var paragraphs = targetTextRange.paragraphs;
        for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
            var paragraphAttributes = paragraphs[paragraphIndex].paragraphAttributes;

            // Word Spacing
            paragraphAttributes.minimumWordSpacing = 80;
            paragraphAttributes.desiredWordSpacing = 100;
            paragraphAttributes.maximumWordSpacing = 133;

            // Letter Spacing
            paragraphAttributes.minimumLetterSpacing = 0;
            paragraphAttributes.desiredLetterSpacing = 0;
            paragraphAttributes.maximumLetterSpacing = 0;

            // Glyph Scaling
            paragraphAttributes.minimumGlyphScaling = 100;
            paragraphAttributes.desiredGlyphScaling = 100;
            paragraphAttributes.maximumGlyphScaling = 100;
        }
    }

    var activeDocument = app.activeDocument;
    var selectionItems = activeDocument.selection;

    if (selectionItems.length === 0) {
        alert("テキストを選択してください。");
    } else {
        var targetTextRange = null;

        if (selectionItems[0].typename === "TextFrame") {
            targetTextRange = selectionItems[0].textRange;
        } else if (selectionItems[0].story) {
            targetTextRange = selectionItems[0];
        }

        if (targetTextRange === null) {
            alert("テキストが選択されていません。");
        } else {
            resetJustification(targetTextRange);
            alert("Justification を初期化しました。");
        }
    }
})();
