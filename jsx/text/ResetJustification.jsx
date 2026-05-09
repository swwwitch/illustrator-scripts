#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview

選択したテキストの段落に対して、ジャスティフィケーション関連の設定（ワードスペース、文字間、グリフスケーリング）を初期値にリセットします。

・Word Spacing：80 / 100 / 133
・Letter Spacing：0 / 0 / 0
・Glyph Scaling：100 / 100 / 100

選択がない場合、またはテキストでない場合はアラートを表示します。

Resets justification-related settings (word spacing, letter spacing, glyph scaling)
of the selected text paragraphs to default values.

- Word Spacing: 80 / 100 / 133
- Letter Spacing: 0 / 0 / 0
- Glyph Scaling: 100 / 100 / 100

Shows an alert if no valid text is selected.
*/

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