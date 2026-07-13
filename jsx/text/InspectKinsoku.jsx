#target illustrator

// 選択している段落で使われている禁則の値を列挙して確認する

// 選択内容から対象の段落を集めて配列で返す
// 文字ツールでの文字選択は TextRange、選択ツールでのオブジェクト選択は TextFrame の配列になる
function collectSelectedParagraphs(doc) {
    var currentSelection = doc.selection;
    var targetParagraphs = [];

    if (!currentSelection) {
        return targetParagraphs;
    }

    // 文字ツールでテキストを選択した場合
    if (currentSelection.constructor && currentSelection.constructor.name === "TextRange") {
        var selectedParagraphs = currentSelection.paragraphs;
        for (var i = 0; i < selectedParagraphs.length; i++) {
            targetParagraphs.push(selectedParagraphs[i]);
        }
        return targetParagraphs;
    }

    // 選択ツールでオブジェクトを選択した場合（配列）
    for (var i = 0; i < currentSelection.length; i++) {
        var selectedItem = currentSelection[i];
        if (selectedItem.constructor && selectedItem.constructor.name === "TextFrame") {
            var framedParagraphs = selectedItem.textRange.paragraphs;
            for (var j = 0; j < framedParagraphs.length; j++) {
                targetParagraphs.push(framedParagraphs[j]);
            }
        }
    }
    return targetParagraphs;
}

// 段落配列を走査し、検出した禁則値を集合として返す
function collectKinsokuValues(paragraphs) {
    var detectedKinsokuSet = {};

    for (var i = 0; i < paragraphs.length; i++) {
        try {
            // 禁則「なし」の段落は kinsoku を読むだけで Error 9563 を投げるため try は必須
            var kinsokuValue = paragraphs[i].paragraphAttributes.kinsoku;
            detectedKinsokuSet[String(kinsokuValue)] = true;
        } catch (e) {
            // 禁則「なし」の段落は属性が undefined 扱いとなり Error 9563 を投げる
            if (e.number === 9563) {
                detectedKinsokuSet["なし"] = true;
            } else {
                detectedKinsokuSet["ERROR: " + e.message] = true;
            }
        }
    }

    return detectedKinsokuSet;
}

(function () {
    var targetParagraphs = collectSelectedParagraphs(app.activeDocument);
    if (!targetParagraphs.length) {
        alert("段落が選択されていません。\nテキスト、またはテキストフレームを選択してください。");
        return;
    }

    var detectedKinsokuSet = collectKinsokuValues(targetParagraphs);

    var reportText = "検出された禁則値:\n";
    for (var detectedValue in detectedKinsokuSet) {
        reportText += "  → \"" + detectedValue + "\"\n";
    }
    alert(reportText);
})();