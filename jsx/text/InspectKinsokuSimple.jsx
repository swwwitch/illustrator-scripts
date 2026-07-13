#target illustrator

// 選択したテキストフレームの段落で使われている禁則の値を列挙して確認する
(function () {
    var selection = app.activeDocument.selection;
    var kinsokuSet = {};

    for (var i = 0; i < selection.length; i++) {
        if (selection[i].constructor.name !== "TextFrame") continue;
        var paragraphs = selection[i].textRange.paragraphs;
        for (var j = 0; j < paragraphs.length; j++) {
            try {
                kinsokuSet[String(paragraphs[j].paragraphAttributes.kinsoku)] = true;
            } catch (e) {
                // 禁則「なし」は属性が undefined 扱いで Error 9563 を投げる
                kinsokuSet[e.number === 9563 ? "なし" : "ERROR: " + e.message] = true;
            }
        }
    }

    var report = "検出された禁則値:\n";
    for (var value in kinsokuSet) {
        report += "  → \"" + value + "\"\n";
    }
    alert(report);
})();
