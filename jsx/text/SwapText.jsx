#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// 選択中の2つのテキストの文字列を入れ替える

if (app.documents.length === 0) {
    alert("ドキュメントが開かれていません。");
} else {
    var doc = app.activeDocument;
    var selectedItems = doc.selection;

    if (selectedItems.length !== 2) {
        alert("テキストオブジェクトを2つ選択してください。");
    } else if (selectedItems[0].typename !== "TextFrame" || selectedItems[1].typename !== "TextFrame") {
        alert("選択した2つは両方ともテキストオブジェクトである必要があります。");
    } else {
        var firstTextFrame = selectedItems[0];
        var secondTextFrame = selectedItems[1];

        var firstContents = firstTextFrame.contents;
        var secondContents = secondTextFrame.contents;

        firstTextFrame.contents = secondContents;
        secondTextFrame.contents = firstContents;
    }
}
