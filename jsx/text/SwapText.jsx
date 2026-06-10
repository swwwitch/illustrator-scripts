#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中の2つのテキストオブジェクトの文字列（contents）を入れ替えます。

- 選択は2つ、かつ両方ともテキストオブジェクトである必要があります。
- 条件を満たさない場合はアラートを表示します。

### Overview

Swaps the contents of two selected text objects.

- Exactly two objects must be selected, and both must be text objects.
- Shows an alert if the conditions are not met.

### 紹介記事

https://note.com/dtp_tranist/n/n071e09af28a7

*/

var SCRIPT_VERSION = "v1.0.0";

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
