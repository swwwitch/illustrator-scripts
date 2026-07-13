#target illustrator

/* 選択したテキストの文字ツメを30%に設定する */

(function () {
    if (app.documents.length === 0) return;
    var sel = app.activeDocument.selection;
    if (!sel) return;

    /* テキスト編集モードでは selection が TextRange 単体になる */
    var items = (sel.typename === "TextRange") ? [sel] : sel;

    for (var i = 0; i < items.length; i++) {
        var range = items[i].textRange || items[i];
        try {
            range.characterAttributes.Tsume = 30;
        } catch (e) {}
    }
})();
