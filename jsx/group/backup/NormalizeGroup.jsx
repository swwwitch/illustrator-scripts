#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    if (!app.documents.length) return;

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length < 2) return;

    /* Illustrator には undo をグループ化する API が無いため、そのまま実行する。
       InDesign の app.doScript(fn, ScriptLanguage.JAVASCRIPT, ..., UndoModes.ENTIRE_SCRIPT, name)
       は Illustrator では使えない（Illustrator の doScript はアクション再生用で、
       ScriptLanguage / UndoModes も未定義のため ReferenceError になる）。
       取り消しは ungroup / group の 2 ステップに分かれる。
       Illustrator has no undo-grouping API; InDesign's doScript overload is unavailable
       here, so the two menu commands stay as two separate undo steps. */
    app.executeMenuCommand('ungroup');
    app.executeMenuCommand('group');

})();