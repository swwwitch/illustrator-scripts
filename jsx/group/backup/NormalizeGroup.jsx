#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    if (!app.documents.length) return;

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length < 2) return;

    app.doScript(function () {
        app.executeMenuCommand('ungroup');
        app.executeMenuCommand('group');
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'Normalize Group Order');

})();