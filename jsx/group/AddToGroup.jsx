#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    var doc = app.documents.length && app.activeDocument;
    if (!doc) return;

    var sel = doc.selection;

    if (sel.length < 2) return;

    // ungroup
    app.executeMenuCommand('ungroup');

    // group
    app.executeMenuCommand('group');

})();