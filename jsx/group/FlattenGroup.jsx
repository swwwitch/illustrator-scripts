#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

var doc = app.documents.length && app.activeDocument;
if (!doc) return;

var sel = doc.selection;
if (!sel.length) return;

    if (!sel || sel.length < 1) return;

    // ungroup all
    app.executeMenuCommand('ungroupAll');

    // group
    app.executeMenuCommand('group');

})();