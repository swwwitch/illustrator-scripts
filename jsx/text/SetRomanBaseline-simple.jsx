#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

if (app.documents.length) {
    var sel = app.activeDocument.selection;
    for (var i = 0; i < sel.length; i++)
        try { sel[i].textRange.characterAttributes.alignment = StyleRunAlignmentType.ROMANBASELINE; } catch (e) {}
}
