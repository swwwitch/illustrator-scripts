#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// メイン処理 / Main
// =========================================
var selection = app.activeDocument.selection;
for (var i = 0; i < selection.length; i++) {
    selection[i].textRange.characterAttributes.kerningMethod = AutoKernType.OPTICAL;
}
