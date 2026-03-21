// Illustrator用のJavaScript
// 選択しているオブジェクトごとにアピアランスを分割

var doc = app.activeDocument;
var sel = doc.selection;

if (sel.length === 0) {
    alert("オブジェクトを選択してください。");
} else {
    for (var i = sel.length - 1; i >= 0; i--) {
        doc.selection = null;
        sel[i].selected = true;
        app.executeMenuCommand('expandStyle');
        // app.executeMenuCommand('group');
    }
}