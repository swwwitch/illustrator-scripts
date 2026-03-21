#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    if (!app.documents.length) return;

    var doc = app.activeDocument;
    var sel = doc.selection;

    // if (sel.length < 2) return;

     try { app.executeMenuCommand('ungroupAll'); } catch (e) {}

    // 記録：最初に選択されていたオブジェクト
    var initial = [];
    for (var i = 0; i < sel.length; i++) {
        initial.push(sel[i]);
    }

    // ungroupAll は環境/状況によって効かないことがあるため、フォールバック付きで深く解除
    ungroupDeepSelection(doc);

    // 改めてオブジェクトを選択
    doc.selection = null;
    for (var j = 0; j < initial.length; j++) {
        try {
            initial[j].selected = true;
        } catch (e) {}
    }

    app.executeMenuCommand('group');


    function ungroupDeepSelection(doc) {
        // まず ungroupAll を試す
        try { app.executeMenuCommand('ungroupAll'); } catch (e) {}

        // それでもグループが残る場合があるので、ungroup を繰り返す
        // 無限ループ防止のため上限を設ける
        for (var k = 0; k < 200; k++) {
            if (!selectionHasGroup(doc.selection)) break;
            try { app.executeMenuCommand('ungroup'); } catch (e) { break; }
        }
    }

    function selectionHasGroup(sel) {
        if (!sel || !sel.length) return false;
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            try {
                if (it.typename === 'GroupItem') return true;
            } catch (e) {}
        }
        return false;
    }

})();