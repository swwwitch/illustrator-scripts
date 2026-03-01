#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
MenuCommandBatch

### 更新日：
20260301-5

### 概要：
現在選択しているオブジェクトのみに対して、Illustratorのメニューコマンドを連続実行します。
（オフセットパス → グループ化 → パスファインダー：合流 → 形状を拡張）

### 使い方：
1) ドキュメントを開き、対象を選択
2) スクリプトを実行

※メニューコマンドはIllustratorの状態や環境に依存します。エラー時は中断します。
*/

(function () {
    if (app.documents.length === 0) {
        alert('ドキュメントが開かれていません。');
        return;
    }

    var doc = app.activeDocument;

    // selection check
    if (!doc.selection || doc.selection.length === 0) {
        alert('オブジェクトを選択してから実行してください。');
        return;
    }

    function runMenu(cmd) {
        try {
            app.executeMenuCommand(cmd);
            return true;
        } catch (e) {
            alert('メニューコマンドの実行に失敗しました：\n' + cmd + '\n\n' + e);
            return false;
        }
    }

    function snapshotPageItems() {
        var arr = [];
        // doc.pageItems includes many types; we store references to detect newly created items
        for (var i = 0; i < doc.pageItems.length; i++) {
            arr.push(doc.pageItems[i]);
        }
        return arr;
    }

    function refExists(arr, obj) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === obj) return true;
        }
        return false;
    }

    function diffNewItems(beforeArr, afterArr) {
        var out = [];
        for (var i = 0; i < afterArr.length; i++) {
            var it = afterArr[i];
            if (!refExists(beforeArr, it)) {
                // Avoid selecting locked/hidden items when possible
                try {
                    if (it.locked) continue;
                    if (it.hidden) continue;
                } catch (_) { }
                out.push(it);
            }
        }
        return out;
    }

    function restoreSelectionSafe(items) {
        var sel = [];
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            try {
                if (it.locked) continue;
                if (it.hidden) continue;
            } catch (_) { }
            sel.push(it);
        }
        try {
            doc.selection = sel;
        } catch (_) {
            // If assignment fails, fall back to whatever Illustrator keeps selected
        }
    }

    function pushUnique(arr, item) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === item) return;
        }
        arr.push(item);
    }

    function unionSelectable(a, b) {
        var out = [];
        var i, it;
        if (a && a.length) {
            for (i = 0; i < a.length; i++) {
                it = a[i];
                if (!it) continue;
                try {
                    if (it.locked) continue;
                    if (it.hidden) continue;
                } catch (_) { }
                pushUnique(out, it);
            }
        }
        if (b && b.length) {
            for (i = 0; i < b.length; i++) {
                it = b[i];
                if (!it) continue;
                try {
                    if (it.locked) continue;
                    if (it.hidden) continue;
                } catch (_) { }
                pushUnique(out, it);
            }
        }
        return out;
    }

    function main() {
        // Preserve original selection (references)
        var originalSel = [];
        try {
            for (var i = 0; i < doc.selection.length; i++) originalSel.push(doc.selection[i]);
        } catch (_) { }

        // Snapshot existing items so we can detect Offset Path results
        var beforeItems = snapshotPageItems();

        // Offset Path (opens dialog)
        if (!runMenu('OffsetPath v22')) return;

        // If selection got lost/changed unexpectedly, try to fix it by selecting newly created items
        var afterItems = snapshotPageItems();
        var newItems = diffNewItems(beforeItems, afterItems);

        // Goal: after Offset Path, keep BOTH the original selection and the newly created offset results selected
        // so that subsequent Group / Pathfinder / Expand target the intended combined set.
        if (newItems.length > 0) {
            var unionSel = unionSelectable(originalSel, newItems);
            if (unionSel.length > 0) {
                restoreSelectionSafe(unionSel);
            } else {
                // Fallback: at least try selecting the new items
                restoreSelectionSafe(newItems);
            }
        } else {
            // No detectable new items (or diff failed). If nothing is selected, restore original selection.
            try {
                if (!doc.selection || doc.selection.length === 0) {
                    restoreSelectionSafe(originalSel);
                }
            } catch (_) {
                restoreSelectionSafe(originalSel);
            }
        }

        // Group
        if (!runMenu('group')) return;

        // Live Pathfinder Merge
        if (!runMenu('Live Pathfinder Merge')) return;

        // Expand Appearance / Expand Style
        if (!runMenu('expandStyle')) return;
    }

    // Run as a single undo step where possible (AI version dependent)
    if (doc && typeof doc.suspendHistory === 'function') {
        doc.suspendHistory('MenuCommandBatch', 'main()');
    } else {
        // Fallback for environments/versions without suspendHistory
        main();
    }
})();