#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

OverlapRemover.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OverlapRemover.md

### 概要：

- 更新日：2026-03-01
- 選択オブジェクトに対して Illustrator のメニューコマンドを連続実行
- オフセットパス → グループ化 → パスファインダー：合流 → 形状を拡張

### 主な機能：

- 選択オブジェクトのみを処理対象とする
- オフセットパス実行後、元選択と新規生成物を結合して後続処理に引き渡し
- 失敗時はアラートを表示して中断
- 可能な環境では suspendHistory による単一 Undo ステップ化

### 更新履歴：

- v1.0.0 (2026-03-01) : 初期バージョン

*/

/*

### Script Name:

OverlapRemover.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OverlapRemover.md

### Description:

- Last Updated: 2026-03-01
- Run a sequence of Illustrator menu commands against the current selection
- Offset Path → Group → Pathfinder: Merge → Expand Shape

### Main Features:

- Operates only on the current selection
- After Offset Path, combines the original selection with newly created items for follow-up steps
- Aborts with an alert on failure
- Wraps the run in a single Undo step via suspendHistory when available

### Changelog:

- v1.0.0 (2026-03-01) : Initial version

*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

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