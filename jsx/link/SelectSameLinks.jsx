#target illustrator

/*
SelectSameLinks.jsx

### 概要 / Overview
- 現在選択中のリンク画像（PlacedItem）と同じリンクファイルを参照する
  PlacedItem をドキュメント内から抽出し、選択／削除を行う。
- Finds every PlacedItem in the active document that references the
  same linked file(s) as the current selection, then selects or deletes
  them according to the dialog options.

### ダイアログ / Dialog
- 判定方法 : 同じパス（絶対パス） / ファイル名一致（パス無視）
- 動作     : 同一リンクを選択 / リンク画像のみ削除 / クリップグループごと削除
*/

var SCRIPT_VERSION = "v1.1";

// =========================
// ロケール / Locale
// =========================
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "同一リンクを選択／削除",
        en: "Select / Delete Same Links"
    },
    matchModePanel: {
        ja: "判定方法",
        en: "Match by"
    },
    matchByPath: {
        ja: "同じパス",
        en: "Same path"
    },
    matchByName: {
        ja: "ファイル名一致",
        en: "File name only"
    },
    actionPanel: {
        ja: "動作",
        en: "Action"
    },
    actionSelect: {
        ja: "同一リンクを選択",
        en: "Select same links"
    },
    actionDeleteImage: {
        ja: "リンク画像のみを削除",
        en: "Delete linked image only"
    },
    actionDeleteClipGroup: {
        ja: "クリップグループごと削除",
        en: "Delete with clip group"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    alertNoDoc: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertNoSelection: {
        ja: "リンク画像（配置画像）を選択してから実行してください。",
        en: "Please select at least one linked (placed) image first."
    },
    alertNoLinkPath: {
        ja: "選択中の配置画像からリンクパスを取得できませんでした。",
        en: "Could not read the linked file path of the selection."
    },
    alertResultSelected: {
        ja: "件選択しました。",
        en: " item(s) selected."
    },
    alertResultDeleted: {
        ja: "件削除しました。",
        en: " item(s) deleted."
    }
};
function L(key) { return LABELS[key][lang]; }
function labelText(key) { return L(key); }

// =========================
// マッチモード / Match mode
// =========================
// 'path' : placedItem.file.fsName（絶対パス）で同一判定
// 'name' : placedItem.file.name（ファイル名のみ）で同一判定
var MATCH_MODE = 'path';

// =========================
// ユーティリティ / Utilities
// =========================
function tryGet(fn, fallback) {
    try { return fn(); } catch (e) { return fallback; }
}

// 選択範囲（およびその子孫）から PlacedItem を集める
// Collect PlacedItems from the selection (descending into GroupItems / clip groups)
function collectPlacedFromSelection(sel) {
    var out = [];
    if (!sel) return out;
    function visit(node) {
        var tn = tryGet(function () { return node.typename; }, "");
        if (tn === "PlacedItem") {
            out.push(node);
            return;
        }
        if (tn === "GroupItem") {
            var kids = tryGet(function () { return node.pageItems; }, null);
            if (!kids) return;
            for (var i = 0; i < kids.length; i++) visit(kids[i]);
        }
    }
    for (var k = 0; k < sel.length; k++) visit(sel[k]);
    return out;
}

// PlacedItem からリンクのキーを取得（取得不可なら null）
// Returns the link key: absolute path (MATCH_MODE='path') or file name only (MATCH_MODE='name')
function getLinkPath(placedItem) {
    var f = tryGet(function () { return placedItem.file; }, null);
    if (!f) return null;
    if (MATCH_MODE === 'name') {
        return tryGet(function () { return f.name; }, null);
    }
    return tryGet(function () { return f.fsName; }, null);
}

// PlacedItem を内包するクリップグループを取得（無ければ null）
// Returns the enclosing clipped GroupItem of a PlacedItem, or null if none.
function getEnclosingClipGroup(item) {
    var node = tryGet(function () { return item.parent; }, null);
    while (node) {
        var tn = tryGet(function () { return node.typename; }, "");
        if (tn !== "GroupItem") break;
        var isClipped = tryGet(function () { return node.clipped; }, false);
        if (isClipped) return node;
        node = tryGet(function () { return node.parent; }, null);
    }
    return null;
}

// =========================
// ダイアログ / Dialog
// =========================
// 判定方法 + 動作を選ばせる。
// 戻り値：{ matchMode: 'path'|'name', action: 'select'|'deleteImage'|'deleteClipGroup' }
//        Cancel 時は null
function showOptionsDialog(initialMode, initialAction) {
    var dlg = new Window("dialog", L('dialogTitle') + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    // 判定方法 / Match mode
    var modePanel = dlg.add("panel", undefined, L('matchModePanel'));
    modePanel.orientation = "column";
    modePanel.alignChildren = ["left", "top"];
    modePanel.margins = [15, 20, 15, 10];

    var pathRadio = modePanel.add("radiobutton", undefined, L('matchByPath'));
    var nameRadio = modePanel.add("radiobutton", undefined, L('matchByName'));
    if (initialMode === 'name') {
        nameRadio.value = true;
    } else {
        pathRadio.value = true;
    }

    // 動作 / Action
    var actionPanel = dlg.add("panel", undefined, L('actionPanel'));
    actionPanel.orientation = "column";
    actionPanel.alignChildren = ["left", "top"];
    actionPanel.margins = [15, 20, 15, 10];

    var selectRadio = actionPanel.add("radiobutton", undefined, L('actionSelect'));
    var deleteImageRadio = actionPanel.add("radiobutton", undefined, L('actionDeleteImage'));
    var deleteClipRadio = actionPanel.add("radiobutton", undefined, L('actionDeleteClipGroup'));
    if (initialAction === 'deleteImage') {
        deleteImageRadio.value = true;
    } else if (initialAction === 'deleteClipGroup') {
        deleteClipRadio.value = true;
    } else {
        selectRadio.value = true;
    }

    // ボタン / Buttons（Mac 規約：Cancel → OK）
    var btnRow = dlg.add("group");
    btnRow.alignment = ["right", "top"];
    var cancelBtn = btnRow.add("button", undefined, L('cancel'), { name: "cancel" });
    var okBtn = btnRow.add("button", undefined, "OK", { name: "ok" });

    var result = null;
    okBtn.onClick = function () {
        var action = 'select';
        if (deleteImageRadio.value) action = 'deleteImage';
        else if (deleteClipRadio.value) action = 'deleteClipGroup';
        result = {
            matchMode: (nameRadio.value ? 'name' : 'path'),
            action: action
        };
        dlg.close();
    };
    cancelBtn.onClick = function () {
        result = null;
        dlg.close();
    };

    dlg.show();
    return result;
}

// =========================
// メイン / Main
// =========================
function main() {
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return;
    }
    var doc = app.activeDocument;
    var sel = tryGet(function () { return doc.selection; }, null);
    if (!sel || sel.length === 0) {
        alert(L('alertNoSelection'));
        return;
    }

    var selPlaced = collectPlacedFromSelection(sel);
    if (selPlaced.length === 0) {
        alert(L('alertNoSelection'));
        return;
    }

    // ダイアログで判定方法と動作を選択
    var opts = showOptionsDialog(MATCH_MODE, 'select');
    if (!opts) return;
    MATCH_MODE = opts.matchMode;

    // 選択中の PlacedItem からリンクキー集合を作成
    var keySet = {};
    var hasAny = false;
    for (var s = 0; s < selPlaced.length; s++) {
        var p = getLinkPath(selPlaced[s]);
        if (p) { keySet[p] = true; hasAny = true; }
    }
    if (!hasAny) {
        alert(L('alertNoLinkPath'));
        return;
    }

    // ドキュメント内のマッチ対象を収集（後段で参照を保ったまま削除するため一旦バッファ）
    var placedItems = doc.placedItems;
    var matches = [];
    for (var i = 0; i < placedItems.length; i++) {
        var item = placedItems[i];
        var key = getLinkPath(item);
        if (key && keySet[key]) matches.push(item);
    }

    if (opts.action === 'select') {
        doc.selection = null;
        var selectedCount = 0;
        for (var j = 0; j < matches.length; j++) {
            try { matches[j].selected = true; selectedCount++; } catch (e) { }
        }
        app.redraw();
        alert(selectedCount + " " + L('alertResultSelected'));
        return;
    }

    // 削除モード：参照が無効化されないよう逆順で remove
    doc.selection = null;
    var deletedCount = 0;
    for (var k = matches.length - 1; k >= 0; k--) {
        var target = matches[k];
        var removed = false;
        if (opts.action === 'deleteClipGroup') {
            var clipGroup = getEnclosingClipGroup(target);
            if (clipGroup) {
                try { clipGroup.remove(); removed = true; } catch (e) { }
            }
        }
        if (!removed) {
            try { target.remove(); removed = true; } catch (e2) { }
        }
        if (removed) deletedCount++;
    }
    app.redraw();
    alert(deletedCount + " " + L('alertResultDeleted'));
}

main();
