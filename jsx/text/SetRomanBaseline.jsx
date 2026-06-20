#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

文字揃えの変更.jsx

- 選択したテキストの文字揃え（ベースライン）をまとめて変更
- ダイアログを閉じずにライブプレビューで結果を確認
- TextFrame / TextRange / グループ内テキストへ再帰的に適用

### 更新履歴

- v1.0.0 : 初版

### 紹介記事（note: Japanese only）

https://note.com/dtp_tranist/n/n0f1203d5f8f8

### Overview

bunjisoroe-no-henkou.jsx

- Batch-change the character alignment (baseline) of selected text
- Live preview shows the result without closing the dialog
- Recursively applies to TextFrame / TextRange / text inside groups

### Changelog

- v1.0.0 : Initial release

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ローカライズ / Localization
// =========================================
var lang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

var LABELS = {
    dialog: {
        title: { ja: "文字揃えを選択", en: "Select Character Alignment" },
        prompt: { ja: "適用する文字揃えを選択してください：", en: "Select the alignment to apply:" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    align: {
        romanBaseline: { ja: "欧文ベースライン", en: "Roman Baseline" },
        top: { ja: "仮想ボディの上{slash}右", en: "Embox Top{slash}Right" },
        center: { ja: "中央", en: "Embox Centre" },
        bottom: { ja: "仮想ボディの下{slash}左", en: "Embox Bottom{slash}Left" },
        icfTop: { ja: "平均字面の上{slash}右", en: "ICF Box Top{slash}Right" },
        icfBottom: { ja: "平均字面の下{slash}左", en: "ICF Box Bottom{slash}Left" }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Please select a text object." },
        noText: { ja: "選択範囲にテキストがありませんでした。", en: "No text was found in the selection." },
        applied: {
            ja: "{count} 個のテキストに「{label}」を適用しました。",
            en: "Applied “{label}” to {count} text object(s)."
        }
    }
};

/* ラベルを現在の言語で取得し {slash} を / に置換 / Resolve a label for the current language, replacing {slash} with / */
function L(entry) {
    var text = (entry && entry[lang]) ? entry[lang] : (entry ? entry.en : "");
    return text.replace(/\{slash\}/g, "/");
}

(function () {
    if (app.documents.length === 0) {
        alert(L(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {
        alert(L(LABELS.alert.noSelection));
        return;
    }

    var alignList = [
        { label: LABELS.align.romanBaseline, value: StyleRunAlignmentType.ROMANBASELINE },
        { label: LABELS.align.top,           value: StyleRunAlignmentType.top },
        { label: LABELS.align.center,        value: StyleRunAlignmentType.center },
        { label: LABELS.align.bottom,        value: StyleRunAlignmentType.bottom },
        { label: LABELS.align.icfTop,        value: StyleRunAlignmentType.icfTop },
        { label: LABELS.align.icfBottom,     value: StyleRunAlignmentType.icfBottom }
    ];

    // 選択を確定時点で固定 / 取り消し後も参照が壊れないようコピーしておく
    var targetItems = [];
    for (var s = 0; s < doc.selection.length; s++) {
        targetItems.push(doc.selection[s]);
    }

    // プレビュー状態 / Preview state
    var previewState = { isUndo: false };
    var lastResult = { count: 0, label: "" };

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    var dlg = new Window("dialog", L(LABELS.dialog.title) + " " + SCRIPT_VERSION);

    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    dlg.add("statictext", undefined, L(LABELS.dialog.prompt));

    var alignLabels = [];
    for (var a = 0; a < alignList.length; a++) {
        alignLabels.push(L(alignList[a].label));
    }

    var dropdown = dlg.add("dropdownlist", undefined, alignLabels);
    dropdown.selection = 0;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    var cancelBtn = btnGroup.add("button", undefined, L(LABELS.button.cancel), { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, "OK", { name: "ok" });

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* ScriptUI の値から選択テキストへ文字揃えを適用する純粋関数 / Pure function: apply alignment to selected text from ScriptUI values */
    function process() {
        var selectedIndex = dropdown.selection.index;
        var selectedValue = alignList[selectedIndex].value;
        var count = 0;

        for (var i = 0; i < targetItems.length; i++) {
            count += applyStyleRunAlignment(targetItems[i], selectedValue);
        }

        lastResult.count = count;
        lastResult.label = L(alignList[selectedIndex].label);
    }

    /* プレビュー更新：前回分を取り消してから再適用し再描画 / Refresh preview: undo previous, re-apply, redraw */
    function runPreview() {
        if (previewState.isUndo) {
            app.undo();
        }
        process();
        app.redraw();
        previewState.isUndo = true;
    }

    /* プレビューを取り消して元の状態へ戻す / Undo the preview and restore the original state */
    function undoPreview() {
        if (previewState.isUndo) {
            app.undo();
            app.redraw();
            previewState.isUndo = false;
        }
    }

    /* キャンセルや Esc で閉じたときの後始末 / Clean up when closed via cancel or Esc */
    function cleanupPreview() {
        undoPreview();
    }

    // ドロップダウン変更で随時プレビュー更新 / Update preview whenever the dropdown changes
    dropdown.onChange = function () {
        runPreview();
    };

    // 初回プレビューは onShow から起動（同期コンテキストでの app.undo を避ける）/ Fire the first preview from onShow to avoid app.undo in the sync context
    dlg.onShow = function () {
        runPreview();
    };

    // OK: 直前のプレビューを取り消してから 1 回だけ確定適用（履歴は最後の 1 エントリのみ）/ OK: undo the preview, then apply once so only the final entry remains in history
    okBtn.onClick = function () {
        undoPreview();
        process();
        dlg.close(1);
    };

    cancelBtn.onClick = function () {
        dlg.close(0);
    };

    // Cancel / Esc / X の取りこぼし対策（OK 時は isUndo=false のため何もしない）/ Catch-all for cancel/Esc/X (no-op on OK since isUndo is false)
    dlg.onClose = function () {
        cleanupPreview();
    };

    if (dlg.show() !== 1) {
        return;
    }

    if (lastResult.count === 0) {
        alert(L(LABELS.alert.noText));
    } else {
        alert(L(LABELS.alert.applied)
            .replace("{count}", lastResult.count)
            .replace("{label}", lastResult.label));
    }

    // =========================================
    // 文字揃えの適用 / Apply alignment
    // =========================================

    /* テキストオブジェクトへ再帰的に文字揃えを設定（グループは中を走査）/ Recursively set alignment on text objects (descends into groups) */
    function applyStyleRunAlignment(item, alignmentValue) {
        // グループは子要素を再帰処理 / Groups: recurse into child items
        if (item.typename === "GroupItem") {
            var applied = 0;
            for (var j = 0; j < item.pageItems.length; j++) {
                applied += applyStyleRunAlignment(item.pageItems[j], alignmentValue);
            }
            return applied;
        }

        // TextFrame / TextRange で characterAttributes の取得元だけ切り替え / Pick the characterAttributes source per type
        var attrs = null;
        if (item.typename === "TextFrame") {
            attrs = item.textRange.characterAttributes;
        } else if (item.typename === "TextRange") {
            attrs = item.characterAttributes;
        }
        if (attrs === null) {
            return 0;
        }

        try {
            attrs.alignment = alignmentValue;
            return 1;
        } catch (e) {
            return 0;
        }
    }
})();
