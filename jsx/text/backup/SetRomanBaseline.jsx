#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

SetRomanBaseline.jsx

- 選択したテキストの文字揃え（ベースライン）をまとめて変更
- ダイアログを閉じずにライブプレビューで結果を確認
- ダイアログ表示時に選択テキストの現在の文字揃えを初期選択に反映
- TextFrame / TextRange / グループ内テキストへ再帰的に適用

### 更新履歴

- v1.0.0 : 初版
- v1.1.0 : ドロップダウンをラジオボタンに変更、選択テキストの現在設定を初期選択に反映

### 紹介記事（note: Japanese only）

https://note.com/dtp_tranist/n/n0f1203d5f8f8

### Overview

- Batch-change the character alignment (baseline) of selected text
- Live preview shows the result without closing the dialog
- On open, reflects the current alignment of the selected text as the initial choice
- Recursively applies to TextFrame / TextRange / text inside groups

### Changelog

- v1.0.0 : Initial release
- v1.1.0 : Replaced the dropdown with radio buttons; reflect the selection's current setting as the initial choice

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.0";

// =========================================
// ローカライズ / Localization
// =========================================
var lang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

var LABELS = {
    dialog: {
        title: { ja: "文字揃えの変更", en: "Change Character Alignment" },
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

    // 文字揃えをラジオボタンで列挙 / List the alignments as radio buttons
    var radioPanel = dlg.add("group");
    radioPanel.orientation = "column";
    radioPanel.alignChildren = ["left", "top"];

    var radioButtons = [];
    for (var a = 0; a < alignList.length; a++) {
        var radio = radioPanel.add("radiobutton", undefined, L(alignList[a].label));
        radio.onClick = runPreview;
        radioButtons.push(radio);
    }

    // 選択テキストの現在の文字揃えを初期選択に反映 / Reflect the selection's current alignment as the initial choice
    radioButtons[getInitialAlignmentIndex()].value = true;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    var cancelBtn = btnGroup.add("button", undefined, L(LABELS.button.cancel), { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, "OK", { name: "ok" });

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* 指定インデックスの文字揃えを対象テキストへ適用 / Apply the alignment at the given index to the target text */
    function process(selectedIndex) {
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
        process(getSelectedAlignmentIndex());
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

    // 初回プレビューは onShow から起動（同期コンテキストでの app.undo を避ける）/ Fire the first preview from onShow to avoid app.undo in the sync context
    dlg.onShow = function () {
        runPreview();
    };

    // OK: 選択を控えてプレビューを取り消すだけ。確定適用はダイアログを閉じた後に行う（alert とのタイミング競合を回避）/ OK: stash the choice and undo the preview only; the final apply happens after the dialog closes (avoids the alert timing conflict)
    var committedIndex = -1;
    okBtn.onClick = function () {
        committedIndex = getSelectedAlignmentIndex();
        undoPreview();
        dlg.close(1);
    };

    cancelBtn.onClick = function () {
        dlg.close(0);
    };

    // Cancel / Esc / X の取りこぼし対策（OK 時は isUndo=false のため何もしない）/ Catch-all for cancel/Esc/X (no-op on OK since isUndo is false)
    dlg.onClose = function () {
        undoPreview();
    };

    if (dlg.show() !== 1) {
        return;
    }

    // ダイアログを閉じた後に 1 回だけ確定適用し、再描画してから結果を通知 / Apply once after the dialog has closed, redraw, then report
    process(committedIndex);
    app.redraw();

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

    /* TextFrame / TextRange から characterAttributes を取得（それ以外は null）/ Get characterAttributes from a TextFrame / TextRange (null otherwise) */
    function getCharacterAttributes(item) {
        if (item.typename === "TextFrame") {
            return item.textRange.characterAttributes;
        }
        if (item.typename === "TextRange") {
            return item.characterAttributes;
        }
        return null;
    }

    /* グループを潜って最初のテキストの characterAttributes を返す / Descend into groups and return the first text's characterAttributes */
    function getFirstCharacterAttributes(item) {
        if (item.typename === "GroupItem") {
            for (var j = 0; j < item.pageItems.length; j++) {
                var found = getFirstCharacterAttributes(item.pageItems[j]);
                if (found !== null) {
                    return found;
                }
            }
            return null;
        }
        return getCharacterAttributes(item);
    }

    /* 選択テキストの現在の文字揃えに一致する alignList のインデックス（無ければ 0）/ Index in alignList matching the selection's current alignment (0 if none) */
    function getInitialAlignmentIndex() {
        for (var i = 0; i < targetItems.length; i++) {
            var attrs = getFirstCharacterAttributes(targetItems[i]);
            if (attrs === null) {
                continue;
            }
            var current = String(attrs.alignment);
            for (var k = 0; k < alignList.length; k++) {
                if (String(alignList[k].value) === current) {
                    return k;
                }
            }
            return 0;
        }
        return 0;
    }

    /* 選択中のラジオボタンの alignList インデックス / Index of the selected radio in alignList */
    function getSelectedAlignmentIndex() {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) {
                return i;
            }
        }
        return 0;
    }

    /* テキストへ再帰的に文字揃えを設定（グループは中を走査）/ Recursively set alignment on text (descends into groups) */
    function applyStyleRunAlignment(item, alignmentValue) {
        if (item.typename === "GroupItem") {
            var applied = 0;
            for (var j = 0; j < item.pageItems.length; j++) {
                applied += applyStyleRunAlignment(item.pageItems[j], alignmentValue);
            }
            return applied;
        }

        var attrs = getCharacterAttributes(item);
        if (attrs === null) {
            return 0;
        }
        attrs.alignment = alignmentValue;
        return 1;
    }
})();
