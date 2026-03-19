#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();
/*
 * TextLineEditor.jsx
 * 更新日 / Updated: 2026-03-19
 *
 * 概要 / Overview:
 * 選択した単一のテキストフレームを対象に、行単位で並び替え・追加・編集・削除を行う Illustrator 用スクリプトです。
 * 段落改行（\r）と強制改行（\u0003）を行として扱い、リストボックス上で内容を整理してから反映できます。
 * 空行削除にも対応し、選択状態に応じて操作ボタンの有効 / 無効を切り替えます。
 *
 * This Illustrator script lets you reorder, add, edit, and delete lines in a single selected text frame.
 * It treats both paragraph breaks (\r) and forced line breaks (\u0003) as lines, allows you to organize them in a list box,
 * and then writes the result back to the text frame. It also supports removing empty lines and updates button states
 * according to the current selection.
 */

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "行の並び替えと編集",
        en: "Reorder and Edit Lines"
    },
    noDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    selectOneText: {
        ja: "テキストオブジェクトを1つだけ選択してください。",
        en: "Please select exactly one text object."
    },
    selectText: {
        ja: "テキストオブジェクトを選択してください。",
        en: "Please select a text object."
    },
    emptyText: {
        ja: "テキストが空です。",
        en: "The text is empty."
    },
    needMultipleLines: {
        ja: "複数行のテキストを選択してください。",
        en: "Please select multi-line text."
    },
    instruction: {
        ja: "行を選択して順番を変更してください",
        en: "Select a line and change its order."
    },
    up: {
        ja: "上へ",
        en: "Up"
    },
    down: {
        ja: "下へ",
        en: "Down"
    },
    add: {
        ja: "追加",
        en: "Add"
    },
    edit: {
        ja: "編集",
        en: "Edit"
    },
    deleteLabel: {
        ja: "削除",
        en: "Delete"
    },
    removeEmpty: {
        ja: "空行削除",
        en: "Remove Empty Lines"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    promptAdd: {
        ja: "追加する行を入力してください",
        en: "Enter the line to add."
    },
    promptEdit: {
        ja: "行を編集してください",
        en: "Edit the selected line."
    },
    confirmDelete: {
        ja: "選択した行を削除しますか？",
        en: "Delete the selected line?"
    }
};

function L(key) {
    return LABELS[key][lang];
}

/* メイン処理 / Main process */
(function () {
    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }

    if (app.selection.length !== 1) {
        alert(L("selectOneText"));
        return;
    }

    var item = app.selection[0];
    if (!(item.typename === "TextFrame")) {
        alert(L("selectText"));
        return;
    }

    var originalText = item.contents;
    if (!originalText || originalText === "") {
        alert(L("emptyText"));
        return;
    }

    /* 改行コードを統一（段落改行 \r と強制改行 \u0003 の両対応） / Normalize line breaks (supports both paragraph breaks \r and forced line breaks \u0003) */
    var normalized = originalText
        .replace(/\r\n/g, "\r")
        .replace(/\n/g, "\r")
        .replace(/\u0003/g, "\r");

    var lines = normalized.split("\r");

    if (lines.length <= 1) {
        alert(L("needMultipleLines"));
        return;
    }

    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 10;
    win.margins = 16;

    win.add("statictext", undefined, L("instruction"));

    var mainGroup = win.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "fill"];
    mainGroup.spacing = 15;

    var listBox = mainGroup.add("listbox", undefined, [], {
        multiselect: false,
        numberOfColumns: 1,
        showHeaders: false,
        columnTitles: [""]
    });

    /* リストボックスの最小サイズを基準に、内容に応じて自動調整する / Auto-size the list box based on content while keeping a minimum size */
    var minListWidth = 200;
    var minListHeight = 260;
    var maxListWidth = 720;
    var maxListHeight = 520;

    var longestLen = 0;
    for (var i = 0; i < lines.length; i++) {
        if (lines[i].length > longestLen) longestLen = lines[i].length;
    }

    /* ScriptUI では文字幅を正確に測れないため、1文字あたりのおおよその幅で見積もる / Estimate width per character because ScriptUI cannot measure text width accurately */
    var estimatedWidth = Math.max(minListWidth, Math.min(maxListWidth, 40 + longestLen * 9));
    var estimatedHeight = Math.max(minListHeight, Math.min(maxListHeight, 40 + lines.length * 18));

    listBox.preferredSize = [estimatedWidth, estimatedHeight];
    listBox.columnWidths = [estimatedWidth - 24];

    /* ボタンエリア / Button area */
    var buttonArea = mainGroup.add("group");
    buttonArea.orientation = "row";
    buttonArea.alignment = ["center", "fill"];
    buttonArea.alignChildren = ["center", "top"];

    /* ボタン列 / Button column */
    var btnGroup = buttonArea.add("group");
    btnGroup.orientation = "column";
    btnGroup.alignChildren = ["center", "top"];
    btnGroup.spacing = 8;

    var upBtn = btnGroup.add("button", undefined, L("up"));
    var downBtn = btnGroup.add("button", undefined, L("down"));

    /* スペーサー（上下操作と編集操作を分離） / Spacer to separate move operations from edit operations */
    var spacer = btnGroup.add("group");
    spacer.minimumSize.height = 10;

    var addBtn = btnGroup.add("button", undefined, L("add"));
    var editBtn = btnGroup.add("button", undefined, L("edit"));
    var deleteBtn = btnGroup.add("button", undefined, L("deleteLabel"));
    var removeEmptyBtn = btnGroup.add("button", undefined, L("removeEmpty"));

    /* 下部ボタンエリア / Bottom button area */
    var bottomGroup = win.add("group");
    bottomGroup.orientation = "row";
    bottomGroup.alignment = ["center", "fill"];
    bottomGroup.margins = [0, 10, 0, 0];

    var cancelBtn = bottomGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    var okBtn = bottomGroup.add("button", undefined, L("ok"), { name: "ok" });

    /* ボタンの有効 / 無効を更新 / Update button enabled states */
    function updateButtonState() {
        var hasSelection = !!listBox.selection;
        var idx = hasSelection ? listBox.selection.index : -1;

        upBtn.enabled = hasSelection && idx > 0;
        downBtn.enabled = hasSelection && idx >= 0 && idx < lines.length - 1;
        editBtn.enabled = hasSelection;
        deleteBtn.enabled = hasSelection && lines.length > 1;

        var hasEmptyLine = false;
        for (var i = 0; i < lines.length; i++) {
            if (lines[i] === "") {
                hasEmptyLine = true;
                break;
            }
        }
        removeEmptyBtn.enabled = hasEmptyLine;
    }

    /* リスト表示を更新 / Refresh the list display */
    function refreshList(selectIndex) {
        listBox.removeAll();
        listBox.columnWidths = [listBox.preferredSize[0] - 24];
        for (var i = 0; i < lines.length; i++) {
            listBox.add("item", lines[i]);
        }
        if (lines.length > 0) {
            if (selectIndex < 0) selectIndex = 0;
            if (selectIndex >= lines.length) selectIndex = lines.length - 1;
            listBox.selection = selectIndex;
        } else {
            listBox.selection = null;
        }
        updateButtonState();
    }

    /* 選択行を上へ移動 / Move the selected line up */
    function moveUp() {
        if (!listBox.selection) return;
        var idx = listBox.selection.index;
        if (idx <= 0) return;

        var tmp = lines[idx];
        lines[idx] = lines[idx - 1];
        lines[idx - 1] = tmp;

        refreshList(idx - 1);
    }

    /* 選択行を下へ移動 / Move the selected line down */
    function moveDown() {
        if (!listBox.selection) return;
        var idx = listBox.selection.index;
        if (idx >= lines.length - 1) return;

        var tmp = lines[idx];
        lines[idx] = lines[idx + 1];
        lines[idx + 1] = tmp;

        refreshList(idx + 1);
    }

    /* 行を追加 / Add a line */
    function addLine() {
        var result = prompt(L("promptAdd"), "");
        if (result === null) return;
        lines.push(result);
        refreshList(lines.length - 1);
    }

    /* 選択行を編集 / Edit the selected line */
    function editLine() {
        if (!listBox.selection) return;
        var idx = listBox.selection.index;
        var result = prompt(L("promptEdit"), lines[idx]);
        if (result === null) return;
        lines[idx] = result;
        refreshList(idx);
    }

    /* 選択行を削除 / Delete the selected line */
    function deleteLine() {
        if (!listBox.selection) return;
        if (lines.length <= 1) return;
        var idx = listBox.selection.index;
        if (!confirm(L("confirmDelete"))) return;
        lines.splice(idx, 1);
        refreshList(idx);
    }
    /* 空行を削除 / Remove empty lines */
    function removeEmptyLines() {
        var filtered = [];
        for (var i = 0; i < lines.length; i++) {
            if (lines[i] !== "") {
                filtered.push(lines[i]);
            }
        }
        if (filtered.length === 0) return;
        lines = filtered;
        refreshList(0);
    }

    /* ボタンイベントを関連付ける / Bind button events */
    upBtn.onClick = moveUp;
    downBtn.onClick = moveDown;
    addBtn.onClick = addLine;
    editBtn.onClick = editLine;
    deleteBtn.onClick = deleteLine;
    removeEmptyBtn.onClick = removeEmptyLines;

    /* リストボックスイベント / List box events */
    listBox.onChange = function () {
        updateButtonState();
    };

    listBox.onDoubleClick = function () {
        if (!editBtn.enabled) return;
        editLine();
    };

    /* 初期表示を構築 / Build the initial UI state */
    refreshList(0);

    /* ダイアログを表示 / Show the dialog */
    var result = win.show();
    if (result !== 1) {
        return;
    }

    /* 編集結果をテキストフレームへ反映 / Apply the edited result to the text frame */
    item.contents = lines.join("\r");

    /* 完了メッセージ（必要に応じて使用） / Completion message (enable if needed) */
    // alert("並び替えを反映しました。");
})();
