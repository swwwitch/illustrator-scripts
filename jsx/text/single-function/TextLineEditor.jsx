#target illustrator
#targetengine "TextLineEditorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストオブジェクトの行を、一覧で並べ替え・編集します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextLineEditor.md

### Overview

Reorders and edits the lines of the selected text object from a list.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextLineEditor.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextLineEditor";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextLineEditor.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextLineEditor.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /* メイン処理 / Main process */
    (function () {
        if (app.documents.length === 0) {
            alert(getLabel("noDocument"));
            return;
        }

        if (app.selection.length !== 1) {
            alert(getLabel("selectOneText"));
            return;
        }

        var item = app.selection[0];
        if (!(item.typename === "TextFrame")) {
            alert(getLabel("selectText"));
            return;
        }

        var originalText = item.contents;
        if (!originalText || originalText === "") {
            alert(getLabel("emptyText"));
            return;
        }

        /* 改行コードを統一（段落改行 \r と強制改行 \u0003 の両対応） / Normalize line breaks (supports both paragraph breaks \r and forced line breaks \u0003) */
        var normalized = originalText
            .replace(/\r\n/g, "\r")
            .replace(/\n/g, "\r")
            .replace(/\u0003/g, "\r");

        var lines = normalized.split("\r");

        if (lines.length <= 1) {
            alert(getLabel("needMultipleLines"));
            return;
        }

        var win = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing = 10;
        win.margins = 16;

        win.add("statictext", undefined, getLabel("instruction"));

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

        var upBtn = btnGroup.add("button", undefined, getLabel("up"));
        var downBtn = btnGroup.add("button", undefined, getLabel("down"));

        /* スペーサー（上下操作と編集操作を分離） / Spacer to separate move operations from edit operations */
        var spacer = btnGroup.add("group");
        spacer.minimumSize.height = 10;

        var addBtn = btnGroup.add("button", undefined, getLabel("add"));
        var editBtn = btnGroup.add("button", undefined, getLabel("edit"));
        var deleteBtn = btnGroup.add("button", undefined, getLabel("deleteLabel"));
        var removeEmptyBtn = btnGroup.add("button", undefined, getLabel("removeEmpty"));

        /* 下部ボタンエリア / Bottom button area */
        var buttonRow = addButtonRow(win, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("ok"), { name: "ok" });

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
            var result = prompt(getLabel("promptAdd"), "");
            if (result === null) return;
            lines.push(result);
            refreshList(lines.length - 1);
        }

        /* 選択行を編集 / Edit the selected line */
        function editLine() {
            if (!listBox.selection) return;
            var idx = listBox.selection.index;
            var result = prompt(getLabel("promptEdit"), lines[idx]);
            if (result === null) return;
            lines[idx] = result;
            refreshList(idx);
        }

        /* 選択行を削除 / Delete the selected line */
        function deleteLine() {
            if (!listBox.selection) return;
            if (lines.length <= 1) return;
            var idx = listBox.selection.index;
            if (!confirm(getLabel("confirmDelete"))) return;
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
        prepareDialogWindow(win, SCRIPT_NAME);
        var result = win.show();
        if (result !== 1) {
            return;
        }

        /* 編集結果をテキストフレームへ反映 / Apply the edited result to the text frame */
        item.contents = lines.join("\r");

        /* 完了メッセージ（必要に応じて使用） / Completion message (enable if needed) */
        // alert("並び替えを反映しました。");
    })();

})();
