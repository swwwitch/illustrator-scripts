#target illustrator
#targetengine "TextLineEditorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストオブジェクトの行を、一覧で並べ替え・編集します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextLineEditor.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n21bb9a835075

### Overview

Reorders and edits the lines of the selected text object from a list.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextLineEditor.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextLineEditor";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextLineEditor.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextLineEditor.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n21bb9a835075"; /* 紹介記事 / article URL */

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
        dialog: {
            title: { ja: "行の並び替えと編集", en: "Reorder and Edit Lines" },
            instruction: { ja: "行を選択して、並び替えや編集をしてください。", en: "Select a line to reorder or edit it." }
        },
        button: {
            moveUp: { ja: "上へ", en: "Up" },
            moveDown: { ja: "下へ", en: "Down" },
            addLine: { ja: "追加", en: "Add" },
            editLine: { ja: "編集", en: "Edit" },
            removeLine: { ja: "削除", en: "Delete" },
            removeEmptyLines: { ja: "空行を削除", en: "Remove Empty Lines" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            lineList: { ja: "ダブルクリックで行を編集", en: "Double-click a line to edit it" },
            addLine: { ja: "末尾に行を追加", en: "Add a line at the end" },
            removeEmptyLines: { ja: "空の行をすべて削除", en: "Remove all empty lines" }
        },
        prompt: {
            addLine: { ja: "追加する行を入力してください。", en: "Enter the line to add." },
            editLine: { ja: "行を編集してください。", en: "Edit the selected line." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneText: { ja: "テキストオブジェクトを1つだけ選択してください。", en: "Please select exactly one text object." },
            selectText: { ja: "テキストオブジェクトを選択してください。", en: "Please select a text object." },
            emptyText: { ja: "テキストが空です。", en: "The text is empty." },
            needMultipleLines: { ja: "複数行のテキストを選択してください。", en: "Please select multi-line text." },
            confirmRemoveLine: { ja: "選択した行を削除しますか？", en: "Delete the selected line?" }
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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // =========================================
    // レイアウト / Layout
    // =========================================
    var LIST_MIN_WIDTH = 200;      /* 行一覧の最小幅 / min width of the line list */
    var LIST_MIN_HEIGHT = 260;     /* 行一覧の最小の高さ / min height of the line list */
    var LIST_MAX_WIDTH = 720;      /* 行一覧の最大幅 / max width of the line list */
    var LIST_MAX_HEIGHT = 520;     /* 行一覧の最大の高さ / max height of the line list */
    var LIST_CHAR_WIDTH = 9;       /* 1文字あたりの見積もり幅 / estimated width per character */
    var LIST_ROW_HEIGHT = 18;      /* 1行あたりの見積もりの高さ / estimated height per row */
    var LIST_PADDING = 40;         /* 行一覧の見積もりに足す余白 / padding added to the list estimate */
    var LIST_COLUMN_INSET = 24;    /* 列幅を一覧の幅から差し引く量 / column width inset from the list width */

    /**
     * 選択中のテキストフレームを1つ返す。条件に合わなければ警告を出して null を返す
     * @returns {TextFrame|null} 対象のテキストフレーム
     */
    function getTargetTextFrame() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return null;
        }
        var selectedItems = app.activeDocument.selection;
        if (selectedItems.length !== 1) {
            alert(getLabel("alert.selectOneText"));
            return null;
        }
        if (selectedItems[0].typename !== "TextFrame") {
            alert(getLabel("alert.selectText"));
            return null;
        }
        return selectedItems[0];
    }

    /**
     * テキストを行に分ける（段落改行 \r・強制改行 \u0003・\n・\r\n のいずれでも区切る）
     * @param {string} sourceText - 元のテキスト
     * @returns {string[]} 行の配列
     */
    function splitTextIntoLines(sourceText) {
        return sourceText.replace(/\r\n|\n|\u0003/g, "\r").split("\r");
    }

    /**
     * 行の一覧で並べ替え・編集するダイアログを表示する
     * @param {string[]} initialLines - 元の行
     * @returns {string[]|null} 編集後の行。キャンセルしたときは null
     */
    function showLineEditorDialog(initialLines) {
        var lineTexts = initialLines.slice(0);

        var lineEditorDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        lineEditorDialog.orientation = "column";
        lineEditorDialog.alignChildren = ["fill", "top"];
        lineEditorDialog.spacing = 10;
        lineEditorDialog.margins = 16;

        lineEditorDialog.add("statictext", undefined, getLabel("dialog.instruction"));

        var editorRowGroup = lineEditorDialog.add("group");
        editorRowGroup.orientation = "row";
        editorRowGroup.alignChildren = ["fill", "fill"];
        editorRowGroup.spacing = 15;

        var lineListBox = editorRowGroup.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 1,
            showHeaders: false,
            columnTitles: [""]
        });
        lineListBox.preferredSize = estimateLineListSize(lineTexts);
        lineListBox.helpTip = getLabel("tooltip.lineList");

        /* 行を操作するボタンの列 / Column of line operation buttons */
        var lineButtonGroup = editorRowGroup.add("group");
        lineButtonGroup.orientation = "column";
        lineButtonGroup.alignment = ["center", "top"];
        lineButtonGroup.alignChildren = ["center", "top"];
        lineButtonGroup.spacing = 8;

        var btnMoveUp = lineButtonGroup.add("button", undefined, getLabel("button.moveUp"));
        var btnMoveDown = lineButtonGroup.add("button", undefined, getLabel("button.moveDown"));

        /* 並べ替えと編集のボタンを離す / Separate the reorder buttons from the edit buttons */
        var buttonSeparatorSpacer = lineButtonGroup.add("group");
        buttonSeparatorSpacer.minimumSize.height = 10;

        var btnAddLine = lineButtonGroup.add("button", undefined, getLabel("button.addLine"));
        btnAddLine.helpTip = getLabel("tooltip.addLine");
        var btnEditLine = lineButtonGroup.add("button", undefined, getLabel("button.editLine"));
        var btnRemoveLine = lineButtonGroup.add("button", undefined, getLabel("button.removeLine"));
        var btnRemoveEmptyLines = lineButtonGroup.add("button", undefined, getLabel("button.removeEmptyLines"));
        btnRemoveEmptyLines.helpTip = getLabel("tooltip.removeEmptyLines");

        var buttonRow = addButtonRow(lineEditorDialog, { centered: true });
        buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /* 選択中の行番号（未選択は -1） / Index of the selected line (-1 when none) */
        function getSelectedLineIndex() {
            return lineListBox.selection ? lineListBox.selection.index : -1;
        }

        /* ボタンの有効 / 無効を更新 / Update button enabled states */
        function updateLineButtonStates() {
            var selectedIndex = getSelectedLineIndex();
            var hasSelection = selectedIndex >= 0;
            btnMoveUp.enabled = selectedIndex > 0;
            btnMoveDown.enabled = hasSelection && selectedIndex < lineTexts.length - 1;
            btnEditLine.enabled = hasSelection;
            btnRemoveLine.enabled = hasSelection && lineTexts.length > 1;
            btnRemoveEmptyLines.enabled = hasEmptyLine(lineTexts);
        }

        /* 一覧を作り直して、指定した行を選択する / Rebuild the list and select the given line */
        function refreshLineList(selectIndex) {
            lineListBox.removeAll();
            lineListBox.columnWidths = [lineListBox.preferredSize[0] - LIST_COLUMN_INSET];
            for (var i = 0; i < lineTexts.length; i++) {
                lineListBox.add("item", lineTexts[i]);
            }
            lineListBox.selection = lineTexts.length > 0
                ? Math.max(0, Math.min(selectIndex, lineTexts.length - 1))
                : null;
            updateLineButtonStates();
        }

        /* 選択行を offset 行ぶん動かす（-1 で上、1 で下） / Move the selected line by offset (-1 up, 1 down) */
        function moveSelectedLine(offset) {
            var selectedIndex = getSelectedLineIndex();
            var targetIndex = selectedIndex + offset;
            if (selectedIndex < 0 || targetIndex < 0 || targetIndex >= lineTexts.length) return;
            var movingLine = lineTexts[selectedIndex];
            lineTexts[selectedIndex] = lineTexts[targetIndex];
            lineTexts[targetIndex] = movingLine;
            refreshLineList(targetIndex);
        }

        /* 末尾に行を追加 / Add a line at the end */
        function addLine() {
            var inputText = prompt(getLabel("prompt.addLine"), "");
            if (inputText === null) return;
            lineTexts.push(inputText);
            refreshLineList(lineTexts.length - 1);
        }

        /* 選択行を編集 / Edit the selected line */
        function editSelectedLine() {
            var selectedIndex = getSelectedLineIndex();
            if (selectedIndex < 0) return;
            var inputText = prompt(getLabel("prompt.editLine"), lineTexts[selectedIndex]);
            if (inputText === null) return;
            lineTexts[selectedIndex] = inputText;
            refreshLineList(selectedIndex);
        }

        /* 選択行を削除 / Delete the selected line */
        function removeSelectedLine() {
            var selectedIndex = getSelectedLineIndex();
            if (selectedIndex < 0 || lineTexts.length <= 1) return;
            if (!confirm(getLabel("alert.confirmRemoveLine"))) return;
            lineTexts.splice(selectedIndex, 1);
            refreshLineList(selectedIndex);
        }

        /* 空行をすべて削除 / Remove all empty lines */
        function removeEmptyLines() {
            var nonEmptyLines = [];
            for (var i = 0; i < lineTexts.length; i++) {
                if (lineTexts[i] !== "") nonEmptyLines.push(lineTexts[i]);
            }
            if (nonEmptyLines.length === 0) return;
            lineTexts = nonEmptyLines;
            refreshLineList(0);
        }

        btnMoveUp.onClick = function () { moveSelectedLine(-1); };
        btnMoveDown.onClick = function () { moveSelectedLine(1); };
        btnAddLine.onClick = addLine;
        btnEditLine.onClick = editSelectedLine;
        btnRemoveLine.onClick = removeSelectedLine;
        btnRemoveEmptyLines.onClick = removeEmptyLines;
        lineListBox.onChange = updateLineButtonStates;
        lineListBox.onDoubleClick = editSelectedLine;

        refreshLineList(0);

        prepareDialogWindow(lineEditorDialog, SCRIPT_NAME);
        if (lineEditorDialog.show() !== 1) return null;
        return lineTexts;
    }

    /**
     * 行の一覧の大きさを、行数と最長の行の文字数から見積もる。
     * ScriptUI では文字幅を正確に測れないため、1文字あたりのおおよその幅で見積もる
     * @param {string[]} lineTexts - 行の配列
     * @returns {number[]} [幅, 高さ]
     */
    function estimateLineListSize(lineTexts) {
        var longestLineLength = 0;
        for (var i = 0; i < lineTexts.length; i++) {
            if (lineTexts[i].length > longestLineLength) longestLineLength = lineTexts[i].length;
        }
        var listWidth = Math.max(LIST_MIN_WIDTH, Math.min(LIST_MAX_WIDTH, LIST_PADDING + longestLineLength * LIST_CHAR_WIDTH));
        var listHeight = Math.max(LIST_MIN_HEIGHT, Math.min(LIST_MAX_HEIGHT, LIST_PADDING + lineTexts.length * LIST_ROW_HEIGHT));
        return [listWidth, listHeight];
    }

    /**
     * 空の行が含まれるかを返す
     * @param {string[]} lineTexts - 行の配列
     * @returns {boolean} 空の行があれば true
     */
    function hasEmptyLine(lineTexts) {
        for (var i = 0; i < lineTexts.length; i++) {
            if (lineTexts[i] === "") return true;
        }
        return false;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    function main() {
        var targetFrame = getTargetTextFrame();
        if (!targetFrame) return;

        var originalText = targetFrame.contents;
        if (!originalText) {
            alert(getLabel("alert.emptyText"));
            return;
        }

        var sourceLines = splitTextIntoLines(originalText);
        if (sourceLines.length <= 1) {
            alert(getLabel("alert.needMultipleLines"));
            return;
        }

        var editedLines = showLineEditorDialog(sourceLines);
        if (!editedLines) return;

        /* 編集結果をテキストフレームへ反映 / Apply the edited lines to the text frame */
        targetFrame.contents = editedLines.join("\r");
    }

    main();

})();
