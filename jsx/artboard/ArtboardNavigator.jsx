#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

概要

アクティブドキュメントのアートボードを一覧表示し、表示・並び替え・追加・
リネーム・再配置をまとめて行うスクリプト。

■ 一覧と表示
リスト項目をクリックすると、そのアートボードを画面いっぱいに表示する。
⌘+クリックで複数選択でき、選択したすべてのアートボードが収まるように表示する。
数字キー 1〜9 でも該当アートボードを表示できる。
画面下部の「全体を表示」で全アートボードを一覧表示できる。

■ 並び替え（アートボードパネル上）
移動ボタンで選択アートボードの並び順をまとめて変更でき、
「名前順に並び替え」で全アートボードを名前順に整列する。

■ アートボードの追加と複製
追加位置（現在の次／末尾）を選んで、空のアートボードの追加または
現在のアートボードの複製ができる。既存の行・列レイアウトを保ったまま挿入する。

■ 一括リネーム（すべて）
# / ## トークンを含むパターン文字列で、全アートボードを連番付きで一括リネームできる。

■ 再配置（カンバス上）
全アートボードをグリッド状に再配置できる。「スクエアに近づける」「横方向に配置」
「縦方向に配置」のほか、「アートボード名の行列を参照」では「1-2」等の名前から
行・列を決めて配置する。行数・列数や間隔（定規単位、列間／行間の連動可）も指定でき、
アートワークも一緒に移動する。

■ その他
「ビデオ定規」でビデオ定規の表示／非表示を切り替えられる。
スクリプト実行中はアートボードの枠線を太く表示し、終了時に元の太さへ戻す。

Overview

Lists the artboards of the active document and bundles viewing,
reordering, adding, renaming and repositioning into one palette.

■ List & view
Clicking a row fits that artboard in the window; ⌘-click selects
multiple rows and fits them all at once, and number keys 1-9 fit the
matching artboard. The "Fit All" button frames every artboard.

■ Reorder (Artboards panel)
The move buttons reorder the selected artboards as a group, and "Sort
by Name" sorts every artboard alphabetically.

■ Add & duplicate artboards
Choose the insert position (after the current artboard or at the end)
to add a blank artboard or duplicate the current one, preserving the
existing row/column layout.

■ Batch rename (all)
Renames every artboard at once from a pattern string with # / ## number
tokens.

■ Reposition (canvas)
Rearranges every artboard into a grid — near-square, horizontal,
vertical, or by row-column artboard names such as "1-2". Rows, columns
and spacing (in ruler units, with optional column/row link) can be
specified, and the artwork moves with the artboards.

■ Other
The "Video Ruler" button toggles the video ruler. While the script is
running the artboard border is shown thicker, and its original width is
restored on exit.

*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

/* 現在のUI言語を判定（ja / en）/ Detect the current UI language (ja / en) */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "アートボード一覧", en: "Artboard List" },
    noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    untitled: { ja: "名称未設定", en: "Untitled" },
    headerNumber: { ja: "番号", en: "No." },
    headerName: { ja: "名前", en: "Name" },
    fitAll: { ja: "全体を表示", en: "Fit All" },
    fitAllTip: { ja: "すべてのアートボードが収まるように表示します。", en: "Fit all artboards in the window." },
    videoRuler: { ja: "ビデオ定規", en: "Video Ruler" },
    videoRulerTip: { ja: "ビデオ定規の表示／非表示を切り替えます。", en: "Toggle the video ruler." },
    close: { ja: "閉じる", en: "Close" },
    sortByName: { ja: "名前順に並び替え", en: "Sort by Name" },
    sortByNameTip: { ja: "すべてのアートボードを名前順に並び替えます。", en: "Sort all artboards by name." },
    reorderPanel: { ja: "並び替え（アートボードパネル上）", en: "Reorder (Artboards panel)" },
    moveTopBtn: { ja: "↑ 先頭へ", en: "↑ Top" },
    moveUpBtn: { ja: "↑ 上へ", en: "↑ Up" },
    moveDownBtn: { ja: "↓ 下へ", en: "↓ Down" },
    moveBottomBtn: { ja: "↓ 末尾へ", en: "↓ Bottom" },
    moveUpTip: { ja: "選択を1つ上へ移動", en: "Move selection up" },
    moveDownTip: { ja: "選択を1つ下へ移動", en: "Move selection down" },
    moveTopTip: { ja: "選択を先頭へ移動", en: "Move selection to top" },
    moveBottomTip: { ja: "選択を末尾へ移動", en: "Move selection to bottom" },
    reorderError: { ja: "アートボードの並び替えに失敗しました。", en: "Failed to reorder the artboards." },
    artboardAddPanel: { ja: "アートボードの追加と複製", en: "Add & Duplicate Artboards" },
    artboardAddPanelTip: {
        ja: "空のアートボードを追加、または現在のアートボードを複製します。",
        en: "Add a blank artboard or duplicate the current artboard."
    },
    addArtboard: { ja: "新規（空）", en: "New (Blank)" },
    addArtboardTip: {
        ja: "空のアートボードを追加します。",
        en: "Add a blank artboard."
    },
    duplicateArtboard: { ja: "複製", en: "Duplicate" },
    duplicateArtboardTip: {
        ja: "現在のアートボードとアートワークを複製します。",
        en: "Duplicate the current artboard and its artwork."
    },
    addPositionLabel: { ja: "追加位置", en: "Insert position" },
    addAfterCurrent: { ja: "現在の次", en: "After current" },
    addAfterCurrentTip: {
        ja: "現在のアートボードの直後へ追加します。",
        en: "Insert after the current artboard."
    },
    addAtEnd: { ja: "末尾", en: "At the end" },
    addAtEndTip: {
        ja: "アートボード一覧の末尾へ追加します。",
        en: "Append to the end of the artboard list."
    },
    artboardLimitError: { ja: "アートボードの最大作成可能数を超えています。", en: "The maximum number of artboards would be exceeded." },
    noSpaceError: { ja: "アートボードを作成する十分なスペースがありません。", en: "There is not enough space to create the artboard." },
    repositionPanel: { ja: "再配置（カンバス上）", en: "Rearrange (Canvas)" },
    repositionNone: { ja: "しない", en: "None" },
    rowsAndColumnsPanel: { ja: "行と列", en: "Rows and Columns" },
    rowsAndColumnsPanelTip: {
        ja: "再配置方法と、行数・列数を指定します。",
        en: "Choose the arrangement method and set rows/columns."
    },
    rowsLabel: { ja: "行数", en: "Rows" },
    columnsLabel: { ja: "列数", en: "Columns" },
    rowsInputTip: { ja: "行数を指定して、列方向に流し込みます。", en: "Set the row count and fill by columns." },
    columnsInputTip: { ja: "列数を指定して、行方向に流し込みます。", en: "Set the column count and fill by rows." },
    squareGrid: { ja: "スクエアに近づける", en: "Near Square" },
    horizontalRow: { ja: "横方向に配置", en: "Horizontal Flow" },
    singleColumn: { ja: "縦方向に配置", en: "Vertical Flow" },
    squareGridTip: { ja: "外形が正方形に近くなるグリッドへ再配置", en: "Rearrange into a near-square grid" },
    horizontalRowTip: { ja: "横方向に詰めて配置（入りきらなければ次の行へ）", en: "Fill rows (wraps to the next row when needed)" },
    singleColumnTip: { ja: "縦方向に詰めて配置（入りきらなければ次の列へ）", en: "Fill a column (wraps to the next column when needed)" },
    repositionByName: { ja: "アートボード名の行列を参照", en: "Use row-column names" },
    repositionByNameTip: { ja: "「1-2」など 行-列／接頭辞-番号 形式の名前からグリッド配置", en: "Place into a grid from row-column / prefix-number names (e.g. 1-2)" },
    noNameMatch: { ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）形式のアートボード名が見つかりませんでした。", en: "No artboard names in row-column (e.g. 1-2) or prefix-number (e.g. banner-1) format were found." },
    gapPanel: { ja: "間隔", en: "Spacing" },
    gapPanelTip: {
        ja: "アートボード間の間隔を指定します。",
        en: "Set the spacing between artboards."
    },
    columnGapLabel: { ja: "水平方向", en: "Horizontal" },
    rowGapLabel: { ja: "垂直方向", en: "Vertical" },
    linkGapLabel: { ja: "連動", en: "Link" },
    linkGapTip: { ja: "列間と行間を同じ値で連動", en: "Keep column and row gaps in sync" },
    duplicatePanel: { ja: "未指定／重複時", en: "Unmatched / Duplicates" },
    duplicateRowEnd: { ja: "各行の末尾に配置", en: "Append to row end" },
    duplicateRowEndTip: { ja: "形式に合わない名前や重複した位置は、直近の行の末尾に配置します。", en: "Place unmatched names or duplicate positions at the end of the nearest row." },
    duplicateLastRow: { ja: "最終行の次の行にまとめる", en: "Collect below last row" },
    duplicateLastRowTip: { ja: "形式に合わない名前や重複した位置を、最終行の次の行にまとめます。", en: "Collect unmatched names or duplicate positions in the row after the last matched row." },
    needTwoArtboards: { ja: "アートボードが2つ以上必要です。", en: "At least two artboards are required." },
    renamePanel: { ja: "一括リネーム（すべて）", en: "Batch Rename (all)" },
    renamePanelTip: {
        ja: "すべてのアートボード名を連番付きで一括変更します。",
        en: "Rename all artboards at once using sequential numbers."
    },
    renameStringLabel: { ja: "パターン", en: "Pattern" },
    renamePatternTip: { ja: "# または ## を含む名前のパターンを入力します。", en: "Enter a name pattern that includes # or ##." },
    applyRename: { ja: "リネーム", en: "Rename" },
    applyRenameTip: { ja: "入力したパターンで、すべてのアートボード名を更新します。", en: "Rename all artboards using the entered pattern." },
    renameNeedText: { ja: "パターンを入力してください。", en: "Enter a name pattern." },
    renameNumberTip: { ja: "連番（1, 2, 3 …）を挿入します。", en: "Insert a sequential number token (1, 2, 3 …)." },
    renameNumberPaddedTip: { ja: "2桁ゼロ埋めの連番（01, 02, 03 …）を挿入します。", en: "Insert a zero-padded number token (01, 02, 03 …)." },
    renameHyphenTip: { ja: "ハイフンを挿入します。", en: "Insert a hyphen." },
    renameUnderscoreTip: { ja: "アンダースコアを挿入します。", en: "Insert an underscore." }
};

/* ラベルを現在の言語で取得 / Get a label in the current language */
function L(key) {
    return LABELS[key][lang];
}

/* コロン付きラベルを取得（日本語は全角コロン、英語は半角コロン）/ Get a label with a trailing colon (full-width for JA, half-width for EN) */
function labelText(key) {
    return L(key) + (lang === "ja" ? "：" : ":");
}

// =========================================
// メイン処理 / Main process
// =========================================

(function () {

    // ドキュメントが無ければ終了 / Abort if no document is open
    if (app.documents.length === 0) {
        alert(L('noDocument'));
        return;
    }

    var doc = app.activeDocument;

    // -------------------------------------
    // アートボード操作 / Artboard operations
    // -------------------------------------

    /* 指定インデックスのアートボードを画面いっぱいに表示 / Fit the artboard at the given index in the window */
    function fitArtboardToWindow(artboardIndex) {
        doc.artboards.setActiveArtboardIndex(artboardIndex);
        // 「選択中のアートボードを画面に合わせる」メニューコマンド / Menu command: fit the active artboard
        app.executeMenuCommand("fitin");
        // モーダルダイアログ表示中はビュー変更を再描画で反映 / Redraw so the view change shows while the modal dialog is up
        app.redraw();
    }

    /* 複数アートボードの外接矩形を画面いっぱいに表示 / Fit the bounding box of multiple artboards in the window
       Illustrator には「選択アートボードに合わせる」コマンドが無いため、
       外接矩形を求めてビューのズーム率と中心位置を直接調整する
       Illustrator has no "fit selected artboards" command, so the view's
       zoom and center point are adjusted directly to the bounding box */
    function fitArtboardsToWindow(artboardIndices) {
        if (artboardIndices.length === 0) {
            return;
        }
        // 1つだけならメニューコマンドのほうが正確 / Use the menu command when only one artboard is selected
        if (artboardIndices.length === 1) {
            fitArtboardToWindow(artboardIndices[0]);
            return;
        }
        var artboards = doc.artboards;

        // 選択アートボードの外接矩形を求める / Compute the bounding box of the selected artboards
        var firstRect = artboards[artboardIndices[0]].artboardRect;
        var unionLeft = firstRect[0];
        var unionTop = firstRect[1];
        var unionRight = firstRect[2];
        var unionBottom = firstRect[3];
        for (var i = 1; i < artboardIndices.length; i++) {
            var artboardRect = artboards[artboardIndices[i]].artboardRect;
            if (artboardRect[0] < unionLeft) { unionLeft = artboardRect[0]; }
            if (artboardRect[1] > unionTop) { unionTop = artboardRect[1]; }
            if (artboardRect[2] > unionRight) { unionRight = artboardRect[2]; }
            if (artboardRect[3] < unionBottom) { unionBottom = artboardRect[3]; }
        }

        var targetWidth = unionRight - unionLeft;
        var targetHeight = unionTop - unionBottom;
        if (targetWidth <= 0 || targetHeight <= 0) {
            return;
        }

        // 現在のビューから表示領域とズーム率を取得 / Read the visible area and zoom from the current view
        var view = doc.views[0];
        var visibleBounds = view.bounds; // [左, 上, 右, 下] / [left, top, right, bottom]
        var visibleWidth = visibleBounds[2] - visibleBounds[0];
        var visibleHeight = visibleBounds[1] - visibleBounds[3];

        // 周囲に余白を確保 / Keep some padding around the artboards
        var paddingFactor = 1.1;

        // 縦横どちらも収まるよう、小さいほうのズーム率を採用 / Use the smaller zoom so both width and height fit
        var zoomByWidth = view.zoom * (visibleWidth / (targetWidth * paddingFactor));
        var zoomByHeight = view.zoom * (visibleHeight / (targetHeight * paddingFactor));
        var newZoom = (zoomByWidth < zoomByHeight) ? zoomByWidth : zoomByHeight;

        view.zoom = newZoom;
        view.centerPoint = [
            (unionLeft + unionRight) / 2,
            (unionTop + unionBottom) / 2
        ];
        // モーダルダイアログ表示中はビュー変更を再描画で反映 / Redraw so the view change shows while the modal dialog is up
        app.redraw();
    }

    /* 1つのアートボードを別のインデックスへ移動 / Move one artboard to another index
       Illustrator には並び替えAPIが無いため remove/insert で作り直す
       Illustrator has no reorder API, so the artboard is rebuilt via remove/insert */
    function moveArtboard(fromIndex, toIndex) {
        if (fromIndex === toIndex) {
            return;
        }
        var artboards = doc.artboards;
        var sourceArtboard = artboards[fromIndex];

        // 移動元の状態を退避 / Capture the source artboard state
        var savedName = sourceArtboard.name;
        var savedRect = sourceArtboard.artboardRect;
        var savedShowCenter = sourceArtboard.showCenter;
        var savedShowCrossHairs = sourceArtboard.showCrossHairs;
        var savedShowSafeAreas = sourceArtboard.showSafeAreas;

        artboards.remove(fromIndex);
        artboards.insert(savedRect, toIndex);

        // 作り直したアートボードへ状態を復元 / Restore the state onto the rebuilt artboard
        var movedArtboard = artboards[toIndex];
        movedArtboard.name = savedName;
        movedArtboard.showCenter = savedShowCenter;
        movedArtboard.showCrossHairs = savedShowCrossHairs;
        movedArtboard.showSafeAreas = savedShowSafeAreas;
    }

    /* 目標順どおりにアートボードを並び替え / Reorder artboards to match the target order
       targetOrder[p] には「位置 p に来るべきアートボードの現在のインデックス」が入る
       targetOrder[p] holds the current index of the artboard that belongs at position p */
    function reorderArtboards(targetOrder) {
        var artboardCount = targetOrder.length;

        // currentPositions[元インデックス] = 現在のインデックス / current index of the artboard originally at that index
        var currentPositions = [];
        var i;
        for (i = 0; i < artboardCount; i++) {
            currentPositions[i] = i;
        }

        for (var targetPosition = 0; targetPosition < artboardCount; targetPosition++) {
            var originalIndex = targetOrder[targetPosition];
            var fromIndex = currentPositions[originalIndex];
            if (fromIndex === targetPosition) {
                continue;
            }
            moveArtboard(fromIndex, targetPosition);

            // 移動で [targetPosition, fromIndex) のアートボードが1つ後ろへずれる / Items in [targetPosition, fromIndex) shift back by one
            for (i = 0; i < artboardCount; i++) {
                if (currentPositions[i] >= targetPosition && currentPositions[i] < fromIndex) {
                    currentPositions[i]++;
                }
            }
            currentPositions[originalIndex] = targetPosition;
        }
    }

    /* 移動操作後の目標順を計算 / Build the target order for a move operation
       mode: "up"（1つ上）/ "down"（1つ下）/ "top"（先頭）/ "bottom"（末尾）
       選択アートボードはブロックとして移動し、相対順は保持する
       The selected artboards move as a block and keep their relative order */
    function buildTargetOrder(selectedIndices, artboardCount, mode) {
        // 選択インデックスを昇順に複製 / Copy the selected indices and sort them ascending
        var selected = selectedIndices.slice();
        selected.sort(function (a, b) { return a - b; });

        var isSelected = [];
        var i;
        for (i = 0; i < artboardCount; i++) {
            isSelected[i] = false;
        }
        for (i = 0; i < selected.length; i++) {
            isSelected[selected[i]] = true;
        }

        // 各選択アートボードの移動後インデックスを決定 / Decide the new index of each selected artboard
        var newPos = [];
        if (mode === "up") {
            // 上から順に、1つ上へ。先頭や他の選択にぶつかったら止まる / Move up, blocked by the top or another selected row
            for (i = 0; i < selected.length; i++) {
                var lowerBound = (i === 0) ? 0 : newPos[i - 1] + 1;
                var raised = selected[i] - 1;
                newPos[i] = (raised > lowerBound) ? raised : lowerBound;
            }
        } else if (mode === "down") {
            // 下から順に、1つ下へ。末尾や他の選択にぶつかったら止まる / Move down, blocked by the bottom or another selected row
            for (i = selected.length - 1; i >= 0; i--) {
                var upperBound = (i === selected.length - 1) ? artboardCount - 1 : newPos[i + 1] - 1;
                var lowered = selected[i] + 1;
                newPos[i] = (lowered < upperBound) ? lowered : upperBound;
            }
        } else if (mode === "top") {
            for (i = 0; i < selected.length; i++) {
                newPos[i] = i;
            }
        } else { // "bottom"
            for (i = 0; i < selected.length; i++) {
                newPos[i] = artboardCount - selected.length + i;
            }
        }

        // 目標順を組み立て / Assemble the target order
        var order = [];
        for (i = 0; i < artboardCount; i++) {
            order[i] = -1;
        }
        for (i = 0; i < selected.length; i++) {
            order[newPos[i]] = selected[i];
        }
        // 非選択アートボードを現在の順序のまま空きスロットへ / Fill the gaps with non-selected artboards, keeping order
        var slot = 0;
        for (i = 0; i < artboardCount; i++) {
            if (!isSelected[i]) {
                while (order[slot] !== -1) {
                    slot++;
                }
                order[slot] = i;
            }
        }
        return order;
    }

    /* アートボード名を自然順で比較（数字は数値として比較）/ Compare artboard names in natural order (digit runs compared numerically) */
    function compareArtboardName(nameA, nameB) {
        // 数字と非数字のかたまりに分割 / Split each name into runs of digits and non-digits
        var chunksA = String(nameA).toLowerCase().match(/(\d+|\D+)/g) || [];
        var chunksB = String(nameB).toLowerCase().match(/(\d+|\D+)/g) || [];
        var shorter = (chunksA.length < chunksB.length) ? chunksA.length : chunksB.length;

        for (var i = 0; i < shorter; i++) {
            var partA = chunksA[i];
            var partB = chunksB[i];
            if (partA === partB) {
                continue;
            }
            // 両方が数字なら数値として比較 / Compare numerically when both runs are digits
            if (/^\d+$/.test(partA) && /^\d+$/.test(partB)) {
                return parseInt(partA, 10) - parseInt(partB, 10);
            }
            return (partA < partB) ? -1 : 1;
        }
        return chunksA.length - chunksB.length;
    }

    /* すべてのアートボードを名前順に並び替え / Sort every artboard by name
       戻り値は適用した目標順 / Returns the target order that was applied */
    function sortArtboardsByName() {
        var artboards = doc.artboards;
        var artboardCount = artboards.length;

        var names = [];
        var order = [];
        var i;
        for (i = 0; i < artboardCount; i++) {
            names[i] = artboards[i].name;
            order[i] = i;
        }
        // 名前順、同名は元の順序を維持 / Sort by name, keeping the original order for equal names
        order.sort(function (a, b) {
            var diff = compareArtboardName(names[a], names[b]);
            return (diff !== 0) ? diff : (a - b);
        });
        reorderArtboards(order);
        return order;
    }

    /* 並び替え後の選択位置を求める / Map the previous selection to its new positions
       targetOrder[p] = 並び替え前のインデックス / pre-reorder index now at position p */
    function mapSelectionAfterReorder(targetOrder, oldSelectedIndices) {
        var wasSelected = [];
        var i;
        for (i = 0; i < targetOrder.length; i++) {
            wasSelected[i] = false;
        }
        for (i = 0; i < oldSelectedIndices.length; i++) {
            wasSelected[oldSelectedIndices[i]] = true;
        }
        var newSelection = [];
        for (i = 0; i < targetOrder.length; i++) {
            if (wasSelected[targetOrder[i]]) {
                newSelection.push(i);
            }
        }
        return newSelection;
    }

    // -------------------------------------
    // アートボードの再配置 / Artboard repositioning
    // -------------------------------------

    /* Illustrator の最大カンバス範囲を取得（一時レイヤーで原点を測定）/ Get Illustrator's max canvas bounds (origin measured via a temp layer) */
    function getCanvasBounds() {
        var CANVAS_MAX_SIZE = 16383;
        var wasModified = doc.modified; // 計測前の変更フラグを退避 / remember the modified flag before measuring
        var tempLayer = doc.layers.add();
        var tempTextFrame = tempLayer.textFrames.add();
        var left = tempTextFrame.matrix.mValueTX;
        var top = tempTextFrame.matrix.mValueTY;
        tempLayer.remove();
        doc.modified = wasModified; // 一時レイヤー追加で立った変更フラグを元に戻す / restore the modified flag
        // [left, top, right, bottom]
        return [left, top, left + CANVAS_MAX_SIZE, top - CANVAS_MAX_SIZE];
    }

    /* グリッド全体をカンバス中央へ配置する原点を算出 / Compute the origin that centers the whole grid on the canvas */
    function computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, columnGap, rowGap, gridColumns, gridRows) {
        var gridWidth = gridColumns * cellWidth + (gridColumns - 1) * columnGap;
        var gridHeight = gridRows * cellHeight + (gridRows - 1) * rowGap;
        var canvasWidth = canvasBounds[2] - canvasBounds[0];
        var canvasHeight = canvasBounds[1] - canvasBounds[3];
        var leftMargin = Math.round((canvasWidth - gridWidth) / 2);
        var topMargin = Math.round((canvasHeight - gridHeight) / 2);
        return {
            left: canvasBounds[0] + leftMargin,
            top: canvasBounds[1] - topMargin
        };
    }

    /* 元のアートボード矩形と最大セルサイズを取得 / Get original artboard rects and the maximum cell size */
    function getArtboardGridMetrics() {
        var artboardCount = doc.artboards.length;
        var artboardRects = [];
        var maxCellWidth = 0;
        var maxCellHeight = 0;
        for (var i = 0; i < artboardCount; i++) {
            var artboardRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
            artboardRects.push(artboardRect);
            var measuredWidth = artboardRect[2] - artboardRect[0];
            var measuredHeight = artboardRect[1] - artboardRect[3];
            if (measuredWidth > maxCellWidth) maxCellWidth = measuredWidth;
            if (measuredHeight > maxCellHeight) maxCellHeight = measuredHeight;
        }
        return {
            rects: artboardRects,
            cellWidth: maxCellWidth,
            cellHeight: maxCellHeight
        };
    }

    /* 全体の外形が最も正方形に近づく列数を求める / Find the column count whose grid outline is closest to a square */
    function chooseBestColumnCount(itemCount, cellWidth, cellHeight, columnGap, rowGap) {
        var bestColumns = 1, bestScore = null, bestEmptyCells = 0, columnCandidate;
        for (columnCandidate = 1; columnCandidate <= itemCount; columnCandidate++) {
            var rows = Math.ceil(itemCount / columnCandidate);
            var occupiedColumns = (columnCandidate < itemCount) ? columnCandidate : itemCount;
            var gridWidth = occupiedColumns * cellWidth + (occupiedColumns - 1) * columnGap;
            var gridHeight = rows * cellHeight + (rows - 1) * rowGap;
            var aspectRatio = gridWidth / gridHeight;
            // 1 に近いほど正方形に近い / closer to 1 = squarer
            var score = (aspectRatio >= 1) ? aspectRatio : (1 / aspectRatio);
            var emptyCells = rows * columnCandidate - itemCount;
            if (bestScore === null
                || score < bestScore - 0.0001
                || (Math.abs(score - bestScore) <= 0.0001 && emptyCells < bestEmptyCells)) {
                bestScore = score;
                bestEmptyCells = emptyCells;
                bestColumns = columnCandidate;
            }
        }
        return bestColumns;
    }

    /* 横方向に流し込む場合の行・列を決定 / Decide rows and columns for horizontal flow */
    function computeHorizontalGridSize(artboardCount, canvasWidth, cellWidth, columnGap) {
        var maxColumns = Math.floor((canvasWidth + columnGap) / (cellWidth + columnGap));
        if (maxColumns < 1) maxColumns = 1;
        var gridColumns = (artboardCount < maxColumns) ? artboardCount : maxColumns;
        return {
            columns: gridColumns,
            rows: Math.ceil(artboardCount / gridColumns),
            fillByColumn: false
        };
    }

    /* 縦方向に流し込む場合の行・列を決定 / Decide rows and columns for vertical flow */
    function computeVerticalGridSize(artboardCount, canvasHeight, cellHeight, rowGap) {
        var maxRows = Math.floor((canvasHeight + rowGap) / (cellHeight + rowGap));
        if (maxRows < 1) maxRows = 1;
        var gridRows = (artboardCount < maxRows) ? artboardCount : maxRows;
        return {
            columns: Math.ceil(artboardCount / gridRows),
            rows: gridRows,
            fillByColumn: true
        };
    }

    /* 再配置モードから行・列・流し込み方向を決定 / Decide rows, columns and fill direction from the rearrange mode */
    function resolveGridLayout(mode, artboardCount, cellWidth, cellHeight, columnGap, rowGap, canvasWidth, canvasHeight, manualColumns, manualRows, manualFillByColumn) {
        if (typeof manualColumns === "number") {
            return {
                columns: manualColumns,
                rows: manualRows,
                fillByColumn: manualFillByColumn
            };
        }

        if (mode === "horizontal") {
            return computeHorizontalGridSize(artboardCount, canvasWidth, cellWidth, columnGap);
        }

        if (mode === "vertical") {
            return computeVerticalGridSize(artboardCount, canvasHeight, cellHeight, rowGap);
        }

        var squareColumns = chooseBestColumnCount(artboardCount, cellWidth, cellHeight, columnGap, rowGap);
        return {
            columns: squareColumns,
            rows: Math.ceil(artboardCount / squareColumns),
            fillByColumn: false
        };
    }

    /* 全レイヤー（サブレイヤー含む）を再帰的に収集 / Collect all layers recursively, including sublayers */
    function collectAllLayers() {
        var allLayers = [];
        function walkLayers(layers) {
            for (var i = 0; i < layers.length; i++) {
                allLayers.push(layers[i]);
                walkLayers(layers[i].layers);
            }
        }
        walkLayers(doc.layers);
        return allLayers;
    }

    /* レイヤー直下の最上位アイテムのみ収集（グループ内は親と一緒に動く）/ Collect only top-level items (children move with their parent) */
    function collectTopLevelItems() {
        var topLevelItems = [], pageItems = doc.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].parent && pageItems[i].parent.typename === "Layer") {
                topLevelItems.push(pageItems[i]);
            }
        }
        return topLevelItems;
    }

    /* 矩形 [left, top, right, bottom] の中心点 / Center point of a [left, top, right, bottom] rect */
    function getRectCenter(rect) {
        return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    /* ロックを一時解除し、復元用の状態を返す / Temporarily unlock; returns the state needed to restore it */
    function unlockLayersAndItems() {
        var lockState = { layers: [], items: [] }, i;
        var layers = collectAllLayers();
        for (i = 0; i < layers.length; i++) {
            if (layers[i].locked) { lockState.layers.push(layers[i]); layers[i].locked = false; }
        }
        var pageItems = doc.pageItems;
        for (i = 0; i < pageItems.length; i++) {
            if (pageItems[i].locked) { lockState.items.push(pageItems[i]); pageItems[i].locked = false; }
        }
        return lockState;
    }

    /* unlockLayersAndItems で外したロックを元に戻す / Re-apply the locks released by unlockLayersAndItems */
    function restoreLockState(lockState) {
        var i;
        for (i = 0; i < lockState.items.length; i++) { lockState.items[i].locked = true; }
        for (i = 0; i < lockState.layers.length; i++) { lockState.layers[i].locked = true; }
    }

    /* 座標系を元に戻す / Restore the coordinate system */
    function restoreCoordinateSystem(savedCoordinateSystem) {
        if (savedCoordinateSystem !== null) {
            try { app.coordinateSystem = savedCoordinateSystem; } catch (e) { }
        }
    }

    /* アートボード矩形配列から所属アートワークを収集 / Collect artwork assigned to each artboard rect */
    function collectItemsByArtboardRects(artboardRects) {
        var topLevelItems = collectTopLevelItems();
        var itemsByArtboard = [];
        var i, j;
        for (i = 0; i < artboardRects.length; i++) {
            itemsByArtboard.push([]);
        }
        for (i = 0; i < topLevelItems.length; i++) {
            var itemCenter = getRectCenter(topLevelItems[i].geometricBounds);
            for (j = 0; j < artboardRects.length; j++) {
                var candidateRect = artboardRects[j];
                if (itemCenter[0] >= candidateRect[0] && itemCenter[0] <= candidateRect[2]
                    && itemCenter[1] <= candidateRect[1] && itemCenter[1] >= candidateRect[3]) {
                    itemsByArtboard[j].push(topLevelItems[i]);
                    break;
                }
            }
        }
        return itemsByArtboard;
    }

    /* 再配置処理用に座標系とロック状態を一時変更して実行 / Run a reposition operation with temporary coordinate and lock state */
    function runWithUnlockedDocumentCoordinates(operation) {
        var savedCoordinateSystem = null;
        try {
            savedCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }

        var lockState = unlockLayersAndItems();
        try {
            return operation();
        } finally {
            restoreLockState(lockState);
            restoreCoordinateSystem(savedCoordinateSystem);
        }
    }

    /* 再配置結果が見えるよう全体表示 / Fit all artboards after repositioning */
    function fitAllAfterReposition() {
        app.executeMenuCommand("fitall");
        app.redraw();
    }

    /* すべてのアートボードをグリッド状に再配置（アートワークも一緒に移動）/ Rearrange every artboard into a grid (artwork moves with it)
       mode: "square"（外形が正方形に近い）/ "horizontal"（横方向に詰めて折り返し）/ "vertical"（縦方向に詰めて折り返し）
       manualColumns/manualRows/manualFillByColumn を渡すと、その列数・行数・方向で再配置（行/列の手動変更時）
       間隔は列間／行間フィールド、セルは最大アートボードのサイズ。戻り値は成功時 { columns, rows, fillByColumn }、中止時 null
       mode: "square"/"horizontal"/"vertical"; pass manualColumns/Rows/FillByColumn to override.
       Returns { columns, rows, fillByColumn } on success, or null when aborted. */
    function arrangeArtboards(mode, manualColumns, manualRows, manualFillByColumn) {
        var artboardCount = doc.artboards.length;
        if (artboardCount < 2) {
            alert(L('needTwoArtboards'));
            return null;
        }

        // 元の矩形を保存し、最大サイズを 1 セルとする / Capture original rects; cell = largest artboard
        var gridMetrics = getArtboardGridMetrics();
        var originalArtboardRects = gridMetrics.rects;
        var cellWidth = gridMetrics.cellWidth;
        var cellHeight = gridMetrics.cellHeight;
        var i, j;

        // 間隔は列間／行間フィールドの値（pt）/ Gaps come from the column/row gap fields (points)
        var gaps = readGapInputs();
        var columnGap = gaps.columnGap, rowGap = gaps.rowGap;

        // カンバス範囲（中央寄せと折り返し判定に使用）/ Canvas bounds (used for centering and wrap limits)
        var canvasBounds = getCanvasBounds();
        var canvasWidth = canvasBounds[2] - canvasBounds[0];
        var canvasHeight = canvasBounds[1] - canvasBounds[3];

        // 列数・行数・流し込み方向を決定 / Decide columns, rows and fill order
        var gridLayout = resolveGridLayout(
            mode, artboardCount, cellWidth, cellHeight, columnGap, rowGap,
            canvasWidth, canvasHeight, manualColumns, manualRows, manualFillByColumn
        );
        var gridColumns = gridLayout.columns;
        var gridRows = gridLayout.rows;
        var fillByColumn = gridLayout.fillByColumn;

        // アートワークの所属アートボードを中心点で判定（移動前に）/ Assign artwork to artboards by center point (before any move)
        var itemsByArtboard = collectItemsByArtboardRects(originalArtboardRects);

        // グリッド原点（カンバス中央）/ Grid origin centered on the canvas
        var gridOrigin = computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, columnGap, rowGap, gridColumns, gridRows);

        runWithUnlockedDocumentCoordinates(function () {
            for (i = 0; i < artboardCount; i++) {
                // 流し込み方向に応じて行・列を決定 / Row and column depend on the fill order
                var row, column;
                if (fillByColumn) {
                    column = Math.floor(i / gridRows);
                    row = i % gridRows;
                } else {
                    row = Math.floor(i / gridColumns);
                    column = i % gridColumns;
                }
                var cellLeft = gridOrigin.left + column * (cellWidth + columnGap);
                var cellTop = gridOrigin.top - row * (cellHeight + rowGap);

                var originalRect = originalArtboardRects[i];
                var artboardWidth = originalRect[2] - originalRect[0];
                var artboardHeight = originalRect[1] - originalRect[3];
                // サイズが異なるアートボードはセル内で中央に置く / Center smaller artboards within the cell
                var newLeft = cellLeft + (cellWidth - artboardWidth) / 2;
                var newTop = cellTop - (cellHeight - artboardHeight) / 2;

                var deltaX = newLeft - originalRect[0];
                var deltaY = newTop - originalRect[1];

                doc.artboards[i].artboardRect = [newLeft, newTop, newLeft + artboardWidth, newTop - artboardHeight];

                var artboardItems = itemsByArtboard[i];
                for (j = 0; j < artboardItems.length; j++) {
                    try { artboardItems[j].translate(deltaX, deltaY); } catch (e) { }
                }
            }
        });

        fitAllAfterReposition();
        return { columns: gridColumns, rows: gridRows, fillByColumn: fillByColumn };
    }

    // -------------------------------------
    // アートボード名から再配置 / Reposition from artboard names
    // -------------------------------------

    /* 未指定／重複アートボードの境界で間隔を広げる倍率 / Gap multiplier at the boundary into the exception area */
    var EXCEPTION_BOUNDARY_GAP_MULTIPLIER = 3;

    /* 配置候補の初期オブジェクトを作成 / Create an initial placement candidate */
    function createPlacementItem(artboardIndex, artboardRect) {
        return {
            index: artboardIndex,
            rect: artboardRect,
            width: artboardRect[2] - artboardRect[0],
            height: artboardRect[1] - artboardRect[3],
            matched: false,
            matchType: "",
            prefixName: "",
            rowNumber: 0,
            columnNumber: 0,
            assignedRow: 0,
            assignedColumn: 0
        };
    }

    /* 「行-列」形式（例: 1-2）の名前を判定 / Match a "row-column" name such as 1-2 */
    function applyRowColumnNameMatch(placementItem, artboardName) {
        var rowColumnMatch = artboardName.match(/^(\d+)[-_x](\d+)$/i);
        if (!rowColumnMatch) {
            return false;
        }
        placementItem.rowNumber = parseInt(rowColumnMatch[1], 10);
        placementItem.columnNumber = parseInt(rowColumnMatch[2], 10);
        if (placementItem.rowNumber < 1 || placementItem.columnNumber < 1) {
            return false;
        }
        placementItem.matched = true;
        placementItem.matchType = "rowColumn";
        return true;
    }

    /* 「接頭辞-番号」形式（例: banner-1）の名前を判定 / Match a "prefix-number" name such as banner-1 */
    function applyPrefixNumberNameMatch(placementItem, artboardName, prefixOrder, prefixRowOffsetByName) {
        var prefixNumberMatch = artboardName.match(/^(.+?)[-_x](\d+)$/i);
        if (!prefixNumberMatch || /^\d+$/.test(prefixNumberMatch[1])) {
            return false;
        }
        placementItem.prefixName = prefixNumberMatch[1].toLowerCase();
        placementItem.columnNumber = parseInt(prefixNumberMatch[2], 10);
        if (placementItem.columnNumber < 1) {
            return false;
        }
        placementItem.matched = true;
        placementItem.matchType = "prefixNumber";
        // 接頭辞ごとに 1 行を割り当てる（出現順）/ Each distinct prefix gets its own row, in encounter order
        if (prefixRowOffsetByName[placementItem.prefixName] === undefined) {
            prefixRowOffsetByName[placementItem.prefixName] = prefixOrder.length + 1;
            prefixOrder.push(placementItem.prefixName);
        }
        return true;
    }

    /* 全アートボード名を解析して配置候補を作成 / Parse every artboard name into placement candidates */
    function parseArtboardNamePlacements() {
        var artboards = doc.artboards;
        var maxColumnNumber = 0, maxRowNumber = 0;
        var placementItems = [], hasMatchedArtboard = false;
        var prefixOrder = [], prefixRowOffsetByName = {};

        for (var i = 0; i < artboards.length; i++) {
            var placementItem = createPlacementItem(i, artboards[i].artboardRect);
            var artboardName = artboards[i].name;
            if (applyRowColumnNameMatch(placementItem, artboardName)) {
                if (placementItem.columnNumber > maxColumnNumber) maxColumnNumber = placementItem.columnNumber;
                if (placementItem.rowNumber > maxRowNumber) maxRowNumber = placementItem.rowNumber;
                hasMatchedArtboard = true;
            } else if (applyPrefixNumberNameMatch(placementItem, artboardName, prefixOrder, prefixRowOffsetByName)) {
                if (placementItem.columnNumber > maxColumnNumber) maxColumnNumber = placementItem.columnNumber;
                hasMatchedArtboard = true;
            }
            placementItems.push(placementItem);
        }

        return {
            maxColumnNumber: maxColumnNumber,
            maxRowNumber: maxRowNumber,
            maxMatchedRowNumber: maxRowNumber + prefixOrder.length,
            placementItems: placementItems,
            hasMatchedArtboard: hasMatchedArtboard,
            prefixRowOffsetByName: prefixRowOffsetByName
        };
    }

    /* 指定行の末尾（maxColumnNumber + n）に空きスロットを予約 / Reserve the next empty slot at a row's end */
    function reserveNextColumnAtRowEnd(rowNumber, occupiedSlots, nextExceptionColumnByRow, maxColumnNumber) {
        if (nextExceptionColumnByRow[rowNumber] === undefined) {
            nextExceptionColumnByRow[rowNumber] = maxColumnNumber + 1;
        }
        while (true) {
            var candidateColumn = nextExceptionColumnByRow[rowNumber]++;
            var candidateSlotKey = rowNumber + "," + candidateColumn;
            if (!occupiedSlots[candidateSlotKey]) {
                occupiedSlots[candidateSlotKey] = true;
                return candidateColumn;
            }
        }
    }

    /* 配置候補に行・列スロットを割り当てる / Assign row/column slots to placement candidates
       exceptionMode: "rowEnd"（各行の末尾）/ "lastRow"（最終行の次の行にまとめる）
       exceptionMode: "rowEnd" (end of each row) / "lastRow" (collect in the row after the last matched row) */
    function assignArtboardPlacementSlots(placementContext, exceptionMode) {
        var occupiedSlots = {}, nextExceptionColumnByRow = {};
        var currentRowNumber = 1;
        var exceptionRowNumber = placementContext.maxMatchedRowNumber + 1;
        var lastRowExceptionColumn = 1;
        var placementItems = placementContext.placementItems;

        for (var i = 0; i < placementItems.length; i++) {
            var placementItem = placementItems[i];
            if (placementItem.matched) {
                var targetRow = placementItem.rowNumber;
                if (placementItem.matchType === "prefixNumber") {
                    targetRow = placementContext.maxRowNumber + placementContext.prefixRowOffsetByName[placementItem.prefixName];
                }
                placementItem.assignedRow = targetRow;
                var requestedKey = targetRow + "," + placementItem.columnNumber;
                if (!occupiedSlots[requestedKey]) {
                    occupiedSlots[requestedKey] = true;
                    placementItem.assignedColumn = placementItem.columnNumber;
                } else {
                    // 同じ行-列が重複 → その行の末尾へ / Duplicate slot → push to the row end
                    placementItem.assignedColumn = reserveNextColumnAtRowEnd(targetRow, occupiedSlots, nextExceptionColumnByRow, placementContext.maxColumnNumber);
                }
                currentRowNumber = targetRow;
            } else if (exceptionMode === "lastRow") {
                // 名前が一致しないアートボード → 最終行の次の行にまとめる / Unmatched name → collect in the row after the last matched row
                placementItem.assignedRow = exceptionRowNumber;
                while (true) {
                    var lastRowColumn = lastRowExceptionColumn++;
                    var lastRowKey = exceptionRowNumber + "," + lastRowColumn;
                    if (!occupiedSlots[lastRowKey]) {
                        occupiedSlots[lastRowKey] = true;
                        placementItem.assignedColumn = lastRowColumn;
                        break;
                    }
                }
            } else {
                // 名前が一致しないアートボード → 直近の行の末尾へ / Unmatched name → end of the most recent row
                placementItem.assignedRow = currentRowNumber;
                placementItem.assignedColumn = reserveNextColumnAtRowEnd(currentRowNumber, occupiedSlots, nextExceptionColumnByRow, placementContext.maxColumnNumber);
            }
        }
    }

    /* オブジェクトの数値キーを昇順で取り出す / Collect numeric keys in ascending order */
    function collectSortedNumericKeys(numericKeyedObject) {
        var keys = [];
        for (var key in numericKeyedObject) {
            if (numericKeyedObject.hasOwnProperty(key)) {
                keys.push(parseInt(key, 10));
            }
        }
        keys.sort(function (a, b) { return a - b; });
        return keys;
    }

    /* 累積オフセットを計算（boundaryNumber の手前の間隔は EXCEPTION_BOUNDARY_GAP_MULTIPLIER 倍）/ Cumulative offsets; the gap before boundaryNumber is widened */
    function computeCumulativeOffsets(sortedNumbers, sizeByNumber, gap, boundaryNumber) {
        var offsetByNumber = {}, cumulative = 0;
        for (var i = 0; i < sortedNumbers.length; i++) {
            var currentNumber = sortedNumbers[i];
            if (i > 0) {
                cumulative += (currentNumber === boundaryNumber) ? gap * EXCEPTION_BOUNDARY_GAP_MULTIPLIER : gap;
            }
            offsetByNumber[currentNumber] = cumulative;
            cumulative += sizeByNumber[currentNumber];
        }
        return offsetByNumber;
    }

    /* アートボード名の「行-列」「接頭辞-番号」に従ってグリッド状に再配置（アートワークも一緒に移動）
       間隔は列間／行間フィールド、グリッド原点は (0, 0)。exceptionMode で未指定／重複の扱いを指定
       戻り値は成功時 true、一致名が無ければ false
       Rearrange artboards into a grid from "row-col"/"prefix-number" names (artwork moves with them).
       Returns true on success, false when no name matched. */
    function arrangeArtboardsByName(exceptionMode) {
        var artboardCount = doc.artboards.length;
        if (artboardCount < 2) {
            alert(L('needTwoArtboards'));
            return false;
        }

        var placementContext = parseArtboardNamePlacements();
        if (!placementContext.hasMatchedArtboard) {
            alert(L('noNameMatch'));
            return false;
        }
        assignArtboardPlacementSlots(placementContext, exceptionMode);
        var placementItems = placementContext.placementItems;
        var i, j;

        // 列ごとの最大幅・行ごとの最大高さを集計 / Per-column max width and per-row max height
        var columnWidthByNumber = {}, rowHeightByNumber = {};
        for (i = 0; i < placementItems.length; i++) {
            var metricsItem = placementItems[i];
            if (columnWidthByNumber[metricsItem.assignedColumn] === undefined
                || metricsItem.width > columnWidthByNumber[metricsItem.assignedColumn]) {
                columnWidthByNumber[metricsItem.assignedColumn] = metricsItem.width;
            }
            if (rowHeightByNumber[metricsItem.assignedRow] === undefined
                || metricsItem.height > rowHeightByNumber[metricsItem.assignedRow]) {
                rowHeightByNumber[metricsItem.assignedRow] = metricsItem.height;
            }
        }

        // 間隔は列間／行間フィールドの値（pt）/ Gaps come from the column/row gap fields (points)
        var gaps = readGapInputs();
        var columnOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(columnWidthByNumber), columnWidthByNumber, gaps.columnGap, placementContext.maxColumnNumber + 1);
        var rowOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(rowHeightByNumber), rowHeightByNumber, gaps.rowGap, placementContext.maxMatchedRowNumber + 1);

        // アートワークの所属アートボードを中心点で判定（移動前に）/ Assign artwork to artboards by center point (before any move)
        var placementRects = [];
        for (i = 0; i < placementItems.length; i++) {
            placementRects.push(placementItems[i].rect);
        }
        var itemsByArtboard = collectItemsByArtboardRects(placementRects);

        runWithUnlockedDocumentCoordinates(function () {
            for (i = 0; i < placementItems.length; i++) {
                var placementItem = placementItems[i];
                var originalRect = placementItem.rect;
                var newLeft = columnOffsetByNumber[placementItem.assignedColumn];
                var newTop = -rowOffsetByNumber[placementItem.assignedRow];
                var deltaX = newLeft - originalRect[0];
                var deltaY = newTop - originalRect[1];

                doc.artboards[placementItem.index].artboardRect = [newLeft, newTop, newLeft + placementItem.width, newTop - placementItem.height];

                var artboardItems = itemsByArtboard[i];
                for (j = 0; j < artboardItems.length; j++) {
                    try { artboardItems[j].translate(deltaX, deltaY); } catch (e) { }
                }
            }
        });

        fitAllAfterReposition();
        return true;
    }

    // -------------------------------------
    // 簡易リネーム / Quick rename
    // -------------------------------------

    /* リネーム文字列のトークンを展開（## = 2桁ゼロ埋め、# = 連番）/ Expand rename tokens (## = zero-padded, # = plain sequential number) */
    function expandRenameTokens(pattern, number) {
        var paddedNumber = (number < 10) ? ("0" + number) : String(number);
        // ## を先に置換（# より先でないと # が二重展開される）/ Replace ## first, otherwise # would expand twice
        return pattern.replace(/##/g, paddedNumber).replace(/#/g, String(number));
    }

    /* すべてのアートボードをパターン文字列でリネーム（連番は先頭から 1, 2, 3 …）/ Rename every artboard from the pattern (sequence starts at 1) */
    function renameAllArtboards(pattern) {
        var artboards = doc.artboards;
        for (var i = 0; i < artboards.length; i++) {
            artboards[i].name = expandRenameTokens(pattern, i + 1);
        }
    }

    // -------------------------------------
    // アートボードの追加・複製 / Artboard insertion
    // -------------------------------------

    /* 新規アートボード名に付けるサフィックス / Suffixes appended to the new artboard name */
    var ARTBOARD_BLANK_SUFFIX = "_blank";
    var ARTBOARD_COPY_SUFFIX = "_copy";

    /* 指定矩形内に中心がある最上位アイテムを収集 / Collect top-level items whose center lies within the rect */
    function collectItemsOnArtboard(artboardRect) {
        var itemsOnArtboard = [];
        var topLevelItems = collectTopLevelItems();
        for (var i = 0; i < topLevelItems.length; i++) {
            var itemCenter = getRectCenter(topLevelItems[i].geometricBounds);
            if (itemCenter[0] >= artboardRect[0] && itemCenter[0] <= artboardRect[2]
                && itemCenter[1] <= artboardRect[1] && itemCenter[1] >= artboardRect[3]) {
                itemsOnArtboard.push(topLevelItems[i]);
            }
        }
        return itemsOnArtboard;
    }

    /* Illustrator のバージョンに応じたアートボード上限を取得 / Get the artboard limit for the Illustrator version */
    function getArtboardLimit() {
        return (parseFloat(app.version) >= 22) ? 1000 : 100;
    }

    /* 既存レイアウトからグリッド情報を推定 / Infer grid information from the existing layout */
    function inferArtboardGridInfo(artboards, artboardCount) {
        var baseArtboardRect = artboards[0].artboardRect;
        var artboardWidth = baseArtboardRect[2] - baseArtboardRect[0];
        var artboardHeight = baseArtboardRect[3] - baseArtboardRect[1];
        var artboardSpacing = app.preferences.getRealPreference("plugin/ArtboardRearrange/ArtboardSpacing");
        var gridStep = [artboardWidth + artboardSpacing, artboardHeight - artboardSpacing];
        var columns = 0;
        var primaryAxis = 0;
        var secondaryAxis = 1;

        if (artboardCount >= 2) {
            var secondArtboardRect = artboards[1].artboardRect;
            primaryAxis = (baseArtboardRect[0] === secondArtboardRect[0]) ? 1 : 0;
            secondaryAxis = 1 - primaryAxis;
            gridStep[primaryAxis] = secondArtboardRect[primaryAxis] - baseArtboardRect[primaryAxis];

            for (var i = 2; i < artboardCount; i++) {
                var scanArtboardRect = artboards[i].artboardRect;
                if (baseArtboardRect[secondaryAxis] !== scanArtboardRect[secondaryAxis]) {
                    gridStep[secondaryAxis] = scanArtboardRect[secondaryAxis] - baseArtboardRect[secondaryAxis];
                    columns = i;
                    break;
                }
            }
        }

        return {
            baseRect: baseArtboardRect,
            width: artboardWidth,
            height: artboardHeight,
            step: gridStep,
            columns: columns,
            primaryAxis: primaryAxis,
            secondaryAxis: secondaryAxis
        };
    }

    /* カンバスに収まる列数・行数をグリッド情報から計算 / Compute the number of columns and rows that fit on the canvas */
    function computeAvailableGridSize(gridInfo) {
        var canvasBounds = getCanvasBounds();
        var baseRect = gridInfo.baseRect;
        var gridStep = gridInfo.step;
        var primaryAxis = gridInfo.primaryAxis;
        var secondaryAxis = gridInfo.secondaryAxis;
        var gridUnitRect = [
            baseRect[0] + Math.abs(gridStep[0]),
            baseRect[1] - Math.abs(gridStep[1]),
            baseRect[0],
            baseRect[1]
        ];
        var primaryEdge = (primaryAxis ^ +(gridStep[primaryAxis] < 0)) ? primaryAxis : primaryAxis + 2;
        var secondaryEdge = (secondaryAxis ^ +(gridStep[secondaryAxis] < 0)) ? secondaryAxis : secondaryAxis + 2;
        var columns = gridInfo.columns ||
            Math.abs(Math.floor((canvasBounds[primaryEdge] - gridUnitRect[primaryEdge]) / gridStep[primaryAxis]));
        var rows = Math.abs(Math.floor((canvasBounds[secondaryEdge] - gridUnitRect[secondaryEdge]) / gridStep[secondaryAxis]));
        return {
            columns: columns,
            rows: rows
        };
    }

    /* グリッドインデックスに対応する左上座標を取得 / Get the top-left coordinate for a grid index */
    function getGridPositionAt(gridInfo, columns, gridIndex) {
        var offset = [];
        offset[gridInfo.primaryAxis] = (gridIndex % columns) * gridInfo.step[gridInfo.primaryAxis];
        offset[gridInfo.secondaryAxis] = Math.floor(gridIndex / columns) * gridInfo.step[gridInfo.secondaryAxis];
        return [gridInfo.baseRect[0] + offset[0], gridInfo.baseRect[1] + offset[1]];
    }

    /* 挿入位置以降のアートボードとアートワークを1セル後ろへ移動 / Shift artboards and artwork after the insertion point back by one cell */
    function shiftArtboardsAfterInsert(artboards, insertIndex, artboardCount, gridInfo, columns) {
        for (var shiftIndex = artboardCount - 1; shiftIndex >= insertIndex; shiftIndex--) {
            var shiftRect = artboards[shiftIndex].artboardRect;
            var shiftTarget = getGridPositionAt(gridInfo, columns, shiftIndex + 1);
            var dx = shiftTarget[0] - shiftRect[0];
            var dy = shiftTarget[1] - shiftRect[1];
            var shiftItems = collectItemsOnArtboard(shiftRect);
            for (var j = 0; j < shiftItems.length; j++) {
                try { shiftItems[j].translate(dx, dy); } catch (e) { }
            }
            artboards[shiftIndex].artboardRect = [
                shiftRect[0] + dx, shiftRect[1] + dy,
                shiftRect[2] + dx, shiftRect[3] + dy
            ];
        }
    }

    /* 末尾に追加されたアートボードを指定インデックスへ送る / Move the appended artboard into the requested index */
    function moveAppendedArtboardToIndex(artboards, insertIndex) {
        var lastIndex = artboards.length - 1;
        if (insertIndex >= lastIndex) {
            return;
        }
        var appendedRect = artboards[lastIndex].artboardRect;
        var appendedName = artboards[lastIndex].name;
        for (var k = lastIndex; k > insertIndex; k--) {
            artboards[k].artboardRect = artboards[k - 1].artboardRect;
            artboards[k].name = artboards[k - 1].name;
        }
        artboards[insertIndex].artboardRect = appendedRect;
        artboards[insertIndex].name = appendedName;
    }

    /* 複製元アートボードの内容を挿入先へコピー / Copy the source artboard artwork to the inserted artboard */
    function duplicateArtworkToInsertedArtboard(artboards, sourceIndex, destinationIndex) {
        var sourceRect = artboards[sourceIndex].artboardRect;
        var destinationRect = artboards[destinationIndex].artboardRect;
        var offsetX = destinationRect[0] - sourceRect[0];
        var offsetY = destinationRect[1] - sourceRect[1];
        var itemsToCopy = collectItemsOnArtboard(sourceRect);
        for (var m = 0; m < itemsToCopy.length; m++) {
            try { itemsToCopy[m].duplicate().translate(offsetX, offsetY); } catch (e) { }
        }
    }

    /* 既存の行・列レイアウトを保ったまま新規アートボードを挿入（ロジックは AddArtboardPlus.jsx を移植）
       duplicateMode      : true=現在のアートボードを複製 / false=空のアートボード
       insertAfterCurrent : true=現在のアートボードの次 / false=末尾
       戻り値は挿入したアートボードのインデックス、失敗時は -1
       Insert a new artboard while preserving the existing row/column layout.
       Returns the inserted artboard index, or -1 on failure. */
    function insertArtboardPreservingLayout(duplicateMode, insertAfterCurrent) {
        var artboards = doc.artboards;
        var artboardCount = artboards.length;

        if (artboardCount + 1 > getArtboardLimit()) {
            alert(L('artboardLimitError'));
            return -1;
        }

        var activeIndex = artboards.getActiveArtboardIndex();
        var insertedIndex = -1;

        runWithUnlockedDocumentCoordinates(function () {
            var gridInfo = inferArtboardGridInfo(artboards, artboardCount);
            var availableGridSize = computeAvailableGridSize(gridInfo);
            if (artboardCount + 1 > availableGridSize.columns * availableGridSize.rows) {
                alert(L('noSpaceError'));
                return;
            }

            var insertIndex = insertAfterCurrent ? activeIndex + 1 : artboardCount;
            shiftArtboardsAfterInsert(artboards, insertIndex, artboardCount, gridInfo, availableGridSize.columns);

            var insertPosition = getGridPositionAt(gridInfo, availableGridSize.columns, insertIndex);
            artboards.add([
                insertPosition[0], insertPosition[1],
                insertPosition[0] + gridInfo.width, insertPosition[1] + gridInfo.height
            ]);
            moveAppendedArtboardToIndex(artboards, insertIndex);

            artboards[insertIndex].name = artboards[activeIndex].name +
                (duplicateMode ? ARTBOARD_COPY_SUFFIX : ARTBOARD_BLANK_SUFFIX);

            if (duplicateMode) {
                duplicateArtworkToInsertedArtboard(artboards, activeIndex, insertIndex);
            }

            insertedIndex = insertIndex;
        });

        if (insertedIndex < 0) {
            return -1;
        }

        doc.selection = null;
        artboards.setActiveArtboardIndex(insertedIndex);
        app.redraw();
        return insertedIndex;
    }

    // -------------------------------------
    // ダイアログUI / Dialog UI
    // -------------------------------------

    // パネル共通のマージンと行間 / Shared panel margins and spacing
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* パネルの体裁を共通化（縦並び・左寄せ・幅いっぱい・共通マージン）/ Apply the shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ↑↓キーで数値を増減（Shift=±10 スナップ、Option=±0.1、通常=±1）。変更後は onChange を発火
       Arrow keys nudge a numeric field (Shift = ±10 snap, Option = ±0.1, plain = ±1); fires onChange afterward */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") {
                return;
            }
            var numericValue = Number(editText.text);
            if (isNaN(numericValue)) {
                return;
            }
            var keyboard = ScriptUI.environment.keyboardState;
            var direction = (event.keyName === "Up") ? 1 : -1;

            if (keyboard.shiftKey) {
                // Shift：±10、10 の倍数にスナップ / Shift: ±10, snapped to a multiple of 10
                numericValue = (direction > 0) ? Math.ceil((numericValue + 1) / 10) * 10 : Math.floor((numericValue - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                // Option：±0.1 / Option: ±0.1
                numericValue += direction * 0.1;
            } else {
                // 通常：±1 / Plain: ±1
                numericValue += direction;
            }
            if (numericValue < 0) {
                numericValue = 0;
            }
            numericValue = keyboard.altKey ? (Math.round(numericValue * 10) / 10) : Math.round(numericValue);

            editText.text = numericValue;
            event.preventDefault();
            // 既存の onChange を発火（再配置などに反映）/ Fire onChange so handlers run (re-arrange, etc.)
            editText.notify("onChange");
        });
    }

    var dialog = new Window("dialog", L('dialogTitle') + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];
    dialog.spacing = 10;
    dialog.margins = 15;

    // コンテンツ行（左カラム：操作パネル / 右カラム：アートボード一覧）/ Content row (left: control panels, right: artboard list)
    var contentRow = dialog.add("group");
    contentRow.orientation = "row";
    contentRow.alignChildren = ["left", "top"];
    contentRow.alignment = "fill";
    contentRow.spacing = 10;

    // 左カラム：操作パネルを縦に並べる / Left column: control panels stacked vertically
    var leftColumn = contentRow.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["fill", "top"];
    leftColumn.spacing = 10;

    // 右カラム：アートボード一覧と並び替えパネルを縦に並べる / Right column: artboard list with the reorder panel below it
    var rightColumn = contentRow.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["fill", "top"];
    rightColumn.spacing = 10;

    // アートボード一覧（1列目：番号 / 2列目：名前、⌘+クリックで複数選択可）/ Artboard list (col 1: number, col 2: name; Cmd-click for multi-select)
    var artboardList = rightColumn.add("listbox", undefined, [], {
        multiselect: true,
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: [L('headerNumber'), L('headerName')],
        columnWidths: [40, 244]
    });
    artboardList.preferredSize = [300, 300];

    // 並び替えパネル（右カラム下部、1行目：移動ボタン / 2行目：名前順）/ Reorder panel (below the list; row 1: move buttons, row 2: sort by name)
    var reorderPanel = rightColumn.add("panel", undefined, L('reorderPanel'));
    setupPanel(reorderPanel);

    // 移動ボタン（先頭へ / 上へ / 下へ / 末尾へ）/ Move buttons (top / up / down / bottom)
    var moveButtonRow = reorderPanel.add("group");
    moveButtonRow.orientation = "row";
    moveButtonRow.spacing = 4;

    var moveToTopButton = moveButtonRow.add("button", undefined, L('moveTopBtn'));
    var moveUpButton = moveButtonRow.add("button", undefined, L('moveUpBtn'));
    var moveDownButton = moveButtonRow.add("button", undefined, L('moveDownBtn'));
    var moveToBottomButton = moveButtonRow.add("button", undefined, L('moveBottomBtn'));

    // 先頭へ／末尾へはラベルが1文字長いぶん少し広げる / Top/Bottom buttons are slightly wider for their 1-char-longer labels
    var moveButtonHeight = 22;
    var moveButtonWidth = 56;
    var moveEndButtonWidth = 60;
    moveToTopButton.preferredSize = [moveEndButtonWidth, moveButtonHeight];
    moveUpButton.preferredSize = [moveButtonWidth, moveButtonHeight];
    moveDownButton.preferredSize = [moveButtonWidth, moveButtonHeight];
    moveToBottomButton.preferredSize = [moveEndButtonWidth, moveButtonHeight];

    moveToTopButton.helpTip = L('moveTopTip');
    moveUpButton.helpTip = L('moveUpTip');
    moveDownButton.helpTip = L('moveDownTip');
    moveToBottomButton.helpTip = L('moveBottomTip');

    // アートボード名順ボタンは次の行へ / The sort-by-name button sits on its own row
    var sortByNameButton = reorderPanel.add("button", undefined, L('sortByName'));
    sortByNameButton.preferredSize.height = 22;
    sortByNameButton.helpTip = L('sortByNameTip');

    // アートボードの追加と複製パネル（右カラム最下部）/ Add & duplicate artboards panel (bottom of the right column)
    var artboardAddPanel = rightColumn.add("panel", undefined, L('artboardAddPanel'));
    setupPanel(artboardAddPanel);
    artboardAddPanel.helpTip = L('artboardAddPanelTip');

    // 追加位置（現在の次 / 末尾、横並び）/ Insert position (after current / at the end, in a row)
    var addPositionRow = artboardAddPanel.add("group");
    addPositionRow.orientation = "row";
    addPositionRow.alignChildren = ["left", "center"];
    addPositionRow.spacing = 6;
    addPositionRow.add("statictext", undefined, labelText('addPositionLabel'));

    var addPositionGroup = addPositionRow.add("group");
    addPositionGroup.orientation = "row";
    addPositionGroup.alignChildren = ["left", "center"];
    addPositionGroup.spacing = 8;
    var addAfterCurrentRadio = addPositionGroup.add("radiobutton", undefined, L('addAfterCurrent'));
    var addAtEndRadio = addPositionGroup.add("radiobutton", undefined, L('addAtEnd'));
    addAfterCurrentRadio.value = true;
    addAfterCurrentRadio.helpTip = L('addAfterCurrentTip');
    addAtEndRadio.helpTip = L('addAtEndTip');

    // 追加ボタンを横並び（新規（空）／複製）/ Add buttons in a row (New (Blank) / Duplicate)
    var addButtonRow = artboardAddPanel.add("group");
    addButtonRow.orientation = "row";
    addButtonRow.alignChildren = ["left", "center"];
    addButtonRow.spacing = 6;

    var addArtboardButton = addButtonRow.add("button", undefined, L('addArtboard'));
    addArtboardButton.preferredSize.height = 22;
    addArtboardButton.helpTip = L('addArtboardTip');

    var duplicateArtboardButton = addButtonRow.add("button", undefined, L('duplicateArtboard'));
    duplicateArtboardButton.preferredSize.height = 22;
    duplicateArtboardButton.helpTip = L('duplicateArtboardTip');

    /* 追加位置ラジオに従って新規アートボードを挿入し、一覧と表示を更新
       Insert a new artboard per the position radio, then refresh the list and view */
    function runInsertArtboard(duplicateMode) {
        try {
            var insertedIndex = insertArtboardPreservingLayout(duplicateMode, addAfterCurrentRadio.value);
            if (insertedIndex >= 0) {
                refreshArtboardList([insertedIndex]);
                updateByNameRadioState();
                fitArtboardsToWindow([insertedIndex]);
                fitAllButton.enabled = true;
            }
        } catch (e) {
            alert(e);
        }
    }
    addArtboardButton.onClick = function () { runInsertArtboard(false); };
    duplicateArtboardButton.onClick = function () { runInsertArtboard(true); };

    // 簡易リネームパネル（文字列＋トークンで全アートボードを一括リネーム）/ Quick rename panel (rename all artboards from a token pattern)
    var renamePanel = leftColumn.add("panel", undefined, L('renamePanel'));
    setupPanel(renamePanel);
    renamePanel.helpTip = L('renamePanelTip');

    // 文字列入力欄 / Text field
    var renameInputRow = renamePanel.add("group");
    renameInputRow.orientation = "row";
    renameInputRow.alignChildren = ["left", "center"];
    renameInputRow.add("statictext", undefined, labelText('renameStringLabel'));
    var renameInput = renameInputRow.add("edittext", undefined, "");
    renameInput.preferredSize.width = 180;
    renameInput.helpTip = L('renamePatternTip');

    // トークン挿入ボタンとリネームボタンを3カラムに（左：トークン / 中央：スペーサー / 右：リネーム）
    // Token-insert buttons and the rename button in three columns (left / spacer / right)
    var renameTokenRow = renamePanel.add("group");
    renameTokenRow.orientation = "row";
    renameTokenRow.alignment = "fill";
    renameTokenRow.alignChildren = ["left", "center"];

    // 左：トークン挿入ボタン（- _ # ##）/ Left: token-insert buttons
    var renameTokenGroup = renameTokenRow.add("group");
    renameTokenGroup.alignChildren = ["left", "center"];
    var insertHyphenButton = renameTokenGroup.add("button", undefined, "-");
    var insertUnderscoreButton = renameTokenGroup.add("button", undefined, "_");
    var insertNumberButton = renameTokenGroup.add("button", undefined, "#");
    var insertNumberPaddedButton = renameTokenGroup.add("button", undefined, "##");

    // 中央：スペーサー（伸縮）/ Center: stretchable spacer
    var renameTokenSpacer = renameTokenRow.add("group");
    renameTokenSpacer.alignment = ["fill", "fill"];
    renameTokenSpacer.minimumSize.width = 0;

    // 右：リネーム適用ボタン / Right: the rename Apply button
    var renameApplyGroup = renameTokenRow.add("group");
    renameApplyGroup.alignChildren = ["right", "center"];
    var applyRenameButton = renameApplyGroup.add("button", undefined, L('applyRename'));

    insertHyphenButton.helpTip = L('renameHyphenTip');
    insertUnderscoreButton.helpTip = L('renameUnderscoreTip');
    insertNumberButton.helpTip = L('renameNumberTip');
    insertNumberPaddedButton.helpTip = L('renameNumberPaddedTip');
    applyRenameButton.helpTip = L('applyRenameTip');

    var tokenButtonSize = [28, 22];
    insertHyphenButton.preferredSize = tokenButtonSize;
    insertUnderscoreButton.preferredSize = tokenButtonSize;
    insertNumberButton.preferredSize = tokenButtonSize;
    insertNumberPaddedButton.preferredSize = tokenButtonSize;
    applyRenameButton.preferredSize.height = 22;

    // 再配置パネル（カンバス上のレイアウト）/ Reposition panel (layout on the canvas)
    var repositionPanel = leftColumn.add("panel", undefined, L('repositionPanel'));
    setupPanel(repositionPanel);

    // 行と列パネル（再配置モードのラジオ＋行/列の入力欄）/ Rows-and-columns panel (mode radios + rows/cols fields)
    var rowsColumnsPanel = repositionPanel.add("panel", undefined, L('rowsAndColumnsPanel'));
    setupPanel(rowsColumnsPanel);
    rowsColumnsPanel.helpTip = L('rowsAndColumnsPanelTip');

    // 行数・列数の入力欄（パネル最上部。再配置後に結果を表示、手動変更でカンバスへ反映）
    // Rows/columns fields (top of the panel; show the result, editing re-arranges the canvas)
    var gridSizeRow = rowsColumnsPanel.add("group");
    gridSizeRow.orientation = "row";
    gridSizeRow.alignChildren = ["left", "center"];
    gridSizeRow.spacing = 6;

    gridSizeRow.add("statictext", undefined, labelText('rowsLabel'));
    var rowsInput = gridSizeRow.add("edittext", undefined, "");
    rowsInput.preferredSize.width = 40;
    rowsInput.helpTip = L('rowsInputTip');
    gridSizeRow.add("statictext", undefined, labelText('columnsLabel'));
    var columnsInput = gridSizeRow.add("edittext", undefined, "");
    columnsInput.preferredSize.width = 40;
    columnsInput.helpTip = L('columnsInputTip');

    // ↑↓キーで増減 / Arrow keys nudge the values
    changeValueByArrowKey(rowsInput);
    changeValueByArrowKey(columnsInput);

    // 「しない」状態では行・列は空＋編集不可 / While "None" is selected the fields stay empty and disabled
    rowsInput.enabled = false;
    columnsInput.enabled = false;

    // 再配置モードのラジオボタン（縦並び、「しない」＝デフォルト）/ Reposition mode radio buttons (stacked; "None" = default)
    var repositionRadioGroup = rowsColumnsPanel.add("group");
    repositionRadioGroup.orientation = "column";
    repositionRadioGroup.alignChildren = ["left", "top"];
    repositionRadioGroup.spacing = 4;

    var repositionNoneRadio = repositionRadioGroup.add("radiobutton", undefined, L('repositionNone'));
    var squareGridRadio = repositionRadioGroup.add("radiobutton", undefined, L('squareGrid'));
    squareGridRadio.helpTip = L('squareGridTip');
    var horizontalRadio = repositionRadioGroup.add("radiobutton", undefined, L('horizontalRow'));
    horizontalRadio.helpTip = L('horizontalRowTip');
    var verticalRadio = repositionRadioGroup.add("radiobutton", undefined, L('singleColumn'));
    verticalRadio.helpTip = L('singleColumnTip');
    var byNameRadio = repositionRadioGroup.add("radiobutton", undefined, L('repositionByName'));
    byNameRadio.helpTip = L('repositionByNameTip');

    // 既定は「しない」/ Default selection is "None"
    repositionNoneRadio.value = true;

    /* 環境設定の定規単位（rulerType）を取得 / Get the ruler unit from preferences (rulerType) */
    function getRulerUnitInfo() {
        switch (app.preferences.getIntegerPreference("rulerType")) {
            case 0: return { label: "inch", factor: 72.0 };
            case 1: return { label: "mm", factor: 72.0 / 25.4 };
            case 3: return { label: "pica", factor: 12.0 };
            case 4: return { label: "cm", factor: 72.0 / 2.54 };
            case 5: return { label: "Q", factor: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", factor: 1.0 };
            default: return { label: "pt", factor: 1.0 }; // 2 = pt
        }
    }

    /* 表示用に数値を整える（小数3桁で丸め）/ Round a number for display (3 decimals) */
    function formatGapNumber(value) {
        return String(Math.round(value * 1000) / 1000);
    }

    // ギャップパネル（左：列間・行間 / 右：連動チェック）/ Gap panel (left: gaps, right: link checkbox)
    var gapPanel = repositionPanel.add("panel", undefined, L('gapPanel'));
    setupPanel(gapPanel, 4);
    gapPanel.helpTip = L('gapPanelTip');

    // 列間・行間は定規単位で入力（既定値は 100pt 相当）/ Gaps use the ruler unit (default ≈ 100 pt)
    var gapUnitInfo = getRulerUnitInfo();
    var defaultGapText = formatGapNumber(100 / gapUnitInfo.factor);

    // 2カラム構成 / Two-column layout
    var gapColumns = gapPanel.add("group");
    gapColumns.orientation = "row";
    gapColumns.alignChildren = ["left", "center"];
    gapColumns.spacing = 12;

    // 左カラム：列間・行間の入力欄 / Left column: column/row gap fields
    var gapFieldsColumn = gapColumns.add("group");
    gapFieldsColumn.orientation = "column";
    gapFieldsColumn.alignChildren = ["left", "center"];
    gapFieldsColumn.spacing = 4;

    var columnGapRow = gapFieldsColumn.add("group");
    columnGapRow.orientation = "row";
    columnGapRow.alignChildren = ["left", "center"];
    columnGapRow.spacing = 6;
    columnGapRow.add("statictext", undefined, labelText('columnGapLabel'));
    var columnGapInput = columnGapRow.add("edittext", undefined, defaultGapText);
    columnGapInput.preferredSize.width = 50;
    columnGapRow.add("statictext", undefined, gapUnitInfo.label);

    var rowGapRow = gapFieldsColumn.add("group");
    rowGapRow.orientation = "row";
    rowGapRow.alignChildren = ["left", "center"];
    rowGapRow.spacing = 6;
    rowGapRow.add("statictext", undefined, labelText('rowGapLabel'));
    var rowGapInput = rowGapRow.add("edittext", undefined, defaultGapText);
    rowGapInput.preferredSize.width = 50;
    rowGapRow.add("statictext", undefined, gapUnitInfo.label);

    // 右カラム：連動チェックボックス / Right column: link checkbox
    var gapLinkColumn = gapColumns.add("group");
    gapLinkColumn.orientation = "column";
    gapLinkColumn.alignChildren = ["left", "center"];
    var linkGapCheckbox = gapLinkColumn.add("checkbox", undefined, L('linkGapLabel'));
    linkGapCheckbox.helpTip = L('linkGapTip');

    // ↑↓キーで増減 / Arrow keys nudge the values
    changeValueByArrowKey(columnGapInput);
    changeValueByArrowKey(rowGapInput);

    /* 連動チェック中は、編集した側の値をもう一方へ反映 / While linked, mirror the edited value to the other field */
    function syncLinkedGap(sourceInput, targetInput) {
        if (linkGapCheckbox.value) {
            targetInput.text = sourceInput.text;
        }
    }

    /* 連動状態を UI へ反映（行間欄のディムと値そろえ）/ Apply the link state to the UI (dim the row-gap row, match its value) */
    function applyGapLinkState() {
        rowGapRow.enabled = !linkGapCheckbox.value;
        if (linkGapCheckbox.value) {
            rowGapInput.text = columnGapInput.text;
        }
    }

    // 列間・行間の変更でプレビューを更新（連動時は相手へ反映してから）/ Gap edits refresh the preview (mirroring first when linked)
    columnGapInput.onChange = function () {
        syncLinkedGap(columnGapInput, rowGapInput);
        reapplyArrangement();
    };
    rowGapInput.onChange = function () {
        syncLinkedGap(rowGapInput, columnGapInput);
        reapplyArrangement();
    };

    // 連動 ON/OFF：行間欄のディムを切り替え、プレビューを更新 / Toggle the row-gap dim, then refresh the preview
    linkGapCheckbox.onClick = function () {
        applyGapLinkState();
        reapplyArrangement();
    };

    // 既定で連動 ON（行間欄はディム、値は列間にそろえる）/ Linked by default (row gap dimmed and matched to the column gap)
    linkGapCheckbox.value = true;
    applyGapLinkState();

    /* 列間・行間入力欄を pt 値として読み取る（定規単位 → pt、無効値は 0）
       Read the column/row gap fields as points (ruler unit → points, invalid → 0) */
    function readGapInputs() {
        var columnGap = parseFloat(columnGapInput.text);
        var rowGap = parseFloat(rowGapInput.text);
        if (isNaN(columnGap) || columnGap < 0) { columnGap = 0; }
        if (isNaN(rowGap) || rowGap < 0) { rowGap = 0; }
        return {
            columnGap: columnGap * gapUnitInfo.factor,
            rowGap: rowGap * gapUnitInfo.factor
        };
    }

    // 未指定／重複パネル（名前が一致しないアートボードの扱い／「アートボード名の行列を参照」用）/ Unspecified/duplicate panel
    var duplicatePanel = repositionPanel.add("panel", undefined, L('duplicatePanel'));
    setupPanel(duplicatePanel, 4);
    var duplicateRowEndRadio = duplicatePanel.add("radiobutton", undefined, L('duplicateRowEnd'));
    duplicateRowEndRadio.helpTip = L('duplicateRowEndTip');
    var duplicateLastRowRadio = duplicatePanel.add("radiobutton", undefined, L('duplicateLastRow'));
    duplicateLastRowRadio.helpTip = L('duplicateLastRowTip');
    duplicateRowEndRadio.value = true;

    // ボタンエリア（左：全体表示 / スペーサー / 右：閉じる）/ Button area (left / spacer / right)
    var buttonRowGroup = dialog.add("group");
    buttonRowGroup.orientation = "row";
    buttonRowGroup.margins = [0, 10, 0, 0];
    buttonRowGroup.alignment = ["fill", "bottom"];

    // 左側グループ / Left-side button group
    var buttonLeftGroup = buttonRowGroup.add("group");
    buttonLeftGroup.alignChildren = ["left", "center"];
    var fitAllButton = buttonLeftGroup.add("button", undefined, L('fitAll'));
    fitAllButton.helpTip = L('fitAllTip');
    var videoRulerButton = buttonLeftGroup.add("button", undefined, L('videoRuler'));
    videoRulerButton.helpTip = L('videoRulerTip');

    // スペーサー（伸縮）/ Spacer (stretchable)
    var buttonSpacer = buttonRowGroup.add("group");
    buttonSpacer.alignment = ["fill", "fill"];
    buttonSpacer.minimumSize.width = 0;

    // 右側グループ / Right-side button group
    var buttonRightGroup = buttonRowGroup.add("group");
    buttonRightGroup.alignChildren = ["right", "center"];
    var closeButton = buttonRightGroup.add("button", undefined, L('close'), {
        name: "cancel"
    });

    // -------------------------------------
    // リスト更新とイベント処理 / List refresh and event handlers
    // -------------------------------------

    // 再構築中の onChange を抑制するフラグ / Flag that suppresses onChange while the list is being rebuilt
    var suppressListChange = false;

    /* リストボックスを現在のアートボード順で再構築 / Rebuild the listbox from the current artboard order */
    function refreshArtboardList(selectionIndices) {
        suppressListChange = true;
        artboardList.removeAll();

        var artboards = doc.artboards;
        for (var i = 0; i < artboards.length; i++) {
            var artboardName = artboards[i].name;
            if (!artboardName || artboardName === "") {
                artboardName = L('untitled');
            }
            // 1列目に番号、2列目に名前 / Column 1 = number, column 2 = name
            var listItem = artboardList.add("item", String(i + 1));
            listItem.subItems[0].text = artboardName;
        }
        // 選択を復元 / Restore the selection
        if (selectionIndices) {
            for (var i = 0; i < selectionIndices.length; i++) {
                var selectedIndex = selectionIndices[i];
                if (selectedIndex >= 0 && selectedIndex < artboardList.items.length) {
                    artboardList.items[selectedIndex].selected = true;
                }
            }
        }
        suppressListChange = false;
    }

    /* 選択中アートボードのインデックス配列を取得 / Get the indices of the selected artboards */
    function getSelectedIndices() {
        var indices = [];
        var selection = artboardList.selection;
        if (selection === null) {
            return indices;
        }
        if (selection instanceof Array) {
            for (var i = 0; i < selection.length; i++) {
                indices.push(selection[i].index);
            }
        } else {
            indices.push(selection.index);
        }
        return indices;
    }

    /* リストで指定インデックスの行だけを選択（他の選択は解除）/ Select only the given row, clearing any other selection */
    function selectSingleArtboardRow(index) {
        // 複数選択リストでは selection への代入だけでは旧選択が残るため、各行を明示的に設定
        // On a multi-select listbox a plain assignment leaves the old rows selected, so set each row explicitly
        suppressListChange = true;
        for (var k = 0; k < artboardList.items.length; k++) {
            artboardList.items[k].selected = (k === index);
        }
        suppressListChange = false;
    }

    /* 並び替え処理を実行し、失敗時はメッセージを表示 / Run a reorder action with error handling */
    function runReorder(reorderAction) {
        try {
            reorderAction();
        } catch (e) {
            alert(L('reorderError') + "\n" + e);
        }
    }

    /* 選択アートボードを指定モードでまとめて移動 / Move the selected artboards as a group with the given mode */
    function moveSelectedArtboards(mode) {
        var selected = getSelectedIndices();
        if (selected.length === 0) {
            return;
        }
        runReorder(function () {
            var artboardCount = doc.artboards.length;
            var targetOrder = buildTargetOrder(selected, artboardCount, mode);
            reorderArtboards(targetOrder);
            refreshArtboardList(mapSelectionAfterReorder(targetOrder, selected));
        });
    }

    // 選択したアートボードを画面に表示（複数選択時は全体が収まるように）/ Fit the selected artboards (the whole group when multiple are selected)
    artboardList.onChange = function () {
        if (suppressListChange) {
            return;
        }
        var indices = getSelectedIndices();
        fitArtboardsToWindow(indices);
        // 個別のアートボードを拡大表示したので全体表示ボタンを有効化 / A subset is shown — enable the Fit All button
        if (indices.length >= 1) {
            fitAllButton.enabled = true;
        }
    };

    // 選択を1つ上へ移動 / Move the selection up
    moveUpButton.onClick = function () {
        moveSelectedArtboards("up");
    };

    // 選択を1つ下へ移動 / Move the selection down
    moveDownButton.onClick = function () {
        moveSelectedArtboards("down");
    };

    // 選択を先頭へ移動 / Move the selection to the top
    moveToTopButton.onClick = function () {
        moveSelectedArtboards("top");
    };

    // 選択を末尾へ移動 / Move the selection to the bottom
    moveToBottomButton.onClick = function () {
        moveSelectedArtboards("bottom");
    };

    // アートボード名で並び替え / Sort artboards by name
    sortByNameButton.onClick = function () {
        var selected = getSelectedIndices();
        runReorder(function () {
            var order = sortArtboardsByName();
            refreshArtboardList(mapSelectionAfterReorder(order, selected));
        });
    };

    /* 適用ボタンの有効状態を更新 / Update the Apply button state
       # トークンが無い文字列（テキストのみ／記号だけ）は全アートボードが同名になるため、# を含むときだけ有効
       Without a # token (text only, or just a symbol) every artboard gets the same name, so enable Apply only when # is present */
    function updateApplyRenameState() {
        applyRenameButton.enabled = (renameInput.text.indexOf("#") !== -1);
    }

    // 記号・番号トークンを文字列欄の末尾へ挿入 / Append a symbol/number token to the rename field
    insertHyphenButton.onClick = function () { renameInput.text = renameInput.text + "-"; updateApplyRenameState(); };
    insertUnderscoreButton.onClick = function () { renameInput.text = renameInput.text + "_"; updateApplyRenameState(); };
    insertNumberButton.onClick = function () { renameInput.text = renameInput.text + "#"; updateApplyRenameState(); };
    insertNumberPaddedButton.onClick = function () { renameInput.text = renameInput.text + "##"; updateApplyRenameState(); };

    // 文字列欄の入力に応じて適用ボタンの有効状態を更新 / Keep the Apply button state in sync with the field
    renameInput.onChanging = updateApplyRenameState;

    // 入力した文字列で全アートボードをリネーム / Rename every artboard with the entered pattern
    applyRenameButton.onClick = function () {
        var pattern = renameInput.text;
        if (!pattern) {
            alert(L('renameNeedText'));
            return;
        }
        var selected = getSelectedIndices();
        try {
            renameAllArtboards(pattern);
        } catch (e) {
            alert(e);
            return;
        }
        refreshArtboardList(selected);
        // リネームで名前が変わったので参照オプションの可否を再判定 / Re-check the byName option after renaming
        updateByNameRadioState();
    };

    // すべてのアートボードを画面に収める / Fit all artboards in the window
    fitAllButton.onClick = function () {
        app.executeMenuCommand("fitall");
        // モーダルダイアログ表示中はビュー変更を再描画で反映 / Redraw so the view change shows while the modal dialog is up
        app.redraw();
        // 全体表示中はボタン自身をディム / Dim the button itself while the whole view is shown
        fitAllButton.enabled = false;
    };

    // ビデオ定規の表示／非表示を切り替え / Toggle the video ruler
    videoRulerButton.onClick = function () {
        app.executeMenuCommand("videoruler");
    };

    // 直近の再配置で使った流し込み方向（ギャップ変更時の作り直しに使用）/ Fill order of the last arrange (reused when gaps change)
    var lastFillByColumn = false;

    // 行・列フィールドを再配置結果で更新（result が null なら空＋無効）/ Update the rows/cols fields from an arrange result
    function setGridSizeFields(arrangeResult) {
        rowsInput.text = arrangeResult ? String(arrangeResult.rows) : "";
        columnsInput.text = arrangeResult ? String(arrangeResult.columns) : "";
        rowsInput.enabled = !!arrangeResult;
        columnsInput.enabled = !!arrangeResult;
        if (arrangeResult) {
            lastFillByColumn = arrangeResult.fillByColumn;
        }
    }

    // 再配置ラジオ（正方形 / 水平 / 垂直）/ Reposition radios (square / horizontal / vertical)
    function runArrange(mode) {
        try {
            var arrangeResult = arrangeArtboards(mode);
            if (arrangeResult) {
                // 再配置後は全体表示状態。全体表示ボタンをディムし、行・列を表示 / View shows all — dim Fit All and show rows/cols
                fitAllButton.enabled = false;
                setGridSizeFields(arrangeResult);
            }
        } catch (e) {
            alert(e);
        }
    }

    // 指定の列数・行数でカンバスを再配置（行/列フィールドの変更時）/ Re-arrange to an explicit columns/rows spec
    function applyGridSize(columns, rows, fillByColumn) {
        try {
            var arrangeResult = arrangeArtboards(null, columns, rows, fillByColumn);
            if (arrangeResult) {
                fitAllButton.enabled = false;
                setGridSizeFields(arrangeResult);
            }
        } catch (e) {
            alert(e);
        }
    }

    /* 現在の再配置をギャップ変更後に作り直してプレビューを更新
       Re-run the active arrangement so the preview reflects new gap values */
    function reapplyArrangement() {
        if (byNameRadio.value) {
            runArrangeByName();
            return;
        }
        // スクエア／水平／垂直、または手動グリッド：現在の行・列で作り直す
        // Square / horizontal / vertical, or a manual grid: rebuild from the current rows/cols
        var columns = parseInt(columnsInput.text, 10);
        var rows = parseInt(rowsInput.text, 10);
        if (!isNaN(columns) && columns >= 1 && !isNaN(rows) && rows >= 1) {
            applyGridSize(columns, rows, lastFillByColumn);
        }
        // 「しない」のときは行・列が空なので何もしない / "None": fields are empty, nothing to do
    }

    /* 「アートボード名の行列を参照」以外では未指定／重複パネルをディム表示
       Dim the unspecified/duplicate panel unless "from artboard names" is selected */
    function updateDuplicatePanelState() {
        duplicatePanel.enabled = byNameRadio.value;
    }

    /* 再配置モードが「しない」のときは間隔パネルをディム表示
       Dim the spacing panel while the reposition mode is "None" */
    function updateGapPanelState() {
        gapPanel.enabled = !repositionNoneRadio.value;
    }

    squareGridRadio.onClick = function () { runArrange("square"); updateDuplicatePanelState(); updateGapPanelState(); };
    horizontalRadio.onClick = function () { runArrange("horizontal"); updateDuplicatePanelState(); updateGapPanelState(); };
    verticalRadio.onClick = function () { runArrange("vertical"); updateDuplicatePanelState(); updateGapPanelState(); };
    // アートボード名（行-列／接頭辞-番号）から再配置 / Rearrange from artboard names
    function runArrangeByName() {
        try {
            // 未指定／重複ラジオの選択を渡す / Pass the unspecified/duplicate radio choice
            var exceptionMode = duplicateLastRowRadio.value ? "lastRow" : "rowEnd";
            if (arrangeArtboardsByName(exceptionMode)) {
                fitAllButton.enabled = false;
            }
        } catch (e) {
            alert(e);
        }
        // 名前ベースは行/列の数値指定を持たない / Name-based mode has no rows/cols spec
        setGridSizeFields(null);
    }
    byNameRadio.onClick = function () { runArrangeByName(); updateDuplicatePanelState(); updateGapPanelState(); };

    // 未指定／重複の変更は、byName が選択中のときだけ即再配置 / Changing the unspecified/duplicate option re-arranges only while byName is the active mode
    function reapplyDuplicateMode() {
        if (byNameRadio.value) {
            runArrangeByName();
        }
    }
    duplicateRowEndRadio.onClick = reapplyDuplicateMode;
    duplicateLastRowRadio.onClick = reapplyDuplicateMode;
    // 「しない」を選んだら行・列フィールドをクリア / Selecting "None" clears the rows/cols fields
    repositionNoneRadio.onClick = function () { setGridSizeFields(null); updateDuplicatePanelState(); updateGapPanelState(); };

    /* アートボード名に行・列情報がなければ「アートボード名の行列を参照」をディム表示
       Dim the "from artboard names" option when no artboard name carries row/column info */
    function updateByNameRadioState() {
        byNameRadio.enabled = parseArtboardNamePlacements().hasMatchedArtboard;
    }

    // 行数を変更したら、列方向に流し込んで再配置 / Editing the row count re-arranges column-major
    rowsInput.onChange = function () {
        var rows = parseInt(rowsInput.text, 10);
        if (isNaN(rows) || rows < 1) {
            return;
        }
        var artboardCount = doc.artboards.length;
        if (rows > artboardCount) {
            rows = artboardCount;
        }
        applyGridSize(Math.ceil(artboardCount / rows), rows, true);
    };

    // 列数を変更したら、行方向に流し込んで再配置 / Editing the column count re-arranges row-major
    columnsInput.onChange = function () {
        var columns = parseInt(columnsInput.text, 10);
        if (isNaN(columns) || columns < 1) {
            return;
        }
        var artboardCount = doc.artboards.length;
        if (columns > artboardCount) {
            columns = artboardCount;
        }
        applyGridSize(columns, Math.ceil(artboardCount / columns), false);
    };

    // 数字キー 1〜9 で該当アートボードを表示（10 以上は無視）/ Number keys 1-9 fit the matching artboard (10+ ignored)
    dialog.addEventListener("keydown", function (event) {
        // テキスト入力中の数字はそのまま入力させる / Let digits type normally inside text fields
        if (event.target && event.target.type === "edittext") {
            return;
        }
        var keyName = event.keyName;
        if (!keyName || keyName.length !== 1 || keyName < "1" || keyName > "9") {
            return;
        }
        var artboardIndex = parseInt(keyName, 10) - 1;
        // 存在しない番号（アートボード数を超える）は無視 / Ignore numbers beyond the artboard count
        if (artboardIndex >= doc.artboards.length) {
            return;
        }
        selectSingleArtboardRow(artboardIndex);
        fitArtboardsToWindow([artboardIndex]);
        // 個別のアートボードを拡大表示したので全体表示ボタンを有効化 / A subset is shown — enable the Fit All button
        fitAllButton.enabled = true;
        event.preventDefault();
    });

    // 初期表示：先頭を選択（表示は切り替えない）、適用ボタンは空欄なのでディム / Initial state: select the first row (no fit); Apply starts dimmed (field is empty)
    refreshArtboardList([0]);
    updateApplyRenameState();
    updateByNameRadioState();
    updateDuplicatePanelState();
    updateGapPanelState();

    // 実行中はアートボードの枠線を太く表示し、終了時に元の幅へ戻す
    // Thicken the artboard border while running, then restore the original width on close
    var savedArtboardBorderWidth = null;
    try {
        savedArtboardBorderWidth = app.preferences.getRealPreference("ArtboardBBWidth");
    } catch (e) { }

    try {
        try {
            app.preferences.setRealPreference("ArtboardBBWidth", 4);
            app.redraw();
        } catch (e) { }
        dialog.show();
    } finally {
        // 枠線の幅を元に戻す / Restore the artboard border width
        if (savedArtboardBorderWidth !== null) {
            try {
                app.preferences.setRealPreference("ArtboardBBWidth", savedArtboardBorderWidth);
                app.redraw();
            } catch (e) { }
        }
    }

})();