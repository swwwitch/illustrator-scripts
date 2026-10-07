#target illustrator
#targetengine "SortTextByColumnEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストフレーム内のタブ区切りテキストを、指定した列の値で並び替えます。
数値・文字列の昇順／降順とランダム順に対応し、見出し行は自動で判定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortTextByColumn.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ncc89f822d2d2

### Overview

Sorts the tab-separated text inside a text frame by the values in a chosen column.
Numeric and string sorting (ascending or descending) and random order are supported, and a header row is detected automatically.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortTextByColumn.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SortTextByColumn";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.17";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-07";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortTextByColumn.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortTextByColumn.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ncc89f822d2d2"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 区切り文字 / Separators
    // =========================================

    var LINE_SEPARATOR = "\r"; /* 行（段落）の区切り / line (paragraph) separator */
    var CELL_SEPARATOR = "\t"; /* 列の区切り / column separator */

    // =========================================
    // 列のプレビュー / Column preview
    // =========================================

    var COLUMN_PREVIEW_COUNT = 3;      /* 列のラジオボタンに並べる値の数 / number of values shown on each column radio */
    var COLUMN_PREVIEW_MAX_CHARS = 10; /* 値ごとの最大文字数（超えたら「...」で切る）/ max characters per value, truncated with "..." */

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ローカライズ / Localization
    // =========================================

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    /* UIラベル（表示順：列 → 順序 → 見出し）/ UI labels in display order: column, order, header */
    var LABELS = {
        dialog: {
            title: { ja: "列を基準に並べ替え", en: "Sort by Column" }
        },
        panel: {
            sortColumn: { ja: "基準にする列", en: "Sort By" },
            sortOrder:  { ja: "並べ替えの順序", en: "Order" }
        },
        radio: {
            column:     { ja: "【列%1】", en: "[Column %1] " },
            ascending:  { ja: "昇順", en: "Ascending" },
            descending: { ja: "降順", en: "Descending" },
            random:     { ja: "ランダム", en: "Random" }
        },
        checkbox: {
            header:     { ja: "1行目を見出しとして残す", en: "Keep first row as header" }
        },
        button: {
            ok:         { ja: "OK", en: "OK" },
            cancel:     { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            column: {
                ja: "この列の値を基準に行を並べ替えます。数字を含むセルは数値として比べ、空のセルの行は末尾に置きます。",
                en: "Sorts the rows by the values in this column. Cells containing digits are compared as numbers; rows with an empty cell go last."
            },
            ascending: {
                ja: "小さい順に並べます。数値が先、文字はそのあとに文字コード順（A→Z、あ→ん）で並びます。",
                en: "Sorts from smallest to largest. Numbers come first, then text in character-code order (A to Z)."
            },
            descending: { ja: "昇順の逆順に並べます。", en: "Sorts in the reverse of ascending order." },
            random:     { ja: "列の値と関係なく、行の順序をシャッフルします。", en: "Shuffles the rows regardless of the column values." },
            header:     { ja: "1行目は並べ替えず、見出しとして先頭に残します。", en: "Keeps the first row at the top instead of sorting it." }
        },
        alert: {
            noDocument:      { ja: "ドキュメントを開いてください。", en: "Open a document." },
            selectTextFrame: { ja: "1つのテキストオブジェクトを選択してください。", en: "Select one text object." },
            tooFewLines:     { ja: "並べ替えるには2行以上のテキストが必要です。", en: "The text needs at least two lines to sort." }
        }
    };

    // =========================================
    // 並べ替え / Sorting
    // =========================================

    /**
     * 行から指定した列のセルを取り出す（列が足りない行は空文字）
     * @param {string} line - 行
     * @param {number} columnIndex - 列（0 始まり）
     * @returns {string} セルの文字列
     */
    function getCellText(line, columnIndex) {
        return line.split(CELL_SEPARATOR)[columnIndex] || "";
    }

    /**
     * 空、または空白だけの文字列かを返す
     * @param {string} text - 調べる文字列
     * @returns {boolean} 空白だけなら true
     */
    function isBlankText(text) {
        return text.replace(/\s/g, "") === "";
    }

    /**
     * テキストから最初に現れる数値（整数・小数・カンマ区切り）を取り出す
     * @param {string} cellText - セルの文字列
     * @returns {number|null} 数値。見つからなければ null
     */
    function extractFirstNumber(cellText) {
        /* 数字で始める（「,」だけに一致させない）/ Must start with a digit so a lone "," does not match */
        var numberMatch = cellText.match(/\d[\d,]*(\.\d+)?/);
        if (numberMatch) {
            return parseFloat(numberMatch[0].replace(/,/g, ""));
        }
        return null;
    }

    /**
     * 数値を指定の桁数になるまで先頭を 0 で埋める
     * @param {number|string} value - 値
     * @param {number} width - 桁数
     * @returns {string} 0 で埋めた文字列
     */
    function padWithZeros(value, width) {
        var paddedText = String(value);
        while (paddedText.length < width) paddedText = "0" + paddedText;
        return paddedText;
    }

    /**
     * セルを、引数なしの sort() でそのまま比べられる文字列キーにする。
     * 数値は桁をそろえて文字列より前、文字列は大文字小文字を区別しない
     * @param {string} cellText - セルの文字列
     * @returns {string} 並べ替え用のキー
     */
    function buildSortKey(cellText) {
        var firstNumber = extractFirstNumber(cellText);
        if (firstNumber === null) return "1" + cellText.toLowerCase();
        var numberParts = firstNumber.toFixed(6).split(".");
        return "0" + padWithZeros(numberParts[0], 20) + "." + numberParts[1];
    }

    /**
     * 配列の順序をシャッフルする（Fisher–Yates、直接書き換える）
     * @param {string[]} lines - 行
     * @returns {void}
     */
    function shuffleLines(lines) {
        for (var i = lines.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapLine = lines[i];
            lines[i] = lines[j];
            lines[j] = swapLine;
        }
    }

    /**
     * 指定列と並び順で行を並べ替える。セルに数値があれば数値、無ければ文字列で比べる。
     * 比較関数つきの sort() は遅く並びも狂うので、文字列キー＋引数なしの sort() で並べる
     * @param {string[]} lines - 並べ替える行
     * @param {number} columnIndex - 基準にする列（0 始まり）
     * @param {string} sortOrder - "asc" / "desc" / "random"
     * @returns {string[]} 並べ替えた行
     */
    function generateSortedLines(lines, columnIndex, sortOrder) {
        if (sortOrder === "random") {
            var shuffledLines = lines.slice();
            shuffleLines(shuffledLines);
            return shuffledLines;
        }

        /* キーの後ろに行番号を付け、同じキーは元の順に保つ（降順は反転するので番号も逆に振る）
           Append the line number so equal keys keep their order (numbered backwards for descending, which is reversed) */
        var isDescending = (sortOrder === "desc");
        var sortEntries = [];
        var emptyCellLines = []; /* 列が空の行（「合計」など）は順序に関係なく末尾へ / lines with an empty cell (e.g. totals) go last in either order */
        for (var i = 0; i < lines.length; i++) {
            var cellText = getCellText(lines[i], columnIndex);
            if (isBlankText(cellText)) {
                emptyCellLines.push(lines[i]);
                continue;
            }
            var lineNumber = isDescending ? (lines.length - 1 - i) : i;
            sortEntries.push(buildSortKey(cellText) + "\u0001" + padWithZeros(lineNumber, 8));
        }
        sortEntries.sort();
        if (isDescending) sortEntries.reverse();

        var sortedLines = [];
        for (var j = 0; j < sortEntries.length; j++) {
            var entryLineNumber = parseInt(sortEntries[j].slice(sortEntries[j].lastIndexOf("\u0001") + 1), 10);
            sortedLines.push(lines[isDescending ? (lines.length - 1 - entryLineNumber) : entryLineNumber]);
        }
        return sortedLines.concat(emptyCellLines);
    }

    // =========================================
    // 見出しと既定の列 / Header and default column
    // =========================================

    /**
     * プレビュー用に長い値を切り詰める（ダイアログが横に広がらないように）
     * @param {string} cellText - セルの文字列
     * @returns {string} COLUMN_PREVIEW_MAX_CHARS 文字までの文字列。切ったときは末尾に「...」
     */
    function truncatePreviewText(cellText) {
        if (cellText.length <= COLUMN_PREVIEW_MAX_CHARS) return cellText;
        return cellText.substring(0, COLUMN_PREVIEW_MAX_CHARS) + "...";
    }

    /**
     * 各列の値を先頭から COLUMN_PREVIEW_COUNT 件ずつ集める（列のラジオボタンに表示する。長い値は切り詰める）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @returns {string[][]} 列ごとの値
     */
    function collectColumnPreviews(columnCount, lines) {
        var columnPreviews = [];
        for (var i = 0; i < columnCount; i++) {
            columnPreviews[i] = [];
        }
        for (var j = 0; j < lines.length; j++) {
            var rowCells = lines[j].split(CELL_SEPARATOR);
            for (var k = 0; k < columnCount; k++) {
                if (columnPreviews[k].length < COLUMN_PREVIEW_COUNT && rowCells[k]) {
                    columnPreviews[k].push(truncatePreviewText(rowCells[k]));
                }
            }
        }
        return columnPreviews;
    }

    /**
     * 行の指定した列に数値があるかを返す
     * @param {string} line - 行
     * @param {number} columnIndex - 列（0 始まり）
     * @returns {boolean} 数値があれば true
     */
    function cellHasNumber(line, columnIndex) {
        return extractFirstNumber(getCellText(line, columnIndex)) !== null;
    }

    /**
     * 1行目が見出しらしいかを調べる（1行目は数値でなく、2行目以降に数値がある列があるか）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @returns {boolean} 見出しらしければ true
     */
    function looksLikeHeaderRow(columnCount, lines) {
        for (var i = 0; i < columnCount; i++) {
            if (cellHasNumber(lines[0], i)) continue;
            for (var j = 1; j < lines.length; j++) {
                if (cellHasNumber(lines[j], i)) return true;
            }
        }
        return false;
    }

    /**
     * 最初に選んでおく列を決める（先頭3行がすべて数値の最初の列。無ければ1列目）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @param {number} startRow - 調べ始める行（見出しありなら 1）
     * @returns {number} 列の番号（0 始まり）
     */
    function findDefaultColumn(columnCount, lines, startRow) {
        var endRow = Math.min(startRow + 3, lines.length);
        for (var i = 0; i < columnCount; i++) {
            var isNumericColumn = true;
            for (var j = startRow; j < endRow; j++) {
                if (!cellHasNumber(lines[j], i)) {
                    isNumericColumn = false;
                    break;
                }
            }
            if (isNumericColumn) return i;
        }
        return 0;
    }

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

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ［基準にする列］パネルを作り、列ごとのラジオボタンに先頭の値を並べる
     * @param {Window} sortDialog - 追加先のダイアログ
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @returns {RadioButton[]} 列ごとのラジオボタン
     */
    function addColumnPanel(sortDialog, columnCount, lines) {
        var columnPanel = sortDialog.add("panel", undefined, getLabel("panel.sortColumn"));
        setupPanel(columnPanel, 6);

        var columnPreviews = collectColumnPreviews(columnCount, lines);
        var columnRadios = [];
        for (var i = 0; i < columnCount; i++) {
            var columnLabel = getLabel("radio.column", [i + 1]) + columnPreviews[i].join(", ") + "...";
            var columnRadio = columnPanel.add("radiobutton", undefined, columnLabel);
            columnRadio.helpTip = getLabel("tooltip.column");
            columnRadios.push(columnRadio);
        }
        return columnRadios;
    }

    /**
     * ［並べ替えの順序］パネルを作る（昇順・降順・ランダム。初期値は昇順）
     * @param {Window} sortDialog - 追加先のダイアログ
     * @returns {{ascending: RadioButton, descending: RadioButton, random: RadioButton}} 順序のラジオボタン
     */
    function addOrderPanel(sortDialog) {
        var orderPanel = sortDialog.add("panel", undefined, getLabel("panel.sortOrder"));
        setupPanel(orderPanel, 6);

        var orderRadioGroup = orderPanel.add("group");
        setupRow(orderRadioGroup);
        var orderRadios = {};
        var orderKeys = ["ascending", "descending", "random"];
        for (var i = 0; i < orderKeys.length; i++) {
            var orderRadio = orderRadioGroup.add("radiobutton", undefined, getLabel("radio." + orderKeys[i]));
            orderRadio.helpTip = getLabel("tooltip." + orderKeys[i]);
            orderRadios[orderKeys[i]] = orderRadio;
        }
        orderRadios.ascending.value = true;
        return orderRadios;
    }

    /**
     * 選ばれているラジオボタンの番号を返す
     * @param {RadioButton[]} radioButtons - ラジオボタン
     * @returns {number} 番号（0 始まり）。どれも選ばれていなければ 0
     */
    function getSelectedRadioIndex(radioButtons) {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) return i;
        }
        return 0;
    }

    /**
     * 並び替えの列、順序、見出し行の有無を選ぶダイアログを出す
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @param {boolean} hasHeaderCandidate - 1行目だけタブが無く、見出しと判断済みか
     * @returns {Object|null} { column, order, useHeader }。キャンセル時は null
     */
    function showSortOptionsDialog(columnCount, lines, hasHeaderCandidate) {
        var sortDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(sortDialog);

        var columnRadios = addColumnPanel(sortDialog, columnCount, lines);
        var orderRadios = addOrderPanel(sortDialog);

        var headerCheckbox = sortDialog.add("checkbox", undefined, getLabel("checkbox.header"));
        headerCheckbox.helpTip = getLabel("tooltip.header");
        headerCheckbox.value = hasHeaderCandidate || looksLikeHeaderRow(columnCount, lines);

        /* 見出し行があれば2行目以降で既定の列を探す / skip the header row when choosing the default column */
        columnRadios[findDefaultColumn(columnCount, lines, headerCheckbox.value ? 1 : 0)].value = true;

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(sortDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "OK" });

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(sortDialog, SCRIPT_NAME);
        if (sortDialog.show() !== 1) return null;

        return {
            column: getSelectedRadioIndex(columnRadios),
            order: orderRadios.random.value ? "random" : (orderRadios.descending.value ? "desc" : "asc"),
            useHeader: headerCheckbox.value
        };
    }

    // =========================================
    // テキストフレームの書き換え / Rewriting text frames
    // =========================================

    /**
     * テキストフレームの先頭から指定した文字数を削除する（見出しより前の空行を除く）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} characterCount - 削除する文字数（改行も1文字）
     * @returns {void}
     */
    function removeLeadingCharacters(textFrame, characterCount) {
        /* 後ろから消して番号をずらさない / Remove from the end so the indices do not shift */
        var frameCharacters = textFrame.textRange.characters;
        for (var i = characterCount - 1; i >= 0; i--) {
            frameCharacters[i].remove();
        }
    }

    /**
     * 上下に並んだテキストフレームを、上から順に1つのテキストフレームへまとめる
     * @param {TextFrame[]} frames - まとめるテキストフレーム（上から順）
     * @returns {void}
     */
    function mergeTextFramesVertically(frames) {
        if (frames.length < 2) return;

        /* いったん1行ずつのフレームに分ける（作った順がそのまま上からの順）/ split into one frame per line first, created in top-to-bottom order */
        var lineFrames = [];
        for (var i = 0; i < frames.length; i++) {
            var frameLines = frames[i].contents.split(LINE_SEPARATOR);
            for (var j = 0; j < frameLines.length; j++) {
                if (frameLines[j] !== "") {
                    var lineFrame = frames[i].duplicate();
                    lineFrame.contents = frameLines[j];
                    lineFrames.push(lineFrame);
                }
            }
            frames[i].remove();
        }

        var baseFrame = lineFrames[0];
        for (var k = 1; k < lineFrames.length; k++) {
            baseFrame.paragraphs.add('\n');
            var sourceParagraphs = lineFrames[k].paragraphs;
            for (var p = 0; p < sourceParagraphs.length; p++) {
                sourceParagraphs[p].duplicate(baseFrame);
            }
            lineFrames[k].remove();
        }
    }

    /**
     * 見出し行を残して本文だけを並べ替えた内容に置き換える
     * 複製で見出しと本文のフレームに分け、書き換えたあと1つにまとめ直す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string[]} sortedLines - 並べ替えた本文の行
     * @param {number} headerOffset - 見出し行より前にある文字数（先頭の空行の分）
     * @returns {void}
     */
    function replaceBodyKeepingHeader(textFrame, sortedLines, headerOffset) {
        var headerFrame = textFrame;
        var bodyFrame = textFrame.duplicate();

        /* 先頭の空行を除き、見出しを1段落目にする / drop leading blank lines so the header is the first paragraph */
        removeLeadingCharacters(headerFrame, headerOffset);
        removeLeadingCharacters(bodyFrame, headerOffset);

        /* 見出し：2行目以降を削除 / header: remove every line but the first */
        var headerParagraphs = headerFrame.textRange.paragraphs;
        for (var i = headerParagraphs.length - 1; i >= 1; i--) {
            headerParagraphs[i].remove();
        }

        /* 本文：1行目を削除し、本文の書式で書き換える / body: remove the first line so the new text takes the body formatting */
        var bodyParagraphs = bodyFrame.textRange.paragraphs;
        if (bodyParagraphs.length > 1) {
            bodyParagraphs[0].remove();
        }

        bodyFrame.contents = sortedLines.join(LINE_SEPARATOR);
        mergeTextFramesVertically([headerFrame, bodyFrame]);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * テキストを行に分け、空行や空白だけの行を除く。最初の行より前にある文字数も数える
     * @param {string} frameContents - テキストフレームの contents
     * @returns {{lines: string[], headerOffset: number}} 残した行と、最初の行より前の文字数（改行も1文字）
     */
    function collectTextLines(frameContents) {
        var allLines = frameContents.split(LINE_SEPARATOR);
        var lines = [];
        var headerOffset = 0;
        for (var i = 0; i < allLines.length; i++) {
            if (!isBlankText(allLines[i])) {
                lines.push(allLines[i]);
            } else if (lines.length === 0) {
                headerOffset += allLines[i].length + 1; /* 改行の分を足す / plus the line break */
            }
        }
        return { lines: lines, headerOffset: headerOffset };
    }

    /**
     * 選択したテキストフレームのタブ区切りテキストを、選んだ列で並べ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var docSelection = app.activeDocument.selection;
        if (docSelection.length !== 1 || docSelection[0].typename !== "TextFrame") {
            alert(getLabel("alert.selectTextFrame"));
            return;
        }

        var textFrame = docSelection[0];

        var textLines = collectTextLines(textFrame.contents);
        var lines = textLines.lines;
        if (lines.length < 2) {
            alert(getLabel("alert.tooFewLines"));
            return;
        }

        /* 1行目だけタブが無く2行目がタブ区切りなら、1行目を見出しとみなす / a tab-less first line over tabbed lines is a header */
        var hasHeaderCandidate = (lines[0].indexOf(CELL_SEPARATOR) === -1 && lines[1].indexOf(CELL_SEPARATOR) !== -1);
        var columnCount = lines[hasHeaderCandidate ? 1 : 0].split(CELL_SEPARATOR).length;

        var sortOptions = showSortOptionsDialog(columnCount, lines, hasHeaderCandidate);
        if (sortOptions === null) return;

        var dataLines = sortOptions.useHeader ? lines.slice(1) : lines;
        var sortedLines = generateSortedLines(dataLines, sortOptions.column, sortOptions.order);

        if (sortOptions.useHeader) {
            replaceBodyKeepingHeader(textFrame, sortedLines, textLines.headerOffset);
        } else {
            textFrame.contents = sortedLines.join(LINE_SEPARATOR);
        }
        app.redraw();
    }

    main();

})();
