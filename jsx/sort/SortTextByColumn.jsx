#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストフレーム内のタブ区切りテキストを、指定した列の値で並び替えます。
数値・文字列の昇順／降順とランダム順に対応し、見出し行は自動で判定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortTextByColumn.md

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
var SCRIPT_VERSION  = "v1.0.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortTextByColumn.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortTextByColumn.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var LINE_SEPARATOR = "\r"; /* 行（段落）の区切り / line (paragraph) separator */
    var CELL_SEPARATOR = "\t"; /* 列の区切り / column separator */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* UIラベル（表示順：列 → 順序 → 見出し） */
    var LABELS = {
        dialog: {
            title: { ja: "列を基準に並べ替え", en: "Sort by Column" }
        },
        panel: {
            sortColumn: { ja: "ソート対象の列", en: "Sort Target Column" },
            sortOrder:  { ja: "ソート方法", en: "Sort Order" }
        },
        radio: {
            column:     { ja: "【列%1】", en: "[Row %1] " },
            ascending:  { ja: "昇順", en: "Ascending" },
            descending: { ja: "降順", en: "Descending" },
            random:     { ja: "ランダム", en: "Random" }
        },
        checkbox: {
            header:     { ja: "1行目を見出し行として扱う", en: "Treat first row as header" }
        },
        button: {
            ok:         { ja: "OK", en: "OK" },
            cancel:     { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            column:     { ja: "この列の値を基準に、行を並べ替えます。", en: "The rows are sorted by the values in this column." },
            ascending:  { ja: "値の小さい順（あいうえお順）に並べます。", en: "Sorts in ascending order." },
            descending: { ja: "値の大きい順に並べます。", en: "Sorts in descending order." },
            random:     { ja: "値と関係なく、行の順序をシャッフルします。", en: "Shuffles the rows regardless of their values." },
            header:     { ja: "1行目は並べ替えず、見出しとして先頭に残します。", en: "Keeps the first row at the top instead of sorting it." }
        }
    };

    /**
     * ドット区切りのキーから表示言語のラベルを取得する
     * @param {string} labelPath - "panel.sortOrder" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // 並べ替え / Sorting
    // =========================================

    /**
     * テキストから最初に現れる数値（整数・小数・カンマ区切り）を取り出す
     * @param {string} cellText - セルの文字列
     * @returns {number|null} 数値。見つからなければ null
     */
    function extractFirstNumber(cellText) {
        var numberMatch = cellText.match(/[\d,]+(\.\d+)?/);
        if (numberMatch) {
            return parseFloat(numberMatch[0].replace(/,/g, ""));
        }
        return null;
    }

    /**
     * 配列の順序をシャッフルする（Fisher–Yates、直接書き換える）
     * @param {Object[]} lineEntries - { line, key } の配列
     * @returns {void}
     */
    function shuffleLineEntries(lineEntries) {
        for (var i = lineEntries.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapEntry = lineEntries[i];
            lineEntries[i] = lineEntries[j];
            lineEntries[j] = swapEntry;
        }
    }

    /**
     * 数値キーを昇順・降順で比べる
     * @param {Object} a - { line, key } の要素
     * @param {Object} b - { line, key } の要素
     * @param {string} sortOrder - "asc" / "desc"
     * @returns {number} 並べ替え用の差
     */
    function compareNumericKeys(a, b, sortOrder) {
        return sortOrder === "asc" ? (a.key - b.key) : (b.key - a.key);
    }

    /**
     * 文字列キーを昇順・降順で比べる（大文字小文字は区別しない）
     * @param {Object} a - { line, key } の要素
     * @param {Object} b - { line, key } の要素
     * @param {string} sortOrder - "asc" / "desc"
     * @returns {number} -1 / 0 / 1
     */
    function compareStringKeys(a, b, sortOrder) {
        var aText = String(a.key).toLowerCase();
        var bText = String(b.key).toLowerCase();
        if (sortOrder === "asc") {
            return aText < bText ? -1 : aText > bText ? 1 : 0;
        }
        return aText > bText ? -1 : aText < bText ? 1 : 0;
    }

    /**
     * 指定列と並び順で行を並べ替える。セルに数値があれば数値、無ければ文字列で比べる
     * @param {string[]} lines - 並べ替える行
     * @param {number} columnIndex - 基準にする列（0 始まり）
     * @param {string} sortOrder - "asc" / "desc" / "random"
     * @returns {string[]} 並べ替えた行
     */
    function generateSortedLines(lines, columnIndex, sortOrder) {
        var lineEntries = [];
        for (var i = 0; i < lines.length; i++) {
            var cellText = lines[i].split(CELL_SEPARATOR)[columnIndex] || "";
            var firstNumber = extractFirstNumber(cellText);
            lineEntries.push({
                line: lines[i],
                key: firstNumber !== null ? firstNumber : cellText
            });
        }

        if (sortOrder === "random") {
            shuffleLineEntries(lineEntries);
        } else {
            lineEntries.sort(function (a, b) {
                if (typeof a.key === "number" && typeof b.key === "number") {
                    return compareNumericKeys(a, b, sortOrder);
                }
                return compareStringKeys(a, b, sortOrder);
            });
        }

        var sortedLines = [];
        for (var j = 0; j < lineEntries.length; j++) {
            sortedLines.push(lineEntries[j].line);
        }
        return sortedLines;
    }

    // =========================================
    // 見出しと既定の列 / Header and default column
    // =========================================

    /**
     * 各列の値を先頭から3件ずつ集める（列のラジオボタンに表示する）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @returns {string[][]} 列ごとの値
     */
    function collectColumnPreviews(columnCount, lines) {
        var columnPreviews = [];
        for (var i = 0; i < columnCount; i++) {
            columnPreviews[i] = [];
        }
        for (var r = 0; r < lines.length; r++) {
            var cells = lines[r].split(CELL_SEPARATOR);
            for (var c = 0; c < columnCount; c++) {
                if (columnPreviews[c].length < 3 && cells[c]) {
                    columnPreviews[c].push(cells[c]);
                }
            }
        }
        return columnPreviews;
    }

    /**
     * 1行目が見出しらしいかを調べる（1行目は数値でなく、2行目以降に数値がある列があるか）
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @returns {boolean} 見出しらしければ true
     */
    function looksLikeHeaderRow(columnCount, lines) {
        for (var i = 0; i < columnCount; i++) {
            var hasNumericBelow = false;
            for (var r = 1; r < lines.length; r++) {
                if (extractFirstNumber(lines[r].split(CELL_SEPARATOR)[i]) !== null) {
                    hasNumericBelow = true;
                    break;
                }
            }
            var topIsNumber = extractFirstNumber(lines[0].split(CELL_SEPARATOR)[i]) !== null;
            if (!topIsNumber && hasNumericBelow) return true;
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
        for (var c = 0; c < columnCount; c++) {
            var numericOnly = true;
            for (var r = startRow; r < Math.min(startRow + 3, lines.length); r++) {
                if (extractFirstNumber(lines[r].split(CELL_SEPARATOR)[c]) === null) {
                    numericOnly = false;
                    break;
                }
            }
            if (numericOnly) return c;
        }
        return 0;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 並び替えの列、順序、見出し行の有無を選ぶダイアログを出す
     * @param {number} columnCount - 列の数
     * @param {string[]} lines - 行
     * @param {boolean} hasHeaderCandidate - 1行目だけタブが無く、見出しと判断済みか
     * @returns {Object|null} { column, order, useHeader }。キャンセル時は null
     */
    function showSortOptionsDialog(columnCount, lines, hasHeaderCandidate) {
        var sortDialog = new Window("dialog", getLabel("dialog.title"));
        sortDialog.orientation = "column";
        sortDialog.alignChildren = "left";

        var columnPanel = sortDialog.add("panel", undefined, getLabel("panel.sortColumn"));
        columnPanel.orientation = "column";
        columnPanel.alignChildren = "left";
        columnPanel.margins = [10, 25, 10, 10];

        var columnRadioGroup = columnPanel.add("group");
        columnRadioGroup.orientation = "column";
        columnRadioGroup.alignChildren = "left";

        /* 列ごとのラジオボタンに先頭3件の値を表示 / show the first three values on each column radio */
        var columnPreviews = collectColumnPreviews(columnCount, lines);
        var columnRadios = [];
        for (var i = 0; i < columnCount; i++) {
            var columnLabel = getLabel("radio.column").replace("%1", i + 1) + columnPreviews[i].join(", ") + "…";
            var columnRadio = columnRadioGroup.add("radiobutton", undefined, columnLabel);
            columnRadio.helpTip = getLabel("tooltip.column");
            columnRadios.push(columnRadio);
        }

        var orderPanel = sortDialog.add("panel", undefined, getLabel("panel.sortOrder"));
        orderPanel.orientation = "column";
        orderPanel.alignChildren = "left";
        orderPanel.margins = [10, 25, 10, 10];

        var orderRadioGroup = orderPanel.add("group");
        orderRadioGroup.orientation = "row";
        var ascendingRadio = orderRadioGroup.add("radiobutton", undefined, getLabel("radio.ascending"));
        ascendingRadio.helpTip = getLabel("tooltip.ascending");
        var descendingRadio = orderRadioGroup.add("radiobutton", undefined, getLabel("radio.descending"));
        descendingRadio.helpTip = getLabel("tooltip.descending");
        var randomRadio = orderRadioGroup.add("radiobutton", undefined, getLabel("radio.random"));
        randomRadio.helpTip = getLabel("tooltip.random");
        ascendingRadio.value = true;

        var headerCheckbox = sortDialog.add("checkbox", undefined, getLabel("checkbox.header"));
        headerCheckbox.helpTip = getLabel("tooltip.header");
        headerCheckbox.value = hasHeaderCandidate || looksLikeHeaderRow(columnCount, lines);

        /* 見出し行があれば2行目以降で既定の列を探す / skip the header row when choosing the default column */
        columnRadios[findDefaultColumn(columnCount, lines, headerCheckbox.value ? 1 : 0)].value = true;

        var btnRowGroup = sortDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "OK" });

        if (sortDialog.show() !== 1) return null;

        var selectedColumn = 0;
        for (var j = 0; j < columnRadios.length; j++) {
            if (columnRadios[j].value) {
                selectedColumn = j;
                break;
            }
        }

        return {
            column: selectedColumn,
            order: randomRadio.value ? "random" : (descendingRadio.value ? "desc" : "asc"),
            useHeader: headerCheckbox.value
        };
    }

    // =========================================
    // テキストフレームの書き換え / Rewriting text frames
    // =========================================

    /**
     * テキストフレームを位置（Y座標の降順、同じなら X座標の昇順）で並べた配列を返す
     * @param {TextFrame[]} frameList - テキストフレーム
     * @returns {TextFrame[]} 並べ替えた新しい配列（失敗時は元の配列）
     */
    function sortTextFramesByPosition(frameList) {
        try {
            /* position の読み取りに失敗したら元の順のまま / keep the original order if a position cannot be read */
            var sortedList = frameList.slice();
            sortedList.sort(function (a, b) {
                var aPosition = a.position;
                var bPosition = b.position;
                if (aPosition[1] > bPosition[1]) return -1;
                if (aPosition[1] < bPosition[1]) return 1;
                if (aPosition[0] < bPosition[0]) return -1;
                if (aPosition[0] > bPosition[0]) return 1;
                return 0;
            });
            return sortedList;
        } catch (e) {
            alert("ソート中にエラーが発生しました: " + e.message);
            return frameList;
        }
    }

    /**
     * 上下に並んだテキストフレームを、上から順に1つのテキストフレームへまとめる
     * @param {TextFrame[]} frames - まとめるテキストフレーム
     * @returns {void}
     */
    function mergeTextFramesVertically(frames) {
        if (frames.length < 2) return;

        /* いったん1行ずつのフレームに分ける / split into one frame per line first */
        var sortedFrames = sortTextFramesByPosition(frames);
        var lineFrames = [];
        for (var i = 0; i < sortedFrames.length; i++) {
            var frameLines = sortedFrames[i].contents.split(LINE_SEPARATOR);
            for (var j = 0; j < frameLines.length; j++) {
                if (frameLines[j] !== "") {
                    var lineFrame = sortedFrames[i].duplicate();
                    lineFrame.contents = frameLines[j];
                    lineFrame.top -= j * 20; /* 位置調整（ソート用） / offset so the lines sort in order */
                    lineFrames.push(lineFrame);
                }
            }
            sortedFrames[i].remove();
        }
        sortedFrames = sortTextFramesByPosition(lineFrames);

        var baseFrame = sortedFrames[0];
        for (var k = 1; k < sortedFrames.length; k++) {
            baseFrame.paragraphs.add('\n');
            var sourceParagraphs = sortedFrames[k].paragraphs;
            for (var p = 0; p < sourceParagraphs.length; p++) {
                sourceParagraphs[p].duplicate(baseFrame);
            }
            sortedFrames[k].remove();
        }
    }

    /**
     * 見出し行を残して本文だけを並べ替えた内容に置き換える
     * 複製で見出しと本文のフレームに分け、書き換えたあと1つにまとめ直す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string[]} sortedLines - 並べ替えた本文の行
     * @returns {void}
     */
    function replaceBodyKeepingHeader(textFrame, sortedLines) {
        var headerFrame = textFrame;
        var bodyFrame = textFrame.duplicate();

        /* 本文フレームを1行分下に移動（フォントサイズに基づく） / move the body down one line by font size */
        var textSize = bodyFrame.textRange.characterAttributes.size;
        if (!isNaN(textSize)) {
            bodyFrame.top -= textSize * 1.5;
        }

        /* 見出し：2行目以降を削除 / header: remove every line but the first */
        var headerParagraphs = headerFrame.textRange.paragraphs;
        for (var i = headerParagraphs.length - 1; i >= 1; i--) {
            headerParagraphs[i].remove();
        }

        /* 本文：1行目を削除 / body: remove the first line */
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
     * 選択したテキストフレームのタブ区切りテキストを、選んだ列で並べ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてください。");
            return;
        }

        var docSelection = app.activeDocument.selection;
        if (docSelection.length !== 1 || docSelection[0].typename !== "TextFrame") {
            alert("1つのテキストオブジェクトを選択してください。");
            return;
        }

        var textFrame = docSelection[0];

        /* 空行や空白のみの行を除く / drop empty and whitespace-only lines */
        var allLines = textFrame.contents.split(LINE_SEPARATOR);
        var lines = [];
        for (var i = 0; i < allLines.length; i++) {
            if (allLines[i].replace(/\s/g, "").length > 0) {
                lines.push(allLines[i]);
            }
        }

        /* 1行目だけタブが無く2行目がタブ区切りなら、1行目を見出しとみなす / a tab-less first line over tabbed lines is a header */
        var hasHeaderCandidate = (lines.length >= 2 && lines[0].indexOf(CELL_SEPARATOR) === -1 && lines[1].indexOf(CELL_SEPARATOR) !== -1);
        var columnCount = lines[hasHeaderCandidate ? 1 : 0].split(CELL_SEPARATOR).length;

        var sortOptions = showSortOptionsDialog(columnCount, lines, hasHeaderCandidate);
        if (sortOptions === null) return;

        var dataLines = sortOptions.useHeader ? lines.slice(1) : lines;
        var sortedLines = generateSortedLines(dataLines, sortOptions.column, sortOptions.order);

        if (sortOptions.useHeader) {
            replaceBodyKeepingHeader(textFrame, sortedLines);
        } else {
            textFrame.contents = sortedLines.join(LINE_SEPARATOR);
        }
        app.redraw();
    }

    main();

})();
