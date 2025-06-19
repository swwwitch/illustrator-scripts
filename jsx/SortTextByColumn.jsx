#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  スクリプト名：SortTextByColumn.jsx

  概要：
  テキストフレーム内のタブ区切りテキストを指定列の値で並び替えます。
  数値・文字列の昇順・降順・ランダム順に対応しています。

  処理概要：
  1. テキストフレームが1つ選択されているか確認
  2. テキストを行単位で分割し、空行を除去
  3. ダイアログで並び替え対象の列、順序、見出し行の有無を設定
  4. 指定列で並び替えを実行
  5. 見出し行ありの場合はヘッダーと本文でテキストフレームを分割し再構築
  6. 結果をテキストフレームに反映（Undo対応）

  作成日：2025-06-15
  最終更新日：2025-06-17
  - v1.0.0 初版
  - v1.0.1 ［1行目を見出し行として扱う］オプションを追加
  - v1.0.2 数値の抽出処理を改善し、カンマ区切りの数値も対応
  - v1.0.3 ランダムソートの実装を追加
  - v1.0.4 数字だけの列を自動的に選択する機能を追加  
  - v1.0.5 見出し行の有無を自動判定する機能を追加
    - v1.0.6 見出し行にtabコード不完全でもスルー
*/

function main() {
    var lang = getCurrentLang();

    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。");
        return;
    }

    var sel = app.activeDocument.selection;
    if (sel.length !== 1 || sel[0].typename !== "TextFrame") {
        alert("1つのテキストオブジェクトを選択してください。");
        return;
    }

    var textFrame = sel[0];
    var lineBreak = "\r";
    var tabChar = "\t";
    var lines = textFrame.contents.split(lineBreak);

    /* 空行や空白のみの行を除去 */
    var filteredLines = [];
    for (var i = 0; i < lines.length; i++) {
        if (lines[i].replace(/\s/g, "").length > 0) {
            filteredLines.push(lines[i]);
        }
    }
    lines = filteredLines;

    var previewCols;
    var hasHeaderCandidate = false;

    // 一時的に全行をチェックし、2行目以降がタブ区切りかつ1行目がそうでない場合に見出し候補と判断
    if (lines.length >= 2 && lines[0].indexOf(tabChar) === -1 && lines[1].indexOf(tabChar) !== -1) {
        hasHeaderCandidate = true;
        previewCols = lines[1].split(tabChar);
    } else {
        previewCols = lines[0].split(tabChar);
    }

    var sortOptions = showSortOptionsDialog(previewCols, lines, lang, hasHeaderCandidate);
    if (sortOptions === null) return;

    var selectedColumn = sortOptions.column;
    var sortOrder = sortOptions.order;

    var dataLines = sortOptions.useHeader ? lines.slice(1) : lines;
    var sortedLines = generateSortedLines(dataLines, selectedColumn, sortOrder);

    if (sortOptions.useHeader) {
        /* 見出し行あり：元のテキストフレームを複製しヘッダーと本文に分割 */
        var headerFrame = textFrame;
        var bodyFrame = textFrame.duplicate();

        /* 本文フレームを1行分下に移動（フォントサイズに基づく） */
        var textSize = bodyFrame.textRange.characterAttributes.size;
        if (!isNaN(textSize)) {
            bodyFrame.top -= textSize * 1.5;
        }

        /* ヘッダー：2行目以降を削除 */
        var fullText = headerFrame.textRange;
        var para = fullText.paragraphs;
        for (var i = para.length - 1; i >= 1; i--) {
            para[i].remove();
        }

        /* 本文：1行目を削除 */
        var fullTextBody = bodyFrame.textRange;
        var paraBody = fullTextBody.paragraphs;
        if (paraBody.length > 1) {
            paraBody[0].remove();
        }

        bodyFrame.contents = sortedLines.join(lineBreak);

        /* ヘッダーと本文のテキストフレームを垂直方向に統合 */
        mergeTextFramesVertically([headerFrame, bodyFrame]);
    } else {
        textFrame.contents = sortedLines.join(lineBreak);
    }
    app.redraw();
}

/* 指定列と並び順に基づき行をソートし、並び替えた行の配列を返す */
function generateSortedLines(lines, columnIndex, sortOrder) {
    var tabChar = "\t";
    var linesWithKeys = [];

    for (var i = 0; i < lines.length; i++) {
        var cells = lines[i].split(tabChar);
        var num = extractFirstNumber(cells[columnIndex] || "");
        var key = num !== null ? num : (cells[columnIndex] || "");
        linesWithKeys.push({
            line: lines[i],
            key: key
        });
    }

    if (sortOrder === "random") {
        sortLinesRandom(linesWithKeys);
    } else {
        linesWithKeys.sort(function(a, b) {
            if (typeof a.key === "number" && typeof b.key === "number") {
                return sortLinesNumeric(a, b, sortOrder);
            } else {
                return sortLinesString(a, b, sortOrder);
            }
        });
    }

    var sortedLines = [];
    for (var j = 0; j < linesWithKeys.length; j++) {
        sortedLines.push(linesWithKeys[j].line);
    }
    return sortedLines;
}

/* 配列をランダムにシャッフル */
function sortLinesRandom(linesWithKeys) {
    for (var i = linesWithKeys.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = linesWithKeys[i];
        linesWithKeys[i] = linesWithKeys[j];
        linesWithKeys[j] = temp;
    }
}

/* 数値キーの昇順・降順比較 */
function sortLinesNumeric(a, b, sortOrder) {
    return sortOrder === "asc" ? (a.key - b.key) : (b.key - a.key);
}

/* 文字列キーの昇順・降順比較（大文字小文字区別なし） */
function sortLinesString(a, b, sortOrder) {
    var aStr = String(a.key).toLowerCase();
    var bStr = String(b.key).toLowerCase();
    if (sortOrder === "asc") {
        return aStr < bStr ? -1 : aStr > bStr ? 1 : 0;
    } else {
        return aStr > bStr ? -1 : aStr < bStr ? 1 : 0;
    }
}

/* テキストから最初に現れる数値（整数・小数・カンマ区切り）を抽出して数値化 */
function extractFirstNumber(text) {
    var match = text.match(/[\d,]+(\.\d+)?/);
    if (match) {
        var normalized = match[0].replace(/,/g, "");
        return parseFloat(normalized);
    }
    return null;
}

/* 現在のロケールに基づき言語コードを返す（日本語か英語） */
function getCurrentLang() {
    return $.locale.indexOf("ja") === 0 ? "ja" : "en";
}

/* UIラベル（表示順：列 → 順序 → 見出し） */
var LABELS = {
    "SORT_COLUMN": {
        ja: "ソート対象の列",
        en: "Sort Target Column"
    },
    "SORT_ORDER": {
        ja: "ソート方法",
        en: "Sort Order"
    },
    "ASCENDING": {
        ja: "昇順",
        en: "Ascending"
    },
    "DESCENDING": {
        ja: "降順",
        en: "Descending"
    },
    "RANDOM": {
        ja: "ランダム",
        en: "Random"
    },
    "HEADER": {
        ja: "1行目を見出し行として扱う",
        en: "Treat first row as header"
    }
};

/* 並び替え対象の列、順序、見出し行の有無を選択するダイアログ表示 */
function showSortOptionsDialog(columns, lines, lang, hasHeaderCandidate) {
    var dlg = new Window("dialog", LABELS.SORT_COLUMN[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var columnPanel = dlg.add("panel", undefined, LABELS.SORT_COLUMN[lang]);
    columnPanel.orientation = "column";
    columnPanel.alignChildren = "left";
    columnPanel.margins = [10, 25, 10, 10];
    var radioButtons = [];

    var columnGroup = columnPanel.add("group");
    columnGroup.orientation = "column";
    columnGroup.alignChildren = "left";

    /* 各列の先頭3件の値をプレビューとして収集 */
    var previews = [];
    for (var i = 0; i < columns.length; i++) {
        previews[i] = [];
    }
    for (var r = 0; r < lines.length; r++) {
        var cells = lines[r].split("\t");
        for (var c = 0; c < previews.length; c++) {
            if (previews[c].length < 3 && cells[c]) {
                previews[c].push(cells[c]);
            }
        }
    }

    /* 列ごとにラジオボタンを作成しプレビューを表示 */
    for (var i = 0; i < columns.length; i++) {
        var previewStr = previews[i].join(", ");
        var btnLabel = (lang === "ja" ? "【列" : "[Row ") + (i + 1) + (lang === "ja" ? "】" : "] ") + previewStr + "…";
        var btn = columnGroup.add("radiobutton", undefined, btnLabel);
        radioButtons.push(btn);
    }

    var orderPanel = dlg.add("panel", undefined, LABELS.SORT_ORDER[lang]);
    orderPanel.orientation = "column";
    orderPanel.alignChildren = "left";
    orderPanel.margins = [10, 25, 10, 10];

    var orderGroup = orderPanel.add("group");
    orderGroup.orientation = "row";
    var ascBtn = orderGroup.add("radiobutton", undefined, LABELS.ASCENDING[lang]);
    var descBtn = orderGroup.add("radiobutton", undefined, LABELS.DESCENDING[lang]);
    var randomBtn = orderGroup.add("radiobutton", undefined, LABELS.RANDOM[lang]);
    ascBtn.value = true;


    var headerCheckbox = dlg.add("checkbox", undefined, LABELS.HEADER[lang]);
    headerCheckbox.value = false;

    if (hasHeaderCandidate) {
        headerCheckbox.value = true;
    } else {
        // 既存の自動判定ロジックをここに置く
        for (var i = 0; i < columns.length; i++) {
            var hasNumericBelow = false;
            for (var r = 1; r < lines.length; r++) {
                var cell = lines[r].split("\t")[i];
                if (extractFirstNumber(cell) !== null) {
                    hasNumericBelow = true;
                    break;
                }
            }
            var topCell = lines[0].split("\t")[i];
            var topIsNumber = extractFirstNumber(topCell) !== null;
            if (!topIsNumber && hasNumericBelow) {
                headerCheckbox.value = true;
                break;
            }
        }
    }

    /* デフォルト選択：見出し行がある場合は2行目以降を調査 */
    var defaultColIndex = 0;
    var startRow = headerCheckbox.value ? 1 : 0;
    for (var c = 0; c < columns.length; c++) {
        var numericOnly = true;
        for (var r = startRow; r < Math.min(startRow + 3, lines.length); r++) {
            var cell = lines[r].split("\t")[c];
            if (extractFirstNumber(cell) === null) {
                numericOnly = false;
                break;
            }
        }
        if (numericOnly) {
            defaultColIndex = c;
            break;
        }
    }
    radioButtons[defaultColIndex].value = true;


    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "right";
    btnGroup.add("button", undefined, "Cancel", {
        name: "cancel"
    });
    btnGroup.add("button", undefined, "OK", {
        name: "OK"
    });

    if (dlg.show() !== 1) return null;

    var selectedIndex = 0;
    for (var j = 0; j < radioButtons.length; j++) {
        if (radioButtons[j].value) {
            selectedIndex = j;
            break;
        }
    }

    return {
        column: selectedIndex,
        order: randomBtn.value ? "random" : (descBtn.value ? "desc" : "asc"),
        useHeader: headerCheckbox.value
    };
}

/* テキストフレームを垂直方向に統合して1つのテキストフレームにまとめる */
function mergeTextFramesVertically(frames) {
    if (frames.length < 2) return;

    var sortedFrames = sortTextFramesByPosition(frames);
    var splitFrames = [];
    for (var i = 0; i < sortedFrames.length; i++) {
        var lines = sortedFrames[i].contents.split('\r');
        for (var j = 0; j < lines.length; j++) {
            if (lines[j] !== "") {
                var tf = sortedFrames[i].duplicate();
                tf.contents = lines[j];
                tf.top -= j * 20; // 位置調整（ソート用）
                splitFrames.push(tf);
            }
        }
        sortedFrames[i].remove();
    }
    sortedFrames = sortTextFramesByPosition(splitFrames);

    var baseFrame = sortedFrames[0];
    for (var k = 1; k < sortedFrames.length; k++) {
        baseFrame.paragraphs.add('\n');
        var paragraphs = sortedFrames[k].paragraphs;
        for (var p = 0; p < paragraphs.length; p++) {
            paragraphs[p].duplicate(baseFrame);
        }
        sortedFrames[k].remove();
    }
}

/* テキストフレームを位置（Y座標降順、X座標昇順）でソートして返す */
function sortTextFramesByPosition(frameList) {
    try {
        var copyList = frameList.slice();

        copyList.sort(function(a, b) {
            if (a.position[1] > b.position[1]) return -1;
            if (a.position[1] < b.position[1]) return 1;
            if (a.position[1] === b.position[1]) {
                if (a.position[0] < b.position[0]) return -1;
                if (a.position[0] > b.position[0]) return 1;
                return 0;
            }
        });
        return copyList;
    } catch (e) {
        alert("ソート中にエラーが発生しました: " + e.message);
        return frameList;
    }
}

main();