#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

TextMergeToAreaBox.jsx

### 概要：

- 複数のテキストオブジェクトを1つのエリア内文字に連結します。
- 元のオブジェクトのサイズ・フォント・行送りを反映します。

### 主な機能：

- 改行位置の調整（末尾が「。」以外、「.」「?」「!」の場合は連結）
- JUSTIFY 揃えの自動適用
- 元のオブジェクトの削除と置換処理

### 処理の流れ：

1. 選択中のテキストオブジェクトを上から順にソート
2. 幅・高さ・行送りなどを取得
3. テキストを1つに連結し、エリア内文字を作成
4. 元のオブジェクトは削除

### 謝辞

倉田タカシさん（イラレで便利）
https://d-p.2-d.jp/ai-js/

### note

https://note.com/dtp_tranist/n/ne8d31278c266

### 更新履歴：

- v1.0 (20250717) : 初期バージョン
- v1.1 (20250718) : 1行だけに対応、禁則を設定
- v1.2 (20250719) : 行末が英単語の場合の改行処理を追加

---

### Script Name：

TextMergeToAreaBox.jsx

### Overview：

- Merges multiple text items into a single area text box.
- Inherits font, size, and leading from the original text.

### Key Features：

- Removes line breaks except after "。", ".", "!", or "?"
- Applies JUSTIFY alignment automatically
- Replaces and deletes original text items

### Workflow：

1. Sort selected text items from top to bottom
2. Extract dimensions and leading
3. Merge text and create area text
4. Delete original items

### Change Log：

- v1.0 (20250718): Initial release
- v1.1 (20250719): Added support for single line text, set kinsoku rules
- v1.2 (20250720): Added handling for line breaks after English words

*/

var SCRIPT_VERSION = "v1.2";
var LINE_Y_THRESHOLD = 5; // 行分類のY座標差の閾値（ポイント単位）
var MIN_LEADING_RATIO = 1.2; // 最小の行送り倍率

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ラベル定義 / Label definitions */
var LABELS = {
    errorNoText: {
        ja: "変換できるテキストはありません。",
        en: "No convertible text found."
    }
};


function main() {
    var textLines = groupTextFramesByLine(activeDocument.selection);

    if (textLines.length != 0) {

        // 1行だけの場合は別処理：エリア内文字にせず左揃えで出力
        if (textLines.length === 1) {
            var sortedItems = sortItemsByPosition(textLines[0], "x");
            var mergedText = "";
            for (var j = 0; j < sortedItems.length; j++) {
                mergedText += sortedItems[j].contents;
            }
            var originalPosition = sortedItems[0].position;
            var textFrame = sortedItems[0].duplicate();
            textFrame.contents = mergedText;
            textFrame.position = originalPosition;
            textFrame.paragraphs[0].justification = Justification.LEFT;
            for (var k = 0; k < sortedItems.length; k++) {
                sortedItems[k].remove();
            }
            app.selection = [textFrame];
            return;
        }

        var mergedTextFrames = [];
        for (var i = 0; i < textLines.length; i++) {
            var sortedItems = sortItemsByPosition(textLines[i], "x");

            var mergedText = "";
            for (var j = 0; j < sortedItems.length; j++) {
                mergedText += sortedItems[j].contents;
            }
            var originalPosition = sortedItems[0].position;

            var mergedTextFrame = sortedItems[0].duplicate();

            mergedTextFrame.orientation = TextOrientation.HORIZONTAL;
            mergedTextFrame.move(sortedItems[0], ElementPlacement.PLACEBEFORE);

            mergedTextFrame.contents = mergedText;
            mergedTextFrame.position = originalPosition;

            mergedTextFrames.push(mergedTextFrame);

            for (var k = 0; k < sortedItems.length; k++) {
                sortedItems[k].remove();
            }
        }

        /*
          各行のmergedTextFrame.contentsを、「。」「！」「？」で終わる場合のみ改行で連結
          Join mergedTextFrame contents with line breaks only if ending with "。", "！", or "？"
          ただし英単語の末尾がピリオド「.」、疑問符「?」、感嘆符「!」の場合は改行しない
        */
        var finalText = "";
        var fontSize = mergedTextFrames[0].textRange.characterAttributes.size;

        for (var i = 0; i < mergedTextFrames.length; i++) {
            var content = mergedTextFrames[i].contents;

            finalText += content;

            // 次の行の先頭が英単語の場合、末尾と先頭の間にスペースを挿入
            if (i < mergedTextFrames.length - 1) {
                var nextContent = mergedTextFrames[i + 1].contents;
                var endsWithENWord = /[A-Za-z0-9)]$/.test(content);
                var startsWithENWord = /^[A-Za-z0-9(]/.test(nextContent);
                if (endsWithENWord && startsWithENWord) {
                    finalText += " ";
                }
                // 英単語がハイフンで分断されている場合、ハイフンを除去して結合
                if (/[A-Za-z0-9)]-$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                    // ハイフンを削除して連結（末尾のハイフンを除去済みのcontentを使う）
                    finalText = finalText.replace(/-$/, "");
                }
            }

            // 内容に応じて改行を追加
            var endsWithJP = /[。！？]$/.test(content);
            var endsWithEN = /[.!?]$/.test(content);
            var isEnglish = /^[\x00-\x7F]+$/.test(content.replace(/[\s\r\n]/g, ""));

            if (i < mergedTextFrames.length - 1) {
                if (endsWithJP || (endsWithEN && !isEnglish)) {
                    finalText += "\r";
                }
            }
        }

        /*
          ① 選択範囲全体のバウンディングボックスを取得
          Get bounding box of entire selection
        */
        var selectionBounds = getSelectionBounds(activeDocument.selection);
        var selLeft = selectionBounds[0];
        var selTop = selectionBounds[1];
        var selRight = selectionBounds[2];
        var selBottom = selectionBounds[3];
        var selWidth = selRight - selLeft;
        // 長方形（エリアテキスト）の作成幅を1文字分縮める
        // var fontSize = mergedTextFrames[0].textRange.characterAttributes.size;  ← 削除済み
        selWidth = selWidth - fontSize;
        var selHeight = selTop - selBottom;

        /*
          ② 長方形を作成
          Create rectangle for area text
        */
        var rect = activeDocument.pathItems.rectangle(selTop, selLeft, selWidth, selHeight);
        rect.stroked = false;
        rect.filled = false;

        /*
          ③ エリア内文字に変換
          Convert to area text
        */
        var newTextFrame = activeDocument.textFrames.areaText(rect);
        newTextFrame.contents = finalText;
        newTextFrame.textRange.characterAttributes.textFont = mergedTextFrames[0].textRange.characterAttributes.textFont;
        newTextFrame.textRange.characterAttributes.size = mergedTextFrames[0].textRange.characterAttributes.size;
        newTextFrame.paragraphs[0].paragraphAttributes.kinsoku = "Soft";
        newTextFrame.paragraphs[0].justification = Justification.FULLJUSTIFYLASTLINELEFT;
        newTextFrame.textRange.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        newTextFrame.textRange.paragraphAttributes.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        newTextFrame.paragraphs[0].justification = Justification.FULLJUSTIFYLASTLINELEFT;

        var leading = null;
        if (mergedTextFrames.length >= 2) {
            var y1 = mergedTextFrames[0].position[1];
            var y2 = mergedTextFrames[1].position[1];
            // var fontSize = mergedTextFrames[0].textRange.characterAttributes.size;  ← この行を削除
            leading = Math.abs(y1 - y2); // 行送りとして y 差を使用
            if (leading < fontSize) {
                leading = fontSize * MIN_LEADING_RATIO; // 最小でも MIN_LEADING_RATIO にする
            }
        }
        if (leading !== null) {
            newTextFrame.textRange.characterAttributes.autoLeading = false;
            newTextFrame.textRange.characterAttributes.leading = leading;
        }

        // 元のmergedTextFramesを削除
        for (var i = 0; i < mergedTextFrames.length; i++) {
            mergedTextFrames[i].remove();
        }

        // 生成されたエリア内文字を選択状態にする
        app.selection = null; // 選択を一度解除
        app.selection = [newTextFrame]; // 再度選択
        app.redraw();

    } else {
        /* エラーメッセージの表示 / Show error message */
        alert(LABELS.errorNoText[lang]);
    }
}

main();

/*
  選択したテキストアイテムをY位置で行単位に分類
  Group selected text frames by line based on Y position
*/
function groupTextFramesByLine(selection) {
    var textItems = [];
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "TextFrame") {
            textItems.push(selection[i]);
        }
    }
    var sortedItems = sortTextFramesByY(textItems);
    return groupByLineY(sortedItems, LINE_Y_THRESHOLD);
}

/*
  アイテムをXまたはY座標で並び替え
  Sort items by X or Y position
*/
function sortItemsByPosition(itemsToSort, XorY) {
    var itemsArray = itemsToSort.slice(0);
    var sortOrder;
    if (XorY == "x") {
        sortOrder = function(a, b) {
            return a["left"] - b["left"];
        };
    } else if (XorY == "y") {
        sortOrder = function(a, b) {
            return b["top"] - a["top"];
        };
    }
    itemsArray.sort(sortOrder);
    return itemsArray;
}

/*
  選択範囲全体のバウンディングボックスを取得
  Get bounding box of entire selection
*/
function getSelectionBounds(selection) {
    var x1 = selection[0].visibleBounds[0];
    var y1 = selection[0].visibleBounds[1];
    var x2 = selection[0].visibleBounds[2];
    var y2 = selection[0].visibleBounds[3];
    for (var i = 1; i < selection.length; i++) {
        var b = selection[i].visibleBounds;
        if (b[0] < x1) x1 = b[0];
        if (b[1] > y1) y1 = b[1];
        if (b[2] > x2) x2 = b[2];
        if (b[3] < y2) y2 = b[3];
    }
    return [x1, y1, x2, y2];
}

function sortTextFramesByY(textItems) {
    return textItems.sort(function(a, b) {
        return b.position[1] - a.position[1];
    });
}

function groupByLineY(sortedItems, threshold) {
    var lines = [];
    for (var i = 0; i < sortedItems.length; i++) {
        var tf = sortedItems[i];
        var y = tf.position[1];
        var found = false;
        for (var j = 0; j < lines.length; j++) {
            var lineY = lines[j][0].position[1];
            if (Math.abs(lineY - y) <= threshold) {
                lines[j].push(tf);
                found = true;
                break;
            }
        }
        if (!found) {
            lines.push([tf]);
        }
    }
    return lines;
}