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
- v1.2.1 (20260618) : コードを整理（IIFE化・関数分割・命名見直し）、1行のときも横組みに統一、バウンディングボックスを選択状態に依存しないよう修正、ドキュメント未オープン時のガード追加

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
- v1.2.1 (20260618): Refactored (IIFE wrap, function split, renaming); single line now forced horizontal; bounding box no longer depends on selection state; added guard for no open document

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextMergeToAreaBox";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var LINE_Y_THRESHOLD = 5; // 行分類のY座標差の閾値（pt） / Y threshold for grouping lines (pt)
    var MIN_LEADING_RATIO = 1.2; // 最小の行送り倍率 / Minimum leading ratio

    // =========================================
    // ローカライズ / Localize
    // =========================================
    var lang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* ラベル定義 / Label definitions */
    var LABELS = {
        errorNoText: {
            ja: "変換できるテキストはありません。",
            en: "No convertible text found."
        }
    };


    /*
      メイン処理：選択テキストを行ごとに連結し、エリア内文字を生成
      Main: merge selected text by line and build an area text frame
    */
    function main() {
        /* ドキュメント未オープン時は終了 / Exit when no document is open */
        if (app.documents.length === 0) {
            return;
        }

        var textLines = groupTextFramesByLine(activeDocument.selection);

        if (textLines.length === 0) {
            /* エラーメッセージの表示 / Show error message */
            alert(LABELS.errorNoText[lang]);
            return;
        }

        // 1行だけの場合は別処理：エリア内文字にせず左揃えで出力
        if (textLines.length === 1) {
            var singleLineFrame = mergeLineFrames(textLines[0]);
            singleLineFrame.paragraphs[0].justification = Justification.LEFT;
            app.selection = [singleLineFrame];
            return;
        }

        // 各行を1つのテキストフレームに連結
        var mergedTextFrames = [];
        for (var i = 0; i < textLines.length; i++) {
            mergedTextFrames.push(mergeLineFrames(textLines[i]));
        }

        var fontSize = mergedTextFrames[0].textRange.characterAttributes.size;
        var finalText = joinLineContents(mergedTextFrames);

        /*
          マージ済みフレーム全体のバウンディングボックスから作成位置・サイズを決定
          （選択状態に依存しないようmergedTextFramesから取得）
          Decide area-text position and size from the merged frames' bounding box
          (read from mergedTextFrames so it does not depend on the current selection)
        */
        var mergedBounds = getCombinedBounds(mergedTextFrames);
        var boundsLeft = mergedBounds[0];
        var boundsTop = mergedBounds[1];
        var boundsRight = mergedBounds[2];
        var boundsBottom = mergedBounds[3];
        var areaWidth = (boundsRight - boundsLeft) - fontSize; // 作成幅を1文字分縮める
        var areaHeight = boundsTop - boundsBottom;

        /*
          長方形を作成してエリア内文字に変換
          Create a rectangle and convert it to area text
        */
        var areaRect = activeDocument.pathItems.rectangle(boundsTop, boundsLeft, areaWidth, areaHeight);
        areaRect.stroked = false;
        areaRect.filled = false;

        var areaTextFrame = activeDocument.textFrames.areaText(areaRect);
        areaTextFrame.contents = finalText;
        areaTextFrame.textRange.characterAttributes.textFont = mergedTextFrames[0].textRange.characterAttributes.textFont;
        areaTextFrame.textRange.characterAttributes.size = mergedTextFrames[0].textRange.characterAttributes.size;
        areaTextFrame.paragraphs[0].paragraphAttributes.kinsoku = "Soft";
        areaTextFrame.paragraphs[0].justification = Justification.FULLJUSTIFYLASTLINELEFT;
        areaTextFrame.textRange.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        areaTextFrame.textRange.paragraphAttributes.justification = Justification.FULLJUSTIFYLASTLINELEFT;

        var leading = computeLeading(mergedTextFrames, fontSize);
        if (leading !== null) {
            areaTextFrame.textRange.characterAttributes.autoLeading = false;
            areaTextFrame.textRange.characterAttributes.leading = leading;
        }

        // 連結に使った元のmergedTextFramesを削除
        for (var i = 0; i < mergedTextFrames.length; i++) {
            mergedTextFrames[i].remove();
        }

        // 生成されたエリア内文字を選択状態にする
        app.selection = null; // 選択を一度解除
        app.selection = [areaTextFrame]; // 再度選択
        app.redraw();
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
      1行分のテキストフレームをX順に連結し、1つのテキストフレームにまとめる
      Merge one line's text frames (left to right) into a single text frame
    */
    function mergeLineFrames(lineFrames) {
        var framesInLine = sortItemsByPosition(lineFrames, "x");

        var mergedText = "";
        for (var j = 0; j < framesInLine.length; j++) {
            mergedText += framesInLine[j].contents;
        }

        var firstFrame = framesInLine[0];
        var originalPosition = firstFrame.position;

        var mergedFrame = firstFrame.duplicate();
        mergedFrame.orientation = TextOrientation.HORIZONTAL;
        mergedFrame.move(firstFrame, ElementPlacement.PLACEBEFORE);
        mergedFrame.contents = mergedText;
        mergedFrame.position = originalPosition;

        for (var k = 0; k < framesInLine.length; k++) {
            framesInLine[k].remove();
        }

        return mergedFrame;
    }

    /*
      各行のmergedTextFrame.contentsを連結する
      Join each line's contents into the final text
      - 文末が「。」「！」「？」、または英文の「.」「!」「?」のときのみ改行
      - 英単語どうしが隣り合う場合はスペースを挿入
      - 英単語がハイフンで分断されている場合はハイフンを除去して結合
    */
    function joinLineContents(mergedTextFrames) {
        var finalText = "";
        for (var i = 0; i < mergedTextFrames.length; i++) {
            var content = mergedTextFrames[i].contents;
            finalText += content;

            if (i >= mergedTextFrames.length - 1) {
                continue; // 最終行は連結処理なし
            }

            var nextContent = mergedTextFrames[i + 1].contents;
            var endsWithENWord = /[A-Za-z0-9)]$/.test(content);
            var startsWithENWord = /^[A-Za-z0-9(]/.test(nextContent);

            if (startsWithENWord) {
                if (endsWithENWord) {
                    finalText += " "; // 英単語どうしの間にスペース
                } else if (/[A-Za-z0-9)]-$/.test(content)) {
                    finalText = finalText.replace(/-$/, ""); // ハイフン分断を結合
                }
            }

            // 文末でのみ改行を追加（英文判定で日本語末尾の誤改行を防ぐ）
            var endsWithJP = /[。！？]$/.test(content);
            var endsWithEN = /[.!?]$/.test(content);
            var isEnglish = /^[\x00-\x7F]+$/.test(content.replace(/[\s\r\n]/g, ""));
            if (endsWithJP || (endsWithEN && !isEnglish)) {
                finalText += "\r";
            }
        }
        return finalText;
    }

    /*
      隣り合う2行のY差から行送りを求める（2行未満はnull）
      Derive leading from the Y gap of the first two lines (null if fewer than two)
    */
    function computeLeading(mergedTextFrames, fontSize) {
        if (mergedTextFrames.length < 2) {
            return null;
        }
        var y1 = mergedTextFrames[0].position[1];
        var y2 = mergedTextFrames[1].position[1];
        var leading = Math.abs(y1 - y2); // 行送りとして y 差を使用
        if (leading < fontSize) {
            leading = fontSize * MIN_LEADING_RATIO; // 最小でも MIN_LEADING_RATIO 倍にする
        }
        return leading;
    }

    /*
      アイテムをXまたはY座標で並び替え
      Sort items by X or Y position
    */
    function sortItemsByPosition(itemsToSort, axis) {
        var sortedItems = itemsToSort.slice(0);
        var compareItems;
        if (axis === "x") {
            compareItems = function(a, b) {
                return a["left"] - b["left"];
            };
        } else if (axis === "y") {
            compareItems = function(a, b) {
                return b["top"] - a["top"];
            };
        }
        sortedItems.sort(compareItems);
        return sortedItems;
    }

    /*
      複数アイテム全体のバウンディングボックスを取得
      Get the combined bounding box of multiple items
    */
    function getCombinedBounds(items) {
        var x1 = items[0].visibleBounds[0];
        var y1 = items[0].visibleBounds[1];
        var x2 = items[0].visibleBounds[2];
        var y2 = items[0].visibleBounds[3];
        for (var i = 1; i < items.length; i++) {
            var bounds = items[i].visibleBounds;
            if (bounds[0] < x1) x1 = bounds[0];
            if (bounds[1] > y1) y1 = bounds[1];
            if (bounds[2] > x2) x2 = bounds[2];
            if (bounds[3] < y2) y2 = bounds[3];
        }
        return [x1, y1, x2, y2];
    }

    /*
      テキストフレームをY位置で上から下へ並び替え
      Sort text frames from top to bottom by Y position
    */
    function sortTextFramesByY(textItems) {
        return textItems.sort(function(a, b) {
            return b.position[1] - a.position[1];
        });
    }

    /*
      Y位置が近いフレームを同じ行としてグループ化
      Group frames whose Y positions are within the threshold into the same line
    */
    function groupByLineY(sortedItems, yThreshold) {
        var lineGroups = [];
        for (var i = 0; i < sortedItems.length; i++) {
            var frame = sortedItems[i];
            var y = frame.position[1];
            var found = false;
            for (var j = 0; j < lineGroups.length; j++) {
                var lineY = lineGroups[j][0].position[1];
                if (Math.abs(lineY - y) <= yThreshold) {
                    lineGroups[j].push(frame);
                    found = true;
                    break;
                }
            }
            if (!found) {
                lineGroups.push([frame]);
            }
        }
        return lineGroups;
    }

})();
