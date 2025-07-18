#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

MimicDynamicText.jsx

### 概要

- エリア内文字をポイント文字に変換し、改行ごとに分割して個別に再配置するIllustrator用スクリプトです。
- 最終的にポイント文字を連結してエリア内文字に戻すことで、より柔軟な再編集が可能になります。

### 主な機能

- エリア内文字をポイント文字に変換
- 改行ごとに分割してポイント文字フレームを生成
- 複数のテキストフレームを連結して再構築
- 再びエリア内文字に変換して自動行送りを設定
- 日本語／英語インターフェース対応

### 処理の流れ

1. エリア内文字が選択されているか確認
2. ポイント文字に変換後、改行で分割し再配置
3. 分割されたテキストフレームを連結
4. エリア内文字に再変換して調整

### 更新履歴

- v1.0.0 (20250618) : 初期バージョン
- v1.0.1 (20250618) : エリア内文字へ戻す処理を追加

---

### Script Name:

MimicDynamicText.jsx

### Overview

- An Illustrator script that converts area text to point text, splits it line by line, and repositions each line individually.
- Finally merges the point texts and converts them back into area text for flexible editing.

### Main Features

- Convert area text to point text
- Split by line breaks and create individual point text frames
- Merge multiple text frames back into one
- Convert back to area text and apply auto leading
- Japanese and English UI support

### Process Flow

1. Check if area text is selected
2. Convert to point text, split by line, and reposition
3. Merge split text frames
4. Convert back to area text and adjust

### Update History

- v1.0.0 (20250618): Initial version
- v1.0.1 (20250618): Added process to convert back to area text
*/

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    alertSelectAreaText: { ja: "エリア内文字を選択してください。", en: "Please select area text." },
    alertSortError: { ja: "ソート中にエラーが発生しました: ", en: "An error occurred during sorting: " }
};


/* エリア内文字をポイント文字に変換し、変換後のTextFrameを取得 / Convert area text to point text and get the resulting TextFrame */
function convertAreaTextToPointText(areaTextFrame, originalContents, originalPosition) {
    areaTextFrame.convertAreaObjectToPointObject();

    var allTextFrames = app.activeDocument.textFrames;
    for (var i = 0; i < allTextFrames.length; i++) {
        var tf = allTextFrames[i];
        var dx = Math.abs(tf.position[0] - originalPosition[0]);
        var dy = Math.abs(tf.position[1] - originalPosition[1]);

        if (dx < 1 && dy < 1) {
            return tf;
        }
    }
    return null;
}

/* エリア内文字を行単位に分割し、ポイント文字のTextFrameとして再配置 / Split area text by lines and reposition as point text frames */
function splitTextFrameIntoLines(textFrame) {
    var lines = textFrame.contents.split('\r');
    var originalPosition = textFrame.position;
    var currentY = originalPosition[1];
    var textSize = textFrame.textRange.characterAttributes.size;
    var textFont = textFrame.textRange.characterAttributes.textFont;
    var textColor = textFrame.textRange.characterAttributes.fillColor;
    var hScale = textFrame.textRange.characterAttributes.horizontalScale;

    var splitLines = [];

    for (var i = 0; i < lines.length; i++) {
        var newLine = app.activeDocument.textFrames.pointText([originalPosition[0], currentY]);
        newLine.contents = lines[i];
        newLine.textRange.characterAttributes.size = textSize;
        newLine.textRange.characterAttributes.textFont = textFont;
        newLine.textRange.characterAttributes.fillColor = textColor;

        var actualWidth = newLine.width;
        var scaleX = (textFrame.width / actualWidth);
        if (scaleX > 0) {
            newLine.textRange.characterAttributes.size = textSize * scaleX;
            newLine.textRange.characterAttributes.horizontalScale = 100;
            newLine.textRange.characterAttributes.verticalScale = 100;
        }

        currentY -= newLine.height;
        splitLines.push(newLine);
    }
    var mergedFrame = mergeTextFramesVertically(splitLines);
    if (mergedFrame) {
        /* 行間を自動に設定し、autoLeadingAmount を 100 に設定 / Set auto leading and autoLeadingAmount to 110 */
        mergedFrame.textRange.characterAttributes.autoLeading = true;

        var paragraphs = mergedFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            paragraphs[i].paragraphAttributes.autoLeadingAmount = 110;
        }

        redraw();
        textFrame.remove();
        mergedFrame.convertPointObjectToAreaObject();
    }
}

/* 垂直比率を水平比率に合わせて統一 / Match vertical scale to horizontal scale */
function matchVerticalScaleToHorizontal(item) {
    var currentHorizontalScale = item.textRange.characterAttributes.horizontalScale;
    item.textRange.characterAttributes.verticalScale = currentHorizontalScale;
}

/* テキストフレームのみをリストに追加 / Add only text frames to list */
function collectTextFrame(item, list) {
    if (item.typename === 'TextFrame') {
        list.push(item);
    }
}

/* 複数のテキストフレームを縦方向に連結して再構成 / Merge multiple text frames vertically */
function mergeTextFramesVertically(frames) {
    if (frames.length < 2) {
        return;
    }

    var sortedFrames = sortTextFramesByPosition(frames);
    var splitFrames = [];
    for (var i = 0; i < sortedFrames.length; i++) {
        var lines = sortedFrames[i].contents.split('\r');
        for (var j = 0; j < lines.length; j++) {
            if (lines[j] !== "") {
                var tf = sortedFrames[i].duplicate();
                tf.contents = lines[j];
                tf.top -= j * 2000; /* 行順を維持するために位置調整 / Adjust position to maintain line order */
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
    return baseFrame;
}

/* テキストフレームを位置情報（上→下、左→右）でソート / Sort text frames by position (top to bottom, left to right) */
function sortTextFramesByPosition(frameList) {
    try {
        var copyList = [];
        var i;
        for (i = 0; i < frameList.length; i++) {
            copyList.push(frameList[i]);
        }

        copyList.sort(function(a, b) {
            if (a.position[1] > b.position[1]) {
                return -1;
            }
            if (a.position[1] < b.position[1]) {
                return 1;
            }
            if (a.position[1] === b.position[1]) {
                if (a.position[0] < b.position[0]) {
                    return -1;
                }
                if (a.position[0] > b.position[0]) {
                    return 1;
                }
                return 0;
            }
        });
        return copyList;
    } catch (e) {
        alert(LABELS.alertSortError[lang] + e.message);
        return frameList;
    }
}

/* メイン処理 / Main process */
function main() {
    /* 選択確認 / Check selection */
    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
        alert(LABELS.alertSelectAreaText[lang]);
        return;
    }

    var areaTextFrame = app.activeDocument.selection[0];

    if (areaTextFrame.typename !== "TextFrame" || areaTextFrame.kind !== TextType.AREATEXT) {
        alert(LABELS.alertSelectAreaText[lang]);
        return;
    }

    var areaWidth = areaTextFrame.width;

    /* 分割 / Split */
    splitTextFrameIntoLines(areaTextFrame);

    /* 連結 / Merge */
    var selectionItems = app.activeDocument.selection;
    if (selectionItems.length >= 2) {
        var textFrames = [];
        for (var i = 0; i < selectionItems.length; i++) {
            collectTextFrame(selectionItems[i], textFrames);
        }
        if (textFrames.length >= 2) {
            mergeTextFramesVertically(textFrames);
        }
    }
}

main();