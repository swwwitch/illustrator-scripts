/*

MimicDynamicText.jsx

【スクリプトの概要】
選択されたエリア内文字をポイント文字に変換し、改行ごとに分割して個別のテキストフレームとして再配置します。
その後、分割されたポイント文字を連結し、最終的に再びエリア内文字としてまとめ直す処理を行います。

【処理の流れ】
1. エリア内文字が選択されているかを確認
2. エリア内文字をポイント文字に変換し、改行ごとに分割して再配置
3. 分割されたポイント文字を連結して一つのテキストフレームにまとめる
4. 最終的にポイント文字をエリア内文字に再変換し、元の幅や自動行送りを設定

【対象オブジェクト】
- エリア内文字（AREATEXT）

【対象外】
- ポイント文字（POINTTEXT）
- パス上文字（PATHTEXT）

【更新履歴】
2025-06-18: 初版作成
2025-06-18: 最終的にエリア内文字へ戻す処理を追加
*/

#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// エリア内文字をポイント文字に変換し、変換後のTextFrameを取得
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

// エリア内文字を行単位に分割し、ポイント文字のTextFrameとして再配置
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
        // 行間を自動に設定し、autoLeadingAmount を 100 に設定
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

// 垂直比率を水平比率に合わせて統一
function matchVerticalScaleToHorizontal(item) {
    var currentHorizontalScale = item.textRange.characterAttributes.horizontalScale;
    item.textRange.characterAttributes.verticalScale = currentHorizontalScale;
}

// テキストフレームのみをリストに追加
function collectTextFrame(item, list) {
    if (item.typename === 'TextFrame') {
        list.push(item);
    }
}

// 複数のテキストフレームを縦方向に連結して再構成
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
                tf.top -= j * 2000; // 行順を維持するために位置調整
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

// テキストフレームを位置情報（上→下、左→右）でソート
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
        alert("ソート中にエラーが発生しました: " + e.message);
        return frameList;
    }
}

function main() {
    // 選択確認
    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
        alert("エリア内文字を選択してください。");
        return;
    }

    var areaTextFrame = app.activeDocument.selection[0];

    if (areaTextFrame.typename !== "TextFrame" || areaTextFrame.kind !== TextType.AREATEXT) {
        alert("エリア内文字を選択してください。");
        return;
    }

    var areaWidth = areaTextFrame.width;

    // 分割
    splitTextFrameIntoLines(areaTextFrame);

    // 連結
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