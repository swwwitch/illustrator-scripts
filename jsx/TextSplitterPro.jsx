#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：TextSplitterPro.jsx

概要：
選択したテキスト（ポイントテキスト／エリアテキスト／直線パス上のテキスト）を
1文字ずつ分割し、「等間隔」または「視覚的な位置（トラッキング考慮）」で再配置します。
複数行の場合は行単位に分割して処理します。

対象：
- PointText／AreaText／直線パス上のTextPath
- 複数オブジェクトおよび複数行対応

処理の流れ：
1. 配置方法の選択ダイアログを表示
2. 複数行の場合は行単位に分割
3. 各行を1文字ずつ分割して再配置
4. グループ化オプションに対応

更新履歴：
- 2025-06-09 1.0.0 初期バージョン
- 2025-06-09 1.0.1 テキストフレームの位置を考慮して配置するよう修正
- 2025-06-10 1.0.3 特定の文字を選択しているとき、テキストオブジェクトを選択するロジックを追加
- v1.0.4 ローカライズを調整
*/

var LABELS = {
    ja: {
        dialogTitle: "テキスト分割",
        groupLabel: "分割方法",
        even: "等間隔で並べる",
        visual: "見た目どおり",
        groupNone: "グループ化しない",
        groupEachLine: "各行ごとにグループ化",
        groupAll: "全体をグループ化",
        ok: "OK",
        cancel: "キャンセル",
        alertSelectOne: "テキストオブジェクトを1つ以上選択してください。"
    },
    en: {
        dialogTitle: "Select Placement Mode",
        groupLabel: "Placement Method",
        even: "Distribute Evenly",
        visual: "Preserve Visual Position",
        groupNone: "No grouping",
        groupEachLine: "Group by line",
        groupAll: "Group all together",
        ok: "OK",
        cancel: "Cancel",
        alertSelectOne: "Please select one or more text objects."
    }
};

// 現在の言語設定を判定し、'ja' または 'en' を返す
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

// 縦組みかどうかを判定
function isVerticalTextFrame(textFrame) {
    return textFrame.orientation === TextOrientation.VERTICAL;
}

// 選択中の TextRange を含む単一の TextFrame を選択し直すユーティリティ関数
function selectSingleTextFrameFromTextRange() {
    if (app.selection.constructor.name === "TextRange") {
        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length === 1) {
            app.executeMenuCommand("deselectall");
            app.selection = [textFramesInStory[0]];
            try {
                app.selectTool("Adobe Select Tool");
            } catch (e) {}
        }
    }
}

function main() {
    selectSingleTextFrameFromTextRange();

    var lang = getCurrentLang();

    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
        alert(LABELS[lang].alertSelectOne);
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    // 選択中のテキストに複数行が含まれるか判定
    var hasMultiLine = false;
    for (var s = 0; s < selection.length; s++) {
        if (selection[s].typename === "TextFrame" && selection[s].contents.indexOf("\r") !== -1) {
            hasMultiLine = true;
            break;
        }
    }

    var result = showPlacementModeDialog(lang, hasMultiLine);
    if (!result || (result.mode !== "even" && result.mode !== "visual")) return;

    var mode = result.mode;
    var groupMode = result.group;

    for (var s = 0; s < selection.length; s++) {
        var selected = selection[s];
        try {
            if (selected.typename !== "TextFrame") continue;
        } catch (e) {
            continue;
        }

        var baseLayer = selected.layer;

        if (selected.kind === TextType.PATHTEXT && selected.textPath) {
            splitPathText(selected, mode, groupMode);
            continue;
        }

        var isVertical = isVerticalTextFrame(selected);

        var splitFrames = splitTextFrameByLine(selected, isVertical);
        var allCharFrames = [];
        for (var i = 0; i < splitFrames.length; i++) {
            var item = splitFrames[i];
            var isVerticalLine = isVerticalTextFrame(item);
            var frames = (mode === "visual") ? splitTextToCharacters(item, groupMode === "each" ? "each" : "none", hasMultiLine) : splitEvenly(item, groupMode === "each" ? "each" : "none");
            allCharFrames = allCharFrames.concat(frames);
        }

        if (groupMode === "all" && allCharFrames.length > 0) {
            var lastLayer = allCharFrames[0].layer;
            groupItemsIfNeeded(allCharFrames, lastLayer);
        }
        // groupMode が "none" または "each" の場合は既に個別処理済み
    }
}

function splitPathText(textFrame, mode, groupMode) {
    var doc = app.activeDocument;
    var path = textFrame.textPath;
    var charCount = textFrame.contents.length;
    var size = textFrame.textRange.characterAttributes.size;
    var font = textFrame.textRange.characterAttributes.textFont;
    var fillColor = textFrame.textRange.characterAttributes.fillColor;
    var layer = textFrame.layer;

    // Approximate the path into point segments
    var sampleCount = Math.max(20, charCount * 10);
    var points = [];
    for (var i = 0; i <= sampleCount; i++) {
        var t = i / sampleCount;
        var pt = path.getPoint(t);
        var tangent = path.getTangent(t);
        points.push({ anchor: pt, angle: Math.atan2(tangent[1], tangent[0]) * 180 / Math.PI });
    }

    var charFrames = [];

    for (var i = 0; i < charCount; i++) {
        var ch = textFrame.contents.charAt(i);
        var index = Math.round(i * (points.length - 1) / Math.max(1, charCount - 1));
        var info = points[index];
        var tf = doc.textFrames.add();
        tf.contents = ch;
        tf.textRange.characterAttributes.size = size;
        tf.textRange.characterAttributes.textFont = font;
        tf.textRange.characterAttributes.fillColor = fillColor;
        tf.position = info.anchor;
        tf.rotate(info.angle);
        charFrames.push(tf);
    }

    textFrame.remove();

    finalizeCharacterFrames(charFrames, groupMode, layer);
}

function showPlacementModeDialog(lang, isMultiLine) {
    var dlg = new Window("dialog", LABELS[lang].dialogTitle);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var panel = dlg.add("panel", undefined, LABELS[lang].groupLabel);
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.margins = [15, 20, 15, 10];

    var visualBtn = panel.add("radiobutton", undefined, LABELS[lang].visual);
    var evenBtn = panel.add("radiobutton", undefined, LABELS[lang].even);
    visualBtn.value = true;

    var groupPanel = dlg.add("panel", undefined, LABELS[lang].groupLabel);
    groupPanel.orientation = "column";
    groupPanel.alignChildren = "left";
    groupPanel.margins = [15, 20, 15, 10];

    var groupNoneBtn = groupPanel.add("radiobutton", undefined, LABELS[lang].groupNone);
    var groupEachBtn = groupPanel.add("radiobutton", undefined, LABELS[lang].groupEachLine);
    var groupAllBtn = groupPanel.add("radiobutton", undefined, LABELS[lang].groupAll);
    groupNoneBtn.value = true;

    groupEachBtn.enabled = isMultiLine;
    groupAllBtn.enabled = true;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";

    var cancelBtn = btnGroup.add("button", undefined, LABELS[lang].cancel);
    var okBtn = btnGroup.add("button", undefined, LABELS[lang].ok, {
        name: "ok"
    });

    var result = null;
    okBtn.onClick = function() {
        result = {
            mode: evenBtn.value ? "even" : "visual",
            group: groupEachBtn.value ? "each" : (groupAllBtn.value ? "all" : "none")
        };
        dlg.close();
    };
    cancelBtn.onClick = function() {
        result = null;
        dlg.close();
    };

    dlg.show();
    return result;
}

// テキストフレームを行単位で分割し、それぞれ新規テキストフレームを作成
function splitTextFrameByLine(originalText, isVertical) {
    var doc = app.activeDocument;
    var lines = originalText.lines;
    var baseX = originalText.position[0];
    var baseY = originalText.position[1];
    var fontSize = originalText.textRange.characterAttributes.size;
    var leading = originalText.textRange.leading;
    if (isNaN(leading) || leading <= 0) {
        leading = fontSize * 1.2;
    }

    var lineIndex = 0;
    var result = [];

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].contents.replace(/[\r\n]+$/, ""); // 改行削除
        if (line === "") continue;

        var tf = doc.textFrames.add();
        tf.contents = line;

        var attr = tf.textRange.characterAttributes;
        var srcAttr = originalText.textRange.characterAttributes;

        attr.size = srcAttr.size;
        attr.textFont = srcAttr.textFont;
        attr.leading = srcAttr.leading;
        attr.tracking = srcAttr.tracking;
        attr.stroked = false;

        tf.orientation = originalText.orientation;
        tf.position = [baseX, baseY]; // 一時設定
        app.redraw(); // bounding box 計算のため必要

        var bounds = tf.visibleBounds;
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];

        var useLineOffset = lines.length > 1;
        tf.position = [
            baseX - (isVertical && useLineOffset ? width * (lineIndex - 1) : 0),
            baseY - (isVertical ? 0 : leading * lineIndex)
        ];

        result.push(tf);
        lineIndex++;
    }

    originalText.remove();
    return result;
}

// 文字単位に分割し、トラッキングをpt単位に変換して視覚的な位置で配置、グループ化対応
function splitTextToCharacters(textItem, groupMode, isMultiLine) {
    var doc = app.activeDocument;
    var charCount = textItem.characters.length;
    var xCursor = textItem.position[0];
    var yCursor = textItem.position[1];
    var originalX = xCursor;
    var size = textItem.textRange.characterAttributes.size;

    var charFrames = [];
    var layer = textItem.layer;

    var isVertical = isVerticalTextFrame(textItem);
    // 縦組かつ複数行の場合はX位置を保持
    var preserveX = isMultiLine && isVertical;

    var lastBaseline = textItem.characters[0].baseline;

    for (var j = 0; j < charCount; j++) {
        var currentChar = textItem.characters[j];
        var currentBaseline = currentChar.baseline;

        if (j > 0 && currentBaseline !== lastBaseline) {
            yCursor -= size * 1.2;
            xCursor = originalX;
        }

        var charFrame = createCharacterFrame(textItem, j, xCursor, yCursor);
        var charWidth = charFrame.width;

        var trackingPt = (j < charCount - 1) ? getTrackingInPt(currentChar, size) : 0;
        if (isVertical) {
            yCursor -= size + trackingPt;
            if (!preserveX) {
                xCursor = originalX;
            }
        } else {
            xCursor += charWidth + trackingPt;
        }

        charFrames.push(charFrame);
        lastBaseline = currentBaseline;
    }

    textItem.remove();

    if (groupMode !== "all") {
        finalizeCharacterFrames(charFrames, groupMode, layer);
    }
    return charFrames;
}

// 文字を等間隔に分割配置、グループ化対応
function splitEvenly(textItem, groupMode) {
    var charCount = textItem.characters.length;
    var size = textItem.textRange.characterAttributes.size;
    var spacing = size * 1 + getTrackingInPt(textItem.textRange, size);
    var x = textItem.position[0];
    var y = textItem.position[1];
    var charFrames = [];
    var layer = textItem.layer;

    var isVertical = isVerticalTextFrame(textItem);

    for (var i = 0; i < charCount; i++) {
        var offsetX = isVertical ? 0 : spacing * i;
        var offsetY = isVertical ? -spacing * i : 0;
        var frame = createCharacterFrame(textItem, i, x + offsetX, y + offsetY);
        charFrames.push(frame);
    }
    // Remove the original text frame after creating character frames
    textItem.remove();
    if (groupMode !== "all") {
        finalizeCharacterFrames(charFrames, groupMode, layer);
    }
    return charFrames;
}

// 指定文字を複製し位置を設定
function createCharacterFrame(textItem, charIndex, x, y) {
    var charFrame = textItem.duplicate();
    try {
        var orientation = textItem.textRange.characterAttributes.textOrientation;
        charFrame.textRange.characterAttributes.textOrientation = orientation;
    } catch (e) {
        // 何もしない
    }
    charFrame.contents = textItem.characters[charIndex].contents;
    charFrame.position = [x, y];
    try {
        var srcKerning = textItem.characters[charIndex].characterAttributes.kerning;
        charFrame.textRange.characterAttributes.kerning = srcKerning;
    } catch (e) {
        // 何もしない
    }
    return charFrame;
}

// トラッキング値をpt単位に変換して取得
function getTrackingInPt(charRange, size) {
    var tracking = 0;
    try {
        tracking = charRange.characterAttributes.tracking;
    } catch (e) {
        tracking = 0;
    }
    return tracking / 1000 * size;
}

// 複数アイテムをグループ化または単純に選択
function groupItemsIfNeeded(items, layer) {
    if (items.length <= 1) {
        for (var i = 0; i < items.length; i++) {
            items[i].selected = true;
        }
        return;
    }

    var group = layer.groupItems.add();
    for (var i = 0; i < items.length; i++) {
        items[i].move(group, ElementPlacement.PLACEATEND);
    }
    group.selected = true;
}

function finalizeCharacterFrames(charFrames, groupMode, layer) {
    if (groupMode === "all") {
        groupItemsIfNeeded(charFrames, layer);
    } else if (groupMode === "none") {
        for (var i = 0; i < charFrames.length; i++) {
            charFrames[i].selected = true;
        }
    } else if (groupMode === "each") {
        groupItemsIfNeeded(charFrames, layer);
    }
}

main();