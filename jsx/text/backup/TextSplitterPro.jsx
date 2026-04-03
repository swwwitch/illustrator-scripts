#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name：

TextSplitterPro.jsx

### 概要 / Overview

- 選択したテキストを1文字ずつ分割し、等間隔または見た目どおりに再配置するIllustrator用スクリプトです。
- ポイントテキスト、エリアテキスト、パス上テキストに対応し、複数行や複数オブジェクトを処理できます。
- An Illustrator script that splits selected text into individual characters and rearranges them either evenly or visually considering tracking.
- Supports point text, area text, and text on a path, handling multi-line and multiple objects.

### 主な機能 / Main Features

- 等間隔配置またはトラッキング考慮の視覚的配置を選択
- 行単位での分割と再構築
- グループ化オプション（なし／行単位／全体）
- 日本語／英語インターフェース対応
- Choose between evenly spaced or visually positioned arrangement
- Split and reconstruct by line
- Grouping options (none, by line, all together)
- Japanese and English UI support

### 処理の流れ / Process Flow

1. 配置モード選択ダイアログを表示
2. 行単位に分割（複数行テキストの場合）
3. 各行を1文字ずつ分割して再配置
4. グループ化設定に従ってまとめる
1. Show placement mode selection dialog
2. Split by line (for multi-line text)
3. Split each line into characters and rearrange
4. Group according to selected option

### 更新履歴 / Update History

- v1.0 (20250609) : 初期バージョン
- v1.1 (20250609) : テキストフレームの位置考慮処理を追加
- v1.3 (20250610) : 選択文字対応ロジック追加
- v1.4 (20250706) : ローカライズ調整
- v1.0 (20250609): Initial version
- v1.1 (20250609): Added text frame position consideration
- v1.3 (20250610): Added logic for selected characters
- v1.4 (20250706): Localization adjustments
*/

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* -------------------------------
   日英ラベル定義 / Japanese-English label definitions
-------------------------------- */
var LABELS = {
    dialogTitle: { ja: "テキスト分割", en: "Select Placement Mode" },
    groupLabel: { ja: "分割方法", en: "Placement Method" },
    even: { ja: "等間隔で並べる", en: "Distribute Evenly" },
    visual: { ja: "見た目どおり", en: "Preserve Visual Position" },
    groupNone: { ja: "グループ化しない", en: "No grouping" },
    groupEachLine: { ja: "各行ごとにグループ化", en: "Group by line" },
    groupAll: { ja: "全体をグループ化", en: "Group all together" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertSelectOne: { ja: "テキストオブジェクトを1つ以上選択してください。", en: "Please select one or more text objects." }
};

/* 縦組みかどうかを判定 / Check if text frame is vertical */
function isVerticalTextFrame(textFrame) {
    return textFrame.orientation === TextOrientation.VERTICAL;
}

/* 選択中の TextRange を含む単一の TextFrame を選択し直すユーティリティ関数 / Re-select single TextFrame if a TextRange is selected */
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

    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
        alert(LABELS.alertSelectOne[lang]);
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    /* 選択中のテキストに複数行が含まれるか判定 / Check if selected text contains multiple lines */
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
        /* groupMode が "none" または "each" の場合は既に個別処理済み / Already handled individually */
    }
}

/* パステキストを分割して配置 / Split text on path and arrange */
function splitPathText(textFrame, mode, groupMode) {
    var doc = app.activeDocument;
    var path = textFrame.textPath;
    var charCount = textFrame.contents.length;
    var size = textFrame.textRange.characterAttributes.size;
    var font = textFrame.textRange.characterAttributes.textFont;
    var fillColor = textFrame.textRange.characterAttributes.fillColor;
    var layer = textFrame.layer;

    /* パスを点群に近似 / Approximate the path into point segments */
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

/* ダイアログ作成 / Create dialog */
function showPlacementModeDialog(lang, isMultiLine) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var panel = dlg.add("panel", undefined, LABELS.groupLabel[lang]);
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.margins = [15, 20, 15, 10];

    var visualBtn = panel.add("radiobutton", undefined, LABELS.visual[lang]);
    var evenBtn = panel.add("radiobutton", undefined, LABELS.even[lang]);
    visualBtn.value = true;

    var groupPanel = dlg.add("panel", undefined, LABELS.groupLabel[lang]);
    groupPanel.orientation = "column";
    groupPanel.alignChildren = "left";
    groupPanel.margins = [15, 20, 15, 10];

    var groupNoneBtn = groupPanel.add("radiobutton", undefined, LABELS.groupNone[lang]);
    var groupEachBtn = groupPanel.add("radiobutton", undefined, LABELS.groupEachLine[lang]);
    var groupAllBtn = groupPanel.add("radiobutton", undefined, LABELS.groupAll[lang]);
    groupNoneBtn.value = true;

    groupEachBtn.enabled = isMultiLine;
    groupAllBtn.enabled = true;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";

    var cancelBtn = btnGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang], {
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

/* テキストフレームを行単位で分割し、それぞれ新規テキストフレームを作成 / Split text frame by line and create new frames */
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
        var line = lines[i].contents.replace(/[\r\n]+$/, ""); /* 改行削除 / Remove line breaks */
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
        tf.position = [baseX, baseY]; /* 一時設定 / Temporary position */
        app.redraw(); /* bounding box 計算のため必要 / Needed for bounding box calculation */

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

/* 文字単位に分割し、トラッキングをpt単位に変換して視覚的な位置で配置、グループ化対応 / Split text into characters with visual positioning and grouping */
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
    /* 縦組かつ複数行の場合はX位置を保持 / Preserve X position if vertical and multiline */
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

/* 文字を等間隔に分割配置、グループ化対応 / Split characters evenly with grouping */
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
    /* 元のテキストフレームは削除 / Remove the original text frame after creating character frames */
    textItem.remove();
    if (groupMode !== "all") {
        finalizeCharacterFrames(charFrames, groupMode, layer);
    }
    return charFrames;
}

/* 指定文字を複製し位置を設定 / Duplicate specified character and set position */
function createCharacterFrame(textItem, charIndex, x, y) {
    var charFrame = textItem.duplicate();
    try {
        var orientation = textItem.textRange.characterAttributes.textOrientation;
        charFrame.textRange.characterAttributes.textOrientation = orientation;
    } catch (e) {
        /* 何もしない / Do nothing */
    }
    charFrame.contents = textItem.characters[charIndex].contents;
    charFrame.position = [x, y];
    try {
        var srcKerning = textItem.characters[charIndex].characterAttributes.kerning;
        charFrame.textRange.characterAttributes.kerning = srcKerning;
    } catch (e) {
        /* 何もしない / Do nothing */
    }
    return charFrame;
}

/* トラッキング値をpt単位に変換して取得 / Convert tracking value to points */
function getTrackingInPt(charRange, size) {
    var tracking = 0;
    try {
        tracking = charRange.characterAttributes.tracking;
    } catch (e) {
        tracking = 0;
    }
    return tracking / 1000 * size;
}

/* 複数アイテムをグループ化または単純に選択 / Group multiple items or simply select */
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