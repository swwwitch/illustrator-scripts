#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SortByNumbers.jsx

### 概要

- グループ内のテキストから数値を抽出し、数値に基づいてグループを並び替え縦方向に整列するIllustrator用スクリプトです。
- フォント情報によるグループ分けと、昇順・降順・ランダムソートに対応します。

### 主な機能

- テキストから数値抽出（カンマ除去、小数対応）
- フォント名ごとのグループ選択とソート
- 昇順、降順、ランダムの並び替えオプション
- スペーシング（フィットまたはカスタム値）選択
- 日本語／英語インターフェース対応

### 処理の流れ

1. 数値を含むグループオブジェクトを選択
2. ダイアログでフォント、ソート順、スペーシングを設定
3. グループをソートし縦方向に配置
4. 全体位置を元の開始位置に補正

### 更新履歴

- v1.0.0 (20250615) : 初期バージョン
- v1.1.0 (20250616) : ラジオボタン選択状態保持機能追加

---

### Script Name:

SortByNumbers.jsx

### Overview

- An Illustrator script that extracts numbers from text inside groups, sorts groups based on those numbers, and arranges them vertically.
- Supports grouping by font info and sorting in ascending, descending, or random order.

### Main Features

- Extract numbers from text (ignores commas, supports decimals)
- Group and sort by font name
- Ascending, descending, or random sorting options
- Choose spacing mode: fit or custom value
- Japanese and English UI support

### Process Flow

1. Select group objects containing numbers
2. Configure font, sort order, and spacing in dialog
3. Sort and arrange groups vertically
4. Adjust overall position to original start point

### Update History

- v1.0.0 (20250615): Initial version
- v1.1.0 (20250616): Added radio button state retention
*/

var LABELS = {
    ja: {
        title: "グループの数値で整列",
        sortGroup: "数値グループ",
        asc: "昇順",
        desc: "降順",
        ok: "ソート",
        fit: "ぴったり",
        custom: "指定"
    },
    en: {
        title: "Align Groups by Number",
        sortGroup: "Number Group",
        asc: "Ascending",
        desc: "Descending",
        ok: "Sort",
        fit: "Fit",
        custom: "Custom"
    }
};

function getCurrentLang() {
    return ($.locale === "ja" || $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var sel = app.activeDocument.selection;
    if (sel.length === 0) {
        alert("グループオブジェクトを選択してください。");
        return;
    }

    var validGroups = [];
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "GroupItem") {
            var dummyMap = {};
            var firstValue = { value: undefined };
            collectTextWithFontInfoPerGroup(sel[i], dummyMap, sel[i], firstValue);
            if (!isNaN(firstValue.value)) {
                validGroups.push(sel[i]);
            }
        }
    }
    if (validGroups.length === 0) {
        alert("数値を含むグループが見つかりません。");
        return;
    }

    var fontMap = {};
    for (var i = 0; i < validGroups.length; i++) {
        collectTextWithFontInfoPerGroup(validGroups[i], fontMap, validGroups[i]);
    }

    var fontNames = [];
    for (var name in fontMap) fontNames.push(name);
    fontNames.sort();

    if (fontNames.length === 0) {
        alert("選択されたグループ内に数字テキストが含まれていません。");
        return;
    }

    var originalPositions = {};
    for (var i = 0; i < validGroups.length; i++) {
        var g = validGroups[i];
        originalPositions[g.name] = {
            group: g,
            bounds: g.visibleBounds.concat()
        };
    }

    var selected = showFontChoiceDialog(fontMap, originalPositions);
    if (!selected) return;
    var selectedFont = selected.font;
    var isDescending = selected.descending;

    var entries = fontMap[selectedFont];
    var groupValueMap = {};
    var groupCounter = 0;
    for (var i = 0; i < entries.length; i++) {
        var g = entries[i].group;
        var v = entries[i].value;

        var alreadyExists = false;
        for (var k in groupValueMap) {
            if (groupValueMap[k].group === g) {
                alreadyExists = true;
                break;
            }
        }
        if (!alreadyExists) {
            var key = "g" + groupCounter++;
            groupValueMap[key] = {
                group: g,
                value: v
            };
        }
    }
    var groupData = [];
    for (var id in groupValueMap) {
        groupData.push(groupValueMap[id]);
    }

    if (selected.random) {
        // Fisher–Yates shuffle。先頭が元の先頭と同じ場合は別の要素と交換
        var originalFirst = groupData[0];
        for (var i = groupData.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = groupData[i];
            groupData[i] = groupData[j];
            groupData[j] = temp;
        }
        if (groupData.length > 1 && groupData[0].group === originalFirst.group) {
            var swapIndex = 1;
            groupData[0] = groupData[swapIndex];
            groupData[swapIndex] = originalFirst;
        }
    } else {
        groupData.sort(function(a, b) {
            return isDescending ? b.value - a.value : a.value - b.value;
        });
    }

    var spacing = 20;
    if (selected.spacingMode === "custom" && !isNaN(selected.spacingValue)) {
        spacing = selected.spacingValue;
    }

    // 並び替え開始位置（Y座標）を算出。全グループの上端の最大値を基準にする
    var boundsList = [];
    for (var id in originalPositions) {
        boundsList.push(originalPositions[id].bounds);
    }
    var minY = boundsList[0][1];
    for (var i = 1; i < boundsList.length; i++) {
        if (boundsList[i][1] < minY) {
            minY = boundsList[i][1];
        }
    }
    var startTop = minY;

    // 左端は上端がstartTopのグループの左端を使用
    var originalTopLeft = null;
    for (var id in originalPositions) {
        if (originalPositions[id].bounds[1] === startTop) {
            originalTopLeft = originalPositions[id].bounds[0];
            break;
        }
    }
    var currentTop = startTop;
    var startLeft = originalTopLeft;

    // Store the original top-left position before moving groups
    var beforeTopLeftX = startLeft;
    var beforeTopLeftY = startTop;

    for (var j = 0; j < groupData.length; j++) {
        var g = groupData[j].group;
        g.locked = false;
        g.hidden = false;

        // move group to currentTop
        var gBoundsBefore = g.visibleBounds;
        var gLeft = gBoundsBefore[0];
        var gTop = gBoundsBefore[1];
        var dx = startLeft - gLeft;
        var dy = currentTop - gTop;
        g.translate(dx, dy);

        // get new height after translation
        var gBoundsAfter = g.visibleBounds;
        var height = gBoundsAfter[1] - gBoundsAfter[3];

        if (selected.spacingMode === "fit") {
            currentTop -= height;
        } else {
            currentTop -= (height + spacing);
        }
    }

    // 再配置後、最初のグループの位置を元の位置に合わせて全体を調整
    var firstGroupBounds = groupData[0].group.visibleBounds;
    var shiftX = beforeTopLeftX - firstGroupBounds[0];
    var shiftY = beforeTopLeftY - firstGroupBounds[1];
    for (var i = 0; i < groupData.length; i++) {
        groupData[i].group.translate(shiftX, shiftY);
    }

    app.redraw();
}

function restoreGroupPositions(positionMap) {
    for (var id in positionMap) {
        var g = positionMap[id].group;
        var bounds = g.visibleBounds;
        var orig = positionMap[id].bounds;
        g.translate(orig[0] - bounds[0], orig[1] - bounds[1]);
    }
}

/**
 * 再帰的にテキストフレームから数値を抽出し配列に収集
 * @param {PageItem} obj - 対象オブジェクト
 * @param {Array} textFrames - 収集先配列
 */
function collectTextFramesRecursive(obj, textFrames) {
    if (obj.typename === "TextFrame") {
        var text = obj.contents;
        var cleaned = text.replace(/,/g, "");
        if (/^\d+(\.\d+)?$/.test(cleaned)) {
            var number = parseFloat(cleaned);
            if (!isNaN(number)) {
                textFrames.push({
                    item: obj,
                    value: number
                });
            }
        }
    } else if (obj.typename === "GroupItem") {
        for (var i = 0; i < obj.pageItems.length; i++) {
            collectTextFramesRecursive(obj.pageItems[i], textFrames);
        }
    }
}

var doc = app.activeDocument;
var unitLabel;
switch (doc.rulerUnits) {
    case RulerUnits.Millimeters:
        unitLabel = "mm";
        break;
    case RulerUnits.Centimeters:
        unitLabel = "cm";
        break;
    case RulerUnits.Inches:
        unitLabel = "inch";
        break;
    case RulerUnits.Pixels:
        unitLabel = "px";
        break;
    case RulerUnits.Picas:
        unitLabel = "pica";
        break;
    default:
        unitLabel = "pt";
        break;
}

main();

function showFontChoiceDialog(fontMap, originalPositions) {
    var lang = getCurrentLang();
    var dialog = new Window("dialog", LABELS[lang].title);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var radioGroup = dialog.add("panel", undefined, LABELS[lang].sortGroup);
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";
    radioGroup.margins = [10, 20, 10, 10];

    var radioButtons = [];
    var fontKeys = [];
    for (var name in fontMap) {
        fontKeys.push(name);
    }
    fontKeys.sort();

    for (var i = 0; i < fontKeys.length; i++) {
        var name = fontKeys[i];
        var values = [];
        if (fontMap[name] instanceof Array && fontMap[name].length > 0) {
            for (var v = 0; v < fontMap[name].length; v++) {
                values.push(fontMap[name][v].value);
            }
            values.sort(function(a, b) {
                return a - b;
            });
        }
        var label = (values.length > 0) ?
            (values.length > 3 ? values.slice(0, 3).join(", ") + "…" : values.join(", ")) :
            "";
        var rb = radioGroup.add("radiobutton", undefined, label);
        radioButtons.push({
            button: rb,
            key: name
        });
    }
    if (radioButtons.length > 0) {
        radioButtons[0].button.value = true;
    }

    var sortPanel = dialog.add("panel", undefined);
    sortPanel.orientation = "row";
    sortPanel.alignChildren = "left";
    sortPanel.margins = [10, 20, 10, 10];

    var ascRadio = sortPanel.add("radiobutton", undefined, LABELS[lang].asc);
    var descRadio = sortPanel.add("radiobutton", undefined, LABELS[lang].desc);
    var randomRadio = sortPanel.add("radiobutton", undefined, "ランダム");
    ascRadio.value = true;

    var spacingPanel = dialog.add("panel", undefined, LABELS[lang].custom === "指定" || LABELS[lang].custom === "Custom" ? LABELS[lang].custom : "間隔");
    spacingPanel.orientation = "row";
    spacingPanel.alignChildren = "left";
    spacingPanel.margins = [10, 20, 10, 10];

    var fitRadio = spacingPanel.add("radiobutton", undefined, LABELS[lang].fit);
    var customRadio = spacingPanel.add("radiobutton", undefined, LABELS[lang].custom);
    var defaultSpacing = (unitLabel === "mm") ? "1" : "20";
    var spacingInput = spacingPanel.add("edittext", undefined, defaultSpacing);
    spacingInput.characters = 5;
    spacingInput.enabled = false;
    var spacingUnit = spacingPanel.add("statictext", undefined, unitLabel);
    fitRadio.value = true;

    customRadio.onClick = function() {
        spacingInput.enabled = true;
    };
    fitRadio.onClick = function() {
        spacingInput.enabled = false;
    };

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, "キャンセル");
    var okBtn = buttonGroup.add("button", undefined, LABELS[lang].ok, {
        name: "ok"
    });
    cancelBtn.alignment = "left";
    okBtn.alignment = "right";

    var result = null;
    okBtn.onClick = function() {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].button.value) {
                var parsed = parseFloat(spacingInput.text);
                var factor = 1;
                if (unitLabel === "mm") factor = 2.83464567;
                else if (unitLabel === "cm") factor = 28.3464567;
                else if (unitLabel === "inch") factor = 72;
                else if (unitLabel === "pica") factor = 12;
                var referenceValue = parsed * factor;
                result = {
                    font: radioButtons[i].key,
                    descending: descRadio.value,
                    random: randomRadio.value,
                    spacingMode: fitRadio.value ? "fit" : "custom",
                    spacingValue: referenceValue
                };
                break;
            }
        }
        dialog.close();
    };
    cancelBtn.onClick = function() {
        dialog.close();
    };

    dialog.show();
    return result;
}

function collectTextWithFontInfoPerGroup(obj, result, groupRef, firstValueRef) {
    if (obj.typename === "TextFrame") {
        var text = obj.contents;
        var cleaned = text.replace(/,/g, "");
        var value = parseFloat(cleaned);
        var fontName = "不明";
        try {
            var font = obj.textRanges[0].characterAttributes.textFont;
            if (font && font.family && font.style) {
                fontName = font.family + " " + font.style;
            }
        } catch (e) {}
        if (!isNaN(value)) {
            if (!result[fontName]) result[fontName] = [];
            result[fontName].push({
                value: value,
                text: text,
                item: obj,
                group: groupRef
            });
            if (firstValueRef && typeof firstValueRef.value === "undefined") {
                firstValueRef.value = value;
            }
        }
    } else if (obj.typename === "GroupItem") {
        for (var i = 0; i < obj.pageItems.length; i++) {
            var child = obj.pageItems[i];
            collectTextWithFontInfoPerGroup(child, result, (child.typename === "GroupItem" ? child : groupRef), firstValueRef);
        }
    }
}