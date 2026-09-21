#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

グループ内のテキストから数値を抽出し、その数値でグループを並び替えて縦方向に整列します。
フォント情報によるグループ分けと、昇順・降順・ランダム順に対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortByNumbers.md

### Overview

Extracts numbers from the text inside groups, sorts the groups by those numbers, and aligns them vertically.
Groups can be split by font, and the order can be ascending, descending or random.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortByNumbers.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SortByNumbers";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortByNumbers.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortByNumbers.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var LABELS = {
        dialog: {
            title: { ja: "グループの数値で整列", en: "Align Groups by Number" }
        },
        panel: {
            sortGroup: { ja: "数値グループ", en: "Number Group" },
            spacing:   { ja: "間隔", en: "Spacing" }
        },
        radio: {
            asc:    { ja: "昇順", en: "Ascending" },
            desc:   { ja: "降順", en: "Descending" },
            random: { ja: "ランダム", en: "Random" },
            fit:    { ja: "ぴったり", en: "Fit" },
            custom: { ja: "指定", en: "Custom" }
        },
        tooltip: {
            sortGroup: { ja: "並べ替えの基準に使う数値のまとまりを選びます。", en: "Which set of numbers to sort by." },
            asc:       { ja: "数値の小さい順に並べます。", en: "Sorts from the smallest number up." },
            desc:      { ja: "数値の大きい順に並べます。", en: "Sorts from the largest number down." },
            random:    { ja: "数値と関係なく、順序をシャッフルします。", en: "Shuffles the order regardless of the numbers." },
            fit:       { ja: "現在の並びの間隔を保ったまま詰め直します。", en: "Keeps the current spacing and repacks the objects." },
            custom:    { ja: "間隔を数値で指定します。", en: "Sets the spacing to a value you type." },
            spacingInput: { ja: "オブジェクト間にあける間隔です。「指定」を選んだときだけ使われます。", en: "Gap left between objects. Used only when Custom is selected." }
        },
        button: {
            ok:     { ja: "ソート", en: "Sort" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.spacing" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    function getCurrentLang() {
        return ($.locale === "ja" || $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var currentSelection = app.activeDocument.selection;
        if (currentSelection.length === 0) {
            alert("グループオブジェクトを選択してください。");
            return;
        }

        var validGroups = [];
        for (var i = 0; i < currentSelection.length; i++) {
            if (currentSelection[i].typename === "GroupItem") {
                var dummyMap = {};
                var firstValue = { value: undefined };
                collectTextWithFontInfoPerGroup(currentSelection[i], dummyMap, currentSelection[i], firstValue);
                if (!isNaN(firstValue.value)) {
                    validGroups.push(currentSelection[i]);
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
        var uiLang = getCurrentLang();
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        var radioGroup = dialog.add("panel", undefined, getLabel("panel.sortGroup"));
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
            rb.helpTip = getLabel("tooltip.sortGroup");
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

        var ascRadio = sortPanel.add("radiobutton", undefined, getLabel("radio.asc"));
        ascRadio.helpTip = getLabel("tooltip.asc");
        var descRadio = sortPanel.add("radiobutton", undefined, getLabel("radio.desc"));
        descRadio.helpTip = getLabel("tooltip.desc");
        var randomRadio = sortPanel.add("radiobutton", undefined, getLabel("radio.random"));
        randomRadio.helpTip = getLabel("tooltip.random");
        ascRadio.value = true;

        var spacingPanel = dialog.add("panel", undefined, getLabel("panel.spacing"));
        spacingPanel.orientation = "row";
        spacingPanel.alignChildren = "left";
        spacingPanel.margins = [10, 20, 10, 10];

        var fitRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.fit"));
        fitRadio.helpTip = getLabel("tooltip.fit");
        var customRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.custom"));
        customRadio.helpTip = getLabel("tooltip.custom");
        var defaultSpacing = (unitLabel === "mm") ? "1" : "20";
        var spacingInput = spacingPanel.add("edittext", undefined, defaultSpacing);
        spacingInput.helpTip = getLabel("tooltip.spacingInput");
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
        var cancelBtn = buttonGroup.add("button", undefined, getLabel("button.cancel"));
        var okBtn = buttonGroup.add("button", undefined, getLabel("button.ok"), {
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

})();
