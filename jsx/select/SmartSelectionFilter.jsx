#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要

選択中のオブジェクトから、条件に合うテキストやパスだけを残して選択し直すIllustrator用スクリプトです。
ポイント文字／エリア内文字／パス上文字、行揃え、使用フォント、オープンパス、水平線／垂直線、クローズパスの塗り・線状態を条件にできます。
［まとめて選択］では、テキスト系、ケイのみ、塗りのみをチェックボックスでまとめて指定できます。
詳細条件を追加・解除すると、［まとめて選択］側のチェック状態も現在の条件に合わせて同期します。
対象外オブジェクトは、そのまま、非表示、不透明度指定から選べ、不透明度はスライダーで0〜100%に調整できます。
不透明度スライダーは、不透明度指定時のみ有効になります。
チェックボックスやスライダーの変更は、ダイアログを閉じずにカンバスへ反映されます。
アウトライン表示ボタンで表示モードを一時的に切り替えられ、終了時には元の表示に戻します。
OKで現在の選択状態を確定し、キャンセルでスクリプト開始前の選択状態と表示状態に戻します。
*/

// =========================================
// バージョンとローカライズ
// Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLanguage = getCurrentLang();

/* UI／エラーメッセージの日英ラベル定義 / Japanese-English UI and error message label definitions */
var LABELS = {
    dialogTitle: {
        ja: "選択オブジェクトの絞り込み",
        en: "Filter Selected Objects"
    },
    textPanel: {
        ja: "テキスト",
        en: "Text"
    },
    textObject: {
        ja: "ポイント文字",
        en: "Point Text"
    },
    areaText: {
        ja: "エリア内文字",
        en: "Area Text"
    },
    pathText: {
        ja: "パス上文字",
        en: "Path Text"
    },
    alignmentPanel: {
        ja: "行揃え",
        en: "Alignment"
    },
    fontPanel: {
        ja: "フォント",
        en: "Font"
    },
    noFontItems: {
        ja: "対象フォントなし",
        en: "No target fonts"
    },
    alignLeft: {
        ja: "左揃え",
        en: "Left"
    },
    alignCenter: {
        ja: "中央揃え",
        en: "Center"
    },
    alignRight: {
        ja: "右揃え",
        en: "Right"
    },
    pathPanel: {
        ja: "パス",
        en: "Path"
    },
    openPath: {
        ja: "オープンパス",
        en: "Open Path"
    },
    horizontalLine: {
        ja: "水平線",
        en: "Horizontal Line"
    },
    verticalLine: {
        ja: "垂直線",
        en: "Vertical Line"
    },
    closedPathPanel: {
        ja: "クローズパス",
        en: "Closed Path"
    },
    closedPathFillOnly: {
        ja: "塗りのみ",
        en: "Fill Only"
    },
    closedPathStrokeOnly: {
        ja: "線のみ",
        en: "Stroke Only"
    },
    closedPathFillAndStroke: {
        ja: "塗りと線",
        en: "Fill and Stroke"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    btnOutlineOn: {
        ja: "アウトライン表示",
        en: "Outline"
    },
    btnOutlineOff: {
        ja: "プレビュー表示",
        en: "Preview"
    },
    nonSelectedPanel: {
        ja: "対象外オブジェクト",
        en: "Non-target Objects"
    },
    nonSelectedNone: {
        ja: "何もしない",
        en: "Do Nothing"
    },
    nonSelectedHide: {
        ja: "非表示",
        en: "Hide"
    },
    nonSelectedOpacity: {
        ja: "不透明度",
        en: "Opacity"
    },
    simplePanel: {
        ja: "まとめて選択",
        en: "Quick Select"
    },
    simpleText: {
        ja: "テキスト",
        en: "Text"
    },
    simpleStrokeOnly: {
        ja: "ケイのみ",
        en: "Stroke Only"
    },
    simpleFillOnlyPath: {
        ja: "塗りのみ",
        en: "Fill Only"
    },
    noDocumentError: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    noSelectionError: {
        ja: "オブジェクトが選択されていません。",
        en: "No object is selected."
    }
};


function L(key) {
    if (!LABELS[key]) return key;
    var entry = LABELS[key];
    return entry[currentLanguage] || entry.en || key;
}

/* エラー表示ヘルパー / Error display helper */
function showError(key) {
    alert(L(key));
}

/* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
function labelWithCount(key, count) {
    if (currentLanguage === "ja") {
        return L(key) + "（" + count + "）";
    }
    return L(key) + " (" + count + ")";
}

/* 値付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with value (full-width JA parentheses, half-width EN parentheses) */
function labelWithValue(key, valueText) {
    if (currentLanguage === "ja") {
        return L(key) + "（" + valueText + "）";
    }
    return L(key) + " (" + valueText + ")";
}

/* 任意テキストへの値付け（日本語は全角括弧、英語は半角括弧）/ Append value to arbitrary text (full-width JA parentheses, half-width EN parentheses) */
function appendValueText(baseText, valueText) {
    if (currentLanguage === "ja") {
        return baseText + "（" + valueText + "）";
    }
    return baseText + " (" + valueText + ")";
}

/* パネル共通設定 / Common panel setup */
var PANEL_MARGINS = [15, 20, 15, 10];
var LINE_ORIENTATION_EPSILON = 0.01;
var DEFAULT_NON_SELECTED_OPACITY = 30;

function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    if (typeof spacing === "number") {
        panel.spacing = spacing;
    }
}

/* 選択を配列として退避 / Store selection as an array */
function copySelection(selection) {
    var copiedSelection = [];
    for (var selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
        copiedSelection.push(selection[selectionIndex]);
    }
    return copiedSelection;
}

/* 退避した選択状態を復元 / Restore stored selection */
function restoreSelection(items) {
    app.selection = null;
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        try {
            items[itemIndex].selected = true;
        } catch (e) { }
    }
}

/* チェックボックス変更時の処理を設定 / Bind checkbox update handlers */
function bindCheckboxPreviewUpdates(checkboxes, onUpdate) {
    function bindOneCheckbox(sourceCheckbox) {
        sourceCheckbox._optionClickRequested = false;

        sourceCheckbox.addEventListener("mousedown", function (event) {
            sourceCheckbox._optionClickRequested = event.altKey === true;
        });

        sourceCheckbox.onClick = function () {
            if (sourceCheckbox._optionClickRequested) {
                for (var checkboxIndex = 0; checkboxIndex < checkboxes.length; checkboxIndex++) {
                    checkboxes[checkboxIndex].value = sourceCheckbox.value;
                    checkboxes[checkboxIndex]._optionClickRequested = false;
                }
            }

            sourceCheckbox._optionClickRequested = false;
            onUpdate();
        };
    }

    for (var checkboxIndex = 0; checkboxIndex < checkboxes.length; checkboxIndex++) {
        bindOneCheckbox(checkboxes[checkboxIndex]);
    }
}

/* ポイント文字かどうかを判定 / Check whether a text frame is point text */
function isPointTextFrame(textFrame) {
    try {
        return textFrame.kind === TextType.POINTTEXT;
    } catch (e) {
        return false;
    }
}

/* エリア内文字かどうかを判定 / Check whether a text frame is area text */
function isAreaTextFrame(textFrame) {
    try {
        return textFrame.kind === TextType.AREATEXT;
    } catch (e) {
        return false;
    }
}

/* パス上文字かどうかを判定 / Check whether a text frame is path text */
function isPathTextFrame(textFrame) {
    try {
        return textFrame.kind === TextType.PATHTEXT;
    } catch (e) {
        return false;
    }
}

/* テキストの行揃えを取得 / Get text alignment */
function getTextAlignmentType(textFrame) {
    try {
        var justification = textFrame.textRange.paragraphAttributes.justification;
        if (justification === Justification.LEFT) {
            return "left";
        }
        if (justification === Justification.CENTER) {
            return "center";
        }
        if (justification === Justification.RIGHT) {
            return "right";
        }
    } catch (e) { }

    return "other";
}

/* テキストのフォント名を取得 / Get text font display name (family + style) */
function getTextFontName(textFrame) {
    try {
        var font = textFrame.textRange.characterAttributes.textFont;
        var family = "";
        var style = "";
        try { family = font.family || ""; } catch (e1) { }
        try { style = font.style || ""; } catch (e2) { }
        if (family && style) return family + " " + style;
        if (family) return family;
        try { return font.name; } catch (e3) { }
    } catch (e) { }

    return "";
}

/* フォント件数を追加 / Add font count */
function addFontCount(fontCounts, fontName) {
    if (!fontName) {
        return;
    }
    if (!fontCounts[fontName]) {
        fontCounts[fontName] = 0;
    }
    fontCounts[fontName]++;
}

/* フォント名をソートして配列化 / Sort font names into an array */
function getSortedFontNames(fontCounts) {
    var fontNames = [];
    for (var fontName in fontCounts) {
        if (fontCounts.hasOwnProperty(fontName)) {
            fontNames.push(fontName);
        }
    }
    fontNames.sort();
    return fontNames;
}

/* オープンパスの線方向を判定 / Detect line orientation of an open path
   - "horizontal": geometricBounds の高さ ≈ 0
   - "vertical":   geometricBounds の幅 ≈ 0
   - "other":      それ以外
*/
function getOpenPathLineOrientation(pathItem) {
    try {
        var geometricBounds = pathItem.geometricBounds; // [left, top, right, bottom]
        if (!geometricBounds || geometricBounds.length < 4) return "other";
        var width = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var height = Math.abs(geometricBounds[1] - geometricBounds[3]);
        if (height <= LINE_ORIENTATION_EPSILON && width > LINE_ORIENTATION_EPSILON) return "horizontal";
        if (width <= LINE_ORIENTATION_EPSILON && height > LINE_ORIENTATION_EPSILON) return "vertical";
    } catch (e) { }
    return "other";
}

/* クローズパスの塗りと線の状態を取得 / Get fill and stroke state of a closed path */
function getClosedPathPaintType(pathItem) {
    var hasFill = false;
    var hasStroke = false;

    try {
        hasFill = pathItem.filled === true;
    } catch (e1) { }

    try {
        hasStroke = pathItem.stroked === true;
    } catch (e2) { }

    if (hasFill && hasStroke) {
        return "fillAndStroke";
    }
    if (hasFill) {
        return "fillOnly";
    }
    if (hasStroke) {
        return "strokeOnly";
    }

    return "none";
}

/* 選択内の各種オブジェクト数をカウント / Count object types in selection */
function countSelectionItems(selection) {
    var counts = {
        text: 0,
        areaText: 0,
        pathText: 0,
        alignLeft: 0,
        alignCenter: 0,
        alignRight: 0,
        fonts: {},
        openPath: 0,
        horizontalLine: 0,
        verticalLine: 0,
        closedPathFillOnly: 0,
        closedPathStrokeOnly: 0,
        closedPathFillAndStroke: 0
    };

    countItemsRecursive(selection, counts);
    return counts;
}

/* 選択内の子要素も含めて再帰的にカウント / Count selected items recursively including children */
function countItemsRecursive(items, counts) {
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var item = items[itemIndex];

        if (item.typename === "TextFrame") {
            if (isPointTextFrame(item)) {
                counts.text++;
            } else if (isAreaTextFrame(item)) {
                counts.areaText++;
            } else if (isPathTextFrame(item)) {
                counts.pathText++;
            }

            var alignmentType = getTextAlignmentType(item);
            if (alignmentType === "left") {
                counts.alignLeft++;
            } else if (alignmentType === "center") {
                counts.alignCenter++;
            } else if (alignmentType === "right") {
                counts.alignRight++;
            }

            addFontCount(counts.fonts, getTextFontName(item));
        } else if (item.typename === "PathItem") {
            if (item.closed === true) {
                var paintType = getClosedPathPaintType(item);
                if (paintType === "fillOnly") {
                    counts.closedPathFillOnly++;
                } else if (paintType === "strokeOnly") {
                    counts.closedPathStrokeOnly++;
                } else if (paintType === "fillAndStroke") {
                    counts.closedPathFillAndStroke++;
                }
            } else {
                counts.openPath++;
                var orientation = getOpenPathLineOrientation(item);
                if (orientation === "horizontal") {
                    counts.horizontalLine++;
                } else if (orientation === "vertical") {
                    counts.verticalLine++;
                }
            }
        } else if (item.typename === "GroupItem") {
            countItemsRecursive(item.pageItems, counts);
        }
    }
}

/* UIから選択条件を取得 / Read filter options from UI */
function readFilterOptions(dialogUi) {
    return {
        keepText: dialogUi.cbText.value,
        keepAreaText: dialogUi.cbAreaText.value,
        keepPathText: dialogUi.cbPathText.value,
        keepAlignLeft: dialogUi.cbAlignLeft.value,
        keepAlignCenter: dialogUi.cbAlignCenter.value,
        keepAlignRight: dialogUi.cbAlignRight.value,
        selectedFonts: readSelectedFonts(dialogUi.fontCheckboxes),
        keepOpenPath: dialogUi.cbOpenPath.value,
        keepHorizontalLine: dialogUi.cbHorizontalLine.value,
        keepVerticalLine: dialogUi.cbVerticalLine.value,
        keepClosedPathFillOnly: dialogUi.cbClosedPathFillOnly.value,
        keepClosedPathStrokeOnly: dialogUi.cbClosedPathStrokeOnly.value,
        keepClosedPathFillAndStroke: dialogUi.cbClosedPathFillAndStroke.value
    };
}

/* チェックされたフォント名を取得 / Read selected font names */
function readSelectedFonts(fontCheckboxes) {
    var selectedFonts = {};
    for (var fontCheckboxIndex = 0; fontCheckboxIndex < fontCheckboxes.length; fontCheckboxIndex++) {
        if (fontCheckboxes[fontCheckboxIndex].checkbox.value) {
            selectedFonts[fontCheckboxes[fontCheckboxIndex].fontName] = true;
        }
    }
    return selectedFonts;
}

/* 1項目が残す対象かどうかを判定 / Check whether an item should remain selected */
function shouldKeepItem(item, filterOptions) {
    if (item.typename === "TextFrame") {
        var keepByTextType = false;
        if (isPointTextFrame(item)) {
            keepByTextType = filterOptions.keepText;
        } else if (isAreaTextFrame(item)) {
            keepByTextType = filterOptions.keepAreaText;
        } else if (isPathTextFrame(item)) {
            keepByTextType = filterOptions.keepPathText;
        }

        if (!keepByTextType) {
            return false;
        }

        if (filterOptions.keepAlignLeft || filterOptions.keepAlignCenter || filterOptions.keepAlignRight) {
            var alignmentType = getTextAlignmentType(item);
            if (alignmentType === "left" && !filterOptions.keepAlignLeft) {
                return false;
            }
            if (alignmentType === "center" && !filterOptions.keepAlignCenter) {
                return false;
            }
            if (alignmentType === "right" && !filterOptions.keepAlignRight) {
                return false;
            }
            if (alignmentType !== "left" && alignmentType !== "center" && alignmentType !== "right") {
                return false;
            }
        }

        if (hasSelectedFonts(filterOptions.selectedFonts)) {
            return filterOptions.selectedFonts[getTextFontName(item)] === true;
        }

        return true;
    }

    if (item.typename === "PathItem") {
        if (item.closed === true) {
            var paintType = getClosedPathPaintType(item);
            if (paintType === "fillOnly") {
                return filterOptions.keepClosedPathFillOnly;
            }
            if (paintType === "strokeOnly") {
                return filterOptions.keepClosedPathStrokeOnly;
            }
            if (paintType === "fillAndStroke") {
                return filterOptions.keepClosedPathFillAndStroke;
            }
            return false;
        }
        var orientation = getOpenPathLineOrientation(item);
        if (orientation === "horizontal") {
            return filterOptions.keepOpenPath || filterOptions.keepHorizontalLine;
        }
        if (orientation === "vertical") {
            return filterOptions.keepOpenPath || filterOptions.keepVerticalLine;
        }
        return filterOptions.keepOpenPath;
    }

    return false;
}

/* フォント指定があるかどうかを判定 / Check whether any font is selected */
function hasSelectedFonts(selectedFonts) {
    for (var fontName in selectedFonts) {
        if (selectedFonts.hasOwnProperty(fontName)) {
            return true;
        }
    }
    return false;
}

/* 選択状態を条件に合わせて反映 / Apply filter options to selection */
function applyFilterToSelection(items, filterOptions) {
    app.selection = null;
    applyFilterToItemsRecursive(items, filterOptions);
}

/* 子要素も含めて選択状態を条件に合わせて反映 / Apply filter options recursively including children */
function applyFilterToItemsRecursive(items, filterOptions) {
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var item = items[itemIndex];

        try {
            item.selected = shouldKeepItem(item, filterOptions);
        } catch (e1) { }

        if (item.typename === "GroupItem") {
            try {
                applyFilterToItemsRecursive(item.pageItems, filterOptions);
            } catch (e2) { }
        }
    }
}

/* 行揃え/フォントパネルの有効状態を更新 / Update alignment & font panel enabled state */
function updateAlignmentPanelEnabled(dialogUi) {
    var hasTextTarget = dialogUi.cbText.value || dialogUi.cbAreaText.value || dialogUi.cbPathText.value;
    dialogUi.alignmentPanel.enabled = hasTextTarget;
    if (dialogUi.fontPanel) dialogUi.fontPanel.enabled = hasTextTarget;
}

/* 元の表示状態を退避 / Capture original visibility/opacity for items (recursive) */
function captureVisualState(items, snapshot) {
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var item = items[itemIndex];
        var entry = { item: item };
        try { entry.hidden = item.hidden; } catch (e1) { entry.hidden = false; }
        try { entry.opacity = item.opacity; } catch (e2) { entry.opacity = 100; }
        snapshot.push(entry);
        if (item.typename === "GroupItem") {
            try { captureVisualState(item.pageItems, snapshot); } catch (e3) { }
        }
    }
}

/* 退避した表示状態を復元 / Restore captured visibility/opacity */
function restoreVisualState(snapshot) {
    for (var snapshotIndex = 0; snapshotIndex < snapshot.length; snapshotIndex++) {
        var visualState = snapshot[snapshotIndex];
        try { visualState.item.hidden = visualState.hidden; } catch (e1) { }
        try { visualState.item.opacity = visualState.opacity; } catch (e2) { }
    }
}

/* 非選択オブジェクトに挙動を適用 / Apply behavior to non-selected items
   mode: "none" | "hide" | "opacity"
   opacityValue: number (0-100)
*/
function applyNonSelectedBehavior(items, mode, opacityValue) {
    if (mode === "none") return;
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var item = items[itemIndex];
        var isSelected = false;
        try { isSelected = item.selected === true; } catch (e0) { }
        if (!isSelected) {
            try {
                if (mode === "hide") {
                    item.hidden = true;
                } else if (mode === "opacity") {
                    item.opacity = opacityValue;
                }
            } catch (e) { }
        }
        if (item.typename === "GroupItem") {
            try { applyNonSelectedBehavior(item.pageItems, mode, opacityValue); } catch (e2) { }
        }
    }
}

function showFilterDialog(originalSelection) {
    var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var counts = countSelectionItems(originalSelection);

    /* 元の表示状態スナップショット / Snapshot original visual state */
    var visualSnapshot = [];
    captureVisualState(originalSelection, visualSnapshot);

    /* 上部：選択オブジェクト以外の挙動 / Top: behavior for non-selected items */
    var nonSelectedPanel = dialog.add("panel", undefined, L("nonSelectedPanel"));
    nonSelectedPanel.orientation = "column";
    nonSelectedPanel.alignChildren = ["left", "center"];
    nonSelectedPanel.margins = PANEL_MARGINS;
    nonSelectedPanel.spacing = 6;

    var nonSelectedRadioGroup = nonSelectedPanel.add("group");
    nonSelectedRadioGroup.orientation = "row";
    nonSelectedRadioGroup.alignChildren = ["left", "center"];
    nonSelectedRadioGroup.spacing = 12;

    var rbNonSelNone = nonSelectedRadioGroup.add("radiobutton", undefined, L("nonSelectedNone"));
    var rbNonSelHide = nonSelectedRadioGroup.add("radiobutton", undefined, L("nonSelectedHide"));
    var rbNonSelOpacity = nonSelectedRadioGroup.add("radiobutton", undefined, labelWithValue("nonSelectedOpacity", DEFAULT_NON_SELECTED_OPACITY + "%"));

    var nonSelectedOpacitySliderGroup = nonSelectedPanel.add("group");
    nonSelectedOpacitySliderGroup.orientation = "row";
    nonSelectedOpacitySliderGroup.alignChildren = ["center", "center"];
    nonSelectedOpacitySliderGroup.alignment = ["fill", "center"];

    var nonSelectedOpacitySlider = nonSelectedOpacitySliderGroup.add("slider", undefined, DEFAULT_NON_SELECTED_OPACITY, 0, 100);
    nonSelectedOpacitySlider.preferredSize.width = 220;
    rbNonSelNone.value = true;

    function getNonSelectedMode() {
        if (rbNonSelHide.value) return "hide";
        if (rbNonSelOpacity.value) return "opacity";
        return "none";
    }

    function getNonSelectedOpacityValue() {
        return Math.max(0, Math.min(100, Math.round(nonSelectedOpacitySlider.value)));
    }

    function updateNonSelectedOpacityDisplay() {
        rbNonSelOpacity.text = labelWithValue("nonSelectedOpacity", getNonSelectedOpacityValue() + "%");
    }

    function updateNonSelectedOpacitySliderEnabled() {
        nonSelectedOpacitySlider.enabled = rbNonSelOpacity.value === true;
    }

    /* 簡易モード：横並びチェックボックス / Quick mode: horizontal checkboxes */
    var simplePanel = dialog.add("panel", undefined, L("simplePanel"));
    simplePanel.orientation = "row";
    simplePanel.alignChildren = ["left", "center"];
    simplePanel.margins = PANEL_MARGINS;
    simplePanel.spacing = 12;
    var cbSimpleText = simplePanel.add("checkbox", undefined, L("simpleText"));
    var cbSimpleStrokeOnly = simplePanel.add("checkbox", undefined, L("simpleStrokeOnly"));
    var cbSimpleFillOnlyPath = simplePanel.add("checkbox", undefined, L("simpleFillOnlyPath"));

    var columnGroup = dialog.add("group");
    columnGroup.orientation = "row";
    columnGroup.alignChildren = ["fill", "top"];
    columnGroup.alignment = "fill";

    var leftColumn = columnGroup.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = "fill";
    leftColumn.alignment = ["fill", "top"];

    var rightColumn = columnGroup.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = "fill";
    rightColumn.alignment = ["fill", "top"];

    var textPanel = leftColumn.add("panel", undefined, L("textPanel"));
    setupPanel(textPanel, 6);
    var cbText = textPanel.add("checkbox", undefined, labelWithCount("textObject", counts.text));
    var cbAreaText = textPanel.add("checkbox", undefined, labelWithCount("areaText", counts.areaText));
    var cbPathText = textPanel.add("checkbox", undefined, labelWithCount("pathText", counts.pathText));

    var alignmentPanel = leftColumn.add("panel", undefined, L("alignmentPanel"));
    setupPanel(alignmentPanel, 6);
    var cbAlignLeft = alignmentPanel.add("checkbox", undefined, labelWithCount("alignLeft", counts.alignLeft));
    var cbAlignCenter = alignmentPanel.add("checkbox", undefined, labelWithCount("alignCenter", counts.alignCenter));
    var cbAlignRight = alignmentPanel.add("checkbox", undefined, labelWithCount("alignRight", counts.alignRight));

    var pathPanel = rightColumn.add("panel", undefined, L("pathPanel"));
    setupPanel(pathPanel, 6);
    var cbOpenPath = pathPanel.add("checkbox", undefined, labelWithCount("openPath", counts.openPath));
    var cbHorizontalLine = pathPanel.add("checkbox", undefined, labelWithCount("horizontalLine", counts.horizontalLine));
    var cbVerticalLine = pathPanel.add("checkbox", undefined, labelWithCount("verticalLine", counts.verticalLine));

    var closedPathPanel = rightColumn.add("panel", undefined, L("closedPathPanel"));
    setupPanel(closedPathPanel, 6);
    var cbClosedPathFillOnly = closedPathPanel.add("checkbox", undefined, labelWithCount("closedPathFillOnly", counts.closedPathFillOnly));
    var cbClosedPathStrokeOnly = closedPathPanel.add("checkbox", undefined, labelWithCount("closedPathStrokeOnly", counts.closedPathStrokeOnly));
    var cbClosedPathFillAndStroke = closedPathPanel.add("checkbox", undefined, labelWithCount("closedPathFillAndStroke", counts.closedPathFillAndStroke));

    /* フォントパネル：カラム貫通（全幅） / Font panel spanning across columns */
    var fontPanel = dialog.add("panel", undefined, L("fontPanel"));
    setupPanel(fontPanel, 6);
    var fontCheckboxes = [];
    var fontNames = getSortedFontNames(counts.fonts);
    for (var fontIndex = 0; fontIndex < fontNames.length; fontIndex++) {
        var fontName = fontNames[fontIndex];
        var fontCheckboxLabel = appendValueText(fontName, counts.fonts[fontName]);
        var fontCheckbox = fontPanel.add("checkbox", undefined, fontCheckboxLabel);
        fontCheckboxes.push({
            fontName: fontName,
            checkbox: fontCheckbox
        });
    }

    var dialogUi = {
        cbText: cbText,
        cbAreaText: cbAreaText,
        cbPathText: cbPathText,
        alignmentPanel: alignmentPanel,
        cbAlignLeft: cbAlignLeft,
        cbAlignCenter: cbAlignCenter,
        cbAlignRight: cbAlignRight,
        fontPanel: fontPanel,
        fontCheckboxes: fontCheckboxes,
        cbOpenPath: cbOpenPath,
        cbHorizontalLine: cbHorizontalLine,
        cbVerticalLine: cbVerticalLine,
        cbClosedPathFillOnly: cbClosedPathFillOnly,
        cbClosedPathStrokeOnly: cbClosedPathStrokeOnly,
        cbClosedPathFillAndStroke: cbClosedPathFillAndStroke
    };

    var filterCheckboxes = [
        cbText,
        cbAreaText,
        cbPathText,
        cbAlignLeft,
        cbAlignCenter,
        cbAlignRight,
        cbOpenPath,
        cbHorizontalLine,
        cbVerticalLine,
        cbClosedPathFillOnly,
        cbClosedPathStrokeOnly,
        cbClosedPathFillAndStroke
    ];

    for (var fontCheckboxIndex = 0; fontCheckboxIndex < fontCheckboxes.length; fontCheckboxIndex++) {
        filterCheckboxes.push(fontCheckboxes[fontCheckboxIndex].checkbox);
    }

    /* デフォルトでは全てチェックを外す / Uncheck all options by default */
    for (var checkboxIndex = 0; checkboxIndex < filterCheckboxes.length; checkboxIndex++) {
        filterCheckboxes[checkboxIndex].value = false;
    }

    /* テキスト系をまとめてデフォルトON / Enable all text-related options by default */
    cbSimpleText.value = true;
    cbText.value = true;
    cbAreaText.value = true;
    cbPathText.value = true;

    function updateSimpleCheckboxesFromDetailedSelection() {
        cbSimpleText.value = cbText.value === true && cbAreaText.value === true && cbPathText.value === true;
        cbSimpleStrokeOnly.value = cbOpenPath.value === true && cbClosedPathStrokeOnly.value === true;
        cbSimpleFillOnlyPath.value = cbClosedPathFillOnly.value === true;
    }

    function updateCanvasSelection() {
        updateAlignmentPanelEnabled(dialogUi);
        updateSimpleCheckboxesFromDetailedSelection();
        restoreVisualState(visualSnapshot);
        applyFilterToSelection(originalSelection, readFilterOptions(dialogUi));
        applyNonSelectedBehavior(originalSelection, getNonSelectedMode(), getNonSelectedOpacityValue());
        try {
            app.redraw();
        } catch (e) { }
    }

    rbNonSelNone.onClick = function () {
        updateNonSelectedOpacitySliderEnabled();
        updateCanvasSelection();
    };
    rbNonSelHide.onClick = function () {
        updateNonSelectedOpacitySliderEnabled();
        updateCanvasSelection();
    };
    rbNonSelOpacity.onClick = function () {
        updateNonSelectedOpacitySliderEnabled();
        updateCanvasSelection();
    };
    nonSelectedOpacitySlider.onChanging = function () {
        rbNonSelOpacity.value = true;
        updateNonSelectedOpacityDisplay();
        updateNonSelectedOpacitySliderEnabled();
        updateCanvasSelection();
    };
    nonSelectedOpacitySlider.onChange = function () {
        rbNonSelOpacity.value = true;
        updateNonSelectedOpacityDisplay();
        updateNonSelectedOpacitySliderEnabled();
        updateCanvasSelection();
    };


    cbSimpleText.onClick = function () {
        cbText.value = cbSimpleText.value;
        cbAreaText.value = cbSimpleText.value;
        cbPathText.value = cbSimpleText.value;
        updateCanvasSelection();
    };
    cbSimpleStrokeOnly.onClick = function () {
        cbOpenPath.value = cbSimpleStrokeOnly.value;
        cbClosedPathStrokeOnly.value = cbSimpleStrokeOnly.value;
        updateCanvasSelection();
    };
    cbSimpleFillOnlyPath.onClick = function () {
        cbClosedPathFillOnly.value = cbSimpleFillOnlyPath.value;
        updateCanvasSelection();
    };

    bindCheckboxPreviewUpdates(filterCheckboxes, updateCanvasSelection);

    /* ボタングループを3カラムで作成 / Create three-column button group */
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignChildren = ["fill", "center"];
    buttonGroup.alignment = "fill";
    buttonGroup.margins = [0, 10, 0, 0];

    var leftButtonGroup = buttonGroup.add("group");
    leftButtonGroup.orientation = "row";
    leftButtonGroup.alignChildren = ["left", "center"];
    leftButtonGroup.alignment = ["left", "center"];

    var isOutlineMode = false;
    var previewButton = leftButtonGroup.add("button", undefined, L("btnOutlineOn"));
    previewButton.onClick = function () {
        try {
            app.executeMenuCommand("preview");
            isOutlineMode = !isOutlineMode;
            previewButton.text = isOutlineMode ? L("btnOutlineOff") : L("btnOutlineOn");
        } catch (e) { }
    };

    var spacer = buttonGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var rightButtonGroup = buttonGroup.add("group");
    rightButtonGroup.orientation = "row";
    rightButtonGroup.alignChildren = ["right", "center"];
    rightButtonGroup.alignment = ["right", "center"];
    rightButtonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    rightButtonGroup.add("button", undefined, L("ok"), { name: "ok" });

    updateNonSelectedOpacitySliderEnabled();
    updateCanvasSelection();

    var dialogResult = dialog.show();

    if (isOutlineMode) {
        try {
            app.executeMenuCommand("preview");
        } catch (e) { }
    }

    restoreVisualState(visualSnapshot);

    if (dialogResult !== 1) {
        restoreSelection(originalSelection);
        return null;
    }

    return readFilterOptions(dialogUi);
}

(function () {
    if (app.documents.length === 0) {
        showError("noDocumentError");
        return;
    }

    var doc = app.activeDocument;
    var originalSelection = copySelection(doc.selection);

    if (originalSelection.length === 0) {
        showError("noSelectionError");
        return;
    }

    showFilterDialog(originalSelection);
})();