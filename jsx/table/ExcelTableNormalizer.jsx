#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 概要
 -----------------------------------------------------------------------------
 Excel由来のIllustratorデータを、表組みとして扱いやすい状態に整形するスクリプトです。

 主な処理:
 - 実行時にカラーモード関連のメニューコマンドを実行し、以後の色指定を安定させます。
 - クリッピングマスクを解除し、エラーインジケーターなどの小さいマークなどの不要なオブジェクトを削除します。
 - テキストを専用レイヤー（_text_all）へ移動し、重複テキストの削除や印刷用黒への変換を行います。
 - セル背景を専用レイヤー（_cell_rectangle）へ抽出・調整し、必要に応じて高さを均等化します。
 - 長方形状の罫線を中心線化し、配置モードに応じて均等配置・結合セル対応・列幅保持を行います。
 - 外枠の長方形化、線幅統一、罫線の印刷用黒への変換に対応します。
 - 線幅入力はIllustratorの環境設定の［線］を参照し、mm環境では0.1mm、pt環境では0.25ptを初期値にします。

 Overview
 -----------------------------------------------------------------------------
 This script formats Illustrator data imported or copied from Excel so it can be handled like a table.

 Main features:
 - Runs the color-related menu command at startup to stabilize subsequent color settings.
 - Releases clipping masks and removes unnecessary small marks.
 - Moves text to the dedicated layer (_text_all), removes duplicate text, and converts text to print black.
 - Extracts and adjusts cell backgrounds on the dedicated layer (_cell_rectangle), and equalizes their heights when needed.
 - Converts rectangular rules to centerlines and applies uniform placement, merged-cell handling, or column-width preservation according to the placement mode.
 - Supports converting the outer frame to a rectangle, unifying stroke width, and setting rules to print black.
 - Stroke width input follows Illustrator strokeUnits, defaulting to 0.1 mm in mm environments and 0.25 pt in pt environments.

 更新日 / Updated: 2026-04-27
 */

(function () {
    var SCRIPT_VERSION = "v1.0.0";

    var TEXT_LAYER_NAME = "_text_all";
    var PATH_LAYER_NAME = "_cell_rectangle";

    var lang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var PANEL_MARGINS = [15, 20, 15, 10];

    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    /* 日英ラベル定義 / Define Japanese-English labels */
    var LABELS = {
        dialogTitle: {
            ja: "Excelデータを整形",
            en: "Format Excel Data"
        },
        releaseMask: { ja: "クリッピングマスクを解除", en: "Release clipping masks" },
        optionsPanelTitle: { ja: "オプション", en: "Options" },
        textPanelTitle: { ja: "テキスト", en: "Text" },
        cellBgPanelTitle: { ja: "セル背景", en: "Cell Background" },
        moveTextToLayer: {
            ja: "専用レイヤーへ移動",
            en: "Move text to dedicated layer"
        },
        setTextK100: {
            ja: "印刷用の「黒」に",
            en: "Set text to print black"
        },
        removeDuplicateTexts: {
            ja: "重複テキストを削除",
            en: "Remove duplicate texts"
        },
        removeSmallObjects: {
            ja: "小さいマークなどを削除",
            en: "Remove small marks"
        },
        adjustCellBackground: {
            ja: "セル背景を抽出・統合",
            en: "Extract and merge cell backgrounds"
        },
        equalizeHeights: {
            ja: "罫線にスナップ",
            en: "Snap to rules"
        },
        rulesPanelTitle: { ja: "罫線", en: "Rules" },
        stylePanelTitle: { ja: "スタイル", en: "Style" },
        centerline: { ja: "罫線を中心線化", en: "Convert rules to centerlines" },
        placementMode: { ja: "罫線の配置", en: "Rule placement" },
        placementUniformForced: { ja: "行・列を均等に配置", en: "Distribute rows and columns evenly" },
        placementUniformMerged: { ja: "結合セルを考慮して均等配置", en: "Distribute evenly with merged cells" },
        placementMamaIki: { ja: "列幅を保持", en: "Preserve column widths" },
        outerToRect: { ja: "外枠を長方形に変換", en: "Convert outer frame to rectangle" },
        strokeWidth: { ja: "線幅", en: "Stroke width" },
        rulesK100: { ja: "罫線を印刷用の「黒」に", en: "Set rules to print black" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." }
    };

    function L(key) {
        return LABELS[key][lang];
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (lang === 'ja' ? '：' : ':');
    }

    function safeCall(callback) {
        try {
            return callback();
        } catch (e) { }
        return null;
    }

    // 単位コードとラベルのマップ
    var unitLabelMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    // 現在の線単位ラベルを取得（strokeUnits）
    function getCurrentStrokeUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("strokeUnits");
        return unitLabelMap[unitCode] || "pt";
    }

    function getDefaultStrokeWidthText(unitLabel) {
        if (unitLabel === "pt") return "0.25";
        if (unitLabel === "mm") return "0.1";
        return "0.1";
    }

    function convertLengthToPoints(value, unitLabel) {
        switch (unitLabel) {
            case "mm": return value * 72 / 25.4;
            case "cm": return value * 72 / 2.54;
            case "in": return value * 72;
            case "pt": return value;
            case "px": return value; // Illustratorはpx≒pt
            case "Q/H": return value * 0.709; // 約
            default: return value;
        }
    }

    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is pressed
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減 / Change by 0.1 when Option is pressed
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め / Round to 1 decimal
                value = Math.round(value * 10) / 10;
            } else if (keyboard.shiftKey) {
                // Shiftキー押下時のみ整数に丸め / Round to integer only when Shift is pressed
                value = Math.round(value);
            }

            editText.text = String(value);
        });
    }

    /* メイン処理：UIを構築し実行する / Main entry: build UI and execute */
    function main() {
        safeCall(function () { app.executeMenuCommand('Colors8'); });
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";
        dlg.spacing = 10;
        dlg.margins = 20;

        var columnsGroup = dlg.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";
        columnsGroup.spacing = 10;

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";
        leftColumn.spacing = 10;

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";
        rightColumn.spacing = 10;

        var optionsPanel = leftColumn.add("panel", undefined, L('optionsPanelTitle'));
        setupPanel(optionsPanel);

        var releaseMaskCheckbox = optionsPanel.add("checkbox", undefined, L('releaseMask'));
        releaseMaskCheckbox.value = true;

        var removeSmallCheckbox = optionsPanel.add("checkbox", undefined, L('removeSmallObjects'));
        removeSmallCheckbox.value = true;

        var textPanel = leftColumn.add("panel", undefined, L('textPanelTitle'));
        setupPanel(textPanel);

        var moveTextCheckbox = textPanel.add("checkbox", undefined, L('moveTextToLayer'));
        moveTextCheckbox.value = true;

        var setK100Checkbox = textPanel.add("checkbox", undefined, L('setTextK100'));
        setK100Checkbox.value = true;

        var removeDuplicateCheckbox = textPanel.add("checkbox", undefined, L('removeDuplicateTexts'));
        removeDuplicateCheckbox.value = true;

        var cellBgPanel = leftColumn.add("panel", undefined, L('cellBgPanelTitle'));
        setupPanel(cellBgPanel);

        var adjustCellBgCheckbox = cellBgPanel.add("checkbox", undefined, L('adjustCellBackground'));
        adjustCellBgCheckbox.value = true;

        var equalizeHeightsCheckbox = cellBgPanel.add("checkbox", undefined, L('equalizeHeights'));
        equalizeHeightsCheckbox.value = true;

        adjustCellBgCheckbox.onClick = function () {
            equalizeHeightsCheckbox.enabled = adjustCellBgCheckbox.value;
        };

        var rulesPanel = rightColumn.add("panel", undefined, L('rulesPanelTitle'));
        setupPanel(rulesPanel);

        var centerlineCheckbox = rulesPanel.add("checkbox", undefined, L('centerline'));
        centerlineCheckbox.value = true;

        var outerToRectCheckbox = rulesPanel.add("checkbox", undefined, L('outerToRect'));
        outerToRectCheckbox.value = true;

        var placementPanel = rulesPanel.add("panel", undefined, L('placementMode'));
        setupPanel(placementPanel, 6);
        var placementUniformForcedRadio = placementPanel.add("radiobutton", undefined, L('placementUniformForced'));
        var placementUniformMergedRadio = placementPanel.add("radiobutton", undefined, L('placementUniformMerged'));
        var placementMamaIkiRadio = placementPanel.add("radiobutton", undefined, L('placementMamaIki'));
        placementMamaIkiRadio.value = true;


        var stylePanel = rulesPanel.add("panel", undefined, L('stylePanelTitle'));
        setupPanel(stylePanel);
        stylePanel.alignChildren = "fill";

        var strokeWidthGroup = stylePanel.add("group");
        strokeWidthGroup.orientation = "row";
        strokeWidthGroup.add("statictext", undefined, labelText('strokeWidth'));
        var strokeUnitLabel = getCurrentStrokeUnitLabel();
        var strokeWidthInput = strokeWidthGroup.add("edittext", undefined, getDefaultStrokeWidthText(strokeUnitLabel));
        strokeWidthInput.characters = 5;
        changeValueByArrowKey(strokeWidthInput);
        strokeWidthGroup.add("statictext", undefined, strokeUnitLabel);

        var rulesK100Checkbox = stylePanel.add("checkbox", undefined, L('rulesK100'));
        rulesK100Checkbox.value = true;

        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "center";
        btnGroup.margins = [0, 10, 0, 0];

        var btnCancel = btnGroup.add("button", undefined, L('cancel'), { name: "cancel" });
        var btnOk = btnGroup.add("button", undefined, L('ok'), { name: "ok" });

        btnOk.onClick = function () {
            var releaseMask = releaseMaskCheckbox.value;
            var moveText = moveTextCheckbox.value;
            var setK100 = setK100Checkbox.value;
            var removeDuplicate = removeDuplicateCheckbox.value;
            var removeSmall = removeSmallCheckbox.value;
            var adjustCellBg = adjustCellBgCheckbox.value;
            var equalizeHeights = equalizeHeightsCheckbox.value;
            var centerline = centerlineCheckbox.value;
            var placementMode = placementUniformForcedRadio.value ? "forced"
                : placementUniformMergedRadio.value ? "merged"
                    : "keepColumnWidths";
            var outerToRect = outerToRectCheckbox.value;
            var strokeUnitLabel = getCurrentStrokeUnitLabel();
            var strokeWidthText = String(strokeWidthInput.text).replace(/,/g, ".");
            var strokeWidthValue = parseFloat(strokeWidthText);
            if (isNaN(strokeWidthValue) || strokeWidthValue <= 0) {
                strokeWidthValue = parseFloat(getDefaultStrokeWidthText(strokeUnitLabel));
            }
            var strokeWidthPt = convertLengthToPoints(strokeWidthValue, strokeUnitLabel);
            var rulesK100 = rulesK100Checkbox.value;
            dlg.close();
            executeRelease(releaseMask, moveText, setK100, removeDuplicate, removeSmall, adjustCellBg, equalizeHeights, centerline, placementMode, outerToRect, strokeWidthPt, rulesK100);
        };

        dlg.show();
    }

    /* 実処理：選択オブジェクトに対して整形処理を実行 / Core processing pipeline */
    function executeRelease(releaseMask, moveText, setK100, removeDuplicate, removeSmall, adjustCellBg, equalizeHeights, centerline, placementMode, outerToRect, strokeWidthPt, rulesK100) {
        if (!app.documents.length || !app.activeDocument.selection.length) {
            alert(L('noSelection'));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = doc.selection;

        if (removeDuplicate) {
            processClipGroupTexts(selectionItems);
        }

        if (releaseMask) {
            for (var i = 0; i < selectionItems.length; i++) {
                var currentItem = selectionItems[i];
                if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
                    for (var j = 0; j < currentItem.pageItems.length; j++) {
                        var item = currentItem.pageItems[j];
                        if (item.typename === "PathItem" && item.clipping) {
                            item.remove();
                            break;
                        }
                    }
                    ungroup(currentItem);
                }
            }
        }

        if (removeSmall) {
            removeSmallPathItems(doc);
        }

        if (moveText) {
            moveTextsToLayer(doc, TEXT_LAYER_NAME);
        }

        if (setK100) {
            applyK100ToAllTexts(doc);
        }

        if (adjustCellBg) {
            autoSelectAndMerge({
                layerName: PATH_LAYER_NAME,
                excludeLayerNames: [TEXT_LAYER_NAME],
                silent: true
            });
        }
        // セル背景のグループを強制解除（スナップ精度向上のため）
        ungroupAllInLayer(doc, PATH_LAYER_NAME);

        var finalItems = null;
        var centerlineApplied = false;
        if (centerline) {
            var centerLines = convertRectanglesToCenterLines(doc, [TEXT_LAYER_NAME, PATH_LAYER_NAME]);
            if (centerLines && centerLines.length > 0) {
                centerLines = dedupeOverlappingLines(centerLines);
                if (placementMode) {
                    applyPlacementMode(centerLines, placementMode);
                    centerLines = dedupeOverlappingLines(centerLines);
                }
                finalItems = centerLines;
                centerlineApplied = true;
            }
        }
        if (!finalItems) {
            finalItems = collectRuleLines(doc, [TEXT_LAYER_NAME, PATH_LAYER_NAME]);
        }

        // 外枠矩形化より前に、純粋な中心線群に対してセル背景をスナップする
        // Snap cell backgrounds while finalItems is still a homogeneous set of centerlines
        if (adjustCellBg && equalizeHeights && finalItems && finalItems.length > 0) {
            snapCellBackgroundsAcrossLayers(doc, [TEXT_LAYER_NAME], finalItems, 0);
        }

        // 外枠矩形化より前に線幅・色を確定させる
        // 新しい外枠矩形は referencePath（中心線）から strokeWidth/color を継承するため
        // Apply stroke styling before outerToRect so the new rectangle inherits the styled stroke
        if (typeof strokeWidthPt === "number" && strokeWidthPt > 0) {
            applyUniformStrokeWidth(finalItems, strokeWidthPt);
        }
        if (rulesK100) {
            applyK100ToLines(finalItems);
        }

        if (centerlineApplied && outerToRect && finalItems && finalItems.length > 0) {
            finalItems = convertOuterFrameToRectangle(finalItems);
        }
    }

    /* 小さいパスを削除 / Remove small path items based on text size */
    function removeSmallPathItems(doc) {
        var minimumTextSize = null;
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            var textFrame = doc.textFrames[textFrameIndex];
            try {
                var textSize = textFrame.textRange.characterAttributes.size;
                if (typeof textSize === 'number' && textSize > 0) {
                    if (minimumTextSize === null || textSize < minimumTextSize) minimumTextSize = textSize;
                }
            } catch (e) { }
        }
        if (minimumTextSize === null || minimumTextSize <= 0) return;

        var removalThreshold = minimumTextSize / 2;
        var pathItemsToRemove = [];
        for (var pathItemIndex = 0; pathItemIndex < doc.pathItems.length; pathItemIndex++) {
            var pathItem = doc.pathItems[pathItemIndex];
            try {
                if (pathItem.clipping) continue;
                if (pathItem.locked || (pathItem.layer && pathItem.layer.locked)) continue;
                var pathBounds = pathItem.geometricBounds;
                var pathWidth = pathBounds[2] - pathBounds[0];
                var pathHeight = pathBounds[1] - pathBounds[3];
                if (pathWidth < removalThreshold && pathHeight < removalThreshold) {
                    pathItemsToRemove.push(pathItem);
                }
            } catch (e) { }
        }
        for (var removeIndex = 0; removeIndex < pathItemsToRemove.length; removeIndex++) {
            (function (index) {
                safeCall(function () { pathItemsToRemove[index].remove(); });
            })(removeIndex);
        }
    }

    /* クリップグループ内テキスト処理 / Process text inside clipped groups */
    function processClipGroupTexts(selectionItems) {
        for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
            var selectedItem = selectionItems[selectionIndex];
            if (selectedItem.typename !== "GroupItem" || selectedItem.clipped !== true) continue;
            deduplicateTextsInGroup(selectedItem);
            concatenateTextsInGroup(selectedItem);
        }
    }

    /* 重複テキスト削除 / Remove duplicate texts in group */
    function deduplicateTextsInGroup(group) {
        var textFrames = [];
        collectTextFrames(group, textFrames);
        if (textFrames.length < 2) return;
        var seenTextContents = {};
        var textFramesToRemove = [];
        for (var textIndex = textFrames.length - 1; textIndex >= 0; textIndex--) {
            var textKey = "k:" + textFrames[textIndex].contents;
            if (seenTextContents[textKey]) {
                textFramesToRemove.push(textFrames[textIndex]);
            } else {
                seenTextContents[textKey] = true;
            }
        }
        for (var removeIndex = 0; removeIndex < textFramesToRemove.length; removeIndex++) {
            (function (index) {
                safeCall(function () { textFramesToRemove[index].remove(); });
            })(removeIndex);
        }
    }

    /* テキスト連結 / Concatenate texts in reading order */
    function concatenateTextsInGroup(group) {
        var textFrames = [];
        collectTextFrames(group, textFrames);
        if (textFrames.length < 2) return;
        textFrames.sort(function (textFrameA, textFrameB) {
            var yA, yB, xA, xB;
            try {
                yA = textFrameA.position[1]; yB = textFrameB.position[1];
                xA = textFrameA.position[0]; xB = textFrameB.position[0];
            } catch (e) { return 0; }
            if (Math.abs(yA - yB) > 0.5) return yB - yA;
            return xA - xB;
        });
        var combinedText = "";
        for (var textIndex = 0; textIndex < textFrames.length; textIndex++) {
            combinedText += textFrames[textIndex].contents;
        }
        try {
            textFrames[0].contents = combinedText;
        } catch (e) { return; }
        for (var removeIndex = 1; removeIndex < textFrames.length; removeIndex++) {
            (function (index) {
                safeCall(function () { textFrames[index].remove(); });
            })(removeIndex);
        }
    }

    /* テキスト収集 / Collect text frames recursively */
    function collectTextFrames(group, out) {
        for (var pageItemIndex = 0; pageItemIndex < group.pageItems.length; pageItemIndex++) {
            var pageItem = group.pageItems[pageItemIndex];
            try {
                if (pageItem.typename === "TextFrame") {
                    out.push(pageItem);
                } else if (pageItem.typename === "GroupItem") {
                    collectTextFrames(pageItem, out);
                }
            } catch (e) { }
        }
    }

    /* テキストをK100に変換 / Apply K100 to all texts */
    function applyK100ToAllTexts(doc) {
        var blackColor = getK100Black();
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            var textFrame = doc.textFrames[textFrameIndex];
            try {
                textFrame.textRange.characterAttributes.fillColor = blackColor;
            } catch (e) { }
        }
    }

    /* テキストを指定レイヤーへ移動 / Move texts to layer */
    function moveTextsToLayer(doc, layerName) {
        var targetLayer = null;
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            if (doc.layers[layerIndex].name === layerName) {
                targetLayer = doc.layers[layerIndex];
                break;
            }
        }
        if (!targetLayer) {
            targetLayer = doc.layers.add();
            targetLayer.name = layerName;
        }
        if (targetLayer.locked) targetLayer.locked = false;
        if (!targetLayer.visible) targetLayer.visible = true;

        var textFrames = [];
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            textFrames.push(doc.textFrames[textFrameIndex]);
        }
        for (var moveIndex = 0; moveIndex < textFrames.length; moveIndex++) {
            var textFrame = textFrames[moveIndex];
            if (textFrame.parent === targetLayer) continue;
            (function (targetTextFrame) {
                safeCall(function () { targetTextFrame.move(targetLayer, ElementPlacement.PLACEATBEGINNING); });
            })(textFrame);
        }
    }

    /* グループ解除 / Ungroup items */
    function ungroup(groupItem) {
        var parent = groupItem.parent;
        while (groupItem.pageItems.length > 0) {
            groupItem.pageItems[0].move(parent, ElementPlacement.PLACEATEND);
        }
        groupItem.remove();
    }

    /* K100ブラック生成 / Create K100 black color */
    function getK100Black() {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    /* 長方形を中心線に変換 / Convert rectangles to centerlines */
    function convertRectanglesToCenterLines(doc, excludeLayerNames) {
        var rectangleItems = [];
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var layer = doc.layers[layerIndex];
            var shouldSkipLayer = false;
            if (excludeLayerNames) {
                for (var excludeIndex = 0; excludeIndex < excludeLayerNames.length; excludeIndex++) {
                    if (layer.name === excludeLayerNames[excludeIndex]) { shouldSkipLayer = true; break; }
                }
            }
            if (shouldSkipLayer || layer.locked || !layer.visible) continue;
            collectRectsForCenterline(layer.pageItems, rectangleItems);
        }
        var centerLineItems = [];
        for (var rectangleIndex = 0; rectangleIndex < rectangleItems.length; rectangleIndex++) {
            try {
                var centerLine = convertRectToCenterLine(rectangleItems[rectangleIndex], true, 0);
                if (centerLine) centerLineItems.push(centerLine);
            } catch (e) { }
        }
        return centerLineItems;
    }

    /*
     目的:
     Excel由来の罫線は位置や本数が不揃いで、セルグリッドとして扱いにくい。
     この関数は線を水平／垂直に分類し、外接範囲から行列グリッドを再構築するための入口。
     配置モードに応じて、均等配置・結合セル対応・列幅保持のいずれかを適用する。

     Purpose:
     Lines imported from Excel are often inconsistent in position and count.
     This function classifies lines into horizontal/vertical and rebuilds a grid
     from their bounds, then delegates to a specific placement strategy.
    */
    /* 配置モード適用 / Apply placement mode to lines */
    function applyPlacementMode(lines, mode) {
        if (!lines || lines.length === 0) return;

        var horizontalLines = [];
        var verticalLines = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var linePath = lines[lineIndex];
            try {
                if (!linePath.pathPoints || linePath.pathPoints.length !== 2) continue;
                var startAnchor = linePath.pathPoints[0].anchor;
                var endAnchor = linePath.pathPoints[1].anchor;
                var horizontalDistance = Math.abs(startAnchor[0] - endAnchor[0]);
                var verticalDistance = Math.abs(startAnchor[1] - endAnchor[1]);
                if (horizontalDistance >= verticalDistance) {
                    horizontalLines.push({
                        path: linePath,
                        y: (startAnchor[1] + endAnchor[1]) / 2,
                        minX: Math.min(startAnchor[0], endAnchor[0]),
                        maxX: Math.max(startAnchor[0], endAnchor[0])
                    });
                } else {
                    verticalLines.push({
                        path: linePath,
                        x: (startAnchor[0] + endAnchor[0]) / 2,
                        minY: Math.min(startAnchor[1], endAnchor[1]),
                        maxY: Math.max(startAnchor[1], endAnchor[1])
                    });
                }
            } catch (e) { }
        }
        if (horizontalLines.length === 0 || verticalLines.length === 0) return;

        var minX = verticalLines[0].x;
        var maxX = verticalLines[0].x;
        for (var verticalIndex = 1; verticalIndex < verticalLines.length; verticalIndex++) {
            if (verticalLines[verticalIndex].x < minX) minX = verticalLines[verticalIndex].x;
            if (verticalLines[verticalIndex].x > maxX) maxX = verticalLines[verticalIndex].x;
        }
        var minY = horizontalLines[0].y;
        var maxY = horizontalLines[0].y;
        for (var horizontalIndex = 1; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            if (horizontalLines[horizontalIndex].y < minY) minY = horizontalLines[horizontalIndex].y;
            if (horizontalLines[horizontalIndex].y > maxY) maxY = horizontalLines[horizontalIndex].y;
        }
        var lineBounds = { minX: minX, maxX: maxX, minY: minY, maxY: maxY };

        if (mode === "forced") {
            alignLinesEvenly(horizontalLines, verticalLines, lineBounds);
        } else if (mode === "merged") {
            alignLinesWithMergedCells(horizontalLines, verticalLines, lineBounds);
        } else if (mode === "keepColumnWidths") {
            alignLinesKeepingColumnWidths(horizontalLines, verticalLines, lineBounds);
        }
    }

    /* 列幅保持で整列 / Align lines keeping column widths */
    function alignLinesKeepingColumnWidths(horizontalLines, verticalLines, bounds) {
        var horizontalYPositions = [];
        if (horizontalLines.length >= 2) {
            horizontalLines.sort(function (lineA, lineB) { return lineA.y - lineB.y; });
            var horizontalStep = (bounds.maxY - bounds.minY) / (horizontalLines.length - 1);
            for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
                horizontalYPositions.push(bounds.minY + horizontalStep * horizontalIndex);
            }
        } else {
            for (var fallbackHorizontalIndex = 0; fallbackHorizontalIndex < horizontalLines.length; fallbackHorizontalIndex++) {
                horizontalYPositions.push(horizontalLines[fallbackHorizontalIndex].y);
            }
        }

        var topY = Math.max(bounds.minY, bounds.maxY);
        var bottomY = Math.min(bounds.minY, bounds.maxY);

        for (var horizontalLineIndex = 0; horizontalLineIndex < horizontalLines.length; horizontalLineIndex++) {
            setLineEndpointsSafe(horizontalLines[horizontalLineIndex].path, [bounds.minX, horizontalYPositions[horizontalLineIndex]], [bounds.maxX, horizontalYPositions[horizontalLineIndex]]);
        }
        for (var verticalLineIndex = 0; verticalLineIndex < verticalLines.length; verticalLineIndex++) {
            setLineEndpointsSafe(verticalLines[verticalLineIndex].path, [verticalLines[verticalLineIndex].x, topY], [verticalLines[verticalLineIndex].x, bottomY]);
        }
    }

    /* 線端座標を安全に設定 / Safely set line endpoints */
    function setLineEndpointsSafe(path, startPoint, endPoint) {
        safeCall(function () {
            path.pathPoints[0].anchor = startPoint;
            path.pathPoints[0].leftDirection = startPoint;
            path.pathPoints[0].rightDirection = startPoint;
            path.pathPoints[1].anchor = endPoint;
            path.pathPoints[1].leftDirection = endPoint;
            path.pathPoints[1].rightDirection = endPoint;
        });
    }

    /* 均等配置 / Evenly distribute lines */
    function alignLinesEvenly(horizontalLines, verticalLines, bounds) {
        var horizontalYPositions = [];
        if (horizontalLines.length >= 2) {
            horizontalLines.sort(function (lineA, lineB) { return lineA.y - lineB.y; });
            var horizontalStep = (bounds.maxY - bounds.minY) / (horizontalLines.length - 1);
            for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
                horizontalYPositions.push(bounds.minY + horizontalStep * horizontalIndex);
            }
        } else {
            for (var fallbackHorizontalIndex = 0; fallbackHorizontalIndex < horizontalLines.length; fallbackHorizontalIndex++) {
                horizontalYPositions.push(horizontalLines[fallbackHorizontalIndex].y);
            }
        }

        var verticalXPositions = [];
        if (verticalLines.length >= 2) {
            verticalLines.sort(function (lineA, lineB) { return lineA.x - lineB.x; });
            var verticalStep = (bounds.maxX - bounds.minX) / (verticalLines.length - 1);
            for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
                verticalXPositions.push(bounds.minX + verticalStep * verticalIndex);
            }
        } else {
            for (var fallbackVerticalIndex = 0; fallbackVerticalIndex < verticalLines.length; fallbackVerticalIndex++) {
                verticalXPositions.push(verticalLines[fallbackVerticalIndex].x);
            }
        }

        var topY = Math.max(bounds.minY, bounds.maxY);
        var bottomY = Math.min(bounds.minY, bounds.maxY);

        for (var horizontalLineIndex = 0; horizontalLineIndex < horizontalLines.length; horizontalLineIndex++) {
            setLineEndpointsSafe(horizontalLines[horizontalLineIndex].path, [bounds.minX, horizontalYPositions[horizontalLineIndex]], [bounds.maxX, horizontalYPositions[horizontalLineIndex]]);
        }
        for (var verticalLineIndex = 0; verticalLineIndex < verticalLines.length; verticalLineIndex++) {
            setLineEndpointsSafe(verticalLines[verticalLineIndex].path, [verticalXPositions[verticalLineIndex], topY], [verticalXPositions[verticalLineIndex], bottomY]);
        }
    }

    /*
     目的:
     結合セルがある表では、単純な等間隔配置では列境界と一致しない。
     近接座標で線をグルーピングし、実際の列／行境界にスナップさせることで、
     結合セルを壊さずにグリッドを再構築する。

     Purpose:
     With merged cells, uniform spacing breaks true column boundaries.
     This groups lines by nearby coordinates and snaps them to detected
     row/column edges, preserving merged-cell structure.
    */
    /* 結合セル対応配置 / Align lines with merged-cell awareness */
    function alignLinesWithMergedCells(horizontalLines, verticalLines, bounds) {
        var coordinateTolerance = 5.0;
        var horizontalGroups = groupLinesByCoordinate(horizontalLines, "y", coordinateTolerance);
        var verticalGroups = groupLinesByCoordinate(verticalLines, "x", coordinateTolerance);

        var horizontalYPositions = [];
        if (horizontalGroups.length >= 2) {
            var horizontalStep = (bounds.maxY - bounds.minY) / (horizontalGroups.length - 1);
            for (var horizontalGroupIndex = 0; horizontalGroupIndex < horizontalGroups.length; horizontalGroupIndex++) {
                horizontalYPositions.push(bounds.minY + horizontalStep * horizontalGroupIndex);
            }
        } else {
            for (var fallbackHorizontalGroupIndex = 0; fallbackHorizontalGroupIndex < horizontalGroups.length; fallbackHorizontalGroupIndex++) {
                horizontalYPositions.push(horizontalGroups[fallbackHorizontalGroupIndex].coord);
            }
        }

        var verticalXPositions = [];
        if (verticalGroups.length >= 2) {
            var verticalStep = (bounds.maxX - bounds.minX) / (verticalGroups.length - 1);
            for (var verticalGroupIndex = 0; verticalGroupIndex < verticalGroups.length; verticalGroupIndex++) {
                verticalXPositions.push(bounds.minX + verticalStep * verticalGroupIndex);
            }
        } else {
            for (var fallbackVerticalGroupIndex = 0; fallbackVerticalGroupIndex < verticalGroups.length; fallbackVerticalGroupIndex++) {
                verticalXPositions.push(verticalGroups[fallbackVerticalGroupIndex].coord);
            }
        }

        for (var rowGroupIndex = 0; rowGroupIndex < horizontalGroups.length; rowGroupIndex++) {
            var rowY = horizontalYPositions[rowGroupIndex];
            var rowLines = horizontalGroups[rowGroupIndex].lines;
            for (var rowLineIndex = 0; rowLineIndex < rowLines.length; rowLineIndex++) {
                var horizontalLine = rowLines[rowLineIndex];
                var leftColumnIndex = findClosestIndexByProperty(horizontalLine.minX, verticalGroups, "coord");
                var rightColumnIndex = findClosestIndexByProperty(horizontalLine.maxX, verticalGroups, "coord");
                setLineEndpointsSafe(horizontalLine.path, [verticalXPositions[leftColumnIndex], rowY], [verticalXPositions[rightColumnIndex], rowY]);
            }
        }
        for (var columnGroupIndex = 0; columnGroupIndex < verticalGroups.length; columnGroupIndex++) {
            var columnX = verticalXPositions[columnGroupIndex];
            var columnLines = verticalGroups[columnGroupIndex].lines;
            for (var columnLineIndex = 0; columnLineIndex < columnLines.length; columnLineIndex++) {
                var verticalLine = columnLines[columnLineIndex];
                var topRowIndex = findClosestIndexByProperty(verticalLine.maxY, horizontalGroups, "coord");
                var bottomRowIndex = findClosestIndexByProperty(verticalLine.minY, horizontalGroups, "coord");
                setLineEndpointsSafe(verticalLine.path, [columnX, horizontalYPositions[topRowIndex]], [columnX, horizontalYPositions[bottomRowIndex]]);
            }
        }
    }

    /* 座標近傍でグルーピング / Group lines by coordinate tolerance */
    function groupLinesByCoordinate(lines, propName, tolerance) {
        var sorted = lines.slice();
        sorted.sort(function (a, b) { return a[propName] - b[propName]; });
        var groups = [];
        for (var i = 0; i < sorted.length; i++) {
            var line = sorted[i];
            var coord = line[propName];
            if (groups.length === 0 || Math.abs(coord - groups[groups.length - 1].coord) > tolerance) {
                groups.push({ coord: coord, lines: [line] });
            } else {
                var last = groups[groups.length - 1];
                last.lines.push(line);
                var sum = 0;
                for (var k = 0; k < last.lines.length; k++) sum += last.lines[k][propName];
                last.coord = sum / last.lines.length;
            }
        }
        return groups;
    }

    /* 最も近いインデックス取得 / Find closest index by property */
    function findClosestIndexByProperty(target, items, propName) {
        var bestIdx = 0;
        var bestDist = Math.abs(target - items[0][propName]);
        for (var i = 1; i < items.length; i++) {
            var d = Math.abs(target - items[i][propName]);
            if (d < bestDist) { bestIdx = i; bestDist = d; }
        }
        return bestIdx;
    }

    /* 重複線の統合 / Merge overlapping lines */
    function dedupeOverlappingLines(lines) {
        if (!lines || lines.length < 2) return lines || [];
        var TOL = 0.5;
        var groups = [];

        for (var i = 0; i < lines.length; i++) {
            var info = lineKeyInfo(lines[i]);
            if (!info) {
                groups.push({ orient: null, coord: 0, items: [{ path: lines[i], minSpan: 0, maxSpan: 0 }] });
                continue;
            }
            var matched = null;
            for (var g = 0; g < groups.length; g++) {
                var gr = groups[g];
                if (gr.orient !== info.orient) continue;
                if (Math.abs(gr.coord - info.coord) > TOL) continue;
                matched = gr;
                break;
            }
            if (matched) {
                matched.items.push({ path: lines[i], minSpan: info.minSpan, maxSpan: info.maxSpan });
            } else {
                groups.push({
                    orient: info.orient,
                    coord: info.coord,
                    items: [{ path: lines[i], minSpan: info.minSpan, maxSpan: info.maxSpan }]
                });
            }
        }

        var kept = [];
        for (var gi = 0; gi < groups.length; gi++) {
            var grp = groups[gi];
            if (grp.items.length === 1 || grp.orient === null) {
                kept.push(grp.items[0].path);
                continue;
            }
            var unionMin = grp.items[0].minSpan;
            var unionMax = grp.items[0].maxSpan;
            for (var ii = 1; ii < grp.items.length; ii++) {
                if (grp.items[ii].minSpan < unionMin) unionMin = grp.items[ii].minSpan;
                if (grp.items[ii].maxSpan > unionMax) unionMax = grp.items[ii].maxSpan;
            }
            var keeper = grp.items[0].path;
            if (grp.orient === "h") {
                setLineEndpointsSafe(keeper, [unionMin, grp.coord], [unionMax, grp.coord]);
            } else {
                setLineEndpointsSafe(keeper, [grp.coord, unionMax], [grp.coord, unionMin]);
            }
            for (var removeIndex = 1; removeIndex < grp.items.length; removeIndex++) {
                (function (index) {
                    safeCall(function () { grp.items[index].path.remove(); });
                })(removeIndex);
            }
            kept.push(keeper);
        }
        return kept;
    }

    /* 線の特徴抽出 / Extract line orientation and span info */
    function lineKeyInfo(path) {
        try {
            if (!path.pathPoints || path.pathPoints.length !== 2) return null;
            var startAnchor = path.pathPoints[0].anchor;
            var endAnchor = path.pathPoints[1].anchor;
            var horizontalDistance = Math.abs(startAnchor[0] - endAnchor[0]);
            var verticalDistance = Math.abs(startAnchor[1] - endAnchor[1]);
            if (horizontalDistance >= verticalDistance) {
                return {
                    orient: "h",
                    coord: (startAnchor[1] + endAnchor[1]) / 2,
                    minSpan: Math.min(startAnchor[0], endAnchor[0]),
                    maxSpan: Math.max(startAnchor[0], endAnchor[0])
                };
            }
            return {
                orient: "v",
                coord: (startAnchor[0] + endAnchor[0]) / 2,
                minSpan: Math.min(startAnchor[1], endAnchor[1]),
                maxSpan: Math.max(startAnchor[1], endAnchor[1])
            };
        } catch (e) { return null; }
    }

    /* 外枠を長方形化 / Convert outer frame to rectangle */
    function convertOuterFrameToRectangle(lines) {
        if (!lines || lines.length < 4) return lines || [];

        var horizontalLines = [];
        var verticalLines = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var linePath = lines[lineIndex];
            try {
                if (!linePath.pathPoints || linePath.pathPoints.length !== 2) continue;
                var startAnchor = linePath.pathPoints[0].anchor;
                var endAnchor = linePath.pathPoints[1].anchor;
                var horizontalDistance = Math.abs(startAnchor[0] - endAnchor[0]);
                var verticalDistance = Math.abs(startAnchor[1] - endAnchor[1]);
                if (horizontalDistance >= verticalDistance) {
                    horizontalLines.push({ path: linePath, y: (startAnchor[1] + endAnchor[1]) / 2 });
                } else {
                    verticalLines.push({ path: linePath, x: (startAnchor[0] + endAnchor[0]) / 2 });
                }
            } catch (e) { }
        }
        if (horizontalLines.length < 2 || verticalLines.length < 2) return lines;

        var topLine = horizontalLines[0], bottomLine = horizontalLines[0];
        for (var horizontalIndex = 1; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            if (horizontalLines[horizontalIndex].y > topLine.y) topLine = horizontalLines[horizontalIndex];
            if (horizontalLines[horizontalIndex].y < bottomLine.y) bottomLine = horizontalLines[horizontalIndex];
        }
        var leftLine = verticalLines[0], rightLine = verticalLines[0];
        for (var verticalIndex = 1; verticalIndex < verticalLines.length; verticalIndex++) {
            if (verticalLines[verticalIndex].x < leftLine.x) leftLine = verticalLines[verticalIndex];
            if (verticalLines[verticalIndex].x > rightLine.x) rightLine = verticalLines[verticalIndex];
        }

        var outerCandidates = [topLine.path, bottomLine.path, leftLine.path, rightLine.path];
        var uniqueOuterLines = [];
        for (var candidateIndex = 0; candidateIndex < outerCandidates.length; candidateIndex++) {
            var isDuplicate = false;
            for (var uniqueIndex = 0; uniqueIndex < uniqueOuterLines.length; uniqueIndex++) {
                if (uniqueOuterLines[uniqueIndex] === outerCandidates[candidateIndex]) { isDuplicate = true; break; }
            }
            if (!isDuplicate) uniqueOuterLines.push(outerCandidates[candidateIndex]);
        }

        var minX = leftLine.x, maxX = rightLine.x;
        var minY = bottomLine.y, maxY = topLine.y;

        var referencePath = topLine.path;
        var parent = referencePath.parent;
        var rectanglePath;
        try {
            rectanglePath = parent.pathItems.rectangle(maxY, minX, maxX - minX, maxY - minY);
        } catch (e) {
            return lines;
        }
        rectanglePath.filled = false;
        rectanglePath.stroked = referencePath.stroked;
        if (referencePath.stroked) {
            safeCall(function () { rectanglePath.strokeColor = referencePath.strokeColor; });
            safeCall(function () { rectanglePath.strokeWidth = referencePath.strokeWidth; });
            safeCall(function () { rectanglePath.strokeCap = referencePath.strokeCap; });
        }
        safeCall(function () { rectanglePath.move(referencePath, ElementPlacement.PLACEAFTER); });

        for (var removeIndex = 0; removeIndex < uniqueOuterLines.length; removeIndex++) {
            (function (index) {
                safeCall(function () { uniqueOuterLines[index].remove(); });
            })(removeIndex);
        }

        var result = [];
        for (var resultIndex = 0; resultIndex < lines.length; resultIndex++) {
            var shouldKeepLine = true;
            for (var outerIndex = 0; outerIndex < uniqueOuterLines.length; outerIndex++) {
                if (lines[resultIndex] === uniqueOuterLines[outerIndex]) { shouldKeepLine = false; break; }
            }
            if (shouldKeepLine) result.push(lines[resultIndex]);
        }
        result.push(rectanglePath);
        return result;
    }

    /* 線幅統一 / Apply uniform stroke width */
    function applyUniformStrokeWidth(items, widthPt) {
        if (!items || items.length === 0 || !widthPt || widthPt <= 0) return;

        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            (function (index) {
                safeCall(function () {
                    items[index].stroked = true;
                    items[index].strokeWidth = widthPt;
                });
            })(itemIndex);
        }
    }

    /* 線をK100に変換 / Apply K100 to lines */
    function applyK100ToLines(items) {
        if (!items || items.length === 0) return;
        var blackColor = getK100Black();
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            (function (index) {
                safeCall(function () {
                    items[index].stroked = true;
                    items[index].strokeColor = blackColor;
                });
            })(itemIndex);
        }
    }

    /* レイヤー横断でセル背景（塗りクローズパス）を罫線中心線にスナップ
       Snap filled closed paths across all visible layers (FillSnapper-style) */
    function snapCellBackgroundsAcrossLayers(doc, excludeLayerNames, ruleItems, insetPt) {
        var ruleCenters = collectRuleCenterPositionsFromItems(ruleItems);
        if (!ruleCenters || ruleCenters.xPositions.length < 2 || ruleCenters.yPositions.length < 2) return 0;

        var rectangleItems = [];
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var layer = doc.layers[layerIndex];
            var shouldSkipLayer = false;
            if (excludeLayerNames) {
                for (var excludeIndex = 0; excludeIndex < excludeLayerNames.length; excludeIndex++) {
                    if (layer.name === excludeLayerNames[excludeIndex]) { shouldSkipLayer = true; break; }
                }
            }
            if (shouldSkipLayer || layer.locked || !layer.visible) continue;
            collectCellBackgroundSnapTargetsInItems(layer.pageItems, rectangleItems);
        }

        var inset = (typeof insetPt === "number" && insetPt > 0) ? insetPt : 0;
        var snapped = 0;
        for (var rectangleIndex = 0; rectangleIndex < rectangleItems.length; rectangleIndex++) {
            try {
                var rectangleItem = rectangleItems[rectangleIndex];
                var bounds = rectangleItem.geometricBounds;
                var snappedLeft = findNearestValue(bounds[0], ruleCenters.xPositions);
                var snappedTop = findNearestValue(bounds[1], ruleCenters.yPositions);
                var snappedRight = findNearestValue(bounds[2], ruleCenters.xPositions);
                var snappedBottom = findNearestValue(bounds[3], ruleCenters.yPositions);
                if (snappedLeft === null || snappedRight === null || snappedTop === null || snappedBottom === null) continue;
                if (snappedLeft >= snappedRight || snappedTop <= snappedBottom) continue;
                if (fitClosedPathToBounds(
                    rectangleItem,
                    snappedLeft + inset,
                    snappedTop - inset,
                    snappedRight - inset,
                    snappedBottom + inset
                )) {
                    snapped++;
                }
            } catch (e) { }
        }
        return snapped;
    }

    /* 任意のアンカー数のクローズパスを目標境界に比例マッピング / Proportionally map a closed path's anchors to target bounds */
    function fitClosedPathToBounds(pathItem, left, top, right, bottom) {
        if (!pathItem || !pathItem.pathPoints || pathItem.pathPoints.length < 2) return false;

        var sourceBounds = pathItem.geometricBounds;
        var sourceLeft = sourceBounds[0];
        var sourceTop = sourceBounds[1];
        var sourceWidth = sourceBounds[2] - sourceLeft;
        var sourceHeight = sourceTop - sourceBounds[3];
        var targetWidth = right - left;
        var targetHeight = top - bottom;

        if (sourceWidth <= 0.0001 || sourceHeight <= 0.0001) return false;
        if (targetWidth <= 0.0001 || targetHeight <= 0.0001) return false;

        var scaleX = targetWidth / sourceWidth;
        var scaleY = targetHeight / sourceHeight;

        function mapPoint(p) {
            return [
                left + (p[0] - sourceLeft) * scaleX,
                top - (sourceTop - p[1]) * scaleY
            ];
        }

        try {
            var pathPointList = pathItem.pathPoints;
            for (var i = 0; i < pathPointList.length; i++) {
                var pt = pathPointList[i];
                pt.anchor = mapPoint(pt.anchor);
                pt.leftDirection = mapPoint(pt.leftDirection);
                pt.rightDirection = mapPoint(pt.rightDirection);
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    /* 指定レイヤー内のグループをすべて解除 / Ungroup all groups in a layer */
    function ungroupAllInLayer(doc, layerName) {
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var layer = doc.layers[layerIndex];
            if (layer.name !== layerName) continue;
            if (layer.locked || !layer.visible) return;
            ungroupItemsRecursive(layer.pageItems);
            break;
        }
    }

    /* グループを再帰的に解除 / Recursively ungroup items */
    function ungroupItemsRecursive(items) {
        for (var i = items.length - 1; i >= 0; i--) {
            var item = items[i];
            try {
                if (item.typename === "GroupItem") {
                    var parent = item.parent;
                    while (item.pageItems.length > 0) {
                        item.pageItems[0].move(parent, ElementPlacement.PLACEATEND);
                    }
                    item.remove();
                }
            } catch (e) { }
        }
    }

    /* セル背景スナップ対象を再帰収集 / Recursively collect cell background snap targets */
    function collectCellBackgroundSnapTargetsInItems(items, out) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var pageItem = items[itemIndex];
            try {
                if (!pageItem || pageItem.locked || pageItem.hidden) continue;
                if (pageItem.typename === "PathItem") {
                    if (isCellBackgroundSnapTarget(pageItem)) out.push(pageItem);
                } else if (pageItem.typename === "GroupItem") {
                    collectCellBackgroundSnapTargetsInItems(pageItem.pageItems, out);
                } else if (pageItem.typename === "CompoundPathItem") {
                    for (var pathIndex = 0; pathIndex < pageItem.pathItems.length; pathIndex++) {
                        var childPathItem = pageItem.pathItems[pathIndex];
                        if (isCellBackgroundSnapTarget(childPathItem)) out.push(childPathItem);
                    }
                }
            } catch (e) { }
        }
    }

    /* セル背景スナップ対象判定 / Check if path can be snapped as a cell background */
    function isCellBackgroundSnapTarget(pathItem) {
        try {
            if (!pathItem || pathItem.typename !== "PathItem") return false;
            if (!pathItem.closed || pathItem.clipping) return false;
            if (!pathItem.pathPoints || pathItem.pathPoints.length < 2) return false;
            if (!pathItem.filled) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    /* 罫線アイテムから中心線座標を収集 / Collect centerline positions from rule items */
    function collectRuleCenterPositionsFromItems(ruleItems) {
        var xPositions = [];
        var yPositions = [];
        if (!ruleItems || ruleItems.length === 0) return { xPositions: xPositions, yPositions: yPositions };

        for (var itemIndex = 0; itemIndex < ruleItems.length; itemIndex++) {
            collectRuleCenterPositionsFromItem(ruleItems[itemIndex], xPositions, yPositions);
        }

        return {
            xPositions: uniqueSortedValues(xPositions, 0.5),
            yPositions: uniqueSortedValues(yPositions, 0.5)
        };
    }

    /* 単一アイテムから罫線中心座標を収集 / Collect rule center positions from one item */
    function collectRuleCenterPositionsFromItem(pageItem, xPositions, yPositions) {
        try {
            if (!pageItem || pageItem.locked || pageItem.hidden) return;
            if (pageItem.typename === "PathItem") {
                if (!pageItem.closed && pageItem.pathPoints && pageItem.pathPoints.length === 2) {
                    var startAnchor = pageItem.pathPoints[0].anchor;
                    var endAnchor = pageItem.pathPoints[1].anchor;
                    var horizontalDistance = Math.abs(startAnchor[0] - endAnchor[0]);
                    var verticalDistance = Math.abs(startAnchor[1] - endAnchor[1]);
                    if (horizontalDistance >= verticalDistance) {
                        yPositions.push((startAnchor[1] + endAnchor[1]) / 2);
                    } else {
                        xPositions.push((startAnchor[0] + endAnchor[0]) / 2);
                    }
                } else if (pageItem.closed && pageItem.pathPoints && pageItem.pathPoints.length === 4) {
                    var bounds = pageItem.geometricBounds;
                    xPositions.push(bounds[0]);
                    xPositions.push(bounds[2]);
                    yPositions.push(bounds[1]);
                    yPositions.push(bounds[3]);
                }
            } else if (pageItem.typename === "GroupItem") {
                for (var childIndex = 0; childIndex < pageItem.pageItems.length; childIndex++) {
                    collectRuleCenterPositionsFromItem(pageItem.pageItems[childIndex], xPositions, yPositions);
                }
            } else if (pageItem.typename === "CompoundPathItem") {
                for (var pathIndex = 0; pathIndex < pageItem.pathItems.length; pathIndex++) {
                    collectRuleCenterPositionsFromItem(pageItem.pathItems[pathIndex], xPositions, yPositions);
                }
            }
        } catch (e) { }
    }

    /* 近接値を統合して昇順化 / Merge nearby values and sort ascending */
    function uniqueSortedValues(values, tolerance) {
        var sortedValues = values.slice();
        sortedValues.sort(function (valueA, valueB) { return valueA - valueB; });
        var result = [];
        for (var valueIndex = 0; valueIndex < sortedValues.length; valueIndex++) {
            var value = sortedValues[valueIndex];
            if (result.length === 0 || Math.abs(value - result[result.length - 1]) > tolerance) {
                result.push(value);
            } else {
                result[result.length - 1] = (result[result.length - 1] + value) / 2;
            }
        }
        return result;
    }

    /* 最も近い値を取得 / Find nearest value */
    function findNearestValue(targetValue, values) {
        if (!values || values.length === 0) return null;
        var nearestValue = values[0];
        var nearestDistance = Math.abs(targetValue - nearestValue);
        for (var valueIndex = 1; valueIndex < values.length; valueIndex++) {
            var distance = Math.abs(targetValue - values[valueIndex]);
            if (distance < nearestDistance) {
                nearestValue = values[valueIndex];
                nearestDistance = distance;
            }
        }
        return nearestValue;
    }

    /* 中心線用矩形収集 / Collect rectangles for centerline conversion */
    function collectRectsForCenterline(items, out) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var pageItem = items[itemIndex];
            try {
                if (pageItem.locked || pageItem.hidden) continue;
                if (pageItem.typename === "PathItem") {
                    if (pageItem.closed && pageItem.pathPoints.length === 4) out.push(pageItem);
                } else if (pageItem.typename === "GroupItem") {
                    collectRectsForCenterline(pageItem.pageItems, out);
                } else if (pageItem.typename === "CompoundPathItem") {
                    if (pageItem.pathItems.length === 1) {
                        var childPathItem = pageItem.pathItems[0];
                        if (childPathItem.closed && childPathItem.pathPoints.length === 4) out.push(childPathItem);
                    }
                }
            } catch (e) { }
        }
    }

    /* 罫線（2点パス）収集 / Collect rule lines (open 2-point paths) */
    function collectRuleLines(doc, excludeLayerNames) {
        var lines = [];
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var layer = doc.layers[layerIndex];
            var shouldSkipLayer = false;
            if (excludeLayerNames) {
                for (var excludeIndex = 0; excludeIndex < excludeLayerNames.length; excludeIndex++) {
                    if (layer.name === excludeLayerNames[excludeIndex]) { shouldSkipLayer = true; break; }
                }
            }
            if (shouldSkipLayer || layer.locked || !layer.visible) continue;
            collectRuleLinesInItems(layer.pageItems, lines);
        }
        return lines;
    }

    function collectRuleLinesInItems(items, out) {
        for (var i = 0; i < items.length; i++) {
            var pageItem = items[i];
            try {
                if (pageItem.locked || pageItem.hidden) continue;
                if (pageItem.typename === "PathItem") {
                    if (!pageItem.closed && pageItem.pathPoints && pageItem.pathPoints.length === 2) {
                        out.push(pageItem);
                    }
                } else if (pageItem.typename === "GroupItem") {
                    collectRuleLinesInItems(pageItem.pageItems, out);
                } else if (pageItem.typename === "CompoundPathItem") {
                    for (var c = 0; c < pageItem.pathItems.length; c++) {
                        var child = pageItem.pathItems[c];
                        if (!child.closed && child.pathPoints && child.pathPoints.length === 2) {
                            out.push(child);
                        }
                    }
                }
            } catch (e) { }
        }
    }

    /* 長方形→中心線変換 / Convert rectangle to centerline */
    function convertRectToCenterLine(rect, correctRotation, minShortSidePt) {
        if (correctRotation && rect.pathPoints.length === 4) {
            var firstAnchor = rect.pathPoints[0].anchor;
            var secondAnchor = rect.pathPoints[1].anchor;
            var deltaX = secondAnchor[0] - firstAnchor[0];
            var deltaY = secondAnchor[1] - firstAnchor[1];
            var angleDeg = Math.atan2(deltaY, deltaX) * 180 / Math.PI;
            if (angleDeg < 0) angleDeg += 360;
            var nearestAxisDeg = Math.round(angleDeg / 90) * 90;
            var deviationDeg = angleDeg - nearestAxisDeg;
            var absDeviationDeg = Math.abs(deviationDeg);
            if (absDeviationDeg >= 0.5 && absDeviationDeg <= 10) {
                safeCall(function () { rect.rotate(-deviationDeg); });
            }
        }

        var bounds = rect.geometricBounds;
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        var rectWidth = right - left;
        var rectHeight = top - bottom;

        var diffRatio = Math.abs(rectWidth - rectHeight) / Math.max(rectWidth, rectHeight);
        if (diffRatio < 0.05) return null;
        var shortSide = Math.min(rectWidth, rectHeight);
        var longSide = Math.max(rectWidth, rectHeight);
        if (shortSide * 1.5 > longSide) return null;
        if (typeof minShortSidePt === "number" && shortSide < minShortSidePt) return null;

        var parentLayer = rect.layer;
        if (!parentLayer) return null;

        var centerLine = parentLayer.pathItems.add();
        centerLine.stroked = true;
        centerLine.filled = false;
        safeCall(function () { centerLine.strokeColor = rect.fillColor; });

        var startPoint = centerLine.pathPoints.add();
        var endPoint = centerLine.pathPoints.add();

        if (rectHeight <= rectWidth) {
            var centerY = (top + bottom) / 2;
            startPoint.anchor = [left, centerY];
            endPoint.anchor = [right, centerY];
        } else {
            var centerX = (left + right) / 2;
            startPoint.anchor = [centerX, top];
            endPoint.anchor = [centerX, bottom];
        }

        centerLine.strokeWidth = (rectHeight <= rectWidth) ? rectHeight : rectWidth;

        startPoint.leftDirection = startPoint.anchor;
        startPoint.rightDirection = startPoint.anchor;
        endPoint.leftDirection = endPoint.anchor;
        endPoint.rightDirection = endPoint.anchor;

        safeCall(function () { rect.remove(); });
        return centerLine;
    }

    /* 最頻フォントサイズ取得 / Get most frequent text size */
    function getModeTextSize(doc) {
        var counts = {};
        for (var t = 0; t < doc.textFrames.length; t++) {
            try {
                var chars = doc.textFrames[t].characters;
                for (var c = 0; c < chars.length; c++) {
                    try {
                        var key = Math.round(chars[c].size * 10) / 10;
                        counts[key] = (counts[key] || 0) + 1;
                    } catch (e) { }
                }
            } catch (e) { }
        }
        var modeSize = 0, modeCount = 0;
        for (var k in counts) {
            if (counts.hasOwnProperty(k) && counts[k] > modeCount) {
                modeCount = counts[k];
                modeSize = parseFloat(k);
            }
        }
        return modeSize;
    }

    /* 長方形判定 / Check if path is rectangle-like */
    function isRectangleLike(item) {
        if (!item || item.typename !== "PathItem") return false;
        if (!item.closed) return false;
        if (item.pathPoints.length !== 4) return false;
        return true;
    }

    /* 配列内存在チェック / Check if item exists in array */
    function containsItem(arr, item) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === item) return true;
        }
        return false;
    }

    /* 矩形収集 / Collect rectangles from document */
    function collectRectangles(doc, opts) {
        opts = opts || {};
        var excludeNames = opts.excludeLayerNames || [];
        var shortSideMin = (typeof opts.shortSideMin === "number") ? opts.shortSideMin : 0;
        var customFilter = opts.filter || null;
        var result = [];

        function walk(items) {
            for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
                var pageItem = items[itemIndex];
                try {
                    if (pageItem.typename === "GroupItem") {
                        walk(pageItem.pageItems);
                    } else if (isRectangleLike(pageItem)) {
                        var itemWidth = pageItem.width;
                        var itemHeight = pageItem.height;
                        var shortSide = (itemWidth < itemHeight) ? itemWidth : itemHeight;
                        if (shortSide > shortSideMin) {
                            if (customFilter === null || customFilter(pageItem)) {
                                result.push(pageItem);
                            }
                        }
                    }
                } catch (e) { }
            }
        }

        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var layer = doc.layers[layerIndex];
            var shouldSkipLayer = false;
            for (var excludeIndex = 0; excludeIndex < excludeNames.length; excludeIndex++) {
                if (layer.name === excludeNames[excludeIndex]) { shouldSkipLayer = true; break; }
            }
            if (shouldSkipLayer || layer.locked || !layer.visible) continue;
            walk(layer.pageItems);
        }
        return result;
    }

    /* 背面レイヤー確保 / Ensure layer exists and send to back */
    function ensureBackLayer(doc, layerName) {
        var layer = null;
        for (var l = 0; l < doc.layers.length; l++) {
            if (doc.layers[l].name === layerName) { layer = doc.layers[l]; break; }
        }
        if (layer === null) {
            layer = doc.layers.add();
            layer.name = layerName;
        }
        safeCall(function () { layer.zOrder(ZOrderMethod.SENDTOBACK); });
        return layer;
    }

    /* アピアランス展開対象抽出 / Expand appearance targets */
    function expandByAppearance(doc, items) {
        var found = [];
        for (var n = 0; n < items.length; n++) {
            try {
                doc.selection = null;
                items[n].selected = true;
                app.executeMenuCommand('Find Appearance menu item');
                for (var j = 0; j < doc.selection.length; j++) {
                    if (!containsItem(found, doc.selection[j])) {
                        found.push(doc.selection[j]);
                    }
                }
            } catch (e) { }
        }
        return found;
    }

    /* アイテム移動 / Move items to layer */
    function moveItemsToLayer(items, layer) {
        for (var itemIndex = items.length - 1; itemIndex >= 0; itemIndex--) {
            (function (index) {
                safeCall(function () { items[index].move(layer, ElementPlacement.PLACEATBEGINNING); });
            })(itemIndex);
        }
    }

    /* 選択設定 / Set selection */
    function setSelection(doc, items) {
        doc.selection = null;
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            (function (index) {
                safeCall(function () { items[index].selected = true; });
            })(itemIndex);
        }
    }

    /*
     目的:
     セル背景（矩形）は分割・重複・見た目効果（アピアランス）で散在している。
     対象矩形を抽出→同一レイヤーへ集約→グループ化することで、
     後続処理（罫線スナップなど）が扱いやすい状態に正規化する。

     Purpose:
     Cell backgrounds are fragmented and affected by appearances.
     This collects target rectangles, moves them to a single layer, and groups
     them so downstream steps (rule snapping, etc.) can work with them cleanly.
    */
    /* セル背景抽出＋集約 / Auto select and group shapes */
    function autoSelectAndMerge(opts) {
        opts = opts || {};
        if (app.documents.length === 0) return null;
        var doc = app.activeDocument;
        var layerName = opts.layerName || PATH_LAYER_NAME;

        var threshold;
        if (typeof opts.shortSideMin === "function") {
            threshold = opts.shortSideMin(doc);
        } else if (typeof opts.shortSideMin === "number") {
            threshold = opts.shortSideMin;
        } else {
            threshold = getModeTextSize(doc);
        }
        if (!threshold || threshold <= 0) return null;

        var excludeNames = [layerName];
        if (opts.excludeLayerNames) {
            for (var ex = 0; ex < opts.excludeLayerNames.length; ex++) {
                excludeNames.push(opts.excludeLayerNames[ex]);
            }
        }

        var rectangles = collectRectangles(doc, {
            excludeLayerNames: excludeNames,
            shortSideMin: threshold
        });
        if (rectangles.length === 0) return null;

        setSelection(doc, rectangles);

        var targets;
        if (opts.runFindAppearance === false) {
            targets = rectangles;
        } else {
            targets = expandByAppearance(doc, rectangles);
            if (targets.length === 0) return null;
        }

        var targetLayer = ensureBackLayer(doc, layerName);
        moveItemsToLayer(targets, targetLayer);
        setSelection(doc, targets);

        if (opts.runGroup !== false) safeCall(function () { app.executeMenuCommand('group'); });

        return targets;
    }

    main();
})();
