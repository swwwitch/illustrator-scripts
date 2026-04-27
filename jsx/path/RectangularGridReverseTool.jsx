#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要
選択した水平線・垂直線を、長方形グリッド状に再構成する Illustrator スクリプトです。
主な機能：
  配置：均等配置しない、均等に（強制）、均等＋結合セル対応を選択できます。
  線：突出線端、破線の実線化、線幅（最大・最小・平均・指定値）を設定できます。
  後処理：外枠の長方形化、グループ化を選択できます。
  プレビュー：配置・線端・線幅の変更結果を、ダイアログを閉じずに確認できます。
Overview
This Illustrator script reconstructs selected horizontal and vertical lines into a rectangular grid structure.
Main features:
  Distribution: Choose Do not distribute evenly, Evenly (force), or Evenly + merged cells.
  Line: Set projecting caps, convert dashed lines to solid, and set stroke width (maximum, minimum, average, or specified value).
  Post-processing: Convert the outer frame to a rectangle and group processed items.
  Preview: Check distribution, line cap, and stroke width changes without closing the dialog.
*/

(function () {
    // =========================================
    // バージョンとローカライズ / Version and localization
    // =========================================

    var SCRIPT_VERSION = "v1.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "逆・長方形グリッドツール",
            en: "Reverse Rectangle Grid Tool"
        },
        optionEven: {
            ja: "均等に（強制）",
            en: "Evenly (force)"
        },
        panelDistribution: {
            ja: "配置モード",
            en: "Distribution Mode"
        },
        optionDistributionNone: {
            ja: "均等配置しない",
            en: "Do not distribute evenly"
        },
        optionEvenMergedCell: {
            ja: "均等＋結合セル対応",
            en: "Evenly + merged cells"
        },
        panelLine: {
            ja: "線",
            en: "Line"
        },
        panelStrokeWidth: {
            ja: "線幅",
            en: "Stroke Width"
        },
        optionStrokeWidthMax: {
            ja: "最大",
            en: "Maximum"
        },
        optionStrokeWidthMin: {
            ja: "最小",
            en: "Minimum"
        },
        optionStrokeWidthAverage: {
            ja: "平均",
            en: "Average"
        },
        optionProjectingCap: {
            ja: "突出線端にする",
            en: "Projecting cap"
        },
        optionConvertDashedToSolid: {
            ja: "破線を実線にする",
            en: "Convert dashed lines to solid"
        },
        optionStrokeWidthSpecified: {
            ja: "指定",
            en: "Specified"
        },
        panelPostProcessing: {
            ja: "後処理",
            en: "Post-processing"
        },
        optionFrameToRect: {
            ja: "外枠を長方形に変換",
            en: "Convert outer frame to rectangle"
        },
        optionGroup: {
            ja: "グループ化",
            en: "Group"
        },
        buttonCancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        buttonOK: {
            ja: "OK",
            en: "OK"
        },
        optionPreview: {
            ja: "プレビュー",
            en: "Preview"
        },
        errorNoSelection: {
            ja: "水平線と垂直線を選択してください。",
            en: "Select horizontal and vertical lines."
        },
        errorNotEnoughLines: {
            ja: "格子を作るには、最低でも2本の水平線と2本の垂直線が必要です。",
            en: "To create a grid, select at least two horizontal lines and two vertical lines."
        }
    };

    /* ラベル取得 / Get localized label */
    function getLabel(key) {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
        return key;
    }

    // =========================================
    // 単位ユーティリティ / Unit utilities
    // =========================================

    /**
     * Illustrator 単位ユーティリティ関数群 / Illustrator unit utility functions
     *
     * 設定キーの意味 / Preference keys:
     * - "rulerType"         ：一般（定規の単位）/ General ruler unit
     * - "strokeUnits"       ：線 / Stroke unit
     * - "text/units"        ：文字 / Text unit
     * - "text/asianunits"   ：東アジア言語のオプション / East Asian text unit
     */

    // 単位コードとラベルのマップ / Unit code and label map
    var unitMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    // 単位コードと設定キーから適切な単位ラベルを返す / Return the appropriate unit label from a unit code and preference key
    function getUnitLabel(code, prefKey) {
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return unitMap[code] || "pt";
    }

    // 単位コードから pt 換算係数を返す / Return the point conversion factor from a unit code
    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H
            case 6: return 1.0;                         // px
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    // 設定キーから単位情報を取得 / Get unit information from a preference key
    function getUnitInfoFromPreference(prefKey) {
        var unitCode = 2;
        try {
            unitCode = app.preferences.getIntegerPreference(prefKey);
        } catch (unitError) {
            unitCode = 2;
        }
        return {
            code: unitCode,
            label: getUnitLabel(unitCode, prefKey),
            factor: getPtFactorFromUnitCode(unitCode)
        };
    }

    // 指定単位の値をpt値に変換 / Convert unit value to point value
    function unitValueToPoints(unitValue, unitInfo) {
        return unitValue * unitInfo.factor;
    }

    if (app.documents.length === 0) return;

    var selectedItems = app.activeDocument.selection;
    if (!selectedItems || selectedItems.length === 0) {
        alert(getLabel("errorNoSelection"));
        return;
    }

    var classified = classifySelectedStraightLines(selectedItems);
    if (classified.horizontalLines.length < 2 || classified.verticalLines.length < 2) {
        alert(getLabel("errorNotEnoughLines"));
        return;
    }

    var gridBounds = getGridBounds(classified.horizontalLines, classified.verticalLines);
    var originalLineStates = captureLineStates(classified.horizontalLines, classified.verticalLines);
    var strokeUnitInfo = getUnitInfoFromPreference("strokeUnits");

    var dialogOptions = showOptionDialog(classified, gridBounds, originalLineStates, strokeUnitInfo);
    if (!dialogOptions) return;

    restoreLineStates(originalLineStates);
    applyRepresentativeStrokeWidth(classified.horizontalLines, classified.verticalLines, dialogOptions);
    if (dialogOptions.convertDashedToSolid) convertDashedLinesToSolid(classified.horizontalLines, classified.verticalLines);
    alignLinesToGridBounds(classified.horizontalLines, classified.verticalLines, gridBounds, dialogOptions);

    // 外枠を長方形に変換 → 残存パス＋長方形をグループ化 / Convert outer frame to rectangle, then group remaining paths and rectangle
    var groupTargets = [];
    for (var horizontalIndex = 0; horizontalIndex < classified.horizontalLines.length; horizontalIndex++) groupTargets.push(classified.horizontalLines[horizontalIndex].path);
    for (var verticalIndex = 0; verticalIndex < classified.verticalLines.length; verticalIndex++) groupTargets.push(classified.verticalLines[verticalIndex].path);

    if (dialogOptions.frameToRect) {
        var outerPaths = pickOuterPaths(classified.horizontalLines, classified.verticalLines);
        var frameRectangle = createOuterFrameRectangle(outerPaths.referencePath, gridBounds);
        // outerPaths.uniqueOuterPaths を groupTargets から除外し frameRectangle を追加 / Exclude outerPaths.uniqueOuterPaths from groupTargets and add frameRectangle
        var innerGridItems = [];
        for (var targetIndex = 0; targetIndex < groupTargets.length; targetIndex++) {
            var isOuterPath = false;
            for (var outerPathIndex = 0; outerPathIndex < outerPaths.uniqueOuterPaths.length; outerPathIndex++) {
                if (outerPaths.uniqueOuterPaths[outerPathIndex] === groupTargets[targetIndex]) { isOuterPath = true; break; }
            }
            if (!isOuterPath) innerGridItems.push(groupTargets[targetIndex]);
        }
        innerGridItems.push(frameRectangle);
        groupTargets = innerGridItems;
        // 参照が無効になる前にフィルタを終え、最後に削除 / Finish filtering before references become invalid, then remove them at the end
        for (var removeIndex = 0; removeIndex < outerPaths.uniqueOuterPaths.length; removeIndex++) {
            try { outerPaths.uniqueOuterPaths[removeIndex].remove(); } catch (removeError) { }
        }
    }

    if (dialogOptions.group && groupTargets.length > 1) {
        groupProcessedItems(groupTargets);
    }

    // 選択アイテムから直線パス（アンカー2点）を抽出し、水平/垂直に分類 / Extract straight paths (2 anchors) from selected items and classify them as horizontal or vertical
    function classifySelectedStraightLines(selectedItems) {
        var horizontalLines = [];
        var verticalLines = [];
        for (var itemIndex = 0; itemIndex < selectedItems.length; itemIndex++) {
            var pathItem = selectedItems[itemIndex];
            if (!pathItem || pathItem.typename !== "PathItem" || pathItem.pathPoints.length !== 2) continue;
            var firstAnchor = pathItem.pathPoints[0].anchor;
            var secondAnchor = pathItem.pathPoints[1].anchor;
            var horizontalDistance = Math.abs(firstAnchor[0] - secondAnchor[0]);
            var verticalDistance = Math.abs(firstAnchor[1] - secondAnchor[1]);
            if (horizontalDistance >= verticalDistance) {
                horizontalLines.push({
                    path: pathItem,
                    y: (firstAnchor[1] + secondAnchor[1]) / 2,
                    minX: Math.min(firstAnchor[0], secondAnchor[0]),
                    maxX: Math.max(firstAnchor[0], secondAnchor[0])
                });
            } else {
                verticalLines.push({
                    path: pathItem,
                    x: (firstAnchor[0] + secondAnchor[0]) / 2,
                    minY: Math.min(firstAnchor[1], secondAnchor[1]),
                    maxY: Math.max(firstAnchor[1], secondAnchor[1])
                });
            }
        }
        return { horizontalLines: horizontalLines, verticalLines: verticalLines };
    }

    // 格子の外周座標（左端・右端・上端・下端）を計算 / Calculate grid outer coordinates (left, right, top, bottom)
    function getGridBounds(horizontalLines, verticalLines) {
        var minX = verticalLines[0].x, maxX = verticalLines[0].x;
        for (var verticalIndex = 1; verticalIndex < verticalLines.length; verticalIndex++) {
            if (verticalLines[verticalIndex].x < minX) minX = verticalLines[verticalIndex].x;
            if (verticalLines[verticalIndex].x > maxX) maxX = verticalLines[verticalIndex].x;
        }
        var minY = horizontalLines[0].y, maxY = horizontalLines[0].y;
        for (var horizontalIndex = 1; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            if (horizontalLines[horizontalIndex].y < minY) minY = horizontalLines[horizontalIndex].y;
            if (horizontalLines[horizontalIndex].y > maxY) maxY = horizontalLines[horizontalIndex].y;
        }
        return { minX: minX, maxX: maxX, minY: minY, maxY: maxY };
    }

    // 2点の直線パスの両端を設定（方向ハンドルもアンカー位置に揃える）/ Set both endpoints of a 2-point straight path (also align direction handles to anchors)
    function setLineEndpoints(pathItem, firstPoint, secondPoint) {
        pathItem.pathPoints[0].anchor = firstPoint;
        pathItem.pathPoints[0].leftDirection = firstPoint;
        pathItem.pathPoints[0].rightDirection = firstPoint;
        pathItem.pathPoints[1].anchor = secondPoint;
        pathItem.pathPoints[1].leftDirection = secondPoint;
        pathItem.pathPoints[1].rightDirection = secondPoint;
    }

    // プレビュー復元用に線の状態を記録 / Capture line states for preview restoration
    function captureLineStates(horizontalLines, verticalLines) {
        var lineStates = [];
        collectLineStates(horizontalLines, lineStates);
        collectLineStates(verticalLines, lineStates);
        return lineStates;
    }

    // 線情報を配列に追加 / Add line information to an array
    function collectLineStates(lines, lineStates) {
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var pathItem = lines[lineIndex].path;
            lineStates.push({
                path: pathItem,
                firstAnchor: pathItem.pathPoints[0].anchor,
                firstLeftDirection: pathItem.pathPoints[0].leftDirection,
                firstRightDirection: pathItem.pathPoints[0].rightDirection,
                secondAnchor: pathItem.pathPoints[1].anchor,
                secondLeftDirection: pathItem.pathPoints[1].leftDirection,
                secondRightDirection: pathItem.pathPoints[1].rightDirection,
                stroked: pathItem.stroked,
                strokeCap: pathItem.strokeCap,
                strokeWidth: pathItem.strokeWidth,
                strokeDashes: pathItem.strokeDashes,
                strokeDashOffset: pathItem.strokeDashOffset
            });
        }
    }

    // 記録した線の状態を復元 / Restore captured line states
    function restoreLineStates(lineStates) {
        for (var lineIndex = 0; lineIndex < lineStates.length; lineIndex++) {
            var lineState = lineStates[lineIndex];
            var pathItem = lineState.path;
            pathItem.pathPoints[0].anchor = lineState.firstAnchor;
            pathItem.pathPoints[0].leftDirection = lineState.firstLeftDirection;
            pathItem.pathPoints[0].rightDirection = lineState.firstRightDirection;
            pathItem.pathPoints[1].anchor = lineState.secondAnchor;
            pathItem.pathPoints[1].leftDirection = lineState.secondLeftDirection;
            pathItem.pathPoints[1].rightDirection = lineState.secondRightDirection;
            pathItem.stroked = lineState.stroked;
            pathItem.strokeCap = lineState.strokeCap;
            pathItem.strokeWidth = lineState.strokeWidth;
            pathItem.strokeDashes = lineState.strokeDashes;
            pathItem.strokeDashOffset = lineState.strokeDashOffset;
        }
    }

    // プレビューを更新 / Update preview
    function updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates) {
        restoreLineStates(originalLineStates);
        if (!ui.previewCheckbox.value) {
            app.redraw();
            return;
        }
        var previewOptions = readOptionDialogState(ui);
        // プレビュー時は frameToRect, group を適用しない
        previewOptions.frameToRect = false;
        previewOptions.group = false;
        applyRepresentativeStrokeWidth(classified.horizontalLines, classified.verticalLines, previewOptions);
        if (previewOptions.convertDashedToSolid) convertDashedLinesToSolid(classified.horizontalLines, classified.verticalLines);
        alignLinesToGridBounds(classified.horizontalLines, classified.verticalLines, gridBounds, previewOptions);
        app.redraw();
    }

    // ダイアログの現在値を読む / Read current dialog state
    function readOptionDialogState(ui) {
        return {
            distributionMode: ui.distributionNoneRadio.value ? "none"
                : ui.distributionEvenMergedCellRadio.value ? "evenMergedCell"
                    : "even",
            even: !ui.distributionNoneRadio.value,
            projectingCap: ui.projectingCapCheckbox.value,
            convertDashedToSolid: ui.convertDashedToSolidCheckbox.value,
            strokeWidthMode: ui.strokeWidthMaxRadio.value ? "max"
                : ui.strokeWidthMinRadio.value ? "min"
                    : ui.strokeWidthSpecifiedRadio.value ? "specified"
                        : "average",
            specifiedStrokeWidthPt: ui.strokeWidthSpecifiedRadio && ui.strokeWidthSpecifiedRadio.value
                ? unitValueToPoints(readNumericText(ui.strokeWidthInput.text, 0), ui.strokeUnitInfo)
                : 0,
            frameToRect: ui.frameToRectangleCheckbox.value,
            group: ui.groupingCheckbox.value
        };
    }

    // 数値文字列を安全に読む / Safely read numeric text
    function readNumericText(text, fallbackValue) {
        var normalizedText = String(text).replace(/,/g, ".");
        var value = parseFloat(normalizedText);
        if (isNaN(value)) return fallbackValue;
        return value;
    }

    // 指定線幅入力の有効状態を同期 / Sync specified stroke width input enabled state
    function syncSpecifiedStrokeWidthInput(ui) {
        ui.strokeWidthInput.enabled = ui.strokeWidthSpecifiedRadio.value;
    }

    // ↑↓キーで数値を増減 / Change numeric value with Up/Down arrow keys
    function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;
            var shouldUpdateValue = false;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is pressed
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    shouldUpdateValue = true;
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    shouldUpdateValue = true;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減 / Change by 0.1 when Option is pressed
                if (event.keyName === "Up") {
                    value += delta;
                    shouldUpdateValue = true;
                } else if (event.keyName === "Down") {
                    value -= delta;
                    shouldUpdateValue = true;
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    shouldUpdateValue = true;
                } else if (event.keyName === "Down") {
                    value -= delta;
                    shouldUpdateValue = true;
                }
            }

            if (!shouldUpdateValue) return;

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10; // 小数第1位まで / Round to 1 decimal
            } else {
                value = Math.round(value); // 整数に丸め / Round to integer
            }

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            // プレビュー更新 / Update preview
            if (typeof onValueChanged === "function") onValueChanged();
        });
    }

    // 代表線幅を適用 / Apply representative stroke width
    function applyRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions) {
        var representativeStrokeWidth = getRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions);
        if (representativeStrokeWidth <= 0) return;
        applyStrokeWidthToLines(horizontalLines, representativeStrokeWidth);
        applyStrokeWidthToLines(verticalLines, representativeStrokeWidth);
    }

    // 代表線幅を取得 / Get representative stroke width
    function getRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions) {
        if (dialogOptions.strokeWidthMode === "specified") {
            return dialogOptions.specifiedStrokeWidthPt;
        }

        var strokeWidths = [];
        collectStrokeWidths(horizontalLines, strokeWidths);
        collectStrokeWidths(verticalLines, strokeWidths);
        if (strokeWidths.length === 0) return 0;

        if (dialogOptions.strokeWidthMode === "max") {
            var maxStrokeWidth = strokeWidths[0];
            for (var maxIndex = 1; maxIndex < strokeWidths.length; maxIndex++) {
                if (strokeWidths[maxIndex] > maxStrokeWidth) maxStrokeWidth = strokeWidths[maxIndex];
            }
            return maxStrokeWidth;
        }

        if (dialogOptions.strokeWidthMode === "min") {
            var minStrokeWidth = strokeWidths[0];
            for (var minIndex = 1; minIndex < strokeWidths.length; minIndex++) {
                if (strokeWidths[minIndex] < minStrokeWidth) minStrokeWidth = strokeWidths[minIndex];
            }
            return minStrokeWidth;
        }

        var totalStrokeWidth = 0;
        for (var averageIndex = 0; averageIndex < strokeWidths.length; averageIndex++) {
            totalStrokeWidth += strokeWidths[averageIndex];
        }
        return totalStrokeWidth / strokeWidths.length;
    }

    // 線幅一覧を収集 / Collect stroke widths
    function collectStrokeWidths(lines, strokeWidths) {
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var pathItem = lines[lineIndex].path;
            if (pathItem.stroked && pathItem.strokeWidth > 0) strokeWidths.push(pathItem.strokeWidth);
        }
    }

    // 線群に線幅を適用 / Apply stroke width to lines
    function applyStrokeWidthToLines(lines, strokeWidth) {
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var pathItem = lines[lineIndex].path;
            pathItem.stroked = true;
            pathItem.strokeWidth = strokeWidth;
        }
    }

    // 配列要素の指定プロパティが targetValue に最も近いインデックスを返す / Return the index whose property is closest to targetValue
    function findClosestIndexByProperty(targetValue, items, propertyName) {
        var closestIndex = 0;
        var closestDistance = Math.abs(targetValue - items[0][propertyName]);
        for (var itemIndex = 1; itemIndex < items.length; itemIndex++) {
            var currentDistance = Math.abs(targetValue - items[itemIndex][propertyName]);
            if (currentDistance < closestDistance) {
                closestIndex = itemIndex;
                closestDistance = currentDistance;
            }
        }
        return closestIndex;
    }

    // 同じ座標とみなす許容誤差（pt） / Tolerance to treat coordinates as identical (pt)
    var COORDINATE_TOLERANCE = 5.0;

    // 線群を指定プロパティで近接グループ化（昇順）/ Group lines by proximity along a property (ascending)
    function groupLinesByCoordinate(lines, propertyName, tolerance) {
        var sortedLines = lines.slice();
        sortedLines.sort(function (lineA, lineB) { return lineA[propertyName] - lineB[propertyName]; });
        var coordinateGroups = [];
        for (var lineIndex = 0; lineIndex < sortedLines.length; lineIndex++) {
            var currentLine = sortedLines[lineIndex];
            var currentCoordinate = currentLine[propertyName];
            if (coordinateGroups.length === 0 || Math.abs(currentCoordinate - coordinateGroups[coordinateGroups.length - 1].coord) > tolerance) {
                coordinateGroups.push({ coord: currentCoordinate, lines: [currentLine] });
            } else {
                var lastCoordinateGroup = coordinateGroups[coordinateGroups.length - 1];
                lastCoordinateGroup.lines.push(currentLine);
                var coordinateSum = 0;
                for (var groupedLineIndex = 0; groupedLineIndex < lastCoordinateGroup.lines.length; groupedLineIndex++) {
                    coordinateSum += lastCoordinateGroup.lines[groupedLineIndex][propertyName];
                }
                lastCoordinateGroup.coord = coordinateSum / lastCoordinateGroup.lines.length;
            }
        }
        return coordinateGroups;
    }

    // 水平線・垂直線を格子に揃える / Align horizontal and vertical lines to the grid
    function alignLinesToGridBounds(horizontalLines, verticalLines, gridBounds, dialogOptions) {
        var isMergedCellMode = dialogOptions.distributionMode === "evenMergedCell";

        if (isMergedCellMode) {
            alignLinesWithMergedCellSupport(horizontalLines, verticalLines, gridBounds, dialogOptions);
            return;
        }

        // 水平線の Y 位置を決定（均等配置 or 元の Y）/ Determine Y positions for horizontal lines
        var horizontalLineYPositions = [];
        if (dialogOptions.even && horizontalLines.length >= 2) {
            horizontalLines.sort(function (lineA, lineB) { return lineA.y - lineB.y; });
            var horizontalStep = (gridBounds.maxY - gridBounds.minY) / (horizontalLines.length - 1);
            for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
                horizontalLineYPositions.push(gridBounds.minY + horizontalStep * horizontalIndex);
            }
        } else {
            for (var originalHorizontalIndex = 0; originalHorizontalIndex < horizontalLines.length; originalHorizontalIndex++) {
                horizontalLineYPositions.push(horizontalLines[originalHorizontalIndex].y);
            }
        }

        // 垂直線の X 位置を決定（均等配置 or 元の X）/ Determine X positions for vertical lines
        var verticalLineXPositions = [];
        if (dialogOptions.even && verticalLines.length >= 2) {
            verticalLines.sort(function (lineA, lineB) { return lineA.x - lineB.x; });
            var verticalStep = (gridBounds.maxX - gridBounds.minX) / (verticalLines.length - 1);
            for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
                verticalLineXPositions.push(gridBounds.minX + verticalStep * verticalIndex);
            }
        } else {
            for (var originalVerticalIndex = 0; originalVerticalIndex < verticalLines.length; originalVerticalIndex++) {
                verticalLineXPositions.push(verticalLines[originalVerticalIndex].x);
            }
        }

        // Illustrator は Y 上方向が正 / Illustrator uses positive Y upward
        var topY = Math.max(gridBounds.minY, gridBounds.maxY);
        var bottomY = Math.min(gridBounds.minY, gridBounds.maxY);

        for (var horizontalPathIndex = 0; horizontalPathIndex < horizontalLines.length; horizontalPathIndex++) {
            var horizontalPath = horizontalLines[horizontalPathIndex].path;
            if (dialogOptions.projectingCap) horizontalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
            setLineEndpoints(horizontalPath, [gridBounds.minX, horizontalLineYPositions[horizontalPathIndex]], [gridBounds.maxX, horizontalLineYPositions[horizontalPathIndex]]);
        }
        for (var verticalPathIndex = 0; verticalPathIndex < verticalLines.length; verticalPathIndex++) {
            var verticalPath = verticalLines[verticalPathIndex].path;
            if (dialogOptions.projectingCap) verticalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
            setLineEndpoints(verticalPath, [verticalLineXPositions[verticalPathIndex], topY], [verticalLineXPositions[verticalPathIndex], bottomY]);
        }
    }

    // 結合セル対応モード：同じ座標の線を 1 行/1 列としてまとめ、両端は直近の交差点にスナップ / Merged-cell mode: lines on the same coordinate share one row/column; endpoints snap to the nearest intersection
    function alignLinesWithMergedCellSupport(horizontalLines, verticalLines, gridBounds, dialogOptions) {
        var horizontalGroups = groupLinesByCoordinate(horizontalLines, "y", COORDINATE_TOLERANCE);
        var verticalGroups = groupLinesByCoordinate(verticalLines, "x", COORDINATE_TOLERANCE);

        // 各グループの新しい Y 位置（行）/ New Y position for each horizontal group (row)
        var horizontalGroupYPositions = [];
        if (dialogOptions.even && horizontalGroups.length >= 2) {
            var horizontalStep = (gridBounds.maxY - gridBounds.minY) / (horizontalGroups.length - 1);
            for (var horizontalGroupIndex = 0; horizontalGroupIndex < horizontalGroups.length; horizontalGroupIndex++) {
                horizontalGroupYPositions.push(gridBounds.minY + horizontalStep * horizontalGroupIndex);
            }
        } else {
            for (var originalHorizontalGroupIndex = 0; originalHorizontalGroupIndex < horizontalGroups.length; originalHorizontalGroupIndex++) {
                horizontalGroupYPositions.push(horizontalGroups[originalHorizontalGroupIndex].coord);
            }
        }

        // 各グループの新しい X 位置（列）/ New X position for each vertical group (column)
        var verticalGroupXPositions = [];
        if (dialogOptions.even && verticalGroups.length >= 2) {
            var verticalStep = (gridBounds.maxX - gridBounds.minX) / (verticalGroups.length - 1);
            for (var verticalGroupIndex = 0; verticalGroupIndex < verticalGroups.length; verticalGroupIndex++) {
                verticalGroupXPositions.push(gridBounds.minX + verticalStep * verticalGroupIndex);
            }
        } else {
            for (var originalVerticalGroupIndex = 0; originalVerticalGroupIndex < verticalGroups.length; originalVerticalGroupIndex++) {
                verticalGroupXPositions.push(verticalGroups[originalVerticalGroupIndex].coord);
            }
        }

        // 水平線：同一行のすべてに同じ Y、両端は最寄り列にスナップ / Apply same Y to all lines in a row; snap endpoints to nearest column
        for (var horizontalGroupIndex = 0; horizontalGroupIndex < horizontalGroups.length; horizontalGroupIndex++) {
            var rowY = horizontalGroupYPositions[horizontalGroupIndex];
            var rowLines = horizontalGroups[horizontalGroupIndex].lines;
            for (var rowLineIndex = 0; rowLineIndex < rowLines.length; rowLineIndex++) {
                var horizontalLineData = rowLines[rowLineIndex];
                var horizontalPath = horizontalLineData.path;
                if (dialogOptions.projectingCap) horizontalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
                var leftColumnIndex = findClosestIndexByProperty(horizontalLineData.minX, verticalGroups, "coord");
                var rightColumnIndex = findClosestIndexByProperty(horizontalLineData.maxX, verticalGroups, "coord");
                setLineEndpoints(horizontalPath, [verticalGroupXPositions[leftColumnIndex], rowY], [verticalGroupXPositions[rightColumnIndex], rowY]);
            }
        }

        // 垂直線：同一列のすべてに同じ X、両端は最寄り行にスナップ / Apply same X to all lines in a column; snap endpoints to nearest row
        for (var verticalGroupIndex = 0; verticalGroupIndex < verticalGroups.length; verticalGroupIndex++) {
            var columnX = verticalGroupXPositions[verticalGroupIndex];
            var columnLines = verticalGroups[verticalGroupIndex].lines;
            for (var columnLineIndex = 0; columnLineIndex < columnLines.length; columnLineIndex++) {
                var verticalLineData = columnLines[columnLineIndex];
                var verticalPath = verticalLineData.path;
                if (dialogOptions.projectingCap) verticalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
                var topRowIndex = findClosestIndexByProperty(verticalLineData.maxY, horizontalGroups, "coord");
                var bottomRowIndex = findClosestIndexByProperty(verticalLineData.minY, horizontalGroups, "coord");
                setLineEndpoints(verticalPath, [columnX, horizontalGroupYPositions[topRowIndex]], [columnX, horizontalGroupYPositions[bottomRowIndex]]);
            }
        }
    }

    // 外周4本（上下水平線・左右垂直線）を選び、重複排除した参照配列を返す / Pick the four outer lines (top/bottom horizontal, left/right vertical) and return deduplicated references
    function pickOuterPaths(horizontalLines, verticalLines) {
        var topLine = null, bottomLine = null, leftLine = null, rightLine = null;
        for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            if (topLine === null || horizontalLines[horizontalIndex].y > topLine.y) topLine = horizontalLines[horizontalIndex];
            if (bottomLine === null || horizontalLines[horizontalIndex].y < bottomLine.y) bottomLine = horizontalLines[horizontalIndex];
        }
        for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
            if (leftLine === null || verticalLines[verticalIndex].x < leftLine.x) leftLine = verticalLines[verticalIndex];
            if (rightLine === null || verticalLines[verticalIndex].x > rightLine.x) rightLine = verticalLines[verticalIndex];
        }
        var outerPathCandidates = [topLine.path, bottomLine.path, leftLine.path, rightLine.path];
        var uniqueOuterPaths = [];
        for (var candidateIndex = 0; candidateIndex < outerPathCandidates.length; candidateIndex++) {
            var isDuplicate = false;
            for (var uniqueIndex = 0; uniqueIndex < uniqueOuterPaths.length; uniqueIndex++) {
                if (uniqueOuterPaths[uniqueIndex] === outerPathCandidates[candidateIndex]) { isDuplicate = true; break; }
            }
            if (!isDuplicate) uniqueOuterPaths.push(outerPathCandidates[candidateIndex]);
        }
        return { uniqueOuterPaths: uniqueOuterPaths, referencePath: topLine.path };
    }

    // 外周長方形を作成（線属性は referencePath から引き継ぐ）/ Create the outer rectangle (inherit stroke attributes from referencePath)
    function createOuterFrameRectangle(referencePath, gridBounds) {
        var parentLayerOrGroup = referencePath.parent;
        // rectangle(top, left, width, height) — Illustrator は Y 上方向が正 / Illustrator uses positive Y upward
        var frameRectangle = parentLayerOrGroup.pathItems.rectangle(
            gridBounds.maxY, gridBounds.minX,
            gridBounds.maxX - gridBounds.minX, gridBounds.maxY - gridBounds.minY
        );
        frameRectangle.filled = referencePath.filled;
        if (referencePath.filled) frameRectangle.fillColor = referencePath.fillColor;
        frameRectangle.stroked = referencePath.stroked;
        if (referencePath.stroked) {
            frameRectangle.strokeColor = referencePath.strokeColor;
            frameRectangle.strokeWidth = referencePath.strokeWidth;
            frameRectangle.strokeCap = referencePath.strokeCap;
            frameRectangle.strokeJoin = referencePath.strokeJoin;
            frameRectangle.strokeMiterLimit = referencePath.strokeMiterLimit;
            frameRectangle.strokeDashes = referencePath.strokeDashes;
            frameRectangle.strokeDashOffset = referencePath.strokeDashOffset;
        }
        frameRectangle.move(referencePath, ElementPlacement.PLACEAFTER);
        return frameRectangle;
    }

    // 渡されたアイテムを1つのグループにまとめる / Group the given items into one group
    function groupProcessedItems(processedItems) {
        var parent = processedItems[0].parent;
        var group = parent.groupItems.add();
        for (var itemIndex = 0; itemIndex < processedItems.length; itemIndex++) {
            processedItems[itemIndex].move(group, ElementPlacement.PLACEATEND);
        }
        return group;
    }

    // 破線を実線に変換 / Convert dashed lines to solid lines
    function convertDashedLinesToSolid(horizontalLines, verticalLines) {
        applySolidStrokeToLines(horizontalLines);
        applySolidStrokeToLines(verticalLines);
    }

    // 線群の破線設定を解除 / Clear dash settings from lines
    function applySolidStrokeToLines(lines) {
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var pathItem = lines[lineIndex].path;
            pathItem.strokeDashes = [];
            pathItem.strokeDashOffset = 0;
        }
    }

    // ダイアログを表示してオプションを受け取る / Show the dialog and return selected options
    function showOptionDialog(classified, gridBounds, originalLineStates, strokeUnitInfo) {
        var optionDialog = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        optionDialog.alignChildren = "left";
        optionDialog.margins = 16;

        // メインオプション2カラム / Main options in two columns
        var mainOptionsGroup = optionDialog.add("group");
        mainOptionsGroup.orientation = "row";
        mainOptionsGroup.alignChildren = ["fill", "top"];

        var leftColumnGroup = mainOptionsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "fill";

        var rightColumnGroup = mainOptionsGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "fill";

        // 配置ラジオボタン / Distribution options
        var distributionPanel = leftColumnGroup.add("panel", undefined, getLabel("panelDistribution"));
        distributionPanel.orientation = "column";
        distributionPanel.alignChildren = "left";
        distributionPanel.margins = [10, 20, 10, 10];

        var distributionNoneRadio = distributionPanel.add("radiobutton", undefined, getLabel("optionDistributionNone"));
        var distributionEvenRadio = distributionPanel.add("radiobutton", undefined, getLabel("optionEven"));
        var distributionEvenMergedCellRadio = distributionPanel.add("radiobutton", undefined, getLabel("optionEvenMergedCell"));

        // デフォルトは均等＋結合セル対応 / Default is Evenly + merged cells
        distributionEvenMergedCellRadio.value = true;

        // 線パネル / Line panel
        var linePanel = rightColumnGroup.add("panel", undefined, getLabel("panelLine"));
        linePanel.orientation = "column";
        linePanel.alignChildren = "left";
        linePanel.margins = [10, 20, 10, 10];

        var projectingCapCheckbox = linePanel.add("checkbox", undefined, getLabel("optionProjectingCap"));
        projectingCapCheckbox.value = true;

        var convertDashedToSolidCheckbox = linePanel.add("checkbox", undefined, getLabel("optionConvertDashedToSolid"));
        convertDashedToSolidCheckbox.value = false;

        // 線幅パネル / Stroke width panel
        var strokeWidthPanel = linePanel.add("panel", undefined, getLabel("panelStrokeWidth"));
        strokeWidthPanel.orientation = "column";
        strokeWidthPanel.alignChildren = "left";
        strokeWidthPanel.margins = [10, 20, 10, 10];

        var strokeWidthMaxRadio = strokeWidthPanel.add("radiobutton", undefined, getLabel("optionStrokeWidthMax"));
        var strokeWidthMinRadio = strokeWidthPanel.add("radiobutton", undefined, getLabel("optionStrokeWidthMin"));
        var strokeWidthAverageRadio = strokeWidthPanel.add("radiobutton", undefined, getLabel("optionStrokeWidthAverage"));
        var strokeWidthSpecifiedRadio = strokeWidthPanel.add("radiobutton", undefined, getLabel("optionStrokeWidthSpecified"));

        var strokeWidthSpecifiedInputGroup = strokeWidthPanel.add("group");
        strokeWidthSpecifiedInputGroup.orientation = "row";
        strokeWidthSpecifiedInputGroup.alignChildren = ["left", "center"];
        var strokeWidthInput = strokeWidthSpecifiedInputGroup.add("edittext", undefined, "0.25");
        strokeWidthInput.characters = 4;
        strokeWidthSpecifiedInputGroup.add("statictext", undefined, strokeUnitInfo.label);

        // デフォルトは平均
        strokeWidthAverageRadio.value = true;
        strokeWidthInput.enabled = false;

        // 後処理パネル / Post-processing panel
        var postProcessingPanel = leftColumnGroup.add("panel", undefined, getLabel("panelPostProcessing"));
        postProcessingPanel.orientation = "column";
        postProcessingPanel.alignChildren = "left";
        postProcessingPanel.margins = [10, 20, 10, 10];

        var frameToRectangleCheckbox = postProcessingPanel.add("checkbox", undefined, getLabel("optionFrameToRect"));
        frameToRectangleCheckbox.value = false;

        var groupingCheckbox = postProcessingPanel.add("checkbox", undefined, getLabel("optionGroup"));
        groupingCheckbox.value = true;

        var buttonAreaGroup = optionDialog.add("group");
        buttonAreaGroup.orientation = "row";
        buttonAreaGroup.alignment = "fill";
        buttonAreaGroup.alignChildren = ["fill", "center"];

        var previewButtonGroup = buttonAreaGroup.add("group");
        previewButtonGroup.orientation = "row";
        previewButtonGroup.alignChildren = ["left", "center"];
        var previewCheckbox = previewButtonGroup.add("checkbox", undefined, getLabel("optionPreview"));

        var buttonSpacerGroup = buttonAreaGroup.add("group");
        buttonSpacerGroup.alignment = ["fill", "center"];

        var actionButtonGroup = buttonAreaGroup.add("group");
        actionButtonGroup.orientation = "row";
        actionButtonGroup.alignChildren = ["right", "center"];
        actionButtonGroup.add("button", undefined, getLabel("buttonCancel"), { name: "cancel" });
        actionButtonGroup.add("button", undefined, getLabel("buttonOK"), { name: "ok" });

        var ui = {
            previewCheckbox: previewCheckbox,
            distributionNoneRadio: distributionNoneRadio,
            distributionEvenRadio: distributionEvenRadio,
            distributionEvenMergedCellRadio: distributionEvenMergedCellRadio,
            projectingCapCheckbox: projectingCapCheckbox,
            convertDashedToSolidCheckbox: convertDashedToSolidCheckbox,
            strokeWidthMaxRadio: strokeWidthMaxRadio,
            strokeWidthMinRadio: strokeWidthMinRadio,
            strokeWidthAverageRadio: strokeWidthAverageRadio,
            strokeWidthSpecifiedRadio: strokeWidthSpecifiedRadio,
            strokeWidthInput: strokeWidthInput,
            strokeUnitInfo: strokeUnitInfo,
            frameToRectangleCheckbox: frameToRectangleCheckbox,
            groupingCheckbox: groupingCheckbox
        };

        previewCheckbox.onClick = function () {
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        distributionNoneRadio.onClick = function () {
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        distributionEvenRadio.onClick = function () {
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        distributionEvenMergedCellRadio.onClick = function () {
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        projectingCapCheckbox.onClick = function () {
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        convertDashedToSolidCheckbox.onClick = function () {
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        strokeWidthMaxRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(ui);
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        strokeWidthMinRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(ui);
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        strokeWidthAverageRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(ui);
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        strokeWidthSpecifiedRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(ui);
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        strokeWidthInput.onChanging = function () {
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        };
        changeValueByArrowKey(strokeWidthInput, false, function () {
            if (!strokeWidthSpecifiedRadio.value) strokeWidthSpecifiedRadio.value = true;
            syncSpecifiedStrokeWidthInput(ui);
            updatePreviewFromDialogState(ui, classified, gridBounds, originalLineStates);
        });

        var dialogResult = optionDialog.show();
        restoreLineStates(originalLineStates);
        if (dialogResult !== 1) return null;
        return readOptionDialogState(ui);
    }
})();