#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要
選択した水平線・垂直線を解析し、長方形グリッドとして再構成する Illustrator スクリプトです。
不揃いな罫線や、結合セルを含むレイアウトを整理し、整った格子構造に変換します。
主な機能：
  配置：均等配置なし／均等（強制）／均等＋結合セル対応 を選択可能
  対象：縦罫・横罫ごとに均等化対象を制御
  線：突出線端、破線→実線、線幅（最大・最小・平均・指定）を設定
  後処理：外枠の長方形化、グループ化
  プレビュー：ダイアログを閉じずに結果を確認
Overview
This Illustrator script analyzes selected horizontal and vertical lines and reconstructs them into a rectangular grid.
It is designed to normalize uneven rule lines and layouts that include merged-cell-like structures.
Main features:
  Distribution: Choose none, evenly (force), or evenly with merged cell support
  Targets: Control which lines (horizontal/vertical) are equalized
  Line: Set projecting caps, convert dashed lines to solid, and define stroke width (max/min/average/specified)
  Post-processing: Convert outer frame to a rectangle and optionally group items
  Preview: See results without closing the dialog
*/

(function () {
    // =========================================
    // バージョンとローカライズ / Version and localization
    // =========================================

    var SCRIPT_VERSION = "v1.2.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "グリッド再構築",
            en: "Rebuild Grid"
        },
        optionEven: {
            ja: "均等に（強制）",
            en: "Evenly (force)"
        },
        panelPreprocessing: {
            ja: "前処理",
            en: "Pre-processing"
        },
        optionSplitFrameToFourSides: {
            ja: "外枠を四辺に分割",
            en: "Split outer frame into four sides"
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
        panelEqualize: {
            ja: "均等配置の対象",
            en: "Equalize Targets"
        },
        optionEqualizeVertical: {
            ja: "縦罫",
            en: "Vertical lines"
        },
        optionEqualizeHorizontal: {
            ja: "横罫",
            en: "Horizontal lines"
        },
        optionLockFirstColumn: {
            ja: "1列目を固定",
            en: "Lock first column"
        },
        optionLockFirstRow: {
            ja: "1行目を固定",
            en: "Lock first row"
        },
        panelLine: {
            ja: "線（後処理）",
            en: "Line (post-process)"
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
        optionCenterPointTextVertically: {
            ja: "ポイント文字をセル内で上下中央",
            en: "Center point text vertically in cells"
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
        if (LABELS[key] && LABELS[key][currentLanguage]) return LABELS[key][currentLanguage];
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

    // グループ等を再帰展開して PathItem の平坦リストにする / Flatten selection (descending into groups) to PathItems
    var flatPathItems = flattenSelectionToPathItems(selectedItems);

    // 塗りのあるパスは対象外 / Exclude paths that have a fill
    flatPathItems = excludeFilledPaths(flatPathItems);

    // 軸並行の長方形は4本の罫線に分解 / Decompose axis-aligned rectangles into 4 lines
    var expansionResult = expandRectanglesInSelection(flatPathItems);
    var workingItems = expansionResult.items;
    var rectangleExpansions = expansionResult.expansions;

    var classified = classifySelectedStraightLines(workingItems);
    if (classified.horizontalLines.length < 2 || classified.verticalLines.length < 2) {
        restoreRectangleExpansions(rectangleExpansions);
        alert(getLabel("errorNotEnoughLines"));
        return;
    }

    var gridBounds = getGridBounds(classified.horizontalLines, classified.verticalLines);
    var originalLineStates = captureLineStates(classified.horizontalLines, classified.verticalLines);
    var strokeUnitInfo = getUnitInfoFromPreference("strokeUnits");

    var hasExpandedRectangles = rectangleExpansions.length > 0;
    var dialogOptions = showOptionDialog(classified, gridBounds, originalLineStates, strokeUnitInfo, hasExpandedRectangles);
    if (!dialogOptions) {
        restoreRectangleExpansions(rectangleExpansions);
        return;
    }

    // 前処理：外枠を四辺に分割しない場合は分解を取り消して終了 / If split is disabled, undo expansion and exit
    if (!dialogOptions.splitOuterFrame && hasExpandedRectangles) {
        restoreRectangleExpansions(rectangleExpansions);
        return;
    }

    restoreLineStates(originalLineStates);
    applyRepresentativeStrokeWidth(classified.horizontalLines, classified.verticalLines, dialogOptions);
    if (dialogOptions.convertDashedToSolid) convertDashedLinesToSolid(classified.horizontalLines, classified.verticalLines);
    alignLinesToGridBounds(classified.horizontalLines, classified.verticalLines, gridBounds, dialogOptions);

    // ポイント文字をセル内で上下中央に配置 / Center point text vertically within cells
    if (dialogOptions.centerPointTextVertically) {
        centerPointTextVerticallyInCells(classified.horizontalLines, gridBounds);
    }

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

    // 選択を再帰展開して PathItem 平坦リストを返す / Flatten a selection into a flat PathItem list (descend into groups and compound paths)
    function flattenSelectionToPathItems(items) {
        var pathItems = [];
        collectPathItemsRecursively(items, pathItems);
        return pathItems;
    }

    function collectPathItemsRecursively(items, pathItems) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var item = items[itemIndex];
            if (!item) continue;
            if (item.typename === "PathItem") {
                pathItems.push(item);
            } else if (item.typename === "GroupItem") {
                collectPathItemsRecursively(item.pageItems, pathItems);
            } else if (item.typename === "CompoundPathItem") {
                collectPathItemsRecursively(item.pathItems, pathItems);
            }
        }
    }

    // 塗りのあるパスを対象から除外 / Exclude filled paths from the working set
    function excludeFilledPaths(pathItems) {
        var result = [];
        for (var itemIndex = 0; itemIndex < pathItems.length; itemIndex++) {
            if (!pathItems[itemIndex].filled) result.push(pathItems[itemIndex]);
        }
        return result;
    }

    // 軸並行の長方形か判定（閉じた4点・各辺が水平または垂直）/ Check if path is an axis-aligned rectangle
    function isAxisAlignedRectangle(pathItem) {
        if (!pathItem || pathItem.typename !== "PathItem") return false;
        if (!pathItem.closed) return false;
        if (pathItem.pathPoints.length !== 4) return false;
        var rectangleTolerance = 0.01;
        for (var pointIndex = 0; pointIndex < 4; pointIndex++) {
            var currentAnchor = pathItem.pathPoints[pointIndex].anchor;
            var nextAnchor = pathItem.pathPoints[(pointIndex + 1) % 4].anchor;
            var horizontalDelta = Math.abs(currentAnchor[0] - nextAnchor[0]);
            var verticalDelta = Math.abs(currentAnchor[1] - nextAnchor[1]);
            if (horizontalDelta > rectangleTolerance && verticalDelta > rectangleTolerance) return false;
            if (horizontalDelta < rectangleTolerance && verticalDelta < rectangleTolerance) return false;
        }
        return true;
    }

    // 長方形パスから外接座標を取得 / Get bounding coordinates from a rectangle path
    function getRectangleBoundsFromPath(pathItem) {
        var rectangleMinX = pathItem.pathPoints[0].anchor[0], rectangleMaxX = rectangleMinX;
        var rectangleMinY = pathItem.pathPoints[0].anchor[1], rectangleMaxY = rectangleMinY;
        for (var pointIndex = 1; pointIndex < pathItem.pathPoints.length; pointIndex++) {
            var anchor = pathItem.pathPoints[pointIndex].anchor;
            if (anchor[0] < rectangleMinX) rectangleMinX = anchor[0];
            if (anchor[0] > rectangleMaxX) rectangleMaxX = anchor[0];
            if (anchor[1] < rectangleMinY) rectangleMinY = anchor[1];
            if (anchor[1] > rectangleMaxY) rectangleMaxY = anchor[1];
        }
        return { minX: rectangleMinX, maxX: rectangleMaxX, minY: rectangleMinY, maxY: rectangleMaxY };
    }

    // パスの塗り・線属性を記録 / Capture fill and stroke appearance from a path
    function capturePathAppearance(pathItem) {
        var appearance = {
            filled: pathItem.filled,
            stroked: pathItem.stroked
        };
        if (pathItem.filled) appearance.fillColor = pathItem.fillColor;
        if (pathItem.stroked) {
            appearance.strokeColor = pathItem.strokeColor;
            appearance.strokeWidth = pathItem.strokeWidth;
            appearance.strokeCap = pathItem.strokeCap;
            appearance.strokeJoin = pathItem.strokeJoin;
            appearance.strokeMiterLimit = pathItem.strokeMiterLimit;
            appearance.strokeDashes = pathItem.strokeDashes;
            appearance.strokeDashOffset = pathItem.strokeDashOffset;
        }
        return appearance;
    }

    // 記録した塗り・線属性を適用 / Apply captured fill and stroke appearance
    function applyPathAppearance(pathItem, appearance) {
        pathItem.filled = appearance.filled;
        if (appearance.filled) pathItem.fillColor = appearance.fillColor;
        pathItem.stroked = appearance.stroked;
        if (appearance.stroked) {
            pathItem.strokeColor = appearance.strokeColor;
            pathItem.strokeWidth = appearance.strokeWidth;
            pathItem.strokeCap = appearance.strokeCap;
            pathItem.strokeJoin = appearance.strokeJoin;
            pathItem.strokeMiterLimit = appearance.strokeMiterLimit;
            pathItem.strokeDashes = appearance.strokeDashes;
            pathItem.strokeDashOffset = appearance.strokeDashOffset;
        }
    }

    // 2点の直線パスを生成（線属性を引き継ぐ）/ Create a 2-point line path inheriting stroke appearance
    function createLineFromAppearance(parentContainer, firstPoint, secondPoint, appearance) {
        var newLine = parentContainer.pathItems.add();
        newLine.setEntirePath([firstPoint, secondPoint]);
        newLine.closed = false;
        applyPathAppearance(newLine, appearance);
        newLine.filled = false;
        return newLine;
    }

    // 選択内の長方形を4本の罫線に分解 / Expand axis-aligned rectangles in selection into 4 lines each
    function expandRectanglesInSelection(selectedItems) {
        var sourceItems = [];
        for (var sourceIndex = 0; sourceIndex < selectedItems.length; sourceIndex++) {
            sourceItems.push(selectedItems[sourceIndex]);
        }
        var expandedItems = [];
        var rectangleExpansions = [];
        for (var itemIndex = 0; itemIndex < sourceItems.length; itemIndex++) {
            var sourceItem = sourceItems[itemIndex];
            if (isAxisAlignedRectangle(sourceItem)) {
                var bounds = getRectangleBoundsFromPath(sourceItem);
                var appearance = capturePathAppearance(sourceItem);
                var parentContainer = sourceItem.parent;
                var topLeft = [bounds.minX, bounds.maxY];
                var topRight = [bounds.maxX, bounds.maxY];
                var bottomLeft = [bounds.minX, bounds.minY];
                var bottomRight = [bounds.maxX, bounds.minY];
                var generatedLines = [
                    createLineFromAppearance(parentContainer, topLeft, topRight, appearance),
                    createLineFromAppearance(parentContainer, bottomLeft, bottomRight, appearance),
                    createLineFromAppearance(parentContainer, bottomLeft, topLeft, appearance),
                    createLineFromAppearance(parentContainer, bottomRight, topRight, appearance)
                ];
                for (var generatedIndex = 0; generatedIndex < generatedLines.length; generatedIndex++) {
                    expandedItems.push(generatedLines[generatedIndex]);
                }
                rectangleExpansions.push({
                    bounds: bounds,
                    appearance: appearance,
                    parent: parentContainer,
                    generatedLines: generatedLines
                });
                try { sourceItem.remove(); } catch (rectangleRemoveError) { }
            } else {
                expandedItems.push(sourceItem);
            }
        }
        return { items: expandedItems, expansions: rectangleExpansions };
    }

    // 長方形分解を取り消す（キャンセル時用）/ Undo rectangle expansions (used on cancel)
    function restoreRectangleExpansions(rectangleExpansions) {
        for (var expansionIndex = 0; expansionIndex < rectangleExpansions.length; expansionIndex++) {
            var expansion = rectangleExpansions[expansionIndex];
            for (var generatedLineIndex = 0; generatedLineIndex < expansion.generatedLines.length; generatedLineIndex++) {
                try { expansion.generatedLines[generatedLineIndex].remove(); } catch (lineRemoveError) { }
            }
            var bounds = expansion.bounds;
            var restoredRectangle = expansion.parent.pathItems.rectangle(
                bounds.maxY, bounds.minX,
                bounds.maxX - bounds.minX, bounds.maxY - bounds.minY
            );
            applyPathAppearance(restoredRectangle, expansion.appearance);
        }
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

    // 格子の外周座標（左端・右端・上端・下端）を計算
    // 上下左右に貫通する罫線（最長線の90%以上の長さ）が複数あればそれを基準とし、
    // それより外側にはみ出した線は計算対象から除外する
    // / Calculate grid outer bounds. If there are spanning lines (>=90% of the longest in that orientation),
    // use those as the basis and ignore lines that protrude beyond them.
    function getGridBounds(horizontalLines, verticalLines) {
        var SPAN_RATIO_THRESHOLD = 0.9;

        var maxHorizontalSpan = 0;
        for (var hSpanIndex = 0; hSpanIndex < horizontalLines.length; hSpanIndex++) {
            var hSpanLength = horizontalLines[hSpanIndex].maxX - horizontalLines[hSpanIndex].minX;
            if (hSpanLength > maxHorizontalSpan) maxHorizontalSpan = hSpanLength;
        }
        var maxVerticalSpan = 0;
        for (var vSpanIndex = 0; vSpanIndex < verticalLines.length; vSpanIndex++) {
            var vSpanLength = verticalLines[vSpanIndex].maxY - verticalLines[vSpanIndex].minY;
            if (vSpanLength > maxVerticalSpan) maxVerticalSpan = vSpanLength;
        }

        var spanningHorizontals = [];
        var horizontalThreshold = maxHorizontalSpan * SPAN_RATIO_THRESHOLD;
        for (var hPickIndex = 0; hPickIndex < horizontalLines.length; hPickIndex++) {
            if ((horizontalLines[hPickIndex].maxX - horizontalLines[hPickIndex].minX) >= horizontalThreshold) {
                spanningHorizontals.push(horizontalLines[hPickIndex]);
            }
        }
        var spanningVerticals = [];
        var verticalThreshold = maxVerticalSpan * SPAN_RATIO_THRESHOLD;
        for (var vPickIndex = 0; vPickIndex < verticalLines.length; vPickIndex++) {
            if ((verticalLines[vPickIndex].maxY - verticalLines[vPickIndex].minY) >= verticalThreshold) {
                spanningVerticals.push(verticalLines[vPickIndex]);
            }
        }

        // 貫通線が2本以上あれば外枠候補として採用、なければ全線をフォールバック
        // / Use spanning lines as outer-frame candidates when at least 2 exist; otherwise fall back to all lines
        var verticalSource = spanningVerticals.length >= 2 ? spanningVerticals : verticalLines;
        var horizontalSource = spanningHorizontals.length >= 2 ? spanningHorizontals : horizontalLines;

        var gridMinX = verticalSource[0].x, gridMaxX = verticalSource[0].x;
        for (var verticalIndex = 1; verticalIndex < verticalSource.length; verticalIndex++) {
            if (verticalSource[verticalIndex].x < gridMinX) gridMinX = verticalSource[verticalIndex].x;
            if (verticalSource[verticalIndex].x > gridMaxX) gridMaxX = verticalSource[verticalIndex].x;
        }
        var gridMinY = horizontalSource[0].y, gridMaxY = horizontalSource[0].y;
        for (var horizontalIndex = 1; horizontalIndex < horizontalSource.length; horizontalIndex++) {
            if (horizontalSource[horizontalIndex].y < gridMinY) gridMinY = horizontalSource[horizontalIndex].y;
            if (horizontalSource[horizontalIndex].y > gridMaxY) gridMaxY = horizontalSource[horizontalIndex].y;
        }
        return { minX: gridMinX, maxX: gridMaxX, minY: gridMinY, maxY: gridMaxY };
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
    function updatePreviewFromDialogState(dialogUi, classified, gridBounds, originalLineStates) {
        restoreLineStates(originalLineStates);
        if (!dialogUi.previewCheckbox.value) {
            app.redraw();
            return;
        }
        // 外枠を四辺に分割が OFF の場合、本処理（罫線整列）を行わず元の状態のまま表示 / If split is OFF, skip the main alignment and show original state
        if (!dialogUi.splitOuterFrameCheckbox.value) {
            app.redraw();
            return;
        }
        var previewOptions = readOptionDialogState(dialogUi);
        // プレビュー時は frameToRect, group を適用しない
        previewOptions.frameToRect = false;
        previewOptions.group = false;
        applyRepresentativeStrokeWidth(classified.horizontalLines, classified.verticalLines, previewOptions);
        if (previewOptions.convertDashedToSolid) convertDashedLinesToSolid(classified.horizontalLines, classified.verticalLines);
        alignLinesToGridBounds(classified.horizontalLines, classified.verticalLines, gridBounds, previewOptions);
        app.redraw();
    }

    // ダイアログの現在値を読む / Read current dialog state
    function readOptionDialogState(dialogUi) {
        var distributionMode = dialogUi.distributionNoneRadio.value ? "none"
            : dialogUi.distributionEvenMergedCellRadio.value ? "evenMergedCell"
                : "even";
        var isEvenMode = distributionMode !== "none";
        var evenHorizontal = isEvenMode && dialogUi.equalizeHorizontalCheckbox.value;
        var evenVertical = isEvenMode && dialogUi.equalizeVerticalCheckbox.value;
        return {
            distributionMode: distributionMode,
            even: isEvenMode,
            evenHorizontal: evenHorizontal,
            evenVertical: evenVertical,
            projectingCap: dialogUi.projectingCapCheckbox.value,
            convertDashedToSolid: dialogUi.convertDashedToSolidCheckbox.value,
            strokeWidthMode: dialogUi.strokeWidthMaxRadio.value ? "max"
                : dialogUi.strokeWidthMinRadio.value ? "min"
                    : dialogUi.strokeWidthSpecifiedRadio.value ? "specified"
                        : "average",
            specifiedStrokeWidthPt: dialogUi.strokeWidthSpecifiedRadio && dialogUi.strokeWidthSpecifiedRadio.value
                ? unitValueToPoints(readNumericText(dialogUi.strokeWidthInput.text, 0), dialogUi.strokeUnitInfo)
                : 0,
            lockFirstColumn: dialogUi.lockFirstColumnCheckbox.value,
            lockFirstRow: dialogUi.lockFirstRowCheckbox.value,
            frameToRect: dialogUi.frameToRectangleCheckbox.value,
            centerPointTextVertically: dialogUi.centerPointTextVerticallyCheckbox.value,
            group: dialogUi.groupingCheckbox.value,
            splitOuterFrame: dialogUi.splitOuterFrameCheckbox.value
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
    function syncSpecifiedStrokeWidthInput(dialogUi) {
        dialogUi.strokeWidthInput.enabled = dialogUi.strokeWidthSpecifiedRadio.value;
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
        if (dialogOptions.evenHorizontal && horizontalLines.length >= 2) {
            horizontalLines.sort(function (lineA, lineB) { return lineA.y - lineB.y; });
            if (dialogOptions.lockFirstRow && horizontalLines.length >= 3) {
                // 1行目（最上）・2行目を固定し、2行目と最下を基準に残りを均等配置 / Lock 1st & 2nd from top; distribute the rest between 2nd-from-top and bottom
                var topY = horizontalLines[horizontalLines.length - 1].y;
                var lockedAnchorY = horizontalLines[horizontalLines.length - 2].y;
                var bottommostY = horizontalLines[0].y;
                var lockedHorizontalStep = (lockedAnchorY - bottommostY) / (horizontalLines.length - 2);
                for (var lockedHorizontalIndex = 0; lockedHorizontalIndex < horizontalLines.length - 1; lockedHorizontalIndex++) {
                    horizontalLineYPositions.push(bottommostY + lockedHorizontalStep * lockedHorizontalIndex);
                }
                horizontalLineYPositions.push(topY);
            } else {
                var horizontalStep = (gridBounds.maxY - gridBounds.minY) / (horizontalLines.length - 1);
                for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
                    horizontalLineYPositions.push(gridBounds.minY + horizontalStep * horizontalIndex);
                }
            }
        } else {
            for (var originalHorizontalIndex = 0; originalHorizontalIndex < horizontalLines.length; originalHorizontalIndex++) {
                horizontalLineYPositions.push(horizontalLines[originalHorizontalIndex].y);
            }
        }

        // 垂直線の X 位置を決定（均等配置 or 元の X）/ Determine X positions for vertical lines
        var verticalLineXPositions = [];
        if (dialogOptions.evenVertical && verticalLines.length >= 2) {
            verticalLines.sort(function (lineA, lineB) { return lineA.x - lineB.x; });
            if (dialogOptions.lockFirstColumn && verticalLines.length >= 3) {
                // 1本目・2本目を固定し、2本目と最右を基準に残りを均等配置 / Lock 1st & 2nd; distribute the rest between 2nd and rightmost
                var lockedAnchorX = verticalLines[1].x;
                var rightmostVerticalX = verticalLines[verticalLines.length - 1].x;
                var lockedVerticalStep = (rightmostVerticalX - lockedAnchorX) / (verticalLines.length - 2);
                verticalLineXPositions.push(verticalLines[0].x);
                for (var lockedVerticalIndex = 1; lockedVerticalIndex < verticalLines.length; lockedVerticalIndex++) {
                    verticalLineXPositions.push(lockedAnchorX + lockedVerticalStep * (lockedVerticalIndex - 1));
                }
            } else {
                var verticalStep = (gridBounds.maxX - gridBounds.minX) / (verticalLines.length - 1);
                for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
                    verticalLineXPositions.push(gridBounds.minX + verticalStep * verticalIndex);
                }
            }
        } else {
            for (var originalVerticalIndex = 0; originalVerticalIndex < verticalLines.length; originalVerticalIndex++) {
                verticalLineXPositions.push(verticalLines[originalVerticalIndex].x);
            }
        }

        // Illustrator は Y 上方向が正 / Illustrator uses positive Y upward
        var topCoordinateY = Math.max(gridBounds.minY, gridBounds.maxY);
        var bottomCoordinateY = Math.min(gridBounds.minY, gridBounds.maxY);

        for (var horizontalPathIndex = 0; horizontalPathIndex < horizontalLines.length; horizontalPathIndex++) {
            var horizontalPath = horizontalLines[horizontalPathIndex].path;
            if (dialogOptions.projectingCap) horizontalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
            setLineEndpoints(horizontalPath, [gridBounds.minX, horizontalLineYPositions[horizontalPathIndex]], [gridBounds.maxX, horizontalLineYPositions[horizontalPathIndex]]);
        }
        for (var verticalPathIndex = 0; verticalPathIndex < verticalLines.length; verticalPathIndex++) {
            var verticalPath = verticalLines[verticalPathIndex].path;
            if (dialogOptions.projectingCap) verticalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
            setLineEndpoints(verticalPath, [verticalLineXPositions[verticalPathIndex], topCoordinateY], [verticalLineXPositions[verticalPathIndex], bottomCoordinateY]);
        }
    }

    // 結合セル対応モード：同じ座標の線を 1 行/1 列としてまとめ、両端は直近の交差点にスナップ / Merged-cell mode: lines on the same coordinate share one row/column; endpoints snap to the nearest intersection
    function alignLinesWithMergedCellSupport(horizontalLines, verticalLines, gridBounds, dialogOptions) {
        var horizontalGroups = groupLinesByCoordinate(horizontalLines, "y", COORDINATE_TOLERANCE);
        var verticalGroups = groupLinesByCoordinate(verticalLines, "x", COORDINATE_TOLERANCE);

        // 各グループの新しい Y 位置（行）/ New Y position for each horizontal group (row)
        var horizontalGroupYPositions = [];
        if (dialogOptions.evenHorizontal && horizontalGroups.length >= 2) {
            if (dialogOptions.lockFirstRow && horizontalGroups.length >= 3) {
                // 1行目（最上）・2行目を固定し、2行目と最下を基準に残りを均等配置 / Lock 1st & 2nd rows from top; distribute the rest between 2nd-from-top and bottom
                var topRowY = horizontalGroups[horizontalGroups.length - 1].coord;
                var lockedAnchorRowY = horizontalGroups[horizontalGroups.length - 2].coord;
                var bottommostRowY = horizontalGroups[0].coord;
                var lockedRowStep = (lockedAnchorRowY - bottommostRowY) / (horizontalGroups.length - 2);
                for (var lockedRowIndex = 0; lockedRowIndex < horizontalGroups.length - 1; lockedRowIndex++) {
                    horizontalGroupYPositions.push(bottommostRowY + lockedRowStep * lockedRowIndex);
                }
                horizontalGroupYPositions.push(topRowY);
            } else {
                var horizontalStep = (gridBounds.maxY - gridBounds.minY) / (horizontalGroups.length - 1);
                for (var horizontalGroupIndex = 0; horizontalGroupIndex < horizontalGroups.length; horizontalGroupIndex++) {
                    horizontalGroupYPositions.push(gridBounds.minY + horizontalStep * horizontalGroupIndex);
                }
            }
        } else {
            for (var originalHorizontalGroupIndex = 0; originalHorizontalGroupIndex < horizontalGroups.length; originalHorizontalGroupIndex++) {
                horizontalGroupYPositions.push(horizontalGroups[originalHorizontalGroupIndex].coord);
            }
        }

        // 各グループの新しい X 位置（列）/ New X position for each vertical group (column)
        var verticalGroupXPositions = [];
        if (dialogOptions.evenVertical && verticalGroups.length >= 2) {
            if (dialogOptions.lockFirstColumn && verticalGroups.length >= 3) {
                // 1列目・2列目を固定し、2列目と最右を基準に残りを均等配置 / Lock 1st & 2nd columns; distribute the rest between 2nd and rightmost
                var lockedAnchorColumnX = verticalGroups[1].coord;
                var rightmostColumnX = verticalGroups[verticalGroups.length - 1].coord;
                var lockedColumnStep = (rightmostColumnX - lockedAnchorColumnX) / (verticalGroups.length - 2);
                verticalGroupXPositions.push(verticalGroups[0].coord);
                for (var lockedColumnIndex = 1; lockedColumnIndex < verticalGroups.length; lockedColumnIndex++) {
                    verticalGroupXPositions.push(lockedAnchorColumnX + lockedColumnStep * (lockedColumnIndex - 1));
                }
            } else {
                var verticalStep = (gridBounds.maxX - gridBounds.minX) / (verticalGroups.length - 1);
                for (var verticalGroupIndex = 0; verticalGroupIndex < verticalGroups.length; verticalGroupIndex++) {
                    verticalGroupXPositions.push(gridBounds.minX + verticalStep * verticalGroupIndex);
                }
            }
        } else {
            for (var originalVerticalGroupIndex = 0; originalVerticalGroupIndex < verticalGroups.length; originalVerticalGroupIndex++) {
                verticalGroupXPositions.push(verticalGroups[originalVerticalGroupIndex].coord);
            }
        }

        // 水平線：同一行のすべてに同じ Y、両端は最寄り列にスナップ / Apply same Y to all lines in a row; snap endpoints to nearest column
        for (var horizontalGroupIndex = 0; horizontalGroupIndex < horizontalGroups.length; horizontalGroupIndex++) {
            var rowCoordinateY = horizontalGroupYPositions[horizontalGroupIndex];
            var rowLines = horizontalGroups[horizontalGroupIndex].lines;
            for (var rowLineIndex = 0; rowLineIndex < rowLines.length; rowLineIndex++) {
                var horizontalLineData = rowLines[rowLineIndex];
                var horizontalPath = horizontalLineData.path;
                if (dialogOptions.projectingCap) horizontalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
                var leftColumnIndex = findClosestIndexByProperty(horizontalLineData.minX, verticalGroups, "coord");
                var rightColumnIndex = findClosestIndexByProperty(horizontalLineData.maxX, verticalGroups, "coord");
                setLineEndpoints(horizontalPath, [verticalGroupXPositions[leftColumnIndex], rowCoordinateY], [verticalGroupXPositions[rightColumnIndex], rowCoordinateY]);
            }
        }

        // 垂直線：同一列のすべてに同じ X、両端は最寄り行にスナップ / Apply same X to all lines in a column; snap endpoints to nearest row
        for (var verticalGroupIndex = 0; verticalGroupIndex < verticalGroups.length; verticalGroupIndex++) {
            var columnCoordinateX = verticalGroupXPositions[verticalGroupIndex];
            var columnLines = verticalGroups[verticalGroupIndex].lines;
            for (var columnLineIndex = 0; columnLineIndex < columnLines.length; columnLineIndex++) {
                var verticalLineData = columnLines[columnLineIndex];
                var verticalPath = verticalLineData.path;
                if (dialogOptions.projectingCap) verticalPath.strokeCap = StrokeCap.PROJECTINGENDCAP;
                var topRowIndex = findClosestIndexByProperty(verticalLineData.maxY, horizontalGroups, "coord");
                var bottomRowIndex = findClosestIndexByProperty(verticalLineData.minY, horizontalGroups, "coord");
                setLineEndpoints(verticalPath, [columnCoordinateX, horizontalGroupYPositions[topRowIndex]], [columnCoordinateX, horizontalGroupYPositions[bottomRowIndex]]);
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

    // ポイント文字をセル内で上下中央に移動 / Move point text frames to the vertical center of their cell row
    function centerPointTextVerticallyInCells(horizontalLines, gridBounds) {
        var rowYs = [];
        for (var lineIndex = 0; lineIndex < horizontalLines.length; lineIndex++) {
            rowYs.push(horizontalLines[lineIndex].y);
        }
        rowYs.sort(function (firstY, secondY) { return firstY - secondY; });
        if (rowYs.length < 2) return;

        var documentTextFrames = app.activeDocument.textFrames;
        for (var textIndex = 0; textIndex < documentTextFrames.length; textIndex++) {
            var textFrame = documentTextFrames[textIndex];
            if (textFrame.kind !== TextType.POINTTEXT) continue;
            var textBounds = textFrame.geometricBounds; // [left, top, right, bottom]（Y上方向が正）
            var textCenterX = (textBounds[0] + textBounds[2]) / 2;
            var textCenterY = (textBounds[1] + textBounds[3]) / 2;
            // グリッド外のテキストは対象外 / Skip text outside the grid bounds
            if (textCenterX < gridBounds.minX || textCenterX > gridBounds.maxX) continue;
            if (textCenterY < gridBounds.minY || textCenterY > gridBounds.maxY) continue;
            // テキストの中心が含まれる行を探す / Find the row whose Y range contains the text center
            var rowBottomY = null, rowTopY = null;
            for (var rowIndex = 0; rowIndex < rowYs.length - 1; rowIndex++) {
                if (textCenterY >= rowYs[rowIndex] && textCenterY <= rowYs[rowIndex + 1]) {
                    rowBottomY = rowYs[rowIndex];
                    rowTopY = rowYs[rowIndex + 1];
                    break;
                }
            }
            if (rowBottomY === null) continue;
            var rowCenterY = (rowBottomY + rowTopY) / 2;
            var deltaY = rowCenterY - textCenterY;
            if (Math.abs(deltaY) > 0.001) textFrame.translate(0, deltaY);
        }
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
    function showOptionDialog(classified, gridBounds, originalLineStates, strokeUnitInfo, hasExpandedRectangles) {
        function refreshPreview() {
            updatePreviewFromDialogState(dialogUi, classified, gridBounds, originalLineStates);
        }
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

        // 前処理パネル / Pre-processing panel
        var preprocessingPanel = leftColumnGroup.add("panel", undefined, getLabel("panelPreprocessing"));
        preprocessingPanel.orientation = "column";
        preprocessingPanel.alignChildren = "left";
        preprocessingPanel.margins = [10, 20, 10, 10];

        var splitOuterFrameCheckbox = preprocessingPanel.add("checkbox", undefined, getLabel("optionSplitFrameToFourSides"));
        splitOuterFrameCheckbox.value = !!hasExpandedRectangles;

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

        // 平均化パネル / Equalize panel
        var equalizePanel = leftColumnGroup.add("panel", undefined, getLabel("panelEqualize"));
        equalizePanel.orientation = "column";
        equalizePanel.alignChildren = "left";
        equalizePanel.margins = [10, 20, 10, 10];

        var equalizeVerticalRow = equalizePanel.add("group");
        equalizeVerticalRow.orientation = "row";
        equalizeVerticalRow.alignChildren = ["left", "center"];
        var equalizeVerticalCheckbox = equalizeVerticalRow.add("checkbox", undefined, getLabel("optionEqualizeVertical"));
        var lockFirstColumnCheckbox = equalizeVerticalRow.add("checkbox", undefined, getLabel("optionLockFirstColumn"));

        var equalizeHorizontalRow = equalizePanel.add("group");
        equalizeHorizontalRow.orientation = "row";
        equalizeHorizontalRow.alignChildren = ["left", "center"];
        var equalizeHorizontalCheckbox = equalizeHorizontalRow.add("checkbox", undefined, getLabel("optionEqualizeHorizontal"));
        var lockFirstRowCheckbox = equalizeHorizontalRow.add("checkbox", undefined, getLabel("optionLockFirstRow"));

        // デフォルトは両方 OFF / Default both to OFF
        equalizeVerticalCheckbox.value = false;
        equalizeHorizontalCheckbox.value = false;

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
        frameToRectangleCheckbox.value = !!hasExpandedRectangles;

        var centerPointTextVerticallyCheckbox = postProcessingPanel.add("checkbox", undefined, getLabel("optionCenterPointTextVertically"));
        centerPointTextVerticallyCheckbox.value = false;

        var groupingCheckbox = postProcessingPanel.add("checkbox", undefined, getLabel("optionGroup"));
        groupingCheckbox.value = false;

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

        var dialogUi = {
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
            centerPointTextVerticallyCheckbox: centerPointTextVerticallyCheckbox,
            groupingCheckbox: groupingCheckbox,
            splitOuterFrameCheckbox: splitOuterFrameCheckbox,
            equalizeVerticalCheckbox: equalizeVerticalCheckbox,
            equalizeHorizontalCheckbox: equalizeHorizontalCheckbox,
            lockFirstColumnCheckbox: lockFirstColumnCheckbox,
            lockFirstRowCheckbox: lockFirstRowCheckbox
        };

        // 配置モードに応じて平均化パネルの有効/無効を同期 / Sync equalize panel enabled state with distribution mode
        function syncEqualizePanelEnabled() {
            var enabled = !distributionNoneRadio.value;
            equalizeVerticalCheckbox.enabled = enabled;
            equalizeHorizontalCheckbox.enabled = enabled;
            lockFirstColumnCheckbox.enabled = enabled && equalizeVerticalCheckbox.value;
            lockFirstRowCheckbox.enabled = enabled && equalizeHorizontalCheckbox.value;
        }

        syncEqualizePanelEnabled();

        previewCheckbox.onClick = function () {
            refreshPreview();
        };
        distributionNoneRadio.onClick = function () {
            syncEqualizePanelEnabled();
            refreshPreview();
        };
        distributionEvenRadio.onClick = function () {
            syncEqualizePanelEnabled();
            refreshPreview();
        };
        distributionEvenMergedCellRadio.onClick = function () {
            syncEqualizePanelEnabled();
            refreshPreview();
        };
        equalizeVerticalCheckbox.onClick = function () {
            if (ScriptUI.environment.keyboardState.altKey) {
                equalizeHorizontalCheckbox.value = equalizeVerticalCheckbox.value;
            }
            syncEqualizePanelEnabled();
            refreshPreview();
        };
        equalizeHorizontalCheckbox.onClick = function () {
            if (ScriptUI.environment.keyboardState.altKey) {
                equalizeVerticalCheckbox.value = equalizeHorizontalCheckbox.value;
            }
            syncEqualizePanelEnabled();
            refreshPreview();
        };
        lockFirstColumnCheckbox.onClick = function () {
            refreshPreview();
        };
        lockFirstRowCheckbox.onClick = function () {
            refreshPreview();
        };
        projectingCapCheckbox.onClick = function () {
            refreshPreview();
        };
        convertDashedToSolidCheckbox.onClick = function () {
            refreshPreview();
        };
        strokeWidthMaxRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(dialogUi);
            refreshPreview();
        };
        strokeWidthMinRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(dialogUi);
            refreshPreview();
        };
        strokeWidthAverageRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(dialogUi);
            refreshPreview();
        };
        strokeWidthSpecifiedRadio.onClick = function () {
            syncSpecifiedStrokeWidthInput(dialogUi);
            refreshPreview();
        };
        strokeWidthInput.onChanging = function () {
            refreshPreview();
        };
        changeValueByArrowKey(strokeWidthInput, false, function () {
            if (!strokeWidthSpecifiedRadio.value) strokeWidthSpecifiedRadio.value = true;
            syncSpecifiedStrokeWidthInput(dialogUi);
            refreshPreview();
        });
        // 外枠分割トグル：分割済みの線に対して整列処理（プレビュー）を再実行 / Toggle: re-run alignment preview on the already-split lines
        splitOuterFrameCheckbox.onClick = function () {
            refreshPreview();
        };

        var dialogResult = optionDialog.show();
        restoreLineStates(originalLineStates);
        if (dialogResult !== 1) return null;
        return readOptionDialogState(dialogUi);
    }
})();