#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した水平線・垂直線を解析し、長方形グリッドとして再構成します。
不揃いな罫線や、結合セルを含むレイアウトを整理し、整った格子構造に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangularGridReverseTool.md

### Overview

Analyzes the selected horizontal and vertical lines and reconstructs them into a rectangular grid.
Uneven rules and layouts containing merged cells are tidied into a regular lattice.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangularGridReverseTool.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RectangularGridReverseTool";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangularGridReverseTool.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangularGridReverseTool.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;                /* ダイアログの余白 / Dialog margins */
    var PANEL_MARGINS = [10, 20, 10, 10];   /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グリッド再構築", en: "Rebuild Grid" }
        },
        panel: {
            preprocessing: { ja: "前処理", en: "Pre-processing" },
            distribution: { ja: "配置モード", en: "Distribution Mode" },
            equalize: { ja: "均等配置の対象", en: "Equalize Targets" },
            line: { ja: "線（後処理）", en: "Line (post-process)" },
            strokeWidth: { ja: "線幅", en: "Stroke Width" },
            postProcessing: { ja: "後処理", en: "Post-processing" }
        },
        radio: {
            distributionNone: { ja: "均等配置しない", en: "Do not distribute evenly" },
            distributionEven: { ja: "均等に（強制）", en: "Evenly (force)" },
            distributionEvenMergedCell: { ja: "均等＋結合セル対応", en: "Evenly + merged cells" },
            strokeWidthMax: { ja: "最大", en: "Maximum" },
            strokeWidthMin: { ja: "最小", en: "Minimum" },
            strokeWidthAverage: { ja: "平均", en: "Average" },
            strokeWidthSpecified: { ja: "指定", en: "Specified" }
        },
        checkbox: {
            splitFrameToFourSides: { ja: "外枠を四辺に分割", en: "Split outer frame into four sides" },
            equalizeVertical: { ja: "縦罫", en: "Vertical lines" },
            equalizeHorizontal: { ja: "横罫", en: "Horizontal lines" },
            lockFirstColumn: { ja: "1列目を固定", en: "Lock first column" },
            lockFirstRow: { ja: "1行目を固定", en: "Lock first row" },
            projectingCap: { ja: "突出線端にする", en: "Projecting cap" },
            convertDashedToSolid: { ja: "破線を実線にする", en: "Convert dashed lines to solid" },
            frameToRect: { ja: "外枠を長方形に変換", en: "Convert outer frame to rectangle" },
            centerPointTextVertically: { ja: "ポイント文字をセル内で上下中央", en: "Center point text vertically in cells" },
            grouping: { ja: "グループ化", en: "Group" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            splitFrameToFourSides: {
                ja: "外枠の長方形を、上下左右4本の線に分割します。",
                en: "Splits the outer rectangle into four separate lines."
            },
            distributionNone: { ja: "行・列の間隔は変えません。", en: "Leaves the row and column spacing alone." },
            distributionEven: { ja: "行・列の間隔を均等にそろえます。", en: "Evens out the row and column spacing." },
            distributionEvenMergedCell: {
                ja: "結合セルを考慮しながら、行・列の間隔を均等にそろえます。",
                en: "Evens out the spacing while respecting merged cells."
            },
            equalizeVertical: { ja: "列の幅をそろえます。", en: "Gives the columns the same width." },
            lockFirstColumn: {
                ja: "1列目の幅は変えずに、残りをそろえます。",
                en: "Keeps the first column as it is and evens out the rest."
            },
            equalizeHorizontal: { ja: "行の高さをそろえます。", en: "Gives the rows the same height." },
            lockFirstRow: {
                ja: "1行目の高さは変えずに、残りをそろえます。",
                en: "Keeps the first row as it is and evens out the rest."
            },
            projectingCap: {
                ja: "線の端を太さの半分だけ延ばして、角の隙間をなくします。",
                en: "Extends the line ends by half the weight so the corners close up."
            },
            convertDashedToSolid: { ja: "点線・破線を実線に変えます。", en: "Turns dashed lines into solid ones." },
            strokeWidthMax: { ja: "いちばん太い線に合わせます。", en: "Matches the thickest line." },
            strokeWidthMin: { ja: "いちばん細い線に合わせます。", en: "Matches the thinnest line." },
            strokeWidthAverage: { ja: "線幅の平均に合わせます。", en: "Matches the average line weight." },
            strokeWidthSpecified: { ja: "線幅を数値で指定します。", en: "Sets the line weight to a value you type." },
            frameToRect: {
                ja: "外周の4本の線を1つの長方形にまとめます。",
                en: "Merges the four outer lines back into a single rectangle."
            },
            centerPointTextVertically: {
                ja: "ポイント文字をセルの天地中央に置き直します。",
                en: "Re-centers point text vertically within its cell."
            },
            grouping: { ja: "できあがった表を1つのグループにまとめます。", en: "Groups the finished table together." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noSelection: { ja: "水平線と垂直線を選択してください。", en: "Select horizontal and vertical lines." },
            notEnoughLines: {
                ja: "格子を作るには、最低でも2本の水平線と2本の垂直線が必要です。",
                en: "To create a grid, select at least two horizontal lines and two vertical lines."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
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

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した罫線を解析し、ダイアログの設定でグリッドに組み直す
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var selectedItems = app.activeDocument.selection;
        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        // グループ等を再帰展開して PathItem の平坦リストにする / Flatten selection (descending into groups) to PathItems
        var flatPathItems = [];
        collectPathItemsRecursively(selectedItems, flatPathItems);

        // 塗りのあるパスは対象外 / Exclude paths that have a fill
        flatPathItems = excludeFilledPaths(flatPathItems);

        // 軸並行の長方形は4本の罫線に分解 / Decompose axis-aligned rectangles into 4 lines
        var expansionResult = expandRectanglesInSelection(flatPathItems);
        var rectangleExpansions = expansionResult.expansions;

        var classified = classifySelectedStraightLines(expansionResult.items);
        if (classified.horizontalLines.length < 2 || classified.verticalLines.length < 2) {
            restoreRectangleExpansions(rectangleExpansions);
            alert(getLabel("alert.notEnoughLines"));
            return;
        }

        var gridBounds = getGridBounds(classified.horizontalLines, classified.verticalLines);
        var originalLineStates = captureLineStates(classified.horizontalLines, classified.verticalLines);
        var strokeUnitInfo = getUnitInfo("strokeUnits");

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
        applyLineOptions(classified, gridBounds, dialogOptions);

        // ポイント文字をセル内で上下中央に配置 / Center point text vertically within cells
        if (dialogOptions.centerPointTextVertically) {
            centerPointTextVerticallyInCells(classified.horizontalLines, gridBounds);
        }

        // 外枠を長方形に変換 → 残存パス＋長方形をグループ化 / Convert outer frame to rectangle, then group remaining paths and rectangle
        var groupTargets = collectLinePaths(classified.horizontalLines.concat(classified.verticalLines));
        if (dialogOptions.frameToRect) {
            groupTargets = replaceOuterLinesWithRectangle(classified, gridBounds, groupTargets);
        }

        if (dialogOptions.group && groupTargets.length > 1) {
            groupProcessedItems(groupTargets);
        }
    }

    main();

    // =========================================
    // 選択の前処理 / Preparing the selection
    // =========================================

    /**
     * 選択を再帰的にたどり、グループや複合パスの中の PathItem も集める
     * @param {PageItem[]} sourceItems - 走査するアイテム
     * @param {PathItem[]} collectedPathItems - 見つけたパスを追加する配列
     * @returns {void}
     */
    function collectPathItemsRecursively(sourceItems, collectedPathItems) {
        for (var itemIndex = 0; itemIndex < sourceItems.length; itemIndex++) {
            var pageItem = sourceItems[itemIndex];
            if (!pageItem) continue;
            if (pageItem.typename === "PathItem") {
                collectedPathItems.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                collectPathItemsRecursively(pageItem.pageItems, collectedPathItems);
            } else if (pageItem.typename === "CompoundPathItem") {
                collectPathItemsRecursively(pageItem.pathItems, collectedPathItems);
            }
        }
    }

    /**
     * 塗りのあるパスを対象から除外する
     * @param {PathItem[]} pathItems - 対象のパス
     * @returns {PathItem[]} 塗りのないパス
     */
    function excludeFilledPaths(pathItems) {
        var unfilledPaths = [];
        for (var itemIndex = 0; itemIndex < pathItems.length; itemIndex++) {
            if (!pathItems[itemIndex].filled) unfilledPaths.push(pathItems[itemIndex]);
        }
        return unfilledPaths;
    }

    /**
     * 軸に平行な長方形（閉じた4点で、各辺が水平または垂直）かどうかを判定する
     * @param {PathItem} pathItem - 判定するパス
     * @returns {boolean} 軸に平行な長方形なら true
     */
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

    /**
     * 長方形パスのアンカーから外接座標を求める
     * @param {PathItem} pathItem - 長方形のパス
     * @returns {{minX: number, maxX: number, minY: number, maxY: number}} 外接座標
     */
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

    /**
     * パスの塗り・線の属性を記録する
     * @param {PathItem} pathItem - 記録するパス
     * @returns {Object} 塗り・線の属性
     */
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

    /**
     * 記録した塗り・線の属性をパスに適用する
     * @param {PathItem} pathItem - 適用先のパス
     * @param {Object} appearance - capturePathAppearance() の戻り値
     * @returns {void}
     */
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

    /**
     * 線の属性を引き継いだ2点の直線パスを作る（塗りなし）
     * @param {Object} parentContainer - 作成先のレイヤーまたはグループ
     * @param {number[]} firstPoint - 始点 [x, y]
     * @param {number[]} secondPoint - 終点 [x, y]
     * @param {Object} appearance - capturePathAppearance() の戻り値
     * @returns {PathItem} 作成した直線
     */
    function createLineFromAppearance(parentContainer, firstPoint, secondPoint, appearance) {
        var newLine = parentContainer.pathItems.add();
        newLine.setEntirePath([firstPoint, secondPoint]);
        newLine.closed = false;
        applyPathAppearance(newLine, appearance);
        newLine.filled = false;
        return newLine;
    }

    /**
     * 軸に平行な長方形を4本の罫線に分解する（元の長方形は削除）
     * @param {PathItem[]} pathItems - 対象のパス
     * @returns {{items: PathItem[], expansions: Object[]}} 分解後のパスと、取り消し用の記録
     */
    function expandRectanglesInSelection(pathItems) {
        var expandedItems = [];
        var rectangleExpansions = [];
        for (var itemIndex = 0; itemIndex < pathItems.length; itemIndex++) {
            var sourceItem = pathItems[itemIndex];
            if (!isAxisAlignedRectangle(sourceItem)) {
                expandedItems.push(sourceItem);
                continue;
            }
            var rectangleBounds = getRectangleBoundsFromPath(sourceItem);
            var appearance = capturePathAppearance(sourceItem);
            var parentContainer = sourceItem.parent;
            var topLeft = [rectangleBounds.minX, rectangleBounds.maxY];
            var topRight = [rectangleBounds.maxX, rectangleBounds.maxY];
            var bottomLeft = [rectangleBounds.minX, rectangleBounds.minY];
            var bottomRight = [rectangleBounds.maxX, rectangleBounds.minY];
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
                bounds: rectangleBounds,
                appearance: appearance,
                parent: parentContainer,
                generatedLines: generatedLines
            });
            /* 削除できない場合も分解は続ける / keep going even if the rectangle cannot be removed */
            try { sourceItem.remove(); } catch (rectangleRemoveError) { }
        }
        return { items: expandedItems, expansions: rectangleExpansions };
    }

    /**
     * 長方形の分解を取り消す（キャンセル時用）
     * @param {Object[]} rectangleExpansions - expandRectanglesInSelection() の記録
     * @returns {void}
     */
    function restoreRectangleExpansions(rectangleExpansions) {
        for (var expansionIndex = 0; expansionIndex < rectangleExpansions.length; expansionIndex++) {
            var expansion = rectangleExpansions[expansionIndex];
            for (var generatedLineIndex = 0; generatedLineIndex < expansion.generatedLines.length; generatedLineIndex++) {
                /* 既に消えている線は飛ばす / skip lines that are already gone */
                try { expansion.generatedLines[generatedLineIndex].remove(); } catch (lineRemoveError) { }
            }
            var rectangleBounds = expansion.bounds;
            var restoredRectangle = expansion.parent.pathItems.rectangle(
                rectangleBounds.maxY, rectangleBounds.minX,
                rectangleBounds.maxX - rectangleBounds.minX, rectangleBounds.maxY - rectangleBounds.minY
            );
            applyPathAppearance(restoredRectangle, expansion.appearance);
        }
    }

    // =========================================
    // 罫線の分類と外周 / Line classification and grid bounds
    // =========================================

    /**
     * アンカー2点の直線パスを抜き出し、水平線と垂直線に分類する
     * @param {PathItem[]} pathItems - 対象のパス
     * @returns {{horizontalLines: Object[], verticalLines: Object[]}} 水平線（path, y, minX, maxX）と垂直線（path, x, minY, maxY）
     */
    function classifySelectedStraightLines(pathItems) {
        var horizontalLines = [];
        var verticalLines = [];
        for (var itemIndex = 0; itemIndex < pathItems.length; itemIndex++) {
            var pathItem = pathItems[itemIndex];
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

    /**
     * 格子の外周座標（左端・右端・上端・下端）を計算する
     * 上下左右に貫通する罫線（最長線の90%以上の長さ）が複数あればそれを基準とし、
     * それより外側にはみ出した線は計算対象から除外する
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {{minX: number, maxX: number, minY: number, maxY: number}} 外周座標
     */
    function getGridBounds(horizontalLines, verticalLines) {
        var xRange = getValueRange(pickSpanningLines(verticalLines, "minY", "maxY"), "x");
        var yRange = getValueRange(pickSpanningLines(horizontalLines, "minX", "maxX"), "y");
        return { minX: xRange.min, maxX: xRange.max, minY: yRange.min, maxY: yRange.max };
    }

    /**
     * 最長線の 90% 以上の長さがある「貫通線」を選ぶ。2本未満なら全線を返す
     * @param {Object[]} lines - 線の記録
     * @param {string} startKey - 線の始端のプロパティ名（minX / minY）
     * @param {string} endKey - 線の終端のプロパティ名（maxX / maxY）
     * @returns {Object[]} 外枠の候補にする線
     */
    function pickSpanningLines(lines, startKey, endKey) {
        var SPAN_RATIO_THRESHOLD = 0.9;

        var maxSpan = 0;
        for (var spanIndex = 0; spanIndex < lines.length; spanIndex++) {
            var spanLength = lines[spanIndex][endKey] - lines[spanIndex][startKey];
            if (spanLength > maxSpan) maxSpan = spanLength;
        }

        var spanningLines = [];
        var spanThreshold = maxSpan * SPAN_RATIO_THRESHOLD;
        for (var pickIndex = 0; pickIndex < lines.length; pickIndex++) {
            if ((lines[pickIndex][endKey] - lines[pickIndex][startKey]) >= spanThreshold) {
                spanningLines.push(lines[pickIndex]);
            }
        }

        // 貫通線が2本以上あれば外枠候補として採用、なければ全線をフォールバック
        // / Use spanning lines as outer-frame candidates when at least 2 exist; otherwise fall back to all lines
        return spanningLines.length >= 2 ? spanningLines : lines;
    }

    /**
     * 線の記録から、指定プロパティの最小値と最大値を求める
     * @param {Object[]} lines - 線の記録（1本以上）
     * @param {string} propertyName - 調べるプロパティ名（x / y）
     * @returns {{min: number, max: number}} 最小値と最大値
     */
    function getValueRange(lines, propertyName) {
        var minValue = lines[0][propertyName], maxValue = lines[0][propertyName];
        for (var lineIndex = 1; lineIndex < lines.length; lineIndex++) {
            if (lines[lineIndex][propertyName] < minValue) minValue = lines[lineIndex][propertyName];
            if (lines[lineIndex][propertyName] > maxValue) maxValue = lines[lineIndex][propertyName];
        }
        return { min: minValue, max: maxValue };
    }

    /**
     * 線の記録からパスだけを取り出す
     * @param {Object[]} lines - 線の記録
     * @returns {PathItem[]} パス
     */
    function collectLinePaths(lines) {
        var linePaths = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) linePaths.push(lines[lineIndex].path);
        return linePaths;
    }

    /**
     * 線の記録から、指定プロパティの値を順に取り出す
     * @param {Object[]} lines - 線の記録（または座標グループ）
     * @param {string} propertyName - 取り出すプロパティ名
     * @returns {number[]} 値の配列
     */
    function collectLineValues(lines, propertyName) {
        var lineValues = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) lineValues.push(lines[lineIndex][propertyName]);
        return lineValues;
    }

    /**
     * 配列に同じ参照が含まれているかどうか
     * @param {Array} itemList - 調べる配列
     * @param {Object} targetItem - 探す参照
     * @returns {boolean} 含まれていれば true
     */
    function containsItem(itemList, targetItem) {
        for (var itemIndex = 0; itemIndex < itemList.length; itemIndex++) {
            if (itemList[itemIndex] === targetItem) return true;
        }
        return false;
    }

    // =========================================
    // 線の状態の記録と復元 / Capturing and restoring line states
    // =========================================

    /**
     * 2点の直線パスの両端を設定する（方向ハンドルもアンカー位置に揃える）
     * @param {PathItem} pathItem - 2点のパス
     * @param {number[]} firstPoint - 始点 [x, y]
     * @param {number[]} secondPoint - 終点 [x, y]
     * @returns {void}
     */
    function setLineEndpoints(pathItem, firstPoint, secondPoint) {
        pathItem.pathPoints[0].anchor = firstPoint;
        pathItem.pathPoints[0].leftDirection = firstPoint;
        pathItem.pathPoints[0].rightDirection = firstPoint;
        pathItem.pathPoints[1].anchor = secondPoint;
        pathItem.pathPoints[1].leftDirection = secondPoint;
        pathItem.pathPoints[1].rightDirection = secondPoint;
    }

    /**
     * プレビュー復元用に、水平線・垂直線の座標と線の属性を記録する
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {Object[]} 線ごとの状態
     */
    function captureLineStates(horizontalLines, verticalLines) {
        var allLines = horizontalLines.concat(verticalLines);
        var lineStates = [];
        for (var lineIndex = 0; lineIndex < allLines.length; lineIndex++) {
            var pathItem = allLines[lineIndex].path;
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
        return lineStates;
    }

    /**
     * 記録した線の状態を復元する
     * @param {Object[]} lineStates - captureLineStates() の戻り値
     * @returns {void}
     */
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

    // =========================================
    // プレビューとダイアログの値 / Preview and dialog values
    // =========================================

    /**
     * 線幅・破線・整列をまとめて適用する（プレビューと確定で共通）
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function applyLineOptions(classified, gridBounds, dialogOptions) {
        applyRepresentativeStrokeWidth(classified.horizontalLines, classified.verticalLines, dialogOptions);
        if (dialogOptions.convertDashedToSolid) convertDashedLinesToSolid(classified.horizontalLines, classified.verticalLines);
        alignLinesToGridBounds(classified.horizontalLines, classified.verticalLines, gridBounds, dialogOptions);
    }

    /**
     * プレビューを更新する（元に戻してから、プレビューと外枠分割が ON のときだけ整列する）
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {Object[]} originalLineStates - 元の線の状態
     * @returns {void}
     */
    function updatePreviewFromDialogState(dialogControls, classified, gridBounds, originalLineStates) {
        restoreLineStates(originalLineStates);
        // 外枠を四辺に分割が OFF の場合、本処理（罫線整列）を行わず元の状態のまま表示 / If split is OFF, skip the main alignment and show original state
        if (dialogControls.previewCheckbox.value && dialogControls.splitOuterFrameCheckbox.value) {
            var previewOptions = readOptionDialogState(dialogControls);
            // プレビュー時は frameToRect, group を適用しない / frameToRect and group are not applied in preview
            previewOptions.frameToRect = false;
            previewOptions.group = false;
            applyLineOptions(classified, gridBounds, previewOptions);
        }
        app.redraw();
    }

    /**
     * ダイアログの現在値を読み取る
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @returns {Object} ダイアログの設定
     */
    function readOptionDialogState(dialogControls) {
        var distributionMode = dialogControls.distributionNoneRadio.value ? "none"
            : dialogControls.distributionEvenMergedCellRadio.value ? "evenMergedCell"
                : "even";
        var isEvenMode = distributionMode !== "none";
        return {
            distributionMode: distributionMode,
            even: isEvenMode,
            evenHorizontal: isEvenMode && dialogControls.equalizeHorizontalCheckbox.value,
            evenVertical: isEvenMode && dialogControls.equalizeVerticalCheckbox.value,
            projectingCap: dialogControls.projectingCapCheckbox.value,
            convertDashedToSolid: dialogControls.convertDashedToSolidCheckbox.value,
            strokeWidthMode: dialogControls.strokeWidthMaxRadio.value ? "max"
                : dialogControls.strokeWidthMinRadio.value ? "min"
                    : dialogControls.strokeWidthSpecifiedRadio.value ? "specified"
                        : "average",
            specifiedStrokeWidthPt: dialogControls.strokeWidthSpecifiedRadio.value
                ? readNumericText(dialogControls.strokeWidthInput.text, 0) * dialogControls.strokeUnitInfo.pointsPerUnit
                : 0,
            lockFirstColumn: dialogControls.lockFirstColumnCheckbox.value,
            lockFirstRow: dialogControls.lockFirstRowCheckbox.value,
            frameToRect: dialogControls.frameToRectangleCheckbox.value,
            centerPointTextVertically: dialogControls.centerPointTextVerticallyCheckbox.value,
            group: dialogControls.groupingCheckbox.value,
            splitOuterFrame: dialogControls.splitOuterFrameCheckbox.value
        };
    }

    /**
     * 数値の文字列を読む（カンマは小数点として扱う）
     * @param {string} text - 入力された文字列
     * @param {number} fallbackValue - 数値にならないときの値
     * @returns {number} 読み取った数値
     */
    function readNumericText(text, fallbackValue) {
        var normalizedText = String(text).replace(/,/g, ".");
        var value = parseFloat(normalizedText);
        if (isNaN(value)) return fallbackValue;
        return value;
    }

    /**
     * 「指定」を選んだときだけ線幅の入力欄を有効にする
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @returns {void}
     */
    function syncSpecifiedStrokeWidthInput(dialogControls) {
        dialogControls.strokeWidthInput.enabled = dialogControls.strokeWidthSpecifiedRadio.value;
    }

    /**
     * ↑↓キーで数値を増減する（Shift は 10 刻み、Option は 0.1 刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すかどうか
     * @param {Function} onValueChanged - 値を変えたあとに呼ぶ関数
     * @returns {void}
     */
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

    // =========================================
    // 線幅と破線 / Stroke width and dashes
    // =========================================

    /**
     * 代表線幅をすべての線に適用する（代表線幅が 0 以下なら何もしない）
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function applyRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions) {
        var representativeStrokeWidth = getRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions);
        if (representativeStrokeWidth <= 0) return;
        var allLines = horizontalLines.concat(verticalLines);
        for (var lineIndex = 0; lineIndex < allLines.length; lineIndex++) {
            var pathItem = allLines[lineIndex].path;
            pathItem.stroked = true;
            pathItem.strokeWidth = representativeStrokeWidth;
        }
    }

    /**
     * 設定に応じた代表線幅（最大・最小・平均・指定）を返す
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {number} 代表線幅（pt）。線のある線が無ければ 0
     */
    function getRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions) {
        if (dialogOptions.strokeWidthMode === "specified") {
            return dialogOptions.specifiedStrokeWidthPt;
        }

        var strokeWidths = collectStrokeWidths(horizontalLines.concat(verticalLines));
        if (strokeWidths.length === 0) return 0;

        if (dialogOptions.strokeWidthMode === "max") return Math.max.apply(null, strokeWidths);
        if (dialogOptions.strokeWidthMode === "min") return Math.min.apply(null, strokeWidths);

        var totalStrokeWidth = 0;
        for (var averageIndex = 0; averageIndex < strokeWidths.length; averageIndex++) {
            totalStrokeWidth += strokeWidths[averageIndex];
        }
        return totalStrokeWidth / strokeWidths.length;
    }

    /**
     * 線のある線から線幅を集める
     * @param {Object[]} lines - 線の記録
     * @returns {number[]} 0 より大きい線幅
     */
    function collectStrokeWidths(lines) {
        var strokeWidths = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var pathItem = lines[lineIndex].path;
            if (pathItem.stroked && pathItem.strokeWidth > 0) strokeWidths.push(pathItem.strokeWidth);
        }
        return strokeWidths;
    }

    /**
     * 破線・点線の設定を外して実線にする
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {void}
     */
    function convertDashedLinesToSolid(horizontalLines, verticalLines) {
        var allLines = horizontalLines.concat(verticalLines);
        for (var lineIndex = 0; lineIndex < allLines.length; lineIndex++) {
            var pathItem = allLines[lineIndex].path;
            pathItem.strokeDashes = [];
            pathItem.strokeDashOffset = 0;
        }
    }

    // =========================================
    // 整列 / Alignment
    // =========================================

    /**
     * 配列要素の指定プロパティが targetValue に最も近いインデックスを返す
     * @param {number} targetValue - 基準の値
     * @param {Object[]} items - 候補（1件以上）
     * @param {string} propertyName - 比べるプロパティ名
     * @returns {number} 最も近い要素のインデックス
     */
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

    /* 同じ座標とみなす許容誤差（pt） / Tolerance to treat coordinates as identical (pt)
       注意：main() の実行より後で代入されるため、実行中は undefined のまま使われている（既存の不具合。未修正）
       Note: assigned after main() has run, so it is still undefined while the script runs (known bug, left as is) */
    var COORDINATE_TOLERANCE = 5.0;

    /**
     * 線を指定プロパティの近さでグループにまとめる（昇順）
     * @param {Object[]} lines - 線の記録
     * @param {string} propertyName - まとめる基準のプロパティ名（x / y）
     * @param {number} tolerance - 同じ座標とみなす差
     * @returns {{coord: number, lines: Object[]}[]} 座標ごとのグループ（coord はグループ内の平均）
     */
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

    /**
     * 最小値から最大値までを等間隔に分けた座標を返す
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @param {number} count - 座標の数（2以上）
     * @returns {number[]} 等間隔の座標
     */
    function getEvenlySpacedValues(minValue, maxValue, count) {
        var step = (maxValue - minValue) / (count - 1);
        var evenValues = [];
        for (var valueIndex = 0; valueIndex < count; valueIndex++) {
            evenValues.push(minValue + step * valueIndex);
        }
        return evenValues;
    }

    /**
     * 行（水平線）の新しい Y 座標を決める
     * 1行目を固定するときは、最上と上から2本目を残し、2本目と最下の間を均等に分ける
     * @param {number[]} rowYs - 下から上に並んだ Y 座標
     * @param {Object} gridBounds - 外周座標
     * @param {boolean} shouldEven - 均等にするかどうか（false なら rowYs をそのまま返す）
     * @param {boolean} lockFirstRow - 1行目を固定するかどうか
     * @returns {number[]} 新しい Y 座標
     */
    function resolveRowPositions(rowYs, gridBounds, shouldEven, lockFirstRow) {
        if (!shouldEven) return rowYs;
        var rowCount = rowYs.length;
        if (!(lockFirstRow && rowCount >= 3)) return getEvenlySpacedValues(gridBounds.minY, gridBounds.maxY, rowCount);

        // 1行目（最上）・2行目を固定し、2行目と最下を基準に残りを均等配置 / Lock 1st & 2nd from top; distribute the rest between 2nd-from-top and bottom
        var bottommostY = rowYs[0];
        var lockedAnchorY = rowYs[rowCount - 2];
        var lockedRowStep = (lockedAnchorY - bottommostY) / (rowCount - 2);
        var rowPositions = [];
        for (var rowIndex = 0; rowIndex < rowCount - 1; rowIndex++) {
            rowPositions.push(bottommostY + lockedRowStep * rowIndex);
        }
        rowPositions.push(rowYs[rowCount - 1]);
        return rowPositions;
    }

    /**
     * 列（垂直線）の新しい X 座標を決める
     * 1列目を固定するときは、左端と2本目を残し、2本目と最右の間を均等に分ける
     * @param {number[]} columnXs - 左から右に並んだ X 座標
     * @param {Object} gridBounds - 外周座標
     * @param {boolean} shouldEven - 均等にするかどうか（false なら columnXs をそのまま返す）
     * @param {boolean} lockFirstColumn - 1列目を固定するかどうか
     * @returns {number[]} 新しい X 座標
     */
    function resolveColumnPositions(columnXs, gridBounds, shouldEven, lockFirstColumn) {
        if (!shouldEven) return columnXs;
        var columnCount = columnXs.length;
        if (!(lockFirstColumn && columnCount >= 3)) return getEvenlySpacedValues(gridBounds.minX, gridBounds.maxX, columnCount);

        // 1本目・2本目を固定し、2本目と最右を基準に残りを均等配置 / Lock 1st & 2nd; distribute the rest between 2nd and rightmost
        var lockedAnchorX = columnXs[1];
        var lockedColumnStep = (columnXs[columnCount - 1] - lockedAnchorX) / (columnCount - 2);
        var columnPositions = [columnXs[0]];
        for (var columnIndex = 1; columnIndex < columnCount; columnIndex++) {
            columnPositions.push(lockedAnchorX + lockedColumnStep * (columnIndex - 1));
        }
        return columnPositions;
    }

    /**
     * 線の両端を設定する（必要なら先に突出線端にする）
     * @param {PathItem} pathItem - 2点のパス
     * @param {number[]} firstPoint - 始点 [x, y]
     * @param {number[]} secondPoint - 終点 [x, y]
     * @param {boolean} projectingCap - 突出線端にするかどうか
     * @returns {void}
     */
    function placeGridLine(pathItem, firstPoint, secondPoint, projectingCap) {
        if (projectingCap) pathItem.strokeCap = StrokeCap.PROJECTINGENDCAP;
        setLineEndpoints(pathItem, firstPoint, secondPoint);
    }

    /**
     * 水平線・垂直線を格子に揃える
     * @param {Object[]} horizontalLines - 水平線（均等にするときは Y の昇順に並べ替える）
     * @param {Object[]} verticalLines - 垂直線（均等にするときは X の昇順に並べ替える）
     * @param {Object} gridBounds - 外周座標
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function alignLinesToGridBounds(horizontalLines, verticalLines, gridBounds, dialogOptions) {
        if (dialogOptions.distributionMode === "evenMergedCell") {
            alignLinesWithMergedCellSupport(horizontalLines, verticalLines, gridBounds, dialogOptions);
            return;
        }

        // 水平線の Y 位置を決定（均等配置 or 元の Y）/ Determine Y positions for horizontal lines
        var shouldEvenRows = dialogOptions.evenHorizontal && horizontalLines.length >= 2;
        if (shouldEvenRows) horizontalLines.sort(function (lineA, lineB) { return lineA.y - lineB.y; });
        var horizontalLineYPositions = resolveRowPositions(collectLineValues(horizontalLines, "y"), gridBounds, shouldEvenRows, dialogOptions.lockFirstRow);

        // 垂直線の X 位置を決定（均等配置 or 元の X）/ Determine X positions for vertical lines
        var shouldEvenColumns = dialogOptions.evenVertical && verticalLines.length >= 2;
        if (shouldEvenColumns) verticalLines.sort(function (lineA, lineB) { return lineA.x - lineB.x; });
        var verticalLineXPositions = resolveColumnPositions(collectLineValues(verticalLines, "x"), gridBounds, shouldEvenColumns, dialogOptions.lockFirstColumn);

        // Illustrator は Y 上方向が正 / Illustrator uses positive Y upward
        var topCoordinateY = Math.max(gridBounds.minY, gridBounds.maxY);
        var bottomCoordinateY = Math.min(gridBounds.minY, gridBounds.maxY);

        for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            placeGridLine(horizontalLines[horizontalIndex].path,
                [gridBounds.minX, horizontalLineYPositions[horizontalIndex]],
                [gridBounds.maxX, horizontalLineYPositions[horizontalIndex]], dialogOptions.projectingCap);
        }
        for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
            placeGridLine(verticalLines[verticalIndex].path,
                [verticalLineXPositions[verticalIndex], topCoordinateY],
                [verticalLineXPositions[verticalIndex], bottomCoordinateY], dialogOptions.projectingCap);
        }
    }

    /**
     * 結合セル対応モード：同じ座標の線を 1 行／1 列としてまとめ、両端は直近の交差点にスナップする
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @param {Object} gridBounds - 外周座標
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function alignLinesWithMergedCellSupport(horizontalLines, verticalLines, gridBounds, dialogOptions) {
        var horizontalGroups = groupLinesByCoordinate(horizontalLines, "y", COORDINATE_TOLERANCE);
        var verticalGroups = groupLinesByCoordinate(verticalLines, "x", COORDINATE_TOLERANCE);

        // 各グループの新しい位置（行＝Y、列＝X）/ New position for each group (rows: Y, columns: X)
        var horizontalGroupYPositions = resolveRowPositions(collectLineValues(horizontalGroups, "coord"), gridBounds,
            dialogOptions.evenHorizontal && horizontalGroups.length >= 2, dialogOptions.lockFirstRow);
        var verticalGroupXPositions = resolveColumnPositions(collectLineValues(verticalGroups, "coord"), gridBounds,
            dialogOptions.evenVertical && verticalGroups.length >= 2, dialogOptions.lockFirstColumn);

        // 水平線：同一行のすべてに同じ Y、両端は最寄り列にスナップ / Apply same Y to all lines in a row; snap endpoints to nearest column
        for (var horizontalGroupIndex = 0; horizontalGroupIndex < horizontalGroups.length; horizontalGroupIndex++) {
            var rowCoordinateY = horizontalGroupYPositions[horizontalGroupIndex];
            var rowLines = horizontalGroups[horizontalGroupIndex].lines;
            for (var rowLineIndex = 0; rowLineIndex < rowLines.length; rowLineIndex++) {
                var horizontalLineData = rowLines[rowLineIndex];
                var leftColumnIndex = findClosestIndexByProperty(horizontalLineData.minX, verticalGroups, "coord");
                var rightColumnIndex = findClosestIndexByProperty(horizontalLineData.maxX, verticalGroups, "coord");
                placeGridLine(horizontalLineData.path,
                    [verticalGroupXPositions[leftColumnIndex], rowCoordinateY],
                    [verticalGroupXPositions[rightColumnIndex], rowCoordinateY], dialogOptions.projectingCap);
            }
        }

        // 垂直線：同一列のすべてに同じ X、両端は最寄り行にスナップ / Apply same X to all lines in a column; snap endpoints to nearest row
        for (var verticalGroupIndex = 0; verticalGroupIndex < verticalGroups.length; verticalGroupIndex++) {
            var columnCoordinateX = verticalGroupXPositions[verticalGroupIndex];
            var columnLines = verticalGroups[verticalGroupIndex].lines;
            for (var columnLineIndex = 0; columnLineIndex < columnLines.length; columnLineIndex++) {
                var verticalLineData = columnLines[columnLineIndex];
                var topRowIndex = findClosestIndexByProperty(verticalLineData.maxY, horizontalGroups, "coord");
                var bottomRowIndex = findClosestIndexByProperty(verticalLineData.minY, horizontalGroups, "coord");
                placeGridLine(verticalLineData.path,
                    [columnCoordinateX, horizontalGroupYPositions[topRowIndex]],
                    [columnCoordinateX, horizontalGroupYPositions[bottomRowIndex]], dialogOptions.projectingCap);
            }
        }
    }

    // =========================================
    // 後処理 / Post-processing
    // =========================================

    /**
     * 外周4本（上下の水平線・左右の垂直線）を選び、重複を除いた参照を返す
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {{uniqueOuterPaths: PathItem[], referencePath: PathItem}} 外周のパスと、線の属性を引き継ぐ基準（上辺）
     */
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
            if (!containsItem(uniqueOuterPaths, outerPathCandidates[candidateIndex])) uniqueOuterPaths.push(outerPathCandidates[candidateIndex]);
        }
        return { uniqueOuterPaths: uniqueOuterPaths, referencePath: topLine.path };
    }

    /**
     * 外周の長方形を作成する（塗り・線の属性は referencePath から引き継ぐ）
     * @param {PathItem} referencePath - 属性と重ね順の基準にするパス
     * @param {Object} gridBounds - 外周座標
     * @returns {PathItem} 作成した長方形
     */
    function createOuterFrameRectangle(referencePath, gridBounds) {
        var parentLayerOrGroup = referencePath.parent;
        // rectangle(top, left, width, height) — Illustrator は Y 上方向が正 / Illustrator uses positive Y upward
        var frameRectangle = parentLayerOrGroup.pathItems.rectangle(
            gridBounds.maxY, gridBounds.minX,
            gridBounds.maxX - gridBounds.minX, gridBounds.maxY - gridBounds.minY
        );
        applyPathAppearance(frameRectangle, capturePathAppearance(referencePath));
        frameRectangle.move(referencePath, ElementPlacement.PLACEAFTER);
        return frameRectangle;
    }

    /**
     * 外周4本の線を1つの長方形に置き換え、グループ化の対象を差し替える
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {PathItem[]} groupTargets - グループ化の対象（全罫線）
     * @returns {PathItem[]} 外周の線を除き、長方形を加えたグループ化の対象
     */
    function replaceOuterLinesWithRectangle(classified, gridBounds, groupTargets) {
        var outerPaths = pickOuterPaths(classified.horizontalLines, classified.verticalLines);
        var frameRectangle = createOuterFrameRectangle(outerPaths.referencePath, gridBounds);
        // outerPaths.uniqueOuterPaths を groupTargets から除外し frameRectangle を追加 / Exclude outerPaths.uniqueOuterPaths from groupTargets and add frameRectangle
        var innerGridItems = [];
        for (var targetIndex = 0; targetIndex < groupTargets.length; targetIndex++) {
            if (!containsItem(outerPaths.uniqueOuterPaths, groupTargets[targetIndex])) innerGridItems.push(groupTargets[targetIndex]);
        }
        innerGridItems.push(frameRectangle);
        // 参照が無効になる前にフィルタを終え、最後に削除 / Finish filtering before references become invalid, then remove them at the end
        for (var removeIndex = 0; removeIndex < outerPaths.uniqueOuterPaths.length; removeIndex++) {
            /* 削除できない線は残す / leave lines that cannot be removed */
            try { outerPaths.uniqueOuterPaths[removeIndex].remove(); } catch (removeError) { }
        }
        return innerGridItems;
    }

    /**
     * ポイント文字を、中心が含まれる行の上下中央へ移動する
     * @param {Object[]} horizontalLines - 水平線（y は分類時の座標）
     * @param {Object} gridBounds - 外周座標
     * @returns {void}
     */
    function centerPointTextVerticallyInCells(horizontalLines, gridBounds) {
        var rowYs = collectLineValues(horizontalLines, "y");
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

    /**
     * 渡されたアイテムを1つのグループにまとめる
     * @param {PageItem[]} processedItems - まとめるアイテム
     * @returns {GroupItem} 作成したグループ
     */
    function groupProcessedItems(processedItems) {
        var parentContainer = processedItems[0].parent;
        var tableGroup = parentContainer.groupItems.add();
        for (var itemIndex = 0; itemIndex < processedItems.length; itemIndex++) {
            processedItems[itemIndex].move(tableGroup, ElementPlacement.PLACEATEND);
        }
        return tableGroup;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 縦並びのオプションパネルを追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelPath - パネル見出しの LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentGroup, labelPath) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(labelPath));
        optionPanel.orientation = "column";
        optionPanel.alignChildren = "left";
        optionPanel.margins = PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * 横並びの行グループを追加する
     * @param {Panel} parentPanel - 追加先
     * @returns {Group} 追加した行
     */
    function addOptionRow(parentPanel) {
        var optionRow = parentPanel.add("group");
        optionRow.orientation = "row";
        optionRow.alignChildren = ["left", "center"];
        return optionRow;
    }

    /**
     * tooltip 付きのチェックボックスを追加する
     * @param {Object} parentContainer - 追加先のパネルまたはグループ
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentContainer, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentContainer.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * tooltip 付きのラジオボタンを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addOptionRadio(parentPanel, labelPath, tooltipPath) {
        var optionRadio = parentPanel.add("radiobutton", undefined, getLabel(labelPath));
        optionRadio.helpTip = getLabel(tooltipPath);
        return optionRadio;
    }

    /**
     * ダイアログを組み立てる
     * @param {{label: string, pointsPerUnit: number}} strokeUnitInfo - 線の単位
     * @param {boolean} hasExpandedRectangles - 外枠の長方形を分解したかどうか
     * @returns {Object} ダイアログ（optionDialog）と各コントロール
     */
    function buildOptionDialog(strokeUnitInfo, hasExpandedRectangles) {
        var optionDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        optionDialog.alignChildren = "left";
        optionDialog.margins = DIALOG_MARGINS;

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
        var preprocessingPanel = addOptionPanel(leftColumnGroup, "panel.preprocessing");
        var splitOuterFrameCheckbox = addOptionCheckbox(preprocessingPanel,
            "checkbox.splitFrameToFourSides", "tooltip.splitFrameToFourSides", !!hasExpandedRectangles);

        // 配置ラジオボタン / Distribution options
        var distributionPanel = addOptionPanel(leftColumnGroup, "panel.distribution");
        var distributionNoneRadio = addOptionRadio(distributionPanel, "radio.distributionNone", "tooltip.distributionNone");
        var distributionEvenRadio = addOptionRadio(distributionPanel, "radio.distributionEven", "tooltip.distributionEven");
        var distributionEvenMergedCellRadio = addOptionRadio(distributionPanel, "radio.distributionEvenMergedCell", "tooltip.distributionEvenMergedCell");

        // デフォルトは均等＋結合セル対応 / Default is Evenly + merged cells
        distributionEvenMergedCellRadio.value = true;

        // 平均化パネル（デフォルトは両方 OFF）/ Equalize panel (both OFF by default)
        var equalizePanel = addOptionPanel(leftColumnGroup, "panel.equalize");
        var equalizeVerticalRow = addOptionRow(equalizePanel);
        var equalizeVerticalCheckbox = addOptionCheckbox(equalizeVerticalRow, "checkbox.equalizeVertical", "tooltip.equalizeVertical", false);
        var lockFirstColumnCheckbox = addOptionCheckbox(equalizeVerticalRow, "checkbox.lockFirstColumn", "tooltip.lockFirstColumn", false);
        var equalizeHorizontalRow = addOptionRow(equalizePanel);
        var equalizeHorizontalCheckbox = addOptionCheckbox(equalizeHorizontalRow, "checkbox.equalizeHorizontal", "tooltip.equalizeHorizontal", false);
        var lockFirstRowCheckbox = addOptionCheckbox(equalizeHorizontalRow, "checkbox.lockFirstRow", "tooltip.lockFirstRow", false);

        // 線パネル / Line panel
        var linePanel = addOptionPanel(rightColumnGroup, "panel.line");
        var projectingCapCheckbox = addOptionCheckbox(linePanel, "checkbox.projectingCap", "tooltip.projectingCap", true);
        var convertDashedToSolidCheckbox = addOptionCheckbox(linePanel, "checkbox.convertDashedToSolid", "tooltip.convertDashedToSolid", false);

        // 線幅パネル / Stroke width panel
        var strokeWidthPanel = addOptionPanel(linePanel, "panel.strokeWidth");
        var strokeWidthMaxRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthMax", "tooltip.strokeWidthMax");
        var strokeWidthMinRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthMin", "tooltip.strokeWidthMin");
        var strokeWidthAverageRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthAverage", "tooltip.strokeWidthAverage");
        var strokeWidthSpecifiedRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthSpecified", "tooltip.strokeWidthSpecified");

        var strokeWidthSpecifiedInputGroup = addOptionRow(strokeWidthPanel);
        var strokeWidthInput = strokeWidthSpecifiedInputGroup.add("edittext", undefined, "0.25");
        strokeWidthInput.helpTip = getLabel("tooltip.strokeWidthSpecified");
        strokeWidthInput.characters = 4;
        strokeWidthSpecifiedInputGroup.add("statictext", undefined, strokeUnitInfo.label);

        // デフォルトは平均 / Default is average
        strokeWidthAverageRadio.value = true;
        strokeWidthInput.enabled = false;

        // 後処理パネル / Post-processing panel
        var postProcessingPanel = addOptionPanel(leftColumnGroup, "panel.postProcessing");
        var frameToRectangleCheckbox = addOptionCheckbox(postProcessingPanel,
            "checkbox.frameToRect", "tooltip.frameToRect", !!hasExpandedRectangles);
        var centerPointTextVerticallyCheckbox = addOptionCheckbox(postProcessingPanel,
            "checkbox.centerPointTextVertically", "tooltip.centerPointTextVertically", false);
        var groupingCheckbox = addOptionCheckbox(postProcessingPanel, "checkbox.grouping", "tooltip.grouping", false);

        // 下段：左＝プレビュー、右＝ボタン / Bottom row: left=preview, right=buttons
        var btnRowGroup = optionDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "fill";
        btnRowGroup.alignChildren = ["fill", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheckbox = addOptionCheckbox(btnLeftGroup, "checkbox.preview", "tooltip.preview", false);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "center"];

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            optionDialog: optionDialog,
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
    }

    /**
     * ダイアログのコントロールにクリック等の処理を結び付ける
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @param {Function} refreshPreview - プレビューを更新する関数
     * @returns {void}
     */
    function bindOptionDialogEvents(dialogControls, refreshPreview) {
        var equalizeVerticalCheckbox = dialogControls.equalizeVerticalCheckbox;
        var equalizeHorizontalCheckbox = dialogControls.equalizeHorizontalCheckbox;

        /**
         * 配置モードに応じて平均化パネルの有効／無効を同期する
         * @returns {void}
         */
        function syncEqualizePanelEnabled() {
            var isEqualizeEnabled = !dialogControls.distributionNoneRadio.value;
            equalizeVerticalCheckbox.enabled = isEqualizeEnabled;
            equalizeHorizontalCheckbox.enabled = isEqualizeEnabled;
            dialogControls.lockFirstColumnCheckbox.enabled = isEqualizeEnabled && equalizeVerticalCheckbox.value;
            dialogControls.lockFirstRowCheckbox.enabled = isEqualizeEnabled && equalizeHorizontalCheckbox.value;
        }

        /**
         * 縦罫／横罫のチェックボックスの処理を設定する。option＋クリックでもう一方も同じ状態にする
         * @param {Checkbox} clickedCheckbox - クリックされるチェックボックス
         * @param {Checkbox} otherCheckbox - そろえるチェックボックス
         * @returns {void}
         */
        function bindEqualizeCheckbox(clickedCheckbox, otherCheckbox) {
            clickedCheckbox.onClick = function () {
                if (ScriptUI.environment.keyboardState.altKey) {
                    otherCheckbox.value = clickedCheckbox.value;
                }
                syncEqualizePanelEnabled();
                refreshPreview();
            };
        }

        var onDistributionClick = function () {
            syncEqualizePanelEnabled();
            refreshPreview();
        };
        var onStrokeWidthClick = function () {
            syncSpecifiedStrokeWidthInput(dialogControls);
            refreshPreview();
        };

        syncEqualizePanelEnabled();

        dialogControls.previewCheckbox.onClick = refreshPreview;
        dialogControls.distributionNoneRadio.onClick = onDistributionClick;
        dialogControls.distributionEvenRadio.onClick = onDistributionClick;
        dialogControls.distributionEvenMergedCellRadio.onClick = onDistributionClick;
        bindEqualizeCheckbox(equalizeVerticalCheckbox, equalizeHorizontalCheckbox);
        bindEqualizeCheckbox(equalizeHorizontalCheckbox, equalizeVerticalCheckbox);
        dialogControls.lockFirstColumnCheckbox.onClick = refreshPreview;
        dialogControls.lockFirstRowCheckbox.onClick = refreshPreview;
        dialogControls.projectingCapCheckbox.onClick = refreshPreview;
        dialogControls.convertDashedToSolidCheckbox.onClick = refreshPreview;
        dialogControls.strokeWidthMaxRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthMinRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthAverageRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthSpecifiedRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthInput.onChanging = refreshPreview;
        changeValueByArrowKey(dialogControls.strokeWidthInput, false, function () {
            if (!dialogControls.strokeWidthSpecifiedRadio.value) dialogControls.strokeWidthSpecifiedRadio.value = true;
            syncSpecifiedStrokeWidthInput(dialogControls);
            refreshPreview();
        });
        // 外枠分割トグル：分割済みの線に対して整列処理（プレビュー）を再実行 / Toggle: re-run alignment preview on the already-split lines
        dialogControls.splitOuterFrameCheckbox.onClick = refreshPreview;
    }

    /**
     * ダイアログを表示してオプションを受け取る
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {Object[]} originalLineStates - 元の線の状態
     * @param {{label: string, pointsPerUnit: number}} strokeUnitInfo - 線の単位
     * @param {boolean} hasExpandedRectangles - 外枠の長方形を分解したかどうか
     * @returns {Object|null} ダイアログの設定。キャンセルなら null
     */
    function showOptionDialog(classified, gridBounds, originalLineStates, strokeUnitInfo, hasExpandedRectangles) {
        var dialogControls = buildOptionDialog(strokeUnitInfo, hasExpandedRectangles);
        bindOptionDialogEvents(dialogControls, function () {
            updatePreviewFromDialogState(dialogControls, classified, gridBounds, originalLineStates);
        });

        var dialogResult = dialogControls.optionDialog.show();
        restoreLineStates(originalLineStates);
        if (dialogResult !== 1) return null;
        return readOptionDialogState(dialogControls);
    }
})();
