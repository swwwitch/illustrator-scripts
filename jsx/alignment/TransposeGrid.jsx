#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した複数オブジェクトを行・列として自動判定し、元の配置ピッチを推定してグリッドを転置（行⇄列）します。
歯抜けのある配置や、1行だけ／1列だけの転置にも対応します。

詳細は README を参照してください。

### Overview

Detects the rows and columns of the selected objects, estimates the original pitch, and transposes the grid so rows become columns and vice versa.
Gaps in the grid, and single-row or single-column layouts, are handled too.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TransposeGrid";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TransposeGrid.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TransposeGrid.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 同じ列とみなす左端Xの許容差（pt）/ tolerance for treating lefts as one column */
    var SAME_COLUMN_TOLERANCE_PT = 8.0;
    /* 同じ行とみなす上端Yの許容差（pt）/ tolerance for treating tops as one row */
    var SAME_ROW_TOLERANCE_PT = 8.0;

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを行・列に割り当て、グリッドを転置（行⇄列）する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var selectedObjects = app.activeDocument.selection;
        if (!selectedObjects || selectedObjects.length < 1) {
            alert("オブジェクトを選択してください。");
            return;
        }

        var selectedItems = [];
        for (var i = 0; i < selectedObjects.length; i++) selectedItems.push(selectedObjects[i]);

        /* すべての左端X・上端Yを集めて行・列の候補を作る / collect lefts and tops to derive the axes */
        var leftXValues = [];
        var topYValues = [];
        for (var k = 0; k < selectedItems.length; k++) {
            leftXValues.push(getLeftX(selectedItems[k]));
            topYValues.push(getTopY(selectedItems[k]));
        }

        var columnXValues = clusterValues(leftXValues, SAME_COLUMN_TOLERANCE_PT); /* 左→右 */
        var rowYValues = clusterValues(topYValues, SAME_ROW_TOLERANCE_PT);
        /* Illustrator座標では上ほどYが大きいので「上→下」に並べ直す / top-down order */
        rowYValues.sort(function (valueA, valueB) { return valueB - valueA; });

        var rowCount = rowYValues.length;
        var columnCount = columnXValues.length;

        var gridAssignments = assignItemsToGrid(selectedItems, rowYValues, columnXValues);
        if (gridAssignments === null) return; /* 同一セル衝突。assignItemsToGrid が通知済み */

        var targetPitch = resolveTargetPitch(columnXValues, rowYValues, rowCount, columnCount);
        if (targetPitch === null) return; /* ピッチを推定できず。resolveTargetPitch が通知済み */

        /* 転置後グリッドの基準は元の左上に固定する / keep the original top-left as the origin */
        var originLeft = columnXValues[0];
        var originTop = rowYValues[0];

        for (var m = 0; m < gridAssignments.length; m++) {
            var targetItem = gridAssignments[m].item;
            /* 行と列を入れ替えた位置へ移動する / rows become columns and vice versa */
            var targetLeft = originLeft + gridAssignments[m].rowIndex * targetPitch.columnPitchPt;
            var targetTop = originTop - gridAssignments[m].columnIndex * targetPitch.rowPitchPt;

            var itemBounds = targetItem.visibleBounds;
            targetItem.translate(targetLeft - itemBounds[0], targetTop - itemBounds[1]);
        }
    }

    /**
     * オブジェクトの左端X（visibleBounds[0]）を返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number} 左端X（pt）
     */
    function getLeftX(pageItem) {
        return pageItem.visibleBounds[0];
    }

    /**
     * オブジェクトの上端Y（visibleBounds[1]）を返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number} 上端Y（pt）
     */
    function getTopY(pageItem) {
        return pageItem.visibleBounds[1];
    }

    /**
     * 近い値をまとめて、行または列の代表座標一覧を作る
     * @param {number[]} values - まとめる座標値
     * @param {number} tolerancePt - 同じ行／列とみなす許容差（pt）
     * @returns {number[]} 昇順に並べた代表座標
     */
    function clusterValues(values, tolerancePt) {
        values.sort(function (valueA, valueB) { return valueA - valueB; });

        var clusterCenters = [];
        for (var i = 0; i < values.length; i++) {
            var matchedIndex = -1;
            for (var j = 0; j < clusterCenters.length; j++) {
                if (Math.abs(values[i] - clusterCenters[j]) <= tolerancePt) {
                    matchedIndex = j;
                    break;
                }
            }
            if (matchedIndex < 0) {
                clusterCenters.push(values[i]);
            } else {
                /* 中心を軽く平均して安定させる / average lightly to keep the center stable */
                clusterCenters[matchedIndex] = (clusterCenters[matchedIndex] + values[i]) / 2.0;
            }
        }

        clusterCenters.sort(function (valueA, valueB) { return valueA - valueB; });
        return clusterCenters;
    }

    /**
     * 値に最も近い要素の添字を返す
     * @param {number[]} candidateValues - 探索対象の座標一覧
     * @param {number} value - 基準となる座標
     * @returns {number} 最も近い要素の添字
     */
    function findNearestIndex(candidateValues, value) {
        var nearestIndex = 0;
        var nearestDistance = Math.abs(value - candidateValues[0]);
        for (var i = 1; i < candidateValues.length; i++) {
            var distance = Math.abs(value - candidateValues[i]);
            if (distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /**
     * 各オブジェクトを (行, 列) に割り当てる。歯抜けは許容するが、同一セルの重複は中止する
     * @param {PageItem[]} selectedItems - 割り当てるオブジェクト
     * @param {number[]} rowYValues - 行の代表Y（上→下）
     * @param {number[]} columnXValues - 列の代表X（左→右）
     * @returns {Array|null} {item, rowIndex, columnIndex} の配列。重複があれば null
     */
    function assignItemsToGrid(selectedItems, rowYValues, columnXValues) {
        var occupiedCells = {};
        var gridAssignments = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var rowIndex = findNearestIndex(rowYValues, getTopY(selectedItems[i]));
            var columnIndex = findNearestIndex(columnXValues, getLeftX(selectedItems[i]));
            var cellKey = rowIndex + "," + columnIndex;

            if (occupiedCells[cellKey]) {
                alert("同一セルに複数オブジェクトが割り当てられました。\n" +
                    "許容値（SAME_COLUMN_TOLERANCE_PT / SAME_ROW_TOLERANCE_PT）を下げるか、整列状態を確認してください。\n" +
                    "衝突セル: (" + rowIndex + "," + columnIndex + ")");
                return null;
            }

            occupiedCells[cellKey] = true;
            gridAssignments.push({ item: selectedItems[i], rowIndex: rowIndex, columnIndex: columnIndex });
        }

        return gridAssignments;
    }

    /**
     * 隣接差の中央値から代表ピッチを求める
     * @param {number[]} sortedValues - 昇順または降順に並んだ座標
     * @returns {number} 代表ピッチ（要素が1つ以下なら0）
     */
    function medianAdjacentGap(sortedValues) {
        if (sortedValues.length < 2) return 0;

        var gaps = [];
        for (var i = 1; i < sortedValues.length; i++) {
            gaps.push(Math.abs(sortedValues[i] - sortedValues[i - 1]));
        }
        gaps.sort(function (valueA, valueB) { return valueA - valueB; });
        return gaps[Math.floor(gaps.length / 2)];
    }

    /**
     * 転置後に使う行・列のピッチを決める
     * 1行だけ／1列だけのときは、取れている側のピッチをもう一方へ流用する。
     * @param {number[]} columnXValues - 列の代表X
     * @param {number[]} rowYValues - 行の代表Y
     * @param {number} rowCount - 検出した行数
     * @param {number} columnCount - 検出した列数
     * @returns {{columnPitchPt: number, rowPitchPt: number}|null} 使用するピッチ。推定できなければ null
     */
    function resolveTargetPitch(columnXValues, rowYValues, rowCount, columnCount) {
        var columnPitchPt = medianAdjacentGap(columnXValues);
        var rowPitchPt = medianAdjacentGap(rowYValues);

        if (rowCount === 1 && columnCount === 1) {
            alert("1つしか選択されていないため、転置できません。");
            return null;
        }

        /* 1行 → 1列。横ピッチを縦ピッチとして流用する / reuse the column pitch for the new rows */
        if (rowCount === 1) {
            if (columnPitchPt === 0) {
                alert("1行は検出できましたが、列ピッチが推定できませんでした。");
                return null;
            }
            return { columnPitchPt: columnPitchPt, rowPitchPt: columnPitchPt };
        }

        /* 1列 → 1行。縦ピッチを横ピッチとして流用する / reuse the row pitch for the new columns */
        if (columnCount === 1) {
            if (rowPitchPt === 0) {
                alert("1列は検出できましたが、行ピッチが推定できませんでした。");
                return null;
            }
            return { columnPitchPt: rowPitchPt, rowPitchPt: rowPitchPt };
        }

        if (columnPitchPt === 0 || rowPitchPt === 0) {
            alert("行または列のピッチが推定できませんでした。許容値や整列状態を確認してください。");
            return null;
        }

        return { columnPitchPt: columnPitchPt, rowPitchPt: rowPitchPt };
    }

    main();

})();
