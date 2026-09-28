#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した複数オブジェクトを行・列として自動判定し、元の配置ピッチを推定してグリッドを転置（行⇄列）します。
歯抜けのある配置や、1行だけ／1列だけの転置にも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TransposeGrid.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb5600abd495a

### Overview

Detects the rows and columns of the selected objects, estimates the original pitch, and transposes the grid so rows become columns and vice versa.
Gaps in the grid, and single-row or single-column layouts, are handled too.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TransposeGrid.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TransposeGrid";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TransposeGrid.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TransposeGrid.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb5600abd495a"; /* 紹介記事 / article URL */

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
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            needTwoItems: { ja: "オブジェクトを2つ以上選択してください。", en: "Select two or more objects." },
            cellConflict: {
                ja: "複数のオブジェクトが同じ位置（{row}行目・{column}列目）に割り当てられました。\n" +
                    "オブジェクトの並びを確認するか、許容値（SAME_COLUMN_TOLERANCE_PT / SAME_ROW_TOLERANCE_PT）を小さくしてください。",
                en: "More than one object was assigned to the same cell (row {row}, column {column}).\n" +
                    "Check the arrangement, or lower SAME_COLUMN_TOLERANCE_PT / SAME_ROW_TOLERANCE_PT."
            },
            pitchNotFound: {
                ja: "行または列の間隔を推定できませんでした。\n" +
                    "オブジェクトの並びと許容値を確認してください。",
                en: "Could not estimate the row or column spacing.\n" +
                    "Check the arrangement and the tolerance values."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを行・列に割り当て、グリッドを転置（行⇄列）する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var selectedItems = app.activeDocument.selection;
        if (!selectedItems || selectedItems.length < 2) {
            alert(getLabel("alert.needTwoItems"));
            return;
        }

        var itemEntries = collectItemEntries(selectedItems);
        var gridAxes = detectGridAxes(itemEntries);

        var gridAssignments = assignItemsToGrid(itemEntries, gridAxes);
        if (gridAssignments === null) return; /* 同一セル衝突。assignItemsToGrid が通知済み */

        var targetPitch = resolveTargetPitch(gridAxes);
        if (targetPitch === null) {
            alert(getLabel("alert.pitchNotFound"));
            return;
        }

        moveToTransposedCells(gridAssignments, gridAxes, targetPitch);
    }

    /**
     * 選択オブジェクトと、その左端X・上端Yの組を作る（visibleBounds の読み取りを1回にする）
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {Array} {item, left, top} の配列
     */
    function collectItemEntries(selectedItems) {
        var itemEntries = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = selectedItems[i].visibleBounds;
            itemEntries.push({ item: selectedItems[i], left: itemBounds[0], top: itemBounds[1] });
        }
        return itemEntries;
    }

    /**
     * 左端X・上端Yをまとめて、列と行の代表座標を求める
     * @param {Array} itemEntries - collectItemEntries() の戻り値
     * @returns {{columnXValues: number[], rowYValues: number[]}} 列は左→右、行は上→下
     */
    function detectGridAxes(itemEntries) {
        var leftXValues = [];
        var topYValues = [];
        for (var i = 0; i < itemEntries.length; i++) {
            leftXValues.push(itemEntries[i].left);
            topYValues.push(itemEntries[i].top);
        }

        var rowYValues = clusterCoordinates(topYValues, SAME_ROW_TOLERANCE_PT);
        /* Illustrator座標では上ほどYが大きいので「上→下」に並べ直す / top-down order */
        rowYValues.reverse();

        return {
            columnXValues: clusterCoordinates(leftXValues, SAME_COLUMN_TOLERANCE_PT),
            rowYValues: rowYValues
        };
    }

    /**
     * 近い値をまとめて、行または列の代表座標一覧を作る
     * @param {number[]} coordinateValues - まとめる座標値（この関数内で並べ替える）
     * @param {number} tolerancePt - 同じ行／列とみなす許容差（pt）
     * @returns {number[]} 昇順に並べた代表座標
     */
    function clusterCoordinates(coordinateValues, tolerancePt) {
        coordinateValues.sort(compareAscending);

        var clusterCenters = [];
        for (var i = 0; i < coordinateValues.length; i++) {
            var matchedIndex = -1;
            for (var j = 0; j < clusterCenters.length; j++) {
                if (Math.abs(coordinateValues[i] - clusterCenters[j]) <= tolerancePt) {
                    matchedIndex = j;
                    break;
                }
            }
            if (matchedIndex < 0) {
                clusterCenters.push(coordinateValues[i]);
            } else {
                /* 中心を軽く平均して安定させる / average lightly to keep the center stable */
                clusterCenters[matchedIndex] = (clusterCenters[matchedIndex] + coordinateValues[i]) / 2.0;
            }
        }

        clusterCenters.sort(compareAscending);
        return clusterCenters;
    }

    /**
     * sort() 用の昇順比較
     * @param {number} valueA - 比較する値
     * @param {number} valueB - 比較する値
     * @returns {number} 負なら valueA が先
     */
    function compareAscending(valueA, valueB) {
        return valueA - valueB;
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
     * @param {Array} itemEntries - collectItemEntries() の戻り値
     * @param {{columnXValues: number[], rowYValues: number[]}} gridAxes - 列と行の代表座標
     * @returns {Array|null} {item, rowIndex, columnIndex} の配列。重複があれば null
     */
    function assignItemsToGrid(itemEntries, gridAxes) {
        var occupiedCells = {};
        var gridAssignments = [];

        for (var i = 0; i < itemEntries.length; i++) {
            var rowIndex = findNearestIndex(gridAxes.rowYValues, itemEntries[i].top);
            var columnIndex = findNearestIndex(gridAxes.columnXValues, itemEntries[i].left);
            var cellKey = rowIndex + "," + columnIndex;

            if (occupiedCells[cellKey]) {
                /* 表示は1始まり / show 1-based positions */
                alert(getLabel("alert.cellConflict")
                    .replace("{row}", rowIndex + 1)
                    .replace("{column}", columnIndex + 1));
                return null;
            }

            occupiedCells[cellKey] = true;
            gridAssignments.push({ item: itemEntries[i].item, rowIndex: rowIndex, columnIndex: columnIndex });
        }

        return gridAssignments;
    }

    /**
     * 隣接差の中央値から代表ピッチを求める
     * @param {number[]} sortedValues - 昇順または降順に並んだ座標
     * @returns {number} 代表ピッチ（要素が1つ以下なら0）
     */
    function getMedianAdjacentGap(sortedValues) {
        if (sortedValues.length < 2) return 0;

        var adjacentGaps = [];
        for (var i = 1; i < sortedValues.length; i++) {
            adjacentGaps.push(Math.abs(sortedValues[i] - sortedValues[i - 1]));
        }
        adjacentGaps.sort(compareAscending);
        return adjacentGaps[Math.floor(adjacentGaps.length / 2)];
    }

    /**
     * 転置後に使う行・列のピッチを決める
     * 1行だけ／1列だけのときは、取れている側のピッチをもう一方へ流用する。
     * @param {{columnXValues: number[], rowYValues: number[]}} gridAxes - 列と行の代表座標
     * @returns {{columnPitchPt: number, rowPitchPt: number}|null} 使用するピッチ。推定できなければ null
     */
    function resolveTargetPitch(gridAxes) {
        var columnPitchPt = getMedianAdjacentGap(gridAxes.columnXValues);
        var rowPitchPt = getMedianAdjacentGap(gridAxes.rowYValues);

        /* 1行 → 1列。横ピッチを縦ピッチとして流用する / reuse the column pitch for the new rows */
        if (gridAxes.rowYValues.length === 1) rowPitchPt = columnPitchPt;
        /* 1列 → 1行。縦ピッチを横ピッチとして流用する / reuse the row pitch for the new columns */
        if (gridAxes.columnXValues.length === 1) columnPitchPt = rowPitchPt;

        if (columnPitchPt === 0 || rowPitchPt === 0) return null;
        return { columnPitchPt: columnPitchPt, rowPitchPt: rowPitchPt };
    }

    /**
     * 行と列を入れ替えた位置へ各オブジェクトを移動する。基準は元の左上に固定する
     * @param {Array} gridAssignments - assignItemsToGrid() の戻り値
     * @param {{columnXValues: number[], rowYValues: number[]}} gridAxes - 列と行の代表座標
     * @param {{columnPitchPt: number, rowPitchPt: number}} targetPitch - 転置後のピッチ
     * @returns {void}
     */
    function moveToTransposedCells(gridAssignments, gridAxes, targetPitch) {
        var originLeft = gridAxes.columnXValues[0];
        var originTop = gridAxes.rowYValues[0];

        for (var i = 0; i < gridAssignments.length; i++) {
            var targetItem = gridAssignments[i].item;
            /* 元の行番号が新しい列番号、元の列番号が新しい行番号になる / rows become columns and vice versa */
            var targetLeft = originLeft + gridAssignments[i].rowIndex * targetPitch.columnPitchPt;
            var targetTop = originTop - gridAssignments[i].columnIndex * targetPitch.rowPitchPt;

            var itemBounds = targetItem.visibleBounds;
            targetItem.translate(targetLeft - itemBounds[0], targetTop - itemBounds[1]);
        }
    }

    main();

})();
