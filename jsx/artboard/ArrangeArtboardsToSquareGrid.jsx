#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のアートボードを、全体の外形ができるだけ正方形に近づく行列で再配置します。
各アートボード内のアートワークも一緒に移動し、グリッドはカンバス中央に配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeArtboardsToSquareGrid.md

### Overview

Re-lays out every artboard so the whole grid's outline is as close to a square as possible.
Each artboard's artwork moves with it, and the grid is centered on the canvas.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeArtboardsToSquareGrid.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArrangeArtboardsToSquareGrid"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeArtboardsToSquareGrid.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeArtboardsToSquareGrid.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_MARGINS = 16;                 /* ダイアログ外周の余白 / dialog margin */
    var DIALOG_SPACING = 12;                 /* ダイアログ内の要素間隔 / dialog spacing */
    var PANEL_MARGINS = 16;                  /* 「整列設定」パネルの余白 / settings panel margin */
    var PANEL_SPACING = 10;                  /* 「整列設定」パネル内の要素間隔 / settings panel spacing */
    var NUMBER_FIELD_CHARS = 4;              /* 間隔・列数の入力欄の文字数 / width of the gap and column fields */
    var AUTO_BUTTON_HEIGHT = 22;             /* ［自動］ボタンの高さ / height of the Auto button */
    var PREVIEW_TEXT_SIZE = [330, 36];       /* 行列・外形サイズの表示欄 / size of the grid summary text */

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "アートボードを正方形に整列", en: "Arrange Artboards to Square" }
        },
        panel: {
            settings: { ja: "整列設定", en: "Layout settings" }
        },
        fieldLabel: {
            artboardCount: { ja: "アートボード数", en: "Artboards" },
            gap: { ja: "間隔", en: "Gap" },
            columns: { ja: "列数", en: "Columns" }
        },
        button: {
            auto: { ja: "自動", en: "Auto" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        status: {
            recommend: { ja: "推奨", en: "Recommended" },
            columnUnit: { ja: "列", en: "col" },
            rowUnit: { ja: "行", en: "row" }
        },
        tooltip: {
            gap: { ja: "アートボードどうしのあいだにあける間隔です。", en: "Space left between neighbouring artboards." },
            columns: {
                ja: "横に並べるアートボードの数です。行数はこの値から決まります。",
                en: "How many artboards to place in a row. The row count follows from this."
            },
            auto: {
                ja: "アートボード数から、できるだけ正方形に近くなる列数を入れ直します。",
                en: "Fills in the column count that comes closest to a square arrangement."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            needTwo: { ja: "アートボードが 2 つ以上必要です。", en: "At least two artboards are required." },
            badColumns: { ja: "列数は 1 以上の整数で入力してください。", en: "Enter the column count as an integer of 1 or more." },
            badGap: { ja: "間隔は 0 以上の数値で入力してください。", en: "Enter the gap as a number of 0 or more." },
            done: { ja: "整列が完了しました。", en: "Artboards arranged." },
            moveFailed: { ja: " 個のオブジェクトは移動できませんでした（ロック等）。", en: " object(s) could not be moved (locked, etc.)." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "fieldLabel.gap" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // カンバスとグリッド / Canvas & grid
    // =========================================

    /**
     * Illustrator の最大カンバス範囲を取得する（一時レイヤーで原点を測定）
     * Get Illustrator's max canvas bounds (measures the origin via a temp layer)
     * @returns {number[]} [left, top, right, bottom]
     */
    function getCanvasBounds() {
        var CANVAS_MAX_SIZE = 16383;
        var targetDoc = app.activeDocument;
        var wasModified = targetDoc.modified; // 計測前の変更フラグを退避 / remember the modified flag before measuring
        var tempLayer = targetDoc.layers.add();
        var tempTextFrame = tempLayer.textFrames.add();
        var canvasLeft = tempTextFrame.matrix.mValueTX;
        var canvasTop = tempTextFrame.matrix.mValueTY;
        tempLayer.remove();
        targetDoc.modified = wasModified; // 一時レイヤー追加で立った変更フラグを元に戻す / restore the modified flag
        return [canvasLeft, canvasTop, canvasLeft + CANVAS_MAX_SIZE, canvasTop - CANVAS_MAX_SIZE];
    }

    /**
     * 列数からグリッドの行数と外形サイズを求める
     * @param {number} itemCount - アートボード数
     * @param {number} columnCount - 列数
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {number} gap - 間隔（pt）
     * @returns {{rows: number, occupiedColumns: number, gridWidth: number, gridHeight: number}} 行数・実際に埋まる列数・外形サイズ
     */
    function computeGridSize(itemCount, columnCount, cellWidth, cellHeight, gap) {
        var rowCount = Math.ceil(itemCount / columnCount);
        var occupiedColumns = (columnCount < itemCount) ? columnCount : itemCount;
        return {
            rows: rowCount,
            occupiedColumns: occupiedColumns,
            gridWidth: occupiedColumns * cellWidth + (occupiedColumns - 1) * gap,
            gridHeight: rowCount * cellHeight + (rowCount - 1) * gap
        };
    }

    /**
     * グリッド全体をカンバスの天地・左右中央へ配置する原点を算出する
     * Compute the top-left origin that centers the whole artboard grid on the canvas.
     * @param {number[]} canvasBounds - カンバスの範囲 [left, top, right, bottom]
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {number} gap - 間隔（pt）
     * @param {number} columnCount - 列数
     * @param {number} itemCount - アートボード数
     * @returns {object} { left, top, cols, rows, gridWidth, gridHeight }
     */
    function computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, gap, columnCount, itemCount) {
        var gridSize = computeGridSize(itemCount, columnCount, cellWidth, cellHeight, gap);
        var canvasWidth = canvasBounds[2] - canvasBounds[0];
        var canvasHeight = canvasBounds[1] - canvasBounds[3];
        var leftMargin = Math.round((canvasWidth - gridSize.gridWidth) / 2);
        var topMargin = Math.round((canvasHeight - gridSize.gridHeight) / 2);
        return {
            left: canvasBounds[0] + leftMargin,
            top: canvasBounds[1] - topMargin,
            cols: gridSize.occupiedColumns,
            rows: gridSize.rows,
            gridWidth: gridSize.gridWidth,
            gridHeight: gridSize.gridHeight
        };
    }

    /**
     * 全体の外形が最も正方形に近づく列数を求める（セルの縦横比と間隔を考慮する）
     * Find the column count whose overall grid outline is closest to a square (accounts for cell aspect ratio and gap).
     * @param {number} itemCount - アートボード数
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {number} gap - 間隔（pt）
     * @returns {number} 列数
     */
    function chooseBestColumnCount(itemCount, cellWidth, cellHeight, gap) {
        var bestColumns = 1, bestScore = null, bestEmptyCells = 0, columnCandidate;
        for (columnCandidate = 1; columnCandidate <= itemCount; columnCandidate++) {
            var gridSize = computeGridSize(itemCount, columnCandidate, cellWidth, cellHeight, gap);
            var aspectRatio = gridSize.gridWidth / gridSize.gridHeight;
            var score = (aspectRatio >= 1) ? aspectRatio : (1 / aspectRatio); // 1 に近いほど正方形 / closer to 1 = squarer
            var emptyCells = gridSize.rows * columnCandidate - itemCount;      // 余りセル数 / unused cells
            if (bestScore === null
                || score < bestScore - 0.0001
                || (Math.abs(score - bestScore) <= 0.0001 && emptyCells < bestEmptyCells)) {
                bestScore = score;
                bestEmptyCells = emptyCells;
                bestColumns = columnCandidate;
            }
        }
        return bestColumns;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 全レイヤー（サブレイヤー含む）を再帰的に収集する / Collect all layers recursively, including sublayers
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer[]} すべてのレイヤー
     */
    function collectAllLayers(doc) {
        var allLayers = [];

        /**
         * レイヤーとその下のサブレイヤーを allLayers へ積む
         * @param {Layers} childLayers - 走査するレイヤー
         * @returns {void}
         */
        function collectLayersRecursive(childLayers) {
            for (var i = 0; i < childLayers.length; i++) {
                allLayers.push(childLayers[i]);
                collectLayersRecursive(childLayers[i].layers);
            }
        }
        collectLayersRecursive(doc.layers);
        return allLayers;
    }

    /**
     * レイヤー直下の最上位アイテムのみ収集する（グループ内は親ごと動かす）/ Collect only top-level items (children move with their parent)
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 最上位のアイテム
     */
    function collectTopLevelItems(doc) {
        var topLevelItems = [], pageItems = doc.pageItems, i;
        for (i = 0; i < pageItems.length; i++) {
            if (pageItems[i].parent && pageItems[i].parent.typename === "Layer") {
                topLevelItems.push(pageItems[i]);
            }
        }
        return topLevelItems;
    }

    /**
     * 矩形 [left, top, right, bottom] の中心点 / Center point of a [left, top, right, bottom] rect
     * @param {number[]} rect - 矩形
     * @returns {number[]} [x, y]
     */
    function getRectCenter(rect) {
        return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    /**
     * ロックを一時解除し、復元用の状態を返す / Temporarily unlock; returns the state needed to restore it
     * @param {Document} doc - 対象ドキュメント
     * @returns {{layers: Layer[], items: PageItem[]}} ロックを外したレイヤーとアイテム
     */
    function unlockLayersAndItems(doc) {
        var lockState = { layers: [], items: [] }, i;
        var allLayers = collectAllLayers(doc);
        for (i = 0; i < allLayers.length; i++) {
            if (allLayers[i].locked) { lockState.layers.push(allLayers[i]); allLayers[i].locked = false; }
        }
        var pageItems = doc.pageItems;
        for (i = 0; i < pageItems.length; i++) {
            if (pageItems[i].locked) { lockState.items.push(pageItems[i]); pageItems[i].locked = false; }
        }
        return lockState;
    }

    /**
     * unlockLayersAndItems で外したロックを元に戻す / Re-apply the locks released by unlockLayersAndItems
     * @param {{layers: Layer[], items: PageItem[]}} lockState - unlockLayersAndItems() の戻り値
     * @returns {void}
     */
    function restoreLockState(lockState) {
        var i;
        for (i = 0; i < lockState.items.length; i++) { lockState.items[i].locked = true; }
        for (i = 0; i < lockState.layers.length; i++) { lockState.layers[i].locked = true; }
    }

    /**
     * 小数 1 桁に丸めて文字列化 / Round to 1 decimal place and stringify
     * @param {number} value - 丸める値
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return String(Math.round(value * 10) / 10);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 列数・間隔の入力を読み取り、検証する
     * @param {EditText} columnInput - 列数の入力欄
     * @param {EditText} gapInput - 間隔の入力欄（定規単位）
     * @returns {{columns: number, gapInRulerUnit: number, errorLabelPath: string|null}} 読み取った値と、不正なときのエラー文言のパス
     */
    function readLayoutInputs(columnInput, gapInput) {
        var columns = parseInt(columnInput.text, 10);
        var gapInRulerUnit = parseFloat(gapInput.text);
        var errorLabelPath = null;
        if (isNaN(columns) || columns < 1) {
            errorLabelPath = "alert.badColumns";
        } else if (isNaN(gapInRulerUnit) || gapInRulerUnit < 0) {
            errorLabelPath = "alert.badGap";
        }
        return { columns: columns, gapInRulerUnit: gapInRulerUnit, errorLabelPath: errorLabelPath };
    }

    /**
     * 列数と間隔を指定するダイアログを表示する
     * @param {number} artboardCount - アートボード数
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {{columns: number, gapInPoints: number}|null} 指定した列数と間隔（キャンセルなら null）
     */
    function showGridDialog(artboardCount, cellWidth, cellHeight, rulerUnit) {
        var dialogResult = null;

        var gridDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        gridDialog.orientation = "column";
        gridDialog.alignChildren = "fill";
        gridDialog.margins = DIALOG_MARGINS;
        gridDialog.spacing = DIALOG_SPACING;

        gridDialog.add("statictext", undefined, getLabel("fieldLabel.artboardCount") + ": " + artboardCount);

        var settingsPanel = gridDialog.add("panel", undefined, getLabel("panel.settings"));
        settingsPanel.orientation = "column";
        settingsPanel.alignChildren = "left";
        settingsPanel.margins = PANEL_MARGINS;
        settingsPanel.spacing = PANEL_SPACING;

        /* 間隔 / Gap（既定 = アートボード幅の 1/5）/ default = 1/5 of artboard width */
        var defaultGapInPoints = cellWidth / 5;
        var gapGroup = settingsPanel.add("group");
        gapGroup.add("statictext", undefined, labelText("fieldLabel.gap"));
        var gapInput = gapGroup.add("edittext", undefined, formatNumber(defaultGapInPoints / rulerUnit.pointsPerUnit));
        gapInput.helpTip = getLabel("tooltip.gap");
        gapInput.characters = NUMBER_FIELD_CHARS;
        gapGroup.add("statictext", undefined, rulerUnit.label);

        /* 列数 / Columns */
        var recommendedColumns = chooseBestColumnCount(artboardCount, cellWidth, cellHeight, defaultGapInPoints);
        var columnGroup = settingsPanel.add("group");
        columnGroup.add("statictext", undefined, labelText("fieldLabel.columns"));
        var columnInput = columnGroup.add("edittext", undefined, String(recommendedColumns));
        columnInput.helpTip = getLabel("tooltip.columns");
        columnInput.characters = NUMBER_FIELD_CHARS;
        var btnAuto = columnGroup.add("button", undefined, getLabel("button.auto"));
        btnAuto.helpTip = getLabel("tooltip.auto");
        btnAuto.preferredSize.height = AUTO_BUTTON_HEIGHT;

        /* プレビュー / Preview */
        var previewText = settingsPanel.add("statictext", undefined, "", { multiline: true });
        previewText.preferredSize = PREVIEW_TEXT_SIZE;

        /**
         * 入力値から行列と外形サイズ、推奨列数を表示し直す
         * @returns {void}
         */
        function updatePreview() {
            var layoutInputs = readLayoutInputs(columnInput, gapInput);
            if (layoutInputs.errorLabelPath) { previewText.text = getLabel(layoutInputs.errorLabelPath); return; }
            var columns = layoutInputs.columns;
            var gapInPoints = layoutInputs.gapInRulerUnit * rulerUnit.pointsPerUnit;
            var gridSize = computeGridSize(artboardCount, columns, cellWidth, cellHeight, gapInPoints);
            var bestColumns = chooseBestColumnCount(artboardCount, cellWidth, cellHeight, gapInPoints);
            previewText.text = columns + " " + getLabel("status.columnUnit") + " x " + gridSize.rows + " " + getLabel("status.rowUnit")
                + "  /  " + formatNumber(gridSize.gridWidth / rulerUnit.pointsPerUnit)
                + " x " + formatNumber(gridSize.gridHeight / rulerUnit.pointsPerUnit) + " " + rulerUnit.label
                + "\n" + getLabel("status.recommend") + ": " + bestColumns + " " + getLabel("status.columnUnit");
        }

        gapInput.onChanging = updatePreview;
        columnInput.onChanging = updatePreview;
        btnAuto.onClick = function () {
            var gapInRulerUnit = parseFloat(gapInput.text);
            if (isNaN(gapInRulerUnit) || gapInRulerUnit < 0) gapInRulerUnit = 0;
            columnInput.text = String(chooseBestColumnCount(
                artboardCount, cellWidth, cellHeight, gapInRulerUnit * rulerUnit.pointsPerUnit));
            updatePreview();
        };
        updatePreview();

        /* ボタン / Buttons (Mac 規約: Cancel -> OK) */
        var btnRowGroup = gridDialog.add("group");
        btnRowGroup.alignment = "right";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        gridDialog.cancelElement = btnCancel;
        gridDialog.defaultElement = btnOK;

        btnOK.onClick = function () {
            var layoutInputs = readLayoutInputs(columnInput, gapInput);
            if (layoutInputs.errorLabelPath) { alert(getLabel(layoutInputs.errorLabelPath)); return; }
            dialogResult = { columns: layoutInputs.columns, gapInPoints: layoutInputs.gapInRulerUnit * rulerUnit.pointsPerUnit };
            gridDialog.close();
        };

        gridDialog.show();
        return dialogResult;
    }

    // =========================================
    // 再配置 / Rearrangement
    // =========================================

    /**
     * 座標系を元へ戻す（控えられなかったときは何もしない）
     * @param {CoordinateSystem|null} savedCoordinateSystem - 控えた座標系
     * @returns {void}
     */
    function restoreCoordinateSystem(savedCoordinateSystem) {
        if (savedCoordinateSystem !== null) {
            try { app.coordinateSystem = savedCoordinateSystem; } catch (e) { }
        }
    }

    /**
     * アートワークを中心点が含まれるアートボードに振り分ける（移動前に呼ぶ）
     * @param {PageItem[]} topLevelItems - 最上位のアイテム
     * @param {number[][]} artboardRects - 移動前のアートボード矩形
     * @returns {PageItem[][]} アートボードごとのアイテム
     */
    function assignItemsToArtboards(topLevelItems, artboardRects) {
        var itemsByArtboard = [];
        for (var i = 0; i < artboardRects.length; i++) itemsByArtboard.push([]);
        for (var j = 0; j < topLevelItems.length; j++) {
            var itemCenter = getRectCenter(topLevelItems[j].geometricBounds); // 中心点で所属を決定 / assign by center point
            for (var k = 0; k < artboardRects.length; k++) {
                var candidateRect = artboardRects[k];
                if (itemCenter[0] >= candidateRect[0] && itemCenter[0] <= candidateRect[2]
                    && itemCenter[1] <= candidateRect[1] && itemCenter[1] >= candidateRect[3]) {
                    itemsByArtboard[k].push(topLevelItems[j]);
                    break;
                }
            }
        }
        return itemsByArtboard;
    }

    /**
     * アートボードをグリッドのセルへ移し、所属アートワークも同じだけ動かす
     * @param {Document} doc - 対象ドキュメント
     * @param {number[][]} originalArtboardRects - 移動前のアートボード矩形
     * @param {PageItem[][]} itemsByArtboard - アートボードごとのアイテム
     * @param {object} gridLayout - gridOrigin / cellWidth / cellHeight / gapInPoints / columnCount
     * @returns {number} 移動できなかったアイテムの数
     */
    function moveArtboardsToGrid(doc, originalArtboardRects, itemsByArtboard, gridLayout) {
        var failedMoveCount = 0;
        for (var i = 0; i < originalArtboardRects.length; i++) {
            var gridRow = Math.floor(i / gridLayout.columnCount);
            var gridColumn = i % gridLayout.columnCount;
            var cellLeft = gridLayout.gridOrigin.left + gridColumn * (gridLayout.cellWidth + gridLayout.gapInPoints);
            var cellTop = gridLayout.gridOrigin.top - gridRow * (gridLayout.cellHeight + gridLayout.gapInPoints);

            var originalRect = originalArtboardRects[i];
            var artboardWidth = originalRect[2] - originalRect[0];
            var artboardHeight = originalRect[1] - originalRect[3];
            /* サイズが異なるアートボードはセル内で中央に置く / center within the cell */
            var newLeft = cellLeft + (gridLayout.cellWidth - artboardWidth) / 2;
            var newTop = cellTop - (gridLayout.cellHeight - artboardHeight) / 2;

            var deltaX = newLeft - originalRect[0];
            var deltaY = newTop - originalRect[1];

            doc.artboards[i].artboardRect = [newLeft, newTop, newLeft + artboardWidth, newTop - artboardHeight];

            var artboardItems = itemsByArtboard[i];
            for (var j = 0; j < artboardItems.length; j++) {
                try { artboardItems[j].translate(deltaX, deltaY); } catch (e) { failedMoveCount++; }
            }
        }
        return failedMoveCount;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで列数と間隔を決め、アートボードを正方形に近いグリッドへ並べ直す
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
        var doc = app.activeDocument;
        var artboardCount = doc.artboards.length;
        if (artboardCount < 2) { alert(getLabel("alert.needTwo")); return; }

        /* artboardRect と geometricBounds を同じ座標系に揃える / unify the coordinate space */
        var savedCoordinateSystem = null;
        try {
            savedCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }

        /* 元の矩形を保存し、最大サイズを 1 セルとする / capture original rects; cell = largest artboard */
        var originalArtboardRects = [], cellWidth = 0, cellHeight = 0;
        for (var i = 0; i < artboardCount; i++) {
            var artboardRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
            originalArtboardRects.push(artboardRect);
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];
            if (artboardWidth > cellWidth) cellWidth = artboardWidth;
            if (artboardHeight > cellHeight) cellHeight = artboardHeight;
        }

        var rulerUnit = getUnitInfo();

        var layoutSettings = showGridDialog(artboardCount, cellWidth, cellHeight, rulerUnit);
        if (!layoutSettings) { restoreCoordinateSystem(savedCoordinateSystem); return; }
        var gapInPoints = layoutSettings.gapInPoints;
        var columnCount = layoutSettings.columns;

        /* アートワークの所属アートボードを判定（移動前に）/ assign artwork to artboards (before any move) */
        var itemsByArtboard = assignItemsToArtboards(collectTopLevelItems(doc), originalArtboardRects);

        /* グリッド原点（カンバス中央）/ grid origin centered on the canvas */
        var canvasBounds = getCanvasBounds();
        var gridOrigin = computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, gapInPoints, columnCount, artboardCount);

        /* 再配置 / reposition */
        var lockState = unlockLayersAndItems(doc);
        var failedMoveCount = 0;
        try {
            failedMoveCount = moveArtboardsToGrid(doc, originalArtboardRects, itemsByArtboard, {
                gridOrigin: gridOrigin,
                cellWidth: cellWidth,
                cellHeight: cellHeight,
                gapInPoints: gapInPoints,
                columnCount: columnCount
            });
        } finally {
            restoreLockState(lockState);
            restoreCoordinateSystem(savedCoordinateSystem);
        }

        app.redraw();

        var resultMessage = getLabel("alert.done") + "\n"
            + getLabel("fieldLabel.artboardCount") + ": " + artboardCount + "  /  "
            + columnCount + " " + getLabel("status.columnUnit") + " x " + gridOrigin.rows + " " + getLabel("status.rowUnit");
        if (failedMoveCount > 0) resultMessage += "\n" + failedMoveCount + getLabel("alert.moveFailed");
        alert(resultMessage);
    }

    // =========================================
    // 実行 / Run
    // =========================================
    try {
        main();
    } catch (err) {
        alert("Error: " + err + (err.line ? " (line " + err.line + ")" : ""));
    }

})();
