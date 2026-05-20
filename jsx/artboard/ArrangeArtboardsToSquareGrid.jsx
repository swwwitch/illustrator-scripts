#target illustrator

/* ============================================================
 * アートボードを正方形グリッドに整列 / Arrange Artboards into a Square Grid
 *
 * 複数のアートボードを、全体の外形ができるだけ正方形（縦横比 1:1）
 * に近づく行列で再配置する。各アートボード内のアートワークも一緒に
 * 移動する。グリッドはカンバス中央に配置される。
 *
 * Re-lays out every artboard so the whole grid's outline is as
 * close to a square as possible, and moves each artboard's
 * artwork together with it. The grid is centered on the canvas.
 * ============================================================ */

var SCRIPT_VERSION = "v1.0";

// ============================================================
// ローカライズ / Localization
// ============================================================
var L = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

var LABELS = {
    dialogTitle: { ja: "アートボードを正方形に整列 " + SCRIPT_VERSION, en: "Arrange Artboards to Square " + SCRIPT_VERSION },
    abCount: { ja: "アートボード数", en: "Artboards" },
    settings: { ja: "整列設定", en: "Layout settings" },
    gap: { ja: "間隔：", en: "Gap: " },
    columns: { ja: "列数：", en: "Columns: " },
    autoBtn: { ja: "自動", en: "Auto" },
    recommend: { ja: "推奨", en: "Recommended" },
    colWord: { ja: "列", en: "col" },
    rowWord: { ja: "行", en: "row" },
    cancelBtn: { ja: "キャンセル", en: "Cancel" },
    noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    needTwo: { ja: "アートボードが 2 つ以上必要です。", en: "At least two artboards are required." },
    badColumns: { ja: "列数は 1 以上の整数で入力してください。", en: "Enter the column count as an integer of 1 or more." },
    badGap: { ja: "間隔は 0 以上の数値で入力してください。", en: "Enter the gap as a number of 0 or more." },
    done: { ja: "整列が完了しました。", en: "Artboards arranged." },
    moveFailed: { ja: " 個のオブジェクトは移動できませんでした（ロック等）。", en: " object(s) could not be moved (locked, etc.)." }
};

function labelText(key) {
    return LABELS[key][L];
}

// ============================================================
// 雛形（提供された関数） / Provided helper functions
// ============================================================

/* Illustrator の最大カンバス範囲を取得（一時レイヤーで原点を測定）/ Get Illustrator's max canvas bounds (measures the origin via a temp layer) */
function getCanvasBounds() {
    var CANVAS_MAX_SIZE = 16383;
    var targetDoc = app.activeDocument;
    var wasModified = targetDoc.modified; // 計測前の変更フラグを退避 / remember the modified flag before measuring
    var tempLayer = targetDoc.layers.add();
    var tempTextFrame = tempLayer.textFrames.add();
    var left = tempTextFrame.matrix.mValueTX;
    var top = tempTextFrame.matrix.mValueTY;
    tempLayer.remove();
    targetDoc.modified = wasModified; // 一時レイヤー追加で立った変更フラグを元に戻す / restore the modified flag
    // [left, top, right, bottom]
    return [left, top, left + CANVAS_MAX_SIZE, top - CANVAS_MAX_SIZE];
}

/*
    グリッド全体をカンバスの天地・左右中央へ配置する原点を算出
    Compute the top-left origin that centers the whole artboard grid on the canvas.
    返り値 / Returns: { left, top, cols, rows, gridWidth, gridHeight }
*/
function computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, gap, columnCount, itemCount) {
    var rows = Math.ceil(itemCount / columnCount);
    var occupiedColumns = (columnCount < itemCount) ? columnCount : itemCount;
    var gridWidth = occupiedColumns * cellWidth + (occupiedColumns - 1) * gap;
    var gridHeight = rows * cellHeight + (rows - 1) * gap;
    var canvasWidth = canvasBounds[2] - canvasBounds[0];
    var canvasHeight = canvasBounds[1] - canvasBounds[3];
    var leftMargin = Math.round((canvasWidth - gridWidth) / 2);
    var topMargin = Math.round((canvasHeight - gridHeight) / 2);
    return {
        left: canvasBounds[0] + leftMargin,
        top: canvasBounds[1] - topMargin,
        cols: occupiedColumns,
        rows: rows,
        gridWidth: gridWidth,
        gridHeight: gridHeight
    };
}

// ============================================================
// ユーティリティ / Utilities
// ============================================================

/* 現在の定規単位の情報を返す（表示名と pt 換算係数）/ Ruler unit info: display label + points-per-unit */
function getRulerUnitInfo() {
    var unitCode = 2;
    try { unitCode = app.preferences.getIntegerPreference("rulerType"); } catch (e) { }
    // 0:inch 1:mm 2:pt 3:pica 4:cm 5:Q/H 6:px
    var unitTable = [
        { label: "inch", pointsPerUnit: 72.0 },
        { label: "mm", pointsPerUnit: 72.0 / 25.4 },
        { label: "pt", pointsPerUnit: 1.0 },
        { label: "pica", pointsPerUnit: 12.0 },
        { label: "cm", pointsPerUnit: 72.0 / 2.54 },
        { label: "Q", pointsPerUnit: (72.0 / 25.4) * 0.25 },
        { label: "px", pointsPerUnit: 1.0 }
    ];
    if (unitCode < 0 || unitCode > 6) unitCode = 2;
    return unitTable[unitCode];
}

/*
    全体の外形が最も正方形に近づく列数を求める
    Find the column count whose overall grid outline is closest to a square.
    セルの縦横比と間隔を考慮する / accounts for cell aspect ratio and gap.
*/
function chooseBestColumnCount(itemCount, cellWidth, cellHeight, gap) {
    var bestColumns = 1, bestScore = null, bestEmptyCells = 0, columnCandidate;
    for (columnCandidate = 1; columnCandidate <= itemCount; columnCandidate++) {
        var rows = Math.ceil(itemCount / columnCandidate);
        var occupiedColumns = (columnCandidate < itemCount) ? columnCandidate : itemCount;
        var gridWidth = occupiedColumns * cellWidth + (occupiedColumns - 1) * gap;
        var gridHeight = rows * cellHeight + (rows - 1) * gap;
        var aspectRatio = gridWidth / gridHeight;
        var score = (aspectRatio >= 1) ? aspectRatio : (1 / aspectRatio); // 1 に近いほど正方形 / closer to 1 = squarer
        var emptyCells = rows * columnCandidate - itemCount;               // 余りセル数 / unused cells
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

/* 全レイヤー（サブレイヤー含む）を再帰的に収集 / Collect all layers recursively, including sublayers */
function collectAllLayers(document) {
    var allLayers = [];
    function walk(layers) {
        var i;
        for (i = 0; i < layers.length; i++) {
            allLayers.push(layers[i]);
            walk(layers[i].layers);
        }
    }
    walk(document.layers);
    return allLayers;
}

/* レイヤー直下の最上位アイテムのみ収集（グループ内は親ごと動かす）/ Collect only top-level items (children move with their parent) */
function collectTopLevelItems(document) {
    var topLevelItems = [], pageItems = document.pageItems, i;
    for (i = 0; i < pageItems.length; i++) {
        if (pageItems[i].parent && pageItems[i].parent.typename === "Layer") {
            topLevelItems.push(pageItems[i]);
        }
    }
    return topLevelItems;
}

/* 矩形 [left, top, right, bottom] の中心点 / Center point of a [left, top, right, bottom] rect */
function getRectCenter(rect) {
    return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
}

/* ロックを一時解除し、復元用の状態を返す / Temporarily unlock; returns the state needed to restore it */
function unlockLayersAndItems(document) {
    var lockState = { layers: [], items: [] }, i;
    var layers = collectAllLayers(document);
    for (i = 0; i < layers.length; i++) {
        if (layers[i].locked) { lockState.layers.push(layers[i]); layers[i].locked = false; }
    }
    var pageItems = document.pageItems;
    for (i = 0; i < pageItems.length; i++) {
        if (pageItems[i].locked) { lockState.items.push(pageItems[i]); pageItems[i].locked = false; }
    }
    return lockState;
}

/* unlockLayersAndItems で外したロックを元に戻す / Re-apply the locks released by unlockLayersAndItems */
function restoreLockState(lockState) {
    var i;
    for (i = 0; i < lockState.items.length; i++) { lockState.items[i].locked = true; }
    for (i = 0; i < lockState.layers.length; i++) { lockState.layers[i].locked = true; }
}

/* 小数 1 桁に丸めて文字列化 / Round to 1 decimal place and stringify */
function formatNumber(value) {
    return String(Math.round(value * 10) / 10);
}

// ============================================================
// ダイアログ / Dialog
// ============================================================
function showDialog(artboardCount, cellWidth, cellHeight, rulerUnit) {
    var dialogResult = null;

    var dialog = new Window("dialog", labelText("dialogTitle"));
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = 16;
    dialog.spacing = 12;

    dialog.add("statictext", undefined, labelText("abCount") + ": " + artboardCount);

    var panel = dialog.add("panel", undefined, labelText("settings"));
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.margins = 16;
    panel.spacing = 10;

    /* 間隔 / Gap（既定 = アートボード幅の 1/5）/ default = 1/5 of artboard width */
    var defaultGapInPoints = cellWidth / 5;
    var gapGroup = panel.add("group");
    gapGroup.add("statictext", undefined, labelText("gap"));
    var gapInput = gapGroup.add("edittext", undefined, formatNumber(defaultGapInPoints / rulerUnit.pointsPerUnit));
    gapInput.characters = 4;
    gapGroup.add("statictext", undefined, rulerUnit.label);

    /* 列数 / Columns */
    var recommendedColumns = chooseBestColumnCount(artboardCount, cellWidth, cellHeight, defaultGapInPoints);
    var columnGroup = panel.add("group");
    columnGroup.add("statictext", undefined, labelText("columns"));
    var columnInput = columnGroup.add("edittext", undefined, String(recommendedColumns));
    columnInput.characters = 4;
    var autoButton = columnGroup.add("button", undefined, labelText("autoBtn"));
    autoButton.preferredSize.height = 22;

    /* プレビュー / Preview */
    var preview = panel.add("statictext", undefined, "", { multiline: true });
    preview.preferredSize = [330, 36];

    function updatePreview() {
        var columns = parseInt(columnInput.text, 10);
        var gapInRulerUnit = parseFloat(gapInput.text);
        if (isNaN(columns) || columns < 1) { preview.text = labelText("badColumns"); return; }
        if (isNaN(gapInRulerUnit) || gapInRulerUnit < 0) { preview.text = labelText("badGap"); return; }
        var gapInPoints = gapInRulerUnit * rulerUnit.pointsPerUnit;
        var rows = Math.ceil(artboardCount / columns);
        var occupiedColumns = (columns < artboardCount) ? columns : artboardCount;
        var gridWidth = occupiedColumns * cellWidth + (occupiedColumns - 1) * gapInPoints;
        var gridHeight = rows * cellHeight + (rows - 1) * gapInPoints;
        var bestColumns = chooseBestColumnCount(artboardCount, cellWidth, cellHeight, gapInPoints);
        preview.text = columns + " " + labelText("colWord") + " x " + rows + " " + labelText("rowWord")
            + "  /  " + formatNumber(gridWidth / rulerUnit.pointsPerUnit)
            + " x " + formatNumber(gridHeight / rulerUnit.pointsPerUnit) + " " + rulerUnit.label
            + "\n" + labelText("recommend") + ": " + bestColumns + " " + labelText("colWord");
    }

    gapInput.onChanging = updatePreview;
    columnInput.onChanging = updatePreview;
    autoButton.onClick = function () {
        var gapInRulerUnit = parseFloat(gapInput.text);
        if (isNaN(gapInRulerUnit) || gapInRulerUnit < 0) gapInRulerUnit = 0;
        columnInput.text = String(chooseBestColumnCount(
            artboardCount, cellWidth, cellHeight, gapInRulerUnit * rulerUnit.pointsPerUnit));
        updatePreview();
    };
    updatePreview();

    /* ボタン / Buttons (Mac 規約: Cancel -> OK) */
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var cancelButton = buttonGroup.add("button", undefined, labelText("cancelBtn"), { name: "cancel" });
    var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });
    dialog.cancelElement = cancelButton;
    dialog.defaultElement = okButton;

    okButton.onClick = function () {
        var columns = parseInt(columnInput.text, 10);
        var gapInRulerUnit = parseFloat(gapInput.text);
        if (isNaN(columns) || columns < 1) { alert(labelText("badColumns")); return; }
        if (isNaN(gapInRulerUnit) || gapInRulerUnit < 0) { alert(labelText("badGap")); return; }
        dialogResult = { columns: columns, gapInPoints: gapInRulerUnit * rulerUnit.pointsPerUnit };
        dialog.close();
    };

    dialog.show();
    return dialogResult;
}

// ============================================================
// メイン処理 / Main
// ============================================================
function restoreCoordinateSystem(savedCoordinateSystem) {
    if (savedCoordinateSystem !== null) {
        try { app.coordinateSystem = savedCoordinateSystem; } catch (e) { }
    }
}

function main() {
    if (app.documents.length === 0) { alert(labelText("noDoc")); return; }
    var document = app.activeDocument;
    var artboardCount = document.artboards.length;
    if (artboardCount < 2) { alert(labelText("needTwo")); return; }

    /* artboardRect と geometricBounds を同じ座標系に揃える / unify the coordinate space */
    var savedCoordinateSystem = null;
    try {
        savedCoordinateSystem = app.coordinateSystem;
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
    } catch (e) { }

    /* 元の矩形を保存し、最大サイズを 1 セルとする / capture original rects; cell = largest artboard */
    var originalArtboardRects = [], cellWidth = 0, cellHeight = 0, i, j;
    for (i = 0; i < artboardCount; i++) {
        var artboardRect = document.artboards[i].artboardRect; // [left, top, right, bottom]
        originalArtboardRects.push(artboardRect);
        var artboardWidth = artboardRect[2] - artboardRect[0];
        var artboardHeight = artboardRect[1] - artboardRect[3];
        if (artboardWidth > cellWidth) cellWidth = artboardWidth;
        if (artboardHeight > cellHeight) cellHeight = artboardHeight;
    }

    var rulerUnit = getRulerUnitInfo();

    var layoutSettings = showDialog(artboardCount, cellWidth, cellHeight, rulerUnit);
    if (!layoutSettings) { restoreCoordinateSystem(savedCoordinateSystem); return; }
    var gapInPoints = layoutSettings.gapInPoints;
    var columnCount = layoutSettings.columns;

    /* アートワークの所属アートボードを判定（移動前に）/ assign artwork to artboards (before any move) */
    var topLevelItems = collectTopLevelItems(document);
    var itemsByArtboard = [];
    for (i = 0; i < artboardCount; i++) itemsByArtboard.push([]);
    for (i = 0; i < topLevelItems.length; i++) {
        var itemCenter = getRectCenter(topLevelItems[i].geometricBounds); // 中心点で所属を決定 / assign by center point
        for (j = 0; j < artboardCount; j++) {
            var candidateRect = originalArtboardRects[j];
            if (itemCenter[0] >= candidateRect[0] && itemCenter[0] <= candidateRect[2]
                && itemCenter[1] <= candidateRect[1] && itemCenter[1] >= candidateRect[3]) {
                itemsByArtboard[j].push(topLevelItems[i]);
                break;
            }
        }
    }

    /* グリッド原点（カンバス中央）/ grid origin centered on the canvas */
    var canvasBounds = getCanvasBounds();
    var gridOrigin = computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, gapInPoints, columnCount, artboardCount);

    /* 再配置 / reposition */
    var lockState = unlockLayersAndItems(document);
    var failedMoveCount = 0;
    try {
        for (i = 0; i < artboardCount; i++) {
            var row = Math.floor(i / columnCount);
            var col = i % columnCount;
            var cellLeft = gridOrigin.left + col * (cellWidth + gapInPoints);
            var cellTop = gridOrigin.top - row * (cellHeight + gapInPoints);

            var originalRect = originalArtboardRects[i];
            var artboardWidth = originalRect[2] - originalRect[0];
            var artboardHeight = originalRect[1] - originalRect[3];
            /* サイズが異なるアートボードはセル内で中央に置く / center within the cell */
            var newLeft = cellLeft + (cellWidth - artboardWidth) / 2;
            var newTop = cellTop - (cellHeight - artboardHeight) / 2;

            var deltaX = newLeft - originalRect[0];
            var deltaY = newTop - originalRect[1];

            document.artboards[i].artboardRect = [newLeft, newTop, newLeft + artboardWidth, newTop - artboardHeight];

            var artboardItems = itemsByArtboard[i];
            for (j = 0; j < artboardItems.length; j++) {
                try { artboardItems[j].translate(deltaX, deltaY); } catch (e) { failedMoveCount++; }
            }
        }
    } finally {
        restoreLockState(lockState);
        restoreCoordinateSystem(savedCoordinateSystem);
    }

    app.redraw();

    var message = labelText("done") + "\n"
        + labelText("abCount") + ": " + artboardCount + "  /  "
        + columnCount + " " + labelText("colWord") + " x " + gridOrigin.rows + " " + labelText("rowWord");
    if (failedMoveCount > 0) message += "\n" + failedMoveCount + labelText("moveFailed");
    alert(message);
}

// ============================================================
// 実行 / Run
// ============================================================
try {
    main();
} catch (err) {
    alert("Error: " + err + (err.line ? " (line " + err.line + ")" : ""));
}
