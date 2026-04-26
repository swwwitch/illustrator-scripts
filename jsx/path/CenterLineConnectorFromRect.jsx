#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要 / Overview

- 選択した長方形を中心線に変換し、交差する線を連結／統合する Illustrator 用スクリプト
- 2本／3本／4本は幾何ロジック、5本以上は Pathfinder（Divide → Add）で処理
- 除外プレビュー、回転補正、線幅共通化に対応し、共通線幅は5本以上の統合結果にも反映

- Converts selected rectangles into center lines and connects / unifies intersecting lines in Illustrator.
- Uses geometric logic for 2 / 3 / 4 lines and Pathfinder (Divide → Add) for 5+ lines.
- Supports exclusion preview, rotation correction, and common stroke width, including 5+ line union results.

### 主な機能 / Main Features

- 縦長／横長の長方形に対応し、適切な方向の中心線を描画
- 回転補正（任意でON/OFF）で角度ズレを修正
- 複数オブジェクト同時処理に対応
- 除外条件（正方形に近い形状、短辺×1.5＞長辺、短辺が指定値未満）あり
- 日本語／英語インターフェース対応

- Supports vertical / horizontal rectangles, draws appropriate center line
- Optional rotation correction to fix angle misalignments
- Supports processing multiple objects at once
- Exclusion conditions (near-square, short × 1.5 > long, short side below threshold)
- Japanese / English UI support

### 処理の流れ / Process Flow

1. 閉じた4点パスの長方形を抽出 / Extract closed 4-point rectangle paths
2. 除外条件に該当する長方形をプレビュー表示 / Preview rectangles that match exclusion conditions
3. 縦横比に基づき長方形を中心線に変換 / Convert rectangles into center lines based on aspect ratio
4. 交差関係に応じて、4本・3本・2本・5本以上の中心線を連結／統合処理 / Convert 4 / 3 / 2 / 5+ intersecting center lines into connected or unified paths
5. 5本以上は Pathfinder（Divide → Add）でアウトライン統合し、塗りなし・線ありを再設定 / For 5+ lines, unify outlines via Pathfinder (Divide → Add) and restore no-fill / stroked appearance
6. 必要に応じて生成結果の線幅を共通化し、生成結果を選択状態に / Optionally normalize stroke widths and select generated results

### 更新履歴 / Update History

- v1.0.0 (20250612) : 初版作成 / Initial version
- v1.3.1 (20260427) : 除外条件のUI文言を整理し、印刷用の黒にするオプションを追加 / Refined exclusion UI wording and added an option to use print black

*/

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.3.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "長方形を線に変換し、連結処理", en: "Convert Rectangles to Lines and Connect" },
    chkExclude: { ja: "長方形を除外", en: "Exclude rectangles" },
    lblShortSideLength: { ja: "短辺が", en: "Short side is" },
    lblGreaterEqual: { ja: "mm以上", en: "mm or more" },
    pnlCenterLineConversion: { ja: "中心線化", en: "Center Line Conversion" },
    pnlOption: { ja: "オプション", en: "Options" },
    chkConnectAll: { ja: "連結処理", en: "Connect lines" },
    chkPrintBlack: { ja: "印刷用の黒にする", en: "Use print black" },
    lblStrokePref: { ja: "線幅", en: "Stroke Width" },
    chkCommonStroke: { ja: "線幅を共通にする", en: "Make stroke widths common" },
    strokeMax: { ja: "最大", en: "Max" },
    strokeMin: { ja: "最小", en: "Min" },
    strokeAvg: { ja: "平均", en: "Average" },
    btnOutlineOn: { ja: "アウトライン表示", en: "Outline View" },
    btnOutlineOff: { ja: "プレビュー表示", en: "Preview View" },
    btnOk: { ja: "OK", en: "OK" },
    btnCancel: { ja: "キャンセル", en: "Cancel" },
    alertNoSelection: { ja: "長方形を1つ以上選択してください。", en: "Please select at least one rectangle." },
    alertError: { ja: "エラーが発生しました", en: "An error occurred" }
};

/* ラベル取得 / Get label */
function L(key) {
    return LABELS[key][lang];
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return L(key) + (lang === 'ja' ? '：' : ':');
}

// =========================================
// 対象外マーカー（プレビュー）/ Exclusion markers (preview)
// =========================================

var EXCLUSION_PREVIEW_LAYER_NAME = "__center_line_preview__";


function makeMagentaYellow() {
    var markerColor = new CMYKColor();
    markerColor.cyan = 0; markerColor.magenta = 100; markerColor.yellow = 100; markerColor.black = 0;
    return markerColor;
}

function makePrintBlack() {
    var blackColor = new CMYKColor();
    blackColor.cyan = 0;
    blackColor.magenta = 0;
    blackColor.yellow = 0;
    blackColor.black = 100;
    return blackColor;
}

function ensureExclusionPreviewLayer() {
    var doc = app.activeDocument;
    var layer;
    try {
        layer = doc.layers.getByName(EXCLUSION_PREVIEW_LAYER_NAME);
    } catch (e) {
        layer = doc.layers.add();
        layer.name = EXCLUSION_PREVIEW_LAYER_NAME;
    }
    layer.locked = false;
    layer.visible = true;
    return layer;
}

function clearExclusionPreviewLayer() {
    try {
        var layer = app.activeDocument.layers.getByName(EXCLUSION_PREVIEW_LAYER_NAME);
        layer.remove();
    } catch (e) { }
}

/* グループを再帰的に展開して4点閉じ長方形を収集
   Recursively walk groups and collect 4-point closed rectangles */
function collectRectsRecursive(items, results) {
    for (var i = 0; i < items.length; i++) {
        var pageItem = items[i];
        try { if (pageItem.locked || pageItem.hidden) continue; } catch (e) { }
        if (pageItem.typename === "PathItem") {
            if (pageItem.closed && pageItem.pathPoints.length === 4) results.push(pageItem);
        } else if (pageItem.typename === "CompoundPathItem") {
            if (pageItem.pathItems.length === 1) {
                var compoundSubPath = pageItem.pathItems[0];
                if (compoundSubPath.closed && compoundSubPath.pathPoints.length === 4) results.push(compoundSubPath);
            }
        } else if (pageItem.typename === "GroupItem") {
            collectRectsRecursive(pageItem.pageItems, results);
        }
    }
}

/* convertRectToCenterLine と同じ除外条件を判定 / Same exclusion criteria as convertRectToCenterLine */

function isExcludedRect(rect, minShortSidePt) {
    var bounds = rect.geometricBounds;
    var rectWidth = bounds[2] - bounds[0];
    var rectHeight = bounds[1] - bounds[3];
    var diffRatio = Math.abs(rectWidth - rectHeight) / Math.max(rectWidth, rectHeight);
    if (diffRatio < 0.05) return true;
    var shortSide = Math.min(rectWidth, rectHeight);
    var longSide = Math.max(rectWidth, rectHeight);
    if (shortSide * 1.5 > longSide) return true;
    if (shortSide < minShortSidePt) return true;
    return false;
}

/* 選択中の長方形から、短辺の最大値を mm で取得
   Get the largest short side among selected rectangles in millimeters */
function getMaxShortSideMmFromItems(items) {
    var rects = [];
    collectRectsRecursive(items, rects);
    var maxShortSidePt = 0;

    for (var rectIndex = 0; rectIndex < rects.length; rectIndex++) {
        try {
            var bounds = rects[rectIndex].geometricBounds;
            var rectWidth = Math.abs(bounds[2] - bounds[0]);
            var rectHeight = Math.abs(bounds[1] - bounds[3]);
            var shortSide = Math.min(rectWidth, rectHeight);
            if (shortSide > maxShortSidePt) maxShortSidePt = shortSide;
        } catch (e) { }
    }

    return Math.round((maxShortSidePt * 25.4 / 72) * 10) / 10;
}

/* 対象外オブジェクトを M100Y100・半透明で複製してマーカー表示
   Duplicate excluded objects as M100Y100 semi-transparent markers */
function refreshExclusionPreview(items, minShortSidePt) {
    clearExclusionPreviewLayer();
    if (!items || items.length === 0) { app.redraw(); return; }
    var layer = ensureExclusionPreviewLayer();
    var color = makeMagentaYellow();

    var rects = [];
    collectRectsRecursive(items, rects);

    for (var i = 0; i < rects.length; i++) {
        if (!isExcludedRect(rects[i], minShortSidePt)) continue;
        try {
            var markerRect = rects[i].duplicate(layer, ElementPlacement.PLACEATBEGINNING);
            markerRect.filled = true;
            markerRect.fillColor = color;
            markerRect.stroked = false;
            markerRect.opacity = 50;
        } catch (e) { }
    }
    app.redraw();
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* オプション設定用ダイアログを表示 / Show options dialog */
function showOptionDialog() {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "row";
    dlg.alignChildren = ["fill", "fill"];
    dlg.margins = 16;
    dlg.spacing = 12;

    var PANEL_MARGINS = [15, 20, 15, 10];

    /* 左カラム（パネル群）/ Left column (panels) */
    var leftCol = dlg.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "fill";
    leftCol.spacing = 12;

    /* 直線化パネル / Straighten panel */
    var centerLinePanel = leftCol.add("panel", undefined, L('pnlCenterLineConversion'));
    centerLinePanel.alignChildren = "left";
    centerLinePanel.margins = PANEL_MARGINS;
    centerLinePanel.spacing = 6;


    /* 短辺の下限（mm）：スライダー上限は選択中の長方形の短辺最大値
       Min short side (mm): slider maximum = largest short side among selected rectangles */
    var SHORT_MIN = 0;
    var SHORT_MAX = getMaxShortSideMmFromItems(app.activeDocument.selection);
    if (SHORT_MAX <= 0) SHORT_MAX = 10;
    var SHORT_DEFAULT = Math.min(10, SHORT_MAX);

    var cbMinShortSide = centerLinePanel.add("checkbox", undefined, L('chkExclude'));
    cbMinShortSide.value = false;

    var minShortSideGroup = centerLinePanel.add("group");
    minShortSideGroup.alignment = "left";
    minShortSideGroup.add("statictext", undefined, L('lblShortSideLength'));
    var minShortSideInput = minShortSideGroup.add("edittext", undefined, SHORT_DEFAULT.toFixed(1));
    minShortSideInput.characters = 5;
    minShortSideGroup.add("statictext", undefined, L('lblGreaterEqual'));

    var minShortSideSlider = centerLinePanel.add("slider", undefined, SHORT_DEFAULT, SHORT_MIN, SHORT_MAX);
    minShortSideSlider.preferredSize.width = 190;

    function syncShortEnabled() {
        var on = cbMinShortSide.value;
        minShortSideInput.enabled = on;
        minShortSideSlider.enabled = on;
    }
    cbMinShortSide.onClick = function () {
        syncShortEnabled();
        updateExclusionPreview();
    };
    syncShortEnabled();

    function roundTo1(v) { return Math.round(v * 10) / 10; }

    /* プレビュー用に元の選択を保存 / Snapshot original selection for preview */
    var snapshotSelection = [];
    for (var ssi = 0; ssi < app.activeDocument.selection.length; ssi++) {
        snapshotSelection.push(app.activeDocument.selection[ssi]);
    }

    function updateExclusionPreview() {
        var minShortSidePt = 0;
        if (cbMinShortSide.value) {
            var minShortSideMm = parseFloat(minShortSideInput.text);
            if (isNaN(minShortSideMm)) minShortSideMm = SHORT_DEFAULT;
            minShortSidePt = minShortSideMm * 72 / 25.4;
        }
        refreshExclusionPreview(snapshotSelection, minShortSidePt);
    }

    minShortSideSlider.onChanging = function () {
        minShortSideInput.text = roundTo1(minShortSideSlider.value).toFixed(1);
        updateExclusionPreview();
    };
    minShortSideInput.onChange = function () {
        var minShortSideMm = parseFloat(minShortSideInput.text);
        if (isNaN(minShortSideMm)) minShortSideMm = minShortSideSlider.value;
        if (minShortSideMm < SHORT_MIN) minShortSideMm = SHORT_MIN;
        if (minShortSideMm > SHORT_MAX) minShortSideMm = SHORT_MAX;
        minShortSideSlider.value = minShortSideMm;
        minShortSideInput.text = roundTo1(minShortSideMm).toFixed(1);
        updateExclusionPreview();
    };

    /* オプションパネル（1カラム）/ Options panel (single column) */
    var pnlOpt = leftCol.add("panel", undefined, L('pnlOption'));
    pnlOpt.orientation = "column";
    pnlOpt.alignChildren = "fill";
    pnlOpt.margins = PANEL_MARGINS;
    pnlOpt.spacing = 8;

    /* 連結処理チェックボックス / Connect lines checkbox */
    var cbConnectAll = pnlOpt.add("checkbox", undefined, L('chkConnectAll'));
    cbConnectAll.value = true;

    var cbPrintBlack = pnlOpt.add("checkbox", undefined, L('chkPrintBlack'));
    cbPrintBlack.value = false;

    /* 線幅の決定パネル / Stroke width panel */
    var pnlStroke = pnlOpt.add("panel", undefined, L('lblStrokePref'));
    pnlStroke.orientation = "column";
    pnlStroke.alignChildren = "left";
    pnlStroke.margins = PANEL_MARGINS;
    pnlStroke.spacing = 6;
    var strokeRadioGroup = pnlStroke.add("group");
    strokeRadioGroup.orientation = "row";
    strokeRadioGroup.alignChildren = "left";
    strokeRadioGroup.spacing = 8;
    var rbStrokeMax = strokeRadioGroup.add("radiobutton", undefined, L('strokeMax'));
    var rbStrokeMin = strokeRadioGroup.add("radiobutton", undefined, L('strokeMin'));
    var rbStrokeAvg = strokeRadioGroup.add("radiobutton", undefined, L('strokeAvg'));
    rbStrokeMin.value = true;
    var cbCommonStroke = pnlStroke.add("checkbox", undefined, L('chkCommonStroke'));
    cbCommonStroke.value = false;

    /* 右カラム（上：OK／キャンセル、下：アウトライン表示）
       Right column (top: OK / Cancel, bottom: outline toggle) */
    var rightCol = dlg.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];
    rightCol.alignment = ["right", "fill"];
    rightCol.spacing = 8;

    var btnOk = rightCol.add("button", undefined, L('btnOk'), { name: "ok" });

    var btnCancel = rightCol.add("button", undefined, L('btnCancel'), { name: "cancel" });

    var verticalSpacer = rightCol.add("statictext", undefined, "");
    verticalSpacer.alignment = ["fill", "fill"];

    var outlineGroup = rightCol.add("group");
    outlineGroup.orientation = "column";
    outlineGroup.alignChildren = ["fill", "bottom"];
    var isOutlineMode = false;
    var btnOutlineToggle = outlineGroup.add("button", undefined, L('btnOutlineOn'));
    btnOutlineToggle.onClick = function () {
        try {
            app.executeMenuCommand('preview');
            isOutlineMode = !isOutlineMode;
            btnOutlineToggle.text = isOutlineMode ? L('btnOutlineOff') : L('btnOutlineOn');
        } catch (e) { }
    };

    /* 初回の除外マーカー表示 / Initial exclusion marker preview */
    updateExclusionPreview();

    var dlgResult = dlg.show();

    /* プレビューレイヤー削除＆選択復元 / Clean up preview layer and restore selection */
    clearExclusionPreviewLayer();
    try { app.activeDocument.selection = snapshotSelection; } catch (e) { }
    app.redraw();

    if (dlgResult !== 1) return null;

    var strokeStrategy = rbStrokeMin.value ? "min" : (rbStrokeAvg.value ? "avg" : "max");

    return {
        correctRotation: true,
        minShortSideMm: cbMinShortSide.value ? parseFloat(minShortSideInput.text) : 0,
        connectLines: cbConnectAll.value,
        printBlack: cbPrintBlack.value,
        commonStroke: cbCommonStroke.value,
        strokeStrategy: strokeStrategy
    };
}

// =========================================
// 中心線描画処理 / Center Line Drawing
// =========================================

/* 長方形を中心線に変換し、元の長方形を削除 / Convert rectangle into a center line and remove original */
function convertRectToCenterLine(rect, correctRotation, minShortSidePt) {
    /* 回転補正（最初の2点のアンカーから角度を取得）/ Rotation correction (angle from first two anchors) */
    if (correctRotation && rect.pathPoints.length === 4) {
        var firstAnchor = rect.pathPoints[0].anchor;
        var secondAnchor = rect.pathPoints[1].anchor;

        var deltaX = secondAnchor[0] - firstAnchor[0];
        var deltaY = secondAnchor[1] - firstAnchor[1];
        var angleRad = Math.atan2(deltaY, deltaX);
        var angleDeg = angleRad * 180 / Math.PI;
        if (angleDeg < 0) angleDeg += 360;

        var rotationAmount = angleDeg;
        var normalized = rotationAmount % 90;
        if (normalized > 45) normalized = 90 - normalized;

        if (normalized >= 0.5 && normalized <= 10) {
            rect.rotate(-rotationAmount);
        }
    }

    /* 補正後にバウンディングボックスを取得 / Get bounding box after correction */
    var bounds = rect.geometricBounds;
    var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
    var rectWidth = right - left;
    var rectHeight = top - bottom;

    /* 正方形に近い場合は除外（5%未満の差）/ Skip near-square shapes (<5% diff) */
    var diffRatio = Math.abs(rectWidth - rectHeight) / Math.max(rectWidth, rectHeight);
    if (diffRatio < 0.05) return null;
    /* 短辺×1.5 ＞ 長辺の場合は除外 / Skip when short × 1.5 > long */
    var shortSide = Math.min(rectWidth, rectHeight);
    var longSide = Math.max(rectWidth, rectHeight);
    if (shortSide * 1.5 > longSide) return null;
    /* 短辺が指定値未満の場合は除外（指定値以降のみ直線化） / Skip when short side is below the threshold (only linearize at/above) */
    if (typeof minShortSidePt === "number" && shortSide < minShortSidePt) return null;

    /* 中心線を生成 / Create center line */
    var centerLine = app.activeDocument.pathItems.add();
    centerLine.stroked = true;
    centerLine.filled = false;
    centerLine.strokeColor = rect.fillColor;

    var startPoint = centerLine.pathPoints.add();
    var endPoint = centerLine.pathPoints.add();

    if (rectHeight <= rectWidth) {
        /* 横長：中央に水平線 / Horizontal: draw horizontal line at center */
        var centerY = (top + bottom) / 2;
        startPoint.anchor = [left, centerY];
        endPoint.anchor = [right, centerY];
    } else {
        /* 縦長：中央に垂直線 / Vertical: draw vertical line at center */
        var centerX = (left + right) / 2;
        startPoint.anchor = [centerX, top];
        endPoint.anchor = [centerX, bottom];
    }

    centerLine.strokeWidth = (rectHeight <= rectWidth) ? rectHeight : rectWidth;

    startPoint.leftDirection = startPoint.anchor;
    startPoint.rightDirection = startPoint.anchor;
    endPoint.leftDirection = endPoint.anchor;
    endPoint.rightDirection = endPoint.anchor;

    rect.remove();
    return centerLine;
}

/* 太さが混在するときの代表値を取得 / Pick a representative stroke width when widths differ */
function getRepresentativeStrokeWidth(lines, strategy) {
    var initialStrokeWidth = lines[0].strokeWidth;
    if (strategy === "min") {
        var minimumStrokeWidth = initialStrokeWidth;
        for (var i = 1; i < lines.length; i++) if (lines[i].strokeWidth < minimumStrokeWidth) minimumStrokeWidth = lines[i].strokeWidth;
        return minimumStrokeWidth;
    } else if (strategy === "avg") {
        var sum = 0;
        for (var j = 0; j < lines.length; j++) sum += lines[j].strokeWidth;
        return sum / lines.length;
    }
    /* default: max */
    var maximumStrokeWidth = initialStrokeWidth;
    for (var k = 1; k < lines.length; k++) if (lines[k].strokeWidth > maximumStrokeWidth) maximumStrokeWidth = lines[k].strokeWidth;
    return maximumStrokeWidth;
}

/* 生成結果の線幅を共通化（コンパウンド／グループ内のサブパスも対象）
   Make stroke width common across generated results (recurses into compounds and groups) */

function applyCommonStrokeWidth(items, strategy) {
    if (!items || items.length === 0) return;

    var strokeItems = [];
    function collectStrokedRecursive(item) {
        try {
            if (item.typename === 'PathItem') {
                if (item.stroked) strokeItems.push(item);
            } else if (item.typename === 'CompoundPathItem') {
                for (var n = 0; n < item.pathItems.length; n++) collectStrokedRecursive(item.pathItems[n]);
            } else if (item.typename === 'GroupItem') {
                for (var m = 0; m < item.pageItems.length; m++) collectStrokedRecursive(item.pageItems[m]);
            }
        } catch (e) { }
    }
    for (var i = 0; i < items.length; i++) {
        if (items[i]) collectStrokedRecursive(items[i]);
    }
    if (strokeItems.length === 0) return;

    var commonWidth = getRepresentativeStrokeWidth(strokeItems, strategy);
    for (var j = 0; j < strokeItems.length; j++) {
        try { strokeItems[j].strokeWidth = commonWidth; } catch (e) { }
    }
}

/* 生成結果の線色を印刷用の黒（C0 M0 Y0 K100）に統一
   Apply print black (C0 M0 Y0 K100) to generated result strokes */
function applyPrintBlackStroke(items) {
    if (!items || items.length === 0) return;

    var blackColor = makePrintBlack();

    function applyBlackRecursive(item) {
        try {
            if (item.typename === 'PathItem') {
                item.stroked = true;
                item.strokeColor = blackColor;
            } else if (item.typename === 'CompoundPathItem') {
                for (var compoundPathIndex = 0; compoundPathIndex < item.pathItems.length; compoundPathIndex++) {
                    applyBlackRecursive(item.pathItems[compoundPathIndex]);
                }
            } else if (item.typename === 'GroupItem') {
                for (var groupItemIndex = 0; groupItemIndex < item.pageItems.length; groupItemIndex++) {
                    applyBlackRecursive(item.pageItems[groupItemIndex]);
                }
            }
        } catch (e) { }
    }

    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        if (items[itemIndex]) applyBlackRecursive(items[itemIndex]);
    }
}

/* 4本の中心線が「#」状に交差していれば、交点を頂点とする長方形に変換
   If 4 center lines form a "#" shape, convert into a rectangle whose corners are the intersections */
function tryConnectFourLinesIntoRect(lines, strokeStrategy) {
    if (!lines || lines.length !== 4) return null;

    var EPS = 0.01;
    var horizontalLineInfos = [];
    var verticalLineInfos = [];

    for (var i = 0; i < 4; i++) {
        var lineItem = lines[i];
        if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) return null;
        var startAnchor = lineItem.pathPoints[0].anchor;
        var endAnchor = lineItem.pathPoints[1].anchor;
        if (Math.abs(startAnchor[1] - endAnchor[1]) < EPS) {
            horizontalLineInfos.push({
                y: (startAnchor[1] + endAnchor[1]) / 2,
                xMin: Math.min(startAnchor[0], endAnchor[0]),
                xMax: Math.max(startAnchor[0], endAnchor[0]),
                line: lineItem
            });
        } else if (Math.abs(startAnchor[0] - endAnchor[0]) < EPS) {
            verticalLineInfos.push({
                x: (startAnchor[0] + endAnchor[0]) / 2,
                yMin: Math.min(startAnchor[1], endAnchor[1]),
                yMax: Math.max(startAnchor[1], endAnchor[1]),
                line: lineItem
            });
        } else {
            return null;
        }
    }

    if (horizontalLineInfos.length !== 2 || verticalLineInfos.length !== 2) return null;

    /* 各 H が両 V の x を内包し、各 V が両 H の y を内包しているか
       Each H must span across both Vs in x; each V across both Hs in y */
    var firstVerticalX = verticalLineInfos[0].x;
    var secondVerticalX = verticalLineInfos[1].x;
    var firstHorizontalY = horizontalLineInfos[0].y;
    var secondHorizontalY = horizontalLineInfos[1].y;
    var xMin = Math.min(firstVerticalX, secondVerticalX);
    var xMax = Math.max(firstVerticalX, secondVerticalX);
    var yMin = Math.min(firstHorizontalY, secondHorizontalY);
    var yMax = Math.max(firstHorizontalY, secondHorizontalY);

    for (var horizontalIndex = 0; horizontalIndex < 2; horizontalIndex++) {
        if (horizontalLineInfos[horizontalIndex].xMin > xMin + EPS || horizontalLineInfos[horizontalIndex].xMax < xMax - EPS) return null;
    }
    for (var verticalIndex = 0; verticalIndex < 2; verticalIndex++) {
        if (verticalLineInfos[verticalIndex].yMin > yMin + EPS || verticalLineInfos[verticalIndex].yMax < yMax - EPS) return null;
    }

    /* 交点を頂点とする閉じた長方形を生成 / Create closed rectangle through intersections */
    var rectanglePath = app.activeDocument.pathItems.add();
    rectanglePath.closed = true;
    rectanglePath.stroked = lines[0].stroked;
    rectanglePath.strokeColor = lines[0].strokeColor;
    rectanglePath.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy);
    rectanglePath.filled = false;

    var corners = [
        [xMin, yMax],
        [xMax, yMax],
        [xMax, yMin],
        [xMin, yMin]
    ];
    for (var c = 0; c < 4; c++) {
        var cornerPoint = rectanglePath.pathPoints.add();
        cornerPoint.anchor = corners[c];
        cornerPoint.leftDirection = corners[c];
        cornerPoint.rightDirection = corners[c];
    }

    for (var k = 0; k < lines.length; k++) {
        try { lines[k].remove(); } catch (e) { }
    }

    return rectanglePath;
}

/* 3本（2H+1V または 1H+2V）の中心線をコの字型の開いたパスに変換
   Convert 3 lines (2H+1V or 1H+2V) into an open U-shape path */
function tryConnectThreeLinesIntoUShape(lines, strokeStrategy) {
    if (!lines || lines.length !== 3) return null;

    var EPS = 0.01;
    var horizontalLineInfos = [], verticalLineInfos = [];

    for (var i = 0; i < 3; i++) {
        var lineItem = lines[i];
        if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) return null;
        var startAnchor = lineItem.pathPoints[0].anchor;
        var endAnchor = lineItem.pathPoints[1].anchor;
        if (Math.abs(startAnchor[1] - endAnchor[1]) < EPS) {
            horizontalLineInfos.push({ y: (startAnchor[1] + endAnchor[1]) / 2, xMin: Math.min(startAnchor[0], endAnchor[0]), xMax: Math.max(startAnchor[0], endAnchor[0]) });
        } else if (Math.abs(startAnchor[0] - endAnchor[0]) < EPS) {
            verticalLineInfos.push({ x: (startAnchor[0] + endAnchor[0]) / 2, yMin: Math.min(startAnchor[1], endAnchor[1]), yMax: Math.max(startAnchor[1], endAnchor[1]) });
        } else {
            return null;
        }
    }

    var anchors = null;

    if (horizontalLineInfos.length === 2 && verticalLineInfos.length === 1) {
        /* 2H + 1V：V は両 H に交差し、両 H の x 範囲の片端付近にある
           V crosses both Hs and sits near one extreme of the H x-range */
        var verticalLineInfo = verticalLineInfos[0];
        var firstHorizontalY = horizontalLineInfos[0].y;
        var secondHorizontalY = horizontalLineInfos[1].y;
        var yMin = Math.min(firstHorizontalY, secondHorizontalY);
        var yMax = Math.max(firstHorizontalY, secondHorizontalY);
        if (verticalLineInfo.yMin > yMin + EPS || verticalLineInfo.yMax < yMax - EPS) return null;
        for (var horizontalIndex = 0; horizontalIndex < 2; horizontalIndex++) {
            if (horizontalLineInfos[horizontalIndex].xMin > verticalLineInfo.x + EPS || horizontalLineInfos[horizontalIndex].xMax < verticalLineInfo.x - EPS) return null;
        }
        var horizontalXMin = Math.min(horizontalLineInfos[0].xMin, horizontalLineInfos[1].xMin);
        var horizontalXMax = Math.max(horizontalLineInfos[0].xMax, horizontalLineInfos[1].xMax);
        var openEndX = (Math.abs(verticalLineInfo.x - horizontalXMin) < Math.abs(verticalLineInfo.x - horizontalXMax)) ? horizontalXMax : horizontalXMin;
        anchors = [
            [openEndX, firstHorizontalY],
            [verticalLineInfo.x, firstHorizontalY],
            [verticalLineInfo.x, secondHorizontalY],
            [openEndX, secondHorizontalY]
        ];
    } else if (horizontalLineInfos.length === 1 && verticalLineInfos.length === 2) {
        /* 1H + 2V：H は両 V に交差し、両 V の y 範囲の片端付近にある
           H crosses both Vs and sits near one extreme of the V y-range */
        var horizontalLineInfo = horizontalLineInfos[0];
        var firstVerticalX = verticalLineInfos[0].x;
        var secondVerticalX = verticalLineInfos[1].x;
        var xMin = Math.min(firstVerticalX, secondVerticalX);
        var xMax = Math.max(firstVerticalX, secondVerticalX);
        if (horizontalLineInfo.xMin > xMin + EPS || horizontalLineInfo.xMax < xMax - EPS) return null;
        for (var verticalIndex = 0; verticalIndex < 2; verticalIndex++) {
            if (verticalLineInfos[verticalIndex].yMin > horizontalLineInfo.y + EPS || verticalLineInfos[verticalIndex].yMax < horizontalLineInfo.y - EPS) return null;
        }
        var verticalYMin = Math.min(verticalLineInfos[0].yMin, verticalLineInfos[1].yMin);
        var verticalYMax = Math.max(verticalLineInfos[0].yMax, verticalLineInfos[1].yMax);
        var openEndY = (Math.abs(horizontalLineInfo.y - verticalYMin) < Math.abs(horizontalLineInfo.y - verticalYMax)) ? verticalYMax : verticalYMin;
        anchors = [
            [firstVerticalX, openEndY],
            [firstVerticalX, horizontalLineInfo.y],
            [secondVerticalX, horizontalLineInfo.y],
            [secondVerticalX, openEndY]
        ];
    } else {
        return null;
    }

    var path = app.activeDocument.pathItems.add();
    path.closed = false;
    path.stroked = lines[0].stroked;
    path.strokeColor = lines[0].strokeColor;
    path.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy);
    path.filled = false;

    for (var anchorIndex = 0; anchorIndex < anchors.length; anchorIndex++) {
        var pathPoint = path.pathPoints.add();
        pathPoint.anchor = anchors[anchorIndex];
        pathPoint.leftDirection = anchors[anchorIndex];
        pathPoint.rightDirection = anchors[anchorIndex];
    }

    for (var removeIndex = 0; removeIndex < lines.length; removeIndex++) {
        try { lines[removeIndex].remove(); } catch (e) { }
    }

    return path;
}

/* 2本（1H+1V）の中心線をL字型の開いたパスに変換
   Convert 2 lines (1H+1V) into an open L-shape path */
function tryConnectTwoLinesIntoLShape(lines, strokeStrategy) {
    if (!lines || lines.length !== 2) return null;

    var EPS = 0.01;
    var horizontalLineInfos = [], verticalLineInfos = [];

    for (var i = 0; i < 2; i++) {
        var lineItem = lines[i];
        if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) return null;
        var startAnchor = lineItem.pathPoints[0].anchor;
        var endAnchor = lineItem.pathPoints[1].anchor;
        if (Math.abs(startAnchor[1] - endAnchor[1]) < EPS) {
            horizontalLineInfos.push({ y: (startAnchor[1] + endAnchor[1]) / 2, xMin: Math.min(startAnchor[0], endAnchor[0]), xMax: Math.max(startAnchor[0], endAnchor[0]) });
        } else if (Math.abs(startAnchor[0] - endAnchor[0]) < EPS) {
            verticalLineInfos.push({ x: (startAnchor[0] + endAnchor[0]) / 2, yMin: Math.min(startAnchor[1], endAnchor[1]), yMax: Math.max(startAnchor[1], endAnchor[1]) });
        } else {
            return null;
        }
    }

    if (horizontalLineInfos.length !== 1 || verticalLineInfos.length !== 1) return null;

    var horizontalLineInfo = horizontalLineInfos[0];
    var verticalLineInfo = verticalLineInfos[0];

    /* 交差していること / Must intersect */
    if (horizontalLineInfo.xMin > verticalLineInfo.x + EPS || horizontalLineInfo.xMax < verticalLineInfo.x - EPS) return null;
    if (verticalLineInfo.yMin > horizontalLineInfo.y + EPS || verticalLineInfo.yMax < horizontalLineInfo.y - EPS) return null;

    var farEndY = (Math.abs(horizontalLineInfo.y - verticalLineInfo.yMin) < Math.abs(horizontalLineInfo.y - verticalLineInfo.yMax)) ? verticalLineInfo.yMax : verticalLineInfo.yMin;
    var farEndX = (Math.abs(verticalLineInfo.x - horizontalLineInfo.xMin) < Math.abs(verticalLineInfo.x - horizontalLineInfo.xMax)) ? horizontalLineInfo.xMax : horizontalLineInfo.xMin;

    var anchors = [
        [verticalLineInfo.x, farEndY],
        [verticalLineInfo.x, horizontalLineInfo.y],
        [farEndX, horizontalLineInfo.y]
    ];

    var path = app.activeDocument.pathItems.add();
    path.closed = false;
    path.stroked = lines[0].stroked;
    path.strokeColor = lines[0].strokeColor;
    path.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy);
    path.filled = false;

    for (var anchorIndex = 0; anchorIndex < anchors.length; anchorIndex++) {
        var pathPoint = path.pathPoints.add();
        pathPoint.anchor = anchors[anchorIndex];
        pathPoint.leftDirection = anchors[anchorIndex];
        pathPoint.rightDirection = anchors[anchorIndex];
    }

    for (var removeIndex = 0; removeIndex < lines.length; removeIndex++) {
        try { lines[removeIndex].remove(); } catch (e) { }
    }

    return path;
}

/* Pathfinder Divide の Live Effect を XML で適用するヘルパー
   Helper to apply a Pathfinder Divide Live Effect via XML */
function applyPathfinderDivideEffect(item, removeUnpainted, expandAppearance) {
    var shouldRemoveUnpainted = (removeUnpainted !== false);
    var shouldExpandAppearance = (expandAppearance === true);
    var values = [
        5,                         /* Command: Divide */
        1,                         /* ConvertCustom */
        shouldRemoveUnpainted ? 1 : 0,
        0.5,                       /* Mix */
        10,                        /* Precision */
        1,                         /* RemovePoints */
        'Divide'
    ];
    var xml = ('<LiveEffect name="Adobe Pathfinder" isPre="1"><Dict data="I Command #1 B ConvertCustom #2 B ExtractUnpainted #3 R Mix #4 R Precision #5 B RemovePoints #6"><Entry name="DisplayString" value="#7" valueType="S"/></Dict></LiveEffect>')
        .replace(/#(\d+)/g, function (_, n) { return values[parseInt(n, 10) - 1]; });

    item.applyEffect(xml);
    if (shouldExpandAppearance) app.executeMenuCommand("expandStyle");
}

/* PageItem 内のすべての PathItem に再帰的に処理を適用
   Recursively apply a callback to every PathItem inside a PageItem */
function walkPathItemsRecursive(item, callback) {
    try {
        if (item.typename === 'PathItem') {
            callback(item);
        } else if (item.typename === 'CompoundPathItem') {
            for (var compoundPathIndex = 0; compoundPathIndex < item.pathItems.length; compoundPathIndex++) {
                callback(item.pathItems[compoundPathIndex]);
            }
        } else if (item.typename === 'GroupItem') {
            for (var groupItemIndex = 0; groupItemIndex < item.pageItems.length; groupItemIndex++) {
                walkPathItemsRecursive(item.pageItems[groupItemIndex], callback);
            }
        }
    } catch (e) { }
}

/* 現在の選択内のすべてのパスを塗りなし・線ありに整える
   Restore no-fill / stroked appearance for all paths in the current selection */
function restoreNoFillStrokeToCurrentSelection(strokeColor, strokeWidth) {
    for (var selectionIndex = 0; selectionIndex < app.selection.length; selectionIndex++) {
        walkPathItemsRecursive(app.selection[selectionIndex], function (pathItem) {
            pathItem.filled = false;
            pathItem.stroked = true;
            if (strokeColor) {
                try { pathItem.strokeColor = strokeColor; } catch (e) { }
            }
            if (strokeWidth !== null && strokeWidth !== undefined) {
                pathItem.strokeWidth = strokeWidth;
            }
        });
    }
}

/* 選択中のアピアランスをアウトライン統合し、塗りなし・線ありに戻す
   Union the selected appearance into outlines, then restore no-fill / stroked paths */
function outlineUnionCurrentSelection(strokeColor, strokeWidth) {
    app.executeMenuCommand('Adobe New Fill Shortcut');
    app.executeMenuCommand('expandStyle');
    app.executeMenuCommand('Live Pathfinder Add');
    app.executeMenuCommand('expandStyle');

    restoreNoFillStrokeToCurrentSelection(strokeColor, strokeWidth);
}

/* 5本以上の中心線をグループ化し、Pathfinder Divide で分割後、
   新規塗り・展開・Pathfinder Add・再展開を経てアウトラインに統合する。
   展開後の選択結果に対して、塗りなし・線ありを再設定する。
   Group 5+ center lines, split them with Pathfinder Divide,
   then unify the outline through new fill, expand, Pathfinder Add, and expand again.
   Restore no-fill / stroked appearance on the expanded selection result. */
function tryConnectManyLinesIntoOutline(lines, strokeStrategy, commonStrokeWidth) {
    if (!lines || lines.length < 5) return null;

    var doc = app.activeDocument;
    var representativeStrokeWidth = (commonStrokeWidth !== null && commonStrokeWidth !== undefined)
        ? commonStrokeWidth
        : getRepresentativeStrokeWidth(lines, strokeStrategy);
    var representativeStrokeColor = lines[0].strokeColor;

    /* グループの配置先（最初の線の親レイヤー or 親グループ）/ Determine parent for the wrapping group */
    var groupParent = doc.activeLayer;
    try {
        var firstParent = lines[0].parent;
        if (firstParent && firstParent.typename === "CompoundPathItem") firstParent = firstParent.parent;
        if (firstParent && (firstParent.typename === "Layer" || firstParent.typename === "GroupItem")) {
            groupParent = firstParent;
        }
    } catch (e) { }

    /* 1. 線をすべて1つのグループに集める / Move all lines into a single group */
    var group;
    try { group = groupParent.groupItems.add(); }
    catch (e) { group = doc.activeLayer.groupItems.add(); }
    for (var i = lines.length - 1; i >= 0; i--) {
        try { lines[i].move(group, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
    }

    /* 2. Live Effect の Pathfinder Divide を適用（交差点で分割）
       Apply Pathfinder Divide as a Live Effect to split at intersections */
    applyPathfinderDivideEffect(group, false, false);

    /* 3. グループを選択状態にする / Make the group the current selection */
    app.selection = null;
    group.selected = true;
    app.redraw();

    /* 4. アウトライン統合し、現在の選択結果に対して塗りなし・線ありを再設定
       Union outlines, then clear fills and restore strokes on the current selection result */
    outlineUnionCurrentSelection(representativeStrokeColor, representativeStrokeWidth);

    /* 5. グループ内が単一要素なら親に昇格、複数要素なら包むグループのまま返す
       If the group contains a single child, promote it; otherwise return the group */
    var result = group;
    try {
        if (group.pageItems.length === 1) {
            result = group.pageItems[0];
            try { result.move(groupParent, ElementPlacement.PLACEATEND); } catch (e) { }
            try { group.remove(); } catch (e) { }
        }
    } catch (e) { }

    return result;
}

/* 軸並行な2線分が交差するか判定（端点接触も許容）
   Test whether 2 axis-aligned segments intersect (endpoint touch allowed) */
function intersectsAxisAligned(lnA, lnB) {
    var EPS = 0.01;
    var firstLineStartAnchor = lnA.pathPoints[0].anchor;
    var firstLineEndAnchor = lnA.pathPoints[1].anchor;
    var secondLineStartAnchor = lnB.pathPoints[0].anchor;
    var secondLineEndAnchor = lnB.pathPoints[1].anchor;
    var isFirstLineHorizontal = Math.abs(firstLineStartAnchor[1] - firstLineEndAnchor[1]) < EPS;
    var isSecondLineHorizontal = Math.abs(secondLineStartAnchor[1] - secondLineEndAnchor[1]) < EPS;
    if (isFirstLineHorizontal === isSecondLineHorizontal) return false; /* 平行は非交差 / Parallel: not intersecting */
    var horizontalStartAnchor = isFirstLineHorizontal ? firstLineStartAnchor : secondLineStartAnchor;
    var horizontalEndAnchor = isFirstLineHorizontal ? firstLineEndAnchor : secondLineEndAnchor;
    var verticalStartAnchor = isFirstLineHorizontal ? secondLineStartAnchor : firstLineStartAnchor;
    var verticalEndAnchor = isFirstLineHorizontal ? secondLineEndAnchor : firstLineEndAnchor;
    var horizontalY = (horizontalStartAnchor[1] + horizontalEndAnchor[1]) / 2;
    var verticalX = (verticalStartAnchor[0] + verticalEndAnchor[0]) / 2;
    var horizontalXMin = Math.min(horizontalStartAnchor[0], horizontalEndAnchor[0]);
    var horizontalXMax = Math.max(horizontalStartAnchor[0], horizontalEndAnchor[0]);
    var verticalYMin = Math.min(verticalStartAnchor[1], verticalEndAnchor[1]);
    var verticalYMax = Math.max(verticalStartAnchor[1], verticalEndAnchor[1]);
    return (verticalX >= horizontalXMin - EPS && verticalX <= horizontalXMax + EPS &&
        horizontalY >= verticalYMin - EPS && horizontalY <= verticalYMax + EPS);
}

/* 中心線を交差連結成分でクラスタ分け / Cluster lines by intersection-connected components */
function clusterLinesByIntersection(lines) {
    var n = lines.length;
    var visited = [];
    for (var i = 0; i < n; i++) visited.push(false);
    var clusters = [];
    for (var startIndex = 0; startIndex < n; startIndex++) {
        if (visited[startIndex]) continue;
        var cluster = [];
        var stack = [startIndex];
        visited[startIndex] = true;
        while (stack.length > 0) {
            var idx = stack.pop();
            cluster.push(lines[idx]);
            for (var j = 0; j < n; j++) {
                if (!visited[j] && intersectsAxisAligned(lines[idx], lines[j])) {
                    visited[j] = true;
                    stack.push(j);
                }
            }
        }
        clusters.push(cluster);
    }
    return clusters;
}

// =========================================
// メイン処理 / Main
// =========================================

function main() {
    try {
        if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
            alert(L('alertNoSelection'));
            return;
        }

        var opts = showOptionDialog();
        if (opts === null) return;

        /* mm → pt（1 mm = 72/25.4 pt）/ mm → pt conversion */
        var minShortSidePt = opts.minShortSideMm * 72 / 25.4;

        var selectedItems = app.activeDocument.selection;
        var generatedCenterLines = [];

        /* グループ内も含めて再帰的に長方形を収集 / Recursively gather rectangles, including inside groups */
        var rectsToProcess = [];
        collectRectsRecursive(selectedItems, rectsToProcess);

        for (var i = 0; i < rectsToProcess.length; i++) {
            var rect = rectsToProcess[i];
            /* CompoundPathItem 内のサブパスはコンパウンドの親へ配置 / For subpaths in compounds, use compound's parent */
            var placeTarget = rect.parent;
            try {
                if (placeTarget && placeTarget.typename === "CompoundPathItem") {
                    placeTarget = placeTarget.parent;
                }
            } catch (e) { placeTarget = null; }

            var newLine = convertRectToCenterLine(rect, true, minShortSidePt);
            if (newLine) {
                if (placeTarget) {
                    try { newLine.move(placeTarget, ElementPlacement.PLACEATEND); } catch (e) { }
                }
                generatedCenterLines.push(newLine);
            }
        }
        /* 交差連結成分ごとに結合パターンを決定 / Combine per intersection-connected cluster */
        if (generatedCenterLines.length > 0) {
            var commonStrokeWidth = opts.commonStroke
                ? getRepresentativeStrokeWidth(generatedCenterLines, opts.strokeStrategy)
                : null;
            var clusters = clusterLinesByIntersection(generatedCenterLines);
            var resultLines = [];
            for (var clusterIndex = 0; clusterIndex < clusters.length; clusterIndex++) {
                var cluster = clusters[clusterIndex];
                var combined = null;
                if (opts.connectLines) {
                    if (cluster.length === 4) {
                        combined = tryConnectFourLinesIntoRect(cluster, opts.strokeStrategy);
                    } else if (cluster.length === 3) {
                        combined = tryConnectThreeLinesIntoUShape(cluster, opts.strokeStrategy);
                    } else if (cluster.length === 2) {
                        combined = tryConnectTwoLinesIntoLShape(cluster, opts.strokeStrategy);
                    } else if (cluster.length >= 5) {
                        combined = tryConnectManyLinesIntoOutline(cluster, opts.strokeStrategy, commonStrokeWidth);
                    }
                }
                if (combined) {
                    resultLines.push(combined);
                } else {
                    for (var lineIndex = 0; lineIndex < cluster.length; lineIndex++) resultLines.push(cluster[lineIndex]);
                }
            }
            generatedCenterLines = resultLines;
            if (opts.commonStroke) {
                applyCommonStrokeWidth(generatedCenterLines, opts.strokeStrategy);
            }
            if (opts.printBlack) {
                applyPrintBlackStroke(generatedCenterLines);
            }
            app.activeDocument.selection = generatedCenterLines;
        }
    } catch (e) {
        alert(labelText('alertError') + "\n" + e);
    }
}

main();