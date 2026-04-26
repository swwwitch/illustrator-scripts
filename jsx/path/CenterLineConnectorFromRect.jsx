#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
- Excelなどからコピー＆ペーストされた長方形から罫線（グリッド・枠線）を自動生成し、線同士の関係に応じて格子化・結合・統合まで行う Illustrator 用スクリプト
- 中心線化（任意ON/OFF）、回転補正、除外条件、格子モード、外枠長方形化に対応
- 線幅は Illustrator の strokeUnits、短辺しきい値は rulerType を参照し、表示値と内部pt換算を一元管理
- 線幅の共通化・代表値選択・指定値、印刷用ブラック化、グループ化など仕上げ処理も含む
- IIFE構成、UI構築／イベント配線／値取得／実行フローを分離し、保守しやすい構造に整理
- Illustrator DOM の不安定な操作（move/remove/selection 等）は安全ヘルパー経由で実行し、例外処理を集約

- Illustrator script that generates rule lines (grids / frames) from rectangles copied and pasted from sources such as Excel, then grids, connects, or unifies them based on line relationships
- Supports center-line conversion (toggleable), rotation correction, exclusion rules, grid mode, and outer-frame rectangle conversion
- Uses Illustrator strokeUnits for stroke width and rulerType for the short-side threshold, with centralized conversion between display units and internal points
- Includes finishing controls such as stroke normalization, representative width selection, custom width, print black, and grouping
- Uses an IIFE structure and separates UI building, event binding, value reading, and execution flow for easier maintenance
- Runs unstable Illustrator DOM operations such as move/remove/selection through safe helper functions to centralize exception handling

### 更新履歴 / Update History

- v1.0.0 (20250612) : 初版作成 / Initial version
- v1.6.0 (20260427) : UI構造の分離（パネル・イベント・状態・プレビュー）、単位管理の整理（strokeUnits / rulerType）、安全操作ヘルパーの導入、命名整理 / Refactored UI structure (panels, events, state, preview), unified unit handling (strokeUnits / rulerType), added safe operation helpers, improved naming consistency

*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version & Localization
    // =========================================

    var SCRIPT_VERSION = "v1.6.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "長方形を線に変換し、連結処理", en: "Convert Rectangles to Lines and Connect" },
        lblShortSideLength: { ja: "短辺が", en: "Short side is" },
        lblGreaterEqual: { ja: "以上の長方形は対象外", en: "below threshold — skip" },
        pnlCenterLineConversion: { ja: "中心線化", en: "Center Line Conversion" },
        chkAngleCorrect: { ja: "角度補正", en: "Correct rotation" },
        pnlOption: { ja: "オプション", en: "Options" },
        pnlRule: { ja: "罫線", en: "Rule" },
        chkConnectAll: { ja: "連結処理", en: "Connect lines" },
        rdoConnectNone: { ja: "なし", en: "None" },
        rdoConnectGrid: { ja: "格子状に", en: "Grid" },
        chkOuterRect: { ja: "外枠を長方形化", en: "Outer frame as rectangle" },
        chkPrintBlack: { ja: "印刷用の「黒」（K100）にする", en: "Use print black (K100)" },
        lblStrokePref: { ja: "線幅", en: "Stroke Width" },
        chkCommonStroke: { ja: "線幅を共通にする", en: "Make stroke widths common" },
        chkGroup: { ja: "グループ化", en: "Group result" },
        strokeMax: { ja: "最大", en: "Max" },
        strokeMin: { ja: "最小", en: "Min" },
        strokeAvg: { ja: "平均", en: "Average" },
        strokeCustom: { ja: "指定", en: "Custom" },
        btnOutlineOn: { ja: "アウトライン表示", en: "Outline View" },
        btnOutlineOff: { ja: "プレビュー表示", en: "Preview View" },
        btnOk: { ja: "OK", en: "OK" },
        btnCancel: { ja: "キャンセル", en: "Cancel" },
        alertNoSelection: { ja: "長方形を1つ以上選択してください。", en: "Please select at least one rectangle." },
        alertError: { ja: "エラーが発生しました", en: "An error occurred" }
    };

    /* ラベル取得 / Get label */
    function getLabel(key) {
        return LABELS[key][lang];
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return getLabel(key) + (lang === 'ja' ? '：' : ':');
    }

    // =========================================
    // 安全操作ヘルパー / Safe Operation Helpers
    // =========================================

    /* ロック／非表示なら true / True when item is locked or hidden */
    function isLockedOrHidden(pageItem) {
        try {
            return !!(pageItem.locked || pageItem.hidden);
        } catch (e) {
            return false;
        }
    }

    /* PageItem を安全に削除 / Safely remove a PageItem */
    function removePageItemSafely(pageItem) {
        try { pageItem.remove(); } catch (e) { }
    }

    /* PageItem を安全に移動 / Safely move a PageItem */
    function movePageItemSafely(pageItem, destination, placement) {
        try { pageItem.move(destination, placement); } catch (e) { }
    }

    /* 線端形状を安全に設定 / Safely set stroke cap */
    function setStrokeCapSafely(pathItem, strokeCap) {
        try { pathItem.strokeCap = strokeCap; } catch (e) { }
    }

    /* 選択を安全に復元 / Safely restore document selection */
    function restoreSelectionSafely(selectionItems) {
        try { app.activeDocument.selection = selectionItems; } catch (e) { }
    }

    // =========================================
    // 単位変換 / Unit Conversion
    // =========================================

    /* 単位コードとラベルのマップ / Unit code to label map */
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

    /* 指定単位の1単位が何ptかを定義 / Point factor for one unit of each preference unit */
    var unitToPointFactorMap = {
        0: 72,                // in
        1: 72 / 25.4,         // mm
        2: 1,                 // pt
        3: 12,                // pica
        4: 72 / 2.54,         // cm
        5: 72 / 25.4 / 4,     // Q/H
        6: 1,                 // px
        7: 72,                // ft/in
        8: 72 / 0.0254,       // m
        9: 72 * 36,           // yd
        10: 72 * 12           // ft
    };

    /* 指定単位のpt換算係数を取得 / Get point factor for the given unit code */
    function getUnitToPointFactor(unitCode) {
        return unitToPointFactorMap[unitCode] || 1;
    }

    /* 線幅用の環境設定単位コードを取得 / Get current stroke unit code */
    function getCurrentStrokeUnitCode() {
        try {
            return app.preferences.getIntegerPreference("strokeUnits");
        } catch (e) {
            return 2;
        }
    }

    /* 線幅用の現在の単位ラベルを取得 / Get current stroke unit label */
    function getCurrentStrokeUnitLabel() {
        var unitCode = getCurrentStrokeUnitCode();
        return unitLabelMap[unitCode] || "pt";
    }

    /* 長さ用の環境設定単位コードを取得 / Get current ruler unit code */
    function getCurrentRulerUnitCode() {
        try {
            return app.preferences.getIntegerPreference("rulerType");
        } catch (e) {
            return 2;
        }
    }

    /* 長さ用の現在の単位ラベルを取得 / Get current ruler unit label */
    function getCurrentRulerUnitLabel() {
        var unitCode = getCurrentRulerUnitCode();
        return unitLabelMap[unitCode] || "pt";
    }

    /* pt値を指定単位の表示値へ変換 / Convert points to display unit value */
    function pointsToUnitValue(points, unitCode) {
        return points / getUnitToPointFactor(unitCode);
    }

    /* 指定単位の入力値をptへ変換 / Convert display unit value to points */
    function unitValueToPoints(value, unitCode) {
        return value * getUnitToPointFactor(unitCode);
    }

    // =========================================
    // 対象外マーカー（プレビュー）/ Exclusion markers (preview)
    // =========================================

    var EXCLUSION_PREVIEW_LAYER_NAME = "__center_line_preview__";


    function createExclusionMarkerColor() {
        var markerColor = new CMYKColor();
        markerColor.cyan = 0; markerColor.magenta = 100; markerColor.yellow = 100; markerColor.black = 0;
        return markerColor;
    }

    function createPrintBlackColor() {
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

    /* グループを再帰的に展開して2点パス（線分）を収集
       Recursively walk groups and collect 2-point paths (line segments) */
    function collectLinesRecursive(items, results) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var pageItem = items[itemIndex];
            if (isLockedOrHidden(pageItem)) continue;
            if (pageItem.typename === "PathItem") {
                if (pageItem.pathPoints.length === 2) results.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                collectLinesRecursive(pageItem.pageItems, results);
            }
        }
    }

    /* グループを再帰的に展開して4点閉じ長方形を収集
       Recursively walk groups and collect 4-point closed rectangles */
    function collectRectsRecursive(items, results) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var pageItem = items[itemIndex];
            if (isLockedOrHidden(pageItem)) continue;
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

    /* 選択中の長方形から、短辺の最大値を指定単位で取得
       Get the largest short side among selected rectangles in the given unit */
    function getMaxShortSideInUnitFromItems(items, unitCode) {
        var rectangleItems = [];
        collectRectsRecursive(items, rectangleItems);
        var maxShortSidePt = 0;

        for (var rectIndex = 0; rectIndex < rectangleItems.length; rectIndex++) {
            try {
                var bounds = rectangleItems[rectIndex].geometricBounds;
                var rectWidth = Math.abs(bounds[2] - bounds[0]);
                var rectHeight = Math.abs(bounds[1] - bounds[3]);
                var shortSide = Math.min(rectWidth, rectHeight);
                if (shortSide > maxShortSidePt) maxShortSidePt = shortSide;
            } catch (e) { }
        }

        return Math.round(pointsToUnitValue(maxShortSidePt, unitCode) * 10) / 10;
    }

    /* 対象外オブジェクトを M100Y100・半透明で複製してマーカー表示
       Duplicate excluded objects as M100Y100 semi-transparent markers */
    function refreshExclusionPreview(items, minShortSidePt) {
        clearExclusionPreviewLayer();
        if (!items || items.length === 0) { app.redraw(); return; }
        var layer = ensureExclusionPreviewLayer();
        var markerColor = createExclusionMarkerColor();

        var rectangleItems = [];
        collectRectsRecursive(items, rectangleItems);

        for (var rectangleIndex = 0; rectangleIndex < rectangleItems.length; rectangleIndex++) {
            if (!isExcludedRect(rectangleItems[rectangleIndex], minShortSidePt)) continue;
            try {
                var markerRect = rectangleItems[rectangleIndex].duplicate(layer, ElementPlacement.PLACEATBEGINNING);
                markerRect.filled = true;
                markerRect.fillColor = markerColor;
                markerRect.stroked = false;
                markerRect.opacity = 50;
            } catch (e) { }
        }
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 中心線パネルを構築 / Build center-line panel */
    function buildCenterLinePanel(parent, PANEL_MARGINS, rulerUnitCode, rulerUnitLabel) {
        var panel = parent.add("panel", undefined, getLabel('pnlCenterLineConversion'));
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        panel.spacing = 6;

        var initialRectsForDetect = [];
        collectRectsRecursive(app.activeDocument.selection, initialRectsForDetect);
        var hasRectsInSelection = initialRectsForDetect.length > 0;

        var cbCenterLine = panel.add("checkbox", undefined, getLabel('pnlCenterLineConversion'));
        cbCenterLine.value = hasRectsInSelection;

        var cbAngleCorrect = panel.add("checkbox", undefined, getLabel('chkAngleCorrect'));
        cbAngleCorrect.value = true;

        var SHORT_MIN = 0;
        var SHORT_MAX = getMaxShortSideInUnitFromItems(app.activeDocument.selection, rulerUnitCode);
        if (SHORT_MAX <= 0) SHORT_MAX = 10;
        var SHORT_DEFAULT = Math.min(10, SHORT_MAX);

        var minShortSideGroup = panel.add("group");
        minShortSideGroup.orientation = "row";
        minShortSideGroup.alignChildren = "center";
        minShortSideGroup.spacing = 4;

        var cbMinShortSide = minShortSideGroup.add("checkbox", undefined, "");
        cbMinShortSide.value = false;
        minShortSideGroup.add("statictext", undefined, getLabel('lblShortSideLength'));

        var minShortSideInput = minShortSideGroup.add("edittext", undefined, SHORT_DEFAULT.toFixed(1));
        minShortSideInput.characters = 5;

        minShortSideGroup.add("statictext", undefined, rulerUnitLabel + getLabel('lblGreaterEqual'));

        var minShortSideSlider = panel.add("slider", undefined, SHORT_DEFAULT, SHORT_MIN, SHORT_MAX);
        minShortSideSlider.preferredSize.width = 247;

        return {
            panel: panel,
            cbCenterLine: cbCenterLine,
            cbAngleCorrect: cbAngleCorrect,
            SHORT_MIN: SHORT_MIN,
            SHORT_MAX: SHORT_MAX,
            SHORT_DEFAULT: SHORT_DEFAULT,
            cbMinShortSide: cbMinShortSide,
            minShortSideInput: minShortSideInput,
            minShortSideSlider: minShortSideSlider
        };
    }

    /* 連結オプションパネルを構築 / Build connect options panel */
    function buildConnectOptionsPanel(parent, PANEL_MARGINS) {
        var panel = parent.add("panel", undefined, getLabel('pnlOption'));
        panel.orientation = "column";
        panel.alignChildren = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = 8;

        var rbConnectNone = panel.add("radiobutton", undefined, getLabel('rdoConnectNone'));
        var rbConnectAll = panel.add("radiobutton", undefined, getLabel('chkConnectAll'));

        var gridRow = panel.add("group");
        gridRow.orientation = "row";
        gridRow.alignChildren = "left";
        gridRow.spacing = 12;
        gridRow.margins = 0;

        var rbConnectGrid = gridRow.add("radiobutton", undefined, getLabel('rdoConnectGrid'));
        var cbOuterRect = gridRow.add("checkbox", undefined, getLabel('chkOuterRect'));

        cbOuterRect.value = false;
        rbConnectGrid.value = true;

        return {
            panel: panel,
            rbConnectNone: rbConnectNone,
            rbConnectAll: rbConnectAll,
            rbConnectGrid: rbConnectGrid,
            cbOuterRect: cbOuterRect,
            connectRadios: [rbConnectNone, rbConnectAll, rbConnectGrid]
        };
    }

    /* 罫線パネルを構築 / Build rule panel */
    function buildRulePanel(parent, PANEL_MARGINS, strokeUnitLabel) {
        var panel = parent.add("panel", undefined, getLabel('pnlRule'));
        panel.orientation = "column";
        panel.alignChildren = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = 8;

        var cbPrintBlack = panel.add("checkbox", undefined, getLabel('chkPrintBlack'));
        cbPrintBlack.value = false;

        var strokeWidthPanel = panel.add("panel", undefined, getLabel('lblStrokePref'));
        strokeWidthPanel.orientation = "column";
        strokeWidthPanel.alignChildren = "left";
        strokeWidthPanel.margins = PANEL_MARGINS;
        strokeWidthPanel.spacing = 6;

        var strokeRadioGroup = strokeWidthPanel.add("group");
        strokeRadioGroup.orientation = "row";
        strokeRadioGroup.alignChildren = "center";
        strokeRadioGroup.spacing = 8;

        var rbStrokeMax = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeMax'));
        var rbStrokeMin = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeMin'));
        var rbStrokeAvg = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeAvg'));
        var rbStrokeCustom = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeCustom'));

        rbStrokeMin.value = true;

        var customStrokeInput = strokeRadioGroup.add("edittext", undefined, "0.25");
        customStrokeInput.characters = 4;

        strokeRadioGroup.add("statictext", undefined, strokeUnitLabel);

        var cbCommonStroke = strokeWidthPanel.add("checkbox", undefined, getLabel('chkCommonStroke'));
        cbCommonStroke.value = false;

        var cbGroup = panel.add("checkbox", undefined, getLabel('chkGroup'));
        cbGroup.value = true;

        return {
            panel: panel,
            cbPrintBlack: cbPrintBlack,
            rbStrokeMax: rbStrokeMax,
            rbStrokeMin: rbStrokeMin,
            rbStrokeAvg: rbStrokeAvg,
            rbStrokeCustom: rbStrokeCustom,
            customStrokeInput: customStrokeInput,
            cbCommonStroke: cbCommonStroke,
            cbGroup: cbGroup
        };
    }

    /* ボタン行を構築 / Build dialog button row */
    function buildDialogButtonRow(dialog) {
        var row = dialog.add("group");
        row.orientation = "row";
        row.alignChildren = ["fill", "center"];
        row.alignment = "fill";
        row.spacing = 8;

        var left = row.add("group");
        left.alignment = ["left", "center"];

        var btnOutlineToggle = left.add("button", undefined, getLabel('btnOutlineOn'));

        var spacer = row.add("statictext", undefined, "");
        spacer.alignment = ["fill", "center"];

        var right = row.add("group");
        right.orientation = "row";
        right.alignment = ["right", "center"];
        right.spacing = 8;

        right.add("button", undefined, getLabel('btnCancel'), { name: "cancel" });
        right.add("button", undefined, getLabel('btnOk'), { name: "ok" });

        return {
            btnOutlineToggle: btnOutlineToggle,
            isOutlineMode: false
        };
    }

    /* ダイアログUI構築 / Build dialog UI */
    function buildOptionDialogUI(dialog) {
        var PANEL_MARGINS = [15, 20, 15, 10];
        var strokeUnitCode = getCurrentStrokeUnitCode();
        var strokeUnitLabel = getCurrentStrokeUnitLabel();
        var rulerUnitCode = getCurrentRulerUnitCode();
        var rulerUnitLabel = getCurrentRulerUnitLabel();

        var panelColumn = dialog.add("group");
        panelColumn.orientation = "column";
        panelColumn.alignChildren = "fill";
        panelColumn.spacing = 12;

        var center = buildCenterLinePanel(panelColumn, PANEL_MARGINS, rulerUnitCode, rulerUnitLabel);
        var connect = buildConnectOptionsPanel(panelColumn, PANEL_MARGINS);
        var rule = buildRulePanel(panelColumn, PANEL_MARGINS, strokeUnitLabel);
        var buttons = buildDialogButtonRow(dialog);

        var snapshotSelection = [];
        for (var ssi = 0; ssi < app.activeDocument.selection.length; ssi++) {
            snapshotSelection.push(app.activeDocument.selection[ssi]);
        }

        return {
            strokeUnitCode: strokeUnitCode,
            rulerUnitCode: rulerUnitCode,

            cbCenterLine: center.cbCenterLine,
            cbAngleCorrect: center.cbAngleCorrect,
            SHORT_MIN: center.SHORT_MIN,
            SHORT_MAX: center.SHORT_MAX,
            SHORT_DEFAULT: center.SHORT_DEFAULT,
            cbMinShortSide: center.cbMinShortSide,
            minShortSideInput: center.minShortSideInput,
            minShortSideSlider: center.minShortSideSlider,

            snapshotSelection: snapshotSelection,

            rbConnectNone: connect.rbConnectNone,
            rbConnectAll: connect.rbConnectAll,
            rbConnectGrid: connect.rbConnectGrid,
            cbOuterRect: connect.cbOuterRect,
            connectRadios: connect.connectRadios,

            cbPrintBlack: rule.cbPrintBlack,
            rbStrokeMax: rule.rbStrokeMax,
            rbStrokeMin: rule.rbStrokeMin,
            rbStrokeAvg: rule.rbStrokeAvg,
            rbStrokeCustom: rule.rbStrokeCustom,
            customStrokeInput: rule.customStrokeInput,
            cbCommonStroke: rule.cbCommonStroke,
            cbGroup: rule.cbGroup,

            btnOutlineToggle: buttons.btnOutlineToggle,
            isOutlineMode: buttons.isOutlineMode
        };
    }

    /* UIから値を読み取る / Read values from UI */
    function readOptionDialogValues(ui) {
        var strokeStrategy;
        if (ui.rbStrokeCustom.value) strokeStrategy = "custom";
        else if (ui.rbStrokeMin.value) strokeStrategy = "min";
        else if (ui.rbStrokeAvg.value) strokeStrategy = "avg";
        else strokeStrategy = "max";

        var customStrokeWidth = parseFloat(ui.customStrokeInput.text);
        if (isNaN(customStrokeWidth) || customStrokeWidth <= 0) customStrokeWidth = 0.25;

        // Strengthen NaN handling for minShortSideInput
        var minShortSideValue = parseFloat(ui.minShortSideInput.text);
        if (isNaN(minShortSideValue)) minShortSideValue = ui.SHORT_DEFAULT;
        if (minShortSideValue < ui.SHORT_MIN) minShortSideValue = ui.SHORT_MIN;
        if (minShortSideValue > ui.SHORT_MAX) minShortSideValue = ui.SHORT_MAX;

        var connectMode = ui.rbConnectGrid.value ? "grid" : (ui.rbConnectAll.value ? "connect" : "none");

        return {
            correctRotation: ui.cbAngleCorrect.value,
            enableCenterLine: ui.cbCenterLine.value,
            minShortSidePt: ui.cbMinShortSide.value ? unitValueToPoints(minShortSideValue, ui.rulerUnitCode) : 0,
            connectMode: connectMode,
            outerRect: ui.cbOuterRect.value,
            printBlack: ui.cbPrintBlack.value,
            commonStroke: ui.cbCommonStroke.value,
            strokeStrategy: strokeStrategy,
            customStrokeWidth: unitValueToPoints(customStrokeWidth, ui.strokeUnitCode),
            groupResult: ui.cbGroup.value
        };
    }

    /* UI状態をまとめて同期 / Synchronize all UI enabled states */
    function syncUIState(ui) {
        var centerLineActive = ui.cbCenterLine.value;
        var minShortSideActive = centerLineActive && ui.cbMinShortSide.value;
        var gridActive = ui.rbConnectGrid.value;

        ui.cbMinShortSide.enabled = centerLineActive;
        ui.cbAngleCorrect.enabled = centerLineActive;
        ui.minShortSideInput.enabled = minShortSideActive;
        ui.minShortSideSlider.enabled = minShortSideActive;

        ui.cbOuterRect.enabled = gridActive;
        if (ui.cbGroup) ui.cbGroup.enabled = gridActive;

        ui.customStrokeInput.enabled = ui.rbStrokeCustom.value;
    }

    /* ダイアログ表示前の初期状態を同期 / Initialize dialog state before showing */
    function initializeOptionDialogState(ui) {
        syncUIState(ui);
        updateExclusionPreviewFromUIState(ui);
    }

    /* UI状態からプレビューを更新 / Update preview from UI state */
    function updateExclusionPreviewFromUIState(ui) {
        if (!ui.cbCenterLine.value) {
            clearExclusionPreviewLayer();
            app.redraw();
            return;
        }

        var minShortSidePt = 0;
        if (ui.cbMinShortSide.value) {
            var minShortSideValue = parseFloat(ui.minShortSideInput.text);
            if (isNaN(minShortSideValue)) minShortSideValue = ui.SHORT_DEFAULT;
            minShortSidePt = unitValueToPoints(minShortSideValue, ui.rulerUnitCode);
        }

        refreshExclusionPreview(ui.snapshotSelection, minShortSidePt);
    }

    /* ダイアログイベントを配線 / Bind dialog events */
    function bindOptionDialogEvents(ui) {

        function roundTo1(value) {
            return Math.round(value * 10) / 10;
        }

        function selectConnectRadio(selected) {
            for (var radioIndex = 0; radioIndex < ui.connectRadios.length; radioIndex++) {
                ui.connectRadios[radioIndex].value = (ui.connectRadios[radioIndex] === selected);
            }
        }

        ui.cbMinShortSide.onClick = function () {
            syncUIState(ui);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.cbCenterLine.onClick = function () {
            syncUIState(ui);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.minShortSideSlider.onChanging = function () {
            ui.minShortSideInput.text = roundTo1(ui.minShortSideSlider.value).toFixed(1);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.minShortSideInput.onChange = function () {
            var minShortSideValue = parseFloat(ui.minShortSideInput.text);
            if (isNaN(minShortSideValue)) minShortSideValue = ui.minShortSideSlider.value;
            if (minShortSideValue < ui.SHORT_MIN) minShortSideValue = ui.SHORT_MIN;
            if (minShortSideValue > ui.SHORT_MAX) minShortSideValue = ui.SHORT_MAX;
            ui.minShortSideSlider.value = minShortSideValue;
            ui.minShortSideInput.text = roundTo1(minShortSideValue).toFixed(1);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.rbConnectNone.onClick = function () {
            selectConnectRadio(ui.rbConnectNone);
            syncUIState(ui);
        };

        ui.rbConnectAll.onClick = function () {
            selectConnectRadio(ui.rbConnectAll);
            syncUIState(ui);
        };

        ui.rbConnectGrid.onClick = function () {
            selectConnectRadio(ui.rbConnectGrid);
            syncUIState(ui);
            ui.rbStrokeCustom.value = true;
            syncUIState(ui);
            ui.cbCommonStroke.value = true;
        };

        ui.rbStrokeMax.onClick = function () { syncUIState(ui); };
        ui.rbStrokeMin.onClick = function () { syncUIState(ui); };
        ui.rbStrokeAvg.onClick = function () { syncUIState(ui); };
        ui.rbStrokeCustom.onClick = function () {
            syncUIState(ui);
            ui.cbCommonStroke.value = true;
        };

        ui.btnOutlineToggle.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                ui.isOutlineMode = !ui.isOutlineMode;
                ui.btnOutlineToggle.text = ui.isOutlineMode ? getLabel('btnOutlineOff') : getLabel('btnOutlineOn');
            } catch (e) { }
        };
    }

    /* オプション設定用ダイアログを表示 / Show options dialog */
    function showOptionDialog() {
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;

        var ui = buildOptionDialogUI(dialog);

        /* イベント配線 / Bind dialog events */
        bindOptionDialogEvents(ui);

        /* 初期状態の同期 / Synchronize initial UI state */
        initializeOptionDialogState(ui);

        var dialogResult = dialog.show();

        // Clean up preview layer and restore selection
        clearExclusionPreviewLayer();
        restoreSelectionSafely(ui.snapshotSelection);
        app.redraw();

        if (dialogResult !== 1) return null;

        return readOptionDialogValues(ui);
    }

    // =========================================
    // 中心線描画処理 / Center Line Drawing
    // =========================================

    /* 長方形を中心線に変換し、元の長方形を削除 / Convert rectangle into a center line and remove original */
    function convertRectToCenterLine(rect, correctRotation, minShortSidePt) {
        /* 回転補正：最初の辺の角度を測り、最寄りの 90° 軸（水平 0/180、垂直 90/270）から
           0.5°〜10° 以内のずれなら、その軸にスナップする
           Rotation correction: measure first-edge angle and, if it deviates 0.5°–10° from
           the nearest 90° axis (horizontal 0/180, vertical 90/270), snap to that axis */
        if (correctRotation && rect.pathPoints.length === 4) {
            var firstAnchor = rect.pathPoints[0].anchor;
            var secondAnchor = rect.pathPoints[1].anchor;

            var deltaX = secondAnchor[0] - firstAnchor[0];
            var deltaY = secondAnchor[1] - firstAnchor[1];
            var angleRad = Math.atan2(deltaY, deltaX);
            var angleDeg = angleRad * 180 / Math.PI;
            if (angleDeg < 0) angleDeg += 360;

            var nearestAxisDeg = Math.round(angleDeg / 90) * 90;
            var deviationDeg = angleDeg - nearestAxisDeg;
            var absDeviationDeg = Math.abs(deviationDeg);

            if (absDeviationDeg >= 0.5 && absDeviationDeg <= 10) {
                rect.rotate(-deviationDeg);
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

        removePageItemSafely(rect);
        return centerLine;
    }

    /* 太さが混在するときの代表値を取得（"custom" 指定時は customWidth を返す）
       Pick a representative stroke width when widths differ ("custom" returns customWidth) */
    function getRepresentativeStrokeWidth(lines, strategy, customWidth) {
        if (strategy === "custom" && typeof customWidth === "number" && customWidth > 0) {
            return customWidth;
        }
        var initialStrokeWidth = lines[0].strokeWidth;
        if (strategy === "min") {
            var minimumStrokeWidth = initialStrokeWidth;
            for (var lineIndex = 1; lineIndex < lines.length; lineIndex++) {
                if (lines[lineIndex].strokeWidth < minimumStrokeWidth) minimumStrokeWidth = lines[lineIndex].strokeWidth;
            }
            return minimumStrokeWidth;
        } else if (strategy === "avg") {
            var sum = 0;
            for (var averageLineIndex = 0; averageLineIndex < lines.length; averageLineIndex++) {
                sum += lines[averageLineIndex].strokeWidth;
            }
            return sum / lines.length;
        }
        /* default: max */
        var maximumStrokeWidth = initialStrokeWidth;
        for (var maximumLineIndex = 1; maximumLineIndex < lines.length; maximumLineIndex++) {
            if (lines[maximumLineIndex].strokeWidth > maximumStrokeWidth) maximumStrokeWidth = lines[maximumLineIndex].strokeWidth;
        }
        return maximumStrokeWidth;
    }

    /* 生成結果の線幅を共通化（コンパウンド／グループ内のサブパスも対象）
       Make stroke width common across generated results (recurses into compounds and groups) */

    function applyCommonStrokeWidth(items, strategy, customWidth) {
        if (!items || items.length === 0) return;

        var strokeItems = [];
        function collectStrokedRecursive(item) {
            try {
                if (item.typename === 'PathItem') {
                    if (item.stroked) strokeItems.push(item);
                } else if (item.typename === 'CompoundPathItem') {
                    for (var compoundPathIndex = 0; compoundPathIndex < item.pathItems.length; compoundPathIndex++) {
                        collectStrokedRecursive(item.pathItems[compoundPathIndex]);
                    }
                } else if (item.typename === 'GroupItem') {
                    for (var groupPageItemIndex = 0; groupPageItemIndex < item.pageItems.length; groupPageItemIndex++) {
                        collectStrokedRecursive(item.pageItems[groupPageItemIndex]);
                    }
                }
            } catch (e) { }
        }
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            if (items[itemIndex]) collectStrokedRecursive(items[itemIndex]);
        }
        if (strokeItems.length === 0) return;

        var commonWidth = getRepresentativeStrokeWidth(strokeItems, strategy, customWidth);
        for (var strokeItemIndex = 0; strokeItemIndex < strokeItems.length; strokeItemIndex++) {
            try { strokeItems[strokeItemIndex].strokeWidth = commonWidth; } catch (e) { }
        }
    }

    /* 生成結果の線色を印刷用の黒（C0 M0 Y0 K100）に統一
       Apply print black (C0 M0 Y0 K100) to generated result strokes */
    function applyPrintBlackStroke(items) {
        if (!items || items.length === 0) return;

        var printBlackColor = createPrintBlackColor();

        function applyBlackRecursive(item) {
            try {
                if (item.typename === 'PathItem') {
                    item.stroked = true;
                    item.strokeColor = printBlackColor;
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
    function tryConnectFourLinesIntoRect(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length !== 4) return null;

        var EPS = 0.01;
        var horizontalLineInfos = [];
        var verticalLineInfos = [];

        for (var lineIndex = 0; lineIndex < 4; lineIndex++) {
            var lineItem = lines[lineIndex];
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
        rectanglePath.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
        rectanglePath.filled = false;

        var corners = [
            [xMin, yMax],
            [xMax, yMax],
            [xMax, yMin],
            [xMin, yMin]
        ];
        for (var cornerIndex = 0; cornerIndex < 4; cornerIndex++) {
            var cornerPoint = rectanglePath.pathPoints.add();
            cornerPoint.anchor = corners[cornerIndex];
            cornerPoint.leftDirection = corners[cornerIndex];
            cornerPoint.rightDirection = corners[cornerIndex];
        }

        for (var removeLineIndex = 0; removeLineIndex < lines.length; removeLineIndex++) {
            removePageItemSafely(lines[removeLineIndex]);
        }

        return rectanglePath;
    }

    /* 3本（2H+1V または 1H+2V）の中心線をコの字型の開いたパスに変換
       Convert 3 lines (2H+1V or 1H+2V) into an open U-shape path */
    function tryConnectThreeLinesIntoUShape(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length !== 3) return null;

        var EPS = 0.01;
        var horizontalLineInfos = [], verticalLineInfos = [];

        for (var lineIndex = 0; lineIndex < 3; lineIndex++) {
            var lineItem = lines[lineIndex];
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
        path.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
        path.filled = false;

        for (var anchorIndex = 0; anchorIndex < anchors.length; anchorIndex++) {
            var pathPoint = path.pathPoints.add();
            pathPoint.anchor = anchors[anchorIndex];
            pathPoint.leftDirection = anchors[anchorIndex];
            pathPoint.rightDirection = anchors[anchorIndex];
        }

        for (var removeIndex = 0; removeIndex < lines.length; removeIndex++) {
            removePageItemSafely(lines[removeIndex]);
        }

        return path;
    }

    /* 2本（1H+1V）の中心線をL字型の開いたパスに変換
       Convert 2 lines (1H+1V) into an open L-shape path */
    function tryConnectTwoLinesIntoLShape(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length !== 2) return null;

        var EPS = 0.01;
        var horizontalLineInfos = [], verticalLineInfos = [];

        for (var lineIndex = 0; lineIndex < 2; lineIndex++) {
            var lineItem = lines[lineIndex];
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
        path.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
        path.filled = false;

        for (var anchorIndex = 0; anchorIndex < anchors.length; anchorIndex++) {
            var pathPoint = path.pathPoints.add();
            pathPoint.anchor = anchors[anchorIndex];
            pathPoint.leftDirection = anchors[anchorIndex];
            pathPoint.rightDirection = anchors[anchorIndex];
        }

        for (var removeIndex = 0; removeIndex < lines.length; removeIndex++) {
            removePageItemSafely(lines[removeIndex]);
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
    function tryConnectManyLinesIntoOutline(lines, strokeStrategy, commonStrokeWidth, customWidth) {
        if (!lines || lines.length < 5) return null;

        var doc = app.activeDocument;
        var representativeStrokeWidth = (commonStrokeWidth !== null && commonStrokeWidth !== undefined)
            ? commonStrokeWidth
            : getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
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
        for (var lineIndex = lines.length - 1; lineIndex >= 0; lineIndex--) {
            movePageItemSafely(lines[lineIndex], group, ElementPlacement.PLACEATBEGINNING);
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
                movePageItemSafely(result, groupParent, ElementPlacement.PLACEATEND);
                removePageItemSafely(group);
            }
        } catch (e) { }

        return result;
    }

    /* クラスタを格子状に整列：水平線は X レンジ全体、垂直線は Y レンジ全体に伸ばし、線端は突出端で揃える
       Align a cluster into a grid: extend H lines to full X range, V lines to full Y range, with projecting end caps */
    function tryAlignClusterAsGrid(lines) {
        if (!lines || lines.length < 4) return null;

        var horizontalInfos = [];
        var verticalInfos = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var lineItem = lines[lineIndex];
            if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) continue;
            var startAnchor = lineItem.pathPoints[0].anchor;
            var endAnchor = lineItem.pathPoints[1].anchor;
            var deltaX = Math.abs(startAnchor[0] - endAnchor[0]);
            var deltaY = Math.abs(startAnchor[1] - endAnchor[1]);
            if (deltaX >= deltaY) {
                horizontalInfos.push({ path: lineItem, y: (startAnchor[1] + endAnchor[1]) / 2 });
            } else {
                verticalInfos.push({ path: lineItem, x: (startAnchor[0] + endAnchor[0]) / 2 });
            }
        }
        if (horizontalInfos.length < 2 || verticalInfos.length < 2) return null;

        var minX = verticalInfos[0].x, maxX = verticalInfos[0].x;
        for (var verticalInfoIndex = 1; verticalInfoIndex < verticalInfos.length; verticalInfoIndex++) {
            if (verticalInfos[verticalInfoIndex].x < minX) minX = verticalInfos[verticalInfoIndex].x;
            if (verticalInfos[verticalInfoIndex].x > maxX) maxX = verticalInfos[verticalInfoIndex].x;
        }
        var minY = horizontalInfos[0].y, maxY = horizontalInfos[0].y;
        for (var horizontalInfoIndex = 1; horizontalInfoIndex < horizontalInfos.length; horizontalInfoIndex++) {
            if (horizontalInfos[horizontalInfoIndex].y < minY) minY = horizontalInfos[horizontalInfoIndex].y;
            if (horizontalInfos[horizontalInfoIndex].y > maxY) maxY = horizontalInfos[horizontalInfoIndex].y;
        }

        function setLineEndpoints(pathItem, anchorA, anchorB) {
            pathItem.pathPoints[0].anchor = anchorA;
            pathItem.pathPoints[0].leftDirection = anchorA;
            pathItem.pathPoints[0].rightDirection = anchorA;
            pathItem.pathPoints[1].anchor = anchorB;
            pathItem.pathPoints[1].leftDirection = anchorB;
            pathItem.pathPoints[1].rightDirection = anchorB;
        }

        for (var horizontalLineIndex = 0; horizontalLineIndex < horizontalInfos.length; horizontalLineIndex++) {
            var horizontalPath = horizontalInfos[horizontalLineIndex].path;
            setStrokeCapSafely(horizontalPath, StrokeCap.PROJECTINGENDCAP);
            setLineEndpoints(horizontalPath, [minX, horizontalInfos[horizontalLineIndex].y], [maxX, horizontalInfos[horizontalLineIndex].y]);
        }
        /* Illustrator は Y 上方向が正 / Illustrator Y axis points up */
        var topY = Math.max(minY, maxY);
        var bottomY = Math.min(minY, maxY);
        for (var verticalLineIndex = 0; verticalLineIndex < verticalInfos.length; verticalLineIndex++) {
            var verticalPath = verticalInfos[verticalLineIndex].path;
            setStrokeCapSafely(verticalPath, StrokeCap.PROJECTINGENDCAP);
            setLineEndpoints(verticalPath, [verticalInfos[verticalLineIndex].x, topY], [verticalInfos[verticalLineIndex].x, bottomY]);
        }

        return lines;
    }

    /* 整列済みクラスタの最外周4本（top/bottom H と left/right V）を長方形に置換し、内側の線はそのまま残す
       Replace the outermost 4 lines (top/bottom H, left/right V) of an aligned cluster with a rectangle, keeping the inner lines */
    function convertOuterFrameToRect(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length < 4) return null;

        var horizontalInfos = [];
        var verticalInfos = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var lineItem = lines[lineIndex];
            if (!lineItem || !lineItem.pathPoints || lineItem.pathPoints.length !== 2) continue;
            var startAnchor = lineItem.pathPoints[0].anchor;
            var endAnchor = lineItem.pathPoints[1].anchor;
            var deltaX = Math.abs(startAnchor[0] - endAnchor[0]);
            var deltaY = Math.abs(startAnchor[1] - endAnchor[1]);
            if (deltaX >= deltaY) {
                horizontalInfos.push({ path: lineItem, y: (startAnchor[1] + endAnchor[1]) / 2 });
            } else {
                verticalInfos.push({ path: lineItem, x: (startAnchor[0] + endAnchor[0]) / 2 });
            }
        }
        if (horizontalInfos.length < 2 || verticalInfos.length < 2) return null;

        var topHorizontalInfo = horizontalInfos[0];
        var bottomHorizontalInfo = horizontalInfos[0];
        for (var horizontalInfoIndex = 1; horizontalInfoIndex < horizontalInfos.length; horizontalInfoIndex++) {
            if (horizontalInfos[horizontalInfoIndex].y > topHorizontalInfo.y) topHorizontalInfo = horizontalInfos[horizontalInfoIndex];
            if (horizontalInfos[horizontalInfoIndex].y < bottomHorizontalInfo.y) bottomHorizontalInfo = horizontalInfos[horizontalInfoIndex];
        }
        var leftVerticalInfo = verticalInfos[0];
        var rightVerticalInfo = verticalInfos[0];
        for (var verticalInfoIndex = 1; verticalInfoIndex < verticalInfos.length; verticalInfoIndex++) {
            if (verticalInfos[verticalInfoIndex].x < leftVerticalInfo.x) leftVerticalInfo = verticalInfos[verticalInfoIndex];
            if (verticalInfos[verticalInfoIndex].x > rightVerticalInfo.x) rightVerticalInfo = verticalInfos[verticalInfoIndex];
        }

        var outerLines = [topHorizontalInfo.path, bottomHorizontalInfo.path, leftVerticalInfo.path, rightVerticalInfo.path];

        var minX = leftVerticalInfo.x;
        var maxX = rightVerticalInfo.x;
        var minY = bottomHorizontalInfo.y;
        var maxY = topHorizontalInfo.y;

        var rectanglePath = app.activeDocument.pathItems.add();
        rectanglePath.closed = true;
        rectanglePath.stroked = topHorizontalInfo.path.stroked;
        rectanglePath.strokeColor = topHorizontalInfo.path.strokeColor;
        rectanglePath.strokeWidth = getRepresentativeStrokeWidth(outerLines, strokeStrategy, customWidth);
        rectanglePath.filled = false;

        var corners = [
            [minX, maxY],
            [maxX, maxY],
            [maxX, minY],
            [minX, minY]
        ];
        for (var cornerIndex = 0; cornerIndex < 4; cornerIndex++) {
            var cornerPoint = rectanglePath.pathPoints.add();
            cornerPoint.anchor = corners[cornerIndex];
            cornerPoint.leftDirection = corners[cornerIndex];
            cornerPoint.rightDirection = corners[cornerIndex];
        }

        /* 内側線のみ残す（参照比較で外周4本を除外） / Keep only inner lines (exclude outer 4 by reference) */
        var remainingLines = [rectanglePath];
        for (var remainingLineIndex = 0; remainingLineIndex < lines.length; remainingLineIndex++) {
            var isOuterLine = false;
            for (var outerLineIndex = 0; outerLineIndex < outerLines.length; outerLineIndex++) {
                if (lines[remainingLineIndex] === outerLines[outerLineIndex]) { isOuterLine = true; break; }
            }
            if (!isOuterLine) remainingLines.push(lines[remainingLineIndex]);
        }

        for (var removeOuterLineIndex = 0; removeOuterLineIndex < outerLines.length; removeOuterLineIndex++) {
            removePageItemSafely(outerLines[removeOuterLineIndex]);
        }

        return remainingLines;
    }

    /* 軸並行な2線分が交差するか判定（端点接触も許容、tolerance で「少し届かない線」も交差扱い）
       Test whether 2 axis-aligned segments intersect (endpoint touch allowed; tolerance lets near-misses count as intersecting) */
    function intersectsAxisAligned(lnA, lnB, tolerance) {
        var EPS = (typeof tolerance === "number" && tolerance > 0) ? tolerance : 0.01;
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

    /* 中心線を交差連結成分でクラスタ分け（tolerance で許容幅を指定可）
       Cluster lines by intersection-connected components (optional tolerance for near-miss tolerance) */
    function clusterLinesByIntersection(lines, tolerance) {
        var lineCount = lines.length;
        var visited = [];
        for (var visitedIndex = 0; visitedIndex < lineCount; visitedIndex++) visited.push(false);
        var clusters = [];
        for (var startIndex = 0; startIndex < lineCount; startIndex++) {
            if (visited[startIndex]) continue;
            var cluster = [];
            var stack = [startIndex];
            visited[startIndex] = true;
            while (stack.length > 0) {
                var currentLineIndex = stack.pop();
                cluster.push(lines[currentLineIndex]);
                for (var compareLineIndex = 0; compareLineIndex < lineCount; compareLineIndex++) {
                    if (!visited[compareLineIndex] && intersectsAxisAligned(lines[currentLineIndex], lines[compareLineIndex], tolerance)) {
                        visited[compareLineIndex] = true;
                        stack.push(compareLineIndex);
                    }
                }
            }
            clusters.push(cluster);
        }
        return clusters;
    }

    /* 格子モード用のクラスタ許容幅：最大ストローク幅の3倍（最低5pt）。
       thin な線も拾えるよう floor を持たせる
       Grid-mode cluster tolerance: max stroke width × 3, floored at 5pt so thin lines still get coverage */
    function computeGridClusterTolerance(lines) {
        if (!lines || lines.length === 0) return 0;
        var maxStrokeWidth = 0;
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            try {
                if (lines[lineIndex].strokeWidth > maxStrokeWidth) maxStrokeWidth = lines[lineIndex].strokeWidth;
            } catch (e) { }
        }
        var tolerance = maxStrokeWidth * 3;
        if (tolerance < 5) tolerance = 5;
        return tolerance;
    }


    // =========================================
    // 実行フロー補助 / Execution Flow Helpers
    // =========================================

    /* 選択内容から処理対象の中心線を生成または収集
       Create or collect target center lines from the current selection */
    function buildTargetCenterLines(selectedItems, options) {
        var generatedCenterLines = [];

        if (options.enableCenterLine) {
            generatedCenterLines = convertSelectedRectanglesToCenterLines(selectedItems, options);
        } else {
            collectLinesRecursive(selectedItems, generatedCenterLines);
        }

        return generatedCenterLines;
    }

    /* 選択中の長方形を中心線化して返す
       Convert selected rectangles to center lines and return them */
    function convertSelectedRectanglesToCenterLines(selectedItems, options) {
        var generatedCenterLines = [];
        var rectangleItemsToProcess = [];
        collectRectsRecursive(selectedItems, rectangleItemsToProcess);

        for (var rectangleIndex = 0; rectangleIndex < rectangleItemsToProcess.length; rectangleIndex++) {
            var rectangleItem = rectangleItemsToProcess[rectangleIndex];
            var placeTarget = getCenterLinePlacementTarget(rectangleItem);
            var newLine = convertRectToCenterLine(rectangleItem, options.correctRotation, options.minShortSidePt);

            if (newLine) {
                if (placeTarget) {
                    movePageItemSafely(newLine, placeTarget, ElementPlacement.PLACEATEND);
                }
                generatedCenterLines.push(newLine);
            }
        }

        return generatedCenterLines;
    }

    /* 中心線の配置先を取得（CompoundPathItem 内サブパスは親へ配置）
       Get placement target for a generated center line (subpaths in compounds use compound parent) */
    function getCenterLinePlacementTarget(rectangleItem) {
        var placeTarget = rectangleItem.parent;
        try {
            if (placeTarget && placeTarget.typename === "CompoundPathItem") {
                placeTarget = placeTarget.parent;
            }
        } catch (e) {
            placeTarget = null;
        }
        return placeTarget;
    }

    /* 生成済み中心線を連結モードに応じてクラスタ単位で処理
       Process generated center lines cluster-by-cluster according to connect mode */
    function processCenterLineClusters(generatedCenterLines, options) {
        if (!generatedCenterLines || generatedCenterLines.length === 0) return generatedCenterLines;

        var commonStrokeWidth = options.commonStroke
            ? getRepresentativeStrokeWidth(generatedCenterLines, options.strokeStrategy, options.customStrokeWidth)
            : null;
        var clusterTolerance = (options.connectMode === "grid")
            ? computeGridClusterTolerance(generatedCenterLines)
            : 0;
        var clusters = clusterLinesByIntersection(generatedCenterLines, clusterTolerance);
        var resultLines = [];

        for (var clusterIndex = 0; clusterIndex < clusters.length; clusterIndex++) {
            appendProcessedCluster(resultLines, clusters[clusterIndex], options, commonStrokeWidth);
        }

        return resultLines;
    }

    /* 1つのクラスタを処理し、結果配列へ追加
       Process one cluster and append its result to the output array */
    function appendProcessedCluster(resultLines, cluster, options, commonStrokeWidth) {
        var combined = null;
        var gridAligned = null;

        if (options.connectMode === "connect") {
            combined = connectClusterByLineCount(cluster, options, commonStrokeWidth);
        } else if (options.connectMode === "grid") {
            gridAligned = alignClusterAsGridResult(cluster, options);
        }

        if (combined) {
            resultLines.push(combined);
        } else if (gridAligned) {
            for (var gridLineIndex = 0; gridLineIndex < gridAligned.length; gridLineIndex++) {
                resultLines.push(gridAligned[gridLineIndex]);
            }
        } else {
            for (var clusterLineIndex = 0; clusterLineIndex < cluster.length; clusterLineIndex++) {
                resultLines.push(cluster[clusterLineIndex]);
            }
        }
    }

    /* 本数に応じてクラスタを結合
       Connect a cluster according to its line count */
    function connectClusterByLineCount(cluster, options, commonStrokeWidth) {
        if (cluster.length === 4) {
            return tryConnectFourLinesIntoRect(cluster, options.strokeStrategy, options.customStrokeWidth);
        }
        if (cluster.length === 3) {
            return tryConnectThreeLinesIntoUShape(cluster, options.strokeStrategy, options.customStrokeWidth);
        }
        if (cluster.length === 2) {
            return tryConnectTwoLinesIntoLShape(cluster, options.strokeStrategy, options.customStrokeWidth);
        }
        if (cluster.length >= 5) {
            return tryConnectManyLinesIntoOutline(cluster, options.strokeStrategy, commonStrokeWidth, options.customStrokeWidth);
        }
        return null;
    }

    /* クラスタを格子化し、必要に応じて外枠を長方形化
       Align a cluster as a grid and optionally convert its outer frame to a rectangle */
    function alignClusterAsGridResult(cluster, options) {
        var gridAligned = tryAlignClusterAsGrid(cluster);
        if (gridAligned && options.outerRect) {
            var withFrame = convertOuterFrameToRect(gridAligned, options.strokeStrategy, options.customStrokeWidth);
            if (withFrame) gridAligned = withFrame;
        }
        return gridAligned;
    }

    /* 線幅・色などの仕上げ処理を適用
       Apply finishing options such as stroke width and print black */
    function applyFinishingOptions(generatedCenterLines, options) {
        if (options.commonStroke || options.strokeStrategy === "custom") {
            applyCommonStrokeWidth(generatedCenterLines, options.strokeStrategy, options.customStrokeWidth);
        }
        if (options.printBlack) {
            applyPrintBlackStroke(generatedCenterLines);
        }
    }

    /* 結果を選択状態にし、必要に応じて1つのグループにまとめる
       Select results and optionally wrap them into one group */
    function selectOrGroupResults(generatedCenterLines, options) {
        if (options.groupResult && options.connectMode === "grid" && generatedCenterLines.length > 0) {
            var resultGroup = groupGeneratedCenterLines(generatedCenterLines);
            app.activeDocument.selection = [resultGroup];
        } else {
            app.activeDocument.selection = generatedCenterLines;
        }
    }

    /* 生成結果を1つのグループへまとめる
       Wrap generated results into one group */
    function groupGeneratedCenterLines(generatedCenterLines) {
        var groupParent = getGeneratedResultGroupParent(generatedCenterLines);
        var resultGroup = groupParent.groupItems.add();

        for (var groupItemIndex = generatedCenterLines.length - 1; groupItemIndex >= 0; groupItemIndex--) {
            movePageItemSafely(generatedCenterLines[groupItemIndex], resultGroup, ElementPlacement.PLACEATBEGINNING);
        }

        return resultGroup;
    }

    /* 生成結果をまとめる親レイヤー／親グループを決定
       Determine parent layer/group for wrapping generated results */
    function getGeneratedResultGroupParent(generatedCenterLines) {
        var groupParent = app.activeDocument.activeLayer;
        try {
            var firstParent = generatedCenterLines[0].parent;
            if (firstParent && firstParent.typename === "CompoundPathItem") firstParent = firstParent.parent;
            if (firstParent && (firstParent.typename === "Layer" || firstParent.typename === "GroupItem")) {
                groupParent = firstParent;
            }
        } catch (e) { }
        return groupParent;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
        try {
            if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
                alert(getLabel('alertNoSelection'));
                return;
            }

            var options = showOptionDialog();
            if (options === null) return;

            var selectedItems = app.activeDocument.selection;
            var generatedCenterLines = buildTargetCenterLines(selectedItems, options);

            if (generatedCenterLines.length > 0) {
                generatedCenterLines = processCenterLineClusters(generatedCenterLines, options);
                applyFinishingOptions(generatedCenterLines, options);
                selectOrGroupResults(generatedCenterLines, options);
            }
        } catch (e) {
            alert(labelText('alertError') + "\n" + e);
        }
    }

    main();
})();