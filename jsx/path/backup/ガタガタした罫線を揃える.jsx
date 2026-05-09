#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "長方形を線に変換し、連結処理", en: "Convert Rectangles to Lines and Connect" },
    pnlOption: { ja: "オプション", en: "Options" },
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
// カラー / Colors
// =========================================

function makePrintBlack() {
    var blackColor = new CMYKColor();
    blackColor.cyan = 0;
    blackColor.magenta = 0;
    blackColor.yellow = 0;
    blackColor.black = 100;
    return blackColor;
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

    /* オプションパネル（1カラム）/ Options panel (single column) */
    var pnlOpt = leftCol.add("panel", undefined, L('pnlOption'));
    pnlOpt.orientation = "column";
    pnlOpt.alignChildren = "fill";
    pnlOpt.margins = PANEL_MARGINS;
    pnlOpt.spacing = 8;

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

    var dlgResult = dlg.show();

    if (dlgResult !== 1) return null;

    var strokeStrategy = rbStrokeMin.value ? "min" : (rbStrokeAvg.value ? "avg" : "max");

    return {
        correctRotation: true,
        printBlack: cbPrintBlack.value,
        commonStroke: cbCommonStroke.value,
        strokeStrategy: strokeStrategy
    };
}

// =========================================
// 後処理ユーティリティ / Post-processing Utilities
// =========================================

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

/* 5本以上の中心線をグループ化し、Pathfinder Divide で交差点分割した状態で返す
   Group 5+ center lines and apply Pathfinder Divide as a Live Effect, then return the group */
function tryConnectManyLinesIntoOutline(lines) {
    if (!lines || lines.length < 5) return null;

    var doc = app.activeDocument;

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

    /* 4. 新しい線をアピアランスに追加 / Add a new stroke to the appearance */
    app.executeMenuCommand('Adobe New Stroke Shortcut');

    /* 5. アピアランスを展開 / Expand appearance */
    app.executeMenuCommand('expandStyle');

    return group;
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

        /* 選択をスナップショット / Snapshot selection */
        var selectedItems = [];
        for (var i = 0; i < app.activeDocument.selection.length; i++) {
            selectedItems.push(app.activeDocument.selection[i]);
        }

        var resultItems = [];
        if (selectedItems.length >= 5) {
            var combined = tryConnectManyLinesIntoOutline(selectedItems);
            if (combined) {
                resultItems.push(combined);
            } else {
                for (var j = 0; j < selectedItems.length; j++) resultItems.push(selectedItems[j]);
            }
        } else {
            for (var k = 0; k < selectedItems.length; k++) resultItems.push(selectedItems[k]);
        }

        if (opts.commonStroke) {
            applyCommonStrokeWidth(resultItems, opts.strokeStrategy);
        }
        if (opts.printBlack) {
            applyPrintBlackStroke(resultItems);
        }
        app.activeDocument.selection = resultItems;
    } catch (e) {
        alert(labelText('alertError') + "\n" + e);
    }
}

main();
