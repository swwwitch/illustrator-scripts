#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * 概要：
 * 選択した塗りオブジェクトに対して、元の塗り色を始点にした線形グラデーションを作成します。
 * 単色オブジェクトを選択している場合はその塗り色を使用し、オブジェクトが1つで塗りがグラデーションの場合は、黒・白・透明を除いたグラデーションストップから始点色を選択できます。
 * 複数オブジェクト選択時は始点カラーパネル全体を無効化し、単一グラデーション選択時はドロップダウン内に実際の色情報を表示します。Auto は先頭の有効な候補色を使用し、同じ色のグラデーションストップは重複候補として扱わず、1色にまとめます。
 * 終点カラーと角度を2カラムで表示し、角度は 0、30、45、60、90 のラジオボタンから選択できます。
 * 終点カラーには、黒、白、透明、補色を選択できます。
 * セパレートグラデーション、反転、プレビューに対応し、キャンセル時には元の塗り色へ戻します。
 * CompoundPathItem、GroupItem 内の再帰処理、クリッピンググループ内のオブジェクトにも対応します。
 * グラデーション角度はオブジェクトごとに前回角度との差分で適用します。
 *
 * Overview:
 * Creates a linear gradient for selected filled objects, using the original fill color as the starting color.
 * For solid-color objects, the object's fill color is used directly. When exactly one gradient-filled object is selected, the source color can be chosen from gradient stops excluding black, white, and fully transparent stops.
 * When multiple objects are selected, the entire source-color panel is disabled. For a single gradient-filled object, the dropdown shows actual color information for each available source color. Auto uses the first valid source color, and duplicate stop colors are merged into a single option.
 * The end color and angle are shown in a two-column layout, and the angle can be selected with radio buttons for 0, 30, 45, 60, or 90 degrees.
 * The end color can be set to black, white, transparent, or a complementary color.
 * Supports separate gradients, reverse, and preview, and restores the original fill color when canceled.
 * Compound paths, recursive traversal inside groups, and objects inside clipping groups are also supported.
 * The gradient angle is applied per object using the difference from the previously applied angle.
 */

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
        ja: "グラデーション作成",
        en: "Create Gradient"
    },
    selectObjectAlert: {
        ja: "オブジェクトを選択してください。",
        en: "Please select an object."
    },
    endpointColorPanel: {
        ja: "終点のカラー",
        en: "End Color"
    },
    anglePanel: {
        ja: "角度",
        en: "Angle"
    },
    sourceColorPanel: {
        ja: "始点カラー",
        en: "Source Color"
    },
    auto: {
        ja: "自動（先頭）",
        en: "Auto (First)"
    },
    black: {
        ja: "黒",
        en: "Black"
    },
    white: {
        ja: "白",
        en: "White"
    },
    transparent: {
        ja: "透明",
        en: "Transparent"
    },
    complementary: {
        ja: "補色",
        en: "Complementary"
    },
    optionsPanel: {
        ja: "オプション",
        en: "Options"
    },
    separateGradient: {
        ja: "セパレートグラデーション",
        en: "Separate Gradient"
    },
    reverse: {
        ja: "反転",
        en: "Reverse"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    unsupportedSpotComplementary: {
        ja: "スポットカラーの補色計算には未対応です。",
        en: "Complementary color calculation is not supported for spot colors."
    }
};

function L(key) {
    return LABELS[key][lang];
}

function main() {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;
    var originalSelection = [];
    for (var selIndex = 0; selIndex < sel.length; selIndex++) {
        originalSelection.push(sel[selIndex]);
    }

    if (sel.length === 0) {
        alert(L("selectObjectAlert"));
        return;
    }

    var isCMYK = isDocumentCMYK(doc);
    var targetObjects = [];
    var shouldRestoreOriginal = false;
    var shouldCleanupAppliedGradients = false;

    try {
        /* 選択オブジェクトの情報を保持 / Store selected object data */
        targetObjects = collectGradientTargets(sel, isCMYK);

        if (targetObjects.length === 0) {
            return;
        }

        // =========================================
        // ダイアログの作成 / Create dialog
        // =========================================
        var ui = buildDialogUI();
        var dlg = ui.dlg;
        var radioBlack = ui.radioBlack;
        var radioWhite = ui.radioWhite;
        var radioTransparent = ui.radioTransparent;
        var radioComplementary = ui.radioComplementary;
        var angle0 = ui.angle0;
        var angle30 = ui.angle30;
        var angle45 = ui.angle45;
        var angle60 = ui.angle60;
        var angle90 = ui.angle90;
        var sourceDropdown = ui.sourceDropdown;
        var sourcePanel = ui.sourcePanel;
        var chkSeparate = ui.chkSeparate;
        var chkReverse = ui.chkReverse;
        var chkPreview = ui.chkPreview;

        populateSourceDropdown(sourceDropdown, targetObjects, isCMYK);
        updateSourcePanelEnabled(sourcePanel, sourceDropdown, targetObjects);


        // =========================================
        // 処理関数 / Processing functions
        // =========================================

        function applyGradient() {
            for (var i = 0; i < targetObjects.length; i++) {
                var data = targetObjects[i];
                var obj = data.item;
                var orgColor = getSourceColorForFill(data.originalFillColor, isCMYK, sourceDropdown);
                var targetColor = orgColor;
                var targetOpacity = 100.0;

                if (radioTransparent.value) {
                    targetColor = orgColor;
                    targetOpacity = 0.0;
                } else if (radioBlack.value) {
                    targetColor = createBlackColor(isCMYK);
                } else if (radioWhite.value) {
                    targetColor = createWhiteColor(isCMYK);
                } else if (radioComplementary.value) {
                    targetColor = createComplementaryColor(orgColor, isCMYK);
                }

                var startColor = orgColor;
                var startOpacity = 100.0;
                var endColor = targetColor;
                var endOpacity = targetOpacity;

                if (chkReverse.value) {
                    startColor = targetColor;
                    startOpacity = targetOpacity;
                    endColor = orgColor;
                    endOpacity = 100.0;
                }

                /* グラデーションの新規作成または再利用 / Create or reuse gradient */
                if (data.appliedGrad === null) {
                    data.appliedGrad = doc.gradients.add();
                    data.appliedGrad.type = GradientType.LINEAR;
                }
                var activeGrad = data.appliedGrad;

                var isSeparate = chkSeparate.value;
                var requiredStops = isSeparate ? 4 : 2;

                while (activeGrad.gradientStops.length < requiredStops) {
                    activeGrad.gradientStops.add();
                }
                while (activeGrad.gradientStops.length > requiredStops) {
                    activeGrad.gradientStops[activeGrad.gradientStops.length - 1].remove();
                }

                if (isSeparate) {
                    /* セパレート（0, 50, 50, 100） / Separate stops (0, 50, 50, 100) */
                    setStop(activeGrad.gradientStops[0], 0, startColor, startOpacity);
                    setStop(activeGrad.gradientStops[1], 50.0, startColor, startOpacity);
                    setStop(activeGrad.gradientStops[2], 50.0, endColor, endOpacity);
                    setStop(activeGrad.gradientStops[3], 100.0, endColor, endOpacity);
                } else {
                    /* 通常（0, 100） / Standard stops (0, 100) */
                    setStop(activeGrad.gradientStops[0], 0, startColor, startOpacity);
                    setStop(activeGrad.gradientStops[1], 100.0, endColor, endOpacity);
                }

                var gradColor = new GradientColor();
                gradColor.gradient = activeGrad;
                setTargetFillColor(obj, gradColor);

                doc.selection = null;
                obj.selected = true;
                applyGradientAngle(obj, gradColor, data, angle0, angle30, angle45, angle60, angle90);
            }
        }

        function setStop(stop, ramp, color, opacity) {
            stop.rampPoint = ramp;
            stop.color = color;
            stop.opacity = opacity;
        }

        function restoreOriginal() {
            for (var i = 0; i < targetObjects.length; i++) {
                var data = targetObjects[i];
                setTargetFillColor(data.item, data.originalFillColor);
                data.lastAngle = 0;
            }
        }

        function cleanupAppliedGradients() {
            for (var i = 0; i < targetObjects.length; i++) {
                var data = targetObjects[i];
                if (data.appliedGrad !== null) {
                    try { data.appliedGrad.remove(); } catch (e) { }
                }
                data.appliedGrad = null;
                data.lastAngle = 0;
            }
        }

        function restoreSelection() {
            doc.selection = null;
            for (var i = 0; i < originalSelection.length; i++) {
                try {
                    originalSelection[i].selected = true;
                } catch (e) { }
            }
        }

        function updatePreview() {
            if (chkPreview.value) {
                applyGradient();
            } else {
                restoreOriginal();
            }
            restoreSelection();
            try {
                app.redraw();
            } catch (e) { }
        }

        /* イベントの設定 / Set event handlers */
        chkPreview.onClick = updatePreview;
        chkSeparate.onClick = updatePreview;
        chkReverse.onClick = updatePreview;
        angle0.onClick = updatePreview;
        angle30.onClick = updatePreview;
        angle45.onClick = updatePreview;
        angle60.onClick = updatePreview;
        angle90.onClick = updatePreview;
        sourceDropdown.onChange = updatePreview;
        radioBlack.onClick = updatePreview;
        radioWhite.onClick = updatePreview;
        radioTransparent.onClick = updatePreview;
        radioComplementary.onClick = updatePreview;
        addColorKeyHandler(dlg, radioBlack, radioWhite, radioTransparent, radioComplementary, updatePreview);

        if (dlg.show() === 1) {
            if (!chkPreview.value) {
                applyGradient();
            }
            shouldCleanupAppliedGradients = false;
            shouldRestoreOriginal = false;
            restoreSelection();
        } else {
            shouldRestoreOriginal = true;
            shouldCleanupAppliedGradients = true;
            restoreSelection();
        }
    } finally {
        if (shouldRestoreOriginal) {
            try {
                restoreOriginal();
            } catch (e) { }
        }
        if (shouldCleanupAppliedGradients) {
            try {
                cleanupAppliedGradients();
            } catch (e) { }
        }
        try {
            restoreSelection();
        } catch (e) { }
        try {
            app.redraw();
        } catch (e) { }
    }
}

// =========================================
// UI構築 / UI construction
// =========================================

function buildDialogUI() {
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    var topGroup = dlg.add("group");
    topGroup.orientation = "row";
    topGroup.alignChildren = ["fill", "top"];

    /* 終点カラーの設定 / Set end color options */
    var panel = topGroup.add("panel", undefined, L("endpointColorPanel"));
    panel.orientation = "column";
    panel.alignChildren = ["left", "top"];
    panel.margins = [15, 20, 15, 10];

    var radioBlack = panel.add("radiobutton", undefined, L("black"));
    var radioWhite = panel.add("radiobutton", undefined, L("white"));
    var radioTransparent = panel.add("radiobutton", undefined, L("transparent"));
    var radioComplementary = panel.add("radiobutton", undefined, L("complementary"));

    radioTransparent.value = true; // デフォルト / Default

    var anglePanel = topGroup.add("panel", undefined, L("anglePanel"));
    anglePanel.orientation = "column";
    anglePanel.alignChildren = ["left", "top"];
    anglePanel.margins = [15, 20, 15, 10];

    var angle0 = anglePanel.add("radiobutton", undefined, "0");
    var angle30 = anglePanel.add("radiobutton", undefined, "30");
    var angle45 = anglePanel.add("radiobutton", undefined, "45");
    var angle60 = anglePanel.add("radiobutton", undefined, "60");
    var angle90 = anglePanel.add("radiobutton", undefined, "90");
    angle0.value = true; // デフォルト / Default

    var sourcePanel = dlg.add("panel", undefined, L("sourceColorPanel"));
    sourcePanel.orientation = "row";
    sourcePanel.alignChildren = ["left", "center"];
    sourcePanel.margins = [15, 20, 15, 10];

    var sourceDropdown = sourcePanel.add("dropdownlist", undefined, [L("auto")]);
    sourceDropdown.selection = 0; // デフォルト / Default

    /* オプションの設定 / Set options */
    var optPanel = dlg.add("panel", undefined, L("optionsPanel"));
    optPanel.orientation = "column";
    optPanel.alignChildren = ["left", "top"];
    optPanel.margins = [15, 20, 15, 10];

    var chkSeparate = optPanel.add("checkbox", undefined, L("separateGradient"));
    var chkReverse = optPanel.add("checkbox", undefined, L("reverse"));
    var chkPreview = optPanel.add("checkbox", undefined, L("preview"));

    /* ボタンの設定 / Set button layout */
    var btnGroup = dlg.add("group");
    btnGroup.alignment = ["center", "center"]; // 中央揃え / Center align
    btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    btnGroup.add("button", undefined, L("ok"), { name: "ok" });

    return {
        dlg: dlg,
        radioBlack: radioBlack,
        radioWhite: radioWhite,
        radioTransparent: radioTransparent,
        radioComplementary: radioComplementary,
        angle0: angle0,
        angle30: angle30,
        angle45: angle45,
        angle60: angle60,
        angle90: angle90,
        sourcePanel: sourcePanel,
        sourceDropdown: sourceDropdown,
        chkSeparate: chkSeparate,
        chkReverse: chkReverse,
        chkPreview: chkPreview
    };
}

// =========================================
// 補助関数 / Helper functions
// =========================================

function populateSourceDropdown(sourceDropdown, targetObjects, isCMYK) {
    if (!sourceDropdown) {
        return;
    }

    removeAllDropdownItems(sourceDropdown);

    var gradientFill = getSingleGradientFillColor(targetObjects);
    if (!gradientFill) {
        sourceDropdown.add("item", L("auto"));
        sourceDropdown.selection = 0;
        return;
    }

    var sourceColors = collectSourceColorsFromGradient(gradientFill, isCMYK);
    if (sourceColors.length === 0) {
        sourceDropdown.add("item", L("auto"));
        sourceDropdown.selection = 0;
        return;
    }

    sourceDropdown.add("item", L("auto"));
    for (var i = 0; i < sourceColors.length; i++) {
        sourceDropdown.add("item", formatSourceColorLabel(sourceColors[i], i));
    }
    sourceDropdown.selection = 0;
}

function removeAllDropdownItems(dropdown) {
    while (dropdown.items.length > 0) {
        dropdown.remove(dropdown.items[0]);
    }
}

function getSingleGradientFillColor(targetObjects) {
    if (!targetObjects || targetObjects.length !== 1) {
        return null;
    }

    var fillColor = targetObjects[0].originalFillColor;
    if (fillColor && fillColor.typename === "GradientColor") {
        return fillColor;
    }
    return null;
}

function updateSourcePanelEnabled(sourcePanel, sourceDropdown, targetObjects) {
    var isEnabled = !!getSingleGradientFillColor(targetObjects);

    if (sourcePanel) {
        sourcePanel.enabled = isEnabled;
    }
    if (sourceDropdown) {
        sourceDropdown.enabled = isEnabled;
    }
}

function isDocumentCMYK(doc) {
    return doc.documentColorSpace === DocumentColorSpace.CMYK;
}
function isGradientTarget(obj) {
    if (!obj) {
        return false;
    }

    if (obj.typename === "PathItem") {
        return obj.filled && !obj.clipping;
    }

    if (obj.typename === "CompoundPathItem") {
        return isFilledCompoundPath(obj);
    }

    return false;
}

function isFilledCompoundPath(compoundPath) {
    if (!compoundPath || !compoundPath.pathItems || compoundPath.pathItems.length === 0) {
        return false;
    }

    for (var i = 0; i < compoundPath.pathItems.length; i++) {
        var pathItem = compoundPath.pathItems[i];
        if (pathItem.filled && !pathItem.clipping) {
            return true;
        }
    }

    return false;
}

function createGradientTargetRecord(obj, isCMYK) {
    return {
        item: obj,
        originalFillColor: getTargetFillColor(obj),
        appliedGrad: null,
        lastAngle: 0
    };
}

function collectGradientTargets(selection, isCMYK) {
    var targetObjects = [];
    for (var i = 0; i < selection.length; i++) {
        collectGradientTargetsFromItem(selection[i], targetObjects, isCMYK);
    }
    return targetObjects;
}

function collectGradientTargetsFromItem(item, targetObjects, isCMYK) {
    if (!item) {
        return;
    }

    if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            collectGradientTargetsFromItem(item.pageItems[i], targetObjects, isCMYK);
        }
        return;
    }

    if (isGradientTarget(item)) {
        targetObjects.push(createGradientTargetRecord(item, isCMYK));
    }
}

function getTargetFillColor(obj) {
    if (!obj) {
        return null;
    }

    if (obj.typename === "CompoundPathItem") {
        return getCompoundPathFillColor(obj);
    }

    return obj.fillColor;
}

function setTargetFillColor(obj, fillColor) {
    if (!obj) {
        return;
    }

    if (obj.typename === "CompoundPathItem") {
        setCompoundPathFillColor(obj, fillColor);
        return;
    }

    obj.fillColor = fillColor;
}

function getCompoundPathFillColor(compoundPath) {
    if (!compoundPath || !compoundPath.pathItems || compoundPath.pathItems.length === 0) {
        return null;
    }

    for (var i = 0; i < compoundPath.pathItems.length; i++) {
        var pathItem = compoundPath.pathItems[i];
        if (pathItem.filled && !pathItem.clipping) {
            return pathItem.fillColor;
        }
    }

    return compoundPath.pathItems[0].fillColor;
}

function setCompoundPathFillColor(compoundPath, fillColor) {
    if (!compoundPath || !compoundPath.pathItems || compoundPath.pathItems.length === 0) {
        return;
    }

    for (var i = 0; i < compoundPath.pathItems.length; i++) {
        var pathItem = compoundPath.pathItems[i];
        if (!pathItem.clipping) {
            pathItem.fillColor = fillColor;
        }
    }
}

function getSourceColorForFill(fillColor, isCMYK, sourceDropdown) {
    if (!fillColor || fillColor.typename !== "GradientColor") {
        return cloneSimpleColor(fillColor, isCMYK);
    }

    var sourceColors = collectSourceColorsFromGradient(fillColor, isCMYK);
    if (sourceColors.length === 0) {
        return getFallbackGradientColor(fillColor, isCMYK);
    }

    var selectedIndex = getSelectedSourceColorIndex(sourceDropdown);
    if (selectedIndex < 0 || selectedIndex >= sourceColors.length) {
        selectedIndex = 0;
    }

    return cloneSimpleColor(sourceColors[selectedIndex], isCMYK);
}

function collectSourceColorsFromGradient(fillColor, isCMYK) {
    var colors = [];
    var grad = fillColor.gradient;
    if (!grad || !grad.gradientStops) {
        return colors;
    }

    for (var i = 0; i < grad.gradientStops.length; i++) {
        var stop = grad.gradientStops[i];
        if (!isExcludedSourceStop(stop)) {
            var clonedColor = cloneSimpleColor(stop.color, isCMYK);
            if (!containsEquivalentColor(colors, clonedColor)) {
                colors.push(clonedColor);
            }
        }
    }
    return colors;
}

function containsEquivalentColor(colors, targetColor) {
    for (var i = 0; i < colors.length; i++) {
        if (isSameColorValue(colors[i], targetColor)) {
            return true;
        }
    }
    return false;
}

function isSameColorValue(colorA, colorB) {
    if (!colorA || !colorB) {
        return false;
    }

    if (colorA.typename !== colorB.typename) {
        return false;
    }

    if (colorA.typename === "CMYKColor") {
        return isSameNumber(colorA.cyan, colorB.cyan) &&
            isSameNumber(colorA.magenta, colorB.magenta) &&
            isSameNumber(colorA.yellow, colorB.yellow) &&
            isSameNumber(colorA.black, colorB.black);
    }

    if (colorA.typename === "RGBColor") {
        return isSameNumber(colorA.red, colorB.red) &&
            isSameNumber(colorA.green, colorB.green) &&
            isSameNumber(colorA.blue, colorB.blue);
    }

    if (colorA.typename === "GrayColor") {
        return isSameNumber(colorA.gray, colorB.gray);
    }

    if (colorA.typename === "SpotColor") {
        var spotNameA = (colorA.spot && colorA.spot.name) ? colorA.spot.name : "";
        var spotNameB = (colorB.spot && colorB.spot.name) ? colorB.spot.name : "";
        return spotNameA === spotNameB && isSameNumber(colorA.tint, colorB.tint);
    }

    return false;
}

function isSameNumber(a, b) {
    var na = Number(a);
    var nb = Number(b);
    if (isNaN(na) || isNaN(nb)) {
        return false;
    }
    return Math.abs(na - nb) < 0.001;
}

function getSelectedSourceColorIndex(sourceDropdown) {
    if (!sourceDropdown || !sourceDropdown.selection) {
        return 0;
    }

    if (sourceDropdown.selection.index <= 0) {
        return 0;
    }

    return Math.max(0, sourceDropdown.selection.index - 1);
}

function formatSourceColorLabel(color, index) {
    return String(index + 1) + ': ' + formatColorLabel(color);
}

function formatColorLabel(color) {
    if (!color) {
        return '-';
    }

    if (color.typename === "CMYKColor") {
        return 'C' + formatColorNumber(color.cyan) + ' M' + formatColorNumber(color.magenta) + ' Y' + formatColorNumber(color.yellow) + ' K' + formatColorNumber(color.black);
    }

    if (color.typename === "RGBColor") {
        return 'R' + formatColorNumber(color.red) + ' G' + formatColorNumber(color.green) + ' B' + formatColorNumber(color.blue);
    }

    if (color.typename === "GrayColor") {
        return 'Gray ' + formatColorNumber(color.gray);
    }

    if (color.typename === "SpotColor") {
        var spotName = (color.spot && color.spot.name) ? color.spot.name : 'Spot';
        return spotName + ' ' + formatColorNumber(color.tint) + '%';
    }

    return color.typename;
}

function formatColorNumber(value) {
    if (typeof value !== "number") {
        return '0';
    }

    var rounded = Math.round(value * 10) / 10;
    if (Math.abs(rounded - Math.round(rounded)) < 0.001) {
        return String(Math.round(rounded));
    }
    return String(rounded);
}

function isExcludedSourceStop(stop) {
    if (!stop) {
        return true;
    }
    if (typeof stop.opacity === "number" && stop.opacity <= 0) {
        return true;
    }
    return isBlackOrWhiteColor(stop.color);
}

function isBlackOrWhiteColor(color) {
    if (!color) {
        return false;
    }

    if (color.typename === "RGBColor") {
        var isBlackRgb = (color.red === 0 && color.green === 0 && color.blue === 0);
        var isWhiteRgb = (color.red === 255 && color.green === 255 && color.blue === 255);
        return isBlackRgb || isWhiteRgb;
    }

    if (color.typename === "CMYKColor") {
        var isBlackCmyk = (color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 100);
        var isWhiteCmyk = (color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0);
        return isBlackCmyk || isWhiteCmyk;
    }

    if (color.typename === "GrayColor") {
        return color.gray === 0 || color.gray === 100;
    }

    return false;
}

function getFallbackGradientColor(fillColor, isCMYK) {
    var grad = fillColor.gradient;
    if (grad && grad.gradientStops && grad.gradientStops.length > 0) {
        return cloneSimpleColor(grad.gradientStops[0].color, isCMYK);
    }
    return createBlackColor(isCMYK);
}

function cloneSimpleColor(color, isCMYK) {
    if (!color) {
        return createBlackColor(isCMYK);
    }

    if (color.typename === "CMYKColor") {
        var cmyk = new CMYKColor();
        cmyk.cyan = color.cyan;
        cmyk.magenta = color.magenta;
        cmyk.yellow = color.yellow;
        cmyk.black = color.black;
        return cmyk;
    }

    if (color.typename === "RGBColor") {
        var rgb = new RGBColor();
        rgb.red = color.red;
        rgb.green = color.green;
        rgb.blue = color.blue;
        return rgb;
    }

    if (color.typename === "GrayColor") {
        var gray = new GrayColor();
        gray.gray = color.gray;
        return gray;
    }

    if (color.typename === "SpotColor") {
        var spot = new SpotColor();
        spot.spot = color.spot;
        spot.tint = color.tint;
        return spot;
    }

    return createBlackColor(isCMYK);
}

function addColorKeyHandler(dlg, blackRadio, whiteRadio, transparentRadio, complementaryRadio, onChange) {
    dlg.addEventListener("keydown", function (event) {
        if (event.keyName == "B") {
            blackRadio.value = true;
            event.preventDefault();
        } else if (event.keyName == "W") {
            whiteRadio.value = true;
            event.preventDefault();
        } else if (event.keyName == "T") {
            transparentRadio.value = true;
            event.preventDefault();
        } else if (event.keyName == "C") {
            complementaryRadio.value = true;
            event.preventDefault();
        } else {
            return;
        }

        onChange();
    });
}

function getSelectedAngle(angle0, angle30, angle45, angle60, angle90) {
    if (angle30.value) {
        return 30;
    }
    if (angle45.value) {
        return 45;
    }
    if (angle60.value) {
        return 60;
    }
    if (angle90.value) {
        return 90;
    }
    return 0;
}

function applyGradientAngle(obj, gradColor, data, angle0, angle30, angle45, angle60, angle90) {
    var angle = getSelectedAngle(angle0, angle30, angle45, angle60, angle90);
    var previousAngle = (typeof data.lastAngle === "number") ? data.lastAngle : 0;
    var deltaAngle = angle - previousAngle;

    gradColor.angle = angle;

    if (deltaAngle !== 0) {
        rotateTargetGradient(obj, deltaAngle);
    }

    data.lastAngle = angle;
}

function rotateTargetGradient(obj, deltaAngle) {
    if (!obj || deltaAngle === 0) {
        return;
    }

    if (obj.typename === "CompoundPathItem") {
        rotateCompoundPathGradient(obj, deltaAngle);
        return;
    }

    obj.rotate(deltaAngle, false, false, true, false, Transformation.CENTER);
}

function rotateCompoundPathGradient(compoundPath, deltaAngle) {
    if (!compoundPath || !compoundPath.pathItems || compoundPath.pathItems.length === 0) {
        return;
    }

    compoundPath.rotate(deltaAngle, false, false, true, false, Transformation.CENTER);
}


function createBlackColor(isCMYK) {
    if (isCMYK) {
        var c = new CMYKColor();
        c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100;
        return c;
    } else {
        var c = new RGBColor();
        c.red = 0; c.green = 0; c.blue = 0;
        return c;
    }
}

function createWhiteColor(isCMYK) {
    if (isCMYK) {
        var c = new CMYKColor();
        c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 0;
        return c;
    } else {
        var c = new RGBColor();
        c.red = 255; c.green = 255; c.blue = 255;
        return c;
    }
}


function createComplementaryColor(orgColor, isCMYK) {
    if (orgColor.typename === "CMYKColor") {
        var c = new CMYKColor();
        c.cyan = 100 - orgColor.cyan;
        c.magenta = 100 - orgColor.magenta;
        c.yellow = 100 - orgColor.yellow;
        c.black = orgColor.black;
        return c;
    } else if (orgColor.typename === "RGBColor") {
        var c = new RGBColor();
        c.red = 255 - orgColor.red;
        c.green = 255 - orgColor.green;
        c.blue = 255 - orgColor.blue;
        return c;
    } else if (orgColor.typename === "GrayColor") {
        var c = new GrayColor();
        c.gray = 100 - orgColor.gray;
        return c;
    } else if (orgColor.typename === "SpotColor") {
        throw new Error(L("unsupportedSpotComplementary"));
    }
    return createBlackColor(isCMYK);
}

main();