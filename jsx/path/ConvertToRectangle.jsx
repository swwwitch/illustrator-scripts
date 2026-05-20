#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
## 概要

選択したオブジェクトの境界に合わせて長方形を作成する Illustrator スクリプトです。
作成単位は「オブジェクトごと」または「選択範囲全体」から選べます。
プレビュー境界での計測、テキストのアウトライン化計測、定規単位に連動したマージン（負値で内側）を指定できます。
ダイアログ表示中は選択オブジェクトの不透明度を一時的に 50% へ下げてプレビューを見やすくできます（不透明度を変えるプリセット選択時は自動で無効化）。
塗り・線プリセット（線幅は Illustrator の「線幅」単位に連動）、重ね順、元オブジェクトの扱い（残す／クリッピングマスクにする／削除）を設定できます。
クリッピングマスク化では元オブジェクトの重ね順を維持します。「選択範囲全体」では「前面／背面」に応じて最前面／最背面の項目を基準にします。
ロック・非表示オブジェクトは対象外です。プレビューで結果を確認しながら調整できます（マスク／削除は OK 後に反映）。
スクリプト実行中はバウンディングボックスとエッジ表示を一時的に切り替え、終了時（OK／キャンセル／エラーいずれも）に元の状態へ戻します。

### キーボード操作

マージン入力：↑↓ で ±1、Shift+↑↓ で ±10、Alt+↑↓ で ±0.1。
ショートカット：S/G（作成単位）、P/O/A/T（オプション）、F/B（重ね順）、N/M/D（元オブジェクト）。

## Overview

Creates Illustrator rectangles that match the bounds of selected objects.
The creation unit can be set to either per object or the whole selection.
You can measure preview bounds, measure text as outlines, and add a ruler-unit margin (negative shrinks).
While the dialog is open, the selected objects can be temporarily dimmed to 50% opacity to make the preview easier to see (auto-disabled when a preset that changes opacity is selected).
Choose a fill/stroke preset (stroke width follows the Stroke Units preference), rectangle order, and original handling (keep / make clipping mask / delete).
Clipping-mask mode preserves the original z-order. Whole-selection mode anchors the new rectangle to the top-/bottom-most item based on the front/back setting.
Locked and hidden objects are excluded. Preview shows the result while you adjust (mask/delete is applied on OK).
The bounding box and edge display are temporarily toggled during the script and restored on exit (OK, Cancel, or error).

### Keyboard

Margin field: ↑↓ ±1, Shift+↑↓ ±10, Alt+↑↓ ±0.1.
Shortcuts: S/G (creation unit), P/O/A/T (options), F/B (order), N/M/D (original).
*/

// ==============================
// スクリプト情報 / Script information
// ==============================
var SCRIPT_VERSION = "v1.0.1";

// ==============================
// 言語判定 / Language detection
// ==============================
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

// ==============================
// ラベル定義 / Label definitions
// ==============================
var LABELS = {
    /* Dialog / ダイアログ */
    dialogTitle: { ja: "長方形を作成", en: "Create Rectangle" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    preview: { ja: "プレビュー", en: "Preview" },
    previewTip: { ja: "オンで作成される長方形をその場で表示します（マスク／削除はOK後に反映）", en: "Shows the resulting rectangle while you adjust (mask/delete applied on OK)" },

    /* Panels / パネル */
    panelTarget: { ja: "作成単位", en: "Creation unit" },
    panelOptions: { ja: "オプション", en: "Options" },
    panelAppearance: { ja: "塗りと線", en: "Fill and stroke" },
    panelOrder: { ja: "長方形の重ね順", en: "Rectangle order" },
    panelOriginal: { ja: "元オブジェクトの扱い", en: "Original handling" },

    /* Target radios / 作成単位 */
    targetEach: { ja: "オブジェクトごと", en: "Per object" },
    targetEachTip: { ja: "選択中の各オブジェクトに対して、それぞれ長方形を作成します", en: "Creates a separate rectangle for each selected object" },
    targetGroup: { ja: "選択範囲全体", en: "Whole selection" },
    targetGroupTip: { ja: "選択中のオブジェクト全体を1つの境界として、長方形を1つ作成します", en: "Creates one rectangle using the bounds of the entire selection" },

    /* Options / オプション */
    previewBounds: { ja: "プレビュー境界で計測", en: "Use preview bounds" },
    boundsTip: { ja: "オンで線幅・効果を含む見た目のサイズを基準にします。オフではパス本体の境界を使います", en: "When on, measures the visible size including stroke and effects. When off, uses the path bounds" },
    outlineText: { ja: "テキストはアウトライン化して計測", en: "Measure text as outlines" },
    outlineTextTip: { ja: "元のテキストは変更せず、複製を一時的にアウトライン化して境界を計測します", en: "Measures a temporary outlined copy without modifying the original text" },
    useMargin: { ja: "マージンを追加", en: "Add margin" },
    marginTip: { ja: "オンにすると、計測した境界から指定値だけ外側へ広げます。単位は Illustrator の定規単位に連動します（負値で内側）", en: "When on, expands the measured bounds outward by this amount. The unit follows Illustrator's ruler unit (negative shrinks)" },
    dimSelection: { ja: "実行中は不透明度を下げる", en: "Dim selection while running" },
    dimSelectionTip: { ja: "ダイアログを開いている間、選択オブジェクトの不透明度を一時的に 50% にして、閉じたときに元に戻します", en: "Temporarily sets the selected objects to 50% opacity while the dialog is open, and restores it when the dialog closes" },

    /* Appearance / 塗りと線 */
    currentTip: { ja: "ドキュメントの現在の既定値（直前に使った塗り・線）をそのまま使います", en: "Uses the document's current default fill and stroke" },
    imageNotice: { ja: "画像では直前の塗り・線を取得できないため選択できません", en: "Current fill/stroke is not available for image items" },

    /* Order radios / 重ね順 */
    orderFront: { ja: "前面に作成", en: "Create in front" },
    orderFrontTip: { ja: "生成する長方形を元オブジェクトの前面に配置します。「元オブジェクトの扱い」が「残す」のときだけ有効です", en: "Places the new rectangle in front of the original. Available only when the original is kept" },
    orderBack: { ja: "背面に作成", en: "Create behind" },
    orderBackTip: { ja: "生成する長方形を元オブジェクトの背面に配置します。「元オブジェクトの扱い」が「残す」のときだけ有効です", en: "Places the new rectangle behind the original. Available only when the original is kept" },

    /* Original handling / 元オブジェクト */
    originalKeep: { ja: "残す", en: "Keep" },
    originalKeepTip: { ja: "元オブジェクトを残し、長方形だけを追加します", en: "Keeps the original object and adds the rectangle" },
    originalMask: { ja: "クリッピングマスクにする", en: "Make clipping mask" },
    maskTip: { ja: "生成した長方形をクリップパスとして、元オブジェクトをクリッピングマスク化します", en: "Uses the new rectangle as the clipping path for the original object" },
    originalDelete: { ja: "削除", en: "Delete" },
    originalDeleteTip: { ja: "長方形の作成後、元オブジェクトを削除します", en: "Deletes the original object after creating the rectangle" },

    /* Alerts / 警告 */
    noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    noSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." },
    noResult: { ja: "長方形を作成できるオブジェクトがありませんでした。", en: "No rectangles could be created." }
};

// ラベル取得 / Get label
function L(labelText) {
    var labelEntry = LABELS[labelText];
    if (!labelEntry) return labelText;
    return labelEntry[lang] || labelEntry.en || labelText;
}

// ==============================
// 塗り・線プリセット / Fill & stroke presets
// ==============================
// applyAppearance: true … 生成した長方形にこのプリセットの塗り・線を適用する
// applyAppearance: false（直前の塗り・線）… 適用せず、作成時の既定値のままにする
// strokeWidth: Illustrator の「線幅」環境設定単位で指定し、適用時に pt へ変換する
// opacity (省略時 100): プリセットで不透明度を変える場合に指定
var FILL_STROKE_PRESETS = [
    { ja: "塗り：なし、線：なし", en: "Fill: none, Stroke: none", applyAppearance: true, filled: false, stroked: false, strokeWidth: 0 },
    { ja: "線：黒、{width}{unit}", en: "Stroke: black, {width}{unit}", applyAppearance: true, filled: false, stroked: true, strokeWidth: 1 },
    { ja: "線：黒、{width}{unit}", en: "Stroke: black, {width}{unit}", applyAppearance: true, filled: false, stroked: true, strokeWidth: 0.25 },
    { ja: "塗り：黒、不透明度：40%", en: "Fill: black 40%, Stroke: none", applyAppearance: true, filled: true, stroked: false, strokeWidth: 0, opacity: 40 },
    { ja: "直前の塗り・線の情報", en: "Current fill / stroke", applyAppearance: false }
];
var DEFAULT_PRESET_INDEX = 1; // 既定：線：黒、1（線幅設定単位）

// 「直前の塗り・線の情報」プリセットの位置（applyAppearance: false）
var CURRENT_PRESET_INDEX = (function () {
    for (var i = 0; i < FILL_STROKE_PRESETS.length; i++) {
        if (!FILL_STROKE_PRESETS[i].applyAppearance) return i;
    }
    return -1;
})();

// ==============================
// 単位変換 / Unit conversion
// ==============================

// Illustrator の strokeUnits 設定値を取得する / Get Illustrator strokeUnits preference
function getStrokeUnitsPreference() {
    var strokeUnits = 2; // 既定：pt / Default: points
    try {
        strokeUnits = app.preferences.getIntegerPreference("strokeUnits");
    } catch (ignoreStrokeUnits) { }
    return strokeUnits;
}

// Illustrator の単位設定値を pt 変換係数に変換する / Convert Illustrator unit preference value to a points factor
function getUnitToPointFactor(unitType) {
    switch (unitType) {
        case 0: return 72;              // inch
        case 1: return 72 / 25.4;       // mm
        case 2: return 1;               // pt
        case 3: return 12;              // pica
        case 4: return 72 / 2.54;       // cm
        case 5: return 0.25;            // Q/H
        case 6: return 1;               // px
        default: return 1;
    }
}

// Illustrator の単位設定値を表示用単位ラベルに変換する / Convert Illustrator unit preference value to a display label
function getUnitLabel(unitType) {
    switch (unitType) {
        case 0: return "in";
        case 1: return "mm";
        case 2: return "pt";
        case 3: return "pc";
        case 4: return "cm";
        case 5: return "Q";
        case 6: return "px";
        default: return "pt";
    }
}

// Illustrator の strokeUnits 設定値を pt 変換係数に変換する / Convert Illustrator strokeUnits to a points factor
function getStrokeUnitToPointFactor() {
    return getUnitToPointFactor(getStrokeUnitsPreference());
}

// Illustrator の strokeUnits 設定値を表示用単位ラベルに変換する / Convert Illustrator strokeUnits to a display label
function getStrokeUnitLabel() {
    return getUnitLabel(getStrokeUnitsPreference());
}

// Illustrator の rulerType 設定値を取得する / Get Illustrator rulerType preference
function getRulerTypePreference() {
    var rulerType = 2; // 既定：pt / Default: points
    try {
        rulerType = app.preferences.getIntegerPreference("rulerType");
    } catch (ignoreRulerType) { }
    return rulerType;
}

// Illustrator の rulerType 設定値を pt 変換係数に変換する / Convert Illustrator rulerType to a points factor
function getRulerUnitToPointFactor() {
    return getUnitToPointFactor(getRulerTypePreference());
}

// Illustrator の rulerType 設定値を表示用単位ラベルに変換する / Convert Illustrator rulerType to a display label
function getRulerUnitLabel() {
    return getUnitLabel(getRulerTypePreference());
}

// 定規単位の値を pt に変換する / Convert ruler-unit value to points
function rulerUnitValueToPoints(value) {
    return value * getRulerUnitToPointFactor();
}

// 線幅設定単位の値を pt に変換する / Convert stroke-unit value to points
function strokeUnitValueToPoints(value) {
    return value * getStrokeUnitToPointFactor();
}

// ドキュメントのカラースペースに合わせた黒を返す / Returns black for the document color space
function makeBlack(isRgbDocument) {
    if (isRgbDocument) {
        var rgb = new RGBColor();
        rgb.red = rgb.green = rgb.blue = 0;
        return rgb;
    }
    var cmyk = new CMYKColor();
    cmyk.cyan = cmyk.magenta = cmyk.yellow = 0;
    cmyk.black = 100;
    return cmyk;
}

// 線幅プリセット名に現在の線幅単位を反映する / Apply the current stroke unit to preset labels
function formatFillStrokePresetLabel(preset) {
    var presetLabel = preset[lang] || preset.en;
    if (presetLabel.indexOf("{width}") < 0) return presetLabel;

    return presetLabel
        .replace("{width}", preset.strokeWidth)
        .replace("{unit}", getStrokeUnitLabel());
}

// ==============================
// 選択オブジェクトの判定 / Selection inspection
// ==============================

// リンク画像（PlacedItem）または配置画像（RasterItem）かどうか
function isImageItem(item) {
    return item.typename === "PlacedItem" || item.typename === "RasterItem";
}

// 選択オブジェクトがすべて画像かどうか / True when every selected item is an image
function isSelectionAllImages(pageItems) {
    if (!pageItems || pageItems.length === 0) return false;
    for (var i = 0; i < pageItems.length; i++) {
        if (!isImageItem(pageItems[i])) return false;
    }
    return true;
}

// ==============================
// パネル構築ヘルパー / Panel building helpers
// ==============================
var PANEL_MARGINS = [15, 20, 15, 10];
var PANEL_SPACING = 8;

// パネル共通設定。spacing 省略時は PANEL_SPACING
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ['fill', 'top'];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// 縦並びパネルを追加（チェックボックス・ラジオを縦に並べる用）
function addColumnPanel(parent, labelKey) {
    var panel = parent.add("panel", undefined, L(labelKey));
    setupPanel(panel);
    return panel;
}

// ==============================
// ダイアログ用ユーティリティ / Dialog utilities
// ==============================

// value が true のラジオの index を返す。見つからなければ fallback
function findSelectedRadioIndex(radios, fallback) {
    for (var i = 0; i < radios.length; i++) {
        if (radios[i].value) return i;
    }
    return fallback;
}

// 元オブジェクトの扱いラジオから "keep" | "mask" | "delete" を返す
function getOriginalMode(maskRadio, deleteRadio) {
    if (maskRadio.value) return "mask";
    if (deleteRadio.value) return "delete";
    return "keep";
}

// ダイアログ全体にキーボードショートカットを付与する
// handlers: { "S": function(){...}, "G": function(){...}, ... }
// テキスト入力にフォーカスがあるとき、および修飾キー併用時は発火しない
function addKeyboardShortcuts(dialog, handlers) {
    dialog.addEventListener("keydown", function (event) {
        // テキスト編集中はショートカットを横取りしない
        if (event.target && event.target.type === "edittext") return;
        // 修飾キー併用時はシステムショートカットの可能性があるため無視
        var keyboard = ScriptUI.environment.keyboardState;
        if (keyboard.shiftKey || keyboard.ctrlKey || keyboard.altKey || keyboard.metaKey) return;
        var handler = handlers[event.keyName];
        if (handler) {
            handler();
            event.preventDefault();
        }
    });
}

// テキスト入力フィールドに ↑/↓ キーでの数値増減を付与する
// ↑↓: ±1 / Shift+↑↓: ±10（10の倍数にスナップ）/ Alt+↑↓: ±0.1
// allowNegative=false で負値を 0 にクランプ。onAfterChange は値変更後のコールバック
function changeValueByArrowKey(editText, allowNegative, onAfterChange) {
    editText.addEventListener("keydown", function (event) {
        /* Up/Down 以外は触らない（タイピング中の値クランプ・onAfterChange 連発を防ぐ） */
        if (event.keyName != "Up" && event.keyName != "Down") return;

        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (keyboard.shiftKey) {
            var delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
            } else {
                value = Math.floor((value - 1) / delta) * delta;
            }
        } else if (keyboard.altKey) {
            value += (event.keyName == "Up") ? 0.1 : -0.1;
            value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
        } else {
            value += (event.keyName == "Up") ? 1 : -1;
            value = Math.round(value); /* 整数に丸め / Round to integer */
        }

        if (!allowNegative && value < 0) value = 0;

        editText.text = value;
        event.preventDefault();
        /* 値変更後のコールバック（プレビュー更新等）/ Optional callback after value change */
        if (typeof onAfterChange === "function") {
            onAfterChange();
        }
    });
}

// ==============================
// 設定ダイアログ / Settings dialog
// ==============================
function createSettingsDialogWindow() {
    var dialog = new Window("dialog", L("dialogTitle") + "  " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = 16;
    dialog.spacing = 12;
    return dialog;
}

function createDialogColumns(dialog) {
    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = "top";
    columnsGroup.spacing = 12;

    var leftColumn = columnsGroup.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = "fill";
    leftColumn.spacing = 12;

    var rightColumn = columnsGroup.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = "fill";
    rightColumn.spacing = 12;

    return {
        leftColumn: leftColumn,
        rightColumn: rightColumn
    };
}

function buildTargetPanel(parent) {
    var targetPanel = addColumnPanel(parent, "panelTarget");
    var eachRadio = targetPanel.add("radiobutton", undefined, L("targetEach"));
    eachRadio.helpTip = L("targetEachTip");
    var groupRadio = targetPanel.add("radiobutton", undefined, L("targetGroup"));
    groupRadio.helpTip = L("targetGroupTip");
    eachRadio.value = true;
    return {
        eachRadio: eachRadio,
        groupRadio: groupRadio
    };
}

function buildOptionsPanel(parent) {
    var optionsPanel = addColumnPanel(parent, "panelOptions");
    var previewBoundsCheck = optionsPanel.add("checkbox", undefined, L("previewBounds"));
    previewBoundsCheck.helpTip = L("boundsTip");
    var outlineTextCheck = optionsPanel.add("checkbox", undefined, L("outlineText"));
    outlineTextCheck.helpTip = L("outlineTextTip");

    var marginRow = optionsPanel.add("group");
    marginRow.orientation = "row";
    marginRow.alignChildren = ["left", "center"];
    var useMarginCheck = marginRow.add("checkbox", undefined, L("useMargin"));
    useMarginCheck.helpTip = L("marginTip");
    var marginInput = marginRow.add("edittext", undefined, "0");
    marginInput.characters = 5;
    marginInput.enabled = false;
    marginInput.helpTip = L("marginTip");
    var marginUnitLabel = marginRow.add("statictext", undefined, getRulerUnitLabel());
    marginUnitLabel.enabled = false;

    var dimCheck = optionsPanel.add("checkbox", undefined, L("dimSelection"));
    dimCheck.helpTip = L("dimSelectionTip");
    dimCheck.value = true;

    return {
        previewBoundsCheck: previewBoundsCheck,
        outlineTextCheck: outlineTextCheck,
        useMarginCheck: useMarginCheck,
        marginInput: marginInput,
        marginUnitLabel: marginUnitLabel,
        dimCheck: dimCheck
    };
}

function buildOriginalPanel(parent) {
    var originalPanel = addColumnPanel(parent, "panelOriginal");
    var keepRadio = originalPanel.add("radiobutton", undefined, L("originalKeep"));
    keepRadio.helpTip = L("originalKeepTip");
    var maskRadio = originalPanel.add("radiobutton", undefined, L("originalMask"));
    maskRadio.helpTip = L("maskTip");
    var deleteRadio = originalPanel.add("radiobutton", undefined, L("originalDelete"));
    deleteRadio.helpTip = L("originalDeleteTip");
    keepRadio.value = true;
    return {
        keepRadio: keepRadio,
        maskRadio: maskRadio,
        deleteRadio: deleteRadio
    };
}

function buildAppearancePanel(parent, selectionAllImages) {
    var appearancePanel = addColumnPanel(parent, "panelAppearance");
    var radios = [];
    for (var i = 0; i < FILL_STROKE_PRESETS.length; i++) {
        var presetLabel = formatFillStrokePresetLabel(FILL_STROKE_PRESETS[i]);
        radios.push(appearancePanel.add("radiobutton", undefined, presetLabel));
    }
    radios[DEFAULT_PRESET_INDEX].value = true;

    if (CURRENT_PRESET_INDEX >= 0) {
        radios[CURRENT_PRESET_INDEX].helpTip = L("currentTip");
    }
    if (selectionAllImages && CURRENT_PRESET_INDEX >= 0) {
        radios[CURRENT_PRESET_INDEX].enabled = false;
        radios[CURRENT_PRESET_INDEX].helpTip = L("imageNotice");
    }
    return radios;
}

function buildOrderPanel(parent) {
    var orderPanel = addColumnPanel(parent, "panelOrder");
    var frontRadio = orderPanel.add("radiobutton", undefined, L("orderFront"));
    frontRadio.helpTip = L("orderFrontTip");
    var backRadio = orderPanel.add("radiobutton", undefined, L("orderBack"));
    backRadio.helpTip = L("orderBackTip");
    frontRadio.value = true;
    return {
        frontRadio: frontRadio,
        backRadio: backRadio
    };
}

function buildButtonRow(dialog) {
    var buttonRow = dialog.add("group");
    buttonRow.orientation = "row";
    buttonRow.alignment = "fill";
    buttonRow.alignChildren = ["fill", "center"];

    var buttonLeft = buttonRow.add("group");
    buttonLeft.alignment = ["left", "center"];
    var previewCheck = buttonLeft.add("checkbox", undefined, L("preview"));
    previewCheck.helpTip = L("previewTip");

    var buttonSpacer = buttonRow.add("group");
    buttonSpacer.alignment = ["fill", "fill"];

    var buttonRight = buttonRow.add("group");
    buttonRight.alignment = ["right", "center"];
    var cancelButton = buttonRight.add("button", undefined, L("cancel"), { name: "cancel" });
    var okButton = buttonRight.add("button", undefined, "OK", { name: "ok" });

    return {
        previewCheck: previewCheck,
        cancelButton: cancelButton,
        okButton: okButton
    };
}

function getMarginValue(ui) {
    if (!ui.options.useMarginCheck.value) return 0;
    return rulerUnitValueToPoints(parseFloat(ui.options.marginInput.text) || 0);
}

function buildDialogSettings(ui, forPreview) {
    var appearanceIndex = findSelectedRadioIndex(ui.appearanceRadios, DEFAULT_PRESET_INDEX);
    return {
        groupAsOne: ui.target.groupRadio.value,
        useVisibleBounds: ui.options.previewBoundsCheck.value,
        outlineText: ui.options.outlineTextCheck.value,
        margin: getMarginValue(ui),
        appearancePreset: FILL_STROKE_PRESETS[appearanceIndex],
        placeInFront: ui.order.frontRadio.value,
        originalMode: forPreview ? "keep" : getOriginalMode(ui.original.maskRadio, ui.original.deleteRadio),
        preview: ui.buttons.previewCheck.value
    };
}
function syncMarginEnabled(ui) {
    ui.options.marginInput.enabled = ui.options.useMarginCheck.value;
    ui.options.marginUnitLabel.enabled = ui.options.useMarginCheck.value;
}

function clearPreviewItems(previewItems) {
    for (var i = 0; i < previewItems.length; i++) {
        try { previewItems[i].remove(); } catch (ignorePreview) { }
    }
    previewItems.splice(0, previewItems.length);
}

// 元の不透明度を退避（後で確実に復元するため）
function captureItemOpacities(items) {
    var snapshots = [];
    for (var i = 0; i < items.length; i++) {
        var current = 100;
        try { current = items[i].opacity; } catch (ignoreOpacityRead) { }
        snapshots.push({ item: items[i], opacity: current });
    }
    return snapshots;
}

function applyDimmedOpacity(items) {
    for (var i = 0; i < items.length; i++) {
        try { items[i].opacity = 50; } catch (ignoreDim) { }
    }
}

function restoreItemOpacities(snapshots) {
    for (var i = 0; i < snapshots.length; i++) {
        try { snapshots[i].item.opacity = snapshots[i].opacity; } catch (ignoreRestore) { }
    }
}

function renderDialogPreview(doc, eligibleItems, ui, previewItems, outlinedBoundsCache) {
    clearPreviewItems(previewItems);

    if (ui.buttons.previewCheck.value) {
        try {
            var previewSettings = buildDialogSettings(ui, true);
            var createdPreviewItems = produceRectangles(doc, eligibleItems, previewSettings, outlinedBoundsCache);
            for (var i = 0; i < createdPreviewItems.length; i++) {
                previewItems.push(createdPreviewItems[i]);
            }
        } catch (ignoreRender) { }
    }

    app.redraw();
}

function createDialogState(doc, eligibleItems, ui) {
    var previewItems = [];
    var outlinedBoundsCache = [];
    var opacitySnapshots = captureItemOpacities(eligibleItems);
    var dimmed = false;
    /* 不透明度付きプリセットで強制 OFF にする前の値を保存する（戻すときに使う）/
       Saves the dim value before being forced off by an opacity preset */
    var savedDimValue = null;

    function syncDimAvailability() {
        var appearanceIndex = findSelectedRadioIndex(ui.appearanceRadios, DEFAULT_PRESET_INDEX);
        var preset = FILL_STROKE_PRESETS[appearanceIndex];
        var presetHasOpacity = preset && typeof preset.opacity === "number";
        if (presetHasOpacity) {
            if (savedDimValue === null) {
                savedDimValue = ui.options.dimCheck.value;
            }
            ui.options.dimCheck.value = false;
            ui.options.dimCheck.enabled = false;
        } else {
            if (savedDimValue !== null) {
                ui.options.dimCheck.value = savedDimValue;
                savedDimValue = null;
            }
            ui.options.dimCheck.enabled = true;
        }
    }

    function syncDimming() {
        var shouldDim = ui.options.dimCheck.value;
        if (shouldDim && !dimmed) {
            applyDimmedOpacity(eligibleItems);
            dimmed = true;
        } else if (!shouldDim && dimmed) {
            restoreItemOpacities(opacitySnapshots);
            dimmed = false;
        }
        app.redraw();
    }

    function restoreOpacity() {
        if (dimmed) {
            restoreItemOpacities(opacitySnapshots);
            dimmed = false;
        }
    }

    return {
        buildSettings: function (forPreview) {
            return buildDialogSettings(ui, forPreview);
        },
        clearPreview: function () {
            clearPreviewItems(previewItems);
        },
        renderPreview: function () {
            renderDialogPreview(doc, eligibleItems, ui, previewItems, outlinedBoundsCache);
        },
        syncDimAvailability: syncDimAvailability,
        syncDimming: syncDimming,
        restoreOpacity: restoreOpacity
    };
}

function syncOrderEnabled(ui) {
    ui.order.frontRadio.enabled = ui.original.keepRadio.value;
    ui.order.backRadio.enabled = ui.original.keepRadio.value;
}

function bindPreviewEvents(ui, state) {
    ui.target.eachRadio.onClick = state.renderPreview;
    ui.target.groupRadio.onClick = state.renderPreview;
    ui.options.previewBoundsCheck.onClick = state.renderPreview;
    ui.options.outlineTextCheck.onClick = state.renderPreview;
    ui.options.useMarginCheck.onClick = function () {
        syncMarginEnabled(ui);
        state.renderPreview();
    };
    ui.options.marginInput.onChange = state.renderPreview;
    ui.options.dimCheck.onClick = state.syncDimming;
    for (var appearanceRadioIndex = 0; appearanceRadioIndex < ui.appearanceRadios.length; appearanceRadioIndex++) {
        ui.appearanceRadios[appearanceRadioIndex].onClick = function () {
            state.syncDimAvailability();
            state.syncDimming();
            state.renderPreview();
        };
    }
    ui.order.frontRadio.onClick = state.renderPreview;
    ui.order.backRadio.onClick = state.renderPreview;
    ui.buttons.previewCheck.onClick = state.renderPreview;
}

function bindOriginalHandlingEvents(ui) {
    ui.original.keepRadio.onClick = function () { syncOrderEnabled(ui); };
    ui.original.maskRadio.onClick = function () { syncOrderEnabled(ui); };
    ui.original.deleteRadio.onClick = function () { syncOrderEnabled(ui); };
}

function bindDialogKeyboardShortcuts(dialog, ui, state) {
    addKeyboardShortcuts(dialog, {
        "S": function () { ui.target.eachRadio.value = true; state.renderPreview(); },
        "G": function () { ui.target.groupRadio.value = true; state.renderPreview(); },
        "P": function () { ui.options.previewBoundsCheck.value = !ui.options.previewBoundsCheck.value; state.renderPreview(); },
        "O": function () { ui.options.outlineTextCheck.value = !ui.options.outlineTextCheck.value; state.renderPreview(); },
        "A": function () { ui.options.useMarginCheck.value = !ui.options.useMarginCheck.value; syncMarginEnabled(ui); state.renderPreview(); },
        "T": function () {
            if (!ui.options.dimCheck.enabled) return; // 不透明度プリセット選択中は無効化されているのでスキップ
            ui.options.dimCheck.value = !ui.options.dimCheck.value;
            state.syncDimming();
        },
        "F": function () { ui.order.frontRadio.value = true; state.renderPreview(); },
        "B": function () { ui.order.backRadio.value = true; state.renderPreview(); },
        "N": function () { ui.original.keepRadio.value = true; syncOrderEnabled(ui); },
        "M": function () { ui.original.maskRadio.value = true; syncOrderEnabled(ui); },
        "D": function () { ui.original.deleteRadio.value = true; syncOrderEnabled(ui); }
    });
}

function bindDialogEvents(dialog, ui, state) {
    bindPreviewEvents(ui, state);
    bindOriginalHandlingEvents(ui);
    bindDialogKeyboardShortcuts(dialog, ui, state);
}

function buildDialogUI(dialog, selectionAllImages) {
    var columns = createDialogColumns(dialog);
    return {
        target: buildTargetPanel(columns.leftColumn),
        options: buildOptionsPanel(columns.leftColumn),
        original: buildOriginalPanel(columns.leftColumn),
        appearanceRadios: buildAppearancePanel(columns.rightColumn, selectionAllImages),
        order: buildOrderPanel(columns.rightColumn),
        buttons: buildButtonRow(dialog)
    };
}

function showSettingsDialog(doc, eligibleItems, selectionAllImages) {
    var dialog = createSettingsDialogWindow();
    var ui = buildDialogUI(dialog, selectionAllImages);
    var state = createDialogState(doc, eligibleItems, ui);
    changeValueByArrowKey(ui.options.marginInput, true, state.renderPreview);
    syncMarginEnabled(ui);
    syncOrderEnabled(ui);
    bindDialogEvents(dialog, ui, state);
    state.syncDimAvailability();
    state.syncDimming();

    dialog.onClose = function () {
        state.clearPreview();
        state.restoreOpacity();
    };

    var dialogResult = null;
    ui.buttons.okButton.onClick = function () {
        dialogResult = state.buildSettings(false);
        dialog.close();
    };
    ui.buttons.cancelButton.onClick = function () {
        dialog.close();
    };

    dialog.show();
    return dialogResult;
}

// ==============================
// 塗り・線の適用 / Fill & stroke
// ==============================
// 生成した長方形そのものにプリセットの塗り・線・不透明度を適用する（元の選択オブジェクトには影響しない）
function applyRectAppearance(rect, preset, isRgbDocument) {
    if (!preset.applyAppearance) {
        return; // 直前の塗り・線：作成時の既定値のままにする
    }
    rect.filled = preset.filled;
    rect.stroked = preset.stroked;
    if (preset.filled) {
        rect.fillColor = makeBlack(isRgbDocument);
    }
    if (preset.stroked) {
        rect.strokeColor = makeBlack(isRgbDocument);
        rect.strokeWidth = strokeUnitValueToPoints(preset.strokeWidth);
    }
    if (typeof preset.opacity === "number") {
        rect.opacity = preset.opacity;
    }
}

// ==============================
// 長方形の作成 / Rectangle creation
// ==============================

// オブジェクトの外接矩形を取得する
// geometricBounds: パス本体の境界 / visibleBounds: 線幅・効果を含むプレビュー境界
function getItemBounds(item, useVisibleBounds) {
    return useVisibleBounds ? item.visibleBounds : item.geometricBounds;
}

// テキストをアウトライン化して計測した境界をキャッシュする
// 各エントリ: { item: TextFrame参照, geometric: [...], visible: [...] }
function findOutlinedBoundsCache(outlinedBoundsCache, item) {
    for (var i = 0; i < outlinedBoundsCache.length; i++) {
        if (outlinedBoundsCache[i].item === item) return outlinedBoundsCache[i];
    }
    return null;
}

// 計測用の外接矩形を取得する
// テキスト＋「テキストをアウトライン化」設定時は、複製をアウトライン化して計測し複製は破棄する
// 一度測定した結果は渡された outlinedBoundsCache に保存し、以降は再アウトライン化せず使い回す
function measureItemBounds(item, settings, outlinedBoundsCache) {
    if (settings.outlineText && item.typename === "TextFrame") {
        var cached = findOutlinedBoundsCache(outlinedBoundsCache, item);
        if (cached) {
            return settings.useVisibleBounds ? cached.visible : cached.geometric;
        }
        var duplicatedText = null;
        var outlinedGroup = null;
        try {
            duplicatedText = item.duplicate();
            outlinedGroup = duplicatedText.createOutline(); // 成功すると duplicatedText は消費され GroupItem になる
            // useVisibleBounds が後で切り替わっても再計測しなくて済むよう両方を保存する
            var geometricBounds = getItemBounds(outlinedGroup, false);
            var visibleBounds = getItemBounds(outlinedGroup, true);
            outlinedGroup.remove();
            outlinedBoundsCache.push({ item: item, geometric: geometricBounds, visible: visibleBounds });
            return settings.useVisibleBounds ? visibleBounds : geometricBounds;
        } catch (e) {
            // 計測用の複製・アウトラインがドキュメントに残らないよう後始末する
            if (outlinedGroup) {
                try { outlinedGroup.remove(); } catch (ignoreOutlined) { }
            } else if (duplicatedText) {
                try { duplicatedText.remove(); } catch (ignoreDuplicate) { }
            }
        }
    }
    return getItemBounds(item, settings.useVisibleBounds);
}

// 複数オブジェクトをまとめた外接矩形を返す / Combined bounding box of multiple objects
function getCombinedBounds(pageItems, settings, outlinedBoundsCache) {
    if (pageItems.length === 0) return null;
    var first = measureItemBounds(pageItems[0], settings, outlinedBoundsCache);
    var combined = [first[0], first[1], first[2], first[3]];
    for (var i = 1; i < pageItems.length; i++) {
        var b = measureItemBounds(pageItems[i], settings, outlinedBoundsCache);
        if (b[0] < combined[0]) combined[0] = b[0];
        if (b[1] > combined[1]) combined[1] = b[1];
        if (b[2] > combined[2]) combined[2] = b[2];
        if (b[3] < combined[3]) combined[3] = b[3];
    }
    return combined;
}

// 指定した外接矩形から長方形を作成し、プリセットの塗り・線を適用する
function createRectFromBounds(doc, bounds, settings, referenceItem) {
    // マージン分だけ境界を外側へ拡張（負値なら内側に縮む）
    var margin = (typeof settings.margin === "number") ? settings.margin : 0;
    var rectLeft = bounds[0] - margin;
    var rectTop = bounds[1] + margin;
    var rectRight = bounds[2] + margin;
    var rectBottom = bounds[3] - margin;

    var rectWidth = rectRight - rectLeft;
    var rectHeight = rectTop - rectBottom;
    if (rectWidth <= 0 || rectHeight <= 0) {
        return null;
    }

    var rect = doc.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
    // keep モードのみ「前面/背面」を反映。mask は後でグループ内に再配置、delete は元の位置に置く
    var placement;
    if (settings.originalMode === "keep") {
        placement = settings.placeInFront ? ElementPlacement.PLACEBEFORE : ElementPlacement.PLACEAFTER;
    } else {
        placement = ElementPlacement.PLACEBEFORE;
    }
    // move() の戻り値は環境により undefined になるため、元の参照をそのまま使う
    rect.move(referenceItem, placement);

    // 塗り・線は生成した長方形に直接適用する（選択中の元オブジェクトには触れない）
    var isRgbDocument = (doc.documentColorSpace === DocumentColorSpace.RGB);
    applyRectAppearance(rect, settings.appearancePreset, isRgbDocument);
    return rect;
}

// ==============================
// 元のオブジェクトの処理 / Handling the original object
// ==============================

// 元オブジェクト群のうち最前面（z 順で最も上）の項目を返す
function findTopMostItem(items) {
    var top = items[0];
    for (var i = 1; i < items.length; i++) {
        if (items[i].absoluteZOrderPosition > top.absoluteZOrderPosition) {
            top = items[i];
        }
    }
    return top;
}

// 長方形をクリッピングマスクとして適用し、マスクグループを返す
function applyClippingMask(doc, clipRect, contentItems) {
    // 元の最前面項目の位置にグループを差し込み、元の z 順序を維持する
    var topMost = findTopMostItem(contentItems);
    var clipGroup = doc.groupItems.add();
    clipGroup.move(topMost, ElementPlacement.PLACEBEFORE);
    // 元オブジェクトをグループへ移動
    for (var i = 0; i < contentItems.length; i++) {
        contentItems[i].move(clipGroup, ElementPlacement.MOVETOEND);
    }
    // クリップパス（長方形）は最前面に置く
    clipRect.move(clipGroup, ElementPlacement.MOVETOBEGINNING);
    clipRect.clipping = true;
    clipGroup.clipped = true;
    return clipGroup;
}

// 元のオブジェクトの扱い（そのまま／マスク／削除）を適用し、選択対象の項目を返す
function applyOriginalMode(doc, rect, originalItems, originalMode) {
    if (originalMode === "mask") {
        return applyClippingMask(doc, rect, originalItems);
    }
    if (originalMode === "delete") {
        for (var i = 0; i < originalItems.length; i++) {
            originalItems[i].remove();
        }
        return rect;
    }
    return rect; // そのまま / keep
}

// ==============================
// ロック・非表示判定 / Locked & hidden checks
// ==============================

// 親レイヤー・親グループを含めてロック状態かどうかを返す
function isItemLocked(item) {
    var current = item;
    while (current && current.typename !== "Document") {
        if (current.locked) return true;
        current = current.parent;
    }
    return false;
}

// 親レイヤー・親グループを含めて非表示状態かどうかを返す
function isItemHidden(item) {
    var current = item;
    while (current && current.typename !== "Document") {
        if (current.hidden) return true;
        current = current.parent;
    }
    return false;
}

// ==============================
// メイン処理のステップ / Main steps
// ==============================

// doc.selection を配列にコピー（後段で selection が書き換わっても参照を保持するため）
function collectSelectedItems(doc) {
    var items = [];
    for (var i = 0; i < doc.selection.length; i++) {
        items.push(doc.selection[i]);
    }
    return items;
}

// ロック・非表示オブジェクトを除いた処理対象を返す
function filterEligibleItems(items) {
    var eligible = [];
    for (var i = 0; i < items.length; i++) {
        if (!isItemLocked(items[i]) && !isItemHidden(items[i])) {
            eligible.push(items[i]);
        }
    }
    return eligible;
}

// 元オブジェクト群のうち最背面（z 順で最も下）の項目を返す
function findBottomMostItem(items) {
    var bottom = items[0];
    for (var i = 1; i < items.length; i++) {
        if (items[i].absoluteZOrderPosition < bottom.absoluteZOrderPosition) {
            bottom = items[i];
        }
    }
    return bottom;
}

// 選択範囲全体をまとめて1つの長方形（またはマスクグループ）にする
function produceGroupedRectangle(doc, items, settings, outlinedBoundsCache) {
    var combinedBounds = getCombinedBounds(items, settings, outlinedBoundsCache);
    if (!combinedBounds) return [];

    // 「前面に作成」は選択中の最前面項目、「背面に作成」は最背面項目を基準にする
    var referenceItem = settings.placeInFront ? findTopMostItem(items) : findBottomMostItem(items);
    var groupRect = createRectFromBounds(doc, combinedBounds, settings, referenceItem);
    if (!groupRect) return [];

    return [applyOriginalMode(doc, groupRect, items, settings.originalMode)];
}

// オブジェクトごとに長方形（またはマスクグループ）を作成する
function produceIndividualRectangles(doc, items, settings, outlinedBoundsCache) {
    var created = [];
    for (var i = 0; i < items.length; i++) {
        var itemBounds = measureItemBounds(items[i], settings, outlinedBoundsCache);
        var rect = createRectFromBounds(doc, itemBounds, settings, items[i]);
        if (rect) {
            created.push(applyOriginalMode(doc, rect, [items[i]], settings.originalMode));
        }
    }
    return created;
}

// 設定に従って長方形（またはマスクグループ）を生成し、結果配列を返す
function produceRectangles(doc, items, settings, outlinedBoundsCache) {
    if (!outlinedBoundsCache) outlinedBoundsCache = [];

    if (settings.groupAsOne) {
        return produceGroupedRectangle(doc, items, settings, outlinedBoundsCache);
    }
    return produceIndividualRectangles(doc, items, settings, outlinedBoundsCache);
}

// ==============================
// メイン処理 / Main
// ==============================
function main() {
    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert(L("noSelection"));
        return;
    }

    var selectedItems = collectSelectedItems(doc);
    var eligibleItems = filterEligibleItems(selectedItems);
    if (eligibleItems.length === 0) {
        alert(L("noResult"));
        return;
    }

    app.executeMenuCommand('AI Bounding Box Toggle');
    app.executeMenuCommand('edge');
    try {
        var settings = showSettingsDialog(doc, eligibleItems, isSelectionAllImages(eligibleItems));
        if (!settings) {
            return; // キャンセル / Cancelled
        }

        var createdItems = produceRectangles(doc, eligibleItems, settings);

        if (createdItems.length === 0) {
            alert(L("noResult"));
            return;
        }

        // 作成した長方形（マスク時はマスクグループ）のみを選択
        doc.selection = createdItems;
    } finally {
        app.executeMenuCommand('edge');
        app.executeMenuCommand('AI Bounding Box Toggle');
    }
}

main();
