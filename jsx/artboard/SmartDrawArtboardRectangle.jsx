#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

アートボードサイズの長方形を描画

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- アクティブまたは全アートボードと**同サイズ**の長方形を描画します。
- カラー（なし / K100・不透明度15%）や重ね順（最前面 / 最背面 / bgレイヤー）を選択可能。プレビューは常時ONで描画専用レイヤーに表示します。

### 主な機能：

- オフセット指定（±で内外に均等拡張）
- 対象範囲の切替（現在のアートボード / すべて）
- 重ね順の指定（最前面 / 最背面 / bgレイヤーに配置）
- プレビュー（破線1pt・50%トーン、Previewレイヤーに描画）
- ダイアログ位置・透明度の設定、位置の記憶（#targetengine 利用）

### 処理の流れ：

1) ダイアログ表示 → 入力（オフセット/カラー/重ね順/対象）
2) 入力に応じてプレビューを更新（デバウンスあり）
3) OK で本描画（対象の各アートボードに長方形を生成）

### オリジナル、謝辞：

（該当なし）

### note：

- 「ディム表示（指定）」は UI のみ実装。描画ロジックは未実装です。

### 動画：

（該当なし）

### 更新履歴：

- v1.0 (20250820) : 初期バージョン

---

### Script Name:

Draw Artboard-Sized Rectangle

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Draws rectangles **matching the size** of the active or all artboards.
- Select color (None / K100 with 15% opacity) and stacking order (Front / Back / bg layer). Live preview is always on and drawn on a dedicated layer.

### Key Features:

- Offset (expand/shrink evenly outward/inward with ± values)
- Target scope (Current artboard / All artboards)
- Z-order (Bring to Front / Send to Back / Place in bg layer)
- Preview (1pt dashed, 50% tone on a Preview layer)
- Dialog opacity/position and position memory (#targetengine)

### Flow:

1) Show dialog → Enter options (offset/color/z-order/target)
2) Live preview updates with debounce
3) On OK, draw rectangles for the selected target

### Credits / Acknowledgements:

(none)

### Notes:

- "Dim" mode is UI-only at the moment; rendering logic is not implemented yet.

### Videos:

(none)

### Changelog:

- v1.0 (20250820): Initial version

*/


var SCRIPT_VERSION = "v1.0";

// ===== Dialog appearance & position (tunable) =====
var DIALOG_OFFSET_X = 300; // shift right (+) / left (-)
var DIALOG_OFFSET_Y = 0; // shift down (+) / up (-)
var DIALOG_OPACITY = 0.98; // 0.0 - 1.0

function setDialogOpacity(dlg, opacityValue) {
    try {
        dlg.opacity = opacityValue;
    } catch (e) {}
}

function shiftDialogPositionOnce(dlg, offsetX, offsetY) {
    try {
        var loc = dlg.location;
        dlg.location = [loc[0] + offsetX, loc[1] + offsetY];
    } catch (e) {}
}

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ラベル定義 / Label definitions (UI order) */
var LABELS = {
    dialogTitle: {
        ja: "アートボードサイズの長方形を描画" + ' ' + SCRIPT_VERSION,
        en: "Draw Artboard-Sized Rectangle"+ ' ' + SCRIPT_VERSION
    },
    // Panels
    offsetTitle: {
        ja: "オフセット",
        en: "Offset"
    },

    colorTitle: {
        ja: "カラー",
        en: "Color"
    },
    zorderTitle: {
        ja: "重ね順",
        en: "Stacking Order"
    },
    targetTitle: {
        ja: "対象",
        en: "Target"
    },
    // Color options
    colorNone: {
        ja: "なし",
        en: "None"
    },
    colorK100: {
        ja: "K100、不透明度15%",
        en: "K100, Opacity 15%"
    },
    colorSpecified: {
        ja: "ディム表示",
        en: "Dim"
    },
    // Z-order options
    front: {
        ja: "最前面",
        en: "Bring to Front"
    },
    back: {
        ja: "最背面",
        en: "Send to Back"
    },
    bg: {
        ja: "「bg」レイヤー",
        en: "\"bg\" Layer"
    },
    // Target options
    currentAB: {
        ja: "現在のアートボードのみ",
        en: "Current Artboard Only"
    },
    allAB: {
        ja: "すべて",
        en: "All"
    },
    // Buttons
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    // Names
    previewLayer: {
        ja: "プレビュー",
        en: "Preview"
    },
    rectName: {
        ja: "アートボード境界",
        en: "Artboard Bounds"
    },
    previewRect: {
        ja: "__プレビュー_アートボード境界",
        en: "__Preview_ArtboardBounds"
    }
};

function createBlackColor(doc) {
    if (doc.documentColorSpace == DocumentColorSpace.RGB) {
        var blackColor = new RGBColor();
        blackColor.red = 0;
        blackColor.green = 0;
        blackColor.blue = 0;
        return blackColor;
    } else {
        var blackColor = new CMYKColor();
        blackColor.black = 100;
        blackColor.cyan = 0;
        blackColor.magenta = 0;
        blackColor.yellow = 0;
        return blackColor;
    }
}

// 単位コードとラベルのマップ / Map rulerType codes to labels
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

// 現在の単位ラベルを取得 / Get current unit label from prefs
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

function changeValueByArrowKey(editText, onValueChange) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            // 小数第1位までに丸め
            value = Math.round(value * 10) / 10;
        } else {
            // 整数に丸め
            value = Math.round(value);
        }

        editText.text = value;
        try {
            if (typeof onValueChange === 'function') onValueChange();
        } catch (e) {}
    });
}

// ===== Preview helpers =====
var __previewItems = [];

var __previewDebounceTask = null;

function schedulePreview(choice, delayMs) {
    try {
        if (__previewDebounceTask) app.cancelTask(__previewDebounceTask);
    } catch (e) {}
    $.global.__lastPreviewChoice = choice;
    var code = 'try{renderPreview(app.activeDocument, $.global.__lastPreviewChoice);}catch(e){}';
    try {
        __previewDebounceTask = app.scheduleTask(code, Math.max(0, delayMs | 0), false);
    } catch (e) {
        try {
            renderPreview(app.activeDocument, choice);
        } catch (_) {}
    }
}

/*
 * プレビュー要求（即時/遅延）/ Request preview (immediate or debounced)
 */
function requestPreview(choice, immediate) {
    if (immediate) {
        try { if (__previewDebounceTask) app.cancelTask(__previewDebounceTask); } catch(_) {}
        try { renderPreview(app.activeDocument, choice); } catch(_) {}
    } else {
        schedulePreview(choice, 75);
    }
}

function clearPreview() {
    try {
        var doc = app.activeDocument;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === LABELS.previewLayer[lang]) {
                try {
                    doc.layers[i].remove();
                } catch (e) {}
                break;
            }
        }
    } catch (e) {}
    __previewItems = [];
}

function getOrCreatePreviewLayer(doc) {
    var name = LABELS.previewLayer[lang];
    var layer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) {
            layer = doc.layers[i];
            break;
        }
    }
    if (!layer) {
        layer = doc.layers.add();
        layer.name = name;
    }
    layer.visible = true;
    layer.locked = false;
    try {
        layer.move(doc, ElementPlacement.PLACEATBEGINNING);
    } catch (e) {}
    return layer;
}

// --- Helper: returns 50% gray/black stroke color matching document color space
function getPreviewStrokeColor(doc) {
    if (doc.documentColorSpace == DocumentColorSpace.RGB) {
        var c = new RGBColor();
        c.red = 128;
        c.green = 128;
        c.blue = 128; // ~50% gray
        return c;
    } else {
        var c = new CMYKColor();
        c.cyan = 0;
        c.magenta = 0;
        c.yellow = 0;
        c.black = 50; // K=50%
        return c;
    }
}

/*
 * プレビュー描画（Previewレイヤーへ一時オブジェクトを生成）
 * Render live preview into the dedicated Preview layer.
 */
function renderPreview(doc, choice) {
    clearPreview();
    if (!doc || !choice) return;

    var prevCS = null;
    try {
        prevCS = app.coordinateSystem;
    } catch (e) {}
    try {
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
    } catch (e) {}

    var prevIndex = doc.artboards.getActiveArtboardIndex();

    var previewLayer = getOrCreatePreviewLayer(doc);

    function previewOne(idx) {
        try {
            // Removed setting active artboard to avoid missed drawing
            // doc.artboards.setActiveArtboardIndex(idx);
        } catch (e) {}
        var ab = doc.artboards[idx];
        var abRect = ab.artboardRect;
        var abWidth = abRect[2] - abRect[0];
        var abHeight = abRect[1] - abRect[3];
        var o = choice.offset || 0;
        var rect = previewLayer.pathItems.rectangle(
            abRect[1] + o, abRect[0] - o, abWidth + o * 2, abHeight + o * 2
        );

        if (choice.colorMode === 'none') {
            rect.filled = false;
            rect.stroked = false;
        } else if (choice.colorMode === 'k100') {
            rect.filled = true;
            rect.fillColor = createBlackColor(doc);
            rect.stroked = false;
            rect.opacity = 15;
        } else if (choice.colorMode === 'specified') {
            // 仕様未定：とりあえず none と同等
            rect.filled = false;
            rect.stroked = false;
        }

        // --- Preview stroke (for visibility): dashed 1pt, 50% tone ---
        rect.stroked = true; // ensure stroke is on even if colorMode is 'none'
        rect.strokeWidth = 1;
        try {
            rect.strokeDashes = [6, 4];
        } catch (e) {} // dashed if supported
        rect.strokeColor = getPreviewStrokeColor(doc);

        rect.name = LABELS.previewRect[lang];
        rect.selected = false;
        if (choice.zOrder === 'front') {
            rect.zOrder(ZOrderMethod.BRINGTOFRONT);
        } else if (choice.zOrder === 'back') {
            rect.zOrder(ZOrderMethod.SENDTOBACK);
        }
    }

    if (choice.target === 'all') {
        for (var i = 0; i < doc.artboards.length; i++) previewOne(i);
    } else {
        previewOne(doc.artboards.getActiveArtboardIndex());
    }

    try {
        if (prevCS !== null) app.coordinateSystem = prevCS;
    } catch (e) {}
    app.redraw();
}

function showDialog() {
    var dlg = new Window('dialog', LABELS.dialogTitle[lang]);
    setDialogOpacity(dlg, DIALOG_OPACITY);
    dlg.alignChildren = 'left';

    // --- Add two-column group container ---
    var cols = dlg.add('group');
    cols.orientation = 'row';
    cols.alignChildren = ['fill', 'top'];
    cols.spacing = 12;

    var leftCol = cols.add('group');
    leftCol.orientation = 'column';
    leftCol.alignChildren = 'fill';
    leftCol.spacing = 10;

    var rightCol = cols.add('group');
    rightCol.orientation = 'column';
    rightCol.alignChildren = 'fill';
    rightCol.spacing = 10;

    // Offset panel (with margins and row)
    var offsetPanel = leftCol.add('panel', undefined, LABELS.offsetTitle[lang]);
    offsetPanel.orientation = 'column';
    offsetPanel.alignChildren = 'left';
    offsetPanel.margins = [15, 20, 15, 10];

    var offsetRow = offsetPanel.add('group');
    offsetRow.orientation = 'row';
    offsetRow.alignChildren = 'center';
    offsetRow.alignment = 'center';


    var offsetInput = offsetRow.add('edittext', undefined, '0');
    offsetInput.characters = 5;
    changeValueByArrowKey(offsetInput, function() {
        updatePreview();
    });

    // Enterキーでも明示的に更新
    offsetInput.addEventListener('keydown', function(e) {
        if (e.keyName == 'Enter') {
            try { requestPreview(buildChoiceFromUI(), true); } catch (_) {}
        }
    });

    var unitLabel = getCurrentUnitLabel();
    offsetRow.add('statictext', undefined, unitLabel);

    dlg.onShow = function() {
        try {
            // Focus offset field first
            offsetInput.active = true;
        } catch (e) {}
        // Shift dialog position once on show
        shiftDialogPositionOnce(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        // Render initial preview
        updatePreview();
    };

    // Add new panel for color
    var colorPanel = leftCol.add('panel', undefined, LABELS.colorTitle[lang]);
    colorPanel.orientation = 'column';
    colorPanel.alignChildren = 'left';
    colorPanel.margins = [15, 20, 15, 10];

    var noneRadio = colorPanel.add('radiobutton', undefined, LABELS.colorNone[lang]);
    var k100Radio = colorPanel.add('radiobutton', undefined, LABELS.colorK100[lang]);
    var specifiedRadio = colorPanel.add('radiobutton', undefined, LABELS.colorSpecified[lang]);

    k100Radio.value = true;
    // ディム表示は未実装のため使用不可 / Disable 'Dim' option (not implemented yet)
    specifiedRadio.enabled = false;
    specifiedRadio.value = false;

    // Add new panel for zOrder
    var zOrderPanel = rightCol.add('panel', undefined, LABELS.zorderTitle[lang]);
    zOrderPanel.orientation = 'column';
    zOrderPanel.alignChildren = 'left';
    zOrderPanel.margins = [15, 20, 15, 10];

    var frontRadio = zOrderPanel.add('radiobutton', undefined, LABELS.front[lang]);
    var backRadio = zOrderPanel.add('radiobutton', undefined, LABELS.back[lang]);
    var bgLayerRadio = zOrderPanel.add('radiobutton', undefined, LABELS.bg[lang]);

    backRadio.value = true;

    // Add new panel for target
    var targetPanel = rightCol.add('panel', undefined, LABELS.targetTitle[lang]);
    targetPanel.orientation = 'column';
    targetPanel.alignChildren = 'left';
    targetPanel.margins = [15, 20, 15, 10];

    var currentRadio = targetPanel.add('radiobutton', undefined, LABELS.currentAB[lang]);
    var allRadio = targetPanel.add('radiobutton', undefined, LABELS.allAB[lang]);

    // 自動選択: アートボード数で切り替え
    var abCount = (app.documents.length ? app.activeDocument.artboards.length : 0);
    if (abCount <= 1) {
        currentRadio.value = true; // 1つ以下のときは「現在のみ」
        allRadio.value = false;
    } else {
        currentRadio.value = false;
        allRadio.value = true; // 複数あるときは「すべて」
    }

    function buildChoiceFromUI() {
        var colorMode = noneRadio.value ? 'none' : (k100Radio.value ? 'k100' : (specifiedRadio.value ? 'specified' : 'none'));
        var zOrder = frontRadio.value ? 'front' : (backRadio.value ? 'back' : (bgLayerRadio.value ? 'bg' : 'back'));
        var target = currentRadio.value ? 'current' : (allRadio.value ? 'all' : 'current');
        var offset = parseFloat(offsetInput.text);
        if (isNaN(offset)) offset = 0;
        return {
            colorMode: colorMode,
            offset: offset,
            zOrder: zOrder,
            target: target
        };
    }

    function updatePreview() {
        try { requestPreview(buildChoiceFromUI(), false); } catch (e) {}
    }

    offsetInput.onChanging = updatePreview;
    offsetInput.onChange = updatePreview;

    noneRadio.onClick = updatePreview;
    k100Radio.onClick = updatePreview;
    specifiedRadio.onClick = updatePreview;

    frontRadio.onClick = updatePreview;
    backRadio.onClick = updatePreview;
    bgLayerRadio.onClick = updatePreview;

    currentRadio.onClick = updatePreview;
    allRadio.onClick = updatePreview;
    currentRadio.onChanging = updatePreview;
    allRadio.onChanging = updatePreview;

    var btnGroup = dlg.add('group');
    btnGroup.alignment = 'center';
    var cancelBtn = btnGroup.add('button', undefined, LABELS.cancel[lang]);
    var okBtn = btnGroup.add('button', undefined, LABELS.ok[lang]);

    okBtn.onClick = function() {
        try {
            if (__previewDebounceTask) app.cancelTask(__previewDebounceTask);
        } catch (e) {}
        clearPreview();
        dlg.close(1);
    };
    cancelBtn.onClick = function() {
        try {
            if (__previewDebounceTask) app.cancelTask(__previewDebounceTask);
        } catch (e) {}
        clearPreview();
        dlg.close(0);
    };

    var result = dlg.show();
    if (result != 1) {
        return null;
    }

    var colorMode = null;
    if (noneRadio.value) {
        colorMode = 'none';
    } else if (k100Radio.value) {
        colorMode = 'k100';
    } else if (specifiedRadio.value) {
        colorMode = 'specified';
    }

    var offset = parseFloat(offsetInput.text);
    if (isNaN(offset)) offset = 0;

    var zOrder = null;
    if (frontRadio.value) {
        zOrder = 'front';
    } else if (backRadio.value) {
        zOrder = 'back';
    } else if (bgLayerRadio.value) {
        zOrder = 'bg';
    }

    var target = null;
    if (currentRadio.value) {
        target = 'current';
    } else if (allRadio.value) {
        target = 'all';
    }

    return {
        colorMode: colorMode,
        offset: offset,
        zOrder: zOrder,
        target: target
    };
}

/*
 * 「bg」レイヤーの取得/作成 / Get or create the "bg" layer
 */
function getOrCreateBgLayer(doc) {
    var name = 'bg';
    var layer = null;
    // 検索
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) {
            layer = doc.layers[i];
            break;
        }
    }
    // 作成
    if (!layer) {
        layer = doc.layers.add();
        layer.name = name;
    }
    // 見える＆編集可能に
    layer.visible = true;
    layer.locked = false;
    // 最背面へ
    try {
        layer.move(doc, ElementPlacement.PLACEATEND);
    } catch (e) {}
    return layer;
}

function drawRectangleForArtboard(doc, ab, choice) {
    var abRect = ab.artboardRect; // [left, top, right, bottom]

    var abWidth = abRect[2] - abRect[0];
    var abHeight = abRect[1] - abRect[3];

    var o = choice.offset;
    var targetContainer = doc;
    if (choice.zOrder === 'bg') {
        targetContainer = getOrCreateBgLayer(doc);
    }
    var rect = (choice.zOrder === 'bg' ? targetContainer.pathItems : doc.pathItems).rectangle(abRect[1] + o, abRect[0] - o, abWidth + o * 2, abHeight + o * 2);

    if (choice.colorMode === 'none') {
        rect.filled = false;
        rect.stroked = false;
    } else if (choice.colorMode === 'k100') {
        rect.filled = true;
        rect.fillColor = createBlackColor(doc);
        rect.stroked = false;
        rect.opacity = 15;
    }

    rect.name = LABELS.rectName[lang];
    rect.selected = true;

    if (choice.zOrder === 'front') {
        rect.zOrder(ZOrderMethod.BRINGTOFRONT);
    } else if (choice.zOrder === 'back') {
        rect.zOrder(ZOrderMethod.SENDTOBACK);
    }
}

function main() {
    if (app.documents.length === 0) return;

    var choice = showDialog();
    if (choice === null) return;

    var doc = app.activeDocument;

    var __prevCS = null;
    try {
        __prevCS = app.coordinateSystem;
    } catch (e) {}
    try {
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
    } catch (e) {}

    app.executeMenuCommand('deselectall'); // 既存選択を解除

    if (choice.target === 'current') {
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        drawRectangleForArtboard(doc, ab, choice);
    } else if (choice.target === 'all') {
        var prevIndex = doc.artboards.getActiveArtboardIndex();
        for (var i = 0; i < doc.artboards.length; i++) {
            try {
                doc.artboards.setActiveArtboardIndex(i);
            } catch (e) {}
            var ab = doc.artboards[i];
            drawRectangleForArtboard(doc, ab, choice);
        }
        try {
            doc.artboards.setActiveArtboardIndex(prevIndex);
        } catch (e) {}
    }

    try {
        if (__prevCS !== null) app.coordinateSystem = __prevCS;
    } catch (e) {}
}

main();