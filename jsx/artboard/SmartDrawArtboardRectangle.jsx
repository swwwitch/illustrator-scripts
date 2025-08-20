#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

アートボードサイズを調整

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- アクティブまたは全アートボードと**同サイズ**の長方形を、オフセットを考慮して描画します。
- カラー（なし／K100・15%）・HEX・CMYK、重ね順（最前面／最背面／bgレイヤー）を指定でき、ライブプレビューで確認できます。

### 主な機能：

- オフセット指定（裁ち落としプリセット：3mm／12H／0.125in）
- カラー指定（None／K100 15%／HEX／CMYK）
- 重ね順（Front／Back／bgレイヤー）
- 対象範囲（作業アートボード／すべて）
- プレビュー（1ptの破線、50%トーン、専用レイヤー）
- ダイアログ位置・不透明度の設定（位置記憶に #targetengine 利用）

### 処理の流れ：

1) ダイアログでオプションを入力（オフセット／カラー／重ね順／対象）
2) 入力と同時にプレビューを更新（デバウンスあり）
3) OKで本描画（対象アートボードに長方形を生成）

### note：

- 「ディム表示（指定）」は UI のみ実装。描画ロジックは未実装です。

### 更新履歴：

- v1.2 (20250821) : CMYKを独立UI化したため、parseCustomColor() から旧CMYK文字列解釈（"C0 M0 Y0 K100" 等）を削除。
- v1.1 (20250821) : ダイアログの位置・透明度を設定可能に。プレビューの破線を1ptに変更。オフセットの単位を現在の単位に合わせるよう修正。
- v1.0 (20250820) : 初期バージョン

---

### Script Name:

Adjust Artboard Size

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Draws rectangles that **match the active or all artboards**, with optional offset.
- Supports color (None/K100 15%/HEX/CMYK), stacking order (Front/Back/bg layer), and live preview.

### Key Features:

- Offset with Bleed presets (3mm / 12H / 0.125in)
- Color modes (None / K100 15% / HEX / CMYK)
- Z-order (Front / Back / bg layer)
- Target scope (Current artboard / All artboards)
- Preview (1pt dashed stroke, 50% tone, dedicated layer)
- Dialog position & opacity settings (position memory via #targetengine)

### Flow:

1) Enter options (offset/color/z-order/target) in the dialog
2) Live preview updates with debounce
3) On OK, draw rectangles for the selected target

### Notes:

- "Dim" mode is UI-only for now; rendering logic is not implemented yet.

### Changelog:

- v1.2 (20250821): Removed legacy CMYK string parsing (e.g., "C0 M0 Y0 K100") from parseCustomColor() since CMYK has its own UI.
- v1.1 (20250821): Added dialog position/opacity settings. Changed preview stroke to 1pt dashed. Offset unit now matches current ruler setting.
- v1.0 (20250820): Initial version

*/

var SCRIPT_VERSION = "v1.2";

/*
 * Color mode constants / カラーモード定数
 */
var ColorMode = {
    NONE: 'none',
    K100: 'k100',
    HEX: 'hex',
    CMYK: 'cmyk'
};

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ラベル定義 / Label definitions (UI order) */
var LABELS = {
    dialogTitle: {
        ja: "アートボードサイズを調整 " + SCRIPT_VERSION,
        en: "Adjust Artboard Size " + SCRIPT_VERSION
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
    // Offset options
    bleed: {
        ja: "裁ち落とし",
        en: "Bleed"
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
        ja: "HEX",
        en: "HEX"
    },
    colorCustomCMYK: {
        ja: "CMYK",
        en: "CMYK"
    },
    colorCustomHint: {
        ja: "例: #FF0000 / 255,0,0",
        en: "e.g., #FF0000 / 255,0,0"
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
        ja: "作業アートボード",
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
        ja: "_preview",
        en: "_preview"
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

// --- Helpers: color constructors & parsers ---
function _clamp(v, min, max) {
    return v < min ? min : (v > max ? max : v);
}

function makeRGB(r, g, b) {
    var c = new RGBColor();
    c.red = _clamp(Math.round(r), 0, 255);
    c.green = _clamp(Math.round(g), 0, 255);
    c.blue = _clamp(Math.round(b), 0, 255);
    return c;
}


function makeCMYK(cy, mg, yl, k) {
    var c = new CMYKColor();
    c.cyan = _clamp(cy, 0, 100);
    c.magenta = _clamp(mg, 0, 100);
    c.yellow = _clamp(yl, 0, 100);
    c.black = _clamp(k, 0, 100);
    return c;
}

// --- RGB/CMYK conversion helpers ---
function rgbToCmyk(r, g, b) {
    // r,g,b: 0-255 → return [C,M,Y,K] 0-100
    r = _clamp(r, 0, 255) / 255;
    g = _clamp(g, 0, 255) / 255;
    b = _clamp(b, 0, 255) / 255;
    var k = 1 - Math.max(r, g, b);
    if (k >= 0.9999) return [0, 0, 0, 100];
    var c = (1 - r - k) / (1 - k);
    var m = (1 - g - k) / (1 - k);
    var y = (1 - b - k) / (1 - k);
    return [Math.round(c * 100), Math.round(m * 100), Math.round(y * 100), Math.round(k * 100)];
}

function cmykToRgb(c, m, y, k) {
    // c,m,y,k: 0-100 → return [R,G,B] 0-255
    c = _clamp(c, 0, 100) / 100;
    m = _clamp(m, 0, 100) / 100;
    y = _clamp(y, 0, 100) / 100;
    k = _clamp(k, 0, 100) / 100;
    var r = 255 * (1 - c) * (1 - k);
    var g = 255 * (1 - m) * (1 - k);
    var b = 255 * (1 - y) * (1 - k);
    return [Math.round(r), Math.round(g), Math.round(b)];
}

function _toInt(x) {
    var n = parseFloat(x);
    return isNaN(n) ? NaN : n;
}

/*
 * customValue の解釈と RGBColor/CMYKColor へのマッピング
 * Accepts:
 *  - "#RRGGBB"
 *  - "R,G,B"  (0-255)
 *  - 名前色: black, white, red, green, blue, cyan, magenta, yellow, orange, grayXX (0-100)
 */
function parseCustomColor(doc, customValue) {
    try {
        if (!customValue) return null;
        var s = String(customValue).replace(/^\s+|\s+$/g, '').toLowerCase();
        if (!s) return null;

        // Normalize separators: convert full-width spaces and JP commas to ASCII
        s = s.replace(/\u3000/g, ' '); // full-width space → space
        s = s.replace(/[，、]/g, ','); // Japanese comma/ton-ten → comma
        // Normalize full-width digits and punctuation to ASCII
        s = s.replace(/[０-９]/g, function(ch) {
            return String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30);
        });
        s = s.replace(/．/g, '.'); // full-width period
        s = s.replace(/／/g, '/'); // full-width slash

        // #RRGGBB
        if (s.charAt(0) === '#' && s.length === 7) {
            var r = parseInt(s.substr(1, 2), 16);
            var g = parseInt(s.substr(3, 2), 16);
            var b = parseInt(s.substr(5, 2), 16);
            if (!isNaN(r) && !isNaN(g) && !isNaN(b)) return makeRGB(r, g, b);
        }

        // R,G,B  (commas or spaces)
        var threeNums = s.split(/[\s\u3000,，\/、]+/);
        if (threeNums.length === 3 && threeNums.every(function(x) {
                return /^\d+(?:\.\d+)?$/.test(x);
            })) {
            var r2 = _toInt(threeNums[0]),
                g2 = _toInt(threeNums[1]),
                b2 = _toInt(threeNums[2]);
            if (!isNaN(r2) && !isNaN(g2) && !isNaN(b2)) {
                if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK) {
                    var cmyk = rgbToCmyk(r2, g2, b2);
                    return makeCMYK(cmyk[0], cmyk[1], cmyk[2], cmyk[3]);
                }
                return makeRGB(r2, g2, b2);
            }
        }

        // Named colors (basic)
        var named = {
            black: {
                rgb: [0, 0, 0],
                cmyk: [0, 0, 0, 100]
            },
            white: {
                rgb: [255, 255, 255],
                cmyk: [0, 0, 0, 0]
            },
            red: {
                rgb: [255, 0, 0],
                cmyk: [0, 100, 100, 0]
            },
            green: {
                rgb: [0, 128, 0],
                cmyk: [100, 0, 100, 50]
            },
            blue: {
                rgb: [0, 0, 255],
                cmyk: [100, 100, 0, 0]
            },
            cyan: {
                rgb: [0, 255, 255],
                cmyk: [100, 0, 0, 0]
            },
            magenta: {
                rgb: [255, 0, 255],
                cmyk: [0, 100, 0, 0]
            },
            yellow: {
                rgb: [255, 255, 0],
                cmyk: [0, 0, 100, 0]
            },
            orange: {
                rgb: [255, 165, 0],
                cmyk: [0, 35, 100, 0]
            }
        };
        if (named[s]) {
            // Prefer document color space, else RGB
            if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK)
                return makeCMYK(named[s].cmyk[0], named[s].cmyk[1], named[s].cmyk[2], named[s].cmyk[3]);
            return makeRGB(named[s].rgb[0], named[s].rgb[1], named[s].rgb[2]);
        }
        // grayNN (0-100)
        var m = s.match(/^gray\s*(\d{1,3})$/);
        if (m) {
            var gk = _clamp(parseInt(m[1], 10), 0, 100);
            if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK) return makeCMYK(0, 0, 0, gk);
            var v = Math.round(255 * (100 - gk) / 100);
            return makeRGB(v, v, v);
        }
    } catch (e) {}
    return null;
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

// 現在の単位コードを取得 / Get current rulerType code
function getCurrentUnitCode() {
    try {
        return app.preferences.getIntegerPreference("rulerType");
    } catch (e) {
        return 2; // fallback to pt
    }
}

// 単位コード→pt係数 / Convert unit code to points factor
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0; // in
        case 1:
            return 72.0 / 25.4; // mm
        case 2:
            return 1.0; // pt
        case 3:
            return 12.0; // pica
        case 4:
            return 72.0 / 2.54; // cm
        case 5:
            return 72.0 / 25.4 * 0.25; // Q or H
        case 6:
            return 1.0; // px (Illustrator=1px=1pt)
        case 7:
            return 72.0 * 12.0; // ft/in
        case 8:
            return 72.0 / 25.4 * 1000.0; // m
        case 9:
            return 72.0 * 36.0; // yd
        case 10:
            return 72.0 * 12.0; // ft
        default:
            return 1.0;
    }
}

/*
 * Resolve offset display text & internal pt value in one place
 * 単位変換と裁ち落としプリセットを一元化
 * - offsetText: current edit field text (string)
 * - unitCode: app.preferences.getIntegerPreference("rulerType")
 * - bleedEnabled: true when Bleed is ON
 * Return: { pt: Number, displayText: String, disabled: Boolean }
 */
function resolveOffsetToPt(offsetText, unitCode, bleedEnabled) {
    var displayText = String(offsetText == null ? '' : offsetText);
    var pt = 0;

    if (bleedEnabled) {
        // Decide preset by unit and compute pt via canonical factors
        if (unitCode === 1) { // mm
            displayText = '3'; // 3mm
            pt = 3 * getPtFactorFromUnitCode(1);
        } else if (unitCode === 5) { // Q/H
            displayText = '12'; // 12H
            pt = 12 * getPtFactorFromUnitCode(5);
        } else if (unitCode === 2) { // pt
            displayText = '0.125 inches'; // show inches text (UI意図どおり)
            pt = 0.125 * 72.0; // 0.125in = 9pt
        } else {
            // Fallback: treat as 3mm equivalent for any other unit
            displayText = displayText || '3';
            pt = 3 * getPtFactorFromUnitCode(1);
        }
        return {
            pt: pt,
            displayText: displayText,
            disabled: true
        };
    }

    // Normal (non-bleed) case: multiply by current unit factor
    var n = parseFloat(displayText);
    if (isNaN(n)) n = 0;
    pt = n * getPtFactorFromUnitCode(unitCode);
    return {
        pt: pt,
        displayText: displayText,
        disabled: false
    };
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
        try {
            if (__previewDebounceTask) app.cancelTask(__previewDebounceTask);
        } catch (_) {}
        try {
            renderPreview(app.activeDocument, choice);
        } catch (_) {}
    } else {
        schedulePreview(choice, 75);
    }
}

function clearPreview() {
    try {
        var doc = app.activeDocument;
        var names = [LABELS.previewLayer[lang], "プレビュー", "Preview", "_preview"]; // support legacy names
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var nm = doc.layers[i].name;
            for (var j = 0; j < names.length; j++) {
                if (nm === names[j]) {
                    try {
                        doc.layers[i].remove();
                    } catch (e) {}
                    break;
                }
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
            // Avoid switching active artboard for preview / プレビューではアクティブAB切替を行わない
        } catch (e) {}
        var ab = doc.artboards[idx];
        var abRect = ab.artboardRect;
        var abWidth = abRect[2] - abRect[0];
        var abHeight = abRect[1] - abRect[3];
        var o = choice.offset || 0;
        var rect = previewLayer.pathItems.rectangle(
            abRect[1] + o, abRect[0] - o, abWidth + o * 2, abHeight + o * 2
        );


        // Map-driven color fill handlers / マップ駆動の塗り設定
        var previewFillers = {
            'none': function() {
                rect.filled = false;
                rect.stroked = false;
            },
            'k100': function() {
                rect.filled = true;
                rect.fillColor = createBlackColor(doc);
                rect.stroked = false;
                rect.opacity = 15;
            },
            'hex': function() {
                var col = parseCustomColor(doc, choice.customValue);
                if (col) {
                    rect.filled = true;
                    rect.fillColor = col;
                    rect.stroked = false;
                    rect.opacity = 100;
                } else {
                    rect.filled = false;
                    rect.stroked = false;
                }
            },
            'cmyk': function() {
                var c = choice.customCMYK && choice.customCMYK.c,
                    m = choice.customCMYK && choice.customCMYK.m,
                    y = choice.customCMYK && choice.customCMYK.y,
                    k = choice.customCMYK && choice.customCMYK.k;
                var ok = (typeof c === 'number' && !isNaN(c)) && (typeof m === 'number' && !isNaN(m)) && (typeof y === 'number' && !isNaN(y)) && (typeof k === 'number' && !isNaN(k));
                if (!ok) {
                    rect.filled = false;
                    rect.stroked = false;
                    return;
                }
                var col;
                if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
                    var rgb = cmykToRgb(c, m, y, k);
                    col = makeRGB(rgb[0], rgb[1], rgb[2]);
                } else {
                    col = makeCMYK(c, m, y, k);
                }
                rect.filled = true;
                rect.fillColor = col;
                rect.stroked = false;
                rect.opacity = 100;
            }
        };
        (previewFillers[choice.colorMode] || previewFillers[ColorMode.NONE])();

        // Common preview stroke (visibility)
        rect.stroked = true;
        rect.strokeWidth = 1;
        try {
            rect.strokeDashes = [6, 4];
        } catch (e) {}
        rect.strokeColor = getPreviewStrokeColor(doc);

        rect.name = LABELS.previewRect[lang];
        rect.selected = false;
        if (choice.zOrder === 'front') rect.zOrder(ZOrderMethod.BRINGTOFRONT);
        else if (choice.zOrder === 'back') rect.zOrder(ZOrderMethod.SENDTOBACK);
    }

    // --- Draw previews for target scope ---
    if (choice.target === 'all') {
        for (var i = 0; i < doc.artboards.length; i++) previewOne(i);
        try {
            doc.artboards.setActiveArtboardIndex(prevIndex);
        } catch (e) {}
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

    // Limit column widths to keep layout balanced

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
            try {
                requestPreview(buildChoiceFromUI(), true);
            } catch (_) {}
        }
    });

    var unitLabel = getCurrentUnitLabel();
    offsetRow.add('statictext', undefined, unitLabel);

    // Bleed checkbox row
    var bleedRow = offsetPanel.add('group');
    bleedRow.orientation = 'row';
    bleedRow.alignChildren = 'center';
    bleedRow.alignment = 'center';
    var cbBleed = bleedRow.add('checkbox', undefined, LABELS.bleed[lang]);
    cbBleed.alignment = 'center';
    cbBleed.value = false; // default OFF

    // Preserve user's manual offset when toggling Bleed
    var __lastUserOffsetText = '0';

    function applyBleedPreset(toPreview) {
        var unitCode = getCurrentUnitCode();
        var res = resolveOffsetToPt(offsetInput.text, unitCode, true);
        try {
            offsetInput.text = res.displayText;
        } catch (e) {}
        offsetInput.enabled = !res.disabled; // disabled when bleed ON
        if (toPreview === true) {
            updatePreview();
        }
    }

    function removeBleedPreset(toPreview) {
        try {
            offsetInput.text = __lastUserOffsetText;
        } catch (e) {}
        offsetInput.enabled = true;
        if (toPreview === true) {
            updatePreview();
        }
    }

    dlg.onShow = function() {
        try {
            // Focus offset field first
            offsetInput.active = true;
        } catch (e) {}
        // Shift dialog position once on show
        shiftDialogPositionOnce(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        // Remember initial value and set field state
        __lastUserOffsetText = String(offsetInput.text);
        if (cbBleed.value) {
            applyBleedPreset(false);
        } else {
            offsetInput.enabled = true;
        }
        // Render initial preview
        updatePreview();
    };

    // Add new panel for color
    var colorPanel = rightCol.add('panel', undefined, LABELS.colorTitle[lang]);
    colorPanel.orientation = 'column';
    colorPanel.alignChildren = 'left';
    colorPanel.margins = [15, 20, 15, 10];

    var noneRadio = colorPanel.add('radiobutton', undefined, LABELS.colorNone[lang]);
    var k100Radio = colorPanel.add('radiobutton', undefined, LABELS.colorK100[lang]);
    var specifiedRadio = colorPanel.add('radiobutton', undefined, LABELS.colorSpecified[lang]);

    // Custom value text field (UI only; logic TBD)
    var customRow = colorPanel.add('group');
    customRow.orientation = 'row';
    customRow.alignment = 'left';
    var customInput = customRow.add('edittext', undefined, '');
    customInput.characters = 14; // narrower to avoid column growth

    // (spacing) separate HEX and CMYK blocks
    var _hexCmykSpacer = colorPanel.add('group');
    try {
        _hexCmykSpacer.preferredSize.height = 6;
    } catch (e) {}

    // CMYK mode radio
    var cmykRadio = colorPanel.add('radiobutton', undefined, LABELS.colorCustomCMYK[lang]);

    // Custom CMYK input fields (horizontal)
    var cmykRow = colorPanel.add('group');
    cmykRow.orientation = 'row';
    cmykRow.alignment = 'left';
    cmykRow.spacing = 6;

    var lblC = cmykRow.add('statictext', undefined, 'C');
    var etC = cmykRow.add('edittext', undefined, '');
    etC.characters = 3; // width 3
    var lblM = cmykRow.add('statictext', undefined, 'M');
    var etM = cmykRow.add('edittext', undefined, '');
    etM.characters = 3; // width 3
    var lblY = cmykRow.add('statictext', undefined, 'Y');
    var etY = cmykRow.add('edittext', undefined, '');
    etY.characters = 3; // width 3
    var lblK = cmykRow.add('statictext', undefined, 'K');
    var etK = cmykRow.add('edittext', undefined, '');
    etK.characters = 3; // width 3

    // --- CMYK validation helpers (empty→0, clamp 0–100, red text warning) ---
    function setEtWarn(et, warn) {
        try {
            var g = et.graphics;
            var pen = g.newPen(g.PenType.SOLID_COLOR, warn ? [1, 0, 0] : [0, 0, 0], 1);
            g.foregroundColor = pen; // text color as fallback to 'red border'
            et.helpTip = warn ? '0–100 の範囲にしてください（未入力は 0 として扱います）' : '';
        } catch (e) {}
    }

    function validateCmykField(et) {
        try {
            var t = String(et.text || '');
            if (t === '') {
                setEtWarn(et, false);
                return;
            } // typing phase, don't warn
            var n = parseFloat(t);
            var warn = (isNaN(n) || n < 0 || n > 100);
            setEtWarn(et, warn);
        } catch (e) {}
    }

    function clampCmykField(et) {
        try {
            var t = String(et.text || '');
            var n = parseFloat(t);
            if (isNaN(n)) n = 0; // empty/invalid -> 0
            n = _clamp(n, 0, 100);
            et.text = String(n);
            setEtWarn(et, false);
        } catch (e) {}
    }


    // Update preview on typing/commit with validation
    etC.onChanging = function() {
        validateCmykField(etC);
        updatePreview();
    };
    etC.onChange = function() {
        clampCmykField(etC);
        updatePreview();
    };

    etM.onChanging = function() {
        validateCmykField(etM);
        updatePreview();
    };
    etM.onChange = function() {
        clampCmykField(etM);
        updatePreview();
    };

    etY.onChanging = function() {
        validateCmykField(etY);
        updatePreview();
    };
    etY.onChange = function() {
        clampCmykField(etY);
        updatePreview();
    };

    etK.onChanging = function() {
        validateCmykField(etK);
        updatePreview();
    };
    etK.onChange = function() {
        clampCmykField(etK);
        updatePreview();
    };

    // Default selection
    k100Radio.value = true;

    // Enable custom field only when "Custom" is selected

    function syncCustomEnable() {
        var hexOn = !!specifiedRadio.value; // HEX モード
        var cmykOn = !!cmykRadio.value; // CMYK モード
        // HEX テキスト欄
        customInput.enabled = hexOn;
        // CMYK 入力欄
        try {
            etC.enabled = cmykOn;
            lblC.enabled = cmykOn;
            etM.enabled = cmykOn;
            lblM.enabled = cmykOn;
            etY.enabled = cmykOn;
            lblY.enabled = cmykOn;
            etK.enabled = cmykOn;
            lblK.enabled = cmykOn;
            if (!cmykOn) {
                setEtWarn(etC, false);
                setEtWarn(etM, false);
                setEtWarn(etY, false);
                setEtWarn(etK, false);
            }
        } catch (e) {}
    }

    syncCustomEnable();

    // Add new panel for zOrder
    var zOrderPanel = leftCol.add('panel', undefined, LABELS.zorderTitle[lang]);
    zOrderPanel.orientation = 'column';
    zOrderPanel.alignChildren = 'left';
    zOrderPanel.margins = [15, 20, 15, 10];

    var frontRadio = zOrderPanel.add('radiobutton', undefined, LABELS.front[lang]);
    var backRadio = zOrderPanel.add('radiobutton', undefined, LABELS.back[lang]);
    var bgLayerRadio = zOrderPanel.add('radiobutton', undefined, LABELS.bg[lang]);

    backRadio.value = true;

    // Add new panel for target
    var targetPanel = leftCol.add('panel', undefined, LABELS.targetTitle[lang]);
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
        var colorMode = (function() {
            if (noneRadio.value) return ColorMode.NONE;
            if (k100Radio.value) return ColorMode.K100;
            if (specifiedRadio.value) return ColorMode.HEX;
            if (cmykRadio.value) return ColorMode.CMYK;
            return ColorMode.NONE;
        })();

        var zOrder = frontRadio.value ? 'front' : (backRadio.value ? 'back' : (bgLayerRadio.value ? 'bg' : 'back'));
        var target = currentRadio.value ? 'current' : (allRadio.value ? 'all' : 'current');

        // offset計算を一元化
        var unitCode = getCurrentUnitCode();
        var resolved = resolveOffsetToPt(offsetInput.text, unitCode, !!cbBleed.value);
        var offsetPt = resolved.pt;

        var customValue = '';
        try {
            customValue = String(customInput.text || '').replace(/^\s+|\s+$/g, '');
        } catch (e) {}

        var cmykObj = {
            c: 0,
            m: 0,
            y: 0,
            k: 0
        };
        try {
            var cTmp = parseFloat(etC.text);
            if (isNaN(cTmp)) cTmp = 0;
            cmykObj.c = _clamp(cTmp, 0, 100);
            var mTmp = parseFloat(etM.text);
            if (isNaN(mTmp)) mTmp = 0;
            cmykObj.m = _clamp(mTmp, 0, 100);
            var yTmp = parseFloat(etY.text);
            if (isNaN(yTmp)) yTmp = 0;
            cmykObj.y = _clamp(yTmp, 0, 100);
            var kTmp = parseFloat(etK.text);
            if (isNaN(kTmp)) kTmp = 0;
            cmykObj.k = _clamp(kTmp, 0, 100);
        } catch (e) {}

        return {
            colorMode: colorMode,
            customValue: customValue, // HEX文字列
            customCMYK: cmykObj, // CMYK値
            offset: offsetPt,
            zOrder: zOrder,
            target: target,
            bleed: !!cbBleed.value
        };
    }

    function updatePreview() {
        try {
            requestPreview(buildChoiceFromUI(), false);
        } catch (e) {}
    }

    offsetInput.onChanging = updatePreview;
    offsetInput.onChange = updatePreview;

    noneRadio.onClick = function() {
        // Enforce exclusivity
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        noneRadio.value = true;
        syncCustomEnable();
        updatePreview();
    };

    k100Radio.onClick = function() {
        noneRadio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        k100Radio.value = true;
        syncCustomEnable();
        updatePreview();
    };

    specifiedRadio.onClick = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        cmykRadio.value = false;
        specifiedRadio.value = true;
        syncCustomEnable();
        try {
            requestPreview(buildChoiceFromUI(), true);
        } catch (_) {
            updatePreview();
        }
    };

    cmykRadio.onClick = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = true;
        syncCustomEnable();
        updatePreview();
    };

    noneRadio.onChanging = function() {
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        noneRadio.value = true;
        syncCustomEnable();
        updatePreview();
    };
    k100Radio.onChanging = function() {
        noneRadio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        k100Radio.value = true;
        syncCustomEnable();
        updatePreview();
    };
    specifiedRadio.onChanging = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        cmykRadio.value = false;
        specifiedRadio.value = true;
        syncCustomEnable();
        try {
            requestPreview(buildChoiceFromUI(), true);
        } catch (_) {
            updatePreview();
        }
    };
    cmykRadio.onChanging = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = true;
        syncCustomEnable();
        updatePreview();
    };

    frontRadio.onClick = updatePreview;
    backRadio.onClick = updatePreview;
    bgLayerRadio.onClick = updatePreview;

    currentRadio.onClick = updatePreview;
    allRadio.onClick = updatePreview;
    currentRadio.onChanging = updatePreview;
    allRadio.onChanging = updatePreview;

    cbBleed.onClick = function() {
        if (cbBleed.value) {
            __lastUserOffsetText = String(offsetInput.text);
            applyBleedPreset(true);
        } else {
            removeBleedPreset(true);
        }
    };

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
    if (noneRadio.value) colorMode = ColorMode.NONE;
    else if (k100Radio.value) colorMode = ColorMode.K100;
    else if (specifiedRadio.value) colorMode = ColorMode.HEX;
    else if (cmykRadio.value) colorMode = ColorMode.CMYK;

    var unitCode2 = getCurrentUnitCode();
    var resolvedFinal = resolveOffsetToPt(offsetInput.text, unitCode2, !!cbBleed.value);
    var offset = resolvedFinal.pt;

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

    var customValueFinal = '';
    try {
        customValueFinal = String(customInput.text || '').replace(/^\s+|\s+$/g, '');
    } catch (e) {}
    return {
        colorMode: colorMode,
        customValue: customValueFinal,
        customCMYK: (function() {
            try {
                var cTmp = parseFloat(etC.text);
                if (isNaN(cTmp)) cTmp = 0;
                cTmp = _clamp(cTmp, 0, 100);
                var mTmp = parseFloat(etM.text);
                if (isNaN(mTmp)) mTmp = 0;
                mTmp = _clamp(mTmp, 0, 100);
                var yTmp = parseFloat(etY.text);
                if (isNaN(yTmp)) yTmp = 0;
                yTmp = _clamp(yTmp, 0, 100);
                var kTmp = parseFloat(etK.text);
                if (isNaN(kTmp)) kTmp = 0;
                kTmp = _clamp(kTmp, 0, 100);
                return {
                    c: cTmp,
                    m: mTmp,
                    y: yTmp,
                    k: kTmp
                };
            } catch (e) {
                return {
                    c: 0,
                    m: 0,
                    y: 0,
                    k: 0
                };
            }
        })(),
        offset: offset,
        zOrder: zOrder,
        target: target,
        bleed: !!cbBleed.value,
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

    // Map-driven color fill for final drawing / 本描画のマップ駆動塗り設定
    var drawFillers = {
        'none': function() {
            rect.filled = false;
            rect.stroked = false;
        },
        'k100': function() {
            rect.filled = true;
            rect.fillColor = createBlackColor(doc);
            rect.stroked = false;
            rect.opacity = 15;
        },
        'hex': function() {
            var col = parseCustomColor(doc, choice.customValue);
            if (col) {
                rect.filled = true;
                rect.fillColor = col;
                rect.stroked = false;
                rect.opacity = 100;
            } else {
                rect.filled = false;
                rect.stroked = false;
            }
        },
        'cmyk': function() {
            var c = choice.customCMYK && choice.customCMYK.c,
                m = choice.customCMYK && choice.customCMYK.m,
                y = choice.customCMYK && choice.customCMYK.y,
                k = choice.customCMYK && choice.customCMYK.k;
            var ok = (typeof c === 'number' && !isNaN(c)) && (typeof m === 'number' && !isNaN(m)) && (typeof y === 'number' && !isNaN(y)) && (typeof k === 'number' && !isNaN(k));
            if (!ok) {
                rect.filled = false;
                rect.stroked = false;
                return;
            }
            var col;
            if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
                var rgb = cmykToRgb(c, m, y, k);
                col = makeRGB(rgb[0], rgb[1], rgb[2]);
            } else {
                col = makeCMYK(c, m, y, k);
            }
            rect.filled = true;
            rect.fillColor = col;
            rect.stroked = false;
            rect.opacity = 100;
        }
    };
    (drawFillers[choice.colorMode] || drawFillers[ColorMode.NONE])();

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