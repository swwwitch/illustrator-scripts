#target illustrator
#targetengine "DialogEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

アートボードサイズを調整

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddPageNumberFromTextSelection.jsx

### 概要：

- アクティブまたは全アートボードと同サイズの長方形を、オフセットを考慮して描画します。
- カラー（なし／K100 15%／HEX／CMYK）、重ね順（最前面／最背面／bgレイヤー）を指定でき、ライブプレビューで確認できます。

### 主な機能：

- オフセット指定（裁ち落としプリセット：3mm／12H／0.125in）
- カラー指定（None／K100 15%／HEX／CMYK）
- 重ね順（Front／Back／bgレイヤー）
- 対象範囲（作業アートボード／すべて）
- プレビュー（1ptの破線、50%トーン、専用レイヤー）
- ダイアログ位置・不透明度の設定（位置記憶に #targetengine 利用）

### note：

- 「ディム表示（指定）」は UI のみ実装。描画ロジックは未実装です。

### 更新履歴：

- v1.0 (20250820) : 初期バージョン
- v1.1 (20250821) : ダイアログの位置・透明度設定を追加。プレビュー破線を1ptに変更。オフセット単位を現在の単位に合わせるよう修正。
- v1.2 (20250821) : CMYKを独立UI化し、parseCustomColor()から旧CMYK文字列解釈を削除。
- v1.3 (20250821) : UIまわりをブラッシュアップ
- v1.4 (20250821) : 微調整
- v1.5（20250824）：カラーまわりのロジックを調整

---

### Script Name:

Adjust Artboard Size

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddPageNumberFromTextSelection.jsx

### Overview:

- Draws rectangles that match the active or all artboards, with optional offset.
- Supports color (None / K100 15% / HEX / CMYK), stacking order (Front / Back / bg layer), and live preview.

### Key Features:

- Offset with Bleed presets (3mm / 12H / 0.125in)
- Color modes (None / K100 15% / HEX / CMYK)
- Z-order (Front / Back / bg layer)
- Target scope (Current artboard / All artboards)
- Preview (1pt dashed stroke, 50% tone, dedicated layer)
- Dialog position & opacity settings (position memory via #targetengine)

### Notes:

- "Dim" mode is UI-only for now; rendering logic is not implemented yet.

### Changelog:

- v1.0 (20250820): Initial version
- v1.1 (20250821): Added dialog position/opacity settings. Changed preview stroke to 1pt dashed. Offset unit now matches current ruler setting.
- v1.2 (20250821): Removed legacy CMYK string parsing from parseCustomColor() since CMYK has its own UI.
- v1.3 (20250821): UI improvements, including better layout and error handling.
- v1.4 (20250821): Minor adjustments.
- v1.5（20250824）

*/

var SCRIPT_VERSION = "v1.5";

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
        ja: "アートボードサイズの長方形を描画 " + SCRIPT_VERSION,
        en: "Draw Artboard-Sized Rectangle " + SCRIPT_VERSION
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
        ja: "例: #FF0000",
        en: "e.g., #FF0000"
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
        ja: "作業中のアートボードのみ",
        en: "Current Artboard Only"
    },
    allAB: {
        ja: "すべてのアートボード",
        en: "All Artboards"
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

// ===== Preview timing (tunable) =====
// 入力中のプレビュー遅延（タイプしやすさ優先）/ Delay during typing
var PREVIEW_DELAY_TYPING_MS = 110; // recommend 100–120ms


/* =========================================
 * DialogPersist util (extractable)
 * ダイアログの不透明度・初期位置・位置記憶を共通化するユーティリティ。
 * 使い方:
 *   DialogPersist.setOpacity(dlg, 0.95);
 *   DialogPersist.restorePosition(dlg, "__YourDialogKey", offsetX, offsetY);
 *   DialogPersist.rememberOnMove(dlg, "__YourDialogKey");
 *   DialogPersist.savePosition(dlg, "__YourDialogKey"); // 閉じる直前などに
 * ========================================= */
(function(g){
    if (!g.DialogPersist) {
        g.DialogPersist = {
            setOpacity: function(dlg, v){
                try { dlg.opacity = v; } catch (e) {}
            },
            _getSaved: function(key){
                return g[key] && g[key].length === 2 ? g[key] : null;
            },
            _setSaved: function(key, loc){
                g[key] = [loc[0], loc[1]];
            },
            _clampToScreen: function(loc){
                try {
                    var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
                    var x = Math.max(vb[0] + 10, Math.min(loc[0], vb[2] - 10));
                    var y = Math.max(vb[1] + 10, Math.min(loc[1], vb[3] - 10));
                    return [x, y];
                } catch (e) {
                    return loc;
                }
            },
            restorePosition: function(dlg, key, offsetX, offsetY){
                var loc = this._getSaved(key);
                try {
                    if (loc) {
                        dlg.location = this._clampToScreen(loc);
                    } else {
                        var l = dlg.location;
                        dlg.location = [l[0] + (offsetX|0), l[1] + (offsetY|0)];
                    }
                } catch (e) {}
            },
            rememberOnMove: function(dlg, key){
                var self = this;
                dlg.onMove = function(){
                    try {
                        self._setSaved(key, [dlg.location[0], dlg.location[1]]);
                    } catch (e) {}
                };
            },
            savePosition: function(dlg, key){
                try { this._setSaved(key, [dlg.location[0], dlg.location[1]]); } catch (e) {}
            }
        };
    }
})($.global);

/* 入力欄の強調表示ヘルパー / Helper to highlight EditText fields */
function setEditHighlight(et, on) {
    try {
        var g = et.graphics;
        if (on) {
            // Light yellow highlight
            g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, [1, 1, 0.85]);
            g.foregroundColor = g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.2, 0], 1);
        } else {
            // Reset to default-looking white
            g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, [1, 1, 1]);
            g.foregroundColor = g.newPen(g.PenType.SOLID_COLOR, [0, 0, 0], 1);
        }
        et.notify('onDraw'); // redraw
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

// --- Common fill applier for both preview and final draw ---
/*
 * applyFillByMode(doc, rect, mode, payload, opts)
 * - mode: 'none' | 'k100' | 'hex' | 'cmyk'
 * - payload: { customValue: String, customCMYK: {c,m,y,k} }
 * - opts: { k100Opacity: Number }
 */
function applyFillByMode(doc, rect, mode, payload, opts) {
    try {
        var k100Opacity = (opts && typeof opts.k100Opacity === 'number') ? opts.k100Opacity : 15;
        if (mode === 'none') {
            rect.filled = false;
            rect.stroked = false;
            return;
        }
        if (mode === 'k100') {
            rect.filled = true;
            rect.fillColor = createBlackColor(doc);
            rect.stroked = false;
            rect.opacity = k100Opacity;
            return;
        }
        if (mode === 'hex') {
            var col = parseCustomColor(doc, payload && payload.customValue);
            if (col) {
                rect.filled = true;
                rect.fillColor = col;
                rect.stroked = false;
                rect.opacity = 100;
            } else {
                rect.filled = false;
                rect.stroked = false;
            }
            return;
        }
        if (mode === 'cmyk') {
            var c = payload && payload.customCMYK && payload.customCMYK.c,
                m = payload && payload.customCMYK && payload.customCMYK.m,
                y = payload && payload.customCMYK && payload.customCMYK.y,
                k = payload && payload.customCMYK && payload.customCMYK.k;
            var ok = (typeof c === 'number' && !isNaN(c)) &&
                (typeof m === 'number' && !isNaN(m)) &&
                (typeof y === 'number' && !isNaN(y)) &&
                (typeof k === 'number' && !isNaN(k));
            if (!ok) {
                rect.filled = false;
                rect.stroked = false;
                return;
            }
            var color;
            if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
                var rgb = cmykToRgb(c, m, y, k);
                color = makeRGB(rgb[0], rgb[1], rgb[2]);
            } else {
                color = makeCMYK(c, m, y, k);
            }
            rect.filled = true;
            rect.fillColor = color;
            rect.stroked = false;
            rect.opacity = 100;
            return;
        }
        // Fallback
        rect.filled = false;
        rect.stroked = false;
    } catch (e) {}
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

        // Expand shorthand HEX notations
        // #RGB → #RRGGBB, #RR → #RRRRRR, #R → #RRRRRR
        if (s.charAt(0) === '#') {
            if (s.length === 4) { // #RGB
                var r1 = s.charAt(1), g1 = s.charAt(2), b1 = s.charAt(3);
                s = '#' + r1 + r1 + g1 + g1 + b1 + b1;
            } else if (s.length === 3) { // #RR
                var r2 = s.charAt(1), r3 = s.charAt(2);
                s = '#' + r2 + r3 + r2 + r3 + r2 + r3;
            } else if (s.length === 2) { // #R
                var r4 = s.charAt(1);
                s = '#' + r4 + r4 + r4 + r4 + r4 + r4; // #3 → #333333
            }
        }

        // #RRGGBB
        if (s.charAt(0) === '#' && s.length === 7) {
            var r = parseInt(s.substr(1, 2), 16);
            var g = parseInt(s.substr(3, 2), 16);
            var b = parseInt(s.substr(5, 2), 16);
            if (!isNaN(r) && !isNaN(g) && !isNaN(b)) return makeRGB(r, g, b);
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

/* =========================================
 * PreviewHistory util (extractable)
 * ヒストリーを残さないプレビューのための小さなユーティリティ。
 * 他スクリプトでもこのブロックをコピペすれば再利用できます。
 * 使い方:
 *   PreviewHistory.start();     // ダイアログ表示時などにカウンタ初期化
 *   PreviewHistory.bump();      // プレビュー描画ごとにカウント(+1)
 *   PreviewHistory.undo();      // 閉じる/キャンセル時に一括Undo
 *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
 * ========================================= */
(function(g){
    if (!g.PreviewHistory) {
        g.PreviewHistory = {
            start: function(){
                g.__previewUndoCount = 0;
            },
            bump: function(){
                g.__previewUndoCount = (g.__previewUndoCount | 0) + 1;
            },
            undo: function(){
                var n = g.__previewUndoCount | 0;
                try {
                    for (var i = 0; i < n; i++) app.executeMenuCommand('undo');
                } catch (e) {}
                g.__previewUndoCount = 0;
            },
            cancelTask: function(taskId){
                try { if (taskId) app.cancelTask(taskId); } catch (e) {}
            }
        };
    }
})($.global);


var __previewDebounceTask = null;

function schedulePreview(choice, delayMs) {
    try {
        if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
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
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
        } catch (_) {}
        try {
            renderPreview(app.activeDocument, choice);
        } catch (_) {}
    } else {
        schedulePreview(choice, PREVIEW_DELAY_TYPING_MS);
    }
}

function clearPreview(removeLayer) {
    try {
        var doc = app.activeDocument;
        var names = [LABELS.previewLayer[lang], "プレビュー", "Preview", "_preview"]; // legacy names
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var layer = doc.layers[i];
            var nm = layer.name;
            for (var j = 0; j < names.length; j++) {
                if (nm === names[j]) {
                    if (removeLayer) {
                        try {
                            layer.remove();
                        } catch (e) {}
                    } else {
                        // live update: just hide items instead of removing layer itself
                        try {
                            for (var k = layer.pathItems.length - 1; k >= 0; k--) {
                                try {
                                    layer.pathItems[k].hidden = true;
                                } catch (_) {}
                            }
                        } catch (_) {}
                    }
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

// Helper to fetch or create a preview rectangle for a specific artboard index
function getOrCreatePreviewRect(previewLayer, idx, top, left, width, height) {
    // Try to reuse existing item named with index suffix
    var nameBase = LABELS.previewRect[lang] + "#" + idx;
    var item = null;
    try {
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var it = previewLayer.pathItems[i];
            if (it.name === nameBase) {
                item = it;
                break;
            }
        }
    } catch (e) {}
    if (!item) {
        // create once
        item = previewLayer.pathItems.rectangle(top, left, width, height);
        item.name = nameBase;
    } else {
        // update geometry in-place
        try {
            item.top = top;
            item.left = left;
            item.width = width;
            item.height = height;
        } catch (e) {}
        try {
            item.hidden = false;
        } catch (e) {}
    }
    return item;
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
    // Hide existing preview items instead of deleting the whole layer
    clearPreview(false);
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
        var rect = getOrCreatePreviewRect(
            previewLayer,
            idx,
            abRect[1] + o,
            abRect[0] - o,
            abWidth + o * 2,
            abHeight + o * 2
        );

        // Unified fill application (preview)
        applyFillByMode(doc, rect, choice.colorMode, {
            customValue: choice.customValue,
            customCMYK: choice.customCMYK
        }, {
            k100Opacity: 15
        });

        // Common preview stroke (visibility)
        rect.stroked = true;
        rect.strokeWidth = 1;
        try {
            rect.strokeDashes = [6, 4];
        } catch (e) {}
        rect.strokeColor = getPreviewStrokeColor(doc);

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
        // Hide non-target preview items to avoid deletions during typing
        var maxIdx = (choice.target === 'all') ? doc.artboards.length - 1 : doc.artboards.getActiveArtboardIndex();
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var it = previewLayer.pathItems[i];
            // items are named like "__Preview_...#<idx>"
            var m = /#(\d+)$/.exec(it.name || "");
            if (m) {
                var idn = parseInt(m[1], 10);
                if (choice.target === 'all') {
                    if (idn < 0 || idn > doc.artboards.length - 1) {
                        try {
                            it.hidden = true;
                        } catch (_) {}
                    }
                } else {
                    var active = doc.artboards.getActiveArtboardIndex();
                    if (idn !== active) {
                        try {
                            it.hidden = true;
                        } catch (_) {}
                    }
                }
            }
        }
    } catch (e) {}

    try {
        if (prevCS !== null) app.coordinateSystem = prevCS;
    } catch (e) {}
    PreviewHistory.bump();
    app.redraw();
}

function showDialog() {
    var dlg = new Window('dialog', LABELS.dialogTitle[lang]);
    DialogPersist.setOpacity(dlg, DIALOG_OPACITY);
    var __DLG_KEY = "__SmartDrawABRect_Dialog"; // unique key per dialog
    if ($.global[__DLG_KEY] === undefined) $.global[__DLG_KEY] = null; // ensure slot
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
        DialogPersist.restorePosition(dlg, __DLG_KEY, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        try {
            offsetInput.active = true;
        } catch (e) {}
        // Remember initial value and set field state
        __lastUserOffsetText = String(offsetInput.text);
        if (cbBleed.value) {
            applyBleedPreset(false);
        } else {
            offsetInput.enabled = true;
        }
        PreviewHistory.start(); // reset preview history counter
        // Render initial preview
        updatePreviewCommit();
    };
    DialogPersist.rememberOnMove(dlg, __DLG_KEY);

    // Add new panel for color
    var colorPanel = rightCol.add('panel', undefined, LABELS.colorTitle[lang]);
    colorPanel.orientation = 'column';
    colorPanel.alignChildren = 'left';
    colorPanel.margins = [15, 20, 15, 10];
    colorPanel.spacing = 10; // increase vertical gap between rows

    var noneRadio = colorPanel.add('radiobutton', undefined, LABELS.colorNone[lang]);
    var k100Radio = colorPanel.add('radiobutton', undefined, LABELS.colorK100[lang]);

    // HEX radio + input on the same row
    var hexRow = colorPanel.add('group');
    hexRow.orientation = 'row';
    hexRow.alignment = 'left';
    hexRow.alignChildren = ['left', 'center'];
    hexRow.spacing = 6;
    var specifiedRadio = hexRow.add('radiobutton', undefined, LABELS.colorSpecified[lang]);
    var customInput = hexRow.add('edittext', undefined, '#');
    customInput.characters = 14; // narrower to avoid column growth

    // --- HEX validation & feedback ---
    function setHexWarn(et, warn, msg) {
        try {
            var g = et.graphics;
            var pen = g.newPen(g.PenType.SOLID_COLOR, warn ? [1, 0, 0] : [0, 0, 0], 1);
            g.foregroundColor = pen; // text color fallback for border
            if (warn) {
                et.helpTip = (lang === 'ja') ? (msg || '正しい #RRGGBB を入力してください') : (msg || 'Enter a valid #RRGGBB value');
            } else {
                et.helpTip = '';
            }
            et.notify('onDraw');
        } catch (e) {}
    }

    customInput.onChanging = function() {
        try {
            var t = String(customInput.text || '').replace(/\s+/g, '');
            if (t === '') {
                setHexWarn(customInput, false);
                updatePreviewTyping();
                return;
            }
            if (t === '#') {
                setHexWarn(customInput, true, (lang === 'ja') ? 'HEX未入力（# のみ）' : 'HEX not entered (# only)');
                updatePreviewTyping();
                return;
            }
            // Validate only exact #RRGGBB on-the-fly
            var valid = /^#([0-9a-fA-F]{6})$/.test(t);
            setHexWarn(customInput, !valid);
            updatePreviewTyping();
        } catch (e) {}
    };

    customInput.onChange = function() {
        try {
            var t = String(customInput.text || '').replace(/\s+/g, '');
            if (/^[0-9a-fA-F]{6}$/.test(t)) {
                t = '#' + t.toUpperCase();
                customInput.text = t;
                setHexWarn(customInput, false);
            } else if (/^#([0-9a-fA-F]{6})$/.test(t)) {
                customInput.text = ('#' + RegExp.$1.toUpperCase());
                setHexWarn(customInput, false);
            } else if (t === '#') {
                setHexWarn(customInput, true, (lang === 'ja') ? 'HEX未入力（# のみ）' : 'HEX not entered (# only)');
            } else {
                setHexWarn(customInput, true);
            }
            updatePreviewCommit();
        } catch (e) {}
    };


    // CMYK mode radio
    var cmykRadio = colorPanel.add('radiobutton', undefined, LABELS.colorCustomCMYK[lang]);

    // Custom CMYK input fields (two-row grid: labels on top, fields below)
    var cmykRow = colorPanel.add('group');
    cmykRow.orientation = 'column';
    cmykRow.alignment = 'left';
    cmykRow.spacing = 4;

    var cmykHead = cmykRow.add('group');
    cmykHead.orientation = 'row';
    cmykHead.alignChildren = ['left', 'center'];
    cmykHead.spacing = 10;

    var cmykInputs = cmykRow.add('group');
    cmykInputs.orientation = 'row';
    cmykInputs.alignChildren = ['left', 'center'];
    cmykInputs.spacing = 10;

    var colWidth = 40; // fixed width to align columns

    var lblC = cmykHead.add('statictext', undefined, '  C');
    lblC.preferredSize.width = colWidth;
    var lblM = cmykHead.add('statictext', undefined, '  M');
    lblM.preferredSize.width = colWidth;
    var lblY = cmykHead.add('statictext', undefined, '  Y');
    lblY.preferredSize.width = colWidth;
    var lblK = cmykHead.add('statictext', undefined, '  K');
    lblK.preferredSize.width = colWidth;

    var etC = cmykInputs.add('edittext', undefined, '');
    etC.characters = 3;
    etC.preferredSize.width = colWidth;
    var etM = cmykInputs.add('edittext', undefined, '');
    etM.characters = 3;
    etM.preferredSize.width = colWidth;
    var etY = cmykInputs.add('edittext', undefined, '');
    etY.characters = 3;
    etY.preferredSize.width = colWidth;
    var etK = cmykInputs.add('edittext', undefined, '');
    etK.characters = 3;
    etK.preferredSize.width = colWidth;

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
    // If the field currently holds exactly "0", clear it on focus for easier typing
    function clearZeroOnFocus(et) {
        try {
            et.addEventListener('focus', function() {
                try {
                    if (String(et.text) === '0') {
                        et.text = '';
                        // caret will be at the end by default
                    }
                } catch (e) {}
            });
        } catch (e) {}
    }

    // Prevent any leading-zero integer like "03" from being typed (but allow decimals like "0.5")
    function replaceZeroOnFirstDigit(et) {
        try {
            et.addEventListener('keydown', function(ev) {
                var k = String(ev.keyName || '');
                // Only care about single digit keys 0-9
                if (!/^[0-9]$/.test(k)) return;
                try {
                    var t = String(et.text || '');
                    // If user is composing a decimal number, do nothing here
                    if (/\./.test(t)) return;

                    // If the field is exactly "0" (or series of zeros), clear it **before** the new digit is inserted
                    // so typing "3" results directly in "3" (never "03").
                    if (/^0+$/.test(t)) {
                        et.text = '';
                        return; // let default insertion append the digit
                    }

                    // If there are leading zeros with other digits (e.g. "007"), normalize immediately.
                    if (/^0\d+$/.test(t)) {
                        et.text = t.replace(/^0+/, '');
                        // Do not prevent the default; allow the typed digit to insert normally after normalization.
                        return;
                    }
                } catch (e) {}
            });
        } catch (e) {}
    }

    // Bind common handlers to a CMYK EditText
    function bindCmykField(et) {
        et.onChanging = function() {
            try {
                var t = String(et.text || '');
                // Normalize leading zeros for integers: "03" -> "3", "007" -> "7"
                // Do NOT touch decimals like "0.5" (only pure digits)
                if (/^0\d+$/.test(t)) {
                    et.text = t.replace(/^0+/, '');
                    t = String(et.text || '');
                }
            } catch (e) {}
            validateCmykField(et);
            updatePreviewTyping();
        };
        et.onChange = function() {
            clampCmykField(et);
            updatePreviewCommit();
        };
        changeValueByArrowKey(et, function() {
            clampCmykField(et);
            updatePreviewTyping();
        });
    }

    // --- Hotkey guard: disable N/K/H/C while typing in fields ---
    var __hotkeyBlocked = {
        v: false
    };

    function _attachBlockOnFocusBlur(ctrl) {
        try {
            ctrl.addEventListener('focus', function() {
                __hotkeyBlocked.v = true;
            });
        } catch (e) {}
        try {
            ctrl.addEventListener('blur', function() {
                __hotkeyBlocked.v = false;
            });
        } catch (e) {}
    }
    // Block on all edit fields
    _attachBlockOnFocusBlur(offsetInput);
    _attachBlockOnFocusBlur(customInput);
    _attachBlockOnFocusBlur(etC);
    _attachBlockOnFocusBlur(etM);
    _attachBlockOnFocusBlur(etY);
    _attachBlockOnFocusBlur(etK);


    clearZeroOnFocus(etC);
    clearZeroOnFocus(etM);
    clearZeroOnFocus(etY);
    clearZeroOnFocus(etK);

    replaceZeroOnFirstDigit(etC);
    replaceZeroOnFirstDigit(etM);
    replaceZeroOnFirstDigit(etY);
    replaceZeroOnFirstDigit(etK);

    // --- Bind CMYK fields (common handlers)
    bindCmykField(etC);
    bindCmykField(etM);
    bindCmykField(etY);
    bindCmykField(etK);

    // Default selection
    k100Radio.value = true;

    // Enable custom field only when "Custom" is selected

    /* HEX 有効/無効を切替 / Enable-Disable HEX input */
    function setHexEnabled(on) {
        try {
            customInput.enabled = !!on;
        } catch (e) {}
    }

    /* CMYK 有効/無効を切替 / Enable-Disable CMYK inputs */
    function setCmykEnabled(on) {
        var v = !!on;
        try {
            etC.enabled = v;
            lblC.enabled = v;
            etM.enabled = v;
            lblM.enabled = v;
            etY.enabled = v;
            lblY.enabled = v;
            etK.enabled = v;
            lblK.enabled = v;
            if (!v) {
                setEtWarn(etC, false);
                setEtWarn(etM, false);
                setEtWarn(etY, false);
                setEtWarn(etK, false);
            }
        } catch (e) {}
    }

    /* ラジオ選択に応じて一括反映 / Apply enable states from radio values */
    function updateColorEnableFromRadios() {
        setHexEnabled(!!specifiedRadio.value);
        setCmykEnabled(!!cmykRadio.value);
    }


    updateColorEnableFromRadios();

    // Add new panel for zOrder
    var zOrderPanel = leftCol.add('panel', undefined, LABELS.zorderTitle[lang]);
    zOrderPanel.orientation = 'column';
    zOrderPanel.alignChildren = 'left';
    zOrderPanel.margins = [15, 20, 15, 10];

    var frontRadio = zOrderPanel.add('radiobutton', undefined, LABELS.front[lang]);
    var backRadio = zOrderPanel.add('radiobutton', undefined, LABELS.back[lang]);
    var bgLayerRadio = zOrderPanel.add('radiobutton', undefined, LABELS.bg[lang]);

    backRadio.value = true;

    // Add new panel for target (moved to right column)
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

    // 1枚しかない場合は「すべてのアートボード」をディム（無効化）
    if (abCount <= 1) {
        try {
            allRadio.enabled = false;
            allRadio.helpTip = (lang === 'ja') ? 'アートボードが1つのため選択できません' : 'Disabled: only one artboard exists';
        } catch (e) {}
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

    function updatePreviewTyping() {
        try {
            schedulePreview(buildChoiceFromUI(), PREVIEW_DELAY_TYPING_MS);
        } catch (e) {}
    }

    function updatePreviewCommit() {
        try {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
        } catch (_) {}
        try {
            renderPreview(app.activeDocument, buildChoiceFromUI());
        } catch (_) {}
    }

    function updatePreview() {
        try {
            requestPreview(buildChoiceFromUI(), false);
        } catch (e) {}
    }

    offsetInput.onChanging = updatePreviewTyping;
    offsetInput.onChange = updatePreviewCommit;

    noneRadio.onClick = function() {
        // Enforce exclusivity
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        noneRadio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, false);
        setEditHighlight(etC, false);
        updatePreviewCommit();
    };

    k100Radio.onClick = function() {
        noneRadio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        k100Radio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, false);
        setEditHighlight(etC, false);
        updatePreviewCommit();
    };

    specifiedRadio.onClick = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        cmykRadio.value = false;
        specifiedRadio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, true);
        setEditHighlight(etC, false);
        try {
            customInput.active = true;
        } catch (e) {}
        try {
            requestPreview(buildChoiceFromUI(), true);
        } catch (_) {
            updatePreviewCommit();
        }
    };

    cmykRadio.onClick = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, false);
        setEditHighlight(etC, true);
        try {
            etC.active = true;
        } catch (e) {}
        updatePreviewCommit();
    };

    noneRadio.onChanging = function() {
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        noneRadio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, false);
        setEditHighlight(etC, false);
        updatePreviewCommit();
    };
    k100Radio.onChanging = function() {
        noneRadio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = false;
        k100Radio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, false);
        setEditHighlight(etC, false);
        updatePreviewCommit();
    };
    specifiedRadio.onChanging = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        cmykRadio.value = false;
        specifiedRadio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, true);
        setEditHighlight(etC, false);
        try {
            customInput.active = true;
        } catch (e) {}
        try {
            requestPreview(buildChoiceFromUI(), true);
        } catch (_) {
            updatePreviewCommit();
        }
    };
    cmykRadio.onChanging = function() {
        noneRadio.value = false;
        k100Radio.value = false;
        specifiedRadio.value = false;
        cmykRadio.value = true;
        updateColorEnableFromRadios();
        setEditHighlight(customInput, false);
        setEditHighlight(etC, true);
        try {
            etC.active = true;
        } catch (e) {}
        updatePreviewCommit();
    };

    // --- Hotkeys: N/K/H/C to switch color mode radios ---
    function addColorHotkeys(dialog) {
        dialog.addEventListener('keydown', function(event) {
            if (__hotkeyBlocked.v) return; // typing in a field
            var key = (event && event.keyName) ? String(event.keyName).toUpperCase() : '';
            if (key === 'N') {
                noneRadio.notify('onClick');
                event.preventDefault();
            } else if (key === 'K') {
                k100Radio.notify('onClick');
                event.preventDefault();
            } else if (key === 'H') {
                specifiedRadio.notify('onClick');
                event.preventDefault();
            } else if (key === 'C') {
                cmykRadio.notify('onClick');
                event.preventDefault();
            }
        });
    }
    addColorHotkeys(dlg);

    // --- Hotkeys: Target (S/A) and Z-order (F/B/L) ---
    /* 作業アートボード/全体 と 重ね順 をホットキーで切替 / Toggle target & z-order via hotkeys */
    function addScopeAndZHotkeys(dialog) {
        dialog.addEventListener('keydown', function(event) {
            if (__hotkeyBlocked.v) return; // ignore when typing in fields
            var key = (event && event.keyName) ? String(event.keyName).toUpperCase() : '';

            // Target scope: S = current (Single), A = All
            if (key === 'S') {
                try {
                    currentRadio.notify('onClick');
                } catch (_) {}
                event.preventDefault();
                return;
            }
            if (key === 'A') {
                try {
                    if (allRadio.enabled) allRadio.notify('onClick');
                } catch (_) {}
                event.preventDefault();
                return;
            }

            // Z-order: F = Front, B = Back, L = bg Layer
            if (key === 'F') {
                try {
                    frontRadio.notify('onClick');
                } catch (_) {}
                event.preventDefault();
                return;
            }
            if (key === 'B') {
                try {
                    backRadio.notify('onClick');
                } catch (_) {}
                event.preventDefault();
                return;
            }
            if (key === 'L') {
                try {
                    bgLayerRadio.notify('onClick');
                } catch (_) {}
                event.preventDefault();
                return;
            }
        });
    }
    addScopeAndZHotkeys(dlg);

    frontRadio.onClick = updatePreviewCommit;
    backRadio.onClick = updatePreviewCommit;
    bgLayerRadio.onClick = updatePreviewCommit;

    currentRadio.onClick = updatePreviewCommit;
    allRadio.onClick = updatePreviewCommit;
    currentRadio.onChanging = updatePreviewCommit;
    allRadio.onChanging = updatePreviewCommit;

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
        DialogPersist.savePosition(dlg, __DLG_KEY);
        try {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
        } catch (e) {}
        PreviewHistory.undo();
        dlg.close(1);
    };
    cancelBtn.onClick = function() {
        DialogPersist.savePosition(dlg, __DLG_KEY);
        try {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
        } catch (e) {}
        PreviewHistory.undo();
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
 * 編集可能なレイヤーを取得（なければ作成）/ Get an editable layer or create one
 */
function getWritableLayer(doc) {
    try {
        var lyr = doc.activeLayer;
        if (lyr && !lyr.locked && lyr.visible) return lyr;
    } catch (e) {}
    // try find first unlocked & visible layer
    try {
        for (var i = 0; i < doc.layers.length; i++) {
            var l = doc.layers[i];
            if (!l.locked && l.visible) return l;
        }
    } catch (e) {}
    // last resort: create a new layer at top
    try {
        var nl = doc.layers.add();
        nl.name = "_auto_draw";
        nl.visible = true;
        nl.locked = false;
        return nl;
    } catch (e) {}
    return doc.activeLayer; // fallback
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
    // 見える＆編集可能に / Ensure editable
    layer.visible = true;
    layer.locked = false;
    try {
        layer.printable = true;
    } catch (e) {}
    // 最背面へ / Send to back of layer stack
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
    var targetLayer;
    if (choice.zOrder === 'bg') {
        targetLayer = getOrCreateBgLayer(doc);
    } else {
        targetLayer = getWritableLayer(doc);
    }
    // Create on the chosen layer to avoid "Target layer cannot be modified"
    var rect = targetLayer.pathItems.rectangle(
        abRect[1] + o,
        abRect[0] - o,
        abWidth + o * 2,
        abHeight + o * 2
    );
    // Ensure the new rect is selected before running a selection-based menu command
    try { rect.selected = true; } catch (e) {}
    // Apply Effect > Convert to Shape using the last-used parameters
    try { app.executeMenuCommand('Convert to Shape'); } catch (e) {}

    // Unified fill application (final draw)
    applyFillByMode(doc, rect, choice.colorMode, {
        customValue: choice.customValue,
        customCMYK: choice.customCMYK
    }, {
        k100Opacity: 15
    });

    rect.name = LABELS.rectName[lang];
    rect.selected = true;
    try {
        rect.hidden = false;
    } catch (e) {}
    try {
        targetLayer.visible = true;
        targetLayer.locked = false;
    } catch (e) {}

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