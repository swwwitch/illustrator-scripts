#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*
### スクリプト名：

背面に長方形を作成

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択オブジェクトの外接バウンディングボックスを基準に、オフセットを加えた長方形を生成
- プレビュー機能で即時確認可能、作成した長方形は常に最背面に配置

### 主な機能：

- オフセット（現在の定規単位に追従）
- 角丸（ライブエフェクト適用、非展開）
- 塗り/線 カラー指定（K100 / ホワイト / HEX / CMYK）
- 対象：個別／グループとして
- プレビュー（専用レイヤー、ヒストリーに残らない）
- ダイアログ位置・不透明度の記憶

### 処理の流れ：

1. 対象オブジェクトのバウンディングボックスを計算
2. オフセット・角丸・塗り/線を適用した長方形を作成
3. プレビューは専用レイヤーに生成、確定時に本番描画
4. 「テキストとグループ化」オプションで元テキストとグループ化可能

### 更新履歴：

- v1.0 (2025-08-22) : 初期バージョン
- v1.1 (2025-08-23) : プレビュー・カラー選択機能を追加
- v1.2 (2025-08-23) : 種別（塗り/線）、線幅指定、プリセット保存機能を追加

---
### Script Name:

Draw Rectangle Behind Selection

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Generate rectangles offset from the bounding box of selected objects
- Live preview with immediate feedback; created rectangles are always sent to back

### Key Features:

- Offset (follows current ruler units)
- Corner radius (applied via Live Effect, kept unexpanded)
- Fill/Stroke color options (K100 / White / HEX / CMYK)
- Target: Individual or as Group
- Preview on a dedicated layer (does not pollute history)
- Dialog position & opacity persistence

### Processing Flow:

1. Compute bounding box of target objects
2. Apply offset, corner radius, and fill/stroke settings to rectangle
3. Render preview to dedicated layer; finalize on OK
4. Optionally group rectangle with original text

### Update History:

- v1.0 (2025-08-22): Initial version
- v1.1 (2025-08-23): Added preview and color selection
- v1.2 (2025-08-23): Added type (fill/stroke), stroke width, and preset saving
*/

var SCRIPT_VERSION = "v1.2";

/*
 * Color mode constants / カラーモード定数
 */
var ColorMode = {
    NONE: 'none',
    K100: 'k100',
    WHITE: 'white',
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
        ja: "背面に長方形を作成" + " " + SCRIPT_VERSION,
        en: "Draw Rectangle Behind Selection" + " " + SCRIPT_VERSION
    },
    alertNoSelection: {
        ja: "オブジェクトを選択してください。",
        en: "Please select at least one object."
    },
    offsetTitle: {
        ja: "マージン",
        en: "Margins"
    },
    colorTitle: {
        ja: "塗り",
        en: "Fill"
    },
    roundTitle: {
        ja: "角丸",
        en: "Corner Radius"
    },
    targetTitle: {
        ja: "対象",
        en: "Target"
    },
    typeTitle: {
        ja: "種別",
        en: "Type"
    },
    strokeWidth: {
        ja: "線幅",
        en: "Stroke Width"
    },
    previewBounds: {
        ja: "テキストとグループ化",
        en: "Group with Text"
    },
    previewOutline: {
        ja: "テキストをアウトラインで計算",
        en: "Outline Text for Preview"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    previewLayer: {
        ja: "_プレビュー",
        en: "_preview"
    },
    rectName: {
        ja: "アートボード境界",
        en: "Artboard Bounds"
    },
    previewRect: {
        ja: "__プレビュー_アートボード境界",
        en: "__Preview_ArtboardBounds"
    },
    // --- Color & type options ---
    typeFill: {
        ja: "塗り",
        en: "Fill"
    },
    typeStroke: {
        ja: "線",
        en: "Stroke"
    },
    colorK100: {
        ja: "ブラック",
        en: "Black"
    },
    colorWhite: {
        ja: "ホワイト",
        en: "White"
    },
    colorSpecified: {
        ja: "HEX",
        en: "HEX"
    },
    colorCustomCMYK: {
        ja: "CMYK",
        en: "CMYK"
    },
    // --- Target options ---
    currentAB: {
        ja: "個別",
        en: "Create Individually"
    },
    allAB: {
        ja: "グループとして",
        en: "Create as Group"
    }
};


/* ===== Dialog appearance & position (tunable) ===== */
var DIALOG_OFFSET_X = 300; // shift right (+) / left (-)
var DIALOG_OFFSET_Y = 0; // shift down (+) / up (-)
var DIALOG_OPACITY = 0.98; // 0.0 - 1.0

/*
 * プレビュー遅延時間 / Preview delay timings
 * - 入力中は軽め（タイプしやすさ優先）/ Lighter while typing for responsiveness
 * - アウトライン計算時はやや重め / Slightly heavier when outlining text
 */
var PREVIEW_DELAY_TYPING_MS = 110; // recommend 100–120ms
var PREVIEW_DELAY_OUTLINE_MS = 240; // heavier when outlining text during preview

/* =========================================
 * DialogPersist util (extractable)
 * ダイアログの不透明度・初期位置・位置記憶を共通化するユーティリティ。 / Utility for dialog opacity, initial position, and remembering position.
 * 使い方 / Usage:
 *   DialogPersist.setOpacity(dlg, 0.95);
 *   DialogPersist.restorePosition(dlg, "__YourDialogKey", offsetX, offsetY);
 *   DialogPersist.rememberOnMove(dlg, "__YourDialogKey");
 *   DialogPersist.savePosition(dlg, "__YourDialogKey"); // 閉じる直前などに / e.g. just before closing
 * ========================================= */
(function(g) {
    if (!g.DialogPersist) {
        g.DialogPersist = {
            setOpacity: function(dlg, v) {
                try {
                    dlg.opacity = v;
                } catch (e) {}
            },
            _getSaved: function(key) {
                return g[key] && g[key].length === 2 ? g[key] : null;
            },
            _setSaved: function(key, loc) {
                g[key] = [loc[0], loc[1]];
            },
            _clampToScreen: function(loc) {
                try {
                    var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
                    var x = Math.max(vb[0] + 10, Math.min(loc[0], vb[2] - 10));
                    var y = Math.max(vb[1] + 10, Math.min(loc[1], vb[3] - 10));
                    return [x, y];
                } catch (e) {
                    return loc;
                }
            },
            restorePosition: function(dlg, key, offsetX, offsetY) {
                var loc = this._getSaved(key);
                try {
                    if (loc) {
                        dlg.location = this._clampToScreen(loc);
                    } else {
                        var l = dlg.location;
                        dlg.location = [l[0] + (offsetX | 0), l[1] + (offsetY | 0)];
                    }
                } catch (e) {}
            },
            rememberOnMove: function(dlg, key) {
                var self = this;
                dlg.onMove = function() {
                    try {
                        self._setSaved(key, [dlg.location[0], dlg.location[1]]);
                    } catch (e) {}
                };
            },
            savePosition: function(dlg, key) {
                try {
                    this._setSaved(key, [dlg.location[0], dlg.location[1]]);
                } catch (e) {}
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

// --- 色生成・パーサーヘルパー / Helpers: color constructors & parsers ---
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

// --- RGB/CMYK 変換ヘルパー / RGB/CMYK conversion helpers ---
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

// --- 塗り適用ヘルパー（責務分離）/ Fill helpers split by responsibility ---
/*
 * resolveFillColor(doc, mode, payload)
 * 目的: カラーオブジェクトの解決のみを担当（不透明度/線は扱わない）
 * Purpose: Resolve and return only the color object; do not touch opacity/stroke
 * - mode: ColorMode.K100 | ColorMode.WHITE | ColorMode.HEX | ColorMode.CMYK | ColorMode.NONE
 * - payload: { customValue: String, customCMYK: {c,m,y,k} }
 * 戻り値: RGBColor/CMYKColor または null（塗りなし）
 */
function resolveFillColor(doc, mode, payload) {
    try {
        if (mode === ColorMode.NONE) return null;
        if (mode === ColorMode.K100) {
            return createBlackColor(doc);
        }
        if (mode === ColorMode.WHITE) {
            if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
                return makeRGB(255, 255, 255);
            } else {
                return makeCMYK(0, 0, 0, 0);
            }
        }
        if (mode === ColorMode.HEX) {
            var col = parseCustomColor(doc, payload && payload.customValue);
            return col ? col : null;
        }
        if (mode === ColorMode.CMYK) {
            var c = payload && payload.customCMYK && payload.customCMYK.c,
                m = payload && payload.customCMYK && payload.customCMYK.m,
                y = payload && payload.customCMYK && payload.customCMYK.y,
                k = payload && payload.customCMYK && payload.customCMYK.k;
            var ok = (typeof c === 'number' && !isNaN(c)) &&
                (typeof m === 'number' && !isNaN(m)) &&
                (typeof y === 'number' && !isNaN(y)) &&
                (typeof k === 'number' && !isNaN(k));
            if (!ok) return null;
            if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
                var rgb = cmykToRgb(c, m, y, k);
                return makeRGB(rgb[0], rgb[1], rgb[2]);
            } else {
                return makeCMYK(c, m, y, k);
            }
        }
    } catch (e) {}
    return null;
}

/*
 * applyFill(rect, color, strokeOff)
 * 目的: 実際の塗り適用のみ（不透明度は今後も設定しない）
 * Purpose: Apply fill only; do not set opacity in any case
 * - strokeOff: true なら線をオフにする
 */
function applyFill(rect, color, strokeOff) {
    try {
        if (color) {
            rect.filled = true;
            rect.fillColor = color;
        } else {
            rect.filled = false;
        }
        if (strokeOff) {
            rect.stroked = false;
        }
        // 不透明度は設定しない / Do not touch opacity
    } catch (e) {}
}

/* applyLiveEffect: ライブエフェクトを適用（展開しない） */
function applyLiveEffect(item, effectName, dictData) {
    try {
        item.applyEffect(
            '<LiveEffect name="' + effectName + '">' +
            '<Dict data="' + dictData + '"/>' +
            '</LiveEffect>'
        );
    } catch (e) {
        alert("エフェクト適用エラー: " + e);
    }
}

function _toInt(x) {
    var n = parseFloat(x);
    return isNaN(n) ? NaN : n;
}

/*
 * customValue の解釈と RGBColor/CMYKColor へのマッピング / Parse customValue and map to RGBColor/CMYKColor
 * 受け入れ形式 / Accepts:
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

        // R,G,B (0-255, comma-separated; optional spaces, JP commas normalized above)
        var rgbCsv = s.match(/^\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*$/);
        if (rgbCsv) {
            var r0 = _clamp(parseInt(rgbCsv[1], 10), 0, 255);
            var g0 = _clamp(parseInt(rgbCsv[2], 10), 0, 255);
            var b0 = _clamp(parseInt(rgbCsv[3], 10), 0, 255);
            return makeRGB(r0, g0, b0);
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
 * 単位変換を一元化（Bleedプリセットなし）/ Centralize unit conversion (no Bleed preset)
 * - offsetText: current edit field text (string)
 * - unitCode: app.preferences.getIntegerPreference("rulerType")
 * Return: { pt: Number, displayText: String, disabled: Boolean }
 */
function resolveOffsetToPt(offsetText, unitCode) {
    var displayText = String(offsetText == null ? '' : offsetText);
    var n = parseFloat(displayText);
    if (isNaN(n)) n = 0;
    var pt = n * getPtFactorFromUnitCode(unitCode);
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

// ===== プレビュー用ヘルパー / Preview helpers =====


/* =========================================
 * PreviewHistory util (extractable)
 * ヒストリーを残さないプレビューのための小さなユーティリティ。/ Small utility for previews that do not leave history.
 * 他スクリプトでもこのブロックをコピペすれば再利用できます。/ You can reuse this block in other scripts.
 * 使い方 / Usage:
 *   PreviewHistory.start();     // ダイアログ表示時などにカウンタ初期化 / Initialize counter when dialog appears, etc.
 *   PreviewHistory.bump();      // プレビュー描画ごとにカウント(+1) / Increment for each preview rendering
 *   PreviewHistory.undo();      // 閉じる/キャンセル時に一括Undo / Undo all at close/cancel
 *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助 / Helper to cancel app.scheduleTask
 * ========================================= */
(function(g) {
    if (!g.PreviewHistory) {
        g.PreviewHistory = {
            start: function() {
                g.__previewUndoCount = 0;
            },
            bump: function() {
                g.__previewUndoCount = (g.__previewUndoCount | 0) + 1;
            },
            undo: function() {
                var n = g.__previewUndoCount | 0;
                try {
                    for (var i = 0; i < n; i++) app.executeMenuCommand('undo');
                } catch (e) {}
                g.__previewUndoCount = 0;
            },
            cancelTask: function(taskId) {
                try {
                    if (taskId) app.cancelTask(taskId);
                } catch (e) {}
            }
        };
    }
})($.global);


var __previewDebounceTask = null;

function schedulePreview(choice, delayMs) {
    try {
        if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
    } catch (e) {}
    var delay;
    if (typeof delayMs === 'number') {
        delay = delayMs;
    } else {
        delay = (choice && choice.usePreviewOutline) ? PREVIEW_DELAY_OUTLINE_MS : PREVIEW_DELAY_TYPING_MS;
    }
    $.global.__lastPreviewChoice = choice;
    var code = 'try{renderPreview(app.activeDocument, $.global.__lastPreviewChoice);}catch(e){}';
    try {
        __previewDebounceTask = app.scheduleTask(code, Math.max(0, delay | 0), false);
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
        layer.move(doc, ElementPlacement.PLACEATEND);
    } catch (e) {}
    return layer;
}

function getOrCreatePreviewRect(previewLayer, idx, top, left, width, height) {
    var nameBase = LABELS.previewRect[lang] + "#" + idx;
    // Remove existing item to avoid stacking Live Effects across previews
    try {
        for (var i = previewLayer.pathItems.length - 1; i >= 0; i--) {
            var it = previewLayer.pathItems[i];
            if (it.name === nameBase) {
                try {
                    it.remove();
                } catch (_) {}
                break;
            }
        }
    } catch (e) {}
    // Create a fresh rectangle each time
    var item = previewLayer.pathItems.rectangle(top, left, width, height);
    item.name = nameBase;
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

// Convert [L,T,R,B] bounds to a rectSpec {left, top, width, height}
function boundsToRectSpec(boundsArr) {
    if (!boundsArr || boundsArr.length !== 4) return null;
    var L = boundsArr[0],
        T = boundsArr[1],
        R = boundsArr[2],
        B = boundsArr[3];
    return {
        left: L,
        top: T,
        width: (R - L),
        height: (T - B)
    };
}

/**
 * 共通プレビュー矩形作成 / Build preview rectangle with fill, stroke, and live corner.
 * @param {Layer} previewLayer
 * @param {number} idx - index for naming (group=0, item=i)
 * @param {Object} rectSpec - {left, top, width, height}
 * @param {Object} choice - { offsetV, offsetH, colorMode, customValue, customCMYK, roundPt, isPill }
 * @param {Document} doc
 * @returns {PathItem|null}
 */
function buildPreviewRect(previewLayer, idx, rectSpec, choice, doc) {
    if (!rectSpec) return null;
    var L = rectSpec.left,
        T = rectSpec.top,
        w = rectSpec.width,
        h = rectSpec.height;
    var R = L + w,
        B = T - h;

    var oV = (choice && typeof choice.offsetV === 'number') ? choice.offsetV : 0;
    var oH = (choice && typeof choice.offsetH === 'number') ? choice.offsetH : 0;

    var rect = getOrCreatePreviewRect(previewLayer, idx, T + oV, L - oH, w + oH * 2, h + oV * 2);

    // Resolve once (use for fill or stroke)
    var col = resolveFillColor(doc, choice.colorMode, {
        customValue: choice.customValue,
        customCMYK: choice.customCMYK
    });

    if (choice && choice.type === 'stroke') {
        // Preview as stroke-only in chosen color
        try {
            rect.filled = false;
        } catch (e) {}
        try {
            rect.stroked = !!col;
            if (col) rect.strokeColor = col;
            rect.strokeWidth = (choice && typeof choice.strokeWidth === 'number' && choice.strokeWidth > 0) ? choice.strokeWidth : 1;
        } catch (e) {}
    } else {
        // Default: fill color + subtle gray outline
        applyFill(rect, col, true);
        try {
            rect.stroked = true;
            rect.strokeColor = getPreviewStrokeColor(doc);
            rect.strokeWidth = 0.5;
        } catch (e) {}
    }

    // Corner radius via Live Effect (kept unexpanded)
    try {
        var r = (choice && typeof choice.roundPt === 'number') ? choice.roundPt : 0;
        if (choice && choice.isPill) {
            r = (h + oV * 2) / 2; // pill: radius = height/2 with margins
        }
        if (r > 0) {
            applyLiveEffect(rect, "Adobe Round Corners", "R radius " + r + " ");
        }
    } catch (e) {}

    try {
        rect.selected = false;
    } catch (e) {}
    try {
        rect.zOrder(ZOrderMethod.SENDTOBACK);
    } catch (e) {}
    return rect;
}

// Small dispatcher: choose proper bounds getters based on outline/preview flags
function makeBoundsGetter(doc, useOutline, usePreview) {
    return {
        item: function(it) {
            return useOutline ? getFinalItemBounds(doc, it, usePreview) :
                getItemBounds(it, usePreview);
        },
        group: function(sel) {
            return useOutline ? getCombinedFinalBounds(doc, sel, usePreview) :
                getCombinedGeometricBounds(sel, usePreview);
        }
    };
}

/**
 * withTargetBounds: Iterate selection and provide rectSpec from the start.
 * - For group mode: calls groupFn(rectSpec)
 * - For individual mode: calls itemFn(item, rectSpec, index)
 */
function withTargetBounds(doc, sel, choice, groupFn, itemFn) {
    try {
        if (!doc || !choice) return;
        var useOutline = !!choice.usePreviewOutline;
        var usePreview = !!choice.usePreviewBounds;
        var G = makeBoundsGetter(doc, useOutline, usePreview);
        if (choice.target === 'group') {
            var gb = G.group(sel);
            var rs = boundsToRectSpec(gb);
            if (rs && typeof groupFn === 'function') groupFn(rs);
        } else {
            for (var i = 0; i < sel.length; i++) {
                var it = sel[i];
                if (!it) continue;
                var ib = G.item(it);
                var rs2 = boundsToRectSpec(ib);
                if (rs2 && typeof itemFn === 'function') itemFn(it, rs2, i);
            }
        }
    } catch (e) {}
}

/*
 * プレビュー描画（Previewレイヤーへ一時オブジェクトを生成）/ Draw preview shapes in the Preview layer (temporary objects)
 * Render live preview into the dedicated Preview layer.
 */
function renderPreview(doc, choice) {
    clearPreview(false);
    if (!doc || !choice) return;

    var prevCS = null;
    try {
        prevCS = app.coordinateSystem;
    } catch (e) {}
    try {
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
    } catch (e) {}

    var previewLayer = getOrCreatePreviewLayer(doc);
    var sel = [];
    try {
        sel = doc.selection || [];
    } catch (e) {
        sel = [];
    }

    try {
        withTargetBounds(doc, sel, choice,
            function(rectSpec) {
                buildPreviewRect(previewLayer, 0, rectSpec, choice, doc);
            },
            function(it, rectSpec, idx) {
                buildPreviewRect(previewLayer, idx, rectSpec, choice, doc);
            }
        );
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

    // ==== Top bar: Preset dropdown + Register button (logic to be added later) ====
    var topBar = dlg.add('group');
    topBar.orientation = 'row';
    topBar.alignChildren = ['center', 'center'];
    topBar.alignment = ['center', 'top'];
    topBar.spacing = 10;

    // Preset dropdown (placeholder items; populated later)
    var presetDropdown = topBar.add('dropdownlist', undefined, []);
    try {
        presetDropdown.preferredSize = {
            width: 200,
            height: -1
        };
    } catch (e) {}

    // Register button
    var btnRegister = topBar.add('button', undefined, '登録');
    try {
        btnRegister.preferredSize = {
            width: 60,
            height: 20
        };
    } catch (e) {}

    // ==== Preset helpers =====================================================
    var __presetFile = (function() {
        try {
            var folder = Folder.userData.fsName + "/ai-scripts";
            var f = new Folder(folder);
            if (!f.exists) f.create();
            return folder + "/DrawRectangleBehindSelectedObject_presets.json";
        } catch (e) {
            return Folder.myDocuments.fsName + "/DrawRectangleBehindSelectedObject_presets.json";
        }
    })();

    function __readPresets() {
        var list = [];
        try {
            var file = new File(__presetFile);
            if (file.exists) {
                if (file.open('r')) {
                    var txt = file.read();
                    file.close();
                    try {
                        list = JSON.parse(txt) || [];
                    } catch (__) {
                        list = [];
                    }
                }
            }
        } catch (e) {}
        if (!(list instanceof Array)) list = [];
        return list;
    }

    function __writePresets(list) {
        try {
            var file = new File(__presetFile);
            if (file.open('w')) {
                file.write(JSON.stringify(list, null, 2));
                file.close();
                return true;
            }
        } catch (e) {}
        return false;
    }

    function __refreshPresetDropdown(list, selectName) {
        try {
            while (presetDropdown.items.length) presetDropdown.remove(0);
            for (var i = 0; i < list.length; i++) {
                presetDropdown.add('item', list[i].name);
            }
            if (selectName) {
                for (var j = 0; j < presetDropdown.items.length; j++) {
                    if (presetDropdown.items[j].text === selectName) {
                        presetDropdown.selection = j;
                        break;
                    }
                }
            } else if (presetDropdown.items.length) {
                presetDropdown.selection = 0;
            }
        } catch (e) {}
    }

    function __promptPresetName(defaultName) {
        var w = new Window('dialog', 'プリセット名 / Preset Name');
        w.orientation = 'column';
        w.alignChildren = ['fill', 'center'];
        w.margins = 15;
        var g1 = w.add('group');
        g1.orientation = 'row';
        g1.alignChildren = ['left', 'center'];
        g1.add('statictext', undefined, '名称');
        var et = g1.add('edittext', undefined, defaultName || '');
        et.characters = 24;
        try {
            et.active = true;
        } catch (e) {}
        var g2 = w.add('group');
        g2.alignment = ['center', 'center'];
        var bOk = g2.add('button', undefined, 'OK');
        var bCancel = g2.add('button', undefined, 'キャンセル');
        var result = null;
        bOk.onClick = function() {
            result = String(et.text || '').replace(/^\s+|\s+$/g, '');
            w.close(1);
        };
        bCancel.onClick = function() {
            result = null;
            w.close(0);
        };
        var r = w.show();
        if (r == 1 && result) return result;
        return null;
    }

    var __presets = __readPresets();
    __refreshPresetDropdown(__presets);

    btnRegister.onClick = function() {
        try {
            var name = __promptPresetName('');
            if (!name) return; // cancelled
            var choiceNow = __collectChoice(true);
            // store only necessary fields (avoid circulars)
            var entry = {
                name: name,
                choice: serializeChoice(choiceNow)
            };
            // Upsert by name
            var replaced = false;
            for (var i = 0; i < __presets.length; i++) {
                if (__presets[i].name === name) {
                    __presets[i] = entry;
                    replaced = true;
                    break;
                }
            }
            if (!replaced) __presets.push(entry);
            __writePresets(__presets);
            __refreshPresetDropdown(__presets, name);
        } catch (e) {}
    };

    // Preset dropdown selection handler: restore settings from preset into UI
    presetDropdown.onChange = function() {
        try {
            var selIdx = presetDropdown.selection ? presetDropdown.selection.index : -1;
            if (selIdx < 0 || selIdx >= __presets.length) return;
            var chosen = __presets[selIdx];
            if (!chosen || !chosen.choice) return;
            applyChoiceToUI(chosen.choice);
        } catch (e) {}
    };

    // --- Two-column layout container ---
    var mainRow = dlg.add('group');
    mainRow.orientation = 'row';
    mainRow.alignChildren = 'top';
    mainRow.spacing = 16;

    var leftCol = mainRow.add('group');
    leftCol.orientation = 'column';
    leftCol.alignChildren = 'fill';
    leftCol.spacing = 12;

    var rightCol = mainRow.add('group');
    rightCol.orientation = 'column';
    rightCol.alignChildren = 'fill';
    rightCol.spacing = 12;

    // ==== Margin panel layout adjusted (referenced structure) ====
    // PANEL1 (offsetPanel)
    var offsetPanel = leftCol.add('panel', undefined, LABELS.offsetTitle[lang]);
    offsetPanel.orientation = 'row';
    offsetPanel.alignChildren = ['left', 'top'];
    offsetPanel.spacing = 10;
    offsetPanel.margins = [15, 15, 20, 15];

    // LEFTCOL container (row)
    var marginLeftCol = offsetPanel.add('group');
    marginLeftCol.orientation = 'row';
    marginLeftCol.alignChildren = ['left', 'center'];
    marginLeftCol.spacing = 10;
    marginLeftCol.margins = 0;

    // GROUP1 (column for two rows: 上下 / 左右)
    var marginGroup1 = marginLeftCol.add('group');
    marginGroup1.orientation = 'column';
    marginGroup1.alignChildren = ['left', 'center'];
    marginGroup1.spacing = 10;
    marginGroup1.margins = 0;

    // GROUP2 (row: 上下)
    var groupV = marginGroup1.add('group');
    groupV.orientation = 'row';
    groupV.alignChildren = ['left', 'center'];
    groupV.spacing = 10;
    groupV.margins = 0;

    var offsetVInputLabel = groupV.add('statictext', undefined, '上下');
    var offsetVInput = groupV.add('edittext', undefined, '2');
    offsetVInput.preferredSize = {
        width: 35,
        height: -1
    }; // widen a bit
    offsetVInput.characters = 3;
    changeValueByArrowKey(offsetVInput, function() {
        updatePreview();
    });
    offsetVInput.addEventListener('keydown', function(e) {
        if (e.keyName == 'Enter') {
            try {
                if (cbLinkMargins.value) {
                    offsetHInput.text = String(offsetVInput.text);
                }
            } catch (__) {}
            try {
                requestPreview(buildChoiceFromUI(), true);
            } catch (_) {}
        }
    });
    groupV.add('statictext', undefined, getCurrentUnitLabel());

    // GROUP3 (row: 左右)
    var groupH = marginGroup1.add('group');
    groupH.orientation = 'row';
    groupH.alignChildren = ['left', 'center'];
    groupH.spacing = 10;
    groupH.margins = 0;

    var offsetHInputLabel = groupH.add('statictext', undefined, '左右');
    var offsetHInput = groupH.add('edittext', undefined, '2');
    offsetHInput.preferredSize = {
        width: 35,
        height: -1
    }; // widen a bit
    offsetHInput.characters = 3;
    changeValueByArrowKey(offsetHInput, function() {
        updatePreview();
    });
    offsetHInput.addEventListener('keydown', function(e) {
        if (e.keyName == 'Enter') {
            try {
                if (cbLinkMargins.value) {
                    offsetVInput.text = String(offsetHInput.text);
                }
            } catch (__) {}
            try {
                requestPreview(buildChoiceFromUI(), true);
            } catch (_) {}
        }
    });
    groupH.add('statictext', undefined, getCurrentUnitLabel());

    // GROUP4 (right column in the same row: checkbox "連動")
    var groupLink = marginLeftCol.add('group');
    groupLink.orientation = 'row';
    groupLink.alignChildren = ['center', 'top'];
    groupLink.spacing = 10;
    groupLink.margins = 0;
    groupLink.alignment = ['left', 'center'];

    var cbLinkMargins = groupLink.add('checkbox', undefined, '連動');
    cbLinkMargins.value = true; // default ON

    function __updateLinkDim() {
        try {
            offsetHInput.enabled = !cbLinkMargins.value;
        } catch (e) {}
        try {
            offsetHInputLabel.enabled = !cbLinkMargins.value;
        } catch (e) {}
    }

    cbLinkMargins.onClick = function() {
        __updateLinkDim();
        try {
            if (cbLinkMargins.value) {
                offsetHInput.text = String(offsetVInput.text);
            }
        } catch (__) {}
        updatePreviewCommit();
    };
    cbLinkMargins.onChanging = cbLinkMargins.onClick;

    var roundPanel = leftCol.add('panel', undefined, LABELS.roundTitle[lang]);
    roundPanel.orientation = 'column';
    roundPanel.alignChildren = 'left';
    roundPanel.margins = [15, 20, 15, 10];

    var roundRow = roundPanel.add('group');
    roundRow.orientation = 'row';
    roundRow.alignChildren = ['left', 'center'];
    roundRow.alignment = ['center', 'top'];
    // Checkbox before the numeric field (UI only for now)
    var cbRoundEnable = roundRow.add('checkbox', undefined, '');
    var __lastRoundValue = '2'; // will be updated after roundInput is created
    cbRoundEnable.value = true; // default ON
    try {
        cbRoundEnable.alignment = ['left', 'center'];
    } catch (e) {}

    // Enable/disable for corner UI (checkbox governs round input & pill)
    function setCornerUIEnabled(on) {
        var v = !!on;
        try {
            roundInput.enabled = v && !(cbPill && cbPill.value);
        } catch (e) {}
        try {
            roundUnitLabel.enabled = v;
        } catch (e) {}
        try {
            cbPill.enabled = v;
        } catch (e) {}
    }
    cbRoundEnable.onClick = function() {
        if (cbRoundEnable.value) {
            // Restoring previous value
            try {
                if (__lastRoundValue) roundInput.text = __lastRoundValue;
            } catch (e) {}
        } else {
            // Saving current value before disabling
            try {
                __lastRoundValue = String(roundInput.text);
            } catch (e) {}
        }
        setCornerUIEnabled(cbRoundEnable.value);
        updatePreviewCommit();
    };
    cbRoundEnable.onChanging = cbRoundEnable.onClick;

    try {
        roundInputLabel.preferredSize.width = __LABEL_WIDTH;
        roundInputLabel.alignment = ['right', 'center'];
    } catch (e) {}

    // --- Helper: when Pill is ON, show height/2 (with margins) in roundInput (read-only)
    function updatePillRoundField() {
        if (cbRoundEnable && !cbRoundEnable.value) return; // disabled: do not compute or display
        try {
            if (!cbPill || !cbPill.value) return; // only when pill mode
            var doc = app.activeDocument;
            if (!doc) return;
            var sel = doc.selection || [];
            if (!sel.length) return;

            var usePreview = true; // always ON (UI label only for now)
            var unitCode = getCurrentUnitCode();
            var oVpt = resolveOffsetToPt(offsetVInput.text, unitCode).pt;

            // Determine bounds according to target mode
            var gb = null;
            if (allRadio.value) {
                gb = getCombinedFinalBounds(doc, sel, usePreview);
            } else {
                gb = getFinalItemBounds(doc, sel[0], usePreview);
            }
            if (!gb) return;

            var left = gb[0],
                top = gb[1],
                right = gb[2],
                bottom = gb[3];
            var h = top - bottom;

            var pillRadiusPt = (h + oVpt * 2) / 2; // height/2 including margins
            var factor = getPtFactorFromUnitCode(unitCode);
            var displayVal = Math.round((pillRadiusPt / factor) * 100) / 100; // 2桁表示

            roundInput.text = String(displayVal);
            try {
                roundInput.enabled = false;
            } catch (e) {}
            // 角丸欄を更新したら、PillがONのとき左右マージンにもコピー
            try {
                if (cbPill && cbPill.value) {
                    offsetHInput.text = String(roundInput.text);
                }
            } catch (__) {}
        } catch (e) {}
    }

    var roundInput = roundRow.add('edittext', undefined, '2');
    __lastRoundValue = String(roundInput.text);
    roundInput.preferredSize = {
        width: 35,
        height: -1
    }; // widen a bit
    roundInput.characters = 3;
    changeValueByArrowKey(roundInput, function() {
        updatePreview(); // UIのみ。描画ロジックは後日実装
    });

    roundInput.addEventListener('keydown', function(e) {
        if (e.keyName == 'Enter') {
            try {
                requestPreview(buildChoiceFromUI(), true);
            } catch (_) {}
        }
    });

    var unitLabel2 = getCurrentUnitLabel();
    var roundUnitLabel = roundRow.add('statictext', undefined, unitLabel2);

    // --- Pill shape option: move pill checkbox to same row as input/unit ---
    var cbPill = roundRow.add('checkbox', undefined, 'ピル形状');
    // Vertically center controls in the round row
    try {
        roundRow.alignment = ['fill', 'center'];
    } catch (e) {}
    try {
        roundInput.alignment = ['left', 'center'];
    } catch (e) {}
    try {
        roundUnitLabel.alignment = ['left', 'center'];
    } catch (e) {}
    try {
        cbPill.alignment = ['left', 'center'];
    } catch (e) {}
    cbPill.value = false; // default OFF
    cbPill.onClick = function() {
        try {
            roundInput.enabled = !cbPill.value;
        } catch (e) {}
        if (cbPill.value) {
            // 1) 計算して角丸欄へ反映
            try {
                updatePillRoundField();
            } catch (__) {}
            // 2) 連動を自動的にOFF + ディム更新
            try {
                cbLinkMargins.value = false;
            } catch (__) {}
            try {
                if (typeof __updateLinkDim === 'function') __updateLinkDim();
            } catch (__) {}
            // 3) 角丸の値を左右マージンに入れる（単位表示と一致）
            try {
                offsetHInput.text = String(roundInput.text);
            } catch (__) {}
        }
        updatePreviewCommit();
    };
    cbPill.onChanging = cbPill.onClick;


    dlg.onShow = function() {
        try {
            if (typeof setStrokeWidthUIEnabled === 'function') setStrokeWidthUIEnabled(!!typeStrokeRadio.value);
        } catch (e) {}
        DialogPersist.restorePosition(dlg, __DLG_KEY, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        try {
            offsetVInput.active = true;
        } catch (e) {}
        PreviewHistory.start();
        try {
            cbGroupWithText.value = true;
        } catch (e) {}
        try {
            if (typeof __updateLinkDim === 'function') __updateLinkDim();
        } catch (e) {}
        try {
            if (cbPill && cbPill.value) {
                roundInput.enabled = false;
                updatePillRoundField();
            }
        } catch (e) {}
        try {
            if (typeof setCornerUIEnabled === 'function') setCornerUIEnabled(!!cbRoundEnable.value);
        } catch (e) {}
        updatePreviewCommit();
    };
    DialogPersist.rememberOnMove(dlg, __DLG_KEY);

    // Add new panel for color
    var colorPanel = rightCol.add('panel', undefined, LABELS.colorTitle[lang]);
    colorPanel.orientation = 'column';
    colorPanel.alignChildren = 'left';
    colorPanel.margins = [15, 20, 15, 10];
    colorPanel.spacing = 10; // increase vertical gap between rows

    var k100Radio = colorPanel.add('radiobutton', undefined, LABELS.colorK100[lang]);
    var whiteRadio = colorPanel.add('radiobutton', undefined, LABELS.colorWhite[lang]);

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

    // If the field shows exactly "0", typing a digit should replace it (avoid "03")
    function replaceZeroOnFirstDigit(et) {
        try {
            et.addEventListener('keydown', function(ev) {
                var k = String(ev.keyName || '');
                if (/^[0-9]$/.test(k)) {
                    try {
                        var t = String(et.text || '');
                        if (t === '0') {
                            et.text = k; // replace instead of append
                            if (ev && typeof ev.preventDefault === 'function') ev.preventDefault();
                            validateCmykField(et);
                            updatePreviewTyping();
                        }
                    } catch (e) {}
                }
            });
        } catch (e) {}
    }

    // Bind common handlers to a CMYK EditText
    function bindCmykField(et) {
        et.onChanging = function() {
            // Drop a single leading zero when typing the first digit (avoid "03")
            try {
                var t = String(et.text || '');
                var m = t.match(/^0([0-9])$/);
                if (m) et.text = m[1];
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
    var __hotkeyGuard = {
        active: false
    };

    function attachTypingBlockOnFocusBlur(ctrl) {
        try {
            ctrl.addEventListener('focus', function() {
                __hotkeyGuard.active = true;
            });
        } catch (e) {}
        try {
            ctrl.addEventListener('blur', function() {
                __hotkeyGuard.active = false;
            });
        } catch (e) {}
    }
    // Block hotkeys while typing in these fields
    attachTypingBlockOnFocusBlur(offsetVInput);
    attachTypingBlockOnFocusBlur(offsetHInput);
    attachTypingBlockOnFocusBlur(customInput);
    attachTypingBlockOnFocusBlur(etC);
    attachTypingBlockOnFocusBlur(etM);
    attachTypingBlockOnFocusBlur(etY);
    attachTypingBlockOnFocusBlur(etK);


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
    whiteRadio.value = false;

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

    // --- Centralized color-mode selector (K100 / WHITE / HEX / CMYK) ---
    function selectColorMode(mode) {
        // 1) Toggle radios
        k100Radio.value = (mode === ColorMode.K100);
        if (typeof whiteRadio !== 'undefined') whiteRadio.value = (mode === ColorMode.WHITE);
        specifiedRadio.value = (mode === ColorMode.HEX);
        cmykRadio.value = (mode === ColorMode.CMYK);

        // 2) Enable/disable inputs
        updateColorEnableFromRadios();

        // 3) Field highlights
        setEditHighlight(customInput, mode === ColorMode.HEX);
        setEditHighlight(etC, mode === ColorMode.CMYK);

        // 4) Defaults / focus per mode
        try {
            if (mode === ColorMode.HEX) {
                var t = String(customInput.text || '').trim();
                if (t === '' || t === '#') customInput.text = '#ffcc00';
                customInput.active = true;
            } else if (mode === ColorMode.CMYK) {
                etC.active = true;
            }
        } catch (e) {}

        // 5) Commit preview
        updatePreviewCommit();
    }

    // Add new panel for target (moved back to left column)
    var targetPanel = leftCol.add('panel', undefined, LABELS.targetTitle[lang]);
    targetPanel.orientation = 'row';
    targetPanel.alignChildren = ['left', 'center'];
    targetPanel.margins = [15, 20, 15, 10];
    targetPanel.spacing = 20;

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

    // Shared collector to build a choice object; `finalMode` toggles any final-only tweaks.
    function __collectChoice(finalMode) {
        var colorMode = (function() {
            if (k100Radio.value) return ColorMode.K100;
            if (whiteRadio.value) return ColorMode.WHITE;
            if (specifiedRadio.value) return ColorMode.HEX;
            if (cmykRadio.value) return ColorMode.CMYK;
            return ColorMode.K100;
        })();
        var unitCode = getCurrentUnitCode();
        var resolvedV = resolveOffsetToPt(offsetVInput.text, unitCode);
        var resolvedH = resolveOffsetToPt(offsetHInput.text, unitCode);
        var offsetVPt = resolvedV.pt;
        var offsetHPt = resolvedH.pt;
        try {
            if (cbLinkMargins && cbLinkMargins.value) offsetHPt = offsetVPt;
        } catch (e) {}
        var resolvedRound = resolveOffsetToPt(roundInput.text, unitCode);
        var roundPt = Math.max(0, resolvedRound.pt);
        var __roundEnabled = true;
        try {
            __roundEnabled = !!cbRoundEnable.value;
        } catch (e) {
            __roundEnabled = true;
        }
        if (!__roundEnabled) roundPt = 0;
        var target = currentRadio.value ? 'individual' : (allRadio.value ? 'group' : 'individual');
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
        var choice = {
            colorMode: colorMode,
            customValue: customValue,
            customCMYK: cmykObj,
            offsetV: offsetVPt,
            offsetH: offsetHPt,
            roundPt: roundPt,
            isPill: (__roundEnabled ? !!cbPill.value : false),
            type: (function() {
                try {
                    return typeStrokeRadio.value ? 'stroke' : 'fill';
                } catch (e) {
                    return 'fill';
                }
            })(),
            strokeWidth: (function() {
                try {
                    var uc = getCurrentUnitCode();
                    return resolveOffsetToPt(strokeWidthInput.text, uc).pt;
                } catch (e) {
                    return 1;
                }
            })(),
            target: target,
            groupWithText: (function() {
                try {
                    return !!cbGroupWithText.value;
                } catch (e) {
                    return true;
                }
            })(),
            usePreviewBounds: true, // always on
            usePreviewOutline: true
        };
        return choice;
    }

    // ---- Preset schema helpers (centralize what to save & how to apply) ----
    function serializeChoice(choice) {
        // Keep only stable, serializable fields
        return {
            colorMode: choice.colorMode,
            customValue: choice.customValue,
            customCMYK: choice.customCMYK,
            offsetV: choice.offsetV,
            offsetH: choice.offsetH,
            roundPt: choice.roundPt,
            isPill: !!choice.isPill,
            type: choice.type,
            strokeWidth: choice.strokeWidth,
            target: choice.target,
            groupWithText: !!choice.groupWithText
        };
    }

    function applyChoiceToUI(c) {
        if (!c) return;
        // Color mode
        if (c.colorMode === ColorMode.K100) k100Radio.notify('onClick');
        else if (c.colorMode === ColorMode.WHITE) whiteRadio.notify('onClick');
        else if (c.colorMode === ColorMode.HEX) {
            specifiedRadio.notify('onClick');
            try {
                customInput.text = c.customValue || '#';
            } catch (e) {}
        } else if (c.colorMode === ColorMode.CMYK) {
            cmykRadio.notify('onClick');
            try {
                etC.text = c.customCMYK.c;
                etM.text = c.customCMYK.m;
                etY.text = c.customCMYK.y;
                etK.text = c.customCMYK.k;
            } catch (e) {}
        }

        // Margins (pt → current unit display)
        try {
            offsetVInput.text = String(Math.round(c.offsetV / getPtFactorFromUnitCode(getCurrentUnitCode())));
        } catch (e) {}
        try {
            offsetHInput.text = String(Math.round(c.offsetH / getPtFactorFromUnitCode(getCurrentUnitCode())));
        } catch (e) {}

        // Corner radius & pill
        try {
            cbRoundEnable.value = (c.roundPt > 0 || c.isPill);
            roundInput.text = String(Math.round(c.roundPt / getPtFactorFromUnitCode(getCurrentUnitCode())));
            cbPill.value = !!c.isPill;
            setCornerUIEnabled(cbRoundEnable.value);
        } catch (e) {}

        // Type (fill/stroke)
        try {
            if (c.type === 'stroke') typeStrokeRadio.notify('onClick');
            else typeFillRadio.notify('onClick');
        } catch (e) {}

        // Stroke width
        try {
            strokeWidthInput.text = String(Math.round((c.strokeWidth || 1) / getPtFactorFromUnitCode(getCurrentUnitCode())));
        } catch (e) {}

        // Target
        try {
            if (c.target === 'group') allRadio.notify('onClick');
            else currentRadio.notify('onClick');
        } catch (e) {}

        // Group with text
        try {
            cbGroupWithText.value = !!c.groupWithText;
        } catch (e) {}

        // Refresh preview
        updatePreviewCommit();
    }

    function buildChoiceFromUI() {
        return __collectChoice(false);
    }

    function updatePreviewTyping() {
        try {
            updatePillRoundField();
        } catch (e) {}
        try {
            schedulePreview(buildChoiceFromUI());
        } catch (e) {}
    }

    function updatePreviewCommit() {
        try {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
        } catch (_) {}
        try {
            updatePillRoundField();
        } catch (e) {}
        try {
            renderPreview(app.activeDocument, buildChoiceFromUI());
        } catch (_) {}
    }

    function updatePreview() {
        try {
            requestPreview(buildChoiceFromUI(), false);
        } catch (e) {}
    }

    offsetVInput.onChanging = function() {
        try {
            if (cbLinkMargins.value) {
                offsetHInput.text = String(offsetVInput.text);
            }
        } catch (e) {}
        updatePreviewTyping();
    };
    offsetVInput.onChange = function() {
        try {
            if (cbLinkMargins.value) {
                offsetHInput.text = String(offsetVInput.text);
            }
        } catch (e) {}
        updatePreviewCommit();
    };
    offsetHInput.onChanging = function() {
        try {
            if (cbLinkMargins.value) {
                offsetVInput.text = String(offsetHInput.text);
            }
        } catch (e) {}
        updatePreviewTyping();
    };
    offsetHInput.onChange = function() {
        try {
            if (cbLinkMargins.value) {
                offsetVInput.text = String(offsetHInput.text);
            }
        } catch (e) {}
        updatePreviewCommit();
    };

    k100Radio.onClick = function() {
        selectColorMode(ColorMode.K100);
    };
    k100Radio.onChanging = function() {
        selectColorMode(ColorMode.K100);
    };

    whiteRadio.onClick = function() {
        selectColorMode(ColorMode.WHITE);
    };
    whiteRadio.onChanging = whiteRadio.onClick;

    specifiedRadio.onClick = function() {
        selectColorMode(ColorMode.HEX);
    };
    specifiedRadio.onChanging = function() {
        selectColorMode(ColorMode.HEX);
    };

    cmykRadio.onClick = function() {
        selectColorMode(ColorMode.CMYK);
    };
    cmykRadio.onChanging = function() {
        selectColorMode(ColorMode.CMYK);
    };

    // --- Hotkeys: N/K/H/C to switch color mode radios ---
    function addColorHotkeys(dialog) {
        dialog.addEventListener('keydown', function(event) {
            if (__hotkeyGuard.active) return; // typing in a field
            var key = (event && event.keyName) ? String(event.keyName).toUpperCase() : '';
            if (key === 'K') {
                selectColorMode(ColorMode.K100);
                event.preventDefault();
            } else if (key === 'W') {
                selectColorMode(ColorMode.WHITE);
                event.preventDefault();
            } else if (key === 'H') {
                selectColorMode(ColorMode.HEX);
                event.preventDefault();
            } else if (key === 'C') {
                selectColorMode(ColorMode.CMYK);
                event.preventDefault();
            }
        });
    }
    addColorHotkeys(dlg);

    // Add new panel for type (under color panel)
    var typePanel = rightCol.add('panel', undefined, LABELS.typeTitle[lang]);
    typePanel.orientation = 'row';
    typePanel.alignChildren = ['left', 'center'];
    typePanel.margins = [15, 20, 15, 10];
    typePanel.spacing = 20;

    var typeFillRadio = typePanel.add('radiobutton', undefined, LABELS.typeFill[lang]);
    var typeStrokeRadio = typePanel.add('radiobutton', undefined, LABELS.typeStroke[lang]);

    // Stroke width row (under type radios)
    var strokeWidthRow = typePanel.add('group');
    strokeWidthRow.orientation = 'row';
    strokeWidthRow.alignChildren = ['left', 'center'];
    strokeWidthRow.spacing = 6;

    var strokeWidthLabel = strokeWidthRow.add('statictext', undefined, LABELS.strokeWidth[lang]);
    var strokeWidthInput = strokeWidthRow.add('edittext', undefined, '1');
    strokeWidthInput.preferredSize = {
        width: 35,
        height: -1
    };
    strokeWidthInput.characters = 3;
    var strokeWidthUnit = strokeWidthRow.add('statictext', undefined, getCurrentUnitLabel());

    // Enable/disable stroke width depending on type selection
    function setStrokeWidthUIEnabled(on) {
        var v = !!on;
        try {
            strokeWidthInput.enabled = v;
        } catch (e) {}
        try {
            strokeWidthLabel.enabled = v;
        } catch (e) {}
        try {
            strokeWidthUnit.enabled = v;
        } catch (e) {}
    }

    function updateTypeEnableAndPreview() {
        setStrokeWidthUIEnabled(!!typeStrokeRadio.value);
        updatePreviewCommit();
    }

    // Radios toggle enable + refresh preview
    typeFillRadio.onClick = updateTypeEnableAndPreview;
    typeFillRadio.onChanging = updateTypeEnableAndPreview;
    typeStrokeRadio.onClick = updateTypeEnableAndPreview;
    typeStrokeRadio.onChanging = updateTypeEnableAndPreview;

    // Stroke width field handlers
    changeValueByArrowKey(strokeWidthInput, function() {
        updatePreviewTyping();
    });
    strokeWidthInput.onChanging = function() {
        updatePreviewTyping();
    };
    strokeWidthInput.onChange = function() {
        updatePreviewCommit();
    };

    // Default selection: Fill
    typeFillRadio.value = true;
    typeStrokeRadio.value = false;

    // --- Hotkeys: Target (S/A) and Z-order (F/B/L) ---
    /* 対象スコープ: I/G（個別/グループ）、S/A（レガシー対応）ホットキー */
    function addScopeAndZHotkeys(dialog) {
        dialog.addEventListener('keydown', function(event) {
            if (__hotkeyGuard.active) return; // ignore when typing in fields
            var key = (event && event.keyName) ? String(event.keyName).toUpperCase() : '';

            // Target scope: I = 個別 (Individual), G = グループ (Group)
            // Legacy: S = current (Single), A = All
            if (key === 'I' || key === 'S') {
                try {
                    currentRadio.notify('onClick');
                } catch (_) {}
                event.preventDefault();
                return;
            }
            if (key === 'G' || key === 'A') {
                try {
                    allRadio.notify('onClick');
                } catch (_) {}
                event.preventDefault();
                return;
            }
        });
    }
    addScopeAndZHotkeys(dlg);



    currentRadio.onClick = updatePreviewCommit;
    allRadio.onClick = updatePreviewCommit;
    currentRadio.onChanging = updatePreviewCommit;
    allRadio.onChanging = updatePreviewCommit;

    // Grouping option (formerly preview bounds)
    var groupRow = leftCol.add('group');
    groupRow.orientation = 'column';
    groupRow.alignment = 'center';
    groupRow.alignChildren = ['left', 'center'];

    var cbGroupWithText = groupRow.add('checkbox', undefined, LABELS.previewBounds[lang]);
    cbGroupWithText.value = true; // default ON
    cbGroupWithText.onClick = updatePreviewCommit;
    cbGroupWithText.onChanging = updatePreviewCommit;


    // (UI removed) Always compute preview with outlined text for accuracy

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
        // Delete temporary outline layer after dialog closes
        try {
            var doc = app.activeDocument;
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                if (doc.layers[i].name === '__tmp_outline_bounds__') {
                    try {
                        doc.layers[i].remove();
                    } catch (e) {}
                }
            }
        } catch (e) {}
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
    return __collectChoice(true);
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

// --- Helper: Get or create a temporary outline layer (visible & editable for correct bounds) ---
function getOrCreateTempOutlineLayer(doc) {
    var name = '__tmp_outline_bounds__';
    var layer = null;
    try {
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
        // IMPORTANT: keep visible & editable during outline so bounds are valid
        layer.visible = true;
        layer.locked = false;
        try {
            layer.printable = false;
        } catch (e) {}
        try {
            layer.move(doc, ElementPlacement.PLACEATEND);
        } catch (e) {}
    } catch (e) {}
    return layer;
}


function getItemBounds(it, usePreview) {
    try {
        if (!it) return null;
        if (usePreview && it.visibleBounds) return it.visibleBounds; // [L,T,R,B]
        return it.geometricBounds; // fallback
    } catch (e) {
        return null;
    }
}

// --- Helper: Get bounds for a TextFrame by outlining (for final drawing) ---
function getTextOutlinedBounds(doc, tf, usePreview) {
    try {
        if (!tf || tf.typename !== 'TextFrame') return null;
        var tmp = getOrCreateTempOutlineLayer(doc);
        ensureLayerEditable(doc, tmp); // make active/editable

        // Duplicate the text into temp layer (must be visible)
        var dup = tf.duplicate(tmp, ElementPlacement.PLACEATBEGINNING);
        try {
            dup.hidden = false;
        } catch (e) {}
        try {
            dup.locked = false;
        } catch (e) {}
        var outlined = null;
        try {
            // Preferred API
            outlined = dup.createOutline(); // GroupItem
        } catch (e) {
            outlined = null;
        }

        // Fallback via menu command if direct API failed (older/edge environments)
        if (!outlined) {
            var prevSel = [];
            try {
                if (doc.selection && doc.selection.length) {
                    for (var i = 0; i < doc.selection.length; i++) prevSel.push(doc.selection[i]);
                }
            } catch (e) {}
            try {
                app.executeMenuCommand('deselectall');
            } catch (e) {}
            try {
                dup.selected = true;
            } catch (e) {}
            try {
                app.executeMenuCommand('createOutlines');
            } catch (e) {}
            try {
                outlined = (doc.selection && doc.selection.length) ? doc.selection[0] : null;
            } catch (e) {
                outlined = null;
            }
            // restore selection
            try {
                app.executeMenuCommand('deselectall');
            } catch (e) {}
            try {
                for (var j = 0; j < prevSel.length; j++) prevSel[j].selected = true;
            } catch (e) {}
        }

        var gb = null;
        if (outlined) {
            try {
                gb = usePreview && outlined.visibleBounds ? outlined.visibleBounds : outlined.geometricBounds;
            } catch (e) {
                gb = null;
            }
        }

        // cleanup
        try {
            if (outlined) outlined.remove();
        } catch (e) {}
        try {
            if (dup) dup.remove();
        } catch (e) {}
        return gb || null;
    } catch (e) {
        return null;
    }
}

// --- Helper: Get final bounds for any item (uses outlined bounds for text) ---
function getFinalItemBounds(doc, it, usePreview) {
    try {
        if (!it) return null;
        if (it.typename === 'TextFrame') {
            var gbText = getTextOutlinedBounds(doc, it, usePreview);
            if (gbText) return gbText;
        }
        return getItemBounds(it, usePreview);
    } catch (e) {
        return null;
    }
}


function getCombinedGeometricBounds(sel, usePreview) {
    try {
        if (!sel || !sel.length) return null;
        var left = null,
            top = null,
            right = null,
            bottom = null;
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            var gb = getItemBounds(it, usePreview); // [L,T,R,B]
            if (!gb) continue;
            if (left === null || gb[0] < left) left = gb[0];
            if (top === null || gb[1] > top) top = gb[1];
            if (right === null || gb[2] > right) right = gb[2];
            if (bottom === null || gb[3] < bottom) bottom = gb[3];
        }
        if (left === null) return null;
        return [left, top, right, bottom];
    } catch (e) {
        return null;
    }
}

// --- Helper: Combine final bounds with outlined text support (for group mode, final drawing) ---
function getCombinedFinalBounds(doc, sel, usePreview) {
    try {
        if (!sel || !sel.length) return null;
        var left = null,
            top = null,
            right = null,
            bottom = null;
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            var gb = getFinalItemBounds(doc, it, usePreview);
            if (!gb) continue;
            if (left === null || gb[0] < left) left = gb[0];
            if (top === null || gb[1] > top) top = gb[1];
            if (right === null || gb[2] > right) right = gb[2];
            if (bottom === null || gb[3] < bottom) bottom = gb[3];
        }
        if (left === null) return null;
        return [left, top, right, bottom];
    } catch (e) {
        return null;
    }
}

// --- Helper: Ensure a layer is editable and activate it ---
function ensureLayerEditable(doc, layer) {
    if (!doc || !layer) return;
    try {
        layer.visible = true;
    } catch (e) {}
    try {
        layer.locked = false;
    } catch (e) {}
    // Unlock ancestors if any
    try {
        var p = layer.parent;
        while (p && p.typename === 'Layer') {
            try {
                p.visible = true;
            } catch (e) {}
            try {
                p.locked = false;
            } catch (e) {}
            p = p.parent;
        }
    } catch (e) {}
    // Make it the active layer (some environments require this for insertion)
    try {
        doc.activeLayer = layer;
    } catch (e) {}
}

function drawRectangleForItem(doc, itemOrRect, choice) {
    // Accept either a legacy item with .geometricBounds or a rectSpec {left, top, width, height}
    var left, top, right, bottom, w, h, targetLayer = null;
    if (itemOrRect && typeof itemOrRect.left === 'number' && typeof itemOrRect.top === 'number') {
        // New rectSpec path
        left = itemOrRect.left;
        top = itemOrRect.top;
        w = itemOrRect.width;
        h = itemOrRect.height;
        right = left + w;
        bottom = top - h;
        try {
            targetLayer = itemOrRect.layer || null;
        } catch (e) {
            targetLayer = null;
        }
    } else {
        // Legacy path: expects .geometricBounds on the object
        var gb = itemOrRect.geometricBounds; // [left, top, right, bottom]
        left = gb[0];
        top = gb[1];
        right = gb[2];
        bottom = gb[3];
        w = right - left;
        h = top - bottom;
        try {
            targetLayer = itemOrRect.layer || null;
        } catch (e) {
            targetLayer = null;
        }
    }
    var oV = (choice && typeof choice.offsetV === 'number') ? choice.offsetV : 0;
    var oH = (choice && typeof choice.offsetH === 'number') ? choice.offsetH : 0;
    // ターゲットレイヤーは選択オブジェクトの属するレイヤー（なければアクティブレイヤー）
    if (!targetLayer) {
        try {
            targetLayer = doc.activeLayer;
        } catch (_) {}
    }
    ensureLayerEditable(doc, targetLayer);

    var rect = null;
    try {
        rect = targetLayer.pathItems.rectangle(
            top + oV,
            left - oH,
            w + oH * 2,
            h + oV * 2
        );
    } catch (e) {
        // Fallback: create on activeLayer, then move into targetLayer
        try {
            var tmpLayer = doc.activeLayer;
            rect = tmpLayer.pathItems.rectangle(
                top + oV,
                left - oH,
                w + oH * 2,
                h + oV * 2
            );
            try {
                rect.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (_) {}
        } catch (_) {
            throw e; // rethrow original if fallback also fails
        }
    }
    try {
        rect.selected = true;
    } catch (e) {}
    try {
        app.executeMenuCommand('Convert to Shape');
    } catch (e) {}

    var __finalColor = resolveFillColor(doc, choice.colorMode, {
        customValue: choice.customValue,
        customCMYK: choice.customCMYK
    });
    if (choice && choice.type === 'stroke') {
        try {
            rect.filled = false;
            rect.stroked = !!__finalColor;
            if (__finalColor) rect.strokeColor = __finalColor;
            rect.strokeWidth = (choice && typeof choice.strokeWidth === 'number' && choice.strokeWidth > 0) ? choice.strokeWidth : 1;
        } catch (e) {}
    } else {
        applyFill(rect, __finalColor, true); // filled; stroke off
    }
    // 角丸をライブエフェクトで適用（非展開）
    try {
        var r = (choice && typeof choice.roundPt === 'number') ? choice.roundPt : 0;
        if (choice && choice.isPill) {
            r = (h + oV * 2) / 2; // pill: 高さの1/2（上下マージン込み）
        }
        if (r > 0) {
            applyLiveEffect(rect, "Adobe Round Corners", "R radius " + r + " ");
        }
    } catch (e) {}

    rect.name = LABELS.rectName[lang];
    rect.selected = true;
    try {
        rect.hidden = false;
    } catch (e) {}
    try {
        targetLayer.visible = true;
        targetLayer.locked = false;
    } catch (e) {}

    rect.zOrder(ZOrderMethod.SENDTOBACK); // 常に最背面
    return rect;
}

function main() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert(LABELS.alertNoSelection[lang]);
        return;
    }

    var choice = showDialog();
    if (choice === null) return;

    // Save current selection BEFORE deselecting
    var sel = [];
    try {
        if (doc.selection && doc.selection.length) {
            // Robust copy for Illustrator's array-like Selection
            for (var i = 0; i < doc.selection.length; i++) sel.push(doc.selection[i]);
        }
    } catch (e) {}

    var __prevCS = null;
    try {
        __prevCS = app.coordinateSystem;
    } catch (e) {}
    try {
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
    } catch (e) {}

    // 選択解除（描画中の誤操作を避ける）
    app.executeMenuCommand('deselectall');

    withTargetBounds(doc, sel, choice,
        function(rectSpec) {
            // Group mode: use representative layer (first selection's layer or active)
            var repLayer = null;
            try {
                repLayer = (sel && sel.length && sel[0] && sel[0].layer) ? sel[0].layer : doc.activeLayer;
            } catch (__) {
                repLayer = doc.activeLayer;
            }
            rectSpec.layer = repLayer;
            var rectDrawn = drawRectangleForItem(doc, rectSpec, choice);
            try {
                if (choice.groupWithText && rectDrawn) {
                    var grp = repLayer.groupItems.add();
                    for (var gi = 0; gi < sel.length; gi++) {
                        try {
                            sel[gi].move(grp, ElementPlacement.PLACEATEND);
                        } catch (__) {}
                    }
                    try {
                        rectDrawn.move(grp, ElementPlacement.PLACEATBEGINNING);
                    } catch (__) {}
                    try {
                        rectDrawn.zOrder(ZOrderMethod.SENDTOBACK);
                    } catch (__) {}
                }
            } catch (__) {}
        },
        function(it, rectSpec, idx) {
            rectSpec.layer = (function() {
                try {
                    return it.layer;
                } catch (__) {
                    return null;
                }
            })();
            var rectDrawn2 = drawRectangleForItem(doc, rectSpec, choice);
            try {
                if (choice.groupWithText && rectDrawn2) {
                    var tgtLayer = null;
                    try {
                        tgtLayer = it.layer || rectDrawn2.layer;
                    } catch (__) {
                        tgtLayer = rectDrawn2.layer;
                    }
                    ensureLayerEditable(doc, tgtLayer);
                    var gi2 = tgtLayer.groupItems.add();
                    try {
                        it.move(gi2, ElementPlacement.PLACEATEND);
                    } catch (__) {}
                    try {
                        rectDrawn2.move(gi2, ElementPlacement.PLACEATBEGINNING);
                    } catch (__) {}
                    try {
                        rectDrawn2.zOrder(ZOrderMethod.SENDTOBACK);
                    } catch (__) {}
                }
            } catch (__) {}
        }
    );

    // Remove temporary outline layer if exists
    try {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name === '__tmp_outline_bounds__') {
                try {
                    doc.layers[i].remove();
                } catch (e) {}
            }
        }
    } catch (e) {}

    try {
        if (__prevCS !== null) app.coordinateSystem = __prevCS;
    } catch (e) {}
}
main();