#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*
============================================================
スクリプト名 / Script Name
------------------------------------------------------------
背面に長方形を作成 / Draw Rectangle Behind Selection

説明 / Description
------------------------------------------------------------
選択中のオブジェクト（または「グループとして」選択全体）に対して、外接バウンディングボックスを基準に
オフセットを加えた長方形を選択オブジェクトの属するレイヤーに作成します。作成された長方形は常に最背面に配置されます。
ライブプレビューにより、数値入力の変更が即座に確認できます。

● 主な機能 / Key Features
- オフセット（現在の定規単位に追従） / Offset respects current ruler units
- 角丸（UIのみ・ロジック追加予定） / Corner radius (UI only for now; logic TBD)
- カラー：K100 / ホワイト / HEX / CMYK / Colors: K100, White, HEX, CMYK
- 対象：個別／グループとして（選択単位の切替） / Target: Individual or as Group (treat selection as one)
- プレビュー：専用レイヤーに破線で描画、ダイアログ操作中はヒストリーに残りません
  Preview: dashed temporary shapes on dedicated layer; preview actions do not pollute history
- ダイアログ位置・不透明度の記憶 / Persist dialog position & opacity

注意事項 / Notes
------------------------------------------------------------
- 角丸の数値は現状プレビューのみで使用します。実描画ロジックは今後追加します。
- Zオーダーは常に最背面、ターゲットレイヤーは選択しているオブジェクトです。

更新履歴 / Update History
------------------------------------------------------------
- v1.0 (2025-08-22): 初期バージョン。選択オブジェクトを基準に、オフセット長方形を bg レイヤー最背面に作成。
  Added initial version: draws offset rectangles behind selection on bg layer with live preview.
============================================================
*/

var SCRIPT_VERSION = "v1.0";

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
    // Panels
    offsetTitle: {
        ja: "オフセットと角丸",
        en: "Offset & Corner Radius"
    },
    colorTitle: {
        ja: "カラー",
        en: "Color"
    },
    roundTitle: {
        ja: "角丸",
        en: "Corner Radius"
    },

    targetTitle: {
        ja: "対象",
        en: "Target"
    },
    // Color options
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
    colorCustomHint: {
        ja: "例: #FF0000",
        en: "e.g., #FF0000"
    },
    // Target options
    currentAB: {
        ja: "個別",
        en: "Create Individually"
    },
    allAB: {
        ja: "グループとして",
        en: "Create as Group"
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
    },
    previewBounds: {
        ja: "プレビュー境界",
        en: "Preview Bounds"
    },
    previewOutline: {
        ja: "テキストをアウトラインで計算",
        en: "Outline Text for Preview"
    }
};


/* ===== Dialog appearance & position (tunable) ===== */
var DIALOG_OFFSET_X = 300; // shift right (+) / left (-)
var DIALOG_OFFSET_Y = 0; // shift down (+) / up (-)
var DIALOG_OPACITY = 0.98; // 0.0 - 1.0

/* ===== Preview timing (tunable) ===== */
// 入力中のプレビュー遅延（タイプしやすさ優先）/ Delay during typing
var PREVIEW_DELAY_TYPING_MS = 110; // recommend 100–120ms
var PREVIEW_DELAY_OUTLINE_MS = 240; // heavier when outlining text during preview

/* =========================================
 * DialogPersist util (extractable)
 * ダイアログの不透明度・初期位置・位置記憶を共通化するユーティリティ。
 * 使い方:
 *   DialogPersist.setOpacity(dlg, 0.95);
 *   DialogPersist.restorePosition(dlg, "__YourDialogKey", offsetX, offsetY);
 *   DialogPersist.rememberOnMove(dlg, "__YourDialogKey");
 *   DialogPersist.savePosition(dlg, "__YourDialogKey"); // 閉じる直前などに
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

// --- Fill helpers split by responsibility ---
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
 * 単位変換を一元化（Bleedプリセットなし）
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
        layer.move(doc, ElementPlacement.PLACEATEND);
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

    function previewGroup(sel) {
        if (!sel || !sel.length) return;
        var bounds = choice.usePreviewOutline ?
            getCombinedFinalBounds(doc, sel, !!choice.usePreviewBounds) :
            getCombinedGeometricBounds(sel, !!choice.usePreviewBounds);
        if (!bounds) return;
        var left = bounds[0],
            top = bounds[1],
            right = bounds[2],
            bottom = bounds[3];
        var w = right - left;
        var h = top - bottom;
        var o = choice.offset || 0;

        var rect = getOrCreatePreviewRect(previewLayer, 0, top + o, left - o, w + o * 2, h + o * 2);
        var __colG = resolveFillColor(doc, choice.colorMode, {
            customValue: choice.customValue,
            customCMYK: choice.customCMYK
        });
        applyFill(rect, __colG, true);
        rect.stroked = false; // プレビューは線なし
        rect.selected = false;
        rect.zOrder(ZOrderMethod.SENDTOBACK);
    }

    function previewOne(item, idx) {

        var gb = choice.usePreviewOutline ?
            getFinalItemBounds(doc, item, !!choice.usePreviewBounds) :
            getItemBounds(item, !!choice.usePreviewBounds); // [left, top, right, bottom]
        if (!gb) return;
        var left = gb[0],
            top = gb[1],
            right = gb[2],
            bottom = gb[3];
        var w = right - left;
        var h = top - bottom;
        var o = choice.offset || 0;

        var rect = getOrCreatePreviewRect(
            previewLayer,
            idx,
            top + o,
            left - o,
            w + o * 2,
            h + o * 2
        );

        var __col = resolveFillColor(doc, choice.colorMode, {
            customValue: choice.customValue,
            customCMYK: choice.customCMYK
        });
        applyFill(rect, __col, true);

        rect.stroked = false; // プレビューは線なし

        rect.selected = false;
        rect.zOrder(ZOrderMethod.SENDTOBACK); // 常に最背面
    }

    try {
        var sel = doc.selection || [];
        if (choice && choice.target === 'group') {
            previewGroup(sel);
        } else {
            for (var i = 0; i < sel.length; i++) {
                var it = sel[i];
                if (it && it.geometricBounds) previewOne(it, i);
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

    // --- Switch to single-column container ---
    var mainCol = dlg.add('group');
    mainCol.orientation = 'column';
    mainCol.alignChildren = 'fill';
    mainCol.spacing = 12;

    // Offset panel (with margins and row)
    var offsetPanel = mainCol.add('panel', undefined, LABELS.offsetTitle[lang]);
    offsetPanel.orientation = 'column';
    offsetPanel.alignChildren = 'left';
    offsetPanel.margins = [15, 20, 15, 10];

    var offsetRow = offsetPanel.add('group');
    offsetRow.orientation = 'row';
    offsetRow.alignChildren = 'center';
    offsetRow.alignment = 'center';

    var offsetInputLabel = offsetRow.add('statictext', undefined, "オフセット");
    var __LABEL_WIDTH = 70; // unify label width
    try {
        offsetInputLabel.preferredSize.width = __LABEL_WIDTH;
        offsetInputLabel.alignment = ['right', 'center'];
    } catch (e) {}
    var offsetInput = offsetRow.add('edittext', undefined, '2');
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

    var roundRow = offsetPanel.add('group');
    roundRow.orientation = 'row';
    roundRow.alignChildren = 'center';
    roundRow.alignment = 'center';

    var roundInputLabel = roundRow.add('statictext', undefined, "角丸");
    try {
        roundInputLabel.preferredSize.width = __LABEL_WIDTH;
        roundInputLabel.alignment = ['right', 'center'];
    } catch (e) {}

    var roundInput = roundRow.add('edittext', undefined, '2');
    roundInput.characters = 5;
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
    roundRow.add('statictext', undefined, unitLabel2);


    dlg.onShow = function() {
        DialogPersist.restorePosition(dlg, __DLG_KEY, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        try {
            offsetInput.active = true;
        } catch (e) {}
        PreviewHistory.start();
        updatePreviewCommit();
    };
    DialogPersist.rememberOnMove(dlg, __DLG_KEY);

    // Add new panel for color
    var colorPanel = mainCol.add('panel', undefined, LABELS.colorTitle[lang]);
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

    // Add new panel for target (moved to main column)
    var targetPanel = mainCol.add('panel', undefined, LABELS.targetTitle[lang]);
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

    function buildChoiceFromUI() {
        var colorMode = (function() {
            if (k100Radio.value) return ColorMode.K100;
            if (whiteRadio.value) return ColorMode.WHITE;
            if (specifiedRadio.value) return ColorMode.HEX;
            if (cmykRadio.value) return ColorMode.CMYK;
            return ColorMode.K100;
        })();

        // offset計算を一元化
        var unitCode = getCurrentUnitCode();
        var resolved = resolveOffsetToPt(offsetInput.text, unitCode);
        var offsetPt = resolved.pt;

        // Target: 個別 or グループ
        var target = currentRadio.value ? 'individual' :
            (allRadio.value ? 'group' : 'individual');

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
            target: target,
            usePreviewBounds: !!cbPreviewBounds.value,
            usePreviewOutline: !!cbPreviewOutline.value
        };
    }

    function updatePreviewTyping() {
        try {
            schedulePreview(buildChoiceFromUI());
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
            if (__hotkeyBlocked.v) return; // typing in a field
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

    // --- Hotkeys: Target (S/A) and Z-order (F/B/L) ---
    /* 対象スコープ: I/G（個別/グループ）、S/A（レガシー対応）ホットキー */
    function addScopeAndZHotkeys(dialog) {
        dialog.addEventListener('keydown', function(event) {
            if (__hotkeyBlocked.v) return; // ignore when typing in fields
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

    // --- Preview bounds toggle ---
    var boundsRow = dlg.add('group');
    boundsRow.orientation = 'column';
    boundsRow.alignment = 'left';

    boundsRow.alignChildren = ['left', 'center'];

    var cbPreviewBounds = boundsRow.add('checkbox', undefined, LABELS.previewBounds[lang]);
    cbPreviewBounds.value = true; // default ON
    cbPreviewBounds.onClick = updatePreviewCommit;
    cbPreviewBounds.onChanging = updatePreviewCommit;

    var cbPreviewOutline = boundsRow.add('checkbox', undefined, LABELS.previewOutline[lang]);
    cbPreviewOutline.value = true; // default ON
    cbPreviewOutline.onClick = updatePreviewCommit;
    cbPreviewOutline.onChanging = updatePreviewCommit;

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
                    try { doc.layers[i].remove(); } catch (e) {}
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

    var colorMode = null;
    if (k100Radio.value) colorMode = ColorMode.K100;
    else if (whiteRadio && whiteRadio.value) colorMode = ColorMode.WHITE;
    else if (specifiedRadio.value) colorMode = ColorMode.HEX;
    else if (cmykRadio.value) colorMode = ColorMode.CMYK;

    var unitCode2 = getCurrentUnitCode();
    var resolvedFinal = resolveOffsetToPt(offsetInput.text, unitCode2);
    var offset = resolvedFinal.pt;
    var targetFinal = currentRadio.value ? 'individual' : (allRadio.value ? 'group' : 'individual');
    var usePreviewBoundsFinal = !!cbPreviewBounds.value;
    var usePreviewOutlineFinal = !!cbPreviewOutline.value;

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
        target: targetFinal,
        usePreviewBounds: usePreviewBoundsFinal,
        usePreviewOutline: usePreviewOutlineFinal
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

function drawRectangleForItem(doc, item, choice) {
    var gb = item.geometricBounds; // [left, top, right, bottom]
    var left = gb[0],
        top = gb[1],
        right = gb[2],
        bottom = gb[3];
    var w = right - left;
    var h = top - bottom;
    var o = choice.offset;
    // ターゲットレイヤーは選択オブジェクトの属するレイヤー（なければアクティブレイヤー）
    var targetLayer = null;
    try {
        targetLayer = item.layer || null;
    } catch (e) {}
    if (!targetLayer) {
        try {
            targetLayer = doc.activeLayer;
        } catch (_) {}
    }
    ensureLayerEditable(doc, targetLayer);

    var rect = null;
    try {
        rect = targetLayer.pathItems.rectangle(
            top + o,
            left - o,
            w + o * 2,
            h + o * 2
        );
    } catch (e) {
        // Fallback: create on activeLayer, then move into targetLayer
        try {
            var tmpLayer = doc.activeLayer;
            rect = tmpLayer.pathItems.rectangle(
                top + o,
                left - o,
                w + o * 2,
                h + o * 2
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
    applyFill(rect, __finalColor, true);

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

    if (choice && choice.target === 'group') {
        var bounds = getCombinedFinalBounds(doc, sel, !!choice.usePreviewBounds);
        if (bounds) {
            // 代表レイヤー: 最初の選択オブジェクトのレイヤーを採用（混在レイヤー時の基準）
            var repLayer = null;
            try {
                repLayer = (sel && sel.length && sel[0] && sel[0].layer) ? sel[0].layer : doc.activeLayer;
            } catch (__) {
                repLayer = doc.activeLayer;
            }

            var dummy = {
                geometricBounds: bounds,
                layer: repLayer
            };
            drawRectangleForItem(doc, dummy, choice);
        }
    } else {
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            var gb = getFinalItemBounds(doc, it, !!choice.usePreviewBounds);
            if (gb) drawRectangleForItem(doc, {
                geometricBounds: gb,
                layer: (function() {
                    try {
                        return it.layer;
                    } catch (__) {
                        return null;
                    }
                })()
            }, choice);
        }
    }

    // Remove temporary outline layer if exists
    try {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name === '__tmp_outline_bounds__') {
                try { doc.layers[i].remove(); } catch (e) {}
            }
        }
    } catch (e) {}

    try {
        if (__prevCS !== null) app.coordinateSystem = __prevCS;
    } catch (e) {}
}

main();