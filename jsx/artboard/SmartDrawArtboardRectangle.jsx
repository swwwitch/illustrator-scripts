#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要：

- アクティブまたは全アートボードと同サイズの長方形を、オフセットを考慮して描画します。
- カラー（なし／K100 15%／HEX／CMYK）、重ね順（最前面／最背面／bgレイヤー）を指定でき、ライブプレビューで確認できます。
- 描画後に「ガイド化」「ライブシェイプに変換」をオプションで適用できます。
- ライブシェイプには中心の○（中心点）を常に表示します。

### 主な機能：

- オフセット指定（裁ち落としプリセット：3mm／12H／0.125in）
- カラー指定（None／K100 15%／HEX／CMYK）
- 重ね順（Front／Back／bgレイヤー、デフォルトは最前面）
- 対象範囲（作業アートボード／すべて）
- オプション（ガイド化：デフォルトOFF／ライブシェイプに変換：デフォルトON）
- 中心の○（中心点）を常に表示（記録済みアクションで適用）
- プレビュー（1ptの破線、50%トーン、専用レイヤー）
- ホットキー（F：最前面／B：最背面／L：bgレイヤー／G：ガイド化／C：現在のアートボード／A：すべて）
- ダイアログの初期位置・不透明度の設定

### 更新履歴：

- v1.0 (20250820) : 初期バージョン
- v1.5.1 (20250824) : ライブシェイプ化・ガイド化などのオプション追加
- v1.5.2 (20260531) : 中心の○を常に表示、単位テーブル統合、CMYK入力修正、ホットキー再編
- v1.5.3 (20260531) : オブジェクト名を「<長方形>」に変更、オフセット入力欄の幅を調整、最新バージョン

---

### Overview:

- Draws rectangles that match the active or all artboards, with optional offset.
- Supports color (None / K100 15% / HEX / CMYK), stacking order (Front / Back / bg layer), and live preview.
- Optional post-draw actions: "Make guides" and "Convert to Live Shape".
- Always shows the center widget (center point) on live shapes.

### Key Features:

- Offset with Bleed presets (3mm / 12H / 0.125in)
- Color modes (None / K100 15% / HEX / CMYK)
- Z-order (Front / Back / bg layer; defaults to Bring to Front)
- Target scope (Current artboard / All artboards)
- Options (Make guides: default OFF / Convert to Live Shape: default ON)
- Always shows the center widget (applied via a recorded action)
- Preview (1pt dashed stroke, 50% tone, dedicated layer)
- Hotkeys (F: Front / B: Back / L: bg layer / G: Make guides / C: Current artboard / A: All)
- Dialog initial position & opacity settings

### Changelog:

- v1.0 (20250820): Initial version
- v1.5.1 (20250824): Added post-draw options (Convert to Live Shape, Make guides)
- v1.5.2 (20260531): Always show center widget, unified unit table, CMYK input fix, hotkey rework
- v1.5.3 (20260531): Renamed object to "<Rectangle>", tweaked offset field width; latest version

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "v1.5.3";

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期位置・不透明度（調整可）/ Dialog position & opacity (tunable) */
    var DIALOG_OFFSET_X = 300; // 右(+)／左(-) / shift right (+) / left (-)
    var DIALOG_OFFSET_Y = 0; //   下(+)／上(-) / shift down (+) / up (-)
    var DIALOG_OPACITY = 0.97; // 0.0 - 1.0

    /* 入力中のプレビュー遅延（タイプしやすさ優先）/ Preview delay while typing */
    var PREVIEW_DELAY_TYPING_MS = 110; // recommend 100–120ms

    /* 全パネル共通の余白・行間 / Shared panel margins & spacing */
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* カラーモード定数 / Color mode constants */
    var ColorMode = {
        NONE: 'none',
        K100: 'k100',
        HEX: 'hex',
        CMYK: 'cmyk'
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のUI言語を判定 / Detect current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義（カテゴリ別）/ Label definitions (by category) */
    var LABELS = {
        dialog: {
            title: {
                ja: "アートボードサイズの長方形を描画",
                en: "Draw Artboard Size Rectangle"
            }
        },
        panel: {
            offset: {
                ja: "オフセット",
                en: "Offset"
            },
            color: {
                ja: "カラー",
                en: "Color"
            },
            zorder: {
                ja: "配置位置",
                en: "Placement"
            },
            target: {
                ja: "対象",
                en: "Target"
            },
            options: {
                ja: "オプション",
                en: "Options"
            }
        },
        checkbox: {
            bleed: {
                ja: "裁ち落とし",
                en: "Bleed"
            },
            makeGuide: {
                ja: "ガイドに変換",
                en: "Convert to Guides"
            },
            convertToLiveShape: {
                ja: "ライブシェイプ化",
                en: "Convert to Live Shape"
            }
        },
        color: {
            none: {
                ja: "なし",
                en: "None"
            },
            k100: {
                ja: "K100、不透明度15%",
                en: "K100, Opacity 15%"
            },
            hex: {
                ja: "HEX",
                en: "HEX"
            },
            cmyk: {
                ja: "CMYK",
                en: "CMYK"
            },
            hint: {
                ja: "例: #FF0000",
                en: "e.g., #FF0000"
            }
        },
        zorder: {
            front: {
                ja: "最前面",
                en: "Front"
            },
            back: {
                ja: "最背面",
                en: "Back"
            },
            bg: {
                ja: "bgレイヤー",
                en: "bg Layer"
            }
        },
        target: {
            current: {
                ja: "現在のアートボード",
                en: "Current Artboard"
            },
            all: {
                ja: "すべてのアートボード",
                en: "All Artboards"
            }
        },
        button: {
            ok: {
                ja: "OK",
                en: "OK"
            },
            cancel: {
                ja: "キャンセル",
                en: "Cancel"
            },
            previewOutline: {
                ja: "アウトライン表示",
                en: "Outline/Preview"
            },
            previewPreview: {
                ja: "プレビュー表示",
                en: "Outline/Preview"
            }
        },
        helpTip: {
            offsetInput: {
                ja: "アートボード境界から外側へ広げる量を指定します。負の値で内側へ縮めます。",
                en: "Set how far the bounds expand outward from the artboard. Use a negative value to shrink inward."
            },
            bleed: {
                ja: "現在の単位に応じて、裁ち落とし相当の値を自動入力します。",
                en: "Automatically fills a bleed-equivalent offset based on the current unit."
            },
            hexInput: {
                ja: "#RRGGBB 形式で塗りカラーを指定します。",
                en: "Enter a fill color in #RRGGBB format."
            },
            cmykInput: {
                ja: "0〜100の範囲でCMYK値を指定します。未入力は0として扱います。",
                en: "Enter CMYK values from 0 to 100. Empty fields are treated as 0."
            },
            bgLayer: {
                ja: "bgレイヤーを作成または使用し、レイヤーの最背面へ配置します。",
                en: "Creates or uses the bg layer and places it at the back of the layer stack."
            },
            previewToggle: {
                ja: "Illustratorのアウトライン表示／プレビュー表示を切り替えます。",
                en: "Toggles Illustrator's Outline and Preview display modes."
            },
            convertToLiveShape: {
                ja: "Illustratorのメニューコマンドで長方形をライブシェイプ化します。中心点も表示されます。",
                en: "Uses Illustrator's menu command to convert rectangles to Live Shapes. The center point is also shown."
            },
            makeGuide: {
                ja: "描画した長方形をガイドに変換します。",
                en: "Converts the drawn rectangles to guides."
            }
        },
        name: {
            previewLayer: {
                ja: "_preview",
                en: "_preview"
            },
            rect: {
                ja: "<長方形>",
                en: "<Rectangle>"
            },
            previewRect: {
                ja: "__プレビュー_アートボードサイズの長方形",
                en: "__Preview_ArtboardSizeRectangle"
            }
        }
    };

    /* ラベル取得（ドット区切りキー、{slash}→「/」に展開）/ Resolve label by dotted key; {slash}→"/" */
    function getLocalizedText(key) {
        var node = LABELS;
        var parts = String(key).split('.');
        for (var i = 0; i < parts.length; i++) {
            if (node == null) break;
            node = node[parts[i]];
        }
        var text = (node && node[currentLanguage] != null) ? node[currentLanguage] : key;
        return String(text).replace(/\{slash\}/g, '/');
    }

    /* =========================================
     * DialogPersist util (extractable)
     * ダイアログの不透明度・初期位置を共通化するユーティリティ。
     * 使い方:
     *   DialogPersist.setOpacity(dialog, 0.95);
     *   DialogPersist.applyInitialOffset(dialog, offsetX, offsetY); // onShow などで
     * ========================================= */
    (function (g) {
        if (!g.DialogPersist) {
            g.DialogPersist = {
                setOpacity: function (dialog, v) {
                    try { dialog.opacity = v; } catch (e) { }
                },
                applyInitialOffset: function (dialog, offsetX, offsetY) {
                    try {
                        var l = dialog.location;
                        dialog.location = [l[0] + (offsetX | 0), l[1] + (offsetY | 0)];
                    } catch (e) { }
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
        } catch (e) { }
    }

    /* ドキュメントのカラースペースに合わせた黒を生成 / Build black matching the document color space */
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
    function clampValue(v, min, max) {
        return v < min ? min : (v > max ? max : v);
    }

    /* RGBColor を生成（0–255にクランプ）/ Build an RGBColor (clamped to 0–255) */
    function makeRGB(r, g, b) {
        var c = new RGBColor();
        c.red = clampValue(Math.round(r), 0, 255);
        c.green = clampValue(Math.round(g), 0, 255);
        c.blue = clampValue(Math.round(b), 0, 255);
        return c;
    }

    /* CMYKColor を生成（0–100にクランプ）/ Build a CMYKColor (clamped to 0–100) */
    function makeCMYK(cy, mg, yl, k) {
        var c = new CMYKColor();
        c.cyan = clampValue(cy, 0, 100);
        c.magenta = clampValue(mg, 0, 100);
        c.yellow = clampValue(yl, 0, 100);
        c.black = clampValue(k, 0, 100);
        return c;
    }

    // --- RGB/CMYK conversion helpers ---
    function rgbToCmyk(r, g, b) {
        // r,g,b: 0-255 → return [C,M,Y,K] 0-100
        r = clampValue(r, 0, 255) / 255;
        g = clampValue(g, 0, 255) / 255;
        b = clampValue(b, 0, 255) / 255;
        var k = 1 - Math.max(r, g, b);
        if (k >= 0.9999) return [0, 0, 0, 100];
        var c = (1 - r - k) / (1 - k);
        var m = (1 - g - k) / (1 - k);
        var y = (1 - b - k) / (1 - k);
        return [Math.round(c * 100), Math.round(m * 100), Math.round(y * 100), Math.round(k * 100)];
    }

    function cmykToRgb(c, m, y, k) {
        // c,m,y,k: 0-100 → return [R,G,B] 0-255
        c = clampValue(c, 0, 100) / 100;
        m = clampValue(m, 0, 100) / 100;
        y = clampValue(y, 0, 100) / 100;
        k = clampValue(k, 0, 100) / 100;
        var r = 255 * (1 - c) * (1 - k);
        var g = 255 * (1 - m) * (1 - k);
        var b = 255 * (1 - y) * (1 - k);
        return [Math.round(r), Math.round(g), Math.round(b)];
    }

    // --- Common fill applier for both preview and final draw ---
    /*
     * applyFillByMode(doc, targetRect, mode, payload, opts)
     * - mode: 'none' | 'k100' | 'hex' | 'cmyk'
     * - payload: { customValue: String, customCMYK: {c,m,y,k} }
     * - opts: { k100Opacity: Number }
     */
    function applyFillByMode(doc, targetRect, mode, payload, opts) {
        try {
            var k100Opacity = (opts && typeof opts.k100Opacity === 'number') ? opts.k100Opacity : 15;
            if (mode === 'none') {
                targetRect.filled = false;
                targetRect.stroked = false;
                return;
            }
            if (mode === 'k100') {
                targetRect.filled = true;
                targetRect.fillColor = createBlackColor(doc);
                targetRect.stroked = false;
                targetRect.opacity = k100Opacity;
                return;
            }
            if (mode === 'hex') {
                var col = parseCustomColor(doc, payload && payload.customValue);
                if (col) {
                    targetRect.filled = true;
                    targetRect.fillColor = col;
                    targetRect.stroked = false;
                    targetRect.opacity = 100;
                } else {
                    targetRect.filled = false;
                    targetRect.stroked = false;
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
                    targetRect.filled = false;
                    targetRect.stroked = false;
                    return;
                }
                var color;
                if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
                    var rgb = cmykToRgb(c, m, y, k);
                    color = makeRGB(rgb[0], rgb[1], rgb[2]);
                } else {
                    color = makeCMYK(c, m, y, k);
                }
                targetRect.filled = true;
                targetRect.fillColor = color;
                targetRect.stroked = false;
                targetRect.opacity = 100;
                return;
            }
            // Fallback
            targetRect.filled = false;
            targetRect.stroked = false;
        } catch (e) { }
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
            s = s.replace(/[０-９]/g, function (ch) {
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
                var gk = clampValue(parseInt(m[1], 10), 0, 100);
                if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK) return makeCMYK(0, 0, 0, gk);
                var v = Math.round(255 * (100 - gk) / 100);
                return makeRGB(v, v, v);
            }
        } catch (e) { }
        return null;
    }

    // 単位コード→ラベルとpt係数のテーブル（rulerType基準）
    // Map rulerType codes to label & points-per-unit factor
    var UNIT_TABLE = {
        0:  { label: "in",    factor: 72.0 },                 // inch
        1:  { label: "mm",    factor: 72.0 / 25.4 },          // mm
        2:  { label: "pt",    factor: 1.0 },                  // pt
        3:  { label: "pica",  factor: 12.0 },                 // pica
        4:  { label: "cm",    factor: 72.0 / 2.54 },          // cm
        5:  { label: "Q/H",   factor: 72.0 / 25.4 * 0.25 },   // Q or H
        6:  { label: "px",    factor: 1.0 },                  // px (Illustrator: 1px=1pt)
        7:  { label: "ft/in", factor: 72.0 * 12.0 },          // ft/in
        8:  { label: "m",     factor: 72.0 / 25.4 * 1000.0 }, // m
        9:  { label: "yd",    factor: 72.0 * 36.0 },          // yd
        10: { label: "ft",    factor: 72.0 * 12.0 }           // ft
    };

    // 現在の単位コードを取得 / Get current rulerType code
    function getCurrentUnitCode() {
        try {
            return app.preferences.getIntegerPreference("rulerType");
        } catch (e) {
            return 2; // fallback to pt
        }
    }

    // 現在の単位ラベルを取得 / Get current unit label from prefs
    function getCurrentUnitLabel() {
        var entry = UNIT_TABLE[getCurrentUnitCode()];
        return entry ? entry.label : "pt";
    }

    // 単位コード→pt係数 / Convert unit code to points factor
    function getPtFactorFromUnitCode(code) {
        var entry = UNIT_TABLE[code];
        return entry ? entry.factor : 1.0;
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

    /* 上下キーで入力値を増減（Shift=10, Option=0.1）/ Increment value with arrow keys (Shift=10, Option=0.1) */
    function changeValueByArrowKey(editText, onValueChange) {
        editText.addEventListener("keydown", function (event) {
            // 上下キー以外（通常の文字入力など）には一切干渉しない
            // Only react to Up/Down; never touch the field on normal character input
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) value = 0; // 空欄などは0起点 / treat empty/invalid as 0

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
            } catch (e) { }
        });
    }

    // ===== Preview helpers =====

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 使い方:
     *   PreviewHistory.start();     // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();      // プレビュー描画ごとにカウント(+1)
     *   PreviewHistory.undo();      // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */
    (function (g) {
        if (!g.PreviewHistory) {
            g.PreviewHistory = {
                start: function () {
                    g.__previewUndoCount = 0;
                },
                bump: function () {
                    g.__previewUndoCount = (g.__previewUndoCount | 0) + 1;
                },
                undo: function () {
                    var n = g.__previewUndoCount | 0;
                    try {
                        for (var i = 0; i < n; i++) app.executeMenuCommand('undo');
                    } catch (e) { }
                    g.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

    var __previewDebounceTask = null;

    /* プレビュー描画を遅延スケジュール（デバウンス）/ Schedule a debounced preview render */
    function schedulePreview(choice, delayMs) {
        if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
        $.global.__lastPreviewChoice = choice;
        // scheduleTask の文字列はグローバルスコープで評価されるため、IIFE内の renderPreview を $.global 経由で参照
        // scheduleTask runs its string in global scope, so expose renderPreview via $.global (survives the IIFE wrapper)
        $.global.__renderPreviewRef = renderPreview;
        var code = 'try{$.global.__renderPreviewRef(app.activeDocument, $.global.__lastPreviewChoice);}catch(e){}';
        try {
            __previewDebounceTask = app.scheduleTask(code, Math.max(0, delayMs | 0), false);
        } catch (e) {
            try {
                renderPreview(app.activeDocument, choice);
            } catch (_) { }
        }
    }

    /*
     * プレビュー要求（即時/遅延）/ Request preview (immediate or debounced)
     */
    function requestPreview(choice, immediate) {
        if (immediate) {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
            try {
                renderPreview(app.activeDocument, choice);
            } catch (_) { }
        } else {
            schedulePreview(choice, PREVIEW_DELAY_TYPING_MS);
        }
    }

    function clearPreview(removeLayer) {
        try {
            var doc = app.activeDocument;
            var names = [getLocalizedText('name.previewLayer'), "プレビュー", "Preview", "_preview"]; // legacy names
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var layer = doc.layers[i];
                var nm = layer.name;
                for (var j = 0; j < names.length; j++) {
                    if (nm === names[j]) {
                        if (removeLayer) {
                            try {
                                layer.remove();
                            } catch (e) { }
                        } else {
                            // live update: hide only preview items created by this script
                            var previewNameBase = getLocalizedText('name.previewRect') + "#";
                            try {
                                for (var k = layer.pathItems.length - 1; k >= 0; k--) {
                                    try {
                                        var itemName = String(layer.pathItems[k].name || "");
                                        if (itemName.indexOf(previewNameBase) === 0) {
                                            layer.pathItems[k].hidden = true;
                                        }
                                    } catch (e) { }
                                }
                            } catch (e) { }
                        }
                        break;
                    }
                }
            }
        } catch (e) { }
    }

    /* プレビュー専用レイヤーを取得（なければ作成し最前面へ）/ Get or create the dedicated preview layer (front-most) */
    function getOrCreatePreviewLayer(doc) {
        var name = getLocalizedText('name.previewLayer');
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
        } catch (e) { }
        return layer;
    }

    // Helper to fetch or create a preview rectangle for a specific artboard index
    function getOrCreatePreviewRect(previewLayer, artboardIndex, top, left, width, height) {
        // Try to reuse existing item named with index suffix
        var nameBase = getLocalizedText('name.previewRect') + "#" + artboardIndex;
        var item = null;
        try {
            for (var i = 0; i < previewLayer.pathItems.length; i++) {
                var it = previewLayer.pathItems[i];
                if (it.name === nameBase) {
                    item = it;
                    break;
                }
            }
        } catch (e) { }
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
            } catch (e) { }
            try {
                item.hidden = false;
            } catch (e) { }
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
        } catch (e) { }
        try {
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }

        var prevIndex = doc.artboards.getActiveArtboardIndex();
        var previewLayer = getOrCreatePreviewLayer(doc);

        function renderPreviewForArtboard(artboardIndex) {
            try {
                // Avoid switching active artboard for preview / プレビューではアクティブAB切替を行わない
            } catch (e) { }
            var artboard = doc.artboards[artboardIndex];
            var artboardRect = artboard.artboardRect;
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];
            var o = choice.offset || 0;
            var artboardRectangle = getOrCreatePreviewRect(
                previewLayer,
                artboardIndex,
                artboardRect[1] + o,
                artboardRect[0] - o,
                artboardWidth + o * 2,
                artboardHeight + o * 2
            );

            // Unified fill application (preview)
            applyFillByMode(doc, artboardRectangle, choice.colorMode, {
                customValue: choice.customValue,
                customCMYK: choice.customCMYK
            }, {
                k100Opacity: 15
            });

            // Common preview stroke (visibility)
            artboardRectangle.stroked = true;
            artboardRectangle.strokeWidth = 1;
            try {
                artboardRectangle.strokeDashes = [6, 4];
            } catch (e) { }
            artboardRectangle.strokeColor = getPreviewStrokeColor(doc);

            artboardRectangle.selected = false;
            if (choice.zOrder === 'front') artboardRectangle.zOrder(ZOrderMethod.BRINGTOFRONT);
            else if (choice.zOrder === 'back') artboardRectangle.zOrder(ZOrderMethod.SENDTOBACK);
        }

        // --- Draw previews for target scope ---
        if (choice.target === 'all') {
            for (var i = 0; i < doc.artboards.length; i++) renderPreviewForArtboard(i);
            try {
                doc.artboards.setActiveArtboardIndex(prevIndex);
            } catch (e) { }
        } else {
            renderPreviewForArtboard(doc.artboards.getActiveArtboardIndex());
        }

        try {
            // Hide non-target preview items to avoid deletions during typing
            for (var i = 0; i < previewLayer.pathItems.length; i++) {
                var it = previewLayer.pathItems[i];
                // items are named like "__Preview_...#<artboardIndex>"
                var m = /#(\d+)$/.exec(it.name || "");
                if (m) {
                    var idn = parseInt(m[1], 10);
                    if (choice.target === 'all') {
                        if (idn < 0 || idn > doc.artboards.length - 1) {
                            try {
                                it.hidden = true;
                            } catch (_) { }
                        }
                    } else {
                        var active = doc.artboards.getActiveArtboardIndex();
                        if (idn !== active) {
                            try {
                                it.hidden = true;
                            } catch (_) { }
                        }
                    }
                }
            }
        } catch (e) { }

        try {
            if (prevCS !== null) app.coordinateSystem = prevCS;
        } catch (e) { }
        PreviewHistory.bump();
        app.redraw();
    }

    /* パネルの体裁を共通設定（余白・行間・整列）/ Apply shared panel layout (margins, spacing, alignment) */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* オプションパネルを構築（ガイド化／ライブシェイプ変換）
     * Build the options panel (make guides / convert to live shape).
     * 各オプションは独立。必要なものだけ描画後に適用（applyDrawOptions）
     * Options are independent; only enabled options are applied after drawing (see applyDrawOptions)
     */
    function buildPostProcessOptionsPanel(parent) {
        var panel = parent.add('panel', undefined, getLocalizedText('panel.options'));
        setupPanel(panel);

        var makeGuide = panel.add('checkbox', undefined, getLocalizedText('checkbox.makeGuide'));
        makeGuide.value = false; // デフォルトOFF / default OFF

        var convertToLiveShape = panel.add('checkbox', undefined, getLocalizedText('checkbox.convertToLiveShape'));
        convertToLiveShape.value = true; // デフォルトON / default ON

        // 各オプションは独立。必要なものだけ描画後に適用（applyDrawOptions）
        // Options are independent; only enabled options are applied after drawing (see applyDrawOptions)

        return { makeGuide: makeGuide, convertToLiveShape: convertToLiveShape };
    }

    /* 設定ダイアログを構築して結果を返す / Build the settings dialog and return the user's choice */
    function showDialog() {
        var dialog = new Window('dialog', getLocalizedText('dialog.title') + ' ' + SCRIPT_VERSION);
        DialogPersist.setOpacity(dialog, DIALOG_OPACITY);
        dialog.alignChildren = 'left';

        // --- Add two-column group container ---
        var mainColumnsGroup = dialog.add('group');
        mainColumnsGroup.orientation = 'row';
        mainColumnsGroup.alignChildren = ['fill', 'top'];
        mainColumnsGroup.spacing = 12;

        var leftColumnGroup = mainColumnsGroup.add('group');
        leftColumnGroup.orientation = 'column';
        leftColumnGroup.alignChildren = 'fill';
        leftColumnGroup.spacing = 10;

        var rightColumnGroup = mainColumnsGroup.add('group');
        rightColumnGroup.orientation = 'column';
        rightColumnGroup.alignChildren = 'fill';
        rightColumnGroup.spacing = 10;

        // Limit column widths to keep layout balanced

        // Offset panel (with margins and row)
        var offsetSettingsPanel = leftColumnGroup.add('panel', undefined, getLocalizedText('panel.offset'));
        setupPanel(offsetSettingsPanel);

        var offsetRow = offsetSettingsPanel.add('group');
        offsetRow.orientation = 'row';
        offsetRow.alignChildren = 'center';
        offsetRow.alignment = 'center';

        var offsetInput = offsetRow.add('edittext', undefined, '0');
        offsetInput.characters = 4;
        offsetInput.helpTip = getLocalizedText('helpTip.offsetInput');
        changeValueByArrowKey(offsetInput, function () {
            requestDelayedPreviewUpdate();
        });

        // Enterキーでも明示的に更新
        offsetInput.addEventListener('keydown', function (e) {
            if (e.keyName == 'Enter') {
                try {
                    requestPreview(buildDialogSettings(), true);
                } catch (_) { }
            }
        });

        var unitLabel = getCurrentUnitLabel();
        offsetRow.add('statictext', undefined, unitLabel);

        // Bleed checkbox row
        var bleedRow = offsetSettingsPanel.add('group');
        bleedRow.orientation = 'row';
        bleedRow.alignChildren = 'center';
        bleedRow.alignment = 'center';
        var bleedCheckbox = bleedRow.add('checkbox', undefined, getLocalizedText('checkbox.bleed'));
        bleedCheckbox.alignment = 'center';
        bleedCheckbox.value = false; // default OFF
        bleedCheckbox.helpTip = getLocalizedText('helpTip.bleed');

        // Preserve user's manual offset when toggling Bleed
        var lastUserOffsetText = '0';

        function applyBleedPreset(toPreview) {
            var unitCode = getCurrentUnitCode();
            var res = resolveOffsetToPt(offsetInput.text, unitCode, true);
            try {
                offsetInput.text = res.displayText;
            } catch (e) { }
            offsetInput.enabled = !res.disabled; // disabled when bleed ON
            if (toPreview === true) {
                requestDelayedPreviewUpdate();
            }
        }

        function removeBleedPreset(toPreview) {
            try {
                offsetInput.text = lastUserOffsetText;
            } catch (e) { }
            offsetInput.enabled = true;
            if (toPreview === true) {
                requestDelayedPreviewUpdate();
            }
        }

        dialog.onShow = function () {
            DialogPersist.applyInitialOffset(dialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
            try {
                offsetInput.active = true;
            } catch (e) { }
            // Remember initial value and set field state
            lastUserOffsetText = String(offsetInput.text);
            if (bleedCheckbox.value) {
                applyBleedPreset(false);
            } else {
                offsetInput.enabled = true;
            }
            PreviewHistory.start(); // reset preview history counter
            // Render initial preview
            updatePreviewImmediately();
        };

        // Add new panel for color
        var colorSettingsPanel = rightColumnGroup.add('panel', undefined, getLocalizedText('panel.color'));
        setupPanel(colorSettingsPanel, 10); // やや広めの行間 / a bit more vertical gap between rows

        var noneRadio = colorSettingsPanel.add('radiobutton', undefined, getLocalizedText('color.none'));
        var k100Radio = colorSettingsPanel.add('radiobutton', undefined, getLocalizedText('color.k100'));

        // HEX radio + input on the same row
        var hexRow = colorSettingsPanel.add('group');
        hexRow.orientation = 'row';
        hexRow.alignment = 'left';
        hexRow.alignChildren = ['left', 'center'];
        hexRow.spacing = 6;
        var hexRadio = hexRow.add('radiobutton', undefined, getLocalizedText('color.hex'));
        var hexInput = hexRow.add('edittext', undefined, '#');
        hexInput.characters = 14; // narrower to avoid column growth
        hexInput.helpTip = getLocalizedText('helpTip.hexInput');

        // --- HEX validation & feedback ---
        function setHexWarn(et, warn, msg) {
            try {
                var g = et.graphics;
                var pen = g.newPen(g.PenType.SOLID_COLOR, warn ? [1, 0, 0] : [0, 0, 0], 1);
                g.foregroundColor = pen; // text color fallback for border
                if (warn) {
                    et.helpTip = (currentLanguage === 'ja') ? (msg || '正しい #RRGGBB を入力してください') : (msg || 'Enter a valid #RRGGBB value');
                } else {
                    et.helpTip = getLocalizedText('helpTip.hexInput');
                }
                et.notify('onDraw');
            } catch (e) { }
        }

        hexInput.onChanging = function () {
            try {
                var t = String(hexInput.text || '').replace(/\s+/g, '');
                if (t === '') {
                    setHexWarn(hexInput, false);
                    updatePreviewWhileTyping();
                    return;
                }
                if (t === '#') {
                    setHexWarn(hexInput, true, (currentLanguage === 'ja') ? 'HEX未入力（# のみ）' : 'HEX not entered (# only)');
                    updatePreviewWhileTyping();
                    return;
                }
                // Validate only exact #RRGGBB on-the-fly
                var valid = /^#([0-9a-fA-F]{6})$/.test(t);
                setHexWarn(hexInput, !valid);
                updatePreviewWhileTyping();
            } catch (e) { }
        };

        hexInput.onChange = function () {
            try {
                var t = String(hexInput.text || '').replace(/\s+/g, '');
                if (/^[0-9a-fA-F]{6}$/.test(t)) {
                    t = '#' + t.toUpperCase();
                    hexInput.text = t;
                    setHexWarn(hexInput, false);
                } else if (/^#([0-9a-fA-F]{6})$/.test(t)) {
                    hexInput.text = ('#' + RegExp.$1.toUpperCase());
                    setHexWarn(hexInput, false);
                } else if (t === '#') {
                    setHexWarn(hexInput, true, (currentLanguage === 'ja') ? 'HEX未入力（# のみ）' : 'HEX not entered (# only)');
                } else {
                    setHexWarn(hexInput, true);
                }
                updatePreviewImmediately();
            } catch (e) { }
        };

        // CMYK mode radio
        var cmykRadio = colorSettingsPanel.add('radiobutton', undefined, getLocalizedText('color.cmyk'));

        // Custom CMYK input fields (two-row grid: labels on top, fields below)
        var cmykRow = colorSettingsPanel.add('group');
        cmykRow.orientation = 'column';
        cmykRow.alignment = 'left';
        cmykRow.spacing = 4;

        var cmykLabelRow = cmykRow.add('group');
        cmykLabelRow.orientation = 'row';
        cmykLabelRow.alignChildren = ['left', 'center'];
        cmykLabelRow.spacing = 10;

        var cmykFieldRow = cmykRow.add('group');
        cmykFieldRow.orientation = 'row';
        cmykFieldRow.alignChildren = ['left', 'center'];
        cmykFieldRow.spacing = 10;

        var cmykColWidth = 40; // fixed width to align columns

        var cyanLabel = cmykLabelRow.add('statictext', undefined, '  C');
        cyanLabel.preferredSize.width = cmykColWidth;
        var magentaLabel = cmykLabelRow.add('statictext', undefined, '  M');
        magentaLabel.preferredSize.width = cmykColWidth;
        var yellowLabel = cmykLabelRow.add('statictext', undefined, '  Y');
        yellowLabel.preferredSize.width = cmykColWidth;
        var blackLabel = cmykLabelRow.add('statictext', undefined, '  K');
        blackLabel.preferredSize.width = cmykColWidth;

        var cyanInput = cmykFieldRow.add('edittext', undefined, '');
        cyanInput.characters = 3;
        cyanInput.preferredSize.width = cmykColWidth;
        var magentaInput = cmykFieldRow.add('edittext', undefined, '');
        magentaInput.characters = 3;
        magentaInput.preferredSize.width = cmykColWidth;
        var yellowInput = cmykFieldRow.add('edittext', undefined, '');
        yellowInput.characters = 3;
        yellowInput.preferredSize.width = cmykColWidth;
        var blackInput = cmykFieldRow.add('edittext', undefined, '');
        blackInput.characters = 3;
        blackInput.preferredSize.width = cmykColWidth;
        cyanInput.helpTip = getLocalizedText('helpTip.cmykInput');
        magentaInput.helpTip = getLocalizedText('helpTip.cmykInput');
        yellowInput.helpTip = getLocalizedText('helpTip.cmykInput');
        blackInput.helpTip = getLocalizedText('helpTip.cmykInput');

        // --- CMYK validation helpers (empty→0, clamp 0–100, red text warning) ---
        function setCmykFieldWarn(et, warn) {
            try {
                var g = et.graphics;
                var pen = g.newPen(g.PenType.SOLID_COLOR, warn ? [1, 0, 0] : [0, 0, 0], 1);
                g.foregroundColor = pen; // text color as fallback to 'red border'
                et.helpTip = warn ? '0–100 の範囲にしてください（未入力は 0 として扱います）' : getLocalizedText('helpTip.cmykInput');
            } catch (e) { }
        }

        function validateCmykField(et) {
            try {
                var t = String(et.text || '');
                if (t === '') {
                    setCmykFieldWarn(et, false);
                    return;
                } // typing phase, don't warn
                var n = parseFloat(t);
                var warn = (isNaN(n) || n < 0 || n > 100);
                setCmykFieldWarn(et, warn);
            } catch (e) { }
        }

        function clampCmykField(et) {
            try {
                var t = String(et.text || '');
                // 未入力は空のまま（計算時に0扱い）。不要な「0」を表示しない
                // Keep empty fields empty (treated as 0 at compute time); don't insert a spurious "0"
                if (t.replace(/^\s+|\s+$/g, '') === '') {
                    setCmykFieldWarn(et, false);
                    return;
                }
                var n = parseFloat(t);
                if (isNaN(n)) n = 0; // invalid -> 0
                n = clampValue(n, 0, 100);
                et.text = String(n);
                setCmykFieldWarn(et, false);
            } catch (e) { }
        }
        // If the field currently holds exactly "0", clear it on focus for easier typing
        function clearZeroOnFocus(et) {
            try {
                et.addEventListener('focus', function () {
                    try {
                        if (String(et.text) === '0') {
                            et.text = '';
                            // caret will be at the end by default
                        }
                    } catch (e) { }
                });
            } catch (e) { }
        }

        // Prevent any leading-zero integer like "03" from being typed (but allow decimals like "0.5")
        function replaceZeroOnFirstDigit(et) {
            try {
                et.addEventListener('keydown', function (ev) {
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
                    } catch (e) { }
                });
            } catch (e) { }
        }

        // Bind common handlers to a CMYK EditText
        function bindCmykField(et) {
            et.onChanging = function () {
                try {
                    var t = String(et.text || '');
                    // Normalize leading zeros for integers: "03" -> "3", "007" -> "7"
                    // Do NOT touch decimals like "0.5" (only pure digits)
                    if (/^0\d+$/.test(t)) {
                        et.text = t.replace(/^0+/, '');
                        t = String(et.text || '');
                    }
                } catch (e) { }
                validateCmykField(et);
                updatePreviewWhileTyping();
            };
            et.onChange = function () {
                clampCmykField(et);
                updatePreviewImmediately();
            };
            changeValueByArrowKey(et, function () {
                clampCmykField(et);
                updatePreviewWhileTyping();
            });
        }

        // --- Hotkey guard: disable N/K/H/C while typing in fields ---
        var hotkeyState = {
            v: false
        };

        function _attachBlockOnFocusBlur(ctrl) {
            try {
                ctrl.addEventListener('focus', function () {
                    hotkeyState.v = true;
                });
            } catch (e) { }
            try {
                ctrl.addEventListener('blur', function () {
                    hotkeyState.v = false;
                });
            } catch (e) { }
        }
        // Block on all edit fields
        _attachBlockOnFocusBlur(offsetInput);
        _attachBlockOnFocusBlur(hexInput);
        _attachBlockOnFocusBlur(cyanInput);
        _attachBlockOnFocusBlur(magentaInput);
        _attachBlockOnFocusBlur(yellowInput);
        _attachBlockOnFocusBlur(blackInput);

        clearZeroOnFocus(cyanInput);
        clearZeroOnFocus(magentaInput);
        clearZeroOnFocus(yellowInput);
        clearZeroOnFocus(blackInput);

        replaceZeroOnFirstDigit(cyanInput);
        replaceZeroOnFirstDigit(magentaInput);
        replaceZeroOnFirstDigit(yellowInput);
        replaceZeroOnFirstDigit(blackInput);

        // --- Bind CMYK fields (common handlers)
        bindCmykField(cyanInput);
        bindCmykField(magentaInput);
        bindCmykField(yellowInput);
        bindCmykField(blackInput);

        // Default selection
        k100Radio.value = true;

        // Enable custom field only when "Custom" is selected

        /* HEX 有効/無効を切替 / Enable-Disable HEX input */
        function setHexEnabled(on) {
            try {
                hexInput.enabled = !!on;
            } catch (e) { }
        }

        /* CMYK 有効/無効を切替 / Enable-Disable CMYK inputs */
        function setCmykEnabled(on) {
            var v = !!on;
            try {
                cyanInput.enabled = v;
                cyanLabel.enabled = v;
                magentaInput.enabled = v;
                magentaLabel.enabled = v;
                yellowInput.enabled = v;
                yellowLabel.enabled = v;
                blackInput.enabled = v;
                blackLabel.enabled = v;
                if (!v) {
                    setCmykFieldWarn(cyanInput, false);
                    setCmykFieldWarn(magentaInput, false);
                    setCmykFieldWarn(yellowInput, false);
                    setCmykFieldWarn(blackInput, false);
                }
            } catch (e) { }
        }

        /* ラジオ選択に応じて一括反映 / Apply enable states from radio values */
        function updateColorEnableFromRadios() {
            setHexEnabled(!!hexRadio.value);
            setCmykEnabled(!!cmykRadio.value);
        }

        updateColorEnableFromRadios();

        // Add new panel for zOrder
        var placementPanel = leftColumnGroup.add('panel', undefined, getLocalizedText('panel.zorder'));
        setupPanel(placementPanel);

        var frontRadio = placementPanel.add('radiobutton', undefined, getLocalizedText('zorder.front'));
        var backRadio = placementPanel.add('radiobutton', undefined, getLocalizedText('zorder.back'));
        var bgLayerRadio = placementPanel.add('radiobutton', undefined, getLocalizedText('zorder.bg'));
        bgLayerRadio.helpTip = getLocalizedText('helpTip.bgLayer');

        frontRadio.value = true; // デフォルトは最前面 / default to Bring to Front

        // オプションパネル（ガイド化／ライブシェイプ変換）/ Options panel (make guides / convert to live shape)
        var optionControls = buildPostProcessOptionsPanel(leftColumnGroup);
        var makeGuideCheckbox = optionControls.makeGuide;
        var convertToLiveShapeCheckbox = optionControls.convertToLiveShape;
        makeGuideCheckbox.helpTip = getLocalizedText('helpTip.makeGuide');
        convertToLiveShapeCheckbox.helpTip = getLocalizedText('helpTip.convertToLiveShape');

        // Add new panel for target (moved to right column)
        var targetScopePanel = rightColumnGroup.add('panel', undefined, getLocalizedText('panel.target'));
        setupPanel(targetScopePanel);

        var currentArtboardRadio = targetScopePanel.add('radiobutton', undefined, getLocalizedText('target.current'));
        var allArtboardsRadio = targetScopePanel.add('radiobutton', undefined, getLocalizedText('target.all'));

        // 常に「作業中のアートボードのみ」をデフォルト選択
        var artboardCount = (app.documents.length ? app.activeDocument.artboards.length : 0);
        currentArtboardRadio.value = true;
        allArtboardsRadio.value = false;

        // 1枚しかない場合は「すべてのアートボード」をディム（無効化）
        if (artboardCount <= 1) {
            try {
                allArtboardsRadio.enabled = false;
                allArtboardsRadio.helpTip = (currentLanguage === 'ja') ? 'アートボードが1つのため選択できません' : 'Disabled: only one artboard exists';
            } catch (e) { }
        }

        function buildDialogSettings() {
            var colorMode = (function () {
                if (noneRadio.value) return ColorMode.NONE;
                if (k100Radio.value) return ColorMode.K100;
                if (hexRadio.value) return ColorMode.HEX;
                if (cmykRadio.value) return ColorMode.CMYK;
                return ColorMode.NONE;
            })();

            var zOrder = frontRadio.value ? 'front' : (backRadio.value ? 'back' : (bgLayerRadio.value ? 'bg' : 'back'));
            var target = currentArtboardRadio.value ? 'current' : (allArtboardsRadio.value ? 'all' : 'current');

            // offset計算を一元化
            var unitCode = getCurrentUnitCode();
            var resolved = resolveOffsetToPt(offsetInput.text, unitCode, !!bleedCheckbox.value);
            var offsetPt = resolved.pt;

            var customValue = '';
            try {
                customValue = String(hexInput.text || '').replace(/^\s+|\s+$/g, '');
            } catch (e) { }

            var cmykObj = {
                c: 0,
                m: 0,
                y: 0,
                k: 0
            };
            try {
                var cTmp = parseFloat(cyanInput.text);
                if (isNaN(cTmp)) cTmp = 0;
                cmykObj.c = clampValue(cTmp, 0, 100);
                var mTmp = parseFloat(magentaInput.text);
                if (isNaN(mTmp)) mTmp = 0;
                cmykObj.m = clampValue(mTmp, 0, 100);
                var yTmp = parseFloat(yellowInput.text);
                if (isNaN(yTmp)) yTmp = 0;
                cmykObj.y = clampValue(yTmp, 0, 100);
                var kTmp = parseFloat(blackInput.text);
                if (isNaN(kTmp)) kTmp = 0;
                cmykObj.k = clampValue(kTmp, 0, 100);
            } catch (e) { }

            return {
                colorMode: colorMode,
                customValue: customValue, // HEX文字列
                customCMYK: cmykObj, // CMYK値
                offset: offsetPt,
                zOrder: zOrder,
                target: target,
                bleed: !!bleedCheckbox.value,
                makeGuide: !!makeGuideCheckbox.value,
                convertToLiveShape: !!convertToLiveShapeCheckbox.value
            };
        }

        function updatePreviewWhileTyping() {
            try {
                schedulePreview(buildDialogSettings(), PREVIEW_DELAY_TYPING_MS);
            } catch (e) { }
        }

        function updatePreviewImmediately() {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
            try {
                renderPreview(app.activeDocument, buildDialogSettings());
            } catch (_) { }
        }

        function requestDelayedPreviewUpdate() {
            try {
                requestPreview(buildDialogSettings(), false);
            } catch (e) { }
        }

        offsetInput.onChanging = updatePreviewWhileTyping;
        offsetInput.onChange = updatePreviewImmediately;

        /* カラーモードのラジオを排他選択し、ハイライト・フォーカス・プレビューを更新
         * Select a color-mode radio exclusively, then sync highlight/focus/preview
         */
        function selectColorMode(mode) {
            noneRadio.value = (mode === ColorMode.NONE);
            k100Radio.value = (mode === ColorMode.K100);
            hexRadio.value = (mode === ColorMode.HEX);
            cmykRadio.value = (mode === ColorMode.CMYK);
            updateColorEnableFromRadios();
            setEditHighlight(hexInput, mode === ColorMode.HEX);
            setEditHighlight(cyanInput, mode === ColorMode.CMYK);
            if (mode === ColorMode.HEX) {
                try { hexInput.active = true; } catch (e) { }
            } else if (mode === ColorMode.CMYK) {
                try { cyanInput.active = true; } catch (e) { }
            }
            updatePreviewImmediately();
        }

        noneRadio.onClick = noneRadio.onChanging = function () { selectColorMode(ColorMode.NONE); };
        k100Radio.onClick = k100Radio.onChanging = function () { selectColorMode(ColorMode.K100); };
        hexRadio.onClick = hexRadio.onChanging = function () { selectColorMode(ColorMode.HEX); };
        cmykRadio.onClick = cmykRadio.onChanging = function () { selectColorMode(ColorMode.CMYK); };

        // --- Hotkeys: N/K/H/C to switch color mode radios ---
        // --- Hotkeys: Z-order (F/B/L), Make Guides (G), Target scope (C/A) ---
        /* 重ね順・ガイド化・対象範囲をホットキーで切替 / Toggle z-order, make-guides & target scope via hotkeys */
        function addDialogHotkeys(dialog) {
            dialog.addEventListener('keydown', function (event) {
                if (hotkeyState.v) return; // ignore when typing in fields
                var key = (event && event.keyName) ? String(event.keyName).toUpperCase() : '';

                // Z-order: F = Front, B = Back, L = bg Layer
                if (key === 'F') {
                    frontRadio.value = true;
                    updatePreviewImmediately();
                    event.preventDefault();
                    return;
                }
                if (key === 'B') {
                    backRadio.value = true;
                    updatePreviewImmediately();
                    event.preventDefault();
                    return;
                }
                if (key === 'L') {
                    bgLayerRadio.value = true;
                    updatePreviewImmediately();
                    event.preventDefault();
                    return;
                }

                // Make guides: G toggles the checkbox (post-draw option; not previewed)
                if (key === 'G') {
                    makeGuideCheckbox.value = !makeGuideCheckbox.value;
                    event.preventDefault();
                    return;
                }

                // Target scope: C = current artboard, A = all artboards
                if (key === 'C') {
                    currentArtboardRadio.value = true;
                    updatePreviewImmediately();
                    event.preventDefault();
                    return;
                }
                if (key === 'A') {
                    if (allArtboardsRadio.enabled) {
                        allArtboardsRadio.value = true;
                        updatePreviewImmediately();
                    }
                    event.preventDefault();
                    return;
                }
            });
        }
        addDialogHotkeys(dialog);

        frontRadio.onClick = updatePreviewImmediately;
        backRadio.onClick = updatePreviewImmediately;
        bgLayerRadio.onClick = updatePreviewImmediately;

        currentArtboardRadio.onClick = updatePreviewImmediately;
        allArtboardsRadio.onClick = updatePreviewImmediately;
        currentArtboardRadio.onChanging = updatePreviewImmediately;
        allArtboardsRadio.onChanging = updatePreviewImmediately;

        bleedCheckbox.onClick = function () {
            if (bleedCheckbox.value) {
                lastUserOffsetText = String(offsetInput.text);
                applyBleedPreset(true);
            } else {
                removeBleedPreset(true);
            }
        };

        var buttonRowGroup = dialog.add('group');
        buttonRowGroup.orientation = 'row';
        buttonRowGroup.alignChildren = ['fill', 'center'];
        buttonRowGroup.alignment = 'fill';

        var leftButtonGroup = buttonRowGroup.add('group');
        leftButtonGroup.orientation = 'row';

        var isPreviewDisplay = true;
        var previewButton = leftButtonGroup.add('button', undefined, getLocalizedText('button.previewOutline'));
        previewButton.helpTip = getLocalizedText('helpTip.previewToggle');

        var spacer = buttonRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];

        var rightButtonGroup = buttonRowGroup.add('group');
        rightButtonGroup.orientation = 'row';
        rightButtonGroup.alignment = ['right', 'center']; // 右カラムは右揃え / right-align the right column

        var cancelButton = rightButtonGroup.add('button', undefined, getLocalizedText('button.cancel'));
        var okButton = rightButtonGroup.add('button', undefined, getLocalizedText('button.ok'));

        previewButton.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                isPreviewDisplay = !isPreviewDisplay;
                previewButton.text = isPreviewDisplay ? getLocalizedText('button.previewOutline') : getLocalizedText('button.previewPreview');
            } catch (e) { }
        };

        okButton.onClick = function () {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
            PreviewHistory.undo();
            dialog.close(1);
        };
        cancelButton.onClick = function () {
            if (__previewDebounceTask) PreviewHistory.cancelTask(__previewDebounceTask);
            PreviewHistory.undo();
            dialog.close(0);
        };

        var result = dialog.show();
        if (result != 1) {
            return null;
        }

        // 確定値は buildDialogSettings() に一元化（プレビューと同じ計算）
        // Final values come from buildDialogSettings() (same computation as the preview)
        return buildDialogSettings();
    }

    /*
     * 編集可能なレイヤーを取得（なければ作成）/ Get an editable layer or create one
     */
    function getWritableLayer(doc) {
        try {
            var lyr = doc.activeLayer;
            if (lyr && !lyr.locked && lyr.visible) return lyr;
        } catch (e) { }
        // try find first unlocked & visible layer
        try {
            for (var i = 0; i < doc.layers.length; i++) {
                var l = doc.layers[i];
                if (!l.locked && l.visible) return l;
            }
        } catch (e) { }
        // last resort: create a new layer at top
        try {
            var nl = doc.layers.add();
            nl.name = "_auto_draw";
            nl.visible = true;
            nl.locked = false;
            return nl;
        } catch (e) { }
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
        } catch (e) { }
        // 最背面へ / Send to back of layer stack
        try {
            layer.move(doc, ElementPlacement.PLACEATEND);
        } catch (e) { }
        return layer;
    }

    /* アートボードと同サイズ（オフセット込み）の長方形を1枚描画して返す
     * Draw one rectangle matching the artboard (incl. offset) and return it
     */
    function drawRectangleForArtboard(doc, artboard, choice) {
        var artboardRect = artboard.artboardRect; // [left, top, right, bottom]

        var artboardWidth = artboardRect[2] - artboardRect[0];
        var artboardHeight = artboardRect[1] - artboardRect[3];

        var o = choice.offset;
        var targetLayer = (choice.zOrder === 'bg') ? getOrCreateBgLayer(doc) : getWritableLayer(doc);
        // Create on the chosen layer to avoid "Target layer cannot be modified"
        var artboardRectangle = targetLayer.pathItems.rectangle(
            artboardRect[1] + o,
            artboardRect[0] - o,
            artboardWidth + o * 2,
            artboardHeight + o * 2
        );

        // Unified fill application (final draw)
        applyFillByMode(doc, artboardRectangle, choice.colorMode, {
            customValue: choice.customValue,
            customCMYK: choice.customCMYK
        }, {
            k100Opacity: 15
        });

        artboardRectangle.name = getLocalizedText('name.rect');
        artboardRectangle.selected = true;
        try {
            artboardRectangle.hidden = false;
        } catch (e) { }
        try {
            targetLayer.visible = true;
            targetLayer.locked = false;
        } catch (e) { }

        if (choice.zOrder === 'front') {
            artboardRectangle.zOrder(ZOrderMethod.BRINGTOFRONT);
        } else if (choice.zOrder === 'back') {
            artboardRectangle.zOrder(ZOrderMethod.SENDTOBACK);
        }

        return artboardRectangle;
    }

    /* 中心の○（属性パネル「中心点を表示」）を選択オブジェクトへ適用 / Show the center widget on selected items
     * API・メニューコマンドからは設定できないため、記録済みアクション(.aia)を一時ファイルへ書き出して
     * loadAction→doScript で再生する（メモリ: reference_illustrator_temp_action）。
     * セット名 "object" / アクション名 "ヘソ表示"（日本語名は .aia 本文をそのまま使う）。
     * 呼び出し側で「対象だけを選択した状態」にしてから実行すること。
     */
    function showShapeCenterWidget() {
        var ACTION_SET_NAME = 'object';
        var ACTION_NAME = 'ヘソ表示';
        var ACTION_BODY = [
            '/version 3',
            '/name [ 6',
            '\t6f626a656374',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '\t/name [ 12',
            '\t\te38398e382bde8a1a8e7a4ba',
            '\t]',
            '\t/keyIndex 0',
            '\t/colorIndex 0',
            '\t/isOpen 1',
            '\t/eventCount 1',
            '\t/event-1 {',
            '\t\t/useRulersIn1stQuadrant 0',
            '\t\t/internalName (adobe_attributePalette)',
            '\t\t/localizedName [ 12',
            '\t\t\te5b19ee680a7e8a8ade5ae9a',
            '\t\t]',
            '\t\t/isOpen 1',
            '\t\t/isOn 1',
            '\t\t/hasDialog 0',
            '\t\t/parameterCount 1',
            '\t\t/parameter-1 {',
            '\t\t\t/key 1668183154',
            '\t\t\t/showInPalette 4294967295',
            '\t\t\t/type (boolean)',
            '\t\t\t/value 1',
            '\t\t}',
            '\t}',
            '}',
            ''
        ].join('\n');

        var tmpFile = null;
        try {
            // 一時ファイルへ書き出し / write the recorded action to a temp file
            tmpFile = new File(Folder.temp + '/SmartDrawArtboardRectangle_center.aia');
            tmpFile.encoding = 'UTF-8';
            tmpFile.open('w');
            tmpFile.write(ACTION_BODY);
            tmpFile.close();

            // 同名セットを解放してからロード→再生 / unload any same-named set, then load & play
            try { app.unloadAction(ACTION_SET_NAME, ''); } catch (e) { }
            app.loadAction(tmpFile);
            app.doScript(ACTION_NAME, ACTION_SET_NAME, false);
        } catch (e) {
        } finally {
            try { app.unloadAction(ACTION_SET_NAME, ''); } catch (e) { }
            try { if (tmpFile && tmpFile.exists) tmpFile.remove(); } catch (e) { }
        }
    }

    /* 描画後のオプション（ガイド化／ライブシェイプ変換／中心の○表示）を選択中の長方形へ適用
     * Apply post-draw options (make guides / convert to live shape / show center) to the drawn rectangles
     */
    function applyDrawOptions(createdRectangles, choice) {
        if (!createdRectangles || !createdRectangles.length) return;

        // ライブシェイプに変換：ONのときだけ選択ベースのメニューコマンドを実行
        // Convert to Live Shape: run the selection-based menu command only when ON
        if (choice.convertToLiveShape) {
            try { app.executeMenuCommand('deselectall'); } catch (e) { }
            for (var j = 0; j < createdRectangles.length; j++) {
                try { createdRectangles[j].selected = true; } catch (e) { }
            }
            try { app.executeMenuCommand('Convert to Shape'); } catch (e) { }
        }

        // 中心の○を表示（常にON・UI非掲載）：新規長方形だけを選択し直してからアクション再生
        // Show center widget (always ON, no UI): re-select only the new rectangles, then replay the action
        try { app.executeMenuCommand('deselectall'); } catch (e) { }
        for (var c = 0; c < createdRectangles.length; c++) {
            try { createdRectangles[c].selected = true; } catch (e) { }
        }
        showShapeCenterWidget();

        // ガイド化：PathItem.guides を直接立てる（選択・メニュー状態に依存せず確実）
        // Make guides: set PathItem.guides directly — robust, independent of selection/menu state
        if (choice.makeGuide) {
            for (var g = 0; g < createdRectangles.length; g++) {
                try { createdRectangles[g].guides = true; } catch (e) { }
            }
        }
    }

    /* エントリポイント：ダイアログ→描画→オプション適用 / Entry point: dialog → draw → options */
    function main() {
        if (app.documents.length === 0) return;

        var choice = showDialog();
        if (choice === null) return;

        var doc = app.activeDocument;

        var previousCoordinateSystem = null;
        try {
            previousCoordinateSystem = app.coordinateSystem;
        } catch (e) { }
        try {
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }

        app.executeMenuCommand('deselectall'); // 既存選択を解除

        var createdRectangles = [];
        if (choice.target === 'current') {
            var currentArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            createdRectangles.push(drawRectangleForArtboard(doc, currentArtboard, choice));
        } else if (choice.target === 'all') {
            var prevIndex = doc.artboards.getActiveArtboardIndex();
            for (var i = 0; i < doc.artboards.length; i++) {
                try {
                    doc.artboards.setActiveArtboardIndex(i);
                } catch (e) { }
                var targetArtboard = doc.artboards[i];
                createdRectangles.push(drawRectangleForArtboard(doc, targetArtboard, choice));
            }
            try {
                doc.artboards.setActiveArtboardIndex(prevIndex);
            } catch (e) { }
        }

        applyDrawOptions(createdRectangles, choice);

        try {
            if (previousCoordinateSystem !== null) app.coordinateSystem = previousCoordinateSystem;
        } catch (e) { }
    }

    main();

})();