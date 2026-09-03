#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブまたはすべてのアートボードと同じサイズの長方形を、オフセットを考慮して描画します。
カラー・配置位置・対象をライブプレビューで確かめながら指定でき、描画後に「ガイドに変換」「ライブシェイプ化」を適用できます。

詳細は README を参照してください。

### Overview

Draws a rectangle the size of the active artboard, or of every artboard, taking an offset into account.
Color, placement and target scope are set with a live preview, and the result can be converted to guides or to a live shape.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartDrawArtboardRectangle";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDrawArtboardRectangle.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDrawArtboardRectangle.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1ba88513a9c8"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 入力中のプレビュー遅延（タイプしやすさ優先）/ Preview delay while typing */
    var PREVIEW_DELAY_TYPING_MS = 110; /* 推奨 100–120ms / recommend 100–120ms */

    /* K100モードの不透明度（%）/ Opacity (%) used by the K100 color mode */
    var K100_OPACITY = 15;

    /* 「bgレイヤー」配置で使うレイヤー名 / Layer name used by the "bg layer" placement */
    var BG_LAYER_NAME = 'bg';

    /* 描画先が見つからないときに作るレイヤー名 / Layer created when no writable layer is found */
    var FALLBACK_LAYER_NAME = '_auto_draw';

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログの初期位置・不透明度 / Dialog position & opacity */
    var DIALOG_OFFSET_X = 300;  /* 右(+)／左(-) / shift right (+) / left (-) */
    var DIALOG_OFFSET_Y = 0;    /* 下(+)／上(-) / shift down (+) / up (-) */
    var DIALOG_OPACITY = 0.98;  /* 0.0 - 1.0 */

    /* 余白と間隔 / Margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING = 8;                /* パネル内の要素間隔 */
    var COLUMN_SPACING = 12;              /* 2カラムの間隔 */
    var STACK_SPACING = 10;               /* カラム内のパネル間隔・広めの行間 */
    var TIGHT_SPACING = 6;                /* 詰めた行間 */

    /* CMYK入力欄の固定幅（ラベルと桁を揃える）/ Fixed width that aligns CMYK labels and fields */
    var CMYK_FIELD_WIDTH = 40;

    /**
     * パネルの共通設定
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループの共通設定（row/column で整列を切り替え）
     * @param {Group} group - 対象グループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // カラーモード / Color modes
    // =========================================

    /* 塗りの決め方を表す定数 / How the fill color is decided */
    var ColorMode = {
        NONE: 'none',
        K100: 'k100',
        HEX: 'hex',
        CMYK: 'cmyk'
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* ラベル定義（カテゴリ別）/ Label definitions (by category) */
    var LABELS = {
        dialog: {
            title: {
                ja: "アートボードサイズの長方形を描画",
                en: "Draw Artboard Size Rectangle"
            }
        },
        panel: {
            offset: { ja: "オフセット", en: "Offset" },
            color: { ja: "カラー", en: "Color" },
            zorder: { ja: "配置位置", en: "Placement" },
            target: { ja: "対象", en: "Target" },
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            bleed: { ja: "裁ち落とし", en: "Bleed" },
            makeGuide: { ja: "ガイドに変換", en: "Convert to Guides" },
            convertToLiveShape: { ja: "ライブシェイプ化", en: "Convert to Live Shape" }
        },
        color: {
            none: { ja: "なし", en: "None" },
            k100: { ja: "K100、不透明度15%", en: "K100, Opacity 15%" },
            hex: { ja: "HEX", en: "HEX" },
            cmyk: { ja: "CMYK", en: "CMYK" },
            hint: { ja: "例: #FF0000", en: "e.g., #FF0000" }
        },
        zorder: {
            front: { ja: "最前面", en: "Front" },
            back: { ja: "最背面", en: "Back" },
            bg: { ja: "bgレイヤー", en: "bg Layer" }
        },
        target: {
            current: { ja: "現在のアートボード", en: "Current Artboard" },
            all: { ja: "すべてのアートボード", en: "All Artboards" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            previewOutline: { ja: "アウトライン表示", en: "Outline" },
            previewPreview: { ja: "プレビュー表示", en: "Preview" }
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
                ja: "#RRGGBB（#RGB 短縮・red などの色名・gray50 も可）で塗りカラーを指定します。",
                en: "Enter a fill color: #RRGGBB (also #RGB shorthand, color names like red, or gray50)."
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
        warning: {
            hexInvalid: {
                ja: "正しい #RRGGBB を入力してください",
                en: "Enter a valid #RRGGBB value"
            },
            hexEmpty: {
                ja: "HEX未入力（# のみ）",
                en: "HEX not entered (# only)"
            },
            cmykRange: {
                ja: "0–100 の範囲にしてください（未入力は 0 として扱います）",
                en: "Enter a value from 0 to 100 (empty fields are treated as 0)"
            },
            singleArtboard: {
                ja: "アートボードが1つのため選択できません",
                en: "Disabled: only one artboard exists"
            }
        },
        name: {
            previewLayer: { ja: "_preview", en: "_preview" },
            rect: { ja: "<長方形>", en: "<Rectangle>" },
            guide: { ja: "<ガイド>", en: "<Guide>" },
            previewRect: {
                ja: "__プレビュー_アートボードサイズの長方形",
                en: "__Preview_ArtboardSizeRectangle"
            }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー、{slash}→「/」に展開）
     * @param {string} key - "panel.offset" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(key) {
        var labelNode = LABELS;
        var keyParts = String(key).split('.');
        for (var i = 0; i < keyParts.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[keyParts[i]];
        }
        var text = (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : key;
        return String(text).replace(/\{slash\}/g, '/');
    }

    // =========================================
    // ダイアログ共通ユーティリティ / Dialog utilities
    // =========================================

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
                setOpacity: function (dialog, opacity) {
                    try { dialog.opacity = opacity; } catch (e) { }
                },
                applyInitialOffset: function (dialog, offsetX, offsetY) {
                    try {
                        var location = dialog.location;
                        dialog.location = [location[0] + (offsetX | 0), location[1] + (offsetY | 0)];
                    } catch (e) { }
                }
            };
        }
    })($.global);

    /* 入力中のホットキー抑止用に、フォーカス中のコントロールを保持 / Control that currently owns focus */
    var focusedField = null;

    /**
     * 入力欄のフォーカスを追跡し、入力中はダイアログのホットキーを無効にする
     * 単一 boolean だと「新フィールドの focus → 旧フィールドの blur」の順で false に落ちるため、
     * コントロール自体を保持して自分の blur のときだけクリアする（順序非依存）。
     * @param {EditText} fieldControl - 追跡対象の入力欄
     * @returns {void}
     */
    function trackFocusForHotkeys(fieldControl) {
        fieldControl.addEventListener('focus', function () {
            focusedField = fieldControl;
        });
        fieldControl.addEventListener('blur', function () {
            if (focusedField === fieldControl) focusedField = null;
        });
    }

    /**
     * 入力欄を淡黄色でハイライト表示する（選択中のカラーモードを示す）
     * @param {EditText} fieldControl - 対象の入力欄
     * @param {boolean} highlighted - true でハイライト、false で通常表示
     * @returns {void}
     */
    function setFieldHighlight(fieldControl, highlighted) {
        try {
            var graphics = fieldControl.graphics;
            var backgroundRgb = highlighted ? [1, 1, 0.85] : [1, 1, 1];
            var foregroundRgb = highlighted ? [0.2, 0.2, 0] : [0, 0, 0];
            graphics.backgroundColor = graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundRgb);
            graphics.foregroundColor = graphics.newPen(graphics.PenType.SOLID_COLOR, foregroundRgb, 1);
            fieldControl.notify('onDraw');
        } catch (e) { }
    }

    /**
     * 入力欄の文字色を警告(赤)／通常(黒)に切り替える
     * @param {EditText} fieldControl - 対象の入力欄
     * @param {boolean} isWarning - true で赤、false で黒
     * @returns {void}
     */
    function setFieldWarnColor(fieldControl, isWarning) {
        try {
            var graphics = fieldControl.graphics;
            graphics.foregroundColor = graphics.newPen(graphics.PenType.SOLID_COLOR, isWarning ? [1, 0, 0] : [0, 0, 0], 1);
        } catch (e) { }
    }

    /**
     * 上下キーで入力値を増減する（Shift=10刻み、Option=0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {function} onValueChange - 値が変わったあとに呼ぶコールバック
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onValueChange) {
        editText.addEventListener("keydown", function (event) {
            /* 上下キー以外（通常の文字入力など）には一切干渉しない
               Only react to Up/Down; never touch the field on normal character input */
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) value = 0; /* 空欄などは0起点 / treat empty or invalid as 0 */

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            event.preventDefault();

            if (keyboard.shiftKey) {
                /* Shift押下時は10の倍数へスナップ / Snap to multiples of 10 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                /* Option押下時は0.1単位、小数第1位に丸め / 0.1 steps, rounded to one decimal */
                value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                value = Math.round(value + (isUp ? 1 : -1));
            }

            editText.text = value;
            if (typeof onValueChange === 'function') onValueChange();
        });
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コード→ラベルとpt係数のテーブル（rulerType基準）
       Map rulerType codes to label & points-per-unit factor */
    var UNIT_TABLE = {
        0: { label: "in", factor: 72.0 },                 /* inch */
        1: { label: "mm", factor: 72.0 / 25.4 },          /* mm */
        2: { label: "pt", factor: 1.0 },                  /* pt */
        3: { label: "pica", factor: 12.0 },               /* pica */
        4: { label: "cm", factor: 72.0 / 2.54 },          /* cm */
        5: { label: "Q/H", factor: 72.0 / 25.4 * 0.25 },  /* Q or H */
        6: { label: "px", factor: 1.0 },                  /* px（Illustratorでは 1px=1pt）*/
        7: { label: "ft/in", factor: 72.0 * 12.0 },       /* ft/in */
        8: { label: "m", factor: 72.0 / 25.4 * 1000.0 },  /* m */
        9: { label: "yd", factor: 72.0 * 36.0 },          /* yd */
        10: { label: "ft", factor: 72.0 * 12.0 }          /* ft */
    };

    /**
     * 現在の単位コード（rulerType）を取得する
     * @returns {number} 単位コード（取得できない場合は 2 = pt）
     */
    function getCurrentUnitCode() {
        try {
            return app.preferences.getIntegerPreference("rulerType");
        } catch (e) {
            return 2;
        }
    }

    /**
     * 現在の単位ラベルを取得する
     * @returns {string} "mm" などの単位ラベル
     */
    function getCurrentUnitLabel() {
        var unitEntry = UNIT_TABLE[getCurrentUnitCode()];
        return unitEntry ? unitEntry.label : "pt";
    }

    /**
     * 単位コードからpt換算係数を取得する
     * @param {number} unitCode - rulerType の単位コード
     * @returns {number} 1単位あたりのpt数
     */
    function getPtFactorFromUnitCode(unitCode) {
        var unitEntry = UNIT_TABLE[unitCode];
        return unitEntry ? unitEntry.factor : 1.0;
    }

    /**
     * 入力欄の表示値と内部pt値を一元的に解決する（裁ち落としプリセットを含む）
     * @param {string} offsetText - 入力欄の現在のテキスト
     * @param {number} unitCode - rulerType の単位コード
     * @param {boolean} bleedEnabled - 裁ち落としがONかどうか
     * @returns {object} { pt: number, displayText: string, disabled: boolean }
     */
    function resolveOffsetToPt(offsetText, unitCode, bleedEnabled) {
        var displayText = String(offsetText == null ? '' : offsetText);

        if (bleedEnabled) {
            /* 単位ごとの裁ち落とし相当値。表示値と pt 値を必ず同じ量にする
               Bleed preset per unit; the shown value and the pt value always describe the same amount */
            var bleedAmount = 3;      /* 既定は 3mm 相当 / defaults to 3mm */
            var bleedUnitCode = 1;
            if (unitCode === 5) {          /* Q/H */
                bleedAmount = 12;
                bleedUnitCode = 5;
            } else if (unitCode === 2) {   /* pt（0.125in = 9pt）*/
                bleedAmount = 9;
                bleedUnitCode = 2;
            } else if (unitCode === 1) {   /* mm */
                bleedAmount = 3;
                bleedUnitCode = 1;
            } else {
                /* mm・Q/H・pt 以外は 3mm 相当を現在の単位へ換算して表示
                   For other units, convert the 3mm equivalent into the current unit */
                bleedAmount = 3 * getPtFactorFromUnitCode(1) / getPtFactorFromUnitCode(unitCode);
                bleedAmount = Math.round(bleedAmount * 1000) / 1000;
                bleedUnitCode = unitCode;
            }
            return {
                pt: bleedAmount * getPtFactorFromUnitCode(bleedUnitCode),
                displayText: String(bleedAmount),
                disabled: true
            };
        }

        /* 通常時は現在の単位の係数を掛ける / Normal case: multiply by the current unit factor */
        var offsetValue = parseFloat(displayText);
        if (isNaN(offsetValue)) offsetValue = 0;
        return {
            pt: offsetValue * getPtFactorFromUnitCode(unitCode),
            displayText: displayText,
            disabled: false
        };
    }

    // =========================================
    // カラー / Color
    // =========================================

    /* 色名テーブル（RGB/CMYK 両方を持つ）/ Named colors, with both an RGB and a CMYK value */
    var NAMED_COLOR_TABLE = {
        black: { rgb: [0, 0, 0], cmyk: [0, 0, 0, 100] },
        white: { rgb: [255, 255, 255], cmyk: [0, 0, 0, 0] },
        red: { rgb: [255, 0, 0], cmyk: [0, 100, 100, 0] },
        green: { rgb: [0, 128, 0], cmyk: [100, 0, 100, 50] },
        blue: { rgb: [0, 0, 255], cmyk: [100, 100, 0, 0] },
        cyan: { rgb: [0, 255, 255], cmyk: [100, 0, 0, 0] },
        magenta: { rgb: [255, 0, 255], cmyk: [0, 100, 0, 0] },
        yellow: { rgb: [255, 255, 0], cmyk: [0, 0, 100, 0] },
        orange: { rgb: [255, 165, 0], cmyk: [0, 35, 100, 0] }
    };

    /**
     * 数値を指定範囲に収める
     * @param {number} value - 対象の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clampValue(value, minValue, maxValue) {
        return value < minValue ? minValue : (value > maxValue ? maxValue : value);
    }

    /**
     * RGBColor を生成する（0–255にクランプ）
     * @param {number} red - 赤（0–255）
     * @param {number} green - 緑（0–255）
     * @param {number} blue - 青（0–255）
     * @returns {RGBColor} 生成した色
     */
    function makeRgbColor(red, green, blue) {
        var rgbColor = new RGBColor();
        rgbColor.red = clampValue(Math.round(red), 0, 255);
        rgbColor.green = clampValue(Math.round(green), 0, 255);
        rgbColor.blue = clampValue(Math.round(blue), 0, 255);
        return rgbColor;
    }

    /**
     * CMYKColor を生成する（0–100にクランプ）
     * @param {number} cyan - シアン（0–100）
     * @param {number} magenta - マゼンタ（0–100）
     * @param {number} yellow - イエロー（0–100）
     * @param {number} black - ブラック（0–100）
     * @returns {CMYKColor} 生成した色
     */
    function makeCmykColor(cyan, magenta, yellow, black) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = clampValue(cyan, 0, 100);
        cmykColor.magenta = clampValue(magenta, 0, 100);
        cmykColor.yellow = clampValue(yellow, 0, 100);
        cmykColor.black = clampValue(black, 0, 100);
        return cmykColor;
    }

    /**
     * CMYK値をRGB値へ変換する
     * @param {number} cyan - シアン（0–100）
     * @param {number} magenta - マゼンタ（0–100）
     * @param {number} yellow - イエロー（0–100）
     * @param {number} black - ブラック（0–100）
     * @returns {number[]} [R, G, B]（0–255）
     */
    function cmykToRgb(cyan, magenta, yellow, black) {
        var c = clampValue(cyan, 0, 100) / 100;
        var m = clampValue(magenta, 0, 100) / 100;
        var y = clampValue(yellow, 0, 100) / 100;
        var k = clampValue(black, 0, 100) / 100;
        return [
            Math.round(255 * (1 - c) * (1 - k)),
            Math.round(255 * (1 - m) * (1 - k)),
            Math.round(255 * (1 - y) * (1 - k))
        ];
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を生成する
     * @param {Document} doc - 対象ドキュメント
     * @returns {RGBColor|CMYKColor} 黒
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) return makeRgbColor(0, 0, 0);
        return makeCmykColor(0, 0, 0, 100);
    }

    /**
     * カラー入力欄の文字列を色として解釈する
     * #RRGGBB／#RGB・#RG・#R の短縮形／色名（red など）／grayNN（0–100）を受け付ける
     * @param {Document} doc - 対象ドキュメント（カラースペース判定に使う）
     * @param {string} colorText - 入力文字列
     * @returns {RGBColor|CMYKColor|null} 解釈できた色。できなければ null
     */
    function parseColorText(doc, colorText) {
        if (!colorText) return null;
        var text = String(colorText).replace(/^\s+|\s+$/g, '').toLowerCase();
        if (!text) return null;

        /* 全角の空白・読点・数字・記号をASCIIへ正規化 / Normalize full-width characters to ASCII */
        text = text.replace(/　/g, ' ').replace(/[，、]/g, ',');
        text = text.replace(/[０-９]/g, function (ch) {
            return String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30);
        });
        text = text.replace(/．/g, '.').replace(/／/g, '/');

        /* 短縮HEXを #RRGGBB へ展開 / Expand shorthand hex notations */
        if (text.charAt(0) === '#') {
            var digits = text.substr(1);
            if (digits.length === 1) {          /* #R → #RRRRRR */
                text = '#' + digits + digits + digits + digits + digits + digits;
            } else if (digits.length === 2) {   /* #RG → #RGRGRG */
                text = '#' + digits + digits + digits;
            } else if (digits.length === 3) {   /* #RGB → #RRGGBB */
                text = '#' + digits.charAt(0) + digits.charAt(0) +
                    digits.charAt(1) + digits.charAt(1) +
                    digits.charAt(2) + digits.charAt(2);
            }
        }

        /* #RRGGBB */
        if (/^#[0-9a-f]{6}$/.test(text)) {
            return makeRgbColor(parseInt(text.substr(1, 2), 16), parseInt(text.substr(3, 2), 16), parseInt(text.substr(5, 2), 16));
        }

        /* 色名 / Named colors — ドキュメントのカラースペースを優先 */
        var namedColor = NAMED_COLOR_TABLE[text];
        if (namedColor) {
            if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK) {
                return makeCmykColor(namedColor.cmyk[0], namedColor.cmyk[1], namedColor.cmyk[2], namedColor.cmyk[3]);
            }
            return makeRgbColor(namedColor.rgb[0], namedColor.rgb[1], namedColor.rgb[2]);
        }

        /* grayNN（0–100）/ grayNN (0-100) */
        var grayMatch = text.match(/^gray\s*(\d{1,3})$/);
        if (grayMatch) {
            var grayLevel = clampValue(parseInt(grayMatch[1], 10), 0, 100);
            if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK) return makeCmykColor(0, 0, 0, grayLevel);
            var grayByte = Math.round(255 * (100 - grayLevel) / 100);
            return makeRgbColor(grayByte, grayByte, grayByte);
        }

        return null;
    }

    /**
     * CMYK入力値からドキュメントのカラースペースに合う色を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {object} cmykValues - { c, m, y, k }（0–100）
     * @returns {RGBColor|CMYKColor|null} 4値が揃っていなければ null
     */
    function buildCmykFillColor(doc, cmykValues) {
        if (!cmykValues) return null;
        var channels = [cmykValues.c, cmykValues.m, cmykValues.y, cmykValues.k];
        for (var i = 0; i < channels.length; i++) {
            if (typeof channels[i] !== 'number' || isNaN(channels[i])) return null;
        }
        if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
            var rgb = cmykToRgb(channels[0], channels[1], channels[2], channels[3]);
            return makeRgbColor(rgb[0], rgb[1], rgb[2]);
        }
        return makeCmykColor(channels[0], channels[1], channels[2], channels[3]);
    }

    /**
     * カラーモードに応じた塗りを適用する（プレビューと本描画で共通）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} targetRectangle - 塗りを適用する長方形
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function applyFillByMode(doc, targetRectangle, drawSettings) {
        var fillColor = null;
        var fillOpacity = 100;

        if (drawSettings.colorMode === ColorMode.K100) {
            fillColor = createBlackColor(doc);
            fillOpacity = K100_OPACITY;
        } else if (drawSettings.colorMode === ColorMode.HEX) {
            fillColor = parseColorText(doc, drawSettings.customValue);
        } else if (drawSettings.colorMode === ColorMode.CMYK) {
            fillColor = buildCmykFillColor(doc, drawSettings.customCMYK);
        }

        /* 解釈できない値・「なし」は塗りなし。線は呼び出し側（プレビュー）で付け直す
           Unparsable values and "None" mean no fill; the caller re-applies any stroke */
        targetRectangle.stroked = false;
        targetRectangle.filled = !!fillColor;
        if (fillColor) targetRectangle.fillColor = fillColor;
        targetRectangle.opacity = fillOpacity;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 使い方:
     *   PreviewHistory.start();      // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();       // プレビュー描画ごとにカウント(+1)
     *   PreviewHistory.undo();       // 閉じる/キャンセル時に一括Undo
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
                    var undoCount = g.__previewUndoCount | 0;
                    try {
                        for (var i = 0; i < undoCount; i++) app.executeMenuCommand('undo');
                    } catch (e) { }
                    g.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

    /* デバウンス中のプレビュータスクID / Task id of the pending debounced preview */
    var previewDebounceTaskId = null;

    /**
     * プレビュー描画を遅延スケジュールする（デバウンス）
     * @param {object} drawSettings - ダイアログの設定値
     * @param {number} delayMs - 遅延ミリ秒
     * @returns {void}
     */
    function schedulePreview(drawSettings, delayMs) {
        PreviewHistory.cancelTask(previewDebounceTaskId);
        /* scheduleTask の文字列はグローバルスコープで評価されるため、IIFE内の関数を $.global 経由で渡す
           scheduleTask runs its string in global scope, so expose the renderer via $.global */
        $.global.__previewSettings = drawSettings;
        $.global.__previewRenderer = renderPreview;
        var scheduledCode = 'try{$.global.__previewRenderer(app.activeDocument, $.global.__previewSettings);}catch(e){}';
        try {
            previewDebounceTaskId = app.scheduleTask(scheduledCode, Math.max(0, delayMs | 0), false);
        } catch (e) {
            try { renderPreview(app.activeDocument, drawSettings); } catch (err) { }
        }
    }

    /**
     * プレビュー専用レイヤーの名前かどうかを判定する
     * @param {string} layerName - レイヤー名
     * @returns {boolean} プレビュー専用レイヤーなら true
     */
    function isPreviewLayerName(layerName) {
        return layerName === getLabel('name.previewLayer') || layerName === '_preview';
    }

    /**
     * プレビューを片付ける
     * @param {boolean} removeLayer - true でレイヤーごと削除、false ではプレビュー長方形を隠すだけ
     * @returns {void}
     */
    function clearPreview(removeLayer) {
        try {
            var doc = app.activeDocument;
            var previewItemPrefix = getLabel('name.previewRect') + "#";
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var layer = doc.layers[i];
                if (!isPreviewLayerName(layer.name)) continue;
                if (removeLayer) {
                    layer.remove();
                    continue;
                }
                /* 入力中は削除せず、このスクリプトが作った長方形だけ隠す
                   While typing, hide only the items this script created instead of deleting them */
                for (var k = layer.pathItems.length - 1; k >= 0; k--) {
                    if (String(layer.pathItems[k].name || "").indexOf(previewItemPrefix) === 0) {
                        layer.pathItems[k].hidden = true;
                    }
                }
            }
        } catch (e) { }
    }

    /**
     * プレビュー専用レイヤーを取得する（なければ作成し最前面へ）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} プレビュー専用レイヤー
     */
    function getOrCreatePreviewLayer(doc) {
        var previewLayerName = getLabel('name.previewLayer');
        var previewLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === previewLayerName) {
                previewLayer = doc.layers[i];
                break;
            }
        }
        if (!previewLayer) {
            previewLayer = doc.layers.add();
            previewLayer.name = previewLayerName;
        }
        previewLayer.visible = true;
        previewLayer.locked = false;
        try {
            previewLayer.move(doc, ElementPlacement.PLACEATBEGINNING);
        } catch (e) { }
        return previewLayer;
    }

    /**
     * アートボード番号に対応するプレビュー長方形を取得する（なければ作成、あれば再利用）
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {number} artboardIndex - アートボード番号
     * @param {number} top - 上端座標
     * @param {number} left - 左端座標
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @returns {PathItem} プレビュー長方形
     */
    function getOrCreatePreviewRectangle(previewLayer, artboardIndex, top, left, width, height) {
        var previewItemName = getLabel('name.previewRect') + "#" + artboardIndex;
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var existingRectangle = previewLayer.pathItems[i];
            if (existingRectangle.name !== previewItemName) continue;
            /* 既存を使い回して再作成のヒストリーを増やさない / Reuse in place to avoid extra history entries */
            existingRectangle.top = top;
            existingRectangle.left = left;
            existingRectangle.width = width;
            existingRectangle.height = height;
            existingRectangle.hidden = false;
            return existingRectangle;
        }
        var previewRectangle = previewLayer.pathItems.rectangle(top, left, width, height);
        previewRectangle.name = previewItemName;
        return previewRectangle;
    }

    /**
     * プレビュー用の50%グレー（ドキュメントのカラースペースに合わせる）
     * @param {Document} doc - 対象ドキュメント
     * @returns {RGBColor|CMYKColor} 線色
     */
    function getPreviewStrokeColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) return makeRgbColor(128, 128, 128);
        return makeCmykColor(0, 0, 0, 50);
    }

    /**
     * 1つのアートボードぶんのプレビュー長方形を描く
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {number} artboardIndex - アートボード番号
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function drawPreviewRectangle(doc, previewLayer, artboardIndex, drawSettings) {
        var artboardRect = doc.artboards[artboardIndex].artboardRect; /* [left, top, right, bottom] */
        var artboardWidth = artboardRect[2] - artboardRect[0];
        var artboardHeight = artboardRect[1] - artboardRect[3];
        var offsetPt = drawSettings.offset || 0;

        var previewRectangle = getOrCreatePreviewRectangle(
            previewLayer,
            artboardIndex,
            artboardRect[1] + offsetPt,
            artboardRect[0] - offsetPt,
            artboardWidth + offsetPt * 2,
            artboardHeight + offsetPt * 2
        );

        applyFillByMode(doc, previewRectangle, drawSettings);

        /* 塗りなしでも位置が分かるよう、プレビューは常に破線で縁取る
           Always outline the preview so it stays visible even with no fill */
        previewRectangle.stroked = true;
        previewRectangle.strokeWidth = 1;
        previewRectangle.strokeDashes = [6, 4];
        previewRectangle.strokeColor = getPreviewStrokeColor(doc);
        previewRectangle.selected = false;

        if (drawSettings.zOrder === 'front') previewRectangle.zOrder(ZOrderMethod.BRINGTOFRONT);
        else if (drawSettings.zOrder === 'back') previewRectangle.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 対象外のアートボードに対応するプレビュー長方形を隠す
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function hidePreviewItemsOutOfScope(doc, previewLayer, drawSettings) {
        /* すべてのアートボードが対象のときは -1（範囲外の番号だけ隠す）/ -1 means "every artboard is in scope" */
        var visibleArtboardIndex = (drawSettings.target === 'all') ? -1 : doc.artboards.getActiveArtboardIndex();
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var previewItem = previewLayer.pathItems[i];
            /* アイテム名は "__Preview_...#<アートボード番号>" 形式 */
            var indexMatch = /#(\d+)$/.exec(previewItem.name || "");
            if (!indexMatch) continue;
            var itemArtboardIndex = parseInt(indexMatch[1], 10);
            var keepVisible = (visibleArtboardIndex < 0) ?
                (itemArtboardIndex < doc.artboards.length) :
                (itemArtboardIndex === visibleArtboardIndex);
            if (!keepVisible) previewItem.hidden = true;
        }
    }

    /**
     * プレビューを描画する（専用レイヤーへ一時オブジェクトを生成）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function renderPreview(doc, drawSettings) {
        /* レイヤーごと消さずに既存プレビューを隠す / Hide existing preview items instead of deleting the layer */
        clearPreview(false);
        if (!doc || !drawSettings) return;

        var previousCoordinateSystem = null;
        try {
            previousCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }

        var previewLayer = getOrCreatePreviewLayer(doc);
        if (drawSettings.target === 'all') {
            for (var i = 0; i < doc.artboards.length; i++) drawPreviewRectangle(doc, previewLayer, i, drawSettings);
        } else {
            drawPreviewRectangle(doc, previewLayer, doc.artboards.getActiveArtboardIndex(), drawSettings);
        }
        hidePreviewItemsOutOfScope(doc, previewLayer, drawSettings);

        try {
            if (previousCoordinateSystem !== null) app.coordinateSystem = previousCoordinateSystem;
        } catch (e) { }
        PreviewHistory.bump();
        app.redraw();
    }

    // =========================================
    // 入力欄の検証 / Field validation
    // =========================================

    /**
     * HEX入力欄の警告表示を切り替える
     * @param {EditText} hexField - HEX入力欄
     * @param {boolean} isWarning - true で警告表示
     * @param {string} [warningKey] - 警告時のヘルプチップのラベルキー
     * @returns {void}
     */
    function setHexWarning(hexField, isWarning, warningKey) {
        setFieldWarnColor(hexField, isWarning);
        hexField.helpTip = isWarning ? getLabel(warningKey || 'warning.hexInvalid') : getLabel('helpTip.hexInput');
        try { hexField.notify('onDraw'); } catch (e) { }
    }

    /**
     * CMYK入力欄の警告表示を切り替える
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄
     * @param {boolean} isWarning - true で警告表示
     * @returns {void}
     */
    function setCmykWarning(channelInput, isWarning) {
        setFieldWarnColor(channelInput, isWarning);
        channelInput.helpTip = isWarning ? getLabel('warning.cmykRange') : getLabel('helpTip.cmykInput');
    }

    /**
     * CMYK入力欄の値が0–100の範囲かを判定して警告表示に反映する（入力途中は警告しない）
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄
     * @returns {void}
     */
    function validateCmykField(channelInput) {
        var fieldText = String(channelInput.text || '');
        if (fieldText === '') {
            setCmykWarning(channelInput, false);
            return;
        }
        var channelValue = parseFloat(fieldText);
        setCmykWarning(channelInput, isNaN(channelValue) || channelValue < 0 || channelValue > 100);
    }

    /**
     * CMYK入力欄の値を0–100へ丸める（未入力は空のまま。計算時に0として扱う）
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄
     * @returns {void}
     */
    function clampCmykField(channelInput) {
        var fieldText = String(channelInput.text || '').replace(/^\s+|\s+$/g, '');
        if (fieldText !== '') {
            var channelValue = parseFloat(fieldText);
            if (isNaN(channelValue)) channelValue = 0;
            channelInput.text = String(clampValue(channelValue, 0, 100));
        }
        setCmykWarning(channelInput, false);
    }

    /**
     * CMYK入力欄へ共通のハンドラをまとめて登録する
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred, trackFocus }
     * @returns {void}
     */
    function bindCmykField(channelInput, previewHooks) {
        channelInput.addEventListener('focus', function () {
            /* ちょうど "0" のときは入力しやすいようクリア / Clear a lone "0" so typing replaces it */
            if (String(channelInput.text) === '0') channelInput.text = '';
        });

        channelInput.addEventListener('keydown', function (event) {
            /* 先頭ゼロ（"03"）を作らせない。小数 "0.5" は触らない
               Prevent leading-zero integers; leave decimals like "0.5" alone */
            var typedKey = String(event.keyName || '');
            if (!/^[0-9]$/.test(typedKey)) return;
            var fieldText = String(channelInput.text || '');
            if (/\./.test(fieldText)) return;
            if (/^0+$/.test(fieldText)) channelInput.text = '';
            else if (/^0\d+$/.test(fieldText)) channelInput.text = fieldText.replace(/^0+/, '');
        });

        channelInput.onChanging = function () {
            var fieldText = String(channelInput.text || '');
            if (/^0\d+$/.test(fieldText)) channelInput.text = fieldText.replace(/^0+/, '');
            validateCmykField(channelInput);
            previewHooks.deferred();
        };

        channelInput.onChange = function () {
            clampCmykField(channelInput);
            previewHooks.immediate();
        };

        changeValueByArrowKey(channelInput, function () {
            clampCmykField(channelInput);
            previewHooks.deferred();
        });

        previewHooks.trackFocus(channelInput);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * オフセットパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred, trackFocus }
     * @returns {object} { offsetInput, bleedCheckbox, initFieldState }
     */
    function buildOffsetPanel(parentGroup, previewHooks) {
        var offsetPanel = parentGroup.add('panel', undefined, getLabel('panel.offset'));
        setupPanel(offsetPanel);

        var offsetRow = offsetPanel.add('group');
        setupGroup(offsetRow, 'row');
        offsetRow.alignChildren = 'center';
        offsetRow.alignment = 'center';

        var offsetInput = offsetRow.add('edittext', undefined, '0');
        offsetInput.characters = 4;
        offsetInput.helpTip = getLabel('helpTip.offsetInput');
        offsetRow.add('statictext', undefined, getCurrentUnitLabel());

        var bleedRow = offsetPanel.add('group');
        setupGroup(bleedRow, 'row');
        bleedRow.alignChildren = 'center';
        bleedRow.alignment = 'center';

        var bleedCheckbox = bleedRow.add('checkbox', undefined, getLabel('checkbox.bleed'));
        bleedCheckbox.alignment = 'center';
        bleedCheckbox.value = false; /* デフォルトOFF / default OFF */
        bleedCheckbox.helpTip = getLabel('helpTip.bleed');

        /* 裁ち落としON/OFFの往復で手入力値を失わないよう控えておく
           Remember the manual offset so toggling Bleed does not lose it */
        var manualOffsetText = '0';

        /* 裁ち落としの状態を入力欄へ反映 / Reflect the current Bleed state in the field */
        function applyBleedState(refreshPreview) {
            if (bleedCheckbox.value) {
                var resolvedOffset = resolveOffsetToPt(offsetInput.text, getCurrentUnitCode(), true);
                offsetInput.text = resolvedOffset.displayText;
                offsetInput.enabled = !resolvedOffset.disabled;
            } else {
                offsetInput.text = manualOffsetText;
                offsetInput.enabled = true;
            }
            if (refreshPreview) previewHooks.deferred();
        }

        bleedCheckbox.onClick = function () {
            if (bleedCheckbox.value) manualOffsetText = String(offsetInput.text);
            applyBleedState(true);
        };

        offsetInput.onChanging = previewHooks.deferred;
        offsetInput.onChange = previewHooks.immediate;
        offsetInput.addEventListener('keydown', function (event) {
            /* Enterでも即座に反映 / Enter refreshes the preview immediately */
            if (event.keyName == 'Enter') previewHooks.immediate();
        });
        changeValueByArrowKey(offsetInput, previewHooks.deferred);
        previewHooks.trackFocus(offsetInput);

        return {
            offsetInput: offsetInput,
            bleedCheckbox: bleedCheckbox,
            initFieldState: function () {
                manualOffsetText = String(offsetInput.text);
                applyBleedState(false);
            }
        };
    }

    /**
     * カラーパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred, trackFocus }
     * @returns {object} 各ラジオ・入力欄をまとめたオブジェクト
     */
    function buildColorPanel(parentGroup, previewHooks) {
        var colorPanel = parentGroup.add('panel', undefined, getLabel('panel.color'));
        setupPanel(colorPanel, STACK_SPACING); /* やや広めの行間 / a bit more vertical gap */

        var noneRadio = colorPanel.add('radiobutton', undefined, getLabel('color.none'));
        var k100Radio = colorPanel.add('radiobutton', undefined, getLabel('color.k100'));

        /* HEXはラジオと入力欄を同じ行に / HEX radio and its field share one row */
        var hexRow = colorPanel.add('group');
        setupGroup(hexRow, 'row', TIGHT_SPACING);
        var hexRadio = hexRow.add('radiobutton', undefined, getLabel('color.hex'));
        var hexInput = hexRow.add('edittext', undefined, '#');
        hexInput.characters = 14; /* カラム幅が伸びないよう控えめに / narrow enough to keep the column width */
        hexInput.helpTip = getLabel('helpTip.hexInput');

        var cmykRadio = colorPanel.add('radiobutton', undefined, getLabel('color.cmyk'));

        /* ラベル行とフィールド行の2段グリッド / Two-row grid: labels on top, fields below */
        var cmykGrid = colorPanel.add('group');
        setupGroup(cmykGrid, 'column', 4);
        var cmykLabelRow = cmykGrid.add('group');
        setupGroup(cmykLabelRow, 'row', STACK_SPACING);
        var cmykFieldRow = cmykGrid.add('group');
        setupGroup(cmykFieldRow, 'row', STACK_SPACING);

        var cmykChannelTexts = ['  C', '  M', '  Y', '  K'];
        var cmykLabels = [];
        var cmykInputs = [];
        for (var i = 0; i < cmykChannelTexts.length; i++) {
            var channelLabel = cmykLabelRow.add('statictext', undefined, cmykChannelTexts[i]);
            channelLabel.preferredSize.width = CMYK_FIELD_WIDTH;
            var channelInput = cmykFieldRow.add('edittext', undefined, '');
            channelInput.characters = 3;
            channelInput.preferredSize.width = CMYK_FIELD_WIDTH;
            channelInput.helpTip = getLabel('helpTip.cmykInput');
            cmykLabels.push(channelLabel);
            cmykInputs.push(channelInput);
            bindCmykField(channelInput, previewHooks);
        }

        hexInput.onChanging = function () {
            var hexText = String(hexInput.text || '').replace(/\s+/g, '');
            if (hexText === '') setHexWarning(hexInput, false);
            else if (hexText === '#') setHexWarning(hexInput, true, 'warning.hexEmpty');
            /* parseColorText が解釈できる入力（#RRGGBB／短縮HEX／色名／grayNN）はすべて有効
               Anything parseColorText can resolve is valid */
            else setHexWarning(hexInput, !parseColorText(app.activeDocument, hexText));
            previewHooks.deferred();
        };

        hexInput.onChange = function () {
            var hexText = String(hexInput.text || '').replace(/\s+/g, '');
            if (/^#?[0-9a-fA-F]{6}$/.test(hexText)) {
                /* 6桁HEXは # 付き・大文字へ正規化 / Normalize 6-digit hex to "#" + uppercase */
                hexInput.text = '#' + hexText.replace(/^#/, '').toUpperCase();
                setHexWarning(hexInput, false);
            } else if (hexText === '#') {
                setHexWarning(hexInput, true, 'warning.hexEmpty');
            } else {
                /* 色名・短縮HEX・grayNN は整形せずそのまま受理 / Names, shorthand hex and grayNN pass through */
                setHexWarning(hexInput, !parseColorText(app.activeDocument, hexText));
            }
            previewHooks.immediate();
        };

        previewHooks.trackFocus(hexInput);

        /* ラジオ選択に応じて入力欄の有効・無効を反映 / Sync field enable state with the radios */
        function updateColorFieldStates() {
            hexInput.enabled = !!hexRadio.value;
            var cmykEnabled = !!cmykRadio.value;
            for (var i = 0; i < cmykInputs.length; i++) {
                cmykInputs[i].enabled = cmykEnabled;
                cmykLabels[i].enabled = cmykEnabled;
                if (!cmykEnabled) setCmykWarning(cmykInputs[i], false);
            }
        }

        /* カラーモードを排他選択し、ハイライト・フォーカス・プレビューを更新
           Select a color mode exclusively, then sync highlight, focus and preview */
        function selectColorMode(colorMode) {
            noneRadio.value = (colorMode === ColorMode.NONE);
            k100Radio.value = (colorMode === ColorMode.K100);
            hexRadio.value = (colorMode === ColorMode.HEX);
            cmykRadio.value = (colorMode === ColorMode.CMYK);
            updateColorFieldStates();
            setFieldHighlight(hexInput, colorMode === ColorMode.HEX);
            setFieldHighlight(cmykInputs[0], colorMode === ColorMode.CMYK);
            try {
                if (colorMode === ColorMode.HEX) hexInput.active = true;
                else if (colorMode === ColorMode.CMYK) cmykInputs[0].active = true;
            } catch (e) { }
            previewHooks.immediate();
        }

        var colorRadioModes = [
            [noneRadio, ColorMode.NONE],
            [k100Radio, ColorMode.K100],
            [hexRadio, ColorMode.HEX],
            [cmykRadio, ColorMode.CMYK]
        ];
        for (var j = 0; j < colorRadioModes.length; j++) {
            (function (radio, colorMode) {
                radio.onClick = radio.onChanging = function () { selectColorMode(colorMode); };
            })(colorRadioModes[j][0], colorRadioModes[j][1]);
        }

        k100Radio.value = true; /* デフォルトはK100 / default to K100 */
        updateColorFieldStates();

        return {
            noneRadio: noneRadio,
            k100Radio: k100Radio,
            hexRadio: hexRadio,
            cmykRadio: cmykRadio,
            hexInput: hexInput,
            cmykInputs: cmykInputs
        };
    }

    /**
     * 配置位置（重ね順）パネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック
     * @returns {object} { frontRadio, backRadio, bgLayerRadio }
     */
    function buildPlacementPanel(parentGroup, previewHooks) {
        var placementPanel = parentGroup.add('panel', undefined, getLabel('panel.zorder'));
        setupPanel(placementPanel, TIGHT_SPACING);

        var frontRadio = placementPanel.add('radiobutton', undefined, getLabel('zorder.front'));
        var backRadio = placementPanel.add('radiobutton', undefined, getLabel('zorder.back'));
        var bgLayerRadio = placementPanel.add('radiobutton', undefined, getLabel('zorder.bg'));
        bgLayerRadio.helpTip = getLabel('helpTip.bgLayer');

        frontRadio.value = true; /* デフォルトは最前面 / default to Bring to Front */

        var placementRadios = [frontRadio, backRadio, bgLayerRadio];
        for (var i = 0; i < placementRadios.length; i++) {
            placementRadios[i].onClick = previewHooks.immediate;
        }

        return { frontRadio: frontRadio, backRadio: backRadio, bgLayerRadio: bgLayerRadio };
    }

    /**
     * 対象アートボードのパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック
     * @returns {object} { currentArtboardRadio, allArtboardsRadio }
     */
    function buildTargetPanel(parentGroup, previewHooks) {
        var targetPanel = parentGroup.add('panel', undefined, getLabel('panel.target'));
        setupPanel(targetPanel);

        var currentArtboardRadio = targetPanel.add('radiobutton', undefined, getLabel('target.current'));
        var allArtboardsRadio = targetPanel.add('radiobutton', undefined, getLabel('target.all'));

        /* 常に「現在のアートボード」をデフォルト選択 / Always default to the current artboard */
        currentArtboardRadio.value = true;
        allArtboardsRadio.value = false;

        /* 1枚しかない場合は「すべてのアートボード」をディム / Dim "All Artboards" when there is only one */
        var artboardCount = app.documents.length ? app.activeDocument.artboards.length : 0;
        if (artboardCount <= 1) {
            allArtboardsRadio.enabled = false;
            allArtboardsRadio.helpTip = getLabel('warning.singleArtboard');
        }

        /* クリックだけでなくキーボード操作（onChanging）でも更新 / Refresh on click and on keyboard change */
        currentArtboardRadio.onClick = currentArtboardRadio.onChanging = previewHooks.immediate;
        allArtboardsRadio.onClick = allArtboardsRadio.onChanging = previewHooks.immediate;

        return { currentArtboardRadio: currentArtboardRadio, allArtboardsRadio: allArtboardsRadio };
    }

    /**
     * オプションパネル（ガイド化／ライブシェイプ変換）を構築する
     * 各オプションは独立。必要なものだけ描画後に適用（applyDrawOptions）
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @returns {object} { makeGuideCheckbox, convertToLiveShapeCheckbox }
     */
    function buildOptionsPanel(parentGroup) {
        var optionsPanel = parentGroup.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel);

        var makeGuideCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.makeGuide'));
        makeGuideCheckbox.value = false; /* デフォルトOFF / default OFF */
        makeGuideCheckbox.helpTip = getLabel('helpTip.makeGuide');

        var convertToLiveShapeCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.convertToLiveShape'));
        convertToLiveShapeCheckbox.value = true; /* デフォルトON / default ON */
        convertToLiveShapeCheckbox.helpTip = getLabel('helpTip.convertToLiveShape');

        return { makeGuideCheckbox: makeGuideCheckbox, convertToLiveShapeCheckbox: convertToLiveShapeCheckbox };
    }

    /**
     * ダイアログのホットキーを登録する（F/B/L=重ね順、C/A=対象、G=ガイド化）
     * @param {Window} dialog - 対象ダイアログ
     * @param {object} dialogControls - 各パネルのコントロール
     * @param {function} refreshPreview - プレビューを即時更新するコールバック
     * @returns {void}
     */
    function addDialogHotkeys(dialog, dialogControls, refreshPreview) {
        dialog.addEventListener('keydown', function (event) {
            if (focusedField) return; /* 入力中は無効 / ignore while typing in a field */
            var pressedKey = (event && event.keyName) ? String(event.keyName).toUpperCase() : '';

            if (pressedKey === 'G') {
                /* ガイド化は描画後の処理なのでプレビューには反映しない
                   Make-guides is a post-draw option and is not previewed */
                var makeGuideCheckbox = dialogControls.options.makeGuideCheckbox;
                makeGuideCheckbox.value = !makeGuideCheckbox.value;
                event.preventDefault();
                return;
            }

            var selectedRadio = null;
            if (pressedKey === 'F') selectedRadio = dialogControls.placement.frontRadio;
            else if (pressedKey === 'B') selectedRadio = dialogControls.placement.backRadio;
            else if (pressedKey === 'L') selectedRadio = dialogControls.placement.bgLayerRadio;
            else if (pressedKey === 'C') selectedRadio = dialogControls.target.currentArtboardRadio;
            else if (pressedKey === 'A') selectedRadio = dialogControls.target.allArtboardsRadio;
            else return;

            if (selectedRadio.enabled) {
                selectedRadio.value = true;
                refreshPreview();
            }
            event.preventDefault();
        });
    }

    /**
     * ダイアログの入力内容から描画設定を組み立てる（プレビューと確定値で共通）
     * @param {object} dialogControls - 各パネルのコントロール
     * @returns {object} 描画設定
     */
    function collectDrawSettings(dialogControls) {
        var colorControls = dialogControls.color;
        var placementControls = dialogControls.placement;
        var offsetControls = dialogControls.offset;

        var colorMode = ColorMode.NONE;
        if (colorControls.k100Radio.value) colorMode = ColorMode.K100;
        else if (colorControls.hexRadio.value) colorMode = ColorMode.HEX;
        else if (colorControls.cmykRadio.value) colorMode = ColorMode.CMYK;

        var zOrder = placementControls.frontRadio.value ? 'front' :
            (placementControls.bgLayerRadio.value ? 'bg' : 'back');

        /* オフセット計算は resolveOffsetToPt に一元化 / All offset math lives in resolveOffsetToPt */
        var resolvedOffset = resolveOffsetToPt(offsetControls.offsetInput.text, getCurrentUnitCode(), !!offsetControls.bleedCheckbox.value);

        /* 各欄を0–100にクランプ（空欄・不正は0）/ Clamp each field to 0-100 (empty or invalid becomes 0) */
        var cmykChannelKeys = ['c', 'm', 'y', 'k'];
        var cmykValues = { c: 0, m: 0, y: 0, k: 0 };
        for (var i = 0; i < colorControls.cmykInputs.length; i++) {
            var channelValue = parseFloat(colorControls.cmykInputs[i].text);
            if (isNaN(channelValue)) channelValue = 0;
            cmykValues[cmykChannelKeys[i]] = clampValue(channelValue, 0, 100);
        }

        return {
            colorMode: colorMode,
            customValue: String(colorControls.hexInput.text || '').replace(/^\s+|\s+$/g, ''), /* HEX文字列 */
            customCMYK: cmykValues,
            offset: resolvedOffset.pt,
            zOrder: zOrder,
            target: dialogControls.target.allArtboardsRadio.value ? 'all' : 'current',
            bleed: !!offsetControls.bleedCheckbox.value,
            makeGuide: !!dialogControls.options.makeGuideCheckbox.value,
            convertToLiveShape: !!dialogControls.options.convertToLiveShapeCheckbox.value
        };
    }

    /**
     * 設定ダイアログを構築して結果を返す
     * @returns {object|null} 描画設定。キャンセル時は null
     */
    function showDialog() {
        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        DialogPersist.setOpacity(dialog, DIALOG_OPACITY);
        dialog.alignChildren = 'left';

        /* 各パネルより先に定義してコールバックとして配る（実行はパネル構築後）
           Declared before the panels so they can be handed out as callbacks */
        var dialogControls = null;

        function updatePreviewImmediately() {
            PreviewHistory.cancelTask(previewDebounceTaskId);
            try {
                renderPreview(app.activeDocument, collectDrawSettings(dialogControls));
            } catch (e) { }
        }

        function updatePreviewDeferred() {
            try {
                schedulePreview(collectDrawSettings(dialogControls), PREVIEW_DELAY_TYPING_MS);
            } catch (e) { }
        }

        var previewHooks = {
            immediate: updatePreviewImmediately,
            deferred: updatePreviewDeferred,
            trackFocus: trackFocusForHotkeys
        };

        /* 2カラム構成 / Two-column layout */
        var mainColumnsGroup = dialog.add('group');
        setupGroup(mainColumnsGroup, 'row', COLUMN_SPACING);
        mainColumnsGroup.alignChildren = ['fill', 'top']; /* 2カラムを上揃え・横いっぱいに */

        var leftColumnGroup = mainColumnsGroup.add('group');
        setupGroup(leftColumnGroup, 'column', STACK_SPACING);
        leftColumnGroup.alignChildren = 'fill'; /* パネルを列幅いっぱいに / panels fill the column */

        var rightColumnGroup = mainColumnsGroup.add('group');
        setupGroup(rightColumnGroup, 'column', STACK_SPACING);
        rightColumnGroup.alignChildren = 'fill';

        dialogControls = {
            offset: buildOffsetPanel(leftColumnGroup, previewHooks),
            placement: buildPlacementPanel(leftColumnGroup, previewHooks),
            options: buildOptionsPanel(leftColumnGroup),
            color: buildColorPanel(rightColumnGroup, previewHooks),
            target: buildTargetPanel(rightColumnGroup, previewHooks)
        };

        addDialogHotkeys(dialog, dialogControls, updatePreviewImmediately);

        /* ボタン行 / Button row */
        var btnRowGroup = dialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['fill', 'center'];
        btnRowGroup.alignment = 'fill';

        var btnLeftGroup = btnRowGroup.add('group');
        setupGroup(btnLeftGroup, 'row');

        var isPreviewDisplayMode = true;
        var btnDisplayToggle = btnLeftGroup.add('button', undefined, getLabel('button.previewOutline'));
        btnDisplayToggle.helpTip = getLabel('helpTip.previewToggle');

        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.orientation = 'row';
        btnRightGroup.alignment = ['right', 'center']; /* 右カラムは右揃え / right-align the right column */

        var btnCancel = btnRightGroup.add('button', undefined, getLabel('button.cancel'));
        var btnOK = btnRightGroup.add('button', undefined, getLabel('button.ok'));

        btnDisplayToggle.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                isPreviewDisplayMode = !isPreviewDisplayMode;
                btnDisplayToggle.text = isPreviewDisplayMode ? getLabel('button.previewOutline') : getLabel('button.previewPreview');
            } catch (e) { }
        };

        /* プレビューを片付けてから閉じる / Clean the preview up, then close */
        function closeWithCleanup(resultCode) {
            PreviewHistory.cancelTask(previewDebounceTaskId);
            PreviewHistory.undo();
            clearPreview(true); /* undo回数に依存せず _preview レイヤーを確実に削除 */
            dialog.close(resultCode);
        }

        btnOK.onClick = function () { closeWithCleanup(1); };
        btnCancel.onClick = function () { closeWithCleanup(0); };

        dialog.onShow = function () {
            DialogPersist.applyInitialOffset(dialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
            dialogControls.offset.initFieldState();
            try { dialogControls.offset.offsetInput.active = true; } catch (e) { }
            PreviewHistory.start(); /* プレビューのUndoカウンタを初期化 */
            updatePreviewImmediately();
        };

        if (dialog.show() != 1) return null;

        /* 確定値もプレビューと同じ計算経路から取る / Final values come from the same computation as the preview */
        return collectDrawSettings(dialogControls);
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 編集可能なレイヤーを取得する（なければ作成）
     * テンプレートレイヤーは locked が false でも編集できない（Error 8705）ので除外し、
     * プレビュー用レイヤーは本番の描画先にしない（残存時の誤描画防止）。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 描画先レイヤー
     */
    function getWritableLayer(doc) {
        function isWritableLayer(layer) {
            return !!layer && !layer.locked && layer.visible && !layer.template && !isPreviewLayerName(layer.name);
        }
        try {
            if (isWritableLayer(doc.activeLayer)) return doc.activeLayer;
            for (var i = 0; i < doc.layers.length; i++) {
                if (isWritableLayer(doc.layers[i])) return doc.layers[i];
            }
            var newLayer = doc.layers.add();
            newLayer.name = FALLBACK_LAYER_NAME;
            return newLayer;
        } catch (e) { }
        return doc.activeLayer;
    }

    /**
     * 「bg」レイヤーを取得する（なければ作成し、最背面へ移動）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} bgレイヤー
     */
    function getOrCreateBgLayer(doc) {
        var bgLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === BG_LAYER_NAME) {
                bgLayer = doc.layers[i];
                break;
            }
        }
        if (!bgLayer) {
            bgLayer = doc.layers.add();
            bgLayer.name = BG_LAYER_NAME;
        }
        /* 見える＆編集可能にしてから最背面へ / Make it visible and editable, then send it to the back */
        bgLayer.visible = true;
        bgLayer.locked = false;
        try {
            bgLayer.printable = true;
            bgLayer.move(doc, ElementPlacement.PLACEATEND);
        } catch (e) { }
        return bgLayer;
    }

    /**
     * アートボードと同サイズ（オフセット込み）の長方形を1枚描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {Artboard} artboard - 対象アートボード
     * @param {object} drawSettings - 描画設定
     * @returns {PathItem} 描画した長方形
     */
    function drawRectangleForArtboard(doc, artboard, drawSettings) {
        var artboardRect = artboard.artboardRect; /* [left, top, right, bottom] */
        var artboardWidth = artboardRect[2] - artboardRect[0];
        var artboardHeight = artboardRect[1] - artboardRect[3];
        var offsetPt = drawSettings.offset;

        var targetLayer = (drawSettings.zOrder === 'bg') ? getOrCreateBgLayer(doc) : getWritableLayer(doc);
        /* アクティブレイヤーがロックされたままだと、別の編集可能レイヤーへ作成しても
           Illustrator が Error 8705（対象レイヤーは編集できません）を投げる。
           Make the target layer editable AND active before creating, or a locked active layer
           triggers Error 8705 "Target layer cannot be modified". */
        try {
            targetLayer.locked = false;
            targetLayer.visible = true;
            doc.activeLayer = targetLayer;
        } catch (e) { }

        var artboardRectangle = targetLayer.pathItems.rectangle(
            artboardRect[1] + offsetPt,
            artboardRect[0] - offsetPt,
            artboardWidth + offsetPt * 2,
            artboardHeight + offsetPt * 2
        );

        applyFillByMode(doc, artboardRectangle, drawSettings);
        artboardRectangle.name = getLabel('name.rect');
        artboardRectangle.selected = true;

        if (drawSettings.zOrder === 'front') artboardRectangle.zOrder(ZOrderMethod.BRINGTOFRONT);
        else if (drawSettings.zOrder === 'back') artboardRectangle.zOrder(ZOrderMethod.SENDTOBACK);

        return artboardRectangle;
    }

    /**
     * 中心の○（属性パネル「中心点を表示」）を選択オブジェクトへ適用する
     * API・メニューコマンドからは設定できないため、記録済みアクション(.aia)を一時ファイルへ書き出して
     * loadAction→doScript で再生する。呼び出し側で「対象だけを選択した状態」にしてから実行すること。
     * @returns {void}
     */
    function showShapeCenterWidget() {
        var ACTION_SET_NAME = 'SmartDrawArtboardRectangle';
        var ACTION_NAME = 'CenterPoint';
        var ACTION_BODY = [
            '/version 3',
            '/name [ 26',
            '\t536d61727444726177417274626f61726452656374616e676c65',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '\t/name [ 11',
            '\t\t43656e746572506f696e74',
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

        var actionFile = null;
        try {
            /* 一時ファイルへ書き出し / Write the recorded action to a temp file */
            actionFile = new File(Folder.temp + '/SmartDrawArtboardRectangle_center.aia');
            actionFile.encoding = 'UTF-8';
            actionFile.open('w');
            actionFile.write(ACTION_BODY);
            actionFile.close();

            /* 同名セットを解放してからロード→再生 / Unload any same-named set, then load and play */
            try { app.unloadAction(ACTION_SET_NAME, ''); } catch (e) { }
            app.loadAction(actionFile);
            app.doScript(ACTION_NAME, ACTION_SET_NAME, false);
        } catch (e) {
        } finally {
            try { app.unloadAction(ACTION_SET_NAME, ''); } catch (e) { }
            try { if (actionFile && actionFile.exists) actionFile.remove(); } catch (e) { }
        }
    }

    /**
     * 指定アイテムだけを選択状態にする
     * @param {PathItem[]} items - 選択したいアイテム
     * @returns {void}
     */
    function selectOnly(items) {
        try { app.executeMenuCommand('deselectall'); } catch (e) { }
        try {
            for (var i = 0; i < items.length; i++) items[i].selected = true;
        } catch (e) { }
    }

    /**
     * 描画後のオプション（ライブシェイプ化／中心の○表示／ガイド化）を適用する
     * @param {PathItem[]} createdRectangles - 描画した長方形
     * @param {object} drawSettings - 描画設定
     * @returns {void}
     */
    function applyDrawOptions(createdRectangles, drawSettings) {
        if (!createdRectangles || !createdRectangles.length) return;

        if (drawSettings.convertToLiveShape) {
            /* 選択ベースのメニューコマンドなので、対象だけを選択してから実行 */
            selectOnly(createdRectangles);
            try { app.executeMenuCommand('Convert to Shape'); } catch (e) { }
            /* 中心の○はライブシェイプにのみ表示される。変換後に選択し直してから適用
               The center widget only renders on live shapes, so re-select after converting */
            selectOnly(createdRectangles);
            showShapeCenterWidget();
        }

        if (drawSettings.makeGuide) {
            /* PathItem.guides を直接立てる（選択・メニュー状態に依存せず確実）
               Set PathItem.guides directly - robust, independent of selection and menu state */
            for (var i = 0; i < createdRectangles.length; i++) {
                createdRectangles[i].name = getLabel('name.guide');
                createdRectangles[i].guides = true;
            }
            try { app.executeMenuCommand('deselectall'); } catch (e) { }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント：ダイアログ→描画→オプション適用
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var drawSettings = showDialog();
        if (drawSettings === null) return;

        var doc = app.activeDocument;

        var previousCoordinateSystem = null;
        try {
            previousCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }

        app.executeMenuCommand('deselectall'); /* 既存選択を解除 / clear any existing selection */

        var createdRectangles = [];
        if (drawSettings.target === 'all') {
            for (var i = 0; i < doc.artboards.length; i++) {
                createdRectangles.push(drawRectangleForArtboard(doc, doc.artboards[i], drawSettings));
            }
        } else {
            var currentArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            createdRectangles.push(drawRectangleForArtboard(doc, currentArtboard, drawSettings));
        }

        applyDrawOptions(createdRectangles, drawSettings);

        try {
            if (previousCoordinateSystem !== null) app.coordinateSystem = previousCoordinateSystem;
        } catch (e) { }
    }

    main();

})();
