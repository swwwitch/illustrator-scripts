#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した配置画像・ラスター画像・ベクター・テキストから代表色を抽出し、16／11／8／5色のカラーパレットを元オブジェクトの下に描画します。
オブジェクトを選択していないときは、スウォッチパネルで選択中のスウォッチから色玉パレットを作成します。

詳細は README を参照してください。

### Overview

Extracts representative colors from the selected placed image, raster image, vector art or text and draws a 16, 11, 8 or 5 color palette below the original.
With nothing selected, it builds the palette from the swatches selected in the Swatches panel instead.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ColorPaletteFromImage";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-15";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorPaletteFromImage.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorPaletteFromImage.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8b57cf662462"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User configuration
    // =========================================

    /*
    ほぼ白の除外しきい値 / Near-white exclusion thresholds

    目的:
    11色の代表色選出時に「背景の白」などを除外し、中間色を拾いやすくする。

    判定ロジック:
    1. まず RGB の見た目の明るさで判定
       - R,G,B がすべて NEAR_WHITE_RGB_MIN 以上なら「白候補」
    2. CMYK が取得できる場合
       - C+M+Y+K が NEAR_WHITE_CMYK_TOTAL_MAX 以下なら「ほぼ白」
    3. CMYK が取得できない場合
       - RGB 条件のみで「ほぼ白」と判定

    ※ 淡い色（薄いピンク・水色など）を誤って白扱いしないため、
       RGB 判定を先に行う保守的な条件になっている。
    */
    var NEAR_WHITE_CMYK_TOTAL_MAX = 20;   /* この値以下の C+M+Y+K を「ほぼ白」とみなす / near-white when C+M+Y+K is at or below this */
    var NEAR_WHITE_RGB_MIN = 235;         /* この値以上の RGB 各チャンネルを「ほぼ白」とみなす / near-white when every RGB channel is at or above this */

    /*
    ほぼ黒の除外しきい値 / Near-black exclusion thresholds

    目的:
    11色の代表色選出時に「背景の黒」などの極端に暗い色を除外する。

    判定ロジック:
    1. まず RGB の暗さで判定
       - R,G,B がすべて NEAR_BLACK_RGB_MAX 以下なら「黒候補」
    2. CMYK が取得できる場合
       - (C+M+Y+K が NEAR_BLACK_CMYK_TOTAL_MIN 以上) または
       - (K が NEAR_BLACK_K_MIN 以上)
       のどちらかを満たす場合「ほぼ黒」
    3. CMYK が取得できない場合
       - RGB 条件のみで「ほぼ黒」と判定

    ※ 濃い色（濃紺・濃茶など）を黒として誤除外しないよう、
       RGB と CMYK の両方で確認する保守的な条件になっている。
    */
    var NEAR_BLACK_CMYK_TOTAL_MIN = 280;  /* この値以上の C+M+Y+K を「ほぼ黒」とみなす / near-black when C+M+Y+K is at or above this */
    var NEAR_BLACK_K_MIN = 85;            /* この値以上の K 単独値を「ほぼ黒」とみなす / near-black when K alone is at or above this */
    var NEAR_BLACK_RGB_MAX = 40;          /* この値以下の RGB 各チャンネルを「ほぼ黒」とみなす / near-black when every RGB channel is at or below this */

    /* クリップグループのラスタライズ設定 / Rasterize settings for clip groups */
    var RASTERIZE_RESOLUTION = 300;       /* 解像度（ppi） / resolution in ppi */
    var RASTERIZE_TRANSPARENCY = false;   /* 背景 false=白 true=透明 / background: false = white, true = transparent */
    var RASTERIZE_PADDING = 0;            /* 余白（px） / padding in px */
    var RASTERIZE_ANTIALIAS = true;       /* アンチエイリアス / anti-aliasing */

    /* 既定で選ぶ画像トレースプリセット。名前がローカライズされるため候補を順に試す。
       角カッコ・空白・大文字小文字・全角数字の違いは正規化して吸収する。
       「16色変換」に誤ヒットしないよう、部分一致ではなく完全一致で探す。
       Default image trace preset. The name is localized, so candidates are tried in order.
       Brackets, spaces, letter case, and full-width digits are normalized away before comparing.
       Matched exactly, not by substring, so that "[16 Colors]" is not picked up by mistake. */
    var DEFAULT_TRACING_PRESETS = ["6色変換", "6色", "6カラー", "6 Colors", "6 Color"];

    /* CMYK補正の丸め幅（%） / Rounding step for the CMYK-adjusted row */
    var CMYK_ROUND_STEP = 5;

    /* スウォッチ名の重複回避で試す連番の上限 / Highest suffix tried when making a swatch name unique */
    var SWATCH_NAME_MAX_SUFFIX = 999;

    /* ラベルに使うフォント（先頭から順に試す） / Label fonts, tried in order */
    var LABEL_FONT_NAMES = ["MyriadPro-Regular", "Myriad Pro"];

    /* 一時的に作成するレイヤー・グループの名前 / Names of temporary layers and groups */
    var WORK_LAYER_NAME = "__workLayer__";
    var PALETTE_GROUP_NAME = "__ColorPalette__";
    var PREVIEW_GROUP_NAME = "__ColorPalettePreview__";

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* ダイアログ内の寸法 / Sizes inside the dialogs */
    var PRESET_LIST_SIZE = [150, 300];       /* プリセット一覧の寸法 / preset listbox size */
    var DIALOG_BUTTON_SIZE = [90, 26];       /* ボタンの寸法 / dialog button size */
    var PROGRESS_BAR_SIZE = [260, 14];       /* 進捗バーの寸法 / progress bar size */
    var PROGRESS_TEXT_CHARS = 28;            /* 進捗テキストの最小幅 / minimum width of the progress text */

    /* パレットの寸法比 / Palette proportions */
    var PALETTE_MAX_COLUMNS = 16;            /* 最上段の色数（寸法の基準） / column count that defines the cell size */
    var PALETTE_GAP_RATIO = 0.10;            /* 色玉の間隔比 / gap as a ratio of the cell size */
    var PALETTE_ROW_GAP_DIVISOR = 120;       /* 行間＝元オブジェクト幅 ÷ この値 / row gap = source width divided by this */
    var SWATCH_MODE_SQUARE_SIZE = 60;        /* スウォッチ選択時の色玉の一辺 / square size in swatch mode */
    var LABEL_SIZE_RATIO = 10;               /* 色玉の一辺に対するラベル文字サイズの比 / square size to label font size ratio */
    var ADJUSTED_ROW_GAP_RATIO = 0.25;       /* CMYK補正行を下げる量の比 / offset ratio for the CMYK-adjusted row */
    var FIT_VIEW_MARGIN = 1.05;              /* 画面にフィットしたときの余裕 / breathing room when fitting the view */

    /**
     * ウィンドウの余白と間隔をそろえる
     * @param {Window} win - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの余白と間隔をそろえる
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING
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
     * 横並びグループの配置と間隔をそろえる（ボタン列など）
     * @param {Group} group - 対象グループ
     * @param {string|Array} [alignment] - グループ自身の alignment。省略時は "left"
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 見出し付きパネルを追加する
     * @param {Window|Group} parent - 追加先
     * @param {string} labelText - パネルの見出し
     * @param {number} [spacing] - パネル内の要素間隔
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelText, spacing) {
        var panel = parent.add("panel", undefined, labelText);
        setupPanel(panel, spacing);
        return panel;
    }

    /**
     * 縦並びのカラムグループを追加する
     * @param {Window|Group} parent - 追加先
     * @param {string|Array} [alignment] - グループ自身の alignment
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parent, alignment) {
        var column = parent.add("group");
        column.orientation = "column";
        column.alignChildren = "fill";
        if (alignment) column.alignment = alignment;
        return column;
    }

    /**
     * パネル幅いっぱいに広げないチェックボックスを追加する
     * @param {Panel} panel - 追加先のパネル
     * @param {string} labelText - チェックボックスのラベル
     * @param {string} [tooltipText] - ツールチップ
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addLeftCheckbox(panel, labelText, tooltipText) {
        var checkbox = panel.add("checkbox", undefined, labelText);
        checkbox.alignment = "left";
        if (tooltipText) checkbox.helpTip = tooltipText;
        return checkbox;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境のロケールから表示言語を決める / Pick the UI language from the locale */
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            output: { ja: "カラーパレット作成", en: "Create Color Palette" },
            preset: { ja: "画像トレース（プリセット選択）", en: "Image Trace (Preset Selection)" },
            progress: { ja: "処理中", en: "Processing" }
        },
        panel: {
            outputRows: { ja: "出力する行", en: "Rows to Output" },
            colorInfo: { ja: "カラー情報", en: "Color Info" }
        },
        radio: {
            allRows: { ja: "すべての行", en: "All rows" },
            fiveRowsOnly: { ja: "5色の行のみ", en: "5-color rows only" }
        },
        listHeader: {
            presetBuiltIn: { ja: "標準", en: "Built-in" },
            presetCustom: { ja: "ユーザー定義", en: "Custom" }
        },
        checkbox: {
            count16: { ja: "16色", en: "16 Colors" },
            count11: { ja: "11色", en: "11 Colors" },
            count8: { ja: "8色", en: "8 Colors" },
            count5: { ja: "5色", en: "5 Colors" },
            count5Adjusted: { ja: "5色（CMYK補正）", en: "5 Colors (CMYK Rounded)" },
            hex: { ja: "HEX", en: "HEX" },
            cmyk: { ja: "CMYK", en: "CMYK" },
            fitView: { ja: "画面にフィット", en: "Fit View" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            retry: { ja: "選び直す", en: "Reselect" }
        },
        tooltip: {
            allRows: {
                ja: "16色・11色・8色の行もあわせて出力します。",
                en: "Also outputs the 16, 11, and 8 color rows."
            },
            fiveRowsOnly: {
                ja: "16色・11色・8色の行を出力せず、5色の行だけを出力します。",
                en: "Outputs only the 5-color rows, skipping the 16, 11, and 8 color rows."
            },
            count16: {
                ja: "元オブジェクトの幅にフィットする16色の行を出力します。他の行の幅もこの行にそろえます。",
                en: "Outputs a 16-color row fitted to the source width. The other rows align to it."
            },
            count11: {
                ja: "「ほぼ白」「ほぼ黒」を除いた11色の行を出力します。",
                en: "Outputs an 11-color row with near-white and near-black colors excluded."
            },
            count5Adjusted: {
                ja: "5色の各値を5%刻みに丸めた行を出力します。",
                en: "Outputs a row with each value rounded to the nearest 5%."
            },
            hexLabel: {
                ja: "5色の行に HEX 値を表示します。",
                en: "Shows HEX values under the 5-color row."
            },
            cmykLabel: {
                ja: "5色（CMYK補正）の行に CMYK 値を表示します。",
                en: "Shows CMYK values under the CMYK-rounded 5-color row."
            },
            cmykDocOnly: {
                ja: "CMYKラベルはCMYKドキュメントでのみ利用できます。",
                en: "CMYK labels are available only in a CMYK document."
            },
            retry: {
                ja: "画像トレースのプリセット選択からやり直します。",
                en: "Starts over from the image trace preset selection."
            },
            fitView: {
                ja: "設定を変えるたびに、元オブジェクトとプレビューが収まるよう表示倍率を合わせます。",
                en: "Refits the view to the source object and the preview whenever a setting changes."
            },
            presetBuiltIn: {
                ja: "Illustrator にあらかじめ用意されているプリセットです。",
                en: "Presets that ship with Illustrator."
            },
            presetCustom: {
                ja: "自分で保存したプリセットです。",
                en: "Presets you saved yourself."
            }
        },
        progress: {
            preparing: { ja: "準備中…", en: "Preparing…" },
            tracing: { ja: "色を解析中…", en: "Analyzing Colors…" },
            palette: { ja: "パレット生成中…", en: "Generating Palette…" },
            done: { ja: "完了", en: "Done" }
        },
        group: {
            colors16: { ja: "16色", en: "16 Colors" },
            colorsSuffix: { ja: "色", en: "Colors" },
            swatchPalette: { ja: "スウォッチパレット", en: "Swatch Palette" }
        },
        fallbackName: {
            itemPrefix: { ja: "パレット_", en: "Palette_" },
            unknownColor: { ja: "不明な色", en: "Unknown Color" }
        },
        prefix: {
            cmyk: { ja: "CMYK: ", en: "CMYK: " }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noRow: { ja: "出力する行が選ばれていません。", en: "No rows selected to output." },
            noSwatchSelected: {
                ja: "対象のオブジェクトまたはスウォッチが選択されていません。",
                en: "No applicable objects or swatches are selected."
            }
        }
    };

    /**
     * ラベルをドット区切りのキーで引く
     * @param {string} labelKey - "alert.noDocument" のようなドット区切りキー
     * @returns {string} 現在の表示言語のラベル。見つからない場合はキーをそのまま返す
     */
    function getLabel(labelKey) {
        var keyParts = labelKey.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) return labelKey;
            labelNode = labelNode[keyParts[i]];
        }
        return (labelNode && labelNode[uiLang]) ? labelNode[uiLang] : labelKey;
    }

    /**
     * 握りつぶした例外を ExtendScript コンソールに書き出す
     * @param {Error} e - 発生した例外
     * @param {string} [context] - 発生箇所を示す短い説明
     * @returns {void}
     */
    function logError(e, context) {
        var message = "[" + SCRIPT_NAME + "]";
        if (context) message += " " + context + ":";
        $.writeln(message + " " + e);
    }

    // =========================================
    // カラーの変換 / Color conversion
    // =========================================

    /**
     * 値を 0〜255 に収める
     * @param {number} value - 入力値
     * @returns {number} 0〜255 に収めた値
     */
    function clamp255(value) {
        return Math.max(0, Math.min(255, value));
    }

    /**
     * 値を 0〜100 に収める
     * @param {number} value - 入力値
     * @returns {number} 0〜100 に収めた値
     */
    function clampPercent(value) {
        return Math.max(0, Math.min(100, value));
    }

    /**
     * CMYK_ROUND_STEP 刻みに丸める（0/5 は 2.5、5/10 は 7.5 を境界とする）
     * @param {number} value - 入力値
     * @returns {number} 丸めた値
     */
    function roundToStep(value) {
        var half = CMYK_ROUND_STEP / 2;
        return Math.floor((value + half) / CMYK_ROUND_STEP) * CMYK_ROUND_STEP;
    }

    /**
     * 0〜255 の値を2桁の16進表記に変換する
     * @param {number} value - 0〜255 の値
     * @returns {string} 2桁の16進文字列
     */
    function toHex2(value) {
        var hex = Math.round(value).toString(16).toUpperCase();
        return (hex.length === 1) ? ("0" + hex) : hex;
    }

    /**
     * RGB 値を HEX 文字列に変換する
     * @param {Array<number>} rgb - [R, G, B]
     * @returns {string} "#RRGGBB" 形式の文字列
     */
    function rgbToHex(rgb) {
        return "#" + toHex2(rgb[0]) + toHex2(rgb[1]) + toHex2(rgb[2]);
    }

    /**
     * app.convertSampleColor でカラーを変換する
     * @param {ImageColorSpace} sourceSpace - 変換元のカラースペース
     * @param {Array<number>} sourceValues - 変換元の値
     * @param {ImageColorSpace} targetSpace - 変換先のカラースペース
     * @returns {Array<number>|null} 変換後の値。変換できない場合は null
     */
    function convertSampleColor(sourceSpace, sourceValues, targetSpace) {
        try {
            if (!app.convertSampleColor || typeof ImageColorSpace === "undefined" || typeof ColorConvertPurpose === "undefined") {
                return null;
            }
            var converted = app.convertSampleColor(
                sourceSpace,
                sourceValues,
                targetSpace,
                ColorConvertPurpose.defaultpurpose,
                false,
                false
            );
            return (converted && converted.length >= 3) ? converted : null;
        } catch (e) {
            logError(e, "convertSampleColor");
            return null;
        }
    }

    /**
     * カラーを RGB 値に変換する
     * @param {Color} color - 変換元のカラー
     * @returns {Array<number>} [R, G, B]（0〜255）
     */
    function colorToRGB(color) {
        if (!color) return [0, 0, 0];

        if (color.typename === "RGBColor") {
            return [clamp255(color.red), clamp255(color.green), clamp255(color.blue)];
        }

        /* CMYK はドキュメントの見えに近づけるため Illustrator 自身の変換を優先する / Prefer Illustrator's own conversion for CMYK */
        if (color.typename === "CMYKColor") {
            var fromCmyk = convertSampleColor(
                ImageColorSpace.CMYK,
                [color.cyan, color.magenta, color.yellow, color.black],
                ImageColorSpace.RGB
            );
            if (fromCmyk) return [clamp255(fromCmyk[0]), clamp255(fromCmyk[1]), clamp255(fromCmyk[2])];

            /* 変換が使えない場合の簡易近似 / Simple approximation when sample conversion is unavailable */
            return [
                clamp255(255 * (1 - color.cyan / 100) * (1 - color.black / 100)),
                clamp255(255 * (1 - color.magenta / 100) * (1 - color.black / 100)),
                clamp255(255 * (1 - color.yellow / 100) * (1 - color.black / 100))
            ];
        }

        if (color.typename === "GrayColor") {
            var fromGray = convertSampleColor(ImageColorSpace.GrayScale, [color.gray], ImageColorSpace.RGB);
            if (fromGray) return [clamp255(fromGray[0]), clamp255(fromGray[1]), clamp255(fromGray[2])];

            var grayValue = clamp255(255 * (1 - color.gray / 100));
            return [grayValue, grayValue, grayValue];
        }

        if (color.typename === "LabColor") {
            var fromLab = convertSampleColor(ImageColorSpace.LAB, [color.l, color.a, color.b], ImageColorSpace.RGB);
            if (fromLab) return [clamp255(fromLab[0]), clamp255(fromLab[1]), clamp255(fromLab[2])];
        }

        return [0, 0, 0];
    }

    /**
     * カラーを CMYK 値に変換する
     * @param {Color} color - 変換元のカラー
     * @returns {Array<number>|null} [C, M, Y, K]（0〜100）。変換できない場合は null
     */
    function colorToCMYKVals(color) {
        if (!color) return null;
        if (color.typename === "CMYKColor") {
            return [color.cyan, color.magenta, color.yellow, color.black];
        }

        var converted = null;
        if (color.typename === "RGBColor") {
            converted = convertSampleColor(ImageColorSpace.RGB, [color.red, color.green, color.blue], ImageColorSpace.CMYK);
        } else if (color.typename === "LabColor") {
            converted = convertSampleColor(ImageColorSpace.LAB, [color.l, color.a, color.b], ImageColorSpace.CMYK);
        }

        return (converted && converted.length >= 4) ? [converted[0], converted[1], converted[2], converted[3]] : null;
    }

    /**
     * CMYK 各値を CMYK_ROUND_STEP 刻みに丸めたカラーを作る
     * @param {Color} color - 元のカラー
     * @returns {CMYKColor|null} 丸めた CMYK カラー。変換できない場合は null
     */
    function buildCmykAdjustedColor(color) {
        var cmykVals = colorToCMYKVals(color);
        if (!cmykVals || cmykVals.length < 4) return null;

        var adjustedColor = new CMYKColor();
        adjustedColor.cyan = clampPercent(roundToStep(cmykVals[0]));
        adjustedColor.magenta = clampPercent(roundToStep(cmykVals[1]));
        adjustedColor.yellow = clampPercent(roundToStep(cmykVals[2]));
        adjustedColor.black = clampPercent(roundToStep(cmykVals[3]));
        return adjustedColor;
    }

    /**
     * カラーの一致判定に使うキーを作る
     * @param {Color} color - 対象のカラー
     * @returns {string} RGB 値をまとめた文字列キー
     */
    function colorKey(color) {
        var rgb = colorToRGB(color);
        return Math.round(rgb[0]) + "," + Math.round(rgb[1]) + "," + Math.round(rgb[2]);
    }

    /**
     * カラー値からスウォッチ名を生成する
     * @param {Color} color - 対象のカラー
     * @returns {string} スウォッチ名
     */
    function colorToName(color) {
        if (color.typename === "RGBColor") {
            return "R=" + Math.round(color.red) + " G=" + Math.round(color.green) + " B=" + Math.round(color.blue);
        }
        if (color.typename === "CMYKColor") {
            return "C=" + Math.round(color.cyan) + " M=" + Math.round(color.magenta) +
                " Y=" + Math.round(color.yellow) + " K=" + Math.round(color.black);
        }
        if (color.typename === "GrayColor") {
            return "Gray=" + Math.round(color.gray);
        }
        return getLabel("fallbackName.unknownColor");
    }

    /**
     * ドキュメントが CMYK かどうかを判定する
     * @param {Document} doc - 対象ドキュメント
     * @returns {boolean} CMYK ドキュメントなら true
     */
    function isCmykDocument(doc) {
        try {
            return !!(doc && doc.documentColorSpace === DocumentColorSpace.CMYK);
        } catch (e) {
            return false;
        }
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} ラベル用の黒
     */
    function getLabelBlackColor(doc) {
        if (isCmykDocument(doc)) {
            var cmykBlack = new CMYKColor();
            cmykBlack.cyan = 0;
            cmykBlack.magenta = 0;
            cmykBlack.yellow = 0;
            cmykBlack.black = 100;
            return cmykBlack;
        }
        var rgbBlack = new RGBColor();
        rgbBlack.red = 0;
        rgbBlack.green = 0;
        rgbBlack.blue = 0;
        return rgbBlack;
    }

    // =========================================
    // カラーの抽出 / Color extraction
    // =========================================

    /**
     * グラデーションのカラーストップを個別のカラーとして取り出す
     * @param {GradientColor} gradientColor - 対象のグラデーションカラー
     * @returns {Array<Color>} 取り出したカラーの配列
     */
    function getGradientStopColors(gradientColor) {
        var stopColors = [];
        try {
            var gradientStops = gradientColor.gradient.gradientStops;
            for (var i = 0; i < gradientStops.length; i++) {
                var stopColor = gradientStops[i].color;
                if (stopColor.typename === "SpotColor") stopColor = stopColor.spot.color;
                if (stopColor.typename === "NoColor" || stopColor.typename === "PatternColor") continue;
                stopColors.push(stopColor);
            }
        } catch (e) {
            logError(e, "gradient stops");
        }
        return stopColors;
    }

    /**
     * オブジェクトツリーから塗りのカラーを面積付きで集める
     * @param {PageItem} item - 走査対象のオブジェクト
     * @param {Array<object>} collected - 収集先の配列
     * @param {object} [stats] - 正しく読めた保証がない塗りを数える { unreadableCount: number }
     * @returns {Array<object>} { color: Color, area: number } の配列
     */
    function collectFillColors(item, collected, stats) {
        if (!collected) collected = [];
        if (!item) return collected;

        if (item.typename === "PathItem") {
            /* フリーグラデーションなど DOM から読めない塗りがあっても、1つのパスで全体を止めない
               A fill the DOM cannot read (freeform gradient and the like) must not abort the whole walk */
            var fillColor = null;
            try {
                if (item.filled) fillColor = item.fillColor;
            } catch (e) {
                if (stats) stats.unreadableCount++;
                logError(e, "path fill color");
            }
            if (!fillColor) return collected;

            var area = 1;
            try {
                area = Math.abs(item.area);
            } catch (e) {
                logError(e, "path area");
            }
            if (area < 1) area = 1;

            if (fillColor.typename === "GradientColor") {
                var stopColors = getGradientStopColors(fillColor);
                /* ストップを1つも読めないグラデーション（フリーグラデーションなど）
                   A gradient whose stops cannot be read at all (a freeform gradient and the like) */
                if (stopColors.length === 0 && stats) stats.unreadableCount++;
                for (var i = 0; i < stopColors.length; i++) {
                    collected.push({ color: stopColors[i], area: area / stopColors.length });
                }
            } else if (fillColor.typename !== "NoColor" &&
                fillColor.typename !== "PatternColor" &&
                fillColor.typename !== "SpotColor") {
                /* フリーグラデーションは GrayColor（gray=0）として返るため、本物のグレー塗りと DOM 上は区別できない。
                   取りこぼすより余分にラスタライズするほうが安全なので、読めない塗りとして扱う。
                   Illustrator reports a freeform gradient as GrayColor(gray=0), which is indistinguishable from a
                   real gray fill. Flag it as unreadable: an extra rasterize costs less than losing the colors. */
                if (fillColor.typename === "GrayColor" && stats) stats.unreadableCount++;
                collected.push({ color: fillColor, area: area });
            }
            return collected;
        }

        if (item.typename === "GroupItem" || item.typename === "CompoundPathItem") {
            var children = (item.typename === "GroupItem") ? item.pageItems : item.pathItems;
            for (var k = 0; k < children.length; k++) {
                collectFillColors(children[k], collected, stats);
            }
        }
        return collected;
    }

    /**
     * 同じカラーをまとめ、面積を合算する
     * @param {Array<object>} colors - { color: Color, area: number } の配列
     * @returns {Array<object>} 重複を除いた { color: Color, area: number } の配列
     */
    function deduplicateColors(colors) {
        var seenIndex = {};
        var merged = [];
        for (var i = 0; i < colors.length; i++) {
            var key = colorKey(colors[i].color);
            if (seenIndex[key] === undefined) {
                seenIndex[key] = merged.length;
                merged.push({ color: colors[i].color, area: colors[i].area || 1 });
            } else {
                merged[seenIndex[key]].area += (colors[i].area || 1);
            }
        }
        return merged;
    }

    /**
     * スウォッチパネルで選択中のスウォッチからカラーを取得する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<object>} { color: Color, area: number } の配列
     */
    function getSelectedSwatchColors(doc) {
        var colors = [];
        try {
            var selectedSwatches = doc.swatches.getSelected();
            if (!selectedSwatches || selectedSwatches.length === 0) return colors;

            for (var i = 0; i < selectedSwatches.length; i++) {
                var swatchColor = selectedSwatches[i].color;
                if (!swatchColor) continue;

                /* [なし]・パターンは対象外 / Skip [None] and pattern swatches */
                if (swatchColor.typename === "NoColor" || swatchColor.typename === "PatternColor") continue;

                if (swatchColor.typename === "GradientColor") {
                    var stopColors = getGradientStopColors(swatchColor);
                    for (var k = 0; k < stopColors.length; k++) {
                        colors.push({ color: stopColors[k], area: 1 });
                    }
                    continue;
                }
                colors.push({ color: swatchColor, area: 1 });
            }
        } catch (e) {
            logError(e, "getSelectedSwatchColors");
        }
        return colors;
    }

    // =========================================
    // 代表色の選出 / Representative color selection
    // =========================================

    /**
     * エントリの明度を返す
     * @param {object} entry - { r, g, b } を持つエントリ
     * @returns {number} 明度
     */
    function getLuminance(entry) {
        return entry.r * 0.299 + entry.g * 0.587 + entry.b * 0.114;
    }

    /**
     * 2つのエントリの RGB 距離の2乗を返す
     * @param {object} entryA - 比較元のエントリ
     * @param {object} entryB - 比較先のエントリ
     * @returns {number} 距離の2乗
     */
    function getColorDistance(entryA, entryB) {
        var dr = entryA.r - entryB.r;
        var dg = entryA.g - entryB.g;
        var db = entryA.b - entryB.b;
        return dr * dr + dg * dg + db * db;
    }

    /**
     * 最も暗いエントリの位置を返す
     * @param {Array<object>} entries - { r, g, b } を持つエントリの配列
     * @returns {number} 最も暗いエントリのインデックス
     */
    function findDarkestIndex(entries) {
        var darkestIndex = 0;
        var minLuminance = Infinity;
        for (var i = 0; i < entries.length; i++) {
            var luminance = getLuminance(entries[i]);
            if (luminance < minLuminance) {
                minLuminance = luminance;
                darkestIndex = i;
            }
        }
        return darkestIndex;
    }

    /**
     * 最も暗い色から始めて、最近傍法でグラデーション風に並べ替える
     * @param {Array<object>} entries - { r, g, b } を持つエントリの配列
     * @returns {Array<object>} 並べ替えたエントリの配列
     */
    function sortByNearest(entries) {
        if (entries.length <= 1) return entries.slice();

        var remaining = entries.slice();
        var sorted = [remaining.splice(findDarkestIndex(remaining), 1)[0]];

        while (remaining.length > 0) {
            var last = sorted[sorted.length - 1];
            var nearestIndex = 0;
            var nearestDistance = Infinity;
            for (var i = 0; i < remaining.length; i++) {
                var distance = getColorDistance(last, remaining[i]);
                if (distance < nearestDistance) {
                    nearestDistance = distance;
                    nearestIndex = i;
                }
            }
            sorted.push(remaining.splice(nearestIndex, 1)[0]);
        }
        return sorted;
    }

    /**
     * カラー配列を RGB 付きのエントリに変換し、最近傍法で並べ替える
     * @param {Array<object>} colorEntries - { color, area } または { swatch: { color, area } } の配列
     * @returns {Array<object>} { swatch, r, g, b, area } の配列
     */
    function buildOrderedPaletteEntries(colorEntries) {
        var paletteEntries = [];
        for (var i = 0; i < colorEntries.length; i++) {
            var entry = colorEntries[i];
            var swatch = (entry && entry.swatch) ? entry.swatch : entry;
            if (!swatch) continue;

            var area = 1;
            if (swatch.area !== undefined) {
                area = swatch.area;
            } else if (entry && entry.area !== undefined) {
                area = entry.area;
            }

            var rgb = colorToRGB(swatch.color);
            paletteEntries.push({
                swatch: { color: swatch.color, area: area },
                r: rgb[0],
                g: rgb[1],
                b: rgb[2],
                area: area || 1
            });
        }
        return sortByNearest(paletteEntries);
    }

    /**
     * 最大距離法で N 色を選ぶ（面積で重み付け）
     * スコア = 既選択色との最小距離 × pow(面積 / 平均面積, 0.75)
     * 面積が平均の4倍なら重み約2.83倍、1/4なら約0.35倍で、sqrt より強く面積を反映する。
     * @param {Array<object>} colorList - { swatch, r, g, b, area } の配列
     * @param {number} pickCount - 選出する色数
     * @returns {Array<object>} 選出したエントリの配列
     */
    function selectByMaxDistance(colorList, pickCount) {
        if (colorList.length <= pickCount) return colorList.slice();

        var totalArea = 0;
        for (var i = 0; i < colorList.length; i++) {
            totalArea += (colorList[i].area || 1);
        }
        var averageArea = totalArea / colorList.length;

        var areaWeights = [];
        var isUsed = [];
        for (var k = 0; k < colorList.length; k++) {
            areaWeights.push(Math.pow((colorList[k].area || 1) / averageArea, 0.75));
            isUsed.push(false);
        }

        /* 最初の色は最も暗い色 / The first color is the darkest one */
        var firstIndex = findDarkestIndex(colorList);
        var selectedEntries = [colorList[firstIndex]];
        isUsed[firstIndex] = true;

        /* 既選択色から最も遠い色を順に選ぶ / Repeatedly pick the color farthest from the selected set */
        while (selectedEntries.length < pickCount) {
            var bestIndex = -1;
            var bestScore = -1;

            for (var c = 0; c < colorList.length; c++) {
                if (isUsed[c]) continue;

                var minDistance = Infinity;
                for (var s = 0; s < selectedEntries.length; s++) {
                    var distance = getColorDistance(colorList[c], selectedEntries[s]);
                    if (distance < minDistance) minDistance = distance;
                }

                var score = minDistance * areaWeights[c];
                if (score > bestScore) {
                    bestScore = score;
                    bestIndex = c;
                }
            }

            selectedEntries.push(colorList[bestIndex]);
            isUsed[bestIndex] = true;
        }

        return selectedEntries;
    }

    /**
     * エントリが「ほぼ白」かどうかを判定する
     * RGB の明るさで白候補を絞り、可能なら CMYK 合計量で確認する。
     * @param {object} entry - { swatch, r, g, b } を持つエントリ
     * @returns {boolean} ほぼ白なら true
     */
    function isNearlyWhite(entry) {
        if (entry.r < NEAR_WHITE_RGB_MIN || entry.g < NEAR_WHITE_RGB_MIN || entry.b < NEAR_WHITE_RGB_MIN) {
            return false;
        }
        var cmyk = colorToCMYKVals(entry.swatch.color);
        /* CMYK が取得できない場合は RGB 条件のみで判定する / Fall back to the RGB condition alone */
        if (!cmyk) return true;
        return (cmyk[0] + cmyk[1] + cmyk[2] + cmyk[3]) <= NEAR_WHITE_CMYK_TOTAL_MAX;
    }

    /**
     * エントリが「ほぼ黒」かどうかを判定する
     * RGB の暗さで黒候補を絞り、可能なら CMYK 合計量または K 値で確認する。
     * @param {object} entry - { swatch, r, g, b } を持つエントリ
     * @returns {boolean} ほぼ黒なら true
     */
    function isNearlyBlack(entry) {
        if (entry.r > NEAR_BLACK_RGB_MAX || entry.g > NEAR_BLACK_RGB_MAX || entry.b > NEAR_BLACK_RGB_MAX) {
            return false;
        }
        var cmyk = colorToCMYKVals(entry.swatch.color);
        /* CMYK が取得できない場合は RGB 条件のみで判定する / Fall back to the RGB condition alone */
        if (!cmyk) return true;
        return ((cmyk[0] + cmyk[1] + cmyk[2] + cmyk[3]) >= NEAR_BLACK_CMYK_TOTAL_MIN || cmyk[3] >= NEAR_BLACK_K_MIN);
    }

    /**
     * 11色行の入力色を作る（カスケード時も「ほぼ白／ほぼ黒」の除外を効かせる）
     * @param {Array<object>} colorList - 全カラーのエントリ配列
     * @param {Array<object>} cascadePrev - 直前の行で選ばれたエントリ配列
     * @param {boolean} useCascade - 段階的減色を使うかどうか
     * @returns {Array<object>} 11色行の候補エントリ配列
     */
    function getElevenRowSourceColors(colorList, cascadePrev, useCascade) {
        var filtered = [];
        for (var i = 0; i < colorList.length; i++) {
            if (!isNearlyWhite(colorList[i]) && !isNearlyBlack(colorList[i])) filtered.push(colorList[i]);
        }
        if (filtered.length === 0) filtered = colorList;
        if (!useCascade || !cascadePrev) return filtered;

        var filteredKeys = {};
        for (var k = 0; k < filtered.length; k++) {
            filteredKeys[colorKey(filtered[k].swatch.color)] = true;
        }

        var cascadedFiltered = [];
        for (var j = 0; j < cascadePrev.length; j++) {
            if (filteredKeys[colorKey(cascadePrev[j].swatch.color)]) cascadedFiltered.push(cascadePrev[j]);
        }
        return (cascadedFiltered.length > 0) ? cascadedFiltered : filtered;
    }

    /**
     * 16 → 11 → 8 → 5 の段階的減色で全行の代表色をまとめて選ぶ
     * @param {Array<object>} colorList - 全カラーのエントリ配列
     * @param {object} outputOptions - 出力オプション
     * @returns {object} 色数をキーにしたエントリ配列のマップ
     */
    function buildAllPaletteRows(colorList, outputOptions) {
        var rowsByCount = { 16: [], 11: [], 8: [], 5: [] };
        if (!colorList || !colorList.length) return rowsByCount;

        var useCascade = !outputOptions || outputOptions.cascade !== false;
        var cascadePrev = null;

        if (useCascade || !outputOptions || outputOptions.out16) {
            rowsByCount[16] = sortByNearest(selectByMaxDistance(colorList, 16));
            cascadePrev = rowsByCount[16];
        }

        var rowCounts = [11, 8, 5];
        for (var i = 0; i < rowCounts.length; i++) {
            var rowCount = rowCounts[i];
            var sourceColors;
            if (rowCount === 11) {
                sourceColors = getElevenRowSourceColors(colorList, cascadePrev, useCascade);
            } else {
                sourceColors = (useCascade && cascadePrev) ? cascadePrev : colorList;
            }

            var selectedEntries = selectByMaxDistance(sourceColors, rowCount);
            if (useCascade) cascadePrev = selectedEntries;
            rowsByCount[rowCount] = sortByNearest(selectedEntries);
        }
        return rowsByCount;
    }

    /**
     * 描画とスウォッチ登録で共有する行の選出結果と寸法を求める
     * @param {number} sourceWidth - 元オブジェクトの幅
     * @param {Array<object>} colors - { color, area } の配列
     * @param {object} outputOptions - 出力オプション
     * @returns {object} { rowsByCount, gap, squareSize }
     */
    function buildPaletteRowPlan(sourceWidth, colors, outputOptions) {
        var gap = (sourceWidth / PALETTE_MAX_COLUMNS) * PALETTE_GAP_RATIO;
        return {
            rowsByCount: buildAllPaletteRows(buildOrderedPaletteEntries(colors), outputOptions),
            gap: gap,
            squareSize: (sourceWidth - (PALETTE_MAX_COLUMNS - 1) * gap) / PALETTE_MAX_COLUMNS
        };
    }

    // =========================================
    // スウォッチ登録 / Swatch registration
    // =========================================

    /**
     * 指定した名前のスウォッチが既にあるかを調べる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} swatchName - 調べる名前
     * @returns {boolean} 存在すれば true
     */
    function swatchNameExists(doc, swatchName) {
        try {
            doc.swatches.getByName(swatchName);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 既存のスウォッチ名と衝突しない名前を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} baseName - 基準となる名前
     * @returns {string} 未使用のスウォッチ名
     */
    function buildUniqueSwatchName(doc, baseName) {
        if (!swatchNameExists(doc, baseName)) return baseName;
        for (var i = 2; i <= SWATCH_NAME_MAX_SUFFIX; i++) {
            var candidate = baseName + " " + i;
            if (!swatchNameExists(doc, candidate)) return candidate;
        }
        return baseName;
    }

    /**
     * カラーをスウォッチグループに追加する（毎回新しいスウォッチを作る）
     * @param {Document} doc - 対象ドキュメント
     * @param {SwatchGroup} swatchGroup - 追加先のスウォッチグループ
     * @param {Color} color - 追加するカラー
     * @returns {void}
     */
    function addSwatchToGroup(doc, swatchGroup, color) {
        if (color.typename === "SpotColor" || color.typename === "PatternColor" ||
            color.typename === "GradientColor" || color.typename === "NoColor") {
            return;
        }

        /* 1色の登録に失敗しても残りの色は登録する / A single failure must not drop the remaining colors */
        try {
            /* 同名の既存スウォッチを流用してはいけない。Illustrator ではスウォッチは1つのグループにしか
               属せないため、addSwatch() が既存スウォッチを元のグループから「移動」させてしまう。
               colorToName() は既定スウォッチと同じ "C=0 M=100 Y=100 K=0" 形式のため衝突しやすい。
               Never reuse an existing swatch: a swatch belongs to only one group, so addSwatch() would
               move the user's swatch out of its group. colorToName() collides with the default swatch names. */
            var swatch = doc.swatches.add();
            swatch.name = buildUniqueSwatchName(doc, colorToName(color));
            swatch.color = color;
            swatchGroup.addSwatch(swatch);
        } catch (err) {
            logError(err, "addSwatchToGroup");
        }
    }

    /**
     * カラー配列からスウォッチグループを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} groupName - スウォッチグループ名
     * @param {Array<object>} colors - { color, area } の配列
     * @returns {SwatchGroup} 作成したスウォッチグループ
     */
    function createSwatchGroupFromColors(doc, groupName, colors) {
        var swatchGroup = doc.swatchGroups.add();
        swatchGroup.name = groupName;
        for (var i = 0; i < colors.length; i++) {
            addSwatchToGroup(doc, swatchGroup, colors[i].color);
        }
        return swatchGroup;
    }

    /**
     * パレットのエントリをスウォッチ登録用のカラー配列に変換する
     * @param {Array<object>} rowEntries - { swatch: { color, area } } の配列
     * @param {boolean} useCmykAdjusted - CMYK補正した色にするかどうか
     * @returns {Array<object>} { color, area } の配列
     */
    function toSwatchColors(rowEntries, useCmykAdjusted) {
        var colors = [];
        for (var i = 0; i < rowEntries.length; i++) {
            var baseColor = rowEntries[i].swatch.color;
            var color = baseColor;
            if (useCmykAdjusted) color = buildCmykAdjustedColor(baseColor) || baseColor;
            colors.push({ color: color, area: rowEntries[i].swatch.area || 1 });
        }
        return colors;
    }

    /**
     * 5色・5色（CMYK補正）のスウォッチグループを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} baseName - グループ名の基準となる名前
     * @param {Array<object>} colors - { color, area } の配列
     * @param {object} outputOptions - 出力オプション
     * @returns {void}
     */
    function createSwatchGroupsFor5Only(doc, baseName, colors, outputOptions) {
        if (!outputOptions) return;

        var fiveRowEntries = buildAllPaletteRows(buildOrderedPaletteEntries(colors), outputOptions)[5];
        if (!fiveRowEntries || !fiveRowEntries.length) return;

        if (outputOptions.out5) {
            createSwatchGroupFromColors(doc, baseName + " - " + getLabel("checkbox.count5"),
                toSwatchColors(fiveRowEntries, false));
        }
        if (outputOptions.out5Adj) {
            createSwatchGroupFromColors(doc, baseName + " - " + getLabel("checkbox.count5Adjusted"),
                toSwatchColors(fiveRowEntries, true));
        }
    }

    // =========================================
    // パレットの描画 / Palette drawing
    // =========================================

    /**
     * ラベル用のフォントを適用する
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @returns {void}
     */
    function applyLabelFont(textRange) {
        for (var i = 0; i < LABEL_FONT_NAMES.length; i++) {
            try {
                textRange.characterAttributes.textFont = app.textFonts.getByName(LABEL_FONT_NAMES[i]);
                return;
            } catch (e) {
                /* 見つからなければ次の候補を試す / Try the next candidate */
            }
        }
    }

    /**
     * 色玉に添えるラベル文字列を作る
     * @param {Color} color - 対象のカラー
     * @param {string} labelMode - "hex" / "cmyk" / "both"
     * @returns {string} ラベル文字列
     */
    function buildColorLabelText(color, labelMode) {
        var hex = rgbToHex(colorToRGB(color));
        if (labelMode === "hex") return hex;

        var cmyk = colorToCMYKVals(color);
        if (!cmyk) return hex;

        var cmykText = getLabel("prefix.cmyk") +
            roundToStep(cmyk[0]) + ", " + roundToStep(cmyk[1]) + ", " +
            roundToStep(cmyk[2]) + ", " + roundToStep(cmyk[3]);

        return (labelMode === "cmyk") ? cmykText : (cmykText + "\r" + hex);
    }

    /**
     * 色玉の下にラベルを追加する
     * @param {Layer} targetLayer - テキストフレームを作るレイヤー
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} swatchRect - 対象の色玉
     * @param {number} squareSize - 色玉の一辺
     * @param {GroupItem} labelGroup - ラベルの移動先グループ
     * @param {string} labelMode - "hex" / "cmyk" / "both"
     * @returns {void}
     */
    function addPaletteLabel(targetLayer, doc, swatchRect, squareSize, labelGroup, labelMode) {
        if (!swatchRect || !swatchRect.filled) return;

        try {
            var fontSize = squareSize / LABEL_SIZE_RATIO;
            var labelFrame = targetLayer.textFrames.add();
            labelFrame.contents = buildColorLabelText(swatchRect.fillColor, labelMode || "both");
            labelFrame.textRange.justification = Justification.LEFT;
            labelFrame.textRange.characterAttributes.size = fontSize;
            labelFrame.textRange.fillColor = getLabelBlackColor(doc);
            applyLabelFont(labelFrame.textRange);
            labelFrame.position = [swatchRect.left, swatchRect.top - swatchRect.height - (fontSize / 2)];
            labelFrame.move(labelGroup, ElementPlacement.PLACEATEND);
        } catch (e) {
            logError(e, "palette label add");
        }
    }

    /**
     * パレット1行分の色玉を描画する
     * @param {GroupItem|Layer} container - 描画先のコンテナ
     * @param {string} rowName - 行のグループ名
     * @param {Array<object>} rowEntries - 行のエントリ配列
     * @param {number} rowLeft - 行の左端
     * @param {number} rowTop - 行の上端
     * @param {number} squareSize - 色玉の一辺
     * @param {number} gap - 色玉の間隔
     * @param {string|null} labelMode - ラベルの種類。null ならラベルなし
     * @param {Layer} targetLayer - ラベルを作るレイヤー
     * @param {Document} doc - 対象ドキュメント
     * @returns {GroupItem} 作成した行グループ
     */
    function drawPaletteRow(container, rowName, rowEntries, rowLeft, rowTop, squareSize, gap, labelMode, targetLayer, doc) {
        var rowGroup = container.groupItems.add();
        rowGroup.name = rowName;

        var cellCount = Math.max(1, rowEntries.length);
        for (var i = 0; i < cellCount; i++) {
            var swatchRect = rowGroup.pathItems.rectangle(rowTop, rowLeft + i * (squareSize + gap), squareSize, squareSize);
            swatchRect.stroked = false;
            swatchRect.filled = (rowEntries.length > 0);
            if (swatchRect.filled) {
                swatchRect.fillColor = rowEntries[i % rowEntries.length].swatch.color;
            }
            if (labelMode) addPaletteLabel(targetLayer, doc, swatchRect, squareSize, rowGroup, labelMode);
        }
        return rowGroup;
    }

    /**
     * 行内の色玉を CMYK補正した色に置き換える
     * @param {GroupItem} rowGroup - 対象の行グループ
     * @returns {void}
     */
    function applyAdjustedColorsToRow(rowGroup) {
        for (var i = 0; i < rowGroup.pathItems.length; i++) {
            var swatchRect = rowGroup.pathItems[i];
            if (!swatchRect.filled) continue;
            var adjustedColor = buildCmykAdjustedColor(swatchRect.fillColor);
            if (adjustedColor) swatchRect.fillColor = adjustedColor;
        }
    }

    /**
     * CMYK補正行のラベルを作り直す
     * @param {GroupItem} rowGroup - 対象の行グループ
     * @param {number} squareSize - 色玉の一辺
     * @param {object} outputOptions - 出力オプション
     * @param {Layer} targetLayer - ラベルを作るレイヤー
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function rebuildAdjustedRowLabels(rowGroup, squareSize, outputOptions, targetLayer, doc) {
        for (var i = rowGroup.textFrames.length - 1; i >= 0; i--) {
            rowGroup.textFrames[i].remove();
        }
        if (!outputOptions.showCMYK) return;

        for (var k = 0; k < rowGroup.pathItems.length; k++) {
            addPaletteLabel(targetLayer, doc, rowGroup.pathItems[k], squareSize, rowGroup, "cmyk");
        }
    }

    /**
     * 5色行から CMYK補正した行を作る
     * 「5色」がOFFで「5色（CMYK補正）」だけONのときは、複製せず基準行をそのまま補正する。
     * @param {GroupItem} baseRowGroup - 基準となる5色行
     * @param {boolean} duplicateRow - 複製して別の行を作るかどうか
     * @param {number} squareSize - 色玉の一辺
     * @param {object} outputOptions - 出力オプション
     * @param {Layer} targetLayer - ラベルを作るレイヤー
     * @param {Document} doc - 対象ドキュメント
     * @returns {GroupItem} CMYK補正した行
     */
    function buildAdjustedFiveRow(baseRowGroup, duplicateRow, squareSize, outputOptions, targetLayer, doc) {
        var adjustedRow = baseRowGroup;
        if (duplicateRow) {
            adjustedRow = baseRowGroup.duplicate();
            adjustedRow.translate(0, -(squareSize + squareSize * ADJUSTED_ROW_GAP_RATIO));
        }
        applyAdjustedColorsToRow(adjustedRow);
        rebuildAdjustedRowLabels(adjustedRow, squareSize, outputOptions, targetLayer, doc);
        return adjustedRow;
    }

    /**
     * 指定した色数の行を出力するかどうかを判定する
     * @param {number} rowCount - 行の色数
     * @param {object} outputOptions - 出力オプション
     * @returns {boolean} 出力する場合は true
     */
    function isRowEnabled(rowCount, outputOptions) {
        if (rowCount === 11) return !!outputOptions.out11;
        if (rowCount === 8) return !!outputOptions.out8;
        if (rowCount === 5) return !!(outputOptions.out5 || outputOptions.out5Adj);
        return true;
    }

    /**
     * 行のグループ名を作る
     * @param {number} rowCount - 行の色数
     * @returns {string} グループ名
     */
    function buildRowName(rowCount) {
        var suffix = getLabel("group.colorsSuffix");
        return (uiLang === "ja") ? (rowCount + suffix) : (rowCount + " " + suffix);
    }

    /**
     * 元オブジェクトの下にカラーパレットを描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem|object} originalItem - 配置の基準となるオブジェクトまたはアンカー
     * @param {Array<object>} colors - { color, area } の配列
     * @param {object} outputOptions - 出力オプション
     * @param {GroupItem} [containerGroup] - 描画先グループ。省略時はレイヤー直下
     * @returns {void}
     */
    function drawSwatchSquares(doc, originalItem, colors, outputOptions, containerGroup) {
        var sourceLeft = originalItem.left;
        var sourceWidth = originalItem.width;
        var sourceBottom = originalItem.top - originalItem.height;

        var rowPlan = buildPaletteRowPlan(sourceWidth, colors, outputOptions);
        var gap = rowPlan.gap;
        var rowGap = sourceWidth / PALETTE_ROW_GAP_DIVISOR;
        var firstRowTop = sourceBottom - rowPlan.squareSize;

        var targetLayer = originalItem.layer;
        var container = containerGroup || targetLayer;

        if (outputOptions.out16) {
            drawPaletteRow(container, getLabel("group.colors16"), rowPlan.rowsByCount[16],
                sourceLeft, firstRowTop, rowPlan.squareSize, gap, null, targetLayer, doc);
        }

        var rowCounts = [11, 8, 5];
        var prevBottom = outputOptions.out16 ? (firstRowTop - rowPlan.squareSize) : (firstRowTop + rowGap);

        for (var i = 0; i < rowCounts.length; i++) {
            var rowCount = rowCounts[i];
            if (!isRowEnabled(rowCount, outputOptions)) continue;

            var squareSize = (sourceWidth - gap * (rowCount - 1)) / rowCount;
            var rowTop = prevBottom - rowGap;

            /* 「5色」OFF＋「5色（CMYK補正）」ONのときは、基準行をそのまま補正行として使う / Reuse the base row when only the adjusted output is requested */
            var drawBaseRow = !(rowCount === 5 && !outputOptions.out5 && outputOptions.out5Adj);
            var labelMode = (rowCount === 5 && drawBaseRow && outputOptions.showHEX) ? "hex" : null;

            var rowGroup = drawPaletteRow(container, buildRowName(rowCount), rowPlan.rowsByCount[rowCount] || [],
                sourceLeft, rowTop, squareSize, gap, labelMode, targetLayer, doc);

            if (rowCount === 5 && outputOptions.out5Adj) {
                try {
                    buildAdjustedFiveRow(rowGroup, drawBaseRow, squareSize, outputOptions, targetLayer, doc);
                } catch (e) {
                    logError(e, "adjusted five row");
                }
            }

            prevBottom = rowTop - squareSize;
        }
    }

    /**
     * スウォッチ選択から固定サイズの色玉パレットを描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<object>} colors - { color, area } の配列
     * @returns {void}
     */
    function drawPaletteFromSwatches(doc, colors) {
        var squareSize = SWATCH_MODE_SQUARE_SIZE;
        var gap = squareSize * PALETTE_GAP_RATIO;
        var colorCount = colors.length;

        /* アクティブアートボードの中央に配置 / Place at the center of the active artboard */
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var totalWidth = colorCount * squareSize + (colorCount - 1) * gap;
        var rowLeft = artboardRect[0] + ((artboardRect[2] - artboardRect[0]) - totalWidth) / 2;
        var rowTop = artboardRect[1] - ((artboardRect[1] - artboardRect[3]) - squareSize) / 2;

        var targetLayer = doc.activeLayer;

        /* 代表色の選出とは違い、スウォッチパネルで選んだ順のまま並べる / Unlike the extracted palettes, keep the order the swatches were selected in */
        var rowEntries = [];
        for (var i = 0; i < colorCount; i++) {
            rowEntries.push({ swatch: colors[i] });
        }

        drawPaletteRow(targetLayer, getLabel("group.swatchPalette"), rowEntries,
            rowLeft, rowTop, squareSize, gap, isCmykDocument(doc) ? "both" : "hex", targetLayer, doc);

        try {
            createSwatchGroupFromColors(doc, getLabel("group.swatchPalette"), colors);
        } catch (e) {
            logError(e, "swatch group registration");
        }

        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialogs
    // =========================================

    /* セッション中だけ保持するダイアログ位置 / Dialog position remembered for this session only */
    var outputDialogBounds = null;

    /**
     * プリセット名を照合用に正規化する（角カッコ・空白を除き、小文字と半角数字にそろえる）
     * @param {string} presetName - プリセット名
     * @returns {string} 正規化した名前
     */
    function normalizePresetName(presetName) {
        var normalized = String(presetName).replace(/[\[\]\s]/g, "").toLowerCase();
        return normalized.replace(/[０-９]/g, function (fullWidthDigit) {
            return String.fromCharCode(fullWidthDigit.charCodeAt(0) - 0xFEE0);
        });
    }

    /**
     * 既定のトレースプリセットが一覧の何番目にあるかを返す
     * @param {Array<string>} presetNames - プリセット名の配列
     * @returns {number} 見つかった位置。無ければ -1
     */
    function findDefaultPresetIndex(presetNames) {
        for (var i = 0; i < DEFAULT_TRACING_PRESETS.length; i++) {
            var wanted = normalizePresetName(DEFAULT_TRACING_PRESETS[i]);
            for (var k = 0; k < presetNames.length; k++) {
                if (normalizePresetName(presetNames[k]) === wanted) return k;
            }
        }
        return -1;
    }

    /**
     * 利用できる既定のトレースプリセット名を返す
     * @returns {string|null} プリセット名。見つからない場合は null
     */
    function getDefaultTracingPresetName() {
        var presets = null;
        try {
            presets = app.tracingPresetsList;
        } catch (e) {
            logError(e, "tracing presets");
            return null;
        }
        if (!presets || !presets.length) return null;

        var index = findDefaultPresetIndex(presets);
        return (index >= 0) ? presets[index] : null;
    }

    /**
     * トレースプリセットを標準（[]付き）とユーザー定義に分類する
     * @param {Array<string>} presets - プリセット名の配列
     * @returns {object} { builtIn: Array<string>, custom: Array<string> }
     */
    function categorizePresets(presets) {
        var builtIn = [];
        var custom = [];
        for (var i = 0; i < presets.length; i++) {
            if (presets[i].charAt(0) === "[") {
                builtIn.push(presets[i]);
            } else {
                custom.push(presets[i]);
            }
        }
        return { builtIn: builtIn, custom: custom };
    }

    /**
     * 見出し付きのプリセット一覧を追加する
     * @param {Group} parent - 追加先の行グループ
     * @param {string} headerText - 一覧の見出し
     * @param {Array<string>} presetNames - プリセット名の配列
     * @param {string} tooltipText - ツールチップ
     * @returns {ListBox} 追加した一覧
     */
    function addPresetList(parent, headerText, presetNames, tooltipText) {
        var column = addColumnGroup(parent);
        column.add("statictext", undefined, headerText);

        var presetList = column.add("listbox", undefined, presetNames);
        presetList.preferredSize = PRESET_LIST_SIZE;
        presetList.helpTip = tooltipText;
        return presetList;
    }

    /**
     * トレースプリセット選択ダイアログを表示する
     * @param {Array<string>} presets - プリセット名の配列
     * @returns {string|null} 選択したプリセット名。キャンセル時は null
     */
    function showPresetDialog(presets) {
        var presetCategories = categorizePresets(presets);

        var presetDialog = new Window("dialog", getLabel("dialog.preset"));
        setupWindow(presetDialog);

        var presetListRow = presetDialog.add("group");
        setupRow(presetListRow, "fill", COLUMN_SPACING);
        presetListRow.alignChildren = ["fill", "fill"];

        var builtInPresetList = addPresetList(presetListRow, getLabel("listHeader.presetBuiltIn"),
            presetCategories.builtIn, getLabel("tooltip.presetBuiltIn"));
        var customPresetList = addPresetList(presetListRow, getLabel("listHeader.presetCustom"),
            presetCategories.custom, getLabel("tooltip.presetCustom"));

        /* 2つのリストで同時に選択させない / Keep only one list selected at a time */
        builtInPresetList.onChange = function () {
            if (builtInPresetList.selection) customPresetList.selection = null;
        };
        customPresetList.onChange = function () {
            if (customPresetList.selection) builtInPresetList.selection = null;
        };

        /* 既定プリセット（6色変換）を標準・ユーザー定義の順に探し、無ければ最後のユーザー定義を選ぶ
           Look for the default preset (6 Colors) in the built-in list then the custom one, otherwise fall back to the last custom preset */
        var defaultBuiltInIndex = findDefaultPresetIndex(presetCategories.builtIn);
        var defaultCustomIndex = findDefaultPresetIndex(presetCategories.custom);
        if (defaultBuiltInIndex >= 0) {
            builtInPresetList.selection = defaultBuiltInIndex;
        } else if (defaultCustomIndex >= 0) {
            customPresetList.selection = defaultCustomIndex;
        } else if (presetCategories.custom.length > 0) {
            customPresetList.selection = presetCategories.custom.length - 1;
        } else if (presetCategories.builtIn.length > 0) {
            builtInPresetList.selection = 0;
        }

        var presetButtonRow = presetDialog.add("group");
        setupRow(presetButtonRow, ["center", "top"]);
        presetButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        presetButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        if (presetDialog.show() !== 1) return null;
        if (builtInPresetList.selection) return builtInPresetList.selection.text;
        if (customPresetList.selection) return customPresetList.selection.text;
        return null;
    }

    /**
     * 出力オプションダイアログを表示する
     * @param {function} onPreviewChange - 設定変更時に呼ばれるプレビュー更新関数
     * @param {function} onFitView - 「画面にフィット」がONのときに呼ばれる関数
     * @returns {object|string|null} 出力オプション、"__RETRY__"、キャンセル時は null
     */
    function showOutputOptionsDialog(onPreviewChange, onFitView) {
        var outputDialog = new Window("dialog", getLabel("dialog.output") + " " + SCRIPT_VERSION);
        setupWindow(outputDialog);

        outputDialog.onMove = outputDialog.onResize = function () {
            outputDialogBounds = outputDialog.bounds;
        };

        var rowScopeRow = outputDialog.add("group");
        setupRow(rowScopeRow, "center");
        var allRowsRadio = rowScopeRow.add("radiobutton", undefined, getLabel("radio.allRows"));
        allRowsRadio.helpTip = getLabel("tooltip.allRows");
        var fiveRowsOnlyRadio = rowScopeRow.add("radiobutton", undefined, getLabel("radio.fiveRowsOnly"));
        fiveRowsOnlyRadio.helpTip = getLabel("tooltip.fiveRowsOnly");

        var optionColumnsRow = outputDialog.add("group");
        setupRow(optionColumnsRow, "fill", COLUMN_SPACING);
        optionColumnsRow.alignChildren = ["fill", "top"];

        var outputRowsColumn = addColumnGroup(optionColumnsRow);
        var colorInfoColumn = addColumnGroup(optionColumnsRow);

        var outputRowsPanel = addPanel(outputRowsColumn, getLabel("panel.outputRows"), 6);
        var count16Checkbox = addLeftCheckbox(outputRowsPanel, getLabel("checkbox.count16"), getLabel("tooltip.count16"));
        var count11Checkbox = addLeftCheckbox(outputRowsPanel, getLabel("checkbox.count11"), getLabel("tooltip.count11"));
        var count8Checkbox = addLeftCheckbox(outputRowsPanel, getLabel("checkbox.count8"));
        var count5Checkbox = addLeftCheckbox(outputRowsPanel, getLabel("checkbox.count5"));
        var count5AdjustedCheckbox = addLeftCheckbox(outputRowsPanel, getLabel("checkbox.count5Adjusted"), getLabel("tooltip.count5Adjusted"));

        var colorInfoPanel = addPanel(colorInfoColumn, getLabel("panel.colorInfo"), 6);
        var hexCheckbox = addLeftCheckbox(colorInfoPanel, getLabel("checkbox.hex"), getLabel("tooltip.hexLabel"));
        var cmykCheckbox = addLeftCheckbox(colorInfoPanel, getLabel("checkbox.cmyk"), getLabel("tooltip.cmykLabel"));

        /* CMYKラベルは CMYK ドキュメントでのみ意味を持つ / CMYK labels only make sense in a CMYK document */
        var canUseCmykLabels = isCmykDocument(app.activeDocument);
        if (!canUseCmykLabels) cmykCheckbox.helpTip = getLabel("tooltip.cmykDocOnly");

        var fitViewRow = outputDialog.add("group");
        setupRow(fitViewRow, "left");
        fitViewRow.margins = [20, 0, 0, 0];
        var fitViewCheckbox = fitViewRow.add("checkbox", undefined, getLabel("checkbox.fitView"));
        fitViewCheckbox.value = true;
        fitViewCheckbox.helpTip = getLabel("tooltip.fitView");

        /* 「5色の行のみ」で伏せる行と、状態を保存する対象 / Rows hidden by "5-color rows only", and the checkboxes whose state is saved */
        var wideRowCheckboxes = [count16Checkbox, count11Checkbox, count8Checkbox];
        var rowCheckboxes = [count16Checkbox, count11Checkbox, count8Checkbox, count5Checkbox, count5AdjustedCheckbox];
        var stateCheckboxes = [count16Checkbox, count11Checkbox, count8Checkbox, count5Checkbox,
            count5AdjustedCheckbox, hexCheckbox, cmykCheckbox];

        for (var i = 0; i < stateCheckboxes.length; i++) stateCheckboxes[i].value = true;
        allRowsRadio.value = false;
        fiveRowsOnlyRadio.value = true;

        var isInitializing = true;
        var savedRowState = null;

        /**
         * 現在の設定から出力オプションを組み立てる
         * @returns {object} 出力オプション
         */
        function getCurrentOptions() {
            return {
                out16: count16Checkbox.value,
                out11: count11Checkbox.value,
                out8: count8Checkbox.value,
                out5: count5Checkbox.value,
                out5Adj: count5AdjustedCheckbox.value,
                showHEX: hexCheckbox.value,
                showCMYK: cmykCheckbox.value,
                cascade: true
            };
        }

        /**
         * 「画面にフィット」がONのときだけビューを合わせる
         * @returns {void}
         */
        function applyFitView() {
            if (!onFitView || !fitViewCheckbox.value) return;
            try {
                onFitView();
            } catch (e) {
                logError(e, "fit view");
            }
        }

        /**
         * プレビュー更新を通知する
         * @returns {void}
         */
        function notifyPreviewChange() {
            /* 初期化中の重複描画を避ける / Avoid redundant redraws while initializing */
            if (isInitializing || !onPreviewChange) return;
            try {
                onPreviewChange(getCurrentOptions());
            } catch (e) {
                logError(e, "notify preview");
            }
            applyFitView();
        }

        /**
         * カラー情報チェックボックスの有効／無効を更新する
         * @param {boolean} doNotify - プレビュー更新を通知するかどうか
         * @returns {void}
         */
        function updateColorInfoAvailability(doNotify) {
            hexCheckbox.enabled = count5Checkbox.value;
            if (!count5Checkbox.value) hexCheckbox.value = false;

            cmykCheckbox.enabled = (count5AdjustedCheckbox.value && canUseCmykLabels);
            if (!count5AdjustedCheckbox.value || !canUseCmykLabels) cmykCheckbox.value = false;

            if (doNotify !== false) notifyPreviewChange();
        }

        /**
         * 「すべての行」／「5色の行のみ」の切り替えを反映する
         * @param {boolean} doNotify - プレビュー更新を通知するかどうか
         * @returns {void}
         */
        function updateRowScope(doNotify) {
            var i;
            if (fiveRowsOnlyRadio.value) {
                /* 5色の行のみ: 他の行をOFFにして操作できないようにする / 5-rows-only: force the other rows off and disable them */
                savedRowState = [];
                for (i = 0; i < stateCheckboxes.length; i++) savedRowState.push(stateCheckboxes[i].value);
                for (i = 0; i < wideRowCheckboxes.length; i++) {
                    wideRowCheckboxes[i].value = false;
                    wideRowCheckboxes[i].enabled = false;
                }
            } else {
                for (i = 0; i < wideRowCheckboxes.length; i++) wideRowCheckboxes[i].enabled = true;
                if (savedRowState) {
                    for (i = 0; i < stateCheckboxes.length; i++) stateCheckboxes[i].value = !!savedRowState[i];
                }
                updateColorInfoAvailability(false);
            }

            if (doNotify !== false) notifyPreviewChange();
        }

        allRowsRadio.onClick = function () { updateRowScope(true); };
        fiveRowsOnlyRadio.onClick = function () { updateRowScope(true); };

        fitViewCheckbox.onClick = applyFitView;

        count16Checkbox.onClick = notifyPreviewChange;
        count11Checkbox.onClick = notifyPreviewChange;
        count8Checkbox.onClick = notifyPreviewChange;
        hexCheckbox.onClick = notifyPreviewChange;
        cmykCheckbox.onClick = notifyPreviewChange;

        count5Checkbox.onClick = function () {
            updateColorInfoAvailability(false);
            notifyPreviewChange();
        };
        count5AdjustedCheckbox.onClick = function () {
            updateColorInfoAvailability(false);
            /* 5色（CMYK補正）をONにしたらCMYKラベルも自動でONにする / Turn CMYK labels on together with the adjusted row */
            if (count5AdjustedCheckbox.value && cmykCheckbox.enabled) cmykCheckbox.value = true;
            notifyPreviewChange();
        };

        /* 最下部のボタン列: 左＝選び直す／右＝キャンセル・OK / Bottom button row: reselect on the left, cancel & OK on the right */
        var dialogButtonRow = outputDialog.add("group");
        setupRow(dialogButtonRow, "fill");

        var retryButton = dialogButtonRow.add("button", undefined, getLabel("button.retry"));
        retryButton.preferredSize = DIALOG_BUTTON_SIZE;
        retryButton.helpTip = getLabel("tooltip.retry");
        retryButton.onClick = function () {
            outputDialogBounds = outputDialog.bounds;
            outputDialog.close(2);
        };

        var buttonRowSpacer = dialogButtonRow.add("group");
        buttonRowSpacer.alignment = ["fill", "fill"];

        var cancelButton = dialogButtonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        cancelButton.preferredSize = DIALOG_BUTTON_SIZE;
        var okButton = dialogButtonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        okButton.preferredSize = DIALOG_BUTTON_SIZE;

        okButton.onClick = function () {
            for (var i = 0; i < rowCheckboxes.length; i++) {
                if (!rowCheckboxes[i].value) continue;
                outputDialogBounds = outputDialog.bounds;
                outputDialog.close(1);
                return;
            }
            alert(getLabel("alert.noRow"));
        };
        cancelButton.onClick = function () {
            outputDialogBounds = outputDialog.bounds;
            outputDialog.close(0);
        };

        /* すべてのコントロールを作ってから初期状態を反映する / Apply the initial state after every control exists */
        updateColorInfoAvailability(false);
        updateRowScope(false);
        isInitializing = false;
        notifyPreviewChange();

        /* 前回位置を復元する（サイズは復元しない） / Restore the previous position only, not the size */
        if (outputDialogBounds) {
            outputDialog.location = [outputDialogBounds.x, outputDialogBounds.y];
        } else {
            outputDialog.center();
        }

        var dialogResult = outputDialog.show();
        outputDialogBounds = outputDialog.bounds;

        if (dialogResult === 2) return "__RETRY__";
        if (dialogResult !== 1) return null;
        return getCurrentOptions();
    }

    /**
     * 進捗ウィンドウを作る
     * @param {number} maxValue - 進捗バーの最大値
     * @returns {object} { window: Window, set: function, close: function }
     */
    function createProgressWindow(maxValue) {
        var progressWindow = new Window("palette", getLabel("dialog.progress"));
        setupWindow(progressWindow);

        var progressText = progressWindow.add("statictext", undefined, getLabel("progress.preparing"));
        progressText.characters = PROGRESS_TEXT_CHARS;
        progressText.alignment = ["fill", "center"];

        var progressBar = progressWindow.add("progressbar", undefined, 0, Math.max(1, maxValue));
        progressBar.preferredSize = PROGRESS_BAR_SIZE;

        progressWindow.show();
        progressWindow.update();

        return {
            window: progressWindow,
            set: function (value, labelText) {
                try {
                    if (labelText) progressText.text = labelText;
                    progressBar.value = Math.max(0, Math.min(progressBar.maxvalue, value));
                    progressWindow.update();
                    app.redraw();
                } catch (e) {
                    logError(e, "progress update");
                }
            },
            close: function () {
                try {
                    progressWindow.close();
                } catch (e) {
                    logError(e, "progress close");
                }
            }
        };
    }

    // =========================================
    // 処理対象の組み立て / Building the task list
    // =========================================

    /**
     * 座標範囲からパレット配置用のアンカーを作る
     * @param {Array<number>} bounds - [左, 上, 右, 下]
     * @param {Layer} layer - 対象のレイヤー
     * @returns {object} パレット配置用のアンカー
     */
    function createBoundsAnchor(bounds, layer) {
        return {
            typename: "BoundsAnchor",
            left: bounds[0],
            top: bounds[1],
            width: bounds[2] - bounds[0],
            height: bounds[1] - bounds[3],
            layer: layer,
            geometricBounds: bounds
        };
    }

    /**
     * 複数オブジェクト全体の境界からパレット配置用アンカーを作る
     * @param {Array<PageItem>} items - 対象のオブジェクト配列
     * @param {PageItem} fallbackItem - 境界が取れない場合に返すオブジェクト
     * @returns {object|PageItem} アンカー、または fallbackItem
     */
    function buildPaletteAnchorFromItems(items, fallbackItem) {
        if (!items || !items.length || !fallbackItem) return fallbackItem;

        var left = Infinity;
        var top = -Infinity;
        var right = -Infinity;
        var bottom = Infinity;
        var found = false;

        for (var i = 0; i < items.length; i++) {
            try {
                var bounds = items[i].geometricBounds;
                if (!bounds || bounds.length < 4) continue;
                if (bounds[0] < left) left = bounds[0];
                if (bounds[1] > top) top = bounds[1];
                if (bounds[2] > right) right = bounds[2];
                if (bounds[3] < bottom) bottom = bounds[3];
                found = true;
            } catch (e) {
                logError(e, "palette anchor bounds");
            }
        }

        if (!found) return fallbackItem;
        return createBoundsAnchor([left, top, right, bottom], fallbackItem.layer);
    }

    /**
     * クリップグループのクリッピングパスの境界を取得する
     * @param {GroupItem} clipGroup - 対象のクリップグループ
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getClippingBounds(clipGroup) {
        try {
            for (var i = 0; i < clipGroup.pageItems.length; i++) {
                if (clipGroup.pageItems[i].clipping) return clipGroup.pageItems[i].geometricBounds;
            }
        } catch (e) {
            logError(e, "clip bounds");
        }
        return clipGroup.geometricBounds;
    }

    /**
     * ユーザー設定に従ったラスタライズ設定を作る
     * @returns {RasterizeOptions} ラスタライズ設定
     */
    function createRasterizeOptions() {
        var rasterizeOptions = new RasterizeOptions();
        rasterizeOptions.resolution = RASTERIZE_RESOLUTION;
        rasterizeOptions.transparency = RASTERIZE_TRANSPARENCY;
        rasterizeOptions.padding = RASTERIZE_PADDING;
        rasterizeOptions.antiAliasing = RASTERIZE_ANTIALIAS;
        rasterizeOptions.backgroundBlack = false;
        return rasterizeOptions;
    }

    /**
     * 選択内のクリップグループをラスタライズし、処理対象の一覧を返す
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} selection - 選択中のオブジェクト
     * @returns {Array<object>} { item: PageItem, anchor: object|null } の配列
     */
    function rasterizeClippedGroups(doc, selection) {
        var sourceEntries = [];
        var rasterizedCount = 0;

        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename !== "GroupItem" || !selection[i].clipped) {
                sourceEntries.push({ item: selection[i], anchor: null });
                continue;
            }

            /* クリップ範囲はパレット配置の基準として残す / Keep the clip bounds as the palette anchor */
            var clipBounds = getClippingBounds(selection[i]);
            var rasterizedItem = doc.rasterize(selection[i], clipBounds, createRasterizeOptions());
            rasterizedCount++;
            sourceEntries.push({ item: rasterizedItem, anchor: createBoundsAnchor(clipBounds, rasterizedItem.layer) });
        }

        if (rasterizedCount > 0) {
            var newSelection = [];
            for (var k = 0; k < sourceEntries.length; k++) newSelection.push(sourceEntries[k].item);
            app.selection = newSelection;
        }
        return sourceEntries;
    }

    /**
     * 処理対象を「何から色を取るか」と「どこにパレットを置くか」に分けて組み立てる
     * ラスター／配置画像は個別に、ベクターとテキストはまとめて1件として扱う。
     * @param {Array<object>} sourceEntries - rasterizeClippedGroups() の戻り値
     * @returns {object} { tasks, vectorItems, textItems }
     */
    function buildExtractionPlan(sourceEntries) {
        var paletteTasks = [];
        var vectorItems = [];
        var textItems = [];

        for (var i = 0; i < sourceEntries.length; i++) {
            var item = sourceEntries[i].item;
            var typeName = item.typename;

            if (typeName === "PlacedItem" || typeName === "RasterItem") {
                /* アンカーがある場合、配置はアンカー・複製元は実アイテム / With an anchor, place by the anchor but duplicate the real item */
                paletteTasks.push({
                    type: "raster",
                    originalItem: sourceEntries[i].anchor || item,
                    workSource: sourceEntries[i].anchor ? item : null
                });
            } else if (typeName === "PathItem" || typeName === "CompoundPathItem" || typeName === "GroupItem") {
                vectorItems.push(item);
            } else if (typeName === "TextFrame") {
                textItems.push(item);
            }
        }

        if (vectorItems.length > 0 || textItems.length > 0) {
            var anchorItems = vectorItems.concat(textItems);
            var anchorBase = (vectorItems.length > 0) ? vectorItems[0] : textItems[0];
            paletteTasks.push({ type: "vector", originalItem: buildPaletteAnchorFromItems(anchorItems, anchorBase) });
        }

        return { tasks: paletteTasks, vectorItems: vectorItems, textItems: textItems };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ベクターとテキストを作業用レイヤーに複製し、1つのグループにまとめる
     * テキストは複製後にアウトライン化する。
     * @param {Layer} workLayer - 作業用レイヤー
     * @param {Array<PageItem>} vectorItems - ベクターオブジェクトの配列
     * @param {Array<TextFrame>} textItems - テキストフレームの配列
     * @returns {GroupItem} 複製をまとめたグループ
     */
    function duplicateVectorsAndText(workLayer, vectorItems, textItems) {
        var vectorGroup = workLayer.groupItems.add();

        for (var i = vectorItems.length - 1; i >= 0; i--) {
            vectorItems[i].duplicate().move(vectorGroup, ElementPlacement.PLACEATEND);
        }

        for (var k = textItems.length - 1; k >= 0; k--) {
            var duplicatedText = textItems[k].duplicate();
            duplicatedText.move(vectorGroup, ElementPlacement.PLACEATEND);

            try {
                /* createOutline() の戻り値を明示的に扱う。環境によっては新規オブジェクトが返り、
                   元の TextFrame が残ることがあるため、戻り値を無視しない。
                   失敗した場合は複製テキストを残して続行する（collectFillColors() は TextFrame を拾わないため色抽出の対象外になる）。 */
                var outlinedText = duplicatedText.createOutline();
                if (outlinedText) outlinedText.move(vectorGroup, ElementPlacement.PLACEATEND);
            } catch (e) {
                logError(e, "text outline");
            }
        }

        return vectorGroup;
    }

    /**
     * 処理対象を作業用レイヤーに複製する
     * @param {Layer} workLayer - 作業用レイヤー
     * @param {object} extractionPlan - buildExtractionPlan() の戻り値
     * @returns {Array<object>} { type, originalItem, workItem } の配列
     */
    function duplicateTasksToWorkLayer(workLayer, extractionPlan) {
        var workTasks = [];

        for (var i = 0; i < extractionPlan.tasks.length; i++) {
            var paletteTask = extractionPlan.tasks[i];
            if (paletteTask.type === "vector") {
                workTasks.push({
                    type: "vector",
                    originalItem: paletteTask.originalItem,
                    workItem: duplicateVectorsAndText(workLayer, extractionPlan.vectorItems, extractionPlan.textItems)
                });
            } else {
                var duplicatedItem = (paletteTask.workSource || paletteTask.originalItem).duplicate();
                duplicatedItem.move(workLayer, ElementPlacement.PLACEATEND);
                workTasks.push({ type: "raster", originalItem: paletteTask.originalItem, workItem: duplicatedItem });
            }
        }
        return workTasks;
    }

    /**
     * ラスター／配置画像をトレースして拡張する
     * @param {PageItem} itemToTrace - トレース対象のオブジェクト
     * @param {string|null} tracingPresetName - 適用するプリセット名
     * @returns {GroupItem|null} 拡張結果。失敗時は null
     */
    function traceAndExpand(itemToTrace, tracingPresetName) {
        try {
            var traceObject = itemToTrace.trace();
            if (tracingPresetName) {
                try {
                    traceObject.tracing.tracingOptions.loadFromPreset(tracingPresetName);
                } catch (e) {
                    logError(e, "trace preset");
                }
            }
            return traceObject.tracing.expandTracing();
        } catch (e) {
            logError(e, "trace expand");
            return null;
        }
    }

    /**
     * 処理対象からカラーを抽出する
     * ベクターは直接、ラスター／配置画像はトレース→拡張してから抽出する。
     * ベクターに読めない塗りが含まれていた場合は、ラスタライズしての抽出に切り替える。
     * @param {Document} doc - 対象ドキュメント
     * @param {object} workTask - { type, originalItem, workItem }
     * @param {string|null} tracingPresetName - 適用するプリセット名
     * @returns {Array<object>|null} { color, area } の配列。抽出できない場合は null
     */
    function extractTaskColors(doc, workTask, tracingPresetName) {
        if (workTask.type !== "vector") {
            var expanded = traceAndExpand(workTask.workItem, tracingPresetName);
            return expanded ? deduplicateColors(collectFillColors(expanded, [])) : null;
        }

        var stats = { unreadableCount: 0 };
        var colors = deduplicateColors(collectFillColors(workTask.workItem, [], stats));
        if (stats.unreadableCount === 0) return colors;

        /* 読めない塗りがあったぶんは色が欠けるので、複製をラスタライズしてトレースし直す
           Some fills could not be read, so rasterize the duplicate and trace it instead */
        var fallbackColors = extractColorsByRasterizing(doc, workTask.workItem, tracingPresetName);
        return (fallbackColors && fallbackColors.length > 0) ? fallbackColors : colors;
    }

    /**
     * 作業用レイヤーの複製をラスタライズし、トレースしてカラーを抽出する
     * フリーグラデーションなど、DOM から塗りを読めないオブジェクトの受け皿。
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} workItem - 作業用レイヤー上の複製
     * @param {string|null} tracingPresetName - 適用するプリセット名
     * @returns {Array<object>|null} { color, area } の配列。失敗した場合は null
     */
    function extractColorsByRasterizing(doc, workItem, tracingPresetName) {
        var rasterizedItem;
        try {
            rasterizedItem = doc.rasterize(workItem, workItem.geometricBounds, createRasterizeOptions());
        } catch (e) {
            logError(e, "vector fallback rasterize");
            return null;
        }

        /* ベクター選択ではプリセット選択ダイアログを出さないため、既定プリセットで補う
           The preset dialog is skipped for a vector selection, so fall back to the default preset */
        var presetName = tracingPresetName || getDefaultTracingPresetName();
        var expanded = traceAndExpand(rasterizedItem, presetName);
        return expanded ? deduplicateColors(collectFillColors(expanded, [])) : null;
    }

    /**
     * パレットのグループ名を決める
     * @param {PageItem|object} originalItem - 配置の基準となるオブジェクト
     * @param {number} taskIndex - 処理対象の連番（0起点）
     * @returns {string} グループ名
     */
    function buildPaletteGroupName(originalItem, taskIndex) {
        if (originalItem.typename === "PlacedItem" && originalItem.file) return originalItem.file.name;
        return getLabel("fallbackName.itemPrefix") + (taskIndex + 1);
    }

    /**
     * パレット描画用の空グループを作る
     * @param {PageItem|object} originalItem - 配置の基準となるオブジェクト
     * @param {string} groupName - グループ名
     * @returns {GroupItem|null} 作成したグループ。失敗時は null
     */
    function createPaletteContainer(originalItem, groupName) {
        try {
            var paletteContainer = originalItem.layer.groupItems.add();
            paletteContainer.name = groupName;
            return paletteContainer;
        } catch (e) {
            logError(e, "palette group create");
            return null;
        }
    }

    /**
     * グループを安全に削除する
     * @param {GroupItem} group - 削除するグループ
     * @returns {void}
     */
    function removeGroup(group) {
        if (!group) return;
        try {
            group.remove();
        } catch (e) {
            logError(e, "group remove");
        }
    }

    /**
     * 1件分のスウォッチ登録とパレット描画を行う
     * @param {Document} doc - 対象ドキュメント
     * @param {object} workTask - { type, originalItem, workItem }
     * @param {number} taskIndex - 処理対象の連番（0起点）
     * @param {Array<object>} colors - { color, area } の配列
     * @param {object} outputOptions - 出力オプション
     * @returns {void}
     */
    function outputPaletteForTask(doc, workTask, taskIndex, colors, outputOptions) {
        /* スウォッチ登録は5色・5色（CMYK補正）だけを対象にする / Register swatches only for the 5-color rows */
        createSwatchGroupsFor5Only(doc, buildPaletteGroupName(workTask.originalItem, taskIndex), colors, outputOptions);

        var paletteContainer = createPaletteContainer(workTask.originalItem, PALETTE_GROUP_NAME);
        try {
            drawSwatchSquares(doc, workTask.originalItem, colors, outputOptions, paletteContainer);
        } catch (e) {
            logError(e, "palette draw");
        }
    }

    /**
     * 対象オブジェクトとプレビューが収まるようにビューを合わせる
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} items - 対象のオブジェクト配列
     * @returns {void}
     */
    function fitViewToItems(doc, items) {
        if (!items.length) return;

        var left = Infinity;
        var top = -Infinity;
        var right = -Infinity;
        var bottom = Infinity;
        for (var i = 0; i < items.length; i++) {
            var bounds = items[i].geometricBounds;
            if (bounds[0] < left) left = bounds[0];
            if (bounds[1] > top) top = bounds[1];
            if (bounds[2] > right) right = bounds[2];
            if (bounds[3] < bottom) bottom = bounds[3];
        }

        var activeView = doc.activeView;
        activeView.centerPoint = [(left + right) / 2, (top + bottom) / 2];

        var viewBounds = activeView.bounds;
        var itemWidth = right - left;
        var itemHeight = top - bottom;
        if (itemWidth > 0 && itemHeight > 0) {
            var scale = Math.min(
                (viewBounds[2] - viewBounds[0]) / itemWidth,
                (viewBounds[1] - viewBounds[3]) / itemHeight
            ) / FIT_VIEW_MARGIN;
            activeView.zoom = activeView.zoom * scale;
        }
        app.redraw();
    }

    /**
     * プレビューを表示しながら出力オプションを尋ねる
     * @param {Document} doc - 対象ドキュメント
     * @param {object} workTask - { type, originalItem, workItem }
     * @param {Array<object>} colors - { color, area } の配列
     * @param {object} progress - 進捗ウィンドウ
     * @returns {object} { result: "ok"|"retry"|"cancel", options, previewGroup }
     */
    function askOutputOptions(doc, workTask, colors, progress) {
        var previewOptions = {
            out16: true, out11: true, out8: true, out5: true, out5Adj: true,
            showHEX: true, showCMYK: true, cascade: true
        };
        var previewGroup = createPaletteContainer(workTask.originalItem, PREVIEW_GROUP_NAME);

        try {
            if (previewGroup) drawSwatchSquares(doc, workTask.originalItem, colors, previewOptions, previewGroup);
        } catch (e) {
            logError(e, "preview initial draw");
        }

        /* モーダルを開く前にプレビューを描き切らせる / Make sure the preview is rendered before the modal opens */
        app.redraw();
        $.sleep(80);

        progress.window.hide();
        var outputOptions = showOutputOptionsDialog(function (currentOptions) {
            /* 取り残しを防ぐため、プレビューグループは毎回作り直す / Recreate the preview group each time so nothing lingers */
            removeGroup(previewGroup);
            previewGroup = createPaletteContainer(workTask.originalItem, PREVIEW_GROUP_NAME);
            if (previewGroup) {
                try {
                    drawSwatchSquares(doc, workTask.originalItem, colors, currentOptions, previewGroup);
                } catch (e) {
                    logError(e, "preview redraw");
                }
            }
            app.redraw();
        }, function () {
            var items = [workTask.originalItem];
            if (previewGroup) items.push(previewGroup);
            fitViewToItems(doc, items);
        });
        progress.window.show();

        if (outputOptions === "__RETRY__" || !outputOptions) {
            removeGroup(previewGroup);
            return { result: (outputOptions === "__RETRY__") ? "retry" : "cancel", options: null, previewGroup: null };
        }

        /* 最終出力を描き終えるまでプレビューは残す / Keep the preview until the final output is drawn */
        return { result: "ok", options: outputOptions, previewGroup: previewGroup };
    }

    /**
     * 処理対象を順に処理する
     * 最初にカラーを取得できた時点で出力オプションを尋ね、以降はその設定を使う。
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<object>} workTasks - duplicateTasksToWorkLayer() の戻り値
     * @param {object} progress - 進捗ウィンドウ
     * @param {string|null} tracingPresetName - 適用するプリセット名
     * @returns {string} "done" または "retry"
     */
    function processPaletteTasks(doc, workTasks, progress, tracingPresetName) {
        var outputOptions = null;

        for (var i = 0; i < workTasks.length; i++) {
            try {
                var workTask = workTasks[i];

                progress.set(i * 2 + 1, getLabel("progress.tracing"));
                var colors = extractTaskColors(doc, workTask, tracingPresetName);

                progress.set(i * 2 + 2, getLabel("progress.palette"));
                if (!colors || colors.length === 0) continue;

                if (outputOptions !== null) {
                    outputPaletteForTask(doc, workTask, i, colors, outputOptions);
                    continue;
                }

                var optionsChoice = askOutputOptions(doc, workTask, colors, progress);
                if (optionsChoice.result !== "ok") {
                    return (optionsChoice.result === "retry") ? "retry" : "done";
                }

                outputOptions = optionsChoice.options;
                outputPaletteForTask(doc, workTask, i, colors, outputOptions);
                removeGroup(optionsChoice.previewGroup);
            } catch (e) {
                logError(e, "per-item processing");
            }
        }
        return "done";
    }

    /**
     * ラスター／配置画像の処理対象が含まれるかを判定する
     * @param {Array<object>} paletteTasks - 処理対象の配列
     * @returns {boolean} 含まれる場合は true
     */
    function hasRasterTask(paletteTasks) {
        for (var i = 0; i < paletteTasks.length; i++) {
            if (paletteTasks[i].type === "raster") return true;
        }
        return false;
    }

    /**
     * プリセット選択からパレット出力までを1回分実行する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} extractionPlan - buildExtractionPlan() の戻り値
     * @returns {string} "done" または "retry"
     */
    function runPaletteSession(doc, extractionPlan) {
        var tracingPresetName = null;

        /* プリセット選択はラスター／配置画像があるときだけ尋ねる / Ask for a preset only when a raster task exists */
        if (hasRasterTask(extractionPlan.tasks)) {
            var tracingPresets = app.tracingPresetsList;
            if (tracingPresets && tracingPresets.length) {
                tracingPresetName = showPresetDialog(tracingPresets);
                if (tracingPresetName === null) return "done";
            }
        }

        var progress = createProgressWindow(extractionPlan.tasks.length * 2);
        progress.set(0, getLabel("progress.preparing"));

        var workLayer = doc.layers.add();
        workLayer.name = WORK_LAYER_NAME;

        var sessionResult = "done";
        try {
            var workTasks = duplicateTasksToWorkLayer(workLayer, extractionPlan);
            sessionResult = processPaletteTasks(doc, workTasks, progress, tracingPresetName);
            progress.set(extractionPlan.tasks.length * 2, getLabel("progress.done"));
        } finally {
            try {
                workLayer.remove();
            } catch (e) {
                logError(e, "work layer remove");
            }
            progress.close();
        }

        app.redraw();
        return sessionResult;
    }

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var extractionPlan = buildExtractionPlan(rasterizeClippedGroups(doc, doc.selection));

        /* オブジェクト未選択のときはスウォッチパネルの選択を使う / With no objects selected, fall back to the swatch selection */
        if (extractionPlan.tasks.length === 0) {
            var swatchColors = getSelectedSwatchColors(doc);
            if (swatchColors.length === 0) {
                alert(getLabel("alert.noSwatchSelected"));
                return;
            }
            drawPaletteFromSwatches(doc, swatchColors);
            return;
        }

        while (runPaletteSession(doc, extractionPlan) === "retry") {
            /* 「選び直す」が選ばれている間は繰り返す / Repeat while the user keeps choosing Reselect */
        }
    }

    main();

})();
