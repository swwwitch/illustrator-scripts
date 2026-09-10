#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを一時アートボードに収め、背景・マージン・罫線・書き出しサイズ・ファイル名を指定してPNG書き出しします。
設定した内容はアートボード上でそのままプレビューでき、よく使う組み合わせはプリセットとして呼び出せます。

詳細は README を参照してください。

### Overview

Places the selection on a temporary artboard and exports it as PNG with a chosen background, margin, border, size and filename.
Every setting is previewed on the artboard itself, and frequently used combinations can be recalled as presets.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectExporter";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-11";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectExporter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectExporter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/necf308c39f5d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 作業用に一時生成するレイヤー・オブジェクトの名前 / Names of the temporary layer and artwork */
    var PREVIEW_LAYER_NAME      = "__preview";
    var PREVIEW_BACKGROUND_NAME = "preview_background";
    var PREVIEW_BORDER_NAME     = "preview_border";

    /* 単位ごとの罫線幅の初期値 / Initial border width per ruler unit */
    var DEFAULT_BORDER_BY_UNIT = { mm: 0.1, _fallback: 1 };

    /* 透明グリッド1マスの基準サイズ（pt、100%のとき）/ Checker tile size at 100%, in points */
    var CHECKER_TILE_PT = 8;

    /* 透明グリッドの最大マス数（細かすぎる指定で描画が終わらなくなるのを防ぐ）
       / Cap on checker tiles, so a tiny percentage cannot stall the redraw */
    var MAX_CHECKER_TILES = 4000;

    /* PNG書き出しで指定できる最大倍率 / Maximum scale the PNG export accepts */
    var MAX_EXPORT_SCALE = 776.19;

    /* 倍率ラジオボタンに並べる値と初期選択 / Scale choices and the initial selection */
    var SCALE_CHOICES = [100, 200, 300, 400];
    var DEFAULT_SCALE = 400;

    /* プリセット保存ファイルの接頭辞（デスクトップに書き出す）/ Prefix of the preset file saved to the desktop */
    var PRESET_FILE_PREFIX = "export-setting-";

    /* 書き出しファイル名で接尾辞を省いたときの既定語 / Fallback word used when no suffix is given */
    var DEFAULT_SUFFIX_WORD = "selection";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between columns */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */

    var NUMBER_FIELD_CHARS = 4;                /* 数値入力欄の文字数 / width of a numeric field */
    var MARGIN_FIELD_CHARS = 3;                /* マージン入力欄の文字数 / width of a margin field */
    var COLOR_FIELD_CHARS  = 12;               /* カラーコード入力欄の文字数（C0M100Y100K0 が収まる幅）/ width of a color code field */
    var SUFFIX_FIELD_CHARS = 14;               /* 接尾辞入力欄の文字数 / width of the suffix field */
    var SIZE_RADIO_WIDTH   = 60;               /* 倍率・横幅ラジオのラベル幅 / label width of the scale rows */
    var MARGIN_CELL_WIDTH  = { ja: 66, en: 82 }; /* マージン3×3グリッドの1マス幅 / cell width of the 3x3 margin grid */
    var FILENAME_ROW_HEIGHT = 22;              /* ファイル名プレビューの行高（ディセンダー切れ防止）/ row height of the filename preview */

    var DIALOG_OFFSET_X = 300;                 /* 画面中央からの横オフセット / horizontal offset from the screen center */

    /* 倍率ラジオの選択中／非選択の文字色。暗いUIでは黒が沈むので明暗で切り替える
       / Foreground colors of the scale radios; black disappears on a dark UI, so switch by brightness */
    function isLightUserInterface() {
        return app.preferences.getRealPreference("uiBrightness") > 0.5;
    }
    var SCALE_ACTIVE_COLOR   = isLightUserInterface() ? [0, 0, 0] : [0.9, 0.9, 0.9];
    var SCALE_INACTIVE_COLOR = [0.5, 0.5, 0.5];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (($.locale || "") + "").indexOf("ja") === 0 ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /**
     * ラベル定義から現在の言語の文字列を取得する
     * @param {Object} labelSet - { ja: string, en: string } 形式のラベル定義
     * @returns {string} 表示文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        return (labelSet[uiLang] != null) ? labelSet[uiLang] : (labelSet.en || "");
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: string, en: string } 形式のラベル定義
     * @returns {string} コロン付きの表示文字列
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ": ");
    }

    var LABELS = {
        dialog: {
            title: { ja: "選択オブジェクトを書き出し", en: "Export Selected Objects" }
        },
        panel: {
            background: { ja: "背景色", en: "Background" },
            margin: { ja: "マージン", en: "Margin" },
            border: { ja: "罫線", en: "Border" },
            size: { ja: "書き出しサイズ（px）", en: "Export Size (px)" },
            fileName: { ja: "書き出しファイル名", en: "Export Filename" },
            location: { ja: "書き出し先", en: "Export Location" }
        },
        radio: {
            transparent: { ja: "透過", en: "Transparent" },
            black: { ja: "黒", en: "Black" },
            white: { ja: "白", en: "White" },
            checker: { ja: "透明グリッド", en: "Transparency Grid" },
            colorCode: { ja: "カラー指定", en: "Color Code" },
            useDocName: { ja: "参照する", en: "Use" },
            ignoreDocName: { ja: "参照しない", en: "Ignore" },
            none: { ja: "なし", en: "None" },
            desktop: { ja: "デスクトップ", en: "Desktop" },
            documentFolder: { ja: "ファイルと同じ場所", en: "Same as File" }
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            borderWidth: { ja: "線幅", en: "Width" },
            borderColor: { ja: "罫線カラー", en: "Border Color" },
            customScale: { ja: "倍率", en: "Scale" },
            targetWidth: { ja: "横幅", en: "Width" },
            documentName: { ja: "ファイル名", en: "Filename" },
            marginTop: { ja: "上", en: "Top" },
            marginBottom: { ja: "下", en: "Bottom" },
            marginLeft: { ja: "左", en: "Left" },
            marginRight: { ja: "右", en: "Right" },
            roundMode: { ja: "サイズの微調整", en: "Size fine-tuning" },
            delimiter: { ja: "区切り文字", en: "delimiter" },
            suffix: { ja: "接尾辞", en: "Suffix" }
        },
        checkbox: {
            linkMargin: { ja: "連動", en: "Same Value" },
            showFolder: { ja: "書き出し後、フォルダーを表示", en: "Show Folder After Export" }
        },
        roundMode: {
            pixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to pixel grid" },
            currentUnit: { ja: "現在の単位で値を整数値に", en: "Round values in current unit" },
            none: { ja: "何もしない", en: "Do nothing" }
        },
        tooltip: {
            linkMargin: { ja: "上の値を下・左・右にも適用します。", en: "Apply the top value to bottom, left, and right." },
            roundPixel: {
                ja: "書き出し範囲を整数ピクセルまで広げます（倍率100%で1pt＝1px）。",
                en: "Grow the export area to whole pixels (at 100%, one pt equals one px)."
            },
            roundUnit: {
                ja: "書き出し範囲を現在の定規単位で整数になるまで広げます。",
                en: "Grow the export area to whole units in the current ruler unit."
            },
            roundNone: {
                ja: "丸めずに、計測した範囲のまま書き出します。",
                en: "Export the measured area as is, without rounding."
            }
        },
        dropdown: {
            custom: { ja: "カスタム", en: "Custom" }
        },
        button: {
            savePreset: { ja: "プリセットを保存", en: "Save Preset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        preset: {
            transparent200: { ja: "透過・マージンなし・倍率200%", en: "Transparent / No Margin / 200%" },
            whiteHorizontal300: { ja: "白背景・左右3mm・倍率300%", en: "White BG / Horizontal 3mm / 300%" },
            whiteVerticalWidth1000: { ja: "白背景・上下5mm・幅指定1000px", en: "White BG / Vertical 5mm / Width 1000px" },
            blackAll200: { ja: "黒背景・四辺10mm・倍率200%", en: "Black BG / All 10mm / 200%" }
        },
        prompt: {
            presetName: { ja: "プリセット名を入力してください", en: "Enter preset name" },
            defaultPresetName: { ja: "マイプリセット", en: "MyPreset" }
        },
        alert: {
            noSelection: {
                ja: "ドキュメントが開かれていないか、オブジェクトが選択されていません。",
                en: "No document open or no object selected."
            },
            invalidSize: { ja: "選択範囲のサイズが無効です。", en: "Invalid selection size." },
            exportFailed: { ja: "書き出しに失敗しました：", en: "Export failed: " },
            scaleLimited: {
                ja: "書き出し倍率が上限を超えたため、次の倍率で書き出しました：",
                en: "The requested scale exceeds the maximum, so the image was exported at: "
            },
            presetSaved: { ja: "プリセットを保存しました：", en: "Preset saved: " },
            presetSaveFailed: { ja: "プリセットの保存に失敗しました：", en: "Failed to save the preset: " }
        }
    };

    /* 初期プリセット（値の書式は background / margin / round / border / location / delimiter / suffix / size）
       / Built-in presets, encoded the same way as a saved preset file */
    var PRESETS = [
        {
            label: LABELS.preset.transparent200,
            background: "transparent",
            margin: "0,0,0,0",
            round: "pixelGrid",
            border: "none",
            location: "desktop",
            delimiter: "",
            suffix: "200",
            size: "scale:200"
        },
        {
            label: LABELS.preset.whiteHorizontal300,
            background: "white",
            margin: "0,0,3,3",
            round: "pixelGrid",
            border: "none",
            location: "desktop",
            delimiter: "_",
            suffix: "300",
            size: "scale:300"
        },
        {
            label: LABELS.preset.whiteVerticalWidth1000,
            background: "white",
            margin: "5,5,0,0",
            round: "pixelGrid",
            border: "1,black",
            location: "desktop",
            delimiter: "_",
            suffix: "1000",
            size: "width:1000"
        },
        {
            label: LABELS.preset.blackAll200,
            background: "black",
            margin: "10,10,10,10",
            round: "pixelGrid",
            border: "none",
            location: "desktop",
            delimiter: "-",
            suffix: "200",
            size: "scale:200"
        }
    ];

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ダイアログウィンドウに共通レイアウトを適用する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {Object} titleLabelSet - パネル見出しのラベル定義
     * @param {string} [unitLabel] - 見出しに添える単位（省略時は付けない）
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, titleLabelSet, unitLabel) {
        var panelTitle = getLabel(titleLabelSet);
        if (unitLabel) panelTitle += (uiLang === "ja") ? "（" + unitLabel + "）" : " (" + unitLabel + ")";
        var createdPanel = parentContainer.add("panel", undefined, panelTitle);
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 横並びの行グループを生成する
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @param {string} [horizontalAlign] - 横方向の揃え
     * @returns {Group} 生成したグループ
     */
    function addRow(parentContainer, horizontalAlign) {
        var createdGroup = parentContainer.add("group");
        setupRow(createdGroup, horizontalAlign);
        return createdGroup;
    }

    /**
     * 行の項目名（コロン付き）を追加する
     * @param {Group} parentGroup - 追加先の行グループ
     * @param {Object} labelSet - ラベル定義
     * @returns {StaticText} 追加した項目名
     */
    function addRowLabel(parentGroup, labelSet) {
        return parentGroup.add("statictext", undefined, labelText(labelSet));
    }

    /**
     * 数値入力欄を追加する（↑↓キーでの増減付き）
     * @param {Group} parentGroup - 追加先の行グループ
     * @param {string} initialText - 初期値
     * @param {number} charWidth - 入力欄の文字数
     * @param {function} [onValueChanged] - 値が変わったときに呼ぶコールバック
     * @returns {EditText} 追加した入力欄
     */
    function addNumberField(parentGroup, initialText, charWidth, onValueChanged) {
        var inputField = parentGroup.add("edittext", undefined, initialText);
        inputField.characters = charWidth;
        changeValueByArrowKey(inputField, onValueChanged);
        if (typeof onValueChanged === "function") {
            inputField.onChange = function() {
                onValueChanged(inputField.text);
            };
        }
        return inputField;
    }

    /**
     * 入力欄に↑↓キーでの値増減を追加する（Shift併用で10単位スナップ、option併用で0.1単位）
     * @param {EditText} inputField - 対象の入力欄
     * @param {function} [onValueChanged] - 値を更新したあとに呼ぶコールバック
     * @returns {void}
     */
    function changeValueByArrowKey(inputField, onValueChanged) {
        inputField.addEventListener("keydown", function(event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var currentValue = Number(inputField.text);
            if (isNaN(currentValue)) return;

            /* 修飾キーは keyboardState から読む / Read the modifiers from keyboardState */
            var keyboardState = ScriptUI.environment.keyboardState;
            var stepDirection = (event.keyName == "Up") ? 1 : -1;

            if (keyboardState.shiftKey) {
                /* Shift併用は「10の倍数」にスナップ / Snap to multiples of 10 with Shift */
                currentValue = Math.round(currentValue / 10) * 10 + stepDirection * 10;
                if (currentValue < 0) currentValue = 0;
            } else if (keyboardState.altKey) {
                /* option併用は0.1単位（小数第1位に丸め）/ Step by 0.1 with option */
                currentValue = Math.round((currentValue + stepDirection * 0.1) * 10) / 10;
            } else {
                currentValue += stepDirection;
                if (currentValue < 0) currentValue = 0;
            }

            event.preventDefault();
            inputField.text = currentValue;
            if (typeof onValueChanged === "function") {
                onValueChanged(inputField.text);
            }
        });
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * 定規の単位ラベルとpt換算係数を取得する
     * @returns {{label: string, factor: number}} 単位ラベルと1単位あたりのpt数
     */
    function getRulerUnitInfo() {
        var rulerType = app.preferences.getIntegerPreference("rulerType");
        var unitTable = {
            0: { label: "inch", factor: 72.0 },
            1: { label: "mm", factor: 72.0 / 25.4 },
            3: { label: "pica", factor: 12.0 },
            4: { label: "cm", factor: 72.0 / 2.54 },
            5: { label: "Q", factor: 72.0 / 25.4 * 0.25 },
            6: { label: "px", factor: 1.0 }
        };
        /* 未対応の rulerType（2 を含む）は pt 扱い / Unknown ruler types, 2 included, fall back to pt */
        return unitTable[rulerType] || { label: "pt", factor: 1.0 };
    }

    /**
     * 単位ごとの既定値を取り出す
     * @param {Object} defaultsByUnit - 単位ラベルをキーにした既定値の表
     * @param {string} unitLabel - 単位ラベル
     * @returns {number} 既定値
     */
    function getDefaultForUnit(defaultsByUnit, unitLabel) {
        return (defaultsByUnit[unitLabel] != null) ? defaultsByUnit[unitLabel] : defaultsByUnit._fallback;
    }

    /**
     * 入力文字列を数値として読み取る（不正値は0）
     * @param {string} inputText - 入力文字列
     * @returns {number} 読み取った数値
     */
    function toNumber(inputText) {
        var parsedValue = parseFloat(inputText);
        return isNaN(parsedValue) ? 0 : parsedValue;
    }

    /**
     * pt値をピクセル数（整数）に切り上げる
     * @param {number} valuePt - pt値
     * @returns {number} 切り上げたピクセル数
     */
    function ceilToPixel(valuePt) {
        return Math.ceil(valuePt);
    }

    // =========================================
    // カラー / Colors
    // =========================================

    /* 受け付けるカラーコードの書式 / Accepted color code formats */
    var COLOR_CODE_PATTERN = /^(#[0-9A-F]{6}|R\d{1,3}G\d{1,3}B\d{1,3}|C\d{1,3}M\d{1,3}Y\d{1,3}K\d{1,3})$/;

    /* 「カラー指定」を選んだことだけを表す内部キーワード（入力欄の値は書き換えない）
       / Internal keyword meaning "the color code radio is selected", leaving the field untouched */
    var COLOR_CODE_KEYWORD = "colorCode";

    /**
     * ドキュメントのカラースペースに合わせた白を生成する
     * @returns {RGBColor|CMYKColor} 白
     */
    function createWhiteColor() {
        if (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB) {
            var whiteRgb = new RGBColor();
            whiteRgb.red = 255;
            whiteRgb.green = 255;
            whiteRgb.blue = 255;
            return whiteRgb;
        }
        var whiteCmyk = new CMYKColor();
        whiteCmyk.cyan = 0;
        whiteCmyk.magenta = 0;
        whiteCmyk.yellow = 0;
        whiteCmyk.black = 0;
        return whiteCmyk;
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を生成する
     * @returns {RGBColor|CMYKColor} 黒
     */
    function createBlackColor() {
        if (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB) {
            var blackRgb = new RGBColor();
            blackRgb.red = 0;
            blackRgb.green = 0;
            blackRgb.blue = 0;
            return blackRgb;
        }
        var blackCmyk = new CMYKColor();
        blackCmyk.cyan = 0;
        blackCmyk.magenta = 0;
        blackCmyk.yellow = 0;
        blackCmyk.black = 100;
        return blackCmyk;
    }

    /**
     * カラーコード文字列からカラーを生成する（#RRGGBB / R255G255B255 / C0M100Y100K0）
     * @param {string} colorCode - カラーコード文字列
     * @returns {RGBColor|CMYKColor|null} 生成したカラー。解釈できないときは null
     */
    function createColorFromCode(colorCode) {
        var normalizedCode = normalizeColorCode(colorCode);

        if (/^#[0-9A-F]{6}$/.test(normalizedCode)) {
            var hexColor = new RGBColor();
            hexColor.red = parseInt(normalizedCode.substring(1, 3), 16);
            hexColor.green = parseInt(normalizedCode.substring(3, 5), 16);
            hexColor.blue = parseInt(normalizedCode.substring(5, 7), 16);
            return hexColor;
        }

        var rgbMatch = normalizedCode.match(/^R(\d{1,3})G(\d{1,3})B(\d{1,3})$/);
        if (rgbMatch) {
            var rgbColor = new RGBColor();
            rgbColor.red = Math.min(255, parseInt(rgbMatch[1], 10));
            rgbColor.green = Math.min(255, parseInt(rgbMatch[2], 10));
            rgbColor.blue = Math.min(255, parseInt(rgbMatch[3], 10));
            return rgbColor;
        }

        var cmykMatch = normalizedCode.match(/^C(\d{1,3})M(\d{1,3})Y(\d{1,3})K(\d{1,3})$/);
        if (cmykMatch) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = Math.min(100, parseInt(cmykMatch[1], 10));
            cmykColor.magenta = Math.min(100, parseInt(cmykMatch[2], 10));
            cmykColor.yellow = Math.min(100, parseInt(cmykMatch[3], 10));
            cmykColor.black = Math.min(100, parseInt(cmykMatch[4], 10));
            return cmykColor;
        }

        return null;
    }

    /**
     * カラーコードを正規化する（空白を除き、# の無い6桁16進数には # を補う）
     * @param {string} colorCode - 入力されたカラーコード
     * @returns {string} 正規化したカラーコード
     */
    function normalizeColorCode(colorCode) {
        var normalizedCode = String(colorCode).replace(/\s+/g, "").toUpperCase();
        return /^[0-9A-F]{6}$/.test(normalizedCode) ? "#" + normalizedCode : normalizedCode;
    }

    /**
     * カラーコードとして解釈できる文字列かどうかを判定する
     * @param {string} value - 判定する文字列
     * @returns {boolean} 対応書式なら true
     */
    function isColorCode(value) {
        return (typeof value === "string") && COLOR_CODE_PATTERN.test(normalizeColorCode(value));
    }

    // =========================================
    // 表示状態の保存と復元 / Saving and restoring visibility
    // =========================================

    /**
     * 指定レイヤー以外をすべて非表示にする
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} exceptLayer - 表示したままにするレイヤー
     * @returns {Layer[]} 非表示にしたレイヤー
     */
    function hideOtherLayers(doc, exceptLayer) {
        var hiddenLayers = [];
        for (var i = 0; i < doc.layers.length; i++) {
            var targetLayer = doc.layers[i];
            if (targetLayer === exceptLayer || !targetLayer.visible) continue;
            targetLayer.visible = false;
            hiddenLayers.push(targetLayer);
        }
        return hiddenLayers;
    }

    /**
     * 非表示にしたレイヤーを再表示する
     * @param {Layer[]} hiddenLayers - 非表示にしたレイヤー
     * @returns {void}
     */
    function restoreLayerVisibility(hiddenLayers) {
        for (var i = 0; i < hiddenLayers.length; i++) {
            hiddenLayers[i].visible = true;
        }
    }

    /**
     * 選択オブジェクトを指定レイヤーへ複製する（重ね順を維持）
     * @param {PageItem[]} selectedItems - 複製元の選択オブジェクト
     * @param {Layer} targetLayer - 複製先レイヤー
     * @returns {PageItem[]} 複製したオブジェクト
     */
    function duplicateSelectionToLayer(selectedItems, targetLayer) {
        var duplicatedItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            duplicatedItems.push(selectedItems[i].duplicate(targetLayer, ElementPlacement.PLACEATEND));
        }
        return duplicatedItems;
    }

    /**
     * プレビュー用レイヤーを中身ごと削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - 削除するレイヤー名
     * @returns {void}
     */
    function removePreviewLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name !== layerName) continue;
            var targetLayer = doc.layers[i];
            targetLayer.locked = false;
            targetLayer.visible = true;
            targetLayer.remove();
            return;
        }
    }

    // =========================================
    // 背景・罫線の描画 / Drawing the background and border
    // =========================================

    /**
     * 選択オブジェクトの外接範囲を求める
     * テキストは字面の枠ではなく実際の字形で測るため、複製をアウトライン化してから計測する
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getSelectionBounds(selectedItems) {
        var temporaryItems = [];
        try {
            var measureTargets = [];
            for (var i = 0; i < selectedItems.length; i++) {
                measureTargets.push(prepareMeasureTarget(selectedItems[i], temporaryItems));
            }
            return unionVisibleBounds(measureTargets);
        } finally {
            /* 途中で失敗しても計測用の一時オブジェクトは必ず片付ける / The temporary artwork goes away whatever fails */
            for (var j = temporaryItems.length - 1; j >= 0; j--) {
                try {
                    temporaryItems[j].remove();
                } catch (e) { /* アウトライン化で無効になった参照は無視 / A reference invalidated by createOutline is fine to skip */ }
            }
        }
    }

    /**
     * 計測に使うオブジェクトを用意する
     * テキストを含むものだけ複製し、複製側のテキストをアウトライン化する（元のオブジェクトは触らない）
     * @param {PageItem} item - 計測したいオブジェクト
     * @param {PageItem[]} temporaryItems - 計測用に作った一時オブジェクトの記録先
     * @returns {PageItem} 計測に使うオブジェクト
     */
    function prepareMeasureTarget(item, temporaryItems) {
        if (!containsText(item)) return item;

        /* 複製を先に控えてからアウトライン化する（失敗しても複製を消せるように）
           / Track the copy before outlining, so it can be removed even on failure */
        var itemCopy = item.duplicate();
        temporaryItems.push(itemCopy);

        if (itemCopy.typename !== "TextFrame") {
            outlineTextsInPlace(itemCopy);
            return itemCopy;
        }

        /* createOutline() は元のテキストを差し替えるので、戻り値のグループを計測対象にする
           / createOutline() replaces the text frame, so the returned group is what gets measured */
        try {
            var outlinedGroup = itemCopy.createOutline();
            temporaryItems.push(outlinedGroup);
            return outlinedGroup;
        } catch (e) {
            /* 空のテキストなどアウトライン化できないものは複製のまま測る / Text that cannot be outlined is measured as is */
            return itemCopy;
        }
    }

    /**
     * オブジェクトがテキストを含むかを判定する（グループの中も見る）
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} テキストを含むなら true
     */
    function containsText(item) {
        if (item.typename === "TextFrame") return true;
        if (item.typename !== "GroupItem") return false;
        for (var i = 0; i < item.pageItems.length; i++) {
            if (containsText(item.pageItems[i])) return true;
        }
        return false;
    }

    /**
     * グループの中のテキストをその場でアウトライン化する（複製に対してのみ使う）
     * @param {PageItem} item - 走査するオブジェクト
     * @returns {void}
     */
    function outlineTextsInPlace(item) {
        if (item.typename !== "GroupItem") return;
        /* createOutline() が要素を差し替えるので末尾から走査する / createOutline() swaps the item, so walk backwards */
        for (var i = item.pageItems.length - 1; i >= 0; i--) {
            var childItem = item.pageItems[i];
            if (childItem.typename === "TextFrame") {
                try {
                    childItem.createOutline();
                } catch (e) { /* アウトライン化できないテキストはそのまま測る / Leave text that cannot be outlined as is */ }
            } else {
                outlineTextsInPlace(childItem);
            }
        }
    }

    /**
     * 複数オブジェクトの外接範囲を合成する
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function unionVisibleBounds(items) {
        var left = Number.POSITIVE_INFINITY;
        var top = Number.NEGATIVE_INFINITY;
        var right = Number.NEGATIVE_INFINITY;
        var bottom = Number.POSITIVE_INFINITY;

        for (var i = 0; i < items.length; i++) {
            var itemBounds;
            try {
                itemBounds = items[i].visibleBounds;
            } catch (e) {
                /* アウトライン化で中身が無くなったものなどは飛ばす / Skip anything left without geometry */
                continue;
            }
            if (itemBounds[0] < left) left = itemBounds[0];
            if (itemBounds[1] > top) top = itemBounds[1];
            if (itemBounds[2] > right) right = itemBounds[2];
            if (itemBounds[3] < bottom) bottom = itemBounds[3];
        }

        return [left, top, right, bottom];
    }

    /**
     * マージン指定（"上,下,左,右" 形式の "3,3,0,0" など）をpt値に展開する
     * @param {string} marginSpec - マージン指定
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {{top: number, bottom: number, left: number, right: number}} 各辺のマージン（pt）
     */
    function resolveMarginOffsets(marginSpec, unitFactor) {
        var marginValues = String(marginSpec).split(",");
        return {
            top: toNumber(marginValues[0]) * unitFactor,
            bottom: toNumber(marginValues[1]) * unitFactor,
            left: toNumber(marginValues[2]) * unitFactor,
            right: toNumber(marginValues[3]) * unitFactor
        };
    }

    /**
     * 丸めモードから、書き出し範囲を合わせるグリッドの間隔を求める
     * @param {string} roundMode - 丸めモード（pixelGrid / currentUnit / none）
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {number} グリッド間隔（pt）。丸めないときは 0
     */
    function resolveSnapStep(roundMode, unitFactor) {
        if (roundMode === "currentUnit") return unitFactor;
        if (roundMode === "none") return 0;
        /* 既定はピクセルグリッド。倍率100%で 1pt＝1px / Pixel grid by default; at 100% one pt equals one px */
        return 1;
    }

    /**
     * 書き出し範囲を求める（丸めモードに従い、オブジェクトが欠けないよう外側へ広げる）
     * @param {number[]} selectionBounds - 選択オブジェクトの外接範囲 [左, 上, 右, 下]
     * @param {{top: number, bottom: number, left: number, right: number}} marginOffsets - 各辺のマージン（pt）
     * @param {string} roundMode - 丸めモード（pixelGrid / currentUnit / none）
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 書き出し範囲
     */
    function buildExportRect(selectionBounds, marginOffsets, roundMode, unitFactor) {
        var left = selectionBounds[0] - marginOffsets.left;
        var top = selectionBounds[1] + marginOffsets.top;
        var right = selectionBounds[2] + marginOffsets.right;
        var bottom = selectionBounds[3] - marginOffsets.bottom;

        var snapStep = resolveSnapStep(roundMode, unitFactor);
        if (snapStep > 0) {
            left = Math.floor(left / snapStep) * snapStep;
            top = Math.ceil(top / snapStep) * snapStep;
            right = Math.ceil(right / snapStep) * snapStep;
            bottom = Math.floor(bottom / snapStep) * snapStep;
        }
        return {
            left: left,
            top: top,
            right: right,
            bottom: bottom,
            width: right - left,
            height: top - bottom
        };
    }

    /**
     * 罫線指定（"none" / "0.1,black" など）から線幅（pt、最小1px）を求める
     * @param {string} borderSpec - 罫線指定
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {number} 線幅（pt）。罫線なしなら 0
     */
    function resolveBorderWidth(borderSpec, unitFactor) {
        if (!borderSpec || borderSpec === "none") return 0;
        var borderWidth = Math.ceil(toNumber(borderSpec.split(",")[0]) * unitFactor);
        return (borderWidth > 0) ? Math.max(1, borderWidth) : 0;
    }

    /**
     * 罫線指定からカラーを生成する
     * @param {string} borderSpec - 罫線指定
     * @returns {RGBColor|CMYKColor|null} 罫線カラー
     */
    function resolveBorderColor(borderSpec) {
        var colorName = String(borderSpec).split(",")[1];
        if (colorName === "black") return createBlackColor();
        if (colorName === "white") return createWhiteColor();
        return isColorCode(colorName) ? createColorFromCode(colorName) : null;
    }

    /**
     * 背景オブジェクトを生成する（書き出し・プレビュー共用）
     * @param {string} backgroundChoice - 背景指定（transparent / white / black / transparentGrid / #RRGGBB）
     * @param {Object} exportRect - 書き出し範囲
     * @param {number} checkerPercent - 透明グリッドの倍率（%）
     * @returns {PageItem|null} 生成した背景。透過のときは null
     */
    function createExportBackground(backgroundChoice, exportRect, checkerPercent) {
        var doc = app.activeDocument;

        if (backgroundChoice === "transparentGrid") {
            var tileSize = CHECKER_TILE_PT * (checkerPercent / 100);
            var checkerGroup = doc.groupItems.add();
            checkerGroup.name = PREVIEW_BACKGROUND_NAME;
            drawCheckerPattern(checkerGroup, exportRect, tileSize);
            return checkerGroup;
        }

        var backgroundColor = null;
        if (backgroundChoice === "white") backgroundColor = createWhiteColor();
        else if (backgroundChoice === "black") backgroundColor = createBlackColor();
        else if (isColorCode(backgroundChoice)) backgroundColor = createColorFromCode(backgroundChoice);

        /* 解釈できないカラーコードは描かずに透過のまま見せる / An unreadable color code stays transparent */
        if (!backgroundColor) return null;

        var backgroundRect = doc.pathItems.rectangle(exportRect.top, exportRect.left, exportRect.width, exportRect.height);
        backgroundRect.name = PREVIEW_BACKGROUND_NAME;
        backgroundRect.filled = true;
        backgroundRect.stroked = false;
        backgroundRect.fillColor = backgroundColor;
        backgroundRect.zOrder(ZOrderMethod.SENDTOBACK);
        return backgroundRect;
    }

    /**
     * 透明グリッド（市松模様）を描画する
     * @param {GroupItem} parentGroup - 追加先グループ
     * @param {Object} exportRect - 書き出し範囲
     * @param {number} tileSize - 1マスの大きさ（pt）
     * @returns {void}
     */
    function drawCheckerPattern(parentGroup, exportRect, tileSize) {
        var isRgbDocument = (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB);
        var grayColor = isRgbDocument ? new RGBColor() : new CMYKColor();
        var whiteColor = createWhiteColor();

        if (isRgbDocument) {
            grayColor.red = 204;
            grayColor.green = 204;
            grayColor.blue = 204;
        } else {
            grayColor.cyan = 0;
            grayColor.magenta = 0;
            grayColor.yellow = 0;
            grayColor.black = 30;
        }

        var columnCount = Math.ceil(exportRect.width / tileSize);
        var rowCount = Math.ceil(exportRect.height / tileSize);

        /* マスが細かすぎるときはマスを大きくして描画量を抑える / Enlarge the tiles when there would be too many */
        if (columnCount * rowCount > MAX_CHECKER_TILES) {
            tileSize = tileSize * Math.sqrt(columnCount * rowCount / MAX_CHECKER_TILES);
            columnCount = Math.ceil(exportRect.width / tileSize);
            rowCount = Math.ceil(exportRect.height / tileSize);
        }

        for (var i = 0; i < rowCount; i++) {
            for (var j = 0; j < columnCount; j++) {
                var tileRect = parentGroup.pathItems.rectangle(
                    exportRect.top - (i * tileSize),
                    exportRect.left + (j * tileSize),
                    tileSize,
                    tileSize
                );
                tileRect.filled = true;
                tileRect.stroked = false;
                tileRect.fillColor = ((i + j) % 2 === 0) ? grayColor : whiteColor;
            }
        }
        parentGroup.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 書き出し範囲の内側に罫線を描画する
     * @param {Object} exportRect - 書き出し範囲
     * @param {number} borderWidth - 線幅（pt）
     * @param {RGBColor|CMYKColor|null} borderColor - 罫線カラー
     * @returns {PathItem|null} 生成した罫線。描画しないときは null
     */
    function drawBorderRectangle(exportRect, borderWidth, borderColor) {
        if (borderWidth <= 0) return null;
        /* 線幅が書き出し範囲より太いと矩形を作れない / A stroke wider than the area leaves no rectangle to draw */
        if (exportRect.width <= borderWidth || exportRect.height <= borderWidth) return null;

        /* 線の中心が範囲の内側に収まるよう半分だけ内側に寄せる / Inset by half the stroke so it stays inside */
        var halfStroke = borderWidth / 2;
        var borderRect = app.activeDocument.pathItems.rectangle(
            exportRect.top - halfStroke,
            exportRect.left + halfStroke,
            exportRect.width - borderWidth,
            exportRect.height - borderWidth
        );
        borderRect.name = PREVIEW_BORDER_NAME;
        borderRect.filled = false;
        borderRect.stroked = true;
        borderRect.strokeWidth = borderWidth;
        if (borderColor) borderRect.strokeColor = borderColor;
        return borderRect;
    }

    /**
     * プレビュー・書き出し用に生成した背景と罫線を削除する
     * @returns {void}
     */
    function removePreviewArtwork() {
        var doc = app.activeDocument;
        var itemsToRemove = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var targetItem = doc.pageItems[i];
            if (targetItem.name === PREVIEW_BACKGROUND_NAME || targetItem.name === PREVIEW_BORDER_NAME) {
                itemsToRemove.push(targetItem);
            }
        }
        for (var j = 0; j < itemsToRemove.length; j++) {
            /* 親ごと削除済みのことがあるため、失敗しても続行 / A parent may already be gone, so keep going */
            try {
                itemsToRemove[j].remove();
            } catch (e) {}
        }
    }

    // =========================================
    // ファイル名 / Filename
    // =========================================

    /**
     * ファイル名の禁則文字・空白類を区切り文字に置き換える
     * @param {string} fileName - 置換前のファイル名
     * @param {string} delimiter - 区切り文字（"-" または "_"）
     * @returns {string} 置換後のファイル名
     */
    function sanitizeFileName(fileName, delimiter) {
        var replacement = (delimiter === "-") ? "-" : "_";
        return fileName.replace(/[¥\/:*?"<>|\r\n\t　 ]/g, replacement);
    }

    /**
     * 書き出しファイル名を組み立てる（プレビューと書き出しで共用）
     * @param {string} documentBaseName - 拡張子を除いたドキュメント名
     * @param {string} delimiter - 区切り文字
     * @param {string} suffix - 接尾辞
     * @param {boolean} useDocumentName - ドキュメント名を使うかどうか
     * @returns {string} 書き出しファイル名
     */
    function buildExportFileName(documentBaseName, delimiter, suffix, useDocumentName) {
        var baseName = useDocumentName ? documentBaseName : "";
        var fileName = suffix ?
            baseName + delimiter + suffix + ".png" :
            baseName + "_" + DEFAULT_SUFFIX_WORD + ".png";
        return sanitizeFileName(fileName, delimiter);
    }

    // =========================================
    // ダイアログの設定値 / Reading the dialog settings
    // =========================================

    /**
     * 背景の選択状態を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {string} 背景指定（transparent / white / black / transparentGrid / #RRGGBB）
     */
    function getBackgroundChoice(controls) {
        var background = controls.background;
        if (background.checker.value) return "transparentGrid";
        if (background.white.value) return "white";
        if (background.black.value) return "black";
        if (background.colorCode.value) return normalizeColorCode(background.colorCodeInput.text);
        return "transparent";
    }

    /**
     * 透明グリッドの倍率を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {number} 倍率（%）
     */
    function getCheckerPercent(controls) {
        var checkerPercent = toNumber(controls.background.checkerScaleInput.text);
        return (checkerPercent > 0) ? checkerPercent : 100;
    }

    /**
     * マージンの入力値を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {string} マージン指定（上,下,左,右 / 例 "3,3,0,0"）
     */
    function getMarginSpec(controls) {
        var margin = controls.margin;
        return [
            toNumber(margin.topInput.text),
            toNumber(margin.bottomInput.text),
            toNumber(margin.leftInput.text),
            toNumber(margin.rightInput.text)
        ].join(",");
    }

    /**
     * 丸めモードの選択状態を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {string} 丸めモード（pixelGrid / currentUnit / none）
     */
    function getRoundMode(controls) {
        var margin = controls.margin;
        if (margin.roundUnit.value) return "currentUnit";
        if (margin.roundNone.value) return "none";
        return "pixelGrid";
    }

    /**
     * 罫線の選択状態を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {string} 罫線指定（none / 0.1,black など）
     */
    function getBorderSpec(controls) {
        return controls.border.enabled.value ? readBorderSpec(controls.border) : "none";
    }

    /**
     * 罫線パネルの入力値から罫線指定を組み立てる
     * @param {Object} border - 罫線のコントロール一式
     * @returns {string} 罫線指定（線幅,カラー）
     */
    function readBorderSpec(border) {
        var borderColorName = "black";
        if (border.white.value) borderColorName = "white";
        else if (border.colorCode.value) borderColorName = normalizeColorCode(border.colorCodeInput.text);
        return border.widthInput.text + "," + borderColorName;
    }

    /**
     * 書き出しサイズの選択状態を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {string} サイズ指定（scale:200 / width:1000 など）
     */
    function getSizeSpec(controls) {
        var size = controls.size;
        for (var i = 0; i < size.scaleRadios.length; i++) {
            if (size.scaleRadios[i].value) return "scale:" + SCALE_CHOICES[i];
        }
        if (size.customScale.value) return "scale:" + size.customScaleInput.text;
        if (size.targetWidth.value) return "width:" + size.targetWidthInput.text;
        return "scale:100";
    }

    /**
     * 区切り文字の選択状態を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {string} 区切り文字（"-" / "_" / ""）
     */
    function getDelimiter(controls) {
        if (controls.fileName.delimiterDash.value) return "-";
        if (controls.fileName.delimiterUnderscore.value) return "_";
        return "";
    }

    /**
     * 接尾辞の入力値を取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {string} 接尾辞
     */
    function getSuffix(controls) {
        return controls.fileName.suffixEnabled.value ? controls.fileName.suffixInput.text : "";
    }

    /**
     * 書き出し先フォルダーを取得する
     * @param {Object} controls - ダイアログのコントロール一式
     * @param {Document} doc - 対象ドキュメント
     * @returns {Folder} 書き出し先フォルダー
     */
    function getDestinationFolder(controls, doc) {
        if (controls.location.desktop.value) return Folder.desktop;
        /* 未保存の書類には保存先が無いのでデスクトップへ逃がす / An unsaved document has no folder, so fall back to the desktop */
        try {
            return doc.fullName.parent;
        } catch (e) {
            return Folder.desktop;
        }
    }

    /**
     * ダイアログの入力内容を書き出し設定にまとめる
     * @param {Object} controls - ダイアログのコントロール一式
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} 書き出し設定
     */
    function collectExportSettings(controls, doc) {
        return {
            backgroundChoice: getBackgroundChoice(controls),
            checkerPercent: getCheckerPercent(controls),
            marginSpec: getMarginSpec(controls),
            roundMode: getRoundMode(controls),
            borderSpec: getBorderSpec(controls),
            sizeSpec: getSizeSpec(controls),
            fileName: controls.fileName.previewText.text,
            destinationFolder: getDestinationFolder(controls, doc),
            showFolder: controls.location.showFolder.value
        };
    }

    /**
     * 現在の設定をプリセット定義としてデスクトップに書き出す
     * @param {Object} controls - ダイアログのコントロール一式
     * @returns {void}
     */
    function savePresetToFile(controls) {
        var presetName = prompt(getLabel(LABELS.prompt.presetName), getLabel(LABELS.prompt.defaultPresetName));
        if (!presetName) return;

        var today = new Date();
        var dateStamp = today.getFullYear() +
            ("0" + (today.getMonth() + 1)).slice(-2) +
            ("0" + today.getDate()).slice(-2);
        var presetFile = new File(Folder.desktop + "/" + PRESET_FILE_PREFIX + dateStamp + ".txt");
        for (var serialNumber = 2; presetFile.exists; serialNumber++) {
            presetFile = new File(Folder.desktop + "/" + PRESET_FILE_PREFIX + dateStamp + "_" + serialNumber + ".txt");
        }

        /* PRESETS にそのまま貼り込める形で書き出す / Written so it can be pasted straight into PRESETS */
        var presetLines = [
            "{",
            '    label: { ja: "' + presetName + '", en: "' + presetName + '" },',
            '    background: "' + getBackgroundChoice(controls) + '",',
            '    margin: "' + getMarginSpec(controls) + '",',
            '    round: "' + getRoundMode(controls) + '",',
            '    border: "' + getBorderSpec(controls) + '",',
            '    location: "' + (controls.location.desktop.value ? "desktop" : "documentFolder") + '",',
            '    delimiter: "' + getDelimiter(controls) + '",',
            '    suffix: "' + getSuffix(controls) + '",',
            '    size: "' + getSizeSpec(controls) + '"',
            "}"
        ];

        try {
            presetFile.encoding = "UTF-8";
            presetFile.open("w");
            presetFile.write(presetLines.join("\n"));
            presetFile.close();
        } catch (e) {
            alert(getLabel(LABELS.alert.presetSaveFailed) + e.message);
            return;
        }
        alert(getLabel(LABELS.alert.presetSaved) + presetFile.name);
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * 現在の設定でプレビュー用の背景・罫線を描き直す
     * @param {Object} controls - ダイアログのコントロール一式
     * @param {number[]} selectionBounds - 選択オブジェクトの外接範囲
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {void}
     */
    function renderPreview(controls, selectionBounds, unitFactor) {
        removePreviewArtwork();

        var exportRect = buildExportRect(selectionBounds, resolveMarginOffsets(getMarginSpec(controls), unitFactor),
            getRoundMode(controls), unitFactor);
        createExportBackground(getBackgroundChoice(controls), exportRect, getCheckerPercent(controls));

        var borderSpec = getBorderSpec(controls);
        drawBorderRectangle(exportRect, resolveBorderWidth(borderSpec, unitFactor), resolveBorderColor(borderSpec));

        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 書き出しオプションのダイアログを表示する
     * @param {number[]} selectionBounds - 選択オブジェクトの外接範囲 [左, 上, 右, 下]
     * @param {{label: string, factor: number}} rulerUnit - 定規の単位情報
     * @param {string} documentBaseName - 拡張子を除いたドキュメント名
     * @returns {Object|null} 書き出し設定。キャンセル時は null
     */
    function showExportOptionsDialog(selectionBounds, rulerUnit, documentBaseName) {
        var doc = app.activeDocument;
        var selectionWidth = selectionBounds[2] - selectionBounds[0];
        var selectionHeight = selectionBounds[1] - selectionBounds[3];
        var controls = {};

        /* 現在のマージン設定を含めた書き出し範囲 / Export rect including the current margin */
        function getExportRectFromUI() {
            return buildExportRect(selectionBounds, resolveMarginOffsets(getMarginSpec(controls), rulerUnit.factor),
                getRoundMode(controls), rulerUnit.factor);
        }

        /* 設定が変わるたびにプレビューと倍率ラベルを描き直す / Redraw the preview and the scale labels on every change */
        function refreshPreview() {
            if (controls.size) updateScaleLabels();
            renderPreview(controls, selectionBounds, rulerUnit.factor);
        }

        /* ファイル名プレビューを更新する / Refresh the filename preview */
        function updateFileNamePreview() {
            controls.fileName.previewText.text = buildExportFileName(
                documentBaseName,
                getDelimiter(controls),
                getSuffix(controls),
                !controls.fileName.ignoreDocName.value
            );
        }

        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        // -----------------------------------------
        // プリセット行 / Preset row
        // -----------------------------------------
        var presetRow = addRow(dialog);
        addRowLabel(presetRow, LABELS.fieldLabel.preset);
        var presetNames = [getLabel(LABELS.dropdown.custom)];
        for (var i = 0; i < PRESETS.length; i++) {
            presetNames.push(getLabel(PRESETS[i].label));
        }
        var presetDropdown = presetRow.add("dropdownlist", undefined, presetNames);
        presetDropdown.selection = 0;
        var btnSavePreset = presetRow.add("button", undefined, getLabel(LABELS.button.savePreset));

        // -----------------------------------------
        // 2カラム / Two columns
        // -----------------------------------------
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = COLUMN_SPACING;

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.spacing = COLUMN_SPACING;

        controls.background = buildBackgroundPanel(leftColumn);
        controls.margin = buildMarginPanel(leftColumn);
        controls.border = buildBorderPanel(leftColumn);
        controls.size = buildSizePanel(rightColumn);
        controls.fileName = buildFileNamePanel(rightColumn);
        controls.location = buildLocationPanel(rightColumn);

        /**
         * 背景色パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 背景色のコントロール一式
         */
        function buildBackgroundPanel(parentColumn) {
            var backgroundPanel = addPanel(parentColumn, LABELS.panel.background);

            var basicRow = addRow(backgroundPanel);
            var background = {
                transparent: basicRow.add("radiobutton", undefined, getLabel(LABELS.radio.transparent)),
                black: basicRow.add("radiobutton", undefined, getLabel(LABELS.radio.black)),
                white: basicRow.add("radiobutton", undefined, getLabel(LABELS.radio.white))
            };

            var checkerRow = addRow(backgroundPanel);
            background.checker = checkerRow.add("radiobutton", undefined, getLabel(LABELS.radio.checker));
            background.checkerScaleInput = addNumberField(checkerRow, "100", NUMBER_FIELD_CHARS, refreshPreview);
            checkerRow.add("statictext", undefined, "%");

            var colorCodeRow = addRow(backgroundPanel);
            background.colorCode = colorCodeRow.add("radiobutton", undefined, getLabel(LABELS.radio.colorCode));
            background.colorCodeInput = colorCodeRow.add("edittext", undefined, "#ffcc00");
            background.colorCodeInput.characters = COLOR_FIELD_CHARS;
            background.colorCodeInput.onChange = refreshPreview;

            /* 背景の排他選択と入力欄の有効・無効をまとめて切り替える / Switch the background choice and its fields together */
            background.select = function(backgroundChoice) {
                background.transparent.value = (backgroundChoice === "transparent");
                background.black.value = (backgroundChoice === "black");
                background.white.value = (backgroundChoice === "white");
                background.checker.value = (backgroundChoice === "transparentGrid");
                /* 決め打ちの4種以外はカラー指定として扱う（読めない値でも選択は外さない）
                   / Anything but the four keywords means the color code, even when it cannot be parsed */
                background.colorCode.value = !(background.transparent.value || background.black.value ||
                    background.white.value || background.checker.value);
                if (background.colorCode.value && backgroundChoice !== COLOR_CODE_KEYWORD) {
                    background.colorCodeInput.text = backgroundChoice;
                }
                background.colorCodeInput.enabled = background.colorCode.value;
                background.checkerScaleInput.enabled = background.checker.value;
            };

            var backgroundChoices = ["transparent", "black", "white", "transparentGrid", COLOR_CODE_KEYWORD];
            var backgroundRadios = [background.transparent, background.black, background.white, background.checker, background.colorCode];
            for (var i = 0; i < backgroundRadios.length; i++) {
                backgroundRadios[i].onClick = createBackgroundClickHandler(background, backgroundChoices[i]);
            }

            background.select("transparent");
            return background;
        }

        /**
         * 背景ラジオボタンのクリックハンドラーを作る
         * @param {Object} background - 背景色のコントロール一式
         * @param {string} backgroundChoice - 選択される背景指定
         * @returns {function} クリックハンドラー
         */
        function createBackgroundClickHandler(background, backgroundChoice) {
            return function() {
                background.select(backgroundChoice);
                refreshPreview();
            };
        }

        /**
         * マージンパネルを作る（上下左右を3×3に配置し、中央を連動チェックボックスにする）
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} マージンのコントロール一式
         */
        function buildMarginPanel(parentColumn) {
            var marginPanel = addPanel(parentColumn, LABELS.panel.margin, rulerUnit.label);
            var margin = {};

            /* 1行目：［空］［上］［空］/ Row 1: [empty][top][empty] */
            var topRow = addMarginGridRow(marginPanel);
            addMarginGridCell(topRow);
            var topCell = addMarginField(topRow, LABELS.fieldLabel.marginTop, onTopMarginChanged);
            addMarginGridCell(topRow);

            /* 2行目：［左］［連動］［右］/ Row 2: [left][link][right] */
            var middleRow = addMarginGridRow(marginPanel);
            var leftCell = addMarginField(middleRow, LABELS.fieldLabel.marginLeft, refreshPreview);
            var linkCell = addMarginGridCell(middleRow);
            margin.link = linkCell.add("checkbox", undefined, getLabel(LABELS.checkbox.linkMargin));
            margin.link.helpTip = getLabel(LABELS.tooltip.linkMargin);
            var rightCell = addMarginField(middleRow, LABELS.fieldLabel.marginRight, refreshPreview);

            /* 3行目：［空］［下］［空］/ Row 3: [empty][bottom][empty] */
            var bottomRow = addMarginGridRow(marginPanel);
            addMarginGridCell(bottomRow);
            var bottomCell = addMarginField(bottomRow, LABELS.fieldLabel.marginBottom, refreshPreview);
            addMarginGridCell(bottomRow);

            /* サイズの微調整（書き出し範囲の丸め方）/ Size fine-tuning: how the export area is rounded */
            addRowLabel(addRow(marginPanel), LABELS.fieldLabel.roundMode);

            var roundColumn = marginPanel.add("group");
            roundColumn.orientation = "column";
            roundColumn.alignment = ["fill", "top"];
            roundColumn.alignChildren = ["left", "center"];
            roundColumn.spacing = PANEL_SPACING;
            margin.roundPixel = roundColumn.add("radiobutton", undefined, getLabel(LABELS.roundMode.pixelGrid));
            margin.roundUnit = roundColumn.add("radiobutton", undefined, getLabel(LABELS.roundMode.currentUnit));
            margin.roundNone = roundColumn.add("radiobutton", undefined, getLabel(LABELS.roundMode.none));
            margin.roundPixel.helpTip = getLabel(LABELS.tooltip.roundPixel);
            margin.roundUnit.helpTip = getLabel(LABELS.tooltip.roundUnit);
            margin.roundNone.helpTip = getLabel(LABELS.tooltip.roundNone);

            /* 丸めモードをUIへ反映する / Apply a round mode to the UI */
            margin.selectRoundMode = function(roundMode) {
                margin.roundUnit.value = (roundMode === "currentUnit");
                margin.roundNone.value = (roundMode === "none");
                /* 未指定・不明な値はピクセルグリッドに寄せる / An unknown or missing value falls back to the pixel grid */
                margin.roundPixel.value = !(margin.roundUnit.value || margin.roundNone.value);
            };

            var roundRadios = [margin.roundPixel, margin.roundUnit, margin.roundNone];
            for (var i = 0; i < roundRadios.length; i++) {
                roundRadios[i].onClick = refreshPreview;
            }
            margin.selectRoundMode("pixelGrid");

            margin.topInput = topCell.input;
            margin.bottomInput = bottomCell.input;
            margin.leftInput = leftCell.input;
            margin.rightInput = rightCell.input;

            /* 連動中は上の値を残る3辺へ写し、入力できないようにする / While linked, copy the top value and lock the other three */
            function syncLinkedMargins() {
                var isLinked = margin.link.value;
                if (isLinked) {
                    margin.bottomInput.text = margin.topInput.text;
                    margin.leftInput.text = margin.topInput.text;
                    margin.rightInput.text = margin.topInput.text;
                }
                bottomCell.group.enabled = !isLinked;
                leftCell.group.enabled = !isLinked;
                rightCell.group.enabled = !isLinked;
            }

            function onTopMarginChanged() {
                syncLinkedMargins();
                refreshPreview();
            }

            margin.link.onClick = function() {
                syncLinkedMargins();
                refreshPreview();
            };

            /* マージン指定（"3,3,0,0"）をUIへ反映する / Apply a margin spec to the UI */
            margin.select = function(marginSpec) {
                var marginValues = String(marginSpec).split(",");
                margin.topInput.text = String(toNumber(marginValues[0]));
                margin.bottomInput.text = String(toNumber(marginValues[1]));
                margin.leftInput.text = String(toNumber(marginValues[2]));
                margin.rightInput.text = String(toNumber(marginValues[3]));
                /* 四辺が同じ値のときだけ連動状態に戻す / Re-link only when all four sides match */
                margin.link.value = (margin.topInput.text === margin.bottomInput.text &&
                    margin.topInput.text === margin.leftInput.text &&
                    margin.topInput.text === margin.rightInput.text);
                syncLinkedMargins();
            };

            margin.select("0,0,0,0");
            return margin;
        }

        /**
         * マージングリッドの1行を作る
         * @param {Panel} parentPanel - 追加先のパネル
         * @returns {Group} 生成した行グループ
         */
        function addMarginGridRow(parentPanel) {
            var rowGroup = parentPanel.add("group");
            rowGroup.orientation = "row";
            /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
            rowGroup.alignment = ["center", "top"];
            rowGroup.alignChildren = ["center", "center"];
            rowGroup.spacing = 0;
            return rowGroup;
        }

        /**
         * マージングリッドの1マス（固定幅の空グループ）を作る
         * @param {Group} parentRow - 追加先の行グループ
         * @returns {Group} 生成したグループ
         */
        function addMarginGridCell(parentRow) {
            var cellGroup = parentRow.add("group");
            cellGroup.orientation = "row";
            cellGroup.alignment = ["center", "center"];
            cellGroup.alignChildren = ["center", "center"];
            cellGroup.spacing = PANEL_SPACING;
            cellGroup.minimumSize.width = MARGIN_CELL_WIDTH[uiLang];
            return cellGroup;
        }

        /**
         * マージングリッドに項目名＋数値欄のマスを作る
         * @param {Group} parentRow - 追加先の行グループ
         * @param {Object} labelSet - 項目名のラベル定義
         * @param {function} onValueChanged - 値が変わったときに呼ぶコールバック
         * @returns {{group: Group, input: EditText}} マスのグループと入力欄
         */
        function addMarginField(parentRow, labelSet, onValueChanged) {
            var cellGroup = addMarginGridCell(parentRow);
            cellGroup.add("statictext", undefined, labelText(labelSet));
            return { group: cellGroup, input: addNumberField(cellGroup, "0", MARGIN_FIELD_CHARS, onValueChanged) };
        }

        /**
         * 罫線パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 罫線のコントロール一式
         */
        function buildBorderPanel(parentColumn) {
            var borderPanel = addPanel(parentColumn, LABELS.panel.border);

            var widthRow = addRow(borderPanel);
            var border = { enabled: widthRow.add("checkbox", undefined, labelText(LABELS.fieldLabel.borderWidth)) };
            var defaultBorderWidth = getDefaultForUnit(DEFAULT_BORDER_BY_UNIT, rulerUnit.label);
            border.widthInput = addNumberField(widthRow, String(defaultBorderWidth), NUMBER_FIELD_CHARS, refreshPreview);
            widthRow.add("statictext", undefined, rulerUnit.label);

            border.colorLabelRow = addRow(borderPanel);
            addRowLabel(border.colorLabelRow, LABELS.fieldLabel.borderColor);

            border.colorRadioRow = addRow(borderPanel);
            border.black = border.colorRadioRow.add("radiobutton", undefined, getLabel(LABELS.radio.black));
            border.white = border.colorRadioRow.add("radiobutton", undefined, getLabel(LABELS.radio.white));

            border.colorCodeRow = addRow(borderPanel);
            border.colorCode = border.colorCodeRow.add("radiobutton", undefined, getLabel(LABELS.radio.colorCode));
            border.colorCodeInput = border.colorCodeRow.add("edittext", undefined, "#333333");
            border.colorCodeInput.characters = COLOR_FIELD_CHARS;
            border.colorCodeInput.onChange = refreshPreview;

            /* 罫線指定（none / 0.1,black など）をUIへ反映する / Apply a border spec to the UI */
            border.select = function(borderSpec) {
                var hasBorder = (borderSpec !== "none");
                border.enabled.value = hasBorder;

                if (hasBorder) {
                    var specParts = String(borderSpec).split(",");
                    var borderColorName = specParts[1];
                    border.widthInput.text = specParts[0];
                    border.black.value = (borderColorName === "black");
                    border.white.value = (borderColorName === "white");
                    /* 黒・白以外はカラー指定として扱う / Anything but black or white means the color code */
                    border.colorCode.value = !(border.black.value || border.white.value);
                    if (border.colorCode.value && borderColorName !== COLOR_CODE_KEYWORD) {
                        border.colorCodeInput.text = borderColorName;
                    }
                }

                border.widthInput.enabled = hasBorder;
                border.colorLabelRow.enabled = hasBorder;
                border.colorRadioRow.enabled = hasBorder;
                border.colorCodeRow.enabled = hasBorder;
                border.colorCodeInput.enabled = hasBorder && border.colorCode.value;
            };

            border.enabled.onClick = function() {
                border.select(border.enabled.value ? readBorderSpec(border) : "none");
                refreshPreview();
            };
            border.black.onClick = createBorderColorHandler(border, "black");
            border.white.onClick = createBorderColorHandler(border, "white");
            border.colorCode.onClick = createBorderColorHandler(border, COLOR_CODE_KEYWORD);

            border.select("none");
            return border;
        }

        /**
         * 罫線カラーのラジオボタンのクリックハンドラーを作る
         * @param {Object} border - 罫線のコントロール一式
         * @param {string} borderColorName - 選択されるカラー名（black / white / COLOR_CODE_KEYWORD）
         * @returns {function} クリックハンドラー
         */
        function createBorderColorHandler(border, borderColorName) {
            return function() {
                border.select(border.widthInput.text + "," + borderColorName);
                refreshPreview();
            };
        }

        /**
         * 書き出しサイズパネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 書き出しサイズのコントロール一式
         */
        function buildSizePanel(parentColumn) {
            var sizePanel = addPanel(parentColumn, LABELS.panel.size);
            var size = { scaleRadios: [] };

            var scaleColumn = sizePanel.add("group");
            scaleColumn.orientation = "column";
            /* ラベルが伸びても切れないよう、ラジオはパネル幅いっぱいに広げる
               / Stretch the radios to the panel width so a longer label is never clipped */
            scaleColumn.alignment = ["fill", "top"];
            scaleColumn.alignChildren = ["fill", "center"];
            scaleColumn.spacing = PANEL_SPACING;

            var exportRect = getExportRectFromUI();
            for (var i = 0; i < SCALE_CHOICES.length; i++) {
                size.scaleRadios.push(scaleColumn.add("radiobutton", undefined, buildScaleLabel(SCALE_CHOICES[i], exportRect)));
            }

            var customScaleRow = addRow(sizePanel);
            size.customScale = customScaleRow.add("radiobutton", undefined, labelText(LABELS.fieldLabel.customScale));
            size.customScale.preferredSize.width = SIZE_RADIO_WIDTH;
            size.customScaleInput = addNumberField(customScaleRow, String(DEFAULT_SCALE), NUMBER_FIELD_CHARS + 1, function(value) {
                size.select("scale:" + value, false, true);
            });
            customScaleRow.add("statictext", undefined, "%");

            var targetWidthRow = addRow(sizePanel);
            size.targetWidth = targetWidthRow.add("radiobutton", undefined, labelText(LABELS.fieldLabel.targetWidth));
            size.targetWidth.preferredSize.width = SIZE_RADIO_WIDTH;
            size.targetWidthInput = addNumberField(targetWidthRow, "", NUMBER_FIELD_CHARS + 3, function(value) {
                size.select("width:" + value);
            });
            targetWidthRow.add("statictext", undefined, "px");

            /* サイズ指定（scale:200 / width:1000）をUIへ反映する / Apply a size spec to the UI */
            size.select = function(sizeSpec, keepSuffix, forceCustomScale) {
                var isWidthMode = (String(sizeSpec).indexOf("width:") === 0);
                var specValue = String(sizeSpec).split(":")[1];
                var matchedScale = false;

                for (var i = 0; i < size.scaleRadios.length; i++) {
                    /* 倍率指定を選んだときは、値が 1x〜4x と同じでも固定倍率に吸われないようにする
                       / When the custom scale is picked, a matching value must not jump to a fixed radio */
                    var isSelected = !isWidthMode && !forceCustomScale && (specValue === String(SCALE_CHOICES[i]));
                    size.scaleRadios[i].value = isSelected;
                    setScaleRadioColor(size.scaleRadios[i], isSelected);
                    if (isSelected) matchedScale = true;
                }

                size.customScale.value = (!isWidthMode && !matchedScale);
                size.targetWidth.value = isWidthMode;
                size.customScaleInput.enabled = size.customScale.value;
                size.targetWidthInput.enabled = isWidthMode;

                if (isWidthMode) {
                    size.targetWidthInput.text = specValue;
                } else {
                    size.customScaleInput.text = specValue;
                    size.targetWidthInput.text = String(ceilToPixel(getExportRectFromUI().width * toNumber(specValue) / 100));
                }

                /* 倍率・横幅の値をそのまま接尾辞に流用する / Reuse the scale or width value as the suffix */
                if (!keepSuffix) controls.fileName.setSuffix(specValue);
            };

            for (var j = 0; j < size.scaleRadios.length; j++) {
                size.scaleRadios[j].onClick = createScaleClickHandler(size, SCALE_CHOICES[j]);
            }
            size.customScale.onClick = function() {
                size.select("scale:" + size.customScaleInput.text, false, true);
            };
            size.targetWidth.onClick = function() {
                size.select("width:" + size.targetWidthInput.text);
            };

            return size;
        }

        /**
         * 倍率ラジオボタンのラベルを組み立てる（書き出しピクセル数を併記）
         * @param {number} scalePercent - 倍率（%）
         * @param {Object} exportRect - 現在のマージンを含めた書き出し範囲
         * @returns {string} ラベル文字列
         */
        function buildScaleLabel(scalePercent, exportRect) {
            var scaleRatio = scalePercent / 100;
            return (scaleRatio + "x") + (uiLang === "ja" ? "：" : ": ") +
                ceilToPixel(exportRect.width * scaleRatio) + " × " + ceilToPixel(exportRect.height * scaleRatio);
        }

        /**
         * 倍率ラジオボタンのラベルを現在のマージンに合わせて描き直す
         * @returns {void}
         */
        function updateScaleLabels() {
            var exportRect = getExportRectFromUI();
            var labelChanged = false;

            for (var i = 0; i < controls.size.scaleRadios.length; i++) {
                var scaleRadio = controls.size.scaleRadios[i];
                var scaleLabel = buildScaleLabel(SCALE_CHOICES[i], exportRect);
                if (scaleRadio.text !== scaleLabel) {
                    scaleRadio.text = scaleLabel;
                    /* 幅は自動計算に戻す（前のラベル幅のままだと桁が増えたときに切れる）
                       / Reset to auto width, or the control keeps the previous label's width */
                    scaleRadio.preferredSize.width = -1;
                    labelChanged = true;
                }
                setScaleRadioColor(scaleRadio, scaleRadio.value);
            }

            /* 桁が増えてもラベルが切れないよう、変わったときだけ組み直す / Re-layout only when a label changed, so longer text is not clipped */
            if (labelChanged) dialog.layout.layout(true);
        }

        /**
         * 倍率ラジオボタンのクリックハンドラーを作る
         * @param {Object} size - 書き出しサイズのコントロール一式
         * @param {number} scalePercent - 選択される倍率（%）
         * @returns {function} クリックハンドラー
         */
        function createScaleClickHandler(size, scalePercent) {
            return function() {
                size.select("scale:" + scalePercent);
            };
        }

        /**
         * 倍率ラジオボタンの文字色を選択状態に合わせる
         * @param {RadioButton} scaleRadio - 対象のラジオボタン
         * @param {boolean} isSelected - 選択中かどうか
         * @returns {void}
         */
        function setScaleRadioColor(scaleRadio, isSelected) {
            var radioGraphics = scaleRadio.graphics;
            radioGraphics.foregroundColor = radioGraphics.newPen(
                radioGraphics.PenType.SOLID_COLOR,
                isSelected ? SCALE_ACTIVE_COLOR : SCALE_INACTIVE_COLOR,
                1
            );
        }

        /**
         * 書き出しファイル名パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} ファイル名のコントロール一式
         */
        function buildFileNamePanel(parentColumn) {
            var fileNamePanel = addPanel(parentColumn, LABELS.panel.fileName);
            var fileName = {};

            var documentNameRow = addRow(fileNamePanel);
            addRowLabel(documentNameRow, LABELS.fieldLabel.documentName);
            fileName.useDocName = documentNameRow.add("radiobutton", undefined, getLabel(LABELS.radio.useDocName));
            fileName.ignoreDocName = documentNameRow.add("radiobutton", undefined, getLabel(LABELS.radio.ignoreDocName));
            fileName.useDocName.value = true;

            var delimiterRow = addRow(fileNamePanel);
            addRowLabel(delimiterRow, LABELS.fieldLabel.delimiter);
            fileName.delimiterNone = delimiterRow.add("radiobutton", undefined, getLabel(LABELS.radio.none));
            fileName.delimiterDash = delimiterRow.add("radiobutton", undefined, "-");
            fileName.delimiterUnderscore = delimiterRow.add("radiobutton", undefined, "_");

            var suffixRow = addRow(fileNamePanel);
            fileName.suffixEnabled = suffixRow.add("checkbox", undefined, labelText(LABELS.fieldLabel.suffix));
            fileName.suffixInput = suffixRow.add("edittext", undefined, "");
            fileName.suffixInput.characters = SUFFIX_FIELD_CHARS;
            fileName.suffixInput.enabled = false;

            var previewPanel = fileNamePanel.add("panel");
            previewPanel.alignment = ["fill", "top"];
            previewPanel.margins = [10, 10, 10, 10];
            fileName.previewText = previewPanel.add("statictext", undefined, "");
            fileName.previewText.alignment = ["fill", "center"];
            fileName.previewText.preferredSize.height = FILENAME_ROW_HEIGHT;

            /* 接尾辞を指定してファイル名プレビューを更新する / Set the suffix and refresh the preview */
            fileName.setSuffix = function(suffixValue) {
                fileName.suffixEnabled.value = true;
                fileName.suffixInput.enabled = true;
                fileName.suffixInput.text = suffixValue;
                updateFileNamePreview();
            };

            /* 接尾辞をなしに戻す / Clear the suffix */
            fileName.clearSuffix = function() {
                fileName.suffixEnabled.value = false;
                fileName.suffixInput.enabled = false;
                updateFileNamePreview();
            };

            /* 区切り文字を選択する / Select the delimiter */
            fileName.setDelimiter = function(delimiter) {
                fileName.delimiterNone.value = (delimiter === "");
                fileName.delimiterDash.value = (delimiter === "-");
                fileName.delimiterUnderscore.value = (delimiter === "_");
                updateFileNamePreview();
            };

            fileName.useDocName.onClick = updateFileNamePreview;
            fileName.ignoreDocName.onClick = function() {
                /* ドキュメント名を使わないときは接尾辞だけが頼りになる / Without the document name, only the suffix identifies the file */
                fileName.setDelimiter("");
                fileName.setSuffix(fileName.suffixInput.text);
                fileName.suffixInput.active = true;
            };
            fileName.delimiterNone.onClick = updateFileNamePreview;
            fileName.delimiterDash.onClick = updateFileNamePreview;
            fileName.delimiterUnderscore.onClick = updateFileNamePreview;
            fileName.suffixEnabled.onClick = function() {
                fileName.suffixInput.enabled = fileName.suffixEnabled.value;
                updateFileNamePreview();
            };
            fileName.suffixInput.onChange = updateFileNamePreview;

            return fileName;
        }

        /**
         * 書き出し先パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 書き出し先のコントロール一式
         */
        function buildLocationPanel(parentColumn) {
            var locationPanel = addPanel(parentColumn, LABELS.panel.location);
            var location = {};

            var folderRow = addRow(locationPanel);
            location.desktop = folderRow.add("radiobutton", undefined, getLabel(LABELS.radio.desktop));
            location.documentFolder = folderRow.add("radiobutton", undefined, getLabel(LABELS.radio.documentFolder));
            location.desktop.value = true;

            var showFolderRow = addRow(locationPanel);
            location.showFolder = showFolderRow.add("checkbox", undefined, getLabel(LABELS.checkbox.showFolder));
            location.showFolder.value = true;
            /* フォルダーを開く処理は macOS のみ / Opening the folder is macOS only */
            showFolderRow.visible = (Folder.fs === "Macintosh");

            /* 書き出し先を選択する / Select the destination */
            location.select = function(locationName) {
                location.desktop.value = (locationName === "desktop");
                location.documentFolder.value = (locationName !== "desktop");
            };

            return location;
        }

        /**
         * プリセットの内容をダイアログへ反映する
         * @param {Object} preset - プリセット定義
         * @returns {void}
         */
        function applyPreset(preset) {
            controls.background.select(preset.background);
            controls.margin.select(preset.margin);
            controls.margin.selectRoundMode(preset.round);
            controls.border.select(preset.border);
            controls.location.select(preset.location);
            controls.size.select(preset.size, true);
            controls.fileName.setDelimiter(preset.delimiter);
            if (preset.suffix) controls.fileName.setSuffix(preset.suffix);
            else controls.fileName.clearSuffix();
            refreshPreview();
        }

        presetDropdown.onChange = function() {
            var selectedIndex = presetDropdown.selection.index;
            /* 先頭は「カスタム」なので何も反映しない / The first entry is "Custom" and applies nothing */
            if (selectedIndex > 0) applyPreset(PRESETS[selectedIndex - 1]);
        };
        btnSavePreset.onClick = function() {
            savePresetToFile(controls);
        };

        // -----------------------------------------
        // ボタンエリア / Button row
        // -----------------------------------------
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        dialog.defaultElement = btnOK;
        dialog.cancelElement = btnCancel;

        // -----------------------------------------
        // 初期状態とプレビュー / Initial state and preview
        // -----------------------------------------
        controls.fileName.setDelimiter("-");
        controls.size.select("scale:" + DEFAULT_SCALE);

        refreshPreview();

        dialog.addEventListener("show", function() {
            updateFileNamePreview();
            dialog.center();
            dialog.location = [dialog.location[0] + DIALOG_OFFSET_X, dialog.location[1]];
        });

        var dialogResult = dialog.show();

        /* プレビュー用の描画は、OK・キャンセルどちらでも必ず消す / Always remove the preview artwork, whichever button was used */
        removePreviewArtwork();

        return (dialogResult === 1) ? collectExportSettings(controls, doc) : null;
    }

    // =========================================
    // 書き出し / Export
    // =========================================

    /**
     * サイズ指定から書き出し倍率を求める
     * @param {string} sizeSpec - サイズ指定（scale:200 / width:1000）
     * @param {number} exportWidthPt - 書き出し範囲の幅（pt）
     * @returns {number} 書き出し倍率（%）
     */
    function resolveExportScale(sizeSpec, exportWidthPt) {
        var specValue = toNumber(String(sizeSpec).split(":")[1]);
        var scalePercent = specValue;

        /* 倍率100%で1pt＝1pxになるため、横幅指定はpt値との比で倍率を求める / At 100% one pt equals one px, so a target width is a ratio against the pt size */
        if (String(sizeSpec).indexOf("width:") === 0) {
            scalePercent = (exportWidthPt > 0) ? (specValue / exportWidthPt * 100) : 100;
        }
        return (scalePercent > 0) ? scalePercent : 100;
    }

    /**
     * 書き出し倍率を指定できる範囲に収める
     * @param {number} scalePercent - 求めた倍率（%）
     * @returns {number} 範囲内に収めた倍率（%）
     */
    function limitExportScale(scalePercent) {
        return Math.min(scalePercent, MAX_EXPORT_SCALE);
    }

    /**
     * 一時アートボードを作ってPNG書き出しする
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} settings - 書き出し設定
     * @param {Object} exportRect - 書き出し範囲
     * @param {{label: string, factor: number}} rulerUnit - 定規の単位情報
     * @returns {void}
     */
    function exportAsPng(doc, settings, exportRect, rulerUnit) {
        var temporaryArtboardIndex = doc.artboards.length;
        doc.artboards.add([exportRect.left, exportRect.top, exportRect.right, exportRect.bottom]);
        doc.artboards.setActiveArtboardIndex(temporaryArtboardIndex);

        var requestedScale = resolveExportScale(settings.sizeSpec, exportRect.width);
        var exportScale = limitExportScale(requestedScale);

        /* 背景・罫線・一時アートボードは、書き出しが失敗しても必ず片付ける
           / The background, border and temporary artboard go away even when the export fails */
        try {
            createExportBackground(settings.backgroundChoice, exportRect, settings.checkerPercent);
            var borderRect = drawBorderRectangle(
                exportRect,
                resolveBorderWidth(settings.borderSpec, rulerUnit.factor),
                resolveBorderColor(settings.borderSpec)
            );
            if (borderRect) borderRect.zOrder(ZOrderMethod.BRINGTOFRONT);

            var exportOptions = new ExportOptionsPNG24();
            exportOptions.artBoardClipping = true;
            exportOptions.transparency = (settings.backgroundChoice === "transparent");
            exportOptions.horizontalScale = exportScale;
            exportOptions.verticalScale = exportScale;

            var exportFile = new File(settings.destinationFolder + "/" + settings.fileName);
            try {
                doc.exportFile(exportFile, ExportType.PNG24, exportOptions);
            } catch (e) {
                alert(getLabel(LABELS.alert.exportFailed) + "\n" + e.message);
            }
        } finally {
            removePreviewArtwork();
            doc.artboards.remove(temporaryArtboardIndex);
        }

        if (exportScale < requestedScale) {
            alert(getLabel(LABELS.alert.scaleLimited) + Math.floor(exportScale) + "%");
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0 || app.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var doc = app.activeDocument;
        var rulerUnit = getRulerUnitInfo();
        var documentBaseName = doc.name.replace(/\.ai$/i, "");
        var originalArtboardIndex = doc.artboards.getActiveArtboardIndex();

        /* 選択オブジェクトだけを写した作業用レイヤーで、他のオブジェクトを写り込ませずにプレビューする
           / Work on a copy of the selection so nothing else shows up in the preview */
        var previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;
        var previewItems = duplicateSelectionToLayer(doc.selection, previewLayer);
        var hiddenLayers = hideOtherLayers(doc, previewLayer);
        var selectionBounds = getSelectionBounds(previewItems);

        var settings = null;
        /* 途中で失敗しても、レイヤーを隠したままドキュメントを放置しない
           / Never leave the document with its layers hidden, whatever fails on the way */
        try {
            settings = showExportOptionsDialog(selectionBounds, rulerUnit, documentBaseName);

            if (settings) {
                var exportRect = buildExportRect(selectionBounds, resolveMarginOffsets(settings.marginSpec, rulerUnit.factor),
                    settings.roundMode, rulerUnit.factor);
                if (exportRect.width > 0 && exportRect.height > 0) {
                    exportAsPng(doc, settings, exportRect, rulerUnit);
                } else {
                    alert(getLabel(LABELS.alert.invalidSize));
                    settings = null;
                }
            }
        } finally {
            /* 作業用レイヤー・レイヤー表示・アクティブアートボードを元に戻す / Undo the temporary layer, visibility and active artboard */
            removePreviewArtwork();
            removePreviewLayerByName(doc, PREVIEW_LAYER_NAME);
            restoreLayerVisibility(hiddenLayers);
            doc.artboards.setActiveArtboardIndex(originalArtboardIndex);
            app.redraw();
        }

        if (settings && settings.showFolder && Folder.fs === "Macintosh") {
            settings.destinationFolder.execute();
        }
    }

    main();

})();
