#target illustrator
#targetengine "TileInsideShapeEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.PatternFill
Name=大きいオブジェクトの形の内側に小さいオブジェクトを敷き詰める
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=2.0.0
Release-Date=2025-10-26
Description-BEGIN
オブジェクトを2つ以上選択し、いちばん大きいものの形の内側に、ほかのオブジェクトを複製して敷き詰めます。
形からはみ出すものは残さず（マスクして残すこともできます）、元のオブジェクトはそのまま残します。

With two or more objects selected, duplicates the smaller ones inside the shape of the largest one.
Copies that stick out of the shape are removed (or kept under a mask), and the originals are left in place.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

オブジェクトを2つ以上選択し、いちばん大きいものの形の内側に、ほかのオブジェクトを複製して敷き詰めます。
形からはみ出すものは残さず（マスクして残すこともできます）、元のオブジェクトはそのまま残します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TileInsideShape.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne393f59cbb77

### Overview

With two or more objects selected, duplicates the smaller ones inside the shape of the largest one.
Copies that stick out of the shape are removed (or kept under a mask), and the originals are left in place.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TileInsideShape.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TileInsideShape";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-27";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "大きいオブジェクトの形の内側に小さいオブジェクトを敷き詰める"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TileInsideShape.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TileInsideShape.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne393f59cbb77"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/TileInsideShape.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 間隔の初期値を決める比率（タイルの幅に掛ける）/ Initial spacing as a ratio of the tile width */
    var DEFAULT_GAP_RATIO = 0.2;

    /* 並べ方の初期値（"grid" グリッド / "brick" レンガ状（横）/ "brickVertical" レンガ状（縦）/ "random" ランダム）
       Initial layout: "grid", "brick" (rows), "brickVertical" (columns) or "random" */
    var DEFAULT_LAYOUT_TYPE = "brick";

    /* ［シンボル化］［プレビュー］［中央にそろえる］［はみ出しも残してマスク］の初期値
       Initial state of Symbolize, Preview, Center and Keep Overflow and Mask */
    var DEFAULT_SYMBOLIZE = true;
    var DEFAULT_PREVIEW = true;
    var DEFAULT_CENTER_ALIGN = true;
    var DEFAULT_CLIP_OVERFLOW = false;

    /* ［大きいオブジェクト］の初期値（"opacity" 不透明度 / "remove" 削除 / "guide" ガイド化）
       Initial choice for the larger object: "opacity", "remove" or "guide" */
    var DEFAULT_CONTAINER_ACTION = "opacity";

    /* スライダーの範囲（スケールは%、間隔・オフセットはタイルの幅に掛ける比率。入力欄には範囲外の値も入れられる）
       Slider ranges: scale in %, spacing and offset as ratios of the tile width; the fields accept values outside them */
    var SCALE_SLIDER_MIN = 10;
    var SCALE_SLIDER_MAX = 200;
    var SPACING_SLIDER_MAX_RATIO = 1;  /* 間隔は 0〜タイルの幅 / spacing: 0 to the tile width */
    var OFFSET_SLIDER_RANGE_RATIO = 1; /* オフセットは ±タイルの幅 / offset: plus or minus the tile width */
    var LAYOUT_ANGLE_SLIDER_RANGE = 90; /* 並びの角度は ±90° / layout angle: plus or minus 90° */

    /* ［ゆらぎ］の［大きさ］の上限（%）/ Upper limit of the Size jitter (%) */
    var SCALE_JITTER_MAX = 90;

    /* プレビューの不透明度（%）/ Opacity of the preview (%) */
    var PREVIEW_OPACITY = 60;

    /* 敷き詰めたタイルを入れるグループ名（プレビューは名前で探して消す）/ Group names; previews are found and removed by name */
    var GRID_GROUP_NAME = 'TiledGrid';
    var PREVIEW_GROUP_NAME = 'TiledGrid_preview';

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* ［OK］で確定した設定を Illustrator の終了まで覚えるキー（間隔・オフセットは pt で控える）
       Key for the settings confirmed with OK, kept until Illustrator quits (spacing and offset in pt) */
    var LAST_SETTINGS_KEY = "__" + SCRIPT_NAME + "_LastSettings";

    /**
     * 前回［OK］で確定した設定を返す
     * @returns {Object|null} 設定（readFillSettings() の形。乱数の種は除く）。まだ無ければ null
     */
    function loadLastSettings() {
        return $.global[LAST_SETTINGS_KEY] || null;
    }

    /**
     * ［OK］で確定した設定を控える
     * @param {Object} fillSettings - 設定（readFillSettings() の形。スケールは%、長さは pt）
     * @returns {void}
     */
    function saveLastSettings(fillSettings) {
        /* 乱数の種は控えない（開くたびに新しい並びにする）/ The random seed is not kept, so each run gets a new arrangement */
        var savedSettings = {};
        for (var settingName in fillSettings) {
            if (fillSettings.hasOwnProperty(settingName) && settingName !== 'randomSeed') savedSettings[settingName] = fillSettings[settingName];
        }
        $.global[LAST_SETTINGS_KEY] = savedSettings;
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
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
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    /* ダイアログ固有の寸法 / Dialog-specific sizes */
    var LABEL_WIDTH = 80;       /* 項目名の幅 / Field label width */
    var FIELD_CHARACTERS = 4;     /* 数値入力欄の桁数 / Width of the number fields in characters */
    var LENGTH_FIELD_CHARACTERS = 6; /* 単位付きの数値入力欄（スケール・間隔・オフセット）の桁数 / Width of the fields that show a unit */
    var PERCENT_UNIT = ' %';         /* スケールの欄の中に付ける単位 / Unit shown inside the scale field */
    var SLIDER_INTEGER_RANGE = 50;   /* スライダーの範囲がこれより広ければ整数刻み、狭ければ 0.1 刻み / Wider slider ranges step by 1, narrower by 0.1 */
    var SLIDER_WIDTH = 300;          /* スライダーの幅（すべての行で同じ長さ）/ Slider width, the same in every row */
    var SHIFT_SLIDER_RANGE = 50;     /* ［調整X］［調整Y］のスライダーの端で、並びの周期の何%ずらすか（中央が0）/ % of the pattern period at each end of the Adjust X/Y sliders (0 at the center) */
    var RADIO_TEXT_INDENT = 18;      /* ラジオの下に置く欄の字下げ（ラジオの文字の頭にそろえる）/ Indent of a field under a radio, lined up with the radio text */
    var TILE_LABEL_WIDTH = 110;      /* ［タイル］パネルの項目名の幅 / Label width in the Tiles panel */
    var DEGREE_UNIT = '°';           /* 角度の欄の中に付ける単位 / Unit shown inside the angle fields */
    var COUNT_TEXT_WIDTH = 80;       /* ［配置数］の値の幅（あとから伸びないので先に取る）/ Width reserved for the tile count, since controls do not grow */
    var COUNT_ROW_BOTTOM_MARGIN = 5; /* ［配置数］の行の下の余白 / Bottom margin under the tile count row */
    var GAP_LINK_SPACING = 6;        /* 間隔の2行と連動アイコンの間 / Gap between the spacing rows and the link icon */

    /**
     * 右揃えの項目名と入力欄を並べる行を作る
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {number} [labelWidth] - 項目名の幅（省略時は LABEL_WIDTH）
     * @returns {Group} 項目名を追加済みの行
     */
    function addFieldRow(parentContainer, labelPath, labelWidth) {
        var fieldRow = parentContainer.add('group');
        fieldRow.alignment = ['fill', 'top'];
        fieldRow.alignChildren = ['left', 'center'];
        var fieldLabel = fieldRow.add('statictext', undefined, labelText(labelPath));
        fieldLabel.justify = 'right';
        fieldLabel.preferredSize.width = labelWidth || LABEL_WIDTH;
        return fieldRow;
    }

    /**
     * 項目名・∧∨付きの数値欄・スライダーを並べた行を作る。欄とスライダーは同じ値を指し、どちらを動かしても onValueChange を呼ぶ
     * （スライダーの範囲外の値は欄にだけ入り、スライダーは端に寄る）。
     * sliderRange に toSlider / fromSlider を渡すと、欄の値とスライダーの値を換算する
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {string} initialText - 欄の初期値（単位付き）
     * @param {string} tooltipPath - helpTip のラベルのパス
     * @param {Object} stepOptions - ∧∨の設定（min / unit など。addStepper() を参照）
     * @param {{min: number, max: number, toSlider: Function, fromSlider: Function, fieldStep: number}} sliderRange - スライダーの範囲と、省略できる換算（欄の値→スライダー、スライダー→欄の値）・欄に書く刻み
     * @param {Function} onValueChange - 値が変わったときに呼ぶ関数
     * @returns {{row: Group, field: EditText, slider: Slider, syncSlider: Function}} 行・欄・スライダーと、欄の値をスライダーに写し直す関数
     */
    function addSliderFieldRow(parentContainer, labelPath, initialText, tooltipPath, stepOptions, sliderRange, onValueChange) {
        var fieldRow = addFieldRow(parentContainer, labelPath);
        var valueSlider;
        stepOptions.onStep = function () {
            syncSliderToField();
            onValueChange();
        };
        var numberField = addNumberField(fieldRow, initialText, tooltipPath, stepOptions);
        numberField.characters = LENGTH_FIELD_CHARACTERS;

        valueSlider = fieldRow.add('slider', undefined, sliderRange.min, sliderRange.min, sliderRange.max);
        valueSlider.preferredSize.width = SLIDER_WIDTH;
        valueSlider.alignment = ['left', 'center'];
        valueSlider.helpTip = numberField.helpTip;
        /* 幅の広い範囲は整数、狭い範囲は 0.1 刻みで欄に書く / Write whole numbers for wide ranges, tenths for narrow ones */
        var fieldStep = sliderRange.fieldStep || ((sliderRange.max - sliderRange.min > SLIDER_INTEGER_RANGE) ? 1 : 0.1);

        /**
         * 欄の値をスライダーに写す（範囲外なら端に寄せる）
         * @returns {void}
         */
        function syncSliderToField() {
            var fieldValue = parseFloat(numberField.text);
            if (isNaN(fieldValue)) return;
            var sliderValue = sliderRange.toSlider ? sliderRange.toSlider(fieldValue) : fieldValue;
            valueSlider.value = Math.max(sliderRange.min, Math.min(sliderRange.max, sliderValue));
        }

        valueSlider.onChanging = function () {
            var fieldValue = sliderRange.fromSlider ? sliderRange.fromSlider(valueSlider.value) : valueSlider.value;
            var sliderText = formatStepperNumber(Math.round(fieldValue / fieldStep) * fieldStep) + (stepOptions.unit || '');
            /* 刻みが変わらない細かい動きでは描き直さない / Skip redraws until the value changes by a step */
            if (sliderText === numberField.text) return;
            numberField.text = sliderText;
            onValueChange();
        };
        numberField.onChanging = function () {
            syncSliderToField();
            onValueChange();
        };
        /* 計算式を確定したあとの値も写す / Also follow the value after an expression is committed */
        numberField.onChange = numberField.onChanging;
        syncSliderToField();
        return { row: fieldRow, field: numberField, slider: valueSlider, syncSlider: syncSliderToField };
    }

    /**
     * カラムを横に並べた行を作る（各カラムは中身を上から縦に並べる）
     * @param {Window|Group} parentContainer - 追加先
     * @param {number} columnCount - カラムの数
     * @returns {{row: Group, columns: Group[]}} 行とカラム
     */
    function addColumnRow(parentContainer, columnCount) {
        var columnRow = parentContainer.add('group');
        columnRow.orientation = 'row';
        columnRow.alignChildren = ['fill', 'top'];
        columnRow.spacing = COLUMN_SPACING;
        var columns = [];
        for (var i = 0; i < columnCount; i++) {
            var column = columnRow.add('group');
            column.orientation = 'column';
            column.alignChildren = ['fill', 'top'];
            column.spacing = WINDOW_SPACING;
            columns.push(column);
        }
        return { row: columnRow, columns: columns };
    }

    /**
     * 行に数値入力欄を、左に∧∨を付けて追加する（↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} fieldRow - 追加先の行
     * @param {string} initialText - 初期値
     * @param {string} tooltipPath - helpTip のラベルのパス
     * @param {Object} stepOptions - ∧∨の設定（min / integer / onStep など。addStepper() を参照）
     * @returns {EditText} 追加した入力欄
     */
    function addNumberField(fieldRow, initialText, tooltipPath, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldRow.add('group');
        stepperFieldGroup.orientation = 'row';
        stepperFieldGroup.alignChildren = ['left', 'center'];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var numberField;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberField; }, stepOptions);
        numberField = stepperFieldGroup.add('edittext', undefined, initialText);
        numberField.characters = FIELD_CHARACTERS;
        numberField.helpTip = getLabel(tooltipPath);
        bindSteppedArrowKeys(numberField, stepperGroup);
        return numberField;
    }

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 }, /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 }, /* 歯 / H */
        "p": { unit: "pc", amount: 1 }     /* パイカ / pica */
    };

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) return;
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    var LINK_UI_DARK = isDarkUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @param {number[]} [iconSize] - アイコンの [幅, 高さ]（省略時は LINK_ICON_SIZE。絵は 22px 基準から拡大縮小する）
     * @param {number} [chainRatio] - 鎖の絵の大きさの比率（省略時は 1。枠・地の大きさは変えず、鎖だけ縮める）
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle, iconSize, chainRatio) {
        var toggleSize = iconSize || LINK_ICON_SIZE;
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = toggleSize;
        linkToggle.minimumSize = toggleSize;
        linkToggle.maximumSize = toggleSize;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = toggleSize[0];
            var iconHeight = toggleSize[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR, chainRatio);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @param {number} [chainRatio] - 鎖の大きさの比率（省略時は 1）。中央に置いたまま縮める
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

    /**
     * pt の値を単位の値に換算する
     * @param {number} ptValue - pt の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} 単位の値
     */
    function ptToUnit(ptValue, unit) {
        return ptValue / unit.pointsPerUnit;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトの内側に敷き詰め", en: "Tile Inside Shape" }
        },
        fieldLabel: {
            scale: { ja: "スケール", en: "Scale" },
            spacingX: { ja: "間隔（横）", en: "Spacing X" },
            spacingY: { ja: "間隔（縦）", en: "Spacing Y" },
            offset: { ja: "オフセット", en: "Offset" },
            shiftX: { ja: "調整X", en: "Adjust X" },
            shiftY: { ja: "調整Y", en: "Adjust Y" },
            layoutAngle: { ja: "並びの角度", en: "Angle" },
            count: { ja: "配置数", en: "Tiles" },
            density: { ja: "密度", en: "Density" },
            rotation: { ja: "個別に回転", en: "Rotate Each" },
            rotationJitter: { ja: "回転", en: "Rotation" },
            scaleJitter: { ja: "大きさ", en: "Size" },
            tileOrder: { ja: "選び方", en: "Order" }
        },
        panel: {
            layout: { ja: "並べ方", en: "Layout" },
            tile: { ja: "タイル", en: "Tiles" },
            containerAction: { ja: "敷き詰め先", en: "Fill Area" }
        },
        radio: {
            grid: { ja: "グリッド", en: "Grid" },
            brick: { ja: "レンガ状（横）", en: "Brick (Horizontal)" },
            brickVertical: { ja: "レンガ状（縦）", en: "Brick (Vertical)" },
            hexagon: { ja: "六角形", en: "Hexagonal" },
            random: { ja: "ランダム", en: "Random" },
            orderSequence: { ja: "順番", en: "In Order" },
            orderRandom: { ja: "ランダム", en: "Random" },
            opacityContainer: { ja: "不透明度", en: "Opacity" },
            removeContainer: { ja: "削除", en: "Delete" },
            guideContainer: { ja: "ガイド化", en: "Make Guides" }
        },
        checkbox: {
            centerAlign: { ja: "中央にそろえる", en: "Center" },
            clipOverflow: { ja: "はみ出しも残してマスク", en: "Keep Overflow and Mask" },
            alternateFlip: { ja: "1つおきに180°回転", en: "Flip Every Other 180°" },
            jitter: { ja: "ゆらぎ", en: "Jitter" },
            preview: { ja: "プレビュー", en: "Preview" },
            symbolize: { ja: "シンボル化", en: "Symbolize" }
        },
        tooltip: {
            scale: {
                ja: "敷き詰めるタイルの大きさです。縦横比はそのままで拡大・縮小します。",
                en: "Size of the placed tiles, scaled with the aspect ratio kept."
            },
            spacingX: { ja: "横に隣り合うタイルの間隔です。", en: "Horizontal space between neighboring tiles." },
            spacingY: { ja: "縦に隣り合うタイルの間隔です。", en: "Vertical space between neighboring tiles." },
            linkSpacing: { ja: "横と縦の間隔を連動させます。", en: "Keeps the horizontal and vertical spacing the same." },
            offset: {
                ja: "敷き詰め先の形をこの量だけオフセットした内側に収めます。負の値で内側に縮め、正の値でははみ出しを許します。",
                en: "Fits the tiles inside the fill area's shape offset by this amount. Negative values shrink it; positive values let tiles stick out."
            },
            shiftX: {
                ja: "敷き詰める位置を右へずらします。負の値で左へずらします。スライダーは中央がずらさない位置で、端で並びの半周期ずれます。",
                en: "Shifts the tiles to the right. Negative values shift them to the left. The slider's center leaves them in place; each end shifts by half the pattern period."
            },
            shiftY: {
                ja: "敷き詰める位置を下へずらします。負の値で上へずらします。スライダーは中央がずらさない位置で、右へ動かすと下へ、端で並びの半周期ずれます。",
                en: "Shifts the tiles down. Negative values shift them up. The slider's center leaves them in place; moving it right shifts them down, by up to half the pattern period."
            },
            layoutAngle: {
                ja: "並び全体を、敷き詰め先の中心を軸にこの角度だけ回します。タイルも同じ角度で回ります。",
                en: "Rotates the whole arrangement by this angle around the center of the fill area. The tiles turn with it."
            },
            count: { ja: "今の設定で置かれるタイルの数です。", en: "Number of tiles placed with the current settings." },
            grid: { ja: "縦横にそろえて並べます。", en: "Lines the tiles up in rows and columns." },
            brick: {
                ja: "1行おきに半個分横にずらして、レンガのように並べます。",
                en: "Offsets every other row sideways by half a tile, like brickwork."
            },
            brickVertical: {
                ja: "1列おきに半個分縦にずらして、縦に積んだレンガのように並べます。",
                en: "Offsets every other column vertically by half a tile, like brickwork turned on its side."
            },
            hexagon: {
                ja: "1行おきに半個分ずらし、行の間を詰めて六角形に並べます。円形のタイル向きです。",
                en: "Offsets every other row by half a tile and closes up the rows into a hexagonal pattern. Best for round tiles."
            },
            random: {
                ja: "間隔を空けて重ならないように、ランダムな位置に並べます。ダイアログボックスを開くたび、または［再配置］で並びが変わります。［調整X］［調整Y］［並びの角度］［中央にそろえる］は使いません。",
                en: "Places the tiles at random positions without overlapping, keeping the spacing. The arrangement changes each time the dialog opens or with Reshuffle. Adjust X, Adjust Y, Angle and Center are not used."
            },
            density: {
                ja: "ランダムに置いたタイルのうち、残す割合です。100% で詰められるだけ詰めます。",
                en: "Share of the randomly placed tiles to keep. At 100% as many as fit are placed."
            },
            centerAlign: {
                ja: "並びを敷き詰め先の中心にそろえ、上下左右の余白を均等にします。中心にタイルを置くか間を置くかは、多く収まる方を選びます。",
                en: "Centers the arrangement in the fill area so the margins are even. A tile or a gap goes at the center, whichever fits more tiles."
            },
            clipOverflow: {
                ja: "形からはみ出すタイルも残し、敷き詰め先の形でマスクします。［オフセット］は使いません。",
                en: "Keeps tiles that stick out of the shape and masks them with the fill area's shape. Offset is not used."
            },
            reshuffle: {
                ja: "ランダムな並び・ばらつき・タイルの選び方を引き直します。",
                en: "Draws a new random arrangement, jitter and tile order."
            },
            rotation: {
                ja: "オンにすると、タイルをそれぞれの中心でこの角度だけ回転させます。オフでも値は残ります。",
                en: "When on, rotates each tile around its own center by this angle. The value is kept while it is off."
            },
            alternateFlip: {
                ja: "1つおきにタイルを180°回します。三角形を互い違いに並べるときなどに使います。",
                en: "Turns every other tile by 180°, e.g. to interlock triangles."
            },
            jitter: {
                ja: "オンにすると、下の［回転］［大きさ］のばらつきを使います。オフでも値は残ります。",
                en: "When on, the Rotation and Size jitter below are used. The values are kept while it is off."
            },
            rotationJitter: {
                ja: "タイルごとに、この角度の範囲（±）でランダムに回します。",
                en: "Rotates each tile by a random angle within this range (±)."
            },
            scaleJitter: {
                ja: "タイルごとに、この割合までランダムに小さくします。",
                en: "Shrinks each tile at random by up to this percentage."
            },
            tileOrder: {
                ja: "敷き詰めるオブジェクトを複数選んだとき、どの順で使うかを選びます。［順番］は選択の前面から順に使います。",
                en: "When several objects to tile are selected, chooses how they are used. In Order cycles through them from the frontmost."
            },
            preview: { ja: "設定をドキュメント上にプレビューします。", en: "Previews the result in the document." },
            symbolize: {
                ja: "タイルをシンボルとして複製します。あとからまとめて差し替えられます。",
                en: "Duplicates the tile as a symbol, so every copy can be swapped later at once."
            },
            opacityContainer: {
                ja: "敷き詰めたあと、敷き詰め先の不透明度をこの値にします。100% なら何も変えません。",
                en: "Sets the fill area's opacity to this value after tiling. At 100% it is left unchanged."
            },
            removeContainer: { ja: "敷き詰めたあと、敷き詰め先を削除します。", en: "Deletes the fill area after tiling." },
            guideContainer: {
                ja: "敷き詰めたあと、敷き詰め先をガイドに変換します。",
                en: "Converts the fill area into guides after tiling."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開いていません。", en: "No document is open." },
            selectTwo: {
                ja: "敷き詰め先と、敷き詰めるオブジェクトを1つ以上選択してください。いちばん大きいものが敷き詰め先になります。",
                en: "Select the fill area and at least one object to tile. The largest object becomes the fill area."
            },
            emptyTile: { ja: "敷き詰めるオブジェクトの幅か高さが0です。", en: "An object to tile has no width or height." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reshuffle: { ja: "再配置", en: "Reshuffle" }
        }
    };

    // =========================================
    // 形状と配置の計算 / Geometry
    // =========================================

    /**
     * オブジェクトの表示上の境界と寸法を取得する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number, area: number}} 境界情報
     */
    function getBoundsInfo(pageItem) {
        var visibleBounds = pageItem.visibleBounds; /* [left, top, right, bottom] */
        var itemWidth = visibleBounds[2] - visibleBounds[0];
        var itemHeight = visibleBounds[1] - visibleBounds[3];
        return {
            left: visibleBounds[0],
            top: visibleBounds[1],
            right: visibleBounds[2],
            bottom: visibleBounds[3],
            width: itemWidth,
            height: itemHeight,
            area: Math.abs(itemWidth * itemHeight)
        };
    }

    /* 並べ方ごとのずらし方（"rows" は1行おきに横、"columns" は1列おきに縦）
       How each layout offsets: "rows" shifts every other row, "columns" every other column */
    var BRICK_DIRECTIONS = { brick: 'rows', brickVertical: 'columns', hexagon: 'rows' };

    /* 六角形で行の送りに掛ける比率（√3/2。円を隙間なく詰める）/ Row step ratio for the hexagonal layout (√3/2, packs circles tightly) */
    var HEXAGON_ROW_RATIO = Math.sqrt(3) / 2;

    /* ランダムに並べるとき、グリッドのマス1つあたりに位置を試す回数 / Positions tried per grid cell when placing at random */
    var RANDOM_ATTEMPTS_PER_CELL = 20;

    /* 回転していないとみなす角度の差（度）/ Angle difference (degrees) treated as no rotation */
    var ANGLE_TOLERANCE = 1e-6;

    /* 回した形を角度ごとに控える数の上限（超えたら作り直す）/ Max rotated shapes kept per angle before the cache is rebuilt */
    var ROTATED_SHAPE_CACHE_LIMIT = 64;

    /**
     * 種から毎回同じ順に 0 以上 1 未満の乱数を返す関数を作る（プレビューと実行で同じ並びにする）
     * @param {number} seed - 種（1〜2147483646 の整数）
     * @returns {Function} 呼ぶたびに次の乱数を返す関数
     */
    function createSeededRandom(seed) {
        var state = seed;
        return function () {
            /* Park-Miller 法。積は 2^53 未満なので誤差は出ない / Park-Miller; the product stays below 2^53, so it is exact */
            state = (state * 16807) % 2147483647;
            return (state - 1) / 2147483646;
        };
    }

    /**
     * 種から、用途ごとに別の乱数列を引くための種を作る
     * @param {number} seed - 元の種（1〜2147483646 の整数）
     * @param {number} salt - 用途ごとの番号
     * @returns {number} 新しい種（1〜2147483646 の整数）
     */
    function deriveSeed(seed, salt) {
        return ((seed - 1 + salt * 7919) % 2147483646) + 1;
    }

    /**
     * 整数が奇数か判定する（負の数も扱う）
     * @param {number} value - 整数
     * @returns {boolean} 奇数なら true
     */
    function isOdd(value) {
        return ((value % 2) + 2) % 2 === 1;
    }

    /**
     * 点を原点のまわりに回す（反時計回りが正）
     * @param {number} x - X 座標
     * @param {number} y - Y 座標
     * @param {number} angleDegrees - 角度（度）
     * @returns {number[]} 回した点の [x, y]
     */
    function rotatePoint(x, y, angleDegrees) {
        var radians = angleDegrees * Math.PI / 180;
        var cosValue = Math.cos(radians);
        var sinValue = Math.sin(radians);
        return [x * cosValue - y * sinValue, x * sinValue + y * cosValue];
    }

    /**
     * 角度を -180°〜180° に収める
     * @param {number} angleDegrees - 角度（度）
     * @returns {number} 収めた角度
     */
    function normalizeFullTurn(angleDegrees) {
        var angle = ((angleDegrees % 360) + 360) % 360;
        return (angle > 180) ? angle - 360 : angle;
    }

    /**
     * 角度を -90°〜90° に収める（長方形は180°回しても同じ形になる）
     * @param {number} angleDegrees - 角度（度）
     * @returns {number} 収めた角度
     */
    function normalizeHalfTurn(angleDegrees) {
        var angle = ((angleDegrees % 180) + 180) % 180;
        return (angle > 90) ? angle - 180 : angle;
    }

    /**
     * ［個別に回転］の角度を返す（オフなら欄の値にかかわらず 0）
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @returns {number} 角度（度）
     */
    function getTileRotation(fillSettings) {
        return fillSettings.tileRotationOn ? fillSettings.tileRotation : 0;
    }

    /**
     * タイルの大きさから並びの送りを決める。
     * 送りはいちばん大きいタイルを［回転］で回した外接矩形に間隔を足したもの（六角形は行の送りを詰める）
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @returns {{footWidth: number, footHeight: number, stepX: number, stepY: number, maxRadius: number}|null} 送りなど（pt）。決められなければ null
     */
    function getLayoutMetrics(fillTarget, fillSettings) {
        var scaleRatio = fillSettings.scale / 100;
        var rotationRadians = getTileRotation(fillSettings) * Math.PI / 180;
        var cosAbs = Math.abs(Math.cos(rotationRadians));
        var sinAbs = Math.abs(Math.sin(rotationRadians));
        var footWidth = 0;
        var footHeight = 0;
        var maxRadius = 0;
        for (var i = 0; i < fillTarget.tiles.length; i++) {
            var tileWidth = fillTarget.tiles[i].info.width * scaleRatio;
            var tileHeight = fillTarget.tiles[i].info.height * scaleRatio;
            footWidth = Math.max(footWidth, tileWidth * cosAbs + tileHeight * sinAbs);
            footHeight = Math.max(footHeight, tileWidth * sinAbs + tileHeight * cosAbs);
            maxRadius = Math.max(maxRadius, Math.sqrt(tileWidth * tileWidth + tileHeight * tileHeight) / 2);
        }
        var stepX = footWidth + fillSettings.gapX;
        var stepY = footHeight + fillSettings.gapY;
        if (!(stepX > 0) || !(stepY > 0)) return null;
        if (fillSettings.layoutType === 'hexagon') stepY *= HEXAGON_ROW_RATIO;
        return { footWidth: footWidth, footHeight: footHeight, stepX: stepX, stepY: stepY, maxRadius: maxRadius };
    }

    /**
     * タイルの中心を置く範囲を、容器の外接矩形から広げる量を返す
     * （マスクするときははみ出すタイルも拾うのでタイルの半径、収めるときは正のオフセット）
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @param {Object} layoutMetrics - getLayoutMetrics() の結果
     * @returns {number} 広げる量（pt）
     */
    function getCoverMargin(fillSettings, layoutMetrics) {
        if (fillSettings.clipOverflow) return layoutMetrics.maxRadius;
        return Math.max(0, fillSettings.offset);
    }

    /**
     * 矩形を広げて中心のまわりに回し、回したあとの外接矩形を中心からの座標で返す
     * @param {Object} boundsInfo - 境界情報
     * @param {number} centerX - 中心の X 座標
     * @param {number} centerY - 中心の Y 座標
     * @param {number} angleDegrees - 回す角度（度）
     * @param {number} margin - 上下左右に広げる量（pt）
     * @returns {{left: number, top: number, right: number, bottom: number}} 中心を原点にした外接矩形
     */
    function getRotatedBox(boundsInfo, centerX, centerY, angleDegrees, margin) {
        var corners = [
            [boundsInfo.left - margin, boundsInfo.top + margin],
            [boundsInfo.right + margin, boundsInfo.top + margin],
            [boundsInfo.right + margin, boundsInfo.bottom - margin],
            [boundsInfo.left - margin, boundsInfo.bottom - margin]
        ];
        var rotatedBox = null;
        for (var i = 0; i < corners.length; i++) {
            var point = rotatePoint(corners[i][0] - centerX, corners[i][1] - centerY, angleDegrees);
            if (!rotatedBox) {
                rotatedBox = { left: point[0], right: point[0], top: point[1], bottom: point[1] };
                continue;
            }
            rotatedBox.left = Math.min(rotatedBox.left, point[0]);
            rotatedBox.right = Math.max(rotatedBox.right, point[0]);
            rotatedBox.top = Math.max(rotatedBox.top, point[1]);
            rotatedBox.bottom = Math.min(rotatedBox.bottom, point[1]);
        }
        return rotatedBox;
    }

    /**
     * 並べ方に応じて、タイルを置く位置をまとめて決める（容器に収まるもの、マスクするときは形に触れるものだけ）
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @returns {Object[]} 置く位置の配列（{x, y, angle, scale, tileIndex}。x・y はタイルの中心、angle は度、scale は%）
     */
    function computePlacements(fillTarget, fillSettings) {
        var layoutMetrics = getLayoutMetrics(fillTarget, fillSettings);
        if (!layoutMetrics) return [];
        /* マスクするときはオフセットを使わない / Masking ignores the offset */
        var containerShape = buildContainerShape(fillTarget.container, fillSettings.clipOverflow ? 0 : fillSettings.offset);
        if (fillSettings.layoutType === 'random') {
            return computeRandomPlacements(fillTarget, fillSettings, layoutMetrics, containerShape);
        }
        if (!fillSettings.centerAlign) {
            return computeLatticePlacements(fillTarget, fillSettings, layoutMetrics, containerShape, null);
        }
        /* 中心にタイルを置くか、タイルの間を置くかを縦横で試し、いちばん多く収まる並びにする
           Try a tile or a gap at the center on each axis and keep the arrangement that fits the most */
        var centerPhases = [[0, 0], [0.5, 0], [0, 0.5], [0.5, 0.5]];
        var bestPlacements = [];
        for (var i = 0; i < centerPhases.length; i++) {
            var candidatePlacements = computeLatticePlacements(fillTarget, fillSettings, layoutMetrics, containerShape, centerPhases[i]);
            if (candidatePlacements.length > bestPlacements.length) bestPlacements = candidatePlacements;
        }
        return bestPlacements;
    }

    /**
     * グリッド・レンガ状・六角形の並びで、タイルを置く位置を決める。
     * 並びは容器の中心を原点に［並びの角度］だけ回した座標で作り、容器の外接矩形を覆う範囲を調べる
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @param {Object} layoutMetrics - getLayoutMetrics() の結果
     * @param {Object} containerShape - buildContainerShape() の結果
     * @param {number[]|null} centerPhase - 中央にそろえるときの起点（送りに対する割合 [横, 縦]）。null なら左上の角にそろえる
     * @returns {Object[]} 置く位置の配列
     */
    function computeLatticePlacements(fillTarget, fillSettings, layoutMetrics, containerShape, centerPhase) {
        var containerInfo = fillTarget.containerInfo;
        var centerX = (containerInfo.left + containerInfo.right) / 2;
        var centerY = (containerInfo.top + containerInfo.bottom) / 2;
        var layoutAngle = fillSettings.layoutAngle;
        var stepX = layoutMetrics.stepX;
        var stepY = layoutMetrics.stepY;
        var brickDirection = BRICK_DIRECTIONS[fillSettings.layoutType] || '';

        /* 並びの向きで測った容器の外接矩形（中心が原点）/ The container's bounds measured along the layout, centered on the origin */
        var containerBox = getRotatedBox(containerInfo, centerX, centerY, -layoutAngle, 0);
        var coverBox = getRotatedBox(containerInfo, centerX, centerY, -layoutAngle, getCoverMargin(fillSettings, layoutMetrics));

        /* 起点：中央そろえは中心から、そうでなければ外接矩形の左上の角にタイルの角を合わせる
           Origin: from the center when centering, otherwise a tile corner on the top-left corner of the bounds */
        var originX;
        var originY;
        if (centerPhase) {
            originX = centerPhase[0] * stepX;
            originY = -centerPhase[1] * stepY;
        } else {
            originX = containerBox.left + layoutMetrics.footWidth / 2;
            originY = containerBox.top - layoutMetrics.footHeight / 2;
        }
        originX += fillSettings.shiftX;
        originY -= fillSettings.shiftY;

        /* 覆う範囲に1マスずつ余分を足して、ずらした行・列の端も拾う / One extra cell on each side catches the ends of offset rows and columns */
        var firstColumn = Math.floor((coverBox.left - originX) / stepX) - 1;
        var lastColumn = Math.ceil((coverBox.right - originX) / stepX) + 1;
        var firstRow = Math.floor((originY - coverBox.top) / stepY) - 1;
        var lastRow = Math.ceil((originY - coverBox.bottom) / stepY) + 1;

        var placementBuilder = createPlacementBuilder(fillTarget, fillSettings, containerShape);
        for (var rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            for (var columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++) {
                var localX = originX + columnIndex * stepX + ((brickDirection === 'rows' && isOdd(rowIndex)) ? stepX / 2 : 0);
                var localY = originY - rowIndex * stepY - ((brickDirection === 'columns' && isOdd(columnIndex)) ? stepY / 2 : 0);
                var worldPoint = rotatePoint(localX, localY, layoutAngle);
                var candidate = placementBuilder.makeCandidate(centerX + worldPoint[0], centerY + worldPoint[1], layoutAngle, isOdd(rowIndex + columnIndex));
                if (placementBuilder.isFitting(candidate)) placementBuilder.accept(candidate);
            }
        }
        return placementBuilder.placements;
    }

    /**
     * 容器の外接矩形の中にランダムな位置を選び、置いたタイルと間隔を空けて重ならず、容器に収まる位置を決める。
     * ［密度］が100%未満なら、置いた中からその割合だけ残す
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @param {Object} layoutMetrics - getLayoutMetrics() の結果
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {Object[]} 置く位置の配列
     */
    function computeRandomPlacements(fillTarget, fillSettings, layoutMetrics, containerShape) {
        var containerInfo = fillTarget.containerInfo;
        var coverMargin = getCoverMargin(fillSettings, layoutMetrics);
        var minX = containerInfo.left - coverMargin;
        var maxX = containerInfo.right + coverMargin;
        var minY = containerInfo.bottom - coverMargin;
        var maxY = containerInfo.top + coverMargin;
        var nextRandom = createSeededRandom(deriveSeed(fillSettings.randomSeed, 1));
        var placementBuilder = createPlacementBuilder(fillTarget, fillSettings, containerShape);

        var cellCount = Math.ceil((maxX - minX) / layoutMetrics.stepX) * Math.ceil((maxY - minY) / layoutMetrics.stepY);
        var attemptCount = Math.max(1, cellCount) * RANDOM_ATTEMPTS_PER_CELL;

        /* 置いたタイルを、間隔込みの外接円の直径を1マスにして振り分け、重なりは周りの9マスだけ調べる
           Bucket placed tiles by the diameter of their spaced bounding circle, so overlaps are checked against the 9 neighbors only */
        var bucketSize = (layoutMetrics.maxRadius + Math.max(fillSettings.gapX, fillSettings.gapY)) * 2;
        var footprintsByBucket = {};

        /**
         * 置いたタイルのどれかと、間隔を含めて重なるか判定する
         * @param {Object} footprint - makeFootprint() の結果
         * @param {number} bucketColumn - マスの列
         * @param {number} bucketRow - マスの行
         * @returns {boolean} 重なれば true
         */
        function overlapsPlacedTile(footprint, bucketColumn, bucketRow) {
            for (var i = bucketColumn - 1; i <= bucketColumn + 1; i++) {
                for (var j = bucketRow - 1; j <= bucketRow + 1; j++) {
                    var bucketFootprints = footprintsByBucket[i + ',' + j];
                    if (!bucketFootprints) continue;
                    for (var k = 0; k < bucketFootprints.length; k++) {
                        if (doFootprintsOverlap(bucketFootprints[k], footprint)) return true;
                    }
                }
            }
            return false;
        }

        for (var attemptIndex = 0; attemptIndex < attemptCount; attemptIndex++) {
            var candidate = placementBuilder.makeCandidate(minX + (maxX - minX) * nextRandom(), minY + (maxY - minY) * nextRandom(),
                0, isOdd(placementBuilder.placements.length));
            var footprint = makeFootprint(candidate, fillTarget.tiles[candidate.tileIndex].info, fillSettings);
            var bucketColumn = Math.floor((candidate.x - minX) / bucketSize);
            var bucketRow = Math.floor((candidate.y - minY) / bucketSize);
            /* 軽い重なりの判定を先にする / Run the cheap overlap check first */
            if (overlapsPlacedTile(footprint, bucketColumn, bucketRow)) continue;
            if (!placementBuilder.isFitting(candidate)) continue;
            placementBuilder.accept(candidate);
            var bucketKey = bucketColumn + ',' + bucketRow;
            if (!footprintsByBucket[bucketKey]) footprintsByBucket[bucketKey] = [];
            footprintsByBucket[bucketKey].push(footprint);
        }

        var placements = placementBuilder.placements;
        if (fillSettings.density >= 100) return placements;
        /* 置いた中から密度の割合だけ残す（別の乱数列で選ぶ）/ Keep the density's share of the placed tiles, picked with a separate sequence */
        var thinningRandom = createSeededRandom(deriveSeed(fillSettings.randomSeed, 2));
        var keptPlacements = [];
        for (var i = 0; i < placements.length; i++) {
            if (thinningRandom() * 100 < fillSettings.density) keptPlacements.push(placements[i]);
        }
        return keptPlacements;
    }

    /**
     * 置く位置の候補を作り、収まるか調べて控えるための道具をまとめる。
     * 乱数（タイルの選び方・回転・大きさのばらつき）は、置けるかどうかに関係なく候補ごとに同じ順で引く
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {{placements: Object[], makeCandidate: Function, isFitting: Function, accept: Function}} 控えた位置の配列と、候補を作る・調べる・控える関数
     */
    function createPlacementBuilder(fillTarget, fillSettings, containerShape) {
        var nextRandom = createSeededRandom(fillSettings.randomSeed);
        var tileCount = fillTarget.tiles.length;
        var placements = [];
        /* ［ゆらぎ］がオフならばらつきは使わない（欄の値は残す）/ With Jitter off the jitter values are kept but not used */
        var rotationJitter = fillSettings.jitter ? fillSettings.rotationJitter : 0;
        var scaleJitter = fillSettings.jitter ? fillSettings.scaleJitter : 0;
        return {
            placements: placements,
            /**
             * 置く位置の候補を作る
             * @param {number} x - タイルの中心の X 座標
             * @param {number} y - タイルの中心の Y 座標
             * @param {number} baseAngle - 並びの角度（度）
             * @param {boolean} isAlternate - 1つおきに回す側か
             * @returns {{x: number, y: number, angle: number, scale: number, tileIndex: number}} 候補
             */
            makeCandidate: function (x, y, baseAngle, isAlternate) {
                var orderRandom = nextRandom();
                var rotationRandom = nextRandom();
                var scaleRandom = nextRandom();
                var tileIndex = placements.length % tileCount;
                if (fillSettings.tileOrder === 'random') tileIndex = Math.min(tileCount - 1, Math.floor(orderRandom * tileCount));
                return {
                    x: x,
                    y: y,
                    tileIndex: tileIndex,
                    angle: baseAngle + getTileRotation(fillSettings) + ((fillSettings.alternateFlip && isAlternate) ? 180 : 0) +
                        rotationJitter * (rotationRandom * 2 - 1),
                    scale: fillSettings.scale * (1 - scaleJitter / 100 * scaleRandom)
                };
            },
            /**
             * 候補が容器に収まるか（マスクするときは形に触れるか）判定する
             * @param {Object} candidate - makeCandidate() の結果
             * @returns {boolean} 置けるなら true
             */
            isFitting: function (candidate) {
                return isPlacementFitting(candidate, fillTarget.tiles[candidate.tileIndex].info, containerShape, fillSettings.clipOverflow);
            },
            /**
             * 候補を置く位置として控える
             * @param {Object} candidate - makeCandidate() の結果
             * @returns {void}
             */
            accept: function (candidate) {
                placements.push(candidate);
            }
        };
    }

    /**
     * タイルが間隔込みで占める、回した長方形を作る（ランダムの重なり判定用）
     * @param {Object} placement - 置く位置
     * @param {Object} tileInfo - タイルの境界情報
     * @param {Object} fillSettings - showFillDialog() と同じ形の設定
     * @returns {{x: number, y: number, halfWidth: number, halfHeight: number, axisX: number[], axisY: number[]}} 中心・半分の幅と高さ・向きの単位ベクトル
     */
    function makeFootprint(placement, tileInfo, fillSettings) {
        var radians = placement.angle * Math.PI / 180;
        var cosValue = Math.cos(radians);
        var sinValue = Math.sin(radians);
        return {
            x: placement.x,
            y: placement.y,
            halfWidth: (tileInfo.width * placement.scale / 100 + fillSettings.gapX) / 2,
            halfHeight: (tileInfo.height * placement.scale / 100 + fillSettings.gapY) / 2,
            axisX: [cosValue, sinValue],
            axisY: [-sinValue, cosValue]
        };
    }

    /**
     * 回した長方形どうしが重なるか判定する（分離軸の定理。接するだけなら重ならない）
     * @param {Object} footprintA - makeFootprint() の結果
     * @param {Object} footprintB - makeFootprint() の結果
     * @returns {boolean} 重なれば true
     */
    function doFootprintsOverlap(footprintA, footprintB) {
        var dx = footprintB.x - footprintA.x;
        var dy = footprintB.y - footprintA.y;
        var separatingAxes = [footprintA.axisX, footprintA.axisY, footprintB.axisX, footprintB.axisY];
        for (var i = 0; i < separatingAxes.length; i++) {
            var axis = separatingAxes[i];
            var centerDistance = Math.abs(dx * axis[0] + dy * axis[1]);
            var reachA = footprintA.halfWidth * Math.abs(footprintA.axisX[0] * axis[0] + footprintA.axisX[1] * axis[1]) +
                footprintA.halfHeight * Math.abs(footprintA.axisY[0] * axis[0] + footprintA.axisY[1] * axis[1]);
            var reachB = footprintB.halfWidth * Math.abs(footprintB.axisX[0] * axis[0] + footprintB.axisX[1] * axis[1]) +
                footprintB.halfHeight * Math.abs(footprintB.axisY[0] * axis[0] + footprintB.axisY[1] * axis[1]);
            if (centerDistance >= reachA + reachB - EDGE_TOLERANCE) return false;
        }
        return true;
    }

    /**
     * 置く位置のタイルが、容器に収まるか（マスクするときは形に触れるか）判定する。
     * 回したタイルは、容器の形を逆に回してタイルを軸にそろえてから調べる
     * @param {Object} placement - 置く位置
     * @param {Object} tileInfo - タイルの境界情報
     * @param {Object} containerShape - buildContainerShape() の結果
     * @param {boolean} isTouchEnough - 形に触れていればよいなら true（マスクするとき）
     * @returns {boolean} 置けるなら true
     */
    function isPlacementFitting(placement, tileInfo, containerShape, isTouchEnough) {
        var halfWidth = tileInfo.width * placement.scale / 200;
        var halfHeight = tileInfo.height * placement.scale / 200;
        var checkAngle = normalizeHalfTurn(placement.angle);
        var checkShape = containerShape;
        var centerX = placement.x;
        var centerY = placement.y;
        if (Math.abs(checkAngle) > ANGLE_TOLERANCE) {
            checkShape = getRotatedContainerShape(containerShape, -checkAngle);
            var rotatedCenter = rotatePoint(centerX, centerY, -checkAngle);
            centerX = rotatedCenter[0];
            centerY = rotatedCenter[1];
        }
        var tileBounds = { left: centerX - halfWidth, right: centerX + halfWidth, top: centerY + halfHeight, bottom: centerY - halfHeight };
        return isTouchEnough ? isTileTouching(tileBounds, checkShape) : isTileInside(tileBounds, checkShape);
    }

    /* 角度ごとに回した容器の形の控え / Container shapes rotated per angle */
    var rotatedShapeCache = { sourceParts: null, byAngle: {}, size: 0 };

    /**
     * 容器の形を原点のまわりに回したものを返す（同じ角度は控えを使う）
     * @param {Object} containerShape - buildContainerShape() の結果
     * @param {number} angleDegrees - 回す角度（度）
     * @returns {{parts: Object[], offset: number}} 回した形
     */
    function getRotatedContainerShape(containerShape, angleDegrees) {
        if (rotatedShapeCache.sourceParts !== containerShape.parts || rotatedShapeCache.size >= ROTATED_SHAPE_CACHE_LIMIT) {
            rotatedShapeCache = { sourceParts: containerShape.parts, byAngle: {}, size: 0 };
        }
        var angleKey = angleDegrees.toFixed(6);
        var rotatedParts = rotatedShapeCache.byAngle[angleKey];
        if (!rotatedParts) {
            rotatedParts = [];
            for (var i = 0; i < containerShape.parts.length; i++) {
                var sourceEdges = containerShape.parts[i].edges;
                var rotatedEdges = [];
                for (var j = 0; j < sourceEdges.length; j++) {
                    var fromPoint = rotatePoint(sourceEdges[j].x1, sourceEdges[j].y1, angleDegrees);
                    var toPoint = rotatePoint(sourceEdges[j].x2, sourceEdges[j].y2, angleDegrees);
                    rotatedEdges.push(makeEdge(fromPoint, toPoint));
                }
                rotatedParts.push(makeShapePart(rotatedEdges));
            }
            rotatedShapeCache.byAngle[angleKey] = rotatedParts;
            rotatedShapeCache.size++;
        }
        return { parts: rotatedParts, offset: containerShape.offset };
    }

    /* 曲線の区間を折れ線に分ける数 / Line segments per curved segment when flattening */
    var CURVE_FLATTEN_STEPS = 16;

    /* 輪郭にちょうど接するタイルを残すための許容値（pt）/ Tolerance (pt) so tiles that just touch the outline are kept */
    var EDGE_TOLERANCE = 0.001;

    /**
     * パスを折れ線の頂点列に分ける（開いたパスは塗りと同じく始点と終点を結んで扱う）
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number[][]} 頂点の [x, y] の配列
     */
    function flattenPath(pathItem) {
        var pathPoints = pathItem.pathPoints;
        var pointCount = pathPoints.length;
        var vertices = [];
        for (var i = 0; i < pointCount; i++) {
            /* 始点・始点の方向線・終点の方向線・終点 / start, its handle, end handle, end */
            var startAnchor = pathPoints[i].anchor;
            var startHandle = pathPoints[i].rightDirection;
            var endHandle = pathPoints[(i + 1) % pointCount].leftDirection;
            var endAnchor = pathPoints[(i + 1) % pointCount].anchor;
            var isStraight = (startHandle[0] === startAnchor[0] && startHandle[1] === startAnchor[1] &&
                endHandle[0] === endAnchor[0] && endHandle[1] === endAnchor[1]);
            var stepCount = isStraight ? 1 : CURVE_FLATTEN_STEPS;
            for (var j = 0; j < stepCount; j++) {
                var t = j / stepCount;
                var u = 1 - t;
                var w0 = u * u * u, w1 = 3 * u * u * t, w2 = 3 * u * t * t, w3 = t * t * t;
                vertices.push([
                    w0 * startAnchor[0] + w1 * startHandle[0] + w2 * endHandle[0] + w3 * endAnchor[0],
                    w0 * startAnchor[1] + w1 * startHandle[1] + w2 * endHandle[1] + w3 * endAnchor[1]
                ]);
            }
        }
        return vertices;
    }

    /**
     * オブジェクトの輪郭を折れ線の頂点列にして配列に足す。
     * パスと複合パスは形そのもの、テキストは複製をアウトライン化した字形、
     * それ以外（配置画像など）は外接矩形を輪郭にする
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {number[][][]} outlines - 輪郭を足していく配列
     * @returns {number[][][]} outlines（輪郭ごとの頂点の [x, y] の配列）
     */
    function appendItemOutlines(pageItem, outlines) {
        if (pageItem.typename === 'PathItem') {
            outlines.push(flattenPath(pageItem));
        } else if (pageItem.typename === 'CompoundPathItem') {
            for (var i = 0; i < pageItem.pathItems.length; i++) outlines.push(flattenPath(pageItem.pathItems[i]));
        } else if (pageItem.typename === 'TextFrame') {
            var textOutlines = getTextOutlines(pageItem);
            for (var j = 0; j < textOutlines.length; j++) outlines.push(textOutlines[j]);
        } else {
            var itemInfo = getBoundsInfo(pageItem);
            outlines.push([
                [itemInfo.left, itemInfo.top],
                [itemInfo.right, itemInfo.top],
                [itemInfo.right, itemInfo.bottom],
                [itemInfo.left, itemInfo.bottom]
            ]);
        }
        return outlines;
    }

    /**
     * テキストの複製をアウトライン化し、字形の輪郭を返す（アウトライン化したものは削除する）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[][][]} 輪郭ごとの頂点の [x, y] の配列
     */
    function getTextOutlines(textFrame) {
        /* createOutline() は複製を消費し、アウトラインのグループだけが残る。失敗したら複製を片付けて形なしとする
           createOutline() consumes the duplicate; on failure remove the duplicate and treat it as having no shape */
        var textCopy = textFrame.duplicate();
        var outlineGroup;
        try {
            outlineGroup = textCopy.createOutline();
        } catch (e) {
            try { textCopy.remove(); } catch (removeError) {}
            return [];
        }
        var outlines = [];
        try {
            var glyphParts = collectContainerParts(outlineGroup, []);
            for (var i = 0; i < glyphParts.length; i++) appendItemOutlines(glyphParts[i], outlines);
        } finally {
            outlineGroup.remove();
        }
        return outlines;
    }

    /**
     * 容器を形の単位に分ける。グループは構成するオブジェクトごと（入れ子も展開）、
     * クリップグループはクリッピングパスだけを形にする。非表示のものは除く
     * @param {PageItem} container - 容器
     * @param {PageItem[]} parts - 見つけたオブジェクトを足していく配列
     * @returns {PageItem[]} parts（形として扱うオブジェクトの配列）
     */
    function collectContainerParts(container, parts) {
        if (container.typename !== 'GroupItem') {
            parts.push(container);
            return parts;
        }
        var groupMembers = container.pageItems;
        /* クリップグループはマスクだけ。マスクは先頭（テキストのマスクには clipping が無い）/ Only the mask; it is the first member (text masks have no clipping flag) */
        if (container.clipped && groupMembers.length) {
            parts.push(groupMembers[0]);
            return parts;
        }
        for (var i = 0; i < groupMembers.length; i++) {
            if (!groupMembers[i].hidden) collectContainerParts(groupMembers[i], parts);
        }
        return parts;
    }

    /**
     * 頂点列を判定用の辺に変換する
     * @param {number[][][]} outlines - 輪郭ごとの頂点の [x, y] の配列
     * @returns {Object[]} 辺（端点と外接矩形）の配列
     */
    function outlinesToEdges(outlines) {
        var edges = [];
        for (var k = 0; k < outlines.length; k++) {
            var vertices = outlines[k];
            if (vertices.length < 3) continue; /* 面を作らない線 / a line encloses nothing */
            for (var j = 0; j < vertices.length; j++) {
                edges.push(makeEdge(vertices[j], vertices[(j + 1) % vertices.length]));
            }
        }
        return edges;
    }

    /**
     * 2点を結ぶ判定用の辺を作る
     * @param {number[]} fromPoint - 始点の [x, y]
     * @param {number[]} toPoint - 終点の [x, y]
     * @returns {Object} 辺（端点と外接矩形）
     */
    function makeEdge(fromPoint, toPoint) {
        return {
            x1: fromPoint[0], y1: fromPoint[1], x2: toPoint[0], y2: toPoint[1],
            minX: Math.min(fromPoint[0], toPoint[0]), maxX: Math.max(fromPoint[0], toPoint[0]),
            minY: Math.min(fromPoint[1], toPoint[1]), maxY: Math.max(fromPoint[1], toPoint[1])
        };
    }

    /**
     * 容器の輪郭を、形ごとに折れ線の辺に分けて判定用にまとめる
     * @param {PageItem} container - 容器
     * @param {number} offsetPt - オフセット（pt。負で内側、正で外側）
     * @returns {{parts: Object[], offset: number}} 形ごとの辺と外接矩形（makeShapePart() の結果）とオフセット
     */
    function buildContainerShape(container, offsetPt) {
        /* 形はオフセットに関係しないので、同じ容器なら作り直さない（テキストのアウトライン化を繰り返さない）
           The outline does not depend on the offset, so reuse it for the same container (avoids re-outlining text) */
        if (containerPartsCache.container !== container) {
            var containerParts = collectContainerParts(container, []);
            var parts = [];
            var outlines = [];
            for (var i = 0; i < containerParts.length; i++) {
                var partOutlines = appendItemOutlines(containerParts[i], []);
                var edges = outlinesToEdges(partOutlines);
                if (edges.length) parts.push(makeShapePart(edges));
                for (var j = 0; j < partOutlines.length; j++) outlines.push(partOutlines[j]);
            }
            containerPartsCache = { container: container, parts: parts, outlines: outlines };
        }
        return { parts: containerPartsCache.parts, offset: offsetPt || 0 };
    }

    /**
     * 容器の輪郭を、折れ線の頂点列で返す（マスクを作るときに使う。buildContainerShape() と同じ控えを使う）
     * @param {PageItem} container - 容器
     * @returns {number[][][]} 輪郭ごとの頂点の [x, y] の配列
     */
    function getContainerOutlines(container) {
        buildContainerShape(container, 0);
        return containerPartsCache.outlines;
    }

    /**
     * 形の辺に外接矩形を添える（タイルと重ならない形を先に除くため）
     * @param {Object[]} edges - 辺（端点と外接矩形）の配列
     * @returns {{edges: Object[], minX: number, maxX: number, minY: number, maxY: number}} 辺と外接矩形
     */
    function makeShapePart(edges) {
        var shapePart = { edges: edges, minX: edges[0].minX, maxX: edges[0].maxX, minY: edges[0].minY, maxY: edges[0].maxY };
        for (var i = 1; i < edges.length; i++) {
            if (edges[i].minX < shapePart.minX) shapePart.minX = edges[i].minX;
            if (edges[i].maxX > shapePart.maxX) shapePart.maxX = edges[i].maxX;
            if (edges[i].minY < shapePart.minY) shapePart.minY = edges[i].minY;
            if (edges[i].maxY > shapePart.maxY) shapePart.maxY = edges[i].maxY;
        }
        return shapePart;
    }

    /* buildContainerShape() で作った形と輪郭の控え（実行ごとに main() で空にする）/ Shape and outline cache, cleared by main() on every run */
    var containerPartsCache = { container: null, parts: [], outlines: [] };

    /**
     * 点が輪郭の内側にあるか判定する（偶奇規則。複合パスの穴は外側になる）
     * @param {number} x - X 座標
     * @param {number} y - Y 座標
     * @param {Object[]} edges - buildContainerShape() の辺
     * @returns {boolean} 内側なら true
     */
    function isPointInsideOutline(x, y, edges) {
        var isInside = false;
        for (var i = 0; i < edges.length; i++) {
            var edge = edges[i];
            if ((edge.y1 > y) !== (edge.y2 > y) &&
                x < edge.x1 + (y - edge.y1) * (edge.x2 - edge.x1) / (edge.y2 - edge.y1)) {
                isInside = !isInside;
            }
        }
        return isInside;
    }

    /**
     * 輪郭の辺が矩形に触れるか判定する（Liang–Barsky の線分クリップ）
     * @param {Object} edge - 辺（端点と外接矩形）
     * @param {{left: number, top: number, right: number, bottom: number}} rect - 矩形
     * @returns {boolean} 触れていれば true
     */
    function isEdgeTouchingRect(edge, rect) {
        if (edge.maxX < rect.left || edge.minX > rect.right || edge.maxY < rect.bottom || edge.minY > rect.top) return false;
        var dx = edge.x2 - edge.x1;
        var dy = edge.y2 - edge.y1;
        var directions = [-dx, dx, -dy, dy];
        var distances = [edge.x1 - rect.left, rect.right - edge.x1, edge.y1 - rect.bottom, rect.top - edge.y1];
        var enterT = 0;
        var exitT = 1;
        for (var k = 0; k < 4; k++) {
            if (directions[k] === 0) {
                if (distances[k] < 0) return false;
                continue;
            }
            var t = distances[k] / directions[k];
            if (directions[k] < 0) {
                if (t > exitT) return false;
                if (t > enterT) enterT = t;
            } else {
                if (t < enterT) return false;
                if (t < exitT) exitT = t;
            }
        }
        return true;
    }

    /**
     * 点から線分までの距離を返す
     * @param {number} x - 点の X 座標
     * @param {number} y - 点の Y 座標
     * @param {Object} edge - 辺（端点と外接矩形）
     * @returns {number} 距離（pt）
     */
    function getPointToEdgeDistance(x, y, edge) {
        var dx = edge.x2 - edge.x1;
        var dy = edge.y2 - edge.y1;
        var lengthSquared = dx * dx + dy * dy;
        var t = (lengthSquared > 0) ? ((x - edge.x1) * dx + (y - edge.y1) * dy) / lengthSquared : 0;
        t = Math.max(0, Math.min(1, t));
        var nearestX = edge.x1 + t * dx;
        var nearestY = edge.y1 + t * dy;
        return Math.sqrt((x - nearestX) * (x - nearestX) + (y - nearestY) * (y - nearestY));
    }

    /**
     * 点から矩形までの距離を返す（内側なら 0）
     * @param {number} x - 点の X 座標
     * @param {number} y - 点の Y 座標
     * @param {{left: number, top: number, right: number, bottom: number}} rect - 矩形
     * @returns {number} 距離（pt）
     */
    function getPointToRectDistance(x, y, rect) {
        var dx = Math.max(rect.left - x, 0, x - rect.right);
        var dy = Math.max(rect.bottom - y, 0, y - rect.top);
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 輪郭の辺が矩形から指定の距離より近いか判定する
     * @param {Object} edge - 辺（端点と外接矩形）
     * @param {{left: number, top: number, right: number, bottom: number}} rect - 矩形
     * @param {number} clearance - 空けたい距離（pt）
     * @returns {boolean} 近ければ true
     */
    function isEdgeCloserThan(edge, rect, clearance) {
        if (edge.maxX < rect.left - clearance || edge.minX > rect.right + clearance ||
            edge.maxY < rect.bottom - clearance || edge.minY > rect.top + clearance) return false;
        if (isEdgeTouchingRect(edge, rect)) return true;
        /* 交わらない線分と矩形の距離は、端点か矩形の角のどれかで決まる / The gap is set by an endpoint or a rectangle corner */
        return (getPointToRectDistance(edge.x1, edge.y1, rect) < clearance ||
            getPointToRectDistance(edge.x2, edge.y2, rect) < clearance ||
            getPointToEdgeDistance(rect.left, rect.top, edge) < clearance ||
            getPointToEdgeDistance(rect.right, rect.top, edge) < clearance ||
            getPointToEdgeDistance(rect.left, rect.bottom, edge) < clearance ||
            getPointToEdgeDistance(rect.right, rect.bottom, edge) < clearance);
    }

    /**
     * タイルが、容器のどれか1つの形をオフセットした内側に収まっているか判定する
     * @param {Object} tileBounds - タイルの境界情報
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {boolean} 残すなら true
     */
    function isTileInside(tileBounds, containerShape) {
        for (var i = 0; i < containerShape.parts.length; i++) {
            var shapePart = containerShape.parts[i];
            /* 外接矩形が重ならない形は調べない / Skip shapes whose bounding box does not overlap the tile */
            if (tileBounds.right < shapePart.minX || tileBounds.left > shapePart.maxX ||
                tileBounds.top < shapePart.minY || tileBounds.bottom > shapePart.maxY) continue;
            if (isTileInsidePart(tileBounds, shapePart.edges, containerShape.offset)) return true;
        }
        return false;
    }

    /**
     * タイルが、1つの形をオフセットした内側に収まっているか判定する。
     * 負のオフセットはタイルの外接矩形と輪郭の距離で測り、正のオフセットは外接矩形をその分縮めて判定する
     * @param {Object} tileBounds - タイルの境界情報
     * @param {Object[]} edges - 形の辺（端点と外接矩形）
     * @param {number} offsetPt - オフセット（pt。負で内側、正で外側）
     * @returns {boolean} 収まっていれば true
     */
    function isTileInsidePart(tileBounds, edges, offsetPt) {
        /* はみ出しを許す分だけ縮める。潰れたときはタイルの中心で判定 / Shrink by the allowed overhang; a collapsed box becomes the tile center */
        var shrinkBy = Math.max(0, offsetPt) + EDGE_TOLERANCE;
        var rect = {
            left: tileBounds.left + shrinkBy,
            right: tileBounds.right - shrinkBy,
            top: tileBounds.top - shrinkBy,
            bottom: tileBounds.bottom + shrinkBy
        };
        if (rect.left > rect.right) rect.left = rect.right = (tileBounds.left + tileBounds.right) / 2.0;
        if (rect.bottom > rect.top) rect.top = rect.bottom = (tileBounds.top + tileBounds.bottom) / 2.0;

        if (!isPointInsideOutline(rect.left, rect.top, edges) ||
            !isPointInsideOutline(rect.right, rect.top, edges) ||
            !isPointInsideOutline(rect.left, rect.bottom, edges) ||
            !isPointInsideOutline(rect.right, rect.bottom, edges)) return false;

        /* 内側へのオフセットは輪郭から空ける距離になる / An inward offset is the clearance from the outline */
        var clearance = Math.max(0, -offsetPt - EDGE_TOLERANCE);
        for (var i = 0; i < edges.length; i++) {
            if (clearance > 0 ? isEdgeCloserThan(edges[i], rect, clearance) : isEdgeTouchingRect(edges[i], rect)) return false;
        }
        return true;
    }

    /**
     * タイルが、容器のどれか1つの形に少しでも重なるか判定する（マスクするときに残すタイル）。
     * 輪郭の辺がタイルに触れていれば重なり、触れていなければタイルの中心が形の内側にあるかで決まる
     * @param {Object} tileBounds - タイルの境界情報
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {boolean} 重なれば true
     */
    function isTileTouching(tileBounds, containerShape) {
        var centerX = (tileBounds.left + tileBounds.right) / 2;
        var centerY = (tileBounds.top + tileBounds.bottom) / 2;
        for (var i = 0; i < containerShape.parts.length; i++) {
            var shapePart = containerShape.parts[i];
            if (tileBounds.right < shapePart.minX || tileBounds.left > shapePart.maxX ||
                tileBounds.top < shapePart.minY || tileBounds.bottom > shapePart.maxY) continue;
            for (var j = 0; j < shapePart.edges.length; j++) {
                if (isEdgeTouchingRect(shapePart.edges[j], tileBounds)) return true;
            }
            if (isPointInsideOutline(centerX, centerY, shapePart.edges)) return true;
        }
        return false;
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 前回の設定の値を返す（無ければ代わりの値）
     * @param {Object|null} lastSettings - loadLastSettings() の結果
     * @param {string} settingName - 設定の名前
     * @param {*} fallbackValue - 前回の値が無いときの値
     * @returns {*} 値
     */
    function getSavedSetting(lastSettings, settingName, fallbackValue) {
        return (lastSettings && lastSettings[settingName] !== undefined) ? lastSettings[settingName] : fallbackValue;
    }

    /**
     * ラジオ・チェックボックスの状態を、前回の設定（無ければ初期値）から作る。
     * 表示前の value は読み戻せないので、状態はこのオブジェクトに控えて読む
     * @param {Object|null} lastSettings - loadLastSettings() の結果
     * @returns {Object} ダイアログの状態（layoutType / centerAlign / clipOverflow / tileRotationOn / alternateFlip / jitter / tileOrder / symbolize / preview / containerAction / randomSeed）
     */
    function createDialogState(lastSettings) {
        var layoutType = getSavedSetting(lastSettings, 'layoutType', null);
        /* v1.7.0 で控えた isBrick も読む / Also read isBrick kept by v1.7.0 */
        if (!layoutType && lastSettings) layoutType = lastSettings.isBrick ? 'brick' : 'grid';
        if (!/^(grid|brick|brickVertical|hexagon|random)$/.test(layoutType)) layoutType = DEFAULT_LAYOUT_TYPE;
        var containerAction = getSavedSetting(lastSettings, 'containerAction', DEFAULT_CONTAINER_ACTION);
        /* 以前の「何もしない」（keep）などは初期値に戻す / Fall back for choices that no longer exist, such as the old "keep" */
        if (!/^(opacity|remove|guide)$/.test(containerAction)) containerAction = DEFAULT_CONTAINER_ACTION;
        return {
            layoutType: layoutType,
            centerAlign: getSavedSetting(lastSettings, 'centerAlign', DEFAULT_CENTER_ALIGN),
            clipOverflow: getSavedSetting(lastSettings, 'clipOverflow', DEFAULT_CLIP_OVERFLOW),
            /* 切り替えの無い版で控えた角度があればオンで始める / Start on if an angle was kept without the switch */
            tileRotationOn: getSavedSetting(lastSettings, 'tileRotationOn', getSavedSetting(lastSettings, 'tileRotation', 0) !== 0),
            alternateFlip: getSavedSetting(lastSettings, 'alternateFlip', false),
            jitter: getSavedSetting(lastSettings, 'jitter', false),
            tileOrder: (getSavedSetting(lastSettings, 'tileOrder', 'sequence') === 'random') ? 'random' : 'sequence',
            symbolize: getSavedSetting(lastSettings, 'symbolize', DEFAULT_SYMBOLIZE),
            preview: getSavedSetting(lastSettings, 'preview', DEFAULT_PREVIEW),
            containerAction: containerAction,
            /* ランダムな並び・ばらつきはダイアログを開くたびに変え、開いている間は［再配置］まで固定する（プレビューと実行をそろえる）
               New random arrangement and jitter per dialog, kept until Reshuffle so the preview matches the result */
            randomSeed: createRandomSeed()
        };
    }

    /**
     * ランダムな並びの種を新しく作る
     * @returns {number} 種（1〜2147483646 の整数）
     */
    function createRandomSeed() {
        return Math.floor(Math.random() * 2147483646) + 1;
    }

    /**
     * ダイアログの状態に結びついたチェックボックスを足す（クリックで状態を書き換えてから onToggle を呼ぶ）
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - helpTip のラベルのパス
     * @param {Object} dialogState - createDialogState() の結果
     * @param {string} stateName - 書き換える状態の名前
     * @param {Function} [onToggle] - 切り替えたあとに呼ぶ関数
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addStateCheckbox(parent, labelPath, tooltipPath, dialogState, stateName, onToggle) {
        var stateCheckbox = parent.add('checkbox', undefined, getLabel(labelPath));
        stateCheckbox.helpTip = getLabel(tooltipPath);
        stateCheckbox.value = dialogState[stateName];
        stateCheckbox.onClick = function () {
            dialogState[stateName] = stateCheckbox.value;
            if (onToggle) onToggle();
        };
        return stateCheckbox;
    }

    /**
     * ダイアログの状態に結びついたラジオボタンを足す（選ぶと状態を書き換えてから onSelect を呼ぶ）。
     * 別のグループに分かれたラジオは自動で排他にならないので、radioGroup を渡すと手で切り替える
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - helpTip のラベルのパス
     * @param {Object} dialogState - createDialogState() の結果
     * @param {string} stateName - 書き換える状態の名前
     * @param {string} stateValue - このラジオを選んだときの値
     * @param {Function} [onSelect] - 選んだあとに呼ぶ関数
     * @param {RadioButton[]} [radioGroup] - 手で排他にするラジオの配列（このラジオも足す）
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addStateRadio(parent, labelPath, tooltipPath, dialogState, stateName, stateValue, onSelect, radioGroup) {
        var stateRadio = parent.add('radiobutton', undefined, getLabel(labelPath));
        stateRadio.helpTip = getLabel(tooltipPath);
        stateRadio.value = (dialogState[stateName] === stateValue);
        stateRadio.onClick = function () {
            dialogState[stateName] = stateValue;
            if (radioGroup) {
                for (var i = 0; i < radioGroup.length; i++) radioGroup[i].value = (radioGroup[i] === stateRadio);
            }
            if (onSelect) onSelect();
        };
        if (radioGroup) radioGroup.push(stateRadio);
        return stateRadio;
    }

    /**
     * パネルに、項目名と∧∨付きの数値欄の行を足す
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {number} labelWidth - 項目名の幅
     * @param {string} initialText - 欄の初期値（単位付き）
     * @param {string} tooltipPath - helpTip のラベルのパス
     * @param {Object} stepOptions - ∧∨の設定（min / max / unit）
     * @param {Function} onValueChange - 値が変わったときに呼ぶ関数
     * @returns {EditText} 追加した欄（行は .parent.parent）
     */
    function addPanelNumberRow(parentPanel, labelPath, labelWidth, initialText, tooltipPath, stepOptions, onValueChange) {
        var numberRow = addFieldRow(parentPanel, labelPath, labelWidth);
        stepOptions.onStep = onValueChange;
        var numberField = addNumberField(numberRow, initialText, tooltipPath, stepOptions);
        numberField.characters = LENGTH_FIELD_CHARACTERS;
        numberField.onChanging = onValueChange;
        return numberField;
    }

    /**
     * スライダーで並びの周期の ±半分を動かす範囲と換算を作る（スライダーは周期の %、欄は今の単位の長さ）
     * @param {Function} getPeriodInUnit - 今の周期（今の単位）を返す関数
     * @returns {{min: number, max: number, toSlider: Function, fromSlider: Function, fieldStep: number}} addSliderFieldRow() に渡す範囲
     */
    function makePeriodSliderRange(getPeriodInUnit) {
        return {
            min: -SHIFT_SLIDER_RANGE,
            max: SHIFT_SLIDER_RANGE,
            fieldStep: 0.1,
            toSlider: function (shiftValue) {
                var period = getPeriodInUnit();
                if (!(period > 0)) return 0;
                /* 周期の倍数のずれは同じ並びなので、中央からの ±半周期に折り返す / Shifts a whole period apart look the same, so wrap to ±half a period */
                var periodRatio = shiftValue / period;
                return (periodRatio - Math.round(periodRatio)) * 100;
            },
            fromSlider: function (sliderValue) {
                return sliderValue / 100 * getPeriodInUnit();
            }
        };
    }

    /**
     * 欄とスライダーの値を、別の欄とスライダーに写す
     * @param {Object} sourceControls - 写す元（addSliderFieldRow() の結果）
     * @param {Object} targetControls - 写す先（addSliderFieldRow() の結果）
     * @returns {void}
     */
    function copySliderFieldValue(sourceControls, targetControls) {
        targetControls.field.text = sourceControls.field.text;
        targetControls.syncSlider();
    }

    /**
     * 欄の数値を読む（数値でなければ代わりの値。範囲があれば収める）
     * @param {EditText} numberField - 入力欄
     * @param {number} fallbackValue - 数値でないときの値
     * @param {number} [minValue] - 下限
     * @param {number} [maxValue] - 上限
     * @returns {number} 値
     */
    function readFieldNumber(numberField, fallbackValue, minValue, maxValue) {
        var fieldValue = parseFloat(numberField.text);
        if (isNaN(fieldValue)) fieldValue = fallbackValue;
        if (minValue !== undefined) fieldValue = Math.max(minValue, fieldValue);
        if (maxValue !== undefined) fieldValue = Math.min(maxValue, fieldValue);
        return fieldValue;
    }

    /**
     * 行やグループの有効・無効を切り替える（変わったときだけ。中の∧∨も描き直す）
     * @param {Group} targetGroup - 対象の行・グループ
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setGroupEnabled(targetGroup, isEnabled) {
        if (targetGroup.enabled === isEnabled) return;
        targetGroup.enabled = isEnabled;
        redrawSteppersIn(targetGroup);
    }

    /**
     * 敷き詰め設定のダイアログを表示する
     * @param {Object} fillTarget - 容器・タイルの参照。間隔の初期値とスライダーの範囲、送りの計算に使う
     * @param {Function} onPreview - 入力が変わるたびに設定（pt 換算済み）を受け取り、置くタイルの数を返す関数
     * @returns {Object|null} 設定（readFillSettings() の形。スケールは%、長さは pt、角度は度）。キャンセル時は null
     */
    function showFillDialog(fillTarget, onPreview) {
        /* ダイアログのタイトルはラベル＋バージョン / Dialog title = label + version */
        var fillDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(fillDialog);

        var rulerUnit = getUnitInfo("rulerType");
        /* 長さの欄の中に付ける単位（「10 mm」の形）/ Unit shown inside the length fields */
        var lengthUnit = ' ' + rulerUnit.label;

        /* いちばん幅の広いタイル（今の単位）から、間隔の初期値とスライダーの範囲を決める
           The widest tile, in the current unit, sets the initial spacing and the slider ranges */
        var maxTileWidthInUnit = 0;
        for (var i = 0; i < fillTarget.tiles.length; i++) {
            maxTileWidthInUnit = Math.max(maxTileWidthInUnit, ptToUnit(fillTarget.tiles[i].info.width, rulerUnit));
        }

        /* 前回の設定があれば初期値にする（長さは今の単位に換算）/ Start from the last settings if any (lengths converted to the current unit) */
        var lastSettings = loadLastSettings();
        var dialogState = createDialogState(lastSettings);
        /* 欄がすべてそろうまでプレビューを描かない / No preview until every field exists */
        var isDialogReady = false;

        /**
         * 前回の数値を、単位付きの欄の文字列にする
         * @param {string} settingName - 設定の名前
         * @param {number} fallbackValue - 前回の値が無いときの値
         * @param {string} fieldUnit - 欄の単位（長さの欄なら lengthUnit。値は pt から換算する）
         * @returns {string} 欄に入れる文字列
         */
        function getSavedFieldText(settingName, fallbackValue, fieldUnit) {
            var savedValue = getSavedSetting(lastSettings, settingName, fallbackValue);
            if (fieldUnit === lengthUnit) savedValue = ptToUnit(savedValue, rulerUnit);
            return formatStepperNumber(savedValue) + fieldUnit;
        }

        /* 上段は1段組み：項目名・欄・スライダー（スライダーはどの行も同じ長さ。間隔の2行だけ、その右に連動アイコン）
           Upper area in one column of label, field and slider; every slider has the same length, and the spacing rows carry the link icon to their right */

        /* スケール（縦横比は固定）/ Scale (aspect ratio kept) */
        var scaleControls = addSliderFieldRow(fillDialog, 'fieldLabel.scale', getSavedFieldText('scale', 100, PERCENT_UNIT), 'tooltip.scale',
            { min: 1, unit: PERCENT_UNIT }, { min: SCALE_SLIDER_MIN, max: SCALE_SLIDER_MAX }, updatePeriodAndPreview);

        /* 間隔（横・縦）。v1.7.0 で控えた間隔（gap）は両方に使う。2行のスライダーの右に連動アイコン
           Horizontal and vertical spacing; the single gap kept by v1.7.0 serves both. The link icon sits to the right of both sliders */
        var defaultSpacingPt = getSavedSetting(lastSettings, 'gap', unitToPt(Math.round(maxTileWidthInUnit * DEFAULT_GAP_RATIO), rulerUnit));
        var spacingSliderRange = { min: 0, max: maxTileWidthInUnit * SPACING_SLIDER_MAX_RATIO };
        var spacingPairRow = fillDialog.add('group');
        spacingPairRow.orientation = 'row';
        spacingPairRow.alignment = ['fill', 'top'];
        spacingPairRow.alignChildren = ['left', 'center'];
        spacingPairRow.spacing = GAP_LINK_SPACING;
        var spacingFieldColumn = spacingPairRow.add('group');
        spacingFieldColumn.orientation = 'column';
        spacingFieldColumn.alignment = ['left', 'center'];
        spacingFieldColumn.alignChildren = ['fill', 'top'];
        spacingFieldColumn.spacing = WINDOW_SPACING;
        var spacingXControls = addSliderFieldRow(spacingFieldColumn, 'fieldLabel.spacingX', getSavedFieldText('gapX', defaultSpacingPt, lengthUnit), 'tooltip.spacingX',
            { min: 0, unit: lengthUnit }, spacingSliderRange, function () { onSpacingChange(spacingXControls); });
        var spacingYControls = addSliderFieldRow(spacingFieldColumn, 'fieldLabel.spacingY', getSavedFieldText('gapY', defaultSpacingPt, lengthUnit), 'tooltip.spacingY',
            { min: 0, unit: lengthUnit }, spacingSliderRange, function () { onSpacingChange(spacingYControls); });
        var spacingLinkToggle = addLinkToggle(spacingPairRow, getSavedSetting(lastSettings, 'gapLinked', true), function () { onSpacingChange(spacingXControls); });
        spacingLinkToggle.helpTip = getLabel('tooltip.linkSpacing');
        if (spacingLinkToggle.value) copySliderFieldValue(spacingXControls, spacingYControls);
        spacingXControls.field.active = true;

        /* オフセット（負で内側、正で外側）/ Offset (negative inward, positive outward) */
        var offsetSliderRange = maxTileWidthInUnit * OFFSET_SLIDER_RANGE_RATIO;
        var offsetControls = addSliderFieldRow(fillDialog, 'fieldLabel.offset', getSavedFieldText('offset', 0, lengthUnit), 'tooltip.offset',
            { unit: lengthUnit }, { min: -offsetSliderRange, max: offsetSliderRange }, updatePreviewFromFields);

        /* 調整X・Y（欄は長さ、スライダーは端で並びの周期の半分ずれる）/ Adjust X and Y: the field holds a length; each slider end shifts by half the pattern period */
        var shiftXControls = addSliderFieldRow(fillDialog, 'fieldLabel.shiftX', getSavedFieldText('shiftX', 0, lengthUnit), 'tooltip.shiftX',
            { unit: lengthUnit }, makePeriodSliderRange(function () { return getShiftPeriodInUnit(false); }), updatePreviewFromFields);
        var shiftYControls = addSliderFieldRow(fillDialog, 'fieldLabel.shiftY', getSavedFieldText('shiftY', 0, lengthUnit), 'tooltip.shiftY',
            { unit: lengthUnit }, makePeriodSliderRange(function () { return getShiftPeriodInUnit(true); }), updatePreviewFromFields);

        /* 並びの角度（容器の中心を軸に、並び全体とタイルを回す）/ Layout angle: turns the whole arrangement and the tiles around the container's center */
        var layoutAngleControls = addSliderFieldRow(fillDialog, 'fieldLabel.layoutAngle', getSavedFieldText('layoutAngle', 0, DEGREE_UNIT), 'tooltip.layoutAngle',
            { unit: DEGREE_UNIT }, { min: -LAYOUT_ANGLE_SLIDER_RANGE, max: LAYOUT_ANGLE_SLIDER_RANGE }, updatePreviewFromFields);

        /* 置くタイルの数 / Number of tiles to place */
        var tileCountRow = addFieldRow(fillDialog, 'fieldLabel.count');
        tileCountRow.margins = [0, 0, 0, COUNT_ROW_BOTTOM_MARGIN];
        var tileCountText = tileCountRow.add('statictext', undefined, '');
        tileCountText.preferredSize.width = COUNT_TEXT_WIDTH;
        tileCountText.helpTip = getLabel('tooltip.count');

        /* 下段：左＝並べ方とシンボル化、右＝タイルとその下に敷き詰め先（左右の高さをそろえる）
           Lower row: layout and Symbolize on the left; tiles with the fill area below on the right, both sides of equal height */
        var panelColumns = addColumnRow(fillDialog, 2);
        panelColumns.row.alignChildren = ['fill', 'fill'];

        /* 並べ方（ラジオは同じパネルに続けて入れるので排他は自動）/ Layout; the radios follow one another in one panel, so exclusivity is automatic */
        var layoutPanel = panelColumns.columns[0].add('panel', undefined, getLabel('panel.layout'));
        setupPanel(layoutPanel, 6);
        addStateRadio(layoutPanel, 'radio.grid', 'tooltip.grid', dialogState, 'layoutType', 'grid', onLayoutChange);
        addStateRadio(layoutPanel, 'radio.brick', 'tooltip.brick', dialogState, 'layoutType', 'brick', onLayoutChange);
        addStateRadio(layoutPanel, 'radio.brickVertical', 'tooltip.brickVertical', dialogState, 'layoutType', 'brickVertical', onLayoutChange);
        addStateRadio(layoutPanel, 'radio.hexagon', 'tooltip.hexagon', dialogState, 'layoutType', 'hexagon', onLayoutChange);
        addStateRadio(layoutPanel, 'radio.random', 'tooltip.random', dialogState, 'layoutType', 'random', onLayoutChange);

        /* 密度（ランダムだけ）は［ランダム］の下に、ラジオの文字にそろえて置く / Density, for Random only, under the Random radio and lined up with its text */
        var densityRow = layoutPanel.add('group');
        setupRow(densityRow, ['left', 'center'], 0);
        var densityIndent = densityRow.add('group');
        densityIndent.preferredSize.width = RADIO_TEXT_INDENT;
        var densityLabel = densityRow.add('statictext', undefined, labelText('fieldLabel.density'));
        densityLabel.helpTip = getLabel('tooltip.density');
        var densityField = addNumberField(densityRow, getSavedFieldText('density', 100, PERCENT_UNIT), 'tooltip.density',
            { min: 1, max: 100, unit: PERCENT_UNIT, onStep: updatePreviewFromFields });
        densityField.characters = LENGTH_FIELD_CHARACTERS;
        densityField.onChanging = updatePreviewFromFields;

        var centerAlignCheckbox = addStateCheckbox(layoutPanel, 'checkbox.centerAlign', 'tooltip.centerAlign', dialogState, 'centerAlign', updatePreviewFromFields);
        addStateCheckbox(layoutPanel, 'checkbox.clipOverflow', 'tooltip.clipOverflow', dialogState, 'clipOverflow', updateEnabledStatesAndPreview);

        /* 再配置（ランダムな並び・ばらつき・選び方を引き直す）/ Reshuffle: draw a new random arrangement, jitter and order */
        var reshuffleButton = layoutPanel.add('button', undefined, getLabel('button.reshuffle'));
        reshuffleButton.alignment = 'left';
        reshuffleButton.helpTip = getLabel('tooltip.reshuffle');
        reshuffleButton.onClick = function () {
            dialogState.randomSeed = createRandomSeed();
            updatePreviewFromFields();
        };

        /* シンボル化は［並べ方］パネルの下 / Symbolize sits under the Layout panel */
        addStateCheckbox(panelColumns.columns[0], 'checkbox.symbolize', 'tooltip.symbolize', dialogState, 'symbolize');

        /* タイル（個別に回転・1つおきに180°回転・ゆらぎ・選び方）/ Tiles: Rotate Each, Flip Every Other, Jitter and Order */
        var tilePanel = panelColumns.columns[1].add('panel', undefined, getLabel('panel.tile'));
        setupPanel(tilePanel, 6);

        /* 個別に回転：項目名の代わりにチェックボックスを置き、オフなら欄をディム表示にする。回転は送り（外接矩形）を変えるので周期も写し直す
           Rotate Each: a checkbox in place of the label dims the field when off. Rotation changes the step, so the period is refreshed too */
        var rotationRow = tilePanel.add('group');
        setupRow(rotationRow, ['fill', 'center']);
        var rotationCheckbox = addStateCheckbox(rotationRow, 'fieldLabel.rotation', 'tooltip.rotation', dialogState, 'tileRotationOn', function () {
            updateEnabledStates();
            updatePeriodAndPreview();
        });
        rotationCheckbox.preferredSize.width = TILE_LABEL_WIDTH;
        var rotationField = addNumberField(rotationRow, getSavedFieldText('tileRotation', 0, DEGREE_UNIT), 'tooltip.rotation',
            { unit: DEGREE_UNIT, onStep: updatePeriodAndPreview });
        rotationField.characters = LENGTH_FIELD_CHARACTERS;
        rotationField.onChanging = updatePeriodAndPreview;

        /* 1つおきに180°（項目名の幅だけ空けて欄の列にそろえる）/ Flip every other tile, lined up with the field column */
        var alternateFlipRow = tilePanel.add('group');
        setupRow(alternateFlipRow, ['left', 'center']);
        var alternateFlipIndent = alternateFlipRow.add('group');
        alternateFlipIndent.preferredSize.width = TILE_LABEL_WIDTH;
        addStateCheckbox(alternateFlipRow, 'checkbox.alternateFlip', 'tooltip.alternateFlip', dialogState, 'alternateFlip', updatePreviewFromFields);

        /* ゆらぎ（回転・大きさのばらつきをまとめてオン・オフ）/ Jitter: turns rotation and size jitter on and off together */
        addStateCheckbox(tilePanel, 'checkbox.jitter', 'tooltip.jitter', dialogState, 'jitter', updateEnabledStatesAndPreview);
        var rotationJitterField = addPanelNumberRow(tilePanel, 'fieldLabel.rotationJitter', TILE_LABEL_WIDTH, getSavedFieldText('rotationJitter', 0, DEGREE_UNIT),
            'tooltip.rotationJitter', { min: 0, max: 180, unit: DEGREE_UNIT }, updatePreviewFromFields);
        var scaleJitterField = addPanelNumberRow(tilePanel, 'fieldLabel.scaleJitter', TILE_LABEL_WIDTH, getSavedFieldText('scaleJitter', 0, PERCENT_UNIT),
            'tooltip.scaleJitter', { min: 0, max: SCALE_JITTER_MAX, unit: PERCENT_UNIT }, updatePreviewFromFields);

        /* 選び方（敷き詰めるオブジェクトが複数のときだけ）。2つのラジオは同じ行に入れるので排他は自動
           Order, only with several objects to tile; both radios share the row, so exclusivity is automatic */
        var tileOrderRow = addFieldRow(tilePanel, 'fieldLabel.tileOrder', TILE_LABEL_WIDTH);
        addStateRadio(tileOrderRow, 'radio.orderSequence', 'tooltip.tileOrder', dialogState, 'tileOrder', 'sequence', updatePreviewFromFields);
        addStateRadio(tileOrderRow, 'radio.orderRandom', 'tooltip.tileOrder', dialogState, 'tileOrder', 'random', updatePreviewFromFields);
        tileOrderRow.enabled = (fillTarget.tiles.length > 1);

        /* 敷き詰め先（敷き詰めたあとの扱い）。［不透明度］は値の欄と同じ行に入るので、ラジオの排他は手で切り替える
           Fill area, after tiling. The Opacity radio shares a row with its field, so exclusivity is handled by hand */
        var containerActionPanel = panelColumns.columns[1].add('panel', undefined, getLabel('panel.containerAction'));
        setupPanel(containerActionPanel, 6);
        var containerActionRadios = [];
        addStateRadio(containerActionPanel, 'radio.removeContainer', 'tooltip.removeContainer', dialogState, 'containerAction', 'remove', updateEnabledStates, containerActionRadios);
        addStateRadio(containerActionPanel, 'radio.guideContainer', 'tooltip.guideContainer', dialogState, 'containerAction', 'guide', updateEnabledStates, containerActionRadios);
        /* 不透明度は下げるだけ。100% なら何もしない / Opacity only lowers; at 100% nothing changes */
        var opacityRow = containerActionPanel.add('group');
        setupRow(opacityRow, ['left', 'center']);
        addStateRadio(opacityRow, 'radio.opacityContainer', 'tooltip.opacityContainer', dialogState, 'containerAction', 'opacity', updateEnabledStates, containerActionRadios);
        var opacityField = addNumberField(opacityRow, getSavedFieldText('containerOpacity', 100, PERCENT_UNIT), 'tooltip.opacityContainer',
            { min: 0, max: 100, unit: PERCENT_UNIT });
        opacityField.characters = LENGTH_FIELD_CHARACTERS;

        /* ボタン（左にプレビュー）/ Buttons, with Preview on the left */
        var buttonRow = addButtonRow(fillDialog);
        addStateCheckbox(buttonRow.leftGroup, 'checkbox.preview', 'tooltip.preview', dialogState, 'preview', updatePreviewFromFields);
        buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        /**
         * 今の入力内容での並びの周期を、今の単位の長さで返す（送り。レンガ状・六角形でずらす向きは2個分）
         * @param {boolean} isVertical - 縦の周期なら true
         * @returns {number} 周期（欄がそろう前や決められないときは 0）
         */
        function getShiftPeriodInUnit(isVertical) {
            if (!isDialogReady) return 0;
            var layoutMetrics = getLayoutMetrics(fillTarget, readFillSettings());
            if (!layoutMetrics) return 0;
            var brickDirection = BRICK_DIRECTIONS[dialogState.layoutType] || '';
            var stepPt = isVertical ? layoutMetrics.stepY : layoutMetrics.stepX;
            var isDoubled = isVertical ? (brickDirection === 'rows') : (brickDirection === 'columns');
            return ptToUnit(isDoubled ? stepPt * 2 : stepPt, rulerUnit);
        }

        /**
         * 間隔が変わったら、連動中はもう一方へ写し、周期とプレビューを更新する
         * @param {Object} sourceControls - 変わった側（addSliderFieldRow() の結果）
         * @returns {void}
         */
        function onSpacingChange(sourceControls) {
            if (spacingLinkToggle && spacingLinkToggle.value) {
                copySliderFieldValue(sourceControls, (sourceControls === spacingXControls) ? spacingYControls : spacingXControls);
            }
            updatePeriodAndPreview();
        }

        /**
         * 並べ方を選んだら、ディム表示と周期・プレビューを更新する
         * @returns {void}
         */
        function onLayoutChange() {
            updateEnabledStates();
            updatePeriodAndPreview();
        }

        /**
         * ディム表示を更新してからプレビューを描き直す
         * @returns {void}
         */
        function updateEnabledStatesAndPreview() {
            updateEnabledStates();
            updatePreviewFromFields();
        }

        /**
         * 周期が変わる入力（スケール・間隔・並べ方・回転）のあとに、調整X・Y のスライダーを写し直してプレビューを描き直す
         * @returns {void}
         */
        function updatePeriodAndPreview() {
            if (!isDialogReady) return;
            shiftXControls.syncSlider();
            shiftYControls.syncSlider();
            updatePreviewFromFields();
        }

        /**
         * 入力欄の値を pt に換算して読む
         * @returns {Object} 設定（スケール・密度・大きさのばらつき・不透明度は%、長さは pt、角度は度）
         */
        function readFillSettings() {
            var scaleValue = readFieldNumber(scaleControls.field, 100);
            return {
                scale: (scaleValue > 0) ? scaleValue : 100,
                gapX: unitToPt(readFieldNumber(spacingXControls.field, 0, 0), rulerUnit),
                gapY: unitToPt(readFieldNumber(spacingYControls.field, 0, 0), rulerUnit),
                gapLinked: spacingLinkToggle.value,
                offset: unitToPt(readFieldNumber(offsetControls.field, 0), rulerUnit),
                shiftX: unitToPt(readFieldNumber(shiftXControls.field, 0), rulerUnit),
                shiftY: unitToPt(readFieldNumber(shiftYControls.field, 0), rulerUnit),
                layoutAngle: readFieldNumber(layoutAngleControls.field, 0),
                layoutType: dialogState.layoutType,
                centerAlign: dialogState.centerAlign,
                clipOverflow: dialogState.clipOverflow,
                density: readFieldNumber(densityField, 100, 1, 100),
                randomSeed: dialogState.randomSeed,
                tileRotationOn: dialogState.tileRotationOn,
                tileRotation: readFieldNumber(rotationField, 0),
                alternateFlip: dialogState.alternateFlip,
                jitter: dialogState.jitter,
                rotationJitter: readFieldNumber(rotationJitterField, 0, 0, 180),
                scaleJitter: readFieldNumber(scaleJitterField, 0, 0, SCALE_JITTER_MAX),
                tileOrder: dialogState.tileOrder,
                symbolize: dialogState.symbolize,
                preview: dialogState.preview,
                containerAction: dialogState.containerAction,
                containerOpacity: readFieldNumber(opacityField, 100, 0, 100)
            };
        }

        /**
         * 現在の入力内容でプレビューを描き直し、置くタイルの数を表示する
         * @returns {void}
         */
        function updatePreviewFromFields() {
            if (!isDialogReady) return;
            tileCountText.text = String(onPreview(readFillSettings()));
        }

        /**
         * 選択に合わせて、使わない項目をディム表示にする
         * （ランダムは調整X・Y・並びの角度・中央そろえを使わず、密度はランダムだけ。マスクはオフセットを使わない。
         * ばらつきは［ゆらぎ］、個別に回転の角度は［個別に回転］、不透明度の値は［不透明度］がオンのときだけ）
         * @returns {void}
         */
        function updateEnabledStates() {
            var isRandom = (dialogState.layoutType === 'random');
            setGroupEnabled(shiftXControls.row, !isRandom);
            setGroupEnabled(shiftYControls.row, !isRandom);
            setGroupEnabled(layoutAngleControls.row, !isRandom);
            setGroupEnabled(densityRow, isRandom);
            centerAlignCheckbox.enabled = !isRandom;
            setGroupEnabled(offsetControls.row, !dialogState.clipOverflow);
            setGroupEnabled(opacityField.parent, dialogState.containerAction === 'opacity');
            /* 欄の行は「∧∨と欄のグループ」の親 / A field's row is the parent of its stepper group */
            setGroupEnabled(rotationJitterField.parent.parent, dialogState.jitter);
            setGroupEnabled(scaleJitterField.parent.parent, dialogState.jitter);
            setGroupEnabled(rotationField.parent, dialogState.tileRotationOn);
        }

        /* 欄がすべてそろってから、ディム表示・調整X・Y のスライダー・最初のプレビューを整える
           Once every field exists, set the dimming, the Adjust X/Y sliders and the first preview */
        isDialogReady = true;
        updateEnabledStates();
        updatePeriodAndPreview();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(fillDialog, SCRIPT_NAME);
        if (fillDialog.show() !== 1) return null;

        /* OKで確定値をptに変換し、次回のために控える / Convert to pt on OK and keep it for the next run */
        var confirmedSettings = readFillSettings();
        saveLastSettings(confirmedSettings);
        return confirmedSettings;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したオブジェクトを容器（いちばん大きいもの）とタイル（残り）に振り分け、ダイアログで設定して敷き詰める
     * @returns {void}
     */
    function main() {
        containerPartsCache = { container: null, parts: [], outlines: [] };
        var doc = app.documents.length ? app.activeDocument : null;
        if (!doc) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        if (!doc.selection || doc.selection.length < 2) {
            alert(getLabel('alert.selectTwo'));
            return;
        }

        /* 面積がいちばん大きいものを容器、残りをタイルにする（タイルは選択の順＝前面から）
           The largest by area is the container; the rest are tiles, in selection order (front first) */
        var selectedItems = [];
        var containerIndex = 0;
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push({ item: doc.selection[i], info: getBoundsInfo(doc.selection[i]) });
            if (selectedItems[i].info.area > selectedItems[containerIndex].info.area) containerIndex = i;
        }
        var tiles = [];
        for (var j = 0; j < selectedItems.length; j++) {
            if (j === containerIndex) continue;
            if (selectedItems[j].info.width <= 0 || selectedItems[j].info.height <= 0) {
                alert(getLabel('alert.emptyTile'));
                return;
            }
            tiles.push(selectedItems[j]);
        }
        var fillTarget = {
            container: selectedItems[containerIndex].item,
            containerInfo: selectedItems[containerIndex].info,
            tiles: tiles,
            previewGroup: null
        };

        var fillSettings = showFillDialog(fillTarget, function(previewSettings) {
            var placements = computePlacements(fillTarget, previewSettings);
            if (previewSettings.preview) {
                renderPreview(fillTarget, previewSettings, placements);
            } else {
                clearPreview(fillTarget);
                app.redraw();
            }
            return placements.length;
        });
        clearPreview(fillTarget);
        if (!fillSettings) {
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            return;
        }

        placeTiles(doc, fillTarget, fillSettings);
        applyContainerAction(doc, fillTarget.container, fillSettings);
    }

    /**
     * 敷き詰めたあと、大きいオブジェクトの不透明度を変える・削除する・ガイド化する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} container - 大きいオブジェクト
     * @param {Object} fillSettings - showFillDialog() の結果（containerAction は "opacity" / "remove" / "guide"）
     * @returns {void}
     */
    function applyContainerAction(doc, container, fillSettings) {
        var containerAction = fillSettings.containerAction;
        if (containerAction === 'opacity') {
            if (fillSettings.containerOpacity < 100) container.opacity = fillSettings.containerOpacity;
        } else if (containerAction === 'remove') {
            container.remove();
        } else if (containerAction === 'guide') {
            /* メニューコマンド makeGuide は yeKB で失敗するので DOM で変換する / makeGuide fails with yeKB, so convert through the DOM */
            convertToGuides(container);
            doc.selection = null;
        }
    }

    /**
     * パスをガイドに変換する。グループは中まで（クリップグループはマスクを解除して）、
     * 複合パスは構成するパスを外に出してから、テキストはアウトライン化してから変換する。
     * 画像・シンボルなどガイドにできないものはそのまま残す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {void}
     */
    function convertToGuides(pageItem) {
        if (pageItem.typename === 'PathItem') {
            if (pageItem.clipping) pageItem.clipping = false; /* クリッピングパスのままではガイドにできない / a clipping path cannot become a guide */
            pageItem.guides = true;
        } else if (pageItem.typename === 'CompoundPathItem') {
            for (var i = pageItem.pathItems.length - 1; i >= 0; i--) {
                var memberPath = pageItem.pathItems[i];
                memberPath.move(pageItem, ElementPlacement.PLACEBEFORE);
                convertToGuides(memberPath);
            }
            pageItem.remove();
        } else if (pageItem.typename === 'TextFrame') {
            /* テキストはアウトライン化してからガイドにする / Outline text first */
            convertToGuides(pageItem.createOutline());
        } else if (pageItem.typename === 'GroupItem') {
            if (pageItem.clipped) pageItem.clipped = false;
            for (var j = pageItem.pageItems.length - 1; j >= 0; j--) convertToGuides(pageItem.pageItems[j]);
        }
    }

    /**
     * プレビューを削除する（参照で消し、さらに名前で残りを掃除する）
     * @param {Object} fillTarget - 容器・タイル・プレビューの参照
     * @returns {void}
     */
    function clearPreview(fillTarget) {
        /* 削除済みのグループは参照が無効になる / The group may already be gone */
        try {
            if (fillTarget.previewGroup && fillTarget.previewGroup.isValid) {
                fillTarget.previewGroup.remove();
            }
        } catch (e) {}
        fillTarget.previewGroup = null;

        /* 名前で余分なプレビューを掃除。ロックされたレイヤーなどで削除できないものは飛ばす
           Sweep stray previews by name; skip any that cannot be removed, e.g. on a locked layer */
        var containerLayer = fillTarget.container.layer;
        for (var i = containerLayer.groupItems.length - 1; i >= 0; i--) {
            var groupItem = containerLayer.groupItems[i];
            if (groupItem.name.indexOf(PREVIEW_GROUP_NAME) !== 0) continue;
            try {
                groupItem.remove();
            } catch (e) {}
        }
    }

    /**
     * プレビューを描画する（半透明のグループに、決めた位置へタイルを複製する。マスクするときはマスクも作る）
     * @param {Object} fillTarget - 容器・タイル・プレビューの参照
     * @param {Object} previewSettings - showFillDialog() と同じ形の設定
     * @param {Object[]} placements - computePlacements() の結果
     * @returns {void}
     */
    function renderPreview(fillTarget, previewSettings, placements) {
        clearPreview(fillTarget);

        var previewGroup = fillTarget.container.layer.groupItems.add();
        previewGroup.name = PREVIEW_GROUP_NAME;
        fillTarget.previewGroup = previewGroup;

        for (var i = 0; i < placements.length; i++) {
            var tileEntry = fillTarget.tiles[placements[i].tileIndex];
            placeTileCopy(tileEntry.item.duplicate(previewGroup, ElementPlacement.PLACEATBEGINNING), placements[i], tileEntry.info);
        }
        if (previewSettings.clipOverflow && placements.length) addClipPath(fillTarget.container, previewGroup);

        previewGroup.opacity = PREVIEW_OPACITY;
        app.redraw();
    }

    /**
     * タイルの複製を縦横同じ比率で拡大・縮小し（線幅・パターン・グラデーションも）、中心を置く位置に合わせて回す
     * @param {PageItem} tileCopy - 複製したタイル（シンボルのインスタンスも可）
     * @param {Object} placement - 置く位置（computePlacements() の要素）
     * @param {Object|null} sourceInfo - 複製元の境界情報。拡大・縮小しても中心は変わらないので測らずに使う。null なら置いてから測る
     * @returns {void}
     */
    function placeTileCopy(tileCopy, placement, sourceInfo) {
        if (Math.abs(placement.scale - 100) > ANGLE_TOLERANCE) {
            tileCopy.resize(placement.scale, placement.scale, true, true, true, true, placement.scale);
        }
        var currentInfo = sourceInfo || getBoundsInfo(tileCopy);
        tileCopy.translate(placement.x - (currentInfo.left + currentInfo.right) / 2, placement.y - (currentInfo.top + currentInfo.bottom) / 2);
        var turnAngle = normalizeFullTurn(placement.angle);
        if (Math.abs(turnAngle) > ANGLE_TOLERANCE) {
            tileCopy.rotate(turnAngle, true, true, true, true, Transformation.CENTER);
        }
    }

    /**
     * 置く位置が、元のタイルそのものと同じか判定する（等倍・回転なしで中心が重なる）
     * @param {Object} placement - 置く位置
     * @param {Object} tileInfo - 元のタイルの境界情報
     * @returns {boolean} 同じなら true
     */
    function isOriginalTileSpot(placement, tileInfo) {
        return Math.abs(placement.scale - 100) < ANGLE_TOLERANCE &&
            Math.abs(normalizeFullTurn(placement.angle)) < ANGLE_TOLERANCE &&
            Math.abs(placement.x - (tileInfo.left + tileInfo.right) / 2) < 0.01 &&
            Math.abs(placement.y - (tileInfo.top + tileInfo.bottom) / 2) < 0.01;
    }

    /**
     * グループの最前面に容器の形のマスクを置き、クリップグループにする。
     * パス・複合パスは複製をそのまま使い、グループやテキストは輪郭を折れ線にした複合パスで作る
     * @param {PageItem} container - 容器
     * @param {GroupItem} targetGroup - マスクするグループ
     * @returns {void}
     */
    function addClipPath(container, targetGroup) {
        if (container.typename === 'PathItem' || container.typename === 'CompoundPathItem') {
            container.duplicate(targetGroup, ElementPlacement.PLACEATBEGINNING);
        } else {
            var clipCompound = targetGroup.compoundPathItems.add();
            var containerOutlines = getContainerOutlines(container);
            for (var i = 0; i < containerOutlines.length; i++) {
                if (containerOutlines[i].length < 3) continue;
                var outlinePath = clipCompound.pathItems.add();
                outlinePath.setEntirePath(containerOutlines[i]);
                outlinePath.closed = true;
            }
            clipCompound.move(targetGroup, ElementPlacement.PLACEATBEGINNING);
        }
        targetGroup.clipped = true;
    }

    /**
     * タイルをシンボルとして登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} tile - タイル
     * @returns {Symbol|null} 登録したシンボル。失敗時は null（通常の複製に切り替える）
     */
    function createTileSymbol(doc, tile) {
        /* シンボルにできないオブジェクトがある / Some items cannot become symbols */
        try {
            var symbolSource = tile.duplicate();
            var symbolDefinition = doc.symbols.add(symbolSource);
            try { symbolSource.remove(); } catch (e) {}
            return symbolDefinition;
        } catch (e) {
            return null; /* 失敗時は通常複製 / fallback */
        }
    }

    /**
     * 確定した設定で、決めた位置にタイルを敷き詰める（マスクするときはクリップグループにする）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() の結果
     * @returns {void}
     */
    function placeTiles(doc, fillTarget, fillSettings) {
        var placements = computePlacements(fillTarget, fillSettings);

        /* 必要ならタイルごとにシンボルを作る / Make a symbol per tile if needed */
        var symbolDefinitions = [];
        for (var i = 0; i < fillTarget.tiles.length; i++) {
            symbolDefinitions.push(fillSettings.symbolize ? createTileSymbol(doc, fillTarget.tiles[i].item) : null);
        }

        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        app.redraw();

        /* 複製グループを作成 / Create group for final duplicates */
        var gridGroup = fillTarget.container.layer.groupItems.add();
        gridGroup.name = GRID_GROUP_NAME;

        /**
         * タイルを1つ複製してグリッドのグループに入れる（シンボル化するならシンボルのインスタンス）
         * @param {number} tileIndex - タイルの番号
         * @returns {PageItem} 複製したタイル
         */
        function addTileCopy(tileIndex) {
            var symbolDefinition = symbolDefinitions[tileIndex];
            if (!symbolDefinition) return fillTarget.tiles[tileIndex].item.duplicate(gridGroup, ElementPlacement.PLACEATBEGINNING);
            var symbolCopy = doc.symbolItems.add(symbolDefinition);
            symbolCopy.move(gridGroup, ElementPlacement.PLACEATBEGINNING);
            return symbolCopy;
        }

        var placedCount = 0;
        for (var j = 0; j < placements.length; j++) {
            var tileEntry = fillTarget.tiles[placements[j].tileIndex];
            /* 元のタイルと同じ場所はスキップ / Skip the original tile's own spot */
            if (isOriginalTileSpot(placements[j], tileEntry.info)) continue;
            /* シンボルのインスタンスは置いた位置が決まっていないので測る / Symbol instances land anywhere, so measure them */
            placeTileCopy(addTileCopy(placements[j].tileIndex), placements[j], symbolDefinitions[placements[j].tileIndex] ? null : tileEntry.info);
            placedCount++;
        }
        if (fillSettings.clipOverflow && placedCount) addClipPath(fillTarget.container, gridGroup);

        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        app.redraw();
    }

    /**
     * メイン処理を実行し、例外はアラートで知らせる
     * @returns {void}
     */
    function runMainWithErrorAlert() {
        /* 実行中は境界線を隠し、終わったら（キャンセル・エラーでも）切り替え直す
           Hide edges while running and toggle them back afterwards, even on cancel or error */
        var isEdgeToggled = toggleEdges();
        try {
            main();
        } catch (e) {
            alert('[' + SCRIPT_NAME + '] Error:\n' + e);
        } finally {
            if (isEdgeToggled) toggleEdges();
        }
    }

    /**
     * ［境界線を隠す／表示］を切り替える（ドキュメントが無いときは何もしない）
     * @returns {boolean} 切り替えたら true
     */
    function toggleEdges() {
        if (!app.documents.length) return false;
        try {
            app.executeMenuCommand('edge');
            return true;
        } catch (e) {
            return false;
        }
    }

    /* Illustrator には app.doScript が無いので、そのまま実行する / Illustrator has no app.doScript, so run directly */
    runMainWithErrorAlert();

})();
