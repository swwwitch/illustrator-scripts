#target illustrator
#targetengine "PlaceImagesInGridEngine"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.PlaceImagesInGrid
Name=画像を選んで長方形やアートボードにグリッド状に配置
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.2
Release-Date=2026-11-12
Description-BEGIN
ファイル選択ダイアログボックスで選んだ画像をリンクで配置し、選択中の長方形（または作業中のアートボード）にグリッド状に並べます。
列数・行数・マージンと、各画像のマスク（角丸・線）は、配置後のダイアログボックスでプレビューを見ながら調整できます。

Places the images chosen in the file dialog box as links and lays them out in a grid inside the selected rectangle (or the active artboard).
Columns, rows, margins and a mask for each image (corner radius and stroke) can be adjusted with a live preview in the dialog box that opens after placing.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

ファイル選択ダイアログボックスで選んだ画像をリンクで配置し、選択中の長方形（または作業中のアートボード）にグリッド状に並べます。
列数・行数・マージンと、各画像のマスク（角丸・線）は、配置後のダイアログボックスでプレビューを見ながら調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PlaceImagesInGrid.md

### Overview

Places the images chosen in the file dialog box as links and lays them out in a grid inside the selected rectangle (or the active artboard).
Columns, rows, margins and a mask for each image (corner radius and stroke) can be adjusted with a live preview in the dialog box that opens after placing.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PlaceImagesInGrid.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PlaceImagesInGrid";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-12";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "画像を選んで長方形やアートボードにグリッド状に配置"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PlaceImagesInGrid.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PlaceImagesInGrid.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/PlaceImagesInGrid.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 対象の拡張子 / Target extensions */
    var IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "tif", "tiff", "psd", "gif", "bmp", "pdf", "eps", "ai", "svg", "heic", "webp"];

    /* 間隔（画像どうしのマージン）の初期値（pt） / Default margin between images (pt) */
    var DEFAULT_IMAGE_GAP = 10;

    /* 対象（長方形・アートボード）の内側に取るマージンの初期値（pt） / Default margin inside the target (pt) */
    var DEFAULT_FRAME_MARGIN = 0;

    /* マスクの種類 / Mask modes */
    var MASK_MODE_NONE = "none"; /* マスクしない / no mask */
    var MASK_MODE_CELL = "cell"; /* セルの形（セルを覆う大きさにして切り抜く） / cell shape (scaled to cover the cell, then clipped) */
    var MASK_MODE_ITEM = "item"; /* アイテムの形（セルに収めた画像の大きさ） / item shape (the image fitted in its cell) */

    /* マスクの種類・角丸の半径（pt）の初期値 / Defaults for the mask mode and the corner radius (pt) */
    var DEFAULT_MASK_MODE = MASK_MODE_NONE;
    var DEFAULT_USE_CORNER_RADIUS = false;
    var DEFAULT_CORNER_RADIUS = 10;

    /* マスクの長方形に線を付けるか・線幅（pt）の初期値。色の初期値は黒 / Defaults for stroking the mask rectangle and its width (pt); the default color is black */
    var DEFAULT_USE_STROKE = false;
    var DEFAULT_STROKE_WIDTH = 1;

    /* 列数・行数の決め方 / How columns and rows are set */
    var GRID_MODE_AUTO = "auto";
    var GRID_MODE_MANUAL = "manual";

    /* 最終行の揃え / Alignment of the last row */
    var LAST_ROW_ALIGN_LEFT = "left";
    var LAST_ROW_ALIGN_CENTER = "center";
    var DEFAULT_LAST_ROW_ALIGN = LAST_ROW_ALIGN_CENTER; /* 初期値 / default */

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

    var FIELD_LABEL_WIDTH = 52;  /* 項目名の幅（グリッド・マージン。「最終行 :」が入る幅） / label width in the grid and margin panels (fits the last-row label) */
    var MASK_OPTIONS_LABEL_WIDTH = 64; /* マスクのオプションの項目名の幅（チェックボックスが入る） / label width in the mask options panel (holds checkboxes) */
    var LAST_ROW_TOP_MARGIN = 4; /* 「最終行」の行の上に足す余白 / extra space above the last-row option */
    var COLOR_SWATCH_SIZE = [40, 20]; /* 線の色見本の大きさ / size of the stroke color swatch */

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
    // UI の部品 / UI parts
    // =========================================

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
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
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
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
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
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
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
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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
    // 設定の保存 / Settings store
    // =========================================

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^\uFEFF/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    /* 保存する設定の既定値（長さは pt） / Default settings to store (lengths in pt) */
    var DEFAULT_SETTINGS = {
        gridMode: GRID_MODE_AUTO,
        columns: 1,
        lastRowAlign: DEFAULT_LAST_ROW_ALIGN,
        gapPt: DEFAULT_IMAGE_GAP,
        frameMarginPt: DEFAULT_FRAME_MARGIN,
        maskMode: DEFAULT_MASK_MODE,
        useCornerRadius: DEFAULT_USE_CORNER_RADIUS,
        cornerRadiusPt: DEFAULT_CORNER_RADIUS,
        useStroke: DEFAULT_USE_STROKE,
        strokeWidthPt: DEFAULT_STROKE_WIDTH,
        strokeColor: null
    };

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

    var LABELS = {
        dialog: {
            title: { ja: "画像をグリッドに配置", en: "Place Images in Grid" },
            chooseImages: { ja: "配置する画像を選択", en: "Choose images to place" }
        },
        panel: {
            grid: { ja: "グリッド", en: "Grid" },
            margin: { ja: "マージン", en: "Margins" },
            mask: { ja: "マスク形状", en: "Mask shape" },
            maskOptions: { ja: "マスクのオプション", en: "Mask options" }
        },
        fieldLabel: {
            columns: { ja: "列数", en: "Columns" },
            rows: { ja: "行数", en: "Rows" },
            gap: { ja: "間隔", en: "Gutter" },
            frameMargin: { ja: "外周", en: "Inset" },
            strokeColor: { ja: "線", en: "Stroke" },
            lastRow: { ja: "最終行", en: "Last row" }
        },
        radio: {
            maskNone: { ja: "なし", en: "None" },
            maskCell: { ja: "セル", en: "Cell" },
            maskItem: { ja: "画像", en: "Image" },
            gridAuto: { ja: "自動", en: "Auto" },
            gridManual: { ja: "手動", en: "Manual" },
            lastRowLeft: { ja: "左揃え", en: "Align Left" },
            lastRowCenter: { ja: "中央揃え", en: "Align Center" }
        },
        checkbox: {
            useCornerRadius: { ja: "角丸", en: "Corner radius" },
            useStroke: { ja: "線幅", en: "Stroke weight" }
        },
        button: {
            shuffle: { ja: "シャッフル", en: "Shuffle" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            gridAuto: { ja: "並べる範囲がいちばん埋まる列数・行数にする", en: "Use the columns and rows that fill the area best" },
            shuffle: { ja: "画像の並び順をランダムに入れ替える", en: "Rearrange the images in random order" },
            strokeColor: { ja: "クリックで線の色を選ぶ", en: "Click to choose the stroke color" },
            maskCell: { ja: "セルを覆う大きさにして、セルの形で切り抜く（大きさがそろう）", en: "Scale each image to cover its cell and clip it to the cell (all the same size)" },
            maskItem: { ja: "画像と同じ大きさでマスクする（切り取られる部分はない）", en: "Mask each image at its own size (nothing is cropped)" },
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
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document first." },
            notRectangle: {
                ja: "長方形を1つ選択してから実行してください。\n何も選択していないときは、アートボードに並べます。",
                en: "Select one rectangle before running the script.\nWith nothing selected, images are laid out on the artboard."
            },
            layerUnavailable: {
                ja: "作業中のレイヤー「{name}」がロックまたは非表示のため配置できません。",
                en: "Cannot place images because the active layer \"{name}\" is locked or hidden."
            },
            placeFailed: {
                ja: "次の画像を配置できませんでした。\n{list}",
                en: "The following images could not be placed.\n{list}"
            }
        }
    };

    // =========================================
    // ファイル選択 / File selection
    // =========================================

    /**
     * ファイル名の拡張子が対象の画像かどうかを返す。
     * @param {string} fileName - ファイル名
     * @returns {boolean} 対象の拡張子なら true
     */
    function isImageName(fileName) {
        var extMatch = /\.([^.]+)$/.exec(fileName);
        if (!extMatch) return false;
        var ext = extMatch[1].toLowerCase();
        for (var i = 0; i < IMAGE_EXTENSIONS.length; i++) {
            if (IMAGE_EXTENSIONS[i] === ext) return true;
        }
        return false;
    }

    /**
     * ファイル選択ダイアログボックスを開いて画像を選ばせる（複数選択可）。
     * @returns {File[]} 選んだファイル。キャンセル時は空の配列
     */
    function chooseImageFiles() {
        var fileFilter;
        if ($.os.indexOf("Windows") !== -1) {
            /* Windows はフィルター文字列 / Windows uses a filter string */
            fileFilter = "Images:*." + IMAGE_EXTENSIONS.join(";*.");
        } else {
            /* Mac は関数（フォルダーは辿れるように true） / Mac uses a filter function */
            fileFilter = function (f) {
                return (f instanceof Folder) || isImageName(f.name);
            };
        }
        var chosen = File.openDialog(getLabel("dialog.chooseImages"), fileFilter, true);
        if (!chosen) return [];
        return (chosen instanceof Array) ? chosen : [chosen];
    }

    // =========================================
    // 配置 / Placement
    // =========================================

    /**
     * 画像を配置する。失敗したファイルは failedNames に名前を積む。
     * @param {Document} doc - 配置先のドキュメント
     * @param {File[]} imageFiles - 配置する画像
     * @param {string[]} failedNames - 配置できなかったファイル名を積む配列
     * @returns {PlacedItem[]} 配置したアイテム
     */
    function placeImages(doc, imageFiles, failedNames) {
        var placedItems = [];
        for (var i = 0; i < imageFiles.length; i++) {
            var placedItem = doc.placedItems.add();
            try {
                placedItem.file = imageFiles[i];
                placedItems.push(placedItem);
            } catch (e) {
                placedItem.remove();
                failedNames.push(decodeURI(imageFiles[i].name));
            }
        }
        return placedItems;
    }

    /**
     * 配置先の長方形を選択から取り出す。
     * @param {Document} doc - 対象のドキュメント
     * @returns {PathItem|null} 選択中の長方形。何も選択していなければ null
     * @throws {Error} 長方形1つ以外を選択しているとき
     */
    function getTargetRectangle(doc) {
        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0) return null;
        if (selectedItems.length !== 1 || selectedItems[0].typename !== "PathItem") {
            throw new Error(getLabel("alert.notRectangle"));
        }
        return selectedItems[0];
    }

    // =========================================
    // 行列の計算と並べ方 / Grid calculation and layout
    // =========================================

    /**
     * @typedef {Object} GridLayout
     * @property {number} columns - 列数
     * @property {number} rows - 行数
     * @property {number} cellWidth - セルの幅（pt）
     * @property {number} cellHeight - セルの高さ（pt）
     */

    /**
     * 行列の数と枠の大きさからセルの大きさを求める。
     * @param {number} columns - 列数
     * @param {number} rows - 行数
     * @param {number} frameWidth - 枠の幅（pt）
     * @param {number} frameHeight - 枠の高さ（pt）
     * @param {number} gap - 画像どうしの間隔（pt）
     * @returns {GridLayout|null} 行列とセルの大きさ。間隔だけで枠が埋まるときは null
     */
    function makeGrid(columns, rows, frameWidth, frameHeight, gap) {
        var cellWidth = (frameWidth - gap * (columns - 1)) / columns;
        var cellHeight = (frameHeight - gap * (rows - 1)) / rows;
        if (cellWidth <= 0 || cellHeight <= 0) return null;
        return { columns: columns, rows: rows, cellWidth: cellWidth, cellHeight: cellHeight };
    }

    /**
     * 枠がいちばん埋まる行列の数を求める。
     * 列数を1〜画像数まで試し、各画像をセルに収めたときの面積の合計が最大になるものを選ぶ。
     * @param {PlacedItem[]} placedItems - 並べるアイテム
     * @param {number} frameWidth - 枠の幅（pt）
     * @param {number} frameHeight - 枠の高さ（pt）
     * @param {number} gap - 画像どうしの間隔（pt）
     * @returns {GridLayout|null} 行列とセルの大きさ。どの行列でも間隔だけで枠が埋まるときは null
     */
    function findBestGrid(placedItems, frameWidth, frameHeight, gap) {
        var imageCount = placedItems.length;
        var bestGrid = null;
        var bestCoveredArea = 0;
        for (var columns = 1; columns <= imageCount; columns++) {
            var rows = Math.ceil(imageCount / columns);
            /* 同じ行数なら列の少ないほうで足りる / Skip column counts that do not reduce rows */
            if (columns > 1 && Math.ceil(imageCount / (columns - 1)) === rows) continue;
            var grid = makeGrid(columns, rows, frameWidth, frameHeight, gap);
            if (!grid) continue;

            var coveredArea = 0;
            for (var i = 0; i < imageCount; i++) {
                var fitScale = Math.min(grid.cellWidth / placedItems[i].width, grid.cellHeight / placedItems[i].height);
                coveredArea += placedItems[i].width * placedItems[i].height * fitScale * fitScale;
            }
            if (coveredArea > bestCoveredArea) {
                bestCoveredArea = coveredArea;
                bestGrid = grid;
            }
        }
        return bestGrid;
    }

    /**
     * 行数を指定されたときの列数を求める（左上から行ごとに詰めるので、行数は列数から決まる）。
     * 増やしたのに実際の行数が変わらないときは、変わるまで列を減らす（▲で値が戻って進まなくなるのを防ぐ）。
     * @param {number} imageCount - 画像の数
     * @param {number} requestedRows - 指定された行数
     * @param {number} previousRows - 直前の行数
     * @returns {number} 列数
     */
    function columnsForRows(imageCount, requestedRows, previousRows) {
        var columns = Math.ceil(imageCount / requestedRows);
        if (requestedRows > previousRows) {
            while (columns > 1 && Math.ceil(imageCount / columns) <= previousRows) columns--;
        }
        return columns;
    }

    /**
     * 配置したアイテムを等倍のまま横一列に並べ、全体を枠の中央に置く（枠が小さすぎて行列に並べられないとき用）。
     * @param {PlacedItem[]} placedItems - 並べるアイテム（選んだ順）
     * @param {number[]} frameBounds - 枠の [左, 上, 右, 下]（ドキュメント座標）
     * @param {number} gap - 画像どうしの間隔（pt）
     * @returns {void}
     */
    function arrangeInRow(placedItems, frameBounds, gap) {
        var rowWidth = gap * (placedItems.length - 1);
        for (var i = 0; i < placedItems.length; i++) {
            rowWidth += placedItems[i].width;
        }

        /* 左から順に、天地中央で並べる / Place left to right, vertically centered */
        var centerX = (frameBounds[0] + frameBounds[2]) / 2;
        var centerY = (frameBounds[1] + frameBounds[3]) / 2;
        var nextLeft = centerX - rowWidth / 2;
        for (var j = 0; j < placedItems.length; j++) {
            var item = placedItems[j];
            item.position = [nextLeft, centerY + item.height / 2];
            nextLeft += item.width + gap;
        }
    }

    /**
     * 配置したアイテムを行列に並べる。
     * 各画像は縦横比を保ったまま、セルに収まる大きさ（fillCells が true ならセルを覆う大きさ）に拡大・縮小する
     * （何度呼んでも同じ大きさになる）。
     * 行・列の大きさは、収めるときは実際の画像に詰め、覆うときはセルの大きさにする。画像どうしの間隔は縦横とも gap にそろえる。
     * 余った分は枠の上下（左右）に均等に振り分ける。端数の出た最終行は lastRowAlign で左揃えか左右中央にする。
     * @param {PlacedItem[]} placedItems - 並べるアイテム（選んだ順に左上から右へ）
     * @param {number[]} frameBounds - 枠の [左, 上, 右, 下]（ドキュメント座標）
     * @param {GridLayout} grid - 行列とセルの大きさ
     * @param {number} gap - 画像どうしの間隔（pt）
     * @param {string} lastRowAlign - 最終行の揃え（"left" または "center"）
     * @param {boolean} fillCells - セルを覆う大きさにする（マスク用）なら true
     * @returns {number[][]} 各アイテムの区画 [左, 上, 幅, 高さ]（placedItems と同じ順）
     */
    function arrangeGridInFrame(placedItems, frameBounds, grid, gap, lastRowAlign, fillCells) {
        var frameWidth = frameBounds[2] - frameBounds[0];
        var frameHeight = frameBounds[1] - frameBounds[3];
        var imageCount = placedItems.length;
        var usedRows = Math.ceil(imageCount / grid.columns);
        var slotRects = [];
        var i;

        /* セルに合わせて拡大・縮小し、列の幅・行の高さを求める / Scale to the cell, then measure columns and rows */
        var columnWidths = [];
        var rowHeights = [];
        for (i = 0; i < grid.columns; i++) columnWidths[i] = fillCells ? grid.cellWidth : 0;
        for (i = 0; i < usedRows; i++) rowHeights[i] = fillCells ? grid.cellHeight : 0;
        for (i = 0; i < imageCount; i++) {
            var scaledItem = placedItems[i];
            var widthScale = grid.cellWidth / scaledItem.width;
            var heightScale = grid.cellHeight / scaledItem.height;
            var cellScale = fillCells ? Math.max(widthScale, heightScale) : Math.min(widthScale, heightScale);
            scaledItem.resize(cellScale * 100, cellScale * 100);
            if (fillCells) continue;
            var columnOf = i % grid.columns;
            var rowOf = Math.floor(i / grid.columns);
            if (scaledItem.width > columnWidths[columnOf]) columnWidths[columnOf] = scaledItem.width;
            if (scaledItem.height > rowHeights[rowOf]) rowHeights[rowOf] = scaledItem.height;
        }

        /* 全体の大きさ / Size of the whole block */
        var blockColumns = Math.min(grid.columns, imageCount);
        var blockWidth = sumWithGaps(columnWidths, blockColumns, gap);
        var blockHeight = sumWithGaps(rowHeights, usedRows, gap);
        var blockLeft = frameBounds[0] + (frameWidth - blockWidth) / 2;
        var rowTop = frameBounds[1] - (frameHeight - blockHeight) / 2;

        for (var rowIndex = 0; rowIndex < usedRows; rowIndex++) {
            var firstIndex = rowIndex * grid.columns;
            var rowCount = Math.min(grid.columns, imageCount - firstIndex);

            /* 端数の出た最終行は左揃えか左右中央 / A short last row is left-aligned or centered */
            var cellLeft = blockLeft;
            if (lastRowAlign === LAST_ROW_ALIGN_CENTER) {
                cellLeft += (blockWidth - sumWithGaps(columnWidths, rowCount, gap)) / 2;
            }
            for (var columnIndex = 0; columnIndex < rowCount; columnIndex++) {
                var item = placedItems[firstIndex + columnIndex];
                item.position = [
                    cellLeft + (columnWidths[columnIndex] - item.width) / 2,
                    rowTop - (rowHeights[rowIndex] - item.height) / 2
                ];
                slotRects[firstIndex + columnIndex] = [cellLeft, rowTop, columnWidths[columnIndex], rowHeights[rowIndex]];
                cellLeft += columnWidths[columnIndex] + gap;
            }
            rowTop -= rowHeights[rowIndex] + gap;
        }
        return slotRects;
    }

    /**
     * 先頭から count 個の寸法を、間に gap をはさんで合計する。
     * @param {number[]} sizes - 列の幅または行の高さ
     * @param {number} count - 合計する個数
     * @param {number} gap - 間隔（pt）
     * @returns {number} 合計（pt）
     */
    function sumWithGaps(sizes, count, gap) {
        var total = gap * (count - 1);
        for (var i = 0; i < count; i++) total += sizes[i];
        return total;
    }

    /**
     * 配列の並び順をランダムに入れ替える（Fisher–Yates。元の配列を書き換える）。
     * 2つ以上あるときは、元と同じ並びにはしない。
     * @param {PlacedItem[]} items - 入れ替える配列
     * @returns {void}
     */
    function shuffleItemsInPlace(items) {
        if (items.length < 2) return;
        var originalOrder = items.slice(0);
        var isSameOrder = true;
        while (isSameOrder) {
            for (var i = items.length - 1; i > 0; i--) {
                var j = Math.floor(Math.random() * (i + 1));
                var swapped = items[i];
                items[i] = items[j];
                items[j] = swapped;
            }
            isSameOrder = true;
            for (var k = 0; k < items.length; k++) {
                if (items[k] !== originalOrder[k]) {
                    isSameOrder = false;
                    break;
                }
            }
        }
    }

    /**
     * 枠を内側へ縮める（左右・上下に同じ量）。
     * @param {number[]} bounds - 枠の [左, 上, 右, 下]（ドキュメント座標）
     * @param {number} inset - 内側へ縮める量（pt）
     * @returns {number[]} 縮めた [左, 上, 右, 下]
     */
    function insetBounds(bounds, inset) {
        return [bounds[0] + inset, bounds[1] - inset, bounds[2] - inset, bounds[3] + inset];
    }

    /**
     * 行列の数と間隔で並べ直す。間隔だけで枠が埋まるときは等倍で横一列にする。
     * @param {PlacedItem[]} placedItems - 並べるアイテム
     * @param {number[]} frameBounds - 枠の [左, 上, 右, 下]（ドキュメント座標）
     * @param {number} columns - 列数
     * @param {number} rows - 行数
     * @param {number} gap - 画像どうしの間隔（pt）
     * @param {string} lastRowAlign - 最終行の揃え（"left" または "center"）
     * @param {boolean} fillCells - セルを覆う大きさにする（マスク用）なら true
     * @returns {number[][]|null} 各アイテムの区画 [左, 上, 幅, 高さ]。横一列にしたときは null
     */
    function layoutPlacedItems(placedItems, frameBounds, columns, rows, gap, lastRowAlign, fillCells) {
        var grid = makeGrid(columns, rows, frameBounds[2] - frameBounds[0], frameBounds[1] - frameBounds[3], gap);
        if (grid) {
            return arrangeGridInFrame(placedItems, frameBounds, grid, gap, lastRowAlign, fillCells);
        }
        arrangeInRow(placedItems, frameBounds, gap);
        return null;
    }

    // =========================================
    // マスク / Masks
    // =========================================

    /**
     * 角丸の効果（Adobe Round Corners）の XML を返す。
     * @param {number} radius - 半径（pt）
     * @returns {string} applyEffect() に渡す XML
     */
    function roundCornersEffectXml(radius) {
        return '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>';
    }

    /**
     * 各アイテムの外接矩形を区画の形 [左, 上, 幅, 高さ] で返す。
     * @param {PlacedItem[]} placedItems - 対象のアイテム
     * @returns {number[][]} 各アイテムの [左, 上, 幅, 高さ]
     */
    function itemRects(placedItems) {
        var rects = [];
        for (var i = 0; i < placedItems.length; i++) {
            var item = placedItems[i];
            rects.push([item.position[0], item.position[1], item.width, item.height]);
        }
        return rects;
    }

    /**
     * 各アイテムを区画の長方形でクリップしたグループに入れる。半径が0より大きければ、グループに角丸の効果を掛ける。
     * @param {Document} doc - 対象のドキュメント
     * @param {PlacedItem[]} placedItems - マスクするアイテム
     * @param {number[][]} slotRects - 各アイテムの区画 [左, 上, 幅, 高さ]
     * @param {number} cornerRadius - 角丸の半径（pt）
     * @param {{color: Color, width: number}|null} maskStroke - クリップパスに付ける線（付けないなら null）
     * @returns {GroupItem[]} 作ったクリップグループ
     */
    function wrapItemsInMasks(doc, placedItems, slotRects, cornerRadius, maskStroke) {
        var maskGroups = [];
        for (var i = 0; i < placedItems.length; i++) {
            var item = placedItems[i];
            var slotRect = slotRects[i];
            var maskGroup = item.layer.groupItems.add();
            maskGroup.move(item, ElementPlacement.PLACEBEFORE);
            item.move(maskGroup, ElementPlacement.PLACEATEND);

            /* クリップパスは区画の長方形（グループのいちばん上） / The clipping path is the slot rectangle, on top */
            var maskPath = doc.pathItems.rectangle(slotRect[1], slotRect[0], slotRect[2], slotRect[3]);
            maskPath.move(maskGroup, ElementPlacement.PLACEATBEGINNING);
            maskGroup.clipped = true;

            /* 線はマスクにしたあとで付ける（マスクにすると塗り・線が消えるため） / Stroke after clipping, which clears fill and stroke */
            if (maskStroke && maskStroke.width > 0) {
                maskPath.stroked = true;
                maskPath.strokeColor = maskStroke.color;
                maskPath.strokeWidth = maskStroke.width;
            }

            if (cornerRadius > 0) {
                maskGroup.applyEffect(roundCornersEffectXml(cornerRadius));
            }
            maskGroups.push(maskGroup);
        }
        return maskGroups;
    }

    /**
     * クリップグループからアイテムを出し、グループ（とクリップパス）を削除する。
     * @param {GroupItem[]} maskGroups - wrapItemsInMasks() で作ったグループ
     * @returns {void}
     */
    function unwrapMasks(maskGroups) {
        for (var i = 0; i < maskGroups.length; i++) {
            var maskGroup = maskGroups[i];
            /* 画像はクリップパス以外の子 / The image is the child that is not the clipping path */
            for (var j = maskGroup.pageItems.length - 1; j >= 0; j--) {
                var child = maskGroup.pageItems[j];
                if (child.typename === "PlacedItem") child.move(maskGroup, ElementPlacement.PLACEBEFORE);
            }
            maskGroup.remove();
        }
        maskGroups.length = 0;
    }

    // =========================================
    // 線の色 / Stroke color
    // =========================================

    /**
     * ドキュメントのカラーモードの黒を返す。
     * @param {Document} doc - 対象のドキュメント
     * @returns {Color} RGBColor または CMYKColor
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = 100;
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = 0;
        rgbColor.green = 0;
        rgbColor.blue = 0;
        return rgbColor;
    }

    /**
     * 色を設定に保存できる形にする（RGB・CMYK・グレー。それ以外は null）。
     * @param {Color} strokeColor - 色
     * @returns {Object|null} { model: "rgb"|"cmyk"|"gray", values: number[] }
     */
    function colorToSettings(strokeColor) {
        if (strokeColor.typename === "RGBColor") return { model: "rgb", values: [strokeColor.red, strokeColor.green, strokeColor.blue] };
        if (strokeColor.typename === "CMYKColor") return { model: "cmyk", values: [strokeColor.cyan, strokeColor.magenta, strokeColor.yellow, strokeColor.black] };
        if (strokeColor.typename === "GrayColor") return { model: "gray", values: [strokeColor.gray] };
        return null;
    }

    /**
     * 保存した色を Color に戻す。読めないときはドキュメントのカラーモードの黒。
     * @param {Document} doc - 対象のドキュメント
     * @param {Object|null} savedColor - colorToSettings() の戻り値
     * @returns {Color} 色
     */
    function colorFromSettings(doc, savedColor) {
        if (!savedColor || !(savedColor.values instanceof Array)) return createBlackColor(doc);
        var values = savedColor.values;
        if (savedColor.model === "rgb" && values.length === 3) {
            var rgbColor = new RGBColor();
            rgbColor.red = values[0];
            rgbColor.green = values[1];
            rgbColor.blue = values[2];
            return rgbColor;
        }
        if (savedColor.model === "cmyk" && values.length === 4) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = values[0];
            cmykColor.magenta = values[1];
            cmykColor.yellow = values[2];
            cmykColor.black = values[3];
            return cmykColor;
        }
        if (savedColor.model === "gray" && values.length === 1) {
            var grayColor = new GrayColor();
            grayColor.gray = values[0];
            return grayColor;
        }
        return createBlackColor(doc);
    }

    /**
     * 色を色見本に描くための RGB（0〜1）にする（CMYK・グレーは簡易換算）。
     * @param {Color} swatchColor - 色
     * @returns {number[]} [r, g, b]（0〜1）
     */
    function colorToScreenRgb(swatchColor) {
        if (swatchColor.typename === "RGBColor") return [swatchColor.red / 255, swatchColor.green / 255, swatchColor.blue / 255];
        if (swatchColor.typename === "GrayColor") return [1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100];
        if (swatchColor.typename === "CMYKColor") {
            var blackRatio = 1 - swatchColor.black / 100;
            return [
                (1 - swatchColor.cyan / 100) * blackRatio,
                (1 - swatchColor.magenta / 100) * blackRatio,
                (1 - swatchColor.yellow / 100) * blackRatio
            ];
        }
        /* スポットカラーなどは灰色で示す / show other color types in gray */
        return [0.5, 0.5, 0.5];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {Object} GridSettings
     * @property {string} gridMode - 列数・行数の決め方（"auto" または "manual"）
     * @property {number} columns - 列数（"manual" のとき）
     * @property {string} lastRowAlign - 最終行の揃え（"left" または "center"）
     * @property {number} gapPt - 間隔（画像どうしのマージン、pt）
     * @property {number} frameMarginPt - 外周（並べる範囲の内側に取るマージン、pt）
     * @property {string} maskMode - マスクの種類（"none" / "cell" / "item"）
     * @property {boolean} useCornerRadius - 角丸を付けるなら true
     * @property {number} cornerRadiusPt - 角丸の半径（pt）
     * @property {boolean} useStroke - マスクの長方形に線を付けるなら true
     * @property {number} strokeWidthPt - 線幅（pt）
     * @property {Object|null} strokeColor - 線の色（colorToSettings() の形）
     */

    /**
     * 配置したアイテムをプレビューしながら、列数・行数・マージン・マスクなどを指定するダイアログボックスを表示する。
     * @param {Document} doc - 対象のドキュメント
     * @param {PlacedItem[]} placedItems - 配置したアイテム
     * @param {number[]} frameBounds - 対象（長方形・アートボード）の [左, 上, 右, 下]（ドキュメント座標）
     * @param {GridSettings} initialSettings - 最初の設定（前回の設定）
     * @param {GroupItem[]} maskGroups - マスクのクリップグループを入れる配列（ダイアログボックスの中で作り直す）
     * @returns {GridSettings|null} OK なら確定した設定、キャンセルなら null
     */
    function showGridDialog(doc, placedItems, frameBounds, initialSettings, maskGroups) {
        var imageCount = placedItems.length;
        var rulerUnit = getUnitInfo();
        var strokeUnit = getUnitInfo("strokeUnits");
        var rulerSuffix = " " + rulerUnit.label;

        /* show() 前に代入した value は読み戻せないので、状態は変数で持つ / Keep state in variables; values set before show() cannot be read back */
        var isGridAuto = (initialSettings.gridMode !== GRID_MODE_MANUAL);
        var currentGap = initialSettings.gapPt;
        var currentFrameMargin = initialSettings.frameMarginPt;
        var currentLastRowAlign = initialSettings.lastRowAlign;
        var currentMaskMode = initialSettings.maskMode;
        if (currentMaskMode !== MASK_MODE_CELL && currentMaskMode !== MASK_MODE_ITEM) currentMaskMode = MASK_MODE_NONE;
        var currentUseMask = (currentMaskMode !== MASK_MODE_NONE);
        var currentUseCornerRadius = initialSettings.useCornerRadius;
        var currentCornerRadius = initialSettings.cornerRadiusPt;
        var currentUseStroke = initialSettings.useStroke;
        var currentStrokeWidth = initialSettings.strokeWidthPt;
        var currentStrokeColor = colorFromSettings(doc, initialSettings.strokeColor);
        var currentColumns = Math.max(1, Math.min(imageCount, Math.round(initialSettings.columns)));
        var currentRows = Math.ceil(imageCount / currentColumns);

        /**
         * いちばん埋まる行列の数を今の値にする。
         * @returns {void}
         */
        function applyBestGrid() {
            var innerBounds = insetBounds(frameBounds, currentFrameMargin);
            var bestGrid = findBestGrid(placedItems, innerBounds[2] - innerBounds[0], innerBounds[1] - innerBounds[3], currentGap);
            currentColumns = bestGrid ? bestGrid.columns : imageCount;
            currentRows = bestGrid ? bestGrid.rows : 1;
        }

        /**
         * 今の値で並べ直して画面を描き直す。
         * @returns {void}
         */
        function updatePreview() {
            /* マスクは並べる前に外し、並べたあとで作り直す / Remove the masks before laying out, then rebuild them */
            unwrapMasks(maskGroups);
            var slotRects = layoutPlacedItems(placedItems, insetBounds(frameBounds, currentFrameMargin),
                currentColumns, currentRows, currentGap, currentLastRowAlign, currentMaskMode === MASK_MODE_CELL);
            if (currentUseMask && slotRects) {
                /* アイテムの形なら、セルに収めた画像そのものの大きさでマスク / Item shape: mask to the fitted image itself */
                var maskRects = (currentMaskMode === MASK_MODE_ITEM) ? itemRects(placedItems) : slotRects;
                var maskStroke = currentUseStroke ? { color: currentStrokeColor, width: currentStrokeWidth } : null;
                var newGroups = wrapItemsInMasks(doc, placedItems, maskRects, currentUseCornerRadius ? currentCornerRadius : 0, maskStroke);
                for (var i = 0; i < newGroups.length; i++) maskGroups.push(newGroups[i]);
            }
            app.redraw();
        }

        if (isGridAuto) applyBestGrid();
        updatePreview();

        var gridDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(gridDialog);

        /* 2カラム：左にグリッド・マージン、右にマスク・オプション / Two columns: grid and margins on the left, mask and options on the right */
        var columnsGroup = gridDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = WINDOW_SPACING;
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.spacing = WINDOW_SPACING;

        // -----------------------------------------
        // グリッド / Grid
        // -----------------------------------------
        var gridPanel = leftColumn.add("panel", undefined, getLabel("panel.grid"));
        setupPanel(gridPanel, 6);

        var columnsFieldOptions = {
            label: labelText("fieldLabel.columns"), labelWidth: FIELD_LABEL_WIDTH,
            text: String(currentColumns), characters: 5, step: 1, min: 1, max: imageCount, integer: true,
            onStep: function () { onColumnsChanged(); }
        };
        var rowsFieldOptions = {
            label: labelText("fieldLabel.rows"), labelWidth: FIELD_LABEL_WIDTH,
            text: String(currentRows), characters: 5, step: 1, min: 1, max: imageCount, integer: true,
            onStep: function () { onRowsChanged(); }
        };

        /* 列数・行数の決め方（自動・指定） / How columns and rows are set (auto / manual) */
        var gridModeGroup = gridPanel.add("group");
        setupRow(gridModeGroup, "center");
        var gridAutoRadio = gridModeGroup.add("radiobutton", undefined, getLabel("radio.gridAuto"));
        var gridManualRadio = gridModeGroup.add("radiobutton", undefined, getLabel("radio.gridManual"));
        gridAutoRadio.helpTip = getLabel("tooltip.gridAuto");
        gridAutoRadio.value = isGridAuto;
        gridManualRadio.value = !isGridAuto;

        var columnsInput = addSteppedField(gridPanel, columnsFieldOptions);
        var rowsInput = addSteppedField(gridPanel, rowsFieldOptions);
        setSteppedFieldEnabled(columnsInput, !isGridAuto);
        setSteppedFieldEnabled(rowsInput, !isGridAuto);

        /* 最終行の揃え（左・中央） / Alignment of the last row (left / center) */
        var lastRowGroup = gridPanel.add("group");
        setupRow(lastRowGroup, "left", STEPPER_FIELD_SPACING);
        lastRowGroup.margins = [0, LAST_ROW_TOP_MARGIN, 0, 0]; /* 列数・行数と少し離す / a little apart from the columns and rows */
        var lastRowLabel = lastRowGroup.add("statictext", undefined, labelText("fieldLabel.lastRow"));
        lastRowLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        lastRowLabel.justify = "right";
        /* ラジオボタンは縦に並べ、項目名は1つ目の行に合わせる / Stack the radios; the label lines up with the first one */
        lastRowGroup.alignChildren = ["left", "top"];
        var lastRowRadioGroup = lastRowGroup.add("group");
        lastRowRadioGroup.orientation = "column";
        lastRowRadioGroup.alignChildren = ["left", "center"];
        lastRowRadioGroup.spacing = 6;
        var lastRowLeftRadio = lastRowRadioGroup.add("radiobutton", undefined, getLabel("radio.lastRowLeft"));
        var lastRowCenterRadio = lastRowRadioGroup.add("radiobutton", undefined, getLabel("radio.lastRowCenter"));
        lastRowLeftRadio.value = (currentLastRowAlign === LAST_ROW_ALIGN_LEFT);
        lastRowCenterRadio.value = (currentLastRowAlign !== LAST_ROW_ALIGN_LEFT);
        lastRowLeftRadio.onClick = function () {
            currentLastRowAlign = LAST_ROW_ALIGN_LEFT;
            updatePreview();
        };
        lastRowCenterRadio.onClick = function () {
            currentLastRowAlign = LAST_ROW_ALIGN_CENTER;
            updatePreview();
        };

        // -----------------------------------------
        // マージン / Margins
        // -----------------------------------------
        var gapFieldOptions = {
            label: labelText("fieldLabel.gap"), labelWidth: FIELD_LABEL_WIDTH,
            text: "", characters: 5, step: 1, min: 0, unit: rulerSuffix,
            onStep: function () { onMarginChanged(); }
        };
        var frameMarginFieldOptions = {
            label: labelText("fieldLabel.frameMargin"), labelWidth: FIELD_LABEL_WIDTH,
            text: "", characters: 5, step: 1, min: 0, unit: rulerSuffix,
            onStep: function () { onMarginChanged(); }
        };
        var marginPanel = leftColumn.add("panel", undefined, getLabel("panel.margin"));
        setupPanel(marginPanel, 6);
        var gapInput = addSteppedField(marginPanel, gapFieldOptions);
        var frameMarginInput = addSteppedField(marginPanel, frameMarginFieldOptions);
        writeSteppedValue(gapInput, ptToUnit(currentGap, rulerUnit), gapFieldOptions);
        writeSteppedValue(frameMarginInput, ptToUnit(currentFrameMargin, rulerUnit), frameMarginFieldOptions);

        // -----------------------------------------
        // マスク（角丸・線） / Mask (corner radius, stroke)
        // -----------------------------------------
        var maskPanel = rightColumn.add("panel", undefined, getLabel("panel.mask"));
        setupPanel(maskPanel, 6);
        /* マスクの種類（なし・セルの形・アイテムの形） / Mask mode (none / cell shape / item shape) */
        var maskModeGroup = maskPanel.add("group");
        maskModeGroup.orientation = "column";
        maskModeGroup.alignChildren = ["left", "center"];
        maskModeGroup.spacing = 6;
        var maskNoneRadio = maskModeGroup.add("radiobutton", undefined, getLabel("radio.maskNone"));
        var maskCellRadio = maskModeGroup.add("radiobutton", undefined, getLabel("radio.maskCell"));
        var maskItemRadio = maskModeGroup.add("radiobutton", undefined, getLabel("radio.maskItem"));
        maskCellRadio.helpTip = getLabel("tooltip.maskCell");
        maskItemRadio.helpTip = getLabel("tooltip.maskItem");
        maskNoneRadio.value = (currentMaskMode === MASK_MODE_NONE);
        maskCellRadio.value = (currentMaskMode === MASK_MODE_CELL);
        maskItemRadio.value = (currentMaskMode === MASK_MODE_ITEM);

        /**
         * 項目名の位置に、右揃えのチェックボックスを置く（「☐ 角丸 :」の形）。
         * @param {Group} rowGroup - 行のグループ
         * @param {string} labelKey - LABELS.checkbox のキー
         * @param {boolean} isChecked - 最初のオンオフ
         * @returns {Checkbox} 作ったチェックボックス
         */
        function addCheckboxLabel(rowGroup, labelKey, isChecked) {
            var checkboxLabelGroup = rowGroup.add("group");
            checkboxLabelGroup.preferredSize.width = MASK_OPTIONS_LABEL_WIDTH;
            checkboxLabelGroup.alignChildren = ["right", "center"];
            checkboxLabelGroup.margins = 0;
            var labelCheckbox = checkboxLabelGroup.add("checkbox", undefined, labelText("checkbox." + labelKey));
            labelCheckbox.alignment = ["right", "center"];
            labelCheckbox.value = isChecked;
            return labelCheckbox;
        }

        /* マスクのオプション（角丸・線・線幅） / Mask options (corner radius, stroke, stroke weight) */
        var maskOptionsPanel = rightColumn.add("panel", undefined, getLabel("panel.maskOptions"));
        setupPanel(maskOptionsPanel, 6);

        /* 角丸：チェックボックス・半径 / Corner radius: checkbox and radius */
        var cornerRadiusRow = maskOptionsPanel.add("group");
        setupRow(cornerRadiusRow, "left", STEPPER_FIELD_SPACING);
        var useCornerRadiusCheckbox = addCheckboxLabel(cornerRadiusRow, "useCornerRadius", currentUseCornerRadius);
        var cornerRadiusFieldOptions = {
            label: "",
            text: "", characters: 5, step: 1, min: 0, unit: rulerSuffix,
            onStep: function () { onCornerRadiusChanged(); }
        };
        var cornerRadiusInput = addSteppedField(cornerRadiusRow, cornerRadiusFieldOptions);
        /* 項目名はチェックボックスが兼ねるので、空の項目名の幅と間隔を詰める / The checkbox is the label, so collapse the empty label */
        cornerRadiusInput.fieldLabel.maximumSize.width = 0;
        cornerRadiusInput.fieldLabel.parent.spacing = 0;
        writeSteppedValue(cornerRadiusInput, ptToUnit(currentCornerRadius, rulerUnit), cornerRadiusFieldOptions);

        /* 線幅：チェックボックス・線幅 / Stroke weight: checkbox and weight */
        var strokeWidthRow = maskOptionsPanel.add("group");
        setupRow(strokeWidthRow, "left", STEPPER_FIELD_SPACING);
        var useStrokeCheckbox = addCheckboxLabel(strokeWidthRow, "useStroke", currentUseStroke);
        var strokeWidthFieldOptions = {
            label: "", characters: 5, step: 1, min: 0, unit: " " + strokeUnit.label,
            onStep: function () { onStrokeWidthChanged(); }
        };
        var strokeWidthInput = addSteppedField(strokeWidthRow, strokeWidthFieldOptions);
        writeSteppedValue(strokeWidthInput, ptToUnit(currentStrokeWidth, strokeUnit), strokeWidthFieldOptions);
        /* 項目名はチェックボックスが兼ねるので、空の項目名の幅と間隔を詰める / The checkbox is the label, so collapse the empty label */
        strokeWidthInput.fieldLabel.maximumSize.width = 0;
        strokeWidthInput.fieldLabel.parent.spacing = 0;

        /* 線：項目名・色見本 / Stroke: label and color swatch */
        var strokeColorRow = maskOptionsPanel.add("group");
        setupRow(strokeColorRow, "left", STEPPER_FIELD_SPACING);
        var strokeColorLabel = strokeColorRow.add("statictext", undefined, labelText("fieldLabel.strokeColor"));
        strokeColorLabel.preferredSize.width = MASK_OPTIONS_LABEL_WIDTH;
        strokeColorLabel.justify = "right";
        /* 色見本の左に、数値欄の∧∨と同じ余白を足す / Same left margin as the steppers of the numeric fields */
        var strokeSwatchHolder = strokeColorRow.add("group");
        strokeSwatchHolder.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0];
        var strokeSwatch = strokeSwatchHolder.add("group");
        strokeSwatch.preferredSize = COLOR_SWATCH_SIZE;
        strokeSwatch.helpTip = getLabel("tooltip.strokeColor");
        strokeSwatch.onDraw = function () {
            var swatchGraphics = strokeSwatch.graphics;
            var swatchWidth = strokeSwatch.size[0];
            var swatchHeight = strokeSwatch.size[1];
            /* 使えないときは薄く描く（自作描画は enabled で淡色にならない） / draw faintly when unavailable */
            var swatchAlpha = (currentUseMask && currentUseStroke) ? 1 : 0.25;
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, colorToScreenRgb(currentStrokeColor).concat(swatchAlpha)));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, swatchAlpha], 1));
        };
        /* クリックで標準のカラーピッカーを開く。キャンセルは渡した色が返る / Clicking opens the standard color picker; Cancel returns the color passed in */
        strokeSwatch.addEventListener("click", function () {
            if (!currentUseMask || !currentUseStroke) return;
            currentStrokeColor = app.showColorPicker(currentStrokeColor);
            redrawStrokeSwatch();
            updatePreview();
        });


        /**
         * 線の色見本を描き直す（group には notify() が無いので隠して再表示する）。
         * @returns {void}
         */
        function redrawStrokeSwatch() {
            strokeSwatch.hide();
            strokeSwatch.show();
        }

        /**
         * マスク・線のオンオフに合わせて、角丸・線の項目をディムする。
         * @returns {void}
         */
        function updateMaskControls() {
            useCornerRadiusCheckbox.enabled = currentUseMask;
            setSteppedFieldEnabled(cornerRadiusInput, currentUseMask && currentUseCornerRadius);
            useStrokeCheckbox.enabled = currentUseMask;
            setSteppedFieldEnabled(strokeWidthInput, currentUseMask && currentUseStroke);
            strokeColorLabel.enabled = currentUseMask && currentUseStroke;
            redrawStrokeSwatch();
        }
        updateMaskControls();

        /**
         * マスクの種類を変えたとき。
         * @param {string} maskMode - "none" / "cell" / "item"
         * @returns {void}
         */
        function onMaskModeChanged(maskMode) {
            currentMaskMode = maskMode;
            currentUseMask = (maskMode !== MASK_MODE_NONE);
            updateMaskControls();
            updatePreview();
        }
        maskNoneRadio.onClick = function () { onMaskModeChanged(MASK_MODE_NONE); };
        maskCellRadio.onClick = function () { onMaskModeChanged(MASK_MODE_CELL); };
        maskItemRadio.onClick = function () { onMaskModeChanged(MASK_MODE_ITEM); };
        useCornerRadiusCheckbox.onClick = function () {
            currentUseCornerRadius = useCornerRadiusCheckbox.value;
            updateMaskControls();
            updatePreview();
        };
        useStrokeCheckbox.onClick = function () {
            currentUseStroke = useStrokeCheckbox.value;
            updateMaskControls();
            updatePreview();
        };

        // -----------------------------------------
        // 入力の処理 / Input handlers
        // -----------------------------------------

        /**
         * 入力欄の確定（直接入力）でも、∧∨と同じ処理を呼ぶ。部品の値の正規化のあとに続ける。
         * @param {EditText} numberInput - addSteppedField() で作った入力欄
         * @param {Function} onCommitted - 確定後に呼ぶ処理
         * @returns {void}
         */
        function chainFieldChange(numberInput, onCommitted) {
            var normalizeValue = numberInput.onChange;
            numberInput.onChange = function () {
                normalizeValue();
                onCommitted();
            };
        }

        /**
         * 列数と行数の欄を今の値に書き直す。
         * @returns {void}
         */
        function writeGridFields() {
            writeSteppedValue(columnsInput, currentColumns, columnsFieldOptions);
            writeSteppedValue(rowsInput, currentRows, rowsFieldOptions);
        }

        /**
         * 列数を変えたとき：行数は全部が入る数にする。
         * @returns {void}
         */
        function onColumnsChanged() {
            currentColumns = parseInt(columnsInput.text, 10);
            currentRows = Math.ceil(imageCount / currentColumns);
            writeGridFields();
            updatePreview();
        }

        /**
         * 行数を変えたとき：列数は全部が入る数にし、行数は実際に使う数にそろえる。
         * @returns {void}
         */
        function onRowsChanged() {
            currentColumns = columnsForRows(imageCount, parseInt(rowsInput.text, 10), currentRows);
            currentRows = Math.ceil(imageCount / currentColumns);
            writeGridFields();
            updatePreview();
        }

        /**
         * マージン（間隔・外周）を変えたとき。
         * @returns {void}
         */
        function onMarginChanged() {
            currentGap = unitToPt(parseFloat(gapInput.text), rulerUnit);
            currentFrameMargin = unitToPt(parseFloat(frameMarginInput.text), rulerUnit);
            if (isGridAuto) {
                /* 自動なら、マージンに合わせて行列を選び直す / In auto mode, re-pick the grid for the new margin */
                applyBestGrid();
                writeGridFields();
            }
            updatePreview();
        }

        /**
         * 角丸の半径を変えたとき。
         * @returns {void}
         */
        function onCornerRadiusChanged() {
            currentCornerRadius = unitToPt(parseFloat(cornerRadiusInput.text), rulerUnit);
            updatePreview();
        }

        /**
         * 線幅を変えたとき。
         * @returns {void}
         */
        function onStrokeWidthChanged() {
            currentStrokeWidth = unitToPt(parseFloat(strokeWidthInput.text), strokeUnit);
            updatePreview();
        }

        /**
         * 自動・指定に合わせて列数・行数の欄を切り替える。自動に戻したときは行列を選び直す。
         * @returns {void}
         */
        function onGridModeChanged() {
            isGridAuto = gridAutoRadio.value;
            setSteppedFieldEnabled(columnsInput, !isGridAuto);
            setSteppedFieldEnabled(rowsInput, !isGridAuto);
            if (isGridAuto) {
                applyBestGrid();
                writeGridFields();
                updatePreview();
            }
        }

        chainFieldChange(columnsInput, onColumnsChanged);
        chainFieldChange(rowsInput, onRowsChanged);
        chainFieldChange(gapInput, onMarginChanged);
        chainFieldChange(frameMarginInput, onMarginChanged);
        chainFieldChange(cornerRadiusInput, onCornerRadiusChanged);
        chainFieldChange(strokeWidthInput, onStrokeWidthChanged);
        gridAutoRadio.onClick = onGridModeChanged;
        gridManualRadio.onClick = onGridModeChanged;

        // -----------------------------------------
        // ボタン / Buttons
        // -----------------------------------------
        var buttonRow = addButtonRow(gridDialog);
        var btnShuffle = buttonRow.leftGroup.add("button", undefined, getLabel("button.shuffle"));
        btnShuffle.helpTip = getLabel("tooltip.shuffle");
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnShuffle.onClick = function () {
            shuffleItemsInPlace(placedItems);
            updatePreview();
        };
        btnCancel.onClick = function () { gridDialog.close(2); };
        btnOK.onClick = function () { gridDialog.close(1); };

        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(gridDialog);
        prepareDialogWindow(gridDialog, SCRIPT_NAME);
        if (gridDialog.show() !== 1) return null;

        return {
            gridMode: isGridAuto ? GRID_MODE_AUTO : GRID_MODE_MANUAL,
            columns: currentColumns,
            lastRowAlign: currentLastRowAlign,
            gapPt: currentGap,
            frameMarginPt: currentFrameMargin,
            maskMode: currentMaskMode,
            useCornerRadius: currentUseCornerRadius,
            cornerRadiusPt: currentCornerRadius,
            useStroke: currentUseStroke,
            strokeWidthPt: currentStrokeWidth,
            strokeColor: colorToSettings(currentStrokeColor)
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理。
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        var activeLayer = doc.activeLayer;
        if (activeLayer.locked || !activeLayer.visible) {
            alert(getLabel("alert.layerUnavailable", { name: activeLayer.name }));
            return;
        }

        var targetRectangle;
        try {
            targetRectangle = getTargetRectangle(doc);
        } catch (e) {
            alert(e.message);
            return;
        }

        var imageFiles = chooseImageFiles();
        if (imageFiles.length === 0) return;

        var failedNames = [];
        var placedItems = placeImages(doc, imageFiles, failedNames);
        if (failedNames.length > 0) {
            alert(getLabel("alert.placeFailed", { list: failedNames.join("\n") }));
        }
        if (placedItems.length === 0) return;

        var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
        var initialSettings = settingsStore.load(DEFAULT_SETTINGS);
        var maskGroups = [];
        var confirmedSettings;

        var savedCoordinateSystem = app.coordinateSystem;
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        try {
            var frameBounds;
            if (targetRectangle) {
                /* 長方形の中に並べ、長方形のすぐ前面へ / Inside the rectangle, just in front of it */
                frameBounds = targetRectangle.geometricBounds;
                for (var i = 0; i < placedItems.length; i++) {
                    placedItems[i].move(targetRectangle, ElementPlacement.PLACEBEFORE);
                }
            } else {
                /* 作業中のアートボードに並べる / Inside the active artboard */
                frameBounds = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            }

            confirmedSettings = showGridDialog(doc, placedItems, frameBounds, initialSettings, maskGroups);
            if (!confirmedSettings) {
                /* キャンセルならマスクと配置した画像を取り除く / Remove the masks and the placed images on cancel */
                unwrapMasks(maskGroups);
                for (var j = placedItems.length - 1; j >= 0; j--) {
                    placedItems[j].remove();
                }
                app.redraw();
                return;
            }
        } finally {
            app.coordinateSystem = savedCoordinateSystem;
        }
        settingsStore.save(confirmedSettings);

        doc.selection = (maskGroups.length > 0) ? maskGroups : placedItems;
    }

    main();

})();
