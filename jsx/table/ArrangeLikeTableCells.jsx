#target illustrator
#targetengine "ArrangeLikeTableCellsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ArrangeLikeTableCells
Name=ArrangeLikeTableCells
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.0
Release-Date=2026-10-20
Description-BEGIN
外枠と罫線で区切った表で、外枠の大きさ・罫線の間隔・線幅を整え、各テキストをセルの中央に字形の見た目で配置します。

In a table drawn with an outer frame and rules, adjusts the frame size, the rule spacing and the stroke weights, then centers each text in its cell based on its glyph outlines.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

外枠と罫線で区切った表で、外枠の大きさ・罫線の間隔・線幅を整え、各テキストをセルの中央に字形の見た目で配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeLikeTableCells.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n22aa0e767930

### Overview

In a table drawn with an outer frame and rules, adjusts the frame size, the rule spacing and the stroke weights, then centers each text in its cell based on its glyph outlines.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeLikeTableCells.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArrangeLikeTableCells";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeLikeTableCells.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeLikeTableCells.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n22aa0e767930"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［縦罫の間隔を調整］［横罫の間隔を調整］の初期値 / Initial value of "Adjust vertical / horizontal rule spacing" */
    var ADJUST_RULE_SPACING = true;

    /*
        ［間隔］の初期値 / Initial value of "Spacing"
        "even"    : 罫線を等間隔に並べる
                    Space the rules evenly.
        "content" : セルの大きさをテキストに合わせ、テキストの両側の余白をそろえる
                    Size each cell to its texts, with equal space around them.
    */
    var DISTRIBUTE_MODE = "even";

    /*
        ［長さを外枠に合わせる］の初期値 / Initial value of "Match rule lengths to frame"
        true  : 縦罫の上下、横罫の左右を外枠に合わせる（2点の罫線のみ）
                Snap rule endpoints to the outer frame (two-point rules only).
        false : 罫線の長さは変えず、位置だけ平均化する
                Keep rule lengths; redistribute positions only.
    */
    var MATCH_RULE_LENGTHS = true;

    /* ［外枠を長方形に変換］の初期値 / Initial value of "Convert frame to rectangle" */
    var CONVERT_FRAME_TO_RECT = false;

    /* ［角を丸くする］の初期値と、半径の初期値（pt） / Initial value of "Round corners" and the initial radius (pt) */
    var ROUND_CORNERS = false;
    var CORNER_RADIUS_PT = 8;

    /* ［線幅を揃える］の初期値（線幅の初期値は、外枠・罫線でもっとも多い線幅） / Initial value of "Match stroke weights" (the weight starts at the most common one among the frame and rules) */
    var MATCH_STROKE_WIDTHS = false;

    /* ［図形を作成］の初期値 / Initial value of "Create cell shapes" */
    var CREATE_CELL_SHAPES = false;

    /*
        セルの図形の塗り方の初期値 / Initial fill of the cell shapes
        "uniform" : すべてのセルを FILL_UNIFORM_TINT の濃さで塗る
                    Fill every cell with the FILL_UNIFORM_TINT tint.
        "random"  : セルごとに FILL_TINT_MIN〜FILL_TINT_MAX のランダムな濃さで塗る
                    Fill each cell with a random tint from FILL_TINT_MIN to FILL_TINT_MAX.
    */
    var CELL_FILL_MODE = "uniform";

    /* ［すべてのセルを…で塗る］の濃さと、ランダムに選ぶ濃さの範囲（K%） / Tint for "Fill all cells" and the range of random tints (K%) */
    var FILL_UNIFORM_TINT = 15;
    var FILL_TINT_MIN = 10;
    var FILL_TINT_MAX = 60;

    /* ［グループ化］の初期値（セルの図形を1つのグループにまとめる） / Initial value of "Group" (group the cell shapes) */
    var GROUP_CELL_FILLS = true;

    /* ［セルの中央に配置］の初期値 / Initial value of "Center in cells" */
    var CENTER_IN_CELLS = true;

    /* ［中央揃えにする］の初期値 / Initial value of "Align center" */
    var JUSTIFY_CENTER = true;

    /* 縦・横の辺・罫線とみなす傾きの許容値（pt） / Tolerance for treating a segment or rule as vertical or horizontal (pt) */
    var TOLERANCE = 0.5;

    /* 位置の比較に挟む許容値（pt） / Tolerance for comparing positions (pt) */
    var EPSILON = 0.001;

    /* 縦・横の辺とみなす傾き（長さに対するずれの割合。0.02 で約1°） / Slope still treated as vertical or horizontal (offset per length; 0.02 is about 1 degree) */
    var SLOPE_TOLERANCE = 0.02;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログの余白・字下げ・欄の幅 / Gaps, indents and field widths in the dialog */
    var SEPARATE_OPTION_GAP       = 6;   /* 別の機能の行の上に足す余白 / extra space above a row for a separate feature */
    var FRAME_OPTIONS_TOP_GAP     = 5;   /* ［外枠を長方形に変換］の上に足す余白 / extra space above "Convert frame to rectangle" */
    var MATCH_LENGTHS_TOP_GAP     = 3;   /* ［長さを外枠に合わせる］の上に足す余白 / extra space above "Match rule lengths to frame" */
    var SUB_OPTION_INDENT         = 12;  /* チェックボックスの下に置く項目（塗り方のラジオボタン・線幅の欄）の字下げ / indent of items placed under a checkbox (fill radios, stroke weight field) */
    var CORNER_RADIUS_FIELD_CHARS = 4;   /* ［角を丸くする］の欄の幅（文字数） / width of the "Round corners" field (characters) */

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
            title: { ja: "表のように整える", en: "Arrange Like a Table" }
        },
        panel: {
            frame: { ja: "外枠", en: "Frame" },
            rules: { ja: "罫線", en: "Rules" },
            text: { ja: "テキスト", en: "Text" },
            cells: { ja: "セル", en: "Cells" }
        },
        checkbox: {
            adjustVerticalSpacing: { ja: "縦罫の間隔を調整", en: "Adjust vertical rule spacing" },
            adjustHorizontalSpacing: { ja: "横罫の間隔を調整", en: "Adjust horizontal rule spacing" },
            matchRuleLengths: { ja: "長さを外枠に合わせる", en: "Match rule lengths to frame" },
            convertFrame: { ja: "外枠を長方形に変換", en: "Convert frame to rectangle" },
            roundCorners: { ja: "角を丸くする", en: "Round corners" },
            matchStrokeWidths: { ja: "線幅を揃える", en: "Match stroke weights" },
            createCellShapes: { ja: "図形を作成", en: "Create cell shapes" },
            groupFills: { ja: "グループ化", en: "Group" },
            centerInCells: { ja: "セルの中央に配置", en: "Center in cells" },
            justifyCenter: { ja: "中央揃えにする", en: "Align center" }
        },
        fieldLabel: {
            frameWidth: { ja: "幅", en: "Width" },
            frameHeight: { ja: "高さ", en: "Height" },
            spacing: { ja: "間隔", en: "Spacing" },
            fill: { ja: "塗り", en: "Fill" }
        },
        radio: {
            even: { ja: "均等", en: "Even" },
            content: { ja: "テキストに合わせる", en: "Fit to text" },
            uniformFill: { ja: "同じ濃さ（K{tint}）", en: "Same tint (K{tint})" },
            randomFill: { ja: "ランダムな濃さ", en: "Random tints" }
        },
        tooltip: {
            frameSize: {
                ja: "外枠の大きさを変えます。左上を基準に、外枠と罫線を伸縮します（線幅とテキストの大きさは変えません）。",
                en: "Changes the size of the frame. The frame and rules are scaled from the top-left corner (stroke widths and text sizes are kept)."
            },
            adjustVerticalSpacing: {
                ja: "外枠の中にある縦罫の位置を、［間隔］の指定どおりに調整します。Option キーを押しながらクリックすると横罫もまとめて切り替えます。",
                en: "Adjusts the positions of the vertical rules inside the frame as set in Spacing. Option-click to toggle the horizontal rules too."
            },
            adjustHorizontalSpacing: {
                ja: "外枠の中にある横罫の位置を、［間隔］の指定どおりに調整します。Option キーを押しながらクリックすると縦罫もまとめて切り替えます。",
                en: "Adjusts the positions of the horizontal rules inside the frame as set in Spacing. Option-click to toggle the vertical rules too."
            },
            even: {
                ja: "罫線を等間隔に並べます。",
                en: "Spaces the rules evenly."
            },
            content: {
                ja: "セルの幅（高さ）をテキストに合わせ、テキストの両側の余白をそろえます。",
                en: "Sizes each column (row) to its texts, with equal space on both sides."
            },
            matchRuleLengths: {
                ja: "間隔を調整した罫線の長さを、外枠の辺に合わせます。",
                en: "Matches the length of the adjusted rules to the sides of the frame."
            },
            roundCorners: {
                ja: "外枠の長方形に、指定した半径で［角を丸くする］効果を適用します。外枠が長方形のとき（［外枠を長方形に変換］を含む）に使えます。セルの塗りも外枠の角に合わせて丸めます。",
                en: "Applies the Round Corners effect with the given radius to the frame rectangle. Available when the frame is a rectangle (including \"Convert frame to rectangle\"). Cell fills are rounded at the frame corners too."
            },
            matchStrokeWidths: {
                ja: "外枠と罫線の線幅を、指定した値にそろえます（線の無いパスは変えません）。",
                en: "Sets the stroke weight of the frame and rules to the given value (paths without a stroke are left as they are)."
            },
            groupFills: {
                ja: "セルの図形を1つのグループにまとめます。オフにすると、図形を1つずつ最背面に置きます。",
                en: "Groups the cell shapes. When off, the shapes are placed individually at the back."
            },
            createCellShapes: {
                ja: "各セルに、グレーで塗った長方形を作り、最背面に置きます。",
                en: "Adds a rectangle filled with gray in each cell, sent to the back."
            },
            uniformFill: {
                ja: "すべてのセルを K{tint} のグレーで塗ります。",
                en: "Fills every cell with a K{tint} gray."
            },
            randomFill: {
                ja: "セルごとに、K{min}〜K{max} のランダムな濃さのグレーで塗ります。プレビューを更新しても濃さは変わりません。",
                en: "Fills each cell with a random gray tint from K{min} to K{max}. The tints stay the same while the preview updates."
            },
            convertFrame: {
                ja: "外枠を作っている罫線4本を、同じ位置の長方形1つに置き換えます。線の設定は上辺の罫線から引き継ぎます。長方形を使わず罫線で囲んだ枠のときだけ使えます。",
                en: "Replaces the four rules forming the frame with a single rectangle in the same place, taking the stroke from the top rule. Available only when the frame is drawn with rules."
            },
            centerInCells: {
                ja: "テキストを、枠・罫線で区切られたセルの中央へ字形の見た目で移動します。",
                en: "Moves each text to the center of its cell, based on its glyph outlines."
            },
            justifyCenter: {
                ja: "段落の行揃えを中央揃えにします。",
                en: "Sets the paragraph alignment to center."
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
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noText: {
                ja: "テキストと、セルの枠・罫線をいっしょに選択してください。",
                en: "Select the texts together with the cell frame and rules."
            },
            noRule: {
                ja: "セルの枠・罫線が選択されていません。",
                en: "No cell frame or rules are selected."
            },
            skipped: {
                ja: "{count} 個のテキストは、囲むセルが見つからなかったため移動しませんでした。",
                en: "{count} text(s) were not moved because no enclosing cell was found."
            }
        }
    };

    // =========================================
    // 選択 / Selection
    // =========================================

    /**
     * 処理対象の選択オブジェクトを返す
     * @returns {PageItem[]|null} 選択オブジェクト（前面→背面の順）。ドキュメントがない・未選択・文字の選択中は null
     */
    function getValidSelection() {
        if (!app.documents.length) return null;
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る
           A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || !selectedItems.length || selectedItems.typename === "TextRange") return null;
        return selectedItems;
    }

    /**
     * グループ・複合パスの中までたどって、テキストとパスを振り分ける
     * @param {PageItem} item - 対象のオブジェクト
     * @param {TextFrame[]} textFrames - テキストの追加先
     * @param {PathItem[]} pathItems - パスの追加先
     * @returns {void}
     */
    function collectItems(item, textFrames, pathItems) {
        var i;
        switch (item.typename) {
            case "TextFrame":
                textFrames.push(item);
                break;
            case "PathItem":
                if (!item.guides) pathItems.push(item);
                break;
            case "CompoundPathItem":
                for (i = 0; i < item.pathItems.length; i++) pathItems.push(item.pathItems[i]);
                break;
            case "GroupItem":
                for (i = 0; i < item.pageItems.length; i++) collectItems(item.pageItems[i], textFrames, pathItems);
                break;
        }
    }

    // =========================================
    // セルの辺 / Cell edges
    // =========================================

    /**
     * パスのアンカー間の線分から、垂直・水平の辺を拾う（SLOPE_TOLERANCE までの傾きも含める）
     * @param {PathItem[]} pathItems - 枠・罫線のパス
     * @returns {{vertical: Object[], horizontal: Object[]}} 辺ごとの {pathItem, startAlong, startPos, endAlong, endPos}。
     *     along は辺に沿う軸（垂直なら Y）、pos は直交する軸（垂直なら X）の座標で、startAlong ≦ endAlong
     */
    function collectEdges(pathItems) {
        var edges = { vertical: [], horizontal: [] };
        for (var i = 0; i < pathItems.length; i++) {
            var points = pathItems[i].pathPoints;
            var pointCount = points.length;
            var segmentCount = pathItems[i].closed ? pointCount : pointCount - 1;
            for (var j = 0; j < segmentCount; j++) {
                var startPoint = points[j].anchor;
                var endPoint = points[(j + 1) % pointCount].anchor;
                var dx = Math.abs(startPoint[0] - endPoint[0]);
                var dy = Math.abs(startPoint[1] - endPoint[1]);
                if (dy > TOLERANCE && dx <= Math.max(TOLERANCE, dy * SLOPE_TOLERANCE)) {
                    edges.vertical.push(createEdge(pathItems[i], startPoint[1], startPoint[0], endPoint[1], endPoint[0]));
                } else if (dx > TOLERANCE && dy <= Math.max(TOLERANCE, dx * SLOPE_TOLERANCE)) {
                    edges.horizontal.push(createEdge(pathItems[i], startPoint[0], startPoint[1], endPoint[0], endPoint[1]));
                }
            }
        }
        return edges;
    }

    /**
     * 線分を、沿う軸の小さい端から大きい端への辺の情報にする
     * @param {PathItem} pathItem - 線分を含むパス
     * @param {number} alongA - 一方の端の、辺に沿う軸の座標
     * @param {number} posA - 一方の端の、直交する軸の座標
     * @param {number} alongB - もう一方の端の、辺に沿う軸の座標
     * @param {number} posB - もう一方の端の、直交する軸の座標
     * @returns {{pathItem: PathItem, startAlong: number, startPos: number, endAlong: number, endPos: number}} 辺の情報
     */
    function createEdge(pathItem, alongA, posA, alongB, posB) {
        if (alongA <= alongB) return { pathItem: pathItem, startAlong: alongA, startPos: posA, endAlong: alongB, endPos: posB };
        return { pathItem: pathItem, startAlong: alongB, startPos: posB, endAlong: alongA, endPos: posA };
    }

    /**
     * 辺に沿う軸の座標での、辺の位置を返す（傾いた辺は両端から補間する）
     * @param {Object} edge - createEdge() の戻り値
     * @param {number} along - 辺に沿う軸の座標
     * @returns {number} 直交する軸の座標
     */
    function getEdgePosAt(edge, along) {
        var span = edge.endAlong - edge.startAlong;
        if (span <= 0) return edge.startPos;
        var ratio = Math.max(0, Math.min(1, (along - edge.startAlong) / span));
        return edge.startPos + (edge.endPos - edge.startPos) * ratio;
    }

    /**
     * 基準位置の両側で、いちばん近い辺を返す
     * @param {Object[]} edges - collectEdges の vertical または horizontal
     * @param {number} pos - 辺と直交する軸での基準位置
     * @param {number} along - 辺に沿う軸での基準位置（この位置を覆う辺だけを使う）
     * @returns {{lower: number, upper: number, lowerItem: PathItem, upperItem: PathItem}|null} 両側の辺の位置とパス。片側でも無ければ null
     */
    function findNearestEdges(edges, pos, along) {
        var nearest = { lower: null, upper: null, lowerItem: null, upperItem: null };
        for (var i = 0; i < edges.length; i++) {
            var edge = edges[i];
            if (along < edge.startAlong - TOLERANCE || along > edge.endAlong + TOLERANCE) continue;
            var edgePos = getEdgePosAt(edge, along);
            if (edgePos < pos - EPSILON && (nearest.lower === null || edgePos > nearest.lower)) {
                nearest.lower = edgePos;
                nearest.lowerItem = edge.pathItem;
            }
            if (edgePos > pos + EPSILON && (nearest.upper === null || edgePos < nearest.upper)) {
                nearest.upper = edgePos;
                nearest.upperItem = edge.pathItem;
            }
        }
        if (nearest.lower === null || nearest.upper === null) return null;
        return nearest;
    }

    /**
     * 罫線を動かす前の辺の位置を、動かしたあとの位置に置き換える
     * @param {Object[]} positionMap - {pathItem, from: 前の位置, to: 後の位置} の配列
     * @param {number} pos - 動かす前の辺の位置
     * @param {PathItem} pathItem - その辺を含むパス
     * @param {{pathItems: PathItem[], origin: number, ratio: number}} axisScale - 外枠の伸縮（この軸の基準と倍率、伸縮したパス）
     * @returns {number} 動かしたあとの位置。並べ直した罫線ならその位置、伸縮したパスなら伸縮後の位置、どちらでもなければそのまま
     */
    function mapPosition(positionMap, pos, pathItem, axisScale) {
        var i;
        for (i = 0; i < positionMap.length; i++) {
            /* DOM の参照は === で同じものと判定できる / DOM references compare equal with === */
            if (positionMap[i].pathItem === pathItem) return positionMap[i].to;
        }
        for (i = 0; i < axisScale.pathItems.length; i++) {
            if (axisScale.pathItems[i] === pathItem) return axisScale.origin + (pos - axisScale.origin) * axisScale.ratio;
        }
        return pos;
    }

    // =========================================
    // 外枠と罫線の判定 / Detecting the frame and rules
    // =========================================

    /**
     * もっとも面積の大きい閉じたパスを外枠とみなす
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @returns {PathItem|null} 外枠のパス。見つからなければ null
     */
    function findOuterRectangle(candidatePathItems) {
        var outerRectangleCandidate = null;
        var maxArea = 0;

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];
            if (!pathItem.closed) continue;

            var geometricBounds = pathItem.geometricBounds;
            var pathWidth = geometricBounds[2] - geometricBounds[0];
            var pathHeight = geometricBounds[1] - geometricBounds[3];
            if (pathWidth <= 0 || pathHeight <= 0) continue;

            var pathArea = pathWidth * pathHeight;
            if (pathArea > maxArea) {
                maxArea = pathArea;
                outerRectangleCandidate = pathItem;
            }
        }
        return outerRectangleCandidate;
    }

    /**
     * 外枠の座標と寸法をまとめて返す
     * @param {PathItem} outerRect - 外枠のパス
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 外枠の範囲
     */
    function getRectBounds(outerRect) {
        var outerBounds = outerRect.geometricBounds;
        return {
            left: outerBounds[0],
            top: outerBounds[1],
            right: outerBounds[2],
            bottom: outerBounds[3],
            width: outerBounds[2] - outerBounds[0],
            height: outerBounds[1] - outerBounds[3]
        };
    }

    /**
     * 外枠の中にある縦罫・横罫を振り分け、縦罫は左から右、横罫は上から下に並べる
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @param {PathItem|null} outerRect - 外枠のパス（対象から除く）。罫線で囲んだ枠なら null
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {{vertical: PathItem[], horizontal: PathItem[]}} 縦罫と横罫
     */
    function classifyRules(candidatePathItems, outerRect, rectBounds) {
        var verticalRules = [];
        var horizontalRules = [];

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];
            if (pathItem === outerRect) continue;

            if (isRuleInRect(pathItem, rectBounds, true)) {
                verticalRules.push(pathItem);
            } else if (isRuleInRect(pathItem, rectBounds, false)) {
                horizontalRules.push(pathItem);
            }
        }

        /* 縦罫は左から右へ、横罫は上から下へソート / Sort vertical L→R, horizontal T→B */
        verticalRules.sort(function (firstRule, secondRule) {
            return getCenterX(firstRule) - getCenterX(secondRule);
        });
        horizontalRules.sort(function (firstRule, secondRule) {
            return getCenterY(secondRule) - getCenterY(firstRule);
        });

        return { vertical: verticalRules, horizontal: horizontalRules };
    }

    /**
     * 外枠の中にある縦罫（または横罫）かどうかを判定する
     * 縦罫は「長さ＝高さ、並ぶ方向＝X」、横罫は「長さ＝幅、並ぶ方向＝Y」として同じ条件で見る
     * @param {PathItem} pathItem - 判定するパス
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} isVertical - true なら縦罫、false なら横罫として判定する
     * @returns {boolean} 罫線とみなせるとき true
     */
    function isRuleInRect(pathItem, rectBounds, isVertical) {
        if (pathItem.closed) return false;

        var geometricBounds = pathItem.geometricBounds;
        var pathWidth = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var pathHeight = Math.abs(geometricBounds[1] - geometricBounds[3]);
        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;

        /* 罫線の幅（傾いた罫線ならその振れ幅）ごと外枠の内側にあるものだけを罫線とする。外枠の辺に重なる線は辺とみなす
           Only lines lying entirely inside the frame across their thickness (or tilt) are rules; lines on a frame edge count as the edge */
        if (isVertical) {
            return hasRuleShape(pathHeight, pathWidth, rectBounds.height) &&
                isInsideRange(geometricBounds[0], rectBounds.left, rectBounds.right) &&
                isInsideRange(geometricBounds[2], rectBounds.left, rectBounds.right) &&
                isNearRange(centerY, rectBounds.bottom, rectBounds.top);
        }
        return hasRuleShape(pathWidth, pathHeight, rectBounds.width) &&
            isInsideRange(geometricBounds[3], rectBounds.bottom, rectBounds.top) &&
            isInsideRange(geometricBounds[1], rectBounds.bottom, rectBounds.top) &&
            isNearRange(centerX, rectBounds.left, rectBounds.right);
    }

    /**
     * 細長い形で、外枠の辺の 20% 以上の長さがあるかどうか
     * @param {number} ruleLength - 罫線の方向の長さ
     * @param {number} ruleThickness - 罫線と直交する方向の幅
     * @param {number} outerLength - 同じ方向の外枠の長さ
     * @returns {boolean} 罫線の形とみなせるとき true
     */
    function hasRuleShape(ruleLength, ruleThickness, outerLength) {
        if (ruleLength <= 0) return false;
        if (!(ruleThickness <= TOLERANCE || ruleLength > ruleThickness * 5)) return false;
        return !(ruleLength < outerLength * 0.2);
    }

    /**
     * 値が範囲の内側に、許容値より離れて収まっているかどうか（両端ちょうどは外）
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 内側なら true
     */
    function isInsideRange(value, rangeMin, rangeMax) {
        return value > rangeMin + TOLERANCE && value < rangeMax - TOLERANCE;
    }

    /**
     * 値が範囲に、許容値の はみ出しまで含めて収まっているかどうか
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 収まっていれば true
     */
    function isNearRange(value, rangeMin, rangeMax) {
        return value >= rangeMin - TOLERANCE && value <= rangeMax + TOLERANCE;
    }

    /**
     * パスの外接矩形の中心 X を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の X 座標
     */
    function getCenterX(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[0] + geometricBounds[2]) / 2;
    }

    /**
     * パスの外接矩形の中心 Y を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の Y 座標
     */
    function getCenterY(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[1] + geometricBounds[3]) / 2;
    }

    // =========================================
    // 罫線の再配置 / Redistributing rules
    // =========================================

    /**
     * 外枠と、その中の縦罫・横罫を見つける。
     * 外枠は、いちばん大きい閉じたパス。閉じたパスが無いときだけ、罫線4本で囲んだ枠を使う
     * （長方形を領域として置き、罫線をその外まで伸ばした表でも、長方形を外枠にする）
     * @param {PathItem[]} pathItems - 選択中のパス
     * @returns {{rectBounds: Object, vertical: PathItem[], horizontal: PathItem[], frameLines: Object|null, outerRect: PathItem|null}|null} 外枠の範囲と罫線（罫線で囲んだ枠なら、その4辺の罫線も。閉じたパスの外枠ならそのパスも）。外枠が無ければ null
     */
    function findRules(pathItems) {
        var outerRect = findOuterRectangle(pathItems);
        var rectBounds = outerRect ? getRectBounds(outerRect) : null;
        var lineFrame = outerRect ? null : findLineFrame(pathItems);
        if (lineFrame) {
            /* 外側の罫線は外枠の辺になり、classifyRules() で罫線から外れる / The outermost rules become the frame and are left out by classifyRules() */
            outerRect = null;
            rectBounds = lineFrame.bounds;
        }
        if (!rectBounds) return null;

        var rules = classifyRules(pathItems, outerRect, rectBounds);
        return {
            rectBounds: rectBounds,
            vertical: rules.vertical,
            horizontal: rules.horizontal,
            frameLines: lineFrame ? { vertical: lineFrame.vertical, horizontal: lineFrame.horizontal } : null,
            outerRect: outerRect
        };
    }

    /**
     * 長方形を使わず、罫線4本で囲んだ枠を見つける。
     * 左右は外側の縦罫の位置、上下は外側の横罫の位置で決める（罫線の端がはみ出していても枠は広がらない）
     * @param {PathItem[]} pathItems - 選択中のパス
     * @returns {{bounds: Object, vertical: PathItem[], horizontal: PathItem[]}|null} 枠の範囲と、枠になる縦罫・横罫。縦・横とも2本以上なければ null
     */
    function findLineFrame(pathItems) {
        var verticalLines = [];
        var horizontalLines = [];
        var i;
        for (i = 0; i < pathItems.length; i++) {
            if (pathItems[i].closed) continue;
            var geometricBounds = pathItems[i].geometricBounds;
            var pathWidth = geometricBounds[2] - geometricBounds[0];
            var pathHeight = geometricBounds[1] - geometricBounds[3];
            /* セルの辺と同じく、SLOPE_TOLERANCE までの傾きは縦・横とみなす / Allow the same slope as cell edges */
            if (pathHeight > TOLERANCE && pathWidth <= Math.max(TOLERANCE, pathHeight * SLOPE_TOLERANCE)) verticalLines.push(pathItems[i]);
            else if (pathWidth > TOLERANCE && pathHeight <= Math.max(TOLERANCE, pathWidth * SLOPE_TOLERANCE)) horizontalLines.push(pathItems[i]);
        }
        if (verticalLines.length < 2 || horizontalLines.length < 2) return null;

        var left = Infinity;
        var right = -Infinity;
        var top = -Infinity;
        var bottom = Infinity;
        for (i = 0; i < verticalLines.length; i++) {
            left = Math.min(left, getCenterX(verticalLines[i]));
            right = Math.max(right, getCenterX(verticalLines[i]));
        }
        for (i = 0; i < horizontalLines.length; i++) {
            top = Math.max(top, getCenterY(horizontalLines[i]));
            bottom = Math.min(bottom, getCenterY(horizontalLines[i]));
        }
        if (right - left <= TOLERANCE || top - bottom <= TOLERANCE) return null;

        var frameVertical = [];
        var frameHorizontal = [];
        for (i = 0; i < verticalLines.length; i++) {
            var centerX = getCenterX(verticalLines[i]);
            if (Math.abs(centerX - left) <= TOLERANCE || Math.abs(centerX - right) <= TOLERANCE) frameVertical.push(verticalLines[i]);
        }
        for (i = 0; i < horizontalLines.length; i++) {
            var centerY = getCenterY(horizontalLines[i]);
            if (Math.abs(centerY - top) <= TOLERANCE || Math.abs(centerY - bottom) <= TOLERANCE) frameHorizontal.push(horizontalLines[i]);
        }

        return {
            bounds: { left: left, top: top, right: right, bottom: bottom, width: right - left, height: top - bottom },
            vertical: frameVertical,
            horizontal: frameHorizontal
        };
    }

    /**
     * 罫線で囲んだ枠の4辺を、枠の角から角までの長さにそろえる（2点の罫線のみ）
     * @param {Object} frameLines - findRules() の frameLines
     * @param {Object} rectBounds - 枠の範囲
     * @returns {void}
     */
    function matchFrameLineLengths(frameLines, rectBounds) {
        var i;
        for (i = 0; i < frameLines.vertical.length; i++) {
            if (frameLines.vertical[i].pathPoints.length !== 2) continue;
            setTwoPointVerticalLine(frameLines.vertical[i], getCenterX(frameLines.vertical[i]), rectBounds.top, rectBounds.bottom);
        }
        for (i = 0; i < frameLines.horizontal.length; i++) {
            if (frameLines.horizontal[i].pathPoints.length !== 2) continue;
            setTwoPointHorizontalLine(frameLines.horizontal[i], getCenterY(frameLines.horizontal[i]), rectBounds.left, rectBounds.right);
        }
    }

    /**
     * 縦罫・横罫を、設定に従って並べ直す
     * @param {Object|null} rules - findRules() の戻り値（columnSizes・rowSizes を足したもの）
     * @param {{adjustVerticalSpacing: boolean, adjustHorizontalSpacing: boolean, distributeMode: string, matchRuleLengths: boolean}} settings - 設定
     * @returns {{vertical: Object[], horizontal: Object[]}} 罫線ごとの {from: 前の位置, to: 後の位置}
     */
    function adjustRuleSpacing(rules, settings) {
        var positionMap = { vertical: [], horizontal: [] };
        if (!rules) return positionMap;

        var fitsContent = settings.distributeMode === "content";
        if (settings.adjustVerticalSpacing) {
            var columnOffsets = computeRuleOffsets(rules.rectBounds.width, rules.vertical.length, fitsContent ? rules.columnSizes : null);
            positionMap.vertical = redistributeVerticalRules(rules.vertical, rules.rectBounds, settings.matchRuleLengths, columnOffsets);
        }
        if (settings.adjustHorizontalSpacing) {
            var rowOffsets = computeRuleOffsets(rules.rectBounds.height, rules.horizontal.length, fitsContent ? rules.rowSizes : null);
            positionMap.horizontal = redistributeHorizontalRules(rules.horizontal, rules.rectBounds, settings.matchRuleLengths, rowOffsets);
        }
        if (rules.frameLines && settings.matchRuleLengths && (settings.adjustVerticalSpacing || settings.adjustHorizontalSpacing)) {
            matchFrameLineLengths(rules.frameLines, rules.rectBounds);
        }
        return positionMap;
    }

    /**
     * 罫線を置く位置を、外枠の辺からの距離で返す。
     * 区画ごとに「中身の大きさ＋共通の余白」を割り当てる（中身が無ければ等間隔）
     * @param {number} totalLength - 外枠の幅（高さ）
     * @param {number} ruleCount - 罫線の数
     * @param {number[]|null} contentSizes - 区画ごとの中身の大きさ（罫線の数＋1 個）。null なら等間隔
     * @returns {number[]} 罫線ごとの、外枠の左辺（上辺）からの距離
     */
    function computeRuleOffsets(totalLength, ruleCount, contentSizes) {
        var sectionCount = ruleCount + 1;
        var sectionSizes = [];
        var contentTotal = 0;
        var i;
        for (i = 0; i < sectionCount; i++) {
            var sectionSize = contentSizes ? (contentSizes[i] || 0) : 0;
            sectionSizes.push(sectionSize);
            contentTotal += sectionSize;
        }

        var padding = (totalLength - contentTotal) / sectionCount;
        var scale = 1;
        /* 中身が外枠に収まらないときは、余白なしで中身の比率に縮める / If the content does not fit, drop the padding and scale by the content ratio */
        if (padding < 0) {
            padding = 0;
            scale = totalLength / contentTotal;
        }

        var offsets = [];
        var offset = 0;
        for (i = 0; i < ruleCount; i++) {
            offset += sectionSizes[i] * scale + padding;
            offsets.push(offset);
        }
        return offsets;
    }

    /**
     * 罫線で区切られた区画ごとに、中にあるテキストの最大の幅（高さ）を測る
     * @param {Object[]} textCells - main() で集めたテキストの情報
     * @param {Object} rules - findRules() の戻り値
     * @param {boolean} isVertical - true なら縦罫で区切った列の幅、false なら横罫で区切った行の高さ
     * @returns {number[]} 区画ごとの大きさ（左から右、上から下。罫線の数＋1 個）
     */
    function measureContentSizes(textCells, rules, isVertical) {
        var sortedRules = isVertical ? rules.vertical : rules.horizontal;
        var rulePositions = [];
        var sizes = [0];
        var i;
        for (i = 0; i < sortedRules.length; i++) {
            rulePositions.push(isVertical ? getCenterX(sortedRules[i]) : getCenterY(sortedRules[i]));
            sizes.push(0);
        }

        var rectBounds = rules.rectBounds;
        for (i = 0; i < textCells.length; i++) {
            var glyphCenter = textCells[i].glyphCenter;
            if (!glyphCenter) continue;
            /* 外枠の外にあるテキストは数えない / Ignore texts outside the frame */
            if (!isNearRange(glyphCenter[0], rectBounds.left, rectBounds.right) ||
                !isNearRange(glyphCenter[1], rectBounds.bottom, rectBounds.top)) continue;

            var center = isVertical ? glyphCenter[0] : glyphCenter[1];
            var size = isVertical ? textCells[i].glyphSize[0] : textCells[i].glyphSize[1];
            /* 縦罫は左から右、横罫は上から下（Y は下ほど小さい）に並んでいる / Vertical rules run left to right, horizontal top to bottom (Y decreases downward) */
            var sectionIndex = 0;
            for (var j = 0; j < rulePositions.length; j++) {
                if (isVertical ? rulePositions[j] < center : rulePositions[j] > center) sectionIndex++;
            }
            if (size > sizes[sectionIndex]) sizes[sectionIndex] = size;
        }
        return sizes;
    }

    /**
     * 縦罫を指定の位置に並べ直す
     * @param {PathItem[]} verticalRules - 左から右に並んだ縦罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} matchRuleLengths - 2点の罫線の上下を外枠に合わせるかどうか
     * @param {number[]} ruleOffsets - 罫線ごとの、外枠の左辺からの距離
     * @returns {Object[]} 罫線ごとの {from: 前の X, to: 後の X}
     */
    function redistributeVerticalRules(verticalRules, rectBounds, matchRuleLengths, ruleOffsets) {
        var positionMap = [];
        var ruleCount = verticalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var verticalRule = verticalRules[i];
            var currentX = getCenterX(verticalRule);
            var targetX = rectBounds.left + ruleOffsets[i];
            positionMap.push({ pathItem: verticalRule, from: currentX, to: targetX });

            if (matchRuleLengths && verticalRule.pathPoints.length === 2) {
                setTwoPointVerticalLine(verticalRule, targetX, rectBounds.top, rectBounds.bottom);
            } else {
                translatePath(verticalRule, targetX - currentX, 0);
            }
        }
        return positionMap;
    }

    /**
     * 横罫を指定の位置に並べ直す
     * @param {PathItem[]} horizontalRules - 上から下に並んだ横罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} matchRuleLengths - 2点の罫線の左右を外枠に合わせるかどうか
     * @param {number[]} ruleOffsets - 罫線ごとの、外枠の上辺からの距離
     * @returns {Object[]} 罫線ごとの {from: 前の Y, to: 後の Y}
     */
    function redistributeHorizontalRules(horizontalRules, rectBounds, matchRuleLengths, ruleOffsets) {
        var positionMap = [];
        var ruleCount = horizontalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var horizontalRule = horizontalRules[i];
            var currentY = getCenterY(horizontalRule);
            var targetY = rectBounds.top - ruleOffsets[i];
            positionMap.push({ pathItem: horizontalRule, from: currentY, to: targetY });

            if (matchRuleLengths && horizontalRule.pathPoints.length === 2) {
                setTwoPointHorizontalLine(horizontalRule, targetY, rectBounds.left, rectBounds.right);
            } else {
                translatePath(horizontalRule, 0, targetY - currentY);
            }
        }
        return positionMap;
    }

    /**
     * パスのすべてのアンカーポイントとハンドルを平行移動する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function translatePath(pathItem, deltaX, deltaY) {
        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            movePathPoint(pathItem.pathPoints[i], deltaX, deltaY);
        }
    }

    /**
     * 2点の縦罫を、指定の X 位置で上端から下端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} x - 配置する X 座標
     * @param {number} top - 上端の Y 座標
     * @param {number} bottom - 下端の Y 座標
     * @returns {void}
     */
    function setTwoPointVerticalLine(pathItem, x, top, bottom) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnTop = firstPoint.anchor[1] >= secondPoint.anchor[1];

        movePointTo(isFirstOnTop ? firstPoint : secondPoint, x, top);
        movePointTo(isFirstOnTop ? secondPoint : firstPoint, x, bottom);
    }

    /**
     * 2点の横罫を、指定の Y 位置で左端から右端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} y - 配置する Y 座標
     * @param {number} left - 左端の X 座標
     * @param {number} right - 右端の X 座標
     * @returns {void}
     */
    function setTwoPointHorizontalLine(pathItem, y, left, right) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnLeft = firstPoint.anchor[0] <= secondPoint.anchor[0];

        movePointTo(isFirstOnLeft ? firstPoint : secondPoint, left, y);
        movePointTo(isFirstOnLeft ? secondPoint : firstPoint, right, y);
    }

    /**
     * アンカーポイントを指定の座標へ移動する（ハンドルも同じ量だけ動かす）
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} targetX - 移動先の X 座標
     * @param {number} targetY - 移動先の Y 座標
     * @returns {void}
     */
    function movePointTo(pathPoint, targetX, targetY) {
        movePathPoint(pathPoint, targetX - pathPoint.anchor[0], targetY - pathPoint.anchor[1]);
    }

    /**
     * アンカーポイントと両ハンドルを同じ量だけ動かす
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function movePathPoint(pathPoint, deltaX, deltaY) {
        pathPoint.anchor = [pathPoint.anchor[0] + deltaX, pathPoint.anchor[1] + deltaY];
        pathPoint.leftDirection = [pathPoint.leftDirection[0] + deltaX, pathPoint.leftDirection[1] + deltaY];
        pathPoint.rightDirection = [pathPoint.rightDirection[0] + deltaX, pathPoint.rightDirection[1] + deltaY];
    }

    // =========================================
    // 外枠の大きさ / Frame size
    // =========================================

    /**
     * 外枠の大きさを変えるときに伸縮するパス（外枠・枠の罫線・中の罫線）を返す
     * @param {Object} rules - findRules() の戻り値
     * @returns {PathItem[]} 伸縮するパス
     */
    function getTablePathItems(rules) {
        var tablePaths = rules.outerRect ? [rules.outerRect] : [];
        if (rules.frameLines) tablePaths = tablePaths.concat(rules.frameLines.vertical, rules.frameLines.horizontal);
        return tablePaths.concat(rules.vertical, rules.horizontal);
    }

    /**
     * 線のあるパスのうち、もっとも多い線幅を返す
     * @param {PathItem[]} targetPathItems - 対象のパス
     * @returns {number} 線幅（pt）。線のあるパスが無ければ 1
     */
    function getMostCommonStrokeWidth(targetPathItems) {
        var widthCounts = {};
        var mostCommonWidth = 1;
        var maxCount = 0;
        for (var i = 0; i < targetPathItems.length; i++) {
            if (!targetPathItems[i].stroked) continue;
            var strokeWidth = targetPathItems[i].strokeWidth;
            var widthKey = String(Math.round(strokeWidth * 1000) / 1000);
            widthCounts[widthKey] = (widthCounts[widthKey] || 0) + 1;
            if (widthCounts[widthKey] > maxCount) {
                maxCount = widthCounts[widthKey];
                mostCommonWidth = strokeWidth;
            }
        }
        return mostCommonWidth;
    }

    /**
     * 線のあるパスの線幅をそろえる（線の無いパスはそのまま）
     * @param {PathItem[]} targetPathItems - 対象のパス
     * @param {number} strokeWidth - 線幅（pt）
     * @returns {void}
     */
    function setStrokeWidths(targetPathItems, strokeWidth) {
        for (var i = 0; i < targetPathItems.length; i++) {
            if (targetPathItems[i].stroked) targetPathItems[i].strokeWidth = strokeWidth;
        }
    }

    /**
     * パスの線端をそろえる
     * @param {PathItem[]} targetPathItems - 対象のパス
     * @param {StrokeCap} strokeCap - 線端
     * @returns {void}
     */
    function setStrokeCaps(targetPathItems, strokeCap) {
        for (var i = 0; i < targetPathItems.length; i++) targetPathItems[i].strokeCap = strokeCap;
    }

    /**
     * パスのアンカーポイントとハンドルを、基準点から縦横の倍率で伸縮する（線幅は変えない）
     * @param {PathItem[]} targetPathItems - 対象のパス
     * @param {number} originX - 基準点の X
     * @param {number} originY - 基準点の Y
     * @param {number} ratioX - 横の倍率
     * @param {number} ratioY - 縦の倍率
     * @returns {void}
     */
    function scalePathItems(targetPathItems, originX, originY, ratioX, ratioY) {
        /**
         * 座標を伸縮する
         * @param {number[]} point - 座標
         * @returns {number[]} 伸縮後の座標
         */
        function scalePoint(point) {
            return [originX + (point[0] - originX) * ratioX, originY + (point[1] - originY) * ratioY];
        }
        for (var i = 0; i < targetPathItems.length; i++) {
            var pathPoints = targetPathItems[i].pathPoints;
            for (var j = 0; j < pathPoints.length; j++) {
                var pathPoint = pathPoints[j];
                pathPoint.anchor = scalePoint(pathPoint.anchor);
                pathPoint.leftDirection = scalePoint(pathPoint.leftDirection);
                pathPoint.rightDirection = scalePoint(pathPoint.rightDirection);
            }
        }
    }

    // =========================================
    // 外枠の変換 / Converting the frame
    // =========================================

    /**
     * 罫線で囲んだ枠と同じ位置に長方形を作り、枠の罫線を隠す（プレビュー中は隠すだけで、確定時に削除する）
     * @param {Object} frameLines - findRules() の frameLines
     * @param {Object} rectBounds - 枠の範囲
     * @returns {PathItem} 作った長方形
     */
    function createFrameRectangle(frameLines, rectBounds) {
        /* 線の設定は上辺の罫線から引き継ぐ / Take the stroke from the top rule */
        var styleSource = frameLines.horizontal[0];
        for (var i = 1; i < frameLines.horizontal.length; i++) {
            if (getCenterY(frameLines.horizontal[i]) > getCenterY(styleSource)) styleSource = frameLines.horizontal[i];
        }

        var frameRect = styleSource.parent.pathItems.rectangle(rectBounds.top, rectBounds.left, rectBounds.width, rectBounds.height);
        frameRect.move(styleSource, ElementPlacement.PLACEBEFORE);
        frameRect.filled = false;
        frameRect.stroked = styleSource.stroked;
        if (styleSource.stroked) {
            frameRect.strokeColor = styleSource.strokeColor;
            frameRect.strokeWidth = styleSource.strokeWidth;
            frameRect.strokeDashes = styleSource.strokeDashes;
            frameRect.strokeDashOffset = styleSource.strokeDashOffset;
            frameRect.strokeOverprint = styleSource.strokeOverprint;
            /* 角が欠けないようにマイター結合にする / Use miter joins so the corners stay square */
            frameRect.strokeJoin = StrokeJoin.MITERENDJOIN;
        }
        frameRect.opacity = styleSource.opacity;

        setFrameLinesHidden(frameLines, true);
        return frameRect;
    }

    /**
     * 枠の罫線の表示・非表示を切り替える
     * @param {Object} frameLines - findRules() の frameLines
     * @param {boolean} isHidden - 隠すなら true
     * @returns {void}
     */
    function setFrameLinesHidden(frameLines, isHidden) {
        var lines = frameLines.vertical.concat(frameLines.horizontal);
        for (var i = 0; i < lines.length; i++) lines[i].hidden = isHidden;
    }

    /**
     * 枠の罫線を削除する（［OK］で変換を確定したとき）
     * @param {Object} frameLines - findRules() の frameLines
     * @returns {void}
     */
    function removeFrameLines(frameLines) {
        var lines = frameLines.vertical.concat(frameLines.horizontal);
        for (var i = 0; i < lines.length; i++) lines[i].remove();
    }

    // =========================================
    // 角丸とセルの塗り / Rounded corners and cell fills
    // =========================================

    /**
     * ［効果］→［スタイライズ］→［角を丸くする］を適用する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} radiusPt - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(pathItem, radiusPt) {
        pathItem.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>');
    }

    /* 円弧をベジェで近似するハンドルの長さ（半径に対する比） / Handle length approximating a quarter circle (ratio to the radius) */
    var CIRCLE_KAPPA = 0.5523;

    /**
     * 角ごとに半径を指定した長方形のアンカーポイントを、左上から時計回りに返す
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} right - 右端
     * @param {number} bottom - 下端
     * @param {number[]} radii - [左上, 右上, 右下, 左下] の半径（0 なら角は丸めない）
     * @returns {Object[]} {anchor, leftDirection, rightDirection} の配列
     */
    function buildRoundedRectPoints(left, top, right, bottom, radii) {
        var maxRadius = Math.min(right - left, top - bottom) / 2;
        var clampedRadii = [];
        for (var i = 0; i < 4; i++) clampedRadii.push(Math.max(0, Math.min(radii[i], maxRadius)));
        var points = [];

        /**
         * 角を1つ足す（丸めるなら円弧の両端の2点、丸めないなら角の1点）
         * @param {number[]} cornerPoint - 角の座標
         * @param {number[]} arcStart - 円弧の始点
         * @param {number[]} arcEnd - 円弧の終点
         * @param {number} radius - 半径
         * @returns {void}
         */
        function addCorner(cornerPoint, arcStart, arcEnd, radius) {
            if (radius <= 0) {
                points.push({ anchor: cornerPoint, leftDirection: cornerPoint, rightDirection: cornerPoint });
                return;
            }
            /* ハンドルは円弧の端から角の方向へ伸ばす / Handles point from each arc end toward the corner */
            var startHandle = [arcStart[0] + (cornerPoint[0] - arcStart[0]) * CIRCLE_KAPPA, arcStart[1] + (cornerPoint[1] - arcStart[1]) * CIRCLE_KAPPA];
            var endHandle = [arcEnd[0] + (cornerPoint[0] - arcEnd[0]) * CIRCLE_KAPPA, arcEnd[1] + (cornerPoint[1] - arcEnd[1]) * CIRCLE_KAPPA];
            points.push({ anchor: arcStart, leftDirection: arcStart, rightDirection: startHandle });
            points.push({ anchor: arcEnd, leftDirection: endHandle, rightDirection: arcEnd });
        }

        addCorner([left, top], [left, top - clampedRadii[0]], [left + clampedRadii[0], top], clampedRadii[0]);
        addCorner([right, top], [right - clampedRadii[1], top], [right, top - clampedRadii[1]], clampedRadii[1]);
        addCorner([right, bottom], [right, bottom + clampedRadii[2]], [right - clampedRadii[2], bottom], clampedRadii[2]);
        addCorner([left, bottom], [left + clampedRadii[3], bottom], [left, bottom + clampedRadii[3]], clampedRadii[3]);
        return points;
    }

    /**
     * パスの形を、指定したアンカーポイントで置き換える（閉じたパスにする）
     * @param {PathItem} pathItem - 対象のパス
     * @param {Object[]} points - buildRoundedRectPoints() の戻り値
     * @returns {void}
     */
    function setPathShape(pathItem, points) {
        var anchors = [];
        for (var i = 0; i < points.length; i++) anchors.push(points[i].anchor);
        pathItem.setEntirePath(anchors);
        for (var j = 0; j < points.length; j++) {
            var pathPoint = pathItem.pathPoints[j];
            pathPoint.leftDirection = points[j].leftDirection;
            pathPoint.rightDirection = points[j].rightDirection;
            pathPoint.pointType = PointType.CORNER;
        }
        pathItem.closed = true;
    }

    /**
     * セルの塗りに使う、K の濃さの色を返す（CMYK ドキュメントは K だけの CMYK、RGB ドキュメントはグレースケール）
     * @param {number} tint - 濃さ（%）
     * @returns {Color} 塗りの色
     */
    function createTintColor(tint) {
        if (app.activeDocument.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = tint;
            return cmykColor;
        }
        var grayColor = new GrayColor();
        grayColor.gray = tint;
        return grayColor;
    }

    /**
     * 外枠と、今の罫線の位置で区切られたセルを、左上から右下の順に返す
     * @param {Object} rules - findRules() の戻り値（罫線は並べ直したあとの位置で読む）
     * @returns {{columnEdges: number[], rowEdges: number[]}} 列の境界の X（左から右）と、行の境界の Y（上から下）
     */
    function getCellGrid(rules) {
        var rectBounds = rules.rectBounds;
        var columnEdges = [rectBounds.left];
        var rowEdges = [rectBounds.top];
        var i;
        for (i = 0; i < rules.vertical.length; i++) columnEdges.push(getCenterX(rules.vertical[i]));
        for (i = 0; i < rules.horizontal.length; i++) rowEdges.push(getCenterY(rules.horizontal[i]));
        columnEdges.push(rectBounds.right);
        rowEdges.push(rectBounds.bottom);
        /* 並べ直さなかった罫線も含め、位置の順に並べる / Sort by position, including rules that were not redistributed */
        columnEdges.sort(function (a, b) { return a - b; });
        rowEdges.sort(function (a, b) { return b - a; });
        return { columnEdges: columnEdges, rowEdges: rowEdges };
    }

    /**
     * 各セルに塗りの長方形を作り、最背面に置く（1つのグループにまとめるか、ばらのまま）
     * @param {Object} rules - findRules() の戻り値
     * @param {number} cornerRadius - 外枠の角丸の半径（pt。角丸にしないなら 0）
     * @param {string} fillMode - "uniform"（すべて FILL_UNIFORM_TINT）または "random"（セルごとにランダム）
     * @param {number[]} cellTints - セルごとのランダムな濃さ（足りない分はここに足していき、プレビューのたびに色が変わらないようにする）
     * @param {PageItem} backmostItem - 選択の中でいちばん背面のオブジェクト（このオブジェクトのレイヤー・グループの最背面に置く）
     * @param {boolean} isGrouped - グループにまとめるなら true
     * @returns {PageItem[]} 作ったもの（グループなら1つ、ばらなら長方形の配列）
     */
    function createCellFills(rules, cornerRadius, fillMode, cellTints, backmostItem, isGrouped) {
        var grid = getCellGrid(rules);
        var lastColumn = grid.columnEdges.length - 2;
        var lastRow = grid.rowEdges.length - 2;
        var fillGroup = app.activeDocument.groupItems.add();

        for (var row = 0; row <= lastRow; row++) {
            for (var column = 0; column <= lastColumn; column++) {
                var cellIndex = row * (lastColumn + 1) + column;
                if (fillMode === "random" && cellTints[cellIndex] === undefined) {
                    cellTints[cellIndex] = FILL_TINT_MIN + Math.floor(Math.random() * (FILL_TINT_MAX - FILL_TINT_MIN + 1));
                }
                /* 外枠の角に当たる角だけを丸める / Round only the corners that sit on the frame corners */
                var radii = [
                    (row === 0 && column === 0) ? cornerRadius : 0,
                    (row === 0 && column === lastColumn) ? cornerRadius : 0,
                    (row === lastRow && column === lastColumn) ? cornerRadius : 0,
                    (row === lastRow && column === 0) ? cornerRadius : 0
                ];
                var cellFill = fillGroup.pathItems.add();
                setPathShape(cellFill, buildRoundedRectPoints(
                    grid.columnEdges[column], grid.rowEdges[row], grid.columnEdges[column + 1], grid.rowEdges[row + 1], radii));
                cellFill.stroked = false;
                cellFill.filled = true;
                cellFill.fillColor = createTintColor(fillMode === "random" ? cellTints[cellIndex] : FILL_UNIFORM_TINT);
            }
        }

        /* 選択のいちばん背面のオブジェクトと同じレイヤー（グループ）に入れ、その中の最背面へ送る
           Put it in the same layer (or group) as the backmost selected item, then send it to the back there */
        fillGroup.move(backmostItem, ElementPlacement.PLACEAFTER);
        fillGroup.zOrder(ZOrderMethod.SENDTOBACK);
        if (isGrouped) return [fillGroup];

        /* グループにしないときは、前面の長方形から順にグループの直前へ出して重なり順を保ち、空のグループを消す
           When not grouping, move the rectangles out front to back just before the group to keep their order, then delete the empty group */
        var looseFills = [];
        while (fillGroup.pageItems.length > 0) {
            var looseFill = fillGroup.pageItems[0];
            looseFill.move(fillGroup, ElementPlacement.PLACEBEFORE);
            looseFills.push(looseFill);
        }
        fillGroup.remove();
        return looseFills;
    }

    // =========================================
    // テキスト / Text
    // =========================================

    /**
     * 字形の外形で測ったテキストの範囲を返す（アウトライン化した複製で測る）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[]|null} [左, 上, 右, 下]。字形が無ければ null
     */
    function getGlyphBounds(textFrame) {
        var duplicateFrame = textFrame.duplicate();
        var outlineGroup = null;
        var bounds = null;
        try {
            outlineGroup = duplicateFrame.createOutline();
            bounds = outlineGroup.geometricBounds;
        } catch (e) {
            bounds = null;
        }
        /* createOutline() は複製を消費するので、残っていれば消す / createOutline() consumes the duplicate; remove it only if it remains */
        try { duplicateFrame.remove(); } catch (e) { }
        if (outlineGroup) outlineGroup.remove();
        return bounds;
    }


    /**
     * 字形の範囲から、中心と大きさを求める
     * @param {number[]|null} bounds - getGlyphBounds() の戻り値
     * @returns {{center: number[], size: number[]}|null} 中心 [X, Y] と大きさ [幅, 高さ]。範囲が無ければ null
     */
    function getBoundsGeometry(bounds) {
        if (!bounds) return null;
        return {
            center: [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2],
            size: [bounds[2] - bounds[0], bounds[1] - bounds[3]]
        };
    }

    /**
     * テキストごとに、字形の中心・大きさと、それを囲むセルの辺を集める
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @param {{vertical: Object[], horizontal: Object[]}} cellEdges - collectEdges() の戻り値
     * @returns {Object[]} テキストごとの {textFrame, glyphCenter, glyphSize, column, row}（字形やセルが無ければ null）
     */
    function collectTextCells(textFrames, cellEdges) {
        var textCells = [];
        for (var i = 0; i < textFrames.length; i++) {
            var glyphGeometry = getBoundsGeometry(getGlyphBounds(textFrames[i]));
            var glyphCenter = glyphGeometry ? glyphGeometry.center : null;
            textCells.push({
                textFrame: textFrames[i],
                glyphCenter: glyphCenter,
                glyphSize: glyphGeometry ? glyphGeometry.size : null,
                column: glyphCenter ? findNearestEdges(cellEdges.vertical, glyphCenter[0], glyphCenter[1]) : null,
                row: glyphCenter ? findNearestEdges(cellEdges.horizontal, glyphCenter[1], glyphCenter[0]) : null
            });
        }
        return textCells;
    }

    /**
     * 行揃えを設定し、テキストを動かしたあとのセルの中央へ移動する
     * @param {Object[]} textCells - collectTextCells() の戻り値
     * @param {{centerInCells: boolean, justifyCenter: boolean}} settings - 設定
     * @param {{vertical: Object[], horizontal: Object[]}} positionMap - adjustRuleSpacing() の戻り値
     * @param {{vertical: Object, horizontal: Object}} frameScale - 外枠の伸縮（mapPosition() に渡す）
     * @returns {number} セルが見つからず移動しなかったテキストの数
     */
    function placeTextsInCells(textCells, settings, positionMap, frameScale) {
        var skippedCount = 0;
        for (var i = 0; i < textCells.length; i++) {
            var textCell = textCells[i];
            if (settings.justifyCenter) setJustifyCenter(textCell.textFrame);
            if (!settings.centerInCells) continue;

            if (!textCell.column || !textCell.row) {
                skippedCount++;
                continue;
            }

            var currentCenter = textCell.glyphCenter;
            if (settings.justifyCenter) {
                /* 行揃えを変えると字形が動くので測り直す / Re-measure, since changing the alignment moves the glyphs */
                var currentGeometry = getBoundsGeometry(getGlyphBounds(textCell.textFrame));
                currentCenter = currentGeometry ? currentGeometry.center : null;
            }
            if (!currentCenter) {
                skippedCount++;
                continue;
            }
            var column = textCell.column;
            var row = textCell.row;
            var cellCenterX = (mapPosition(positionMap.vertical, column.lower, column.lowerItem, frameScale.vertical) +
                mapPosition(positionMap.vertical, column.upper, column.upperItem, frameScale.vertical)) / 2;
            var cellCenterY = (mapPosition(positionMap.horizontal, row.lower, row.lowerItem, frameScale.horizontal) +
                mapPosition(positionMap.horizontal, row.upper, row.upperItem, frameScale.horizontal)) / 2;
            textCell.textFrame.translate(cellCenterX - currentCenter[0], cellCenterY - currentCenter[1]);
        }
        return skippedCount;
    }

    /**
     * 段落の行揃えを中央揃えにする
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {void}
     */
    function setJustifyCenter(textFrame) {
        textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
    }

    // =========================================
    // 状態の保存と復元 / Saving and restoring state
    // =========================================

    /**
     * 編集前のパスの座標・線幅・線端を保存する
     * @param {PathItem[]} targetPathItems - 保存するパス
     * @returns {Object[]} パスごとのアンカーポイントの状態・線幅・線端
     */
    function savePathState(targetPathItems) {
        var pathStates = [];

        for (var pathIndex = 0; pathIndex < targetPathItems.length; pathIndex++) {
            var pathItem = targetPathItems[pathIndex];
            var pointStates = [];

            for (var pointIndex = 0; pointIndex < pathItem.pathPoints.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];

                pointStates.push({
                    anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                    leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                    rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]],
                    pointType: pathPoint.pointType
                });
            }

            pathStates.push({ pathItem: pathItem, pointStates: pointStates, strokeWidth: pathItem.strokeWidth, strokeCap: pathItem.strokeCap });
        }

        return pathStates;
    }

    /**
     * 保存しておいた座標・線幅・線端へ戻す
     * @param {Object[]} pathStates - savePathState() の戻り値
     * @returns {void}
     */
    function restorePathState(pathStates) {
        for (var pathIndex = 0; pathIndex < pathStates.length; pathIndex++) {
            var pathItem = pathStates[pathIndex].pathItem;
            var pointStates = pathStates[pathIndex].pointStates;

            for (var pointIndex = 0; pointIndex < pointStates.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];
                var pointState = pointStates[pointIndex];

                pathPoint.anchor = [pointState.anchor[0], pointState.anchor[1]];
                pathPoint.leftDirection = [pointState.leftDirection[0], pointState.leftDirection[1]];
                pathPoint.rightDirection = [pointState.rightDirection[0], pointState.rightDirection[1]];
                pathPoint.pointType = pointState.pointType;
            }
            pathItem.strokeWidth = pathStates[pathIndex].strokeWidth;
            pathItem.strokeCap = pathStates[pathIndex].strokeCap;
        }
    }

    /**
     * 編集前のテキストの位置と、段落ごとの行揃えを保存する
     * @param {TextFrame[]} textFrames - 保存するテキスト
     * @returns {Object[]} テキストごとの {textFrame, position, justifications}
     */
    function saveTextState(textFrames) {
        var textStates = [];
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var justifications = [];
            for (var j = 0; j < textFrame.paragraphs.length; j++) {
                justifications.push(textFrame.paragraphs[j].paragraphAttributes.justification);
            }
            textStates.push({
                textFrame: textFrame,
                position: [textFrame.position[0], textFrame.position[1]],
                justifications: justifications
            });
        }
        return textStates;
    }

    /**
     * 保存しておいた行揃えと位置へ戻す（行揃えで字形が動くので、位置はそのあとで戻す）
     * @param {Object[]} textStates - saveTextState() の戻り値
     * @returns {void}
     */
    function restoreTextState(textStates) {
        for (var i = 0; i < textStates.length; i++) {
            var textState = textStates[i];
            var paragraphs = textState.textFrame.paragraphs;
            for (var j = 0; j < textState.justifications.length && j < paragraphs.length; j++) {
                paragraphs[j].paragraphAttributes.justification = textState.justifications[j];
            }
            textState.textFrame.position = [textState.position[0], textState.position[1]];
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択からテキスト・枠・罫線を集め、ダイアログで均等化・配置を確定する
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        if (!selectedItems) {
            alert(getLabel("alert.noText"));
            return;
        }

        var textFrames = [];
        var pathItems = [];
        for (var i = 0; i < selectedItems.length; i++) collectItems(selectedItems[i], textFrames, pathItems);
        if (!textFrames.length) {
            alert(getLabel("alert.noText"));
            return;
        }

        var cellEdges = collectEdges(pathItems);
        if (!cellEdges.vertical.length || !cellEdges.horizontal.length) {
            alert(getLabel("alert.noRule"));
            return;
        }

        /* 罫線を動かすとセルの判定がずれるので、先に今のセルを決めておく
           Find each text's cell before moving rules, since moving them changes the result */
        var textCells = collectTextCells(textFrames, cellEdges);

        var rules = findRules(pathItems);
        if (rules) {
            /* ［テキストに合わせる］で使う区画ごとの大きさは、動かす前に測っておく / Measure section sizes for "Fit to text" before anything moves */
            rules.columnSizes = measureContentSizes(textCells, rules, true);
            rules.rowSizes = measureContentSizes(textCells, rules, false);
        }
        var verticalRuleCount = rules ? rules.vertical.length : 0;
        var horizontalRuleCount = rules ? rules.horizontal.length : 0;
        var pathState = savePathState(pathItems);
        var textState = saveTextState(textFrames);
        var frameRectangle = null; /* ［外枠を長方形に変換］で作った長方形 / rectangle made by "Convert frame to rectangle" */
        var roundedFramePreview = null; /* 角丸のプレビュー用に作った外枠の複製（効果は取り消せないので、元は隠しておく） / duplicate of the frame for the round-corner preview (effects cannot be removed, so the original stays hidden) */
        var cellFillItems = [];    /* ［図形を作成］で作ったグループ、またはばらの長方形 / group or loose rectangles made by "Create cell shapes" */
        var cellTints = [];        /* セルごとの塗りの濃さ / fill tint of each cell */
        var backmostItem = selectedItems[selectedItems.length - 1]; /* 選択は前面→背面の順 / the selection runs front to back */
        var originalBounds = rules ? rules.rectBounds : null; /* 大きさを変える前の外枠の範囲 / frame bounds before resizing */
        var tablePathItems = rules ? getTablePathItems(rules) : []; /* 外枠の大きさに合わせて伸縮するパス / paths scaled with the frame size */

        /**
         * 罫線・テキストを元に戻し、作った長方形・塗り・角丸の複製を消して、隠したパスを表示し直す
         * @returns {void}
         */
        function restoreOriginalState() {
            for (var fillIndex = 0; fillIndex < cellFillItems.length; fillIndex++) cellFillItems[fillIndex].remove();
            cellFillItems = [];
            if (roundedFramePreview) {
                roundedFramePreview.remove();
                roundedFramePreview = null;
                rules.outerRect.hidden = false;
            }
            if (frameRectangle) {
                frameRectangle.remove();
                frameRectangle = null;
                setFrameLinesHidden(rules.frameLines, false);
            }
            restorePathState(pathState);
            restoreTextState(textState);
            if (rules) rules.rectBounds = originalBounds;
        }

        /**
         * 外枠と罫線を、左上を基準に指定の大きさへ伸縮し、rules.rectBounds を伸縮後の範囲にする
         * @param {number} frameWidth - 外枠の幅（pt）
         * @param {number} frameHeight - 外枠の高さ（pt）
         * @returns {{vertical: Object, horizontal: Object}} mapPosition() に渡す、軸ごとの伸縮
         */
        function resizeFrame(frameWidth, frameHeight) {
            var ratioX = frameWidth / originalBounds.width;
            var ratioY = frameHeight / originalBounds.height;
            if (Math.abs(frameWidth - originalBounds.width) > EPSILON || Math.abs(frameHeight - originalBounds.height) > EPSILON) {
                scalePathItems(tablePathItems, originalBounds.left, originalBounds.top, ratioX, ratioY);
            }
            rules.rectBounds = {
                left: originalBounds.left,
                top: originalBounds.top,
                right: originalBounds.left + frameWidth,
                bottom: originalBounds.top - frameHeight,
                width: frameWidth,
                height: frameHeight
            };
            return {
                vertical: { pathItems: tablePathItems, origin: originalBounds.left, ratio: ratioX },
                horizontal: { pathItems: tablePathItems, origin: originalBounds.top, ratio: ratioY }
            };
        }

        /**
         * 元の状態に戻したうえで、設定どおりに均等化・配置する
         * @param {{frameWidth: number, frameHeight: number, adjustVerticalSpacing: boolean, adjustHorizontalSpacing: boolean, distributeMode: string, matchRuleLengths: boolean, convertFrame: boolean, roundCorners: boolean, cornerRadius: number, matchStrokeWidths: boolean, strokeWidth: number, createCellShapes: boolean, fillMode: string, groupFills: boolean, centerInCells: boolean, justifyCenter: boolean}} settings - 設定（frameWidth・frameHeight・cornerRadius・strokeWidth は pt）
         * @param {boolean} [isFinal] - ［OK］で確定するとき true（角丸の効果を複製ではなく外枠そのものに適用する）
         * @returns {number} セルが見つからず移動しなかったテキストの数
         */
        function applySettings(settings, isFinal) {
            restoreOriginalState();

            var identityScale = { pathItems: [], origin: 0, ratio: 1 };
            var frameScale = rules ? resizeFrame(settings.frameWidth, settings.frameHeight) : { vertical: identityScale, horizontal: identityScale };
            var positionMap = adjustRuleSpacing(rules, settings);
            if (settings.convertFrame && rules && rules.frameLines) {
                frameRectangle = createFrameRectangle(rules.frameLines, rules.rectBounds);
            } else if (rules && rules.frameLines) {
                /* 罫線で囲んだ外枠は、角が欠けないよう突出線端にする / A frame drawn with rules gets projecting caps so the corners close */
                setStrokeCaps(rules.frameLines.vertical.concat(rules.frameLines.horizontal), StrokeCap.PROJECTINGENDCAP);
            }
            if (settings.matchStrokeWidths && rules) {
                /* 角丸の複製より先に掛け、複製にも同じ線幅を引き継がせる / Set before the round-corner duplicate so it inherits the width */
                setStrokeWidths(frameRectangle ? tablePathItems.concat([frameRectangle]) : tablePathItems, settings.strokeWidth);
            }

            /* 角丸は外枠が長方形のときだけ / Round corners only when the frame is a rectangle */
            var frameShape = frameRectangle || (rules ? rules.outerRect : null);
            var cornerRadius = (settings.roundCorners && frameShape) ? settings.cornerRadius : 0;
            if (cornerRadius > 0) {
                /* 作った長方形は消せば戻るので直接、元からある外枠はプレビュー中だけ複製に掛ける
                   The created rectangle is simply removed on restore; an existing frame gets the effect on a duplicate while previewing */
                if (frameShape === rules.outerRect && !isFinal) {
                    /* duplicate() は hidden を引き継ぐので、隠す前に複製する / duplicate() inherits hidden, so duplicate before hiding */
                    roundedFramePreview = rules.outerRect.duplicate();
                    rules.outerRect.hidden = true;
                    frameShape = roundedFramePreview;
                }
                applyRoundCornersEffect(frameShape, cornerRadius);
            }
            if (settings.createCellShapes && rules) {
                cellFillItems = createCellFills(rules, cornerRadius, settings.fillMode, cellTints, backmostItem, settings.groupFills);
            }
            return placeTextsInCells(textCells, settings, positionMap, frameScale);
        }

        /* 開いた時点の設定（見つかった罫線は ON） / Initial settings: rules found start ON */
        var initialSettings = {
            frameWidth: rules ? rules.rectBounds.width : 0,
            frameHeight: rules ? rules.rectBounds.height : 0,
            adjustVerticalSpacing: ADJUST_RULE_SPACING && verticalRuleCount > 0,
            adjustHorizontalSpacing: ADJUST_RULE_SPACING && horizontalRuleCount > 0,
            distributeMode: DISTRIBUTE_MODE,
            matchRuleLengths: MATCH_RULE_LENGTHS,
            convertFrame: CONVERT_FRAME_TO_RECT && !!(rules && rules.frameLines),
            roundCorners: ROUND_CORNERS,
            cornerRadius: CORNER_RADIUS_PT,
            matchStrokeWidths: MATCH_STROKE_WIDTHS && tablePathItems.length > 0,
            strokeWidth: getMostCommonStrokeWidth(tablePathItems),
            createCellShapes: CREATE_CELL_SHAPES && !!rules,
            fillMode: CELL_FILL_MODE,
            groupFills: GROUP_CELL_FILLS,
            centerInCells: CENTER_IN_CELLS,
            justifyCenter: JUSTIFY_CENTER
        };

        var unitInfo = getUnitInfo();
        var strokeUnitInfo = getUnitInfo("strokeUnits");
        var dialogControls = buildDialog(verticalRuleCount, horizontalRuleCount, rules, unitInfo, strokeUnitInfo, initialSettings);
        var frameWidthInput = dialogControls.frameWidthInput;
        var frameHeightInput = dialogControls.frameHeightInput;
        var adjustVerticalCheckbox = dialogControls.adjustVerticalCheckbox;
        var adjustHorizontalCheckbox = dialogControls.adjustHorizontalCheckbox;
        var spacingRow = dialogControls.spacingRow;
        var evenSpacingRadio = dialogControls.evenSpacingRadio;
        var contentSpacingRadio = dialogControls.contentSpacingRadio;
        var matchRuleLengthsCheckbox = dialogControls.matchRuleLengthsCheckbox;
        var convertFrameCheckbox = dialogControls.convertFrameCheckbox;
        var cornerRow = dialogControls.cornerRow;
        var roundCornersCheckbox = dialogControls.roundCornersCheckbox;
        var cornerRadiusInput = dialogControls.cornerRadiusInput;
        var matchStrokeCheckbox = dialogControls.matchStrokeCheckbox;
        var strokeWidthInput = dialogControls.strokeWidthInput;
        var createCellShapesCheckbox = dialogControls.createCellShapesCheckbox;
        var fillModeGroup = dialogControls.fillModeGroup;
        var uniformFillRadio = dialogControls.uniformFillRadio;
        var randomFillRadio = dialogControls.randomFillRadio;
        var groupFillsCheckbox = dialogControls.groupFillsCheckbox;
        var centerInCellsCheckbox = dialogControls.centerInCellsCheckbox;
        var justifyCenterCheckbox = dialogControls.justifyCenterCheckbox;

        /**
         * チェックボックスの状態を設定にまとめる（表示前は value を読み戻せないので、ユーザー操作のあとだけ呼ぶ）
         * @returns {Object} applySettings() に渡す設定
         */
        function getCheckedSettings() {
            return {
                frameWidth: frameWidthPt,
                frameHeight: frameHeightPt,
                adjustVerticalSpacing: adjustVerticalCheckbox.value,
                adjustHorizontalSpacing: adjustHorizontalCheckbox.value,
                distributeMode: contentSpacingRadio.value ? "content" : "even",
                matchRuleLengths: matchRuleLengthsCheckbox.value,
                convertFrame: convertFrameCheckbox.value,
                roundCorners: roundCornersCheckbox.value,
                cornerRadius: cornerRadiusPt,
                matchStrokeWidths: matchStrokeCheckbox.value,
                strokeWidth: strokeWidthPt,
                createCellShapes: createCellShapesCheckbox.value,
                fillMode: randomFillRadio.value ? "random" : "uniform",
                groupFills: groupFillsCheckbox.value,
                centerInCells: centerInCellsCheckbox.value,
                justifyCenter: justifyCenterCheckbox.value
            };
        }

        /* 有効／無効の切り替え。表示前は checkbox.value を読み戻せないので、状態は引数で受け取る
           Enabling helpers take the state as an argument, since checkbox.value cannot be read back before show() */

        /**
         * ［間隔］［長さを外枠に合わせる］は、どちらかの罫線の間隔を調整するときだけ操作できるようにする
         * @param {boolean} isAdjustingSpacing - 縦罫か横罫の間隔を調整するなら true
         * @returns {void}
         */
        function updateRuleOptionsEnabled(isAdjustingSpacing) {
            spacingRow.enabled = isAdjustingSpacing;
            matchRuleLengthsCheckbox.enabled = isAdjustingSpacing;
        }

        /**
         * ［角を丸くする］は、外枠が長方形のとき（閉じたパスか、［外枠を長方形に変換］がオン）だけ操作できるようにする
         * @param {boolean} isConvertingFrame - ［外枠を長方形に変換］がオンなら true
         * @returns {void}
         */
        function updateCornerRowEnabled(isConvertingFrame) {
            cornerRow.enabled = !!(rules && (rules.outerRect || (rules.frameLines && isConvertingFrame)));
            /* ∧∨は自作描画なので、描き直してディム表示を切り替える / Redraw the custom-drawn stepper to update its dimming */
            redrawSteppersIn(cornerRow);
        }

        /**
         * 塗り方と［グループ化］は、セルの図形を作るときだけ操作できるようにする
         * @param {boolean} isCreating - ［図形を作成］がオンなら true
         * @returns {void}
         */
        function updateCellShapeOptionsEnabled(isCreating) {
            fillModeGroup.enabled = isCreating;
            groupFillsCheckbox.enabled = isCreating;
        }

        updateRuleOptionsEnabled(initialSettings.adjustVerticalSpacing || initialSettings.adjustHorizontalSpacing);
        updateCornerRowEnabled(initialSettings.convertFrame);
        updateCellShapeOptionsEnabled(initialSettings.createCellShapes);

        /* 欄の値（pt）。不正な値が入ったら直前の値に戻す / Field values in pt; invalid input falls back to the previous value */
        var frameWidthPt = initialSettings.frameWidth;
        var frameHeightPt = initialSettings.frameHeight;
        var cornerRadiusPt = initialSettings.cornerRadius;
        var strokeWidthPt = initialSettings.strokeWidth;

        /**
         * 長さの欄の入力を読み取り、pt に換算して返す（計算式・別の単位も欄の単位へ換算する）。
         * 欄の表示は小数第2位までなので、表示が前回読んだときのままなら控えた値を使い、丸めた分だけ外枠が伸縮しないようにする
         * @param {EditText} valueInput - 対象の欄
         * @param {number} previousPt - 直前の値（pt）
         * @param {boolean} allowsZero - 0 を受け付けるなら true（半径）。false なら 0 より大きい値だけ（幅・高さ）
         * @returns {number} 読み取った値（pt）。不正な値なら直前の値
         */
        function readLengthInput(valueInput, previousPt, allowsZero) {
            if (valueInput.text === valueInput.readText) return previousPt;
            var inputValue = evaluateArithmetic(valueInput.text, valueInput.stepperGroup.stepOptions.unit);
            if (isNaN(inputValue) || (allowsZero ? inputValue < 0 : inputValue <= 0)) return previousPt;
            valueInput.readText = valueInput.text;
            return inputValue * valueInput.pointsPerUnit;
        }

        /**
         * 幅・高さ・半径・線幅の欄をすべて読み取る
         * @returns {void}
         */
        function readLengthInputs() {
            frameWidthPt = readLengthInput(frameWidthInput, frameWidthPt, false);
            frameHeightPt = readLengthInput(frameHeightInput, frameHeightPt, false);
            cornerRadiusPt = readLengthInput(cornerRadiusInput, cornerRadiusPt, true);
            strokeWidthPt = readLengthInput(strokeWidthInput, strokeWidthPt, true);
        }

        /**
         * プレビューを更新する（プレビューは常にオン。［キャンセル］で元に戻す）
         * @returns {void}
         */
        function updatePreview() {
            applySettings(getCheckedSettings());
            app.redraw();
        }

        /**
         * 罫線チェックボックスのクリック処理（Option／Alt キーで縦罫・横罫をまとめて切り替える）
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @returns {void}
         */
        function handleRuleCheckboxClick(clickedCheckbox) {
            if (ScriptUI.environment.keyboardState.altKey) {
                if (adjustVerticalCheckbox.enabled) adjustVerticalCheckbox.value = clickedCheckbox.value;
                if (adjustHorizontalCheckbox.enabled) adjustHorizontalCheckbox.value = clickedCheckbox.value;
            }
            updateRuleOptionsEnabled(adjustVerticalCheckbox.value || adjustHorizontalCheckbox.value);
            updatePreview();
        }

        adjustVerticalCheckbox.onClick = function () {
            handleRuleCheckboxClick(adjustVerticalCheckbox);
        };
        adjustHorizontalCheckbox.onClick = function () {
            handleRuleCheckboxClick(adjustHorizontalCheckbox);
        };
        evenSpacingRadio.onClick = updatePreview;
        contentSpacingRadio.onClick = updatePreview;
        matchRuleLengthsCheckbox.onClick = updatePreview;
        convertFrameCheckbox.onClick = function () {
            updateCornerRowEnabled(convertFrameCheckbox.value);
            updatePreview();
        };
        roundCornersCheckbox.onClick = updatePreview;
        matchStrokeCheckbox.onClick = updatePreview;

        var lengthInputs = [frameWidthInput, frameHeightInput, cornerRadiusInput, strokeWidthInput];
        for (var inputIndex = 0; inputIndex < lengthInputs.length; inputIndex++) {
            /* 開いた時点の表示は初期値そのものとして扱う / The initial text stands for the exact initial value */
            lengthInputs[inputIndex].readText = lengthInputs[inputIndex].text;
        }

        /**
         * 欄の値を読み直してプレビューを更新する（∧∨・↑↓キー・直接入力のあと）
         * @returns {void}
         */
        function handleLengthChange() {
            readLengthInputs();
            updatePreview();
        }

        /**
         * 直接入力をそろえる addSteppedField() の onChange のあとで、プレビューを更新する
         * @param {EditText} valueInput - 対象の欄
         * @returns {void}
         */
        function chainLengthOnChange(valueInput) {
            var normalizeInput = valueInput.onChange;
            valueInput.onChange = function () {
                normalizeInput();
                handleLengthChange();
            };
        }

        /**
         * チェックボックスの横の欄（半径・線幅）の直接入力をそろえ、プレビューを更新する
         * @param {EditText} valueInput - 対象の欄
         * @returns {void}
         */
        function wireCheckboxLengthOnChange(valueInput) {
            valueInput.onChange = function () {
                /* 数値でなければ直前の値に戻し、下限と単位をそろえる / Revert non-numbers; apply the minimum and the unit */
                var stepOptions = valueInput.stepperGroup.stepOptions;
                var inputValue = evaluateArithmetic(valueInput.text, stepOptions.unit);
                if (isNaN(inputValue)) valueInput.text = valueInput.lastValidText;
                else writeSteppedValue(valueInput, inputValue, stepOptions);
                handleLengthChange();
            };
        }

        chainLengthOnChange(frameWidthInput);
        chainLengthOnChange(frameHeightInput);
        wireCheckboxLengthOnChange(cornerRadiusInput);
        wireCheckboxLengthOnChange(strokeWidthInput);
        for (inputIndex = 0; inputIndex < lengthInputs.length; inputIndex++) {
            /* 入力中も1文字ごとにプレビューする（数値として読めない途中の入力は直前の値のまま） / Preview on every keystroke; incomplete input keeps the previous value */
            lengthInputs[inputIndex].onChanging = handleLengthChange;
            lengthInputs[inputIndex].stepperGroup.stepOptions.onStep = handleLengthChange;
        }
        uniformFillRadio.onClick = updatePreview;
        randomFillRadio.onClick = updatePreview;
        createCellShapesCheckbox.onClick = function () {
            updateCellShapeOptionsEnabled(createCellShapesCheckbox.value);
            updatePreview();
        };
        groupFillsCheckbox.onClick = updatePreview;
        centerInCellsCheckbox.onClick = updatePreview;
        justifyCenterCheckbox.onClick = updatePreview;

        var skippedCount = 0;
        dialogControls.btnOK.onClick = function () {
            /* 欄を確定せずに［OK］を押したときも、入力中の値を使う / Use the value being typed even when OK is pressed without committing it */
            readLengthInputs();
            skippedCount = applySettings(getCheckedSettings(), true);
            app.redraw();
            dialogControls.mainDialog.close(1);
        };
        dialogControls.btnCancel.onClick = function () {
            dialogControls.mainDialog.close(0);
        };

        /* 「キャンセル」の文字では Esc の取消ボタンと認識されない / "キャンセル" is not recognized as the Esc button */
        dialogControls.mainDialog.cancelElement = dialogControls.btnCancel;

        /* 初回のプレビュー（プレビューは常にオン） / Initial preview (the preview is always on) */
        applySettings(initialSettings);
        app.redraw();

        prepareDialogWindow(dialogControls.mainDialog, SCRIPT_NAME);
        /* 表は横に長く、選択範囲を避けると画面の端まで追いやられるので、避けずに前回の位置で開く
           Tables are wide, so avoiding the selection pushes the dialog to the screen edge; open at the last location instead */
        dialogControls.mainDialog.dialogWindowState.selectionSpan = null;

        /* Esc など［OK］以外で閉じたときも元に戻す / Restore however the dialog is closed, unless by OK */
        if (dialogControls.mainDialog.show() !== 1) {
            restoreOriginalState();
            app.redraw();
            return;
        }

        /* 変換を確定したら、隠しておいた枠の罫線を削除する / Once the conversion is confirmed, delete the hidden frame rules */
        if (frameRectangle) removeFrameLines(rules.frameLines);

        if (skippedCount) alert(getLabel("alert.skipped", { count: skippedCount }));
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる
     * @param {number} verticalRuleCount - 見つかった縦罫の数
     * @param {number} horizontalRuleCount - 見つかった横罫の数
     * @param {Object|null} rules - findRules() の戻り値
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 定規の単位
     * @param {{label: string, pointsPerUnit: number}} strokeUnitInfo - 線の単位
     * @param {{frameWidth: number, frameHeight: number, adjustVerticalSpacing: boolean, adjustHorizontalSpacing: boolean, distributeMode: string, matchRuleLengths: boolean, convertFrame: boolean, roundCorners: boolean, cornerRadius: number, matchStrokeWidths: boolean, strokeWidth: number, createCellShapes: boolean, fillMode: string, groupFills: boolean, centerInCells: boolean, justifyCenter: boolean}} initialSettings - コントロールの初期値
     * @returns {{mainDialog: Window, frameWidthInput: EditText, frameHeightInput: EditText, adjustVerticalCheckbox: Checkbox, adjustHorizontalCheckbox: Checkbox, spacingRow: Group, evenSpacingRadio: RadioButton, contentSpacingRadio: RadioButton, matchRuleLengthsCheckbox: Checkbox, convertFrameCheckbox: Checkbox, cornerRow: Group, roundCornersCheckbox: Checkbox, cornerRadiusInput: EditText, matchStrokeCheckbox: Checkbox, strokeWidthInput: EditText, createCellShapesCheckbox: Checkbox, fillModeGroup: Group, uniformFillRadio: RadioButton, randomFillRadio: RadioButton, groupFillsCheckbox: Checkbox, centerInCellsCheckbox: Checkbox, justifyCenterCheckbox: Checkbox, btnCancel: Button, btnOK: Button}} ダイアログとコントロール
     */
    function buildDialog(verticalRuleCount, horizontalRuleCount, rules, unitInfo, strokeUnitInfo, initialSettings) {
        var hasFrameLines = !!(rules && rules.frameLines);
        var mainDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        /* 2カラム：左に［罫線］［テキスト］、右に［外枠］［セル］ / Two columns: Rules and Text on the left, Frame and Cells on the right */
        var columnsGroup = mainDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = columnsGroup.add("group");
        var rightColumn = columnsGroup.add("group");
        setupColumn(leftColumn);
        setupColumn(rightColumn);

        var framePanel = rightColumn.add("panel", undefined, getLabel("panel.frame"));
        setupPanel(framePanel, 6);
        /* 幅・高さの項目名は長いほうの幅にそろえて右揃え / Size labels share the wider label's width, right-aligned */
        var sizeLabelWidth = Math.max(
            mainDialog.graphics.measureString(labelText("fieldLabel.frameWidth"))[0],
            mainDialog.graphics.measureString(labelText("fieldLabel.frameHeight"))[0]);
        var frameWidthInput = addLengthField(framePanel, "fieldLabel.frameWidth", sizeLabelWidth, initialSettings.frameWidth, unitInfo);
        var frameHeightInput = addLengthField(framePanel, "fieldLabel.frameHeight", sizeLabelWidth, initialSettings.frameHeight, unitInfo);
        /* 大きさの欄とチェックボックスの間を少し空ける（コントロールには margins が無いのでグループで包む）
           Leave a little space below the size fields (controls have no margins, so wrap it in a group) */
        var convertFrameGroup = framePanel.add("group");
        convertFrameGroup.alignChildren = ["left", "center"];
        convertFrameGroup.margins = [0, FRAME_OPTIONS_TOP_GAP, 0, 0];
        var convertFrameCheckbox = addOptionCheckbox(convertFrameGroup,
            "checkbox.convertFrame", "tooltip.convertFrame", hasFrameLines, initialSettings.convertFrame);
        var cornerRowControls = addCheckboxLengthRow(framePanel, "checkbox.roundCorners", "tooltip.roundCorners",
            initialSettings.roundCorners, initialSettings.cornerRadius, unitInfo);
        var cornerRow = cornerRowControls.row;
        cornerRowControls.input.characters = CORNER_RADIUS_FIELD_CHARS;
        framePanel.enabled = !!rules;

        var rulesPanel = leftColumn.add("panel", undefined, getLabel("panel.rules"));
        setupPanel(rulesPanel, 6);

        var adjustVerticalCheckbox = addOptionCheckbox(rulesPanel,
            "checkbox.adjustVerticalSpacing", "tooltip.adjustVerticalSpacing", verticalRuleCount > 0, initialSettings.adjustVerticalSpacing);
        var adjustHorizontalCheckbox = addOptionCheckbox(rulesPanel,
            "checkbox.adjustHorizontalSpacing", "tooltip.adjustHorizontalSpacing", horizontalRuleCount > 0, initialSettings.adjustHorizontalSpacing);

        var spacingRow = rulesPanel.add("group");
        setupRow(spacingRow, "left", 8);
        spacingRow.alignChildren = ["left", "top"];
        spacingRow.add("statictext", undefined, labelText("fieldLabel.spacing"));
        /* ラジオは同じ親の中だけ排他になるので、2つを1つのグループに縦に積む / Radios are exclusive only within one parent, so stack both in one group */
        var spacingRadioGroup = spacingRow.add("group");
        spacingRadioGroup.orientation = "column";
        spacingRadioGroup.alignChildren = ["left", "top"];
        spacingRadioGroup.spacing = 6;
        var evenSpacingRadio = spacingRadioGroup.add("radiobutton", undefined, getLabel("radio.even"));
        evenSpacingRadio.helpTip = getLabel("tooltip.even");
        var contentSpacingRadio = spacingRadioGroup.add("radiobutton", undefined, getLabel("radio.content"));
        contentSpacingRadio.helpTip = getLabel("tooltip.content");
        contentSpacingRadio.value = initialSettings.distributeMode === "content";
        evenSpacingRadio.value = !contentSpacingRadio.value;

        /* ［間隔］のラジオボタンとの間を少し空ける（コントロールには margins が無いのでグループで包む）
           Leave a little space below the spacing radios (controls have no margins, so wrap it in a group) */
        var matchRuleLengthsGroup = rulesPanel.add("group");
        matchRuleLengthsGroup.alignChildren = ["left", "center"];
        matchRuleLengthsGroup.margins = [0, MATCH_LENGTHS_TOP_GAP, 0, 0];
        var matchRuleLengthsCheckbox = addOptionCheckbox(matchRuleLengthsGroup,
            "checkbox.matchRuleLengths", "tooltip.matchRuleLengths", true, initialSettings.matchRuleLengths);

        var strokeRowControls = addCheckboxLengthRow(rulesPanel, "checkbox.matchStrokeWidths", "tooltip.matchStrokeWidths",
            initialSettings.matchStrokeWidths, initialSettings.strokeWidth, strokeUnitInfo, true);
        /* 間隔の調整とは別の機能なので、上と少し離す / A separate feature from the spacing, so set it apart */
        strokeRowControls.row.margins = [0, SEPARATE_OPTION_GAP, 0, 0];
        strokeRowControls.row.enabled = !!rules;

        var textPanel = leftColumn.add("panel", undefined, getLabel("panel.text"));
        setupPanel(textPanel, 6);

        var centerInCellsCheckbox = addOptionCheckbox(textPanel,
            "checkbox.centerInCells", "tooltip.centerInCells", true, initialSettings.centerInCells);
        var justifyCenterCheckbox = addOptionCheckbox(textPanel,
            "checkbox.justifyCenter", "tooltip.justifyCenter", true, initialSettings.justifyCenter);

        var cellsPanel = rightColumn.add("panel", undefined, getLabel("panel.cells"));
        setupPanel(cellsPanel, 6);

        var createCellShapesCheckbox = addOptionCheckbox(cellsPanel,
            "checkbox.createCellShapes", "tooltip.createCellShapes", !!rules, initialSettings.createCellShapes);
        /* ラジオは同じ親の中だけ排他になるので、2つを1つのグループに入れる / Radios are exclusive only within one parent, so keep both in one group */
        /* 「塗り : ○ 同じ濃さ（K15） ○ ランダムな濃さ」を［罫線］パネルの［間隔］と同じ形で置き、［図形を作成］に従うので少し字下げする
           Laid out like "Spacing" in the Rules panel, indented under "Create cell shapes" */
        var fillModeGroup = cellsPanel.add("group");
        setupRow(fillModeGroup, "left", 8);
        fillModeGroup.alignChildren = ["left", "top"];
        fillModeGroup.margins = [SUB_OPTION_INDENT, 0, 0, 0];
        fillModeGroup.add("statictext", undefined, labelText("fieldLabel.fill"));
        var fillRadioGroup = fillModeGroup.add("group");
        fillRadioGroup.orientation = "column";
        fillRadioGroup.alignChildren = ["left", "top"];
        fillRadioGroup.spacing = 6;
        var uniformFillRadio = fillRadioGroup.add("radiobutton", undefined, getLabel("radio.uniformFill", { tint: FILL_UNIFORM_TINT }));
        uniformFillRadio.helpTip = getLabel("tooltip.uniformFill", { tint: FILL_UNIFORM_TINT });
        var randomFillRadio = fillRadioGroup.add("radiobutton", undefined, getLabel("radio.randomFill"));
        randomFillRadio.helpTip = getLabel("tooltip.randomFill", { min: FILL_TINT_MIN, max: FILL_TINT_MAX });
        randomFillRadio.value = initialSettings.fillMode === "random";
        uniformFillRadio.value = !randomFillRadio.value;
        var groupFillsCheckbox = addOptionCheckbox(cellsPanel,
            "checkbox.groupFills", "tooltip.groupFills", true, initialSettings.groupFills);

        var buttonRow = addButtonRow(mainDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"));
        alignRightOnlyButtonRow(buttonRow);

        return {
            mainDialog: mainDialog,
            frameWidthInput: frameWidthInput,
            frameHeightInput: frameHeightInput,
            adjustVerticalCheckbox: adjustVerticalCheckbox,
            adjustHorizontalCheckbox: adjustHorizontalCheckbox,
            spacingRow: spacingRow,
            evenSpacingRadio: evenSpacingRadio,
            contentSpacingRadio: contentSpacingRadio,
            matchRuleLengthsCheckbox: matchRuleLengthsCheckbox,
            convertFrameCheckbox: convertFrameCheckbox,
            cornerRow: cornerRow,
            roundCornersCheckbox: cornerRowControls.checkbox,
            cornerRadiusInput: cornerRowControls.input,
            matchStrokeCheckbox: strokeRowControls.checkbox,
            strokeWidthInput: strokeRowControls.input,
            createCellShapesCheckbox: createCellShapesCheckbox,
            fillModeGroup: fillModeGroup,
            uniformFillRadio: uniformFillRadio,
            randomFillRadio: randomFillRadio,
            groupFillsCheckbox: groupFillsCheckbox,
            centerInCellsCheckbox: centerInCellsCheckbox,
            justifyCenterCheckbox: justifyCenterCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * ダイアログのカラム（パネルを縦に積むグループ）の共通設定
     * @param {Group} columnGroup - 対象のグループ
     * @returns {void}
     */
    function setupColumn(columnGroup) {
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.alignment = ["fill", "top"];
        columnGroup.spacing = WINDOW_SPACING;
        columnGroup.margins = 0;
    }

    /**
     * 欄に付ける単位（「20 mm」の「 mm」）を返す
     * @param {{label: string}} unitInfo - 定規の単位
     * @returns {string} 単位の文字列
     */
    function getLengthUnitSuffix(unitInfo) {
        return " " + unitInfo.label;
    }

    /**
     * 長さ（pt）を、定規の単位に換算して単位を付けた欄の文字列にする（小数第2位まで）
     * @param {number} lengthPt - 長さ（pt）
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 定規の単位
     * @returns {string} 欄に表示する文字列（例 "20 mm"）
     */
    function formatLength(lengthPt, unitInfo) {
        return formatStepperNumber(lengthPt / unitInfo.pointsPerUnit) + getLengthUnitSuffix(unitInfo);
    }

    /**
     * 「項目名 : ∧∨[欄]」の長さの欄（幅・高さ）を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 項目名の LABELS パス
     * @param {number} labelWidth - 項目名の幅（右揃え）
     * @param {number} valuePt - 初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 定規の単位
     * @returns {EditText} 追加した欄
     */
    function addLengthField(parentPanel, labelPath, labelWidth, valuePt, unitInfo) {
        var valueInput = addSteppedField(parentPanel, {
            label: labelText(labelPath), labelWidth: labelWidth,
            text: formatLength(valuePt, unitInfo), characters: 8, step: 1, min: 0.01, unit: getLengthUnitSuffix(unitInfo)
        });
        valueInput.helpTip = getLabel("tooltip.frameSize");
        valueInput.pointsPerUnit = unitInfo.pointsPerUnit;
        return valueInput;
    }

    /**
     * 「□チェックボックス ∧∨[欄]」の行（角丸の半径・線幅）を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} checkboxLabelPath - チェックボックスの LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} isChecked - チェックボックスの初期値
     * @param {number} valuePt - 欄の初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} lengthUnitInfo - 欄の単位
     * @param {boolean} [placesFieldBelow] - true なら欄をチェックボックスの次の行に字下げして置く
     * @returns {{row: Group, checkbox: Checkbox, input: EditText}} 行（欄を次の行に置くときは2行をまとめたグループ）とチェックボックスと欄
     */
    function addCheckboxLengthRow(parentPanel, checkboxLabelPath, tooltipPath, isChecked, valuePt, lengthUnitInfo, placesFieldBelow) {
        var checkboxRow = parentPanel.add("group");
        if (placesFieldBelow) {
            checkboxRow.orientation = "column";
            checkboxRow.alignChildren = ["left", "top"];
            checkboxRow.spacing = 6;
        } else {
            setupRow(checkboxRow, "left", 6);
        }
        var rowCheckbox = checkboxRow.add("checkbox", undefined, getLabel(checkboxLabelPath));
        rowCheckbox.helpTip = getLabel(tooltipPath);
        rowCheckbox.value = isChecked === true;
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = checkboxRow.add("group");
        setupRow(stepperInputGroup, "left", 0);
        if (placesFieldBelow) stepperInputGroup.margins = [SUB_OPTION_INDENT, 0, 0, 0];
        var valueInput;
        var rowStepper = addStepper(stepperInputGroup, function () { return valueInput; },
            { step: 1, min: 0, unit: getLengthUnitSuffix(lengthUnitInfo) });
        valueInput = stepperInputGroup.add("edittext", undefined, formatLength(valuePt, lengthUnitInfo));
        valueInput.characters = 6;
        valueInput.helpTip = getLabel(tooltipPath);
        valueInput.lastValidText = valueInput.text;
        valueInput.pointsPerUnit = lengthUnitInfo.pointsPerUnit;
        bindSteppedArrowKeys(valueInput, rowStepper);
        return { row: checkboxRow, checkbox: rowCheckbox, input: valueInput };
    }

    /**
     * チェックボックスを tooltip 付きで追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} isEnabled - 操作できるかどうか
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, isEnabled, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue === true;
        optionCheckbox.enabled = isEnabled;
        return optionCheckbox;
    }

    main();

})();
