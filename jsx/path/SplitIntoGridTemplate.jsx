#target illustrator
#targetengine "SplitIntoGridTemplateEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパスを、外接矩形を基準に行×列の格子で等分割します。行数・列数と間隔はダイアログで指定し、左右・上下の分割ボタンでも入れられます。

### Overview

Splits the selected paths into an equal grid of rows and columns across their bounding box. Set the rows, columns, and gutters in the dialog, or fill them in with the left/right and top/bottom split buttons.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitIntoGridTemplate";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_ROWS = 1;                /* 行数の初期値 / initial number of rows */
    var DEFAULT_COLUMNS = 2;             /* 列数の初期値 / initial number of columns */
    var MAX_DIVISIONS = 20;              /* 行数・列数の上限 / upper limit for rows and columns */
    var DEFAULT_USE_GUTTER = false;      /* ［間隔を空ける］の初期値 / initial state of Add Gutters */
    var DEFAULT_ROW_GUTTER = 0;          /* 行間の初期値（定規の単位）/ initial row gutter (ruler units) */
    var DEFAULT_COLUMN_GUTTER = 0;       /* 列間の初期値（定規の単位）/ initial column gutter (ruler units) */
    var DEFAULT_LINK_GUTTERS = true;     /* 行間・列間の連動の初期値 / initial state of the gap link */
    var DEFAULT_GUIDE_EDGE = false;      /* ガイドの［エッジ］の初期値 / initial state of guide Edges */
    var DEFAULT_GUIDE_CENTER_V = false;  /* ガイドの［中心（縦）］の初期値 / initial state of guide Vertical Center */
    var DEFAULT_GUIDE_CENTER_H = false;  /* ガイドの［中心（横）］の初期値 / initial state of guide Horizontal Center */
    var DEFAULT_GUIDE_EXTENSION = 0;     /* ガイドの伸張の初期値（定規の単位）/ initial guide extension (ruler units) */
    var DEFAULT_LIVE_SHAPE = true;       /* ［ライブシェイプにする］の初期値 / initial state of Convert to Live Shape */
    var DEFAULT_SHOW_CENTER = true;      /* ［中心点を表示］の初期値 / initial state of Show Center Point */
    var DEFAULT_PREVIEW = true;          /* ［プレビュー］の初期値 / initial state of Preview */
    var CUT_OVERHANG = 1;                /* 切り抜き矩形のはみ出し量（pt）/ overhang of the cutting rectangles (pt) */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_MARGINS = 18;                 /* ダイアログ外周の余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];    /* パネルの余白 [左,上,右,下] / panel margins */
    var PRESET_PIECE_COUNTS = [2, 3, 4];     /* 分割ボタンの分割数 / piece counts on the split buttons */
    var GRID_INPUT_CHARS = 3;                /* 行・列の入力欄の文字数 / characters for the row and column fields */
    var GUTTER_INPUT_CHARS = 5;              /* 間隔の入力欄の文字数 / characters for the gutter fields */
    var TWO_BY_TWO_TOP_MARGIN = 5;           /* ［4分割（2行2列）］の上に足す余白 / extra space above the 2 × 2 button */
    var GRID_LABEL_WIDTH = { ja: 30, en: 60 };   /* 行・列の項目名の幅（言語別）/ row and column label width per language */
    var EXTENSION_INPUT_CHARS = 5;           /* 伸張の入力欄の文字数 / characters for the extension field */
    var GUTTER_LABEL_WIDTH = { ja: 40, en: 80 }; /* 行間・列間の項目名の幅（言語別）/ gap label width per language */

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
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
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は SETTINGS_STORE_* / createSettingsStore / readSettingsLegacyFile / readSettingsLegacyPreference / settingsStore*
    // 2. 寿命は今のスクリプトに合わせて選ぶ。
    //      "session"    … $.global に置く。Illustrator を終了するまで残る。#targetengine が必須（無いと毎回消える）
    //      "persistent" … Folder.userData/illustrator-scripts/<storeName>.json に書く。再起動しても残る
    //    storeName はふつう SCRIPT_NAME。ダイアログの位置は DialogPosition の部品が持つので、ここには入れない
    // 3. 既定値を1か所にまとめ、load で受け取る。戻り値は毎回新しいオブジェクト（書き換えても保存されない）
    //      var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
    //      var DEFAULT_SETTINGS = { widthPt: 10, addFrame: true, modeKey: "fit", corners: { tl: 0, tr: 0 } };
    //      var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
    //      …OK で閉じたら…
    //      settingsStore.save({ widthPt: …, addFrame: …, modeKey: …, corners: { tl: …, tr: … } });
    //    型は既定値に合わせる（数値の既定値には "12" も 12 として読む。真偽は "1"/"0"/"true"/"false" も読む）。
    //    合わない値・既定値に無い項目は捨てて既定値を使う。{} と null の既定値は中身を問わずそのまま受け取る
    //    （名前をキーにしたプリセット集など）。配列は配列ならそのまま受け取る
    // 4. 保存できるのは文字列・数値・真偽・null と、その配列・入れ子のオブジェクトだけ。
    //    DOM オブジェクト・File・関数は入れない（パスは fsName の文字列で持つ）。長さは pt で持つ
    // 5. 旧形式の設定を読み継ぐときは、3つ目の引数に legacy 関数を渡す。
    //    新しい保存が1度も無いとき（ファイルが無い・$.global に無い）だけ呼ばれ、戻り値を保存値として既定値と突き合わせる。
    //    旧ファイル・旧キーは消さない。キー名が変わったときは legacy の中で詰め替える
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");  … key=value / toSource / JSON を自動判別
    //      } });
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyPreference("SmartTextFindReplace/settings");  … app.preferences の文字列
    //      } });
    // 6. clear() は保存を消す。legacy を渡したストアでは空の保存（{}）を書き、旧設定が戻ってこないようにする
    // 7. 失敗は例外にせず、load は既定値、save は false を返す（$.writeln に理由を出す）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            return textFile.read().replace(/^﻿/, "");
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
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // セッション中の記憶 / Session memory
    // =========================================

    /* 常駐エンジン（#targetengine）の $.global に載せ、Illustrator を終了するまで前回の値を残す。旧版の $.global のキーを1度だけ読み継ぐ
       Kept on $.global of the persistent engine, so the last values survive until Illustrator quits; the old $.global key is read once */
    var LEGACY_SESSION_STATE_KEY = "__SplitIntoGridTemplate_State__";
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_STATE_KEY] || null; }
    });

    /**
     * @typedef {Object} DialogState
     * @property {string} rows - 行の入力欄の文字列
     * @property {string} columns - 列の入力欄の文字列
     * @property {boolean} useGutter - ［間隔を空ける］
     * @property {string} rowGutter - 行間の入力欄の文字列（単位付き）
     * @property {string} columnGutter - 列間の入力欄の文字列（単位付き）
     * @property {boolean} linkGutters - 行間・列間の連動
     * @property {boolean} guideEdge - ガイドの［エッジ］
     * @property {boolean} guideCenterVertical - ガイドの［中心（縦）］
     * @property {boolean} guideCenterHorizontal - ガイドの［中心（横）］
     * @property {string} guideExtension - 伸張の入力欄の文字列（単位付き）
     * @property {boolean} liveShape - ［ライブシェイプにする］
     * @property {boolean} showCenter - ［中心点を表示］
     * @property {boolean} preview - ［プレビュー］
     * @property {string} unitLabel - 保存したときの定規の単位
     */

    /**
     * ダイアログの初期状態を返す。前回OKしたときの値があればそれを使い、無い項目は初期値で埋める。
     * 定規の単位が前回と違うときは、長さの欄だけ初期値に戻す（「5 mm」を pt として読まないように）。
     * @param {string} unitLabel - いまの定規の単位
     * @returns {DialogState} 初期状態
     */
    function getInitialDialogState(unitLabel) {
        var unitSuffix = " " + unitLabel;
        var defaults = {
            rows: String(DEFAULT_ROWS),
            columns: String(DEFAULT_COLUMNS),
            useGutter: DEFAULT_USE_GUTTER,
            rowGutter: DEFAULT_ROW_GUTTER + unitSuffix,
            columnGutter: DEFAULT_COLUMN_GUTTER + unitSuffix,
            linkGutters: DEFAULT_LINK_GUTTERS,
            guideEdge: DEFAULT_GUIDE_EDGE,
            guideCenterVertical: DEFAULT_GUIDE_CENTER_V,
            guideCenterHorizontal: DEFAULT_GUIDE_CENTER_H,
            guideExtension: DEFAULT_GUIDE_EXTENSION + unitSuffix,
            liveShape: DEFAULT_LIVE_SHAPE,
            showCenter: DEFAULT_SHOW_CENTER,
            preview: DEFAULT_PREVIEW,
            unitLabel: unitLabel
        };
        var dialogState = settingsStore.load(defaults);
        if (dialogState.unitLabel !== unitLabel) {
            dialogState.rowGutter = defaults.rowGutter;
            dialogState.columnGutter = defaults.columnGutter;
            dialogState.guideExtension = defaults.guideExtension;
            dialogState.unitLabel = unitLabel;
        }
        return dialogState;
    }

    /**
     * ダイアログの状態を次回のために残す。
     * @param {DialogState} dialogState - 残す状態
     * @returns {void}
     */
    function saveDialogState(dialogState) {
        settingsStore.save(dialogState);
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    var LABELS = {
        dialog: {
            title: { ja: "パスを格子状に分割", en: "Split Path into Grid" }
        },
        panel: {
            gridSize: { ja: "行と列", en: "Rows and Columns" },
            preset: { ja: "分割", en: "Split" },
            leftRight: { ja: "左右", en: "Left/Right" },
            topBottom: { ja: "上下", en: "Top/Bottom" },
            gutter: { ja: "間隔", en: "Gutter" },
            guides: { ja: "ガイドを作成", en: "Create Guides" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            rows: { ja: "行", en: "Rows" },
            columns: { ja: "列", en: "Columns" },
            rowGutter: { ja: "行間", en: "Row Gap" },
            columnGutter: { ja: "列間", en: "Column Gap" },
            guideExtension: { ja: "伸張", en: "Extension" }
        },
        checkbox: {
            useGutter: { ja: "間隔を空ける", en: "Add Gutters" },
            guideEdge: { ja: "エッジ", en: "Edges" },
            guideCenterVertical: { ja: "中心（縦）", en: "Vertical Center" },
            guideCenterHorizontal: { ja: "中心（横）", en: "Horizontal Center" },
            liveShape: { ja: "ライブシェイプにする", en: "Convert to Live Shape" },
            showCenter: { ja: "中心点を表示", en: "Show Center Point" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            pieces: { ja: "{n}分割", en: "{n} pieces" },
            twoByTwo: { ja: "4分割（2行2列）", en: "4 pieces (2 × 2)" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            leftRightPreset: { ja: "行を1、列を{n}にする", en: "Set rows to 1 and columns to {n}" },
            topBottomPreset: { ja: "行を{n}、列を1にする", en: "Set rows to {n} and columns to 1" },
            useGutter: { ja: "OFFのときは行間・列間を0として分割", en: "When off, split with no row or column gaps" },
            linkGutters: { ja: "行間と列間を同じ値にそろえる（クリックで切り替え）", en: "Keep the row and column gaps equal (click to toggle)" },
            guideEdge: {
                ja: "分割した各片の四辺にガイドを引く（option＋クリックで、すべてON／これだけON を切り替え）",
                en: "Draw guides along the four sides of each piece (Option-click to switch between all on and only this)"
            },
            guideCenterVertical: {
                ja: "各片の左右中央に垂直のガイドを引く（option＋クリックで、すべてON／これだけON を切り替え）",
                en: "Draw a vertical guide through the center of each piece (Option-click to switch between all on and only this)"
            },
            guideCenterHorizontal: {
                ja: "各片の上下中央に水平のガイドを引く（option＋クリックで、すべてON／これだけON を切り替え）",
                en: "Draw a horizontal guide through the center of each piece (Option-click to switch between all on and only this)"
            },
            guideExtension: { ja: "元のパスの外側へガイドを伸ばす距離", en: "Distance to extend the guides beyond the original path" },
            liveShape: {
                ja: "OK後、［シェイプに変換］で長方形などの片をライブシェイプにします。",
                en: "After OK, turn rectangular and other pieces into live shapes with Convert to Shape."
            },
            showCenter: {
                ja: "分割した各パスの中心点を属性パネルで表示します（OK時に適用）。",
                en: "Show the center point of each split path in the Attributes panel (applied on OK)."
            },
            preview: {
                ja: "分割結果を表示し、元のパスを一時的に隠す",
                en: "Show the split result and temporarily hide the original paths"
            },
            /* ステップボタン用 / for the stepper */
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
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noPath: {
                ja: "パス（または複合パス）を選択してから実行してください。",
                en: "Please select a path (or compound path) before running this script."
            },
            nothingSplit: {
                ja: "分割できませんでした。\n行・列の数と間隔を確認してください。",
                en: "Nothing was split.\nCheck the rows, columns, and gutters."
            },
            failed: { ja: "分割に失敗しました：\n", en: "Failed to split:\n" }
        }
    };

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 分割の指定を受け取るダイアログを表示する。閉じるときにプレビューは必ず消す。
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} targets - 分割するパス
     * @returns {SplitSettings|null} 指定内容。キャンセル時は null
     */
    function showDialog(doc, targets) {
        var unitInfo = getUnitInfo();
        var dlg = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = DIALOG_MARGINS;

        /* 左列に行・列・間隔・オプション、右列に分割パネル / Left column: rows, columns, gutters, options; right column: Split panel */
        var columnsRowGroup = dlg.add("group");
        columnsRowGroup.orientation = "row";
        columnsRowGroup.alignChildren = ["fill", "top"];
        var leftColumnGroup = columnsRowGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];
        var rightColumnGroup = columnsRowGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = ["fill", "top"];

        var dialogState = getInitialDialogState(unitInfo.label);
        var gridInputs = addGridSizeFields(leftColumnGroup, dialogState, refreshPreview);
        var gutterInputs = addGutterPanel(leftColumnGroup, unitInfo, dialogState, refreshPreview);
        var presetButtons = addPresetButtons(rightColumnGroup);

        var guideInputs = addGuidePanel(leftColumnGroup, unitInfo, dialogState, refreshPreview);

        var optionsPanel = leftColumnGroup.add("panel", undefined, getLabel(LABELS.panel.options));
        optionsPanel.orientation = "column";
        optionsPanel.alignChildren = ["left", "top"];
        optionsPanel.margins = PANEL_MARGINS;
        var cbLiveShape = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.liveShape));
        cbLiveShape.value = dialogState.liveShape;
        cbLiveShape.helpTip = getLabel(LABELS.tooltip.liveShape);
        var cbShowCenter = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.showCenter));
        cbShowCenter.value = dialogState.showCenter;
        cbShowCenter.helpTip = getLabel(LABELS.tooltip.showCenter);

        // ボタン行（左：プレビュー／右：キャンセル・OK）/ Button row: preview on the left, Cancel/OK on the right
        var buttonRow = addButtonRow(dlg);
        var cbPreview = buttonRow.leftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        cbPreview.value = dialogState.preview;
        cbPreview.helpTip = getLabel(LABELS.tooltip.preview);

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        var splitPreview = createSplitPreview(doc, targets);

        /* 入力欄とチェックボックスから分割の指定を読む。［間隔を空ける］がOFFなら間隔は0
           Read the split settings from the controls; gutters are 0 while Add Gutters is off */
        function collectSplitSettings() {
            var gutterScale = gutterInputs.useGutter.value ? unitInfo.pointsPerUnit : 0;
            return {
                rows: parseInt(gridInputs.rows.text, 10),
                columns: parseInt(gridInputs.columns.text, 10),
                rowGutter: parseFloat(gutterInputs.row.text) * gutterScale,
                columnGutter: parseFloat(gutterInputs.column.text) * gutterScale,
                guideEdge: guideInputs.edge.value,
                guideCenterVertical: guideInputs.centerVertical.value,
                guideCenterHorizontal: guideInputs.centerHorizontal.value,
                guideExtension: parseFloat(guideInputs.extension.text) * unitInfo.pointsPerUnit,
                liveShape: cbLiveShape.value,
                showCenter: cbShowCenter.value
            };
        }

        /* 次回のために、いまの入力をそのまま控える / Capture the current inputs as they are for next time */
        function collectDialogState() {
            return {
                rows: gridInputs.rows.text,
                columns: gridInputs.columns.text,
                useGutter: gutterInputs.useGutter.value,
                rowGutter: gutterInputs.row.text,
                columnGutter: gutterInputs.column.text,
                linkGutters: gutterInputs.linkToggle.value,
                guideEdge: guideInputs.edge.value,
                guideCenterVertical: guideInputs.centerVertical.value,
                guideCenterHorizontal: guideInputs.centerHorizontal.value,
                guideExtension: guideInputs.extension.text,
                liveShape: cbLiveShape.value,
                showCenter: cbShowCenter.value,
                preview: cbPreview.value,
                unitLabel: unitInfo.label
            };
        }

        /* 行数・列数に合わせて間隔の欄をディム表示にし、プレビューを作り直す
           Dim the unused gap fields for the current rows and columns, then rebuild the preview */
        function refreshPreview() {
            gutterInputs.setGridDivisions(parseInt(gridInputs.rows.text, 10), parseInt(gridInputs.columns.text, 10));
            if (cbPreview.value) {
                splitPreview.show(collectSplitSettings());
            } else {
                splitPreview.clear();
            }
        }

        /* ボタンの値を入力欄へ入れて反映する / Put a button's values into the fields and apply them */
        function applyGridSize(rows, columns) {
            gridInputs.rows.text = String(rows);
            gridInputs.columns.text = String(columns);
            refreshPreview();
        }

        for (var i = 0; i < PRESET_PIECE_COUNTS.length; i++) {
            (function (pieceCount) {
                presetButtons.leftRight[i].onClick = function () { applyGridSize(1, pieceCount); };
                presetButtons.topBottom[i].onClick = function () { applyGridSize(pieceCount, 1); };
            })(PRESET_PIECE_COUNTS[i]);
        }
        presetButtons.twoByTwo.onClick = function () { applyGridSize(2, 2); };
        cbPreview.onClick = refreshPreview;
        dlg.onShow = refreshPreview;

        /* プレビューの後片付けは呼び出し側で目印まで取り消して行う / The caller undoes back to the marker to clean up the preview */
        prepareDialogWindow(dlg, SCRIPT_NAME);
        var dialogResult = dlg.show();
        if (dialogResult !== 1) return null;
        saveDialogState(collectDialogState());
        return collectSplitSettings();
    }

    /**
     * ［分割］パネルを作り、左右・上下の分割ボタンと［4分割（2行2列）］ボタンを入れる。
     * @param {Group} parentColumn - 追加先の列
     * @returns {{leftRight: Array<Button>, topBottom: Array<Button>, twoByTwo: Button}} 作ったボタン
     */
    function addPresetButtons(parentColumn) {
        var presetPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.preset));
        presetPanel.orientation = "column";
        presetPanel.alignChildren = ["fill", "top"];
        presetPanel.margins = PANEL_MARGINS;

        /* 左右・上下のパネルを縦に並べる / Stack the Left/Right and Top/Bottom panels */
        var leftRightButtons = addPresetPanel(presetPanel, LABELS.panel.leftRight, LABELS.tooltip.leftRightPreset);
        var topBottomButtons = addPresetPanel(presetPanel, LABELS.panel.topBottom, LABELS.tooltip.topBottomPreset);
        /* ボタンに margins は無いので、group で包んで上に余白を足す / Buttons have no margins, so wrap it in a group to add space above */
        var twoByTwoGroup = presetPanel.add("group");
        twoByTwoGroup.margins = [0, TWO_BY_TWO_TOP_MARGIN, 0, 0];
        twoByTwoGroup.alignment = ["center", "top"]; /* 幅いっぱいに伸ばさない / do not stretch to full width */
        var btnTwoByTwo = twoByTwoGroup.add("button", undefined, getLabel(LABELS.button.twoByTwo));
        return { leftRight: leftRightButtons, topBottom: topBottomButtons, twoByTwo: btnTwoByTwo };
    }

    /**
     * 分割ボタン（2〜4分割）を縦に並べたパネルを作る。
     * @param {Panel} parentGroup - パネルを置く［分割］パネル
     * @param {Object} titleLabel - パネル名の { ja, en }
     * @param {Object} tooltipLabel - ボタンの説明の { ja, en }（{n} に分割数が入る）
     * @returns {Array<Button>} PRESET_PIECE_COUNTS の順に並んだボタン
     */
    function addPresetPanel(parentGroup, titleLabel, tooltipLabel) {
        var directionPanel = parentGroup.add("panel", undefined, getLabel(titleLabel));
        directionPanel.orientation = "column";
        directionPanel.alignChildren = ["fill", "top"];
        directionPanel.margins = PANEL_MARGINS;

        var buttons = [];
        for (var i = 0; i < PRESET_PIECE_COUNTS.length; i++) {
            var pieceCount = PRESET_PIECE_COUNTS[i];
            var presetButton = directionPanel.add("button", undefined, getLabel(LABELS.button.pieces).replace("{n}", pieceCount));
            presetButton.helpTip = getLabel(tooltipLabel).replace("{n}", pieceCount);
            buttons.push(presetButton);
        }
        return buttons;
    }

    /**
     * ［行と列］パネルに、行・列の数値欄（ステップボタン付き、1〜MAX_DIVISIONS の整数）を追加する。
     * @param {Group} parentColumn - 追加先の列
     * @param {DialogState} dialogState - 初期状態
     * @param {Function} onValueChange - 値が変わったときに呼ぶ関数
     * @returns {{rows: EditText, columns: EditText}} 作った入力欄
     */
    function addGridSizeFields(parentColumn, dialogState, onValueChange) {
        var gridSizePanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.gridSize));
        gridSizePanel.orientation = "column";
        gridSizePanel.alignChildren = ["left", "top"];
        gridSizePanel.margins = PANEL_MARGINS;

        var gridFieldOptions = { labelWidth: GRID_LABEL_WIDTH[uiLang], characters: GRID_INPUT_CHARS, step: 1, min: 1, max: MAX_DIVISIONS, integer: true, onStep: onValueChange };
        var rowsInput = addSteppedField(gridSizePanel, mergeFieldOptions(gridFieldOptions, {
            label: labelText(LABELS.fieldLabel.rows), text: dialogState.rows
        }));
        var columnsInput = addSteppedField(gridSizePanel, mergeFieldOptions(gridFieldOptions, {
            label: labelText(LABELS.fieldLabel.columns), text: dialogState.columns
        }));
        chainOnChange(rowsInput, onValueChange);
        chainOnChange(columnsInput, onValueChange);
        return { rows: rowsInput, columns: columnsInput };
    }

    /**
     * ［間隔］パネル（［間隔を空ける］・行間・列間・連動アイコン）を追加する。連動中は片方を変えるともう片方も同じ値になる。
     * @param {Group} parentColumn - 追加先の列
     * @param {{label: string}} unitInfo - 定規の単位
     * @param {DialogState} dialogState - 初期状態
     * @param {Function} onValueChange - 値が変わったときに呼ぶ関数
     * @returns {{useGutter: Checkbox, row: EditText, column: EditText, setGridDivisions: Function}} 作ったチェックボックスと入力欄、
     *     行数・列数を渡してディム表示を更新する関数
     */
    function addGutterPanel(parentColumn, unitInfo, dialogState, onValueChange) {
        var gutterPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.gutter));
        gutterPanel.orientation = "column";
        gutterPanel.alignChildren = ["left", "top"];
        gutterPanel.margins = PANEL_MARGINS;

        var cbUseGutter = gutterPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.useGutter));
        cbUseGutter.value = dialogState.useGutter;
        cbUseGutter.helpTip = getLabel(LABELS.tooltip.useGutter);

        var unitSuffix = " " + unitInfo.label;
        var gutterFieldOptions = { labelWidth: GUTTER_LABEL_WIDTH[uiLang], characters: GUTTER_INPUT_CHARS, step: 1, min: 0, unit: unitSuffix };
        /* 行間・列間の2行の右に連動アイコンを置く / Put the link icon to the right of the two gap rows */
        var gapFieldsRowGroup = gutterPanel.add("group");
        gapFieldsRowGroup.orientation = "row";
        gapFieldsRowGroup.alignChildren = ["left", "center"];
        var gapFieldsColumnGroup = gapFieldsRowGroup.add("group");
        gapFieldsColumnGroup.orientation = "column";
        gapFieldsColumnGroup.alignChildren = ["left", "top"];

        var rowGutterInput = addSteppedField(gapFieldsColumnGroup, mergeFieldOptions(gutterFieldOptions, {
            label: labelText(LABELS.fieldLabel.rowGutter), text: dialogState.rowGutter,
            onStep: function () { handleGutterChange(rowGutterInput); }
        }));
        var columnGutterInput = addSteppedField(gapFieldsColumnGroup, mergeFieldOptions(gutterFieldOptions, {
            label: labelText(LABELS.fieldLabel.columnGutter), text: dialogState.columnGutter,
            onStep: function () { handleGutterChange(columnGutterInput); }
        }));
        chainOnChange(rowGutterInput, function () { handleGutterChange(rowGutterInput); });
        chainOnChange(columnGutterInput, function () { handleGutterChange(columnGutterInput); });

        var linkToggle = addLinkToggle(gapFieldsRowGroup, dialogState.linkGutters, function () { handleGutterChange(rowGutterInput); });
        linkToggle.helpTip = getLabel(LABELS.tooltip.linkGutters);
        if (linkToggle.value) columnGutterInput.text = rowGutterInput.text;

        /* 行数・列数（行が1なら行間、列が1なら列間は使わない）/ Rows and columns; a single row or column has no gap to set */
        var gridDivisions = { rows: parseInt(dialogState.rows, 10), columns: parseInt(dialogState.columns, 10) };

        /* 使わない間隔の欄をディム表示にする。連動アイコンは両方使うときだけ有効
           Dim the gap fields not in use; the link icon is enabled only when both are */
        function updateGutterEnabled() {
            var isRowGapEnabled = cbUseGutter.value && gridDivisions.rows > 1;
            var isColumnGapEnabled = cbUseGutter.value && gridDivisions.columns > 1;
            var isLinkEnabled = isRowGapEnabled && isColumnGapEnabled;
            /* 描き直しでちらつかないよう、変わったときだけ切り替える / Switch only on change to avoid flicker */
            if (rowGutterInput.enabled !== isRowGapEnabled) setSteppedFieldEnabled(rowGutterInput, isRowGapEnabled);
            if (columnGutterInput.enabled !== isColumnGapEnabled) setSteppedFieldEnabled(columnGutterInput, isColumnGapEnabled);
            setLinkToggleEnabled(linkToggle, isLinkEnabled);
        }

        /* 行数・列数が変わったときに呼ぶ / Call when the rows or columns change */
        function setGridDivisions(rows, columns) {
            gridDivisions.rows = rows;
            gridDivisions.columns = columns;
            updateGutterEnabled();
        }
        updateGutterEnabled();
        cbUseGutter.onClick = function () {
            updateGutterEnabled();
            onValueChange();
        };

        /* 連動中なら、変えた欄の値をもう片方へ写してから反映する / When linked, copy the edited value to the other field first */
        function handleGutterChange(editedInput) {
            if (linkToggle.value) {
                var otherInput = (editedInput === rowGutterInput) ? columnGutterInput : rowGutterInput;
                otherInput.text = editedInput.text;
            }
            onValueChange();
        }

        return { useGutter: cbUseGutter, row: rowGutterInput, column: columnGutterInput, linkToggle: linkToggle, setGridDivisions: setGridDivisions };
    }

    /**
     * ［ガイドを作成］パネル（エッジ・中心（縦）・中心（横）・伸張）を追加する。
     * ガイドの種類が1つも選ばれていない間は、伸張をディム表示にする。
     * @param {Group} parentColumn - 追加先の列
     * @param {{label: string}} unitInfo - 定規の単位
     * @param {DialogState} dialogState - 初期状態
     * @param {Function} onValueChange - 値が変わったときに呼ぶ関数
     * @returns {{edge: Checkbox, centerVertical: Checkbox, centerHorizontal: Checkbox, extension: EditText}} 作ったチェックボックスと入力欄
     */
    function addGuidePanel(parentColumn, unitInfo, dialogState, onValueChange) {
        var guidePanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.guides));
        guidePanel.orientation = "column";
        guidePanel.alignChildren = ["left", "top"];
        guidePanel.margins = PANEL_MARGINS;

        var cbGuideEdge = addGuideCheckbox(guidePanel, LABELS.checkbox.guideEdge, LABELS.tooltip.guideEdge, dialogState.guideEdge);
        /* 中心（縦）・中心（横）は横に並べる / Put the two center options side by side */
        var guideCenterRowGroup = guidePanel.add("group");
        guideCenterRowGroup.orientation = "row";
        guideCenterRowGroup.alignChildren = ["left", "center"];
        var cbGuideCenterVertical = addGuideCheckbox(guideCenterRowGroup, LABELS.checkbox.guideCenterVertical, LABELS.tooltip.guideCenterVertical, dialogState.guideCenterVertical);
        var cbGuideCenterHorizontal = addGuideCheckbox(guideCenterRowGroup, LABELS.checkbox.guideCenterHorizontal, LABELS.tooltip.guideCenterHorizontal, dialogState.guideCenterHorizontal);

        var unitSuffix = " " + unitInfo.label;
        var extensionInput = addSteppedField(guidePanel, {
            label: labelText(LABELS.fieldLabel.guideExtension), text: dialogState.guideExtension,
            characters: EXTENSION_INPUT_CHARS, step: 1, min: 0, unit: unitSuffix, onStep: onValueChange
        });
        extensionInput.helpTip = getLabel(LABELS.tooltip.guideExtension);
        chainOnChange(extensionInput, onValueChange);

        /* ガイドの種類が1つも無ければ伸張は使わない / The extension is unused when no guide type is chosen */
        function updateExtensionEnabled() {
            var isEnabled = cbGuideEdge.value || cbGuideCenterVertical.value || cbGuideCenterHorizontal.value;
            if (extensionInput.enabled !== isEnabled) setSteppedFieldEnabled(extensionInput, isEnabled);
        }
        updateExtensionEnabled();
        var guideCheckboxes = [cbGuideEdge, cbGuideCenterVertical, cbGuideCenterHorizontal];
        for (var i = 0; i < guideCheckboxes.length; i++) {
            guideCheckboxes[i].onClick = (function (clickedCheckbox) {
                return function () {
                    if (ScriptUI.environment.keyboardState.altKey) applySoloToggle(guideCheckboxes, clickedCheckbox);
                    updateExtensionEnabled();
                    onValueChange();
                };
            })(guideCheckboxes[i]);
        }

        return { edge: cbGuideEdge, centerVertical: cbGuideCenterVertical, centerHorizontal: cbGuideCenterHorizontal, extension: extensionInput };
    }

    /**
     * option＋クリックの切り替え。すべてONなら押したものだけONにし、それ以外のときはすべてONにする。
     * @param {Array<Checkbox>} checkboxes - 対象のチェックボックス
     * @param {Checkbox} clickedCheckbox - 押したチェックボックス
     * @returns {void}
     */
    function applySoloToggle(checkboxes, clickedCheckbox) {
        /* クリックで反転する前の状態で「すべてON」かを見る / Judge "all on" by the state before the click flipped it */
        var wasAllOn = true;
        for (var i = 0; i < checkboxes.length; i++) {
            var wasChecked = (checkboxes[i] === clickedCheckbox) ? !checkboxes[i].value : checkboxes[i].value;
            if (!wasChecked) wasAllOn = false;
        }
        for (var j = 0; j < checkboxes.length; j++) {
            checkboxes[j].value = wasAllOn ? (checkboxes[j] === clickedCheckbox) : true;
        }
    }

    /**
     * ガイドの種類のチェックボックスを1つ追加する。
     * @param {Panel|Group} guidePanel - 追加先
     * @param {Object} titleLabel - 項目名の { ja, en }
     * @param {Object} tooltipLabel - 説明の { ja, en }
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 作ったチェックボックス
     */
    function addGuideCheckbox(guidePanel, titleLabel, tooltipLabel, initialValue) {
        var guideCheckbox = guidePanel.add("checkbox", undefined, getLabel(titleLabel));
        guideCheckbox.value = initialValue;
        guideCheckbox.helpTip = getLabel(tooltipLabel);
        return guideCheckbox;
    }

    /**
     * 共通の数値欄オプションに、欄ごとのオプションを重ねた新しいオブジェクトを返す。
     * @param {Object} baseOptions - 共通のオプション
     * @param {Object} fieldOptions - 欄ごとのオプション（同じキーは上書き）
     * @returns {Object} 重ねたオプション
     */
    function mergeFieldOptions(baseOptions, fieldOptions) {
        var merged = {};
        var key;
        for (key in baseOptions) merged[key] = baseOptions[key];
        for (key in fieldOptions) merged[key] = fieldOptions[key];
        return merged;
    }

    /**
     * addSteppedField() が付けた値の整形（onChange）を残したまま、あとに処理を足す。
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {Function} handler - 整形のあとに呼ぶ関数
     * @returns {void}
     */
    function chainOnChange(numberInput, handler) {
        var normalizeValue = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeValue();
            handler();
        };
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. アイコンを addLinkToggle(親, 初期値, 切り替え後の関数) で作る。helpTip はコピー先で付ける
    //      var linkToggle = addLinkToggle(fieldsRowGroup, true, function () { syncFields(); });
    //      linkToggle.helpTip = getLabel(LABELS.tooltip.linkToggle);
    // 3. 連動中かは linkToggle.value で読む。コードから変えるときは setLinkToggleValue(linkToggle, true)
    // 4. 有効／無効は setLinkToggleEnabled(linkToggle, isEnabled)（無効の間はクリックが効かず、薄く描く）
    // 5. 2つの入力欄の右に置くときは、行 group の中に「入力欄を縦に積んだ group」とアイコンを並べ、
    //    行 group の alignChildren を ["left", "center"] にすると上下の中央に来る
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle) {
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = LINK_ICON_SIZE;
        linkToggle.minimumSize = LINK_ICON_SIZE;
        linkToggle.maximumSize = LINK_ICON_SIZE;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = LINK_ICON_SIZE[0];
            var iconHeight = LINK_ICON_SIZE[1];
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
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
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
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * 複製で分割結果を見せ、元のパスを一時的に隠すプレビューを作る。
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} targets - 分割するパス
     * @returns {{show: Function, clear: Function}} show(splitSettings) で作り直し、clear() で消す
     */
    function createSplitPreview(doc, targets) {
        var splitOutput = { pieces: [], guides: [] };
        var hiddenItems = [];

        /* 作った片とガイドを削除し、隠したパスを表示に戻す / Remove the pieces and guides, and unhide the originals */
        function clear() {
            var createdItems = splitOutput.pieces.concat(splitOutput.guides);
            for (var i = 0; i < createdItems.length; i++) createdItems[i].remove();
            for (var j = 0; j < hiddenItems.length; j++) hiddenItems[j].hidden = false;
            splitOutput = { pieces: [], guides: [] };
            hiddenItems = [];
            app.redraw();
        }

        /* 分割結果を作り直す / Rebuild the split result */
        function show(splitSettings) {
            clear();
            /* メニューコマンドが特殊な形で失敗してもダイアログは閉じない / Keep the dialog alive if a menu command fails on an odd shape */
            try {
                for (var i = 0; i < targets.length; i++) {
                    /* 分割の後で隠す（複製は元の hidden を引き継ぐ）/ Hide after splitting, since duplicates inherit hidden */
                    if (!splitItem(doc, targets[i], splitSettings, splitOutput, false)) continue;
                    targets[i].hidden = true;
                    hiddenItems.push(targets[i]);
                }
            } catch (e) {
                clear();
            }
            app.redraw();
        }

        return { show: show, clear: clear };
    }

    // =========================================
    // 属性パネルの［中心を表示］アクション / "Show Center" action
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は runTemporaryAction / loadTemporaryActionSet / unloadTemporaryActionSet / toActionHex / buildActionNameLines
    // 2. アクション定義は配列＋join("\n") で組み立てる（''' は ES3 の構文エラー）。
    //    セット名・アクション名は英数字にする。/name [ n 16進 ] は buildActionNameLines で作るとバイト数がずれない
    //      var actionSource = [
    //          "/version 3"
    //      ].concat(buildActionNameLines("", "MySet"), [
    //          "/isOpen 1", "/actionCount 1", "/action-1 {"
    //      ], buildActionNameLines("\t", "myAction"), [ … ]).join("\n");
    // 3. 1回だけ実行するとき:
    //      if (!runTemporaryAction(actionSource, "MySet", "myAction")) alert(getLabel("alert.actionFailed"));
    //    何度も実行するとき（オブジェクトごとなど）は、読み込み・解除を1回ずつにする:
    //      if (!loadTemporaryActionSet(actionSource, "MySet")) { alert(…); return; }
    //      try { for (…) app.doScript("myAction", "MySet"); } finally { unloadTemporaryActionSet("MySet"); }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。警告を出すかはコピー先で決める
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 選択中のオブジェクトに属性パネルの［中心を表示］を適用する。
     * API で直接設定できないため、一時アクション（.aia）を読み込んで再生する。
     * @returns {void}
     */
    function runShowCenterAction() {
        var actionSetName = "Attribute";
        var actionName = "ShowCenter";
        var actionSource = [
            "/version 3",
            "/name [ 9 417474726962757465 ]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            " /name [ 10 53686f7743656e746572 ]",
            " /keyIndex 0",
            " /colorIndex 0",
            " /isOpen 1",
            " /eventCount 1",
            " /event-1 {",
            "  /useRulersIn1stQuadrant 0",
            "  /internalName (adobe_attributePalette)",
            "  /localizedName [ 12 e5b19ee680a7e8a8ade5ae9a ]",
            "  /isOpen 1",
            "  /isOn 1",
            "  /hasDialog 0",
            "  /parameterCount 1",
            "  /parameter-1 {",
            "   /key 1668183154",
            "   /showInPalette 4294967295",
            "   /type (boolean)",
            "   /value 1",
            "  }",
            " }",
            "}"
        ].join("\n");

        /* 失敗は従来どおり例外で伝える / Report a failure as an exception, as before */
        if (!runTemporaryAction(actionSource, actionSetName, actionName)) {
            throw new Error("Could not run the action: " + actionSetName + " / " + actionName);
        }
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * 選択から分割対象のパス・複合パスを取り出す。
     * ガイド・クリッピングパス・ロック中・非表示のものは除く。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<PageItem>} 対象オブジェクト
     */
    function collectTargets(doc) {
        var targets = [];
        var selection = doc.selection;
        if (!selection || !selection.length) return targets;
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.locked || item.hidden) continue;
            if (item.typename === "PathItem") {
                if (item.guides || item.clipping) continue;
                targets.push(item);
            } else if (item.typename === "CompoundPathItem") {
                targets.push(item);
            }
        }
        return targets;
    }

    // =========================================
    // 分割 / Splitting
    // =========================================

    /**
     * @typedef {Object} SplitSettings
     * @property {number} rows - 行数
     * @property {number} columns - 列数
     * @property {number} rowGutter - 行間（pt）
     * @property {number} columnGutter - 列間（pt）
     * @property {boolean} guideEdge - 各片の四辺にガイドを引くか
     * @property {boolean} guideCenterVertical - 各片の左右中央に垂直のガイドを引くか
     * @property {boolean} guideCenterHorizontal - 各片の上下中央に水平のガイドを引くか
     * @property {number} guideExtension - ガイドを外側へ伸ばす距離（pt）
     * @property {boolean} liveShape - ライブシェイプにするか
     * @property {boolean} showCenter - 中心点を表示するか
     */

    /**
     * @typedef {Object} GridCell
     * @property {number} left - セルの左端（pt）
     * @property {number} top - セルの上端（pt）
     * @property {number} right - セルの右端（pt）
     * @property {number} bottom - セルの下端（pt）
     * @property {{left: number, top: number, width: number, height: number}} cutRect - 分割に使う矩形（外周ははみ出させる）
     */

    /**
     * 外接矩形を行×列に分けたときの1セルの幅と高さを返す（間隔を差し引く）。
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {SplitSettings} splitSettings - 分割の指定
     * @returns {{width: number, height: number}} セルの大きさ（pt）。0以下なら分割できない
     */
    function computeCellSize(bounds, splitSettings) {
        return {
            width: (bounds[2] - bounds[0] - splitSettings.columnGutter * (splitSettings.columns - 1)) / splitSettings.columns,
            height: (bounds[1] - bounds[3] - splitSettings.rowGutter * (splitSettings.rows - 1)) / splitSettings.rows
        };
    }

    /**
     * 外接矩形を行×列に分けたセルを、左上から行ごとに返す。
     * cutRect は分割に使う矩形で、外周のセルだけ外側へはみ出させて取りこぼしを防ぐ。
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {SplitSettings} splitSettings - 分割の指定
     * @returns {Array<GridCell>} セル。分割できなければ空
     */
    function computeGridCells(bounds, splitSettings) {
        var rows = splitSettings.rows;
        var columns = splitSettings.columns;
        var cellSize = computeCellSize(bounds, splitSettings);
        var gridCells = [];
        if (rows * columns <= 1 || cellSize.width <= 0 || cellSize.height <= 0) return gridCells;

        for (var r = 0; r < rows; r++) {
            var cellTop = bounds[1] - (cellSize.height + splitSettings.rowGutter) * r;
            var topOverhang = (r === 0) ? CUT_OVERHANG : 0;
            var bottomOverhang = (r === rows - 1) ? CUT_OVERHANG : 0;
            for (var c = 0; c < columns; c++) {
                var cellLeft = bounds[0] + (cellSize.width + splitSettings.columnGutter) * c;
                var leftOverhang = (c === 0) ? CUT_OVERHANG : 0;
                var rightOverhang = (c === columns - 1) ? CUT_OVERHANG : 0;
                gridCells.push({
                    left: cellLeft,
                    top: cellTop,
                    right: cellLeft + cellSize.width,
                    bottom: cellTop - cellSize.height,
                    cutRect: {
                        left: cellLeft - leftOverhang,
                        top: cellTop + topOverhang,
                        width: cellSize.width + leftOverhang + rightOverhang,
                        height: cellSize.height + topOverhang + bottomOverhang
                    }
                });
            }
        }
        return gridCells;
    }

    /**
     * グループを再帰的にほどき、中身をグループの位置へ出してグループを削除する。
     * @param {PageItem} item - 展開結果のオブジェクト
     * @param {Array<PageItem>} results - 取り出したオブジェクトを積む配列
     * @returns {void}
     */
    function releaseGroup(item, results) {
        if (item.typename !== "GroupItem") {
            results.push(item);
            return;
        }
        while (item.pageItems.length > 0) {
            var child = item.pageItems[0];
            child.move(item, ElementPlacement.PLACEBEFORE);
            releaseGroup(child, results);
        }
        item.remove();
    }

    /**
     * 複製をすべてのセルの矩形でまとめて分割し、セルに入る片だけを残す。
     * セルの数にかかわらずメニューコマンドは［分割］と［アピアランスを分割］の2回で済む。
     * 矩形は塗り・線なしで複製の背面に置くので、パスの中の面は複製の塗り・線を引き継ぎ、
     * パスの外に余った矩形の面は塗り・線なしになる。
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} source - 元のパス
     * @param {Array<GridCell>} gridCells - セル
     * @param {Array<PageItem>} pieces - できた片を積む配列
     * @returns {void}
     */
    function divideIntoCells(doc, source, gridCells, pieces) {
        var duplicate = source.duplicate(source, ElementPlacement.PLACEBEFORE);
        var divideGroup = duplicate.parent.groupItems.add();
        divideGroup.move(duplicate, ElementPlacement.PLACEBEFORE);
        duplicate.move(divideGroup, ElementPlacement.PLACEATBEGINNING);
        for (var i = 0; i < gridCells.length; i++) {
            var cutRect = gridCells[i].cutRect;
            var cutter = duplicate.layer.pathItems.rectangle(cutRect.top, cutRect.left, cutRect.width, cutRect.height);
            cutter.filled = false;
            cutter.stroked = false;
            cutter.move(divideGroup, ElementPlacement.PLACEATEND);
        }

        doc.selection = null;
        divideGroup.selected = true;
        app.executeMenuCommand("Live Pathfinder Divide");
        app.executeMenuCommand("expandStyle");

        var dividedItems = [];
        var expanded = doc.selection;
        for (var j = 0; j < expanded.length; j++) {
            releaseGroup(expanded[j], dividedItems);
        }
        /* 矩形の余り（塗り・線なし）と、間隔の部分にできた片を捨てる / Drop leftover rectangle faces and pieces in the gutters */
        for (var k = 0; k < dividedItems.length; k++) {
            if (isUnpainted(dividedItems[k]) || !isInsideAnyCell(dividedItems[k], gridCells)) {
                dividedItems[k].remove();
            } else {
                pieces.push(dividedItems[k]);
            }
        }
    }

    /**
     * 塗りも線も無いオブジェクトかを判定する（複合パスは最初のパスで見る）。
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} 塗りも線も無ければ true
     */
    function isUnpainted(item) {
        var paintSource = item;
        if (item.typename === "CompoundPathItem") {
            if (item.pathItems.length === 0) return true;
            paintSource = item.pathItems[0];
        }
        return !paintSource.filled && !paintSource.stroked;
    }

    /**
     * オブジェクトの中心がいずれかのセルに入っているかを判定する。
     * 分割でできた面はセルか間隔のどちらかに収まるので、中心で見分けられる。
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Array<GridCell>} gridCells - セル
     * @returns {boolean} セルに入っていれば true
     */
    function isInsideAnyCell(item, gridCells) {
        var itemBounds = item.geometricBounds;
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        for (var i = 0; i < gridCells.length; i++) {
            var gridCell = gridCells[i];
            if (centerX >= gridCell.left && centerX <= gridCell.right && centerY <= gridCell.top && centerY >= gridCell.bottom) return true;
        }
        return false;
    }

    /**
     * ガイドを引く位置を、垂直（X）と水平（Y）に分けて返す。同じ位置は1本にまとめる。
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {SplitSettings} splitSettings - 分割の指定
     * @returns {{xs: number[], ys: number[]}} 垂直ガイドのX座標と水平ガイドのY座標
     */
    function computeGuidePositions(bounds, splitSettings) {
        var cellSize = computeCellSize(bounds, splitSettings);
        var guidePositions = { xs: [], ys: [] };
        for (var c = 0; c < splitSettings.columns; c++) {
            var cellLeft = bounds[0] + (cellSize.width + splitSettings.columnGutter) * c;
            if (splitSettings.guideEdge) {
                addUniquePosition(guidePositions.xs, cellLeft);
                addUniquePosition(guidePositions.xs, cellLeft + cellSize.width);
            }
            if (splitSettings.guideCenterVertical) addUniquePosition(guidePositions.xs, cellLeft + cellSize.width / 2);
        }
        for (var r = 0; r < splitSettings.rows; r++) {
            var cellTop = bounds[1] - (cellSize.height + splitSettings.rowGutter) * r;
            if (splitSettings.guideEdge) {
                addUniquePosition(guidePositions.ys, cellTop);
                addUniquePosition(guidePositions.ys, cellTop - cellSize.height);
            }
            if (splitSettings.guideCenterHorizontal) addUniquePosition(guidePositions.ys, cellTop - cellSize.height / 2);
        }
        return guidePositions;
    }

    /**
     * 座標を配列に足す。ほぼ同じ位置（0.001pt 未満の差）がすでにあれば足さない。
     * @param {number[]} positions - 座標の配列
     * @param {number} position - 足す座標
     * @returns {void}
     */
    function addUniquePosition(positions, position) {
        for (var i = 0; i < positions.length; i++) {
            if (Math.abs(positions[i] - position) < 0.001) return;
        }
        positions.push(position);
    }

    /**
     * 元のパスの外接矩形に、伸張の分だけ伸ばしたガイドを引く。
     * @param {Layer} layer - ガイドを置くレイヤー
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {SplitSettings} splitSettings - 分割の指定
     * @param {Array<PathItem>} guides - 引いたガイドを積む配列
     * @returns {void}
     */
    function drawGuides(layer, bounds, splitSettings, guides) {
        var guidePositions = computeGuidePositions(bounds, splitSettings);
        var extension = splitSettings.guideExtension;
        var i;
        for (i = 0; i < guidePositions.xs.length; i++) {
            var lineX = guidePositions.xs[i];
            guides.push(addGuideLine(layer, [lineX, bounds[1] + extension], [lineX, bounds[3] - extension]));
        }
        for (i = 0; i < guidePositions.ys.length; i++) {
            var lineY = guidePositions.ys[i];
            guides.push(addGuideLine(layer, [bounds[0] - extension, lineY], [bounds[2] + extension, lineY]));
        }
    }

    /**
     * ガイド線を1本追加する（塗り・線なしのパスをガイドにする）。
     * @param {Layer} layer - 追加先レイヤー
     * @param {number[]} startPoint - 始点 [x, y]
     * @param {number[]} endPoint - 終点 [x, y]
     * @returns {PathItem} 追加したガイド
     */
    function addGuideLine(layer, startPoint, endPoint) {
        var guideLine = layer.pathItems.add();
        guideLine.setEntirePath([startPoint, endPoint]);
        guideLine.stroked = false;
        guideLine.filled = false;
        guideLine.guides = true;
        return guideLine;
    }

    /**
     * パスを外接矩形を基準に行×列の格子で分割し、指定があればガイドも引く。確定実行のときだけ元のパスを削除する。
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} source - 元のパス
     * @param {SplitSettings} splitSettings - 分割の指定
     * @param {{pieces: Array<PageItem>, guides: Array<PathItem>}} splitOutput - できた片とガイドを積む先
     * @param {boolean} isCommitRun - 確定実行なら true、プレビューなら false
     * @returns {boolean} 分割したら true（1行1列や間隔が大きすぎるときは何もせず false）
     */
    function splitItem(doc, source, splitSettings, splitOutput, isCommitRun) {
        var bounds = source.geometricBounds;
        var gridCells = computeGridCells(bounds, splitSettings);
        if (gridCells.length === 0) return false;
        divideIntoCells(doc, source, gridCells, splitOutput.pieces);
        drawGuides(source.layer, bounds, splitSettings, splitOutput.guides);
        if (isCommitRun) source.remove();
        return true;
    }

    // =========================================
    // ヒストリー汚染対策 / Undo history cleanup
    // =========================================

    /* プレビューは［分割］などのメニューコマンドを使うため、取り消しの段数を数えて戻すことができない。
       ダイアログを開く前に目印を置き、閉じたら目印が消えるまで取り消して、プレビューの履歴を残さない
       Preview uses menu commands, so its undo steps cannot be counted. Place a marker before the dialog
       and undo until the marker is gone, leaving no preview steps in the history */
    var HISTORY_MARKER_NAME = "__SplitIntoGridTemplate_HistoryMarker__";
    var MAX_ROLLBACK_UNDOS = 5000; /* 取り消しの上限（無限ループ防止）/ cap on undos to avoid an endless loop */

    /**
     * 取り消しの目印になる非表示のパスを置く。前後で redraw して、目印だけで1段の履歴にする。
     * @param {Layer} layer - 目印を置くレイヤー（対象と同じ、ロックされていないレイヤー）
     * @returns {void}
     */
    function placeHistoryMarker(layer) {
        app.redraw();
        var historyMarker = layer.pathItems.add();
        historyMarker.name = HISTORY_MARKER_NAME;
        historyMarker.hidden = true;
        app.redraw();
    }

    /**
     * 目印がドキュメントに残っているかを判定する。
     * @param {Document} doc - 対象ドキュメント
     * @returns {boolean} 残っていれば true
     */
    function hasHistoryMarker(doc) {
        /* getByName は見つからないと例外になる / getByName throws when nothing matches */
        try {
            doc.pageItems.getByName(HISTORY_MARKER_NAME);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 目印が消えるまで取り消し、ダイアログを開く前の状態と履歴に戻す。
     * 取り消しきれなかったときは、目印だけ直接削除する。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function rollbackToHistoryMarker(doc) {
        for (var i = 0; i < MAX_ROLLBACK_UNDOS && hasHistoryMarker(doc); i++) {
            /* 取り消す履歴が尽きると例外になる / app.undo() throws when the history runs out */
            try {
                app.undo();
            } catch (e) {
                break;
            }
        }
        if (hasHistoryMarker(doc)) doc.pageItems.getByName(HISTORY_MARKER_NAME).remove();
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで分割の指定を受け取り、選択したパスを順に分割する。
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} targets - 分割するパス
     * @returns {void}
     */
    function runSplit(doc, targets) {
        placeHistoryMarker(targets[0].layer);
        var splitSettings = showDialog(doc, targets);
        /* プレビューで積んだ履歴を目印ごと取り消す（OKでもキャンセルでも）/ Undo the preview history along with the marker */
        rollbackToHistoryMarker(doc);
        if (!splitSettings) return;

        /* 取り消しで古い参照が使えなくなることがあるので、戻った選択から対象を取り直す
           Undo can invalidate old references, so re-collect the targets from the restored selection */
        targets = collectTargets(doc);
        if (targets.length === 0) return;

        var splitOutput = { pieces: [], guides: [] };
        /* メニューコマンドが特殊な形で失敗することがある / Menu commands can fail on odd shapes */
        try {
            for (var i = 0; i < targets.length; i++) {
                splitItem(doc, targets[i], splitSettings, splitOutput, true);
            }
        } catch (e) {
            alert(getLabel(LABELS.alert.failed) + e);
        }

        doc.selection = null;
        var pieces = splitOutput.pieces;
        if (pieces.length === 0) {
            alert(getLabel(LABELS.alert.nothingSplit));
            return;
        }
        /* ガイドは選択しない（後処理の対象にしない）/ Leave the guides unselected so post-processing skips them */
        for (var j = 0; j < pieces.length; j++) {
            pieces[j].selected = true;
        }
        /* 片を選択したまま後処理をかける / Post-process while the pieces stay selected */
        if (splitSettings.liveShape) app.executeMenuCommand("Convert to Shape");
        if (splitSettings.showCenter) runShowCenterAction();
    }

    /**
     * 対象を確かめ、境界線を隠した状態で分割を実行する。
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        var targets = collectTargets(doc);
        if (targets.length === 0) {
            alert(getLabel(LABELS.alert.noPath));
            return;
        }

        /* ［境界線を隠す］はトグルなので、キャンセルや例外でも必ずもう一度実行して戻す
           Hide Edges is a toggle, so run it again on every exit, including cancel and errors */
        app.executeMenuCommand("edge");
        try {
            runSplit(doc, targets);
        } finally {
            app.executeMenuCommand("edge");
        }
    }

    main();

})();
