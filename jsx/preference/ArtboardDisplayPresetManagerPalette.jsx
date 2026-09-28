#target illustrator
#targetengine "ArtboardDisplayPresetManagerPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード関連のIllustrator環境設定を、常駐パレットでまとめて切り替えます。
アートボードのサイズ確認とリサイズ、名前表示と枠線、プリセット、カンバスカラーの切り替えに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardDisplayPresetManagerPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9eba8ab03170

### Overview

A persistent palette for switching the artboard-related Illustrator preferences.
It covers checking and resizing artboards, name and border display, presets, and the canvas color.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardDisplayPresetManagerPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardDisplayPresetManagerPalette"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardDisplayPresetManagerPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardDisplayPresetManagerPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9eba8ab03170"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 枠線カラーの選択肢（ドロップダウンの並び順。colorKey は LABELS.dropdown.borderColor のキー）
       Border color choices (dropdown order; colorKey is a key of LABELS.dropdown.borderColor) */
    var BORDER_COLOR_PRESETS = [
        { colorKey: "lightBlue",  r: 0.29, g: 0.52, b: 1.0 },
        { colorKey: "lightRed",   r: 1.0,  g: 0.29, b: 0.29 },
        { colorKey: "green",      r: 0.0,  g: 0.65, b: 0.31 },
        { colorKey: "mediumBlue", r: 0.0,  g: 0.45, b: 0.78 },
        { colorKey: "magenta",    r: 1.0,  g: 0.0,  b: 1.0 },
        { colorKey: "cyan",       r: 0.0,  g: 1.0,  b: 1.0 },
        { colorKey: "lightGray",  r: 0.65, g: 0.65, b: 0.65 },
        { colorKey: "black",      r: 0.0,  g: 0.0,  b: 0.0 },
        { colorKey: "yellow",     r: 1.0,  g: 1.0,  b: 0.0 }
    ];

    /* 枠線の太さの選択肢（1〜4）/ Border width choices (1-4) */
    var BORDER_WIDTH_CHOICES = [1, 2, 3, 4];

    /* 表示プリセット（適用・判定の両方で使う単一の真実）/ Display presets (single source for both apply and detect) */
    /* "default" は ES3 予約語のため引用符付きキー / "default" is an ES3 reserved word, so quote the key */
    /* printBleed は現在未使用（PRINT_BLEED_WIDGET を参照）/ printBleed is currently unused (see the PRINT_BLEED_WIDGET note) */
    var PRESET_KEYS = ["default", "emphasis", "light"];
    var PRESETS = {
        "default": { showName: true,  colorKey: "black",     borderWidth: 1, printBleed: true,  moveLocked: false },
        emphasis:  { showName: false, colorKey: "lightRed",  borderWidth: 3, printBleed: false, moveLocked: true },
        light:     { showName: false, colorKey: "lightGray", borderWidth: 1, printBleed: false, moveLocked: true }
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;               /* パネル内の要素間隔 / panel spacing */
    var SIZE_LABEL_WIDTH = 48;             /* 幅・高さラベルの幅（px）/ width of the width/height labels */

    /* 9軸ウィジェット / 9-axis anchor widget */
    var DEFAULT_ANCHOR_INDEX = 0;          /* 既定の基準点（0=左上〜8=右下の行優先）/ default anchor (row-major, 0=top-left..8=bottom-right) */

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
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
     * 行グループの共通設定を適用する（ボタン列・ラベル＋入力欄など）
     * statictext は edittext より低く ScriptUI(mac) では上寄せになりやすいため、天地中央を明示する
     * @param {Group} rowGroup - 対象グループ
     * @param {string} [alignment] - グループ自身の整列（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）
     * @param {Button} button - 対象ボタン
     * @param {number} px - 詰める高さ（px）
     * @returns {void}
     */
    function trimButtonHeight(button, px) {
        button.size = [button.size.width, button.size.height - px];
    }

    /**
     * ↑↓キーを離したときに onChange を呼ぶ（押している間のキーリピートでは確定しない）
     * Illustrator にはタイマーが無いため、キーを離すまでを待ち時間の代わりにする
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function commitOnArrowKeyUp(editText) {
        /* 押している間は目印を立て、∧∨の onStep で確定させない / flag the key as held so onStep does not commit */
        editText.addEventListener("keydown", function (event) {
            if (event.keyName == "Up" || event.keyName == "Down") editText.isArrowKeyDown = true;
        });
        editText.addEventListener("keyup", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            editText.isArrowKeyDown = false;
            if (typeof editText.onChange === "function") editText.onChange();
        });
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
    // 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は ANCHOR_WIDGET_* / *AnchorWidget* / getAnchor* の名前。貼る前に、コピー先にある旧版の
    //    ANCHOR_WIDGET_SIZE・ANCHOR_CELL_*・ANCHOR_CONNECTIONS・ANCHOR_*_COLOR / FILL・addAnchorWidget・drawAnchorWidget・
    //    drawAnchorCell・redrawAnchorWidget・initAnchorColors（9軸の配色だけを決めているもの）・clampGridIndex を消す
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. ウィジェットを作る。初期値は 0〜8（0=左上, 4=中央, 8=右下）か名前（"topLeft" / "top" / "topRight" /
    //    "left" / "center" / "right" / "bottomLeft" / "bottom" / "bottomRight"）
    //      var anchorWidget = addAnchorWidget(anchorPanel, "center", function (anchorIndex) { updatePreview(); });
    //      anchorWidget.helpTip = getLabel(LABELS.tooltip.anchor);
    //    onChange はクリックのたびに呼ぶ（同じセルでも呼ぶ）。setAnchorWidgetValue() からは呼ばない
    //    未選択（-1）を許すときは addAnchorWidget(parent, -1, onChange, { allowNone: true })
    //    選べないセルは { disabledCells: [4] } か setAnchorWidgetCellsDisabled(anchorWidget, [4])（薄く描き、クリックも無視）
    // 3. 値を読む: getAnchorWidgetIndex(anchorWidget) … 0〜8（未選択は -1）/ getAnchorWidgetName(anchorWidget) … "topLeft" など
    //    値を書く: setAnchorWidgetValue(anchorWidget, 2) または setAnchorWidgetValue(anchorWidget, "topRight")（描き直す）
    //    旧版の widget.selectedAnchorIndex / widget.anchorIndex への直接代入は描き直されないので使わない
    // 4. Illustrator の変形に渡す:
    //      pageItem.resize(150, 150, true, true, true, true, 150, getAnchorTransformation(getAnchorWidgetIndex(anchorWidget)));
    //    座標で使うときは getAnchorPointOnBounds(geometricBounds, anchorIndex) → [x, y]、
    //    割合で使うときは getAnchorRatio(anchorIndex) → [0|0.5|1, 0|0.5|1]（左上が [0, 0]）、
    //    シンボル登録は getAnchorSymbolRegistrationPoint(anchorIndex)
    // 5. 有効／無効は setAnchorWidgetEnabled(anchorWidget, isEnabled)（薄い色で描き直し、クリックも無視）。
    //    パネル・行など親の enabled を切り替えたときは、そのあとで redrawAnchorWidgetsIn(親) を呼ぶ
    //    （親の無効化は子の enabled に出ないので、描画とクリックの判定は親までたどる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    // -----------------------------------------
    // 基準点ウィジェットの寸法 / Anchor widget metrics
    // -----------------------------------------
    var ANCHOR_WIDGET_SIZE      = 66;   /* ウィジェット全体の一辺 / overall size of the widget */
    var ANCHOR_WIDGET_CELL_SIZE = 9;    /* □1個の一辺 / size of one square */
    var ANCHOR_WIDGET_CELL_GAP  = 7.5;  /* □どうしの間隔 / gap between squares */
    var ANCHOR_WIDGET_NONE      = -1;   /* 未選択のインデックス / index while nothing is selected */

    /* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
       Cell names in row-major order, matching the Transformation enumeration */
    var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線 / Rules joining the outer squares (the center stands alone) */
    var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    // -----------------------------------------
    // 基準点ウィジェットの配色 / Anchor widget colors
    // -----------------------------------------
    var ANCHOR_WIDGET_UI_DARK = isDarkUI();
    /* 枠線・ケイ線はグレー、選択セルの塗りはライトで濃いグレー・ダークで明るいグレー（既存スクリプトの配色を踏襲）。
       無効時は同じ色を半透明にして背景へ沈める（不透明の薄いグレーだとダークUIで逆に明るく浮くため）
       Gray rules; the selected fill is dark gray on light UI and light gray on dark UI (as in the existing scripts).
       Disabled colors are translucent versions so they sink into any background */
    var ANCHOR_WIDGET_LINE_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.55, 0.55, 0.55, 1]   : [0.6, 0.6, 0.6, 1];  /* 枠線・ケイ線 / rules */
    var ANCHOR_WIDGET_FILL_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.8, 0.8, 0.8, 1]      : [0.4, 0.4, 0.4, 1];  /* 選択セルの塗り / selected fill */
    var ANCHOR_WIDGET_DIM_LINE_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.55, 0.55, 0.55, 0.4] : [0.6, 0.6, 0.6, 0.4];  /* 無効時の枠線 / rules when disabled */
    var ANCHOR_WIDGET_DIM_FILL_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.8, 0.8, 0.8, 0.3]    : [0.4, 0.4, 0.4, 0.3];  /* 無効時の塗り / fill when disabled */

    // -----------------------------------------
    // ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 基準点（3×3）を選ぶウィジェットを追加する。クリックしたセルを選び、onChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
     * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
     * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）/ size（一辺。既定 66）
     * @returns {Button} ウィジェット（値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
     */
    function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
        var anchorOptions = widgetOptions || {};
        var widgetSize = anchorOptions.size || ANCHOR_WIDGET_SIZE;
        var anchorWidget = parent.add("button", undefined, "");
        anchorWidget.minimumSize = [widgetSize, widgetSize];
        anchorWidget.preferredSize = [widgetSize, widgetSize];
        anchorWidget.maximumSize = [widgetSize, widgetSize];
        anchorWidget.isAnchorWidget = true; /* redrawAnchorWidgetsIn() の目印 / marker for redrawAnchorWidgetsIn() */
        anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(anchorOptions.disabledCells);
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(initialValue, anchorWidget.anchorAllowNone);
        anchorWidget.onDraw = function () { drawAnchorWidget(anchorWidget); };
        anchorWidget.onClick = function () {}; /* セルの判定は mousedown で行う / hit-testing happens in mousedown */

        /* クリック座標（コントロール基準）を3分割してセルを判定する / split the control-relative click into thirds */
        anchorWidget.addEventListener("mousedown", function (event) {
            if (!isAnchorWidgetEnabledInTree(anchorWidget)) return;
            var cellIndex = getAnchorCellAt(event.clientX, event.clientY, anchorWidget.size[0], anchorWidget.size[1]);
            if (anchorWidget.anchorDisabledCells[cellIndex]) return;
            anchorWidget.anchorWidgetIndex = cellIndex;
            redrawAnchorWidget(anchorWidget);
            if (onChange) onChange(cellIndex, anchorWidget);
        });
        return anchorWidget;
    }

    /**
     * 選択中のセルのインデックスを返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {number} 0〜8（行優先）。未選択なら -1
     */
    function getAnchorWidgetIndex(anchorWidget) {
        return anchorWidget.anchorWidgetIndex;
    }

    /**
     * 選択中のセルの名前を返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {string} "topLeft" など。未選択なら ""
     */
    function getAnchorWidgetName(anchorWidget) {
        return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
    }

    /**
     * 選択するセルを変えて描き直す（onChange は呼ばない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * ウィジェットの有効／無効を切り替えて描き直す（無効の間は薄く描き、クリックも無視する）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
        anchorWidget.enabled = isEnabled;
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * 選べないセルを指定し直して描き直す（選択中のセルは変えない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
     * @returns {void}
     */
    function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(disabledCells);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * コンテナ以下にある基準点ウィジェットをすべて描き直す。パネルや行の enabled を切り替えたあとに呼ぶ
     * @param {Object} container - パネル・グループ・ウィンドウなど
     * @returns {void}
     */
    function redrawAnchorWidgetsIn(container) {
        if (container.isAnchorWidget) {
            redrawAnchorWidget(container);
            return;
        }
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            redrawAnchorWidgetsIn(container.children[i]);
        }
    }

    // -----------------------------------------
    // 値の変換 / Value helpers
    // -----------------------------------------
    /**
     * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
     * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
     * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
     * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
     */
    function resolveAnchorWidgetIndex(anchorValue, allowNone) {
        if (typeof anchorValue === "string") {
            for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
                if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
            }
            return 4;
        }
        if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
        if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
            return anchorValue;
        }
        return 4;
    }

    /**
     * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [横の割合, 縦の割合]
     */
    function getAnchorRatio(anchorValue) {
        var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    }

    /**
     * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
     * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [x, y]
     */
    function getAnchorPointOnBounds(bounds, anchorValue) {
        var anchorRatio = getAnchorRatio(anchorValue);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
            bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
        ];
    }

    /**
     * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
     * 基準は効果を含まない境界（geometricBounds）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {Transformation} Transformation.TOPLEFT など
     */
    function getAnchorTransformation(anchorValue) {
        var transformations = [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ];
        return transformations[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * symbols.add() に渡す登録点を返す（Illustrator 専用）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
     */
    function getAnchorSymbolRegistrationPoint(anchorValue) {
        var registrationPoints = [
            SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
        ];
        return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * クリック位置からセルのインデックスを求める（ウィジェットを縦横3等分し、外にはみ出した座標は端のセルに寄せる）
     * @param {number} clickX - コントロール基準の x
     * @param {number} clickY - コントロール基準の y
     * @param {number} widgetWidth - ウィジェットの幅
     * @param {number} widgetHeight - ウィジェットの高さ
     * @returns {number} 0〜8
     */
    function getAnchorCellAt(clickX, clickY, widgetWidth, widgetHeight) {
        var column = Math.min(2, Math.max(0, Math.floor(clickX / (widgetWidth / 3))));
        var row = Math.min(2, Math.max(0, Math.floor(clickY / (widgetHeight / 3))));
        return row * 3 + column;
    }

    /**
     * セルのインデックスの配列を、9個の真偽値に直す
     * @param {number[]} [cellIndexes] - セルのインデックスの配列
     * @returns {boolean[]} 含まれるセルだけ true
     */
    function toAnchorCellFlags(cellIndexes) {
        var cellFlags = [false, false, false, false, false, false, false, false, false];
        if (!cellIndexes) return cellFlags;
        for (var i = 0; i < cellIndexes.length; i++) {
            if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
        }
        return cellFlags;
    }

    // -----------------------------------------
    // 描画 / Drawing
    // -----------------------------------------
    /**
     * ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立。選択セルだけ塗る）
     * @param {Button} anchorWidget - 描くウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        var halfCell = cellSize / 2;
        /* 自作描画は自動でディムにならないので、親までたどって判定する / custom drawing is not dimmed automatically */
        var isEnabled = isAnchorWidgetEnabledInTree(anchorWidget);

        /* ボタンの地をコントロールの地色で塗り、パネルに溶け込ませる（backgroundColor が無い環境では例外）
           Paint the control's own background so the widget blends into the panel; throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
        var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);
        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1);
        for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
                graphics.lineTo(cellB[0], cellB[1] + halfCell);
            } else {
                /* 縦方向：下の□へ / vertical: to the square below */
                graphics.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
                graphics.lineTo(cellB[0] + halfCell, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < 9; i++) {
            var isCellEnabled = isEnabled && !anchorWidget.anchorDisabledCells[i];
            drawAnchorWidgetCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.anchorWidgetIndex, isCellEnabled);
        }
    }

    /**
     * □を1つ描く（選択中だけ塗り、枠は塗りの上に重ねる）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} cellX - 左端
     * @param {number} cellY - 上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {boolean} isEnabled - 選べるセルなら true（false なら薄く描く）
     * @returns {void}
     */
    function drawAnchorWidgetCell(graphics, cellX, cellY, isSelected, isEnabled) {
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        /* rectPath の前には毎回 newPath()（呼ばないとパスが累積して塗りが線画になる）
           Always call newPath() before rectPath(), or paths accumulate and fills turn into outlines */
        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(cellX, cellY, cellSize, cellSize);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_FILL_COLOR : ANCHOR_WIDGET_DIM_FILL_COLOR));
        }
        graphics.newPath();
        graphics.rectPath(cellX, cellY, cellSize, cellSize);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isAnchorWidgetEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false) return false;
        }
        return true;
    }

    /**
     * ウィジェットの onDraw を呼び直す。notify("onDraw") は環境によって例外や空振りになるため、隠して再表示して描き直させる
     * @param {Button} anchorWidget - 描き直すウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        anchorWidget.hide();
        anchorWidget.show();
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 常駐エンジン / Persistent engine
    // =========================================

    /* 常駐エンジンにパレット参照を保持するキー / Key holding the palette reference in the persistent engine */
    var PALETTE_GLOBAL_KEY = "__artboardDisplayPresetPalette";

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
     * pt 値を指定単位の表示文字列へ変換する（小数2桁で丸め）
     * @param {number} pointValue - pt 値
     * @param {object} unitInfo - getUnitInfo() の戻り値
     * @returns {string} 表示用文字列（数値でない場合は空文字）
     */
    function pointToUnitText(pointValue, unitInfo) {
        if (isNaN(pointValue)) return "";
        return String(Math.round(pointValue / unitInfo.pointsPerUnit * 100) / 100);
    }

    // =========================================
    // 環境設定アクセス / Preferences access
    // =========================================

    var appPreferences = app.preferences;

    /* -----------------------------------------
       PRINT_BLEED_WIDGET：「裁ち落としを印刷」生成AIボタンの表示切り替えについて

       enablePrintBleedWidget は setBooleanPreference で確実に書き換えられ、環境設定ダイアログの
       表示にも反映されるが、カンバス上のウィジェットは再評価されない。app.redraw()、ズーム、
       ツール切り替え、プレビュー／アウトライン切り替え、アートボード再設定、ドキュメント切り替えの
       いずれでも反映せず、スクリプトから即時反映させる手段が見つからなかった（Illustrator が
       起動時か環境設定ダイアログの適用時にしか読まないと思われる）。

       将来のバージョンで挙動が戻る可能性があるため、関連コードは削除せずコメントアウトで保留する。
       復活させるときは、この定数・LABELS の checkbox.showPrintBleedAI・buildOptionsPanel・
       applyPrintBleedWidgetSetting・applyOptionSettings・applyPresetToUI・findMatchingPresetKey・
       reflectPreferences・wirePaletteEvents のコメントアウトを戻すこと。

       The preference is written reliably, but Illustrator never re-evaluates the canvas widget,
       and no scripted refresh triggers it. The item is parked (commented out) rather than removed.
       ----------------------------------------- */
    // var PRINT_BLEED_WIDGET_KEY = "enablePrintBleedWidget";

    /**
     * 環境設定を型指定で読み取る
     * 未登録のキーや型の食い違いで例外になることがあるため try で受ける
     * @param {string} valueType - "Real" | "Boolean" | "Integer"
     * @param {string} key - 環境設定キー
     * @param {*} fallback - 取得に失敗したときの値
     * @returns {*} 取得値（失敗時は fallback）
     */
    function readPref(valueType, key, fallback) {
        try { return appPreferences["get" + valueType + "Preference"](key); } catch (e) { return fallback; }
    }

    /**
     * メニューコマンドを実行する
     * ドキュメントが無いときなど、実行できない状況の例外は無視する
     * @param {string} command - メニューコマンド名
     * @returns {void}
     */
    function runMenuCommand(command) {
        try { app.executeMenuCommand(command); } catch (e) {}
    }

    /**
     * アートボード表示を強制的に再描画する（環境設定の変更は自動では反映されないため）
     * @returns {void}
     */
    function refreshArtboardDisplay() {
        runMenuCommand("zoomout");
        runMenuCommand("zoomin");
    }

    // =========================================
    // アートボード枠線カラー / Artboard border color
    // =========================================

    /* colorKey が見つからないときに使うブラックの index / Black index used when a colorKey is not found */
    var BORDER_COLOR_BLACK_INDEX = 7;

    /**
     * ドロップダウン用のカラー名配列を生成する
     * @returns {string[]} 現在言語のカラー名の配列
     */
    function buildBorderColorNames() {
        var colorNames = [];
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            colorNames.push(getLabel("dropdown.borderColor." + BORDER_COLOR_PRESETS[i].colorKey));
        }
        return colorNames;
    }

    /**
     * 指定 RGB にもっとも近いカラープリセットの index を返す
     * @param {number} red - 赤成分（0〜1）
     * @param {number} green - 緑成分（0〜1）
     * @param {number} blue - 青成分（0〜1）
     * @returns {number} もっとも近いプリセットの index
     */
    function findClosestBorderColorIndex(red, green, blue) {
        var closestIndex = 0;
        var closestDistance = Infinity;
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            var colorPreset = BORDER_COLOR_PRESETS[i];
            var distance = Math.abs(colorPreset.r - red) + Math.abs(colorPreset.g - green) + Math.abs(colorPreset.b - blue);
            if (distance < closestDistance) {
                closestDistance = distance;
                closestIndex = i;
            }
        }
        return closestIndex;
    }

    /**
     * colorKey からカラープリセットの index を取得する
     * @param {string} colorKey - BORDER_COLOR_PRESETS の colorKey（例: "lightRed"）
     * @returns {number} プリセットの index（見つからない場合はブラック）
     */
    function findBorderColorIndexByKey(colorKey) {
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            if (BORDER_COLOR_PRESETS[i].colorKey === colorKey) return i;
        }
        return BORDER_COLOR_BLACK_INDEX;
    }

    // =========================================
    // アートボード情報（BridgeTalk 委譲）/ Artboard info (BridgeTalk delegation)
    // 常駐パレットのイベントハンドラ内では DOM 接続を失うため、DOM の読み書きはメインエンジンへ委譲する
    // Inside persistent-palette event handlers the DOM connection is lost, so all DOM access is delegated to the main engine
    // =========================================

    /* フィールド区切り（アートボード名に現れにくい ASCII 文字列／エスケープ不要）/ Field separator (ASCII, unlikely in names, no escaping) */
    var ARTBOARD_FIELD_SEPARATOR = "<|>";

    /**
     * メインエンジンへコードを委譲し、結果文字列をコールバックへ渡す
     * @param {string} bodyCode - メインエンジンで実行するコード
     * @param {function} [onResult] - 結果文字列を受け取るコールバック（失敗時は空文字）
     * @returns {void}
     */
    function delegateToMainEngine(bodyCode, onResult) {
        if (typeof onResult !== "function") onResult = function () {};
        var bridge = new BridgeTalk();
        bridge.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
        bridge.body = bodyCode;
        bridge.onResult = function (response) { onResult(String(response.body)); };
        bridge.onError = function () { onResult(""); };
        bridge.send();
    }

    /**
     * メインエンジンで実行するコード本体を組み立てる（read／丸め／リサイズ後に最新情報を返す）
     * @param {string} operation - "read" | "round" | "resize"
     * @param {number} [widthPoint] - resize 時の幅（pt）
     * @param {number} [heightPoint] - resize 時の高さ（pt）
     * @param {number} [anchorIndex] - resize 時の基準点（0=左上〜8=右下の行優先）
     * @returns {string} メインエンジンへ送るコード
     */
    function buildArtboardBridgeCode(operation, widthPoint, heightPoint, anchorIndex) {
        var mutationCode = "";
        if (operation === "round") {
            mutationCode = "artboard.artboardRect=[Math.round(rect[0]),Math.round(rect[1]),Math.round(rect[2]),Math.round(rect[3])];rect=artboard.artboardRect;";
        } else if (operation === "resize") {
            /* 負数連結による '--' 構文エラーを避けるため括弧で囲む / Wrap in parens to avoid '--' from negative numbers */
            /* 基準点の側に寄せる：増減分に 0／0.5／1 を掛けて左上をずらす / Shift the top-left by 0, 0.5 or 1 of the size change toward the anchor */
            var anchorColumnRatio = (anchorIndex % 3) / 2;
            var anchorRowRatio = Math.floor(anchorIndex / 3) / 2;
            mutationCode = "var w=" + Number(widthPoint) + ",h=" + Number(heightPoint) + ";" +
                "var left=rect[0]+(rect[2]-rect[0]-w)*" + anchorColumnRatio + ";" +
                "var top=rect[1]-(rect[1]-rect[3]-h)*" + anchorRowRatio + ";" +
                "artboard.artboardRect=[left,top,left+w,top-h];rect=artboard.artboardRect;";
        }
        var separatorCode = "+\"" + ARTBOARD_FIELD_SEPARATOR + "\"+";
        return "" +
            "(function(){try{" +
            "var doc=app.activeDocument;" +
            "var artboardIndex=doc.artboards.getActiveArtboardIndex();" +
            "var artboard=doc.artboards[artboardIndex];var rect=artboard.artboardRect;" +
            mutationCode +
            "return (artboardIndex+1)" + separatorCode + "artboard.name" + separatorCode + "(rect[2]-rect[0])" + separatorCode + "(rect[1]-rect[3]);" +
            "}catch(e){return \"\";}})();";
    }

    // =========================================
    // UI構築 / UI construction
    // =========================================

    /**
     * ラベル・入力欄・単位表示を並べたサイズ入力欄を作る
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} labelPath - ラベルの LABELS キー
     * @param {string} unitLabel - 単位ラベル（例: "mm"）
     * @returns {object} { input: EditText, unitText: StaticText }
     */
    function addSizeField(parentRow, labelPath, unitLabel) {
        var sizeFieldGroup = parentRow.add("group");
        setupRow(sizeFieldGroup, "left", 4);
        var sizeLabel = sizeFieldGroup.add("statictext", undefined, labelText(labelPath));
        sizeLabel.preferredSize.width = SIZE_LABEL_WIDTH;
        sizeLabel.justify = "right";
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = sizeFieldGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var sizeInput;
        var sizeStepper = addStepper(stepperInputGroup, function () { return sizeInput; }, {
            min: 0,
            /* ∧∨のクリックはその場でリサイズ。↑↓キーは離したときに確定（commitOnArrowKeyUp）
               A click resizes at once; the arrow keys commit on key-up (commitOnArrowKeyUp) */
            onStep: function (numberInput) {
                if (!numberInput.isArrowKeyDown && typeof numberInput.onChange === "function") numberInput.onChange();
            }
        });
        sizeInput = stepperInputGroup.add("edittext", undefined, "");
        sizeInput.helpTip = getLabel("tooltip.sizeField");
        sizeInput.characters = 5;
        commitOnArrowKeyUp(sizeInput);
        bindSteppedArrowKeys(sizeInput, sizeStepper);
        var unitText = sizeFieldGroup.add("statictext", undefined, unitLabel);
        return { input: sizeInput, unitText: unitText };
    }

    /**
     * ヘルプ付きのボタンを追加する
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} labelKey - LABELS.button / LABELS.tooltip 共通のキー
     * @param {Array|string} alignment - ボタンの整列
     * @returns {Button} 追加したボタン
     */
    function addButtonWithTip(parentRow, labelKey, alignment) {
        var button = parentRow.add("button", undefined, getLabel("button." + labelKey));
        button.helpTip = getLabel("tooltip." + labelKey);
        button.alignment = alignment;
        return button;
    }

    /**
     * 「現在のアートボード」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildCurrentArtboardPanel(parentWindow) {
        var currentArtboardPanel = parentWindow.add("panel", undefined, getLabel("panel.currentArtboard"));
        setupPanel(currentArtboardPanel, 8);

        /* 番号・名前（左右中央）/ Number and name (centered) */
        var artboardInfoRow = currentArtboardPanel.add("group");
        setupRow(artboardInfoRow, "fill");
        artboardInfoRow.alignChildren = ["center", "center"];
        var artboardInfoText = artboardInfoRow.add("statictext", undefined, "—");
        artboardInfoText.characters = 28;
        artboardInfoText.justify = "center";

        /* 幅・高さ（縦並び）＋基準点の9軸 / Width and height (stacked) + 9-axis anchor */
        var sizeRow = currentArtboardPanel.add("group");
        setupRow(sizeRow, "left", 16);
        var sizeColumn = sizeRow.add("group");
        sizeColumn.orientation = "column";
        sizeColumn.alignChildren = ["left", "center"];
        sizeColumn.alignment = "left";
        sizeColumn.spacing = 6;
        var rulerUnitLabel = getUnitInfo().label;
        var widthField = addSizeField(sizeColumn, "fieldLabel.width", rulerUnitLabel);
        var heightField = addSizeField(sizeColumn, "fieldLabel.height", rulerUnitLabel);
        var anchorWidget = addAnchorWidget(sizeRow, DEFAULT_ANCHOR_INDEX);
        anchorWidget.helpTip = getLabel("tooltip.anchor");

        /* ボタン行（パネル幅いっぱいには広げない）/ Button row (do not stretch to the panel width) */
        var artboardButtonRow = currentArtboardPanel.add("group");
        setupRow(artboardButtonRow, "left");

        return {
            artboardInfoText: artboardInfoText,
            widthInput: widthField.input,
            heightInput: heightField.input,
            widthUnitText: widthField.unitText,
            heightUnitText: heightField.unitText,
            anchorWidget: anchorWidget,
            optimizePixelGridButton: addButtonWithTip(artboardButtonRow, "optimizePixelGrid", "left")
        };
    }

    /**
     * 同じヘルプを付けたラジオボタンを並べる
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string[]} radioTexts - ラジオボタンの文言
     * @param {string} tooltipText - 共通のヘルプ
     * @returns {RadioButton[]} 追加したラジオボタン
     */
    function addRadioRow(parentRow, radioTexts, tooltipText) {
        var radios = [];
        for (var i = 0; i < radioTexts.length; i++) {
            var radio = parentRow.add("radiobutton", undefined, radioTexts[i]);
            radio.helpTip = tooltipText;
            radios.push(radio);
        }
        return radios;
    }

    /**
     * 「アートボード名と枠線」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildArtboardDisplayPanel(parentWindow) {
        var artboardDisplayPanel = parentWindow.add("panel", undefined, getLabel("panel.artboardDisplay"));
        setupPanel(artboardDisplayPanel);

        var showNameCheckbox = artboardDisplayPanel.add("checkbox", undefined, getLabel("checkbox.showArtboardName"));
        showNameCheckbox.helpTip = getLabel("tooltip.showArtboardName");

        /* 枠線サブパネル / Border sub-panel */
        var borderPanel = artboardDisplayPanel.add("panel", undefined, getLabel("panel.artboardBorder"));
        setupPanel(borderPanel, 8);

        var borderColorRow = borderPanel.add("group");
        setupRow(borderColorRow, "left");
        borderColorRow.add("statictext", undefined, labelText("fieldLabel.borderColor"));
        var borderColorList = borderColorRow.add("dropdownlist", undefined, buildBorderColorNames());
        borderColorList.helpTip = getLabel("tooltip.borderColor");

        var borderWidthRow = borderPanel.add("group");
        setupRow(borderWidthRow, "left");
        borderWidthRow.add("statictext", undefined, labelText("fieldLabel.borderWidth"));
        var borderWidthTexts = [];
        for (var i = 0; i < BORDER_WIDTH_CHOICES.length; i++) {
            borderWidthTexts.push(String(BORDER_WIDTH_CHOICES[i]));
        }

        /* プリセット（パネル最下部・左右中央）/ Presets (bottom of panel, centered) */
        var presetRow = artboardDisplayPanel.add("group");
        setupRow(presetRow, "center");
        var presetTexts = [];
        for (var j = 0; j < PRESET_KEYS.length; j++) {
            presetTexts.push(getLabel("radio.preset." + PRESET_KEYS[j]));
        }

        return {
            showNameCheckbox: showNameCheckbox,
            borderColorList: borderColorList,
            borderWidthRadios: addRadioRow(borderWidthRow, borderWidthTexts, getLabel("tooltip.borderWidth")),
            presetRadios: addRadioRow(presetRow, presetTexts, getLabel("tooltip.preset"))
        };
    }

    /**
     * 「オプション」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildOptionsPanel(parentWindow) {
        var optionsPanel = parentWindow.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 8);

        /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        // var printBleedCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.showPrintBleedAI"));

        var moveLockedHiddenCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.moveLockedHidden"));
        moveLockedHiddenCheckbox.helpTip = getLabel("tooltip.moveLockedHidden");

        return {
            // printBleedCheckbox: printBleedCheckbox,
            moveLockedHiddenCheckbox: moveLockedHiddenCheckbox
        };
    }

    /**
     * 下部のボタン行（カンバスカラー／ビデオ定規）を構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} 行内のボタン
     */
    function buildFooterRow(parentWindow) {
        var footerRow = parentWindow.add("group");
        setupRow(footerRow, "fill");

        var canvasColorButton = addButtonWithTip(footerRow, "canvasColor", ["left", "center"]);
        var footerSpacer = footerRow.add("group");
        footerSpacer.alignment = ["fill", "fill"];
        var videoRulerButton = addButtonWithTip(footerRow, "videoRuler", ["right", "center"]);

        return { canvasColorButton: canvasColorButton, videoRulerButton: videoRulerButton };
    }

    /**
     * パレット全体を構築する
     * @returns {object} ウィンドウと全コントロールをまとめたオブジェクト
     */
    function buildPalette() {
        var paletteWindow = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(paletteWindow);

        var paletteUI = { paletteWindow: paletteWindow };
        var sectionControls = [
            buildCurrentArtboardPanel(paletteWindow),
            buildArtboardDisplayPanel(paletteWindow),
            buildOptionsPanel(paletteWindow),
            buildFooterRow(paletteWindow)
        ];
        for (var i = 0; i < sectionControls.length; i++) {
            for (var controlName in sectionControls[i]) {
                if (sectionControls[i].hasOwnProperty(controlName)) paletteUI[controlName] = sectionControls[i][controlName];
            }
        }
        return paletteUI;
    }

    // =========================================
    // 即時反映：UI → 環境設定 / Immediate apply: UI to preferences
    // =========================================

    /**
     * 選択中の枠線の太さ（1〜4）を取得する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {number} 選択中の太さ
     */
    function getSelectedBorderWidth(paletteUI) {
        for (var i = 0; i < paletteUI.borderWidthRadios.length; i++) {
            if (paletteUI.borderWidthRadios[i].value) return BORDER_WIDTH_CHOICES[i];
        }
        return BORDER_WIDTH_CHOICES[0];
    }

    /**
     * アートボード名・枠線カラー・枠線の太さを環境設定へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function applyArtboardDisplaySettings(paletteUI) {
        appPreferences.setBooleanPreference("showArtboardLabelOnCanvas", paletteUI.showNameCheckbox.value);
        var colorIndex = paletteUI.borderColorList.selection ? paletteUI.borderColorList.selection.index : BORDER_COLOR_BLACK_INDEX;
        var borderColor = BORDER_COLOR_PRESETS[colorIndex];
        appPreferences.setRealPreference("ArtboardBBColorRed", borderColor.r);
        appPreferences.setRealPreference("ArtboardBBColorGreen", borderColor.g);
        appPreferences.setRealPreference("ArtboardBBColorBlue", borderColor.b);
        appPreferences.setRealPreference("ArtboardBBWidth", getSelectedBorderWidth(paletteUI));
        refreshArtboardDisplay();
    }

    // PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note
    // 「裁ち落としを印刷」生成AIボタンの表示設定を、書き込み・再描画・読み戻しまで
    // メインエンジンへ委譲する（書き込めなかった場合はチェックを実際の値へ戻す）
    // function applyPrintBleedWidgetSetting(paletteUI) {
    //     var enabled = paletteUI.printBleedCheckbox.value;
    //     delegateToMainEngine("" +
    //         "(function(){" +
    //         "app.preferences.setBooleanPreference('" + PRINT_BLEED_WIDGET_KEY + "'," + (enabled ? "true" : "false") + ");" +
    //         "try{if(app.documents.length){app.activeDocument.activate();app.redraw();" +
    //         "app.executeMenuCommand('zoomout');app.executeMenuCommand('zoomin');}}catch(e){}" +
    //         "return String(app.preferences.getBooleanPreference('" + PRINT_BLEED_WIDGET_KEY + "'));" +
    //         "})();",
    //         function (result) {
    //             if (result === "true" || result === "false") {
    //                 paletteUI.printBleedCheckbox.value = (result === "true");
    //             }
    //         });
    // }

    /**
     * オプション（ロックまたは非表示オブジェクトを一緒に移動）を環境設定へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function applyOptionSettings(paletteUI) {
        // applyPrintBleedWidgetSetting(paletteUI); /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        appPreferences.setBooleanPreference("moveLockedAndHiddenArt", paletteUI.moveLockedHiddenCheckbox.value);
    }

    // =========================================
    // 値反映：環境設定 → UI / Reflect preferences into the UI
    // =========================================

    /**
     * 枠線の太さのラジオボタンを選択する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {number} borderWidth - 選択する太さ（1〜4）
     * @returns {void}
     */
    function selectBorderWidthRadio(paletteUI, borderWidth) {
        for (var i = 0; i < paletteUI.borderWidthRadios.length; i++) {
            paletteUI.borderWidthRadios[i].value = (BORDER_WIDTH_CHOICES[i] === borderWidth);
        }
    }

    /**
     * プリセットの内容を UI へ適用する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {string} presetKey - PRESETS のキー
     * @returns {void}
     */
    function applyPresetToUI(paletteUI, presetKey) {
        var preset = PRESETS[presetKey];
        if (!preset) return;
        paletteUI.showNameCheckbox.value = preset.showName;
        paletteUI.borderColorList.selection = findBorderColorIndexByKey(preset.colorKey);
        selectBorderWidthRadio(paletteUI, preset.borderWidth);
        // paletteUI.printBleedCheckbox.value = preset.printBleed; /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        paletteUI.moveLockedHiddenCheckbox.value = preset.moveLocked;
    }

    /**
     * 現在の UI 状態に一致するプリセットキーを返す
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {number} colorIndex - 現在のカラープリセット index
     * @param {number} borderWidth - 現在の枠線の太さ
     * @returns {string|null} 一致するプリセットキー（なければ null）
     */
    function findMatchingPresetKey(paletteUI, colorIndex, borderWidth) {
        for (var i = 0; i < PRESET_KEYS.length; i++) {
            var preset = PRESETS[PRESET_KEYS[i]];
            if (paletteUI.showNameCheckbox.value === preset.showName &&
                colorIndex === findBorderColorIndexByKey(preset.colorKey) &&
                borderWidth === preset.borderWidth &&
                /* paletteUI.printBleedCheckbox.value === preset.printBleed && */ /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
                paletteUI.moveLockedHiddenCheckbox.value === preset.moveLocked) {
                return PRESET_KEYS[i];
            }
        }
        return null;
    }

    /**
     * 現在の環境設定を UI へ反映し、一致するプリセットのラジオボタンを選択する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function reflectPreferences(paletteUI) {
        paletteUI.showNameCheckbox.value = !!readPref("Boolean", "showArtboardLabelOnCanvas", false);

        var colorIndex = findClosestBorderColorIndex(
            readPref("Real", "ArtboardBBColorRed", 0.0),
            readPref("Real", "ArtboardBBColorGreen", 0.0),
            readPref("Real", "ArtboardBBColorBlue", 0.0)
        );
        paletteUI.borderColorList.selection = colorIndex;

        /* 1〜4 の範囲に収める / Clamp to 1-4 */
        var borderWidth = Math.max(1, Math.min(BORDER_WIDTH_CHOICES.length, Math.round(readPref("Real", "ArtboardBBWidth", 1.0))));
        selectBorderWidthRadio(paletteUI, borderWidth);

        // paletteUI.printBleedCheckbox.value = !!readPref("Boolean", PRINT_BLEED_WIDGET_KEY, false); /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        paletteUI.moveLockedHiddenCheckbox.value = !!readPref("Boolean", "moveLockedAndHiddenArt", false);

        var matchedPresetKey = findMatchingPresetKey(paletteUI, colorIndex, borderWidth);
        for (var i = 0; i < PRESET_KEYS.length; i++) {
            paletteUI.presetRadios[i].value = (PRESET_KEYS[i] === matchedPresetKey);
        }
    }

    // =========================================
    // 現在のアートボード / Current artboard
    // =========================================

    /**
     * 委譲結果（番号<|>名前<|>幅pt<|>高さpt）を UI へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {string} result - メインエンジンからの結果文字列
     * @param {boolean} alertOnEmpty - 空結果（ドキュメントなし）でアラートを出すか
     * @returns {void}
     */
    function applyArtboardResult(paletteUI, result, alertOnEmpty) {
        var rulerUnit = getUnitInfo();
        paletteUI.widthUnitText.text = rulerUnit.label;
        paletteUI.heightUnitText.text = rulerUnit.label;

        var resultFields = result ? result.split(ARTBOARD_FIELD_SEPARATOR) : [];
        if (resultFields.length < 4) {
            paletteUI.artboardInfoText.text = "—";
            paletteUI.widthInput.text = "";
            paletteUI.heightInput.text = "";
            if (alertOnEmpty) alert(getLabel("alert.noDocument"));
            return;
        }
        var infoSeparator = (uiLang === "ja") ? "：" : ": ";
        paletteUI.artboardInfoText.text = "#" + resultFields[0] + infoSeparator + resultFields[1];
        paletteUI.widthInput.text = pointToUnitText(parseFloat(resultFields[2]), rulerUnit);
        paletteUI.heightInput.text = pointToUnitText(parseFloat(resultFields[3]), rulerUnit);
    }

    /**
     * アートボード操作をメインエンジンへ委譲し、結果を UI へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @param {string} operation - "read" | "round" | "resize"
     * @param {number} widthPoint - resize 時の幅（pt）
     * @param {number} heightPoint - resize 時の高さ（pt）
     * @param {boolean} alertOnEmpty - 空結果でアラートを出すか
     * @returns {void}
     */
    function runArtboardOperation(paletteUI, operation, widthPoint, heightPoint, alertOnEmpty) {
        var anchorIndex = getAnchorWidgetIndex(paletteUI.anchorWidget);
        delegateToMainEngine(buildArtboardBridgeCode(operation, widthPoint, heightPoint, anchorIndex), function (result) {
            applyArtboardResult(paletteUI, result, alertOnEmpty);
        });
    }

    /**
     * 現在のアートボード情報を取得して UI へ反映する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function refreshArtboardInfo(paletteUI) {
        runArtboardOperation(paletteUI, "read", 0, 0, false);
    }

    /**
     * 幅・高さの入力値でアクティブアートボードをリサイズする（不正値は現在値へ戻す）
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function resizeArtboardFromFields(paletteUI) {
        var rulerUnit = getUnitInfo();
        var widthValue = parseFloat(paletteUI.widthInput.text);
        var heightValue = parseFloat(paletteUI.heightInput.text);
        if (isNaN(widthValue) || isNaN(heightValue) || widthValue <= 0 || heightValue <= 0) {
            refreshArtboardInfo(paletteUI);
            return;
        }
        runArtboardOperation(paletteUI, "resize", widthValue * rulerUnit.pointsPerUnit, heightValue * rulerUnit.pointsPerUnit, true);
    }

    // =========================================
    // イベント設定 / Event wiring
    // =========================================

    /**
     * 表示設定（プリセット・アートボード名・枠線・オプション）のイベントを設定する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function wireDisplaySettingEvents(paletteUI) {
        var i;
        var applyDisplaySettings = function () { applyArtboardDisplaySettings(paletteUI); };

        /* プリセット：UI へ展開してから環境設定へ反映 / Presets: expand into the UI, then write */
        for (i = 0; i < PRESET_KEYS.length; i++) {
            paletteUI.presetRadios[i].presetKey = PRESET_KEYS[i];
            paletteUI.presetRadios[i].onClick = function () {
                applyPresetToUI(paletteUI, this.presetKey);
                applyArtboardDisplaySettings(paletteUI);
                applyOptionSettings(paletteUI);
            };
        }

        paletteUI.showNameCheckbox.onClick = applyDisplaySettings;
        paletteUI.borderColorList.onChange = applyDisplaySettings;
        for (i = 0; i < paletteUI.borderWidthRadios.length; i++) {
            paletteUI.borderWidthRadios[i].onClick = applyDisplaySettings;
        }

        // paletteUI.printBleedCheckbox.onClick = function () { applyOptionSettings(paletteUI); }; /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        paletteUI.moveLockedHiddenCheckbox.onClick = function () { applyOptionSettings(paletteUI); };
    }

    /**
     * アートボード操作・下部ボタン・ウィンドウのイベントを設定する
     * @param {object} paletteUI - パレットのコントロール一式
     * @returns {void}
     */
    function wireArtboardAndWindowEvents(paletteUI) {
        /* ピクセルグリッドに最適化：XYWH を整数値へ丸める / Optimize: round XYWH to integers */
        paletteUI.optimizePixelGridButton.onClick = function () { runArtboardOperation(paletteUI, "round", 0, 0, true); };

        /* 幅・高さの確定でアートボードをリサイズ / Resize the artboard when width/height are committed */
        paletteUI.widthInput.onChange = function () { resizeArtboardFromFields(paletteUI); };
        paletteUI.heightInput.onChange = function () { resizeArtboardFromFields(paletteUI); };

        /* カンバスカラーの変更（uiCanvasIsWhite: 1=白 / 0=グレー）/ Toggle the canvas color */
        paletteUI.canvasColorButton.onClick = function () {
            appPreferences.setIntegerPreference("uiCanvasIsWhite", readPref("Integer", "uiCanvasIsWhite", 0) === 1 ? 0 : 1);
            refreshArtboardDisplay();
        };
        paletteUI.videoRulerButton.onClick = function () { runMenuCommand("videoruler"); };

        /* アクティブ時に Esc で閉じる / Close on Esc while active */
        paletteUI.paletteWindow.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") paletteUI.paletteWindow.close();
        });

        /* 再アクティブ時：外部変更とアートボードの切り替えに追従 / On re-activate: follow external changes */
        paletteUI.paletteWindow.onActivate = function () {
            reflectPreferences(paletteUI);
            refreshArtboardInfo(paletteUI);
        };

        /* 常駐エンジンの参照は閉じたらクリア / Clear the persistent-engine reference on close */
        paletteUI.paletteWindow.onClose = function () {
            $.global[PALETTE_GLOBAL_KEY] = null;
        };
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

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "アートボード関連の環境設定", en: "Artboard-Related Preferences" }
        },
        panel: {
            currentArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            artboardDisplay: { ja: "アートボード名と枠線", en: "Artboard Name & Border" },
            artboardBorder: { ja: "アートボードの枠線", en: "Artboard Border" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            borderColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            borderWidth: { ja: "ストロークの幅", en: "Stroke Width" }
        },
        checkbox: {
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            /* showPrintBleedAI は現在未使用（PRINT_BLEED_WIDGET を参照）/ showPrintBleedAI is currently unused (see the PRINT_BLEED_WIDGET note) */
            showPrintBleedAI: {
                ja: "「裁ち落としを印刷」生成AIボタンを表示",
                en: "Show the \"Print Bleed\" Generative AI Button"
            },
            moveLockedHidden: {
                ja: "ロックまたは非表示オブジェクトを一緒に移動",
                en: "Move Locked or Hidden Objects Together"
            }
        },
        radio: {
            preset: {
                /* "default" は ES3 予約語のため引用符付きキーにする / "default" is an ES3 reserved word, so quote the key */
                "default": { ja: "デフォルト", en: "Default" },
                emphasis: { ja: "強調", en: "Emphasis" },
                light: { ja: "ライト", en: "Light" }
            }
        },
        dropdown: {
            borderColor: {
                lightBlue: { ja: "ライトブルー", en: "Light Blue" },
                lightRed: { ja: "サーモンピンク", en: "Light Red" },
                green: { ja: "グリーン", en: "Green" },
                mediumBlue: { ja: "ミディアムブルー", en: "Medium Blue" },
                magenta: { ja: "マゼンタ", en: "Magenta" },
                cyan: { ja: "シアン", en: "Cyan" },
                lightGray: { ja: "ライトグレー", en: "Light Gray" },
                black: { ja: "ブラック", en: "Black" },
                yellow: { ja: "イエロー", en: "Yellow" }
            }
        },
        button: {
            optimizePixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to Pixel Grid" },
            canvasColor: { ja: "カンバスカラーの変更", en: "Change Canvas Color" },
            videoRuler: { ja: "ビデオ定規", en: "Video Ruler" }
        },
        tooltip: {
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            sizeField: {
                ja: "値を確定すると、アクティブなアートボードを右の基準点を基準にリサイズします。↑↓で±1、shift併用で±10、option併用で±0.1。",
                en: "Commit a value to resize the active artboard around the reference point on the right. Arrow keys: ±1, Shift ±10, Option ±0.1."
            },
            anchor: { ja: "リサイズの基準点です。", en: "Reference point for resizing." },
            showArtboardName: { ja: "カンバス上にアートボード名を表示します。", en: "Shows the artboard names on the canvas." },
            borderColor: { ja: "アートボードの境界線の色です。", en: "Color of the artboard borders." },
            borderWidth: { ja: "アートボードの境界線の太さです。", en: "Width of the artboard borders." },
            preset: { ja: "まとめて切り替える表示設定の組み合わせです。", en: "A set of display settings applied together." },
            moveLockedHidden: {
                ja: "ロックや非表示のオブジェクトも、アートボードと一緒に動かします。",
                en: "Moves locked and hidden objects along with the artboard."
            },
            optimizePixelGrid: {
                ja: "アクティブなアートボードの位置とサイズを整数値に丸めます。",
                en: "Rounds the active artboard's position and size to whole numbers."
            },
            canvasColor: {
                ja: "アートボード外のカンバスを、白とグレーで切り替えます。",
                en: "Toggles the canvas outside the artboards between white and gray."
            },
            videoRuler: { ja: "ビデオ定規の表示／非表示を切り替えます。", en: "Shows or hides the video ruler." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * すでに開いているパレットを返す（無効な参照・表示前の残骸は null）
     * 表示中のものだけを有効と見なす。構築途中でエラーになったウィンドウを掴むと、
     * イベント未配線のパレットを開き続けることになるため
     * 閉じて破棄されたウィンドウは visible の参照で例外になるため try で受ける
     * @returns {Window|null} 既存のパレット
     */
    function getExistingPalette() {
        try {
            var palette = $.global[PALETTE_GLOBAL_KEY];
            return (palette && palette.visible === true) ? palette : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * パレットを構築して表示する（重複起動時は既存のパレットを前面に出す）
     * @returns {void}
     */
    function main() {
        var existingPalette = getExistingPalette();
        if (existingPalette) {
            existingPalette.show();
            return;
        }

        var paletteUI = buildPalette();
        wireDisplaySettingEvents(paletteUI);
        wireArtboardAndWindowEvents(paletteUI);
        reflectPreferences(paletteUI);
        refreshArtboardInfo(paletteUI);

        paletteUI.paletteWindow.center();
        paletteUI.paletteWindow.show();

        /* 構築と表示がすべて通ってから参照を保持する（途中で失敗した窓を残さない）
           Store the reference only after everything succeeded, so a half-built window is never kept */
        $.global[PALETTE_GLOBAL_KEY] = paletteUI.paletteWindow;

        /* レイアウト確定後にボタン高さを 2px 詰める / Trim the button heights by 2px after layout */
        trimButtonHeight(paletteUI.optimizePixelGridButton, 2);
    }

    main();

}());
