#target illustrator
#targetengine "BlendSpEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択内容に応じて、ブレンドの作成・設定・調整を1つのダイアログで行います。
ステップ数と方向はライブプレビューで確認でき、解除／拡張／ブレンド軸の置き換えは［OK］で確定したときだけ実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendSp.md

### Overview

Creates, configures and adjusts a blend from a single dialog, depending on what is selected.
Steps and orientation are shown as a live preview, while Release, Expand and Replace Spine run only when the dialog is confirmed.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BlendSp.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "BlendSp";                      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendSp.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BlendSp.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php
(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ブレンド以外を選択して開いたときのステップ数 / Steps used when no blend is selected at launch */
    var DEFAULT_BLEND_STEPS = 8;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;              /* ダイアログの余白 / dialog margins */
    var DIALOG_SPACING = 10;              /* ダイアログ内の要素間隔 / dialog spacing */
    var STEP_AREA_SPACING = 6;            /* ステップ数の行とスライダーの間隔 / gap between the Steps row and the slider */
    var COLUMN_SPACING = 12;              /* 2カラムの間隔 / gap between the two columns */
    var COLUMN_STACK_SPACING = 10;        /* カラム内のパネル間隔 / gap between panels in a column */
    var PANEL_MARGINS = [12, 18, 12, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var OPTION_LIST_SPACING = 4;          /* ラジオ・チェックボックスの行間 / gap between radios and checkboxes */

    /**
     * オプションパネルの共通設定
     * @param {Panel} optionPanel - 対象パネル
     * @param {string} horizontalAlign - 子の横方向の揃え（"fill" / "left"）
     * @param {number} spacing - 要素間隔
     * @returns {void}
     */
    function setupOptionPanel(optionPanel, horizontalAlign, spacing) {
        optionPanel.orientation = 'column';
        optionPanel.alignChildren = [horizontalAlign, 'top'];
        optionPanel.spacing = spacing;
        optionPanel.margins = PANEL_MARGINS;
    }

    /**
     * パネルを縦に積むカラムを追加する
     * @param {Group} columnsGroup - 追加先の横並びグループ
     * @returns {Group} 追加したカラム
     */
    function addColumnGroup(columnsGroup) {
        var columnGroup = columnsGroup.add('group');
        columnGroup.orientation = 'column';
        columnGroup.alignChildren = ['fill', 'top'];
        columnGroup.spacing = COLUMN_STACK_SPACING;
        return columnGroup;
    }

    /**
     * ラジオ・チェックボックスを縦に並べるグループを追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @returns {Group} 追加したグループ
     */
    function addOptionList(parentContainer) {
        var optionList = parentContainer.add('group');
        optionList.orientation = 'column';
        optionList.alignChildren = ['left', 'top'];
        optionList.spacing = OPTION_LIST_SPACING;
        return optionList;
    }

    /**
     * ラベルと tooltip 付きのコントロールを追加する
     * @param {Group} optionList - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox"
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - tooltip のパス
     * @returns {RadioButton|Checkbox} 追加したコントロール
     */
    function addOptionControl(optionList, controlType, labelPath, tooltipPath) {
        var optionControl = optionList.add(controlType, undefined, getLabel(labelPath));
        optionControl.helpTip = getLabel(tooltipPath);
        return optionControl;
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

    // =========================================
    // ブレンドの設定 / Blend settings
    // =========================================

    /* ブレンドオプション > 方向（0: Align to Page〈垂直方向〉/ 1: Align to Path〈パスに沿う〉）
       Blend Options > Orientation */
    var BlendOrientation = {
        'Align to Page': 0,
        'Align to Path': 1
    };

    /* ステップ数の上限 / Maximum number of steps */
    var MAX_BLEND_STEPS = 1000;

    /* ステップ数スライダーの上限（通常／Option 併用／Shift 併用）/ Slider maximum: plain / with Option / with Shift */
    var STEP_SLIDER_MAX = { normal: 32, option: 128, shift: 1000 };

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
            title: { ja: "ブレンドSpecial", en: "Blend Special" }
        },
        panel: {
            orientation: { ja: "方向", en: "Orientation" },
            adjust: { ja: "その他", en: "Misc" },
            reverse: { ja: "反転", en: "Reverse" }
        },
        fieldLabel: {
            steps: { ja: "ステップ数", en: "Steps" }
        },
        radio: {
            alignToPage: { ja: "垂直方向", en: "Align to Page" },
            alignToPath: { ja: "パスに沿う", en: "Align to Path" },
            adjustNone: { ja: "なし", en: "None" },
            adjustRelease: { ja: "解除", en: "Release" },
            adjustExpand: { ja: "拡張", en: "Expand" },
            adjustReplace: { ja: "ブレンド軸を置き換え", en: "Replace Spine" }
        },
        checkbox: {
            reverseSpine: { ja: "ブレンド軸を反転", en: "Reverse Spine" },
            reverseStack: { ja: "前後を反転", en: "Reverse Front to Back" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            step: {
                ja: "ブレンドの中間オブジェクト数です。0 で中間なし。スライダーでも変えられます。",
                en: "How many intermediate objects the blend creates. 0 means none. The slider changes it too."
            },
            alignToPage: {
                ja: "中間オブジェクトの向きを、ページの垂直方向に固定します。",
                en: "Keeps the intermediate objects upright with respect to the page."
            },
            alignToPath: {
                ja: "中間オブジェクトの向きを、スパイン（軸のパス）の傾きに合わせます。",
                en: "Rotates the intermediate objects to follow the spine."
            },
            adjustNone: { ja: "ブレンドをそのまま残します。", en: "Leaves the blend as a live blend." },
            adjustRelease: {
                ja: "ブレンドを解除して、元のオブジェクトとスパインに戻します。",
                en: "Releases the blend back into the original objects and the spine."
            },
            adjustExpand: {
                ja: "ブレンドを分割・拡張して、中間オブジェクトを実体のあるパスにします。",
                en: "Expands the blend so the intermediate steps become real paths."
            },
            adjustReplace: {
                ja: "スパインを、選択しておいた別のパスに置き換えます。",
                en: "Replaces the spine with another path you selected."
            },
            reverseSpine: {
                ja: "スパインの向きを反転し、始点と終点を入れ替えます。",
                en: "Reverses the spine so the start and the end swap places."
            },
            reverseStack: { ja: "ブレンド内の重ね順を逆にします。", en: "Reverses the stacking order inside the blend." },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        alert: {
            invalidStep: { ja: "ステップ数は 0〜1000 の整数で入力してください。", en: "Please enter an integer from 0 to 1000." },
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." }
        }
    };

    // =========================================
    // 選択とブレンドの判定 / Selection and blend helpers
    // =========================================

    /**
     * 選択から最初の BlendItem または PluginItem を返す（BlendItem を優先）
     * @param {PageItem[]} selection - 探索する選択オブジェクト
     * @returns {PageItem|null} 見つかったブレンドオブジェクト。無ければ null
     */
    function findFirstBlendObject(selection) {
        try {
            if (!selection || selection.length <= 0) {
                return null;
            }
            /* BlendItem を優先（blendOptions.steps を持つことが多い）/ Prefer BlendItem; it usually exposes blendOptions.steps */
            for (var i = 0; i < selection.length; i++) {
                if (selection[i] && selection[i].typename === 'BlendItem') {
                    return selection[i];
                }
            }
            for (var j = 0; j < selection.length; j++) {
                if (selection[j] && selection[j].typename === 'PluginItem') {
                    return selection[j];
                }
            }
        } catch (e) { }
        return null;
    }

    /**
     * 選択がブレンドオブジェクトのみで構成されているかを判定する
     *
     * 「ブレンド軸を置き換え」の有効／無効制御に使う。ブレンドとパスの混在選択では false になる。
     * @param {PageItem[]} selection - 判定する選択オブジェクト
     * @returns {boolean} すべてブレンドなら true
     */
    function isOnlyBlendObjects(selection) {
        try {
            if (!selection || selection.length <= 0) {
                return false;
            }
            /* 1つ以上あり、全てが BlendItem / PluginItem のとき「ブレンドのみ」/ Every item is a blend */
            for (var i = 0; i < selection.length; i++) {
                var selectedItem = selection[i];
                if (!selectedItem) {
                    return false;
                }
                var itemType = selectedItem.typename;
                if (itemType !== 'PluginItem' && itemType !== 'BlendItem') {
                    return false;
                }
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 指定したアイテムだけを選択状態にする（失敗しても例外は投げない）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} targetItem - 選択したいアイテム
     * @returns {void}
     */
    function selectOnlyItem(doc, targetItem) {
        try {
            if (!doc || !targetItem) {
                return;
            }
            doc.selection = null;
            targetItem.selected = true;
        } catch (e) { }
    }

    /**
     * 現在の選択を配列に控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択していたアイテム
     */
    function snapshotSelection(doc) {
        var selectedItems = [];
        try {
            var currentSelection = doc.selection;
            if (currentSelection && currentSelection.length) {
                for (var i = 0; i < currentSelection.length; i++) {
                    selectedItems.push(currentSelection[i]);
                }
            }
        } catch (e) { }
        return selectedItems;
    }

    /**
     * snapshotSelection() で控えた選択へ戻す（選択できないアイテムは飛ばす）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 控えておいたアイテム
     * @returns {void}
     */
    function restoreSelection(doc, selectedItems) {
        try {
            doc.selection = null;
            for (var i = 0; i < selectedItems.length; i++) {
                /* ロック・非表示などで選択できないアイテムは飛ばす / Skip items that cannot be selected */
                try { selectedItems[i].selected = true; } catch (err) { }
            }
        } catch (e) { }
    }

    /**
     * 選択のうち最前面のパスを削除する
     * @param {PageItem[]} selection - 対象の選択オブジェクト
     * @returns {boolean} 削除できた場合は true、対象が無い場合は false
     */
    function deleteFrontmostPath(selection) {
        try {
            if (!selection || selection.length <= 0) {
                return false;
            }

            var frontmostPath = null;
            var frontmostZ = null;

            for (var i = 0; i < selection.length; i++) {
                var selectedItem = selection[i];
                if (!selectedItem || selectedItem.typename !== 'PathItem') {
                    continue;
                }

                var zPosition = null;
                try {
                    zPosition = selectedItem.zOrderPosition;
                } catch (err) {
                    zPosition = null;
                }

                /* zOrderPosition が大きいものを優先。取れなければ先に見つかったもの（比較できる値が無いうちは後のもので置き換え）
                   Prefer the higher zOrderPosition; while no comparable value exists, the later path takes over */
                var hasFrontmostZ = (frontmostZ !== null && frontmostZ !== undefined);
                var hasZ = (zPosition !== null && zPosition !== undefined);
                if (!hasFrontmostZ || (hasZ && zPosition > frontmostZ)) {
                    frontmostPath = selectedItem;
                    frontmostZ = zPosition;
                }
            }

            if (frontmostPath) {
                frontmostPath.remove();
                return true;
            }
        } catch (e) { }

        return false;
    }

    /**
     * ブレンドのメニューコマンドを実行し、選択に残った最前面のパスを削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} menuCommand - 実行するメニューコマンド
     * @returns {void}
     */
    function runBlendCommandAndDeleteFrontmostPath(doc, menuCommand) {
        app.executeMenuCommand(menuCommand);
        /* コマンド実行後に選択が変わる可能性があるため、doc.selection を参照する
           The command may change the selection, so read doc.selection afresh */
        try {
            deleteFrontmostPath(doc.selection);
        } catch (e) { }
    }

    /**
     * ステップ数を 0〜MAX_BLEND_STEPS に収める
     * @param {number} stepValue - 対象の値
     * @returns {number} 範囲に収めた値
     */
    function clampBlendSteps(stepValue) {
        if (stepValue < 0) return 0;
        if (stepValue > MAX_BLEND_STEPS) return MAX_BLEND_STEPS;
        return stepValue;
    }

    /**
     * PluginItem のブレンドから中間ステップ数を取得する
     *
     * 元データを壊さないよう、複製してから分割・拡張してカウントし、後片付けする。
     * @param {PageItem} blendItem - ステップ数を調べるブレンド（PluginItem）
     * @returns {number} 中間ステップ数。取得できない場合は -1
     */
    function getBlendStepsFromPluginItem(blendItem) {
        /* PluginItem 以外は対象外 / Only PluginItem is supported */
        if (!blendItem || blendItem.typename !== 'PluginItem') {
            return -1;
        }

        var doc = app.activeDocument;
        var stepCount = -1;
        var duplicatedBlend = null;

        /* 選択状態を退避（この関数内で選択を変更するため）/ Save the selection; this function changes it */
        var originalSelection = snapshotSelection(doc);

        try {
            /* 1) 複製（元データ保護）/ Duplicate to protect the original */
            duplicatedBlend = blendItem.duplicate();

            /* 2) 複製したものだけを選択 / Select only the duplicate */
            doc.selection = null;
            duplicatedBlend.selected = true;

            /* 3) 分割・拡張 / Expand */
            app.executeMenuCommand('Path Blend Expand');

            /* 4) 分割・拡張後の選択（展開結果）を確認 / Check the expanded result */
            if (doc.selection && doc.selection.length > 0) {
                /* 5) トップレベルのグループを解除（ネストまでは無理に追わない）/ Ungroup the top level only */
                try { app.executeMenuCommand('ungroup'); } catch (err) { }

                /* 6) 数を数え、7) 始点・終点を除外（全要素数 - 2）/ Count, excluding the start and end objects */
                var expandedCount = doc.selection.length;
                stepCount = (expandedCount >= 2) ? (expandedCount - 2) : 0;
            }
        } catch (e) {
            stepCount = -1;
        } finally {
            /* 展開残骸を削除 / Remove what the expansion left */
            try {
                var expandedItems = doc.selection;
                if (expandedItems && expandedItems.length) {
                    for (var i = 0; i < expandedItems.length; i++) {
                        try { expandedItems[i].remove(); } catch (removeError) { }
                    }
                }
            } catch (selectionError) { }

            /* 複製が残っていた場合の保険 / In case the duplicate survived */
            try {
                if (duplicatedBlend) { duplicatedBlend.remove(); }
            } catch (duplicateError) { }

            /* 選択状態を復元 / Restore the selection */
            restoreSelection(doc, originalSelection);
        }

        return stepCount;
    }

    /**
     * ダイアログのステップ数の初期値を決める
     *
     * ブレンドを選択して開いたときは、そのステップ数を読めれば使う（読めなければ既定値）。
     * @param {PageItem[]} selection - 実行時の選択オブジェクト
     * @param {boolean} wasBlendSelectedAtOpen - ダイアログを開いた時点でブレンドを選択していたか
     * @returns {number} ステップ数の初期値
     */
    function readDefaultStep(selection, wasBlendSelectedAtOpen) {
        var defaultStep = DEFAULT_BLEND_STEPS;
        if (!wasBlendSelectedAtOpen) {
            return defaultStep;
        }
        try {
            var blendObject = findFirstBlendObject(selection);
            /* BlendItem は blendOptions.steps を持つことが多い。PluginItem は持たないことがある
               BlendItem typically exposes blendOptions.steps; PluginItem may not */
            if (blendObject && blendObject.blendOptions && typeof blendObject.blendOptions.steps === 'number') {
                defaultStep = blendObject.blendOptions.steps;
            } else if (blendObject && blendObject.typename === 'PluginItem') {
                var countedSteps = getBlendStepsFromPluginItem(blendObject);
                if (typeof countedSteps === 'number' && countedSteps >= 0) {
                    defaultStep = countedSteps;
                }
            }
        } catch (e) { }
        return defaultStep;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ステップ数の入力欄とスライダーを組み立てる
     * @param {Window} blendDialog - 追加先のダイアログ
     * @param {number} defaultStep - ステップ数の初期値
     * @returns {{stepInput: EditText, stepSlider: Slider}} 入力欄とスライダー
     */
    function buildStepArea(blendDialog, defaultStep) {
        var stepArea = blendDialog.add('group');
        stepArea.orientation = 'column';
        stepArea.alignChildren = ['fill', 'top'];
        stepArea.spacing = STEP_AREA_SPACING;

        var stepRow = stepArea.add('group');
        stepRow.orientation = 'row';
        stepRow.alignChildren = ['left', 'center'];
        stepRow.add('statictext', undefined, getLabel('fieldLabel.steps'));

        /* ∧∨と入力欄は隙間0で突き合わせる。0〜上限の整数（onStep は bindStepControls() で入れる）
           Butt the stepper against the field; integers within 0..MAX_BLEND_STEPS */
        var stepperInputGroup = stepRow.add('group');
        stepperInputGroup.orientation = 'row';
        stepperInputGroup.alignChildren = ['left', 'center'];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var stepOptions = { step: 1, min: 0, max: MAX_BLEND_STEPS, integer: true };
        var stepInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return stepInput; }, stepOptions);
        stepInput = stepperInputGroup.add('edittext', undefined, String(defaultStep));
        stepInput.helpTip = getLabel('tooltip.step');
        stepInput.characters = 4;
        stepInput.stepOptions = stepOptions;
        bindSteppedArrowKeys(stepInput, stepperGroup);

        /* スライダー（ステップ数の下、横いっぱい）/ Slider under Steps, full width */
        var stepSlider = stepArea.add('slider', undefined, defaultStep, 0, STEP_SLIDER_MAX.normal);
        stepSlider.helpTip = getLabel('tooltip.step');
        stepSlider.alignment = ['fill', 'center'];

        return { stepInput: stepInput, stepSlider: stepSlider };
    }

    /**
     * 「方向」「反転」（左カラム）と「その他」（右カラム）のパネルを組み立てる
     * @param {Window} blendDialog - 追加先のダイアログ
     * @returns {Object} 各パネルとラジオ・チェックボックス
     */
    function buildOptionColumns(blendDialog) {
        var columnsGroup = blendDialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignChildren = ['fill', 'top'];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = addColumnGroup(columnsGroup);
        var rightColumnGroup = addColumnGroup(columnsGroup);

        /* 方向 / Orientation */
        var orientationPanel = leftColumnGroup.add('panel', undefined, getLabel('panel.orientation'));
        setupOptionPanel(orientationPanel, 'fill', 8);

        var orientationRow = orientationPanel.add('group');
        orientationRow.orientation = 'row';
        orientationRow.alignChildren = ['left', 'top'];

        var orientationList = addOptionList(orientationRow);
        var alignToPageRadio = addOptionControl(orientationList, 'radiobutton', 'radio.alignToPage', 'tooltip.alignToPage');
        var alignToPathRadio = addOptionControl(orientationList, 'radiobutton', 'radio.alignToPath', 'tooltip.alignToPath');
        alignToPageRadio.value = true;

        /* その他（解除・拡張・ブレンド軸の置き換え）/ Misc: release, expand, replace spine */
        var adjustPanel = rightColumnGroup.add('panel', undefined, getLabel('panel.adjust'));
        setupOptionPanel(adjustPanel, 'left', 6);

        var adjustList = addOptionList(adjustPanel);
        var adjustNoneRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustNone', 'tooltip.adjustNone');
        var adjustReleaseRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustRelease', 'tooltip.adjustRelease');
        var adjustExpandRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustExpand', 'tooltip.adjustExpand');
        var adjustReplaceRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustReplace', 'tooltip.adjustReplace');
        adjustNoneRadio.value = true;

        /* 反転 / Reverse */
        var reversePanel = leftColumnGroup.add('panel', undefined, getLabel('panel.reverse'));
        setupOptionPanel(reversePanel, 'left', 6);

        var reverseList = addOptionList(reversePanel);
        var reverseSpineCheckbox = addOptionControl(reverseList, 'checkbox', 'checkbox.reverseSpine', 'tooltip.reverseSpine');
        var reverseStackCheckbox = addOptionControl(reverseList, 'checkbox', 'checkbox.reverseStack', 'tooltip.reverseStack');
        reverseSpineCheckbox.value = false;
        reverseStackCheckbox.value = false;

        return {
            orientationPanel: orientationPanel,
            alignToPageRadio: alignToPageRadio,
            alignToPathRadio: alignToPathRadio,
            adjustNoneRadio: adjustNoneRadio,
            adjustReleaseRadio: adjustReleaseRadio,
            adjustExpandRadio: adjustExpandRadio,
            adjustReplaceRadio: adjustReplaceRadio,
            reversePanel: reversePanel,
            reverseSpineCheckbox: reverseSpineCheckbox,
            reverseStackCheckbox: reverseStackCheckbox
        };
    }

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

    /**
     * 選択中の方向をブレンドオプションの値で返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {number} BlendOrientation の値
     */
    function readOrientation(dialogControls) {
        return (dialogControls.alignToPathRadio.value) ? BlendOrientation['Align to Path'] : BlendOrientation['Align to Page'];
    }

    /**
     * 選択中の「その他」の調整モードを返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "none" / "release" / "expand" / "replaceSpine"
     */
    function readAdjustMode(dialogControls) {
        if (dialogControls.adjustReleaseRadio.value) return 'release';
        if (dialogControls.adjustExpandRadio.value) return 'expand';
        if (dialogControls.adjustReplaceRadio.value) return 'replaceSpine';
        return 'none';
    }

    /**
     * ステップ数の入力欄とスライダーを連動させ、値が変わるたびにプレビューを更新する
     * @param {EditText} stepInput - ステップ数の入力欄
     * @param {Slider} stepSlider - ステップ数のスライダー
     * @param {function} refreshPreview - プレビューを更新するコールバック
     * @returns {void}
     */
    function bindStepControls(stepInput, stepSlider, refreshPreview) {
        /* 範囲を切り替えるとき（32 ⇔ 128 ⇔ 1000）につまみの位置を保つための控え
           Last slider maximum, used to keep the thumb position when the range switches */
        var lastSliderMax = STEP_SLIDER_MAX.normal;

        /**
         * 修飾キーの状態に応じてステップ数スライダーの上限を切り替える
         *
         * 通常0〜32、Option併用で0〜128、Shift併用で0〜1000。つまみの相対位置は保つ。
         * @returns {void}
         */
        function updateStepSliderRangeFromKeyboard() {
            try {
                var keyboardState = ScriptUI.environment.keyboardState;

                /* 優先順位: Shift (0-1000) > Option (0-128) > 通常 (0-32) / Priority: Shift > Option > default */
                var nextMax = STEP_SLIDER_MAX.normal;
                if (keyboardState && keyboardState.shiftKey) {
                    nextMax = STEP_SLIDER_MAX.shift;
                } else if (keyboardState && keyboardState.altKey) {
                    nextMax = STEP_SLIDER_MAX.option;
                }

                if (lastSliderMax !== nextMax) {
                    /* 範囲を変えてもつまみの相対位置（比率）を保つ / Keep the thumb's relative position */
                    var thumbRatio = (lastSliderMax > 0) ? (stepSlider.value / lastSliderMax) : 0;
                    if (!isFinite(thumbRatio) || isNaN(thumbRatio)) thumbRatio = 0;

                    stepSlider.maxvalue = nextMax;

                    var nextValue = Math.round(thumbRatio * nextMax);
                    if (nextValue < 0) nextValue = 0;
                    if (nextValue > nextMax) nextValue = nextMax;

                    stepSlider.value = nextValue;
                    /* 入力欄も新しい値にそろえる / Keep the field in step with the new value */
                    stepInput.text = String(nextValue);

                    lastSliderMax = nextMax;
                } else if (stepSlider.maxvalue !== nextMax) {
                    /* 上限が他所で変わっていても範囲をそろえる / Re-apply the range if it drifted */
                    stepSlider.maxvalue = nextMax;
                }
            } catch (e) { }
        }

        /**
         * ステップ数を整数に丸め、0以上スライダー上限以下に収める
         * @param {number} stepValue - 丸める前の値
         * @returns {number} 0〜スライダー上限に収めた整数
         */
        function clampToSliderRange(stepValue) {
            var roundedValue = Math.round(stepValue);
            if (roundedValue < 0) roundedValue = 0;
            if (roundedValue > stepSlider.maxvalue) roundedValue = stepSlider.maxvalue;
            return roundedValue;
        }

        /**
         * 入力欄の値をスライダーへ反映する
         * @returns {void}
         */
        function syncSliderFromEditText() {
            var stepValue = parseInt(stepInput.text, 10);
            if (isNaN(stepValue)) return;
            stepSlider.value = clampToSliderRange(stepValue);
        }

        /**
         * スライダーの値を入力欄へ反映する
         * @returns {void}
         */
        function syncEditTextFromSlider() {
            stepInput.text = String(clampToSliderRange(stepSlider.value));
        }

        /* ∧∨・↑↓キー：スライダーとプレビューを更新 / The stepper and arrow keys update the slider and the preview */
        stepInput.stepOptions.onStep = function () {
            syncSliderFromEditText();
            refreshPreview();
        };

        /* スライダー：入力欄とプレビューを更新 / The slider updates the field and the preview */
        stepSlider.onChanging = stepSlider.onChange = function () {
            updateStepSliderRangeFromKeyboard();
            syncEditTextFromSlider();
            refreshPreview();
        };

        /* 手入力でもプレビュー（数値として読めるときだけ。入力中は書き換えない）
           Preview while typing, only when the text parses; the field is not rewritten per keystroke */
        stepInput.onChanging = function () {
            if (isNaN(parseInt(stepInput.text, 10))) {
                return;
            }
            syncSliderFromEditText();
            refreshPreview();
        };

        /* 確定時（Enter・フォーカス移動）に 0〜上限の整数へ整える / Sanitize on commit */
        stepInput.onChange = function () {
            var stepValue = parseInt(stepInput.text, 10);
            if (isNaN(stepValue)) {
                stepValue = 0;
            }
            stepInput.text = String(clampBlendSteps(stepValue));
            syncSliderFromEditText();
            refreshPreview();
        };
    }

    /**
     * ダイアログのライブプレビューを用意する
     *
     * キャンセルで戻せるよう、プレビューで実行した操作の数を数える。
     * 反転コマンドはトグルなので、前回適用時から状態が変わったときだけ実行する。
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {Object} applyStep / applyReverse / undoAll / keepChanges / wasReverseApplied を持つオブジェクト
     */
    function createBlendPreview(dialogControls) {
        /* プレビューで実行した操作の数（キャンセル時に取り消す）/ Operations to undo on Cancel */
        var previewUndoDepth = 0;

        /* プレビュー対象のブレンド（コマンド実行時に選択されている必要がある）/ Blend the preview acts on */
        var targetBlendItem = null;
        try {
            targetBlendItem = findFirstBlendObject(app.activeDocument.selection);
        } catch (e) { }

        /* 反転の前回適用時の状態と、プレビューで反転を実行したか / Last applied reverse state */
        var lastReverseSpine = dialogControls.reverseSpineCheckbox.value;
        var lastReverseStack = dialogControls.reverseStackCheckbox.value;
        var reversePreviewTouched = false;

        /**
         * 対象のブレンドを特定し、それだけを選択状態にする
         * @returns {void}
         */
        function ensureTargetBlendSelected() {
            try {
                if (!targetBlendItem) {
                    targetBlendItem = findFirstBlendObject(app.activeDocument.selection);
                }
                if (targetBlendItem) {
                    selectOnlyItem(app.activeDocument, targetBlendItem);
                }
            } catch (e) { }
        }

        /**
         * UIの現在値（ステップ数と方向）でブレンド設定を適用し、プレビューを更新する
         * @returns {void}
         */
        function applyStepPreview() {
            var stepValue = parseInt(dialogControls.stepInput.text, 10);
            if (isNaN(stepValue)) {
                return;
            }
            try {
                ensureTargetBlendSelected();
                setBlendOption(clampBlendSteps(stepValue), readOrientation(dialogControls));
                previewUndoDepth++;
                app.redraw();
            } catch (e) { }
        }

        /**
         * 反転チェックボックスの状態をブレンドへ即時プレビューする
         * @returns {void}
         */
        function applyReversePreview() {
            var wantSpine = !!dialogControls.reverseSpineCheckbox.value;
            var wantStack = !!dialogControls.reverseStackCheckbox.value;

            ensureTargetBlendSelected();

            try {
                if (wantSpine !== lastReverseSpine) {
                    app.executeMenuCommand('Path Blend Reverse Spine');
                    previewUndoDepth++;
                    lastReverseSpine = wantSpine;
                    reversePreviewTouched = true;
                }
            } catch (e) { }

            try {
                if (wantStack !== lastReverseStack) {
                    app.executeMenuCommand('Path Blend Reverse Stack');
                    previewUndoDepth++;
                    lastReverseStack = wantStack;
                    reversePreviewTouched = true;
                }
            } catch (e) { }

            app.redraw();
        }

        /**
         * プレビューで実行した操作をすべて取り消す
         * @returns {void}
         */
        function undoAll() {
            try {
                while (previewUndoDepth > 0) {
                    app.executeMenuCommand('undo');
                    previewUndoDepth--;
                }
                app.redraw();
            } catch (e) { }
        }

        return {
            applyStep: applyStepPreview,
            applyReverse: applyReversePreview,
            undoAll: undoAll,
            /* 確定時はプレビューの結果を残す / Keep the preview result on OK */
            keepChanges: function () { previewUndoDepth = 0; },
            wasReverseApplied: function () { return reversePreviewTouched; }
        };
    }

    /**
     * 「その他」「反転」のラジオ・チェックボックスの連動を登録し、初期状態を反映する
     * @param {Window} blendDialog - 対象ダイアログ（淡色表示のペンを作る）
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {PageItem[]} selection - 実行時の選択オブジェクト
     * @param {Object} blendPreview - createBlendPreview() の戻り値
     * @returns {function} 調整モードに合わせて各パネルの状態を更新する関数
     */
    function bindAdjustControls(blendDialog, dialogControls, selection, blendPreview) {
        var adjustRadios = {
            none: dialogControls.adjustNoneRadio,
            release: dialogControls.adjustReleaseRadio,
            expand: dialogControls.adjustExpandRadio,
            replaceSpine: dialogControls.adjustReplaceRadio
        };

        /**
         * 調整モードのラジオボタンを排他的に切り替える
         *
         * ScriptUI は同じ親の中でしか自動排他しないため、パネルをまたぐ排他をここで行う。
         * @param {string} adjustMode - "none" / "release" / "expand" / "replaceSpine" のいずれか
         * @returns {void}
         */
        function selectAdjustMode(adjustMode) {
            adjustRadios.none.value = (adjustMode === 'none');
            adjustRadios.release.value = (adjustMode === 'release');
            adjustRadios.expand.value = (adjustMode === 'expand');
            adjustRadios.replaceSpine.value = (adjustMode === 'replaceSpine');
        }

        var dimPen = null;
        var normalPen = null;

        /**
         * 「その他」が「なし」のとき、他の選択肢のラベルを淡色にする
         *
         * `enabled = false` にすると選択できなくなるため、文字色だけを変える。
         * @returns {void}
         */
        function syncAdjustOptionDimming() {
            try {
                if (!dimPen) {
                    dimPen = blendDialog.graphics.newPen(PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1);
                }
                if (!normalPen) {
                    normalPen = blendDialog.graphics.newPen(PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
                }

                var useDim = !!adjustRadios.none.value;

                /* 無効化中のラジオ（ブレンドのみ選択時の置き換えなど）は常に淡色。有効なものだけ切り替える
                   Disabled radios stay dim; only enabled ones follow the None state */
                var dimmableRadios = [adjustRadios.release, adjustRadios.expand, adjustRadios.replaceSpine];
                for (var i = 0; i < dimmableRadios.length; i++) {
                    var adjustRadio = dimmableRadios[i];
                    if (!adjustRadio) continue;

                    if (adjustRadio.enabled === false) {
                        adjustRadio.graphics.foregroundColor = dimPen;
                        continue;
                    }

                    adjustRadio.graphics.foregroundColor = useDim ? dimPen : normalPen;
                }
            } catch (e) { }
        }

        /**
         * 選択内容に応じて「ブレンド軸を置き換え」の有効／無効を更新する
         *
         * 選択がブレンドのみのときは無効にし、選択済みなら「なし」へ戻す。
         * @returns {void}
         */
        function updateAdjustAvailability() {
            try {
                adjustRadios.replaceSpine.enabled = !isOnlyBlendObjects(selection);
                if (!adjustRadios.replaceSpine.enabled && adjustRadios.replaceSpine.value) {
                    selectAdjustMode('none');
                }
                syncAdjustOptionDimming();
            } catch (e) { }
        }

        /**
         * 「その他」が「なし」以外のとき、方向パネルを無効化し、フォーカスを移す
         * @returns {void}
         */
        function syncBlendPanelEnabled() {
            var isAdjustNone = !!adjustRadios.none.value;
            dialogControls.orientationPanel.enabled = isAdjustNone;

            /* フォーカスを分かりやすい位置へ / Keep focus sensible */
            try {
                if (isAdjustNone) {
                    dialogControls.stepInput.active = true;
                    dialogControls.stepInput.setSelection(0, dialogControls.stepInput.text.length);
                } else {
                    /* 選択中の選択肢にフォーカス / Focus the selected option */
                    if (adjustRadios.release.value) adjustRadios.release.active = true;
                    else if (adjustRadios.expand.value) adjustRadios.expand.active = true;
                    else if (adjustRadios.replaceSpine.value) adjustRadios.replaceSpine.active = true;
                    else adjustRadios.none.active = true;
                }
            } catch (e) { }
        }

        /**
         * 「その他」が「なし」以外のとき、反転パネルを無効化する
         *
         * チェックボックスの値は書き換えない（プレビュー済みの反転は OK／キャンセルの処理まで残す）。
         * @returns {void}
         */
        function syncReversePanelEnabled() {
            dialogControls.reversePanel.enabled = !!adjustRadios.none.value;
        }

        /**
         * 調整モードに合わせて各パネルの状態を更新する
         * @returns {void}
         */
        function refreshAdjustState() {
            updateAdjustAvailability();
            syncBlendPanelEnabled();
            syncReversePanelEnabled();
        }

        /* 「なし」以外を選ぶと方向・反転パネルを無効化 / Choosing an adjustment disables the other panels */
        for (var adjustMode in adjustRadios) {
            if (!adjustRadios.hasOwnProperty(adjustMode)) continue;
            (function (radioMode) {
                adjustRadios[radioMode].onClick = function () {
                    selectAdjustMode(radioMode);
                    refreshAdjustState();
                };
            })(adjustMode);
        }

        /* 反転は独立した操作。混乱を避けるため「その他」は「なし」に戻す
           Reverse actions are independent; keep Misc on None to avoid mixed modes */
        dialogControls.reverseSpineCheckbox.onClick = dialogControls.reverseStackCheckbox.onClick = function () {
            selectAdjustMode('none');
            refreshAdjustState();
            blendPreview.applyReverse();
        };

        refreshAdjustState();
        return refreshAdjustState;
    }

    /**
     * 設定ダイアログを組み立てて表示し、確定した入力値を返す
     * @param {PageItem[]} selection - 実行時の選択オブジェクト
     * @param {boolean} wasBlendSelectedAtOpen - ダイアログを開いた時点でブレンドを選択していたか
     * @returns {Object|null} ステップ数・方向・反転・調整モードを持つ入力値。キャンセル時は null
     */
    function showBlendDialog(selection, wasBlendSelectedAtOpen) {
        var doc = app.activeDocument;

        /* ダイアログ開始時の選択を保存（Replace Spine などで「ブレンド＋パス」の混在選択を維持するため）
           Keep the launch selection; Replace Spine needs the blend and the path selected together */
        var selectionAtOpen = snapshotSelection(doc);
        var defaultStep = readDefaultStep(selection, wasBlendSelectedAtOpen);

        var blendDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        blendDialog.orientation = 'column';
        blendDialog.alignChildren = ['fill', 'top'];
        blendDialog.spacing = DIALOG_SPACING;
        blendDialog.margins = DIALOG_MARGINS;

        var stepControls = buildStepArea(blendDialog, defaultStep);
        var dialogControls = buildOptionColumns(blendDialog);
        dialogControls.stepInput = stepControls.stepInput;
        dialogControls.stepSlider = stepControls.stepSlider;

        var blendPreview = createBlendPreview(dialogControls);
        bindStepControls(dialogControls.stepInput, dialogControls.stepSlider, blendPreview.applyStep);
        dialogControls.alignToPageRadio.onClick = blendPreview.applyStep;
        dialogControls.alignToPathRadio.onClick = blendPreview.applyStep;
        var refreshAdjustState = bindAdjustControls(blendDialog, dialogControls, selection, blendPreview);

        var buttonRow = addButtonRow(blendDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        var dialogResult = null;

        blendDialog.onShow = function () {
            blendPreview.applyStep();
            refreshAdjustState();
        };

        btnOK.onClick = function () {
            var reverseSpine = !!dialogControls.reverseSpineCheckbox.value;
            var reverseStack = !!dialogControls.reverseStackCheckbox.value;

            var stepValue = parseInt(dialogControls.stepInput.text, 10);
            if (isNaN(stepValue) || stepValue < 0 || stepValue > MAX_BLEND_STEPS) {
                alert(getLabel('alert.invalidStep'));
                return;
            }
            var adjustMode = readAdjustMode(dialogControls);

            /* Replace Spine は「ブレンド＋置き換え用パス」の同時選択が必要。
               プレビュー中にブレンド単体選択へ切り替わっていることがあるため、ここで元の選択を復元する。
               Replace Spine needs the blend and the new path selected; restore the launch selection */
            if (adjustMode === 'replaceSpine') {
                restoreSelection(doc, selectionAtOpen);
            }

            /* プレビューで反転済みなら文書は既に目的の状態。true を渡すと main() で再度トグルして打ち消してしまう
               If the reverse preview already ran, passing true would toggle it back in main() */
            var reverseApplied = blendPreview.wasReverseApplied();
            dialogResult = {
                step: stepValue,
                orientation: readOrientation(dialogControls),
                adjustMode: adjustMode,
                reverseSpine: reverseApplied ? false : reverseSpine,
                reverseStack: reverseApplied ? false : reverseStack
            };
            blendPreview.keepChanges();
            blendDialog.close(1);
        };

        btnCancel.onClick = function () {
            /* このダイアログで適用したプレビューをすべて戻す / Revert every preview change */
            blendPreview.undoAll();
            /* 取り消し時も、ダイアログ開始時の選択に戻す / Restore the launch selection */
            restoreSelection(doc, selectionAtOpen);
            blendDialog.close(0);
        };

        prepareDialogWindow(blendDialog, SCRIPT_NAME);
        if (blendDialog.show() !== 1) {
            return null;
        }
        return dialogResult;
    }

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
     * ステップ数と方向を指定して、ブレンドオプションを一時アクションで適用する
     * @param {number} step - 中間ステップ数
     * @param {number} orientationValue - ブレンドの方向を表す値
     * @returns {void}
     */
    function setBlendOption(step, orientationValue) {
        var actionCode = [
            "/version 3",
            "/name [ 5",
            "	426c656e64",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "	/name [ 7",
            "		73657453746570",
            "	]",
            "	/keyIndex 0",
            "	/colorIndex 0",
            "	/isOpen 1",
            "	/eventCount 1",
            "	/event-1 {",
            "		/useRulersIn1stQuadrant 0",
            "		/internalName (ai_plugin_liveblend)",
            "		/localizedName [ 12",
            "			e38396e383ace383b3e38389",
            "		]",
            "		/isOpen 0",
            "		/isOn 1",
            "		/hasDialog 1",
            "		/showDialog 0",
            "		/parameterCount 3",
            "		/parameter-1 {",
            "			/key 1835363957",
            "			/showInPalette 4294967295",
            "			/type (enumerated)",
            "			/name [ 15",
            "				e382aae38397e382b7e383a7e383b3",
            "			]",
            "			/value 5",
            "		}",
            "		/parameter-2 {",
            "			/key 1937007984",
            "			/showInPalette 4294967295",
            "			/type (integer)",
            "			/value " + String(step),
            "		}",
            "		/parameter-3 {",
            "			/key 1919906913",
            "			/showInPalette 4294967295",
            "			/type (enumerated)",
            "			/name [ 12",
            "				e59e82e79bb4e696b9e59091",
            "			]",
            "			/value " + String(orientationValue),
            "		}",
            "	}",
            "}"
        ].join("\n");

        if (!runTemporaryAction(actionCode, "Blend", "setStep")) {
            alert(getLabel('alert.actionFailed'));
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確認し、ダイアログを開いてブレンドの作成・設定・調整を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length <= 0) {
            return;
        }
        var doc = app.activeDocument;
        var selection = doc.selection;

        /* ダイアログを開いた時点で「ブレンドオブジェクトを選択していたか」を記録
           （この後 Path Blend Make を実行して選択がブレンドに変わることがあるため）
           Remember whether a blend was selected at launch; Path Blend Make may change the selection */
        var wasBlendSelectedAtOpen = !!findFirstBlendObject(selection);

        /* 選択にブレンド（PluginItem/BlendItem）が含まれない場合のみ、ブレンドを作成
           （既に含まれる場合は、パスが混ざっていても選択はそのまま。作成せずオプション/調整のみ行う）
           Make a blend only when none is selected; otherwise keep the selection as is */
        if (!wasBlendSelectedAtOpen) {
            try {
                app.executeMenuCommand('Path Blend Make');
            } catch (e) { }

            /* コマンド実行後に選択が変わる可能性があるため、再取得 / The command may change the selection */
            selection = doc.selection;
            if (!selection || selection.length <= 0) {
                return;
            }

            /* 生成された PluginItem（ブレンド）だけを選択状態にする / Select only the new blend */
            var createdBlend = findFirstBlendObject(selection);
            if (createdBlend) {
                selectOnlyItem(doc, createdBlend);
                selection = doc.selection;
            }
        }

        /* ダイアログを表示してユーザー入力を取得（UI操作専用。引数実行は行わない）
           Show the dialog; this script is UI-only */
        var dialogInput = showBlendDialog(selection, wasBlendSelectedAtOpen);
        if (!dialogInput) {
            return; /* キャンセル / cancelled */
        }

        var adjustMode = dialogInput.adjustMode || 'none';
        if (adjustMode === 'release') {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Release');
            return;
        }
        if (adjustMode === 'expand') {
            app.executeMenuCommand('Path Blend Expand');
            return;
        }
        if (adjustMode === 'replaceSpine') {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Replace Spine');
            return;
        }

        var reverseSpine = !!dialogInput.reverseSpine;
        var reverseStack = !!dialogInput.reverseStack;
        if (reverseSpine) {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Reverse Spine');
        }
        if (reverseStack) {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Reverse Stack');
        }
        if (reverseSpine || reverseStack) {
            return;
        }

        /* オプションだけを適用（新しいブレンドは作らない）/ Apply options only; no new blend */
        setBlendOption(dialogInput.step, dialogInput.orientation);
    }

    main();

})();
