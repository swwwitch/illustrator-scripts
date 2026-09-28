#target illustrator
#targetengine "ConvertToAreaTypeLikeButtonEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToAreaTypeLikeButton.md

### Overview

Creates and adjusts area text from point text, text on a path, shapes, or existing area text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToAreaTypeLikeButton.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertToAreaTypeLikeButton";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToAreaTypeLikeButton.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToAreaTypeLikeButton.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ボタン風の拡大倍率 / Button-style expansion ratios */
    var BUTTON_WIDTH_RATIO  = 1.2;  /* 元の幅に対する倍率 / Ratio of original width */
    var BUTTON_HEIGHT_RATIO = 1.6;  /* 元の高さに対する倍率 / Ratio of original height */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS        = 20;                /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS         = [16, 20, 16, 12];  /* パネルの余白 / panel margins */
    var PANEL_SPACING         = 8;                 /* パネル内の間隔 / panel spacing */
    var SIZE_LABEL_WIDTH      = 44;                /* 幅・高さのラベル幅 / width of the width/height labels */
    var INDENT_CHECKBOX_WIDTH = 52;                /* 左右インデントのチェックボックス幅 / width of the indent checkboxes */

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は PANEL_SPACING）
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
     * グループの共通設定を適用する（row/column で整列を切り替え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    /* 日英ラベル定義（カテゴリ構造）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "AreaType Toolkit", en: "AreaType Toolkit" }
        },
        /* パネル見出し / Panel titles */
        panel: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            frameSize: { ja: "フレームサイズ", en: "Frame size" },
            indent: { ja: "インデント", en: "Indent" },
            options: { ja: "オプション", en: "Options" }
        },
        /* 入力ラベル / Field labels */
        fieldLabel: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        /* チェックボックス / Checkboxes */
        checkbox: {
            indentLeft: { ja: "左", en: "Left" },
            indentRight: { ja: "右", en: "Right" },
            margin: { ja: "外側からの間隔", en: "Spacing" }
        },
        /* ボタン / Buttons */
        button: {
            overset: { ja: "文字あふれ解消", en: "Make overset" },
            run: { ja: "実行", en: "Run" },
            close: { ja: "閉じる", en: "Close" }
        },
        /* ツールチップ / Tooltips */
        tooltip: {
            fontSize: { ja: "エリア内文字のフォントサイズです。", en: "Font size of the area text." },
            overset: { ja: "文字があふれないところまでフォントサイズを下げます。", en: "Lowers the font size until the text no longer overflows." },
            width: { ja: "テキストフレームの幅です。", en: "Width of the text frame." },
            height: { ja: "テキストフレームの高さです。", en: "Height of the text frame." },
            indentLeft: { ja: "段落の左インデントを設定します。", en: "Sets the left indent of the paragraphs." },
            indentRight: { ja: "段落の右インデントを設定します。", en: "Sets the right indent of the paragraphs." },
            indentValue: { ja: "インデントの量です。", en: "Amount of the indent." },
            sync: { ja: "左右のインデントを同じ値にします。", en: "Keeps the left and right indents the same." },
            margin: { ja: "テキストフレームの内側に空ける余白です。", en: "Inset kept inside the text frame." },
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
        /* 警告メッセージ / Alerts */
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字・エリア内文字を選択してください。",
                en: "Please select point text, path text, or area text."
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 選択の判定 / Selection checks
    // =========================================

    /**
     * パス上文字かを返す
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} パス上文字なら true
     */
    function isPathTextFrame(pageItem) {
        return pageItem.typename === "TextFrame" && pageItem.kind === TextType.PATHTEXT;
    }

    /**
     * エリア内文字かを返す
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} エリア内文字なら true
     */
    function isAreaTextFrame(pageItem) {
        return pageItem.typename === "TextFrame" && pageItem.kind === TextType.AREATEXT;
    }

    /**
     * 選択にポイント文字（パス上文字を含む）・エリア内文字・パスが含まれるかを調べる
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {{hasPointText: boolean, hasAreaText: boolean, hasPathItem: boolean}} 含まれる種類
     */
    function classifySelection(selectedItems) {
        var selectionKinds = { hasPointText: false, hasAreaText: false, hasPathItem: false };
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem.typename === "TextFrame") {
                if (selectedItem.kind === TextType.POINTTEXT || selectedItem.kind === TextType.PATHTEXT) selectionKinds.hasPointText = true;
                if (selectedItem.kind === TextType.AREATEXT) selectionKinds.hasAreaText = true;
            }
            if (selectedItem.typename === "PathItem" || selectedItem.typename === "CompoundPathItem") {
                selectionKinds.hasPathItem = true;
            }
        }
        return selectionKinds;
    }

    /**
     * 選択から最初のエリア内文字を返す
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {TextFrame|null} 最初のエリア内文字（無ければ null）
     */
    function findFirstAreaText(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (isAreaTextFrame(selectedItems[i])) return selectedItems[i];
        }
        return null;
    }

    // =========================================
    // パス上文字 → ポイント文字（変換前処理）/ Path text → Point text (pre-process)
    // =========================================

    /**
     * 関数を実行し、例外は握りつぶす（属性ごとに失敗しても残りを続けるため）
     * @param {Function} attemptAction - 実行する処理
     * @returns {*} 処理の戻り値（例外時は undefined）
     */
    function runIgnoringErrors(attemptAction) {
        try { return attemptAction(); } catch (e) { return undefined; }
    }

    /**
     * 文字ごとの属性を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} 文字ごとの属性
     */
    function snapshotCharacterAttributes(textFrame) {
        var attributeSnapshots = [];
        for (var i = 0; i < textFrame.characters.length; i++) {
            var sourceAttributes = textFrame.characters[i].characterAttributes;
            attributeSnapshots.push({
                font: sourceAttributes.textFont,
                size: sourceAttributes.size,
                fillColor: sourceAttributes.fillColor,
                strokeColor: sourceAttributes.strokeColor,
                strokeWeight: sourceAttributes.strokeWeight,
                autoLeading: sourceAttributes.autoLeading,
                leading: sourceAttributes.leading
            });
        }
        return attributeSnapshots;
    }

    /**
     * 控えた文字属性を書き戻す（ベースライン移動と比率はリセット）。属性ごとに失敗しても続ける
     * @param {TextFrame} textFrame - 書き戻し先のテキストフレーム
     * @param {Object[]} attributeSnapshots - snapshotCharacterAttributes() の戻り値
     * @returns {void}
     */
    function restoreCharacterAttributes(textFrame, attributeSnapshots) {
        var restoreCount = Math.min(textFrame.characters.length, attributeSnapshots.length);
        for (var i = 0; i < restoreCount; i++) {
            var targetAttributes = textFrame.characters[i].characterAttributes;
            var savedAttributes = attributeSnapshots[i];

            runIgnoringErrors(function () { targetAttributes.textFont = savedAttributes.font; });
            runIgnoringErrors(function () { targetAttributes.size = savedAttributes.size; });
            runIgnoringErrors(function () { targetAttributes.fillColor = savedAttributes.fillColor; });
            runIgnoringErrors(function () {
                var savedStrokeColor = savedAttributes.strokeColor;
                targetAttributes.strokeColor = savedStrokeColor;
                targetAttributes.strokeWeight = (savedStrokeColor && savedStrokeColor.typename === "NoColor") ? 0 : savedAttributes.strokeWeight;
            });
            runIgnoringErrors(function () { targetAttributes.baselineShift = 0; });
            runIgnoringErrors(function () { targetAttributes.horizontalScale = 100; });
            runIgnoringErrors(function () { targetAttributes.verticalScale = 100; });
            runIgnoringErrors(function () { targetAttributes.autoLeading = savedAttributes.autoLeading; });
            if (!savedAttributes.autoLeading) runIgnoringErrors(function () { targetAttributes.leading = savedAttributes.leading; });
        }
    }

    /**
     * パス上文字を字形を保ったままポイント文字へ分離する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} pathTextFrames - 分離するパス上文字
     * @returns {TextFrame[]} 作成したポイント文字
     */
    function detachPathTextToPointText(doc, pathTextFrames) {
        var createdPointTexts = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return createdPointTexts;

        /* 新規テキストだけ選べるよう選択を解除 / Clear selection */
        runIgnoringErrors(function () { doc.selection = null; });

        for (var j = pathTextFrames.length - 1; j >= 0; j--) {
            var pathText = pathTextFrames[j];
            if (!pathText || !isPathTextFrame(pathText)) continue;

            var originalPath = null;
            runIgnoringErrors(function () { originalPath = pathText.textPath; });
            if (!originalPath) continue;

            /* 1) 文字ごとの属性を退避 / Snapshot per-character attributes */
            var attributeSnapshots = snapshotCharacterAttributes(pathText);

            var textContents = "";
            runIgnoringErrors(function () { textContents = pathText.contents; });

            var justification = null;
            runIgnoringErrors(function () {
                if (pathText.paragraphs && pathText.paragraphs.length > 0) {
                    justification = pathText.paragraphs[0].paragraphAttributes.justification;
                }
            });

            /* 2) パス始点にポイント文字を新規作成 / Create new point text at path start anchor */
            var pointText = doc.textFrames.add();
            var anchorPoint = null;
            runIgnoringErrors(function () {
                if (originalPath.pathPoints && originalPath.pathPoints.length > 0) {
                    anchorPoint = originalPath.pathPoints[0].anchor;
                }
            });
            if (anchorPoint) {
                pointText.position = [anchorPoint[0], anchorPoint[1]];
            }

            pointText.contents = textContents;

            if (justification !== null && pointText.paragraphs && pointText.paragraphs.length > 0) {
                runIgnoringErrors(function () { pointText.paragraphs[0].paragraphAttributes.justification = justification; });
            }

            /* 既定の線を一旦消し、後で文字ごとに復元 / Clear default stroke, restore per-character later */
            runIgnoringErrors(function () {
                pointText.textRange.characterAttributes.strokeColor = new NoColor();
                pointText.textRange.characterAttributes.strokeWeight = 0;
            });

            /* 文字ごとの属性を復元 / Restore per-character attributes */
            restoreCharacterAttributes(pointText, attributeSnapshots);

            /* 3) 元のパス上文字を削除（パスも一緒に消える）/ Remove original path text */
            runIgnoringErrors(function () { pathText.remove(); });

            /* 4) 新規テキストを選択して返す / Select and return new text */
            runIgnoringErrors(function () { pointText.selected = true; });
            createdPointTexts.push(pointText);
        }

        return createdPointTexts;
    }

    /**
     * 選択内のパス上文字をポイント文字へ置き換え、置き換えた選択配列を返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {PageItem[]} 置き換え後の選択（パス上文字が無ければ元の選択）
     */
    function preprocessPathTextSelection(doc, currentSelection) {
        if (!doc || !currentSelection || !currentSelection.length) return currentSelection;

        var pathTexts = [];
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            /* 無効オブジェクト（削除済み等）は読めない / Invalid (deleted) objects cannot be read */
            try {
                if (selectedItem && isPathTextFrame(selectedItem)) pathTexts.push(selectedItem);
            } catch (e0) { }
        }
        if (!pathTexts.length) return currentSelection;

        var createdPointTexts = detachPathTextToPointText(doc, pathTexts);
        if (!createdPointTexts.length) return currentSelection;

        /* パス上文字を新ポイント文字に差し替えた新しい選択配列を構築 / Build replaced selection array */
        var replacedSelection = [];
        for (var j = 0; j < currentSelection.length; j++) {
            var remainingItem = currentSelection[j];
            /* 削除したパス上文字は読めないので飛ばす / Removed path text cannot be read, so it is skipped */
            try {
                if (remainingItem && !isPathTextFrame(remainingItem)) replacedSelection.push(remainingItem);
            } catch (e1) { }
        }
        for (var k = 0; k < createdPointTexts.length; k++) replacedSelection.push(createdPointTexts[k]);

        try { doc.selection = replacedSelection; } catch (e) { }
        app.redraw();

        return replacedSelection;
    }

    // =========================================
    // オーバーセット判定・文字サイズ調整 / Overset detection & font sizing
    // =========================================

    /**
     * overflows プロパティを安全に取得する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean|null} あふれていれば true（読めなければ null）
     */
    function readOverflows(textFrame) {
        try {
            if (textFrame && typeof textFrame.overflows !== "undefined") return !!textFrame.overflows;
        } catch (e) { }
        return null;
    }

    /**
     * 表示されている行に収まらない文字があるか判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function hasHiddenCharacters(textFrame) {
        var lineCount = textFrame.lines.length;
        if (lineCount === 0) return textFrame.characters.length > 0;
        var visibleCharacters = 0;
        for (var i = 0; i < lineCount; i++) { visibleCharacters += textFrame.lines[i].characters.length; }
        return visibleCharacters < textFrame.characters.length;
    }

    /**
     * テキストフレームがあふれているか判定する（エリア内文字は overflows を優先）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function isOversetFrame(textFrame) {
        if (textFrame && textFrame.kind === TextType.AREATEXT) {
            var overflowState = readOverflows(textFrame);
            if (overflowState !== null) return overflowState;
        }
        try { return hasHiddenCharacters(textFrame); } catch (e) { return false; }
    }

    /**
     * 行送り比率（行送り/サイズ）を取得する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{ratio: number}|null} 行送りの比率（自動行送りなどのときは null）
     */
    function getLeadingInfo(textFrame) {
        try {
            var textAttributes = textFrame.textRange.characterAttributes;
            if (textAttributes.autoLeading) return null;
            var fontSize = textAttributes.size, leading = textAttributes.leading;
            if (fontSize > 0 && leading > 0) return { ratio: leading / fontSize };
        } catch (e) { }
        return null;
    }

    /**
     * 比率を保ったまま行送りを更新する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} newSize - 変更後の文字サイズ（pt）
     * @param {{ratio: number}|null} leadingInfo - getLeadingInfo() の戻り値
     * @returns {void}
     */
    function applyLeading(textFrame, newSize, leadingInfo) {
        if (!leadingInfo) return;
        try { textFrame.textRange.characterAttributes.leading = newSize * leadingInfo.ratio; } catch (e) { }
    }

    /**
     * 文字サイズを設定し、行送りを比率に合わせる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {{ratio: number}|null} leadingInfo - getLeadingInfo() の戻り値
     * @returns {void}
     */
    function setFontSizeKeepingLeading(textFrame, fontSize, leadingInfo) {
        textFrame.textRange.characterAttributes.size = fontSize;
        applyLeading(textFrame, fontSize, leadingInfo);
    }

    /**
     * あふれなくなる最大サイズを二分探索で探して縮小する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function shrinkFont(textFrame) {
        if (textFrame.characters.length <= 0 || !isOversetFrame(textFrame)) return;
        var leadingInfo = getLeadingInfo(textFrame);
        var upperSize = textFrame.textRange.characterAttributes.size;
        var lowerSize = 0.1;

        /* 最小でもあふれるならそのまま終了 / If even lowerSize overflows, keep min size */
        setFontSizeKeepingLeading(textFrame, lowerSize, leadingInfo);
        if (isOversetFrame(textFrame)) return;

        /* 二分探索 / Binary search */
        for (var i = 0; i < 40; i++) {
            var middleSize = (lowerSize + upperSize) / 2;
            setFontSizeKeepingLeading(textFrame, middleSize, leadingInfo);
            if (isOversetFrame(textFrame)) {
                upperSize = middleSize;
            } else {
                lowerSize = middleSize;
            }
            if (upperSize - lowerSize < 0.1) break;
        }

        /* あふれない側（lowerSize）に確定 / Settle on the non-overset side */
        setFontSizeKeepingLeading(textFrame, lowerSize, leadingInfo);
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

    // =========================================
    // ダイナミックアクション / Dynamic actions
    // =========================================

    /**
     * /name ブロック（ASCII）を生成する
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(actionName) {
        return "/name [ " + actionName.length + " " + toActionHex(actionName).toUpperCase() + " ]";
    }

    /**
     * アクションセット定義（.aia）文字列を組み立てる
     * @param {string} setName - アクションセット名
     * @param {string} internalName - イベントの内部名
     * @param {string} localizedNameHex - ローカライズ名（長さと16進。空なら省略）
     * @param {number} parameterKey - パラメーターのキー
     * @param {Array<{name: string, value: number}>} actionDefinitions - アクション名と値の組
     * @returns {string} .aia 形式のアクションセット定義
     */
    function buildActionSetAia(setName, internalName, localizedNameHex, parameterKey, actionDefinitions) {
        var aiaText = "/version 3" +
            buildActionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefinitions.length;

        for (var i = 0; i < actionDefinitions.length; i++) {
            var actionDef = actionDefinitions[i];
            aiaText += "/action-" + (i + 1) + " {" +
                " " + buildActionNameBlock(actionDef.name) +
                " /keyIndex 0" +
                " /colorIndex 0" +
                " /isOpen 1" +
                " /eventCount 1" +
                " /event-1 {" +
                " /useRulersIn1stQuadrant 0" +
                " /internalName (" + internalName + ")" +
                (localizedNameHex ? (" /localizedName [ " + localizedNameHex + " ]") : "") +
                " /isOpen 0" +
                " /isOn 1" +
                " /hasDialog 0" +
                " /parameterCount 1" +
                " /parameter-1 {" +
                " /key " + parameterKey +
                " /showInPalette 4294967295" +
                " /type (integer)" +
                " /value " + actionDef.value +
                " }" +
                " }" +
                "}";
        }
        return aiaText;
    }

    /* フレーム整列アクションセット名 / Frame-alignment action set name */
    var AREA_TEXT_ACTION_SET = "AreaText";

    /* フレーム整列のアクション名（添字がアクションの値：0=上, 1=中央, 2=下, 3=均等）
       Frame-alignment action names; the index is the action value (0=top, 1=center, 2=bottom, 3=justify) */
    var FRAME_ALIGNMENT_ACTIONS = ["AlignTop", "AlignCenter", "AlignBottom", "AlignJustify"];

    /**
     * フレーム整列アクション（AlignTop/Center/Bottom/Justify）を読み込む
     * @returns {void}
     */
    function loadAreaTextActions() {
        var actionDefinitions = [];
        for (var i = 0; i < FRAME_ALIGNMENT_ACTIONS.length; i++) {
            actionDefinitions.push({ name: FRAME_ALIGNMENT_ACTIONS[i], value: i });
        }
        var aiaText = buildActionSetAia(
            AREA_TEXT_ACTION_SET,
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            actionDefinitions
        );
        /* 読み込めなくても変換・調整は続ける（整列だけ効かない）/ Keep going without the set; only the alignment is skipped */
        loadTemporaryActionSet(aiaText, AREA_TEXT_ACTION_SET);
    }

    /**
     * フレーム整列アクションを破棄する
     * @returns {void}
     */
    function unloadAreaTextActions() {
        unloadTemporaryActionSet(AREA_TEXT_ACTION_SET);
    }

    /**
     * テキストの配置（垂直方向のフレーム整列）をアクションで変更する
     * @param {number} alignValue - 0=上 / 1=中央 / 2=下 / 3=均等
     * @returns {void}
     */
    function runFrameAlignmentAction(alignValue) {
        if (alignValue !== 0 && alignValue !== 1 && alignValue !== 2 && alignValue !== 3) return;
        try { app.doScript(FRAME_ALIGNMENT_ACTIONS[alignValue], AREA_TEXT_ACTION_SET, false); } catch (e) { }
    }

    /**
     * 指定フレームにフレーム整列を適用する（プレビュー中はスキップ）
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @param {number} alignValue - 0=上 / 1=中央 / 2=下 / 3=均等
     * @param {boolean} forPreview - プレビュー中か
     * @returns {void}
     */
    function applyAreaTextFrameAlignment(areaTextFrame, alignValue, forPreview) {
        /* app.doScript はプレビュー中に呼ぶと不安定なためスキップ / Unstable during preview, so skip */
        if (forPreview) return;
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaTextFrame];
            app.redraw(); /* 選択状態を確定 / Commit the selection */
            runFrameAlignmentAction(alignValue);
        } catch (e) { }
    }

    // =========================================
    // UIユーティリティ / UI utilities
    // =========================================

    /**
     * 行に、左に∧∨を付けた数値欄を追加する（↑↓キーも∧∨と同じ処理で増減する。負の値にはしない）。
     * 増減後の処理は、あとから input.stepOptions.onStep に入れる
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 初期値
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup、設定は .stepOptions で参照できる）
     */
    function addStepperInput(parentRow, initialText) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = parentRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;
        var stepOptions = { min: 0 };
        var numberInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperFieldGroup.add("edittext", undefined, initialText);
        numberInput.stepperGroup = stepperGroup;
        numberInput.stepOptions = stepOptions;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * addStepperInput() で作った入力欄の有効／無効を、∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - 対象の入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * 幅・高さの行（ラベル＋入力欄＋単位）を追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelKey - LABELS.fieldLabel と LABELS.tooltip のキー
     * @param {string} unitLabel - 単位の表示
     * @returns {EditText} 追加した入力欄
     */
    function addSizeRow(parentContainer, labelKey, unitLabel) {
        var sizeRowGroup = parentContainer.add("group");
        var sizeLabel = sizeRowGroup.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        sizeLabel.preferredSize.width = SIZE_LABEL_WIDTH;
        var sizeInput = addStepperInput(sizeRowGroup, "");
        sizeInput.characters = 5;
        sizeInput.helpTip = getLabel("tooltip." + labelKey);
        sizeRowGroup.add("statictext", undefined, unitLabel);
        return sizeInput;
    }

    /**
     * インデントの行（チェックボックス＋入力欄＋単位）を追加する（入力欄は無効で始まる）
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelKey - LABELS.checkbox と LABELS.tooltip のキー
     * @param {string} unitLabel - 単位の表示
     * @returns {{checkbox: Checkbox, input: EditText}} 追加したチェックボックスと入力欄
     */
    function addIndentRow(parentContainer, labelKey, unitLabel) {
        var indentRowGroup = parentContainer.add("group");
        var indentCheckbox = indentRowGroup.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        indentCheckbox.preferredSize.width = INDENT_CHECKBOX_WIDTH;
        indentCheckbox.helpTip = getLabel("tooltip." + labelKey);
        var indentInput = addStepperInput(indentRowGroup, "0");
        indentInput.characters = 4;
        indentInput.helpTip = getLabel("tooltip.indentValue");
        indentRowGroup.add("statictext", undefined, unitLabel);
        setStepperInputEnabled(indentInput, false);
        return { checkbox: indentCheckbox, input: indentInput };
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
    // 変換 / Conversion
    // =========================================

    /**
     * テキストのフォントとサイズを読む
     * @param {TextFrame} textFrame - 読み取り元
     * @returns {{font: TextFont|null, size: number}} フォントとサイズ（読めなければ null / 0）
     */
    function readFontAndSize(textFrame) {
        var fontStyle = { font: null, size: 0 };
        try {
            fontStyle.font = textFrame.textRange.characterAttributes.textFont;
            fontStyle.size = textFrame.textRange.characterAttributes.size;
        } catch (e) { }
        return fontStyle;
    }

    /**
     * フォントとサイズを適用する（未インストールのフォントなどは失敗しうる）
     * @param {TextFrame} textFrame - 適用先
     * @param {{font: TextFont|null, size: number}} fontStyle - readFontAndSize() の戻り値
     * @returns {void}
     */
    function applyFontAndSize(textFrame, fontStyle) {
        try {
            if (fontStyle.font) textFrame.textRange.characterAttributes.textFont = fontStyle.font;
            if (fontStyle.size > 0) textFrame.textRange.characterAttributes.size = fontStyle.size;
        } catch (e) { }
    }

    /**
     * 選択からモードを判定して変換し、調整ダイアログを開く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {void}
     */
    function convertToAreaTypeAndAdjust(doc, currentSelection) {
        preprocessPathTextSelection(doc, currentSelection);

        var updatedSelection = doc.selection;
        if (!updatedSelection || updatedSelection.length === 0) { return; }

        /* 選択内容に応じてモードを決定（ポイント文字は常にボタン風）/ Decide mode (point text is always button style) */
        var selectionKinds = classifySelection(updatedSelection);
        if (!selectionKinds.hasPointText) return;

        var createdFrames = selectionKinds.hasPathItem
            ? createAreaTextFromTextAndShape(doc, updatedSelection)
            : createButtonAreaTexts(doc, updatedSelection);
        if (createdFrames.length > 0) {
            doc.selection = createdFrames;
            app.redraw();
            showAdjustDialog(doc, createdFrames[0], createdFrames);
        }
    }

    /**
     * 選択したテキストと閉じたパスから、エリア内文字を1つ作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {TextFrame[]} 作成したエリア内文字（作れなければ空）
     */
    function createAreaTextFromTextAndShape(doc, selectedItems) {
        var createdFrames = [];
        var sourceText = null, destPath = null;
        for (var i = 0; i < selectedItems.length; i++) {
            if (!sourceText && selectedItems[i].typename === "TextFrame") { sourceText = selectedItems[i]; }
            else if (!destPath && selectedItems[i].typename === "PathItem" && selectedItems[i].closed) { destPath = selectedItems[i]; }
        }
        if (!sourceText || !destPath) return createdFrames;

        /* 種類によってはエリア内文字にできない / Some shapes cannot become area type */
        try {
            var sourceContents = sourceText.contents;
            var sourceStyle = readFontAndSize(sourceText);
            var framePath = destPath.duplicate();
            framePath.filled = false; framePath.stroked = false;
            var areaTextFrame = doc.textFrames.areaText(framePath);
            areaTextFrame.contents = sourceContents;
            applyFontAndSize(areaTextFrame, sourceStyle);
            sourceText.remove();
            destPath.remove();
            createdFrames.push(areaTextFrame);
        } catch (e) { }
        return createdFrames;
    }

    /**
     * ポイント文字ごとに、幅×BUTTON_WIDTH_RATIO・高さ×BUTTON_HEIGHT_RATIO の枠で中央揃えのエリア内文字を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {TextFrame[]} 作成したエリア内文字
     */
    function createButtonAreaTexts(doc, selectedItems) {
        var createdFrames = [];
        for (var i = selectedItems.length - 1; i >= 0; i--) {
            var pointText = selectedItems[i];
            if (pointText.typename !== "TextFrame" || pointText.kind !== TextType.POINTTEXT) continue;
            /* 1件で失敗しても残りを処理する / One failure must not stop the rest */
            try {
                var textBounds = pointText.geometricBounds;
                var originalWidth = textBounds[2] - textBounds[0], originalHeight = textBounds[1] - textBounds[3];
                var buttonWidth = originalWidth * BUTTON_WIDTH_RATIO, buttonHeight = originalHeight * BUTTON_HEIGHT_RATIO;
                var buttonRect = doc.pathItems.rectangle(
                    textBounds[1] + (buttonHeight - originalHeight) / 2, textBounds[0] - (buttonWidth - originalWidth) / 2, buttonWidth, buttonHeight);
                buttonRect.filled = false; buttonRect.stroked = false;
                var textContents = pointText.contents;
                var sourceStyle = readFontAndSize(pointText);
                var areaTextFrame = doc.textFrames.areaText(buttonRect);
                areaTextFrame.contents = textContents;
                applyFontAndSize(areaTextFrame, sourceStyle);
                /* 行揃えを中央に / Horizontal center */
                try { areaTextFrame.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
                /* 縦位置はDOMで不安定なためアクションで中央 / Vertical center via dynamic action (unreliable via DOM) */
                applyAreaTextFrameAlignment(areaTextFrame, 1, false);
                createdFrames.push(areaTextFrame);
                pointText.remove();
            } catch (e) { }
        }
        return createdFrames;
    }

    // =========================================
    // 調整ダイアログ / Adjust dialog
    // =========================================

    /**
     * 幅・高さの入力を検証して有効値を返す（不正なら最終正常値へ戻す）
     * @param {EditText} editText - 対象の入力欄
     * @param {number|null} lastValue - 最終正常値（ルーラー単位）
     * @param {number} pointsPerUnit - ルーラー単位1あたりのポイント数
     * @returns {number|null} 有効値（ルーラー単位。不正なら null）
     */
    function validateSizeField(editText, lastValue, pointsPerUnit) {
        var sizeValue = parseFloat(String(editText.text));
        if (isNaN(sizeValue) || !isFinite(sizeValue) || sizeValue <= 0) {
            if (lastValue !== null) editText.text = lastValue;
            return null;
        }
        /* 表示単位での上限値（極端値防止）/ Max size in ruler units (guards extreme values) */
        var maxSize = 100000 / pointsPerUnit;
        if (sizeValue > maxSize) {
            sizeValue = maxSize;
            editText.text = Math.round(sizeValue * 100) / 100;
        }
        if (sizeValue < 0.01) {
            sizeValue = 0.01;
            editText.text = Math.round(sizeValue * 100) / 100;
        }
        return sizeValue;
    }

    /**
     * 選択からエリア内文字だけを集める
     * @param {Document} doc - 対象ドキュメント
     * @returns {TextFrame[]|null} エリア内文字（無ければ・読めなければ null）
     */
    function collectSelectedAreaTexts(doc) {
        /* undo 後などは選択が読めないことがある / The selection may be unreadable, e.g. after an undo */
        try {
            var currentSelection = doc.selection;
            var areaTextFrames = [];
            if (currentSelection && currentSelection.length) {
                for (var i = 0; i < currentSelection.length; i++) {
                    if (currentSelection[i] && isAreaTextFrame(currentSelection[i])) {
                        areaTextFrames.push(currentSelection[i]);
                    }
                }
            }
            return areaTextFrames.length ? areaTextFrames : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 調整ダイアログを組み立てる（イベントは showAdjustDialog() で結び付ける）
     * @param {string} unitLabel - ルーラー単位の表示
     * @returns {Object} ダイアログと各コントロール
     */
    function buildAdjustDialog(unitLabel) {
        var adjustDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        adjustDialog.alignChildren = "fill";
        adjustDialog.margins = DIALOG_MARGINS;

        var mainColumnGroup = adjustDialog.add("group");
        setupGroup(mainColumnGroup, "column");

        /* フォントサイズ / Font size */
        var fontSizePanel = mainColumnGroup.add("panel", undefined, getLabel("panel.fontSize"));
        setupPanel(fontSizePanel);
        var fontSizeRow = fontSizePanel.add("group");
        fontSizeRow.alignment = "left";
        fontSizeRow.add("statictext", undefined, labelText("fieldLabel.fontSize"));
        var fontSizeInput = addStepperInput(fontSizeRow, "");
        fontSizeInput.characters = 5;
        fontSizeInput.helpTip = getLabel("tooltip.fontSize");
        fontSizeRow.add("statictext", undefined, "pt");
        var oversetButtonRow = fontSizePanel.add("group");
        oversetButtonRow.orientation = "row";
        var btnFixOverset = oversetButtonRow.add("button", undefined, getLabel("button.overset"));
        btnFixOverset.helpTip = getLabel("tooltip.overset");

        /* フレームサイズ / Frame size */
        var frameSizePanel = mainColumnGroup.add("panel", undefined, getLabel("panel.frameSize"));
        setupPanel(frameSizePanel);
        var widthInput = addSizeRow(frameSizePanel, "width", unitLabel);
        var heightInput = addSizeRow(frameSizePanel, "height", unitLabel);

        /* インデント / Indent */
        var indentPanel = mainColumnGroup.add("panel", undefined, getLabel("panel.indent"));
        setupPanel(indentPanel);
        /* 連動のリンクアイコンを右側に並べるため行方向へ上書き / Override to row so the link icon sits to the right */
        indentPanel.orientation = "row";
        indentPanel.alignChildren = ["left", "top"];
        var indentFieldsColumn = indentPanel.add("group");
        indentFieldsColumn.orientation = "column";
        indentFieldsColumn.alignChildren = "left";
        var leftIndentRow = addIndentRow(indentFieldsColumn, "indentLeft", unitLabel);
        var rightIndentRow = addIndentRow(indentFieldsColumn, "indentRight", unitLabel);
        /* 連動のリンクアイコンは2つの欄の右、上下の中央に置く。切り替え後の処理は syncToggle.onToggleHandler に後から入れる
           The link icon sits right of the two fields, vertically centred; the handler is set later on syncToggle.onToggleHandler */
        var syncToggle = addLinkToggle(indentPanel, false, function () {
            if (syncToggle.onToggleHandler) { syncToggle.onToggleHandler(); }
        });
        syncToggle.alignment = ["left", "center"];
        syncToggle.helpTip = getLabel("tooltip.sync");

        /* オプション / Options */
        var optionsPanel = mainColumnGroup.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel);
        var marginRow = optionsPanel.add("group");
        var marginCheckbox = marginRow.add("checkbox", undefined, getLabel("checkbox.margin"));
        marginCheckbox.helpTip = getLabel("tooltip.margin");
        var marginInput = addStepperInput(marginRow, "0");
        marginInput.characters = 6;
        marginInput.helpTip = getLabel("tooltip.margin");
        var marginUnitLabel = marginRow.add("statictext", undefined, unitLabel);
        setStepperInputEnabled(marginInput, false);
        marginUnitLabel.enabled = false;

        /* ボタンエリア / Button area */
        var buttonRow = addButtonRow(adjustDialog);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("button.close"), { name: "cancel" });
        var btnRun = buttonRow.rightGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });

        return {
            adjustDialog: adjustDialog,
            fontSizeInput: fontSizeInput,
            btnFixOverset: btnFixOverset,
            widthInput: widthInput,
            heightInput: heightInput,
            leftIndentCheckbox: leftIndentRow.checkbox,
            leftIndentInput: leftIndentRow.input,
            rightIndentCheckbox: rightIndentRow.checkbox,
            rightIndentInput: rightIndentRow.input,
            syncToggle: syncToggle,
            marginCheckbox: marginCheckbox,
            marginInput: marginInput,
            marginUnitLabel: marginUnitLabel,
            btnClose: btnClose,
            btnRun: btnRun
        };
    }

    /**
     * エリア内文字の調整ダイアログを表示する（プレビューは常時ON）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame|null} initialFrame - 初期値を読むエリア内文字
     * @param {TextFrame[]|null} targetFrames - 調整対象（null なら現在の選択）
     * @returns {void}
     */
    function showAdjustDialog(doc, initialFrame, targetFrames) {
        var rulerUnit = getUnitInfo("rulerType");
        var pointsPerUnit = rulerUnit.pointsPerUnit;

        /* 受け取った変換結果を確実に対象にする / Make the passed frames the active target */
        var framesToSelect = (targetFrames && targetFrames.length) ? targetFrames : (initialFrame ? [initialFrame] : null);
        if (framesToSelect) {
            try { doc.selection = framesToSelect; } catch (e) { }
        }
        app.redraw();

        /* モーダル中は selection が変動するため、渡された配列を優先して固定 / Pin targets (selection drifts in modal) */
        var fixedTargets = null;
        if (targetFrames && targetFrames.length) {
            fixedTargets = targetFrames.slice(0);
        } else {
            /* 選択が読めないことがある / The selection may be unreadable */
            try {
                var initialSelection = doc.selection;
                if (initialSelection && initialSelection.length) {
                    fixedTargets = [];
                    for (var i = 0; i < initialSelection.length; i++) { fixedTargets.push(initialSelection[i]); }
                }
            } catch (e4) { }
        }

        var dialogControls = buildAdjustDialog(rulerUnit.label);
        var adjustDialog = dialogControls.adjustDialog;
        var fontSizeInput = dialogControls.fontSizeInput;
        var widthInput = dialogControls.widthInput;
        var heightInput = dialogControls.heightInput;
        var leftIndentCheckbox = dialogControls.leftIndentCheckbox;
        var leftIndentInput = dialogControls.leftIndentInput;
        var rightIndentCheckbox = dialogControls.rightIndentCheckbox;
        var rightIndentInput = dialogControls.rightIndentInput;
        var syncToggle = dialogControls.syncToggle;
        var marginCheckbox = dialogControls.marginCheckbox;
        var marginInput = dialogControls.marginInput;
        var marginUnitLabel = dialogControls.marginUnitLabel;

        /* 状態変数 / State */
        var isPreviewActive = false;

        /* 入力バリデーション用の最終正常値（ルーラー単位）/ Last valid values for validation (ruler units) */
        var lastValidWidth = null;
        var lastValidHeight = null;

        /**
         * ポイント値をルーラー単位に換算して小数第2位で丸める
         * @param {number} valueInPt - ポイント値
         * @returns {number} ルーラー単位の値
         */
        function toRoundedRulerUnits(valueInPt) {
            return Math.round((valueInPt / pointsPerUnit) * 100) / 100;
        }

        /**
         * 対象フレームから現在値をUIに読み込む
         * @param {TextFrame} sourceFrame - 読み取り元のエリア内文字
         * @returns {void}
         */
        function loadValuesFromFrame(sourceFrame) {
            var hasMultiParagraph;
            try { hasMultiParagraph = (sourceFrame.paragraphs && sourceFrame.paragraphs.length >= 2); } catch (e) { hasMultiParagraph = false; }
            var initialWidth = toRoundedRulerUnits(sourceFrame.textPath.width);
            var initialHeight = toRoundedRulerUnits(sourceFrame.textPath.height);
            var fontSize = 0;
            try { fontSize = sourceFrame.textRange.characterAttributes.size || 0; } catch (e) { }
            if (fontSize > 0) { fontSizeInput.text = Math.round(fontSize * 100) / 100; }
            widthInput.text = initialWidth;
            heightInput.text = initialHeight;
            lastValidWidth = parseFloat(widthInput.text);
            lastValidHeight = parseFloat(heightInput.text);
            try {
                var frameSpacing = sourceFrame.spacing || 0;
                marginInput.text = toRoundedRulerUnits(frameSpacing);
                marginCheckbox.value = (frameSpacing !== 0);
                setStepperInputEnabled(marginInput, marginCheckbox.value);
                marginUnitLabel.enabled = marginCheckbox.value;
            } catch (e) { }
            try {
                var leftIndentPt = sourceFrame.paragraphs.length > 0 ? (sourceFrame.paragraphs[0].leftIndent || 0) : 0;
                var rightIndentPt = sourceFrame.paragraphs.length > 0 ? (sourceFrame.paragraphs[0].rightIndent || 0) : 0;
                leftIndentCheckbox.value = (leftIndentPt !== 0);
                setStepperInputEnabled(leftIndentInput, leftIndentCheckbox.value);
                leftIndentInput.text = leftIndentCheckbox.value ? toRoundedRulerUnits(leftIndentPt) : "0";
                rightIndentCheckbox.value = (rightIndentPt !== 0);
                setStepperInputEnabled(rightIndentInput, rightIndentCheckbox.value);
                rightIndentInput.text = rightIndentCheckbox.value ? toRoundedRulerUnits(rightIndentPt) : "0";
            } catch (e) { }
            dialogControls.btnFixOverset.enabled = !hasMultiParagraph;
        }

        /**
         * UIの値をフレームへ適用する（行揃え・配置は常に中央）
         * @param {boolean} forPreview - プレビューとして適用するか（フレーム整列のアクションは実行しない）
         * @param {boolean} [shrinkToFit] - あふれないところまでフォントサイズを下げるか
         * @returns {void}
         */
        function runAdjust(forPreview, shrinkToFit) {
            var justification = Justification.CENTER;
            var frameAlignment = 1; /* 中央 / center */

            var leftIndentPt = (leftIndentCheckbox.value || syncToggle.value)
                ? (parseFloat(leftIndentInput.text) || 0) * pointsPerUnit : 0;
            var rightIndentPt = syncToggle.value
                ? leftIndentPt
                : (rightIndentCheckbox.value ? (parseFloat(rightIndentInput.text) || 0) * pointsPerUnit : 0);
            var marginPt = marginCheckbox.value ? (parseFloat(marginInput.text) || 0) * pointsPerUnit : 0;

            var savedSelection = [];
            var originalSelection = doc.selection;
            for (var j = 0; j < originalSelection.length; j++) { savedSelection.push(originalSelection[j]); }

            /* 固定ターゲットを優先 / Prefer pinned targets */
            if (!(fixedTargets && fixedTargets.length)) { fixedTargets = collectSelectedAreaTexts(doc); }
            var adjustTargets = fixedTargets && fixedTargets.length ? fixedTargets : doc.selection;

            for (var i = adjustTargets.length - 1; i >= 0; i--) {
                var targetFrame = adjustTargets[i];
                if (!isAreaTextFrame(targetFrame)) continue;
                try { targetFrame.spacing = marginPt; } catch (e) { }
                /* 幅/高さ：NaN・0以下・極端値をガード / Guard NaN, non-positive, extreme values */
                var widthInRulerUnits = validateSizeField(widthInput, lastValidWidth, pointsPerUnit);
                var heightInRulerUnits = validateSizeField(heightInput, lastValidHeight, pointsPerUnit);
                if (widthInRulerUnits !== null) { lastValidWidth = widthInRulerUnits; try { targetFrame.textPath.width = widthInRulerUnits * pointsPerUnit; } catch (e) { } }
                if (heightInRulerUnits !== null) { lastValidHeight = heightInRulerUnits; try { targetFrame.textPath.height = heightInRulerUnits * pointsPerUnit; } catch (e) { } }
                if (shrinkToFit) { shrinkFont(targetFrame); }
                try {
                    var paragraphAttrs = targetFrame.textRange.paragraphAttributes;
                    paragraphAttrs.justification = justification;
                    paragraphAttrs.leftIndent = leftIndentPt;
                    paragraphAttrs.rightIndent = rightIndentPt;
                } catch (e) { }
                applyAreaTextFrameAlignment(targetFrame, frameAlignment, forPreview);
            }
            if (savedSelection.length > 0) { try { doc.selection = savedSelection; } catch (e) { } }
            app.redraw();
        }

        /**
         * 直前のプレビューを undo で取り消す
         * @returns {void}
         */
        function undoPreview() {
            try { app.undo(); } catch (e) { }
            app.redraw();
            isPreviewActive = false;
        }

        /**
         * プレビューを更新する（常時ON。直前のプレビューは undo で戻す）
         * @returns {void}
         */
        function updatePreview() {
            if (isPreviewActive) {
                undoPreview();
                /* undo 後は参照が無効化されるため対象を取り直す / Refresh targets after undo */
                fixedTargets = collectSelectedAreaTexts(doc);
            }
            if (!(fixedTargets && fixedTargets.length)) { fixedTargets = collectSelectedAreaTexts(doc); }
            runAdjust(true);
            isPreviewActive = true;
        }

        /**
         * フォントサイズ入力を選択中のエリア内文字へ即時反映する
         * @returns {void}
         */
        function applyFontSizeFromField() {
            var newSize = parseFloat(fontSizeInput.text) || 0;
            if (newSize <= 0) return;
            var selectedItems = doc.selection;
            for (var i = 0; i < selectedItems.length; i++) {
                if (isAreaTextFrame(selectedItems[i])) {
                    try { selectedItems[i].textRange.characterAttributes.size = newSize; } catch (e) { }
                }
            }
            app.redraw();
        }

        /**
         * 幅入力を検証してプレビューする
         * @returns {void}
         */
        function onWidthChange() {
            var widthInRulerUnits = validateSizeField(widthInput, lastValidWidth, pointsPerUnit);
            if (widthInRulerUnits !== null) { lastValidWidth = widthInRulerUnits; }
            updatePreview();
        }

        /**
         * 高さ入力を検証してプレビューする
         * @returns {void}
         */
        function onHeightChange() {
            var heightInRulerUnits = validateSizeField(heightInput, lastValidHeight, pointsPerUnit);
            if (heightInRulerUnits !== null) { lastValidHeight = heightInRulerUnits; }
            updatePreview();
        }

        /**
         * 連動時に右インデントを同期してから幅変更扱いにする
         * @returns {void}
         */
        function onAdjustmentChange() {
            if (syncToggle.value) { rightIndentInput.text = leftIndentInput.text; }
            onWidthChange();
        }

        /* --- イベントハンドラ / Event handlers --- */
        dialogControls.btnFixOverset.onClick = function () { runAdjust(false, true); };
        marginCheckbox.onClick = function () {
            setStepperInputEnabled(marginInput, marginCheckbox.value);
            marginUnitLabel.enabled = marginCheckbox.value;
            if (marginCheckbox.value) { marginInput.text = "1"; }
            onAdjustmentChange();
        };
        syncToggle.onToggleHandler = function () {
            if (syncToggle.value) {
                leftIndentCheckbox.value = true; setStepperInputEnabled(leftIndentInput, true);
                rightIndentCheckbox.enabled = false; setStepperInputEnabled(rightIndentInput, false);
                rightIndentInput.text = leftIndentInput.text;
            } else {
                rightIndentCheckbox.enabled = true; setStepperInputEnabled(rightIndentInput, rightIndentCheckbox.value);
            }
            onAdjustmentChange();
        };
        leftIndentCheckbox.onClick = function () {
            if (!syncToggle.value) { setStepperInputEnabled(leftIndentInput, leftIndentCheckbox.value); }
            if (!leftIndentCheckbox.value) { leftIndentInput.text = "0"; }
            onAdjustmentChange();
        };
        rightIndentCheckbox.onClick = function () {
            setStepperInputEnabled(rightIndentInput, rightIndentCheckbox.value);
            if (!rightIndentCheckbox.value) { rightIndentInput.text = "0"; }
            onAdjustmentChange();
        };

        fontSizeInput.onChange = applyFontSizeFromField;
        marginInput.onChange = onAdjustmentChange;
        widthInput.onChange = onWidthChange;
        heightInput.onChange = onHeightChange;
        leftIndentInput.onChange = onAdjustmentChange;
        rightIndentInput.onChange = onAdjustmentChange;
        /* ∧∨・↑↓キーで増減したあとの処理 / what runs after stepping */
        fontSizeInput.stepOptions.onStep = applyFontSizeFromField;
        marginInput.stepOptions.onStep = onAdjustmentChange;
        widthInput.stepOptions.onStep = onWidthChange;
        heightInput.stepOptions.onStep = updatePreview;
        leftIndentInput.stepOptions.onStep = onAdjustmentChange;
        rightIndentInput.stepOptions.onStep = onAdjustmentChange;

        dialogControls.btnRun.onClick = function () {
            if (isPreviewActive) { undoPreview(); }
            runAdjust(false);
            adjustDialog.close(1);
        };
        dialogControls.btnClose.onClick = function () {
            if (isPreviewActive) { app.undo(); app.redraw(); }
            adjustDialog.close(0);
        };

        /* 初期値読み込み / Load initial values */
        if (initialFrame) { loadValuesFromFrame(initialFrame); }

        /* 開いたらプレビュー実行（常時ON）/ Run preview on open (always on) */
        updatePreview();

        prepareDialogWindow(adjustDialog, SCRIPT_NAME);
        adjustDialog.show();
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    /**
     * 選択を確かめ、変換または調整ダイアログを実行する（アクションは終了時に必ず破棄）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var selectionKinds = classifySelection(currentSelection);
        if (!selectionKinds.hasPointText && !selectionKinds.hasPathItem && !selectionKinds.hasAreaText) {
            alert(getLabel("alert.selectText"));
            return;
        }

        /* アクションを実行時に読み込み、終了時に破棄 / Load actions at start, unload on exit */
        loadAreaTextActions();
        try {
            if (selectionKinds.hasPointText || selectionKinds.hasPathItem) {
                /* ポイント文字 または 図形 → 変換（常にボタン風）してダイアログ / Point or shape → convert then dialog */
                convertToAreaTypeAndAdjust(doc, currentSelection);
            } else {
                /* エリア内文字のみ → 調整ダイアログ / Area type only → adjust dialog */
                showAdjustDialog(doc, findFirstAreaText(currentSelection), null);
            }
        } finally {
            unloadAreaTextActions();
        }
    }

    main();

})();
