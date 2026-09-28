#target illustrator
#targetengine "MultiAreaTextEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のテキストを1つのエリア内文字にまとめたり、逆に分割したりします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MultiAreaText.md

### Overview

Merges several text objects into a single area text, or splits one back out.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MultiAreaText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MultiAreaText";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-03-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MultiAreaText.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MultiAreaText.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];      /* パネルの余白 / panel margins */
    var SPACING_FIELD_CHARACTERS = 4;          /* 外枠からの間隔の入力欄の文字数 / width of the inset spacing field */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "複数のテキスト", en: "Multiple Text Objects" }
        },
        panel: {
            operation: { ja: "操作", en: "Action" },
            order: { ja: "順序", en: "Order" },
            threadText: { ja: "スレッドテキスト", en: "Threaded Text" },
            mergeSettings: { ja: "マージの設定", en: "Merge Settings" },
            areaTextHeight: { ja: "エリア内文字の高さ", en: "Area Text Height" }
        },
        radio: {
            merge: { ja: "マージ", en: "Merge" },
            thread: { ja: "スレッドテキスト", en: "Threaded Text" },
            swap: { ja: "交換（文字列のみ）", en: "Swap (Text Only)" },
            topToBottom: { ja: "上から", en: "Top to Bottom" },
            leftToRight: { ja: "左から", en: "Left to Right" },
            threadLink: { ja: "作成", en: "Create" },
            threadUnlink: { ja: "スレッドのリンクを解除", en: "Remove Threading" },
            threadAdd: { ja: "スレッドに追加", en: "Add to Thread" },
            threadRelease: { ja: "スレッドから除外", en: "Release from Thread" },
            releaseByDuplicate: { ja: "複製してスレッドから除外", en: "Duplicate and Release" },
            heightNone: { ja: "何もしない", en: "None" },
            heightFit: { ja: "フィット", en: "Fit" },
            heightAuto: { ja: "自動サイズ調整", en: "Auto Size" }
        },
        checkbox: {
            removeLineBreaks: { ja: "改行を削除", en: "Remove Line Breaks" },
            preserveFormatting: { ja: "書式を保持（段落単位）", en: "Preserve Formatting (Per Paragraph)" },
            insetSpacing: { ja: "外枠からの間隔", en: "Inset Spacing" },
            justify: { ja: "均等配置（最終行左揃え）", en: "Justify (Last Line Left)" },
            addBorder: { ja: "長方形の枠を付ける", en: "Add Rectangle Border" }
        },
        tooltip: {
            merge: {
                ja: "選択したテキストを連結し、全体を囲む1つのエリア内文字にまとめます",
                en: "Joins the selected text into one area text that covers them all."
            },
            thread: {
                ja: "［スレッドテキスト］パネルで選んだ操作を実行します",
                en: "Runs the action chosen in the Threaded Text panel."
            },
            swap: {
                ja: "2つのテキストの文字列を入れ替えます。位置と書式はそのままです",
                en: "Swaps the text of the two objects; their position and formatting stay put."
            },
            topToBottom: { ja: "上にあるテキストから順に連結します", en: "Joins the text from top to bottom." },
            leftToRight: { ja: "左にあるテキストから順に連結します", en: "Joins the text from left to right." },
            threadLink: {
                ja: "選択したエリア内文字をスレッドでつなぎます。すでにつながっているものがあれば、いったん解除してからつなぎ直します",
                en: "Threads the selected area text, removing any existing threading first."
            },
            threadUnlink: {
                ja: "［スレッドのリンクを解除］を実行します",
                en: "Runs Remove Threading."
            },
            threadAdd: {
                ja: "スレッドをいったん解除し、選択したエリア内文字をまとめてつなぎ直します",
                en: "Removes the threading, then threads the selected area text again."
            },
            threadRelease: {
                ja: "選択したエリア内文字をスレッドから外し、中の文字を同じ位置に独立したエリア内文字として残します（文字属性は段落ごと）",
                en: "Releases the selected area text and keeps its text as separate area text in place (character attributes per paragraph)."
            },
            releaseByDuplicate: {
                ja: "選択したエリア内文字を複製して独立させ、元のエリア内文字とその文字をスレッドから削除します",
                en: "Duplicates each selected area text as a standalone copy, then deletes the original and its text from the thread."
            },
            removeLineBreaks: {
                ja: "連結したテキストの改行をすべて削除します（［書式を保持］がオンのときは使えません）",
                en: "Removes every line break from the joined text (unavailable while Preserve Formatting is on)."
            },
            preserveFormatting: {
                ja: "段落ごとに文字属性を復元します（段落設定は完全ではありません）",
                en: "Restores character attributes paragraph by paragraph (paragraph settings are not fully kept)."
            },
            insetSpacing: {
                ja: "まとめたエリア内文字の［外枠からの間隔］を、定規の単位で設定します",
                en: "Sets the inset spacing of the merged area text, in ruler units."
            },
            addBorder: {
                ja: "新しい線を追加し、［形状に変換］の長方形効果でテキストを枠で囲みます",
                en: "Adds a new stroke and a Convert to Rectangle effect to frame the text."
            },
            heightFit: {
                ja: "テキストに合わせて高さを一度だけ調整します（自動サイズ調整はオフに戻します）",
                en: "Fits the height to the text once, then turns Auto Size off again."
            },
            heightAuto: {
                ja: "エリア内文字の自動サイズ調整をオンにします",
                en: "Turns on Auto Size for the area text."
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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            noTextFrame: { ja: "選択にテキストが含まれていません。", en: "Selection does not contain any text." },
            notThreaded: { ja: "選択されたテキストはスレッドテキストではありません。", en: "The selected text is not threaded text." },
            needTwo: { ja: "テキストを2つ以上選択してください。", en: "Please select 2 or more text objects." },
            releaseFailed: {
                ja: "スレッドから除外できなかったため、内容は変更していません。",
                en: "Could not release the text from the thread, so nothing was changed."
            }
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
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
    // 自動サイズ調整（一時アクション） / Auto size via a temporary action
    // =========================================

    /**
     * 一時アクションで、選択中のエリア内文字の自動サイズ調整を切り替える
     * @param {number} autoSizeValue - 1 = オン、2 = オフ
     * @returns {void}
     */
    function runAutoSizeAction(autoSizeValue) {
        var actionSource = '/version 3' + '/name [ 8' + ' 4172656154797065' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 8' + ' 4175746f53697a65' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOAreaTextDialog)' + ' /localizedName [ 33' + ' e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1952539754' + ' /showInPalette 4294967295' + ' /type (integer)' + ' /value ' + autoSizeValue + ' }' + ' }' + '}';

        /* 失敗したら従来どおり例外で処理を止める（解除と一時ファイルの削除は済んでいる）
           On failure, stop with an exception as before (the set and temp file are already cleaned up) */
        if (!runTemporaryAction(actionSource, "AreaType", "AutoSize")) {
            throw new Error("AutoSize action failed");
        }
    }

    /**
     * エリア内文字を選択し、自動サイズ調整をオン／オフする
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @param {boolean} autoSizeOn - オンにするなら true
     * @returns {void}
     */
    function setAreaTextAutoSize(textFrame, autoSizeOn) {
        app.activeDocument.selection = [textFrame];
        runAutoSizeAction(autoSizeOn ? 1 : 2);
    }

    // =========================================
    // 形状に変換（ライブ効果） / Convert to Shape live effect
    // =========================================

    /* ［形状に変換］の長方形：相対指定で幅・高さに 10pt 追加（CornerRadius は角丸長方形用で、長方形では使われない）
       Convert to Rectangle, relative: 10 pt extra width and height (CornerRadius only matters for rounded rectangles) */
    var RECTANGLE_SHAPE_EFFECT_XML = '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 ' +
        'R RelWidth 10 R RelHeight 10 R AbsWidth 10 R AbsHeight 10 R Absolute 0 R CornerRadius 9 "/></LiveEffect>';

    /**
     * ［形状に変換］の長方形効果を適用する。失敗したらメッセージを出す
     * @param {PageItem} targetItem - 適用先
     * @returns {void}
     */
    function applyRectangleShapeEffect(targetItem) {
        try {
            targetItem.applyEffect(RECTANGLE_SHAPE_EFFECT_XML);
        } catch (error) {
            alert(error.message);
        }
    }

    // =========================================
    // 選択と文字属性 / Selection and character attributes
    // =========================================

    /**
     * テキストフレームがスレッドでつながっているかを返す
     * @param {TextFrame} textFrame - 判定するフレーム
     * @returns {boolean} 前後どちらかにフレームがつながっていれば true
     */
    function isThreadedFrame(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return false;
        /* つながりが無いときや無効なフレームで例外になることがある / may throw when unthreaded or invalid */
        try { if (textFrame.nextFrame) return true; } catch (e) { }
        try { if (textFrame.previousFrame) return true; } catch (e) { }
        return false;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。使わない関数も消さずに残してよい（互いに呼び合う）。
    //    識別子は SELECTION_ITEMS_TOLERANCE / normalizeSelectionItems / resolveTextRangeFrame /
    //    collectSelectionItems / getTextFrameKindKey / collectSelectionTextFrames / collectSelectionPathItems /
    //    isClipMaskItem / getClipMaskItem / hasClippedDescendant / readUsePreviewBoundsPreference /
    //    getClipAwareBounds / filterMeasurableChildren / getClipAwareUnionBounds / isNearlySameCoordinate / areBoundsNearlyEqual
    // 2. 選択は normalizeSelectionItems(doc.selection) で配列にする。文字カーソルの選択（TextRange）は
    //    配列ではなく1個で返り、しかも .length（文字数）を持つので、length だけで配列と見なさない
    // 3. テキストフレーム:
    //      var frames = collectSelectionTextFrames(doc.selection);                           // 全種類
    //      var frames = collectSelectionTextFrames(doc.selection, { kinds: ["point", "path"] });
    //    パス:
    //      var paths = collectSelectionPathItems(doc.selection);                             // 複合パスは中のパスへ
    //      var paths = collectSelectionPathItems(doc.selection, { compoundPaths: "whole", skipClipMasks: true });
    //    それ以外は collectSelectionItems(source, { accept: function (item) { … } }) で条件を書く
    // 4. 並びは選択と同じ前面→背面（グループの中も pageItems の順）。重なり順を使う処理はこの順を前提にしてよい
    // 5. doc.selection に代入し直す配列は skipLocked / skipHidden を true にする。
    //    ロック・非表示を選択に代入すると例外になり、中の子が選択に残る
    // 6. 境界は getClipAwareBounds(item, usePreviewBounds) / getClipAwareUnionBounds(items, usePreviewBounds)。
    //    usePreviewBounds を省くと環境設定の［プレビュー境界を使用］に従う。返り値は [左, 上, 右, 下] の新しい配列
    //    （書き換えても元のオブジェクトに影響しない）。測れないときは null
    // 7. 座標の一致・前後の判定は isNearlySameCoordinate / areBoundsNearlyEqual で許容値を挟む
    //    （吸着させた辺とガイドは 1e-12 ほどずれる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 選択を処理対象の並びに展開する。グループは中のテキストフレームに置き換え、それ以外はそのまま残す
     * グループの中はロック・非表示のもの（中ごと）とクリッピングマスクを除く
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {Array} 展開した処理対象
     */
    function expandGroupsInSelection(selectedItems) {
        var expandedItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem") {
                expandedItems = expandedItems.concat(collectSelectionTextFrames(selectedItems[i].pageItems, {
                    skipLocked: true, skipHidden: true, skipClipMasks: true, textRangeToFrame: false, unique: false
                }));
            } else {
                expandedItems.push(selectedItems[i]);
            }
        }
        return expandedItems;
    }

    /**
     * 選択にテキストフレームが1つでも含まれるかを返す
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 含まれていれば true
     */
    function containsTextFrame(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame") return true;
        }
        return false;
    }

    /**
     * 選択にポイント文字またはパス上文字が含まれるかを返す
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 含まれていれば true
     */
    function containsNonAreaText(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame" &&
                (selectedItems[i].kind === TextType.POINTTEXT || selectedItems[i].kind === TextType.PATHTEXT)) {
                return true;
            }
        }
        return false;
    }

    /* 引き継ぐ文字属性（読み書きの順） / Character attributes carried over, in read/write order */
    var COPIED_CHARACTER_ATTRIBUTES = ["textFont", "size", "leading", "fillColor", "tracking", "kerningMethod"];

    /**
     * 文字属性のうち引き継ぐものを控える
     * @param {CharacterAttributes} sourceAttributes - 読み取り元の文字属性
     * @returns {Object} 属性名をキーにした控え
     */
    function readCharacterAttributes(sourceAttributes) {
        var attributeSnapshot = {};
        for (var attrIndex = 0; attrIndex < COPIED_CHARACTER_ATTRIBUTES.length; attrIndex++) {
            attributeSnapshot[COPIED_CHARACTER_ATTRIBUTES[attrIndex]] = sourceAttributes[COPIED_CHARACTER_ATTRIBUTES[attrIndex]];
        }
        return attributeSnapshot;
    }

    /**
     * 控えた文字属性を書き込む
     * @param {CharacterAttributes} targetAttributes - 書き込み先の文字属性
     * @param {Object} attributeSnapshot - readCharacterAttributes() の結果
     * @returns {void}
     */
    function writeCharacterAttributes(targetAttributes, attributeSnapshot) {
        for (var attrIndex = 0; attrIndex < COPIED_CHARACTER_ATTRIBUTES.length; attrIndex++) {
            targetAttributes[COPIED_CHARACTER_ATTRIBUTES[attrIndex]] = attributeSnapshot[COPIED_CHARACTER_ATTRIBUTES[attrIndex]];
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 縦並びの列グループを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加した列グループ
     */
    function addColumnGroup(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        return columnGroup;
    }

    /**
     * 共通レイアウトのパネルを追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} titlePath - パネル見出しのラベルパス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentGroup, titlePath) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        optionPanel.orientation = "column";
        optionPanel.alignment = ["fill", "top"];
        optionPanel.alignChildren = ["left", "center"];
        optionPanel.margins = PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * ツールチップを付けてコントロールを追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox"
     * @param {string} labelPath - ラベルのパス
     * @param {string} [tooltipPath] - ツールチップのラベルパス
     * @returns {Object} 追加したコントロール
     */
    function addControl(parentContainer, controlType, labelPath, tooltipPath) {
        var addedControl = parentContainer.add(controlType, undefined, getLabel(labelPath));
        if (tooltipPath) addedControl.helpTip = getLabel(tooltipPath);
        return addedControl;
    }

    /**
     * 左カラム（操作・順序・スレッドテキスト）を組み立てる
     * @param {Group} leftColumn - 追加先の列グループ
     * @param {Object} dialogControls - コントロールを登録するオブジェクト
     * @param {boolean} isSingle - テキストが1つだけなら true
     * @returns {void}
     */
    function addOperationColumn(leftColumn, dialogControls, isSingle) {
        dialogControls.operationPanel = addOptionPanel(leftColumn, "panel.operation");
        dialogControls.rbMerge = addControl(dialogControls.operationPanel, "radiobutton", "radio.merge", "tooltip.merge");
        dialogControls.rbThread = addControl(dialogControls.operationPanel, "radiobutton", "radio.thread", "tooltip.thread");
        dialogControls.rbSwap = addControl(dialogControls.operationPanel, "radiobutton", "radio.swap", "tooltip.swap");
        dialogControls.rbMerge.value = !isSingle;
        dialogControls.rbThread.value = isSingle;

        dialogControls.orderPanel = addOptionPanel(leftColumn, "panel.order");
        dialogControls.rbTopToBottom = addControl(dialogControls.orderPanel, "radiobutton", "radio.topToBottom", "tooltip.topToBottom");
        dialogControls.rbLeftToRight = addControl(dialogControls.orderPanel, "radiobutton", "radio.leftToRight", "tooltip.leftToRight");
        dialogControls.rbTopToBottom.value = true;

        dialogControls.threadPanel = addOptionPanel(leftColumn, "panel.threadText");
        dialogControls.rbThreadLink = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadLink", "tooltip.threadLink");
        dialogControls.rbThreadUnlink = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadUnlink", "tooltip.threadUnlink");
        dialogControls.rbThreadAdd = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadAdd", "tooltip.threadAdd");
        dialogControls.rbThreadRelease = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadRelease", "tooltip.threadRelease");
        dialogControls.rbReleaseByDuplicate = addControl(dialogControls.threadPanel, "radiobutton", "radio.releaseByDuplicate", "tooltip.releaseByDuplicate");
        dialogControls.rbThreadRelease.value = isSingle;
        dialogControls.rbThreadLink.value = !isSingle;
    }

    /**
     * 右カラム（マージの設定・エリア内文字の高さ）を組み立てる
     * @param {Group} rightColumn - 追加先の列グループ
     * @param {Object} dialogControls - コントロールを登録するオブジェクト
     * @param {string} rulerLabel - 定規の単位の表示名
     * @returns {void}
     */
    function addMergeSettingsColumn(rightColumn, dialogControls, rulerLabel) {
        var settingsPanel = addOptionPanel(rightColumn, "panel.mergeSettings");
        dialogControls.mergeSettingsPanel = settingsPanel;
        dialogControls.cbRemoveLineBreaks = addControl(settingsPanel, "checkbox", "checkbox.removeLineBreaks", "tooltip.removeLineBreaks");
        dialogControls.cbPreserveFormatting = addControl(settingsPanel, "checkbox", "checkbox.preserveFormatting", "tooltip.preserveFormatting");
        dialogControls.cbPreserveFormatting.value = true;
        dialogControls.cbRemoveLineBreaks.enabled = false;
        dialogControls.cbPreserveFormatting.onClick = function () {
            dialogControls.cbRemoveLineBreaks.enabled = !dialogControls.cbPreserveFormatting.value;
        };

        var insetSpacingGroup = settingsPanel.add("group");
        dialogControls.cbInsetSpacing = addControl(insetSpacingGroup, "checkbox", "checkbox.insetSpacing", "tooltip.insetSpacing");
        /* ∧∨と入力欄は隙間0で突き合わせる。間隔は0未満にしない / butt the stepper against the field; never below 0 */
        var insetStepperGroup = insetSpacingGroup.add("group");
        insetStepperGroup.orientation = "row";
        insetStepperGroup.alignChildren = ["left", "center"];
        insetStepperGroup.spacing = 0;
        insetStepperGroup.margins = 0;
        var insetStepper = addStepper(insetStepperGroup, function () { return dialogControls.insetSpacingInput; }, { min: 0 });
        dialogControls.insetSpacingInput = insetStepperGroup.add("edittext", undefined, "1");
        dialogControls.insetSpacingInput.characters = SPACING_FIELD_CHARACTERS;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(dialogControls.insetSpacingInput, insetStepper);
        insetSpacingGroup.add("statictext", undefined, rulerLabel);

        /**
         * 間隔の入力欄を、∧∨ごと有効／無効にする
         * @param {boolean} isEnabled - 有効にするなら true
         * @returns {void}
         */
        function setInsetSpacingEnabled(isEnabled) {
            dialogControls.insetSpacingInput.enabled = isEnabled;
            insetStepper.enabled = isEnabled;
            redrawSteppersIn(insetStepper);
        }
        setInsetSpacingEnabled(false);
        dialogControls.cbInsetSpacing.onClick = function () {
            setInsetSpacingEnabled(dialogControls.cbInsetSpacing.value);
        };

        dialogControls.cbJustify = addControl(settingsPanel, "checkbox", "checkbox.justify");
        dialogControls.cbJustify.value = true;
        dialogControls.cbAddBorder = addControl(settingsPanel, "checkbox", "checkbox.addBorder", "tooltip.addBorder");

        dialogControls.heightPanel = addOptionPanel(rightColumn, "panel.areaTextHeight");
        dialogControls.rbHeightNone = addControl(dialogControls.heightPanel, "radiobutton", "radio.heightNone");
        dialogControls.rbHeightFit = addControl(dialogControls.heightPanel, "radiobutton", "radio.heightFit", "tooltip.heightFit");
        dialogControls.rbHeightAuto = addControl(dialogControls.heightPanel, "radiobutton", "radio.heightAuto", "tooltip.heightAuto");
        dialogControls.rbHeightFit.value = true;
    }

    /**
     * 選んだ操作に合わせて、パネルとラジオボタンの有効／無効を切り替える
     * 子の有効／無効は各チェックボックスの onClick が保つので、ここではパネル単位で切り替える
     * @param {Object} dialogControls - buildDialog() のコントロール
     * @param {Array} selectedItems - 処理対象
     * @returns {void}
     */
    function updatePanelStates(dialogControls, selectedItems) {
        var isMerge = dialogControls.rbMerge.value;
        /* 1つのときはマージ不可、交換は2つのときだけ / Merge needs two or more, swap exactly two */
        dialogControls.rbMerge.enabled = (selectedItems.length > 1);
        dialogControls.rbSwap.enabled = (selectedItems.length === 2);
        /* ポイント文字・パス上文字はスレッドにできない / Point and path text cannot be threaded */
        dialogControls.rbThread.enabled = !containsNonAreaText(selectedItems);
        dialogControls.orderPanel.enabled = isMerge;
        dialogControls.threadPanel.enabled = dialogControls.rbThread.value;
        dialogControls.mergeSettingsPanel.enabled = isMerge;
        redrawSteppersIn(dialogControls.mergeSettingsPanel);
        dialogControls.heightPanel.enabled = isMerge;
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
     * ダイアログを組み立てる（表示はしない）
     * @param {Array} selectedItems - 処理対象
     * @param {string} rulerLabel - 定規の単位の表示名
     * @param {boolean} heightOnly - ［エリア内文字の高さ］だけを有効にするなら true
     * @returns {Object} ダイアログ本体（window）と各コントロール
     */
    function buildDialog(selectedItems, rulerLabel, heightOnly) {
        var dialogControls = {};

        dialogControls.window = new Window("dialog", getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        var columnsGroup = dialogControls.window.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];

        addOperationColumn(addColumnGroup(columnsGroup), dialogControls, selectedItems.length === 1);
        addMergeSettingsColumn(addColumnGroup(columnsGroup), dialogControls, rulerLabel);

        if (heightOnly) {
            /* スレッドでないエリア内文字1つ：エリア内文字の高さ以外を無効化 / A single unthreaded area text: disable everything but Area Text Height */
            dialogControls.operationPanel.enabled = false;
            dialogControls.orderPanel.enabled = false;
            dialogControls.threadPanel.enabled = false;
            dialogControls.mergeSettingsPanel.enabled = false;
            redrawSteppersIn(dialogControls.mergeSettingsPanel);
        } else {
            var onOperationClick = function () { updatePanelStates(dialogControls, selectedItems); };
            dialogControls.rbMerge.onClick = onOperationClick;
            dialogControls.rbThread.onClick = onOperationClick;
            dialogControls.rbSwap.onClick = onOperationClick;
            onOperationClick();
        }

        var buttonRow = addButtonRow(dialogControls.window);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return dialogControls;
    }

    /**
     * ダイアログのコントロールから処理設定を読み取る
     * @param {Object} dialogControls - buildDialog() の結果
     * @returns {Object} 処理設定
     */
    function readDialogSettings(dialogControls) {
        var threadAction = "link";
        if (dialogControls.rbThreadUnlink.value) threadAction = "unlink";
        else if (dialogControls.rbThreadRelease.value) threadAction = "release";
        else if (dialogControls.rbReleaseByDuplicate.value) threadAction = "releaseByDuplicate";
        else if (dialogControls.rbThreadAdd.value) threadAction = "add";

        var frameHeightMode = "none";
        if (dialogControls.rbHeightFit.value) frameHeightMode = "fit";
        else if (dialogControls.rbHeightAuto.value) frameHeightMode = "auto";

        return {
            operation: dialogControls.rbThread.value ? "thread" : (dialogControls.rbSwap.value ? "swap" : "merge"),
            threadAction: threadAction,
            leftToRight: dialogControls.rbLeftToRight.value,
            removeLineBreaks: dialogControls.cbRemoveLineBreaks.value,
            preserveFormatting: dialogControls.cbPreserveFormatting.value,
            insetSpacingOn: dialogControls.cbInsetSpacing.value,
            insetSpacingText: dialogControls.insetSpacingInput.text,
            justify: dialogControls.cbJustify.value,
            addBorder: dialogControls.cbAddBorder.value,
            frameHeightMode: frameHeightMode
        };
    }

    // =========================================
    // スレッドテキスト / Threaded text
    // =========================================

    /**
     * 選択したフレームをスレッドでつなぐ（つながっているものがあれば先に解除する）
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function linkThreadFrames(selectedItems) {
        /* スレッドテキストが混在しているか判定 / Check if selection has threaded frames */
        var hasThreaded = false;
        for (var i = 0; i < selectedItems.length; i++) {
            if (isThreadedFrame(selectedItems[i])) {
                hasThreaded = true;
                break;
            }
        }
        if (hasThreaded) {
            app.executeMenuCommand('removeThreading');
        }
        app.executeMenuCommand('threadTextCreate');
    }

    /**
     * スレッドから外す前に、フレームの位置・文字列・段落ごとの文字属性を控える
     * @param {TextFrame} textFrame - 対象のフレーム
     * @returns {Object} 控え（frame / left / top / width / height / text / paragraphAttributes）
     */
    function snapshotFrameForRelease(textFrame) {
        var frameBounds = textFrame.geometricBounds;
        var frameText = textFrame.contents;

        var paragraphAttributes = [];
        for (var paragraphIndex = 0; paragraphIndex < textFrame.paragraphs.length; paragraphIndex++) {
            /* 読めない段落（末尾の空段落など）で打ち切る / stop at a paragraph that cannot be read, such as a trailing empty one */
            try {
                paragraphAttributes.push(readCharacterAttributes(textFrame.paragraphs[paragraphIndex].characterAttributes));
            } catch (e) {
                break;
            }
        }

        return {
            frame: textFrame,
            left: frameBounds[0],
            top: frameBounds[1],
            width: frameBounds[2] - frameBounds[0],
            height: frameBounds[1] - frameBounds[3],
            text: frameText,
            paragraphAttributes: paragraphAttributes
        };
    }

    /**
     * 控えた内容から、同じ位置・大きさのエリア内文字を作り直す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} frameSnapshot - snapshotFrameForRelease() の結果
     * @returns {void}
     */
    function recreateAreaText(doc, frameSnapshot) {
        var newRect = doc.pathItems.rectangle(frameSnapshot.top, frameSnapshot.left, frameSnapshot.width, frameSnapshot.height);
        var newTextFrame = doc.textFrames.areaText(newRect);
        newTextFrame.contents = frameSnapshot.text;

        var paragraphAttributes = frameSnapshot.paragraphAttributes;
        /* まず先頭段落の書式をtextRange全体に適用 / Apply first para attrs to entire textRange */
        if (paragraphAttributes.length > 0) {
            writeCharacterAttributes(newTextFrame.textRange.characterAttributes, paragraphAttributes[0]);
        }

        /* 段落ごとに書式を適用。段落が足りなければ打ち切る / Apply attributes per paragraph; stop when paragraphs run out */
        for (var paragraphIndex = 1; paragraphIndex < paragraphAttributes.length; paragraphIndex++) {
            try {
                writeCharacterAttributes(newTextFrame.paragraphs[paragraphIndex].characterAttributes, paragraphAttributes[paragraphIndex]);
            } catch (e) {
                break;
            }
        }
    }

    /**
     * 選択を控える（あとで元に戻すため）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択されていたオブジェクト
     */
    function copySelection(doc) {
        var savedSelection = [];
        for (var i = 0; i < doc.selection.length; i++) savedSelection.push(doc.selection[i]);
        return savedSelection;
    }

    /**
     * 控えた選択を戻す。削除済みのオブジェクトは飛ばす
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} savedSelection - copySelection() の結果
     * @returns {void}
     */
    function restoreSelection(doc, savedSelection) {
        var itemsToSelect = [];
        for (var i = 0; i < savedSelection.length; i++) {
            /* 削除済みのオブジェクトは typename の参照で例外になる / a removed item throws on typename */
            try {
                if (savedSelection[i] && savedSelection[i].typename) itemsToSelect.push(savedSelection[i]);
            } catch (e) { }
        }
        if (itemsToSelect.length === 0) return;
        try {
            doc.selection = itemsToSelect;
        } catch (e) { }
    }

    /**
     * 選択したフレームをスレッドから外し、中の文字を同じ位置の独立したエリア内文字として残す
     * 先に contents を消すと、コマンド失敗時にテキストが失われるため、release の成功を確かめてから消す
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function releaseFromThread(doc, selectedItems) {
        var frameSnapshots = [];
        var releaseTargets = [];

        /* 選択を控えて、最後に戻す / Save the selection to restore it at the end */
        var savedSelection = copySelection(doc);

        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename !== "TextFrame") continue;
            frameSnapshots.push(snapshotFrameForRelease(selectedItems[i]));
            releaseTargets.push(selectedItems[i]);
        }

        /* 先に release を実行（まだ内容は変えない） / Execute release first, before any destructive change */
        var releaseOk = false;
        try {
            doc.selection = releaseTargets;
            app.executeMenuCommand('releaseThreadedTextSelection');

            /* すべての対象がスレッドから外れていれば成功 / Success when no target is threaded any more */
            releaseOk = true;
            for (var checkIndex = 0; checkIndex < frameSnapshots.length; checkIndex++) {
                if (isThreadedFrame(frameSnapshots[checkIndex].frame)) {
                    releaseOk = false;
                    break;
                }
            }
        } catch (e) {
            releaseOk = false;
        }

        if (!releaseOk) {
            alert(getLabel("alert.releaseFailed"));
            return;
        }

        /* スレッドから外れたので、元の内容を消しても連結先に影響しない / Safe to clear now that the chain is broken */
        for (var clearIndex = 0; clearIndex < frameSnapshots.length; clearIndex++) {
            frameSnapshots[clearIndex].frame.contents = "";
        }

        /* 保存した情報から新しいエリア内文字を作成 / Create new area text from saved info */
        for (var createIndex = 0; createIndex < frameSnapshots.length; createIndex++) {
            recreateAreaText(doc, frameSnapshots[createIndex]);
        }

        /* 空になり、スレッドから外れた元のフレームだけを削除 / Remove only the originals that are empty and unthreaded */
        for (var removeIndex = 0; removeIndex < frameSnapshots.length; removeIndex++) {
            /* 無効になったフレームは読み取りや削除で例外になる / an invalid frame throws on read or remove */
            try {
                var originalFrame = frameSnapshots[removeIndex].frame;
                if (!originalFrame || originalFrame.typename !== "TextFrame") continue;
                if (originalFrame.contents !== "") continue;
                if (isThreadedFrame(originalFrame)) continue;
                originalFrame.remove();
            } catch (e) { }
        }

        restoreSelection(doc, savedSelection);
    }

    /**
     * 選択したフレームを複製して独立させ、元のフレームを削除する
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function releaseByDuplicate(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename !== "TextFrame") continue;
            var originalFrame = selectedItems[i];
            /* 複製（スレッドに属さない独立コピーが作られる） / The duplicate is a standalone copy */
            var duplicatedFrame = originalFrame.duplicate();
            /* 複製がスレッドから独立しているか簡易検証 / Make sure the duplicate is not threaded */
            if (isThreadedFrame(duplicatedFrame)) {
                duplicatedFrame.remove();
                continue;
            }
            /* 元フレームのテキストをスレッドから削除し、元フレームも削除 / Clear the original text from the thread, then remove the frame */
            originalFrame.contents = "";
            originalFrame.remove();
        }
    }

    /**
     * ［スレッドテキスト］パネルで選んだ操作を実行する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {string} threadAction - "link" / "unlink" / "add" / "release" / "releaseByDuplicate"
     * @returns {void}
     */
    function runThreadAction(doc, selectedItems, threadAction) {
        /* メニューコマンドは選択に効くので、グループを展開した並びを選び直す / Menu commands act on the selection, so reselect the expanded items */
        doc.selection = selectedItems;
        if (threadAction === "unlink") {
            app.executeMenuCommand('removeThreading');
        } else if (threadAction === "release") {
            releaseFromThread(doc, selectedItems);
        } else if (threadAction === "releaseByDuplicate") {
            releaseByDuplicate(selectedItems);
        } else if (threadAction === "add") {
            app.executeMenuCommand('removeThreading');
            app.executeMenuCommand('threadTextCreate');
        } else {
            linkThreadFrames(selectedItems);
        }
    }

    // =========================================
    // マージ / Merge
    // =========================================

    /**
     * マージの対象になるテキストフレームを集め、指定の順に並べる
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {boolean} leftToRight - 左から並べるなら true（false なら上から）
     * @returns {TextFrame[]} 並べ替えたテキストフレーム
     */
    function collectMergeTargets(selectedItems, leftToRight) {
        var textFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame") textFrames.push(selectedItems[i]);
        }

        /* ソート（左から右、または上から下） / Sort (left-to-right or top-to-bottom) */
        if (leftToRight) {
            textFrames.sort(function (firstFrame, secondFrame) {
                return firstFrame.geometricBounds[0] - secondFrame.geometricBounds[0];
            });
        } else {
            textFrames.sort(function (firstFrame, secondFrame) {
                return secondFrame.geometricBounds[1] - firstFrame.geometricBounds[1];
            });
        }
        return textFrames;
    }

    /**
     * 複数のフレーム全体を囲む矩形を求める
     * @param {TextFrame[]} textFrames - 対象のフレーム
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getCombinedBounds(textFrames) {
        var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        for (var i = 0; i < textFrames.length; i++) {
            var frameBounds = textFrames[i].geometricBounds;
            minX = Math.min(minX, frameBounds[0]);
            maxY = Math.max(maxY, frameBounds[1]);
            maxX = Math.max(maxX, frameBounds[2]);
            minY = Math.min(minY, frameBounds[3]);
        }
        return [minX, maxY, maxX, minY];
    }

    /**
     * まとめたエリア内文字に、元のフレームごとの文字属性を段落単位で書き戻す
     * @param {TextFrame} newTextFrame - まとめたエリア内文字
     * @param {Object[]} frameAttributes - フレームごとの文字属性の控え
     * @param {number[]} frameParagraphCounts - フレームごとの段落数
     * @returns {void}
     */
    function applyFormattingPerFrame(newTextFrame, frameAttributes, frameParagraphCounts) {
        /* 先頭フレームの書式をtextRange全体に適用後、段落ごとに上書き / Apply the first frame to the whole range, then override per paragraph */
        writeCharacterAttributes(newTextFrame.textRange.characterAttributes, frameAttributes[0]);
        var paragraphIndex = 0;
        for (var frameIndex = 0; frameIndex < frameAttributes.length; frameIndex++) {
            for (var framePara = 0; framePara < frameParagraphCounts[frameIndex]; framePara++) {
                /* 末尾の空段落は Error 1302 になるのでスキップ / skip the trailing empty paragraph (Error 1302) */
                try {
                    writeCharacterAttributes(newTextFrame.paragraphs[paragraphIndex].characterAttributes, frameAttributes[frameIndex]);
                } catch (e) { }
                paragraphIndex++;
            }
        }
    }

    /**
     * 行揃えと外枠からの間隔を適用する
     * @param {TextFrame} textFrame - 対象のフレーム
     * @param {Object} mergeSettings - readDialogSettings() の結果
     * @param {number} rulerToPoint - 定規の単位 1 あたりのポイント数
     * @returns {void}
     */
    function applyStyle(textFrame, mergeSettings, rulerToPoint) {
        if (mergeSettings.justify) {
            textFrame.textRange.paragraphAttributes.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        }
        if (mergeSettings.insetSpacingOn) {
            var spacingValue = parseFloat(mergeSettings.insetSpacingText);
            if (!isNaN(spacingValue)) textFrame.spacing = spacingValue * rulerToPoint;
        }
    }

    /**
     * ［エリア内文字の高さ］の設定を適用する
     * @param {TextFrame} textFrame - 対象のフレーム
     * @param {string} frameHeightMode - "none" / "fit" / "auto"
     * @returns {void}
     */
    function applyFrameHeight(textFrame, frameHeightMode) {
        if (frameHeightMode === "none") return;
        if (textFrame.kind !== TextType.AREATEXT) return;
        setAreaTextAutoSize(textFrame, true);
        if (frameHeightMode === "fit") {
            /* 一度だけ合わせて、自動サイズ調整はオフに戻す / Fit once, then turn Auto Size back off */
            setAreaTextAutoSize(textFrame, false);
        }
    }

    /**
     * 選択したテキストを連結し、全体を囲む1つのエリア内文字にまとめる
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {Object} mergeSettings - readDialogSettings() の結果
     * @param {number} rulerToPoint - 定規の単位 1 あたりのポイント数
     * @returns {void}
     */
    function mergeTextFrames(doc, selectedItems, mergeSettings, rulerToPoint) {
        var textFrames = collectMergeTargets(selectedItems, mergeSettings.leftToRight);
        if (textFrames.length < 2) {
            alert(getLabel("alert.needTwo"));
            return;
        }

        var combinedText = [];
        for (var i = 0; i < textFrames.length; i++) {
            combinedText.push(textFrames[i].contents);
        }

        /* 書式保持用：フレームごとの属性と段落数を収集 / Collect frame-level attributes and paragraph counts */
        var frameAttributes = [];
        var frameParagraphCounts = [];
        if (mergeSettings.preserveFormatting) {
            for (var frameIndex = 0; frameIndex < textFrames.length; frameIndex++) {
                frameAttributes.push(readCharacterAttributes(textFrames[frameIndex].textRange.characterAttributes));
                frameParagraphCounts.push(textFrames[frameIndex].paragraphs.length);
            }
        }

        var combinedBounds = getCombinedBounds(textFrames);
        var areaRect = doc.pathItems.rectangle(combinedBounds[1], combinedBounds[0],
            combinedBounds[2] - combinedBounds[0], combinedBounds[1] - combinedBounds[3]);
        var newTextFrame = doc.textFrames.areaText(areaRect);
        var joinedText = combinedText.join("\r");
        if (mergeSettings.removeLineBreaks) {
            joinedText = joinedText.replace(/[\r\n]/g, '');
        }
        newTextFrame.contents = joinedText;

        /* 書式の適用 / Apply character attributes */
        if (mergeSettings.preserveFormatting && !mergeSettings.removeLineBreaks) {
            applyFormattingPerFrame(newTextFrame, frameAttributes, frameParagraphCounts);
        } else {
            /* 先頭フレームの書式を一括適用 / Apply first frame attributes uniformly */
            writeCharacterAttributes(newTextFrame.textRange.characterAttributes,
                readCharacterAttributes(textFrames[0].textRange.characterAttributes));
        }

        for (var removeIndex = textFrames.length - 1; removeIndex >= 0; removeIndex--) {
            textFrames[removeIndex].remove();
        }

        /* スタイルを適用 / Apply style */
        applyStyle(newTextFrame, mergeSettings, rulerToPoint);

        /* 新しいエリア内文字を選択 / Select new area text frame */
        newTextFrame.selected = true;

        /* エリア内文字の高さを適用 / Apply area text height */
        applyFrameHeight(newTextFrame, mergeSettings.frameHeightMode);

        /* 長方形の枠を付ける：新規線を追加し、長方形効果で囲む / Add a border: a new stroke plus a rectangle effect */
        if (mergeSettings.addBorder) {
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            applyRectangleShapeEffect(newTextFrame);
        }
        app.redraw();
    }

    // =========================================
    // 交換 / Swap
    // =========================================

    /**
     * 2つのテキストの文字列を入れ替える（位置と書式はそのまま）
     * @param {TextFrame} firstFrame - 1つ目のテキスト
     * @param {TextFrame} secondFrame - 2つ目のテキスト
     * @returns {void}
     */
    function swapTextContents(firstFrame, secondFrame) {
        var firstContents = firstFrame.contents;
        firstFrame.contents = secondFrame.contents;
        secondFrame.contents = firstContents;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを表示し、選んだ操作（マージ／スレッド／交換）を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        if (!doc.selection || doc.selection.length < 1) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* グループは中のテキストフレームに展開 / Expand groups into the text frames inside */
        var selectedItems = expandGroupsInSelection(doc.selection);

        /* テキストフレームが含まれているか検証 / Verify selection contains text frames */
        if (!containsTextFrame(selectedItems)) {
            alert(getLabel("alert.noTextFrame"));
            return;
        }

        /* 1つ選択時：スレッドでなければエリア内文字の高さだけ。ポイント文字などは対象外 / Single selection: unthreaded area text gets Area Text Height only; other kinds exit */
        var heightOnly = false;
        if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame" && !isThreadedFrame(selectedItems[0])) {
            if (selectedItems[0].kind !== TextType.AREATEXT) {
                alert(getLabel("alert.notThreaded"));
                return;
            }
            heightOnly = true;
        }

        /* ルーラー単位 / Ruler units */
        var rulerUnit = getUnitInfo("rulerType");

        var dialogControls = buildDialog(selectedItems, rulerUnit.label, heightOnly);
        prepareDialogWindow(dialogControls.window, SCRIPT_NAME);
        if (dialogControls.window.show() !== 1) return;
        var dialogSettings = readDialogSettings(dialogControls);

        /* エリア内文字の高さだけ / Area Text Height only */
        if (heightOnly) {
            applyFrameHeight(selectedItems[0], dialogSettings.frameHeightMode);
            app.redraw();
            return;
        }

        /* スレッドテキスト / Thread text */
        if (dialogSettings.operation === "thread") {
            runThreadAction(doc, selectedItems, dialogSettings.threadAction);
            return;
        }

        /* 交換 / Swap */
        if (dialogSettings.operation === "swap") {
            swapTextContents(selectedItems[0], selectedItems[1]);
            return;
        }

        /* マージ / Merge */
        mergeTextFrames(doc, selectedItems, dialogSettings, rulerUnit.pointsPerUnit);
    }

    main();

})();
