#target illustrator
#targetengine "AiQuickPrefsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorの各種環境設定を、常駐パレットでまとめて切り替えます。設定は操作した時点で即時反映されます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManagerForTransformAndAlignPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n41d8dc1961be

### Overview

A persistent palette for switching a range of Illustrator preferences. Every change takes effect the moment you make it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManagerForTransformAndAlignPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PreferenceManagerForTransformAndAlignPalette"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManagerForTransformAndAlignPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManagerForTransformAndAlignPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n41d8dc1961be"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DIALOG_OPACITY = 0.98;   /* パレットの不透明度 / Palette opacity */
    var SAVE_DEBOUNCE_MS = 40;   /* 保存デバウンス(ms) / Save debounce (ms) */

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
            title: { ja: "環境設定：変形と整列", en: "Preferences: Transform & Align" }
        },
        panel: {
            keyInput: { ja: "キー増加", en: "Key Input" },
            transform: { ja: "変形と整列", en: "Transform & Align" },
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            guide: { ja: "ガイドと定規", en: "Guides & Rulers" },
            artboard: { ja: "アートボード名と枠線", en: "Artboard Name & Border" },
            artboardBorder: { ja: "アートボードの枠線", en: "Artboard Border" },
            etc: { ja: "その他", en: "Other" }
        },
        tooltip: {
            keyValue:      { ja: "環境設定［一般］の「キー入力」の値です。矢印キー1回で動く距離になります。", en: "The Keyboard Increment from the General preferences: how far one arrow key press moves things." },
            keyUnit:       { ja: "「キー入力」の値を入力する単位です。", en: "The unit the Keyboard Increment is entered in." },
            previewBounds: { ja: "線幅や効果を含めた見た目の端を、オブジェクトの境界として扱います。", en: "Treats the visible edges including strokes and effects as the object bounds." },
            transformPattern: { ja: "オブジェクトを変形したとき、パターン塗りも一緒に変形します。", en: "Transforms pattern fills along with the object." },
            scaleCorners:  { ja: "拡大・縮小したとき、ライブコーナーの角丸も一緒に変わります。", en: "Scales live corner radii along with the object." },
            scaleStroke:   { ja: "拡大・縮小したとき、線幅と効果も一緒に変わります。", en: "Scales stroke weights and effects along with the object." },
            glyphBounds:   { ja: "整列の基準を、仮想ボディではなく字形の実際の輪郭にします。", en: "Aligns text by the actual glyph outlines instead of the em box." },
            guideShow:     { ja: "ガイドの表示・非表示を切り替えます。", en: "Shows or hides the guides." },
            guideLock:     { ja: "ガイドをロックして、選択・移動できないようにします。", en: "Locks the guides so they cannot be selected or moved." },
            showArtboardName: { ja: "カンバス上にアートボード名を表示します。", en: "Shows the artboard names on the canvas." },
            canvasWhite:   { ja: "アートボードの外の背景を白にします。", en: "Makes the canvas outside the artboards white." },
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
        checkbox: {
            pointText: { ja: "ポイント文字", en: "Point Type" },
            areaText: { ja: "エリア内文字", en: "Area Type" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
            transformPattern: { ja: "パターンを変形", en: "Transform Pattern Tiles" },
            scaleCorners: { ja: "角を拡大・縮小", en: "Scale Corners" },
            scaleStroke: { ja: "線幅と効果も拡大・縮小", en: "Scale Strokes & Effects" },
            guideShow: { ja: "ガイドを表示", en: "Show Guides" },
            guideLock: { ja: "ガイドをロック", en: "Lock Guides" },
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            canvasWhite: { ja: "カンバスカラーをホワイトに", en: "Set Canvas Color to White" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-time Drawing & Editing" },
            pastePlain: { ja: "書式なしペースト", en: "Paste without Formatting" }
        },
        label: {
            strokeColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            strokeWidth: { ja: "ストロークの幅", en: "Stroke Width" }
        },
        button: {
            videoRuler: { ja: "ビデオ定規", en: "Video Ruler" }
        },
        color: {
            lightBlue: { ja: "ライトブルー", en: "Light Blue" },
            red: { ja: "サーモンピンク", en: "Light Red" },
            green: { ja: "グリーン", en: "Green" },
            blue: { ja: "ミディアムブルー", en: "Medium Blue" },
            magenta: { ja: "マゼンタ", en: "Magenta" },
            cyan: { ja: "シアン", en: "Cyan" },
            grey: { ja: "ライトグレー", en: "Light Gray" },
            black: { ja: "ブラック", en: "Black" },
            yellow: { ja: "イエロー", en: "Yellow" }
        }
    };

    // =========================================
    // 単位 / Unit
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）
       decimals：1pt 未満に潰れないように、大きい単位ほど桁数を増やす（in で 1mm ≒ 0.039）
       popup：単位ポップアップに並べるかどうか
       Unit table; the array index equals the rulerType code.
       decimals: larger units need more digits so small values do not collapse to 0 (1mm is 0.039in).
       popup: whether the unit appears in the unit popup. */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72,               popup: true },   /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4,        popup: true },   /* 1 */
        { label: "pt",    pointsPerUnit: 1,                popup: true },   /* 2 */
        { label: "pica",  pointsPerUnit: 12,               popup: true },   /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54,        popup: true },   /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25, popup: true },   /* 5 */
        { label: "px",    pointsPerUnit: 1,                popup: true },   /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12,          popup: false },  /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000, popup: false },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36,          popup: false },  /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12,          popup: false }   /* 10 */
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

    /* 定規単位のコードから単位ラベルを取得（歯は H 表示、未対応コードは pt）/ Get the label for a rulerType code (code 5 shows as H; unknown codes fall back to pt) */
    function getRulerUnitLabelByCode(unitCode) {
        var unit = UNITS[unitCode] || UNITS[2];
        return (unitCode === 5 && HA_UNIT_PREF_KEYS["rulerType"]) ? "H" : unit.label;
    }

    /* ポップアップに表示する単位コード（表示順、UNITS から派生）/ Unit codes shown in the popup (in order, derived from UNITS) */
    var UNIT_POPUP_CODES = (function () {
        var codes = [];
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].popup) codes.push(i);
        }
        return codes;
    })();

    /* 単位コード → ポップアップのインデックス（未対応コードは -1）/ Unit code -> popup index (-1 if unsupported) */
    function unitCodeToPopupIndex(code) {
        for (var i = 0; i < UNIT_POPUP_CODES.length; i++) {
            if (UNIT_POPUP_CODES[i] === code) return i;
        }
        return -1;
    }

    // =========================================
    // アートボード枠線 / Artboard border
    // =========================================

    /* 枠線カラーのプリセット（ドロップダウンの並び順）/ Border color presets (dropdown order) */
    var STROKE_COLOR_PRESETS = [
        { labelKey: "color.lightBlue", r: 0.29, g: 0.52, b: 1.0 },
        { labelKey: "color.red",       r: 1.0,  g: 0.29, b: 0.29 },
        { labelKey: "color.green",     r: 0.0,  g: 0.65, b: 0.31 },
        { labelKey: "color.blue",      r: 0.0,  g: 0.45, b: 0.78 },
        { labelKey: "color.magenta",   r: 1.0,  g: 0.0,  b: 1.0 },
        { labelKey: "color.cyan",      r: 0.0,  g: 1.0,  b: 1.0 },
        { labelKey: "color.grey",      r: 0.65, g: 0.65, b: 0.65 },
        { labelKey: "color.black",     r: 0.0,  g: 0.0,  b: 0.0 },
        { labelKey: "color.yellow",    r: 1.0,  g: 1.0,  b: 0.0 }
    ];
    var STROKE_COLOR_BLACK_INDEX = 7;

    /* ドロップダウン用のカラー名配列を生成 / Build the color name list for the dropdown */
    function buildStrokeColorNames() {
        var names = [];
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            names.push(getLabel(STROKE_COLOR_PRESETS[i].labelKey));
        }
        return names;
    }

    /* RGB に最も近いプリセットの index を返す / Return the index of the preset closest to the given RGB */
    function findClosestStrokeColor(r, g, b) {
        var bestIdx = 0;
        var bestDist = Infinity;
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            var p = STROKE_COLOR_PRESETS[i];
            var dist = Math.abs(p.r - r) + Math.abs(p.g - g) + Math.abs(p.b - b);
            if (dist < bestDist) {
                bestDist = dist;
                bestIdx = i;
            }
        }
        return bestIdx;
    }

    // =========================================
    // 状態（キャッシュ）/ State (cache)
    // =========================================

    /* 常駐エンジンでは app.preferences への都度アクセスを避け、読み出した値をここへ保持 */
    /* In a persistent engine we avoid per-event app.preferences access; fetched values are cached here */
    var PREF_STATE = {
        rulerType: 2,
        cursorKeyLengthPt: 1.0
    };

    /* 現在の定規単位の pt 換算係数を取得 / Get the pt factor for the current ruler unit */
    function getCurrentPtPerUnit() {
        /* 常駐エンジンなので環境設定は読まず、キャッシュ済みのコードで UNITS を引く / cached code, no per-event preference read */
        return (UNITS[PREF_STATE.rulerType] || UNITS[2]).pointsPerUnit;
    }

    // =========================================
    // 環境設定の読み出し（同期・直接）/ Reading preferences (direct & synchronous)
    // 読み出しはエンジンを跨いでも安全なため、パレットエンジンで直接取得する
    // Reads are safe across engines, so fetch them directly in the palette engine
    // =========================================

    function readAllPreferences() {
        var p = app.preferences;
        function gb(key) { try { return p.getBooleanPreference(key); } catch (e) { return false; } }
        function gi(key) { try { return p.getIntegerPreference(key); } catch (e) { return 0; } }
        function gr(key) { try { return p.getRealPreference(key); } catch (e) { return 0; } }
        return {
            EnableActualPointTextSpaceAlign: gb("EnableActualPointTextSpaceAlign"),
            EnableActualAreaTextSpaceAlign: gb("EnableActualAreaTextSpaceAlign"),
            includeStrokeInBounds: gb("includeStrokeInBounds"),
            transformPatterns: gb("transformPatterns"),
            scaleLineWeight: gb("scaleLineWeight"),
            showGuides: gb("showGuides"),
            lockGuides: gb("lockGuides"),
            showArtboardLabelOnCanvas: gb("showArtboardLabelOnCanvas"),
            ArtboardBBColorRed: gr("ArtboardBBColorRed"),
            ArtboardBBColorGreen: gr("ArtboardBBColorGreen"),
            ArtboardBBColorBlue: gr("ArtboardBBColorBlue"),
            ArtboardBBWidth: gr("ArtboardBBWidth"),
            LiveEdit_State_Machine: gb("LiveEdit_State_Machine"),
            pastePlain: gb("plugin/FileClipboard/pasteWithoutFormatting"),
            policyForPreservingCorners: gi("policyForPreservingCorners"),
            uiCanvasIsWhite: gi("uiCanvasIsWhite"),
            rulerType: gi("rulerType"),
            cursorKeyLength: gr("cursorKeyLength")
        };
    }

    // =========================================
    // BridgeTalk 委譲（書き込み）/ BridgeTalk delegation (writes)
    // =========================================

    /* メインエンジン（target="illustrator"）へ環境設定コードを送って実行 / Send preference code to the main engine and run it */
    function runInMainEngine(bodyCode) {
        try {
            var bt = new BridgeTalk();
            bt.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
            bt.body = bodyCode;
            bt.onError = function (message) {
                /* no-op: 失敗時は既存の値を保持 / keep existing values on failure */
            };
            bt.send();
        } catch (e) {
            /* BridgeTalk 不可時は同一エンジンで直接実行 / Fallback: run directly in this engine */
            try {
                eval(bodyCode);
            } catch (e2) {
                // no-op
            }
        }
    }

    /* Boolean 環境設定をメインエンジンで設定 / Set a boolean preference on the main engine */
    function btSetBooleanPreference(prefKey, value) {
        runInMainEngine('app.preferences.setBooleanPreference("' + prefKey + '", ' + (value ? 'true' : 'false') + ');');
    }

    /* Integer 環境設定をメインエンジンで設定 / Set an integer preference on the main engine */
    function btSetIntegerPreference(prefKey, value) {
        runInMainEngine('app.preferences.setIntegerPreference("' + prefKey + '", ' + parseInt(value, 10) + ');');
    }

    /* Real 環境設定をメインエンジンで設定 / Set a real preference on the main engine */
    function btSetRealPreference(prefKey, value) {
        runInMainEngine('app.preferences.setRealPreference("' + prefKey + '", ' + Number(value) + ');');
    }

    /* アートボード名・枠線カラー・幅をまとめて設定し、キャンバスを再描画 / Apply artboard name/border color/width together, then refresh the canvas */
    function btApplyArtboardBorder(showName, r, g, b, width) {
        var body = '' +
            'var p=app.preferences;' +
            'p.setBooleanPreference("showArtboardLabelOnCanvas",' + (showName ? 'true' : 'false') + ');' +
            'p.setRealPreference("ArtboardBBColorRed",' + Number(r) + ');' +
            'p.setRealPreference("ArtboardBBColorGreen",' + Number(g) + ');' +
            'p.setRealPreference("ArtboardBBColorBlue",' + Number(b) + ');' +
            'p.setRealPreference("ArtboardBBWidth",' + Number(width) + ');' +
            'try{app.executeMenuCommand("zoomout");app.executeMenuCommand("zoomin");}catch(e){}';
        runInMainEngine(body);
    }

    // =========================================
    // カーソル移動量 / Cursor step
    // =========================================

    /* 現在単位の値を cursorKeyLength(pt) として保存 / Save value (in current unit) to cursorKeyLength as pt */
    function saveCursorKeyLengthInCurrentUnit(unitValue) {
        if (isNaN(unitValue) || unitValue < 0) return false;
        PREF_STATE.cursorKeyLengthPt = unitValue * getCurrentPtPerUnit();
        btSetRealPreference("cursorKeyLength", PREF_STATE.cursorKeyLengthPt);
        return true;
    }

    /* キャッシュ済み cursorKeyLength(pt) を現在単位の文字列(小数1桁)で取得 / Read cached cursorKeyLength as a current-unit string */
    function readCursorKeyLengthInCurrentUnit() {
        return (PREF_STATE.cursorKeyLengthPt / getCurrentPtPerUnit()).toFixed(1);
    }

    /* キー増加フィールドの値を保存 / Save the value from the key field */
    function saveCursorKeyLengthFromField(editText) {
        saveCursorKeyLengthInCurrentUnit(parseFloat(editText.text));
    }

    /* ===== デバウンス保存 / Debounced saving ===== */
    var __cursorKeyDebounceTaskId = null;
    var __cursorKeyPendingText = null;

    /* 保留中のテキストを実際に保存（scheduleTask から呼ばれる）/ Save the pending text (called from scheduleTask) */
    function __runSaveCursorKeyLength() {
        try {
            if (__cursorKeyPendingText !== null) {
                saveCursorKeyLengthInCurrentUnit(parseFloat(__cursorKeyPendingText));
            }
        } catch (e) {
            // no-op on failure
        }
        __cursorKeyDebounceTaskId = null;
    }
    /* scheduleTask の文字列はグローバルスコープで評価されるため $.global 経由で公開 / scheduleTask strings run in global scope, so expose via $.global */
    $.global.__aiQuickPrefsRunSave = __runSaveCursorKeyLength;

    /* 一定遅延後に保存をスケジュール（不可時は即時保存）/ Schedule a save after a short delay (immediate if unavailable) */
    function scheduleSaveCursorKeyLengthDebounced(editText, delayMs) {
        try {
            __cursorKeyPendingText = String(editText.text);
            if (__cursorKeyDebounceTaskId) {
                app.cancelTask(__cursorKeyDebounceTaskId);
                __cursorKeyDebounceTaskId = null;
            }
            __cursorKeyDebounceTaskId = app.scheduleTask("$.global.__aiQuickPrefsRunSave()", delayMs, false);
        } catch (e) {
            saveCursorKeyLengthFromField(editText);
        }
    }

    // =========================================
    // UI ヘルパー / UI helpers
    // =========================================

    /* 複数チェックボックスを Boolean 環境設定キーにバインド / Bind multiple checkboxes to boolean preference keys */
    function bindCheckboxes(pairs) {
        for (var i = 0; i < pairs.length; i++) {
            (function (pair) {
                pair.checkbox.onClick = function () {
                    btSetBooleanPreference(pair.prefKey, pair.checkbox.value === true);
                };
            })(pairs[i]);
        }
    }

    /* 標準パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel) {
        panel.orientation = 'column';
        panel.alignChildren = ['left', 'top'];
        panel.margins = [8, 20, 8, 15];
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /* パレットを構築して表示 / Build and show the palette */
    function main() {

        /* すでにパレットが開いていれば前面に出して終了 / If a palette already exists, bring it forward and return */
        try {
            if ($.global.__aiQuickPrefsPalette) {
                $.global.__aiQuickPrefsPalette.show();
                return;
            }
        } catch (e) {
            $.global.__aiQuickPrefsPalette = null;
        }

        /* 初期表示用に現在値を読み込む / Load current values for the initial display */
        var initialPrefs = readAllPreferences();
        PREF_STATE.rulerType = parseInt(initialPrefs.rulerType, 10);
        if (isNaN(PREF_STATE.rulerType)) PREF_STATE.rulerType = 2;
        PREF_STATE.cursorKeyLengthPt = parseFloat(initialPrefs.cursorKeyLength);
        if (isNaN(PREF_STATE.cursorKeyLengthPt)) PREF_STATE.cursorKeyLengthPt = 1.0;

        var dialog = new Window('palette', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        dialog.opacity = DIALOG_OPACITY;
        $.global.__aiQuickPrefsPalette = dialog;

        /* 閉じたら参照をクリア（次回は再構築）/ Clear the reference on close (rebuild next time) */
        dialog.onClose = function () {
            $.global.__aiQuickPrefsPalette = null;
        };

        var mainGroup = dialog.add('group');
        mainGroup.orientation = 'column';
        mainGroup.alignChildren = ['fill', 'top'];

        /* ===== 2カラム行（左右の列）/ Two-column row ===== */
        var columnsRow = mainGroup.add('group');
        columnsRow.orientation = 'row';
        columnsRow.alignChildren = ['fill', 'top'];

        var leftColumn = columnsRow.add('group');
        leftColumn.orientation = 'column';
        leftColumn.alignChildren = ['fill', 'top'];

        var rightColumn = columnsRow.add('group');
        rightColumn.orientation = 'column';
        rightColumn.alignChildren = ['fill', 'top'];

        /* ----- 左列：キー増加 / 変形と整列 / Left column: Key input / Transform & Align ----- */

        /* キー増加パネル（カーソル移動量）と単位ポップアップ / Key input panel (cursor step) with the unit popup */
        var keyInputPanel = leftColumn.add('panel', undefined, getLabel('panel.keyInput'));
        keyInputPanel.orientation = 'row';
        keyInputPanel.alignChildren = ['left', 'center'];
        keyInputPanel.margins = [8, 20, 8, 15];

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var keyFieldGroup = keyInputPanel.add('group');
        keyFieldGroup.orientation = 'row';
        keyFieldGroup.alignChildren = ['left', 'center'];
        keyFieldGroup.spacing = 0;
        keyFieldGroup.margins = 0;
        var keyField;
        var keyFieldStepper = addStepper(keyFieldGroup, function () { return keyField; }, {
            min: 0,
            /* 常に小数第1位で表示し、デバウンス保存 / always show one decimal and save (debounced) */
            onStep: function (numberInput) {
                numberInput.text = parseFloat(numberInput.text).toFixed(1);
                scheduleSaveCursorKeyLengthDebounced(numberInput, SAVE_DEBOUNCE_MS);
            }
        });
        keyField = keyFieldGroup.add('edittext', undefined, "1.0");
        keyField.helpTip = getLabel('tooltip.keyValue');
        keyField.characters = 4;

        var suppressUnitChange = false;
        var unitDropdown = keyInputPanel.add('dropdownlist', undefined, []);
        unitDropdown.helpTip = getLabel('tooltip.keyUnit');
        for (var u = 0; u < UNIT_POPUP_CODES.length; u++) {
            unitDropdown.add('item', getRulerUnitLabelByCode(UNIT_POPUP_CODES[u]));
        }
        unitDropdown.preferredSize.width = 70;

        /* 単位ポップアップ：選んだ単位を定規単位(rulerType)へ反映し、表示を再計算 / Unit popup: apply the chosen unit to rulerType and recompute the display */
        unitDropdown.onChange = function () {
            if (suppressUnitChange || !unitDropdown.selection) return;
            var code = UNIT_POPUP_CODES[unitDropdown.selection.index];
            PREF_STATE.rulerType = code;
            btSetIntegerPreference("rulerType", code);
            /* 保存済み pt 値は不変。新しい単位で再表示 / Stored pt value is unchanged; redisplay in the new unit */
            keyField.text = readCursorKeyLengthInCurrentUnit();
        };

        bindSteppedArrowKeys(keyField, keyFieldStepper); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */

        /* 変形と整列パネル / Transform & Align panel */
        var transformPanel = leftColumn.add('panel', undefined, getLabel('panel.transform'));
        setupPanel(transformPanel);

        /* プレビュー境界 / Preview bounds */
        var checkboxPreview = transformPanel.add('checkbox', undefined, getLabel('checkbox.previewBounds'));
        checkboxPreview.helpTip = getLabel('tooltip.previewBounds');
        checkboxPreview.onClick = function () {
            btSetBooleanPreference("includeStrokeInBounds", checkboxPreview.value === true);
        };

        /* パターンを変形 / Transform patterns */
        var checkboxPattern = transformPanel.add('checkbox', undefined, getLabel('checkbox.transformPattern'));
        checkboxPattern.helpTip = getLabel('tooltip.transformPattern');
        checkboxPattern.onClick = function () {
            btSetBooleanPreference("transformPatterns", checkboxPattern.value === true);
        };

        /* 角を拡大・縮小（1=ON, 2=OFF）/ Scale corners (1=ON, 2=OFF) */
        var checkboxCorner = transformPanel.add('checkbox', undefined, getLabel('checkbox.scaleCorners'));
        checkboxCorner.helpTip = getLabel('tooltip.scaleCorners');
        checkboxCorner.onClick = function () {
            btSetIntegerPreference("policyForPreservingCorners", checkboxCorner.value ? 1 : 2);
        };

        /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
        var checkboxStroke = transformPanel.add('checkbox', undefined, getLabel('checkbox.scaleStroke'));
        checkboxStroke.helpTip = getLabel('tooltip.scaleStroke');
        checkboxStroke.onClick = function () {
            btSetBooleanPreference("scaleLineWeight", checkboxStroke.value === true);
        };

        /* ----- 右列：字形の境界に整列 / ガイドと定規 / Right column: Glyph bounds / Guides & Rulers ----- */

        /* 字形の境界に整列パネル / Align to glyph bounds panel */
        var glyphPanel = rightColumn.add('panel', undefined, getLabel('panel.glyphBounds'));
        setupPanel(glyphPanel);

        var checkboxPoint = glyphPanel.add('checkbox', undefined, getLabel('checkbox.pointText'));
        checkboxPoint.helpTip = getLabel('tooltip.glyphBounds');
        var checkboxArea = glyphPanel.add('checkbox', undefined, getLabel('checkbox.areaText'));
        checkboxArea.helpTip = getLabel('tooltip.glyphBounds');

        bindCheckboxes([
            { checkbox: checkboxPoint, prefKey: 'EnableActualPointTextSpaceAlign' },
            { checkbox: checkboxArea, prefKey: 'EnableActualAreaTextSpaceAlign' }
        ]);

        /* ガイドと定規パネル / Guides & Rulers panel */
        var guidePanel = rightColumn.add('panel', undefined, getLabel('panel.guide'));
        setupPanel(guidePanel);

        /* ガイドを表示 / Show guides */
        var checkboxGuideShow = guidePanel.add('checkbox', undefined, getLabel('checkbox.guideShow'));
        checkboxGuideShow.helpTip = getLabel('tooltip.guideShow');
        checkboxGuideShow.onClick = function () {
            btSetBooleanPreference("showGuides", checkboxGuideShow.value === true);
        };

        /* ガイドをロック / Lock guides */
        var checkboxGuideLock = guidePanel.add('checkbox', undefined, getLabel('checkbox.guideLock'));
        checkboxGuideLock.helpTip = getLabel('tooltip.guideLock');
        checkboxGuideLock.onClick = function () {
            btSetBooleanPreference("lockGuides", checkboxGuideLock.value === true);
        };

        /* ビデオ定規（メニューコマンドのトグル）/ Video ruler (menu-command toggle) */
        var btnVideoRuler = guidePanel.add('button', undefined, getLabel('button.videoRuler'));
        btnVideoRuler.alignment = ['left', 'top']; /* 幅いっぱいにしない（ラベル幅）/ Do not fill width (size to label) */
        btnVideoRuler.onClick = function () {
            runInMainEngine('try{app.executeMenuCommand("videoruler");}catch(e){}');
        };

        /* ----- 全幅：アートボード名と枠線 / その他（一番下）/ Full width: Artboard / Other (bottom) ----- */

        /* アートボード名と枠線パネル / Artboard name & border panel */
        var artboardPanel = mainGroup.add('panel', undefined, getLabel('panel.artboard'));
        artboardPanel.orientation = 'column';
        artboardPanel.alignChildren = ['fill', 'top'];
        artboardPanel.margins = [8, 20, 8, 15];

        var suppressArtboardChange = false;

        /* アートボード名を表示 / Show artboard name */
        var cbShowArtboardName = artboardPanel.add('checkbox', undefined, getLabel('checkbox.showArtboardName'));
        cbShowArtboardName.helpTip = getLabel('tooltip.showArtboardName');
        cbShowArtboardName.onClick = function () {
            applyArtboard();
        };

        /* カンバスカラーをホワイトに（ON=1, OFF=0）/ Canvas color white (ON=1, OFF=0) */
        var checkboxCanvasWhite = artboardPanel.add('checkbox', undefined, getLabel('checkbox.canvasWhite'));
        checkboxCanvasWhite.helpTip = getLabel('tooltip.canvasWhite');
        checkboxCanvasWhite.onClick = function () {
            btSetIntegerPreference("uiCanvasIsWhite", checkboxCanvasWhite.value ? 1 : 0);
        };

        /* アートボードの枠線サブパネル / Artboard border sub-panel */
        var artboardBorderPanel = artboardPanel.add('panel', undefined, getLabel('panel.artboardBorder'));
        artboardBorderPanel.orientation = 'column';
        artboardBorderPanel.alignChildren = ['left', 'top'];
        artboardBorderPanel.margins = [8, 20, 8, 15];

        /* ハイライトのカラー / Highlight color */
        var strokeColorRow = artboardBorderPanel.add('group');
        strokeColorRow.orientation = 'row';
        strokeColorRow.alignChildren = ['left', 'center'];
        strokeColorRow.add('statictext', undefined, labelText('label.strokeColor'));
        var ddStrokeColor = strokeColorRow.add('dropdownlist', undefined, buildStrokeColorNames());
        ddStrokeColor.onChange = function () {
            if (suppressArtboardChange || !ddStrokeColor.selection) return;
            applyArtboard();
        };

        /* ストロークの幅（1〜4）/ Stroke width (1-4) */
        var strokeWidthRow = artboardBorderPanel.add('group');
        strokeWidthRow.orientation = 'row';
        strokeWidthRow.alignChildren = ['left', 'center'];
        strokeWidthRow.add('statictext', undefined, labelText('label.strokeWidth'));
        var rbStrokeWidth1 = strokeWidthRow.add('radiobutton', undefined, '1');
        var rbStrokeWidth2 = strokeWidthRow.add('radiobutton', undefined, '2');
        var rbStrokeWidth3 = strokeWidthRow.add('radiobutton', undefined, '3');
        var rbStrokeWidth4 = strokeWidthRow.add('radiobutton', undefined, '4');
        var rbStrokeWidths = [rbStrokeWidth1, rbStrokeWidth2, rbStrokeWidth3, rbStrokeWidth4];
        for (var sw = 0; sw < rbStrokeWidths.length; sw++) {
            rbStrokeWidths[sw].onClick = function () {
                applyArtboard();
            };
        }

        /* 選択中のストローク幅（1〜4）を取得 / Get the selected stroke width (1-4) */
        function getSelectedStrokeWidth() {
            for (var i = 0; i < rbStrokeWidths.length; i++) {
                if (rbStrokeWidths[i].value) return i + 1;
            }
            return 1;
        }

        /* アートボード名・枠線の現在 UI 値をまとめて保存 / Save the current artboard name/border UI state */
        function applyArtboard() {
            var idx = ddStrokeColor.selection ? ddStrokeColor.selection.index : STROKE_COLOR_BLACK_INDEX;
            var c = STROKE_COLOR_PRESETS[idx];
            btApplyArtboardBorder(cbShowArtboardName.value === true, c.r, c.g, c.b, getSelectedStrokeWidth());
        }

        /* その他パネル（一番下）/ Other panel (bottom) */
        var etcPanel = mainGroup.add('panel', undefined, getLabel('panel.etc'));
        setupPanel(etcPanel);

        /* リアルタイムの描画と編集 / Real-time drawing & editing */
        var checkboxRealtime = etcPanel.add('checkbox', undefined, getLabel('checkbox.realtimeDrawing'));
        checkboxRealtime.onClick = function () {
            btSetBooleanPreference("LiveEdit_State_Machine", checkboxRealtime.value === true);
        };

        /* 書式なしペースト / Paste without formatting */
        var checkboxPastePlain = etcPanel.add('checkbox', undefined, getLabel('checkbox.pastePlain'));
        checkboxPastePlain.onClick = function () {
            btSetBooleanPreference("plugin/FileClipboard/pasteWithoutFormatting", checkboxPastePlain.value === true);
        };

        /* 読み出した環境設定を UI へ反映 / Apply fetched preferences to the UI */
        function applyPreferencesToUI(map) {
            function asBool(prefKey) {
                return String(map[prefKey]) === "true";
            }
            checkboxPoint.value = asBool('EnableActualPointTextSpaceAlign');
            checkboxArea.value = asBool('EnableActualAreaTextSpaceAlign');
            checkboxPreview.value = asBool('includeStrokeInBounds');
            checkboxPattern.value = asBool('transformPatterns');
            checkboxStroke.value = asBool('scaleLineWeight');
            checkboxGuideShow.value = asBool('showGuides');
            checkboxGuideLock.value = asBool('lockGuides');
            checkboxRealtime.value = asBool('LiveEdit_State_Machine');
            checkboxPastePlain.value = asBool('pastePlain');
            checkboxCorner.value = (parseInt(map['policyForPreservingCorners'], 10) === 1);

            /* アートボード名と枠線 / Artboard name & border */
            cbShowArtboardName.value = asBool('showArtboardLabelOnCanvas');
            checkboxCanvasWhite.value = (parseInt(map['uiCanvasIsWhite'], 10) === 1);
            var cr = parseFloat(map['ArtboardBBColorRed']);
            var cg = parseFloat(map['ArtboardBBColorGreen']);
            var cb2 = parseFloat(map['ArtboardBBColorBlue']);
            if (isNaN(cr)) cr = 0;
            if (isNaN(cg)) cg = 0;
            if (isNaN(cb2)) cb2 = 0;
            suppressArtboardChange = true;
            ddStrokeColor.selection = findClosestStrokeColor(cr, cg, cb2);
            suppressArtboardChange = false;
            var w = Math.round(parseFloat(map['ArtboardBBWidth']));
            if (isNaN(w)) w = 1;
            if (w < 1) w = 1;
            if (w > 4) w = 4;
            for (var wi = 0; wi < rbStrokeWidths.length; wi++) {
                rbStrokeWidths[wi].value = (wi === w - 1);
            }

            var ruler = parseInt(map['rulerType'], 10);
            PREF_STATE.rulerType = isNaN(ruler) ? PREF_STATE.rulerType : ruler;
            var ckl = parseFloat(map['cursorKeyLength']);
            PREF_STATE.cursorKeyLengthPt = isNaN(ckl) ? PREF_STATE.cursorKeyLengthPt : ckl;

            /* 単位ポップアップを現在の rulerType に同期（onChange を発火させない）/ Sync the unit popup to the current rulerType (without firing onChange) */
            var unitIdx = unitCodeToPopupIndex(PREF_STATE.rulerType);
            if (unitIdx >= 0) {
                suppressUnitChange = true;
                unitDropdown.selection = unitIdx;
                suppressUnitChange = false;
            }
            keyField.text = readCursorKeyLengthInCurrentUnit();
        }

        /* キー増加の確定時に保存（不正値は元へ戻す）/ Save on commit (restore on invalid input) */
        keyField.onChange = function () {
            if (!saveCursorKeyLengthInCurrentUnit(parseFloat(keyField.text))) {
                keyField.text = readCursorKeyLengthInCurrentUnit();
            }
        };

        /* 構築直後に現在値を反映（常に最新の環境設定を表示）/ Populate from current values right after building */
        applyPreferencesToUI(initialPrefs);

        /* 表示時：フォーカスと最新値の再読込 / On show: focus and reload current values */
        dialog.onShow = function () {
            keyField.active = true;
            applyPreferencesToUI(readAllPreferences());
        };

        /* 再アクティブ時：外部変更（環境設定ダイアログ等）へクリックで追従 / On re-activate: follow external changes via click */
        dialog.onActivate = function () {
            applyPreferencesToUI(readAllPreferences());
        };

        dialog.show();
    }

    main();

}());
