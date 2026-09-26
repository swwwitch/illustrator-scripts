#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択している文字を対象に、フォントサイズと水平比率／垂直比率を調整します。
ライブプレビューで結果を確認しながら調整でき、キャンセルすると開く前の状態に戻ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSize.md

note記事も参照してください。
https://note.com/dtp_tranist/n/xxxxxxxx

### Overview

Adjusts the font size and the horizontal and vertical scale of the selected characters.
A live preview shows the result, and cancelling restores the state from before the dialog opened.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSize.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustFontSize";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSize.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSize.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DIALOG_OPACITY  = 0.98;  /* ダイアログ透明度 / dialog opacity */
    var DIALOG_OFFSET_X = 0;     /* 表示位置の横オフセット / horizontal offset on show */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [16, 20, 16, 12];  /* パネル余白 / panel margins */
    var PANEL_SPACING = 8;                 /* パネル内の標準間隔 / default spacing inside panels */
    var FIELD_SPACING = 6;                 /* 入力行どうしの間隔 / spacing between field rows */
    var LABEL_WIDTH   = 118;               /* ラベル幅（揃える）/ unified label width */
    var BUTTON_WIDTH  = 90;                /* OK・キャンセルの幅 / width of OK and Cancel */
    var CONVERT_BUTTON_WIDTH = 150;        /* 実サイズ↔見かけボタンの幅 / width of the actual↔apparent button */
    var EDIT_CHARACTERS    = 4;            /* 入力欄の文字数 / width of an edittext in characters */
    var READOUT_CHARACTERS = 5;            /* 表示専用テキストの文字数 / width of a readout in characters */

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
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
    // 6. この欄に別の↑↓キー処理を付けない（二重に効く）
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
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkStepperUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var STEPPER_UI_DARK           = isDarkStepperUI();
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
        var upTooltip = stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp";
        var downTooltip = stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown";
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

    /**
     * 実行環境のロケールからUIの表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "フォントサイズの調整", en: "Font Size Adjuster" }
        },
        panel: {
            fontSize: { ja: "フォントサイズの調整", en: "Font Size Adjustment" }
        },
        fieldLabel: {
            fontSize: { ja: "フォントサイズ", en: "Font Size" },
            scale: { ja: "水平比率/垂直比率", en: "Scale" },
            apparent: { ja: "見かけ", en: "Apparent" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" },
            toApparent: { ja: "実サイズ↔見かけ", en: "Actual ↔ Apparent" }
        },
        alert: {
            selectText: { ja: "テキストを選択してください", en: "Please select text" },
            previewError: { ja: "プレビュー更新エラー / Preview update error: ", en: "Preview update error: " }
        },
        tooltip: {
            scale: { ja: "水平比率・垂直比率を同じ値でまとめて設定します。", en: "Sets the horizontal and vertical scale together to the same value." },
            apparent: { ja: "フォントサイズ×比率で計算した、実際の見た目のサイズです。", en: "The actual visual size, computed as font size × scale." },
            toApparent: {
                ja: "サイズ×比率を見かけサイズとして実フォントサイズに焼き込み、比率を100%にします。もう一度押すと焼き込み前の比率付き状態に戻ります",
                en: "Bakes size × scale into the actual font size at 100%. Press again to restore the previous scaled state."
            },
            reset: {
                ja: "調整を取り消して、開いた直後の状態に戻します。optionキーを押しながらクリックすると、選択している文字すべてを先頭文字のフォントサイズに統一し、水平比率・垂直比率を100%に揃えて適用します。",
                en: "Discards the adjustments and returns to the just-opened state. Option-click to unify every selected character to the first character's font size and set the horizontal / vertical scale to 100%."
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
        }
    };

    /**
     * キーからラベルを現在の言語で取得する（"panel.fontSize" のようにドット区切り）
     * @param {string} key - カテゴリ名とキー名をドットでつないだラベルキー
     * @returns {string} 現在の言語のラベル文字列（未定義の場合は英語にフォールバック）
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelEntry = LABELS[keyParts[0]][keyParts[1]];
        return labelEntry[uiLang] || labelEntry.en;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} key - ラベルキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ":");
    }

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

    // =========================================
    // 数値ユーティリティ / Number helpers
    // =========================================

    /**
     * コントロールの文字列を数値として読み取る
     * @param {EditText|StaticText} control - 対象のコントロール
     * @returns {number|null} 数値（空欄や数値でない場合は null）
     */
    function readNumber(control) {
        var parsedValue = parseFloat(control.text);
        return isNaN(parsedValue) ? null : parsedValue;
    }

    /**
     * 小数第1位に丸める
     * @param {number} value - 丸める値
     * @returns {number} 小数第1位までの値
     */
    function roundToTenth(value) {
        return Math.round(value * 10) / 10;
    }

    /**
     * フォントサイズと比率から見かけのサイズを求める
     * @param {number} size - フォントサイズ
     * @param {number} scale - 比率（%）
     * @returns {number} 見かけのサイズ（小数第2位まで）
     */
    function calculateApparentSize(size, scale) {
        return Math.round(size * scale) / 100;
    }

    // =========================================
    // レイアウト補助 / Layout helpers
    // =========================================

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は既定値）
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
     * グループに共通のレイアウト設定を適用する
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は既定値）
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

    /**
     * ラベル＋コントロール＋単位の行を追加する
     * @param {Panel|Group} parent - 追加先のコンテナ
     * @param {string} labelKey - ラベルキー
     * @param {string} controlType - "edittext"（入力欄）または "statictext"（表示専用）
     * @param {string} initialText - コントロールの初期表示文字列
     * @param {string} unitText - 単位表示の文字列
     * @param {Object} [stepOptions] - 入力欄の左に∧∨を付けるときの min / max（増減後は onChange を通知する）
     * @returns {{label: StaticText, control: EditText|StaticText, unit: StaticText}} 生成した行の各コントロール
     */
    function addRow(parent, labelKey, controlType, initialText, unitText, stepOptions) {
        var rowGroup = parent.add("group");
        setupGroup(rowGroup, "row");
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelKey));
        rowLabel.justify = "right";
        var isEditable = (controlType === "edittext");
        var controlParent = rowGroup;
        var valueControl;
        if (stepOptions) {
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            controlParent = rowGroup.add("group");
            controlParent.orientation = "row";
            controlParent.alignChildren = ["left", "center"];
            controlParent.spacing = 0;
            controlParent.margins = 0;
            /* text の代入では onChange が発火しないので通知する / assigning text does not fire onChange */
            stepOptions.onStep = function (numberInput) { numberInput.notify("onChange"); };
            var stepperGroup = addStepper(controlParent, function () { return valueControl; }, stepOptions);
        }
        valueControl = controlParent.add(controlType, undefined, initialText);
        valueControl.characters = isEditable ? EDIT_CHARACTERS : READOUT_CHARACTERS;
        if (isEditable) valueControl.justify = "right";
        if (stepOptions) bindSteppedArrowKeys(valueControl, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */
        var unitLabel = rowGroup.add("statictext", undefined, unitText);
        return { label: rowLabel, control: valueControl, unit: unitLabel };
    }

    /**
     * 行（ラベル＋コントロール＋単位）にまとめてヘルプチップを設定する
     * @param {{label: StaticText, control: EditText|StaticText, unit: StaticText}} rowControls - addRow が返した行オブジェクト
     * @param {string} tooltipText - 設定するヘルプチップ文字列
     * @returns {void}
     */
    function setRowTooltip(rowControls, tooltipText) {
        rowControls.label.helpTip = tooltipText;
        rowControls.control.helpTip = tooltipText;
        rowControls.unit.helpTip = tooltipText;
    }

    /**
     * 行のラベル・コントロール・単位をまとめて有効／無効にする
     * @param {{label: StaticText, control: EditText|StaticText, unit: StaticText}} rowControls - addRow が返した行オブジェクト
     * @param {boolean} enabled - 有効にするなら true
     * @returns {void}
     */
    function setRowEnabled(rowControls, enabled) {
        rowControls.label.enabled = enabled;
        rowControls.control.enabled = enabled;
        rowControls.unit.enabled = enabled;
    }

    /**
     * 複数ラベルの幅を揃える
     * @param {number} labelWidth - 設定する幅（px）
     * @param {StaticText[]} labelControls - 幅を揃えるラベルの配列
     * @returns {void}
     */
    function alignLabelWidths(labelWidth, labelControls) {
        for (var i = 0; i < labelControls.length; i++) {
            labelControls[i].preferredSize.width = labelWidth;
        }
    }

    // =========================================
    // 選択取得 / Selection
    // =========================================

    /**
     * 選択中のテキスト範囲を取得する
     * @returns {TextRange[]} 選択されているテキスト範囲の配列（なければ空配列）
     */
    function getTextSelection() {
        var docSelection = app.activeDocument.selection;
        var textRanges = [];
        if (!docSelection) return textRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (docSelection.constructor.name === "TextRange") {
            textRanges.push(docSelection);
            return textRanges;
        }
        for (var i = 0; i < docSelection.length; i++) {
            var selectedItem = docSelection[i];
            if (selectedItem.constructor.name === "TextFrame") {
                textRanges.push(selectedItem.textRange);
            } else if (selectedItem.constructor.name === "TextRange") {
                textRanges.push(selectedItem);
            }
        }
        return textRanges;
    }

    /**
     * テキスト範囲の先頭文字を取得する
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @returns {TextRange|null} 最初の文字（文字がなければ null）
     */
    function findFirstChar(textRanges) {
        for (var i = 0; i < textRanges.length; i++) {
            if (textRanges[i].characters.length > 0) return textRanges[i].characters[0];
        }
        return null;
    }

    /**
     * テキスト範囲のすべての文字にコールバックを適用する
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @param {function} charCallback - 各文字に対して実行する処理
     * @returns {void}
     */
    function forEachChar(textRanges, charCallback) {
        for (var i = 0; i < textRanges.length; i++) {
            var characters = textRanges[i].characters;
            for (var j = 0; j < characters.length; j++) {
                charCallback(characters[j]);
            }
        }
    }

    // =========================================
    // PreviewManager
    // プレビュー時にUndo履歴を汚さないための小さな管理クラス。
    // - addStep(): 変更処理を実行してundoDepthをカウント
    // - rollback(): 適用済みのプレビューをすべて取り消し
    // undoDepth = 適用してまだ取り消していないステップ数 / steps applied and not yet undone
    // 適用中の例外でダイアログごと落ちないよう addStep だけ try で囲み、
    // 上下キー連打で同じアラートが溢れないよう同一メッセージは1度だけ表示する。
    // =========================================

    /**
     * プレビュー適用とUndoの深さを管理するクラス
     * @returns {void}
     */
    function PreviewManager() {
        this.undoDepth = 0;
        var lastReportedError = "";

        /**
         * 変更処理を1ステップとして実行し、Undoの深さを数える
         * @param {function} applyChange - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (applyChange) {
            /* 文字属性の書き込みが失敗しうる / writing character attributes can fail */
            try {
                applyChange();
                lastReportedError = "";
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                var errorMessage = getLabel("alert.previewError") + e;
                if (errorMessage === lastReportedError) return;
                lastReportedError = errorMessage;
                alert(errorMessage);
            }
        };

        /**
         * 適用済みのプレビューをすべて取り消す
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    /**
     * フォントサイズの調整パネルを構築する（イベントの配線は呼び出し側で行う）
     * @param {Window} parentDialog - 追加先のダイアログ
     * @param {string} unitLabel - サイズ欄に表示する単位ラベル
     * @returns {{sizeRow: object, scaleRow: object, apparentRow: object, convertButton: Button}} 構築したコントロール
     */
    function buildFontSizePanel(parentDialog, unitLabel) {
        var fontSizePanel = parentDialog.add("panel", undefined, getLabel("panel.fontSize"));
        setupPanel(fontSizePanel, FIELD_SPACING);

        var sizeRow = addRow(fontSizePanel, "fieldLabel.fontSize", "edittext", "0", unitLabel, { min: 0.1 });
        var scaleRow = addRow(fontSizePanel, "fieldLabel.scale", "edittext", "100", "%", { min: 1 });
        setRowTooltip(scaleRow, getLabel("tooltip.scale"));
        var apparentRow = addRow(fontSizePanel, "fieldLabel.apparent", "statictext", "--", unitLabel);
        setRowTooltip(apparentRow, getLabel("tooltip.apparent"));

        var convertButton = fontSizePanel.add("button", undefined, getLabel("button.toApparent"));
        convertButton.helpTip = getLabel("tooltip.toApparent");
        convertButton.alignment = "right";
        convertButton.preferredSize.width = CONVERT_BUTTON_WIDTH;

        alignLabelWidths(LABEL_WIDTH, [sizeRow.label, scaleRow.label, apparentRow.label]);
        return { sizeRow: sizeRow, scaleRow: scaleRow, apparentRow: apparentRow, convertButton: convertButton };
    }

    /**
     * 下部のボタン行を構築する（左＝リセット／右＝キャンセル・OK）
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{btnReset: Button, btnCancel: Button, btnOK: Button}} 構築したボタン
     */
    function buildButtonRow(parentDialog) {
        var btnRowGroup = parentDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "fill";
        btnRowGroup.alignChildren = ["fill", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignment = ["left", "center"];
        var btnReset = btnLeftGroup.add("button", undefined, getLabel("button.reset"));
        btnReset.helpTip = getLabel("tooltip.reset");

        /* 左右のボタンを両端に押し広げるスペーサー / spacer that pushes both sides apart */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "center"];

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;

        return { btnReset: btnReset, btnCancel: btnCancel, btnOK: btnOK };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択している文字のフォントサイズと比率を調整するダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length <= 0) {
            return;
        }

        var targetRanges = getTextSelection();
        if (targetRanges.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var previewManager = new PreviewManager();
        var textUnit = getUnitInfo("text/units");
        var unitLabel = textUnit.label;
        var unitFactor = textUnit.pointsPerUnit;

        var adjustDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        adjustDialog.alignChildren = "fill";
        adjustDialog.opacity = DIALOG_OPACITY;

        var fontSizeUI = buildFontSizePanel(adjustDialog, unitLabel);
        var buttonUI = buildButtonRow(adjustDialog);
        var sizeInput = fontSizeUI.sizeRow.control;
        var scaleInput = fontSizeUI.scaleRow.control;
        var apparentRow = fontSizeUI.apparentRow;

        /* 焼き込み前の状態（順方向で保存→逆方向で復元）。手動でサイズ/比率を変えたら無効化
           pre-bake state (saved on forward, restored on back); cleared when size/scale is edited by hand */
        var apparentToggleState = null;

        // ---- 値の適用・プレビュー / Apply values & preview ----

        /**
         * 現在の入力値を選択している文字にまとめて適用する（空欄の項目は適用しない）
         * @returns {void}
         */
        function applyCurrentValues() {
            var size = readNumber(sizeInput);
            var scale = readNumber(scaleInput);
            if (size === null && scale === null) return;
            var sizeInPt = (size === null) ? null : size * unitFactor;
            forEachChar(targetRanges, function (character) {
                if (sizeInPt !== null) character.size = sizeInPt;
                if (scale !== null) {
                    character.characterAttributes.horizontalScale = scale;
                    character.characterAttributes.verticalScale = scale;
                }
            });
        }

        /**
         * Undo履歴を汚さずにプレビューを更新し、見かけサイズの表示も更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(applyCurrentValues);
            updateApparentSizeDisplay();
        }

        // ---- 表示更新 / Display updates ----

        /**
         * 見かけサイズの表示を更新する（比率100%のときはディム表示）
         * @returns {void}
         */
        function updateApparentSizeDisplay() {
            var size = readNumber(sizeInput);
            var scale = readNumber(scaleInput);
            var hasValue = (size !== null && scale !== null);
            apparentRow.control.text = hasValue ? calculateApparentSize(size, scale) + "" : "--";
            setRowEnabled(apparentRow, scale !== 100);
        }

        /**
         * 選択している文字の先頭の現在値を読み取って入力欄に反映する
         * @returns {void}
         */
        function loadValuesFromSelection() {
            /* 実際の値を読み直すので、焼き込み前の保存状態（トグル）は破棄する
               Reloading the actual values invalidates the saved pre-bake (toggle) state */
            apparentToggleState = null;
            var firstChar = findFirstChar(targetRanges);
            sizeInput.text = firstChar ? roundToTenth(firstChar.size / unitFactor) + "" : "";
            scaleInput.text = firstChar ? roundToTenth(firstChar.characterAttributes.horizontalScale) + "" : "";
            updateApparentSizeDisplay();
        }

        // ---- イベント / Events ----

        /**
         * サイズ・比率の確定入力を受けてプレビューと表示を更新する
         * @returns {void}
         */
        function onValueChanged() {
            apparentToggleState = null; /* 手動編集でトグル復元を無効化 / manual edit invalidates the toggle */
            updatePreview();
        }

        /* サイズ・比率は「入力値をそのまま適用」。loadValuesFromSelection() で入力欄を読み直すと
           入力値が丸めで戻る恐れがあるため onChange では呼ばない（見かけ表示だけ更新する）
           apply the typed value as-is; do NOT reload the fields on change (re-reading them
           could snap the typed value back via rounding). Only refresh the apparent readout */
        sizeInput.onChange = onValueChanged;
        scaleInput.onChange = onValueChanged;
        sizeInput.onChanging = updateApparentSizeDisplay;
        scaleInput.onChanging = updateApparentSizeDisplay;

        /* 実サイズ↔見かけのトグル / toggle between actual size and apparent (baked) size
           順方向：サイズ×比率を実サイズに焼き込み比率100%へ。逆方向：直前の比率付き状態へ戻す
           forward: bake size × scale into the actual size at 100%; back: restore the previous scaled state */
        fontSizeUI.convertButton.onClick = function () {
            var size = readNumber(sizeInput);
            var scale = readNumber(scaleInput);
            if (size === null || scale === null) return;
            if (apparentToggleState !== null) {
                sizeInput.text = apparentToggleState.size + "";
                scaleInput.text = apparentToggleState.scale + "";
                apparentToggleState = null;
            } else {
                apparentToggleState = { size: size, scale: scale };
                sizeInput.text = calculateApparentSize(size, scale) + "";
                scaleInput.text = "100";
            }
            updatePreview();
        };

        /* リセット：プレビューを取り消して開いた直後の状態に戻す。optionキー併用のときは
           そこからさらに、選択している文字すべてを先頭文字のサイズ・比率100%に統一して適用する
           Reset: undo the preview and return to the just-opened state; with the option key,
           also unify every selected character to the first character's size at 100% scale */
        buttonUI.btnReset.onClick = function () {
            previewManager.rollback();
            loadValuesFromSelection(); /* 調整前の先頭文字の値を読み直す / re-read the pre-adjustment values */
            if (!ScriptUI.environment.keyboardState.altKey) return;
            scaleInput.text = "100";
            updatePreview();
        };

        buttonUI.btnOK.onClick = function () {
            /* プレビュー分を戻し、本適用を1回だけ実行して確定（Undo履歴は1つ）
               undo the preview, then apply once so it lands as a single undo entry */
            previewManager.rollback();
            applyCurrentValues();
            adjustDialog.close();
        };

        buttonUI.btnCancel.onClick = function () {
            /* 開いてから適用した分をすべて取り消してから閉じる / undo everything applied since open, then close */
            previewManager.rollback();
            adjustDialog.close(2);
        };

        adjustDialog.onShow = function () {
            adjustDialog.location = [adjustDialog.location[0] + DIALOG_OFFSET_X, adjustDialog.location[1]];
            scaleInput.active = true;
        };

        /* 開いた時点では何も適用しない（現在の状態をそのまま保持）。値を変更したときだけプレビュー適用
           apply nothing on open (keep the current state as-is); preview only kicks in once a value changes */
        loadValuesFromSelection();
        adjustDialog.show();
    }

    main();

})();
