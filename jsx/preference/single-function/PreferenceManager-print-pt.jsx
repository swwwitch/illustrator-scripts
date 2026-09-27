#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

単位と数値インクリメントを、ダイアログから指定・変更します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManager-print-pt.md

### Overview

Sets the units and the numeric increment from a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager-print-pt.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PreferenceManager-print-pt";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManager-print-pt.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager-print-pt.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "単位とインクリメント設定", en: "Units and Increments" }
        },
        panel: {
            units:      { ja: "単位", en: "Units" },
            increments: { ja: "増減値", en: "Increments" }
        },
        radio: {
            modePt: { ja: "プリント（pt）", en: "Print (pt)" },
            modeQ:  { ja: "プリント（Q）", en: "Print (Q)" },
            modePx: { ja: "オンスクリーン（px）", en: "Onscreen (px)" }
        },
        fieldLabel: {
            general:  { ja: "一般", en: "General" },
            stroke:   { ja: "線", en: "Stroke" },
            type:     { ja: "文字", en: "Text" },
            asian:    { ja: "東アジア言語", en: "East Asian" },
            key:      { ja: "キー増加", en: "Keyboard increment" },
            radius:   { ja: "角丸の半径", en: "Corner radius" },
            size:     { ja: "フォントサイズ", en: "Font size" },
            baseline: { ja: "ベースラインシフト", en: "Baseline shift" }
        },
        tooltip: {
            modePt:   { ja: "一般=mm、線=pt、文字=pt にまとめて切り替えます。", en: "Sets General=mm, Stroke=pt, Text=pt." },
            modeQ:    { ja: "一般=mm、線=mm、文字=Q にまとめて切り替えます。", en: "Sets General=mm, Stroke=mm, Text=Q." },
            modePx:   { ja: "すべての単位を px に切り替えます。", en: "Sets every unit to px." },
            general:  { ja: "定規やパネルに表示される、既定の長さの単位です。", en: "Default unit shown on rulers and panels." },
            stroke:   { ja: "線幅の入力・表示に使う単位です。", en: "Unit used for stroke weights." },
            type:     { ja: "フォントサイズや行送りに使う単位です。", en: "Unit used for font size and leading." },
            asian:    { ja: "東アジア言語のオプションで使う単位です。", en: "Unit used for East Asian typography options." },
            key:      { ja: "矢印キー1回で動く距離です。", en: "How far one arrow key press moves things." },
            radius:   { ja: "角丸ツールの既定の半径です。", en: "Default radius used by the rounded rectangle tool." },
            size:     { ja: "文字サイズ・行送りを増減する1回ぶんの量です。", en: "How much one step changes the type size or leading." },
            baseline: { ja: "ベースラインシフトを増減する1回ぶんの量です。", en: "How much one step changes the baseline shift." },
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
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.units" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのドット区切りキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ": ");
    }

    /**
     * ∧∨付きの増減値の入力欄を行に追加する（∧∨と入力欄は隙間0で突き合わせ、↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} rowGroup - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @returns {EditText} 入力欄
     */
    function addIncrementInput(rowGroup, initialText) {
        var stepperInputGroup = rowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        /* 増減値は正の数。単位は隣の表示に任せ、欄には数値だけを入れる / increments stay positive; the unit is shown next to the field */
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, { min: 0.01 });
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    function main() {
        // --- 単位ラベル更新関数 ---
        // 左パネルの単位選択に応じて右パネルの単位表示を更新する
        function updateIncrementLabels() {
            lblKeyUnit.text      = ddGeneral.selection.text;
            lblRadiusUnit.text   = ddGeneral.selection.text;
            lblSizeUnit.text     = ddType.selection.text;
            lblBaselineUnit.text = ddType.selection.text;
        }

        // ドキュメントが開かれていない場合は処理を中止
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてから実行してください。");
            return;
        }

        // Illustratorのバージョン番号と現在の単位設定を取得
        var versionParts = app.version.split(".");
        var aiMajor = parseInt(versionParts[0], 10);
        var aiMinor = parseInt(versionParts[1] || "0", 10);

        /* 4つの環境設定キーで同じドロップダウンを使うため、単位コード5は「Q/H」の中立表記にする
           （他のスクリプトの UNITS / getUnitInfo とは用途が異なり、pt 換算は行わない）
           The same dropdown serves four preference keys, so unit code 5 uses the neutral "Q/H" label
           (unlike the UNITS / getUnitInfo tables elsewhere, this file does no pt conversion) */
        var unitOptions = ["pt","pc","in","mm","cm","Q/H","px"];

        // 現在の単位プリファレンスを取得
        var currentGeneral = getUnitKey("rulerType");
        var currentStroke  = getUnitKey("strokeUnits");
        var currentType    = getUnitKey("text/units");
        var currentAsian   = getUnitKey("text/asianunits");

        // ダイアログ作成
        var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dlg.alignChildren = "fill";

        // --- ダイアログ位置・透明度調整 ---
        var offsetX = 300;
        var dialogOpacity = 0.97;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function () {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }

        setDialogOpacity(dlg, dialogOpacity);
        shiftDialogPosition(dlg, offsetX, 0);

        // 左カラムラベル幅を約6文字分に短縮 (~60px)
        var labelWidthLeft = 80;
        var labelWidthRight = 120;

        // --- 単位プリセット選択用ラジオボタン ---
        // 「プリント（pt）」「プリント（Q）」「オンスクリーン（px）」の3種
        var modeGroup = dlg.add("group");
        modeGroup.orientation = "row";
        modeGroup.alignment = "center";
        var rbPt = modeGroup.add("radiobutton", undefined, getLabel("radio.modePt"));
        rbPt.helpTip = getLabel("tooltip.modePt");
        var rbQ  = modeGroup.add("radiobutton", undefined, getLabel("radio.modeQ"));
        rbQ.helpTip = getLabel("tooltip.modeQ");
        var rbPx = modeGroup.add("radiobutton", undefined, getLabel("radio.modePx"));
        rbPx.helpTip = getLabel("tooltip.modePx");
        rbPt.value = true;

        // プリント（pt）選択時の単位設定
        rbPt.onClick = function() {
            ddGeneral.selection = arrayIndexOf(unitOptions, "mm");
            ddStroke.selection  = arrayIndexOf(unitOptions, "pt");
            ddType.selection    = arrayIndexOf(unitOptions, "pt");
            ddAsian.selection   = arrayIndexOf(unitOptions, "pt");
            updateIncrementLabels();
        };

        // プリント（Q）選択時の単位設定
        rbQ.onClick = function() {
            ddGeneral.selection = arrayIndexOf(unitOptions, "mm");
            ddStroke.selection  = arrayIndexOf(unitOptions, "Q/H");
            ddType.selection    = arrayIndexOf(unitOptions, "Q/H");
            ddAsian.selection   = arrayIndexOf(unitOptions, "Q/H");
            updateIncrementLabels();
        };

        // オンスクリーン（px）選択時の単位設定
        rbPx.onClick = function() {
            ddGeneral.selection = arrayIndexOf(unitOptions, "px");
            ddStroke.selection  = arrayIndexOf(unitOptions, "pt");
            ddType.selection    = arrayIndexOf(unitOptions, "pt");
            ddAsian.selection   = arrayIndexOf(unitOptions, "px");
            updateIncrementLabels();
        };

        // --- ダイアログ本体（2カラム構成）---
        // 左：単位選択ドロップダウン
        // 右：数値入力フィールド
        var mainGroup = dlg.add("group");
        mainGroup.orientation = "row";

        // 左カラムをパネルに変更
        var leftPanel = mainGroup.add("panel", undefined, getLabel("panel.units"));
        leftPanel.orientation = "column";
        leftPanel.alignChildren = "left";

        // 右カラムをパネルに変更
        var rightPanel = mainGroup.add("panel", undefined, getLabel("panel.increments"));
        rightPanel.orientation = "column";
        rightPanel.alignChildren = "left";

        // 左カラムに単位選択4つ
        var grpGeneral = leftPanel.add("group");
        grpGeneral.orientation = "row";
        var lblGeneral = grpGeneral.add("statictext", undefined, labelText("fieldLabel.general"));
        lblGeneral.preferredSize.width = labelWidthLeft;
        lblGeneral.justify = "right";
        var ddGeneral = grpGeneral.add("dropdownlist", undefined, unitOptions);
        ddGeneral.helpTip = getLabel("tooltip.general");
        ddGeneral.selection = arrayIndexOf(unitOptions, currentGeneral);

        var grpStroke = leftPanel.add("group");
        grpStroke.orientation = "row";
        var lblStroke = grpStroke.add("statictext", undefined, labelText("fieldLabel.stroke"));
        lblStroke.preferredSize.width = labelWidthLeft;
        lblStroke.justify = "right";
        var ddStroke = grpStroke.add("dropdownlist", undefined, unitOptions);
        ddStroke.helpTip = getLabel("tooltip.stroke");
        ddStroke.selection = arrayIndexOf(unitOptions, currentStroke);

        var grpType = leftPanel.add("group");
        grpType.orientation = "row";
        var lblType = grpType.add("statictext", undefined, labelText("fieldLabel.type"));
        lblType.preferredSize.width = labelWidthLeft;
        lblType.justify = "right";
        var ddType = grpType.add("dropdownlist", undefined, unitOptions);
        ddType.helpTip = getLabel("tooltip.type");
        ddType.selection = arrayIndexOf(unitOptions, currentType);

        var grpAsian = leftPanel.add("group");
        grpAsian.orientation = "row";
        var lblAsian = grpAsian.add("statictext", undefined, labelText("fieldLabel.asian"));
        lblAsian.preferredSize.width = labelWidthLeft;
        lblAsian.justify = "right";
        var ddAsian = grpAsian.add("dropdownlist", undefined, unitOptions);
        ddAsian.helpTip = getLabel("tooltip.asian");
        ddAsian.selection = arrayIndexOf(unitOptions, currentAsian);

        // キー増加インクリメント
        var grpKey = rightPanel.add("group");
        grpKey.orientation = "row";
        var lblKey = grpKey.add("statictext", undefined, labelText("fieldLabel.key"));
        lblKey.preferredSize.width = labelWidthRight;
        lblKey.justify = "right";
        var etKeyValue = addIncrementInput(grpKey, "0.1");
        etKeyValue.helpTip = getLabel("tooltip.key");
        etKeyValue.characters = 5;
        var lblKeyUnit = grpKey.add("statictext", undefined, currentGeneral);
        lblKeyUnit.preferredSize.width = 40;

        // 角丸半径の増減値
        var grpRadius = rightPanel.add("group");
        grpRadius.orientation = "row";
        var lblRadius = grpRadius.add("statictext", undefined, labelText("fieldLabel.radius"));
        lblRadius.preferredSize.width = labelWidthRight;
        lblRadius.justify = "right";
        var etRadiusValue = addIncrementInput(grpRadius, "1");
        etRadiusValue.helpTip = getLabel("tooltip.radius");
        etRadiusValue.characters = 5;
        var lblRadiusUnit = grpRadius.add("statictext", undefined, currentGeneral);
        lblRadiusUnit.preferredSize.width = 40;

        // フォントサイズ増減値
        var grpSize = rightPanel.add("group");
        grpSize.orientation = "row";
        var lblSize = grpSize.add("statictext", undefined, labelText("fieldLabel.size"));
        lblSize.preferredSize.width = labelWidthRight;
        lblSize.justify = "right";
        var etSizeValue = addIncrementInput(grpSize, "1");
        etSizeValue.helpTip = getLabel("tooltip.size");
        etSizeValue.characters = 5;
        var lblSizeUnit = grpSize.add("statictext", undefined, currentType);
        lblSizeUnit.preferredSize.width = 40;

        // ベースラインシフト増減値
        var grpBaseline = rightPanel.add("group");
        grpBaseline.orientation = "row";
        var lblBaseline = grpBaseline.add("statictext", undefined, labelText("fieldLabel.baseline"));
        lblBaseline.preferredSize.width = labelWidthRight;
        lblBaseline.justify = "right";
        var etBaselineValue = addIncrementInput(grpBaseline, "0.1");
        etBaselineValue.helpTip = getLabel("tooltip.baseline");
        etBaselineValue.characters = 5;
        var lblBaselineUnit = grpBaseline.add("statictext", undefined, currentType);
        lblBaselineUnit.preferredSize.width = 40;

        // OK/キャンセル
        var btnGroup = dlg.add("group");
        btnGroup.alignment = "center";

        var btnCancel = btnGroup.add("button", undefined, "Cancel", {name: "cancel"});
        var btnOK = btnGroup.add("button", undefined, "OK", {name: "ok"});

        // ドロップダウン変更時に単位ラベルを更新
        ddGeneral.onChange = updateIncrementLabels;
        ddStroke.onChange = updateIncrementLabels;
        ddType.onChange = updateIncrementLabels;
        ddAsian.onChange = updateIncrementLabels;

        dlg.onShow = function() {
            updateIncrementLabels();
        };

        // 初期ラベル反映
        updateIncrementLabels();

        if (dlg.show() != 1) return; // Cancelなら中止

        // 値を取得
        var units_general = ddGeneral.selection.text;
        var units_stroke  = ddStroke.selection.text;
        var units_type    = ddType.selection.text;
        var units_asian   = ddAsian.selection.text;

        var num_key      = etKeyValue.text;
        var unit_key     = currentGeneral;
        var num_r        = etRadiusValue.text;
        var unit_r       = currentGeneral;
        var num_size     = etSizeValue.text;
        var unit_size    = currentType;
        var num_baseline = etBaselineValue.text;
        var unit_baseline = currentType;

        // 数値プリファレンスを設定
        setPref("cursorKeyLength", unit_key, num_key);
        setPref("ovalRadius", unit_r, num_r);
        setPref("text/sizeIncrement", unit_size, num_size);
        setPref("text/riseIncrement", unit_baseline, num_baseline);

        // 単位プリファレンスを設定（バージョンチェック不要のマッピング）
        setUnits("rulerType", units_general);
        setUnits("strokeUnits", units_stroke);
        setUnits("text/units", units_type);
        setUnits("text/asianunits", units_asian);

        // alert("単位と数値インクリメントを設定しました。");
    }

    /**
     * 単位をptに変換してプリファレンスに保存
     * 数値プリファレンスを設定
     */
    function setPref(prefKey, dstUnits, num) {
        var value = calcWithUnit(num, dstUnits);
        if (dstUnits !== "pt") {
            value = convertUnit(value, dstUnits, "pt");
        }
        if (value !== undefined) {
            app.preferences.setRealPreference(prefKey, value);
        }
    }

    /**
     * 単位プリファレンスを設定（バージョンごとのマッピング + フォールバック）
     */
    function setUnits(prefKey, dstUnits) {
        // バージョンごとのマッピング定義
        var DB_DEFAULT = {
            "pt": 2, "pc": 3, "in": 0, "mm": 1, "cm": 4, "Q/H": 5, "px": 6
        };
        var DB_ALT = {
            "pt": 3, "pc": 2, "in": 1, "mm": 0, "cm": 1, "Q/H": 6, "px": 5
        };

        var versionParts = app.version.split(".");
        var aiMajor = parseInt(versionParts[0], 10);
        var aiMinor = parseInt(versionParts[1] || "0", 10);

        // バージョン別に選択
        var db;
        if (aiMajor >= 25) {
            db = DB_DEFAULT;
        } else {
            db = DB_ALT;
        }

        if (db[dstUnits] !== undefined) {
            app.preferences.setIntegerPreference(prefKey, db[dstUnits]);
        } else {
            alert("未対応の単位: " + dstUnits + " (" + prefKey + ")。既定値 pt を使用します。");
            app.preferences.setIntegerPreference(prefKey, db["pt"]);
        }
    }

    /**
     * ユーザー入力を数値に変換（単位付き数値をサポート）
     * 単位付き文字列を計算
     */
    function calcWithUnit(str, defaultUnits) {
        if (!defaultUnits) defaultUnits = "mm";
        var newStr = str.replace(/[　\s]+/g, " ").replace(",", "");
        var regUnitValue = /([0-9]+(?:\.[0-9]+)?)( ?[a-zA-Z\/]+)/g;
        newStr = newStr.replace(regUnitValue, function (m0, m1, m2) {
            return convertUnit(m1, m2.replace(/^\s+|\s+$/g, ""), defaultUnits);
        });

        var result;
        try {
            result = eval(newStr);
        } catch (e) {
            result = undefined;
        }
        return result;
    }

    /**
     * Q/H（級／歯）をmm単位に変換し、UnitValueで指定単位に変換
     * 単位変換
     */
    function convertUnit(value, srcUnit, dstUnits) {
        var qReg = /^(?:Q\/H|[qh])$/i;
        var mmUnit = "mm";

        if (qReg.test(srcUnit)) {
            value = Number(value) * 0.25;
            srcUnit = mmUnit;
        }
        if (qReg.test(dstUnits)) {
            value = Number(value) * 4;
            dstUnits = mmUnit;
        }

        var res;
        try {
            res = new UnitValue(value, srcUnit).as(dstUnits);
        } catch (e) {
            res = undefined;
        }
        return res;
    }

    /**
     * 現在のプリファレンスIDを対応する単位キーに変換
     * 現在の単位プリファレンスのキーを取得
     */
    function getUnitKey(prefKey) {
        var id = app.preferences.getIntegerPreference(prefKey);
        var db = {
            0: "in",
            1: "mm",
            2: "pt",
            3: "pica",
            4: "cm",
            5: "Q/H",
            6: "px",
            7: "ft/in",
            8: "m",
            9: "yd",
            10: "ft"
        };
        return db[id] || "pt";
    }

    // ES3互換 indexOf
    function arrayIndexOf(arr, value) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] == value) return i;
        }
        return -1;
    }

    main();

})();
