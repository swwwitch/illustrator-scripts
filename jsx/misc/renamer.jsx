#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクトやレイヤーなどの名前を、条件を指定して一括で変更します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/renamer.md

### Overview

Renames objects, layers and the like in bulk, according to the conditions you set.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/renamer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "renamer";                      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/renamer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/renamer.md"; /* README (English) */

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
        var upTooltip = stepOptions.integer ? "tipStepUpInteger" : "tipStepUp";
        var downTooltip = stepOptions.integer ? "tipStepDownInteger" : "tipStepDown";
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
    // バージョンとローカライズ / Version and Localization
    // =========================================

    /* 現在のロケールを判定 / Detect current locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialogTitle:         { ja: "名前の検索置換", en: "Find and Replace Names" },
        noDoc:               { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        target:              { ja: "対象", en: "Target" },
        artboard:            { ja: "アートボード", en: "Artboard" },
        layer:               { ja: "レイヤー", en: "Layer" },
        symbol:              { ja: "シンボル", en: "Symbol" },
        graphicStyle:        { ja: "グラフィックスタイル", en: "Graphic Style" },
        findReplace:         { ja: "検索・置換", en: "Find & Replace" },
        findReplaceEnable:   { ja: "検索置換", en: "Find & Replace" },
        find:                { ja: "検索", en: "Find" },
        replace:             { ja: "置換", en: "Replace" },
        regex:               { ja: "正規表現", en: "Regex" },
        prefix:              { ja: "接頭辞", en: "Prefix" },
        suffix:              { ja: "接尾辞", en: "Suffix" },
        numberingEnable:     { ja: "ナンバリング", en: "Numbering" },
        separator:           { ja: "区切り", en: "Separator" },
        startNumber:         { ja: "開始番号", en: "Start" },
        sort:                { ja: "並び替え", en: "Sort" },
        sortOriginal:        { ja: "元の順", en: "Original" },
        sortNameAsc:         { ja: "名前 ↑", en: "Name ↑" },
        sortNameDesc:        { ja: "名前 ↓", en: "Name ↓" },
        sortChanged:         { ja: "変更あり優先", en: "Changed first" },
        moveTop:             { ja: "↑↑", en: "↑↑" },
        moveUp:              { ja: "↑", en: "↑" },
        moveDown:            { ja: "↓", en: "↓" },
        moveBottom:          { ja: "↓↓", en: "↓↓" },
        cancel:              { ja: "キャンセル", en: "Cancel" },
        needInput:           { ja: "検索文字を入力するか、接頭辞・接尾辞のナンバリングを有効にしてください。", en: "Enter a search string or enable prefix/suffix numbering." },
        noMatchArtboard:     { ja: "該当するアートボード名はありませんでした。", en: "No artboard names matched." },
        noMatchLayer:        { ja: "該当するレイヤー名はありませんでした。", en: "No layer names matched." },
        noMatchSymbol:       { ja: "該当するシンボル名はありませんでした。", en: "No symbol names matched." },
        noMatchGraphicStyle: { ja: "該当するグラフィックスタイル名はありませんでした。", en: "No graphic style names matched." },
        done:                { ja: "完了しました。", en: "Done." },
        targetArtboards:     { ja: "対象：アートボード名", en: "Target: Artboards" },
        targetLayers:        { ja: "対象：レイヤー名", en: "Target: Layers" },
        targetSymbols:       { ja: "対象：シンボル名", en: "Target: Symbols" },
        targetGraphicStyles: { ja: "対象：グラフィックスタイル名", en: "Target: Graphic Styles" },
        renamed:             { ja: "変更数：", en: "Renamed: " },
        suffixed:            { ja: "同名回避で連番追加：", en: "Suffixed to avoid duplicates: " },
        errorsLabel:         { ja: "エラー：", en: "Errors: " },
        countSuffix:         { ja: " 件", en: "" },
        tipArtboard:         { ja: "アートボード名を対象にします。", en: "Renames artboard names." },
        tipLayer:            { ja: "レイヤー名を対象にします。", en: "Renames layer names." },
        tipSymbol:           { ja: "シンボル名を対象にします。", en: "Renames symbol names." },
        tipGraphicStyle:     { ja: "グラフィックスタイル名を対象にします。", en: "Renames graphic style names." },
        tipFindReplaceEnable:{ ja: "オフにすると、接頭辞・接尾辞のナンバリングだけを行います。", en: "When off, only the prefix/suffix numbering is applied." },
        tipFind:             { ja: "名前の中から探す文字列です。空欄だとすべてが対象になります。", en: "Text to look for in the name. Leave blank to target everything." },
        tipReplace:          { ja: "置き換える文字列です。空欄にすると検索文字列を削除します。", en: "Replacement text. Leave blank to delete the found text." },
        tipRegex:            { ja: "検索文字列を正規表現として扱います（$1 などの後方参照も使えます）。", en: "Treats the search text as a regular expression, including back-references such as $1." },
        tipPrefixEnable:     { ja: "名前の先頭に連番を付けます。", en: "Adds a sequential number to the start of the name." },
        tipSuffixEnable:     { ja: "名前の末尾に連番を付けます。", en: "Adds a sequential number to the end of the name." },
        tipSeparator:        { ja: "連番と名前の間に入れる文字です。", en: "Character placed between the number and the name." },
        tipStartNumber:      { ja: "連番の最初の数字です。", en: "First number of the sequence." },
        tipStepUp:           { ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）", en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)" },
        tipStepDown:         { ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）", en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)" },
        tipStepUpInteger:    { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        tipStepDownInteger:  { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
        tipSort:             { ja: "プレビューと連番を振る順序です。", en: "Order used for the preview and for numbering." },
        tipMoveTop:          { ja: "選んだ項目を先頭へ移動します。", en: "Moves the selected item to the top." },
        tipMoveUp:           { ja: "選んだ項目を1つ上へ移動します。", en: "Moves the selected item up one position." },
        tipMoveDown:         { ja: "選んだ項目を1つ下へ移動します。", en: "Moves the selected item down one position." },
        tipMoveBottom:       { ja: "選んだ項目を末尾へ移動します。", en: "Moves the selected item to the bottom." },
        tipPreview:          { ja: "変更前と変更後の名前の一覧です。並べ替えたい項目はここで選びます。", en: "List of the names before and after the change. Select an item here to reorder it." }
    };

    /* ローカライズ文字列を取得 / Get localized string */
    function getLabel(key) {
        var entry = LABELS[key];
        if (!entry) return key;
        return entry[uiLang] || entry.en || entry.ja || key;
    }

    /* コロン付きの項目名を返す（日本語は全角、英語は半角） / Return a label with a colon */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ": ");
    }

        if (app.documents.length === 0) {
            alert(getLabel("noDoc"));
            return;
        }

        var doc = app.activeDocument;

        // =========================================
        // ダイアログ定数 / ヘルパー / Dialog constants & helpers
        // =========================================

        var PANEL_MARGINS = [15, 20, 15, 10];
        var PANEL_SPACING = 8;
        var PREVIEW_LINE_HEIGHT = 16; // Mac の listbox 行高さ目安 / Approx. line height on Mac
        var PREVIEW_VISIBLE_LINES = 20;

        /* パネルの共通設定を適用 / Apply common panel settings */
        function setupPanel(panel, spacing) {
            panel.orientation = "column";
            panel.alignChildren = "left";
            panel.alignment = "fill";
            panel.margins = PANEL_MARGINS;
            panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        }

        // =========================================
        // 共通関数 / Common functions
        // =========================================

        /* 全置換（非正規表現） / Replace all (non-regex) */
        function replaceAll(text, search, replacement) {
            return text.split(search).join(replacement);
        }

        /* Object.prototype 衝突回避のキー / Key to avoid Object.prototype collisions */
        function nameKey(name) {
            return "@" + name;
        }

        /* 衝突しないユニーク名を生成 / Generate unique name avoiding collisions */
        function makeUniqueName(baseName, usedNames) {
            var name = baseName;
            var suffix = 2;

            while (usedNames[nameKey(name)]) {
                name = baseName + "_" + suffix;
                suffix++;
            }

            return name;
        }

        /* 名前取得の例外を握りつぶす / Safely get item name */
        function safeGetName(item) {
            try { return item.name; } catch (e) { return ""; }
        }

        /* 検索文字 / 正規表現でマッチ判定 / Test match by string or regex */
        function matchesFind(text, search, useRegex) {
            if (search === "") return false;
            if (useRegex) {
                try { return new RegExp(search).test(text); }
                catch (e) { return false; }
            }
            return text.indexOf(search) >= 0;
        }

        /* 置換を実行 / Apply replacement */
        function applyReplace(text, search, replacement, useRegex) {
            if (useRegex) {
                try { return text.replace(new RegExp(search, "g"), replacement); }
                catch (e) { return text; }
            }
            return replaceAll(text, search, replacement);
        }

        /* 0埋め / Zero-pad number */
        function padLeftZero(num, width) {
            var s = "" + num;
            while (s.length < width) s = "0" + s;
            return s;
        }

        /**
         * 開始番号の欄を「∧∨・入力欄」の組で追加する。↑↓キーも∧∨と同じ処理で増減する。
         * 「001」のように0埋めした桁数は連番の桁数になるので、増減後も手入力したときの桁数で0埋めし直す
         * @param {Group} parentRow - 追加先の行
         * @returns {EditText} 入力欄（桁数は .zeroPadWidth に控える）
         */
        function addStartNumberInput(parentRow) {
            var stepperInputGroup = parentRow.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;

            var startInput;
            var stepperGroup = addStepper(stepperInputGroup, function () { return startInput; }, {
                integer: true,
                min: 0,
                onStep: function (numberInput) {
                    /* 増減で「001」→「2」と桁が落ちないよう、手入力の桁数で0埋めし直す / keep the typed zero-padding */
                    numberInput.text = padLeftZero(parseInt(numberInput.text, 10), numberInput.zeroPadWidth);
                    numberInput.lastValidText = numberInput.text;
                    rerender();
                }
            });
            startInput = stepperInputGroup.add("edittext", undefined, "1");
            startInput.zeroPadWidth = 1;
            startInput.stepperGroup = stepperGroup;
            bindSteppedArrowKeys(startInput, stepperGroup);
            return startInput;
        }

        /**
         * 開始番号の手入力時に、0埋めの桁数を控えてプレビューを更新する（onChanging に割り当てる）
         * @returns {void}
         */
        function onStartNumberChanging() {
            this.zeroPadWidth = this.text.length > 0 ? this.text.length : 1;
            rerender();
        }

        /* 入力文字列から開始番号と桁数を取得 / Parse start number and width */
        function parseStartSpec(s) {
            var n = parseInt(s, 10);
            if (isNaN(n)) n = 1;
            return { start: n, width: s.length > 0 ? s.length : 1 };
        }

        /* ネストレイヤーを再帰的に収集 / Recursively collect nested layers */
        function collectLayers(layers, collected) {
            for (var i = 0; i < layers.length; i++) {
                var layer = layers[i];

                collected.push(layer);

                if (layer.layers && layer.layers.length > 0) {
                    collectLayers(layer.layers, collected);
                }
            }
        }

        /* 対象モードに応じた項目配列を返す / Return items for the given mode */
        function getTargetItems(mode) {
            var items = [];
            var i;
            if (mode === "artboard") {
                for (i = 0; i < doc.artboards.length; i++) items.push(doc.artboards[i]);
            } else if (mode === "layer") {
                collectLayers(doc.layers, items);
            } else if (mode === "symbol") {
                for (i = 0; i < doc.symbols.length; i++) items.push(doc.symbols[i]);
            } else if (mode === "graphicStyle") {
                for (i = 0; i < doc.graphicStyles.length; i++) items.push(doc.graphicStyles[i]);
            }
            return items;
        }

        /* ラジオボタンから対象モードを取得 / Get current target mode from radios */
        function getModeFromRadios() {
            if (rbArtboard.value) return "artboard";
            if (rbLayer.value) return "layer";
            if (rbSymbol.value) return "symbol";
            if (rbGraphicStyle.value) return "graphicStyle";
            return "symbol";
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        var win = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];

        // 対象選択 (2カラムを貫通) / Target selection (spans both columns)
        var targetPanel = win.add("panel", undefined, getLabel("target"));
        targetPanel.orientation = "row";
        targetPanel.alignChildren = ["left", "center"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = 10;

        var rbArtboard = targetPanel.add("radiobutton", undefined, getLabel("artboard"));
        rbArtboard.helpTip = getLabel("tipArtboard");
        var rbLayer = targetPanel.add("radiobutton", undefined, getLabel("layer"));
        rbLayer.helpTip = getLabel("tipLayer");
        var rbSymbol = targetPanel.add("radiobutton", undefined, getLabel("symbol"));
        rbSymbol.helpTip = getLabel("tipSymbol");
        var rbGraphicStyle = targetPanel.add("radiobutton", undefined, getLabel("graphicStyle"));
        rbGraphicStyle.helpTip = getLabel("tipGraphicStyle");

        rbSymbol.value = true;

        // 2カラム / Two-column layout
        var mainGroup = win.add("group");
        mainGroup.orientation = "row";
        mainGroup.alignChildren = ["fill", "fill"];

        // 左カラム / Left column
        var leftColumn = mainGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];

        // 検索・置換パネル / Find & Replace panel
        var findReplacePanel = leftColumn.add("panel", undefined, getLabel("findReplace"));
        setupPanel(findReplacePanel, 6);

        var findReplaceCheckbox = findReplacePanel.add("checkbox", undefined, getLabel("findReplaceEnable"));
        findReplaceCheckbox.helpTip = getLabel("tipFindReplaceEnable");
        findReplaceCheckbox.value = true;

        var findGroup = findReplacePanel.add("group");
        findGroup.orientation = "row";
        findGroup.alignChildren = ["left", "center"];
        findGroup.add("statictext", undefined, labelText("find"));
        var findInput = findGroup.add("edittext", undefined, "");
        findInput.helpTip = getLabel("tipFind");
        findInput.characters = 15;

        var replaceGroup = findReplacePanel.add("group");
        replaceGroup.orientation = "row";
        replaceGroup.alignChildren = ["left", "center"];
        replaceGroup.add("statictext", undefined, labelText("replace"));
        var replaceInput = replaceGroup.add("edittext", undefined, "");
        replaceInput.helpTip = getLabel("tipReplace");
        replaceInput.characters = 15;

        var regexGroup = findReplacePanel.add("group");
        regexGroup.orientation = "row";
        regexGroup.alignChildren = ["right", "center"];
        regexGroup.alignment = "fill";
        var regexCheckbox = regexGroup.add("checkbox", undefined, getLabel("regex"));
        regexCheckbox.helpTip = getLabel("tipRegex");

        // 接頭辞パネル / Prefix panel
        var prefixPanel = leftColumn.add("panel", undefined, getLabel("prefix"));
        setupPanel(prefixPanel, 6);

        var prefixCheckbox = prefixPanel.add("checkbox", undefined, getLabel("numberingEnable"));
        prefixCheckbox.helpTip = getLabel("tipPrefixEnable");

        var prefixSepGroup = prefixPanel.add("group");
        prefixSepGroup.orientation = "row";
        prefixSepGroup.alignChildren = ["left", "center"];
        prefixSepGroup.add("statictext", undefined, labelText("separator"));
        var rbPrefixSepDash = prefixSepGroup.add("radiobutton", undefined, "-");
        rbPrefixSepDash.helpTip = getLabel("tipSeparator");
        var rbPrefixSepUnderscore = prefixSepGroup.add("radiobutton", undefined, "_");
        rbPrefixSepUnderscore.helpTip = getLabel("tipSeparator");
        rbPrefixSepDash.value = true;

        var prefixStartGroup = prefixPanel.add("group");
        prefixStartGroup.orientation = "row";
        prefixStartGroup.alignChildren = ["left", "center"];
        prefixStartGroup.add("statictext", undefined, labelText("startNumber"));
        var prefixStartInput = addStartNumberInput(prefixStartGroup);
        prefixStartInput.helpTip = getLabel("tipStartNumber");
        prefixStartInput.characters = 4;

        // 接尾辞パネル / Suffix panel
        var suffixPanel = leftColumn.add("panel", undefined, getLabel("suffix"));
        setupPanel(suffixPanel, 6);

        var suffixCheckbox = suffixPanel.add("checkbox", undefined, getLabel("numberingEnable"));
        suffixCheckbox.helpTip = getLabel("tipSuffixEnable");

        var suffixSepGroup = suffixPanel.add("group");
        suffixSepGroup.orientation = "row";
        suffixSepGroup.alignChildren = ["left", "center"];
        suffixSepGroup.add("statictext", undefined, labelText("separator"));
        var rbSuffixSepDash = suffixSepGroup.add("radiobutton", undefined, "-");
        rbSuffixSepDash.helpTip = getLabel("tipSeparator");
        var rbSuffixSepUnderscore = suffixSepGroup.add("radiobutton", undefined, "_");
        rbSuffixSepUnderscore.helpTip = getLabel("tipSeparator");
        rbSuffixSepDash.value = true;

        var suffixStartGroup = suffixPanel.add("group");
        suffixStartGroup.orientation = "row";
        suffixStartGroup.alignChildren = ["left", "center"];
        suffixStartGroup.add("statictext", undefined, labelText("startNumber"));
        var suffixStartInput = addStartNumberInput(suffixStartGroup);
        suffixStartInput.helpTip = getLabel("tipStartNumber");
        suffixStartInput.characters = 4;

        // 右カラム: 並び替え + プレビュー / Right column: sort + preview
        var rightColumn = mainGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "fill"];

        var sortGroup = rightColumn.add("group");
        sortGroup.orientation = "row";
        sortGroup.alignChildren = ["left", "center"];
        sortGroup.add("statictext", undefined, labelText("sort"));
        var sortDropdown = sortGroup.add("dropdownlist", undefined, [
            getLabel("sortOriginal"),
            getLabel("sortNameAsc"),
            getLabel("sortNameDesc"),
            getLabel("sortChanged")
        ]);
        sortDropdown.selection = 0;
        sortDropdown.helpTip = getLabel("tipSort");

        var moveTopBtn = sortGroup.add("button", undefined, getLabel("moveTop"));
        moveTopBtn.helpTip = getLabel("tipMoveTop");
        var moveUpBtn = sortGroup.add("button", undefined, getLabel("moveUp"));
        moveUpBtn.helpTip = getLabel("tipMoveUp");
        var moveDownBtn = sortGroup.add("button", undefined, getLabel("moveDown"));
        moveDownBtn.helpTip = getLabel("tipMoveDown");
        var moveBottomBtn = sortGroup.add("button", undefined, getLabel("moveBottom"));
        moveBottomBtn.helpTip = getLabel("tipMoveBottom");
        moveTopBtn.preferredSize.width = 36;
        moveUpBtn.preferredSize.width = 32;
        moveDownBtn.preferredSize.width = 32;
        moveBottomBtn.preferredSize.width = 36;

        var previewList = rightColumn.add("listbox", undefined, []);
        previewList.helpTip = getLabel("tipPreview");
        previewList.preferredSize.width = 340;
        previewList.preferredSize.height = PREVIEW_VISIBLE_LINES * PREVIEW_LINE_HEIGHT;

        // ボタン (Mac規約: Cancel → OK) / Buttons (Mac convention)
        var buttonGroup = win.add("group");
        buttonGroup.alignment = "right";

        buttonGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
        var okBtn = buttonGroup.add("button", undefined, "OK", { name: "ok" });

        okBtn.enabled = false;

        // =========================================
        // プレビュー / ナンバリング / Preview & numbering
        // =========================================

        var lastItems = [];

        /* 区切り文字を取得 (接頭辞/接尾辞) / Get separator for prefix/suffix */
        function getPrefixSeparator() {
            return rbPrefixSepDash.value ? "-" : "_";
        }
        function getSuffixSeparator() {
            return rbSuffixSepDash.value ? "-" : "_";
        }

        /* 新名前を全件再計算 / Recompute all new names */
        function recomputeNewNames() {
            var prefixOn = prefixCheckbox.value;
            var suffixOn = suffixCheckbox.value;

            var prefixSep = getPrefixSeparator();
            var prefixSpec = parseStartSpec(prefixStartInput.text);
            var prefixCounter = prefixSpec.start;

            var suffixSep = getSuffixSeparator();
            var suffixSpec = parseStartSpec(suffixStartInput.text);
            var suffixCounter = suffixSpec.start;

            for (var i = 0; i < lastItems.length; i++) {
                var e = lastItems[i];
                var base = e.matched ? e.replacedName : e.oldName;
                var finalName = base;

                if (prefixOn) {
                    finalName = padLeftZero(prefixCounter, prefixSpec.width) + prefixSep + finalName;
                    prefixCounter++;
                }
                if (suffixOn) {
                    finalName = finalName + suffixSep + padLeftZero(suffixCounter, suffixSpec.width);
                    suffixCounter++;
                }

                e.newName = finalName;
            }
        }

        /* 並び替えを適用 / Apply sort */
        function applySort() {
            var mode = sortDropdown.selection ? sortDropdown.selection.index : 0;
            if (mode === 1) {
                lastItems.sort(function (a, b) {
                    return a.oldName < b.oldName ? -1 : (a.oldName > b.oldName ? 1 : 0);
                });
            } else if (mode === 2) {
                lastItems.sort(function (a, b) {
                    return a.oldName < b.oldName ? 1 : (a.oldName > b.oldName ? -1 : 0);
                });
            } else if (mode === 3) {
                lastItems.sort(function (a, b) {
                    return (b.matched ? 1 : 0) - (a.matched ? 1 : 0);
                });
            }
        }

        /* プレビュー (listbox) を描画 / Render preview list */
        function renderPreview(keepSelectionIndex) {
            previewList.removeAll();
            for (var i = 0; i < lastItems.length; i++) {
                var e = lastItems[i];
                var text;
                if (e.newName !== e.oldName) {
                    text = e.oldName + "  →  " + e.newName;
                } else {
                    text = e.oldName;
                }
                previewList.add("item", text);
            }
            if (typeof keepSelectionIndex === "number" && keepSelectionIndex >= 0 && keepSelectionIndex < previewList.items.length) {
                previewList.selection = keepSelectionIndex;
            }
        }

        /* 対象を再走査してプレビューを更新 / Rescan items and refresh preview */
        function refreshPreview() {
            var mode = getModeFromRadios();
            var items = getTargetItems(mode);
            var findReplaceOn = findReplaceCheckbox.value;
            var search = findInput.text;
            var replacement = replaceInput.text;
            var useRegexNow = regexCheckbox.value;

            lastItems = [];
            for (var i = 0; i < items.length; i++) {
                var oldName = safeGetName(items[i]);
                if (oldName === "") continue;

                var replaced = oldName;
                var matched = false;
                if (findReplaceOn && search !== "" && matchesFind(oldName, search, useRegexNow)) {
                    var candidate = applyReplace(oldName, search, replacement, useRegexNow);
                    if (candidate !== oldName) {
                        replaced = candidate;
                        matched = true;
                    }
                }

                lastItems.push({
                    item: items[i],
                    oldName: oldName,
                    replacedName: replaced,
                    matched: matched,
                    newName: replaced
                });
            }

            applySort();
            recomputeNewNames();
            renderPreview();
        }

        /* 再計算と再描画のみ実行 / Recompute & re-render without rescan */
        function rerender() {
            recomputeNewNames();
            var selIdx = previewList.selection ? previewList.selection.index : -1;
            renderPreview(selIdx);
        }

        /* 接頭辞/接尾辞 UI の有効/無効を切替 / Toggle prefix/suffix UI */
        function updatePrefixEnabled() {
            var on = prefixCheckbox.value;
            rbPrefixSepDash.enabled = on;
            rbPrefixSepUnderscore.enabled = on;
            prefixStartInput.enabled = on;
            /* ∧∨は入力欄の兄弟なので、包む group ごと切り替えて描き直す / the stepper is a sibling, so toggle and redraw the wrapper */
            prefixStartInput.parent.enabled = on;
            redrawSteppersIn(prefixStartInput.parent);
        }
        function updateSuffixEnabled() {
            var on = suffixCheckbox.value;
            rbSuffixSepDash.enabled = on;
            rbSuffixSepUnderscore.enabled = on;
            suffixStartInput.enabled = on;
            /* ∧∨は入力欄の兄弟なので、包む group ごと切り替えて描き直す / the stepper is a sibling, so toggle and redraw the wrapper */
            suffixStartInput.parent.enabled = on;
            redrawSteppersIn(suffixStartInput.parent);
        }

        /* OK ボタンの有効/無効を切替 / Toggle OK button */
        function updateOkEnabled() {
            var findReplaceActive = findReplaceCheckbox.value && findInput.text.length > 0;
            okBtn.enabled = findReplaceActive || prefixCheckbox.value || suffixCheckbox.value;
        }

        /* 選択行を上下に動かす / Move selected row up or down */
        function moveSelected(direction) {
            var currentSelection = previewList.selection;
            if (!currentSelection) return;
            var idx = currentSelection.index;
            var newIdx = idx + direction;
            if (newIdx < 0 || newIdx >= lastItems.length) return;

            var tmp = lastItems[idx];
            lastItems[idx] = lastItems[newIdx];
            lastItems[newIdx] = tmp;

            recomputeNewNames();
            renderPreview(newIdx);
        }

        sortDropdown.onChange = function () {
            applySort();
            recomputeNewNames();
            renderPreview();
        };

        moveUpBtn.onClick = function () { moveSelected(-1); };
        moveDownBtn.onClick = function () { moveSelected(1); };

        rbArtboard.onClick = refreshPreview;
        rbLayer.onClick = refreshPreview;
        rbSymbol.onClick = refreshPreview;
        rbGraphicStyle.onClick = refreshPreview;

        findInput.onChanging = function () {
            updateOkEnabled();
            refreshPreview();
        };

        replaceInput.onChanging = refreshPreview;
        regexCheckbox.onClick = refreshPreview;

        prefixCheckbox.onClick = function () {
            updatePrefixEnabled();
            updateOkEnabled();
            rerender();
        };
        rbPrefixSepDash.onClick = rerender;
        rbPrefixSepUnderscore.onClick = rerender;
        prefixStartInput.onChanging = onStartNumberChanging;

        suffixCheckbox.onClick = function () {
            updateSuffixEnabled();
            updateOkEnabled();
            rerender();
        };
        rbSuffixSepDash.onClick = rerender;
        rbSuffixSepUnderscore.onClick = rerender;
        suffixStartInput.onChanging = onStartNumberChanging;

        updatePrefixEnabled();
        updateSuffixEnabled();
        findInput.active = true;
        refreshPreview();

        if (win.show() !== 1) {
            return;
        }

        var findText = findInput.text;
        var enablePrefix = prefixCheckbox.value;
        var enableSuffix = suffixCheckbox.value;

        if (findText === "" && !enablePrefix && !enableSuffix) {
            alert(getLabel("needInput"));
            return;
        }

        var targetMode = getModeFromRadios();

        // リネーム計画 (実際に名前が変わる項目のみ) / Rename plan (only items whose name changes)
        var renamePlan = [];
        for (var rp = 0; rp < lastItems.length; rp++) {
            var entry = lastItems[rp];
            if (entry.newName !== entry.oldName) {
                renamePlan.push({
                    item: entry.item,
                    oldName: entry.oldName,
                    newName: entry.newName
                });
            }
        }

        /* リネーム計画から該当アイテムを検索 / Find planned rename for an item */
        function plannedRenameFor(refItem) {
            for (var i = 0; i < renamePlan.length; i++) {
                if (renamePlan[i].item === refItem) return renamePlan[i];
            }
            return null;
        }

        // =========================================
        // アートボード名 / Artboard names
        // =========================================

        /* アートボード名をリネーム / Rename artboards */
        function renameArtboards() {
            if (renamePlan.length === 0) {
                alert(getLabel("noMatchArtboard"));
                return;
            }

            var count = 0;
            var errors = 0;

            for (var i = 0; i < renamePlan.length; i++) {
                try {
                    renamePlan[i].item.name = renamePlan[i].newName;
                    count++;
                } catch (e) {
                    errors++;
                }
            }

            alert(
                getLabel("done") + "\n\n" +
                getLabel("targetArtboards") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
            );
        }

        // =========================================
        // レイヤー名 / Layer names
        // =========================================

        /* レイヤー名をリネーム / Rename layers */
        function renameLayers() {
            if (renamePlan.length === 0) {
                alert(getLabel("noMatchLayer"));
                return;
            }

            var count = 0;
            var errors = 0;

            for (var i = 0; i < renamePlan.length; i++) {
                try {
                    renamePlan[i].item.name = renamePlan[i].newName;
                    count++;
                } catch (e) {
                    errors++;
                }
            }

            alert(
                getLabel("done") + "\n\n" +
                getLabel("targetLayers") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
            );
        }

        // =========================================
        // シンボル名 / Symbol names
        // =========================================
        // シンボル名は同名不可のため、一度一時名にしてから最終名に変更します。
        // 同名になる場合は _2, _3 のように付けます。
        // Symbol names must be unique, so we rename via temporary names first.
        // Duplicates are suffixed as _2, _3, ...

        /* シンボル名をリネーム / Rename symbols */
        function renameSymbols() {
            if (renamePlan.length === 0) {
                alert(getLabel("noMatchSymbol"));
                return;
            }

            var symbols = doc.symbols;
            var usedNames = {};

            // リネーム対象外のシンボル名を予約 / Reserve names of symbols not being renamed
            for (var i = 0; i < symbols.length; i++) {
                if (!plannedRenameFor(symbols[i])) {
                    usedNames[nameKey(symbols[i].name)] = true;
                }
            }

            var count = 0;
            var suffixed = 0;
            var errors = 0;

            // 一時名へ変更 / Rename to temporary names
            var tempPrefix = "__symbol_rename_temp__" + new Date().getTime() + "__";
            var renamedToTemp = [];

            for (var j = 0; j < renamePlan.length; j++) {
                try {
                    var tempName = tempPrefix + j;
                    renamePlan[j].item.name = tempName;
                    renamedToTemp.push(renamePlan[j]);
                } catch (e1) {
                    errors++;
                }
            }

            // 最終名へ変更 / Rename to final names
            for (var k = 0; k < renamedToTemp.length; k++) {
                try {
                    var desiredName = renamedToTemp[k].newName;
                    var finalName = makeUniqueName(desiredName, usedNames);

                    if (finalName !== desiredName) {
                        suffixed++;
                    }

                    renamedToTemp[k].item.name = finalName;
                    usedNames[nameKey(finalName)] = true;
                    count++;
                } catch (e2) {
                    errors++;
                }
            }

            alert(
                getLabel("done") + "\n\n" +
                getLabel("targetSymbols") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("suffixed") + suffixed + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
            );
        }

        // =========================================
        // グラフィックスタイル名 / Graphic style names
        // =========================================
        // グラフィックスタイル名は同名不可のため、一度一時名にしてから最終名に変更します。
        // 初期スタイルなど、一部のスタイルは変更できない場合があります。
        // 同名になる場合は _2, _3 のように付けます。
        // Graphic style names must be unique, so we rename via temporary names first.
        // Some built-in styles may be immutable. Duplicates are suffixed as _2, _3, ...

        /* グラフィックスタイル名をリネーム / Rename graphic styles */
        function renameGraphicStyles() {
            if (renamePlan.length === 0) {
                alert(getLabel("noMatchGraphicStyle"));
                return;
            }

            var styles = doc.graphicStyles;
            var usedNames = {};

            // リネーム対象外のスタイル名を予約 / Reserve names of styles not being renamed
            for (var i = 0; i < styles.length; i++) {
                if (!plannedRenameFor(styles[i])) {
                    try { usedNames[nameKey(styles[i].name)] = true; }
                    catch (e0) { /* 名前取得できないスタイルは無視 / Skip styles whose name is unavailable */ }
                }
            }

            var count = 0;
            var suffixed = 0;
            var errors = 0;

            // 一時名へ変更 / Rename to temporary names
            var tempPrefix = "__gstyle_rename_temp__" + new Date().getTime() + "__";
            var renamedToTemp = [];

            for (var j = 0; j < renamePlan.length; j++) {
                try {
                    var tempName = tempPrefix + j;
                    renamePlan[j].item.name = tempName;
                    renamedToTemp.push(renamePlan[j]);
                } catch (e1) {
                    errors++;
                }
            }

            // 最終名へ変更 / Rename to final names
            for (var k = 0; k < renamedToTemp.length; k++) {
                try {
                    var desiredName = renamedToTemp[k].newName;
                    var finalName = makeUniqueName(desiredName, usedNames);

                    if (finalName !== desiredName) {
                        suffixed++;
                    }

                    renamedToTemp[k].item.name = finalName;
                    usedNames[nameKey(finalName)] = true;
                    count++;
                } catch (e2) {
                    errors++;
                }
            }

            alert(
                getLabel("done") + "\n\n" +
                getLabel("targetGraphicStyles") + "\n" +
                getLabel("renamed") + count + getLabel("countSuffix") + "\n" +
                getLabel("suffixed") + suffixed + getLabel("countSuffix") + "\n" +
                getLabel("errorsLabel") + errors + getLabel("countSuffix")
            );
        }

        // =========================================
        // 実行 / Execute
        // =========================================

        if (targetMode === "artboard") {
            renameArtboards();
        } else if (targetMode === "layer") {
            renameLayers();
        } else if (targetMode === "symbol") {
            renameSymbols();
        } else if (targetMode === "graphicStyle") {
            renameGraphicStyles();
        }

    })();
