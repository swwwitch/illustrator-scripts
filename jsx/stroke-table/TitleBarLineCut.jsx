#target illustrator
#targetengine "TitleBarLineCutEngine"
try { app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); } catch (e) { }

/*

### 概要

テキスト1つと長方形パス1つを選択して実行すると、テキストの周りで罫線が欠けたタイトル帯を作成します。
マージン、角丸、塗り、ノッチ、線幅をダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TitleBarLineCut.md

### Overview

With one text frame and one rectangle path selected, builds a title bar whose rule is cut away around the text.
Margin, corner radius, fill, notch and stroke weight are set in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TitleBarLineCut.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TitleBarLineCut";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TitleBarLineCut.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TitleBarLineCut.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    (function initSessionState() {
        if (!$.global.__tblc_state) {
            $.global.__tblc_state = {
                marginText: null,
                roundOn: false,
                roundText: null,
                fillOn: true,
                notchOn: false,
                strokeOn: true,
                widthText: null,
                capIndex: 0 // 0:none 1:round
            };
        }
    })();

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

    /**
     * 行に「∧∨＋入力欄」を追加する。∧∨・↑↓キーで増減したら onChanging を呼んで既存のプレビュー更新に乗せる
     * @param {Group} rowGroup - 追加先の行
     * @param {string} initialText - 初期値
     * @param {Object} stepOptions - min / max / integer（addStepper() に渡す）
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSteppedInput(rowGroup, initialText, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = rowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        stepOptions.onStep = function (steppedInput) {
            try {
                if (typeof steppedInput.onChanging === "function") steppedInput.onChanging();
            } catch (e) { }
        };
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * addSteppedInput() で作った入力欄の有効・無効を、∧∨のディム表示とあわせて切り替える
     * @param {EditText} numberInput - 対象の入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

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

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "タイトル帯の罫線欠け処理",
            en: "Title Bar Line Cut"
        },

        // Alerts
        alertOpenDoc: {
            ja: "ドキュメントを開いてください。",
            en: "Please open a document."
        },
        alertSelectTwo: {
            ja: "テキスト1つと長方形パス1つを選択してください。",
            en: "Select one text object and one rectangle path."
        },
        alertSelectTypes: {
            ja: "テキスト（1つ）と、長方形パス（1つ）または欠け罫線（1つ）を選択してください。",
            en: "Select 1 text object and 1 rectangle path, or 1 already-cut stroke."
        },
        alertMarginNonNegative: {
            ja: "マージン(%u)は0以上の数値で指定してください。",
            en: "Margin (%u) must be 0 or greater."
        },

        // Buttons
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },

        // Panels
        panelFill: { ja: "塗り", en: "Fill" },
        panelStroke: { ja: "線", en: "Stroke" },

        // Labels / Controls
        margin: { ja: "マージン", en: "Margin" },
        tipMargin: { ja: "文字の外側に足す余白です。", en: "Space added around the text." },
        tipRound: { ja: "背景の角を丸めます。半径は右の欄で指定します。", en: "Rounds the corners of the backing shape. The field on the right sets the radius." },
        tipFillOn: { ja: "背景に塗りを付けます。", en: "Fills the backing shape." },
        tipNotch: { ja: "背景の端に切り欠きを入れます。", en: "Cuts a notch into the edge of the backing shape." },
        tipStrokeOn: { ja: "背景に線を付けます。太さは右の欄で指定します。", en: "Strokes the backing shape. The field on the right sets the weight." },
        tipCapNone: { ja: "線の端に飾りを付けません。", en: "Leaves the line ends plain." },
        tipCapRound: { ja: "線の端を丸くします。", en: "Rounds the line ends." },
        roundCorners: { ja: "角丸", en: "Round" },
        fillOn: { ja: "塗り", en: "Fill" },
        notch: { ja: "ノッチ", en: "Notch" },
        strokeWidth: { ja: "線幅", en: "Stroke" },
        lineCap: { ja: "線端", en: "Cap" },
        capNone: { ja: "なし", en: "None" },
        capRound: { ja: "丸型", en: "Round" },

        // Stepper buttons
        tipStepUp: {
            ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        tipStepDown: {
            ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        tipStepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        tipStepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
    };

    function getLabel(key) {
        try {
            var v = LABELS[key];
            if (!v) return key;
            return v[uiLang] || v.en || v.ja || key;
        } catch (e) {
            return key;
        }
    }

    function LF(key, unitLabel) {
        // %u を単位ラベルで置換
        var s = getLabel(key);
        return s.replace(/%u/g, String(unitLabel));
    }

    (function () {
        if (app.documents.length === 0) { alert(getLabel('alertOpenDoc')); return; }
        var doc = app.activeDocument;

        if (!doc.selection || doc.selection.length !== 2) {
            alert(getLabel('alertSelectTwo'));
            return;
        }

        function isTextItem(it) { return it && it.typename === "TextFrame"; }

        function isPathItem(it) {
            return it && it.typename === "PathItem";
        }

        function isGroupItem(it) {
            return it && it.typename === "GroupItem";
        }

        function isStrokeCandidate(it) {
            // Rectangle path, already-cut open path, or a group of open paths (2-edge mode)
            return isPathItem(it) || isGroupItem(it);
        }

        function isRectPathItem(it) {
            // Strict rectangle: closed + 4 points
            return isPathItem(it) && it.closed === true && it.pathPoints.length === 4;
        }

        function extractRectBoundsFromPathItem(p) {
            // Use anchors (not visibleBounds) to avoid stroke width affecting bounds.
            // Returns [L,T,R,B]
            var L = 1e10, T = -1e10, R = -1e10, B = 1e10;
            try {
                for (var i = 0; i < p.pathPoints.length; i++) {
                    var a = p.pathPoints[i].anchor;
                    var x = a[0], y = a[1];
                    if (x < L) L = x;
                    if (x > R) R = x;
                    if (y > T) T = y;
                    if (y < B) B = y;
                }
            } catch (e) { }
            return [L, T, R, B];
        }

        function extractBoundsFromGroupItem(g) {
            // Returns [L,T,R,B]
            try {
                return g.visibleBounds.slice(0);
            } catch (e) {
                return [0, 0, 0, 0];
            }
        }

        function findFirstPathItemInGroup(g) {
            try {
                for (var i = 0; i < g.pageItems.length; i++) {
                    var it = g.pageItems[i];
                    if (it && it.typename === "PathItem") return it;
                }
            } catch (e) { }
            return null;
        }

        function makeClosedRectFromBounds(bounds, layer) {
            // bounds: [L,T,R,B]
            var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
            var w = R - L;
            var h = T - B;
            var p = doc.pathItems.rectangle(T, L, w, h);
            p.closed = true;
            return p;
        }

        function normalizeRectCandidate(rectCandidate) {
            // If user selects an already-cut open path / non-rect path / group, reconstruct a 4-pt closed rectangle.
            // Returns a PathItem that is a closed 4-pt rectangle.
            if (!rectCandidate) return null;

            // If it's already a strict rectangle, accept as-is
            if (isRectPathItem(rectCandidate)) return rectCandidate;

            var layer = null;
            var b = null;
            var refForAppearance = null;

            if (isPathItem(rectCandidate)) {
                layer = rectCandidate.layer;
                b = extractRectBoundsFromPathItem(rectCandidate);
                refForAppearance = rectCandidate;
            } else if (isGroupItem(rectCandidate)) {
                layer = rectCandidate.layer;
                b = extractBoundsFromGroupItem(rectCandidate);
                refForAppearance = findFirstPathItemInGroup(rectCandidate) || rectCandidate;
            } else {
                return null;
            }

            var newRect = makeClosedRectFromBounds(b, layer);

            // Copy appearance from the candidate (or first path in group) to the new rectangle
            try {
                newRect.stroked = refForAppearance.stroked;
                newRect.filled = refForAppearance.filled;
                try { newRect.strokeColor = refForAppearance.strokeColor; } catch (e) { }
                try { newRect.fillColor = refForAppearance.fillColor; } catch (e) { }
                try { newRect.strokeWidth = refForAppearance.strokeWidth; } catch (e) { }
                try { newRect.strokeDashes = refForAppearance.strokeDashes; } catch (e) { }
                try { newRect.dashOffset = refForAppearance.dashOffset; } catch (e) { }
                try { newRect.strokeCap = refForAppearance.strokeCap; } catch (e) { }
                try { newRect.strokeJoin = refForAppearance.strokeJoin; } catch (e) { }
                try { newRect.miterLimit = refForAppearance.miterLimit; } catch (e) { }
                try { newRect.strokeOverprint = refForAppearance.strokeOverprint; } catch (e) { }
                try { newRect.opacity = refForAppearance.opacity; } catch (e) { }
                try { newRect.blendingMode = refForAppearance.blendingMode; } catch (e) { }
            } catch (e) { }

            // Keep stacking position roughly similar: put the new rect where the old one was.
            try { newRect.move(rectCandidate, ElementPlacement.PLACEAFTER); } catch (e) { }

            // Remove old candidate (cut path / group)
            try { rectCandidate.remove(); } catch (e) { }

            return newRect;
        }

        var a = doc.selection[0], b = doc.selection[1];
        var tf = null, rectA = null;

        // Accept either a strict rectangle, or an already-cut path (open) / non-4pt path or group and normalize it.
        if (isTextItem(a) && isStrokeCandidate(b)) { tf = a; rectA = b; }
        else if (isTextItem(b) && isStrokeCandidate(a)) { tf = b; rectA = a; }
        else { alert(getLabel('alertSelectTypes')); return; }

        rectA = normalizeRectCandidate(rectA);
        if (!rectA) { alert(getLabel('alertSelectTypes')); return; }

        // rulerType を参照して単位ラベルと pt 変換を決める
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

        var rulerUnitInfo = getUnitInfo("rulerType");
        var strokeUnitInfo = getUnitInfo("strokeUnits");

        function unitToPt(v) {
            return v * rulerUnitInfo.pointsPerUnit;
        }

        function strokeUnitToPt(v) {
            return v * strokeUnitInfo.pointsPerUnit;
        }

        function ptToStrokeUnit(vPt) {
            return vPt / strokeUnitInfo.pointsPerUnit;
        }

        // ---- 状態保持（プレビュー用）----
        var rectLayer = rectA.layer;
        var rectZ = rectA.zOrderPosition; // 同レイヤー内の相対順（AIにより不安定な場合あり）

        var orig = {
            id: rectA.uuid || null,
            // boundsは再生成に使う
            rBounds: rectA.visibleBounds.slice(0), // [L,T,R,B]
            // appearance (A基準)
            stroked: rectA.stroked,
            filled: rectA.filled,
            fillColor: null,
            strokeColor: null,
            strokeWidth: rectA.strokeWidth,
            strokeDashes: rectA.strokeDashes,
            dashOffset: rectA.dashOffset,
            strokeCap: rectA.strokeCap,
            strokeJoin: rectA.strokeJoin,
            miterLimit: rectA.miterLimit,
            strokeOverprint: rectA.strokeOverprint,
            opacity: rectA.opacity,
            blendingMode: rectA.blendingMode
        };
        try { orig.strokeColor = rectA.strokeColor; } catch (e) { }
        try { orig.fillColor = rectA.fillColor; } catch (e) { }

        var previewItem = null; // 生成したプレビュー用（PathItem または GroupItem）

        // A（元）/ B（塗りのみ）/ C（線のみ）
        var rectB = null;

        var rectC = null;

        function makeFillOnlyFromA(src) {
            // Aの塗りが「なし」の場合はBを作らない
            try { if (!src.filled) return null; } catch (e) { }

            var b = src.duplicate();
            try { b.hidden = false; } catch (e) { }
            b.stroked = false;
            b.filled = true;
            return b;
        }

        function makeStrokeOnlyFromA(src) {
            var c = src.duplicate();
            try { c.hidden = false; } catch (e) { }
            c.filled = false;
            c.stroked = true;
            return c;
        }

        function rebuildFillOnlyB() {
            // 角丸のON/OFFや半径変更に追従させるため、Bは作り直す（effectの二重掛け防止）
            try { if (rectB) rectB.remove(); } catch (e) { }
            rectB = null;

            // Aの塗りが「なし」の場合はBを生成しない
            try {
                if (!rectA.filled) return;
            } catch (e) { }

            rectB = makeFillOnlyFromA(rectA);
            if (!rectB) return;
            try { rectB.hidden = false; } catch (e) { }
            // 可能なら C の直前に配置（A/B/C の重なりを崩しにくい）
            try {
                if (rectC) rectB.move(rectC, ElementPlacement.PLACEBEFORE);
            } catch (e) { }

            // 角丸値（pt）を先に取得
            var rrB = 0;
            try {
                if (typeof cbRound !== "undefined" && cbRound.value) {
                    rrB = getRoundRadiusPt();
                }
            } catch (e) { rrB = 0; }

            // ノッチONならBに欠き取り（ライブ：パスファインダー減算）
            var didNotch = false;
            try {
                if (typeof cbNotch !== "undefined" && cbNotch.value) {
                    // UIができている（et/parseMarginが使える）ときだけ
                    if (typeof et !== "undefined" && typeof parseMargin === "function") {
                        var m = parseMargin();
                        if (m !== null) {
                            var mp = unitToPt(m);
                            var newB = applyNotchSubtractToFill(rectB, mp);
                            if (newB) {
                                rectB = newB; // GroupItemになる場合がある
                                didNotch = true;
                            }
                        }
                    }
                }
            } catch (e) { }

            // 角丸：ノッチON時はグループ（Z）に対して適用、OFF時は従来通りBに適用
            try {
                if (rrB > 0) {
                    applyRoundCornersEffect(rectB, rrB);
                }
            } catch (e) { }

            // Bは最背面へ（塗りを常に背面に）
            try {
                if (rectB) rectB.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (e) { }
        }

        rectC = makeStrokeOnlyFromA(rectA);
        rectB = makeFillOnlyFromA(rectA);
        // Aの塗りが「なし」の場合は rectB は null のまま

        // 線はデフォルトで有効
        var __strokeEnabled = true;

        // ダイアログ中はAを隠す（最終的にOK時にAは削除する）
        var __rectAHiddenForDialog = false;
        try { rectA.hidden = true; __rectAHiddenForDialog = true; } catch (e) { }

        function copyAppearance(srcObj, dst) {
            dst.stroked = srcObj.stroked;
            dst.strokeWidth = srcObj.strokeWidth;
            dst.strokeDashes = srcObj.strokeDashes;
            dst.dashOffset = srcObj.dashOffset;
            dst.strokeCap = srcObj.strokeCap;
            dst.strokeJoin = srcObj.strokeJoin;
            dst.miterLimit = srcObj.miterLimit;
            dst.strokeOverprint = srcObj.strokeOverprint;
            try { dst.strokeColor = srcObj.strokeColor; } catch (e) { }
            try { dst.opacity = srcObj.opacity; } catch (e) { }
            try { dst.blendingMode = srcObj.blendingMode; } catch (e) { }
        }

        // テキスト外接（計算用）
        // テキストは見た目上の外接とズレることがあるため、
        // 一時的に複製→アウトライン化して、その外接で計算する（1回だけ）
        var __textBoundsForCalc = null; // [L,T,R,B]

        function buildTextBoundsForCalcOnce() {
            if (__textBoundsForCalc) return;
            try {
                // 複製 → アウトライン化（元のtfは触らない）
                var dup = tf.duplicate();
                // createOutline() は複製テキストをアウトラインに変換し、戻り値はGroupItemになる
                var outlined = dup.createOutline();
                try { outlined.hidden = true; } catch (e) { }
                __textBoundsForCalc = outlined.visibleBounds.slice(0);
                try { outlined.remove(); } catch (e) { }
            } catch (e) {
                // フォールバック：通常のvisibleBounds
                try { __textBoundsForCalc = tf.visibleBounds.slice(0); } catch (e) { __textBoundsForCalc = null; }
            }
        }

        function getDefaultMarginValueInt() {
            // テキスト高の半分を、現在単位に変換して整数に丸める
            try {
                buildTextBoundsForCalcOnce();
                var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
                var hPt = Math.abs(tb[1] - tb[3]);
                var halfPt = hPt / 2;
                var v = halfPt / rulerUnitInfo.pointsPerUnit; // pt -> current unit
                if (isNaN(v) || !isFinite(v)) return 2;
                return Math.round(v);
            } catch (e) {
                return 2;
            }
        }

        var __defaultMarginInt = getDefaultMarginValueInt();

        function getDefaultRoundValueInt() {
            // テキスト高の1/4を、現在単位に変換して整数に丸める
            try {
                buildTextBoundsForCalcOnce();
                var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
                var hPt = Math.abs(tb[1] - tb[3]);
                var qPt = hPt / 4;
                var v = qPt / rulerUnitInfo.pointsPerUnit; // pt -> current unit
                if (isNaN(v) || !isFinite(v)) return 2;
                return Math.round(v);
            } catch (e) {
                return 2;
            }
        }

        var __defaultRoundInt = getDefaultRoundValueInt();

        function getDefaultStrokeWidthValue() {
            // 元の線幅（pt）を strokeUnits に変換して表示用にする
            try {
                var v = ptToStrokeUnit(orig.strokeWidth);
                if (isNaN(v) || !isFinite(v) || v <= 0) v = 1;
                // 見やすさ優先：小数1桁に丸め（整数なら整数表示になる）
                v = Math.round(v * 10) / 10;
                return String(v);
            } catch (e) {
                return "1";
            }
        }

        var __defaultStrokeWidthText = getDefaultStrokeWidthValue();

        function getCutRange(addPt) {
            buildTextBoundsForCalcOnce();
            var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
            return { cutL: tb[0] - addPt, cutR: tb[2] + addPt };
        }

        function clampCutToRect(cut, rL, rR) {
            var cutL = cut.cutL, cutR = cut.cutR;
            if (cutL < rL) cutL = rL;
            if (cutR > rR) cutR = rR;
            return { cutL: cutL, cutR: cutR };
        }

        function makeRectByBounds(bounds, layer) {
            // bounds: [L,T,R,B]
            var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
            var w = R - L;
            var h = T - B;
            var p = doc.pathItems.rectangle(T, L, w, h);
            try { p.stroked = false; } catch (e) { }
            try { p.filled = false; } catch (e) { }
            if (layer) {
                try { p.move(layer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
            }
            return p;
        }

        function applyNotchSubtractToFill(fillItem, marginPt) {
            // Returns the new container (GroupItem) that has live pathfinder subtract applied.
            // If it fails, returns the original fillItem.
            if (!fillItem || !cbNotch || !cbNotch.value) return fillItem;
            if (marginPt === undefined || marginPt === null) return fillItem;

            try {
                buildTextBoundsForCalcOnce();
                var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]

                // X: text bounds rectangle (calc-only)
                var rectX = makeRectByBounds(tb, fillItem.layer);

                // Y: expanded by margin on all sides
                var yb = [tb[0] - marginPt, tb[1] + marginPt, tb[2] + marginPt, tb[3] - marginPt];
                var rectY = makeRectByBounds(yb, fillItem.layer);

                // Yは減算用カッターなので filled=true にして面として成立させる
                try { rectY.stroked = false; } catch (e) { }
                try { rectY.filled = true; } catch (e) { }
                try { rectY.hidden = false; } catch (e) { }
                try { rectY.fillColor = fillItem.fillColor; } catch (e) { }
                try { rectY.opacity = 100; } catch (e) { }

                // Z: group (fill + Y)
                var g = doc.groupItems.add();
                try { g.move(fillItem.layer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }

                // order: fill behind, cutter(Y) in front
                try { fillItem.move(g, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
                try { rectY.move(g, ElementPlacement.PLACEATEND); } catch (e) { }

                // X is calc-only
                try { rectX.remove(); } catch (e) { }

                // Apply Live Pathfinder Subtract to Z
                try {
                    doc.selection = null;
                    g.selected = true;
                    app.executeMenuCommand('Live Pathfinder Minus Back');
                } catch (e) { }

                try { doc.selection = null; } catch (e) { }

                return g;
            } catch (e) {
                return fillItem;
            }
        }

        // プレビュー適用：元rectは残したまま、上に“欠けた罫線”を重ねて見せる
        //（元の上辺を消さずに重ねる方式なので、プレビューは「完成形と同じ」にするため
        //  いったん元rectを非表示にして、プレビューで置き換えて表示します）
        var origVisible = rectA.hidden;

        function clearPreview() {
            if (previewItem) {
                try { previewItem.remove(); } catch (e) { }
                previewItem = null;
            }
            try { if (rectC) rectC.hidden = false; } catch (e) { }
        }

        function chooseCutEdge(tb, rL, rT, rR, rBot) {
            // Decide which rectangle edge (top/bottom/left/right) is closest to text center
            // tb: [L,T,R,B]
            try {
                var cx = (tb[0] + tb[2]) / 2;
                var cy = (tb[1] + tb[3]) / 2;

                var dTop = Math.abs(rT - cy);
                var dBottom = Math.abs(cy - rBot);
                var dLeft = Math.abs(cx - rL);
                var dRight = Math.abs(rR - cx);

                var min = Math.min(dTop, dBottom, dLeft, dRight);
                if (min === dTop) return "top";
                if (min === dBottom) return "bottom";
                if (min === dLeft) return "left";
                return "right";
            } catch (e) {
                return "top";
            }
        }

        function getCutRangeV(addPt) {
            buildTextBoundsForCalcOnce();
            var tb = __textBoundsForCalc || tf.visibleBounds; // [L,T,R,B]
            return { cutT: tb[1] + addPt, cutB: tb[3] - addPt };
        }

        function clampCutVToRect(cutV, rT, rBot) {
            var cutT = cutV.cutT, cutB = cutV.cutB;
            if (cutT > rT) cutT = rT;
            if (cutB < rBot) cutB = rBot;
            return { cutT: cutT, cutB: cutB };
        }

        function chooseCutEdges(tb, rL, rT, rR, rBot, addPt) {
            // Return 1 edge normally; return 2 *adjacent* edges when near a corner.
            try {
                var cx = (tb[0] + tb[2]) / 2;
                var cy = (tb[1] + tb[3]) / 2;
                var dTop = Math.abs(rT - cy);
                var dBottom = Math.abs(cy - rBot);
                var dLeft = Math.abs(cx - rL);
                var dRight = Math.abs(rR - cx);
                var arr = [
                    { e: "top", d: dTop },
                    { e: "bottom", d: dBottom },
                    { e: "left", d: dLeft },
                    { e: "right", d: dRight }
                ];
                arr.sort(function (a, b) { return a.d - b.d; });
                var e1 = arr[0].e, d1 = arr[0].d;
                var e2 = arr[1].e, d2 = arr[1].d;

                // Threshold: if the 2nd-closest edge is nearly as close as the closest edge, treat as corner.
                var th = Math.max(2, addPt * 0.8);

                function isAdjacent(a, b) {
                    return ((a === "top" || a === "bottom") && (b === "left" || b === "right")) ||
                        ((b === "top" || b === "bottom") && (a === "left" || a === "right"));
                }

                if (isAdjacent(e1, e2) && (d2 - d1) <= th) return [e1, e2];
                return [e1];
            } catch (e) {
                return [chooseCutEdge(tb, rL, rT, rR, rBot)];
            }
        }

        function applyAppearanceToPreview(preview, refStrokeObj) {
            if (!preview) return;
            try {
                if (preview.typename === "GroupItem") {
                    for (var i = 0; i < preview.pageItems.length; i++) {
                        var it = preview.pageItems[i];
                        if (it && it.typename === "PathItem") {
                            try { copyAppearance(refStrokeObj, it); } catch (e) { }
                        }
                    }
                } else if (preview.typename === "PathItem") {
                    try { copyAppearance(refStrokeObj, preview); } catch (e) { }
                }
            } catch (e) { }
        }

        function setStrokeWidthToPreview(preview, swPt) {
            if (!preview) return;
            try {
                if (preview.typename === "GroupItem") {
                    for (var i = 0; i < preview.pageItems.length; i++) {
                        var it = preview.pageItems[i];
                        if (it && it.typename === "PathItem") {
                            try { it.strokeWidth = swPt; } catch (e) { }
                        }
                    }
                } else if (preview.typename === "PathItem") {
                    preview.strokeWidth = swPt;
                }
            } catch (e) { }
        }

        function setStrokeCapToPreview(preview, cap) {
            if (!preview) return;
            try {
                if (preview.typename === "GroupItem") {
                    for (var i = 0; i < preview.pageItems.length; i++) {
                        var it = preview.pageItems[i];
                        if (it && it.typename === "PathItem") {
                            try { it.strokeCap = cap; } catch (e) { }
                        }
                    }
                } else if (preview.typename === "PathItem") {
                    preview.strokeCap = cap;
                }
            } catch (e) { }
        }

        function isHorizontalEdge(edge) {
            return edge === "top" || edge === "bottom";
        }
        function isVerticalEdge(edge) {
            return edge === "left" || edge === "right";
        }

        function applyPreview(marginMM) {
            clearPreview();

            // 線が無効ならCを作らない（=欠き取りもしない）
            try {
                if (!cbStrokeOn.value || !rectC) return;
            } catch (e) { return; }

            var addPt = unitToPt(marginMM);
            var rB = orig.rBounds;
            var rL = rB[0], rT = rB[1], rR = rB[2], rBot = rB[3];

            // Decide which edge(s) to cut based on the text position
            buildTextBoundsForCalcOnce();
            var tbEdge = __textBoundsForCalc || tf.visibleBounds;
            // For 2-edge support, use chooseCutEdges:
            var cutEdges = chooseCutEdges(tbEdge, rL, rT, rR, rBot, addPt);
            var cutEdge = cutEdges[0];
            var cut = getCutRange(addPt);
            var cutV = getCutRangeV(addPt);

            // 交差しないなら何もしない（元を表示）
            // - top/bottom: need horizontal overlap
            // - left/right: need vertical overlap
            var ccTmp = clampCutToRect(cut, rL, rR);
            var cvTmp = clampCutVToRect(cutV, rT, rBot);

            var needH = false;
            var needV = false;
            for (var ei = 0; ei < cutEdges.length; ei++) {
                if (isHorizontalEdge(cutEdges[ei])) needH = true;
                if (isVerticalEdge(cutEdges[ei])) needV = true;
            }

            var okH = (!needH) || (ccTmp.cutR > ccTmp.cutL);
            var okV = (!needV) || (cvTmp.cutT > cvTmp.cutB);

            if (!okH || !okV) {
                try { if (rectC) rectC.hidden = false; } catch (e) { }
                return;
            }

            var cc = ccTmp;
            var cv = cvTmp;

            // 元rectを隠して、開いたパス（欠け罫線）で見せる
            rectC.hidden = true;
            function makePath(points) {
                // Avoid corner dots/artifacts:
                // - remove consecutive duplicate points
                // - nudge endpoints toward a *distinct* neighbor
                // - if degenerate (<2 distinct points), skip creating a path

                function samePt(a, b) {
                    return a && b && Math.abs(a[0] - b[0]) < 1e-6 && Math.abs(a[1] - b[1]) < 1e-6;
                }

                function compactConsecutive(pts) {
                    if (!pts || pts.length === 0) return [];
                    var out = [pts[0]];
                    for (var i = 1; i < pts.length; i++) {
                        if (!samePt(pts[i], out[out.length - 1])) out.push(pts[i]);
                    }
                    return out;
                }

                try {
                    points = compactConsecutive(points);

                    // If all points are the same or only one distinct point, do nothing
                    if (!points || points.length < 2) return null;

                    var eps = 0.05; // pt
                    try {
                        var sw0 = (rectC && rectC.stroked) ? rectC.strokeWidth : 1;
                        eps = Math.max(0.05, sw0 * 0.25);
                    } catch (e) { }

                    function nudgeEnd(idx, step) {
                        // step: +1 to search forward, -1 to search backward for a distinct neighbor
                        var p0 = points[idx];
                        var j = idx + step;
                        while (j >= 0 && j < points.length && samePt(points[j], p0)) {
                            j += step;
                        }
                        if (j < 0 || j >= points.length) return;

                        var p1 = points[j];
                        var vx = p1[0] - p0[0];
                        var vy = p1[1] - p0[1];
                        var len = Math.sqrt(vx * vx + vy * vy);
                        if (len <= 1e-6) return;
                        vx /= len; vy /= len;
                        points[idx] = [p0[0] + vx * eps, p0[1] + vy * eps];
                    }

                    // first point nudged toward next distinct
                    nudgeEnd(0, +1);
                    // last point nudged toward previous distinct
                    nudgeEnd(points.length - 1, -1);

                    // Re-compact after nudging (in case endpoints collapsed)
                    points = compactConsecutive(points);
                    if (!points || points.length < 2) return null;
                } catch (e) { }

                var p = doc.pathItems.add();
                p.setEntirePath(points);
                p.closed = false;
                p.filled = false;
                p.stroked = true;
                return p;
            }

            // If only one edge, behave as before
            if (cutEdges.length === 1) {
                if (cutEdge === "bottom") {
                    previewItem = makePath([
                        [cc.cutR, rBot],
                        [rR, rBot],
                        [rR, rT],
                        [rL, rT],
                        [rL, rBot],
                        [cc.cutL, rBot]
                    ]);
                } else if (cutEdge === "left") {
                    previewItem = makePath([
                        [rL, cv.cutT],
                        [rL, rT],
                        [rR, rT],
                        [rR, rBot],
                        [rL, rBot],
                        [rL, cv.cutB]
                    ]);
                } else if (cutEdge === "right") {
                    previewItem = makePath([
                        [rR, cv.cutT],
                        [rR, rT],
                        [rL, rT],
                        [rL, rBot],
                        [rR, rBot],
                        [rR, cv.cutB]
                    ]);
                } else {
                    // top (default)
                    previewItem = makePath([
                        [cc.cutR, rT],
                        [rR, rT],
                        [rR, rBot],
                        [rL, rBot],
                        [rL, rT],
                        [cc.cutL, rT]
                    ]);
                }
            } else {
                // corner case: two adjacent edges (top+left, top+right, bottom+left, bottom+right)
                // Build TWO open paths that split the perimeter, so no segment is duplicated.
                var e1 = cutEdges[0];
                var e2 = cutEdges[1];

                // Normalize to (h, v)
                var h = (e1 === "top" || e1 === "bottom") ? e1 : e2;
                var v = (h === e1) ? e2 : e1;

                if (!((h === "top" || h === "bottom") && (v === "left" || v === "right"))) {
                    previewItem = null; // unexpected, will fallback
                } else {
                    var g = doc.groupItems.add();

                    if (h === "top" && v === "left") {
                        // Missing: top [cc.cutL..cc.cutR] and left [cv.cutB..cv.cutT]
                        // PathA: top-right gap end -> around to left gap bottom
                        var pA = makePath([
                            [cc.cutR, rT],
                            [rR, rT],
                            [rR, rBot],
                            [rL, rBot],
                            [rL, cv.cutB]
                        ]);
                        // PathB: left gap top -> corner -> top gap left
                        var pB = makePath([
                            [rL, cv.cutT],
                            [rL, rT],
                            [cc.cutL, rT]
                        ]);
                        if (pA) pA.move(g, ElementPlacement.PLACEATEND);
                        if (pB) pB.move(g, ElementPlacement.PLACEATEND);
                    } else if (h === "top" && v === "right") {
                        // Missing: top [cc.cutL..cc.cutR] and right [cv.cutB..cv.cutT]
                        // PathA: top gap right -> corner -> right gap top
                        var pA2 = makePath([
                            [cc.cutR, rT],
                            [rR, rT],
                            [rR, cv.cutT]
                        ]);
                        // PathB: right gap bottom -> around to top gap left
                        var pB2 = makePath([
                            [rR, cv.cutB],
                            [rR, rBot],
                            [rL, rBot],
                            [rL, rT],
                            [cc.cutL, rT]
                        ]);
                        if (pA2) pA2.move(g, ElementPlacement.PLACEATEND);
                        if (pB2) pB2.move(g, ElementPlacement.PLACEATEND);
                    } else if (h === "bottom" && v === "left") {
                        // Missing: bottom [cc.cutL..cc.cutR] and left [cv.cutB..cv.cutT]
                        // PathA: bottom gap right -> around to left gap top
                        var pA3 = makePath([
                            [cc.cutR, rBot],
                            [rR, rBot],
                            [rR, rT],
                            [rL, rT],
                            [rL, cv.cutT]
                        ]);
                        // PathB: left gap bottom -> corner -> bottom gap left
                        var pB3 = makePath([
                            [rL, cv.cutB],
                            [rL, rBot],
                            [cc.cutL, rBot]
                        ]);
                        if (pA3) pA3.move(g, ElementPlacement.PLACEATEND);
                        if (pB3) pB3.move(g, ElementPlacement.PLACEATEND);
                    } else {
                        // bottom + right
                        // Missing: bottom [cc.cutL..cc.cutR] and right [cv.cutB..cv.cutT]
                        // PathA: bottom gap right -> corner -> right gap bottom
                        var pA4 = makePath([
                            [cc.cutR, rBot],
                            [rR, rBot],
                            [rR, cv.cutB]
                        ]);
                        // PathB: right gap top -> around to bottom gap left
                        var pB4 = makePath([
                            [rR, cv.cutT],
                            [rR, rT],
                            [rL, rT],
                            [rL, rBot],
                            [cc.cutL, rBot]
                        ]);
                        if (pA4) pA4.move(g, ElementPlacement.PLACEATEND);
                        if (pB4) pB4.move(g, ElementPlacement.PLACEATEND);
                    }

                    // If paths were not created, fallback
                    try {
                        if (g.pageItems.length === 0) {
                            g.remove();
                            previewItem = null;
                        } else {
                            previewItem = g;
                        }
                    } catch (e) {
                        previewItem = null;
                    }
                }
            }

            if (!previewItem) {
                // Fallback: use the nearest single edge
                var fe = chooseCutEdge(tbEdge, rL, rT, rR, rBot);
                if (fe === "bottom") {
                    previewItem = makePath([[cc.cutR, rBot], [rR, rBot], [rR, rT], [rL, rT], [rL, rBot], [cc.cutL, rBot]]);
                } else if (fe === "left") {
                    previewItem = makePath([[rL, cv.cutT], [rL, rT], [rR, rT], [rR, rBot], [rL, rBot], [rL, cv.cutB]]);
                } else if (fe === "right") {
                    previewItem = makePath([[rR, cv.cutT], [rR, rT], [rL, rT], [rL, rBot], [rR, rBot], [rR, cv.cutB]]);
                } else {
                    previewItem = makePath([[cc.cutR, rT], [rR, rT], [rR, rBot], [rL, rBot], [rL, rT], [cc.cutL, rT]]);
                }
            }
            try { applyAppearanceToPreview(previewItem, rectC); } catch (e) { }

            // 線幅（入力が有効なら上書き）
            try {
                var sw = parseStrokeWidth();
                if (sw !== null) setStrokeWidthToPreview(previewItem, strokeUnitToPt(sw));
            } catch (e) { }

            try {
                var selCap = getSelectedCap();
                setStrokeCapToPreview(previewItem, selCap);
            } catch (e) { }

            // 角丸（重複適用を防ぐ）
            try {
                var rr = getRoundRadiusPt();
                if (rr > 0) {
                    if (previewItem && previewItem.typename === "GroupItem") {
                        for (var i = 0; i < previewItem.pageItems.length; i++) {
                            var it2 = previewItem.pageItems[i];
                            if (it2 && it2.typename === "PathItem") {
                                applyRoundCornersEffect(it2, rr);
                            }
                        }
                    } else {
                        applyRoundCornersEffect(previewItem, rr);
                    }
                }
            } catch (e) { }

            // レイヤーを合わせる
            try { previewItem.move((rectC ? rectC.layer : rectLayer), ElementPlacement.PLACEATBEGINNING); } catch (e) { }
            // 可能なら元rectの直後/直前へ近づけたいが、zOrderPositionの厳密復元は難しいため最前面寄せ
            try { previewItem.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e) { }
        }

        // ---- ダイアログ / Dialog ----
        var dlg = new Window("dialog", getLabel('dialogTitle') + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        // ダイアログの位置 / Dialog position
        var offsetX = 300;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function () {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        shiftDialogPosition(dlg, offsetX, 0);

        var row = dlg.add("group");
        row.alignChildren = ["left", "center"];

        var stMarginLabel = row.add("statictext", undefined, getLabel('margin'));
        stMarginLabel.preferredSize.width = 60;
        stMarginLabel.justify = "right";

        var et = addSteppedInput(row, String(__defaultMarginInt), { min: 0 }); // マージンは負値NG
        et.helpTip = getLabel('tipMargin');
        et.characters = 3;
        try {
            if ($.global.__tblc_state.marginText !== null && $.global.__tblc_state.marginText !== undefined) {
                et.text = String($.global.__tblc_state.marginText);
            }
        } catch (e) { }

        var stUnitMargin = row.add("statictext", undefined, rulerUnitInfo.label);

        // 角丸（マージンの直下）
        var rowR = dlg.add("group");
        rowR.alignChildren = ["left", "center"];

        var cbRound = rowR.add("checkbox", undefined, getLabel('roundCorners'));
        cbRound.helpTip = getLabel('tipRound');
        cbRound.value = false;
        cbRound.value = !!$.global.__tblc_state.roundOn;

        var etRound = addSteppedInput(rowR, String(__defaultRoundInt), { min: 0 }); // 角丸半径も負値NG
        etRound.helpTip = getLabel('tipRound');
        try {
            if ($.global.__tblc_state.roundText !== null && $.global.__tblc_state.roundText !== undefined) {
                etRound.text = String($.global.__tblc_state.roundText);
            }
        } catch (e) { }
        etRound.characters = 3;

        var stUnitRound = rowR.add("statictext", undefined, rulerUnitInfo.label);
        setSteppedInputEnabled(etRound, false);

        // 塗り
        var fillPanel = dlg.add("panel", undefined, getLabel('panelFill'));
        fillPanel.orientation = "column";
        fillPanel.alignChildren = ["fill", "top"];
        fillPanel.margins = [15, 20, 15, 10];

        var rowF = fillPanel.add("group");
        rowF.alignChildren = ["left", "center"];

        var cbFillOn = rowF.add("checkbox", undefined, getLabel('fillOn'));
        cbFillOn.helpTip = getLabel('tipFillOn');
        cbFillOn.value = true;
        cbFillOn.value = !!$.global.__tblc_state.fillOn;

        var rowN = fillPanel.add("group");
        rowN.alignChildren = ["left", "center"];

        var cbNotch = rowN.add("checkbox", undefined, getLabel('notch'));
        cbNotch.helpTip = getLabel('tipNotch');
        cbNotch.value = false;
        cbNotch.value = !!$.global.__tblc_state.notchOn;

        // 線（線幅・線端）
        var linePanel = dlg.add("panel", undefined, getLabel('panelStroke'));
        linePanel.orientation = "column";
        linePanel.alignChildren = ["fill", "top"];
        linePanel.margins = [15, 20, 15, 10];

        // 線幅（strokeUnits）
        var rowW = linePanel.add("group");
        rowW.alignChildren = ["left", "center"];

        var cbStrokeOn = rowW.add("checkbox", undefined, "");
        cbStrokeOn.helpTip = getLabel('tipStrokeOn');
        cbStrokeOn.value = true;
        cbStrokeOn.value = !!$.global.__tblc_state.strokeOn;

        var stWidthLabel = rowW.add("statictext", undefined, getLabel('strokeWidth'));

        var etWidth = addSteppedInput(rowW, __defaultStrokeWidthText, { min: 0.1 }); // 線幅は最低0.1
        etWidth.helpTip = getLabel('tipStrokeOn');
        etWidth.characters = 3;
        try {
            if ($.global.__tblc_state.widthText !== null && $.global.__tblc_state.widthText !== undefined) {
                etWidth.text = String($.global.__tblc_state.widthText);
            }
        } catch (e) { }

        var stUnitWidth = rowW.add("statictext", undefined, strokeUnitInfo.label);

        // 線端
        var capRow = linePanel.add("group");
        capRow.orientation = "row";
        capRow.alignChildren = ["left", "center"];

        var stCapLabel = capRow.add("statictext", undefined, getLabel('lineCap'));

        var capBtns = capRow.add("group");
        capBtns.alignment = ["left", "center"];
        capBtns.orientation = "row";
        capBtns.alignChildren = ["left", "center"];

        var rbCapNone = capBtns.add("radiobutton", undefined, getLabel('capNone'));
        rbCapNone.helpTip = getLabel('tipCapNone');
        var rbCapRound = capBtns.add("radiobutton", undefined, getLabel('capRound'));
        rbCapRound.helpTip = getLabel('tipCapRound');

        // セッション復元（線端）
        var __capIdx = 0;
        try { __capIdx = ($.global.__tblc_state.capIndex === 1) ? 1 : 0; } catch (e) { __capIdx = 0; }
        rbCapNone.value = (__capIdx === 0);
        rbCapRound.value = (__capIdx === 1);

        function getSelectedCap() {
            // Always return an explicit cap.
            // 「なし」= BUTT / 「丸型」= ROUND
            if (rbCapRound.value) return StrokeCap.ROUNDENDCAP;
            return StrokeCap.BUTTENDCAP;
        }

        function parseRoundRadius() {
            if (!cbRound.value) return 0;
            var v = parseFloat(etRound.text);
            if (isNaN(v) || v <= 0) return 0;
            return v;
        }

        function getRoundRadiusPt() {
            return unitToPt(parseRoundRadius());
        }

        function createRoundCornersEffectXML(radiusPt) {
            // Reference format: <Dict data="R radius #value# "/>
            // radiusPt is in points
            var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>';
            return xml;
        }

        function applyRoundCornersEffect(item, radiusPt) {
            if (!item || radiusPt <= 0) return;
            // Apply via applyEffect() for reliable rendering, but make it idempotent by removing existing blocks first.
            try {
                removeRoundCornersEffect(item);
            } catch (e) { }
            try {
                item.applyEffect(createRoundCornersEffectXML(radiusPt));
            } catch (e) { }
        }

        function removeRoundCornersEffect(item) {
            // Remove existing "Adobe Round Corners" live effect only (do not wipe other effects).
            // Illustrator may emit name attribute with single/double quotes, extra attrs, and whitespace/newlines.
            try {
                if (!item) return;
                var ae = item.appliedEffect;
                if (!ae) return;

                // Match any LiveEffect tag whose name is Adobe Round Corners
                var re = /<LiveEffect\b[^>]*\bname\s*=\s*(['"])\s*Adobe Round Corners\s*\1[^>]*>[\s\S]*?<\/LiveEffect>/g;

                // Loop until stable (in case Illustrator nests/duplicates blocks)
                var prev;
                do {
                    prev = ae;
                    ae = ae.replace(re, "");
                } while (ae !== prev);

                item.appliedEffect = ae;
            } catch (e) { }
        }

        function syncRoundUI() {
            setSteppedInputEnabled(etRound, cbRound.value);
        }

        function syncFillUI() {
            cbNotch.enabled = !!cbFillOn.value;
        }

        function syncStrokeUIAndObjects() {
            var on = true;
            try { on = !!cbStrokeOn.value; } catch (e) { on = true; }

            // UI enable/disable
            setSteppedInputEnabled(etWidth, on);
            capRow.enabled = on;

            // Object create/remove
            if (!on) {
                try { clearPreview(); } catch (e) { }
                try { if (rectC) rectC.remove(); } catch (e) { }
                rectC = null;
                return;
            }

            // ON: rectCが無ければ作る
            if (!rectC) {
                try {
                    rectC = makeStrokeOnlyFromA(rectA);
                    // 可能ならBの後ろへ
                    try { if (rectB) rectC.move(rectB, ElementPlacement.PLACEAFTER); } catch (e) { }
                } catch (e) {
                    rectC = null;
                }
            }
        }

        function persistDialogValues() {
            try {
                $.global.__tblc_state.marginText = et.text;
                $.global.__tblc_state.roundOn = !!cbRound.value;
                $.global.__tblc_state.roundText = etRound.text;
                $.global.__tblc_state.fillOn = !!cbFillOn.value;
                $.global.__tblc_state.notchOn = !!cbNotch.value;
                $.global.__tblc_state.strokeOn = !!cbStrokeOn.value;
                $.global.__tblc_state.widthText = etWidth.text;
                $.global.__tblc_state.capIndex = rbCapRound.value ? 1 : 0;
            } catch (e) { }
        }

        var btns = dlg.add("group");
        btns.alignment = "right";
        var btCancel = btns.add("button", undefined, getLabel('cancel'), { name: "cancel" });
        var btOK = btns.add("button", undefined, getLabel('ok'), { name: "ok" });

        function parseMargin() {
            var v = parseFloat(et.text);
            if (isNaN(v) || v < 0) return null;
            return v;
        }

        function parseStrokeWidth() {
            var v = parseFloat(etWidth.text);
            if (isNaN(v) || v < 0.1) return null; // 0は不可、最低0.1
            return v;
        }

        function updatePreview() {
            // B（塗りのみ）を角丸設定に追従（塗りOFFなら削除）
            try {
                if (typeof cbFillOn !== "undefined" && !cbFillOn.value) {
                    if (rectB) { try { rectB.remove(); } catch (e) { } }
                    rectB = null;
                } else {
                    rebuildFillOnlyB();
                }
            } catch (e) { }

            var m = parseMargin();
            if (m === null) return; // 不正値は触らない

            // 線が有効なときだけ欠き取りプレビュー
            try {
                if (cbStrokeOn.value && rectC) applyPreview(m);
                else clearPreview();
            } catch (e) {
                clearPreview();
            }
            app.redraw();
        }

        et.onChanging = updatePreview;
        etRound.onChanging = updatePreview;
        etWidth.onChanging = updatePreview;

        cbStrokeOn.onClick = function () {
            syncStrokeUIAndObjects();
            updatePreview();
        };

        cbRound.onClick = function () {
            syncRoundUI();
            // 角丸ONのときは線端を「丸型」に寄せる
            if (cbRound.value) {
                rbCapRound.value = true;
            }
            updatePreview();
        };
        cbNotch.onClick = updatePreview;
        cbFillOn.onClick = function () {
            syncFillUI();
            updatePreview();
        };
        rbCapNone.onClick = updatePreview;
        rbCapRound.onClick = updatePreview;

        // 初期状態
        syncRoundUI();
        syncFillUI();
        syncStrokeUIAndObjects();

        // テキスト外接（計算用）を先に確定（1回だけ）
        buildTextBoundsForCalcOnce();

        // 初期のBを角丸設定に合わせて作り直し
        rebuildFillOnlyB();

        // 初期プレビュー
        updatePreview();

        prepareDialogWindow(dlg, SCRIPT_NAME);
        var result = dlg.show();
        persistDialogValues();

        if (result !== 1) {
            // キャンセル：プレビュー解除し、B/Cを破棄してAを戻す
            clearPreview();
            try { if (rectB) rectB.remove(); } catch (e) { }
            try { if (rectC) rectC.remove(); } catch (e) { }
            try { if (__rectAHiddenForDialog) rectA.hidden = false; } catch (e) { }
            app.redraw();
            return;
        }

        // OK：プレビューを確定（元rectを置き換え）
        var marginMM = parseMargin();
        if (marginMM === null) {
            clearPreview();
            alert(LF('alertMarginNonNegative', rulerUnitInfo.label));
            return;
        }

        // 線が無効なら：Cも欠き取りも無し。Bのみ残してAを削除（Bが無い場合はAを戻して終了）
        if (!cbStrokeOn.value) {
            try { if (rectB) rectB.hidden = false; } catch (e) { }
            if (!rectB) {
                try { if (__rectAHiddenForDialog) rectA.hidden = false; } catch (e) { }
                return;
            }
            try { rectA.remove(); } catch (e) { }
            return;
        }

        // プレビューがあるならそれを採用、なければ生成
        if (!previewItem) {
            applyPreview(marginMM);
        }

        // 交差しない等でpreviewItemが作られてない場合は何もしない
        if (!previewItem) {
            // 交差しない等で欠き取りが発生しない場合：Aを残すのではなく、B/Cを破棄してAを戻す
            try { if (rectB) rectB.remove(); } catch (e) { }
            try { if (rectC) rectC.remove(); } catch (e) { }
            try { if (__rectAHiddenForDialog) rectA.hidden = false; } catch (e) { }
            return;
        }

        // Cを削除し、previewItemを本番にする（Aは削除、Bは残す）
        try { if (rectC) rectC.remove(); } catch (e) { }
        try { rectA.remove(); } catch (e) { }

        // 念のため表示
        try { previewItem.hidden = false; } catch (e) { }

        // 最終反映（線端）
        try {
            var finalCap = getSelectedCap();
            setStrokeCapToPreview(previewItem, finalCap);
        } catch (e) { }

        // 最終反映（線幅）
        try {
            var finalSW = parseStrokeWidth();
            if (finalSW !== null) setStrokeWidthToPreview(previewItem, strokeUnitToPt(finalSW));
        } catch (e) { }

        // 角丸はプレビュー生成（applyPreview）時点で previewItem に反映済み。
        // ここで再適用すると二重にスタックするため、最終反映では行わない。

    })();

})();
