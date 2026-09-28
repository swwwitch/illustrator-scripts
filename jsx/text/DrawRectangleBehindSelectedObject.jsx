#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### 概要

選択したオブジェクトの背面に、マージン・角丸・カラーを指定した長方形を作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawRectangleBehindSelectedObject.md

### Overview

Draws a rectangle with margins, rounded corners, and a color of your choice behind the selected objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawRectangleBehindSelectedObject.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DrawRectangleBehindSelectedObject"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawRectangleBehindSelectedObject.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawRectangleBehindSelectedObject.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // デバッグ / Debug
    // =========================================
    var DEBUG_MODE = false;              /* true でエラーを ExtendScript コンソールに出す / Log errors to the console */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X      = 300;                 /* 初回表示位置の横ずらし（+右 / -左）/ Initial shift right (+) or left (-) */
    var DIALOG_OFFSET_Y      = 0;                   /* 初回表示位置の縦ずらし（+下 / -上）/ Initial shift down (+) or up (-) */
    var COLUMN_SPACING       = 16;                  /* 2カラムの間隔 / Gap between columns */
    var COLUMN_INNER_SPACING = 12;                  /* カラム内のパネル間隔 / Gap between panels in a column */
    var PANEL_MARGINS        = [15, 20, 15, 10];    /* パネル余白 [左,上,右,下] / Panel margins */
    var MARGIN_PANEL_MARGINS = [15, 15, 20, 15];    /* マージンパネルの余白 / Margin panel margins */
    var NUMBER_FIELD_WIDTH   = 35;                  /* 数値欄の幅 / Numeric field width */
    var CMYK_COLUMN_WIDTH    = 40;                  /* CMYK の列幅 / CMYK column width */
    var OPACITY_FIELD_WIDTH  = 40;                  /* 不透明度欄の幅 / Opacity field width */

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

    // =========================================
    // 内部名・定数 / Internal names and constants
    // =========================================
    var PREVIEW_LAYER_NAME = "_preview";                              /* プレビューレイヤー名 / Preview layer name */
    var PREVIEW_LAYER_NAMES = [PREVIEW_LAYER_NAME, "プレビュー", "Preview"]; /* 旧版の名前も片付け対象 / Includes legacy names */
    var PREVIEW_RECT_BASE_NAMES = { ja: "__プレビュー_アートボード境界", en: "__Preview_ArtboardBounds" }; /* プレビュー矩形名の接頭辞 / Preview rect name prefix */
    var TEMP_OUTLINE_LAYER_NAME = "__tmp_outline_bounds__";          /* アウトライン計測用の一時レイヤー / Temp layer for outline measuring */
    var FINAL_RECT_NAME = "BG_Rect";                                  /* 確定した長方形の名前 / Name of the final rectangle */
    var STATE_FILE_NAME = "DrawRectangleBehindSelectedObject.state";  /* 前回の設定を保存するファイル / Last-settings file */

    /* カラーモード / Color modes */
    var ColorMode = {
        NONE: 'none',
        K100: 'k100',
        WHITE: 'white',
        HEX: 'hex',
        CMYK: 'cmyk'
    };

    /* 名前で指定できる色 / Named colors accepted in the HEX field */
    var NAMED_COLORS = {
        black:   { rgb: [0, 0, 0],       cmyk: [0, 0, 0, 100] },
        white:   { rgb: [255, 255, 255], cmyk: [0, 0, 0, 0] },
        red:     { rgb: [255, 0, 0],     cmyk: [0, 100, 100, 0] },
        green:   { rgb: [0, 128, 0],     cmyk: [100, 0, 100, 50] },
        blue:    { rgb: [0, 0, 255],     cmyk: [100, 100, 0, 0] },
        cyan:    { rgb: [0, 255, 255],   cmyk: [100, 0, 0, 0] },
        magenta: { rgb: [255, 0, 255],   cmyk: [0, 100, 0, 0] },
        yellow:  { rgb: [255, 255, 0],   cmyk: [0, 0, 100, 0] },
        orange:  { rgb: [255, 165, 0],   cmyk: [0, 35, 100, 0] }
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "背面に長方形を作成", en: "Draw Rectangle Behind Selection" }
        },
        panel: {
            margin: { ja: "マージン", en: "Margins" },
            corner: { ja: "角丸", en: "Corner Radius" },
            color: { ja: "カラー", en: "Color" },
            target: { ja: "対象", en: "Target" },
            opacity: { ja: "不透明度", en: "Opacity" },
            paintType: { ja: "種別", en: "Type" }
        },
        fieldLabel: {
            offsetV: { ja: "上下", en: "Vertical" },
            offsetH: { ja: "左右", en: "Horizontal" },
            strokeWidth: { ja: "線幅", en: "Stroke Width" }
        },
        checkbox: {
            pillShape: { ja: "ピル形状", en: "Pill Shape" },
            opacityApply: { ja: "適用", en: "Apply" },
            groupWithObjects: { ja: "オブジェクトとグループ化", en: "Group with Objects" }
        },
        radio: {
            colorBlack: { ja: "ブラック", en: "Black" },
            colorWhite: { ja: "ホワイト", en: "White" },
            colorHex: { ja: "HEX", en: "HEX" },
            colorCmyk: { ja: "CMYK", en: "CMYK" },
            paintFill: { ja: "塗り", en: "Fill" },
            paintStroke: { ja: "線", en: "Stroke" },
            targetIndividual: { ja: "個別", en: "Individually" },
            targetGroup: { ja: "グループとして", en: "As a Group" }
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
            linkMargins: { ja: "左右にも上下と同じ値を使います", en: "Use the vertical value for the horizontal margin too" },
            roundEnable: { ja: "角丸を適用します", en: "Apply rounded corners" },
            pillShape: {
                ja: "高さの半分を角丸の半径にします（左右のマージンにも反映）",
                en: "Use half the height as the corner radius (also copied to the horizontal margin)"
            },
            colorBlack: { ja: "ショートカット：K", en: "Shortcut: K" },
            colorWhite: { ja: "ショートカット：W", en: "Shortcut: W" },
            colorHex: { ja: "ショートカット：H", en: "Shortcut: H" },
            colorCmyk: { ja: "ショートカット：C", en: "Shortcut: C" },
            hexInput: {
                ja: "#RRGGBB、R,G,B、色名（red など）で指定します",
                en: "Enter #RRGGBB, R,G,B, or a color name (e.g. red)"
            },
            hexInvalid: { ja: "正しい #RRGGBB を入力してください", en: "Enter a valid #RRGGBB value" },
            hexEmpty: { ja: "HEX未入力（# のみ）", en: "HEX not entered (# only)" },
            cmykRange: {
                ja: "0〜100 の範囲にしてください（未入力は 0 として扱います）",
                en: "Enter a value from 0 to 100 (empty is treated as 0)"
            },
            targetIndividual: {
                ja: "選択したオブジェクトごとに作成します（ショートカット：I）",
                en: "Create one rectangle per selected object (shortcut: I)"
            },
            targetGroup: {
                ja: "選択全体を囲む1つを作成します（ショートカット：G）",
                en: "Create one rectangle around the whole selection (shortcut: G)"
            },
            groupWithObjects: {
                ja: "作成した長方形を、対象のオブジェクトとグループ化します",
                en: "Group each rectangle with the objects it was drawn behind"
            },
            opacityApply: { ja: "OFF のときは 100% で作成します", en: "When off, the rectangle is created at 100%" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        if (labelSet[uiLang] != null) return labelSet[uiLang];
        return (labelSet.en != null) ? labelSet.en : "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} コロン付きラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 共通ヘルパー / Common helpers
    // =========================================

    /**
     * DEBUG_MODE のときだけエラーをコンソールに出す
     * @param {string} context - 発生箇所
     * @param {Error} e - 例外
     * @returns {void}
     */
    function logError(context, e) {
        if (!DEBUG_MODE) return;
        $.writeln("[ERROR] " + context + ": " + e);
    }

    /**
     * プロパティを代入する。種類やバージョンによって持たないプロパティがあるため失敗は無視する
     * @param {Object} target - 代入先（DOM オブジェクトや ScriptUI コントロール）
     * @param {string} propName - プロパティ名
     * @param {*} value - 値
     * @returns {boolean} 代入できたら true
     */
    function trySetProperty(target, propName, value) {
        try {
            target[propName] = value;
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 数値を範囲内に収める
     * @param {number} value - 値
     * @param {number} min - 下限
     * @param {number} max - 上限
     * @returns {number} 範囲内の値
     */
    function clampNumber(value, min, max) {
        return value < min ? min : (value > max ? max : value);
    }

    /**
     * 数値ならその値、そうでなければ既定値を返す
     * @param {*} value - 値
     * @param {number} fallback - 既定値
     * @returns {number} 数値
     */
    function numberOr(value, fallback) {
        return (typeof value === 'number') ? value : fallback;
    }

    /**
     * 入力欄の文字列を現在の定規単位として pt に換算する（負や不正な値は 0）
     * @param {string} fieldText - 入力欄の文字列
     * @param {number} pointsPerUnit - 1単位あたりの pt
     * @returns {number} pt 値
     */
    function fieldTextToPt(fieldText, pointsPerUnit) {
        var n = parseFloat(String(fieldText == null ? '' : fieldText));
        if (isNaN(n) || n < 0) n = 0;
        return n * pointsPerUnit;
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

    // =========================================
    // カラー / Colors
    // =========================================

    /**
     * RGBColor を作る（0〜255 に丸める）
     * @param {number} r - 赤
     * @param {number} g - 緑
     * @param {number} b - 青
     * @returns {RGBColor} カラー
     */
    function makeRGB(r, g, b) {
        var color = new RGBColor();
        color.red = clampNumber(Math.round(r), 0, 255);
        color.green = clampNumber(Math.round(g), 0, 255);
        color.blue = clampNumber(Math.round(b), 0, 255);
        return color;
    }

    /**
     * CMYKColor を作る（0〜100 に収める）
     * @param {number} c - シアン
     * @param {number} m - マゼンタ
     * @param {number} y - イエロー
     * @param {number} k - ブラック
     * @returns {CMYKColor} カラー
     */
    function makeCMYK(c, m, y, k) {
        var color = new CMYKColor();
        color.cyan = clampNumber(c, 0, 100);
        color.magenta = clampNumber(m, 0, 100);
        color.yellow = clampNumber(y, 0, 100);
        color.black = clampNumber(k, 0, 100);
        return color;
    }

    /**
     * ドキュメントのカラーモードが RGB か
     * @param {Document} doc - ドキュメント
     * @returns {boolean} RGB なら true
     */
    function isRgbDocument(doc) {
        return doc.documentColorSpace == DocumentColorSpace.RGB;
    }

    /**
     * ドキュメントのカラーモードに合わせて RGB か CMYK のカラーを作る
     * @param {Document} doc - ドキュメント
     * @param {number[]} rgb - [R, G, B]
     * @param {number[]} cmyk - [C, M, Y, K]
     * @returns {RGBColor|CMYKColor} カラー
     */
    function makeDocColor(doc, rgb, cmyk) {
        if (isRgbDocument(doc)) return makeRGB(rgb[0], rgb[1], rgb[2]);
        return makeCMYK(cmyk[0], cmyk[1], cmyk[2], cmyk[3]);
    }

    /**
     * RGB（0〜255）を CMYK（0〜100）に換算する
     * @param {number} r - 赤
     * @param {number} g - 緑
     * @param {number} b - 青
     * @returns {number[]} [C, M, Y, K]
     */
    function rgbToCmyk(r, g, b) {
        r = clampNumber(r, 0, 255) / 255;
        g = clampNumber(g, 0, 255) / 255;
        b = clampNumber(b, 0, 255) / 255;
        var k = 1 - Math.max(r, g, b);
        if (k >= 0.9999) return [0, 0, 0, 100];
        var c = (1 - r - k) / (1 - k);
        var m = (1 - g - k) / (1 - k);
        var y = (1 - b - k) / (1 - k);
        return [Math.round(c * 100), Math.round(m * 100), Math.round(y * 100), Math.round(k * 100)];
    }

    /**
     * CMYK（0〜100）を RGB（0〜255）に換算する
     * @param {number} c - シアン
     * @param {number} m - マゼンタ
     * @param {number} y - イエロー
     * @param {number} k - ブラック
     * @returns {number[]} [R, G, B]
     */
    function cmykToRgb(c, m, y, k) {
        c = clampNumber(c, 0, 100) / 100;
        m = clampNumber(m, 0, 100) / 100;
        y = clampNumber(y, 0, 100) / 100;
        k = clampNumber(k, 0, 100) / 100;
        return [
            Math.round(255 * (1 - c) * (1 - k)),
            Math.round(255 * (1 - m) * (1 - k)),
            Math.round(255 * (1 - y) * (1 - k))
        ];
    }

    /**
     * 全角の数字・区切り文字を半角にそろえ、小文字にする
     * @param {string} text - 入力文字列
     * @returns {string} 正規化した文字列
     */
    function normalizeColorText(text) {
        var s = String(text).replace(/^\s+|\s+$/g, '').toLowerCase();
        s = s.replace(/　/g, ' ');
        s = s.replace(/[，、]/g, ',');
        s = s.replace(/[０-９]/g, function (ch) {
            return String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30);
        });
        s = s.replace(/．/g, '.');
        s = s.replace(/／/g, '/');
        return s;
    }

    /**
     * HEX 欄の文字列をカラーにする
     * 受け付ける形式: "#RRGGBB" / "RRGGBB" / "R,G,B"（0〜255）/ 色名 / grayNN（0〜100）
     * @param {Document} doc - ドキュメント
     * @param {string} customValue - 入力文字列
     * @returns {RGBColor|CMYKColor|null} カラー。解釈できなければ null
     */
    function parseCustomColor(doc, customValue) {
        if (!customValue) return null;
        var s = normalizeColorText(customValue);
        if (!s) return null;

        /* #RRGGBB / RRGGBB */
        var hexDigits = null;
        if (s.charAt(0) === '#' && s.length === 7) hexDigits = s.substr(1);
        else if (/^[0-9a-f]{6}$/.test(s)) hexDigits = s;
        if (hexDigits) {
            var r = parseInt(hexDigits.substr(0, 2), 16);
            var g = parseInt(hexDigits.substr(2, 2), 16);
            var b = parseInt(hexDigits.substr(4, 2), 16);
            if (!isNaN(r) && !isNaN(g) && !isNaN(b)) return makeRGB(r, g, b);
        }

        /* R,G,B（0〜255）/ Comma-separated RGB */
        var rgbCsv = s.match(/^\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*$/);
        if (rgbCsv) {
            return makeRGB(parseInt(rgbCsv[1], 10), parseInt(rgbCsv[2], 10), parseInt(rgbCsv[3], 10));
        }

        /* 色名（ドキュメントのカラーモードに合わせる）/ Named colors follow the document color mode */
        if (NAMED_COLORS.hasOwnProperty(s)) {
            return isRgbDocument(doc) ?
                makeRGB(NAMED_COLORS[s].rgb[0], NAMED_COLORS[s].rgb[1], NAMED_COLORS[s].rgb[2]) :
                makeCMYK(NAMED_COLORS[s].cmyk[0], NAMED_COLORS[s].cmyk[1], NAMED_COLORS[s].cmyk[2], NAMED_COLORS[s].cmyk[3]);
        }

        /* grayNN（0〜100）/ Gray percentage */
        var grayMatch = s.match(/^gray\s*(\d{1,3})$/);
        if (grayMatch) {
            var grayPercent = clampNumber(parseInt(grayMatch[1], 10), 0, 100);
            var level = Math.round(255 * (100 - grayPercent) / 100);
            return makeDocColor(doc, [level, level, level], [0, 0, 0, grayPercent]);
        }
        return null;
    }

    /**
     * 設定からカラーを決める（塗りにも線にも使う）
     * @param {Document} doc - ドキュメント
     * @param {Object} choice - ダイアログの設定
     * @returns {RGBColor|CMYKColor|null} カラー。無しなら null
     */
    function resolveChoiceColor(doc, choice) {
        var mode = choice.colorMode;
        if (mode === ColorMode.K100) return makeDocColor(doc, [0, 0, 0], [0, 0, 0, 100]);
        if (mode === ColorMode.WHITE) return makeDocColor(doc, [255, 255, 255], [0, 0, 0, 0]);
        if (mode === ColorMode.HEX) {
            var color = parseCustomColor(doc, choice.customValue);
            /* CMYK ドキュメントでは RGB を CMYK に換算 / Convert RGB to CMYK in CMYK documents */
            if (color && color.typename === 'RGBColor' && doc.documentColorSpace == DocumentColorSpace.CMYK) {
                var cmyk = rgbToCmyk(color.red | 0, color.green | 0, color.blue | 0);
                return makeCMYK(cmyk[0], cmyk[1], cmyk[2], cmyk[3]);
            }
            return color;
        }
        if (mode === ColorMode.CMYK) {
            var values = choice.customCMYK || {};
            var keys = ['c', 'm', 'y', 'k'];
            for (var i = 0; i < keys.length; i++) {
                if (typeof values[keys[i]] !== 'number' || isNaN(values[keys[i]])) return null;
            }
            if (isRgbDocument(doc)) {
                var rgb = cmykToRgb(values.c, values.m, values.y, values.k);
                return makeRGB(rgb[0], rgb[1], rgb[2]);
            }
            return makeCMYK(values.c, values.m, values.y, values.k);
        }
        return null;
    }

    // =========================================
    // 長方形の見た目 / Rectangle appearance
    // =========================================

    /**
     * オーバープリントを外し、描画モードを通常にする
     * @param {PathItem} rect - 長方形
     * @returns {void}
     */
    function resetOverprintAndBlend(rect) {
        trySetProperty(rect, "fillOverprint", false);
        trySetProperty(rect, "strokeOverprint", false);
        trySetProperty(rect, "blendingMode", BlendingMode.NORMAL);
    }

    /**
     * 塗りを設定し、線を外す（不透明度は触らない）
     * @param {PathItem} rect - 長方形
     * @param {RGBColor|CMYKColor|null} color - 塗りの色。null なら塗りなし
     * @returns {void}
     */
    function applyFill(rect, color) {
        try {
            if (color) {
                rect.filled = true;
                rect.fillColor = color;
            } else {
                rect.filled = false;
            }
            rect.stroked = false;
        } catch (e) {
            logError("applyFill", e);
        }
        resetOverprintAndBlend(rect);
    }

    /**
     * 設定の線幅（pt）を返す。0 以下や未指定なら 1
     * @param {Object} choice - ダイアログの設定
     * @returns {number} 線幅（pt）
     */
    function getStrokeWidthPt(choice) {
        return (typeof choice.strokeWidth === 'number' && choice.strokeWidth > 0) ? choice.strokeWidth : 1;
    }

    /**
     * 種別（塗り／線）に応じて色を付ける
     * @param {PathItem} rect - 長方形
     * @param {Object} choice - ダイアログの設定
     * @param {RGBColor|CMYKColor|null} color - 色
     * @returns {void}
     */
    function applyRectPaint(rect, choice, color) {
        if (choice.type !== 'stroke') {
            applyFill(rect, color);
            return;
        }
        try {
            rect.filled = false;
            rect.stroked = !!color;
            if (color) rect.strokeColor = color;
            rect.strokeWidth = getStrokeWidthPt(choice);
        } catch (e) {
            logError("applyRectPaint", e);
        }
    }

    /**
     * 設定の不透明度（0〜100）を適用する
     * @param {PathItem} rect - 長方形
     * @param {Object} choice - ダイアログの設定
     * @returns {void}
     */
    function applyRectOpacity(rect, choice) {
        rect.opacity = clampNumber(Math.round(numberOr(choice.opacity, 100)), 0, 100);
    }

    /**
     * ライブエフェクトを適用する（展開しない）
     * @param {PageItem} item - 対象
     * @param {string} effectName - エフェクト名
     * @param {string} dictData - Dict の data 属性
     * @returns {void}
     */
    function applyLiveEffect(item, effectName, dictData) {
        try {
            item.applyEffect(
                '<LiveEffect name="' + effectName + '">' +
                '<Dict data="' + dictData + '"/>' +
                '</LiveEffect>'
            );
        } catch (e) {
            logError("applyLiveEffect", e);
        }
    }

    /**
     * 角丸をライブエフェクトで付ける。ピル形状なら高さの半分を半径にする
     * @param {PathItem} rect - 長方形
     * @param {Object} choice - ダイアログの設定
     * @param {number} rectHeight - マージン込みの高さ（pt）
     * @returns {void}
     */
    function applyCornerEffect(rect, choice, rectHeight) {
        var radius = choice.isPill ? rectHeight / 2 : numberOr(choice.roundPt, 0);
        if (radius > 0) applyLiveEffect(rect, "Adobe Round Corners", "R radius " + radius + " ");
    }

    // =========================================
    // 選択とレイヤー / Selection and layers
    // =========================================

    /* 今回ダイアログを開いたときの選択（プレビューはこれを基準にする）/ Selection captured when the dialog opened */
    var sessionSelection = [];

    /**
     * 選択を配列に写し取る
     * @param {Document} doc - ドキュメント
     * @returns {PageItem[]} 選択オブジェクト
     */
    function snapshotSelection(doc) {
        var items = [];
        try {
            var selection = doc.selection || [];
            for (var i = 0; i < selection.length; i++) items.push(selection[i]);
        } catch (e) {
            logError("snapshotSelection", e);
        }
        return items;
    }

    /**
     * プレビューの対象にする選択を返す（ダイアログを開いたときの選択を優先）
     * @param {Document} doc - ドキュメント
     * @returns {PageItem[]} 対象オブジェクト
     */
    function getPreviewSelection(doc) {
        if (sessionSelection && sessionSelection.length) return sessionSelection;
        try {
            return doc.selection || [];
        } catch (e) {
            return [];
        }
    }

    /**
     * オブジェクトの属するレイヤーを返す。取れなければアクティブレイヤー
     * @param {PageItem} item - オブジェクト
     * @param {Document} doc - ドキュメント
     * @returns {Layer} レイヤー
     */
    function getItemLayer(item, doc) {
        var layer = null;
        try {
            layer = item.layer;
        } catch (e) {
            layer = null;
        }
        return layer || doc.activeLayer;
    }

    /**
     * 選択の代表レイヤー（先頭オブジェクトのレイヤー）を返す
     * @param {Document} doc - ドキュメント
     * @param {PageItem[]} items - 選択オブジェクト
     * @returns {Layer} レイヤー
     */
    function getRepresentativeLayer(doc, items) {
        if (!items || !items.length) return doc.activeLayer;
        return getItemLayer(items[0], doc);
    }

    /**
     * 選択オブジェクトの属するレイヤーを重複なく集める（名前と型で同一視）
     * @param {PageItem[]} items - 選択オブジェクト
     * @returns {Layer[]} レイヤー
     */
    function getUniqueLayersFromSelection(items) {
        var layers = [];
        var seen = {};
        for (var i = 0; i < items.length; i++) {
            var layer = null;
            try {
                layer = items[i].layer;
            } catch (e) {
                layer = null;
            }
            if (!layer) continue;
            var key = (layer.name || "") + "#" + (layer.typename || "");
            if (!seen[key]) {
                seen[key] = true;
                layers.push(layer);
            }
        }
        return layers;
    }

    /**
     * レイヤーと親レイヤーを表示・ロック解除し、アクティブにする
     * @param {Document} doc - ドキュメント
     * @param {Layer} layer - レイヤー
     * @returns {void}
     */
    function ensureLayerEditable(doc, layer) {
        if (!doc || !layer) return;
        trySetProperty(layer, "visible", true);
        trySetProperty(layer, "locked", false);
        var parentLayer = layer.parent;
        while (parentLayer && parentLayer.typename === 'Layer') {
            trySetProperty(parentLayer, "visible", true);
            trySetProperty(parentLayer, "locked", false);
            parentLayer = parentLayer.parent;
        }
        /* 環境によってはアクティブでないと追加できない / Some environments need an active layer for insertion */
        trySetProperty(doc, "activeLayer", layer);
    }

    // =========================================
    // 境界の計測 / Bounds measuring
    // =========================================

    /**
     * 名前でレイヤーを探す（最上位のみ）
     * @param {Document} doc - ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} レイヤー
     */
    function findLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return doc.layers[i];
        }
        return null;
    }

    /**
     * アウトライン計測用の一時レイヤーを取得／作成する（正しい境界を得るため表示・編集可能にする）
     * @param {Document} doc - ドキュメント
     * @returns {Layer|null} 一時レイヤー
     */
    function getOrCreateTempOutlineLayer(doc) {
        try {
            var layer = findLayerByName(doc, TEMP_OUTLINE_LAYER_NAME);
            if (!layer) {
                layer = doc.layers.add();
                layer.name = TEMP_OUTLINE_LAYER_NAME;
            }
            layer.visible = true;
            layer.locked = false;
            trySetProperty(layer, "printable", false);
            try {
                layer.move(doc, ElementPlacement.PLACEATEND);
            } catch (e) {
                logError("getOrCreateTempOutlineLayer.move", e);
            }
            return layer;
        } catch (e) {
            logError("getOrCreateTempOutlineLayer", e);
            return null;
        }
    }

    /**
     * 空になった一時レイヤーを削除する
     * @param {Document} doc - ドキュメント
     * @returns {void}
     */
    function removeTempOutlineLayer(doc) {
        try {
            var layer = findLayerByName(doc, TEMP_OUTLINE_LAYER_NAME);
            if (layer && layer.pageItems.length === 0 && layer.layers.length === 0) layer.remove();
        } catch (e) {
            logError("removeTempOutlineLayer", e);
        }
    }

    /**
     * オブジェクトの境界（線幅込み）を返す
     * @param {PageItem} item - オブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]
     */
    function getItemBounds(item) {
        try {
            if (!item) return null;
            return item.visibleBounds || item.geometricBounds;
        } catch (e) {
            return null;
        }
    }

    /**
     * メニューコマンドでテキストをアウトライン化する（createOutline() が使えないときの予備）
     * @param {Document} doc - ドキュメント
     * @param {TextFrame} textFrame - 複製したテキスト
     * @returns {PageItem|null} アウトライン化したオブジェクト
     */
    function outlineTextByMenu(doc, textFrame) {
        var previousSelection = snapshotSelection(doc);
        var outlined = null;
        try {
            app.executeMenuCommand('deselectall');
            textFrame.selected = true;
            app.executeMenuCommand('createOutlines');
            outlined = (doc.selection && doc.selection.length) ? doc.selection[0] : null;
        } catch (e) {
            logError("outlineTextByMenu", e);
        }
        /* 選択を戻す / Restore selection */
        try {
            app.executeMenuCommand('deselectall');
            for (var i = 0; i < previousSelection.length; i++) previousSelection[i].selected = true;
        } catch (e) {
            logError("outlineTextByMenu.restore", e);
        }
        return outlined;
    }

    /**
     * テキストを複製してアウトライン化し、字形の境界を測る
     * @param {Document} doc - ドキュメント
     * @param {TextFrame} textFrame - テキスト
     * @returns {number[]|null} [左, 上, 右, 下]
     */
    function getTextOutlinedBounds(doc, textFrame) {
        try {
            if (!textFrame || textFrame.typename !== 'TextFrame') return null;
            var tempLayer = getOrCreateTempOutlineLayer(doc);
            ensureLayerEditable(doc, tempLayer);

            var duplicated = textFrame.duplicate(tempLayer, ElementPlacement.PLACEATBEGINNING);
            trySetProperty(duplicated, "hidden", false);
            trySetProperty(duplicated, "locked", false);

            var outlined = null;
            try {
                outlined = duplicated.createOutline();
            } catch (e) {
                outlined = null;
            }
            if (!outlined) outlined = outlineTextByMenu(doc, duplicated);

            var bounds = null;
            if (outlined) {
                try {
                    bounds = outlined.visibleBounds || outlined.geometricBounds;
                } catch (e) {
                    bounds = null;
                }
            }

            /* 後片付け。createOutline() は複製を消費するので remove() が例外になることがある
               Cleanup; createOutline() consumes the duplicate, so remove() may throw */
            try {
                if (outlined) outlined.remove();
            } catch (e) {}
            try {
                duplicated.remove();
            } catch (e) {}
            return bounds;
        } catch (e) {
            logError("getTextOutlinedBounds", e);
            return null;
        }
    }

    /**
     * 最終的な境界を返す（テキストはアウトラインの境界）
     * @param {Document} doc - ドキュメント
     * @param {PageItem} item - オブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]
     */
    function getFinalItemBounds(doc, item) {
        try {
            if (!item) return null;
            if (item.typename === 'TextFrame') {
                var textBounds = getTextOutlinedBounds(doc, item);
                if (textBounds) return textBounds;
            }
            return getItemBounds(item);
        } catch (e) {
            return null;
        }
    }

    /**
     * 複数オブジェクトの最終的な境界を合わせた外接矩形を返す
     * @param {Document} doc - ドキュメント
     * @param {PageItem[]} items - オブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]
     */
    function getCombinedFinalBounds(doc, items) {
        if (!items || !items.length) return null;
        var combined = null;
        for (var i = 0; i < items.length; i++) {
            var bounds = getFinalItemBounds(doc, items[i]);
            if (!bounds) continue;
            if (!combined) {
                combined = [bounds[0], bounds[1], bounds[2], bounds[3]];
                continue;
            }
            if (bounds[0] < combined[0]) combined[0] = bounds[0];
            if (bounds[1] > combined[1]) combined[1] = bounds[1];
            if (bounds[2] > combined[2]) combined[2] = bounds[2];
            if (bounds[3] < combined[3]) combined[3] = bounds[3];
        }
        return combined;
    }

    // =========================================
    // 境界のキャッシュ / Bounds cache
    // =========================================

    /* アウトライン計測の結果を選択ごとに控える（選択が変わったら作り直す）
       Outline bounds cached per selection; rebuilt when the selection signature changes */
    var outlineCache = { signature: "", groupBounds: null, itemBounds: {} };

    /**
     * 選択の識別文字列（個数・名前・境界）を作る
     * @param {PageItem[]} items - 選択オブジェクト
     * @returns {string} 識別文字列
     */
    function buildSelectionSignature(items) {
        var parts = [items ? items.length : 0];
        for (var i = 0; i < (items ? items.length : 0); i++) {
            var itemName = "";
            var itemBounds = "";
            try {
                itemName = String(items[i].name || "");
            } catch (e) {}
            /* 空白だけの文字グループなどは geometricBounds が例外になる / may throw for empty text groups */
            try {
                itemBounds = String(items[i].geometricBounds);
            } catch (e) {}
            parts.push(itemName + "|" + itemBounds);
        }
        return parts.join(";");
    }

    /**
     * 選択が変わっていたらキャッシュを作り直す
     * @param {PageItem[]} items - 選択オブジェクト
     * @returns {void}
     */
    function syncOutlineCache(items) {
        var signature = buildSelectionSignature(items);
        if (outlineCache.signature !== signature) {
            outlineCache = { signature: signature, groupBounds: null, itemBounds: {} };
        }
    }

    /**
     * 選択の index 番目の最終境界をキャッシュ経由で返す
     * @param {Document} doc - ドキュメント
     * @param {PageItem} item - オブジェクト
     * @param {number} index - 選択内の番号
     * @returns {number[]|null} [左, 上, 右, 下]
     */
    function getFinalItemBoundsCached(doc, item, index) {
        var key = String(index);
        if (outlineCache.itemBounds.hasOwnProperty(key)) return outlineCache.itemBounds[key];
        var bounds = getFinalItemBounds(doc, item);
        outlineCache.itemBounds[key] = bounds;
        return bounds;
    }

    /**
     * 選択全体の最終境界をキャッシュ経由で返す
     * @param {Document} doc - ドキュメント
     * @param {PageItem[]} items - 選択オブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]
     */
    function getCombinedFinalBoundsCached(doc, items) {
        if (outlineCache.groupBounds) return outlineCache.groupBounds;
        var bounds = getCombinedFinalBounds(doc, items);
        outlineCache.groupBounds = bounds;
        return bounds;
    }

    /**
     * [左, 上, 右, 下] を {left, top, width, height} にする
     * @param {number[]} bounds - 境界
     * @returns {{left: number, top: number, width: number, height: number}|null} 矩形の指定
     */
    function boundsToRectSpec(bounds) {
        if (!bounds || bounds.length !== 4) return null;
        return {
            left: bounds[0],
            top: bounds[1],
            width: bounds[2] - bounds[0],
            height: bounds[1] - bounds[3]
        };
    }

    /**
     * 対象（グループ／個別）ごとの矩形指定を visitor に渡す
     * @param {Document} doc - ドキュメント
     * @param {PageItem[]} items - 選択オブジェクト
     * @param {Object} choice - ダイアログの設定
     * @param {Function} visitor - function(rectSpec, index)。グループは index 0
     * @returns {void}
     */
    function forEachTargetRectSpec(doc, items, choice, visitor) {
        try {
            if (choice.target === 'group') {
                var groupSpec = boundsToRectSpec(getCombinedFinalBoundsCached(doc, items));
                if (groupSpec) visitor(groupSpec, 0);
                return;
            }
            for (var i = 0; i < items.length; i++) {
                if (!items[i]) continue;
                var itemSpec = boundsToRectSpec(getFinalItemBoundsCached(doc, items[i], i));
                if (itemSpec) visitor(itemSpec, i);
            }
        } catch (e) {
            logError("forEachTargetRectSpec", e);
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビュー用のレイヤー名か（旧版の名前を含む）
     * @param {string} layerName - レイヤー名
     * @returns {boolean} プレビュー用なら true
     */
    function isPreviewLayerName(layerName) {
        for (var i = 0; i < PREVIEW_LAYER_NAMES.length; i++) {
            if (layerName === PREVIEW_LAYER_NAMES[i]) return true;
        }
        return false;
    }

    /**
     * プレビューレイヤーを探す
     * @param {Document} doc - ドキュメント
     * @returns {Layer|null} プレビューレイヤー
     */
    function findPreviewLayer(doc) {
        try {
            for (var i = 0; i < doc.layers.length; i++) {
                if (isPreviewLayerName(doc.layers[i].name)) return doc.layers[i];
            }
        } catch (e) {
            logError("findPreviewLayer", e);
        }
        return null;
    }

    /**
     * プレビューを片付ける。removeLayer が false なら長方形を隠すだけ、true ならレイヤーごと削除する
     * @param {Document} doc - ドキュメント
     * @param {boolean} removeLayer - レイヤーを削除するか
     * @returns {void}
     */
    function clearPreview(doc, removeLayer) {
        if (!doc) return;
        try {
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var layer = doc.layers[i];
                if (!isPreviewLayerName(layer.name)) continue;
                if (removeLayer) {
                    try {
                        layer.remove();
                    } catch (e) {
                        logError("clearPreview.remove", e);
                    }
                } else {
                    for (var k = layer.pathItems.length - 1; k >= 0; k--) {
                        trySetProperty(layer.pathItems[k], "hidden", true);
                    }
                }
            }
        } catch (e) {
            logError("clearPreview", e);
        }
        if (removeLayer) removeTempOutlineLayer(doc);
    }

    /**
     * プレビューレイヤーを取得／作成し、基準レイヤーの直下に置く
     * @param {Document} doc - ドキュメント
     * @param {Layer} referenceLayer - 基準レイヤー
     * @returns {Layer} プレビューレイヤー
     */
    function getOrCreatePreviewLayer(doc, referenceLayer) {
        var layer = findLayerByName(doc, PREVIEW_LAYER_NAME);
        if (!layer) {
            layer = doc.layers.add();
            layer.name = PREVIEW_LAYER_NAME;
        }
        layer.visible = true;
        layer.locked = false;
        /* 描画モードなどの影響を受けないようにする / Render at full strength */
        trySetProperty(layer, "opacity", 100);
        trySetProperty(layer, "blendingMode", BlendingMode.NORMAL);
        trySetProperty(layer, "transparencyIsolated", false);
        trySetProperty(layer, "transparencyKnockoutGroup", false);
        try {
            if (referenceLayer && referenceLayer.typename === 'Layer') {
                layer.move(referenceLayer, ElementPlacement.PLACEAFTER);
            } else {
                layer.move(doc, ElementPlacement.PLACEATEND);
            }
        } catch (e) {
            logError("getOrCreatePreviewLayer.move", e);
        }
        return layer;
    }

    /**
     * プレビュー矩形の名前（"<接頭辞>#<番号>"）を作る
     * @param {number} index - 番号（グループは 0）
     * @returns {string} 名前
     */
    function makePreviewRectName(index) {
        return getLabel(PREVIEW_RECT_BASE_NAMES) + "#" + String(index | 0);
    }

    /**
     * プレビュー矩形の名前から番号を取り出す
     * @param {string} itemName - 名前
     * @returns {number|null} 番号。プレビュー矩形でなければ null
     */
    function parsePreviewIndex(itemName) {
        var match = /^(.*)#(\d+)$/.exec(String(itemName || ''));
        if (!match) return null;
        if (match[1] !== PREVIEW_RECT_BASE_NAMES.ja && match[1] !== PREVIEW_RECT_BASE_NAMES.en) return null;
        return parseInt(match[2], 10);
    }

    /**
     * プレビュー矩形を作り直す（ライブエフェクトが重ならないよう同名の旧矩形は削除）
     * @param {Layer} previewLayer - プレビューレイヤー
     * @param {number} index - 番号
     * @param {number} top - 上端
     * @param {number} left - 左端
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @returns {PathItem} 長方形
     */
    function createPreviewRect(previewLayer, index, top, left, width, height) {
        var rectName = makePreviewRectName(index);
        try {
            for (var i = previewLayer.pathItems.length - 1; i >= 0; i--) {
                if (String(previewLayer.pathItems[i].name || '') === rectName) {
                    previewLayer.pathItems[i].remove();
                    break;
                }
            }
        } catch (e) {
            logError("createPreviewRect.remove", e);
        }
        var rect = previewLayer.pathItems.rectangle(top, left, width, height);
        rect.name = rectName;
        return rect;
    }

    /**
     * プレビュー矩形を描く（マージン・色・不透明度・角丸を反映）
     * @param {Layer} previewLayer - プレビューレイヤー
     * @param {number} index - 番号（グループは 0）
     * @param {Object} rectSpec - {left, top, width, height}
     * @param {Object} choice - ダイアログの設定
     * @param {Document} doc - ドキュメント
     * @returns {PathItem} 長方形
     */
    function buildPreviewRect(previewLayer, index, rectSpec, choice, doc) {
        var offsetV = numberOr(choice.offsetV, 0);
        var offsetH = numberOr(choice.offsetH, 0);
        var rectHeight = rectSpec.height + offsetV * 2;
        var rect = createPreviewRect(previewLayer, index,
            rectSpec.top + offsetV, rectSpec.left - offsetH, rectSpec.width + offsetH * 2, rectHeight);

        applyRectPaint(rect, choice, resolveChoiceColor(doc, choice));
        if (choice.type !== 'stroke') {
            /* ホワイトは見えにくいので、プレビューだけグレーの補助線を付ける / Gray helper stroke for White */
            try {
                if (choice.colorMode === ColorMode.WHITE) {
                    rect.stroked = true;
                    rect.strokeColor = makeDocColor(doc, [128, 128, 128], [0, 0, 0, 50]);
                    rect.strokeWidth = 0.5;
                } else {
                    rect.stroked = false;
                }
            } catch (e) {
                logError("buildPreviewRect.helperStroke", e);
            }
        }
        resetOverprintAndBlend(rect);
        applyRectOpacity(rect, choice);
        applyCornerEffect(rect, choice, rectHeight);
        rect.selected = false;
        rect.zOrder(ZOrderMethod.SENDTOBACK);
        return rect;
    }

    /**
     * プレビューを描き直す
     * @param {Document} doc - ドキュメント
     * @param {Object} choice - ダイアログの設定
     * @returns {void}
     */
    function renderPreview(doc, choice) {
        clearPreview(doc, false);
        if (!doc || !choice) return;

        var previousCoordinateSystem = app.coordinateSystem;
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;

        var items = getPreviewSelection(doc);
        syncOutlineCache(items);
        var previewLayer = getOrCreatePreviewLayer(doc, getRepresentativeLayer(doc, items));
        forEachTargetRectSpec(doc, items, choice, function (rectSpec, index) {
            buildPreviewRect(previewLayer, index, rectSpec, choice, doc);
        });

        app.coordinateSystem = previousCoordinateSystem;
        app.redraw();
    }

    // =========================================
    // 確定 / Finalize
    // =========================================

    /**
     * プレビューレイヤーの長方形を番号ごとに集める
     * @param {Layer} previewLayer - プレビューレイヤー
     * @returns {Object} 番号 → PathItem
     */
    function collectPreviewRectsByIndex(previewLayer) {
        var rectsByIndex = {};
        try {
            for (var i = 0; i < previewLayer.pathItems.length; i++) {
                var rect = previewLayer.pathItems[i];
                var index = parsePreviewIndex(rect.name);
                if (index == null || isNaN(index)) index = i; /* 名前で取れなければ並び順 / fall back to order */
                rectsByIndex[index] = rect;
            }
        } catch (e) {
            logError("collectPreviewRectsByIndex", e);
        }
        return rectsByIndex;
    }

    /**
     * 確定用の見た目にする（プレビューの補助線を外し、不透明度を適用）
     * @param {PathItem} rect - 長方形
     * @param {Object} choice - ダイアログの設定
     * @param {RGBColor|CMYKColor|null} color - 色
     * @returns {void}
     */
    function applyFinalStyle(rect, choice, color) {
        applyRectPaint(rect, choice, color);
        applyRectOpacity(rect, choice);
    }

    /**
     * 長方形を表示して名前を付け、レイヤーの最背面へ移す
     * @param {PathItem} rect - 長方形
     * @param {Layer} layer - 移動先レイヤー
     * @returns {void}
     */
    function moveRectToLayerBack(rect, layer) {
        trySetProperty(rect, "hidden", false);
        trySetProperty(rect, "name", FINAL_RECT_NAME);
        /* 移動できなくても最背面へは送る / Still send to back if the move fails */
        try {
            rect.move(layer, ElementPlacement.PLACEATBEGINNING);
        } catch (e) {
            logError("moveRectToLayerBack.move", e);
        }
        try {
            rect.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (e) {
            logError("moveRectToLayerBack.zOrder", e);
        }
    }

    /**
     * 長方形とオブジェクトを新しいグループにまとめる（長方形は最背面）
     * @param {Layer} layer - グループを作るレイヤー
     * @param {PathItem} rect - 長方形
     * @param {PageItem[]} items - まとめるオブジェクト
     * @returns {void}
     */
    function groupRectWithItems(layer, rect, items) {
        try {
            var groupItem = layer.groupItems.add();
            for (var i = 0; i < items.length; i++) {
                try {
                    items[i].move(groupItem, ElementPlacement.PLACEATEND);
                } catch (e) {
                    logError("groupRectWithItems.item", e);
                }
            }
            rect.move(groupItem, ElementPlacement.PLACEATBEGINNING);
            rect.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (e) {
            logError("groupRectWithItems", e);
        }
    }

    /**
     * 「グループとして」の長方形を配置する。選択が複数レイヤーにまたがるときはレイヤーごとに複製する
     * @param {Document} doc - ドキュメント
     * @param {PathItem} rect - 長方形
     * @param {Layer[]} layers - 選択の属するレイヤー
     * @param {PageItem[]} items - 選択オブジェクト
     * @param {boolean} groupWithObjects - オブジェクトとグループ化するか
     * @returns {void}
     */
    function placeGroupRect(doc, rect, layers, items, groupWithObjects) {
        for (var i = 0; i < layers.length; i++) {
            if (!layers[i]) continue;
            ensureLayerEditable(doc, layers[i]);
            moveRectToLayerBack((i === 0) ? rect : rect.duplicate(), layers[i]);
        }
        /* グループ化は同じレイヤーにそろっているときだけ / Group only when everything is on one layer */
        if (layers.length === 1 && groupWithObjects) {
            groupRectWithItems(layers[0] || getRepresentativeLayer(doc, items), rect, items);
        }
    }

    /**
     * 「個別」の長方形を、それぞれのオブジェクトのレイヤーへ配置する
     * @param {Document} doc - ドキュメント
     * @param {PathItem[]} rects - 選択と同じ並びの長方形（無ければ null）
     * @param {PageItem[]} items - 選択オブジェクト
     * @param {boolean} groupWithObjects - オブジェクトとグループ化するか
     * @returns {void}
     */
    function placeItemRects(doc, rects, items, groupWithObjects) {
        for (var i = 0; i < items.length; i++) {
            if (!items[i] || !rects[i]) continue;
            var layer = getItemLayer(items[i], doc);
            ensureLayerEditable(doc, layer);
            moveRectToLayerBack(rects[i], layer);
            if (groupWithObjects) groupRectWithItems(layer, rects[i], [items[i]]);
        }
    }

    /**
     * プレビューの長方形をそのまま確定する
     * @param {Document} doc - ドキュメント
     * @param {PageItem[]} items - 選択オブジェクト
     * @param {Object} choice - ダイアログの設定
     * @returns {void}
     */
    function convertPreviewToFinal(doc, items, choice) {
        if (!doc || !choice) return;
        var previewLayer = findPreviewLayer(doc);
        if (!previewLayer) return;

        var rectsByIndex = collectPreviewRectsByIndex(previewLayer);
        var color = resolveChoiceColor(doc, choice);
        var groupWithObjects = !!choice.groupWithText;

        if (choice.target === 'group') {
            var groupRect = rectsByIndex[0];
            if (!groupRect) return;
            applyFinalStyle(groupRect, choice, color);
            var layers = getUniqueLayersFromSelection(items);
            if (!layers.length) layers = [getRepresentativeLayer(doc, items)];
            placeGroupRect(doc, groupRect, layers, items, groupWithObjects);
        } else {
            var rects = [];
            for (var i = 0; i < items.length; i++) {
                rects[i] = rectsByIndex[i] || null;
                if (rects[i]) applyFinalStyle(rects[i], choice, color);
            }
            placeItemRects(doc, rects, items, groupWithObjects);
        }

        try {
            previewLayer.remove();
        } catch (e) {
            logError("convertPreviewToFinal.remove", e);
        }
    }

    // =========================================
    // 前回の設定 / Last settings
    // =========================================

    /**
     * 前回の設定を保存するファイルを返す
     * @returns {File|null} 設定ファイル
     */
    function getStateFile() {
        try {
            var stateFolder = new Folder(Folder.userData.fsName + "/ai-scripts");
            if (!stateFolder.exists) stateFolder.create();
            return new File(stateFolder.fsName + "/" + STATE_FILE_NAME);
        } catch (e) {
            logError("getStateFile", e);
            return null;
        }
    }

    /**
     * 設定を保存する（toSource 形式）
     * @param {Object} choiceData - 保存する設定
     * @returns {void}
     */
    function saveLastChoice(choiceData) {
        try {
            var stateFile = getStateFile();
            if (stateFile && stateFile.open('w')) {
                stateFile.write(choiceData.toSource());
                stateFile.close();
            }
        } catch (e) {
            logError("saveLastChoice", e);
        }
    }

    /**
     * 前回の設定を読み込む
     * @returns {Object|null} 設定。無ければ null
     */
    function loadLastChoice() {
        try {
            var stateFile = getStateFile();
            if (!stateFile || !stateFile.exists || !stateFile.open('r')) return null;
            var source = stateFile.read();
            stateFile.close();
            var loaded = eval(source);
            return (loaded && typeof loaded === 'object') ? loaded : null;
        } catch (e) {
            logError("loadLastChoice", e);
            return null;
        }
    }

    /**
     * 保存する項目だけを取り出す（キー名は既存の保存ファイルとの互換のため据え置き）
     * @param {Object} choice - ダイアログの設定
     * @returns {Object} 保存用の設定
     */
    function serializeChoice(choice) {
        return {
            colorMode: choice.colorMode,
            customValue: choice.customValue,
            customCMYK: choice.customCMYK,
            offsetV: choice.offsetV,
            offsetH: choice.offsetH,
            roundPt: choice.roundPt,
            isPill: !!choice.isPill,
            type: choice.type,
            strokeWidth: choice.strokeWidth,
            target: choice.target,
            groupWithText: !!choice.groupWithText,
            opacity: clampNumber(numberOr(choice.opacity, 100), 0, 100),
            opacityEnabled: (typeof choice.opacityEnabled === 'boolean') ? choice.opacityEnabled : true
        };
    }

    // =========================================
    // 入力欄 / Input fields
    // =========================================

    /**
     * 入力欄の背景を黄色で強調する／戻す
     * @param {EditText} editText - 入力欄
     * @param {boolean} highlighted - 強調するか
     * @returns {void}
     */
    function setEditHighlight(editText, highlighted) {
        /* graphics の描画設定は環境によって失敗することがある / graphics may fail on some platforms */
        try {
            var gfx = editText.graphics;
            gfx.backgroundColor = gfx.newBrush(gfx.BrushType.SOLID_COLOR, highlighted ? [1, 1, 0.85] : [1, 1, 1]);
            gfx.foregroundColor = gfx.newPen(gfx.PenType.SOLID_COLOR, highlighted ? [0.2, 0.2, 0] : [0, 0, 0], 1);
            editText.notify('onDraw');
        } catch (e) {
            logError("setEditHighlight", e);
        }
    }

    /**
     * 入力欄の文字色を警告色（赤）にする／戻し、ツールチップを差し替える
     * @param {EditText} editText - 入力欄
     * @param {boolean} warn - 警告するか
     * @param {string} warnTip - 警告時のツールチップ
     * @param {string} normalTip - 通常時のツールチップ
     * @returns {void}
     */
    function setFieldWarning(editText, warn, warnTip, normalTip) {
        editText.helpTip = warn ? warnTip : normalTip;
        try {
            var gfx = editText.graphics;
            gfx.foregroundColor = gfx.newPen(gfx.PenType.SOLID_COLOR, warn ? [1, 0, 0] : [0, 0, 0], 1);
            editText.notify('onDraw');
        } catch (e) {
            logError("setFieldWarning", e);
        }
    }

    /**
     * 数値欄を、左に∧∨を付けて作る（↑↓キーも∧∨と同じ処理で増減する）。
     * 下限・上限と増減後の処理は、あとで bindNumericField() などが .stepOptions に入れる
     * @param {Group} parent - 親グループ
     * @param {string} initialText - 初期値
     * @param {boolean} [isInteger] - 整数の欄なら true（∧∨のツールチップが変わる）
     * @returns {EditText} 入力欄（∧∨は .stepperGroup、増減の設定は .stepOptions で参照できる）
     */
    function addNumberField(parent, initialText, isInteger) {
        var editText = addStepperEditText(parent, initialText, isInteger);
        editText.preferredSize = { width: NUMBER_FIELD_WIDTH, height: -1 };
        editText.characters = 3;
        return editText;
    }

    /**
     * 同じ行に∧∨と入力欄を隙間0で並べて追加する
     * @param {Group|Panel} parent - 追加先の行
     * @param {string} initialText - 初期値
     * @param {boolean} [isInteger] - 整数の欄なら true
     * @returns {EditText} 入力欄（∧∨は .stepperGroup、増減の設定は .stepOptions で参照できる）
     */
    function addStepperEditText(parent, initialText, isInteger) {
        var stepperInputGroup = parent.add('group');
        stepperInputGroup.orientation = 'row';
        stepperInputGroup.alignChildren = ['left', 'center'];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var editText;
        var stepOptions = { step: 1, min: 0, integer: !!isInteger };
        var stepperGroup = addStepper(stepperInputGroup, function () { return editText; }, stepOptions);
        editText = stepperInputGroup.add('edittext', undefined, initialText);
        editText.stepperGroup = stepperGroup;
        editText.stepOptions = stepOptions;
        bindSteppedArrowKeys(editText, stepperGroup);
        return editText;
    }

    /**
     * 数値欄と∧∨の有効／無効をまとめて切り替える
     * @param {EditText} editText - addStepperEditText() で作った入力欄
     * @param {boolean} enabled - 有効にするか
     * @returns {void}
     */
    function setNumberFieldEnabled(editText, enabled) {
        editText.enabled = enabled;
        editText.stepperGroup.enabled = enabled;
        redrawSteppersIn(editText.stepperGroup);
    }

    /**
     * 数値欄に入力の制限（範囲・整数化・負号の禁止）と反映処理を付ける
     * @param {EditText} editText - 入力欄
     * @param {Object} fieldOptions - { min, max, integer, onTyping, onCommit, mirror }
     * @returns {void}
     */
    function bindNumericField(editText, fieldOptions) {
        var min = (typeof fieldOptions.min === 'number') ? fieldOptions.min : -Infinity;
        var max = (typeof fieldOptions.max === 'number') ? fieldOptions.max : Infinity;
        var roundToInteger = !!fieldOptions.integer;
        var onTyping = fieldOptions.onTyping || function () {};
        var onCommit = fieldOptions.onCommit || function () {};
        var mirror = fieldOptions.mirror || null;

        /**
         * 範囲内に収める
         * @param {number} value - 値
         * @param {boolean} roundInt - 整数に丸めるか
         * @returns {number} 収めた値
         */
        function clampFieldValue(value, roundInt) {
            if (isNaN(value)) value = 0;
            value = clampNumber(value, min, max);
            return roundInt ? Math.round(value) : value;
        }

        /**
         * 値を確定形にして書き戻し、連動先にも写す
         * @returns {void}
         */
        function commitFieldValue() {
            editText.text = String(clampFieldValue(parseFloat(editText.text), roundToInteger));
            if (mirror) mirror(editText.text);
        }

        /* 入力中：先頭の「-」を禁止し、範囲内に収める / While typing: no leading minus, keep in range */
        editText.onChanging = function () {
            editText.text = String(editText.text || '').replace(/^-+/, '');
            var value = parseFloat(editText.text);
            if (!isNaN(value)) {
                editText.text = String(clampFieldValue(value, false));
                if (mirror) mirror(editText.text);
            }
            onTyping();
        };
        editText.onChange = function () {
            commitFieldValue();
            onCommit();
        };
        /* ∧∨・↑↓：範囲は入力と同じにし、増減後に確定形へそろえて連動先にも写す / steppers share the field's range */
        editText.stepOptions.min = (typeof fieldOptions.min === 'number') ? fieldOptions.min : undefined;
        editText.stepOptions.max = (typeof fieldOptions.max === 'number') ? fieldOptions.max : undefined;
        editText.stepOptions.integer = roundToInteger;
        editText.stepOptions.onStep = function () {
            commitFieldValue();
            onTyping();
        };
    }

    /**
     * フォーカス中はダイアログのショートカットを止める
     * @param {EditText} control - 入力欄
     * @param {Object} hotkeyGuard - { active: boolean }
     * @returns {void}
     */
    function attachHotkeyGuard(control, hotkeyGuard) {
        control.addEventListener('focus', function () {
            hotkeyGuard.active = true;
        });
        control.addEventListener('blur', function () {
            hotkeyGuard.active = false;
        });
    }

    // =========================================
    // HEX 欄 / HEX field
    // =========================================

    /**
     * HEX 欄の文字列を検証する
     * @param {string} text - 入力文字列
     * @param {boolean} coerceUpper - 大文字にそろえるか
     * @returns {{valid: boolean, text: string, message: string|null}} 検証結果
     */
    function validateHex(text, coerceUpper) {
        if (!text) return { valid: false, text: "", message: null };
        var s = String(text).replace(/^\s+|\s+$/g, '');
        if (s === "#") return { valid: false, text: s, message: getLabel(LABELS.tooltip.hexEmpty) };
        var match = s.match(/^#?([0-9a-fA-F]{6})$/);
        if (match) {
            return { valid: true, text: "#" + (coerceUpper ? match[1].toUpperCase() : match[1]), message: null };
        }
        return { valid: false, text: s, message: getLabel(LABELS.tooltip.hexInvalid) };
    }

    /**
     * HEX 欄を検証して警告表示とプレビューを更新する
     * @param {Object} ui - ダイアログのコントロール
     * @param {boolean} coerceUpper - 大文字にそろえるか
     * @returns {void}
     */
    function handleHexInput(ui, coerceUpper) {
        var hexInput = ui.hexInput;
        var result = validateHex(hexInput.text, coerceUpper);
        hexInput.text = result.text || hexInput.text;
        setFieldWarning(hexInput, !result.valid,
            result.message || getLabel(LABELS.tooltip.hexInvalid), getLabel(LABELS.tooltip.hexInput));
        refreshPreview(ui);
    }

    // =========================================
    // CMYK 欄 / CMYK fields
    // =========================================

    /**
     * CMYK 欄の警告表示を切り替える
     * @param {EditText} editText - 入力欄
     * @param {boolean} warn - 警告するか
     * @returns {void}
     */
    function setCmykWarning(editText, warn) {
        setFieldWarning(editText, warn, getLabel(LABELS.tooltip.cmykRange), '');
    }

    /**
     * 0〜100 の範囲外なら警告する（空欄は入力途中とみなす）
     * @param {EditText} editText - 入力欄
     * @returns {void}
     */
    function validateCmykField(editText) {
        var text = String(editText.text || '');
        if (text === '') {
            setCmykWarning(editText, false);
            return;
        }
        var value = parseFloat(text);
        setCmykWarning(editText, isNaN(value) || value < 0 || value > 100);
    }

    /**
     * 0〜100 に収める（空欄・不正な値は 0）
     * @param {EditText} editText - 入力欄
     * @returns {void}
     */
    function clampCmykField(editText) {
        var value = parseFloat(String(editText.text || ''));
        if (isNaN(value)) value = 0;
        editText.text = String(clampNumber(value, 0, 100));
        setCmykWarning(editText, false);
    }

    /**
     * CMYK 欄に入力の補助（0 の消去・先頭 0 の置き換え・範囲の制限）とプレビューの反映を付ける
     * @param {EditText} editText - 入力欄
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function bindCmykField(editText, ui) {
        attachHotkeyGuard(editText, ui.hotkeyGuard);

        /* フォーカス時に「0」を空にして打ちやすくする / Clear "0" on focus */
        editText.addEventListener('focus', function () {
            if (String(editText.text) === '0') editText.text = '';
        });

        /* 「0」の状態で数字を打ったら置き換える（「03」にしない）/ Replace a lone "0" with the typed digit */
        editText.addEventListener('keydown', function (event) {
            var keyName = String(event.keyName || '');
            if (!/^[0-9]$/.test(keyName) || String(editText.text || '') !== '0') return;
            editText.text = keyName;
            event.preventDefault();
            validateCmykField(editText);
            refreshPreview(ui);
        });

        editText.onChanging = function () {
            editText.text = String(editText.text || '').replace(/^-+/, '').replace(/^0([0-9])$/, '$1');
            validateCmykField(editText);
            refreshPreview(ui);
        };
        editText.onChange = function () {
            clampCmykField(editText);
            refreshPreview(ui);
        };
        editText.stepOptions.max = 100;
        editText.stepOptions.onStep = function () {
            clampCmykField(editText);
            refreshPreview(ui);
        };
    }

    /**
     * CMYK 欄の値を 0〜100 の数値で読む
     * @param {Object} ui - ダイアログのコントロール
     * @returns {{c: number, m: number, y: number, k: number}} CMYK 値
     */
    function readCmykValues(ui) {
        var keys = ['c', 'm', 'y', 'k'];
        var values = {};
        for (var i = 0; i < keys.length; i++) {
            var value = parseFloat(ui.cmykInputs[i].text);
            values[keys[i]] = clampNumber(isNaN(value) ? 0 : value, 0, 100);
        }
        return values;
    }

    // =========================================
    // ダイアログの状態 / Dialog state
    // =========================================

    /**
     * 選択中のカラーモードを返す
     * @param {Object} ui - ダイアログのコントロール
     * @returns {string} ColorMode の値
     */
    function getSelectedColorMode(ui) {
        if (ui.whiteRadio.value) return ColorMode.WHITE;
        if (ui.hexRadio.value) return ColorMode.HEX;
        if (ui.cmykRadio.value) return ColorMode.CMYK;
        return ColorMode.K100;
    }

    /**
     * 不透明度欄を読む（0〜100、不正なら 100）
     * @param {EditText} editText - 入力欄
     * @returns {number} 不透明度
     */
    function readOpacityValue(editText) {
        var value = parseFloat(editText.text);
        if (isNaN(value)) value = 100;
        return clampNumber(Math.round(value), 0, 100);
    }

    /**
     * ダイアログの内容から設定を組み立てる（長さは pt）
     * @param {Object} ui - ダイアログのコントロール
     * @returns {Object} 設定
     */
    function collectChoice(ui) {
        var pointsPerUnit = getUnitInfo().pointsPerUnit;
        var offsetVPt = fieldTextToPt(ui.offsetVInput.text, pointsPerUnit);
        var roundEnabled = !!ui.cbRoundEnable.value;
        var opacityEnabled = !!ui.cbOpacityApply.value;
        return {
            colorMode: getSelectedColorMode(ui),
            customValue: String(ui.hexInput.text || '').replace(/^\s+|\s+$/g, ''),
            customCMYK: readCmykValues(ui),
            offsetV: offsetVPt,
            offsetH: ui.linkMarginsToggle.value ? offsetVPt : fieldTextToPt(ui.offsetHInput.text, pointsPerUnit),
            roundPt: roundEnabled ? fieldTextToPt(ui.roundInput.text, pointsPerUnit) : 0,
            isPill: roundEnabled ? !!ui.cbPill.value : false,
            type: ui.paintStrokeRadio.value ? 'stroke' : 'fill',
            strokeWidth: fieldTextToPt(ui.strokeWidthInput.text, pointsPerUnit),
            target: (!ui.individualRadio.value && ui.groupRadio.value) ? 'group' : 'individual',
            /* キー名は保存ファイルとの互換のため据え置き / Key kept for saved-state compatibility */
            groupWithText: !!ui.cbGroupWithObjects.value,
            /* OFF のときは 100% として扱う / Treated as 100% when off */
            opacity: opacityEnabled ? readOpacityValue(ui.opacityInput) : 100,
            opacityEnabled: opacityEnabled
        };
    }

    /**
     * 保存した設定をダイアログに反映する
     * @param {Object} ui - ダイアログのコントロール
     * @param {Object} saved - 保存した設定
     * @returns {void}
     */
    function applyChoiceToUI(ui, saved) {
        var pointsPerUnit = getUnitInfo().pointsPerUnit;

        if (saved.colorMode === ColorMode.K100) {
            ui.blackRadio.notify('onClick');
        } else if (saved.colorMode === ColorMode.WHITE) {
            ui.whiteRadio.notify('onClick');
        } else if (saved.colorMode === ColorMode.HEX) {
            ui.hexRadio.notify('onClick');
            ui.hexInput.text = saved.customValue || '#';
        } else if (saved.colorMode === ColorMode.CMYK) {
            ui.cmykRadio.notify('onClick');
            /* 保存ファイルに customCMYK が無いことがある / customCMYK may be missing */
            try {
                ui.cmykInputs[0].text = saved.customCMYK.c;
                ui.cmykInputs[1].text = saved.customCMYK.m;
                ui.cmykInputs[2].text = saved.customCMYK.y;
                ui.cmykInputs[3].text = saved.customCMYK.k;
            } catch (e) {
                logError("applyChoiceToUI.cmyk", e);
            }
        }

        ui.offsetVInput.text = String(Math.round(saved.offsetV / pointsPerUnit));
        ui.offsetHInput.text = String(Math.round(saved.offsetH / pointsPerUnit));

        ui.cbRoundEnable.value = (saved.roundPt > 0 || saved.isPill);
        ui.roundInput.text = String(Math.round(saved.roundPt / pointsPerUnit));
        ui.cbPill.value = !!saved.isPill;
        setCornerUIEnabled(ui, ui.cbRoundEnable.value);

        if (saved.type === 'stroke') ui.paintStrokeRadio.notify('onClick');
        else ui.paintFillRadio.notify('onClick');
        ui.strokeWidthInput.text = String(Math.round((saved.strokeWidth || 1) / pointsPerUnit));

        if (saved.target === 'group') ui.groupRadio.notify('onClick');
        else ui.individualRadio.notify('onClick');

        ui.cbGroupWithObjects.value = !!saved.groupWithText;
        ui.cbOpacityApply.value = (typeof saved.opacityEnabled === 'boolean') ? saved.opacityEnabled : true;
        setOpacityUIEnabled(ui, ui.cbOpacityApply.value);

        refreshPreview(ui);
    }

    /**
     * ピル形状のとき、マージン込みの高さの半分を角丸欄に表示し、左右マージンにも写す
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function updatePillRoundField(ui) {
        if (!ui.cbRoundEnable.value || !ui.cbPill.value) return;
        try {
            var doc = ui.doc;
            var items = doc.selection || [];
            if (!items.length) return;

            var pointsPerUnit = getUnitInfo().pointsPerUnit;
            var offsetVPt = fieldTextToPt(ui.offsetVInput.text, pointsPerUnit);
            syncOutlineCache(items);
            var bounds = ui.groupRadio.value ?
                getCombinedFinalBoundsCached(doc, items) :
                getFinalItemBoundsCached(doc, items[0], 0);
            if (!bounds) return;

            var pillRadiusPt = (bounds[1] - bounds[3] + offsetVPt * 2) / 2;
            ui.roundInput.text = String(Math.round((pillRadiusPt / pointsPerUnit) * 100) / 100);
            setNumberFieldEnabled(ui.roundInput, false);
            ui.offsetHInput.text = String(ui.roundInput.text);
        } catch (e) {
            logError("updatePillRoundField", e);
        }
    }

    /**
     * ピル形状の表示を更新してからプレビューを描き直す
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function refreshPreview(ui) {
        updatePillRoundField(ui);
        try {
            renderPreview(ui.doc, collectChoice(ui));
        } catch (e) {
            logError("refreshPreview", e);
        }
    }

    /**
     * 「連動」がONのとき左右の欄を無効にする
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function updateLinkDim(ui) {
        setNumberFieldEnabled(ui.offsetHInput, !ui.linkMarginsToggle.value);
        ui.offsetHLabel.enabled = !ui.linkMarginsToggle.value;
    }

    /**
     * 角丸の欄とピル形状の有効／無効を切り替える
     * @param {Object} ui - ダイアログのコントロール
     * @param {boolean} enabled - 有効にするか
     * @returns {void}
     */
    function setCornerUIEnabled(ui, enabled) {
        setNumberFieldEnabled(ui.roundInput, enabled && !ui.cbPill.value);
        ui.roundUnitLabel.enabled = enabled;
        ui.cbPill.enabled = enabled;
    }

    /**
     * 不透明度欄の有効／無効を切り替える。無効時は「60」を薄く表示する（計算上は 100%）
     * @param {Object} ui - ダイアログのコントロール
     * @param {boolean} enabled - 有効にするか
     * @returns {void}
     */
    function setOpacityUIEnabled(ui, enabled) {
        setNumberFieldEnabled(ui.opacityInput, enabled);
        ui.opacityPercentLabel.enabled = enabled;
        setEditHighlight(ui.opacityInput, enabled);
        if (!enabled) ui.opacityInput.text = '60';
    }

    /**
     * 不透明度が 100 なら［適用］をOFF、それ以外ならONにそろえる
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function syncOpacityApplyAuto(ui) {
        var shouldEnable = (readOpacityValueRaw(ui.opacityInput) !== 100);
        if (ui.cbOpacityApply.value !== shouldEnable) {
            ui.cbOpacityApply.value = shouldEnable;
            setOpacityUIEnabled(ui, shouldEnable);
        }
    }

    /**
     * 不透明度欄を丸めた数値で読む（範囲は制限しない、不正なら 100）
     * @param {EditText} editText - 入力欄
     * @returns {number} 値
     */
    function readOpacityValueRaw(editText) {
        var value = parseFloat(editText.text);
        return isNaN(value) ? 100 : Math.round(value);
    }

    /**
     * 線幅の欄の有効／無効を切り替える
     * @param {Object} ui - ダイアログのコントロール
     * @param {boolean} enabled - 有効にするか
     * @returns {void}
     */
    function setStrokeWidthUIEnabled(ui, enabled) {
        setNumberFieldEnabled(ui.strokeWidthInput, enabled);
        ui.strokeWidthLabel.enabled = enabled;
        ui.strokeWidthUnitLabel.enabled = enabled;
    }

    /**
     * カラーモードのラジオに合わせて HEX／CMYK 欄の有効／無効を切り替える
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function updateColorFieldsEnabled(ui) {
        var cmykEnabled = !!ui.cmykRadio.value;
        ui.hexInput.enabled = !!ui.hexRadio.value;
        for (var i = 0; i < ui.cmykInputs.length; i++) {
            setNumberFieldEnabled(ui.cmykInputs[i], cmykEnabled);
            ui.cmykLabels[i].enabled = cmykEnabled;
            if (!cmykEnabled) setCmykWarning(ui.cmykInputs[i], false);
        }
    }

    /**
     * カラーモードを切り替える（ラジオ・欄の有効状態・強調・フォーカス・プレビュー）
     * ラジオが別々のグループにあるため、排他は手動でそろえる
     * @param {Object} ui - ダイアログのコントロール
     * @param {string} mode - ColorMode の値
     * @returns {void}
     */
    function applyColorMode(ui, mode) {
        ui.blackRadio.value = (mode === ColorMode.K100);
        ui.whiteRadio.value = (mode === ColorMode.WHITE);
        ui.hexRadio.value = (mode === ColorMode.HEX);
        ui.cmykRadio.value = (mode === ColorMode.CMYK);
        updateColorFieldsEnabled(ui);

        setEditHighlight(ui.hexInput, mode === ColorMode.HEX);
        setEditHighlight(ui.cmykInputs[0], mode === ColorMode.CMYK);
        if (mode === ColorMode.HEX) ui.hexInput.active = true;
        else if (mode === ColorMode.CMYK) ui.cmykInputs[0].active = true;

        /* ホワイトを選んだら種別を一度「塗り」にする（あとから変更可）/ White nudges the type to Fill once */
        if (mode === ColorMode.WHITE) ui.paintFillRadio.notify('onClick');

        refreshPreview(ui);
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * 行グループを作る
     * @param {Group|Panel} parent - 親
     * @param {Array} alignChildren - 子の揃え
     * @param {number} spacing - 間隔
     * @returns {Group} 行グループ
     */
    function addRowGroup(parent, alignChildren, spacing) {
        var rowGroup = parent.add('group');
        rowGroup.orientation = 'row';
        rowGroup.alignChildren = alignChildren;
        if (typeof spacing === 'number') rowGroup.spacing = spacing;
        return rowGroup;
    }

    /**
     * パネルを作る
     * @param {Group} parent - 親
     * @param {Object} titleLabel - パネル名のラベル
     * @param {string} orientation - 'row' / 'column'
     * @param {Array|string} alignChildren - 子の揃え
     * @param {number[]} margins - 余白
     * @param {number} [spacing] - 間隔
     * @returns {Panel} パネル
     */
    function addPanel(parent, titleLabel, orientation, alignChildren, margins, spacing) {
        var panel = parent.add('panel', undefined, getLabel(titleLabel));
        panel.orientation = orientation;
        panel.alignChildren = alignChildren;
        if (typeof spacing === 'number') panel.spacing = spacing;
        panel.margins = margins;
        return panel;
    }

    /**
     * 縦に並べるカラムを作る
     * @param {Group} parent - 親
     * @returns {Group} カラム
     */
    function addColumn(parent) {
        var column = parent.add('group');
        column.orientation = 'column';
        column.alignChildren = 'fill';
        column.spacing = COLUMN_INNER_SPACING;
        return column;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
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
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkLinkToggleUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var LINK_UI_DARK = isDarkLinkToggleUI();
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

    /**
     * マージンパネル（上下・左右・連動）を作る
     * @param {Group} parent - 親カラム
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildMarginPanel(parent, ui) {
        var marginPanel = addPanel(parent, LABELS.panel.margin, 'row', ['left', 'top'], MARGIN_PANEL_MARGINS, 10);
        var marginBody = addRowGroup(marginPanel, ['left', 'center'], 10);
        marginBody.margins = 0;

        var fieldColumn = marginBody.add('group');
        fieldColumn.orientation = 'column';
        fieldColumn.alignChildren = ['left', 'center'];
        fieldColumn.spacing = 10;
        fieldColumn.margins = 0;
        var unitLabel = getUnitInfo().label;

        var verticalRow = addRowGroup(fieldColumn, ['left', 'center'], 10);
        verticalRow.margins = 0;
        ui.offsetVLabel = verticalRow.add('statictext', undefined, labelText(LABELS.fieldLabel.offsetV));
        ui.offsetVInput = addNumberField(verticalRow, '2', true);
        verticalRow.add('statictext', undefined, unitLabel);

        var horizontalRow = addRowGroup(fieldColumn, ['left', 'center'], 10);
        horizontalRow.margins = 0;
        ui.offsetHLabel = horizontalRow.add('statictext', undefined, labelText(LABELS.fieldLabel.offsetH));
        ui.offsetHInput = addNumberField(horizontalRow, '2', true);
        horizontalRow.add('statictext', undefined, unitLabel);

        /* 連動のリンクアイコンは2つの欄の右、上下の中央に置く / The link icon sits right of the two fields, vertically centred */
        ui.linkMarginsToggle = addLinkToggle(marginBody, true, function () {
            updateLinkDim(ui);
            if (ui.linkMarginsToggle.value) ui.offsetHInput.text = String(ui.offsetVInput.text);
            refreshPreview(ui);
        });
        ui.linkMarginsToggle.helpTip = getLabel(LABELS.tooltip.linkMargins);

        bindMarginField(ui.offsetVInput, function () { return ui.offsetHInput; }, ui);
        bindMarginField(ui.offsetHInput, function () { return ui.offsetVInput; }, ui);
        attachHotkeyGuard(ui.offsetVInput, ui.hotkeyGuard);
        attachHotkeyGuard(ui.offsetHInput, ui.hotkeyGuard);
    }

    /**
     * マージン欄を数値欄にし、「連動」がONなら相手の欄へ値を写す
     * @param {EditText} editText - マージン欄
     * @param {Function} getPartnerField - 相手の欄を返す関数
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function bindMarginField(editText, getPartnerField, ui) {
        bindNumericField(editText, {
            min: 0,
            integer: true,
            mirror: function (value) {
                if (ui.linkMarginsToggle.value) getPartnerField().text = String(value);
            },
            onTyping: function () { refreshPreview(ui); },
            onCommit: function () { refreshPreview(ui); }
        });
    }

    /**
     * 角丸パネル（有効化・半径・ピル形状）を作る
     * @param {Group} parent - 親カラム
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildCornerPanel(parent, ui) {
        var cornerPanel = addPanel(parent, LABELS.panel.corner, 'column', ['left', 'top'], PANEL_MARGINS);

        var radiusRow = addRowGroup(cornerPanel, ['left', 'center']);
        ui.cbRoundEnable = radiusRow.add('checkbox', undefined, '');
        ui.cbRoundEnable.value = true;
        ui.cbRoundEnable.alignment = ['left', 'center'];
        ui.cbRoundEnable.helpTip = getLabel(LABELS.tooltip.roundEnable);

        ui.roundInput = addNumberField(radiusRow, '2', true);
        ui.lastRoundValue = String(ui.roundInput.text);
        bindNumericField(ui.roundInput, {
            min: 0,
            integer: true,
            onTyping: function () { refreshPreview(ui); },
            onCommit: function () { refreshPreview(ui); }
        });
        ui.roundUnitLabel = radiusRow.add('statictext', undefined, getUnitInfo().label);

        var pillRow = addRowGroup(cornerPanel, ['left', 'center']);
        ui.cbPill = pillRow.add('checkbox', undefined, getLabel(LABELS.checkbox.pillShape));
        ui.cbPill.value = false;
        ui.cbPill.helpTip = getLabel(LABELS.tooltip.pillShape);

        ui.cbRoundEnable.onClick = function () {
            /* OFF にするときは値を控えて 0 に、ON で戻す / Remember the value when turning off */
            if (ui.cbRoundEnable.value) {
                if (ui.lastRoundValue) ui.roundInput.text = ui.lastRoundValue;
            } else {
                ui.lastRoundValue = String(ui.roundInput.text);
                ui.roundInput.text = '0';
            }
            setCornerUIEnabled(ui, ui.cbRoundEnable.value);
            if (ui.cbRoundEnable.value) ui.roundInput.active = true;
            refreshPreview(ui);
        };

        ui.cbPill.onClick = function () {
            setNumberFieldEnabled(ui.roundInput, !ui.cbPill.value);
            if (ui.cbPill.value) {
                /* 半径を計算し、連動をOFFにして左右マージンへ写す / Compute radius, unlink, copy to horizontal */
                updatePillRoundField(ui);
                setLinkToggleValue(ui.linkMarginsToggle, false);
                updateLinkDim(ui);
                ui.offsetHInput.text = String(ui.roundInput.text);
            }
            refreshPreview(ui);
        };
    }

    /**
     * カラーパネル（ブラック・ホワイト・HEX・CMYK）を作る
     * @param {Group} parent - 親カラム
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildColorPanel(parent, ui) {
        var colorPanel = addPanel(parent, LABELS.panel.color, 'column', 'left', PANEL_MARGINS, 10);

        ui.blackRadio = colorPanel.add('radiobutton', undefined, getLabel(LABELS.radio.colorBlack));
        ui.blackRadio.helpTip = getLabel(LABELS.tooltip.colorBlack);
        ui.whiteRadio = colorPanel.add('radiobutton', undefined, getLabel(LABELS.radio.colorWhite));
        ui.whiteRadio.helpTip = getLabel(LABELS.tooltip.colorWhite);

        var hexRow = addRowGroup(colorPanel, ['left', 'center'], 6);
        hexRow.alignment = 'left';
        ui.hexRadio = hexRow.add('radiobutton', undefined, getLabel(LABELS.radio.colorHex));
        ui.hexRadio.helpTip = getLabel(LABELS.tooltip.colorHex);
        ui.hexInput = hexRow.add('edittext', undefined, '#');
        ui.hexInput.characters = 14;
        ui.hexInput.helpTip = getLabel(LABELS.tooltip.hexInput);
        ui.hexInput.onChanging = function () { handleHexInput(ui, false); };
        ui.hexInput.onChange = function () { handleHexInput(ui, true); };
        attachHotkeyGuard(ui.hexInput, ui.hotkeyGuard);

        ui.cmykRadio = colorPanel.add('radiobutton', undefined, getLabel(LABELS.radio.colorCmyk));
        ui.cmykRadio.helpTip = getLabel(LABELS.tooltip.colorCmyk);
        buildCmykFields(colorPanel, ui);

        ui.blackRadio.value = true;
        ui.whiteRadio.value = false;
        updateColorFieldsEnabled(ui);

        ui.blackRadio.onClick = function () { applyColorMode(ui, ColorMode.K100); };
        ui.whiteRadio.onClick = function () { applyColorMode(ui, ColorMode.WHITE); };
        ui.hexRadio.onClick = function () { applyColorMode(ui, ColorMode.HEX); };
        ui.cmykRadio.onClick = function () { applyColorMode(ui, ColorMode.CMYK); };
    }

    /**
     * CMYK の見出し行と入力行を作る
     * @param {Panel} parent - カラーパネル
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildCmykFields(parent, ui) {
        var cmykGroup = parent.add('group');
        cmykGroup.orientation = 'column';
        cmykGroup.alignment = 'left';
        cmykGroup.spacing = 4;
        var headingRow = addRowGroup(cmykGroup, ['left', 'center'], 10);
        var inputRow = addRowGroup(cmykGroup, ['left', 'center'], 10);

        var channelNames = ['C', 'M', 'Y', 'K'];
        ui.cmykLabels = [];
        ui.cmykInputs = [];
        for (var i = 0; i < channelNames.length; i++) {
            var channelLabel = headingRow.add('statictext', undefined, '  ' + channelNames[i]);
            /* 入力欄の左に∧∨が付くので、その幅を足して列をそろえる / include the stepper so the headings line up */
            channelLabel.preferredSize.width = CMYK_COLUMN_WIDTH + STEPPER_BUTTON_WIDTH + STEPPER_SIDE_MARGIN;
            ui.cmykLabels.push(channelLabel);
        }
        for (var j = 0; j < channelNames.length; j++) {
            var channelInput = addStepperEditText(inputRow, '', false);
            channelInput.characters = 3;
            channelInput.preferredSize.width = CMYK_COLUMN_WIDTH;
            bindCmykField(channelInput, ui);
            ui.cmykInputs.push(channelInput);
        }
    }

    /**
     * 対象パネル（個別・グループとして）を作る。初期値はアートボードが複数なら「グループとして」
     * @param {Group} parent - 親カラム
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildTargetPanel(parent, ui) {
        var targetPanel = addPanel(parent, LABELS.panel.target, 'row', ['left', 'center'], PANEL_MARGINS, 20);
        ui.individualRadio = targetPanel.add('radiobutton', undefined, getLabel(LABELS.radio.targetIndividual));
        ui.individualRadio.helpTip = getLabel(LABELS.tooltip.targetIndividual);
        ui.groupRadio = targetPanel.add('radiobutton', undefined, getLabel(LABELS.radio.targetGroup));
        ui.groupRadio.helpTip = getLabel(LABELS.tooltip.targetGroup);

        var hasMultipleArtboards = ui.doc.artboards.length > 1;
        ui.individualRadio.value = !hasMultipleArtboards;
        ui.groupRadio.value = hasMultipleArtboards;

        ui.individualRadio.onClick = function () { refreshPreview(ui); };
        ui.groupRadio.onClick = function () { refreshPreview(ui); };
    }

    /**
     * 不透明度パネル（適用・値）を作る
     * @param {Group} parent - 親カラム
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildOpacityPanel(parent, ui) {
        var opacityPanel = addPanel(parent, LABELS.panel.opacity, 'row', ['left', 'center'], PANEL_MARGINS, 10);
        ui.cbOpacityApply = opacityPanel.add('checkbox', undefined, getLabel(LABELS.checkbox.opacityApply));
        ui.cbOpacityApply.value = true;
        ui.cbOpacityApply.helpTip = getLabel(LABELS.tooltip.opacityApply);

        ui.opacityInput = addStepperEditText(opacityPanel, '100', true);
        ui.opacityInput.characters = 3;
        ui.opacityInput.preferredSize = { width: OPACITY_FIELD_WIDTH, height: -1 };
        ui.opacityPercentLabel = opacityPanel.add('statictext', undefined, '%');

        ui.cbOpacityApply.onClick = function () {
            setOpacityUIEnabled(ui, ui.cbOpacityApply.value);
            refreshPreview(ui);
        };

        bindNumericField(ui.opacityInput, {
            min: 0,
            max: 100,
            integer: true,
            onTyping: function () { refreshPreview(ui); },
            onCommit: function () { refreshPreview(ui); }
        });

        /* 入力・確定の後で［適用］を値に合わせる（∧∨・↑↓キーでは合わせない）
           Sync the Apply checkbox after typing and commit (not after arrow keys) */
        var numericOnChanging = ui.opacityInput.onChanging;
        var numericOnChange = ui.opacityInput.onChange;
        ui.opacityInput.onChanging = function () {
            numericOnChanging();
            syncOpacityApplyAuto(ui);
        };
        ui.opacityInput.onChange = function () {
            numericOnChange();
            syncOpacityApplyAuto(ui);
        };
    }

    /**
     * 種別パネル（塗り・線・線幅）を作る
     * @param {Group} parent - 親カラム
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildPaintTypePanel(parent, ui) {
        var paintTypePanel = addPanel(parent, LABELS.panel.paintType, 'row', ['left', 'center'], PANEL_MARGINS, 20);
        ui.paintFillRadio = paintTypePanel.add('radiobutton', undefined, getLabel(LABELS.radio.paintFill));
        ui.paintStrokeRadio = paintTypePanel.add('radiobutton', undefined, getLabel(LABELS.radio.paintStroke));

        var strokeWidthRow = addRowGroup(paintTypePanel, ['left', 'center'], 6);
        ui.strokeWidthLabel = strokeWidthRow.add('statictext', undefined, labelText(LABELS.fieldLabel.strokeWidth));
        ui.strokeWidthInput = addNumberField(strokeWidthRow, '1');
        bindNumericField(ui.strokeWidthInput, {
            min: 0,
            integer: false,
            onTyping: function () { refreshPreview(ui); },
            onCommit: function () { refreshPreview(ui); }
        });
        ui.strokeWidthUnitLabel = strokeWidthRow.add('statictext', undefined, getUnitInfo().label);

        var onPaintTypeClick = function () {
            setStrokeWidthUIEnabled(ui, !!ui.paintStrokeRadio.value);
            refreshPreview(ui);
        };
        ui.paintFillRadio.onClick = onPaintTypeClick;
        ui.paintStrokeRadio.onClick = onPaintTypeClick;

        ui.paintFillRadio.value = true;
        ui.paintStrokeRadio.value = false;
    }

    /**
     * 「オブジェクトとグループ化」のチェックボックスを作る
     * @param {Group} parent - 親カラム
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildGroupOption(parent, ui) {
        var groupOptionRow = parent.add('group');
        groupOptionRow.orientation = 'column';
        groupOptionRow.alignment = 'center';
        groupOptionRow.alignChildren = ['left', 'center'];
        ui.cbGroupWithObjects = groupOptionRow.add('checkbox', undefined, getLabel(LABELS.checkbox.groupWithObjects));
        ui.cbGroupWithObjects.value = true;
        ui.cbGroupWithObjects.helpTip = getLabel(LABELS.tooltip.groupWithObjects);
        ui.cbGroupWithObjects.onClick = function () { refreshPreview(ui); };
    }

    /**
     * ボタン行（キャンセル・OK）を作る
     * @param {Window} dlg - ダイアログ
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function buildButtonRow(dlg, ui) {
        var btnRowGroup = dlg.add('group');
        btnRowGroup.alignment = 'center';
        var btnCancel = btnRowGroup.add('button', undefined, getLabel(LABELS.button.cancel), { name: 'cancel' });
        var btnOK = btnRowGroup.add('button', undefined, getLabel(LABELS.button.ok), { name: 'ok' });

        btnOK.onClick = function () {
            saveLastChoice(serializeChoice(collectChoice(ui)));
            dlg.close(1);
        };
        btnCancel.onClick = function () {
            clearPreview(ui.doc, true);
            dlg.close(0);
        };
    }

    /**
     * ショートカット（K/W/H/C：カラー、I/G：対象。旧 S/A も有効）を付ける。入力欄にフォーカス中は無効
     * @param {Window} dlg - ダイアログ
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function bindDialogHotkeys(dlg, ui) {
        var colorModeByKey = { K: ColorMode.K100, W: ColorMode.WHITE, H: ColorMode.HEX, C: ColorMode.CMYK };
        dlg.addEventListener('keydown', function (event) {
            if (ui.hotkeyGuard.active) return;
            var key = event.keyName ? String(event.keyName).toUpperCase() : '';
            if (colorModeByKey.hasOwnProperty(key)) {
                applyColorMode(ui, colorModeByKey[key]);
                event.preventDefault();
            } else if (key === 'I' || key === 'S') {
                ui.individualRadio.notify('onClick');
                event.preventDefault();
            } else if (key === 'G' || key === 'A') {
                ui.groupRadio.notify('onClick');
                event.preventDefault();
            }
        });
    }

    /**
     * 表示時の初期化（前回の設定の復元・位置・有効状態・最初のプレビュー）
     * @param {Object} ui - ダイアログのコントロール
     * @returns {void}
     */
    function initDialogOnShow(ui) {
        /* 保存ファイルが壊れていても開けるようにする / Keep opening even if the saved file is broken */
        try {
            var lastChoice = loadLastChoice();
            if (lastChoice) applyChoiceToUI(ui, lastChoice);
        } catch (e) {
            logError("initDialogOnShow.restore", e);
        }
        setStrokeWidthUIEnabled(ui, !!ui.paintStrokeRadio.value);
        /* 初回の表示位置。前回の位置があれば prepareDialogWindow() が上書きする
           First-run placement; prepareDialogWindow() overrides it with the last position */
        var defaultLocation = ui.dialog.location;
        ui.dialog.location = [defaultLocation[0] + DIALOG_OFFSET_X, defaultLocation[1] + DIALOG_OFFSET_Y];
        ui.offsetVInput.active = true;
        ui.cbGroupWithObjects.value = true;
        updateLinkDim(ui);
        if (ui.cbPill.value) {
            setNumberFieldEnabled(ui.roundInput, false);
            updatePillRoundField(ui);
        }
        setCornerUIEnabled(ui, !!ui.cbRoundEnable.value);
        syncOpacityApplyAuto(ui);
        refreshPreview(ui);
    }

    /**
     * ダイアログを表示し、OK なら設定を返す
     * @param {Document} doc - ドキュメント
     * @returns {Object|null} 設定。キャンセルなら null
     */
    function showDialog(doc) {
        /* 今回の選択を固定する / Pin the selection for this session */
        sessionSelection = snapshotSelection(doc);

        var dlg = new Window('dialog', getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dlg.alignChildren = 'left';

        var ui = { dialog: dlg, doc: doc, hotkeyGuard: { active: false }, lastRoundValue: '2' };

        var columnsRow = addRowGroup(dlg, 'top', COLUMN_SPACING);
        var leftColumn = addColumn(columnsRow);
        var rightColumn = addColumn(columnsRow);

        buildMarginPanel(leftColumn, ui);
        buildCornerPanel(leftColumn, ui);
        buildColorPanel(rightColumn, ui);
        buildTargetPanel(leftColumn, ui);
        buildOpacityPanel(rightColumn, ui);
        buildPaintTypePanel(rightColumn, ui);
        buildGroupOption(leftColumn, ui);
        buildButtonRow(dlg, ui);
        bindDialogHotkeys(dlg, ui);

        dlg.onShow = function () { initDialogOnShow(ui); };

        prepareDialogWindow(dlg, SCRIPT_NAME);
        if (dlg.show() != 1) return null;
        return collectChoice(ui);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで設定し、プレビューの長方形を確定する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var choice = showDialog(doc);
        if (choice === null) {
            /* Esc やクローズボックスで閉じたときもプレビューを片付ける / Also clean up on Esc or close box */
            clearPreview(doc, true);
            return;
        }

        /* 確定前の選択を控えてから解除する / Keep the selection, then deselect */
        var targetItems = snapshotSelection(doc);
        app.executeMenuCommand('deselectall');

        convertPreviewToFinal(doc, targetItems, choice);
        clearPreview(doc, true);
    }

    try {
        main();
    } catch (e) {
        logError("main", e);
    }

})();
