#target illustrator
#targetengine "SortByNumbersEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

グループ内のテキストから数値を抽出し、その数値でグループを並び替えて縦方向に整列します。
フォント情報によるグループ分けと、昇順・降順・ランダム順に対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortByNumbers.md

### Overview

Extracts numbers from the text inside groups, sorts the groups by those numbers, and aligns them vertically.
Groups can be split by font, and the order can be ascending, descending or random.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortByNumbers.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SortByNumbers";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortByNumbers.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortByNumbers.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「指定」の値が読めないときに使う間隔（pt） / Gap used when the custom value cannot be read, in points */
    var FALLBACK_GAP = 20;

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
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "グループの数値で整列", en: "Align Groups by Number" }
        },
        panel: {
            sortGroup: { ja: "数値グループ", en: "Number Group" },
            spacing:   { ja: "間隔", en: "Spacing" }
        },
        radio: {
            asc:    { ja: "昇順", en: "Ascending" },
            desc:   { ja: "降順", en: "Descending" },
            random: { ja: "ランダム", en: "Random" },
            fit:    { ja: "ぴったり", en: "Fit" },
            custom: { ja: "指定", en: "Custom" }
        },
        tooltip: {
            sortGroup: { ja: "並べ替えの基準に使う数値のまとまりを選びます。", en: "Which set of numbers to sort by." },
            asc:       { ja: "数値の小さい順に並べます。", en: "Sorts from the smallest number up." },
            desc:      { ja: "数値の大きい順に並べます。", en: "Sorts from the largest number down." },
            random:    { ja: "数値と関係なく、順序をシャッフルします。", en: "Shuffles the order regardless of the numbers." },
            fit:       { ja: "現在の並びの間隔を保ったまま詰め直します。", en: "Keeps the current spacing and repacks the objects." },
            custom:    { ja: "間隔を数値で指定します。", en: "Sets the spacing to a value you type." },
            spacingInput: {
                ja: "オブジェクト間にあける間隔です。「指定」を選んだときだけ使われます。",
                en: "Gap left between objects. Used only when Custom is selected."
            },
            /* ステップボタン用 / for the stepper */
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
            ok:     { ja: "ソート", en: "Sort" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー。ステップボタンからは { ja, en } のオブジェクトでも呼ばれる）
     * @param {string|Object} labelPath - "panel.spacing" のようなドット区切りキー、または { ja, en }
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        if (labelPath && typeof labelPath === "object") return labelPath[uiLang] || labelPath.ja;
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * ドキュメントの定規の単位の表示名と、1単位あたりのポイント数を返す
     * @param {Document} targetDocument - 対象ドキュメント
     * @returns {{label: string, pointsPerUnit: number}} 単位の情報（未対応の単位は pt）
     */
    function getRulerUnit(targetDocument) {
        switch (targetDocument.rulerUnits) {
            case RulerUnits.Millimeters:
                return { label: "mm", pointsPerUnit: 2.83464567 };
            case RulerUnits.Centimeters:
                return { label: "cm", pointsPerUnit: 28.3464567 };
            case RulerUnits.Inches:
                return { label: "inch", pointsPerUnit: 72 };
            case RulerUnits.Pixels:
                return { label: "px", pointsPerUnit: 1 };
            case RulerUnits.Picas:
                return { label: "pica", pointsPerUnit: 12 };
            default:
                return { label: "pt", pointsPerUnit: 1 };
        }
    }

    // =========================================
    // 数値の収集 / Number collection
    // =========================================

    /**
     * テキストフレームの数値をフォント（ファミリー＋スタイル）ごとに集める（グループは再帰的にたどる）
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {Object} fontMap - フォント名 → { value, group } の配列（ここに追加する）
     * @param {GroupItem} ownerGroup - 数値を持たせるグループ（入れ子のグループではそのグループ）
     * @param {Object} [firstNumberRef] - 最初に見つかった数値を value に入れる入れ物
     * @returns {void}
     */
    function collectNumbersByFont(pageItem, fontMap, ownerGroup, firstNumberRef) {
        if (pageItem.typename === "TextFrame") {
            var numberValue = parseFloat(pageItem.contents.replace(/,/g, ""));
            var fontName = "不明";
            try {
                /* 空のテキストでは textRanges[0] が取れないことがある / textRanges[0] may be missing on empty text */
                var textFont = pageItem.textRanges[0].characterAttributes.textFont;
                if (textFont && textFont.family && textFont.style) {
                    fontName = textFont.family + " " + textFont.style;
                }
            } catch (e) {}
            if (!isNaN(numberValue)) {
                if (!fontMap[fontName]) fontMap[fontName] = [];
                fontMap[fontName].push({ value: numberValue, group: ownerGroup });
                if (firstNumberRef && typeof firstNumberRef.value === "undefined") {
                    firstNumberRef.value = numberValue;
                }
            }
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                var childItem = pageItem.pageItems[i];
                collectNumbersByFont(childItem, fontMap, (childItem.typename === "GroupItem" ? childItem : ownerGroup), firstNumberRef);
            }
        }
    }

    /**
     * 選択の中から、数値のテキストを含むグループだけを取り出す
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {GroupItem[]} 数値を含むグループ
     */
    function findNumberedGroups(selectedItems) {
        var numberedGroups = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem") {
                var firstNumberRef = { value: undefined };
                collectNumbersByFont(selectedItems[i], {}, selectedItems[i], firstNumberRef);
                if (!isNaN(firstNumberRef.value)) {
                    numberedGroups.push(selectedItems[i]);
                }
            }
        }
        return numberedGroups;
    }

    /**
     * 同じグループの重複を除き、グループごとに最初の数値だけを残す
     * @param {Object[]} fontEntries - collectNumbersByFont() が集めた1フォント分の配列
     * @returns {Object[]} { group, value } の配列
     */
    function getUniqueGroupEntries(fontEntries) {
        var groupEntries = [];
        for (var i = 0; i < fontEntries.length; i++) {
            var alreadyExists = false;
            for (var j = 0; j < groupEntries.length; j++) {
                if (groupEntries[j].group === fontEntries[i].group) {
                    alreadyExists = true;
                    break;
                }
            }
            if (!alreadyExists) {
                groupEntries.push({ group: fontEntries[i].group, value: fontEntries[i].value });
            }
        }
        return groupEntries;
    }

    // =========================================
    // 並べ替えと配置 / Sorting and placement
    // =========================================

    /**
     * 順序をシャッフルする（Fisher–Yates）。先頭が元の先頭のままなら2番目と入れ替える
     * @param {Object[]} groupEntries - 並べ替える配列（直接書き換える）
     * @returns {void}
     */
    function shuffleGroupEntries(groupEntries) {
        var originalFirst = groupEntries[0];
        for (var i = groupEntries.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapEntry = groupEntries[i];
            groupEntries[i] = groupEntries[j];
            groupEntries[j] = swapEntry;
        }
        if (groupEntries.length > 1 && groupEntries[0].group === originalFirst.group) {
            groupEntries[0] = groupEntries[1];
            groupEntries[1] = originalFirst;
        }
    }

    /**
     * 並べ始める位置（上端がいちばん低いグループの左上）を求める
     * グループ名をキーにして控えるため、同じ名前のグループは後のものだけが候補になる
     * @param {GroupItem[]} numberedGroups - 数値を含むグループ
     * @returns {{left: number, top: number}} 並べ始める左上の座標
     */
    function getStackOrigin(numberedGroups) {
        var boundsByName = {};
        for (var i = 0; i < numberedGroups.length; i++) {
            boundsByName[numberedGroups[i].name] = numberedGroups[i].visibleBounds.concat();
        }
        var startTop = null;
        for (var groupName in boundsByName) {
            if (startTop === null || boundsByName[groupName][1] < startTop) {
                startTop = boundsByName[groupName][1];
            }
        }
        var startLeft = null;
        for (var candidateName in boundsByName) {
            if (boundsByName[candidateName][1] === startTop) {
                startLeft = boundsByName[candidateName][0];
                break;
            }
        }
        return { left: startLeft, top: startTop };
    }

    /**
     * グループを並べた順に、起点から下へ積み上げる
     * @param {Object[]} groupEntries - { group, value } の配列（並べる順）
     * @param {{left: number, top: number}} stackOrigin - 起点の左上
     * @param {boolean} fitsTightly - true なら間隔 0（ぴったり）
     * @param {number} gap - グループ間の間隔（pt）
     * @returns {void}
     */
    function stackGroups(groupEntries, stackOrigin, fitsTightly, gap) {
        var currentTop = stackOrigin.top;
        for (var i = 0; i < groupEntries.length; i++) {
            var groupItem = groupEntries[i].group;
            groupItem.locked = false;
            groupItem.hidden = false;

            var boundsBefore = groupItem.visibleBounds;
            groupItem.translate(stackOrigin.left - boundsBefore[0], currentTop - boundsBefore[1]);

            /* 移動後の高さで次の上端を決める / next top from the height after moving */
            var boundsAfter = groupItem.visibleBounds;
            var groupHeight = boundsAfter[1] - boundsAfter[3];
            if (fitsTightly) {
                currentTop -= groupHeight;
            } else {
                currentTop -= (groupHeight + gap);
            }
        }

        /* 先頭のグループが起点に来るよう全体をずらす / shift everything so the first group sits on the origin */
        var firstBounds = groupEntries[0].group.visibleBounds;
        var shiftX = stackOrigin.left - firstBounds[0];
        var shiftY = stackOrigin.top - firstBounds[1];
        for (var j = 0; j < groupEntries.length; j++) {
            groupEntries[j].group.translate(shiftX, shiftY);
        }
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値のまとまり（フォント）ごとのラジオボタンに出す、先頭3つの数値を返す
     * @param {Object[]} fontEntries - 1フォント分の { value, group } の配列
     * @returns {string} "1, 2, 3…" のような表示
     */
    function getNumberPreviewLabel(fontEntries) {
        var numberValues = [];
        for (var i = 0; i < fontEntries.length; i++) {
            numberValues.push(fontEntries[i].value);
        }
        numberValues.sort(function (a, b) {
            return a - b;
        });
        return (numberValues.length > 3) ? numberValues.slice(0, 3).join(", ") + "…" : numberValues.join(", ");
    }

    /**
     * 数値のまとまり・並び順・間隔を選ぶダイアログを出す
     * @param {Object} fontMap - フォント名 → 数値の配列
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {Object|null} { font, descending, random, spacingMode, spacingValue }。キャンセル時は null
     */
    function showFontChoiceDialog(fontMap, rulerUnit) {
        var sortDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        sortDialog.orientation = "column";
        sortDialog.alignChildren = "fill";

        var numberGroupPanel = sortDialog.add("panel", undefined, getLabel("panel.sortGroup"));
        numberGroupPanel.orientation = "column";
        numberGroupPanel.alignChildren = "left";
        numberGroupPanel.margins = [10, 20, 10, 10];

        var fontNames = [];
        for (var fontName in fontMap) {
            fontNames.push(fontName);
        }
        fontNames.sort();

        var fontRadios = [];
        for (var i = 0; i < fontNames.length; i++) {
            var fontRadio = numberGroupPanel.add("radiobutton", undefined, getNumberPreviewLabel(fontMap[fontNames[i]]));
            fontRadio.helpTip = getLabel("tooltip.sortGroup");
            fontRadios.push({ button: fontRadio, key: fontNames[i] });
        }
        if (fontRadios.length > 0) {
            fontRadios[0].button.value = true;
        }

        var sortOrderPanel = sortDialog.add("panel", undefined);
        sortOrderPanel.orientation = "row";
        sortOrderPanel.alignChildren = "left";
        sortOrderPanel.margins = [10, 20, 10, 10];

        var ascRadio = sortOrderPanel.add("radiobutton", undefined, getLabel("radio.asc"));
        ascRadio.helpTip = getLabel("tooltip.asc");
        var descRadio = sortOrderPanel.add("radiobutton", undefined, getLabel("radio.desc"));
        descRadio.helpTip = getLabel("tooltip.desc");
        var randomRadio = sortOrderPanel.add("radiobutton", undefined, getLabel("radio.random"));
        randomRadio.helpTip = getLabel("tooltip.random");
        ascRadio.value = true;

        var spacingPanel = sortDialog.add("panel", undefined, getLabel("panel.spacing"));
        spacingPanel.orientation = "row";
        spacingPanel.alignChildren = "left";
        spacingPanel.margins = [10, 20, 10, 10];

        var fitRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.fit"));
        fitRadio.helpTip = getLabel("tooltip.fit");
        var customRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.custom"));
        customRadio.helpTip = getLabel("tooltip.custom");
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var spacingStepperInputGroup = spacingPanel.add("group");
        spacingStepperInputGroup.orientation = "row";
        spacingStepperInputGroup.alignChildren = ["left", "center"];
        spacingStepperInputGroup.spacing = 0;
        spacingStepperInputGroup.margins = 0;
        var spacingInput;
        var spacingStepperGroup = addStepper(spacingStepperInputGroup, function () { return spacingInput; }, {});
        spacingInput = spacingStepperInputGroup.add("edittext", undefined, (rulerUnit.label === "mm") ? "1" : "20");
        spacingInput.helpTip = getLabel("tooltip.spacingInput");
        spacingInput.characters = 5;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(spacingInput, spacingStepperGroup);
        spacingPanel.add("statictext", undefined, rulerUnit.label);
        fitRadio.value = true;

        /**
         * 間隔の入力欄と∧∨の有効／無効を切り替え、∧∨を描き直す
         * @param {boolean} isEnabled - 有効にするなら true
         * @returns {void}
         */
        function setSpacingInputEnabled(isEnabled) {
            spacingStepperInputGroup.enabled = isEnabled;
            redrawSteppersIn(spacingStepperInputGroup);
        }
        setSpacingInputEnabled(false);

        customRadio.onClick = function () {
            setSpacingInputEnabled(true);
        };
        fitRadio.onClick = function () {
            setSpacingInputEnabled(false);
        };

        var btnRowGroup = sortDialog.add("group");
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOk = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnCancel.alignment = "left";
        btnOk.alignment = "right";

        var sortOptions = null;
        btnOk.onClick = function () {
            for (var i = 0; i < fontRadios.length; i++) {
                if (fontRadios[i].button.value) {
                    sortOptions = {
                        font: fontRadios[i].key,
                        descending: descRadio.value,
                        random: randomRadio.value,
                        spacingMode: fitRadio.value ? "fit" : "custom",
                        spacingValue: parseFloat(spacingInput.text) * rulerUnit.pointsPerUnit
                    };
                    break;
                }
            }
            sortDialog.close();
        };
        btnCancel.onClick = function () {
            sortDialog.close();
        };

        prepareDialogWindow(sortDialog, SCRIPT_NAME);
        sortDialog.show();
        return sortOptions;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 数値を含むグループを選んだ順に並べ替え、縦に積み上げる
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {void}
     */
    function main(rulerUnit) {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var docSelection = app.activeDocument.selection;
        if (docSelection.length === 0) {
            alert("グループオブジェクトを選択してください。");
            return;
        }

        var numberedGroups = findNumberedGroups(docSelection);
        if (numberedGroups.length === 0) {
            alert("数値を含むグループが見つかりません。");
            return;
        }

        var fontMap = {};
        for (var i = 0; i < numberedGroups.length; i++) {
            collectNumbersByFont(numberedGroups[i], fontMap, numberedGroups[i]);
        }
        var stackOrigin = getStackOrigin(numberedGroups);

        var sortOptions = showFontChoiceDialog(fontMap, rulerUnit);
        if (!sortOptions) return;

        var groupEntries = getUniqueGroupEntries(fontMap[sortOptions.font]);
        if (sortOptions.random) {
            shuffleGroupEntries(groupEntries);
        } else {
            groupEntries.sort(function (a, b) {
                return sortOptions.descending ? b.value - a.value : a.value - b.value;
            });
        }

        var gap = FALLBACK_GAP;
        if (sortOptions.spacingMode === "custom" && !isNaN(sortOptions.spacingValue)) {
            gap = sortOptions.spacingValue;
        }
        stackGroups(groupEntries, stackOrigin, sortOptions.spacingMode === "fit", gap);

        app.redraw();
    }

    /* 定規の単位はドキュメントの有無を確かめる前に読む（従来どおり） / read before the document check, as before */
    main(getRulerUnit(app.activeDocument));

})();
