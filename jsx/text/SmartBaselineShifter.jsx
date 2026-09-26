#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストのベースラインシフトを調整します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBaselineShifter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n5e41727cf265

### Overview

Adjusts the baseline shift of the selected text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBaselineShifter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartBaselineShifter";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBaselineShifter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBaselineShifter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n5e41727cf265"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_REFERENCE_CHAR = "0"; /* 基準文字の初期値 / Initial reference character */
    var SHIFT_DECIMAL_PLACES   = 4;   /* 計算したシフト量の小数桁数 / Decimal places of the calculated shift amount */

    /* 対象文字の初期値から外す文字（空白・改行・英数字・ひらがな・カタカナ・漢字）
       Characters left out of the initial target (whitespace, line breaks, alphanumerics, kana, kanji) */
    var NON_SYMBOL_CHAR_PATTERN = /^[\x00-\x20 　A-Za-z0-9぀-ゟ゠-ヿ一-鿿]$/;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var INPUT_COLUMN_MARGINS       = [15, 5, 15, 5];  /* 入力欄の列の余白 [左,上,右,下] / Input column margins */
    var SHIFT_ROW_MARGINS          = [0, 0, 0, 10];   /* シフト量の行の余白 [左,上,右,下] / Shift amount row margins */
    var AUTO_ADJUST_PANEL_MARGINS  = [15, 20, 15, 5]; /* 自動調整パネルの余白 [左,上,右,下] / Auto adjust panel margins */
    var TEXT_INPUT_CHARACTERS      = 6;               /* 対象文字・シフト量の欄の幅（文字数）/ Width of the target and shift fields */
    var REFERENCE_INPUT_CHARACTERS = 3;               /* 基準文字の欄の幅（文字数）/ Width of the reference field */
    var CALCULATE_BUTTON_BOUNDS    = [0, 0, 60, 25];  /* 計算ボタンの大きさ / Calculate button bounds */
    var BUTTON_SPACER_BOUNDS       = [0, 0, 0, 30];   /* キャンセルとリセットの間隔 / Gap between Cancel and Reset */
    var DIALOG_OFFSET_X            = 300;             /* ダイアログを右へずらす量 / Horizontal dialog offset */
    var DIALOG_OPACITY             = 0.97;            /* ダイアログの不透明度 / Dialog opacity */

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
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "ベースライン調整", en: "Adjust Baseline" }
        },
        panel: {
            autoAdjust: { ja: "自動調整（天地）", en: "Auto Adjust (Vertical)" }
        },
        fieldLabel: {
            targetChars: { ja: "対象文字", en: "Target Character" },
            shiftAmount: { ja: "シフト量", en: "Shift Amount" },
            referenceChar: { ja: "基準文字", en: "Reference Character" }
        },
        button: {
            adjust: { ja: "調整", en: "Adjust" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" },
            calculate: { ja: "計算", en: "Calculate" }
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
            targetChars: {
                ja: "ベースラインをシフトする対象文字を入力します。",
                en: "Enter the character(s) to shift."
            },
            shiftAmount: {
                ja: "手動で指定するベースラインシフト量（数値）です。",
                en: "Specify the baseline shift amount manually."
            },
            referenceChar: {
                ja: "基準となる文字を1文字入力します。",
                en: "Enter the reference character (1 character)."
            },
            calculate: {
                ja: "対象文字と基準文字からシフト量を自動計算します。",
                en: "Calculate shift amount automatically."
            },
            adjust: {
                ja: "指定したシフト量を確定して適用します。",
                en: "Apply the specified shift amount."
            },
            reset: {
                ja: "選択しているテキストのベースラインシフトを全リセットします。",
                en: "Reset baseline shifts in all text frames."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            selectTextFrame: { ja: "テキストフレームを選択してください。", en: "Select one or more text frames." },
            invalidChars: {
                ja: "対象文字は1文字以上、基準文字は1文字を入力してください。",
                en: "Enter at least one target character and exactly one reference character."
            },
            targetNotFound: { ja: "対象文字が含まれていません。", en: "Target character not found." },
            shiftNotNumber: { ja: "シフト量は数値で入力してください。", en: "Shift amount must be a number." },
            errorPrefix: { ja: "エラー: ", en: "Error: " }
        }
    };

    /**
     * 現在の言語のラベルを取得する
     * @param {Object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} コロン付きのラベル文字列
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューで加えた変更を数えておき、app.undo() でまとめて取り消す
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;
    }

    /**
     * 変更を実行し、取り消す段数を1つ増やす
     * @param {function} changeFunc - 実行する変更
     * @returns {void}
     */
    PreviewManager.prototype.addStep = function (changeFunc) {
        /* プレビュー中の失敗はアラートを出さずに見送る / Preview failures are skipped without an alert */
        try {
            changeFunc();
            this.undoDepth++;
            app.redraw();
        } catch (e) { }
    };

    /**
     * プレビューで加えた変更をすべて取り消す
     * @returns {void}
     */
    PreviewManager.prototype.rollback = function () {
        while (this.undoDepth > 0) {
            /* 取り消せなくなったら段数を捨てて数え違いを残さない / If undo fails, drop the count so it cannot drift */
            try {
                app.undo();
            } catch (e) {
                this.undoDepth = 0;
                break;
            }
            this.undoDepth--;
        }
        app.redraw();
    };

    /**
     * プレビューを取り消してから本番の処理を1回だけ実行する（取り消しを1段にまとめる）
     * @param {function} finalAction - 本番の処理
     * @returns {void}
     */
    PreviewManager.prototype.confirm = function (finalAction) {
        this.rollback();
        finalAction();
    };

    // =========================================
    // 対象の収集 / Collecting targets
    // =========================================

    /**
     * 文字の編集中で、そのストーリーがテキストフレーム1つだけなら、そのフレームを選択し直す
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function selectFrameOfEditedText(doc) {
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.typename !== "TextRange") return;

        var storyFrames = currentSelection.story.textFrames;
        if (storyFrames.length !== 1) return;

        app.executeMenuCommand("deselectall");
        doc.selection = [storyFrames[0]];
        app.selectTool("Adobe Select Tool");
    }

    /**
     * アイテムがテキストフレームなら追加し、グループなら中を再帰的にたどる
     * @param {PageItem} pageItem - 調べるアイテム
     * @param {TextFrame[]} textFrames - 見つけたフレームを追加する配列
     * @returns {void}
     */
    function appendTextFrames(pageItem, textFrames) {
        if (pageItem.typename === "TextFrame") {
            textFrames.push(pageItem);
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                appendTextFrames(pageItem.pageItems[i], textFrames);
            }
        }
    }

    /**
     * 選択（グループの中を含む）からテキストフレームを集める
     * @param {PageItem[]|TextRange} currentSelection - ドキュメントの選択
     * @returns {TextFrame[]} テキストフレーム（文字の編集中は空）
     */
    function collectTextFrames(currentSelection) {
        var textFrames = [];
        if (!currentSelection || currentSelection.typename === "TextRange") return textFrames;
        for (var i = 0; i < currentSelection.length; i++) {
            appendTextFrames(currentSelection[i], textFrames);
        }
        return textFrames;
    }

    /**
     * 英数字・かな・漢字以外の文字を、出てきた順に重複なく集める（対象文字の初期値）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {string} 集めた文字
     */
    function collectSymbolChars(textFrames) {
        var symbolChars = "";
        for (var i = 0; i < textFrames.length; i++) {
            var frameText = textFrames[i].contents;
            for (var j = 0; j < frameText.length; j++) {
                var character = frameText.charAt(j);
                if (!NON_SYMBOL_CHAR_PATTERN.test(character) && symbolChars.indexOf(character) === -1) symbolChars += character;
            }
        }
        return symbolChars;
    }

    // =========================================
    // ベースラインシフト / Baseline shift
    // =========================================

    /**
     * テキストフレームのベースラインシフトをすべて0に戻す
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function resetBaselineShift(textFrames) {
        for (var i = 0; i < textFrames.length; i++) {
            /* 書き換えられないフレームは飛ばして残りを続ける / Skip frames that cannot be modified and carry on */
            try {
                textFrames[i].textRange.characterAttributes.baselineShift = 0;
            } catch (e) { }
        }
    }

    /**
     * フレーム内の対象文字すべてにベースラインシフトを設定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（複数可）
     * @param {number} shiftAmount - ベースラインシフトの値（pt）
     * @returns {void}
     */
    function applyBaselineShift(textFrame, targetChars, shiftAmount) {
        var frameCharacters = textFrame.textRange.characters;
        for (var i = 0; i < frameCharacters.length; i++) {
            var charText = frameCharacters[i].contents;
            if (charText && targetChars.indexOf(charText) !== -1) frameCharacters[i].characterAttributes.baselineShift = shiftAmount;
        }
    }

    /**
     * ベースラインシフトをいったん全部戻し、対象文字だけにシフト量を設定する
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（複数可）
     * @param {number} shiftAmount - ベースラインシフトの値（pt）
     * @returns {void}
     */
    function applyShiftToAll(textFrames, targetChars, shiftAmount) {
        resetBaselineShift(textFrames);
        for (var i = 0; i < textFrames.length; i++) {
            applyBaselineShift(textFrames[i], targetChars, shiftAmount);
        }
    }

    // =========================================
    // 文字の中心の実測 / Measuring character centers
    // =========================================

    /**
     * アイテムの天地中央のY座標を求める
     * @param {PageItem} pageItem - 対象のアイテム
     * @returns {number} 天地中央のY座標
     */
    function getCenterY(pageItem) {
        var bounds = pageItem.geometricBounds;
        return (bounds[1] + bounds[3]) / 2;
    }

    /**
     * テキストフレームを複製して1文字だけにし、アウトライン化した字形の天地中央を測る
     * @param {TextFrame} textFrame - 書式の元になるテキストフレーム
     * @param {string} character - 測る文字
     * @returns {number} 字形の天地中央のY座標
     */
    function measureCharCenterY(textFrame, character) {
        var tempFrame = textFrame.duplicate();
        var outlineGroup = null;
        try {
            tempFrame.contents = character;
            outlineGroup = tempFrame.createOutline(); /* 複製はここで消費される / The duplicate is consumed here */
            return getCenterY(outlineGroup);
        } finally {
            /* 途中で失敗しても一時オブジェクトを残さない / Never leave the temporary objects behind, even on failure */
            if (outlineGroup) outlineGroup.remove();
            else tempFrame.remove();
        }
    }

    /**
     * 対象文字を含む最初のフレームで、基準文字と対象文字の天地中央の差（シフト量）を求める
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（最初に見つかった1文字で測る）
     * @param {string} referenceChar - 基準文字（1文字）
     * @returns {number|null} シフト量（pt）。対象文字が見つからなければ null
     */
    function calculateShiftAmount(textFrames, targetChars, referenceChar) {
        for (var i = 0; i < textFrames.length; i++) {
            var frameText = textFrames[i].contents;
            for (var j = 0; j < targetChars.length; j++) {
                var targetChar = targetChars.charAt(j);
                if (frameText.indexOf(targetChar) === -1) continue;

                var referenceCenterY = measureCharCenterY(textFrames[i], referenceChar);
                return referenceCenterY - measureCharCenterY(textFrames[i], targetChar);
            }
        }
        return null;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名＋入力欄の1行を追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} initialText - 入力欄の初期値
     * @param {number} inputCharacters - 入力欄の幅（文字数）
     * @returns {EditText} 追加した入力欄（行のグループは parent で取れる）
     */
    function addFieldRow(parentContainer, labelSet, initialText, inputCharacters) {
        var fieldRow = parentContainer.add("group");
        fieldRow.add("statictext", undefined, labelText(labelSet));
        var fieldInput = fieldRow.add("edittext", undefined, initialText);
        fieldInput.characters = inputCharacters;
        return fieldInput;
    }

    /**
     * 左の列（対象文字・シフト量・自動調整パネル）を組む
     * @param {Group} parentContainer - 追加先
     * @param {string} defaultTargetChars - 対象文字の初期値
     * @param {string} shiftUnitLabel - シフト量の単位の表示
     * @returns {{targetInput: EditText, shiftInput: EditText, referenceInput: EditText, btnCalculate: Button}} 入力欄と計算ボタン
     */
    function buildInputColumn(parentContainer, defaultTargetChars, shiftUnitLabel) {
        var inputColumn = parentContainer.add("group");
        inputColumn.orientation = "column";
        inputColumn.alignChildren = "left";
        inputColumn.margins = INPUT_COLUMN_MARGINS;

        var targetInput = addFieldRow(inputColumn, LABELS.fieldLabel.targetChars, defaultTargetChars, TEXT_INPUT_CHARACTERS);
        targetInput.helpTip = getLabel(LABELS.tooltip.targetChars);

        /* シフト量は項目名・∧∨・入力欄・単位の1行。∧∨と入力欄は隙間0で突き合わせ、負の値も許す
           Shift row: label, stepper, field and unit; the stepper butts the field and negatives are allowed */
        var shiftRow = inputColumn.add("group");
        shiftRow.margins = SHIFT_ROW_MARGINS;
        shiftRow.add("statictext", undefined, labelText(LABELS.fieldLabel.shiftAmount));
        var shiftFieldGroup = shiftRow.add("group");
        shiftFieldGroup.orientation = "row";
        shiftFieldGroup.alignChildren = ["left", "center"];
        shiftFieldGroup.spacing = 0;
        shiftFieldGroup.margins = 0;
        var shiftInput;
        var shiftStepper = addStepper(shiftFieldGroup, function () { return shiftInput; }, {
            onStep: function (numberInput) { numberInput.notify("onChanging"); } /* プレビュー更新 / refresh preview */
        });
        shiftInput = shiftFieldGroup.add("edittext", undefined, "0");
        shiftInput.characters = TEXT_INPUT_CHARACTERS;
        shiftInput.helpTip = getLabel(LABELS.tooltip.shiftAmount);
        bindSteppedArrowKeys(shiftInput, shiftStepper); /* ↑↓キーも∧∨と同じ処理 / arrow keys share the stepper */
        shiftRow.add("statictext", undefined, shiftUnitLabel);
        shiftInput.active = true;

        var autoAdjustPanel = inputColumn.add("panel", undefined, getLabel(LABELS.panel.autoAdjust));
        autoAdjustPanel.orientation = "column";
        autoAdjustPanel.alignChildren = "left";
        autoAdjustPanel.margins = AUTO_ADJUST_PANEL_MARGINS;

        var referenceInput = addFieldRow(autoAdjustPanel, LABELS.fieldLabel.referenceChar, DEFAULT_REFERENCE_CHAR, REFERENCE_INPUT_CHARACTERS);
        referenceInput.helpTip = getLabel(LABELS.tooltip.referenceChar);
        var btnCalculate = referenceInput.parent.add("button", CALCULATE_BUTTON_BOUNDS, getLabel(LABELS.button.calculate));
        btnCalculate.helpTip = getLabel(LABELS.tooltip.calculate);

        return { targetInput: targetInput, shiftInput: shiftInput, referenceInput: referenceInput, btnCalculate: btnCalculate };
    }

    /**
     * 右の列（調整・キャンセル・リセット）を組む
     * @param {Group} parentContainer - 追加先
     * @returns {{btnOK: Button, btnCancel: Button, btnReset: Button}} ボタン
     */
    function buildButtonColumn(parentContainer) {
        var buttonColumn = parentContainer.add("group");
        buttonColumn.orientation = "column";
        buttonColumn.alignChildren = "fill";

        var btnOK = buttonColumn.add("button", undefined, getLabel(LABELS.button.adjust), { name: "ok" });
        btnOK.helpTip = getLabel(LABELS.tooltip.adjust);
        var btnCancel = buttonColumn.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });

        buttonColumn.add("statictext", BUTTON_SPACER_BOUNDS, " "); /* スペーサー / Spacer */
        var btnReset = buttonColumn.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);

        return { btnOK: btnOK, btnCancel: btnCancel, btnReset: btnReset };
    }

    /**
     * 対象文字とシフト量を指定するダイアログを表示する（入力のたびにプレビュー）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {PreviewManager} previewManager - プレビューの取り消し管理
     * @returns {{targetChars: string, shiftAmount: number}|null} 対象文字とシフト量（pt）。キャンセル時は null
     */
    function showShiftDialog(textFrames, previewManager) {
        /* シフト量の欄は環境設定の「東アジア言語」の単位で表示し、適用するときに pt へ換算する
           The shift field uses the East Asian type unit; values are converted to points when applied */
        var shiftUnit = getUnitInfo("text/asianunits");
        var shiftDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        shiftDialog.orientation = "column";
        shiftDialog.alignChildren = "left";
        shiftDialog.opacity = DIALOG_OPACITY;
        shiftDialog.onShow = function () {
            shiftDialog.location = [shiftDialog.location[0] + DIALOG_OFFSET_X, shiftDialog.location[1]];
        };

        var columnsGroup = shiftDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];

        var inputControls = buildInputColumn(columnsGroup, collectSymbolChars(textFrames), shiftUnit.label);
        var dialogButtons = buildButtonColumn(columnsGroup);
        var targetInput = inputControls.targetInput;
        var shiftInput = inputControls.shiftInput;
        var referenceInput = inputControls.referenceInput;

        /**
         * 直前のプレビューを取り消してから、いまの入力で掛け直す
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            if (!targetInput.text) return;

            var shiftAmount = parseFloat(shiftInput.text);
            if (isNaN(shiftAmount)) shiftAmount = 0;
            previewManager.addStep(function () {
                applyShiftToAll(textFrames, targetInput.text, shiftAmount * shiftUnit.pointsPerUnit);
            });
        }

        targetInput.onChanging = updatePreview;
        shiftInput.onChanging = updatePreview; /* ∧∨・↑↓キーも onChanging 経由で更新 / steppers and arrow keys go through onChanging */

        /* 基準文字との天地中央の差をシフト量に入れる / Put the center difference from the reference character into the shift field */
        inputControls.btnCalculate.onClick = function () {
            if (targetInput.text.length === 0 || referenceInput.text.length !== 1) {
                alert(getLabel(LABELS.alert.invalidChars));
                return;
            }
            var shiftAmount = calculateShiftAmount(textFrames, targetInput.text, referenceInput.text);
            if (shiftAmount === null) {
                alert(getLabel(LABELS.alert.targetNotFound));
                return;
            }

            shiftInput.text = (shiftAmount / shiftUnit.pointsPerUnit).toFixed(SHIFT_DECIMAL_PLACES);
            updatePreview();
        };

        /* ベースラインシフトを全部0に戻すのも、プレビューの1段として扱う / Resetting everything is also a single preview step */
        dialogButtons.btnReset.onClick = function () {
            previewManager.rollback();
            previewManager.addStep(function () {
                resetBaselineShift(textFrames);
            });
            shiftInput.text = "0";
        };

        /* 確定はダイアログを閉じてから main() で行う / The final apply happens in main() after the dialog closes */
        dialogButtons.btnOK.onClick = function () {
            if (targetInput.text.length === 0) {
                alert(getLabel(LABELS.alert.invalidChars));
                return;
            }
            if (isNaN(Number(shiftInput.text))) {
                alert(getLabel(LABELS.alert.shiftNotNumber));
                return;
            }
            shiftDialog.close(1);
        };

        if (shiftDialog.show() !== 1) return null;
        return { targetChars: targetInput.text, shiftAmount: Number(shiftInput.text) * shiftUnit.pointsPerUnit };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前提を確かめてダイアログを表示し、確定したシフト量を適用する
     * @returns {void}
     */
    function main() {
        try {
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }

            var doc = app.activeDocument;
            selectFrameOfEditedText(doc);

            var textFrames = collectTextFrames(doc.selection);
            if (textFrames.length === 0) {
                alert(getLabel(LABELS.alert.selectTextFrame));
                return;
            }

            var previewManager = new PreviewManager();
            var dialogResult = showShiftDialog(textFrames, previewManager);
            if (!dialogResult) {
                /* キャンセルでも閉じるボタンでもプレビューを戻す / Undo the preview on Cancel and on the close box */
                previewManager.rollback();
                return;
            }

            previewManager.confirm(function () {
                applyShiftToAll(textFrames, dialogResult.targetChars, dialogResult.shiftAmount);
            });
        } catch (e) {
            alert(getLabel(LABELS.alert.errorPrefix) + e);
        }
    }

    main();

})();
