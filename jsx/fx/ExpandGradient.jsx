#target illustrator
#targetengine "ExpandGradientEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトのグラデーションを、指定した数の単色オブジェクトに分割します。
分割後の後処理として、重なりを整理してひとまとめにする／両端からブレンドを作成する、を選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandGradient.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nbe084e691ba5

### Overview

Splits the gradient on the selected objects into a given number of solid-color objects.
Post-processing can tidy the overlaps into a single set or build a blend from the two end objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandGradient.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExpandGradient";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandGradient.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandGradient.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nbe084e691ba5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 分割数（既定値、2 以上の整数） / Default gradient step count (integer ≥ 2) */
    var DEFAULT_GRADIENT_STEPS = 5;

    /* 分割・拡張の補正値（Illustrator は指定値より1つ少ないオブジェクトを生成する） / Offset for Expand (Illustrator yields one object fewer than specified) */
    var EXPAND_STEP_OFFSET = 1;

    /* 「ブレンドに変換」時のステップ数（固定・ダイアログではディム表示） / Fixed step count for "Convert to blend" (dimmed in the dialog) */
    var BLEND_FIXED_STEPS = 2;

    /* 後処理モードの既定値（"none" / "simple" / "blend"） / Default post-process mode ("none" / "simple" / "blend") */
    var DEFAULT_POST_PROCESS_MODE = "simple";

    /* ダイアログ表示の有無（false で既定値のまま即実行） / Show dialog (false: run silently with default) */
    var SHOW_DIALOG = true;

    /* 一時アクションのセット名（衝突回避のためユニーク名） / Temporary action set name (unique to avoid collisions) */
    var ACTION_SET_NAME = "ExpandGradient_tmp";

    /* 一時アクションのアクション名 / Temporary action name */
    var ACTION_NAME = "Expand-gradient";

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
    // 6. この欄には別の↑↓キー増減処理を付けない（↑↓キーが二重に効く）
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
        var upTooltip = stepOptions.integer ? "stepUpInteger" : "stepUp";
        var downTooltip = stepOptions.integer ? "stepDownInteger" : "stepDown";
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

    /**
     * 表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {

        /* === 共通 / Common === */
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        /* === ダイアログ / Dialog === */
        dialogTitle: {
            ja: "グラデーションを分割・拡張",
            en: "Expand Gradient"
        },
        steps: {
            ja: "ステップ数",
            en: "Steps"
        },
        postProcessPanel: {
            ja: "実行後の処理",
            en: "Post-Processing"
        },
        postProcessNone: {
            ja: "なし",
            en: "None"
        },
        postProcessSimple: {
            ja: "単純に拡張",
            en: "Simple expand"
        },
        postProcessBlend: {
            ja: "ブレンドに変換",
            en: "Convert to blend"
        },

        /* === ツールチップ / Tooltips === */
        tipSteps: { ja: "グラデーションを何段階に分けるかです。多いほどなめらかになります。", en: "How many bands the gradient is split into. More bands look smoother." },
        tipPostProcessNone: { ja: "分割したまま、後処理は行いません。", en: "Leaves the expanded bands as they are." },
        tipPostProcessSimple: { ja: "隣り合う同色の帯をまとめて、パスの数を減らします。", en: "Merges neighbouring bands of the same color to cut down the number of paths." },
        tipPostProcessBlend: { ja: "両端の帯からブレンドを作り直して、中間をなめらかにします。", en: "Rebuilds a blend from the end bands so the middle stays smooth." },
        stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
        stepUp: {
            ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepDown: {
            ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },

        /* === アラート / Alerts === */
        alertNoDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        alertNoSelection: {
            ja: "オブジェクトを選択してください。",
            en: "Please select an object."
        },
        alertInvalidSteps: {
            ja: "ステップ数は 2 以上の整数で指定してください。",
            en: "Steps must be an integer of 2 or more."
        },
        alertMergeTargetNotFound: {
            ja: "Pathfinder Merge を適用できる対象を取得できませんでした。処理を中断します。",
            en: "Could not find a valid target for applying Pathfinder Merge. The process will stop."
        },
        alertBlendTargetNotFound: {
            ja: "ブレンドに必要なオブジェクトを取得できませんでした。処理を中断します。",
            en: "Could not get the objects needed for the blend. The process will stop."
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {string} key - LABELS のキー
     * @returns {string} ラベル文字列（見つからなければキーそのもの）
     */
    function getLabel(key) {
        return (LABELS[key] && LABELS[key][currentLanguage]) ? LABELS[key][currentLanguage] : key;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} key - LABELS のキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    (function () {

        // =========================================
        // 入口チェック / Entry checks
        // =========================================

        if (app.documents.length === 0) {
            alert(getLabel("alertNoDocument"));
            return;
        }

        var activeDoc = app.activeDocument;
        if (activeDoc.selection.length === 0) {
            alert(getLabel("alertNoSelection"));
            return;
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        var gradientSteps = DEFAULT_GRADIENT_STEPS;
        var postProcessMode = DEFAULT_POST_PROCESS_MODE;

        if (SHOW_DIALOG) {
            var dialogResult = showStepsDialog(SCRIPT_VERSION, DEFAULT_GRADIENT_STEPS, DEFAULT_POST_PROCESS_MODE);
            if (dialogResult === null) return;
            gradientSteps = dialogResult.steps;
            postProcessMode = dialogResult.postProcessMode;
        }

        playEmbeddedAction(buildExpandActionSource(gradientSteps + EXPAND_STEP_OFFSET, ACTION_SET_NAME, ACTION_NAME), ACTION_SET_NAME, ACTION_NAME);

        if (postProcessMode === "simple") {
            if (!furtherExpandSelection(activeDoc)) {
                alert(getLabel("alertMergeTargetNotFound"));
                return;
            }
        } else if (postProcessMode === "blend") {
            if (!convertToBlend(activeDoc)) {
                alert(getLabel("alertBlendTargetNotFound"));
                return;
            }
        }

    })();

    // =========================================
    // 後処理 / Post-processing
    // =========================================

    /**
     * 分割・拡張の結果をクロップし、Pathfinder Merge で同色の重なりを整理する。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {boolean} 整理できたら true、対象が見つからなければ false
     */
    function furtherExpandSelection(targetDoc) {
        targetDoc.activate();
        app.executeMenuCommand('Live Pathfinder Crop');
        app.executeMenuCommand('expandStyle');

        /* Pathfinder Merge ライブエフェクト（command 8）を適用し、アピアランスを分割
           Apply the Pathfinder Merge live effect (command 8) and expand appearance */
        var pathfinderMergeXml = '<LiveEffect name="Adobe Pathfinder" isPre="1">'
            + '<Dict data="I Command 8 B ConvertCustom 1 B ExtractUnpainted 1 R Mix 0.5 R Precision 10 B RemovePoints 1 R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0 R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">'
            + '<Entry name="DisplayString" value="Merge" valueType="S"/>'
            + '</Dict></LiveEffect>';

        var mergeTarget = getFirstEffectApplicableSelectionItem(targetDoc);
        if (mergeTarget === null) return false;

        mergeTarget.applyEffect(pathfinderMergeXml);
        app.redraw();
        app.executeMenuCommand("deselectall");
        mergeTarget.selected = true;
        app.executeMenuCommand('expandStyle');

        return true;
    }

    /**
     * 分割結果の両端だけを残し、ブレンドに置き換える。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {boolean} 変換できたら true、対象が足りなければ false
     */
    function convertToBlend(targetDoc) {
        /* furtherExpandSelection と同じ前処理（Crop → expandStyle → Merge → expandStyle）
           Same pre-processing as furtherExpandSelection */
        if (!furtherExpandSelection(targetDoc)) return false;

        var endPaths = reduceToEndPaths(targetDoc);
        if (endPaths === null) return false;

        /* 2点を選択してブレンド作成 / Select the two items and run Blend Make */
        app.executeMenuCommand("deselectall");
        endPaths.backItem.selected = true;
        endPaths.frontItem.selected = true;
        app.executeMenuCommand('Path Blend Make');

        return true;
    }

    /**
     * 選択配下の塗りパスから最前面・最背面だけを残し、レイヤー直下へ移動する。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {object} frontItem / backItem を持つオブジェクト、2つ揃わなければ null
     */
    function reduceToEndPaths(targetDoc) {
        /* 包んでいるグループ／コンパウンドは後で片付けるため記録
           Remember the wrapping containers so we can drop them later */
        var originalContainers = [];
        var paintedPaths = [];
        for (var i = 0; i < targetDoc.selection.length; i++) {
            originalContainers.push(targetDoc.selection[i]);
            collectPaintedDescendantPaths(targetDoc.selection[i], paintedPaths);
        }
        if (paintedPaths.length < 2) return null;

        /* 親共通なら親内 z 順（前面=0）で並べ替え / Sort by parent's z-order when parents are shared */
        paintedPaths = sortByZOrderWhenShared(paintedPaths);

        var frontItem = paintedPaths[0];
        var backItem = paintedPaths[paintedPaths.length - 1];
        for (var i = 1; i < paintedPaths.length - 1; i++) {
            paintedPaths[i].remove();
        }

        /* front/back をレイヤー直下へ移動し、空になった包みを除去
           Lift front/back to the layer and drop the now-empty wrappers */
        var hostLayer = targetDoc.activeLayer;
        frontItem.move(hostLayer, ElementPlacement.PLACEATEND);
        backItem.move(hostLayer, ElementPlacement.PLACEATEND);
        for (var i = 0; i < originalContainers.length; i++) {
            /* 選択項目自体が両端のパスだった場合は消さない / Keep the end paths when they were selected directly */
            if (originalContainers[i] === frontItem || originalContainers[i] === backItem) continue;
            /* 移動で空になった包みは既に消えている場合がある / The wrapper may already be gone */
            try { originalContainers[i].remove(); } catch (e) { }
        }

        return { frontItem: frontItem, backItem: backItem };
    }

    /**
     * 塗りまたは線を持つ葉パスを再帰的に収集する。
     * @param {object} item - 走査対象のページアイテム
     * @param {object[]} result - 収集先の配列
     * @returns {void}
     */
    function collectPaintedDescendantPaths(item, result) {
        if (!item) return;
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectPaintedDescendantPaths(item.pageItems[i], result);
            }
        } else if (item.typename === "CompoundPathItem") {
            for (var i = 0; i < item.pathItems.length; i++) {
                collectPaintedDescendantPaths(item.pathItems[i], result);
            }
        } else if (item.typename === "PathItem") {
            if (item.clipping) return;
            if (!item.filled && !item.stroked) return;
            if (item.pathPoints && item.pathPoints.length === 2) return;
            result.push(item);
        }
    }

    /**
     * 親が共通のときだけ、親内の z 順（前面が先頭）に並べ替える。
     * @param {object[]} paths - 並べ替えるパスの配列
     * @returns {object[]} 並べ替えた配列、親がばらけていれば元の配列
     */
    function sortByZOrderWhenShared(paths) {
        if (paths.length < 2) return paths;

        var sharedParent = paths[0].parent;
        for (var i = 1; i < paths.length; i++) {
            if (paths[i].parent !== sharedParent) return paths;
        }

        /* 親の pageItems は前面から後面の順なので、その順に拾い直す
           The parent's pageItems run front to back, so collect in that order */
        var sorted = [];
        for (var i = 0; i < sharedParent.pageItems.length; i++) {
            for (var j = 0; j < paths.length; j++) {
                if (paths[j] === sharedParent.pageItems[i]) {
                    sorted.push(paths[j]);
                    break;
                }
            }
        }

        return (sorted.length === paths.length) ? sorted : paths;
    }

    /**
     * 選択中でライブ効果を適用できる最初のアイテムを返す。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {object} 適用できるアイテム、見つからなければ null
     */
    function getFirstEffectApplicableSelectionItem(targetDoc) {
        if (!targetDoc || !targetDoc.selection || targetDoc.selection.length === 0) return null;

        for (var i = 0; i < targetDoc.selection.length; i++) {
            var selectedItem = targetDoc.selection[i];
            if (selectedItem && typeof selectedItem.applyEffect === "function") {
                return selectedItem;
            }
        }

        return null;
    }

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * ステップ数と実行後の処理を指定するダイアログを表示する。
     * @param {string} scriptVersion - タイトルに表示するバージョン
     * @param {number} defaultSteps - ステップ数の初期値
     * @param {string} defaultPostProcessMode - 実行後の処理の初期値（"none" / "simple" / "blend"）
     * @returns {object} steps / postProcessMode を持つオブジェクト、キャンセル時は null
     */
    function showStepsDialog(scriptVersion, defaultSteps, defaultPostProcessMode) {

        var stepsDialog = new Window("dialog", getLabel("dialogTitle") + " " + scriptVersion);
        stepsDialog.orientation = "column";
        stepsDialog.alignChildren = ["fill", "top"];
        stepsDialog.margins = 16;

        var stepsRow = stepsDialog.add("group");
        stepsRow.orientation = "row";
        stepsRow.alignChildren = ["left", "center"];
        stepsRow.add("statictext", undefined, labelText("steps"));
        /* ∧∨と入力欄は隙間0で突き合わせる。2以上の整数 / butt the stepper against the field; integers of 2 or more */
        var stepperInputGroup = stepsRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var stepsInput;
        var stepsStepperGroup = addStepper(stepperInputGroup, function () { return stepsInput; }, { step: 1, min: 2, integer: true });
        stepsInput = stepperInputGroup.add("edittext", undefined, String(defaultSteps));
        stepsInput.helpTip = getLabel("tipSteps");
        stepsInput.characters = 5;
        stepsInput.active = true;
        bindSteppedArrowKeys(stepsInput, stepsStepperGroup);

        var postProcessPanel = stepsDialog.add("panel", undefined, getLabel("postProcessPanel"));
        postProcessPanel.orientation = "column";
        postProcessPanel.alignChildren = ["left", "top"];
        postProcessPanel.margins = [15, 20, 15, 10];

        var postProcessNoneRb = postProcessPanel.add("radiobutton", undefined, getLabel("postProcessNone"));
        postProcessNoneRb.helpTip = getLabel("tipPostProcessNone");
        var postProcessSimpleRb = postProcessPanel.add("radiobutton", undefined, getLabel("postProcessSimple"));
        postProcessSimpleRb.helpTip = getLabel("tipPostProcessSimple");
        var postProcessBlendRb = postProcessPanel.add("radiobutton", undefined, getLabel("postProcessBlend"));
        postProcessBlendRb.helpTip = getLabel("tipPostProcessBlend");

        postProcessSimpleRb.value = (defaultPostProcessMode === "simple");
        postProcessBlendRb.value = (defaultPostProcessMode === "blend");
        postProcessNoneRb.value = (!postProcessSimpleRb.value && !postProcessBlendRb.value);

        /* 「ブレンドに変換」はステップ数を固定し、入力欄をディムにする（戻したときは元の値へ）
           Convert to blend fixes the step count and dims the field; the previous value returns on switch back */
        var keptStepsText = String(defaultSteps);
        /**
         * 後処理の選択に合わせて、ステップ数の欄と∧∨を有効／無効にする
         * @returns {void}
         */
        function updateStepsAvailability() {
            var isBlend = postProcessBlendRb.value;
            if (isBlend) {
                if (stepsInput.enabled) keptStepsText = stepsInput.text;
                stepsInput.text = String(BLEND_FIXED_STEPS);
            } else if (!stepsInput.enabled) {
                stepsInput.text = keptStepsText;
            }
            stepsInput.enabled = !isBlend;
            stepsStepperGroup.enabled = !isBlend;
            redrawSteppersIn(stepsStepperGroup);
        }
        postProcessNoneRb.onClick = updateStepsAvailability;
        postProcessSimpleRb.onClick = updateStepsAvailability;
        postProcessBlendRb.onClick = updateStepsAvailability;
        updateStepsAvailability();

        var okCancelGroup = stepsDialog.add("group");
        okCancelGroup.alignment = ["right", "center"];
        okCancelGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
        okCancelGroup.add("button", undefined, "OK", { name: "ok" });

        prepareDialogWindow(stepsDialog, SCRIPT_NAME);
        if (stepsDialog.show() !== 1) return null;

        var parsedSteps = parseInt(stepsInput.text, 10);
        if (isNaN(parsedSteps) || parsedSteps < 2) {
            alert(getLabel("alertInvalidSteps"));
            return null;
        }

        var selectedMode = "none";
        if (postProcessSimpleRb.value) selectedMode = "simple";
        else if (postProcessBlendRb.value) selectedMode = "blend";

        return { steps: parsedSteps, postProcessMode: selectedMode };

    }

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

    /**
     * 分割・拡張（ai_plugin_expand）のアクションソースを組み立てる。
     * @param {number} gradientSteps - グラデーションの分割数
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @returns {string} .aia のソース
     */
    function buildExpandActionSource(gradientSteps, setName, actionName) {
        var parameters = [
            buildParameterLine(1, 1868720756, "boolean", 0, ""),
            buildParameterLine(2, 1718185068, "boolean", 1, ""),
            buildParameterLine(3, 1937011307, "boolean", 0, ""),
            buildParameterLine(4, 1936553064, "boolean", 0, ""),
            buildParameterLine(5, 1937007984, "integer", gradientSteps, "")
        ];
        /* 表示名「分割・拡張」 / Localized name */
        return buildActionSource(setName, actionName, "ai_plugin_expand", "e58886e589b2e383bbe68ba1e5bcb5", parameters);
    }

    /**
     * 1イベントだけのアクションセットを組み立てる。
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} internalName - プラグインの内部名
     * @param {string} localizedNameHex - パネル表示名のUTF-8 16進表現
     * @param {string[]} parameters - パラメーター行の配列
     * @returns {string} .aia のソース
     */
    function buildActionSource(setName, actionName, internalName, localizedNameHex, parameters) {
        return ''
            + '/version 3'
            + buildActionNameLine(setName)
            + '/isOpen 1'
            + '/actionCount 1'
            + '/action-1 {'
            + ' ' + buildActionNameLine(actionName)
            + ' /keyIndex 0'
            + ' /colorIndex 0'
            + ' /isOpen 1'
            + ' /eventCount 1'
            + ' /event-1 {'
            + ' /useRulersIn1stQuadrant 0'
            + ' /internalName (' + internalName + ')'
            + ' /localizedName ' + buildHexTextBlock(localizedNameHex)
            + ' /isOpen 1'
            + ' /isOn 1'
            + ' /hasDialog 1'
            + ' /showDialog 0'
            + ' /parameterCount ' + parameters.length
            + parameters.join('')
            + ' }'
            + '}';
    }

    /**
     * アクションのパラメーター1行を組み立てる。
     * @param {number} index - パラメーター番号（1始まり）
     * @param {number} key - パラメーターキー
     * @param {string} type - 値の型（boolean / integer / enumerated）
     * @param {number} value - 値
     * @param {string} localizedNameHex - 表示名のUTF-8 16進表現（不要なら空文字）
     * @returns {string} パラメーター行
     */
    function buildParameterLine(index, key, type, value, localizedNameHex) {
        return ' /parameter-' + index
            + ' { /key ' + key
            + ' /showInPalette 4294967295'
            + ' /type (' + type + ')'
            + (localizedNameHex ? ' /name ' + buildHexTextBlock(localizedNameHex) : '')
            + ' /value ' + value + ' }';
    }

    /**
     * アクション名の行を組み立てる。
     * @param {string} actionName - アクション名（ASCII）
     * @returns {string} /name の行
     */
    function buildActionNameLine(actionName) {
        return '/name ' + buildHexTextBlock(stringToHex(actionName)) + '\n';
    }

    /**
     * .aia の文字列表記 [ バイト数 16進 ] を組み立てる。
     * @param {string} hexText - 16進表現
     * @returns {string} [ バイト数 16進 ] の形式
     */
    function buildHexTextBlock(hexText) {
        return '[ ' + (hexText.length / 2) + ' ' + hexText + ' ]';
    }

    /**
     * 文字列を16進表現に変換する。
     * @param {string} sourceText - 変換元の文字列（ASCII）
     * @returns {string} 16進表現
     */
    function stringToHex(sourceText) {
        var hexText = "";
        for (var i = 0; i < sourceText.length; i++) {
            var hexValue = sourceText.charCodeAt(i).toString(16);
            if (hexValue.length < 2) hexValue = "0" + hexValue;
            hexText += hexValue;
        }
        return hexText;
    }

    // =========================================
    // 一時アクション実行 / Temporary action playback
    // =========================================

    /**
     * 一時的な .aia を書き出してロード・再生し、後始末まで行う。
     * @param {string} actionSource - .aia のソース
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @returns {void}
     */
    function playEmbeddedAction(actionSource, setName, actionName) {

        var actionFile = new File('~/ExpandGradientAction.aia');

        /* 既存の同名セットが残っているとロードが効かないので、先に外す / Unload first in case the set is still loaded */
        try { app.unloadAction(setName, ""); } catch (e) { }

        try {
            if (!actionFile.open('w')) {
                throw new Error('Failed to open temporary action file for writing.');
            }
            actionFile.write(actionSource);
            actionFile.close();

            app.loadAction(actionFile);
            app.doScript(actionName, setName, false);
        } finally {
            /* close / remove は失敗しても false を返すだけ / close and remove just return false on failure */
            actionFile.close();
            actionFile.remove();
            try { app.unloadAction(setName, ""); } catch (e) { }
        }

    }

})();
