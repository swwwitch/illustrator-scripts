#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した1つのテキストフレームの各行をアウトライン幅で測定し、最長行の幅にそろうよう行ごとの文字サイズを変倍します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DynamicTextGeneratorSimple.md

note記事も参照してください。
https://note.com/dtp_tranist/n/xxxxxxxx

### Overview

Measures each line of a selected text frame by its outline width and scales the per-line font size so every line matches the longest one.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGeneratorSimple.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DynamicTextGeneratorSimple";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DynamicTextGeneratorSimple.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGeneratorSimple.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 幅をそろえたあとに行送りを自動へ切り替えるか（ダイアログの初期値）
       Switch leading to auto after fitting the widths (initial state in the dialog) */
    var APPLY_AUTO_LEADING = true;

    /* 自動行送りの比率（％）。applyAutoLeading() の既定値でもある
       Auto-leading ratio in percent; also the default used by applyAutoLeading() */
    var AUTO_LEADING_AMOUNT = 100;

    /* ダイアログを開いた直後からプレビューを表示するか / Show the preview as soon as the dialog opens */
    var PREVIEW_ON_OPEN = true;

    /* 変倍率がこの範囲内なら誤差とみなして変更しない / Treat ratios within this range as no change */
    var RATIO_EPSILON = 0.001;

    /* ダイアログの透明度 / Dialog opacity */
    var DIALOG_OPACITY = 0.98;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS   = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING   = 8;                 /* パネル内の標準間隔 / default spacing inside panels */
    var FIELD_SPACING   = 6;                 /* 入力行どうしの間隔 / spacing between field rows */
    var FIELD_INDENT    = 20;                /* チェックボックスに従属する行の字下げ / indent for rows owned by a checkbox */
    var BUTTON_WIDTH    = 90;                /* OK・キャンセルの幅 / width of OK and Cancel */
    var EDIT_CHARACTERS = 4;                 /* 入力欄の文字数 / width of an edittext in characters */
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
        var upTooltip = stepOptions.integer ? "stepUpInteger" : "stepUp";
        var downTooltip = stepOptions.integer ? "stepDownInteger" : "stepDown";
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel("tooltip", upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel("tooltip", downTooltip);
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
            title: { ja: "行の幅そろえ", en: "Fit Lines to Widest" }
        },
        panel: {
            leading: { ja: "行送り", en: "Leading" }
        },
        checkbox: {
            autoLeading: { ja: "行送りを自動にする", en: "Switch leading to auto" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            autoLeadingAmount: { ja: "自動行送り", en: "Auto leading" }
        },
        unit: {
            percent: { ja: "%", en: "%" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            preview: {
                ja: "ONの間は結果を仮表示します。OFFにするかキャンセルすると元に戻ります。",
                en: "Shows a temporary result while enabled. Turning it off or cancelling restores the original."
            },
            autoLeading: {
                ja: "幅をそろえたあとに、行送りを自動へ切り替えます。",
                en: "Switches leading to auto after the widths are fitted."
            },
            autoLeadingAmount: {
                ja: "自動行送りの比率です。文字サイズに対する行送りの割合を指定します。",
                en: "The auto-leading ratio, given as a percentage of the font size."
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
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneTextFrame: { ja: "テキストフレームを1つだけ選択してください。", en: "Select exactly one text frame." },
            needTwoLines: { ja: "2行以上のテキストを選択してください。", en: "Select text with two or more lines." },
            noMeasurableLine: { ja: "有効なテキスト行が見つかりませんでした。", en: "No measurable text line was found." },
            previewError: { ja: "プレビュー更新エラー：", en: "Preview update error: " }
        }
    };

    /**
     * LABELS を上から順にたどって現在のUI言語のラベルを返す
     * @param {...string} - LABELS をたどるキー
     * @returns {string} ラベル文字列（見つからない場合は空文字）
     */
    function getLabel() {
        var node = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (node == null) break;
            node = node[arguments[i]];
        }
        return (node && node[uiLang] != null) ? node[uiLang] : "";
    }

    /**
     * コロン付きのラベルを返す（日本語は全角コロン、英語は半角コロン）
     * @param {...string} - LABELS をたどるキー
     * @returns {string} コロンを付けたラベル文字列
     */
    function getLabelWithColon() {
        return getLabel.apply(null, arguments) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 数値ユーティリティ / Number helpers
    // =========================================

    /**
     * 入力欄の文字列を数値として読み取る
     * @param {EditText} editText - 対象の入力欄
     * @param {number} fallback - 数値として読めない場合に返す値
     * @returns {number} 読み取った数値
     */
    function readNumber(editText, fallback) {
        var value = Number(editText.text);
        return (editText.text !== "" && !isNaN(value)) ? value : fallback;
    }

    // =========================================
    // UI部品 / UI helpers
    // =========================================

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループに共通のレイアウト設定を適用する
     * @param {Group} group - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル＋入力欄＋単位の行を追加する
     * @param {Panel|Group} parent - 追加先のコンテナ
     * @param {string} labelText - 行の見出し（コロン付き）
     * @param {string} initialText - 入力欄の初期値
     * @param {string} unitText - 単位表示の文字列
     * @param {Object} stepOptions - ∧∨と↑↓キーの増減設定（min / onStep など。addStepper() に渡す）
     * @returns {{label: StaticText, input: EditText, unit: StaticText, stepper: Group}} 生成した行の各コントロール
     */
    function addFieldRow(parent, labelText, initialText, unitText, stepOptions) {
        var row = parent.add("group");
        setupGroup(row, "row", FIELD_SPACING);
        row.margins = [FIELD_INDENT, 0, 0, 0];

        var label = row.add("statictext", undefined, labelText);

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = row.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var input;
        var stepper = addStepper(stepperInputGroup, function () { return input; }, stepOptions);
        input = stepperInputGroup.add("edittext", undefined, initialText);
        input.characters = EDIT_CHARACTERS;
        input.justify = "right";
        bindSteppedArrowKeys(input, stepper);
        var unit = row.add("statictext", undefined, unitText);

        return { label: label, input: input, unit: unit, stepper: stepper };
    }

    /**
     * 行（ラベル＋入力欄＋単位）をまとめて有効／無効にする
     * @param {{label: StaticText, input: EditText, unit: StaticText, stepper: Group}} row - addFieldRow が返した行オブジェクト
     * @param {boolean} enabled - 有効にするなら true
     * @returns {void}
     */
    function setFieldRowEnabled(row, enabled) {
        row.label.enabled = enabled;
        row.input.enabled = enabled;
        row.unit.enabled = enabled;
        row.stepper.enabled = enabled;
        redrawSteppersIn(row.stepper); /* 自作描画の∧∨を描き直す / redraw the custom-drawn stepper */
    }

    /**
     * 行（ラベル＋入力欄＋単位）にまとめてヘルプチップを設定する
     * @param {{label: StaticText, input: EditText, unit: StaticText}} row - addFieldRow が返した行オブジェクト
     * @param {string} tooltip - 設定するヘルプチップ文字列
     * @returns {void}
     */
    function setFieldRowTooltip(row, tooltip) {
        row.label.helpTip = tooltip;
        row.input.helpTip = tooltip;
        row.unit.helpTip = tooltip;
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
         * @param {function} func - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (func) {
            try {
                func();
                lastReportedError = "";
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                var message = getLabel("alert", "previewError") + e;
                if (message === lastReportedError) return;
                lastReportedError = message;
                alert(message);
            }
        };

        /**
         * 適用済みのプレビューをすべて取り消す
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try {
                    app.undo();
                } catch (e) {
                    /* Undoに失敗したらカウンターのずれを防ぐため打ち切る / stop on failure so the counter cannot drift */
                    this.undoDepth = 0;
                    break;
                }
                this.undoDepth--;
            }
            app.redraw();
        };
    }

    // =========================================
    // 測定と変倍 / Measure & scale
    // =========================================

    /**
     * 行の測定結果
     * @typedef {object} LineMetric
     * @property {number} start - ストーリー内の開始文字インデックス
     * @property {number} end - ストーリー内の終了文字インデックス（この位置は含まない）
     * @property {number} width - 行の外形幅（pt）。測定できない行は0
     */

    /**
     * 改行や空白しか含まない行かどうかを判定する
     * @param {TextRange} line - 判定する行
     * @returns {boolean} 内容が空とみなせる場合 true
     */
    function isBlankLine(line) {
        return line.contents.replace(/[\r\n\x03\s　]/g, "").length === 0;
    }

    /**
     * 行の内容を一時テキストフレームへ複製し、アウトライン化して外形幅を測る
     * 一時オブジェクトは成否にかかわらず必ず削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextRange} line - 測定する行
     * @returns {number} 行の外形幅（pt）。測定できない場合は0
     */
    function measureLineWidth(doc, line) {
        var tempTextFrame = null;
        var outlineGroup = null;
        var width = 0;

        try {
            tempTextFrame = doc.textFrames.add();
            line.duplicate(tempTextFrame, ElementPlacement.INSIDE);
            /* createOutline() は元のテキストフレームを消費するので参照を手放す
               createOutline() consumes the source frame, so drop the reference */
            outlineGroup = tempTextFrame.createOutline();
            tempTextFrame = null;
            width = outlineGroup.width;
        } catch (e) {
            /* アウトライン化できないときはフレーム幅で代用 / Fall back to the frame width */
            try {
                width = (tempTextFrame !== null) ? tempTextFrame.width : 0;
            } catch (err) {
                width = 0;
            }
        }

        /* 残っている一時オブジェクトを後始末する / Clean up whichever temporary object survived */
        try { if (outlineGroup !== null) outlineGroup.remove(); } catch (e) {}
        try { if (tempTextFrame !== null) tempTextFrame.remove(); } catch (e) {}

        return width;
    }

    /**
     * ストーリー内の指定範囲の文字サイズを変倍する（固定行送りの場合は行送りも追従させる）
     * @param {Story} story - 対象ストーリー
     * @param {number} startIndex - 開始文字インデックス
     * @param {number} endIndex - 終了文字インデックス（この位置は含まない）
     * @param {number} ratio - 変倍率
     * @returns {void}
     */
    function scaleCharacterSizes(story, startIndex, endIndex, ratio) {
        for (var i = startIndex; i < endIndex; i++) {
            try {
                var charAttr = story.characters[i].characterAttributes;
                charAttr.size *= ratio;
                /* 固定行送りのときだけ行が重ならないよう行送りも変倍する
                   Scale leading as well, but only when it is fixed */
                if (!charAttr.autoLeading) {
                    charAttr.leading *= ratio;
                }
            } catch (e) {
                /* 設定できない文字はスキップ / Skip characters that reject the change */
            }
        }
    }

    /**
     * 各行の文字サイズを最長行の幅にそろえる
     * 先に全行を測ってから変倍する（変倍で行が再合成されても対象がずれないようにするため）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {boolean} 1行でも測定できた場合 true
     */
    function fitLinesToWidestLine(doc, textFrame) {
        var lines = textFrame.lines;
        /** @type {LineMetric[]} */
        var lineMetrics = [];
        var maxWidth = 0;

        /* 1. 全行の外形幅と最大幅を測る（この時点ではテキストを変更しない）
           1. Measure every line and the widest width; the text is not touched yet */
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            var width = isBlankLine(line) ? 0 : measureLineWidth(doc, line);
            lineMetrics.push({ start: line.start, end: line.end, width: width });
            if (width > maxWidth) {
                maxWidth = width;
            }
        }

        if (maxWidth === 0) {
            return false;
        }

        /* 2. 行ごとに変倍する。行オブジェクトではなくストーリー内の文字インデックスで指定して
              変倍による行の再合成の影響を受けないようにする
           2. Scale line by line, addressing characters by story index rather than by line object
              so re-composition during scaling cannot shift the target range */
        var story = textFrame.story;
        for (var i = 0; i < lineMetrics.length; i++) {
            var metric = lineMetrics[i];
            if (metric.width <= 0) continue;

            var ratio = maxWidth / metric.width;
            /* 幅がほぼ同等の行はスキップ / Skip lines that already match the widest one */
            if (Math.abs(ratio - 1) < RATIO_EPSILON) continue;

            scaleCharacterSizes(story, metric.start, metric.end, ratio);
        }

        return true;
    }

    /**
     * テキストフレームの行送りを自動に切り替え、各段落の自動行送り比率を設定する
     * 他のスクリプトからも単体で呼べるよう、対象と比率を引数で受け取る
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number} autoLeadingAmount - 自動行送りの比率（％）。省略時は100
     * @returns {void}
     */
    function applyAutoLeading(textFrame, autoLeadingAmount) {
        var amount = (typeof autoLeadingAmount === "number" && !isNaN(autoLeadingAmount)) ? autoLeadingAmount : 100;

        /* フレーム全体を自動行送りにする / Switch the whole frame to auto leading */
        textFrame.textRange.characterAttributes.autoLeading = true;

        var paragraphs = textFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            try {
                paragraphs[i].characterAttributes.autoLeading = true;
                paragraphs[i].paragraphAttributes.autoLeadingAmount = amount;
            } catch (e) {
                /* 空段落など設定できないものはスキップ / Skip paragraphs that reject the setting */
            }
        }
    }

    /**
     * 幅そろえと行送りの設定をまとめて適用する（プレビューでも本適用でも共通）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {{applyAutoLeading: boolean, autoLeadingAmount: number}} settings - ダイアログの設定値
     * @returns {boolean} 1行でも測定できた場合 true
     */
    function applySettings(doc, textFrame, settings) {
        if (!fitLinesToWidestLine(doc, textFrame)) {
            return false;
        }
        if (settings.applyAutoLeading) {
            applyAutoLeading(textFrame, settings.autoLeadingAmount);
        }
        return true;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログのUIを構築する（イベントの配線は呼び出し側で行う）
     * @returns {object} ダイアログとコントロール群
     */
    function createSettingsDialog() {
        var dialog = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.opacity = DIALOG_OPACITY;

        var leadingPanel = dialog.add("panel", undefined, getLabel("panel", "leading"));
        setupPanel(leadingPanel, FIELD_SPACING);

        var autoLeadingCheckbox = leadingPanel.add("checkbox", undefined, getLabel("checkbox", "autoLeading"));
        autoLeadingCheckbox.value = APPLY_AUTO_LEADING;
        autoLeadingCheckbox.helpTip = getLabel("tooltip", "autoLeading");

        var amountRow = addFieldRow(
            leadingPanel,
            getLabelWithColon("fieldLabel", "autoLeadingAmount"),
            AUTO_LEADING_AMOUNT + "",
            getLabel("unit", "percent"),
            {
                min: 0,
                /* 手入力の確定と同じくプレビューを更新する / update the preview as a typed value would */
                onStep: function (numberInput) { numberInput.notify("onChange"); }
            }
        );
        setFieldRowTooltip(amountRow, getLabel("tooltip", "autoLeadingAmount"));
        setFieldRowEnabled(amountRow, autoLeadingCheckbox.value);

        /* 下段：左にプレビュー、右にキャンセル・OK / Footer: preview on the left, Cancel and OK on the right */
        var footerGroup = dialog.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = "fill";
        footerGroup.alignChildren = ["fill", "center"];

        var leftGroup = footerGroup.add("group");
        leftGroup.alignment = ["left", "center"];
        var previewCheckbox = leftGroup.add("checkbox", undefined, getLabel("checkbox", "preview"));
        previewCheckbox.value = PREVIEW_ON_OPEN;
        previewCheckbox.helpTip = getLabel("tooltip", "preview");

        /* 左右のボタンを両端に押し広げるスペーサー / spacer that pushes both sides apart */
        var spacerGroup = footerGroup.add("group");
        spacerGroup.alignment = ["fill", "center"];

        var rightGroup = footerGroup.add("group");
        rightGroup.alignment = ["right", "center"];
        var cancelButton = rightGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
        var okButton = rightGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" });
        cancelButton.preferredSize.width = BUTTON_WIDTH;
        okButton.preferredSize.width = BUTTON_WIDTH;

        return {
            dialog: dialog,
            autoLeadingCheckbox: autoLeadingCheckbox,
            amountRow: amountRow,
            previewCheckbox: previewCheckbox,
            cancelButton: cancelButton,
            okButton: okButton
        };
    }

    /**
     * 設定ダイアログを表示し、プレビューを見ながら結果を確定する
     * OKのときはプレビュー分を取り消してから1回だけ適用するので、Undoは1回で戻る
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function showSettingsDialog(doc, textFrame) {
        var ui = createSettingsDialog();
        var previewManager = new PreviewManager();

        /**
         * ダイアログの現在の入力値を読み取る
         * @returns {{applyAutoLeading: boolean, autoLeadingAmount: number}} 設定値
         */
        function readSettings() {
            var amount = readNumber(ui.amountRow.input, AUTO_LEADING_AMOUNT);
            /* 0以下や数値でない入力は既定値に読み替える / fall back to the default for non-positive or invalid input */
            if (!(amount > 0)) amount = AUTO_LEADING_AMOUNT;
            return {
                applyAutoLeading: ui.autoLeadingCheckbox.value,
                autoLeadingAmount: amount
            };
        }

        /**
         * プレビューを最新の設定で貼り直す（Undo履歴は汚さない）
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            if (!ui.previewCheckbox.value) return;
            previewManager.addStep(function () {
                applySettings(doc, textFrame, readSettings());
            });
        }

        ui.previewCheckbox.onClick = updatePreview;

        ui.autoLeadingCheckbox.onClick = function () {
            setFieldRowEnabled(ui.amountRow, ui.autoLeadingCheckbox.value);
            updatePreview();
        };

        ui.amountRow.input.onChange = updatePreview;

        ui.okButton.onClick = function () {
            /* プレビュー分を戻し、本適用を1回だけ実行して確定（Undo履歴は1つ）
               undo the preview, then apply once so it lands as a single undo entry */
            previewManager.rollback();
            var applied = applySettings(doc, textFrame, readSettings());
            ui.dialog.close(1);
            if (!applied) {
                alert(getLabel("alert", "noMeasurableLine"));
                return;
            }
            app.redraw();
        };

        ui.cancelButton.onClick = function () {
            /* 開いてから適用した分をすべて取り消してから閉じる / undo everything applied since open, then close */
            previewManager.rollback();
            ui.dialog.close(2);
        };

        updatePreview();
        ui.dialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = doc.selection;

        /* 選択オブジェクトのチェック / Validate the selection */
        if (!selectionItems || selectionItems.length !== 1 || selectionItems[0].typename !== "TextFrame") {
            alert(getLabel("alert", "selectOneTextFrame"));
            return;
        }

        var targetTextFrame = selectionItems[0];
        if (targetTextFrame.lines.length <= 1) {
            alert(getLabel("alert", "needTwoLines"));
            return;
        }

        showSettingsDialog(doc, targetTextFrame);
    }

    main();

})();
