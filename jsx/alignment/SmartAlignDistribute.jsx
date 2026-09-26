#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを縦または横に並べ、指定した間隔で分布します。
方向は自動判定でき、揃え（左右／上下）、プレビュー境界、ランダム並べ替えにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md

### Overview

Lines the selected objects up vertically or horizontally and distributes them at the spacing you specify.
The direction can be detected automatically, and alignment, preview bounds and random reordering are all supported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignDistribute";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビューを再描画する最小間隔（ミリ秒）/ Minimum interval between preview renders (ms) */
    var PREVIEW_MIN_INTERVAL_MS = 80;

    /* 操作からプレビューを実行するまでの待ち時間（ミリ秒）/ Delay before a requested preview runs (ms) */
    var PREVIEW_SCHEDULE_MS = 60;

    // =========================================
    // 内部キー / Internal keys
    // =========================================

    /* ダイアログ位置をセッション内に控える $.global のキー / $.global key for the session dialog position */
    var DIALOG_POSITION_KEY = "__SmartAlignDistribute_DialogPosition__";

    /* 計測用に一時的に作るグループの名前 / name of the throwaway measuring group */
    var TEMP_MEASURE_GROUP_NAME = "__SmartAlignDistribute_TempMeasure__";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OPACITY = 0.97;               /* ダイアログの不透明度 / dialog opacity */
    var PANEL_MARGINS = [15, 20, 15, 10];    /* パネルの余白 [左,上,右,下] / panel margins */
    var OPTIONS_MARGINS = [15, 5, 15, 5];    /* オプション欄の余白 / options group margins */
    var SPACING_FIELD_CHARS = 3;             /* 間隔の入力欄の幅（文字数）/ width of the spacing field */

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
        if (direction > 0) return Math.floor(value / multiple) * multiple + multiple;
        return Math.ceil(value / multiple) * multiple - multiple;
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
     * 現在のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },
        panel: {
            direction: { ja: "方向", en: "Direction" },
            spacing: { ja: "間隔", en: "Spacing" },
            alignHorizontal: { ja: "揃え（左右）", en: "Align (H)" },
            alignVertical: { ja: "揃え（上下）", en: "Align (V)" }
        },
        radio: {
            directionAuto: { ja: "自動", en: "Auto" },
            directionVertical: { ja: "縦", en: "Vertical" },
            directionHorizontal: { ja: "横", en: "Horizontal" },
            alignNone: { ja: "なし", en: "None" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignTop: { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        checkbox: {
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            measureText: { ja: "テキストの高さを計測", en: "Measure text height" },
            random: { ja: "ランダム", en: "Random" }
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
            directionAuto: {
                ja: "選択範囲が横長なら横並び、縦長なら縦並びとして扱います。",
                en: "Lays the objects out in a row when the selection is wider than tall, in a column otherwise."
            },
            directionVertical: { ja: "上から下へ縦に並べます。", en: "Stacks the objects from top to bottom." },
            directionHorizontal: { ja: "左から右へ横に並べます。", en: "Lays the objects out from left to right." },
            spacing: {
                ja: "オブジェクト間のすき間。マイナス値で重ねられます。",
                en: "Gap between objects. Negative values overlap them."
            },
            alignHorizontal: {
                ja: "縦に並べたときの左右の揃え方です。N／L／C／R キーでも切り替えられます。",
                en: "Horizontal alignment used when stacking vertically. The keys N / L / C / R switch it."
            },
            alignVertical: {
                ja: "横に並べたときの上下の揃え方です。N／T／M／B キーでも切り替えられます。",
                en: "Vertical alignment used when laying out horizontally. The keys N / T / M / B switch it."
            },
            usePreviewBounds: {
                ja: "線や効果を含む見た目の境界で整列します。オフはパスのみのジオメトリ境界。",
                en: "Align by visible bounds (incl. strokes/effects). Off uses geometric (path-only) bounds."
            },
            measureText: {
                ja: "縦並び時のみ、テキストを一度だけ複製→アウトライン化して境界を計測します（ダイアログ中だけキャッシュ）。",
                en: "Only in vertical layout, measures text by duplicating and outlining once (cached for this dialog only)."
            },
            random: {
                ja: "並び順をランダムに入れ替えます（左上の位置は維持）。",
                en: "Shuffle the stacking order at random (top-left position is kept)."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            errorPrefix: { ja: "エラーが発生しました: ", en: "An error has occurred: " }
        }
    };

    /**
     * ドット区切りのパスで LABELS から現在の言語の文字列を取り出す（{slash} は "/" に置き換える）
     * @param {string} labelPath - "panel.direction" のようなドット区切りのキー
     * @returns {string} 現在の言語の文字列（見つからなければ labelPath そのもの）
     */
    function getLabel(labelPath) {
        var labelNode = LABELS;
        var pathKeys = labelPath.split(".");
        for (var i = 0; i < pathKeys.length; i++) {
            if (labelNode == null) return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        if (labelNode == null) return labelPath;
        var localizedText = (labelNode[uiLang] != null) ? labelNode[uiLang] : labelNode.en;
        if (localizedText == null) return labelPath;
        return String(localizedText).replace(/\{slash\}/g, "/");
    }

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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューを作り直す（前回分を undo してから processFn を実行）
     * @param {{isUndo: boolean}} previewState - プレビューの状態（undo すべき変更があるか）
     * @param {function} processFn - プレビューとして実行する処理
     * @param {boolean} isEnabled - プレビューを表示するか（false なら前回分を戻すだけ）
     * @returns {void}
     */
    function runPreview(previewState, processFn, isEnabled) {
        /* app.undo() と DOM 操作は状況によって失敗する / undo and DOM edits may fail */
        try {
            if (isEnabled) {
                if (previewState.isUndo) app.undo();
                else previewState.isUndo = true;
                processFn();
                app.redraw();
            } else if (previewState.isUndo) {
                app.undo();
                app.redraw();
                previewState.isUndo = false;
            }
        } catch (err) { }
    }

    /**
     * プレビュー分を巻き戻す（確定処理の直前・キャンセル時）
     * @param {{isUndo: boolean}} previewState - プレビューの状態
     * @returns {void}
     */
    function undoPreview(previewState) {
        /* 取り消す履歴が無いと app.undo() が失敗することがある / undo may fail with no history */
        try {
            if (previewState.isUndo) app.undo();
        } catch (err) { }
        previewState.isUndo = false;
    }

    /**
     * セッション内に控えたダイアログ位置を返す
     * @returns {number[]|null} [x, y]（控えが無ければ null）
     */
    function loadDialogPosition() {
        var savedPosition = $.global[DIALOG_POSITION_KEY];
        return (savedPosition && savedPosition.length === 2) ? [savedPosition[0], savedPosition[1]] : null;
    }

    /**
     * ダイアログ位置をセッション内に控える
     * @param {number[]} dialogLocation - ダイアログの位置 [x, y]
     * @returns {void}
     */
    function saveDialogPosition(dialogLocation) {
        if (!dialogLocation || dialogLocation.length !== 2) return;
        $.global[DIALOG_POSITION_KEY] = [Math.round(dialogLocation[0]), Math.round(dialogLocation[1])];
    }

    // =========================================
    // 境界と並び順 / Bounds and ordering
    // =========================================

    /**
     * クリップグループのクリッピングパスを返す
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {PageItem|null} クリッピングパス（クリップグループでなければ null）
     */
    function getClippingPath(targetItem) {
        if (!targetItem || targetItem.typename !== "GroupItem" || !targetItem.clipped) return null;

        var childItems = targetItem.pageItems;
        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];
            /* clipping を持たない種類のアイテムが混ざると例外になる / some item kinds do not expose clipping */
            try {
                if (childItem.clipping === true) return childItem;
            } catch (e) { }
            /* 複合パスはマスク本体に clipping が無く、内部パスに付く / Compound path: clipping flag sits on inner path, not the wrapper */
            if (childItem.typename === "CompoundPathItem" && childItem.pathItems && childItem.pathItems.length) {
                try {
                    if (childItem.pathItems[0].clipping === true) return childItem;
                } catch (e) { }
            }
        }
        return null;
    }

    /**
     * オブジェクトの境界を返す（クリップグループはクリッピングパスを測る）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界（visibleBounds）を使うか
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(targetItem, usePreviewBounds) {
        var clipPath = getClippingPath(targetItem);
        var measuredItem = clipPath ? clipPath : targetItem;
        return usePreviewBounds ? measuredItem.visibleBounds : measuredItem.geometricBounds;
    }

    /**
     * 選択全体の幅・高さを返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {{spanX: number, spanY: number}} 幅と高さ
     */
    function getSelectionSpan(targetItems, usePreviewBounds) {
        var minLeft = null, maxRight = null, maxTop = null, minBottom = null;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getItemBounds(targetItem, usePreviewBounds);
            if (minLeft === null || bounds[0] < minLeft) minLeft = bounds[0];
            if (maxRight === null || bounds[2] > maxRight) maxRight = bounds[2];
            if (maxTop === null || bounds[1] > maxTop) maxTop = bounds[1];
            if (minBottom === null || bounds[3] < minBottom) minBottom = bounds[3];
        }
        if (minLeft === null) return { spanX: 0, spanY: 0 };
        return { spanX: (maxRight - minLeft), spanY: (maxTop - minBottom) };
    }

    /**
     * 選択の縦横比から並べる方向を判定する（geometricBounds で安定させる）
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {string} "horizontal" または "vertical"
     */
    function detectDirection(targetItems) {
        var selectionSpan = getSelectionSpan(targetItems, false);
        return (selectionSpan.spanX >= selectionSpan.spanY) ? "horizontal" : "vertical";
    }

    /**
     * 上→下（同じ高さなら左→右）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortTopToBottom(targetItems) {
        var sortedItems = targetItems.slice();
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.top !== itemB.top) return itemB.top - itemA.top;
            return itemA.left - itemB.left;
        });
        return sortedItems;
    }

    /**
     * 左→右（同じ位置なら上→下）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortLeftToRight(targetItems) {
        var sortedItems = targetItems.slice();
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.left !== itemB.left) return itemA.left - itemB.left;
            return itemB.top - itemA.top;
        });
        return sortedItems;
    }

    /**
     * 0〜count-1 のインデックスを Fisher-Yates でシャッフルした配列を返す
     * @param {number} count - 要素数
     * @returns {number[]} シャッフルしたインデックス
     */
    function makeShuffledIndices(count) {
        var indices = [];
        for (var i = 0; i < count; i++) indices.push(i);
        for (var j = indices.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var swappedIndex = indices[j]; indices[j] = indices[k]; indices[k] = swappedIndex;
        }
        return indices;
    }

    /**
     * オブジェクト（または子孫）がテキストを含むかを返す
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {boolean} テキストを含めば true
     */
    function containerHasText(targetItem) {
        if (!targetItem) return false;
        /* textFrames / pageItems を持たない種類がある / not every item kind has textFrames or pageItems */
        try {
            if (targetItem.typename === "TextFrame") return true;
            if (targetItem.textFrames && targetItem.textFrames.length > 0) return true;
            if (targetItem.pageItems && targetItem.pageItems.length) {
                for (var i = 0; i < targetItem.pageItems.length; i++) {
                    if (containerHasText(targetItem.pageItems[i])) return true;
                }
            }
        } catch (e) { }
        return false;
    }

    /**
     * いずれかのオブジェクトがテキストを含むかを返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {boolean} テキストを含むものがあれば true
     */
    function anyItemHasText(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            if (containerHasText(targetItems[i])) return true;
        }
        return false;
    }

    // =========================================
    // キー操作 / Keyboard
    // =========================================

    /* 揃えのショートカットキーと、対応するラジオのキー / Shortcut keys mapped to a radio in each row */
    var HORIZONTAL_ALIGN_RADIO_BY_KEY = { "N": "none", "L": "left", "C": "center", "R": "right" };
    var VERTICAL_ALIGN_RADIO_BY_KEY = { "N": "none", "T": "top", "M": "middle", "B": "bottom" };

    /**
     * N / L / C / R / T / M / B キーで揃えのラジオを切り替える
     * @param {object} keyTarget - キーを受けるダイアログまたはコントロール
     * @param {{horizontal: Object, vertical: Object, getDirection: function}} alignRadios - 揃えのラジオ一式と方向の取得関数
     * @param {function} [onUpdate] - 切り替えたあとに呼ぶ処理
     * @returns {void}
     */
    function addAlignKeyHandler(keyTarget, alignRadios, onUpdate) {
        keyTarget.addEventListener("keydown", function (event) {
            /* 横並びのときは上下の揃え、縦並びのときは左右の揃えを操作する
               A horizontal layout adjusts the vertical align, and vice versa */
            var isHorizontalLayout = (alignRadios.getDirection() === "horizontal");
            var radioByKey = isHorizontalLayout ? VERTICAL_ALIGN_RADIO_BY_KEY : HORIZONTAL_ALIGN_RADIO_BY_KEY;
            var radioSet = isHorizontalLayout ? alignRadios.vertical : alignRadios.horizontal;

            var radioKey = radioByKey[event.keyName];
            if (!radioKey) return;

            for (var setKey in radioSet) {
                radioSet[setKey].value = (setKey === radioKey);
            }
            event.preventDefault();
            if (onUpdate) onUpdate();
        });
    }

    // =========================================
    // アウトライン計測 / Outline measurement
    // =========================================

    /**
     * コンテナ内のすべてのテキストをアウトライン化する
     * @param {PageItem} containerItem - テキスト、またはテキストを含むグループ
     * @returns {void}
     */
    function outlineAllTextInContainer(containerItem) {
        if (!containerItem) return;
        /* createOutline() は空のテキストなどで失敗する / createOutline() can fail, e.g. on empty text */
        try {
            if (containerItem.typename === "TextFrame") {
                containerItem.createOutline();
                return;
            }
            if (containerItem.textFrames && containerItem.textFrames.length) {
                /* createOutline で要素が消えるため後ろから走査 / iterate backwards: createOutline removes the frame */
                for (var i = containerItem.textFrames.length - 1; i >= 0; i--) {
                    try { containerItem.textFrames[i].createOutline(); } catch (e1) { }
                }
            }
        } catch (e) { }
    }

    /**
     * 複製をアウトライン化して境界を測る（boundsCache に控えて再利用）
     * @param {PageItem} originalItem - 測るオブジェクト
     * @param {Array} boundsCache - [{item: PageItem, bounds: number[]}] の控え
     * @returns {number[]|null} [左, 上, 右, 下]（テキストを含まない・測れないときは null）
     */
    function measureOutlineBoundsOnce(originalItem, boundsCache) {
        if (!originalItem) return null;

        for (var i = 0; i < boundsCache.length; i++) {
            if (boundsCache[i].item === originalItem) return boundsCache[i].bounds;
        }
        if (!containerHasText(originalItem)) return null;

        var doc = app.activeDocument;
        var measureGroup = null;
        try {
            var measureLayer = null;
            try { measureLayer = originalItem.layer; } catch (eLayer) { }
            if (!measureLayer) measureLayer = doc.activeLayer;
            measureGroup = measureLayer.groupItems.add();
            measureGroup.name = TEMP_MEASURE_GROUP_NAME;

            var duplicatedItem = null;
            try {
                duplicatedItem = originalItem.duplicate(measureGroup, ElementPlacement.PLACEATEND);
            } catch (eDup) {
                duplicatedItem = originalItem.duplicate();
                duplicatedItem.move(measureGroup, ElementPlacement.PLACEATEND);
            }

            outlineAllTextInContainer(duplicatedItem);

            var measuredBounds = null;
            try { measuredBounds = measureGroup.visibleBounds; }
            catch (eVisible) { try { measuredBounds = measureGroup.geometricBounds; } catch (eGeometric) { } }
            if (!measuredBounds) return null;

            boundsCache.push({ item: originalItem, bounds: [measuredBounds[0], measuredBounds[1], measuredBounds[2], measuredBounds[3]] });
            return [measuredBounds[0], measuredBounds[1], measuredBounds[2], measuredBounds[3]];
        } catch (e) {
            return null;
        } finally {
            try { if (measureGroup) measureGroup.remove(); } catch (eRemove) { }
        }
    }

    // =========================================
    // レイアウト適用 / Layout application
    // =========================================

    /**
     * 並べる順（並べ替え、またはシャッフル）と、ランダム時の基準位置を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {string} direction - "horizontal" / "vertical"
     * @param {boolean} useRandom - ランダムに並べるか
     * @param {number[]|null} previousShuffleOrder - 前回のシャッフル順（件数が同じなら使い回す）
     * @returns {{items: PageItem[], baseLeft: number, baseTop: number, shuffleOrder: number[]}} 並べる順と基準位置
     */
    function buildArrangementOrder(targetItems, direction, useRandom, previousShuffleOrder) {
        var arrangement = { items: [], baseLeft: null, baseTop: null, shuffleOrder: previousShuffleOrder };

        if (useRandom) {
            if (!arrangement.shuffleOrder || arrangement.shuffleOrder.length !== targetItems.length) {
                arrangement.shuffleOrder = makeShuffledIndices(targetItems.length);
            }
            for (var i = 0; i < arrangement.shuffleOrder.length; i++) {
                arrangement.items.push(targetItems[arrangement.shuffleOrder[i]]);
            }
            /* ランダム時は元の左上を基準として保持 / Preserve top-left base for random */
            for (var k = 0; k < targetItems.length; k++) {
                var targetItem = targetItems[k];
                if (!targetItem) continue;
                if (arrangement.baseLeft === null || targetItem.left < arrangement.baseLeft) arrangement.baseLeft = targetItem.left;
                if (arrangement.baseTop === null || targetItem.top > arrangement.baseTop) arrangement.baseTop = targetItem.top;
            }
            if (arrangement.baseLeft === null && arrangement.items[0]) arrangement.baseLeft = arrangement.items[0].left;
            if (arrangement.baseTop === null && arrangement.items[0]) arrangement.baseTop = arrangement.items[0].top;
        } else {
            arrangement.shuffleOrder = null;
            arrangement.items = (direction === "horizontal") ? sortLeftToRight(targetItems) : sortTopToBottom(targetItems);
        }
        return arrangement;
    }

    /**
     * オブジェクトの幅または高さの最大値を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {function} getBounds - 境界 [左, 上, 右, 下] を返す関数
     * @param {boolean} measureWidth - true なら幅、false なら高さ
     * @returns {number} 最大値
     */
    function computeMaxSize(targetItems, getBounds, measureWidth) {
        var maxSize = 0;
        for (var i = 0; i < targetItems.length; i++) {
            var bounds = getBounds(targetItems[i]);
            var itemSize = measureWidth ? (bounds[2] - bounds[0]) : (bounds[1] - bounds[3]);
            if (itemSize > maxSize) maxSize = itemSize;
        }
        return maxSize;
    }

    /**
     * 左→右に横並びで配置し、上下揃えを適用する
     * @param {PageItem[]} targetItems - 並べる順のオブジェクト
     * @param {number} startX - 先頭の左端
     * @param {number} startY - 揃えの基準にする上端
     * @param {number} referenceHeight - 揃えの基準にする高さ
     * @param {number} spacingPt - 間隔（pt）
     * @param {string} vAlignMode - "none" / "top" / "middle" / "bottom"
     * @param {function} getBounds - 境界を返す関数
     * @returns {void}
     */
    function applyHorizontalLayout(targetItems, startX, startY, referenceHeight, spacingPt, vAlignMode, getBounds) {
        var currentX = startX;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getBounds(targetItem);
            var itemWidth = bounds[2] - bounds[0];

            /* Xは左基準で配置 / Place by left edge on X */
            targetItem.left = targetItem.left + (currentX - bounds[0]);

            /* 上下揃え / Vertical alignment */
            if (vAlignMode !== "none") {
                var cellTop = startY;
                var cellBottom = startY - referenceHeight;
                var dy = 0;
                if (vAlignMode === "middle") dy = ((cellTop + cellBottom) / 2) - ((bounds[1] + bounds[3]) / 2);
                else if (vAlignMode === "bottom") dy = cellBottom - bounds[3];
                else dy = cellTop - bounds[1]; /* top */
                targetItem.top = targetItem.top + dy;
            }

            if (i < targetItems.length - 1) {
                currentX += itemWidth + spacingPt;
            }
        }
    }

    /**
     * 上→下に縦並びで配置し、左右揃えを適用する
     * @param {PageItem[]} targetItems - 並べる順のオブジェクト
     * @param {number} startX - 揃えの基準にする左端
     * @param {number} startY - 先頭の上端
     * @param {number} referenceWidth - 揃えの基準にする幅
     * @param {number} spacingPt - 間隔（pt）
     * @param {string} hAlignMode - "none" / "left" / "center" / "right"
     * @param {function} getBounds - 境界を返す関数
     * @returns {void}
     */
    function applyVerticalLayout(targetItems, startX, startY, referenceWidth, spacingPt, hAlignMode, getBounds) {
        var currentY = startY;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getBounds(targetItem);
            var itemHeight = bounds[1] - bounds[3];

            /* 左右揃え / Horizontal alignment */
            if (hAlignMode !== "none") {
                var cellLeft = startX;
                var cellRight = cellLeft + referenceWidth;
                var dx = 0;
                if (hAlignMode === "center") dx = ((cellLeft + cellRight) / 2) - ((bounds[0] + bounds[2]) / 2);
                else if (hAlignMode === "right") dx = cellRight - bounds[2];
                else dx = cellLeft - bounds[0]; /* left */
                targetItem.left = targetItem.left + dx;
            }

            /* YはTop揃えで積む / Stack by top on Y */
            targetItem.top = targetItem.top + (currentY - bounds[1]);

            if (i < targetItems.length - 1) {
                currentY -= itemHeight + spacingPt;
            }
        }
    }

    /**
     * ランダム時に、先頭のオブジェクトが元の左上に来るよう全体をずらす
     * @param {PageItem[]} targetItems - 並べたオブジェクト
     * @param {number} baseLeft - 元の左端
     * @param {number} baseTop - 元の上端
     * @returns {void}
     */
    function applyRandomBaseOffset(targetItems, baseLeft, baseTop) {
        if (!targetItems.length) return;
        var dx = baseLeft - targetItems[0].left;
        var dy = baseTop - targetItems[0].top;
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            targetItems[i].left += dx;
            targetItems[i].top += dy;
        }
    }

    /**
     * 並べる順に従ってオブジェクトを配置する
     * @param {{items: PageItem[], baseLeft: number, baseTop: number}} arrangement - buildArrangementOrder() の結果
     * @param {{direction: string, spacingPt: number, useRandom: boolean, hAlignMode: string, vAlignMode: string}} layoutSettings - 並べ方
     * @param {function} getBounds - 境界 [左, 上, 右, 下] を返す関数
     * @returns {void}
     */
    function placeArrangement(arrangement, layoutSettings, getBounds) {
        var arrangedItems = arrangement.items;
        if (!arrangedItems.length) return;

        var startBounds = getBounds(arrangedItems[0]);
        var startX = startBounds[0];
        var startY = startBounds[1];

        if (layoutSettings.direction === "horizontal") {
            var referenceHeight = computeMaxSize(arrangedItems, getBounds, false);
            applyHorizontalLayout(arrangedItems, startX, startY, referenceHeight, layoutSettings.spacingPt, layoutSettings.vAlignMode, getBounds);
        } else {
            var referenceWidth = computeMaxSize(arrangedItems, getBounds, true);
            applyVerticalLayout(arrangedItems, startX, startY, referenceWidth, layoutSettings.spacingPt, layoutSettings.hAlignMode, getBounds);
        }

        if (layoutSettings.useRandom) {
            applyRandomBaseOffset(arrangedItems, arrangement.baseLeft, arrangement.baseTop);
        }
    }

    // =========================================
    // ダイアログ部品 / Dialog parts
    // =========================================

    /**
     * ［方向］パネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{auto: RadioButton, vertical: RadioButton, horizontal: RadioButton}} 方向のラジオ
     */
    function addDirectionPanel(parentDialog) {
        var directionPanel = parentDialog.add("panel", undefined, getLabel("panel.direction"));
        directionPanel.orientation = "row";
        directionPanel.alignChildren = ["center", "center"];
        directionPanel.margins = PANEL_MARGINS;
        var directionRadios = {
            auto: addTooltipRadio(directionPanel, getLabel("radio.directionAuto"), getLabel("tooltip.directionAuto")),
            vertical: addTooltipRadio(directionPanel, getLabel("radio.directionVertical"), getLabel("tooltip.directionVertical")),
            horizontal: addTooltipRadio(directionPanel, getLabel("radio.directionHorizontal"), getLabel("tooltip.directionHorizontal"))
        };
        directionRadios.vertical.value = true; /* デフォルトを「縦」に / Default to Vertical */
        return directionRadios;
    }

    /**
     * ツールチップ付きのラジオボタンを追加する
     * @param {object} parentGroup - 追加先のパネルまたはグループ
     * @param {string} radioText - ラジオの表示名
     * @param {string} tooltipText - ツールチップ
     * @returns {RadioButton} 追加したラジオ
     */
    function addTooltipRadio(parentGroup, radioText, tooltipText) {
        var radioButton = parentGroup.add("radiobutton", undefined, radioText);
        radioButton.helpTip = tooltipText;
        return radioButton;
    }

    /**
     * ［間隔］パネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @param {function} onSpacingStep - ∧∨・↑↓キーで値を変えたあとに呼ぶ処理
     * @returns {EditText} 間隔の入力欄
     */
    function addSpacingPanel(parentDialog, onSpacingStep) {
        var spacingPanel = parentDialog.add("panel", undefined, getLabel("panel.spacing"));
        spacingPanel.orientation = "column";
        spacingPanel.alignChildren = ["center", "center"];
        spacingPanel.margins = PANEL_MARGINS;
        var spacingRowGroup = spacingPanel.add("group");
        spacingRowGroup.orientation = "row";
        spacingRowGroup.alignChildren = ["left", "center"];
        /* ∧∨と入力欄は隙間0で突き合わせる。マイナス値も許す（重ねる）/ Stepper butts the field; negatives allowed (overlap) */
        var spacingFieldGroup = spacingRowGroup.add("group");
        spacingFieldGroup.orientation = "row";
        spacingFieldGroup.alignChildren = ["left", "center"];
        spacingFieldGroup.spacing = 0;
        spacingFieldGroup.margins = 0;
        var spacingInput;
        var spacingStepper = addStepper(spacingFieldGroup, function () { return spacingInput; }, {
            onStep: function () { onSpacingStep(); }
        });
        spacingInput = spacingFieldGroup.add("edittext", undefined, "0");
        spacingInput.characters = SPACING_FIELD_CHARS;
        spacingInput.helpTip = getLabel("tooltip.spacing");
        bindSteppedArrowKeys(spacingInput, spacingStepper); /* ↑↓キーも∧∨と同じ処理 / arrow keys share the stepper */
        spacingRowGroup.add("statictext", undefined, getUnitInfo().label);
        return spacingInput;
    }

    /**
     * 揃えのラジオを1行ぶん追加する
     * @param {Panel} alignPanel - 追加先のパネル
     * @param {Array} radioDefs - [ラジオのキー, LABELS のパス] の配列
     * @param {string} defaultKey - 最初に選ぶラジオのキー
     * @param {string} tooltipText - 行のラジオすべてに付けるツールチップ
     * @returns {{rowGroup: Group, radios: Object}} 行のグループと、キー → ラジオ
     */
    function addAlignRadioRow(alignPanel, radioDefs, defaultKey, tooltipText) {
        var rowGroup = alignPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        var alignRadios = {};
        for (var i = 0; i < radioDefs.length; i++) {
            alignRadios[radioDefs[i][0]] = rowGroup.add("radiobutton", undefined, getLabel(radioDefs[i][1]));
        }
        alignRadios[defaultKey].value = true;
        for (var radioKey in alignRadios) {
            alignRadios[radioKey].helpTip = tooltipText;
        }
        return { rowGroup: rowGroup, radios: alignRadios };
    }

    /**
     * ［揃え］パネル（左右の揃えと上下の揃えの2行）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{panel: Panel, horizontalRow: Group, verticalRow: Group, horizontal: Object, vertical: Object}} パネルと各行のラジオ
     */
    function addAlignPanel(parentDialog) {
        var alignPanel = parentDialog.add("panel", undefined, "");
        alignPanel.orientation = "column";
        alignPanel.alignChildren = ["left", "center"];
        alignPanel.margins = PANEL_MARGINS;

        /* 既定は「中央」/ Default to Center (Middle) */
        var horizontalRow = addAlignRadioRow(alignPanel, [
            ["none", "radio.alignNone"], ["left", "radio.alignLeft"], ["center", "radio.alignCenter"], ["right", "radio.alignRight"]
        ], "center", getLabel("tooltip.alignHorizontal"));
        var verticalRow = addAlignRadioRow(alignPanel, [
            ["none", "radio.alignNone"], ["top", "radio.alignTop"], ["middle", "radio.alignMiddle"], ["bottom", "radio.alignBottom"]
        ], "middle", getLabel("tooltip.alignVertical"));

        return {
            panel: alignPanel,
            horizontalRow: horizontalRow.rowGroup,
            verticalRow: verticalRow.rowGroup,
            horizontal: horizontalRow.radios,
            vertical: verticalRow.radios
        };
    }

    /**
     * オプションのチェックボックス（プレビュー境界・テキスト計測・ランダム）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{usePreviewBounds: Checkbox, measureText: Checkbox, random: Checkbox}} チェックボックス
     */
    function addOptionsGroup(parentDialog) {
        var optionsGroup = parentDialog.add("group");
        optionsGroup.orientation = "column";
        optionsGroup.alignChildren = ["left", "center"];
        optionsGroup.alignment = ["fill", "top"];
        optionsGroup.margins = OPTIONS_MARGINS;

        return {
            usePreviewBounds: addOptionCheckbox(optionsGroup, getLabel("checkbox.usePreviewBounds"), getLabel("tooltip.usePreviewBounds"), true),
            measureText: addOptionCheckbox(optionsGroup, getLabel("checkbox.measureText"), getLabel("tooltip.measureText"), false),
            random: addOptionCheckbox(optionsGroup, getLabel("checkbox.random"), getLabel("tooltip.random"), false)
        };
    }

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} checkboxText - 表示名
     * @param {string} tooltipText - ツールチップ
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentGroup, checkboxText, tooltipText, initialValue) {
        var optionCheckbox = parentGroup.add("checkbox", undefined, checkboxText);
        optionCheckbox.value = initialValue;
        optionCheckbox.helpTip = tooltipText;
        return optionCheckbox;
    }

    /**
     * ［キャンセル］［OK］のボタン行を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addButtonRow(parentDialog) {
        var btnRowGroup = parentDialog.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    }

    /**
     * ラジオの組から選ばれているもののキーを返す
     * @param {object} radioSet - キー → ラジオ
     * @param {string} fallbackKey - どれも選ばれていないときのキー
     * @returns {string} 選ばれているラジオのキー
     */
    function getCheckedKey(radioSet, fallbackKey) {
        for (var radioKey in radioSet) {
            if (radioSet[radioKey].value) return radioKey;
        }
        return fallbackKey;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを表示し、プレビューしながら整列・分布を実行する
     * @returns {boolean} OK で確定したら true、キャンセルなら false
     */
    function showArrangeDialog() {
        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = "fill";
        alignDialog.opacity = DIALOG_OPACITY;

        var lastPosition = loadDialogPosition();
        if (lastPosition) alignDialog.location = lastPosition;

        /* キャンセル時に復元するため現在のプリファレンスを保存 / Preserve preference for cancel */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

        /* 選択スナップショット / Snapshot selection */
        var originalSelection = app.activeDocument.selection.slice();
        var previewState = { isUndo: false };
        var detectedDirection = detectDirection(originalSelection);
        var selectionHasText = anyItemHasText(originalSelection);

        var directionRadios = addDirectionPanel(alignDialog);
        var spacingInput = addSpacingPanel(alignDialog, requestPreviewUpdate);
        var alignControls = addAlignPanel(alignDialog);
        var optionCheckboxes = addOptionsGroup(alignDialog);
        addButtonRow(alignDialog);

        /* ---- 内部状態 / Internal state ---- */
        var randomOrderCache = null;
        var outlineMeasureCache = []; /* [{item:PageItem, bounds:[l,t,r,b]}] */
        var isPreviewUpdating = false;
        var lastPreviewTime = 0;
        var previewTaskId = 0;

        /**
         * 選択中のラジオから実際に並べる方向を返す（自動は判定結果）
         * @returns {string} "horizontal" または "vertical"
         */
        function getEffectiveDirection() {
            if (directionRadios.vertical.value) return "vertical";
            if (directionRadios.horizontal.value) return "horizontal";
            return detectedDirection;
        }

        /**
         * レイアウト用の境界を返す（縦並びでテキスト計測が ON ならアウトラインの実測値）
         * @param {PageItem} targetItem - 対象のオブジェクト
         * @returns {number[]} [左, 上, 右, 下]
         */
        function getBoundsForLayout(targetItem) {
            if (optionCheckboxes.measureText.value && getEffectiveDirection() !== "horizontal") {
                var outlinedBounds = measureOutlineBoundsOnce(targetItem, outlineMeasureCache);
                if (outlinedBounds) return outlinedBounds;
            }
            return getItemBounds(targetItem, optionCheckboxes.usePreviewBounds.value);
        }

        /**
         * 間隔の入力値を現在の単位から pt に換算する
         * @returns {number} 間隔（pt）
         */
        function getSpacingPt() {
            var spacingValue = parseFloat(spacingInput.text);
            if (isNaN(spacingValue)) spacingValue = 0;
            return spacingValue * getUnitInfo().pointsPerUnit;
        }

        /**
         * 現在の設定で選択を並べる
         * @returns {void}
         */
        function applyLayoutToSelection() {
            if (!originalSelection || originalSelection.length === 0) return;
            var direction = getEffectiveDirection();
            var spacingPt = getSpacingPt();
            var useRandom = optionCheckboxes.random.value;
            var arrangement = buildArrangementOrder(originalSelection, direction, useRandom, randomOrderCache);
            randomOrderCache = arrangement.shuffleOrder;
            placeArrangement(arrangement, {
                direction: direction,
                spacingPt: spacingPt,
                useRandom: useRandom,
                hAlignMode: getCheckedKey(alignControls.horizontal, "left"),
                vAlignMode: getCheckedKey(alignControls.vertical, "top")
            }, getBoundsForLayout);
        }

        /**
         * 方向に合わせて揃えパネルの見出しと有効状態、テキスト計測の有効状態を切り替える
         * @returns {void}
         */
        function syncAlignUI() {
            var isHorizontal = (getEffectiveDirection() === "horizontal");
            alignControls.panel.text = getLabel(isHorizontal ? "panel.alignVertical" : "panel.alignHorizontal");
            alignControls.horizontalRow.enabled = !isHorizontal;
            alignControls.verticalRow.enabled = isHorizontal;

            /* テキスト計測は縦並び時かつ選択にテキストがある場合のみ有効 / Enable text-measure only when vertical & selection has text */
            optionCheckboxes.measureText.enabled = !isHorizontal && selectionHasText;

            /* 表示前のレイアウトは失敗することがある / layout before show may fail */
            try { alignDialog.layout.layout(true); } catch (e) { }
        }

        /**
         * 連続呼び出しを間引きながらプレビューを作り直す
         * @returns {void}
         */
        function updatePreview() {
            if (isPreviewUpdating) return;
            var now = new Date().getTime();
            if (now - lastPreviewTime < PREVIEW_MIN_INTERVAL_MS) return;
            lastPreviewTime = now;

            isPreviewUpdating = true;
            try {
                try {
                    app.preferences.setBooleanPreference("includeStrokeInBounds", optionCheckboxes.usePreviewBounds.value);
                } catch (e) { }
                runPreview(previewState, applyLayoutToSelection, true);
            } finally {
                isPreviewUpdating = false;
            }
        }

        /**
         * プレビューを予約する（ScriptUI のイベント中に DOM を触ると落ちるため scheduleTask で遅らせる）
         * @returns {void}
         */
        function requestPreviewUpdate() {
            try { if (previewTaskId) app.cancelTask(previewTaskId); } catch (e) { }
            $.global.__SAT_updatePreview = updatePreview;
            try {
                previewTaskId = app.scheduleTask('$.global.__SAT_updatePreview && $.global.__SAT_updatePreview();', PREVIEW_SCHEDULE_MS, false);
            } catch (e) {
                updatePreview();
            }
        }

        /**
         * 方向を変えたら揃えパネルを同期してプレビューを更新する
         * @returns {void}
         */
        function onDirectionChanged() {
            syncAlignUI();
            requestPreviewUpdate();
        }

        /* ---- イベントバインド / Event bindings ---- */
        var alignRadios = {
            horizontal: alignControls.horizontal,
            vertical: alignControls.vertical,
            getDirection: getEffectiveDirection
        };
        addAlignKeyHandler(alignDialog, alignRadios, requestPreviewUpdate);
        addAlignKeyHandler(spacingInput, alignRadios, requestPreviewUpdate);

        var radioKey;
        for (radioKey in alignControls.horizontal) alignControls.horizontal[radioKey].onClick = requestPreviewUpdate;
        for (radioKey in alignControls.vertical) alignControls.vertical[radioKey].onClick = requestPreviewUpdate;

        directionRadios.auto.onClick = onDirectionChanged;
        directionRadios.vertical.onClick = onDirectionChanged;
        directionRadios.horizontal.onClick = onDirectionChanged;

        optionCheckboxes.usePreviewBounds.onClick = function () {
            /* 境界モードが変わるとアウトライン計測結果も再計算 / Outline cache invalid when bounds mode changes */
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };
        optionCheckboxes.random.onClick = function () {
            randomOrderCache = null;
            requestPreviewUpdate();
        };
        optionCheckboxes.measureText.onClick = function () {
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };

        /* 初期化 / Init */
        syncAlignUI();
        requestPreviewUpdate();
        spacingInput.active = true;

        var dialogResult = alignDialog.show();
        try { if (previewTaskId) app.cancelTask(previewTaskId); } catch (e) { }
        saveDialogPosition(alignDialog.location);

        if (dialogResult !== 1) {
            /* キャンセル: プレビューを undo してプリファレンスも戻す / Cancel: undo preview & restore preference */
            undoPreview(previewState);
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return false;
        }

        /* 確定: プレビュー分を巻き戻して本実行（undo 履歴を 1 件にまとめる） / OK: undo preview, then commit as single history entry */
        undoPreview(previewState);
        try { app.preferences.setBooleanPreference("includeStrokeInBounds", optionCheckboxes.usePreviewBounds.value); } catch (e) { }
        applyLayoutToSelection();
        app.redraw();
        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        /* ドキュメントが無いと app.activeDocument が例外になる / activeDocument throws with no document */
        try {
            var selectedItems = app.activeDocument.selection;
            if (!selectedItems || selectedItems.length === 0) {
                alert(getLabel("alert.noSelection"));
                return;
            }
            showArrangeDialog();
        } catch (e) {
            alert(getLabel("alert.errorPrefix") + e.message);
        }
    }

    main();

})();
