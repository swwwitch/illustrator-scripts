#target illustrator
#targetengine "TextWithShapeToAreaTypeEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextWithShapeToAreaType.md

### Overview

Converts point text, text on a path, or a shape plus text into area text while preserving the appearance.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextWithShapeToAreaType.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextWithShapeToAreaType";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextWithShapeToAreaType.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextWithShapeToAreaType.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ［大きさ調整：する］の既定の倍率 / Default ratios for size adjustment "On" */
    var BUTTON_WIDTH_RATIO  = 1.2;  /* 元の幅に対する倍率 / Ratio of original width */
    var BUTTON_HEIGHT_RATIO = 1.6;  /* 元の高さに対する倍率 / Ratio of original height */

    // =========================================
    // 設定ファイル / Preferences file
    // =========================================

    /* 参照した AI ファイルとスタイル名を記憶する設定ファイル / Prefs file remembering the picked AI file and its style names */
    var PREFS_FILE_NAME = "styles_for_TextWithShapeToAreaType.txt";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS         = 15;                /* ダイアログの余白 / Dialog margins */
    var PANEL_MARGINS          = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] / Panel margins */
    var PANEL_SPACING          = 8;                 /* パネル内の要素間隔 / Panel spacing */
    var RATIO_LABEL_WIDTH      = 44;                /* 幅・高さの項目名の幅 / Width of the ratio labels */
    var RATIO_INPUT_CHARACTERS = 5;                 /* 倍率欄の幅（文字数）/ Width of the ratio fields */
    var FILE_NAME_WIDTH        = 240;               /* ファイル名表示の幅 / Width of the file-name label */

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

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

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義（UI パーツ別）/ Japanese-English label definitions (by UI part) */
    var LABELS = {
        dialog: {
            title: { ja: "エリア内文字に変換", en: "Convert to Area Type" },
            pickFile: { ja: "スタイルの AI ファイルを選択", en: "Select a style AI file" }
        },
        panel: {
            sizeAdjust: { ja: "大きさ調整", en: "Size adjustment" },
            style: { ja: "グラフィックスタイル", en: "Graphic style" },
            loadStyles: { ja: "スタイルの読み込み", en: "Load Styles" }
        },
        radio: {
            doAdjust: { ja: "する", en: "On" },
            dontAdjust: { ja: "しない", en: "Off" },
            styleOriginal: { ja: "元の見た目", en: "Original appearance" }
        },
        fieldLabel: {
            widthRatio: { ja: "幅", en: "Width" },
            heightRatio: { ja: "高さ", en: "Height" }
        },
        status: {
            noFileSelected: { ja: "ファイル未選択", en: "No file selected" }
        },
        button: {
            load: { ja: "読み込み", en: "Load" },
            reload: { ja: "再読み込み", en: "Reload" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            load: {
                ja: "「読み込み」でスタイルの AI ファイルを選択してください。",
                en: "Click “Load” to choose a style AI file."
            },
            reload: {
                ja: "記憶したファイルからスタイルを取り込み直します。",
                en: "Re-import styles from the remembered file."
            },
            doAdjust: {
                ja: "テキストだけを変換するとき、枠を幅・高さの倍率で広げます。「元の見た目」ではボタン風の背景（塗り2つと長方形の効果）も付けます。図形と一緒に選んだときは使われません。",
                en: "When converting text alone, enlarges the frame by the width and height ratios. With Original appearance, also adds a button-like background (two fills and a rectangle effect). Not used when a shape is selected with the text."
            },
            ratio: {
                ja: "実測した大きさに対する百分率。数値以外や0以下のときは既定値を使います。",
                en: "Percentage of the measured size. Falls back to the default when the value is not a positive number."
            },
            styleOriginal: {
                ja: "元のテキストの見た目（塗り・線・効果）を一時的なグラフィックスタイルにして引き継ぎます。",
                en: "Carries over the source text's appearance (fill, stroke, effects) through a temporary graphic style."
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
            selectText: {
                ja: "ポイント文字・パス上文字、またはテキストと図形を選択してください。",
                en: "Please select point text, path text, or text and a shape."
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            rectangleOnly: {
                ja: "フレームに使える図形は長方形のみです。長方形を選択してください。",
                en: "Only a rectangle can serve as the frame. Please select a rectangle."
            },
            fileNotFound: {
                ja: "指定されたファイルが見つかりません：\n",
                en: "The specified file was not found:\n"
            },
            styleNotFound: {
                ja: "指定したグラフィックスタイルが見つかりません：\n",
                en: "The graphic style was not found:\n"
            }
        }
    };

    /**
     * ドット区切りのパス、または { ja, en } のラベルから表示言語の文字列を引く
     * @param {string|Object} labelPath - "alert.noDocument" のようなドット区切りのキーか { ja: string, en: string }
     * @returns {string} 表示言語の文字列（無ければ英語、それも無ければ labelPath）
     */
    function getLabel(labelPath) {
        /* ステップボタンは LABELS.tooltip.stepUp などをオブジェクトで渡す / the stepper passes label objects */
        if (typeof labelPath !== "string") return labelPath[uiLang] || labelPath.en;
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            if (!labelNode || typeof labelNode[pathKeys[i]] === "undefined") return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        if (labelNode[uiLang]) return labelNode[uiLang];
        if (labelNode.en) return labelNode.en;
        return labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // テキストの種類 / Text kinds
    // =========================================

    /**
     * ポイント文字またはパス上文字かどうか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} ポイント文字・パス上文字なら true
     */
    function isPointOrPathText(pageItem) {
        return pageItem.typename === "TextFrame" && (pageItem.kind === TextType.POINTTEXT || pageItem.kind === TextType.PATHTEXT);
    }

    /**
     * 選択に変換できるテキスト（ポイント文字・パス上文字）が含まれるか
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {boolean} 含まれれば true
     */
    function containsConvertibleText(currentSelection) {
        for (var i = 0; i < currentSelection.length; i++) {
            if (isPointOrPathText(currentSelection[i])) return true;
        }
        return false;
    }

    // =========================================
    // パス上文字 → ポイント文字（変換前処理）/ Path text → Point text (pre-process)
    // =========================================

    /**
     * 関数を実行し、例外は握りつぶす
     * @param {Function} domAction - 実行する関数（DOM の読み書き）
     * @returns {*} 関数の戻り値。失敗したら undefined
     */
    function tryQuietly(domAction) { try { return domAction(); } catch (e) { return undefined; } }

    /**
     * パス上文字の文字ごとの属性を退避する
     * @param {TextFrame} textFrame - パス上文字
     * @returns {Object[]} 文字ごとの属性（font / size / fillColor / strokeColor / strokeWeight / autoLeading / leading）
     */
    function snapshotCharAttrs(textFrame) {
        var charAttrs = [];
        for (var i = 0; i < textFrame.characters.length; i++) {
            var charAttr = textFrame.characters[i].characterAttributes;
            charAttrs.push({
                font: charAttr.textFont,
                size: charAttr.size,
                fillColor: charAttr.fillColor,
                strokeColor: charAttr.strokeColor,
                strokeWeight: charAttr.strokeWeight,
                autoLeading: charAttr.autoLeading,
                leading: charAttr.leading
            });
        }
        return charAttrs;
    }

    /**
     * 退避した文字属性を新しいポイント文字へ戻す（ベースラインシフト・比率は既定に戻す）
     * @param {TextFrame} pointText - 新しいポイント文字
     * @param {Object[]} charAttrs - snapshotCharAttrs() の結果
     * @returns {void}
     */
    function restoreCharAttrs(pointText, charAttrs) {
        var restoreCount = Math.min(pointText.characters.length, charAttrs.length);
        for (var i = 0; i < restoreCount; i++) {
            var targetAttr = pointText.characters[i].characterAttributes;
            var sourceAttr = charAttrs[i];

            tryQuietly(function () { targetAttr.textFont = sourceAttr.font; });
            tryQuietly(function () { targetAttr.size = sourceAttr.size; });
            tryQuietly(function () { targetAttr.fillColor = sourceAttr.fillColor; });
            tryQuietly(function () {
                var sourceStroke = sourceAttr.strokeColor;
                targetAttr.strokeColor = sourceStroke;
                targetAttr.strokeWeight = (sourceStroke && sourceStroke.typename === "NoColor") ? 0 : sourceAttr.strokeWeight;
            });
            tryQuietly(function () { targetAttr.baselineShift = 0; });
            tryQuietly(function () { targetAttr.horizontalScale = 100; });
            tryQuietly(function () { targetAttr.verticalScale = 100; });
            tryQuietly(function () { targetAttr.autoLeading = sourceAttr.autoLeading; });
            if (!sourceAttr.autoLeading) tryQuietly(function () { targetAttr.leading = sourceAttr.leading; });
        }
    }

    /**
     * パス上文字と同じ内容・行揃え・文字属性のポイント文字を、パスの始点に作る
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame} pathText - 元のパス上文字
     * @returns {TextFrame|null} 作ったポイント文字。パスを取れなければ null
     */
    function createPointTextFromPathText(doc, pathText) {
        var originalPath = null;
        tryQuietly(function () { originalPath = pathText.textPath; });
        if (!originalPath) return null;

        /* 1) 文字ごとの属性・内容・行揃えを退避 / Snapshot per-character attributes, contents, justification */
        var charAttrs = snapshotCharAttrs(pathText);

        var textContents = "";
        tryQuietly(function () { textContents = pathText.contents; });

        var justification = null;
        tryQuietly(function () {
            if (pathText.paragraphs && pathText.paragraphs.length > 0) {
                justification = pathText.paragraphs[0].paragraphAttributes.justification;
            }
        });

        /* 2) パス始点にポイント文字を新規作成 / Create new point text at path start anchor */
        var pointText = doc.textFrames.add();
        var anchorPoint = null;
        tryQuietly(function () {
            if (originalPath.pathPoints && originalPath.pathPoints.length > 0) {
                anchorPoint = originalPath.pathPoints[0].anchor;
            }
        });
        if (anchorPoint) {
            pointText.position = [anchorPoint[0], anchorPoint[1]];
        }

        pointText.contents = textContents;

        if (justification !== null && pointText.paragraphs && pointText.paragraphs.length > 0) {
            tryQuietly(function () { pointText.paragraphs[0].paragraphAttributes.justification = justification; });
        }

        /* 既定の線を一旦消し、後で文字ごとに復元 / Clear default stroke, restore per-character later */
        tryQuietly(function () {
            var noColor = new NoColor();
            pointText.textRange.characterAttributes.strokeColor = noColor;
            pointText.textRange.characterAttributes.strokeWeight = 0;
        });

        /* 文字ごとの属性を復元 / Restore per-character attributes */
        restoreCharAttrs(pointText, charAttrs);
        return pointText;
    }

    /**
     * パス上文字を字形を保ったままポイント文字へ分離する（元のパス上文字は削除）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} pathTextFrames - パス上文字
     * @returns {TextFrame[]} 作ったポイント文字（選択状態）
     */
    function detachPathTextToPointText(doc, pathTextFrames) {
        var createdTexts = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return createdTexts;

        /* 新規テキストだけ選べるよう選択を解除 / Clear selection */
        tryQuietly(function () { doc.selection = null; });

        for (var i = pathTextFrames.length - 1; i >= 0; i--) {
            var pathText = pathTextFrames[i];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;

            var pointText = createPointTextFromPathText(doc, pathText);
            if (!pointText) continue;

            /* 3) 元のパス上文字を削除（パスも一緒に消える）/ Remove original path text */
            tryQuietly(function () { pathText.remove(); });

            /* 4) 新規テキストを選択して返す / Select and return new text */
            tryQuietly(function () { pointText.selected = true; });
            createdTexts.push(pointText);
        }

        return createdTexts;
    }

    /**
     * 選択内のパス上文字をポイント文字へ置き換え、選択し直す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {PageItem[]} 置き換えた選択（パス上文字が無ければ元の選択）
     */
    function preprocessPathTextSelection(doc, currentSelection) {
        if (!doc || !currentSelection || !currentSelection.length) return currentSelection;

        var pathTexts = [];
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            try {
                if (selectedItem && selectedItem.typename === "TextFrame" && selectedItem.kind === TextType.PATHTEXT) {
                    pathTexts.push(selectedItem);
                }
            } catch (e0) {
                /* 無効オブジェクト（削除済み等）はスキップ / Skip invalid objects */
            }
        }
        if (!pathTexts.length) return currentSelection;

        var newTexts = detachPathTextToPointText(doc, pathTexts);
        if (!newTexts.length) return currentSelection;

        /* パス上文字を新ポイント文字に差し替えた新しい選択配列を構築。削除した旧オブジェクトは参照で例外になるので飛ばす
           Build the replaced selection; the removed originals throw on access and are skipped */
        var replacedSelection = [];
        for (var j = 0; j < currentSelection.length; j++) {
            var keptItem = currentSelection[j];
            try {
                if (keptItem && !(keptItem.typename === "TextFrame" && keptItem.kind === TextType.PATHTEXT)) {
                    replacedSelection.push(keptItem);
                }
            } catch (e1) {
                /* 無効オブジェクトはスキップ / Skip invalid objects */
            }
        }
        for (var k = 0; k < newTexts.length; k++) replacedSelection.push(newTexts[k]);

        try { doc.selection = replacedSelection; } catch (e) { }
        app.redraw();

        return replacedSelection;
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    // =========================================

    /**
     * 文字列を16進数表現にする
     * @param {string} text - ASCII 文字列
     * @returns {string} 16進数の文字列
     */
    function _hexAscii(text) {
        var hex = "";
        for (var i = 0; i < text.length; i++) {
            var pair = text.charCodeAt(i).toString(16);
            if (pair.length < 2) pair = "0" + pair;
            hex += pair;
        }
        return hex;
    }

    /**
     * アクション定義の /name ブロック（ASCII）を作る
     * @param {string} text - 名前
     * @returns {string} /name ブロック
     */
    function _nameBlockAscii(text) {
        return "/name [ " + text.length + " " + _hexAscii(text).toUpperCase() + " ]";
    }

    /**
     * アクションセット定義（.aia）の文字列を組み立てる
     * @param {string} setName - アクションセット名
     * @param {string} internalName - イベントの内部名
     * @param {string} localizedNameHex - ローカライズ名（"長さ 16進" 形式。空なら省略）
     * @param {number} paramKeyInt - パラメーターのキー
     * @param {Object[]} actionDefs - { name, value } のアクション定義
     * @returns {string} .aia の文字列
     */
    function _buildActionSetAIA(setName, internalName, localizedNameHex, paramKeyInt, actionDefs) {
        var aiaString = "/version 3" +
            _nameBlockAscii(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefs.length;

        for (var i = 0; i < actionDefs.length; i++) {
            var actionDef = actionDefs[i];
            aiaString += "/action-" + (i + 1) + " {" +
                " " + _nameBlockAscii(actionDef.name) +
                " /keyIndex 0" +
                " /colorIndex 0" +
                " /isOpen 1" +
                " /eventCount 1" +
                " /event-1 {" +
                " /useRulersIn1stQuadrant 0" +
                " /internalName (" + internalName + ")" +
                (localizedNameHex ? (" /localizedName [ " + localizedNameHex + " ]") : "") +
                " /isOpen 0" +
                " /isOn 1" +
                " /hasDialog 0" +
                " /parameterCount 1" +
                " /parameter-1 {" +
                " /key " + paramKeyInt +
                " /showInPalette 4294967295" +
                " /type (integer)" +
                " /value " + actionDef.value +
                " }" +
                " }" +
                "}";
        }
        return aiaString;
    }

    /* フレーム整列アクションセット名 / Frame-alignment action set name */
    var AREA_TEXT_ACTION_SET = "AreaText";

    /**
     * アクションセットを読み込む（一時フォルダーに .aia を書き出して loadAction）
     * 同名のセットがあれば先に外して衝突を避ける
     * @param {string} setName - アクションセット名
     * @param {string} aiaString - .aia の文字列
     * @returns {void}
     */
    function loadActionSet(setName, aiaString) {
        unloadActionSet(setName);
        try {
            var actionFile = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
            actionFile.open("w");
            actionFile.write(aiaString);
            actionFile.close();

            app.loadAction(actionFile);
            actionFile.remove();
        } catch (e) { }
    }

    /**
     * アクションセットを破棄する
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadActionSet(setName) {
        try { app.unloadAction(setName, ""); } catch (e) { }
    }

    /**
     * フレーム整列アクション（AlignTop / Center / Bottom / Justify）を読み込む
     * @returns {void}
     */
    function loadAreaTextActions() {
        var aiaString = _buildActionSetAIA(
            AREA_TEXT_ACTION_SET,
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            [
                { name: "AlignTop", value: 0 },
                { name: "AlignCenter", value: 1 },
                { name: "AlignBottom", value: 2 },
                { name: "AlignJustify", value: 3 }
            ]
        );
        loadActionSet(AREA_TEXT_ACTION_SET, aiaString);
    }

    /**
     * フレーム整列アクションを破棄する
     * @returns {void}
     */
    function unloadAreaTextActions() {
        unloadActionSet(AREA_TEXT_ACTION_SET);
    }

    /**
     * エリア内文字のフレーム整列（縦方向の配置）アクションを実行する
     * @param {number} valueInt - 0=上 / 1=中央 / 2=下 / 3=均等
     * @returns {void}
     */
    function runAreaTextFrameAlignmentAction(valueInt) {
        if (valueInt !== 0 && valueInt !== 1 && valueInt !== 2 && valueInt !== 3) return;

        var actionName = "AlignTop";
        if (valueInt === 1) actionName = "AlignCenter";
        else if (valueInt === 2) actionName = "AlignBottom";
        else if (valueInt === 3) actionName = "AlignJustify";

        try { app.doScript(actionName, AREA_TEXT_ACTION_SET, false); } catch (e) { }
    }

    /**
     * 指定のエリア内文字だけを選択して、フレーム整列（縦方向の配置）を適用する
     * @param {TextFrame} areaType - 対象のエリア内文字
     * @param {number} valueInt - 0=上 / 1=中央 / 2=下 / 3=均等
     * @returns {void}
     */
    function applyAreaTextFrameAlignment(areaType, valueInt) {
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaType];
            app.redraw(); /* 選択状態を確定 / Commit the selection */
            runAreaTextFrameAlignmentAction(valueInt);
        } catch (e) { }
    }

    // =========================================
    // グラフィックスタイル登録 / Graphic style registration
    // =========================================

    /* 一時グラフィックスタイル名とアクション / Temp graphic style name and action */
    var TEMP_STYLE_NAME = "temp_style";
    var TEMP_STYLE_ACTION_SET = "GraphicStyle";
    var TEMP_STYLE_ACTION_NAME = "AddNewWithoutName";

    /* 選択オブジェクトの見た目を無名グラフィックスタイルとして末尾に追加するアクション定義(.aia)
       Dynamic action (.aia) that appends the selection's appearance as an unnamed graphic style */
    var GRAPHIC_STYLE_AIA = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /value 1 } } }';

    /**
     * テキストフレームの見た目を temp_style として登録する（既存があれば作り直す）
     * @param {TextFrame} textFrame - 見た目の元になるテキストフレーム
     * @returns {string|null} 登録したスタイル名。登録できなければ null
     */
    function registerTextFrameAsTempGraphicStyle(textFrame) {
        if (!textFrame) return null;
        var doc = app.activeDocument;
        var graphicStyles = doc.graphicStyles;

        /* 既存の temp_style を削除 / Remove any existing temp_style */
        removeGraphicStyleByName(TEMP_STYLE_NAME);

        /* 登録用に対象だけを選択 / Select only the target for registration */
        doc.selection = null;
        try { textFrame.selected = true; } catch (selectError) { return null; }

        var countBefore = graphicStyles.length;
        loadActionSet(TEMP_STYLE_ACTION_SET, GRAPHIC_STYLE_AIA);
        try { app.doScript(TEMP_STYLE_ACTION_NAME, TEMP_STYLE_ACTION_SET, false); } catch (runError) { }
        unloadActionSet(TEMP_STYLE_ACTION_SET);

        /* 末尾に増えたスタイルを temp_style に改名 / Rename the newly appended style to temp_style */
        if (graphicStyles.length <= countBefore) return null;
        graphicStyles[graphicStyles.length - 1].name = TEMP_STYLE_NAME;
        return TEMP_STYLE_NAME;
    }

    /**
     * 名前でグラフィックスタイルを適用する
     * @param {string} styleName - グラフィックスタイル名
     * @param {PageItem} targetItem - 適用先
     * @returns {boolean} 適用できたら true
     */
    function applyGraphicStyleByName(styleName, targetItem) {
        if (!styleName || !targetItem) return false;
        try {
            app.activeDocument.graphicStyles.getByName(styleName).applyTo(targetItem);
            return true;
        } catch (e) { return false; }
    }

    /**
     * 名前でグラフィックスタイルを削除する（無ければ何もしない）
     * @param {string} styleName - グラフィックスタイル名
     * @returns {void}
     */
    function removeGraphicStyleByName(styleName) {
        if (!styleName) return;
        try { app.activeDocument.graphicStyles.getByName(styleName).remove(); } catch (e) { }
    }

    // =========================================
    // スタイル読み込み / Import graphic styles
    // =========================================

    /**
     * 取り込み用レイヤー「// _imported」を取得する（無ければ作る）。ロックと非表示は解除
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @returns {Layer} 取り込み用レイヤー
     */
    function getOrCreateImportLayer(destinationDoc) {
        var importLayerName = "// _imported";
        var importLayer;
        try {
            importLayer = destinationDoc.layers.getByName(importLayerName);
        } catch (e) {
            importLayer = destinationDoc.layers.add();
            importLayer.name = importLayerName;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    /**
     * パスから表示用のファイル名を取り出す
     * @param {string} filePath - ファイルのパス
     * @returns {string} デコードしたファイル名（失敗したらパスのまま）
     */
    function getDisplayFileName(filePath) {
        try { return decodeURI(new File(filePath).name); } catch (e) { return filePath; }
    }

    /**
     * 名前でグラフィックスタイルを取得する
     * @param {Document} destinationDoc - 探すドキュメント
     * @param {string} styleName - グラフィックスタイル名
     * @returns {GraphicStyle|null} 見つかったスタイル。無ければ null
     */
    function findGraphicStyle(destinationDoc, styleName) {
        try { return destinationDoc.graphicStyles.getByName(styleName); } catch (e) { return null; }
    }

    /**
     * 設定ファイル（前回のスタイルファイルのパスとスタイル名を記憶）を返す
     * @returns {File} 設定ファイル
     */
    function getPrefsFile() {
        return new File(Folder.userData + "/" + PREFS_FILE_NAME);
    }

    /**
     * 記憶しているスタイルファイルのパスとスタイル名を読み込む
     * @returns {{filePath: string, styleNames: string[]}} 記憶していた内容（無ければ空）
     */
    function loadSavedStyleState() {
        var savedState = { filePath: "", styleNames: [] };
        var prefsFile = getPrefsFile();
        if (!prefsFile.exists) return savedState;
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("r");
            var prefsText = prefsFile.read();
            prefsFile.close();
            var prefsLines = prefsText.split(/\r\n|\r|\n/);
            for (var i = 0; i < prefsLines.length; i++) {
                var separatorIndex = prefsLines[i].indexOf("=");
                if (separatorIndex < 0) continue;
                var entryKey = prefsLines[i].substring(0, separatorIndex);
                var entryValue = prefsLines[i].substring(separatorIndex + 1);
                if (entryKey === "styleFilePath") savedState.filePath = entryValue;
                else if (entryKey === "styleNames") savedState.styleNames = entryValue ? entryValue.split("\t") : [];
            }
        } catch (e) { }
        return savedState;
    }

    /**
     * スタイルファイルのパスとスタイル名を記憶する（key=value 形式）
     * @param {string} filePath - スタイルファイルのパス
     * @param {string[]} styleNames - 取り込んだスタイル名
     * @returns {void}
     */
    function saveStyleState(filePath, styleNames) {
        var prefsFile = getPrefsFile();
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("w");
            prefsFile.write("styleFilePath=" + filePath + "\n");
            prefsFile.write("styleNames=" + styleNames.join("\t") + "\n");
            prefsFile.close();
        } catch (e) { }
    }

    /**
     * スタイル用の AI ファイルを選ばせる
     * @returns {string} 選んだファイルのパス。キャンセルなら空文字
     */
    function pickStyleFile() {
        var pickedFile = File.openDialog(getLabel("dialog.pickFile"), function (candidate) {
            return (candidate instanceof Folder) || /\.ai$/i.test(candidate.name);
        });
        return pickedFile ? pickedFile.fsName : "";
    }

    /**
     * スタイルの AI ファイルを開いてスタイル名を取得 → コピー → 一時レイヤーへ貼り付け → レイヤーごと削除する
     * （グラフィックスタイルだけが取り込み先に残る）
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @param {string} filePath - スタイルの AI ファイルのパス
     * @returns {string[]|null} 取り込み先に登録できたスタイル名。ファイルが無ければ警告して null
     */
    function importStylesFrom(destinationDoc, filePath) {
        var styleFile = new File(filePath);
        if (!styleFile.exists) {
            alert(getLabel("alert.fileNotFound") + getDisplayFileName(filePath));
            return null;
        }
        var styleSourceDoc = app.open(styleFile);

        /* 元ファイルのグラフィックスタイル名を取得（index 0 の既定スタイルは除外）
           Collect style names from the source (skip the default style at index 0) */
        var sourceStyleNames = [];
        for (var i = 1; i < styleSourceDoc.graphicStyles.length; i++) {
            sourceStyleNames.push(styleSourceDoc.graphicStyles[i].name);
        }

        /* 作業アートボード内の全てをコピーして保存せず閉じる / Copy in-artboard objects, close without saving */
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);

        /* 一時レイヤーへ貼り付け→登録後にレイヤーごと削除 / Paste to temp layer, then remove it */
        app.activeDocument = destinationDoc;
        var importLayer = getOrCreateImportLayer(destinationDoc);
        destinationDoc.activeLayer = importLayer;
        app.executeMenuCommand("paste");
        try { importLayer.remove(); } catch (e) { }
        /* モーダルダイアログ表示中でも貼り付けの残像を即座に消す / Redraw now so the paste doesn't linger under the modal dialog */
        app.redraw();

        /* 実際にドキュメントへ登録されたスタイル名だけを返す / Keep only names actually registered in the destination */
        var importedStyleNames = [];
        for (var k = 0; k < sourceStyleNames.length; k++) {
            if (findGraphicStyle(destinationDoc, sourceStyleNames[k])) importedStyleNames.push(sourceStyleNames[k]);
        }
        return importedStyleNames;
    }

    /**
     * 指定のスタイルを取り込み先に用意する（未登録なら記憶したファイルから取り込む）
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @param {string} styleName - グラフィックスタイル名
     * @param {string} filePath - スタイルの AI ファイルのパス
     * @returns {boolean} 用意できたら true
     */
    function ensureExternalStyle(destinationDoc, styleName, filePath) {
        if (findGraphicStyle(destinationDoc, styleName)) return true;
        if (importStylesFrom(destinationDoc, filePath) === null) return false;
        if (!findGraphicStyle(destinationDoc, styleName)) {
            alert(getLabel("alert.styleNotFound") + styleName);
            return false;
        }
        return true;
    }

    /**
     * 適用するグラフィックスタイルを決める（読み込んだスタイル、または元テキストの一時スタイル）
     * @param {Object} convertOptions - ダイアログの結果
     * @param {TextFrame} sourceText - 元テキスト
     * @returns {{name: string|null, isTemp: boolean}} スタイル名と、一時スタイルかどうか
     */
    function resolveStyleForSource(convertOptions, sourceText) {
        if (convertOptions && convertOptions.externalStyleName) {
            return { name: convertOptions.externalStyleName, isTemp: false };
        }
        return { name: registerTextFrameAsTempGraphicStyle(sourceText), isTemp: true };
    }

    // =========================================
    // 正確なサイズ計測 / Accurate size measurement
    // =========================================

    /**
     * 選択オブジェクト群の可視バウンディングボックスの和を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]。測れるものが無ければ null
     */
    function getSelectionVisibleBounds(targetItems) {
        if (!targetItems || !targetItems.length) return null;
        var left = null, top = null, right = null, bottom = null;
        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds;
            try { itemBounds = targetItems[i].visibleBounds; } catch (e) { continue; }
            if (!itemBounds) continue;
            if (left === null || itemBounds[0] < left) left = itemBounds[0];
            if (top === null || itemBounds[1] > top) top = itemBounds[1];
            if (right === null || itemBounds[2] > right) right = itemBounds[2];
            if (bottom === null || itemBounds[3] < bottom) bottom = itemBounds[3];
        }
        if (left === null) return null;
        return [left, top, right, bottom];
    }

    /**
     * 複製 → アピアランス分割 → アウトラインで正確な可視サイズを測り、複製を破棄する
     * @param {PageItem} sourceItem - 測るオブジェクト
     * @returns {Object|null} { left, top, right, bottom, width, height }。測れなければ null
     */
    function measureAccurateBounds(sourceItem) {
        var doc = app.activeDocument;
        var savedSelection = doc.selection;
        var measuredBounds = null;
        var duplicatedItem = null;
        try {
            duplicatedItem = sourceItem.duplicate();
            doc.selection = null;
            duplicatedItem.selected = true;
            app.redraw();

            /* アピアランスを分割 / Expand appearance */
            try { app.executeMenuCommand('expandStyle'); } catch (e) { }
            /* テキストのアウトライン / Create outlines */
            try { app.executeMenuCommand('outline'); } catch (e) { }
            app.redraw();

            /* 分割・アウトライン後の選択全体のサイズを計測 / Measure bounds of the resulting selection */
            var expandedItems = doc.selection;
            var bounds = getSelectionVisibleBounds(expandedItems);
            if (bounds) {
                measuredBounds = { left: bounds[0], top: bounds[1], right: bounds[2], bottom: bounds[3], width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
            }

            /* 複製（分割・アウトライン結果）を削除 / Remove the duplicate (expanded/outlined result) */
            for (var i = expandedItems.length - 1; i >= 0; i--) {
                try { expandedItems[i].remove(); } catch (e2) { }
            }
        } catch (e0) {
            if (duplicatedItem) { try { duplicatedItem.remove(); } catch (e3) { } }
        }

        try { doc.selection = savedSelection; } catch (e4) { }
        return measuredBounds;
    }

    // =========================================
    // シェイプ効果 / Shape effect
    // =========================================

    /* ［形状に変換：長方形］ライブエフェクトの定義 / "Convert to Shape: Rectangle" live-effect definition */
    var RECTANGLE_SHAPE_EFFECT_XML = '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 R RelWidth 0 R RelHeight 0 R AbsWidth 0 R AbsHeight 0 R Absolute 0 R CornerRadius 9 "/></LiveEffect>';

    /**
     * 塗りを2枚追加し、長方形のシェイプ効果でボタン状の背景にする（塗り色は設定しない）
     * @param {TextFrame} areaType - 対象のエリア内文字
     * @returns {void}
     */
    function applyButtonShapeAppearance(areaType) {
        try {
            app.activeDocument.selection = null;
            areaType.selected = true;
            app.redraw();
            app.executeMenuCommand('Adobe New Fill Shortcut');
            app.executeMenuCommand('Adobe New Fill Shortcut');
            areaType.applyEffect(RECTANGLE_SHAPE_EFFECT_XML);
        } catch (e) { }
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /**
     * 直線コーナーだけでできた、軸に平行な長方形のパスかどうか
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 長方形なら true
     */
    function isRectanglePath(pageItem) {
        if (!pageItem || pageItem.typename !== "PathItem" || !pageItem.closed) return false;
        var pathPoints;
        try { pathPoints = pageItem.pathPoints; } catch (e) { return false; }
        if (!pathPoints || pathPoints.length !== 4) return false;

        /**
         * 2つの座標がほぼ等しいか
         * @param {number} firstValue - 座標1
         * @param {number} secondValue - 座標2
         * @returns {boolean} 差が 0.01 未満なら true
         */
        function approxEqual(firstValue, secondValue) { return Math.abs(firstValue - secondValue) < 0.01; }

        /* 各コーナーがハンドルを持たない直線か / Each corner is a straight (handle-less) point */
        for (var i = 0; i < pathPoints.length; i++) {
            var anchor = pathPoints[i].anchor, leftDirection = pathPoints[i].leftDirection, rightDirection = pathPoints[i].rightDirection;
            if (!approxEqual(anchor[0], leftDirection[0]) || !approxEqual(anchor[1], leftDirection[1])) return false;
            if (!approxEqual(anchor[0], rightDirection[0]) || !approxEqual(anchor[1], rightDirection[1])) return false;
        }

        /* x・y がそれぞれ2値だけ＝軸並行の長方形 / Exactly two distinct x and y values = axis-aligned rectangle */
        var distinctX = [], distinctY = [];

        /**
         * ほぼ等しい値が無ければ配列に加える
         * @param {number[]} distinctValues - 追加先
         * @param {number} value - 加える値
         * @returns {void}
         */
        function pushDistinct(distinctValues, value) {
            for (var k = 0; k < distinctValues.length; k++) { if (approxEqual(distinctValues[k], value)) return; }
            distinctValues.push(value);
        }
        for (var j = 0; j < pathPoints.length; j++) {
            pushDistinct(distinctX, pathPoints[j].anchor[0]);
            pushDistinct(distinctY, pathPoints[j].anchor[1]);
        }
        return distinctX.length === 2 && distinctY.length === 2;
    }

    /**
     * エリア内文字のフレームに使える図形か（長方形のみ）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 使えるなら true
     */
    function isAreaFrameCandidate(pageItem) {
        return isRectanglePath(pageItem);
    }

    /**
     * geometricBounds を { left, top, width, height } で返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {Object} { left, top, width, height }
     */
    function geometricBoundsOf(pageItem) {
        var bounds = pageItem.geometricBounds;
        return { left: bounds[0], top: bounds[1], width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
    }

    /**
     * 中心を保ったまま、大きさに倍率を掛けた枠を返す
     * @param {Object} bounds - { left, top, width, height }
     * @param {number} widthRatio - 幅の倍率
     * @param {number} heightRatio - 高さの倍率
     * @returns {Object} { left, top, width, height }
     */
    function scaleFrameAroundCenter(bounds, widthRatio, heightRatio) {
        var frameWidth = bounds.width * widthRatio;
        var frameHeight = bounds.height * heightRatio;
        var centerX = bounds.left + bounds.width / 2;
        var centerY = bounds.top - bounds.height / 2;
        return { left: centerX - frameWidth / 2, top: centerY + frameHeight / 2, width: frameWidth, height: frameHeight };
    }

    /**
     * 元テキストの自動カーニングと文字組みアキ量設定を取得する
     * @param {TextFrame} sourceText - 元テキスト
     * @returns {Object} { kerningMethod, mojikumi }（取れなければ null）
     */
    function readKerningAndMojikumi(sourceText) {
        var snapshot = { kerningMethod: null, mojikumi: null };
        try { snapshot.kerningMethod = sourceText.textRange.characterAttributes.kerningMethod; } catch (e) { }
        /* 文字組みアキ量設定「なし」は読むと例外 / Reading mojikumi "None" throws */
        try {
            if (sourceText.paragraphs.length > 0) {
                snapshot.mojikumi = sourceText.paragraphs[0].paragraphAttributes.mojikumi;
            }
        } catch (e2) { }
        return snapshot;
    }

    /**
     * 自動カーニングと文字組みアキ量設定をエリア内文字へ適用する
     * @param {TextFrame} areaType - 適用先のエリア内文字
     * @param {Object} snapshot - readKerningAndMojikumi() の結果
     * @returns {void}
     */
    function applyKerningAndMojikumi(areaType, snapshot) {
        if (snapshot.kerningMethod !== null) {
            try { areaType.textRange.characterAttributes.kerningMethod = snapshot.kerningMethod; } catch (e) { }
        }
        if (snapshot.mojikumi !== null && snapshot.mojikumi !== undefined) {
            var paragraphs = areaType.paragraphs;
            for (var i = 0; i < paragraphs.length; i++) {
                try { paragraphs[i].paragraphAttributes.mojikumi = snapshot.mojikumi; } catch (e2) { }
            }
        }
    }

    /**
     * 図形パスをエリア内文字にして、元テキストの内容・フォント・カーニング・文字組み・グラフィックスタイルを移す
     * @param {Document} doc - 対象のドキュメント
     * @param {PathItem} framePath - 枠にするパス
     * @param {TextFrame} sourceText - 元テキスト
     * @param {string|null} graphicStyleName - 適用するグラフィックスタイル名（null なら適用しない）
     * @returns {TextFrame} 作ったエリア内文字
     */
    function fillAreaTypeFromSourceText(doc, framePath, sourceText, graphicStyleName) {
        var sourceFont = null, sourceSize = 0;
        try {
            var sourceAttributes = sourceText.textRange.characterAttributes;
            sourceFont = sourceAttributes.textFont;
            sourceSize = sourceAttributes.size;
        } catch (e) { }
        var typeSnapshot = readKerningAndMojikumi(sourceText);

        var areaType = doc.textFrames.areaText(framePath);
        areaType.contents = sourceText.contents;
        try {
            if (sourceFont) areaType.textRange.characterAttributes.textFont = sourceFont;
            if (sourceSize > 0) areaType.textRange.characterAttributes.size = sourceSize;
        } catch (e2) { }
        applyKerningAndMojikumi(areaType, typeSnapshot);

        /* 登録したグラフィックスタイルを適用（削除は呼び出し側が finally で行う）
           Apply the registered graphic style (the caller removes it in a finally block) */
        if (graphicStyleName) { applyGraphicStyleByName(graphicStyleName, areaType); }
        return areaType;
    }

    /**
     * エリア内文字の内容を水平・垂直とも中央にする
     * @param {TextFrame} areaType - 対象のエリア内文字
     * @returns {void}
     */
    function centerAreaTypeContents(areaType) {
        try { areaType.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
        /* 縦位置は DOM で不安定なためアクションで中央 / Vertical center via dynamic action (unreliable via DOM) */
        applyAreaTextFrameAlignment(areaType, 1);
    }

    /**
     * 選択からモード（テキスト＋図形／テキストのみ）を判定してエリア内文字へ変換する
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @param {Object} convertOptions - ダイアログの結果
     * @returns {void}
     */
    function convertSelectionToAreaType(doc, currentSelection, convertOptions) {
        preprocessPathTextSelection(doc, currentSelection);

        /* 前処理で選択が差し替わるので取り直す / The pre-process replaces the selection, so read it again */
        var refreshedSelection = doc.selection;
        if (!refreshedSelection || refreshedSelection.length === 0) { return; }

        /* 変換元テキストと、フレームに使える図形（長方形のみ）の有無を判定 / Detect source text and a usable frame shape (rectangles only) */
        var hasSourceText = false, hasFrameShape = false, hasNonRectShape = false;
        for (var i = 0; i < refreshedSelection.length; i++) {
            var selectedItem = refreshedSelection[i];
            if (isPointOrPathText(selectedItem)) hasSourceText = true;
            if (isAreaFrameCandidate(selectedItem)) hasFrameShape = true;
            else if (selectedItem.typename === "PathItem" && selectedItem.closed) hasNonRectShape = true;
        }
        if (!hasSourceText) { return; }

        /* 図形は長方形のみ対応。長方形以外の閉じたパスだけが図形として選ばれている場合は中止
           Only rectangles are supported. Abort when a non-rectangle closed path is the only shape selected. */
        if (hasNonRectShape && !hasFrameShape) { alert(getLabel("alert.rectangleOnly")); return; }

        /* テキスト＋図形 → 図形をフレームに / テキストのみ → 計測した実寸でフレーム
           Text + shape → use the shape as the frame; text only → frame at the measured real size */
        var createdAreaTypes = hasFrameShape
            ? convertTextIntoShape(doc, refreshedSelection, convertOptions)
            : convertPointTextToMeasuredArea(doc, refreshedSelection, convertOptions);

        if (createdAreaTypes.length > 0) {
            try { doc.selection = createdAreaTypes; app.redraw(); } catch (e) { }
        }
    }

    /**
     * テキスト＋図形：図形を複製してエリア内文字にする（元のテキストと図形は削除）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} targetItems - 対象（最初のテキストフレームと最初の長方形を使う）
     * @param {Object} convertOptions - ダイアログの結果
     * @returns {TextFrame[]} 作ったエリア内文字
     */
    function convertTextIntoShape(doc, targetItems, convertOptions) {
        var createdAreaTypes = [];

        var sourceText = null, sourceShape = null;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!sourceText && targetItem.typename === "TextFrame") { sourceText = targetItem; }
            else if (!sourceShape && isAreaFrameCandidate(targetItem)) { sourceShape = targetItem; }
        }
        if (!sourceText || !sourceShape) return createdAreaTypes;

        /* 適用スタイルを決定（外部スタイル or 元テキストの見た目）/ Resolve the style to apply */
        var styleInfo = resolveStyleForSource(convertOptions, sourceText);
        try {
            var framePath = sourceShape.duplicate();
            framePath.filled = false; framePath.stroked = false;
            var areaType = fillAreaTypeFromSourceText(doc, framePath, sourceText, styleInfo.name);
            centerAreaTypeContents(areaType);
            /* 外部スタイル未使用時のみ塗り2枚＋長方形シェイプ効果でボタン状の背景に
               Add a button-like background only when no external style is used */
            if (styleInfo.isTemp) { applyButtonShapeAppearance(areaType); }
            sourceText.remove();
            sourceShape.remove();
            createdAreaTypes.push(areaType);
        } catch (e) {
        } finally {
            /* 一時グラフィックスタイルのみ削除（外部スタイルは残す）/ Remove only the temp style; keep external styles */
            if (styleInfo.isTemp) { removeGraphicStyleByName(styleInfo.name); }
        }

        return createdAreaTypes;
    }

    /**
     * ポイント文字のみ：計測した実寸（大きさ調整の倍率つき）でフレームを作り、中央配置する（元のポイント文字は削除）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} targetItems - 対象（ポイント文字だけを変換）
     * @param {Object} convertOptions - ダイアログの結果
     * @returns {TextFrame[]} 作ったエリア内文字
     */
    function convertPointTextToMeasuredArea(doc, targetItems, convertOptions) {
        var createdAreaTypes = [];

        /* 大きさ調整が有効なら倍率、無効なら等倍 / Ratios when size adjustment is on, 1x when off */
        var isAdjusting = !!(convertOptions && convertOptions.adjust);
        var widthRatio = isAdjusting ? convertOptions.widthRatio : 1;
        var heightRatio = isAdjusting ? convertOptions.heightRatio : 1;

        for (var i = targetItems.length - 1; i >= 0; i--) {
            var sourceText = targetItems[i];
            if (!(sourceText.typename === "TextFrame" && sourceText.kind === TextType.POINTTEXT)) continue;
            /* 適用スタイルを決定（外部スタイル or 元テキストの見た目）/ Resolve the style to apply */
            var styleInfo = resolveStyleForSource(convertOptions, sourceText);
            try {
                /* 複製→アピアランス分割→アウトラインで正確な可視サイズを計測 / Measure accurate visible size */
                var measuredBounds = measureAccurateBounds(sourceText) || geometricBoundsOf(sourceText);
                /* 大きさ調整の倍率を掛け、中心を保ったままフレーム化 / Apply the size-adjustment ratios, keeping the center */
                var frameBounds = scaleFrameAroundCenter(measuredBounds, widthRatio, heightRatio);
                var framePath = doc.pathItems.rectangle(frameBounds.top, frameBounds.left, frameBounds.width, frameBounds.height);
                framePath.filled = false; framePath.stroked = false;
                var areaType = fillAreaTypeFromSourceText(doc, framePath, sourceText, styleInfo.name);
                centerAreaTypeContents(areaType);
                /* 「する」のときは塗り2枚＋長方形シェイプ効果でボタン状の背景を付与（外部スタイル未使用時のみ）
                   When size adjustment is on, add a button-like background (only when no external style is used) */
                if (isAdjusting && styleInfo.isTemp) { applyButtonShapeAppearance(areaType); }
                createdAreaTypes.push(areaType);
                sourceText.remove();
            } catch (e) {
            } finally {
                /* 一時グラフィックスタイルのみ削除（外部スタイルは残す）/ Remove only the temp style; keep external styles */
                if (styleInfo.isTemp) { removeGraphicStyleByName(styleInfo.name); }
            }
        }

        return createdAreaTypes;
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
     * 項目名＋∧∨＋百分率の入力欄＋「%」の行を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - 項目名の LABELS パス
     * @param {number} defaultRatio - 既定の倍率
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addRatioRow(parentPanel, labelPath, defaultRatio) {
        var ratioRow = parentPanel.add("group");
        var ratioLabel = ratioRow.add("statictext", undefined, labelText(labelPath));
        ratioLabel.preferredSize.width = RATIO_LABEL_WIDTH;
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = ratioRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var ratioInput;
        /* 1%以上（0以下は parseRatioPercent() で既定値に戻るため） / at least 1% (0 or less falls back to the default) */
        var stepperGroup = addStepper(stepperInputGroup, function () { return ratioInput; }, { min: 1 });
        ratioInput = stepperInputGroup.add("edittext", undefined, String(Math.round(defaultRatio * 100)));
        bindSteppedArrowKeys(ratioInput, stepperGroup);
        ratioInput.stepperGroup = stepperGroup; /* 有効／無効の切り替えで参照 / used when toggling enabled */
        ratioInput.characters = RATIO_INPUT_CHARACTERS;
        ratioInput.helpTip = getLabel("tooltip.ratio");
        ratioRow.add("statictext", undefined, "%");
        return ratioInput;
    }

    /**
     * 百分率の入力を倍率にする（数値以外・0以下は既定値）
     * @param {string} percentText - 入力された百分率
     * @param {number} fallbackRatio - 既定の倍率
     * @returns {number} 倍率
     */
    function parseRatioPercent(percentText, fallbackRatio) {
        var percent = parseFloat(percentText);
        return (isNaN(percent) || percent <= 0) ? fallbackRatio : percent / 100;
    }

    /**
     * 「大きさ調整」パネル（する／しない＋幅・高さの倍率）を組む
     * @param {Window} optionsDialog - 追加先のダイアログ
     * @returns {{adjustOnRadio: RadioButton, widthInput: EditText, heightInput: EditText}} コントロール
     */
    function buildSizeAdjustPanel(optionsDialog) {
        var sizeAdjustPanel = optionsDialog.add("panel", undefined, getLabel("panel.sizeAdjust"));
        setupPanel(sizeAdjustPanel, 6);

        /* 大きさ調整 する / しない / Size adjustment on / off */
        var adjustModeGroup = sizeAdjustPanel.add("group");
        var adjustOnRadio = adjustModeGroup.add("radiobutton", undefined, getLabel("radio.doAdjust"));
        var adjustOffRadio = adjustModeGroup.add("radiobutton", undefined, getLabel("radio.dontAdjust"));
        adjustOnRadio.helpTip = getLabel("tooltip.doAdjust");
        adjustOffRadio.value = true; /* 既定は「しない」/ Default: off */

        /* 幅・高さの倍率（別々の行、百分率 % で入力）/ Width and height ratios (separate rows, entered as %) */
        var widthInput = addRatioRow(sizeAdjustPanel, "fieldLabel.widthRatio", BUTTON_WIDTH_RATIO);
        var heightInput = addRatioRow(sizeAdjustPanel, "fieldLabel.heightRatio", BUTTON_HEIGHT_RATIO);

        /**
         * 「する」のときだけ倍率欄を使えるようにする
         * @returns {void}
         */
        function updateRatioInputsEnabled() {
            widthInput.enabled = adjustOnRadio.value;
            heightInput.enabled = adjustOnRadio.value;
            /* ∧∨も入力欄と一緒に無効にし、自作描画をディム表示に合わせて描き直す / dim the steppers with the fields and redraw them */
            widthInput.stepperGroup.enabled = adjustOnRadio.value;
            heightInput.stepperGroup.enabled = adjustOnRadio.value;
            redrawSteppersIn(sizeAdjustPanel);
        }
        adjustOnRadio.onClick = updateRatioInputsEnabled;
        adjustOffRadio.onClick = updateRatioInputsEnabled;
        updateRatioInputsEnabled();

        return { adjustOnRadio: adjustOnRadio, widthInput: widthInput, heightInput: heightInput };
    }

    /**
     * 「グラフィックスタイル」「スタイルの読み込み」パネルを組み、読み込みボタンを接続する
     * 「読み込み」で別の AI ファイルを選ぶと、その場でラジオを組み直す
     * @param {Window} optionsDialog - 追加先のダイアログ
     * @param {Document} destinationDoc - スタイルの取り込み先
     * @param {{filePath: string, styleNames: string[]}} styleState - 記憶しているファイルとスタイル名（読み込みで更新）
     * @returns {{refresh: Function, getExternalStyleName: Function}} 表示の更新と、選ばれたスタイル名の取得
     */
    function buildStylePanels(optionsDialog, destinationDoc, styleState) {
        /* グラフィックスタイル（元の見た目／読み込んだスタイル）/ Graphic style (original appearance / loaded styles) */
        var stylePanel = optionsDialog.add("panel", undefined, getLabel("panel.style"));
        setupPanel(stylePanel, 6);
        var styleOriginalRadio = stylePanel.add("radiobutton", undefined, getLabel("radio.styleOriginal"));
        styleOriginalRadio.helpTip = getLabel("tooltip.styleOriginal");
        styleOriginalRadio.value = true; /* 既定は「元の見た目」/ Default: original appearance */

        /* 読み込んだスタイルのラジオを差し替えるためのコンテナ / Container whose radios get rebuilt on reload */
        var importedRadioGroup = stylePanel.add("group");
        importedRadioGroup.orientation = "column";
        importedRadioGroup.alignChildren = ["left", "top"];
        importedRadioGroup.spacing = 6;
        var importedRadios = []; /* { styleRadio, styleName } */

        /* スタイルの読み込みパネル（ボタンの下にファイル名を表示）/ Load-styles panel (filename shown below the button) */
        var loadPanel = optionsDialog.add("panel", undefined, getLabel("panel.loadStyles"));
        setupPanel(loadPanel, 6);
        /* 読み込み / 再読み込みボタンを左寄せで横並び / Load & Reload buttons in a left-aligned row */
        var loadButtonRow = loadPanel.add("group");
        loadButtonRow.alignment = "left";
        var btnLoad = loadButtonRow.add("button", undefined, getLabel("button.load"));
        btnLoad.helpTip = getLabel("tooltip.load"); /* 使い方はツールチップで案内 / Usage hint shown as a tooltip */
        var btnReload = loadButtonRow.add("button", undefined, getLabel("button.reload"));
        btnReload.helpTip = getLabel("tooltip.reload"); /* 記憶したファイルから再取り込み / Re-import from the remembered file */
        var fileNameText = loadPanel.add("statictext", undefined, "", { truncate: "middle" });
        fileNameText.preferredSize.width = FILE_NAME_WIDTH;

        /**
         * 選択中のファイル名表示と再読み込みボタンの有効状態を更新する
         * @returns {void}
         */
        function refreshFileLabel() {
            fileNameText.text = styleState.filePath ? getDisplayFileName(styleState.filePath) : getLabel("status.noFileSelected");
            btnReload.enabled = !!styleState.filePath; /* 記憶したファイルが無ければ再読み込み不可 / Disable Reload without a remembered file */
        }

        /**
         * 読み込んだスタイル名でラジオを組み直す
         * @returns {void}
         */
        function rebuildImportedRadios() {
            for (var i = importedRadioGroup.children.length - 1; i >= 0; i--) {
                importedRadioGroup.remove(importedRadioGroup.children[i]);
            }
            importedRadios = [];
            for (var j = 0; j < styleState.styleNames.length; j++) {
                var styleRadio = importedRadioGroup.add("radiobutton", undefined, styleState.styleNames[j]);
                importedRadios.push({ styleRadio: styleRadio, styleName: styleState.styleNames[j] });
            }
            optionsDialog.layout.layout(true);
            optionsDialog.layout.resize();
        }

        /**
         * スタイルを取り込み、記憶と表示を更新する
         * @param {string} filePath - スタイルの AI ファイルのパス
         * @returns {void}
         */
        function importAndRemember(filePath) {
            var importedStyleNames = importStylesFrom(destinationDoc, filePath);
            if (importedStyleNames === null) return; /* ファイル未検出は importStylesFrom 側で警告済み / Already alerted */
            styleState.filePath = filePath;
            styleState.styleNames = importedStyleNames;
            saveStyleState(filePath, importedStyleNames); /* 次回以降このファイルを参照 / Remember for next runs */
            refreshFileLabel();
            rebuildImportedRadios();
        }

        /* onClick で連結（addEventListener は発火しない環境があるため）/ Use onClick, not addEventListener */
        btnLoad.onClick = function () {
            var pickedPath = pickStyleFile();
            if (!pickedPath) return;
            importAndRemember(pickedPath);
        };

        /* 記憶したファイルを選び直さずに再取り込み（別ドキュメントでも同じファイルを再利用）
           Re-import from the remembered file without re-picking (reuse the same file in another document) */
        btnReload.onClick = function () {
            if (!styleState.filePath) return;
            importAndRemember(styleState.filePath);
        };

        return {
            refresh: function () {
                refreshFileLabel();
                rebuildImportedRadios();
            },
            /* 「元の見た目」なら null、読み込んだスタイルなら選択中の名前 / null for original, else the checked style name */
            getExternalStyleName: function () {
                for (var k = 0; k < importedRadios.length; k++) {
                    if (importedRadios[k].styleRadio.value) return importedRadios[k].styleName;
                }
                return null;
            }
        };
    }

    /**
     * オプションダイアログを表示する
     * @param {Document} destinationDoc - スタイルの取り込み先
     * @param {{filePath: string, styleNames: string[]}} savedStyleState - 記憶しているファイルとスタイル名
     * @returns {Object|null} OK なら { adjust, widthRatio, heightRatio, externalStyleName, styleFilePath }、キャンセルなら null
     */
    function showOptionsDialog(destinationDoc, savedStyleState) {
        var styleState = {
            filePath: (savedStyleState && savedStyleState.filePath) || "",
            styleNames: (savedStyleState && savedStyleState.styleNames) || []
        };

        var optionsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        optionsDialog.orientation = "column";
        optionsDialog.alignChildren = "fill";
        optionsDialog.margins = DIALOG_MARGINS;

        var sizeControls = buildSizeAdjustPanel(optionsDialog);
        var styleControls = buildStylePanels(optionsDialog, destinationDoc, styleState);

        /* ボタン（Mac規約：キャンセル → OK）/ Buttons (Mac order: Cancel → OK) */
        var btnRowGroup = optionsDialog.add("group");
        btnRowGroup.alignment = "right";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        styleControls.refresh();

        var dialogResult = null;
        btnOK.onClick = function () {
            /* 幅・高さは百分率 % 入力を倍率へ換算 / Width/height: convert the % input to a ratio */
            dialogResult = {
                adjust: sizeControls.adjustOnRadio.value,
                widthRatio: parseRatioPercent(sizeControls.widthInput.text, BUTTON_WIDTH_RATIO),
                heightRatio: parseRatioPercent(sizeControls.heightInput.text, BUTTON_HEIGHT_RATIO),
                externalStyleName: styleControls.getExternalStyleName(),
                styleFilePath: styleState.filePath
            };
            optionsDialog.close();
        };
        btnCancel.onClick = function () { dialogResult = null; optionsDialog.close(); };

        prepareDialogWindow(optionsDialog, SCRIPT_NAME);
        optionsDialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、ダイアログの設定でエリア内文字へ変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0 || !containsConvertibleText(currentSelection)) {
            alert(getLabel("alert.selectText"));
            return;
        }

        /* 記憶した参照ファイルとスタイル名を読み込み、ダイアログを表示（キャンセルで中止）
           Load the remembered file/style names, then show the dialog (Cancel aborts) */
        var convertOptions = showOptionsDialog(doc, loadSavedStyleState());
        if (!convertOptions) return;

        /* 読み込んだスタイル選択時は変換前にドキュメントへ用意（未登録なら取り込み）/ Ensure the chosen style is in the document first */
        if (convertOptions.externalStyleName) {
            if (!ensureExternalStyle(doc, convertOptions.externalStyleName, convertOptions.styleFilePath)) return;
            /* 取り込みで選択が外れるため復帰 / Restore selection (import clears it) */
            try { doc.selection = currentSelection; } catch (eSel) { }
        }

        /* フレーム整列アクションを読み込み、終了時に破棄 / Load frame-alignment actions, unload on exit */
        loadAreaTextActions();
        try {
            convertSelectionToAreaType(doc, currentSelection, convertOptions);
        } finally {
            unloadAreaTextActions();
        }
    }

    main();

})();
