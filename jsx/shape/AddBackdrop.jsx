#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "MyScriptEngine"

/*

### 概要

選択したテキストやオブジェクトの背面に、正円・スーパー楕円・長方形の図形を敷きます。
既存の背面図形を一緒に選んでいれば、形状を読み取って置き換えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na8af4a7016ad

### Overview

Lays a circle, superellipse or rectangle behind the selected text or objects.
When an existing backdrop is selected too, the script reads its shape and replaces it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBackdrop.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddBackdrop";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBackdrop.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na8af4a7016ad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var SUPERELLIPSE_EXPONENT     = 2.5;   /* スーパー楕円の指数（2で円、大きいほど四角に近づく） / Superellipse exponent */
    var SUPERELLIPSE_POINT_COUNT  = 8;     /* スーパー楕円のアンカー数（既存図形の判定にも使う） / Anchor count, also used for detection */
    var SUPERELLIPSE_HANDLE_RATIO = 0.35;  /* ハンドル長の係数（大きいほど丸い） / Handle length ratio */
    var ONE_CHAR_DIAMETER_RATIO   = 1.5;   /* 1文字モードの直径（文字サイズに対する倍率） / Diameter per font size in single-character mode */
    var DEFAULT_MARGIN_DIVISOR    = 4;     /* マージンの初期値＝短辺÷この値 / Default margin = short side / this */
    var DEFAULT_ROUND_DIVISOR     = 5;     /* 角丸の初期値＝短辺÷この値 / Default corner radius = short side / this */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X         = 300;               /* 表示時の横方向のずらし量 / Horizontal shift on show */
    var DIALOG_OFFSET_Y         = 0;                 /* 表示時の縦方向のずらし量 / Vertical shift on show */
    var PANEL_MARGINS           = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / Panel margins [left, top, right, bottom] */
    var FIELD_CHARS             = 4;                 /* 数値欄の文字数 / Characters of numeric fields */
    var SHAPE_ROW_BOTTOM_MARGIN = 5;                 /* 形状ラジオの下の余白 / Space below the shape radios */
    var SHORT_FIELD_CHARS       = 3;                 /* 短い数値欄（倍率・マージン・CMYK）の文字数 / Characters of short fields */

    /**
     * 縦並びのパネルを追加する
     * @param {Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @param {string[]} [childAlignment] - alignChildren（省略時は ["fill", "top"]）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentContainer, panelTitle, childAlignment) {
        var createdPanel = parentContainer.add('panel', undefined, panelTitle);
        createdPanel.orientation = 'column';
        createdPanel.alignChildren = childAlignment || ['fill', 'top'];
        createdPanel.margins = PANEL_MARGINS;
        return createdPanel;
    }

    /**
     * 縦並びのグループを追加する
     * @param {Group} parentContainer - 追加先
     * @param {string[]} childAlignment - alignChildren
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parentContainer, childAlignment) {
        var createdGroup = parentContainer.add('group');
        createdGroup.orientation = 'column';
        createdGroup.alignChildren = childAlignment;
        return createdGroup;
    }

    /**
     * 数値欄を、左に∧∨を付けて追加する。下限・上限は setFieldRange()、増減後の処理は bindNumberField() があとから入れる
     * @param {Group} parentContainer - 追加先
     * @param {string} initialText - 初期値
     * @param {string} tooltipText - ツールチップ
     * @param {Object} [stepOptions] - ∧∨の設定（integer など。省略時は小数あり）
     * @returns {EditText} 追加した数値欄（∧∨は .stepperGroup、設定は .stepOptions で参照できる）
     */
    function addNumberField(parentContainer, initialText, tooltipText, stepOptions) {
        var fieldStepOptions = stepOptions || {};
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = parentContainer.add('group');
        stepperFieldGroup.orientation = 'row';
        stepperFieldGroup.alignChildren = ['left', 'center'];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;
        var field;
        var fieldStepper = addStepper(stepperFieldGroup, function () { return field; }, fieldStepOptions);
        field = stepperFieldGroup.add('edittext', undefined, initialText);
        field.helpTip = tooltipText;
        field.characters = FIELD_CHARS;
        field.stepOptions = fieldStepOptions;
        field.stepperGroup = fieldStepper;
        bindSteppedArrowKeys(field, fieldStepper);
        return field;
    }

    /**
     * 数値欄の有効／無効を∧∨ごと切り替える（∧∨は自作描画なので描き直してディム表示をそろえる）
     * @param {EditText} field - addNumberField() で作った数値欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setNumberFieldEnabled(field, isEnabled) {
        field.enabled = isEnabled;
        field.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(field.stepperGroup);
    }

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

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
    var STEPPER_UI_DARK           = isDarkUI();
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

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

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
    var LINK_UI_DARK = isDarkUI();
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

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^﻿/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    /* #targetengine 下の $.global は Illustrator の起動中は残るので、前回の UI 値をここに控える
       $.global survives while Illustrator runs under #targetengine; the last UI values live here */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session");

    /* 保存する値の既定値。_hasSaved が false の間はダイアログの初期値を使う
       Defaults; while _hasSaved is false the dialog keeps its own initial values */
    var DEFAULT_SETTINGS = {
        _hasSaved: false,
        shape: 'circle',          /* 'circle' | 'superellipse' | 'rect' */
        scale: '90',
        oneChar: false,
        marginV: '0',
        marginH: '0',
        marginLink: true,
        marginSquare: false,
        roundEnable: false,
        roundValue: '2',
        pill: false,
        groupWithText: true,
        exclude: false,
        offsetX: '0',
        offsetY: '0',
        kind: 'fill',             /* 'fill' | 'stroke' */
        strokeWidth: '1',
        colorMode: 'black',       /* 'text' | 'black' | 'white' | 'cmyk' */
        cmykC: '0',
        cmykM: '0',
        cmykY: '0',
        cmykK: '0',
        opacityApply: false,
        opacity: '60'
    };
    /* 前回の値（main で読み込む）/ Previous values, loaded in main */
    var dialogState = null;

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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "背面に図形を敷く", en: "Add Backdrop" }
        },
        panel: {
            scale:    { ja: "スケール", en: "Scale" },
            margin:   { ja: "マージン", en: "Margin" },
            round:    { ja: "角丸", en: "Round" },
            grouping: { ja: "グループ", en: "Group" },
            offset:   { ja: "位置の調整", en: "Offset" },
            kind:     { ja: "塗りと線", en: "Fill & Stroke" },
            color:    { ja: "カラー", en: "Color" },
            opacity:  { ja: "不透明度", en: "Opacity" }
        },
        fieldLabel: {
            scale:   { ja: "倍率", en: "Size" },
            marginV: { ja: "上下", en: "Vertical" },
            marginH: { ja: "左右", en: "Horizontal" },
            offsetX: { ja: "X", en: "X" },
            offsetY: { ja: "Y", en: "Y" }
        },
        radio: {
            circle:       { ja: "正円", en: "Circle" },
            superellipse: { ja: "スーパー楕円", en: "Superellipse" },
            rectangle:    { ja: "長方形", en: "Rectangle" },
            fill:         { ja: "塗り", en: "Fill" },
            stroke:       { ja: "線", en: "Stroke" },
            textColor:    { ja: "文字の色", en: "Text Color" },
            black:        { ja: "ブラック", en: "Black" },
            white:        { ja: "ホワイト", en: "White" },
            cmyk:         { ja: "CMYK", en: "CMYK" }
        },
        checkbox: {
            oneChar:       { ja: "1文字", en: "Single Character" },
            square:        { ja: "正方形", en: "Square" },
            pill:          { ja: "ピル形状", en: "Pill shape" },
            groupWithText: { ja: "テキストとグループ化", en: "Group with Text" },
            exclude:       { ja: "中マド処理", en: "Exclude" }
        },
        tooltip: {
            circle:        { ja: "背景を正円にします。（E キー）", en: "Makes the backdrop a circle. (E key)" },
            superellipse: {
                ja: "背景をスーパー楕円（角の丸い四角に近い形）にします。（S キー）",
                en: "Makes the backdrop a superellipse, between a circle and a rounded square. (S key)"
            },
            rectangle:     { ja: "背景を長方形にします。（R キー）", en: "Makes the backdrop a rectangle. (R key)" },
            scale:         { ja: "文字に対する背景の大きさ（％）です。", en: "Size of the backdrop relative to the text, in percent." },
            oneChar: {
                ja: "文字サイズを基準に、1文字分の円を作ります。",
                en: "Sizes the circle from the font size to fit a single character."
            },
            marginV:       { ja: "文字の上下に足す余白です。", en: "Space added above and below the text." },
            marginH:       { ja: "文字の左右に足す余白です。", en: "Space added left and right of the text." },
            marginLink: {
                ja: "上下と左右の余白を同じ値にそろえます。",
                en: "Uses the same value for the vertical and horizontal margins."
            },
            square:        { ja: "背景を正方形にします。", en: "Makes the backdrop a square." },
            round: {
                ja: "背景の角を丸めます。半径は右の欄で指定します。",
                en: "Rounds the corners of the backdrop. The field on the right sets the radius."
            },
            pill:          { ja: "左右の端を半円にして、丸いピル型にします。", en: "Rounds both ends into a pill shape." },
            groupWithText: { ja: "背景と文字を1つのグループにまとめます。", en: "Groups the backdrop with the text." },
            exclude:       { ja: "背景と文字を「中マド」にして、文字部分を抜きます。", en: "Knocks the text out of the backdrop using Exclude." },
            offset:        { ja: "背景の位置を、この値だけずらします。", en: "Nudges the backdrop by this amount." },
            fill:          { ja: "背景に塗りを付けます。", en: "Fills the backdrop." },
            stroke: {
                ja: "背景に線を付けます。太さは右の欄で指定します。",
                en: "Strokes the backdrop. The field on the right sets the weight."
            },
            opacity:       { ja: "背景の不透明度（％）です。", en: "Opacity of the backdrop, in percent." },
            textColor:     { ja: "文字の色をそのまま背景の色に使います。", en: "Uses the text color for the backdrop." },
            black:         { ja: "背景を黒にします。", en: "Makes the backdrop black." },
            white:         { ja: "背景を白にします。", en: "Makes the backdrop white." },
            cmyk:          { ja: "背景の色をCMYKで指定します。", en: "Sets the backdrop color in CMYK." },
            cmykValue:     { ja: "この版の濃度（％）です。", en: "Ink percentage for this plate." },
            transparencyGrid: {
                ja: "透明グリッドの表示／非表示を切り替えます。白い背景の見え方を確かめるときに使います。",
                en: "Toggles the transparency grid, which helps check a white backdrop."
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
        button: {
            transparencyGrid: { ja: "透明グリッドを表示", en: "Show Transparency Grid" },
            ok:               { ja: "OK", en: "OK" },
            cancel:           { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:   { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            genericError:  { ja: "エラーが発生しました", en: "An error occurred" },
            previewError:  { ja: "プレビュー中にエラーが発生しました", en: "An error occurred during preview" },
            groupFailed:   { ja: "グループ化に失敗しました", en: "Failed to group objects" },
            excludeFailed: { ja: "中マド処理（Exclude）の適用に失敗しました", en: "Failed to apply Pathfinder Exclude" }
        }
    };

    /**
     * 見出しに単位を添える（日本語は全角かっこ、英語は半角かっこ）
     * @param {string} titleText - 見出し
     * @param {string} unitLabel - 単位
     * @returns {string} 単位付きの見出し
     */
    function withUnit(titleText, unitLabel) {
        return (uiLang === "ja") ? titleText + "（" + unitLabel + "）" : titleText + " (" + unitLabel + ")";
    }

    // =========================================
    // 共通処理 / Helpers
    // =========================================

    /**
     * デバッグ用のログを ExtendScript Toolkit のコンソールに出す
     * @param {string} message - ログの内容
     * @returns {void}
     */
    function logDebug(message) {
        $.writeln('[' + SCRIPT_NAME + '] ' + message);
    }

    /**
     * オブジェクトを削除する（失敗してもログを出して続行）
     * @param {PageItem} item - 削除するオブジェクト
     * @param {string} itemLabel - ログに出す名前
     * @returns {void}
     */
    function safeRemove(item, itemLabel) {
        if (!item) return;
        try {
            item.remove();
        } catch (e) {
            logDebug('remove failed (' + itemLabel + '): ' + e);
        }
    }

    /**
     * オブジェクトだけを選択してメニューコマンドを実行し、結果のオブジェクトを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象オブジェクト
     * @param {string} menuCommand - executeMenuCommand に渡すコマンド名
     * @returns {PageItem} 実行後に選択されているオブジェクト（取れなければ元のオブジェクト）
     */
    function runMenuCommandOnItem(doc, item, menuCommand) {
        doc.selection = null;
        item.selected = true;
        app.executeMenuCommand(menuCommand);
        var resultItem = (doc.selection && doc.selection.length) ? doc.selection[0] : item;
        doc.selection = null;
        return resultItem;
    }

    /**
     * グループ内のオブジェクトをグループの直前（前面側）へ出す
     * @param {GroupItem} group - 解除するグループ
     * @param {PageItem} keepItem - グループに残すオブジェクト
     * @returns {void}
     */
    function releaseGroupChildren(group, keepItem) {
        for (var i = group.pageItems.length - 1; i >= 0; i--) {
            var child = group.pageItems[i];
            if (child === keepItem) continue;
            child.move(group, ElementPlacement.PLACEBEFORE);
        }
    }

    /**
     * 小数第1位に丸める
     * @param {number} value - 元の値
     * @returns {number} 丸めた値
     */
    function roundToTenth(value) {
        return Math.round(value * 10) / 10;
    }

    /**
     * グレーのカラーを作る
     * @param {number} grayValue - 濃度（0＝白、100＝黒）
     * @returns {GrayColor} カラー
     */
    function makeGrayColor(grayValue) {
        var color = new GrayColor();
        color.gray = grayValue;
        return color;
    }

    /**
     * 境界 [左, 上, 右, 下] を寸法付きの矩形情報にする
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function boundsToRect(bounds) {
        return {
            left: bounds[0],
            top: bounds[1],
            right: bounds[2],
            bottom: bounds[3],
            w: bounds[2] - bounds[0],
            h: bounds[1] - bounds[3]
        };
    }

    // =========================================
    // 数値欄 / Numeric fields
    // =========================================

    /**
     * 数値欄に許容範囲を設定する（入力時と∧∨・↑↓キーの両方で使う）
     * @param {EditText} editText - 対象の数値欄
     * @param {number|null} minValue - 最小値（制限しないときは null）
     * @param {number|null} maxValue - 最大値（制限しないときは null）
     * @returns {void}
     */
    function setFieldRange(editText, minValue, maxValue) {
        editText.rangeMin = minValue;
        editText.rangeMax = maxValue;
        /* ∧∨の範囲は undefined で「制限なし」 / the stepper treats undefined as unlimited */
        editText.stepOptions.min = (minValue === null) ? undefined : minValue;
        editText.stepOptions.max = (maxValue === null) ? undefined : maxValue;
    }

    /**
     * 数値を数値欄の許容範囲に収める
     * @param {EditText} editText - 対象の数値欄
     * @param {number} value - 元の値
     * @returns {number} 範囲内に収めた値
     */
    function clampToFieldRange(editText, value) {
        if (typeof editText.rangeMin === 'number' && value < editText.rangeMin) return editText.rangeMin;
        if (typeof editText.rangeMax === 'number' && value > editText.rangeMax) return editText.rangeMax;
        return value;
    }

    /**
     * 数値欄の入力と∧∨・↑↓キーに更新処理をつなぐ（入力中の値も許容範囲に収める）
     * @param {EditText} editText - addNumberField() で作った数値欄
     * @param {Function} onUpdate - 値が変わったときに呼ぶ処理
     * @returns {void}
     */
    function bindNumberField(editText, onUpdate) {
        editText.onChanging = function () {
            var value = parseFloat(editText.text);
            if (!isNaN(value) && clampToFieldRange(editText, value) !== value) {
                editText.text = String(clampToFieldRange(editText, value));
            }
            onUpdate();
        };
        editText.stepOptions.onStep = function () { onUpdate(); };
    }

    // =========================================
    // ラジオボタン / Radio buttons
    // =========================================

    /**
     * ラジオボタンの対応表から、選択中のキーを返す
     * @param {Object} radioMap - キー → RadioButton
     * @param {string} fallbackKey - どれも選ばれていないときのキー
     * @returns {string} 選択中のキー
     */
    function getSelectedRadioKey(radioMap, fallbackKey) {
        for (var key in radioMap) {
            if (radioMap[key].value) return key;
        }
        return fallbackKey;
    }

    /**
     * ラジオボタンの対応表で、指定したキーのボタンだけを選ぶ
     * @param {Object} radioMap - キー → RadioButton
     * @param {string} selectedKey - 選ぶキー
     * @param {string} fallbackKey - selectedKey が対応表に無いときに選ぶキー
     * @returns {void}
     */
    function selectRadioKey(radioMap, selectedKey, fallbackKey) {
        if (!radioMap[selectedKey]) selectedKey = fallbackKey;
        for (var key in radioMap) {
            radioMap[key].value = (key === selectedKey);
        }
    }

    // =========================================
    // 形状 / Shape geometry
    // =========================================

    /**
     * 符号を返す
     * @param {number} value - 対象の値
     * @returns {number} 1 / -1 / 0
     */
    function signOf(value) {
        return ((value > 0) - (value < 0)) || +value;
    }

    /**
     * スーパー楕円のアンカーポイントを、角度で等分して求める
     * @param {number} cx - 中心X
     * @param {number} cy - 中心Y
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @param {number} exponent - 指数
     * @param {number} pointCount - アンカー数
     * @returns {number[][]} アンカー座標の配列
     */
    function buildSuperellipseAnchorPoints(cx, cy, width, height, exponent, pointCount) {
        var anchors = [];
        for (var i = 0; i < pointCount; i++) {
            var theta = (Math.PI * 2 * i) / pointCount;
            var cosTheta = Math.cos(theta);
            var sinTheta = Math.sin(theta);
            var x = Math.pow(Math.abs(cosTheta), 2 / exponent) * (width / 2) * signOf(cosTheta);
            var y = Math.pow(Math.abs(sinTheta), 2 / exponent) * (height / 2) * signOf(sinTheta);
            anchors.push([cx + x, cy + y]);
        }
        return anchors;
    }

    /**
     * パスをスーパー楕円に置き換え、各アンカーにスムーズなハンドルを付ける
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} cx - 中心X
     * @param {number} cy - 中心Y
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @returns {void}
     */
    function morphPathToSuperellipse(pathItem, cx, cy, width, height) {
        var anchors = buildSuperellipseAnchorPoints(cx, cy, width, height, SUPERELLIPSE_EXPONENT, SUPERELLIPSE_POINT_COUNT);
        pathItem.setEntirePath(anchors);
        pathItem.closed = true;

        var pathPoints = pathItem.pathPoints;
        var pointCount = anchors.length;
        for (var i = 0; i < pointCount; i++) {
            var prevAnchor = anchors[(i - 1 + pointCount) % pointCount];
            var anchor = anchors[i];
            var nextAnchor = anchors[(i + 1) % pointCount];

            /* 前後のアンカーを結ぶ向きを接線にする / Tangent from the previous to the next anchor */
            var tangentX = nextAnchor[0] - prevAnchor[0];
            var tangentY = nextAnchor[1] - prevAnchor[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY);
            if (tangentLength === 0) continue;
            tangentX /= tangentLength;
            tangentY /= tangentLength;

            /* 短いほうの辺に合わせてハンドル長を決める / Handle length from the shorter neighboring segment */
            var prevLength = Math.sqrt(Math.pow(anchor[0] - prevAnchor[0], 2) + Math.pow(anchor[1] - prevAnchor[1], 2));
            var nextLength = Math.sqrt(Math.pow(nextAnchor[0] - anchor[0], 2) + Math.pow(nextAnchor[1] - anchor[1], 2));
            var handleLength = Math.min(prevLength, nextLength) * SUPERELLIPSE_HANDLE_RATIO;

            pathPoints[i].anchor = anchor;
            pathPoints[i].leftDirection = [anchor[0] - tangentX * handleLength, anchor[1] - tangentY * handleLength];
            pathPoints[i].rightDirection = [anchor[0] + tangentX * handleLength, anchor[1] + tangentY * handleLength];
            pathPoints[i].pointType = PointType.SMOOTH;
        }
    }

    /**
     * ライブ効果「角を丸くする」を適用する
     * @param {PageItem} item - 対象オブジェクト
     * @param {number} radius - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersLive(item, radius) {
        if (!radius || radius <= 0) return;
        try {
            item.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>');
        } catch (e) {
            logDebug('applyRoundCornersLive failed: ' + e);
        }
    }

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // =========================================
    // 対象の特定 / Target resolution
    // =========================================

    /**
     * 選択から、置き換え対象の既存の背面図形を探す
     * - 単一グループ（テキスト＋図形）を選択しているとき：グループ内のテキスト以外の最初のオブジェクト
     * - 複数選択（テキストまたはグループ＋図形）のとき：最初のパス／複合パス
     * @param {PageItem[]} selectionItems - 選択オブジェクト
     * @param {TextFrame|null} textItem - 対象テキスト
     * @returns {{backdrop: PageItem, enclosingGroup: GroupItem, target: PageItem}} 見つからなければ各値は null
     */
    function findExistingBackdrop(selectionItems, textItem) {
        var found = { backdrop: null, enclosingGroup: null, target: null };
        var isTextMode = !!textItem;

        if (selectionItems.length === 1 && selectionItems[0].typename === 'GroupItem' && isTextMode) {
            var outerGroup = selectionItems[0];
            for (var i = 0; i < outerGroup.pageItems.length; i++) {
                var groupChild = outerGroup.pageItems[i];
                if (groupChild === textItem || groupChild.typename === 'TextFrame') continue;
                found.backdrop = groupChild;
                found.enclosingGroup = outerGroup;
                return found;
            }
        }

        if (selectionItems.length >= 2) {
            var otherTargets = [];
            var shapeCandidates = [];
            for (var j = 0; j < selectionItems.length; j++) {
                var selectedItem = selectionItems[j];
                if (selectedItem === textItem) continue;
                if (selectedItem.typename === 'TextFrame' || selectedItem.typename === 'GroupItem') {
                    otherTargets.push(selectedItem);
                } else if (selectedItem.typename === 'PathItem' || selectedItem.typename === 'CompoundPathItem') {
                    shapeCandidates.push(selectedItem);
                }
            }
            if (shapeCandidates.length > 0 && (isTextMode || otherTargets.length > 0)) {
                found.backdrop = shapeCandidates[0];
                if (!isTextMode) found.target = otherTargets[0];
            }
        }
        return found;
    }

    /**
     * グループ内から形状判定に使う最初のパスを探す
     * @param {GroupItem} group - 対象グループ
     * @returns {PathItem|null} 見つかったパス
     */
    function findFirstPathInGroup(group) {
        for (var i = 0; i < group.pageItems.length; i++) {
            var child = group.pageItems[i];
            if (child.typename === 'PathItem') return child;
            if (child.typename === 'CompoundPathItem' && child.pathItems.length > 0) return child.pathItems[0];
        }
        return null;
    }

    /**
     * 既存の図形から形状の種類を判定する
     * @param {PageItem} item - 対象オブジェクト
     * @returns {string} 'circle' | 'superellipse' | 'rect'
     */
    function detectShapeType(item) {
        try {
            var path = item;
            if (item.typename === 'CompoundPathItem') path = item.pathItems[0];
            if (item.typename === 'GroupItem') {
                path = findFirstPathInGroup(item);
                if (!path) {
                    /* パスが無ければ縦横比で判定 / Without a path, judge by the aspect ratio */
                    var groupRect = boundsToRect(item.geometricBounds);
                    return (Math.abs(groupRect.w - groupRect.h) < Math.max(groupRect.w, groupRect.h) * 0.05) ? 'circle' : 'rect';
                }
            }
            if (!path || path.typename !== 'PathItem') return 'rect';

            var pathPoints = path.pathPoints;
            if (pathPoints.length === SUPERELLIPSE_POINT_COUNT) return 'superellipse';
            if (pathPoints.length !== 4) return 'rect';

            /* 4点でハンドルがあれば楕円系（正円扱い）、無ければ長方形 / 4 points with handles = ellipse, else rectangle */
            for (var i = 0; i < 4; i++) {
                var anchor = pathPoints[i].anchor;
                var leftHandle = pathPoints[i].leftDirection;
                var rightHandle = pathPoints[i].rightDirection;
                var leftDistance = Math.abs(anchor[0] - leftHandle[0]) + Math.abs(anchor[1] - leftHandle[1]);
                var rightDistance = Math.abs(anchor[0] - rightHandle[0]) + Math.abs(anchor[1] - rightHandle[1]);
                if (leftDistance > 0.5 || rightDistance > 0.5) return 'circle';
            }
            return 'rect';
        } catch (e) {
            return 'rect';
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * Undo でプレビューを巻き戻す管理役
     * - プレビューの更新ごとに Undo 段数を数え、巻き戻しで消す
     * - OK 時は巻き戻してから確定処理を1回だけ実行する（Undo 1段）
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;

        /**
         * プレビューを1段分実行する
         * @param {Function} previewAction - プレビューの処理
         * @returns {void}
         */
        this.addStep = function (previewAction) {
            try {
                previewAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(labelValueText("alert.previewError", e));
            }
        };

        /**
         * これまでのプレビューをすべて巻き戻す
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try {
                    app.undo();
                } catch (e) {
                    logDebug('rollback undo failed: ' + e);
                    break;
                }
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * プレビューを巻き戻してから確定処理を実行する
         * @param {Function} finalAction - 確定処理
         * @returns {void}
         */
        this.confirm = function (finalAction) {
            this.rollback();
            finalAction();
            this.undoDepth = 0;
        };
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 形状のラジオボタンを追加する（ダイアログ上部）
     * @param {Window} dialog - ダイアログ
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addShapeRadios(dialog, ui) {
        var shapeRadioGroup = dialog.add('group');
        shapeRadioGroup.orientation = 'row';
        shapeRadioGroup.alignChildren = ['center', 'center'];
        shapeRadioGroup.alignment = 'center';
        shapeRadioGroup.margins = [0, 0, 0, SHAPE_ROW_BOTTOM_MARGIN];

        ui.rbCircle = shapeRadioGroup.add('radiobutton', undefined, getLabel("radio.circle"));
        ui.rbCircle.helpTip = getLabel("tooltip.circle");
        ui.rbSuperellipse = shapeRadioGroup.add('radiobutton', undefined, getLabel("radio.superellipse"));
        ui.rbSuperellipse.helpTip = getLabel("tooltip.superellipse");
        ui.rbRectangle = shapeRadioGroup.add('radiobutton', undefined, getLabel("radio.rectangle"));
        ui.rbRectangle.helpTip = getLabel("tooltip.rectangle");
        ui.rbCircle.value = true;

        ui.shapeRadios = { circle: ui.rbCircle, superellipse: ui.rbSuperellipse, rect: ui.rbRectangle };
    }

    /**
     * スケールパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addScalePanel(column, ui) {
        ui.scalePanel = addPanel(column, getLabel("panel.scale"));

        var scaleRow = ui.scalePanel.add('group');
        scaleRow.add('statictext', undefined, labelText("fieldLabel.scale"));
        ui.scaleInput = addNumberField(scaleRow, '90', getLabel("tooltip.scale"));
        ui.scaleInput.characters = SHORT_FIELD_CHARS;
        scaleRow.add('statictext', undefined, '%');

        var oneCharRow = ui.scalePanel.add('group');
        oneCharRow.orientation = 'row';
        oneCharRow.alignChildren = ['left', 'center'];
        ui.cbOneChar = oneCharRow.add('checkbox', undefined, getLabel("checkbox.oneChar"));
        ui.cbOneChar.helpTip = getLabel("tooltip.oneChar");
    }

    /**
     * マージンの1行（ラベル・数値欄・単位）を追加する
     * @param {Group} parentContainer - 追加先
     * @param {string} labelKey - fieldLabel / tooltip のキー
     * @param {string} unitLabel - 単位
     * @returns {EditText} 追加した数値欄
     */
    function addMarginRow(parentContainer, labelKey, unitLabel) {
        var marginRow = parentContainer.add('group');
        marginRow.orientation = 'row';
        marginRow.alignChildren = ['left', 'center'];
        marginRow.spacing = 10;
        marginRow.margins = 0;
        marginRow.add('statictext', undefined, labelText("fieldLabel." + labelKey));
        var marginInput = addNumberField(marginRow, '0', getLabel("tooltip." + labelKey));
        marginInput.characters = SHORT_FIELD_CHARS;
        marginRow.add('statictext', undefined, unitLabel);
        return marginInput;
    }

    /**
     * マージンパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} rulerUnitLabel - 定規の単位
     * @returns {void}
     */
    function addMarginPanel(column, ui, rulerUnitLabel) {
        ui.marginPanel = addPanel(column, getLabel("panel.margin"));

        /* 上下・左右の2行の右に連動アイコンを置く / Put the link icon to the right of the two margin rows */
        var marginBodyRow = ui.marginPanel.add('group');
        marginBodyRow.orientation = 'row';
        marginBodyRow.alignChildren = ['left', 'center'];
        marginBodyRow.spacing = 10;
        marginBodyRow.margins = 0;

        var marginFieldsColumn = addColumnGroup(marginBodyRow, ['left', 'center']);
        marginFieldsColumn.spacing = 10;
        marginFieldsColumn.margins = 0;
        ui.marginVInput = addMarginRow(marginFieldsColumn, 'marginV', rulerUnitLabel);
        ui.marginHInput = addMarginRow(marginFieldsColumn, 'marginH', rulerUnitLabel);

        /* 切り替えたあとの処理は bindDialogEvents() で onToggle に入れる / The toggle handler is set in the event wiring */
        ui.marginLinkToggle = addLinkToggle(marginBodyRow, true, function () {
            if (ui.onMarginLinkToggle) ui.onMarginLinkToggle();
        });
        ui.marginLinkToggle.helpTip = getLabel("tooltip.marginLink");

        ui.cbSquare = ui.marginPanel.add('checkbox', undefined, getLabel("checkbox.square"));
        ui.cbSquare.helpTip = getLabel("tooltip.square");
    }

    /**
     * 角丸パネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} rulerUnitLabel - 定規の単位
     * @returns {void}
     */
    function addRoundPanel(column, ui, rulerUnitLabel) {
        ui.roundPanel = addPanel(column, getLabel("panel.round"), ['left', 'top']);

        var roundRow = ui.roundPanel.add('group');
        roundRow.orientation = 'row';
        roundRow.alignChildren = ['left', 'center'];
        ui.cbRoundEnable = roundRow.add('checkbox', undefined, '');
        ui.cbRoundEnable.helpTip = getLabel("tooltip.round");
        ui.roundInput = addNumberField(roundRow, '2', getLabel("tooltip.round"));
        roundRow.add('statictext', undefined, rulerUnitLabel);
        /* OFF→ON で戻す半径 / Radius restored when rounding is turned back on */
        ui.lastRoundRadius = '2';

        var pillRow = ui.roundPanel.add('group');
        pillRow.orientation = 'row';
        pillRow.alignChildren = ['left', 'center'];
        ui.cbPill = pillRow.add('checkbox', undefined, getLabel("checkbox.pill"));
        ui.cbPill.helpTip = getLabel("tooltip.pill");
    }

    /**
     * グループパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addGroupingPanel(column, ui) {
        var groupingPanel = addPanel(column, getLabel("panel.grouping"));
        ui.cbGroupWithText = groupingPanel.add('checkbox', undefined, getLabel("checkbox.groupWithText"));
        ui.cbGroupWithText.helpTip = getLabel("tooltip.groupWithText");
        ui.cbGroupWithText.value = true;

        ui.cbExclude = groupingPanel.add('checkbox', undefined, getLabel("checkbox.exclude"));
        ui.cbExclude.helpTip = getLabel("tooltip.exclude");
    }

    /**
     * 座標パネル（X / Y のずらし量）を追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} rulerUnitLabel - 定規の単位
     * @returns {void}
     */
    function addOffsetPanel(column, ui, rulerUnitLabel) {
        var offsetPanel = addPanel(column, withUnit(getLabel("panel.offset"), rulerUnitLabel));

        var offsetRow = offsetPanel.add('group');
        offsetRow.orientation = 'row';
        offsetRow.alignChildren = ['left', 'center'];

        offsetRow.add('statictext', undefined, labelText("fieldLabel.offsetX"));
        ui.offsetXInput = addNumberField(offsetRow, '0', getLabel("tooltip.offset"));
        offsetRow.add('statictext', undefined, labelText("fieldLabel.offsetY"));
        ui.offsetYInput = addNumberField(offsetRow, '0', getLabel("tooltip.offset"));
    }

    /**
     * 種別パネル（塗り / 線 / 線幅）を追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} strokeUnitLabel - 線幅の単位
     * @returns {void}
     */
    function addKindPanel(column, ui, strokeUnitLabel) {
        var kindPanel = addPanel(column, getLabel("panel.kind"));

        var kindRow = kindPanel.add('group');
        kindRow.orientation = 'row';
        kindRow.alignChildren = ['left', 'center'];
        kindRow.spacing = 10;

        ui.rbFill = kindRow.add('radiobutton', undefined, getLabel("radio.fill"));
        ui.rbFill.helpTip = getLabel("tooltip.fill");
        ui.rbStroke = kindRow.add('radiobutton', undefined, getLabel("radio.stroke"));
        ui.rbStroke.helpTip = getLabel("tooltip.stroke");
        ui.rbFill.value = true;
        ui.kindRadios = { fill: ui.rbFill, stroke: ui.rbStroke };

        ui.strokeWidthInput = addNumberField(kindRow, '1', getLabel("tooltip.stroke"));
        setFieldRange(ui.strokeWidthInput, 0, null);
        kindRow.add('statictext', undefined, strokeUnitLabel);
    }

    /**
     * CMYK の1版分（ラベルの下に数値欄）を追加する
     * @param {Group} parentContainer - 追加先
     * @param {string} plateLabel - 版の名前（C / M / Y / K）
     * @returns {EditText} 追加した数値欄
     */
    function addCmykPlateField(parentContainer, plateLabel) {
        var plateColumn = addColumnGroup(parentContainer, ['fill', 'top']);
        var plateText = plateColumn.add('statictext', undefined, plateLabel);
        plateText.justify = 'center';
        var plateInput = addNumberField(plateColumn, '0', getLabel("tooltip.cmykValue"), { integer: true }); /* 読み取りは parseInt / read with parseInt */
        plateInput.characters = SHORT_FIELD_CHARS;
        setFieldRange(plateInput, 0, 100);
        return plateInput;
    }

    /**
     * カラーパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addColorPanel(column, ui) {
        var colorPanel = addPanel(column, getLabel("panel.color"));

        var colorModeColumn = addColumnGroup(colorPanel, ['left', 'top']);
        ui.rbTextColor = colorModeColumn.add('radiobutton', undefined, getLabel("radio.textColor"));
        ui.rbTextColor.helpTip = getLabel("tooltip.textColor");
        ui.rbBlack = colorModeColumn.add('radiobutton', undefined, getLabel("radio.black"));
        ui.rbBlack.helpTip = getLabel("tooltip.black");
        ui.rbWhite = colorModeColumn.add('radiobutton', undefined, getLabel("radio.white"));
        ui.rbWhite.helpTip = getLabel("tooltip.white");
        ui.rbCmyk = colorModeColumn.add('radiobutton', undefined, getLabel("radio.cmyk"));
        ui.rbCmyk.helpTip = getLabel("tooltip.cmyk");
        ui.rbBlack.value = true;
        ui.colorRadios = { text: ui.rbTextColor, black: ui.rbBlack, white: ui.rbWhite, cmyk: ui.rbCmyk };

        ui.cmykGroup = colorPanel.add('group');
        ui.cmykGroup.orientation = 'row';
        ui.cmykGroup.alignChildren = ['center', 'top'];
        /* キーはセッション記憶の項目名と同じ / Keys match the session-memory fields */
        ui.cmykInputs = {
            cmykC: addCmykPlateField(ui.cmykGroup, 'C'),
            cmykM: addCmykPlateField(ui.cmykGroup, 'M'),
            cmykY: addCmykPlateField(ui.cmykGroup, 'Y'),
            cmykK: addCmykPlateField(ui.cmykGroup, 'K')
        };
    }

    /**
     * 不透明度パネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addOpacityPanel(column, ui) {
        var opacityPanel = addPanel(column, getLabel("panel.opacity"));

        var opacityRow = opacityPanel.add('group');
        opacityRow.orientation = 'row';
        ui.cbOpacityApply = opacityRow.add('checkbox', undefined, '');
        ui.cbOpacityApply.helpTip = getLabel("tooltip.opacity");
        ui.opacityInput = addNumberField(opacityRow, '60', getLabel("tooltip.opacity"));
        ui.opacityInput.characters = 3;
        opacityRow.add('statictext', undefined, '%');
    }

    /**
     * ダイアログとコントロールを組み立てる（イベントはつながない）
     * @returns {Object} ダイアログ（dialog）と各コントロール
     */
    function buildDialog() {
        var ui = {};
        var rulerUnitLabel = getUnitInfo('rulerType').label;
        var strokeUnitLabel = getUnitInfo('strokeUnits').label;

        var dialog = new Window('dialog', getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        ui.dialog = dialog;

        addShapeRadios(dialog, ui);

        var columnsRow = dialog.add('group');
        columnsRow.orientation = 'row';
        columnsRow.alignChildren = ['fill', 'top'];
        var leftColumn = addColumnGroup(columnsRow, ['fill', 'top']);
        var rightColumn = addColumnGroup(columnsRow, ['fill', 'top']);

        addScalePanel(leftColumn, ui);
        addMarginPanel(leftColumn, ui, rulerUnitLabel);
        addRoundPanel(leftColumn, ui, rulerUnitLabel);
        addGroupingPanel(leftColumn, ui);

        addOffsetPanel(rightColumn, ui, rulerUnitLabel);
        addKindPanel(rightColumn, ui, strokeUnitLabel);
        addColorPanel(rightColumn, ui);
        addOpacityPanel(rightColumn, ui);

        /* ボタンエリア（左・スペーサー・右） / Button row (left, spacer, right) */
        var buttonRow = addButtonRow(dialog);
        ui.btnTransparencyGrid = buttonRow.leftGroup.add('button', undefined, getLabel("button.transparencyGrid"));
        ui.btnTransparencyGrid.helpTip = getLabel("tooltip.transparencyGrid");
        ui.btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel("button.cancel"), { name: 'cancel' });
        ui.btnOK = buttonRow.rightGroup.add('button', undefined, getLabel("button.ok"), { name: 'ok' });

        return ui;
    }

    // =========================================
    // ダイアログの連動 / Dialog state sync
    // =========================================

    /**
     * スケールパネルの有効／無効を切り替える（長方形は100%固定。ただし正方形ONのときは有効）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncScalePanel(ui) {
        if (ui.rbRectangle.value && !ui.cbSquare.value) {
            ui.scaleInput.text = '100';
            ui.cbOneChar.value = false;
            ui.scalePanel.enabled = false;
        } else {
            ui.scalePanel.enabled = true;
        }
        redrawSteppersIn(ui.scalePanel);
    }

    /**
     * マージンの「連動」を反映する（ONなら左右を上下にそろえて無効化）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncMarginLink(ui) {
        if (ui.marginLinkToggle.value) ui.marginHInput.text = String(ui.marginVInput.text);
        setNumberFieldEnabled(ui.marginHInput, !ui.marginLinkToggle.value);
    }

    /**
     * マージンパネルは長方形のときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncMarginPanel(ui) {
        ui.marginPanel.enabled = ui.rbRectangle.value;
        redrawSteppersIn(ui.marginPanel);
        /* 親の無効化は自作描画に出ないので描き直す / Parent disabling does not reach custom drawing, so redraw */
        redrawLinkToggle(ui.marginLinkToggle);
    }

    /**
     * 角丸の数値欄の有効／無効を切り替える（ピル形状は半径が自動なので無効）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncRoundInput(ui) {
        if (ui.cbPill.value) {
            setNumberFieldEnabled(ui.roundInput, false);
            return;
        }
        if (ui.cbRoundEnable.value) {
            /* OFF→ON で控えた値を戻す / Restore the saved value when turned on */
            if (!ui.roundInput.enabled) ui.roundInput.text = String(ui.lastRoundRadius || ui.roundInput.text || '0');
            setNumberFieldEnabled(ui.roundInput, true);
        } else {
            /* ON→OFF で値を控えて無効化 / Save the value and disable when turned off */
            ui.lastRoundRadius = String(ui.roundInput.text);
            setNumberFieldEnabled(ui.roundInput, false);
        }
    }

    /**
     * 角丸パネルは長方形のときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncRoundPanel(ui) {
        ui.roundPanel.enabled = ui.rbRectangle.value;
        syncRoundInput(ui);
        redrawSteppersIn(ui.roundPanel);
    }

    /**
     * 種別に合わせて線幅と中マド処理を切り替える（線のときは中マド処理を使えない）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncKind(ui) {
        var strokeOn = ui.rbStroke.value;
        setNumberFieldEnabled(ui.strokeWidthInput, strokeOn);
        if (strokeOn) ui.cbExclude.value = false;
        ui.cbExclude.enabled = !strokeOn;
    }

    /**
     * CMYK 欄は CMYK を選んだときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncColor(ui) {
        ui.cmykGroup.enabled = ui.rbCmyk.value;
        redrawSteppersIn(ui.cmykGroup);
    }

    /**
     * 不透明度の数値欄はチェックONのときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncOpacity(ui) {
        setNumberFieldEnabled(ui.opacityInput, ui.cbOpacityApply.value);
    }

    /**
     * 形状に連動するパネルをまとめて切り替える
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncShapeDependentPanels(ui) {
        syncRoundPanel(ui);
        syncMarginPanel(ui);
        syncScalePanel(ui);
    }

    /**
     * すべての連動をまとめて反映する
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncAllPanels(ui) {
        syncColor(ui);
        syncKind(ui);
        syncMarginLink(ui);
        syncOpacity(ui);
        syncShapeDependentPanels(ui);
    }

    // =========================================
    // ダイアログの記憶 / Dialog memory
    // =========================================

    /**
     * 数値欄とセッション記憶の項目の対応表を返す
     * @param {Object} ui - コントロール
     * @returns {Object} 項目名 → EditText
     */
    function getTextFieldMap(ui) {
        var fieldMap = {
            scale: ui.scaleInput,
            marginV: ui.marginVInput,
            marginH: ui.marginHInput,
            roundValue: ui.roundInput,
            offsetX: ui.offsetXInput,
            offsetY: ui.offsetYInput,
            strokeWidth: ui.strokeWidthInput,
            opacity: ui.opacityInput
        };
        for (var key in ui.cmykInputs) fieldMap[key] = ui.cmykInputs[key];
        return fieldMap;
    }

    /**
     * チェックボックスとセッション記憶の項目の対応表を返す
     * @param {Object} ui - コントロール
     * @returns {Object} 項目名 → Checkbox
     */
    function getCheckboxMap(ui) {
        return {
            oneChar: ui.cbOneChar,
            marginSquare: ui.cbSquare,
            roundEnable: ui.cbRoundEnable,
            pill: ui.cbPill,
            groupWithText: ui.cbGroupWithText,
            exclude: ui.cbExclude,
            opacityApply: ui.cbOpacityApply
        };
    }

    /**
     * ダイアログの値をセッション記憶に控える
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function saveDialogState(ui) {
        var fieldMap = getTextFieldMap(ui);
        var checkboxMap = getCheckboxMap(ui);
        var key;
        for (key in fieldMap) dialogState[key] = String(fieldMap[key].text);
        for (key in checkboxMap) dialogState[key] = !!checkboxMap[key].value;
        dialogState.marginLink = !!ui.marginLinkToggle.value;
        dialogState.shape = getSelectedRadioKey(ui.shapeRadios, 'circle');
        dialogState.kind = getSelectedRadioKey(ui.kindRadios, 'fill');
        dialogState.colorMode = getSelectedRadioKey(ui.colorRadios, 'black');
        dialogState._hasSaved = true;
        settingsStore.save(dialogState);
    }

    /**
     * セッション記憶の値をダイアログに戻す
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function restoreDialogState(ui) {
        if (!dialogState._hasSaved) return;
        var fieldMap = getTextFieldMap(ui);
        var checkboxMap = getCheckboxMap(ui);
        var key;
        for (key in fieldMap) fieldMap[key].text = String(dialogState[key]);
        for (key in checkboxMap) checkboxMap[key].value = !!dialogState[key];
        setLinkToggleValue(ui.marginLinkToggle, !!dialogState.marginLink);
        selectRadioKey(ui.shapeRadios, dialogState.shape, 'circle');
        selectRadioKey(ui.kindRadios, dialogState.kind, 'fill');
        selectRadioKey(ui.colorRadios, dialogState.colorMode, 'black');
        ui.lastRoundRadius = String(dialogState.roundValue);
    }

    // =========================================
    // 計測 / Measurement
    // =========================================

    /**
     * テキストの見た目の境界を測る（複製をアウトライン化して測り、結果は控えて使い回す）
     * @param {Object} context - 対象の情報
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function measureTextVisualBounds(context) {
        if (context.textVisualRect) return context.textVisualRect;
        var textDuplicate = null;
        var outlinedGroup = null;
        try {
            textDuplicate = context.textItem.duplicate();
            outlinedGroup = textDuplicate.createOutline();
            context.textVisualRect = boundsToRect(outlinedGroup.geometricBounds);
        } catch (e) {
            /* アウトライン化できないときはフレームの境界 / Fall back to the frame bounds */
            context.textVisualRect = boundsToRect(context.textItem.geometricBounds);
        } finally {
            safeRemove(outlinedGroup, 'outlinedGroup');
            safeRemove(textDuplicate, 'textDuplicate');
        }
        return context.textVisualRect;
    }

    /**
     * 選択オブジェクト全体の境界を測る（置き換える既存の背面図形は除く）
     * @param {Object} context - 対象の情報
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function measureSelectionRect(context) {
        var measuredItems = [];
        for (var i = 0; i < context.selectionItems.length; i++) {
            var item = context.selectionItems[i];
            if (!item || item === context.existingBackdrop) continue;
            measuredItems.push(item);
        }
        /* プレビュー境界で測る（クリップグループはマスクの範囲） / Measure preview bounds (clip groups by their mask) */
        var unionBounds = getClipAwareUnionBounds(measuredItems, true);
        return boundsToRect(unionBounds || getClipAwareBounds(context.targetItem, true));
    }

    /**
     * 背面図形の基準にする矩形を返す
     * - テキスト：見た目の境界（1文字モードではフレームの境界）
     * - テキスト以外：選択全体の境界
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function measureBaseRect(ui, context) {
        if (!context.isTextMode) return measureSelectionRect(context);
        if (ui.cbOneChar.value) return boundsToRect(context.textItem.geometricBounds);
        return measureTextVisualBounds(context);
    }

    /**
     * マージンとスケールの初期値を、対象の短辺から決める（前回の値が無いときだけ呼ぶ）
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function applyDefaultsFromTarget(ui, context) {
        var baseRect = context.isTextMode ? measureTextVisualBounds(context) : measureSelectionRect(context);
        var shortSide = Math.min(Math.abs(baseRect.w), Math.abs(baseRect.h));
        if (!isFinite(shortSide) || shortSide <= 0) return;

        var defaultMargin = String(roundToTenth(shortSide / DEFAULT_MARGIN_DIVISOR));
        ui.marginVInput.text = defaultMargin;
        ui.marginHInput.text = defaultMargin;
        syncMarginLink(ui);

        ui.roundInput.text = String(roundToTenth(shortSide / DEFAULT_ROUND_DIVISOR));
        ui.lastRoundRadius = ui.roundInput.text;
    }

    /**
     * 数値欄の値を数値で読む
     * @param {EditText} editText - 数値欄
     * @param {number} fallbackValue - 数値にならないときの値
     * @returns {number} 読み取った値
     */
    function readNumber(editText, fallbackValue) {
        var value = parseFloat(editText.text);
        return isNaN(value) ? fallbackValue : value;
    }

    /**
     * ダイアログの値から、背面図形の中心と寸法を求める
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {{cx: number, cy: number, diameter: number, rectW: number, rectH: number, baseW: number, scaleRatio: number}} 図形の寸法
     */
    function computeBackdropGeometry(ui, context) {
        var baseRect = measureBaseRect(ui, context);

        var marginY = readNumber(ui.marginVInput, 0);
        var marginX = ui.marginLinkToggle.value ? marginY : readNumber(ui.marginHInput, 0);
        var paddedW = baseRect.w + marginX * 2;
        var paddedH = baseRect.h + marginY * 2;

        var cx = baseRect.left + baseRect.w / 2;
        var cy = baseRect.top - baseRect.h / 2;

        var scalePercent = readNumber(ui.scaleInput, 100);
        if (scalePercent <= 0) scalePercent = 100;
        var scaleRatio = scalePercent / 100;

        var diameter;
        if (context.isTextMode && ui.cbOneChar.value) {
            /* 1文字モード：文字サイズを基準に、上端から文字サイズの半分を中心にする / Size from the font size */
            var fontSize;
            try {
                fontSize = context.textItem.textRange.characterAttributes.size;
            } catch (e) {
                fontSize = Math.max(baseRect.w, baseRect.h);
            }
            diameter = fontSize * ONE_CHAR_DIAMETER_RATIO * scaleRatio;
            cy = baseRect.top - fontSize / 2;
        } else {
            /* 正方形がすっぽり入る円の直径 / Diameter of the circle that encloses the square */
            diameter = Math.max(paddedW, paddedH) * Math.SQRT2 * scaleRatio;
        }

        cx += readNumber(ui.offsetXInput, 0);
        cy += readNumber(ui.offsetYInput, 0);

        /* 正方形ONのときは正円と同じ直径を一辺にする / A square uses the circle's diameter as its side */
        var squareOn = ui.cbSquare.value;
        return {
            cx: cx,
            cy: cy,
            diameter: diameter,
            rectW: squareOn ? diameter : paddedW * scaleRatio,
            rectH: squareOn ? diameter : paddedH * scaleRatio,
            baseW: baseRect.w,
            scaleRatio: scaleRatio
        };
    }

    // =========================================
    // 背面図形の作成 / Backdrop creation
    // =========================================

    /**
     * 背面図形のカラーを決める
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {Color} カラー
     */
    function resolveBackdropColor(ui, context) {
        if (ui.rbTextColor.value) {
            /* 混在した書式などで読めないことがある。そのときはブラック / Falls back to black when unreadable */
            try {
                var textColor = context.textItem.textRange.characterAttributes.fillColor;
                if (textColor) return textColor;
            } catch (e) { }
        }
        if (ui.rbWhite.value) return makeGrayColor(0);
        if (ui.rbCmyk.value) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykC.text, 10) || 0));
            cmykColor.magenta = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykM.text, 10) || 0));
            cmykColor.yellow = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykY.text, 10) || 0));
            cmykColor.black = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykK.text, 10) || 0));
            return cmykColor;
        }
        return makeGrayColor(100);
    }

    /**
     * 背面図形に塗り／線と不透明度を設定する
     * @param {PageItem} item - 背面図形
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function applyBackdropStyle(item, ui, context) {
        /* 図形の変換後などに色を受け付けないことがある。そのときは薄いグレーの塗り / Falls back to a light gray fill */
        try {
            var backdropColor = resolveBackdropColor(ui, context);
            if (ui.rbFill.value) {
                item.filled = true;
                item.stroked = false;
                item.fillColor = backdropColor;
            } else {
                item.filled = false;
                item.stroked = true;
                var strokeWidth = readNumber(ui.strokeWidthInput, 1);
                item.strokeWidth = (strokeWidth < 0) ? 1 : strokeWidth;
                item.strokeColor = backdropColor;
            }
        } catch (e) {
            try {
                item.fillColor = makeGrayColor(20);
                item.filled = true;
                item.stroked = false;
            } catch (eFallback) { }
        }

        var opacityPercent = Math.max(0, Math.min(100, readNumber(ui.opacityInput, 100)));
        item.opacity = ui.cbOpacityApply.value ? opacityPercent : 100;
    }

    /**
     * 長方形の背面図形を作る（角丸・ピル形状はライブ効果）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - コントロール
     * @param {Object} geometry - 図形の寸法
     * @returns {PathItem} 作った図形
     */
    function createRectangleBackdrop(doc, ui, geometry) {
        var rectTop = geometry.cy + geometry.rectH / 2;

        if (ui.cbPill.value) {
            /* ピル形状：高さの半分を半径にし、幅は少なくとも「元の幅＋高さ」 / Pill: radius is half the height */
            var minPillWidth = Math.abs(geometry.baseW) * geometry.scaleRatio + geometry.rectH;
            var pillWidth = Math.max(geometry.rectW, minPillWidth);
            var pillShape = doc.pathItems.rectangle(rectTop, geometry.cx - pillWidth / 2, pillWidth, geometry.rectH);
            applyRoundCornersLive(pillShape, geometry.rectH / 2);
            return pillShape;
        }

        var rectShape = doc.pathItems.rectangle(rectTop, geometry.cx - geometry.rectW / 2, geometry.rectW, geometry.rectH);
        if (ui.cbRoundEnable.value) applyRoundCornersLive(rectShape, readNumber(ui.roundInput, 0));
        return rectShape;
    }

    /**
     * 背面図形を作り、対象の背面に置いて色を付ける
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {PathItem} 作った図形
     */
    function createBackdrop(doc, ui, context) {
        var geometry = computeBackdropGeometry(ui, context);
        var backdropShape;

        if (ui.rbRectangle.value) {
            backdropShape = createRectangleBackdrop(doc, ui, geometry);
        } else {
            var diameter = geometry.diameter;
            backdropShape = doc.pathItems.ellipse(geometry.cy + diameter / 2, geometry.cx - diameter / 2, diameter, diameter);
            if (ui.rbSuperellipse.value) morphPathToSuperellipse(backdropShape, geometry.cx, geometry.cy, diameter, diameter);
        }

        /* 対象の直後（背面）へ。移動できない位置なら最背面へ / Place behind the target, or send to back */
        try {
            backdropShape.move(context.targetItem, ElementPlacement.PLACEAFTER);
        } catch (e) {
            try {
                backdropShape.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (eZOrder) {
                logDebug('zOrder fallback failed: ' + eZOrder);
            }
        }

        applyBackdropStyle(backdropShape, ui, context);
        return backdropShape;
    }

    /**
     * 確定時の仕上げ：正円は「シェイプに変換」、長方形は「ライブパスファインダー（合体）」
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} backdropShape - 背面図形
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {PageItem} 仕上げたオブジェクト
     */
    function finalizeBackdropShape(doc, backdropShape, ui, context) {
        if (ui.rbSuperellipse.value) return backdropShape;
        var menuCommand = ui.rbCircle.value ? 'Convert to Shape' : 'Live Pathfinder Add';
        try {
            var finalizedShape = runMenuCommandOnItem(doc, backdropShape, menuCommand);
            applyBackdropStyle(finalizedShape, ui, context);
            return finalizedShape;
        } catch (e) {
            logDebug(menuCommand + ' failed: ' + e);
            return backdropShape;
        }
    }

    /**
     * 背面図形を対象の背面に置く（グループ化がONなら対象と1つのグループにする）
     * @param {PageItem} backdropShape - 背面図形
     * @param {PageItem} targetItem - 対象オブジェクト
     * @param {boolean} groupWithTarget - 対象とグループ化するか
     * @returns {PageItem} 配置後のオブジェクト（グループ化したときはグループ）
     */
    function placeBackdrop(backdropShape, targetItem, groupWithTarget) {
        if (!groupWithTarget) {
            try {
                backdropShape.move(targetItem, ElementPlacement.PLACEAFTER);
            } catch (e) {
                logDebug('ungrouped backdrop move failed: ' + e);
            }
            return backdropShape;
        }

        try {
            var resultGroup = targetItem.parent.groupItems.add();
            try {
                resultGroup.move(targetItem, ElementPlacement.PLACEAFTER);
            } catch (eMoveGroup) {
                logDebug('group move failed: ' + eMoveGroup);
            }
            backdropShape.move(resultGroup, ElementPlacement.PLACEATEND);
            targetItem.move(resultGroup, ElementPlacement.PLACEATEND);
            try {
                backdropShape.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (eSendToBack) {
                logDebug('grouped backdrop zOrder failed: ' + eSendToBack);
            }
            return resultGroup;
        } catch (eGroup) {
            alert(labelValueText("alert.groupFailed", eGroup));
            return backdropShape;
        }
    }

    /**
     * 「ライブパスファインダー（中マド）」で文字部分を抜く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} resultItem - 配置後のオブジェクト
     * @returns {PageItem} 処理後のオブジェクト
     */
    function applyExclude(doc, resultItem) {
        try {
            return runMenuCommandOnItem(doc, resultItem, 'Live Pathfinder Exclude');
        } catch (e) {
            alert(labelValueText("alert.excludeFailed", e));
            doc.selection = null;
            return resultItem;
        }
    }

    // =========================================
    // 既存の背面図形の置き換え / Replacing an existing backdrop
    // =========================================

    /**
     * プレビュー用：既存の背面図形を隠す（グループはいったん解除して隠す。Undo で戻る）
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function hideReplacedBackdrop(context) {
        if (context.enclosingGroup) {
            try {
                releaseGroupChildren(context.enclosingGroup, context.existingBackdrop);
                context.enclosingGroup.hidden = true;
            } catch (e) {
                logDebug('hideReplacedBackdrop ungroup failed: ' + e);
            }
        }
        if (context.existingBackdrop) {
            try {
                context.existingBackdrop.hidden = true;
            } catch (e) {
                logDebug('hideReplacedBackdrop hide failed: ' + e);
            }
        }
    }

    /**
     * 確定用：既存の背面図形を削除する（グループは中身を出してから削除）
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function removeReplacedBackdrop(context) {
        if (context.enclosingGroup) {
            try {
                releaseGroupChildren(context.enclosingGroup, context.existingBackdrop);
            } catch (e) {
                logDebug('removeReplacedBackdrop ungroup failed: ' + e);
            }
        }
        safeRemove(context.existingBackdrop, 'existingBackdrop');
        safeRemove(context.enclosingGroup, 'enclosingGroup');
        context.existingBackdrop = null;
        context.enclosingGroup = null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から処理対象を決める
     * @param {PageItem[]} selectionItems - 選択オブジェクト
     * @returns {Object} 対象の情報（テキスト、配置の基準、既存の背面図形など）
     */
    function resolveContext(selectionItems) {
        var textItem = collectSelectionTextFrames(selectionItems)[0] || null;
        var found = findExistingBackdrop(selectionItems, textItem);
        return {
            selectionItems: selectionItems,
            textItem: textItem,
            isTextMode: !!textItem,
            /* テキストが無いときは、グループなど選択の先頭を配置の基準にする / Without text, the first selected item is the anchor */
            targetItem: textItem || found.target || selectionItems[0],
            existingBackdrop: found.backdrop,
            enclosingGroup: found.enclosingGroup,
            textVisualRect: null
        };
    }

    /**
     * 対象に合わせてダイアログの初期状態を調整する
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function adjustDialogForTarget(ui, context) {
        /* 既存の背面図形があれば形状を合わせる / Match the shape of an existing backdrop */
        if (context.existingBackdrop) {
            selectRadioKey(ui.shapeRadios, detectShapeType(context.existingBackdrop), 'rect');
            syncShapeDependentPanels(ui);
        }

        if (!context.isTextMode) {
            /* テキストが無いときは「1文字」と「テキストカラー」を使わない / No single-character or text color without text */
            ui.cbOneChar.value = false;
            ui.cbOneChar.enabled = false;
            if (ui.rbTextColor.value) selectRadioKey(ui.colorRadios, 'black', 'black');
            ui.rbTextColor.enabled = false;
            return;
        }

        /* 空白を除いて1文字だけなら「1文字」をON / Turn on single-character mode for a single visible character */
        if (String(context.textItem.contents || '').replace(/\s+/g, '').length === 1) {
            ui.cbOneChar.value = true;
        }
    }

    /**
     * キー入力で形状を切り替える（E: 正円 / S: スーパー楕円 / R: 長方形）
     * 数値欄にフォーカスがあっても効く（押した文字は欄に入らない）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function addShapeKeyHandler(ui) {
        /* ダイアログの入力欄はすべて数値 / Every field in this dialog is numeric */
        var numericFields = [
            ui.scaleInput, ui.marginVInput, ui.marginHInput, ui.roundInput,
            ui.offsetXInput, ui.offsetYInput, ui.strokeWidthInput, ui.opacityInput
        ];
        for (var plateKey in ui.cmykInputs) numericFields.push(ui.cmykInputs[plateKey]);
        /* ラジオを選んで onClick（形状に応じたパネルの同期とプレビュー）/ Select the radio and run its onClick */
        addKeyShortcuts(ui.dialog, {
            "E": ui.shapeRadios.circle,
            "S": ui.shapeRadios.superellipse,
            "R": ui.shapeRadios.rect
        }, { numericFields: numericFields });
    }

    /**
     * ダイアログのイベントをつなぐ
     * @param {Object} ui - コントロール
     * @param {Function} updatePreview - プレビューを更新する処理
     * @returns {void}
     */
    function bindDialogEvents(ui, updatePreview) {
        /* 形状 / Shape */
        var onShapeChanged = function () {
            syncShapeDependentPanels(ui);
            updatePreview();
        };
        for (var shapeKey in ui.shapeRadios) ui.shapeRadios[shapeKey].onClick = onShapeChanged;
        addShapeKeyHandler(ui);

        /* スケール / Scale */
        bindNumberField(ui.scaleInput, updatePreview);
        ui.cbOneChar.onClick = updatePreview;

        /* マージン / Margin */
        bindNumberField(ui.marginVInput, function () {
            syncMarginLink(ui);
            updatePreview();
        });
        bindNumberField(ui.marginHInput, updatePreview);
        ui.onMarginLinkToggle = function () {
            syncMarginLink(ui);
            updatePreview();
        };
        ui.cbSquare.onClick = function () {
            /* 正方形ONでピル形状をOFFにし、スケールを使えるようにする / Square turns pill off and enables scale */
            if (ui.cbSquare.value) {
                ui.cbPill.value = false;
                syncRoundInput(ui);
            }
            syncScalePanel(ui);
            if (ui.cbSquare.value) ui.scaleInput.active = true;
            updatePreview();
        };

        /* 角丸 / Round */
        bindNumberField(ui.roundInput, function () {
            ui.lastRoundRadius = String(ui.roundInput.text);
            updatePreview();
        });
        ui.cbRoundEnable.onClick = function () {
            syncRoundInput(ui);
            if (ui.cbRoundEnable.value && !ui.cbPill.value) ui.roundInput.active = true;
            updatePreview();
        };
        ui.cbPill.onClick = function () {
            syncRoundInput(ui);
            updatePreview();
        };

        /* グループ / Grouping */
        ui.cbExclude.onClick = function () {
            /* 中マド処理にはグループ化とテキストカラーが要る / Exclude needs grouping and the text color */
            if (ui.cbExclude.value) {
                ui.cbGroupWithText.value = true;
                selectRadioKey(ui.colorRadios, 'text', 'text');
                syncColor(ui);
            }
            updatePreview();
        };
        ui.cbGroupWithText.onClick = function () {
            if (!ui.cbGroupWithText.value) ui.cbExclude.value = false;
            updatePreview();
        };

        /* 座標 / Offset */
        bindNumberField(ui.offsetXInput, updatePreview);
        bindNumberField(ui.offsetYInput, updatePreview);

        /* 種別 / Kind */
        ui.rbFill.onClick = ui.rbStroke.onClick = function () {
            syncKind(ui);
            updatePreview();
        };
        bindNumberField(ui.strokeWidthInput, updatePreview);

        /* カラー / Color */
        var onColorChanged = function () {
            syncColor(ui);
            updatePreview();
        };
        for (var colorKey in ui.colorRadios) ui.colorRadios[colorKey].onClick = onColorChanged;
        for (var plateKey in ui.cmykInputs) bindNumberField(ui.cmykInputs[plateKey], updatePreview);

        /* 不透明度 / Opacity */
        ui.cbOpacityApply.onClick = function () {
            syncOpacity(ui);
            updatePreview();
        };
        bindNumberField(ui.opacityInput, updatePreview);
    }

    /**
     * エントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var selectionItems = doc.selection;
        if (!selectionItems || selectionItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var context = resolveContext(selectionItems);
        dialogState = settingsStore.load(DEFAULT_SETTINGS);
        var hadSavedState = dialogState._hasSaved;
        var previewManager = new PreviewManager();

        var ui = buildDialog();
        restoreDialogState(ui);
        syncAllPanels(ui);
        adjustDialogForTarget(ui, context);

        /**
         * プレビューを作り直し、ダイアログの値を控える
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(function () {
                hideReplacedBackdrop(context);
                createBackdrop(doc, ui, context);
            });
            saveDialogState(ui);
        }

        bindDialogEvents(ui, updatePreview);

        ui.btnTransparencyGrid.onClick = function () {
            /* 表示の切り替えなので Undo 履歴には残らない / A view toggle, so it adds no undo step */
            app.executeMenuCommand('TransparencyGrid Menu Item');
            app.redraw();
        };

        ui.btnCancel.onClick = function () {
            previewManager.rollback();
            ui.dialog.close(0);
        };

        ui.dialog.onShow = function () {
            var location = ui.dialog.location;
            ui.dialog.location = [location[0] + DIALOG_OFFSET_X, location[1] + DIALOG_OFFSET_Y];
            ui.scaleInput.active = true;
            if (!hadSavedState) {
                /* 計測に失敗しても初期値のまま開く / Keep the stock values if measuring fails */
                try {
                    applyDefaultsFromTarget(ui, context);
                } catch (e) {
                    logDebug('applyDefaultsFromTarget failed: ' + e);
                }
            }
            updatePreview();
        };

        prepareDialogWindow(ui.dialog, SCRIPT_NAME);
        var dialogResult = ui.dialog.show();
        saveDialogState(ui);
        if (dialogResult !== 1) {
            previewManager.rollback();
            return;
        }

        /* プレビューを巻き戻し、確定処理を1回だけ実行（Undo 1段）/ Roll back, then run the final action once */
        previewManager.confirm(function () {
            removeReplacedBackdrop(context);
            var backdropShape = createBackdrop(doc, ui, context);
            backdropShape = finalizeBackdropShape(doc, backdropShape, ui, context);
            var resultItem = placeBackdrop(backdropShape, context.targetItem, ui.cbGroupWithText.value);
            if (ui.cbExclude.value) resultItem = applyExclude(doc, resultItem);
            doc.selection = null;
            try {
                resultItem.selected = true;
            } catch (e) {
                logDebug('result selection failed: ' + e);
            }
        });
    }

    try {
        main();
    } catch (err) {
        alert(labelValueText("alert.genericError", err));
    }

})();
