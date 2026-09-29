#target illustrator
#targetengine "SplitBackgroundForTwoEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

2つのオブジェクトを選択して実行すると、背面に2分割の背景を作成します。
左右・上下のどちらで分けるかは位置関係から自動で判別し、大きさや分割位置はプレビューを見ながら調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitBackgroundForTwo.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1b7b8759e53b

### Overview

Creates a two-part background behind two selected objects.
Whether it splits left/right or top/bottom follows from how the objects are placed, and the size and split position are adjusted with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitBackgroundForTwo.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitBackgroundForTwo";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.10.3";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitBackgroundForTwo.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitBackgroundForTwo.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1b7b8759e53b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_SIZE_PERCENT = 200;        /* サイズ（%）の初期値 / initial size (%) */
    var DEFAULT_STROKE_PT = 1;             /* 線幅の初期値（pt）/ initial stroke width (pt) */
    var DEFAULT_CORNER_PT = 0;             /* 角丸の初期値（pt）/ initial corner radius (pt) */
    var FIRST_FILL_RGB = [220, 220, 220];  /* 左（上）側の塗り / fill of the left (top) side */
    var SECOND_FILL_RGB = [128, 128, 128]; /* 右（下）側の塗り / fill of the right (bottom) side */
    var LINE_CMYK = [0, 0, 0, 100];        /* 全体の枠と区切り線の色（K100）/ color of the frame and divider (K100) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 18;               /* ダイアログ外周の余白 / dialog margins */
    var DIALOG_OFFSET_X = 300;             /* 表示位置の横のずらし量 / horizontal dialog offset */
    var DIALOG_OFFSET_Y = 0;               /* 表示位置の縦のずらし量 / vertical dialog offset */
    var PANEL_MARGINS = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 12;               /* 2カラムの間隔 / column gutter */
    var BALANCE_SPACING = 6;               /* 幅の入力欄とスライダーの間隔 / gap between the width field and slider */
    var SIZE_INPUT_CHARS = 6;              /* サイズ・幅の入力欄の文字数 / characters for the size and width fields */
    var LENGTH_INPUT_CHARS = 4;            /* 線幅・角丸の入力欄の文字数 / characters for the stroke and corner fields */
    var BALANCE_SLIDER_SIZE = [220, 20];   /* 幅スライダーの大きさ / width slider size */

    /**
     * 共通レイアウトのパネルを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} titleText - パネルのタイトル
     * @param {string} [childAlignment] - 子の横方向の揃え（既定は "left"）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleText, childAlignment) {
        var panel = parent.add("panel", undefined, titleText);
        panel.orientation = "column";
        panel.alignChildren = [childAlignment || "left", "top"];
        panel.margins = PANEL_MARGINS;
        return panel;
    }

    /**
     * 横並びの行グループを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @returns {Group} 追加した行グループ
     */
    function addRow(parent) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} labelKey - ラベルの LABELS キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @param {boolean} checked - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parent, labelKey, tooltipKey, checked) {
        var checkbox = parent.add("checkbox", undefined, getLabel(labelKey));
        checkbox.helpTip = getLabel(tooltipKey);
        checkbox.value = !!checked;
        return checkbox;
    }

    /**
     * ツールチップ付きのラジオボタンを追加する
     * @param {Group} parent - 追加先
     * @param {string} labelKey - ラベルの LABELS キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parent, labelKey, tooltipKey) {
        var radio = parent.add("radiobutton", undefined, getLabel(labelKey));
        radio.helpTip = getLabel(tooltipKey);
        return radio;
    }

    /**
     * ∧∨と入力欄をひと組で追加する（隙間0で突き合わせ、↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} rowGroup - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {number} characters - 入力欄の文字数
     * @param {Object} stepOptions - addStepper() に渡す増減の設定。onStep は bindDialogEvents() で入れる
     * @returns {EditText} 入力欄（増減の設定は .stepOptions で参照できる）
     */
    function addStepperInput(rowGroup, initialText, characters, stepOptions) {
        var stepperInputGroup = rowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = characters;
        numberInput.stepOptions = stepOptions;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 長さの入力欄と単位ラベルを追加する
     * @param {Group} rowGroup - 追加先の行
     * @param {number} valuePt - 初期値（pt）
     * @param {string} prefKey - 単位の環境設定キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @returns {EditText} 追加した入力欄
     */
    function addLengthInput(rowGroup, valuePt, prefKey, tooltipKey) {
        var lengthInput = addStepperInput(rowGroup, formatUnitValue(ptToUnit(valuePt, prefKey), prefKey), LENGTH_INPUT_CHARS, { step: 1, min: 0 });
        lengthInput.helpTip = getLabel(tooltipKey);
        rowGroup.add("statictext", undefined, getUnitInfo(prefKey).label);
        return lengthInput;
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

    // =========================================
    // 単位 / Units
    // =========================================

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

    /**
     * 環境設定の単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} pt の値
     */
    function unitToPt(value, prefKey) {
        return value * getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * pt の値を環境設定の単位に換算する
     * @param {number} valuePt - pt の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} 単位の値
     */
    function ptToUnit(valuePt, prefKey) {
        return valuePt / getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * 単位の大きさに合わせて丸める（pt・Q は小数1桁、mm は2桁、in・cm・ft など1単位が10pt以上は3桁）
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} 丸めた値
     */
    function roundUnitValue(value, prefKey) {
        var pointsPerUnit = getUnitInfo(prefKey).pointsPerUnit;
        var scale = (pointsPerUnit >= 10) ? 1000 : (pointsPerUnit >= 2) ? 100 : 10;
        return Math.round(value * scale) / scale;
    }

    /**
     * 単位の値を入力欄の文字列にする
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {string} 表示用の文字列
     */
    function formatUnitValue(value, prefKey) {
        return String(roundUnitValue(value, prefKey));
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "2つのオブジェクトの背景を作成", en: "Create Background for Two Objects" }
        },
        panel: {
            draw: { ja: "描画", en: "描画" },
            options: { ja: "オプション", en: "Options" },
            balance: { ja: "バランス", en: "Balance" }
        },
        fieldLabel: {
            height: { ja: "高さ", en: "Height" },
            width: { ja: "幅", en: "Width" },
            balanceWidth: { ja: "幅", en: "Width" },
            strokeWidth: { ja: "線幅", en: "Stroke width" },
            cornerRadius: { ja: "角丸", en: "Corner radius" }
        },
        checkbox: {
            fillLeft: { ja: "塗り（左）", en: "Fill (Left)" },
            fillRight: { ja: "塗り（右）", en: "Fill (Right)" },
            fillTop: { ja: "塗り（上）", en: "Fill (Top)" },
            fillBottom: { ja: "塗り（下）", en: "Fill (Bottom)" },
            overallFrame: { ja: "全体の枠", en: "Overall frame" },
            divider: { ja: "区切り線", en: "Divider" }
        },
        radio: {
            none: { ja: "なし", en: "None" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" },
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" }
        },
        tooltip: {
            sizePercent: { ja: "分割の割合（％）です。", en: "Split ratio, in percent." },
            fillSide: { ja: "この側に塗りを付けます。", en: "Fills this side." },
            overallFrame: {
                ja: "分割した全体を1つの枠線で囲みます。",
                en: "Draws a single frame around the whole split shape."
            },
            divider: { ja: "分割の境目にケイ線を引きます。", en: "Draws a rule along the split." },
            strokeWidth: { ja: "ケイ線の太さです。", en: "Weight of the rules." },
            cornerRadius: {
                ja: "角丸の半径です。0 で角のままになります。",
                en: "Corner radius. 0 leaves the corners square."
            },
            pinNone: {
                ja: "どちらの幅も固定せず、割合だけで分けます。",
                en: "Pins neither side; the ratio alone decides the split."
            },
            pinLeft: {
                ja: "左側の幅を固定し、残りを右側にします。",
                en: "Pins the width of the left side and gives the rest to the right."
            },
            pinRight: {
                ja: "右側の幅を固定し、残りを左側にします。",
                en: "Pins the width of the right side and gives the rest to the left."
            },
            pinTop: {
                ja: "上側の幅を固定し、残りを下側にします。",
                en: "Pins the width of the top side and gives the rest to the bottom."
            },
            pinBottom: {
                ja: "下側の幅を固定し、残りを上側にします。",
                en: "Pins the width of the bottom side and gives the rest to the top."
            },
            balanceWidth: { ja: "固定する側の幅です。", en: "Width of the pinned side." },
            balanceWidthSlider: {
                ja: "固定する側の幅をドラッグで決めます。",
                en: "Drag to set the width of the pinned side."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            openDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectTwoItems: { ja: "2つのオブジェクトを選択してください。", en: "Please select two objects." },
            sizeInvalid: {
                ja: "サイズ（%）は0より大きい数値を入力してください。",
                en: "Enter a size (%) greater than 0."
            }
        }
    };

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
            return textFile.read().replace(/^\uFEFF/, "");
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
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
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

    // =========================================
    // セッション設定 / Session settings
    // =========================================
    // Illustrator の起動中だけダイアログの値を保持する（長さは単位を変えても崩れないよう pt で持つ）
    // Dialog values are kept while Illustrator is running; lengths are stored in pt so unit changes do not skew them

    /* 旧版の $.global のキーを1度だけ読み継ぐ / The old $.global key is read once */
    var LEGACY_SESSION_KEY = "SplitBackgroundForTwo_settings";
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_KEY] || null; }
    });

    /**
     * セッションに残したダイアログの値を返す（欠けている項目は初期値で補う）
     * @returns {Object} セッション設定（書き換えても保存されない。saveDialogToSession() で保存する）
     */
    function getSessionSettings() {
        var defaults = {
            percent: DEFAULT_SIZE_PERCENT,
            fillFirst: true,
            fillSecond: true,
            overallFrame: false,
            divider: false,
            strokeWidthPt: DEFAULT_STROKE_PT,
            cornerRadiusPt: DEFAULT_CORNER_PT,
            balanceMode: "none",
            balanceWidthPt: 0
        };
        return settingsStore.load(defaults);
    }

    // =========================================
    // ダイアログ共通 / Dialog helpers
    // =========================================

    /**
     * ダイアログの表示位置をずらす（既存の onShow は先に呼ぶ）
     * @param {Window} dialog - 対象のダイアログ
     * @param {number} offsetX - 横のずらし量
     * @param {number} offsetY - 縦のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(dialog, offsetX, offsetY) {
        var previousOnShow = dialog.onShow;
        dialog.onShow = function () {
            if (typeof previousOnShow === "function") previousOnShow();
            dialog.location = [dialog.location[0] + offsetX, dialog.location[1] + offsetY];
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
    // 実行時の状態 / Run state
    // =========================================

    var doc = null;               /* 対象ドキュメント / target document */
    var targetItems = [];         /* 選択した2つのオブジェクト（左または上が先）/ the two selected objects, left or top first */
    var targetBounds = [];        /* targetItems の外接矩形 [左, 上, 右, 下] / bounds of targetItems [left, top, right, bottom] */
    var isVerticalSplit = false;  /* 上下に分けるか / whether the split runs top/bottom */

    /* プレビュー専用レイヤー（ダイアログを閉じると消す）/ Dedicated preview layer, removed when the dialog closes */
    var PREVIEW_LAYER_NAME = "__SplitBackgroundForTwo__PreviewLayer__";

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
    // 計測 / Measuring
    // =========================================

    /**
     * テキストを複製してアウトライン化し、その外接矩形を返す（サイドベアリングを含めない）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[]|null} 外接矩形。測れなければ null
     */
    function getOutlineBounds(textFrame) {
        var duplicateText = null;
        var outlineGroup = null;
        var bounds = null;
        try {
            duplicateText = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATBEGINNING);
            outlineGroup = duplicateText.createOutline();
            bounds = outlineGroup.geometricBounds;
        } catch (e) {
            /* 空白だけのテキストは中身の無いグループになり、geometricBounds が例外になる / whitespace-only text outlines to an empty group whose bounds throw */
            bounds = null;
        } finally {
            removeItem(outlineGroup);
            /* createOutline() が成功していれば複製は消費済み / the duplicate is already consumed when createOutline() succeeds */
            removeItem(duplicateText);
        }
        return bounds;
    }

    /**
     * オブジェクトの外接矩形を返す。テキストはアウトライン、クリップグループはマスクの範囲で測る
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]} 外接矩形 [左, 上, 右, 下]
     */
    function getItemBounds(item) {
        /* クリップグループはマスクを測る / Measure the mask of a clip group */
        var measuredItem = getClipMaskItem(item) || item;
        var bounds = null;
        if (measuredItem.typename === "TextFrame") bounds = getOutlineBounds(measuredItem);
        if (!bounds) bounds = getClipAwareBounds(measuredItem, false);
        if (!bounds) bounds = measuredItem.geometricBounds;
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    }

    /**
     * 2つの外接矩形のすき間を返す（重なっていれば 0）
     * @param {Array} boundsPair - 2つの外接矩形
     * @param {boolean} vertical - 上下のすき間を測るか（false なら左右）
     * @returns {number} すき間（pt）
     */
    function getGapPt(boundsPair, vertical) {
        var bounds1 = boundsPair[0];
        var bounds2 = boundsPair[1];
        var gapPt;
        if (vertical) {
            gapPt = (bounds1[1] >= bounds2[1]) ? (bounds1[3] - bounds2[1]) : (bounds2[3] - bounds1[1]);
        } else {
            gapPt = (bounds1[0] <= bounds2[0]) ? (bounds2[0] - bounds1[2]) : (bounds1[0] - bounds2[2]);
        }
        return (gapPt > 0) ? gapPt : 0;
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * レイヤーの重ね順を、最上位からの添字の並びで返す（添字が大きいほど背面）
     * @param {Layer} layer - 対象のレイヤー
     * @returns {number[]} 最上位レイヤーから順の添字
     */
    function getLayerStackPath(layer) {
        var stackPath = [];
        while (layer.typename === "Layer") {
            var siblingLayers = layer.parent.layers;
            for (var i = 0; i < siblingLayers.length; i++) {
                if (siblingLayers[i] === layer) break;
            }
            stackPath.unshift(i);
            layer = layer.parent;
        }
        return stackPath;
    }

    /**
     * 2つのレイヤーのうち背面側を返す（サブレイヤーも比べる）
     * @param {Layer} layerA - 1つ目のレイヤー
     * @param {Layer} layerB - 2つ目のレイヤー
     * @returns {Layer} 背面側のレイヤー
     */
    function getBackmostLayer(layerA, layerB) {
        var pathA = getLayerStackPath(layerA);
        var pathB = getLayerStackPath(layerB);
        for (var i = 0; i < pathA.length && i < pathB.length; i++) {
            if (pathA[i] !== pathB[i]) return (pathA[i] > pathB[i]) ? layerA : layerB;
        }
        /* 同じレイヤーか親子なら、浅いほう（親）に描く / same layer or parent and child: use the shallower one */
        return (pathA.length <= pathB.length) ? layerA : layerB;
    }

    /**
     * サブレイヤーをたどって最上位のレイヤーを返す
     * @param {Layer} layer - 対象のレイヤー
     * @returns {Layer} 最上位のレイヤー
     */
    function getTopLevelLayer(layer) {
        while (layer.parent.typename === "Layer") layer = layer.parent;
        return layer;
    }

    /**
     * プレビューレイヤーを探す
     * @returns {Layer|null} プレビューレイヤー。無ければ null
     */
    function findPreviewLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === PREVIEW_LAYER_NAME) return doc.layers[i];
        }
        return null;
    }

    /**
     * プレビューレイヤーを用意し、2つのオブジェクトのレイヤーより背面に置く
     * @returns {Layer} プレビューレイヤー
     */
    function ensurePreviewLayer() {
        var previewLayer = findPreviewLayer();
        if (previewLayer) return previewLayer;

        previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;
        /* 背面側のレイヤーの、さらに背面へ（サブレイヤーなら最上位の親の背面へ）/ behind the backmost layer, or behind its top-level parent for a sublayer */
        var backLayer = getTopLevelLayer(getBackmostLayer(targetItems[0].layer, targetItems[1].layer));
        previewLayer.move(backLayer, ElementPlacement.PLACEAFTER);
        return previewLayer;
    }

    /**
     * プレビューレイヤーを中身ごと削除する
     * @returns {void}
     */
    function removePreviewLayer() {
        var previewLayer = findPreviewLayer();
        if (previewLayer) previewLayer.remove();
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * RGB の配列から色を作る
     * @param {number[]} rgbValues - [R, G, B]
     * @returns {RGBColor} 色
     */
    function makeRGBColor(rgbValues) {
        var color = new RGBColor();
        color.red = rgbValues[0];
        color.green = rgbValues[1];
        color.blue = rgbValues[2];
        return color;
    }

    /**
     * CMYK の配列から色を作る
     * @param {number[]} cmykValues - [C, M, Y, K]
     * @returns {CMYKColor} 色
     */
    function makeCMYKColor(cmykValues) {
        var color = new CMYKColor();
        color.cyan = cmykValues[0];
        color.magenta = cmykValues[1];
        color.yellow = cmykValues[2];
        color.black = cmykValues[3];
        return color;
    }

    var FIRST_FILL_COLOR = makeRGBColor(FIRST_FILL_RGB);
    var SECOND_FILL_COLOR = makeRGBColor(SECOND_FILL_RGB);
    var LINE_COLOR = makeCMYKColor(LINE_CMYK);

    /**
     * 削除する（createOutline() に消費された複製など、すでに無いものは無視する）
     * @param {PageItem} item - 削除するオブジェクト
     * @returns {void}
     */
    function removeItem(item) {
        if (!item) return;
        try {
            item.remove();
        } catch (e) {
            /* すでに無い / already gone */
        }
    }

    /**
     * すき間を1つ目側と2つ目側の余白に振り分ける
     * @param {number} gapPt - すき間（pt）
     * @param {string} balanceMode - 'none' | 'first' | 'second'
     * @param {number} balanceWidthPt - 固定する側の幅（pt）
     * @returns {number[]} [1つ目側の余白, 2つ目側の余白]
     */
    function splitGap(gapPt, balanceMode, balanceWidthPt) {
        var pinnedPt = Math.min(Math.max(balanceWidthPt, 0), gapPt);
        if (balanceMode === "first") return [pinnedPt, gapPt - pinnedPt];
        if (balanceMode === "second") return [gapPt - pinnedPt, pinnedPt];
        return [gapPt / 2, gapPt / 2];
    }

    /**
     * 背景全体の範囲と分割位置を求める
     * @param {Object} drawOptions - 描画の設定（collectDrawOptions() の戻り値）
     * @returns {{left: number, top: number, right: number, bottom: number, split: number}} 範囲と分割位置（左右なら X、上下なら Y）
     */
    function computeSplitLayout(drawOptions) {
        var margins = splitGap(getGapPt(targetBounds, isVerticalSplit), drawOptions.balanceMode, drawOptions.balanceWidthPt);
        return isVerticalSplit
            ? computeTopBottomLayout(margins, drawOptions.sizeRatio)
            : computeLeftRightLayout(margins, drawOptions.sizeRatio);
    }

    /**
     * 左右分割の範囲：高さを倍率で広げ、横はすき間の余白ぶん外へ伸ばす
     * @param {number[]} margins - [左側の余白, 右側の余白]
     * @param {number} sizeRatio - 高さの倍率
     * @returns {{left: number, top: number, right: number, bottom: number, split: number}} 範囲と分割位置（X）
     */
    function computeLeftRightLayout(margins, sizeRatio) {
        var leftBounds = targetBounds[0];
        var rightBounds = targetBounds[1];
        var contentTop = Math.max(leftBounds[1], rightBounds[1]);
        var contentHeight = contentTop - Math.min(leftBounds[3], rightBounds[3]);
        var rectHeight = contentHeight * sizeRatio;
        var rectTop = contentTop - (contentHeight / 2) + (rectHeight / 2);
        return {
            left: leftBounds[0] - margins[0],
            top: rectTop,
            right: rightBounds[2] + margins[1],
            bottom: rectTop - rectHeight,
            split: leftBounds[2] + margins[0]
        };
    }

    /**
     * 上下分割の範囲：横幅を倍率で広げ、縦はすき間の余白ぶん外へ伸ばす
     * @param {number[]} margins - [上側の余白, 下側の余白]
     * @param {number} sizeRatio - 横幅の倍率
     * @returns {{left: number, top: number, right: number, bottom: number, split: number}} 範囲と分割位置（Y）
     */
    function computeTopBottomLayout(margins, sizeRatio) {
        var topBounds = targetBounds[0];
        var bottomBounds = targetBounds[1];
        var contentLeft = Math.min(topBounds[0], bottomBounds[0]);
        var contentWidth = Math.max(Math.max(topBounds[2], bottomBounds[2]) - contentLeft, 0);
        var rectWidth = contentWidth * sizeRatio;
        var rectLeft = contentLeft + (contentWidth / 2) - (rectWidth / 2);
        var rectTop = topBounds[1] + margins[0];
        return {
            left: rectLeft,
            top: rectTop,
            right: rectLeft + rectWidth,
            bottom: Math.min(bottomBounds[3] - margins[1], rectTop),
            split: topBounds[3] - margins[0]
        };
    }

    /**
     * パスのアンカーの選択をすべて外す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function deselectAnchors(pathItem) {
        var points = pathItem.pathPoints;
        for (var i = 0; i < points.length; i++) {
            points[i].selected = PathPointSelection.NOSELECTION;
        }
    }

    /**
     * 長方形の、指定した辺の両端のアンカーだけを選択する
     * @param {PathItem} rectPath - 4点の長方形
     * @param {string} side - "left" | "right" | "top" | "bottom"
     * @returns {boolean} 選択できたか
     */
    function selectRectangleCorners(rectPath, side) {
        var points = rectPath.pathPoints;
        if (points.length !== 4) return false;

        var axis = (side === "left" || side === "right") ? 0 : 1;
        var order = [0, 1, 2, 3];
        order.sort(function (a, b) { return points[a].anchor[axis] - points[b].anchor[axis]; });
        /* 右と上は座標の大きい2点、左と下は小さい2点 / right and top take the two larger coordinates, left and bottom the two smaller */
        var picked = (side === "right" || side === "top") ? [order[2], order[3]] : [order[0], order[1]];

        deselectAnchors(rectPath);
        points[picked[0]].selected = PathPointSelection.ANCHORPOINT;
        points[picked[1]].selected = PathPointSelection.ANCHORPOINT;
        return true;
    }

    /**
     * 塗りの長方形の外側の2つの角を丸める
     * @param {PathItem} fillRect - 塗りの長方形
     * @param {string} side - 丸める辺（"left" | "right" | "top" | "bottom"）
     * @param {number} cornerRadiusPt - 角丸の半径（pt）
     * @returns {void}
     */
    function roundFillCorners(fillRect, side, cornerRadiusPt) {
        if (!(cornerRadiusPt > 0)) return;
        if (!selectRectangleCorners(fillRect, side)) return;
        try {
            roundAnyCorner([fillRect], { rr: cornerRadiusPt });
        } catch (e) {
            /* 幅0の長方形などで計算に失敗しても、角のまま描く / keep square corners if rounding fails, e.g. on a zero-width rectangle */
        }
        deselectAnchors(fillRect);
    }

    /**
     * 「角を丸くする」効果を適用する
     * @param {PathItem} targetItem - 対象のパス
     * @param {number} cornerRadiusPt - 角丸の半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(targetItem, cornerRadiusPt) {
        if (!(cornerRadiusPt > 0)) return;
        try {
            targetItem.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerRadiusPt + ' "/></LiveEffect>');
        } catch (e) {
            /* 効果を適用できない環境では角のまま描く / keep square corners where the effect cannot be applied */
        }
    }

    /**
     * 線を K100 のケイ線にする
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function styleLine(pathItem, strokeWidthPt) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeWidth = strokeWidthPt;
        pathItem.strokeColor = LINE_COLOR;
    }

    /**
     * 塗りの長方形を追加し、外側の角を丸める
     * @param {PathItems} pathItems - 追加先
     * @param {number[]} rectBounds - [左, 上, 右, 下]
     * @param {RGBColor} fillColor - 塗りの色
     * @param {string} roundSide - 丸める辺
     * @param {number} cornerRadiusPt - 角丸の半径（pt）
     * @returns {PathItem} 追加した長方形
     */
    function addFillRect(pathItems, rectBounds, fillColor, roundSide, cornerRadiusPt) {
        var fillRect = pathItems.rectangle(rectBounds[1], rectBounds[0], rectBounds[2] - rectBounds[0], rectBounds[1] - rectBounds[3]);
        fillRect.fillColor = fillColor;
        fillRect.stroked = false;
        roundFillCorners(fillRect, roundSide, cornerRadiusPt);
        return fillRect;
    }

    /**
     * 背景の塗り・全体の枠・区切り線を描く
     * @param {Object} drawOptions - 描画の設定（collectDrawOptions() の戻り値）
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @returns {PageItem[]} 描いたオブジェクト
     */
    function drawSplitBackground(drawOptions, targetLayer) {
        var layout = computeSplitLayout(drawOptions);
        var pathItems = targetLayer.pathItems;
        var fillRects = [];
        var drawnItems = [];

        if (drawOptions.fillFirst) {
            fillRects.push(addFillRect(pathItems,
                isVerticalSplit ? [layout.left, layout.top, layout.right, layout.split] : [layout.left, layout.top, layout.split, layout.bottom],
                FIRST_FILL_COLOR, isVerticalSplit ? "top" : "left", drawOptions.cornerRadiusPt));
        }
        if (drawOptions.fillSecond) {
            fillRects.push(addFillRect(pathItems,
                isVerticalSplit ? [layout.left, layout.split, layout.right, layout.bottom] : [layout.split, layout.top, layout.right, layout.bottom],
                SECOND_FILL_COLOR, isVerticalSplit ? "bottom" : "right", drawOptions.cornerRadiusPt));
        }

        if (drawOptions.overallFrame) {
            var frameRect = pathItems.rectangle(layout.top, layout.left, layout.right - layout.left, layout.top - layout.bottom);
            styleLine(frameRect, drawOptions.strokeWidthPt);
            applyRoundCornersEffect(frameRect, drawOptions.cornerRadiusPt);
            frameRect.zOrder(ZOrderMethod.BRINGTOFRONT);
            drawnItems.push(frameRect);
        }

        if (drawOptions.divider) {
            var dividerLine = pathItems.add();
            styleLine(dividerLine, drawOptions.strokeWidthPt);
            dividerLine.setEntirePath(isVerticalSplit
                ? [[layout.left, layout.split], [layout.right, layout.split]]
                : [[layout.split, layout.top], [layout.split, layout.bottom]]);
            dividerLine.zOrder(ZOrderMethod.BRINGTOFRONT);
            drawnItems.push(dividerLine);
        }

        /* 塗りは最背面へ（1つ目が2つ目の上に来る順）/ send the fills to the back, keeping the first above the second */
        for (var i = 0; i < fillRects.length; i++) {
            fillRects[i].zOrder(ZOrderMethod.SENDTOBACK);
            drawnItems.push(fillRects[i]);
        }
        return drawnItems;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューレイヤーの中身を消す
     * @returns {void}
     */
    function clearPreview() {
        var previewLayer = findPreviewLayer();
        if (!previewLayer) return;
        for (var i = previewLayer.pageItems.length - 1; i >= 0; i--) {
            previewLayer.pageItems[i].remove();
        }
    }

    /**
     * プレビューを描き直す
     * @param {Object|null} drawOptions - 描画の設定。null ならプレビューを消すだけ
     * @returns {void}
     */
    function drawPreview(drawOptions) {
        clearPreview();
        if (drawOptions) {
            try {
                drawSplitBackground(drawOptions, ensurePreviewLayer());
            } catch (e) {
                /* 描けなくてもダイアログの入力は続けられるようにする / keep the dialog usable even if drawing fails */
                clearPreview();
            }
            /* 角丸の処理が塗りを選択するので外す / the corner rounding selects the fill; deselect it */
            doc.selection = null;
        }
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを組み立てる（イベントは bindDialogEvents() で付ける）
     * @param {Object} session - セッションに残した前回の値
     * @returns {Object} ダイアログとコントロールの参照
     */
    function buildSettingsDialog(session) {
        var controls = {};
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = DIALOG_MARGINS;
        controls.dialog = dialog;

        /* サイズ（%）：左右なら高さ、上下なら幅 / Size (%): height for a left/right split, width for top/bottom */
        var sizeRowGroup = addRow(dialog);
        sizeRowGroup.alignment = ["center", "top"];
        sizeRowGroup.add("statictext", undefined, labelText(isVerticalSplit ? "fieldLabel.width" : "fieldLabel.height"));
        controls.sizeInput = addStepperInput(sizeRowGroup, String(session.percent), SIZE_INPUT_CHARS, { step: 1, min: 0 });
        controls.sizeInput.helpTip = getLabel("tooltip.sizePercent");
        controls.sizeInput.active = true;
        sizeRowGroup.add("statictext", undefined, "%");

        /* 2カラム：左＝描画、右＝オプション / Two columns: drawing on the left, options on the right */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var drawPanel = addPanel(columnsGroup, getLabel("panel.draw"));
        controls.fillFirstCheckbox = addCheckbox(drawPanel, isVerticalSplit ? "checkbox.fillTop" : "checkbox.fillLeft", "tooltip.fillSide", session.fillFirst);
        controls.fillSecondCheckbox = addCheckbox(drawPanel, isVerticalSplit ? "checkbox.fillBottom" : "checkbox.fillRight", "tooltip.fillSide", session.fillSecond);
        controls.overallFrameCheckbox = addCheckbox(drawPanel, "checkbox.overallFrame", "tooltip.overallFrame", session.overallFrame);
        controls.dividerCheckbox = addCheckbox(drawPanel, "checkbox.divider", "tooltip.divider", session.divider);

        var optionsPanel = addPanel(columnsGroup, getLabel("panel.options"));
        controls.strokeRowGroup = addRow(optionsPanel);
        controls.strokeRowGroup.add("statictext", undefined, getLabel("fieldLabel.strokeWidth"));
        controls.strokeInput = addLengthInput(controls.strokeRowGroup, session.strokeWidthPt, "strokeUnits", "tooltip.strokeWidth");
        controls.cornerRowGroup = addRow(optionsPanel);
        controls.cornerRowGroup.add("statictext", undefined, getLabel("fieldLabel.cornerRadius"));
        controls.cornerInput = addLengthInput(controls.cornerRowGroup, session.cornerRadiusPt, "rulerType", "tooltip.cornerRadius");

        buildBalancePanel(dialog, controls, session);

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        controls.btnCancel = btnCancel;
        controls.btnOK = btnOK;
        dialog.defaultElement = controls.btnOK;
        dialog.cancelElement = controls.btnCancel;

        return controls;
    }

    /**
     * ［バランス］パネル（固定する側と、その幅）を組み立てる
     * @param {Window} dialog - 追加先のダイアログ
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @param {Object} session - セッションに残した前回の値
     * @returns {void}
     */
    function buildBalancePanel(dialog, controls, session) {
        var balancePanel = addPanel(dialog, getLabel("panel.balance"), "fill");
        balancePanel.alignment = ["fill", "top"];

        var balanceRadioGroup = addRow(balancePanel);
        controls.pinNoneRadio = addRadio(balanceRadioGroup, "radio.none", "tooltip.pinNone");
        controls.pinFirstRadio = addRadio(balanceRadioGroup, isVerticalSplit ? "radio.top" : "radio.left", isVerticalSplit ? "tooltip.pinTop" : "tooltip.pinLeft");
        controls.pinSecondRadio = addRadio(balanceRadioGroup, isVerticalSplit ? "radio.bottom" : "radio.right", isVerticalSplit ? "tooltip.pinBottom" : "tooltip.pinRight");
        controls.pinFirstRadio.value = (session.balanceMode === "first");
        controls.pinSecondRadio.value = (session.balanceMode === "second");
        controls.pinNoneRadio.value = !controls.pinFirstRadio.value && !controls.pinSecondRadio.value;

        controls.balanceWidthGroup = balancePanel.add("group");
        controls.balanceWidthGroup.orientation = "column";
        controls.balanceWidthGroup.alignChildren = ["fill", "top"];
        controls.balanceWidthGroup.spacing = BALANCE_SPACING;

        var balanceInputRow = addRow(controls.balanceWidthGroup);
        balanceInputRow.add("statictext", undefined, getLabel("fieldLabel.balanceWidth"));
        controls.balanceWidthInput = addStepperInput(balanceInputRow, "", SIZE_INPUT_CHARS, { step: 1, min: 0 });
        controls.balanceWidthInput.helpTip = getLabel("tooltip.balanceWidth");
        balanceInputRow.add("statictext", undefined, getUnitInfo("rulerType").label);

        /* 幅の上限は2つのオブジェクトのすき間 / the width cannot exceed the gap between the objects */
        controls.balanceWidthMax = roundUnitValue(ptToUnit(getGapPt(targetBounds, isVerticalSplit), "rulerType"), "rulerType");
        controls.balanceWidthInput.stepOptions.max = controls.balanceWidthMax;
        var balanceSliderRow = addRow(controls.balanceWidthGroup);
        balanceSliderRow.alignChildren = ["fill", "center"];
        controls.balanceWidthSlider = balanceSliderRow.add("slider", undefined, 0, 0, controls.balanceWidthMax);
        controls.balanceWidthSlider.helpTip = getLabel("tooltip.balanceWidthSlider");
        controls.balanceWidthSlider.preferredSize = BALANCE_SLIDER_SIZE;

        setBalanceWidth(controls, clampBalanceWidth(controls, ptToUnit(session.balanceWidthPt, "rulerType"), false));
    }

    /**
     * 線幅の欄を使うか（全体の枠か区切り線があるとき）
     * @param {Object} controls - コントロールの参照
     * @returns {boolean} 使うか
     */
    function needsStroke(controls) {
        return controls.overallFrameCheckbox.value || controls.dividerCheckbox.value;
    }

    /**
     * 角丸の欄を使うか（塗りか全体の枠があるとき）
     * @param {Object} controls - コントロールの参照
     * @returns {boolean} 使うか
     */
    function needsCorner(controls) {
        return controls.fillFirstCheckbox.value || controls.fillSecondCheckbox.value || controls.overallFrameCheckbox.value;
    }

    /**
     * 使わない欄を無効にする
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function updateEnabledStates(controls) {
        controls.strokeRowGroup.enabled = needsStroke(controls);
        controls.cornerRowGroup.enabled = needsCorner(controls);
        controls.balanceWidthGroup.enabled = !controls.pinNoneRadio.value;
        /* ∧∨は自作描画なので描き直してディム表示を切り替える / redraw the custom-drawn steppers */
        redrawSteppersIn(controls.strokeRowGroup);
        redrawSteppersIn(controls.cornerRowGroup);
        redrawSteppersIn(controls.balanceWidthGroup);
    }

    /**
     * サイズ（%）を読む
     * @param {Object} controls - コントロールの参照
     * @returns {number|null} サイズ（%）。0以下や数値でなければ null
     */
    function readSizePercent(controls) {
        var sizePercent = Number(controls.sizeInput.text);
        return (isNaN(sizePercent) || sizePercent <= 0) ? null : sizePercent;
    }

    /**
     * 長さの入力欄を pt に換算して読む
     * @param {EditText} lengthInput - 入力欄
     * @param {string} prefKey - 単位の環境設定キー
     * @param {boolean} allowZero - 0 を許すか
     * @returns {number|null} pt の値（小数1桁）。不正な値なら null
     */
    function readLengthPt(lengthInput, prefKey, allowZero) {
        var valueUnit = Number(lengthInput.text);
        if (isNaN(valueUnit) || valueUnit < 0) return null;
        var valuePt = Math.round(unitToPt(roundUnitValue(valueUnit, prefKey), prefKey) * 10) / 10;
        return (valuePt > 0 || allowZero) ? valuePt : null;
    }

    /**
     * バランスの幅を 0〜上限に収める
     * @param {Object} controls - コントロールの参照
     * @param {number|string} value - 幅（定規の単位）
     * @param {boolean} forceInteger - 整数に丸めるか
     * @returns {number} 収めた幅
     */
    function clampBalanceWidth(controls, value, forceInteger) {
        var widthUnit = Number(value);
        if (isNaN(widthUnit) || widthUnit < 0) widthUnit = 0;
        if (widthUnit > controls.balanceWidthMax) widthUnit = controls.balanceWidthMax;
        return forceInteger ? Math.round(widthUnit) : roundUnitValue(widthUnit, "rulerType");
    }

    /**
     * バランスの幅を入力欄とスライダーの両方に入れる
     * @param {Object} controls - コントロールの参照
     * @param {number} widthUnit - 幅（定規の単位）
     * @returns {void}
     */
    function setBalanceWidth(controls, widthUnit) {
        controls.balanceWidthSlider.value = widthUnit;
        controls.balanceWidthInput.text = String(widthUnit);
    }

    /**
     * バランスの幅を pt で読む
     * @param {Object} controls - コントロールの参照
     * @returns {number} 幅（pt、小数1桁）
     */
    function readBalanceWidthPt(controls) {
        var widthUnit = clampBalanceWidth(controls, controls.balanceWidthInput.text, false);
        return Math.round(unitToPt(widthUnit, "rulerType") * 10) / 10;
    }

    /**
     * 固定する側を返す
     * @param {Object} controls - コントロールの参照
     * @returns {string} 'none' | 'first' | 'second'
     */
    function getBalanceMode(controls) {
        if (controls.pinFirstRadio.value) return "first";
        if (controls.pinSecondRadio.value) return "second";
        return "none";
    }

    /**
     * ダイアログの値を描画の設定にまとめる
     * @param {Object} controls - コントロールの参照
     * @returns {Object|null} 描画の設定。使う欄に不正な値があれば null
     */
    function collectDrawOptions(controls) {
        var sizePercent = readSizePercent(controls);
        var strokeWidthPt = readLengthPt(controls.strokeInput, "strokeUnits", false);
        var cornerRadiusPt = readLengthPt(controls.cornerInput, "rulerType", true);
        if (sizePercent === null) return null;

        /* 使わない欄は、不正な値でも既定値で進める / a field that is not in use falls back to its default */
        if (strokeWidthPt === null) {
            if (needsStroke(controls)) return null;
            strokeWidthPt = DEFAULT_STROKE_PT;
        }
        if (cornerRadiusPt === null) {
            if (needsCorner(controls)) return null;
            cornerRadiusPt = DEFAULT_CORNER_PT;
        }

        return {
            sizeRatio: sizePercent / 100,
            fillFirst: controls.fillFirstCheckbox.value,
            fillSecond: controls.fillSecondCheckbox.value,
            overallFrame: controls.overallFrameCheckbox.value,
            divider: controls.dividerCheckbox.value,
            strokeWidthPt: strokeWidthPt,
            cornerRadiusPt: cornerRadiusPt,
            balanceMode: getBalanceMode(controls),
            balanceWidthPt: readBalanceWidthPt(controls)
        };
    }

    /**
     * ダイアログの値でプレビューを描き直す（使う欄に不正な値があれば消すだけ）
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function refreshPreview(controls) {
        drawPreview(collectDrawOptions(controls));
    }

    /**
     * ダイアログのイベントを付ける
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function bindDialogEvents(controls) {
        var refresh = function () { refreshPreview(controls); };
        var onOptionClick = function () {
            updateEnabledStates(controls);
            refresh();
        };
        var syncBalanceWidthFromInput = function () {
            setBalanceWidth(controls, clampBalanceWidth(controls, controls.balanceWidthInput.text, false));
            refresh();
        };

        /* ∧∨・↑↓キーで増減したあとの更新 / refresh after stepping with the stepper or arrow keys */
        controls.sizeInput.stepOptions.onStep = refresh;
        controls.strokeInput.stepOptions.onStep = refresh;
        controls.cornerInput.stepOptions.onStep = refresh;
        controls.balanceWidthInput.stepOptions.onStep = syncBalanceWidthFromInput;

        controls.sizeInput.addEventListener("changing", refresh);
        controls.strokeInput.addEventListener("changing", refresh);
        controls.cornerInput.addEventListener("changing", refresh);

        /* 入力中は欄を書き換えない（小数点を打てるように）。確定したら範囲内に収める
           Leave the field alone while typing so a decimal point can be entered; clamp it on commit */
        controls.balanceWidthInput.addEventListener("changing", function () {
            controls.balanceWidthSlider.value = clampBalanceWidth(controls, controls.balanceWidthInput.text, false);
            refresh();
        });
        controls.balanceWidthInput.onChange = syncBalanceWidthFromInput;
        controls.balanceWidthSlider.onChanging = function () {
            /* option を押しながらドラッグすると整数に丸める / hold Option while dragging to snap to whole units */
            var snapToInteger = ScriptUI.environment.keyboardState.altKey;
            setBalanceWidth(controls, clampBalanceWidth(controls, controls.balanceWidthSlider.value, snapToInteger));
            refresh();
        };

        controls.fillFirstCheckbox.onClick = onOptionClick;
        controls.fillSecondCheckbox.onClick = onOptionClick;
        controls.overallFrameCheckbox.onClick = onOptionClick;
        controls.dividerCheckbox.onClick = onOptionClick;
        controls.pinNoneRadio.onClick = onOptionClick;
        controls.pinFirstRadio.onClick = onOptionClick;
        controls.pinSecondRadio.onClick = onOptionClick;

        controls.btnOK.onClick = function () {
            if (readSizePercent(controls) === null) {
                alert(getLabel("alert.sizeInvalid"));
                return;
            }
            if (!collectDrawOptions(controls)) {
                /* 線幅か角丸が不正なときは、その欄に戻す / return to the invalid stroke or corner field */
                var strokeInvalid = needsStroke(controls) && readLengthPt(controls.strokeInput, "strokeUnits", false) === null;
                (strokeInvalid ? controls.strokeInput : controls.cornerInput).active = true;
                return;
            }
            controls.dialog.close(1);
        };
        controls.btnCancel.onClick = function () { controls.dialog.close(0); };

        controls.dialog.onShow = refresh;
        shiftDialogPosition(controls.dialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        updateEnabledStates(controls);
    }

    /**
     * ダイアログの値をセッションに残して保存する（不正な値の欄は前回の値のまま）
     * @param {Object} controls - コントロールの参照
     * @param {Object} session - セッション設定
     * @returns {void}
     */
    function saveDialogToSession(controls, session) {
        var sizePercent = readSizePercent(controls);
        var strokeWidthPt = readLengthPt(controls.strokeInput, "strokeUnits", false);
        var cornerRadiusPt = readLengthPt(controls.cornerInput, "rulerType", true);
        if (sizePercent !== null) session.percent = sizePercent;
        if (strokeWidthPt !== null) session.strokeWidthPt = strokeWidthPt;
        if (cornerRadiusPt !== null) session.cornerRadiusPt = cornerRadiusPt;

        session.fillFirst = controls.fillFirstCheckbox.value;
        session.fillSecond = controls.fillSecondCheckbox.value;
        session.overallFrame = controls.overallFrameCheckbox.value;
        session.divider = controls.dividerCheckbox.value;
        session.balanceMode = getBalanceMode(controls);
        session.balanceWidthPt = readBalanceWidthPt(controls);
        settingsStore.save(session);
    }

    /**
     * 設定ダイアログを表示する。閉じたらプレビューを片付ける
     * @returns {Object|null} 描画の設定。キャンセルなら null
     */
    function showSettingsDialog() {
        var session = getSessionSettings();
        var controls = buildSettingsDialog(session);
        bindDialogEvents(controls);

        /* 中断した実行で残ったプレビューレイヤーを消してから始める / start without a preview layer left over from an interrupted run */
        removePreviewLayer();
        prepareDialogWindow(controls.dialog, SCRIPT_NAME);
        var dialogResult = controls.dialog.show();
        removePreviewLayer();
        saveDialogToSession(controls, session);

        return (dialogResult === 1) ? collectDrawOptions(controls) : null;
    }

    // =========================================
    // 角丸アルゴリズム / Round Any Corner
    // 選択したアンカーだけを丸める / rounds only the selected anchors
    // Based on: Hiroyuki Sato (MIT) https://github.com/shspage
    // =========================================

    /**
     * 選択したアンカーの角を半径 rr で丸める
     * @param {PathItem[]} s - 対象のパス
     * @param {Object} conf - 設定（rr: 半径）
     * @returns {void}
     */
    function roundAnyCorner(s, conf) {
        var rr = conf.rr;

        var p, op, pnts;
        var skipList, adjRdirAtEnd, redrawFlg;
        var i, nxi, pvi, q, d, ds, r, g, t, qb;
        var anc1, ldir1, rdir1, anc2, ldir2, rdir2;

        var hanLen = 4 * (Math.sqrt(2) - 1) / 3;
        var ptyp = PointType.SMOOTH;

        for (var j = 0; j < s.length; j++) {
            p = s[j].pathPoints;
            if (readjustAnchors(p) < 2) continue;
            op = !s[j].closed;
            pnts = op ? [getDat(p[0])] : [];
            redrawFlg = false;
            adjRdirAtEnd = 0;

            skipList = [(op || !isSelected(p[0]) || !isCorner(p, 0))];
            for (i = 1; i < p.length; i++) {
                skipList.push((!isSelected(p[i])
                    || !isCorner(p, i)
                    || (op && i == p.length - 1)));
            }

            for (i = 0; i < p.length; i++) {
                nxi = parseIdx(p, i + 1);
                if (nxi < 0) break;

                pvi = parseIdx(p, i - 1);

                q = [p[i].anchor, p[i].rightDirection,
                p[nxi].leftDirection, p[nxi].anchor];

                ds = dist(q[0], q[3]) / 2;
                if (arrEq(q[0], q[1]) && arrEq(q[2], q[3])) {
                    r = Math.min(ds, rr);
                    g = getRad(q[0], q[3]);
                    anc1 = getPnt(q[0], g, r);
                    ldir1 = getPnt(anc1, g + Math.PI, r * hanLen);

                    if (skipList[nxi]) {
                        if (!skipList[i]) {
                            pnts.push([anc1, anc1, ldir1, ptyp]);
                            redrawFlg = true;
                        }
                        pnts.push(getDat(p[nxi]));
                    } else {
                        if (r < rr) {
                            pnts.push([anc1,
                                getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                ldir1,
                                ptyp]);
                        } else {
                            if (!skipList[i]) pnts.push([anc1, anc1, ldir1, ptyp]);
                            anc2 = getPnt(q[3], g + Math.PI, r);
                            pnts.push([anc2,
                                getPnt(anc2, g, r * hanLen),
                                anc2,
                                ptyp]);
                        }
                        redrawFlg = true;
                    }
                } else {
                    d = getT4Len(q, 0) / 2;
                    r = Math.min(d, rr);
                    t = getT4Len(q, r);
                    anc1 = bezier(q, t);
                    rdir1 = defHan(t, q, 1);
                    ldir1 = getPnt(anc1, getRad(rdir1, anc1), r * hanLen);

                    if (skipList[nxi]) {
                        if (skipList[i]) {
                            pnts.push(getDat(p[nxi]));
                        } else {
                            pnts.push([anc1, rdir1, ldir1, ptyp]);
                            with (p[nxi]) pnts.push([anchor,
                                rightDirection,
                                adjHan(anchor, leftDirection, 1 - t),
                                ptyp]);
                            redrawFlg = true;
                        }
                    } else {
                        if (r < rr) {
                            if (skipList[i]) {
                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    defHan(t, q, 0),
                                    ptyp]);
                            } else {
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    ldir1,
                                    ptyp]);
                            }
                        } else {
                            if (skipList[i]) {
                                t = getT4Len(q, -r);
                                anc2 = bezier(q, t);

                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }

                                ldir2 = defHan(t, q, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);

                                pnts.push([anc2, rdir2, ldir2, ptyp]);
                            } else {
                                qb = [anc1, rdir1, adjHan(q[3], q[2], 1 - t), q[3]];
                                t = getT4Len(qb, -r);
                                anc2 = bezier(qb, t);
                                ldir2 = defHan(t, qb, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);
                                rdir1 = adjHan(anc1, rdir1, t);

                                pnts.push([anc1, rdir1, ldir1, ptyp],
                                    [anc2, rdir2, ldir2, ptyp]);
                            }
                        }
                        redrawFlg = true;
                    }
                }
            }
            if (adjRdirAtEnd > 0) {
                pnts[pnts.length - 1][1] = adjHan(p[0].anchor, p[0].rightDirection, adjRdirAtEnd);
            }

            if (redrawFlg) {
                for (i = p.length - 1; i > 0; i--) p[i].remove();

                for (i = 0; i < pnts.length; i++) {
                    var pt = i > 0 ? p.add() : p[0];
                    with (pt) {
                        anchor = pnts[i][0];
                        rightDirection = pnts[i][1];
                        leftDirection = pnts[i][2];
                        pointType = pnts[i][3];
                    }
                }
            }
        }
        app.activeDocument.selection = s;
    }

    /**
     * 点から角度 rad の方向へ len 進んだ点を返す
     * @param {number[]} pt - 起点
     * @param {number} rad - 角度（ラジアン）
     * @param {number} len - 距離
     * @returns {number[]} 点
     */
    function getPnt(pt, rad, len) {
        return [pt[0] + Math.cos(rad) * len,
        pt[1] + Math.sin(rad) * len];
    }

    /**
     * ベジェ曲線の t における接線方向のハンドルを返す
     * @param {number} t - 曲線上の位置（0〜1）
     * @param {Array} q - 制御点4つ
     * @param {number} n - 0 なら前側、1 なら後側
     * @returns {number[]} ハンドルの点
     */
    function defHan(t, q, n) {
        return [t * (t * (q[n][0] - 2 * q[n + 1][0] + q[n + 2][0]) + 2 * (q[n + 1][0] - q[n][0])) + q[n][0],
        t * (t * (q[n][1] - 2 * q[n + 1][1] + q[n + 2][1]) + 2 * (q[n + 1][1] - q[n][1])) + q[n][1]];
    }

    /**
     * ベジェ曲線の t における点を返す
     * @param {Array} q - 制御点4つ
     * @param {number} t - 曲線上の位置（0〜1）
     * @returns {number[]} 点
     */
    function bezier(q, t) {
        var u = 1 - t;
        return [u * u * u * q[0][0] + 3 * u * t * (u * q[1][0] + t * q[2][0]) + t * t * t * q[3][0],
        u * u * u * q[0][1] + 3 * u * t * (u * q[1][1] + t * q[2][1]) + t * t * t * q[3][1]];
    }

    /**
     * ハンドルの長さを m 倍にする
     * @param {number[]} anc - アンカー
     * @param {number[]} dir - ハンドル
     * @param {number} m - 倍率
     * @returns {number[]} 新しいハンドルの点
     */
    function adjHan(anc, dir, m) {
        return [anc[0] + (dir[0] - anc[0]) * m,
        anc[1] + (dir[1] - anc[1]) * m];
    }

    /**
     * アンカーが角（折れ点）か
     * @param {PathPoints} p - パスポイント
     * @param {number} idx - 添字
     * @returns {boolean} 角か
     */
    function isCorner(p, idx) {
        var pnt0 = getAnglePnt(p, idx, -1);
        var pnt1 = getAnglePnt(p, idx, 1);
        if (!pnt0 || !pnt1) return false;
        if (pnt0.length < 1 || pnt1.length < 1) return false;
        var rad = getRad2(pnt0, p[idx].anchor, pnt1, true);
        if (rad > Math.PI - 0.1) return false;
        return true;
    }

    /**
     * 角度を測るための隣の点を返す
     * @param {PathPoints} p - パスポイント
     * @param {number} idx1 - 添字
     * @param {number} dir - -1 なら前、1 なら後
     * @returns {number[]|null} 点。無ければ null、重なっていれば空配列
     */
    function getAnglePnt(p, idx1, dir) {
        if (!dir) dir = -1;
        var idx2 = parseIdx(p, idx1 + dir);
        if (idx2 < 0) return null;
        var p2 = p[idx2];
        with (p[idx1]) {
            if (dir < 0) {
                if (arrEq(leftDirection, anchor)) {
                    if (arrEq(p2.anchor, anchor)) return [];
                    if (arrEq(p2.anchor, p2.rightDirection)
                        || arrEq(p2.rightDirection, anchor)) return p2.anchor;
                    else return p2.rightDirection;
                } else {
                    return leftDirection;
                }
            } else {
                if (arrEq(anchor, rightDirection)) {
                    if (arrEq(anchor, p2.anchor)) return [];
                    if (arrEq(p2.anchor, p2.leftDirection)
                        || arrEq(anchor, p2.leftDirection)) return p2.anchor;
                    else return p2.leftDirection;
                } else {
                    return rightDirection;
                }
            }
        }
    }

    /**
     * 2つの配列の要素が等しいか
     * @param {number[]} arr1 - 配列1
     * @param {number[]} arr2 - 配列2
     * @returns {boolean} 等しいか
     */
    function arrEq(arr1, arr2) {
        for (var i = 0; i < arr1.length; i++) {
            if (arr1[i] != arr2[i]) return false;
        }
        return true;
    }

    /**
     * 2点間の距離
     * @param {number[]} p1 - 点1
     * @param {number[]} p2 - 点2
     * @returns {number} 距離
     */
    function dist(p1, p2) {
        return Math.sqrt(Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2));
    }

    /**
     * 2点間の距離の2乗
     * @param {number[]} p1 - 点1
     * @param {number[]} p2 - 点2
     * @returns {number} 距離の2乗
     */
    function dist2(p1, p2) {
        return Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2);
    }

    /**
     * p1 から p2 への角度
     * @param {number[]} p1 - 起点
     * @param {number[]} p2 - 終点
     * @returns {number} 角度（ラジアン）
     */
    function getRad(p1, p2) {
        return Math.atan2(p2[1] - p1[1],
            p2[0] - p1[0]);
    }

    /**
     * o を頂点とする p1-o-p2 の角度
     * @param {number[]} p1 - 点1
     * @param {number[]} o - 頂点
     * @param {number[]} p2 - 点2
     * @returns {number} 角度（ラジアン）
     */
    function getRad2(p1, o, p2) {
        var v1 = normalize(p1, o);
        var v2 = normalize(p2, o);
        return Math.acos(v1[0] * v2[0] + v1[1] * v2[1]);
    }

    /**
     * o から p への単位ベクトル
     * @param {number[]} p - 点
     * @param {number[]} o - 原点
     * @returns {number[]} 単位ベクトル
     */
    function normalize(p, o) {
        var d = dist(p, o);
        return d == 0 ? [0, 0] : [(p[0] - o[0]) / d,
        (p[1] - o[1]) / d];
    }

    /**
     * ベジェ曲線上で長さ len の位置の t を返す（len が 0 なら全長）
     * @param {Array} q - 制御点4つ
     * @param {number} len - 長さ（負なら終点側から）
     * @returns {number} t、または全長
     */
    function getT4Len(q, len) {
        var m = [q[3][0] - q[0][0] + 3 * (q[1][0] - q[2][0]),
        q[0][0] - 2 * q[1][0] + q[2][0],
        q[1][0] - q[0][0]];
        var n = [q[3][1] - q[0][1] + 3 * (q[1][1] - q[2][1]),
        q[0][1] - 2 * q[1][1] + q[2][1],
        q[1][1] - q[0][1]];
        var k = [m[0] * m[0] + n[0] * n[0],
        4 * (m[0] * m[1] + n[0] * n[1]),
        2 * ((m[0] * m[2] + n[0] * n[2]) + 2 * (m[1] * m[1] + n[1] * n[1])),
        4 * (m[1] * m[2] + n[1] * n[2]),
        m[2] * m[2] + n[2] * n[2]];

        var fullLen = getLength(k, 1);

        if (len == 0) {
            return fullLen;
        } else if (len < 0) {
            len += fullLen;
            if (len < 0) return 0;
        } else if (len > fullLen) {
            return 1;
        }

        var t, d;
        var t0 = 0;
        var t1 = 1;
        var torelance = 0.001;

        for (var h = 1; h < 30; h++) {
            t = t0 + (t1 - t0) / 2;
            d = len - getLength(k, t);
            if (Math.abs(d) < torelance) break;
            else if (d < 0) t1 = t;
            else t0 = t;
        }
        return t;
    }

    /**
     * ベジェ曲線の 0〜t の長さ（シンプソン則）
     * @param {number[]} k - 係数
     * @param {number} t - 終わりの位置
     * @returns {number} 長さ
     */
    function getLength(k, t) {
        var h = t / 128;
        var hh = h * 2;
        var fc = function (t, k) {
            return Math.sqrt(t * (t * (t * (t * k[0] + k[1]) + k[2]) + k[3]) + k[4]) || 0;
        };
        var total = (fc(0, k) - fc(t, k)) / 2;
        for (var i = h; i < t; i += hh) total += 2 * fc(i, k) + fc(i + h, k);
        return total * hh;
    }

    /**
     * 重なったアンカーを1つにまとめる
     * @param {PathPoints} p - パスポイント
     * @returns {number} まとめたあとの点の数
     */
    function readjustAnchors(p) {
        var minDist = 0.0025;
        if (p.length < 2) return 1;
        var i;

        if (p.parent.closed) {
            for (i = p.length - 1; i >= 1; i--) {
                if (dist2(p[0].anchor, p[i].anchor) < minDist) {
                    p[0].leftDirection = p[i].leftDirection;
                    p[i].remove();
                } else {
                    break;
                }
            }
        }

        for (i = p.length - 1; i >= 1; i--) {
            if (dist2(p[i].anchor, p[i - 1].anchor) < minDist) {
                p[i - 1].rightDirection = p[i].rightDirection;
                p[i].remove();
            }
        }

        return p.length;
    }

    /**
     * 添字を点の数の範囲に収める（閉じたパスは巡回、開いたパスは範囲外で -1）
     * @param {PathPoints} p - パスポイント
     * @param {number} n - 添字
     * @returns {number} 収めた添字
     */
    function parseIdx(p, n) {
        var len = p.length;
        if (p.parent.closed) {
            return n >= 0 ? n % len : len - Math.abs(n % len);
        } else {
            return (n < 0 || n > len - 1) ? -1 : n;
        }
    }

    /**
     * パスポイントの座標と種類を配列で返す
     * @param {PathPoint} p - パスポイント
     * @returns {Array} [アンカー, 右ハンドル, 左ハンドル, 種類]
     */
    function getDat(p) {
        with (p) return [anchor, rightDirection, leftDirection, pointType];
    }

    /**
     * アンカーが選択されているか
     * @param {PathPoint} p - パスポイント
     * @returns {boolean} 選択されているか
     */
    function isSelected(p) {
        return p.selected == PathPointSelection.ANCHORPOINT;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 2つのオブジェクトを左から右（上下なら上から下）の順に並べ替える
     * @returns {void}
     */
    function orderTargetsByPosition() {
        var bounds1 = targetBounds[0];
        var bounds2 = targetBounds[1];
        var secondComesFirst = isVerticalSplit ? (bounds1[1] < bounds2[1]) : (bounds1[0] > bounds2[0]);
        if (secondComesFirst) {
            targetItems.reverse();
            targetBounds.reverse();
        }
    }

    /**
     * 選択した2つのオブジェクトの背面に、2分割の背景を作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocument"));
            return;
        }

        doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字の編集中は TextRange が返り、length は文字数になる / while editing text, the selection is a TextRange whose length counts characters */
        if (!selectedItems || selectedItems.typename === "TextRange" || selectedItems.length !== 2) {
            alert(getLabel("alert.selectTwoItems"));
            return;
        }

        targetItems = [selectedItems[0], selectedItems[1]];
        targetBounds = [getItemBounds(targetItems[0]), getItemBounds(targetItems[1])];
        /* すき間の大きいほうの向きに分ける（同じなら左右）/ split along the larger gap; a tie goes left/right */
        isVerticalSplit = (getGapPt(targetBounds, true) > getGapPt(targetBounds, false));
        orderTargetsByPosition();

        var originalActiveLayer = doc.activeLayer;
        var drawOptions = showSettingsDialog();
        if (drawOptions) {
            drawSplitBackground(drawOptions, getBackmostLayer(targetItems[0].layer, targetItems[1].layer));
            /* 全体の枠と区切り線より前面に元のオブジェクトを出す / bring the objects in front of the frame and divider */
            targetItems[0].zOrder(ZOrderMethod.BRINGTOFRONT);
            targetItems[1].zOrder(ZOrderMethod.BRINGTOFRONT);
            doc.selection = null;
        } else {
            /* キャンセルしたら選択を戻す / restore the selection on cancel */
            doc.selection = targetItems;
        }

        try {
            doc.activeLayer = originalActiveLayer;
        } catch (e) {
            /* ロックされたレイヤーなどは戻せないことがある / a locked layer may refuse to become active again */
        }
    }

    main();

})();
