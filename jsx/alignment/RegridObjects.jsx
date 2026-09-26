#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

だいたいグリッド状に並んでいる選択オブジェクトを、指定した左右・上下の間隔で再配置します。
間隔は現在の定規単位で入力でき、常時プレビューで結果を確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegridObjects.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n08861d0e40c3

### Overview

Re-lays out a roughly grid-shaped selection at the horizontal and vertical spacing you specify.
The spacing is entered in the current ruler units, with a continuous preview of the result.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegridObjects.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RegridObjects";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegridObjects.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegridObjects.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n08861d0e40c3"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 行列入れ替えで「同じ列」とみなす左端Xの許容差（pt）/ Tolerance for treating lefts as one column when transposing (pt) */
    var TRANSPOSE_SNAP_X_TOLERANCE = 8.0;

    /* 行列入れ替えで「同じ行」とみなす上端Yの許容差（pt）/ Tolerance for treating tops as one row when transposing (pt) */
    var TRANSPOSE_SNAP_Y_TOLERANCE = 8.0;

    /* ハニカムの行送りに掛ける係数（通常・レンガ状は 1.0）/ Row-step factor for the honeycomb layout (normal and brick use 1.0) */
    var HONEYCOMB_ROW_STEP_FACTOR = 0.75;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS     = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 8;                  /* パネル内の要素間隔 / panel spacing */
    var OPTION_SPACING     = 6;                  /* オプションパネル内の要素間隔 / spacing in the Options panel */
    var COLUMN_SPACING     = 12;                 /* 2カラムの間隔 / gap between columns */
    var SUB_OPTION_MARGINS = [15, 0, 0, 0];      /* サブオプションの字下げ / indent of sub-options */
    var GAP_INPUT_CHARS    = 4;                  /* 間隔の入力欄の文字数 / width of the gap fields, in characters */

    /**
     * ウィンドウに共通のレイアウト設定（縦並び・外周余白・要素間隔）を適用する
     * @param {Window} targetWindow - 対象のダイアログウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウト設定（縦並び・横いっぱい・余白・要素間隔）を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループ（ボタン列など）に共通の横並び設定を適用する
     * @param {Group} rowGroup - 対象のグループ
     * @param {string} [alignment] - グループの配置（"left" / "center" / "right" など。省略時は "left"）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment) {
        rowGroup.orientation = "row";
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = PANEL_SPACING;
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
    // セッション記憶 / Session memory
    // =========================================

    /* 強制グリッド（ダイアログから切替）/ Force Grid mode (toggled from dialog) */
    $.global.__regridForceGrid = $.global.__regridForceGrid || false;

    /* 中央揃え：各セルの天地左右中央に整列（強制グリッドのサブオプション）/ Center each object in its cell (sub-option of Force Grid) */
    $.global.__regridCenterInCell = $.global.__regridCenterInCell || false;

    // =========================================
    // プレビュー履歴ユーティリティ / Preview history util
    // =========================================

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 他スクリプトでもこのブロックをコピペすれば再利用できます。
     * $.global に載せて共有するので、このスクリプトで使わない cancelTask も残しておく。
     * 使い方:
     *   PreviewHistory.start();      // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();       // プレビュー描画ごと（または操作ごと）にカウント(+1)
     *   PreviewHistory.undo();       // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */

    (function (g) {
        if (!g.PreviewHistory) {
            g.PreviewHistory = {
                start: function () { g.__previewUndoCount = 0; },
                bump: function () { g.__previewUndoCount = (g.__previewUndoCount | 0) + 1; },
                undo: function () {
                    var n = g.__previewUndoCount | 0;
                    try { for (var i = 0; i < n; i++) app.executeMenuCommand('undo'); } catch (e) { }
                    g.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

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
     * 現在のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* ラベル定義（カテゴリ分け）/ Label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "グリッドの再配置", en: "Regrid Objects" }
        },
        panel: {
            spacing: { ja: "間隔", en: "Spacing" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            horizontal: { ja: "左右", en: "H" },
            vertical: { ja: "上下", en: "V" }
        },
        checkbox: {
            link: { ja: "連動", en: "Link" },
            brick: { ja: "レンガ状", en: "Brick" },
            honeycomb: { ja: "ハニカム", en: "Honeycomb" },
            forceGrid: { ja: "強制グリッド", en: "Force Grid" },
            centerInCell: { ja: "中央揃え", en: "Center in Cell" },
            transpose: { ja: "行列入れ替え", en: "Swap Rows{slash}Columns" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            horizontal: {
                ja: "左右に隣り合うオブジェクトのあいだにあける間隔です。↑↓キーや∧∨で増減できます。",
                en: "Gap left between horizontally adjacent objects. The arrow keys and the stepper buttons step the value."
            },
            vertical: {
                ja: "上下に隣り合うオブジェクトのあいだにあける間隔です。↑↓キーや∧∨で増減できます。",
                en: "Gap left between vertically adjacent objects. The arrow keys and the stepper buttons step the value."
            },
            link: {
                ja: "左右の間隔と同じ値を上下にも使います。オフにすると上下を個別に指定できます。",
                en: "Uses the horizontal gap for the vertical one too. Turn it off to set them separately."
            },
            brick: {
                ja: "1行ごとに半ピッチずらして、レンガ積みのように配置します。",
                en: "Offsets every other row by half a pitch, like a brick wall."
            },
            honeycomb: {
                ja: "レンガ状に加えて行送りを詰め、六角形に近い並びにします。",
                en: "Adds to the brick offset a tighter row step, giving a honeycomb-like arrangement."
            },
            forceGrid: {
                ja: "歯抜けや行ごとの個数違いがあっても、行数・列数をそろえた格子として並べ直します。",
                en: "Rebuilds the layout as an even grid even when rows have gaps or different counts."
            },
            centerInCell: {
                ja: "各セルの中でオブジェクトを天地左右中央にそろえます（強制グリッドのときだけ使えます）。",
                en: "Centers each object inside its cell. Available only with Force Grid."
            },
            transpose: {
                ja: "行と列を入れ替えて並べ直します。歯抜けのある配置にも対応します。",
                en: "Swaps rows and columns. Layouts with gaps are handled too."
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
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            noSelection: { ja: "グリッド状に並んだオブジェクトを選択してください。", en: "Select objects arranged in a grid." },
            needTwo: { ja: "2つ以上のオブジェクトを選択してください。", en: "Select at least two objects." },
            cellConflict: {
                ja: "行列を入れ替えられません。同じセルに複数のオブジェクトが入ります。\n重なっているオブジェクトがないか確認してください。\n衝突したセル：行 {row}・列 {col}",
                en: "Cannot swap rows and columns: more than one object falls into the same cell.\nCheck for overlapping objects.\nConflicting cell: row {row}, column {col}"
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す。"{slash}" は "/" に置き換える
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (labelNode == null) return labelPath;
        }
        var labelString = labelNode[uiLang] || labelNode.ja || labelNode.en || labelPath;
        return String(labelString).replace(/\{slash\}/g, "/");
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object|string} labelSet - ラベル、またはラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名と間隔の入力欄を1行追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {string} initialText - 入力欄の初期値
     * @param {string} tooltipPath - 入力欄の tooltip のラベルのパス
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addGapField(parentGroup, labelPath, initialText, tooltipPath) {
        var gapRow = parentGroup.add('group');
        gapRow.add('statictext', undefined, labelText(labelPath));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = gapRow.add('group');
        stepperInputGroup.orientation = 'row';
        stepperInputGroup.alignChildren = ['left', 'center'];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var gapInput;
        /* 増減後は手入力と同じくプレビューを更新する（onChanging は bindDialogEvents() で結線）
           after stepping, refresh the preview just like typing does */
        var stepperGroup = addStepper(stepperInputGroup, function () { return gapInput; }, {
            onStep: function (numberInput) { if (numberInput.onChanging) numberInput.onChanging(); }
        });
        gapInput = stepperInputGroup.add('edittext', undefined, initialText);
        gapInput.helpTip = getLabel(tooltipPath);
        gapInput.characters = GAP_INPUT_CHARS;
        gapInput.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(gapInput, stepperGroup);
        return gapInput;
    }

    /**
     * 間隔の入力欄の有効／無効を、∧∨ごとまとめて切り替える
     * @param {EditText} gapInput - addGapField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setGapFieldEnabled(gapInput, isEnabled) {
        gapInput.enabled = isEnabled;
        gapInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(gapInput.stepperGroup);
    }

    /**
     * オプションのチェックボックスを1行追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - チェックボックスのラベルのパス
     * @param {string} tooltipPath - tooltip のラベルのパス
     * @param {boolean} isSubOption - サブオプションとして字下げするか
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, isSubOption) {
        var optionRow = parentPanel.add('group');
        setupRow(optionRow, 'left');
        if (isSubOption) optionRow.margins = SUB_OPTION_MARGINS;
        var optionCheckbox = optionRow.add('checkbox', undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        return optionCheckbox;
    }

    /**
     * 間隔設定ダイアログを組み立てる（イベント結線は呼び出し側で行う）
     * @param {number} initialGapX - 左右間隔の初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {object} ダイアログと各コントロールの参照
     */
    function buildGridSpacingDialog(initialGapX, rulerUnit) {
        // タイトルとバージョンを合成 / combine title and version
        var spacingDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(spacingDialog);

        // パネル名に単位を出す（日本語は全角かっこ）/ show unit in panel title (full-width parentheses in Japanese)
        var unitSuffix = (uiLang === "ja") ? "（" + rulerUnit.label + "）" : " (" + rulerUnit.label + ")";
        var spacingPanel = spacingDialog.add('panel', undefined, getLabel('panel.spacing') + unitSuffix);
        // 2カラム構成なので row のまま余白のみ共通化 / two-column panel: keep row, share margins
        spacingPanel.orientation = 'row';
        spacingPanel.alignChildren = 'top';
        spacingPanel.alignment = 'fill';
        spacingPanel.margins = PANEL_MARGINS;
        spacingPanel.spacing = COLUMN_SPACING;

        // 左カラム / left column
        var gapInputColumn = spacingPanel.add('group');
        gapInputColumn.orientation = 'column';
        gapInputColumn.alignChildren = 'left';

        // 初期値は pt を表示単位に換算して表示 / show initial value converted from pt to the display unit
        var horizontalGapInput = addGapField(gapInputColumn, 'fieldLabel.horizontal',
            (initialGapX / rulerUnit.pointsPerUnit).toFixed(1), 'tooltip.horizontal');
        // 連動ONで始めるので上下は左右と同じ値 / Link starts on, so vertical mirrors horizontal
        var verticalGapInput = addGapField(gapInputColumn, 'fieldLabel.vertical', horizontalGapInput.text, 'tooltip.vertical');

        // 右カラム（連動）/ right column (link)
        var linkColumn = spacingPanel.add('group');
        linkColumn.orientation = 'column';
        linkColumn.alignChildren = 'center';
        linkColumn.alignment = ['fill', 'fill'];
        var linkColumnSpacer = linkColumn.add('statictext', undefined, '');
        linkColumnSpacer.alignment = ['fill', 'fill'];
        var linkCheckbox = linkColumn.add('checkbox', undefined, getLabel('checkbox.link'));
        linkCheckbox.helpTip = getLabel('tooltip.link');

        // オプション（チェックボックスをまとめる）/ Options panel
        var optionsPanel = spacingDialog.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel, OPTION_SPACING);

        // レンガ状／ハニカム（レンガ状のサブオプション）/ Brick and Honeycomb (sub-option of Brick)
        var brickCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.brick', 'tooltip.brick', false);
        var honeycombCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.honeycomb', 'tooltip.honeycomb', true);
        honeycombCheckbox.enabled = false;

        // 強制グリッド／中央揃え（強制グリッドのサブオプション）/ Force Grid and Center in cell (sub-option of Force Grid)
        var forceGridCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.forceGrid', 'tooltip.forceGrid', false);
        forceGridCheckbox.value = !!$.global.__regridForceGrid;
        var centerInCellCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.centerInCell', 'tooltip.centerInCell', true);
        centerInCellCheckbox.value = !!$.global.__regridCenterInCell;
        centerInCellCheckbox.enabled = forceGridCheckbox.value;

        // 行列入れ替え / Swap rows/columns
        var transposeCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.transpose', 'tooltip.transpose', false);

        // 初期状態 / initial state
        linkCheckbox.value = true;
        horizontalGapInput.active = true;
        setGapFieldEnabled(verticalGapInput, false);

        // ボタン行（パネル外・中央寄せ、いっぱいに広げない）/ buttons (outside panels, centered, not stretched)
        var btnRowGroup = spacingDialog.add('group');
        setupRow(btnRowGroup, 'center');
        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        return {
            spacingDialog: spacingDialog,
            horizontalGapInput: horizontalGapInput,
            verticalGapInput: verticalGapInput,
            linkCheckbox: linkCheckbox,
            brickCheckbox: brickCheckbox,
            honeycombCheckbox: honeycombCheckbox,
            forceGridCheckbox: forceGridCheckbox,
            centerInCellCheckbox: centerInCellCheckbox,
            transposeCheckbox: transposeCheckbox
        };
    }

    /**
     * 左右・上下の入力値を pt で読む（数値でなければ 0）。［連動］がONなら上下は左右と同じ値
     * @param {object} dialogControls - buildGridSpacingDialog() の戻り値
     * @param {number} pointsPerUnit - 表示単位1あたりのポイント数
     * @returns {{gapX: number, gapY: number}} 左右・上下の間隔（pt）
     */
    function readGapsInPoints(dialogControls, pointsPerUnit) {
        var gapX = parseFloat(dialogControls.horizontalGapInput.text);
        var gapY = dialogControls.linkCheckbox.value ? gapX : parseFloat(dialogControls.verticalGapInput.text);
        if (isNaN(gapX)) gapX = 0;
        if (isNaN(gapY)) gapY = 0;
        return { gapX: gapX * pointsPerUnit, gapY: gapY * pointsPerUnit };
    }

    /**
     * ［レンガ状］［ハニカム］［行列入れ替え］の状態に合わせて間隔を適用する。
     * 行列入れ替えONのときは、間隔0で並べ直す→転置→基準を取り直す→間隔を適用、の順に進める
     * @param {object} layoutActions - main が用意する配置処理一式
     * @param {object} dialogControls - buildGridSpacingDialog() の戻り値
     * @param {{gapX: number, gapY: number}} gaps - 左右・上下の間隔（pt）
     * @param {boolean} bumpHistory - プレビュー時は各ステップで PreviewHistory.bump() する（確定時は false）
     * @returns {void}
     */
    function applyLayoutFromDialog(layoutActions, dialogControls, gaps, bumpHistory) {
        var isBrick = dialogControls.brickCheckbox.value;
        var isHoneycomb = isBrick && dialogControls.honeycombCheckbox.value;
        var rowStepFactor = isHoneycomb ? HONEYCOMB_ROW_STEP_FACTOR : 1.0;

        if (dialogControls.transposeCheckbox.value) {
            // 転置は間隔0で実行し、その後に間隔を適用 / Transpose with zero gaps, then apply the gaps
            layoutActions.applySpacing(0, 0, false, 1.0);
            if (bumpHistory) PreviewHistory.bump();

            layoutActions.transpose();
            if (bumpHistory) PreviewHistory.bump();

            // 転置後の配置を新しい基準に / adopt the transposed layout as the baseline
            layoutActions.resetBaselineToCurrent();
        }
        layoutActions.applySpacing(gaps.gapX, gaps.gapY, isBrick, rowStepFactor);
        if (bumpHistory) PreviewHistory.bump();
    }

    /**
     * ダイアログのイベントを結線する
     * @param {object} dialogControls - buildGridSpacingDialog() の戻り値
     * @param {Function} updatePreview - プレビューを描き直す関数
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, updatePreview) {
        var horizontalGapInput = dialogControls.horizontalGapInput;
        var verticalGapInput = dialogControls.verticalGapInput;
        var linkCheckbox = dialogControls.linkCheckbox;
        var brickCheckbox = dialogControls.brickCheckbox;
        var honeycombCheckbox = dialogControls.honeycombCheckbox;
        var forceGridCheckbox = dialogControls.forceGridCheckbox;
        var centerInCellCheckbox = dialogControls.centerInCellCheckbox;

        horizontalGapInput.onChanging = updatePreview;
        // 連動ON中の上下欄は無効なので、ここに来るのは連動OFFのときだけ / only reachable with Link off (the field is disabled otherwise)
        verticalGapInput.onChanging = updatePreview;

        linkCheckbox.onClick = function () {
            setGapFieldEnabled(verticalGapInput, !linkCheckbox.value);
            updatePreview();
        };

        // レンガ状OFFでハニカムも解除 / Brick off also clears Honeycomb
        brickCheckbox.onClick = function () {
            honeycombCheckbox.enabled = brickCheckbox.value;
            if (!brickCheckbox.value) honeycombCheckbox.value = false;
            updatePreview();
        };
        honeycombCheckbox.onClick = updatePreview;

        // 強制グリッドOFFで中央揃えも解除 / Force Grid off also clears Center in Cell
        forceGridCheckbox.onClick = function () {
            centerInCellCheckbox.enabled = forceGridCheckbox.value;
            if (!forceGridCheckbox.value) centerInCellCheckbox.value = false;
            updatePreview();
        };
        centerInCellCheckbox.onClick = updatePreview;

        // 行列入れ替えはトグル：ONで転置、OFFで転置前に戻す。連動OFFなら左右・上下の値も入れ替える
        // Toggle: ON transposes, OFF reverts. With Link off, the H/V values are swapped too
        dialogControls.transposeCheckbox.onClick = function () {
            if (!linkCheckbox.value) {
                var horizontalText = horizontalGapInput.text;
                horizontalGapInput.text = verticalGapInput.text;
                verticalGapInput.text = horizontalText;
            }
            updatePreview();
        };
    }

    /**
     * 間隔設定ダイアログを表示し、プレビューと確定適用を行う
     * @param {object} layoutActions - main が用意する配置処理一式
     *   applySpacing（(gapX, gapY, isBrick, rowStepFactor) => void）、transpose（行列入れ替え）、
     *   restoreInitialPositions（ダイアログ開始時点へ戻す）、resetBaselineToCurrent（現在位置を新しい基準にする）
     * @param {number} initialGapX - 左右間隔の初期値（pt）
     * @returns {void}
     */
    function showGridSpacingDialog(layoutActions, initialGapX) {
        // 入力値は表示単位、内部処理は pt / inputs are in display units; internal geometry is in pt
        var rulerUnit = getUnitInfo();
        var dialogControls = buildGridSpacingDialog(initialGapX, rulerUnit);

        /**
         * 直前のプレビューを一括Undoし、現在のUI状態で描き直す（ヒストリーを汚さない）
         * @returns {void}
         */
        function updatePreview() {
            PreviewHistory.undo();
            // 強制グリッド／中央揃えの状態をグローバルに反映 / store Force Grid / Center in Cell globally
            $.global.__regridForceGrid = !!dialogControls.forceGridCheckbox.value;
            $.global.__regridCenterInCell = !!dialogControls.centerInCellCheckbox.value;
            // Undo後の現在位置を基準に作り直す / rebuild the baseline from the post-undo positions
            layoutActions.resetBaselineToCurrent();

            if (dialogControls.linkCheckbox.value) {
                dialogControls.verticalGapInput.text = dialogControls.horizontalGapInput.text;
            }
            applyLayoutFromDialog(layoutActions, dialogControls, readGapsInPoints(dialogControls, rulerUnit.pointsPerUnit), true);
        }

        bindDialogEvents(dialogControls, updatePreview);

        // プレビュー用ヒストリー管理を開始し、開いたときに一度プレビュー / start the preview history and draw once
        PreviewHistory.start();
        updatePreview();

        var dialogResult = dialogControls.spacingDialog.show();

        // プレビュー分を一括Undo / undo the preview
        PreviewHistory.undo();
        if (dialogResult == 1) {
            // OK：最終値で確定適用 / OK: apply with the final values
            layoutActions.resetBaselineToCurrent();
            applyLayoutFromDialog(layoutActions, dialogControls, readGapsInPoints(dialogControls, rulerUnit.pointsPerUnit), false);
        } else {
            // キャンセル：ダイアログ開始時点へ戻す / Cancel: restore the initial positions
            layoutActions.restoreInitialPositions();
        }

        // 念のためカウンタ初期化 / reset counter
        PreviewHistory.start();
    }

    // =========================================
    // レイアウト計算の補助 / Layout helpers
    // =========================================

    /**
     * レイアウト計算に使う外接矩形を返す。
     * すでにグループになっているものは中身を分解せず「グループ＝1つのオブジェクト」として扱う。
     * クリップグループはクリップパスの geometricBounds（＝可視領域）を優先し、
     * それ以外は pageItem.geometricBounds を使う
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number[]} [left, top, right, bottom]
     */
    function getLayoutBounds(pageItem) {
        if (pageItem.typename === 'GroupItem' && pageItem.clipped) {
            // GroupItem の中から clipping パスを探す / look for the clipping path
            for (var i = 0; i < pageItem.pathItems.length; i++) {
                if (pageItem.pathItems[i].clipping) return pageItem.pathItems[i].geometricBounds;
            }
            // CompoundPath が clipping のケース / compound path used as the mask
            for (var j = 0; j < pageItem.compoundPathItems.length; j++) {
                var compoundPath = pageItem.compoundPathItems[j];
                if (compoundPath.pathItems.length > 0 && compoundPath.pathItems[0].clipping) {
                    return compoundPath.pathItems[0].geometricBounds;
                }
            }
            // テキストのマスクなどクリップパスが見つからないときはグループの外接矩形 / e.g. a text mask: fall back to the group's bounds
        }
        return pageItem.geometricBounds;
    }

    /**
     * 各オブジェクトの左上（getLayoutBounds 基準）を控える
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Array<Object>} item / left / top を持つ位置情報の配列
     */
    function snapshotPositions(selectedItems) {
        var positionSnapshot = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = getLayoutBounds(selectedItems[i]);
            positionSnapshot.push({ item: selectedItems[i], left: itemBounds[0], top: itemBounds[1] });
        }
        return positionSnapshot;
    }

    /**
     * 控えた位置情報を複製する（元の控えを書き換えても影響しないように）
     * @param {Array<Object>} positionSnapshot - item / left / top を持つ位置情報の配列
     * @returns {Array<Object>} 複製した位置情報の配列
     */
    function clonePositions(positionSnapshot) {
        var clonedSnapshot = [];
        for (var i = 0; i < positionSnapshot.length; i++) {
            clonedSnapshot.push({ item: positionSnapshot[i].item, left: positionSnapshot[i].left, top: positionSnapshot[i].top });
        }
        return clonedSnapshot;
    }

    /**
     * 控えておいた位置へオブジェクトを戻す
     * @param {Array<Object>} positionSnapshot - item / left / top を持つ位置情報の配列
     * @returns {void}
     */
    function restorePositions(positionSnapshot) {
        for (var i = 0; i < positionSnapshot.length; i++) {
            var snapshotEntry = positionSnapshot[i];
            var currentBounds = getLayoutBounds(snapshotEntry.item);
            snapshotEntry.item.translate(snapshotEntry.left - currentBounds[0], snapshotEntry.top - currentBounds[1]);
        }
    }

    /**
     * 選択の並びから左右間隔の初期値を測る（上→下、同じ行は左→右で並べた先頭2つの間隔）
     * @param {PageItem[]} selectedItems - 対象のオブジェクト（2つ以上）
     * @returns {number} 左右間隔（pt）。負の値は 0
     */
    function measureInitialGapX(selectedItems) {
        // 上端の降順、同じ行（差1pt未満）は左端の昇順 / Sort by top descending and left ascending within same row
        var sortedByPosition = selectedItems.slice().sort(function (itemA, itemB) {
            var boundsA = getLayoutBounds(itemA);
            var boundsB = getLayoutBounds(itemB);
            if (Math.abs(boundsA[1] - boundsB[1]) < 1) {
                return boundsA[0] - boundsB[0]; // same row → compare left
            }
            return boundsB[1] - boundsA[1]; // sort by top descending
        });

        var firstBounds = getLayoutBounds(sortedByPosition[0]);
        var secondBounds = getLayoutBounds(sortedByPosition[1]);
        // ダイアログ初期値では負の値を使わない / no negative value in the dialog
        var initialGapX = secondBounds[0] - firstBounds[2];
        return (initialGapX < 0) ? 0 : initialGapX;
    }

    /**
     * 昇順に並んだ数値配列の、隣接する差分の中央値を返す
     * @param {number[]} sortedAsc - 昇順に並んだ数値配列
     * @returns {number} 隣接差分の中央値。要素が2未満なら0
     */
    function medianAdjacentDiff(sortedAsc) {
        if (!sortedAsc || sortedAsc.length < 2) return 0;
        var diffs = [];
        for (var i = 1; i < sortedAsc.length; i++) {
            diffs.push(Math.abs(sortedAsc[i] - sortedAsc[i - 1]));
        }
        diffs.sort(function (a, b) { return a - b; });
        return diffs[Math.floor(diffs.length / 2)];
    }

    /**
     * 選択オブジェクトの外接境界一覧と、最小の幅・高さを収集する
     * （buildLayoutInfo / buildLayoutInfoForceGrid の共通前処理）
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {{boundsList: Array, minWidth: number, minHeight: number}} {item, bounds} の配列と、最小の幅・高さ
     */
    function collectBoundsList(selectedItems) {
        var boundsList = [];
        var minWidth = Number.MAX_VALUE;
        var minHeight = Number.MAX_VALUE;

        for (var j = 0; j < selectedItems.length; j++) {
            var itemBounds = getLayoutBounds(selectedItems[j]);
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];
            if (itemWidth < minWidth) minWidth = itemWidth;
            if (itemHeight < minHeight) minHeight = itemHeight;
            boundsList.push({ item: selectedItems[j], bounds: itemBounds });
        }
        return { boundsList: boundsList, minWidth: minWidth, minHeight: minHeight };
    }

    /**
     * 行・列をまとめる許容差を返す（最小の幅・高さの半分。1pt 未満にはしない）
     * @param {number} minSize - 選択中の最小の幅または高さ
     * @returns {number} 許容差（pt）
     */
    function getClusterTolerance(minSize) {
        var tolerance = minSize * 0.5;
        return (tolerance < 1) ? 1 : tolerance;
    }

    /**
     * 並べ済みの境界レコードを、座標の近いものどうしでまとめる。
     * 既存のまとまりに入るたびに、まとまりの座標を平均へ更新する
     * @param {Array<Object>} sortedRecords - {item, bounds} の配列（まとめる順に並べたもの）
     * @param {number} boundsIndex - 比べる境界の要素（0 = 左端、1 = 上端）
     * @param {string} centerKey - まとまりの座標を入れるキー（"x" / "y"）
     * @param {number} tolerance - 同じまとまりとみなす許容差
     * @returns {Array<Object>} まとまり（centerKey の座標と members）の配列
     */
    function clusterBoundsRecords(sortedRecords, boundsIndex, centerKey, tolerance) {
        var clusters = [];
        for (var i = 0; i < sortedRecords.length; i++) {
            var boundsRecord = sortedRecords[i];
            var coordinate = boundsRecord.bounds[boundsIndex];
            var isMerged = false;
            for (var c = 0; c < clusters.length; c++) {
                if (Math.abs(clusters[c][centerKey] - coordinate) <= tolerance) {
                    clusters[c].members.push(boundsRecord);
                    clusters[c][centerKey] = (clusters[c][centerKey] * (clusters[c].members.length - 1) + coordinate) / clusters[c].members.length;
                    isMerged = true;
                    break;
                }
            }
            if (!isMerged) {
                var newCluster = { members: [boundsRecord] };
                newCluster[centerKey] = coordinate;
                clusters.push(newCluster);
            }
        }
        return clusters;
    }

    /**
     * 境界レコードの中で最大の幅または高さを返す
     * @param {Array<Object>} boundsRecords - {item, bounds} の配列
     * @param {boolean} measureWidth - true なら幅、false なら高さ
     * @returns {number} 最大の幅または高さ
     */
    function getMaxExtent(boundsRecords, measureWidth) {
        var maxExtent = 0;
        for (var i = 0; i < boundsRecords.length; i++) {
            var recordBounds = boundsRecords[i].bounds;
            var extent = measureWidth ? (recordBounds[2] - recordBounds[0]) : (recordBounds[1] - recordBounds[3]);
            if (extent > maxExtent) maxExtent = extent;
        }
        return maxExtent;
    }

    /**
     * 選択オブジェクトの現在位置から、行と列の構成を推定する
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Object} 行・列の中心座標と、各オブジェクトの行列位置を持つレイアウト情報
     */
    function buildLayoutInfo(selectedItems) {
        var collectedBounds = collectBoundsList(selectedItems);
        var boundsList = collectedBounds.boundsList;

        // 列まとめ（左→右）/ group columns (left -> right)
        boundsList.sort(function (recordA, recordB) { return recordA.bounds[0] - recordB.bounds[0]; });
        var colCenters = clusterBoundsRecords(boundsList, 0, "x", getClusterTolerance(collectedBounds.minWidth));

        // 行まとめ（上→下）/ group rows (top -> bottom)
        var boundsSortedByTop = boundsList.slice().sort(function (recordA, recordB) { return recordB.bounds[1] - recordA.bounds[1]; });
        var rowCenters = clusterBoundsRecords(boundsSortedByTop, 1, "y", getClusterTolerance(collectedBounds.minHeight));

        // 並び順の確定 / sort
        colCenters.sort(function (a, b) { return a.x - b.x; });
        rowCenters.sort(function (a, b) { return b.y - a.y; });

        // 各列の最大幅・各行の最大高さ / max width per column, max height per row
        var colWidths = [];
        for (var i = 0; i < colCenters.length; i++) colWidths.push(getMaxExtent(colCenters[i].members, true));
        var rowHeights = [];
        for (var j = 0; j < rowCenters.length; j++) rowHeights.push(getMaxExtent(rowCenters[j].members, false));

        return {
            boundsList: boundsList,
            colCenters: colCenters,
            rowCenters: rowCenters,
            colWidths: colWidths,
            rowHeights: rowHeights,
            baseX: colCenters[0].x,
            baseY: rowCenters[0].y
        };
    }

    /**
     * 選択オブジェクトを強制的に格子とみなして、行と列の構成を組み立てる。
     * 行ごと（上→下）に左→右で(行,列)を割り当てる：行は上端Yの近さでまとめ、各行の中は左端Xで並べる。
     * 欠け（歯抜け）は許容する（行ごとに列数が異なってよい）
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Object} 行・列の中心座標と、各オブジェクトの行列位置を持つレイアウト情報
     */
    function buildLayoutInfoForceGrid(selectedItems) {
        var collectedBounds = collectBoundsList(selectedItems);
        var boundsList = collectedBounds.boundsList;

        // 行まとめ（上→下）/ group rows (top -> bottom)
        var boundsSortedByTop = boundsList.slice().sort(function (recordA, recordB) { return recordB.bounds[1] - recordA.bounds[1]; });
        var rowClusters = clusterBoundsRecords(boundsSortedByTop, 1, "y", getClusterTolerance(collectedBounds.minHeight));
        rowClusters.sort(function (a, b) { return b.y - a.y; });

        // 各行の中を左→右で確定し、rowIndex/colIndexを付与 / sort within row and assign indices
        var maxCols = 0;
        for (var r = 0; r < rowClusters.length; r++) {
            var rowMembers = rowClusters[r].members;
            rowMembers.sort(function (recordA, recordB) { return recordA.bounds[0] - recordB.bounds[0]; });
            if (rowMembers.length > maxCols) maxCols = rowMembers.length;
            for (var c = 0; c < rowMembers.length; c++) {
                rowMembers[c].rowIndex = r;
                rowMembers[c].colIndex = c;
            }
        }

        // 列番号ごとの最大幅・行ごとの最大高さ / max width per column index, max height per row
        var colWidths = [];
        for (var colIndex = 0; colIndex < maxCols; colIndex++) {
            var maxWidth = 0;
            for (var rowIndex = 0; rowIndex < rowClusters.length; rowIndex++) {
                var memberInColumn = rowClusters[rowIndex].members[colIndex];
                if (memberInColumn) {
                    var memberWidth = memberInColumn.bounds[2] - memberInColumn.bounds[0];
                    if (memberWidth > maxWidth) maxWidth = memberWidth;
                }
            }
            colWidths.push(maxWidth);
        }
        var rowHeights = [];
        for (var i = 0; i < rowClusters.length; i++) rowHeights.push(getMaxExtent(rowClusters[i].members, false));

        // 基準は最も左の左端と最も上の上端 / origin = leftmost left and topmost top
        // （Illustratorの座標はアートボードより下で負になるので、上端の初期値は -MAX_VALUE）
        // (y is negative below the artboard origin, so start the top at -MAX_VALUE)
        var baseX = Number.MAX_VALUE;
        var baseY = -Number.MAX_VALUE;
        for (var j = 0; j < boundsList.length; j++) {
            if (boundsList[j].bounds[0] < baseX) baseX = boundsList[j].bounds[0];
            if (boundsList[j].bounds[1] > baseY) baseY = boundsList[j].bounds[1];
        }

        // 列・行センター（列は割り当て済みの colIndex を使うのでダミー）/ centers (columns are dummies; colIndex is pre-assigned)
        var colCenters = [];
        for (var k = 0; k < maxCols; k++) colCenters.push({ x: baseX, members: [] });

        return {
            boundsList: boundsList,
            colCenters: colCenters,
            rowCenters: rowClusters,
            colWidths: colWidths,
            rowHeights: rowHeights,
            baseX: baseX,
            baseY: baseY
        };
    }

    /**
     * ［強制グリッド］の状態に合わせてレイアウト情報を作る
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Object} レイアウト情報
     */
    function buildCurrentLayoutInfo(selectedItems) {
        return ($.global.__regridForceGrid) ? buildLayoutInfoForceGrid(selectedItems) : buildLayoutInfo(selectedItems);
    }

    /**
     * 対象の列・行インデックスを解決する。
     * Force Grid で事前割り当て済み（colIndex/rowIndex）ならそれを優先し、
     * なければ列／行センターへの最近傍で推定する
     * @param {object} currentEntry - layoutInfo.boundsList の要素
     * @param {Array} colCenters - 列センター配列
     * @param {Array} rowCenters - 行センター配列
     * @returns {{col: number, row: number}} 列・行インデックス
     */
    function resolveColRow(currentEntry, colCenters, rowCenters) {
        var colIndex = (typeof currentEntry.colIndex === 'number') ? currentEntry.colIndex : null;
        var rowIndex = (typeof currentEntry.rowIndex === 'number') ? currentEntry.rowIndex : null;

        if (colIndex === null) {
            colIndex = 0;
            var minDistanceX = Number.MAX_VALUE;
            for (var c = 0; c < colCenters.length; c++) {
                var dx = Math.abs(colCenters[c].x - currentEntry.bounds[0]);
                if (dx < minDistanceX) { minDistanceX = dx; colIndex = c; }
            }
        }

        if (rowIndex === null) {
            rowIndex = 0;
            var minDistanceY = Number.MAX_VALUE;
            for (var r = 0; r < rowCenters.length; r++) {
                var dy = Math.abs(rowCenters[r].y - currentEntry.bounds[1]);
                if (dy < minDistanceY) { minDistanceY = dy; rowIndex = r; }
            }
        }

        return { col: colIndex, row: rowIndex };
    }

    /**
     * レンガ／ハニカムで使う半ピッチを算出する。
     * 目標列ピッチ＝中央値幅 + gapX。推定できない場合は列センター間隔の中央値を使う
     * @param {number[]} colWidths - 列ごとの最大幅
     * @param {Array} colCenters - 列センター配列
     * @param {number} gapX - 左右間隔
     * @returns {number} 半ピッチ（pitch / 2）
     */
    function computeHalfPitch(colWidths, colCenters, gapX) {
        var sortedColWidths = [];
        if (colWidths && colWidths.length > 0) {
            for (var i = 0; i < colWidths.length; i++) sortedColWidths.push(colWidths[i]);
            sortedColWidths.sort(function (a, b) { return a - b; });
        }
        var medianColWidth = (sortedColWidths.length > 0) ? sortedColWidths[Math.floor(sortedColWidths.length / 2)] : 0; // median width
        var pitch = (medianColWidth > 0) ? (medianColWidth + gapX) : 0;

        // fallback：列が1つ等で推定できない場合は現状の列位置差 / fallback to current centers diff
        if (pitch === 0) {
            var colLefts = [];
            for (var j = 0; j < colCenters.length; j++) colLefts.push(colCenters[j].x);
            colLefts.sort(function (a, b) { return a - b; });
            pitch = medianAdjacentDiff(colLefts);
        }
        return pitch / 2.0;
    }

    /**
     * グリッド配置を適用する共通処理。通常／レンガ／ハニカムを引数で切り替える。
     * 直前にプレビュー基準位置へ戻してから、列幅・行高さの累積で再配置する。
     * layoutInfo は baselinePositions と同時に作るので、boundsList の左上がそのまま基準位置になる
     * @param {Object} layoutInfo - レイアウト情報
     * @param {Array<Object>} baselinePositions - プレビュー基準の位置情報
     * @param {number} gapX - 左右間隔
     * @param {number} gapY - 上下間隔
     * @param {boolean} isBrick - 奇数行を半ピッチ横にずらすか（レンガ／ハニカム）
     * @param {number} rowStepFactor - 行送りに掛ける係数（通常・レンガ=1.0、ハニカム=HONEYCOMB_ROW_STEP_FACTOR）
     * @returns {void}
     */
    function applyGridLayout(layoutInfo, baselinePositions, gapX, gapY, isBrick, rowStepFactor) {
        // いったん基準位置に戻す / restore the baseline first
        restorePositions(baselinePositions);

        var boundsList = layoutInfo.boundsList;
        var colWidths = layoutInfo.colWidths;
        var rowHeights = layoutInfo.rowHeights;

        var halfPitch = isBrick ? computeHalfPitch(colWidths, layoutInfo.colCenters, gapX) : 0;

        // 中央揃え：各セルの天地左右中央に整列（強制グリッドのサブオプションなので Force Grid 時のみ有効）
        // Center each object within its cell (sub-option of Force Grid, so only when Force Grid is on)
        var centerInCell = !!$.global.__regridCenterInCell && !!$.global.__regridForceGrid;

        for (var i = 0; i < boundsList.length; i++) {
            var currentEntry = boundsList[i];

            var cellIndex = resolveColRow(currentEntry, layoutInfo.colCenters, layoutInfo.rowCenters);
            var colIndex = cellIndex.col;
            var rowIndex = cellIndex.row;

            // 新しいX（セル左端）/ new X (cell left)
            var newX = layoutInfo.baseX;
            for (var c = 0; c < colIndex; c++) {
                newX += colWidths[c] + gapX;
            }

            // 新しいY（セル上端）/ new Y (cell top)
            var newY = layoutInfo.baseY;
            for (var r = 0; r < rowIndex; r++) {
                newY -= (rowHeights[r] * rowStepFactor) + gapY;
            }

            // 中央揃え：セル内でオブジェクトを左右・天地中央へ寄せる
            // Center within the cell (cell size = column width × row height)
            if (centerInCell) {
                var itemWidth = currentEntry.bounds[2] - currentEntry.bounds[0];
                var itemHeight = currentEntry.bounds[1] - currentEntry.bounds[3];
                newX += (colWidths[colIndex] - itemWidth) / 2;
                newY -= (rowHeights[rowIndex] - itemHeight) / 2;
            }

            // レンガ状：奇数行を半ピッチずらす / Brick: shift odd rows by half pitch
            if (isBrick && halfPitch !== 0 && (rowIndex % 2 === 1)) {
                newX += halfPitch;
            }

            currentEntry.item.translate(newX - currentEntry.bounds[0], newY - currentEntry.bounds[1]);
        }

        // 再描画 / redraw
        app.redraw();
    }

    // =========================================
    // 行列入れ替え / Transpose
    // getLayoutBounds() の left/top を基準に、行・列をクラスタリングして推定し、
    // 推定したピッチ（隣接差の中央値）で左上基準に再配置する（歯抜け対応）。
    // グループは中身を見ず、グループ全体の外接 bbox（クリップグループはクリップパス）で1オブジェクトとして扱う。
    // 1行→1列 / 1列→1行 も対応（ピッチ流用）
    // =========================================

    /**
     * 近い値どうしをまとめて、クラスタの中心値の配列を作る（coordinateValues は昇順に並べ替わる）
     * @param {number[]} coordinateValues - まとめる対象の値
     * @param {number} tolerance - 同じクラスタとみなす許容差
     * @returns {number[]} 昇順に並べたクラスタ中心値の配列
     */
    function clusterValues(coordinateValues, tolerance) {
        coordinateValues.sort(function (a, b) { return a - b; });
        var clusterCenters = [];
        for (var i = 0; i < coordinateValues.length; i++) {
            var coordinate = coordinateValues[i];
            var foundIndex = -1;
            for (var c = 0; c < clusterCenters.length; c++) {
                if (Math.abs(coordinate - clusterCenters[c]) <= tolerance) { foundIndex = c; break; }
            }
            if (foundIndex < 0) clusterCenters.push(coordinate);
            else clusterCenters[foundIndex] = (clusterCenters[foundIndex] + coordinate) / 2.0;
        }
        clusterCenters.sort(function (a, b) { return a - b; });
        return clusterCenters;
    }

    /**
     * 中心値の配列から、指定した値に最も近い要素の位置を返す
     * @param {number[]} sortedCenters - 並べた中心値の配列
     * @param {number} coordinate - 探す値
     * @returns {number} 最も近い要素のインデックス
     */
    function findNearestIndex(sortedCenters, coordinate) {
        var bestIndex = 0;
        var bestDistance = Math.abs(coordinate - sortedCenters[0]);
        for (var i = 1; i < sortedCenters.length; i++) {
            var distance = Math.abs(coordinate - sortedCenters[i]);
            if (distance < bestDistance) { bestDistance = distance; bestIndex = i; }
        }
        return bestIndex;
    }

    /**
     * 各オブジェクトを(行,列)に割り当てる。同じセルに2つ入ったらアラートを出して中止する
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @param {number[]} rowClusters - 行の中心値（上→下）
     * @param {number[]} colClusters - 列の中心値（左→右）
     * @returns {Array<Object>|null} {item, row, col} の配列。衝突したら null
     */
    function mapItemsToCells(selectedItems, rowClusters, colClusters) {
        var occupiedCells = {}; // cellKey "row,col" -> item
        var cellMapping = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var targetItem = selectedItems[i];
            var itemBounds = getLayoutBounds(targetItem);
            var rowIndex = findNearestIndex(rowClusters, itemBounds[1]);
            var colIndex = findNearestIndex(colClusters, itemBounds[0]);
            var cellKey = rowIndex + "," + colIndex;
            if (occupiedCells[cellKey]) {
                // 行・列は1から数えて表示 / show row and column counting from 1
                alert(getLabel('alert.cellConflict').replace("{row}", rowIndex + 1).replace("{col}", colIndex + 1));
                return null;
            }
            occupiedCells[cellKey] = targetItem;
            cellMapping.push({ item: targetItem, row: rowIndex, col: colIndex });
        }
        return cellMapping;
    }

    /**
     * 転置後に使う横・縦のピッチを求める（元の列・行の隣接差の中央値）。
     * 1行しかないときは横ピッチを縦にも、1列しかないときは縦ピッチを横にも流用する
     * @param {number[]} rowClusters - 行の中心値
     * @param {number[]} colClusters - 列の中心値
     * @returns {{x: number, y: number}|null} ピッチ。推定できなければ null
     */
    function getTransposePitch(rowClusters, colClusters) {
        var rowCount = rowClusters.length;
        var colCount = colClusters.length;
        // ※行や列が1つしかない場合は0になる / 0 when there is only one row or column
        var pitchX = (colCount >= 2) ? medianAdjacentDiff(colClusters) : 0;
        var pitchY = (rowCount >= 2) ? medianAdjacentDiff(rowClusters) : 0;

        // ほぼ重なり等で1セル扱いになったケース / everything collapsed into one cell
        if (rowCount === 1 && colCount === 1) return null;

        if (rowCount === 1 && colCount > 1) {
            // 1行 → 1列：横ピッチを縦へ流用 / one row -> one column
            return (pitchX === 0) ? null : { x: pitchX, y: pitchX };
        }
        if (colCount === 1 && rowCount > 1) {
            // 1列 → 1行：縦ピッチを横へ流用 / one column -> one row
            return (pitchY === 0) ? null : { x: pitchY, y: pitchY };
        }
        // 通常（2行以上 かつ 2列以上）/ two or more rows and columns
        return (pitchX === 0 || pitchY === 0) ? null : { x: pitchX, y: pitchY };
    }

    /**
     * 歯抜けを許容したままグリッドを転置（行⇄列）する
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {void}
     */
    function transposeGridWithHoles(selectedItems) {
        if (!selectedItems || selectedItems.length < 1) return;

        // left/top を集める / collect left/top
        var leftValues = [], topValues = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = getLayoutBounds(selectedItems[i]);
            leftValues.push(itemBounds[0]);
            topValues.push(itemBounds[1]);
        }

        var colClusters = clusterValues(leftValues, TRANSPOSE_SNAP_X_TOLERANCE); // left -> right
        var rowClusters = clusterValues(topValues, TRANSPOSE_SNAP_Y_TOLERANCE); // will sort top -> bottom next

        // Illustrator座標では上ほどYが大きいことが多いので「上→下」/ sort top -> bottom
        rowClusters.sort(function (a, b) { return b - a; });

        var cellMapping = mapItemsToCells(selectedItems, rowClusters, colClusters);
        if (!cellMapping) return;

        var pitch = getTransposePitch(rowClusters, colClusters);
        if (!pitch) return;

        // 転置後グリッドの基準（左上固定）/ origin at top-left
        var originLeft = colClusters[0];
        var originTop = rowClusters[0];

        // 転置: 新しい列 = 元の行、新しい行 = 元の列 / newCol = oldRow, newRow = oldCol
        for (var j = 0; j < cellMapping.length; j++) {
            var mappedItem = cellMapping[j].item;
            var targetLeft = originLeft + cellMapping[j].row * pitch.x;
            var targetTop = originTop - cellMapping[j].col * pitch.y;

            var currentBounds = getLayoutBounds(mappedItem);
            mappedItem.translate(targetLeft - currentBounds[0], targetTop - currentBounds[1]);
        }

        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを収集し、間隔設定ダイアログを起動する
     * @returns {void}
     */
    function main() {
        // ドキュメントチェック / document check
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel('alert.noSelection'));
            return;
        }

        // 選択を拾う / collect selection
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }
        if (selectedItems.length < 2) {
            alert(getLabel('alert.needTwo'));
            return;
        }

        var initialGapX = measureInitialGapX(selectedItems);

        // プレビューの基準位置と、ダイアログ開始時点の位置（キャンセルで必ずここへ戻す）
        // Preview baseline, and the snapshot Cancel always returns to
        var baselinePositions = snapshotPositions(selectedItems);
        var initialPositions = clonePositions(baselinePositions);

        var layoutInfo = buildCurrentLayoutInfo(selectedItems);

        /**
         * 現在の位置を新しい基準として控え直し、レイアウト情報を作り直す
         * @returns {void}
         */
        function resetBaselineToCurrent() {
            baselinePositions = snapshotPositions(selectedItems);
            layoutInfo = buildCurrentLayoutInfo(selectedItems);
        }

        // ダイアログ表示 / show dialog
        showGridSpacingDialog({
            /* 間隔を適用（通常／レンガ状／ハニカム）/ apply spacing (normal / brick / honeycomb) */
            applySpacing: function (gapX, gapY, isBrick, rowStepFactor) {
                applyGridLayout(layoutInfo, baselinePositions, gapX, gapY, isBrick, rowStepFactor);
            },
            /* 行列入れ替え / swap rows and columns */
            transpose: function () {
                transposeGridWithHoles(selectedItems);
            },
            /* キャンセル時に戻す（ダイアログ開始時点）/ restore initial positions on Cancel */
            restoreInitialPositions: function () {
                restorePositions(initialPositions);
            },
            resetBaselineToCurrent: resetBaselineToCurrent
        }, initialGapX);
    }

    // 実行 / run
    main();

})();
