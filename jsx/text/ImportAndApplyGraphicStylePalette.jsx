#target illustrator
#targetengine "ImportAndApplyGraphicStyle"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換します。
あわせて、指定したファイルからグラフィックスタイルを取り込んで適用できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportAndApplyGraphicStylePalette.md

### Overview

Converts point text, text on a path, or a shape plus text into area text while preserving the appearance.
A graphic style can be imported from a file you choose and applied at the same time.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportAndApplyGraphicStylePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportAndApplyGraphicStylePalette";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportAndApplyGraphicStylePalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportAndApplyGraphicStylePalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 大きさ調整「する」の既定倍率（幅・高さ）/ Default size-adjustment ratios (width, height) */
    var BUTTON_WIDTH_RATIO = 1.2;   /* 元の幅に対する倍率 / Ratio of original width */
    var BUTTON_HEIGHT_RATIO = 1.6;  /* 元の高さに対する倍率 / Ratio of original height */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PALETTE_MARGINS        = 15;                /* パレットの余白 / palette margins */
    var PANEL_MARGINS          = [16, 20, 16, 12];  /* パネルの余白 / panel margins */
    var PANEL_SPACING          = 8;                 /* パネル内の既定の間隔 / default panel spacing */
    var INNER_PANEL_SPACING    = 6;                 /* このパレットのパネル内の間隔 / spacing inside this palette's panels */
    var SIZE_LABEL_WIDTH       = 44;                /* 「サイズ」ラベルの幅 / width of the Size label */
    var RATIO_INPUT_WIDTH      = 40;                /* 幅・高さの入力欄の幅 / width of the ratio fields */
    var STYLE_LIST_HEIGHT      = 120;               /* スタイル一覧の高さ / height of the style list */
    var STYLE_FOOTER_MARGINS   = [0, 5, 0, 0];      /* スタイルパネル下部のボタン行の余白 / margins of the style panel's button row */
    var FILE_NAME_WIDTH        = 240;               /* ファイル名表示の幅 / width of the file name label */
    var LOAD_BUTTON_ROW_MARGINS = [0, 7, 0, 0];     /* 読み込みボタン行の余白 / margins of the load button row */
    var BUTTON_HEIGHT_TRIM     = 2;                 /* パネル内のボタンの高さを詰める量（px）/ px trimmed from panel buttons */

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は PANEL_SPACING）
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
     * ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無いことがある / size may be unavailable before layout */
        try {
            targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
        } catch (e) { }
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
    // 設定ファイル / Preferences file
    // =========================================

    /* 参照した AI ファイルとスタイル名を記憶する設定ファイル / Prefs file remembering the picked AI file and its style names */
    var PREFS_FILE_NAME = "styles_for_TextWithShapeToAreaType.txt";

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

    /* 日英ラベル定義（カテゴリ構造）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "エリア内文字に変換", en: "Convert to Area Type" },
            pickFile: { ja: "スタイルの AI ファイルを選択", en: "Select a style AI file" }
        },
        panel: {
            areaTypeOption: { ja: "エリア内文字オプション", en: "Area Type Options" },
            graphicStyle: { ja: "グラフィックスタイル", en: "Graphic style" },
            loadStyles: { ja: "スタイルの読み込み", en: "Load Styles" }
        },
        checkbox: {
            convertToAreaType: { ja: "エリア内文字に変換", en: "Convert to area type" },
            usedOnly: { ja: "ドキュメント内で使用しているもののみ", en: "Only those used in the document" }
        },
        radio: {
            doAdjust: { ja: "調整する", en: "Adjust" },
            dontAdjust: { ja: "調整しない", en: "Don't adjust" }
        },
        fieldLabel: {
            size: { ja: "サイズ", en: "Size" },
            widthRatio: { ja: "幅", en: "Width" },
            heightRatio: { ja: "高さ", en: "Height" }
        },
        listItem: {
            styleOriginal: { ja: "元の見た目", en: "Original appearance" }
        },
        status: {
            noFileSelected: { ja: "ファイル未選択", en: "No file selected" }
        },
        button: {
            runConvert: { ja: "変換", en: "Convert" },
            clearStyle: { ja: "クリア", en: "Clear" },
            openStylePanel: { ja: "パネル", en: "Panel" },
            load: { ja: "読み込み", en: "Load" },
            reload: { ja: "再読み込み", en: "Reload" }
        },
        tooltip: {
            convertToAreaType: {
                ja: "オンにすると、［変換］でポイント文字・パス上文字（またはテキストと長方形）をエリア内文字に変換します。オフのときは選んだグラフィックスタイルを適用するだけです。",
                en: "When on, Convert turns point text, path text, or text plus a rectangle into area type. When off, Convert only applies the chosen graphic style."
            },
            doAdjust: {
                ja: "ポイント文字を変換するとき、実寸に幅・高さの倍率を掛けた枠にし、元の見た目ならボタン状の背景を付けます。",
                en: "When converting point text, sizes the frame to the measured size times the width and height ratios, and adds a button-shaped background for the original appearance."
            },
            dontAdjust: { ja: "ポイント文字を変換するとき、実寸と同じ大きさの枠にします。", en: "When converting point text, makes the frame the same size as the text." },
            widthRatio: { ja: "枠の幅を、テキストの実寸に対する割合（%）で指定します。", en: "Frame width as a percentage of the text's measured width." },
            heightRatio: { ja: "枠の高さを、テキストの実寸に対する割合（%）で指定します。", en: "Frame height as a percentage of the text's measured height." },
            styleList: {
                ja: "選ぶと、選択中のオブジェクトにすぐ適用します。「元の見た目」は何も適用しません。",
                en: "Picking a style applies it to the selected objects right away. “Original appearance” applies nothing."
            },
            runConvert: {
                ja: "［エリア内文字に変換］がオンなら選択をエリア内文字に変換し、オフなら選んだグラフィックスタイルを選択に適用します。",
                en: "Converts the selection to area type when “Convert to area type” is on; otherwise applies the chosen graphic style to the selection."
            },
            openStylePanel: { ja: "グラフィックスタイルパネルに移動", en: "Go to the Graphic Styles panel." },
            clearStyle: { ja: "未使用のグラフィックスタイルを削除", en: "Delete unused graphic styles." },
            load: {
                ja: "「読み込み」でスタイルの AI ファイルを選択してください。",
                en: "Click “Load” to choose a style AI file."
            },
            reload: { ja: "記憶したファイルからスタイルを取り込み直します。", en: "Re-import styles from the remembered file." },
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
            }
        }
    };

    /**
     * "category.key" 形式のキー、または { ja, en } のラベルから現在の言語の文字列を取得する
     * @param {string|Object} labelPath - ラベルキー（例: "panel.graphicStyle"）か { ja: string, en: string }
     * @returns {string} 現在の言語のラベル文字列（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        /* ステップボタンは LABELS.tooltip.stepUp などをオブジェクトで渡す / the stepper passes label objects */
        if (typeof labelPath !== "string") return labelPath[uiLang] || labelPath.en;
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            if (labelNode && typeof labelNode[labelPathKeys[i]] !== "undefined") { labelNode = labelNode[labelPathKeys[i]]; }
            else { return labelPath; }
        }
        if (labelNode[uiLang]) return labelNode[uiLang];
        if (labelNode.en) return labelNode.en;
        return labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelSet - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // BridgeTalk 委譲基盤 / BridgeTalk delegation
    // 常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理は worker 関数に集約し、
    // 押下のたびにメインエンジン（illustrator）へ同期委譲する。
    // A resident palette's app loses its DOM connection while shown, so DOM work lives in worker
    // functions and is delegated synchronously to the main engine (illustrator) on each action.
    // =========================================

    /* 戻り値マーカー / Result markers */
    var MARKER_OK = "OK";
    var MARKER_NODOC = "NODOC";
    var MARKER_NOSEL = "NOSEL";
    var MARKER_ERR = "ERR";

    /* 委譲する worker 関数を全登録（追加漏れ防止）。
       worker 関数内は「// 行コメント禁止・/* *\/ のみ・必ずセミコロンで終える」（toString が改行を消すため）。
       Register every worker function here (avoid omissions). Inside workers: no // comments, /* *\/ only,
       always end statements with a semicolon (toString strips newlines). */
    var WORKER_FUNCS = [];

    /* 再入防止ガード / Re-entrancy guard */
    var isBusy = false;

    /**
     * 登録された worker 関数を連結して worker ソースを生成する
     * @returns {string} worker 関数のソース
     */
    function buildWorkerSource() {
        var workerSource = "";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            workerSource += WORKER_FUNCS[i].toString() + "\n";
        }
        return workerSource;
    }

    /**
     * worker 呼び出し式をメインエンジンへ同期委譲し、戻り値の文字列を返す
     * @param {string} callExpr - 文字列を返す式（例: 'workerOpenStylePanel()'）
     * @returns {string} worker の戻り値（失敗時は "ERR:..."）
     */
    function delegate(callExpr) {
        if (isBusy) return MARKER_ERR + ":busy";
        isBusy = true;
        var resultHolder = { result: null };
        /* BridgeTalk の送信は失敗しうる / Sending through BridgeTalk can fail */
        try {
            var payload = buildWorkerSource() + "\n" + callExpr + ";";
            var bridgeTalk = new BridgeTalk();
            bridgeTalk.target = "illustrator";
            /* バックスラッシュ・多バイト・改行の破損を避けるため encodeURIComponent で包む
               Wrap via encodeURIComponent to avoid backslash / multibyte / newline corruption */
            bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(payload) + "\"));";
            bridgeTalk.onResult = function (message) { resultHolder.result = message.body; };
            bridgeTalk.onError = function (message) { resultHolder.result = MARKER_ERR + ":" + message.body; };
            bridgeTalk.send(10); /* 同期送信 / Synchronous send */
        } catch (e) {
            resultHolder.result = MARKER_ERR + ":" + e;
        } finally {
            isBusy = false;
        }
        return (resultHolder.result === null) ? (MARKER_ERR + ":noresult") : resultHolder.result;
    }

    /**
     * 文字列を委譲用の JS 文字列リテラルへエスケープする
     * @param {string} text - エスケープする文字列
     * @returns {string} ダブルクォートで囲んだ文字列リテラル
     */
    function jsStringLiteral(text) {
        var sourceText = String(text);
        var literalText = '"';
        for (var i = 0; i < sourceText.length; i++) {
            var character = sourceText.charAt(i);
            if (character === '\\') literalText += '\\\\';
            else if (character === '"') literalText += '\\"';
            else if (character === '\n') literalText += '\\n';
            else if (character === '\r') literalText += '\\r';
            else if (character === '\t') literalText += '\\t';
            else literalText += character;
        }
        return literalText + '"';
    }

    /**
     * worker の "OK\t<改行連結>" 形式の戻り値を解析する
     * @param {string} resultText - worker の戻り値
     * @returns {{ok: boolean, marker: string, items: string[]}} 解析結果
     */
    function parseMarkerList(resultText) {
        if (!resultText) return { ok: false, marker: MARKER_ERR, items: [] };
        var tabIndex = resultText.indexOf("\t");
        var marker = (tabIndex >= 0) ? resultText.substring(0, tabIndex) : resultText;
        if (marker !== MARKER_OK) return { ok: false, marker: marker, items: [] };
        var listText = (tabIndex >= 0) ? resultText.substring(tabIndex + 1) : "";
        var listItems = listText.length ? listText.split("\n") : [];
        return { ok: true, marker: MARKER_OK, items: listItems };
    }

    // ---- worker 関数 / Worker functions（メインエンジンで実行 / run in the main engine）----
    // toString() で送るため JSDoc を付けず、説明は関数の外の1行コメントにする
    // Serialized with toString(), so no JSDoc: each worker is described by a one-line comment outside it

    /* worker: グラフィックスタイルパネルを表示 / Show the Graphic Styles panel */
    function workerOpenStylePanel() {
        try {
            app.executeMenuCommand("Adobe Style Palette");
            return "OK";
        } catch (e) {
            return "ERR:" + e;
        }
    }
    WORKER_FUNCS.push(workerOpenStylePanel);

    /* worker: 文字列を16進へ / Convert a string to hex */
    function workerHexAscii(text) {
        var hexText = "";
        var i;
        for (i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) { hexPair = "0" + hexPair; }
            hexText += hexPair;
        }
        return hexText;
    }
    WORKER_FUNCS.push(workerHexAscii);

    /* worker: メニューコマンド1件のイベントブロックを組み立て / Build one menu-command event block */
    function workerMenuEventBlock(eventIndex, internalName, localizedNameHex, commandNameHex, value, hasDialog) {
        var blockLines = [];
        blockLines.push("\t/event-" + eventIndex + " {");
        blockLines.push("\t\t/useRulersIn1stQuadrant 1");
        blockLines.push("\t\t/internalName (" + internalName + ")");
        blockLines.push("\t\t/localizedName [ " + (localizedNameHex.length / 2));
        blockLines.push("\t\t\t" + localizedNameHex);
        blockLines.push("\t\t]");
        blockLines.push("\t\t/isOpen 0");
        blockLines.push("\t\t/isOn 1");
        blockLines.push("\t\t/hasDialog " + (hasDialog ? "1" : "0"));
        if (hasDialog) { blockLines.push("\t\t/showDialog 0"); }
        blockLines.push("\t\t/parameterCount 1");
        blockLines.push("\t\t/parameter-1 {");
        blockLines.push("\t\t\t/key 1835363957");
        blockLines.push("\t\t\t/showInPalette 1");
        blockLines.push("\t\t\t/type (enumerated)");
        blockLines.push("\t\t\t/name [ " + (commandNameHex.length / 2));
        blockLines.push("\t\t\t\t" + commandNameHex);
        blockLines.push("\t\t\t]");
        blockLines.push("\t\t\t/value " + value);
        blockLines.push("\t\t}");
        blockLines.push("\t}");
        return blockLines.join("\n");
    }
    WORKER_FUNCS.push(workerMenuEventBlock);

    /* worker: グラフィックスタイル用「未使用をすべて選択 → 削除」の一時アクション定義を組み立て
       Build the "Select All Unused -> Delete" action for graphic styles */
    function workerBuildPruneAction(setName, actionName) {
        var internalName = "ai_plugin_styles";
        var localizedNameHex = "e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab";
        var selectValue = 14;
        var deleteValue = 3;
        var deleteNameHex = "44656c657465205374796c65";
        var selectAllUnusedHex = "53656c65637420416c6c20556e75736564";
        var aiaParts = [];
        aiaParts.push("/version 3");
        aiaParts.push("/name [ " + setName.length);
        aiaParts.push("\t" + workerHexAscii(setName));
        aiaParts.push("]");
        aiaParts.push("/isOpen 1");
        aiaParts.push("/actionCount 1");
        aiaParts.push("/action-1 {");
        aiaParts.push("\t/name [ " + actionName.length);
        aiaParts.push("\t\t" + workerHexAscii(actionName));
        aiaParts.push("\t]");
        aiaParts.push("\t/keyIndex 0");
        aiaParts.push("\t/colorIndex 0");
        aiaParts.push("\t/isOpen 1");
        aiaParts.push("\t/eventCount 2");
        aiaParts.push(workerMenuEventBlock(1, internalName, localizedNameHex, selectAllUnusedHex, selectValue, false));
        aiaParts.push(workerMenuEventBlock(2, internalName, localizedNameHex, deleteNameHex, deleteValue, true));
        aiaParts.push("}");
        return aiaParts.join("\n");
    }
    WORKER_FUNCS.push(workerBuildPruneAction);

    /* worker: 失敗を無視する後始末 / Best-effort cleanup, ignoring failures */
    function workerIgnoringErrors(fn) {
        try { fn(); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerIgnoringErrors);

    /* worker: アクションを一時ファイルに書き出して再生 / Write, load, and play the action */
    function workerPlayTempAction(actionSource, setName, actionName, fileName) {
        var actionFile = new File(fileName);
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) { return false; }
            actionFile.write(actionSource);
            actionFile.close();
            workerIgnoringErrors(function () { app.unloadAction(setName, ""); });
            app.loadAction(actionFile);
            app.doScript(actionName, setName);
        } catch (e) {
            /* ignore */
        }
        workerIgnoringErrors(function () { actionFile.close(); });
        workerIgnoringErrors(function () { app.unloadAction(setName, ""); });
        workerIgnoringErrors(function () { actionFile.remove(); });
        return true;
    }
    WORKER_FUNCS.push(workerPlayTempAction);

    /* worker: 未使用グラフィックスタイルをアクションで削除し、削除件数を返す / Prune unused graphic styles, return the count */
    function workerPruneUnusedCore(doc) {
        var countBefore = doc.graphicStyles.length;
        var source = workerBuildPruneAction("TemporaryActionSet", "TemporaryActionName");
        workerPlayTempAction(source, "TemporaryActionSet", "TemporaryActionName", Folder.temp + "/ImportAndApplyGraphicStyle_prune.aia");
        try { app.redraw(); } catch (e) { /* ignore */ }
        var removed = countBefore - doc.graphicStyles.length;
        if (removed < 0) { removed = 0; }
        return removed;
    }
    WORKER_FUNCS.push(workerPruneUnusedCore);

    /* worker: 未使用スタイル名の集合を取得（未使用削除→残りから判定→削除があれば undo で復元）
       Get the set of unused style names (delete unused, classify by what remains, undo only if deleted) */
    function workerComputeUnusedSet(doc) {
        var unusedNames = {};
        var namesBefore = [];
        var i;
        for (i = 1; i < doc.graphicStyles.length; i++) { namesBefore.push(doc.graphicStyles[i].name); }
        if (namesBefore.length === 0) { return unusedNames; }
        var lengthBefore = doc.graphicStyles.length;
        workerPruneUnusedCore(doc);
        var remainingNames = {};
        var j;
        for (j = 1; j < doc.graphicStyles.length; j++) { remainingNames[doc.graphicStyles[j].name] = true; }
        var k;
        for (k = 0; k < namesBefore.length; k++) {
            if (!remainingNames[namesBefore[k]]) { unusedNames[namesBefore[k]] = true; }
        }
        if (doc.graphicStyles.length < lengthBefore) {
            try { app.undo(); app.redraw(); } catch (e) { /* ignore */ }
        }
        return unusedNames;
    }
    WORKER_FUNCS.push(workerComputeUnusedSet);

    /* worker: 現在のドキュメントのグラフィックスタイル名を返す（既定と temp_style を除外、usedOnly で未使用も除外）
       戻り値: "OK\t" + 名前を \n 連結 / "NODOC"
       Return the document's graphic-style names (skip default and temp_style; usedOnly also skips unused).
       Returns "OK\t" + names joined by \n, or "NODOC" */
    function workerGetStyleNames(usedOnly) {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var unusedNames = {};
        if (usedOnly) { unusedNames = workerComputeUnusedSet(doc); }
        var styleNames = [];
        var i;
        for (i = 1; i < doc.graphicStyles.length; i++) {
            var styleName = doc.graphicStyles[i].name;
            if (styleName === "temp_style") { continue; }
            if (usedOnly && unusedNames[styleName]) { continue; }
            styleNames.push(styleName);
        }
        return "OK\t" + styleNames.join("\n");
    }
    WORKER_FUNCS.push(workerGetStyleNames);

    /* worker: 未使用グラフィックスタイルを削除。戻り値: "OK\t" + 件数 / "NODOC"
       Prune unused graphic styles. Returns "OK\t" + count, or "NODOC" */
    function workerPruneUnused() {
        if (app.documents.length === 0) { return "NODOC"; }
        var removed = workerPruneUnusedCore(app.activeDocument);
        return "OK\t" + removed;
    }
    WORKER_FUNCS.push(workerPruneUnused);

    /* worker: 指定 AI ファイルからグラフィックスタイルを現在のドキュメントへ取り込む
       戻り値: "OK\t" + 実際に登録された名前を \n 連結 / "NOFILE" / "NODOC"
       Import graphic styles from the AI file into the current document.
       Returns "OK\t" + registered names joined by \n, or "NOFILE" / "NODOC" */
    function workerImportStyles(filePath) {
        var styleFile = new File(filePath);
        if (!styleFile.exists) { return "NOFILE"; }
        if (app.documents.length === 0) { return "NODOC"; }
        var destinationDoc = app.activeDocument;
        var styleSourceDoc = app.open(styleFile);
        var sourceStyleNames = [];
        var i;
        for (i = 1; i < styleSourceDoc.graphicStyles.length; i++) { sourceStyleNames.push(styleSourceDoc.graphicStyles[i].name); }
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = destinationDoc;
        var importLayerName = "// _imported";
        var importLayer;
        try { importLayer = destinationDoc.layers.getByName(importLayerName); }
        catch (e) { importLayer = destinationDoc.layers.add(); importLayer.name = importLayerName; }
        importLayer.locked = false;
        importLayer.visible = true;
        destinationDoc.activeLayer = importLayer;
        app.executeMenuCommand("paste");
        try { importLayer.remove(); } catch (e2) { /* ignore */ }
        try { app.redraw(); } catch (e3) { /* ignore */ }
        var importedNames = [];
        var k;
        for (k = 0; k < sourceStyleNames.length; k++) {
            var styleExists = false;
            try { destinationDoc.graphicStyles.getByName(sourceStyleNames[k]); styleExists = true; } catch (e4) { styleExists = false; }
            if (styleExists) { importedNames.push(sourceStyleNames[k]); }
        }
        return "OK\t" + importedNames.join("\n");
    }
    WORKER_FUNCS.push(workerImportStyles);

    /* worker: 直前の操作を取り消す（★ドキュメント確認を app.undo() より先に）
       Undo the last operation (doc check MUST come before app.undo()) */
    function workerUndoLast() {
        if (app.documents.length === 0) { return "NODOC"; }
        try { app.undo(); app.redraw(); } catch (e) { return "ERR:" + e; }
        return "OK";
    }
    WORKER_FUNCS.push(workerUndoLast);

    /* worker: 選択オブジェクトに指定グラフィックスタイルを適用
       undoFirst=true なら直前のプレビューを app.undo() で取り消してから再適用（★ドキュメント確認を先に）
       戻り値: "OK" / "NODOC" / "NOSEL" / "NOSTYLE"
       Apply the named graphic style to the current selection. When undoFirst is true, undo the previous
       preview first, then re-apply (doc check comes before app.undo()). Returns OK/NODOC/NOSEL/NOSTYLE */
    function workerApplyStyleToSelection(styleName, undoFirst) {
        if (app.documents.length === 0) { return "NODOC"; }
        if (undoFirst) {
            try { app.undo(); app.redraw(); } catch (e) { /* ignore */ }
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) { return "NOSEL"; }
        var graphicStyle = null;
        try { graphicStyle = doc.graphicStyles.getByName(styleName); } catch (e2) { graphicStyle = null; }
        if (!graphicStyle) { return "NOSTYLE"; }
        var i;
        for (i = 0; i < currentSelection.length; i++) {
            try { graphicStyle.applyTo(currentSelection[i]); } catch (e3) { /* ignore */ }
        }
        try { app.redraw(); } catch (e4) { /* ignore */ }
        return "OK";
    }
    WORKER_FUNCS.push(workerApplyStyleToSelection);

    // ---- エリア内文字変換 worker 群 / Area-type conversion workers ----

    /* worker: /name ブロック（ASCII, 大文字hex）/ /name block (ASCII, uppercase hex) */
    function workerNameBlockAscii(text) {
        return "/name [ " + text.length + " " + workerHexAscii(text).toUpperCase() + " ]";
    }
    WORKER_FUNCS.push(workerNameBlockAscii);

    /* worker: アクションセット定義(.aia)文字列を組み立て / Build an action set (.aia) string */
    function workerBuildActionSetAIA(setName, internalName, localizedNameHex, paramKeyInt, actionDefs) {
        var aiaText = "/version 3" + workerNameBlockAscii(setName) + "/isOpen 1" + "/actionCount " + actionDefs.length;
        var i;
        for (i = 0; i < actionDefs.length; i++) {
            var actionDef = actionDefs[i];
            aiaText += "/action-" + (i + 1) + " {" +
                " " + workerNameBlockAscii(actionDef.name) +
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
        return aiaText;
    }
    WORKER_FUNCS.push(workerBuildActionSetAIA);

    /* worker: アクションセットを読み込む（既存は先に外し、temp に .aia を書き出して loadAction）/ Load an action set */
    function workerLoadActionSet(setName, aiaString) {
        workerUnloadActionSet(setName);
        try {
            var actionFile = new File(Folder.temp + "/IAAGS_action_" + setName + ".aia");
            actionFile.open("w");
            actionFile.write(aiaString);
            actionFile.close();
            app.loadAction(actionFile);
            actionFile.remove();
        } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerLoadActionSet);

    /* worker: アクションセットを破棄 / Unload an action set */
    function workerUnloadActionSet(setName) {
        try { app.unloadAction(setName, ""); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerUnloadActionSet);

    /* worker: フレーム整列アクション（AlignTop/Center/Bottom/Justify）を読み込む / Load frame-alignment actions */
    function workerLoadAreaTextActions() {
        var aiaText = workerBuildActionSetAIA(
            "AreaText",
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
        workerLoadActionSet("AreaText", aiaText);
    }
    WORKER_FUNCS.push(workerLoadAreaTextActions);

    /* worker: フレーム整列アクションを破棄 / Unload frame-alignment actions */
    function workerUnloadAreaTextActions() {
        workerUnloadActionSet("AreaText");
    }
    WORKER_FUNCS.push(workerUnloadAreaTextActions);

    /* worker: 縦方向の配置アクションを実行（0=上,1=中央,2=下,3=均等）/ Run vertical-placement action */
    function workerRunFrameAlignment(valueInt) {
        var actionNames = ["AlignTop", "AlignCenter", "AlignBottom", "AlignJustify"];
        if (valueInt !== 0 && valueInt !== 1 && valueInt !== 2 && valueInt !== 3) { return; }
        try { app.doScript(actionNames[valueInt], "AreaText", false); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerRunFrameAlignment);

    /* worker: 指定フレームに縦方向の配置を適用 / Apply vertical placement to a frame */
    function workerApplyFrameAlignment(areaType, valueInt) {
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaType];
            app.redraw();
            workerRunFrameAlignment(valueInt);
        } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerApplyFrameAlignment);

    /* worker: TextFrame の見た目を temp_style として登録し、登録名を返す / Register the appearance as temp_style */
    function workerRegisterTempStyle(textFrame) {
        if (!textFrame) { return null; }
        var doc = app.activeDocument;
        if (!doc) { return null; }
        var graphicStyles = doc.graphicStyles;
        if (!graphicStyles) { return null; }
        try { graphicStyles.getByName("temp_style").remove(); } catch (e) { /* ignore */ }
        doc.selection = null;
        try { textFrame.selected = true; } catch (selErr) { return null; }
        var countBefore = graphicStyles.length;
        var aiaText = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /value 1 } } }';
        workerLoadActionSet("GraphicStyle", aiaText);
        try { app.doScript("AddNewWithoutName", "GraphicStyle", false); } catch (runErr) { /* ignore */ }
        workerUnloadActionSet("GraphicStyle");
        if (graphicStyles.length <= countBefore) { return null; }
        graphicStyles[graphicStyles.length - 1].name = "temp_style";
        return "temp_style";
    }
    WORKER_FUNCS.push(workerRegisterTempStyle);

    /* worker: 名前でグラフィックスタイルを適用 / Apply a graphic style by name */
    function workerApplyStyleByName(styleName, targetItem) {
        if (!styleName || !targetItem) { return false; }
        try { app.activeDocument.graphicStyles.getByName(styleName).applyTo(targetItem); return true; } catch (e) { return false; }
    }
    WORKER_FUNCS.push(workerApplyStyleByName);

    /* worker: 名前でグラフィックスタイルを削除 / Remove a graphic style by name */
    function workerRemoveStyleByName(styleName) {
        if (!styleName) { return; }
        try { app.activeDocument.graphicStyles.getByName(styleName).remove(); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerRemoveStyleByName);

    /* worker: 選択の可視バウンディングボックスの和 / Union of visibleBounds over a selection */
    function workerSelectionVisibleBounds(currentSelection) {
        if (!currentSelection || !currentSelection.length) { return null; }
        var left = null, top = null, right = null, bottom = null;
        var i;
        for (i = 0; i < currentSelection.length; i++) {
            var itemBounds;
            try { itemBounds = currentSelection[i].visibleBounds; } catch (e) { continue; }
            if (!itemBounds) { continue; }
            if (left === null || itemBounds[0] < left) { left = itemBounds[0]; }
            if (top === null || itemBounds[1] > top) { top = itemBounds[1]; }
            if (right === null || itemBounds[2] > right) { right = itemBounds[2]; }
            if (bottom === null || itemBounds[3] < bottom) { bottom = itemBounds[3]; }
        }
        if (left === null) { return null; }
        return [left, top, right, bottom];
    }
    WORKER_FUNCS.push(workerSelectionVisibleBounds);

    /* worker: 複製→アピアランス分割→アウトラインで正確な可視サイズを計測 / Measure accurate visible size */
    function workerMeasureAccurateBounds(sourceItem) {
        var doc = app.activeDocument;
        var savedSelection = doc.selection;
        var measured = null;
        var duplicatedItem = null;
        try {
            duplicatedItem = sourceItem.duplicate();
            doc.selection = null;
            duplicatedItem.selected = true;
            app.redraw();
            try { app.executeMenuCommand("expandStyle"); } catch (e) { /* ignore */ }
            try { app.executeMenuCommand("outline"); } catch (e5) { /* ignore */ }
            app.redraw();
            var resultSelection = doc.selection;
            var unionBounds = workerSelectionVisibleBounds(resultSelection);
            if (unionBounds) {
                measured = { left: unionBounds[0], top: unionBounds[1], right: unionBounds[2], bottom: unionBounds[3], width: unionBounds[2] - unionBounds[0], height: unionBounds[1] - unionBounds[3] };
            }
            var k;
            for (k = resultSelection.length - 1; k >= 0; k--) {
                try { resultSelection[k].remove(); } catch (e2) { /* ignore */ }
            }
        } catch (e0) {
            if (duplicatedItem) { try { duplicatedItem.remove(); } catch (e3) { /* ignore */ } }
        }
        try { doc.selection = savedSelection; } catch (e4) { /* ignore */ }
        return measured;
    }
    WORKER_FUNCS.push(workerMeasureAccurateBounds);

    /* worker: geometricBounds を {left, top, width, height} で返す / geometricBounds as an object */
    function workerGeometricBounds(sourceItem) {
        var itemBounds = sourceItem.geometricBounds;
        return { left: itemBounds[0], top: itemBounds[1], width: itemBounds[2] - itemBounds[0], height: itemBounds[1] - itemBounds[3] };
    }
    WORKER_FUNCS.push(workerGeometricBounds);

    /* worker: 軸並行・直線コーナーの長方形か判定 / Is this an axis-aligned straight-corner rectangle */
    function workerIsRectanglePath(pageItem) {
        if (!pageItem || pageItem.typename !== "PathItem" || !pageItem.closed) { return false; }
        var points;
        try { points = pageItem.pathPoints; } catch (e) { return false; }
        if (!points || points.length !== 4) { return false; }
        var i;
        for (i = 0; i < points.length; i++) {
            var anchor = points[i].anchor, left = points[i].leftDirection, right = points[i].rightDirection;
            if (Math.abs(anchor[0] - left[0]) >= 0.01 || Math.abs(anchor[1] - left[1]) >= 0.01) { return false; }
            if (Math.abs(anchor[0] - right[0]) >= 0.01 || Math.abs(anchor[1] - right[1]) >= 0.01) { return false; }
        }
        var distinctX = [], distinctY = [];
        var j, k;
        for (j = 0; j < points.length; j++) {
            var ax = points[j].anchor[0], ay = points[j].anchor[1];
            var foundX = false, foundY = false;
            for (k = 0; k < distinctX.length; k++) { if (Math.abs(distinctX[k] - ax) < 0.01) { foundX = true; } }
            if (!foundX) { distinctX.push(ax); }
            for (k = 0; k < distinctY.length; k++) { if (Math.abs(distinctY[k] - ay) < 0.01) { foundY = true; } }
            if (!foundY) { distinctY.push(ay); }
        }
        return distinctX.length === 2 && distinctY.length === 2;
    }
    WORKER_FUNCS.push(workerIsRectanglePath);

    /* worker: 元テキストの自動カーニングと文字組みを取得 / Capture auto-kerning and mojikumi */
    function workerReadKerningMojikumi(sourceText) {
        var textSnapshot = { kerningMethod: null, mojikumi: null };
        try { textSnapshot.kerningMethod = sourceText.textRange.characterAttributes.kerningMethod; } catch (e) { /* ignore */ }
        try {
            if (sourceText.paragraphs.length > 0) { textSnapshot.mojikumi = sourceText.paragraphs[0].paragraphAttributes.mojikumi; }
        } catch (e2) { /* ignore */ }
        return textSnapshot;
    }
    WORKER_FUNCS.push(workerReadKerningMojikumi);

    /* worker: 自動カーニングと文字組みをエリア内文字へ適用 / Apply auto-kerning and mojikumi */
    function workerApplyKerningMojikumi(areaType, textSnapshot) {
        if (textSnapshot.kerningMethod !== null) {
            try { areaType.textRange.characterAttributes.kerningMethod = textSnapshot.kerningMethod; } catch (e) { /* ignore */ }
        }
        if (textSnapshot.mojikumi !== null && textSnapshot.mojikumi !== undefined) {
            var areaParagraphs = areaType.paragraphs;
            var i;
            for (i = 0; i < areaParagraphs.length; i++) {
                try { areaParagraphs[i].paragraphAttributes.mojikumi = textSnapshot.mojikumi; } catch (e2) { /* ignore */ }
            }
        }
    }
    WORKER_FUNCS.push(workerApplyKerningMojikumi);

    /* worker: 図形パスをエリア内文字にして内容・書式・スタイルを移す / Turn a path into area type carrying text */
    function workerFillAreaType(doc, framePath, sourceText, graphicStyleName) {
        var sourceFont = null, sourceSize = 0;
        try {
            var sourceAttributes = sourceText.textRange.characterAttributes;
            sourceFont = sourceAttributes.textFont;
            sourceSize = sourceAttributes.size;
        } catch (e) { /* ignore */ }
        var textSnapshot = workerReadKerningMojikumi(sourceText);
        var areaType = doc.textFrames.areaText(framePath);
        areaType.contents = sourceText.contents;
        try {
            if (sourceFont) { areaType.textRange.characterAttributes.textFont = sourceFont; }
            if (sourceSize > 0) { areaType.textRange.characterAttributes.size = sourceSize; }
        } catch (e2) { /* ignore */ }
        workerApplyKerningMojikumi(areaType, textSnapshot);
        if (graphicStyleName) { workerApplyStyleByName(graphicStyleName, areaType); }
        return areaType;
    }
    WORKER_FUNCS.push(workerFillAreaType);

    /* worker: エリア内文字を水平・垂直とも中央に / Center area-type contents */
    function workerCenterAreaType(areaType) {
        try { areaType.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { /* ignore */ }
        workerApplyFrameAlignment(areaType, 1);
    }
    WORKER_FUNCS.push(workerCenterAreaType);

    /* worker: 塗り2枚＋長方形シェイプ効果でボタン状の背景に / Add two fills + rectangle shape effect */
    function workerApplyButtonShape(areaType) {
        try {
            app.activeDocument.selection = null;
            areaType.selected = true;
            app.redraw();
            app.executeMenuCommand("Adobe New Fill Shortcut");
            app.executeMenuCommand("Adobe New Fill Shortcut");
            areaType.applyEffect('<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 R RelWidth 0 R RelHeight 0 R AbsWidth 0 R AbsHeight 0 R Absolute 0 R CornerRadius 9 "/></LiveEffect>');
        } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerApplyButtonShape);

    /* worker: パス上文字を字形を保ったままポイント文字へ分離 / Detach path text into point text */
    function workerDetachPathText(doc, pathTextFrames) {
        function ignoreErrors(attemptAction) { try { return attemptAction(); } catch (e) { return undefined; } }
        var createdTexts = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) { return createdTexts; }
        ignoreErrors(function () { doc.selection = null; });
        var j;
        for (j = pathTextFrames.length - 1; j >= 0; j--) {
            var pathText = pathTextFrames[j];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) { continue; }
            var originalPath = null;
            ignoreErrors(function () { originalPath = pathText.textPath; });
            if (!originalPath) { continue; }
            var attributeSnapshots = [];
            var i;
            for (i = 0; i < pathText.characters.length; i++) {
                var sourceAttributes = pathText.characters[i].characterAttributes;
                attributeSnapshots.push({ font: sourceAttributes.textFont, size: sourceAttributes.size, fillColor: sourceAttributes.fillColor, strokeColor: sourceAttributes.strokeColor, strokeWeight: sourceAttributes.strokeWeight, autoLeading: sourceAttributes.autoLeading, leading: sourceAttributes.leading });
            }
            var textContents = "";
            ignoreErrors(function () { textContents = pathText.contents; });
            var justification = null;
            ignoreErrors(function () {
                if (pathText.paragraphs && pathText.paragraphs.length > 0) { justification = pathText.paragraphs[0].paragraphAttributes.justification; }
            });
            var pointText = doc.textFrames.add();
            var anchorPoint = null;
            ignoreErrors(function () {
                if (originalPath.pathPoints && originalPath.pathPoints.length > 0) { anchorPoint = originalPath.pathPoints[0].anchor; }
            });
            if (anchorPoint) { pointText.position = [anchorPoint[0], anchorPoint[1]]; }
            pointText.contents = textContents;
            if (justification !== null && pointText.paragraphs && pointText.paragraphs.length > 0) {
                ignoreErrors(function () { pointText.paragraphs[0].paragraphAttributes.justification = justification; });
            }
            ignoreErrors(function () {
                var noColor = new NoColor();
                pointText.textRange.characterAttributes.strokeColor = noColor;
                pointText.textRange.characterAttributes.strokeWeight = 0;
            });
            var restoreCount = Math.min(pointText.characters.length, attributeSnapshots.length);
            var k;
            for (k = 0; k < restoreCount; k++) {
                var targetAttr = pointText.characters[k].characterAttributes;
                var savedAttr = attributeSnapshots[k];
                ignoreErrors(function () { targetAttr.textFont = savedAttr.font; });
                ignoreErrors(function () { targetAttr.size = savedAttr.size; });
                ignoreErrors(function () { targetAttr.fillColor = savedAttr.fillColor; });
                ignoreErrors(function () {
                    var savedStroke = savedAttr.strokeColor;
                    targetAttr.strokeColor = savedStroke;
                    targetAttr.strokeWeight = (savedStroke && savedStroke.typename === "NoColor") ? 0 : savedAttr.strokeWeight;
                });
                ignoreErrors(function () { targetAttr.baselineShift = 0; });
                ignoreErrors(function () { targetAttr.horizontalScale = 100; });
                ignoreErrors(function () { targetAttr.verticalScale = 100; });
                ignoreErrors(function () { targetAttr.autoLeading = savedAttr.autoLeading; });
                if (!savedAttr.autoLeading) { ignoreErrors(function () { targetAttr.leading = savedAttr.leading; }); }
            }
            ignoreErrors(function () { pathText.remove(); });
            ignoreErrors(function () { pointText.selected = true; });
            createdTexts.push(pointText);
        }
        return createdTexts;
    }
    WORKER_FUNCS.push(workerDetachPathText);

    /* worker: 選択内のパス上文字をポイント文字へ置き換えた選択配列を返す / Replace path text with point text */
    function workerPreprocessPathText(doc, currentSelection) {
        if (!doc || !currentSelection || !currentSelection.length) { return currentSelection; }
        var pathTexts = [];
        var i;
        for (i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            try {
                if (selectedItem && selectedItem.typename === "TextFrame" && selectedItem.kind === TextType.PATHTEXT) { pathTexts.push(selectedItem); }
            } catch (e0) { /* ignore */ }
        }
        if (!pathTexts.length) { return currentSelection; }
        var pointTexts = workerDetachPathText(doc, pathTexts);
        if (!pointTexts.length) { return currentSelection; }
        var replaced = [];
        var j;
        for (j = 0; j < currentSelection.length; j++) {
            var keepItem = currentSelection[j];
            try {
                if (keepItem && keepItem.typename === "TextFrame" && keepItem.kind === TextType.PATHTEXT) { /* skip old */ }
                else if (keepItem) { replaced.push(keepItem); }
            } catch (e1) { /* ignore */ }
        }
        var k;
        for (k = 0; k < pointTexts.length; k++) { replaced.push(pointTexts[k]); }
        try { doc.selection = replaced; } catch (e) { /* ignore */ }
        try { app.redraw(); } catch (e2) { /* ignore */ }
        return replaced;
    }
    WORKER_FUNCS.push(workerPreprocessPathText);

    /* worker: 適用スタイルを決定（外部スタイル or 元テキストの一時スタイル）/ Resolve the style to apply */
    function workerResolveStyle(externalStyleName, sourceText) {
        if (externalStyleName) { return { name: externalStyleName, isTemp: false }; }
        return { name: workerRegisterTempStyle(sourceText), isTemp: true };
    }
    WORKER_FUNCS.push(workerResolveStyle);

    /* worker: テキスト＋図形 → 図形を複製してエリア内文字に / Text + shape -> duplicate the shape into area type */
    function workerConvertTextIntoShape(doc, selectedItems, externalStyleName) {
        var createdFrames = [];
        var sourceText = null, sourceShape = null;
        var i;
        for (i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!sourceText && selectedItem.typename === "TextFrame") { sourceText = selectedItem; }
            else if (!sourceShape && workerIsRectanglePath(selectedItem)) { sourceShape = selectedItem; }
        }
        if (!sourceText || !sourceShape) { return createdFrames; }
        var styleInfo = workerResolveStyle(externalStyleName, sourceText);
        try {
            var framePath = sourceShape.duplicate();
            framePath.filled = false;
            framePath.stroked = false;
            var areaType = workerFillAreaType(doc, framePath, sourceText, styleInfo.name);
            workerCenterAreaType(areaType);
            if (styleInfo.isTemp) { workerApplyButtonShape(areaType); }
            sourceText.remove();
            sourceShape.remove();
            createdFrames.push(areaType);
        } catch (e) {
            /* ignore */
        } finally {
            if (styleInfo.isTemp) { workerRemoveStyleByName(styleInfo.name); }
        }
        return createdFrames;
    }
    WORKER_FUNCS.push(workerConvertTextIntoShape);

    /* worker: ポイント文字のみ → 計測した実寸でフレームを作り中央配置 / Point text -> frame at measured size, centered */
    function workerConvertPointText(doc, selectedItems, adjust, widthRatio, heightRatio, externalStyleName) {
        var createdFrames = [];
        var appliedWidthRatio = adjust ? widthRatio : 1;
        var appliedHeightRatio = adjust ? heightRatio : 1;
        var i;
        for (i = selectedItems.length - 1; i >= 0; i--) {
            var sourceText = selectedItems[i];
            if (!(sourceText.typename === "TextFrame" && sourceText.kind === TextType.POINTTEXT)) { continue; }
            var styleInfo = workerResolveStyle(externalStyleName, sourceText);
            try {
                var measuredBounds = workerMeasureAccurateBounds(sourceText);
                if (!measuredBounds) { measuredBounds = workerGeometricBounds(sourceText); }
                var frameWidth = measuredBounds.width * appliedWidthRatio;
                var frameHeight = measuredBounds.height * appliedHeightRatio;
                var centerX = measuredBounds.left + measuredBounds.width / 2;
                var centerY = measuredBounds.top - measuredBounds.height / 2;
                var frameLeft = centerX - frameWidth / 2;
                var frameTop = centerY + frameHeight / 2;
                var framePath = doc.pathItems.rectangle(frameTop, frameLeft, frameWidth, frameHeight);
                framePath.filled = false;
                framePath.stroked = false;
                var areaType = workerFillAreaType(doc, framePath, sourceText, styleInfo.name);
                workerCenterAreaType(areaType);
                if (adjust) {
                    if (styleInfo.isTemp) { workerApplyButtonShape(areaType); }
                }
                createdFrames.push(areaType);
                sourceText.remove();
            } catch (e) {
                /* ignore */
            } finally {
                if (styleInfo.isTemp) { workerRemoveStyleByName(styleInfo.name); }
            }
        }
        return createdFrames;
    }
    WORKER_FUNCS.push(workerConvertPointText);

    /* worker: 選択をエリア内文字へ変換（エントリ）
       戻り値: "OK" / "NODOC" / "NOSEL" / "NOTEXT" / "RECTONLY" / "ERR:..."
       Convert the selection to area type (entry). Returns markers. */
    function workerConvertSelection(adjust, widthRatio, heightRatio, externalStyleName, undoFirst) {
        if (app.documents.length === 0) { return "NODOC"; }
        if (undoFirst) {
            try { app.undo(); app.redraw(); } catch (eU) { /* ignore */ }
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) { return "NOSEL"; }
        workerLoadAreaTextActions();
        var resultMarker = "OK";
        try {
            workerPreprocessPathText(doc, currentSelection);
            var updatedSelection = doc.selection;
            if (!updatedSelection || updatedSelection.length === 0) { return "NOSEL"; }
            var hasSourceText = false, hasFrameShape = false, hasNonRectShape = false;
            var i;
            for (i = 0; i < updatedSelection.length; i++) {
                var selectedItem = updatedSelection[i];
                if (selectedItem.typename === "TextFrame" && (selectedItem.kind === TextType.POINTTEXT || selectedItem.kind === TextType.PATHTEXT)) { hasSourceText = true; }
                if (workerIsRectanglePath(selectedItem)) { hasFrameShape = true; }
                else if (selectedItem.typename === "PathItem" && selectedItem.closed) { hasNonRectShape = true; }
            }
            if (!hasSourceText) { return "NOTEXT"; }
            if (hasNonRectShape && !hasFrameShape) { return "RECTONLY"; }
            var createdFrames = hasFrameShape
                ? workerConvertTextIntoShape(doc, updatedSelection, externalStyleName)
                : workerConvertPointText(doc, updatedSelection, adjust, widthRatio, heightRatio, externalStyleName);
            if (createdFrames.length > 0) {
                try { doc.selection = createdFrames; app.redraw(); } catch (e) { /* ignore */ }
            }
        } catch (eMain) {
            resultMarker = "ERR:" + eMain;
        } finally {
            workerUnloadAreaTextActions();
        }
        return resultMarker;
    }
    WORKER_FUNCS.push(workerConvertSelection);

    // =========================================
    // 設定ファイルとファイル選択 / Prefs & file picking
    // =========================================

    /**
     * パスから表示用ファイル名を取得する
     * @param {string} filePath - ファイルのパス
     * @returns {string} 表示用のファイル名（デコードできなければパスのまま）
     */
    function getDisplayFileName(filePath) {
        /* 不正なエスケープは decodeURI が例外を投げる / decodeURI throws on malformed escapes */
        try { return decodeURI(new File(filePath).name); } catch (e) { return filePath; }
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
                var prefsKey = prefsLines[i].substring(0, separatorIndex);
                var prefsValue = prefsLines[i].substring(separatorIndex + 1);
                if (prefsKey === "styleFilePath") savedState.filePath = prefsValue;
                else if (prefsKey === "styleNames") savedState.styleNames = prefsValue ? prefsValue.split("\t") : [];
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
     * スタイル用 AI ファイルを選ばせる
     * @returns {string} 選んだファイルのパス（キャンセルで空文字）
     */
    function pickStyleFile() {
        var pickedFile = File.openDialog(getLabel("dialog.pickFile"), function (candidate) {
            return (candidate instanceof Folder) || /\.ai$/i.test(candidate.name);
        });
        return pickedFile ? pickedFile.fsName : "";
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * ［エリア内文字オプション］パネルを組み立てる（マスターのチェックでほかの項目を有効化）
     * @param {Window} paletteWindow - 追加先のパレット
     * @returns {object} パネル内のコントロール（convertCheckbox / adjustOnRadio / widthInput / heightInput / convertButton）
     */
    function addAreaTypeOptionPanel(paletteWindow) {
        var areaTypeOptionPanel = paletteWindow.add("panel", undefined, getLabel("panel.areaTypeOption"));
        setupPanel(areaTypeOptionPanel, INNER_PANEL_SPACING);

        /* マスターチェック：オンのときだけ他パーツを有効化 / Master checkbox: enables the other parts only when checked */
        var convertCheckbox = areaTypeOptionPanel.add("checkbox", undefined, getLabel("checkbox.convertToAreaType"));
        convertCheckbox.helpTip = getLabel("tooltip.convertToAreaType");
        convertCheckbox.value = false; /* 既定はオフ / Default: off */

        /* サイズ：調整する / 調整しない / Size: adjust / don't adjust */
        var adjustModeGroup = areaTypeOptionPanel.add("group");
        var sizeLabel = adjustModeGroup.add("statictext", undefined, labelText("fieldLabel.size"));
        sizeLabel.preferredSize.width = SIZE_LABEL_WIDTH;
        var adjustOnRadio = adjustModeGroup.add("radiobutton", undefined, getLabel("radio.doAdjust"));
        adjustOnRadio.helpTip = getLabel("tooltip.doAdjust");
        var adjustOffRadio = adjustModeGroup.add("radiobutton", undefined, getLabel("radio.dontAdjust"));
        adjustOffRadio.helpTip = getLabel("tooltip.dontAdjust");
        adjustOffRadio.value = true; /* 既定は「しない」/ Default: off */

        /* 幅・高さの倍率（1行に横並び、百分率 % で入力）/ Width and height ratios (one row, entered as %) */
        var ratioRow = areaTypeOptionPanel.add("group");
        var widthInput = addRatioField(ratioRow, "widthRatio", BUTTON_WIDTH_RATIO);
        var heightInput = addRatioField(ratioRow, "heightRatio", BUTTON_HEIGHT_RATIO);

        /* 変換ボタン（このパネル内・主アクション）。パネル幅いっぱいに広げず右寄せ・自然幅に
           onClick は各パネル構築後に割り当て
           Convert button (primary action, inside this panel); right-aligned at its natural width (not full width)
           onClick is assigned after all panels are built */
        var convertButton = areaTypeOptionPanel.add("button", undefined, getLabel("button.runConvert"), { name: "ok" });
        convertButton.helpTip = getLabel("tooltip.runConvert");
        convertButton.alignment = ["right", "center"]; /* 親の fill を上書きして広げない / Override the panel's fill so it doesn't stretch */

        /**
         * マスターのチェックと「する／しない」に合わせて各項目の有効・無効を切り替える
         * @returns {void}
         */
        function updateAreaTypeOptionEnabled() {
            var convertOn = convertCheckbox.value;
            /* マスターがオフなら する/しない ラジオごとディム / Dim the on/off radios when the master is off */
            adjustOnRadio.enabled = convertOn;
            adjustOffRadio.enabled = convertOn;
            /* 幅・高さはマスターオン かつ「する」のときのみ / Width/height only when the master is on and "On" is chosen */
            var ratioEnabled = convertOn && adjustOnRadio.value;
            widthInput.enabled = ratioEnabled;
            heightInput.enabled = ratioEnabled;
            /* ∧∨も入力欄と一緒に無効にし、自作描画をディム表示に合わせて描き直す / dim the steppers with the fields and redraw them */
            widthInput.stepperGroup.enabled = ratioEnabled;
            heightInput.stepperGroup.enabled = ratioEnabled;
            redrawSteppersIn(ratioRow);
        }
        convertCheckbox.onClick = updateAreaTypeOptionEnabled;
        adjustOnRadio.onClick = updateAreaTypeOptionEnabled;
        adjustOffRadio.onClick = updateAreaTypeOptionEnabled;
        updateAreaTypeOptionEnabled();

        return {
            convertCheckbox: convertCheckbox,
            adjustOnRadio: adjustOnRadio,
            widthInput: widthInput,
            heightInput: heightInput,
            convertButton: convertButton
        };
    }

    /**
     * 倍率の入力欄（ラベル＋∧∨＋入力欄＋%）を追加する
     * @param {Group} ratioRow - 追加先の行
     * @param {string} labelKey - LABELS.fieldLabel と LABELS.tooltip のキー
     * @param {number} defaultRatio - 初期値の倍率（1 = 100%）
     * @returns {EditText} 追加した入力欄
     */
    function addRatioField(ratioRow, labelKey, defaultRatio) {
        ratioRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = ratioRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var ratioInput;
        /* 1%以上（0以下は readRatio() で既定の倍率に戻るため） / at least 1% (0 or less falls back to the default ratio) */
        var stepperGroup = addStepper(stepperInputGroup, function () { return ratioInput; }, { min: 1 });
        ratioInput = stepperInputGroup.add("edittext", undefined, String(Math.round(defaultRatio * 100)));
        bindSteppedArrowKeys(ratioInput, stepperGroup);
        ratioInput.stepperGroup = stepperGroup; /* 有効／無効の切り替えで参照 / used when toggling enabled */
        ratioInput.characters = 4;
        ratioInput.preferredSize.width = RATIO_INPUT_WIDTH;
        ratioInput.helpTip = getLabel("tooltip." + labelKey);
        ratioRow.add("statictext", undefined, "%");
        return ratioInput;
    }

    /**
     * ［グラフィックスタイル］パネルを組み立てる
     * @param {Window} paletteWindow - 追加先のパレット
     * @returns {object} パネル内のコントロール（usedOnlyCheckbox / styleListbox / clearButton / openStylePanelButton）
     */
    function addStylePanel(paletteWindow) {
        var stylePanel = paletteWindow.add("panel", undefined, getLabel("panel.graphicStyle"));
        setupPanel(stylePanel, INNER_PANEL_SPACING);

        /* 使用中のみ表示するフィルタ / Filter to show only styles used in the document */
        var usedOnlyCheckbox = stylePanel.add("checkbox", undefined, getLabel("checkbox.usedOnly"));
        usedOnlyCheckbox.value = false;

        /* スタイル選択リスト（先頭が「元の見た目」、以降が現在のドキュメントのグラフィックスタイル）
           Style list (first row = original appearance, then the document's graphic styles) */
        var styleListbox = stylePanel.add("listbox", undefined, [], { multiselect: false });
        styleListbox.preferredSize.height = STYLE_LIST_HEIGHT;
        styleListbox.helpTip = getLabel("tooltip.styleList");

        /* 最下部：左右分割（左＝クリア／スペーサー／右＝パネル）/ Bottom row: split layout (left = Clear, spacer, right = Panel) */
        var styleFooterRow = stylePanel.add("group");
        styleFooterRow.orientation = "row";
        styleFooterRow.alignment = ["fill", "bottom"];
        styleFooterRow.margins = STYLE_FOOTER_MARGINS; /* ボタンエリア上部にマージン / Top margin above the button area */

        /* 左側グループ：未使用のグラフィックスタイルを削除 / Left group: delete unused graphic styles */
        var btnLeftGroup = styleFooterRow.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var clearButton = btnLeftGroup.add("button", undefined, getLabel("button.clearStyle"));
        clearButton.helpTip = getLabel("tooltip.clearStyle");

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = styleFooterRow.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ：グラフィックスタイルパネルへ移動 / Right group: jump to the Graphic Styles panel */
        var btnRightGroup = styleFooterRow.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var openStylePanelButton = btnRightGroup.add("button", undefined, getLabel("button.openStylePanel"));
        openStylePanelButton.helpTip = getLabel("tooltip.openStylePanel");

        return {
            usedOnlyCheckbox: usedOnlyCheckbox,
            styleListbox: styleListbox,
            clearButton: clearButton,
            openStylePanelButton: openStylePanelButton
        };
    }

    /**
     * ［スタイルの読み込み］パネルを組み立てる（ファイル名を上、ボタンを下に表示）
     * @param {Window} paletteWindow - 追加先のパレット
     * @returns {object} パネル内のコントロール（fileNameText / loadButton / reloadButton）
     */
    function addLoadPanel(paletteWindow) {
        var loadPanel = paletteWindow.add("panel", undefined, getLabel("panel.loadStyles"));
        setupPanel(loadPanel, INNER_PANEL_SPACING);
        var fileNameText = loadPanel.add("statictext", undefined, "", { truncate: "middle" });
        fileNameText.preferredSize.width = FILE_NAME_WIDTH;
        /* 読み込み / 再読み込みボタンを左寄せで横並び / Load & Reload buttons in a left-aligned row */
        var loadButtonRow = loadPanel.add("group");
        loadButtonRow.alignment = "left";
        loadButtonRow.margins = LOAD_BUTTON_ROW_MARGINS;
        var loadButton = loadButtonRow.add("button", undefined, getLabel("button.load"));
        loadButton.helpTip = getLabel("tooltip.load"); /* 使い方はツールチップで案内 / Usage hint shown as a tooltip */
        var reloadButton = loadButtonRow.add("button", undefined, getLabel("button.reload"));
        reloadButton.helpTip = getLabel("tooltip.reload"); /* 記憶したファイルから再取り込み / Re-import from the remembered file */

        return { fileNameText: fileNameText, loadButton: loadButton, reloadButton: reloadButton };
    }

    /**
     * 百分率の入力を倍率に換算する（不正な値は既定の倍率）
     * @param {EditText} ratioInput - 倍率の入力欄（%）
     * @param {number} defaultRatio - 既定の倍率
     * @returns {number} 倍率
     */
    function readRatio(ratioInput, defaultRatio) {
        var percentValue = parseFloat(ratioInput.text);
        return (isNaN(percentValue) || percentValue <= 0) ? defaultRatio : percentValue / 100;
    }

    /**
     * 常駐パレットを構築して表示する。「変換」ボタンで選択をエリア内文字化（DOM 処理はメインエンジンへ委譲）
     * 「読み込み」ボタンで別の AI ファイルを選ぶと、その場でリストを組み直す
     * @param {{filePath: string, styleNames: string[]}} savedStyleState - 記憶していたスタイルファイルとスタイル名
     * @returns {void}
     */
    function showPalette(savedStyleState) {
        /* すでにパレットが開いていれば前面に出して終了 / If the palette already exists, bring it forward and return */
        /* 閉じたパレットの参照は読めないことがある / A closed palette's reference may be unreadable */
        try {
            if ($.global.__importAndApplyGraphicStylePalette && $.global.__importAndApplyGraphicStylePalette.visible) {
                $.global.__importAndApplyGraphicStylePalette.show();
                return;
            }
        } catch (eExisting) { /* ignore */ }

        var styleState = {
            filePath: (savedStyleState && savedStyleState.filePath) || "",
            styleNames: (savedStyleState && savedStyleState.styleNames) || []
        };

        var paletteWindow = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        $.global.__importAndApplyGraphicStylePalette = paletteWindow;
        paletteWindow.orientation = "column";
        paletteWindow.alignChildren = "fill";
        paletteWindow.margins = PALETTE_MARGINS;

        var optionControls = addAreaTypeOptionPanel(paletteWindow);
        var styleControls = addStylePanel(paletteWindow);
        var loadControls = addLoadPanel(paletteWindow);
        var styleListbox = styleControls.styleListbox;

        var styleListValues = []; /* 各行に対応する値（null=元の見た目、以降はスタイル名）/ Value per row (null = original, else style name) */

        /* リビルド中は onChange の即時適用を抑止 / Suppress live-apply during a rebuild */
        var isRebuildingStyleList = false;

        /**
         * 選択中のファイル名表示と再読み込みボタンの有効状態を更新する
         * @returns {void}
         */
        function refreshFileLabel() {
            loadControls.fileNameText.text = styleState.filePath ? getDisplayFileName(styleState.filePath) : getLabel("status.noFileSelected");
            loadControls.reloadButton.enabled = !!styleState.filePath; /* 記憶したファイルが無ければ再読み込み不可 / Disable Reload without a remembered file */
        }

        /**
         * 現在のグラフィックスタイルでリストを組み直す（一覧取得はメインエンジンへ委譲）
         * @returns {void}
         */
        function rebuildStyleList() {
            styleListbox.removeAll();
            styleListValues = [];
            /* 先頭に「元の見た目」/ First row = original appearance */
            styleListbox.add("item", getLabel("listItem.styleOriginal"));
            styleListValues.push(null);

            /* 「使用中のみ」ON のときだけ未使用を除外（判定は worker 内で実施）
               Skip unused only when the filter is on (classification happens inside the worker) */
            var parsedNames = parseMarkerList(delegate("workerGetStyleNames(" + (styleControls.usedOnlyCheckbox.value ? "true" : "false") + ")"));
            for (var j = 0; j < parsedNames.items.length; j++) {
                styleListbox.add("item", parsedNames.items[j]);
                styleListValues.push(parsedNames.items[j]);
            }
            isRebuildingStyleList = true;
            styleListbox.selection = 0; /* 既定は「元の見た目」/ Default: original appearance */
            isRebuildingStyleList = false;
        }

        /**
         * AI ファイルからスタイルを取り込み（メインエンジンへ委譲）、記憶してリストを組み直す
         * @param {string} filePath - スタイルの AI ファイルのパス
         * @returns {void}
         */
        function importAndRemember(filePath) {
            var parsedNames = parseMarkerList(delegate("workerImportStyles(" + jsStringLiteral(filePath) + ")"));
            if (!parsedNames.ok) {
                if (parsedNames.marker === "NOFILE") alert(getLabel("alert.fileNotFound") + getDisplayFileName(filePath));
                return;
            }
            styleState.filePath = filePath;
            styleState.styleNames = parsedNames.items;
            saveStyleState(filePath, parsedNames.items); /* 次回以降このファイルを参照 / Remember for next runs */
            refreshFileLabel();
            rebuildStyleList();
        }

        styleControls.clearButton.onClick = function () {
            /* 未使用削除はメインエンジンへ委譲（ダイナミックアクション）/ Delegate the unused-prune to the main engine (dynamic action) */
            delegate("workerPruneUnused()");
            rebuildStyleList();
        };
        styleControls.openStylePanelButton.onClick = function () {
            /* DOM/メニュー操作はメインエンジンへ委譲 / Delegate the DOM/menu op to the main engine */
            delegate("workerOpenStylePanel()");
        };
        styleControls.usedOnlyCheckbox.onClick = function () { rebuildStyleList(); };

        /* listbox で選んだグラフィックスタイルを、その時点の選択オブジェクトへ即適用（「適用」ボタン無し）
           「元の見た目」行・リビルド時は適用しない。選択なし等はクリックごとに警告せず無言で無視
           Live-apply the picked style to the current selection (no Apply button);
           skip the "Original" row and rebuilds, and ignore no-selection silently */
        styleListbox.onChange = function () {
            if (isRebuildingStyleList) return;
            if (!styleListbox.selection) return;
            var pickedStyleName = styleListValues[styleListbox.selection.index];
            if (!pickedStyleName) return; // 「元の見た目」は適用対象なし / "Original appearance" has nothing to apply
            delegate("workerApplyStyleToSelection(" + jsStringLiteral(pickedStyleName) + ",false)");
        };

        /* onClick で連結（addEventListener は発火しない環境があるため）/ Use onClick, not addEventListener */
        loadControls.loadButton.onClick = function () {
            var pickedPath = pickStyleFile();
            if (!pickedPath) return;
            importAndRemember(pickedPath);
        };

        /* 記憶したファイルを選び直さずに再取り込み（別ドキュメントでも同じファイルを再利用）
           Re-import from the remembered file without re-picking (reuse the same file in another document) */
        loadControls.reloadButton.onClick = function () {
            if (!styleState.filePath) return;
            importAndRemember(styleState.filePath);
        };

        refreshFileLabel();
        rebuildStyleList();

        /* 「変換」：現在の UI 設定で選択をエリア内文字化。DOM 処理はメインエンジンへ委譲し、パレットは開いたまま
           "Convert": convert the selection using the current UI settings. DOM work is delegated; the palette stays open. */
        optionControls.convertButton.onClick = function () {
            /* 幅・高さは百分率 % 入力を倍率へ換算 / Width/height: convert the % input to a ratio */
            var widthRatio = readRatio(optionControls.widthInput, BUTTON_WIDTH_RATIO);
            var heightRatio = readRatio(optionControls.heightInput, BUTTON_HEIGHT_RATIO);
            /* 「元の見た目」なら null、スタイル選択時はその名前 / null for original, else the selected style name */
            var externalStyleName = null;
            if (styleListbox.selection) { externalStyleName = styleListValues[styleListbox.selection.index]; }

            if (optionControls.convertCheckbox.value) {
                /* エリア内文字変換はメインエンジンへ委譲（スタイルは読み込み済みで現在のドキュメントに存在）
                   Delegate the area-type conversion to the main engine (the style is already imported in the doc) */
                var convertResult = delegate("workerConvertSelection(" +
                    (optionControls.adjustOnRadio.value ? "true" : "false") + "," +
                    widthRatio + "," +
                    heightRatio + "," +
                    jsStringLiteral(externalStyleName ? externalStyleName : "") + ",false)");
                if (convertResult === "NODOC") { alert(getLabel("alert.noDocument")); }
                else if (convertResult === "RECTONLY") { alert(getLabel("alert.rectangleOnly")); }
                else if (convertResult === "NOSEL" || convertResult === "NOTEXT") { alert(getLabel("alert.selectText")); }
            } else if (externalStyleName) {
                /* 変換OFF：選択したグラフィックスタイルを選択オブジェクトへ適用（メインエンジンへ委譲）
                   Convert off: apply the chosen graphic style to the selection (delegated to the main engine) */
                var applyResult = delegate("workerApplyStyleToSelection(" + jsStringLiteral(externalStyleName) + ",false)");
                if (applyResult === "NODOC") { alert(getLabel("alert.noDocument")); }
                else if (applyResult === "NOSEL") { alert(getLabel("alert.selectText")); }
            }
        };

        /* パレットをアクティブにして Esc キーで閉じる / Close the palette with Esc while it is active */
        paletteWindow.addEventListener("keydown", function (event) {
            if (event && event.keyName === "Escape") { paletteWindow.close(); }
        });

        /* パネル内のボタンはレイアウト確定後に高さを詰める / Trim panel buttons' height after layout */
        paletteWindow.onShow = function () {
            trimButtonHeight(loadControls.loadButton, BUTTON_HEIGHT_TRIM);
            trimButtonHeight(loadControls.reloadButton, BUTTON_HEIGHT_TRIM);
            trimButtonHeight(styleControls.clearButton, BUTTON_HEIGHT_TRIM);
            trimButtonHeight(styleControls.openStylePanelButton, BUTTON_HEIGHT_TRIM);
        };
        /* 閉じたらグローバル参照を解放 / Release the global reference when closed */
        paletteWindow.onClose = function () {
            $.global.__importAndApplyGraphicStylePalette = null;
        };

        paletteWindow.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 記憶した参照ファイルとスタイル名を読み込み、常駐パレットを表示する
     * 変換・スタイル適用は「変換」ボタンから実行する（DOM 処理はメインエンジンへ委譲）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        showPalette(loadSavedStyleState());
    }

    main();

})();
