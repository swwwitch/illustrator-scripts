#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "DialogEngine"

/*

### 概要

アートボードのサイズを「操作」×「対象」×「サイズ」の組み合わせで自動調整します。
選択オブジェクトに合わせるほか、各アートボード内のオブジェクトに合わせて全アートボードを個別に調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardWithMargin.md

note記事も参照してください。
https://note.com/dtp_transit/n/n15d3c6c5a1e5

### Overview

Adjusts artboard size by operation, target and size (width & height).
Fits to the selection, or fits every artboard individually to the objects it contains.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitArtboardWithMargin.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitArtboardWithMargin";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.10.1";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardWithMargin.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitArtboardWithMargin.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_transit/n/n15d3c6c5a1e5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値（保存済みの設定があればそちらが優先） / Dialog defaults (stored settings win) */
    var DIALOG_DEFAULTS = {
        marginByUnit: {
            mm: '5',
            px: '20',
            pt: '10',
            _fallback: '0'
        },
        previewBounds: true,        // プレビュー境界(visibleBounds)を既定に / use visibleBounds by default
        roundMode: 'pixelGrid',     // 既定の丸めモード（pixelGrid / currentUnit / none）/ default rounding mode
        link: true                  // 上下左右の連動を既定ON / link margins by default
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var BUTTON_ROW_TOP_MARGIN = 5;           /* ボタンエリアの上余白 / top margin of the button row */

    /* 初回表示時の画面中央からの横オフセット / First-run offset from screen center */
    var DIALOG_FIRST_RUN_OFFSET_X = 300;

    /* 行ラベルの幅（言語別） / Row label widths per language */
    var BASIS_LABEL_WIDTHS = { ja: 40, en: 76 };    /* 調整基準パネル / adjustment basis panel */
    var MARGIN_LABEL_WIDTHS = { ja: 32, en: 62 };   /* マージンパネル / margin panel */

    /* マージン入力欄の桁数 / Margin field width in characters */
    var MARGIN_FIELD_CHARACTERS = 4;

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

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

    /**
     * 横並びの行グループ（ボタン列など）の共通設定を適用する
     * @param {Group} rowGroup - 対象のグループ
     * @param {string} [alignment] - グループ自身の配置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
    // 2. アイコンを addLinkToggle(親, 初期値, 切り替え後の関数) で作る。helpTip はコピー先で付ける
    //      var linkToggle = addLinkToggle(fieldsRowGroup, true, function () { syncFields(); });
    //      linkToggle.helpTip = getLabel(LABELS.tooltip.linkToggle);
    // 3. 連動中かは linkToggle.value で読む。コードから変えるときは setLinkToggleValue(linkToggle, true)
    // 4. 有効／無効は setLinkToggleEnabled(linkToggle, isEnabled)（無効の間はクリックが効かず、薄く描く）
    // 5. 2つの入力欄の右に置くときは、行 group の中に「入力欄を縦に積んだ group」とアイコンを並べ、
    //    行 group の alignChildren を ["left", "center"] にすると上下の中央に来る
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkLinkToggleUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var LINK_UI_DARK = isDarkLinkToggleUI();
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle
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

    /**
     * 現在の定規単位を UnitValue に渡せる文字列で返す（UnitValue が扱えない単位は pt に寄せる）
     * @returns {string} 単位の文字列
     */
    function getRulerUnitString() {
        var unit = getUnitInfo();
        return (unit.code >= 0 && unit.code <= 6) ? unit.label : 'pt';
    }

    /**
     * 単位ごとの初期マージン値を返す
     * @param {string} unit - 単位の文字列
     * @returns {string} 初期マージン値
     */
    function getDefaultMargin(unit) {
        return DIALOG_DEFAULTS.marginByUnit.hasOwnProperty(unit) ?
            DIALOG_DEFAULTS.marginByUnit[unit] :
            DIALOG_DEFAULTS.marginByUnit._fallback;
    }

    /**
     * 数値＋単位を pt に変換する
     * @param {number|string} value - 値
     * @param {string} unit - 単位の文字列
     * @returns {number} pt。変換できなければ NaN
     */
    function toPt(value, unit) {
        var numericValue = Number(value);
        if (isNaN(numericValue)) return NaN;
        // 歯/Q は UnitValue が非対応のため手計算（1H = 1Q = 0.25mm、1mm = 72/25.4pt） / H and Q are unsupported by UnitValue
        if (unit === "H" || unit === "Q") {
            return numericValue * 0.25 * 72 / 25.4;
        }
        /* UnitValue が単位を受け付けないときに備える / guard against units UnitValue rejects */
        try {
            return new UnitValue(numericValue, unit).as('pt');
        } catch (e) {
            return NaN;
        }
    }

    // =========================================
    // ローカライズ / Localize
    // =========================================

    /**
     * 実行環境の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "アートボードサイズを調整", en: "Adjust Artboard Size" }
        },
        panel: {
            adjustmentBasis: { ja: "調整基準", en: "Adjustment basis" },
            margin: { ja: "マージン", en: "Margin" },
            fineTuning: { ja: "アートボードサイズの微調整", en: "Artboard size fine-tuning" }
        },
        fieldLabel: {
            operation: { ja: "操作", en: "Operation" },
            scope: { ja: "対象", en: "Target" },
            size: { ja: "サイズ", en: "Size" },
            vertical: { ja: "上下", en: "Vertical" },
            horizontal: { ja: "左右", en: "Horizontal" }
        },
        // 操作（fit/expand）と対象（current/all）の2軸、丸めモード / operation (fit/expand), scope (current/all), rounding mode
        radio: {
            fit: { ja: "オブジェクトに合わせる", en: "Fit to objects" },
            expand: { ja: "アートボードを拡張", en: "Expand artboard" },
            currentArtboard: { ja: "現在のアートボード", en: "Current artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
            roundPixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to pixel grid" },
            roundCurrentUnit: { ja: "現在の単位で値を整数値に", en: "Round values in current unit" },
            roundNone: { ja: "何もしない", en: "Do nothing" }
        },
        checkbox: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document first." },
            enterNumber: { ja: "数値を入力してください。", en: "Please enter a number." },
            errorOccurred: { ja: "エラーが発生しました：", en: "An error occurred: " },
            marginTooLarge: {
                ja: "マージンが大きすぎて有効なサイズにできないため、適用をスキップしました。",
                en: "The margin is too large to produce a valid size; skipped."
            }
        },
        tooltip: {
            fit: {
                ja: "オブジェクトの外接＋マージンのサイズにアートボードを合わせます（選択が無いときは各アートボード内のオブジェクトが対象）",
                en: "Resize artboards to the objects' bounds plus margins (with no selection, each artboard uses the objects it contains)"
            },
            expand: {
                ja: "アートボード自身のサイズにマージンを加減します（マイナス値で縮小）",
                en: "Grow/shrink the artboards themselves by the margins (negative shrinks)"
            },
            currentArtboard: {
                ja: "現在のアートボードのみを対象にします",
                en: "Apply to the current artboard only"
            },
            allArtboards: {
                ja: "すべてのアートボードを対象にします（「合わせる」では各アートボード内のオブジェクトに合わせます）",
                en: "Apply to all artboards (Fit uses the objects each artboard contains)"
            },
            marginInput: {
                ja: "↑↓・∧∨で次の整数へ、Shift+↑↓で10の倍数にスナップ、Option+↑↓で±0.1",
                en: "Arrow/stepper: to the next whole number, Shift: snap to 10, Option: ±0.1"
            },
            axisEnable: {
                ja: "OFFにすると実行時のサイズのまま固定します（連動は自動でOFF）。Option+クリックでこちらだけON",
                en: "Off keeps this dimension at its original size (auto-unlinks). Option-click to solo it"
            },
            linked: {
                ja: "上下の値を左右にも自動で適用します",
                en: "Apply the vertical value to horizontal as well"
            },
            previewBounds: {
                ja: "ON：線・効果を含む見た目の境界（プレビュー境界）で計測／OFF：パスの幾何境界で計測",
                en: "On: measure with preview (visible) bounds incl. strokes/effects; Off: geometric path bounds"
            },
            roundPixelGrid: {
                ja: "座標とサイズを整数ピクセルに丸めます",
                en: "Round position and size to integer pixels"
            },
            roundCurrentUnit: {
                ja: "現在の定規単位で座標とサイズを整数に丸めます",
                en: "Round position and size to integers in the current ruler unit"
            },
            roundNone: {
                ja: "丸めずに計測値のまま設定します",
                en: "Apply the measured values without rounding"
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
        }
    };

    /**
     * ローカライズ文字列を取得する（キー漏れ時は英語へフォールバック）
     * @param {Object} labelSet - { ja, en } のラベル
     * @returns {string} 表示言語の文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        if (labelSet[uiLang] != null) return labelSet[uiLang];
        return (labelSet.en != null) ? labelSet.en : "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - ラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // エラー処理 / Error handling
    // =========================================

    /**
     * Error を行番号・ファイル名付きで読みやすく整形する
     * @param {Error} error - 例外
     * @returns {string} 整形した文字列
     */
    function formatError(error) {
        var messageText = (error && error.message) ? String(error.message) : String(error);
        var lineText = (error && error.line) ? (" line " + error.line) : "";
        var fileText = (error && error.fileName) ? (" (" + error.fileName + ")") : "";
        return messageText + lineText + fileText;
    }

    // =========================================
    // プレビュー管理 / Preview manager
    // =========================================

    /**
     * プレビューの適用と復元を制御するクラス / Preview apply/restore manager
     *
     * - updatePreview() のたびに rollback() で開いた時点の状態へ戻してから addStep() で最新状態を適用
     * - OK/Cancel 時に rollback() してプレビューを開いた時点へ戻す
     *
     * app.undo() の回数に依存すると、複数アートボード書き換え時に undo 粒度とズレて
     * 戻しすぎ/戻し不足が起きる。そのため巻き戻しは restoreFn（スナップショット復元）で行う。
     * @param {function} restoreFn - プレビュー前の状態へ戻す関数
     */
    function PreviewManager(restoreFn) {
        this.restoreFn = restoreFn;
        this.dirty = false; // プレビューによる変更が未復元か / preview changes pending restore

        /**
         * 変更操作を実行し、未復元フラグを立てる
         * @param {function} func - 変更操作
         * @returns {void}
         */
        this.addStep = function (func) {
            try {
                func();
                this.dirty = true;
                app.redraw();
            } catch (e) {
                $.writeln("[PreviewManager] addStep error: " + e);
            }
        };

        /**
         * プレビューを開いた時点の状態へ戻す
         * @returns {void}
         */
        this.rollback = function () {
            if (this.dirty && typeof this.restoreFn === "function") {
                try {
                    this.restoreFn();
                } catch (e) {
                    $.writeln("[PreviewManager] rollback error: " + e);
                }
            }
            this.dirty = false;
            app.redraw();
        };

        /**
         * 現在の状態を確定する（OK時）
         * OK時は一度 rollback() で元に戻してから main() 側で本処理を1回だけ実行するため、ここでは rollback のみ。
         * @returns {void}
         */
        this.confirm = function () {
            this.rollback();
        };
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
    // 設定の記憶（セッション内） / Settings persistence (session only)
    // =========================================
    // $.global に設定を保持。#targetengine のためセッション中は保持されるが、再起動でリセット。
    // Kept in $.global; persists during the session but resets when Illustrator restarts.

    var SETTINGS_KEY = "__FitArtboardWithMargin_Settings";

    /**
     * 保存済みの設定を取得する
     * @returns {Object|null} 設定。無ければ null
     */
    function getStoredSettings() {
        var stored = $.global[SETTINGS_KEY];
        return (stored && typeof stored === "object") ? stored : null;
    }

    /**
     * 設定をセッションに保存する
     * @param {Object} settings - 設定
     * @returns {void}
     */
    function storeSettings(settings) {
        $.global[SETTINGS_KEY] = settings;
    }

    /**
     * 保存済み設定と文脈から、ダイアログの初期値を解決する
     * operation: "fit"（オブジェクトに合わせる、要選択）/ "expand"（アートボードを拡張）
     * scope: "current"（現在のアートボード）/ "all"（すべてのアートボード）
     * @param {string} defaultMargin - 保存が無いときのマージン
     * @param {number} artboardCount - アートボードの数
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @returns {Object} ダイアログの初期値
     */
    function resolveInitialSettings(defaultMargin, artboardCount, hasSelection) {
        var saved = getStoredSettings();

        // 操作：選択があれば fit、無ければ expand を既定に / operation default
        var operation = hasSelection ? "fit" : "expand";
        if (saved && (saved.operation === "fit" || saved.operation === "expand")) {
            operation = saved.operation;
        }

        // 対象：選択なし・複数アートボードなら all、それ以外は current を既定に / scope default
        var scope = (!hasSelection && artboardCount > 1) ? "all" : "current";
        if (saved && (saved.scope === "current" || saved.scope === "all")) {
            scope = saved.scope;
        }
        // 「合わせる」で選択が無いときは、各アートボード内のオブジェクトが対象になるため「すべて」固定
        // Fit without a selection works per artboard, so the scope is locked to all.
        if (operation === "fit" && !hasSelection) scope = "all";

        var link = (saved && typeof saved.link === "boolean") ? saved.link : DIALOG_DEFAULTS.link;
        // 連動ONのときは上下・左右とも有効に揃える / when linked, both axes are enabled
        var verticalEnabled = link ? true : ((saved && typeof saved.verticalEnabled === "boolean") ? saved.verticalEnabled : true);
        var horizontalEnabled = link ? true : ((saved && typeof saved.horizontalEnabled === "boolean") ? saved.horizontalEnabled : true);

        return {
            marginV: (saved && saved.marginV != null) ? saved.marginV : defaultMargin,
            marginH: (saved && saved.marginH != null) ? saved.marginH : defaultMargin,
            link: link,
            verticalEnabled: verticalEnabled,
            horizontalEnabled: horizontalEnabled,
            previewBounds: (saved && typeof saved.previewBounds === "boolean") ? saved.previewBounds : DIALOG_DEFAULTS.previewBounds,
            roundMode: (saved && saved.roundMode) ? saved.roundMode : DIALOG_DEFAULTS.roundMode,
            operation: operation,
            scope: scope
        };
    }

    // =========================================
    // 矩形・境界のユーティリティ / Rect & bounds utilities
    // =========================================

    /**
     * 矩形にマージンを加えた新しい矩形を返す
     * Illustrator の artboardRect は [left, top, right, bottom]（上が大・下が小）。
     * @param {number[]} rect - 元の矩形
     * @param {number} verticalMarginPt - 上下のマージン（pt）
     * @param {number} horizontalMarginPt - 左右のマージン（pt）
     * @returns {number[]} マージンを加えた矩形
     */
    function expandRectByMargin(rect, verticalMarginPt, horizontalMarginPt) {
        return [
            rect[0] - horizontalMarginPt,
            rect[1] + verticalMarginPt,
            rect[2] + horizontalMarginPt,
            rect[3] - verticalMarginPt
        ];
    }

    /**
     * アートボード矩形をピクセルグリッドに最適化する（X/Y/W/H を各1回だけ整数化）
     * X(左)・Y(上)・幅・高さの4値をそれぞれ整数に丸め、右下は X+幅 / Y−高さ で再構成する（二重丸めしない）。
     * @param {number[]} rect - 元の矩形
     * @returns {number[]} 丸めた矩形
     */
    function snapRectToPixelGrid(rect) {
        var x = Math.round(rect[0]);
        var y = Math.round(rect[1]);
        var width = Math.round(rect[2] - rect[0]);
        var height = Math.round(rect[1] - rect[3]);
        return [x, y, x + width, y - height];
    }

    /**
     * 矩形の X/Y/W/H を指定単位で各1回だけ整数化する
     * 単位変換に失敗した場合はピクセルグリッドにフォールバック。
     * @param {number[]} rect - 元の矩形
     * @param {string} unit - 単位の文字列
     * @returns {number[]} 丸めた矩形
     */
    function snapRectToUnitGrid(rect, unit) {
        var ptPerUnit = toPt(1, unit);
        if (isNaN(ptPerUnit) || ptPerUnit === 0) return snapRectToPixelGrid(rect);
        var x = Math.round(rect[0] / ptPerUnit) * ptPerUnit;
        var y = Math.round(rect[1] / ptPerUnit) * ptPerUnit;
        var width = Math.round((rect[2] - rect[0]) / ptPerUnit) * ptPerUnit;
        var height = Math.round((rect[1] - rect[3]) / ptPerUnit) * ptPerUnit;
        return [x, y, x + width, y - height];
    }

    /**
     * 矩形の幅・高さが正か（left<right かつ bottom<top）
     * @param {number[]} rect - 矩形
     * @returns {boolean} 有効なら true
     */
    function isValidRect(rect) {
        return rect[2] > rect[0] && rect[1] > rect[3];
    }

    /**
     * 無効化した軸を元アートボードの座標に固定する（＝その軸は動かさない）
     * 横(左右)OFFなら left/right を、縦(上下)OFFなら top/bottom を元アートボード値に戻す。
     * @param {number[]} rect - 計算した矩形
     * @param {number[]} artboardRect - 元のアートボード矩形
     * @param {boolean} verticalEnabled - 上下（高さ）を調整するか
     * @param {boolean} horizontalEnabled - 左右（幅）を調整するか
     * @returns {number[]} 無効軸を固定した矩形
     */
    function lockDisabledAxes(rect, artboardRect, verticalEnabled, horizontalEnabled) {
        var lockedRect = rect.slice();
        if (!horizontalEnabled) { lockedRect[0] = artboardRect[0]; lockedRect[2] = artboardRect[2]; }
        if (!verticalEnabled) { lockedRect[1] = artboardRect[1]; lockedRect[3] = artboardRect[3]; }
        return lockedRect;
    }

    /**
     * 丸めモードに従って矩形を整数化する
     * "pixelGrid" = ピクセル整数、"currentUnit" = 現在の単位で整数、"none" = 丸めなし。
     * @param {number[]} rect - 元の矩形
     * @param {string} roundMode - 丸めモード
     * @param {string} unit - 単位の文字列
     * @returns {number[]} 丸めた矩形
     */
    function applyRounding(rect, roundMode, unit) {
        if (roundMode === "pixelGrid") return snapRectToPixelGrid(rect);
        if (roundMode === "currentUnit") return snapRectToUnitGrid(rect, unit);
        return rect; // "none"
    }

    /**
     * 2つの矩形が実質的に同一か判定する
     * マージン0などで値が変わらない場合は書き換えを省き、取り消し履歴を増やさないために使う。
     * @param {number[]} rectA - 矩形A
     * @param {number[]} rectB - 矩形B
     * @returns {boolean} 同一なら true
     */
    function rectsEqual(rectA, rectB) {
        if (!rectA || !rectB) return false;
        for (var i = 0; i < 4; i++) {
            if (Math.abs(rectA[i] - rectB[i]) > 0.0001) return false;
        }
        return true;
    }

    /**
     * オブジェクトのバウンディングボックスを取得する
     * usePreviewBounds=true なら visibleBounds（塗り/線を含む）、false なら geometricBounds（パス外形のみ）。
     * @param {PageItem} item - 対象
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]} [left, top, right, bottom]
     */
    function getItemBounds(item, usePreviewBounds) {
        return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
    }

    /**
     * 複数アイテムの外接バウンディングボックスを取得する
     * @param {PageItem[]} items - 対象
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。空なら null
     */
    function getUnionBounds(items, usePreviewBounds) {
        if (!items || items.length === 0) return null;
        var unionBounds = getItemBounds(items[0], usePreviewBounds);
        for (var i = 1; i < items.length; i++) {
            var itemBounds = getItemBounds(items[i], usePreviewBounds);
            unionBounds[0] = Math.min(unionBounds[0], itemBounds[0]);
            unionBounds[1] = Math.max(unionBounds[1], itemBounds[1]);
            unionBounds[2] = Math.max(unionBounds[2], itemBounds[2]);
            unionBounds[3] = Math.min(unionBounds[3], itemBounds[3]);
        }
        return unionBounds;
    }

    /**
     * 計測可能なページアイテムか（geometricBounds を持つか）
     * TextRange 等の非ページアイテムを除外し、選択の型不整合を防ぐ。
     * @param {Object} item - 選択の要素
     * @returns {boolean} 計測できれば true
     */
    function isMeasurableItem(item) {
        if (!item) return false;
        /* TextRange などは geometricBounds で例外になる / non-page items throw on geometricBounds */
        try {
            var bounds = item.geometricBounds;
            return (bounds && bounds.length === 4);
        } catch (e) {
            return false;
        }
    }

    /**
     * クリップグループのクリッピングパスを返す
     * @param {GroupItem} groupItem - クリップグループ
     * @returns {PageItem|null} クリッピングパス。無ければ null
     */
    function getClippingPath(groupItem) {
        try {
            for (var j = 0; j < groupItem.pageItems.length; j++) {
                if (groupItem.pageItems[j].clipping) return groupItem.pageItems[j];
            }
        } catch (e) { /* ignore */ }
        return null;
    }

    /**
     * 選択アイテムを正規化する
     * ・計測できない要素（TextRange 等）は除外
     * ・クリップグループはクリッピングパスのみを採用、それ以外はそのまま
     * @param {Object[]} items - 選択の要素
     * @returns {PageItem[]} 計測に使うアイテム
     */
    function collectEffectiveItems(items) {
        var effectiveItems = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!isMeasurableItem(item)) continue; // 非ページアイテムをスキップ / skip non-page items
            if (item.typename === "GroupItem" && item.clipped) {
                var clippingPath = getClippingPath(item);
                if (clippingPath) effectiveItems.push(clippingPath);
            } else {
                effectiveItems.push(item);
            }
        }
        return effectiveItems;
    }

    /**
     * 計測対象を再帰的に収集する
     * ・TextFrame は複製をアウトライン化（元は不変）。グループ内テキストも対象。
     * ・クリップグループはクリッピングパスのみ。通常グループは中身へ再帰。
     * 一時複製・アウトラインは tempObjects に登録し、呼び出し側で必ず削除する。
     * @param {PageItem[]} items - 対象
     * @param {PageItem[]} tempObjects - 一時オブジェクトの登録先
     * @param {PageItem[]} measureItems - 計測対象の格納先
     * @returns {void}
     */
    function collectMeasureTargets(items, tempObjects, measureItems) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!isMeasurableItem(item)) continue;
            if (item.typename === "TextFrame") {
                // 複製を先に登録 → アウトライン化（失敗時も複製を削除できる） / track clone before outlining
                var clone = item.duplicate();
                tempObjects.push(clone);
                var outlined = clone.createOutline(); // GroupItem を返す / returns a GroupItem
                tempObjects.push(outlined);
                measureItems.push(outlined);
            } else if (item.typename === "GroupItem" && item.clipped) {
                var clippingPath = getClippingPath(item);
                if (clippingPath) measureItems.push(clippingPath);
            } else if (item.typename === "GroupItem") {
                // 通常グループは中身を再帰（ネストされたテキストも非破壊計測） / recurse into groups
                collectMeasureTargets(item.pageItems, tempObjects, measureItems);
            } else {
                measureItems.push(item);
            }
        }
    }

    /**
     * 選択オブジェクトの外接境界を取得する
     * テキストは複製をアウトライン化して計測し、計測後に一時オブジェクトを必ず削除する。
     * 元の TextFrame には一切触れないため、ID・重なり順・名前・タグ・ノート・Variable 等が保持される。
     * @param {PageItem[]} items - 対象
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。空なら null
     */
    function measureSelectionBounds(items, usePreviewBounds) {
        var tempObjects = []; // 計測用に作った一時複製・アウトライン（必ず削除） / temp objects to remove
        try {
            var measureItems = [];
            collectMeasureTargets(items, tempObjects, measureItems);
            return getUnionBounds(measureItems, usePreviewBounds);
        } finally {
            // 途中で例外が起きても一時オブジェクトは必ず削除 / always remove temp objects, even on error
            for (var k = 0; k < tempObjects.length; k++) {
                /* アウトライン化で消費された複製の remove() は例外になる / a clone consumed by createOutline() throws on remove() */
                try { tempObjects[k].remove(); } catch (e) { }
            }
        }
    }

    /**
     * 計測に使えるアイテムか（ロック・非表示・ガイド、およびそのレイヤーを除外）
     * @param {PageItem} item - 対象
     * @returns {boolean} 使えるなら true
     */
    function isUsableItem(item) {
        try {
            if (!item || item.locked || item.hidden || item.guides) return false;
            var parentLayer = item.parent;
            while (parentLayer && parentLayer.typename === "Layer") {
                if (!parentLayer.visible || parentLayer.locked) return false;
                parentLayer = parentLayer.parent;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * アートボードに重なるページアイテムを収集する
     * レイヤー直下のアイテムだけを見る（グループの中身はグループごと1件として扱う）。
     * @param {number[]} artboardRect - アートボード矩形
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {PageItem[]} 重なるアイテム
     */
    function getItemsInArtboard(artboardRect, usePreviewBounds) {
        var doc = app.activeDocument;
        var overlappingItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            try {
                if (item.parent.typename !== "Layer") continue; // 入れ子はグループと一緒に扱う / nested items travel with their group
                if (!isUsableItem(item)) continue;
                var bounds = getItemBounds(item, usePreviewBounds);
                // 一辺でも外れていれば非交差 / no overlap when any edge clears the artboard
                if (bounds[2] <= artboardRect[0] || bounds[0] >= artboardRect[2] || bounds[3] >= artboardRect[1] || bounds[1] <= artboardRect[3]) continue;
                overlappingItems.push(item);
            } catch (e) { /* ignore */ }
        }
        return overlappingItems;
    }

    /**
     * アートボード内オブジェクトの外接境界を取得する
     * @param {number[]} artboardRect - アートボード矩形
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。対象が無ければ null
     */
    function measureArtboardContentBounds(artboardRect, usePreviewBounds) {
        var items = getItemsInArtboard(artboardRect, usePreviewBounds);
        if (items.length === 0) return null;
        return measureSelectionBounds(items, usePreviewBounds);
    }

    // =========================================
    // アートボード矩形の計算 / Artboard rect planning
    // =========================================
    // プレビューと確定で同じ計算（丸め・無効軸固定を含む）を使い、プレビューと結果を一致させる。
    // marginSettings: { verticalPt, horizontalPt, roundMode, unit, verticalEnabled, horizontalEnabled }

    /**
     * マージン適用の共通パイプライン：拡張 → 丸め → 無効軸を元座標に固定
     * @param {number[]} baseRect - 基準の矩形（アートボード自身、またはオブジェクトの外接）
     * @param {number[]} artboardOriginalRect - 実行時のアートボード矩形（無効軸の固定用）
     * @param {Object} marginSettings - マージン適用の設定
     * @returns {number[]} 新しいアートボード矩形
     */
    function computeMarginRect(baseRect, artboardOriginalRect, marginSettings) {
        var rect = expandRectByMargin(baseRect, marginSettings.verticalPt, marginSettings.horizontalPt);
        rect = applyRounding(rect, marginSettings.roundMode, marginSettings.unit);
        return lockDisabledAxes(rect, artboardOriginalRect, marginSettings.verticalEnabled, marginSettings.horizontalEnabled);
    }

    /**
     * 全アートボードの矩形を控える
     * @param {Artboards} artboards - アートボードのコレクション
     * @returns {number[][]} 矩形の配列
     */
    function snapshotArtboardRects(artboards) {
        var rects = [];
        for (var i = 0; i < artboards.length; i++) {
            rects.push(artboards[i].artboardRect.slice());
        }
        return rects;
    }

    /**
     * 操作と対象から、書き換えるアートボードの番号と新しい矩形を求める
     * ・拡張：アートボード自身の矩形にマージンを加減
     * ・合わせる×すべて：各アートボードを、その内側のオブジェクトに合わせる（対象が無いアートボードは据え置き）
     * ・合わせる×現在：現在のアートボードを選択の外接に合わせる
     * @param {string} operation - "fit" / "expand"
     * @param {string} scope - "current" / "all"
     * @param {number} activeIndex - 現在のアートボードの番号
     * @param {number[][]} originalRects - 実行時のアートボード矩形
     * @param {Object} boundsProvider - getContentBounds(index) と getSelectionBounds() を持つ計測元
     * @param {Object} marginSettings - マージン適用の設定
     * @returns {Object[]} { index, rect } の配列
     */
    function planArtboardRects(operation, scope, activeIndex, originalRects, boundsProvider, marginSettings) {
        var targetIndexes = [];
        if (scope === "all") {
            for (var i = 0; i < originalRects.length; i++) targetIndexes.push(i);
        } else {
            targetIndexes.push(activeIndex);
        }

        var rectPlans = [];
        for (var j = 0; j < targetIndexes.length; j++) {
            var artboardIndex = targetIndexes[j];
            var baseRect;
            if (operation === "expand") {
                baseRect = originalRects[artboardIndex];
            } else if (scope === "all") {
                baseRect = boundsProvider.getContentBounds(artboardIndex);
            } else {
                baseRect = boundsProvider.getSelectionBounds();
            }
            if (!baseRect) continue; // 計測対象が無い / nothing to measure
            rectPlans.push({ index: artboardIndex, rect: computeMarginRect(baseRect, originalRects[artboardIndex], marginSettings) });
        }
        return rectPlans;
    }

    /**
     * プレビューとしてアートボード矩形を書き換える
     * 無効な矩形と変化の無い矩形は書き換えない（取り消し履歴のノイズを減らす）。
     * @param {Object[]} rectPlans - planArtboardRects() の結果
     * @returns {void}
     */
    function previewArtboardRects(rectPlans) {
        var artboards = app.activeDocument.artboards;
        for (var i = 0; i < rectPlans.length; i++) {
            var rectPlan = rectPlans[i];
            if (isValidRect(rectPlan.rect) && !rectsEqual(artboards[rectPlan.index].artboardRect, rectPlan.rect)) {
                artboards[rectPlan.index].artboardRect = rectPlan.rect;
            }
        }
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * 調整基準パネルに「項目名＋コントロール群」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} contentOrientation - コントロール群の並び（"column" = 縦 / "row" = 横）
     * @returns {Group} コントロールを入れるグループ
     */
    function addBasisRow(parentPanel, labelSet, contentOrientation) {
        var isColumn = (contentOrientation === "column");
        var basisRow = parentPanel.add("group");
        basisRow.orientation = "row";
        basisRow.alignChildren = ["left", isColumn ? "top" : "center"];
        var rowLabel = basisRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = BASIS_LABEL_WIDTHS[uiLang];
        var contentGroup = basisRow.add("group");
        contentGroup.orientation = contentOrientation;
        contentGroup.alignChildren = isColumn ? "left" : ["left", "center"];
        return contentGroup;
    }

    /**
     * 調整基準パネル（操作＋対象＋サイズ）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildBasisPanel(marginDialog, initialSettings, dialogControls) {
        var basisPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.adjustmentBasis));
        setupPanel(basisPanel);

        /* 操作：オブジェクトに合わせる / アートボードを拡張（ラジオは縦並び） / Operation group (vertical radios) */
        var operationGroup = addBasisRow(basisPanel, LABELS.fieldLabel.operation, "column");
        dialogControls.fitRadio = operationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.fit));
        dialogControls.fitRadio.helpTip = getLabel(LABELS.tooltip.fit);
        dialogControls.expandRadio = operationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.expand));
        dialogControls.expandRadio.helpTip = getLabel(LABELS.tooltip.expand);

        /* 対象：現在のアートボード / すべてのアートボード（ラジオは縦並び） / Scope group (vertical radios) */
        var scopeGroup = addBasisRow(basisPanel, LABELS.fieldLabel.scope, "column");
        dialogControls.currentRadio = scopeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.currentArtboard));
        dialogControls.currentRadio.helpTip = getLabel(LABELS.tooltip.currentArtboard);
        dialogControls.allRadio = scopeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.allArtboards));
        dialogControls.allRadio.helpTip = getLabel(LABELS.tooltip.allArtboards);

        /* サイズ：幅／高さ（OFFにした方は実行時のサイズのまま固定） / Axis targets: width & height (off keeps the original size) */
        var axisGroup = addBasisRow(basisPanel, LABELS.fieldLabel.size, "row");
        axisGroup.spacing = COLUMN_SPACING;
        dialogControls.widthCheckbox = axisGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.width));
        dialogControls.widthCheckbox.value = initialSettings.horizontalEnabled;
        dialogControls.widthCheckbox.helpTip = getLabel(LABELS.tooltip.axisEnable);
        dialogControls.heightCheckbox = axisGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.height));
        dialogControls.heightCheckbox.value = initialSettings.verticalEnabled;
        dialogControls.heightCheckbox.helpTip = getLabel(LABELS.tooltip.axisEnable);

        /* 初期値を適用 / apply initial selection */
        dialogControls.fitRadio.value = (initialSettings.operation === "fit");
        dialogControls.expandRadio.value = (initialSettings.operation !== "fit");
        dialogControls.currentRadio.value = (initialSettings.scope === "current");
        dialogControls.allRadio.value = (initialSettings.scope === "all");
    }

    /**
     * マージンの入力行（項目名＋入力欄）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} initialText - 入力欄の初期値
     * @returns {{label: StaticText, input: EditText}} 項目名と入力欄（∧∨は input.stepperGroup で参照できる）
     */
    function addMarginField(parentGroup, labelSet, initialText) {
        var fieldRow = parentGroup.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        var fieldLabel = fieldRow.add("statictext", undefined, getLabel(labelSet));
        fieldLabel.justify = "right";
        fieldLabel.preferredSize.width = MARGIN_LABEL_WIDTHS[uiLang]; /* 入力欄の位置を揃える / line up the inputs */

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var fieldInput;
        /* 増減後は手入力と同じ処理（連動・プレビュー更新）を通す。onChanging は showMarginDialog() で結線
           after stepping, run the same handler as typing (link and preview) */
        var stepperGroup = addStepper(stepperInputGroup, function () { return fieldInput; }, {
            onStep: function (numberInput) { if (numberInput.onChanging) numberInput.onChanging(); }
        });
        fieldInput = stepperInputGroup.add("edittext", undefined, initialText);
        fieldInput.characters = MARGIN_FIELD_CHARACTERS;
        fieldInput.helpTip = getLabel(LABELS.tooltip.marginInput);
        fieldInput.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(fieldInput, stepperGroup);
        return { label: fieldLabel, input: fieldInput };
    }

    /**
     * マージン入力パネル（入力欄と連動アイコン＋プレビュー境界の2カラム）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {string} rulerUnit - 定規単位
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildMarginPanel(marginDialog, rulerUnit, initialSettings, dialogControls) {
        var marginPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.margin) + " (" + rulerUnit + ")");
        setupPanel(marginPanel);
        marginPanel.orientation = "row";
        marginPanel.alignChildren = ["left", "top"];
        marginPanel.spacing = COLUMN_SPACING;

        /* 上下・左右の2行の右に連動アイコンを置く / Put the link icon to the right of the two margin rows */
        var marginFieldsRow = marginPanel.add("group");
        marginFieldsRow.orientation = "row";
        marginFieldsRow.alignChildren = ["left", "center"];

        var marginFieldsColumn = marginFieldsRow.add("group");
        marginFieldsColumn.orientation = "column";
        marginFieldsColumn.alignChildren = ["left", "center"];

        var optionColumn = marginPanel.add("group");
        optionColumn.orientation = "column";
        optionColumn.alignChildren = ["left", "center"];
        optionColumn.alignment = ["left", "center"];

        /* 上下マージン（「高さ」OFFで無効）・左右マージン（「幅」OFFで無効） / Vertical & horizontal margin inputs */
        var verticalField = addMarginField(marginFieldsColumn, LABELS.fieldLabel.vertical, initialSettings.marginV);
        dialogControls.verticalLabel = verticalField.label;
        dialogControls.verticalInput = verticalField.input;
        var horizontalField = addMarginField(marginFieldsColumn, LABELS.fieldLabel.horizontal,
            initialSettings.link ? initialSettings.marginV : initialSettings.marginH);
        dialogControls.horizontalLabel = horizontalField.label;
        dialogControls.horizontalInput = horizontalField.input;

        /* 連動アイコン。切り替え後の処理は showMarginDialog() で handleLinkToggle に結線
           Link icon; the handler is wired to handleLinkToggle in showMarginDialog() */
        dialogControls.linkToggle = addLinkToggle(marginFieldsRow, initialSettings.link, function () {
            if (dialogControls.handleLinkToggle) dialogControls.handleLinkToggle();
        });
        dialogControls.linkToggle.helpTip = getLabel(LABELS.tooltip.linked);

        /* プレビュー境界（visibleBounds を採用するか。「合わせる」のときのみ有効） / use visibleBounds; only for fit */
        dialogControls.previewBoundsCheckbox = optionColumn.add("checkbox", undefined, getLabel(LABELS.checkbox.previewBounds));
        dialogControls.previewBoundsCheckbox.value = initialSettings.previewBounds;
        dialogControls.previewBoundsCheckbox.helpTip = getLabel(LABELS.tooltip.previewBounds);
        dialogControls.previewBoundsCheckbox.enabled = (initialSettings.operation === "fit");
    }

    /**
     * アートボードサイズの微調整パネル（丸めモード）を作る
     * プレビューにも確定と同じ丸めを反映する。
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildFineTuningPanel(marginDialog, initialSettings, dialogControls) {
        var fineTuningPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.fineTuning));
        setupPanel(fineTuningPanel);

        var roundModeGroup = fineTuningPanel.add("group");
        roundModeGroup.orientation = "column";
        roundModeGroup.alignChildren = "left";
        dialogControls.roundPixelRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundPixelGrid));
        dialogControls.roundUnitRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundCurrentUnit));
        dialogControls.roundNoneRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundNone));
        dialogControls.roundPixelRadio.helpTip = getLabel(LABELS.tooltip.roundPixelGrid);
        dialogControls.roundUnitRadio.helpTip = getLabel(LABELS.tooltip.roundCurrentUnit);
        dialogControls.roundNoneRadio.helpTip = getLabel(LABELS.tooltip.roundNone);
        dialogControls.roundUnitRadio.value = (initialSettings.roundMode === "currentUnit");
        dialogControls.roundNoneRadio.value = (initialSettings.roundMode === "none");
        dialogControls.roundPixelRadio.value = !dialogControls.roundUnitRadio.value && !dialogControls.roundNoneRadio.value; // 既定 / default
    }

    /**
     * ボタン行（左右中央：Cancel → OK の順）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildButtonRow(marginDialog, dialogControls) {
        var btnRowGroup = marginDialog.add("group");
        setupRow(btnRowGroup, "center");
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        dialogControls.btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        dialogControls.btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    // =========================================
    // ダイアログの状態 / Dialog state
    // =========================================

    /**
     * 選んでいる操作を返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "fit" / "expand"
     */
    function getOperation(dialogControls) {
        return dialogControls.fitRadio.value ? "fit" : "expand";
    }

    /**
     * 選んでいる対象を返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "current" / "all"
     */
    function getScope(dialogControls) {
        return dialogControls.allRadio.value ? "all" : "current";
    }

    /**
     * 選んでいる丸めモードを返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "pixelGrid" / "currentUnit" / "none"
     */
    function getRoundMode(dialogControls) {
        if (dialogControls.roundPixelRadio.value) return "pixelGrid";
        return dialogControls.roundUnitRadio.value ? "currentUnit" : "none";
    }

    /**
     * 「合わせる」で選択が無いときは対象を「すべて」に固定してディムする
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @returns {void}
     */
    function refreshScopeState(dialogControls, hasSelection) {
        var forceAll = (getOperation(dialogControls) === "fit" && !hasSelection);
        if (forceAll) {
            dialogControls.allRadio.value = true;
            dialogControls.currentRadio.value = false;
        }
        dialogControls.currentRadio.enabled = !forceAll;
    }

    /**
     * マージンの入力欄の有効／無効を、∧∨ごとまとめて切り替える
     * @param {EditText} fieldInput - addMarginField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setMarginInputEnabled(fieldInput, isEnabled) {
        fieldInput.enabled = isEnabled;
        fieldInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(fieldInput.stepperGroup);
    }

    /**
     * 入力欄と連動アイコンの有効/無効を現在の状態から更新する
     * 上下=「高さ」ON、左右=「幅」ON かつ 非連動、連動=幅・高さが両方ONのときだけ有効。
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {void}
     */
    function refreshMarginInputStates(dialogControls) {
        var heightOn = dialogControls.heightCheckbox.value;
        var widthOn = dialogControls.widthCheckbox.value;
        dialogControls.verticalLabel.enabled = heightOn;
        setMarginInputEnabled(dialogControls.verticalInput, heightOn);
        dialogControls.horizontalLabel.enabled = widthOn;
        setMarginInputEnabled(dialogControls.horizontalInput, widthOn && !dialogControls.linkToggle.value);
        // 幅・高さのどちらかがOFFなら連動は使えない（自動OFFのうえディム） / disable link when either axis is off
        setLinkToggleEnabled(dialogControls.linkToggle, heightOn && widthOn);
    }

    /**
     * マージン欄の文字列を pt に変換する。無効な軸は 0 とみなす
     * @param {boolean} axisEnabled - その軸を調整するか
     * @param {string} marginText - 入力欄の文字列
     * @param {string} unit - 単位の文字列
     * @returns {number} pt。数値でなければ NaN
     */
    function marginTextToPoints(axisEnabled, marginText, unit) {
        return axisEnabled ? toPt(parseFloat(marginText), unit) : 0;
    }

    /**
     * ダイアログの入力から、マージン適用の設定を作る
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {string} rulerUnit - 定規単位
     * @returns {Object} マージン適用の設定。valid は有効な軸の値がすべて数値のとき true
     */
    function readMarginSettings(dialogControls, rulerUnit) {
        var verticalEnabled = dialogControls.heightCheckbox.value;
        var horizontalEnabled = dialogControls.widthCheckbox.value;
        var verticalPt = marginTextToPoints(verticalEnabled, dialogControls.verticalInput.text, rulerUnit);
        var horizontalPt = marginTextToPoints(horizontalEnabled, dialogControls.horizontalInput.text, rulerUnit);
        return {
            valid: !isNaN(verticalPt) && !isNaN(horizontalPt),
            verticalPt: verticalPt,
            horizontalPt: horizontalPt,
            roundMode: getRoundMode(dialogControls),
            unit: rulerUnit,
            verticalEnabled: verticalEnabled,
            horizontalEnabled: horizontalEnabled
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * マージン入力ダイアログを表示し設定を返す（ライブプレビュー付き）
     * @param {string} defaultMargin - 保存が無いときのマージン
     * @param {string} rulerUnit - 定規単位
     * @param {number} artboardCount - アートボードの数
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @param {PageItem[]} selectionItems - ダイアログ表示時に固定した選択アイテム（「合わせる」で使用）
     * @returns {Object|null} { operation, scope, previewBounds, marginSettings }。キャンセル時は null
     */
    function showMarginDialog(defaultMargin, rulerUnit, artboardCount, hasSelection, selectionItems) {
        var marginDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);

        setupWindow(marginDialog);
        // 初めて開くときは中央から横にずらす（前回の位置があれば prepareDialogWindow が上書き） / first run: offset from center
        marginDialog.onShow = function () {
            marginDialog.location = [marginDialog.location[0] + DIALOG_FIRST_RUN_OFFSET_X, marginDialog.location[1]];
        };

        // 保存済み設定（セッション内）から初期値を解決 / resolve initial values from stored settings
        var initialSettings = resolveInitialSettings(defaultMargin, artboardCount, hasSelection);

        var dialogControls = {};
        buildBasisPanel(marginDialog, initialSettings, dialogControls);
        buildMarginPanel(marginDialog, rulerUnit, initialSettings, dialogControls);
        buildFineTuningPanel(marginDialog, initialSettings, dialogControls);
        buildButtonRow(marginDialog, dialogControls);
        refreshScopeState(dialogControls, hasSelection);
        refreshMarginInputStates(dialogControls);

        var verticalInput = dialogControls.verticalInput;
        var horizontalInput = dialogControls.horizontalInput;
        var widthCheckbox = dialogControls.widthCheckbox;
        var heightCheckbox = dialogControls.heightCheckbox;
        var linkToggle = dialogControls.linkToggle;
        var previewBoundsCheckbox = dialogControls.previewBoundsCheckbox;

        /* 現在／全アートボードの rect を保存（プレビュー復元用） / Save artboard rects for preview restore */
        var artboards = app.activeDocument.artboards;
        var activeArtboardIndex = artboards.getActiveArtboardIndex();
        var originalArtboardRects = snapshotArtboardRects(artboards);

        /* プレビュー復元：開いた時点の全 artboardRect へ書き戻す（undo回数に依存しない） / Restore all artboard rects to the opening snapshot */
        var previewManager = new PreviewManager(function () {
            var currentArtboards = app.activeDocument.artboards;
            for (var i = 0; i < currentArtboards.length && i < originalArtboardRects.length; i++) {
                if (!rectsEqual(currentArtboards[i].artboardRect, originalArtboardRects[i])) {
                    currentArtboards[i].artboardRect = originalArtboardRects[i];
                }
            }
        });

        /* 計測結果のキャッシュ（元の矩形と固定した選択で計測するのでダイアログ表示中は不変）
           Measured bounds cache; stable while the dialog is open */
        var boundsCache = {};
        var previewBoundsProvider = {
            getContentBounds: function (index) {
                var cacheKey = index + (previewBoundsCheckbox.value ? ":v" : ":g");
                if (!boundsCache.hasOwnProperty(cacheKey)) {
                    boundsCache[cacheKey] = measureArtboardContentBounds(originalArtboardRects[index], previewBoundsCheckbox.value);
                }
                return boundsCache[cacheKey];
            },
            getSelectionBounds: function () {
                var cacheKey = previewBoundsCheckbox.value ? "v" : "g";
                if (!boundsCache.hasOwnProperty(cacheKey)) {
                    boundsCache[cacheKey] = (selectionItems && selectionItems.length > 0) ?
                        measureSelectionBounds(selectionItems, previewBoundsCheckbox.value) : null;
                }
                return boundsCache[cacheKey];
            }
        };

        /* プレビュー更新：直前分を rollback してから最新状態を1回だけ適用 / Refresh preview via PreviewManager */
        function updatePreview() {
            previewManager.rollback();

            var marginSettings = readMarginSettings(dialogControls, rulerUnit);
            if (!marginSettings.valid) return;
            var operation = getOperation(dialogControls);
            var scope = getScope(dialogControls);

            previewManager.addStep(function () {
                previewArtboardRects(planArtboardRects(operation, scope, activeArtboardIndex, originalArtboardRects, previewBoundsProvider, marginSettings));
            });
        }

        /* 操作/対象の変更時：対象の固定・プレビュー境界の有効/無効を切り替えてプレビュー更新 / On change: refresh scope & previewBounds */
        function onBasisChange() {
            refreshScopeState(dialogControls, hasSelection); // 選択なしの「合わせる」は対象を「すべて」に固定 / fit without selection locks scope to all
            previewBoundsCheckbox.enabled = (getOperation(dialogControls) === "fit"); // 合わせる時のみ有効 / only for fit
            updatePreview();
        }

        /* 入力・ラジオ・連動のハンドラ登録（∧∨・↑↓キーも onChanging を通る） / Wire up input handlers (steppers and arrow keys go through onChanging) */
        verticalInput.onChanging = function () {
            if (linkToggle.value) horizontalInput.text = verticalInput.text;
            updatePreview();
        };
        horizontalInput.onChanging = function () {
            if (linkToggle.value) return; // 連動中は水平の直接編集は無効
            updatePreview();
        };

        /* Option(Alt)クリック検出：mousedown で event.altKey を捕捉（onClick では修飾キーを取得できない） / capture Alt on mousedown */
        var axisSoloRequested = { width: false, height: false };
        widthCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.width = (event.altKey === true); });
        heightCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.height = (event.altKey === true); });

        /**
         * 幅・高さチェックの変更処理
         * Option+クリック＝ソロ（クリックした方のみON、もう片方OFF）。どちらかOFFなら連動を自動OFF。
         * @param {boolean} solo - Option+クリックか
         * @param {Checkbox} clickedCheckbox - クリックしたチェックボックス
         * @param {Checkbox} otherCheckbox - もう片方のチェックボックス
         * @returns {void}
         */
        function handleAxisToggle(solo, clickedCheckbox, otherCheckbox) {
            if (solo) {
                clickedCheckbox.value = true;   // クリックした軸をON / keep the clicked axis on
                otherCheckbox.value = false;    // もう片方をOFF / turn the other off
            }
            if (!heightCheckbox.value || !widthCheckbox.value) setLinkToggleValue(linkToggle, false);
            refreshMarginInputStates(dialogControls);
            updatePreview();
        }
        widthCheckbox.onClick = function () {
            var solo = axisSoloRequested.width; axisSoloRequested.width = false;
            handleAxisToggle(solo, widthCheckbox, heightCheckbox);
        };
        heightCheckbox.onClick = function () {
            var solo = axisSoloRequested.height; axisSoloRequested.height = false;
            handleAxisToggle(solo, heightCheckbox, widthCheckbox);
        };

        dialogControls.handleLinkToggle = function () {
            if (linkToggle.value) {
                // 連動ON：幅・高さを有効に揃え、値を上下に統一 / linking re-enables both axes and mirrors V→H
                heightCheckbox.value = true;
                widthCheckbox.value = true;
                horizontalInput.text = verticalInput.text;
            }
            refreshMarginInputStates(dialogControls);
            updatePreview();
        };
        previewBoundsCheckbox.onClick = updatePreview;
        dialogControls.fitRadio.onClick = onBasisChange;
        dialogControls.expandRadio.onClick = onBasisChange;
        dialogControls.currentRadio.onClick = onBasisChange;
        dialogControls.allRadio.onClick = onBasisChange;
        dialogControls.roundPixelRadio.onClick = updatePreview;
        dialogControls.roundUnitRadio.onClick = updatePreview;
        dialogControls.roundNoneRadio.onClick = updatePreview;

        var dialogResult = null;
        dialogControls.btnOK.onClick = function () {
            var marginSettings = readMarginSettings(dialogControls, rulerUnit);
            if (!marginSettings.valid) {
                alert(getLabel(LABELS.alert.enterNumber));
                return;
            }
            dialogResult = {
                operation: getOperation(dialogControls),
                scope: getScope(dialogControls),
                previewBounds: previewBoundsCheckbox.value,
                marginSettings: marginSettings
            };
            // 設定をセッションに保存（次回の初期値に） / store settings for next run
            storeSettings({
                marginV: verticalInput.text,
                marginH: horizontalInput.text,
                link: linkToggle.value,
                verticalEnabled: marginSettings.verticalEnabled,
                horizontalEnabled: marginSettings.horizontalEnabled,
                previewBounds: dialogResult.previewBounds,
                roundMode: marginSettings.roundMode,
                operation: dialogResult.operation,
                scope: dialogResult.scope
            });
            // プレビューを開いた時点へ復元してから閉じる / restore to the opening snapshot, then close
            previewManager.confirm();
            marginDialog.close(1);
        };
        dialogControls.btnCancel.onClick = function () {
            // キャンセル時は必ずロールバックして閉じる / rollback preview and close
            previewManager.rollback();
            marginDialog.close(0);
        };

        updatePreview();

        // 開いたら有効な方のマージン入力欄にフォーカス（上下→無ければ左右） / focus a margin field on open
        if (verticalInput.enabled) {
            verticalInput.active = true;
        } else if (horizontalInput.enabled) {
            horizontalInput.active = true;
        }

        prepareDialogWindow(marginDialog, SCRIPT_NAME);
        marginDialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 対象を決定し、ダイアログの設定に従ってアートボードを調整する
     * @returns {void}
     */
    function main() {
        try {
            // ドキュメント未オープンなら分かりやすく案内して終了 / friendly guard when no document is open
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }
            var doc = app.activeDocument;

            // ダイアログ表示時点の選択を固定（「合わせる」で共通利用） / freeze selection at dialog time
            var selectionItems = collectEffectiveItems(doc.selection);
            var hasSelection = selectionItems.length > 0; // 計測可能な選択があるか / measurable selection?

            var artboards = doc.artboards;
            var rulerUnit = getRulerUnitString();

            /* 単位ごとの初期マージン値。選択なし・複数アートボード時は0（保存済み設定があればそちらが優先される）
               Default margin for the unit; 0 for the expand-all default (stored settings win) */
            var defaultMarginValue = (!hasSelection && artboards.length > 1) ? '0' : getDefaultMargin(rulerUnit);

            var dialogResult = showMarginDialog(defaultMarginValue, rulerUnit, artboards.length, hasSelection, selectionItems);
            if (!dialogResult) return; // キャンセル時は選択ツールにも切り替えない / cancel: leave tool unchanged

            var originalRects = snapshotArtboardRects(artboards);
            var previewBounds = dialogResult.previewBounds;
            var confirmBoundsProvider = {
                getContentBounds: function (index) { return measureArtboardContentBounds(originalRects[index], previewBounds); },
                getSelectionBounds: function () { return measureSelectionBounds(selectionItems, previewBounds); }
            };
            var rectPlans = planArtboardRects(dialogResult.operation, dialogResult.scope, artboards.getActiveArtboardIndex(),
                originalRects, confirmBoundsProvider, dialogResult.marginSettings);

            // 負マージン等で無効サイズになるものは適用しない / skip rects that end up invalid
            var skippedCount = 0;
            for (var i = 0; i < rectPlans.length; i++) {
                if (isValidRect(rectPlans[i].rect)) artboards[rectPlans[i].index].artboardRect = rectPlans[i].rect;
                else skippedCount++;
            }
            if (skippedCount > 0) alert(getLabel(LABELS.alert.marginTooLarge));

            // 適用が完了したときのみ選択ツールへ切り替え（キャンセル・エラー時は切り替えない） / switch tool only after a successful run
            app.selectTool("Adobe Select Tool");

        } catch (e) {
            $.writeln("[FitArtboardWithMargin] ERROR: " + formatError(e));
            alert(getLabel(LABELS.alert.errorOccurred) + formatError(e));
        }
    }

    main();

})();
