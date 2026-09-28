#targetengine "SmartGridMakerEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

長方形の選択、またはアートボードを基準に、囲み罫とグリッドを一括生成します。
外枠・タイトルエリア・内側エリアの分割や線種、裁ち落とし対応のフレームを、プレビューを見ながら1つのダイアログで設定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartGridMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n2b01f896c423

### Overview

Builds a border and a grid from a selected rectangle, or from the artboard.
The outer frame, the title area, the inner-area divisions, the line types, and a bleed-aware frame are all set in one dialog with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartGridMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartGridMaker";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartGridMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartGridMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2b01f896c423"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 生成の設定 / Generation settings
    // =========================================

    /* 生成と既定値の設定（長さの既定値はmmで持ち、現在の定規単位に換算して使う）
       Generation and default values; lengths are in mm and converted to the current ruler unit */
    var GENERATION_SETTINGS = {
        bleedMm: 3,                /* 裁ち落とし幅（mm） / bleed width in mm */
        innerOffsetDivisor: 40,    /* 内側オフセット初期値＝(幅+高さ)/この値 / inner offset default divisor */
        titleSizeDivisor: 5,       /* タイトルエリア初期値＝外側エリアの幅または高さ/この値 / title size default divisor */
        maxGridCount: 100,         /* 列数・行数の上限 / maximum number of columns and rows */
        defaultMarginMm: 15,       /* マージンの初期値（mm） / default margin */
        defaultEdgeScaleMm: -5,    /* 辺の伸縮の初期値（mm） / default edge scale */
        defaultFrameWidthMm: 10,   /* フレームON時の既定幅（mm） / frame width applied when enabled */
        defaultRoundMm: 2          /* 角丸ON時の既定値（mm） / radius applied when rounding is enabled */
    };

    /* 生成物のグレーの濃度（CMYKのK%） / Gray tints of generated items (K in CMYK) */
    var GRAY_TINTS = {
        innerCell: 15,  /* 内側エリアのセル / inner area cells */
        titleBand: 30,  /* タイトル帯 / title band */
        frame: 50,      /* フレーム / frame */
        rule: 100       /* 罫線 / rules */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログとパネルの外観 / Dialog and panel appearance */
    var DIALOG_LAYOUT = {
        dialogOffsetX: 300,       /* ダイアログの表示位置オフセットX / dialog offset X */
        dialogOffsetY: 0,         /* ダイアログの表示位置オフセットY / dialog offset Y */
        tabSize: [300, 460],      /* タブパネルの最小サイズ / minimum size of the tabbed panel */
        panelMargins: [15, 20, 15, 10], /* パネル余白 [左,上,右,下] / panel margins */
        viewLabelWidth: 58,       /* 画面表示タブのラベル幅 / label width in the Display tab */
        viewSliderWidth: 200,     /* 画面表示タブのスライダー幅 / slider width in the Display tab */
        viewButtonWidth: 220,     /* 表示コマンドボタンの幅 / view command button width */
        viewButtonHeight: 22,     /* 表示コマンドボタンの高さ / view command button height */
        toggleCheckboxWidth: 20   /* ラベルなしチェックボックスの幅（空文字ぶんの余白を抑える）
                                     width of the label-less checkboxes, to drop the phantom text space */
    };

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。使わない関数も消さずに残してよい（互いに呼び合う）。
    //    識別子は SELECTION_ITEMS_TOLERANCE / normalizeSelectionItems / resolveTextRangeFrame /
    //    collectSelectionItems / getTextFrameKindKey / collectSelectionTextFrames / collectSelectionPathItems /
    //    isClipMaskItem / getClipMaskItem / hasClippedDescendant / readUsePreviewBoundsPreference /
    //    getClipAwareBounds / filterMeasurableChildren / getClipAwareUnionBounds / isNearlySameCoordinate / areBoundsNearlyEqual
    // 2. 選択は normalizeSelectionItems(doc.selection) で配列にする。文字カーソルの選択（TextRange）は
    //    配列ではなく1個で返り、しかも .length（文字数）を持つので、length だけで配列と見なさない
    // 3. テキストフレーム:
    //      var frames = collectSelectionTextFrames(doc.selection);                           // 全種類
    //      var frames = collectSelectionTextFrames(doc.selection, { kinds: ["point", "path"] });
    //    パス:
    //      var paths = collectSelectionPathItems(doc.selection);                             // 複合パスは中のパスへ
    //      var paths = collectSelectionPathItems(doc.selection, { compoundPaths: "whole", skipClipMasks: true });
    //    それ以外は collectSelectionItems(source, { accept: function (item) { … } }) で条件を書く
    // 4. 並びは選択と同じ前面→背面（グループの中も pageItems の順）。重なり順を使う処理はこの順を前提にしてよい
    // 5. doc.selection に代入し直す配列は skipLocked / skipHidden を true にする。
    //    ロック・非表示を選択に代入すると例外になり、中の子が選択に残る
    // 6. 境界は getClipAwareBounds(item, usePreviewBounds) / getClipAwareUnionBounds(items, usePreviewBounds)。
    //    usePreviewBounds を省くと環境設定の［プレビュー境界を使用］に従う。返り値は [左, 上, 右, 下] の新しい配列
    //    （書き換えても元のオブジェクトに影響しない）。測れないときは null
    // 7. 座標の一致・前後の判定は isNearlySameCoordinate / areBoundsNearlyEqual で許容値を挟む
    //    （吸着させた辺とガイドは 1e-12 ほどずれる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "囲み罫とグリッド", en: "Border and Grid" }
        },
        tab: {
            margin: { ja: "アートボード", en: "Artboard" },
            outer: { ja: "外枠", en: "Outer" },
            inner: { ja: "内側エリア", en: "Inner Area" },
            display: { ja: "画面表示", en: "Display" }
        },
        panel: {
            margin: { ja: "マージン", en: "Margin" },
            frame: { ja: "フレーム", en: "Frame" },
            outer: { ja: "外側エリア", en: "Outer Area" },
            strokeCap: { ja: "線端", en: "Line Caps" },
            titleArea: { ja: "タイトルエリア", en: "Title Area" },
            innerArea: { ja: "内側エリア", en: "Inner Area" },
            offset: { ja: "オフセット", en: "Offset" },
            columns: { ja: "列", en: "Columns" },
            rows: { ja: "行", en: "Rows" },
            lineType: { ja: "線の種類", en: "Line Type" },
            zoomPan: { ja: "ズームとパン", en: "Zoom & Pan" },
            viewCommands: { ja: "表示コマンド", en: "View Commands" }
        },
        checkbox: {
            keepOuter: { ja: "外枠を残す", en: "Keep outer frame" },
            edgeScale: { ja: "辺の伸縮", en: "Extend edges" },
            round: { ja: "角丸", en: "Round" },
            fill: { ja: "塗り", en: "Fill" },
            titleDivider: { ja: "仕切り線", en: "Divider" },
            dividerScale: { ja: "線の伸縮", en: "Extend divider" },
            bleed: { ja: "裁ち落とし", en: "Bleed" },
            divider: { ja: "分割線", en: "Dividers" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        radio: {
            capButt: { ja: "なし", en: "Butt" },
            capRound: { ja: "丸型", en: "Round" },
            capProject: { ja: "突出", en: "Projecting" },
            lineSolid: { ja: "実線", en: "Solid" },
            lineDash: { ja: "破線", en: "Dashed" },
            lineDots: { ja: "点線", en: "Dotted" },
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" }
        },
        fieldLabel: {
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" },
            width: { ja: "幅", en: "Width" },
            titleSize: { ja: "幅／高さ", en: "Size" },
            columnCount: { ja: "列数", en: "Columns" },
            rowCount: { ja: "行数", en: "Rows" },
            spacing: { ja: "間隔", en: "Spacing" },
            zoom: { ja: "ズーム", en: "Zoom" },
            panX: { ja: "左右", en: "Pan L/R" },
            panY: { ja: "上下", en: "Pan U/D" }
        },
        button: {
            fitArtboard: { ja: "アートボードを全体表示", en: "Fit Artboard in Window" },
            actualSize: { ja: "100%表示", en: "Actual Size" },
            fitAll: { ja: "すべてのアートボードを全体表示", en: "Fit All in Window" },
            zoomOut10: { ja: "10%縮小", en: "Zoom Out 10%" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            keepOuter: {
                ja: "OFFにすると、基準にした長方形を削除します",
                en: "Turning this off removes the rectangle used as the base"
            },
            outerRound: {
                ja: "外枠の角を丸めます（ライブエフェクト）。タイトル帯の角にも同じ値を使います。\n辺の伸縮とは同時に使えません",
                en: "Rounds the corners of the outer frame (live effect); the title band reuses the same radius.\nCannot be combined with the edge extension"
            },
            edgeScale: {
                ja: "＋で各辺を両端から伸ばし、−で縮めます。0以外にすると外枠を4本の線に分解します",
                en: "Positive extends each edge at both ends, negative shortens it. Any value but zero splits the outer frame into four lines"
            },
            strokeCap: {
                ja: "辺を4本の線に分解したときだけ有効です",
                en: "Only available while the edges are split into four lines"
            },
            titleArea: {
                ja: "タイトル用の帯を、外枠の内側に作ります",
                en: "Adds a title band inside the outer frame"
            },
            titleSize: {
                ja: "上・下に置くときは高さ、左・右に置くときは幅です",
                en: "A height for a title at the top or bottom, a width for one at the left or right"
            },
            titleFill: {
                ja: "タイトル帯をK30%で塗ります",
                en: "Fills the title band with 30% black"
            },
            titleDivider: {
                ja: "タイトル帯と本文の境目に線を引きます",
                en: "Draws a line between the title band and the body"
            },
            dividerScale: {
                ja: "＋で仕切り線の両端を短く、−で長くします",
                en: "Positive shortens the divider at both ends, negative extends it"
            },
            frame: {
                ja: "アートボードの外周に太い帯を作ります（内側は穴あき）",
                en: "Adds a thick band around the artboard, with the inside cut out"
            },
            bleed: {
                ja: "フレームを裁ち落とし（3mm）まで広げます",
                en: "Extends the frame out to the bleed (3 mm)"
            },
            frameRound: {
                ja: "内側（穴）の角を丸めます",
                en: "Rounds the corners of the inner cutout"
            },
            offset: {
                ja: "外枠（タイトルエリアを除く）から内側エリアまでの距離",
                en: "Distance from the outer frame, excluding the title area, to the inner area"
            },
            spacing: {
                ja: "列／行が2以上のときに有効です",
                en: "Only available for two or more columns or rows"
            },
            innerFill: {
                ja: "各セルをK15%で塗ります。OFFのときは実行時に削除されます",
                en: "Fills each cell with 15% black; the cells are dropped on OK while this is off"
            },
            divider: {
                ja: "間隔の中央に線を1本ずつ引きます",
                en: "Draws one line at the centre of each gutter"
            },
            link: {
                ja: "上の値を下・左・右にも使います",
                en: "Uses the top value for the bottom, left and right as well"
            },
            numberInput: {
                ja: "↑↓で次の整数へ、Shift＋↑↓で次の10の倍数へ、Option＋↑↓で±0.1",
                en: "Up/Down: to the next whole number, Shift+Up/Down: to the next multiple of 10, Option+Up/Down: ±0.1"
            },
            viewSlider: {
                ja: "Optionキーを押しながらドラッグすると微調整できます",
                en: "Hold Option while dragging for fine adjustment"
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
        alert: {
            baseRectFailed: {
                ja: "アートボードを基準にする長方形を作成できませんでした。\nレイヤーのロックや非表示を解除してから実行してください。",
                en: "The rectangle used as the artboard base could not be created.\nUnlock the layer and make it visible, then run the script again."
            }
        }
    };
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
     * mmをptに換算します。
     *
     * @param {number} mm - mmの値。
     * @returns {number} ptの値。
     */
    function mmToPt(mm) {
        return UNITS[1].pointsPerUnit * mm;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は SETTINGS_STORE_* / createSettingsStore / readSettingsLegacyFile / readSettingsLegacyPreference / settingsStore*
    // 2. 寿命は今のスクリプトに合わせて選ぶ。
    //      "session"    … $.global に置く。Illustrator を終了するまで残る。#targetengine が必須（無いと毎回消える）
    //      "persistent" … Folder.userData/illustrator-scripts/<storeName>.json に書く。再起動しても残る
    //    storeName はふつう SCRIPT_NAME。ダイアログの位置は DialogPosition の部品が持つので、ここには入れない
    // 3. 既定値を1か所にまとめ、load で受け取る。戻り値は毎回新しいオブジェクト（書き換えても保存されない）
    //      var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
    //      var DEFAULT_SETTINGS = { widthPt: 10, addFrame: true, modeKey: "fit", corners: { tl: 0, tr: 0 } };
    //      var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
    //      …OK で閉じたら…
    //      settingsStore.save({ widthPt: …, addFrame: …, modeKey: …, corners: { tl: …, tr: … } });
    //    型は既定値に合わせる（数値の既定値には "12" も 12 として読む。真偽は "1"/"0"/"true"/"false" も読む）。
    //    合わない値・既定値に無い項目は捨てて既定値を使う。{} と null の既定値は中身を問わずそのまま受け取る
    //    （名前をキーにしたプリセット集など）。配列は配列ならそのまま受け取る
    // 4. 保存できるのは文字列・数値・真偽・null と、その配列・入れ子のオブジェクトだけ。
    //    DOM オブジェクト・File・関数は入れない（パスは fsName の文字列で持つ）。長さは pt で持つ
    // 5. 旧形式の設定を読み継ぐときは、3つ目の引数に legacy 関数を渡す。
    //    新しい保存が1度も無いとき（ファイルが無い・$.global に無い）だけ呼ばれ、戻り値を保存値として既定値と突き合わせる。
    //    旧ファイル・旧キーは消さない。キー名が変わったときは legacy の中で詰め替える
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");  … key=value / toSource / JSON を自動判別
    //      } });
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyPreference("SmartTextFindReplace/settings");  … app.preferences の文字列
    //      } });
    // 6. clear() は保存を消す。legacy を渡したストアでは空の保存（{}）を書き、旧設定が戻ってこないようにする
    // 7. 失敗は例外にせず、load は既定値、save は false を返す（$.writeln に理由を出す）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // セッション状態 / Session state
    // =========================================
    /* Illustratorの起動中だけダイアログの値を保持する（#targetengine が必須）。旧版の $.global のキーを1度だけ読み継ぐ
       Dialog values are kept only while Illustrator is running (needs #targetengine); the old $.global key is read once */

    var LEGACY_SESSION_STATE_KEY = "__SmartGridMaker__";
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_STATE_KEY] || null; }
    });

    /**
     * 前回のダイアログ設定を読み込みます。
     * 項目はダイアログ側の定義表（ドット区切りのパス）で決まるため、既定値は {}（中身を問わず受け取る）
     *
     * @returns {Object} 保存されていた状態。未保存なら空のオブジェクト。
     */
    function loadSessionState() {
        return settingsStore.load({});
    }

    /**
     * 現在のダイアログ設定を保存します。
     *
     * @param {Object} state - 保存する状態。
     * @returns {void}
     */
    function saveSessionState(state) {
        settingsStore.save(state);
    }

    /**
     * 入れ子のオブジェクトから、ドット区切りのパスで値を取り出します。
     *
     * @param {Object} source - 探索するオブジェクト。
     * @param {string} path - 例 "inner.grid.columns" のようなパス。
     * @returns {*} 見つかった値。存在しない場合は undefined。
     */
    function getStateValue(source, path) {
        var segments = path.split(".");
        var node = source;

        for (var i = 0; i < segments.length; i++) {
            if (node == null) return undefined;
            node = node[segments[i]];
        }
        return node;
    }

    /**
     * 入れ子のオブジェクトに、ドット区切りのパスで値を書き込みます（中間オブジェクトは自動作成）。
     *
     * @param {Object} target - 書き込み先のオブジェクト。
     * @param {string} path - 例 "inner.grid.columns" のようなパス。
     * @param {*} value - 書き込む値。
     * @returns {void}
     */
    function setStateValue(target, path, value) {
        var segments = path.split(".");
        var node = target;

        for (var i = 0; i < segments.length - 1; i++) {
            if (!node[segments[i]]) node[segments[i]] = {};
            node = node[segments[i]];
        }
        node[segments[segments.length - 1]] = value;
    }

    // =========================================
    // ズームとパン / ViewControl
    // =========================================
    /* 画面表示のズームとパン（左右・上下）をスライダーで操作する部品（単体で切り出して使える）
       Zoom and pan sliders for the Illustrator view; self-contained so it can be extracted

       - パンの基準はアクティブなアートボードの中心、可動範囲はアートボードの半分
         Pan is relative to the active artboard centre, within half the artboard size
       - ズームしてもパン量は保つ / Zooming keeps the pan offsets
       - Option(Alt)を押しながらドラッグすると移動量が1/10 / Option(Alt)-drag moves at one tenth speed

       使い方 / Usage:
         var viewControl = ViewControl.create(doc);
         viewControl.buildUI(group, { labelWidth: 58, sliderWidth: 200, zoomLabel: "ズーム：", panXLabel: "左右：", panYLabel: "上下：" });
         viewControl.restore(); // 元の表示に戻す / back to the original view */

    var ViewControl = (function () {

        var ZOOM_MIN_PERCENT = 10;     /* スライダーの最小倍率（%） / slider minimum */
        var ZOOM_MAX_PERCENT = 1600;   /* スライダーの最大倍率（%） / slider maximum */

        /**
         * 値を下限・上限の範囲に収めます。
         *
         * @param {number} value - 対象の値。
         * @param {number} minValue - 下限。
         * @param {number} maxValue - 上限。
         * @returns {number} 範囲に収めた値。
         */
        function clamp(value, minValue, maxValue) {
            return Math.min(Math.max(value, minValue), maxValue);
        }

        /**
         * Option(Alt)キーが押されているかを返します。
         *
         * @returns {boolean} 押されている場合は true。
         */
        function isAltKeyDown() {
            var keyboard = ScriptUI.environment.keyboardState;
            return !!(keyboard && keyboard.altKey);
        }

        /**
         * スライダーの値を処理に渡します。Option(Alt)を押しながらのドラッグは移動量を1/10にします。
         *
         * @param {Slider} slider - 対象のスライダー。
         * @param {Object} sliderState - {rawValue, effectiveValue}。呼び出し側がスライダーごとに保持します。
         * @param {Function} applyValue - 補正後の値を受け取る処理。
         * @returns {void}
         */
        function applySliderValue(slider, sliderState, applyValue) {
            var rawValue = Number(slider.value);
            if (isNaN(rawValue)) rawValue = 0;

            var effectiveValue = rawValue;
            if (isAltKeyDown() && sliderState.rawValue != null) {
                effectiveValue = sliderState.effectiveValue + (rawValue - sliderState.rawValue) * 0.1;
                slider.value = effectiveValue;
            }

            sliderState.rawValue = rawValue;
            sliderState.effectiveValue = effectiveValue;
            applyValue(effectiveValue);
        }

        /**
         * アクティブなアートボードの矩形を返します。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {number[]} [左, 上, 右, 下] の座標。
         */
        function getActiveArtboardRect(doc) {
            return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        }

        /**
         * パンの可動範囲を返します（アートボードの半分を 100〜50000pt に収めた値）。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {Object} {xMax, yMax}（pt）。
         */
        function getPanRange(doc) {
            var rect = getActiveArtboardRect(doc);
            return {
                xMax: clamp(Math.round(Math.abs(rect[2] - rect[0]) / 2), 100, 50000),
                yMax: clamp(Math.round(Math.abs(rect[1] - rect[3]) / 2), 100, 50000)
            };
        }

        /**
         * ドキュメントの表示を操作する ViewControl を作ります。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {Object} {zoomBy, restore, buildUI}
         */
        function create(doc) {
            var view = doc.views[0];
            var originalZoom = view.zoom;
            var originalCenter = view.centerPoint;

            var panX = 0; /* pt */
            var panY = 0; /* pt（UIでは下が正 / positive is down in the UI） */
            var zoomSlider = null;

            /**
             * 表示倍率と表示中心を反映します（中心はアートボード中心＋パン量）。
             *
             * @param {number|null} zoomFactor - 表示倍率（1＝100%）。null なら倍率は変えません。
             * @returns {void}
             */
            function applyView(zoomFactor) {
                try {
                    if (zoomFactor != null) view.zoom = clamp(zoomFactor, 0.0313, 640.0);

                    /* Illustratorは上が＋Y、UIは下が正なので引く / Illustrator's +Y is up, so subtract */
                    var rect = getActiveArtboardRect(doc);
                    view.centerPoint = [(rect[0] + rect[2]) / 2 + panX, (rect[1] + rect[3]) / 2 - panY];
                    app.redraw();

                    /* ズームスライダーを現在の表示倍率に追従させる / Keep the zoom slider in step */
                    if (zoomFactor != null && zoomSlider) zoomSlider.value = Math.round(view.zoom * 100);
                } catch (e) { }
            }

            /**
             * パン量を可動範囲に収めて反映します。
             *
             * @param {string} axis - "x" または "y"。
             * @param {number} value - パン量（pt）。
             * @returns {void}
             */
            function setPan(axis, value) {
                var panRange = getPanRange(doc);
                var panPt = Math.round(Number(value) || 0);

                if (axis === "x") panX = clamp(panPt, -panRange.xMax, panRange.xMax);
                else panY = clamp(panPt, -panRange.yMax, panRange.yMax);
                applyView(null);
            }

            /**
             * ラベル＋スライダーの1行を追加します。
             *
             * @param {Group} parent - 追加先のグループ。
             * @param {Object} options - buildUI() に渡された設定。
             * @param {string} labelText - 項目名。
             * @param {number} value - 初期値。
             * @param {number} minValue - 最小値。
             * @param {number} maxValue - 最大値。
             * @param {Function} applyValue - 補正後の値を受け取る処理。
             * @returns {Slider} 追加したスライダー。
             */
            function addSliderRow(parent, options, labelText, value, minValue, maxValue, applyValue) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = ["left", "center"];

                var label = row.add("statictext", undefined, labelText);
                label.preferredSize.width = options.labelWidth;

                var slider = row.add("slider", undefined, value, minValue, maxValue);
                slider.preferredSize.width = options.sliderWidth;
                if (options.sliderHelpTip) slider.helpTip = options.sliderHelpTip;

                var sliderState = { rawValue: null, effectiveValue: null };
                slider.onChanging = function () {
                    applySliderValue(slider, sliderState, applyValue);
                };
                return slider;
            }

            return {
                /**
                 * 現在の表示倍率に指定の倍率を掛けます（0.9 なら10%縮小）。
                 *
                 * @param {number} factor - 掛ける倍率。
                 * @returns {void}
                 */
                zoomBy: function (factor) {
                    applyView(view.zoom * factor);
                },

                /**
                 * ダイアログを開く前の表示倍率と表示中心に戻します。
                 *
                 * @returns {void}
                 */
                restore: function () {
                    try {
                        view.zoom = originalZoom;
                        view.centerPoint = originalCenter;
                        app.redraw();
                    } catch (e) { }
                    panX = 0;
                    panY = 0;
                },

                /**
                 * ズーム・左右・上下のスライダーを追加します。
                 *
                 * @param {Group} parent - 追加先のグループ。
                 * @param {Object} options - labelWidth / sliderWidth / zoomLabel / panXLabel / panYLabel / sliderHelpTip。
                 * @returns {void}
                 */
                buildUI: function (parent, options) {
                    var initialZoomPercent = Math.round(originalZoom * 100);
                    if (!(initialZoomPercent >= ZOOM_MIN_PERCENT)) initialZoomPercent = 100;

                    var panRange = getPanRange(doc);

                    zoomSlider = addSliderRow(parent, options, options.zoomLabel, initialZoomPercent, ZOOM_MIN_PERCENT, ZOOM_MAX_PERCENT, function (value) {
                        applyView(clamp(Math.round(value), ZOOM_MIN_PERCENT, ZOOM_MAX_PERCENT) / 100);
                    });
                    addSliderRow(parent, options, options.panXLabel, 0, -panRange.xMax, panRange.xMax, function (value) {
                        setPan("x", value);
                    });
                    addSliderRow(parent, options, options.panYLabel, 0, -panRange.yMax, panRange.yMax, function (value) {
                        setPan("y", value);
                    });
                }
            };
        }

        return { create: create };
    })();

    // =========================================
    // 生成物のタグ / Tags of generated items
    // =========================================
    /* 生成したオブジェクトは name と note の両方にタグを持たせ、後処理で見分けます
       Generated items carry the same tag in name and note so later passes can find them */

    var TAG_OUTER_EDGE = "__OuterEdge__";          /* 外枠の4辺 / the four outer edges */
    var TAG_OUTER_ROUND = "__OuterRoundPreview__"; /* 外枠の角丸プレビュー / outer round preview */
    var TAG_TITLE_FILL = "__TitleFill__";          /* タイトル帯の塗り / title band fill */
    var TAG_TITLE_DIVIDER = "__TitleDivider__";    /* タイトル帯の分割線 / title band divider */
    var TAG_INNER_FILL = "__InnerBoxFill__";       /* 内側エリアの塗り / inner area fill */
    var TAG_FRAME_FILL = "__FrameFill__";          /* フレーム / frame */

    /**
     * 生成したオブジェクトにタグを付けます（name と note の両方）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @param {string} tag - 付与するタグ文字列。
     * @returns {void}
     */
    function tagItem(item, tag) {
        item.name = tag;
        item.note = tag;
    }

    /**
     * 内部タグをレイヤーパネルから隠します（name をクリアし、判定用の note は残します）。
     *
     * @param {PageItem[]} items - 対象のオブジェクト。
     * @returns {void}
     */
    function clearTagNames(items) {
        for (var i = 0; i < items.length; i++) {
            /* パスファインダーでグループが置き換わり、参照が無効になっていることがある
               A pathfinder result may have replaced a group, leaving a stale reference */
            try { items[i].name = ""; } catch (e) { }
        }
    }

    // =========================================
    // オブジェクトの操作 / Item helpers
    // =========================================

    /**
     * K版のみのグレー（CMYK）を作成します。
     *
     * @param {number} blackPercent - K版の濃度（0〜100）。
     * @returns {CMYKColor} 生成したカラー。
     */
    function makeGrayColor(blackPercent) {
        var color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 0;
        color.yellow = 0;
        color.black = blackPercent;
        return color;
    }

    /**
     * 線なし・グレーの塗りにします。
     *
     * @param {PathItem} item - 対象のパス。
     * @param {number} blackPercent - K版の濃度（0〜100）。
     * @returns {void}
     */
    function setGrayFill(item, blackPercent) {
        item.stroked = false;
        item.filled = true;
        item.fillColor = makeGrayColor(blackPercent);
    }

    /**
     * 塗りなし・指定の線にします。
     *
     * @param {PathItem} item - 対象のパス。
     * @param {Color} color - 線の色。
     * @param {number} width - 線幅（pt）。
     * @returns {void}
     */
    function setStroke(item, color, width) {
        item.filled = false;
        item.stroked = true;
        item.strokeColor = color;
        item.strokeWidth = width;
    }

    /**
     * 線と塗りの見た目をコピーします（角丸プレビュー用の複製に使用）。
     *
     * @param {PathItem} source - コピー元のパス。
     * @param {PathItem} target - コピー先のパス。
     * @returns {void}
     */
    function copyAppearance(source, target) {
        target.stroked = source.stroked;
        target.filled = source.filled;
        target.strokeColor = source.strokeColor;
        target.fillColor = source.fillColor;
        target.strokeWidth = source.strokeWidth;
    }

    /**
     * オブジェクトを最背面へ送ります（環境差で落ちることがあるため保護）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {void}
     */
    function sendToBack(item) {
        try { item.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
    }

    /**
     * オブジェクトを削除します（すでに消えている場合は何もしません）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {void}
     */
    function removeItem(item) {
        try { item.remove(); } catch (e) { }
    }

    /**
     * オブジェクトの配列をまとめて削除します。
     *
     * @param {PageItem[]} items - 対象のオブジェクト。
     * @returns {void}
     */
    function removeItems(items) {
        for (var i = 0; i < items.length; i++) {
            removeItem(items[i]);
        }
    }

    /**
     * 角丸のライブエフェクトを適用します。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @param {number} radiusPt - 角丸の半径（pt）。0以下なら何もしません。
     * @returns {void}
     */
    function applyRoundCornersEffect(item, radiusPt) {
        if (!(radiusPt > 0)) return;
        var effectXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        try { item.applyEffect(effectXml.replace('#value#', radiusPt)); } catch (e) { }
    }

    (function () {
        // =========================================
        // 準備 / Setup
        // =========================================
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        var rulerUnit = getUnitInfo();

        /* 一部環境で StrokeCap が未定義になるため、最低限の定数を用意
           Provide the StrokeCap constants for hosts that do not expose them */
        if (typeof StrokeCap === "undefined") {
            StrokeCap = {
                BUTTENDCAP: 0,
                ROUNDENDCAP: 1,
                PROJECTINGENDCAP: 2
            };
        }

        /* 選択パスの元の見た目（キャンセル時に戻すため）
           The original look of the selected paths, restored on cancel */
        var originalAppearances = [];

        /* 基準の長方形（選択中のパス、またはアートボード基準の一時矩形）
           Base rectangles: the selected paths, or the temporary artboard rectangle */
        var baseRects = collectSelectedPaths();

        /* パスが選択されていなければ、アクティブなアートボードを基準にする
           With no path selected, the active artboard becomes the base */
        var isArtboardBased = (baseRects.length === 0);
        var artboardBaseRect = null;

        /* 生成したオブジェクト（プレビューと実行の両方） / Items generated for the preview or the final run */
        var generatedItems = [];

        /* 自動ONの判定に使う、直前の状態と手動操作のフラグ
           Previous states and manual-input flags used by the auto-on rules */
        var titleHadSize = false;           /* タイトルエリアにサイズがあったか / the title area had a size */
        var gridWasSplittable = false;      /* 列・行が分割可能だったか / the grid could carry dividers */
        var innerFillManuallySet = false;   /* 内側エリアの［塗り］を操作したか / inner fill was set by hand */
        var bleedManuallySet = false;       /* フレームの［裁ち落とし］を操作したか / bleed was set by hand */

        /* ∧∨付きの入力欄（有効／無効を∧∨にそろえるため控える） / Fields with steppers, kept to sync their enabled state */
        var stepperInputs = [];

        /* ダイアログのコントロール / Dialog controls */
        var viewControl = ViewControl.create(doc);
        var marginFields, innerOffsetFields;
        var frameCheckbox, frameWidthGroup, frameWidthInput, bleedCheckbox, frameRoundCheckbox, frameRoundInput;
        var keepOuterCheckbox, outerRoundCheckbox, outerRoundInput;
        var outerEdgeScaleCheckbox, outerEdgeScaleValueGroup, outerEdgeScaleInput;
        var strokeCapPanel, capButtRadio, capRoundRadio, capProjectRadio;
        var titleCheckbox, titleSizeInput;
        var titlePositionGroup, titleTopRadio, titleBottomRadio, titleLeftRadio, titleRightRadio;
        var titleOptionGroup, titleFillCheckbox, titleLineCheckbox;
        var titleEdgeScaleRow, titleEdgeScaleCheckbox, titleEdgeScaleInput;
        var columnCountInput, columnGutterInput, rowCountInput, rowGutterInput;
        var innerFillCheckbox, innerDividerCheckbox;
        var lineTypePanel, lineSolidRadio, lineDashRadio, lineDotsRadio;
        var previewCheckbox;

        // =========================================
        // 基準の長方形 / Base rectangles
        // =========================================

        /**
         * 選択からパスアイテムだけを取り出し、見た目を塗りなし・K100の1pt線にそろえます。
         *
         * 元の見た目は originalAppearances に控え、キャンセル時に戻せるようにします。
         * テキスト編集中は doc.selection が TextRange になるため、選択なしとして扱います。
         *
         * @returns {PathItem[]} 選択中のパスアイテム。
         */
        function collectSelectedPaths() {
            var selectedPaths = [];
            var selectedItems = doc.selection;

            /* 配列でなければ（テキスト編集中の TextRange など）選択なし扱い
               Anything but an array of items, such as a TextRange, counts as no selection */
            if (!selectedItems || !(selectedItems instanceof Array)) return selectedPaths;

            /* グループの中のパスも拾う。見た目を書き換えるので、マスク・ガイド・ロック・非表示は除く
               Also pick up paths inside groups; skip masks, guides, locked and hidden items since their look gets changed */
            var candidatePaths = collectSelectionPathItems(selectedItems, {
                compoundPaths: "skip",
                skipClipMasks: true,
                skipGuides: true,
                skipLocked: true,
                skipHidden: true
            });

            for (var i = 0; i < candidatePaths.length; i++) {
                var item = candidatePaths[i];

                originalAppearances.push({
                    item: item,
                    filled: item.filled,
                    fillColor: item.fillColor,
                    stroked: item.stroked,
                    strokeColor: item.strokeColor,
                    strokeWidth: item.strokeWidth
                });

                setStroke(item, makeGrayColor(GRAY_TINTS.rule), 1);
                selectedPaths.push(item);
            }
            return selectedPaths;
        }

        /**
         * 選択パスの見た目を、スクリプト実行前の状態に戻します（キャンセル時に使用）。
         *
         * @returns {void}
         */
        function restoreSelectedAppearances() {
            for (var i = 0; i < originalAppearances.length; i++) {
                var appearance = originalAppearances[i];
                try {
                    appearance.item.filled = appearance.filled;
                    appearance.item.fillColor = appearance.fillColor;
                    appearance.item.stroked = appearance.stroked;
                    appearance.item.strokeColor = appearance.strokeColor;
                    appearance.item.strokeWidth = appearance.strokeWidth;
                } catch (e) { }
            }
        }

        /**
         * アクティブなアートボードの矩形を返します。
         *
         * @returns {number[]} [左, 上, 右, 下] の座標。
         */
        function getActiveArtboardRect() {
            return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        }

        /**
         * アートボード基準の一時矩形を削除します（baseRects からも外します）。
         *
         * 削除後に参照へ触ると Error 45 になるため、参照を先に外します。
         *
         * @returns {void}
         */
        function removeArtboardBaseRect() {
            if (!artboardBaseRect) return;

            for (var i = baseRects.length - 1; i >= 0; i--) {
                if (baseRects[i] === artboardBaseRect) baseRects.splice(i, 1);
            }

            removeItem(artboardBaseRect);
            artboardBaseRect = null;
        }

        /**
         * アートボード基準の一時矩形を、マージンの分だけ内側に作り直します。
         *
         * マージンが大きすぎて領域が残らない場合は、マージンを 0 とみなします。
         * 裁ち落としはここでは適用しません（フレームだけに適用します）。
         *
         * @param {Object} marginPt - {top, right, bottom, left}（pt、0以上）。
         * @returns {void}
         */
        function rebuildArtboardBaseRect(marginPt) {
            var artboardRect = getActiveArtboardRect(); // [L, T, R, B]
            var left = artboardRect[0] + marginPt.left;
            var top = artboardRect[1] - marginPt.top;
            var width = (artboardRect[2] - marginPt.right) - left;
            var height = top - (artboardRect[3] + marginPt.bottom);

            if (!(width > 0) || !(height > 0)) {
                left = artboardRect[0];
                top = artboardRect[1];
                width = artboardRect[2] - artboardRect[0];
                height = artboardRect[1] - artboardRect[3];
            }

            removeArtboardBaseRect();

            /* ロックされたレイヤーなどで作成できなくても、ダイアログは続行する
               Keep the dialog running even if the rectangle cannot be created */
            try {
                artboardBaseRect = doc.activeLayer.pathItems.rectangle(top, left, width, height);
                setStroke(artboardBaseRect, makeGrayColor(GRAY_TINTS.rule), 1);
                baseRects.push(artboardBaseRect);
            } catch (e) { }
        }

        /**
         * 基準の長方形をすべて削除します（アートボード基準の一時矩形も含む）。
         *
         * @returns {void}
         */
        function removeBaseRects() {
            removeItems(baseRects);
            baseRects = [];
            artboardBaseRect = null;
        }

        /**
         * 基準の長方形の表示／非表示をまとめて切り替えます。
         *
         * @param {boolean} hidden - 非表示にする場合は true。
         * @returns {void}
         */
        function setBaseRectsHidden(hidden) {
            for (var i = 0; i < baseRects.length; i++) {
                baseRects[i].hidden = hidden;
            }
        }

        // =========================================
        // 生成物の管理 / Generated item tracking
        // =========================================

        /**
         * 生成したオブジェクトを追跡リストに登録して返します。
         *
         * 生成した直後に登録するのが重要です。見た目の設定などで例外が起きても、
         * すでに登録済みならプレビュー解除時に必ず削除できます
         * （登録前に中断すると、消せないオブジェクトがドキュメントに残ります）。
         *
         * @param {PageItem} item - 登録するオブジェクト。
         * @returns {PageItem} 受け取ったオブジェクトをそのまま返します。
         */
        function trackGeneratedItem(item) {
            generatedItems.push(item);
            return item;
        }

        /**
         * 生成したオブジェクトをすべて削除します。
         *
         * @returns {void}
         */
        function removeGeneratedItems() {
            removeItems(generatedItems);
            generatedItems = [];
        }

        // =========================================
        // 入力値の変換 / Input conversion
        // =========================================

        /**
         * 入力欄の文字列を、現在の単位からpt値に換算します。
         *
         * @param {string} text - 入力欄の文字列。
         * @returns {number} pt値。数値として読めない場合は 0。
         */
        function toPt(text) {
            var value = parseFloat(text);
            return isNaN(value) ? 0 : (value * rulerUnit.pointsPerUnit);
        }

        /**
         * 入力欄の文字列を、0以上のpt値に換算します（マージンなど負を許さない項目用）。
         *
         * @param {string} text - 入力欄の文字列。
         * @returns {number} 0以上のpt値。
         */
        function toPositivePt(text) {
            var valuePt = toPt(text);
            return (valuePt > 0) ? valuePt : 0;
        }

        /**
         * 入力欄の文字列を、1以上・上限以下の整数（列数・行数）に変換します。
         *
         * 上限を設けないと、入力のたびに膨大な数のセルを生成してIllustratorが止まります。
         *
         * @param {string} text - 入力欄の文字列。
         * @returns {number} 1〜maxGridCount の整数。
         */
        function toCount(text) {
            var value = parseInt(text, 10);
            if (isNaN(value) || value < 1) return 1;
            return Math.min(value, GENERATION_SETTINGS.maxGridCount);
        }

        /**
         * 列数・行数の入力欄を、1〜上限の範囲に補正します（空欄のあいだは触りません）。
         *
         * @param {EditText} input - 対象の入力欄。
         * @returns {void}
         */
        function snapCountInput(input) {
            if (input.text === "") return;

            var countText = String(toCount(input.text));
            if (countText !== input.text) input.text = countText;
        }

        /**
         * mmで持っている既定値を、現在の定規単位の文字列にします。
         *
         * @param {number} valueMm - mmでの既定値。
         * @returns {string} 現在の定規単位での既定値（小数第1位まで）。
         */
        function defaultValueText(valueMm) {
            return String(Math.round(mmToPt(valueMm) / rulerUnit.pointsPerUnit * 10) / 10);
        }

        /**
         * チェックがONのときだけ、入力欄の値を0以上のpt値で返します。
         *
         * @param {Checkbox} checkbox - 有効／無効を決めるチェックボックス。
         * @param {EditText} input - 値の入力欄。
         * @returns {number} pt値。チェックがOFFなら 0。
         */
        function readCheckedPt(checkbox, input) {
            return checkbox.value ? toPositivePt(input.text) : 0;
        }

        /**
         * 入力欄に正の数値が入っているかを返します。
         *
         * @param {EditText} input - 対象の入力欄。
         * @returns {boolean} 正の数値なら true。
         */
        function hasPositiveValue(input) {
            return (parseFloat(input.text) > 0);
        }

        /**
         * 入力欄が空欄または0なら、既定値を入れます（チェックをONにしたときに使用）。
         *
         * @param {EditText} input - 対象の入力欄。
         * @param {string} defaultText - 入れる既定値。
         * @returns {void}
         */
        function fillDefaultIfZero(input, defaultText) {
            var value = parseFloat(input.text);
            if (isNaN(value) || value === 0) input.text = defaultText;
        }

        /**
         * 内側エリアのオフセット初期値を求めます（外形の (幅+高さ)/40 を10単位に丸めた値）。
         *
         * @returns {number} 現在の定規単位でのオフセット初期値。
         */
        function calcDefaultInnerOffset() {
            if (baseRects.length === 0) return 0;

            var bounds = baseRects[0].geometricBounds; // [L, T, R, B]
            var widthPt = bounds[2] - bounds[0];
            var heightPt = bounds[1] - bounds[3];
            if (!(widthPt > 0) || !(heightPt > 0)) return 0;

            var offsetValue = (widthPt + heightPt) / GENERATION_SETTINGS.innerOffsetDivisor / rulerUnit.pointsPerUnit;

            /* 10以上なら10単位、それ未満は0.1単位に丸める
               （cmやinchのように値が小さくなる単位で0にならないようにする）
               Round to the nearest ten, or to one decimal for units that give small numbers */
            if (offsetValue >= 10) return Math.round(offsetValue / 10) * 10;
            return Math.round(offsetValue * 10) / 10;
        }

        /**
         * タイトルエリアのサイズ初期値を求めます（帯が伸びる方向の長さ / titleSizeDivisor）。
         *
         * 上下に置くときは高さ、左右に置くときは幅を基準にします
         * （高さで決めてしまうと、縦長の長方形で幅を超えてタイトルが作れなくなります）。
         *
         * @returns {number} 現在の定規単位でのサイズ初期値。
         */
        function calcDefaultTitleSize() {
            if (baseRects.length === 0) return 10;

            var bounds = baseRects[0].geometricBounds; // [L, T, R, B]
            var positionKey = getTitlePositionKey();
            var basePt = (positionKey === "top" || positionKey === "bottom")
                ? Math.abs(bounds[1] - bounds[3])
                : Math.abs(bounds[2] - bounds[0]);

            var titleSize = basePt / rulerUnit.pointsPerUnit / GENERATION_SETTINGS.titleSizeDivisor;
            if (!(titleSize > 0)) return 10;

            return Math.max(Math.round(titleSize * 10) / 10, 0.1);
        }

        // =========================================
        // UIの部品 / UI building blocks
        // =========================================

        /**
         * タブを追加し、その中身を並べる縦組みのグループを返します。
         *
         * @param {TabbedPanel} tabbedPanel - 追加先のタブパネル。
         * @param {Object} labelSet - タブ名のラベル。
         * @returns {Group} 中身を追加するグループ。
         */
        function addTabColumn(tabbedPanel, labelSet) {
            var tab = tabbedPanel.add("tab", undefined, getLabel(labelSet));
            tab.orientation = "column";
            tab.alignChildren = ["fill", "top"];
            tab.spacing = 8;
            tab.margins = 10;

            var column = tab.add("group");
            column.orientation = "column";
            column.alignChildren = ["fill", "top"];
            column.spacing = 8;
            return column;
        }

        /**
         * パネル名の後ろに付ける単位を、言語別の括弧で返します（日本語は全角）。
         *
         * @returns {string} 例）"（mm）" / " (mm)"
         */
        function unitSuffix() {
            return (uiLang === "ja") ? ("（" + rulerUnit.label + "）") : (" (" + rulerUnit.label + ")");
        }

        /**
         * 入力欄にツールチップを設定します（↑↓キーの説明を必ず添えます）。
         *
         * @param {EditText} input - 対象の入力欄。
         * @param {Object} [labelSet] - 入力欄ごとの説明。省略時は↑↓キーの説明だけ。
         * @returns {void}
         */
        function setInputHelpTip(input, labelSet) {
            input.helpTip = (labelSet ? getLabel(labelSet) + "\n" : "") + getLabel(LABELS.tooltip.numberInput);
        }

        /**
         * 縦組みのパネルを追加します。
         *
         * @param {Group|Panel} parent - 追加先。
         * @param {string} title - パネル名。
         * @param {number} [spacing] - パネル内の要素間隔。省略時は既定値。
         * @returns {Panel} 追加したパネル。
         */
        function addPanel(parent, title, spacing) {
            var panel = parent.add("panel", undefined, title);
            panel.orientation = "column";
            panel.alignChildren = ["fill", "top"];
            panel.margins = DIALOG_LAYOUT.panelMargins;
            if (spacing) panel.spacing = spacing;
            return panel;
        }

        /**
         * ラジオボタンを横に並べるパネルを追加します。
         *
         * @param {Group|Panel} parent - 追加先。
         * @param {string} title - パネル名。
         * @returns {Panel} 追加したパネル。
         */
        function addRadioPanel(parent, title) {
            var panel = parent.add("panel", undefined, title);
            panel.orientation = "row";
            panel.alignChildren = ["left", "center"];
            panel.margins = DIALOG_LAYOUT.panelMargins;
            return panel;
        }

        /**
         * 左揃えの1行グループを追加します。
         *
         * @param {Group|Panel} parent - 追加先。
         * @param {number} [spacing] - 要素間隔。省略時は既定値。
         * @returns {Group} 追加したグループ。
         */
        function addRow(parent, spacing) {
            var row = parent.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            if (spacing != null) row.spacing = spacing;
            return row;
        }

        /**
         * パネルを非表示にし、レイアウト上の高さも潰します（長方形スタート時に使用）。
         *
         * @param {Panel} panel - 対象のパネル。
         * @param {boolean} collapsed - 畳む場合は true。
         * @returns {void}
         */
        function setPanelCollapsed(panel, collapsed) {
            panel.visible = !collapsed;
            panel.minimumSize.height = 0;
            panel.maximumSize.height = collapsed ? 0 : 10000;
        }

        /**
         * 複数のコントロールに同じイベントハンドラーを割り当てます。
         *
         * @param {Object[]} controls - 対象のコントロール。
         * @param {string} eventName - "onClick" などのハンドラー名。
         * @param {Function} handler - 割り当てる関数。
         * @returns {void}
         */
        function bindAll(controls, eventName, handler) {
            for (var i = 0; i < controls.length; i++) {
                controls[i][eventName] = handler;
            }
        }

        /**
         * 「ラベルなしチェックボックス＋項目名：入力欄 単位」の1行を作ります
         * （タイトルエリアとフレームの有効／無効で共用）。
         *
         * @param {Panel|Group} parent - 追加先。
         * @param {Object} labelSet - 項目名のラベル。
         * @param {Object} [inputTooltipSet] - 入力欄のツールチップ。
         * @returns {Object} {checkbox, valueGroup, input}
         */
        function addToggleValueRow(parent, labelSet, inputTooltipSet) {
            var row = addRow(parent, 0);

            var checkbox = row.add("checkbox", undefined, "");
            /* ラベルがないので、空文字ぶんの幅が入らないよう抑える
               No label, so cap the width to avoid the phantom text space */
            checkbox.preferredSize.width = DIALOG_LAYOUT.toggleCheckboxWidth;

            var valueGroup = addRow(row, 4);
            valueGroup.margins = 0;
            valueGroup.add("statictext", undefined, labelText(labelSet));

            var input = addStepperInput(valueGroup, "0", 4, { step: 1, min: 0 });
            valueGroup.add("statictext", undefined, rulerUnit.label);
            setInputHelpTip(input, inputTooltipSet);

            return { checkbox: checkbox, valueGroup: valueGroup, input: input };
        }

        /**
         * 「チェックボックス＋入力欄＋単位」の1行を作ります（角丸・辺の伸縮で共用）。
         *
         * @param {Panel|Group} parent - 追加先。
         * @param {Object} labelSet - チェックボックスのラベル。
         * @param {string} initialText - 入力欄の初期値。
         * @param {boolean} allowNegative - 負の値を許す場合は true。
         * @returns {Object} {row, checkbox, valueGroup, input}
         */
        function addCheckboxValueRow(parent, labelSet, initialText, allowNegative) {
            var row = addRow(parent);
            var checkbox = row.add("checkbox", undefined, getLabel(labelSet));

            var valueGroup = addRow(row);
            var input = addStepperInput(valueGroup, initialText, 4, allowNegative ? { step: 1 } : { step: 1, min: 0 });
            valueGroup.add("statictext", undefined, rulerUnit.label);
            setInputHelpTip(input);

            return { row: row, checkbox: checkbox, valueGroup: valueGroup, input: input };
        }

        /**
         * 上／左＋連動＋右／下 の3段組の入力欄を作ります（マージンと内側エリアのオフセットで共用）。
         *
         * ［連動］がONのあいだは、上の値を他の3つへコピーし、3つをディム表示にします。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @param {string} initialText - 入力欄の初期値。
         * @param {number} characters - 入力欄の文字数。
         * @param {string} unitLabel - 上下の入力欄に付ける単位。空文字なら付けません。
         * @param {Object} [inputTooltipSet] - 入力欄のツールチップ。
         * @returns {Object} {top, bottom, left, right, linkToggle, applyLinkState}
         */
        function buildLinkedQuadUI(parent, initialText, characters, unitLabel, inputTooltipSet) {
            var inputs = {};
            var fieldGroups = {};
            var followerKeys = ["bottom", "left", "right"];
            var isSyncing = false;

            /**
             * 中央寄せの1段を追加します。
             *
             * @param {number} [spacing] - 要素間隔。
             * @returns {Group} 追加したグループ。
             */
            function addCenteredRow(spacing) {
                var row = parent.add("group");
                row.orientation = "row";
                row.alignChildren = ["center", "center"];
                row.alignment = ["fill", "top"];
                if (spacing) row.spacing = spacing;
                return row;
            }

            /**
             * 「項目名：入力欄（＋単位）」のひとまとまりを追加します。
             *
             * @param {Group} row - 追加先の行。
             * @param {string} positionKey - "top" / "bottom" / "left" / "right"。
             * @param {boolean} withUnit - 単位を付ける場合は true。
             * @returns {void}
             */
            function addField(row, positionKey, withUnit) {
                var fieldGroup = addRow(row);
                fieldGroup.add("statictext", undefined, labelText(LABELS.fieldLabel[positionKey]));

                var input = addStepperInput(fieldGroup, initialText, characters, { step: 1, min: 0 });
                setInputHelpTip(input, inputTooltipSet);

                if (withUnit && unitLabel) fieldGroup.add("statictext", undefined, unitLabel);

                inputs[positionKey] = input;
                fieldGroups[positionKey] = fieldGroup;
            }

            /**
             * ［連動］の状態を反映します（連動ONなら上の値を他へコピーし、3つをディム表示）。
             *
             * @returns {void}
             */
            function applyLinkState() {
                var linked = linkToggle.value;

                /* 値の書き込みで onChanging が走らないようにフラグで抑える
                   The flag keeps our own writes from triggering onChanging */
                isSyncing = true;
                for (var i = 0; i < followerKeys.length; i++) {
                    fieldGroups[followerKeys[i]].enabled = !linked;
                    if (linked) inputs[followerKeys[i]].text = inputs.top.text;
                }
                isSyncing = false;
                syncStepperStates();
            }

            // 1段目：上（中央寄せ）
            addField(addCenteredRow(), "top", true);

            // 2段目：左 ＋ 連動（中央）＋ 右
            var middleRow = addCenteredRow(12);
            addField(middleRow, "left", false);
            var linkToggle = addLinkToggle(middleRow, true, function () {
                applyLinkState();
                requestPreview();
            });
            linkToggle.helpTip = getLabel(LABELS.tooltip.link);
            /* セッションの復元で setLinkToggleValue() を通すための目印 / Marks it so the session restore goes through setLinkToggleValue() */
            linkToggle.isLinkToggle = true;
            addField(middleRow, "right", false);

            // 3段目：下（中央寄せ）
            addField(addCenteredRow(), "bottom", true);

            inputs.top.onChanging = function () {
                if (isSyncing) return;
                if (linkToggle.value) applyLinkState();
                requestPreview();
            };

            bindAll([inputs.bottom, inputs.left, inputs.right], "onChanging", function () {
                if (isSyncing) return;
                requestPreview();
            });

            applyLinkState();

            return {
                top: inputs.top,
                bottom: inputs.bottom,
                left: inputs.left,
                right: inputs.right,
                linkToggle: linkToggle,
                applyLinkState: applyLinkState
            };
        }

        /**
         * 「◯数」＋「間隔」の1行パネルを作ります（列／行で共用）。
         *
         * @param {Group} parent - 追加先。
         * @param {Object} titleLabelSet - パネル名のラベル。
         * @param {Object} countLabelSet - 個数の項目名のラベル。
         * @returns {Object} {countInput, gutterInput}
         */
        function addGridCountPanel(parent, titleLabelSet, countLabelSet) {
            var panel = addPanel(parent, getLabel(titleLabelSet), 8);
            var row = addRow(panel);

            row.add("statictext", undefined, labelText(countLabelSet));
            /* 1〜上限の整数 / integers from 1 to the cap */
            var countInput = addStepperInput(row, "1", 3, { step: 1, min: 1, max: GENERATION_SETTINGS.maxGridCount, integer: true });
            setInputHelpTip(countInput);

            row.add("statictext", undefined, labelText(LABELS.fieldLabel.spacing));
            var gutterInput = addStepperInput(row, "0", 4, { step: 1, min: 0 });
            setInputHelpTip(gutterInput, LABELS.tooltip.spacing);
            row.add("statictext", undefined, rulerUnit.label);

            return { countInput: countInput, gutterInput: gutterInput };
        }

        /**
         * 表示コマンド用の小さめボタンを1つ追加します。
         *
         * @param {Group} parent - 追加先のグループ。
         * @param {Object} labelSet - ボタンのラベル。
         * @param {Function} action - クリック時に実行する処理。
         * @returns {void}
         */
        function addViewCommandButton(parent, labelSet, action) {
            var button = parent.add("button", undefined, getLabel(labelSet));
            button.alignment = "left";
            button.preferredSize = [DIALOG_LAYOUT.viewButtonWidth, DIALOG_LAYOUT.viewButtonHeight];
            button.minimumSize = [DIALOG_LAYOUT.viewButtonWidth, DIALOG_LAYOUT.viewButtonHeight];
            button.maximumSize = [DIALOG_LAYOUT.viewButtonWidth, DIALOG_LAYOUT.viewButtonHeight];
            button.onClick = action;
        }

        /**
         * ∧∨と入力欄をひと組で追加します（隙間0で突き合わせ、↑↓キーも∧∨と同じ処理で増減します）。
         * 増減したあとは入力欄の onChanging を呼び、プレビューや連動を手入力と同じように更新します。
         *
         * @param {Group} parent - 追加先の行。
         * @param {string} initialText - 入力欄の初期値。
         * @param {number} characters - 入力欄の文字数。
         * @param {Object} stepOptions - addStepper() に渡す増減の設定（min / max / integer）。
         * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できます）。
         */
        function addStepperInput(parent, initialText, characters, stepOptions) {
            var stepperInputGroup = parent.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;

            stepOptions.onStep = function (steppedInput) {
                if (typeof steppedInput.onChanging === "function") steppedInput.onChanging();
            };
            var input;
            var stepperGroup = addStepper(stepperInputGroup, function () { return input; }, stepOptions);
            input = stepperInputGroup.add("edittext", undefined, initialText);
            input.characters = characters;
            input.stepperGroup = stepperGroup;
            bindSteppedArrowKeys(input, stepperGroup);
            stepperInputs.push(input);
            return input;
        }

        /**
         * ∧∨の有効／無効を入力欄にそろえ、入力欄や親の状態が変わった∧∨だけ描き直します。
         * 入力欄・行・パネルの enabled を切り替えたあとに呼びます。
         *
         * @returns {void}
         */
        function syncStepperStates() {
            for (var i = 0; i < stepperInputs.length; i++) {
                var stepperGroup = stepperInputs[i].stepperGroup;
                stepperGroup.enabled = stepperInputs[i].enabled;
                var isEnabled = isStepperEnabledInTree(stepperGroup);
                if (stepperGroup.lastDrawnEnabled === isEnabled) continue;
                stepperGroup.lastDrawnEnabled = isEnabled;
                redrawSteppersIn(stepperGroup);
            }
        }

        // =========================================
        // コントロールの有効／無効 / Enabled states
        // =========================================

        /**
         * 辺の伸縮の値を返します（チェックがOFFのときは 0 とみなします）。
         *
         * @returns {number} 現在の単位での辺の伸縮量。
         */
        function getOuterEdgeScaleValue() {
            if (!outerEdgeScaleCheckbox.value) return 0;
            var value = parseFloat(outerEdgeScaleInput.text);
            return isNaN(value) ? 0 : value;
        }

        /**
         * 線端パネルの有効／無効を反映します。
         *
         * 線端は「4辺に分解するとき」だけ意味を持ちます（＝外枠を残す＋辺の伸縮≠0）。
         *
         * @returns {void}
         */
        function applyStrokeCapPanelEnabledState() {
            strokeCapPanel.enabled = (keepOuterCheckbox.value && getOuterEdgeScaleValue() !== 0);
        }

        /**
         * 外側エリアの［辺の伸縮］［角丸］［線端］の有効／無効を反映します。
         *
         * ［角丸］のチェック自体は、タイトルエリアの角丸も参照する値のため常に操作できます。
         *
         * @returns {void}
         */
        function applyOuterAreaEnabledState() {
            outerEdgeScaleCheckbox.enabled = keepOuterCheckbox.value;
            outerEdgeScaleValueGroup.enabled = (keepOuterCheckbox.value && outerEdgeScaleCheckbox.value);

            outerRoundInput.enabled = outerRoundCheckbox.value;
            if (!outerRoundCheckbox.value) outerRoundInput.text = "0";

            applyStrokeCapPanelEnabledState();
            syncStepperStates();
        }

        /**
         * タイトルエリアの［辺の伸縮］の入力欄の有効／無効を反映します。
         *
         * @returns {void}
         */
        function applyTitleEdgeScaleEnabledState() {
            titleEdgeScaleInput.enabled = titleEdgeScaleCheckbox.value;
            if (!titleEdgeScaleCheckbox.value) titleEdgeScaleInput.text = "0";
            syncStepperStates();
        }

        /**
         * タイトルエリアの各コントロールの有効／無効を反映します。
         *
         * 幅／高さはディムせず常に入力できます（チェックを外しても値は残します）。
         *
         * @returns {void}
         */
        function applyTitleAreaEnabledState() {
            var areaEnabled = titleCheckbox.value;
            titleOptionGroup.enabled = areaEnabled;

            /* 位置・塗り・仕切り線は「有効」かつ「サイズ>0」のときだけ
               The position, fill and divider need both the checkbox and a size */
            var usable = (areaEnabled && hasPositiveValue(titleSizeInput));
            titlePositionGroup.enabled = usable;
            titleFillCheckbox.enabled = usable;
            titleLineCheckbox.enabled = usable;

            if (!usable) {
                titleFillCheckbox.value = false;
                titleLineCheckbox.value = false;
            } else if (!titleHadSize) {
                /* 仕切り線：サイズが 0→>0 になった瞬間だけ自動ON（ユーザーは後からOFF可）
                   The divider is auto-enabled only on the zero-to-positive transition */
                titleLineCheckbox.value = true;
            }

            titleHadSize = usable;

            /* ［線の伸縮］は仕切り線がONのときだけ操作できる（自動ONの結果を見てから判定する）
               The divider scale follows the divider checkbox, after the auto-on rule above */
            titleEdgeScaleRow.enabled = (usable && titleLineCheckbox.value);
            if (!titleEdgeScaleRow.enabled) {
                titleEdgeScaleCheckbox.value = false;
                applyTitleEdgeScaleEnabledState();
            }
            syncStepperStates();
        }

        /**
         * フレーム関連UIの有効／無効を、基準の種類とフレーム幅に応じて反映します。
         *
         * フレームはアートボード基準のときだけ使えるため、長方形スタート時は値ごとリセットします。
         *
         * @returns {void}
         */
        function applyFrameEnabledState() {
            if (!isArtboardBased) {
                frameCheckbox.value = false;
                bleedCheckbox.value = false;
                frameRoundCheckbox.value = false;
                frameWidthInput.text = "0";
                frameRoundInput.text = "0";

                frameWidthGroup.enabled = false;
                bleedCheckbox.enabled = false;
                frameRoundCheckbox.enabled = false;
                frameRoundInput.enabled = false;
                syncStepperStates();
                return;
            }

            /* 幅はディムせず常に入力可。OFFのときはチェックを外すだけで値は残す
               Never dim the width; turning the checkbox off keeps the value as-is */
            frameWidthGroup.enabled = true;

            var usable = (frameCheckbox.value && hasPositiveValue(frameWidthInput));
            bleedCheckbox.enabled = usable;
            frameRoundCheckbox.enabled = usable;
            if (!frameCheckbox.value) {
                bleedCheckbox.value = false;
                frameRoundCheckbox.value = false;
            }
            frameRoundInput.enabled = (usable && frameRoundCheckbox.value);
            syncStepperStates();
        }

        /**
         * 列数・行数から分割線を引けるかを判定します。
         *
         * @param {number} columnCount - 列数。
         * @param {number} rowCount - 行数。
         * @returns {boolean} 分割線を引ける場合は true。
         */
        function isGridSplittable(columnCount, rowCount) {
            return (columnCount > 1 || rowCount > 1);
        }

        /**
         * ［分割線］と［線の種類］パネルの有効／無効を反映します。
         *
         * @param {number} columnCount - 列数。
         * @param {number} rowCount - 行数。
         * @param {boolean} allowAutoOn - 1/1から分割可能になった瞬間に自動ONしてよい場合は true。
         * @returns {void}
         */
        function applyInnerDividerEnabledState(columnCount, rowCount, allowAutoOn) {
            var splittable = isGridSplittable(columnCount, rowCount);
            innerDividerCheckbox.enabled = splittable;

            if (!splittable) {
                // 分割できないなら分割線は不要
                innerDividerCheckbox.value = false;
            } else if (allowAutoOn && !gridWasSplittable) {
                // 1/1 から分割可能になった瞬間だけ自動ON
                innerDividerCheckbox.value = true;
            }

            gridWasSplittable = splittable;
            lineTypePanel.enabled = (splittable && innerDividerCheckbox.value);
        }

        /**
         * 他の入力値に依存するコントロールを更新します（collectOptions の副作用をまとめたもの）。
         *
         * @param {Object} options - collectOptions() が組み立てた生成条件。
         * @returns {void}
         */
        function syncDependentControls(options) {
            /* ガターは列／行が2以上のときだけ入力可 / Gutters are editable for 2+ columns or rows only */
            columnGutterInput.enabled = (options.columnCount > 1);
            rowGutterInput.enabled = (options.rowCount > 1);

            /* ガターが入ったら塗りを自動ON（手動操作があれば尊重）
               Turn the fill on once a gutter is set, unless the user set it manually */
            if (!innerFillManuallySet && (options.columnGutterPt !== 0 || options.rowGutterPt !== 0)) {
                innerFillCheckbox.value = true;
            }

            applyInnerDividerEnabledState(options.columnCount, options.rowCount, false);
            applyStrokeCapPanelEnabledState();
            syncStepperStates();
        }

        // =========================================
        // パネルの組み立て / Panel builders
        // =========================================

        /**
         * ［マージン］パネルを組み立てます（アートボード基準のときだけ使います）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildMarginPanel(parent) {
            var marginPanel = addPanel(parent, getLabel(LABELS.panel.margin), 10);
            marginFields = buildLinkedQuadUI(marginPanel, defaultValueText(GENERATION_SETTINGS.defaultMarginMm), 4, rulerUnit.label);

            /* アートボード基準のときだけ表示・操作できる
               The panel is shown and enabled for artboard-based runs only */
            marginPanel.enabled = isArtboardBased;
            setPanelCollapsed(marginPanel, !isArtboardBased);
        }

        /**
         * ［フレーム］パネルを組み立てます（アートボード基準のときだけ使います）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildFramePanel(parent) {
            var framePanel = addPanel(parent, getLabel(LABELS.panel.frame), 10);
            setPanelCollapsed(framePanel, !isArtboardBased);

            var widthRow = addToggleValueRow(framePanel, LABELS.fieldLabel.width);
            frameCheckbox = widthRow.checkbox;
            frameCheckbox.helpTip = getLabel(LABELS.tooltip.frame);
            frameWidthGroup = widthRow.valueGroup;
            frameWidthInput = widthRow.input;

            /* 裁ち落とし（表示は常時。長方形スタート時は enabled で制御）
               The bleed row is always visible; rectangle-based runs disable it instead */
            bleedCheckbox = addRow(framePanel).add("checkbox", undefined, getLabel(LABELS.checkbox.bleed));
            bleedCheckbox.helpTip = getLabel(LABELS.tooltip.bleed);

            var roundRow = addCheckboxValueRow(framePanel, LABELS.checkbox.round, "0", false);
            frameRoundCheckbox = roundRow.checkbox;
            frameRoundCheckbox.helpTip = getLabel(LABELS.tooltip.frameRound);
            frameRoundInput = roundRow.input;

            frameCheckbox.onClick = function () {
                if (frameCheckbox.value) {
                    /* 幅が0なら既定値を入れ、裁ち落としも自動でON（手動操作があれば尊重）
                       Fill in the default width at zero and turn the bleed on, unless the user set it */
                    fillDefaultIfZero(frameWidthInput, defaultValueText(GENERATION_SETTINGS.defaultFrameWidthMm));
                    if (!bleedManuallySet) bleedCheckbox.value = true;
                }

                applyFrameEnabledState();
                requestPreview();
            };

            frameWidthInput.onChanging = function () {
                /* OFFのときは入力だけ受け付け、プレビューには反映しない
                   While unchecked the field just stores the value; nothing is previewed */
                if (!frameCheckbox.value) return;

                /* 幅が0→>0になったら裁ち落としを自動ON（手動操作があれば尊重）
                   Turn the bleed on once a width is set, unless the user set it manually */
                var hasWidth = hasPositiveValue(frameWidthInput);
                bleedCheckbox.enabled = hasWidth;
                if (!hasWidth) bleedCheckbox.value = false;
                else if (!bleedManuallySet) bleedCheckbox.value = true;

                frameRoundCheckbox.enabled = hasWidth;
                if (!hasWidth) {
                    frameRoundCheckbox.value = false;
                    frameRoundInput.enabled = false;
                    syncStepperStates();
                }
                requestPreview();
            };

            bleedCheckbox.onClick = function () {
                /* ユーザーが操作したら以後は自動ONしない / Stop auto-enabling once the user decides */
                bleedManuallySet = true;
                requestPreview();
            };

            frameRoundCheckbox.onClick = function () {
                frameRoundInput.enabled = frameRoundCheckbox.value;
                syncStepperStates();

                /* ONにしたとき、0なら既定値を入れる / Fill in the default radius when enabled at zero */
                if (frameRoundCheckbox.value) fillDefaultIfZero(frameRoundInput, defaultValueText(GENERATION_SETTINGS.defaultRoundMm));
                requestPreview();
            };

            frameRoundInput.onChanging = requestPreview;

            /* 初期反映（フレーム幅0なら裁ち落とし・角丸はディム）
               Initial state: a zero width dims both the bleed and the rounding */
            applyFrameEnabledState();
        }

        /**
         * ［外側エリア］パネルを組み立てます。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildOuterPanel(parent) {
            var outerPanel = addPanel(parent, getLabel(LABELS.panel.outer), 10);

            keepOuterCheckbox = outerPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.keepOuter));
            keepOuterCheckbox.value = true;
            keepOuterCheckbox.helpTip = getLabel(LABELS.tooltip.keepOuter);

            var roundRow = addCheckboxValueRow(outerPanel, LABELS.checkbox.round, "0", false);
            outerRoundCheckbox = roundRow.checkbox;
            outerRoundCheckbox.helpTip = getLabel(LABELS.tooltip.outerRound);
            outerRoundInput = roundRow.input;

            var edgeScaleRow = addCheckboxValueRow(outerPanel, LABELS.checkbox.edgeScale, defaultValueText(GENERATION_SETTINGS.defaultEdgeScaleMm), true);
            outerEdgeScaleCheckbox = edgeScaleRow.checkbox;
            outerEdgeScaleCheckbox.value = true;
            outerEdgeScaleCheckbox.helpTip = getLabel(LABELS.tooltip.edgeScale);
            outerEdgeScaleValueGroup = edgeScaleRow.valueGroup;
            outerEdgeScaleInput = edgeScaleRow.input;
            outerEdgeScaleInput.active = true;

            buildStrokeCapPanel(outerPanel);

            keepOuterCheckbox.onClick = function () {
                applyOuterAreaEnabledState();
                requestPreview();
            };

            outerRoundCheckbox.onClick = function () {
                if (outerRoundCheckbox.value) {
                    /* ONにしたとき、0なら既定値を入れる / Fill in the default radius when enabled at zero */
                    fillDefaultIfZero(outerRoundInput, defaultValueText(GENERATION_SETTINGS.defaultRoundMm));

                    /* 角丸と辺の伸縮は同時に使えないので、ONにした側を残して他方をOFFにする
                       The radius and the edge scale cannot be combined, so turning one on turns the other off */
                    outerEdgeScaleCheckbox.value = false;
                }

                applyOuterAreaEnabledState();
                requestPreview();
            };

            outerEdgeScaleCheckbox.onClick = function () {
                if (outerEdgeScaleCheckbox.value) outerRoundCheckbox.value = false;

                applyOuterAreaEnabledState();
                requestPreview();
            };

            outerRoundInput.onChanging = requestPreview;

            outerEdgeScaleInput.onChanging = function () {
                if (!outerEdgeScaleCheckbox.value) return;
                applyStrokeCapPanelEnabledState();
                requestPreview();
            };

            applyOuterAreaEnabledState();
        }

        /**
         * ［線端］パネルを組み立てます。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @returns {void}
         */
        function buildStrokeCapPanel(parent) {
            strokeCapPanel = addRadioPanel(parent, getLabel(LABELS.panel.strokeCap));
            strokeCapPanel.helpTip = getLabel(LABELS.tooltip.strokeCap);

            capButtRadio = strokeCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capButt));
            capRoundRadio = strokeCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capRound));
            capProjectRadio = strokeCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capProject));

            /* 初期値は基準の長方形の線端を優先 / Prefer the base rectangle's cap */
            var currentCap = (baseRects.length > 0 && baseRects[0].stroked) ? baseRects[0].strokeCap : null;
            if (currentCap === StrokeCap.ROUNDENDCAP) capRoundRadio.value = true;
            else if (currentCap === StrokeCap.PROJECTINGENDCAP) capProjectRadio.value = true;
            else capButtRadio.value = true;

            bindAll([capButtRadio, capRoundRadio, capProjectRadio], "onClick", requestPreview);
        }

        /**
         * ［タイトルエリア］パネルを組み立てます。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildTitlePanel(parent) {
            var titlePanel = addPanel(parent, getLabel(LABELS.panel.titleArea), 10);

            // 有効 ＋ 幅／高さ（1行）
            var sizeRow = addToggleValueRow(titlePanel, LABELS.fieldLabel.titleSize, LABELS.tooltip.titleSize);
            titleCheckbox = sizeRow.checkbox;
            titleCheckbox.helpTip = getLabel(LABELS.tooltip.titleArea);
            titleSizeInput = sizeRow.input;

            // 位置（上／下／左／右）
            titlePositionGroup = addRow(titlePanel);
            titleTopRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.top));
            titleBottomRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.bottom));
            titleLeftRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.left));
            titleRightRadio = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.right));
            titleTopRadio.value = true;

            // 塗り／線
            titleOptionGroup = addRow(titlePanel);
            titleFillCheckbox = titleOptionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.fill));
            titleFillCheckbox.helpTip = getLabel(LABELS.tooltip.titleFill);

            titleLineCheckbox = titleOptionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.titleDivider));
            titleLineCheckbox.value = true;
            titleLineCheckbox.helpTip = getLabel(LABELS.tooltip.titleDivider);

            // 仕切り線の伸縮（両端の詰め量）
            var edgeScaleRow = addCheckboxValueRow(titlePanel, LABELS.checkbox.dividerScale, "0", true);
            titleEdgeScaleRow = edgeScaleRow.row;
            titleEdgeScaleCheckbox = edgeScaleRow.checkbox;
            titleEdgeScaleCheckbox.helpTip = getLabel(LABELS.tooltip.dividerScale);
            titleEdgeScaleInput = edgeScaleRow.input;

            titleCheckbox.onClick = function () {
                /* ONにしたとき、サイズが0ならデフォルト値を入れる
                   Fill in the default size when enabled at zero */
                if (titleCheckbox.value && !hasPositiveValue(titleSizeInput)) {
                    titleSizeInput.text = String(calcDefaultTitleSize());
                }

                applyTitleAreaEnabledState();
                requestPreview();
            };

            titleSizeInput.onChanging = function () {
                /* OFFのときは入力だけ受け付け、プレビューには反映しない
                   While unchecked the field just stores the value; nothing is previewed */
                if (!titleCheckbox.value) return;

                applyTitleAreaEnabledState();
                requestPreview();
            };

            bindAll([titleTopRadio, titleBottomRadio, titleLeftRadio, titleRightRadio], "onClick", requestPreview);

            titleFillCheckbox.onClick = requestPreview;

            titleLineCheckbox.onClick = function () {
                applyTitleAreaEnabledState();
                requestPreview();
            };

            titleEdgeScaleCheckbox.onClick = function () {
                applyTitleEdgeScaleEnabledState();
                requestPreview();
            };

            titleEdgeScaleInput.onChanging = function () {
                if (!titleEdgeScaleCheckbox.value) return;
                requestPreview();
            };

            applyTitleEdgeScaleEnabledState();
            applyTitleAreaEnabledState();
        }

        /**
         * ［内側エリア］パネルを組み立てます（オフセット／列・行／線の種類）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildInnerPanel(parent) {
            var innerPanel = addPanel(parent, getLabel(LABELS.panel.innerArea));

            /* 単位はパネル名に入れているので入力欄には付けない
               The unit is in the panel title, so the fields carry none */
            var offsetPanel = addPanel(innerPanel, getLabel(LABELS.panel.offset) + unitSuffix());
            innerOffsetFields = buildLinkedQuadUI(offsetPanel, String(calcDefaultInnerOffset()), 3, "", LABELS.tooltip.offset);

            buildInnerGridPanels(innerPanel);
            buildLineTypePanel(innerPanel);

            /* 初期状態：列／行が1/1なら分割線はディム（OFF）
               Initially the dividers are dimmed while the grid is 1 by 1 */
            applyInnerDividerEnabledState(toCount(columnCountInput.text), toCount(rowCountInput.text), false);
        }

        /**
         * 内側エリアの列・行と、塗り／分割線のオプションを組み立てます。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @returns {void}
         */
        function buildInnerGridPanels(parent) {
            var gridWrapper = parent.add("group");
            gridWrapper.orientation = "row";
            gridWrapper.alignChildren = ["left", "top"];
            gridWrapper.alignment = ["fill", "top"];

            var gridColumn = gridWrapper.add("group");
            gridColumn.orientation = "column";
            gridColumn.alignChildren = ["left", "top"];
            gridColumn.alignment = ["left", "top"];
            gridColumn.spacing = 12;

            var columnFields = addGridCountPanel(gridColumn, LABELS.panel.columns, LABELS.fieldLabel.columnCount);
            columnCountInput = columnFields.countInput;
            columnGutterInput = columnFields.gutterInput;

            var rowFields = addGridCountPanel(gridColumn, LABELS.panel.rows, LABELS.fieldLabel.rowCount);
            rowCountInput = rowFields.countInput;
            rowGutterInput = rowFields.gutterInput;

            // 塗り・分割線のオプション（中央寄せ）
            var optionWrapper = gridColumn.add("group");
            optionWrapper.orientation = "row";
            optionWrapper.alignChildren = ["center", "center"];
            optionWrapper.alignment = ["fill", "top"];

            var optionGroup = addRow(optionWrapper);
            optionGroup.alignment = ["center", "center"];

            innerFillCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.fill));
            innerFillCheckbox.helpTip = getLabel(LABELS.tooltip.innerFill);

            innerDividerCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.divider));
            innerDividerCheckbox.helpTip = getLabel(LABELS.tooltip.divider);

            bindAll([columnCountInput, rowCountInput], "onChanging", function () {
                /* 空欄のあいだは書き換えない（打ち直すたびに「1」が残って桁が増えるため）
                   Leave a blank field alone; rewriting it would prepend a 1 to the next digit */
                snapCountInput(columnCountInput);
                snapCountInput(rowCountInput);

                var columnCount = toCount(columnCountInput.text);
                var rowCount = toCount(rowCountInput.text);

                columnGutterInput.enabled = (columnCount > 1);
                rowGutterInput.enabled = (rowCount > 1);
                syncStepperStates();
                applyInnerDividerEnabledState(columnCount, rowCount, true);
                requestPreview();
            });

            bindAll([columnGutterInput, rowGutterInput], "onChanging", function () {
                /* ガターが入ったら塗りを自動ON（手動操作があれば尊重）
                   Turn the fill on once a gutter is set, unless the user set it manually */
                var gutter = parseFloat(this.text);
                if (!innerFillManuallySet && !isNaN(gutter) && gutter !== 0) {
                    innerFillCheckbox.value = true;
                }
                requestPreview();
            });

            innerFillCheckbox.onClick = function () {
                /* ユーザーが操作したら以後は自動ONしない / Stop auto-enabling once the user decides */
                innerFillManuallySet = true;
                requestPreview();
            };
        }

        /**
         * 内側エリアの［線の種類］パネルを組み立てます。
         *
         * @param {Panel} parent - 追加先のパネル。
         * @returns {void}
         */
        function buildLineTypePanel(parent) {
            lineTypePanel = addRadioPanel(parent, getLabel(LABELS.panel.lineType));

            lineSolidRadio = lineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineSolid));
            lineDashRadio = lineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineDash));
            lineDotsRadio = lineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineDots));
            lineSolidRadio.value = true;

            bindAll([lineSolidRadio, lineDashRadio, lineDotsRadio], "onClick", requestPreview);

            innerDividerCheckbox.onClick = function () {
                lineTypePanel.enabled = (innerDividerCheckbox.enabled && innerDividerCheckbox.value);
                requestPreview();
            };
        }

        /**
         * ［画面表示］タブを組み立てます（ズームとパン、表示コマンド）。
         *
         * @param {Group} parent - 追加先のグループ。
         * @returns {void}
         */
        function buildDisplayTab(parent) {
            var zoomPanPanel = addPanel(parent, getLabel(LABELS.panel.zoomPan), 10);

            var zoomPanGroup = zoomPanPanel.add("group");
            zoomPanGroup.orientation = "column";
            zoomPanGroup.alignChildren = "left";
            zoomPanGroup.spacing = 8;

            viewControl.buildUI(zoomPanGroup, {
                labelWidth: DIALOG_LAYOUT.viewLabelWidth,
                sliderWidth: DIALOG_LAYOUT.viewSliderWidth,
                zoomLabel: labelText(LABELS.fieldLabel.zoom),
                panXLabel: labelText(LABELS.fieldLabel.panX),
                panYLabel: labelText(LABELS.fieldLabel.panY),
                sliderHelpTip: getLabel(LABELS.tooltip.viewSlider)
            });

            var viewCommandPanel = addPanel(parent, getLabel(LABELS.panel.viewCommands), 10);

            var viewCommandGroup = viewCommandPanel.add("group");
            viewCommandGroup.orientation = "column";
            viewCommandGroup.alignChildren = ["left", "top"];
            viewCommandGroup.alignment = ["fill", "top"];
            viewCommandGroup.spacing = 6;

            addViewCommandButton(viewCommandGroup, LABELS.button.fitArtboard, function () {
                app.executeMenuCommand("fitin");
            });
            addViewCommandButton(viewCommandGroup, LABELS.button.actualSize, function () {
                app.executeMenuCommand("actualsize");
            });
            addViewCommandButton(viewCommandGroup, LABELS.button.fitAll, function () {
                app.executeMenuCommand("fitall");
            });
            addViewCommandButton(viewCommandGroup, LABELS.button.zoomOut10, function () {
                /* 現在の表示倍率を10%縮小（スライダーも追従）
                   Shrink the current zoom by 10%; the slider follows */
                viewControl.zoomBy(0.9);
            });
        }

        // =========================================
        // セッションの保存と復元 / Save and restore the session
        // =========================================
        /* 保存と復元はこの定義表を共有します（項目の追加・変更はここだけ）
           Save and restore share these tables, so a field is defined in one place */

        /**
         * セッションに保存する入力欄・チェックボックスの一覧を返します。
         *
         * @returns {Array} [コントロール, 保存先のパス] の配列。
         */
        function getSessionControls() {
            return [
                [marginFields.top, "margin.top"],
                [marginFields.bottom, "margin.bottom"],
                [marginFields.left, "margin.left"],
                [marginFields.right, "margin.right"],
                [marginFields.linkToggle, "margin.link"],
                [frameCheckbox, "frame.enabled"],
                [frameWidthInput, "frame.width"],
                [bleedCheckbox, "frame.bleed"],
                [frameRoundCheckbox, "frame.round.enabled"],
                [frameRoundInput, "frame.round.value"],
                [keepOuterCheckbox, "outer.keepOuter"],
                [outerRoundCheckbox, "outer.round.enabled"],
                [outerRoundInput, "outer.round.value"],
                [outerEdgeScaleCheckbox, "outer.edgeScale.enabled"],
                [outerEdgeScaleInput, "outer.edgeScale.value"],
                [titleCheckbox, "title.enabled"],
                [titleSizeInput, "title.size"],
                [titleFillCheckbox, "title.fill"],
                [titleLineCheckbox, "title.line"],
                [titleEdgeScaleCheckbox, "title.edgeScale.enabled"],
                [titleEdgeScaleInput, "title.edgeScale.value"],
                [innerOffsetFields.top, "inner.offset.top"],
                [innerOffsetFields.bottom, "inner.offset.bottom"],
                [innerOffsetFields.left, "inner.offset.left"],
                [innerOffsetFields.right, "inner.offset.right"],
                [innerOffsetFields.linkToggle, "inner.offset.link"],
                [columnCountInput, "inner.grid.columns"],
                [columnGutterInput, "inner.grid.columnGutter"],
                [rowCountInput, "inner.grid.rows"],
                [rowGutterInput, "inner.grid.rowGutter"],
                [innerFillCheckbox, "inner.grid.fill"],
                [innerDividerCheckbox, "inner.grid.divider"],
                [previewCheckbox, "preview"]
            ];
        }

        /**
         * セッションに保存するラジオボタン群の一覧を返します。
         *
         * @returns {Array} [保存先のパス, 保存値→コントロールの対応] の配列。
         */
        function getSessionRadioGroups() {
            return [
                ["outer.strokeCap", { butt: capButtRadio, round: capRoundRadio, project: capProjectRadio }],
                ["title.position", { top: titleTopRadio, bottom: titleBottomRadio, left: titleLeftRadio, right: titleRightRadio }],
                ["inner.grid.lineType", { solid: lineSolidRadio, dash: lineDashRadio, dots: lineDotsRadio }]
            ];
        }

        /**
         * ラジオボタン群から、選択中の保存値を返します。
         *
         * @param {Object} radiosByKey - 保存値→コントロールの対応。
         * @returns {string|undefined} 選択中の保存値。どれも選択されていなければ undefined。
         */
        function readRadioKey(radiosByKey) {
            for (var key in radiosByKey) {
                if (!radiosByKey.hasOwnProperty(key)) continue;
                if (radiosByKey[key].value) return key;
            }
            return undefined;
        }

        /**
         * 前回のダイアログ設定を復元します（Illustratorの起動中のみ有効）。
         *
         * @returns {void}
         */
        function restoreDialogState() {
            var state = loadSessionState();
            var controls = getSessionControls();
            var radioGroups = getSessionRadioGroups();
            var i, value;

            for (i = 0; i < controls.length; i++) {
                value = getStateValue(state, controls[i][1]);
                if (typeof value === "undefined") continue;

                if (controls[i][0].type === "edittext") controls[i][0].text = String(value);
                else if (controls[i][0].isLinkToggle) setLinkToggleValue(controls[i][0], !!value);
                else controls[i][0].value = !!value;
            }

            for (i = 0; i < radioGroups.length; i++) {
                var radio = radioGroups[i][1][getStateValue(state, radioGroups[i][0])];
                if (radio) radio.value = true;
            }

            /* 手動操作のフラグも復元する（復元直後に自動ONで上書きされないように）
               Restore the manual-input flags too, so the auto-on rules do not override them */
            innerFillManuallySet = !!getStateValue(state, "manual.innerFill");
            bleedManuallySet = !!getStateValue(state, "manual.bleed");

            /* 角丸と辺の伸縮は同時にONにできない（角丸を優先）
               The radius and the edge scale cannot both be on; the radius wins */
            if (outerRoundCheckbox.value) outerEdgeScaleCheckbox.value = false;

            /* 復元は「サイズが0→>0になった瞬間」ではないため、
               タイトルの［線］が自動ONで上書きされないように直前の状態をそろえる
               A restore is not a zero-to-positive transition, so seed the previous state */
            titleHadSize = (titleCheckbox.value && hasPositiveValue(titleSizeInput));

            /* 他のコントロールに依存する有効／無効を反映し直す
               Re-apply the enabled states that depend on other controls */
            marginFields.applyLinkState();
            innerOffsetFields.applyLinkState();
            applyOuterAreaEnabledState();
            applyTitleEdgeScaleEnabledState();
            applyTitleAreaEnabledState();
            applyFrameEnabledState();
            applyInnerDividerEnabledState(toCount(columnCountInput.text), toCount(rowCountInput.text), false);
        }

        /**
         * 現在のダイアログ設定をセッションに保存します。
         *
         * @returns {void}
         */
        function saveDialogState() {
            var state = {};
            var controls = getSessionControls();
            var radioGroups = getSessionRadioGroups();
            var i;

            for (i = 0; i < controls.length; i++) {
                setStateValue(state, controls[i][1],
                    (controls[i][0].type === "edittext") ? controls[i][0].text : controls[i][0].value);
            }

            for (i = 0; i < radioGroups.length; i++) {
                setStateValue(state, radioGroups[i][0], readRadioKey(radioGroups[i][1]));
            }

            /* 手動操作のフラグ（次回の自動ONを抑えるために保存）
               The manual-input flags, saved so the auto-on rules stay suppressed next time */
            setStateValue(state, "manual.innerFill", innerFillManuallySet);
            setStateValue(state, "manual.bleed", bleedManuallySet);

            saveSessionState(state);
        }

        // =========================================
        // プレビューと生成 / Preview and generation
        // =========================================
        // collectOptions()        : UIを読み、pt単位の生成条件にまとめる
        // generateFromOptions()   : 生成条件からオブジェクトを作る
        // rebuildGeneratedItems() : 前回分を消して生成し、再描画する（唯一の入口）
        // -----------------------------------------

        /**
         * プレビューがONのときだけプレビューを更新します。
         *
         * @returns {void}
         */
        function requestPreview() {
            if (previewCheckbox.value) rebuildGeneratedItems(false);
        }

        /**
         * プレビューまたは最終結果を作り直して再描画します。
         *
         * @param {boolean} isFinal - 実行（OK）時の最終生成なら true。
         * @returns {void}
         */
        function rebuildGeneratedItems(isFinal) {
            removeGeneratedItems();
            generateFromOptions(collectOptions(), isFinal);
            app.redraw();
        }

        /**
         * プレビューを消して元の状態に戻します。
         *
         * @returns {void}
         */
        function clearPreview() {
            removeGeneratedItems();

            /* アートボード基準の一時矩形を先に破棄（baseRectsに残っていると無効参照になる）
               Remove the temporary artboard rectangle first to avoid stale references */
            removeArtboardBaseRect();

            /* 角丸プレビューなどで隠した元オブジェクトを表示に戻す
               Show the originals that the preview had hidden */
            setBaseRectsHidden(false);

            app.redraw();
        }

        /**
         * UIの入力値を読み、pt単位の生成条件にまとめます。
         *
         * @returns {Object} 生成条件。
         */
        function collectOptions() {
            var columnCount = toCount(columnCountInput.text);
            var rowCount = toCount(rowCountInput.text);
            var bleedPt = bleedCheckbox.value ? mmToPt(GENERATION_SETTINGS.bleedMm) : 0;

            var options = {
                /* マージン（アートボード基準のみ） / Margins, artboard-based runs only */
                marginPt: {
                    top: toPositivePt(marginFields.top.text),
                    right: toPositivePt(marginFields.right.text),
                    bottom: toPositivePt(marginFields.bottom.text),
                    left: toPositivePt(marginFields.left.text)
                },

                /* フレーム：幅に裁ち落としを加算し、アートボードも裁ち落としぶん広げる
                   Frame: the bleed widens both the frame and the artboard bounds */
                bleedPt: bleedPt,
                framePt: (frameCheckbox.value ? toPositivePt(frameWidthInput.text) : 0) + bleedPt,
                frameRoundPt: readCheckedPt(frameRoundCheckbox, frameRoundInput),

                /* 外側エリア：辺の伸縮と角丸（角丸はタイトル帯にも使う）
                   Outer area: the edge scale and the radius, which the title band reuses */
                outerEdgeScalePt: getOuterEdgeScaleValue() * rulerUnit.pointsPerUnit,
                outerRoundPt: readCheckedPt(outerRoundCheckbox, outerRoundInput),

                /* タイトルエリア：仕切り線の詰め量は入力値の符号を反転（＋で両端が短くなる）
                   Title area: the divider inset is the negated input (positive shortens the line) */
                titleSizePt: titleCheckbox.value ? toPositivePt(titleSizeInput.text) : 0,
                titleDividerInsetPt: (titleLineCheckbox.value && titleEdgeScaleCheckbox.value) ? -toPt(titleEdgeScaleInput.text) : 0,

                /* 内側エリア：オフセット / Inner area offsets */
                innerOffsetPt: {
                    top: toPositivePt(innerOffsetFields.top.text),
                    right: toPositivePt(innerOffsetFields.right.text),
                    bottom: toPositivePt(innerOffsetFields.bottom.text),
                    left: toPositivePt(innerOffsetFields.left.text)
                },

                /* 内側エリア：列・行とガター（列／行が1のときガターは0扱い）
                   Columns, rows and gutters; gutters are ignored for a single column or row */
                columnCount: columnCount,
                rowCount: rowCount,
                columnGutterPt: (columnCount > 1) ? toPositivePt(columnGutterInput.text) : 0,
                rowGutterPt: (rowCount > 1) ? toPositivePt(rowGutterInput.text) : 0
            };

            syncDependentControls(options);
            return options;
        }

        /**
         * 生成条件からオブジェクトを作成します（プレビュー・実行の共通処理）。
         *
         * @param {Object} options - collectOptions() が返す生成条件。
         * @param {boolean} isFinal - 実行（OK）時の最終生成なら true。
         * @returns {void}
         */
        function generateFromOptions(options, isFinal) {
            if (isArtboardBased) rebuildArtboardBaseRect(options.marginPt);

            var splitEdges = (options.outerEdgeScalePt !== 0);
            var i;

            if (splitEdges) {
                /* 4辺に分解するため、元の長方形は常に隠す
                   Always hide the original rectangle while the edges are split */
                setBaseRectsHidden(true);
                if (keepOuterCheckbox.value) {
                    for (i = 0; i < baseRects.length; i++) {
                        createOuterEdgeLines(baseRects[i], options.outerEdgeScalePt);
                    }
                }
            } else {
                /* 外枠の表示は「外枠を残す」に従う
                   Show the original frame according to the keep-outer checkbox */
                setBaseRectsHidden(!keepOuterCheckbox.value);
            }

            /* フレーム（アートボード基準） / Frame, based on the artboard */
            if (options.framePt > 0 && baseRects.length > 0) {
                createFrame(baseRects[0].layer, getActiveArtboardBounds(options.bleedPt), options.framePt, options.frameRoundPt);
            }

            if (options.titleSizePt > 0) {
                for (i = 0; i < baseRects.length; i++) {
                    createTitleArea(baseRects[i], options);
                }
            }

            if (!splitEdges) applyOuterAreaRound(options.outerRoundPt, isFinal);

            for (i = 0; i < baseRects.length; i++) {
                createInnerArea(baseRects[i], options);
            }
        }

        /**
         * アクティブなアートボードの矩形を返します（裁ち落としのぶん外側に広げます）。
         *
         * @param {number} bleedPt - 裁ち落とし幅（pt）。
         * @returns {number[]} [左, 上, 右, 下] の座標。
         */
        function getActiveArtboardBounds(bleedPt) {
            var rect = getActiveArtboardRect(); // [L, T, R, B]
            return [rect[0] - bleedPt, rect[1] + bleedPt, rect[2] + bleedPt, rect[3] - bleedPt];
        }

        /**
         * 長方形の各辺を、伸縮させた直線として生成します（4辺に分解）。
         *
         * @param {PathItem} baseRect - 基準の長方形。
         * @param {number} edgeScalePt - 伸縮量（pt。正で伸ばし、負で縮める）。
         * @returns {void}
         */
        function createOuterEdgeLines(baseRect, edgeScalePt) {
            var points = baseRect.pathPoints;
            var edgeCount = baseRect.closed ? points.length : points.length - 1;
            var scaleAmount = Math.abs(edgeScalePt);

            /* 正なら両端を外へ、負なら内へ動かす / Positive extends the ends, negative pulls them in */
            var direction = (edgeScalePt >= 0) ? -1 : 1;

            for (var i = 0; i < edgeCount; i++) {
                var startAnchor = points[i].anchor;
                var endAnchor = points[(i + 1) % points.length].anchor;

                var dx = endAnchor[0] - startAnchor[0];
                var dy = endAnchor[1] - startAnchor[1];
                var edgeLength = Math.sqrt(dx * dx + dy * dy);

                /* 長さ0の辺（重なったアンカー）は計算できないので飛ばす
                   A zero-length edge (duplicated anchors) cannot be scaled */
                if (!(edgeLength > 0)) continue;

                /* 縮めるとき、辺が縮小量の2倍以下なら線が反転するので生成しない
                   While shrinking, an edge shorter than twice the amount would turn inside out */
                if (edgeScalePt < 0 && edgeLength <= scaleAmount * 2) continue;

                var ratio = scaleAmount / edgeLength;
                var edgeLine = trackGeneratedItem(baseRect.layer.pathItems.add());
                edgeLine.setEntirePath([
                    [startAnchor[0] + dx * ratio * direction, startAnchor[1] + dy * ratio * direction],
                    [endAnchor[0] - dx * ratio * direction, endAnchor[1] - dy * ratio * direction]
                ]);

                setStroke(edgeLine, baseRect.strokeColor, baseRect.strokeWidth);
                edgeLine.strokeCap = getSelectedStrokeCap();

                // 外枠（4辺）として識別できるようタグ付け
                tagItem(edgeLine, TAG_OUTER_EDGE);
            }
        }

        /**
         * 外側エリアの角丸を適用します（辺の伸縮OFFのときだけ）。
         *
         * プレビューでは元の長方形を隠し、同じ位置に角丸用の一時矩形を作ります。
         * 実行時は元の長方形にライブエフェクトを適用します。
         *
         * @param {number} radiusPt - 角丸の半径（pt）。
         * @param {boolean} isFinal - 実行（OK）時の最終生成なら true。
         * @returns {void}
         */
        function applyOuterAreaRound(radiusPt, isFinal) {
            if (!keepOuterCheckbox.value || outerEdgeScaleCheckbox.value || !(radiusPt > 0)) return;

            for (var i = 0; i < baseRects.length; i++) {
                var baseRect = baseRects[i];

                if (isFinal) {
                    applyRoundCornersEffect(baseRect, radiusPt);
                    baseRect.hidden = false;
                    continue;
                }

                var bounds = baseRect.geometricBounds; // [L, T, R, B]
                var width = bounds[2] - bounds[0];
                var height = bounds[1] - bounds[3];
                if (!(width > 0) || !(height > 0)) continue;

                baseRect.hidden = true;

                var previewRect = trackGeneratedItem(baseRect.layer.pathItems.rectangle(bounds[1], bounds[0], width, height));
                copyAppearance(baseRect, previewRect);
                tagItem(previewRect, TAG_OUTER_ROUND);
                applyRoundCornersEffect(previewRect, radiusPt);
            }
        }

        /**
         * フレーム（外側はグレー、内側は透明の穴）を作成します。
         *
         * 穴あきはグループに Live Pathfinder Exclude を適用して作ります。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {number[]} bounds - 基準領域 [左, 上, 右, 下]。
         * @param {number} framePt - フレームの幅（pt）。
         * @param {number} roundPt - 穴側の角丸の半径（pt）。
         * @returns {void}
         */
        function createFrame(layer, bounds, framePt, roundPt) {
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
            var width = right - left;
            var height = top - bottom;

            /* 内側は最終的に穴になる矩形 / The inner rectangle becomes the hole */
            var holeWidth = width - framePt * 2;
            var holeHeight = height - framePt * 2;
            if (!(holeWidth > 0) || !(holeHeight > 0)) return;

            /* パスファインダーはメニューコマンド経由のため、失敗しても続行する
               The pathfinder runs as a menu command, so keep going if it fails */
            try {
                var outerRect = trackGeneratedItem(layer.pathItems.rectangle(top, left, width, height));
                setGrayFill(outerRect, GRAY_TINTS.frame);

                var holeRect = trackGeneratedItem(layer.pathItems.rectangle(top - framePt, left + framePt, holeWidth, holeHeight));
                /* 内側は一時的な塗り（最終的に Exclude の結果で穴になる）
                   A temporary fill; the Exclude result turns it into a hole */
                setGrayFill(holeRect, 0);

                // 角丸は「内側の長方形」に適用する（穴側を丸める）
                applyRoundCornersEffect(holeRect, roundPt);

                var frameGroup = trackGeneratedItem(layer.groupItems.add());
                outerRect.move(frameGroup, ElementPlacement.PLACEATEND);
                holeRect.move(frameGroup, ElementPlacement.PLACEATEND);

                var frameItem = applyPathfinderExclude(frameGroup);
                if (frameItem !== frameGroup) trackGeneratedItem(frameItem);

                tagItem(frameItem, TAG_FRAME_FILL);
                sendToBack(frameItem);
            } catch (e) { }
        }

        /**
         * グループに Live Pathfinder Exclude を適用し、結果のオブジェクトを返します。
         *
         * 選択を一時的に置き換えるため、実行後に元の選択へ戻します。
         *
         * @param {GroupItem} group - 対象のグループ。
         * @returns {PageItem} Exclude の結果（取得できない場合は元のグループ）。
         */
        function applyPathfinderExclude(group) {
            var previousSelection = doc.selection;
            var resultItem = group;

            try {
                doc.selection = null;
                group.selected = true;
                app.executeMenuCommand('Live Pathfinder Exclude');

                /* 結果は selection の先頭に入る / The result lands at the head of the selection */
                if (doc.selection.length > 0) resultItem = doc.selection[0];
            } catch (e) { }

            try { doc.selection = previousSelection; } catch (e) { }

            return resultItem;
        }

        /**
         * 選択中のタイトルの位置を返します。
         *
         * @returns {string} "top" / "bottom" / "left" / "right"。
         */
        function getTitlePositionKey() {
            if (titleRightRadio.value) return "right";
            if (titleBottomRadio.value) return "bottom";
            if (titleLeftRadio.value) return "left";
            return "top";
        }

        /**
         * タイトルエリアの配置を計算します（位置による分岐をここに集約）。
         *
         * @param {number[]} bounds - 基準領域 [左, 上, 右, 下]。
         * @param {number} sizePt - タイトルエリアの幅／高さ（pt）。
         * @param {number} dividerInsetPt - 仕切り線の両端の詰め量（pt。＋で短く／−で長く）。
         * @returns {Object|null} band（帯の矩形）／divider（仕切り線）／inner（帯を除いた領域）。
         *                        領域が成立しない場合は null。
         */
        function calcTitleAreaLayout(bounds, sizePt, dividerInsetPt) {
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
            var width = right - left;
            var height = top - bottom;
            if (!(width > 0) || !(height > 0)) return null;

            var positionKey = getTitlePositionKey();
            var isHorizontal = (positionKey === "top" || positionKey === "bottom");
            if (sizePt >= (isHorizontal ? height : width)) return null;

            /* 横並び（上／下）：帯は横いっぱい、仕切り線は水平
               Top or bottom: a full-width band with a horizontal divider */
            if (isHorizontal) {
                var dividerY = (positionKey === "top") ? (top - sizePt) : (bottom + sizePt);

                var startX = left + dividerInsetPt;
                var endX = right - dividerInsetPt;
                if (startX >= endX) { startX = left; endX = right; }

                return {
                    band: {
                        top: (positionKey === "top") ? top : (bottom + sizePt),
                        left: left,
                        width: width,
                        height: sizePt
                    },
                    divider: [[startX, dividerY], [endX, dividerY]],
                    inner: (positionKey === "top") ? [left, dividerY, right, bottom] : [left, top, right, dividerY]
                };
            }

            /* 縦並び（左／右）：帯は縦いっぱい、仕切り線は垂直
               Left or right: a full-height band with a vertical divider */
            var dividerX = (positionKey === "left") ? (left + sizePt) : (right - sizePt);

            var startY = top - dividerInsetPt;
            var endY = bottom + dividerInsetPt;
            if (startY <= endY) { startY = top; endY = bottom; }

            return {
                band: {
                    top: top,
                    left: (positionKey === "left") ? left : (right - sizePt),
                    width: sizePt,
                    height: height
                },
                divider: [[dividerX, startY], [dividerX, endY]],
                inner: (positionKey === "left") ? [dividerX, top, right, bottom] : [left, top, dividerX, bottom]
            };
        }

        /**
         * タイトル帯の塗りと、本文との仕切り線を作成します。
         *
         * @param {PathItem} baseRect - 基準の長方形。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {void}
         */
        function createTitleArea(baseRect, options) {
            var layout = calcTitleAreaLayout(baseRect.geometricBounds, options.titleSizePt, options.titleDividerInsetPt);
            if (!layout) return;

            var layer = baseRect.layer;

            if (titleFillCheckbox.value) {
                var bandRect = trackGeneratedItem(layer.pathItems.rectangle(
                    layout.band.top, layout.band.left, layout.band.width, layout.band.height));
                setGrayFill(bandRect, GRAY_TINTS.titleBand);
                tagItem(bandRect, TAG_TITLE_FILL);

                /* 外側エリアの角丸値で、位置に応じた2角だけ角丸にする
                   Round the two corners that match the title position, with the outer radius */
                roundCornersOnSide(bandRect, getTitlePositionKey(), options.outerRoundPt);

                /* 背面へ（他の罫線や要素の下に敷く） / Send behind the rules */
                sendToBack(bandRect);
            }

            if (titleLineCheckbox.value) {
                /* 他の生成物と同じレイヤーに作る（doc.activeLayer は使わない）
                   Create it on the same layer as the other generated items */
                var dividerLine = trackGeneratedItem(layer.pathItems.add());
                dividerLine.setEntirePath(layout.divider);
                setStroke(dividerLine, makeGrayColor(GRAY_TINTS.rule), 1);
                tagItem(dividerLine, TAG_TITLE_DIVIDER);
            }
        }

        /**
         * 長方形の、指定した辺に接する2角だけを角丸にします（同じパスを書き換えます）。
         *
         * @param {PathItem} rect - 対象の長方形（閉じたパス）。
         * @param {string} sideKey - "top" / "bottom" / "left" / "right"。
         * @param {number} radiusPt - 角丸の半径（pt）。0以下なら何もしません。
         * @returns {void}
         */
        function roundCornersOnSide(rect, sideKey, radiusPt) {
            var bounds = rect.geometricBounds; // [L, T, R, B]
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];

            /* 半径は辺の半分までに抑える / Cap the radius at half the shorter side */
            var radius = Math.min(radiusPt, (right - left) / 2, (top - bottom) / 2);
            if (!(radius > 0)) return;

            /* 四分円をベジェで近似するハンドル長 / Bezier handle length for a quarter circle */
            var handleLength = radius * 0.5522847498307936;
            var POINT_TYPE = (typeof PointType !== "undefined") ? PointType : { CORNER: 0, SMOOTH: 1 };

            var roundTopLeft = (sideKey === "top" || sideKey === "left");
            var roundTopRight = (sideKey === "top" || sideKey === "right");
            var roundBottomRight = (sideKey === "bottom" || sideKey === "right");
            var roundBottomLeft = (sideKey === "bottom" || sideKey === "left");

            var cornerPoints = [];

            /**
             * アンカーと左右のハンドルを1点ぶん記録します（ハンドル省略時はアンカーと同じ位置）。
             *
             * @param {number[]} anchor - アンカーの座標 [x, y]。
             * @param {number[]} [leftDirection] - 左方向ハンドルの座標。
             * @param {number[]} [rightDirection] - 右方向ハンドルの座標。
             * @returns {void}
             */
            function addPoint(anchor, leftDirection, rightDirection) {
                cornerPoints.push({
                    anchor: anchor,
                    leftDirection: leftDirection || anchor,
                    rightDirection: rightDirection || anchor,
                    pointType: (leftDirection || rightDirection) ? POINT_TYPE.SMOOTH : POINT_TYPE.CORNER
                });
            }

            // 始点：左上（上辺側）/ Start at the top-left corner, on the top edge
            if (roundTopLeft) addPoint([left + radius, top], [left + radius - handleLength, top], null);
            else addPoint([left, top], null, null);

            // 右上 / Top-right
            if (roundTopRight) {
                addPoint([right - radius, top], null, [right - radius + handleLength, top]);
                addPoint([right, top - radius], [right, top - radius + handleLength], null);
            } else {
                addPoint([right, top], null, null);
            }

            // 右下 / Bottom-right
            if (roundBottomRight) {
                addPoint([right, bottom + radius], null, [right, bottom + radius - handleLength]);
                addPoint([right - radius, bottom], [right - radius + handleLength, bottom], null);
            } else {
                addPoint([right, bottom], null, null);
            }

            // 左下 / Bottom-left
            if (roundBottomLeft) {
                addPoint([left + radius, bottom], null, [left + radius - handleLength, bottom]);
                addPoint([left, bottom + radius], [left, bottom + radius - handleLength], null);
            } else {
                addPoint([left, bottom], null, null);
            }

            // 終点：左上（左辺側）/ Close at the top-left corner, on the left edge
            if (roundTopLeft) addPoint([left, top - radius], null, [left, top - radius + handleLength]);

            var anchors = [];
            for (var i = 0; i < cornerPoints.length; i++) {
                anchors.push(cornerPoints[i].anchor);
            }
            rect.setEntirePath(anchors);
            rect.closed = true;

            var pathPoints = rect.pathPoints;
            for (var j = 0; j < cornerPoints.length; j++) {
                pathPoints[j].leftDirection = cornerPoints[j].leftDirection;
                pathPoints[j].rightDirection = cornerPoints[j].rightDirection;
                pathPoints[j].pointType = cornerPoints[j].pointType;
            }
        }

        /**
         * タイトル領域を除いた「内側エリア」の計算領域を返します。
         *
         * @param {PathItem} baseRect - 基準の長方形。
         * @param {number} titleSizePt - タイトルエリアの幅／高さ（pt）。
         * @returns {number[]|null} [左, 上, 右, 下] の座標。成立しない場合は null。
         */
        function getInnerAreaBounds(baseRect, titleSizePt) {
            var bounds = baseRect.geometricBounds; // [L, T, R, B]
            if (!((bounds[2] - bounds[0]) > 0 && (bounds[1] - bounds[3]) > 0)) return null;

            if (!(titleSizePt > 0)) return bounds;

            /* タイトルが入らないサイズのときは帯を作らないので、内側エリアは外形いっぱい
               When the title does not fit, no band is drawn, so the inner area keeps the full bounds */
            var layout = calcTitleAreaLayout(bounds, titleSizePt, 0);
            return layout ? layout.inner : bounds;
        }

        /**
         * 内側エリアのグリッド配置を計算します。
         *
         * @param {number[]} bounds - 基準領域 [左, 上, 右, 下]。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {Object|null} セル配置。成立しない場合は null。
         */
        function calcInnerGrid(bounds, options) {
            var offset = options.innerOffsetPt;

            /* オフセットが0でも内側エリアは描画する / The inner area is drawn even at zero offset */
            var width = (bounds[2] - bounds[0]) - (offset.left + offset.right);
            var height = (bounds[1] - bounds[3]) - (offset.top + offset.bottom);
            if (!(width > 0) || !(height > 0)) return null;

            /* ガターを除いた1セルの大きさ / Cell size with the gutters removed */
            var cellWidth = (width - options.columnGutterPt * (options.columnCount - 1)) / options.columnCount;
            var cellHeight = (height - options.rowGutterPt * (options.rowCount - 1)) / options.rowCount;
            if (!(cellWidth > 0) || !(cellHeight > 0)) return null;

            return {
                left: bounds[0] + offset.left,
                top: bounds[1] - offset.top,
                width: width,
                height: height,
                columnCount: options.columnCount,
                rowCount: options.rowCount,
                columnGutter: options.columnGutterPt,
                rowGutter: options.rowGutterPt,
                cellWidth: cellWidth,
                cellHeight: cellHeight
            };
        }

        /**
         * 内側エリアのセル（塗り）と分割線を作成します。
         *
         * @param {PathItem} baseRect - 基準の長方形（レイヤーと線の色の参照元）。
         * @param {Object} options - collectOptions() が返す生成条件。
         * @returns {void}
         */
        function createInnerArea(baseRect, options) {
            var areaBounds = getInnerAreaBounds(baseRect, options.titleSizePt);
            if (!areaBounds) return;

            var grid = calcInnerGrid(areaBounds, options);
            if (!grid) return;

            var layer = baseRect.layer;

            /* セルの塗りは［塗り］がONのときだけ（プレビューと実行結果を一致させる）
               The cells are only drawn while the fill is on, so the preview matches the result */
            if (innerFillCheckbox.value) createInnerCellFills(layer, grid);

            /* 分割線OFF、または分割できない構成ならここまで
               Stop here when the dividers are off or the grid cannot carry them */
            if (!innerDividerCheckbox.value || !isGridSplittable(grid.columnCount, grid.rowCount)) return;

            /* 分割線の色は基準の長方形から取る（線がなければ K100）
               Take the divider colour from the base rectangle, falling back to K100 */
            createInnerDividers(layer, grid, baseRect.stroked ? baseRect.strokeColor : makeGrayColor(GRAY_TINTS.rule));
        }

        /**
         * 内側エリアのセル（塗り）を作成します。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {Object} grid - calcInnerGrid() が返すセル配置。
         * @returns {void}
         */
        function createInnerCellFills(layer, grid) {
            for (var i = 0; i < grid.rowCount; i++) {
                var cellTop = grid.top - (grid.cellHeight + grid.rowGutter) * i;

                for (var j = 0; j < grid.columnCount; j++) {
                    var cellLeft = grid.left + (grid.cellWidth + grid.columnGutter) * j;

                    var cellRect = trackGeneratedItem(layer.pathItems.rectangle(cellTop, cellLeft, grid.cellWidth, grid.cellHeight));
                    setGrayFill(cellRect, GRAY_TINTS.innerCell);
                    tagItem(cellRect, TAG_INNER_FILL);

                    /* 背面へ（罫線などのパスの下に敷く） / Send behind the rules */
                    sendToBack(cellRect);
                }
            }
        }

        /**
         * 内側エリアの分割線を、各ガターの中心に1本ずつ作成します。
         *
         * @param {Layer} layer - 作成先のレイヤー。
         * @param {Object} grid - calcInnerGrid() が返すセル配置。
         * @param {Color} strokeColor - 分割線の色。
         * @returns {void}
         */
        function createInnerDividers(layer, grid, strokeColor) {
            var bottomY = grid.top - grid.height;
            var rightX = grid.left + grid.width;

            /**
             * 分割線を1本作成します。
             *
             * @param {number[][]} points - [[x1, y1], [x2, y2]] 形式の始点と終点。
             * @returns {void}
             */
            function addDivider(points) {
                var dividerLine = trackGeneratedItem(layer.pathItems.add());
                dividerLine.setEntirePath(points);
                setStroke(dividerLine, strokeColor, 1);
                applyInnerLineStyle(dividerLine);
            }

            // 列の分割線 / Column dividers
            for (var i = 1; i < grid.columnCount; i++) {
                var gutterCenterX = grid.left + (grid.cellWidth * i) + (grid.columnGutter * (i - 1)) + (grid.columnGutter / 2);
                addDivider([[gutterCenterX, grid.top], [gutterCenterX, bottomY]]);
            }

            // 行の分割線 / Row dividers
            for (var j = 1; j < grid.rowCount; j++) {
                var gutterCenterY = grid.top - (grid.cellHeight * j) - (grid.rowGutter * (j - 1)) - (grid.rowGutter / 2);
                addDivider([[grid.left, gutterCenterY], [rightX, gutterCenterY]]);
            }
        }

        /**
         * UIで選択された線端を返します。
         *
         * @returns {StrokeCap} 選択中の線端。
         */
        function getSelectedStrokeCap() {
            if (capRoundRadio.value) return StrokeCap.ROUNDENDCAP;
            if (capProjectRadio.value) return StrokeCap.PROJECTINGENDCAP;
            return StrokeCap.BUTTENDCAP; // 線端なし
        }

        /**
         * 内側の分割線に線種（実線・点線・ドット点線）を適用します。
         *
         * @param {PathItem} dividerLine - 対象の分割線。
         * @returns {void}
         */
        function applyInnerLineStyle(dividerLine) {
            if (lineDotsRadio.value) {
                /* ドット点線：線端を丸型にして strokeDashes=[0, 線幅*2]
                   Dotted line: round caps with strokeDashes = [0, width * 2] */
                dividerLine.strokeWidth = 2;
                dividerLine.strokeCap = StrokeCap.ROUNDENDCAP;
                dividerLine.strokeDashes = [0, dividerLine.strokeWidth * 2];
                try { dividerLine.strokeJoin = StrokeJoin.ROUNDENDJOIN; } catch (e) { }
                return;
            }

            dividerLine.strokeWidth = 1;
            // 点線（ダッシュ）は破線パターン、実線は破線なし
            dividerLine.strokeDashes = lineDashRadio.value ? [4, 2] : [];
            // 線端は外枠の線端設定に合わせる
            dividerLine.strokeCap = getSelectedStrokeCap();
        }

        /**
         * 実行（OK）後の後処理です。
         *
         * 基準の長方形と内側エリアの塗りの扱いを確定し、内部タグをレイヤーパネルから隠します。
         *
         * @returns {void}
         */
        function finalizeGeneratedItems() {
            /* 外枠を残し、4辺に分解していないときだけ元の長方形を表示する。
               それ以外は削除する（非表示のまま残すと、実行のたびに見えない残骸がたまる）
               Show the original only when it is kept and not replaced by the four edges */
            if (keepOuterCheckbox.value && getOuterEdgeScaleValue() === 0) setBaseRectsHidden(false);
            else removeBaseRects();

            /* 内部タグは note だけに残し、レイヤーパネルに出る name はクリアする
               Keep the tag in note only and clear the name shown in the Layers panel */
            clearTagNames(generatedItems);

            doc.selection = null;
        }

        // =========================================
        // ダイアログの組み立てと実行 / Build the dialog and run
        // =========================================

        /* アートボード基準のときは、先に基準の矩形を作る（初期値の計算にも使う）。
           ロックや非表示のレイヤーでは作成できないので、その場合は理由を知らせて終了する
           Create the base rectangle first; a locked or hidden layer makes it impossible, so report and stop */
        if (isArtboardBased) {
            rebuildArtboardBaseRect({ top: 0, right: 0, bottom: 0, left: 0 });
            if (!artboardBaseRect) {
                alert(getLabel(LABELS.alert.baseRectFailed));
                return;
            }
        }

        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 20;
        dialog.margins = 16;

        /* 初めて開くときは表示位置をずらす（生成物が隠れないように）。2回目からは prepareDialogWindow() が前回の位置に戻す
           Shift the dialog on first open so it does not cover the artwork; later opens return to the last position via prepareDialogWindow() */
        dialog.onShow = function () {
            /* 表示前に切り替えた enabled（マージンのパネル・タブなど）を∧∨にも反映 / reflect enabled states set before showing */
            syncStepperStates();
            dialog.location = [
                dialog.location[0] + DIALOG_LAYOUT.dialogOffsetX,
                dialog.location[1] + DIALOG_LAYOUT.dialogOffsetY
            ];
        };

        /* 4タブ：マージン（＋フレーム）／外側エリア（＋タイトルエリア）／内側エリア／画面表示
           Four tabs: margin and frame, outer and title, inner area, display */
        var settingsTabs = dialog.add("tabbedpanel");
        settingsTabs.alignChildren = ["fill", "top"];
        settingsTabs.alignment = ["fill", "top"];
        settingsTabs.margins = [5, 20, 0, 0];

        /* tabbedpanel は内容量に応じて自動で高さが伸びないことがあるため、最低サイズを与える
           A tabbedpanel does not always grow with its content, so give it a minimum size */
        settingsTabs.minimumSize = DIALOG_LAYOUT.tabSize;
        settingsTabs.preferredSize = DIALOG_LAYOUT.tabSize;

        var marginTabColumn = addTabColumn(settingsTabs, LABELS.tab.margin);
        var outerTabColumn = addTabColumn(settingsTabs, LABELS.tab.outer);
        var innerTabColumn = addTabColumn(settingsTabs, LABELS.tab.inner);
        var displayTabColumn = addTabColumn(settingsTabs, LABELS.tab.display);

        buildMarginPanel(marginTabColumn);
        buildFramePanel(marginTabColumn);
        buildOuterPanel(outerTabColumn);
        buildTitlePanel(outerTabColumn);
        buildInnerPanel(innerTabColumn);
        buildDisplayTab(displayTabColumn);

        /* 長方形スタート時（アートボード基準でない場合）は左タブ全体を非表示
           Hide the whole left tab for rectangle-based runs */
        if (!isArtboardBased) {
            marginTabColumn.parent.visible = false;
            marginTabColumn.parent.enabled = false;
            settingsTabs.selection = outerTabColumn.parent;
        }

        // =========================================
        // ボタンエリア（左：プレビュー ／ 右：ボタン） / Footer: preview on the left, buttons on the right
        // =========================================
        var buttonRow = addButtonRow(dialog);

        previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.value = true; /* 最初からプレビューON / preview starts enabled */

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        previewCheckbox.onClick = function () {
            if (previewCheckbox.value) rebuildGeneratedItems(false);
            else clearPreview();
        };

        restoreDialogState();

        // レイアウト確定（tabbedpanel の内容が潰れるのを防ぐ）
        dialog.layout.layout(true);
        dialog.layout.resize();

        /* プレビューがOFFなら描画しない。UIの依存関係だけは collectOptions() でそろえる
           Draw nothing while the preview is off; collectOptions() still syncs the dependent controls */
        if (previewCheckbox.value) rebuildGeneratedItems(false);
        else collectOptions();

        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();
        saveDialogState();

        /* キャンセル時：生成物を削除し、選択パスの見た目と画面表示を元に戻して終了
           On cancel: drop the generated items and restore both the appearance and the view */
        if (dialogResult !== 1) {
            clearPreview();
            restoreSelectedAppearances();
            viewControl.restore();
            return;
        }

        /* 最終生成（この結果はヒストリーに残す） / Final generation, kept in the history */
        rebuildGeneratedItems(true);
        finalizeGeneratedItems();

    })();

})();
