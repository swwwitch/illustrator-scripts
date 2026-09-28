#target illustrator
#targetengine "NewGuideMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログで方向・位置・単位・対象（カンバス／アートボード）を指定してガイドを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/NewGuideMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1085336d7265

### Overview

Creates guides by specifying direction, position, unit, and target (canvas or artboard) in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/NewGuideMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "NewGuideMaker";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/NewGuideMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/NewGuideMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1085336d7265"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ガイドを作成するレイヤー名 / Layer that receives the guides */
    var GUIDE_LAYER_NAME = "_guide";

    /* カンバス端まで届く十分な長さ（Illustrator の最大カンバス 227inch 相当）/ Length long enough to span the canvas (227 inch ≈ Illustrator max canvas) */
    var CANVAS_SPAN_PT = 227 * 72;

    /* プレビュー線の色（CMYK ドキュメント用）/ Preview stroke color for CMYK documents */
    var PREVIEW_COLOR_CMYK = { cyan: 70, magenta: 50, yellow: 0, black: 0 };

    /* プレビュー線の色（CMYK 以外へのフォールバック）/ Preview stroke color for non-CMYK documents */
    var PREVIEW_COLOR_RGB = { red: 74, green: 132, blue: 255 };

    /* プレビュー線の太さ（pt）/ Stroke width of the preview paths (pt) */
    var PREVIEW_STROKE_WIDTH = 1.0;

    /* ガイド化後の線幅（pt）/ Stroke width once converted to a guide (pt) */
    var GUIDE_STROKE_WIDTH = 0.1;

    /* リピートで一度に作れるガイドの上限（桁の打ち間違いで固まるのを防ぐ）/ Cap on repeated guides, so a mistyped digit cannot freeze Illustrator */
    var MAX_REPEAT_COUNT = 1000;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between columns */
    var FIELD_ROW_SPACING  = 6;                /* ラベル・入力欄・単位表記の間隔 / gap inside a labeled field row */
    var UNIT_TEXT_WIDTH    = 34;               /* 数値欄に添える単位表記の幅 / width of the unit label next to a field */

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ガイド作成", en: "Create Guide" }
        },
        target: {
            panelTitle: { ja: "対象", en: "Target" },
            canvas:     { ja: "カンバス", en: "Canvas" },
            artboard:   { ja: "アートボード", en: "Artboard" },
            extension:  { ja: "延長", en: "Extension" }
        },
        direction: {
            panelTitle: { ja: "方向", en: "Direction" },
            horizontal: { ja: "水平方向", en: "Horizontal" },
            vertical:   { ja: "垂直方向", en: "Vertical" },
            position:   { ja: "開始位置", en: "Start Position" }
        },
        layer: {
            panelTitle:  { ja: "作成レイヤー", en: "Target Layer" },
            guideLayer:  { ja: "_guideレイヤー", en: "_guide Layer" },
            activeLayer: { ja: "現在のレイヤー", en: "Current Layer" }
        },
        repeat: {
            panelTitle: { ja: "リピート", en: "Repeat" },
            count:      { ja: "ガイド数", en: "Guide Count" },
            distance:   { ja: "距離", en: "Distance" }
        },
        unit: {
            fieldLabel: { ja: "単位", en: "Unit" }
        },
        tooltip: {
            extension: { ja: "ガイドをアートボードの外側へ伸ばす量（アートボード対象時のみ）", en: "How far to extend guides beyond the artboard (artboard target only)" },
            position:  { ja: "ガイドの開始位置。↑↓で次の整数へ、Shift+↑↓で次の10の倍数へ", en: "Guide start position. Up/Down steps to the next whole number, Shift+Up/Down to the next multiple of 10" },
            count:     { ja: "作成するガイドの本数", en: "Number of guides to create" },
            distance:  { ja: "リピート時のガイドの間隔", en: "Spacing between repeated guides" },
            direction: { ja: "H / V キーでも切り替えできます", en: "Toggle with the H / V keys too" },
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
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            lockedLayer: { ja: "アクティブレイヤーがロックされています。", en: "The active layer is locked." },
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * ラベル末尾に付けるコロンを返す（日本語は全角、英語は半角）
     * @returns {string} コロン記号
     */
    function getUiColon() {
        return (uiLang === "ja") ? "：" : ":";
    }

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
     * 単位ラベルの一覧を返す（ドロップダウン用）
     * @returns {string[]} 単位ラベルの配列
     */
    function getUnitLabels() {
        var labelList = [];
        for (var i = 0; i < UNITS.length; i++) {
            labelList.push(UNITS[i].label);
        }
        return labelList;
    }

    /**
     * ルーラー環境設定の単位インデックスを取得する（= rulerType コード）
     * @returns {number} UNITS の添字（範囲外なら pt の添字）
     */
    function getRulerUnitIndex() {
        return getUnitInfo("rulerType").code;
    }

    /**
     * 値と単位ラベルから pt へ変換する
     * @param {string|number} inputValue - 変換する値（数値以外は0扱い）
     * @param {string} unitLabel - 単位ラベル（"mm" など）
     * @returns {number} pt に変換した値
     */
    function convertToPt(inputValue, unitLabel) {
        var numericValue = Number(inputValue);
        if (isNaN(numericValue)) {
            return 0;
        }
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].label === unitLabel) {
                return numericValue * UNITS[i].pointsPerUnit;
            }
        }
        return numericValue; /* 見つからなければ pt 扱い / Fall back to pt */
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
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
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 左寄せの縦並びグループを生成する（ラジオ列など）
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addLeftAlignedColumn(parentContainer) {
        var createdGroup = parentContainer.add("group");
        createdGroup.orientation = "column";
        createdGroup.alignChildren = ["left", "center"];
        return createdGroup;
    }

    /**
     * 2択のラジオボタン列を生成する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} firstLabel - 1つ目のラベル
     * @param {string} secondLabel - 2つ目のラベル
     * @param {number} selectedIndex - 初期選択（0=1つ目、1=2つ目）
     * @returns {RadioButton[]} [1つ目, 2つ目] のラジオボタン
     */
    function addRadioPair(parentContainer, firstLabel, secondLabel, selectedIndex) {
        var radioColumn = addLeftAlignedColumn(parentContainer);
        var firstRadio = radioColumn.add("radiobutton", undefined, firstLabel);
        var secondRadio = radioColumn.add("radiobutton", undefined, secondLabel);
        ((selectedIndex === 1) ? secondRadio : firstRadio).value = true;
        return [firstRadio, secondRadio];
    }

    /**
     * 幅いっぱいに広がる縦並びカラムを生成する（パネルを積む用）
     * @param {Group} parentRow - 追加先の行グループ
     * @returns {Group} 生成したカラム
     */
    function addSettingsColumn(parentRow) {
        var createdColumn = parentRow.add("group");
        createdColumn.orientation = "column";
        createdColumn.alignChildren = ["fill", "top"];
        createdColumn.spacing = WINDOW_SPACING;
        return createdColumn;
    }

    /**
     * ダイアログウィンドウを生成する（上段に内容、下段にボタンの縦構成）
     * @param {string} windowTitle - ウィンドウタイトル
     * @returns {Window} 生成したダイアログ
     */
    function createDialogWindow(windowTitle) {
        var dialogWindow = new Window("dialog", windowTitle);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.spacing = WINDOW_SPACING;
        dialogWindow.margins = WINDOW_MARGINS;
        return dialogWindow;
    }

    // =========================================
    // ガイド座標と外観 / Guide geometry and appearance
    // =========================================

    /**
     * ガイド1本分の始点・終点座標を返す（カンバスはドキュメント原点基準＝既定のルーラー0点）
     * @param {boolean} isCanvasTarget - カンバス基準なら true、アートボード基準なら false
     * @param {boolean} isHorizontal - 水平ガイドなら true、垂直ガイドなら false
     * @param {number} positionPt - ガイド位置（pt）
     * @param {number} extensionPt - アートボード外への延長量（pt）
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]（カンバス基準では未使用）
     * @returns {number[][]} 始点・終点の座標配列
     */
    function getGuidePathPoints(isCanvasTarget, isHorizontal, positionPt, extensionPt, artboardRect) {
        if (isCanvasTarget) {
            /* Y は上方向が正なので、下向きの位置は減算 / Y is up, so a downward position subtracts */
            return isHorizontal
                ? [[-CANVAS_SPAN_PT, -positionPt], [CANVAS_SPAN_PT, -positionPt]]
                : [[positionPt, CANVAS_SPAN_PT], [positionPt, -CANVAS_SPAN_PT]];
        }
        var artboardLeft = artboardRect[0];
        var artboardTop = artboardRect[1];
        var artboardRight = artboardRect[2];
        var artboardBottom = artboardRect[3];
        return isHorizontal
            ? [[artboardLeft - extensionPt, artboardTop - positionPt], [artboardRight + extensionPt, artboardTop - positionPt]]
            : [[artboardLeft + positionPt, artboardTop + extensionPt], [artboardLeft + positionPt, artboardBottom - extensionPt]];
    }

    /**
     * プレビュー線に使う色を生成する
     * @param {DocumentColorSpace} docColorSpace - ドキュメントのカラースペース
     * @returns {CMYKColor|RGBColor} プレビュー線の色
     */
    function createPreviewColor(docColorSpace) {
        if (docColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = PREVIEW_COLOR_CMYK.cyan;
            cmykColor.magenta = PREVIEW_COLOR_CMYK.magenta;
            cmykColor.yellow = PREVIEW_COLOR_CMYK.yellow;
            cmykColor.black = PREVIEW_COLOR_CMYK.black;
            return cmykColor;
        }
        /* CMYK 以外は青の RGB にフォールバック / Fall back to a blue RGB for non-CMYK modes */
        var rgbColor = new RGBColor();
        rgbColor.red = PREVIEW_COLOR_RGB.red;
        rgbColor.green = PREVIEW_COLOR_RGB.green;
        rgbColor.blue = PREVIEW_COLOR_RGB.blue;
        return rgbColor;
    }

    /**
     * プレビュー線の見た目を設定する
     * @param {PathItem} previewPath - 対象パス
     * @param {CMYKColor|RGBColor} previewColor - 線の色
     * @returns {void}
     */
    function stylePreviewPath(previewPath, previewColor) {
        previewPath.stroked = true;
        previewPath.filled = false;
        previewPath.strokeWidth = PREVIEW_STROKE_WIDTH;
        previewPath.strokeColor = previewColor;
        previewPath.guides = false; /* プレビューはガイド化しない / Preview is not a guide */
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {object} PreviewSettings
     * @property {boolean} isCanvasTarget - カンバス基準なら true
     * @property {boolean} isHorizontal - 水平ガイドなら true
     * @property {number} positionPt - 1本目のガイド位置（pt）
     * @property {number} extensionPt - アートボード外への延長量（pt）
     * @property {number[]} artboardRect - アートボードの矩形（カンバス基準では null）
     * @property {number} repeatCount - 作成するガイドの本数
     * @property {number} repeatDistancePt - リピート間隔（pt）
     */

    /**
     * ガイド作成ダイアログを構築する
     * @returns {Window} 構築済みのダイアログ
     */
    function createGuideDialog() {
        var doc = app.activeDocument;

        /* 単位オプションと現在単位のインデックス / Unit options and current unit index */
        var unitOptions = getUnitLabels();
        var currentUnitIndex = getRulerUnitIndex();
        var unitSuffixTexts = []; /* 各数値欄の単位表記（共有ドロップダウンに追従）/ Per-field unit labels (follow the shared dropdown) */

        /* ガイド用レイヤーは使用時のみ遅延生成 / The guide layer is created lazily, only when used */
        var guideLayer = null;
        var guideLayerWasCreated = false; /* このスクリプトで新規作成したか / Whether this run created the layer */
        var guideLayerWasLocked = false;  /* 元のロック状態 / The layer's lock state before this run */
        var guidesCommitted = false;      /* OKでガイド化を確定したか / Whether OK committed the guides */

        /* ロック警告は選択ごとに1回だけ出す / Warn about a locked layer only once per selection */
        var lockedLayerAlerted = false;

        /* 現在描画中のプレビュー線（常に配列で保持）/ Currently drawn preview paths (always an array) */
        var activePreviewPaths = null;

        /* ダイアログ内で共有する主要コントロール / Controls shared across the dialog's handlers */
        var canvasRadio, artboardRadio, extensionRow, extensionInput;
        var guideLayerRadio, activeLayerRadio;
        var horizontalRadio, verticalRadio, positionInput;
        var repeatCountInput, repeatDistanceInput;
        var unitDropdown;

        /**
         * ガイド用レイヤーを取得する（なければ作成）
         * @returns {Layer} ガイド用レイヤー
         */
        function ensureGuideLayer() {
            if (guideLayer) {
                return guideLayer;
            }
            var docLayers = doc.layers;
            for (var i = 0; i < docLayers.length; i++) {
                if (docLayers[i].name === GUIDE_LAYER_NAME) {
                    guideLayer = docLayers[i];
                    guideLayerWasLocked = guideLayer.locked;
                    return guideLayer;
                }
            }
            guideLayer = docLayers.add();
            guideLayer.name = GUIDE_LAYER_NAME;
            guideLayerWasCreated = true;
            return guideLayer;
        }

        /**
         * ガイド化せずに閉じたとき、ガイド用レイヤーを元の状態に戻す
         * （このスクリプトで作った空レイヤーは削除、既存レイヤーはロック状態を復元）
         * @returns {void}
         */
        function restoreGuideLayer() {
            if (!guideLayer) return;
            guideLayer.locked = false;
            if (guideLayerWasCreated) {
                /* 空のまま残さない（最後の1枚は削除できないので残す）/ Do not leave an empty layer behind (the last layer cannot be removed) */
                if (guideLayer.pageItems.length === 0 && doc.layers.length > 1) {
                    guideLayer.remove();
                }
            } else {
                guideLayer.locked = guideLayerWasLocked;
            }
            guideLayer = null;
        }

        /**
         * プレビュー線を削除する（モーダル中はドキュメント編集不可なのでパスは常に有効）
         * @returns {void}
         */
        function removePreviewPaths() {
            if (!activePreviewPaths) return;
            for (var i = 0; i < activePreviewPaths.length; i++) {
                var previewPath = activePreviewPaths[i];
                if (previewPath && !previewPath.locked && previewPath.layer && !previewPath.layer.locked) {
                    previewPath.remove();
                }
            }
            activePreviewPaths = null;
        }

        /**
         * プレビューの描画先レイヤーを決定する
         * @returns {Layer|null} 描画先レイヤー（アクティブレイヤーがロック中なら null）
         */
        function resolvePreviewLayer() {
            if (guideLayerRadio.value) {
                var targetGuideLayer = ensureGuideLayer();
                targetGuideLayer.locked = false;
                lockedLayerAlerted = false;
                return targetGuideLayer;
            }
            var activeLayer = doc.activeLayer;
            if (activeLayer.locked) {
                /* プレビューは打鍵ごとに走るので、警告は選択ごとに1回だけ / The preview runs on every keystroke, so warn only once per selection */
                if (!lockedLayerAlerted) {
                    lockedLayerAlerted = true;
                    alert(getLabel('alert.lockedLayer'));
                }
                return null;
            }
            lockedLayerAlerted = false;
            return activeLayer;
        }

        /**
         * 入力欄の内容をプレビュー用の設定値に読み取る
         * @returns {PreviewSettings|null} 設定値（数値として読めない欄があれば null）
         */
        function readPreviewSettings() {
            /* 単位は全フィールド共通 / A single shared unit for all fields */
            var unitLabel = unitDropdown.selection.text;

            var positionValue = parseFloat(positionInput.text);
            if (isNaN(positionValue)) {
                return null;
            }

            var repeatCount = parseInt(repeatCountInput.text, 10);
            if (isNaN(repeatCount) || repeatCount < 1) repeatCount = 1;
            /* 桁の打ち間違いで大量生成しないよう上限で丸める / Clamp so a mistyped digit cannot spawn a huge batch */
            if (repeatCount > MAX_REPEAT_COUNT) repeatCount = MAX_REPEAT_COUNT;
            var repeatDistancePt = convertToPt(repeatDistanceInput.text, unitLabel);
            /* 距離0以下なら重複を避けて1本に / Avoid overlapping guides when distance is 0 or less */
            if (repeatDistancePt <= 0) {
                repeatDistancePt = 0;
                repeatCount = 1;
            }

            /* アートボード対象時の延長量と矩形を一度だけ算出 / Compute extension amount and rect once when targeting the artboard */
            var isCanvasTarget = canvasRadio.value;
            var extensionPt = 0;
            var artboardRect = null;
            if (!isCanvasTarget) {
                var extensionValue = parseFloat(extensionInput.text);
                if (isNaN(extensionValue)) {
                    return null;
                }
                extensionPt = convertToPt(extensionValue, unitLabel);
                artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            }

            return {
                isCanvasTarget: isCanvasTarget,
                isHorizontal: horizontalRadio.value,
                positionPt: convertToPt(positionValue, unitLabel),
                extensionPt: extensionPt,
                artboardRect: artboardRect,
                repeatCount: repeatCount,
                repeatDistancePt: repeatDistancePt
            };
        }

        /**
         * 設定値どおりにプレビュー線を描画する
         * @param {PreviewSettings} settings - 入力欄から読み取った設定値
         * @param {Layer} targetLayer - 描画先レイヤー
         * @returns {PathItem[]} 描画したプレビュー線
         */
        function drawPreviewPaths(settings, targetLayer) {
            var previewColor = createPreviewColor(doc.documentColorSpace);
            var drawnPaths = [];
            for (var i = 0; i < settings.repeatCount; i++) {
                var currentPositionPt = settings.positionPt + i * settings.repeatDistancePt;
                var previewPath = targetLayer.pathItems.add();
                previewPath.setEntirePath(getGuidePathPoints(
                    settings.isCanvasTarget,
                    settings.isHorizontal,
                    currentPositionPt,
                    settings.extensionPt,
                    settings.artboardRect
                ));
                stylePreviewPath(previewPath, previewColor);
                drawnPaths.push(previewPath);
            }
            return drawnPaths;
        }

        /**
         * 現在の入力値でプレビュー線を描き直す
         * @returns {void}
         */
        function drawPreview() {
            removePreviewPaths();
            var settings = readPreviewSettings();
            if (!settings) {
                return;
            }
            var targetLayer = resolvePreviewLayer();
            if (!targetLayer) {
                return;
            }
            activePreviewPaths = drawPreviewPaths(settings, targetLayer);
            app.redraw();
        }

        /**
         * プレビュー線をガイドに変換する
         * @returns {void}
         */
        function convertPreviewToGuides() {
            if (!activePreviewPaths) return;
            /* ガイド化のため対象レイヤーのロックを一時解除 / Temporarily unlock so paths can be converted */
            if (guideLayer && guideLayer.locked) {
                guideLayer.locked = false;
            }
            for (var i = 0; i < activePreviewPaths.length; i++) {
                var previewPath = activePreviewPaths[i];
                if (previewPath && previewPath.layer && !previewPath.layer.locked) {
                    previewPath.guides = true;
                    previewPath.strokeWidth = GUIDE_STROKE_WIDTH;
                }
            }
        }

        /**
         * ∧∨付きの数値入力欄を生成する（↑↓キーも∧∨と同じ処理で増減し、逐次プレビューする）
         * @param {Group} parentRow - 追加先の行グループ
         * @param {string} defaultValue - 初期値
         * @param {number} [widthInChars] - 入力欄の幅（文字数、省略時は2）
         * @param {Object} [stepOptions] - addStepper() に渡す増減の設定（省略時は1ずつ・下限なし）
         * @returns {EditText} 生成した入力欄
         */
        function addNumberField(parentRow, defaultValue, widthInChars, stepOptions) {
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var stepperInputGroup = parentRow.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;
            var fieldStepOptions = stepOptions || { step: 1 };
            fieldStepOptions.onStep = drawPreview;
            var numberField;
            var stepperGroup = addStepper(stepperInputGroup, function () { return numberField; }, fieldStepOptions);
            numberField = stepperInputGroup.add('edittext {characters: ' + ((typeof widthInChars === "number") ? widthInChars : 2) + '}');
            numberField.text = defaultValue;
            bindSteppedArrowKeys(numberField, stepperGroup);
            numberField.addEventListener("changing", drawPreview);
            return numberField;
        }

        /**
         * 数値欄の右に単位表記を追加する（共有ドロップダウンに追従）
         * @param {Group} parentRow - 追加先の行グループ
         * @returns {StaticText} 生成した単位表記
         */
        function addUnitSuffixText(parentRow) {
            var unitText = parentRow.add("statictext", undefined, unitOptions[currentUnitIndex]);
            unitText.preferredSize.width = UNIT_TEXT_WIDTH;
            unitSuffixTexts.push(unitText);
            return unitText;
        }

        /**
         * ラベル付きの数値欄を生成する
         * @param {Group|Panel} parentContainer - 追加先
         * @param {string} labelText - ラベル文字列（コロンは自動付与）
         * @param {string} defaultValue - 初期値
         * @param {boolean} showUnitSuffix - 単位表記を添えるかどうか
         * @param {string} [helpTipText] - ツールチップ
         * @param {number} [widthInChars] - 入力欄の幅（文字数）
         * @param {Object} [stepOptions] - ∧∨の増減の設定（省略時は1ずつ・下限なし）
         * @returns {{row: Group, input: EditText}} 行グループと入力欄
         */
        function addLabeledField(parentContainer, labelText, defaultValue, showUnitSuffix, helpTipText, widthInChars, stepOptions) {
            var fieldRow = parentContainer.add("group");
            setupRow(fieldRow, "left", FIELD_ROW_SPACING);
            var fieldLabel = fieldRow.add("statictext", undefined, labelText + getUiColon());
            var numberField = addNumberField(fieldRow, defaultValue, widthInChars, stepOptions);
            if (showUnitSuffix) {
                addUnitSuffixText(fieldRow);
            }
            if (helpTipText) {
                fieldLabel.helpTip = helpTipText;
                numberField.helpTip = helpTipText;
            }
            return { row: fieldRow, input: numberField };
        }

        /**
         * 対象（カンバス／アートボード）パネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildTargetPanel(parentColumn) {
            var targetPanel = addPanel(parentColumn, getLabel('target.panelTitle'));
            var targetRadios = addRadioPair(targetPanel, getLabel('target.canvas'), getLabel('target.artboard'), 1);
            canvasRadio = targetRadios[0];
            artboardRadio = targetRadios[1];

            /* 延長：ガイドをアートボード外へ伸ばす量 / Extension: how far to extend guides beyond the artboard */
            var extensionField = addLabeledField(targetPanel, getLabel('target.extension'), "0", true, getLabel('tooltip.extension'));
            extensionRow = extensionField.row;
            extensionInput = extensionField.input;
        }

        /**
         * 作成レイヤーパネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildLayerPanel(parentColumn) {
            var layerPanel = addPanel(parentColumn, getLabel('layer.panelTitle'));
            var layerRadios = addRadioPair(layerPanel, getLabel('layer.guideLayer'), getLabel('layer.activeLayer'), 0);
            guideLayerRadio = layerRadios[0];
            activeLayerRadio = layerRadios[1];
        }

        /**
         * 方向パネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildDirectionPanel(parentColumn) {
            var directionPanel = addPanel(parentColumn, getLabel('direction.panelTitle'));
            directionPanel.helpTip = getLabel('tooltip.direction');
            var directionRadios = addRadioPair(directionPanel, getLabel('direction.horizontal'), getLabel('direction.vertical'), 0);
            horizontalRadio = directionRadios[0];
            verticalRadio = directionRadios[1];
            positionInput = addLabeledField(directionPanel, getLabel('direction.position'), "0", true, getLabel('tooltip.position')).input;
        }

        /**
         * リピートパネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildRepeatPanel(parentColumn) {
            var repeatPanel = addPanel(parentColumn, getLabel('repeat.panelTitle'));
            var repeatColumn = addLeftAlignedColumn(repeatPanel);
            /* ガイド数は1〜上限の整数 / guide count: integers from 1 to the cap */
            repeatCountInput = addLabeledField(repeatColumn, getLabel('repeat.count'), "1", false, getLabel('tooltip.count'), undefined,
                { step: 1, min: 1, max: MAX_REPEAT_COUNT, integer: true }).input;
            repeatDistanceInput = addLabeledField(repeatColumn, getLabel('repeat.distance'), "0", true, getLabel('tooltip.distance'), 3).input;
        }

        var dialog = createDialogWindow(getLabel('dialog.title') + ' ' + SCRIPT_VERSION);

        /* 上段：2カラムを横並びに収める行 / Top area: a row holding the two columns */
        var columnsRow = dialog.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["fill", "top"];
        columnsRow.spacing = COLUMN_SPACING;

        /* 左カラム：対象・作成レイヤー / Left column: target, layer */
        var leftColumn = addSettingsColumn(columnsRow);
        buildTargetPanel(leftColumn);
        buildLayerPanel(leftColumn);

        /* 右カラム：方向・リピート / Right column: direction, repeat */
        var rightColumn = addSettingsColumn(columnsRow);
        buildDirectionPanel(rightColumn);
        buildRepeatPanel(rightColumn);

        /* 下部のボタン行（左＝単位、右＝キャンセル＋OK）/ Bottom button row (left: unit, right: Cancel + OK) */
        var buttonRow = addButtonRow(dialog);
        buttonRow.leftGroup.add("statictext", undefined, getLabel('unit.fieldLabel') + getUiColon());
        unitDropdown = buttonRow.leftGroup.add("dropdownlist", undefined, unitOptions);
        unitDropdown.selection = currentUnitIndex;
        unitDropdown.onChange = function() {
            /* 各数値欄の単位表記を更新 / Update the per-field unit labels */
            var selectedUnitLabel = unitDropdown.selection.text;
            for (var i = 0; i < unitSuffixTexts.length; i++) {
                unitSuffixTexts[i].text = selectedUnitLabel;
            }
            drawPreview();
        };

        /* キャンセルは既定動作で閉じ、後片付けは dialog.onClose が行う / Cancel closes by default; cleanup happens in dialog.onClose */
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });

        /* OKボタン：プレビュー線をガイド化してレイヤーを再ロック / OK: convert the preview to guides, then re-lock the layer */
        btnOK.onClick = function() {
            convertPreviewToGuides();
            if (guideLayer) {
                guideLayer.locked = true;
            }
            activePreviewPaths = null;
            guidesCommitted = true;
            dialog.close();
        };

        /**
         * 対象の選択状態をUIに反映してプレビューを更新する
         * @returns {void}
         */
        function syncTargetState() {
            /* カンバス対象では延長が効かないので行をディム / The extension has no effect on the canvas, so dim the row */
            extensionRow.enabled = artboardRadio.value;
            redrawSteppersIn(extensionRow);
            drawPreview();
        }
        canvasRadio.onClick = syncTargetState;
        artboardRadio.onClick = syncTargetState;

        /* イベント：その他はプレビュー再描画のみ / Events: others just redraw the preview */
        horizontalRadio.onClick = drawPreview;
        verticalRadio.onClick = drawPreview;
        guideLayerRadio.onClick = drawPreview;
        activeLayerRadio.onClick = drawPreview;

        /* H/Vキーで方向切り替え。数値欄に文字が入らないようキャプチャフェーズで受ける / Switch direction with H/V; capture phase keeps the letter out of the numeric fields */
        dialog.addEventListener("keydown", function(event) {
            var pressedKey = (event.keyName || "").toUpperCase();
            if (pressedKey !== "H" && pressedKey !== "V") return;
            horizontalRadio.value = (pressedKey === "H");
            verticalRadio.value = !horizontalRadio.value;
            drawPreview();
            event.preventDefault();
        }, true);

        /* 初期状態を反映し、そのまま初回プレビューを描画 / Apply the initial state and draw the first preview */
        syncTargetState();

        /* ダイアログ表示時に「開始位置」入力欄へフォーカス / Focus the position input on dialog show */
        positionInput.active = true;

        /* ダイアログを閉じたら後片付け（キャンセル・ESCも含む）/ Clean up on close (Cancel and ESC included) */
        dialog.onClose = function() {
            removePreviewPaths();
            if (!guidesCommitted) {
                /* ガイド化せずに閉じたので、レイヤーを元の状態へ / Closed without committing, so put the layer back */
                restoreGuideLayer();
            }
        };
        return dialog;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * ガイド作成ダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }
        var guideDialog = createGuideDialog();
        prepareDialogWindow(guideDialog, SCRIPT_NAME);
        guideDialog.show();
    }

    main();

})();
