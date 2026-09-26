#targetengine "RulesBetweenObjects"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクト（図形／テキスト）を上から順に並べ、その間に水平の罫線を描画します。
入力単位は環境設定の「線」に追従し、［延長］で罫線を左右方向に伸縮できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawLinesBetween.md

### Overview

Sorts the selected objects (shapes or text) from top to bottom and draws a horizontal rule between each pair.
Input units follow the stroke-units preference, and Extend stretches or shrinks the rules horizontally.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawLinesBetween.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DrawLinesBetween";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawLinesBetween.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawLinesBetween.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /*
      DrawLinesBetween.jsx

      選択したオブジェクト（図形／テキスト）を上から順に並べ、間に水平の罫線（ケイ線）を描画します。
      / Sort selected objects (shapes/text) from top to bottom and draw horizontal rules between them.

      主なポイント / Key features:
      - 入力単位は「線（strokeUnits）」の環境設定に追従（表示ラベル＆内部pt換算）
        / Follows "strokeUnits" preference (UI label + internal pt conversion)
      - 「延長」で左右方向に罫線を伸縮（+で延長、-で短縮）
        / "Extension" expands/contracts the rule length horizontally (+ extends, - shrinks)
      - 線端（なし／丸型／突出）を選択可能
        / Select line cap (Butt / Round / Projecting)
      - ケイ線の長さ：共通（全体幅）／オブジェクトに合わせる（行ペア幅）
        / Rule length: Common (overall width) / Match Objects (per-row-pair width)
      - テキストは複製→アウトライン化した境界で計算し、終了時に一時オブジェクトを削除
        / Text is measured via duplicate→outline bounds, then temporary items are removed
      - 左右に並ぶ要素は同一行として束ねて扱う（行グルーピング）
        / Horizontally aligned items are treated as one row (row grouping)
      - 実行後、作成した罫線を選択状態に
        / Select created rules after execution
      - 最後に使った設定を保存して次回起動時に復元（線幅／延長／線端など）
        / Persist last-used settings and restore on next run (stroke/extension/cap etc.)
    */

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_LINE_WEIGHT = 0.25; /* 線幅の既定値（pt）/ default stroke weight in pt */
    var DEFAULT_EXTENSION   = 0;    /* 延長の既定値（pt）。左右にどれだけ伸ばすか / default extension on each side in pt */
    var ROW_OVERLAP_RATIO   = 0.5;  /* 同じ行とみなす縦方向の重なり率（0.0–1.0）/ vertical overlap ratio for one row */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OFFSET_X    = 300;              /* ダイアログの表示位置のずらし量 / dialog position offset */
    var DIALOG_OFFSET_Y    = 0;
    var DIALOG_OPACITY     = 0.98;             /* ダイアログの不透明度 / dialog opacity */
    var PANEL_MARGINS      = [15, 20, 15, 10]; /* パネル余白 [左,上,右,下] / panel margins */
    var NUMBER_FIELD_CHARS = 6;                /* 数値入力欄の幅（文字数）/ number field width in characters */

    /**
     * ダイアログを表示するときに位置をずらす
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    /**
     * ダイアログの不透明度を設定する
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} opacityValue - 不透明度（0〜1）
     * @returns {void}
     */
    function setDialogOpacity(targetDialog, opacityValue) {
        try {
            targetDialog.opacity = opacityValue;
        } catch (e) { /* 環境によっては opacity を持たない / opacity is not supported in some environments */ }
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
     * 実行環境のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクト間に罫線", en: "Rules Between Objects" }
        },
        panel: {
            lineCap:    { ja: "線端", en: "Line Cap" },
            ruleLength: { ja: "ケイ線の長さ", en: "Rule Length" }
        },
        fieldLabel: {
            lineWeight: { ja: "線幅", en: "Stroke" },
            extension:  { ja: "延長", en: "Extension" }
        },
        radio: {
            capButt:          { ja: "なし", en: "Butt" },
            capRound:         { ja: "丸型線端", en: "Round" },
            capProjecting:    { ja: "突出線端", en: "Projecting" },
            ruleLengthObject: { ja: "オブジェクトに合わせる", en: "Match Objects" },
            ruleLengthCommon: { ja: "共通", en: "Common" }
        },
        tooltip: {
            lineWeight:       { ja: "ケイ線の太さです。", en: "Weight of the rules." },
            extension:        { ja: "オブジェクトの端からケイ線を離す距離です。", en: "How far the rules sit from the edge of the objects." },
            capButt:          { ja: "線の端を切りっぱなしにします。", en: "Leaves the line ends flat." },
            capRound:         { ja: "線の端を丸くします。", en: "Rounds the line ends." },
            capProjecting:    { ja: "線の端を太さの半分だけ延ばします。", en: "Extends the line ends by half the weight." },
            ruleLengthObject: {
                ja: "ケイ線の長さを、上下のオブジェクトの幅に合わせます。",
                en: "Matches each rule to the width of the objects it sits between."
            },
            ruleLengthCommon: { ja: "すべてのケイ線を同じ長さにそろえます。", en: "Gives every rule the same length." },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:    { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger:  { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            needPositiveStroke:  { ja: "線幅は正の数値を入力してください。", en: "Please enter a positive number for stroke." },
            needNumberExtension: { ja: "延長は数値を入力してください。", en: "Please enter a number for extension." },
            noDocument:          { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文言を取り出す
     * @param {string} labelPath - "category.key" 形式のパス
     * @returns {string} 表示用の文言（見つからないときはパスそのもの）
     */
    function getLabel(labelPath) {
        var labelNode = LABELS;
        var pathKeys = labelPath.split(".");
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
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
    // 前回値の記憶 / Session settings
    // =========================================

    var SETTINGS_KEY = "RulesBetweenObjectsSettings";

    /**
     * 前回の設定を読み込む
     * @returns {{lineWeightPt: number, extensionPt: number, capIndex: number}} 保存値（無いときは null）
     */
    function loadSettings() {
        try {
            /* 保存値が無いと例外になる / throws when nothing is saved */
            var settingsDescriptor = app.getCustomOptions(SETTINGS_KEY);
            return {
                lineWeightPt: settingsDescriptor.getReal(stringIDToTypeID("lineWeightPt")),
                extensionPt: settingsDescriptor.getReal(stringIDToTypeID("marginPt")),
                capIndex: settingsDescriptor.getInteger(stringIDToTypeID("capIndex"))
            };
        } catch (e) {
            return null;
        }
    }

    /**
     * 最後に使った設定を保存する（長さは pt で持つ）
     * @param {number} lineWeightPt - 線幅（pt）
     * @param {number} extensionPt - 延長（pt）
     * @param {number} capIndex - 線端（0:なし / 1:丸型 / 2:突出）
     * @returns {void}
     */
    function saveSettings(lineWeightPt, extensionPt, capIndex) {
        try {
            var settingsDescriptor = new ActionDescriptor();
            settingsDescriptor.putReal(stringIDToTypeID("lineWeightPt"), lineWeightPt);
            settingsDescriptor.putReal(stringIDToTypeID("marginPt"), extensionPt);
            settingsDescriptor.putInteger(stringIDToTypeID("capIndex"), capIndex);
            app.putCustomOptions(SETTINGS_KEY, settingsDescriptor, true);
        } catch (e) { }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値を入力欄向けに丸める（小数第2位まで）
     * @param {number} numberValue - 表示したい数値
     * @returns {string} 整形後の文字列
     */
    function formatNumberForUI(numberValue) {
        var roundedValue = Math.round(numberValue * 100) / 100;
        /* -0 を防ぐ / avoid -0 */
        if (Math.abs(roundedValue) < 0.000001) roundedValue = 0;
        return String(roundedValue);
    }

    /**
     * 「項目名＋数値欄＋（単位）」の行を追加する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {number} initialValue - 初期値（現在の線の単位）
     * @param {string} unitLabel - 単位の表示
     * @param {string} tooltipPath - tooltip のラベルのパス
     * @param {Object} stepOptions - ∧∨と↑↓キーの増減設定（min など。addStepper() に渡す）
     * @returns {EditText} 追加した数値欄
     */
    function addNumberRow(parentContainer, labelPath, initialValue, unitLabel, tooltipPath, stepOptions) {
        var numberRow = parentContainer.add("group");
        numberRow.add("statictext", undefined, getLabel(labelPath));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = numberRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, formatNumberForUI(initialValue));
        numberInput.helpTip = getLabel(tooltipPath);
        numberInput.characters = NUMBER_FIELD_CHARS;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        numberRow.add("statictext", undefined, "(" + unitLabel + ")");
        return numberInput;
    }

    /**
     * ラジオボタンを縦に並べるパネルを追加する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {string} titlePath - パネル見出しのラベルのパス
     * @returns {Group} ラジオボタンを入れるグループ
     */
    function addRadioPanel(parentContainer, titlePath) {
        var radioPanel = parentContainer.add("panel", undefined, getLabel(titlePath));
        radioPanel.orientation = "column";
        radioPanel.alignChildren = ["left", "top"];
        radioPanel.margins = PANEL_MARGINS;

        var radioGroup = radioPanel.add("group");
        radioGroup.orientation = "column";
        radioGroup.alignChildren = ["left", "center"];
        return radioGroup;
    }

    /**
     * tooltip 付きのラジオボタンを追加する
     * @param {Group} parentContainer - 追加先のグループ
     * @param {string} labelKey - radio と tooltip に共通のキー
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadioButton(parentContainer, labelKey) {
        var radioButton = parentContainer.add("radiobutton", undefined, getLabel("radio." + labelKey));
        radioButton.helpTip = getLabel("tooltip." + labelKey);
        return radioButton;
    }

    /**
     * ダイアログを組み立てる
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @param {{lineWeight: number, extension: number, capIndex: number}} initialValues - 初期値（長さは現在の線の単位）
     * @returns {Object} ダイアログと入力の読み取りに使うコントロール
     */
    function buildRulesDialog(strokeUnit, initialValues) {
        var rulesDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setDialogOpacity(rulesDialog, DIALOG_OPACITY);
        shiftDialogPosition(rulesDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        rulesDialog.orientation = "column";
        rulesDialog.alignChildren = ["fill", "top"];

        var lineWeightInput = addNumberRow(rulesDialog, "fieldLabel.lineWeight", initialValues.lineWeight,
            strokeUnit.label, "tooltip.lineWeight", { min: 0 });
        var extensionInput = addNumberRow(rulesDialog, "fieldLabel.extension", initialValues.extension,
            strokeUnit.label, "tooltip.extension", {}); /* マイナスで短縮するので下限なし / negative shortens the rule */

        /* 線端（既定は「なし」、前回値があれば反映）/ Line cap: Butt by default, or the saved one */
        var capRadioGroup = addRadioPanel(rulesDialog, "panel.lineCap");
        var capButtRadio = addRadioButton(capRadioGroup, "capButt");
        var capRoundRadio = addRadioButton(capRadioGroup, "capRound");
        var capProjectingRadio = addRadioButton(capRadioGroup, "capProjecting");
        capButtRadio.value = true;
        if (initialValues.capIndex === 1) capRoundRadio.value = true;
        else if (initialValues.capIndex === 2) capProjectingRadio.value = true;

        /* ケイ線の長さ（既定は「共通」）/ Rule length: Common by default */
        var lengthRadioGroup = addRadioPanel(rulesDialog, "panel.ruleLength");
        var lengthObjectRadio = addRadioButton(lengthRadioGroup, "ruleLengthObject");
        var lengthCommonRadio = addRadioButton(lengthRadioGroup, "ruleLengthCommon");
        lengthCommonRadio.value = true;

        var btnRowGroup = rulesDialog.add("group");
        btnRowGroup.alignment = ["right", "center"];
        btnRowGroup.orientation = "row";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnOK.active = true;              /* Enterキー＝OK / Enter = OK */
        rulesDialog.defaultElement = btnOK; /* 環境依存対策 / for environments that ignore active */

        return {
            dialog: rulesDialog,
            lineWeightInput: lineWeightInput,
            extensionInput: extensionInput,
            capRoundRadio: capRoundRadio,
            capProjectingRadio: capProjectingRadio,
            lengthObjectRadio: lengthObjectRadio
        };
    }

    /**
     * ダイアログの入力を検証して設定にまとめる
     * @param {Object} dialogControls - buildRulesDialog() の戻り値
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @returns {Object} 線幅・延長（pt）・線端・長さの基準（入力が不正なときは null）
     */
    function readRuleSettings(dialogControls, strokeUnit) {
        var lineWeightValue = Number(dialogControls.lineWeightInput.text);
        if (isNaN(lineWeightValue) || lineWeightValue <= 0) {
            alert(getLabel("alert.needPositiveStroke"));
            return null;
        }

        var extensionValue = Number(dialogControls.extensionInput.text);
        if (isNaN(extensionValue)) {
            alert(getLabel("alert.needNumberExtension"));
            return null;
        }

        var capIndex = 0;
        if (dialogControls.capRoundRadio.value) capIndex = 1;
        else if (dialogControls.capProjectingRadio.value) capIndex = 2;

        return {
            lineWeight: lineWeightValue * strokeUnit.pointsPerUnit,
            extension: extensionValue * strokeUnit.pointsPerUnit,
            capIndex: capIndex,
            lineCap: [StrokeCap.BUTTENDCAP, StrokeCap.ROUNDENDCAP, StrokeCap.PROJECTINGENDCAP][capIndex],
            ruleLengthMode: dialogControls.lengthObjectRadio.value ? "object" : "common"
        };
    }

    /**
     * ダイアログを表示し、確定した設定を保存して返す
     * @returns {Object} readRuleSettings() の戻り値（キャンセル・入力エラーのときは null）
     */
    function showRulesDialog() {
        var savedSettings = loadSettings();
        var strokeUnit = getUnitInfo("strokeUnits");

        /* 保存値（pt）があれば優先し、現在の線の単位で表示する / Show saved pt values in the current stroke unit */
        var initialValues = {
            lineWeight: (savedSettings ? savedSettings.lineWeightPt : DEFAULT_LINE_WEIGHT) / strokeUnit.pointsPerUnit,
            extension: (savedSettings ? savedSettings.extensionPt : DEFAULT_EXTENSION) / strokeUnit.pointsPerUnit,
            capIndex: savedSettings ? savedSettings.capIndex : 0
        };

        var dialogControls = buildRulesDialog(strokeUnit, initialValues);
        if (dialogControls.dialog.show() !== 1) return null;

        var ruleSettings = readRuleSettings(dialogControls, strokeUnit);
        if (ruleSettings) saveSettings(ruleSettings.lineWeight, ruleSettings.extension, ruleSettings.capIndex);
        return ruleSettings;
    }

    // =========================================
    // 行の検出 / Row detection
    // =========================================

    /**
     * グループ内のテキストをアウトライン化する（境界を安定させるため。複製に対して使う）
     * @param {GroupItem} targetGroup - 対象のグループ
     * @returns {void}
     */
    function outlineTextFramesInContainer(targetGroup) {
        if (!targetGroup || !targetGroup.textFrames || targetGroup.textFrames.length === 0) return;

        /* textFrames はライブコレクションになり得るので、いったん配列化 / snapshot the live collection */
        var textFrameList = [];
        for (var i = 0; i < targetGroup.textFrames.length; i++) {
            textFrameList.push(targetGroup.textFrames[i]);
        }

        for (var j = 0; j < textFrameList.length; j++) {
            try {
                /* createOutline() は元のテキストを消費してアウトラインのグループを返す / consumes the text frame */
                var outlineGroup = textFrameList[j].createOutline();
                outlineGroup.hidden = true;
            } catch (e) { /* 変換できないテキストは無視 / skip frames that cannot be outlined */ }
        }
    }

    /**
     * 境界の計測に使うオブジェクトを返す
     * テキストとグループは複製してアウトライン化したもの（非表示）を返し、tempItems に控える。
     * @param {PageItem} sourceItem - 選択中のオブジェクト
     * @param {PageItem[]} tempItems - 一時オブジェクトの控え（追加される）
     * @returns {PageItem} 計測用のオブジェクト（複製できないときは元のオブジェクト）
     */
    function makeBoundsProxy(sourceItem, tempItems) {
        if (!sourceItem) return sourceItem;

        if (sourceItem.typename === "TextFrame") {
            try {
                /* createOutline() は複製を消費してアウトラインのグループを返す / consumes the duplicate */
                var outlineGroup = sourceItem.duplicate().createOutline();
                try { outlineGroup.hidden = true; } catch (e) { }
                tempItems.push(outlineGroup);
                return outlineGroup;
            } catch (e) {
                /* 変換できない場合は元のテキストを使う / fall back to the original text */
                return sourceItem;
            }
        }

        if (sourceItem.typename === "GroupItem") {
            try {
                var groupDuplicate = sourceItem.duplicate();
                /* グループ内のテキストもアウトライン化してから境界を見る / outline nested text too */
                outlineTextFramesInContainer(groupDuplicate);
                try { groupDuplicate.hidden = true; } catch (e) { }
                tempItems.push(groupDuplicate);
                return groupDuplicate;
            } catch (e) {
                return sourceItem;
            }
        }

        return sourceItem;
    }

    /**
     * 2つのオブジェクトの縦方向の重なり率を返す（低いほうの高さに対する比率）
     * @param {Object} itemA - visibleBounds を持つオブジェクト
     * @param {Object} itemB - visibleBounds を持つオブジェクト
     * @returns {number} 重なり率（重ならないときは 0）
     */
    function getVerticalOverlapRatio(itemA, itemB) {
        /* visibleBounds: [left, top, right, bottom] */
        var boundsA = itemA.visibleBounds;
        var boundsB = itemB.visibleBounds;
        var overlap = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
        if (overlap <= 0) return 0;

        var minHeight = Math.min(boundsA[1] - boundsA[3], boundsB[1] - boundsB[3]);
        return overlap / minHeight;
    }

    /**
     * 複数のオブジェクトを囲む境界を返す
     * @param {Object[]} boundsItems - visibleBounds を持つオブジェクト
     * @returns {{visibleBounds: number[]}} まとめた境界
     */
    function mergeBounds(boundsItems) {
        var firstBounds = boundsItems[0].visibleBounds;
        var left = firstBounds[0];
        var top = firstBounds[1];
        var right = firstBounds[2];
        var bottom = firstBounds[3];

        for (var i = 1; i < boundsItems.length; i++) {
            var itemBounds = boundsItems[i].visibleBounds;
            left = Math.min(left, itemBounds[0]);
            top = Math.max(top, itemBounds[1]);
            right = Math.max(right, itemBounds[2]);
            bottom = Math.min(bottom, itemBounds[3]);
        }

        return { visibleBounds: [left, top, right, bottom] };
    }

    /**
     * 左右に並ぶオブジェクトを同じ行として束ね、行ごとの境界を返す
     * @param {Object[]} boundsItems - visibleBounds を持つオブジェクト
     * @returns {Object[]} 行ごとの境界（{visibleBounds}）
     */
    function groupItemsByRow(boundsItems) {
        var rowMembers = [];

        for (var i = 0; i < boundsItems.length; i++) {
            var isPlaced = false;
            for (var r = 0; r < rowMembers.length; r++) {
                /* 行の代表要素（最初の1つ）と縦方向の重なりを比較 / compare with the first item of the row */
                if (getVerticalOverlapRatio(boundsItems[i], rowMembers[r][0]) >= ROW_OVERLAP_RATIO) {
                    rowMembers[r].push(boundsItems[i]);
                    isPlaced = true;
                    break;
                }
            }
            if (!isPlaced) rowMembers.push([boundsItems[i]]);
        }

        var rowBoundsList = [];
        for (var j = 0; j < rowMembers.length; j++) {
            rowBoundsList.push(mergeBounds(rowMembers[j]));
        }
        return rowBoundsList;
    }

    /**
     * 選択を行にまとめ、上から順に並べた行の境界を返す
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @param {PageItem[]} tempItems - 計測用に作った一時オブジェクトの控え（追加される）
     * @returns {Object[]} 上から順の行の境界（{visibleBounds}）
     */
    function collectSortedRows(selectedItems, tempItems) {
        var proxyItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            proxyItems.push(makeBoundsProxy(selectedItems[i], tempItems));
        }

        var rowBoundsList = groupItemsByRow(proxyItems);
        rowBoundsList.sort(function (rowA, rowB) {
            /* top が大きいほうが上（一般的な定規の設定を想定）/ larger top = higher */
            return rowB.visibleBounds[1] - rowA.visibleBounds[1];
        });
        return rowBoundsList;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 複数の境界の左右の端を返す
     * @param {Object[]} boundsItems - visibleBounds を持つオブジェクト
     * @returns {{left: number, right: number}} 左端と右端
     */
    function getHorizontalSpan(boundsItems) {
        var left = boundsItems[0].visibleBounds[0];
        var right = boundsItems[0].visibleBounds[2];
        for (var i = 1; i < boundsItems.length; i++) {
            left = Math.min(left, boundsItems[i].visibleBounds[0]);
            right = Math.max(right, boundsItems[i].visibleBounds[2]);
        }
        return { left: left, right: right };
    }

    /**
     * 上下に並ぶ行の間にケイ線を描く
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} rowBoundsList - 上から順の行の境界
     * @param {Object} ruleSettings - showRulesDialog() の戻り値
     * @returns {PathItem[]} 描いたケイ線
     */
    function drawRulesBetweenRows(doc, rowBoundsList, ruleSettings) {
        /* 線の色（K=100）/ Line color: black K=100 */
        var lineColor = new CMYKColor();
        lineColor.cyan = 0;
        lineColor.magenta = 0;
        lineColor.yellow = 0;
        lineColor.black = 100;

        /* 延長：0 ならオブジェクトと同じ幅、正の数なら長く、負の数なら短く / 0 = same width, + longer, - shorter */
        var extension = ruleSettings.extension;

        /* 「共通」は全体の左右幅、「オブジェクトに合わせる」は上下の行の左右幅 / common span vs per-pair span */
        var commonSpan = (ruleSettings.ruleLengthMode === "common" && rowBoundsList.length > 1)
            ? getHorizontalSpan(rowBoundsList)
            : null;

        var createdLines = [];
        for (var i = 0; i < rowBoundsList.length - 1; i++) {
            var upperRow = rowBoundsList[i];
            var lowerRow = rowBoundsList[i + 1];

            /* 上の行の下端と下の行の上端の中間 / midway between the upper bottom and the lower top */
            var midY = (upperRow.visibleBounds[3] + lowerRow.visibleBounds[1]) / 2;
            var ruleSpan = commonSpan || getHorizontalSpan([upperRow, lowerRow]);

            var ruleLine = doc.pathItems.add();
            ruleLine.setEntirePath([[ruleSpan.left - extension, midY], [ruleSpan.right + extension, midY]]);
            ruleLine.filled = false;
            ruleLine.stroked = true;
            ruleLine.strokeWidth = ruleSettings.lineWeight;
            ruleLine.strokeColor = lineColor;
            ruleLine.strokeCap = ruleSettings.lineCap;
            createdLines.push(ruleLine);
        }
        return createdLines;
    }

    /**
     * 計測用に作った一時オブジェクトを削除する
     * @param {PageItem[]} tempItems - 一時オブジェクト
     * @returns {void}
     */
    function removeTemporaryItems(tempItems) {
        for (var i = 0; i < tempItems.length; i++) {
            try {
                if (tempItems[i] && tempItems[i].typename) tempItems[i].remove();
            } catch (e) { }
        }
    }

    /**
     * 描いたケイ線を選択する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} createdLines - 描いたケイ線
     * @returns {void}
     */
    function selectCreatedLines(doc, createdLines) {
        if (createdLines.length === 0) return;
        doc.selection = null;
        for (var i = 0; i < createdLines.length; i++) {
            try {
                createdLines[i].selected = true;
            } catch (e) { }
        }
    }

    /**
     * ダイアログで設定を受け取り、選択したオブジェクトの間にケイ線を描く
     * @returns {void}
     */
    function main() {
        var ruleSettings = showRulesDialog();
        if (ruleSettings === null) return;

        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        /* テキストとグループは複製→アウトライン化した境界で計算し、最後に削除する / measured on outlined duplicates */
        var tempItems = [];
        var rowBoundsList = collectSortedRows(doc.selection, tempItems);
        var createdLines = drawRulesBetweenRows(doc, rowBoundsList, ruleSettings);

        removeTemporaryItems(tempItems);
        selectCreatedLines(doc, createdLines);
    }

    main();

})();
