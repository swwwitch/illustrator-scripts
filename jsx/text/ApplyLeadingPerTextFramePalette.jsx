#target illustrator
#targetengine "ApplyLeadingPerTextFrame"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択された各テキストフレームの各行について、行頭数文字のフォントサイズを基準に行送りを再計算して適用します。
適用する行送りの割合はダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyLeadingPerTextFramePalette.md

### Overview

Recalculates the leading of each line in the selected text frames from the font size of the first few characters, and applies it.
The leading percentage is chosen in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyLeadingPerTextFramePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyLeadingPerTextFramePalette";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyLeadingPerTextFramePalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyLeadingPerTextFramePalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var LINE_FONT_SIZE_SAMPLE_COUNT = 5;   /* 行内で参照する文字数 / Characters sampled per line */
    var DEFAULT_AUTO_LEADING_AMOUNT = 175; /* 読み取れないときの自動行送り量（%）/ Auto leading amount (%) when nothing can be read */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PALETTE_MARGINS = 16;              /* パレットの余白 / Palette margins */
    var PALETTE_SPACING = 12;              /* パネル同士の間隔 / Spacing between panels */
    var PANEL_MARGINS = [15, 20, 15, 10];  /* パネルの余白 [左,上,右,下] / Panel margins */
    var PANEL_SPACING = 8;                 /* パネル内の間隔 / Spacing inside panels */
    var SHORT_INPUT_CHARACTERS = 3;        /* 行送り・自動行送り量の欄の幅（文字数）/ Width of the leading fields */
    var SPACE_INPUT_CHARACTERS = 4;        /* 段落前後のアキの欄の幅（文字数）/ Width of the paragraph spacing fields */

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

    var LABELS = {
        dialog: {
            title: { ja: "行送りの設定", en: "Leading Settings" }
        },
        panel: {
            leading: { ja: "行送り", en: "Leading settings" },
            leadingType: { ja: "行送りの基準", en: "Leading basis" },
            paragraphSpacing: { ja: "段落前後のアキ", en: "Paragraph spacing" }
        },
        radio: {
            topToTop: { ja: "仮想ボディの上基準", en: "Top-to-top (virtual body)" },
            bottomToBottom: { ja: "欧文ベースライン基準", en: "Baseline-to-baseline" },
            otherLeading: { ja: "その他", en: "Other" }
        },
        fieldLabel: {
            spaceBefore: { ja: "段落前", en: "Space before" },
            spaceAfter: { ja: "段落後", en: "Space after" }
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
            autoAmount: {
                ja: "自動行送り量（％）。↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1",
                en: "Auto leading amount (%). ↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1"
            },
            leading: {
                ja: "行送り値（［その他］で直接指定）。↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1",
                en: "Leading value (use Other to set directly). ↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1"
            },
            space: { ja: "↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1", en: "↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1" },
            otherLeading: {
                ja: "左の行送り値をそのまま使います。段落ごとに、各行の先頭の文字サイズから自動行送り量（％）を逆算して設定します",
                en: "Uses the leading value on the left as is: for each paragraph, the auto leading percentage is worked back from the font size at the start of its lines."
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

    // =========================================
    // 行送りの選択肢 / Leading choices
    // =========================================

    var LEADING_CHOICES = [
        { label: "110%", ratio: 1.1, token: "110", isOther: false },
        { label: "125%", ratio: 1.25, token: "125", isOther: false },
        { label: "150%", ratio: 1.5, token: "150", isOther: false },
        { label: getLabel("radio.otherLeading"), ratio: undefined, token: "OTHER", isOther: true }
    ];

    var LEADING_TYPE_CHOICES = [
        { label: getLabel("radio.topToTop"), token: "TOPTOTOP" },
        { label: getLabel("radio.bottomToBottom"), token: "BOTTOMTOBOTTOM" }
    ];

    /**
     * トークンに対応する行送りの選択肢の番号を返す
     * @param {string} choiceToken - "110" / "125" / "150" / "OTHER"
     * @returns {number} 選択肢の番号（見つからなければ 0）
     */
    function getLeadingChoiceIndexByToken(choiceToken) {
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            if (LEADING_CHOICES[i].token === choiceToken) return i;
        }
        return 0;
    }

    /**
     * ［その他］の選択肢の番号を返す
     * @returns {number} 選択肢の番号（無ければ -1）
     */
    function findOtherChoiceIndex() {
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            if (LEADING_CHOICES[i].isOther) return i;
        }
        return -1;
    }

    // =========================================
    // worker 関数 / Worker functions (run in the MAIN engine via BridgeTalk)
    // toString() で送るため JSDoc は付けず、説明は1行の /* */ コメントにする
    // 注意: 内部は // 行コメント禁止・/* */ のみ・必ずセミコロンで終える（toString が改行を消すため）
    // =========================================

    /* 文字の編集中なら、そのテキストフレームを選択し直す / While editing text, select its text frame instead */
    function w_normalizeSelection() {
        if (app.documents.length > 0 && app.selection && app.selection.typename === "TextRange") {
            var story = app.selection.story;
            if (story && story.textFrames.length === 1) {
                var parentTextFrame = story.textFrames[0];
                app.executeMenuCommand("deselectall");
                app.selection = [parentTextFrame];
                try { app.selectTool("Adobe Select Tool"); } catch (e) { }
            }
        }
    }

    /* 文字と行のあるテキストフレームだけを返す（それ以外は null）/ Return the item only when it is a text frame with text and lines */
    function w_getProcessableTextFrame(selectedItem) {
        if (!selectedItem || selectedItem.typename !== "TextFrame") { return null; }
        if (!selectedItem.contents) { return null; }
        if (!selectedItem.lines || selectedItem.lines.length === 0) { return null; }
        return selectedItem;
    }

    /* 選択中の処理できるテキストフレームを重複なく集める / Collect the processable selected text frames without duplicates */
    function w_collectTextFrames() {
        var selectionItems = app.activeDocument.selection;
        var textFrames = [];
        if (!selectionItems || selectionItems.length === 0) { return textFrames; }
        for (var i = 0; i < selectionItems.length; i++) {
            var textFrame = w_getProcessableTextFrame(selectionItems[i]);
            if (!textFrame) { continue; }
            var isDuplicate = false;
            for (var j = 0; j < textFrames.length; j++) { if (textFrames[j] === textFrame) { isDuplicate = true; break; } }
            if (!isDuplicate) { textFrames.push(textFrame); }
        }
        return textFrames;
    }

    /* 行頭から sampleCount 文字ぶんのフォントサイズを集める / Sample the font sizes of the first sampleCount characters of a line */
    function w_sampleLineFontSizes(line, sampleCount) {
        var fontSizes = [];
        if (!line || !line.characters || line.characters.length === 0) { return fontSizes; }
        var maxCount = Math.min(line.characters.length, sampleCount);
        for (var i = 0; i < maxCount; i++) {
            try {
                var fontSize = line.characters[i].characterAttributes.size;
                if (!isNaN(fontSize)) { fontSizes.push(fontSize); }
            } catch (e) { }
        }
        return fontSizes;
    }

    /* 最も多い値を返す（同数なら大きいほう）/ Return the most frequent value (the larger one on a tie) */
    function w_getMostFrequentValue(values) {
        if (!values || values.length === 0) { return NaN; }
        var valueCounts = {};
        var bestValue = values[0];
        var bestCount = 0;
        for (var i = 0; i < values.length; i++) {
            var valueKey = String(values[i]);
            if (!valueCounts[valueKey]) { valueCounts[valueKey] = { value: values[i], count: 0 }; }
            valueCounts[valueKey].count++;
            if (valueCounts[valueKey].count > bestCount) { bestCount = valueCounts[valueKey].count; bestValue = valueCounts[valueKey].value; }
            else if (valueCounts[valueKey].count === bestCount && valueCounts[valueKey].value > bestValue) { bestValue = valueCounts[valueKey].value; }
        }
        return bestValue;
    }

    /* 各行の行頭のフォントサイズを sampledSizes に足していく / Append the line-start font sizes of every line to sampledSizes */
    function w_pushLineSampleSizes(textObject, sampledSizes) {
        var textLines = textObject.lines;
        for (var j = 0; j < textLines.length; j++) {
            var lineSizes = w_sampleLineFontSizes(textLines[j], LINE_FONT_SIZE_SAMPLE_COUNT);
            for (var s = 0; s < lineSizes.length; s++) { sampledSizes.push(lineSizes[s]); }
        }
    }

    /* 段落の基準フォントサイズ（途中で失敗しても集めたぶんで判定）/ Base font size of a paragraph (uses what was sampled even on failure) */
    function w_getParagraphBaseFontSize(paragraph) {
        var sampledSizes = [];
        try {
            w_pushLineSampleSizes(paragraph, sampledSizes);
        } catch (e) { }
        if (sampledSizes.length === 0) { return NaN; }
        return w_getMostFrequentValue(sampledSizes);
    }

    /* テキストフレーム全体の基準フォントサイズ / Base font size of a whole text frame */
    function w_getFrameBaseFontSize(textFrame) {
        var sampledSizes = [];
        w_pushLineSampleSizes(textFrame, sampledSizes);
        if (sampledSizes.length === 0) { return NaN; }
        return w_getMostFrequentValue(sampledSizes);
    }

    /* トークンを AutoLeadingType に変換する / Convert a token to AutoLeadingType */
    function w_resolveLeadingType(leadingTypeToken) {
        if (leadingTypeToken === "BOTTOMTOBOTTOM") { return AutoLeadingType.BOTTOMTOBOTTOM; }
        return AutoLeadingType.TOPTOTOP;
    }

    /* 選択中のテキストフレームに行送りと段落前後のアキを適用し、代表の行送り（pt）を返す / Apply leading and paragraph spacing, then return a representative leading (pt) */
    function w_applyLeading(autoAmount, directMode, directLeadingPt, spaceBefore, spaceAfter, leadingTypeToken) {
        if (app.documents.length === 0) { return "NODOC"; }
        w_normalizeSelection();
        var textFrames = w_collectTextFrames();
        if (textFrames.length === 0) { return "NOSEL"; }
        var leadingType = w_resolveLeadingType(leadingTypeToken);
        var useDirect = directMode && !isNaN(directLeadingPt);
        var representativePt = NaN;
        try {
            for (var i = 0; i < textFrames.length; i++) {
                var textFrame = textFrames[i];
                textFrame.textRange.characterAttributes.autoLeading = true;
                var frameParagraphs = textFrame.paragraphs;
                for (var p = 0; p < frameParagraphs.length; p++) {
                    try {
                        var paragraph = frameParagraphs[p];
                        var amount;
                        if (useDirect) {
                            var baseFontSize = w_getParagraphBaseFontSize(paragraph);
                            if (isNaN(baseFontSize) || baseFontSize <= 0) { continue; }
                            amount = (directLeadingPt / baseFontSize) * 100;
                        } else {
                            amount = autoAmount;
                        }
                        paragraph.characterAttributes.autoLeading = true;
                        paragraph.paragraphAttributes.spaceBefore = spaceBefore;
                        paragraph.paragraphAttributes.spaceAfter = spaceAfter;
                        paragraph.paragraphAttributes.autoLeadingAmount = amount;
                    } catch (ep) { }
                }
                try { textFrame.textRange.leadingType = leadingType; } catch (et) { }
            }
            if (useDirect) {
                representativePt = directLeadingPt;
            } else {
                var frameBaseFontSize = w_getFrameBaseFontSize(textFrames[0]);
                if (!isNaN(frameBaseFontSize)) { representativePt = frameBaseFontSize * (autoAmount / 100); }
            }
        } catch (e) {
            return "ERR:" + e.message;
        }
        app.redraw();
        return "OK|" + representativePt;
    }

    /* 選択中のテキストフレームから初期値を読む / Read the initial values from the selected text frames */
    function w_readInitial() {
        if (app.documents.length === 0) { return "NODOC"; }
        w_normalizeSelection();
        var textFrames = w_collectTextFrames();
        if (textFrames.length === 0) { return "NOSEL"; }
        var autoAmount = DEFAULT_AUTO_LEADING_AMOUNT;
        var leadingPt = NaN;
        var leadingTypeToken = "TOPTOTOP";
        var spaceBefore = 0;
        var spaceAfter = 0;
        var choiceToken = "OTHER";
        var isAuto = false;
        for (var i = 0; i < textFrames.length; i++) {
            try {
                var frameLines = textFrames[i].lines;
                if (frameLines && frameLines.length > 0 && frameLines[0].characters.length > 0) {
                    var firstCharAttrs = frameLines[0].characters[0].characterAttributes;
                    isAuto = firstCharAttrs.autoLeading;
                    if (!isNaN(firstCharAttrs.leading)) { leadingPt = firstCharAttrs.leading; }
                }
            } catch (e) { }
            if (!isNaN(leadingPt)) { break; }
        }
        for (var k = 0; k < textFrames.length; k++) {
            try {
                var frameParagraphs = textFrames[k].paragraphs;
                if (frameParagraphs && frameParagraphs.length > 0) {
                    var paraAttrs = frameParagraphs[0].paragraphAttributes;
                    if (!isNaN(paraAttrs.autoLeadingAmount)) { autoAmount = paraAttrs.autoLeadingAmount; }
                    if (!isNaN(paraAttrs.spaceBefore)) { spaceBefore = paraAttrs.spaceBefore; }
                    if (!isNaN(paraAttrs.spaceAfter)) { spaceAfter = paraAttrs.spaceAfter; }
                    break;
                }
            } catch (e2) { }
        }
        for (var t = 0; t < textFrames.length; t++) {
            try {
                var frameLeadingType = textFrames[t].textRange.leadingType;
                if (frameLeadingType !== undefined && frameLeadingType !== null) {
                    if (frameLeadingType === AutoLeadingType.BOTTOMTOBOTTOM) { leadingTypeToken = "BOTTOMTOBOTTOM"; }
                    else { leadingTypeToken = "TOPTOTOP"; }
                    break;
                }
            } catch (e3) { }
        }
        if (isAuto) {
            var presetTokens = ["110", "125", "150"];
            for (var c = 0; c < presetTokens.length; c++) {
                if (Math.abs(autoAmount - Number(presetTokens[c])) < 0.5) { choiceToken = presetTokens[c]; break; }
            }
        }
        return "OK|" + autoAmount + "|" + leadingPt + "|" + leadingTypeToken + "|" + spaceBefore + "|" + spaceAfter + "|" + choiceToken;
    }

    // worker 関数はすべてここに登録（追加漏れ防止） / Register every worker function here
    var WORKER_FUNCS = [
        w_normalizeSelection,
        w_getProcessableTextFrame,
        w_collectTextFrames,
        w_sampleLineFontSizes,
        w_getMostFrequentValue,
        w_pushLineSampleSizes,
        w_getParagraphBaseFontSize,
        w_getFrameBaseFontSize,
        w_resolveLeadingType,
        w_applyLeading,
        w_readInitial
    ];

    // =========================================
    // BridgeTalk 委譲 / Delegation to the main engine
    // =========================================

    var isBusy = false;

    /**
     * worker 関数群のソースを連結して 1 つのコード文字列にする。
     * 先頭で共有定数を宣言し、後続の呼び出し式から参照できるようにする。
     * @returns {string} メインエンジンで eval するソース
     */
    function buildWorkerSource() {
        var workerSource = "var LINE_FONT_SIZE_SAMPLE_COUNT=" + LINE_FONT_SIZE_SAMPLE_COUNT + ";" +
            "var DEFAULT_AUTO_LEADING_AMOUNT=" + DEFAULT_AUTO_LEADING_AMOUNT + ";";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            workerSource += WORKER_FUNCS[i].toString();
        }
        return workerSource;
    }

    /**
     * worker 呼び出し式をメインエンジンへ同期委譲し、マーカー文字列を受け取る。
     * @param {string} callExpr - "w_applyLeading(...)" 等、文字列を返す呼び出し式
     * @returns {string} マーカー（"OK"/"OK|..."/"NODOC"/"NOSEL"/"ERR:..."）
     */
    function runWorker(callExpr) {
        if (isBusy) { return "ERR:busy"; }
        isBusy = true;
        var workerResult = { value: null };
        /* BridgeTalk の送信は失敗しうる / BridgeTalk sending can fail */
        try {
            var workerCode = buildWorkerSource() + "String(" + callExpr + ");";
            var bridgeTalk = new BridgeTalk();
            bridgeTalk.target = "illustrator";
            bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerCode) + "\"));";
            bridgeTalk.onResult = function (resultMessage) { workerResult.value = resultMessage.body; };
            bridgeTalk.onError = function (errorMessage) { workerResult.value = "ERR:" + errorMessage.body; };
            bridgeTalk.send(10);
        } catch (e) {
            workerResult.value = "ERR:" + e.message;
        } finally {
            isBusy = false;
        }
        return workerResult.value;
    }

    /**
     * マーカー文字列を解析する。
     * @param {string} resultText - runWorker の戻り値
     * @returns {{ok: boolean, code: string, extra: ?string, msg: string}} 解析結果
     */
    function parseWorkerResult(resultText) {
        if (resultText == null) { return { ok: false, code: "ERR", msg: "no response" }; }
        var head = resultText;
        var extra = null;
        var barIndex = resultText.indexOf("|");
        if (barIndex >= 0) {
            head = resultText.substring(0, barIndex);
            extra = resultText.substring(barIndex + 1);
        }
        if (head === "OK") { return { ok: true, code: "OK", extra: extra }; }
        if (head === "NODOC") { return { ok: false, code: "NODOC" }; }
        if (head === "NOSEL") { return { ok: false, code: "NOSEL" }; }
        if (head.indexOf("ERR") === 0) { return { ok: false, code: "ERR", msg: resultText.substring(4) }; }
        return { ok: false, code: "ERR", msg: resultText };
    }

    // =========================================
    // 数値ユーティリティ / Numeric helpers
    // =========================================

    /**
     * 文字列を数値にする（数値でなければ代わりの値）
     * @param {string} value - 元の文字列
     * @param {number} fallback - 数値でないときの値
     * @returns {number} 数値
     */
    function toNumber(value, fallback) {
        var parsedValue = parseFloat(value);
        return isNaN(parsedValue) ? fallback : parsedValue;
    }

    /**
     * 手入力値をパースし、下限でクランプする（負数の手入力対策）。
     * @param {string} text - 入力欄の文字列
     * @param {number} minValue - 下限
     * @param {number} fallback - パース失敗時の値
     * @returns {number} クランプ後の値
     */
    function clampMinNumber(text, minValue, fallback) {
        var parsedValue = parseFloat(text);
        if (isNaN(parsedValue)) { parsedValue = fallback; }
        if (parsedValue < minValue) { parsedValue = minValue; }
        return parsedValue;
    }

    /**
     * pt の値を文字の単位に換算して小数第1位の文字列にする
     * @param {number} ptValue - pt の値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {string} 換算した文字列（数値でなければ空文字）
     */
    function formatByUnit(ptValue, textUnit) {
        if (isNaN(ptValue) || ptValue === null) { return ""; }
        return (Math.round((ptValue / textUnit.pointsPerUnit) * 10) / 10).toFixed(1);
    }

    // =========================================
    // UI: ∧∨付きの数値欄 / Number fields with steppers
    // =========================================

    /**
     * ∧∨と数値欄を隙間0で突き合わせて追加する。∧∨・↑↓キーで増減したあとは入力欄の onChange を呼ぶ（負の値は不可）
     * @param {Group} parentGroup - 追加先
     * @param {string} initialText - 初期値
     * @param {number} [decimals] - 増減後に表示する小数の桁数（省略時は丸めた数値のまま）
     * @returns {EditText} 追加した数値欄
     */
    function addSteppedNumberInput(parentGroup, initialText, decimals) {
        var stepperInputGroup = parentGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, {
            min: 0,
            onStep: function (steppedInput) {
                if (typeof decimals === "number") steppedInput.text = parseFloat(steppedInput.text).toFixed(decimals);
                if (typeof steppedInput.onChange === "function") steppedInput.onChange();
            }
        });
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        bindSteppedArrowKeys(numberInput, stepperGroup); /* ↑↓キーも∧∨と同じ処理 / arrow keys share the stepper */
        return numberInput;
    }

    // =========================================
    // 常駐パレット / Persistent palette
    // =========================================

    /**
     * 選択中のテキストフレームから初期値を読む（読めなければ既定値）
     * @returns {{autoAmount: number, leadingPt: number, leadingTypeToken: string, spaceBefore: number, spaceAfter: number, choiceToken: string}} 初期値
     */
    function readInitialSettings() {
        var initResult = parseWorkerResult(runWorker("w_readInitial()"));
        var initialSettings = {
            autoAmount: DEFAULT_AUTO_LEADING_AMOUNT,
            leadingPt: NaN,
            leadingTypeToken: "TOPTOTOP",
            spaceBefore: 0,
            spaceAfter: 0,
            choiceToken: "110"
        };
        if (initResult.ok && initResult.extra != null) {
            var fields = initResult.extra.split("|");
            initialSettings.autoAmount = toNumber(fields[0], DEFAULT_AUTO_LEADING_AMOUNT);
            initialSettings.leadingPt = toNumber(fields[1], NaN);
            initialSettings.leadingTypeToken = fields[2] || "TOPTOTOP";
            initialSettings.spaceBefore = toNumber(fields[3], 0);
            initialSettings.spaceAfter = toNumber(fields[4], 0);
            initialSettings.choiceToken = fields[5] || "110";
        }
        return initialSettings;
    }

    /**
     * 段落前後のアキの1行（項目名＋数値欄＋単位）を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - 項目名のパス
     * @param {number} initialPt - 初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {EditText} 追加した数値欄
     */
    function addSpaceRow(parentPanel, labelPath, initialPt, textUnit) {
        var spaceRow = parentPanel.add("group");
        spaceRow.add("statictext", undefined, labelText(labelPath));
        var spaceInput = addSteppedNumberInput(spaceRow, formatByUnit(initialPt, textUnit));
        spaceInput.characters = SPACE_INPUT_CHARACTERS;
        spaceInput.helpTip = getLabel("tooltip.space");
        spaceRow.add("statictext", undefined, textUnit.label);
        return spaceInput;
    }

    /**
     * パレットを組み立てる（イベントはまだ付けない）
     * @param {Object} initialSettings - readInitialSettings() の戻り値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {Object} パレットと各コントロール
     */
    function buildPalette(initialSettings, textUnit) {
        var leadingPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        leadingPalette.orientation = "column";
        leadingPalette.alignChildren = "fill";
        leadingPalette.margins = PALETTE_MARGINS;
        leadingPalette.spacing = PALETTE_SPACING;

        /* 行送りパネル / Leading panel */
        var leadingPanel = leadingPalette.add("panel", undefined, getLabel("panel.leading"));
        leadingPanel.orientation = "column";
        leadingPanel.alignChildren = "left";
        leadingPanel.margins = PANEL_MARGINS;
        leadingPanel.spacing = PANEL_SPACING;

        var contentGroup = leadingPanel.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["left", "top"];
        contentGroup.spacing = 25;

        /* 左カラム：行送り値（pt 等） / Left column: leading value */
        var leftColumnGroup = contentGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "left";
        leftColumnGroup.spacing = 0;

        var leadingGroup = leftColumnGroup.add("group");
        var leadingInput = addSteppedNumberInput(leadingGroup, formatByUnit(initialSettings.leadingPt, textUnit), 1);
        leadingInput.characters = SHORT_INPUT_CHARACTERS;
        leadingInput.helpTip = getLabel("tooltip.leading");
        leadingGroup.add("statictext", undefined, textUnit.label);

        /* 右カラム：自動行送り量（％）＋プリセット / Right column: auto amount (%) and presets */
        var rightColumnGroup = contentGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "left";
        rightColumnGroup.spacing = 6;

        var autoAmountGroup = rightColumnGroup.add("group");
        autoAmountGroup.orientation = "row";
        autoAmountGroup.alignChildren = "center";
        autoAmountGroup.spacing = 6;
        var autoInput = addSteppedNumberInput(autoAmountGroup, String(Math.round(initialSettings.autoAmount)));
        autoInput.characters = SHORT_INPUT_CHARACTERS;
        autoInput.helpTip = getLabel("tooltip.autoAmount");
        autoAmountGroup.add("statictext", undefined, "%");

        var leadingChoiceGroup = rightColumnGroup.add("group");
        leadingChoiceGroup.orientation = "column";
        leadingChoiceGroup.alignChildren = "left";
        leadingChoiceGroup.spacing = 6;
        leadingChoiceGroup.margins = [0, 5, 0, 0];

        var leadingRadios = [];
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            leadingRadios.push(leadingChoiceGroup.add("radiobutton", undefined, LEADING_CHOICES[i].label));
            if (LEADING_CHOICES[i].isOther) leadingRadios[i].helpTip = getLabel("tooltip.otherLeading");
        }
        leadingRadios[getLeadingChoiceIndexByToken(initialSettings.choiceToken)].value = true;

        /* 行送りの基準パネル（英語 UI ではタイトルなしのグループ）/ Leading type panel (an untitled group in English) */
        var leadingTypeContainer;
        if (uiLang === "ja") {
            leadingTypeContainer = leadingPalette.add("panel", undefined, getLabel("panel.leadingType"));
            leadingTypeContainer.margins = PANEL_MARGINS;
        } else {
            leadingTypeContainer = leadingPalette.add("group");
            leadingTypeContainer.margins = [0, 0, 0, 0];
        }
        leadingTypeContainer.orientation = "column";
        leadingTypeContainer.alignChildren = "left";
        leadingTypeContainer.spacing = PANEL_SPACING;

        var typeRadios = [];
        var initialTypeIndex = 0;
        for (var t = 0; t < LEADING_TYPE_CHOICES.length; t++) {
            typeRadios.push(leadingTypeContainer.add("radiobutton", undefined, LEADING_TYPE_CHOICES[t].label));
            if (LEADING_TYPE_CHOICES[t].token === initialSettings.leadingTypeToken) { initialTypeIndex = t; }
        }
        typeRadios[initialTypeIndex].value = true;

        /* 段落前後のアキパネル / Paragraph spacing panel */
        var spacePanel = leadingPalette.add("panel", undefined, getLabel("panel.paragraphSpacing"));
        spacePanel.orientation = "column";
        spacePanel.alignChildren = "left";
        spacePanel.margins = PANEL_MARGINS;
        spacePanel.spacing = PANEL_SPACING;

        var spaceBeforeInput = addSpaceRow(spacePanel, "fieldLabel.spaceBefore", initialSettings.spaceBefore, textUnit);
        var spaceAfterInput = addSpaceRow(spacePanel, "fieldLabel.spaceAfter", initialSettings.spaceAfter, textUnit);

        return {
            palette: leadingPalette,
            leadingInput: leadingInput,
            autoInput: autoInput,
            leadingRadios: leadingRadios,
            typeRadios: typeRadios,
            spaceBeforeInput: spaceBeforeInput,
            spaceAfterInput: spaceAfterInput
        };
    }

    /**
     * ラジオボタンの配列から選ばれている番号を返す
     * @param {RadioButton[]} radioButtons - ラジオボタンの配列
     * @returns {number} 選ばれている番号（無ければ -1）
     */
    function getSelectedRadioIndex(radioButtons) {
        for (var r = 0; r < radioButtons.length; r++) {
            if (radioButtons[r].value) return r;
        }
        return -1;
    }

    /**
     * UI から適用オプションを読み取る（負数はクランプ）。
     * @param {Object} paletteControls - buildPalette() の戻り値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {Object} { invalid, directMode, autoAmount, directLeadingPt, spaceBefore, spaceAfter, leadingTypeToken }
     */
    function readApplyOptions(paletteControls, textUnit) {
        var choiceIndex = getSelectedRadioIndex(paletteControls.leadingRadios);
        var directMode = (choiceIndex >= 0) && !!LEADING_CHOICES[choiceIndex].isOther;
        var autoAmount = clampMinNumber(paletteControls.autoInput.text, 0, DEFAULT_AUTO_LEADING_AMOUNT);
        var spaceBefore = clampMinNumber(paletteControls.spaceBeforeInput.text, 0, 0) * textUnit.pointsPerUnit;
        var spaceAfter = clampMinNumber(paletteControls.spaceAfterInput.text, 0, 0) * textUnit.pointsPerUnit;
        var typeIndex = getSelectedRadioIndex(paletteControls.typeRadios);
        var leadingTypeToken = LEADING_TYPE_CHOICES[typeIndex >= 0 ? typeIndex : 0].token;
        var directLeadingPt = NaN;
        if (directMode) {
            var leadingValue = parseFloat(paletteControls.leadingInput.text);
            if (isNaN(leadingValue)) { return { invalid: true }; }
            if (leadingValue < 0) { leadingValue = 0; }
            directLeadingPt = leadingValue * textUnit.pointsPerUnit;
        }
        return {
            invalid: false,
            directMode: directMode,
            autoAmount: autoAmount,
            directLeadingPt: directLeadingPt,
            spaceBefore: spaceBefore,
            spaceAfter: spaceAfter,
            leadingTypeToken: leadingTypeToken
        };
    }

    /**
     * 適用オプションから w_applyLeading の呼び出し式を作る
     * @param {Object} applyOptions - readApplyOptions() の戻り値
     * @returns {string} 呼び出し式
     */
    function buildApplyLeadingCall(applyOptions) {
        return "w_applyLeading(" +
            applyOptions.autoAmount + "," +
            applyOptions.directMode + "," +
            applyOptions.directLeadingPt + "," +
            applyOptions.spaceBefore + "," +
            applyOptions.spaceAfter + ",'" +
            applyOptions.leadingTypeToken + "'" +
            ")";
    }

    /**
     * パレットにイベントを付ける
     * @param {Object} paletteControls - buildPalette() の戻り値
     * @param {{label: string, pointsPerUnit: number}} textUnit - 文字の単位
     * @returns {void}
     */
    function bindPaletteEvents(paletteControls, textUnit) {
        var leadingPalette = paletteControls.palette;
        var leadingInput = paletteControls.leadingInput;
        var autoInput = paletteControls.autoInput;
        var leadingRadios = paletteControls.leadingRadios;
        var isSyncingUI = false;

        /* index 以外の選択を外す（-1 ならすべて外す）/ Select only index (-1 clears all) */
        function selectLeadingChoice(index) {
            for (var r = 0; r < leadingRadios.length; r++) { leadingRadios[r].value = (r === index); }
        }

        /**
         * 現在の UI 値を選択中のテキストフレームへ即適用する。
         * 絶対値で上書きする冪等な処理なので、操作のたびに呼んでも累積しない。
         * @returns {void}
         */
        function applyToSelection() {
            if (isSyncingUI) { return; }
            var applyOptions = readApplyOptions(paletteControls, textUnit);
            if (applyOptions.invalid) { return; }

            var applyResult = parseWorkerResult(runWorker(buildApplyLeadingCall(applyOptions)));
            if (applyResult.ok && !applyOptions.directMode && applyResult.extra != null) {
                var representativePt = parseFloat(applyResult.extra);
                if (!isNaN(representativePt)) {
                    isSyncingUI = true;
                    leadingInput.text = formatByUnit(representativePt, textUnit);
                    isSyncingUI = false;
                }
            }
        }

        /* 自動行送り量を手で変えたらプリセットの選択を外す / Editing the amount clears the preset choice */
        function onAutoAmountEdited() {
            selectLeadingChoice(-1);
            applyToSelection();
        }

        /* 行送り値を手で変えたら［その他］を選ぶ / Editing the leading value selects Other */
        function onLeadingPtEdited() {
            var otherIndex = findOtherChoiceIndex();
            if (otherIndex >= 0) { selectLeadingChoice(otherIndex); }
            applyToSelection();
        }

        for (var rk = 0; rk < leadingRadios.length; rk++) {
            (function (index) {
                leadingRadios[index].onClick = function () {
                    selectLeadingChoice(index);
                    var ratio = LEADING_CHOICES[index].ratio;
                    if (typeof ratio === "number") {
                        isSyncingUI = true;
                        autoInput.text = String(Math.round(ratio * 100));
                        isSyncingUI = false;
                    }
                    applyToSelection();
                };
            })(rk);
        }

        for (var tk = 0; tk < paletteControls.typeRadios.length; tk++) {
            paletteControls.typeRadios[tk].onClick = applyToSelection;
        }

        autoInput.onChange = onAutoAmountEdited;
        leadingInput.onChange = onLeadingPtEdited;
        paletteControls.spaceBeforeInput.onChange = applyToSelection;
        paletteControls.spaceAfterInput.onChange = applyToSelection;

        /* Esc で閉じる / Close on Esc */
        leadingPalette.addEventListener("keydown", function (kbEvent) {
            if (kbEvent.keyName === "Escape") { leadingPalette.close(); }
        });

        leadingPalette.onClose = function () {
            $.global.__ALPTF_PALETTE__ = null;
            return true;
        };
    }

    /**
     * 常駐パレットを表示する（開いているものは閉じてから開き直す）
     * @returns {void}
     */
    function showPalette() {
        /* 多重起動防止：既存パレットがあれば閉じる（破棄済みだと例外）/ Prevent multiple launches; a disposed palette throws */
        if ($.global.__ALPTF_PALETTE__) {
            try { $.global.__ALPTF_PALETTE__.close(); } catch (e) { }
            $.global.__ALPTF_PALETTE__ = null;
        }

        var textUnit = getUnitInfo("text/units");
        var initialSettings = readInitialSettings();
        var paletteControls = buildPalette(initialSettings, textUnit);
        bindPaletteEvents(paletteControls, textUnit);

        $.global.__ALPTF_PALETTE__ = paletteControls.palette;
        paletteControls.palette.center();
        paletteControls.palette.show();
    }

    showPalette();

})();
