#target illustrator
#targetengine "AdjustPairGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの間隔と位置を、指定した値にそろえます。
グループ内の等間隔配置・最も近いもの同士のペア・アートボード端からのマージンの3モードがあり、
［固定］で選んだ側は動かさず、残りをライブプレビューで確認しながら動かします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustPairGap.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nc8fab19d8164

### Overview

Sets the gap and the position of the selected objects to a value you specify.
Three modes — even spacing inside a group, nearest-neighbour pairs, and margins from an
artboard edge — hold the side picked as the key object and move the rest, with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustPairGap.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustPairGap";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustPairGap.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustPairGap.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc8fab19d8164"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_GAP = 30; /* 平均間隔を測れないときのフォールバック（pt）/ Fallback gap when no average can be measured (pt) */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS       = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / Panel margins [left, top, right, bottom] */
    var PANEL_SPACING       = 8;                /* パネル内の要素間隔 / Spacing inside panels */
    var GRID_CELL_SIZE      = [22, 20];         /* ［固定］の十字のセルの幅・高さ (px) / Cross-grid cell width, height (px) */
    var JUSTIFY_BUTTON_SIZE = [26, 26];         /* 行揃えボタンの幅・高さ (px) / Justification button width, height (px) */

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
     * グループの共通設定を適用する（row は横並びなので縦中央、column は縦並びなので左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（既定）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    /* ダイアログの数値は定規の単位で表示し、pt に換算して使う / Dialog values are shown in the ruler unit and converted to pt */
    var rulerUnit = getUnitInfo();
    var rulerUnitLabel = rulerUnit.label;
    var pointsPerUnit = rulerUnit.pointsPerUnit; /* 1単位 = pointsPerUnit pt */

    /**
     * pt の値を定規の単位の表示文字列にする（小数第2位で丸める）
     * @param {number} points - pt の値
     * @returns {string} 表示文字列
     */
    function pointsToDisplayText(points) {
        return String(Math.round((points / pointsPerUnit) * 100) / 100);
    }

    /**
     * 表示文字列（定規の単位）を pt に換算する。空・不正は 0
     * @param {string} displayText - 入力欄の文字列
     * @returns {number} pt の値
     */
    function displayTextToPoints(displayText) {
        var value = parseFloat(displayText);
        if (isNaN(value)) value = 0;
        return value * pointsPerUnit;
    }

    /**
     * 保存した pt の文字列を定規の単位の表示文字列へ戻す。無効なら null
     * @param {string} pointsText - 保存済みの pt の文字列
     * @returns {string|null} 表示文字列
     */
    function savedPointsToDisplayText(pointsText) {
        if (pointsText === undefined || pointsText === null || pointsText === "") return null;
        var points = parseFloat(pointsText);
        if (isNaN(points)) return null;
        return pointsToDisplayText(points);
    }

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

    var LABELS = {
        dialog: {
            title: { ja: "ペア配置の調整", en: "Adjust Pair Layout" }
        },
        panel: {
            mode: { ja: "モード", en: "Mode" },
            fixedSide: { ja: "固定", en: "Key Object" },
            offset: { ja: "オフセット", en: "Offset" },
            position: { ja: "位置調整", en: "Position" },
            justify: { ja: "テキストの行揃え", en: "Text alignment" }
        },
        radio: {
            modeGroup: { ja: "グループ", en: "Group" },
            modeAutoPair: { ja: "自動ペア認識", en: "Auto Pair Detection" },
            modeArtboard: { ja: "アートボード", en: "Artboard" },
            alignNone: { ja: "なし", en: "None" },
            alignLeft: { ja: "左", en: "Left" },
            alignTop: { ja: "上", en: "Top" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
        },
        fieldLabel: {
            spacing: { ja: "間隔", en: "Gap" },
            align: { ja: "整列", en: "Align" },
            position: { ja: "位置", en: "Position" }
        },
        /* OK はローカライズしない / "OK" is not localized */
        button: {
            justifyAuto: { ja: "自動", en: "Auto" },
            justifyLeft: { ja: "左", en: "Left" },
            justifyCenter: { ja: "中央", en: "Center" },
            justifyRight: { ja: "右", en: "Right" },
            justifyFull: { ja: "均等配置", en: "Justify" },
            cancel: { ja: "キャンセル", en: "Cancel" }
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
            modeGroup: {
                ja: "選択した各グループの中身を、隣り合う間隔がすべて同じになるように並べます（3個以上も対応）。",
                en: "Distributes the contents of each selected group so every adjacent gap is equal (3+ objects supported)."
            },
            modeAutoPair: {
                ja: "選択オブジェクトを最も近いもの同士でペアにし、各ペアの間隔をそろえます。",
                en: "Pairs the selected objects by nearest neighbor and sets the gap of each pair."
            },
            modeArtboard: {
                ja: "アートボードの端（キーオブジェクトで選んだ上／左／右／下）を基準に、各選択オブジェクトとの間隔（マージン）を指定値にそろえます。",
                en: "Sets each selected object's gap (margin) to the artboard edge chosen in Key Object (top/left/right/bottom)."
            },
            fixedSide: {
                ja: "基準にする側。選んだ側は動かさず残りを移動します（左右＝水平、上下＝垂直）。",
                en: "The anchor side. The chosen side stays put while the rest move (Left/Right = horizontal, Top/Bottom = vertical)."
            },
            spacing: {
                ja: "オブジェクト間の間隔。マイナスにすると重なります。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
                en: "Gap between objects; a negative value overlaps them. Arrow keys ±1, Shift ±10, Option ±0.1."
            },
            offsetHorizontal: {
                ja: "上下をキーにしたとき有効。整列後の移動側を左右へ追加でずらします（正＝右）。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
                en: "Active when the key is top/bottom. Nudges the moved side horizontally after alignment (positive = right). Arrow keys ±1, Shift ±10, Option ±0.1."
            },
            offsetVertical: {
                ja: "左右をキーにしたとき有効。整列後の移動側を上下へ追加でずらします（正＝下）。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
                en: "Active when the key is left/right. Nudges the moved side vertically after alignment (positive = down). Arrow keys ±1, Shift ±10, Option ±0.1."
            },
            alignH: {
                ja: "縦に並べたとき（上下をキーに）、動く側をキーオブジェクトの左端／中央／右端にそろえます。",
                en: "When stacking vertically (top/bottom key), aligns moved objects to the key object's left/center/right."
            },
            alignV: {
                ja: "横に並べたとき（左右をキーに）、動く側をキーオブジェクトの上端／中央／下端にそろえます。",
                en: "When laying out horizontally (left/right key), aligns moved objects to the key object's top/center/bottom."
            },
            previewBounds: {
                ja: "オンで線幅や効果を含む見た目の境界、オフで幾何境界を基準に間隔を測ります。",
                en: "On: measure by visible bounds (incl. stroke/effects). Off: geometric bounds."
            },
            justifyAuto: {
                ja: "整列・キーに連動。エリア内文字は均等配置、ポイント文字は縦並びなら水平整列・横並びならキー側に合わせます。",
                en: "Linked to align/key: area text is justified; point text follows the horizontal align (vertical stack) or the key side (horizontal row)."
            },
            justifyLeft: { ja: "テキストを左揃えにします。", en: "Left-aligns the text." },
            justifyCenter: { ja: "テキストを中央揃えにします。", en: "Center-aligns the text." },
            justifyRight: { ja: "テキストを右揃えにします。", en: "Right-aligns the text." },
            justifyFull: { ja: "テキストを均等配置（最終行左）にします。", en: "Justifies the text (last line left-aligned)." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo: { ja: "オブジェクトを2つ以上選択してください。", en: "Select at least two objects." }
        }
    };

    /**
     * 単位を括弧で添えたパネル名を返す（日本語は全角括弧、英語は半角）。
     * 各行に単位を並べる代わりにパネル名へまとめる
     * @param {string} labelPath - パネル名のラベルのパス
     * @param {string} unitLabel - 単位の表示ラベル
     * @returns {string} 単位付きのパネル名
     */
    function labelWithUnit(labelPath, unitLabel) {
        return getLabel(labelPath) + (uiLang === "ja" ? "（" + unitLabel + "）" : " (" + unitLabel + ")");
    }

    // =========================================
    // キー操作 / Keyboard
    // =========================================

    /**
     * 整列のキーボードショートカットをダイアログに付ける。整列パネルは1枚なので、同じラジオを
     * 今の向きで読み替える。水平（上下キー時）：L=左 / C=中央 / R=右、垂直（左右キー時）：T=上 / M=中央 / B=下。
     * Cmd+C などの修飾キー付きの入力は横取りしない。数値欄にフォーカスがあっても効く
     * @param {Window} targetDialog - キー入力を受けるダイアログ
     * @param {Object} alignRadios - 整列ラジオ（none / start / center / end）
     * @param {Function} isHorizontalAlignActive - 水平の整列が有効なら true を返す関数
     * @param {EditText[]} numericFields - 数値だけの欄（フォーカス中もショートカットを効かせる）
     * @returns {void}
     */
    function addAlignmentKeyHandler(targetDialog, alignRadios, isHorizontalAlignActive, numericFields) {
        /**
         * 向きが合うときだけ整列ラジオを返すショートカットを作る
         * @param {string} radioKey - alignRadios のキー（start / center / end）
         * @param {boolean} forHorizontal - 水平の整列のキーなら true
         * @returns {Function} ショートカットの処理（向きが違えば false、選択済みなら true、それ以外はラジオを返す）
         */
        function pickAlignRadio(radioKey, forHorizontal) {
            return function () {
                /* 今の向きのキーでなければ文字をそのまま通す / Pass the key through for the other orientation */
                if (!!isHorizontalAlignActive() !== forHorizontal) return false;
                var targetRadio = alignRadios[radioKey];
                /* 既に選択済みならキーだけ使って何もしない / Already selected: consume the key only */
                if (targetRadio.value) return true;
                return targetRadio;
            };
        }
        /* ラジオの onClick（onAlignChange）で更新する / The radios' onClick refreshes the dialog */
        addKeyShortcuts(targetDialog, {
            "L": pickAlignRadio("start", true),
            "C": pickAlignRadio("center", true),
            "R": pickAlignRadio("end", true),
            "T": pickAlignRadio("start", false),
            "M": pickAlignRadio("center", false),
            "B": pickAlignRadio("end", false)
        }, { numericFields: numericFields });
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
    // ダイアログの部品 / Dialog parts
    // =========================================

    /**
     * モードパネル（グループ／自動ペア認識／アートボード）を生成する。イベント結線は呼び出し側で行う
     * @param {Window} parentGroup - 追加先
     * @returns {{modeRadios: RadioButton[], getMode: Function, setMode: Function}} モードのラジオと読み書きの関数
     */
    function buildModePanel(parentGroup) {
        var modePanel = parentGroup.add("panel", undefined, getLabel("panel.mode"));
        setupPanel(modePanel, 6);
        modePanel.alignChildren = ["left", "top"]; // ラジオ縦並び / radios stacked
        var modeGroupRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeGroup"));
        var modeAutoPairRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeAutoPair"));
        var modeArtboardRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeArtboard"));
        modeGroupRadio.helpTip = getLabel("tooltip.modeGroup");
        modeAutoPairRadio.helpTip = getLabel("tooltip.modeAutoPair");
        modeArtboardRadio.helpTip = getLabel("tooltip.modeArtboard");

        /**
         * 現在のモードを返す
         * @returns {string} "group" / "artboard" / "auto"
         */
        function getMode() {
            if (modeGroupRadio.value) return "group";
            if (modeArtboardRadio.value) return "artboard";
            return "auto";
        }

        /**
         * モードを選ぶ。"group" / "artboard" 以外は自動ペア認識
         * @param {string} mode - モード
         * @returns {void}
         */
        function setMode(mode) {
            if (mode === "group") modeGroupRadio.value = true;
            else if (mode === "artboard") modeArtboardRadio.value = true;
            else modeAutoPairRadio.value = true;
        }

        return {
            modeRadios: [modeGroupRadio, modeAutoPairRadio, modeArtboardRadio],
            getMode: getMode,
            setMode: setMode
        };
    }

    /**
     * 固定オブジェクト（上 / 左 / 右 / 下）パネルを生成する。
     * 上下左右を厳密な3×3グリッドに配置する（中央・四隅は空セル）。セルを固定幅にして列を確実にそろえる。
     * 4つのラジオは親が別なので自動グループ化されない（排他は手動）。
     * 既定は「右」を固定（左が動く＝水平）。イベント結線は呼び出し側で行う
     * @param {Group} parentGroup - 追加先
     * @returns {{panel: Panel, fixedRadios: RadioButton[], selectFixedRadio: Function, getFixedSide: Function, setFixedSide: Function}} パネルと操作用の関数
     */
    function buildFixedSidePanel(parentGroup) {
        var fixedSidePanel = parentGroup.add("panel", undefined, getLabel("panel.fixedSide"));
        setupPanel(fixedSidePanel, 2);
        // 十字レイアウトの左右に余白を足す（+8）/ Add left/right margin to the cross layout (+8)
        fixedSidePanel.margins = [PANEL_MARGINS[0] + 8, PANEL_MARGINS[1], PANEL_MARGINS[2] + 8, PANEL_MARGINS[3]];
        fixedSidePanel.alignChildren = ["center", "top"]; // 各行を中央そろえで十字に / Center each row

        /**
         * グリッドの1セルを作る。withRadio が真ならラジオを入れて返す
         * @param {Group} gridRow - 追加先の行
         * @param {boolean} withRadio - ラジオを入れるか
         * @returns {RadioButton|null} 入れたラジオ（無ければ null）
         */
        function addGridCell(gridRow, withRadio) {
            var gridCell = gridRow.add("group");
            gridCell.margins = 0;
            gridCell.alignChildren = ["center", "center"];
            gridCell.preferredSize = GRID_CELL_SIZE;
            return withRadio ? gridCell.add("radiobutton", undefined, "") : null;
        }

        /**
         * グリッドの1行（3セル分の入れ物）を作る
         * @returns {Group} 行のグループ
         */
        function addGridRow() {
            var gridRow = fixedSidePanel.add("group");
            gridRow.orientation = "row";
            gridRow.alignment = "center";
            gridRow.spacing = 0;
            gridRow.margins = 0;
            return gridRow;
        }

        //   ・  上  ・
        //   左  ・  右
        //   ・  下  ・
        var gridRowTop = addGridRow();
        addGridCell(gridRowTop, false);
        var topRadio = addGridCell(gridRowTop, true);       // 上 / Top
        addGridCell(gridRowTop, false);

        var gridRowMid = addGridRow();
        var leftRadio = addGridCell(gridRowMid, true);      // 左 / Left
        addGridCell(gridRowMid, false);                     // 中央は空 / center empty
        var rightRadio = addGridCell(gridRowMid, true);     // 右 / Right

        var gridRowBottom = addGridRow();
        addGridCell(gridRowBottom, false);
        var bottomRadio = addGridCell(gridRowBottom, true); // 下 / Bottom
        addGridCell(gridRowBottom, false);

        var fixedRadios = [topRadio, leftRadio, rightRadio, bottomRadio];
        var radioBySide = { top: topRadio, left: leftRadio, right: rightRadio, bottom: bottomRadio };

        /**
         * 指定したラジオだけ ON にして排他制御する
         * @param {RadioButton} chosenRadio - ON にするラジオ
         * @returns {void}
         */
        function selectFixedRadio(chosenRadio) {
            for (var i = 0; i < fixedRadios.length; i++) {
                fixedRadios[i].value = (fixedRadios[i] === chosenRadio);
            }
        }
        for (var i = 0; i < fixedRadios.length; i++) {
            fixedRadios[i].helpTip = getLabel("tooltip.fixedSide");
        }
        rightRadio.value = true; // 既定：右を固定（左が動く＝水平）/ Default: fix right (horizontal)

        /**
         * 現在固定する側を返す
         * @returns {string} "top" / "left" / "right" / "bottom"
         */
        function getFixedSide() {
            if (topRadio.value) return "top";
            if (leftRadio.value) return "left";
            if (bottomRadio.value) return "bottom";
            return "right";
        }

        /**
         * 固定する側を設定する。未知の値は無視
         * @param {string} side - "top" / "left" / "right" / "bottom"
         * @returns {void}
         */
        function setFixedSide(side) {
            if (radioBySide[side]) selectFixedRadio(radioBySide[side]);
        }

        return {
            panel: fixedSidePanel,
            fixedRadios: fixedRadios,
            selectFixedRadio: selectFixedRadio,
            getFixedSide: getFixedSide,
            setFixedSide: setFixedSide
        };
    }

    /**
     * オフセットパネル（間隔の入力・プレビュー境界）を生成する。整列後のずらし量は［位置調整］パネルが
     * 持つので、ここは間隔だけを扱う。間隔は定規の単位で表示し（単位はパネル名に出す）、
     * getSpacingInPoints() が pt に換算して返す。イベント結線は呼び出し側で行う
     * @param {Group} parentGroup - 追加先
     * @param {number} initialGapPoints - 間隔の初期値・空欄時のフォールバック（pt）
     * @returns {{panel: Panel, spacingInput: EditText, previewBoundsCheckbox: Checkbox, getSpacingInPoints: Function, getBoundsType: Function}} パネルと操作用の関数
     */
    function buildGapPanel(parentGroup, initialGapPoints) {
        var gapPanel = parentGroup.add("panel", undefined, labelWithUnit("panel.offset", rulerUnitLabel));
        setupPanel(gapPanel, 6);

        // 間隔行（ラベル＋入力）。単位はパネル名に出しているので行には並べない
        // Gap row (label + input); the unit lives in the panel title instead
        var spacingRow = gapPanel.add("group");
        setupGroup(spacingRow, "row");
        spacingRow.alignment = "left"; // 広げず左寄せ / Keep at natural width, packed left
        spacingRow.add("statictext", undefined, labelText("fieldLabel.spacing"));
        /* ∧∨と入力欄は隙間0で突き合わせる。負の値も許容（オブジェクトを重ねる）/ Stepper butts the field; negatives allowed (overlap) */
        var spacingFieldGroup = spacingRow.add("group");
        spacingFieldGroup.orientation = "row";
        spacingFieldGroup.alignChildren = ["left", "center"];
        spacingFieldGroup.spacing = 0;
        spacingFieldGroup.margins = 0;
        var spacingInput;
        var spacingStepper = addStepper(spacingFieldGroup, function () { return spacingInput; }, {
            onStep: function (numberInput) { numberInput.notify("onChanging"); } /* プレビュー更新 / refresh preview */
        });
        spacingInput = spacingFieldGroup.add("edittext", undefined, pointsToDisplayText(initialGapPoints));
        spacingInput.characters = 4;
        bindSteppedArrowKeys(spacingInput, spacingStepper);
        spacingInput.helpTip = getLabel("tooltip.spacing");

        // チェックボックス：プレビュー境界（左添え）/ Preview-bounds checkbox (left)
        var previewBoundsGroup = gapPanel.add("group");
        previewBoundsGroup.orientation = "row";
        previewBoundsGroup.alignment = "left";
        previewBoundsGroup.margins = [0, 5, 0, 0]; // 上マージン5 / Top margin 5
        var previewBoundsCheckbox = previewBoundsGroup.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.value = false; // OFF=幾何境界 / ON=プレビュー境界
        previewBoundsCheckbox.helpTip = getLabel("tooltip.previewBounds");

        /**
         * 入力値を pt に換算して返す。数値でなければ初期値
         * @returns {number} 間隔（pt）
         */
        function getSpacingInPoints() {
            var value = parseFloat(spacingInput.text);
            if (isNaN(value)) { value = initialGapPoints / pointsPerUnit; }
            return value * pointsPerUnit;
        }

        /**
         * チェックに応じた境界のプロパティ名を返す
         * @returns {string} "visibleBounds" または "geometricBounds"
         */
        function getBoundsType() {
            return previewBoundsCheckbox.value ? "visibleBounds" : "geometricBounds";
        }

        return {
            panel: gapPanel,
            spacingInput: spacingInput,
            previewBoundsCheckbox: previewBoundsCheckbox,
            getSpacingInPoints: getSpacingInPoints,
            getBoundsType: getBoundsType
        };
    }

    /**
     * 位置調整パネルを生成する。整列（なし/開始/中央/終端）と位置の2行を1枚にまとめ、［固定］で選んだ側に
     * 応じて setOrientation() で水平／垂直に切り替える（開始/終端のラベルとツールチップが入れ替わる）。
     * どちらの向きかは 左・右／上・下 のラベルで示すので、パネル名は「位置調整」で固定
     * （オフセットパネルと同じく、単位はパネル名に出す）。
     * 向きごとの値の保持と中央時のオフセット無効化は createOrientationSwitcher() が行う。
     * 開始/終端のラベルはレイアウト確定後に差し替えるので、文字数の多いほうで組み立てて幅を確保しておく
     * @param {Window} parentGroup - 追加先
     * @returns {{alignRadios: Object, offsetRow: Group, offsetInput: EditText, setOrientation: Function}} 整列ラジオ・位置の入力欄と切り替え関数
     */
    function buildAlignmentPanel(parentGroup) {
        /**
         * 文字数の多いほうを返す（ラベル領域の確保用）
         * @param {string} firstText - 候補1
         * @param {string} secondText - 候補2
         * @returns {string} 長いほうの文字列
         */
        function longerText(firstText, secondText) {
            return (firstText.length >= secondText.length) ? firstText : secondText;
        }

        var alignmentPanel = parentGroup.add("panel", undefined, labelWithUnit("panel.position", rulerUnitLabel));
        setupPanel(alignmentPanel, 6);

        // 整列行（ラベル＋なし/開始/中央/終端）/ Alignment row (label + none/start/center/end)
        var alignRow = alignmentPanel.add("group");
        setupGroup(alignRow, "row");
        alignRow.alignment = "left";
        var alignLabel = alignRow.add("statictext", undefined, labelText("fieldLabel.align"));
        var alignRadios = {
            none: alignRow.add("radiobutton", undefined, getLabel("radio.alignNone")),
            start: alignRow.add("radiobutton", undefined,
                longerText(getLabel("radio.alignLeft"), getLabel("radio.alignTop"))),
            center: alignRow.add("radiobutton", undefined, getLabel("radio.alignCenter")),
            end: alignRow.add("radiobutton", undefined,
                longerText(getLabel("radio.alignRight"), getLabel("radio.alignBottom")))
        };
        alignRadios.none.value = true; // 既定：整列なし / Default: no alignment

        // 位置行（ラベル＋入力）。単位はパネル名に出しているので行には並べない
        // Position row (label + input); the unit lives in the panel title instead
        var offsetRow = alignmentPanel.add("group");
        setupGroup(offsetRow, "row");
        offsetRow.alignment = "left";
        var offsetLabel = offsetRow.add("statictext", undefined, labelText("fieldLabel.position"));
        var offsetFieldGroup = offsetRow.add("group");
        offsetFieldGroup.orientation = "row";
        offsetFieldGroup.alignChildren = ["left", "center"];
        offsetFieldGroup.spacing = 0;
        offsetFieldGroup.margins = 0;
        var offsetInput;
        var offsetStepper = addStepper(offsetFieldGroup, function () { return offsetInput; }, {
            onStep: function (numberInput) { numberInput.notify("onChanging"); } /* プレビュー更新 / refresh preview */
        });
        offsetInput = offsetFieldGroup.add("edittext", undefined, "0");
        offsetInput.characters = 4;
        bindSteppedArrowKeys(offsetInput, offsetStepper);

        // ラベル幅をそろえて整列ラジオと入力の開始位置を合わせる / Match label widths
        var labelWidth = Math.max(alignLabel.preferredSize.width, offsetLabel.preferredSize.width);
        alignLabel.preferredSize.width = labelWidth;
        offsetLabel.preferredSize.width = labelWidth;

        /**
         * 水平／垂直を切り替える（開始/終端のラベルとツールチップ）。ラベルの差し替えで
         * レイアウトは組み直さないので、レイアウト確定後（onShow 以降）に呼ぶこと
         * @param {boolean} isHorizontal - 水平の整列にするか
         * @returns {void}
         */
        function setOrientation(isHorizontal) {
            alignRadios.start.text = getLabel(isHorizontal ? "radio.alignLeft" : "radio.alignTop");
            alignRadios.end.text = getLabel(isHorizontal ? "radio.alignRight" : "radio.alignBottom");
            var alignTip = getLabel(isHorizontal ? "tooltip.alignH" : "tooltip.alignV");
            for (var alignKey in alignRadios) { alignRadios[alignKey].helpTip = alignTip; }
            offsetInput.helpTip = getLabel(isHorizontal ? "tooltip.offsetHorizontal" : "tooltip.offsetVertical");
        }

        return {
            alignRadios: alignRadios,
            offsetRow: offsetRow,
            offsetInput: offsetInput,
            setOrientation: setOrientation
        };
    }

    /**
     * 整列ラジオで選択中の値を返す
     * @param {Object} alignRadios - 整列ラジオ（none / start / center / end）
     * @returns {string} "none" / "start" / "center" / "end"
     */
    function getAlignValue(alignRadios) {
        if (alignRadios.start.value) return "start";
        if (alignRadios.center.value) return "center";
        if (alignRadios.end.value) return "end";
        return "none";
    }

    /**
     * 保存済みの整列値をラジオへ反映する。未知の値は無視
     * @param {Object} alignRadios - 整列ラジオ（none / start / center / end）
     * @param {string} alignValue - 整列値
     * @returns {void}
     */
    function applySavedAlign(alignRadios, alignValue) {
        if (!alignRadios || !alignValue) return;
        var targetRadio = alignRadios[alignValue];
        if (targetRadio) targetRadio.value = true;
    }

    /**
     * ［位置調整］パネルの向き（水平／垂直）の切り替えを受け持つ。水平／垂直それぞれの入力内容を控え、
     * ［固定］の側が変わったら退避・復元する。整列「中央」のときはオフセットを 0 にして無効にする。
     * 上下キー（縦並び）→ 水平の整列、左右キー（横並び）→ 垂直の整列
     * @param {Object} fixedSideRefs - buildFixedSidePanel() の戻り値
     * @param {Object} alignmentRefs - buildAlignmentPanel() の戻り値
     * @returns {{orientationValues: Object, isVerticalGap: Function, stashOrientationValues: Function, updateActivePanels: Function}} 向きごとの控えと操作用の関数
     */
    function createOrientationSwitcher(fixedSideRefs, alignmentRefs) {
        var alignRadios = alignmentRefs.alignRadios;
        var offsetInput = alignmentRefs.offsetInput;
        // 水平／垂直それぞれの入力内容。パネルを切り替えるときに退避・復元する
        // Per-orientation values, stashed and restored as the panel switches
        var orientationValues = {
            h: { align: "none", offset: "0" },
            v: { align: "none", offset: "0" }
        };
        var shownOrientation = null; // 今パネルに出ている向き / the orientation currently shown

        /**
         * キーオブジェクトの側からギャップが垂直か（上下キー）を判定する
         * @returns {boolean} 上下をキーにしているなら true
         */
        function isVerticalGap() {
            var side = fixedSideRefs.getFixedSide();
            return side === "top" || side === "bottom";
        }

        /**
         * パネルに出ている内容を、その向きの控えへ退避する
         * @returns {void}
         */
        function stashOrientationValues() {
            if (!shownOrientation) return;
            orientationValues[shownOrientation].align = getAlignValue(alignRadios);
            orientationValues[shownOrientation].offset = offsetInput.text;
        }

        /**
         * キー側に合わせて整列パネルの向きと中身を入れ替え、整列「中央」ならオフセットを 0＋無効にする
         * @returns {void}
         */
        function updateActivePanels() {
            var orientation = isVerticalGap() ? "h" : "v";
            if (orientation !== shownOrientation) {
                stashOrientationValues();
                shownOrientation = orientation;
                alignmentRefs.setOrientation(orientation === "h");
                applySavedAlign(alignRadios, orientationValues[orientation].align);
                offsetInput.text = orientationValues[orientation].offset;
            }
            var isCenter = alignRadios.center.value;
            alignmentRefs.offsetRow.enabled = !isCenter;
            redrawSteppersIn(alignmentRefs.offsetRow); /* ∧∨のディム表示を切り替える / update stepper dimming */
            if (isCenter) offsetInput.text = "0";
        }

        return {
            orientationValues: orientationValues,
            isVerticalGap: isVerticalGap,
            stashOrientationValues: stashOrientationValues,
            updateActivePanels: updateActivePanels
        };
    }

    // =========================================
    // 行揃えボタン / Justification buttons
    // ScriptUI の button では選択状態を表示できないので、背景とアイコンを onDraw で自前描画する。
    // 描画方式は UnifiedTypePalette.jsx にそろえる。
    // ScriptUI buttons cannot show a selected state, so the background and icon are drawn in onDraw;
    // the drawing follows UnifiedTypePalette.jsx.
    // =========================================

    /**
     * 環境設定のUI明るさが明るい側かを返す。
     * uiBrightness は 0（最暗）〜1（最明）。0.5（やや暗め）は暗い側に含めるため 0.5 超で判定
     * @returns {boolean} 明るい側なら true
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * テーマとアクティブ状態に応じたボタンの配色を返す
     * @param {boolean} isLight - 明るいテーマか
     * @param {boolean} isActive - 選択中か
     * @returns {{bg: number[], border: (number[]|null), line: number[]}} 背景・枠・線の色
     */
    function getJustifyColors(isLight, isActive) {
        if (isLight) {
            return {
                bg: isActive ? [0.40, 0.40, 0.40, 1] : [1, 1, 1, 1],
                border: isActive ? [0.30, 0.30, 0.30, 1] : [0.62, 0.62, 0.62, 1],
                line: isActive ? [1, 1, 1, 1] : [0.25, 0.25, 0.25, 1]
            };
        }
        return {
            bg: isActive ? [0.92, 0.92, 0.92, 1] : [0.30, 0.30, 0.30, 1],
            border: null,
            line: isActive ? [0.16, 0.16, 0.16, 1] : [0.82, 0.82, 0.82, 1]
        };
    }

    /**
     * アイコンの行ごとの線幅を返す。均等配置だけ最終行以外を長くする
     * @param {string} iconType - "left" / "center" / "right" / "full"
     * @param {number} longWidth - 長い行の幅
     * @param {number} shortWidth - 短い行の幅
     * @returns {number[]} 4行分の線幅
     */
    function getJustifyLineWidths(iconType, longWidth, shortWidth) {
        if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
        return [longWidth, shortWidth, longWidth, shortWidth];
    }

    /**
     * 行の開始 X（左／中央／右）を返す
     * @param {string} iconType - "left" / "center" / "right" / "full"
     * @param {number} buttonWidth - ボタンの幅
     * @param {number} lineWidth - 線の幅
     * @returns {number} 線の開始 X
     */
    function getJustifyLineX(iconType, buttonWidth, lineWidth) {
        var margin = 5;
        if (iconType === "right") return buttonWidth - margin - lineWidth;
        if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
        return margin;
    }

    /**
     * 行揃えアイコンの罫線を描く
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {string} iconType - "left" / "center" / "right" / "full"
     * @param {number} buttonWidth - ボタンの幅
     * @param {number[]} lineColor - 線の色
     * @returns {void}
     */
    function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var lineYs = [7, 11, 15, 19];
        var lineWidths = getJustifyLineWidths(iconType, 15, 10);
        for (var i = 0; i < lineYs.length; i++) {
            var lineWidth = lineWidths[i];
            var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
            graphics.newPath();
            graphics.moveTo(lineStartX, lineYs[i]);
            graphics.lineTo(lineStartX + lineWidth, lineYs[i]);
            graphics.strokePath(linePen);
        }
    }

    /**
     * 「自動」のアイコンを描く。他の行揃えと同じ高さに収まる「A」を線で描く
     * （drawString はフォントの解決に環境差があり空欄になることがあるため使わない）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} buttonWidth - ボタンの幅
     * @param {number[]} lineColor - 線の色
     * @returns {void}
     */
    function drawAutoIcon(graphics, buttonWidth, lineColor) {
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var centerX = Math.round(buttonWidth / 2);
        var topY = 7;
        var bottomY = 19;
        var halfWidth = 5;
        var crossY = 15; // 横棒。斜線上の位置に合わせて幅を決める / crossbar, width taken from the diagonals
        var crossHalf = Math.round(halfWidth * (crossY - topY) / (bottomY - topY));
        graphics.newPath();
        graphics.moveTo(centerX - halfWidth, bottomY);
        graphics.lineTo(centerX, topY);
        graphics.lineTo(centerX + halfWidth, bottomY);
        graphics.strokePath(linePen);
        graphics.newPath();
        graphics.moveTo(centerX - crossHalf, crossY);
        graphics.lineTo(centerX + crossHalf, crossY);
        graphics.strokePath(linePen);
    }

    /**
     * ボタンの背景とアイコンを描く。描画できない環境では OS 標準のボタン（ラベル付き）に任せる
     * @param {Button} justifyButton - 描くボタン（iconType を持つ）
     * @param {boolean} isActive - 選択中か
     * @param {boolean} isLight - 明るいテーマか
     * @returns {void}
     */
    function drawJustifyButton(justifyButton, isActive, isLight) {
        var graphics = justifyButton.graphics;
        var buttonColors = getJustifyColors(isLight, isActive);
        try {
            graphics.rectPath(0, 0, justifyButton.size[0], justifyButton.size[1]);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, buttonColors.bg));
            if (buttonColors.border) {
                graphics.rectPath(0, 0, justifyButton.size[0], justifyButton.size[1]);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, buttonColors.border, 1));
            }
            if (justifyButton.iconType === "auto") {
                drawAutoIcon(graphics, justifyButton.size[0], buttonColors.line);
            } else {
                drawJustifyIconLines(graphics, justifyButton.iconType, justifyButton.size[0], buttonColors.line);
            }
        } catch (eDraw) {
            // 描画できない環境ではOS標準のボタン（ラベル付き）に任せる / Fall back to the OS control
            try { graphics.drawOSControl(); } catch (eOs) {}
        }
    }

    /**
     * テキストの行揃えパネルを生成する。自動 / 左 / 中央 / 右 / 均等配置（最終行左）をボタンで並べる。
     * 「自動」は整列・キーに連動（エリア内文字は均等配置）。解決は resolveJustification() で行う。
     * イベント結線は呼び出し側で行う
     * @param {Window} parentGroup - 追加先
     * @returns {{panel: Panel, buttons: Button[], getJustifyMode: Function, setJustifyMode: Function}} パネルとボタン、読み書きの関数
     */
    function buildJustifyPanel(parentGroup) {
        var justifyPanel = parentGroup.add("panel", undefined, getLabel("panel.justify"));
        setupPanel(justifyPanel, 4);
        justifyPanel.orientation = "row"; // ボタン横並び / buttons in a row
        justifyPanel.alignChildren = ["center", "center"]; // ボタン列をパネルの左右中央に / Center the button row in the panel

        // アクティブな行揃えと UI 明暗を共有する（onDraw のクロージャから参照）
        // Shared active id + theme, read by the onDraw closures
        var justifyState = { activeId: "auto", isLight: isLightUI() }; // 既定：自動（整列・キーに連動）/ Default: auto
        var justifyOptions = [
            { id: "auto", label: "button.justifyAuto", tip: "tooltip.justifyAuto" },
            { id: "left", label: "button.justifyLeft", tip: "tooltip.justifyLeft" },
            { id: "center", label: "button.justifyCenter", tip: "tooltip.justifyCenter" },
            { id: "right", label: "button.justifyRight", tip: "tooltip.justifyRight" },
            { id: "full", label: "button.justifyFull", tip: "tooltip.justifyFull" }
        ];
        var justifyButtons = [];
        for (var i = 0; i < justifyOptions.length; i++) {
            // ラベルは描画に失敗したときのフォールバック（drawOSControl）でも使う / text is also the fallback label
            var justifyButton = justifyPanel.add("button", undefined, getLabel(justifyOptions[i].label));
            justifyButton.helpTip = getLabel(justifyOptions[i].tip);
            justifyButton.preferredSize = JUSTIFY_BUTTON_SIZE;
            justifyButton.minimumSize = JUSTIFY_BUTTON_SIZE;
            justifyButton.maximumSize = JUSTIFY_BUTTON_SIZE; // 伸ばさない / keep the fixed size
            justifyButton.justifyId = justifyOptions[i].id;
            justifyButton.iconType = justifyOptions[i].id;
            justifyButton.onDraw = function () { drawJustifyButton(this, this.justifyId === justifyState.activeId, justifyState.isLight); };
            justifyButtons.push(justifyButton);
        }

        /**
         * 選択中の行揃えモードを返す
         * @returns {string} "auto" / "left" / "center" / "right" / "full"
         */
        function getJustifyMode() {
            return justifyState.activeId;
        }

        /**
         * 行揃えモードを設定してボタンを描き直す。未知の値は無視
         * @param {string} mode - "auto" / "left" / "center" / "right" / "full"
         * @returns {void}
         */
        function setJustifyMode(mode) {
            var found = false;
            for (var i = 0; i < justifyButtons.length; i++) {
                if (justifyButtons[i].justifyId === mode) { found = true; break; }
            }
            if (!found) return;
            justifyState.activeId = mode;
            for (var j = 0; j < justifyButtons.length; j++) {
                try { justifyButtons[j].notify("onDraw"); } catch (eDraw) {}
            }
            // notify だけでは画面に反映されないことがあるので、ウィンドウの再描画も要求する
            // notify alone may not reach the screen, so ask the window to repaint as well
            try { justifyPanel.window.update(); } catch (eUpdate) {}
        }

        return { panel: justifyPanel, buttons: justifyButtons, getJustifyMode: getJustifyMode, setJustifyMode: setJustifyMode };
    }

    /**
     * 設定ダイアログを組み立てる（イベント結線と値の復元は呼び出し側で行う）
     * @param {number} initialGapPoints - 間隔の初期値（pt）
     * @returns {{settingsDialog: Window, modeRefs: Object, fixedSideRefs: Object, gapRefs: Object, alignmentRefs: Object, justifyRefs: Object, btnCancel: Button, btnOK: Button}} ダイアログと各パネルの参照
     */
    function buildSettingsDialog(initialGapPoints) {
        var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        settingsDialog.orientation = "column";
        settingsDialog.alignChildren = "fill";

        // モード（1カラム・ラジオ縦並び）/ Mode (single column, radios stacked)
        var modeRefs = buildModePanel(settingsDialog);

        // キーオブジェクト と オフセット を2カラムで左右に並べる / Key object + Offset side by side (two columns)
        var keyPositionColumns = settingsDialog.add("group");
        keyPositionColumns.orientation = "row";
        keyPositionColumns.alignChildren = ["fill", "fill"]; // 2パネルの高さをそろえる / Match panel heights
        keyPositionColumns.spacing = PANEL_SPACING;

        // キーオブジェクト（上・左・右・下を十字に配置）/ Key object (arranged as a cross)
        var fixedSideRefs = buildFixedSidePanel(keyPositionColumns);
        // オフセット（間隔・プレビュー境界）/ Offset (gap, preview bounds)
        var gapRefs = buildGapPanel(keyPositionColumns, initialGapPoints);

        // キー／位置の2パネルの高さをそろえる / Match the two panel heights
        fixedSideRefs.panel.alignment = ["fill", "fill"];
        gapRefs.panel.alignment = ["fill", "fill"];

        // 位置調整パネル（整列＋位置）。1枚で、［固定］に応じて水平／垂直に切り替わる
        // One Position panel (alignment + offset); it switches with the Key Object side
        var alignmentRefs = buildAlignmentPanel(settingsDialog);

        // テキストの行揃え（自動 / 左 / 中央 / 右 / 均等配置）/ Text alignment
        var justifyRefs = buildJustifyPanel(settingsDialog);

        // ボタン（Mac 規約：Cancel → OK）/ Buttons (Mac order: Cancel → OK)
        var buttonRow = addButtonRow(settingsDialog, { centered: true }); // ボタンをダイアログの左右中央に / Center the buttons in the dialog
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, "OK", { name: "ok" });
        // 行揃えのボタンが増えたので Enter / ESC の行き先を明示する
        // Spell out where Enter / ESC go, now that the justification buttons are pushbuttons too
        settingsDialog.defaultElement = btnOK;
        settingsDialog.cancelElement = btnCancel;

        return {
            settingsDialog: settingsDialog,
            modeRefs: modeRefs,
            fixedSideRefs: fixedSideRefs,
            gapRefs: gapRefs,
            alignmentRefs: alignmentRefs,
            justifyRefs: justifyRefs,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    // =========================================
    // 軸 / Axis
    // =========================================

    /**
     * 水平 / 垂直を共通の「進行方向の座標 p」に抽象化する。p は増えるほど右（水平）または下（垂直）。
     * これで間隔調整のロジックを 1 本で両軸に使える。geometricBounds = [左, 上, 右, 下]。
     * start = 進行方向の先頭側の辺、end = 後ろ側の辺。
     * 垂直は Illustrator の y が上ほど大きいので符号を反転して「下方向で増加」にそろえる
     * @param {boolean} isVertical - 垂直の軸にするか
     * @returns {{vertical: boolean, start: Function, end: Function}} 軸（境界から先頭辺・後ろ辺を取り出す関数）
     */
    function makeAxis(isVertical) {
        if (isVertical) {
            return {
                vertical: true,
                start: function (bounds) { return -bounds[1]; }, // 上端 / top edge
                end: function (bounds) { return -bounds[3]; }  // 下端 / bottom edge
            };
        }
        return {
            vertical: false,
            start: function (bounds) { return bounds[0]; }, // 左端 / left edge
            end: function (bounds) { return bounds[2]; }  // 右端 / right edge
        };
    }

    /**
     * 固定する側から軸を求める（上下→垂直、左右→水平）
     * @param {string} fixedSide - "top" / "left" / "right" / "bottom"
     * @returns {Object} makeAxis() の軸
     */
    function axisForSide(fixedSide) {
        return makeAxis(fixedSide === "top" || fixedSide === "bottom");
    }

    /**
     * 固定側が進行方向の「後ろ側」(右 or 下) かどうかを返す
     * @param {string} fixedSide - "top" / "left" / "right" / "bottom"
     * @returns {boolean} 右・下なら true
     */
    function isAnchorEnd(fixedSide) {
        return fixedSide === "right" || fixedSide === "bottom";
    }

    // =========================================
    // 行揃え / Text justification
    // =========================================

    /**
     * 左右方向の整列値を段落の行揃えへ対応づける（start=左揃え、center=中央揃え、end=右揃え）
     * @param {string} alignMode - "none" / "start" / "center" / "end"
     * @returns {Justification|null} 行揃え。none・不明は null（行揃えを変更しない）
     */
    function justificationForAlign(alignMode) {
        if (alignMode === "start") return Justification.LEFT;
        if (alignMode === "center") return Justification.CENTER;
        if (alignMode === "end") return Justification.RIGHT;
        return null;
    }

    /**
     * 「テキストの行揃え」パネルの選択を、オブジェクトごとの行揃えへ解決する。
     * "left"/"center"/"right"/"full" は一律。"auto" は連動：エリア内文字＝均等配置、
     * ポイント文字＝縦並びなら水平整列に連動（左/中央/右）、横並びならキー側に連動（左→左/右→右）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {Object} spacingContext - applySpacing() の設定（justifyMode / alignAxis / alignMode / fixedSide）
     * @returns {Justification|null} 行揃え。null は変更しない
     */
    function resolveJustification(targetItem, spacingContext) {
        var justifyMode = spacingContext.justifyMode;
        if (justifyMode === "left") return Justification.LEFT;
        if (justifyMode === "center") return Justification.CENTER;
        if (justifyMode === "right") return Justification.RIGHT;
        if (justifyMode === "full") return Justification.FULLJUSTIFYLASTLINELEFT;
        // auto（連動）/ auto (linked)
        if (targetItem.constructor.name === "TextFrame" && targetItem.kind === TextType.AREATEXT) {
            return Justification.FULLJUSTIFYLASTLINELEFT; // エリア内文字は均等配置 / area text → justify
        }
        if (!spacingContext.alignAxis.vertical) return justificationForAlign(spacingContext.alignMode); // 縦並び：水平整列に連動
        var fixedSide = spacingContext.fixedSide;
        return (fixedSide === "left") ? Justification.LEFT
            : (fixedSide === "right") ? Justification.RIGHT : null;        // 横並び：キー側に連動
    }

    /**
     * テキストの段落行揃えを設定する。ポイント文字は行揃えを変えるとアンカー基準で組み直されて
     * フレームが動く。見た目の位置は間隔・整列の計算で決めるので、どの値でも元の位置へ戻す
     * （戻さないと行揃えを往復するたびに字幅の半分ずつずれ、キャンセルしても戻らない）。
     * Justification.LEFT は代入が無視される Illustrator のバグがあるので、一時 resize（200%→50%）で
     * 段落属性をリフレッシュしてから代入する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Justification} justification - 行揃え
     * @returns {void}
     */
    function setParagraphJustification(textFrame, justification) {
        var savedPosition = [textFrame.position[0], textFrame.position[1]];
        if (justification === Justification.LEFT) {
            textFrame.resize(200, 200);
            textFrame.textRange.paragraphAttributes.justification = Justification.LEFT;
            textFrame.resize(50, 50);
        } else {
            textFrame.textRange.paragraphAttributes.justification = justification;
        }
        try { textFrame.position = savedPosition; } catch (ePos) {}
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
    // ペアリング / Pairing
    // =========================================

    /**
     * 選択オブジェクトを最も近いもの同士でペアにする。
     * ペアリングは常に geometricBounds（中心）で判定する。「プレビュー境界」を ON にしても
     * 間隔計算の基準が変わるだけで、ペアの組み合わせ自体は変わらない
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {Object[]} ペア（{ a, b }）の配列
     */
    function createNearestPairs(selectedItems) {
        /**
         * オブジェクトの中心座標を返す（常に幾何境界。クリップグループはクリッピングパス基準）
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {{x: number, y: number}} 中心座標
         */
        function getObjectCenter(item) {
            var bounds = getClipAwareBounds(item, false);
            return {
                x: (bounds[0] + bounds[2]) / 2,
                y: (bounds[1] + bounds[3]) / 2
            };
        }

        /**
         * 2点間の距離を返す
         * @param {{x: number, y: number}} point1 - 点1
         * @param {{x: number, y: number}} point2 - 点2
         * @returns {number} 距離
         */
        function getPointDistance(point1, point2) {
            var dx = point1.x - point2.x;
            var dy = point1.y - point2.y;
            return Math.sqrt(dx * dx + dy * dy);
        }

        // 中心座標付きの作業リスト / Working list with centers
        var remainingItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            remainingItems.push({
                item: selectedItems[i],
                center: getObjectCenter(selectedItems[i])
            });
        }

        var pairs = [];
        while (remainingItems.length > 0) {
            var currentItem = remainingItems.shift();
            var nearestIndex = -1;
            var nearestDistance = Infinity;
            for (var j = 0; j < remainingItems.length; j++) {
                var distance = getPointDistance(currentItem.center, remainingItems[j].center);
                if (distance < nearestDistance) {
                    nearestDistance = distance;
                    nearestIndex = j;
                }
            }
            if (nearestIndex !== -1) {
                var partnerItem = remainingItems.splice(nearestIndex, 1)[0];
                pairs.push({ a: currentItem.item, b: partnerItem.item });
            }
        }
        return pairs;
    }

    /**
     * GroupItem の直下のオブジェクトを配列で返す
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {PageItem[]|null} 直下のオブジェクト。グループ以外は null
     */
    function getGroupChildren(item) {
        if (item.typename !== "GroupItem") return null;
        var groupChildren = [];
        for (var i = 0; i < item.pageItems.length; i++) {
            groupChildren.push(item.pageItems[i]);
        }
        return groupChildren;
    }

    /**
     * 選択した各グループの中身を1単位にする。グループは2つに限らない。含まれる全オブジェクトを
     * members として持ち、applySpacing 側で先頭辺の順に等間隔へ分配する（固定側を基準に配置）
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {Object[]} 単位（{ members }）の配列。子が2個未満のグループとグループ以外は含まない
     */
    function createGroupPairs(selectedItems) {
        var pairs = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var groupChildren = getGroupChildren(selectedItems[i]);
            // グループでない、または子が2個未満なら間隔を調整できない / Need a group with 2+ children
            if (!groupChildren || groupChildren.length < 2) continue;
            pairs.push({ members: groupChildren });
        }
        return pairs;
    }

    /**
     * オブジェクトの中心を含むアートボードの矩形を返す。該当が無ければアクティブアートボードを使う。
     * geometricBounds と同じ並び（上 > 下、y は上方向で増加）なので軸関数で扱える
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]} artboardRect [左, 上, 右, 下]
     */
    function artboardRectFor(item) {
        var artboards = app.activeDocument.artboards;
        var bounds = getClipAwareBounds(item, false);
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        for (var i = 0; i < artboards.length; i++) {
            var artboardRect = artboards[i].artboardRect; // [左, 上, 右, 下] / [left, top, right, bottom]
            if (centerX >= artboardRect[0] && centerX <= artboardRect[2] &&
                centerY <= artboardRect[1] && centerY >= artboardRect[3]) {
                return artboardRect;
            }
        }
        return artboards[artboards.getActiveArtboardIndex()].artboardRect;
    }

    /**
     * アートボードモードの作業単位を作る。選択オブジェクトを1つずつ独立した単位にし、
     * applySpacing 側で各オブジェクトとアートボード端の間隔（マージン）を指定値にそろえる
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {Object[]} 単位（{ single }）の配列
     */
    function createArtboardUnits(selectedItems) {
        var artboardUnits = [];
        for (var i = 0; i < selectedItems.length; i++) {
            artboardUnits.push({ single: selectedItems[i] });
        }
        return artboardUnits;
    }

    /**
     * 作業単位に含まれるオブジェクトを配列で返す（アートボード：1個、グループ：全メンバー、ペア：2個）
     * @param {Object} pair - 作業単位（{ single } / { members } / { a, b }）
     * @returns {PageItem[]} オブジェクトの配列
     */
    function getPairItems(pair) {
        if (pair.single) return [pair.single];
        if (pair.members) return pair.members;
        return [pair.a, pair.b];
    }

    /**
     * 選択オブジェクトの現在の間隔の平均（pt）を求める。
     * モードに合わせて測る：グループは各グループ内の隣接間隔、アートボードは固定側の端との距離
     * （マージン）、自動ペア認識は各ペアの間隔。測る軸と固定端は fixedSide から決める。
     * 常に geometricBounds（幾何境界）基準
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @param {string} mode - "group" / "artboard" / "auto"
     * @param {string} fixedSide - "top" / "left" / "right" / "bottom"
     * @returns {number|null} 平均間隔（pt）。測れない場合は null
     */
    function computeAverageGap(selectedItems, mode, fixedSide) {
        var axis = axisForSide(fixedSide);
        var gaps = [];

        if (mode === "group") {
            // 各グループ内：進行方向順に並べて隣り合う間隔を測る / Adjacent gaps inside each group
            for (var i = 0; i < selectedItems.length; i++) {
                var groupChildren = getGroupChildren(selectedItems[i]);
                if (!groupChildren || groupChildren.length < 2) continue;
                var sortedChildren = groupChildren.slice(0);
                sortedChildren.sort(function (itemA, itemB) {
                    return axis.start(getClipAwareBounds(itemA, false)) - axis.start(getClipAwareBounds(itemB, false));
                });
                for (var j = 1; j < sortedChildren.length; j++) {
                    // 次の先頭辺 - 前の後ろ辺 / next leading edge - previous trailing edge
                    gaps.push(axis.start(getClipAwareBounds(sortedChildren[j], false)) - axis.end(getClipAwareBounds(sortedChildren[j - 1], false)));
                }
            }
        } else if (mode === "artboard") {
            // アートボード：各オブジェクトと固定側のアートボード端との距離（マージン）を測る
            // Artboard: distance (margin) from each object to the fixed artboard edge
            var anchorEnd = isAnchorEnd(fixedSide);
            for (var i = 0; i < selectedItems.length; i++) {
                var itemBounds = getClipAwareBounds(selectedItems[i], false);
                var artboardRect = artboardRectFor(selectedItems[i]);
                gaps.push(anchorEnd ? axis.end(artboardRect) - axis.end(itemBounds)
                    : axis.start(itemBounds) - axis.start(artboardRect));
            }
        } else {
            // 自動ペア認識：各ペアの間隔を測る / Gap of each nearest pair
            var nearestPairs = createNearestPairs(selectedItems);
            for (var i = 0; i < nearestPairs.length; i++) {
                var boundsA = getClipAwareBounds(nearestPairs[i].a, false);
                var boundsB = getClipAwareBounds(nearestPairs[i].b, false);
                var isALeading = axis.start(boundsA) < axis.start(boundsB);
                var leadBounds = isALeading ? boundsA : boundsB;
                var trailBounds = isALeading ? boundsB : boundsA;
                gaps.push(axis.start(trailBounds) - axis.end(leadBounds));
            }
        }

        if (gaps.length === 0) return null;
        var gapSum = 0;
        for (var i = 0; i < gaps.length; i++) gapSum += gaps[i];
        return gapSum / gaps.length;
    }

    // =========================================
    // 設定の保存 / Persistence
    // =========================================
    // OK で設定を覚え、次回開いたときに復元する
    // （モード・キー・プレビュー境界・整列・行揃え・間隔・左右/上下オフセット。数値は pt で保存）。
    // 保存先は2つ：
    //   - 再起動しても残す（Folder.userData/illustrator-scripts/AdjustPairGap.json）… キー・プレビュー境界・整列・行揃え・オフセット
    //   - Illustrator の終了まで（$.global）… モードと間隔。選択内容に依存するので再起動後は毎回選択から決め直す
    // Remember the settings on OK and restore them next time (mode, key side, preview-bounds,
    // alignment, justification, gap, and the horizontal/vertical offsets; numbers stored in pt).
    // Two stores: a file under Folder.userData/illustrator-scripts that survives restarts, and $.global
    // (until Illustrator quits) for mode and gap, which depend on the selection.

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

    /* 再起動しても残す設定。旧 AdjustPairGapSettings.txt は新しい保存が無いときだけ読み継ぐ
       Settings kept across restarts; the old AdjustPairGapSettings.txt is read only until the first save */
    var persistentStore = createSettingsStore(SCRIPT_NAME, "persistent", {
        legacy: function () {
            return readSettingsLegacyFile(Folder.userData + "/AdjustPairGapSettings.txt");
        }
    });
    /* Illustrator の終了まで覚える設定（モード・間隔）/ Settings kept until Illustrator quits (mode, gap) */
    var sessionStore = createSettingsStore(SCRIPT_NAME + "Session", "session");

    /* 既定値。空文字は「保存なし」（選択から決める・その欄は初期値のまま）/ Defaults; "" means nothing saved */
    var DEFAULT_PERSISTENT_SETTINGS = {
        fixedSide: "",
        previewBounds: false,
        alignH: "",
        alignV: "",
        justify: "",
        offsetH: "",  /* pt の文字列 / points as text */
        offsetV: ""   /* pt の文字列 / points as text */
    };
    var DEFAULT_SESSION_SETTINGS = {
        mode: "",
        gap: ""       /* pt の文字列 / points as text */
    };

    /**
     * 設定を読み込む（再起動しても残す設定に、同一セッションのモード・間隔を重ねる）
     * @returns {Object} 設定（保存が無い項目は既定値）
     */
    function loadSettings() {
        var settings = persistentStore.load(DEFAULT_PERSISTENT_SETTINGS);
        var sessionSettings = sessionStore.load(DEFAULT_SESSION_SETTINGS);
        settings.mode = sessionSettings.mode;
        settings.gap = sessionSettings.gap;
        return settings;
    }

    /**
     * 設定を保存する。モード・間隔は同一セッション用、それ以外はファイルへ
     * @param {Object} settings - 保存する設定
     * @returns {void}
     */
    function saveSettings(settings) {
        sessionStore.save({ mode: settings.mode, gap: settings.gap });
        persistentStore.save({
            fixedSide: settings.fixedSide,
            previewBounds: settings.previewBounds,
            alignH: settings.alignH,
            alignV: settings.alignV,
            justify: settings.justify,
            offsetH: settings.offsetH,
            offsetV: settings.offsetV
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    (function () {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        // doc.selection を配列にコピーしておく（後の選択変更や undo の影響を受けないように）
        // Copy doc.selection into a plain array (immune to later selection changes / undo)
        var liveSelection = doc.selection;
        // selection が無効・未選択のケースを先に弾く（Illustrator では稀に null になる）
        // Guard against an invalid / empty selection (selection can be null in rare cases)
        if (!liveSelection || liveSelection.length === 0) {
            alert(getLabel("alert.selectTwo"));
            return;
        }
        var selectedItems = [];
        for (var i = 0; i < liveSelection.length; i++) {
            selectedItems.push(liveSelection[i]);
        }

        if (selectedItems.length < 2) {
            alert(getLabel("alert.selectTwo"));
            return;
        }

        // モード切り替え時に組み直すペアを保持する / Holds pairs rebuilt when the mode changes
        var objectPairs = [];

        // 前回終了時の設定。初期間隔を「復元後のモード・キー側」で測るため、ダイアログ生成より先に読む。
        // Last-used settings, loaded before building the dialog so the initial gap can be measured
        // with the mode and key side that will actually be restored.
        var savedSettings = loadSettings();

        // 選択がすべてグループなら既定でグループモードにする（それ以外は自動ペア認識）。
        // Default to group mode when every selected object is a group; otherwise auto pair.
        var selectionIsGroupsOnly = true;
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename !== "GroupItem") { selectionIsGroupsOnly = false; break; }
        }
        var defaultMode = selectionIsGroupsOnly ? "group" : "auto";

        // 実際に開くときのモードとキー側（保存値が有効ならそちらを優先）。
        // 固定側はファイルに永続する一方で間隔は永続しないので、再起動直後は必ずここで測り直される。
        // Mode and key side the dialog actually opens with (a valid saved value wins). The key side
        // persists across restarts while the gap does not, so the gap is always re-measured here.
        var initialMode = (savedSettings.mode === "group" || savedSettings.mode === "auto" ||
            savedSettings.mode === "artboard") ? savedSettings.mode : defaultMode;
        var initialFixedSide = (savedSettings.fixedSide === "top" || savedSettings.fixedSide === "left" ||
            savedSettings.fixedSide === "right" || savedSettings.fixedSide === "bottom")
            ? savedSettings.fixedSide : "right";

        // 間隔の初期値は選択オブジェクトの現在の平均間隔。測れなければ DEFAULT_GAP を使う。
        // 負（重なり）の場合は 0 にクランプ。Initial gap = current average gap of the selection
        // (clamped to >= 0); falls back to DEFAULT_GAP when nothing measurable.
        var measuredGap = computeAverageGap(selectedItems, initialMode, initialFixedSide);
        var initialGapPoints = (measuredGap !== null) ? Math.max(0, measuredGap) : DEFAULT_GAP;

        // =========================================
        // プレビュー / Live preview
        //   巻き戻しは移動量の逆適用で行う（app.undo() は使わない）。
        //   app.undo() をキーボードイベント内で同期実行すると Illustrator が
        //   不安定になり得るため、ここでは記録した移動を translate で打ち消す。
        //   Revert by reversing recorded moves — never app.undo(), which is
        //   unstable when called synchronously inside keyboard event handlers.
        // =========================================
        var appliedMoves = []; // 適用済みの移動 / Applied moves: { targetItem, dx, dy }
        var appliedJustifications = []; // 適用済みの行揃え変更 / Applied justification changes: { targetItem, original }
        // 行揃えの変更で境界が変わったか。プレビューは毎回元位置へ巻き戻すので、境界が変わるのは
        // 行揃えを当てた／戻したときだけ。真のときだけ測り直す（全オブジェクトの visibleBounds
        // 再取得は重く、間隔欄の1打鍵ごとに走ると効いてくる）。
        // Whether justification changed the bounds. The preview always reverts to the original
        // positions, so only applying/reverting justification invalidates them; re-measure only then
        // (re-reading visibleBounds for every object on each keystroke is expensive).
        var boundsCacheStale = false;

        /**
         * 記録した移動を逆向きに適用して元に戻す（位置のみ）。
         * 行揃えはリフレッシュ毎に巻き戻すと resize が連発して重いので、ここでは触らない。
         * 行揃えの巻き戻しは整列系ラジオの変更時とキャンセル時にだけ undoJustifications() で行う
         * @returns {void}
         */
        function undoPreview() {
            for (var i = appliedMoves.length - 1; i >= 0; i--) {
                appliedMoves[i].targetItem.translate(-appliedMoves[i].dx, -appliedMoves[i].dy);
            }
            appliedMoves = [];
        }

        /**
         * テキストの行揃えを justification にそろえ、元の値を記録する（巻き戻し用）。
         * 既に同じなら何もしない（無駄な undo を作らない）。テキスト以外は無視
         * @param {PageItem} targetItem - 対象のオブジェクト
         * @param {Justification|null} justification - 行揃え（null は変更しない）
         * @returns {void}
         */
        function applyJustification(targetItem, justification) {
            if (justification === null) return;
            if (targetItem.constructor.name !== "TextFrame") return;
            var currentJustification = targetItem.textRange.paragraphAttributes.justification;
            if (currentJustification === justification) return;
            appliedJustifications.push({ targetItem: targetItem, original: currentJustification });
            setParagraphJustification(targetItem, justification);
            boundsCacheStale = true;
        }

        /**
         * 記録した行揃え変更を元の値へ戻す
         * @returns {void}
         */
        function undoJustifications() {
            if (appliedJustifications.length === 0) return;
            for (var j = appliedJustifications.length - 1; j >= 0; j--) {
                setParagraphJustification(appliedJustifications[j].targetItem, appliedJustifications[j].original);
            }
            appliedJustifications = [];
            boundsCacheStale = true;
        }

        /**
         * 軸に沿って移動し、巻き戻し用に記録する。移動量は右（水平）/ 下（垂直）で正
         * @param {Object} axis - makeAxis() の軸
         * @param {PageItem} targetItem - 動かすオブジェクト
         * @param {number} travelAmount - 進行方向の移動量（pt）
         * @returns {void}
         */
        function moveByAxis(axis, targetItem, travelAmount) {
            if (travelAmount === 0) return;
            // 垂直は Illustrator の y が上ほど大きいので、下方向(+)へは translate(0, -移動量)
            // Vertical: Illustrator y grows upward, so moving down (+) means translate(0, -amount)
            var dx = axis.vertical ? 0 : travelAmount;
            var dy = axis.vertical ? -travelAmount : 0;
            targetItem.translate(dx, dy);
            appliedMoves.push({ targetItem: targetItem, dx: dx, dy: dy });
        }

        /**
         * グループのメンバー境界のキャッシュを選ぶ（プレビュー境界 / 幾何境界）
         * @param {Object} pair - グループの作業単位
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @returns {number[][]} メンバーごとの境界
         */
        function getMemberBounds(pair, useVisible) {
            return useVisible ? pair.memberVis : pair.memberGeo;
        }

        /**
         * 境界を進行方向（先頭辺の昇順）に並べたインデックス配列を返す
         * @param {number[][]} boundsList - メンバーごとの境界
         * @param {Object} axis - makeAxis() の軸
         * @returns {number[]} インデックスの配列
         */
        function getMemberOrder(boundsList, axis) {
            var memberOrder = [];
            for (var i = 0; i < boundsList.length; i++) memberOrder.push(i);
            memberOrder.sort(function (indexA, indexB) { return axis.start(boundsList[indexA]) - axis.start(boundsList[indexB]); });
            return memberOrder;
        }

        /**
         * グループ内の全オブジェクトを進行方向順に並べ、隣り合う間隔を gap にそろえる。
         * 固定側のオブジェクトは動かさず、そこを起点にカスケードで再配置する。
         * axis で水平/垂直を切り替える（左右→水平、上下→垂直）
         * @param {Object} pair - グループの作業単位
         * @param {string} fixedSide - 固定する側
         * @param {number} gapInPoints - 間隔（pt）
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @param {Object} axis - 間隔の軸
         * @returns {void}
         */
        function distributeGroup(pair, fixedSide, gapInPoints, useVisible, axis) {
            var members = pair.members;
            var cachedBounds = getMemberBounds(pair, useVisible);
            var count = members.length;
            var anchorEnd = isAnchorEnd(fixedSide);
            var memberOrder = getMemberOrder(cachedBounds, axis); // 先頭辺の昇順 / ordered by leading edge

            if (anchorEnd) {
                // 後ろ側（右 or 下）を固定し、後ろから前へ配置 / Anchor the trailing end; walk backward
                var nextLeadingEdge = axis.start(cachedBounds[memberOrder[count - 1]]); // 後ろ端オブジェクトの先頭辺（不動）
                for (var i = count - 2; i >= 0; i--) {
                    var memberIndex = memberOrder[i];
                    var desiredEnd = nextLeadingEdge - gapInPoints;
                    var shiftAmount = desiredEnd - axis.end(cachedBounds[memberIndex]);
                    moveByAxis(axis, members[memberIndex], shiftAmount);
                    nextLeadingEdge = axis.start(cachedBounds[memberIndex]) + shiftAmount; // この要素の新しい先頭辺 / its new leading edge
                }
            } else {
                // 先頭側（左 or 上）を固定し、前から後ろへ配置 / Anchor the leading end; walk forward
                var prevTrailingEdge = axis.end(cachedBounds[memberOrder[0]]); // 先頭端オブジェクトの後ろ辺（不動）
                for (var i = 1; i < count; i++) {
                    var memberIndex = memberOrder[i];
                    var desiredStart = prevTrailingEdge + gapInPoints;
                    var shiftAmount = desiredStart - axis.start(cachedBounds[memberIndex]);
                    moveByAxis(axis, members[memberIndex], shiftAmount);
                    prevTrailingEdge = axis.end(cachedBounds[memberIndex]) + shiftAmount; // この要素の新しい後ろ辺 / its new trailing edge
                }
            }
        }

        /**
         * オブジェクトを alignAxis 方向で基準にそろえる。start = 先頭辺（左 or 上）、end = 後ろ辺（右 or 下）、
         * center = 中央。ギャップ調整は gap 軸方向のみ動かすので、整列軸の位置はキャッシュ境界のまま使える
         * @param {Object} alignAxis - 整列の軸
         * @param {PageItem} targetItem - 動かすオブジェクト
         * @param {number[]} itemBounds - 動かすオブジェクトの境界
         * @param {number[]} anchorBounds - 基準の境界
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @returns {void}
         */
        function alignToAnchor(alignAxis, targetItem, itemBounds, anchorBounds, alignMode) {
            if (alignMode === "none") return;
            var alignShift;
            if (alignMode === "start") {
                alignShift = alignAxis.start(anchorBounds) - alignAxis.start(itemBounds);
            } else if (alignMode === "end") {
                alignShift = alignAxis.end(anchorBounds) - alignAxis.end(itemBounds);
            } else { // center
                var anchorCenter = (alignAxis.start(anchorBounds) + alignAxis.end(anchorBounds)) / 2;
                var itemCenter = (alignAxis.start(itemBounds) + alignAxis.end(itemBounds)) / 2;
                alignShift = anchorCenter - itemCenter;
            }
            moveByAxis(alignAxis, targetItem, alignShift);
        }

        /**
         * グループの固定端メンバー（キーオブジェクト）のインデックスを返す。並び順は不変なので
         * キャッシュ境界で判定して十分（ライブ境界を読まない）
         * @param {Object} pair - グループの作業単位
         * @param {Object} gapAxis - 間隔の軸
         * @param {boolean} anchorEnd - 後ろ側を固定するか
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @returns {number} メンバーのインデックス
         */
        function groupAnchorIndex(pair, gapAxis, anchorEnd, useVisible) {
            var memberOrder = getMemberOrder(getMemberBounds(pair, useVisible), gapAxis);
            return anchorEnd ? memberOrder[memberOrder.length - 1] : memberOrder[0];
        }

        /**
         * グループの各メンバーを固定端のメンバー（アンカー）に整列軸方向でそろえる。
         * 行揃えは間隔調整より前に適用して境界を取り直す（cachePairBounds）ので、ここはキャッシュ境界で十分
         * @param {Object} pair - グループの作業単位
         * @param {Object} gapAxis - 間隔の軸
         * @param {boolean} anchorEnd - 後ろ側を固定するか
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @param {Object} alignAxis - 整列の軸
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @returns {void}
         */
        function alignGroup(pair, gapAxis, anchorEnd, useVisible, alignAxis, alignMode) {
            if (alignMode === "none") return;
            var members = pair.members;
            var cachedBounds = getMemberBounds(pair, useVisible);
            var anchorIndex = groupAnchorIndex(pair, gapAxis, anchorEnd, useVisible); // 固定端のメンバー / member at the fixed end
            var anchorBounds = cachedBounds[anchorIndex];
            for (var i = 0; i < members.length; i++) {
                if (i === anchorIndex) continue;
                alignToAnchor(alignAxis, members[i], cachedBounds[i], anchorBounds, alignMode);
            }
        }

        /**
         * アートボードモード：オブジェクトとアートボード端の間隔（マージン）を指定値にそろえ、
         * アートボードを基準に整列し、位置オフセットをかける（アートボード基準なので常に対象）
         * @param {Object} pair - アートボードの作業単位（{ single }）
         * @param {Object} spacingContext - applySpacing() の設定
         * @returns {void}
         */
        function applyArtboardMargin(pair, spacingContext) {
            var axis = spacingContext.axis;
            var gapInPoints = spacingContext.gapInPoints;
            var singleBounds = spacingContext.useVisible ? pair.visibleS : pair.geometricS;
            var artboardRect = pair.artboardRect;
            var marginShift = spacingContext.anchorEnd
                ? (axis.end(artboardRect) - gapInPoints) - axis.end(singleBounds)    // 右/下端から内側へ / inset from trailing edge
                : (axis.start(artboardRect) + gapInPoints) - axis.start(singleBounds); // 左/上端から内側へ / inset from leading edge
            moveByAxis(axis, pair.single, marginShift);
            if (spacingContext.alignMode !== "none") {
                alignToAnchor(spacingContext.alignAxis, pair.single, singleBounds, artboardRect, spacingContext.alignMode);
            }
            moveByAxis(spacingContext.alignAxis, pair.single, spacingContext.offsetAlong); // 位置オフセット / offset
        }

        /**
         * グループモード：全オブジェクトを等間隔に分配（固定側を基準）し、整列してから、
         * キーオブジェクト（固定端メンバー）以外を直交方向へずらす
         * @param {Object} pair - グループの作業単位（{ members }）
         * @param {Object} spacingContext - applySpacing() の設定
         * @returns {void}
         */
        function applyGroupSpacing(pair, spacingContext) {
            var axis = spacingContext.axis;
            var anchorEnd = spacingContext.anchorEnd;
            var useVisible = spacingContext.useVisible;
            distributeGroup(pair, spacingContext.fixedSide, spacingContext.gapInPoints, useVisible, axis);
            alignGroup(pair, axis, anchorEnd, useVisible, spacingContext.alignAxis, spacingContext.alignMode);
            // 位置オフセット：キーオブジェクト（固定端メンバー）以外を直交方向へずらす / Offset non-key members only
            if (spacingContext.offsetAlong !== 0) {
                var keyMemberIndex = groupAnchorIndex(pair, axis, anchorEnd, useVisible);
                for (var j = 0; j < pair.members.length; j++) {
                    if (j === keyMemberIndex) continue;
                    moveByAxis(spacingContext.alignAxis, pair.members[j], spacingContext.offsetAlong);
                }
            }
        }

        /**
         * 自動ペア認識：2オブジェクトの間隔を調整する（行揃え後の境界）。固定側を基準に動く側だけ移動し、
         * 同じ動く側を整列軸でも固定側へそろえ、位置オフセットをかける
         * @param {Object} pair - ペアの作業単位（{ a, b }）
         * @param {Object} spacingContext - applySpacing() の設定
         * @returns {void}
         */
        function applyPairGap(pair, spacingContext) {
            var axis = spacingContext.axis;
            var gapInPoints = spacingContext.gapInPoints;
            var boundsA = spacingContext.useVisible ? pair.visibleA : pair.geometricA;
            var boundsB = spacingContext.useVisible ? pair.visibleB : pair.geometricB;

            // 進行方向の先頭辺で前後を判定 / Decide leading/trailing by the axis start edge
            var leadObject, trailObject, leadBounds, trailBounds;
            if (axis.start(boundsA) < axis.start(boundsB)) {
                leadObject = pair.a; leadBounds = boundsA;
                trailObject = pair.b; trailBounds = boundsB;
            } else {
                leadObject = pair.b; leadBounds = boundsB;
                trailObject = pair.a; trailBounds = boundsA;
            }

            var gapShift, movedObject, movedBounds, anchorBounds;
            if (spacingContext.anchorEnd) {
                // 後ろ側（右 or 下）を固定：先頭側オブジェクトを動かす / Fix trailing end, move the leading object
                gapShift = (axis.start(trailBounds) - gapInPoints) - axis.end(leadBounds);
                moveByAxis(axis, leadObject, gapShift);
                movedObject = leadObject; movedBounds = leadBounds; anchorBounds = trailBounds;
            } else {
                // 先頭側（左 or 上）を固定：後ろ側オブジェクトを動かす / Fix leading end, move the trailing object
                gapShift = (axis.end(leadBounds) + gapInPoints) - axis.start(trailBounds);
                moveByAxis(axis, trailObject, gapShift);
                movedObject = trailObject; movedBounds = trailBounds; anchorBounds = leadBounds;
            }
            // 整列はギャップ軸と直交方向：ギャップ移動で整列軸の値は変わらないのでキャッシュ境界で可。
            // Alignment is perpendicular to the gap; the gap move doesn't change the align-axis value, so cache is fine.
            if (spacingContext.alignMode !== "none") {
                alignToAnchor(spacingContext.alignAxis, movedObject, movedBounds, anchorBounds, spacingContext.alignMode);
            }
            moveByAxis(spacingContext.alignAxis, movedObject, spacingContext.offsetAlong); // 位置オフセット（移動側のみ）/ position offset (moved object only)
        }

        /**
         * 設定値で各ペアの間隔を調整する。固定側から軸（水平/垂直）と固定端（先頭/後ろ）を決める。
         * alignMode は整列（ギャップ軸に直交する方向、固定オブジェクト基準）。
         * 位置オフセット（offsetAlong）も直交方向で、キーオブジェクトでない側（移動側）だけをずらす（右＝正／下＝正）
         * @param {string} fixedSide - 固定する側
         * @param {number} gapInPoints - 間隔（pt）
         * @param {string} boundsType - "visibleBounds" または "geometricBounds"
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @param {number} offsetAlong - 位置オフセット（pt）
         * @param {string} justifyMode - 行揃えモード
         * @returns {void}
         */
        function applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode) {
            var axis = axisForSide(fixedSide);
            var spacingContext = {
                fixedSide: fixedSide,
                gapInPoints: gapInPoints,
                useVisible: (boundsType === "visibleBounds"),
                axis: axis,
                anchorEnd: isAnchorEnd(fixedSide),
                alignAxis: makeAxis(!axis.vertical), // 整列はギャップ軸に直交 / perpendicular to the gap axis
                alignMode: alignMode,
                offsetAlong: offsetAlong,
                justifyMode: justifyMode
            };

            // 行揃えを先に全テキストへ適用し、境界を取り直す（ポイント文字は行揃えで字幅が変わるため、
            // 間隔・整列を行揃え後の実際の形で計算する）。applyJustification は冪等。
            // Apply justification to all text FIRST, then re-measure bounds, so gap/align use the
            // post-justification shape (point-text width changes with justification). Idempotent.
            for (var i = 0; i < objectPairs.length; i++) {
                var pairItems = getPairItems(objectPairs[i]);
                for (var j = 0; j < pairItems.length; j++) {
                    applyJustification(pairItems[j], resolveJustification(pairItems[j], spacingContext));
                }
            }
            // 行揃えで境界が変わったときだけ取り直す / Re-cache only when justification changed the bounds
            if (boundsCacheStale) {
                cachePairBounds(objectPairs);
                boundsCacheStale = false;
            }

            for (var k = 0; k < objectPairs.length; k++) {
                var pair = objectPairs[k];
                if (pair.single) {
                    applyArtboardMargin(pair, spacingContext);
                } else if (pair.members) {
                    applyGroupSpacing(pair, spacingContext);
                } else {
                    applyPairGap(pair, spacingContext);
                }
            }
        }

        /**
         * 直前のプレビューを巻き戻してから再適用する。
         * 行揃え・整列ともプレビュー時点で最終結果と一致するので、OK では別処理は不要
         * @param {string} fixedSide - 固定する側
         * @param {number} gapInPoints - 間隔（pt）
         * @param {string} boundsType - "visibleBounds" または "geometricBounds"
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @param {number} offsetAlong - 位置オフセット（pt）
         * @param {string} justifyMode - 行揃えモード
         * @returns {void}
         */
        function runPreview(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode) {
            undoPreview();
            applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode);
            app.redraw();
        }

        /**
         * 各ペアの元の境界をキャッシュする。
         * 適用時は常に元位置へ巻き戻してから計算するので、ここで一度取れば使い回せる
         * @param {Object[]} pairs - 作業単位の配列
         * @returns {void}
         */
        function cachePairBounds(pairs) {
            for (var i = 0; i < pairs.length; i++) {
                var pair = pairs[i];
                if (pair.single) {
                    // アートボード：オブジェクト境界と所属アートボードの矩形をキャッシュ
                    // Artboard: cache the object's bounds and its artboard rect
                    pair.geometricS = getClipAwareBounds(pair.single, false);
                    pair.visibleS = getClipAwareBounds(pair.single, true);
                    pair.artboardRect = artboardRectFor(pair.single);
                } else if (pair.members) {
                    // グループ：全メンバーの境界をキャッシュ / Group: cache every member's bounds
                    pair.memberGeo = [];
                    pair.memberVis = [];
                    for (var j = 0; j < pair.members.length; j++) {
                        pair.memberGeo.push(getClipAwareBounds(pair.members[j], false));
                        pair.memberVis.push(getClipAwareBounds(pair.members[j], true));
                    }
                } else {
                    pair.geometricA = getClipAwareBounds(pair.a, false);
                    pair.visibleA = getClipAwareBounds(pair.a, true);
                    pair.geometricB = getClipAwareBounds(pair.b, false);
                    pair.visibleB = getClipAwareBounds(pair.b, true);
                }
            }
        }

        /**
         * モードに応じてペアを組み直す。境界キャッシュは必ず元位置で取るため、先にプレビューを巻き戻してから組む
         * @param {string} mode - "auto" / "group" / "artboard"
         * @returns {void}
         */
        function buildPairs(mode) {
            undoPreview();
            var pairs;
            if (mode === "group") {
                pairs = createGroupPairs(selectedItems);
            } else if (mode === "artboard") {
                pairs = createArtboardUnits(selectedItems);
            } else {
                pairs = createNearestPairs(selectedItems);
            }
            cachePairBounds(pairs);
            boundsCacheStale = false;
            objectPairs = pairs;
        }

        /**
         * 前回終了時の設定をダイアログへ戻す（モード・キー・プレビュー境界・整列・行揃え・間隔・左右/上下オフセット）。
         * モードとキー側は初期間隔を測ったときと同じ値を使う。数値は pt で保存しているので定規の単位の表示へ戻す。
         * 整列とオフセットは向きごとの控えに入れておき、パネルを切り替えたときに反映する
         * @param {Object} dialogControls - buildSettingsDialog() の戻り値
         * @param {Object} orientationValues - 向きごとの控え（h / v）
         * @returns {void}
         */
        function restoreDialogSettings(dialogControls, orientationValues) {
            var alignRadios = dialogControls.alignmentRefs.alignRadios;
            dialogControls.modeRefs.setMode(initialMode);
            dialogControls.fixedSideRefs.setFixedSide(initialFixedSide);
            if (savedSettings.previewBounds) dialogControls.gapRefs.previewBoundsCheckbox.value = true;
            // 整列（キーは none/start/center/end）/ Alignment
            if (alignRadios[savedSettings.alignH]) orientationValues.h.align = savedSettings.alignH;
            if (alignRadios[savedSettings.alignV]) orientationValues.v.align = savedSettings.alignV;
            // テキストの行揃え / Text alignment
            if (savedSettings.justify) dialogControls.justifyRefs.setJustifyMode(savedSettings.justify);
            // 数値（間隔・左右・上下）/ Numeric values (gap, horizontal/vertical offsets)
            var savedGapDisplay = savedPointsToDisplayText(savedSettings.gap);
            if (savedGapDisplay !== null) dialogControls.gapRefs.spacingInput.text = savedGapDisplay;
            var savedOffsetHDisplay = savedPointsToDisplayText(savedSettings.offsetH);
            if (savedOffsetHDisplay !== null) orientationValues.h.offset = savedOffsetHDisplay;
            var savedOffsetVDisplay = savedPointsToDisplayText(savedSettings.offsetV);
            if (savedOffsetVDisplay !== null) orientationValues.v.offset = savedOffsetVDisplay;
        }

        /**
         * ダイアログを生成して表示する。OK では現在の設定をすべて保存する（数値は pt で保存）
         * @returns {boolean} OK なら true
         */
        function showSettingsDialog() {
            var dialogControls = buildSettingsDialog(initialGapPoints);
            var settingsDialog = dialogControls.settingsDialog;
            var modeRefs = dialogControls.modeRefs;
            var fixedSideRefs = dialogControls.fixedSideRefs;
            var gapRefs = dialogControls.gapRefs;
            var alignmentRefs = dialogControls.alignmentRefs;
            var justifyRefs = dialogControls.justifyRefs;
            var alignRadios = alignmentRefs.alignRadios;
            var offsetInput = alignmentRefs.offsetInput;

            var orientationSwitcher = createOrientationSwitcher(fixedSideRefs, alignmentRefs);
            var orientationValues = orientationSwitcher.orientationValues;
            restoreDialogSettings(dialogControls, orientationValues);

            /**
             * 現在の設定でプレビューを更新する（行揃え・整列とも最終結果と一致）
             * @returns {void}
             */
            function refreshPreview() {
                runPreview(fixedSideRefs.getFixedSide(), gapRefs.getSpacingInPoints(), gapRefs.getBoundsType(),
                    getAlignValue(alignRadios), displayTextToPoints(offsetInput.text), justifyRefs.getJustifyMode());
            }

            /**
             * 行揃えの対象・向きが変わりうる操作（モード／キー／整列の切替）用：先に行揃えを元へ戻してから
             * プレビューを更新する。これで「なし」へ戻したときや別の向きへ変えたときに正しく反映される
             * @returns {void}
             */
            function refreshPreviewResetJustify() {
                undoJustifications();
                refreshPreview();
            }

            /**
             * モードを切り替えてペアを組み直し、プレビューを更新する。
             * モードで対象オブジェクトが変わるので行揃えを元へ戻してから組み直す
             * @returns {void}
             */
            function onModeChange() {
                undoJustifications();
                buildPairs(modeRefs.getMode());
                refreshPreview();
            }

            /**
             * 整列の変更時：中央↔それ以外でオフセットの有効/無効が変わるので更新してからプレビュー
             * @returns {void}
             */
            function onAlignChange() {
                orientationSwitcher.updateActivePanels();
                refreshPreviewResetJustify();
            }

            // 設定変更でライブプレビュー / Update preview on change
            for (var i = 0; i < modeRefs.modeRadios.length; i++) {
                modeRefs.modeRadios[i].onClick = onModeChange;
            }
            // 固定側のラジオ：手動で排他にしてからプレビュー更新（軸の切替もここで反映）
            // Fixed-side radios: enforce exclusivity by hand, then refresh (axis switch applies here too)
            var fixedRadios = fixedSideRefs.fixedRadios;
            for (var j = 0; j < fixedRadios.length; j++) {
                fixedRadios[j].onClick = (function (fixedRadio) {
                    return function () {
                        fixedSideRefs.selectFixedRadio(fixedRadio);
                        orientationSwitcher.updateActivePanels();
                        refreshPreviewResetJustify();
                    };
                })(fixedRadios[j]);
            }
            // 間隔・位置オフセット：入力でプレビュー更新（∧∨と↑↓キーも onChanging 経由で更新）
            // Gap and offset: refresh on input (steppers and arrow keys go through onChanging too)
            gapRefs.spacingInput.onChanging = refreshPreview;
            offsetInput.onChanging = refreshPreview;
            gapRefs.previewBoundsCheckbox.onClick = refreshPreview;
            // 整列ラジオ：オフセットの有効/無効を更新し、行揃えを戻してから更新 / Alignment radios
            for (var alignKey in alignRadios) { alignRadios[alignKey].onClick = onAlignChange; }
            // 整列のキーボードショートカット（水平 L/C/R・垂直 T/M/B）。今の向きで読み替える / Alignment keyboard shortcuts
            addAlignmentKeyHandler(settingsDialog, alignRadios, orientationSwitcher.isVerticalGap, [gapRefs.spacingInput, offsetInput]);
            // テキストの行揃えボタン：押した値をアクティブにし、行揃えを戻してから再適用
            // Justification buttons: activate the clicked value, revert justification, then refresh
            var justifyButtons = justifyRefs.buttons;
            for (var k = 0; k < justifyButtons.length; k++) {
                justifyButtons[k].onClick = function () {
                    justifyRefs.setJustifyMode(this.justifyId);
                    refreshPreviewResetJustify();
                };
            }

            // ダイアログ表示時に既定モードでペアを組んで初回プレビュー（同期側 undo を避けて onShow から起動）
            // Build pairs for the default mode, then run the first preview (from onShow to avoid sync undo)
            settingsDialog.onShow = function () {
                orientationSwitcher.updateActivePanels(); // 既定のキー側に合わせて水平/垂直パネルとオフセットの有効/無効を初期化 / Init enabled state
                buildPairs(modeRefs.getMode());
                refreshPreview();
            };

            prepareDialogWindow(settingsDialog, SCRIPT_NAME);
            var isAccepted = (settingsDialog.show() === 1);
            if (isAccepted) {
                // プレビュー状態がそのまま最終結果（行揃え・整列とも反映済み）なので、確定処理は保存のみ。
                // The preview already is the final result (justification + alignment), so OK just saves.
                orientationSwitcher.stashOrientationValues(); // 出ている向きの値を控えへ入れてから保存 / stash the shown values first
                saveSettings({
                    mode: modeRefs.getMode(),
                    fixedSide: fixedSideRefs.getFixedSide(),
                    previewBounds: !!gapRefs.previewBoundsCheckbox.value,
                    alignH: orientationValues.h.align,
                    alignV: orientationValues.v.align,
                    justify: justifyRefs.getJustifyMode(),
                    gap: String(gapRefs.getSpacingInPoints()),
                    offsetH: String(displayTextToPoints(orientationValues.h.offset)),
                    offsetV: String(displayTextToPoints(orientationValues.v.offset))
                });
            }
            return isAccepted;
        }

        // OK はプレビューをそのまま確定。キャンセルは位置と行揃えを巻き戻す
        // OK keeps the applied preview; cancel reverts it (position + justification)
        if (!showSettingsDialog()) {
            undoPreview();
            undoJustifications();
            app.redraw();
        }
    })();

})();
