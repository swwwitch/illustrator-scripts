#target illustrator
#targetengine "AutoTouchTypeEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの各文字に対して、ベースライン・比率・回転・カーニング・トラッキングをランダムに付与します。
seed付きの乱数でプレビューの見た目を安定させ、文字回転を適用したときはトラッキングを自動補正します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoTouchType.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne6545c4717af

### Overview

Randomizes the baseline shift, scale, rotation, kerning and tracking of each character in the selected text.
A seeded RNG keeps the preview stable, and the tracking is corrected automatically when character rotation is applied.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoTouchType.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoTouchType";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoTouchType.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoTouchType.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne6545c4717af"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「犯行声明文」風の背景を置くレイヤー名とグループ名 / Layer and group that hold the ransom-note backgrounds */
    var RANSOM_BG_LAYER_NAME = "__AutoTouchType_BG__";
    var RANSOM_BG_GROUP_NAME = "__BGRects__";

    /* 背景長方形の余白・ゆがみ量（pt）とグレー濃度（%）/ Padding, jitter (pt) and gray range (%) of the background rectangles */
    var RANSOM_RECT_PADDING_PT = 1;
    var RANSOM_RECT_JITTER_PT = 1;
    var RANSOM_GRAY_MIN = 10;
    var RANSOM_GRAY_MAX = 50;

    /* 「犯行声明文」風のトラッキング初期値（1/1000em）/ Default ransom-note tracking (1/1000 em) */
    var RANSOM_TRACKING_DEFAULT = 200;

    /* 文字タッチの初期値 / Default touch amounts */
    var DEFAULT_SCALE_PERCENT = 10;
    var DEFAULT_KERNING_EM = 50;
    var DEFAULT_ROTATION_DEG = 5;

    /* スライダーの範囲 / Slider ranges */
    var SCALE_SLIDER_MAX = 200;
    var KERNING_SLIDER_MIN = -200;
    var KERNING_SLIDER_MAX = 200;
    var ROTATION_SLIDER_MAX = 30;
    var BASELINE_SLIDER_MAX_FALLBACK = 50;
    var ZOOM_MIN_PERCENT = 10;
    var ZOOM_MAX_PERCENT = 1600;

    /* 画面更新の間引き（ミリ秒）/ Redraw throttling in milliseconds */
    var PREVIEW_DELAY_MS = 120;
    var ZOOM_THROTTLE_MS = 150;

    /* 文字回転に対するトラッキング補正の係数 / Factors of the rotation tracking compensation */
    var ROTATION_GAP_FACTOR_AFTER_POSITIVE = 680;
    var ROTATION_GAP_FACTOR_AFTER_NEGATIVE = 1050;
    var ROTATION_GAP_FACTOR_BEFORE = 180;
    var ROTATION_GAP_UPPERCASE_BEFORE = 0.6;
    var ROTATION_GAP_UPPERCASE_AFTER = 1.05;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];
    var SLIDER_WIDTH = 180;
    var ZOOM_SLIDER_WIDTH = 270;
    var TOGGLE_WIDTH = 15;
    var VALUE_FIELD_CHARS = 4;
    var RANSOM_FIELD_CHARS = 5;
    var SMALL_BUTTON_SIZE = [74, 22];
    var TOUCH_BUTTON_ROW_TOP_MARGIN = 10;

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オート文字タッチツール", en: "Auto Touch Type Tool" }
        },
        panel: {
            touch: {
                ja: "位置・スケール・回転などの「ゆらぎ」",
                en: "Position, Scale and Rotation Variation"
            },
            font: { ja: "フォント", en: "Font" },
            ransom: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            randomFont: { ja: "1文字ごとに変更", en: "Change per character" },
            japaneseOnly: { ja: "和文フォントに限定", en: "Japanese only" },
            ransomEnabled: { ja: "「犯行声明文」風", en: "Ransom-note style" },
            ransomTracking: { ja: "トラッキング調整", en: "Adjust tracking" },
            rotationTracking: {
                ja: "文字回転によるトラッキング補正",
                en: "Tracking compensation for rotation"
            },
            lightMode: { ja: "軽量モード", en: "Light mode" }
        },
        fieldLabel: {
            baseline: { ja: "ベースライン", en: "Baseline" },
            scale: { ja: "水平／垂直比率", en: "Scale" },
            rotation: { ja: "文字回転", en: "Rotation" },
            kerning: { ja: "カーニング", en: "Kerning" },
            zoom: { ja: "ズーム", en: "Zoom" }
        },
        button: {
            randomize: { ja: "ランダム", en: "Random" },
            reset: { ja: "リセット", en: "Reset" },
            allOn: { ja: "すべてON", en: "All ON" },
            allOff: { ja: "すべてOFF", en: "All OFF" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectText: { ja: "テキストを選択してください。", en: "Please select text." },
            selectTextRange: {
                ja: "テキスト、または文字を選択してください。",
                en: "Please select a text object or characters."
            },
            enterNumber: { ja: "数値を入力してください。", en: "Please enter a number." },
            noJapaneseFonts: {
                ja: "対象の和文フォント（Pr6 / Pr6N）が見つかりません。",
                en: "No target JP fonts (Pr6 / Pr6N) were found."
            }
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
            randomFont: { ja: "1文字ずつフォントを入れ替えます。", en: "Swaps the font of each character." },
            japaneseOnly: {
                ja: "入れ替え先を和文フォント（Pr6／Pr6N）だけに絞ります。",
                en: "Limits the replacement fonts to Japanese fonts (Pr6 / Pr6N)."
            },
            ransomEnabled: {
                ja: "1文字ずつ大きさと書体をばらつかせ、切り貼りしたような見た目にします。",
                en: "Varies the size and typeface of each character, like letters cut from a magazine."
            },
            ransomTracking: {
                ja: "ばらついた文字幅に合わせて字間を詰めます。",
                en: "Tightens the spacing to match the varied character widths."
            },
            ransomTrackingValue: { ja: "詰める量です。単位は1/1000em。", en: "How much to tighten, in 1/1000 em." },
            baselineEnabled: { ja: "ベースラインシフトをかけるかどうかです。", en: "Whether to apply a baseline shift." },
            baseline: { ja: "1文字ずつ上下にずらす最大量です。", en: "Maximum amount each character is shifted up or down." },
            scaleEnabled: { ja: "文字を長体・平体にするかどうかです。", en: "Whether to condense or extend the characters." },
            scale: {
                ja: "1文字ずつ変える水平／垂直比率の最大量です。",
                en: "Maximum change applied to each character's horizontal and vertical scale."
            },
            kerningEnabled: { ja: "字間をばらつかせるかどうかです。", en: "Whether to vary the spacing between characters." },
            kerning: { ja: "1文字ずつ変える字間の最大量です。単位は1/1000em。", en: "Maximum spacing change per character, in 1/1000 em." },
            rotationEnabled: { ja: "文字を回転させるかどうかです。", en: "Whether to rotate the characters." },
            rotation: { ja: "1文字ずつ回転させる最大角度です。", en: "Maximum rotation applied to each character." },
            rotationTracking: {
                ja: "回転で広がった見た目の幅を、字間で打ち消します。",
                en: "Offsets the apparent width added by the rotation with the letter spacing."
            },
            allOn: { ja: "文字タッチの4項目をすべてオンにします。", en: "Turns on all four touch settings." },
            allOff: { ja: "文字タッチの4項目をすべてオフにします。", en: "Turns off all four touch settings." },
            zoom: {
                ja: "作業中の画面表示倍率を変えます。結果には影響しません。",
                en: "Changes the view zoom while you work. It does not affect the result."
            },
            lightMode: {
                ja: "ドラッグ中は画面表示倍率を変えず、離したときにまとめて反映します。",
                en: "Applies the zoom only when the drag ends, instead of while dragging."
            },
            randomize: { ja: "同じ設定のまま、乱数だけ振り直します。", en: "Re-rolls the randomness while keeping the same settings." },
            reset: {
                ja: "選択したテキストの文字属性を初期状態に戻します。",
                en: "Restores the character attributes of the selected text."
            }
        }
    };

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

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
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

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) { alert(getLabel("alert.selectText")); return; }

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

    // -----------------------------------------
    // 画面表示倍率 / View zoom
    // -----------------------------------------

    /**
     * バウンディングボックスの中心点を返す
     * @param {number[]} bounds - [left, top, right, bottom]
     * @returns {number[]} [x, y]
     */
    function getBoundsCenter(bounds) {
        return [bounds[0] + (bounds[2] - bounds[0]) / 2, bounds[1] + (bounds[3] - bounds[1]) / 2];
    }

    /**
     * 表示倍率をスライダーの範囲に収める
     * @param {number} zoomPercent - 表示倍率（%）
     * @returns {number} 範囲内に丸めた表示倍率（%）
     */
    function clampZoomPercent(zoomPercent) {
        var percent = Math.round(zoomPercent);
        if (percent < ZOOM_MIN_PERCENT) percent = ZOOM_MIN_PERCENT;
        if (percent > ZOOM_MAX_PERCENT) percent = ZOOM_MAX_PERCENT;
        return percent;
    }

    var docView = doc.views[0];
    var originalZoomFactor = docView.zoom;
    var originalCenterPoint = docView.centerPoint;
    var zoomCenterPoint = originalCenterPoint;
    /* 選択が TextRange だと境界を測れない（null で例外になる） / A TextRange selection has no bounds (null throws) */
    try {
        zoomCenterPoint = getBoundsCenter(getClipAwareUnionBounds(doc.selection, true));
    } catch (e) { }

    /**
     * 指定した表示倍率でドキュメントを表示し直す
     * @param {number} zoomPercent - 表示倍率（%）
     * @returns {void}
     */
    function applyZoomPercent(zoomPercent) {
        docView.zoom = clampZoomPercent(zoomPercent) / 100.0;
        if (zoomCenterPoint) docView.centerPoint = zoomCenterPoint;
        app.redraw();
    }

    /**
     * 開始時の表示倍率と表示位置に戻す
     * @returns {void}
     */
    function restoreOriginalView() {
        docView.zoom = originalZoomFactor;
        docView.centerPoint = originalCenterPoint;
        app.redraw();
    }

    // -----------------------------------------
    // 「犯行声明文」風の背景 / Ransom-note backgrounds
    // -----------------------------------------

    /**
     * 選択範囲に含まれるテキストフレームを重複なく集める
     * @param {Object[]} selectionItems - 選択オブジェクトの配列
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function getSelectionTextFrames(selectionItems) {
        /* 文字編集中の選択はフレームにする（グループの中は見ない） / A text-editing selection resolves to its frame (groups are not entered) */
        return collectSelectionTextFrames(selectionItems, { enterGroups: false });
    }

    /**
     * 背景用レイヤーを探し、無ければ作る
     * @returns {Layer} 背景用レイヤー
     */
    function ensureRansomBgLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === RANSOM_BG_LAYER_NAME) return doc.layers[i];
        }
        var bgLayer = doc.layers.add();
        bgLayer.name = RANSOM_BG_LAYER_NAME;
        return bgLayer;
    }

    /**
     * 背景用グループを削除する
     * @param {Layer} bgLayer - 背景用レイヤー
     * @returns {void}
     */
    function removeRansomBgGroup(bgLayer) {
        if (!bgLayer) return;
        for (var i = bgLayer.groupItems.length - 1; i >= 0; i--) {
            if (bgLayer.groupItems[i].name === RANSOM_BG_GROUP_NAME) bgLayer.groupItems[i].remove();
        }
    }

    /**
     * min以上max以下の整数を返す
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @returns {number} 乱数
     */
    function randomIntBetween(minValue, maxValue) {
        return Math.floor(Math.random() * (maxValue - minValue + 1)) + minValue;
    }

    /**
     * 濃度をばらつかせたグレーを作る
     * @returns {GrayColor} 背景用のグレー
     */
    function createRandomGrayFill() {
        var grayColor = new GrayColor();
        grayColor.gray = randomIntBetween(RANSOM_GRAY_MIN, RANSOM_GRAY_MAX);
        return grayColor;
    }

    /**
     * 長方形の4隅を外側へランダムにずらす
     * @param {PathItem} rectPath - 対象の長方形
     * @param {number} maxOffsetPt - ずらす最大量（pt）
     * @returns {void}
     */
    function expandRectCornersRandomly(rectPath, maxOffsetPt) {
        if (!rectPath.pathPoints || rectPath.pathPoints.length < 4) return;
        var bounds = rectPath.geometricBounds;
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        for (var i = 0; i < rectPath.pathPoints.length; i++) {
            var pathPoint = rectPath.pathPoints[i];
            var anchorX = pathPoint.anchor[0];
            var anchorY = pathPoint.anchor[1];
            var dx = anchorX - centerX;
            var dy = anchorY - centerY;
            var distance = Math.sqrt(dx * dx + dy * dy);
            if (!distance) continue;
            var offsetDistance = Math.random() * maxOffsetPt;
            var offsetX = (dx / distance) * offsetDistance;
            var offsetY = (dy / distance) * offsetDistance;
            pathPoint.anchor = [anchorX + offsetX, anchorY + offsetY];
            pathPoint.leftDirection = [pathPoint.leftDirection[0] + offsetX, pathPoint.leftDirection[1] + offsetY];
            pathPoint.rightDirection = [pathPoint.rightDirection[0] + offsetX, pathPoint.rightDirection[1] + offsetY];
        }
    }

    /**
     * 1文字分のバウンディングボックスから背景の長方形を作る
     * @param {GroupItem} bgGroup - 背景用グループ
     * @param {number[]} charBounds - [left, top, right, bottom]
     * @returns {PathItem|null} 作成した長方形（作れないときは null）
     */
    function addRansomRect(bgGroup, charBounds) {
        var rectLeft = charBounds[0] - RANSOM_RECT_PADDING_PT;
        var rectTop = charBounds[1] + RANSOM_RECT_PADDING_PT;
        var rectWidth = (charBounds[2] + RANSOM_RECT_PADDING_PT) - rectLeft;
        var rectHeight = rectTop - (charBounds[3] - RANSOM_RECT_PADDING_PT);
        if (rectWidth <= 0 || rectHeight <= 0) return null;
        var bgRect = bgGroup.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
        bgRect.stroked = false;
        bgRect.filled = true;
        bgRect.fillColor = createRandomGrayFill();
        expandRectCornersRandomly(bgRect, RANSOM_RECT_JITTER_PT);
        return bgRect;
    }

    /**
     * アウトライン化した結果から1文字ぶんのまとまりを集める
     * @param {Object} outlinedItem - createOutline() の戻り値
     * @param {Object[]} charItems - 集めた結果を入れる配列
     * @returns {void}
     */
    function collectOutlinedCharItems(outlinedItem, charItems) {
        if (!outlinedItem) return;

        /* 末端のパスまで降りて拾う / Walk down to the leaf paths */
        function pushLeafItems(parentItem) {
            if (!parentItem || !parentItem.pageItems) return;
            for (var k = 0; k < parentItem.pageItems.length; k++) {
                var childItem = parentItem.pageItems[k];
                if (!childItem) continue;
                if (childItem.typename === "GroupItem") pushLeafItems(childItem);
                else if (childItem.typename === "PathItem" || childItem.typename === "CompoundPathItem") charItems.push(childItem);
            }
        }

        if (outlinedItem.typename !== "GroupItem") {
            charItems.push(outlinedItem);
            return;
        }

        /* 行グループ → 文字グループの順に入れ子になっている / Line groups hold the character groups */
        var pushedCharGroup = false;
        for (var i = 0; i < outlinedItem.groupItems.length; i++) {
            var lineGroup = outlinedItem.groupItems[i];
            if (!lineGroup || lineGroup.typename !== "GroupItem") continue;
            if (lineGroup.groupItems.length > 0) {
                for (var j = 0; j < lineGroup.groupItems.length; j++) {
                    charItems.push(lineGroup.groupItems[j]);
                    pushedCharGroup = true;
                }
            } else {
                charItems.push(lineGroup);
                pushedCharGroup = true;
            }
        }
        if (pushedCharGroup) return;

        pushLeafItems(outlinedItem);
        if (charItems.length === 0) charItems.push(outlinedItem);
    }

    /**
     * テキストを複製してアウトライン化し、1文字ごとの背景長方形を作る
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function createRansomBgRects(textFrames) {
        if (!textFrames || textFrames.length === 0) return;
        var bgLayer = ensureRansomBgLayer();
        removeRansomBgGroup(bgLayer);
        var bgGroup = bgLayer.groupItems.add();
        bgGroup.name = RANSOM_BG_GROUP_NAME;

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!textFrame || textFrame.typename !== "TextFrame") continue;

            /* ロックや非表示のレイヤー上では複製・アウトライン化に失敗することがある
               Duplicating or outlining can fail on locked or hidden layers */
            var duplicatedFrame = null;
            var outlinedGroup = null;
            try {
                duplicatedFrame = textFrame.duplicate();
                duplicatedFrame.move(textFrame, ElementPlacement.PLACEAFTER);
            } catch (e) { }
            if (!duplicatedFrame) continue;
            try {
                outlinedGroup = duplicatedFrame.createOutline();
            } catch (e) { }
            if (!outlinedGroup) {
                /* createOutline() は成功すると複製を消費する。失敗して残った複製だけ片付ける
                   createOutline() consumes the duplicate on success; remove the one left behind by a failure */
                try { duplicatedFrame.remove(); } catch (e) { }
                continue;
            }

            var charItems = [];
            collectOutlinedCharItems(outlinedGroup, charItems);
            for (var j = 0; j < charItems.length; j++) {
                /* スペースなど中身の無い文字グループは geometricBounds を取れない
                   An empty character group, such as a space, has no geometricBounds */
                try {
                    addRansomRect(bgGroup, charItems[j].geometricBounds);
                } catch (e) { }
            }
            outlinedGroup.remove();
        }

        bgGroup.zOrder(ZOrderMethod.SENDTOBACK);
        bgLayer.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 背景の長方形を消す（無ければ何もしない）
     * @returns {void}
     */
    function clearRansomBgRects() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === RANSOM_BG_LAYER_NAME) {
                removeRansomBgGroup(doc.layers[i]);
                return;
            }
        }
    }

    // -----------------------------------------
    // フォント / Fonts
    // -----------------------------------------

    var allFonts = app.textFonts;

    /**
     * ひらがな・カタカナ・漢字を含むか判定する
     * @param {string} sourceText - 判定する文字列
     * @returns {boolean} 含むとき true
     */
    function hasJapaneseCharacters(sourceText) {
        if (!sourceText) return false;
        return /[぀-ゟ゠-ヿ一-鿿]/.test(String(sourceText));
    }

    /* 名前に含まれていたら和文として扱わない語 / Markers that rule a font out */
    var NON_JAPANESE_FONT_MARKERS = [
        "Apple LiGothic",
        "RyoGothicStd",
        "-KO", "-KL", "LogoArl",
        "Kana"
    ];

    /* 和文フォントとみなす語（和文ファミリー名とベンダー名）/ Keywords that mark a font as Japanese */
    var JAPANESE_FONT_KEYWORDS = [
        "ゴシック", "明朝", "丸ゴ", "教科書", "楷書",
        "Mincho", "Maru",
        "Hiragino", "ヒラギノ",
        "Yu Gothic", "Yu Mincho", "游ゴシック", "游明朝",
        "Meiryo", "メイリオ",
        "MS Gothic", "MS Mincho", "MS ゴシック", "MS 明朝",
        "Kozuka", "小塚",
        "Morisawa", "モリサワ",
        "Ryumin", "Shin Go", "新ゴ",
        "Heisei", "平成",
        "Klee", "クレー",
        "Tsukushi", "筑紫",
        "A-OTF", "AP-OTF ", "-OTF",
        "FOT", "Pr6N", "Pr6",
        "Noto Sans JP", "Noto Serif JP",
        "Source Han", "源ノ角", "源ノ明",
        "Min2"
    ];

    /**
     * 和文フォントかどうかを名前から判定する
     * @param {TextFont} candidateFont - 判定するフォント
     * @returns {boolean} 和文とみなせるとき true
     */
    function isJapaneseFont(candidateFont) {
        if (!candidateFont) return false;
        var fontNames = [];
        /* 環境にないフォントは名前を取れないことがある / A font missing from the system may not expose its names */
        try {
            fontNames = [String(candidateFont.name || ""), String(candidateFont.family || ""), String(candidateFont.fullName || ""), String(candidateFont.postScriptName || "")];
        } catch (e) {
            return false;
        }

        var i, j;
        for (i = 0; i < NON_JAPANESE_FONT_MARKERS.length; i++) {
            for (j = 0; j < fontNames.length; j++) {
                if (fontNames[j].indexOf(NON_JAPANESE_FONT_MARKERS[i]) !== -1) return false;
            }
        }

        /* 名前そのものが和文表記ならそれで判定できる / A Japanese name settles it */
        for (j = 0; j < 3; j++) {
            if (hasJapaneseCharacters(fontNames[j])) return true;
        }

        for (i = 0; i < JAPANESE_FONT_KEYWORDS.length; i++) {
            for (j = 0; j < fontNames.length; j++) {
                if (fontNames[j].indexOf(JAPANESE_FONT_KEYWORDS[i]) !== -1) return true;
            }
        }
        return false;
    }

    /* 和文フォントの一覧は作るのが重いので使い回す / Building the list is heavy, so cache it */
    var japaneseFontsCache = null;

    /**
     * 環境にある和文フォントの一覧を返す
     * @returns {TextFont[]} 和文フォントの配列
     */
    function getJapaneseFonts() {
        if (japaneseFontsCache) return japaneseFontsCache;
        japaneseFontsCache = [];
        for (var i = 0; i < allFonts.length; i++) {
            if (isJapaneseFont(allFonts[i])) japaneseFontsCache.push(allFonts[i]);
        }
        return japaneseFontsCache;
    }

    // -----------------------------------------
    // 選択テキスト / Selected text
    // -----------------------------------------

    /**
     * 選択オブジェクトから TextRange を集める
     * @param {Object[]} selectionItems - 選択オブジェクトの配列
     * @returns {TextRange[]} TextRange の配列
     */
    function collectTextRanges(selectionItems) {
        var collectedRanges = [];
        for (var i = 0; i < selectionItems.length; i++) {
            var selectedItem = selectionItems[i];
            if (!selectedItem) continue;
            if (selectedItem.typename === "TextRange") collectedRanges.push(selectedItem);
            else if (selectedItem.typename === "TextFrame") collectedRanges.push(selectedItem.textRange);
        }
        return collectedRanges;
    }

    /**
     * 英数字以外（改行を除く）を含むか判定する
     * @param {TextRange[]} targetRanges - 判定する TextRange
     * @returns {boolean} 含むとき true
     */
    function containsNonAlphanumeric(targetRanges) {
        for (var i = 0; i < targetRanges.length; i++) {
            var rangeText = targetRanges[i].contents;
            for (var j = 0; j < rangeText.length; j++) {
                var oneChar = rangeText.charAt(j);
                if (oneChar === "\r" || oneChar === "\n") continue;
                if (!/[A-Za-z0-9]/.test(oneChar)) return true;
            }
        }
        return false;
    }

    /**
     * 手動カーニング値を読む（読めない環境では0）
     * @param {Object} character - 対象の文字
     * @returns {number} カーニング値（1/1000em）
     */
    function getCharacterKerning(character) {
        /* 自動カーニング中の文字では Error 9551 になる / Reading kerning throws 9551 while auto-kerning is on */
        try {
            var kerningValue = character.kerning;
            return (typeof kerningValue === "number") ? kerningValue : 0;
        } catch (e) {
            return 0;
        }
    }

    var textRanges = collectTextRanges(doc.selection);
    var selectedTextFrames = getSelectionTextFrames(doc.selection);
    if (textRanges.length === 0) { alert(getLabel("alert.selectTextRange")); return; }

    // -----------------------------------------
    // 文字属性のスナップショット / Character snapshots
    // -----------------------------------------

    /* 元の文字属性の控え / The attributes each character started with */
    var charSnapshots = [];

    /**
     * 選択中の全文字の属性を控える
     * @returns {void}
     */
    function takeCharSnapshots() {
        charSnapshots = [];
        for (var i = 0; i < textRanges.length; i++) {
            var textRange = textRanges[i];
            for (var j = 0; j < textRange.length; j++) {
                var character = textRange.characters[j];
                var charAttributes = character.characterAttributes;
                charSnapshots.push({
                    character: character,
                    baselineShift: charAttributes.baselineShift,
                    horizontalScale: charAttributes.horizontalScale,
                    verticalScale: charAttributes.verticalScale,
                    rotation: charAttributes.rotation,
                    kerning: getCharacterKerning(character),
                    tracking: (typeof charAttributes.tracking === "number") ? charAttributes.tracking : 0,
                    textFont: charAttributes.textFont
                });
            }
        }
    }

    /**
     * 控えた文字属性に戻す
     * @returns {void}
     */
    function restoreCharSnapshots() {
        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var charAttributes = snapshot.character.characterAttributes;
            charAttributes.baselineShift = snapshot.baselineShift;
            charAttributes.horizontalScale = snapshot.horizontalScale;
            charAttributes.verticalScale = snapshot.verticalScale;
            charAttributes.rotation = snapshot.rotation;
            /* カーニング・トラッキング・フォントは書き込めない環境がある / These three are not writable everywhere */
            try {
                charAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                snapshot.character.kerning = snapshot.kerning;
            } catch (e) { }
            try {
                charAttributes.tracking = snapshot.tracking;
            } catch (e) { }
            try {
                if (snapshot.textFont) charAttributes.textFont = snapshot.textFont;
            } catch (e) { }
        }
    }

    // -----------------------------------------
    // 文字回転によるトラッキング補正 / Tracking compensation for rotation
    // -----------------------------------------

    /**
     * 1文字の回転量から、前後に足すトラッキング量を求める
     * @param {number} rotationDeg - 回転角（度）
     * @param {string} charContent - その文字の内容
     * @returns {{previousGap: number, currentGap: number}} 前の文字と自身に足す量（1/1000em）
     */
    function calcRotationTrackingPair(rotationDeg, charContent) {
        var sine = Math.sin(rotationDeg * (Math.PI / 180));
        var factorAfter = (rotationDeg > 0) ? ROTATION_GAP_FACTOR_AFTER_POSITIVE : ROTATION_GAP_FACTOR_AFTER_NEGATIVE;
        var currentGap = Math.abs(sine) * factorAfter * -1;
        var previousGap = sine * ROTATION_GAP_FACTOR_BEFORE * -1;

        /* 大文字は字形が大きいぶん前後の効き方が変わる / Uppercase letters need a different balance */
        if (charContent && charContent === charContent.toUpperCase() && charContent !== charContent.toLowerCase()) {
            previousGap = previousGap * ROTATION_GAP_UPPERCASE_BEFORE;
            currentGap = currentGap * ROTATION_GAP_UPPERCASE_AFTER;
        }
        return { previousGap: previousGap, currentGap: currentGap };
    }

    /**
     * 文字の回転角を読む
     * @param {Object} character - 対象の文字
     * @returns {number} 回転角（度）
     */
    function getCharacterRotation(character) {
        /* バージョンによって rotation を持たないことがある / Some builds do not expose rotation */
        try {
            if (typeof character.rotation === "number") return character.rotation;
            if (character.characterAttributes && typeof character.characterAttributes.rotation === "number") {
                return character.characterAttributes.rotation;
            }
        } catch (e) { }
        return 0;
    }

    /**
     * 1つの TextRange にトラッキング補正をかける
     * @param {TextRange} textRange - 対象の TextRange
     * @returns {void}
     */
    function applyRotationTrackingToRange(textRange) {
        var characters = textRange.characters;
        var charCount = characters.length;
        if (charCount <= 1) return;

        var trackingDeltas = [];
        var i;
        for (i = 0; i < charCount; i++) trackingDeltas[i] = 0;

        /* 1周目：回転している文字が前後に必要とする量を足し合わせる / Pass 1: accumulate the required gaps */
        for (i = 0; i < charCount; i++) {
            var rotationDeg = getCharacterRotation(characters[i]);
            if (Math.abs(rotationDeg) <= 1.0) continue;
            var gapPair = calcRotationTrackingPair(rotationDeg, characters[i].contents);
            trackingDeltas[i] += gapPair.currentGap;
            if (i > 0) trackingDeltas[i - 1] += gapPair.previousGap;
        }

        /* 2周目：まとめて書き込む / Pass 2: write the values */
        for (i = 0; i < charCount; i++) {
            if (Math.abs(trackingDeltas[i]) <= 0.5) continue;
            var charContent = characters[i].contents;
            if (charContent === "\r" || charContent === "\n") continue;
            /* トラッキングを書き込めない環境がある / Tracking is not writable everywhere */
            try {
                characters[i].characterAttributes.tracking = trackingDeltas[i];
            } catch (e) { }
        }
    }

    /**
     * 選択中のすべての TextRange にトラッキング補正をかける
     * @param {TextRange[]} targetRanges - 対象の TextRange
     * @returns {void}
     */
    function applyRotationTrackingToRanges(targetRanges) {
        for (var i = 0; i < targetRanges.length; i++) {
            applyRotationTrackingToRange(targetRanges[i]);
        }
    }

    // -----------------------------------------
    // ランダム適用 / Randomization
    // -----------------------------------------

    /**
     * seedから同じ並びを再現できる乱数生成器を作る
     * @param {number} randomSeed - 乱数の種
     * @returns {Function} 0以上1未満の乱数を返す関数
     */
    function createSeededRandom(randomSeed) {
        var rngState = randomSeed >>> 0;
        return function () {
            rngState = (1664525 * rngState + 1013904223) >>> 0;
            return rngState / 4294967296;
        };
    }

    /**
     * -1〜1の乱数を返す
     * @param {Function} nextRandom - 乱数生成器
     * @returns {number} -1以上1未満の値
     */
    function randomSigned(nextRandom) {
        return nextRandom() * 2.0 - 1.0;
    }

    /**
     * 入れ替え先のフォント一覧を決める
     * @param {Object} touchOptions - 適用オプション
     * @returns {TextFont[]|null} フォントの配列（入れ替えないときは null）
     */
    function resolveFontPool(touchOptions) {
        if (!touchOptions.randomFont) return null;
        var fontPool = touchOptions.japaneseOnly ? touchOptions.japaneseFonts : touchOptions.allFonts;
        return (fontPool && fontPool.length > 0) ? fontPool : null;
    }

    /**
     * 選択中の各文字にランダムな文字タッチを適用する
     * @param {{baselinePt: number, scalePercent: number, rotationDeg: number, kerningEm: number}} touchAmounts - 各項目の最大量
     * @param {number} randomSeed - 乱数の種
     * @param {Object} touchOptions - フォント入れ替えや「犯行声明文」風などのオプション
     * @returns {void}
     */
    function applyRandomTouch(touchAmounts, randomSeed, touchOptions) {
        var nextRandom = createSeededRandom(randomSeed);
        var fontPool = resolveFontPool(touchOptions);

        /* 「犯行声明文」風のトラッキングは固定値で上書きする / Ransom-note tracking overrides the random spacing */
        var ransomTrackEnabled = (touchOptions.ransomTrack !== false);
        var useFixedTracking = (touchOptions.ransom && ransomTrackEnabled) || touchOptions.previewRansomTracking;
        var fixedTracking = (typeof touchOptions.ransomTrackValue === "number") ? touchOptions.ransomTrackValue : RANSOM_TRACKING_DEFAULT;
        var addFixedToExisting = (touchOptions.rotationTracking !== false);

        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var charAttributes = snapshot.character.characterAttributes;
            var charContent = snapshot.character.contents;

            if (fontPool && !(charContent === "\r" || charContent === "\n" || charContent === " ")) {
                /* 環境にないフォントは適用できない / A font missing from the system cannot be applied */
                try {
                    charAttributes.textFont = fontPool[Math.floor(nextRandom() * fontPool.length)];
                } catch (e) { }
            }

            charAttributes.baselineShift = randomSigned(nextRandom) * touchAmounts.baselinePt;
            charAttributes.rotation = snapshot.rotation + (randomSigned(nextRandom) * touchAmounts.rotationDeg);

            var scaleFactor = 1.0 + (randomSigned(nextRandom) * (touchAmounts.scalePercent / 100.0));
            charAttributes.horizontalScale = snapshot.horizontalScale * scaleFactor;
            charAttributes.verticalScale = snapshot.verticalScale * scaleFactor;

            /* カーニングとトラッキングは書き込めない環境がある / Kerning and tracking are not writable everywhere */
            try {
                charAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                snapshot.character.kerning = snapshot.kerning + (randomSigned(nextRandom) * touchAmounts.kerningEm);
            } catch (e) { }
            try {
                if (!useFixedTracking) charAttributes.tracking = snapshot.tracking;
                else if (addFixedToExisting) charAttributes.tracking = snapshot.tracking + fixedTracking;
                else charAttributes.tracking = fixedTracking;
            } catch (e) { }
        }

        /* 回転で広がった見た目の幅を字間で打ち消す / Offset the width the rotation added */
        if (Math.abs(touchAmounts.rotationDeg) > 0.0001 && touchOptions.rotationTracking !== false && !useFixedTracking) {
            applyRotationTrackingToRanges(textRanges);
        }

        if (!touchOptions.skipRedraw) app.redraw();
    }

    /**
     * 文字列を数値に変換する
     * @param {string} inputText - 入力文字列
     * @returns {number|null} 数値（数値にならないときは null）
     */
    function parseNumber(inputText) {
        var parsedValue = parseFloat(inputText);
        return isNaN(parsedValue) ? null : parsedValue;
    }

    takeCharSnapshots();
    var randomSeed = (new Date()).getTime() & 0xffffffff;

    /* 前回の実行が残した背景を消してから始める / Clear any background left by a previous run */
    clearRansomBgRects();

    // -----------------------------------------
    // プレビュー管理 / Preview state
    // -----------------------------------------

    /* プレビューは取り消し（アンドゥ）ではなく上書きで反映する。キャンセル時だけ控えに戻す
       The preview overwrites the attributes instead of using undo; Cancel restores the snapshots */
    var previewApplied = false;
    /* OKで閉じたときは onClose の後始末をしない / A close via OK must not roll the result back */
    var closedByOK = false;
    /* リセット直後にOKされたら、もう一度ランダム化しない / OK right after Reset must not re-randomize */
    var didReset = false;

    /**
     * プレビューを適用し、キャンセル時に戻せるよう記録する
     * @param {Function} previewFn - 文字属性を書き換える処理
     * @returns {void}
     */
    function runPreview(previewFn) {
        /* 入力のたびに失敗をアラートしない / Do not alert on every keystroke */
        try { previewFn(); } catch (e) { }
        previewApplied = true;
    }

    /**
     * プレビューを取り消して元の文字属性に戻す
     * @returns {void}
     */
    function cancelPreview() {
        if (!previewApplied) return;
        /* 文字が消えるなどして控えが無効になっていることがある / the snapshots may have gone stale */
        try { restoreCharSnapshots(); } catch (e) { }
        previewApplied = false;
    }

    // -----------------------------------------
    // ダイアログ / Dialog
    // -----------------------------------------

    var touchDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    touchDialog.orientation = "column";
    touchDialog.alignChildren = ["fill", "top"];

    /**
     * 縦並びのパネルを追加する
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} titleKey - パネル見出しのキー（LABELS.panel）
     * @param {string} childAlignment - 子の横方向の揃え（"fill" / "left"）
     * @returns {Panel} 追加したパネル
     */
    function addColumnPanel(parentContainer, titleKey, childAlignment) {
        var addedPanel = parentContainer.add("panel", undefined, getLabel("panel." + titleKey));
        addedPanel.orientation = "column";
        addedPanel.alignChildren = [childAlignment, "top"];
        addedPanel.margins = PANEL_MARGINS;
        return addedPanel;
    }

    /**
     * チェックボックスかボタンを、LABELS のキーで tooltip 付きで追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} controlType - "checkbox" / "button"
     * @param {string} labelKey - LABELS.checkbox（または LABELS.button）と LABELS.tooltip のキー
     * @returns {Checkbox|Button} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var addedControl = parentContainer.add(controlType, undefined, getLabel(controlType + "." + labelKey));
        addedControl.helpTip = getLabel("tooltip." + labelKey);
        return addedControl;
    }

    /* 文字タッチ / Touch panel */
    var touchPanel = addColumnPanel(touchDialog, "touch", "fill");

    /* フォントと「犯行声明文」風を横に並べる / Font and ransom-note panels sit side by side */
    var fontRowGroup = touchDialog.add("group");
    fontRowGroup.orientation = "row";
    fontRowGroup.alignChildren = ["fill", "top"];

    var fontPanel = addColumnPanel(fontRowGroup, "font", "left");
    fontPanel.alignment = ["fill", "top"];

    var chkRandomFont = addLabeledControl(fontPanel, "checkbox", "randomFont");
    var chkJapaneseOnly = addLabeledControl(fontPanel, "checkbox", "japaneseOnly");

    var ransomPanel = addColumnPanel(fontRowGroup, "ransom", "left");
    ransomPanel.alignment = ["fill", "top"];

    var chkRansomEnabled = addLabeledControl(ransomPanel, "checkbox", "ransomEnabled");

    var ransomTrackingGroup = ransomPanel.add("group");
    ransomTrackingGroup.orientation = "row";
    ransomTrackingGroup.alignChildren = ["left", "center"];

    var chkRansomTracking = addLabeledControl(ransomTrackingGroup, "checkbox", "ransomTracking");
    /* ∧∨と入力欄は隙間0で突き合わせる。トラッキングはマイナスも受け付ける
       Butt the stepper against the field; tracking accepts negative values */
    var ransomStepperInputGroup = ransomTrackingGroup.add("group");
    ransomStepperInputGroup.orientation = "row";
    ransomStepperInputGroup.alignChildren = ["left", "center"];
    ransomStepperInputGroup.spacing = 0;
    ransomStepperInputGroup.margins = 0;
    var edtRansomTracking;
    var ransomTrackingStepper = addStepper(ransomStepperInputGroup, function () { return edtRansomTracking; }, {
        onStep: function () { requestPreview(); }
    });
    edtRansomTracking = ransomStepperInputGroup.add("edittext", undefined, String(RANSOM_TRACKING_DEFAULT));
    edtRansomTracking.helpTip = getLabel("tooltip.ransomTrackingValue");
    edtRansomTracking.characters = RANSOM_FIELD_CHARS;
    bindSteppedArrowKeys(edtRansomTracking, ransomTrackingStepper);

    /**
     * トラッキング値の入力欄と∧∨の有効／無効をまとめて切り替え、∧∨を描き直す
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setRansomTrackingFieldEnabled(isEnabled) {
        edtRansomTracking.enabled = isEnabled;
        ransomTrackingStepper.enabled = isEnabled;
        redrawSteppersIn(ransomTrackingStepper);
    }

    chkRandomFont.value = false;
    chkJapaneseOnly.value = false;
    chkRansomEnabled.value = false;
    /* トラッキング調整は既定でON。ただし「有効」がOFFの間は操作できない
       Tracking is on by default but stays dimmed until the style is enabled */
    chkRansomTracking.value = true;
    chkRansomTracking.enabled = false;
    setRansomTrackingFieldEnabled(false);

    /**
     * 「ランダム」がOFFのときは「和文フォントに限定」を無効にする
     * @returns {void}
     */
    function updateFontOptionState() {
        chkJapaneseOnly.enabled = chkRandomFont.value;
        if (!chkRandomFont.value) chkJapaneseOnly.value = false;
    }
    updateFontOptionState();

    /**
     * 英数字以外を含む選択では「和文フォントに限定」を自動でONにする
     * @returns {void}
     */
    function autoEnableJapaneseOnly() {
        if (chkJapaneseOnly.enabled && containsNonAlphanumeric(textRanges)) chkJapaneseOnly.value = true;
    }

    /* ベースラインの単位は環境設定の「文字」の単位に従う / The baseline unit follows the "text/asianunits" preference */
    var baselineUnit = getUnitInfo("text/asianunits");
    var baselineUnitLabel = baselineUnit.label;
    var baselinePointsPerUnit = baselineUnit.pointsPerUnit;

    /* 初期値は文字サイズの1/12、スライダーの上限は文字サイズ / The default is 1/12 of the type size; the slider tops out at the type size */
    var firstCharSizePt = (charSnapshots.length > 0) ? charSnapshots[0].character.characterAttributes.size : 0;
    var defaultBaselineValue = firstCharSizePt ? Math.round(Math.round(firstCharSizePt / 12) / baselinePointsPerUnit) : 0;
    if (isNaN(defaultBaselineValue)) defaultBaselineValue = 0;
    var baselineSliderMax = firstCharSizePt
        ? Math.max(1, Math.round(Math.max(1, Math.round(firstCharSizePt)) / baselinePointsPerUnit))
        : BASELINE_SLIDER_MAX_FALLBACK;

    /**
     * 「チェックボックス＋項目名＋数値欄＋単位＋スライダー」の行を作る
     * @param {Panel} parentPanel - 行を置くパネル
     * @param {string} labelKey - 項目名のキー（tooltipのキーにも使う）
     * @param {string} unitLabel - 単位の表示文字列
     * @param {number} defaultValue - 数値欄の初期値
     * @param {number} minValue - スライダーの下限
     * @param {number} maxValue - スライダーの上限
     * @param {boolean} [allowNegative] - ∧∨・↑↓キーでマイナス値まで下げてよいか
     * @returns {{toggle: Checkbox, label: StaticText, field: EditText, unit: StaticText, slider: Slider}} 作成したコントロール
     */
    function addTouchRow(parentPanel, labelKey, unitLabel, defaultValue, minValue, maxValue, allowNegative) {
        var rowGroup = parentPanel.add("group");

        var enableCheckbox = rowGroup.add("checkbox", undefined, "");
        enableCheckbox.helpTip = getLabel("tooltip." + labelKey + "Enabled");
        enableCheckbox.value = true;
        enableCheckbox.preferredSize.width = TOGGLE_WIDTH;

        var rowLabel = rowGroup.add("statictext", undefined, labelText("fieldLabel." + labelKey));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = rowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var valueField;
        var valueSlider;
        /* 値を変えたらスライダーを追従させ、プレビューを予約する / keep the slider in step and schedule the preview */
        var valueStepper = addStepper(stepperInputGroup, function () { return valueField; }, {
            min: allowNegative ? undefined : 0,
            onStep: function (numberInput) {
                syncSliderFromField(numberInput, valueSlider);
                requestPreview();
            }
        });
        valueField = stepperInputGroup.add("edittext", undefined, String(defaultValue));
        valueField.characters = VALUE_FIELD_CHARS;
        valueField.helpTip = getLabel("tooltip." + labelKey);
        bindSteppedArrowKeys(valueField, valueStepper);

        var unitText = rowGroup.add("statictext", undefined, unitLabel);

        valueSlider = rowGroup.add("slider", undefined, defaultValue, minValue, maxValue);
        valueSlider.preferredSize.width = SLIDER_WIDTH;
        valueSlider.helpTip = getLabel("tooltip." + labelKey);

        return { toggle: enableCheckbox, label: rowLabel, field: valueField, unit: unitText, slider: valueSlider };
    }

    var baselineRow = addTouchRow(touchPanel, "baseline", baselineUnitLabel, defaultBaselineValue, 0, baselineSliderMax);
    var scaleRow = addTouchRow(touchPanel, "scale", "%", DEFAULT_SCALE_PERCENT, 0, SCALE_SLIDER_MAX);
    var kerningRow = addTouchRow(touchPanel, "kerning", "em", DEFAULT_KERNING_EM, KERNING_SLIDER_MIN, KERNING_SLIDER_MAX, true);
    var rotationRow = addTouchRow(touchPanel, "rotation", "°", DEFAULT_ROTATION_DEG, 0, ROTATION_SLIDER_MAX);

    /**
     * 複数のラベルの幅を、いちばん広いものに揃える
     * @param {StaticText[]} labelControls - 幅を揃えるラベル
     * @returns {void}
     */
    function alignLabelWidths(labelControls) {
        var widestWidth = 0;
        var i;
        for (i = 0; i < labelControls.length; i++) {
            if (labelControls[i].preferredSize.width > widestWidth) widestWidth = labelControls[i].preferredSize.width;
        }
        for (i = 0; i < labelControls.length; i++) {
            labelControls[i].preferredSize.width = widestWidth;
            labelControls[i].justify = "left";
        }
    }
    alignLabelWidths([baselineRow.label, scaleRow.label, kerningRow.label, rotationRow.label]);
    alignLabelWidths([baselineRow.unit, scaleRow.unit, kerningRow.unit, rotationRow.unit]);

    /* 文字タッチパネル下部：左に一括ON／OFF、右にトラッキング補正
       Bottom of the touch panel: bulk toggles on the left, tracking compensation on the right */
    var touchButtonRow = touchPanel.add("group");
    touchButtonRow.orientation = "row";
    touchButtonRow.alignChildren = ["fill", "center"];
    touchButtonRow.alignment = "fill";
    touchButtonRow.margins = [0, TOUCH_BUTTON_ROW_TOP_MARGIN, 0, 0];

    var touchButtonLeftGroup = touchButtonRow.add("group");
    touchButtonLeftGroup.orientation = "row";
    touchButtonLeftGroup.alignChildren = ["left", "center"];

    var btnAllOn = addLabeledControl(touchButtonLeftGroup, "button", "allOn");
    btnAllOn.preferredSize = SMALL_BUTTON_SIZE;
    var btnAllOff = addLabeledControl(touchButtonLeftGroup, "button", "allOff");
    btnAllOff.preferredSize = SMALL_BUTTON_SIZE;

    var touchButtonSpacer = touchButtonRow.add("group");
    touchButtonSpacer.alignment = ["fill", "fill"];
    touchButtonSpacer.minimumSize.width = 0;

    var touchButtonRightGroup = touchButtonRow.add("group");
    touchButtonRightGroup.orientation = "row";
    touchButtonRightGroup.alignChildren = ["right", "center"];

    var chkRotationTracking = addLabeledControl(touchButtonRightGroup, "checkbox", "rotationTracking");
    chkRotationTracking.value = true;

    /* ズーム / Zoom row */
    var zoomGroup = touchDialog.add("group");
    zoomGroup.orientation = "row";
    zoomGroup.alignChildren = ["center", "center"];
    zoomGroup.margins = [0, 0, 0, 0];

    zoomGroup.add("statictext", undefined, labelText("fieldLabel.zoom"));
    var sldZoom = zoomGroup.add("slider", undefined, clampZoomPercent(originalZoomFactor * 100), ZOOM_MIN_PERCENT, ZOOM_MAX_PERCENT);
    sldZoom.helpTip = getLabel("tooltip.zoom");
    sldZoom.preferredSize.width = ZOOM_SLIDER_WIDTH;

    var chkLightMode = addLabeledControl(zoomGroup, "checkbox", "lightMode");
    chkLightMode.value = false;

    /* ボタンエリア / Button row */
    var buttonRow = addButtonRow(touchDialog);

    var btnRandomize = addLabeledControl(buttonRow.leftGroup, "button", "randomize");
    var btnReset = addLabeledControl(buttonRow.leftGroup, "button", "reset");

    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    // -----------------------------------------
    // 入力値の同期 / Value syncing
    // -----------------------------------------

    /**
     * 値を範囲内に収める
     * @param {number} inputValue - 対象の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clampToRange(inputValue, minValue, maxValue) {
        if (inputValue < minValue) return minValue;
        if (inputValue > maxValue) return maxValue;
        return inputValue;
    }

    /**
     * 数値欄の値をスライダーに反映する
     * @param {EditText} valueField - 数値欄
     * @param {Slider} valueSlider - スライダー
     * @returns {void}
     */
    function syncSliderFromField(valueField, valueSlider) {
        var fieldValue = parseNumber(valueField.text);
        if (fieldValue === null) return;
        valueSlider.value = clampToRange(fieldValue, valueSlider.minvalue, valueSlider.maxvalue);
    }

    /**
     * スライダーの値を数値欄に反映する
     * @param {Slider} valueSlider - スライダー
     * @param {EditText} valueField - 数値欄
     * @returns {void}
     */
    function syncFieldFromSlider(valueSlider, valueField) {
        valueField.text = String(Math.round(valueSlider.value));
    }

    // -----------------------------------------
    // プレビューの実行 / Running the preview
    // -----------------------------------------

    /* 入力のたびに再描画すると重いので、少し待ってからまとめて反映する
       Redrawing on every keystroke is heavy, so the preview is debounced */
    var previewTaskId = null;

    /* scheduleTask はグローバルに置いた関数名でしか呼べない / scheduleTask can only call a global by name */
    $.global.__AutoTouchTypePreview = function () {
        previewTaskId = null;
        updatePreview();
    };

    /**
     * プレビューの更新を予約する
     * @returns {void}
     */
    function requestPreview() {
        /* scheduleTask が使えない環境ではその場で更新する / Fall back to an immediate update */
        try {
            if (previewTaskId != null) app.cancelTask(previewTaskId);
            previewTaskId = app.scheduleTask("$.global.__AutoTouchTypePreview()", PREVIEW_DELAY_MS, false);
        } catch (e) {
            updatePreview();
        }
    }

    /**
     * 文字タッチ4項目の入力値を読む
     * @returns {{baselinePt: number|null, scalePercent: number|null, rotationDeg: number|null, kerningEm: number|null}} 各項目の最大量（空欄は null）
     */
    function readTouchAmounts() {
        var baselineValue = baselineRow.toggle.value ? parseNumber(baselineRow.field.text) : 0;
        return {
            baselinePt: (baselineValue === null) ? null : baselineValue * baselinePointsPerUnit,
            scalePercent: scaleRow.toggle.value ? parseNumber(scaleRow.field.text) : 0,
            rotationDeg: rotationRow.toggle.value ? parseNumber(rotationRow.field.text) : 0,
            kerningEm: kerningRow.toggle.value ? parseNumber(kerningRow.field.text) : 0
        };
    }

    /**
     * 空欄の項目があるか判定する
     * @param {Object} touchAmounts - readTouchAmounts() の戻り値
     * @returns {boolean} 空欄があるとき true
     */
    function hasEmptyAmount(touchAmounts) {
        return touchAmounts.baselinePt === null || touchAmounts.scalePercent === null ||
            touchAmounts.rotationDeg === null || touchAmounts.kerningEm === null;
    }

    /**
     * 適用オプションを組み立てる
     * @param {boolean} forPreview - プレビュー用かどうか
     * @returns {Object} applyRandomTouch() に渡すオプション
     */
    function buildTouchOptions(forPreview) {
        var ransomTrackingOn = chkRansomEnabled.value && chkRansomTracking.value;
        return {
            randomFont: chkRandomFont.value,
            japaneseOnly: chkJapaneseOnly.value,
            /* 背景の生成はプレビューしない（OK時のみ）/ The backgrounds are created on OK only */
            ransom: forPreview ? false : chkRansomEnabled.value,
            ransomTrack: chkRansomTracking.value,
            /* プレビューではトラッキングの固定値だけ反映する / The preview only reflects the fixed tracking */
            previewRansomTracking: forPreview && ransomTrackingOn,
            ransomTrackValue: null,
            rotationTracking: chkRotationTracking.value,
            allFonts: allFonts,
            japaneseFonts: null,
            skipRedraw: forPreview
        };
    }

    /**
     * 和文フォントに限定するとき、一覧をオプションに入れる
     * @param {Object} touchOptions - 適用オプション
     * @returns {boolean} 続行できるとき true（和文フォントが無いときは false）
     */
    function fillJapaneseFonts(touchOptions) {
        if (!(touchOptions.randomFont && touchOptions.japaneseOnly)) return true;
        touchOptions.japaneseFonts = getJapaneseFonts();
        if (touchOptions.japaneseFonts.length > 0) return true;
        alert(getLabel("alert.noJapaneseFonts"));
        return false;
    }

    /**
     * 現在の設定でプレビューを更新する
     * @returns {void}
     */
    function updatePreview() {
        didReset = false;

        var touchAmounts = readTouchAmounts();
        if (hasEmptyAmount(touchAmounts)) return;

        var touchOptions = buildTouchOptions(true);
        if (!fillJapaneseFonts(touchOptions)) return;
        if (touchOptions.previewRansomTracking) {
            var trackingValue = parseNumber(edtRansomTracking.text);
            touchOptions.ransomTrackValue = (trackingValue === null) ? RANSOM_TRACKING_DEFAULT : trackingValue;
        }

        runPreview(function () {
            applyRandomTouch(touchAmounts, randomSeed, touchOptions);
        });
        app.redraw();
    }

    // -----------------------------------------
    // イベント / Events
    // -----------------------------------------

    var touchRows = [baselineRow, scaleRow, kerningRow, rotationRow];

    /**
     * 文字タッチ4項目がすべてOFFか判定する
     * @returns {boolean} すべてOFFのとき true
     */
    function isTouchAllOff() {
        for (var i = 0; i < touchRows.length; i++) {
            if (touchRows[i].toggle.value) return false;
        }
        return true;
    }

    /**
     * 変化を生まない設定のときは「ランダム」ボタンを無効にする
     * @returns {void}
     */
    function updateRandomizeEnabled() {
        btnRandomize.enabled = !(isTouchAllOff() && !chkRandomFont.value);
    }

    /**
     * 文字タッチ4項目をまとめて切り替える
     * @param {boolean} enabled - ONにするとき true
     * @returns {void}
     */
    function setAllTouchToggles(enabled) {
        for (var i = 0; i < touchRows.length; i++) {
            touchRows[i].toggle.value = enabled;
        }
        updatePreview();
        updateRandomizeEnabled();
    }

    /* 数値欄・スライダー・チェックボックスの操作をプレビューにつなぐ
       Wire the fields, sliders and toggles to the preview */
    for (var rowIndex = 0; rowIndex < touchRows.length; rowIndex++) {
        (function (touchRow) {
            touchRow.field.onChanging = function () {
                syncSliderFromField(touchRow.field, touchRow.slider);
                requestPreview();
            };
            touchRow.slider.onChanging = function () {
                syncFieldFromSlider(touchRow.slider, touchRow.field);
                requestPreview();
            };
            touchRow.toggle.onClick = function () {
                requestPreview();
                updateRandomizeEnabled();
            };
            syncSliderFromField(touchRow.field, touchRow.slider);
        })(touchRows[rowIndex]);
    }

    edtRansomTracking.onChanging = function () { requestPreview(); };

    chkRotationTracking.onClick = function () { requestPreview(); };

    chkRandomFont.onClick = function () {
        updateFontOptionState();
        autoEnableJapaneseOnly();
        requestPreview();
        updateRandomizeEnabled();
    };

    chkJapaneseOnly.onClick = function () {
        requestPreview();
        updateRandomizeEnabled();
    };

    chkRansomEnabled.onClick = function () {
        chkRansomTracking.enabled = chkRansomEnabled.value;
        setRansomTrackingFieldEnabled(chkRansomEnabled.value && chkRansomTracking.value);
        updateRandomizeEnabled();
        requestPreview();
    };

    chkRansomTracking.onClick = function () {
        setRansomTrackingFieldEnabled(chkRansomEnabled.value && chkRansomTracking.value);
        requestPreview();
    };

    btnAllOn.onClick = function () { setAllTouchToggles(true); };
    btnAllOff.onClick = function () { setAllTouchToggles(false); };

    /* ドラッグ中の表示倍率の反映は重いので間引く / Applying the zoom while dragging is heavy, so throttle it */
    var lastZoomAppliedMs = 0;

    sldZoom.onChanging = function () {
        /* 軽量モードではドラッグ中は何もしない / Light mode skips the update while dragging */
        if (chkLightMode.value) return;
        var nowMs = (new Date()).getTime();
        if (nowMs - lastZoomAppliedMs < ZOOM_THROTTLE_MS) return;
        lastZoomAppliedMs = nowMs;
        applyZoomPercent(this.value);
    };

    sldZoom.onChange = function () { applyZoomPercent(this.value); };
    chkLightMode.onClick = function () { applyZoomPercent(sldZoom.value); };

    /**
     * 文字属性をリセットし、控えも初期値にそろえる
     * @returns {void}
     */
    function resetCharAttributes() {
        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var charAttributes = snapshot.character.characterAttributes;

            charAttributes.baselineShift = 0;
            charAttributes.horizontalScale = 100;
            charAttributes.verticalScale = 100;
            charAttributes.rotation = 0;
            snapshot.baselineShift = 0;
            snapshot.horizontalScale = 100;
            snapshot.verticalScale = 100;
            snapshot.rotation = 0;

            /* カーニングとトラッキングは書き込めない環境がある / Kerning and tracking are not writable everywhere */
            try {
                charAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                snapshot.character.kerning = 0;
                snapshot.kerning = 0;
            } catch (e) { }
            try {
                charAttributes.tracking = 0;
                snapshot.tracking = 0;
            } catch (e) { }
        }
    }

    /**
     * すべての文字のフォントを先頭文字のフォントにそろえる
     * @returns {void}
     */
    function unifyFontToFirstChar() {
        if (charSnapshots.length === 0) return;
        var firstFont = charSnapshots[0].character.characterAttributes.textFont;
        if (!firstFont) return;
        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var charContent = snapshot.character.contents;
            if (charContent === "\r" || charContent === "\n") continue;
            /* 環境にないフォントは適用できない / A font missing from the system cannot be applied */
            try {
                snapshot.character.characterAttributes.textFont = firstFont;
                snapshot.textFont = firstFont;
            } catch (e) { }
        }
    }

    btnRandomize.onClick = function () {
        didReset = false;
        randomSeed = (new Date()).getTime() & 0xffffffff;
        requestPreview();
        autoEnableJapaneseOnly();
        updateRandomizeEnabled();
    };

    btnReset.onClick = function () {
        /* チェックボックスの状態は変えず、文字属性だけ初期化する / Only the text attributes are reset */
        cancelPreview();
        clearRansomBgRects();
        resetCharAttributes();
        unifyFontToFirstChar();
        app.redraw();

        updateFontOptionState();
        autoEnableJapaneseOnly();
        updateRandomizeEnabled();
        didReset = true;
    };

    btnOK.onClick = function () {
        var touchAmounts = readTouchAmounts();
        /* ベースラインの空欄は0として扱う / An empty baseline field counts as zero */
        if (touchAmounts.baselinePt === null) touchAmounts.baselinePt = 0;
        if (hasEmptyAmount(touchAmounts)) {
            alert(getLabel("alert.enterNumber"));
            return;
        }

        var touchOptions = buildTouchOptions(false);
        if (touchOptions.ransom && touchOptions.ransomTrack) {
            touchOptions.ransomTrackValue = parseNumber(edtRansomTracking.text);
            if (touchOptions.ransomTrackValue === null) {
                alert(getLabel("alert.enterNumber"));
                return;
            }
        }
        if (!fillJapaneseFonts(touchOptions)) return;

        /* リセット直後でプレビューが無ければ、その結果をそのまま確定する
           Right after Reset with no preview, keep the reset result as is */
        if (!(didReset && !previewApplied)) {
            try {
                applyRandomTouch(touchAmounts, randomSeed, touchOptions);
            } catch (e) {
                /* 選択が変わって控えが無効になったときは取り直してやり直す
                   Re-derive the selection when the snapshots went stale */
                textRanges = collectTextRanges(doc.selection);
                selectedTextFrames = getSelectionTextFrames(doc.selection);
                takeCharSnapshots();
                applyRandomTouch(touchAmounts, randomSeed, touchOptions);
            }
        }

        /* 背景の長方形はOK時だけ作る / The background rectangles are created on OK only */
        if (touchOptions.ransom) createRansomBgRects(selectedTextFrames);
        else clearRansomBgRects();

        closedByOK = true;
        didReset = false;
        touchDialog.close(1);
    };

    btnCancel.onClick = function () {
        cancelPreview();
        clearRansomBgRects();
        restoreOriginalView();
        closedByOK = false;
        touchDialog.close(0);
    };

    touchDialog.onClose = function () {
        /* OKで閉じたときは確定済みなので後始末しない / A close via OK keeps the committed result */
        if (closedByOK) return true;

        cancelPreview();
        clearRansomBgRects();
        restoreOriginalView();
        return true;
    };

    updatePreview();
    updateRandomizeEnabled();
    alignRightOnlyButtonRow(buttonRow);
    prepareDialogWindow(touchDialog, SCRIPT_NAME);
    touchDialog.show();
})();
