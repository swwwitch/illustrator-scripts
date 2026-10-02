#target illustrator
#targetengine "CenterLineConnectorFromRectEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Excelなどからコピー＆ペーストした長方形から、罫線（グリッド・枠線）を自動生成します。
線同士の関係に応じて、格子化・結合・統合まで実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterLineConnectorFromRect.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n0a1c70def387

### Overview

Generates table rules — a grid and a border — from rectangles pasted in from Excel or similar.
Depending on how the lines relate to each other, it also lattices, joins and merges them.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CenterLineConnectorFromRect.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CenterLineConnectorFromRect";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterLineConnectorFromRect.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CenterLineConnectorFromRect.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0a1c70def387"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // バージョンとローカライズ / Version & Localization
    // =========================================

    var DEFAULT_STROKE_WIDTH_PT = 0.25;

    // =========================================
    // レイヤー設定 / Layer Settings
    // =========================================

    var GENERATED_WORK_LAYER_NAME = "__generated_center_line__";
    var EXCLUSION_PREVIEW_LAYER_NAME = "__center_line_preview__";

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

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () {
            numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
            numberInput.active = true;
        });

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "長方形を線に変換・連結", en: "Convert Rectangles to Lines & Connect" },
        lblShortSideLength: { ja: "短辺が", en: "Short side is" },
        lblGreaterEqual: { ja: "以上の長方形は対象外", en: "below threshold — skip" },
        pnlCenterLineConversion: { ja: "中心線化", en: "Center Line Conversion" },
        chkAngleCorrect: { ja: "角度補正", en: "Correct rotation" },
        pnlOption: { ja: "線の連結・調整", en: "Line connection & adjustment" },
        pnlRule: { ja: "罫線", en: "Rule" },
        chkConnectAll: { ja: "連結", en: "Connect" },
        rdoConnectNone: { ja: "なし", en: "None" },
        rdoConnectGrid: { ja: "格子化", en: "Grid" },
        chkOuterRect: { ja: "外枠を長方形に", en: "Outer frame as rectangle" },
        chkPrintBlack: { ja: "印刷用の「黒」（K100）にする", en: "Use print black (K100)" },
        lblStrokePref: { ja: "線幅", en: "Stroke Width" },
        chkCommonStroke: { ja: "線幅を共通にする", en: "Make stroke widths common" },
        chkGroup: { ja: "グループ化", en: "Group result" },
        chkReturnToOriginal: { ja: "元のレイヤーに戻す", en: "Return to the original layer" },
        tipCenterLine: { ja: "細長い長方形を、その中心を通る1本の線に置き換えます。", en: "Replaces each long thin rectangle with a single line down its middle." },
        tipAngleCorrect: { ja: "わずかに傾いた線を、水平・垂直にそろえ直します。", en: "Straightens lines that are only slightly off horizontal or vertical." },
        tipMinShortSide: { ja: "短辺がこの値以下の長方形だけを線に置き換えます。", en: "Only rectangles whose short side is at most this value become lines." },
        tipConnectNone: { ja: "線どうしはつなぎません。", en: "Leaves the lines unconnected." },
        tipConnectAll: { ja: "交差しうる線をすべて延長してつなぎます。", en: "Extends every line that could meet another and joins them." },
        tipConnectGrid: { ja: "格子状に並んでいる線だけをつなぎます。", en: "Joins only the lines that form a grid." },
        tipOuterRect: { ja: "外周を囲む長方形も描きます。", en: "Also draws the rectangle that frames the whole set." },
        tipPrintBlack: { ja: "線の色をスミ100%（K100）にします。", en: "Sets the line color to 100% black (K100)." },
        tipStrokeMax: { ja: "元の長方形の短辺のうち、いちばん太いものに線幅をそろえます。", en: "Uses the thickest of the original short sides as the stroke weight." },
        tipStrokeMin: { ja: "元の長方形の短辺のうち、いちばん細いものに線幅をそろえます。", en: "Uses the thinnest of the original short sides as the stroke weight." },
        tipStrokeAvg: { ja: "元の長方形の短辺の平均を線幅にします。", en: "Uses the average of the original short sides as the stroke weight." },
        tipStrokeCustom: { ja: "線幅を数値で指定します。", en: "Sets the stroke weight to a value you type." },
        tipCommonStroke: { ja: "すべての線を同じ太さにそろえます。オフだと元の太さを保ちます。", en: "Gives every line the same weight. Off keeps their original weights." },
        tipGroup: { ja: "作った線を1つのグループにまとめます。", en: "Groups the resulting lines together." },
        tipReturnToOriginal: { ja: "作った線を、元の長方形があったレイヤーへ戻します。", en: "Moves the new lines back onto the layer the rectangles came from." },
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        strokeMax: { ja: "最大", en: "Max" },
        strokeMin: { ja: "最小", en: "Min" },
        strokeAvg: { ja: "平均", en: "Average" },
        strokeCustom: { ja: "指定", en: "Custom" },
        btnOutlineOn: { ja: "アウトライン表示", en: "Outline View" },
        btnOutlineOff: { ja: "プレビュー表示", en: "Preview View" },
        btnOk: { ja: "OK", en: "OK" },
        btnCancel: { ja: "キャンセル", en: "Cancel" },
        alertNoSelection: { ja: "長方形を1つ以上選択してください。", en: "Please select at least one rectangle." },
        alertError: { ja: "エラーが発生しました", en: "An error occurred" }
    };

    // =========================================
    // 安全操作ヘルパー / Safe Operation Helpers
    // =========================================

    /* ロック／非表示なら true / True when item is locked or hidden */
    function isLockedOrHidden(pageItem) {
        try {
            return !!(pageItem.locked || pageItem.hidden);
        } catch (e) {
            return false;
        }
    }

    /* PageItem を安全に削除 / Safely remove a PageItem */
    function removePageItemSafely(pageItem) {
        try { pageItem.remove(); } catch (e) { }
    }

    /* PageItem を安全に移動 / Safely move a PageItem */
    function movePageItemSafely(pageItem, destination, placement) {
        try { pageItem.move(destination, placement); } catch (e) { }
    }

    /* 線端形状を安全に設定 / Safely set stroke cap */
    function setStrokeCapSafely(pathItem, strokeCap) {
        try { pathItem.strokeCap = strokeCap; } catch (e) { }
    }

    /* 選択を安全に復元 / Safely restore document selection */
    function restoreSelectionSafely(selectionItems) {
        try { app.activeDocument.selection = selectionItems; } catch (e) { }
    }

    // =========================================
    // 単位変換 / Unit Conversion
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

    /* pt値を指定単位の表示値へ変換 / Convert points to display unit value */
    function pointsToUnitValue(points, unitCode) {
        return points / (UNITS[unitCode] || UNITS[2]).pointsPerUnit;
    }

    /* 指定単位の入力値をptへ変換 / Convert display unit value to points */
    function unitValueToPoints(value, unitCode) {
        return value * (UNITS[unitCode] || UNITS[2]).pointsPerUnit;
    }

    function createExclusionMarkerColor() {
        var markerColor = new CMYKColor();
        markerColor.cyan = 0; markerColor.magenta = 100; markerColor.yellow = 100; markerColor.black = 0;
        return markerColor;
    }

    function createPrintBlackColor() {
        var blackColor = new CMYKColor();
        blackColor.cyan = 0;
        blackColor.magenta = 0;
        blackColor.yellow = 0;
        blackColor.black = 100;
        return blackColor;
    }

    function ensureExclusionPreviewLayer() {
        var doc = app.activeDocument;
        var layer;
        try {
            layer = doc.layers.getByName(EXCLUSION_PREVIEW_LAYER_NAME);
        } catch (e) {
            layer = addLayerSafely(EXCLUSION_PREVIEW_LAYER_NAME);
        }
        layer.locked = false;
        layer.visible = true;
        return layer;
    }

    /* レイヤー作成時に最前面レイヤーがロック／非表示でも失敗しないよう一時解除
       Temporarily unlock/show the top layer so a new layer can be created safely */
    function addLayerSafely(layerName) {
        var doc = app.activeDocument;
        var previousActiveLayer = null;
        var topLayer = null;
        var topWasLocked = false;
        var topWasHidden = false;
        var newLayer = null;

        try { previousActiveLayer = doc.activeLayer; } catch (eA) { }
        try {
            topLayer = doc.layers[0];
            topWasLocked = topLayer.locked;
            topWasHidden = !topLayer.visible;
            if (topWasLocked) topLayer.locked = false;
            if (topWasHidden) topLayer.visible = true;
        } catch (eB) { }

        try {
            newLayer = doc.layers.add();
            newLayer.name = layerName;
        } finally {
            try {
                if (topLayer) {
                    if (topWasLocked) topLayer.locked = true;
                    if (topWasHidden) topLayer.visible = false;
                }
            } catch (eC) { }
            if (previousActiveLayer) {
                try { doc.activeLayer = previousActiveLayer; } catch (eD) { }
            }
        }

        return newLayer;
    }

    /* 生成専用の安全な作業レイヤーを取得／作成
       Get or create a writable layer dedicated to generated results */
    function ensureGeneratedWorkLayer() {
        var doc = app.activeDocument;
        var layer;
        try {
            layer = doc.layers.getByName(GENERATED_WORK_LAYER_NAME);
        } catch (e) {
            layer = addLayerSafely(GENERATED_WORK_LAYER_NAME);
        }
        layer.locked = false;
        layer.visible = true;
        return layer;
    }

    /* activeLayer や最前面レイヤーに依存せず、安全な生成専用レイヤーに PathItem を作成
       Create a PathItem on the dedicated writable layer instead of relying on active/top layer state */
    function createGeneratedPathItem() {
        return ensureGeneratedWorkLayer().pathItems.add();
    }

    /* 生成結果をまとめる安全なグループを作成
       Create a safe group for generated results */
    function createGeneratedGroup() {
        return ensureGeneratedWorkLayer().groupItems.add();
    }

    function clearExclusionPreviewLayer() {
        try {
            var layer = app.activeDocument.layers.getByName(EXCLUSION_PREVIEW_LAYER_NAME);
            layer.remove();
        } catch (e) { }
    }

    /* グループを再帰的に展開して2点パス（線分）を収集
       Recursively walk groups and collect 2-point paths (line segments) */
    function collectLinesRecursive(items, results) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var pageItem = items[itemIndex];
            if (isLockedOrHidden(pageItem)) continue;
            if (pageItem.typename === "PathItem") {
                if (pageItem.pathPoints.length === 2) results.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                collectLinesRecursive(pageItem.pageItems, results);
            }
        }
    }

    /* グループを再帰的に展開して4点閉じ長方形を収集
       Recursively walk groups and collect 4-point closed rectangles */
    function collectRectsRecursive(items, results) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var pageItem = items[itemIndex];
            if (isLockedOrHidden(pageItem)) continue;
            if (pageItem.typename === "PathItem") {
                if (pageItem.closed && pageItem.pathPoints.length === 4) results.push(pageItem);
            } else if (pageItem.typename === "CompoundPathItem") {
                if (pageItem.pathItems.length === 1) {
                    var compoundSubPath = pageItem.pathItems[0];
                    if (compoundSubPath.closed && compoundSubPath.pathPoints.length === 4) results.push(compoundSubPath);
                }
            } else if (pageItem.typename === "GroupItem") {
                collectRectsRecursive(pageItem.pageItems, results);
            }
        }
    }

    /* convertRectToCenterLine と同じ除外条件を判定 / Same exclusion criteria as convertRectToCenterLine */

    function isExcludedRect(rect, minShortSidePt) {
        var bounds = rect.geometricBounds;
        var rectWidth = bounds[2] - bounds[0];
        var rectHeight = bounds[1] - bounds[3];
        var diffRatio = Math.abs(rectWidth - rectHeight) / Math.max(rectWidth, rectHeight);
        if (diffRatio < 0.05) return true;
        var shortSide = Math.min(rectWidth, rectHeight);
        var longSide = Math.max(rectWidth, rectHeight);
        if (shortSide * 1.5 > longSide) return true;
        if (shortSide < minShortSidePt) return true;
        return false;
    }

    /* 選択中の長方形から、短辺の最大値を指定単位で取得
       Get the largest short side among selected rectangles in the given unit */
    function getMaxShortSideInUnitFromItems(items, unitCode) {
        var rectangleItems = [];
        collectRectsRecursive(items, rectangleItems);
        var maxShortSidePt = 0;

        for (var rectIndex = 0; rectIndex < rectangleItems.length; rectIndex++) {
            try {
                var bounds = rectangleItems[rectIndex].geometricBounds;
                var rectWidth = Math.abs(bounds[2] - bounds[0]);
                var rectHeight = Math.abs(bounds[1] - bounds[3]);
                var shortSide = Math.min(rectWidth, rectHeight);
                if (shortSide > maxShortSidePt) maxShortSidePt = shortSide;
            } catch (e) { }
        }

        return Math.round(pointsToUnitValue(maxShortSidePt, unitCode) * 10) / 10;
    }

    /* 対象外オブジェクトを M100Y100・半透明で複製してマーカー表示
       Duplicate excluded objects as M100Y100 semi-transparent markers */
    function refreshExclusionPreview(items, minShortSidePt) {
        clearExclusionPreviewLayer();
        if (!items || items.length === 0) { app.redraw(); return; }
        var layer = ensureExclusionPreviewLayer();
        var markerColor = createExclusionMarkerColor();

        var rectangleItems = [];
        collectRectsRecursive(items, rectangleItems);

        for (var rectangleIndex = 0; rectangleIndex < rectangleItems.length; rectangleIndex++) {
            if (!isExcludedRect(rectangleItems[rectangleIndex], minShortSidePt)) continue;
            try {
                var markerRect = rectangleItems[rectangleIndex].duplicate(layer, ElementPlacement.PLACEATBEGINNING);
                markerRect.filled = true;
                markerRect.fillColor = markerColor;
                markerRect.stroked = false;
                markerRect.opacity = 50;
            } catch (e) { }
        }
        app.redraw();
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
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
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
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
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 中心線パネルを構築 / Build center-line panel */
    function buildCenterLinePanel(parent, rulerUnitCode, rulerUnitLabel) {
        var panel = parent.add("panel", undefined, getLabel('pnlCenterLineConversion'));
        setupPanel(panel, 6);
        panel.alignChildren = ["left", "top"];

        var initialRectsForDetect = [];
        collectRectsRecursive(app.activeDocument.selection, initialRectsForDetect);
        var hasRectsInSelection = initialRectsForDetect.length > 0;

        var cbCenterLine = panel.add("checkbox", undefined, getLabel('pnlCenterLineConversion'));
        cbCenterLine.helpTip = getLabel('tipCenterLine');
        cbCenterLine.value = hasRectsInSelection;

        var cbAngleCorrect = panel.add("checkbox", undefined, getLabel('chkAngleCorrect'));
        cbAngleCorrect.helpTip = getLabel('tipAngleCorrect');
        cbAngleCorrect.value = true;

        var SHORT_MIN = 0;
        var SHORT_MAX = getMaxShortSideInUnitFromItems(app.activeDocument.selection, rulerUnitCode);
        if (SHORT_MAX <= 0) SHORT_MAX = 10;
        var SHORT_DEFAULT = Math.min(10, SHORT_MAX);

        var minShortSideGroup = panel.add("group");
        minShortSideGroup.orientation = "row";
        minShortSideGroup.alignChildren = "center";
        minShortSideGroup.spacing = 4;

        var cbMinShortSide = minShortSideGroup.add("checkbox", undefined, "");
        cbMinShortSide.helpTip = getLabel('tipMinShortSide');
        cbMinShortSide.value = false;
        minShortSideGroup.add("statictext", undefined, getLabel('lblShortSideLength'));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var minShortSideStepperRow = minShortSideGroup.add("group");
        minShortSideStepperRow.orientation = "row";
        minShortSideStepperRow.alignChildren = ["left", "center"];
        minShortSideStepperRow.spacing = 0;
        minShortSideStepperRow.margins = 0;

        var minShortSideInput;
        /* 範囲（SHORT_MIN〜SHORT_MAX）・丸め・スライダー連動・プレビューは既存の onChange に任せる / reuse the existing onChange */
        var minShortSideStepper = addStepper(minShortSideStepperRow, function () { return minShortSideInput; }, {
            onStep: function (numberInput) { if (numberInput.onChange) numberInput.onChange(); }
        });
        minShortSideInput = minShortSideStepperRow.add("edittext", undefined, SHORT_DEFAULT.toFixed(1));
        minShortSideInput.helpTip = getLabel('tipMinShortSide');
        minShortSideInput.characters = 5;
        bindSteppedArrowKeys(minShortSideInput, minShortSideStepper);

        minShortSideGroup.add("statictext", undefined, rulerUnitLabel + getLabel('lblGreaterEqual'));

        var minShortSideSlider = panel.add("slider", undefined, SHORT_DEFAULT, SHORT_MIN, SHORT_MAX);
        minShortSideSlider.helpTip = getLabel('tipMinShortSide');
        minShortSideSlider.alignment = ["center", "center"];
        minShortSideSlider.preferredSize.width = 247;

        return {
            panel: panel,
            cbCenterLine: cbCenterLine,
            cbAngleCorrect: cbAngleCorrect,
            SHORT_MIN: SHORT_MIN,
            SHORT_MAX: SHORT_MAX,
            SHORT_DEFAULT: SHORT_DEFAULT,
            cbMinShortSide: cbMinShortSide,
            minShortSideInput: minShortSideInput,
            minShortSideStepper: minShortSideStepper,
            minShortSideSlider: minShortSideSlider
        };
    }

    /* 連結オプションパネルを構築 / Build connect options panel */
    function buildConnectOptionsPanel(parent) {
        var panel = parent.add("panel", undefined, getLabel('pnlOption'));
        setupPanel(panel, 6);

        var radioRow = panel.add("group");
        radioRow.orientation = "row";
        radioRow.alignChildren = "center";
        radioRow.spacing = 12;

        var rbConnectNone = radioRow.add("radiobutton", undefined, getLabel('rdoConnectNone'));
        rbConnectNone.helpTip = getLabel('tipConnectNone');
        var rbConnectAll = radioRow.add("radiobutton", undefined, getLabel('chkConnectAll'));
        rbConnectAll.helpTip = getLabel('tipConnectAll');
        var rbConnectGrid = radioRow.add("radiobutton", undefined, getLabel('rdoConnectGrid'));
        rbConnectGrid.helpTip = getLabel('tipConnectGrid');

        var cbOuterRect = radioRow.add("checkbox", undefined, getLabel('chkOuterRect'));
        cbOuterRect.helpTip = getLabel('tipOuterRect');
        cbOuterRect.value = false;
        rbConnectGrid.value = true;

        return {
            panel: panel,
            rbConnectNone: rbConnectNone,
            rbConnectAll: rbConnectAll,
            rbConnectGrid: rbConnectGrid,
            cbOuterRect: cbOuterRect,
            connectRadios: [rbConnectNone, rbConnectAll, rbConnectGrid]
        };
    }

    /* 罫線パネルを構築 / Build rule panel */
    function buildRulePanel(parent, strokeUnitLabel) {
        var panel = parent.add("panel", undefined, getLabel('pnlRule'));
        setupPanel(panel, 6);

        var cbPrintBlack = panel.add("checkbox", undefined, getLabel('chkPrintBlack'));
        cbPrintBlack.helpTip = getLabel('tipPrintBlack');
        cbPrintBlack.value = false;

        var strokeWidthPanel = panel.add("panel", undefined, getLabel('lblStrokePref'));
        setupPanel(strokeWidthPanel, 6);
        strokeWidthPanel.alignChildren = ["left", "top"];

        var strokeRadioGroup = strokeWidthPanel.add("group");
        strokeRadioGroup.orientation = "row";
        strokeRadioGroup.alignChildren = "center";
        strokeRadioGroup.spacing = 8;

        var rbStrokeMax = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeMax'));
        rbStrokeMax.helpTip = getLabel('tipStrokeMax');
        var rbStrokeMin = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeMin'));
        rbStrokeMin.helpTip = getLabel('tipStrokeMin');
        var rbStrokeAvg = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeAvg'));
        rbStrokeAvg.helpTip = getLabel('tipStrokeAvg');
        var rbStrokeCustom = strokeRadioGroup.add("radiobutton", undefined, getLabel('strokeCustom'));
        rbStrokeCustom.helpTip = getLabel('tipStrokeCustom');

        rbStrokeCustom.value = true;

        /* ∧∨と入力欄は別 group に入れる。ラジオは strokeRadioGroup に並んだままなので排他は崩れない
           the stepper and field get their own group; the radios stay together in strokeRadioGroup */
        var customStrokeStepperRow = strokeRadioGroup.add("group");
        customStrokeStepperRow.orientation = "row";
        customStrokeStepperRow.alignChildren = ["left", "center"];
        customStrokeStepperRow.spacing = 0;
        customStrokeStepperRow.margins = 0;

        var customStrokeInput;
        /* 0以下は実行時に既定値へ戻されるため、下限を置く / keep above zero (non-positive values fall back to the default) */
        var customStrokeStepper = addStepper(customStrokeStepperRow, function () { return customStrokeInput; }, { min: 0.01 });
        customStrokeInput = customStrokeStepperRow.add("edittext", undefined, DEFAULT_STROKE_WIDTH_PT.toString());
        customStrokeInput.helpTip = getLabel('tipStrokeCustom');
        customStrokeInput.characters = 4;
        bindSteppedArrowKeys(customStrokeInput, customStrokeStepper);

        strokeRadioGroup.add("statictext", undefined, strokeUnitLabel);

        var cbCommonStroke = strokeWidthPanel.add("checkbox", undefined, getLabel('chkCommonStroke'));
        cbCommonStroke.helpTip = getLabel('tipCommonStroke');
        cbCommonStroke.value = true;

        return {
            panel: panel,
            cbPrintBlack: cbPrintBlack,
            rbStrokeMax: rbStrokeMax,
            rbStrokeMin: rbStrokeMin,
            rbStrokeAvg: rbStrokeAvg,
            rbStrokeCustom: rbStrokeCustom,
            customStrokeInput: customStrokeInput,
            customStrokeStepper: customStrokeStepper,
            cbCommonStroke: cbCommonStroke
        };
    }

    /* 事後処理パネルを構築 / Build post-process panel */
    function buildPostProcessPanel(parent) {
        var panel = parent.add("panel", undefined, "事後処理");
        setupPanel(panel, 6);
        panel.alignChildren = ["left", "top"];

        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.spacing = 12;

        var cbGroup = row.add("checkbox", undefined, getLabel('chkGroup'));
        cbGroup.helpTip = getLabel('tipGroup');
        cbGroup.value = true;

        var cbReturnToOriginal = row.add("checkbox", undefined, getLabel('chkReturnToOriginal'));
        cbReturnToOriginal.helpTip = getLabel('tipReturnToOriginal');
        cbReturnToOriginal.value = false;

        return {
            panel: panel,
            cbGroup: cbGroup,
            cbReturnToOriginal: cbReturnToOriginal
        };
    }

    /* ダイアログUI構築 / Build dialog UI */
    function buildOptionDialogUI(dialog) {
        var strokeUnit = getUnitInfo("strokeUnits");
        var strokeUnitCode = strokeUnit.code;
        var strokeUnitLabel = strokeUnit.label;
        var rulerUnit = getUnitInfo("rulerType");
        var rulerUnitCode = rulerUnit.code;
        var rulerUnitLabel = rulerUnit.label;

        var panelColumn = dialog.add("group");
        panelColumn.orientation = "column";
        panelColumn.alignChildren = "fill";
        panelColumn.spacing = 12;

        var center = buildCenterLinePanel(panelColumn, rulerUnitCode, rulerUnitLabel);
        var connect = buildConnectOptionsPanel(panelColumn);
        var rule = buildRulePanel(panelColumn, strokeUnitLabel);
        var post = buildPostProcessPanel(panelColumn);
        /* ボタン行（左：アウトライン切替／右：キャンセル・OK）/ Button row: outline toggle on the left, Cancel/OK on the right */
        var buttonRow = addButtonRow(dialog);
        var btnOutlineToggle = buttonRow.leftGroup.add("button", undefined, getLabel('btnOutlineOn'));
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('btnCancel'), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('btnOk'), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        var snapshotSelection = [];
        for (var ssi = 0; ssi < app.activeDocument.selection.length; ssi++) {
            snapshotSelection.push(app.activeDocument.selection[ssi]);
        }

        return {
            strokeUnitCode: strokeUnitCode,
            rulerUnitCode: rulerUnitCode,

            cbCenterLine: center.cbCenterLine,
            cbAngleCorrect: center.cbAngleCorrect,
            SHORT_MIN: center.SHORT_MIN,
            SHORT_MAX: center.SHORT_MAX,
            SHORT_DEFAULT: center.SHORT_DEFAULT,
            cbMinShortSide: center.cbMinShortSide,
            minShortSideInput: center.minShortSideInput,
            minShortSideStepper: center.minShortSideStepper,
            minShortSideSlider: center.minShortSideSlider,

            snapshotSelection: snapshotSelection,

            rbConnectNone: connect.rbConnectNone,
            rbConnectAll: connect.rbConnectAll,
            rbConnectGrid: connect.rbConnectGrid,
            cbOuterRect: connect.cbOuterRect,
            connectRadios: connect.connectRadios,
            cbReturnToOriginal: post.cbReturnToOriginal,

            cbPrintBlack: rule.cbPrintBlack,
            rbStrokeMax: rule.rbStrokeMax,
            rbStrokeMin: rule.rbStrokeMin,
            rbStrokeAvg: rule.rbStrokeAvg,
            rbStrokeCustom: rule.rbStrokeCustom,
            customStrokeInput: rule.customStrokeInput,
            customStrokeStepper: rule.customStrokeStepper,
            cbCommonStroke: rule.cbCommonStroke,
            cbGroup: post.cbGroup,

            btnOutlineToggle: btnOutlineToggle,
            isOutlineMode: false
        };
    }

    /* UIから値を読み取る / Read values from UI */
    function readOptionDialogValues(ui) {
        var strokeStrategy;
        if (ui.rbStrokeCustom.value) strokeStrategy = "custom";
        else if (ui.rbStrokeMin.value) strokeStrategy = "min";
        else if (ui.rbStrokeAvg.value) strokeStrategy = "avg";
        else strokeStrategy = "max";

        var customStrokeWidth = parseFloat(ui.customStrokeInput.text);
        if (isNaN(customStrokeWidth) || customStrokeWidth <= 0) customStrokeWidth = DEFAULT_STROKE_WIDTH_PT;

        // Strengthen NaN handling for minShortSideInput
        var minShortSideValue = parseFloat(ui.minShortSideInput.text);
        if (isNaN(minShortSideValue)) minShortSideValue = ui.SHORT_DEFAULT;
        if (minShortSideValue < ui.SHORT_MIN) minShortSideValue = ui.SHORT_MIN;
        if (minShortSideValue > ui.SHORT_MAX) minShortSideValue = ui.SHORT_MAX;

        var connectMode = ui.rbConnectGrid.value ? "grid" : (ui.rbConnectAll.value ? "connect" : "none");

        return {
            correctRotation: ui.cbAngleCorrect.value,
            enableCenterLine: ui.cbCenterLine.value,
            minShortSidePt: ui.cbMinShortSide.value ? unitValueToPoints(minShortSideValue, ui.rulerUnitCode) : 0,
            connectMode: connectMode,
            outerRect: ui.cbOuterRect.value,
            printBlack: ui.cbPrintBlack.value,
            commonStroke: ui.cbCommonStroke.value,
            strokeStrategy: strokeStrategy,
            customStrokeWidth: unitValueToPoints(customStrokeWidth, ui.strokeUnitCode),
            groupResult: ui.cbGroup.value,
            returnToOriginal: ui.cbReturnToOriginal.value
        };
    }

    /* UI状態をまとめて同期 / Synchronize all UI enabled states */
    function syncUIState(ui) {
        var centerLineActive = ui.cbCenterLine.value;
        var minShortSideActive = centerLineActive && ui.cbMinShortSide.value;
        var gridActive = ui.rbConnectGrid.value;

        ui.cbMinShortSide.enabled = centerLineActive;
        ui.cbAngleCorrect.enabled = centerLineActive;
        ui.minShortSideInput.enabled = minShortSideActive;
        ui.minShortSideStepper.enabled = minShortSideActive;
        ui.minShortSideSlider.enabled = minShortSideActive;

        ui.cbOuterRect.enabled = gridActive;
        if (ui.cbGroup) ui.cbGroup.enabled = gridActive;

        ui.customStrokeInput.enabled = ui.rbStrokeCustom.value;
        ui.customStrokeStepper.enabled = ui.rbStrokeCustom.value;

        /* ∧∨は自作描画なので、有効／無効の切替後に描き直す / redraw the custom-drawn steppers */
        redrawSteppersIn(ui.minShortSideStepper);
        redrawSteppersIn(ui.customStrokeStepper);
    }

    /* ダイアログ表示前の初期状態を同期 / Initialize dialog state before showing */
    function initializeOptionDialogState(ui) {
        syncUIState(ui);
        updateExclusionPreviewFromUIState(ui);
    }

    /* UI状態からプレビューを更新 / Update preview from UI state */
    function updateExclusionPreviewFromUIState(ui) {
        if (!ui.cbCenterLine.value) {
            clearExclusionPreviewLayer();
            app.redraw();
            return;
        }

        var minShortSidePt = 0;
        if (ui.cbMinShortSide.value) {
            var minShortSideValue = parseFloat(ui.minShortSideInput.text);
            if (isNaN(minShortSideValue)) minShortSideValue = ui.SHORT_DEFAULT;
            minShortSidePt = unitValueToPoints(minShortSideValue, ui.rulerUnitCode);
        }

        refreshExclusionPreview(ui.snapshotSelection, minShortSidePt);
    }

    /* ダイアログイベントを配線 / Bind dialog events */
    function bindOptionDialogEvents(ui) {

        function roundTo1(value) {
            return Math.round(value * 10) / 10;
        }

        function selectConnectRadio(selected) {
            for (var radioIndex = 0; radioIndex < ui.connectRadios.length; radioIndex++) {
                ui.connectRadios[radioIndex].value = (ui.connectRadios[radioIndex] === selected);
            }
        }

        ui.cbMinShortSide.onClick = function () {
            syncUIState(ui);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.cbCenterLine.onClick = function () {
            syncUIState(ui);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.minShortSideSlider.onChanging = function () {
            ui.minShortSideInput.text = roundTo1(ui.minShortSideSlider.value).toFixed(1);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.minShortSideInput.onChange = function () {
            var minShortSideValue = parseFloat(ui.minShortSideInput.text);
            if (isNaN(minShortSideValue)) minShortSideValue = ui.minShortSideSlider.value;
            if (minShortSideValue < ui.SHORT_MIN) minShortSideValue = ui.SHORT_MIN;
            if (minShortSideValue > ui.SHORT_MAX) minShortSideValue = ui.SHORT_MAX;
            ui.minShortSideSlider.value = minShortSideValue;
            ui.minShortSideInput.text = roundTo1(minShortSideValue).toFixed(1);
            updateExclusionPreviewFromUIState(ui);
        };

        ui.rbConnectNone.onClick = function () {
            selectConnectRadio(ui.rbConnectNone);
            syncUIState(ui);
        };

        ui.rbConnectAll.onClick = function () {
            selectConnectRadio(ui.rbConnectAll);
            syncUIState(ui);

            // 連結処理時は線幅を「最小」に強制
            ui.rbStrokeMin.value = true;
            ui.rbStrokeMax.value = false;
            ui.rbStrokeAvg.value = false;
            ui.rbStrokeCustom.value = false;
        };

        ui.rbConnectGrid.onClick = function () {
            selectConnectRadio(ui.rbConnectGrid);
            syncUIState(ui);
            ui.rbStrokeCustom.value = true;
            syncUIState(ui);
            ui.cbCommonStroke.value = true;
        };

        ui.rbStrokeMax.onClick = function () { syncUIState(ui); };
        ui.rbStrokeMin.onClick = function () { syncUIState(ui); };
        ui.rbStrokeAvg.onClick = function () { syncUIState(ui); };
        ui.rbStrokeCustom.onClick = function () {
            syncUIState(ui);
            ui.cbCommonStroke.value = true;
        };

        ui.btnOutlineToggle.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                ui.isOutlineMode = !ui.isOutlineMode;
                ui.btnOutlineToggle.text = ui.isOutlineMode ? getLabel('btnOutlineOff') : getLabel('btnOutlineOn');
            } catch (e) { }
        };
    }

    /* オプション設定用ダイアログを表示 / Show options dialog */
    function showOptionDialog() {
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        setupWindow(dialog);

        var ui = buildOptionDialogUI(dialog);

        /* イベント配線 / Bind dialog events */
        bindOptionDialogEvents(ui);

        /* 初期状態の同期 / Synchronize initial UI state */
        initializeOptionDialogState(ui);

        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();

        // Clean up preview layer and restore selection
        clearExclusionPreviewLayer();
        restoreSelectionSafely(ui.snapshotSelection);
        app.redraw();

        if (dialogResult !== 1) return null;

        return readOptionDialogValues(ui);
    }

    // =========================================
    // 中心線描画処理 / Center Line Drawing
    // =========================================

    /* 長方形を中心線に変換し、元の長方形を削除 / Convert rectangle into a center line and remove original */
    function convertRectToCenterLine(rect, correctRotation, minShortSidePt) {
        /* 回転補正：最初の辺の角度を測り、最寄りの 90° 軸（水平 0/180、垂直 90/270）から
           0.5°〜10° 以内のずれなら、その軸にスナップする
           Rotation correction: measure first-edge angle and, if it deviates 0.5°–10° from
           the nearest 90° axis (horizontal 0/180, vertical 90/270), snap to that axis */
        if (correctRotation && rect.pathPoints.length === 4) {
            var firstAnchor = rect.pathPoints[0].anchor;
            var secondAnchor = rect.pathPoints[1].anchor;

            var deltaX = secondAnchor[0] - firstAnchor[0];
            var deltaY = secondAnchor[1] - firstAnchor[1];
            var angleRad = Math.atan2(deltaY, deltaX);
            var angleDeg = angleRad * 180 / Math.PI;
            if (angleDeg < 0) angleDeg += 360;

            var nearestAxisDeg = Math.round(angleDeg / 90) * 90;
            var deviationDeg = angleDeg - nearestAxisDeg;
            var absDeviationDeg = Math.abs(deviationDeg);

            if (absDeviationDeg >= 0.5 && absDeviationDeg <= 10) {
                rect.rotate(-deviationDeg);
            }
        }

        /* 補正後にバウンディングボックスを取得 / Get bounding box after correction */
        var bounds = rect.geometricBounds;
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        var rectWidth = right - left;
        var rectHeight = top - bottom;

        /* 正方形に近い場合は除外（5%未満の差）/ Skip near-square shapes (<5% diff) */
        var diffRatio = Math.abs(rectWidth - rectHeight) / Math.max(rectWidth, rectHeight);
        if (diffRatio < 0.05) return null;
        /* 短辺×1.5 ＞ 長辺の場合は除外 / Skip when short × 1.5 > long */
        var shortSide = Math.min(rectWidth, rectHeight);
        var longSide = Math.max(rectWidth, rectHeight);
        if (shortSide * 1.5 > longSide) return null;
        /* 短辺が指定値未満の場合は除外（指定値以降のみ直線化） / Skip when short side is below the threshold (only linearize at/above) */
        if (typeof minShortSidePt === "number" && shortSide < minShortSidePt) return null;

        /* 中心線を生成 / Create center line */
        var centerLine = createGeneratedPathItem();
        centerLine.stroked = true;
        centerLine.filled = false;
        centerLine.strokeColor = rect.fillColor;

        var startPoint = centerLine.pathPoints.add();
        var endPoint = centerLine.pathPoints.add();

        if (rectHeight <= rectWidth) {
            /* 横長：中央に水平線 / Horizontal: draw horizontal line at center */
            var centerY = (top + bottom) / 2;
            startPoint.anchor = [left, centerY];
            endPoint.anchor = [right, centerY];
        } else {
            /* 縦長：中央に垂直線 / Vertical: draw vertical line at center */
            var centerX = (left + right) / 2;
            startPoint.anchor = [centerX, top];
            endPoint.anchor = [centerX, bottom];
        }

        centerLine.strokeWidth = (rectHeight <= rectWidth) ? rectHeight : rectWidth;

        startPoint.leftDirection = startPoint.anchor;
        startPoint.rightDirection = startPoint.anchor;
        endPoint.leftDirection = endPoint.anchor;
        endPoint.rightDirection = endPoint.anchor;

        removePageItemSafely(rect);
        return centerLine;
    }

    /* 太さが混在するときの代表値を取得（"custom" 指定時は customWidth を返す）
       Pick a representative stroke width when widths differ ("custom" returns customWidth) */
    function getRepresentativeStrokeWidth(lines, strategy, customWidth) {
        if (strategy === "custom" && typeof customWidth === "number" && customWidth > 0) {
            return customWidth;
        }
        var initialStrokeWidth = lines[0].strokeWidth;
        if (strategy === "min") {
            var minimumStrokeWidth = initialStrokeWidth;
            for (var lineIndex = 1; lineIndex < lines.length; lineIndex++) {
                if (lines[lineIndex].strokeWidth < minimumStrokeWidth) minimumStrokeWidth = lines[lineIndex].strokeWidth;
            }
            return minimumStrokeWidth;
        } else if (strategy === "avg") {
            var sum = 0;
            for (var averageLineIndex = 0; averageLineIndex < lines.length; averageLineIndex++) {
                sum += lines[averageLineIndex].strokeWidth;
            }
            return sum / lines.length;
        }
        /* default: max */
        var maximumStrokeWidth = initialStrokeWidth;
        for (var maximumLineIndex = 1; maximumLineIndex < lines.length; maximumLineIndex++) {
            if (lines[maximumLineIndex].strokeWidth > maximumStrokeWidth) maximumStrokeWidth = lines[maximumLineIndex].strokeWidth;
        }
        return maximumStrokeWidth;
    }

    /* 生成結果の線幅を共通化（コンパウンド／グループ内のサブパスも対象）
       Make stroke width common across generated results (recurses into compounds and groups) */

    function applyCommonStrokeWidth(items, strategy, customWidth) {
        if (!items || items.length === 0) return;

        var strokeItems = [];
        function collectStrokedRecursive(item) {
            try {
                if (item.typename === 'PathItem') {
                    if (item.stroked) strokeItems.push(item);
                } else if (item.typename === 'CompoundPathItem') {
                    for (var compoundPathIndex = 0; compoundPathIndex < item.pathItems.length; compoundPathIndex++) {
                        collectStrokedRecursive(item.pathItems[compoundPathIndex]);
                    }
                } else if (item.typename === 'GroupItem') {
                    for (var groupPageItemIndex = 0; groupPageItemIndex < item.pageItems.length; groupPageItemIndex++) {
                        collectStrokedRecursive(item.pageItems[groupPageItemIndex]);
                    }
                }
            } catch (e) { }
        }
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            if (items[itemIndex]) collectStrokedRecursive(items[itemIndex]);
        }
        if (strokeItems.length === 0) return;

        var commonWidth = getRepresentativeStrokeWidth(strokeItems, strategy, customWidth);
        for (var strokeItemIndex = 0; strokeItemIndex < strokeItems.length; strokeItemIndex++) {
            try { strokeItems[strokeItemIndex].strokeWidth = commonWidth; } catch (e) { }
        }
    }

    /* 生成結果の線色を印刷用の黒（C0 M0 Y0 K100）に統一
       Apply print black (C0 M0 Y0 K100) to generated result strokes */
    function applyPrintBlackStroke(items) {
        if (!items || items.length === 0) return;

        var printBlackColor = createPrintBlackColor();

        function applyBlackRecursive(item) {
            try {
                if (item.typename === 'PathItem') {
                    item.stroked = true;
                    item.strokeColor = printBlackColor;
                } else if (item.typename === 'CompoundPathItem') {
                    for (var compoundPathIndex = 0; compoundPathIndex < item.pathItems.length; compoundPathIndex++) {
                        applyBlackRecursive(item.pathItems[compoundPathIndex]);
                    }
                } else if (item.typename === 'GroupItem') {
                    for (var groupItemIndex = 0; groupItemIndex < item.pageItems.length; groupItemIndex++) {
                        applyBlackRecursive(item.pageItems[groupItemIndex]);
                    }
                }
            } catch (e) { }
        }

        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            if (items[itemIndex]) applyBlackRecursive(items[itemIndex]);
        }
    }

    /* 4本の中心線が「#」状に交差していれば、交点を頂点とする長方形に変換
       If 4 center lines form a "#" shape, convert into a rectangle whose corners are the intersections */
    function tryConnectFourLinesIntoRect(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length !== 4) return null;

        var EPS = 0.01;
        var horizontalLineInfos = [];
        var verticalLineInfos = [];

        for (var lineIndex = 0; lineIndex < 4; lineIndex++) {
            var lineItem = lines[lineIndex];
            if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) return null;
            var startAnchor = lineItem.pathPoints[0].anchor;
            var endAnchor = lineItem.pathPoints[1].anchor;
            if (Math.abs(startAnchor[1] - endAnchor[1]) < EPS) {
                horizontalLineInfos.push({
                    y: (startAnchor[1] + endAnchor[1]) / 2,
                    xMin: Math.min(startAnchor[0], endAnchor[0]),
                    xMax: Math.max(startAnchor[0], endAnchor[0]),
                    line: lineItem
                });
            } else if (Math.abs(startAnchor[0] - endAnchor[0]) < EPS) {
                verticalLineInfos.push({
                    x: (startAnchor[0] + endAnchor[0]) / 2,
                    yMin: Math.min(startAnchor[1], endAnchor[1]),
                    yMax: Math.max(startAnchor[1], endAnchor[1]),
                    line: lineItem
                });
            } else {
                return null;
            }
        }

        if (horizontalLineInfos.length !== 2 || verticalLineInfos.length !== 2) return null;

        /* 各 H が両 V の x を内包し、各 V が両 H の y を内包しているか
           Each H must span across both Vs in x; each V across both Hs in y */
        var firstVerticalX = verticalLineInfos[0].x;
        var secondVerticalX = verticalLineInfos[1].x;
        var firstHorizontalY = horizontalLineInfos[0].y;
        var secondHorizontalY = horizontalLineInfos[1].y;
        var xMin = Math.min(firstVerticalX, secondVerticalX);
        var xMax = Math.max(firstVerticalX, secondVerticalX);
        var yMin = Math.min(firstHorizontalY, secondHorizontalY);
        var yMax = Math.max(firstHorizontalY, secondHorizontalY);

        for (var horizontalIndex = 0; horizontalIndex < 2; horizontalIndex++) {
            if (horizontalLineInfos[horizontalIndex].xMin > xMin + EPS || horizontalLineInfos[horizontalIndex].xMax < xMax - EPS) return null;
        }
        for (var verticalIndex = 0; verticalIndex < 2; verticalIndex++) {
            if (verticalLineInfos[verticalIndex].yMin > yMin + EPS || verticalLineInfos[verticalIndex].yMax < yMax - EPS) return null;
        }

        /* 交点を頂点とする閉じた長方形を生成 / Create closed rectangle through intersections */
        var rectanglePath = createGeneratedPathItem();
        rectanglePath.closed = true;
        rectanglePath.stroked = lines[0].stroked;
        rectanglePath.strokeColor = lines[0].strokeColor;
        rectanglePath.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
        rectanglePath.filled = false;

        var corners = [
            [xMin, yMax],
            [xMax, yMax],
            [xMax, yMin],
            [xMin, yMin]
        ];
        for (var cornerIndex = 0; cornerIndex < 4; cornerIndex++) {
            var cornerPoint = rectanglePath.pathPoints.add();
            cornerPoint.anchor = corners[cornerIndex];
            cornerPoint.leftDirection = corners[cornerIndex];
            cornerPoint.rightDirection = corners[cornerIndex];
        }

        for (var removeLineIndex = 0; removeLineIndex < lines.length; removeLineIndex++) {
            removePageItemSafely(lines[removeLineIndex]);
        }

        return rectanglePath;
    }

    /* 3本（2H+1V または 1H+2V）の中心線をコの字型の開いたパスに変換
       Convert 3 lines (2H+1V or 1H+2V) into an open U-shape path */
    function tryConnectThreeLinesIntoUShape(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length !== 3) return null;

        var EPS = 0.01;
        var horizontalLineInfos = [], verticalLineInfos = [];

        for (var lineIndex = 0; lineIndex < 3; lineIndex++) {
            var lineItem = lines[lineIndex];
            if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) return null;
            var startAnchor = lineItem.pathPoints[0].anchor;
            var endAnchor = lineItem.pathPoints[1].anchor;
            if (Math.abs(startAnchor[1] - endAnchor[1]) < EPS) {
                horizontalLineInfos.push({ y: (startAnchor[1] + endAnchor[1]) / 2, xMin: Math.min(startAnchor[0], endAnchor[0]), xMax: Math.max(startAnchor[0], endAnchor[0]) });
            } else if (Math.abs(startAnchor[0] - endAnchor[0]) < EPS) {
                verticalLineInfos.push({ x: (startAnchor[0] + endAnchor[0]) / 2, yMin: Math.min(startAnchor[1], endAnchor[1]), yMax: Math.max(startAnchor[1], endAnchor[1]) });
            } else {
                return null;
            }
        }

        var anchors = null;

        if (horizontalLineInfos.length === 2 && verticalLineInfos.length === 1) {
            /* 2H + 1V：V は両 H に交差し、両 H の x 範囲の片端付近にある
               V crosses both Hs and sits near one extreme of the H x-range */
            var verticalLineInfo = verticalLineInfos[0];
            var firstHorizontalY = horizontalLineInfos[0].y;
            var secondHorizontalY = horizontalLineInfos[1].y;
            var yMin = Math.min(firstHorizontalY, secondHorizontalY);
            var yMax = Math.max(firstHorizontalY, secondHorizontalY);
            if (verticalLineInfo.yMin > yMin + EPS || verticalLineInfo.yMax < yMax - EPS) return null;
            for (var horizontalIndex = 0; horizontalIndex < 2; horizontalIndex++) {
                if (horizontalLineInfos[horizontalIndex].xMin > verticalLineInfo.x + EPS || horizontalLineInfos[horizontalIndex].xMax < verticalLineInfo.x - EPS) return null;
            }
            var horizontalXMin = Math.min(horizontalLineInfos[0].xMin, horizontalLineInfos[1].xMin);
            var horizontalXMax = Math.max(horizontalLineInfos[0].xMax, horizontalLineInfos[1].xMax);
            var openEndX = (Math.abs(verticalLineInfo.x - horizontalXMin) < Math.abs(verticalLineInfo.x - horizontalXMax)) ? horizontalXMax : horizontalXMin;
            anchors = [
                [openEndX, firstHorizontalY],
                [verticalLineInfo.x, firstHorizontalY],
                [verticalLineInfo.x, secondHorizontalY],
                [openEndX, secondHorizontalY]
            ];
        } else if (horizontalLineInfos.length === 1 && verticalLineInfos.length === 2) {
            /* 1H + 2V：H は両 V に交差し、両 V の y 範囲の片端付近にある
               H crosses both Vs and sits near one extreme of the V y-range */
            var horizontalLineInfo = horizontalLineInfos[0];
            var firstVerticalX = verticalLineInfos[0].x;
            var secondVerticalX = verticalLineInfos[1].x;
            var xMin = Math.min(firstVerticalX, secondVerticalX);
            var xMax = Math.max(firstVerticalX, secondVerticalX);
            if (horizontalLineInfo.xMin > xMin + EPS || horizontalLineInfo.xMax < xMax - EPS) return null;
            for (var verticalIndex = 0; verticalIndex < 2; verticalIndex++) {
                if (verticalLineInfos[verticalIndex].yMin > horizontalLineInfo.y + EPS || verticalLineInfos[verticalIndex].yMax < horizontalLineInfo.y - EPS) return null;
            }
            var verticalYMin = Math.min(verticalLineInfos[0].yMin, verticalLineInfos[1].yMin);
            var verticalYMax = Math.max(verticalLineInfos[0].yMax, verticalLineInfos[1].yMax);
            var openEndY = (Math.abs(horizontalLineInfo.y - verticalYMin) < Math.abs(horizontalLineInfo.y - verticalYMax)) ? verticalYMax : verticalYMin;
            anchors = [
                [firstVerticalX, openEndY],
                [firstVerticalX, horizontalLineInfo.y],
                [secondVerticalX, horizontalLineInfo.y],
                [secondVerticalX, openEndY]
            ];
        } else {
            return null;
        }

        var path = createGeneratedPathItem();
        path.closed = false;
        path.stroked = lines[0].stroked;
        path.strokeColor = lines[0].strokeColor;
        path.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
        path.filled = false;

        for (var anchorIndex = 0; anchorIndex < anchors.length; anchorIndex++) {
            var pathPoint = path.pathPoints.add();
            pathPoint.anchor = anchors[anchorIndex];
            pathPoint.leftDirection = anchors[anchorIndex];
            pathPoint.rightDirection = anchors[anchorIndex];
        }

        for (var removeIndex = 0; removeIndex < lines.length; removeIndex++) {
            removePageItemSafely(lines[removeIndex]);
        }

        return path;
    }

    /* 2本（1H+1V）の中心線をL字型の開いたパスに変換
       Convert 2 lines (1H+1V) into an open L-shape path */
    function tryConnectTwoLinesIntoLShape(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length !== 2) return null;

        var EPS = 0.01;
        var horizontalLineInfos = [], verticalLineInfos = [];

        for (var lineIndex = 0; lineIndex < 2; lineIndex++) {
            var lineItem = lines[lineIndex];
            if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) return null;
            var startAnchor = lineItem.pathPoints[0].anchor;
            var endAnchor = lineItem.pathPoints[1].anchor;
            if (Math.abs(startAnchor[1] - endAnchor[1]) < EPS) {
                horizontalLineInfos.push({ y: (startAnchor[1] + endAnchor[1]) / 2, xMin: Math.min(startAnchor[0], endAnchor[0]), xMax: Math.max(startAnchor[0], endAnchor[0]) });
            } else if (Math.abs(startAnchor[0] - endAnchor[0]) < EPS) {
                verticalLineInfos.push({ x: (startAnchor[0] + endAnchor[0]) / 2, yMin: Math.min(startAnchor[1], endAnchor[1]), yMax: Math.max(startAnchor[1], endAnchor[1]) });
            } else {
                return null;
            }
        }

        if (horizontalLineInfos.length !== 1 || verticalLineInfos.length !== 1) return null;

        var horizontalLineInfo = horizontalLineInfos[0];
        var verticalLineInfo = verticalLineInfos[0];

        /* 交差していること / Must intersect */
        if (horizontalLineInfo.xMin > verticalLineInfo.x + EPS || horizontalLineInfo.xMax < verticalLineInfo.x - EPS) return null;
        if (verticalLineInfo.yMin > horizontalLineInfo.y + EPS || verticalLineInfo.yMax < horizontalLineInfo.y - EPS) return null;

        var farEndY = (Math.abs(horizontalLineInfo.y - verticalLineInfo.yMin) < Math.abs(horizontalLineInfo.y - verticalLineInfo.yMax)) ? verticalLineInfo.yMax : verticalLineInfo.yMin;
        var farEndX = (Math.abs(verticalLineInfo.x - horizontalLineInfo.xMin) < Math.abs(verticalLineInfo.x - horizontalLineInfo.xMax)) ? horizontalLineInfo.xMax : horizontalLineInfo.xMin;

        var anchors = [
            [verticalLineInfo.x, farEndY],
            [verticalLineInfo.x, horizontalLineInfo.y],
            [farEndX, horizontalLineInfo.y]
        ];

        var path = createGeneratedPathItem();
        path.closed = false;
        path.stroked = lines[0].stroked;
        path.strokeColor = lines[0].strokeColor;
        path.strokeWidth = getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
        path.filled = false;

        for (var anchorIndex = 0; anchorIndex < anchors.length; anchorIndex++) {
            var pathPoint = path.pathPoints.add();
            pathPoint.anchor = anchors[anchorIndex];
            pathPoint.leftDirection = anchors[anchorIndex];
            pathPoint.rightDirection = anchors[anchorIndex];
        }

        for (var removeIndex = 0; removeIndex < lines.length; removeIndex++) {
            removePageItemSafely(lines[removeIndex]);
        }

        return path;
    }

    /* Pathfinder Divide の Live Effect を XML で適用するヘルパー
       Helper to apply a Pathfinder Divide Live Effect via XML */
    function applyPathfinderDivideEffect(item, removeUnpainted, expandAppearance) {
        var shouldRemoveUnpainted = (removeUnpainted !== false);
        var shouldExpandAppearance = (expandAppearance === true);
        var values = [
            5,                         /* Command: Divide */
            1,                         /* ConvertCustom */
            shouldRemoveUnpainted ? 1 : 0,
            0.5,                       /* Mix */
            10,                        /* Precision */
            1,                         /* RemovePoints */
            'Divide'
        ];
        var xml = ('<LiveEffect name="Adobe Pathfinder" isPre="1"><Dict data="I Command #1 B ConvertCustom #2 B ExtractUnpainted #3 R Mix #4 R Precision #5 B RemovePoints #6"><Entry name="DisplayString" value="#7" valueType="S"/></Dict></LiveEffect>')
            .replace(/#(\d+)/g, function (_, n) { return values[parseInt(n, 10) - 1]; });

        item.applyEffect(xml);
        if (shouldExpandAppearance) app.executeMenuCommand("expandStyle");
    }

    /* PageItem 内のすべての PathItem に再帰的に処理を適用
       Recursively apply a callback to every PathItem inside a PageItem */
    function walkPathItemsRecursive(item, callback) {
        try {
            if (item.typename === 'PathItem') {
                callback(item);
            } else if (item.typename === 'CompoundPathItem') {
                for (var compoundPathIndex = 0; compoundPathIndex < item.pathItems.length; compoundPathIndex++) {
                    callback(item.pathItems[compoundPathIndex]);
                }
            } else if (item.typename === 'GroupItem') {
                for (var groupItemIndex = 0; groupItemIndex < item.pageItems.length; groupItemIndex++) {
                    walkPathItemsRecursive(item.pageItems[groupItemIndex], callback);
                }
            }
        } catch (e) { }
    }

    /* 現在の選択内のすべてのパスを塗りなし・線ありに整える
       Restore no-fill / stroked appearance for all paths in the current selection */
    function restoreNoFillStrokeToCurrentSelection(strokeColor, strokeWidth) {
        for (var selectionIndex = 0; selectionIndex < app.selection.length; selectionIndex++) {
            walkPathItemsRecursive(app.selection[selectionIndex], function (pathItem) {
                pathItem.filled = false;
                pathItem.stroked = true;
                if (strokeColor) {
                    try { pathItem.strokeColor = strokeColor; } catch (e) { }
                }
                if (strokeWidth !== null && strokeWidth !== undefined) {
                    pathItem.strokeWidth = strokeWidth;
                }
            });
        }
    }

    /* 選択中のアピアランスをアウトライン統合し、塗りなし・線ありに戻す
       Union the selected appearance into outlines, then restore no-fill / stroked paths */
    function outlineUnionCurrentSelection(strokeColor, strokeWidth) {
        app.executeMenuCommand('Adobe New Fill Shortcut');
        app.executeMenuCommand('expandStyle');
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');

        restoreNoFillStrokeToCurrentSelection(strokeColor, strokeWidth);
    }

    /* 5本以上の中心線をグループ化し、Pathfinder Divide で分割後、
       新規塗り・展開・Pathfinder Add・再展開を経てアウトラインに統合する。
       展開後の選択結果に対して、塗りなし・線ありを再設定する。
       Group 5+ center lines, split them with Pathfinder Divide,
       then unify the outline through new fill, expand, Pathfinder Add, and expand again.
       Restore no-fill / stroked appearance on the expanded selection result. */
    function tryConnectManyLinesIntoOutline(lines, strokeStrategy, commonStrokeWidth, customWidth) {
        if (!lines || lines.length < 5) return null;

        var doc = app.activeDocument;
        var representativeStrokeWidth = (commonStrokeWidth !== null && commonStrokeWidth !== undefined)
            ? commonStrokeWidth
            : getRepresentativeStrokeWidth(lines, strokeStrategy, customWidth);
        var representativeStrokeColor = lines[0].strokeColor;

        /* 生成専用レイヤー上でグループ化し、ロックされた親レイヤー／activeLayer に依存しない
           Group on the dedicated generated layer to avoid relying on locked parent layers or activeLayer */
        var groupParent = ensureGeneratedWorkLayer();

        /* 1. 線をすべて1つのグループに集める / Move all lines into a single group */
        var group = createGeneratedGroup();
        for (var lineIndex = lines.length - 1; lineIndex >= 0; lineIndex--) {
            movePageItemSafely(lines[lineIndex], group, ElementPlacement.PLACEATBEGINNING);
        }

        /* 2. Live Effect の Pathfinder Divide を適用（交差点で分割）
           Apply Pathfinder Divide as a Live Effect to split at intersections */
        applyPathfinderDivideEffect(group, false, false);

        /* 3. グループを選択状態にする / Make the group the current selection */
        app.selection = null;
        group.selected = true;
        app.redraw();

        /* 4. アウトライン統合し、現在の選択結果に対して塗りなし・線ありを再設定
           Union outlines, then clear fills and restore strokes on the current selection result */
        outlineUnionCurrentSelection(representativeStrokeColor, representativeStrokeWidth);

        /* 5. グループ内が単一要素なら親に昇格、複数要素なら包むグループのまま返す
           If the group contains a single child, promote it; otherwise return the group */
        var result = group;
        try {
            if (group.pageItems.length === 1) {
                result = group.pageItems[0];
                movePageItemSafely(result, groupParent, ElementPlacement.PLACEATEND);
                removePageItemSafely(group);
            }
        } catch (e) { }

        return result;
    }

    /* クラスタを格子状に整列：水平線は X レンジ全体、垂直線は Y レンジ全体に伸ばし、線端は突出端で揃える
       Align a cluster into a grid: extend H lines to full X range, V lines to full Y range, with projecting end caps */
    function tryAlignClusterAsGrid(lines) {
        if (!lines || lines.length < 4) return null;

        var horizontalInfos = [];
        var verticalInfos = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var lineItem = lines[lineIndex];
            if (!lineItem.pathPoints || lineItem.pathPoints.length !== 2) continue;
            var startAnchor = lineItem.pathPoints[0].anchor;
            var endAnchor = lineItem.pathPoints[1].anchor;
            var deltaX = Math.abs(startAnchor[0] - endAnchor[0]);
            var deltaY = Math.abs(startAnchor[1] - endAnchor[1]);
            if (deltaX >= deltaY) {
                horizontalInfos.push({ path: lineItem, y: (startAnchor[1] + endAnchor[1]) / 2 });
            } else {
                verticalInfos.push({ path: lineItem, x: (startAnchor[0] + endAnchor[0]) / 2 });
            }
        }
        if (horizontalInfos.length < 2 || verticalInfos.length < 2) return null;

        var minX = verticalInfos[0].x, maxX = verticalInfos[0].x;
        for (var verticalInfoIndex = 1; verticalInfoIndex < verticalInfos.length; verticalInfoIndex++) {
            if (verticalInfos[verticalInfoIndex].x < minX) minX = verticalInfos[verticalInfoIndex].x;
            if (verticalInfos[verticalInfoIndex].x > maxX) maxX = verticalInfos[verticalInfoIndex].x;
        }
        var minY = horizontalInfos[0].y, maxY = horizontalInfos[0].y;
        for (var horizontalInfoIndex = 1; horizontalInfoIndex < horizontalInfos.length; horizontalInfoIndex++) {
            if (horizontalInfos[horizontalInfoIndex].y < minY) minY = horizontalInfos[horizontalInfoIndex].y;
            if (horizontalInfos[horizontalInfoIndex].y > maxY) maxY = horizontalInfos[horizontalInfoIndex].y;
        }

        function setLineEndpoints(pathItem, anchorA, anchorB) {
            pathItem.pathPoints[0].anchor = anchorA;
            pathItem.pathPoints[0].leftDirection = anchorA;
            pathItem.pathPoints[0].rightDirection = anchorA;
            pathItem.pathPoints[1].anchor = anchorB;
            pathItem.pathPoints[1].leftDirection = anchorB;
            pathItem.pathPoints[1].rightDirection = anchorB;
        }

        for (var horizontalLineIndex = 0; horizontalLineIndex < horizontalInfos.length; horizontalLineIndex++) {
            var horizontalPath = horizontalInfos[horizontalLineIndex].path;
            setStrokeCapSafely(horizontalPath, StrokeCap.PROJECTINGENDCAP);
            setLineEndpoints(horizontalPath, [minX, horizontalInfos[horizontalLineIndex].y], [maxX, horizontalInfos[horizontalLineIndex].y]);
        }
        /* Illustrator は Y 上方向が正 / Illustrator Y axis points up */
        var topY = Math.max(minY, maxY);
        var bottomY = Math.min(minY, maxY);
        for (var verticalLineIndex = 0; verticalLineIndex < verticalInfos.length; verticalLineIndex++) {
            var verticalPath = verticalInfos[verticalLineIndex].path;
            setStrokeCapSafely(verticalPath, StrokeCap.PROJECTINGENDCAP);
            setLineEndpoints(verticalPath, [verticalInfos[verticalLineIndex].x, topY], [verticalInfos[verticalLineIndex].x, bottomY]);
        }

        return lines;
    }

    /* 整列済みクラスタの最外周4本（top/bottom H と left/right V）を長方形に置換し、内側の線はそのまま残す
       Replace the outermost 4 lines (top/bottom H, left/right V) of an aligned cluster with a rectangle, keeping the inner lines */
    function convertOuterFrameToRect(lines, strokeStrategy, customWidth) {
        if (!lines || lines.length < 4) return null;

        var horizontalInfos = [];
        var verticalInfos = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var lineItem = lines[lineIndex];
            if (!lineItem || !lineItem.pathPoints || lineItem.pathPoints.length !== 2) continue;
            var startAnchor = lineItem.pathPoints[0].anchor;
            var endAnchor = lineItem.pathPoints[1].anchor;
            var deltaX = Math.abs(startAnchor[0] - endAnchor[0]);
            var deltaY = Math.abs(startAnchor[1] - endAnchor[1]);
            if (deltaX >= deltaY) {
                horizontalInfos.push({ path: lineItem, y: (startAnchor[1] + endAnchor[1]) / 2 });
            } else {
                verticalInfos.push({ path: lineItem, x: (startAnchor[0] + endAnchor[0]) / 2 });
            }
        }
        if (horizontalInfos.length < 2 || verticalInfos.length < 2) return null;

        var topHorizontalInfo = horizontalInfos[0];
        var bottomHorizontalInfo = horizontalInfos[0];
        for (var horizontalInfoIndex = 1; horizontalInfoIndex < horizontalInfos.length; horizontalInfoIndex++) {
            if (horizontalInfos[horizontalInfoIndex].y > topHorizontalInfo.y) topHorizontalInfo = horizontalInfos[horizontalInfoIndex];
            if (horizontalInfos[horizontalInfoIndex].y < bottomHorizontalInfo.y) bottomHorizontalInfo = horizontalInfos[horizontalInfoIndex];
        }
        var leftVerticalInfo = verticalInfos[0];
        var rightVerticalInfo = verticalInfos[0];
        for (var verticalInfoIndex = 1; verticalInfoIndex < verticalInfos.length; verticalInfoIndex++) {
            if (verticalInfos[verticalInfoIndex].x < leftVerticalInfo.x) leftVerticalInfo = verticalInfos[verticalInfoIndex];
            if (verticalInfos[verticalInfoIndex].x > rightVerticalInfo.x) rightVerticalInfo = verticalInfos[verticalInfoIndex];
        }

        var outerLines = [topHorizontalInfo.path, bottomHorizontalInfo.path, leftVerticalInfo.path, rightVerticalInfo.path];

        var minX = leftVerticalInfo.x;
        var maxX = rightVerticalInfo.x;
        var minY = bottomHorizontalInfo.y;
        var maxY = topHorizontalInfo.y;

        var rectanglePath = createGeneratedPathItem();
        rectanglePath.closed = true;
        rectanglePath.stroked = topHorizontalInfo.path.stroked;
        rectanglePath.strokeColor = topHorizontalInfo.path.strokeColor;
        rectanglePath.strokeWidth = getRepresentativeStrokeWidth(outerLines, strokeStrategy, customWidth);
        rectanglePath.filled = false;

        var corners = [
            [minX, maxY],
            [maxX, maxY],
            [maxX, minY],
            [minX, minY]
        ];
        for (var cornerIndex = 0; cornerIndex < 4; cornerIndex++) {
            var cornerPoint = rectanglePath.pathPoints.add();
            cornerPoint.anchor = corners[cornerIndex];
            cornerPoint.leftDirection = corners[cornerIndex];
            cornerPoint.rightDirection = corners[cornerIndex];
        }

        /* 内側線のみ残す（参照比較で外周4本を除外） / Keep only inner lines (exclude outer 4 by reference) */
        var remainingLines = [rectanglePath];
        for (var remainingLineIndex = 0; remainingLineIndex < lines.length; remainingLineIndex++) {
            var isOuterLine = false;
            for (var outerLineIndex = 0; outerLineIndex < outerLines.length; outerLineIndex++) {
                if (lines[remainingLineIndex] === outerLines[outerLineIndex]) { isOuterLine = true; break; }
            }
            if (!isOuterLine) remainingLines.push(lines[remainingLineIndex]);
        }

        for (var removeOuterLineIndex = 0; removeOuterLineIndex < outerLines.length; removeOuterLineIndex++) {
            removePageItemSafely(outerLines[removeOuterLineIndex]);
        }

        return remainingLines;
    }

    /* 軸並行な2線分が交差するか判定（端点接触も許容、tolerance で「少し届かない線」も交差扱い）
       Test whether 2 axis-aligned segments intersect (endpoint touch allowed; tolerance lets near-misses count as intersecting) */
    function intersectsAxisAligned(lnA, lnB, tolerance) {
        var EPS = (typeof tolerance === "number" && tolerance > 0) ? tolerance : 0.01;
        var firstLineStartAnchor = lnA.pathPoints[0].anchor;
        var firstLineEndAnchor = lnA.pathPoints[1].anchor;
        var secondLineStartAnchor = lnB.pathPoints[0].anchor;
        var secondLineEndAnchor = lnB.pathPoints[1].anchor;
        var isFirstLineHorizontal = Math.abs(firstLineStartAnchor[1] - firstLineEndAnchor[1]) < EPS;
        var isSecondLineHorizontal = Math.abs(secondLineStartAnchor[1] - secondLineEndAnchor[1]) < EPS;
        if (isFirstLineHorizontal === isSecondLineHorizontal) return false; /* 平行は非交差 / Parallel: not intersecting */
        var horizontalStartAnchor = isFirstLineHorizontal ? firstLineStartAnchor : secondLineStartAnchor;
        var horizontalEndAnchor = isFirstLineHorizontal ? firstLineEndAnchor : secondLineEndAnchor;
        var verticalStartAnchor = isFirstLineHorizontal ? secondLineStartAnchor : firstLineStartAnchor;
        var verticalEndAnchor = isFirstLineHorizontal ? secondLineEndAnchor : firstLineEndAnchor;
        var horizontalY = (horizontalStartAnchor[1] + horizontalEndAnchor[1]) / 2;
        var verticalX = (verticalStartAnchor[0] + verticalEndAnchor[0]) / 2;
        var horizontalXMin = Math.min(horizontalStartAnchor[0], horizontalEndAnchor[0]);
        var horizontalXMax = Math.max(horizontalStartAnchor[0], horizontalEndAnchor[0]);
        var verticalYMin = Math.min(verticalStartAnchor[1], verticalEndAnchor[1]);
        var verticalYMax = Math.max(verticalStartAnchor[1], verticalEndAnchor[1]);
        return (verticalX >= horizontalXMin - EPS && verticalX <= horizontalXMax + EPS &&
            horizontalY >= verticalYMin - EPS && horizontalY <= verticalYMax + EPS);
    }

    /* 中心線を交差連結成分でクラスタ分け（tolerance で許容幅を指定可）
       Cluster lines by intersection-connected components (optional tolerance for near-miss tolerance) */
    function clusterLinesByIntersection(lines, tolerance) {
        var lineCount = lines.length;
        var visited = [];
        for (var visitedIndex = 0; visitedIndex < lineCount; visitedIndex++) visited.push(false);
        var clusters = [];
        for (var startIndex = 0; startIndex < lineCount; startIndex++) {
            if (visited[startIndex]) continue;
            var cluster = [];
            var stack = [startIndex];
            visited[startIndex] = true;
            while (stack.length > 0) {
                var currentLineIndex = stack.pop();
                cluster.push(lines[currentLineIndex]);
                for (var compareLineIndex = 0; compareLineIndex < lineCount; compareLineIndex++) {
                    if (!visited[compareLineIndex] && intersectsAxisAligned(lines[currentLineIndex], lines[compareLineIndex], tolerance)) {
                        visited[compareLineIndex] = true;
                        stack.push(compareLineIndex);
                    }
                }
            }
            clusters.push(cluster);
        }
        return clusters;
    }

    /* 格子モード用のクラスタ許容幅：最大ストローク幅の3倍（最低5pt）。
       thin な線も拾えるよう floor を持たせる
       Grid-mode cluster tolerance: max stroke width × 3, floored at 5pt so thin lines still get coverage */
    function computeGridClusterTolerance(lines) {
        if (!lines || lines.length === 0) return 0;
        var maxStrokeWidth = 0;
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            try {
                if (lines[lineIndex].strokeWidth > maxStrokeWidth) maxStrokeWidth = lines[lineIndex].strokeWidth;
            } catch (e) { }
        }
        var tolerance = maxStrokeWidth * 3;
        if (tolerance < 5) tolerance = 5;
        return tolerance;
    }

    // =========================================
    // 実行フロー補助 / Execution Flow Helpers
    // =========================================

    /* 選択内容から処理対象の中心線を生成または収集
       Create or collect target center lines from the current selection */
    function buildTargetCenterLines(selectedItems, options) {
        var generatedCenterLines = [];

        if (options.enableCenterLine) {
            generatedCenterLines = convertSelectedRectanglesToCenterLines(selectedItems, options);
        } else {
            collectLinesRecursive(selectedItems, generatedCenterLines);
        }

        return generatedCenterLines;
    }

    /* 選択中の長方形を中心線化して返す
       Convert selected rectangles to center lines and return them */
    function convertSelectedRectanglesToCenterLines(selectedItems, options) {
        var generatedCenterLines = [];
        var rectangleItemsToProcess = [];
        collectRectsRecursive(selectedItems, rectangleItemsToProcess);

        for (var rectangleIndex = 0; rectangleIndex < rectangleItemsToProcess.length; rectangleIndex++) {
            var rectangleItem = rectangleItemsToProcess[rectangleIndex];
            var newLine = convertRectToCenterLine(rectangleItem, options.correctRotation, options.minShortSidePt);

            if (newLine) {
                // 生成レイヤーに統一（元レイヤーには戻さない）
                generatedCenterLines.push(newLine);
            }
        }

        return generatedCenterLines;
    }

    /* 生成済み中心線を連結モードに応じてクラスタ単位で処理
       Process generated center lines cluster-by-cluster according to connect mode */
    function processCenterLineClusters(generatedCenterLines, options) {
        if (!generatedCenterLines || generatedCenterLines.length === 0) return generatedCenterLines;

        var commonStrokeWidth = options.commonStroke
            ? getRepresentativeStrokeWidth(generatedCenterLines, options.strokeStrategy, options.customStrokeWidth)
            : null;
        var clusterTolerance = (options.connectMode === "grid")
            ? computeGridClusterTolerance(generatedCenterLines)
            : 0;
        var clusters = clusterLinesByIntersection(generatedCenterLines, clusterTolerance);
        var resultLines = [];

        for (var clusterIndex = 0; clusterIndex < clusters.length; clusterIndex++) {
            appendProcessedCluster(resultLines, clusters[clusterIndex], options, commonStrokeWidth);
        }

        return resultLines;
    }

    /* 1つのクラスタを処理し、結果配列へ追加
       Process one cluster and append its result to the output array */
    function appendProcessedCluster(resultLines, cluster, options, commonStrokeWidth) {
        var combined = null;
        var gridAligned = null;

        if (options.connectMode === "connect") {
            combined = connectClusterByLineCount(cluster, options, commonStrokeWidth);
        } else if (options.connectMode === "grid") {
            gridAligned = alignClusterAsGridResult(cluster, options);
        }

        if (combined) {
            resultLines.push(combined);
        } else if (gridAligned) {
            for (var gridLineIndex = 0; gridLineIndex < gridAligned.length; gridLineIndex++) {
                resultLines.push(gridAligned[gridLineIndex]);
            }
        } else {
            for (var clusterLineIndex = 0; clusterLineIndex < cluster.length; clusterLineIndex++) {
                resultLines.push(cluster[clusterLineIndex]);
            }
        }
    }

    /* 本数に応じてクラスタを結合
       Connect a cluster according to its line count */
    function connectClusterByLineCount(cluster, options, commonStrokeWidth) {
        if (cluster.length === 4) {
            return tryConnectFourLinesIntoRect(cluster, options.strokeStrategy, options.customStrokeWidth);
        }
        if (cluster.length === 3) {
            return tryConnectThreeLinesIntoUShape(cluster, options.strokeStrategy, options.customStrokeWidth);
        }
        if (cluster.length === 2) {
            return tryConnectTwoLinesIntoLShape(cluster, options.strokeStrategy, options.customStrokeWidth);
        }
        if (cluster.length >= 5) {
            return tryConnectManyLinesIntoOutline(cluster, options.strokeStrategy, commonStrokeWidth, options.customStrokeWidth);
        }
        return null;
    }

    /* クラスタを格子化し、必要に応じて外枠を長方形化
       Align a cluster as a grid and optionally convert its outer frame to a rectangle */
    function alignClusterAsGridResult(cluster, options) {
        var gridAligned = tryAlignClusterAsGrid(cluster);
        if (gridAligned && options.outerRect) {
            var withFrame = convertOuterFrameToRect(gridAligned, options.strokeStrategy, options.customStrokeWidth);
            if (withFrame) {
                gridAligned = withFrame;

                // 外枠長方形化後に「Convert to Shape」を適用
                try {
                    app.activeDocument.selection = gridAligned;
                    app.executeMenuCommand('Convert to Shape');
                } catch (e) { }
            }
        }
        return gridAligned;
    }

    /* 線幅・色などの仕上げ処理を適用
       Apply finishing options such as stroke width and print black */
    function applyFinishingOptions(generatedCenterLines, options) {
        if (options.commonStroke || options.strokeStrategy === "custom") {
            applyCommonStrokeWidth(generatedCenterLines, options.strokeStrategy, options.customStrokeWidth);
        }
        if (options.printBlack) {
            applyPrintBlackStroke(generatedCenterLines);
        }
    }

    /* 結果を選択状態にし、必要に応じて1つのグループにまとめる
       Select results and optionally wrap them into one group (and move all to generated layer) */
    function selectOrGroupResults(generatedCenterLines, options) {
        if (!generatedCenterLines || generatedCenterLines.length === 0) return;

        // すでに生成レイヤーに集約されている前提で、そのまま扱う
        if (options.groupResult && options.connectMode === "grid") {
            var resultGroup = groupGeneratedCenterLines(generatedCenterLines);
            app.activeDocument.selection = [resultGroup];
        } else {
            app.activeDocument.selection = generatedCenterLines;
        }
    }

    /* 生成結果を1つのグループへまとめる
       Wrap generated results into one group */
    function groupGeneratedCenterLines(generatedCenterLines) {
        var groupParent = getGeneratedResultGroupParent(generatedCenterLines);
        var resultGroup = createGeneratedGroup();

        for (var groupItemIndex = generatedCenterLines.length - 1; groupItemIndex >= 0; groupItemIndex--) {
            movePageItemSafely(generatedCenterLines[groupItemIndex], resultGroup, ElementPlacement.PLACEATBEGINNING);
        }

        return resultGroup;
    }

    /* 生成結果をまとめる親レイヤー／親グループを決定
       Determine parent layer/group for wrapping generated results */
    function getGeneratedResultGroupParent(generatedCenterLines) {
        return ensureGeneratedWorkLayer();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
        var savedActiveLayer = null;
        var savedActiveLayerLocked = false;
        var savedActiveLayerHidden = false;
        var savedTopLayer = null;
        var savedTopLayerLocked = false;
        var savedTopLayerHidden = false;
        try {
            if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
                alert(getLabel('alertNoSelection'));
                return;
            }

            /* activeLayer または最前面レイヤーがロック／非表示だと pathItems.add() 等で
               "Target layer cannot be modified" になることがあるため一時解除 */
            try {
                savedActiveLayer = app.activeDocument.activeLayer;
                savedActiveLayerLocked = savedActiveLayer.locked;
                savedActiveLayerHidden = !savedActiveLayer.visible;
                if (savedActiveLayerLocked) savedActiveLayer.locked = false;
                if (savedActiveLayerHidden) savedActiveLayer.visible = true;
            } catch (e) { }

            try {
                savedTopLayer = app.activeDocument.layers[0];
                if (savedTopLayer && savedTopLayer !== savedActiveLayer) {
                    savedTopLayerLocked = savedTopLayer.locked;
                    savedTopLayerHidden = !savedTopLayer.visible;
                    if (savedTopLayerLocked) savedTopLayer.locked = false;
                    if (savedTopLayerHidden) savedTopLayer.visible = true;
                }
            } catch (e) { }

            var originalLayer = app.activeDocument.activeLayer;

            var options = showOptionDialog();
            if (options === null) return;

            var selectedItems = app.activeDocument.selection;
            var generatedCenterLines = buildTargetCenterLines(selectedItems, options);

            if (generatedCenterLines.length > 0) {
                generatedCenterLines = processCenterLineClusters(generatedCenterLines, options);
                applyFinishingOptions(generatedCenterLines, options);
                selectOrGroupResults(generatedCenterLines, options);
                if (options.returnToOriginal && originalLayer) {
                    try {
                        var genLayer = app.activeDocument.layers.getByName(GENERATED_WORK_LAYER_NAME);
                        for (var i = genLayer.pageItems.length - 1; i >= 0; i--) {
                            movePageItemSafely(genLayer.pageItems[i], originalLayer, ElementPlacement.PLACEATEND);
                        }
                        genLayer.remove();
                    } catch (e) { }
                }
            }
        } catch (e) {
            alert(labelText('alertError') + "\n" + e);
        } finally {
            try {
                if (savedTopLayer && savedTopLayer !== savedActiveLayer) {
                    if (savedTopLayerLocked) savedTopLayer.locked = true;
                    if (savedTopLayerHidden) savedTopLayer.visible = false;
                }
            } catch (e) { }

            try {
                if (savedActiveLayer) {
                    if (savedActiveLayerLocked) savedActiveLayer.locked = true;
                    if (savedActiveLayerHidden) savedActiveLayer.visible = false;
                }
            } catch (e) { }
        }
    }

    main();
})();
