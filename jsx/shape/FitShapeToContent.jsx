#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "FitShapeToContentSession"

/*

### 概要

テキストやグループに合わせて、背面の座布団形状を手早く作成・調整します。
図形も一緒に選ぶ（またはテキストと図形のグループを選ぶ）と、その図形を座布団として使います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitShapeToContent.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n6e4a6a2b175f

### Overview

Quickly creates and adjusts a backing shape that fits a text frame or a group.
Add a shape to the selection, or select a text+shape group, to reuse that shape instead of creating a rectangle.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitShapeToContent.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitShapeToContent";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitShapeToContent.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitShapeToContent.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6e4a6a2b175f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 自動作成した座布団の初期不透明度（%）/ Initial opacity of the auto-created backing shape (%) */
    var AUTO_CREATED_SHAPE_OPACITY = 20;

    /* パディングの初期値（pt）。定規の単位に換算して入力欄に入れる / Default padding (pt), shown in the ruler unit */
    var DEFAULT_PADDING_PT = 20;

    /* プレビュー図形の最小サイズ（pt）。0以下は Illustrator が受け付けない / Minimum preview size (pt) */
    var MIN_PREVIEW_SIZE = 0.1;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between columns */
    var FIELD_LABEL_WIDTH  = 50;               /* 項目名の幅 / width of a field label */
    var FIELD_CHARS        = 5;                /* 数値入力欄の文字数 / character width of a numeric field */

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

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

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

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

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
            return textFile.read().replace(/^\uFEFF/, "");
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
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
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

    // =========================================
    // セッション記憶 / Session state
    // =========================================

    /* 設定は Illustrator の終了まで残す（#targetengine が必須）。旧版の $.global のキーを1度だけ読み継ぐ /
       Settings last until Illustrator quits (needs #targetengine); the old $.global key is read once */
    var LEGACY_SESSION_KEY = "__FitShapeToContentSession__";
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_KEY] || null; }
    });

    /**
     * 前回のダイアログ設定を取得する（初回は既定値）
     * @param {number} unitFactor - 表示単位1つあたりの pt 数
     * @returns {object} 保存済みの設定
     */
    function getSessionState(unitFactor) {
        var defaultPadding = formatUnitValue(DEFAULT_PADDING_PT, unitFactor);
        return settingsStore.load({
            addW: defaultPadding,
            addH: defaultPadding,
            radius: "0",
            radiusEnabled: true,
            link: true,
            pill: false
        });
    }

    /**
     * ダイアログ設定を保存する
     * @param {object} state - 保存する設定
     * @returns {void}
     */
    function saveSessionState(state) {
        settingsStore.save(state);
    }

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "座布団メーカー", en: "Fit Shape to Content" }
        },
        panel: {
            padding: { ja: "パディング", en: "Padding" },
            corner:  { ja: "角丸", en: "Rounded Corners" }
        },
        fieldLabel: {
            width:  { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            radius: { ja: "半径", en: "Radius" }
        },
        checkbox: {
            adjustEnabled: { ja: "座布団の調整", en: "Adjust Shape" },
            pill:          { ja: "ピル形状", en: "Pill Shape" }
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
            adjustEnabled: {
                ja: "オフにすると、座布団の大きさと角丸はそのままで位置だけを合わせます。",
                en: "When off, only the position is matched; the shape keeps its current size and corners."
            },
            paddingWidth: {
                ja: "コンテンツの左右に加えるアキです。",
                en: "Space added to the left and right of the content."
            },
            paddingHeight: {
                ja: "コンテンツの上下に加えるアキです。",
                en: "Space added above and below the content."
            },
            link: {
                ja: "幅と高さのパディングを同じ値にします。",
                en: "Keeps the width and height padding the same."
            },
            radiusEnabled: {
                ja: "角を丸めます。オフにすると角のままにします。",
                en: "Rounds the corners. When off, the corners are left square."
            },
            radius: {
                ja: "角丸の半径です。",
                en: "Radius of the rounded corners."
            },
            pill: {
                ja: "短い辺の半分を半径にして、両端が半円のピル形にします。",
                en: "Uses half of the shorter side as the radius, making a pill shape."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectError:   { ja: "選択エラー", en: "Selection Error" },
            invalidNumber: { ja: "数値を入力してください。", en: "Enter a numeric value." },
            invalidRadius: { ja: "角丸の半径は0以上の数値を入力してください。", en: "Enter a radius value of 0 or greater." },
            selectOne: {
                ja: "テキストまたはグループを1つ、もしくはテキスト/グループと図形を計2つ選択して実行してください。",
                en: "Select one text/group item, or one text/group and one shape (2 items total)."
            },
            clippingGroup: {
                ja: "クリッピンググループの計測には未対応です。クリッピングを解除するか、計測対象を単純なグループにしてください。",
                en: "Clipping groups are not supported for measurement. Release the clipping mask or use a simple group as the content item."
            },
            measureFailed: {
                ja: "コンテンツの計測に失敗しました。選択内容を確認してください。",
                en: "Failed to measure the content item. Check the selected objects and try again."
            }
        }
    };

    /**
     * 選択エラーの警告を表示する
     * @param {string} messageKey - LABELS.alert 配下のメッセージキー
     * @returns {void}
     */
    function alertSelectionError(messageKey) {
        alert(getLabel("alert." + messageKey), getLabel("alert.selectError"));
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

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報（不明なら pt）
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * pt 値を表示単位の文字列に変換する（小数第2位まで）
     * @param {number} valueInPt - pt 単位の値
     * @param {number} unitFactor - 表示単位1つあたりの pt 数
     * @returns {string} 表示用の文字列
     */
    function formatUnitValue(valueInPt, unitFactor) {
        return String(Math.round((valueInPt / unitFactor) * 100) / 100);
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ウィンドウに共通レイアウトを適用する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

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
     * 「項目名＋数値入力欄＋単位」の行を生成する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} fieldLabelText - 右揃えで表示する項目名
     * @param {string} initialText - 入力欄の初期値
     * @param {string} unitLabel - 入力欄の右に添える単位
     * @returns {EditText} 生成した入力欄
     */
    function addNumericFieldRow(parentContainer, fieldLabelText, initialText, unitLabel) {
        var fieldRow = parentContainer.add("group");
        setupRow(fieldRow);

        var fieldLabel = fieldRow.add("statictext", undefined, fieldLabelText);
        fieldLabel.preferredSize = [FIELD_LABEL_WIDTH, -1];
        fieldLabel.justify = "right";

        var inputField = addStepperInput(fieldRow, initialText);

        fieldRow.add("statictext", undefined, unitLabel);
        return inputField;
    }

    /**
     * 左に∧∨を付けた数値入力欄を追加する（∧∨と入力欄は隙間0で突き合わせる。下限0）。
     * ↑↓キーも∧∨と同じ処理で増減し、増減後は入力欄の onChanging（連動・プレビュー更新）を呼ぶ
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 初期値
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addStepperInput(parentRow, initialText) {
        var stepperFieldGroup = parentRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var inputField;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return inputField; }, {
            min: 0,
            onStep: function () {
                /* プログラム変更では onChanging が発火しないため明示的に呼ぶ / programmatic changes do not fire onChanging */
                if (typeof inputField.onChanging === "function") inputField.onChanging();
            }
        });
        inputField = stepperFieldGroup.add("edittext", undefined, initialText);
        inputField.characters = FIELD_CHARS;
        inputField.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(inputField, stepperGroup);
        return inputField;
    }

    /**
     * addStepperInput() で作った入力欄の有効／無効を、∧∨ごと切り替える
     * @param {EditText} inputField - 入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(inputField, isEnabled) {
        inputField.enabled = isEnabled;
        inputField.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(inputField.stepperGroup); /* ∧∨は自作描画なので描き直す / redraw the custom-drawn buttons */
    }

    // =========================================
    // 汎用ヘルパー / Generic helpers
    // =========================================

    /**
     * 例外を握り潰して処理を実行する
     * @param {function} fn - 実行する処理
     * @returns {void}
     */
    function safeDo(fn) {
        try { fn(); } catch (e) { }
    }

    /**
     * 例外を握り潰してアイテムを削除する
     * @param {PageItem} item - 削除するアイテム
     * @returns {void}
     */
    function safeRemove(item) {
        safeDo(function () { if (item) item.remove(); });
    }

    /**
     * 文字列を数値として解析し、無効なら既定値を返す
     * @param {string} value - 解析する文字列
     * @param {number} defaultValue - 解析できないときに返す値
     * @returns {number} 解析結果
     */
    function parseNumberOrDefault(value, defaultValue) {
        var num = parseFloat(value);
        return isNaN(num) ? defaultValue : num;
    }

    /**
     * 数値入力欄を検証する（負値は不可）
     * @param {EditText} editText - 対象の入力欄
     * @param {string} messageKey - LABELS.alert 配下のメッセージキー
     * @returns {number|null} 妥当な数値、不正なら null（警告表示済み）
     */
    function validateNumericField(editText, messageKey) {
        var value = parseFloat(editText.text);
        if (isNaN(value) || value < 0) {
            alert(getLabel("alert." + messageKey), getLabel("dialog.title"));
            editText.active = true;
            editText.selection = [0, editText.text.length];
            return null;
        }
        return value;
    }

    // =========================================
    // オブジェクト判定・計測 / Item checks and measurement
    // =========================================

    /**
     * @typedef {object} ContentBounds
     * @property {number} left - 左端
     * @property {number} top - 上端
     * @property {number} width - 幅
     * @property {number} height - 高さ
     * @property {number} centerX - 中心X
     * @property {number} centerY - 中心Y
     */

    /**
     * アイテムの境界情報を取得する
     * @param {PageItem} item - 対象アイテム
     * @returns {ContentBounds} 境界情報
     */
    function getBoundsFromItem(item) {
        var vb = item.visibleBounds;
        return {
            left: vb[0],
            top: vb[1],
            width: vb[2] - vb[0],
            height: vb[1] - vb[3],
            centerX: vb[0] + ((vb[2] - vb[0]) / 2),
            centerY: vb[1] - ((vb[1] - vb[3]) / 2)
        };
    }

    /**
     * 図形の幾何学的中心（線幅・効果を除いたパスの中心）を指定座標に合わせる
     * @param {PageItem} item - 対象アイテム
     * @param {number} centerX - 合わせ先の中心X
     * @param {number} centerY - 合わせ先の中心Y
     * @returns {void}
     */
    function centerByGeometry(item, centerX, centerY) {
        var gb = item.geometricBounds;
        var geoCenterX = gb[0] + ((gb[2] - gb[0]) / 2);
        var geoCenterY = gb[1] - ((gb[1] - gb[3]) / 2);
        item.translate(centerX - geoCenterX, centerY - geoCenterY);
    }

    /**
     * パス自体（線幅を含まない）が指定サイズになるよう拡大縮小し、中心を合わせる。
     * width / height は線幅込みの実寸なので、比率を掛けてから設定する。
     * @param {PageItem} item - 対象アイテム
     * @param {number} targetWidth - パスの目標幅
     * @param {number} targetHeight - パスの目標高さ
     * @param {number} centerX - 合わせ先の中心X
     * @param {number} centerY - 合わせ先の中心Y
     * @returns {void}
     */
    function resizeAndCenterByGeometry(item, targetWidth, targetHeight, centerX, centerY) {
        var gb = item.geometricBounds;
        var geoWidth = gb[2] - gb[0];
        var geoHeight = gb[1] - gb[3];
        if (geoWidth > 0) item.width = item.width * (targetWidth / geoWidth);
        if (geoHeight > 0) item.height = item.height * (targetHeight / geoHeight);
        centerByGeometry(item, centerX, centerY);
    }

    /**
     * 座布団を敷く対象（テキスト／グループ）かどうかを判定する
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} コンテンツ対象なら true
     */
    function isContentItem(item) {
        return !!(item && (item.typename === "TextFrame" || item.typename === "GroupItem"));
    }

    /**
     * 座布団として使える図形かどうかを判定する
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} 図形なら true
     */
    function isShapeItem(item) {
        return !!(item && (item.typename === "PathItem" || item.typename === "CompoundPathItem"));
    }

    /**
     * クリッピンググループかどうかを判定する
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} クリッピンググループなら true
     */
    function isClippingGroupItem(item) {
        return !!(item && item.typename === "GroupItem" && item.clipped);
    }

    // =========================================
    // アピアランス退避・復元 / Appearance capture and restore
    // =========================================

    /**
     * カラー値を複製する
     * @param {object} color - 複製元のカラー
     * @returns {object|null} 複製したカラー
     */
    function cloneColorValue(color) {
        if (!color) return null;

        var cloned;
        switch (color.typename) {
            case "RGBColor":
                cloned = new RGBColor();
                cloned.red = color.red;
                cloned.green = color.green;
                cloned.blue = color.blue;
                return cloned;
            case "CMYKColor":
                cloned = new CMYKColor();
                cloned.cyan = color.cyan;
                cloned.magenta = color.magenta;
                cloned.yellow = color.yellow;
                cloned.black = color.black;
                return cloned;
            case "GrayColor":
                cloned = new GrayColor();
                cloned.gray = color.gray;
                return cloned;
            case "SpotColor":
                cloned = new SpotColor();
                cloned.spot = color.spot;
                cloned.tint = color.tint;
                return cloned;
            case "PatternColor":
                cloned = new PatternColor();
                cloned.pattern = color.pattern;
                return cloned;
            case "GradientColor":
                cloned = new GradientColor();
                cloned.gradient = color.gradient;
                cloned.angle = color.angle;
                cloned.length = color.length;
                cloned.matrix = color.matrix;
                cloned.origin = color.origin;
                cloned.hiliteAngle = color.hiliteAngle;
                cloned.hiliteLength = color.hiliteLength;
                return cloned;
            case "NoColor":
                return new NoColor();
            default:
                return color;
        }
    }

    /**
     * 図形の基本スタイルを退避する
     * @param {PageItem} shapeItem - 対象図形
     * @returns {object} 退避したスタイル情報
     */
    function captureShapeStyle(shapeItem) {
        return {
            filled: !!shapeItem.filled,
            fillColor: shapeItem.filled ? cloneColorValue(shapeItem.fillColor) : null,
            stroked: !!shapeItem.stroked,
            strokeColor: shapeItem.stroked ? cloneColorValue(shapeItem.strokeColor) : null,
            strokeWidth: shapeItem.strokeWidth,
            opacity: shapeItem.opacity
        };
    }

    /**
     * 退避した基本スタイルを図形に復元する
     * @param {PageItem} shapeItem - 対象図形
     * @param {object} styleInfo - captureShapeStyle() の戻り値
     * @returns {void}
     */
    function restoreShapeStyle(shapeItem, styleInfo) {
        if (!shapeItem || !styleInfo) return;

        shapeItem.opacity = styleInfo.opacity;

        shapeItem.filled = styleInfo.filled;
        if (styleInfo.filled && styleInfo.fillColor) {
            shapeItem.fillColor = cloneColorValue(styleInfo.fillColor);
        }

        shapeItem.stroked = styleInfo.stroked;
        if (styleInfo.stroked && styleInfo.strokeColor) {
            shapeItem.strokeColor = cloneColorValue(styleInfo.strokeColor);
            shapeItem.strokeWidth = styleInfo.strokeWidth;
        }
    }

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /**
     * ダイナミックアクションで「アピアランスを消去」を実行する（塗り・線・不透明度は復元）
     * @param {PageItem} targetItem - 対象アイテム
     * @returns {void}
     */
    function clearAppearanceByAction(targetItem) {
        if (!targetItem) return;

        var actionDefinition = [
            '/version 3',
            '/name [ 10',
            ' 417070656172616e6365',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            ' /name [ 5',
            ' 636c656172',
            ' ]',
            ' /keyIndex 0',
            ' /colorIndex 0',
            ' /isOpen 1',
            ' /eventCount 1',
            ' /event-1 {',
            ' /useRulersIn1stQuadrant 0',
            ' /internalName (ai_plugin_appearance)',
            ' /localizedName [ 18',
            ' e382a2e38394e382a2e383a9e383b3e382b9',
            ' ]',
            ' /isOpen 1',
            ' /isOn 1',
            ' /hasDialog 0',
            ' /parameterCount 1',
            ' /parameter-1 {',
            ' /key 1835363957',
            ' /showInPalette 4294967295',
            ' /type (enumerated)',
            ' /name [ 27',
            ' e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb',
            ' ]',
            ' /value 6',
            ' }',
            ' }',
            '}'
        ].join('');

        var doc = app.activeDocument;
        var originalStyle = captureShapeStyle(targetItem);

        try {
            doc.selection = null;
            targetItem.selected = true;

            /* 失敗は従来どおり例外で伝える / Report a failure as an exception, as before */
            if (!runTemporaryAction(actionDefinition, 'Appearance', 'clear')) {
                throw new Error('Could not run the action: Appearance / clear');
            }
            restoreShapeStyle(targetItem, originalStyle);
        } finally {
            doc.selection = null;
        }
    }

    // =========================================
    // プレビュー undo ヘルパー / Preview undo helpers
    //
    // 「設定変更のたびに 前回 undo → 再生成」を回す共通パターン。
    // process() は ScriptUI 値から毎回再構築する純粋関数として書く。
    // - runPreview: UI 変更ごとに呼ぶ
    // - undoPreview: OK 直前・キャンセル・ダイアログクローズ時に前回プレビューを巻き戻す
    // =========================================

    /**
     * 直前のプレビューを巻き戻してから、プレビューを作り直す
     * @param {{isUndo: boolean}} previewState - プレビューの undo 状態
     * @param {function} processFn - プレビューを組み立てる処理
     * @returns {void}
     */
    function runPreview(previewState, processFn) {
        try {
            if (previewState.isUndo) app.undo();
            else previewState.isUndo = true;
            processFn();
            app.redraw();
        } catch (err) { }
    }

    /**
     * 残っているプレビューを巻き戻す
     * @param {{isUndo: boolean}} previewState - プレビューの undo 状態
     * @returns {void}
     */
    function undoPreview(previewState) {
        try {
            if (previewState.isUndo) app.undo();
        } catch (err) { }
        previewState.isUndo = false;
    }

    // =========================================
    // プレビュー値の計算 / Preview value math
    // =========================================

    /**
     * パディングを加えた座布団に収まる角丸半径の上限を求める（短辺の半分）
     * @param {number} addW - 幅のパディング（pt）
     * @param {number} addH - 高さのパディング（pt）
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @returns {number} 半径の上限（pt）
     */
    function getMaxRadius(addW, addH, bounds) {
        var shapeWidth = Math.max(MIN_PREVIEW_SIZE, bounds.width + addW);
        var shapeHeight = Math.max(MIN_PREVIEW_SIZE, bounds.height + addH);
        return Math.min(shapeWidth, shapeHeight) / 2;
    }

    /**
     * UI の入力値から、プレビューに使う実値（pt）と表示用テキストを求める
     * @param {{adjustEnabled:boolean, radiusEnabled:boolean, pill:boolean, addW:number, addH:number, radius:number}} uiValues - 入力欄の状態（数値は pt）
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @param {number} unitFactor - 表示単位1つあたりの pt 数
     * @returns {{addW:number, addH:number, radius:number, widthText:string, heightText:string, radiusText:string}} 計算結果
     */
    function computePreviewValues(uiValues, bounds, unitFactor) {
        var addW = uiValues.addW;
        var addH = uiValues.addH;
        var radius = (uiValues.radius < 0) ? 0 : uiValues.radius;

        if (!uiValues.adjustEnabled) {
            addW = 0;
            addH = 0;
            radius = 0;
        } else if (!uiValues.radiusEnabled) {
            radius = 0;
        } else if (uiValues.pill) {
            /* 半径＝高さの半分。幅も両端の半円分だけ広げる / Radius is half the height; widen by both caps */
            radius = Math.max(MIN_PREVIEW_SIZE, bounds.height + addH) / 2;
            addW = Math.max(MIN_PREVIEW_SIZE, radius * 2);
        } else {
            radius = Math.min(radius, getMaxRadius(addW, addH, bounds));
        }

        return {
            addW: addW,
            addH: addH,
            radius: radius,
            widthText: formatUnitValue(addW, unitFactor),
            heightText: formatUnitValue(addH, unitFactor),
            radiusText: formatUnitValue(radius, unitFactor)
        };
    }

    // =========================================
    // プレビューコントローラ / Preview controller
    // =========================================

    /**
     * @typedef {object} PreviewWidgets
     * @property {Checkbox} chkAdjustEnabled - 座布団の調整
     * @property {Checkbox} chkRadiusEnabled - 角丸の有効化
     * @property {Checkbox} chkPill - ピル形状
     * @property {Group} linkToggle - 幅と高さの連動（リンクアイコン）
     * @property {EditText} inputW - 幅のパディング
     * @property {EditText} inputH - 高さのパディング
     * @property {EditText} inputR - 角丸の半径
     */

    /**
     * @typedef {object} PreviewTemplates
     * @property {PageItem|null} previewSourceShapeItem - 整列専用の複製（元のアピアランスを保持）
     * @property {PageItem} previewBaseShapeItem - 調整用の複製（アピアランス消去済み）
     */

    /**
     * プレビューの計算・描画・確定を束ねたコントローラを生成する
     * @param {PreviewWidgets} widgets - ダイアログの各ウィジェット
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @param {PreviewTemplates} templates - プレビュー元の複製図形
     * @param {{isUndo: boolean}} previewState - プレビューの undo 状態
     * @param {number} unitFactor - 入力欄の単位1つあたりの pt 数
     * @returns {{reflectEnabled: function, refresh: function, getFinalValues: function, commitFinal: function}} コントローラ
     */
    function createPreviewController(widgets, bounds, templates, previewState, unitFactor) {

        /**
         * チェック状態から各ウィジェットの有効／無効を反映する
         * @returns {void}
         */
        function reflectEnabled() {
            var isAdjustEnabled = widgets.chkAdjustEnabled.value;
            var isRadiusEnabled = widgets.chkRadiusEnabled.value;
            var isPill = widgets.chkPill.value;
            var isLinked = widgets.linkToggle.value;

            /* ピル形状は幅を自動計算するので連動とは併用しない / Pill drives the width itself, so unlink */
            if (isPill && isLinked) {
                setLinkToggleValue(widgets.linkToggle, false);
                isLinked = false;
            }

            setStepperInputEnabled(widgets.inputW, isAdjustEnabled && !isPill);
            setStepperInputEnabled(widgets.inputH, isAdjustEnabled && (isPill || !isLinked));
            setStepperInputEnabled(widgets.inputR, isAdjustEnabled && isRadiusEnabled && !isPill);
            widgets.chkPill.enabled = isAdjustEnabled && isRadiusEnabled;
            setLinkToggleEnabled(widgets.linkToggle, isAdjustEnabled && !isPill);
            widgets.chkAdjustEnabled.enabled = true;
            widgets.chkRadiusEnabled.enabled = isAdjustEnabled;
        }

        /**
         * ウィジェットの現在値を読み取る（数値は表示単位から pt に換算）
         * @returns {object} UI の入力値
         */
        function readUIValues() {
            return {
                adjustEnabled: widgets.chkAdjustEnabled.value,
                radiusEnabled: widgets.chkRadiusEnabled.value,
                pill: widgets.chkPill.value,
                addW: parseNumberOrDefault(widgets.inputW.text, 0) * unitFactor,
                addH: parseNumberOrDefault(widgets.inputH.text, 0) * unitFactor,
                radius: parseNumberOrDefault(widgets.inputR.text, 0) * unitFactor
            };
        }

        /**
         * 自動計算した値を入力欄に書き戻す
         * @param {object} previewValues - computePreviewValues() の戻り値
         * @returns {void}
         */
        function applyDerivedUIValues(previewValues) {
            if (!widgets.chkAdjustEnabled.value) return;
            if (!widgets.chkRadiusEnabled.value) {
                widgets.inputR.text = "0";
                return;
            }
            widgets.inputR.text = previewValues.radiusText;
            if (widgets.chkPill.value) {
                widgets.inputW.text = previewValues.widthText;
                widgets.inputH.text = previewValues.heightText;
            }
        }

        /**
         * 現在の UI 値からプレビュー図形をゼロから作り直す。
         * undo のタイミングは runPreview / undoPreview が管理するので、
         * ユーザーの操作履歴には最後の1回分だけが残る。
         * 確定後に特定できるよう、末尾でプレビュー図形を選択状態にする。
         * @returns {void}
         */
        function process() {
            var previewValues = computePreviewValues(readUIValues(), bounds, unitFactor);
            applyDerivedUIValues(previewValues);

            var doc = app.activeDocument;
            var useAdjustedPreview = widgets.chkAdjustEnabled.value || !templates.previewSourceShapeItem;
            var previewTemplate = useAdjustedPreview ? templates.previewBaseShapeItem : templates.previewSourceShapeItem;

            var previewItem = previewTemplate.duplicate();
            previewItem.hidden = false;

            if (useAdjustedPreview) {
                /* パディングは線の外側ではなくパス基準。線幅を変えても余白が変わらない /
                   Padding is measured from the path, not the stroke, so stroke weight does not shift it */
                var newWidth = Math.max(MIN_PREVIEW_SIZE, bounds.width + previewValues.addW);
                var newHeight = Math.max(MIN_PREVIEW_SIZE, bounds.height + previewValues.addH);
                resizeAndCenterByGeometry(
                    previewItem, newWidth, newHeight, bounds.centerX, bounds.centerY
                );

                if (previewValues.radius > 0) {
                    var roundCornersXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + previewValues.radius + ' "/></LiveEffect>';
                    previewItem.applyEffect(roundCornersXml);
                }
            } else {
                /* 調整なしのときは元図形の寸法のまま、コンテンツの中心に合わせるだけ / Align only */
                centerByGeometry(previewItem, bounds.centerX, bounds.centerY);
            }

            /* 確定後に特定するため選択しておく / Select so the caller can identify it after OK */
            doc.selection = null;
            previewItem.selected = true;
        }

        /**
         * プレビューを更新する
         * @returns {void}
         */
        function refresh() {
            runPreview(previewState, process);
        }

        /**
         * 入力値を検証し、連動・ピル形状に合わせて入力欄を整える
         * @returns {boolean|null} 妥当なら true、不正なら null（警告表示済み）
         */
        function validateInputs() {
            if (!widgets.chkAdjustEnabled.value) return true;

            var addW = validateNumericField(widgets.inputW, "invalidNumber");
            if (addW === null) return null;

            var addH;
            if (widgets.linkToggle.value) {
                addH = addW;
            } else {
                addH = validateNumericField(widgets.inputH, "invalidNumber");
                if (addH === null) return null;
            }
            widgets.inputH.text = String(addH);

            if (!widgets.chkRadiusEnabled.value) {
                widgets.inputR.text = "0";
            } else if (!widgets.chkPill.value) {
                var radius = validateNumericField(widgets.inputR, "invalidRadius");
                if (radius === null) return null;
                widgets.inputR.text = String(radius);
            }
            return true;
        }

        /**
         * 入力を検証し、確定処理に必要な値を返す
         * @returns {{shouldRunPathfinder: boolean}|null} 確定値、不正なら null（警告表示済み）
         */
        function getFinalValues() {
            if (validateInputs() === null) return null;
            return {
                /* ピル形状は角丸効果を実体化する必要がある / Pill shapes must flatten the round-corner effect */
                shouldRunPathfinder: widgets.chkAdjustEnabled.value &&
                    widgets.chkRadiusEnabled.value && widgets.chkPill.value
            };
        }

        /**
         * 前回プレビューを巻き戻し、本番として1回だけ生成する
         * @returns {void}
         */
        function commitFinal() {
            undoPreview(previewState);
            process();
        }

        return {
            reflectEnabled: reflectEnabled,
            refresh: refresh,
            getFinalValues: getFinalValues,
            commitFinal: commitFinal
        };
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {object} DialogContext
     * @property {ContentBounds} bounds - コンテンツの境界情報
     * @property {PageItem|null} previewSourceShapeItem - 整列専用の複製
     * @property {PageItem} previewBaseShapeItem - 調整用の複製
     * @property {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     */

    /**
     * ダイアログのウィジェットを組み立てる
     * @param {Window} targetWindow - 追加先のウィンドウ
     * @param {object} sessionState - 前回の設定
     * @param {{code: number, label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     * @returns {PreviewWidgets} 生成したウィジェット一式
     */
    function buildDialogWidgets(targetWindow, sessionState, rulerUnit, shapeIsAutoCreated) {
        var adjustRow = targetWindow.add("group");
        setupRow(adjustRow, "center");
        var chkAdjustEnabled = adjustRow.add("checkbox", undefined, getLabel("checkbox.adjustEnabled"));
        chkAdjustEnabled.helpTip = getLabel("tooltip.adjustEnabled");
        chkAdjustEnabled.value = !!shapeIsAutoCreated;

        /* パディング / Padding */
        var paddingPanel = addPanel(targetWindow, getLabel("panel.padding"));
        var paddingRow = paddingPanel.add("group");
        setupRow(paddingRow, "left", COLUMN_SPACING);

        var paddingFields = paddingRow.add("group");
        paddingFields.orientation = "column";
        paddingFields.alignChildren = ["left", "center"];
        var inputW = addNumericFieldRow(paddingFields, labelText("fieldLabel.width"), sessionState.addW, rulerUnit.label);
        var inputH = addNumericFieldRow(paddingFields, labelText("fieldLabel.height"), sessionState.addH, rulerUnit.label);
        inputW.helpTip = getLabel("tooltip.paddingWidth");
        inputH.helpTip = getLabel("tooltip.paddingHeight");

        /* 幅・高さの2行の右、上下中央にリンクアイコンを置く。切り替え時の処理は bindDialogEvents() で handleToggle に入れる
           Link icon to the right of the two rows, vertically centred; bindDialogEvents() sets handleToggle */
        var linkToggle = addLinkToggle(paddingRow, !!sessionState.link, function () {
            if (linkToggle.handleToggle) linkToggle.handleToggle();
        });
        linkToggle.helpTip = getLabel("tooltip.link");

        /* 角丸 / Rounded corners */
        var cornerPanel = addPanel(targetWindow, getLabel("panel.corner"));
        var radiusRow = cornerPanel.add("group");
        setupRow(radiusRow);
        var chkRadiusEnabled = radiusRow.add("checkbox", undefined, labelText("fieldLabel.radius"));
        chkRadiusEnabled.helpTip = getLabel("tooltip.radiusEnabled");
        chkRadiusEnabled.value = (sessionState.radiusEnabled !== false);
        var inputR = addStepperInput(radiusRow, sessionState.radius);
        inputR.helpTip = getLabel("tooltip.radius");
        radiusRow.add("statictext", undefined, rulerUnit.label);

        var pillRow = cornerPanel.add("group");
        setupRow(pillRow);
        var chkPill = pillRow.add("checkbox", undefined, getLabel("checkbox.pill"));
        chkPill.helpTip = getLabel("tooltip.pill");
        chkPill.value = !!sessionState.pill;

        return {
            chkAdjustEnabled: chkAdjustEnabled,
            chkRadiusEnabled: chkRadiusEnabled,
            chkPill: chkPill,
            linkToggle: linkToggle,
            inputW: inputW,
            inputH: inputH,
            inputR: inputR
        };
    }

    /**
     * 入力欄・チェックボックスにイベントを結び付ける
     * @param {PreviewWidgets} widgets - ダイアログのウィジェット
     * @param {{reflectEnabled: function, refresh: function}} ctrl - プレビューコントローラ
     * @returns {void}
     */
    function bindDialogEvents(widgets, ctrl) {

        /**
         * 連動がONなら高さを幅にそろえる
         * @returns {void}
         */
        function syncHeightToWidth() {
            if (widgets.linkToggle.value) widgets.inputH.text = widgets.inputW.text;
        }

        /**
         * 連動がONなら幅を高さにそろえる
         * @returns {void}
         */
        function syncWidthToHeight() {
            if (widgets.linkToggle.value) widgets.inputW.text = widgets.inputH.text;
        }

        /**
         * 入力欄の負値を 0 に丸める
         * @param {EditText} editText - 対象の入力欄
         * @returns {void}
         */
        function clampNonNegativeText(editText) {
            var value = parseFloat(editText.text);
            if (!isNaN(value) && value < 0) editText.text = "0";
        }

        /**
         * 幅を変えたあと、連動を反映してプレビューを更新する
         * @returns {void}
         */
        function syncAndPreviewW() {
            syncHeightToWidth();
            ctrl.refresh();
        }

        /**
         * 高さを変えたあと、連動を反映してプレビューを更新する
         * @returns {void}
         */
        function syncAndPreviewH() {
            syncWidthToHeight();
            ctrl.refresh();
        }

        /**
         * チェックボックス用のハンドラを作る（変更を適用してから有効状態とプレビューを更新）
         * @param {function|null} mutateFn - 先に適用する処理（不要なら null）
         * @returns {function} onClick に割り当てるハンドラ
         */
        function checkboxHandler(mutateFn) {
            return function () {
                if (mutateFn) mutateFn();
                ctrl.reflectEnabled();
                ctrl.refresh();
            };
        }

        widgets.inputW.onChanging = function () {
            clampNonNegativeText(widgets.inputW);
            syncAndPreviewW();
        };
        widgets.inputH.onChanging = function () {
            clampNonNegativeText(widgets.inputH);
            syncAndPreviewH();
        };
        widgets.inputR.onChanging = function () {
            clampNonNegativeText(widgets.inputR);
            ctrl.refresh();
        };

        widgets.chkAdjustEnabled.onClick = checkboxHandler(null);
        widgets.linkToggle.handleToggle = checkboxHandler(syncHeightToWidth);
        widgets.chkPill.onClick = checkboxHandler(function () {
            /* ピル解除で幅が手入力に戻るので、連動していれば高さを合わせ直す /
               Leaving pill mode hands the width back to the user; re-mirror it when linked */
            if (!widgets.chkPill.value) syncHeightToWidth();
        });
        widgets.chkRadiusEnabled.onClick = checkboxHandler(function () {
            if (!widgets.chkRadiusEnabled.value) {
                widgets.inputR.text = "0";
                widgets.chkPill.value = false;
            }
            syncHeightToWidth();
        });

        /* 初期状態：連動ONなら高さは幅に追従 / Initial state: height follows width when linked */
        syncHeightToWidth();
    }

    /**
     * 現在の UI 状態からセッション保存用の設定を作る
     * @param {PreviewWidgets} widgets - ダイアログのウィジェット
     * @returns {object} 保存する設定
     */
    function buildSessionPayload(widgets) {
        return {
            addW: widgets.inputW.text,
            addH: widgets.inputH.text,
            radius: widgets.inputR.text,
            radiusEnabled: widgets.chkRadiusEnabled.value,
            link: widgets.linkToggle.value,
            pill: widgets.chkPill.value
        };
    }

    /**
     * 設定ダイアログを表示し、確定した内容を返す
     * @param {DialogContext} dialogContext - ダイアログに渡す情報
     * @returns {{shouldRunPathfinder: boolean, previewItem: PageItem}|null} 確定結果、キャンセル時は null
     */
    function showDialog(dialogContext) {
        var rulerUnit = getUnitInfo("rulerType");
        var sessionState = getSessionState(rulerUnit.pointsPerUnit);
        var previewState = { isUndo: false };
        var confirmedValues = null;
        var finalPreviewItem = null;

        var win = new Window('dialog', getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        setupWindow(win);

        var widgets = buildDialogWidgets(win, sessionState, rulerUnit, dialogContext.shapeIsAutoCreated);
        var ctrl = createPreviewController(
            widgets,
            dialogContext.bounds,
            {
                previewSourceShapeItem: dialogContext.previewSourceShapeItem,
                previewBaseShapeItem: dialogContext.previewBaseShapeItem
            },
            previewState,
            rulerUnit.pointsPerUnit
        );
        bindDialogEvents(widgets, ctrl);

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(win);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnOK.onClick = function () {
            confirmedValues = ctrl.getFinalValues();
            if (!confirmedValues) return;

            saveSessionState(buildSessionPayload(widgets));

            /* 前回プレビューを巻き戻して、本番として1回だけ確定 / Undo the preview, then commit once */
            ctrl.commitFinal();

            /* commitFinal はプレビュー図形を選択状態で残す / commitFinal leaves the item selected */
            var currentSelection = app.activeDocument.selection;
            if (currentSelection && currentSelection.length > 0) finalPreviewItem = currentSelection[0];

            win.close(1);
        };

        btnCancel.onClick = function () {
            saveSessionState(buildSessionPayload(widgets));
            /* 残ったプレビューは win.onClose で巻き戻す / onClose reverts the leftover preview */
            win.close(0);
        };

        /* OK / キャンセル / Esc / 閉じるボタン、どの経路でも残プレビューを片付ける /
           Catch-all: revert any leftover preview however the dialog closes */
        win.onClose = function () {
            undoPreview(previewState);
        };

        /* 初回プレビューは win.onShow（コールバック）から起動する。同期側で呼ぶと、
           後続コールバックの app.undo() が main のセットアップ（テンプレート作成等）まで巻き戻すおそれがある。
           Initial preview is fired from win.onShow — calling it synchronously risks a later
           app.undo() rolling back main()'s template setup. */
        win.onShow = function () {
            widgets.inputW.active = true;
            widgets.inputW.selection = [0, widgets.inputW.text.length];
            if (widgets.chkPill.value) setLinkToggleValue(widgets.linkToggle, false);
            ctrl.reflectEnabled();
            ctrl.refresh();
        };

        prepareDialogWindow(win, SCRIPT_NAME);
        if (win.show() !== 1 || !confirmedValues) {
            /* キャンセル時は onClose の undo で未適用に戻っている / onClose already reverted everything */
            safeRemove(finalPreviewItem);
            return null;
        }

        return {
            shouldRunPathfinder: confirmedValues.shouldRunPathfinder,
            previewItem: finalPreviewItem
        };
    }

    // =========================================
    // 選択の解釈 / Selection parsing
    // =========================================

    /**
     * 「コンテンツ1つ＋図形1つ」だけで構成されたグループから、その2つを取り出す。
     * グループは解除も削除もせず、条件に合わなければ null を返す。
     * @param {GroupItem} groupItem - 対象グループ
     * @returns {{contentItem: PageItem, shapeItem: PageItem}|null} 取り出した組、対象外なら null
     */
    function findContentAndShapeInGroup(groupItem) {
        if (!groupItem || groupItem.typename !== "GroupItem") return null;
        if (groupItem.clipped) return null;
        if (groupItem.pageItems.length !== 2) return null;

        var contentChild = null;
        var shapeChild = null;

        for (var i = 0; i < 2; i++) {
            var child = groupItem.pageItems[i];
            if (isContentItem(child) && !contentChild) {
                contentChild = child;
            } else if (isShapeItem(child) && !shapeChild) {
                shapeChild = child;
            } else {
                return null;
            }
        }

        if (!contentChild || !shapeChild) return null;

        return { contentItem: contentChild, shapeItem: shapeChild };
    }

    /**
     * 選択内容を検証してコンテンツと図形に振り分ける
     * @param {Array<PageItem>} currentSelection - ドキュメントの選択
     * @returns {{contentItem: PageItem, shapeItem: PageItem|null, shapeIsAutoCreated: boolean}|null} 振り分け結果、不正なら null
     */
    function parseSelection(currentSelection) {
        if (!currentSelection || currentSelection.length < 1 || currentSelection.length > 2) {
            alertSelectionError("selectOne");
            return null;
        }

        var contentItem = null;
        var shapeItem = null;
        var shapeIsAutoCreated = false;

        if (currentSelection.length === 2) {
            /* 2つ選択：テキスト/グループ＋図形 / Two items: text or group, plus a shape */
            for (var i = 0; i < currentSelection.length; i++) {
                var item = currentSelection[i];
                if (isContentItem(item) && !contentItem) {
                    contentItem = item;
                } else if (isShapeItem(item) && !shapeItem) {
                    shapeItem = item;
                } else {
                    alertSelectionError("selectOne");
                    return null;
                }
            }
        } else {
            /* 1つ選択：テキスト＋図形のグループなら、グループを保ったまま中身を使い分ける。
               それ以外はコンテンツとみなして長方形を自動作成する /
               One item: reuse the members of a text+shape group in place; otherwise auto-create a rectangle */
            var selectedItem = currentSelection[0];
            var groupMembers = null;
            if (selectedItem && selectedItem.typename === "GroupItem" && !isClippingGroupItem(selectedItem)) {
                groupMembers = findContentAndShapeInGroup(selectedItem);
            }
            if (groupMembers) {
                contentItem = groupMembers.contentItem;
                shapeItem = groupMembers.shapeItem;
            } else {
                if (!isContentItem(selectedItem)) {
                    alertSelectionError("selectOne");
                    return null;
                }
                contentItem = selectedItem;
                shapeIsAutoCreated = true;
            }
        }

        if (isClippingGroupItem(contentItem)) {
            alertSelectionError("clippingGroup");
            return null;
        }

        return {
            contentItem: contentItem,
            shapeItem: shapeItem,
            shapeIsAutoCreated: shapeIsAutoCreated
        };
    }

    /**
     * コンテンツを複製して境界を計測する（テキストはアウトライン化して字面を測る）
     * @param {PageItem} contentItem - 計測対象
     * @returns {ContentBounds|null} 境界情報、失敗時は null（警告表示済み）
     */
    function measureContent(contentItem) {
        var measureItem = null;
        var dupText = null;
        var bounds = null;

        try {
            if (contentItem.typename === "TextFrame") {
                dupText = contentItem.duplicate();
                measureItem = dupText.createOutline();
            } else {
                measureItem = contentItem.duplicate();
            }
            bounds = getBoundsFromItem(measureItem);
        } catch (err) {
            bounds = null;
        } finally {
            safeRemove(measureItem);
            safeRemove(dupText);
        }

        if (!bounds) {
            alertSelectionError("measureFailed");
            return null;
        }
        return bounds;
    }

    // =========================================
    // 図形の準備と確定 / Shape preparation and commit
    // =========================================

    /**
     * 必要なら長方形を自動作成し、プレビュー用の複製を用意して元図形を隠す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} contentItem - コンテンツ
     * @param {PageItem|null} shapeItem - 選択された図形（自動作成時は null）
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成するかどうか
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @returns {{shapeItem: PageItem, previewSourceShapeItem: PageItem|null, previewBaseShapeItem: PageItem}} 準備結果
     */
    function prepareShapeAndPreviews(doc, contentItem, shapeItem, shapeIsAutoCreated, bounds) {
        var previewSourceShapeItem = null;
        var previewBaseShapeItem;

        if (shapeIsAutoCreated) {
            /* コンテンツと同じ大きさの長方形を作り、それをプレビューの元にする /
               Create a rectangle the size of the content and use it as the preview template */
            shapeItem = doc.pathItems.rectangle(bounds.top, bounds.left, bounds.width, bounds.height);
            shapeItem.opacity = AUTO_CREATED_SHAPE_OPACITY;
            shapeItem.move(contentItem, ElementPlacement.PLACEAFTER);
            previewBaseShapeItem = shapeItem.duplicate();
        } else {
            /* 既存図形：整列専用（アピアランスそのまま）と調整専用（アピアランス消去）の2本立て。
               アピアランスが載ったままだと拡大縮小で見た目が崩れる /
               Existing shape: one duplicate as-is for align-only, one cleared so resizing cannot distort it */
            previewSourceShapeItem = shapeItem.duplicate();
            previewSourceShapeItem.hidden = true;

            previewBaseShapeItem = shapeItem.duplicate();
            clearAppearanceByAction(previewBaseShapeItem);
        }

        previewBaseShapeItem.hidden = true;
        shapeItem.hidden = true;
        app.redraw();

        return {
            shapeItem: shapeItem,
            previewSourceShapeItem: previewSourceShapeItem,
            previewBaseShapeItem: previewBaseShapeItem
        };
    }

    /**
     * キャンセル時の後片付け（テンプレート削除、自動作成した長方形の削除、元図形の再表示）
     * @param {object} prepared - prepareShapeAndPreviews() の戻り値
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     * @returns {void}
     */
    function cancelDialogResult(prepared, shapeIsAutoCreated) {
        safeRemove(prepared.previewSourceShapeItem);
        safeRemove(prepared.previewBaseShapeItem);
        if (shapeIsAutoCreated) {
            safeRemove(prepared.shapeItem);
        } else {
            safeDo(function () { prepared.shapeItem.hidden = false; });
        }
        app.redraw();
    }

    /**
     * OK 確定時の後処理（テンプレート削除、必要ならライブパスファインダー、元図形削除、最終選択）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} result - showDialog() の戻り値
     * @param {object} prepared - prepareShapeAndPreviews() の戻り値
     * @param {PageItem} contentItem - コンテンツ
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     * @returns {void}
     */
    function commitDialogResult(doc, result, prepared, contentItem, shapeIsAutoCreated) {
        /* プレビュー用テンプレートは複製元なので確定図形とは常に別物 / Templates are never the committed item */
        safeRemove(prepared.previewSourceShapeItem);
        safeRemove(prepared.previewBaseShapeItem);

        if (!result.previewItem) {
            /* 確定図形を特定できなかった：元の状態に戻す / Could not identify the item; roll back */
            cancelDialogResult(prepared, shapeIsAutoCreated);
            return;
        }

        var finalPreviewItem = result.previewItem;
        finalPreviewItem.hidden = false;

        if (result.shouldRunPathfinder) {
            /* ピル形状は角丸効果を実体化してから合成する / Flatten the round-corner effect for pill shapes */
            doc.selection = null;
            finalPreviewItem.selected = true;
            app.executeMenuCommand('Live Pathfinder Add');
            if (!doc.selection || doc.selection.length !== 1) {
                throw new Error('Live Pathfinder Add did not return exactly one selected item.');
            }
            finalPreviewItem = doc.selection[0];
            finalPreviewItem.hidden = false;
        }

        try {
            prepared.shapeItem.remove();
        } catch (removeError) {
            if (!shapeIsAutoCreated) prepared.shapeItem.hidden = false;
            throw removeError;
        }

        doc.selection = null;
        contentItem.selected = true;
        finalPreviewItem.selected = true;
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var parsed = parseSelection(doc.selection);
        if (!parsed) return;

        var bounds = measureContent(parsed.contentItem);
        if (!bounds) return;

        var prepared = prepareShapeAndPreviews(
            doc, parsed.contentItem, parsed.shapeItem, parsed.shapeIsAutoCreated, bounds
        );

        var result = showDialog({
            bounds: bounds,
            previewSourceShapeItem: prepared.previewSourceShapeItem,
            previewBaseShapeItem: prepared.previewBaseShapeItem,
            shapeIsAutoCreated: parsed.shapeIsAutoCreated
        });

        if (result) {
            commitDialogResult(doc, result, prepared, parsed.contentItem, parsed.shapeIsAutoCreated);
        } else {
            cancelDialogResult(prepared, parsed.shapeIsAutoCreated);
        }
    }

    main();

})();
