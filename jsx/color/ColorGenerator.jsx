#target illustrator
#targetengine "ColorGeneratorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

カラーパレットを生成してアートボードに描画し、スウォッチグループにも登録します。
Tailwind / Lightness / Saturation / Complementary / LCH のアルゴリズムに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorGenerator.md

### Overview

Generates color palettes, draws them on the artboard and registers them as a swatch group.
Tailwind, Lightness, Saturation, Complementary and LCH algorithms are supported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorGenerator.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ColorGenerator";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorGenerator.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorGenerator.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    /* 描画の余白は常に mm で指定する / drawing margins are always given in mm */
    var POINTS_PER_MM = UNITS[1].pointsPerUnit;

    var LABELS = {
        dialogTitle: { ja: "カラージェネレーター", en: "Color Generator" },
        panelSettings: { ja: "ベースカラー", en: "Base Color" },
        labelHex: { ja: "HEX", en: "HEX" },
        labelSteps: { ja: "ステップ数", en: "Steps" },
        panelContrast: { ja: "コントラストシフト", en: "Contrast Shift" },
        algoAll: { ja: "すべて", en: "All" },
        panelAlgorithm: { ja: "アルゴリズム", en: "Algorithm" },
        panelPreview: { ja: "プレビュー", en: "Preview" },
        panelSwatch: { ja: "スウォッチ", en: "Swatch" },
        chkRegisterSwatchGroup: { ja: "スウォッチグループに登録", en: "Register as Swatch Group" },
        chkConvertToGlobal: { ja: "グローバルに変換", en: "Convert to Global" },
        panelOutput: { ja: "出力", en: "Output" },
        chkOutputHex: { ja: "HEX", en: "HEX" },
        chkOutputRgb: { ja: "RGB", en: "RGB" },
        btnRedraw: { ja: "再描画", en: "Redraw" },
        btnCancel: { ja: "キャンセル", en: "Cancel" },
        btnGenerate: { ja: "生成", en: "Generate" },
        alertNoDoc: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        tipHex: { ja: "基準にするカラーを16進数で指定します（例: 3366CC）。", en: "The base color, given as a hex value (for example 3366CC)." },
        tipSteps: { ja: "作る色の数です。スライダーでも変えられます。", en: "How many colors to generate. The slider changes it too." },
        tipContrast: { ja: "明暗の開きを調整します。マイナスで中間に寄り、プラスで開きます。", en: "Adjusts how far apart the light and dark ends sit. Negative pulls them together, positive spreads them." },
        tipAlgorithm: { ja: "色の並べ方です。「すべて」を選ぶと各方式を並べて見比べられます。", en: "How the colors are derived. All lays out every method side by side for comparison." },
        tipRegisterSwatchGroup: { ja: "生成した色をスウォッチグループとしてまとめて登録します。", en: "Registers the generated colors together as a swatch group." },
        tipConvertToGlobal: { ja: "登録するスウォッチをグローバルカラーにします。", en: "Registers the swatches as global colors." },
        tipOutputHex: { ja: "各色の下に16進数の値を添えます。", en: "Writes the hex value under each color." },
        tipOutputRgb: { ja: "各色の下にRGB値を添えます。", en: "Writes the RGB value under each color." },
        tipRedraw: { ja: "いまの設定でプレビューを引き直します。", en: "Redraws the preview with the current settings." },
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
        }
    };

    /* 選択オブジェクトの塗りからHEXを取得 / Get HEX from selection fill */
    function rgbToHexString(r, g, b) {
        function to2(n) {
            var s = Math.round(n).toString(16);
            return (s.length === 1) ? ("0" + s) : s;
        }
        return "#" + to2(r) + to2(g) + to2(b);
    }

    /* CMYK → RGB 変換 / Convert CMYK to RGB (using Illustrator if possible) */
    function cmykToRgbApprox(c, m, y, k) {
        // c,m,y,k are 0..100
        try {
            // Illustrator's converter (more accurate than naive formula)
            var out = app.convertSampleColor(
                ImageColorSpace.CMYK,
                [c, m, y, k],
                ImageColorSpace.RGB,
                ColorConvertPurpose.defaultpurpose,
                true
            );
            if (out && out.length >= 3) {
                return { r: out[0], g: out[1], b: out[2] };
            }
        } catch (e) { }

        // Fallback (simple approximation)
        var C = c / 100, M = m / 100, Y = y / 100, K = k / 100;
        var r = 255 * (1 - C) * (1 - K);
        var g = 255 * (1 - M) * (1 - K);
        var b = 255 * (1 - Y) * (1 - K);
        return { r: r, g: g, b: b };
    }

    function normalizeHexString(hex) {
        if (!hex) return null;
        hex = String(hex).replace(/\s+/g, "");
        if (hex.charAt(0) !== "#") hex = "#" + hex;
        // 3桁は6桁へ
        if (/^#[0-9A-Fa-f]{3}$/.test(hex)) {
            hex = "#" + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2) + hex.charAt(3) + hex.charAt(3);
        }
        if (!/^#[0-9A-Fa-f]{6}$/.test(hex)) return null;
        return "#" + hex.substring(1).toLowerCase();
    }

    function tryGetSelectionFillHex() {
        try {
            if (app.documents.length === 0) return null;
            var doc = app.activeDocument;
            if (!doc.selection || doc.selection.length < 1) return null;

            var it = doc.selection[0];

            // グループの場合は最初のパスを探す / If group, find first path
            if (it.typename === "GroupItem") {
                for (var i = 0; i < it.pageItems.length; i++) {
                    if (it.pageItems[i].typename === "PathItem") { it = it.pageItems[i]; break; }
                }
            }

            // テキストの場合は文字の塗り / If text, use character fill
            if (it.typename === "TextFrame") {
                try {
                    var c = it.textRange.characterAttributes.fillColor;
                    if (c && c.typename === "RGBColor") {
                        return normalizeHexString(rgbToHexString(c.red, c.green, c.blue));
                    }
                    if (c && c.typename === "CMYKColor") {
                        var rgb = cmykToRgbApprox(c.cyan, c.magenta, c.yellow, c.black);
                        return normalizeHexString(rgbToHexString(rgb.r, rgb.g, rgb.b));
                    }
                } catch (e) { }
                return null;
            }

            // パスの場合 / PathItem fill
            if (it.typename === "PathItem") {
                if (!it.filled) return null;
                var fc = it.fillColor;
                if (fc && fc.typename === "RGBColor") {
                    return normalizeHexString(rgbToHexString(fc.red, fc.green, fc.blue));
                }
                if (fc && fc.typename === "CMYKColor") {
                    var rgb2 = cmykToRgbApprox(fc.cyan, fc.magenta, fc.yellow, fc.black);
                    return normalizeHexString(rgbToHexString(rgb2.r, rgb2.g, rgb2.b));
                }
            }

            return null;
        } catch (e) {
            return null;
        }
    }

    /* コントラストシフト / Contrast shift (-1.0 .. 1.0)
     * shift = -1 => 全て中間(128)へ寄せる（コントラスト0）
     * shift =  0 => 変化なし
     * shift =  1 => コントラスト2倍（クリップあり）
     */
    function clamp(v, minV, maxV) {
        if (v < minV) return minV;
        if (v > maxV) return maxV;
        return v;
    }

    function applyContrastShiftToChannel(v, shift) {
        var factor = 1 + shift; // 0..2
        var out = (v - 128) * factor + 128;
        return clamp(Math.round(out), 0, 255);
    }

    function applyContrastShiftToRgb(rgb, shift) {
        if (!rgb) return rgb;
        shift = clamp(Number(shift) || 0, -1, 1);
        if (Math.abs(shift) < 1e-9) return rgb;
        return {
            r: applyContrastShiftToChannel(rgb.r, shift),
            g: applyContrastShiftToChannel(rgb.g, shift),
            b: applyContrastShiftToChannel(rgb.b, shift)
        };
    }

    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alertNoDoc"));
            return;
        }

        var win = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing = 15;

        // 上部：2カラム（左=設定 / 右=スウォッチ） / Top: 2 columns (Left=Settings / Right=Swatch)
        var gTop = win.add("group");
        gTop.orientation = "row";
        gTop.alignChildren = ["fill", "top"];
        gTop.alignment = "fill";

        // 左カラム（ベースカラー / ステップ数） / Left column (Base color / Steps)
        var leftCol = gTop.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];
        leftCol.alignment = "fill";

        // --- 1. 入力エリア ---
        var inputPanel = leftCol.add("panel", undefined, getLabel("panelSettings"));
        inputPanel.margins = [15, 20, 15, 10];
        inputPanel.orientation = "column";
        inputPanel.alignChildren = ["left", "top"];

        // 右カラム（スウォッチ / 出力） / Right column (Swatch / Output)
        var rightCol = gTop.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["fill", "top"];
        rightCol.alignment = "fill";

        // 右カラム：スウォッチ / Right column: Swatch
        var swatchPanel = rightCol.add("panel", undefined, getLabel("panelSwatch"));
        swatchPanel.margins = [15, 20, 15, 10];
        swatchPanel.orientation = "column";
        swatchPanel.alignChildren = ["left", "top"];
        // まだロジック未接続のためディム表示 / Disabled until wired
        swatchPanel.enabled = false;

        var chkRegisterSwatchGroup = swatchPanel.add("checkbox", undefined, getLabel("chkRegisterSwatchGroup"));
        chkRegisterSwatchGroup.helpTip = getLabel("tipRegisterSwatchGroup");
        chkRegisterSwatchGroup.value = true;
        var chkConvertToGlobal = swatchPanel.add("checkbox", undefined, getLabel("chkConvertToGlobal"));
        chkConvertToGlobal.helpTip = getLabel("tipConvertToGlobal");
        chkConvertToGlobal.value = false;

        // 右カラム：出力 / Right column: Output
        var outputPanel = rightCol.add("panel", undefined, getLabel("panelOutput"));
        outputPanel.margins = [15, 20, 15, 10];
        outputPanel.orientation = "column";
        outputPanel.alignChildren = ["left", "top"];

        var chkOutputHex = outputPanel.add("checkbox", undefined, getLabel("chkOutputHex"));
        chkOutputHex.helpTip = getLabel("tipOutputHex");
        chkOutputHex.value = true;

        var chkOutputRgb = outputPanel.add("checkbox", undefined, getLabel("chkOutputRgb"));
        chkOutputRgb.helpTip = getLabel("tipOutputRgb");
        chkOutputRgb.value = false;

        var hexRow = inputPanel.add("group");
        hexRow.orientation = "row";
        hexRow.alignChildren = ["left", "center"];

        // 現在のHEXカラー表示（カラーチップ） / Current HEX color swatch
        var colorSwatch = hexRow.add("panel");
        try { colorSwatch.margins = 0; } catch (e) { }
        colorSwatch.preferredSize = [46, 46];

        hexRow.add("statictext", undefined, labelText("labelHex"));

        var __initHex = tryGetSelectionFillHex() || "#3b82f6";
        var inputHex = hexRow.add("edittext", undefined, __initHex);
        inputHex.helpTip = getLabel("tipHex");
        inputHex.characters = 8;

        function getPreviewRgb() {
            try {
                var nh = normalizeHexString(inputHex.text);
                if (!nh) return null;
                var rgb = hexToRgb(nh);
                return rgb;
            } catch (e) { }
            return null;
        }

        function updateColorSwatch() {
            try { colorSwatch.update(); } catch (e) { }
        }

        colorSwatch.onDraw = function () {
            var g = this.graphics;
            var rgb = getPreviewRgb();
            var w = this.size[0];
            var h = this.size[1];

            // 背景（無効時） / Background when invalid
            var fill;
            if (!rgb) {
                fill = g.newBrush(g.BrushType.SOLID_COLOR, [0.85, 0.85, 0.85, 1]);
            } else {
                fill = g.newBrush(g.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
            }

            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(fill);

            // 枠線 / Border
            try {
                var stroke = g.newPen(g.PenType.SOLID_COLOR, [0.6, 0.6, 0.6, 1], 1);
                g.strokePath(stroke);
            } catch (e) { }
        };
        // 初期値を正規化 / Normalize initial value
        try {
            var nh = normalizeHexString(inputHex.text);
            if (nh) inputHex.text = nh;
        } catch (e) { }
        updateColorSwatch();

        // ステップ数パネル / Steps panel
        var stepsPanel = leftCol.add("panel", undefined, getLabel("labelSteps"));
        stepsPanel.margins = [15, 20, 15, 10];
        stepsPanel.orientation = "column";
        stepsPanel.alignChildren = ["left", "top"];

        // 上段：ラベル + 入力 / Top: label + input
        var stepsInputRow = stepsPanel.add("group");
        stepsInputRow.orientation = "row";
        stepsInputRow.alignChildren = ["left", "center"];
        stepsInputRow.alignment = "left";

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var countStepperFieldGroup = stepsInputRow.add("group");
        countStepperFieldGroup.orientation = "row";
        countStepperFieldGroup.alignChildren = ["left", "center"];
        countStepperFieldGroup.spacing = 0;
        countStepperFieldGroup.margins = 0;

        /* 整数・1〜20。増減後は onChange でスライダー同期とプレビュー更新 / integer 1-20; onChange syncs the slider and preview */
        var countStepperGroup = addStepper(countStepperFieldGroup, function () { return inputCount; }, {
            min: 1, max: 20, integer: true,
            onStep: function () { inputCount.onChange(); }
        });
        var inputCount = countStepperFieldGroup.add("edittext", undefined, "5");
        inputCount.helpTip = getLabel("tipSteps");
        inputCount.characters = 3;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(inputCount, countStepperGroup);

        // 下段：スライダー / Bottom: slider
        var stepsSliderRow = stepsPanel.add("group");
        stepsSliderRow.orientation = "row";
        stepsSliderRow.alignChildren = ["left", "center"];
        stepsSliderRow.alignment = "left";

        var sldCount = stepsSliderRow.add("slider", undefined, 5, 1, 20);
        sldCount.helpTip = getLabel("tipSteps");
        sldCount.preferredSize.width = 200;

        // CONTRAST SHIFT パネル / Contrast Shift panel
        var contrastPanel = leftCol.add("panel", undefined, getLabel("panelContrast"));
        contrastPanel.margins = [15, 20, 15, 10];
        contrastPanel.orientation = "column";
        contrastPanel.alignChildren = ["left", "top"];

        var contrastInputRow = contrastPanel.add("group");
        contrastInputRow.orientation = "row";
        contrastInputRow.alignChildren = ["left", "center"];
        contrastInputRow.alignment = "left";

        var inputContrast = contrastInputRow.add("edittext", undefined, "0.0");
        inputContrast.helpTip = getLabel("tipContrast");
        inputContrast.characters = 4;

        var contrastSliderRow = contrastPanel.add("group");
        contrastSliderRow.orientation = "row";
        contrastSliderRow.alignChildren = ["left", "center"];
        contrastSliderRow.alignment = "left";

        var sldContrast = contrastSliderRow.add("slider", undefined, 0, -1, 1);
        sldContrast.helpTip = getLabel("tipContrast");
        sldContrast.preferredSize.width = 200;

        function getContrastShiftValue() {
            var t = String(inputContrast.text);
            t = t.replace(/^\s+|\s+$/g, "");
            t = t.replace(",", ".");
            var v = parseFloat(t);
            if (!isFinite(v) || isNaN(v)) v = Number(sldContrast.value) || 0;
            v = clamp(v, -1, 1);
            v = Math.round(v * 10) / 10; // 表示は小数1桁
            return v;
        }

        function setContrastShiftValue(v) {
            v = clamp(Number(v) || 0, -1, 1);
            v = Math.round(v * 10) / 10;
            inputContrast.text = v.toFixed(1);
            sldContrast.value = v;
        }
        setContrastShiftValue(0);

        // CONTRAST SHIFT UI 有効/無効 / Enable/disable contrast UI
        function setContrastEnabled(v) {
            contrastPanel.enabled = !!v;
            inputContrast.enabled = !!v;
            sldContrast.enabled = !!v;
        }

        // --- 2. アルゴリズム選択 (ラジオボタン) ---
        var algoPanel = win.add("panel", undefined, getLabel("panelAlgorithm"));
        algoPanel.margins = [15, 20, 15, 10];
        algoPanel.orientation = "column";
        algoPanel.alignChildren = ["left", "top"];

        var rbTailwind = algoPanel.add("radiobutton", undefined, "Tailwind CSS");
        rbTailwind.helpTip = getLabel("tipAlgorithm");
        var rbLight = algoPanel.add("radiobutton", undefined, "Lightness Scale");
        rbLight.helpTip = getLabel("tipAlgorithm");
        var rbLightGeo = algoPanel.add("radiobutton", undefined, "Lightness Scale (Geometric)");
        var rbLch = algoPanel.add("radiobutton", undefined, "LCH");
        rbLch.helpTip = getLabel("tipAlgorithm");
        var rbSatur = algoPanel.add("radiobutton", undefined, "Saturation Scale");
        rbSatur.helpTip = getLabel("tipAlgorithm");
        var rbComple = algoPanel.add("radiobutton", undefined, "Complementary");
        rbComple.helpTip = getLabel("tipAlgorithm");
        var rbAll = algoPanel.add("radiobutton", undefined, getLabel("algoAll"));
        rbAll.value = true; // デフォルト

        // --- 3. プレビューエリア ---
        var previewPanel = win.add("panel", undefined, getLabel("panelPreview"));
        previewPanel.margins = [15, 20, 15, 10];
        previewPanel.size = [420, 70];

        previewPanel.onDraw = function () {
            var g = this.graphics;
            var data = getActivePalette();
            if (!data) return;
            var w = this.size[0] / data.length;
            for (var i = 0; i < data.length; i++) {
                var c0 = data[i];
                var shift = getContrastShiftValue();
                var c = applyContrastShiftToRgb(c0, shift);
                var brush = g.newBrush(g.BrushType.SOLID_COLOR, [c.r / 255, c.g / 255, c.b / 255, 1]);
                g.newPath();
                g.rectPath(i * w, 10, w, 40);
                g.fillPath(brush);
            }
        };

        // プレビュー更新（notifyより update の方が確実）
        var __previewEnabled = true;
        var updatePreview = function () {
            try {
                if (!__previewEnabled) return;
                previewPanel.update(); // onDraw を発火させる
            } catch (e) {
                try { previewPanel.notify("onDraw"); } catch (__) { }
            }
        };
        inputHex.onChanging = function () {
            updateColorSwatch();
            updatePreview();
        };

        // CONTRAST SHIFT のUI連動（updatePreview定義後に接続）
        inputContrast.onChanging = function () {
            var v = getContrastShiftValue();
            sldContrast.value = v;
            updatePreview();
        };
        inputContrast.onChange = function () {
            var v = getContrastShiftValue();
            setContrastShiftValue(v);
            updatePreview();
        };
        sldContrast.onChanging = function () {
            var v = Math.round(Number(sldContrast.value) * 10) / 10;
            setContrastShiftValue(v);
            updatePreview();
        };
        sldContrast.onChange = function () {
            var v = Math.round(Number(sldContrast.value) * 10) / 10;
            setContrastShiftValue(v);
            updatePreview();
        };

        // 入力欄 ⇄ スライダー同期
        function clampInt(v, minV, maxV) {
            v = Math.round(Number(v));
            if (!isFinite(v)) v = minV;
            if (v < minV) v = minV;
            if (v > maxV) v = maxV;
            return v;
        }

        inputCount.onChanging = function () {
            var v = clampInt(inputCount.text, 1, 20);
            sldCount.value = v;
            updatePreview();
        };

        // Tab移動など確定時にも反映
        inputCount.onChange = function () {
            var v = clampInt(inputCount.text, 1, 20);
            sldCount.value = v;
            inputCount.text = String(v);
            updatePreview();
        };

        sldCount.onChanging = function () {
            var v = clampInt(sldCount.value, 1, 20);
            inputCount.text = String(v);
            updatePreview();
        };

        // 念のため onChange でも反映
        sldCount.onChange = function () {
            var v = clampInt(sldCount.value, 1, 20);
            inputCount.text = String(v);
            updatePreview();
        };

        function setPreviewEnabled(v) {
            __previewEnabled = !!v;
            // visible は変えない（レイアウトが上下にガタつくため）
            previewPanel.enabled = __previewEnabled;
            // 見た目を即時反映
            try { previewPanel.update(); } catch (e) { }
        }

        // ステップUI有効/無効
        function setStepsEnabled(v) {
            stepsPanel.enabled = !!v;
            inputCount.enabled = !!v;
            sldCount.enabled = !!v;
            redrawSteppersIn(stepsPanel); /* ∧∨のディム表示を切り替える / update the stepper dimming */
        }

        // 初期状態：デフォルトは「すべて」なのでプレビュー無効 / Initial: default is "All" => disable preview
        setPreviewEnabled(false);
        setStepsEnabled(false);
        setContrastEnabled(false);

        rbTailwind.onClick = function () {
            // Tailwind は常に11・ステップ数UIはディム / Tailwind: fixed 11, dim steps UI
            inputCount.text = "11";
            sldCount.value = 11;
            setStepsEnabled(false);
            setContrastEnabled(true);
            setPreviewEnabled(true);
            updatePreview();
        };
        rbLight.onClick = function () {
            setStepsEnabled(true);
            setContrastEnabled(true);
            setPreviewEnabled(true);
            updatePreview();
        };
        rbLightGeo.onClick = function () {
            setStepsEnabled(true);
            setContrastEnabled(true);
            setPreviewEnabled(true);
            updatePreview();
        };
        rbSatur.onClick = function () {
            setStepsEnabled(true);
            setContrastEnabled(true);
            setPreviewEnabled(true);
            updatePreview();
        };
        rbComple.onClick = function () {
            setStepsEnabled(true);
            setContrastEnabled(true);
            setPreviewEnabled(true);
            updatePreview();
        };
        rbLch.onClick = function () {
            // LCH は常に11・ステップ数UIはディム / LCH: fixed 11, dim steps UI
            inputCount.text = "11";
            sldCount.value = 11;
            setStepsEnabled(false);
            setContrastEnabled(true);
            setPreviewEnabled(true);
            updatePreview();
        };
        rbAll.onClick = function () {
            // 「すべて」はステップ数=11で計算（UIも合わせる）
            inputCount.text = "11";
            sldCount.value = 11;
            // 「すべて」選択時はプレビューを無効化（レイアウトは固定）
            setPreviewEnabled(false);
            setStepsEnabled(false);
            setContrastEnabled(false);
        };

        // --- ボタン ---
        var buttonRow = addButtonRow(win);

        // 左：再描画 / Left: Redraw
        var btnRedraw = buttonRow.leftGroup.add("button", undefined, getLabel("btnRedraw"));
        btnRedraw.helpTip = getLabel("tipRedraw");

        // 右：キャンセル / 生成 / Right: Cancel / Generate
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("btnCancel"));
        var btnOk = buttonRow.rightGroup.add("button", undefined, getLabel("btnGenerate"), { name: "ok" });

        // 再描画 / Redraw (preview refresh)
        btnRedraw.onClick = function () {
            try { updateColorSwatch(); } catch (e) { }

            // ダイアログ内プレビューを強制再描画（プレビュー無効時も） / Force dialog preview repaint
            try {
                var __prev = __previewEnabled;
                __previewEnabled = true;
                try { previewPanel.update(); } catch (e) { }
                try { updatePreview(); } catch (e) { }
                __previewEnabled = __prev;
            } catch (e) { }

            // ダイアログ全体の再描画 / Refresh dialog
            try { win.update(); } catch (e) { }

            // Illustrator画面の再描画 / Force Illustrator UI redraw
            app.redraw();
        };

        // キャンセル / Cancel
        btnCancel.onClick = function () {
            try { win.close(); } catch (e) { }
        };

        // ステップ数取得（全角数字→半角、NaN時はスライダー値を採用、1〜20にクランプ）
        function getStepCount() {
            var t = String(inputCount.text);

            // 全角数字を半角へ
            t = t.replace(/[０-９]/g, function (ch) {
                return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
            });

            // 余計な空白を除去
            t = t.replace(/^\s+|\s+$/g, "");

            var v = parseInt(t, 10);

            // 文字入力が不正な場合はスライダーにフォールバック
            if (!isFinite(v) || isNaN(v)) {
                try { v = Math.round(Number(sldCount.value)); } catch (e) { v = 5; }
            }

            // クランプ（UI仕様：1〜20）
            if (v < 1) v = 1;
            if (v > 20) v = 20;

            return v;
        }

        function getActivePalette() {
            var hex = inputHex.text;
            var count = getStepCount();
            if (!/^#?([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/.test(hex)) return null;

            var rgb = hexToRgb(hex);
            var hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);

            if (rbAll.value) return null;
            if (rbTailwind.value) return algoTailwind(hsl);
            if (rbLight.value) return algoLightness(hsl, count);
            if (rbLightGeo.value) return algoLightnessGeometric(hsl, count);
            if (rbSatur.value) return algoSaturation(hsl, count);
            if (rbComple.value) return algoComplementary(hsl, count);
            if (rbLch.value) return algoLchFromRgb(rgb);
            return null;
        }

        function getActiveAlgorithmName() {
            if (rbTailwind.value) return rbTailwind.text;
            if (rbLight.value) return rbLight.text;
            if (rbLightGeo.value) return rbLightGeo.text;
            if (rbSatur.value) return rbSatur.text;
            if (rbComple.value) return rbComple.text;
            if (rbLch.value) return rbLch.text;
            if (rbAll.value) return rbAll.text;
            return "";
        }

        btnOk.onClick = function () {
            var data = getActivePalette();
            // OK押下時に値を正規化してUIへ反映（全角/空欄対策）
            var vFixed = getStepCount();
            inputCount.text = String(vFixed);
            sldCount.value = vFixed;
            // 出力オプション / Output options
            var outHex = !!chkOutputHex.value;
            var outRgb = !!chkOutputRgb.value;
            var contrastShift = getContrastShiftValue();
            if (rbAll.value) {
                var hex = inputHex.text;
                var rgb0 = hexToRgb(hex);
                var hsl0 = rgbToHsl(rgb0.r, rgb0.g, rgb0.b);

                // 「すべて」生成（固定順）
                win.close();
                drawAllToIllustrator(hsl0, rgb0, outHex, outRgb, contrastShift);
                return;
            }

            if (data) {
                var algoName = getActiveAlgorithmName();
                win.close();
                drawToIllustrator(data, algoName, outHex, outRgb, contrastShift);
            }
        };

        // Escキーでも閉じる + EnterでOK / Close with Esc, trigger OK with Enter
        win.addEventListener("keydown", function (e) {
            try {
                if (!e) return;
                if (e.keyName === "Escape") {
                    e.preventDefault();
                    win.close();
                    return;
                }
                if (e.keyName === "Enter" || e.keyName === "Return") {
                    e.preventDefault();
                    try { btnOk.notify("onClick"); } catch (e) {
                        try { btnOk.onClick(); } catch (__) { }
                    }
                    return;
                }
            } catch (e) { }
        });
        prepareDialogWindow(win, SCRIPT_NAME);
        win.show();
    }

    // --- アルゴリズムの実装 ---

    // 1. Tailwind風 (固定の輝度ステップ)
    function algoTailwind(hsl) {
        var weights = [0.95, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1, 0.05]; // 50-950
        var res = [];
        // 左→右を「濃 → 薄」に
        for (var i = 0; i < weights.length; i++) {
            res.push(hslToRgb(hsl.h, hsl.s, 1 - weights[i]));
        }
        return res;
    }

    // 2. Lightness Scale (輝度変化) - 出力は常に count 色 / Always output exactly "count" colors
    function algoLightness(hsl, count) {
        var n = Math.round(Number(count));
        if (!isFinite(n) || n < 1) n = 1;

        // n=1 のときは元色のみ / If n=1, only the base color
        if (n === 1) return [hslToRgb(hsl.h, hsl.s, hsl.l)];

        // 範囲を決めて等間隔にサンプリング / Sample evenly in a lightness range
        // - 暗側は元Lの30%〜最低0.06
        // - 明側は 0.94 まで（白飛び回避）
        var lMin = Math.max(0.06, hsl.l * 0.30);
        var lMax = Math.min(0.94, hsl.l + (1 - hsl.l) * 0.85);

        // もし範囲が狭すぎる場合は安全に広げる / Safety expand if needed
        if (lMax - lMin < 0.10) {
            lMin = Math.max(0.06, hsl.l - 0.15);
            lMax = Math.min(0.94, hsl.l + 0.15);
        }

        var res = [];
        for (var i = 0; i < n; i++) {
            var t = (n === 1) ? 0 : (i / (n - 1));
            var L = lMin + (lMax - lMin) * t;
            res.push(hslToRgb(hsl.h, hsl.s, L));
        }

        // 左→右を「濃 → 薄」にする / Dark -> Light left-to-right
        // （lMinが暗、lMaxが明なのでそのままでOK）
        return res;
    }

    // 2b. Lightness Scale (Geometric) - 出力は常に count 色 / Always output exactly "count" colors
    // 画像の「等比（幾何）スケール」的に、暗部側の変化を細かくする。
    // lMin..lMax の範囲を指数補間でサンプリング（濃→薄）
    function algoLightnessGeometric(hsl, count) {
        var n = Math.round(Number(count));
        if (!isFinite(n) || n < 1) n = 1;

        if (n === 1) return [hslToRgb(hsl.h, hsl.s, hsl.l)];

        // Linear版と同じ範囲をベースにする
        var lMin = Math.max(0.06, hsl.l * 0.30);
        var lMax = Math.min(0.94, hsl.l + (1 - hsl.l) * 0.85);

        if (lMax - lMin < 0.10) {
            lMin = Math.max(0.06, hsl.l - 0.15);
            lMax = Math.min(0.94, hsl.l + 0.15);
        }

        // 比率（lMin>0 を保証） / ratio
        var ratio = lMax / Math.max(0.0001, lMin);

        var res = [];
        for (var i = 0; i < n; i++) {
            var t = (i / (n - 1)); // 0..1
            // 幾何補間: L = lMin * ratio^t
            var L = lMin * Math.pow(ratio, t);
            // 念のためクランプ
            if (L < 0) L = 0;
            if (L > 1) L = 1;
            res.push(hslToRgb(hsl.h, hsl.s, L));
        }
        return res;
    }

    // 3. Saturation Scale (彩度変化) - 出力は常に count 色 / Always output exactly "count" colors
    function algoSaturation(hsl, count) {
        var n = Math.round(Number(count));
        if (!isFinite(n) || n < 1) n = 1;

        if (n === 1) return [hslToRgb(hsl.h, hsl.s, hsl.l)];

        // 彩度は低→高（元S）で等間隔 / Low saturation -> base saturation
        var sMin = 0.0;
        var sMax = Math.min(1.0, Math.max(0.0, hsl.s));

        var res = [];
        for (var i = 0; i < n; i++) {
            var t = (i / (n - 1));
            var S = sMin + (sMax - sMin) * t;
            res.push(hslToRgb(hsl.h, S, hsl.l));
        }
        return res;
    }

    // 4. Complementary (補色 + 輝度変化)
    function algoComplementary(hsl, count) {
        var n = Math.round(Number(count));
        if (!isFinite(n) || n < 1) n = 1;

        var compH = (hsl.h + 0.5) % 1.0;

        // 常に合計 n 色になるように分配（n=11 なら 6+5 など）
        var n1 = Math.ceil(n / 2);
        var n2 = n - n1;

        var res = algoLightness(hsl, n1);
        var compRes = (n2 > 0) ? algoLightness({ h: compH, s: hsl.s, l: hsl.l }, n2) : [];
        return res.concat(compRes);
    }

    /* 5. LCH (固定11ステップ)
     * - Lightness: 8刻みの10ステップ + 99% を追加して合計11
     * - Chroma: 58 を基準に、sRGB 色域内に収まるまで下げる
     * - Hue: 入力色の Hue を使用（LCHに変換して保持）
     */
    function algoLchFromRgb(baseRgb) {
        // base RGB -> LCH
        var lab = rgbToLab(baseRgb.r, baseRgb.g, baseRgb.b);
        var lch = labToLch(lab);

        var H = lch.h;

        // 8刻みの10ステップ（暗→明）＋ 99 を先頭に追加（合計11）
        // 90,82,74,66,58,50,42,34,26,18 に 99 を足す
        var Ls = [99, 90, 82, 74, 66, 58, 50, 42, 34, 26, 18];

        var res = [];
        for (var i = 0; i < Ls.length; i++) {
            var L = Ls[i];
            var C = 58;

            // gamut-fit: Cを下げてsRGB内に収める
            var rgb = null;
            for (var c = C; c >= 0; c -= 1) {
                var lab2 = lchToLab({ l: L, c: c, h: H });
                var rgb2 = labToRgb(lab2.l, lab2.a, lab2.b);
                if (rgb2 && rgb2.inGamut) {
                    rgb = { r: rgb2.r, g: rgb2.g, b: rgb2.b };
                    break;
                }
            }
            // それでも取れない場合はグレーにフォールバック
            if (!rgb) {
                rgb = { r: Math.round(L / 100 * 255), g: Math.round(L / 100 * 255), b: Math.round(L / 100 * 255) };
            }
            res.push(rgb);
        }
        res.reverse(); // 左→右を「濃 → 薄」に
        return res;
    }

    // --- ユーティリティ ---

    /* =========================================
     * Color space utils: sRGB ↔ XYZ(D65) ↔ Lab ↔ LCH
     * - 参照白色点: D65 (Xn=95.047, Yn=100.000, Zn=108.883)
     * - Lab: CIE 1976
     * ========================================= */

    function srgbToLinear(u) {
        u = u / 255;
        return (u <= 0.04045) ? (u / 12.92) : Math.pow((u + 0.055) / 1.055, 2.4);
    }

    function linearToSrgb(u) {
        var v = (u <= 0.0031308) ? (12.92 * u) : (1.055 * Math.pow(u, 1 / 2.4) - 0.055);
        return Math.round(Math.min(1, Math.max(0, v)) * 255);
    }

    function rgbToXyz(r, g, b) {
        var R = srgbToLinear(r);
        var G = srgbToLinear(g);
        var B = srgbToLinear(b);

        // sRGB D65
        var X = (R * 0.4124564 + G * 0.3575761 + B * 0.1804375) * 100;
        var Y = (R * 0.2126729 + G * 0.7151522 + B * 0.0721750) * 100;
        var Z = (R * 0.0193339 + G * 0.1191920 + B * 0.9503041) * 100;

        return { x: X, y: Y, z: Z };
    }

    function xyzToRgb(X, Y, Z) {
        X /= 100; Y /= 100; Z /= 100;

        var R = 3.2404542 * X + -1.5371385 * Y + -0.4985314 * Z;
        var G = -0.9692660 * X + 1.8760108 * Y + 0.0415560 * Z;
        var B = 0.0556434 * X + -0.2040259 * Y + 1.0572252 * Z;

        // gamut判定（linearの段階で 0..1 に入っているか）
        var inGamut = (R >= 0 && R <= 1 && G >= 0 && G <= 1 && B >= 0 && B <= 1);

        // clampしてsRGBへ
        R = Math.min(1, Math.max(0, R));
        G = Math.min(1, Math.max(0, G));
        B = Math.min(1, Math.max(0, B));

        return { r: linearToSrgb(R), g: linearToSrgb(G), b: linearToSrgb(B), inGamut: inGamut };
    }

    function fLab(t) {
        return (t > 0.008856451679) ? Math.pow(t, 1 / 3) : (7.787037037 * t + 16 / 116);
    }

    function finvLab(t) {
        var t3 = t * t * t;
        return (t3 > 0.008856451679) ? t3 : ((t - 16 / 116) / 7.787037037);
    }

    function xyzToLab(xyz) {
        // D65 white
        var Xn = 95.047, Yn = 100.000, Zn = 108.883;

        var x = fLab(xyz.x / Xn);
        var y = fLab(xyz.y / Yn);
        var z = fLab(xyz.z / Zn);

        var L = 116 * y - 16;
        var a = 500 * (x - y);
        var b = 200 * (y - z);

        return { l: L, a: a, b: b };
    }

    function labToXyz(lab) {
        var Xn = 95.047, Yn = 100.000, Zn = 108.883;

        var fy = (lab.l + 16) / 116;
        var fx = fy + (lab.a / 500);
        var fz = fy - (lab.b / 200);

        var xr = finvLab(fx);
        var yr = finvLab(fy);
        var zr = finvLab(fz);

        return { x: xr * Xn, y: yr * Yn, z: zr * Zn };
    }

    function rgbToLab(r, g, b) {
        return xyzToLab(rgbToXyz(r, g, b));
    }

    function labToRgb(L, a, b) {
        var xyz = labToXyz({ l: L, a: a, b: b });
        return xyzToRgb(xyz.x, xyz.y, xyz.z);
    }

    function labToLch(lab) {
        var C = Math.sqrt(lab.a * lab.a + lab.b * lab.b);
        var H = Math.atan2(lab.b, lab.a) * 180 / Math.PI;
        if (H < 0) H += 360;
        return { l: lab.l, c: C, h: H };
    }

    function lchToLab(lch) {
        var hr = lch.h * Math.PI / 180;
        var a = lch.c * Math.cos(hr);
        var b = lch.c * Math.sin(hr);
        return { l: lch.l, a: a, b: b };
    }

    function hexToRgb(hex) {
        if (hex.charAt(0) === '#') hex = hex.substring(1);
        if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        var num = parseInt(hex, 16);
        return { r: num >> 16, g: (num >> 8) & 0x00FF, b: num & 0x0000FF };
    }

    function rgbToHsl(r, g, b) {
        r /= 255, g /= 255, b /= 255;
        var max = Math.max(r, g, b), min = Math.min(r, g, b);
        var h, s, l = (max + min) / 2;
        if (max == min) { h = s = 0; }
        else {
            var d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }
            h /= 6;
        }
        return { h: h, s: s, l: l };
    }

    function hslToRgb(h, s, l) {
        var r, g, b;
        if (s == 0) { r = g = b = l; }
        else {
            function hue2rgb(p, q, t) {
                if (t < 0) t += 1; if (t > 1) t -= 1;
                if (t < 1 / 6) return p + (q - p) * 6 * t;
                if (t < 1 / 2) return q;
                if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
                return p;
            }
            var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            var p = 2 * l - q;
            r = hue2rgb(p, q, h + 1 / 3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1 / 3);
        }
        return { r: Math.round(r * 255), g: Math.round(g * 255), b: Math.round(b * 255) };
    }

    function drawToIllustrator(data, algoName, outHex, outRgb, contrastShift) {
        var doc = app.activeDocument;

        // Active artboard bounds: [left, top, right, bottom]
        var abIndex = doc.artboards.getActiveArtboardIndex();
        var abRect = doc.artboards[abIndex].artboardRect;
        var abLeft = abRect[0], abTop = abRect[1], abRight = abRect[2], abBottom = abRect[3];

        // 5mm マージン（pt換算）
        var m = 5 * POINTS_PER_MM;

        // inset rect
        var left = abLeft + m;
        var right = abRight - m;
        var top = abTop - m;
        var bottom = abBottom + m;

        var abW = right - left;
        var abH = top - bottom;

        var n = data.length;
        if (!n) return;

        // 見出し高さ（タイトル） / Header height
        var headerH = 24;
        if (abH < headerH + 5) headerH = 0;

        // スウォッチ高さ（従来どおり最大50pt） / Swatch height (cap 50pt)
        var swH = Math.min(50, abH - headerH);
        if (swH < 10) swH = 10;

        // 共通描画へ委譲 / Delegate to shared renderer
        drawPaletteRow(doc, left, top, abW, headerH, swH, data, algoName, outHex, outRgb, contrastShift);
    }

    // 「すべて」用：全アルゴリズムを縦に積んで描画
    function drawAllToIllustrator(hsl0, rgb0, outHex, outRgb, contrastShift) {
        var doc = app.activeDocument;

        // Active artboard bounds: [left, top, right, bottom]
        var abIndex = doc.artboards.getActiveArtboardIndex();
        var abRect = doc.artboards[abIndex].artboardRect;
        var abLeft = abRect[0], abTop = abRect[1], abRight = abRect[2], abBottom = abRect[3];

        // 5mm マージン（pt換算）
        var m = 5 * POINTS_PER_MM;

        // inset rect
        var left = abLeft + m;
        var right = abRight - m;
        var top = abTop - m;
        var bottom = abBottom + m;

        var abW = right - left;
        var abH = top - bottom;

        // 描画順（Tailwind, Lightness, Lightness Geometric, LCH, Saturation, Complementary）
        // 「すべて」選択時はステップ数=11固定
        var ALL_STEPS = 11;

        var blocks = [
            { name: "Tailwind CSS", data: algoTailwind(hsl0) },
            { name: "Lightness Scale", data: algoLightness(hsl0, ALL_STEPS) },
            { name: "Lightness Scale (Geometric)", data: algoLightnessGeometric(hsl0, ALL_STEPS) },
            { name: "LCH", data: algoLchFromRgb(rgb0) },
            { name: "Saturation Scale", data: algoSaturation(hsl0, ALL_STEPS) },
            { name: "Complementary", data: algoComplementary(hsl0, ALL_STEPS) }
        ];

        var nBlocks = blocks.length;

        var headerH = 24;
        var gap = 10;

        // ラベル行数に応じた領域（実際に出る行数ベース）
        var labelLines = 0;
        if (outHex) labelLines += 1;
        if (outRgb) labelLines += 1;

        // 5pt文字 + 行送り7pt（上詰めの実装に合わせる）
        var lineGap = 7;
        var labelAreaH = 0;
        if (labelLines > 0) {
            // 最上行の分 + (残り行数-1)*lineGap + 下余白
            labelAreaH = 5 + ((labelLines - 1) * lineGap) + 6;
        }

        // 各ブロックのスウォッチ高さをアートボード内に収める
        var maxSwH = 50;
        var swH = Math.floor((abH - (nBlocks * headerH) - (nBlocks * labelAreaH) - ((nBlocks - 1) * gap)) / nBlocks);
        if (swH > maxSwH) swH = maxSwH;
        if (swH < 10) swH = 10;

        var yTop = top;

        for (var i = 0; i < nBlocks; i++) {
            var b = blocks[i];
            drawPaletteRow(doc, left, yTop, abW, headerH, swH, b.data, b.name, outHex, outRgb, contrastShift);
            yTop = yTop - (headerH + swH + labelAreaH + gap);
        }
    }

    // 共通描画：タイトル + スウォッチ矩形 + HEX/RGB ラベル（単体/すべて共通）
    // Shared renderer for both single output and "All"
    function drawPaletteRow(doc, left, top, width, headerH, swH, data, algoName, outHex, outRgb, contrastShift) {
        if (!data || !data.length) return;

        // --- グループ作成（矩形/テキスト） ---
        var rectGroup = doc.groupItems.add();
        var titleGroup = doc.groupItems.add();
        var hexGroup = doc.groupItems.add();
        var rgbGroup = doc.groupItems.add();

        // 前面順（RGB/HEX/Title を前に）
        try { titleGroup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e) { }
        try { rgbGroup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e) { }
        try { hexGroup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e) { }

        // タイトル / Title
        if (algoName && algoName.length && headerH > 0) {
            try {
                var tf = doc.textFrames.add();
                tf.contents = algoName;

                // position: [x, y] (yはベースライン)
                tf.position = [left, top - 8];

                // 左揃え
                try { tf.textRange.paragraphAttributes.justification = Justification.LEFT; } catch (e) { }

                // フォント設定：Avenir-Book / 11pt
                try {
                    tf.textRange.characterAttributes.textFont = app.textFonts.getByName("Avenir-Book");
                    tf.textRange.size = 11;
                } catch (e) { }

                // 念のため属性を正規化（最初の1個だけ崩れる対策）
                try {
                    tf.textRange.characterAttributes.horizontalScale = 100;
                    tf.textRange.characterAttributes.verticalScale = 100;
                    tf.textRange.characterAttributes.tracking = 0;
                    tf.textRange.characterAttributes.baselineShift = 0;
                } catch (e) { }

                // 塗り：黒
                try {
                    var tcol = new RGBColor();
                    tcol.red = 0; tcol.green = 0; tcol.blue = 0;
                    tf.textRange.fillColor = tcol;
                } catch (e) { }

                // 見た目の左端を left に合わせる（フォント/サイズ確定後に補正）
                try {
                    var gbT = tf.geometricBounds; // [left, top, right, bottom]
                    var dxT = left - gbT[0];
                    tf.position = [tf.position[0] + dxT, tf.position[1]];
                } catch (e) { }

                try { tf.move(titleGroup, ElementPlacement.PLACEATEND); } catch (e) { }
            } catch (e) { }
        }

        // テキスト共通ヘルパー / Text helpers
        function __setTextStyle(tf, sizePt, fontName) {
            try {
                tf.textRange.size = sizePt;
                var f = null;
                try { f = app.textFonts.getByName(fontName); } catch (e) { }
                if (!f) {
                    try { f = app.textFonts.getByName("Avenir-Book"); } catch (__) { }
                }
                if (f) tf.textRange.characterAttributes.textFont = f;
            } catch (e) { }
            try {
                var tcol = new RGBColor();
                tcol.red = 0; tcol.green = 0; tcol.blue = 0;
                tf.textRange.fillColor = tcol;
            } catch (e) { }
        }

        function __addLabel(text, x, y, sizePt, parentGroup) {
            try {
                var tf = doc.textFrames.add();
                tf.contents = text;
                __setTextStyle(tf, sizePt, "Avenir-Book");
                tf.position = [x, y];
                try { tf.move(parentGroup, ElementPlacement.PLACEATEND); } catch (e) { }
                return tf;
            } catch (e) { }
            return null;
        }

        function __alignLeftToRect(tf, targetX) {
            try {
                var gb = tf.geometricBounds; // [left, top, right, bottom]
                var dx = targetX - gb[0];
                tf.position = [tf.position[0] + dx, tf.position[1]];
            } catch (e) { }
        }

        // 矩形位置 / Rectangles
        var n = data.length;
        var swW = width / n;
        var rectTop = top - headerH;

        for (var i = 0; i < n; i++) {
            var c0 = data[i];
            var c = applyContrastShiftToRgb(c0, contrastShift);
            var col = new RGBColor();
            col.red = c.r; col.green = c.g; col.blue = c.b;

            var rect = rectGroup.pathItems.rectangle(
                rectTop,
                left + (i * swW),
                swW,
                swH
            );
            rect.fillColor = col;
            rect.stroked = false;

            // ラベル（HEX / RGB） / Labels
            if (outHex || outRgb) {
                var x0 = left + (i * swW);
                var yBot2 = rectTop - swH; // bottom edge

                // 2行（HEX値 / RGB値）
                // ※HEXがOFFの場合はRGBを上に詰める
                var lineGap = 7; // 5pt想定の行送り
                var yTopLine = yBot2 - 2;
                var yHexVal = yTopLine;
                var yRgbVal = (outHex ? (yTopLine - lineGap) : yTopLine);

                if (outHex) {
                    var hx = rgbToHexString(c.r, c.g, c.b);
                    var tHex = __addLabel(hx, x0, yHexVal, 5, hexGroup);
                    if (tHex) {
                        __setTextStyle(tHex, 5, "Avenir");
                        __alignLeftToRect(tHex, x0);
                    }
                }

                if (outRgb) {
                    var rgbText = "R" + Math.round(c.r) + " G" + Math.round(c.g) + " B" + Math.round(c.b);
                    var tRgb = __addLabel(rgbText, x0, yRgbVal, 5, rgbGroup);
                    if (tRgb) {
                        __setTextStyle(tRgb, 5, "Avenir-Book");
                        __alignLeftToRect(tRgb, x0);
                    }
                }
            }
        }
    }

    main();

})();
