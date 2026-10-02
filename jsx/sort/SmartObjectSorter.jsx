#target illustrator
#targetengine "SmartObjectSorterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを高さ・幅・不透明度・カラーなどの基準で並び替え、横または縦に整列・分布させます。
整列方向、基準、順序、間隔、幅・高さの統一を、プレビューを見ながら指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectSorter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n663264db75ff

### Overview

Sorts the selected objects by height, width, opacity or color and then aligns and distributes them horizontally or vertically.
Direction, sort key, order, spacing and size unification are all set while watching a preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectSorter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectSorter";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v0.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectSorter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectSorter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n663264db75ff"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「指定」間隔の初期値（pt） / Initial custom gap in points */
    var DEFAULT_CUSTOM_GAP = "20";

    /* 「数字」で並べるときの上端の送り量（pt） / Top-to-top step for the Number sort, in points */
    var NUMBER_SORT_STEP = 50;

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
            title: { ja: "オブジェクトの整列", en: "Object Alignment Tool" }
        },
        panel: {
            sortKey: { ja: "基準", en: "Sort by" },
            sortOrder: { ja: "ソート順", en: "Sort Order" },
            vertical: { ja: "縦方向", en: "Vertical" },
            alignVertical: { ja: "揃え", en: "Align Vertically" },
            spacingVertical: { ja: "縦間隔", en: "Vertical Spacing" },
            matchWidth: { ja: "幅を揃える", en: "Match Width" },
            horizontal: { ja: "横方向", en: "Horizontal" },
            alignHorizontal: { ja: "揃え", en: "Align Horizontally" },
            spacingHorizontal: { ja: "横間隔", en: "Horizontal Spacing" },
            matchHeight: { ja: "高さを揃える", en: "Match Height" }
        },
        radio: {
            alongX: { ja: "横並び", en: "Horizontal" },
            alongY: { ja: "縦並び", en: "Vertical" },
            byHeight: { ja: "高さ", en: "Height" },
            byWidth: { ja: "幅", en: "Width" },
            byOpacity: { ja: "不透明度", en: "Opacity" },
            byColor: { ja: "カラー", en: "Color" },
            byNumber: { ja: "数字", en: "Number" },
            byZOrder: { ja: "重ね順", en: "Z-Order" },
            ascending: { ja: "昇順", en: "Ascending" },
            descending: { ja: "降順", en: "Descending" },
            random: { ja: "ランダム", en: "Random" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignTop: { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" },
            spacingEven: { ja: "均等", en: "Even" },
            spacingTight: { ja: "ぴったり", en: "Tight" },
            spacingCustom: { ja: "指定", en: "Custom" },
            matchMax: { ja: "最大", en: "Max" },
            matchMin: { ja: "最小", en: "Min" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
        },
        tooltip: {
            alongX: {
                ja: "基準とソート順に従って、オブジェクトどうしの横位置を入れ替えます",
                en: "Swaps the objects' horizontal positions to follow the sort key and order"
            },
            alongY: {
                ja: "基準とソート順に従って、オブジェクトどうしの縦位置を入れ替えます",
                en: "Swaps the objects' vertical positions to follow the sort key and order"
            },
            byColor: {
                ja: "塗りのカラーをグレースケールに換算した値で並べ替えます",
                en: "Sorts by the fill color converted to a grayscale value"
            },
            byNumber: {
                ja: "数字だけのテキストを含むグループを数値の順に並べ、上端を50 ptずつずらして縦に配置します",
                en: "Orders groups that contain a digits-only text by that number and stacks them, each top 50 pt below the previous one"
            },
            spacingEven: {
                ja: "両端のオブジェクトの位置はそのままで、間隔を均等にします",
                en: "Evens out the gaps while the objects at both ends stay put"
            },
            spacingTight: { ja: "間隔を0にして詰めます", en: "Closes the gaps between objects" },
            customGap: { ja: "オブジェクトの間隔（pt）", en: "Gap between objects, in points" },
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
            matchMaxWidth: { ja: "いちばん広い幅に合わせて、幅だけを変えます", en: "Scales only the width to match the widest object" },
            matchMinWidth: { ja: "いちばん狭い幅に合わせて、幅だけを変えます", en: "Scales only the width to match the narrowest object" },
            matchMaxHeight: { ja: "いちばん高い高さに合わせて、高さだけを変えます", en: "Scales only the height to match the tallest object" },
            matchMinHeight: { ja: "いちばん低い高さに合わせて、高さだけを変えます", en: "Scales only the height to match the shortest object" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "実行", en: "Apply" }
        },
        alert: {
            noSelection: { ja: "ファイルを開き、並べ替えるオブジェクトを選択してください。", en: "Open a file and select objects to sort." }
        }
    };

    // =========================================
    // 並べ替え / Sorting
    // =========================================

    /* 基準・並び方向のキーと、比べる／入れ替えるプロパティ名の対応
       Option key -> property that is compared or reassigned */
    var OPTION_PROPERTIES = {
        h: "height",
        w: "width",
        o: "opacity",
        color: "color",
        z: "zOrderPosition",
        x: "left",
        y: "top"
    };

    /**
     * 数値を昇順に比べる
     * @param {number} a - 比べる値
     * @param {number} b - 比べる値
     * @returns {number} 並べ替え用の差
     */
    function compareAscending(a, b) {
        return a - b;
    }

    /**
     * 数値を降順に比べる
     * @param {number} a - 比べる値
     * @param {number} b - 比べる値
     * @returns {number} 並べ替え用の差
     */
    function compareDescending(a, b) {
        return b - a;
    }

    /**
     * 並べ替えの順番をランダムにする
     * @returns {number} -0.5〜0.5 の乱数
     */
    function compareRandomly() {
        return Math.random() - .5;
    }

    /**
     * カラーを色成分の配列にする
     * @param {Color} sourceColor - 対象のカラー
     * @returns {number[]} CMYK は4つ、RGB は3つ、グレーは1つの成分（その他は [0]）
     */
    function getColorChannels(sourceColor) {
        var color = sourceColor;
        if (color.hasOwnProperty('color')) color = color.color;
        if (color.constructor.name == 'SpotColor') color = color.spot.color;

        if (color.constructor.name === 'CMYKColor') return [color.cyan, color.magenta, color.yellow, color.black];
        if (color.constructor.name === 'RGBColor') return [color.red, color.green, color.blue];
        if (color.constructor.name === 'GrayColor') return [color.gray];
        return [0];
    }

    /**
     * カラーをグレースケールに換算する
     * @param {Color} sourceColor - 対象のカラー
     * @returns {number[]} グレースケールの値（要素1つの配列）
     */
    function getGrayScaleValue(sourceColor) {
        var channels = getColorChannels(sourceColor);
        if (channels.length === 4) {
            return app.convertSampleColor(ImageColorSpace.CMYK, channels, ImageColorSpace.GrayScale, ColorConvertPurpose.defaultpurpose);
        }
        if (channels.length === 3) {
            return app.convertSampleColor(ImageColorSpace.RGB, channels, ImageColorSpace.GrayScale, ColorConvertPurpose.defaultpurpose);
        }
        return channels;
    }

    /**
     * 「カラー」基準で比べる値（塗りのグレースケール値）を返す
     * @param {PageItem} item - 対象オブジェクト
     * @returns {number[]|number} グレースケールの値。塗りが無ければ 0
     */
    function getColorSortValue(item) {
        return item.fillColor ? getGrayScaleValue(item.fillColor) : 0;
    }

    /**
     * 指定プロパティでオブジェクトを比べる関数を作る
     * @param {string} propertyName - 比べるプロパティ名（"color" は塗りの明るさ）
     * @returns {Function} Array.sort 用の比較関数
     */
    function createItemComparator(propertyName) {
        return function (a, b) {
            if (propertyName === "color") {
                return getColorSortValue(a) - getColorSortValue(b);
            }
            return Number(a[propertyName]) - Number(b[propertyName]);
        };
    }

    /**
     * オブジェクトを基準の順に並べ、その順に位置の値を割り当て直す
     * @param {PageItem[]} items - 対象オブジェクト（この配列自体も並べ替える）
     * @param {Function} compareItems - オブジェクトの比較関数
     * @param {Function} comparePositions - 位置の値の比較関数（昇順／降順／ランダム）
     * @param {string} positionProperty - 割り当て直すプロパティ（"left" / "top"）
     * @returns {void}
     */
    function rearrangeItems(items, compareItems, comparePositions, positionProperty) {
        items.sort(compareItems);
        var positions = [];
        for (var i = 0; i < items.length; i++) {
            positions.push(items[i][positionProperty]);
        }
        positions.sort(comparePositions);
        /* 上端は値が大きいほど上なので逆順に / larger top means higher, so reverse */
        if (positionProperty === "top") positions.reverse();
        for (var j = 0; j < items.length; j++) {
            items[j][positionProperty] = positions[j];
        }
    }

    /**
     * グループ内で最初に見つかった数字だけのテキストを数値で返す（入れ子のグループも探す）
     * @param {GroupItem} groupItem - 対象グループ
     * @returns {number} 見つかった数値。無ければ NaN
     */
    function findNumberInGroup(groupItem) {
        var memberItems = groupItem.pageItems;
        for (var i = 0; i < memberItems.length; i++) {
            var memberItem = memberItems[i];
            if (memberItem.typename === "TextFrame") {
                var frameText = memberItem.contents;
                if (/^\d+$/.test(frameText)) {
                    return parseFloat(frameText);
                }
            } else if (memberItem.typename === "GroupItem") {
                var nestedNumber = findNumberInGroup(memberItem);
                if (!isNaN(nestedNumber)) return nestedNumber;
            }
        }
        return NaN;
    }

    /**
     * 数字を含むグループを数値の順に並べ、先頭のグループの位置から縦に配置する
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} orderKey - ソート順のキー（"l" は降順、それ以外は昇順）
     * @returns {void}
     */
    function arrangeGroupsByNumber(items, orderKey) {
        var numberedGroups = [];
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === "GroupItem") {
                var groupNumber = findNumberInGroup(items[i]);
                if (!isNaN(groupNumber)) {
                    numberedGroups.push({ group: items[i], value: groupNumber });
                }
            }
        }
        if (numberedGroups.length === 0) return;
        if (orderKey === "l") {
            numberedGroups.sort(function (a, b) { return b.value - a.value; });
        } else {
            numberedGroups.sort(function (a, b) { return a.value - b.value; });
        }
        var startTop = numberedGroups[0].group.top;
        var startLeft = numberedGroups[0].group.left;
        for (var j = 0; j < numberedGroups.length; j++) {
            numberedGroups[j].group.left = startLeft;
            numberedGroups[j].group.top = startTop - j * NUMBER_SORT_STEP;
        }
    }

    /**
     * 基準・並び方向・ソート順に従ってオブジェクトを並べ替える
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} sortKey - 基準のキー（h / w / o / color / n / z）
     * @param {string} alongKey - 並び方向のキー（x / y）
     * @param {string} orderKey - ソート順のキー（s / l / r）
     * @returns {void}
     */
    function applyArrangement(items, sortKey, alongKey, orderKey) {
        if (sortKey === "n") {
            arrangeGroupsByNumber(items, orderKey);
            return;
        }
        var comparePositions = compareAscending;
        if (orderKey === "l") {
            comparePositions = compareDescending;
        } else if (orderKey === "r") {
            comparePositions = compareRandomly;
        }
        rearrangeItems(items, createItemComparator(OPTION_PROPERTIES[sortKey]), comparePositions, OPTION_PROPERTIES[alongKey]);
    }

    // =========================================
    // 間隔と揃え / Spacing and alignment
    // =========================================

    /**
     * 両端の位置を保ったまま均等にするときの間隔を求める
     * @param {PageItem[]} sortedItems - 並び順に並べたオブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {number} オブジェクト間の間隔（pt）
     */
    function getEvenGap(sortedItems, isHorizontal) {
        var firstItem = sortedItems[0];
        var lastItem = sortedItems[sortedItems.length - 1];
        var totalSize = 0;
        for (var i = 0; i < sortedItems.length; i++) {
            totalSize += isHorizontal ? sortedItems[i].width : sortedItems[i].height;
        }
        var totalGap = isHorizontal
            ? (lastItem.left + lastItem.width) - firstItem.left - totalSize
            : firstItem.top - (lastItem.top - lastItem.height) - totalSize;
        return totalGap / (sortedItems.length - 1);
    }

    /**
     * 先頭のオブジェクトの位置から、指定の間隔で順に並べる
     * @param {PageItem[]} sortedItems - 並び順に並べたオブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @param {number} gap - オブジェクト間の間隔（pt）
     * @returns {void}
     */
    function placeWithGap(sortedItems, isHorizontal, gap) {
        var position = isHorizontal ? sortedItems[0].left : sortedItems[0].top;
        for (var i = 0; i < sortedItems.length; i++) {
            if (isHorizontal) {
                sortedItems[i].left = position;
                position += sortedItems[i].width + gap;
            } else {
                sortedItems[i].top = position;
                position -= sortedItems[i].height + gap;
            }
        }
    }

    /**
     * 間隔の種類に従ってオブジェクトを並べ直す
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} spacingType - "even"（均等）/ "zero"（ぴったり）/ "custom"（指定）
     * @param {EditText} gapInput - 「指定」の間隔の入力欄
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {void}
     */
    function applySpacing(items, spacingType, gapInput, isHorizontal) {
        var sortedItems = items.slice();
        sortedItems.sort(function (a, b) {
            return isHorizontal ? a.left - b.left : b.top - a.top;
        });
        var gap = 0;
        if (spacingType === "even") {
            gap = getEvenGap(sortedItems, isHorizontal);
        } else if (spacingType === "custom" && gapInput.text !== "") {
            gap = parseFloat(gapInput.text);
        }
        placeWithGap(sortedItems, isHorizontal, gap);
    }

    /**
     * 揃えの基準にする位置を、オブジェクトの左端または上端からのずれで返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {string} alignType - "left" / "center" / "right" / "top" / "middle" / "bottom"
     * @returns {number} 左端（上端）から基準位置までのずれ
     */
    function getAlignOffset(bounds, alignType) {
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        if (alignType === "center") return width / 2;
        if (alignType === "right") return width;
        if (alignType === "middle") return -(height / 2);
        if (alignType === "bottom") return -height;
        return 0;
    }

    /**
     * オブジェクトの左端・中央・右端（上端・中央・下端）を揃える
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} alignType - "left" / "center" / "right" / "top" / "middle" / "bottom"
     * @returns {void}
     */
    function alignItems(items, alignType) {
        if (!items || items.length === 0) return;
        var movesVertically = (alignType === "top" || alignType === "middle" || alignType === "bottom");

        /* 注意: previewBoundsCheckbox はダイアログ内のローカル変数で、ここからは見えないため常に false
           Note: previewBoundsCheckbox is local to the dialog, so this is always false here */
        var usePreviewBounds = (typeof previewBoundsCheckbox !== "undefined" && previewBoundsCheckbox.value);

        var itemBounds = [];
        var alignPositions = [];
        for (var i = 0; i < items.length; i++) {
            itemBounds[i] = usePreviewBounds ? items[i].visibleBounds : items[i].geometricBounds;
            var edgePosition = movesVertically ? itemBounds[i][1] : itemBounds[i][0];
            alignPositions.push(edgePosition + getAlignOffset(itemBounds[i], alignType));
        }

        var targetPosition;
        if (alignType === "top" || alignType === "left") {
            targetPosition = Math.min.apply(null, alignPositions);
        } else if (alignType === "bottom" || alignType === "right") {
            targetPosition = Math.max.apply(null, alignPositions);
        } else {
            var positionSum = 0;
            for (var j = 0; j < alignPositions.length; j++) {
                positionSum += alignPositions[j];
            }
            targetPosition = positionSum / alignPositions.length;
        }

        for (var k = 0; k < items.length; k++) {
            var alignOffset = getAlignOffset(itemBounds[k], alignType);
            if (movesVertically) {
                var dy = items[k].top - itemBounds[k][1];
                items[k].top = targetPosition - alignOffset + dy;
            } else {
                var dx = items[k].left - itemBounds[k][0];
                items[k].left = targetPosition - alignOffset + dx;
            }
        }
    }

    /**
     * オブジェクトの幅または高さを、いちばん大きい（小さい）ものに合わせて拡大・縮小する
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {string} sizeProperty - "width" / "height"
     * @param {boolean} useLargest - true なら最大、false なら最小に合わせる
     * @returns {void}
     */
    function matchItemSize(items, sizeProperty, useLargest) {
        var targetSize = useLargest ? 0 : Infinity;
        for (var i = 0; i < items.length; i++) {
            var itemSize = items[i][sizeProperty];
            if (useLargest ? itemSize > targetSize : itemSize < targetSize) {
                targetSize = itemSize;
            }
        }
        for (var j = 0; j < items.length; j++) {
            var scalePercent = targetSize / items[j][sizeProperty] * 100;
            if (sizeProperty === "width") {
                items[j].resize(scalePercent, 100);
            } else {
                items[j].resize(100, scalePercent);
            }
        }
        app.redraw();
    }

    // =========================================
    // 選択と位置 / Selection and positions
    // =========================================

    /**
     * 先頭3つのオブジェクトの中心から、横並びか縦並びかを推定する
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {string} 縦並びなら "y"、それ以外は "x"
     */
    function detectAlongDirection(items) {
        if (!(items.length >= 3)) return "x";
        var centers = [];
        for (var i = 0; i < 3; i++) {
            var visibleBounds = items[i].visibleBounds;
            centers.push({ x: (visibleBounds[0] + visibleBounds[2]) / 2, y: (visibleBounds[1] + visibleBounds[3]) / 2 });
        }
        var averageDx = (Math.abs(centers[0].x - centers[1].x) + Math.abs(centers[1].x - centers[2].x)) / 2;
        var averageDy = (Math.abs(centers[0].y - centers[1].y) + Math.abs(centers[1].y - centers[2].y)) / 2;
        return (averageDx * 1.5 < averageDy) ? "y" : "x";
    }

    /**
     * キャンセル時に戻せるよう、オブジェクトの位置を控える
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {Object[]} { item, left, top } の配列
     */
    function saveItemPositions(items) {
        var savedPositions = [];
        for (var i = 0; i < items.length; i++) {
            savedPositions.push({ item: items[i], left: items[i].left, top: items[i].top });
        }
        return savedPositions;
    }

    /**
     * 控えておいた位置へオブジェクトを戻す
     * @param {Object[]} savedPositions - saveItemPositions() の戻り値
     * @returns {void}
     */
    function restoreItemPositions(savedPositions) {
        for (var i = 0; i < savedPositions.length; i++) {
            savedPositions[i].item.left = savedPositions[i].left;
            savedPositions[i].item.top = savedPositions[i].top;
        }
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

    /**
     * 選択肢のラジオボタンを並べて作る（値は optionKey に持たせる）
     * @param {Object} parentGroup - 追加先のグループまたはパネル
     * @param {Object[]} choices - { key, label, tooltip } の配列（label と tooltip は LABELS のパス）
     * @param {string} selectedKey - 最初に選んでおくキー
     * @returns {RadioButton[]} 作ったラジオボタン
     */
    function addChoiceRadios(parentGroup, choices, selectedKey) {
        var choiceRadios = [];
        for (var i = 0; i < choices.length; i++) {
            var choiceRadio = parentGroup.add("radiobutton", undefined, getLabel(choices[i].label));
            choiceRadio.optionKey = choices[i].key;
            choiceRadio.value = (choices[i].key === selectedKey);
            if (choices[i].tooltip) choiceRadio.helpTip = getLabel(choices[i].tooltip);
            choiceRadios.push(choiceRadio);
        }
        return choiceRadios;
    }

    /**
     * 選ばれているラジオボタンのキーを返す
     * @param {RadioButton[]} choiceRadios - ラジオボタン
     * @returns {string|null} optionKey。どれも選ばれていなければ null
     */
    function getSelectedKey(choiceRadios) {
        for (var i = 0; i < choiceRadios.length; i++) {
            if (choiceRadios[i].value) return choiceRadios[i].optionKey;
        }
        return null;
    }

    /**
     * 基準・ソート順のパネルを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} titlePath - パネル名の LABELS パス
     * @returns {Panel} 作ったパネル
     */
    function addOptionPanel(parentGroup, titlePath) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        setupPanel(optionPanel, 6);
        return optionPanel;
    }

    /**
     * 選択肢を横一列に並べるパネルを作る（揃え・幅／高さを揃える）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @returns {Panel} 作ったパネル
     */
    function addRowPanel(parentPanel, titlePath) {
        var rowPanel = parentPanel.add("panel", undefined, getLabel(titlePath));
        setupPanel(rowPanel);
        rowPanel.orientation = "row"; /* 選択肢を横一列に並べる / choices in a single row */
        rowPanel.alignChildren = "center";
        return rowPanel;
    }

    /**
     * 揃えのパネルを作る。ラジオボタンを押すとすぐに揃える
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {Object[]} choices - { key, label } の配列（key は alignItems() の揃え方）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {void}
     */
    function addAlignPanel(parentPanel, titlePath, choices, targetItems) {
        var alignRadios = addChoiceRadios(addRowPanel(parentPanel, titlePath), choices, null);
        for (var i = 0; i < alignRadios.length; i++) {
            alignRadios[i].onClick = function () {
                alignItems(targetItems, this.optionKey);
                app.redraw();
            };
        }
    }

    /**
     * 幅（高さ）を揃えるパネルを作る。ラジオボタンを押すとすぐに拡大・縮小する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} sizeProperty - "width" / "height"
     * @returns {void}
     */
    function addMatchSizePanel(parentPanel, titlePath, targetItems, sizeProperty) {
        var matchesWidth = (sizeProperty === "width");
        var matchRadios = addChoiceRadios(addRowPanel(parentPanel, titlePath), [
            { key: "max", label: "radio.matchMax", tooltip: matchesWidth ? "tooltip.matchMaxWidth" : "tooltip.matchMaxHeight" },
            { key: "min", label: "radio.matchMin", tooltip: matchesWidth ? "tooltip.matchMinWidth" : "tooltip.matchMinHeight" }
        ], null);
        matchRadios[0].onClick = function () {
            matchItemSize(targetItems, sizeProperty, true);
        };
        matchRadios[1].onClick = function () {
            matchItemSize(targetItems, sizeProperty, false);
        };
    }

    /**
     * 間隔のパネル（均等／ぴったり／指定）を作る
     * 「指定」だけ別のグループに入るので、ラジオボタンの排他は手で管理する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {void}
     */
    function addSpacingPanel(parentPanel, titlePath, targetItems, isHorizontal) {
        var spacingPanel = parentPanel.add("panel", undefined, getLabel(titlePath));
        setupPanel(spacingPanel);

        var spacingRadioGroup = spacingPanel.add("group");
        spacingRadioGroup.orientation = "column";
        spacingRadioGroup.alignChildren = "left";
        var presetRadios = addChoiceRadios(spacingRadioGroup, [
            { key: "even", label: "radio.spacingEven", tooltip: "tooltip.spacingEven" },
            { key: "zero", label: "radio.spacingTight", tooltip: "tooltip.spacingTight" }
        ], null);

        var customGapGroup = spacingRadioGroup.add("group");
        customGapGroup.orientation = "row";
        customGapGroup.alignChildren = "left";
        var customRadios = addChoiceRadios(customGapGroup, [{ key: "custom", label: "radio.spacingCustom" }], null);
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var customGapStepperRow = customGapGroup.add("group");
        customGapStepperRow.orientation = "row";
        customGapStepperRow.alignChildren = ["left", "center"];
        customGapStepperRow.spacing = 0;
        customGapStepperRow.margins = 0;

        var customGapInput;
        /* 負の間隔（重ね）も使えるので下限は置かない。適用は既存の onChange に任せる / negative gaps allowed; reuse onChange */
        var customGapStepper = addStepper(customGapStepperRow, function () { return customGapInput; }, {
            onStep: function (numberInput) { if (numberInput.onChange) numberInput.onChange(); }
        });
        customGapInput = customGapStepperRow.add("edittext", undefined, DEFAULT_CUSTOM_GAP);
        customGapInput.characters = 5;
        customGapInput.enabled = false;
        customGapStepper.enabled = false;
        customGapInput.helpTip = getLabel("tooltip.customGap");
        bindSteppedArrowKeys(customGapInput, customGapStepper);

        var spacingRadios = presetRadios.concat(customRadios);
        for (var i = 0; i < spacingRadios.length; i++) {
            spacingRadios[i].onClick = function () {
                var spacingType = this.optionKey;
                for (var j = 0; j < spacingRadios.length; j++) {
                    spacingRadios[j].value = (spacingRadios[j].optionKey === spacingType);
                }
                customGapInput.enabled = (spacingType === "custom");
                customGapStepper.enabled = customGapInput.enabled;
                redrawSteppersIn(customGapStepper); /* ∧∨のディム表示を描き直す / redraw the stepper dimming */
                applySpacing(targetItems, spacingType, customGapInput, isHorizontal);
                app.redraw();
            };
        }
        customGapInput.onChange = function () {
            if (customRadios[0].value) {
                applySpacing(targetItems, "custom", customGapInput, isHorizontal);
                app.redraw();
            }
        };
    }

    /**
     * 縦方向・横方向のパネル（揃え・間隔・幅／高さを揃える）を作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean} isHorizontal - 横方向のパネルなら true
     * @returns {void}
     */
    function addDirectionPanel(parentGroup, targetItems, isHorizontal) {
        var directionPanel = parentGroup.add("panel", undefined, getLabel(isHorizontal ? "panel.horizontal" : "panel.vertical"));
        setupPanel(directionPanel);
        if (isHorizontal) {
            addAlignPanel(directionPanel, "panel.alignHorizontal", [
                { key: "top", label: "radio.alignTop" },
                { key: "middle", label: "radio.alignMiddle" },
                { key: "bottom", label: "radio.alignBottom" }
            ], targetItems);
            addSpacingPanel(directionPanel, "panel.spacingHorizontal", targetItems, true);
            addMatchSizePanel(directionPanel, "panel.matchHeight", targetItems, "height");
        } else {
            addAlignPanel(directionPanel, "panel.alignVertical", [
                { key: "left", label: "radio.alignLeft" },
                { key: "center", label: "radio.alignCenter" },
                { key: "right", label: "radio.alignRight" }
            ], targetItems);
            addSpacingPanel(directionPanel, "panel.spacingVertical", targetItems, false);
            addMatchSizePanel(directionPanel, "panel.matchWidth", targetItems, "width");
        }
    }

    /**
     * 並べ替え・整列のダイアログを開く。操作はその場でオブジェクトに反映し、キャンセルで位置を戻す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} defaultAlong - 最初に選んでおく並び方向（"x" / "y"）
     * @param {Object[]} savedPositions - キャンセル時に戻す位置
     * @returns {void}
     */
    function showSortDialog(targetItems, defaultAlong, savedPositions) {
        /* 縦並びなら基準を「幅」、横並びなら「高さ」から始める / start from Width when stacked vertically, Height otherwise */
        var defaultSortKey = (defaultAlong === "y") ? "w" : "h";

        var sortDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(sortDialog);

        var alongGroup = sortDialog.add("group");
        alongGroup.alignment = "center";
        alongGroup.orientation = "row";
        alongGroup.spacing = 5;
        alongGroup.margins = [10, 10, 10, 10];
        var alongRadios = addChoiceRadios(alongGroup, [
            { key: "x", label: "radio.alongX", tooltip: "tooltip.alongX" },
            { key: "y", label: "radio.alongY", tooltip: "tooltip.alongY" }
        ], defaultAlong);

        var columnsGroup = sortDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var sortColumn = columnsGroup.add("group");
        sortColumn.orientation = "column";
        sortColumn.alignChildren = "fill";
        sortColumn.spacing = 10;

        var sortKeyRadios = addChoiceRadios(addOptionPanel(sortColumn, "panel.sortKey"), [
            { key: "h", label: "radio.byHeight" },
            { key: "w", label: "radio.byWidth" },
            { key: "o", label: "radio.byOpacity" },
            { key: "color", label: "radio.byColor", tooltip: "tooltip.byColor" },
            { key: "n", label: "radio.byNumber", tooltip: "tooltip.byNumber" },
            { key: "z", label: "radio.byZOrder" }
        ], defaultSortKey);

        var sortOrderRadios = addChoiceRadios(addOptionPanel(sortColumn, "panel.sortOrder"), [
            { key: "s", label: "radio.ascending" },
            { key: "l", label: "radio.descending" },
            { key: "r", label: "radio.random" }
        ], "s");

        addDirectionPanel(columnsGroup, targetItems, false);
        addDirectionPanel(columnsGroup, targetItems, true);

        /* ボタン行（左：プレビュー境界、右：キャンセル・OK） / Button row (left: preview bounds, right: Cancel and OK) */
        var buttonRow = addButtonRow(sortDialog);

        var previewBoundsCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.value = false;

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnOK.active = true;

        /* 並び方向・基準・ソート順を変えたらすぐに並べ替える / rearrange as soon as an option changes */
        var arrangeRadios = alongRadios.concat(sortKeyRadios, sortOrderRadios);
        for (var i = 0; i < arrangeRadios.length; i++) {
            arrangeRadios[i].onClick = function () {
                var sortKey = getSelectedKey(sortKeyRadios);
                var alongKey = getSelectedKey(alongRadios);
                var orderKey = getSelectedKey(sortOrderRadios);
                if (!sortKey || !alongKey || orderKey === null) return;
                applyArrangement(targetItems, sortKey, alongKey, orderKey);
                app.redraw();
            };
        }

        btnOK.onClick = function () {
            sortDialog.close(1);
        };
        btnCancel.onClick = function () {
            restoreItemPositions(savedPositions);
            app.redraw();
            sortDialog.close(0);
        };
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(sortDialog, SCRIPT_NAME);
        sortDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        var docSelection = app.activeDocument.selection;
        if (docSelection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        var targetItems = [];
        for (var i = 0; i < docSelection.length; i++) {
            targetItems.push(docSelection[i]);
        }
        showSortDialog(targetItems, detectAlongDirection(docSelection), saveItemPositions(docSelection));
    }

    main();

})();
