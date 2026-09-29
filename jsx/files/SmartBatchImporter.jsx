#target illustrator
#targetengine "SmartBatchImporterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数の Illustrator ファイル（.ai / .svg / .eps）を一括で読み込み、ファイルまたはアートボードごとに1つのアートボードを作成します。
作成したアートボードは、全体が正方形に近くなるグリッドへ整列配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBatchImporter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8180588e5630

### Overview

Batch-imports several Illustrator files (.ai / .svg / .eps) and creates one artboard per file, or per source artboard.
The artboards are then arranged in a grid that comes out close to square.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBatchImporter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartBatchImporter";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBatchImporter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBatchImporter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8180588e5630"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var CONFIG = {
        spacingX: 100,                 // グループ間の横間隔（pt）/ Horizontal gap between groups (pt)
        spacingY: 100,                 // 行間の縦間隔（pt）/ Vertical gap between rows (pt)
        artboardMargin: 8.5,           // アートボードと内容の余白（pt）/ Margin around content within an artboard (pt)
        labelFont: "HiraginoSans-W3",  // ラベルのフォント / Font used for labels
        labelSize: 9,                  // ラベルの文字サイズ（pt）/ Label font size (pt)
        labelLayerName: "_label"       // ラベルを置くレイヤー名 / Name of the layer that holds labels
    };

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

    var LABELS = {
        dialog: {
            title: { ja: "ファイル一括読み込み", en: "Batch Import Files" }
        },
        panel: {
            source: { ja: "読み込み対象", en: "Source" },
            filter: { ja: "フィルター", en: "Filter" },
            destination: { ja: "読み込み先", en: "Destination" },
            colorMode: { ja: "カラーモード", en: "Color Mode" },
            resolution: { ja: "解像度", en: "Resolution" },
            docSize: { ja: "サイズ", en: "Size" },
            options: { ja: "読み込みオプション", en: "Import Options" },
            afterImport: { ja: "読み込み後", en: "After Import" }
        },
        radio: {
            openFiles: { ja: "現在、開いているファイル", en: "Currently open files" },
            specifyFolder: { ja: "フォルダーを指定", en: "Specify folder" },
            rgb: { ja: "RGB", en: "RGB" },
            cmyk: { ja: "CMYK", en: "CMYK" },
            currentDoc: { ja: "現在のドキュメント", en: "Current document" },
            newDoc: { ja: "新規ドキュメント", en: "New document" },
            artboardOne: { ja: "1のみ", en: "Artboard 1 only" },
            artboardAll: { ja: "すべて", en: "All" },
            artboardSpecify: { ja: "指定", en: "Specify" },
            closeDoc: { ja: "閉じる", en: "Close" },
            keepOpen: { ja: "開いたまま", en: "Keep Open" }
        },
        checkbox: {
            byArtboard: { ja: "アートボード単位", en: "Import per artboard" },
            includeGuides: { ja: "ガイド", en: "Guides" },
            attachLabel: { ja: "ファイル名ラベルを追加", en: "Add file-name labels" },
            scale: { ja: "拡大・縮小", en: "Scale" }
        },
        preset: {
            custom: { ja: "カスタム", en: "Custom" },
            a4: { ja: "A4：210 × 297 mm", en: "A4: 210 × 297 mm" },
            fullHD: { ja: "フルHD：1920 × 1080 px", en: "Full HD: 1920 × 1080 px" },
            largeCanvas: { ja: "ラージカンバス", en: "Large Canvas" }
        },
        field: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            fileType: { ja: "ファイル形式", en: "File format" },
            fileName: { ja: "ファイル名", en: "File name" },
            unit: { ja: "単位", en: "Unit" },
            artboardTarget: { ja: "対象アートボード", en: "Target artboards" }
        },
        hint: {
            filter: {
                ja: "正規表現を使ってファイルを絞り込み",
                en: "Filter files using a regular expression"
            }
        },
        button: {
            specify: { ja: "指定", en: "Choose" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        progress: {
            title: { ja: "ファイルを読み込み中...", en: "Importing files..." },
            count: { ja: "読み込み", en: "Imported" }
        },
        prompt: {
            selectFolder: {
                ja: "読み込むファイルが入ったフォルダーを選択してください",
                en: "Select a folder that contains files to import"
            }
        },
        alert: {
            noValidFile: {
                ja: "有効なファイルが見つかりませんでした。",
                en: "No valid files found."
            },
            noCurrentDoc: {
                ja: "「現在のドキュメント」に読み込むには、ドキュメントを開いておいてください。",
                en: "Open a document first to import into the current document."
            },
            invalidArtboardSpec: {
                ja: "対象アートボードの指定が正しくありません。番号で指定してください。例: 1, 3-5",
                en: "The target artboard specification is invalid. Specify by number, e.g. 1, 3-5"
            },
            noArtboardImported: {
                ja: "取り込める内容が見つかりませんでした。対象アートボードの番号やファイルの内容を確認してください。",
                en: "Nothing could be imported. Check the target artboard numbers and the file contents."
            },
            invalidNumber: {
                ja: "数値が正しくありません。",
                en: "Invalid numeric input."
            },
            pasteFail: {
                ja: "ペーストに失敗しました",
                en: "Paste failed"
            },
            cancelled: {
                ja: "読み込みを中断しました。ここまでに読み込んだ内容はドキュメントに残っています。",
                en: "Import was stopped. Items imported so far remain in the document."
            },
            invalidScale: {
                ja: "拡大・縮小の値が正しくありません。0より大きい数値を入力してください。",
                en: "The scale value is invalid. Enter a number greater than 0."
            },
            invalidFilter: {
                ja: "ファイル名フィルターの正規表現が正しくありません。式を見直すか、空欄にしてください。",
                en: "The file-name filter is not a valid regular expression. Fix it or clear the field."
            }
        },
        confirm: {
            discardUnsaved: {
                ja: "開いているファイルに未保存の変更がある場合、「読み込み後に閉じる」を選ぶと保存せずに閉じます。続行しますか？",
                en: "If any open files have unsaved changes, choosing \"Close After Import\" closes them without saving. Continue?"
            }
        },
        tooltip: {
            openFiles: {
                ja: "現在Illustratorで開いているドキュメントを読み込み対象にします。",
                en: "Uses the documents currently open in Illustrator as the import source."
            },
            specifyFolder: {
                ja: "指定したフォルダー内の .ai / .svg / .eps ファイルを読み込み対象にします。",
                en: "Uses .ai / .svg / .eps files in the chosen folder as the import source."
            },
            fileType: {
                ja: "フォルダー指定時に読み込むファイル形式を選びます。",
                en: "Choose which file types to import in folder mode."
            },
            closeDoc: {
                ja: "読み込み後に元ファイルを保存せずに閉じます。未保存の変更があるファイルでは変更が失われます。",
                en: "Closes source files after import without saving. Unsaved changes in those files will be lost."
            },
            keepOpen: {
                ja: "読み込み後も元ファイルを開いたままにします。アートボード単位読み込みで一時的に解除したロック・非表示状態は復元します。",
                en: "Keeps source files open after import. Lock/hidden states temporarily changed for per-artboard import are restored."
            },
            specify: {
                ja: "読み込むファイル（.ai / .svg / .eps）が入ったフォルダーを選びます。",
                en: "Choose a folder that contains the files to import (.ai / .svg / .eps)."
            },
            currentDoc: {
                ja: "現在開いているドキュメントに読み込みます。下のカラーモード・解像度・サイズ設定は使いません。",
                en: "Imports into the currently active document. The color mode / resolution / size settings below are not used."
            },
            newDoc: {
                ja: "下の設定で新規ドキュメントを作成し、そこに読み込みます。",
                en: "Creates a new document with the settings below and imports into it."
            },
            colorMode: {
                ja: "新規ドキュメントのカラーモードを選びます。読み込み元の色を完全に変換する機能ではありません。",
                en: "Choose the color mode for the new document. This does not fully convert all colors from source files."
            },
            byArtboard: {
                ja: "各アートボードを別々に読み込み、元のアートボードサイズと相対位置を保ちます。ロック・非表示オブジェクトも対象です。",
                en: "Imports each artboard separately and preserves the original artboard size and relative positions. Locked/hidden objects are included."
            },
            artboardTarget: {
                ja: "取り込むアートボードを選びます。「1のみ」「すべて」「指定」から選択します。",
                en: "Choose which artboards to import: only the first, all, or a specified set."
            },
            artboardSpecify: {
                ja: "取り込むアートボードを番号で指定します。例: 1, 3-5（カンマ区切り・範囲指定可）。",
                en: "Specify artboards to import by number, e.g. 1, 3-5 (comma-separated, ranges allowed)."
            },
            includeGuides: {
                ja: "ルーラーガイド（カンバスの半分以上に伸びる長いガイド）を除いて、ガイドを読み込みます。",
                en: "Imports guides, excluding ruler guides (long guides that span half the canvas or more)."
            },
            attachLabel: {
                ja: "読み込んだアートボードの下に、元ファイル名のラベルを追加します。",
                en: "Adds a source file-name label below each imported artboard."
            },
            scale: {
                ja: "読み込む内容とアートボード枠を指定％で拡大・縮小します。線幅も同率で変わります。",
                en: "Scales imported content and artboard cells by the specified percentage. Stroke widths scale by the same ratio."
            },
            size: {
                ja: "新規ドキュメントの初期サイズです。幅・高さの単位は選択したプリセットに連動します（カスタムは px）。",
                en: "Initial size of the new document. Width/height units follow the selected preset (px for Custom)."
            },
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
    // 単位 / Units
    // =========================================

    // 1mm あたりのポイント数 / Points per millimeter
    var MM_TO_PT = 2.8346;

    // 各単位 → ポイントの換算係数 / Conversion factor from each unit to points
    var UNIT_TO_PT = {
        px: 0.75,     // 96dpi: 1px = 0.75pt
        mm: MM_TO_PT, // 1mm = 2.8346pt
        inch: 72      // 1inch = 72pt（ラージカンバスのネイティブ単位、換算用）
    };

    // 単位ドロップダウンの並び / Units shown in the unit dropdown
    var UNIT_LIST = ["mm", "px"];

    // ドキュメントサイズプリセット（値はそれぞれのネイティブ単位）/ Document size presets (values in their native unit)
    // インデックスは presetDropdown の並びと一致。カスタムは unit のみ（値は自動入力しない）。
    var SIZE_PRESETS = [
        { unit: "px" },                             // カスタム / Custom
        { unit: "mm", width: 210, height: 297 },    // A4
        { unit: "px", width: 1920, height: 1080 },  // フルHD / Full HD
        { unit: "inch", width: 2270, height: 2270 } // ラージカンバス / Large Canvas
    ];

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行に「∧∨・入力欄」を隙間0で突き合わせて追加する。↑↓キーも∧∨と同じ処理で増減する
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - addStepper() に渡す step / min / max / integer / unit / onStep
     * @returns {EditText} 入力欄（∧∨は .stepperGroup、包む group は .parent）
     */
    function addSteppedInput(parentRow, initialText, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /* ラベルを指定位置（アートボード下端の左下）に配置する / Place the label at the given bottom-left point (just below the artboard) */
    function positionLabelBelow(label, left, bottomY) {
        label.left = left;
        label.top = bottomY - 4; // アートボード下端のすぐ下 / Just below the artboard's bottom edge
    }

    /* "_label" レイヤーを取得し、無ければ作成する / Get the "_label" layer, creating it if missing */
    function getOrCreateLabelLayer(doc) {
        var layer;
        try {
            layer = doc.layers.getByName(CONFIG.labelLayerName);
        } catch (e) {
            layer = doc.layers.add();
            layer.name = CONFIG.labelLayerName;
        }
        return layer;
    }

    /* レイヤーとそのサブレイヤーのロック・非表示を再帰的に解除する / Recursively unlock and reveal a layer and its sublayers */
    function unlockLayerTree(layer) {
        layer.locked = false;
        layer.visible = true;
        for (var i = 0; i < layer.layers.length; i++) {
            unlockLayerTree(layer.layers[i]);
        }
    }

    /* ドキュメント内のすべてのレイヤー・オブジェクトのロックと非表示を解除する / Unlock and reveal every layer and object in the document */
    function unlockAllLayersAndItems(doc) {
        for (var li = 0; li < doc.layers.length; li++) {
            unlockLayerTree(doc.layers[li]);
        }
        for (var pi = 0; pi < doc.pageItems.length; pi++) {
            doc.pageItems[pi].locked = false;
            doc.pageItems[pi].hidden = false;
        }
    }

    /* レイヤー・オブジェクトのロック/非表示状態を記録する（後で元に戻すため）
       / Capture the lock/hidden state of layers and items so it can be restored later */
    function captureLockHiddenState(doc) {
        var layers = [];
        var items = [];
        function walkLayers(layerList) {
            for (var i = 0; i < layerList.length; i++) {
                var layer = layerList[i];
                layers.push({ ref: layer, locked: layer.locked, visible: layer.visible });
                walkLayers(layer.layers);
            }
        }
        walkLayers(doc.layers);
        for (var pi = 0; pi < doc.pageItems.length; pi++) {
            var pageItem = doc.pageItems[pi];
            items.push({ ref: pageItem, locked: pageItem.locked, hidden: pageItem.hidden });
        }
        return { layers: layers, items: items };
    }

    /* 記録したロック/非表示状態を元に戻す。1件失敗しても残りは復元を続ける。
       / Restore a previously captured lock/hidden state. A failure on one item won't stop the rest. */
    function restoreLockHiddenState(state) {
        for (var i = 0; i < state.layers.length; i++) {
            try { state.layers[i].ref.locked = state.layers[i].locked; } catch (eLayerLock) { }
            try { state.layers[i].ref.visible = state.layers[i].visible; } catch (eLayerVis) { }
        }
        for (var j = 0; j < state.items.length; j++) {
            try { state.items[j].ref.locked = state.items[j].locked; } catch (eItemLock) { }
            try { state.items[j].ref.hidden = state.items[j].hidden; } catch (eItemHidden) { }
        }
    }

    /* 2つの矩形（[L, T, R, B]、T>B）が重なるか / Whether two [L, T, R, B] rects (T > B) overlap */
    function rectsIntersect(rectA, rectB) {
        return !(rectA[2] < rectB[0] || rectA[0] > rectB[2] || rectA[3] > rectB[1] || rectA[1] < rectB[3]);
    }

    /* 指定アートボード（複数可）に重なるオブジェクトだけロック・非表示を解除する（他アートボードのオブジェクトには触れない）。
       レイヤーは選択可能にするため一時的に全解除する。戻り値は restoreLockHiddenState で復元できる状態。
       / Unlock only the objects overlapping the given artboard rects (leaving other artboards' objects untouched).
       Layers are unlocked/shown temporarily so the items become selectable. The result restores via restoreLockHiddenState. */
    function unlockItemsOnArtboards(doc, abRects) {
        var layers = [];
        function walkLayers(layerList) {
            for (var i = 0; i < layerList.length; i++) {
                var layer = layerList[i];
                layers.push({ ref: layer, locked: layer.locked, visible: layer.visible });
                layer.locked = false;
                layer.visible = true;
                walkLayers(layer.layers);
            }
        }
        walkLayers(doc.layers);

        var items = [];
        for (var pi = 0; pi < doc.pageItems.length; pi++) {
            var pageItem = doc.pageItems[pi];
            if (!pageItem.locked && !pageItem.hidden) continue; // 既に選択可能なものは触らない / Leave already-selectable items alone
            var bounds;
            try { bounds = pageItem.geometricBounds; } catch (eBounds) { continue; }
            var overlapsTarget = false;
            for (var ri = 0; ri < abRects.length; ri++) {
                if (rectsIntersect(bounds, abRects[ri])) { overlapsTarget = true; break; }
            }
            if (!overlapsTarget) continue; // 対象アートボードのどれにも重ならないものは対象外 / Skip items outside all target artboards
            items.push({ ref: pageItem, locked: pageItem.locked, hidden: pageItem.hidden });
            pageItem.locked = false;
            pageItem.hidden = false;
        }
        return { layers: layers, items: items };
    }

    /* "1, 3-5" のような指定文字列を 1 始まりのアートボード番号の配列にする（カンマ区切り・範囲対応）。
       / Parse a spec like "1, 3-5" into an array of 1-based artboard numbers (comma-separated, ranges supported). */
    function parseArtboardNumbers(spec) {
        var numbers = [];
        if (!spec) return numbers;
        var parts = spec.split(/[,，\s]+/);
        for (var i = 0; i < parts.length; i++) {
            var part = parts[i];
            if (part === "") continue;
            var range = part.match(/^(\d+)\s*[-–~]\s*(\d+)$/);
            if (range) {
                var from = parseInt(range[1], 10);
                var to = parseInt(range[2], 10);
                if (from > to) { var swap = from; from = to; to = swap; }
                for (var n = from; n <= to; n++) numbers.push(n);
            } else if (/^\d+$/.test(part)) {
                numbers.push(parseInt(part, 10));
            }
        }
        return numbers;
    }

    /* 対象モードと指定文字列から、取り込むアートボードの 0 始まりインデックス配列を返す。
       / Resolve the 0-based artboard indices to import from the mode and spec text. */
    function resolveTargetArtboardIndices(doc, mode, specText) {
        var count = doc.artboards.length;
        var indices = [];
        if (mode === "all") {
            for (var i = 0; i < count; i++) indices.push(i);
        } else if (mode === "specify") {
            var numbers = parseArtboardNumbers(specText);
            var seen = {};
            for (var k = 0; k < numbers.length; k++) {
                var idx = numbers[k] - 1; // 1 始まり → 0 始まり / 1-based → 0-based
                if (idx >= 0 && idx < count && !seen[idx]) { seen[idx] = true; indices.push(idx); }
            }
            indices.sort(function (a, b) { return a - b; });
        } else { // "first"
            if (count > 0) indices.push(0);
        }
        return indices;
    }

    /* 取り込み対象のガイドを集める。長さがカンバスの半分以上のものは無視する。
       ロック・非表示のガイドは選択できないため除外。withinRect を渡すとその矩形に重なるものだけにする。
       / Collect importable guides, ignoring ones as long as half the canvas or more.
       Locked/hidden guides can't be selected, so they're excluded. With withinRect, keep only overlapping ones. */
    function collectImportableGuides(doc, canvasWidth, canvasHeight, withinRect) {
        var result = [];
        var halfWidth = canvasWidth / 2;
        var halfHeight = canvasHeight / 2;
        for (var i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            if (!item.guides || item.locked || item.hidden) continue;
            var guideBounds = item.geometricBounds; // [L, T, R, B]
            if ((guideBounds[2] - guideBounds[0]) >= halfWidth || (guideBounds[1] - guideBounds[3]) >= halfHeight) continue; // 長すぎるガイドは無視
            if (withinRect && !rectsIntersect(guideBounds, withinRect)) continue;
            result.push(item);
        }
        return result;
    }

    /* ドキュメント内の全アートボードを囲う合成矩形 [L, T, R, B] を返す / Union rect [L, T, R, B] of all artboards in the doc */
    function getArtboardsUnionRect(doc) {
        var firstRect = doc.artboards[0].artboardRect; // [L, T, R, B]
        var left = firstRect[0], top = firstRect[1], right = firstRect[2], bottom = firstRect[3];
        for (var i = 1; i < doc.artboards.length; i++) {
            var rect = doc.artboards[i].artboardRect;
            if (rect[0] < left) left = rect[0];
            if (rect[1] > top) top = rect[1];
            if (rect[2] > right) right = rect[2];
            if (rect[3] < bottom) bottom = rect[3];
        }
        return [left, top, right, bottom];
    }

    /* 貼り付け直後のオブジェクトを1グループにまとめてセルとして記録し、必要ならラベルを付ける。
       ctx.createArtboards が true のときだけセルに合わせたアートボードを追加する（OFF時は枠を作らず内容だけ）。
       / Group the just-pasted objects into one cell and optionally add a label.
       Adds a matching artboard only when ctx.createArtboards is true (otherwise the content is placed without a frame).
       artboardCell を渡すと元のアートボードサイズと内容の相対位置を保持する。
       / When artboardCell is provided, the original artboard size and the content's relative position are preserved. */
    function placePastedGroup(ctx, labelName, artboardCell) {
        var newDoc = ctx.newDoc;
        var pastedItems = newDoc.selection;
        if (pastedItems.length === 0) {
            alert(labelValueText('alert.pasteFail', labelName));
            return false;
        }

        var targetLayer = newDoc.activeLayer;
        var pastedGroup = targetLayer.groupItems.add();
        for (var m = pastedItems.length - 1; m >= 0; m--) {
            pastedItems[m].moveToBeginning(pastedGroup);
        }
        newDoc.activeLayer = targetLayer; // 貼り付け直後にアクティブレイヤーを戻す / Restore the active layer right after pasting

        // スケール適用（コンテンツとアートボード枠の両方を同率で）。線幅・パターン・グラデーションも同率で縮める。
        // Apply scaling to both the content and the artboard cell at the same rate; also scale stroke widths / patterns / gradients.
        if (ctx.scalePercent !== 100) {
            pastedGroup.resize(
                ctx.scalePercent, ctx.scalePercent,
                true,             // changePositions
                true,             // changeFillPatterns
                true,             // changeFillGradients
                true,             // changeStrokePattern
                ctx.scalePercent  // changeLineWidths（線幅も同率で / scale stroke widths by the same percent）
            );
        }
        var scaleFactor = ctx.scalePercent / 100;

        /* 貼り付けた各オブジェクトの外接範囲。クリップグループはマスクで測る（元の側と同じ測り方）
           The union of the pasted items; clip groups by their mask, the same way as the source side */
        var bounds = getClipAwareUnionBounds(pastedGroup.pageItems, true) || pastedGroup.visibleBounds;
        var contentWidth = bounds[2] - bounds[0];
        var contentHeight = bounds[1] - bounds[3];

        // セル（＝追加するアートボードの枠）の寸法と、その中での内容の位置
        // Cell (= the artboard to add) size and the content's position within it
        var cellPadding = ctx.artboardPadding;
        var cellWidth, cellHeight, offsetX, offsetY;
        if (artboardCell) {
            // 元のアートボードサイズと相対位置を保持（スケールも反映）/ Preserve the original artboard size and relative position (scaled)
            cellWidth = artboardCell.width * scaleFactor;
            cellHeight = artboardCell.height * scaleFactor;
            offsetX = artboardCell.offsetX * scaleFactor;
            offsetY = artboardCell.offsetY * scaleFactor;
        } else {
            // 内容に余白を足した枠 / A box fitted to the content plus padding
            cellWidth = contentWidth + cellPadding * 2;
            cellHeight = contentHeight + cellPadding * 2;
            offsetX = cellPadding;
            offsetY = cellPadding;
        }

        // 内容の現在位置を基準にしたセルの左上（最終位置は後でグリッド配置）
        // Cell's top-left based on the content's current spot (final position is set later by the grid layout)
        var abLeft = bounds[0] - offsetX;
        var abTop = bounds[1] + offsetY;

        // アートボードを作るのは「アートボード単位」ONのときだけ。OFFのときは枠を作らず内容だけ配置する。
        // Create an artboard only when "per artboard" is on; when off, place the content without adding a frame.
        var newArtboard = null;
        if (ctx.createArtboards) {
            newArtboard = newDoc.artboards.add([abLeft, abTop, abLeft + cellWidth, abTop - cellHeight]);
            newArtboard.name = labelName;
        }

        var labelItem = null;
        if (ctx.showLabel) {
            var labelLayer = getOrCreateLabelLayer(newDoc);
            labelItem = newDoc.textFrames.add();
            labelItem.contents = labelName;
            // フォントが見つからない環境では既定フォントのまま / Keep the default font when the configured one is missing
            try {
                labelItem.textRange.characterAttributes.textFont = app.textFonts.getByName(CONFIG.labelFont);
            } catch (fontError) { }
            labelItem.textRange.characterAttributes.size = CONFIG.labelSize;
            // ラベルはアートボードの下端の左下に置く / Place the label below the artboard's bottom-left
            positionLabelBelow(labelItem, abLeft, abTop - cellHeight);
            if (labelItem.layer != labelLayer) labelItem.layer = labelLayer;
            labelItem.move(labelLayer, ElementPlacement.PLACEATBEGINNING);
        }

        // セルを記録（グループ・アートボード・ラベルを最終グリッド配置で一緒に動かす）。
        // アートボードを作らない場合のために、現在のセル左上（currentLeft/currentTop）も保持する。
        // Record the cell so the group, artboard, and label move together in the final grid layout.
        // Also keep the current cell top-left (currentLeft/currentTop) for the no-artboard case.
        ctx.cells.push({
            group: pastedGroup,
            artboard: newArtboard, // OFF時は null / null when no artboard is created
            label: labelItem,
            width: cellWidth,
            height: cellHeight,
            currentLeft: abLeft,
            currentTop: abTop
        });
        ctx.placedCount++;
        return true;
    }

    /* 記録したセルを「全体が正方形に近い」グリッドに並べ、キャンバス中央に配置する。
       セルごとに group・artboard・label を同じ移動量でまとめて動かす。
       / Lay out the recorded cells in a grid whose overall shape is as square as possible, centered on the canvas.
       For each cell, the group, artboard, and label move together by the same delta. */
    function layoutCellsAsCenteredGrid(ctx) {
        var cells = ctx.cells;
        var count = cells.length;
        if (count === 0) return;

        // スロットサイズ＝最大セル寸法（サイズ混在でも重ならないように）/ Slot size = the largest cell (so mixed sizes never overlap)
        var slotWidth = 0;
        var slotHeight = 0;
        for (var i = 0; i < count; i++) {
            if (cells[i].width > slotWidth) slotWidth = cells[i].width;
            if (cells[i].height > slotHeight) slotHeight = cells[i].height;
        }

        // セル間の間隔。アートボードモードは「スロット幅の1/8」（縦も同値）、それ以外は CONFIG の既定値。
        // Gap between cells. In artboard mode use 1/8 of the slot width (same for vertical); otherwise the CONFIG default.
        var gapX = ctx.byArtboard ? (slotWidth / 8) : CONFIG.spacingX;
        var gapY = ctx.byArtboard ? (slotWidth / 8) : CONFIG.spacingY;

        // 全体のバウンディングボックスが最も正方形に近くなる列数を選ぶ
        // Choose the column count that makes the overall bounding box closest to square
        var columns = 1;
        var bestDiff = -1;
        for (var c = 1; c <= count; c++) {
            var r = Math.ceil(count / c);
            var totalWidthCandidate = c * slotWidth + (c - 1) * gapX;
            var totalHeightCandidate = r * slotHeight + (r - 1) * gapY;
            var diff = Math.abs(totalWidthCandidate - totalHeightCandidate);
            if (bestDiff < 0 || diff < bestDiff) {
                bestDiff = diff;
                columns = c;
            }
        }
        var rows = Math.ceil(count / columns);

        var totalWidth = columns * slotWidth + (columns - 1) * gapX;
        var totalHeight = rows * slotHeight + (rows - 1) * gapY;

        var startLeft, startTop;
        if (ctx.avoidRect) {
            // 現在のドキュメント：既存アートボードの下に、重ならないよう配置（水平は既存の中央に揃える）
            // Current document: place below the existing artboards (no overlap), centered horizontally on them
            var existingRect = ctx.avoidRect; // [L, T, R, B]
            var existingCenterX = (existingRect[0] + existingRect[2]) / 2;
            startLeft = existingCenterX - totalWidth / 2;
            startTop = existingRect[3] - gapY; // 既存アートボードの下端からひと間隔あけた位置をグリッド上端に / Grid top sits one gap below the existing bottom
        } else {
            // グリッド全体をキャンバス中央に揃える / Center the whole grid on the canvas
            startLeft = ctx.canvasCenterX - totalWidth / 2;
            startTop = ctx.canvasCenterY + totalHeight / 2;
        }

        for (var k = 0; k < count; k++) {
            var col = k % columns;
            var row = Math.floor(k / columns);
            var slotLeft = startLeft + col * (slotWidth + gapX);
            var slotTop = startTop - row * (slotHeight + gapY);

            var cell = cells[k];
            // セルをスロット内で中央に / Center the cell within its slot
            var targetLeft = slotLeft + (slotWidth - cell.width) / 2;
            var targetTop = slotTop - (slotHeight - cell.height) / 2;

            // アートボードがある場合はその枠、無い場合は記録した現在のセル左上を基準に移動量を求める
            // Use the artboard frame if present; otherwise the recorded current cell top-left
            var rect = cell.artboard
                ? cell.artboard.artboardRect // [L, T, R, B]
                : [cell.currentLeft, cell.currentTop, cell.currentLeft + cell.width, cell.currentTop - cell.height];
            var dx = targetLeft - rect[0];
            var dy = targetTop - rect[1];

            cell.group.translate(dx, dy);
            if (cell.artboard) cell.artboard.artboardRect = [rect[0] + dx, rect[1] + dy, rect[2] + dx, rect[3] + dy];
            if (cell.label) cell.label.translate(dx, dy);
        }
    }

    /* メイン処理：ダイアログを表示し、選択されたソースを新規ドキュメントへ取り込み配置する
       / Main entry point: show the dialog, then import and arrange the chosen source into a new document */
    (function main() {
        // 現在開いているドキュメントを収集 / Collect currently open documents
        var openDocs = [];
        for (var i = 0; i < app.documents.length; i++) {
            openDocs.push(app.documents[i]);
        }
        // ［指定］ボタンで選んだフォルダと、その中の対象ファイル一覧
        // The folder chosen via the [Choose] button and the matching files inside it
        var selectedFolder = null;
        var folderFiles = [];        // 種別＋名前フィルタ後のファイル / Files after the type + name filter
        var folderFilesTotal = 0;    // 種別フィルタに一致する総数（名前フィルタ前）/ Total matching the type filter (before the name filter)

        var dialog = new Window("dialog", getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["left", "top"];

        // --- 読み込み対象パネル / Source panel ---
        var sourcePanel = dialog.add("panel", undefined, getLabel('panel.source'));
        setupPanel(sourcePanel);
        var openDocsRadio = sourcePanel.add("radiobutton", undefined, getLabel('radio.openFiles'));
        openDocsRadio.helpTip = getLabel('tooltip.openFiles');

        var folderRow = sourcePanel.add("group");
        setupGroup(folderRow, "row");
        var folderRadio = folderRow.add("radiobutton", undefined, getLabel('radio.specifyFolder'));
        folderRadio.helpTip = getLabel('tooltip.specifyFolder');
        var selectFolderBtn = folderRow.add("button", undefined, getLabel('button.specify'));
        selectFolderBtn.helpTip = getLabel('tooltip.specify');

        // 選択したフォルダ名は別行に表示 / Show the chosen folder name on its own row
        var folderNameRow = sourcePanel.add("group");
        setupGroup(folderNameRow, "row");
        var folderNameText = folderNameRow.add("statictext", undefined, "", { truncate: "middle" });
        folderNameText.preferredSize = [360, 20];
        folderNameText.minimumSize = [360, 20];

        // --- フィルターパネル（読み込み対象パネル内。種別＋ファイル名の正規表現で絞り込み）/ Filter panel (nested in the source panel) ---
        var filterPanel = sourcePanel.add("panel", undefined, getLabel('panel.filter'));
        setupPanel(filterPanel);

        // 種別（読み込むファイル形式、チェックボックスは横並び）/ File types (checkboxes laid out horizontally)
        var typeRow = filterPanel.add("group");
        setupGroup(typeRow, "row");
        var typeLabel = typeRow.add("statictext", undefined, labelText('field.fileType'));
        typeLabel.preferredSize = [100, 20];
        var aiCheckbox = typeRow.add("checkbox", undefined, "AI");
        aiCheckbox.helpTip = getLabel('tooltip.fileType');
        var svgCheckbox = typeRow.add("checkbox", undefined, "SVG");
        svgCheckbox.helpTip = getLabel('tooltip.fileType');
        var epsCheckbox = typeRow.add("checkbox", undefined, "EPS");
        epsCheckbox.helpTip = getLabel('tooltip.fileType');
        aiCheckbox.value = true;   // 既定は AI のみ ON / Default: AI only
        svgCheckbox.value = false;
        epsCheckbox.value = false;

        var filterRow = filterPanel.add("group");
        setupGroup(filterRow, "row");
        var filterLabel = filterRow.add("statictext", undefined, labelText('field.fileName'));
        filterLabel.preferredSize = [100, 20];
        var filterInput = filterRow.add("edittext", undefined, "");
        filterInput.characters = 20;
        filterInput.helpTip = getLabel('hint.filter');
        filterLabel.helpTip = getLabel('hint.filter');

        // ファイル数はパネルのタイトルに表示する。フィルター使用時は「読み込み対象（3/5）」（絞り込み後/総数）、
        // 未使用時は「読み込み対象（5）」のように表示する。
        // Show the file count in the panel title. With an active filter, "Source (3/5)" (filtered/total); otherwise "Source (5)".
        function updateFileCount() {
            var filterActive = (filterInput.text !== "" && getNameFilterRegExp() !== null);
            var total, filtered;
            if (folderRadio.value) {
                total = folderFilesTotal;        // 種別フィルタに一致する総数 / total matching the type filter
                filtered = folderFiles.length;   // 種別＋名前フィルタ後 / after the type + name filter
            } else {
                total = openDocs.length;
                filtered = getFilteredOpenDocs().length;
            }
            var countText = filterActive ? (filtered + '/' + total) : String(total);
            sourcePanel.text = getLabel('panel.source') + (uiLang === 'ja' ? '（' + countText + '）' : ' (' + countText + ')');
        }

        // 「読み込み後の動作」は開いているファイルを選んだときのみ有効（フォルダ指定ではディム）
        // The "After Import" action is available only for open files (dimmed for folder import)
        function updateAfterImportState() {
            var enabled = openDocsRadio.value;
            afterImportLabel.enabled = enabled;
            closeRadio.enabled = enabled;
            keepOpenRadio.enabled = enabled;
        }

        // 種別は「フォルダー指定」のときのみ有効 / File types are available only in folder mode
        function updateTypeRowState() {
            typeLabel.enabled = folderRadio.value;
            aiCheckbox.enabled = folderRadio.value;
            svgCheckbox.enabled = folderRadio.value;
            epsCheckbox.enabled = folderRadio.value;
        }

        // ファイル名ラベルはフォルダー指定では既定OFF、開いているファイルではON
        // Default the file-name label off in folder mode and on for open files
        function updateLabelOptionForMode() {
            showLabelCheckbox.value = openDocsRadio.value;
        }

        // 選択された種別の拡張子リスト / Extensions for the currently selected file types
        function getSelectedExtensions() {
            var exts = [];
            if (aiCheckbox.value) exts.push("ai");
            if (svgCheckbox.value) exts.push("svg");
            if (epsCheckbox.value) exts.push("eps");
            return exts;
        }

        // ファイル名フィルタの正規表現を返す。空欄や不正な式なら null（＝全件対象）。
        // Return the file-name filter as a RegExp; null when empty or invalid (= match all).
        function getNameFilterRegExp() {
            var pattern = filterInput.text;
            if (pattern === null || pattern === "") return null;
            try {
                return new RegExp(pattern, "i");
            } catch (e) {
                return null; // 入力途中の不正な式では絞り込まない / Don't filter on an incomplete/invalid pattern
            }
        }

        // ファイル名がフィルタに一致するか / Whether a file name matches the current filter
        function matchesNameFilter(name) {
            var nameRe = getNameFilterRegExp();
            return nameRe === null || nameRe.test(name);
        }

        // 開いているファイルをフィルタで絞り込んだ配列 / Open documents narrowed by the name filter
        function getFilteredOpenDocs() {
            var matched = [];
            for (var i = 0; i < openDocs.length; i++) {
                if (matchesNameFilter(openDocs[i].name)) matched.push(openDocs[i]);
            }
            return matched;
        }

        // 選択中フォルダを現在の種別とファイル名フィルタで再フィルタ / Re-filter the chosen folder by file types and the name filter
        // 種別に一致する総数（folderFilesTotal）と、名前フィルタ後の一覧（folderFiles）の両方を更新する。
        // Updates both the type-matched total (folderFilesTotal) and the name-filtered list (folderFiles).
        function refreshFolderFiles() {
            folderFiles = [];
            folderFilesTotal = 0;
            var exts = getSelectedExtensions();
            if (selectedFolder && exts.length > 0) {
                var typeRe = new RegExp("\\.(" + exts.join("|") + ")$", "i");
                var typeMatched = selectedFolder.getFiles(function (f) {
                    return f instanceof File && typeRe.test(f.name);
                });
                folderFilesTotal = typeMatched.length;
                var matched = [];
                for (var i = 0; i < typeMatched.length; i++) {
                    if (matchesNameFilter(typeMatched[i].name)) matched.push(typeMatched[i]);
                }
                matched.sort(function (a, b) {
                    return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
                });
                folderFiles = matched;
            }
            updateFileCount();
        }

        // フィルタ入力に応じて件数・一覧を更新 / Refresh counts and lists as the filter changes
        filterInput.onChanging = function () {
            refreshFolderFiles(); // updateFileCount を内包（開いているファイル件数も更新）/ also refreshes the open-files count
        };

        aiCheckbox.onClick = function () {
            refreshFolderFiles();
        };
        svgCheckbox.onClick = aiCheckbox.onClick;
        epsCheckbox.onClick = aiCheckbox.onClick;

        // 開いているファイルが無ければフォルダ指定をデフォルトに
        // Default to folder mode when no document is open
        openDocsRadio.enabled = openDocs.length > 0;
        if (openDocs.length > 0) {
            openDocsRadio.value = true;
        } else {
            folderRadio.value = true;
        }

        // 2つのラジオは別コンテナにあるため、排他選択を手動で制御する
        // The two radios live in different containers, so enforce mutual exclusivity manually
        openDocsRadio.onClick = function () {
            folderRadio.value = false;
            updateFileCount();
            updateAfterImportState();
            updateTypeRowState();
            updateLabelOptionForMode();
        };
        folderRadio.onClick = function () {
            openDocsRadio.value = false;
            updateFileCount();
            updateAfterImportState();
            updateTypeRowState();
            updateLabelOptionForMode();
        };

        selectFolderBtn.onClick = function () {
            var folder = Folder.selectDialog(getLabel('prompt.selectFolder'));
            if (!folder) return;
            selectedFolder = folder;
            folderRadio.value = true;
            openDocsRadio.value = false;
            var folderPath = decodeURI(folder.fsName);
            folderNameText.text = folderPath;
            folderNameText.helpTip = folderPath;
            refreshFolderFiles();
            updateAfterImportState();
            updateTypeRowState();
            updateLabelOptionForMode();
        };

        updateFileCount();
        updateTypeRowState();

        // --- 読み込み先パネル（読み込み先の選択＋新規ドキュメント設定）/ Destination panel (target choice + new-document settings) ---
        var destinationPanel = dialog.add("panel", undefined, getLabel('panel.destination'));
        setupPanel(destinationPanel);

        // 読み込み先：現在のドキュメント／新規ドキュメント / Destination: current document or a new one
        var destRow = destinationPanel.add("group");
        setupGroup(destRow, "row");
        destRow.alignment = ["center", "top"]; // 左右中央 / Center horizontally
        destRow.margins = [0, 0, 0, 10];        // 下に10pxの余白 / 10px margin below
        var currentDocRadio = destRow.add("radiobutton", undefined, getLabel('radio.currentDoc'));
        currentDocRadio.helpTip = getLabel('tooltip.currentDoc');
        var newDocRadio = destRow.add("radiobutton", undefined, getLabel('radio.newDoc'));
        newDocRadio.helpTip = getLabel('tooltip.newDoc');
        newDocRadio.value = true; // 既定は新規ドキュメント / Default: new document

        // カラーモード＋解像度とサイズを横2カラムで並べる / Lay out color mode + resolution and size in two columns
        var newDocSettingsRow = destinationPanel.add("group");
        setupGroup(newDocSettingsRow, "row");
        newDocSettingsRow.alignChildren = ["left", "fill"]; // 2つのパネルの高さを揃える / Match the two panels' heights
        newDocSettingsRow.spacing = 15;                     // カラーモード／解像度とサイズの2カラム間隔 / Gap between the color-mode/resolution and size columns

        // 読み込み先が「現在のドキュメント」のときは新規ドキュメント設定をディムにする
        // Dim the new-document settings when the destination is the current document
        function updateDestinationState() {
            newDocSettingsRow.enabled = newDocRadio.value;
            redrawSteppersIn(newDocSettingsRow); /* 幅・高さの∧∨のディム表示を切り替える / update the steppers' dimming */
        }
        currentDocRadio.onClick = updateDestinationState;
        newDocRadio.onClick = updateDestinationState;

        // 左カラム：カラーモードと解像度を縦に積む / Left column: color mode + resolution stacked
        var colorAndResolutionColumn = newDocSettingsRow.add("group");
        setupGroup(colorAndResolutionColumn, "column");
        colorAndResolutionColumn.alignChildren = ["fill", "top"];

        // カラーモード（ラジオは縦並び）/ Color mode (radios stacked vertically)
        var colorModePanel = colorAndResolutionColumn.add("panel", undefined, getLabel('panel.colorMode'));
        setupPanel(colorModePanel);
        var rgbRadio = colorModePanel.add("radiobutton", undefined, getLabel('radio.rgb'));
        rgbRadio.helpTip = getLabel('tooltip.colorMode');
        var cmykRadio = colorModePanel.add("radiobutton", undefined, getLabel('radio.cmyk'));
        cmykRadio.helpTip = getLabel('tooltip.colorMode');
        rgbRadio.value = true;

        // 解像度（ラスタライズ効果設定の ppi）/ Resolution (raster effects ppi)
        var resolutionPanel = colorAndResolutionColumn.add("panel", undefined, getLabel('panel.resolution'));
        setupPanel(resolutionPanel);
        var resolutionDropdown = resolutionPanel.add("dropdownlist", undefined, ["72", "150", "300"]);
        resolutionDropdown.selection = 2; // デフォルトは300 / Default 300

        var sizePanel = newDocSettingsRow.add("panel", undefined, getLabel('panel.docSize'));
        setupPanel(sizePanel);

        var presetDropdown = sizePanel.add("dropdownlist", undefined, [
            getLabel('preset.custom'),
            getLabel('preset.a4'),
            getLabel('preset.fullHD'),
            getLabel('preset.largeCanvas')
        ]);
        presetDropdown.selection = 0;

        // 選択中の単位（プリセット連動、カスタムは px）/ Current unit (follows the preset; px for Custom)
        var currentUnit = "px";

        // 幅と高さはそれぞれ別の行に / Width and height each on their own row
        var widthRow = sizePanel.add("group");
        setupGroup(widthRow, "row");
        var widthLabel = widthRow.add("statictext", undefined, labelText('field.width'));
        widthLabel.preferredSize = [40, 20];
        var widthInput = addSteppedInput(widthRow, "1000", {});
        widthInput.characters = 5;
        widthInput.helpTip = getLabel('tooltip.size');

        var heightRow = sizePanel.add("group");
        setupGroup(heightRow, "row");
        var heightLabel = heightRow.add("statictext", undefined, labelText('field.height'));
        heightLabel.preferredSize = [40, 20];
        var heightInput = addSteppedInput(heightRow, "1000", {});
        heightInput.characters = 5;
        heightInput.helpTip = getLabel('tooltip.size');

        // 単位（mm / px）。A4 は mm、それ以外は px を既定にし、手動切替で値を換算する。
        // Unit (mm / px). Default mm for A4, px otherwise; switching converts the values.
        var unitRow = sizePanel.add("group");
        setupGroup(unitRow, "row");
        var unitLabel = unitRow.add("statictext", undefined, labelText('field.unit'));
        unitLabel.preferredSize = [40, 20];
        var unitDropdown = unitRow.add("dropdownlist", undefined, UNIT_LIST);

        // 入力値を旧単位から新単位へ換算する（物理サイズを保つ）/ Convert an input value from the old unit to the new one (keeps physical size)
        function convertUnitText(textValue, fromUnit, toUnit) {
            var value = parseFloat(textValue);
            if (isNaN(value)) return textValue;
            return String(Math.round(value * UNIT_TO_PT[fromUnit] / UNIT_TO_PT[toUnit]));
        }

        // ドロップダウンの選択をプログラムから変更（onChange を誤発火させない）/ Set the dropdown selection without firing onChange
        var suppressUnitChange = false;
        function selectUnit(unitName) {
            suppressUnitChange = true;
            for (var i = 0; i < UNIT_LIST.length; i++) {
                if (UNIT_LIST[i] === unitName) { unitDropdown.selection = i; break; }
            }
            suppressUnitChange = false;
        }
        selectUnit(currentUnit); // 初期は px（カスタム）/ Initial: px (Custom)

        // 手動で単位を切り替えたら現在の値を換算 / Convert current values when the unit is switched manually
        unitDropdown.onChange = function () {
            if (suppressUnitChange || !unitDropdown.selection) return;
            var newUnit = unitDropdown.selection.text;
            if (newUnit === currentUnit) return;
            widthInput.text = convertUnitText(widthInput.text, currentUnit, newUnit);
            heightInput.text = convertUnitText(heightInput.text, currentUnit, newUnit);
            currentUnit = newUnit;
        };

        presetDropdown.onChange = function () {
            var idx = presetDropdown.selection.index;
            var preset = SIZE_PRESETS[idx];
            var isA4 = idx === 1;
            var newUnit = isA4 ? "mm" : "px"; // A4 は mm、それ以外は px / mm for A4, px otherwise
            if (preset.width !== undefined) {
                // プリセット：ネイティブ単位の値を表示単位へ換算 / Preset: convert its native-unit values to the display unit
                widthInput.text = convertUnitText(String(preset.width), preset.unit, newUnit);
                heightInput.text = convertUnitText(String(preset.height), preset.unit, newUnit);
            } else {
                // カスタム：現在の値を新しい単位へ換算して物理サイズを保つ / Custom: convert current values to keep the physical size
                widthInput.text = convertUnitText(widthInput.text, currentUnit, newUnit);
                heightInput.text = convertUnitText(heightInput.text, currentUnit, newUnit);
            }
            currentUnit = newUnit;
            selectUnit(currentUnit);

            // A4（印刷向け）は CMYK、それ以外（画面向け）は RGB を既定にする
            // Default to CMYK for A4 (print), RGB for the others (screen)
            cmykRadio.value = isA4;
            rgbRadio.value = !isA4;
        };

        // --- 読み込みオプションパネル / Import options panel ---
        var optionsPanel = dialog.add("panel", undefined, getLabel('panel.options'));
        setupPanel(optionsPanel);
        // オプションを2カラムで並べる（左：アートボード単位／ファイル名ラベル／ガイド／拡大・縮小、右：対象アートボード）
        // Lay out options in two columns (left: per-artboard / file-name label / guides / scale, right: target artboards)
        var optionsColumns = optionsPanel.add("group");
        setupGroup(optionsColumns, "row");
        optionsColumns.alignChildren = ["left", "top"]; // 2カラムを上端で揃える / Top-align the two columns
        optionsColumns.spacing = 20;                     // 左右カラムの間隔 / Gap between the two columns

        // 左カラム：アートボード単位／ファイル名ラベル／ガイド／拡大・縮小 / Left column
        var optionsLeftColumn = optionsColumns.add("group");
        setupGroup(optionsLeftColumn, "column");
        var byArtboardCheckbox = optionsLeftColumn.add("checkbox", undefined, getLabel('checkbox.byArtboard'));
        byArtboardCheckbox.value = true;
        byArtboardCheckbox.helpTip = getLabel('tooltip.byArtboard');

        var showLabelCheckbox = optionsLeftColumn.add("checkbox", undefined, getLabel('checkbox.attachLabel'));
        showLabelCheckbox.value = openDocsRadio.value; // フォルダー指定では既定OFF / Default off in folder mode
        showLabelCheckbox.helpTip = getLabel('tooltip.attachLabel');

        var includeGuidesCheckbox = optionsLeftColumn.add("checkbox", undefined, getLabel('checkbox.includeGuides'));
        includeGuidesCheckbox.value = false;
        includeGuidesCheckbox.helpTip = getLabel('tooltip.includeGuides');

        // スケール（チェックON時に％で拡大縮小）/ Scale (resize by percent when checked)
        var scaleRow = optionsLeftColumn.add("group");
        setupGroup(scaleRow, "row");
        var scaleCheckbox = scaleRow.add("checkbox", undefined, getLabel('checkbox.scale'));
        scaleCheckbox.value = false;
        scaleCheckbox.helpTip = getLabel('tooltip.scale');
        var scaleInput = addSteppedInput(scaleRow, "100", {});
        scaleInput.helpTip = getLabel('tooltip.scale');
        scaleInput.characters = 4;
        /* ∧∨は入力欄の兄弟なので、∧∨と入力欄を包む group ごと切り替えてディム表示にする
           the stepper is a sibling of the field, so toggle their wrapper group to dim it */
        scaleInput.parent.enabled = scaleCheckbox.value;
        scaleInput.enabled = scaleCheckbox.value;
        scaleRow.add("statictext", undefined, "%");
        scaleCheckbox.onClick = function () {
            scaleInput.parent.enabled = this.value;
            scaleInput.enabled = this.value;
            redrawSteppersIn(scaleInput.parent);
        };

        // 右カラム：対象アートボード（1のみ／すべて／指定）をパネルに / Right column: target artboards in a panel
        var targetArtboardPanel = optionsColumns.add("panel", undefined, getLabel('field.artboardTarget'));
        setupPanel(targetArtboardPanel);
        var artboardOneRadio = targetArtboardPanel.add("radiobutton", undefined, getLabel('radio.artboardOne'));
        artboardOneRadio.helpTip = getLabel('tooltip.artboardTarget');
        var artboardAllRadio = targetArtboardPanel.add("radiobutton", undefined, getLabel('radio.artboardAll'));
        artboardAllRadio.helpTip = getLabel('tooltip.artboardTarget');
        var artboardSpecRow = targetArtboardPanel.add("group");
        setupGroup(artboardSpecRow, "row");
        var artboardSpecRadio = artboardSpecRow.add("radiobutton", undefined, getLabel('radio.artboardSpecify'));
        artboardSpecRadio.helpTip = getLabel('tooltip.artboardSpecify');
        var artboardSpecInput = artboardSpecRow.add("edittext", undefined, "");
        artboardSpecInput.characters = 7;
        artboardSpecInput.helpTip = getLabel('tooltip.artboardSpecify');

        // 3つのラジオは別コンテナにまたがるため、排他選択を手動で制御する
        // The three radios span different containers, so enforce mutual exclusivity manually
        function selectArtboardTarget(which) {
            artboardOneRadio.value = (which === "one");
            artboardAllRadio.value = (which === "all");
            artboardSpecRadio.value = (which === "spec");
            artboardSpecInput.enabled = artboardSpecRadio.value;
        }
        artboardOneRadio.onClick = function () { selectArtboardTarget("one"); };
        artboardAllRadio.onClick = function () { selectArtboardTarget("all"); };
        artboardSpecRadio.onClick = function () { selectArtboardTarget("spec"); };
        selectArtboardTarget("one"); // 既定：アートボード1のみ / Default: Artboard 1 only

        // 対象アートボードは「アートボード単位」ONのときのみ有効（OFFでディム）
        // Target artboards apply only when "per artboard" is on (dimmed when off)
        function updateArtboardTargetState() {
            targetArtboardPanel.enabled = byArtboardCheckbox.value;
        }
        byArtboardCheckbox.onClick = updateArtboardTargetState;
        updateArtboardTargetState(); // 初期状態を反映 / Apply the initial state

        // 読み込み後の動作（ラベル＋ラジオの横並び。開いているファイル選択時のみ有効）
        // After-import action (label + radios in a row; enabled only when importing open files)
        var afterImportRow = optionsPanel.add("group");
        setupGroup(afterImportRow, "row");
        var afterImportLabel = afterImportRow.add("statictext", undefined, labelText('panel.afterImport'));
        var closeRadio = afterImportRow.add("radiobutton", undefined, getLabel('radio.closeDoc'));
        closeRadio.helpTip = getLabel('tooltip.closeDoc');
        var keepOpenRadio = afterImportRow.add("radiobutton", undefined, getLabel('radio.keepOpen'));
        keepOpenRadio.helpTip = getLabel('tooltip.keepOpen');
        closeRadio.value = true;

        updateAfterImportState(); // 初期状態を反映 / Apply the initial enabled/dimmed state

        var buttonRow = addButtonRow(dialog);
        // name を "cancel"/"ok" にすると、クリックでダイアログが閉じる（Esc/Enter にも対応）
        // Naming them "cancel"/"ok" makes clicks dismiss the dialog (and binds Esc/Enter)
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });

        prepareDialogWindow(dialog, SCRIPT_NAME);
        if (dialog.show() !== 1) return;

        // ファイル名フィルターの正規表現を検証する。入力途中は無視してよいが、OK後は不正な式を黙って
        // 「全件対象」にせず中断する（誤って意図しないファイルまで取り込む事故を防ぐ）。
        // Validate the file-name filter regex. It's fine to ignore while typing, but after OK don't silently
        // fall back to "match all" on an invalid pattern — stop instead (avoids importing unintended files).
        if (filterInput.text !== "") {
            try {
                new RegExp(filterInput.text);
            } catch (eFilterPattern) {
                alert(getLabel('alert.invalidFilter'));
                return;
            }
        }

        // 選択されたソースとオプションを確定 / Resolve the chosen source and options
        var importFromFolder = folderRadio.value;
        var useCurrentDoc = currentDocRadio.value; // 読み込み先：現在のドキュメント / Destination: current document
        var importByArtboard = byArtboardCheckbox.value;
        // アートボード対象：first（1のみ）/ all（すべて）/ specify（指定） / Target artboards
        var artboardTargetMode = artboardOneRadio.value ? "first" : (artboardAllRadio.value ? "all" : "specify");
        var artboardSpecText = artboardSpecInput.text;
        var includeGuides = includeGuidesCheckbox.value;
        var originalDocs = importFromFolder ? folderFiles : getFilteredOpenDocs();

        // 「現在のドキュメント」に読み込む場合は対象ドキュメントを確定する / Resolve the target document when importing into the current one
        var targetDoc = null;
        if (useCurrentDoc) {
            if (app.documents.length === 0) {
                alert(getLabel('alert.noCurrentDoc'));
                return;
            }
            targetDoc = app.activeDocument;
            // 対象ドキュメント自身は取り込み対象から外す（誤って自分を開いて閉じるのを防ぐ）
            // Exclude the target document itself from the sources (avoid opening and then closing it)
            if (importFromFolder) {
                // フォルダー指定：同じファイルパスのファイルを除外（対象が保存済みファイルの場合のみ判定可能）
                // Folder mode: drop any file with the same path as the target (only when the target is a saved file)
                var targetPath = null;
                try {
                    var targetFile = targetDoc.fullName;
                    if (targetFile && targetFile.exists) targetPath = targetFile.fsName;
                } catch (ePath) { targetPath = null; }
                if (targetPath !== null) {
                    var folderSourcesWithoutTarget = [];
                    for (var fi = 0; fi < originalDocs.length; fi++) {
                        if (originalDocs[fi].fsName !== targetPath) folderSourcesWithoutTarget.push(originalDocs[fi]);
                    }
                    originalDocs = folderSourcesWithoutTarget;
                }
            } else {
                // 開いているファイル指定：対象ドキュメントの参照を除外 / Open-files mode: drop the target document reference
                var sourcesWithoutTarget = [];
                for (var od = 0; od < originalDocs.length; od++) {
                    if (originalDocs[od] !== targetDoc) sourcesWithoutTarget.push(originalDocs[od]);
                }
                originalDocs = sourcesWithoutTarget;
            }
        }

        if (originalDocs.length < 1) {
            alert(getLabel('alert.noValidFile'));
            return;
        }
        var originalDocsLength = originalDocs.length;

        // 「指定」モードでアートボード番号が一つも解釈できない場合は中断（無言終了を防ぐ）
        // Abort if "Specify" mode yields no parseable artboard numbers (avoids silently finishing)
        if (importByArtboard && artboardTargetMode === "specify" && parseArtboardNumbers(artboardSpecText).length === 0) {
            alert(getLabel('alert.invalidArtboardSpec'));
            return;
        }

        // スケール（％）。チェックOFFなら100%（等倍）。ONで不正値（数値でない・0以下）は黙って等倍にせず中断する。
        // Scale percent; 100% when unchecked. When checked, abort on an invalid value (non-numeric or <= 0) instead of silently using 100%.
        var scalePercent = 100;
        if (scaleCheckbox.value) {
            var scaleValue = parseFloat(scaleInput.text);
            if (isNaN(scaleValue) || scaleValue <= 0) {
                alert(getLabel('alert.invalidScale'));
                return;
            }
            scalePercent = scaleValue;
        }

        // 新規ドキュメント作成時のみ、寸法をプログレス表示の前に検証する（無効ならパレットを残さず終了）
        // For a new document only, validate the size before showing the progress palette (so it isn't left open on error)
        var docWidthValue, docHeightValue;
        if (!useCurrentDoc) {
            docWidthValue = parseFloat(widthInput.text);
            docHeightValue = parseFloat(heightInput.text);
            if (isNaN(docWidthValue) || isNaN(docHeightValue)) {
                alert(getLabel('alert.invalidNumber'));
                return;
            }
        }

        // フォルダ読み込みで開いた一時ファイルは常に閉じる。開いているファイルのみ「読み込み後の動作」に従う。
        // Temp files opened from a folder are always closed; only open files honor the "After Import" choice.
        var shouldCloseSource = importFromFolder || closeRadio.value;

        // 開いているファイルを閉じる場合、未保存変更があれば確認（変更が失われるため）
        // When closing open files, confirm if any have unsaved changes (they would be lost)
        if (!importFromFolder && closeRadio.value) {
            var hasUnsavedChanges = false;
            for (var d = 0; d < originalDocs.length; d++) {
                if (!originalDocs[d].saved) {
                    hasUnsavedChanges = true;
                    break;
                }
            }
            if (hasUnsavedChanges && !confirm(getLabel('confirm.discardUnsaved'))) {
                return;
            }
        }

        // プログレスバーのダイアログを表示
        var progressWin = new Window("palette", getLabel('progress.title'));
        progressWin.orientation = "column";
        progressWin.alignChildren = ["fill", "top"];
        progressWin.margins = 20;
        var progressTextGroup = progressWin.add("group");
        progressTextGroup.alignment = ["center", "top"];
        var processedCountStatic = progressTextGroup.add("statictext", undefined, labelText('progress.count') + "0/" + originalDocsLength);
        processedCountStatic.preferredSize = [100, 30];

        var progressBar = progressWin.add("progressbar", undefined, 0, originalDocsLength);
        progressBar.preferredSize = [300, 6];

        var cancelGroup = progressWin.add("group");
        cancelGroup.alignment = "right";
        var progressCancelBtn = cancelGroup.add("button", undefined, getLabel('button.cancel'));

        var userCancelled = false;

        progressCancelBtn.onClick = function () {
            userCancelled = true;
        };
        progressWin.addEventListener("keydown", function (e) {
            if (e.keyName === "Escape") {
                userCancelled = true;
            }
        });
        progressWin.show();

        var newDoc;
        if (useCurrentDoc) {
            // 現在のドキュメントに読み込む（新規作成しない。カラーモード・解像度・サイズ設定は使わない）
            // Import into the current document (no new doc; color mode / resolution / size settings are not used)
            newDoc = targetDoc;
            app.activeDocument = newDoc;
        } else {
            // 入力値を現在の単位からポイントへ換算 / Convert the input values from the current unit to points
            var ptPerUnit = UNIT_TO_PT[currentUnit];
            var docWidthPt = docWidthValue * ptPerUnit;
            var docHeightPt = docHeightValue * ptPerUnit;
            var colorSpace = rgbRadio.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK;

            // 通常ドキュメントの最大寸法（pt）。これを超える場合はラージカンバスとして作成。
            // Max dimension (pt) of a standard document; beyond this, create as a large canvas.
            var STANDARD_MAX_PT = 16383;

            if (docWidthPt > STANDARD_MAX_PT || docHeightPt > STANDARD_MAX_PT) {
                var largeCanvasPreset = new DocumentPreset();
                largeCanvasPreset.units = RulerUnits.Points;
                largeCanvasPreset.width = docWidthPt;
                largeCanvasPreset.height = docHeightPt;
                largeCanvasPreset.colorMode = colorSpace;
                largeCanvasPreset.numArtboards = 1;
                newDoc = app.documents.addDocument("Print", largeCanvasPreset);
            } else {
                newDoc = app.documents.add(colorSpace, docWidthPt, docHeightPt);
            }
            app.activeDocument = newDoc;

            // ラスタライズ効果設定の解像度を反映 / Apply the selected raster effects resolution
            var rasterPpi = parseInt(resolutionDropdown.selection.text, 10);
            if (!isNaN(rasterPpi)) {
                var rasterOptions = newDoc.rasterEffectSettings;
                rasterOptions.resolution = rasterPpi;
                newDoc.rasterEffectSettings = rasterOptions;
            }
        }

        // 配置の基準アートボード。新規は初期（仮）アートボード、現在のドキュメントはアクティブなアートボード。
        // Base artboard for placement: the placeholder for a new doc, the active artboard for the current doc.
        var baseArtboardIndex = useCurrentDoc ? newDoc.artboards.getActiveArtboardIndex() : 0;
        var artboardRect = newDoc.artboards[baseArtboardIndex].artboardRect; // [L, T, R, B]
        var canvasCenterX = (artboardRect[0] + artboardRect[2]) / 2;
        var canvasCenterY = (artboardRect[1] + artboardRect[3]) / 2;

        // 現在のドキュメントに読み込む場合は、取り込み前の既存アートボード全体を記録し、その下に重ならないよう配置する。
        // When importing into the current document, capture the existing artboards (before import) so the grid lands below them without overlapping.
        var avoidRect = useCurrentDoc ? getArtboardsUnionRect(newDoc) : null;

        // 配置の状態をまとめて保持。各セルを記録し、ループ後に正方形グリッドへ中央配置する。
        // Shared placement state. Each cell is recorded, then laid out into a centered square grid after the loop.
        var placementContext = {
            newDoc: newDoc,
            cells: [],
            canvasCenterX: canvasCenterX,
            canvasCenterY: canvasCenterY,
            avoidRect: avoidRect,
            byArtboard: importByArtboard,
            createArtboards: importByArtboard, // アートボードを作るのはアートボード単位ONのときだけ / Create artboards only in per-artboard mode
            placedCount: 0,
            showLabel: showLabelCheckbox.value,
            artboardPadding: CONFIG.artboardMargin,
            scalePercent: scalePercent
        };

        // ループ全体を try/finally で囲み、エラーが出てもプログレスは必ず閉じる
        // Wrap the whole loop so the progress palette is always closed, even on error
        try {
            for (var j = 0; j < originalDocs.length; j++) {
                $.sleep(0);
                app.redraw();
                // キャンセルされたらループを抜けて、ここまでの結果で後始末する / On cancel, break and finish with what was placed so far
                if (userCancelled) break;

                var srcDoc = importFromFolder ? app.open(originalDocs[j]) : originalDocs[j];

                // 1ファイル分の処理。エラーが出ても finally で一時ファイルを必ず閉じ、状態も復元する。
                // Process one file; finally always closes the temp file and restores state, even on error.
                var lockState = null;
                try {
                    app.activeDocument = srcDoc;
                    var labelName = srcDoc.name.replace(/\.[^\.]+$/, "");

                    if (importByArtboard) {
                        // アートボード単位：対象アートボードごとに、ロック・非表示も含めて取り込む
                        // Per artboard: import every object (incl. locked/hidden) on each target artboard
                        var targetIndices = resolveTargetArtboardIndices(srcDoc, artboardTargetMode, artboardSpecText);
                        // すべて：ドキュメント全体を解除。1のみ／指定：対象アートボードに重なるオブジェクトだけ解除し、他は触れない。
                        // 閉じない場合のみ後で復元する。
                        // All: unlock the whole doc. First/Specify: unlock only the objects overlapping the target
                        // artboards, leaving the rest untouched. Restore afterward only when keeping the file open.
                        if (artboardTargetMode === "all") {
                            if (!shouldCloseSource) lockState = captureLockHiddenState(srcDoc);
                            unlockAllLayersAndItems(srcDoc);
                        } else {
                            var targetRects = [];
                            for (var ti = 0; ti < targetIndices.length; ti++) {
                                targetRects.push(srcDoc.artboards[targetIndices[ti]].artboardRect);
                            }
                            var scopedState = unlockItemsOnArtboards(srcDoc, targetRects);
                            if (!shouldCloseSource) lockState = scopedState;
                        }
                        for (var ai = 0; ai < targetIndices.length; ai++) {
                            var ab = targetIndices[ai];
                            srcDoc.artboards.setActiveArtboardIndex(ab);
                            srcDoc.selection = null;
                            srcDoc.selectObjectsOnActiveArtboard();
                            if (srcDoc.selection.length === 0) continue;

                            var abRect = srcDoc.artboards[ab].artboardRect; // [L, T, R, B]

                            // ガイドを含める場合は、このアートボードに重なる短いガイドを選択へ追加（相対位置の算出前に）
                            // When including guides, add the short guides overlapping this artboard before measuring bounds
                            if (includeGuides) {
                                var abGuides = collectImportableGuides(srcDoc, abRect[2] - abRect[0], abRect[1] - abRect[3], abRect);
                                if (abGuides.length > 0) {
                                    var combined = [];
                                    for (var g = 0; g < srcDoc.selection.length; g++) combined.push(srcDoc.selection[g]);
                                    srcDoc.selection = combined.concat(abGuides);
                                }
                            }

                            // 元のアートボードサイズと、その中での内容（＋ガイド）の相対位置を記録
                            // Record the original artboard size and the relative position of the content (and guides)
                            /* クリップグループはマスクで測る（貼り付け側の placePastedGroup と同じ測り方）
                               Clip groups are measured by their mask, the same way placePastedGroup measures the pasted side */
                            var selBounds = getClipAwareUnionBounds(srcDoc.selection, true);
                            var artboardCell = {
                                width: abRect[2] - abRect[0],
                                height: abRect[1] - abRect[3],
                                offsetX: selBounds[0] - abRect[0],
                                offsetY: abRect[1] - selBounds[1]
                            };

                            app.copy();
                            app.activeDocument = newDoc;
                            app.paste();
                            var placedArtboard = placePastedGroup(placementContext, labelName, artboardCell);
                            app.activeDocument = srcDoc;
                            // 配置に失敗したアートボードはスキップ（既にアラート済み）/ Skip artboards that failed to place (already alerted)
                            if (!placedArtboard) continue;
                        }
                    } else {
                        // 通常（アートボード単位OFF）：ファイル全体の表示中・選択可能オブジェクトを1つにまとめて取り込む。
                        // 「対象アートボード」はアートボード単位ONのときのみ有効なので、ここでは全選択する。
                        // Default (per-artboard off): import the whole file's visible/selectable objects as one.
                        // "Target artboards" applies only when per-artboard is on, so select everything here.
                        app.executeMenuCommand("selectall");

                        var filteredSelection = [];
                        for (var s = 0; s < srcDoc.selection.length; s++) {
                            var obj = srcDoc.selection[s];
                            if (!obj.locked && !obj.hidden) {
                                filteredSelection.push(obj);
                            }
                        }
                        // ガイドを含める場合は、短いガイド（カンバス＝先頭アートボード基準）を追加
                        // When including guides, add the short guides (canvas = the first artboard)
                        if (includeGuides) {
                            var firstArtboard = srcDoc.artboards[0].artboardRect;
                            filteredSelection = filteredSelection.concat(
                                collectImportableGuides(srcDoc, firstArtboard[2] - firstArtboard[0], firstArtboard[1] - firstArtboard[3], null)
                            );
                        }
                        if (filteredSelection.length > 0) {
                            srcDoc.selection = filteredSelection;
                            app.copy();
                            app.activeDocument = newDoc;
                            app.paste();
                            placePastedGroup(placementContext, labelName);
                        }
                    }
                } finally {
                    if (lockState) restoreLockHiddenState(lockState);                 // 元の状態へ戻す / Restore the original state
                    if (shouldCloseSource) srcDoc.close(SaveOptions.DONOTSAVECHANGES); // 一時ファイルを閉じる / Close the temp file
                }

                progressBar.value = j + 1;
                processedCountStatic.text = labelText('progress.count') + (j + 1) + "/" + originalDocsLength;
                progressWin.update();
            }

            // 全セルを正方形に近いグリッドへ並べ、キャンバス中央に配置
            // Lay out all cells into a near-square grid centered on the canvas
            layoutCellsAsCenteredGrid(placementContext);

            // グループごとにアートボードを追加したので、初期の仮アートボードを削除（新規ドキュメント時のみ）。
            // 現在のドキュメントに読み込む場合は既存アートボードなので削除しない。
            // Each group added its own artboard, so remove the initial placeholder (new document only).
            // When importing into the current document, artboards[0] is the user's existing one — keep it.
            if (!useCurrentDoc && placementContext.placedCount > 0 && newDoc.artboards.length > placementContext.placedCount) {
                newDoc.artboards.remove(0);
            }
            app.activeDocument = newDoc;
        } finally {
            // ScriptUI パレットの close() は Illustrator で稀に "an Illustrator error occurred (MRAP)" を
            // 投げることがある（処理自体は完了済み）。spurious なエラーで全体を落とさないよう保護する。
            // また本体側で起きた本物の例外を close() のエラーで上書きしないためにも try/catch で囲む。
            // The palette's close() occasionally throws a spurious MRAP error in Illustrator even though the
            // work is done; guard it so it neither aborts the run nor masks a real error from the loop body.
            try { progressWin.hide(); } catch (hideErr) { }
            try { progressWin.close(); } catch (closeErr) { }
        }

        // すべてのアートボードがウィンドウに収まるように表示 / Fit all artboards in the window
        app.executeMenuCommand('fitall');

        // キャンセルされた場合はエラーではなく通常のメッセージで知らせる
        // If cancelled, inform the user with a normal message (not an error)
        if (userCancelled) {
            alert(getLabel('alert.cancelled'));
        } else if (placementContext.placedCount === 0) {
            // 1件も配置できなかった場合（対象アートボード番号が存在しない・内容が空など）も無言で終わらせない
            // Don't finish silently when nothing was placed (e.g. target artboard numbers don't exist, or empty content)
            alert(getLabel('alert.noArtboardImported'));
        }
    })();

})();
