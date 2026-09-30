#target illustrator
#targetengine "RectangularGridReverseToolEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した水平線・垂直線を解析し、長方形グリッドとして再構成します。
不揃いな罫線や、結合セルを含むレイアウトを整理し、整った格子構造に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangularGridReverseTool.md

### Overview

Analyzes the selected horizontal and vertical lines and reconstructs them into a rectangular grid.
Uneven rules and layouts containing merged cells are tidied into a regular lattice.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangularGridReverseTool.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RectangularGridReverseTool";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangularGridReverseTool.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangularGridReverseTool.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グリッド再構築", en: "Rebuild Grid" }
        },
        panel: {
            preprocessing: { ja: "前処理", en: "Pre-processing" },
            distribution: { ja: "配置モード", en: "Distribution Mode" },
            equalize: { ja: "均等配置の対象", en: "Equalize Targets" },
            line: { ja: "線（後処理）", en: "Line (post-process)" },
            strokeWidth: { ja: "線幅", en: "Stroke Width" },
            postProcessing: { ja: "後処理", en: "Post-processing" }
        },
        radio: {
            distributionNone: { ja: "均等配置しない", en: "Do not distribute evenly" },
            distributionEven: { ja: "均等に（強制）", en: "Evenly (force)" },
            distributionEvenMergedCell: { ja: "均等＋結合セル対応", en: "Evenly + merged cells" },
            strokeWidthMax: { ja: "最大", en: "Maximum" },
            strokeWidthMin: { ja: "最小", en: "Minimum" },
            strokeWidthAverage: { ja: "平均", en: "Average" },
            strokeWidthSpecified: { ja: "指定", en: "Specified" }
        },
        checkbox: {
            splitFrameToFourSides: { ja: "外枠を四辺に分割", en: "Split outer frame into four sides" },
            equalizeVertical: { ja: "縦罫", en: "Vertical lines" },
            equalizeHorizontal: { ja: "横罫", en: "Horizontal lines" },
            lockFirstColumn: { ja: "1列目を固定", en: "Lock first column" },
            lockFirstRow: { ja: "1行目を固定", en: "Lock first row" },
            projectingCap: { ja: "突出線端にする", en: "Projecting cap" },
            convertDashedToSolid: { ja: "破線を実線にする", en: "Convert dashed lines to solid" },
            frameToRect: { ja: "外枠を長方形に変換", en: "Convert outer frame to rectangle" },
            centerPointTextVertically: { ja: "ポイント文字をセル内で上下中央", en: "Center point text vertically in cells" },
            grouping: { ja: "グループ化", en: "Group" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            splitFrameToFourSides: {
                ja: "外枠の長方形を、上下左右4本の線に分割します。",
                en: "Splits the outer rectangle into four separate lines."
            },
            distributionNone: { ja: "行・列の間隔は変えません。", en: "Leaves the row and column spacing alone." },
            distributionEven: { ja: "行・列の間隔を均等にそろえます。", en: "Evens out the row and column spacing." },
            distributionEvenMergedCell: {
                ja: "結合セルを考慮しながら、行・列の間隔を均等にそろえます。",
                en: "Evens out the spacing while respecting merged cells."
            },
            equalizeVertical: { ja: "列の幅をそろえます。", en: "Gives the columns the same width." },
            lockFirstColumn: {
                ja: "1列目の幅は変えずに、残りをそろえます。",
                en: "Keeps the first column as it is and evens out the rest."
            },
            equalizeHorizontal: { ja: "行の高さをそろえます。", en: "Gives the rows the same height." },
            lockFirstRow: {
                ja: "1行目の高さは変えずに、残りをそろえます。",
                en: "Keeps the first row as it is and evens out the rest."
            },
            projectingCap: {
                ja: "線の端を太さの半分だけ延ばして、角の隙間をなくします。",
                en: "Extends the line ends by half the weight so the corners close up."
            },
            convertDashedToSolid: { ja: "点線・破線を実線に変えます。", en: "Turns dashed lines into solid ones." },
            strokeWidthMax: { ja: "いちばん太い線に合わせます。", en: "Matches the thickest line." },
            strokeWidthMin: { ja: "いちばん細い線に合わせます。", en: "Matches the thinnest line." },
            strokeWidthAverage: { ja: "線幅の平均に合わせます。", en: "Matches the average line weight." },
            strokeWidthSpecified: { ja: "線幅を数値で指定します。", en: "Sets the line weight to a value you type." },
            frameToRect: {
                ja: "外周の4本の線を1つの長方形にまとめます。",
                en: "Merges the four outer lines back into a single rectangle."
            },
            centerPointTextVertically: {
                ja: "ポイント文字をセルの天地中央に置き直します。",
                en: "Re-centers point text vertically within its cell."
            },
            grouping: { ja: "できあがった表を1つのグループにまとめます。", en: "Groups the finished table together." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            },
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
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noSelection: { ja: "水平線と垂直線を選択してください。", en: "Select horizontal and vertical lines." },
            notEnoughLines: {
                ja: "格子を作るには、最低でも2本の水平線と2本の垂直線が必要です。",
                en: "To create a grid, select at least two horizontal lines and two vertical lines."
            }
        }
    };

    // =========================================
    // 単位ユーティリティ / Unit utilities
    // =========================================

    /**
     * Illustrator 単位ユーティリティ関数群 / Illustrator unit utility functions
     *
     * 設定キーの意味 / Preference keys:
     * - "rulerType"         ：一般（定規の単位）/ General ruler unit
     * - "strokeUnits"       ：線 / Stroke unit
     * - "text/units"        ：文字 / Text unit
     * - "text/asianunits"   ：東アジア言語のオプション / East Asian text unit
     */

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

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した罫線を解析し、ダイアログの設定でグリッドに組み直す
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var selectedItems = app.activeDocument.selection;
        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        // グループ等を再帰展開して PathItem の平坦リストにする / Flatten selection (descending into groups) to PathItems
        var flatPathItems = collectSelectionPathItems(selectedItems);

        // 塗りのあるパスは対象外 / Exclude paths that have a fill
        flatPathItems = excludeFilledPaths(flatPathItems);

        // 軸並行の長方形は4本の罫線に分解 / Decompose axis-aligned rectangles into 4 lines
        var expansionResult = expandRectanglesInSelection(flatPathItems);
        var rectangleExpansions = expansionResult.expansions;

        var classified = classifySelectedStraightLines(expansionResult.items);
        if (classified.horizontalLines.length < 2 || classified.verticalLines.length < 2) {
            restoreRectangleExpansions(rectangleExpansions);
            alert(getLabel("alert.notEnoughLines"));
            return;
        }

        var gridBounds = getGridBounds(classified.horizontalLines, classified.verticalLines);
        var originalLineStates = captureLineStates(classified.horizontalLines, classified.verticalLines);
        var strokeUnitInfo = getUnitInfo("strokeUnits");

        var hasExpandedRectangles = rectangleExpansions.length > 0;
        var dialogOptions = showOptionDialog(classified, gridBounds, originalLineStates, strokeUnitInfo, hasExpandedRectangles);
        if (!dialogOptions) {
            restoreRectangleExpansions(rectangleExpansions);
            return;
        }

        // 前処理：外枠を四辺に分割しない場合は分解を取り消して終了 / If split is disabled, undo expansion and exit
        if (!dialogOptions.splitOuterFrame && hasExpandedRectangles) {
            restoreRectangleExpansions(rectangleExpansions);
            return;
        }

        restoreLineStates(originalLineStates);
        applyLineOptions(classified, gridBounds, dialogOptions);

        // ポイント文字をセル内で上下中央に配置 / Center point text vertically within cells
        if (dialogOptions.centerPointTextVertically) {
            centerPointTextVerticallyInCells(classified.horizontalLines, gridBounds);
        }

        // 外枠を長方形に変換 → 残存パス＋長方形をグループ化 / Convert outer frame to rectangle, then group remaining paths and rectangle
        var groupTargets = collectLinePaths(classified.horizontalLines.concat(classified.verticalLines));
        if (dialogOptions.frameToRect) {
            groupTargets = replaceOuterLinesWithRectangle(classified, gridBounds, groupTargets);
        }

        if (dialogOptions.group && groupTargets.length > 1) {
            groupProcessedItems(groupTargets);
        }
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
    // 選択の前処理 / Preparing the selection
    // =========================================

    /**
     * 塗りのあるパスを対象から除外する
     * @param {PathItem[]} pathItems - 対象のパス
     * @returns {PathItem[]} 塗りのないパス
     */
    function excludeFilledPaths(pathItems) {
        var unfilledPaths = [];
        for (var itemIndex = 0; itemIndex < pathItems.length; itemIndex++) {
            if (!pathItems[itemIndex].filled) unfilledPaths.push(pathItems[itemIndex]);
        }
        return unfilledPaths;
    }

    /**
     * 軸に平行な長方形（閉じた4点で、各辺が水平または垂直）かどうかを判定する
     * @param {PathItem} pathItem - 判定するパス
     * @returns {boolean} 軸に平行な長方形なら true
     */
    function isAxisAlignedRectangle(pathItem) {
        if (!pathItem || pathItem.typename !== "PathItem") return false;
        if (!pathItem.closed) return false;
        if (pathItem.pathPoints.length !== 4) return false;
        var rectangleTolerance = 0.01;
        for (var pointIndex = 0; pointIndex < 4; pointIndex++) {
            var currentAnchor = pathItem.pathPoints[pointIndex].anchor;
            var nextAnchor = pathItem.pathPoints[(pointIndex + 1) % 4].anchor;
            var horizontalDelta = Math.abs(currentAnchor[0] - nextAnchor[0]);
            var verticalDelta = Math.abs(currentAnchor[1] - nextAnchor[1]);
            if (horizontalDelta > rectangleTolerance && verticalDelta > rectangleTolerance) return false;
            if (horizontalDelta < rectangleTolerance && verticalDelta < rectangleTolerance) return false;
        }
        return true;
    }

    /**
     * 長方形パスのアンカーから外接座標を求める
     * @param {PathItem} pathItem - 長方形のパス
     * @returns {{minX: number, maxX: number, minY: number, maxY: number}} 外接座標
     */
    function getRectangleBoundsFromPath(pathItem) {
        var rectangleMinX = pathItem.pathPoints[0].anchor[0], rectangleMaxX = rectangleMinX;
        var rectangleMinY = pathItem.pathPoints[0].anchor[1], rectangleMaxY = rectangleMinY;
        for (var pointIndex = 1; pointIndex < pathItem.pathPoints.length; pointIndex++) {
            var anchor = pathItem.pathPoints[pointIndex].anchor;
            if (anchor[0] < rectangleMinX) rectangleMinX = anchor[0];
            if (anchor[0] > rectangleMaxX) rectangleMaxX = anchor[0];
            if (anchor[1] < rectangleMinY) rectangleMinY = anchor[1];
            if (anchor[1] > rectangleMaxY) rectangleMaxY = anchor[1];
        }
        return { minX: rectangleMinX, maxX: rectangleMaxX, minY: rectangleMinY, maxY: rectangleMaxY };
    }

    /**
     * パスの塗り・線の属性を記録する
     * @param {PathItem} pathItem - 記録するパス
     * @returns {Object} 塗り・線の属性
     */
    function capturePathAppearance(pathItem) {
        var appearance = {
            filled: pathItem.filled,
            stroked: pathItem.stroked
        };
        if (pathItem.filled) appearance.fillColor = pathItem.fillColor;
        if (pathItem.stroked) {
            appearance.strokeColor = pathItem.strokeColor;
            appearance.strokeWidth = pathItem.strokeWidth;
            appearance.strokeCap = pathItem.strokeCap;
            appearance.strokeJoin = pathItem.strokeJoin;
            appearance.strokeMiterLimit = pathItem.strokeMiterLimit;
            appearance.strokeDashes = pathItem.strokeDashes;
            appearance.strokeDashOffset = pathItem.strokeDashOffset;
        }
        return appearance;
    }

    /**
     * 記録した塗り・線の属性をパスに適用する
     * @param {PathItem} pathItem - 適用先のパス
     * @param {Object} appearance - capturePathAppearance() の戻り値
     * @returns {void}
     */
    function applyPathAppearance(pathItem, appearance) {
        pathItem.filled = appearance.filled;
        if (appearance.filled) pathItem.fillColor = appearance.fillColor;
        pathItem.stroked = appearance.stroked;
        if (appearance.stroked) {
            pathItem.strokeColor = appearance.strokeColor;
            pathItem.strokeWidth = appearance.strokeWidth;
            pathItem.strokeCap = appearance.strokeCap;
            pathItem.strokeJoin = appearance.strokeJoin;
            pathItem.strokeMiterLimit = appearance.strokeMiterLimit;
            pathItem.strokeDashes = appearance.strokeDashes;
            pathItem.strokeDashOffset = appearance.strokeDashOffset;
        }
    }

    /**
     * 線の属性を引き継いだ2点の直線パスを作る（塗りなし）
     * @param {Object} parentContainer - 作成先のレイヤーまたはグループ
     * @param {number[]} firstPoint - 始点 [x, y]
     * @param {number[]} secondPoint - 終点 [x, y]
     * @param {Object} appearance - capturePathAppearance() の戻り値
     * @returns {PathItem} 作成した直線
     */
    function createLineFromAppearance(parentContainer, firstPoint, secondPoint, appearance) {
        var newLine = parentContainer.pathItems.add();
        newLine.setEntirePath([firstPoint, secondPoint]);
        newLine.closed = false;
        applyPathAppearance(newLine, appearance);
        newLine.filled = false;
        return newLine;
    }

    /**
     * 軸に平行な長方形を4本の罫線に分解する（元の長方形は削除）
     * @param {PathItem[]} pathItems - 対象のパス
     * @returns {{items: PathItem[], expansions: Object[]}} 分解後のパスと、取り消し用の記録
     */
    function expandRectanglesInSelection(pathItems) {
        var expandedItems = [];
        var rectangleExpansions = [];
        for (var itemIndex = 0; itemIndex < pathItems.length; itemIndex++) {
            var sourceItem = pathItems[itemIndex];
            if (!isAxisAlignedRectangle(sourceItem)) {
                expandedItems.push(sourceItem);
                continue;
            }
            var rectangleBounds = getRectangleBoundsFromPath(sourceItem);
            var appearance = capturePathAppearance(sourceItem);
            var parentContainer = sourceItem.parent;
            var topLeft = [rectangleBounds.minX, rectangleBounds.maxY];
            var topRight = [rectangleBounds.maxX, rectangleBounds.maxY];
            var bottomLeft = [rectangleBounds.minX, rectangleBounds.minY];
            var bottomRight = [rectangleBounds.maxX, rectangleBounds.minY];
            var generatedLines = [
                createLineFromAppearance(parentContainer, topLeft, topRight, appearance),
                createLineFromAppearance(parentContainer, bottomLeft, bottomRight, appearance),
                createLineFromAppearance(parentContainer, bottomLeft, topLeft, appearance),
                createLineFromAppearance(parentContainer, bottomRight, topRight, appearance)
            ];
            for (var generatedIndex = 0; generatedIndex < generatedLines.length; generatedIndex++) {
                expandedItems.push(generatedLines[generatedIndex]);
            }
            rectangleExpansions.push({
                bounds: rectangleBounds,
                appearance: appearance,
                parent: parentContainer,
                generatedLines: generatedLines
            });
            /* 削除できない場合も分解は続ける / keep going even if the rectangle cannot be removed */
            try { sourceItem.remove(); } catch (rectangleRemoveError) { }
        }
        return { items: expandedItems, expansions: rectangleExpansions };
    }

    /**
     * 長方形の分解を取り消す（キャンセル時用）
     * @param {Object[]} rectangleExpansions - expandRectanglesInSelection() の記録
     * @returns {void}
     */
    function restoreRectangleExpansions(rectangleExpansions) {
        for (var expansionIndex = 0; expansionIndex < rectangleExpansions.length; expansionIndex++) {
            var expansion = rectangleExpansions[expansionIndex];
            for (var generatedLineIndex = 0; generatedLineIndex < expansion.generatedLines.length; generatedLineIndex++) {
                /* 既に消えている線は飛ばす / skip lines that are already gone */
                try { expansion.generatedLines[generatedLineIndex].remove(); } catch (lineRemoveError) { }
            }
            var rectangleBounds = expansion.bounds;
            var restoredRectangle = expansion.parent.pathItems.rectangle(
                rectangleBounds.maxY, rectangleBounds.minX,
                rectangleBounds.maxX - rectangleBounds.minX, rectangleBounds.maxY - rectangleBounds.minY
            );
            applyPathAppearance(restoredRectangle, expansion.appearance);
        }
    }

    // =========================================
    // 罫線の分類と外周 / Line classification and grid bounds
    // =========================================

    /**
     * アンカー2点の直線パスを抜き出し、水平線と垂直線に分類する
     * @param {PathItem[]} pathItems - 対象のパス
     * @returns {{horizontalLines: Object[], verticalLines: Object[]}} 水平線（path, y, minX, maxX）と垂直線（path, x, minY, maxY）
     */
    function classifySelectedStraightLines(pathItems) {
        var horizontalLines = [];
        var verticalLines = [];
        for (var itemIndex = 0; itemIndex < pathItems.length; itemIndex++) {
            var pathItem = pathItems[itemIndex];
            if (!pathItem || pathItem.typename !== "PathItem" || pathItem.pathPoints.length !== 2) continue;
            var firstAnchor = pathItem.pathPoints[0].anchor;
            var secondAnchor = pathItem.pathPoints[1].anchor;
            var horizontalDistance = Math.abs(firstAnchor[0] - secondAnchor[0]);
            var verticalDistance = Math.abs(firstAnchor[1] - secondAnchor[1]);
            if (horizontalDistance >= verticalDistance) {
                horizontalLines.push({
                    path: pathItem,
                    y: (firstAnchor[1] + secondAnchor[1]) / 2,
                    minX: Math.min(firstAnchor[0], secondAnchor[0]),
                    maxX: Math.max(firstAnchor[0], secondAnchor[0])
                });
            } else {
                verticalLines.push({
                    path: pathItem,
                    x: (firstAnchor[0] + secondAnchor[0]) / 2,
                    minY: Math.min(firstAnchor[1], secondAnchor[1]),
                    maxY: Math.max(firstAnchor[1], secondAnchor[1])
                });
            }
        }
        return { horizontalLines: horizontalLines, verticalLines: verticalLines };
    }

    /**
     * 格子の外周座標（左端・右端・上端・下端）を計算する
     * 上下左右に貫通する罫線（最長線の90%以上の長さ）が複数あればそれを基準とし、
     * それより外側にはみ出した線は計算対象から除外する
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {{minX: number, maxX: number, minY: number, maxY: number}} 外周座標
     */
    function getGridBounds(horizontalLines, verticalLines) {
        var xRange = getValueRange(pickSpanningLines(verticalLines, "minY", "maxY"), "x");
        var yRange = getValueRange(pickSpanningLines(horizontalLines, "minX", "maxX"), "y");
        return { minX: xRange.min, maxX: xRange.max, minY: yRange.min, maxY: yRange.max };
    }

    /**
     * 最長線の 90% 以上の長さがある「貫通線」を選ぶ。2本未満なら全線を返す
     * @param {Object[]} lines - 線の記録
     * @param {string} startKey - 線の始端のプロパティ名（minX / minY）
     * @param {string} endKey - 線の終端のプロパティ名（maxX / maxY）
     * @returns {Object[]} 外枠の候補にする線
     */
    function pickSpanningLines(lines, startKey, endKey) {
        var SPAN_RATIO_THRESHOLD = 0.9;

        var maxSpan = 0;
        for (var spanIndex = 0; spanIndex < lines.length; spanIndex++) {
            var spanLength = lines[spanIndex][endKey] - lines[spanIndex][startKey];
            if (spanLength > maxSpan) maxSpan = spanLength;
        }

        var spanningLines = [];
        var spanThreshold = maxSpan * SPAN_RATIO_THRESHOLD;
        for (var pickIndex = 0; pickIndex < lines.length; pickIndex++) {
            if ((lines[pickIndex][endKey] - lines[pickIndex][startKey]) >= spanThreshold) {
                spanningLines.push(lines[pickIndex]);
            }
        }

        // 貫通線が2本以上あれば外枠候補として採用、なければ全線をフォールバック
        // / Use spanning lines as outer-frame candidates when at least 2 exist; otherwise fall back to all lines
        return spanningLines.length >= 2 ? spanningLines : lines;
    }

    /**
     * 線の記録から、指定プロパティの最小値と最大値を求める
     * @param {Object[]} lines - 線の記録（1本以上）
     * @param {string} propertyName - 調べるプロパティ名（x / y）
     * @returns {{min: number, max: number}} 最小値と最大値
     */
    function getValueRange(lines, propertyName) {
        var minValue = lines[0][propertyName], maxValue = lines[0][propertyName];
        for (var lineIndex = 1; lineIndex < lines.length; lineIndex++) {
            if (lines[lineIndex][propertyName] < minValue) minValue = lines[lineIndex][propertyName];
            if (lines[lineIndex][propertyName] > maxValue) maxValue = lines[lineIndex][propertyName];
        }
        return { min: minValue, max: maxValue };
    }

    /**
     * 線の記録からパスだけを取り出す
     * @param {Object[]} lines - 線の記録
     * @returns {PathItem[]} パス
     */
    function collectLinePaths(lines) {
        var linePaths = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) linePaths.push(lines[lineIndex].path);
        return linePaths;
    }

    /**
     * 線の記録から、指定プロパティの値を順に取り出す
     * @param {Object[]} lines - 線の記録（または座標グループ）
     * @param {string} propertyName - 取り出すプロパティ名
     * @returns {number[]} 値の配列
     */
    function collectLineValues(lines, propertyName) {
        var lineValues = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) lineValues.push(lines[lineIndex][propertyName]);
        return lineValues;
    }

    /**
     * 配列に同じ参照が含まれているかどうか
     * @param {Array} itemList - 調べる配列
     * @param {Object} targetItem - 探す参照
     * @returns {boolean} 含まれていれば true
     */
    function containsItem(itemList, targetItem) {
        for (var itemIndex = 0; itemIndex < itemList.length; itemIndex++) {
            if (itemList[itemIndex] === targetItem) return true;
        }
        return false;
    }

    // =========================================
    // 線の状態の記録と復元 / Capturing and restoring line states
    // =========================================

    /**
     * 2点の直線パスの両端を設定する（方向ハンドルもアンカー位置に揃える）
     * @param {PathItem} pathItem - 2点のパス
     * @param {number[]} firstPoint - 始点 [x, y]
     * @param {number[]} secondPoint - 終点 [x, y]
     * @returns {void}
     */
    function setLineEndpoints(pathItem, firstPoint, secondPoint) {
        pathItem.pathPoints[0].anchor = firstPoint;
        pathItem.pathPoints[0].leftDirection = firstPoint;
        pathItem.pathPoints[0].rightDirection = firstPoint;
        pathItem.pathPoints[1].anchor = secondPoint;
        pathItem.pathPoints[1].leftDirection = secondPoint;
        pathItem.pathPoints[1].rightDirection = secondPoint;
    }

    /**
     * プレビュー復元用に、水平線・垂直線の座標と線の属性を記録する
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {Object[]} 線ごとの状態
     */
    function captureLineStates(horizontalLines, verticalLines) {
        var allLines = horizontalLines.concat(verticalLines);
        var lineStates = [];
        for (var lineIndex = 0; lineIndex < allLines.length; lineIndex++) {
            var pathItem = allLines[lineIndex].path;
            lineStates.push({
                path: pathItem,
                firstAnchor: pathItem.pathPoints[0].anchor,
                firstLeftDirection: pathItem.pathPoints[0].leftDirection,
                firstRightDirection: pathItem.pathPoints[0].rightDirection,
                secondAnchor: pathItem.pathPoints[1].anchor,
                secondLeftDirection: pathItem.pathPoints[1].leftDirection,
                secondRightDirection: pathItem.pathPoints[1].rightDirection,
                stroked: pathItem.stroked,
                strokeCap: pathItem.strokeCap,
                strokeWidth: pathItem.strokeWidth,
                strokeDashes: pathItem.strokeDashes,
                strokeDashOffset: pathItem.strokeDashOffset
            });
        }
        return lineStates;
    }

    /**
     * 記録した線の状態を復元する
     * @param {Object[]} lineStates - captureLineStates() の戻り値
     * @returns {void}
     */
    function restoreLineStates(lineStates) {
        for (var lineIndex = 0; lineIndex < lineStates.length; lineIndex++) {
            var lineState = lineStates[lineIndex];
            var pathItem = lineState.path;
            pathItem.pathPoints[0].anchor = lineState.firstAnchor;
            pathItem.pathPoints[0].leftDirection = lineState.firstLeftDirection;
            pathItem.pathPoints[0].rightDirection = lineState.firstRightDirection;
            pathItem.pathPoints[1].anchor = lineState.secondAnchor;
            pathItem.pathPoints[1].leftDirection = lineState.secondLeftDirection;
            pathItem.pathPoints[1].rightDirection = lineState.secondRightDirection;
            pathItem.stroked = lineState.stroked;
            pathItem.strokeCap = lineState.strokeCap;
            pathItem.strokeWidth = lineState.strokeWidth;
            pathItem.strokeDashes = lineState.strokeDashes;
            pathItem.strokeDashOffset = lineState.strokeDashOffset;
        }
    }

    // =========================================
    // プレビューとダイアログの値 / Preview and dialog values
    // =========================================

    /**
     * 線幅・破線・整列をまとめて適用する（プレビューと確定で共通）
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function applyLineOptions(classified, gridBounds, dialogOptions) {
        applyRepresentativeStrokeWidth(classified.horizontalLines, classified.verticalLines, dialogOptions);
        if (dialogOptions.convertDashedToSolid) convertDashedLinesToSolid(classified.horizontalLines, classified.verticalLines);
        alignLinesToGridBounds(classified.horizontalLines, classified.verticalLines, gridBounds, dialogOptions);
    }

    /**
     * プレビューを更新する（元に戻してから、プレビューと外枠分割が ON のときだけ整列する）
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {Object[]} originalLineStates - 元の線の状態
     * @returns {void}
     */
    function updatePreviewFromDialogState(dialogControls, classified, gridBounds, originalLineStates) {
        restoreLineStates(originalLineStates);
        // 外枠を四辺に分割が OFF の場合、本処理（罫線整列）を行わず元の状態のまま表示 / If split is OFF, skip the main alignment and show original state
        if (dialogControls.previewCheckbox.value && dialogControls.splitOuterFrameCheckbox.value) {
            var previewOptions = readOptionDialogState(dialogControls);
            // プレビュー時は frameToRect, group を適用しない / frameToRect and group are not applied in preview
            previewOptions.frameToRect = false;
            previewOptions.group = false;
            applyLineOptions(classified, gridBounds, previewOptions);
        }
        app.redraw();
    }

    /**
     * ダイアログの現在値を読み取る
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @returns {Object} ダイアログの設定
     */
    function readOptionDialogState(dialogControls) {
        var distributionMode = dialogControls.distributionNoneRadio.value ? "none"
            : dialogControls.distributionEvenMergedCellRadio.value ? "evenMergedCell"
                : "even";
        var isEvenMode = distributionMode !== "none";
        return {
            distributionMode: distributionMode,
            even: isEvenMode,
            evenHorizontal: isEvenMode && dialogControls.equalizeHorizontalCheckbox.value,
            evenVertical: isEvenMode && dialogControls.equalizeVerticalCheckbox.value,
            projectingCap: dialogControls.projectingCapCheckbox.value,
            convertDashedToSolid: dialogControls.convertDashedToSolidCheckbox.value,
            strokeWidthMode: dialogControls.strokeWidthMaxRadio.value ? "max"
                : dialogControls.strokeWidthMinRadio.value ? "min"
                    : dialogControls.strokeWidthSpecifiedRadio.value ? "specified"
                        : "average",
            specifiedStrokeWidthPt: dialogControls.strokeWidthSpecifiedRadio.value
                ? readNumericText(dialogControls.strokeWidthInput.text, 0) * dialogControls.strokeUnitInfo.pointsPerUnit
                : 0,
            lockFirstColumn: dialogControls.lockFirstColumnCheckbox.value,
            lockFirstRow: dialogControls.lockFirstRowCheckbox.value,
            frameToRect: dialogControls.frameToRectangleCheckbox.value,
            centerPointTextVertically: dialogControls.centerPointTextVerticallyCheckbox.value,
            group: dialogControls.groupingCheckbox.value,
            splitOuterFrame: dialogControls.splitOuterFrameCheckbox.value
        };
    }

    /**
     * 数値の文字列を読む（カンマは小数点として扱う）
     * @param {string} text - 入力された文字列
     * @param {number} fallbackValue - 数値にならないときの値
     * @returns {number} 読み取った数値
     */
    function readNumericText(text, fallbackValue) {
        var normalizedText = String(text).replace(/,/g, ".");
        var value = parseFloat(normalizedText);
        if (isNaN(value)) return fallbackValue;
        return value;
    }

    /**
     * 「指定」を選んだときだけ線幅の入力欄と∧∨を有効にする
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @returns {void}
     */
    function syncSpecifiedStrokeWidthInput(dialogControls) {
        var isSpecified = dialogControls.strokeWidthSpecifiedRadio.value;
        dialogControls.strokeWidthInput.enabled = isSpecified;
        dialogControls.strokeWidthInput.stepperGroup.enabled = isSpecified;
        redrawSteppersIn(dialogControls.strokeWidthInput.stepperGroup);
    }

    // =========================================
    // 線幅と破線 / Stroke width and dashes
    // =========================================

    /**
     * 代表線幅をすべての線に適用する（代表線幅が 0 以下なら何もしない）
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function applyRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions) {
        var representativeStrokeWidth = getRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions);
        if (representativeStrokeWidth <= 0) return;
        var allLines = horizontalLines.concat(verticalLines);
        for (var lineIndex = 0; lineIndex < allLines.length; lineIndex++) {
            var pathItem = allLines[lineIndex].path;
            pathItem.stroked = true;
            pathItem.strokeWidth = representativeStrokeWidth;
        }
    }

    /**
     * 設定に応じた代表線幅（最大・最小・平均・指定）を返す
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {number} 代表線幅（pt）。線のある線が無ければ 0
     */
    function getRepresentativeStrokeWidth(horizontalLines, verticalLines, dialogOptions) {
        if (dialogOptions.strokeWidthMode === "specified") {
            return dialogOptions.specifiedStrokeWidthPt;
        }

        var strokeWidths = collectStrokeWidths(horizontalLines.concat(verticalLines));
        if (strokeWidths.length === 0) return 0;

        if (dialogOptions.strokeWidthMode === "max") return Math.max.apply(null, strokeWidths);
        if (dialogOptions.strokeWidthMode === "min") return Math.min.apply(null, strokeWidths);

        var totalStrokeWidth = 0;
        for (var averageIndex = 0; averageIndex < strokeWidths.length; averageIndex++) {
            totalStrokeWidth += strokeWidths[averageIndex];
        }
        return totalStrokeWidth / strokeWidths.length;
    }

    /**
     * 線のある線から線幅を集める
     * @param {Object[]} lines - 線の記録
     * @returns {number[]} 0 より大きい線幅
     */
    function collectStrokeWidths(lines) {
        var strokeWidths = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var pathItem = lines[lineIndex].path;
            if (pathItem.stroked && pathItem.strokeWidth > 0) strokeWidths.push(pathItem.strokeWidth);
        }
        return strokeWidths;
    }

    /**
     * 破線・点線の設定を外して実線にする
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {void}
     */
    function convertDashedLinesToSolid(horizontalLines, verticalLines) {
        var allLines = horizontalLines.concat(verticalLines);
        for (var lineIndex = 0; lineIndex < allLines.length; lineIndex++) {
            var pathItem = allLines[lineIndex].path;
            pathItem.strokeDashes = [];
            pathItem.strokeDashOffset = 0;
        }
    }

    // =========================================
    // 整列 / Alignment
    // =========================================

    /**
     * 配列要素の指定プロパティが targetValue に最も近いインデックスを返す
     * @param {number} targetValue - 基準の値
     * @param {Object[]} items - 候補（1件以上）
     * @param {string} propertyName - 比べるプロパティ名
     * @returns {number} 最も近い要素のインデックス
     */
    function findClosestIndexByProperty(targetValue, items, propertyName) {
        var closestIndex = 0;
        var closestDistance = Math.abs(targetValue - items[0][propertyName]);
        for (var itemIndex = 1; itemIndex < items.length; itemIndex++) {
            var currentDistance = Math.abs(targetValue - items[itemIndex][propertyName]);
            if (currentDistance < closestDistance) {
                closestIndex = itemIndex;
                closestDistance = currentDistance;
            }
        }
        return closestIndex;
    }

    /* 同じ座標とみなす許容誤差（pt） / Tolerance to treat coordinates as identical (pt) */
    var COORDINATE_TOLERANCE = 5.0;

    /**
     * 線を指定プロパティの近さでグループにまとめる（昇順）
     * @param {Object[]} lines - 線の記録
     * @param {string} propertyName - まとめる基準のプロパティ名（x / y）
     * @param {number} tolerance - 同じ座標とみなす差
     * @returns {{coord: number, lines: Object[]}[]} 座標ごとのグループ（coord はグループ内の平均）
     */
    function groupLinesByCoordinate(lines, propertyName, tolerance) {
        var sortedLines = lines.slice();
        sortedLines.sort(function (lineA, lineB) { return lineA[propertyName] - lineB[propertyName]; });
        var coordinateGroups = [];
        for (var lineIndex = 0; lineIndex < sortedLines.length; lineIndex++) {
            var currentLine = sortedLines[lineIndex];
            var currentCoordinate = currentLine[propertyName];
            if (coordinateGroups.length === 0 || Math.abs(currentCoordinate - coordinateGroups[coordinateGroups.length - 1].coord) > tolerance) {
                coordinateGroups.push({ coord: currentCoordinate, lines: [currentLine] });
            } else {
                var lastCoordinateGroup = coordinateGroups[coordinateGroups.length - 1];
                lastCoordinateGroup.lines.push(currentLine);
                var coordinateSum = 0;
                for (var groupedLineIndex = 0; groupedLineIndex < lastCoordinateGroup.lines.length; groupedLineIndex++) {
                    coordinateSum += lastCoordinateGroup.lines[groupedLineIndex][propertyName];
                }
                lastCoordinateGroup.coord = coordinateSum / lastCoordinateGroup.lines.length;
            }
        }
        return coordinateGroups;
    }

    /**
     * 最小値から最大値までを等間隔に分けた座標を返す
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @param {number} count - 座標の数（2以上）
     * @returns {number[]} 等間隔の座標
     */
    function getEvenlySpacedValues(minValue, maxValue, count) {
        var step = (maxValue - minValue) / (count - 1);
        var evenValues = [];
        for (var valueIndex = 0; valueIndex < count; valueIndex++) {
            evenValues.push(minValue + step * valueIndex);
        }
        return evenValues;
    }

    /**
     * 行（水平線）の新しい Y 座標を決める
     * 1行目を固定するときは、最上と上から2本目を残し、2本目と最下の間を均等に分ける
     * @param {number[]} rowYs - 下から上に並んだ Y 座標
     * @param {Object} gridBounds - 外周座標
     * @param {boolean} shouldEven - 均等にするかどうか（false なら rowYs をそのまま返す）
     * @param {boolean} lockFirstRow - 1行目を固定するかどうか
     * @returns {number[]} 新しい Y 座標
     */
    function resolveRowPositions(rowYs, gridBounds, shouldEven, lockFirstRow) {
        if (!shouldEven) return rowYs;
        var rowCount = rowYs.length;
        if (!(lockFirstRow && rowCount >= 3)) return getEvenlySpacedValues(gridBounds.minY, gridBounds.maxY, rowCount);

        // 1行目（最上）・2行目を固定し、2行目と最下を基準に残りを均等配置 / Lock 1st & 2nd from top; distribute the rest between 2nd-from-top and bottom
        var bottommostY = rowYs[0];
        var lockedAnchorY = rowYs[rowCount - 2];
        var lockedRowStep = (lockedAnchorY - bottommostY) / (rowCount - 2);
        var rowPositions = [];
        for (var rowIndex = 0; rowIndex < rowCount - 1; rowIndex++) {
            rowPositions.push(bottommostY + lockedRowStep * rowIndex);
        }
        rowPositions.push(rowYs[rowCount - 1]);
        return rowPositions;
    }

    /**
     * 列（垂直線）の新しい X 座標を決める
     * 1列目を固定するときは、左端と2本目を残し、2本目と最右の間を均等に分ける
     * @param {number[]} columnXs - 左から右に並んだ X 座標
     * @param {Object} gridBounds - 外周座標
     * @param {boolean} shouldEven - 均等にするかどうか（false なら columnXs をそのまま返す）
     * @param {boolean} lockFirstColumn - 1列目を固定するかどうか
     * @returns {number[]} 新しい X 座標
     */
    function resolveColumnPositions(columnXs, gridBounds, shouldEven, lockFirstColumn) {
        if (!shouldEven) return columnXs;
        var columnCount = columnXs.length;
        if (!(lockFirstColumn && columnCount >= 3)) return getEvenlySpacedValues(gridBounds.minX, gridBounds.maxX, columnCount);

        // 1本目・2本目を固定し、2本目と最右を基準に残りを均等配置 / Lock 1st & 2nd; distribute the rest between 2nd and rightmost
        var lockedAnchorX = columnXs[1];
        var lockedColumnStep = (columnXs[columnCount - 1] - lockedAnchorX) / (columnCount - 2);
        var columnPositions = [columnXs[0]];
        for (var columnIndex = 1; columnIndex < columnCount; columnIndex++) {
            columnPositions.push(lockedAnchorX + lockedColumnStep * (columnIndex - 1));
        }
        return columnPositions;
    }

    /**
     * 線の両端を設定する（必要なら先に突出線端にする）
     * @param {PathItem} pathItem - 2点のパス
     * @param {number[]} firstPoint - 始点 [x, y]
     * @param {number[]} secondPoint - 終点 [x, y]
     * @param {boolean} projectingCap - 突出線端にするかどうか
     * @returns {void}
     */
    function placeGridLine(pathItem, firstPoint, secondPoint, projectingCap) {
        if (projectingCap) pathItem.strokeCap = StrokeCap.PROJECTINGENDCAP;
        setLineEndpoints(pathItem, firstPoint, secondPoint);
    }

    /**
     * 水平線・垂直線を格子に揃える
     * @param {Object[]} horizontalLines - 水平線（均等にするときは Y の昇順に並べ替える）
     * @param {Object[]} verticalLines - 垂直線（均等にするときは X の昇順に並べ替える）
     * @param {Object} gridBounds - 外周座標
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function alignLinesToGridBounds(horizontalLines, verticalLines, gridBounds, dialogOptions) {
        if (dialogOptions.distributionMode === "evenMergedCell") {
            alignLinesWithMergedCellSupport(horizontalLines, verticalLines, gridBounds, dialogOptions);
            return;
        }

        // 水平線の Y 位置を決定（均等配置 or 元の Y）/ Determine Y positions for horizontal lines
        var shouldEvenRows = dialogOptions.evenHorizontal && horizontalLines.length >= 2;
        if (shouldEvenRows) horizontalLines.sort(function (lineA, lineB) { return lineA.y - lineB.y; });
        var horizontalLineYPositions = resolveRowPositions(collectLineValues(horizontalLines, "y"), gridBounds, shouldEvenRows, dialogOptions.lockFirstRow);

        // 垂直線の X 位置を決定（均等配置 or 元の X）/ Determine X positions for vertical lines
        var shouldEvenColumns = dialogOptions.evenVertical && verticalLines.length >= 2;
        if (shouldEvenColumns) verticalLines.sort(function (lineA, lineB) { return lineA.x - lineB.x; });
        var verticalLineXPositions = resolveColumnPositions(collectLineValues(verticalLines, "x"), gridBounds, shouldEvenColumns, dialogOptions.lockFirstColumn);

        // Illustrator は Y 上方向が正 / Illustrator uses positive Y upward
        var topCoordinateY = Math.max(gridBounds.minY, gridBounds.maxY);
        var bottomCoordinateY = Math.min(gridBounds.minY, gridBounds.maxY);

        for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            placeGridLine(horizontalLines[horizontalIndex].path,
                [gridBounds.minX, horizontalLineYPositions[horizontalIndex]],
                [gridBounds.maxX, horizontalLineYPositions[horizontalIndex]], dialogOptions.projectingCap);
        }
        for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
            placeGridLine(verticalLines[verticalIndex].path,
                [verticalLineXPositions[verticalIndex], topCoordinateY],
                [verticalLineXPositions[verticalIndex], bottomCoordinateY], dialogOptions.projectingCap);
        }
    }

    /**
     * 結合セル対応モード：同じ座標の線を 1 行／1 列としてまとめ、両端は直近の交差点にスナップする
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @param {Object} gridBounds - 外周座標
     * @param {Object} dialogOptions - ダイアログの設定
     * @returns {void}
     */
    function alignLinesWithMergedCellSupport(horizontalLines, verticalLines, gridBounds, dialogOptions) {
        var horizontalGroups = groupLinesByCoordinate(horizontalLines, "y", COORDINATE_TOLERANCE);
        var verticalGroups = groupLinesByCoordinate(verticalLines, "x", COORDINATE_TOLERANCE);

        // 各グループの新しい位置（行＝Y、列＝X）/ New position for each group (rows: Y, columns: X)
        var horizontalGroupYPositions = resolveRowPositions(collectLineValues(horizontalGroups, "coord"), gridBounds,
            dialogOptions.evenHorizontal && horizontalGroups.length >= 2, dialogOptions.lockFirstRow);
        var verticalGroupXPositions = resolveColumnPositions(collectLineValues(verticalGroups, "coord"), gridBounds,
            dialogOptions.evenVertical && verticalGroups.length >= 2, dialogOptions.lockFirstColumn);

        // 水平線：同一行のすべてに同じ Y、両端は最寄り列にスナップ / Apply same Y to all lines in a row; snap endpoints to nearest column
        for (var horizontalGroupIndex = 0; horizontalGroupIndex < horizontalGroups.length; horizontalGroupIndex++) {
            var rowCoordinateY = horizontalGroupYPositions[horizontalGroupIndex];
            var rowLines = horizontalGroups[horizontalGroupIndex].lines;
            for (var rowLineIndex = 0; rowLineIndex < rowLines.length; rowLineIndex++) {
                var horizontalLineData = rowLines[rowLineIndex];
                var leftColumnIndex = findClosestIndexByProperty(horizontalLineData.minX, verticalGroups, "coord");
                var rightColumnIndex = findClosestIndexByProperty(horizontalLineData.maxX, verticalGroups, "coord");
                placeGridLine(horizontalLineData.path,
                    [verticalGroupXPositions[leftColumnIndex], rowCoordinateY],
                    [verticalGroupXPositions[rightColumnIndex], rowCoordinateY], dialogOptions.projectingCap);
            }
        }

        // 垂直線：同一列のすべてに同じ X、両端は最寄り行にスナップ / Apply same X to all lines in a column; snap endpoints to nearest row
        for (var verticalGroupIndex = 0; verticalGroupIndex < verticalGroups.length; verticalGroupIndex++) {
            var columnCoordinateX = verticalGroupXPositions[verticalGroupIndex];
            var columnLines = verticalGroups[verticalGroupIndex].lines;
            for (var columnLineIndex = 0; columnLineIndex < columnLines.length; columnLineIndex++) {
                var verticalLineData = columnLines[columnLineIndex];
                var topRowIndex = findClosestIndexByProperty(verticalLineData.maxY, horizontalGroups, "coord");
                var bottomRowIndex = findClosestIndexByProperty(verticalLineData.minY, horizontalGroups, "coord");
                placeGridLine(verticalLineData.path,
                    [columnCoordinateX, horizontalGroupYPositions[topRowIndex]],
                    [columnCoordinateX, horizontalGroupYPositions[bottomRowIndex]], dialogOptions.projectingCap);
            }
        }
    }

    // =========================================
    // 後処理 / Post-processing
    // =========================================

    /**
     * 外周4本（上下の水平線・左右の垂直線）を選び、重複を除いた参照を返す
     * @param {Object[]} horizontalLines - 水平線
     * @param {Object[]} verticalLines - 垂直線
     * @returns {{uniqueOuterPaths: PathItem[], referencePath: PathItem}} 外周のパスと、線の属性を引き継ぐ基準（上辺）
     */
    function pickOuterPaths(horizontalLines, verticalLines) {
        var topLine = null, bottomLine = null, leftLine = null, rightLine = null;
        for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            if (topLine === null || horizontalLines[horizontalIndex].y > topLine.y) topLine = horizontalLines[horizontalIndex];
            if (bottomLine === null || horizontalLines[horizontalIndex].y < bottomLine.y) bottomLine = horizontalLines[horizontalIndex];
        }
        for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
            if (leftLine === null || verticalLines[verticalIndex].x < leftLine.x) leftLine = verticalLines[verticalIndex];
            if (rightLine === null || verticalLines[verticalIndex].x > rightLine.x) rightLine = verticalLines[verticalIndex];
        }
        var outerPathCandidates = [topLine.path, bottomLine.path, leftLine.path, rightLine.path];
        var uniqueOuterPaths = [];
        for (var candidateIndex = 0; candidateIndex < outerPathCandidates.length; candidateIndex++) {
            if (!containsItem(uniqueOuterPaths, outerPathCandidates[candidateIndex])) uniqueOuterPaths.push(outerPathCandidates[candidateIndex]);
        }
        return { uniqueOuterPaths: uniqueOuterPaths, referencePath: topLine.path };
    }

    /**
     * 外周の長方形を作成する（塗り・線の属性は referencePath から引き継ぐ）
     * @param {PathItem} referencePath - 属性と重ね順の基準にするパス
     * @param {Object} gridBounds - 外周座標
     * @returns {PathItem} 作成した長方形
     */
    function createOuterFrameRectangle(referencePath, gridBounds) {
        var parentLayerOrGroup = referencePath.parent;
        // rectangle(top, left, width, height) — Illustrator は Y 上方向が正 / Illustrator uses positive Y upward
        var frameRectangle = parentLayerOrGroup.pathItems.rectangle(
            gridBounds.maxY, gridBounds.minX,
            gridBounds.maxX - gridBounds.minX, gridBounds.maxY - gridBounds.minY
        );
        applyPathAppearance(frameRectangle, capturePathAppearance(referencePath));
        frameRectangle.move(referencePath, ElementPlacement.PLACEAFTER);
        return frameRectangle;
    }

    /**
     * 外周4本の線を1つの長方形に置き換え、グループ化の対象を差し替える
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {PathItem[]} groupTargets - グループ化の対象（全罫線）
     * @returns {PathItem[]} 外周の線を除き、長方形を加えたグループ化の対象
     */
    function replaceOuterLinesWithRectangle(classified, gridBounds, groupTargets) {
        var outerPaths = pickOuterPaths(classified.horizontalLines, classified.verticalLines);
        var frameRectangle = createOuterFrameRectangle(outerPaths.referencePath, gridBounds);
        // outerPaths.uniqueOuterPaths を groupTargets から除外し frameRectangle を追加 / Exclude outerPaths.uniqueOuterPaths from groupTargets and add frameRectangle
        var innerGridItems = [];
        for (var targetIndex = 0; targetIndex < groupTargets.length; targetIndex++) {
            if (!containsItem(outerPaths.uniqueOuterPaths, groupTargets[targetIndex])) innerGridItems.push(groupTargets[targetIndex]);
        }
        innerGridItems.push(frameRectangle);
        // 参照が無効になる前にフィルタを終え、最後に削除 / Finish filtering before references become invalid, then remove them at the end
        for (var removeIndex = 0; removeIndex < outerPaths.uniqueOuterPaths.length; removeIndex++) {
            /* 削除できない線は残す / leave lines that cannot be removed */
            try { outerPaths.uniqueOuterPaths[removeIndex].remove(); } catch (removeError) { }
        }
        return innerGridItems;
    }

    /**
     * ポイント文字を、中心が含まれる行の上下中央へ移動する
     * @param {Object[]} horizontalLines - 水平線（y は分類時の座標）
     * @param {Object} gridBounds - 外周座標
     * @returns {void}
     */
    function centerPointTextVerticallyInCells(horizontalLines, gridBounds) {
        var rowYs = collectLineValues(horizontalLines, "y");
        rowYs.sort(function (firstY, secondY) { return firstY - secondY; });
        if (rowYs.length < 2) return;

        var documentTextFrames = app.activeDocument.textFrames;
        for (var textIndex = 0; textIndex < documentTextFrames.length; textIndex++) {
            var textFrame = documentTextFrames[textIndex];
            if (textFrame.kind !== TextType.POINTTEXT) continue;
            var textBounds = textFrame.geometricBounds; // [left, top, right, bottom]（Y上方向が正）
            var textCenterX = (textBounds[0] + textBounds[2]) / 2;
            var textCenterY = (textBounds[1] + textBounds[3]) / 2;
            // グリッド外のテキストは対象外 / Skip text outside the grid bounds
            if (textCenterX < gridBounds.minX || textCenterX > gridBounds.maxX) continue;
            if (textCenterY < gridBounds.minY || textCenterY > gridBounds.maxY) continue;
            // テキストの中心が含まれる行を探す / Find the row whose Y range contains the text center
            var rowBottomY = null, rowTopY = null;
            for (var rowIndex = 0; rowIndex < rowYs.length - 1; rowIndex++) {
                if (textCenterY >= rowYs[rowIndex] && textCenterY <= rowYs[rowIndex + 1]) {
                    rowBottomY = rowYs[rowIndex];
                    rowTopY = rowYs[rowIndex + 1];
                    break;
                }
            }
            if (rowBottomY === null) continue;
            var rowCenterY = (rowBottomY + rowTopY) / 2;
            var deltaY = rowCenterY - textCenterY;
            if (Math.abs(deltaY) > 0.001) textFrame.translate(0, deltaY);
        }
    }

    /**
     * 渡されたアイテムを1つのグループにまとめる
     * @param {PageItem[]} processedItems - まとめるアイテム
     * @returns {GroupItem} 作成したグループ
     */
    function groupProcessedItems(processedItems) {
        var parentContainer = processedItems[0].parent;
        var tableGroup = parentContainer.groupItems.add();
        for (var itemIndex = 0; itemIndex < processedItems.length; itemIndex++) {
            processedItems[itemIndex].move(tableGroup, ElementPlacement.PLACEATEND);
        }
        return tableGroup;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 縦並びのオプションパネルを追加する（チェックボックス・ラジオボタンが並ぶので間隔を詰める）
     * @param {Group} parentGroup - 追加先
     * @param {string} labelPath - パネル見出しの LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentGroup, labelPath) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(labelPath));
        setupPanel(optionPanel, 6);
        return optionPanel;
    }

    /**
     * 横並びの行グループを追加する
     * @param {Panel} parentPanel - 追加先
     * @returns {Group} 追加した行
     */
    function addOptionRow(parentPanel) {
        var optionRow = parentPanel.add("group");
        optionRow.orientation = "row";
        optionRow.alignChildren = ["left", "center"];
        return optionRow;
    }

    /**
     * tooltip 付きのチェックボックスを追加する
     * @param {Object} parentContainer - 追加先のパネルまたはグループ
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentContainer, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentContainer.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * tooltip 付きのラジオボタンを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addOptionRadio(parentPanel, labelPath, tooltipPath) {
        var optionRadio = parentPanel.add("radiobutton", undefined, getLabel(labelPath));
        optionRadio.helpTip = getLabel(tooltipPath);
        return optionRadio;
    }

    /**
     * ダイアログを組み立てる
     * @param {{label: string, pointsPerUnit: number}} strokeUnitInfo - 線の単位
     * @param {boolean} hasExpandedRectangles - 外枠の長方形を分解したかどうか
     * @returns {Object} ダイアログ（optionDialog）と各コントロール
     */
    function buildOptionDialog(strokeUnitInfo, hasExpandedRectangles) {
        var optionDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(optionDialog);

        // メインオプション2カラム / Main options in two columns
        var mainOptionsGroup = optionDialog.add("group");
        mainOptionsGroup.orientation = "row";
        mainOptionsGroup.alignChildren = ["fill", "top"];
        mainOptionsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = mainOptionsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "fill";

        var rightColumnGroup = mainOptionsGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "fill";

        // 前処理パネル / Pre-processing panel
        var preprocessingPanel = addOptionPanel(leftColumnGroup, "panel.preprocessing");
        var splitOuterFrameCheckbox = addOptionCheckbox(preprocessingPanel,
            "checkbox.splitFrameToFourSides", "tooltip.splitFrameToFourSides", !!hasExpandedRectangles);

        // 配置ラジオボタン / Distribution options
        var distributionPanel = addOptionPanel(leftColumnGroup, "panel.distribution");
        var distributionNoneRadio = addOptionRadio(distributionPanel, "radio.distributionNone", "tooltip.distributionNone");
        var distributionEvenRadio = addOptionRadio(distributionPanel, "radio.distributionEven", "tooltip.distributionEven");
        var distributionEvenMergedCellRadio = addOptionRadio(distributionPanel, "radio.distributionEvenMergedCell", "tooltip.distributionEvenMergedCell");

        // デフォルトは均等＋結合セル対応 / Default is Evenly + merged cells
        distributionEvenMergedCellRadio.value = true;

        // 平均化パネル（デフォルトは両方 OFF）/ Equalize panel (both OFF by default)
        var equalizePanel = addOptionPanel(leftColumnGroup, "panel.equalize");
        var equalizeVerticalRow = addOptionRow(equalizePanel);
        var equalizeVerticalCheckbox = addOptionCheckbox(equalizeVerticalRow, "checkbox.equalizeVertical", "tooltip.equalizeVertical", false);
        var lockFirstColumnCheckbox = addOptionCheckbox(equalizeVerticalRow, "checkbox.lockFirstColumn", "tooltip.lockFirstColumn", false);
        var equalizeHorizontalRow = addOptionRow(equalizePanel);
        var equalizeHorizontalCheckbox = addOptionCheckbox(equalizeHorizontalRow, "checkbox.equalizeHorizontal", "tooltip.equalizeHorizontal", false);
        var lockFirstRowCheckbox = addOptionCheckbox(equalizeHorizontalRow, "checkbox.lockFirstRow", "tooltip.lockFirstRow", false);

        // 線パネル / Line panel
        var linePanel = addOptionPanel(rightColumnGroup, "panel.line");
        var projectingCapCheckbox = addOptionCheckbox(linePanel, "checkbox.projectingCap", "tooltip.projectingCap", true);
        var convertDashedToSolidCheckbox = addOptionCheckbox(linePanel, "checkbox.convertDashedToSolid", "tooltip.convertDashedToSolid", false);

        // 線幅パネル / Stroke width panel
        var strokeWidthPanel = addOptionPanel(linePanel, "panel.strokeWidth");
        var strokeWidthMaxRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthMax", "tooltip.strokeWidthMax");
        var strokeWidthMinRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthMin", "tooltip.strokeWidthMin");
        var strokeWidthAverageRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthAverage", "tooltip.strokeWidthAverage");
        var strokeWidthSpecifiedRadio = addOptionRadio(strokeWidthPanel, "radio.strokeWidthSpecified", "tooltip.strokeWidthSpecified");

        var strokeWidthSpecifiedInputGroup = addOptionRow(strokeWidthPanel);
        /* ∧∨と入力欄は隙間0で突き合わせる。0以上（onStep は bindOptionDialogEvents() で入れる）
           Butt the stepper against the field; 0 or more */
        var strokeWidthStepperInputGroup = strokeWidthSpecifiedInputGroup.add("group");
        strokeWidthStepperInputGroup.orientation = "row";
        strokeWidthStepperInputGroup.alignChildren = ["left", "center"];
        strokeWidthStepperInputGroup.spacing = 0;
        strokeWidthStepperInputGroup.margins = 0;
        var strokeWidthStepOptions = { step: 1, min: 0 };
        var strokeWidthInput;
        var strokeWidthStepperGroup = addStepper(strokeWidthStepperInputGroup, function () { return strokeWidthInput; }, strokeWidthStepOptions);
        strokeWidthInput = strokeWidthStepperInputGroup.add("edittext", undefined, "0.25");
        strokeWidthInput.helpTip = getLabel("tooltip.strokeWidthSpecified");
        strokeWidthInput.characters = 4;
        strokeWidthInput.stepperGroup = strokeWidthStepperGroup;
        strokeWidthInput.stepOptions = strokeWidthStepOptions;
        bindSteppedArrowKeys(strokeWidthInput, strokeWidthStepperGroup);
        strokeWidthSpecifiedInputGroup.add("statictext", undefined, strokeUnitInfo.label);

        // デフォルトは平均 / Default is average
        strokeWidthAverageRadio.value = true;
        strokeWidthInput.enabled = false;
        strokeWidthStepperGroup.enabled = false;

        // 後処理パネル / Post-processing panel
        var postProcessingPanel = addOptionPanel(leftColumnGroup, "panel.postProcessing");
        var frameToRectangleCheckbox = addOptionCheckbox(postProcessingPanel,
            "checkbox.frameToRect", "tooltip.frameToRect", !!hasExpandedRectangles);
        var centerPointTextVerticallyCheckbox = addOptionCheckbox(postProcessingPanel,
            "checkbox.centerPointTextVertically", "tooltip.centerPointTextVertically", false);
        var groupingCheckbox = addOptionCheckbox(postProcessingPanel, "checkbox.grouping", "tooltip.grouping", false);

        // 下段：左＝プレビュー、右＝ボタン / Bottom row: left=preview, right=buttons
        var buttonRow = addButtonRow(optionDialog);
        var previewCheckbox = addOptionCheckbox(buttonRow.leftGroup, "checkbox.preview", "tooltip.preview", false);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            optionDialog: optionDialog,
            previewCheckbox: previewCheckbox,
            distributionNoneRadio: distributionNoneRadio,
            distributionEvenRadio: distributionEvenRadio,
            distributionEvenMergedCellRadio: distributionEvenMergedCellRadio,
            projectingCapCheckbox: projectingCapCheckbox,
            convertDashedToSolidCheckbox: convertDashedToSolidCheckbox,
            strokeWidthMaxRadio: strokeWidthMaxRadio,
            strokeWidthMinRadio: strokeWidthMinRadio,
            strokeWidthAverageRadio: strokeWidthAverageRadio,
            strokeWidthSpecifiedRadio: strokeWidthSpecifiedRadio,
            strokeWidthInput: strokeWidthInput,
            strokeUnitInfo: strokeUnitInfo,
            frameToRectangleCheckbox: frameToRectangleCheckbox,
            centerPointTextVerticallyCheckbox: centerPointTextVerticallyCheckbox,
            groupingCheckbox: groupingCheckbox,
            splitOuterFrameCheckbox: splitOuterFrameCheckbox,
            equalizeVerticalCheckbox: equalizeVerticalCheckbox,
            equalizeHorizontalCheckbox: equalizeHorizontalCheckbox,
            lockFirstColumnCheckbox: lockFirstColumnCheckbox,
            lockFirstRowCheckbox: lockFirstRowCheckbox
        };
    }

    /**
     * ダイアログのコントロールにクリック等の処理を結び付ける
     * @param {Object} dialogControls - buildOptionDialog() の戻り値
     * @param {Function} refreshPreview - プレビューを更新する関数
     * @returns {void}
     */
    function bindOptionDialogEvents(dialogControls, refreshPreview) {
        var equalizeVerticalCheckbox = dialogControls.equalizeVerticalCheckbox;
        var equalizeHorizontalCheckbox = dialogControls.equalizeHorizontalCheckbox;

        /**
         * 配置モードに応じて平均化パネルの有効／無効を同期する
         * @returns {void}
         */
        function syncEqualizePanelEnabled() {
            var isEqualizeEnabled = !dialogControls.distributionNoneRadio.value;
            equalizeVerticalCheckbox.enabled = isEqualizeEnabled;
            equalizeHorizontalCheckbox.enabled = isEqualizeEnabled;
            dialogControls.lockFirstColumnCheckbox.enabled = isEqualizeEnabled && equalizeVerticalCheckbox.value;
            dialogControls.lockFirstRowCheckbox.enabled = isEqualizeEnabled && equalizeHorizontalCheckbox.value;
        }

        /**
         * 縦罫／横罫のチェックボックスの処理を設定する。option＋クリックでもう一方も同じ状態にする
         * @param {Checkbox} clickedCheckbox - クリックされるチェックボックス
         * @param {Checkbox} otherCheckbox - そろえるチェックボックス
         * @returns {void}
         */
        function bindEqualizeCheckbox(clickedCheckbox, otherCheckbox) {
            clickedCheckbox.onClick = function () {
                if (ScriptUI.environment.keyboardState.altKey) {
                    otherCheckbox.value = clickedCheckbox.value;
                }
                syncEqualizePanelEnabled();
                refreshPreview();
            };
        }

        var onDistributionClick = function () {
            syncEqualizePanelEnabled();
            refreshPreview();
        };
        var onStrokeWidthClick = function () {
            syncSpecifiedStrokeWidthInput(dialogControls);
            refreshPreview();
        };

        syncEqualizePanelEnabled();

        dialogControls.previewCheckbox.onClick = refreshPreview;
        dialogControls.distributionNoneRadio.onClick = onDistributionClick;
        dialogControls.distributionEvenRadio.onClick = onDistributionClick;
        dialogControls.distributionEvenMergedCellRadio.onClick = onDistributionClick;
        bindEqualizeCheckbox(equalizeVerticalCheckbox, equalizeHorizontalCheckbox);
        bindEqualizeCheckbox(equalizeHorizontalCheckbox, equalizeVerticalCheckbox);
        dialogControls.lockFirstColumnCheckbox.onClick = refreshPreview;
        dialogControls.lockFirstRowCheckbox.onClick = refreshPreview;
        dialogControls.projectingCapCheckbox.onClick = refreshPreview;
        dialogControls.convertDashedToSolidCheckbox.onClick = refreshPreview;
        dialogControls.strokeWidthMaxRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthMinRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthAverageRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthSpecifiedRadio.onClick = onStrokeWidthClick;
        dialogControls.strokeWidthInput.onChanging = refreshPreview;
        /* ∧∨・↑↓キーで増減したあと / after stepping with the stepper or arrow keys */
        dialogControls.strokeWidthInput.stepOptions.onStep = function () {
            if (!dialogControls.strokeWidthSpecifiedRadio.value) dialogControls.strokeWidthSpecifiedRadio.value = true;
            syncSpecifiedStrokeWidthInput(dialogControls);
            refreshPreview();
        };
        // 外枠分割トグル：分割済みの線に対して整列処理（プレビュー）を再実行 / Toggle: re-run alignment preview on the already-split lines
        dialogControls.splitOuterFrameCheckbox.onClick = refreshPreview;
    }

    /**
     * ダイアログを表示してオプションを受け取る
     * @param {{horizontalLines: Object[], verticalLines: Object[]}} classified - 分類済みの線
     * @param {Object} gridBounds - 外周座標
     * @param {Object[]} originalLineStates - 元の線の状態
     * @param {{label: string, pointsPerUnit: number}} strokeUnitInfo - 線の単位
     * @param {boolean} hasExpandedRectangles - 外枠の長方形を分解したかどうか
     * @returns {Object|null} ダイアログの設定。キャンセルなら null
     */
    function showOptionDialog(classified, gridBounds, originalLineStates, strokeUnitInfo, hasExpandedRectangles) {
        var dialogControls = buildOptionDialog(strokeUnitInfo, hasExpandedRectangles);
        bindOptionDialogEvents(dialogControls, function () {
            updatePreviewFromDialogState(dialogControls, classified, gridBounds, originalLineStates);
        });

        prepareDialogWindow(dialogControls.optionDialog, SCRIPT_NAME);
        var dialogResult = dialogControls.optionDialog.show();
        restoreLineStates(originalLineStates);
        if (dialogResult !== 1) return null;
        return readOptionDialogState(dialogControls);
    }

    /* 定数の代入がすべて済んでから実行する / run only after every constant above has been assigned */
    main();

})();
