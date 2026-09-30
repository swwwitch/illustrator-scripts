#target illustrator
#targetengine "DirectPrefs"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

環境設定の「角度の制限」と「キー増加」の変更、およびガイド・グリッドの表示やロックの切り替えをパレットから行います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DirectPrefsPalette.md

### Overview

A palette for changing the constrain angle and the keyboard increment, and for toggling the display and lock state of guides and the grid.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DirectPrefsPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DirectPrefsPalette";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DirectPrefsPalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DirectPrefsPalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 角度の制限のプリセット（アイソメトリック作図の3軸＋0°）/ Constrain-angle presets (the three isometric axes plus 0°) */
    var CONSTRAIN_PRESETS = [0, 150, 90, 30];

    /* キー増加のプリセット（定規単位コードごと。値は各単位そのままの数値）
       / Keyboard-increment presets per ruler unit code (values are in that unit) */
    var KEY_INCREMENT_PRESETS = {
        "1": [1, 5, 10],   /* mm */
        "2": [1, 6, 12],   /* pt */
        "6": [0.1, 1, 8]   /* px */
    };

    /* 上表にない単位（in / cm / Q/H など）で使うプリセット / Presets used for units missing from the table (in, cm, Q/H, ...) */
    var KEY_INCREMENT_PRESETS_DEFAULT = [1, 5, 10];

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

    var GROUP_SPACING = 8; /* setupGroup() の既定の間隔 / default spacing of setupGroup() */
    var INPUT_CHARS = 6;
    var PRESET_BUTTON_WIDTH = 48;
    var UNIT_LABEL_WIDTH = 32;

    /**
     * グループの共通設定をまとめて適用する（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" / "column"（省略時は "column"）
     * @param {number} [spacing] - 子どうしの間隔（省略時は GROUP_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : GROUP_SPACING;
    }

    /**
     * 行に「∧∨・入力欄」を隙間0で突き合わせて追加する。↑↓キーも∧∨と同じ処理で増減する
     * （値を変えるだけで、適用は従来どおり各パネルの変更ボタンで行う）
     * @param {Group} parentRow - 追加先の行
     * @param {Object} stepOptions - addStepper() に渡す step / min / max / integer / unit / onStep
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSteppedInput(parentRow, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, "");
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
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

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "DirectPrefs", en: "DirectPrefs" }
        },
        panel: {
            constrain: { ja: "角度の制限", en: "Constrain Angle" },
            keyIncrement: { ja: "キー増加", en: "Keyboard Increment" },
            guide: { ja: "ガイド", en: "Guides" },
            grid: { ja: "グリッド", en: "Grid" }
        },
        fieldLabel: {
            constrain: { ja: "角度の制限", en: "Constrain angle" },
            keyIncrement: { ja: "キー増加", en: "Keyboard increment" }
        },
        button: {
            applyConstrain: { ja: "「角度の制限」の値を変更", en: "Change constrain angle" },
            applyKeyIncrement: { ja: "変更", en: "Change" },
            resetConstrain: { ja: "リセット", en: "Reset" },
            toggleVisibility: { ja: "表示・非表示", en: "Show/Hide" },
            toggleLock: { ja: "ロック・ロック解除", en: "Lock/Unlock" },
            snapToGrid: { ja: "グリッドにスナップ", en: "Snap to Grid" }
        },
        tooltip: {
            constrainInput: {
                ja: "ビューの回転角度が候補として入ります（ドキュメントがないときは現在の値）。［「角度の制限」の値を変更］で適用します。",
                en: "Prefilled with the view rotation (the current value when no document is open). Click \"Change constrain angle\" to apply it."
            },
            constrainPreset: { ja: "この角度を「角度の制限」にすぐ適用します。", en: "Applies this angle to the Constrain Angle preference right away." },
            resetConstrain: { ja: "角度の制限を0°に戻します。", en: "Resets the constrain angle to 0°." },
            keyIncrementPreset: { ja: "この値をキー増加にすぐ適用します。", en: "Applies this value to the keyboard increment right away." },
            keyIncrementInput: { ja: "定規の単位で入力します。［変更］で適用します。", en: "Entered in the ruler unit. Click Change to apply it." },
            snapToGrid: { ja: "「グリッドにスナップ」のオン・オフを切り替えます。", en: "Turns Snap to Grid on or off." },
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
        status: {
            applied: { ja: "制限角度に適用しました。", en: "Applied to the constrain angle." },
            appliedKeyIncrement: { ja: "キー増加に適用しました。", en: "Applied to the keyboard increment." },
            resetConstrain: { ja: "制限角度を0°にリセットしました。", en: "Reset the constrain angle to 0°." },
            toggledGuideVisibility: { ja: "ガイドの表示を切り替えました。", en: "Toggled guide visibility." },
            toggledGuideLock: { ja: "ガイドのロックを切り替えました。", en: "Toggled the guide lock." },
            toggledGridVisibility: { ja: "グリッドの表示を切り替えました。", en: "Toggled grid visibility." },
            toggledSnapToGrid: { ja: "グリッドにスナップを切り替えました。", en: "Toggled snap to grid." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            invalidAngle: { ja: "角度には数値を入力してください。", en: "Please enter a numeric angle." },
            invalidValue: { ja: "0以上の数値を入力してください。", en: "Please enter a number of 0 or greater." },
            error: { ja: "エラーが発生しました：", en: "An error occurred:" }
        }
    };

    /**
     * ワーカーの結果文字列を表示用のステータス文にする（既知のコードは専用文、未知のものは汎用エラー）
     * @param {string} result - ワーカーから返った結果文字列
     * @returns {string} ステータス文
     */
    function statusFromResult(result) {
        if (result.indexOf("NODOC") !== -1) {
            return getLabel("alert.noDocument");
        }
        return getLabel("alert.error") + " " + result;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位の定義を1か所に集約（コード／ラベル／pt換算係数／表示桁数）
       decimals：1pt 未満に潰れないように、大きい単位ほど桁数を増やす（in で 1mm ≒ 0.039）
       / Single source of unit definitions (code, label, pt factor, display decimals)
       decimals: larger units need more digits so small values do not collapse to 0 (1mm is 0.039in) */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72,               decimals: 3 },  /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4,        decimals: 1 },  /* 1 */
        { label: "pt",    pointsPerUnit: 1,                decimals: 1 },  /* 2 */
        { label: "pica",  pointsPerUnit: 12,               decimals: 2 },  /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54,        decimals: 2 },  /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25, decimals: 1 },  /* 5 */
        { label: "px",    pointsPerUnit: 1,                decimals: 1 },  /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12,          decimals: 4 },  /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000, decimals: 4 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36,          decimals: 4 },  /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12,          decimals: 4 }   /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 単位コードから単位定義を取得する（未対応コードは pt 相当）
     * @param {number} unitCode - 環境設定の単位コード
     * @returns {{label: string, pointsPerUnit: number, decimals: number}} 単位定義
     */
    function getUnitByCode(unitCode) {
        return UNITS[unitCode] || UNITS[2];
    }

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

    /**
     * 単位コードに対応するキー増加のプリセットを取得する
     * @param {number} unitCode - 定規の単位コード
     * @returns {number[]} プリセットの値（その単位での数値）
     */
    function getKeyIncrementPresets(unitCode) {
        var presets = KEY_INCREMENT_PRESETS[String(unitCode)];
        return presets ? presets : KEY_INCREMENT_PRESETS_DEFAULT;
    }

    // =========================================
    // 角度の計算 / Angle helpers
    // =========================================

    /**
     * 角度を -180〜180 に正規化する
     * @param {number} angle - 角度（度）
     * @returns {number} 正規化した角度
     */
    function normalizeAngle(angle) {
        var normalized = angle % 360;
        if (normalized > 180) { normalized -= 360; }
        if (normalized < -180) { normalized += 360; }
        return normalized;
    }

    /**
     * 表示用に小数2桁へ丸める（atan2 由来の 30.00000001 のような桁あふれを抑える）
     * @param {number} angle - 角度（度）
     * @returns {number} 丸めた角度
     */
    function roundAngle(angle) {
        return Math.round(angle * 100) / 100;
    }

    // =========================================
    // メインエンジンへの委譲 / Delegation to the main engine
    // =========================================

    /**
     * メインエンジンでコードを実行する（常駐パレットの app は DOM 接続を失うため）。
     * 本文は encodeURIComponent + eval で送り、バックスラッシュ・多バイト文字を無傷で渡す
     * @param {string} code - 実行するコード
     * @param {Function} onResult - 結果文字列を受け取るコールバック（失敗時は "ERR:" で始まる）
     * @returns {void}
     */
    function runInMainEngine(code, onResult) {
        var bridgeTalk = new BridgeTalk();
        bridgeTalk.target = "illustrator";
        bridgeTalk.body = 'eval(decodeURIComponent("' + encodeURIComponent(code) + '"));';
        bridgeTalk.onResult = function (response) {
            onResult(String(response.body));
        };
        bridgeTalk.onError = function (response) {
            onResult("ERR:" + String(response.body));
        };
        bridgeTalk.send();
    }

    /**
     * ワーカーの本文を IIFE で包む
     * @param {string} workerCode - ワーカーの本文
     * @returns {string} IIFE で包んだコード
     */
    function wrapWorkerBody(workerCode) {
        return "(function(){" + workerCode + "})()";
    }

    /* ワーカー断片：「角度の制限」を読む。実際の拘束方向は constrain/sin・constrain/cos が持っているため、
       constrain/angle ではなくこの2つから角度を復元する（angle は書いても拘束に反映されない）
       / Worker fragment: read the constrain angle. The real constraint direction lives in constrain/sin and
       constrain/cos, so recover the angle from those instead of constrain/angle (writing `angle` alone has no effect) */
    var WORKER_READ_CONSTRAIN =
        "var constrainAngle=Math.atan2(app.preferences.getRealPreference('constrain/sin')," +
        "app.preferences.getRealPreference('constrain/cos'))*180/Math.PI;";

    /**
     * 「角度の制限」を書き込むワーカー断片を作る。angle は環境設定ダイアログの表示用で度、
     * sin・cos は実際の拘束方向でラジアン由来。angle だけでは拘束に効かないので3つとも書く
     * @param {number} angle - 書き込む角度（度）
     * @returns {string} ワーカー断片
     */
    function buildConstrainWriteCode(angle) {
        return "var radians=(" + angle + ")*Math.PI/180;" +
            "app.preferences.setRealPreference('constrain/angle'," + angle + ");" +
            "app.preferences.setRealPreference('constrain/sin',Math.sin(radians));" +
            "app.preferences.setRealPreference('constrain/cos',Math.cos(radians));";
    }

    /**
     * ビュー回転角度・制限角度・定規単位・キー増加(pt)を1回の委譲でまとめて取得する
     * （"OK:回転,制限,単位コード,キー増加"。回転はドキュメントが開いていなければ空）。
     * 環境設定はドキュメントがなくても読めるため、回転だけを条件付きにしている
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function fetchState(onResult) {
        runInMainEngine(wrapWorkerBody(
            "var viewRotation=(app.documents.length>0)?app.activeDocument.activeView.rotateAngle:'';" +
            WORKER_READ_CONSTRAIN +
            "var rulerUnitCode=app.preferences.getIntegerPreference('rulerType');" +
            "var keyIncrementPt=app.preferences.getRealPreference('cursorKeyLength');" +
            "return 'OK:'+viewRotation+','+constrainAngle+','+rulerUnitCode+','+keyIncrementPt;"
        ), onResult);
    }

    /**
     * 制限角度をメインエンジンで環境設定に適用する
     * @param {number} angle - 制限角度（度）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function applyConstrainAngle(angle, onResult) {
        runInMainEngine(wrapWorkerBody(
            buildConstrainWriteCode(angle) +
            "return 'OK';"
        ), onResult);
    }

    /**
     * キー増加をメインエンジンで環境設定に適用する（値は pt）
     * @param {number} lengthPt - キー増加（pt）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function applyKeyIncrement(lengthPt, onResult) {
        runInMainEngine(wrapWorkerBody(
            "app.preferences.setRealPreference('cursorKeyLength'," + lengthPt + ");" +
            "return 'OK';"
        ), onResult);
    }

    /**
     * メニューコマンドをメインエンジンで実行する（ガイド・グリッドの状態は取得できないため、メニューのトグルを呼ぶ）
     * @param {string} menuCommand - メニューコマンド名
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function runMenuCommand(menuCommand, onResult) {
        runInMainEngine(wrapWorkerBody(
            "if(app.documents.length===0){return 'ERR:NODOC';}" +
            "app.executeMenuCommand('" + menuCommand + "');" +
            "return 'OK';"
        ), onResult);
    }

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

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * パレットと各コントロールを作る
     * @returns {Object} パレット（palette）と各コントロールの参照
     */
    function buildPalette() {
        var prefsPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(prefsPalette);

        /* 角度の制限を変更するパネル / Panel for changing the constrain angle */
        var constrainPanel = prefsPalette.add("panel", undefined, getLabel("panel.constrain"));
        setupPanel(constrainPanel, 6);

        /* 角度の制限（編集可。ビューの回転角度が候補値として入るが、反映は適用ボタンを押したときだけ）
           / Constrain angle (editable; seeded with the view rotation as a suggestion, but committed only on the Apply button) */
        var constrainInputGroup = constrainPanel.add("group");
        setupGroup(constrainInputGroup, "row");
        constrainInputGroup.add("statictext", undefined, labelText("fieldLabel.constrain"));
        var constrainInput = addSteppedInput(constrainInputGroup, {});
        constrainInput.characters = INPUT_CHARS;
        constrainInput.helpTip = getLabel("tooltip.constrainInput");
        constrainInputGroup.add("statictext", undefined, "°");

        /* プリセットボタン行（押すとその角度を即座に適用）/ Preset button row (clicking applies that angle immediately) */
        var constrainPresetGroup = constrainPanel.add("group");
        setupGroup(constrainPresetGroup, "row", 4);
        var constrainPresetButtons = [];
        for (var i = 0; i < CONSTRAIN_PRESETS.length; i++) {
            var constrainPresetButton = constrainPresetGroup.add("button", undefined, CONSTRAIN_PRESETS[i] + "°");
            constrainPresetButton.preferredSize.width = PRESET_BUTTON_WIDTH;
            constrainPresetButton.helpTip = getLabel("tooltip.constrainPreset");
            constrainPresetButtons.push({ angle: CONSTRAIN_PRESETS[i], button: constrainPresetButton });
        }

        /* ボタン行：適用とリセット / Button row: Apply and Reset */
        var constrainButtonGroup = constrainPanel.add("group");
        setupGroup(constrainButtonGroup, "row");
        constrainButtonGroup.alignment = "right";

        /* 角度の制限を0°に戻す / Reset the constrain angle to 0° */
        var resetConstrainButton = constrainButtonGroup.add("button", undefined, getLabel("button.resetConstrain"));
        resetConstrainButton.helpTip = getLabel("tooltip.resetConstrain");

        /* 「角度の制限」の値を変更ボタン / Change-constrain-angle button */
        var applyConstrainButton = constrainButtonGroup.add("button", undefined, getLabel("button.applyConstrain"));

        /* キー増加を変更するパネル / Panel for changing the keyboard increment */
        var keyIncrementPanel = prefsPalette.add("panel", undefined, getLabel("panel.keyIncrement"));
        setupPanel(keyIncrementPanel, 6);

        /* キー増加（定規単位で表示・入力。単位ラベルは rulerType に追従）
           / Keyboard increment (shown and entered in the ruler unit; the unit label follows rulerType) */
        var keyIncrementInputGroup = keyIncrementPanel.add("group");
        setupGroup(keyIncrementInputGroup, "row");
        keyIncrementInputGroup.add("statictext", undefined, labelText("fieldLabel.keyIncrement"));
        /* 0以上（適用時の検証と同じ下限） / 0 or greater, the same floor as the apply check */
        var keyIncrementInput = addSteppedInput(keyIncrementInputGroup, { min: 0 });
        keyIncrementInput.characters = INPUT_CHARS;
        keyIncrementInput.helpTip = getLabel("tooltip.keyIncrementInput");
        var keyIncrementUnitLabel = keyIncrementInputGroup.add("statictext", undefined, "pt");
        keyIncrementUnitLabel.preferredSize.width = UNIT_LABEL_WIDTH;

        /* プリセットボタン行（押すとその値を即座に適用。単位が変わっても作り直さず、表示値だけ差し替える）
           / Preset button row (clicking applies that value immediately; the buttons persist and only their values follow the ruler unit) */
        var keyIncrementPresetGroup = keyIncrementPanel.add("group");
        setupGroup(keyIncrementPresetGroup, "row", 4);
        var keyIncrementPresetButtons = [];
        for (var j = 0; j < KEY_INCREMENT_PRESETS_DEFAULT.length; j++) {
            var keyIncrementPresetButton = keyIncrementPresetGroup.add("button", undefined, "");
            keyIncrementPresetButton.preferredSize.width = PRESET_BUTTON_WIDTH;
            keyIncrementPresetButton.helpTip = getLabel("tooltip.keyIncrementPreset");
            keyIncrementPresetButtons.push(keyIncrementPresetButton);
        }

        /* 「キー増加」の値を変更ボタン / Change-keyboard-increment button */
        var applyKeyIncrementButton = keyIncrementPanel.add("button", undefined, getLabel("button.applyKeyIncrement"));
        applyKeyIncrementButton.alignment = "right";

        /* ガイドのパネル / Guides panel */
        var guidePanel = prefsPalette.add("panel", undefined, getLabel("panel.guide"));
        setupPanel(guidePanel, 6);

        var guideButtonGroup = guidePanel.add("group");
        setupGroup(guideButtonGroup, "row", 4);

        /* ガイドの表示・非表示を切り替えるボタン / Button that toggles guide visibility */
        var toggleGuideVisibilityButton = guideButtonGroup.add("button", undefined, getLabel("button.toggleVisibility"));

        /* ガイドのロック・ロック解除を切り替えるボタン / Button that toggles the guide lock */
        var toggleGuideLockButton = guideButtonGroup.add("button", undefined, getLabel("button.toggleLock"));

        /* グリッドのパネル / Grid panel */
        var gridPanel = prefsPalette.add("panel", undefined, getLabel("panel.grid"));
        setupPanel(gridPanel, 6);

        var gridButtonGroup = gridPanel.add("group");
        setupGroup(gridButtonGroup, "row", 4);

        /* グリッドの表示・非表示を切り替えるボタン / Button that toggles grid visibility */
        var toggleGridVisibilityButton = gridButtonGroup.add("button", undefined, getLabel("button.toggleVisibility"));

        /* グリッドにスナップを切り替えるボタン / Button that toggles snap to grid */
        var snapToGridButton = gridButtonGroup.add("button", undefined, getLabel("button.snapToGrid"));
        snapToGridButton.helpTip = getLabel("tooltip.snapToGrid");

        /* ステータス表示 / Status line */
        var statusText = prefsPalette.add("statictext", undefined, "");
        statusText.alignment = "fill";

        return {
            palette: prefsPalette,
            constrainInput: constrainInput,
            constrainPresetButtons: constrainPresetButtons,
            resetConstrainButton: resetConstrainButton,
            applyConstrainButton: applyConstrainButton,
            keyIncrementInput: keyIncrementInput,
            keyIncrementUnitLabel: keyIncrementUnitLabel,
            keyIncrementPresetButtons: keyIncrementPresetButtons,
            applyKeyIncrementButton: applyKeyIncrementButton,
            toggleGuideVisibilityButton: toggleGuideVisibilityButton,
            toggleGuideLockButton: toggleGuideLockButton,
            toggleGridVisibilityButton: toggleGridVisibilityButton,
            snapToGridButton: snapToGridButton,
            statusText: statusText
        };
    }

    /**
     * パレットの表示更新と各コントロールの操作を結び付ける
     * @param {Object} paletteUI - buildPalette() が返したパレットとコントロールの参照
     * @returns {void}
     */
    function bindPaletteHandlers(paletteUI) {
        /* 現在の制限角度と定規単位（リセットボタンのディム判定・単位換算に使用）
           / Current constrain angle and ruler unit (used to dim the Reset button and to convert units) */
        var currentConstrain = 0;
        var currentUnit = getUnitByCode(2);

        /* キー増加のプリセットボタンが押された時点で読む値 / Values the keyboard-increment preset buttons read at click time */
        var keyIncrementPresetValues = KEY_INCREMENT_PRESETS_DEFAULT;

        /**
         * ワーカーの結果を受けるコールバックを作る。失敗ならステータスにエラーを出し、成功なら onSuccess を呼ぶ
         * @param {Function} onSuccess - 成功時の処理
         * @returns {Function} runInMainEngine に渡すコールバック
         */
        function handleWorkerResult(onSuccess) {
            return function (result) {
                if (result.indexOf("OK") === 0) {
                    onSuccess();
                } else {
                    paletteUI.statusText.text = statusFromResult(result);
                }
            };
        }

        /**
         * 適用済みの制限角度を状態に記録し、0°ならリセットボタンをディムする
         * @param {number} angle - 制限角度（度）
         * @returns {void}
         */
        function setConstrain(angle) {
            currentConstrain = angle;
            paletteUI.resetConstrainButton.enabled = (currentConstrain !== 0);
        }

        /**
         * 制限角度を適用して表示・状態を更新する
         * @param {number} angle - 制限角度（度）
         * @returns {void}
         */
        function commitConstrain(angle) {
            applyConstrainAngle(angle, handleWorkerResult(function () {
                setConstrain(angle);
                paletteUI.constrainInput.text = roundAngle(angle);
                paletteUI.statusText.text = getLabel("status.applied");
            }));
        }

        /**
         * キー増加(pt)を現在の定規単位の表示文字列に変換する
         * @param {number} lengthPt - キー増加（pt）
         * @returns {string} 表示用の文字列
         */
        function formatKeyIncrement(lengthPt) {
            return (lengthPt / currentUnit.pointsPerUnit).toFixed(currentUnit.decimals);
        }

        /**
         * キー増加を適用して表示を更新する（値は現在の定規単位）
         * @param {number} unitValue - 定規単位での値
         * @returns {void}
         */
        function commitKeyIncrement(unitValue) {
            var lengthPt = unitValue * currentUnit.pointsPerUnit;
            applyKeyIncrement(lengthPt, handleWorkerResult(function () {
                paletteUI.keyIncrementInput.text = formatKeyIncrement(lengthPt);
                paletteUI.statusText.text = getLabel("status.appliedKeyIncrement");
            }));
        }

        /**
         * 現在の定規単位に合わせてプリセットボタンの値とラベルを差し替える
         * @returns {void}
         */
        function updateKeyIncrementPresets() {
            keyIncrementPresetValues = getKeyIncrementPresets(currentUnit.code);
            for (var i = 0; i < paletteUI.keyIncrementPresetButtons.length; i++) {
                var hasValue = (i < keyIncrementPresetValues.length);
                paletteUI.keyIncrementPresetButtons[i].visible = hasValue;
                if (hasValue) {
                    paletteUI.keyIncrementPresetButtons[i].text = keyIncrementPresetValues[i] + currentUnit.label;
                }
            }
        }

        /**
         * ビュー回転角度・制限角度・定規単位・キー増加を取得し、各表示へ反映する
         * （制限角度の入力欄にはビューの回転角度を候補値として入れる。適用はボタン押下時のみ）
         * @returns {void}
         */
        function refreshState() {
            fetchState(function (result) {
                if (result.indexOf("OK:") !== 0) {
                    paletteUI.statusText.text = statusFromResult(result);
                    return;
                }
                var stateValues = result.substring(3).split(",");
                setConstrain(parseFloat(stateValues[1]));
                currentUnit = getUnitByCode(parseInt(stateValues[2], 10));
                paletteUI.keyIncrementUnitLabel.text = currentUnit.label;
                paletteUI.keyIncrementInput.text = formatKeyIncrement(parseFloat(stateValues[3]));
                updateKeyIncrementPresets();
                if (stateValues[0] === "") {
                    /* ドキュメントがなければ回転を取得できないので、現在の制限角度をそのまま表示
                       / Without a document there is no rotation to read, so show the current constrain angle as-is */
                    paletteUI.constrainInput.text = roundAngle(currentConstrain);
                    paletteUI.statusText.text = getLabel("alert.noDocument");
                } else {
                    paletteUI.constrainInput.text = roundAngle(normalizeAngle(parseFloat(stateValues[0])));
                    paletteUI.statusText.text = "";
                }
            });
        }

        /**
         * 角度の制限のプリセットボタンに適用処理を付ける（ES3にはブロックスコープがないため、クロージャを関数で切り出す）
         * @param {Object} constrainPreset - プリセットの角度（angle）とボタン（button）の対
         * @returns {void}
         */
        function bindConstrainPreset(constrainPreset) {
            constrainPreset.button.onClick = function () {
                commitConstrain(constrainPreset.angle);
            };
        }

        /**
         * キー増加のプリセットボタンに適用処理を付ける（押した時点の keyIncrementPresetValues を参照する）
         * @param {Button} presetButton - 対象のボタン
         * @param {number} presetIndex - プリセットの番号
         * @returns {void}
         */
        function bindKeyIncrementPreset(presetButton, presetIndex) {
            presetButton.onClick = function () {
                commitKeyIncrement(keyIncrementPresetValues[presetIndex]);
            };
        }

        /**
         * メニューコマンドをトグルボタンに割り当てる
         * @param {Button} toggleButton - 対象のボタン
         * @param {string} menuCommand - メニューコマンド名
         * @param {string} statusPath - 成功時に表示するステータスのラベルパス
         * @returns {void}
         */
        function bindMenuCommand(toggleButton, menuCommand, statusPath) {
            toggleButton.onClick = function () {
                runMenuCommand(menuCommand, handleWorkerResult(function () {
                    paletteUI.statusText.text = getLabel(statusPath);
                }));
            };
        }

        for (var i = 0; i < paletteUI.constrainPresetButtons.length; i++) {
            bindConstrainPreset(paletteUI.constrainPresetButtons[i]);
        }
        for (var j = 0; j < paletteUI.keyIncrementPresetButtons.length; j++) {
            bindKeyIncrementPreset(paletteUI.keyIncrementPresetButtons[j], j);
        }

        /* リセット：制限角度を0°に戻して入力欄と状態を更新 / Reset: set the constrain angle to 0° and refresh the field and state */
        paletteUI.resetConstrainButton.onClick = function () {
            applyConstrainAngle(0, handleWorkerResult(function () {
                setConstrain(0);
                paletteUI.constrainInput.text = "0";
                paletteUI.statusText.text = getLabel("status.resetConstrain");
            }));
        };

        /* 適用ボタン：入力値を検証してメインエンジンで適用 / Apply button: validate the input and apply in the main engine */
        paletteUI.applyConstrainButton.onClick = function () {
            var constrainAngle = parseFloat(paletteUI.constrainInput.text);
            if (isNaN(constrainAngle)) {
                paletteUI.statusText.text = getLabel("alert.invalidAngle");
                return;
            }
            commitConstrain(constrainAngle);
        };

        /* キー増加の適用ボタン：入力値（定規単位）を pt に換算して適用 / Keyboard-increment apply button: convert the entered value (ruler unit) to pt and apply */
        paletteUI.applyKeyIncrementButton.onClick = function () {
            var unitValue = parseFloat(paletteUI.keyIncrementInput.text);
            if (isNaN(unitValue) || unitValue < 0) {
                paletteUI.statusText.text = getLabel("alert.invalidValue");
                return;
            }
            commitKeyIncrement(unitValue);
        };

        bindMenuCommand(paletteUI.toggleGuideVisibilityButton, "showguide", "status.toggledGuideVisibility");
        bindMenuCommand(paletteUI.toggleGuideLockButton, "lockguide", "status.toggledGuideLock");
        bindMenuCommand(paletteUI.toggleGridVisibilityButton, "showgrid", "status.toggledGridVisibility");
        bindMenuCommand(paletteUI.snapToGridButton, "snapgrid", "status.toggledSnapToGrid");

        /* 初期表示時、およびパレットがアクティブになるたびに最新の状態を取得
           / Fetch the latest state on first show and whenever the palette becomes active */
        paletteUI.palette.onShow = refreshState;
        paletteUI.palette.onActivate = refreshState;
    }

    /* 常駐エンジンにパレット参照を保持するキー（CloseAllPalettes.jsx から閉じるため）
       Key holding the palette reference in the persistent engine (so CloseAllPalettes.jsx can close it) */
    var PALETTE_GLOBAL_KEY = "__directPrefsPalette";

    var paletteUI = buildPalette();
    bindPaletteHandlers(paletteUI);
    /* 閉じたら常駐エンジンの参照をクリア / Clear the persistent-engine reference on close */
    paletteUI.palette.onClose = function () { $.global[PALETTE_GLOBAL_KEY] = null; };
    /* パレットは Esc で閉じないので、閉じる処理を割り当てる（入力中も効かせる）
       Palettes do not close on Esc by themselves; map it to close, also while typing */
    addKeyShortcuts(paletteUI.palette, {
        "Escape": { target: function () { paletteUI.palette.close(); }, inFields: true }
    });
    paletteUI.palette.show();
    $.global[PALETTE_GLOBAL_KEY] = paletteUI.palette;

})();
