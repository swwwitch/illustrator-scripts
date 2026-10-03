#target illustrator
#targetengine "AdjustVerticalGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した2つのオブジェクトの上下の間隔を、指定した値にそろえる常駐パレットです。
ライブプレビューに対応し、設定を変えるたびに結果を確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAdjustVerticalGapPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8201294835f9

### Overview

A persistent palette that sets the vertical gap between two selected objects to a value you specify.
A live preview shows the result as you change the settings.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAdjustVerticalGapPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAdjustVerticalGapPalette";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAdjustVerticalGapPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAdjustVerticalGapPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8201294835f9"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var DEFAULT_GAP_VALUE = "3"; /* 間隔の初期値（定規の単位）/ Default gap (ruler unit) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var FIELD_CHARS = 5;                   /* 数値欄の幅（文字数）/ width of the number fields */
    var GROUP_SPACING = 8;                 /* setupGroup() の要素間隔 / spacing used by setupGroup() */

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

    /**
     * グループの共通設定（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（既定）
     * @param {number} [spacing] - 要素間隔（省略時は GROUP_SPACING）
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

    /* 入力された単位を欄の単位へ換算するための、1単位あたりのポイント数（値は UNITS 表と同じ。キーは小文字）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Points per unit for converting typed units into the field's unit (same values as the UNITS table; lowercase keys) */
    var STEPPER_POINTS_PER_UNIT = {
        "in": 72, "inch": 72, "mm": 72 / 25.4, "cm": 72 / 2.54, "m": 72 / 25.4 * 1000,
        "pt": 1, "px": 1, "p": 12, "pc": 12, "pica": 12,
        "q": 72 / 25.4 * 0.25, "h": 72 / 25.4 * 0.25, "ft": 72 * 12, "yd": 72 * 36
    };

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
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
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
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
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
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
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) return;
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
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
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldPointsPerUnit = STEPPER_POINTS_PER_UNIT[fieldUnitKey];
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var pointsPerUnit = STEPPER_POINTS_PER_UNIT[unitKey];
            if (pointsPerUnit === undefined || fieldPointsPerUnit === undefined) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = value * pointsPerUnit;
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldPointsPerUnit;
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
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
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
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
    // 常駐パレットの参照 / Resident palette reference
    // =========================================

    /* パレットの参照を常駐エンジンに保持 / Keep the palette reference alive in the resident engine */
    var paletteWindow = null;

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

    var LABELS = {
        dialog: {
            title: { ja: "上下間隔を調整", en: "Adjust Vertical Gap" }
        },
        panel: {
            anchor: { ja: "キーオブジェクト", en: "Key object" },
            gap: { ja: "上下間隔", en: "Vertical Gap" },
            align: { ja: "横方向の整列", en: "Horizontal Alignment" },
            justify: { ja: "テキストの行揃え", en: "Text alignment" }
        },
        radio: {
            anchorTop: { ja: "上", en: "Top" },
            anchorBottom: { ja: "下", en: "Bottom" },
            anchorAuto: { ja: "自動判定", en: "Auto" },
            alignNone: { ja: "なし", en: "None" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            justifyNone: { ja: "変更しない", en: "Keep current" },
            justifyLink: { ja: "整列に連動", en: "Match alignment" },
            justifyFull: { ja: "均等配置（最終行左）", en: "Justify (last line left-aligned)" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Use preview bounds" }
        },
        fieldLabel: {
            alignAdjust: { ja: "横調整", en: "Offset" }
        },
        button: {
            record: { ja: "記録", en: "Record" },
            edit: { ja: "編集", en: "Edit" },
            apply: { ja: "適用", en: "Apply" }
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
            anchorTop: {
                ja: "基準にする（動かさない）キーオブジェクト。上を選択（ショートカット: T）。",
                en: "Key object to keep in place. Select top (shortcut: T)."
            },
            anchorBottom: {
                ja: "基準にする（動かさない）キーオブジェクト。下を選択（ショートカット: B）。",
                en: "Key object to keep in place. Select bottom (shortcut: B)."
            },
            anchorAuto: {
                ja: "Illustrator で設定したキーオブジェクトを判定して、上下どちらを基準にするか決めます（ショートカット: K）。整列コマンドで一時的に動かして判定し、元の位置へ戻します。判定できないときは「上」になります。",
                en: "Detect the key object set in Illustrator and use it as the anchor (shortcut: K). Align commands probe temporarily and every item is moved back. Falls back to Top when it cannot be detected."
            },
            gap: {
                ja: "上下に並ぶ2つのオブジェクト間の距離です。負の値で重なります。↑↓キーで増減（Shiftで±10／Optionで±0.1）。単位は定規に従います。",
                en: "Vertical distance between the two objects (negative values overlap them). Arrow keys change it (Shift: ±10 / Option: ±0.1). Unit follows the ruler."
            },
            previewBounds: {
                ja: "線幅や効果を含む見た目の境界で間隔を計算します。オフにすると線幅や効果を含まないパス境界で計算します。",
                en: "Calculate the gap using visual bounds, including strokes and effects. Turn off to use geometric path bounds."
            },
            align: {
                ja: "移動するオブジェクトを、キーオブジェクトの左・中央・右にそろえます（ショートカット: 整列しない=N／左=L／中央=C／右=R）。",
                en: "Align the moving object to the left, center, or right of the key object (shortcuts: none=N / left=L / center=C / right=R)."
            },
            alignAdjust: {
                ja: "移動するオブジェクトを左右へ追加でずらす量です。正の値で右へ、負の値で左へ移動します。整列「なし」でも有効です。↑↓キーで増減（Shiftで±10／Optionで±0.1）。単位は定規に従います。",
                en: "Additional horizontal offset for the moving object. Positive moves right, negative moves left. Also works when alignment is None. Arrow keys change it (Shift: ±10 / Option: ±0.1). Unit follows the ruler."
            },
            justify: {
                ja: "テキストの段落の行揃え。「整列に連動」は左右の整列（左／中央／右）に合わせます。テキスト以外には影響しません（ショートカット: 整列に連動=S／均等配置=J）。",
                en: "Paragraph alignment of text. \"Match alignment\" follows the horizontal align. Non-text objects are unaffected (shortcuts: match=S / justify=J)."
            },
            justifyFull: {
                ja: "段落を均等配置します（最終行は左揃え）。",
                en: "Justify paragraphs (last line left-aligned)."
            },
            record: {
                ja: "現在の設定（間隔・キー・整列・行揃え）を記録し、パネルをロックします。記録後、複数のグループを選択して［適用］で一括適用できます。",
                en: "Record the current settings (gap, key object, align, justify) and lock the panels. Then select multiple groups and Apply to batch-apply."
            },
            edit: {
                ja: "ロックを解除して設定を編集できるようにします。",
                en: "Unlock the panels to edit the settings again."
            },
            apply: {
                ja: "記録した設定を、選択中のすべての対象（各グループの2点／2点選択）に一括適用します（ショートカット: A）。未記録なら現在の設定を使います。",
                en: "Batch-apply the recorded settings to every target in the selection (each group of two, or two selected objects) (shortcut: A). Falls back to current settings if nothing is recorded."
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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 委譲される処理 / Delegated worker functions
    //   メインエンジンへ送って実行する。toString() で連結するため、
    //   この節の関数には JSDoc を付けず、// コメントを使わず /* */ のみ・必ずセミコロンで終える。
    //   Sent to the main engine: no JSDoc, only /* */ comments and explicit semicolons here.
    // =========================================

    /* クリップグループのクリップパスを返す（無ければ null）/ Return the clipping path of a clip group, or null */
    function getClippingPath(groupItem) {
        var groupPaths = groupItem.pathItems;
        for (var i = 0; i < groupPaths.length; i++) {
            if (groupPaths[i].clipping === true) {
                return groupPaths[i];
            }
        }
        /* 複合パスでクリップしている場合 / When the clip path is a compound path */
        var compoundPaths = groupItem.compoundPathItems;
        for (var j = 0; j < compoundPaths.length; j++) {
            if (compoundPaths[j].pathItems.length > 0 && compoundPaths[j].pathItems[0].clipping === true) {
                return compoundPaths[j];
            }
        }
        return null;
    }

    /* 計算に使う境界を取得（クリップグループはクリップパス基準）/ Bounds for calculation (clip group uses its clipping path) */
    function getItemBounds(targetItem, usePreviewBounds) {
        var boundsTarget = targetItem;
        if (targetItem.constructor.name === "GroupItem" && targetItem.clipped === true) {
            var clipPath = getClippingPath(targetItem);
            if (clipPath !== null) {
                boundsTarget = clipPath;
            }
        }
        return usePreviewBounds ? boundsTarget.visibleBounds : boundsTarget.geometricBounds;
    }

    /* UI の行揃え選択を実際のキーに解決（「整列に連動」は整列値をそのまま使う。「なし」なら "none"）/
       Resolve the choice ("link" reuses the align value; align "none" stays "none") */
    function resolveJustifyKey(justify, align) {
        if (justify === "link") {
            return align;
        }
        return justify;
    }

    /* 行揃えキーを Justification 列挙値に変換 / Map key to Justification enum */
    function resolveJustification(justifyKey) {
        if (justifyKey === "left") {
            return Justification.LEFT;
        }
        if (justifyKey === "center") {
            return Justification.CENTER;
        }
        if (justifyKey === "right") {
            return Justification.RIGHT;
        }
        if (justifyKey === "full") {
            return Justification.FULLJUSTIFYLASTLINELEFT;
        }
        return null;
    }

    /* TextFrame なら段落の行揃えを設定。実際に変更したら true を返す（no-op 検出用）/
       Set justification on a TextFrame; return true if it actually changed (for no-op detection) */
    function applyJustification(targetItem, justification, force) {
        if (justification === null) {
            return false;
        }
        if (targetItem.constructor.name !== "TextFrame") {
            return false;
        }
        /* 既に目的の行揃えなら何もしない（無駄な undo ステップを作らない）。
           ただし確定時（force）は取得値に頼らず必ず代入する。undo 直後の
           paragraphAttributes は前の値を返すことがあり、代入を飛ばすと戻ってしまう /
           Skip if already at the target justification (avoids a spurious undo step), except on
           commit (force): paragraphAttributes can report the pre-undo value right after an undo,
           and skipping the assignment there would leave the justification reverted */
        if (force !== true && targetItem.textRange.paragraphAttributes.justification === justification) {
            return false;
        }
        if (justification === Justification.LEFT) {
            /* Illustrator のバグで Justification.LEFT の代入は無視される（RIGHT/CENTER は可）。
               一時的に resize して段落属性をリフレッシュさせると代入が効く。
               200%->代入->50% で実寸は元に戻り、位置も保存して戻す。
               Assigning Justification.LEFT is ignored by Illustrator; a temporary resize
               refreshes the paragraph attributes so the assignment takes effect (200% then
               50% leaves the real size unchanged; position is saved and restored). */
            var savedPosition = [targetItem.position[0], targetItem.position[1]];
            targetItem.resize(200, 200);
            targetItem.textRange.paragraphAttributes.justification = Justification.LEFT;
            targetItem.resize(50, 50);
            /* 位置を戻せない種類がある / some items refuse the position assignment */
            try {
                targetItem.position = savedPosition;
            } catch (ePos) {}
            return true;
        }
        targetItem.textRange.paragraphAttributes.justification = justification;
        return true;
    }

    /* 対象2点を選択から解決（2個選択 or 2点入りグループ1個）。解決不可は null /
       Resolve the two target items (two selected, or a single non-clip group of two). Null if unresolved */
    function resolveTargetPair(currentSelection) {
        if (currentSelection.length === 2) {
            return [currentSelection[0], currentSelection[1]];
        }
        /* 2点を含む通常グループ1つ（クリップグループは1オブジェクト扱いなので除外）/
           One regular group of exactly two items (clip groups count as a single object, so excluded) */
        if (currentSelection.length === 1 && currentSelection[0].constructor.name === "GroupItem" && currentSelection[0].clipped !== true) {
            var groupChildren = currentSelection[0].pageItems;
            if (groupChildren.length === 2) {
                return [groupChildren[0], groupChildren[1]];
            }
        }
        return null;
    }

    /* 選択2点の現在の上下間隔を pt で返す（メインエンジンで実行）/ Return the current vertical gap (pt) of the two selected items */
    function measureGap(adjustOptions) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        var targetPair = resolveTargetPair(app.activeDocument.selection);
        if (targetPair === null) {
            return "NOSEL";
        }
        var boundsA = getItemBounds(targetPair[0], adjustOptions.usePreviewBounds);
        var boundsB = getItemBounds(targetPair[1], adjustOptions.usePreviewBounds);
        /* top が大きい方が上 / The object with the larger top is the upper one */
        var upperBounds, lowerBounds;
        if (boundsA[1] >= boundsB[1]) {
            upperBounds = boundsA;
            lowerBounds = boundsB;
        } else {
            upperBounds = boundsB;
            lowerBounds = boundsA;
        }
        /* 上のオブジェクトの下辺 − 下のオブジェクトの上辺（重なりは負）/
           Upper object's bottom edge − lower object's top edge (negative when overlapping) */
        return String(upperBounds[3] - lowerBounds[1]);
    }

    /* 直前のプレビューを取り消す（メインエンジンで実行）/ Undo the previous preview (runs in the main engine) */
    function undoLast() {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        app.undo();
        app.redraw();
        return "OK";
    }

    /* 1ペア（上下2点）に間隔・整列・行揃えを適用。実際に変更したら true /
       Apply gap, align and justify to one pair; return true if anything actually changed */
    function applyToPair(itemA, itemB, adjustOptions) {
        var changed = false;
        var MOVE_EPSILON = 0.0001;

        /* 行揃えを先に適用（ポイント文字は揃えで境界が変わるため）/ Justify first; point-text bounds depend on it */
        var justifyKey = resolveJustifyKey(adjustOptions.justify, adjustOptions.align);
        if (justifyKey !== "none") {
            var justification = resolveJustification(justifyKey);
            if (applyJustification(itemA, justification, adjustOptions.forceJustify)) {
                changed = true;
            }
            if (applyJustification(itemB, justification, adjustOptions.forceJustify)) {
                changed = true;
            }
        }

        var boundsA = getItemBounds(itemA, adjustOptions.usePreviewBounds);
        var boundsB = getItemBounds(itemB, adjustOptions.usePreviewBounds);

        /* top が大きい方が上 / The object with the larger top is the upper one */
        var upperItem, lowerItem;
        if (boundsA[1] >= boundsB[1]) {
            upperItem = itemA;
            lowerItem = itemB;
        } else {
            upperItem = itemB;
            lowerItem = itemA;
        }

        var anchorItem = adjustOptions.anchorTop ? upperItem : lowerItem;
        var movingItem = adjustOptions.anchorTop ? lowerItem : upperItem;

        /* 上下方向の移動（移動量が実質ゼロなら translate しない）/ Vertical move (skip if the delta is effectively zero) */
        var dy;
        if (adjustOptions.anchorTop) {
            var targetLowerTop = getItemBounds(upperItem, adjustOptions.usePreviewBounds)[3] - adjustOptions.gapPoints;
            dy = targetLowerTop - getItemBounds(movingItem, adjustOptions.usePreviewBounds)[1];
        } else {
            var targetUpperBottom = getItemBounds(lowerItem, adjustOptions.usePreviewBounds)[1] + adjustOptions.gapPoints;
            dy = targetUpperBottom - getItemBounds(movingItem, adjustOptions.usePreviewBounds)[3];
        }
        if (Math.abs(dy) > MOVE_EPSILON) {
            movingItem.translate(0, dy);
            changed = true;
        }

        /* 左右方向の整列（同上）/ Horizontal alignment (same zero-skip) */
        if (adjustOptions.align !== "none") {
            var anchorBounds = getItemBounds(anchorItem, adjustOptions.usePreviewBounds);
            var movingBounds = getItemBounds(movingItem, adjustOptions.usePreviewBounds);
            var dx = 0;
            if (adjustOptions.align === "left") {
                dx = anchorBounds[0] - movingBounds[0];
            } else if (adjustOptions.align === "right") {
                dx = anchorBounds[2] - movingBounds[2];
            } else if (adjustOptions.align === "center") {
                dx = ((anchorBounds[0] + anchorBounds[2]) / 2) - ((movingBounds[0] + movingBounds[2]) / 2);
            }
            if (Math.abs(dx) > MOVE_EPSILON) {
                movingItem.translate(dx, 0);
                changed = true;
            }
        }

        /* 整列後の左右ずらし（正＝右／負＝左）。整列「なし」でも適用 /
           Extra horizontal offset after alignment (positive = right, negative = left); applies even when align is none */
        if (adjustOptions.adjustPoints && Math.abs(adjustOptions.adjustPoints) > MOVE_EPSILON) {
            movingItem.translate(adjustOptions.adjustPoints, 0);
            changed = true;
        }

        return changed;
    }

    /* ライブプレビュー：選択ペア1組に適用（メインエンジンで実行）/ Live preview: apply to the single selected pair */
    function runAdjustment(adjustOptions) {
        /* ドキュメント確認を最優先（undo より先）/ Check for a document first, before any undo */
        if (app.documents.length === 0) {
            return "NODOC";
        }
        /* ライブプレビュー：前回適用分を取り消してからやり直す / Live preview: undo the previous apply first */
        if (adjustOptions.undoFirst === true) {
            app.undo();
        }
        var targetPair = resolveTargetPair(app.activeDocument.selection);
        if (targetPair === null) {
            app.redraw();
            return "NOSEL";
        }

        var changed = applyToPair(targetPair[0], targetPair[1], adjustOptions);

        app.redraw();
        /* 変更があれば OK（undo ステップ1つ）、無ければ NOCHANGE（undo ステップ無し）/
           OK if something changed (one undo step), otherwise NOCHANGE (no undo step) */
        return changed ? "OK" : "NOCHANGE";
    }

    /* 選択から一括適用の対象ペア群を集める（各グループの2点／グループ無しなら2点選択を1組）/
       Collect target pairs for batch apply (each group's two children; or two loose items as one pair) */
    function collectTargetPairs(currentSelection) {
        var targetPairs = [];
        for (var i = 0; i < currentSelection.length; i++) {
            if (currentSelection[i].constructor.name === "GroupItem" && currentSelection[i].clipped !== true && currentSelection[i].pageItems.length === 2) {
                targetPairs.push([currentSelection[i].pageItems[0], currentSelection[i].pageItems[1]]);
            }
        }
        /* グループが1つも無く、ちょうど2点選択なら単一ペア / no qualifying groups but exactly two loose items */
        if (targetPairs.length === 0 && currentSelection.length === 2) {
            targetPairs.push([currentSelection[0], currentSelection[1]]);
        }
        return targetPairs;
    }

    /* 記録した設定を選択中の全対象ペアへ一括適用・確定（プレビューなし）/
       Batch-apply the recorded settings to every target pair in the selection (committed, no preview) */
    function runBatchAdjustment(adjustOptions) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        /* プレビュー分はここで取り消す（別送信で取り消すとプレビューと条件が変わる）/
           Undo the preview here, in the same send as the apply */
        if (adjustOptions.undoFirst === true) {
            app.undo();
        }
        /* 確定なので行揃えは取得値で判定せず必ず適用する / Commit: always assign the justification */
        adjustOptions.forceJustify = true;
        var targetPairs = collectTargetPairs(app.activeDocument.selection);
        if (targetPairs.length === 0) {
            app.redraw();
            return "NOSEL";
        }
        for (var i = 0; i < targetPairs.length; i++) {
            applyToPair(targetPairs[i][0], targetPairs[i][1], adjustOptions);
        }
        app.redraw();
        return "OK";
    }

    /* オブジェクトの位置を [x, y] で返す（取得できなければ null）/ Return the item position as [x, y], or null */
    function getItemPosition(targetItem) {
        /* position を持たない・読めない種類がある / some items have no readable position */
        try {
            var itemPosition = targetItem.position;
            if (!itemPosition || itemPosition.length !== 2) {
                return null;
            }
            return [Number(itemPosition[0]), Number(itemPosition[1])];
        } catch (e) {
            return null;
        }
    }

    /* 判定で動かしたオブジェクトを元の位置へ戻す / Move every probed item back to its original position */
    function restoreItemPositions(targetItems, savedPositions) {
        var restored = true;
        for (var i = 0; i < targetItems.length; i++) {
            try {
                targetItems[i].position = [savedPositions[i][0], savedPositions[i][1]];
            } catch (e) {
                /* 1つ失敗しても残りは必ず戻す / Keep restoring the rest even if one fails */
                restored = false;
            }
        }
        return restored;
    }

    /* 整列で動いたオブジェクトを候補から外す / Drop candidates that moved during an align probe */
    function rejectMovedCandidates(targetItems, savedPositions, candidates, tolerance) {
        for (var i = 0; i < targetItems.length; i++) {
            if (!candidates[i]) {
                continue;
            }
            var currentPosition = getItemPosition(targetItems[i]);
            if (currentPosition === null ||
                Math.abs(currentPosition[0] - savedPositions[i][0]) > tolerance ||
                Math.abs(currentPosition[1] - savedPositions[i][1]) > tolerance) {
                candidates[i] = false;
            }
        }
    }

    /* 他のオブジェクトを完全に内包しているか。内包している側は整列コマンドで動かないため、
       キーオブジェクトが無くても4方向すべてで残ってしまう /
       Whether the item's bounds enclose every other item: such an item never moves during the
       probes, so it survives all four of them even when no key object is set */
    function enclosesAllOthers(targetItem, targetItems) {
        var itemBounds = getItemBounds(targetItem, false);
        for (var i = 0; i < targetItems.length; i++) {
            if (targetItems[i] === targetItem) {
                continue;
            }
            var otherBounds = getItemBounds(targetItems[i], false);
            /* [left, top, right, bottom] */
            if (otherBounds[0] < itemBounds[0] || otherBounds[2] > itemBounds[2] ||
                otherBounds[1] > itemBounds[1] || otherBounds[3] < itemBounds[3]) {
                return false;
            }
        }
        return true;
    }

    /* キーオブジェクトを判定（4方向の整列で動かなかった1点）。ExtendScript には
       キーオブジェクトを取得する API が無いため、整列コマンドを一時的に実行して
       位置が変わらないオブジェクトを探し、最後に必ず元の位置へ戻す /
       Detect the key object: the single item that stays put through all four align probes.
       ExtendScript has no API for it, so align commands are run temporarily and every
       item is moved back afterwards */
    function findKeyObject(targetItems) {
        var alignCommands = ["Horizontal Align Left", "Vertical Align Top", "Horizontal Align Right", "Vertical Align Bottom"];
        var tolerance = 0.001;
        var savedPositions = [];
        var candidates = [];
        var i;
        for (i = 0; i < targetItems.length; i++) {
            savedPositions[i] = getItemPosition(targetItems[i]);
            if (savedPositions[i] === null) {
                return null;
            }
            candidates[i] = true;
        }
        /* executeMenuCommand が失敗したら判定をあきらめる / give up when a menu command fails */
        try {
            for (i = 0; i < alignCommands.length; i++) {
                /* 2回目以降だけ元位置へ戻す（初回は余計な移動をしない）/ Restore only from the second probe on */
                if (i > 0 && !restoreItemPositions(targetItems, savedPositions)) {
                    return null;
                }
                app.executeMenuCommand(alignCommands[i]);
                /* 候補が1つになっても打ち切らず、4方向すべてで動かないことを確認する /
                   Keep probing all four directions to avoid a false positive */
                rejectMovedCandidates(targetItems, savedPositions, candidates, tolerance);
                /* 候補が尽きたら残りのプローブは結果を変えられないので打ち切る /
                   No candidate left: the remaining probes cannot change the result */
                var remaining = 0;
                for (var j = 0; j < candidates.length; j++) {
                    if (candidates[j]) {
                        remaining++;
                    }
                }
                if (remaining === 0) {
                    break;
                }
            }
        } catch (e) {
            return null;
        } finally {
            /* 判定中は redraw せず、最後に一度だけ元位置へ戻す / No redraw while probing; restore once at the end */
            restoreItemPositions(targetItems, savedPositions);
        }
        var keyItem = null;
        var candidateCount = 0;
        for (i = 0; i < candidates.length; i++) {
            if (candidates[i]) {
                keyItem = targetItems[i];
                candidateCount++;
            }
        }
        /* 一意に決まり、かつ内包による居座りでないときだけ採用 /
           Accept only a unique candidate that is not merely enclosing the others */
        if (candidateCount !== 1) {
            return null;
        }
        return enclosesAllOthers(keyItem, targetItems) ? null : keyItem;
    }

    /* 選択2点のキーオブジェクトが上下どちらかを返す / Report whether the key object is the upper or lower item */
    function findKeyObjectAnchor(adjustOptions) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        var currentSelection = app.activeDocument.selection;
        /* 整列コマンドは選択に対して働くため、2点選択のときだけ判定する /
           Align commands act on the selection, so only exactly two selected items qualify */
        if (!currentSelection || currentSelection.length !== 2) {
            return "NOSEL";
        }
        var keyItem = findKeyObject(currentSelection);
        app.redraw();
        if (keyItem === null) {
            return "NONE";
        }
        var keyBounds = getItemBounds(keyItem, adjustOptions.usePreviewBounds);
        var otherBounds = getItemBounds((currentSelection[0] === keyItem) ? currentSelection[1] : currentSelection[0], adjustOptions.usePreviewBounds);
        /* top が大きい方が上 / The item with the larger top is the upper one */
        return (keyBounds[1] >= otherBounds[1]) ? "TOP" : "BOTTOM";
    }

    // =========================================
    // BridgeTalk 委譲 / BridgeTalk delegation
    // =========================================

    /* 委譲する worker 関数（追加したらここにも登録）/ Worker functions to delegate (register new ones here) */
    var WORKER_FUNCS = [
        getClippingPath,
        getItemBounds,
        resolveJustifyKey,
        resolveJustification,
        applyJustification,
        resolveTargetPair,
        measureGap,
        undoLast,
        applyToPair,
        runAdjustment,
        collectTargetPairs,
        runBatchAdjustment
    ];

    /* キーオブジェクト判定だけで使う worker 関数（プレビューの送信量を増やさないため別立て）/
       Worker functions used only by key-object detection (kept apart so previews stay small) */
    var KEY_OBJECT_FUNCS = [
        getClippingPath,
        getItemBounds,
        getItemPosition,
        restoreItemPositions,
        rejectMovedCandidates,
        enclosesAllOthers,
        findKeyObject,
        findKeyObjectAnchor
    ];

    /**
     * 設定を JS のオブジェクトリテラル文字列にする
     * @param {object} adjustOptions - 設定（anchorTop / gapPoints / align / adjustPoints / justify / usePreviewBounds / undoFirst）
     * @returns {string} オブジェクトリテラル
     */
    function optionsToLiteral(adjustOptions) {
        return "{"
            + "anchorTop:" + adjustOptions.anchorTop + ","
            + "gapPoints:" + adjustOptions.gapPoints + ","
            + "align:\"" + adjustOptions.align + "\","
            + "adjustPoints:" + (adjustOptions.adjustPoints || 0) + ","
            + "justify:\"" + adjustOptions.justify + "\","
            + "usePreviewBounds:" + adjustOptions.usePreviewBounds + ","
            + "undoFirst:" + (adjustOptions.undoFirst === true)
            + "}";
    }

    /**
     * worker 関数群のソースと呼び出し式をつなげたコードを作る
     * @param {function[]} workerFuncs - 送る worker 関数
     * @param {string} dispatchExpr - 最後に評価する呼び出し式
     * @returns {string} メインエンジンで評価するコード
     */
    function buildWorkerCode(workerFuncs, dispatchExpr) {
        var functionSources = [];
        for (var i = 0; i < workerFuncs.length; i++) {
            functionSources.push(workerFuncs[i].toString());
        }
        return functionSources.join("\n") + "\n" + dispatchExpr + ";";
    }

    /**
     * メインエンジンへ同期で委譲する（% エンコードで文字化けを防ぐ）
     * @param {string} workerCode - 評価するコード
     * @returns {string} 結果の文字列（失敗時は "ERR:" で始まる）
     */
    function delegateToMainEngine(workerCode) {
        var bridgeTalk = new BridgeTalk();
        bridgeTalk.target = "illustrator";
        bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerCode) + "\"));";
        var resultHolder = { value: null };
        bridgeTalk.onResult = function (response) {
            resultHolder.value = response.body;
        };
        bridgeTalk.onError = function (response) {
            resultHolder.value = "ERR:" + response.body;
        };
        bridgeTalk.send(10); /* 同期送信（最大10秒）/ Synchronous send (up to 10s) */
        return (resultHolder.value === null) ? "ERR:timeout" : resultHolder.value;
    }

    /**
     * プレビューを適用する（前回分は worker 側で取り消す）
     * @param {object} adjustOptions - 設定
     * @returns {string} OK / NOCHANGE / NODOC / NOSEL / ERR:…
     */
    function runAdjustmentPreview(adjustOptions) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "runAdjustment(" + optionsToLiteral(adjustOptions) + ")"));
    }

    /**
     * 設定を選択中の全対象ペアへ一括適用・確定する
     * @param {object} adjustOptions - 設定
     * @returns {string} OK / NODOC / NOSEL / ERR:…
     */
    function runBatchAdjustmentDelegate(adjustOptions) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "runBatchAdjustment(" + optionsToLiteral(adjustOptions) + ")"));
    }

    /**
     * 選択2点の現在の間隔を測る
     * @param {object} adjustOptions - 設定（usePreviewBounds を使う）
     * @returns {string} pt の数値文字列 / NODOC / NOSEL
     */
    function measureCurrentGap(adjustOptions) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "measureGap(" + optionsToLiteral(adjustOptions) + ")"));
    }

    /**
     * キーオブジェクトが上下どちらかを判定する
     * @param {boolean} usePreviewBounds - プレビュー境界で比べるか
     * @returns {string} TOP / BOTTOM / NONE / NODOC / NOSEL
     */
    function detectKeyObjectAnchor(usePreviewBounds) {
        return delegateToMainEngine(buildWorkerCode(KEY_OBJECT_FUNCS, "findKeyObjectAnchor({usePreviewBounds:" + (usePreviewBounds === true) + "})"));
    }

    /**
     * 直前のプレビューを取り消す
     * @returns {string} OK / NODOC / ERR:…
     */
    function revertLastPreview() {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "undoLast()"));
    }

    // =========================================
    // パネル生成 / Panel builders
    // =========================================

    /**
     * 定規単位の数値欄（∧∨・↑↓キー対応）と単位ラベルを行に追加する
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 初期値
     * @param {string} tooltipText - ツールチップ
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @param {function} onPreview - 値が変わったときの処理
     * @returns {EditText} 追加した入力欄
     */
    function addUnitField(parentRow, initialText, tooltipText, rulerUnit, onPreview) {
        /* ∧∨と入力欄は隙間0で突き合わせる。負の値も許容（重ねる・左へずらす）/ Stepper butts the field; negatives allowed */
        var stepperFieldGroup = parentRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;
        var unitField;
        var unitFieldStepper = addStepper(stepperFieldGroup, function () { return unitField; }, {
            onStep: function () { onPreview(); }
        });
        unitField = stepperFieldGroup.add("edittext", undefined, initialText);
        unitField.characters = FIELD_CHARS;
        unitField.helpTip = tooltipText;
        bindSteppedArrowKeys(unitField, unitFieldStepper); /* ↑↓キーも∧∨と同じ処理 / arrow keys share the stepper */
        unitField.onChange = onPreview;
        parentRow.add("statictext", undefined, rulerUnit.label);
        return unitField;
    }

    /**
     * ラジオの組すべてに同じツールチップと onClick を付ける
     * @param {RadioButton[]} radioList - 対象のラジオ
     * @param {string} tooltipText - ツールチップ
     * @param {function} [clickHandler] - クリック時の処理（省略時は付けない）
     * @returns {void}
     */
    function setupRadios(radioList, tooltipText, clickHandler) {
        for (var i = 0; i < radioList.length; i++) {
            radioList[i].helpTip = tooltipText;
            if (clickHandler) radioList[i].onClick = clickHandler;
        }
    }

    /**
     * キーオブジェクトのパネル（上・下・自動判定）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {function} onPreview - 上・下を選んだときの処理
     * @param {function} onAutoDetect - 自動判定を選んだときの処理
     * @returns {{anchorTopRadio: RadioButton, anchorBottomRadio: RadioButton, anchorAutoRadio: RadioButton, autoAnchorTop: boolean}} ラジオと自動判定の結果
     */
    function buildAnchorPanel(parentGroup, onPreview, onAutoDetect) {
        var anchorPanel = parentGroup.add("panel", undefined, getLabel("panel.anchor"));
        setupPanel(anchorPanel, 6);
        anchorPanel.helpTip = getLabel("tooltip.anchorTop") + " / " + getLabel("tooltip.anchorBottom") + " / " + getLabel("tooltip.anchorAuto");

        /* 上・下・自動判定は横並び（ショートカット T/B/K はラベル非表示）/ Top, bottom and auto in a row (T/B/K shortcuts are not shown) */
        var anchorRow = anchorPanel.add("group");
        setupGroup(anchorRow, "row");
        var anchorTopRadio = anchorRow.add("radiobutton", undefined, getLabel("radio.anchorTop"));
        var anchorBottomRadio = anchorRow.add("radiobutton", undefined, getLabel("radio.anchorBottom"));
        var anchorAutoRadio = anchorRow.add("radiobutton", undefined, getLabel("radio.anchorAuto"));
        anchorAutoRadio.value = true; /* 既定は自動判定 / Auto by default */
        setupRadios([anchorTopRadio], getLabel("tooltip.anchorTop"), onPreview);
        setupRadios([anchorBottomRadio], getLabel("tooltip.anchorBottom"), onPreview);
        setupRadios([anchorAutoRadio], getLabel("tooltip.anchorAuto"), onAutoDetect);

        return {
            anchorTopRadio: anchorTopRadio,
            anchorBottomRadio: anchorBottomRadio,
            anchorAutoRadio: anchorAutoRadio,
            autoAnchorTop: true /* 自動判定の結果（上が基準なら true）/ Detection result (true when the top item is the key object) */
        };
    }

    /**
     * 上下間隔の値とプレビュー境界のパネルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @param {function} onPreview - 値が変わったときの処理
     * @returns {{gapValueInput: EditText, previewBoundsCheckbox: Checkbox}} 入力欄とチェックボックス
     */
    function buildGapPanel(parentGroup, rulerUnit, onPreview) {
        var gapPanel = parentGroup.add("panel", undefined, getLabel("panel.gap"));
        setupPanel(gapPanel, 6);
        gapPanel.helpTip = getLabel("tooltip.gap");

        var gapRow = gapPanel.add("group");
        setupGroup(gapRow, "row");
        var gapValueInput = addUnitField(gapRow, DEFAULT_GAP_VALUE, getLabel("tooltip.gap"), rulerUnit, onPreview);

        var previewBoundsCheckbox = gapPanel.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.value = true;
        previewBoundsCheckbox.helpTip = getLabel("tooltip.previewBounds");
        previewBoundsCheckbox.onClick = onPreview;

        return {
            gapValueInput: gapValueInput,
            previewBoundsCheckbox: previewBoundsCheckbox
        };
    }

    /**
     * 左右の整列のパネル（なし・左・中央・右と横調整）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @param {function} onPreview - 値が変わったときの処理
     * @returns {object} none / left / center / right のラジオ、adjustInput、selectCenter()
     */
    function buildAlignPanel(parentGroup, rulerUnit, onPreview) {
        var alignPanel = parentGroup.add("panel", undefined, getLabel("panel.align"));
        setupPanel(alignPanel, 6);
        alignPanel.helpTip = getLabel("tooltip.align");

        /* ラジオは横並び / Radios in a row */
        var alignRow = alignPanel.add("group");
        setupGroup(alignRow, "row");
        var alignRadios = {
            none: alignRow.add("radiobutton", undefined, getLabel("radio.alignNone")),
            left: alignRow.add("radiobutton", undefined, getLabel("radio.alignLeft")),
            center: alignRow.add("radiobutton", undefined, getLabel("radio.alignCenter")),
            right: alignRow.add("radiobutton", undefined, getLabel("radio.alignRight"))
        };
        alignRadios.none.value = true;
        setupRadios([alignRadios.none, alignRadios.left, alignRadios.right], getLabel("tooltip.align"), onPreview);
        setupRadios([alignRadios.center], getLabel("tooltip.align"));

        /* 整列後の左右ずらし量（正＝右／負＝左）/ Extra horizontal offset (positive = right, negative = left) */
        var adjustRow = alignPanel.add("group");
        setupGroup(adjustRow, "row");
        adjustRow.add("statictext", undefined, labelText("fieldLabel.alignAdjust"));
        var adjustInput = addUnitField(adjustRow, "0", getLabel("tooltip.alignAdjust"), rulerUnit, onPreview);
        alignRadios.adjustInput = adjustInput;

        /**
         * 「中央」を選び、横調整を0へ戻してプレビューする（マウス・キー操作共通）
         * @returns {void}
         */
        function selectCenter() {
            alignRadios.center.value = true;
            adjustInput.text = "0";
            onPreview();
        }
        alignRadios.center.onClick = selectCenter;
        alignRadios.selectCenter = selectCenter;

        return alignRadios;
    }

    /**
     * テキストの行揃えのパネル（変更しない・整列に連動・均等配置）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {function} onPreview - 選択が変わったときの処理
     * @returns {{none: RadioButton, link: RadioButton, full: RadioButton}} ラジオ
     */
    function buildJustifyPanel(parentGroup, onPreview) {
        var justifyPanel = parentGroup.add("panel", undefined, getLabel("panel.justify"));
        setupPanel(justifyPanel, 6);
        justifyPanel.helpTip = getLabel("tooltip.justify");

        var justifyRadios = {
            none: justifyPanel.add("radiobutton", undefined, getLabel("radio.justifyNone")),
            link: justifyPanel.add("radiobutton", undefined, getLabel("radio.justifyLink")),
            full: justifyPanel.add("radiobutton", undefined, getLabel("radio.justifyFull"))
        };
        justifyRadios.link.value = true;
        setupRadios([justifyRadios.none, justifyRadios.link], getLabel("tooltip.justify"), onPreview);
        /* （最終行左）はツールチップに / "(last line left)" lives in the tooltip */
        setupRadios([justifyRadios.full], getLabel("tooltip.justifyFull"), onPreview);
        return justifyRadios;
    }

    /**
     * ［記録］［適用］のボタン行を追加する
     * @param {Window} parentWindow - 追加先のパレット
     * @param {function} onRecord - ［記録］／［編集］の処理
     * @param {function} onApply - ［適用］の処理
     * @returns {Button} ［記録］ボタン（ロック時に表示名を切り替える）
     */
    function buildButtonRow(parentWindow, onRecord, onApply) {
        var buttonRow = addButtonRow(parentWindow);
        var btnRecord = buttonRow.rightGroup.add("button", undefined, getLabel("button.record"));
        btnRecord.helpTip = getLabel("tooltip.record");
        btnRecord.onClick = onRecord;
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLabel("button.apply"));
        btnApply.helpTip = getLabel("tooltip.apply");
        btnApply.onClick = onApply;
        alignRightOnlyButtonRow(buttonRow);
        return btnRecord;
    }

    /**
     * 選択中のラジオに対応するキーを返す
     * @param {object} radioSet - キー → ラジオ
     * @param {string[]} radioKeys - 調べる順のキー（どれも選ばれていなければ先頭）
     * @returns {string} 選ばれているラジオのキー
     */
    function selectedRadioKey(radioSet, radioKeys) {
        for (var i = 0; i < radioKeys.length; i++) {
            if (radioSet[radioKeys[i]].value) {
                return radioKeys[i];
            }
        }
        return radioKeys[0];
    }

    /**
     * 上のオブジェクトを基準にするかを決める（自動判定は直近の判定結果を使う）
     * @param {object} anchorControls - buildAnchorPanel() の結果
     * @returns {boolean} 上を基準にするなら true
     */
    function resolveAnchorTop(anchorControls) {
        if (anchorControls.anchorAutoRadio.value) {
            return anchorControls.autoAnchorTop;
        }
        return anchorControls.anchorTopRadio.value;
    }

    /**
     * パレットの値を読み取り、間隔と横調整を pt に換算する
     * @param {object} paletteControls - 各パネルのコントロール
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @returns {object} 設定（anchorTop / gapPoints / align / adjustPoints / justify / usePreviewBounds）
     */
    function readAdjustOptions(paletteControls, rulerUnit) {
        var gapValue = parseFloat(paletteControls.gap.gapValueInput.text);
        if (isNaN(gapValue)) {
            gapValue = parseFloat(DEFAULT_GAP_VALUE);
        }
        /* 負の値は重なり（オーバーラップ）として許容し、正規化した値を表示にも反映 /
           Negative values are allowed (objects overlap); reflect the normalized value back to the field */
        if (paletteControls.gap.gapValueInput.text !== String(gapValue)) {
            paletteControls.gap.gapValueInput.text = gapValue;
        }

        var adjustValue = parseFloat(paletteControls.align.adjustInput.text);
        if (isNaN(adjustValue)) {
            adjustValue = 0;
        }

        return {
            anchorTop: resolveAnchorTop(paletteControls.anchor),
            gapPoints: gapValue * rulerUnit.pointsPerUnit,
            align: selectedRadioKey(paletteControls.align, ["none", "left", "center", "right"]),
            adjustPoints: adjustValue * rulerUnit.pointsPerUnit,
            justify: selectedRadioKey(paletteControls.justify, ["none", "link", "full"]),
            usePreviewBounds: paletteControls.gap.previewBoundsCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * パレットを組み立てて表示する
     * @returns {Window} 表示したパレット
     */
    function showPalette() {
        /* 多重起動を防ぐ：既存パレットがあれば閉じる / Prevent duplicates: close any existing palette */
        if (paletteWindow) {
            /* 破棄済みのウィンドウは close() が失敗することがある / close() may fail on a disposed window */
            try {
                paletteWindow.close();
            } catch (e) {}
            paletteWindow = null;
        }

        var rulerUnit = getUnitInfo();

        var gapPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        setupWindow(gapPalette);

        var paletteControls = {};
        var previewState = { active: false }; /* プレビューが反映中か / Whether a preview is currently applied */
        var isBusy = false; /* 同期委譲中の再入防止 / Guard against re-entry during a synchronous delegation */
        var recordedOptions = null; /* 記録した設定（間隔・キー・整列・行揃え・境界）/ The recorded recipe */
        var isRecordLocked = false; /* 記録モード（パネルをロック中）か / Whether we are in recorded/locked mode */
        var settingsColumn = null;
        var btnRecord = null;

        /**
         * Illustrator のキーオブジェクトを判定して、自動判定の基準を更新する
         * @returns {boolean} 上下どちらかに決まったら true
         */
        function loadAnchorFromKeyObject() {
            var detectResult = detectKeyObjectAnchor(paletteControls.gap.previewBoundsCheckbox.value);
            if (detectResult === "TOP" || detectResult === "BOTTOM") {
                paletteControls.anchor.autoAnchorTop = (detectResult === "TOP");
                return true;
            }
            paletteControls.anchor.autoAnchorTop = true;
            if (detectResult === "NOSEL" || detectResult === "NODOC") {
                /* 2点選択以外は判定そのものができない。グループ選択や複数ペアでも「自動」を
                   残したまま既定（上）を使う / Detection needs exactly two selected items;
                   keep Auto selected for groups and batches and fall back to the default (top) */
                return false;
            }
            /* 2点選択でキーオブジェクトが無いときだけ「上」へ切り替え、実際に使う基準を見せる /
               Switch to Top only when two items are selected but no key object exists */
            paletteControls.anchor.anchorTopRadio.value = true;
            return false;
        }

        /**
         * 「自動判定」：プレビューを戻してから判定し直し、プレビューする
         * （判定は整列コマンドでドキュメントを触るため、先にプレビューを戻す）
         * @returns {void}
         */
        function refreshAutoAnchor() {
            if (isBusy) return; /* 委譲中の割り込みを無視（同期送信中もUIは動く）/ ignore clicks during a delegation */
            isBusy = true;
            try {
                paletteControls.anchor.anchorAutoRadio.value = true;
                revertActivePreview();
                loadAnchorFromKeyObject();
            } finally {
                isBusy = false;
            }
            /* updatePreview 自身が再入ガードを張るのでガードの外で呼ぶ / updatePreview sets the guard itself */
            updatePreview();
        }

        /**
         * 選択2点の現在の間隔を入力欄へ取り込む
         * @returns {boolean} 取り込めたら true
         */
        function loadGapFromSelection() {
            var measureResult = measureCurrentGap({
                anchorTop: true,
                gapPoints: 0,
                align: "none",
                justify: "none",
                usePreviewBounds: paletteControls.gap.previewBoundsCheckbox.value,
                undoFirst: false
            });
            var gapPoints = parseFloat(measureResult);
            if (isNaN(gapPoints)) {
                return false; /* NODOC / NOSEL など、計測できず初期値のまま / could not measure; keep default */
            }
            /* 重なり（負の間隔）もそのまま取り込む。pt → 定規単位、0.1単位（小数1桁）に丸め /
               Keep negative gaps (overlap) as-is; pt → ruler unit, rounded to 0.1 (1 decimal) */
            var gapValue = Math.round((gapPoints / rulerUnit.pointsPerUnit) * 10) / 10;
            paletteControls.gap.gapValueInput.text = String(gapValue);
            return true;
        }

        /**
         * ライブプレビューを更新する（前回分は worker 側で取り消す）
         * @returns {void}
         */
        function updatePreview() {
            if (isBusy) return; /* 委譲中に発火した変更は無視 / ignore changes fired mid-delegation */
            isBusy = true;
            try {
                var adjustOptions = readAdjustOptions(paletteControls, rulerUnit);
                adjustOptions.undoFirst = previewState.active;
                var previewResult = runAdjustmentPreview(adjustOptions);
                /* OK＝変更あり（undoステップ1つ）。それ以外（NOCHANGE／エラー）は undo ステップが
                   無いので active=false にし、次回 app.undo() で直前のユーザー操作を巻き戻さない /
                   OK = changed (one undo step). Otherwise (NOCHANGE / error) there is no undo step,
                   so keep active false so the next app.undo() won't revert the user's prior action. */
                previewState.active = (previewResult === "OK");
            } finally {
                /* 委譲が例外で抜けても再入ガードを必ず解除 / Always clear the re-entry guard, even on error */
                isBusy = false;
            }
        }

        /**
         * プレビュー中なら戻して、ドキュメントを元の状態にする
         * @returns {void}
         */
        function revertActivePreview() {
            if (previewState.active) {
                revertLastPreview();
                previewState.active = false;
            }
        }

        /**
         * パネルのロック表示を切り替える（ボタン名・ツールチップ・ディムを連動。ボタンエリアは常に有効）
         * @param {boolean} shouldLock - ロックするなら true
         * @returns {void}
         */
        function setLocked(shouldLock) {
            isRecordLocked = shouldLock;
            settingsColumn.enabled = !shouldLock;
            redrawSteppersIn(settingsColumn); /* ∧∨のディム表示を切り替える / update stepper dimming */
            btnRecord.text = shouldLock ? getLabel("button.edit") : getLabel("button.record");
            btnRecord.helpTip = shouldLock ? getLabel("tooltip.edit") : getLabel("tooltip.record");
        }

        /**
         * 「記録」⇔「編集」の切り替え：記録時は設定を控えてパネルをディム、編集時はロック解除（プレビューは戻さない）
         * @returns {void}
         */
        function toggleRecord() {
            if (!isRecordLocked) {
                recordedOptions = readAdjustOptions(paletteControls, rulerUnit);
                /* プレビューを確定（取り消さず保持）。active=false にして、次の［適用］で
                   app.undo() が走り選択が壊れるのを防ぐ / Commit the preview (keep it) and clear
                   active so the next Apply won't app.undo() and clobber the selection */
                previewState.active = false;
                setLocked(true);
            } else {
                setLocked(false);
            }
        }

        /**
         * 「適用」：記録した設定（無ければ現在値）を選択中の全対象ペアへ一括適用・確定する
         * @returns {void}
         */
        function applyBatch() {
            var adjustOptions = recordedOptions ? recordedOptions : readAdjustOptions(paletteControls, rulerUnit);
            /* プレビュー分の取り消しは worker 側でまとめて行う（二重適用を防ぐ）/
               The worker undoes the preview in the same send (avoids double-applying it) */
            adjustOptions.undoFirst = previewState.active;
            var batchResult = runBatchAdjustmentDelegate(adjustOptions);
            /* 送信に失敗した場合はプレビューが適用されたまま残るので active を維持する /
               On a failed send the preview is still applied, so keep tracking it */
            if (String(batchResult).indexOf("ERR:") !== 0) {
                previewState.active = false;
            }
        }

        /**
         * ロック中か（記録中は true）。パネル系ショートカットの抑止に使う
         * @returns {boolean} ロック中なら true
         */
        function isLockedNow() {
            return isRecordLocked;
        }

        /* 1カラム：間隔値・キーオブジェクト・左右の整列・テキストの行揃えを縦に並べる /
           Single column: gap, key object, horizontal align, text alignment stacked vertically */
        settingsColumn = gapPalette.add("group");
        setupGroup(settingsColumn, "column");
        paletteControls.gap = buildGapPanel(settingsColumn, rulerUnit, updatePreview);
        paletteControls.anchor = buildAnchorPanel(settingsColumn, updatePreview, refreshAutoAnchor);
        paletteControls.align = buildAlignPanel(settingsColumn, rulerUnit, updatePreview);
        paletteControls.justify = buildJustifyPanel(settingsColumn, updatePreview);

        /* ボタン：記録／適用 / Buttons: Record / Apply */
        btnRecord = buildButtonRow(gapPalette, toggleRecord, applyBatch);

        /* 閉じる時：未確定のプレビューは取り消す（×・Esc 共通）/ On close: revert an uncommitted preview (X and Esc) */
        gapPalette.onClose = function () {
            revertActivePreview();
            $.global[PALETTE_GLOBAL_KEY] = null; /* 常駐エンジンの参照をクリア / Clear the persistent-engine reference */
            return true;
        };

        /* キー操作：A で適用、T/B で固定対象を選択、K で自動判定、N/L/C/R で整列、S/J で行揃え、Esc で閉じる
           Keys: A applies, T/B pick the anchor, K re-detects, N/L/C/R align, S/J justify, Esc closes */
        /**
         * ロック中はキーを使うだけで何もしない処理を作る（パネル系ショートカット用）
         * @param {Object|Function} shortcutTarget - ラジオ、または関数
         * @returns {Function} addKeyShortcuts に渡す関数
         */
        function unlessLocked(shortcutTarget) {
            return function () {
                if (isLockedNow()) return true;
                if (typeof shortcutTarget === "function") return shortcutTarget();
                return shortcutTarget;
            };
        }

        addKeyShortcuts(gapPalette, {
            "A": applyBatch,
            "Escape": { target: function () { gapPalette.close(); }, inFields: true },
            "T": unlessLocked(paletteControls.anchor.anchorTopRadio),
            "B": unlessLocked(paletteControls.anchor.anchorBottomRadio),
            "K": unlessLocked(refreshAutoAnchor),
            "N": unlessLocked(paletteControls.align.none),
            "L": unlessLocked(paletteControls.align.left),
            "C": unlessLocked(paletteControls.align.center),
            "R": unlessLocked(paletteControls.align.right),
            "S": unlessLocked(paletteControls.justify.link),
            "J": unlessLocked(paletteControls.justify.full)
        }, {
            numericFields: [paletteControls.gap.gapValueInput, paletteControls.align.adjustInput]
        });

        gapPalette.center();
        gapPalette.show();
        loadAnchorFromKeyObject(); /* 「自動判定」の初期値としてキーオブジェクトを判定 / Detect the key object for the initial Auto anchor */
        loadGapFromSelection(); /* 現在の間隔を取り込む（取り込めれば初期プレビューで動かない）/ Load current gap (no movement on initial preview when available) */
        updatePreview(); /* 初期プレビュー / Initial preview */
        paletteControls.gap.gapValueInput.active = true; /* 開いたら間隔値にフォーカス / Focus the gap field on open */
        return gapPalette;
    }

    /* 常駐エンジンにパレット参照を保持するキー（CloseAllPalettes.jsx から閉じるため）
       Key holding the palette reference in the persistent engine (so CloseAllPalettes.jsx can close it) */
    var PALETTE_GLOBAL_KEY = "__aiAdjustVerticalGapPalette";

    paletteWindow = showPalette();
    $.global[PALETTE_GLOBAL_KEY] = paletteWindow;

})();
