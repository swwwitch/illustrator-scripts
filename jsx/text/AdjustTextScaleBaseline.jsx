#target illustrator
#targetengine "AdjustTextScaleBaselineEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの文字比率とベースラインシフトを調整します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustTextScaleBaseline.md

### Overview

Adjusts the character scale and the baseline shift of the selected text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustTextScaleBaseline.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustTextScaleBaseline";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustTextScaleBaseline.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustTextScaleBaseline.md"; /* README (English) */

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

    var TARGET_CHAR_INPUT_CHARACTERS = 10;   /* 対象文字欄の幅（文字数）/ Width of the target character field */
    var NUMBER_INPUT_CHARACTERS = 4;         /* 数値欄の幅（文字数）/ Width of the number fields */
    var APPARENT_SIZE_CHARACTERS = 5;        /* 見かけのサイズ表示の幅（文字数）/ Width of the apparent size display */
    var ROW_LABEL_WIDTH = 120;               /* 行ラベルの幅 / Width of the row labels */
    var BUTTON_WIDTH = 90;                   /* ボタンの幅 / Button width */
    var CANCEL_RESET_GAP = 50;               /* ［キャンセル］と［リセット］の間隔 / Gap between Cancel and Reset */

    // =========================================
    // プレビュー / Preview
    // =========================================

    var PREVIEW_INTERVAL_MS = 80; /* 入力中のプレビューを間引く間隔（ミリ秒）/ Throttle interval for previews while typing (ms) */

    /* ∧∨・↑↓キーで増減したあとは、手入力の確定と同じ onChange を通す（プレビュー更新）
       After stepping, fire the same onChange as a typed commit (refreshes the preview) */
    var STEP_NOTIFY_CHANGE = function (numberInput) { numberInput.notify("onChange"); };

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

    var LABELS = {
        dialog: {
            title: {
                ja: "フォントサイズとベースライン調整 " + SCRIPT_VERSION,
                en: "Font Size & Baseline Adjuster " + SCRIPT_VERSION
            }
        },
        panel: {
            adjust: { ja: "調整", en: "Adjust" }
        },
        fieldLabel: {
            targetChar: { ja: "対象文字", en: "Target Char" },
            fontSize: { ja: "フォントサイズ", en: "Font Size" },
            scale: { ja: "水平比率/垂直比率", en: "Scale" },
            apparent: { ja: "見かけ", en: "Apparent" },
            baselineShift: { ja: "ベースラインシフト", en: "Baseline Shift" },
            kerning: { ja: "カーニング", en: "Kerning" },
            tracking: { ja: "トラッキング", en: "Tracking" }
        },
        tooltip: {
            targetChar: {
                ja: "ここに書いた文字だけを調整します。空欄にすると選択範囲すべてが対象です。",
                en: "Adjusts only the characters listed here. Leave blank to affect the whole selection."
            },
            fontSize: { ja: "対象文字のフォントサイズを増減します。", en: "Changes the font size of the target characters." },
            scale: {
                ja: "対象文字の長体・平体です。100%で変形なしになります。",
                en: "Horizontal and vertical scale of the target characters. 100% means no distortion."
            },
            apparentSize: {
                ja: "フォントサイズに比率を掛けた、見かけの文字サイズを表示します。比率が100%のときは淡色表示になります。",
                en: "Shows the apparent type size: the font size multiplied by the scale. Dimmed when the scale is 100%."
            },
            baselineShift: {
                ja: "対象文字を上下にずらす量です。負の値で下がります。",
                en: "How far the target characters move up. Negative values move them down."
            },
            kerning: { ja: "対象文字の前後の詰めです。単位は1/1000em。", en: "Spacing around the target characters, in 1/1000 em." },
            tracking: { ja: "選択範囲全体の字間です。単位は1/1000em。", en: "Letter spacing across the whole selection, in 1/1000 em." },
            reset: {
                ja: "対象文字とフォントサイズを開いたときの値に戻し、比率を100%、ベースラインシフト・カーニング・トラッキングを0にします。",
                en: "Restores the target characters and font size to their initial values, sets the scale to 100% and baseline shift, kerning and tracking to 0."
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
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" }
        },
        alert: {
            selectText: { ja: "テキストを選択してください", en: "Please select text" }
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

    // =========================================
    // 対象の収集 / Collecting targets
    // =========================================

    /**
     * 選択から調整対象の TextRange を集める（テキストフレームは全文、文字の選択はその範囲）
     * @param {Array|TextRange} currentSelection - 現在の選択
     * @returns {TextRange[]} 調整対象の TextRange
     */
    function collectSelectedTextRanges(currentSelection) {
        var textRanges = [];
        if (!currentSelection || currentSelection.length === 0) return textRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem instanceof TextFrame) {
                textRanges.push(selectedItem.textRange);
            } else if (selectedItem instanceof TextRange) {
                textRanges.push(selectedItem);
            }
        }
        /* 文字の編集中は選択そのものが TextRange / While editing text, the selection itself is a TextRange */
        if (!(currentSelection instanceof Array) && currentSelection instanceof TextRange) {
            textRanges.push(currentSelection);
        }
        return textRanges;
    }

    /**
     * すべての文字列に共通して含まれる文字を、最初の文字列での出現順に重複なく返す
     * @param {string[]} texts - 文字列の配列
     * @returns {string} 共通の文字
     */
    function getCommonCharacters(texts) {
        if (texts.length === 0) return "";
        var commonChars = texts[0];
        for (var i = 1; i < texts.length; i++) {
            var otherText = texts[i];
            var sharedChars = "";
            for (var j = 0; j < commonChars.length; j++) {
                var currentChar = commonChars.charAt(j);
                if (otherText.indexOf(currentChar) !== -1 && sharedChars.indexOf(currentChar) === -1) {
                    sharedChars += currentChar;
                }
            }
            commonChars = sharedChars;
            if (commonChars.length === 0) break;
        }
        return commonChars;
    }

    /**
     * 英数字を除いた文字を、重複なく出現順に返す
     * @param {string} text - 元の文字列
     * @returns {string} 英数字以外の文字（重複なし）
     */
    function getUniqueNonAlphanumerics(text) {
        var strippedText = text.replace(/[0-9A-Za-z]/g, "");
        var uniqueChars = "";
        for (var i = 0; i < strippedText.length; i++) {
            var currentChar = strippedText.charAt(i);
            if (uniqueChars.indexOf(currentChar) === -1) {
                uniqueChars += currentChar;
            }
        }
        return uniqueChars;
    }

    /**
     * 対象文字の初期値を求める（すべての選択範囲に共通する、英数字以外の文字）
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @returns {string} 対象文字の初期値
     */
    function findDefaultTargetChars(targetRanges) {
        var rangeTexts = [];
        for (var i = 0; i < targetRanges.length; i++) {
            rangeTexts.push(targetRanges[i].contents);
        }
        return getUniqueNonAlphanumerics(getCommonCharacters(rangeTexts));
    }

    // =========================================
    // 文字の調整 / Character adjustment
    // =========================================

    /**
     * 文字が対象文字に当たるかを判定する（対象文字が空ならすべて対象）
     * @param {TextRange} textChar - 1文字の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @returns {boolean} 対象なら true
     */
    function isTargetChar(textChar, charFilter) {
        return !charFilter.hasTargetChars || charFilter.targetChars.indexOf(textChar.contents) !== -1;
    }

    /**
     * 対象文字の入力欄から、対象文字の指定を読み取る
     * @param {EditText} targetCharInput - 対象文字の入力欄
     * @returns {{targetChars: string, hasTargetChars: boolean}} 対象文字の指定
     */
    function readCharFilter(targetCharInput) {
        var targetChars = targetCharInput.text;
        return { targetChars: targetChars, hasTargetChars: targetChars.length > 0 };
    }

    /**
     * すべての TextRange で、対象文字に当たる文字それぞれに処理を行う
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @param {Function} charCallback - 1文字の TextRange を受け取る処理
     * @returns {void}
     */
    function forEachTargetChar(targetRanges, charFilter, charCallback) {
        for (var i = 0; i < targetRanges.length; i++) {
            var rangeChars = targetRanges[i].characters;
            for (var j = 0; j < rangeChars.length; j++) {
                if (isTargetChar(rangeChars[j], charFilter)) charCallback(rangeChars[j]);
            }
        }
    }

    /**
     * 対象文字に当たる最初の文字を返す
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @returns {TextRange|null} 最初の対象文字（無ければ null）
     */
    function findFirstTargetChar(targetRanges, charFilter) {
        for (var i = 0; i < targetRanges.length; i++) {
            var rangeChars = targetRanges[i].characters;
            for (var j = 0; j < rangeChars.length; j++) {
                if (isTargetChar(rangeChars[j], charFilter)) return rangeChars[j];
            }
        }
        return null;
    }

    /**
     * 対象文字にサイズ・比率・ベースラインシフト・カーニング・トラッキングを適用する（null の項目は触らない）
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @param {{size: ?number, scale: ?number, baseline: ?number, kerning: ?number, tracking: ?number}} adjustParams - 適用する値
     * @returns {void}
     */
    function applyTextAdjustments(targetRanges, charFilter, adjustParams) {
        if (adjustParams.size !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.size = adjustParams.size;
            });
        }
        if (adjustParams.scale !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.horizontalScale = adjustParams.scale;
                textChar.characterAttributes.verticalScale = adjustParams.scale;
            });
        }
        if (adjustParams.baseline !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                textChar.characterAttributes.baselineShift = adjustParams.baseline;
            });
        }
        if (adjustParams.kerning !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                textChar.kerning = adjustParams.kerning;
            });
        }
        if (adjustParams.tracking !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.tracking = adjustParams.tracking;
            });
        }
    }

    /**
     * 見かけの文字サイズ（フォントサイズ×比率）を求める
     * @param {number} fontSize - フォントサイズ
     * @param {number} scalePercent - 比率（%）
     * @returns {number} 見かけの文字サイズ（小数第2位まで）
     */
    function calculateApparentSize(fontSize, scalePercent) {
        return Math.round(fontSize * scalePercent) / 100;
    }

    // =========================================
    // プレビューの取り消し管理 / Preview undo management
    // =========================================

    /**
     * プレビュー時に Undo 履歴を汚さないための小さな管理クラス
     * - addStep(): 変更処理を実行して undoDepth を数える
     * - rollback(): プレビューで行った変更をすべて取り消す
     * - confirm(finalAction): 一度戻してから本番処理を1回だけ実行し、Undo を1回にまとめる
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;

        /**
         * 変更処理を1ステップとして実行し、Undo の深さを数える
         * @param {Function} stepAction - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (stepAction) {
            /* 文字属性の書き込みは DOM が例外を投げうる / Writing character attributes can throw */
            try {
                stepAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert("Preview Error: " + e);
            }
        };

        /**
         * プレビューで行った変更をすべて取り消す
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * プレビューを戻してから本番処理を1回だけ実行する
         * @param {Function} finalAction - 確定時に実行する処理
         * @returns {void}
         */
        this.confirm = function (finalAction) {
            this.rollback();
            finalAction();
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 右揃えの行ラベル＋数値欄＋単位の1行を追加する
     * @param {Group|Panel} parentGroup - 追加先
     * @param {string} labelPath - 行ラベルのパス
     * @param {string} tooltipPath - 数値欄の tooltip のパス
     * @param {string} initialText - 数値欄の初期値
     * @param {string} unitText - 単位の表示
     * @param {Object} stepOptions - ∧∨の設定（min / integer / onStep など。addStepper() を参照）
     * @returns {EditText} 追加した数値欄（行は .numberRow で取れる）
     */
    function addNumberRow(parentGroup, labelPath, tooltipPath, initialText, unitText, stepOptions) {
        var numberRow = parentGroup.add("group");
        numberRow.orientation = "row";
        var rowLabel = numberRow.add("statictext", undefined, getLabel(labelPath));
        rowLabel.justify = "right";
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = numberRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;
        var numberInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperFieldGroup.add("edittext", undefined, initialText);
        numberInput.characters = NUMBER_INPUT_CHARACTERS;
        numberInput.helpTip = getLabel(tooltipPath);
        numberInput.numberRow = numberRow;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);
        numberRow.add("statictext", undefined, unitText);
        return numberInput;
    }

    /**
     * ベースラインシフト・カーニング・トラッキングの行（値を右揃え）を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - 行ラベルのパス
     * @param {string} tooltipPath - 数値欄の tooltip のパス
     * @param {string} unitText - 単位の表示
     * @param {Object} stepOptions - ∧∨の設定（addStepper() を参照）
     * @returns {EditText} 追加した数値欄
     */
    function addShiftRow(parentPanel, labelPath, tooltipPath, unitText, stepOptions) {
        var shiftInput = addNumberRow(parentPanel, labelPath, tooltipPath, "0", unitText, stepOptions);
        shiftInput.numberRow.alignChildren = ["right", "center"];
        shiftInput.justify = "right";
        return shiftInput;
    }

    /**
     * ダイアログを組み立てる（イベントはまだ付けない）
     * @param {string} defaultTargetChars - 対象文字の初期値
     * @param {string} unitLabel - 文字サイズの単位
     * @returns {Object} ダイアログと各コントロール
     */
    function buildDialog(defaultTargetChars, unitLabel) {
        var adjustDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(adjustDialog);

        var contentGroup = adjustDialog.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = "top";
        contentGroup.spacing = COLUMN_SPACING;

        var leftColumn = contentGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "left";
        leftColumn.spacing = COLUMN_SPACING;

        var rightColumn = contentGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "right";

        /* 対象文字 / Target characters */
        var targetCharGroup = leftColumn.add("group");
        targetCharGroup.orientation = "row";
        targetCharGroup.add("statictext", undefined, labelText("fieldLabel.targetChar"));
        var targetCharInput = targetCharGroup.add("edittext", undefined, defaultTargetChars);
        targetCharInput.characters = TARGET_CHAR_INPUT_CHARACTERS;
        targetCharInput.helpTip = getLabel("tooltip.targetChar");

        var adjustPanel = leftColumn.add("panel", undefined, getLabel("panel.adjust"));
        setupPanel(adjustPanel);
        adjustPanel.alignChildren = ["left", "top"];

        /* フォントサイズと比率 / Font size and scale */
        var sizeScaleGroup = adjustPanel.add("group");
        sizeScaleGroup.orientation = "column";
        sizeScaleGroup.alignChildren = "left";

        /* フォントサイズは0より大きく、比率は1%以上 / font size above 0, scale at least 1% */
        var sizeInput = addNumberRow(sizeScaleGroup, "fieldLabel.fontSize", "tooltip.fontSize", "0", unitLabel,
            { min: 0.1, onStep: STEP_NOTIFY_CHANGE });
        sizeInput.numberRow.margins = [0, 10, 0, 0];
        sizeInput.numberRow.alignChildren = "left";

        var hScaleInput = addNumberRow(sizeScaleGroup, "fieldLabel.scale", "tooltip.scale", "100", "%",
            { min: 1, onStep: STEP_NOTIFY_CHANGE });
        hScaleInput.numberRow.margins = [0, 0, 0, 6];

        /* 見かけのサイズ（表示のみ）/ Apparent size (display only) */
        var apparentGroup = adjustPanel.add("group");
        apparentGroup.orientation = "row";
        var apparentLabel = apparentGroup.add("statictext", undefined, getLabel("fieldLabel.apparent"));
        apparentLabel.justify = "right";
        apparentLabel.preferredSize.width = ROW_LABEL_WIDTH;
        var apparentSizeText = apparentGroup.add("statictext", undefined, "--");
        apparentSizeText.characters = APPARENT_SIZE_CHARACTERS;
        apparentSizeText.helpTip = getLabel("tooltip.apparentSize");
        var apparentUnitLabel = apparentGroup.add("statictext", undefined, unitLabel);

        /* 負の値も可。カーニング・トラッキングは整数（1/1000em） / negatives allowed; kerning and tracking are integers */
        var baselineInput = addShiftRow(adjustPanel, "fieldLabel.baselineShift", "tooltip.baselineShift", unitLabel,
            { onStep: STEP_NOTIFY_CHANGE });
        var kerningInput = addShiftRow(adjustPanel, "fieldLabel.kerning", "tooltip.kerning", "/1000",
            { integer: true, onStep: STEP_NOTIFY_CHANGE });
        var trackingInput = addShiftRow(adjustPanel, "fieldLabel.tracking", "tooltip.tracking", "/1000",
            { integer: true, onStep: STEP_NOTIFY_CHANGE });

        /* ボタン（右列）/ Buttons (right column) */
        var btnColumnGroup = rightColumn.add("group");
        btnColumnGroup.alignment = "right";
        btnColumnGroup.orientation = "column";
        var btnOK = btnColumnGroup.add("button", undefined, getLabel("button.ok"));
        var btnCancel = btnColumnGroup.add("button", undefined, getLabel("button.cancel"));
        var cancelResetSpacer = btnColumnGroup.add("statictext", undefined, "");
        cancelResetSpacer.preferredSize.height = CANCEL_RESET_GAP;
        var btnReset = btnColumnGroup.add("button", undefined, getLabel("button.reset"));
        btnReset.helpTip = getLabel("tooltip.reset");
        btnReset.preferredSize.width = BUTTON_WIDTH;
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;

        return {
            dialog: adjustDialog,
            targetCharInput: targetCharInput,
            sizeInput: sizeInput,
            hScaleInput: hScaleInput,
            apparentLabel: apparentLabel,
            apparentSizeText: apparentSizeText,
            apparentUnitLabel: apparentUnitLabel,
            baselineInput: baselineInput,
            kerningInput: kerningInput,
            trackingInput: trackingInput,
            btnOK: btnOK,
            btnCancel: btnCancel,
            btnReset: btnReset
        };
    }

    /**
     * 入力欄の数値をまとめて読み取る（数値でない欄は null）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {{size: ?number, scale: ?number, baseline: ?number, kerning: ?number, tracking: ?number}} 入力値
     */
    function readAdjustParams(dialogControls) {
        var sizeValue = parseFloat(dialogControls.sizeInput.text);
        var scaleValue = parseFloat(dialogControls.hScaleInput.text);
        var baselineValue = parseFloat(dialogControls.baselineInput.text);
        var kerningValue = parseFloat(dialogControls.kerningInput.text);
        var trackingValue = parseFloat(dialogControls.trackingInput.text);
        return {
            size: isNaN(sizeValue) ? null : sizeValue,
            scale: isNaN(scaleValue) ? null : scaleValue,
            baseline: isNaN(baselineValue) ? null : baselineValue,
            kerning: isNaN(kerningValue) ? null : kerningValue,
            tracking: isNaN(trackingValue) ? null : trackingValue
        };
    }

    /**
     * 見かけのサイズの表示を更新する（比率 100% のときは淡色表示）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {void}
     */
    function updateApparentSizeDisplay(dialogControls) {
        var fontSize = parseFloat(dialogControls.sizeInput.text);
        var scalePercent = parseFloat(dialogControls.hScaleInput.text);
        if (isNaN(fontSize) || isNaN(scalePercent)) {
            dialogControls.apparentSizeText.text = "--";
        } else {
            dialogControls.apparentSizeText.text = calculateApparentSize(fontSize, scalePercent) + "";
        }

        var isDimmed = (scalePercent === 100);
        dialogControls.apparentLabel.enabled = !isDimmed;
        dialogControls.apparentSizeText.enabled = !isDimmed;
        dialogControls.apparentUnitLabel.enabled = !isDimmed;
    }

    /**
     * ダイアログにイベントを付け、初期値の読み込みと最初のプレビューを行う
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {string} defaultTargetChars - 対象文字の初期値
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, targetRanges, defaultTargetChars) {
        var adjustDialog = dialogControls.dialog;
        var targetCharInput = dialogControls.targetCharInput;
        var sizeInput = dialogControls.sizeInput;
        var hScaleInput = dialogControls.hScaleInput;
        var previewManager = new PreviewManager();
        var initialFontSize = null; /* 最初に読み込んだフォントサイズ（リセット用）/ First font size read, for Reset */
        var lastPreviewTime = 0;

        /**
         * 入力欄の値を対象文字に適用する
         * @returns {void}
         */
        function applyCurrentValues() {
            applyTextAdjustments(targetRanges, readCharFilter(targetCharInput), readAdjustParams(dialogControls));
        }

        /**
         * プレビューを取り消してから掛け直す
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(applyCurrentValues);
        }

        /**
         * プレビューを間引いて更新する（onChanging の連打で Undo と再適用が過剰にならないように）
         * @returns {void}
         */
        function updatePreviewThrottled() {
            var now = (new Date()).getTime();
            if (now - lastPreviewTime < PREVIEW_INTERVAL_MS) return;
            lastPreviewTime = now;
            updatePreview();
        }

        /**
         * 最初の対象文字のサイズと比率を入力欄に読み込む
         * @returns {void}
         */
        function loadFirstTargetCharValues() {
            var firstChar = findFirstTargetChar(targetRanges, readCharFilter(targetCharInput));
            if (!firstChar) {
                sizeInput.text = "";
                hScaleInput.text = "";
                dialogControls.apparentSizeText.text = "--";
                return;
            }
            var fontSize = firstChar.size;
            if (initialFontSize === null) {
                initialFontSize = fontSize;
            }
            sizeInput.text = (Math.round(fontSize * 10) / 10) + "";
            hScaleInput.text = (Math.round(firstChar.characterAttributes.horizontalScale * 10) / 10) + "";
            updateApparentSizeDisplay(dialogControls);
        }

        /**
         * サイズ・比率の確定を受けて、プレビューしてから値を読み直す
         * @returns {void}
         */
        function onSizeOrScaleChange() {
            updatePreview();
            loadFirstTargetCharValues();
            updateApparentSizeDisplay(dialogControls);
        }

        /* ベースラインシフト・カーニング・トラッキング / Baseline shift, kerning and tracking */
        var shiftInputs = [dialogControls.baselineInput, dialogControls.kerningInput, dialogControls.trackingInput];
        for (var i = 0; i < shiftInputs.length; i++) {
            shiftInputs[i].onChange = updatePreview;
            shiftInputs[i].onChanging = updatePreviewThrottled;
        }

        sizeInput.onChange = onSizeOrScaleChange;
        sizeInput.onChanging = function () { updateApparentSizeDisplay(dialogControls); };

        hScaleInput.onChange = onSizeOrScaleChange;
        hScaleInput.onChanging = function () { updateApparentSizeDisplay(dialogControls); };

        targetCharInput.onChange = function () {
            loadFirstTargetCharValues();
            updatePreview();
        };

        dialogControls.btnReset.onClick = function () {
            targetCharInput.text = defaultTargetChars;
            if (initialFontSize !== null) {
                sizeInput.text = initialFontSize;
            }
            hScaleInput.text = "100";
            dialogControls.baselineInput.text = "0";
            dialogControls.kerningInput.text = "0";
            dialogControls.trackingInput.text = "0";

            updatePreview();
            loadFirstTargetCharValues();
        };

        dialogControls.btnCancel.onClick = function () {
            previewManager.rollback();
            adjustDialog.close(2);
        };

        dialogControls.btnOK.onClick = function () {
            /* Undo を1回にまとめて確定 / Confirm as a single undo step */
            previewManager.confirm(applyCurrentValues);
            targetCharInput.onChange = null;
            adjustDialog.close();
        };

        adjustDialog.onShow = function () {
            hScaleInput.active = true;
        };

        loadFirstTargetCharValues();
        updateApparentSizeDisplay(dialogControls);

        /* 初期状態も管理下で1回プレビュー適用（この時点で undoDepth=1 になる）
           Apply the initial state once under management (undoDepth becomes 1 here) */
        updatePreview();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のテキストを集めてダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length <= 0) {
            return;
        }

        var targetRanges = collectSelectedTextRanges(app.activeDocument.selection);
        if (targetRanges.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var defaultTargetChars = findDefaultTargetChars(targetRanges);
        var dialogControls = buildDialog(defaultTargetChars, getUnitInfo("text/units").label);
        bindDialogEvents(dialogControls, targetRanges, defaultTargetChars);
        prepareDialogWindow(dialogControls.dialog, SCRIPT_NAME);
        dialogControls.dialog.show();
    }

    main();

})();
