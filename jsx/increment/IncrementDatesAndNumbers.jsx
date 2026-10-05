#target illustrator
#targetengine "IncrementDatesAndNumbersEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のテキストに含まれる日付・曜日・連番・数値などを、一括して増減します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/IncrementDatesAndNumbers.md

### Overview

Increments or decrements the dates, weekday names, sequence numbers and other values found in the selected text, all at once.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/IncrementDatesAndNumbers.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IncrementDatesAndNumbers";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-11-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/IncrementDatesAndNumbers.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/IncrementDatesAndNumbers.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_STEP_VALUE = 1; /* 増減値の初期値（［対象］を切り替えたときもこの値に戻す） / initial step, also restored when the target changes */

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

    var PREVIEW_MARGINS       = [13, 10, 0, 10];   /* オリジナル／結果の行の余白 / margins of the preview rows */
    var PREVIEW_SPACING       = 6;                 /* オリジナル／結果の行どうしの間隔 / spacing between the preview rows */
    var PREVIEW_LABEL_SPACING = 6;                 /* 項目名と値の間隔 / gap between label and value */
    var TARGET_RADIO_SPACING  = 6;                 /* ［対象］パネルのラジオボタンの間隔 / spacing of the target radio buttons */
    var STEP_FIELD_CHARACTERS = 5;                 /* 増減値の入力欄の幅（文字数） / width of the step field (characters) */

    /* 結果欄の幅を測る見本。オリジナルが短いときもこの長さを確保する
       Sample that sizes the result field; used when the original text is short */
    var RESULT_WIDTH_SAMPLE     = "0000年00月00日";
    var RESULT_SAMPLE_MIN_CHARS = 8;

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

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 }, /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 }, /* 歯 / H */
        "p": { unit: "pc", amount: 1 }     /* パイカ / pica */
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
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
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
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
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
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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
            title: { ja: "日付・数値の増減", en: "Increment Dates and Numbers" }
        },
        panel: {
            step: { ja: "増減", en: "Step" },
            mode: { ja: "種別", en: "Type" },
            target: { ja: "対象", en: "Target" }
        },
        fieldLabel: {
            original: { ja: "オリジナル", en: "Original" },
            result: { ja: "結果", en: "Result" },
            step: { ja: "増減値", en: "Amount" }
        },
        radio: {
            modeNumber: { ja: "数字", en: "Number" },
            modeDate: { ja: "日付", en: "Date" }
        },
        tooltip: {
            step: {
                ja: "1回の増減で足す量です（整数）。負の値を入れると減らせます。",
                en: "Amount added per step (whole numbers). Enter a negative value to count down."
            },
            modeNumber: { ja: "「12.1」を数字とみなして増減します。", en: "Treats \"12.1\" as a number." },
            modeDate: {
                ja: "「12.1」を「月.日」の日付とみなして増減します。",
                en: "Treats \"12.1\" as a month.day date."
            },
            target: {
                ja: "どの部分を増減するかです。切り替えると増減値は1に戻ります。",
                en: "Which part to step. Switching resets the amount to 1."
            },
            result: {
                ja: "選択中の最初のテキストの1行目を例に、結果を表示します。複数行のテキストは行ごとに増減します。",
                en: "Previews the first line of the first selected text. Multi-line text is stepped line by line."
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
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            errorTitle: { ja: "エラー", en: "Error" },
            errorGeneric: { ja: "エラーが発生しました：", en: "An error occurred:" }
        }
    };

    // =========================================
    // 曜日・元号 / Weekdays and eras
    // =========================================
    var JAPANESE_WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
    var ENGLISH_WEEKDAYS  = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    var CIRCLED_WEEKDAYS  = ["㊐", "㊊", "㊋", "㊌", "㊍", "㊎", "㊏"];
    /* 元号と元年の西暦 / era name -> first year */
    var ERA_START_YEARS = {
        "令和": 2019,
        "平成": 1989,
        "昭和": 1926,
        "大正": 1912,
        "明治": 1868
    };
    var CURRENT_YEAR = new Date().getFullYear(); /* 年の無い月日は今年とみなす / month-day values assume this year */
    var LINE_BREAK = String.fromCharCode(13);

    // =========================================
    // 状態 / State
    // =========================================
    var stepValue = DEFAULT_STEP_VALUE; /* 1回の増減量 / amount per step */
    var shiftTarget = "day";            /* 増減する部分（"year" / "month" / "day"） / part to shift */

    /* ドット区切り2要素（例：12.1）の解釈（"number" または "date"）
       Mode for interpreting 2-part dot patterns like 12.1 ("number" or "date") */
    var dotPairMode = "number";

    // =========================================
    // 選択の解析 / Selection analysis
    // =========================================

    /**
     * オブジェクトから最初のテキストフレームを探す（グループ内も再帰）
     * @param {PageItem} pageItem - 探索対象
     * @returns {TextFrame|null} 見つかったテキストフレーム
     */
    function findFirstTextFrame(pageItem) {
        if (!pageItem) return null;
        if (pageItem.typename === "TextFrame") return pageItem;
        if (pageItem.typename === "GroupItem") {
            var childItems = pageItem.pageItems;
            for (var i = 0; i < childItems.length; i++) {
                var foundFrame = findFirstTextFrame(childItems[i]);
                if (foundFrame) return foundFrame;
            }
        }
        return null;
    }

    /**
     * 選択から最初のテキストフレームを探す
     * @param {PageItem[]} selectedItems - ドキュメントの選択
     * @returns {TextFrame|null} 見つかったテキストフレーム
     */
    function findFirstTextFrameInSelection(selectedItems) {
        /* 文字ツールで文字を選択中は TextRange が返り、[0] が無い / a TextRange (type tool) has no [0] */
        if (!selectedItems || selectedItems.typename === "TextRange") return null;
        for (var i = 0; i < selectedItems.length; i++) {
            var foundFrame = findFirstTextFrame(selectedItems[i]);
            if (foundFrame) return foundFrame;
        }
        return null;
    }

    /**
     * テキスト全体が年月日・ドット区切り・時刻のどれかなら、分割した各部分を返す
     * （［対象］パネルのラジオボタンの表示に使う）
     * @param {string} trimmedText - 前後の空白を除いたテキスト
     * @returns {{year: string, month: string, day: string, isDotPair: boolean}|null} 各部分。該当しなければ null
     */
    function splitDateParts(trimmedText) {
        /* 2025年11月21日のような和文年月日 / Japanese Y/M/D */
        var kanjiMatch = trimmedText.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日(?:[（(［\[]?(?:日|月|火|水|木|金|土)[）)\]］]?)?$/);
        if (kanjiMatch) return { year: kanjiMatch[1], month: kanjiMatch[2], day: kanjiMatch[3], isDotPair: false };

        /* 2025/11/21 や 2025/11/21（金）のようなスラッシュ区切り / Slash-separated date with optional weekday suffix */
        var slashMatch = trimmedText.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:[ 　\t]*[（(［\[]?(?:日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[）)\]］]?)?$/);
        if (slashMatch) return { year: slashMatch[1], month: slashMatch[2], day: slashMatch[3], isDotPair: false };

        /* 29.3.2 / 29.4 などのドット区切り / dot-separated */
        var dotTripleMatch = trimmedText.match(/^(\d+)\.(\d+)\.(\d+)$/);
        if (dotTripleMatch) return { year: dotTripleMatch[1], month: dotTripleMatch[2], day: dotTripleMatch[3], isDotPair: false };

        /* 2要素のドット区切り（例：12.1）は「数字／日付」のどちらにもなり得るため、［種別］パネルを出す
           Two-part dot patterns (e.g. 12.1) can be numeric or date, so enable mode selection */
        var dotPairMatch = trimmedText.match(/^(\d+)\.(\d+)$/);
        if (dotPairMatch) return { year: dotPairMatch[1], month: dotPairMatch[2], day: "", isDotPair: true };

        /* 時刻 19:00 のようなパターンも分割対象とする（年=時、月=分として扱う）
           Treat time like 19:00 as a split target (map year->hour, month->minute) */
        var timeMatch = trimmedText.match(/^(\d{1,2}):(\d{2})$/);
        if (timeMatch) return { year: timeMatch[1], month: timeMatch[2], day: "", isDotPair: false };

        return null;
    }

    /**
     * 選択中の最初のテキストフレームから、プレビューの見本と［対象］パネルの内容を読み取る
     * @returns {{originalSample: string, hasDateParts: boolean, isAmbiguousDotPair: boolean, yearPart: string, monthPart: string, dayPart: string}} 読み取り結果
     */
    function readSelectionSample() {
        var selectionSample = {
            originalSample: "",
            hasDateParts: false,
            isAmbiguousDotPair: false,
            yearPart: "",
            monthPart: "",
            dayPart: ""
        };
        if (app.documents.length === 0) return selectionSample;

        var targetFrame = findFirstTextFrameInSelection(app.activeDocument.selection);
        if (!targetFrame) return selectionSample;

        var trimmedText = trimText(targetFrame.contents);
        selectionSample.originalSample = trimmedText.split(LINE_BREAK)[0];

        var dateParts = splitDateParts(trimmedText);
        if (dateParts) {
            selectionSample.hasDateParts = true;
            selectionSample.isAmbiguousDotPair = dateParts.isDotPair;
            selectionSample.yearPart = dateParts.year;
            selectionSample.monthPart = dateParts.month;
            selectionSample.dayPart = dateParts.day;
        }
        return selectionSample;
    }

    // =========================================
    // 増減の計算 / Increment logic
    // =========================================

    /**
     * 前後の空白を除く
     * @param {string} sourceText - 元のテキスト
     * @returns {string} 前後の空白を除いたテキスト
     */
    function trimText(sourceText) {
        return sourceText.replace(/^\s+|\s+$/g, "");
    }

    /**
     * 数値を指定の桁数までゼロ埋めする
     * @param {number} numberValue - 数値
     * @param {number} digitCount - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function padZero(numberValue, digitCount) {
        var paddedText = String(numberValue);
        while (paddedText.length < digitCount) paddedText = "0" + paddedText;
        return paddedText;
    }

    /**
     * 数値文字列の整数部に3桁区切りのカンマを入れる
     * @param {string} numberText - 数値文字列
     * @returns {string} カンマ入りの文字列
     */
    function addThousandsCommas(numberText) {
        var numberParts = numberText.split(".");
        var integerPart = numberParts[0];
        var decimalPart = numberParts.length > 1 ? "." + numberParts[1] : "";
        var isNegative = integerPart.charAt(0) === "-";
        if (isNegative) integerPart = integerPart.substr(1);
        var withCommas = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        return (isNegative ? "-" : "") + withCommas + decimalPart;
    }

    /**
     * 増減量を整数の段数にする（0 は 0、正は 1 以上、負は -1 以下）
     * @param {number} stepAmount - 増減量
     * @returns {number} 整数の段数
     */
    function toWholeStep(stepAmount) {
        if (stepAmount === 0) return 0;
        if (stepAmount > 0) return Math.max(1, Math.round(stepAmount));
        return Math.min(-1, Math.round(stepAmount));
    }

    /**
     * 「年・月・日」の数字から日付を作る（月は1始まり）
     * @param {string|number} yearText - 年
     * @param {string} monthText - 月
     * @param {string} dayText - 日
     * @returns {Date} 日付
     */
    function makeDate(yearText, monthText, dayText) {
        return new Date(parseInt(yearText, 10), parseInt(monthText, 10) - 1, parseInt(dayText, 10));
    }

    /**
     * ［対象］の選択に応じて、日付の年・月・日のどれかを増減する
     * @param {Date} targetDate - 対象の日付（直接書き換える）
     * @param {number} stepInt - 増減する段数
     * @returns {void}
     */
    function shiftDateByTarget(targetDate, stepInt) {
        if (shiftTarget === "year") {
            targetDate.setFullYear(targetDate.getFullYear() + stepInt);
        } else if (shiftTarget === "month") {
            targetDate.setMonth(targetDate.getMonth() + stepInt);
        } else {
            targetDate.setDate(targetDate.getDate() + stepInt);
        }
    }

    /**
     * 日付を「11月21日」の形にする
     * @param {Date} targetDate - 日付
     * @returns {string} 月日の文字列
     */
    function formatKanjiMonthDay(targetDate) {
        return (targetDate.getMonth() + 1) + "月" + targetDate.getDate() + "日";
    }

    /**
     * 「11月21日㊎」のような月日＋曜日記号を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftMonthDayWithSymbol(lineText, stepInt) {
        var dateMatch = lineText.match(/(\d{1,2})月(\d{1,2})日(㊐|㊊|㊋|㊌|㊍|㊎|㊏)/);
        if (!dateMatch) return null;
        var shiftedDate = makeDate(CURRENT_YEAR, dateMatch[1], dateMatch[2]);
        shiftDateByTarget(shiftedDate, stepInt);
        return lineText.replace(dateMatch[0], formatKanjiMonthDay(shiftedDate) + CIRCLED_WEEKDAYS[shiftedDate.getDay()]);
    }

    /**
     * 「11月21日（金）」のような月日＋曜日を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftMonthDayWithWeekday(lineText, stepInt) {
        var dateMatch = lineText.match(/(\d{1,2})月(\d{1,2})日([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (!dateMatch) return null;
        var shiftedDate = makeDate(CURRENT_YEAR, dateMatch[1], dateMatch[2]);
        shiftDateByTarget(shiftedDate, stepInt);
        var newText = formatKanjiMonthDay(shiftedDate) + (dateMatch[3] || "") + JAPANESE_WEEKDAYS[shiftedDate.getDay()] + (dateMatch[5] || "");
        return lineText.replace(dateMatch[0], newText);
    }

    /**
     * 「令和7年11月21日（金）」のような和暦の年月日＋曜日を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftEraDateWithWeekday(lineText, stepInt) {
        var dateMatch = lineText.match(/(明治|大正|昭和|平成|令和)(\d{1,2})年(\d{1,2})月(\d{1,2})日([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (!dateMatch) return null;
        var shiftedDate = makeDate(ERA_START_YEARS[dateMatch[1]] + parseInt(dateMatch[2], 10) - 1, dateMatch[3], dateMatch[4]);
        shiftDateByTarget(shiftedDate, stepInt);
        var newEra = dateMatch[1];
        for (var eraName in ERA_START_YEARS) {
            if (shiftedDate.getFullYear() >= ERA_START_YEARS[eraName]) newEra = eraName;
        }
        var eraYear = shiftedDate.getFullYear() - ERA_START_YEARS[newEra] + 1;
        var newText = newEra + eraYear + "年" + formatKanjiMonthDay(shiftedDate) + (dateMatch[5] || "") + JAPANESE_WEEKDAYS[shiftedDate.getDay()] + (dateMatch[7] || "");
        return lineText.replace(dateMatch[0], newText);
    }

    /**
     * 「2025/11/21 (Fri)」「2025.11.21（金）」のような区切り付き日付＋曜日を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftSeparatedDateWithWeekday(lineText, stepInt) {
        var dateMatch = lineText.match(/(\d{4})([\/.])(\d{1,2})\2(\d{1,2})(\s*)([（([\[]?)(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)([）)\]]?)/);
        if (!dateMatch) return null;
        var dateSeparator = dateMatch[2];
        var weekdaySpacing = dateMatch[5] || "";
        var openBracket = dateMatch[6] || "";
        var weekdayText = dateMatch[7];
        var closeBracket = dateMatch[8] || "";
        var shiftedDate = makeDate(dateMatch[1], dateMatch[3], dateMatch[4]);
        shiftDateByTarget(shiftedDate, stepInt);
        var newDateText = shiftedDate.getFullYear() + dateSeparator + (shiftedDate.getMonth() + 1) + dateSeparator + shiftedDate.getDate();
        var weekdayNames = /^[日月火水木金土]$/.test(weekdayText) ? JAPANESE_WEEKDAYS : ENGLISH_WEEKDAYS;
        var originalText = dateMatch[1] + dateSeparator + dateMatch[3] + dateSeparator + dateMatch[4] + weekdaySpacing + openBracket + weekdayText + closeBracket;
        var replacedText = newDateText + weekdaySpacing + openBracket + weekdayNames[shiftedDate.getDay()] + closeBracket;
        return lineText.replace(originalText, replacedText);
    }

    /* 曜日の無い数字だけの日付（20251121 / 2025/11/21 / 2025.11.21） / numeric dates without a weekday */
    var NUMERIC_DATE_FORMATS = [
        { regex: /(\d{4})(\d{2})(\d{2})/, separator: "" },
        { regex: /(\d{4})\/(\d{1,2})\/(\d{1,2})/, separator: "/" },
        { regex: /(\d{4})\.(\d{1,2})\.(\d{1,2})/, separator: "." }
    ];

    /**
     * 「20251121」「2025/11/21」「2025.11.21」のような数字だけの日付を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftNumericDate(lineText, stepInt) {
        for (var i = 0; i < NUMERIC_DATE_FORMATS.length; i++) {
            var dateFormat = NUMERIC_DATE_FORMATS[i];
            var dateMatch = lineText.match(dateFormat.regex);
            if (!dateMatch) continue;
            var shiftedDate = makeDate(dateMatch[1], dateMatch[2], dateMatch[3]);
            shiftDateByTarget(shiftedDate, stepInt);
            /* 8桁の日付はゼロ埋め、区切り付きはゼロ埋めしない / 8-digit dates are zero-padded, separated ones are not */
            var newText = (dateFormat.separator === "") ?
                padZero(shiftedDate.getFullYear(), 4) + padZero(shiftedDate.getMonth() + 1, 2) + padZero(shiftedDate.getDate(), 2) :
                shiftedDate.getFullYear() + dateFormat.separator + (shiftedDate.getMonth() + 1) + dateFormat.separator + shiftedDate.getDate();
            return lineText.replace(dateFormat.regex, newText);
        }
        return null;
    }

    /**
     * 「2025年11月20日(木)」のような曜日付きの年月日を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftKanjiDateWithWeekday(lineText, stepInt) {
        var dateMatch = lineText.match(/(\d{4})年(\d{1,2})月(\d{1,2})日([（(［\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (!dateMatch) return null;
        var shiftedDate = makeDate(dateMatch[1], dateMatch[2], dateMatch[3]);
        shiftDateByTarget(shiftedDate, stepInt);
        var newText = shiftedDate.getFullYear() + "年" + formatKanjiMonthDay(shiftedDate) + (dateMatch[4] || "") + JAPANESE_WEEKDAYS[shiftedDate.getDay()] + (dateMatch[6] || "");
        return lineText.replace(dateMatch[0], newText);
    }

    /**
     * 「2025年11月20日」のような曜日なしの年月日を増減する
     * （直後にかっこが続くものは曜日付きとして別に扱うので除く）
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftKanjiDate(lineText, stepInt) {
        var dateMatch = lineText.match(/(\d{4})年(\d{1,2})月(\d{1,2})日(?![（(［\[])/);
        if (!dateMatch) return null;
        var shiftedDate = makeDate(dateMatch[1], dateMatch[2], dateMatch[3]);
        shiftDateByTarget(shiftedDate, stepInt);
        return lineText.replace(dateMatch[0], shiftedDate.getFullYear() + "年" + formatKanjiMonthDay(shiftedDate));
    }

    /* 日付として増減するパターン（上から順に試し、最初に当たったものだけを使う）
       Date patterns, tried in order; the first match wins */
    var DATE_SHIFTERS = [
        shiftMonthDayWithSymbol,
        shiftMonthDayWithWeekday,
        shiftEraDateWithWeekday,
        shiftSeparatedDateWithWeekday,
        shiftNumericDate,
        shiftKanjiDateWithWeekday,
        shiftKanjiDate
    ];

    /**
     * 単独の曜日記号と曜日名を、増減の段数だけ送る
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string} 置き換えたテキスト
     */
    function shiftWeekdayMarks(lineText, stepInt) {
        var weekdayShift = stepInt % 7;
        if (weekdayShift < 0) weekdayShift += 7;

        var updatedText = lineText;
        var symbolMatch = updatedText.match(/(㊐|㊊|㊋|㊌|㊍|㊎|㊏)/);
        if (symbolMatch) {
            var symbolIndex = CIRCLED_WEEKDAYS.indexOf(symbolMatch[1]);
            if (symbolIndex !== -1) {
                updatedText = updatedText.replace(symbolMatch[1], CIRCLED_WEEKDAYS[(symbolIndex + weekdayShift) % 7]);
            }
        }

        var weekdayMatch = updatedText.match(/([（([\[]?)(日|月|火|水|木|金|土)([）)\]]?)/);
        if (weekdayMatch) {
            var weekdayIndex = JAPANESE_WEEKDAYS.indexOf(weekdayMatch[2]);
            if (weekdayIndex !== -1) {
                updatedText = updatedText.replace(weekdayMatch[0], (weekdayMatch[1] || "") + JAPANESE_WEEKDAYS[(weekdayIndex + weekdayShift) % 7] + (weekdayMatch[3] || ""));
            }
        }
        return updatedText;
    }

    /**
     * 1900〜2099 の4桁の年を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftYearNumber(lineText, stepInt) {
        var yearMatch = lineText.match(/\b(19|20)\d{2}\b/);
        if (!yearMatch) return null;
        return lineText.replace(yearMatch[0], String(parseInt(yearMatch[0], 10) + stepInt));
    }

    /**
     * 「29.3.2」「12.1」のようなドット区切りの数字を増減する
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftDotNumbers(lineText, stepInt) {
        var dotTripleMatch = lineText.match(/^(\d+)\.(\d+)\.(\d+)$/);
        if (dotTripleMatch) {
            /* 3要素のドット区切りは数値として扱う / Three-part dot patterns stay numeric */
            var firstPart = parseInt(dotTripleMatch[1], 10);
            var secondPart = parseInt(dotTripleMatch[2], 10);
            var thirdPart = parseInt(dotTripleMatch[3], 10);
            if (shiftTarget === "year") firstPart += stepInt;
            else if (shiftTarget === "month") secondPart += stepInt;
            else thirdPart += stepInt;
            return firstPart + "." + secondPart + "." + thirdPart;
        }

        var dotPairMatch = lineText.match(/^(\d+)\.(\d+)$/);
        if (!dotPairMatch) return null;
        return (dotPairMode === "date") ? shiftDotPairAsDate(dotPairMatch, stepInt) : shiftDotPairAsNumber(dotPairMatch, stepInt);
    }

    /**
     * 2要素のドット区切りを「月.日」とみなして増減する（年は今年で仮置き）
     * ［対象］が前半なら月、それ以外なら日を増減し、元の桁数でゼロ埋めする（例：12.05 → 12.06）
     * @param {string[]} dotPairMatch - /^(\d+)\.(\d+)$/ の一致結果
     * @param {number} stepInt - 増減する段数
     * @returns {string} 置き換えたテキスト
     */
    function shiftDotPairAsDate(dotPairMatch, stepInt) {
        var monthWidth = dotPairMatch[1].length;
        var dayWidth = dotPairMatch[2].length;
        var shiftedDate = makeDate(CURRENT_YEAR, dotPairMatch[1], dotPairMatch[2]);

        if (shiftTarget === "year") {
            /* 「年」ターゲット（前半）は月を増減 / the first part shifts the month */
            shiftedDate.setMonth(shiftedDate.getMonth() + stepInt);
        } else {
            /* それ以外は日を増減 / other targets shift the day */
            shiftedDate.setDate(shiftedDate.getDate() + stepInt);
        }
        /* padZero は桁数が足りているときはそのまま返す / padZero leaves wide-enough values as they are */
        return padZero(shiftedDate.getMonth() + 1, monthWidth) + "." + padZero(shiftedDate.getDate(), dayWidth);
    }

    /**
     * 2要素のドット区切りを数字として増減する
     * ［対象］が前半なら前半だけ、後半なら後半を増減して桁あふれを前半へ繰り上げ／繰り下げる
     * @param {string[]} dotPairMatch - /^(\d+)\.(\d+)$/ の一致結果
     * @param {number} stepInt - 増減する段数
     * @returns {string} 置き換えたテキスト
     */
    function shiftDotPairAsNumber(dotPairMatch, stepInt) {
        var firstPart = parseInt(dotPairMatch[1], 10);
        var secondPart = parseInt(dotPairMatch[2], 10);

        if (shiftTarget === "year") {
            firstPart += stepInt;
        } else {
            /* 後半の桁数から基数を決め、合計を分解し直す / base from the digit count, then split the total back */
            var secondPartBase = Math.pow(10, String(Math.abs(secondPart)).length || 1);
            var combinedValue = firstPart * secondPartBase + secondPart + stepInt;
            var newFirstPart = (combinedValue >= 0) ? Math.floor(combinedValue / secondPartBase) : -Math.ceil(-combinedValue / secondPartBase);
            secondPart = combinedValue - newFirstPart * secondPartBase;
            firstPart = newFirstPart;
        }
        return firstPart + "." + secondPart;
    }

    /**
     * 「19:00」のような時刻を増減する（［対象］が前半なら時、それ以外は分。時は24時間でループ）
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepInt - 増減する段数
     * @returns {string|null} 置き換えたテキスト。該当しなければ null
     */
    function shiftTime(lineText, stepInt) {
        var timeMatch = lineText.match(/(\d{1,2}):(\d{2})/);
        if (!timeMatch) return null;
        var hour = parseInt(timeMatch[1], 10);
        var minute = parseInt(timeMatch[2], 10);

        if (shiftTarget === "year") {
            hour += stepInt;
        } else {
            minute += stepInt;
        }

        /* 分のあふれを時へ繰り上げ／繰り下げる / carry minute overflow into the hour */
        while (minute < 0) {
            minute += 60;
            hour -= 1;
        }
        while (minute >= 60) {
            minute -= 60;
            hour += 1;
        }
        hour = ((hour % 24) + 24) % 24;

        return lineText.replace(timeMatch[0], hour + ":" + padZero(minute, 2));
    }

    /**
     * 最初に現れる数値（カンマ区切り・小数を含む）を増減する。小数は最下位の桁を1単位とする
     * @param {string} lineText - 対象のテキスト
     * @param {number} stepAmount - 増減量
     * @returns {string} 置き換えたテキスト（数値が無ければそのまま）
     */
    function shiftFirstNumber(lineText, stepAmount) {
        var numberMatch = lineText.match(/-?\d+(?:,\d{3})*(?:\.\d+)?/);
        if (!numberMatch) return lineText;
        var numberText = numberMatch[0];
        var rawNumberText = numberText.replace(/,/g, "");
        var decimalMatch = rawNumberText.match(/\.(\d+)/);
        var decimalDigits = decimalMatch ? decimalMatch[1].length : 0;
        var unitStep = decimalDigits > 0 ? 1 / Math.pow(10, decimalDigits) : 1;
        var shiftedValue = parseFloat(rawNumberText) + unitStep * stepAmount;
        return lineText.replace(numberText, addThousandsCommas(shiftedValue.toFixed(decimalDigits)));
    }

    /**
     * 1行分のテキストに含まれる日付・曜日・時刻・数値を、現在の設定で増減する
     * @param {string} lineText - 対象の1行
     * @returns {string} 増減後のテキスト（失敗したら元のまま）
     */
    function incrementLine(lineText) {
        /* 途中で例外が出たら元の行を返す（ExtendScript の配列には indexOf が無いことがある）
           Return the line untouched on any error (ExtendScript arrays may lack indexOf) */
        try {
            var updatedText = trimText(lineText);
            var stepInt = toWholeStep(stepValue);

            for (var i = 0; i < DATE_SHIFTERS.length; i++) {
                var shiftedDateText = DATE_SHIFTERS[i](updatedText, stepInt);
                if (shiftedDateText !== null) return shiftedDateText;
            }

            updatedText = shiftWeekdayMarks(updatedText, stepInt);

            var shiftedText = shiftYearNumber(updatedText, stepInt);
            if (shiftedText === null) shiftedText = shiftDotNumbers(updatedText, stepInt);
            if (shiftedText === null) shiftedText = shiftTime(updatedText, stepInt);
            if (shiftedText !== null) return shiftedText;

            return shiftFirstNumber(updatedText, stepValue);
        } catch (e) {
            $.writeln("Error in incrementLine: " + e.message);
            return lineText;
        }
    }

    // =========================================
    // 適用 / Apply
    // =========================================

    /**
     * テキストフレームの各行を増減する（グループ内も再帰）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {void}
     */
    function incrementTextInItem(pageItem) {
        if (!pageItem) return;
        if (pageItem.typename === "TextFrame") {
            var lines = pageItem.contents.split(LINE_BREAK);
            for (var i = 0; i < lines.length; i++) {
                lines[i] = incrementLine(lines[i]);
            }
            pageItem.contents = lines.join(LINE_BREAK);
            return;
        }
        if (pageItem.typename === "GroupItem") {
            var childItems = pageItem.pageItems;
            for (var j = 0; j < childItems.length; j++) {
                incrementTextInItem(childItems[j]);
            }
        }
    }

    /**
     * 選択中のオブジェクトに増減を適用する
     * @returns {void}
     */
    function incrementSelectedTexts() {
        if (app.documents.length === 0) return;
        var selectedItems = app.activeDocument.selection;
        /* ロックされたテキストなどは contents の書き換えで例外になる / writing contents throws on locked text */
        try {
            for (var i = 0; i < selectedItems.length; i++) {
                incrementTextInItem(selectedItems[i]);
            }
        } catch (e) {
            alert(getLabel("alert.errorGeneric") + LINE_BREAK + e.message, getLabel("alert.errorTitle"));
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
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

    /**
     * 結果欄の右に出す数値を求める（純粋な数値どうしなら重複するので出さない）
     * @param {string} originalSample - オリジナルの1行
     * @param {string} previewText - 増減後の1行
     * @returns {number|string} 表示する数値。出さないときは空文字
     */
    function computeResultNumber(originalSample, previewText) {
        var pureNumberPattern = /^\s*-?\d+(?:,\d{3})*(?:\.\d+)?\s*$/;
        if (pureNumberPattern.test(originalSample) && pureNumberPattern.test(previewText)) return "";

        var numberPattern = /-?\d+(?:,\d{3})*(?:\.\d+)?/;
        var originalMatch = originalSample.match(numberPattern);
        var resultMatch = previewText.match(numberPattern);
        if (!originalMatch || !resultMatch) return "";

        var originalNumber = Number(originalMatch[0].replace(/,/g, ""));
        var resultNumber = Number(resultMatch[0].replace(/,/g, ""));
        return (!isNaN(originalNumber) && !isNaN(resultNumber)) ? resultNumber : "";
    }

    /**
     * 「項目名：値」の1行を追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelPath - 項目名のラベルのパス
     * @returns {{rowGroup: Group, fieldLabel: StaticText}} 行と項目名
     */
    function addPreviewRow(parentGroup, labelPath) {
        var rowGroup = parentGroup.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = PREVIEW_LABEL_SPACING;
        var fieldLabel = rowGroup.add("statictext", undefined, labelText(labelPath));
        fieldLabel.justify = "right";
        return { rowGroup: rowGroup, fieldLabel: fieldLabel };
    }

    /**
     * オリジナル／結果の2行を追加する
     * @param {Window} incrementDialog - 追加先のダイアログ
     * @param {string} originalSample - オリジナルの1行
     * @returns {{resultValueText: StaticText, resultNumberText: StaticText}} 更新する表示欄
     */
    function addPreviewRows(incrementDialog, originalSample) {
        var previewGroup = incrementDialog.add("group");
        previewGroup.orientation = "column";
        previewGroup.alignChildren = ["left", "top"];
        previewGroup.margins = PREVIEW_MARGINS;
        previewGroup.spacing = PREVIEW_SPACING;

        var originalRow = addPreviewRow(previewGroup, "fieldLabel.original");
        originalRow.rowGroup.add("statictext", undefined, originalSample);

        var resultRow = addPreviewRow(previewGroup, "fieldLabel.result");
        var resultValueText = resultRow.rowGroup.add("statictext", undefined, "");
        resultValueText.helpTip = getLabel("tooltip.result");
        var resultNumberText = resultRow.rowGroup.add("statictext", undefined, "");

        /* 項目名の幅をそろえ、結果欄の幅を確保する / align the labels and reserve the result width */
        var dialogGraphics = incrementDialog.graphics;
        var labelWidth = Math.ceil(Math.max(
            dialogGraphics.measureString(originalRow.fieldLabel.text)[0],
            dialogGraphics.measureString(resultRow.fieldLabel.text)[0]
        ));
        originalRow.fieldLabel.preferredSize = [labelWidth, -1];
        resultRow.fieldLabel.preferredSize = [labelWidth, -1];

        /* 「12.1」など短いオリジナルは、日付相当の見本で幅を確保（結果の切れを防ぐ）
           Short originals fall back to a date-length sample to avoid truncating the result */
        var widthSample = (originalSample.length < RESULT_SAMPLE_MIN_CHARS) ? RESULT_WIDTH_SAMPLE : originalSample;
        resultValueText.preferredSize = [Math.ceil(dialogGraphics.measureString(widthSample)[0]), -1];

        return { resultValueText: resultValueText, resultNumberText: resultNumberText };
    }

    /**
     * ［増減］パネル（増減値の数値欄）を追加する
     * @param {Window} incrementDialog - 追加先のダイアログ
     * @param {Function} onStepChange - 値が変わったときに呼ぶ関数
     * @returns {EditText} 増減値の入力欄
     */
    function addStepPanel(incrementDialog, onStepChange) {
        var stepPanel = incrementDialog.add("panel", undefined, getLabel("panel.step"));
        setupPanel(stepPanel);

        /* 整数のみ・負の値も可（下限なし） / integers only, negatives allowed (no minimum) */
        var stepInput = addSteppedField(stepPanel, {
            label: labelText("fieldLabel.step"),
            text: String(DEFAULT_STEP_VALUE),
            characters: STEP_FIELD_CHARACTERS,
            step: 1,
            integer: true,
            onStep: onStepChange
        });
        stepInput.helpTip = getLabel("tooltip.step");

        /* 入力中もプレビューを追従させる / keep the preview in step while typing */
        stepInput.onChanging = onStepChange;
        /* 部品の onChange（整数化・数値以外は元に戻す）のあとで反映する / apply after the part's normalization */
        var normalizeStepInput = stepInput.onChange;
        stepInput.onChange = function () {
            normalizeStepInput();
            onStepChange();
        };
        return stepInput;
    }

    /**
     * ［種別］パネル（数字／日付）を追加する。「12.1」のような2要素のドット区切りのときだけ使う
     * @param {Window} incrementDialog - 追加先のダイアログ
     * @param {Function} onModeChange - 切り替えたときに呼ぶ関数
     * @returns {void}
     */
    function addModePanel(incrementDialog, onModeChange) {
        var modePanel = incrementDialog.add("panel", undefined, getLabel("panel.mode"));
        setupPanel(modePanel);
        modePanel.orientation = "row"; /* 2つのラジオボタンを横に並べる / two radio buttons side by side */
        modePanel.alignChildren = ["left", "center"];

        var numberModeRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeNumber"));
        numberModeRadio.helpTip = getLabel("tooltip.modeNumber");
        var dateModeRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeDate"));
        dateModeRadio.helpTip = getLabel("tooltip.modeDate");

        /* 既定は「数字」 / default is Number */
        dotPairMode = "number";
        numberModeRadio.value = true;

        numberModeRadio.onClick = function () {
            dotPairMode = "number";
            onModeChange();
        };
        dateModeRadio.onClick = function () {
            dotPairMode = "date";
            onModeChange();
        };
    }

    /**
     * ［対象］パネルを追加する（分割できた部分ごとのラジオボタン。既定は最後の部分）
     * @param {Window} incrementDialog - 追加先のダイアログ
     * @param {{yearPart: string, monthPart: string, dayPart: string}} selectionSample - readSelectionSample() の結果
     * @param {Function} onTargetChange - 切り替えたときに呼ぶ関数
     * @returns {{yearRadio: (RadioButton|undefined), monthRadio: (RadioButton|undefined), dayRadio: (RadioButton|undefined)}} 作ったラジオボタン
     */
    function addTargetPanel(incrementDialog, selectionSample, onTargetChange) {
        var targetPanel = incrementDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, TARGET_RADIO_SPACING);

        var targetRadios = {};
        var partEntries = [
            { key: "yearRadio", text: selectionSample.yearPart },
            { key: "monthRadio", text: selectionSample.monthPart },
            { key: "dayRadio", text: selectionSample.dayPart }
        ];
        var lastRadio = null;
        for (var i = 0; i < partEntries.length; i++) {
            if (partEntries[i].text === "") continue;
            lastRadio = targetPanel.add("radiobutton", undefined, partEntries[i].text);
            lastRadio.helpTip = getLabel("tooltip.target");
            lastRadio.onClick = onTargetChange;
            targetRadios[partEntries[i].key] = lastRadio;
        }
        /* 既定は最後の部分（日、なければ月・年） / default is the last part */
        if (lastRadio) lastRadio.value = true;
        return targetRadios;
    }

    /**
     * ［対象］のラジオボタンから増減する部分を読む
     * @param {Object|null} targetRadios - addTargetPanel() の結果（パネルが無ければ null）
     * @returns {string} "year" / "month" / "day"
     */
    function readShiftTarget(targetRadios) {
        if (!targetRadios) return "day";
        if (targetRadios.yearRadio && targetRadios.yearRadio.value) return "year";
        if (targetRadios.monthRadio && targetRadios.monthRadio.value) return "month";
        return "day";
    }

    /**
     * 増減値の入力欄を整数で読む
     * @param {EditText} stepInput - 増減値の入力欄
     * @returns {number|null} 増減値。数値でなければ null
     */
    function readStepInput(stepInput) {
        var parsedStep = parseFloat(stepInput.text);
        return isNaN(parsedStep) ? null : Math.round(parsedStep);
    }

    /**
     * ダイアログを表示し、OK で選択中のテキストに増減を適用する
     * @returns {void}
     */
    function showIncrementDialog() {
        var selectionSample = readSelectionSample();
        var originalSample = selectionSample.originalSample;

        var incrementDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(incrementDialog);

        var previewFields = addPreviewRows(incrementDialog, originalSample);
        var stepInput;
        var targetRadios = null;

        /**
         * 現在の設定でオリジナルの1行を変換し、結果欄に出す
         * @returns {void}
         */
        function updateResultPreview() {
            if (!originalSample) {
                previewFields.resultValueText.text = "";
                previewFields.resultNumberText.text = "";
                return;
            }
            var previewText = incrementLine(originalSample);
            previewFields.resultValueText.text = previewText;
            /* 年月日などを分割できるときは右側の数値を出さない / no extra number for split targets */
            previewFields.resultNumberText.text = selectionSample.hasDateParts ? "" : computeResultNumber(originalSample, previewText);
        }

        /**
         * 入力欄の増減値を反映してプレビューを更新する（数値でなければ何もしない）
         * @returns {void}
         */
        function handleStepChange() {
            var parsedStep = readStepInput(stepInput);
            if (parsedStep === null) return;
            stepValue = parsedStep;
            updateResultPreview();
        }

        /**
         * ［対象］の選択を反映し、必要なら増減値を初期値に戻して入力欄を編集状態にする
         * @param {boolean} resetStep - 増減値を初期値に戻すなら true
         * @returns {void}
         */
        function applyTargetFromRadios(resetStep) {
            shiftTarget = readShiftTarget(targetRadios);
            if (resetStep) {
                stepValue = DEFAULT_STEP_VALUE;
                stepInput.text = String(DEFAULT_STEP_VALUE);
                stepInput.active = true;
            }
            updateResultPreview();
        }

        stepInput = addStepPanel(incrementDialog, handleStepChange);

        if (selectionSample.isAmbiguousDotPair) {
            addModePanel(incrementDialog, updateResultPreview);
        }

        if (selectionSample.hasDateParts) {
            targetRadios = addTargetPanel(incrementDialog, selectionSample, function () {
                applyTargetFromRadios(true);
            });
            /* 既定の選択を反映（ここでは増減値を戻さない） / sync the default selection without resetting the step */
            applyTargetFromRadios(false);
        }

        var buttonRow = addButtonRow(incrementDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnOK.onClick = function () {
            var parsedStep = readStepInput(stepInput);
            stepValue = (parsedStep === null) ? DEFAULT_STEP_VALUE : parsedStep;
            stepInput.text = String(stepValue);

            /* ラジオボタンの状態を最終的な増減対象に反映（増減値は維持） / sync the target, keeping the step */
            applyTargetFromRadios(false);

            incrementSelectedTexts();
            incrementDialog.close(1);
        };
        btnCancel.onClick = function () {
            incrementDialog.close(0);
        };
        alignRightOnlyButtonRow(buttonRow);

        incrementDialog.onShow = function () {
            stepInput.active = true;
            updateResultPreview();
        };

        incrementDialog.layout.layout(true);
        prepareDialogWindow(incrementDialog, SCRIPT_NAME);
        incrementDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    showIncrementDialog();

})();
