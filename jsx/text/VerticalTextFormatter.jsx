#target illustrator
#targetengine "VerticalTextFormatterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した縦組みのテキストに、欧文回転・縦中横・数字の書き換え（全角・漢数字・和暦）などをまとめて適用します。
ダイアログボックスで項目を選び、変換の結果を一覧で確かめてから実行できます。

### Overview

Applies vertical-text settings to the selected text in one go: rotated Roman characters, Tate-Chu-Yoko,
and number rewriting (full-width digits, kanji numerals, Japanese era years).
Choose the options in a dialog box and check the result list before running.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "VerticalTextFormatter";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-15";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* ダイアログの初期値（前回の値は SettingsStore に残る） / Dialog defaults (the last values are kept by SettingsStore) */
    var DEFAULT_SETTINGS = {
        rotateRoman: true,
        tcyNumbers: true,
        tcyMinDigits: 2,
        tcyMaxDigits: 3,
        tcyExcludeAdjacent: true,
        tcyCapitals: false,
        tcyExclamation: false,
        singleDigitFullwidth: false,
        japaneseEra: false,
        kanji: false,
        kanjiMinDigits: 4,
        kanjiMaxDigits: 20,
        kanjiStyle: "digit",
        halfwidthAlnum: false,
        commaToTouten: true,
        periodToNakaguro: true,
        verticalOnly: true
    };

    var PREVIEW_MAX_LINES = 300;   /* 確認の一覧に出す最大の行数 / max lines in the check list */

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

    var BUTTON_WIDTH        = 90;    /* OK／キャンセルの幅 / OK and Cancel width */
    var OPTION_INDENT       = 20;    /* 親のチェックボックスに従う項目の字下げ / indent for dependent options */
    var PREVIEW_LIST_HEIGHT = 140;   /* 確認の一覧の高さ / height of the check list */

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
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
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
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
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
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
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
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
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
            title: { ja: "縦組みの整形", en: "Vertical Text Formatter" }
        },
        panel: {
            rotateTcy: { ja: "欧文回転・縦中横", en: "Rotation and Tate-Chu-Yoko" },
            replace: { ja: "文字の変換", en: "Character Conversion" },
            numbers: { ja: "数字", en: "Numbers" },
            preview: { ja: "変換の確認", en: "Conversion Check" }
        },
        checkbox: {
            rotateRoman: { ja: "縦組み中の欧文回転", en: "Rotate Roman characters" },
            tcyNumbers: { ja: "数字を縦中横", en: "Tate-Chu-Yoko for numbers" },
            tcyExcludeAdjacent: { ja: "桁区切り・小数点の隣は除く", en: "Skip numbers next to separators" },
            tcyCapitals: { ja: "大文字の英字（2〜3字）を縦中横", en: "Tate-Chu-Yoko for 2–3 capitals" },
            tcyExclamation: { ja: "「!!」「!?」を縦中横", en: "Tate-Chu-Yoko for \"!!\" and \"!?\"" },
            halfwidthAlnum: { ja: "全角英数字を半角に", en: "Full-width letters and digits to half-width" },
            commaToTouten: { ja: "桁区切りの「,」「，」を「、」に", en: "Thousands \",\" to \"、\"" },
            periodToNakaguro: { ja: "小数点の「.」「．」を「・」に", en: "Decimal \".\" to \"・\"" },
            singleDigitFullwidth: { ja: "1桁の数字を全角に", en: "Single digits to full-width" },
            japaneseEra: { ja: "西暦を和暦に（2026年 → 令和8年）", en: "Western years to Japanese era" },
            kanji: { ja: "漢数字に変換", en: "Convert to kanji numerals" },
            verticalOnly: { ja: "縦組みのテキストだけを処理", en: "Vertical text only" }
        },
        radio: {
            kanjiDigit: { ja: "1字ずつ（二〇二六）", en: "Digit by digit (二〇二六)" },
            kanjiPositional: { ja: "十・百・千を入れる（二千二十六）", en: "With 十・百・千 (二千二十六)" }
        },
        field: {
            digits: { ja: "桁数", en: "Digits" },
            rangeTo: { ja: "〜", en: "to" }
        },
        preview: {
            tcyMark: { ja: "［縦中横］", en: " [TCY]" },
            none: { ja: "（変換する文字はありません）", en: "(Nothing to convert)" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noSelection: { ja: "テキストを選択してください。", en: "Select text first." },
            noVerticalText: { ja: "縦組みのテキストがありません。", en: "No vertical text is selected." },
            actionFailed: { ja: "縦組み中の欧文回転・縦中横を適用できませんでした。", en: "Could not apply the vertical text settings." }
        },
        /* ステップボタン用 / for the stepper */
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
    // 一時アクション定義 / Action definition
    // =========================================
    var ACTION_SET_NAME = "RotateRomanAndTateChuYoko";   /* アクションセット名 / action set name */
    var ROTATE_ACTION_NAME = "rotateRoman";              /* 欧文回転のアクション名 / rotate action name */
    var TCY_ACTION_NAME = "tateChuYoko";                 /* 縦中横のアクション名 / Tate-Chu-Yoko action name */

    /**
     * 文字パネルの真偽値の項目を1つだけオンにするアクション1件ぶんの行を返す
     * （文字パネルのイベントは、書いたパラメーターだけを適用する）
     * @param {number} actionIndex - アクションの番号（1始まり）
     * @param {string} actionName - アクション名
     * @param {number} paramKey - パラメーターのキー
     * @returns {string[]} アクション定義の行
     */
    function buildCharacterToggleActionLines(actionIndex, actionName, paramKey) {
        return ["/action-" + actionIndex + " {"].concat(
            buildActionNameLines("\t", actionName),
            [
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (adobe_SLOCharacterPalette)"
            ],
            buildActionNameLines("\t\t", "文字", "localizedName"),
            [
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key " + paramKey,
                "\t\t\t/showInPalette 4294967295",
                "\t\t\t/type (boolean)",
                "\t\t\t/value 1",
                "\t\t}",
                "\t}",
                "}"
            ]
        );
    }

    /**
     * ［縦組み中の欧文回転］（'rriv'）と［縦中横］（'tyen'）の2件を持つアクション定義を返す
     * @returns {string} アクション定義のテキスト
     */
    function buildVerticalTextActions() {
        return ["/version 3"].concat(
            buildActionNameLines("", ACTION_SET_NAME),
            ["/isOpen 1", "/actionCount 2"],
            buildCharacterToggleActionLines(1, ROTATE_ACTION_NAME, 1920100726),
            buildCharacterToggleActionLines(2, TCY_ACTION_NAME, 1954112878),
            [""]
        ).join("\n");
    }

    // =========================================
    // 変換の計画 / Conversion plan
    // =========================================
    var KANJI_DIGITS = ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
    var KANJI_SMALL_UNITS = ["", "十", "百", "千"];
    var KANJI_LARGE_UNITS = ["", "万", "億", "兆", "京"];
    var FULLWIDTH_DIGIT_OFFSET = 0xFF10 - 0x30;   /* 半角数字→全角数字の差 / offset from ASCII digits to full-width digits */

    /**
     * 半角数字か
     * @param {string} ch - 1文字
     * @returns {boolean} 0〜9 なら true
     */
    function isDigitChar(ch) {
        return ch >= "0" && ch <= "9";
    }

    /**
     * 半角英字か
     * @param {string} ch - 1文字
     * @returns {boolean} A〜Z・a〜z なら true
     */
    function isLatinLetter(ch) {
        return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
    }

    /**
     * 感嘆符・疑問符か（半角・全角）
     * @param {string} ch - 1文字
     * @returns {boolean} ! ? ！ ？ なら true
     */
    function isExclamOrQuestion(ch) {
        return ch === "!" || ch === "?" || ch === "！" || ch === "？";
    }

    /**
     * 全角の英数字を半角にする（1文字ずつ置き換えるので文字数は変わらない）
     * @param {string} text - 対象の文字列
     * @returns {string} 置き換えた文字列
     */
    function toHalfwidthAlnum(text) {
        var resultText = "";
        for (var i = 0; i < text.length; i++) {
            var charCode = text.charCodeAt(i);
            var isFullAlnum = (charCode >= 0xFF10 && charCode <= 0xFF19) || (charCode >= 0xFF21 && charCode <= 0xFF3A) || (charCode >= 0xFF41 && charCode <= 0xFF5A);
            resultText += isFullAlnum ? String.fromCharCode(charCode - 0xFEE0) : text.charAt(i);
        }
        return resultText;
    }

    /**
     * 半角数字を全角数字にする
     * @param {string} digitText - 半角数字の文字列
     * @returns {string} 全角数字の文字列
     */
    function toFullwidthDigits(digitText) {
        var resultText = "";
        for (var i = 0; i < digitText.length; i++) {
            resultText += String.fromCharCode(digitText.charCodeAt(i) + FULLWIDTH_DIGIT_OFFSET);
        }
        return resultText;
    }

    /**
     * 数字を1字ずつ漢数字にする（2026 → 二〇二六）
     * @param {string} digitText - 半角数字の文字列
     * @returns {string} 漢数字の文字列
     */
    function toKanjiDigitByDigit(digitText) {
        var resultText = "";
        for (var i = 0; i < digitText.length; i++) {
            resultText += KANJI_DIGITS[digitText.charCodeAt(i) - 48];
        }
        return resultText;
    }

    /**
     * 整数を十・百・千と万・億…を入れた漢数字にする（2026 → 二千二十六、10000 → 一万）
     * 十・百・千の前の「一」は省く。万の位を超える桁は1字ずつにする
     * @param {string} digitText - 半角数字の文字列
     * @returns {string} 漢数字の文字列
     */
    function toKanjiPositional(digitText) {
        var trimmedText = digitText.replace(/^0+/, "");
        if (trimmedText === "") return KANJI_DIGITS[0];
        if (trimmedText.length > KANJI_LARGE_UNITS.length * 4) return toKanjiDigitByDigit(digitText);
        var resultText = "";
        var groupCount = Math.ceil(trimmedText.length / 4);
        for (var g = groupCount - 1; g >= 0; g--) {
            var groupEnd = trimmedText.length - g * 4;
            var groupText = trimmedText.substring(Math.max(0, groupEnd - 4), groupEnd);
            var groupKanji = "";
            for (var i = 0; i < groupText.length; i++) {
                var digitValue = groupText.charCodeAt(i) - 48;
                if (digitValue === 0) continue;
                var place = groupText.length - 1 - i;
                /* 十・百・千の前の「一」は書かない / Omit "one" before ten, hundred and thousand */
                if (!(digitValue === 1 && place > 0)) groupKanji += KANJI_DIGITS[digitValue];
                groupKanji += KANJI_SMALL_UNITS[place];
            }
            if (groupKanji !== "") resultText += groupKanji + KANJI_LARGE_UNITS[g];
        }
        return resultText;
    }

    /**
     * 文字列のその位置から数字のまとまり（桁区切り・小数を含む）を読む
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 数字で始まる位置
     * @returns {{length: number, parts: Array<{type: string, text: string}>, integerDigits: number, hasSeparator: boolean}}
     *   parts の type は "digits" / "separator"（, ，）/ "decimal"（. ．）
     */
    function readNumberToken(text, startIndex) {
        var parts = [];
        var i = startIndex;
        var integerDigits = 0;
        var hasSeparator = false;

        /**
         * その位置から続く数字を読み、parts に足す
         * @returns {number} 読んだ桁数
         */
        function readDigits() {
            var digitStart = i;
            while (i < text.length && isDigitChar(text.charAt(i))) i++;
            parts.push({ type: "digits", text: text.substring(digitStart, i) });
            return i - digitStart;
        }

        integerDigits += readDigits();
        /* 桁区切り：「,」「，」のあとにちょうど3桁 / Group separator: "," or "，" followed by exactly three digits */
        while ((text.charAt(i) === "," || text.charAt(i) === "，") && isDigitChar(text.charAt(i + 1)) && isDigitChar(text.charAt(i + 2)) &&
                isDigitChar(text.charAt(i + 3)) && !isDigitChar(text.charAt(i + 4))) {
            parts.push({ type: "separator", text: text.charAt(i) });
            i++;
            integerDigits += readDigits();
            hasSeparator = true;
        }
        /* 小数点：「.」「．」のあとに数字 / Decimal point: "." or "．" followed by a digit */
        if ((text.charAt(i) === "." || text.charAt(i) === "．") && isDigitChar(text.charAt(i + 1))) {
            parts.push({ type: "decimal", text: text.charAt(i) });
            i++;
            readDigits();
            hasSeparator = true;
        }
        return { length: i - startIndex, parts: parts, integerDigits: integerDigits, hasSeparator: hasSeparator };
    }

    /**
     * 桁区切り・小数点を設定に合わせて置き換える（置き換えないときは元の文字のまま）
     * @param {Object} part - readNumberToken の parts の要素（separator / decimal）
     * @param {Object} options - 変換の設定
     * @returns {string} 置き換えた文字
     */
    function convertNumberMark(part, options) {
        if (part.type === "separator") return options.commaToTouten ? "、" : part.text;
        return options.periodToNakaguro ? "・" : part.text;
    }

    /**
     * 数字のまとまりを漢数字にする（小数部は書き方によらず1字ずつ）
     * @param {Object} numberToken - readNumberToken の戻り値
     * @param {Object} options - 変換の設定
     * @returns {string} 漢数字の文字列
     */
    function convertNumberToKanji(numberToken, options) {
        var parts = numberToken.parts;
        var resultText = "";
        if (options.kanjiStyle === "positional") {
            /* 整数部は桁区切りを外して位取りで読む / Read the integer part positionally, without separators */
            var integerText = "";
            var partIndex = 0;
            while (partIndex < parts.length && parts[partIndex].type !== "decimal") {
                if (parts[partIndex].type === "digits") integerText += parts[partIndex].text;
                partIndex++;
            }
            resultText = toKanjiPositional(integerText);
            if (partIndex < parts.length) {
                resultText += convertNumberMark(parts[partIndex], options) + toKanjiDigitByDigit(parts[partIndex + 1].text);
            }
            return resultText;
        }
        for (var i = 0; i < parts.length; i++) {
            resultText += (parts[i].type === "digits") ? toKanjiDigitByDigit(parts[i].text) : convertNumberMark(parts[i], options);
        }
        return resultText;
    }

    /**
     * 数字だけの並び（1つの digits）を、縦中横・全角・そのままのどれかにして出力に足す
     * @param {string} digitText - 半角数字の文字列
     * @param {Object} options - 変換の設定
     * @param {function(string, boolean): void} emit - 出力に足す関数（文字列、縦中横にするか）
     * @returns {void}
     */
    function emitDigitRun(digitText, options, emit) {
        var digitCount = digitText.length;
        if (options.tcyNumbers && digitCount >= options.tcyMinDigits && digitCount <= options.tcyMaxDigits) {
            emit(digitText, true);
        } else if (options.singleDigitFullwidth && digitCount === 1) {
            emit(toFullwidthDigits(digitText), false);
        } else {
            emit(digitText, false);
        }
    }

    /**
     * 数字だけの並びを、範囲内なら漢数字、そうでなければ縦中横・全角・そのままにして出力に足す
     * @param {string} digitText - 半角数字の文字列
     * @param {Object} options - 変換の設定
     * @param {function(string, boolean): void} emit - 出力に足す関数
     * @returns {void}
     */
    function emitPlainDigits(digitText, options, emit) {
        if (options.kanji && digitText.length >= options.kanjiMinDigits && digitText.length <= options.kanjiMaxDigits) {
            emit(convertNumberToKanji({ parts: [{ type: "digits", text: digitText }] }, options), false);
        } else {
            emitDigitRun(digitText, options, emit);
        }
    }

    /* 元号と開始日（YYYYMMDD）。新しい順 / Japanese eras and their start dates (YYYYMMDD), newest first */
    var JAPANESE_ERAS = [
        { name: "令和", start: 20190501 },
        { name: "平成", start: 19890108 },
        { name: "昭和", start: 19261225 },
        { name: "大正", start: 19120730 },
        { name: "明治", start: 18681023 }
    ];

    /**
     * その位置の「西暦4桁＋年」を読み、元号と年を返す。月日が続けば改元の境目も月日で判定する
     * （月日が無い改元の年は新しい元号にする：2019年 → 令和元年）
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 数字で始まる位置
     * @returns {{eraName: string, eraYear: number}|null} 西暦の年でなければ null
     */
    function matchWesternYear(text, startIndex) {
        var yearMatch = /^(\d{4})年(?:(\d{1,2})月(?:(\d{1,2})日)?)?/.exec(text.substr(startIndex, 12));
        if (!yearMatch) return null;
        var year = parseInt(yearMatch[1], 10);
        var month = yearMatch[2] ? parseInt(yearMatch[2], 10) : 12;
        var day = yearMatch[3] ? parseInt(yearMatch[3], 10) : 31;
        var dateKey = year * 10000 + month * 100 + day;
        for (var i = 0; i < JAPANESE_ERAS.length; i++) {
            if (dateKey >= JAPANESE_ERAS[i].start) {
                return { eraName: JAPANESE_ERAS[i].name, eraYear: year - Math.floor(JAPANESE_ERAS[i].start / 10000) + 1 };
            }
        }
        return null;
    }

    /**
     * テキストの変換内容を決める。書き換え（後ろから順）と、書き換え後の縦中横の位置を返す
     * @param {string} sourceText - 元の文字列
     * @param {Object} options - 変換の設定
     * @returns {{edits: Array<{start: number, length: number, text: string}>, tcySpans: Array<{start: number, length: number}>,
     *     changes: Array<{before: string, after: string, tcy: boolean}>, outputLength: number}}
     *     edits は後ろから順、changes は確認の一覧用（前から順）
     */
    function planVerticalText(sourceText, options) {
        var text = options.halfwidthAlnum ? toHalfwidthAlnum(sourceText) : sourceText;
        var edits = [];
        var tcySpans = [];
        var changes = [];
        var outputLength = 0;
        var segmentText = "";
        var segmentHasTcy = false;

        /**
         * 出力に文字列を足す（縦中横なら位置を控える）
         * @param {string} pieceText - 足す文字列
         * @param {boolean} isTcy - 縦中横にするか
         * @returns {void}
         */
        function emit(pieceText, isTcy) {
            if (isTcy) {
                tcySpans.push({ start: outputLength, length: pieceText.length });
                segmentHasTcy = true;
            }
            outputLength += pieceText.length;
            segmentText += pieceText;
        }

        /**
         * 元の [start, start+length) を出力 segmentText に置き換えたことを控える（変わらなければ控えない）
         * @param {number} start - 元の位置
         * @param {number} length - 元の文字数
         * @returns {void}
         */
        function closeSegment(start, length) {
            var originalText = sourceText.substr(start, length);
            if (segmentText !== originalText) {
                edits.push({ start: start, length: length, text: segmentText });
            }
            if (segmentText !== originalText || segmentHasTcy) {
                changes.push({ before: originalText, after: segmentText, tcy: segmentHasTcy });
            }
            segmentText = "";
            segmentHasTcy = false;
        }

        var i = 0;
        while (i < text.length) {
            var ch = text.charAt(i);
            var runStart = i;
            var westernYear = (options.japaneseEra && isDigitChar(ch)) ? matchWesternYear(text, i) : null;
            if (westernYear) {
                emit(westernYear.eraName, false);
                /* 1年は「元年」 / The first year is written 元 */
                if (westernYear.eraYear === 1) emit("元", false);
                else emitPlainDigits(String(westernYear.eraYear), options, emit);
                i += 4;
            } else if (isDigitChar(ch)) {
                var numberToken = readNumberToken(text, i);
                if (options.kanji && numberToken.integerDigits >= options.kanjiMinDigits && numberToken.integerDigits <= options.kanjiMaxDigits) {
                    emit(convertNumberToKanji(numberToken, options), false);
                } else if (numberToken.hasSeparator && options.tcyExcludeAdjacent) {
                    /* 桁区切り・小数点の隣の数字は縦中横・全角にしない / Leave numbers next to separators as they are */
                    for (var p = 0; p < numberToken.parts.length; p++) {
                        var part = numberToken.parts[p];
                        emit(part.type === "digits" ? part.text : convertNumberMark(part, options), false);
                    }
                } else {
                    for (var q = 0; q < numberToken.parts.length; q++) {
                        var numberPart = numberToken.parts[q];
                        if (numberPart.type === "digits") emitDigitRun(numberPart.text, options, emit);
                        else emit(convertNumberMark(numberPart, options), false);
                    }
                }
                i += numberToken.length;
            } else if (isLatinLetter(ch)) {
                while (i < text.length && isLatinLetter(text.charAt(i))) i++;
                var letterText = text.substring(runStart, i);
                var isCapsTcy = options.tcyCapitals && letterText.length >= 2 && letterText.length <= 3 && letterText === letterText.toUpperCase();
                emit(letterText, isCapsTcy);
            } else if (isExclamOrQuestion(ch)) {
                while (i < text.length && isExclamOrQuestion(text.charAt(i))) i++;
                var markText = text.substring(runStart, i);
                if (options.tcyExclamation && markText.length === 2) {
                    /* 縦中横にする組は半角にそろえる / Use half-width marks for the Tate-Chu-Yoko pair */
                    emit(markText.replace(/！/g, "!").replace(/？/g, "?"), true);
                } else {
                    emit(markText, false);
                }
            } else {
                emit(ch, false);
                i++;
            }
            closeSegment(runStart, i - runStart);
        }
        edits.reverse();
        return { edits: edits, tcySpans: tcySpans, changes: changes, outputLength: outputLength };
    }

    // =========================================
    // 対象のテキスト / Text targets
    // =========================================

    /**
     * テキストフレームが縦組みか
     * @param {TextFrame} textFrame - 判定するフレーム
     * @returns {boolean} 縦組みなら true
     */
    function isVerticalFrame(textFrame) {
        return !!textFrame && textFrame.orientation === TextOrientation.VERTICAL;
    }

    /**
     * 選択から処理の対象を集める。文字選択中はその範囲、オブジェクト選択ならテキストフレームごと
     * @param {Document} doc - 対象のドキュメント
     * @returns {Array<Object>} { frame | range, offset（文字の番号のずれ）, text, isVertical }
     */
    function collectTextTargets(doc) {
        var selection = doc.selection;
        var textTargets = [];
        /* 文字選択中は selection が TextRange（[0] が無い） / While text is selected, selection is a TextRange without [0] */
        if (selection && selection.typename === "TextRange") {
            if (selection.length > 0) {
                textTargets.push({
                    range: selection,
                    offset: selection.characterOffset - 1,
                    text: selection.contents,
                    isVertical: isVerticalFrame(resolveTextRangeFrame(selection))
                });
            }
            return textTargets;
        }
        var textFrames = collectSelectionTextFrames(selection, { skipLocked: true, skipHidden: true });
        for (var i = 0; i < textFrames.length; i++) {
            textTargets.push({ frame: textFrames[i], offset: 0, text: textFrames[i].contents, isVertical: isVerticalFrame(textFrames[i]) });
        }
        return textTargets;
    }

    /**
     * 縦組みだけに絞る（指定がなければそのまま）
     * @param {Array<Object>} textTargets - collectTextTargets() の戻り値
     * @param {boolean} verticalOnly - 縦組みだけにするか
     * @returns {Array<Object>} 対象
     */
    function filterTextTargets(textTargets, verticalOnly) {
        if (!verticalOnly) return textTargets;
        var filteredTargets = [];
        for (var i = 0; i < textTargets.length; i++) {
            if (textTargets[i].isVertical) filteredTargets.push(textTargets[i]);
        }
        return filteredTargets;
    }

    /**
     * 対象の文字のコレクションを返す（文字選択はストーリー全体の番号で扱う）
     * @param {Object} textTarget - 対象
     * @returns {Characters} 文字のコレクション
     */
    function getTargetCharacters(textTarget) {
        return textTarget.frame ? textTarget.frame.characters : textTarget.range.story.characters;
    }

    /**
     * 書き換えを適用する。文字数が同じなら1文字ずつ、異なるなら余分を消して先頭の文字に入れる（書式は先頭の文字のもの）
     * @param {Object} textTarget - 対象
     * @param {Array<Object>} edits - planVerticalText() の edits（後ろから順）
     * @returns {void}
     */
    function applyTextEdits(textTarget, edits) {
        for (var i = 0; i < edits.length; i++) {
            var characters = getTargetCharacters(textTarget);
            var baseIndex = textTarget.offset + edits[i].start;
            var editText = edits[i].text;
            if (editText.length === edits[i].length) {
                for (var k = 0; k < editText.length; k++) {
                    if (characters[baseIndex + k].contents !== editText.charAt(k)) characters[baseIndex + k].contents = editText.charAt(k);
                }
            } else {
                for (var r = edits[i].length - 1; r >= 1; r--) characters[baseIndex + r].remove();
                characters[baseIndex].contents = editText;
            }
        }
    }

    /**
     * 対象の中の範囲だけを選択する
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} textTarget - 対象
     * @param {number} start - 範囲の先頭（対象の中の位置）
     * @param {number} length - 文字数
     * @returns {void}
     */
    function selectTargetSpan(doc, textTarget, start, length) {
        /* select() は選択に足されるので、先に外す / select() adds to the selection, so clear it first */
        doc.selection = null;
        var spanRange = getTargetCharacters(textTarget)[textTarget.offset + start];
        spanRange.length = length;
        spanRange.select();
    }

    /**
     * 文字の書き換え・欧文回転・縦中横を順に適用する
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {Object} options - 変換の設定
     * @returns {boolean} アクションまで実行できたら true
     */
    function runVerticalTextConversion(doc, textTargets, options) {
        var plans = [];
        var needsAction = options.rotateRoman;
        for (var i = 0; i < textTargets.length; i++) {
            plans.push(planVerticalText(textTargets[i].text, options));
            if (plans[i].tcySpans.length > 0) needsAction = true;
        }
        for (var e = 0; e < textTargets.length; e++) {
            applyTextEdits(textTargets[e], plans[e].edits);
        }
        if (!needsAction) return true;

        if (!loadTemporaryActionSet(buildVerticalTextActions(), ACTION_SET_NAME)) return false;
        try {
            if (options.rotateRoman) {
                if (textTargets[0].range) {
                    selectTargetSpan(doc, textTargets[0], 0, plans[0].outputLength);
                } else {
                    doc.selection = null;
                    for (var f = 0; f < textTargets.length; f++) textTargets[f].frame.selected = true;
                }
                app.doScript(ROTATE_ACTION_NAME, ACTION_SET_NAME, false);
            }
            for (var t = 0; t < textTargets.length; t++) {
                var tcySpans = plans[t].tcySpans;
                for (var s = 0; s < tcySpans.length; s++) {
                    selectTargetSpan(doc, textTargets[t], tcySpans[s].start, tcySpans[s].length);
                    app.doScript(TCY_ACTION_NAME, ACTION_SET_NAME, false);
                }
            }
            return true;
        } catch (e) {
            $.writeln("runVerticalTextConversion: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(ACTION_SET_NAME);
            /* 元の選択に戻す / Restore the original selection */
            try {
                if (textTargets[0].range) {
                    selectTargetSpan(doc, textTargets[0], 0, plans[0].outputLength);
                } else {
                    doc.selection = null;
                    for (var u = 0; u < textTargets.length; u++) textTargets[u].frame.selected = true;
                }
            } catch (restoreError) { /* 戻せなくても続ける / keep going */ }
        }
    }

    /**
     * 確認の一覧に出す行を作る（同じ変換はまとめて件数を付ける）
     * @param {Array<Object>} textTargets - 対象
     * @param {Object} options - 変換の設定
     * @returns {string[]} 一覧の行
     */
    function buildPreviewLines(textTargets, options) {
        var lineKeys = [];
        var lineCounts = {};
        for (var i = 0; i < textTargets.length; i++) {
            var changes = planVerticalText(textTargets[i].text, options).changes;
            for (var c = 0; c < changes.length; c++) {
                var change = changes[c];
                var lineText = (change.before === change.after) ? change.before : change.before + " → " + change.after;
                if (change.tcy) lineText += getLabel("preview.tcyMark");
                if (lineCounts.hasOwnProperty(lineText)) {
                    lineCounts[lineText]++;
                } else {
                    lineCounts[lineText] = 1;
                    lineKeys.push(lineText);
                }
            }
        }
        var previewLines = [];
        for (var k = 0; k < lineKeys.length && k < PREVIEW_MAX_LINES; k++) {
            previewLines.push(lineKeys[k] + (lineCounts[lineKeys[k]] > 1 ? "  ×" + lineCounts[lineKeys[k]] : ""));
        }
        return previewLines;
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * チェックボックスを追加する
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelPath - LABELS のパス
     * @param {boolean} value - 初期値
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parent, labelPath, value) {
        var optionCheckbox = parent.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.value = value;
        return optionCheckbox;
    }

    /**
     * 字下げしたグループ（親のチェックボックスに従う項目用）を追加する
     * @param {Group|Panel} parent - 追加先
     * @returns {Group} グループ
     */
    function addIndentedGroup(parent) {
        var indentedGroup = parent.add("group");
        indentedGroup.orientation = "column";
        indentedGroup.alignChildren = ["left", "top"];
        indentedGroup.margins = [OPTION_INDENT, 0, 0, 0];
        indentedGroup.spacing = 6;
        return indentedGroup;
    }

    /* 数値欄の∧∨・↑↓で呼ぶ処理（bindDialogEvents で入れる） / Handler for the steppers and arrow keys (set in bindDialogEvents) */
    var digitStepHandler = null;

    /**
     * 数値欄の∧∨・↑↓で増減したことを知らせる
     * @returns {void}
     */
    function notifyDigitStep() {
        if (digitStepHandler) digitStepHandler();
    }

    /**
     * 「桁数 : [最小] 〜 [最大]」の行を追加する
     * @param {Group|Panel} parent - 追加先
     * @returns {{minInput: EditText, maxInput: EditText}} 入力欄
     */
    function addDigitRangeRow(parent) {
        var rangeRow = parent.add("group");
        setupRow(rangeRow, ["left", "center"], 4);
        var minInput = addSteppedField(rangeRow, { label: labelText("field.digits"), text: "", characters: 3, step: 1, min: 1, max: 99, integer: true, onStep: notifyDigitStep });
        var maxInput = addSteppedField(rangeRow, { label: getLabel("field.rangeTo"), text: "", characters: 3, step: 1, min: 1, max: 99, integer: true, onStep: notifyDigitStep });
        return { minInput: minInput, maxInput: maxInput };
    }

    /**
     * ダイアログを組み立てる
     * @param {Object} settings - 初期値（DEFAULT_SETTINGS の形）
     * @returns {Object} dialog と各コントロール
     */
    function buildDialog(settings) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);
        var controls = { dialog: dialog };

        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = WINDOW_SPACING;
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.spacing = WINDOW_SPACING;

        /* 欧文回転・縦中横 / Rotation and Tate-Chu-Yoko */
        var tcyPanel = leftColumn.add("panel", undefined, getLabel("panel.rotateTcy"));
        setupPanel(tcyPanel, 6);
        controls.rotateRoman = addOptionCheckbox(tcyPanel, "checkbox.rotateRoman", settings.rotateRoman);
        controls.tcyNumbers = addOptionCheckbox(tcyPanel, "checkbox.tcyNumbers", settings.tcyNumbers);
        var tcyNumberGroup = addIndentedGroup(tcyPanel);
        var tcyRange = addDigitRangeRow(tcyNumberGroup);
        controls.tcyMinInput = tcyRange.minInput;
        controls.tcyMaxInput = tcyRange.maxInput;
        controls.tcyExcludeAdjacent = addOptionCheckbox(tcyNumberGroup, "checkbox.tcyExcludeAdjacent", settings.tcyExcludeAdjacent);
        controls.tcyCapitals = addOptionCheckbox(tcyPanel, "checkbox.tcyCapitals", settings.tcyCapitals);
        controls.tcyExclamation = addOptionCheckbox(tcyPanel, "checkbox.tcyExclamation", settings.tcyExclamation);

        /* 文字の変換 / Character conversion */
        var replacePanel = leftColumn.add("panel", undefined, getLabel("panel.replace"));
        setupPanel(replacePanel, 6);
        controls.halfwidthAlnum = addOptionCheckbox(replacePanel, "checkbox.halfwidthAlnum", settings.halfwidthAlnum);
        controls.commaToTouten = addOptionCheckbox(replacePanel, "checkbox.commaToTouten", settings.commaToTouten);
        controls.periodToNakaguro = addOptionCheckbox(replacePanel, "checkbox.periodToNakaguro", settings.periodToNakaguro);

        /* 数字 / Numbers */
        var numberPanel = rightColumn.add("panel", undefined, getLabel("panel.numbers"));
        setupPanel(numberPanel, 6);
        controls.singleDigitFullwidth = addOptionCheckbox(numberPanel, "checkbox.singleDigitFullwidth", settings.singleDigitFullwidth);
        controls.japaneseEra = addOptionCheckbox(numberPanel, "checkbox.japaneseEra", settings.japaneseEra);
        controls.kanji = addOptionCheckbox(numberPanel, "checkbox.kanji", settings.kanji);
        var kanjiGroup = addIndentedGroup(numberPanel);
        var kanjiRange = addDigitRangeRow(kanjiGroup);
        controls.kanjiMinInput = kanjiRange.minInput;
        controls.kanjiMaxInput = kanjiRange.maxInput;
        controls.kanjiDigitRadio = kanjiGroup.add("radiobutton", undefined, getLabel("radio.kanjiDigit"));
        controls.kanjiPositionalRadio = kanjiGroup.add("radiobutton", undefined, getLabel("radio.kanjiPositional"));
        controls.kanjiPositionalRadio.value = (settings.kanjiStyle === "positional");
        controls.kanjiDigitRadio.value = !controls.kanjiPositionalRadio.value;

        controls.verticalOnly = addOptionCheckbox(rightColumn, "checkbox.verticalOnly", settings.verticalOnly);

        controls.tcyMinInput.text = String(settings.tcyMinDigits);
        controls.tcyMaxInput.text = String(settings.tcyMaxDigits);
        controls.kanjiMinInput.text = String(settings.kanjiMinDigits);
        controls.kanjiMaxInput.text = String(settings.kanjiMaxDigits);

        /* 変換の確認 / Conversion check */
        var previewPanel = dialog.add("panel", undefined, getLabel("panel.preview"));
        setupPanel(previewPanel);
        controls.previewList = previewPanel.add("listbox", undefined, []);
        controls.previewList.preferredSize.height = PREVIEW_LIST_HEIGHT;

        var buttonRow = addButtonRow(dialog);
        controls.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        controls.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        controls.btnCancel.preferredSize.width = BUTTON_WIDTH;
        controls.btnOK.preferredSize.width = BUTTON_WIDTH;
        return controls;
    }

    /**
     * 数値欄の整数を読む（読めなければ既定値）
     * @param {EditText} numberInput - 入力欄
     * @param {number} fallbackValue - 読めないときの値
     * @returns {number} 整数
     */
    function readDigitInput(numberInput, fallbackValue) {
        var value = parseInt(numberInput.text, 10);
        return isNaN(value) ? fallbackValue : value;
    }

    /**
     * ダイアログの値を設定のオブジェクトにする（最小と最大が逆なら入れ替える）
     * @param {Object} controls - buildDialog() の戻り値
     * @returns {Object} DEFAULT_SETTINGS と同じ形
     */
    function readDialogOptions(controls) {
        var tcyMin = readDigitInput(controls.tcyMinInput, DEFAULT_SETTINGS.tcyMinDigits);
        var tcyMax = readDigitInput(controls.tcyMaxInput, DEFAULT_SETTINGS.tcyMaxDigits);
        var kanjiMin = readDigitInput(controls.kanjiMinInput, DEFAULT_SETTINGS.kanjiMinDigits);
        var kanjiMax = readDigitInput(controls.kanjiMaxInput, DEFAULT_SETTINGS.kanjiMaxDigits);
        return {
            rotateRoman: controls.rotateRoman.value,
            tcyNumbers: controls.tcyNumbers.value,
            tcyMinDigits: Math.min(tcyMin, tcyMax),
            tcyMaxDigits: Math.max(tcyMin, tcyMax),
            tcyExcludeAdjacent: controls.tcyExcludeAdjacent.value,
            tcyCapitals: controls.tcyCapitals.value,
            tcyExclamation: controls.tcyExclamation.value,
            singleDigitFullwidth: controls.singleDigitFullwidth.value,
            japaneseEra: controls.japaneseEra.value,
            kanji: controls.kanji.value,
            kanjiMinDigits: Math.min(kanjiMin, kanjiMax),
            kanjiMaxDigits: Math.max(kanjiMin, kanjiMax),
            kanjiStyle: controls.kanjiPositionalRadio.value ? "positional" : "digit",
            halfwidthAlnum: controls.halfwidthAlnum.value,
            commaToTouten: controls.commaToTouten.value,
            periodToNakaguro: controls.periodToNakaguro.value,
            verticalOnly: controls.verticalOnly.value
        };
    }

    /**
     * ダイアログのイベントをつなぐ（有効／無効の切り替えと確認の一覧の更新）
     * @param {Object} controls - buildDialog() の戻り値
     * @param {Array<Object>} textTargets - collectTextTargets() の戻り値
     * @returns {void}
     */
    function bindDialogEvents(controls, textTargets) {
        /**
         * 親のチェックボックスに合わせて子の項目を有効／無効にする
         * @returns {void}
         */
        function updateEnabledStates() {
            setSteppedFieldEnabled(controls.tcyMinInput, controls.tcyNumbers.value);
            setSteppedFieldEnabled(controls.tcyMaxInput, controls.tcyNumbers.value);
            controls.tcyExcludeAdjacent.enabled = controls.tcyNumbers.value;
            setSteppedFieldEnabled(controls.kanjiMinInput, controls.kanji.value);
            setSteppedFieldEnabled(controls.kanjiMaxInput, controls.kanji.value);
            controls.kanjiDigitRadio.enabled = controls.kanji.value;
            controls.kanjiPositionalRadio.enabled = controls.kanji.value;
        }

        /**
         * 確認の一覧を今の設定で作り直す
         * @returns {void}
         */
        function updatePreview() {
            var options = readDialogOptions(controls);
            var previewLines = buildPreviewLines(filterTextTargets(textTargets, options.verticalOnly), options);
            controls.previewList.removeAll();
            if (previewLines.length === 0) previewLines.push(getLabel("preview.none"));
            for (var i = 0; i < previewLines.length; i++) controls.previewList.add("item", previewLines[i]);
        }

        var checkboxNames = ["rotateRoman", "tcyNumbers", "tcyExcludeAdjacent", "tcyCapitals", "tcyExclamation", "singleDigitFullwidth",
            "japaneseEra", "kanji", "halfwidthAlnum", "commaToTouten", "periodToNakaguro", "verticalOnly", "kanjiDigitRadio", "kanjiPositionalRadio"];
        for (var i = 0; i < checkboxNames.length; i++) {
            controls[checkboxNames[i]].onClick = function () {
                updateEnabledStates();
                updatePreview();
            };
        }

        digitStepHandler = updatePreview;
        var digitInputs = [controls.tcyMinInput, controls.tcyMaxInput, controls.kanjiMinInput, controls.kanjiMaxInput];
        for (var d = 0; d < digitInputs.length; d++) {
            wrapDigitInputUpdate(digitInputs[d], updatePreview);
        }

        updateEnabledStates();
        updatePreview();
    }

    /**
     * 数値欄を直接入力したときも確認の一覧を更新する
     * （部品の onChange は値をそろえるだけなので、包んで後ろに足す）
     * @param {EditText} numberInput - addSteppedField() の入力欄
     * @param {function(): void} onUpdate - 更新の処理
     * @returns {void}
     */
    function wrapDigitInputUpdate(numberInput, onUpdate) {
        var normalizeOnChange = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeOnChange();
            onUpdate();
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、ダイアログで設定を選んで実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        var doc = app.activeDocument;
        var textTargets = collectTextTargets(doc);
        if (textTargets.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
        var controls = buildDialog(settingsStore.load(DEFAULT_SETTINGS));
        bindDialogEvents(controls, textTargets);
        prepareDialogWindow(controls.dialog, SCRIPT_NAME);
        if (controls.dialog.show() !== 1) return;

        var options = readDialogOptions(controls);
        settingsStore.save(options);
        var runTargets = filterTextTargets(textTargets, options.verticalOnly);
        if (runTargets.length === 0) {
            alert(getLabel("alert.noVerticalText"));
            return;
        }
        if (!runVerticalTextConversion(doc, runTargets, options)) {
            alert(getLabel("alert.actionFailed"));
        }
    }

    main();

})();
