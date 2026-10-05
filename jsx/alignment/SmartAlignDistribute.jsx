#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを縦または横に並べ、指定した間隔で分布します。
方向は自動判定でき、揃え（左右／上下）、プレビュー境界、ランダム並べ替えにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md

### Overview

Lines the selected objects up vertically or horizontally and distributes them at the spacing you specify.
The direction can be detected automatically, and alignment, preview bounds and random reordering are all supported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignDistribute";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビューを再描画する最小間隔（ミリ秒）/ Minimum interval between preview renders (ms) */
    var PREVIEW_MIN_INTERVAL_MS = 80;

    /* 操作からプレビューを実行するまでの待ち時間（ミリ秒）/ Delay before a requested preview runs (ms) */
    var PREVIEW_SCHEDULE_MS = 60;

    // =========================================
    // 内部キー / Internal keys
    // =========================================

    /* 計測用に一時的に作るグループの名前 / name of the throwaway measuring group */
    var TEMP_MEASURE_GROUP_NAME = "__SmartAlignDistribute_TempMeasure__";

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

    var OPTIONS_MARGINS = [15, 5, 15, 5];    /* オプション欄の余白 / options group margins */
    var SPACING_FIELD_CHARS = 3;             /* 間隔の入力欄の幅（文字数）/ width of the spacing field */

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

    var LABELS = {
        dialog: {
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },
        panel: {
            direction: { ja: "方向", en: "Direction" },
            spacing: { ja: "間隔", en: "Spacing" },
            alignHorizontal: { ja: "揃え（左右）", en: "Align (H)" },
            alignVertical: { ja: "揃え（上下）", en: "Align (V)" }
        },
        radio: {
            directionAuto: { ja: "自動", en: "Auto" },
            directionVertical: { ja: "縦", en: "Vertical" },
            directionHorizontal: { ja: "横", en: "Horizontal" },
            alignNone: { ja: "なし", en: "None" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignTop: { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        checkbox: {
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            measureText: { ja: "テキストの高さを計測", en: "Measure text height" },
            random: { ja: "ランダム", en: "Random" }
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
            directionAuto: {
                ja: "選択範囲が横長なら横並び、縦長なら縦並びとして扱います。",
                en: "Lays the objects out in a row when the selection is wider than tall, in a column otherwise."
            },
            directionVertical: { ja: "上から下へ縦に並べます。", en: "Stacks the objects from top to bottom." },
            directionHorizontal: { ja: "左から右へ横に並べます。", en: "Lays the objects out from left to right." },
            spacing: {
                ja: "オブジェクト間のすき間。マイナス値で重ねられます。",
                en: "Gap between objects. Negative values overlap them."
            },
            alignHorizontal: {
                ja: "縦に並べたときの左右の揃え方です。N／L／C／R キーでも切り替えられます。",
                en: "Horizontal alignment used when stacking vertically. The keys N / L / C / R switch it."
            },
            alignVertical: {
                ja: "横に並べたときの上下の揃え方です。N／T／M／B キーでも切り替えられます。",
                en: "Vertical alignment used when laying out horizontally. The keys N / T / M / B switch it."
            },
            usePreviewBounds: {
                ja: "線や効果を含む見た目の境界で整列します。オフはパスのみのジオメトリ境界。",
                en: "Align by visible bounds (incl. strokes/effects). Off uses geometric (path-only) bounds."
            },
            measureText: {
                ja: "縦並び時のみ、テキストを一度だけ複製→アウトライン化して境界を計測します（ダイアログ中だけキャッシュ）。",
                en: "Only in vertical layout, measures text by duplicating and outlining once (cached for this dialog only)."
            },
            random: {
                ja: "並び順をランダムに入れ替えます（左上の位置は維持）。",
                en: "Shuffle the stacking order at random (top-left position is kept)."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            errorPrefix: { ja: "エラーが発生しました: ", en: "An error has occurred: " }
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
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューを作り直す（前回分を undo してから processFn を実行）
     * @param {{isUndo: boolean}} previewState - プレビューの状態（undo すべき変更があるか）
     * @param {function} processFn - プレビューとして実行する処理
     * @param {boolean} isEnabled - プレビューを表示するか（false なら前回分を戻すだけ）
     * @returns {void}
     */
    function runPreview(previewState, processFn, isEnabled) {
        /* app.undo() と DOM 操作は状況によって失敗する / undo and DOM edits may fail */
        try {
            if (isEnabled) {
                if (previewState.isUndo) app.undo();
                else previewState.isUndo = true;
                processFn();
                app.redraw();
            } else if (previewState.isUndo) {
                app.undo();
                app.redraw();
                previewState.isUndo = false;
            }
        } catch (err) { }
    }

    /**
     * プレビュー分を巻き戻す（確定処理の直前・キャンセル時）
     * @param {{isUndo: boolean}} previewState - プレビューの状態
     * @returns {void}
     */
    function undoPreview(previewState) {
        /* 取り消す履歴が無いと app.undo() が失敗することがある / undo may fail with no history */
        try {
            if (previewState.isUndo) app.undo();
        } catch (err) { }
        previewState.isUndo = false;
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
    // 境界と並び順 / Bounds and ordering
    // =========================================

    /**
     * 選択全体の幅・高さを返す（クリップグループはマスクで測る）
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {{spanX: number, spanY: number}} 幅と高さ
     */
    function getSelectionSpan(targetItems, usePreviewBounds) {
        var unionBounds = getClipAwareUnionBounds(targetItems, usePreviewBounds);
        if (!unionBounds) return { spanX: 0, spanY: 0 };
        return { spanX: (unionBounds[2] - unionBounds[0]), spanY: (unionBounds[1] - unionBounds[3]) };
    }

    /**
     * 選択の縦横比から並べる方向を判定する（geometricBounds で安定させる）
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {string} "horizontal" または "vertical"
     */
    function detectDirection(targetItems) {
        var selectionSpan = getSelectionSpan(targetItems, false);
        return (selectionSpan.spanX >= selectionSpan.spanY) ? "horizontal" : "vertical";
    }

    /**
     * 上→下（同じ高さなら左→右）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortTopToBottom(targetItems) {
        var sortedItems = targetItems.slice();
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.top !== itemB.top) return itemB.top - itemA.top;
            return itemA.left - itemB.left;
        });
        return sortedItems;
    }

    /**
     * 左→右（同じ位置なら上→下）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortLeftToRight(targetItems) {
        var sortedItems = targetItems.slice();
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.left !== itemB.left) return itemA.left - itemB.left;
            return itemB.top - itemA.top;
        });
        return sortedItems;
    }

    /**
     * 0〜count-1 のインデックスを Fisher-Yates でシャッフルした配列を返す
     * @param {number} count - 要素数
     * @returns {number[]} シャッフルしたインデックス
     */
    function makeShuffledIndices(count) {
        var indices = [];
        for (var i = 0; i < count; i++) indices.push(i);
        for (var j = indices.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var swappedIndex = indices[j]; indices[j] = indices[k]; indices[k] = swappedIndex;
        }
        return indices;
    }

    /**
     * オブジェクト（または子孫）がテキストを含むかを返す
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {boolean} テキストを含めば true
     */
    function containerHasText(targetItem) {
        if (!targetItem) return false;
        /* textFrames / pageItems を持たない種類がある / not every item kind has textFrames or pageItems */
        try {
            if (targetItem.typename === "TextFrame") return true;
            if (targetItem.textFrames && targetItem.textFrames.length > 0) return true;
            if (targetItem.pageItems && targetItem.pageItems.length) {
                for (var i = 0; i < targetItem.pageItems.length; i++) {
                    if (containerHasText(targetItem.pageItems[i])) return true;
                }
            }
        } catch (e) { }
        return false;
    }

    /**
     * いずれかのオブジェクトがテキストを含むかを返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {boolean} テキストを含むものがあれば true
     */
    function anyItemHasText(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            if (containerHasText(targetItems[i])) return true;
        }
        return false;
    }

    // =========================================
    // キー操作 / Keyboard
    // =========================================

    /* 揃えのショートカットキーと、対応するラジオのキー / Shortcut keys mapped to a radio in each row */
    var HORIZONTAL_ALIGN_RADIO_BY_KEY = { "N": "none", "L": "left", "C": "center", "R": "right" };
    var VERTICAL_ALIGN_RADIO_BY_KEY = { "N": "none", "T": "top", "M": "middle", "B": "bottom" };

    /**
     * N / L / C / R / T / M / B キーで揃えのラジオを選ぶショートカットの表を作る
     * 横並びのときは上下の揃え、縦並びのときは左右の揃えを操作する
     * @param {{horizontal: Object, vertical: Object, getDirection: function}} alignRadios - 揃えのラジオ一式と方向の取得関数
     * @returns {Object} キー → ラジオを返す関数（いまの向きで使わないキーは false）
     */
    function buildAlignShortcutMap(alignRadios) {
        var shortcutMap = {};
        var keyName;
        /**
         * キーに対応するラジオを、いまの向きの行から選ぶ関数を作る
         * @param {string} pressedKey - キー名
         * @returns {Function} ラジオを返す関数（使わないキーなら false）
         */
        function makeAlignKeyTarget(pressedKey) {
            return function () {
                var isHorizontalLayout = (alignRadios.getDirection() === "horizontal");
                var radioByKey = isHorizontalLayout ? VERTICAL_ALIGN_RADIO_BY_KEY : HORIZONTAL_ALIGN_RADIO_BY_KEY;
                var radioSet = isHorizontalLayout ? alignRadios.vertical : alignRadios.horizontal;
                var radioKey = radioByKey[pressedKey];
                return radioKey ? radioSet[radioKey] : false;
            };
        }
        for (keyName in HORIZONTAL_ALIGN_RADIO_BY_KEY) shortcutMap[keyName] = makeAlignKeyTarget(keyName);
        for (keyName in VERTICAL_ALIGN_RADIO_BY_KEY) shortcutMap[keyName] = makeAlignKeyTarget(keyName);
        return shortcutMap;
    }

    // =========================================
    // アウトライン計測 / Outline measurement
    // =========================================

    /**
     * コンテナ内のすべてのテキストをアウトライン化する
     * @param {PageItem} containerItem - テキスト、またはテキストを含むグループ
     * @returns {void}
     */
    function outlineAllTextInContainer(containerItem) {
        if (!containerItem) return;
        /* createOutline() は空のテキストなどで失敗する / createOutline() can fail, e.g. on empty text */
        try {
            if (containerItem.typename === "TextFrame") {
                containerItem.createOutline();
                return;
            }
            if (containerItem.textFrames && containerItem.textFrames.length) {
                /* createOutline で要素が消えるため後ろから走査 / iterate backwards: createOutline removes the frame */
                for (var i = containerItem.textFrames.length - 1; i >= 0; i--) {
                    try { containerItem.textFrames[i].createOutline(); } catch (e1) { }
                }
            }
        } catch (e) { }
    }

    /**
     * 複製をアウトライン化して境界を測る（boundsCache に控えて再利用）
     * @param {PageItem} originalItem - 測るオブジェクト
     * @param {Array} boundsCache - [{item: PageItem, bounds: number[]}] の控え
     * @returns {number[]|null} [左, 上, 右, 下]（テキストを含まない・測れないときは null）
     */
    function measureOutlineBoundsOnce(originalItem, boundsCache) {
        if (!originalItem) return null;

        for (var i = 0; i < boundsCache.length; i++) {
            if (boundsCache[i].item === originalItem) return boundsCache[i].bounds;
        }
        if (!containerHasText(originalItem)) return null;

        var doc = app.activeDocument;
        var measureGroup = null;
        try {
            var measureLayer = null;
            try { measureLayer = originalItem.layer; } catch (eLayer) { }
            if (!measureLayer) measureLayer = doc.activeLayer;
            measureGroup = measureLayer.groupItems.add();
            measureGroup.name = TEMP_MEASURE_GROUP_NAME;

            var duplicatedItem = null;
            try {
                duplicatedItem = originalItem.duplicate(measureGroup, ElementPlacement.PLACEATEND);
            } catch (eDup) {
                duplicatedItem = originalItem.duplicate();
                duplicatedItem.move(measureGroup, ElementPlacement.PLACEATEND);
            }

            outlineAllTextInContainer(duplicatedItem);

            var measuredBounds = null;
            try { measuredBounds = measureGroup.visibleBounds; }
            catch (eVisible) { try { measuredBounds = measureGroup.geometricBounds; } catch (eGeometric) { } }
            if (!measuredBounds) return null;

            boundsCache.push({ item: originalItem, bounds: [measuredBounds[0], measuredBounds[1], measuredBounds[2], measuredBounds[3]] });
            return [measuredBounds[0], measuredBounds[1], measuredBounds[2], measuredBounds[3]];
        } catch (e) {
            return null;
        } finally {
            try { if (measureGroup) measureGroup.remove(); } catch (eRemove) { }
        }
    }

    // =========================================
    // レイアウト適用 / Layout application
    // =========================================

    /**
     * 並べる順（並べ替え、またはシャッフル）と、ランダム時の基準位置を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {string} direction - "horizontal" / "vertical"
     * @param {boolean} useRandom - ランダムに並べるか
     * @param {number[]|null} previousShuffleOrder - 前回のシャッフル順（件数が同じなら使い回す）
     * @returns {{items: PageItem[], baseLeft: number, baseTop: number, shuffleOrder: number[]}} 並べる順と基準位置
     */
    function buildArrangementOrder(targetItems, direction, useRandom, previousShuffleOrder) {
        var arrangement = { items: [], baseLeft: null, baseTop: null, shuffleOrder: previousShuffleOrder };

        if (useRandom) {
            if (!arrangement.shuffleOrder || arrangement.shuffleOrder.length !== targetItems.length) {
                arrangement.shuffleOrder = makeShuffledIndices(targetItems.length);
            }
            for (var i = 0; i < arrangement.shuffleOrder.length; i++) {
                arrangement.items.push(targetItems[arrangement.shuffleOrder[i]]);
            }
            /* ランダム時は元の左上を基準として保持 / Preserve top-left base for random */
            for (var k = 0; k < targetItems.length; k++) {
                var targetItem = targetItems[k];
                if (!targetItem) continue;
                if (arrangement.baseLeft === null || targetItem.left < arrangement.baseLeft) arrangement.baseLeft = targetItem.left;
                if (arrangement.baseTop === null || targetItem.top > arrangement.baseTop) arrangement.baseTop = targetItem.top;
            }
            if (arrangement.baseLeft === null && arrangement.items[0]) arrangement.baseLeft = arrangement.items[0].left;
            if (arrangement.baseTop === null && arrangement.items[0]) arrangement.baseTop = arrangement.items[0].top;
        } else {
            arrangement.shuffleOrder = null;
            arrangement.items = (direction === "horizontal") ? sortLeftToRight(targetItems) : sortTopToBottom(targetItems);
        }
        return arrangement;
    }

    /**
     * オブジェクトの幅または高さの最大値を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {function} getBounds - 境界 [左, 上, 右, 下] を返す関数
     * @param {boolean} measureWidth - true なら幅、false なら高さ
     * @returns {number} 最大値
     */
    function computeMaxSize(targetItems, getBounds, measureWidth) {
        var maxSize = 0;
        for (var i = 0; i < targetItems.length; i++) {
            var bounds = getBounds(targetItems[i]);
            var itemSize = measureWidth ? (bounds[2] - bounds[0]) : (bounds[1] - bounds[3]);
            if (itemSize > maxSize) maxSize = itemSize;
        }
        return maxSize;
    }

    /**
     * 左→右に横並びで配置し、上下揃えを適用する
     * @param {PageItem[]} targetItems - 並べる順のオブジェクト
     * @param {number} startX - 先頭の左端
     * @param {number} startY - 揃えの基準にする上端
     * @param {number} referenceHeight - 揃えの基準にする高さ
     * @param {number} spacingPt - 間隔（pt）
     * @param {string} vAlignMode - "none" / "top" / "middle" / "bottom"
     * @param {function} getBounds - 境界を返す関数
     * @returns {void}
     */
    function applyHorizontalLayout(targetItems, startX, startY, referenceHeight, spacingPt, vAlignMode, getBounds) {
        var currentX = startX;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getBounds(targetItem);
            var itemWidth = bounds[2] - bounds[0];

            /* Xは左基準で配置 / Place by left edge on X */
            targetItem.left = targetItem.left + (currentX - bounds[0]);

            /* 上下揃え / Vertical alignment */
            if (vAlignMode !== "none") {
                var cellTop = startY;
                var cellBottom = startY - referenceHeight;
                var dy = 0;
                if (vAlignMode === "middle") dy = ((cellTop + cellBottom) / 2) - ((bounds[1] + bounds[3]) / 2);
                else if (vAlignMode === "bottom") dy = cellBottom - bounds[3];
                else dy = cellTop - bounds[1]; /* top */
                targetItem.top = targetItem.top + dy;
            }

            if (i < targetItems.length - 1) {
                currentX += itemWidth + spacingPt;
            }
        }
    }

    /**
     * 上→下に縦並びで配置し、左右揃えを適用する
     * @param {PageItem[]} targetItems - 並べる順のオブジェクト
     * @param {number} startX - 揃えの基準にする左端
     * @param {number} startY - 先頭の上端
     * @param {number} referenceWidth - 揃えの基準にする幅
     * @param {number} spacingPt - 間隔（pt）
     * @param {string} hAlignMode - "none" / "left" / "center" / "right"
     * @param {function} getBounds - 境界を返す関数
     * @returns {void}
     */
    function applyVerticalLayout(targetItems, startX, startY, referenceWidth, spacingPt, hAlignMode, getBounds) {
        var currentY = startY;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getBounds(targetItem);
            var itemHeight = bounds[1] - bounds[3];

            /* 左右揃え / Horizontal alignment */
            if (hAlignMode !== "none") {
                var cellLeft = startX;
                var cellRight = cellLeft + referenceWidth;
                var dx = 0;
                if (hAlignMode === "center") dx = ((cellLeft + cellRight) / 2) - ((bounds[0] + bounds[2]) / 2);
                else if (hAlignMode === "right") dx = cellRight - bounds[2];
                else dx = cellLeft - bounds[0]; /* left */
                targetItem.left = targetItem.left + dx;
            }

            /* YはTop揃えで積む / Stack by top on Y */
            targetItem.top = targetItem.top + (currentY - bounds[1]);

            if (i < targetItems.length - 1) {
                currentY -= itemHeight + spacingPt;
            }
        }
    }

    /**
     * ランダム時に、先頭のオブジェクトが元の左上に来るよう全体をずらす
     * @param {PageItem[]} targetItems - 並べたオブジェクト
     * @param {number} baseLeft - 元の左端
     * @param {number} baseTop - 元の上端
     * @returns {void}
     */
    function applyRandomBaseOffset(targetItems, baseLeft, baseTop) {
        if (!targetItems.length) return;
        var dx = baseLeft - targetItems[0].left;
        var dy = baseTop - targetItems[0].top;
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            targetItems[i].left += dx;
            targetItems[i].top += dy;
        }
    }

    /**
     * 並べる順に従ってオブジェクトを配置する
     * @param {{items: PageItem[], baseLeft: number, baseTop: number}} arrangement - buildArrangementOrder() の結果
     * @param {{direction: string, spacingPt: number, useRandom: boolean, hAlignMode: string, vAlignMode: string}} layoutSettings - 並べ方
     * @param {function} getBounds - 境界 [左, 上, 右, 下] を返す関数
     * @returns {void}
     */
    function placeArrangement(arrangement, layoutSettings, getBounds) {
        var arrangedItems = arrangement.items;
        if (!arrangedItems.length) return;

        var startBounds = getBounds(arrangedItems[0]);
        var startX = startBounds[0];
        var startY = startBounds[1];

        if (layoutSettings.direction === "horizontal") {
            var referenceHeight = computeMaxSize(arrangedItems, getBounds, false);
            applyHorizontalLayout(arrangedItems, startX, startY, referenceHeight, layoutSettings.spacingPt, layoutSettings.vAlignMode, getBounds);
        } else {
            var referenceWidth = computeMaxSize(arrangedItems, getBounds, true);
            applyVerticalLayout(arrangedItems, startX, startY, referenceWidth, layoutSettings.spacingPt, layoutSettings.hAlignMode, getBounds);
        }

        if (layoutSettings.useRandom) {
            applyRandomBaseOffset(arrangedItems, arrangement.baseLeft, arrangement.baseTop);
        }
    }

    // =========================================
    // ダイアログ部品 / Dialog parts
    // =========================================

    /**
     * ［方向］パネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{auto: RadioButton, vertical: RadioButton, horizontal: RadioButton}} 方向のラジオ
     */
    function addDirectionPanel(parentDialog) {
        var directionPanel = parentDialog.add("panel", undefined, getLabel("panel.direction"));
        setupPanel(directionPanel, 6);
        directionPanel.orientation = "row";
        directionPanel.alignChildren = ["center", "center"];
        var directionRadios = {
            auto: addTooltipRadio(directionPanel, getLabel("radio.directionAuto"), getLabel("tooltip.directionAuto")),
            vertical: addTooltipRadio(directionPanel, getLabel("radio.directionVertical"), getLabel("tooltip.directionVertical")),
            horizontal: addTooltipRadio(directionPanel, getLabel("radio.directionHorizontal"), getLabel("tooltip.directionHorizontal"))
        };
        directionRadios.vertical.value = true; /* デフォルトを「縦」に / Default to Vertical */
        return directionRadios;
    }

    /**
     * ツールチップ付きのラジオボタンを追加する
     * @param {object} parentGroup - 追加先のパネルまたはグループ
     * @param {string} radioText - ラジオの表示名
     * @param {string} tooltipText - ツールチップ
     * @returns {RadioButton} 追加したラジオ
     */
    function addTooltipRadio(parentGroup, radioText, tooltipText) {
        var radioButton = parentGroup.add("radiobutton", undefined, radioText);
        radioButton.helpTip = tooltipText;
        return radioButton;
    }

    /**
     * ［間隔］パネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @param {function} onSpacingStep - ∧∨・↑↓キーで値を変えたあとに呼ぶ処理
     * @returns {EditText} 間隔の入力欄
     */
    function addSpacingPanel(parentDialog, onSpacingStep) {
        var spacingPanel = parentDialog.add("panel", undefined, getLabel("panel.spacing"));
        setupPanel(spacingPanel);
        spacingPanel.alignChildren = ["center", "center"];
        var spacingRowGroup = spacingPanel.add("group");
        spacingRowGroup.orientation = "row";
        spacingRowGroup.alignChildren = ["left", "center"];
        /* ∧∨と入力欄は隙間0で突き合わせる。マイナス値も許す（重ねる）/ Stepper butts the field; negatives allowed (overlap) */
        var spacingFieldGroup = spacingRowGroup.add("group");
        spacingFieldGroup.orientation = "row";
        spacingFieldGroup.alignChildren = ["left", "center"];
        spacingFieldGroup.spacing = 0;
        spacingFieldGroup.margins = 0;
        var spacingInput;
        var spacingStepper = addStepper(spacingFieldGroup, function () { return spacingInput; }, {
            onStep: function () { onSpacingStep(); }
        });
        spacingInput = spacingFieldGroup.add("edittext", undefined, "0");
        spacingInput.characters = SPACING_FIELD_CHARS;
        spacingInput.helpTip = getLabel("tooltip.spacing");
        bindSteppedArrowKeys(spacingInput, spacingStepper); /* ↑↓キーも∧∨と同じ処理 / arrow keys share the stepper */
        spacingRowGroup.add("statictext", undefined, getUnitInfo().label);
        return spacingInput;
    }

    /**
     * 揃えのラジオを1行ぶん追加する
     * @param {Panel} alignPanel - 追加先のパネル
     * @param {Array} radioDefs - [ラジオのキー, LABELS のパス] の配列
     * @param {string} defaultKey - 最初に選ぶラジオのキー
     * @param {string} tooltipText - 行のラジオすべてに付けるツールチップ
     * @returns {{rowGroup: Group, radios: Object}} 行のグループと、キー → ラジオ
     */
    function addAlignRadioRow(alignPanel, radioDefs, defaultKey, tooltipText) {
        var rowGroup = alignPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        var alignRadios = {};
        for (var i = 0; i < radioDefs.length; i++) {
            alignRadios[radioDefs[i][0]] = rowGroup.add("radiobutton", undefined, getLabel(radioDefs[i][1]));
        }
        alignRadios[defaultKey].value = true;
        for (var radioKey in alignRadios) {
            alignRadios[radioKey].helpTip = tooltipText;
        }
        return { rowGroup: rowGroup, radios: alignRadios };
    }

    /**
     * ［揃え］パネル（左右の揃えと上下の揃えの2行）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{panel: Panel, horizontalRow: Group, verticalRow: Group, horizontal: Object, vertical: Object}} パネルと各行のラジオ
     */
    function addAlignPanel(parentDialog) {
        var alignPanel = parentDialog.add("panel", undefined, "");
        setupPanel(alignPanel, 6);
        alignPanel.alignChildren = ["left", "center"];

        /* 既定は「中央」/ Default to Center (Middle) */
        var horizontalRow = addAlignRadioRow(alignPanel, [
            ["none", "radio.alignNone"], ["left", "radio.alignLeft"], ["center", "radio.alignCenter"], ["right", "radio.alignRight"]
        ], "center", getLabel("tooltip.alignHorizontal"));
        var verticalRow = addAlignRadioRow(alignPanel, [
            ["none", "radio.alignNone"], ["top", "radio.alignTop"], ["middle", "radio.alignMiddle"], ["bottom", "radio.alignBottom"]
        ], "middle", getLabel("tooltip.alignVertical"));

        return {
            panel: alignPanel,
            horizontalRow: horizontalRow.rowGroup,
            verticalRow: verticalRow.rowGroup,
            horizontal: horizontalRow.radios,
            vertical: verticalRow.radios
        };
    }

    /**
     * オプションのチェックボックス（プレビュー境界・テキスト計測・ランダム）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{usePreviewBounds: Checkbox, measureText: Checkbox, random: Checkbox}} チェックボックス
     */
    function addOptionsGroup(parentDialog) {
        var optionsGroup = parentDialog.add("group");
        optionsGroup.orientation = "column";
        optionsGroup.alignChildren = ["left", "center"];
        optionsGroup.alignment = ["fill", "top"];
        optionsGroup.margins = OPTIONS_MARGINS;

        return {
            usePreviewBounds: addOptionCheckbox(optionsGroup, getLabel("checkbox.usePreviewBounds"), getLabel("tooltip.usePreviewBounds"), true),
            measureText: addOptionCheckbox(optionsGroup, getLabel("checkbox.measureText"), getLabel("tooltip.measureText"), false),
            random: addOptionCheckbox(optionsGroup, getLabel("checkbox.random"), getLabel("tooltip.random"), false)
        };
    }

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} checkboxText - 表示名
     * @param {string} tooltipText - ツールチップ
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentGroup, checkboxText, tooltipText, initialValue) {
        var optionCheckbox = parentGroup.add("checkbox", undefined, checkboxText);
        optionCheckbox.value = initialValue;
        optionCheckbox.helpTip = tooltipText;
        return optionCheckbox;
    }

    /**
     * ラジオの組から選ばれているもののキーを返す
     * @param {object} radioSet - キー → ラジオ
     * @param {string} fallbackKey - どれも選ばれていないときのキー
     * @returns {string} 選ばれているラジオのキー
     */
    function getCheckedKey(radioSet, fallbackKey) {
        for (var radioKey in radioSet) {
            if (radioSet[radioKey].value) return radioKey;
        }
        return fallbackKey;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを表示し、プレビューしながら整列・分布を実行する
     * @returns {boolean} OK で確定したら true、キャンセルなら false
     */
    function showArrangeDialog() {
        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(alignDialog);

        /* キャンセル時に復元するため現在のプリファレンスを保存 / Preserve preference for cancel */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

        /* 選択スナップショット / Snapshot selection */
        var originalSelection = app.activeDocument.selection.slice();
        var previewState = { isUndo: false };
        var detectedDirection = detectDirection(originalSelection);
        var selectionHasText = anyItemHasText(originalSelection);

        var directionRadios = addDirectionPanel(alignDialog);
        var spacingInput = addSpacingPanel(alignDialog, requestPreviewUpdate);
        var alignControls = addAlignPanel(alignDialog);
        var optionCheckboxes = addOptionsGroup(alignDialog);
        var buttonRow = addButtonRow(alignDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /* ---- 内部状態 / Internal state ---- */
        var randomOrderCache = null;
        var outlineMeasureCache = []; /* [{item:PageItem, bounds:[l,t,r,b]}] */
        var isPreviewUpdating = false;
        var lastPreviewTime = 0;
        var previewTaskId = 0;

        /**
         * 選択中のラジオから実際に並べる方向を返す（自動は判定結果）
         * @returns {string} "horizontal" または "vertical"
         */
        function getEffectiveDirection() {
            if (directionRadios.vertical.value) return "vertical";
            if (directionRadios.horizontal.value) return "horizontal";
            return detectedDirection;
        }

        /**
         * レイアウト用の境界を返す（縦並びでテキスト計測が ON ならアウトラインの実測値）
         * @param {PageItem} targetItem - 対象のオブジェクト
         * @returns {number[]} [左, 上, 右, 下]
         */
        function getBoundsForLayout(targetItem) {
            if (optionCheckboxes.measureText.value && getEffectiveDirection() !== "horizontal") {
                var outlinedBounds = measureOutlineBoundsOnce(targetItem, outlineMeasureCache);
                if (outlinedBounds) return outlinedBounds;
            }
            return getClipAwareBounds(targetItem, optionCheckboxes.usePreviewBounds.value);
        }

        /**
         * 間隔の入力値を現在の単位から pt に換算する
         * @returns {number} 間隔（pt）
         */
        function getSpacingPt() {
            var spacingValue = parseFloat(spacingInput.text);
            if (isNaN(spacingValue)) spacingValue = 0;
            return spacingValue * getUnitInfo().pointsPerUnit;
        }

        /**
         * 現在の設定で選択を並べる
         * @returns {void}
         */
        function applyLayoutToSelection() {
            if (!originalSelection || originalSelection.length === 0) return;
            var direction = getEffectiveDirection();
            var spacingPt = getSpacingPt();
            var useRandom = optionCheckboxes.random.value;
            var arrangement = buildArrangementOrder(originalSelection, direction, useRandom, randomOrderCache);
            randomOrderCache = arrangement.shuffleOrder;
            placeArrangement(arrangement, {
                direction: direction,
                spacingPt: spacingPt,
                useRandom: useRandom,
                hAlignMode: getCheckedKey(alignControls.horizontal, "left"),
                vAlignMode: getCheckedKey(alignControls.vertical, "top")
            }, getBoundsForLayout);
        }

        /**
         * 方向に合わせて揃えパネルの見出しと有効状態、テキスト計測の有効状態を切り替える
         * @returns {void}
         */
        function syncAlignUI() {
            var isHorizontal = (getEffectiveDirection() === "horizontal");
            alignControls.panel.text = getLabel(isHorizontal ? "panel.alignVertical" : "panel.alignHorizontal");
            alignControls.horizontalRow.enabled = !isHorizontal;
            alignControls.verticalRow.enabled = isHorizontal;

            /* テキスト計測は縦並び時かつ選択にテキストがある場合のみ有効 / Enable text-measure only when vertical & selection has text */
            optionCheckboxes.measureText.enabled = !isHorizontal && selectionHasText;

            /* 表示前のレイアウトは失敗することがある / layout before show may fail */
            try { alignDialog.layout.layout(true); } catch (e) { }
        }

        /**
         * 連続呼び出しを間引きながらプレビューを作り直す
         * @returns {void}
         */
        function updatePreview() {
            if (isPreviewUpdating) return;
            var now = new Date().getTime();
            if (now - lastPreviewTime < PREVIEW_MIN_INTERVAL_MS) return;
            lastPreviewTime = now;

            isPreviewUpdating = true;
            try {
                try {
                    app.preferences.setBooleanPreference("includeStrokeInBounds", optionCheckboxes.usePreviewBounds.value);
                } catch (e) { }
                runPreview(previewState, applyLayoutToSelection, true);
            } finally {
                isPreviewUpdating = false;
            }
        }

        /**
         * プレビューを予約する（ScriptUI のイベント中に DOM を触ると落ちるため scheduleTask で遅らせる）
         * @returns {void}
         */
        function requestPreviewUpdate() {
            try { if (previewTaskId) app.cancelTask(previewTaskId); } catch (e) { }
            $.global.__SAT_updatePreview = updatePreview;
            try {
                previewTaskId = app.scheduleTask('$.global.__SAT_updatePreview && $.global.__SAT_updatePreview();', PREVIEW_SCHEDULE_MS, false);
            } catch (e) {
                updatePreview();
            }
        }

        /**
         * 方向を変えたら揃えパネルを同期してプレビューを更新する
         * @returns {void}
         */
        function onDirectionChanged() {
            syncAlignUI();
            requestPreviewUpdate();
        }

        /* ---- イベントバインド / Event bindings ---- */
        var alignRadios = {
            horizontal: alignControls.horizontal,
            vertical: alignControls.vertical,
            getDirection: getEffectiveDirection
        };
        var radioKey;
        for (radioKey in alignControls.horizontal) alignControls.horizontal[radioKey].onClick = requestPreviewUpdate;
        for (radioKey in alignControls.vertical) alignControls.vertical[radioKey].onClick = requestPreviewUpdate;
        /* 揃えのキーは間隔の欄に入力中でも効かせる / Align keys also work while the spacing field has focus */
        addKeyShortcuts(alignDialog, buildAlignShortcutMap(alignRadios), { numericFields: [spacingInput] });

        directionRadios.auto.onClick = onDirectionChanged;
        directionRadios.vertical.onClick = onDirectionChanged;
        directionRadios.horizontal.onClick = onDirectionChanged;

        optionCheckboxes.usePreviewBounds.onClick = function () {
            /* 境界モードが変わるとアウトライン計測結果も再計算 / Outline cache invalid when bounds mode changes */
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };
        optionCheckboxes.random.onClick = function () {
            randomOrderCache = null;
            requestPreviewUpdate();
        };
        optionCheckboxes.measureText.onClick = function () {
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };

        /* 初期化 / Init */
        syncAlignUI();
        requestPreviewUpdate();
        spacingInput.active = true;

        prepareDialogWindow(alignDialog, SCRIPT_NAME);
        var dialogResult = alignDialog.show();
        try { if (previewTaskId) app.cancelTask(previewTaskId); } catch (e) { }

        if (dialogResult !== 1) {
            /* キャンセル: プレビューを undo してプリファレンスも戻す / Cancel: undo preview & restore preference */
            undoPreview(previewState);
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return false;
        }

        /* 確定: プレビュー分を巻き戻して本実行（undo 履歴を 1 件にまとめる） / OK: undo preview, then commit as single history entry */
        undoPreview(previewState);
        try { app.preferences.setBooleanPreference("includeStrokeInBounds", optionCheckboxes.usePreviewBounds.value); } catch (e) { }
        applyLayoutToSelection();
        app.redraw();
        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        /* ドキュメントが無いと app.activeDocument が例外になる / activeDocument throws with no document */
        try {
            var selectedItems = app.activeDocument.selection;
            if (!selectedItems || selectedItems.length === 0) {
                alert(getLabel("alert.noSelection"));
                return;
            }
            showArrangeDialog();
        } catch (e) {
            alert(getLabel("alert.errorPrefix") + e.message);
        }
    }

    main();

})();
