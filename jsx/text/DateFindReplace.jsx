#target illustrator
#targetengine "DateFindReplaceEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストフレームから日付を検索し、選択した項目だけを置換します。
オブジェクトが選択されている場合は、その選択範囲のテキストフレームだけを検索対象にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DateFindReplace.md

### Overview

Finds dates in the text frames of the document and replaces only the ones you tick.
When objects are selected, only the text frames within that selection are searched.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DateFindReplace.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DateFindReplace";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DateFindReplace.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DateFindReplace.md"; /* README (English) */

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

    var FOUND_PANEL_MIN_WIDTH = 240;       /* ［見つかった日付］パネルの最小幅 / Minimum width of the found-dates panel */
    var DATE_ROW_INDENT = 20;              /* 日付行の左インデント / Left indent of the date rows */
    var DATE_CHECKBOX_WIDTH = 220;         /* 日付チェックボックスの幅（ラベル切れ対策）/ Width of the date checkboxes */
    var CHECK_LABEL_WIDTH = 130;           /* 確認パネルの項目名の幅（値の左端をそろえる）/ Label width in the check panel */

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
    // 和暦 / Japanese eras
    // =========================================

    /* 令和元年 = 2019 年（令和Y = 西暦 Y + 2018）／平成元年 = 1989 年（平成Y = 西暦 Y + 1988） */
    var REIWA_BASE_YEAR = 2018;
    var HEISEI_BASE_YEAR = 1988;

    /* 各元号の有効範囲（getValidationErrorMessage で使用） */
    var REIWA_START_DATE = new Date(2019, 4, 1);                 /* 2019/5/1 〜 */
    var HEISEI_START_DATE = new Date(1989, 0, 8);                /* 1989/1/8 〜 */
    var HEISEI_END_DATE_EXCLUSIVE = new Date(2019, 4, 1);        /* 〜 2019/4/30 */

    /* 確認パネルの和暦表示に使う元号表（新しい順）/ Era table for the check panel (newest first) */
    var ERA_TABLE = [
        { name: "令和", startDate: REIWA_START_DATE, baseYear: REIWA_BASE_YEAR },   /* 2019/5/1〜 */
        { name: "平成", startDate: HEISEI_START_DATE, baseYear: HEISEI_BASE_YEAR }, /* 1989/1/8〜2019/4/30 */
        { name: "昭和", startDate: new Date(1926, 11, 25), baseYear: 1925 },        /* 1926/12/25〜1989/1/7 */
        { name: "大正", startDate: new Date(1912, 6, 30), baseYear: 1911 },         /* 1912/7/30〜1926/12/24 */
        { name: "明治", startDate: new Date(1868, 8, 8), baseYear: 1867 }           /* 1868/9/8〜1912/7/29 */
    ];

    /**
     * 令和の年を西暦にする
     * @param {number} reiwaYear - 令和の年
     * @returns {number} 西暦
     */
    function reiwaToGregorian(reiwaYear) { return reiwaYear + REIWA_BASE_YEAR; }

    /**
     * 西暦を令和の年にする
     * @param {number} gregorianYear - 西暦
     * @returns {number} 令和の年
     */
    function gregorianToReiwa(gregorianYear) { return gregorianYear - REIWA_BASE_YEAR; }

    /**
     * 平成の年を西暦にする
     * @param {number} heiseiYear - 平成の年
     * @returns {number} 西暦
     */
    function heiseiToGregorian(heiseiYear) { return heiseiYear + HEISEI_BASE_YEAR; }

    /**
     * 西暦を平成の年にする
     * @param {number} gregorianYear - 西暦
     * @returns {number} 平成の年
     */
    function gregorianToHeisei(gregorianYear) { return gregorianYear - HEISEI_BASE_YEAR; }

    // =========================================
    // 選択肢 / Choices
    // =========================================

    /*
       出力フォーマット選択。先頭（preserve）の表示名は LABELS.dropdown.preserveFormat。
       「元の形式を保持」を選ぶと、各マッチの元形式（区切り文字・元号）を維持して置換する。
       曜日表記は「元の形式を保持」でも隣の「曜日」ドロップダウンの選択を反映する。
       それ以外は、選択した形式で全マッチを統一して書き換える。
    */
    var FORMAT_VALUES = ["preserve", "jp", "jp-md", "dot", "dot-md", "slash", "slash-md", "reiwa-jp", "r-dot", "r-slash", "heisei-jp", "h-dot", "h-slash"];
    var FORMAT_PATTERN_LABELS = [
        "YYYY年M月D日",
        "M月D日",
        "YYYY.M.D",
        "M.D",
        "YYYY/M/D",
        "M/D",
        "令和Y年M月D日",
        "RY.M.D",
        "RY/M/D",
        "平成Y年M月D日",
        "HY.M.D",
        "HY/M/D"
    ];

    /* 曜日サフィックスのスタイル。先頭（none）の表示名は LABELS.dropdown.weekdayNone
       Weekday suffix styles; the label for "none" comes from LABELS */
    var WEEKDAY_VALUES = ["none", "kanji", "medium", "long", "full-paren", "half-paren", "en-short", "en-full"];
    var WEEKDAY_SAMPLE_LABELS = ["火", "火曜", "火曜日", "（火）", "(火)", "Tue", "Tuesday"];

    /* 曜日生成用の文字テーブル */
    var WEEKDAY_KANJI = ["日", "月", "火", "水", "木", "金", "土"];
    var WEEKDAY_EN_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    var WEEKDAY_EN_FULL = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

    // =========================================
    // 共通ヘルパー / Shared helpers
    // =========================================

    /**
     * エラー件数のメッセージを作る
     * @param {number} errorCount - 処理できなかった件数
     * @returns {string} メッセージ
     */
    function formatErrorMessage(errorCount) {
        return getLabel(LABELS.alert.lockedErrors, [errorCount]);
    }

    /**
     * テキストフレームの中心が乗っているアートボードの番号を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {number} アートボードの番号（どれにも乗っていなければ -1）
     */
    function getArtboardIndex(doc, targetItem) {
        var itemBounds = targetItem.geometricBounds;
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;

        for (var artboardIdx = 0; artboardIdx < doc.artboards.length; artboardIdx++) {
            var artboardRect = doc.artboards[artboardIdx].artboardRect;
            var minX = Math.min(artboardRect[0], artboardRect[2]);
            var maxX = Math.max(artboardRect[0], artboardRect[2]);
            var minY = Math.min(artboardRect[1], artboardRect[3]);
            var maxY = Math.max(artboardRect[1], artboardRect[3]);
            if (centerX >= minX && centerX <= maxX && centerY >= minY && centerY <= maxY) {
                return artboardIdx;
            }
        }
        return -1;
    }

    /**
     * dropdownlist の選択値を返す（選択が外れているときは代わりの値）
     * @param {DropDownList} dropdown - 対象のドロップダウン
     * @param {string[]} valuesArray - 項目ごとの内部値
     * @param {string} fallback - 選択が無いときの値
     * @returns {string} 選択中の内部値
     */
    function getDropdownValue(dropdown, valuesArray, fallback) {
        if (dropdown.selection !== null && dropdown.selection !== undefined) {
            var selectedIndex = dropdown.selection.index;
            if (typeof selectedIndex === 'number' && selectedIndex >= 0 && selectedIndex < valuesArray.length) {
                return valuesArray[selectedIndex];
            }
        }
        return fallback;
    }

    // =========================================
    // 日付の解析 / Parsing dates
    // =========================================

    /**
     * 曜日の文字列の形式名を返す（WEEKDAY_VALUES と同じ命名）
     * @param {string} weekdayText - 判定する文字列
     * @param {boolean} allowSingleKanji - 漢字1文字（「金」など）も曜日として認めるか
     * @returns {string|null} 形式名（曜日でなければ null）
     */
    function classifyWeekdayText(weekdayText, allowSingleKanji) {
        if (/^[日月火水木金土]曜日$/.test(weekdayText)) return 'long';           /* 例：金曜日 */
        if (/^[日月火水木金土]曜$/.test(weekdayText)) return 'medium';           /* 例：金曜 */
        if (/^\([日月火水木金土]\)$/.test(weekdayText)) return 'half-paren';   /* 例：(金) */
        if (/^（[日月火水木金土]）$/.test(weekdayText)) return 'full-paren';    /* 例：（金） */
        if (allowSingleKanji && /^[日月火水木金土]$/.test(weekdayText)) return 'kanji';
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)day$/.test(weekdayText)) return 'en-full';
        if (/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)$/.test(weekdayText)) return 'en-short';
        return null;
    }

    /**
     * 解析済みの日付に付いている曜日サフィックスの形式を返す（漢字1文字は対象外）
     * @param {Object} parsedDate - detectFormatAndParse() の戻り値
     * @returns {string|null} 形式名（サフィックスが無い・曜日でなければ null）
     */
    function getWeekdaySuffixStyle(parsedDate) {
        if (!parsedDate || !parsedDate.parts || !parsedDate.parts.suffix) return null;
        return classifyWeekdayText(parsedDate.parts.suffix, false);
    }

    /**
     * フレーム全体が曜日だけかを判定し、形式名を返す（単独漢字曜日はここでのみ許可）
     * @param {string} frameContent - テキストフレームの内容
     * @returns {string|null} 形式名（曜日だけでなければ null）
     */
    function detectWeekdayOnlyFrame(frameContent) {
        return classifyWeekdayText(String(frameContent).replace(/^\s+|\s+$/g, ""), true);
    }

    /**
     * 元号年の文字列化。漢字フォーマット（reiwa-jp / heisei-jp）では元年 1 を「元」と表記する
     * @param {number} eraYear - 元号の年
     * @param {boolean} useGanText - 1 年を「元」と書くか
     * @returns {string} 年の文字列
     */
    function formatEraYearText(eraYear, useGanText) {
        if (useGanText && eraYear === 1) return "元";
        return String(eraYear);
    }

    /**
     * 元の数字テキストが 2 桁ゼロ詰め（"05" 等）なら、新しい数字も同じ桁数に揃える。
     * 新しい文字列が数字以外（例：「元」）の場合は揃えない
     * @param {string} oldText - 元の数字テキスト
     * @param {number|string} newValue - 新しい値
     * @returns {string} 桁をそろえた文字列
     */
    function applyZeroPaddingFrom(oldText, newValue) {
        var newStr = String(newValue);
        if (oldText.length === 2 && oldText.charAt(0) === '0' && newStr.length === 1 && /^[0-9]$/.test(newStr)) {
            return "0" + newStr;
        }
        return newStr;
    }

    /**
     * 実在する日付かを判定する
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {boolean} 実在すれば true
     */
    function isRealDate(year, month, day) {
        var checkDate = new Date(year, month - 1, day);
        return checkDate.getFullYear() === year &&
            checkDate.getMonth() === month - 1 &&
            checkDate.getDate() === day;
    }

    /**
     * 元号形式の日付が元号の有効範囲内かを判定する
     * @param {Object} parsedDate - detectFormatAndParse() の戻り値
     * @returns {boolean} 範囲内（または元号なし）なら true
     */
    function isEraDateValid(parsedDate) {
        if (!parsedDate) return false;
        var checkDate = new Date(parsedDate.year, parsedDate.month - 1, parsedDate.day);
        if (parsedDate.era === 'reiwa') return checkDate >= REIWA_START_DATE;
        if (parsedDate.era === 'heisei') return checkDate >= HEISEI_START_DATE && checkDate < HEISEI_END_DATE_EXCLUSIVE;
        return true;
    }

    /* 日付の形式ごとの解析規則（元号系を先に並べて優先させる）。
       hasPrefix: 令和・平成・R・H の接頭辞があるか / separator: 「.」「/」区切りの形式の区切り文字（年月日の漢字で区切る形式は null）
       Parse rules per date format (era formats first). The suffix group keeps any weekday text */
    var DATE_PARSE_RULES = [
        { format: 'reiwa-jp', era: 'reiwa', hasPrefix: true, separator: null,
            pattern: /^(令和)([0-9]{1,2})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/ },
        { format: 'heisei-jp', era: 'heisei', hasPrefix: true, separator: null,
            pattern: /^(平成)([0-9]{1,2})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/ },
        { format: 'r-dot', era: 'reiwa', hasPrefix: true, separator: ".",
            pattern: /^(R)([0-9]{1,2})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'r-slash', era: 'reiwa', hasPrefix: true, separator: "/",
            pattern: /^(R)([0-9]{1,2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'h-dot', era: 'heisei', hasPrefix: true, separator: ".",
            pattern: /^(H)([0-9]{1,2})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'h-slash', era: 'heisei', hasPrefix: true, separator: "/",
            pattern: /^(H)([0-9]{1,2})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'jp', era: null, hasPrefix: false, separator: null,
            pattern: /^([0-9]{4})(年)(0?[1-9]|1[0-2])(月)(0?[1-9]|[12][0-9]|3[01])(日)(.*)$/ },
        { format: 'dot', era: null, hasPrefix: false, separator: ".",
            pattern: /^([0-9]{4})\.(0?[1-9]|1[0-2])\.(0?[1-9]|[12][0-9]|3[01])(.*)$/ },
        { format: 'slash', era: null, hasPrefix: false, separator: "/",
            pattern: /^([0-9]{4})\/(0?[1-9]|1[0-2])\/(0?[1-9]|[12][0-9]|3[01])(.*)$/ }
    ];

    /**
     * 年の文字列を西暦の数値にする（元号なら換算）
     * @param {string|null} era - 'reiwa' / 'heisei' / null
     * @param {string} yearText - 年の文字列
     * @returns {number} 西暦
     */
    function toGregorianYear(era, yearText) {
        var yearValue = parseInt(yearText, 10);
        if (era === 'reiwa') return reiwaToGregorian(yearValue);
        if (era === 'heisei') return heiseiToGregorian(yearValue);
        return yearValue;
    }

    /**
     * マッチ文字列を解析し、形式種別と各構成要素を返す。
     * 戻り値の主要フィールド：
     *   format: 'jp' | 'dot' | 'slash' | 'reiwa-jp' | 'r-dot' | 'r-slash' | 'heisei-jp' | 'h-dot' | 'h-slash'
     *   era:    'reiwa' | 'heisei' | null
     *   year, month, day: 西暦の数値（令和・平成形式の場合は変換後）
     *   parts:  元テキストの分解（prefix, year, sep1, month, sep2, day, sep3, suffix）
     * @param {string} matchText - マッチした文字列
     * @returns {Object|null} 解析結果（どの形式にも当たらなければ null）
     */
    function detectFormatAndParse(matchText) {
        var sourceText = String(matchText);
        for (var ruleIndex = 0; ruleIndex < DATE_PARSE_RULES.length; ruleIndex++) {
            var parseRule = DATE_PARSE_RULES[ruleIndex];
            var ruleMatch = sourceText.match(parseRule.pattern);
            if (!ruleMatch) continue;

            /* 接頭辞があると、以降のグループ番号が1つずれる / A prefix shifts the later groups by one */
            var groupOffset = parseRule.hasPrefix ? 1 : 0;
            var prefixText = parseRule.hasPrefix ? ruleMatch[1] : "";
            var dateParts;
            if (parseRule.separator) {
                dateParts = {
                    prefix: prefixText, year: ruleMatch[1 + groupOffset], sep1: parseRule.separator, month: ruleMatch[2 + groupOffset],
                    sep2: parseRule.separator, day: ruleMatch[3 + groupOffset], sep3: "", suffix: ruleMatch[4 + groupOffset] || ""
                };
            } else {
                dateParts = {
                    prefix: prefixText, year: ruleMatch[1 + groupOffset], sep1: ruleMatch[2 + groupOffset], month: ruleMatch[3 + groupOffset],
                    sep2: ruleMatch[4 + groupOffset], day: ruleMatch[5 + groupOffset], sep3: ruleMatch[6 + groupOffset], suffix: ruleMatch[7 + groupOffset] || ""
                };
            }
            return {
                format: parseRule.format, era: parseRule.era,
                year: toGregorianYear(parseRule.era, dateParts.year),
                month: parseInt(dateParts.month, 10),
                day: parseInt(dateParts.day, 10),
                parts: dateParts
            };
        }
        return null;
    }

    // =========================================
    // 確認パネルの表示 / Check panel text
    // =========================================

    /**
     * 2 つの日付の差を日数で返す
     * @param {Date} fromDate - 基準の日付
     * @param {Date} toDate - 比べる日付
     * @returns {number} 日数の差
     */
    function getDaysDifference(fromDate, toDate) {
        var msPerDay = 1000 * 60 * 60 * 24;
        return Math.round((toDate.getTime() - fromDate.getTime()) / msPerDay);
    }

    /**
     * 日数差ラベル（符号付き、例：「+15日」）
     * @param {number} days - 日数の差
     * @returns {string} 表示用の文字列
     */
    function formatDaysDifference(days) {
        var sign = days > 0 ? "+" : "";
        return sign + days + getLabel(LABELS.valueText.daysSuffix);
    }

    /**
     * 曜日表示（例：金曜日 / Friday）
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {string} 曜日（日付でなければ空文字）
     */
    function getWeekdayLabel(year, month, day) {
        if (isNaN(year) || isNaN(month) || isNaN(day)) return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        if (uiLang !== "ja") return WEEKDAY_EN_FULL[checkDate.getDay()];
        return WEEKDAY_KANJI[checkDate.getDay()] + "曜日";
    }

    /**
     * 西暦+月日から和暦表記を返す（例：「令和8年」「平成元年」）。明治より前は空文字
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {string} 和暦表記
     */
    function formatEraLabel(year, month, day) {
        if (isNaN(year) || isNaN(month) || isNaN(day)) return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";

        for (var i = 0; i < ERA_TABLE.length; i++) {
            if (checkDate >= ERA_TABLE[i].startDate) {
                return ERA_TABLE[i].name + formatEraYearText(year - ERA_TABLE[i].baseYear, true) + "年";
            }
        }
        return "";
    }

    // =========================================
    // 日付検索 / Searching dates
    // =========================================

    /*
       対象例：
         2026年5月8日 / 2026年5月8日金曜 / 2026年5月8日金曜日 / 2026年5月8日(金) / 2026年5月8日（金）
         2026.5.8 / 2026.5.8（金） / 2026/5/8 / 2026/5/8（金）
         令和8年5月8日 / 令和8年5月8日（金） / R8.5.8 / R8.5.8（金） / R8/5/8 / R8/5/8（金）
         英語曜日サフィックス（Fri / Friday など）も検出する。
       元号系を先に列挙して優先マッチさせる。
    */

    var MONTH_PATTERN = "(?:0?[1-9]|1[0-2])";
    var DAY_PATTERN = "(?:0?[1-9]|[12][0-9]|3[01])";

    var WEEKDAY_SUFFIX_PATTERN = "(?:[日月火水木金土]曜日|[日月火水木金土]曜|\\([日月火水木金土]\\)|（[日月火水木金土]）|Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sun|Mon|Tue|Wed|Thu|Fri|Sat)?";
    var DATE_SEARCH_REGEX = new RegExp(
        "令和[0-9]{1,2}年" + MONTH_PATTERN + "月" + DAY_PATTERN + "日" + WEEKDAY_SUFFIX_PATTERN +
        "|平成[0-9]{1,2}年" + MONTH_PATTERN + "月" + DAY_PATTERN + "日" + WEEKDAY_SUFFIX_PATTERN +
        "|R[0-9]{1,2}\\." + MONTH_PATTERN + "\\." + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|R[0-9]{1,2}/" + MONTH_PATTERN + "/" + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|H[0-9]{1,2}\\." + MONTH_PATTERN + "\\." + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|H[0-9]{1,2}/" + MONTH_PATTERN + "/" + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|[0-9]{4}年" + MONTH_PATTERN + "月" + DAY_PATTERN + "日" + WEEKDAY_SUFFIX_PATTERN +
        "|[0-9]{4}\\." + MONTH_PATTERN + "\\." + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN +
        "|[0-9]{4}/" + MONTH_PATTERN + "/" + DAY_PATTERN + WEEKDAY_SUFFIX_PATTERN,
        "g"
    );

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

    /**
     * 検索対象のテキストフレームを集める（選択があれば選択範囲内、なければドキュメント全体）
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} 検索対象のテキストフレーム
     */
    function collectTargetTextFrames(doc) {
        var textFrames = [];
        var docSelection = doc.selection;
        if (docSelection && docSelection.length > 0) {
            /* 選択あり：選択範囲内のテキストフレームのみ（グループの中・文字の選択を含む） / With a selection: only its text frames (inside groups and text selections too) */
            return collectSelectionTextFrames(docSelection);
        }
        /* 選択なし：ドキュメント全体 / No selection: the whole document */
        for (var i = 0; i < doc.textFrames.length; i++) {
            textFrames.push(doc.textFrames[i]);
        }
        return textFrames;
    }

    /**
     * テキストフレームから実在する日付を探す
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @returns {Object[]} 見つかった日付（frameIndex / text / matchIndex / artboardIndex / parsed など）
     */
    function findDateMatches(doc, textFrames) {
        var foundMatches = [];
        for (var frameIdx = 0; frameIdx < textFrames.length; frameIdx++) {
            var targetFrame = textFrames[frameIdx];
            var frameContent = "";

            try {
                frameContent = targetFrame.contents;
            } catch (eRead) {
                continue;
            }

            var artboardIdx = -1;
            try {
                artboardIdx = getArtboardIndex(doc, targetFrame);
            } catch (eArtboard) {
                artboardIdx = -1;
            }

            var match;
            DATE_SEARCH_REGEX.lastIndex = 0;

            while ((match = DATE_SEARCH_REGEX.exec(frameContent)) !== null) {
                /* 直前が数字、かつマッチが数字始まり（YYYY 系）の場合は、5 桁以上の連続数字を切り出した
                   誤検出とみなしてスキップ。R/H/令和/平成 始まりはこの判定の対象外 */
                if (match.index > 0 &&
                    /[0-9]/.test(frameContent.charAt(match.index - 1)) &&
                    /^[0-9]/.test(match[0])) {
                    continue;
                }
                var parsedDate = detectFormatAndParse(match[0]);
                if (!parsedDate) continue;
                if (!isRealDate(parsedDate.year, parsedDate.month, parsedDate.day)) continue;
                if (!isEraDateValid(parsedDate)) continue;
                foundMatches.push({
                    frameIndex: frameIdx,
                    text: match[0],
                    matchIndex: match.index,
                    artboardIndex: artboardIdx,
                    parsed: parsedDate,
                    weekdaySuffixStyle: getWeekdaySuffixStyle(parsedDate),
                    weekdayPairs: null,
                    checkbox: null
                });
            }
        }
        return foundMatches;
    }

    /**
     * 日付マッチごとに、その TextFrame の直接の親 GroupItem を見て、
     * 同じ親グループ内の「曜日のみ」TextFrame を曜日ペアとして関連付ける（match.weekdayPairs に入れる）。
     * 同一グループに複数の日付マッチがある場合は対応が曖昧なのでペアリングしない。
     * @param {Object[]} foundMatches - findDateMatches() の戻り値
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @returns {void}
     */
    function pairWeekdayFrames(foundMatches, textFrames) {
        for (var pairingMatchIndex = 0; pairingMatchIndex < foundMatches.length; pairingMatchIndex++) {
            var pairingMatch = foundMatches[pairingMatchIndex];
            var dateFrame = textFrames[pairingMatch.frameIndex];
            var parentGroup = null;
            try {
                if (dateFrame.parent && dateFrame.parent.typename === 'GroupItem') {
                    parentGroup = dateFrame.parent;
                }
            } catch (eParentGroup) { parentGroup = null; }
            if (!parentGroup) continue;

            var hasOtherDateInSameGroup = false;
            for (var otherMatchIndex = 0; otherMatchIndex < foundMatches.length; otherMatchIndex++) {
                if (otherMatchIndex === pairingMatchIndex) continue;
                try {
                    if (textFrames[foundMatches[otherMatchIndex].frameIndex].parent === parentGroup) {
                        hasOtherDateInSameGroup = true;
                        break;
                    }
                } catch (eOtherMatch) { }
            }
            if (hasOtherDateInSameGroup) continue;

            var weekdayPairs = [];
            for (var childItemIndex = 0; childItemIndex < parentGroup.pageItems.length; childItemIndex++) {
                var childItem;
                try { childItem = parentGroup.pageItems[childItemIndex]; } catch (eChildItem) { continue; }
                if (childItem === dateFrame) continue;
                if (childItem.typename !== 'TextFrame') continue;
                try {
                    var weekdayStyle = detectWeekdayOnlyFrame(childItem.contents);
                    if (weekdayStyle) {
                        weekdayPairs.push({ frame: childItem, style: weekdayStyle });
                    }
                } catch (eWeekdayContent) { }
            }
            if (weekdayPairs.length > 0) {
                pairingMatch.weekdayPairs = weekdayPairs;
            }
        }
    }

    /**
     * 曜日ドロップダウンの初期選択を決める（最初の日付の曜日サフィックス、なければ曜日ペアの形式）
     * 日付サフィックスは単独漢字曜日を許可しないため 'kanji' は曜日のみフレームからだけ拾う
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {string} WEEKDAY_VALUES の値
     */
    function detectInitialWeekdayChoice(foundMatches) {
        for (var matchIndex = 0; matchIndex < foundMatches.length; matchIndex++) {
            var matchInfo = foundMatches[matchIndex];
            if (matchInfo.weekdaySuffixStyle) return matchInfo.weekdaySuffixStyle;
            if (matchInfo.weekdayPairs && matchInfo.weekdayPairs.length > 0) return matchInfo.weekdayPairs[0].style;
        }
        return 'none';
    }

    // =========================================
    // 置換 / Replacement
    // =========================================

    /**
     * テキストフレーム内の一致範囲を書式を保持したまま置換する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} matchStart - 置換を始める文字位置
     * @param {number} oldLen - 置換する文字数
     * @param {string} newText - 新しい文字列
     * @returns {number} 実行した文字操作数（プレビュー巻き戻し用）
     */
    function replaceMatchPreserveStyle(textFrame, matchStart, oldLen, newText) {
        if (oldLen <= 0) return 0;
        var newLen = newText.length;
        try {
            if (textFrame.contents.substr(matchStart, oldLen) === newText) return 0;
        } catch (eSameTextCheck) { }
        var minLen = Math.min(oldLen, newLen);
        /* 伸長時は最後の元文字を「最後の新文字＋追加分」で一度だけ書き換えるため、ループは手前で止める */
        var loopEnd = (newLen > oldLen) ? minLen - 1 : minLen;
        var opCount = 0;

        /* 重複部分は文字単位で書き換え。各 character.contents の代入は元の文字属性を維持 */
        for (var i = 0; i < loopEnd; i++) {
            textFrame.characters[matchStart + i].contents = newText.charAt(i);
            opCount++;
        }

        if (newLen > oldLen) {
            /* 末尾文字に追加分の文字列を含めると、元の書式を引き継いだまま文字を増やせる */
            var lastIdx = matchStart + oldLen - 1;
            textFrame.characters[lastIdx].contents = newText.charAt(oldLen - 1) + newText.substring(oldLen);
            opCount++;
        } else if (newLen < oldLen) {
            /* 余った古い文字を末尾から削除 */
            for (var removeIdx = oldLen - 1; removeIdx >= newLen; removeIdx--) {
                textFrame.characters[matchStart + removeIdx].remove();
                opCount++;
            }
        }
        return opCount;
    }

    /**
     * 指定した形式の曜日文字列を作る
     * @param {string} weekdayChoice - WEEKDAY_VALUES の値
     * @param {number} year - 年
     * @param {number} month - 月
     * @param {number} day - 日
     * @returns {string} 曜日の文字列（'none' や不正な日付なら空文字）
     */
    function buildExplicitWeekday(weekdayChoice, year, month, day) {
        if (weekdayChoice === 'none') return "";
        var checkDate = new Date(year, month - 1, day);
        if (isNaN(checkDate.getTime())) return "";
        var weekdayIndex = checkDate.getDay();
        switch (weekdayChoice) {
            case 'kanji': return WEEKDAY_KANJI[weekdayIndex];
            case 'medium': return WEEKDAY_KANJI[weekdayIndex] + "曜";
            case 'long': return WEEKDAY_KANJI[weekdayIndex] + "曜日";
            case 'full-paren': return "（" + WEEKDAY_KANJI[weekdayIndex] + "）";
            case 'half-paren': return "(" + WEEKDAY_KANJI[weekdayIndex] + ")";
            case 'en-short': return WEEKDAY_EN_SHORT[weekdayIndex];
            case 'en-full': return WEEKDAY_EN_FULL[weekdayIndex];
        }
        return "";
    }

    /**
     * @typedef {object} ReplaceOptions
     * @property {number} newYear - 置換後の年（西暦）
     * @property {number} newMonth - 置換後の月
     * @property {number} newDay - 置換後の日
     * @property {string} formatChoice - FORMAT_VALUES の値
     * @property {string} weekdayChoice - WEEKDAY_VALUES の値
     * @property {boolean} preserveNumberFormat - 数字の書式を保持するか
     * @property {string} explicitWeekdayText - 付ける曜日の文字列
     */

    /**
     * 元の形式での新しい年の文字列を作る（元号なら換算し、ゼロ詰めも元に合わせる）
     * @param {Object} parsedDate - 元の日付の解析結果
     * @param {number} newYear - 置換後の年（西暦）
     * @returns {string} 年の文字列
     */
    function buildNewYearText(parsedDate, newYear) {
        var newYearText;
        if (parsedDate.era === 'reiwa') newYearText = formatEraYearText(gregorianToReiwa(newYear), parsedDate.format === 'reiwa-jp');
        else if (parsedDate.era === 'heisei') newYearText = formatEraYearText(gregorianToHeisei(newYear), parsedDate.format === 'heisei-jp');
        else newYearText = String(newYear);
        return applyZeroPaddingFrom(parsedDate.parts.year, newYearText);
    }

    /**
     * 元の形式（区切り文字・元号）を保った置換後の文字列を作る
     * @param {Object} matchInfo - 見つかった日付
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {string|null} 置換後の文字列
     */
    function buildPreservedFormatText(matchInfo, replaceOptions) {
        var parsedDate = matchInfo.parsed;
        if (!parsedDate) return null;
        var oldParts = parsedDate.parts;
        var replacedText =
            oldParts.prefix +
            buildNewYearText(parsedDate, replaceOptions.newYear) +
            oldParts.sep1 +
            applyZeroPaddingFrom(oldParts.month, replaceOptions.newMonth) +
            oldParts.sep2 +
            applyZeroPaddingFrom(oldParts.day, replaceOptions.newDay) +
            oldParts.sep3;
        /* 曜日ドロップダウンの選択を常に反映（「なし」なら元の suffix を削除） */
        replacedText += replaceOptions.explicitWeekdayText;
        return replacedText;
    }

    /**
     * 選んだ形式で置換後の文字列を作る
     * @param {string} format - FORMAT_VALUES の値（preserve 以外）
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {string|null} 置換後の文字列
     */
    function buildExplicitFormatText(format, replaceOptions) {
        var newYear = replaceOptions.newYear;
        var newMonth = replaceOptions.newMonth;
        var newDay = replaceOptions.newDay;
        var weekdayText = replaceOptions.explicitWeekdayText;
        var reiwaYear = gregorianToReiwa(newYear);
        var heiseiYear = gregorianToHeisei(newYear);
        var reiwaJpYearText = formatEraYearText(reiwaYear, true);
        var heiseiJpYearText = formatEraYearText(heiseiYear, true);
        switch (format) {
            case 'jp': return newYear + "年" + newMonth + "月" + newDay + "日" + weekdayText;
            case 'jp-md': return newMonth + "月" + newDay + "日" + weekdayText;
            case 'dot': return newYear + "." + newMonth + "." + newDay + weekdayText;
            case 'dot-md': return newMonth + "." + newDay + weekdayText;
            case 'slash': return newYear + "/" + newMonth + "/" + newDay + weekdayText;
            case 'slash-md': return newMonth + "/" + newDay + weekdayText;
            case 'reiwa-jp': return "令和" + reiwaJpYearText + "年" + newMonth + "月" + newDay + "日" + weekdayText;
            case 'r-dot': return "R" + reiwaYear + "." + newMonth + "." + newDay + weekdayText;
            case 'r-slash': return "R" + reiwaYear + "/" + newMonth + "/" + newDay + weekdayText;
            case 'heisei-jp': return "平成" + heiseiJpYearText + "年" + newMonth + "月" + newDay + "日" + weekdayText;
            case 'h-dot': return "H" + heiseiYear + "." + newMonth + "." + newDay + weekdayText;
            case 'h-slash': return "H" + heiseiYear + "/" + newMonth + "/" + newDay + weekdayText;
        }
        return null;
    }

    /**
     * 数字の書式を保つため、年・月・日などの部分ごとの置換区間を作る
     * @param {Object} matchInfo - 見つかった日付
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {{offset: number, oldLen: number, newText: string}[]|null} 置換区間（マッチ先頭からの位置）
     */
    function buildReplacementSegments(matchInfo, replaceOptions) {
        var parsedDate = matchInfo.parsed;
        if (!parsedDate) return null;
        var oldParts = parsedDate.parts;
        var segments = [];
        var segmentPos = 0;
        function pushSegment(oldText, newText) {
            if (oldText.length === 0) return;
            segments.push({ offset: segmentPos, oldLen: oldText.length, newText: newText });
            segmentPos += oldText.length;
        }
        var hasJpSuffixSlot = (parsedDate.format === 'jp' || parsedDate.format === 'reiwa-jp' || parsedDate.format === 'heisei-jp');

        pushSegment(oldParts.prefix, oldParts.prefix);
        pushSegment(oldParts.year, buildNewYearText(parsedDate, replaceOptions.newYear));
        pushSegment(oldParts.sep1, oldParts.sep1);
        pushSegment(oldParts.month, applyZeroPaddingFrom(oldParts.month, replaceOptions.newMonth));
        pushSegment(oldParts.sep2, oldParts.sep2);

        if (hasJpSuffixSlot) {
            pushSegment(oldParts.day, applyZeroPaddingFrom(oldParts.day, replaceOptions.newDay));
            if (oldParts.suffix.length > 0) {
                /* 元 suffix を曜日ドロップダウンの選択で置換（「なし」のときは空文字で削除） */
                pushSegment(oldParts.sep3, oldParts.sep3);
                pushSegment(oldParts.suffix, replaceOptions.explicitWeekdayText);
            } else {
                /* 元 suffix なし：sep3（"日"）末尾に曜日を伸ばす */
                pushSegment(oldParts.sep3, oldParts.sep3 + replaceOptions.explicitWeekdayText);
            }
        } else {
            /* dot / slash / r-dot / r-slash：曜日サフィックスを含めて日付全体を1セグメントとして置換する */
            var preservedText = buildPreservedFormatText(matchInfo, replaceOptions);
            if (preservedText !== null) {
                segments = [{ offset: 0, oldLen: matchInfo.text.length, newText: preservedText }];
            }
        }
        return segments;
    }

    /**
     * 1つのテキストフレーム内のマッチを後ろから置換する（途中で例外が出ても、それまでの操作数は replaceStats に残る）
     * @param {TextFrame} targetFrame - 対象のテキストフレーム
     * @param {Object[]} frameMatches - このフレームのマッチ（matchIndex の降順）
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @param {{operations: number, errors: number}} replaceStats - 文字操作数を足していく集計
     * @returns {void}
     */
    function replaceFrameMatches(targetFrame, frameMatches, replaceOptions, replaceStats) {
        for (var orderIdx = 0; orderIdx < frameMatches.length; orderIdx++) {
            var currentMatch = frameMatches[orderIdx];
            if (replaceOptions.formatChoice === 'preserve' && replaceOptions.preserveNumberFormat) {
                var segments = buildReplacementSegments(currentMatch, replaceOptions);
                if (!segments) continue;
                for (var segIdx = segments.length - 1; segIdx >= 0; segIdx--) {
                    var segment = segments[segIdx];
                    replaceStats.operations += replaceMatchPreserveStyle(targetFrame, currentMatch.matchIndex + segment.offset, segment.oldLen, segment.newText);
                }
            } else {
                var fullText = (replaceOptions.formatChoice === 'preserve')
                    ? buildPreservedFormatText(currentMatch, replaceOptions)
                    : buildExplicitFormatText(replaceOptions.formatChoice, replaceOptions);
                if (fullText === null) continue;
                replaceStats.operations += replaceMatchPreserveStyle(targetFrame, currentMatch.matchIndex, currentMatch.text.length, fullText);
            }
        }
    }

    /**
     * チェックされたマッチに対して置換を実行する
     * @param {Object[]} foundMatches - 見つかった日付
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @param {ReplaceOptions} replaceOptions - 置換の設定
     * @returns {{operations: number, errors: number, selectedCount: number}} 文字操作数・エラー件数・対象件数
     */
    function performReplacement(foundMatches, textFrames, replaceOptions) {
        /* チェック済みマッチをフレーム単位でまとめる */
        var matchesByFrame = {};
        var selectedCount = 0;
        for (var resultIdx = 0; resultIdx < foundMatches.length; resultIdx++) {
            if (!foundMatches[resultIdx].checkbox || !foundMatches[resultIdx].checkbox.value) continue;
            var targetFrameIdx = foundMatches[resultIdx].frameIndex;
            if (!matchesByFrame[targetFrameIdx]) matchesByFrame[targetFrameIdx] = [];
            matchesByFrame[targetFrameIdx].push(foundMatches[resultIdx]);
            selectedCount++;
        }

        var replaceStats = { operations: 0, errors: 0 };

        for (var frameIdxKey in matchesByFrame) {
            if (!matchesByFrame.hasOwnProperty(frameIdxKey)) continue;
            var frameMatches = matchesByFrame[frameIdxKey];
            /* 後ろから置換すれば前方の matchIndex はずれない */
            frameMatches.sort(function (firstMatch, secondMatch) { return secondMatch.matchIndex - firstMatch.matchIndex; });

            try {
                replaceFrameMatches(textFrames[parseInt(frameIdxKey, 10)], frameMatches, replaceOptions, replaceStats);
            } catch (eReplace) {
                replaceStats.errors++;
            }
        }

        /*
           チェック済みマッチに紐づく曜日ペア（同一グループ内の曜日のみフレーム）を更新。
           曜日ドロップダウンが「なし」の場合は、空文字化せず連動フレームを更新しない。
           各ペアフレームは独立した TextFrame のままで、日付フレームには連結しない。
        */
        if (replaceOptions.weekdayChoice !== 'none') {
            for (var pairMatchIndex = 0; pairMatchIndex < foundMatches.length; pairMatchIndex++) {
                var dateMatchWithPairs = foundMatches[pairMatchIndex];
                if (!dateMatchWithPairs.checkbox || !dateMatchWithPairs.checkbox.value) continue;
                if (!dateMatchWithPairs.weekdayPairs) continue;
                for (var weekdayPairIndex = 0; weekdayPairIndex < dateMatchWithPairs.weekdayPairs.length; weekdayPairIndex++) {
                    var weekdayPair = dateMatchWithPairs.weekdayPairs[weekdayPairIndex];
                    try {
                        /* 元のスタイルを維持しつつ、新しい曜日に置換 */
                        var newWeekdayText = buildExplicitWeekday(weekdayPair.style, replaceOptions.newYear, replaceOptions.newMonth, replaceOptions.newDay);
                        if (newWeekdayText === "") continue;
                        var oldWeekdayText = String(weekdayPair.frame.contents);
                        if (oldWeekdayText === newWeekdayText) continue;
                        replaceStats.operations += replaceMatchPreserveStyle(weekdayPair.frame, 0, oldWeekdayText.length, newWeekdayText);
                    } catch (eWeekdayPair) {
                        replaceStats.errors++;
                    }
                }
            }
        }

        return { operations: replaceStats.operations, errors: replaceStats.errors, selectedCount: selectedCount };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見つかった日付をアートボードごとにまとめる（アートボード外の -1 は末尾）
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {{artboardIndexes: number[], matchesByArtboard: Object}} 表示順のアートボード番号と、番号ごとのマッチ
     */
    function groupMatchesByArtboard(foundMatches) {
        var matchesByArtboard = {};
        var artboardIndexes = [];
        for (var matchIdx = 0; matchIdx < foundMatches.length; matchIdx++) {
            var artboardKey = String(foundMatches[matchIdx].artboardIndex);
            if (!matchesByArtboard[artboardKey]) {
                matchesByArtboard[artboardKey] = [];
                artboardIndexes.push(foundMatches[matchIdx].artboardIndex);
            }
            matchesByArtboard[artboardKey].push(foundMatches[matchIdx]);
        }
        /* -1（アートボード外）は末尾に */
        artboardIndexes.sort(function (firstIndex, secondIndex) {
            if (firstIndex === -1 && secondIndex !== -1) return 1;
            if (secondIndex === -1 && firstIndex !== -1) return -1;
            return firstIndex - secondIndex;
        });
        return { artboardIndexes: artboardIndexes, matchesByArtboard: matchesByArtboard };
    }

    /**
     * ［見つかった日付］パネルを追加し、日付ごとのチェックボックスを作る（match.checkbox に入れる）
     * @param {Window} dateDialog - 追加先のダイアログ
     * @param {Document} doc - 対象のドキュメント
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {Checkbox[]} 作ったチェックボックス
     */
    function addFoundDatesPanel(dateDialog, doc, foundMatches) {
        var dateCheckboxes = [];
        var artboardGroups = groupMatchesByArtboard(foundMatches);

        var foundDatesPanel = dateDialog.add("panel", undefined, getLabel(LABELS.panel.foundDates, [foundMatches.length]));
        setupPanel(foundDatesPanel, 6);
        foundDatesPanel.alignChildren = "left";
        foundDatesPanel.minimumSize.width = FOUND_PANEL_MIN_WIDTH;

        for (var keyIdx = 0; keyIdx < artboardGroups.artboardIndexes.length; keyIdx++) {
            var currentArtboardIdx = artboardGroups.artboardIndexes[keyIdx];
            var artboardHeaderText;
            if (currentArtboardIdx === -1) {
                artboardHeaderText = getLabel(LABELS.heading.outsideArtboards);
            } else {
                artboardHeaderText = getLabel(LABELS.heading.artboard, [currentArtboardIdx + 1, doc.artboards[currentArtboardIdx].name]);
            }
            /* アートボード見出しは下に少し余白を取る */
            var artboardHeaderGroup = foundDatesPanel.add("group");
            artboardHeaderGroup.orientation = "row";
            artboardHeaderGroup.alignment = "left";
            artboardHeaderGroup.margins = [0, 4, 0, 6];
            artboardHeaderGroup.add("statictext", undefined, artboardHeaderText);

            var artboardMatches = artboardGroups.matchesByArtboard[String(currentArtboardIdx)];
            for (var itemIdx = 0; itemIdx < artboardMatches.length; itemIdx++) {
                /* 各日付行は左マージンでインデント */
                var checkboxRow = foundDatesPanel.add("group");
                checkboxRow.orientation = "row";
                checkboxRow.alignment = "left";
                checkboxRow.margins = [DATE_ROW_INDENT, 0, 0, 0];
                var checkboxText = artboardMatches[itemIdx].text;
                var pairCount = (artboardMatches[itemIdx].weekdayPairs ? artboardMatches[itemIdx].weekdayPairs.length : 0);
                if (pairCount > 0) {
                    checkboxText += getLabel(LABELS.checkbox.linkedWeekdaySuffix);
                }
                var dateCheckbox = checkboxRow.add("checkbox", undefined, checkboxText);
                dateCheckbox.value = true;
                /* チェックボックスのラベル切れ対策 */
                dateCheckbox.preferredSize.width = DATE_CHECKBOX_WIDTH;
                dateCheckbox.helpTip = (pairCount > 0)
                    ? getLabel(LABELS.tooltip.dateCheckboxLinked, [pairCount])
                    : getLabel(LABELS.tooltip.dateCheckbox);
                artboardMatches[itemIdx].checkbox = dateCheckbox;
                dateCheckboxes.push(dateCheckbox);
            }
        }
        return dateCheckboxes;
    }

    /**
     * 年・月・日の入力欄を1つ追加する
     * @param {Group} dateInputGroup - 追加先
     * @param {number} initialValue - 初期値
     * @param {number} fieldChars - 入力欄の幅（文字数）
     * @param {Object} unitLabelSet - 右に添える「年」「月」「日」の LABELS リーフ
     * @param {number} [maxValue] - ∧∨・↑↓キーで止める上限（月は12、日は31。年は省略）
     * @returns {EditText} 追加した入力欄
     */
    function addDateField(dateInputGroup, initialValue, fieldChars, unitLabelSet, maxValue) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = dateInputGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var dateField;
        /* 整数・1以上。値を変えたら onChange（確認表示とプレビューの更新）を呼ぶ
           Whole numbers from 1; run onChange (check labels and preview) after each step */
        var dateStepper = addStepper(stepperInputGroup, function () { return dateField; }, {
            integer: true,
            min: 1,
            max: maxValue,
            onStep: function (numberInput) {
                if (typeof numberInput.onChange === "function") numberInput.onChange();
            }
        });
        dateField = stepperInputGroup.add("edittext", undefined, String(initialValue));
        bindSteppedArrowKeys(dateField, dateStepper);
        dateField.characters = fieldChars;
        dateField.helpTip = getLabel(LABELS.tooltip.dateField);
        dateInputGroup.add("statictext", undefined, getLabel(unitLabelSet));
        return dateField;
    }

    /**
     * 確認パネルに「項目名＋値」の行を追加する
     * @param {Panel} checkPanel - 追加先
     * @param {Object} rowLabelSet - 項目名の LABELS リーフ
     * @param {number} valueWidth - 値の欄の幅
     * @returns {StaticText} 値の欄
     */
    function addCheckRow(checkPanel, rowLabelSet, valueWidth) {
        var checkRow = checkPanel.add("group");
        checkRow.orientation = "row";
        var rowLabel = checkRow.add("statictext", undefined, labelText(rowLabelSet));
        rowLabel.preferredSize.width = CHECK_LABEL_WIDTH;
        var valueText = checkRow.add("statictext", undefined, "");
        valueText.preferredSize.width = valueWidth;
        return valueText;
    }

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

    /**
     * ダイアログを組み立てる（イベントはまだ付けない）
     * @param {Document} doc - 対象のドキュメント
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {Object} ダイアログと各コントロール
     */
    function buildDialog(doc, foundMatches) {
        var dateDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dateDialog);

        var dateCheckboxes = [];
        if (foundMatches.length === 0) {
            dateDialog.add("statictext", undefined, getLabel(LABELS.message.noDatesFound));
        } else {
            dateCheckboxes = addFoundDatesPanel(dateDialog, doc, foundMatches);
        }

        /* 置換後の日付入力パネル：［2026］年［5］月［8］日 形式。初期値は今日 */
        var today = new Date();
        var replacementPanel = dateDialog.add("panel", undefined, getLabel(LABELS.panel.replacement));
        setupPanel(replacementPanel, 6);
        replacementPanel.alignChildren = "left";

        var dateInputGroup = replacementPanel.add("group");
        dateInputGroup.orientation = "row";
        var yearInput = addDateField(dateInputGroup, today.getFullYear(), 5, LABELS.unitLabel.year);
        var monthInput = addDateField(dateInputGroup, today.getMonth() + 1, 3, LABELS.unitLabel.month, 12);
        var dayInput = addDateField(dateInputGroup, today.getDate(), 3, LABELS.unitLabel.day, 31);

        var formatRow = replacementPanel.add("group");
        formatRow.orientation = "row";
        formatRow.add("statictext", undefined, labelText(LABELS.fieldLabel.format));
        var formatDropdown = formatRow.add("dropdownlist", undefined, [getLabel(LABELS.dropdown.preserveFormat)].concat(FORMAT_PATTERN_LABELS));
        formatDropdown.selection = 0;
        formatDropdown.helpTip = getLabel(LABELS.tooltip.formatDropdown);

        var weekdayDropdown = formatRow.add("dropdownlist", undefined, [getLabel(LABELS.dropdown.weekdayNone)].concat(WEEKDAY_SAMPLE_LABELS));
        weekdayDropdown.helpTip = getLabel(LABELS.tooltip.weekdayDropdown);

        /* 最初に見つかった置換対象を基準に、曜日サフィックスまたは曜日のみフレームの形式を初期選択にする */
        var initialWeekdayChoice = detectInitialWeekdayChoice(foundMatches);
        var initialWeekdayIndex = 0;
        for (var weekdayValueIdx = 0; weekdayValueIdx < WEEKDAY_VALUES.length; weekdayValueIdx++) {
            if (WEEKDAY_VALUES[weekdayValueIdx] === initialWeekdayChoice) { initialWeekdayIndex = weekdayValueIdx; break; }
        }
        weekdayDropdown.selection = initialWeekdayIndex;

        /* 数字（年・月・日）の文字書式を保持するか。OFF にするとマッチ範囲を一括置換し、
           新しい文字はマッチ先頭の書式に統一される */
        var preserveNumberFormatCheckbox = replacementPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.preserveNumberFormat));
        preserveNumberFormatCheckbox.value = true;
        preserveNumberFormatCheckbox.helpTip = getLabel(LABELS.tooltip.preserveNumberFormat);

        /* プレビュー：ON で現在の入力をドキュメントに反映し、ダイアログを開いたまま結果を確認できる。
           入力変更時には自動で更新（巻き戻し→再適用）。OFF にすると元に戻す */
        var previewCheckbox = replacementPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.value = false;
        previewCheckbox.helpTip = getLabel(LABELS.tooltip.preview);

        /* 確認パネル：和暦・曜日・日数差のプレビュー（置換対象外） */
        var checkPanel = dateDialog.add("panel", undefined, getLabel(LABELS.panel.check));
        setupPanel(checkPanel, 6);
        checkPanel.alignChildren = "left";
        var eraLabel = addCheckRow(checkPanel, LABELS.fieldLabel.era, 120);
        var weekdayLabel = addCheckRow(checkPanel, LABELS.fieldLabel.weekday, 80);
        var daysDiffLabel = addCheckRow(checkPanel, LABELS.fieldLabel.daysDiff, 120);

        /* OK / キャンセル ボタン */
        var buttonRow = addButtonRow(dateDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            dialog: dateDialog,
            dateCheckboxes: dateCheckboxes,
            yearInput: yearInput,
            monthInput: monthInput,
            dayInput: dayInput,
            formatDropdown: formatDropdown,
            weekdayDropdown: weekdayDropdown,
            preserveNumberFormatCheckbox: preserveNumberFormatCheckbox,
            previewCheckbox: previewCheckbox,
            eraLabel: eraLabel,
            weekdayLabel: weekdayLabel,
            daysDiffLabel: daysDiffLabel,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 年・月・日の入力欄を整数として読む（数値でなければ NaN）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {{year: number, month: number, day: number}} 入力された日付
     */
    function readDateInputs(dialogControls) {
        return {
            year: parseInt(dialogControls.yearInput.text, 10),
            month: parseInt(dialogControls.monthInput.text, 10),
            day: parseInt(dialogControls.dayInput.text, 10)
        };
    }

    /**
     * ダイアログの入力から置換の設定を読み取る
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {ReplaceOptions} 置換の設定
     */
    function readReplaceOptions(dialogControls) {
        var inputDate = readDateInputs(dialogControls);
        var weekdayChoice = getDropdownValue(dialogControls.weekdayDropdown, WEEKDAY_VALUES, 'none');
        return {
            newYear: inputDate.year,
            newMonth: inputDate.month,
            newDay: inputDate.day,
            formatChoice: getDropdownValue(dialogControls.formatDropdown, FORMAT_VALUES, 'preserve'),
            weekdayChoice: weekdayChoice,
            preserveNumberFormat: dialogControls.preserveNumberFormatCheckbox.value,
            explicitWeekdayText: buildExplicitWeekday(weekdayChoice, inputDate.year, inputDate.month, inputDate.day)
        };
    }

    /**
     * 入力値を検証する
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Object[]} foundMatches - 見つかった日付
     * @returns {string|null} エラーがあればメッセージ、なければ null
     */
    function getValidationErrorMessage(dialogControls, foundMatches) {
        var inputDate = readDateInputs(dialogControls);
        var year = inputDate.year;
        var month = inputDate.month;
        var day = inputDate.day;
        if (isNaN(year) || isNaN(month) || isNaN(day)) return getLabel(LABELS.alert.notNumber);
        if (year < 1) return getLabel(LABELS.alert.yearRange);
        if (month < 1 || month > 12) return getLabel(LABELS.alert.monthRange);
        if (day < 1 || day > 31) return getLabel(LABELS.alert.dayRange);
        if (!isRealDate(year, month, day)) return getLabel(LABELS.alert.notRealDate);
        var checkDate = new Date(year, month - 1, day);

        /* 元号フォーマット選択時は、各元号の有効範囲に収まることを要求 */
        var formatChoice = getDropdownValue(dialogControls.formatDropdown, FORMAT_VALUES, 'preserve');
        var isReiwaFormat = (formatChoice === 'reiwa-jp' || formatChoice === 'r-dot' || formatChoice === 'r-slash');
        var isHeiseiFormat = (formatChoice === 'heisei-jp' || formatChoice === 'h-dot' || formatChoice === 'h-slash');
        if (isReiwaFormat && checkDate < REIWA_START_DATE) {
            return getLabel(LABELS.alert.reiwaFormatRange);
        }
        if (isHeiseiFormat && (checkDate < HEISEI_START_DATE || checkDate >= HEISEI_END_DATE_EXCLUSIVE)) {
            return getLabel(LABELS.alert.heiseiFormatRange);
        }

        /* 「元の形式を保持」では、チェック済みマッチの元号がそれぞれ有効になる範囲を要求 */
        if (formatChoice === 'preserve') {
            var hasReiwaMatch = false, hasHeiseiMatch = false;
            for (var fmIdx = 0; fmIdx < foundMatches.length; fmIdx++) {
                var foundMatch = foundMatches[fmIdx];
                if (foundMatch.checkbox && !foundMatch.checkbox.value) continue;
                if (!foundMatch.parsed) continue;
                if (foundMatch.parsed.era === 'reiwa') hasReiwaMatch = true;
                if (foundMatch.parsed.era === 'heisei') hasHeiseiMatch = true;
            }
            if (hasReiwaMatch && checkDate < REIWA_START_DATE) {
                return getLabel(LABELS.alert.reiwaMatchRange);
            }
            if (hasHeiseiMatch && (checkDate < HEISEI_START_DATE || checkDate >= HEISEI_END_DATE_EXCLUSIVE)) {
                return getLabel(LABELS.alert.heiseiMatchRange);
            }
        }

        return null;
    }

    /**
     * 確認パネル（和暦・曜日・最初の日付との差）を入力値で更新する
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Date|null} referenceDate - 比較の基準日（最初に見つかった日付）
     * @returns {void}
     */
    function updateCheckPanel(dialogControls, referenceDate) {
        var inputDate = readDateInputs(dialogControls);
        var year = inputDate.year;
        var month = inputDate.month;
        var day = inputDate.day;

        dialogControls.eraLabel.text = formatEraLabel(year, month, day);
        dialogControls.weekdayLabel.text = getWeekdayLabel(year, month, day);

        dialogControls.daysDiffLabel.text = "";
        if (referenceDate && !isNaN(year) && !isNaN(month) && !isNaN(day)) {
            var newDate = new Date(year, month - 1, day);
            if (!isNaN(newDate.getTime())) {
                dialogControls.daysDiffLabel.text = formatDaysDifference(getDaysDifference(referenceDate, newDate));
            }
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューの適用・巻き戻しを受け持つ（巻き戻しは適用した文字操作数だけ app.undo()）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Object[]} foundMatches - 見つかった日付
     * @param {TextFrame[]} textFrames - 検索対象のテキストフレーム
     * @returns {{apply: Function, revert: Function, refresh: Function}} プレビューの操作
     */
    function createPreviewController(dialogControls, foundMatches, textFrames) {
        var isPreviewApplied = false;
        var previewOpCount = 0;

        /* プレビューを巻き戻す / Undo the preview */
        function revertPreview() {
            if (!isPreviewApplied) return;
            for (var i = 0; i < previewOpCount; i++) {
                try { app.undo(); } catch (eUndo) { break; }
            }
            isPreviewApplied = false;
            previewOpCount = 0;
            app.redraw();
        }

        /* プレビューを適用（呼び出し前に入力が有効であることを確認しておく）/ Apply the preview (input must be valid) */
        function applyPreview() {
            var replaceResult = performReplacement(foundMatches, textFrames, readReplaceOptions(dialogControls));
            previewOpCount = replaceResult.operations;
            isPreviewApplied = replaceResult.operations > 0;
            app.redraw();
        }

        /* 入力やチェック変更時の自動更新（プレビュー ON 時のみ巻き戻し→再適用）/ Re-apply while the preview is on */
        function refreshPreview() {
            if (!dialogControls.previewCheckbox.value) return;
            revertPreview();
            if (getValidationErrorMessage(dialogControls, foundMatches) !== null) return;
            applyPreview();
        }

        return {
            apply: applyPreview,
            revert: revertPreview,
            refresh: refreshPreview
        };
    }

    /**
     * ダイアログにイベントを付け、確認パネルなどの初期状態を反映する
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Object[]} foundMatches - 見つかった日付
     * @param {Object} previewController - createPreviewController() の戻り値
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, foundMatches, previewController) {
        var dateDialog = dialogControls.dialog;
        var dateCheckboxes = dialogControls.dateCheckboxes;

        /* 比較用の基準日（最初に見つかった日付。形式に関わらず西暦で扱う） */
        var referenceDate = null;
        if (foundMatches.length > 0 && foundMatches[0].parsed) {
            var referenceParsed = foundMatches[0].parsed;
            referenceDate = new Date(referenceParsed.year, referenceParsed.month - 1, referenceParsed.day);
        }

        /* チェックボックスに Option＋クリックで全切替を割り当てる。クリック後はプレビューも更新。
           click イベントは発火しない環境があるため、onClick と keyboardState で判定する
           Option-click toggles all; use onClick + keyboardState because the click event may not fire */
        function bindOptionClickToggleAll(checkboxControl) {
            checkboxControl.onClick = function () {
                if (ScriptUI.environment.keyboardState.altKey) {
                    var newValue = checkboxControl.value;
                    for (var idx = 0; idx < dateCheckboxes.length; idx++) {
                        dateCheckboxes[idx].value = newValue;
                    }
                }
                previewController.refresh();
            };
        }
        for (var i = 0; i < dateCheckboxes.length; i++) {
            bindOptionClickToggleAll(dateCheckboxes[i]);
        }

        /* 入力値からチェック用の表示を更新。プレビュー ON 時は再適用も走らせる */
        function refreshCheckLabels() {
            updateCheckPanel(dialogControls, referenceDate);
            previewController.refresh();
        }

        var dateInputs = [dialogControls.yearInput, dialogControls.monthInput, dialogControls.dayInput];
        for (var j = 0; j < dateInputs.length; j++) {
            dateInputs[j].onChange = refreshCheckLabels;
        }

        refreshCheckLabels();

        /* 「数字の書式を保持」は preserve 選択時のみ有効。フォーマット切替で enabled を連動させる */
        function updatePreserveNumberFormatEnabled() {
            var formatChoice = getDropdownValue(dialogControls.formatDropdown, FORMAT_VALUES, 'preserve');
            dialogControls.preserveNumberFormatCheckbox.enabled = (formatChoice === 'preserve');
        }
        updatePreserveNumberFormatEnabled();

        /* フォーマット選択 / 曜日 / 数字書式保持の変更でもプレビューを更新 */
        dialogControls.formatDropdown.onChange = function () {
            updatePreserveNumberFormatEnabled();
            previewController.refresh();
        };
        dialogControls.weekdayDropdown.onChange = function () { previewController.refresh(); };
        dialogControls.preserveNumberFormatCheckbox.onClick = function () { previewController.refresh(); };

        /* プレビューチェックボックス本体 */
        dialogControls.previewCheckbox.onClick = function () {
            if (dialogControls.previewCheckbox.value) {
                var validationError = getValidationErrorMessage(dialogControls, foundMatches);
                if (validationError) {
                    alert(validationError);
                    dialogControls.previewCheckbox.value = false;
                    return;
                }
                previewController.apply();
            } else {
                previewController.revert();
            }
        };

        dialogControls.btnOK.onClick = function () {
            var validationError = getValidationErrorMessage(dialogControls, foundMatches);
            if (validationError) {
                alert(validationError);
                return;
            }
            dateDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            previewController.revert();
            dateDialog.close(0);
        };
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "日付を検索・置換", en: "Find and Replace Dates" }
        },
        panel: {
            foundDates: { ja: "見つかった日付（%1）", en: "Dates Found (%1)" },
            replacement: { ja: "置換後の日付", en: "New Date" },
            check: { ja: "置換内容の確認", en: "Review" }
        },
        heading: {
            outsideArtboards: { ja: "[アートボード外]", en: "[Outside Artboards]" },
            artboard: { ja: "[アートボード %1：%2]", en: "[Artboard %1: %2]" }
        },
        message: {
            noDatesFound: { ja: "日付は見つかりませんでした。", en: "No dates were found." }
        },
        dropdown: {
            preserveFormat: { ja: "元の形式を保持", en: "Keep Original Format" },
            weekdayNone: { ja: "なし", en: "None" }
        },
        checkbox: {
            preserveNumberFormat: { ja: "数字の書式を保持する", en: "Keep Number Formatting" },
            preview: { ja: "プレビュー", en: "Preview" },
            linkedWeekdaySuffix: { ja: "  ＋曜日連動", en: "  + linked weekday" }
        },
        fieldLabel: {
            format: { ja: "フォーマット", en: "Format" },
            era: { ja: "和暦", en: "Japanese era" },
            weekday: { ja: "置換後の曜日", en: "New weekday" },
            daysDiff: { ja: "最初の日付との差", en: "Diff. from first date" }
        },
        unitLabel: {
            year: { ja: "年", en: "Y" },
            month: { ja: "月", en: "M" },
            day: { ja: "日", en: "D" }
        },
        valueText: {
            daysSuffix: { ja: "日", en: " days" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            dateField: { ja: "↑↓キーで増減（Shift：±10）", en: "Up/Down arrow keys to change (Shift: ±10)" },
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
            dateCheckbox: { ja: "Option＋クリックで全項目を一括切替", en: "Option-click to toggle all items" },
            dateCheckboxLinked: {
                ja: "Option＋クリックで全項目を一括切替\n同一グループ内の曜日フレーム %1 件も連動して更新されます\n曜日が「なし」の場合は連動フレームを更新しません",
                en: "Option-click to toggle all items\n%1 weekday frame(s) in the same group are updated too\nThey are not updated when the weekday is set to None"
            },
            formatDropdown: {
                ja: "「元の形式を保持」：各マッチの区切り文字・元号を維持\n曜日表記は右の曜日ドロップダウンの選択を反映\nそれ以外：すべてのマッチを選択した形式に統一",
                en: "Keep Original Format: keeps each match's separators and era\nThe weekday follows the weekday menu on the right\nOther formats: rewrite every match in the chosen format"
            },
            weekdayDropdown: {
                ja: "出力に付与する曜日表記。「元の形式を保持」選択時もこの選択を反映します。\n「なし」を選ぶと日付内の曜日表記を削除し、連動する曜日フレームは更新しません",
                en: "Weekday style added to the output, also with Keep Original Format.\nNone removes the weekday from the date and leaves linked weekday frames unchanged"
            },
            preserveNumberFormat: {
                ja: "ON：年・月・日それぞれの元の文字書式（フォント・サイズ・色など）を維持\nOFF：マッチ範囲全体を一括置換し、書式は先頭文字に揃える\n※「元の形式を保持」選択時のみ有効",
                en: "On: keeps the original character formatting (font, size, color, etc.) of the year, month and day\nOff: replaces the whole match at once, using the formatting of its first character\nAvailable only with Keep Original Format"
            },
            preview: {
                ja: "ON でドキュメントに即時反映。入力やチェックを変更すると自動更新。\nOFF・キャンセルで元に戻す",
                en: "When on, changes are applied to the document right away and follow your edits.\nTurning it off or canceling reverts them"
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noDates: { ja: "置換対象の日付がありません。", en: "There are no dates to replace." },
            noneChecked: { ja: "置換する項目が選択されていません。", en: "No items are selected for replacement." },
            lockedErrors: {
                ja: "%1件はロック等の理由で処理できませんでした。",
                en: "%1 item(s) could not be processed because they are locked or otherwise unavailable."
            },
            notNumber: { ja: "年・月・日は半角数字で入力してください。", en: "Enter the year, month and day as numbers." },
            yearRange: { ja: "年は1以上で入力してください。", en: "Enter a year of 1 or later." },
            monthRange: { ja: "月は1〜12で入力してください。", en: "Enter a month from 1 to 12." },
            dayRange: { ja: "日は1〜31で入力してください。", en: "Enter a day from 1 to 31." },
            notRealDate: { ja: "存在しない日付です。", en: "This date does not exist." },
            reiwaFormatRange: {
                ja: "令和形式は 2019/5/1 以降の日付で指定してください。",
                en: "Reiwa formats need a date on or after 2019/5/1."
            },
            heiseiFormatRange: {
                ja: "平成形式は 1989/1/8〜2019/4/30 の範囲で指定してください。",
                en: "Heisei formats need a date from 1989/1/8 to 2019/4/30."
            },
            reiwaMatchRange: {
                ja: "令和形式のマッチが含まれているため、2019/5/1 以降の日付を指定してください。",
                en: "Some matches use a Reiwa format, so enter a date on or after 2019/5/1."
            },
            heiseiMatchRange: {
                ja: "平成形式のマッチが含まれているため、1989/1/8〜2019/4/30 の日付を指定してください。",
                en: "Some matches use a Heisei format, so enter a date from 1989/1/8 to 2019/4/30."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 日付を探してダイアログを表示し、チェックした日付を置換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var textFrames = collectTargetTextFrames(doc);
        var foundMatches = findDateMatches(doc, textFrames);
        pairWeekdayFrames(foundMatches, textFrames);

        var dialogControls = buildDialog(doc, foundMatches);
        var previewController = createPreviewController(dialogControls, foundMatches, textFrames);
        bindDialogEvents(dialogControls, foundMatches, previewController);

        prepareDialogWindow(dialogControls.dialog, SCRIPT_NAME);
        if (dialogControls.dialog.show() !== 1) {
            return;
        }

        if (foundMatches.length === 0) {
            alert(getLabel(LABELS.alert.noDates));
            return;
        }

        /* OK時はプレビュー結果をそのまま確定せず、いったん戻してから本番置換を再実行する */
        previewController.revert();

        var replaceResult = performReplacement(foundMatches, textFrames, readReplaceOptions(dialogControls));

        if (replaceResult.selectedCount === 0) {
            alert(getLabel(LABELS.alert.noneChecked));
            return;
        }

        if (replaceResult.errors > 0) {
            alert(formatErrorMessage(replaceResult.errors));
        }
    }

    main();

})();
