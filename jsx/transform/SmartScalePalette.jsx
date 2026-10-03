#target illustrator
#targetengine "SmartScalePaletteEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトを、％または仕上がりの幅・高さで拡大・縮小する常駐パレットです。プレビューで確かめてから確定でき、個別に拡大・縮小しながら選択全体の端を固定することもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartScalePalette.md

### Overview

A persistent palette that scales the selected objects by percentage or to a finished width and height. Check the preview before committing, and scale objects individually while the edges of the whole selection stay put.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartScalePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartScalePalette";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartScalePalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartScalePalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* すでにパレットが開いていれば前面に出して終了 / If the palette is already open, bring it forward and return */
    try {
        if ($.global.__smartScalePalette) {
            $.global.__smartScalePalette.show();
            $.global.__smartScalePalette.active = true;
            return;
        }
    } catch (staleReferenceError) {
        $.global.__smartScalePalette = null; /* 参照が無効なら作り直す / stale reference: rebuild */
    }

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var DEFAULT_SCALE_PERCENT = 100; /* スケールの初期値（%） / initial scale (%) */
    var MIN_SCALE_PERCENT = 0.1;     /* 受け付けるスケールの下限（%） / smallest scale accepted (%) */
    var DEFAULT_ANCHOR = "center";   /* 基準点の初期値 / initial reference point */
    var MIN_TARGET_LENGTH = 0.01;    /* ∧∨で下げられる幅・高さの下限（定規の単位） / smallest width or height the steppers reach (ruler units) */

    /* スケールのボタン（右カラムに上から並ぶ）。setTo は値そのもの、stepBy は今の値への加減（%）。
       内側の配列が1つのまとまりで、±10% のまとまりだけ上下に余白を取る
       Scale buttons in the right column, top to bottom. setTo sets the value, stepBy adds to it (%).
       Each inner array is a block; only the ±10% block gets top and bottom margins */
    var SCALE_PRESET_BUTTONS = [
        [{ setTo: 50 }, { setTo: 100 }, { setTo: 200 }],
        [{ stepBy: 10 }, { stepBy: -10 }],
        [{ stepBy: 1 }, { stepBy: -1 }]
    ];

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

    var SCALE_FIELD_CHARACTERS = 4;       /* スケールの入力欄の幅（文字数） / width of the scale field in characters */
    var SCALE_PRESET_BUTTON_WIDTH = 60;   /* スケールのボタンの幅 / width of a scale button */
    var PRESET_BUTTON_SPACING = 4;        /* スケールのボタンどうしの間隔 / spacing between scale buttons */
    var PRESET_STEP10_MARGIN = 10;        /* ±10% のまとまりの上下の余白 / top and bottom margin of the ±10% block */
    var RELATIVE_STEP_TOP_MARGIN = 10;    /* ［相対］の上の余白 / top margin of the Relative checkbox */
    var SIZE_LABEL_WIDTH = 64;            /* スケール・幅・高さの項目名の幅 / width of the scale, width and height labels */
    var SIZE_FIELD_CHARACTERS = 6;        /* 幅・高さの入力欄の幅（文字数） / width of the width and height fields in characters */
    var CURRENT_SIZE_WIDTH = 64;          /* 今の幅・高さの表示の幅 / width of the current size text */
    var PIN_OPTIONS_INDENT = 18;          /* ［個別に］のオプションの字下げ / indent of the Each Object options */

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

    /**
     * 数値を指定の桁で丸めて文字列にする（末尾の0は付けない）
     * @param {number} value - 値
     * @param {number} decimalDigits - 小数の桁数
     * @returns {string} 丸めた値
     */
    function formatRounded(value, decimalDigits) {
        var roundingFactor = Math.pow(10, decimalDigits);
        return String(Math.round(value * roundingFactor) / roundingFactor);
    }

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

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        dialog: {
            title: { ja: "スマート拡大・縮小", en: "Smart Scale" }
        },
        fieldLabel: {
            scale: { ja: "スケール", en: "Scale" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            sizeArrow: { ja: "→", en: "→" },
            noSize: { ja: "—", en: "—" },
            percentUnit: { ja: "%", en: "%" }
        },
        panel: {
            size: { ja: "サイズ", en: "Size" },
            mode: { ja: "モード", en: "Mode" },
            anchor: { ja: "基準点", en: "Reference Point" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            perItem: { ja: "個別に", en: "Each Object" },
            asGroup: { ja: "全体で", en: "As a Whole" }
        },
        checkbox: {
            scaleCorners: { ja: "角", en: "Corners" },
            strokeWidth: { ja: "線幅と効果", en: "Strokes & Effects" },
            pattern: { ja: "パターン", en: "Patterns" },
            gradient: { ja: "グラデーション", en: "Gradients" },
            preview: { ja: "プレビュー", en: "Preview" },
            pinHorizontal: { ja: "左右の両端", en: "Left & Right Edges" },
            pinVertical: { ja: "上下の両端", en: "Top & Bottom Edges" },
            relativeStep: { ja: "相対", en: "Relative" }
        },
        tooltip: {
            scale: {
                ja: "拡大・縮小率です。100 で原寸。左の∧∨や↑↓キーで増減できます。幅と高さの倍率が違うときは空欄になります。",
                en: "Scale percentage. 100 is the original size. The steppers on the left and the arrow keys step the value. Empty while the width and height scales differ."
            },
            anchor: {
                ja: "拡大・縮小の基準点です。［個別に］では各オブジェクトの、［全体で］では選択全体の外形に対して決まります。［左右の両端］［上下の両端］を固定した向きには使いません。",
                en: "Reference point for scaling, on each object's bounds with Each Object and on the whole selection's bounds with As a Whole. Not used in a direction pinned by Left & Right Edges or Top & Bottom Edges."
            },
            sizeField: {
                ja: "選択全体をこの大きさにするスケールを求めます。［個別に］でも、倍率は選択全体の大きさから求めます。",
                en: "Finds the scale that brings the whole selection to this size. With Each Object too, the scale comes from the whole selection."
            },
            currentWidth: {
                ja: "選択全体の今の幅です（環境設定［プレビュー境界を使用］に従って測ります）。パレットに戻ったときに測り直します。",
                en: "Current width of the whole selection (measured according to the Use Preview Bounds preference). Measured again when you return to the palette."
            },
            currentHeight: {
                ja: "選択全体の今の高さです（環境設定［プレビュー境界を使用］に従って測ります）。パレットに戻ったときに測り直します。",
                en: "Current height of the whole selection (measured according to the Use Preview Bounds preference). Measured again when you return to the palette."
            },
            presetSetTo: { ja: "スケールを{value}%にします", en: "Sets the scale to {value}%" },
            presetStepUp: {
                ja: "スケールを{value}%増やします。［相対］が OFF なら元の大きさに対して足し（100→110→120）、ON なら今の倍率に掛けます（100→110→121）",
                en: "Increases the scale by {value}%. With Relative off it adds to the original size (100→110→120); with it on it multiplies the current scale (100→110→121)"
            },
            presetStepDown: {
                ja: "スケールを{value}%減らします。［相対］が OFF なら元の大きさに対して引き（100→90→80）、ON なら今の倍率に掛けます（100→90→81）",
                en: "Decreases the scale by {value}%. With Relative off it subtracts from the original size (100→90→80); with it on it multiplies the current scale (100→90→81)"
            },
            relativeStep: {
                ja: "ON にすると、±のボタンで今の倍率に掛けて増減します（-10% を2回で 100→90→81）。OFF なら元の大きさに対して足し引きします（100→90→80）",
                en: "When on, the ± buttons multiply the current scale (-10% twice: 100→90→81). When off, they add to or subtract from the original size (100→90→80)"
            },
            scaleCorners: { ja: "角丸の半径も拡大・縮小します", en: "Also scales the radius of rounded corners" },
            strokeWidth: {
                ja: "線幅と、ドロップシャドウなどの効果の値も拡大・縮小します",
                en: "Also scales stroke widths and effect values such as drop shadows"
            },
            pattern: { ja: "塗りや線のパターンも拡大・縮小します", en: "Also scales patterns in fills and strokes" },
            gradient: { ja: "グラデーションも拡大・縮小します", en: "Also scales gradients" },
            preview: {
                ja: "値を変えたりボタンを押したりするたびに、結果をその場で表示します（［適用］するまで確定しません）",
                en: "Shows the result as you change values or click buttons (nothing is committed until you click Apply)"
            },
            reset: { ja: "スケールを100%に戻し、すべての設定を初期値に戻します", en: "Returns the scale to 100% and every setting to its default" },
            apply: { ja: "選択中のオブジェクトを拡大・縮小します", en: "Scales the selected objects" },
            close: { ja: "パレットを閉じます（Esc）", en: "Closes the palette (Esc)" },
            commit: { ja: "プレビューの結果で確定し、パレットを閉じます", en: "Commits the previewed result and closes the palette" },
            linkToggle: {
                ja: "幅と高さの比率を保つ（クリックで切り替え）。OFF のときは幅と高さを別々に拡大・縮小します。",
                en: "Keep the width-to-height ratio (click to toggle). When off, width and height scale separately."
            },
            perItem: { ja: "オブジェクトごとに、それぞれの基準点で拡大・縮小します。", en: "Scales each object around its own reference point." },
            asGroup: { ja: "選択全体を1つとみなして、全体の基準点で拡大・縮小します。", en: "Scales the whole selection as one, around its reference point." },
            pinHorizontal: {
                ja: "［個別に］で、選択全体の左右の端を動かしません。端にあるものはその端を基準にし、間のものは位置に応じて基準を決めます（基準点の列は使いません）。",
                en: "With Each Object, keeps the left and right edges of the whole selection. Objects at an edge anchor there, and those in between by their position (the anchor column is not used)."
            },
            pinVertical: {
                ja: "［個別に］で、選択全体の上下の端を動かしません。端にあるものはその端を基準にし、間のものは位置に応じて基準を決めます（基準点の行は使いません）。",
                en: "With Each Object, keeps the top and bottom edges of the whole selection. Objects at an edge anchor there, and those in between by their position (the anchor row is not used)."
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
            reset: { ja: "リセット", en: "Reset" },
            close: { ja: "閉じる", en: "Close" },
            apply: { ja: "適用", en: "Apply" },
            commit: { ja: "確定", en: "Commit" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectObjects: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            error: { ja: "エラーが発生しました：{message}", en: "An error has occurred: {message}" }
        }
    };

    // =========================================
    // worker（メインエンジンで実行する DOM 処理） / Worker (DOM work run on the main engine)
    // 常駐パレットのエンジンからは DOM に触れないので、toString() した関数を BridgeTalk で送る。
    // 注意: JSDoc・行コメント（//）・本体内のコメント禁止、各文はセミコロンで終える（toString が改行を消すことがあるため）
    // The palette engine cannot touch the DOM, so these functions are sent through BridgeTalk as source.
    // Note: no JSDoc, no // comments and no comments inside the body; end every statement with ';'
    // =========================================

    /* 選択からページアイテムを集める（テキスト編集中などは空） / Collect page items from the selection (empty while editing text) */
    function workerCollectItems() {
        var docSelection = app.activeDocument.selection;
        var pageItems = [];
        if (!(docSelection instanceof Array)) { return pageItems; };
        for (var i = 0; i < docSelection.length; i++) {
            if (docSelection[i].geometricBounds) { pageItems.push(docSelection[i]); };
        };
        return pageItems;
    };

    /* オブジェクト群を囲む範囲 [左, 上, 右, 下] / Union bounds [left, top, right, bottom] */
    function workerUnionBounds(pageItems, useVisibleBounds) {
        var unionBounds = null;
        for (var i = 0; i < pageItems.length; i++) {
            var itemBounds = useVisibleBounds ? pageItems[i].visibleBounds : pageItems[i].geometricBounds;
            if (!unionBounds) { unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]]; continue; };
            if (itemBounds[0] < unionBounds[0]) { unionBounds[0] = itemBounds[0]; };
            if (itemBounds[1] > unionBounds[1]) { unionBounds[1] = itemBounds[1]; };
            if (itemBounds[2] > unionBounds[2]) { unionBounds[2] = itemBounds[2]; };
            if (itemBounds[3] < unionBounds[3]) { unionBounds[3] = itemBounds[3]; };
        };
        return unionBounds;
    };

    /* 基準点（0〜8、行優先）を範囲の中の割合 [横, 縦]（0〜1）にする / Reference point (0-8, row-major) as ratios [x, y] (0-1) within bounds */
    function workerAnchorRatio(anchorIndex) {
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    };

    /* 範囲の中で、割合 [横, 縦] にあたる点の座標 / Coordinates of the point at ratios [x, y] within bounds */
    function workerAnchorPoint(areaBounds, anchorRatio) {
        return [areaBounds[0] + (areaBounds[2] - areaBounds[0]) * anchorRatio[0], areaBounds[1] + (areaBounds[3] - areaBounds[1]) * anchorRatio[1]];
    };

    /* 全体の外形を保つための、オブジェクトごとの基準の割合 [横, 縦]。全体の余白のうち、左側（上側）にある割合を使う
       （端にあるものはその端が基準になる）。全体と同じ幅（高さ）なら中央
       Anchor ratios [x, y] that keep the overall bounds: the share of the free space on the item's left (top side);
       an item at an edge anchors there, and one spanning the whole width (height) is centered */
    function workerKeepBoundsRatio(itemBounds, unionBounds) {
        var freeWidth = (unionBounds[2] - unionBounds[0]) - (itemBounds[2] - itemBounds[0]);
        var freeHeight = (unionBounds[1] - unionBounds[3]) - (itemBounds[1] - itemBounds[3]);
        var ratioX = (freeWidth > 0) ? Math.min(1, Math.max(0, (itemBounds[0] - unionBounds[0]) / freeWidth)) : 0.5;
        var ratioY = (freeHeight > 0) ? Math.min(1, Math.max(0, (unionBounds[1] - itemBounds[1]) / freeHeight)) : 0.5;
        return [ratioX, ratioY];
    };

    /* オブジェクト群を1つとみなし、基準点を動かさずに拡大・縮小する。各オブジェクトを中心で拡大・縮小してから
       基準点からの距離が倍率どおりになるよう動かし、線幅などで生じたずれを最後に基準点へ合わせ直す。
       changeLineWidths は真偽値ではなく線幅の倍率（%）。縦横で違うときは相乗平均
       Scale the items as one around the reference point: scale each at its center, move it so its distance from the
       reference point follows the scale, then realign the reference point. changeLineWidths is a percentage, not a boolean */
    function workerScaleAsOne(pageItems, scaleSettings, anchorRatio) {
        var lineWidthPercent = scaleSettings.scaleStrokes ? Math.sqrt(scaleSettings.scaleX * scaleSettings.scaleY) : 100;
        var anchorBefore = workerAnchorPoint(workerUnionBounds(pageItems, scaleSettings.useVisibleBounds), anchorRatio);
        for (var i = 0; i < pageItems.length; i++) {
            var geometricBounds = pageItems[i].geometricBounds;
            var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
            var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;
            pageItems[i].resize(scaleSettings.scaleX, scaleSettings.scaleY, true, scaleSettings.scalePatterns, scaleSettings.scaleGradients, scaleSettings.scalePatterns, lineWidthPercent, Transformation.CENTER);
            pageItems[i].translate((centerX - anchorBefore[0]) * (scaleSettings.scaleX / 100 - 1), (centerY - anchorBefore[1]) * (scaleSettings.scaleY / 100 - 1), true, true, true, true);
        };
        var anchorAfter = workerAnchorPoint(workerUnionBounds(pageItems, scaleSettings.useVisibleBounds), anchorRatio);
        var offsetX = anchorBefore[0] - anchorAfter[0];
        var offsetY = anchorBefore[1] - anchorAfter[1];
        if (offsetX === 0 && offsetY === 0) { return; };
        for (var j = 0; j < pageItems.length; j++) {
            pageItems[j].translate(offsetX, offsetY, true, true, true, true);
        };
    };

    /* プレビューを片付ける。複製を消して元を表示し直し、元を選択し直す（ユーザーが別のものを選んでいたらその選択を残す）。
       状態はメインエンジンの $.global に置く
       Clear the preview: remove the copies, show the originals again and reselect them unless the user picked something else.
       The state lives on $.global of the main engine */
    function workerClearPreview() {
        var previewState = $.global.__smartScalePreviewState;
        if (!previewState) { return "OK"; };
        $.global.__smartScalePreviewState = null;
        var keepUserSelection = false;
        try {
            var currentSelection = previewState.doc.selection;
            if (currentSelection instanceof Array && currentSelection.length > 0) {
                keepUserSelection = true;
                for (var i = 0; i < currentSelection.length; i++) {
                    for (var j = 0; j < previewState.copies.length; j++) {
                        if (currentSelection[i] === previewState.copies[j]) { keepUserSelection = false; };
                    };
                };
            };
        } catch (selectionError) {};
        for (var k = 0; k < previewState.copies.length; k++) {
            try { previewState.copies[k].remove(); } catch (removeError) {};
        };
        var restoredItems = [];
        for (var m = 0; m < previewState.items.length; m++) {
            try { previewState.items[m].hidden = false; restoredItems.push(previewState.items[m]); } catch (unhideError) {};
        };
        if (keepUserSelection || restoredItems.length === 0) { return "OK"; };
        try { previewState.doc.selection = restoredItems; } catch (reselectError) {};
        return "OK";
    };

    /* 選択全体の幅・高さ（pt）を "OK:幅,高さ" で返す。外形は環境設定［プレビュー境界を使用］に従う。プレビュー中なら片付けてから測る
       Return the selection size in points as "OK:width,height"; bounds follow the Use Preview Bounds preference. Clears any preview first */
    function workerMeasure() {
        if (app.documents.length === 0) { return "NODOC"; };
        try {
            workerClearPreview();
            var pageItems = workerCollectItems();
            if (pageItems.length === 0) { return "NOSEL"; };
            var selectionBounds = workerUnionBounds(pageItems, app.preferences.getBooleanPreference("includeStrokeInBounds"));
            return "OK:" + (selectionBounds[2] - selectionBounds[0]) + "," + (selectionBounds[1] - selectionBounds[3]);
        } catch (e) {
            return "ERR:" + e;
        };
    };

    /* モードに合わせて拡大・縮小する。個別に両端を固定するときは全体の外形を先に測り、固定する向きの基準をオブジェクトごとに決める
       Scale by mode; when pinning edges per item, measure the overall bounds first and pick each item's anchor on the pinned axes */
    function workerScaleByMode(pageItems, scaleSettings) {
        var anchorRatio = workerAnchorRatio(scaleSettings.anchorIndex);
        if (scaleSettings.mode === "asGroup") { workerScaleAsOne(pageItems, scaleSettings, anchorRatio); return; };
        var unionBounds = (scaleSettings.pinHorizontal || scaleSettings.pinVertical) ? workerUnionBounds(pageItems, scaleSettings.useVisibleBounds) : null;
        for (var i = 0; i < pageItems.length; i++) {
            var itemRatio = anchorRatio;
            if (unionBounds) {
                var itemBounds = scaleSettings.useVisibleBounds ? pageItems[i].visibleBounds : pageItems[i].geometricBounds;
                var keepRatio = workerKeepBoundsRatio(itemBounds, unionBounds);
                itemRatio = [scaleSettings.pinHorizontal ? keepRatio[0] : anchorRatio[0], scaleSettings.pinVertical ? keepRatio[1] : anchorRatio[1]];
            };
            workerScaleAsOne([pageItems[i]], scaleSettings, itemRatio);
        };
    };

    /* オブジェクトを拡大・縮小する。角と線幅・効果は環境設定に従うので、実行中だけ切り替えて元に戻す
       Scale the items; corners and strokes/effects follow preferences, so switch them only while scaling */
    function workerScaleItems(pageItems, scaleSettings) {
        scaleSettings.useVisibleBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");
        var originalScaleLineWeight = app.preferences.getBooleanPreference("scaleLineWeight");
        var originalCornerPolicy = app.preferences.getIntegerPreference("policyForPreservingCorners");
        app.preferences.setBooleanPreference("scaleLineWeight", scaleSettings.scaleStrokes);
        app.preferences.setIntegerPreference("policyForPreservingCorners", scaleSettings.scaleCorners ? 1 : 2);
        var scaleResult = "OK";
        try {
            workerScaleByMode(pageItems, scaleSettings);
            app.redraw();
        } catch (e) {
            scaleResult = "ERR:" + e;
        };
        app.preferences.setBooleanPreference("scaleLineWeight", originalScaleLineWeight);
        app.preferences.setIntegerPreference("policyForPreservingCorners", originalCornerPolicy);
        return scaleResult;
    };

    /* プレビューを作り直す。選択の複製を拡大・縮小し、元は隠して選択を外す（複製は元を隠す前に作る）
       Rebuild the preview: scale copies of the selection, then hide and deselect the originals (copy before hiding) */
    function workerPreview(scaleSettings) {
        if (app.documents.length === 0) { return "NODOC"; };
        var pageItems;
        try {
            workerClearPreview();
            pageItems = workerCollectItems();
        } catch (collectError) { return "ERR:" + collectError; };
        if (pageItems.length === 0) { return "NOSEL"; };
        var previewState = { doc: app.activeDocument, items: pageItems, copies: [] };
        $.global.__smartScalePreviewState = previewState;
        try {
            for (var i = 0; i < pageItems.length; i++) {
                previewState.copies.push(pageItems[i].duplicate(pageItems[i], ElementPlacement.PLACEBEFORE));
            };
            for (var j = 0; j < pageItems.length; j++) { pageItems[j].hidden = true; };
            previewState.doc.selection = null;
        } catch (e) {
            workerClearPreview();
            return "ERR:" + e;
        };
        var scaleResult = workerScaleItems(previewState.copies, scaleSettings);
        if (scaleResult !== "OK") { workerClearPreview(); };
        return scaleResult;
    };

    /* プレビューを片付けてから、選択そのものを拡大・縮小する / Clear the preview, then scale the selection itself */
    function workerApplyScale(scaleSettings) {
        if (app.documents.length === 0) { return "NODOC"; };
        var pageItems;
        try {
            workerClearPreview();
            pageItems = workerCollectItems();
        } catch (collectError) { return "ERR:" + collectError; };
        if (pageItems.length === 0) { return "NOSEL"; };
        return workerScaleItems(pageItems, scaleSettings);
    };

    /* 委譲する worker 関数の全登録（追加漏れ防止） / All worker functions to delegate */
    var WORKER_FUNCTIONS = [workerCollectItems, workerUnionBounds, workerAnchorRatio, workerAnchorPoint, workerKeepBoundsRatio, workerScaleAsOne, workerScaleByMode, workerClearPreview, workerMeasure, workerScaleItems, workerPreview, workerApplyScale];

    // =========================================
    // BridgeTalk 委譲 / BridgeTalk delegation
    // =========================================

    /**
     * 関数のソースから宣言行〜閉じ括弧行だけを切り出す。
     * toString() は改行を CR で返し、周辺のコメント断片まで巻き込むことがあるため、LF に正規化してから取り出す
     * @param {Function} targetFunction - 文字列化する関数
     * @returns {string} 関数宣言だけのソース
     */
    function sliceFunctionSource(targetFunction) {
        var lines = String(targetFunction).replace(/\r\n?/g, "\n").split("\n");
        var firstIndex = -1;
        var lastIndex = -1;
        for (var i = 0; i < lines.length; i++) {
            if (firstIndex < 0 && /^\s*function\s/.test(lines[i])) firstIndex = i;
            if (firstIndex >= 0 && /^\s*\}[;\s]*$/.test(lines[i])) lastIndex = i;
        }
        if (firstIndex < 0) return String(targetFunction);
        if (lastIndex < firstIndex) {
            /* 1行で書かれた関数は、その行だけを取り出す / A function written on one line: keep just that line */
            return /\}[;\s]*$/.test(lines[firstIndex]) ? lines[firstIndex] : lines.slice(firstIndex).join("\n");
        }
        return lines.slice(firstIndex, lastIndex + 1).join("\n");
    }

    /**
     * worker 関数を連結し、末尾の呼び出しを付けて BridgeTalk の本文にする
     * @param {string} callExpression - メインエンジンで評価する呼び出し式
     * @returns {string} 送る本文
     */
    function buildWorkerBody(callExpression) {
        var source = "";
        for (var i = 0; i < WORKER_FUNCTIONS.length; i++) {
            source += sliceFunctionSource(WORKER_FUNCTIONS[i]) + "\n";
        }
        source += callExpression + ";";
        return 'eval(decodeURIComponent("' + encodeURIComponent(source) + '"));';
    }

    /**
     * メインエンジンへ本文を送って同期実行する
     * @param {string} callExpression - メインエンジンで評価する呼び出し式
     * @returns {string} 結果マーカー（"OK…" / "NODOC" / "NOSEL" / "ERR:…"）
     */
    function sendToMainEngine(callExpression) {
        var resultHolder = { result: "ERR:timeout" };
        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = "illustrator";
        bridgeMessage.body = buildWorkerBody(callExpression);
        bridgeMessage.onResult = function (message) { resultHolder.result = String(message.body); };
        bridgeMessage.onError = function (message) { resultHolder.result = "ERR:" + String(message.body); };
        bridgeMessage.send(10); /* 完了まで待つ / wait for completion */
        return resultHolder.result;
    }


    /**
     * 選択全体の大きさを測る
     * @returns {{status: string, width: number, height: number}} status は "OK" / "NODOC" / "NOSEL" / "ERR:…"。OK のときだけ幅・高さ（pt）が入る
     */
    function measureSelection() {
        var measureResult = sendToMainEngine("workerMeasure()");
        if (measureResult.indexOf("OK:") !== 0) return { status: measureResult, width: 0, height: 0 };
        var sizeValues = measureResult.substring(3).split(",");
        return { status: "OK", width: parseFloat(sizeValues[0]), height: parseFloat(sizeValues[1]) };
    }

    /**
     * 拡大・縮小の設定を、worker に渡すオブジェクトリテラルの文字列にする
     * @param {Object} scaleSettings - scaleX / scaleY / mode（"perItem" / "asGroup"）/ pinHorizontal / pinVertical / anchorIndex / scaleCorners / scaleStrokes / scalePatterns / scaleGradients
     * @returns {string} オブジェクトリテラル
     */
    function toScaleSettingsLiteral(scaleSettings) {
        return "{" +
            "scaleX:" + scaleSettings.scaleX + "," +
            "scaleY:" + scaleSettings.scaleY + "," +
            "mode:\"" + scaleSettings.mode + "\"," +
            "pinHorizontal:" + !!scaleSettings.pinHorizontal + "," +
            "pinVertical:" + !!scaleSettings.pinVertical + "," +
            "anchorIndex:" + scaleSettings.anchorIndex + "," +
            "scaleCorners:" + !!scaleSettings.scaleCorners + "," +
            "scaleStrokes:" + !!scaleSettings.scaleStrokes + "," +
            "scalePatterns:" + !!scaleSettings.scalePatterns + "," +
            "scaleGradients:" + !!scaleSettings.scaleGradients +
            "}";
    }

    /**
     * 選択を拡大・縮小する（プレビュー中なら片付けてから）
     * @param {Object} scaleSettings - toScaleSettingsLiteral() と同じ設定
     * @returns {string} 結果マーカー（"OK" / "NODOC" / "NOSEL" / "ERR:…"）
     */
    function applyScaleToSelection(scaleSettings) {
        return sendToMainEngine("workerApplyScale(" + toScaleSettingsLiteral(scaleSettings) + ")");
    }

    /**
     * 選択の複製を拡大・縮小してプレビューを作り直す
     * @param {Object} scaleSettings - toScaleSettingsLiteral() と同じ設定
     * @returns {string} 結果マーカー（"OK" / "NODOC" / "NOSEL" / "ERR:…"）
     */
    function previewScaleOfSelection(scaleSettings) {
        return sendToMainEngine("workerPreview(" + toScaleSettingsLiteral(scaleSettings) + ")");
    }

    /**
     * プレビューを片付ける
     * @param {boolean} [isAsync] - true なら結果を待たずに送る（パレットを閉じる途中は DOM に触れないため）
     * @returns {void}
     */
    function clearScalePreview(isAsync) {
        if (!isAsync) {
            sendToMainEngine("workerClearPreview()");
            return;
        }
        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = "illustrator";
        bridgeMessage.body = buildWorkerBody("workerClearPreview()");
        bridgeMessage.send();
    }

    /**
     * 結果マーカーに応じてアラートを出す
     * @param {string} resultMarker - "NODOC" / "NOSEL" / "ERR:…"（"OK" なら何もしない）
     * @returns {void}
     */
    function alertForResult(resultMarker) {
        if (resultMarker.indexOf("OK") === 0) return;
        if (resultMarker === "NODOC") alert(getLabel("alert.noDocument"));
        else if (resultMarker === "NOSEL") alert(getLabel("alert.selectObjects"));
        else alert(getLabel("alert.error", { message: resultMarker.replace(/^ERR:/, "") }));
    }

    /**
     * 縦横とも 100%（元のまま）かを返す
     * @param {{x: number, y: number}} scalePercent - 横・縦の拡大・縮小率（%）
     * @returns {boolean} 元のままなら true
     */
    function isOriginalScale(scalePercent) {
        return scalePercent.x === 100 && scalePercent.y === 100;
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
    // パレット / Palette
    // =========================================

    // 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)

    // -----------------------------------------
    // 基準点ウィジェットの寸法 / Anchor widget metrics
    // -----------------------------------------
    var ANCHOR_WIDGET_SIZE      = 66;   /* ウィジェット全体の一辺 / overall size of the widget */
    var ANCHOR_WIDGET_CELL_SIZE = 9;    /* □1個の一辺 / size of one square */
    var ANCHOR_WIDGET_CELL_GAP  = 7.5;  /* □どうしの間隔 / gap between squares */
    var ANCHOR_WIDGET_NONE      = -1;   /* 未選択のインデックス / index while nothing is selected */

    /* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
       Cell names in row-major order, matching the Transformation enumeration */
    var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線 / Rules joining the outer squares (the center stands alone) */
    var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    // -----------------------------------------
    // 基準点ウィジェットの配色 / Anchor widget colors
    // -----------------------------------------
    var ANCHOR_WIDGET_UI_DARK = isDarkUI();
    /* 枠線・ケイ線はグレー、選択セルの塗りはライトで濃いグレー・ダークで明るいグレー（既存スクリプトの配色を踏襲）。
       無効時は同じ色を半透明にして背景へ沈める（不透明の薄いグレーだとダークUIで逆に明るく浮くため）
       Gray rules; the selected fill is dark gray on light UI and light gray on dark UI (as in the existing scripts).
       Disabled colors are translucent versions so they sink into any background */
    var ANCHOR_WIDGET_LINE_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 1]     : [0.42, 0.42, 0.42, 1];  /* 枠線・ケイ線 / rules */
    var ANCHOR_WIDGET_FILL_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 1]     : [0.27, 0.27, 0.27, 1];  /* 選択セルの塗り / selected fill */
    var ANCHOR_WIDGET_DIM_LINE_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 0.4]   : [0.42, 0.42, 0.42, 0.4];  /* 無効時の枠線 / rules when disabled */
    var ANCHOR_WIDGET_DIM_FILL_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 0.3]   : [0.27, 0.27, 0.27, 0.3];  /* 無効時の塗り / fill when disabled */

    // -----------------------------------------
    // ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 基準点（3×3）を選ぶウィジェットを追加する。クリックしたセルを選び、onChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
     * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
     * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）/ size（一辺。既定 66）
     * @returns {Button} ウィジェット（値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
     */
    function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
        var anchorOptions = widgetOptions || {};
        var widgetSize = anchorOptions.size || ANCHOR_WIDGET_SIZE;
        var anchorWidget = parent.add("button", undefined, "");
        anchorWidget.minimumSize = [widgetSize, widgetSize];
        anchorWidget.preferredSize = [widgetSize, widgetSize];
        anchorWidget.maximumSize = [widgetSize, widgetSize];
        anchorWidget.isAnchorWidget = true; /* redrawAnchorWidgetsIn() の目印 / marker for redrawAnchorWidgetsIn() */
        anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(anchorOptions.disabledCells);
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(initialValue, anchorWidget.anchorAllowNone);
        anchorWidget.onDraw = function () { drawAnchorWidget(anchorWidget); };
        anchorWidget.onClick = function () {}; /* セルの判定は mousedown で行う / hit-testing happens in mousedown */

        /* クリック座標（コントロール基準）を3分割してセルを判定する / split the control-relative click into thirds */
        anchorWidget.addEventListener("mousedown", function (event) {
            if (!isAnchorWidgetEnabledInTree(anchorWidget)) return;
            var cellIndex = getAnchorCellAt(event.clientX, event.clientY, anchorWidget.size[0], anchorWidget.size[1]);
            if (anchorWidget.anchorDisabledCells[cellIndex]) return;
            anchorWidget.anchorWidgetIndex = cellIndex;
            redrawAnchorWidget(anchorWidget);
            if (onChange) onChange(cellIndex, anchorWidget);
        });
        return anchorWidget;
    }

    /**
     * 選択中のセルのインデックスを返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {number} 0〜8（行優先）。未選択なら -1
     */
    function getAnchorWidgetIndex(anchorWidget) {
        return anchorWidget.anchorWidgetIndex;
    }

    /**
     * 選択中のセルの名前を返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {string} "topLeft" など。未選択なら ""
     */
    function getAnchorWidgetName(anchorWidget) {
        return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
    }

    /**
     * 選択するセルを変えて描き直す（onChange は呼ばない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * ウィジェットの有効／無効を切り替えて描き直す（無効の間は薄く描き、クリックも無視する）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
        anchorWidget.enabled = isEnabled;
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * 選べないセルを指定し直して描き直す（選択中のセルは変えない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
     * @returns {void}
     */
    function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(disabledCells);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * コンテナ以下にある基準点ウィジェットをすべて描き直す。パネルや行の enabled を切り替えたあとに呼ぶ
     * @param {Object} container - パネル・グループ・ウィンドウなど
     * @returns {void}
     */
    function redrawAnchorWidgetsIn(container) {
        if (container.isAnchorWidget) {
            redrawAnchorWidget(container);
            return;
        }
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            redrawAnchorWidgetsIn(container.children[i]);
        }
    }

    // -----------------------------------------
    // 値の変換 / Value helpers
    // -----------------------------------------
    /**
     * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
     * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
     * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
     * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
     */
    function resolveAnchorWidgetIndex(anchorValue, allowNone) {
        if (typeof anchorValue === "string") {
            for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
                if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
            }
            return 4;
        }
        if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
        if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
            return anchorValue;
        }
        return 4;
    }

    /**
     * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [横の割合, 縦の割合]
     */
    function getAnchorRatio(anchorValue) {
        var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    }

    /**
     * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
     * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [x, y]
     */
    function getAnchorPointOnBounds(bounds, anchorValue) {
        var anchorRatio = getAnchorRatio(anchorValue);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
            bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
        ];
    }

    /**
     * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
     * 基準は効果を含まない境界（geometricBounds）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {Transformation} Transformation.TOPLEFT など
     */
    function getAnchorTransformation(anchorValue) {
        var transformations = [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ];
        return transformations[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * symbols.add() に渡す登録点を返す（Illustrator 専用）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
     */
    function getAnchorSymbolRegistrationPoint(anchorValue) {
        var registrationPoints = [
            SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
        ];
        return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * クリック位置からセルのインデックスを求める（ウィジェットを縦横3等分し、外にはみ出した座標は端のセルに寄せる）
     * @param {number} clickX - コントロール基準の x
     * @param {number} clickY - コントロール基準の y
     * @param {number} widgetWidth - ウィジェットの幅
     * @param {number} widgetHeight - ウィジェットの高さ
     * @returns {number} 0〜8
     */
    function getAnchorCellAt(clickX, clickY, widgetWidth, widgetHeight) {
        var column = Math.min(2, Math.max(0, Math.floor(clickX / (widgetWidth / 3))));
        var row = Math.min(2, Math.max(0, Math.floor(clickY / (widgetHeight / 3))));
        return row * 3 + column;
    }

    /**
     * セルのインデックスの配列を、9個の真偽値に直す
     * @param {number[]} [cellIndexes] - セルのインデックスの配列
     * @returns {boolean[]} 含まれるセルだけ true
     */
    function toAnchorCellFlags(cellIndexes) {
        var cellFlags = [false, false, false, false, false, false, false, false, false];
        if (!cellIndexes) return cellFlags;
        for (var i = 0; i < cellIndexes.length; i++) {
            if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
        }
        return cellFlags;
    }

    // -----------------------------------------
    // 描画 / Drawing
    // -----------------------------------------
    /**
     * ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立。選択セルだけ塗る）
     * @param {Button} anchorWidget - 描くウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        var halfCell = cellSize / 2;
        /* 自作描画は自動でディムにならないので、親までたどって判定する / custom drawing is not dimmed automatically */
        var isEnabled = isAnchorWidgetEnabledInTree(anchorWidget);

        /* ボタンの地をコントロールの地色で塗り、パネルに溶け込ませる（backgroundColor が無い環境では例外）
           Paint the control's own background so the widget blends into the panel; throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
        var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);
        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1);
        for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
                graphics.lineTo(cellB[0], cellB[1] + halfCell);
            } else {
                /* 縦方向：下の□へ / vertical: to the square below */
                graphics.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
                graphics.lineTo(cellB[0] + halfCell, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < 9; i++) {
            var isCellEnabled = isEnabled && !anchorWidget.anchorDisabledCells[i];
            drawAnchorWidgetCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.anchorWidgetIndex, isCellEnabled);
        }
    }

    /**
     * □を1つ描く（選択中だけ塗り、枠は塗りの上に重ねる）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} cellX - 左端
     * @param {number} cellY - 上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {boolean} isEnabled - 選べるセルなら true（false なら薄く描く）
     * @returns {void}
     */
    function drawAnchorWidgetCell(graphics, cellX, cellY, isSelected, isEnabled) {
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        /* rectPath の前には毎回 newPath()（呼ばないとパスが累積して塗りが線画になる）
           Always call newPath() before rectPath(), or paths accumulate and fills turn into outlines */
        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(cellX, cellY, cellSize, cellSize);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_FILL_COLOR : ANCHOR_WIDGET_DIM_FILL_COLOR));
        }
        graphics.newPath();
        graphics.rectPath(cellX, cellY, cellSize, cellSize);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isAnchorWidgetEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false) return false;
        }
        return true;
    }

    /**
     * ウィジェットの onDraw を呼び直す。notify("onDraw") は環境によって例外や空振りになるため、隠して再表示して描き直させる
     * @param {Button} anchorWidget - 描き直すウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        anchorWidget.hide();
        anchorWidget.show();
    }

    // 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    var LINK_UI_DARK = isDarkUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @param {number[]} [iconSize] - アイコンの [幅, 高さ]（省略時は LINK_ICON_SIZE。絵は 22px 基準から拡大縮小する）
     * @param {number} [chainRatio] - 鎖の絵の大きさの比率（省略時は 1。枠・地の大きさは変えず、鎖だけ縮める）
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle, iconSize, chainRatio) {
        var toggleSize = iconSize || LINK_ICON_SIZE;
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = toggleSize;
        linkToggle.minimumSize = toggleSize;
        linkToggle.maximumSize = toggleSize;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = toggleSize[0];
            var iconHeight = toggleSize[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR, chainRatio);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @param {number} [chainRatio] - 鎖の大きさの比率（省略時は 1）。中央に置いたまま縮める
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

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

    /**
     * ラジオボタンを縦に並べたパネルを追加する
     * @param {Group} parent - 追加先
     * @param {string} panelTitle - パネル名
     * @param {string[]} radioTitles - ラジオボタンの文言
     * @returns {RadioButton[]} 追加したラジオボタン
     */
    function addRadioPanel(parent, panelTitle, radioTitles) {
        var radioPanel = parent.add("panel", undefined, panelTitle);
        setupPanel(radioPanel, 6);
        var radioButtons = [];
        for (var i = 0; i < radioTitles.length; i++) {
            radioButtons.push(radioPanel.add("radiobutton", undefined, radioTitles[i]));
        }
        return radioButtons;
    }

    /**
     * スケールのボタンの文言を返す（50% / +10% / -10% など）
     * @param {Object} presetButton - setTo か stepBy を持つ設定
     * @returns {string} ボタンの文言
     */
    function getScalePresetLabel(presetButton) {
        if (presetButton.setTo !== undefined) return presetButton.setTo + "%";
        return (presetButton.stepBy > 0 ? "+" : "") + presetButton.stepBy + "%";
    }

    /**
     * スケールのボタンのツールチップを返す
     * @param {Object} presetButton - setTo か stepBy を持つ設定
     * @returns {string} ツールチップ
     */
    function getScalePresetTooltip(presetButton) {
        if (presetButton.setTo !== undefined) return getLabel("tooltip.presetSetTo", { value: presetButton.setTo });
        var tooltipPath = presetButton.stepBy > 0 ? "tooltip.presetStepUp" : "tooltip.presetStepDown";
        return getLabel(tooltipPath, { value: Math.abs(presetButton.stepBy) });
    }

    /**
     * スケールのボタンを縦に並べて追加する
     * @param {Group} parent - 追加先（縦並びのグループ）
     * @param {Function} onPresetClick - クリックで呼ぶ関数（引数はボタンの設定）
     * @returns {Checkbox} ［相対］のチェックボックス
     */
    function addScalePresetButtons(parent, onPresetClick) {
        for (var i = 0; i < SCALE_PRESET_BUTTONS.length; i++) {
            var presetBlock = SCALE_PRESET_BUTTONS[i];
            var presetBlockGroup = parent.add("group");
            presetBlockGroup.orientation = "column";
            presetBlockGroup.alignChildren = ["fill", "top"];
            presetBlockGroup.spacing = PRESET_BUTTON_SPACING;
            if (Math.abs(presetBlock[0].stepBy) === 10) {
                presetBlockGroup.margins = [0, PRESET_STEP10_MARGIN, 0, PRESET_STEP10_MARGIN];
            }
            for (var j = 0; j < presetBlock.length; j++) {
                var presetButton = presetBlockGroup.add("button", undefined, getScalePresetLabel(presetBlock[j]));
                presetButton.preferredSize.width = SCALE_PRESET_BUTTON_WIDTH;
                presetButton.helpTip = getScalePresetTooltip(presetBlock[j]);
                presetButton.onClick = createScalePresetHandler(presetBlock[j], onPresetClick);
            }
        }
        /* ±のボタンを今の倍率に掛けるか / whether the ± buttons multiply the current scale */
        var relativeStepGroup = parent.add("group");
        relativeStepGroup.alignment = ["center", "top"];
        relativeStepGroup.margins = [0, RELATIVE_STEP_TOP_MARGIN, 0, 0];
        var relativeStepCheckbox = relativeStepGroup.add("checkbox", undefined, getLabel("checkbox.relativeStep"));
        relativeStepCheckbox.helpTip = getLabel("tooltip.relativeStep");
        return relativeStepCheckbox;
    }

    /**
     * スケールのボタン1つ分のクリック処理を作る
     * @param {Object} presetButton - setTo か stepBy を持つ設定
     * @param {Function} onPresetClick - クリックで呼ぶ関数
     * @returns {Function} onClick に入れる関数
     */
    function createScalePresetHandler(presetButton, onPresetClick) {
        return function () { onPresetClick(presetButton); };
    }

    /**
     * スケールのボタンを今の倍率に当てはめる。stepBy は縦横それぞれに、足すか（元の大きさに対して）掛けるか（相対）で当てる
     * @param {Object} presetButton - setTo か stepBy を持つ設定
     * @param {{x: number, y: number}} scalePercent - 今の横・縦の拡大・縮小率（%）
     * @param {boolean} isRelative - true なら今の倍率に (100 + stepBy)% を掛ける
     * @returns {{x: number, y: number}} 新しい横・縦の拡大・縮小率（%）
     */
    function applyScalePreset(presetButton, scalePercent, isRelative) {
        if (presetButton.setTo !== undefined) return { x: presetButton.setTo, y: presetButton.setTo };

        /**
         * 1つの向きの倍率を増減する（小数の誤差を丸め、下限で止める）
         * @param {number} axisPercent - 今の倍率（%）
         * @returns {number} 新しい倍率（%）
         */
        function stepAxis(axisPercent) {
            var nextPercent = isRelative
                ? axisPercent * (100 + presetButton.stepBy) / 100
                : axisPercent + presetButton.stepBy;
            return Math.max(MIN_SCALE_PERCENT, Math.round(nextPercent * 1000) / 1000);
        }

        return { x: stepAxis(scalePercent.x), y: stepAxis(scalePercent.y) };
    }


    /**
     * 子を横に並べ、天地中央にそろえる行を追加する
     * @param {Group|Panel} parent - 追加先
     * @returns {Group} 行の group
     */
    function addRowGroup(parent) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * スケール・幅・高さの項目名を、幅をそろえて右揃えで追加する
     * @param {Group} parent - 追加先の行
     * @param {string} labelPath - 項目名の LABELS のパス
     * @returns {StaticText} 項目名
     */
    function addSizeFieldLabel(parent, labelPath) {
        var fieldLabel = parent.add("statictext", undefined, labelText(labelPath));
        fieldLabel.preferredSize.width = SIZE_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return fieldLabel;
    }

    /**
     * 入力欄の直前にある項目名のクリックで、入力欄にフォーカスを移す（addSteppedField() の項目名と同じ挙動。無効の間は移さない）。
     * 単位や「→」に付けないよう、行の先頭にあるか末尾がコロンの statictext だけを項目名とみなす
     * @param {Group} stepperFieldGroup - ∧∨と入力欄をまとめた group（追加した直後で、親の末尾にある）
     * @param {EditText} numberInput - 入力欄
     * @returns {void}
     */
    function focusInputOnPrecedingLabel(stepperFieldGroup, numberInput) {
        var siblings = stepperFieldGroup.parent.children;
        if (siblings.length < 2) return;
        var labelIndex = siblings.length - 2;
        var fieldLabel = siblings[labelIndex];
        if (fieldLabel.type !== "statictext") return;
        if (labelIndex > 0 && !/[:：]\s*$/.test(fieldLabel.text)) return;
        fieldLabel.addEventListener("click", function () {
            if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
        });
    }

    /**
     * ∧∨と入力欄を隙間なく並べて追加する。↑↓キーも∧∨と同じ処理で増減する
     * @param {Group} parent - 追加先の行
     * @param {string} initialText - 入力欄の最初の値
     * @param {number} fieldCharacters - 入力欄の幅（文字数）
     * @param {Object} stepOptions - addStepper() に渡す設定（min / onStep など）
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addStepperInput(parent, initialText, fieldCharacters, stepOptions) {
        var stepperInputGroup = parent.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0; /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        stepperInputGroup.margins = 0;
        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = fieldCharacters;
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        focusInputOnPrecedingLabel(stepperInputGroup, numberInput);
        return numberInput;
    }

    /**
     * 「スケール：[∧∨][100] %」の行を追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} onScaleEdited - 値を変えたときに呼ぶ関数
     * @returns {EditText} スケールの入力欄
     */
    function addScaleRow(parent, onScaleEdited) {
        var scaleRowGroup = addRowGroup(parent);
        addSizeFieldLabel(scaleRowGroup, "fieldLabel.scale");
        var scaleInput = addStepperInput(scaleRowGroup, String(DEFAULT_SCALE_PERCENT), SCALE_FIELD_CHARACTERS, {
            min: MIN_SCALE_PERCENT,
            onStep: function () { onScaleEdited(); }
        });
        scaleInput.helpTip = getLabel("tooltip.scale");
        scaleInput.onChanging = function () { onScaleEdited(); };
        scaleRowGroup.add("statictext", undefined, getLabel("fieldLabel.percentUnit"));
        return scaleInput;
    }

    /**
     * 「幅：230 → [∧∨][230] mm」の行を追加する（単位は入力欄の右だけ）
     * @param {Group} parent - 追加先
     * @param {string} labelPath - 項目名の LABELS のパス
     * @param {string} scaleAxis - この欄が決める倍率の向き（"x" / "y"）
     * @param {string} unitLabel - 定規の単位の表示
     * @param {Function} onSizeEdited - 値を変えたときに呼ぶ関数（引数は入力欄）
     * @returns {EditText} 目標の長さの入力欄（今の長さの表示は .currentSizeText、項目名は .fieldLabel）
     */
    function addSizeRow(parent, labelPath, scaleAxis, unitLabel, onSizeEdited) {
        var sizeRowGroup = addRowGroup(parent);
        var sizeLabel = addSizeFieldLabel(sizeRowGroup, labelPath);
        var currentSizeText = sizeRowGroup.add("statictext", undefined, getLabel("fieldLabel.noSize"));
        currentSizeText.preferredSize.width = CURRENT_SIZE_WIDTH;
        currentSizeText.justify = "right";
        currentSizeText.helpTip = getLabel(scaleAxis === "x" ? "tooltip.currentWidth" : "tooltip.currentHeight");
        sizeRowGroup.add("statictext", undefined, getLabel("fieldLabel.sizeArrow"));

        var sizeInput = addStepperInput(sizeRowGroup, "", SIZE_FIELD_CHARACTERS, {
            min: MIN_TARGET_LENGTH,
            onStep: function () { onSizeEdited(sizeInput); }
        });
        sizeInput.helpTip = getLabel("tooltip.sizeField");
        sizeInput.sourceLength = 0;
        sizeInput.scaleAxis = scaleAxis;
        sizeInput.currentSizeText = currentSizeText;
        sizeInput.fieldLabel = sizeLabel;
        sizeInput.onChanging = function () { onSizeEdited(sizeInput); };
        sizeRowGroup.add("statictext", undefined, unitLabel);
        return sizeInput;
    }

    /**
     * 「サイズ」パネル（スケール・幅・高さと連動ボタン）を追加する
     * @param {Group} parent - 追加先
     * @param {string} unitLabel - 定規の単位の表示
     * @param {Object} sizeHandlers - onScaleEdited() / onSizeEdited(sizeInput) / onLinkToggled()
     * @returns {{scaleInput: EditText, widthInput: EditText, heightInput: EditText, linkToggle: Group}} 作ったコントロール
     */
    function buildSizePanel(parent, unitLabel, sizeHandlers) {
        var sizePanel = parent.add("panel", undefined, getLabel("panel.size"));
        setupPanel(sizePanel);
        var scaleInput = addScaleRow(sizePanel, sizeHandlers.onScaleEdited);

        /* 幅・高さ（今の大きさ → 目標の大きさ）と連動ボタン / Width and height (current -> target) with the link toggle */
        var sizeLinkRowGroup = addRowGroup(sizePanel);
        var sizeFieldsColumn = sizeLinkRowGroup.add("group");
        sizeFieldsColumn.orientation = "column";
        sizeFieldsColumn.alignChildren = ["left", "top"];
        var widthInput = addSizeRow(sizeFieldsColumn, "fieldLabel.width", "x", unitLabel, sizeHandlers.onSizeEdited);
        var heightInput = addSizeRow(sizeFieldsColumn, "fieldLabel.height", "y", unitLabel, sizeHandlers.onSizeEdited);
        var linkToggle = addLinkToggle(sizeLinkRowGroup, true, sizeHandlers.onLinkToggled);
        linkToggle.helpTip = getLabel("tooltip.linkToggle");
        return { scaleInput: scaleInput, widthInput: widthInput, heightInput: heightInput, linkToggle: linkToggle };
    }

    /**
     * 「モード」パネルを追加する。［個別に］の下に、両端を固定するオプションを字下げして置く
     * @param {Group} parent - 追加先
     * @returns {{asGroupRadio: RadioButton, perItemRadio: RadioButton, pinOptionsGroup: Group, pinHorizontalCheckbox: Checkbox, pinVerticalCheckbox: Checkbox}} 作ったコントロール
     */
    function buildModePanel(parent) {
        var modeRadios = addRadioPanel(parent, getLabel("panel.mode"), [getLabel("radio.asGroup"), getLabel("radio.perItem")]);
        var asGroupRadio = modeRadios[0];
        var perItemRadio = modeRadios[1];
        asGroupRadio.helpTip = getLabel("tooltip.asGroup");
        perItemRadio.helpTip = getLabel("tooltip.perItem");

        var pinOptionsGroup = perItemRadio.parent.add("group");
        pinOptionsGroup.orientation = "column";
        pinOptionsGroup.alignChildren = ["left", "center"];
        pinOptionsGroup.margins = [PIN_OPTIONS_INDENT, 0, 0, 0];
        var pinHorizontalCheckbox = pinOptionsGroup.add("checkbox", undefined, getLabel("checkbox.pinHorizontal"));
        var pinVerticalCheckbox = pinOptionsGroup.add("checkbox", undefined, getLabel("checkbox.pinVertical"));
        pinHorizontalCheckbox.helpTip = getLabel("tooltip.pinHorizontal");
        pinVerticalCheckbox.helpTip = getLabel("tooltip.pinVertical");
        return {
            asGroupRadio: asGroupRadio,
            perItemRadio: perItemRadio,
            pinOptionsGroup: pinOptionsGroup,
            pinHorizontalCheckbox: pinHorizontalCheckbox,
            pinVerticalCheckbox: pinVerticalCheckbox
        };
    }

    /**
     * 「基準点」パネルを追加する
     * @param {Group} parent - 追加先
     * @param {Function} onAnchorChange - セルを選んだときに呼ぶ関数
     * @returns {Button} 基準点ウィジェット
     */
    function buildAnchorPanel(parent, onAnchorChange) {
        var anchorPanel = parent.add("panel", undefined, getLabel("panel.anchor"));
        setupPanel(anchorPanel);
        anchorPanel.alignChildren = ["center", "center"]; /* 基準点は中央に置く / Center the anchor widget */
        var anchorWidget = addAnchorWidget(anchorPanel, DEFAULT_ANCHOR, onAnchorChange);
        anchorWidget.helpTip = getLabel("tooltip.anchor");
        return anchorWidget;
    }

    /**
     * 「オプション」パネルを追加する
     * @param {Group} parent - 追加先
     * @returns {{scaleCornersCheckbox: Checkbox, scaleStrokeCheckbox: Checkbox, scalePatternCheckbox: Checkbox, scaleGradientCheckbox: Checkbox}} 作ったチェックボックス
     */
    function buildOptionsPanel(parent) {
        var optionsPanel = parent.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 6);

        /**
         * ツールチップ付きのチェックボックスを追加する
         * @param {string} labelKey - LABELS.checkbox と LABELS.tooltip で共通のキー
         * @returns {Checkbox} 追加したチェックボックス
         */
        function addOptionCheckbox(labelKey) {
            var optionCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox." + labelKey));
            optionCheckbox.helpTip = getLabel("tooltip." + labelKey);
            return optionCheckbox;
        }

        return {
            scaleCornersCheckbox: addOptionCheckbox("scaleCorners"),
            scaleStrokeCheckbox: addOptionCheckbox("strokeWidth"),
            scalePatternCheckbox: addOptionCheckbox("pattern"),
            scaleGradientCheckbox: addOptionCheckbox("gradient")
        };
    }

    /**
     * 2つのラベルのうち、文字数の多いほうを返す（あとで文言を切り替えるコントロールの幅を確保する）
     * @param {string} firstLabelPath - 1つ目の LABELS のパス
     * @param {string} secondLabelPath - 2つ目の LABELS のパス
     * @returns {string} 長いほうの文言
     */
    function getLongerLabel(firstLabelPath, secondLabelPath) {
        var firstLabel = getLabel(firstLabelPath);
        var secondLabel = getLabel(secondLabelPath);
        return (secondLabel.length > firstLabel.length) ? secondLabel : firstLabel;
    }

    /**
     * 下部のボタン行（左：閉じる・リセット、右：プレビュー・適用）を追加する
     * @param {Window} targetWindow - 追加先のパレット
     * @returns {{btnReset: Button, previewCheckbox: Checkbox, btnClose: Button, btnApply: Button}} 作ったコントロール
     */
    function buildPaletteButtons(targetWindow) {
        var buttonRow = addButtonRow(targetWindow);
        var btnClose = buttonRow.leftGroup.add("button", undefined, getLabel("button.close"));
        btnClose.helpTip = getLabel("tooltip.close");
        var btnReset = buttonRow.leftGroup.add("button", undefined, getLabel("button.reset"));
        btnReset.helpTip = getLabel("tooltip.reset");
        /* ［プレビュー］は［適用］の前に結果を見るものなので、［適用］と並べる / Preview sits next to Apply, whose result it shows */
        var previewCheckbox = buttonRow.rightGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        /* 文言は［プレビュー］に合わせて［適用］／［確定］を切り替える。作成後は幅が広がらないので、長いほうの文言で作る
           The label switches between Apply and Commit; buttons do not grow after creation, so start with the longer one */
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLongerLabel("button.apply", "button.commit"));
        return { btnReset: btnReset, previewCheckbox: previewCheckbox, btnClose: btnClose, btnApply: btnApply };
    }

    /**
     * パレットを組み立てる。［適用］で選択中のオブジェクトを拡大・縮小する
     * @returns {Window} 組み立てたパレット
     */
    function createPalette() {
        var scalePalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(scalePalette);

        /* 左カラム：設定、右カラム：スケールのボタン。［適用］などのボタン行はその下
           Left column: settings; right column: scale buttons; the Apply/Close row sits below */
        var paletteColumnsGroup = scalePalette.add("group");
        paletteColumnsGroup.orientation = "row";
        paletteColumnsGroup.alignChildren = ["fill", "top"];
        paletteColumnsGroup.spacing = COLUMN_SPACING;
        var settingsColumn = paletteColumnsGroup.add("group");
        settingsColumn.orientation = "column";
        settingsColumn.alignChildren = ["fill", "top"];
        var presetColumn = paletteColumnsGroup.add("group");
        presetColumn.orientation = "column";
        presetColumn.alignChildren = ["fill", "top"];
        presetColumn.spacing = PRESET_BUTTON_SPACING;

        var rulerUnit = getUnitInfo("rulerType");
        var sizeControls = buildSizePanel(settingsColumn, rulerUnit.label, {
            onScaleEdited: onScaleTyped,
            onSizeEdited: onSizeEdited,
            onLinkToggled: onLinkToggled
        });
        var scaleInput = sizeControls.scaleInput;
        var widthInput = sizeControls.widthInput;
        var heightInput = sizeControls.heightInput;
        var linkToggle = sizeControls.linkToggle;
        var sizeInputs = [widthInput, heightInput];
        var relativeStepCheckbox = addScalePresetButtons(presetColumn, onPresetClick);

        /* モード・基準点・オプションは横に3つ並べる / mode, anchor and options in three columns */
        var panelColumnsGroup = settingsColumn.add("group");
        panelColumnsGroup.orientation = "row";
        panelColumnsGroup.alignChildren = ["fill", "fill"];
        panelColumnsGroup.spacing = COLUMN_SPACING;
        var modeControls = buildModePanel(panelColumnsGroup);
        var asGroupRadio = modeControls.asGroupRadio;
        var perItemRadio = modeControls.perItemRadio;
        var pinHorizontalCheckbox = modeControls.pinHorizontalCheckbox;
        var pinVerticalCheckbox = modeControls.pinVerticalCheckbox;
        var anchorWidget = buildAnchorPanel(panelColumnsGroup, function () { updatePreview(); });
        var optionCheckboxes = buildOptionsPanel(panelColumnsGroup);
        var paletteButtons = buildPaletteButtons(scalePalette);
        var previewCheckbox = paletteButtons.previewCheckbox;

        /* 横・縦の倍率は丸めずに持つ（欄の表示は丸める） / keep the unrounded x and y scale (fields show rounded values) */
        var currentScale = { x: DEFAULT_SCALE_PERCENT, y: DEFAULT_SCALE_PERCENT };
        /* 最後に値を入れた幅・高さの欄。測り直したとき、その大きさになる倍率を求め直す
           The width or height field edited last; after re-measuring, the scale is recomputed to reach that size */
        var lastEditedSizeInput = null;
        var isBusy = false;          /* 委譲の再入防止 / re-entrancy guard for delegation */
        var hasPendingScale = false; /* 倍率を変えてから、まだ適用していない / the scale changed and is not applied yet */
        var isPreviewShown = false;  /* プレビューの複製がある / preview copies exist */

        /**
         * 倍率に合わせて、スケール・幅・高さの欄を書き直す。縦横の倍率が違うときスケール欄は空にする
         * @param {EditText|null} editedInput - 入力中の欄（書き直さない）
         * @returns {void}
         */
        function syncScaleFields(editedInput) {
            if (editedInput !== scaleInput) {
                scaleInput.text = (currentScale.x === currentScale.y) ? formatRounded(currentScale.x, 3) : "";
            }
            for (var i = 0; i < sizeInputs.length; i++) {
                if (sizeInputs[i] === editedInput) continue;
                if (sizeInputs[i].sourceLength <= 0) {
                    sizeInputs[i].text = "";
                    continue;
                }
                var axisPercent = currentScale[sizeInputs[i].scaleAxis];
                sizeInputs[i].text = formatRounded(sizeInputs[i].sourceLength * axisPercent / 100 / rulerUnit.pointsPerUnit, 2);
            }
        }

        /**
         * 倍率を変え、欄を合わせる
         * @param {{x: number, y: number}} nextScale - 新しい横・縦の拡大・縮小率（%）
         * @param {EditText|null} editedInput - 入力中の欄（書き直さない）
         * @returns {void}
         */
        function setCurrentScale(nextScale, editedInput) {
            currentScale = nextScale;
            syncScaleFields(editedInput);
        }

        /**
         * 操作で倍率を変えたとき：欄を合わせ、未適用の印を付けてプレビューを更新する
         * @param {{x: number, y: number}} nextScale - 新しい横・縦の拡大・縮小率（%）
         * @param {EditText|null} editedInput - 入力中の欄（書き直さない）
         * @returns {void}
         */
        function changeScaleByUser(nextScale, editedInput) {
            setCurrentScale(nextScale, editedInput);
            hasPendingScale = true;
            updatePreview();
        }

        /**
         * スケール欄を変えたとき：縦横とも同じ倍率にする（数値でない・下限未満の間は何もしない）
         * @returns {void}
         */
        function onScaleTyped() {
            var scalePercent = parseFloat(scaleInput.text);
            if (isNaN(scalePercent) || scalePercent < MIN_SCALE_PERCENT) return;
            lastEditedSizeInput = null;
            changeScaleByUser({ x: scalePercent, y: scalePercent }, scaleInput);
        }

        /**
         * スケールのボタンを押したとき
         * @param {Object} presetButton - setTo か stepBy を持つ設定
         * @returns {void}
         */
        function onPresetClick(presetButton) {
            lastEditedSizeInput = null;
            changeScaleByUser(applyScalePreset(presetButton, currentScale, relativeStepCheckbox.value), null);
        }

        /**
         * 幅・高さの欄の値から、その向きの倍率を求める
         * @param {EditText} sizeInput - 幅・高さの欄
         * @returns {number|null} 倍率（%）。求められないときは null
         */
        function getAxisPercentFromSize(sizeInput) {
            var targetLength = parseFloat(sizeInput.text);
            if (isNaN(targetLength) || targetLength <= 0 || sizeInput.sourceLength <= 0) return null;
            var axisPercent = targetLength * rulerUnit.pointsPerUnit / sizeInput.sourceLength * 100;
            return (axisPercent < MIN_SCALE_PERCENT) ? null : axisPercent;
        }

        /**
         * 幅・高さの欄が決めた倍率から、横・縦の倍率を作る。連動中は縦横とも同じ倍率にする
         * @param {EditText} sizeInput - 幅・高さの欄
         * @param {number} axisPercent - その欄の向きの倍率（%）
         * @returns {{x: number, y: number}} 横・縦の拡大・縮小率（%）
         */
        function buildScaleFromSize(sizeInput, axisPercent) {
            var nextScale = linkToggle.value ? { x: axisPercent, y: axisPercent } : { x: currentScale.x, y: currentScale.y };
            nextScale[sizeInput.scaleAxis] = axisPercent;
            return nextScale;
        }

        /**
         * 幅・高さを変えたとき：その向きの倍率を求める
         * @param {EditText} sizeInput - 変えた欄
         * @returns {void}
         */
        function onSizeEdited(sizeInput) {
            var axisPercent = getAxisPercentFromSize(sizeInput);
            if (axisPercent === null) return;
            lastEditedSizeInput = sizeInput;
            changeScaleByUser(buildScaleFromSize(sizeInput, axisPercent), sizeInput);
        }

        /**
         * 連動を切り替えたとき：連動にしたら高さの倍率を幅にそろえる
         * @returns {void}
         */
        function onLinkToggled() {
            if (!linkToggle.value || currentScale.x === currentScale.y) return;
            changeScaleByUser({ x: currentScale.x, y: currentScale.x }, null);
        }

        /**
         * 測った大きさを、今の幅・高さの表示に入れる。長さが無い向きの欄は無効にする
         * @param {{status: string, width: number, height: number}} measureResult - measureSelection() の結果
         * @returns {void}
         */
        function showMeasuredSize(measureResult) {
            var hasSize = measureResult.status === "OK";
            widthInput.sourceLength = hasSize ? measureResult.width : 0;
            heightInput.sourceLength = hasSize ? measureResult.height : 0;
            for (var i = 0; i < sizeInputs.length; i++) {
                var sourceLength = sizeInputs[i].sourceLength;
                sizeInputs[i].currentSizeText.text = (sourceLength > 0)
                    ? formatRounded(sourceLength / rulerUnit.pointsPerUnit, 2)
                    : getLabel("fieldLabel.noSize");
                /* 選択が無いときや長さ0（水平・垂直の直線など）は比率を求められない / no ratio without a selection or with zero length */
                setSteppedFieldEnabled(sizeInputs[i], sourceLength > 0);
            }
        }

        /**
         * 選択全体の大きさを測り直し、今の幅・高さの表示と目標の欄を更新する。
         * 最後に幅・高さの欄へ値を入れていたときは、その大きさになる倍率を求め直す
         * @returns {string} 計測の結果マーカー
         */
        function refreshSelectionSize() {
            var measureResult = measureSelection(); /* 測る前にプレビューは片付く / measuring clears the preview */
            isPreviewShown = false;
            showMeasuredSize(measureResult);
            var axisPercent = lastEditedSizeInput ? getAxisPercentFromSize(lastEditedSizeInput) : null;
            if (axisPercent !== null) {
                setCurrentScale(buildScaleFromSize(lastEditedSizeInput, axisPercent), lastEditedSizeInput);
            } else {
                syncScaleFields(null);
            }
            return measureResult.status;
        }

        /**
         * モードに合わせて、両端のオプションと基準点の有効／無効をそろえる。
         * 両端のオプションは［個別に］のときだけ、基準点は左右・上下とも固定したときは使わない。
         * 表示前は代入した value を読み戻せないので、値は呼び出し側から渡す
         * @param {boolean} isPerItem - ［個別に］なら true
         * @param {boolean} isFullyPinned - ［左右の両端］［上下の両端］がともに ON なら true
         * @returns {void}
         */
        function syncModeControls(isPerItem, isFullyPinned) {
            modeControls.pinOptionsGroup.enabled = isPerItem;
            setAnchorWidgetEnabled(anchorWidget, !(isPerItem && isFullyPinned));
        }

        /**
         * モード・両端のオプションを変えたとき
         * @returns {void}
         */
        function onModeOptionChanged() {
            syncModeControls(perItemRadio.value, pinHorizontalCheckbox.value && pinVerticalCheckbox.value);
            updatePreview();
        }

        /**
         * 各コントロールを既定の値にする
         * @returns {void}
         */
        function setDefaultValues() {
            currentScale = { x: DEFAULT_SCALE_PERCENT, y: DEFAULT_SCALE_PERCENT };
            lastEditedSizeInput = null;
            hasPendingScale = false;
            setLinkToggleValue(linkToggle, true);
            syncScaleFields(null);
            perItemRadio.value = true;
            asGroupRadio.value = false;
            pinHorizontalCheckbox.value = false;
            pinVerticalCheckbox.value = false;
            setAnchorWidgetValue(anchorWidget, DEFAULT_ANCHOR);
            syncModeControls(true, false);
            optionCheckboxes.scaleCornersCheckbox.value = true;
            optionCheckboxes.scaleStrokeCheckbox.value = true;
            optionCheckboxes.scalePatternCheckbox.value = true;
            optionCheckboxes.scaleGradientCheckbox.value = true;
            previewCheckbox.value = true;
            syncApplyButton(true);
            relativeStepCheckbox.value = false;
        }

        /**
         * 今の設定を、拡大・縮小の設定にまとめる
         * @returns {Object} scaleX / scaleY / mode / pinHorizontal / pinVertical / anchorIndex / scaleCorners / scaleStrokes / scalePatterns / scaleGradients
         */
        function collectScaleSettings() {
            return {
                scaleX: currentScale.x,
                scaleY: currentScale.y,
                mode: asGroupRadio.value ? "asGroup" : "perItem",
                pinHorizontal: perItemRadio.value && pinHorizontalCheckbox.value,
                pinVertical: perItemRadio.value && pinVerticalCheckbox.value,
                anchorIndex: getAnchorWidgetIndex(anchorWidget),
                scaleCorners: optionCheckboxes.scaleCornersCheckbox.value,
                scaleStrokes: optionCheckboxes.scaleStrokeCheckbox.value,
                scalePatterns: optionCheckboxes.scalePatternCheckbox.value,
                scaleGradients: optionCheckboxes.scaleGradientCheckbox.value
            };
        }

        /**
         * 委譲中でなければ処理を実行する（委譲の途中で別の操作が割り込まないようにする）
         * @param {Function} delegatedTask - 実行する処理
         * @returns {void}
         */
        function runWhenIdle(delegatedTask) {
            if (isBusy) return;
            isBusy = true;
            try {
                delegatedTask();
            } finally {
                isBusy = false;
            }
        }

        /**
         * プレビューを今の設定に合わせる。［プレビュー］が OFF、未適用の倍率が無い、100% のときは片付ける
         * @returns {void}
         */
        function renderPreview() {
            var shouldPreview = previewCheckbox.value && hasPendingScale && !isOriginalScale(currentScale);
            if (!shouldPreview) {
                if (isPreviewShown) clearScalePreview(false);
                isPreviewShown = false;
                return;
            }
            var previewResult = previewScaleOfSelection(collectScaleSettings());
            isPreviewShown = (previewResult === "OK");
            /* 選択が無いときは黙って見送る（入力のたびにアラートを出さない） / skip silently without a selection */
            if (previewResult.indexOf("ERR:") === 0) alertForResult(previewResult);
        }

        /**
         * 値やボタンの操作に合わせてプレビューを更新する
         * @returns {void}
         */
        function updatePreview() {
            runWhenIdle(renderPreview);
        }

        /**
         * 選択を測り直してから、今の設定で拡大・縮小する
         * @returns {boolean} 拡大・縮小した（100% で何もしなかったときも含む）なら true、選択が無い・失敗したなら false
         */
        function applyCurrentScale() {
            var measureStatus = refreshSelectionSize();
            if (measureStatus !== "OK") {
                alertForResult(measureStatus);
                return false;
            }
            if (isOriginalScale(currentScale)) return true;
            var applyResult = applyScaleToSelection(collectScaleSettings());
            hasPendingScale = false;
            alertForResult(applyResult);
            refreshSelectionSize();
            return applyResult.indexOf("OK") === 0;
        }

        /**
         * ［適用］／［確定］を押したとき。［確定］（［プレビュー］が ON）なら、拡大・縮小できたらそのまま閉じる
         * @returns {void}
         */
        function onApplyClick() {
            var isCommit = previewCheckbox.value;
            var isApplied = false;
            runWhenIdle(function () { isApplied = applyCurrentScale(); });
            if (isCommit && isApplied) closePalette();
        }

        /**
         * ［プレビュー］に合わせて、［適用］を［確定］に切り替える。
         * 表示前は代入した value を読み戻せないので、値は呼び出し側から渡す
         * @param {boolean} isPreviewOn - ［プレビュー］が ON なら true
         * @returns {void}
         */
        function syncApplyButton(isPreviewOn) {
            paletteButtons.btnApply.text = getLabel(isPreviewOn ? "button.commit" : "button.apply");
            paletteButtons.btnApply.helpTip = getLabel(isPreviewOn ? "tooltip.commit" : "tooltip.apply");
        }

        /**
         * プレビューを片付けてからパレットを閉じる
         * @returns {void}
         */
        function closePalette() {
            if (isPreviewShown) clearScalePreview(false);
            isPreviewShown = false;
            scalePalette.close();
        }

        asGroupRadio.onClick = onModeOptionChanged;
        perItemRadio.onClick = onModeOptionChanged;
        pinHorizontalCheckbox.onClick = onModeOptionChanged;
        pinVerticalCheckbox.onClick = onModeOptionChanged;
        for (var optionKey in optionCheckboxes) {
            if (optionCheckboxes.hasOwnProperty(optionKey)) optionCheckboxes[optionKey].onClick = updatePreview;
        }
        previewCheckbox.onClick = function () {
            syncApplyButton(previewCheckbox.value);
            updatePreview();
        };
        paletteButtons.btnReset.onClick = function () {
            setDefaultValues();
            updatePreview();
        };
        paletteButtons.btnApply.onClick = onApplyClick;
        paletteButtons.btnClose.onClick = closePalette;
        /* パレットは Esc で閉じないので割り当てる（入力中も効かせる）。keydown の中から完了待ちの BridgeTalk を
           送ると閉じられないため、プレビューの片付けは onClose の非同期側に任せる
           Palettes do not close on Esc by themselves; map it, also while typing. A blocking BridgeTalk inside keydown
           keeps the palette open, so leave the preview cleanup to the async path in onClose */
        addKeyShortcuts(scalePalette, {
            "Escape": { target: function () { scalePalette.close(); }, inFields: true }
        });

        /* パレットに戻ったときに選択を測り直し、プレビューを作り直す（Illustrator には常駐タイマーが無い）
           Re-measure the selection and rebuild the preview when the palette is activated (Illustrator has no idle timer) */
        scalePalette.onActivate = function () {
            runWhenIdle(function () {
                refreshSelectionSize();
                renderPreview();
            });
        };
        /* 閉じるときは DOM に触れない（触ると落ちる）。プレビューが残っていれば、待たずに片付けを頼む
           Do not touch the DOM while closing (it crashes); ask the main engine to clear a leftover preview without waiting */
        scalePalette.onClose = function () {
            if (isPreviewShown) clearScalePreview(true);
            isPreviewShown = false;
            $.global.__smartScalePalette = null;
        };

        setDefaultValues();
        refreshSelectionSize();
        return scalePalette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* パレット本体は $.global に持たせる（IIFE 内の var では GC される）/ Keep the palette on $.global (an IIFE-local var would be garbage-collected) */
    var scalePalette = createPalette();
    prepareDialogWindow(scalePalette, SCRIPT_NAME);
    $.global.__smartScalePalette = scalePalette;
    scalePalette.show();

})();
