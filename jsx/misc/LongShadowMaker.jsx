#target illustrator
#targetengine "LongShadowMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

距離・角度・スケールを指定して、選択したオブジェクトからロングシャドウを生成します。
プレビューを見ながらプリセットやオフセットで調整でき、生成後に「パスの単純化」を実行できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LongShadowMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n0be484dab7fc

### Overview

Generates a long shadow from the selected object using a distance, an angle and a scale.
Presets and an offset are adjusted with a live preview, and a Simplify Path pass can be run afterwards.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LongShadowMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LongShadowMaker";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LongShadowMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LongShadowMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0be484dab7fc"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author こじらせたクマー（オリジナルアイデア） / original idea
 * @discussion https://note.com/nice_lotus120/n/nf406fb3ae2b4
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 影の面を作るときに1セグメントを何点でサンプリングするか / Samples taken per curve segment */
    var CURVE_SAMPLE_COUNT = 8;

    /* 影の塗りの彩度（1=元のまま、0=無彩色）/ Saturation of the shadow fill (1 = as-is, 0 = gray) */
    var SHADOW_SATURATION_FACTOR = 0.7;

    /* オフセット初期値を「幅と高さの平均」の何分の1にするか / Divisor for the initial offset amount */
    var OFFSET_SIZE_DIVISOR = 20;

    /* プレビューの中間コピー（比率と不透明度）/ Intermediate preview copies (ratio and opacity) */
    var PREVIEW_STEPS = [
        { ratio: 1, opacity: 20 },
        { ratio: 0.75, opacity: 40 },
        { ratio: 0.5, opacity: 60 },
        { ratio: 0.25, opacity: 80 }
    ];

    /* プリセット（スケール／角度）。距離は変更しない / Presets (scale / angle); distance is left untouched */
    var PRESET_ITEMS = [
        "100% /  45°",
        "100% /  30°",
        "100% /  60°",
        " 50% /  90°",
        "  1% /  90°",
        "100% / 135°",
        "100% / 120°",
        "100% / 150°"
    ];

    /* 各スライダーの範囲 / Range of each slider */
    var DISTANCE_SLIDER_MIN_RANGE = 500;
    var ANGLE_MIN = -180;
    var ANGLE_MAX = 180;
    var SCALE_MIN = 1;
    var SCALE_MAX = 300;

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

    /* ダイアログ固有の寸法 / Dialog-specific sizes */
    var ROW_LABEL_WIDTH = 60;     /* 行ラベルの幅 / Width of a row label */
    var UNIT_LABEL_WIDTH = 24;    /* 単位ラベルの幅 / Width of a unit label */
    var NUMBER_FIELD_CHARS = 4;   /* 数値欄の文字数 / Character width of a number field */
    var SLIDER_WIDTH = 170;       /* スライダーの幅 / Width of a slider */
    var SIMPLIFY_ROW_MARGINS = [0, 10, 0, 0]; /* 単純化行の余白 / Margins of the simplify row */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ロングシャドウメーカー", en: "Long Shadow Maker" }
        },
        panel: {
            settings: { ja: "設定", en: "Settings" },
            offset: { ja: "オフセット", en: "Offset" }
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            join: { ja: "形状", en: "Join" },
            distance: { ja: "距離", en: "Distance" },
            angle: { ja: "角度", en: "Angle" },
            scale: { ja: "スケール", en: "Scale" }
        },
        radio: {
            joinMiter: { ja: "マイター", en: "Miter" },
            joinRound: { ja: "ラウンド", en: "Round" },
            joinBevel: { ja: "ベベル", en: "Bevel" }
        },
        checkbox: {
            simplify: { ja: "パスの単純化", en: "Simplify" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            preset: {
                ja: "スケールと角度の組み合わせをまとめて設定します。",
                en: "Sets the scale and angle together."
            },
            offsetEnabled: {
                ja: "影を作る前に、元の形を太らせます。",
                en: "Grows the original shape before the shadow is built."
            },
            offsetValue: { ja: "太らせる量です。", en: "How much to grow the shape." },
            join: {
                ja: "太らせたときの角の処理です。",
                en: "How corners are treated when the shape is grown."
            },
            distance: { ja: "影を伸ばす長さです。", en: "Length of the shadow." },
            angle: { ja: "影が伸びる向きです。", en: "Direction the shadow extends." },
            scale: {
                ja: "影の先端の大きさです。100%で元の形と同じ大きさになります。",
                en: "Size of the far end of the shadow. 100% matches the original shape."
            },
            simplify: {
                ja: "影のアンカーポイントを減らして、軽いパスにします。",
                en: "Reduces the number of anchor points in the shadow."
            },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original state."
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
        alert: {
            noDocument: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            },
            selectSingleShape: {
                ja: "閉パス（単一パス／複合パス／グループ）を1つだけ選択してください。",
                en: "Please select exactly one closed shape (path / compound path / group / text)."
            },
            selectClosedPath: {
                ja: "閉パスを選択してください。",
                en: "Please select a closed path."
            },
            selectClosedGroup: {
                ja: "閉パスのグループを選択してください。",
                en: "Please select a group that consists of closed paths."
            },
            groupBuildFailed: {
                ja: "グループから形状を作成できませんでした。閉パスのグループを選択してください。",
                en: "Could not build a shape from the group. Please select a group of closed paths."
            },
            notGroupItem: { ja: "GroupItem ではありません。", en: "This is not a GroupItem." },
            mergeResultMissing: {
                ja: "グループの合体結果を取得できませんでした。",
                en: "Could not retrieve the merged result from the group."
            },
            groupMergeError: {
                ja: "グループの合体中にエラーが発生しました: ",
                en: "An error occurred while merging the group: "
            },
            notTextFrame: { ja: "TextFrame ではありません。", en: "This is not a TextFrame." },
            outlineFailed: { ja: "アウトライン化に失敗しました。", en: "Failed to create outlines." },
            textMergeError: {
                ja: "テキストの合体中にエラーが発生しました: ",
                en: "An error occurred while processing the text: "
            }
        }
    };

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
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を検証し、ダイアログでロングシャドウを作る（処理に使う関数はこの中にまとめる）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;
        var currentSelection = doc.selection;

        /* 一時ベースの残骸を回収するためのタグ名 / Name tag used to sweep temporary base items */
        var TEMP_BASE_NAME = "__LongShadowTempBase__";

        // -----------------------------------------
        // 選択オブジェクトの検証 / Selection checks
        // -----------------------------------------

        /**
         * ロングシャドウの元にできる種類かどうかを判定する
         * @param {PageItem} pageItem - 判定対象
         * @returns {boolean} 対応している種類なら true
         */
        function isSupportedSourceItem(pageItem) {
            return !!pageItem && (
                pageItem.typename === "PathItem" ||
                pageItem.typename === "CompoundPathItem" ||
                pageItem.typename === "GroupItem" ||
                pageItem.typename === "TextFrame"
            );
        }

        /**
         * 渡されたパスがすべて閉じているかを判定する
         * @param {PathItem[]} subPaths - 判定対象のパス
         * @returns {boolean} 1つ以上あり、すべて閉じていれば true
         */
        function areAllPathsClosed(subPaths) {
            if (!subPaths || subPaths.length === 0) return false;
            for (var i = 0; i < subPaths.length; i++) {
                /* 生成直後のパスは closed を読めないことがある / closed may not be readable yet */
                try {
                    if (!subPaths[i].closed) return false;
                } catch (e) {
                    return false;
                }
            }
            return true;
        }

        if (currentSelection.length !== 1 || !isSupportedSourceItem(currentSelection[0])) {
            alert(getLabel('alert.selectSingleShape'));
            return;
        }

        var sourceItem = currentSelection[0];
        var sourceSubPaths = collectSelectionPathItems(sourceItem, { unique: false });

        /* Path/Compound はここで閉パス検証。Group/Text は実行時に一時パスへ変換して検証する
           Paths and compound paths are checked here; groups and text are checked after conversion */
        if (sourceItem.typename !== "GroupItem" && sourceItem.typename !== "TextFrame") {
            if (!areAllPathsClosed(sourceSubPaths)) {
                alert(getLabel('alert.selectClosedPath'));
                return;
            }
        }

        // -----------------------------------------
        // 一時ベースの後始末 / Temporary base cleanup
        // -----------------------------------------

        /**
         * プロパティに代入する。受け付けない種類・状態のアイテムでは何もしない
         * @param {Object} targetItem - 対象のアイテムやレイヤー
         * @param {string} propertyName - プロパティ名
         * @param {*} newValue - 代入する値
         * @returns {void}
         */
        function setPropertySafely(targetItem, propertyName, newValue) {
            /* 名前を持たない種類や、ロック中・無効になったアイテムでは代入が例外になる
               Assignment throws on item types without the property, or on locked / invalid items */
            try { targetItem[propertyName] = newValue; } catch (e) { }
        }

        /**
         * 親（グループ／レイヤー）を返す。読めないときは null
         * @param {Object} childItem - 対象のアイテム
         * @returns {Object|null} 親
         */
        function getParentSafely(childItem) {
            /* 無効になったアイテムでは parent を読めない / parent is unreadable on invalid items */
            try { return childItem.parent; } catch (e) { return null; }
        }

        /**
         * 一時ベースとその子要素に名前タグを付け外しする
         * @param {PageItem} pageItem - 対象アイテム
         * @param {string} tagName - 付ける名前（空文字でタグを外す）
         * @returns {void}
         */
        function setTempBaseTag(pageItem, tagName) {
            if (!pageItem) return;
            setPropertySafely(pageItem, "name", tagName);

            if (pageItem.typename === 'GroupItem') {
                for (var i = 0; i < pageItem.pageItems.length; i++) setTempBaseTag(pageItem.pageItems[i], tagName);
            } else if (pageItem.typename === 'CompoundPathItem') {
                for (var j = 0; j < pageItem.pathItems.length; j++) setPropertySafely(pageItem.pathItems[j], "name", tagName);
            }
        }

        /**
         * アイテム自身と親（グループ／レイヤー）のロックと非表示を解除する
         * @param {PageItem} pageItem - 対象アイテム
         * @returns {void}
         */
        function unlockItemAndAncestors(pageItem) {
            setPropertySafely(pageItem, "locked", false);
            setPropertySafely(pageItem, "hidden", false);

            var ancestor = getParentSafely(pageItem);
            var depthGuard = 0;
            while (ancestor && depthGuard++ < 50) {
                if (ancestor.typename === 'Layer') {
                    setPropertySafely(ancestor, "locked", false);
                    setPropertySafely(ancestor, "visible", true);
                    break;
                }
                if (ancestor.typename === 'GroupItem') {
                    setPropertySafely(ancestor, "locked", false);
                    setPropertySafely(ancestor, "hidden", false);
                }
                ancestor = getParentSafely(ancestor);
            }
        }

        /**
         * ロックや非表示を解除したうえでアイテムを確実に削除する
         * @param {PageItem} pageItem - 削除するアイテム
         * @returns {void}
         */
        function forceRemoveItem(pageItem) {
            try { if (!pageItem || !pageItem.isValid) return; } catch (e) { return; }

            unlockItemAndAncestors(pageItem);

            /* まず通常の削除 / first try a plain remove */
            try { pageItem.remove(); return; } catch (e) { }

            /* 削除できないときは選択してカット / fall back to selecting and clearing */
            try {
                doc.selection = null;
                pageItem.selected = true;
                app.executeMenuCommand('clear');
            } catch (e) { }
            doc.selection = null;

            try { if (pageItem && pageItem.isValid) pageItem.remove(); } catch (e) { }
        }

        /**
         * ドキュメント内に残った一時ベース（名前タグ付き）をすべて削除する
         * @returns {void}
         */
        function removeTempBaseItemsByName() {
            /* 一時ベースはグループであることが多いので先に走査 / temp bases are usually groups */
            var docGroupItems = doc.groupItems;
            for (var i = docGroupItems.length - 1; i >= 0; i--) {
                if (readItemName(docGroupItems[i]) === TEMP_BASE_NAME) forceRemoveItem(docGroupItems[i]);
            }

            /* 取りこぼし（グループ以外）を掃除 / sweep the non-group leftovers */
            var docPageItems = doc.pageItems;
            for (var j = docPageItems.length - 1; j >= 0; j--) {
                if (readItemName(docPageItems[j]) === TEMP_BASE_NAME) forceRemoveItem(docPageItems[j]);
            }
        }

        /**
         * アイテム名を安全に読み取る
         * @param {PageItem} pageItem - 対象アイテム
         * @returns {string} 読み取れない場合は空文字
         */
        function readItemName(pageItem) {
            try {
                if (!pageItem || !pageItem.isValid) return '';
                return pageItem.name;
            } catch (e) {
                return '';
            }
        }

        // -----------------------------------------
        // 一時ベースの生成 / Building the temporary base
        // -----------------------------------------

        /**
         * 一時生成物を控えて、あとでまとめて削除するための入れ物を作る
         * @returns {{track: function, disposeAll: function}} 追跡用と削除用の関数
         */
        function createTempItemTracker() {
            var temporaryItems = [];

            return {
                track: function (pageItem) {
                    if (pageItem) temporaryItems.push(pageItem);
                    return pageItem;
                },
                disposeAll: function () {
                    doc.selection = null;
                    for (var i = temporaryItems.length - 1; i >= 0; i--) {
                        try {
                            if (temporaryItems[i] && temporaryItems[i].isValid) temporaryItems[i].remove();
                        } catch (e) { }
                    }
                }
            };
        }

        /**
         * 選択中のアイテムを「合体」して分割・拡張し、単一のパスにする
         * @returns {void}
         */
        function executePathfinderAddAndExpand() {
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');
        }

        /**
         * GroupItem を一時的に単一パス／複合パスへ合体して返す（元グループは残す）
         * @param {GroupItem} groupItem - 合体するグループ
         * @returns {{item: PageItem, cleanup: function, ok: boolean, message: string}} 合体結果
         */
        function buildMergedItemFromGroup(groupItem) {
            var mergeResult = { item: null, cleanup: function () { }, ok: false, message: "" };

            if (!groupItem || groupItem.typename !== "GroupItem") {
                mergeResult.message = getLabel('alert.notGroupItem');
                return mergeResult;
            }

            var tempTracker = createTempItemTracker();
            mergeResult.cleanup = tempTracker.disposeAll;

            try {
                var duplicatedGroup = tempTracker.track(groupItem.duplicate());

                doc.selection = null;
                duplicatedGroup.selected = true;
                executePathfinderAddAndExpand();

                var expandedItems = doc.selection;
                doc.selection = null;
                trackAll(tempTracker, expandedItems);

                if (!expandedItems || expandedItems.length === 0) {
                    mergeResult.message = getLabel('alert.mergeResultMissing');
                    return mergeResult;
                }

                var mergedItem = expandedItems[0];

                /* 複数残った場合は一度グループ化して再合体 / regroup and merge again when several remain */
                if (expandedItems.length > 1) {
                    var regrouped = tempTracker.track(doc.groupItems.add());
                    for (var i = 0; i < expandedItems.length; i++) {
                        try { expandedItems[i].move(regrouped, ElementPlacement.PLACEATEND); } catch (e) { }
                    }
                    doc.selection = null;
                    regrouped.selected = true;
                    executePathfinderAddAndExpand();

                    var remergedItems = doc.selection;
                    doc.selection = null;
                    trackAll(tempTracker, remergedItems);

                    if (remergedItems && remergedItems.length > 0) mergedItem = remergedItems[0];
                }

                tempTracker.track(mergedItem);
                setTempBaseTag(mergedItem, TEMP_BASE_NAME);
                mergeResult.item = mergedItem;
                mergeResult.ok = true;
                return mergeResult;

            } catch (e) {
                mergeResult.message = getLabel('alert.groupMergeError') + e;
                return mergeResult;
            }
        }

        /**
         * 選択結果の配列をまとめて一時生成物として控える
         * @param {{track: function}} tempTracker - 一時生成物の入れ物
         * @param {PageItem[]} trackedItems - 控えるアイテム（null 可）
         * @returns {void}
         */
        function trackAll(tempTracker, trackedItems) {
            if (!trackedItems || !trackedItems.length) return;
            for (var i = 0; i < trackedItems.length; i++) tempTracker.track(trackedItems[i]);
        }

        /**
         * TextFrame を一時的にアウトライン化して返す（元テキストは残す）
         * @param {TextFrame} textFrame - アウトライン化するテキスト
         * @returns {{item: PageItem, cleanup: function, ok: boolean, message: string}} アウトライン化結果
         */
        function buildMergedItemFromText(textFrame) {
            var mergeResult = { item: null, cleanup: function () { }, ok: false, message: "" };

            if (!textFrame || textFrame.typename !== "TextFrame") {
                mergeResult.message = getLabel('alert.notTextFrame');
                return mergeResult;
            }

            var tempTracker = createTempItemTracker();
            mergeResult.cleanup = tempTracker.disposeAll;

            try {
                /* 複製に対してアウトライン化するので、オリジナルは残る
                   createOutline() consumes the duplicate, so the original survives */
                var duplicatedText = tempTracker.track(textFrame.duplicate());

                var outlinedItem = null;
                try { outlinedItem = duplicatedText.createOutline(); } catch (e) { outlinedItem = null; }

                if (!outlinedItem) {
                    mergeResult.message = getLabel('alert.outlineFailed');
                    return mergeResult;
                }

                /* 控えるのはアウトライン化で生成されたルートだけにする。子要素まで控えると、
                   後段で移動・合体された要素を巻き込んで生成済みの影が消えることがある
                   Track only the outlined root; tracking its children would delete the finished shadow */
                tempTracker.track(outlinedItem);

                /* 選択状態が残ると後段の処理に巻き込まれるので解除 / clear the selection before moving on */
                doc.selection = null;
                setPropertySafely(outlinedItem, "selected", false);

                setTempBaseTag(outlinedItem, TEMP_BASE_NAME);
                mergeResult.item = outlinedItem;
                mergeResult.ok = true;
                return mergeResult;

            } catch (e) {
                mergeResult.message = getLabel('alert.textMergeError') + e;
                return mergeResult;
            }
        }

        // -----------------------------------------
        // 色処理 / Color utilities
        // -----------------------------------------

        /**
         * 0〜1の範囲に丸める
         * @param {number} value - 丸める値
         * @returns {number} 0〜1に収めた値
         */
        function clamp01(value) {
            return Math.max(0, Math.min(1, value));
        }

        /**
         * 塗り色をRGBの成分に変換する（GrayColor と未対応の色は null）
         * @param {Color} fillColor - 元の塗り色
         * @returns {{red: number, green: number, blue: number}|null} 0〜255のRGB成分
         */
        function toRgbComponents(fillColor) {
            if (fillColor.typename === "RGBColor") {
                return { red: fillColor.red, green: fillColor.green, blue: fillColor.blue };
            }

            if (fillColor.typename === "CMYKColor") {
                /* 簡易 CMYK -> RGB（0-100 を 0-1 に換算）/ rough CMYK to RGB conversion */
                var cyan = fillColor.cyan / 100.0;
                var magenta = fillColor.magenta / 100.0;
                var yellow = fillColor.yellow / 100.0;
                var black = fillColor.black / 100.0;
                return {
                    red: 255 * (1 - cyan) * (1 - black),
                    green: 255 * (1 - magenta) * (1 - black),
                    blue: 255 * (1 - yellow) * (1 - black)
                };
            }

            return null;
        }

        /**
         * HSLの中間値から1チャンネル分の値を求める
         * @param {number} lowerBound - 下側の値
         * @param {number} upperBound - 上側の値
         * @param {number} hueFraction - 0〜1に正規化した色相
         * @returns {number} 0〜1のチャンネル値
         */
        function hueToChannel(lowerBound, upperBound, hueFraction) {
            var wrappedHue = hueFraction;
            if (wrappedHue < 0) wrappedHue += 1;
            if (wrappedHue > 1) wrappedHue -= 1;
            if (wrappedHue < 1 / 6) return lowerBound + (upperBound - lowerBound) * 6 * wrappedHue;
            if (wrappedHue < 1 / 2) return upperBound;
            if (wrappedHue < 2 / 3) return lowerBound + (upperBound - lowerBound) * (2 / 3 - wrappedHue) * 6;
            return lowerBound;
        }

        /**
         * 元の塗り色をもとに、彩度を下げた影用の色を作る
         * @param {Color} fillColor - 元の塗り色
         * @param {number} saturationFactor - 彩度の倍率（1=元のまま、0=無彩色）
         * @returns {Color|null} 影用の色。色が無い場合は null
         */
        function desaturateColorFromFill(fillColor, saturationFactor) {
            if (!fillColor) return null;
            if (saturationFactor === undefined || saturationFactor === null) saturationFactor = SHADOW_SATURATION_FACTOR;

            /* グレーはそのまま複製する / gray is copied as-is */
            if (fillColor.typename === "GrayColor") {
                var grayCopy = new GrayColor();
                grayCopy.gray = fillColor.gray;
                return grayCopy;
            }

            var rgbComponents = toRgbComponents(fillColor);
            /* 未対応（NoColor / グラデーション / パターン）はそのまま返す / unsupported fills pass through */
            if (!rgbComponents) return fillColor;

            var red = clamp01(rgbComponents.red / 255.0);
            var green = clamp01(rgbComponents.green / 255.0);
            var blue = clamp01(rgbComponents.blue / 255.0);

            var maxChannel = Math.max(red, green, blue);
            var minChannel = Math.min(red, green, blue);
            var chroma = maxChannel - minChannel;
            var lightness = (maxChannel + minChannel) / 2;
            var hue = 0;
            var saturation = 0;

            if (chroma !== 0) {
                saturation = chroma / (1 - Math.abs(2 * lightness - 1));
                switch (maxChannel) {
                    case red: hue = ((green - blue) / chroma) % 6; break;
                    case green: hue = ((blue - red) / chroma) + 2; break;
                    case blue: hue = ((red - green) / chroma) + 4; break;
                }
                hue = hue * 60;
                if (hue < 0) hue += 360;
            }

            saturation = clamp01(saturation * saturationFactor);

            var resultRed, resultGreen, resultBlue;
            if (saturation === 0) {
                resultRed = resultGreen = resultBlue = lightness;
            } else {
                var upperBound = lightness < 0.5 ? lightness * (1 + saturation) : (lightness + saturation - lightness * saturation);
                var lowerBound = 2 * lightness - upperBound;
                var hueFraction = hue / 360;
                resultRed = hueToChannel(lowerBound, upperBound, hueFraction + 1 / 3);
                resultGreen = hueToChannel(lowerBound, upperBound, hueFraction);
                resultBlue = hueToChannel(lowerBound, upperBound, hueFraction - 1 / 3);
            }

            var desaturatedColor = new RGBColor();
            desaturatedColor.red = Math.round(resultRed * 255);
            desaturatedColor.green = Math.round(resultGreen * 255);
            desaturatedColor.blue = Math.round(resultBlue * 255);
            return desaturatedColor;
        }

        // -----------------------------------------
        // 影の面を作る / Building the shadow faces
        // -----------------------------------------

        /**
         * 3次ベジェ曲線上の点を求める
         * @param {number[]} p0 - 始点
         * @param {number[]} p1 - 始点側の方向点
         * @param {number[]} p2 - 終点側の方向点
         * @param {number[]} p3 - 終点
         * @param {number} t - 0〜1の位置
         * @returns {number[]} [x, y] 座標
         */
        function bezierPoint(p0, p1, p2, p3, t) {
            var u = 1 - t;
            var tt = t * t;
            var uu = u * u;
            var uuu = uu * u;
            var ttt = tt * t;
            return [
                uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
                uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
            ];
        }

        /**
         * アイテムを複製し、中心を基準に拡大縮小してから移動する
         * @param {PageItem} pageItem - 複製するアイテム
         * @param {number} dx - 水平方向の移動量（pt）
         * @param {number} dy - 垂直方向の移動量（pt）
         * @param {number} scalePercent - 拡大率（100=等倍）
         * @returns {PageItem|null} 複製したアイテム
         */
        function duplicateWithOffsetAndScale(pageItem, dx, dy, scalePercent) {
            if (!pageItem) return null;

            var duplicatedItem = null;
            try { duplicatedItem = pageItem.duplicate(); } catch (e) { duplicatedItem = null; }
            if (!duplicatedItem) return null;

            resizeFromCenter(duplicatedItem, scalePercent);
            try { duplicatedItem.translate(dx, dy); } catch (e) { }
            return duplicatedItem;
        }

        /**
         * 中心を基準にアイテムを拡大縮小する
         * @param {PageItem} pageItem - 対象アイテム
         * @param {number} scalePercent - 拡大率（100=等倍なら何もしない）
         * @returns {void}
         */
        function resizeFromCenter(pageItem, scalePercent) {
            var scaleValue = isFinite(scalePercent) ? Number(scalePercent) : 100;
            if (scaleValue === 100) return;
            /* テキストや効果付きのアイテムで resize が失敗することがある / resize can fail on some items */
            try {
                pageItem.resize(scaleValue, scaleValue, true, true, true, true, true, Transformation.CENTER);
            } catch (e) { }
        }

        /**
         * 閉パスを曲線も含めて点列にする
         * @param {PathItem} pathItem - 点列にするパス
         * @param {number} curveSamplesPerSegment - 1セグメントあたりのサンプル数
         * @returns {number[][]} [x, y] の配列
         */
        function samplePathToPoints(pathItem, curveSamplesPerSegment) {
            if (!curveSamplesPerSegment) curveSamplesPerSegment = CURVE_SAMPLE_COUNT;
            var points = [];
            if (!pathItem) return points;

            /* 一時アイテムは走査中に無効化されることがある / temporary items can go invalid mid-scan */
            try {
                var pathPoints = pathItem.pathPoints;
                var pointCount = pathPoints.length;
                if (pointCount < 2) return points;

                for (var i = 0; i < pointCount; i++) {
                    var currentPoint = pathPoints[i];
                    var nextPoint = pathPoints[(i + 1) % pointCount];

                    for (var j = 0; j < curveSamplesPerSegment; j++) {
                        points.push(bezierPoint(
                            currentPoint.anchor,
                            currentPoint.rightDirection,
                            nextPoint.leftDirection,
                            nextPoint.anchor,
                            j / curveSamplesPerSegment
                        ));
                    }
                }
            } catch (e) { }

            return points;
        }

        /**
         * 重なり順を保ったまま閉じた PathItem だけを集める
         * @param {PageItem} pageItem - 走査対象
         * @returns {PathItem[]} 閉じた PathItem の配列
         */
        function collectClosedPaths(pageItem) {
            var closedPaths = [];
            /* グループ・複合パスの中を pageItems 順にたどる（見た目の重なり順に近い）/ follow pageItems order */
            var pathItems = collectSelectionPathItems(pageItem, { unique: false });
            for (var i = 0; i < pathItems.length; i++) pushIfClosed(closedPaths, pathItems[i]);
            return closedPaths;
        }

        /**
         * 閉じた PathItem のときだけ配列に加える
         * @param {PathItem[]} closedPaths - 追加先の配列
         * @param {PageItem} pathCandidate - 判定するアイテム
         * @returns {void}
         */
        function pushIfClosed(closedPaths, pathCandidate) {
            /* closed を読めない種類のアイテムが混ざる / closed is not readable on every item type */
            try {
                if (pathCandidate && pathCandidate.typename === 'PathItem' && pathCandidate.closed) {
                    closedPaths.push(pathCandidate);
                }
            } catch (e) { }
        }

        /**
         * 2つの点列の間を四角形の面で埋める
         * @param {GroupItem} parentGroup - 面を追加するグループ
         * @param {number[][]} nearPoints - 元の形の点列
         * @param {number[][]} farPoints - 影の先端側の点列
         * @param {Color} faceFill - 面の塗り色
         * @returns {PathItem[]} 生成した面
         */
        function buildSideFaces(parentGroup, nearPoints, farPoints, faceFill) {
            var createdFaces = [];
            if (!parentGroup || !nearPoints || !farPoints) return createdFaces;

            var pointCount = Math.min(nearPoints.length, farPoints.length);
            if (pointCount < 2) return createdFaces;

            for (var i = 0; i < pointCount; i++) {
                var nextIndex = (i + 1) % pointCount;
                var facePath = null;
                /* 極端に小さい面はパス生成が失敗することがある / very small faces can fail to build */
                try {
                    facePath = parentGroup.pathItems.add();
                    facePath.setEntirePath([nearPoints[i], farPoints[i], farPoints[nextIndex], nearPoints[nextIndex]]);
                    facePath.closed = true;
                    facePath.stroked = false;
                    facePath.filled = true;
                    if (faceFill) facePath.fillColor = faceFill;
                    createdFaces.push(facePath);
                } catch (e) {
                    try { if (facePath && facePath.isValid) facePath.remove(); } catch (err) { }
                }
            }
            return createdFaces;
        }

        /**
         * 元の形と複製した形の間を面でつないで影のグループを作る
         * @param {PageItem} baseItem - 影の元になる形
         * @param {number} dx - 水平方向の移動量（pt）
         * @param {number} dy - 垂直方向の移動量（pt）
         * @param {number} scalePercent - 先端の拡大率（100=等倍）
         * @returns {GroupItem|null} 生成した影のグループ
         */
        function buildShadowFaces(baseItem, dx, dy, scalePercent) {
            if (!baseItem) return null;

            var nearPaths = collectClosedPaths(baseItem);
            if (nearPaths.length === 0) return null;

            var farItem = duplicateWithOffsetAndScale(baseItem, dx, dy, scalePercent);
            if (!farItem) return null;

            var farPaths = collectClosedPaths(farItem);
            /* 数が違う場合は少ない方に合わせる（落とさない優先）/ pair up to the smaller count */
            var pairCount = Math.min(nearPaths.length, farPaths.length);
            if (pairCount === 0) {
                try { farItem.remove(); } catch (e) { }
                return null;
            }

            var shadowGroup = doc.groupItems.add();

            var shadowFill = null;
            /* 塗りを持たないパスがある / some paths carry no fill */
            try {
                shadowFill = desaturateColorFromFill(nearPaths[0].fillColor, SHADOW_SATURATION_FACTOR);
            } catch (e) {
                shadowFill = null;
            }

            for (var i = 0; i < pairCount; i++) {
                var nearPoints = samplePathToPoints(nearPaths[i], CURVE_SAMPLE_COUNT);
                var farPoints = samplePathToPoints(farPaths[i], CURVE_SAMPLE_COUNT);
                if (nearPoints.length < 2 || farPoints.length < 2) continue;
                buildSideFaces(shadowGroup, nearPoints, farPoints, shadowFill);
            }

            try { farItem.remove(); } catch (e) { }
            return shadowGroup;
        }

        // -----------------------------------------
        // オフセット効果 / Offset Path live effect
        // -----------------------------------------

        /**
         * 選択中の角の処理に対応する Offset Path のコードを返す
         * @returns {number} 0=ラウンド、1=ベベル、2=マイター
         */
        function getJoinCode() {
            if (joinRoundRadio.value) return 0;
            if (joinBevelRadio.value) return 1;
            return 2;
        }

        /**
         * Offset Path のライブ効果XMLを組み立てる
         * @param {number} offsetPt - オフセット量（pt）
         * @param {number} joinCode - 角の処理（0=ラウンド、1=ベベル、2=マイター）
         * @returns {string} applyEffect() に渡すXML
         */
        function buildOffsetEffectXML(offsetPt, joinCode) {
            /* mlim はマイター制限（既定4）、ofst は pt / mlim is the miter limit, ofst is in points */
            return '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetPt + ' I jntp ' + joinCode + ' "/></LiveEffect>';
        }

        /**
         * 生成した影にオフセット（ライブ効果）を適用する
         * @param {GroupItem} shadowGroup - 適用先のグループ
         * @param {boolean} isFinalRun - 本実行なら true（プレビューでは適用しない）
         * @returns {void}
         */
        function applyOffsetEffect(shadowGroup, isFinalRun) {
            if (!shadowGroup || !isFinalRun) return;
            if (!offsetCheckbox.value) return;

            var offsetPt = Number(offsetInput.text);
            if (isNaN(offsetPt)) offsetPt = 0;

            /* 0でもONなら適用する（結果が変わらないだけ）/ apply even at 0; it simply changes nothing */
            var joinCode = getJoinCode();
            try { shadowGroup.applyEffect(buildOffsetEffectXML(offsetPt, joinCode)); } catch (e) { }

            /* ラウンドのときは後処理として Pathfinder Merge を実行 / round joins need a merge pass */
            if (joinCode === 0) {
                try {
                    doc.selection = null;
                    shadowGroup.selected = true;
                    app.executeMenuCommand('Live Pathfinder Merge');
                } catch (e) { }
                doc.selection = null;
            }
        }

        // -----------------------------------------
        // パスの単純化 / Simplify paths
        // -----------------------------------------

        /**
         * 「パスの単純化」を実行する（Illustratorの仕様でダイアログが開く）
         * @param {PageItem} pageItem - 単純化するアイテム
         * @returns {void}
         */
        function simplifyPathsInItem(pageItem) {
            if (!pageItem || !simplifyCheckbox.value) return;

            /* 複合パスはまとめて1件にする / A compound path counts as one target */
            var simplifyTargets = collectSelectionPathItems(pageItem, { compoundPaths: "whole", unique: false });
            if (!simplifyTargets.length) return;

            doc.selection = null;
            for (var i = 0; i < simplifyTargets.length; i++) setPropertySafely(simplifyTargets[i], "selected", true);

            /* このメニューコマンドは単純化ダイアログを開く（Illustratorの制限）
               This menu command opens the Simplify dialog (Illustrator limitation) */
            try { app.executeMenuCommand("simplify menu item"); } catch (e) { }

            doc.selection = null;
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        var isPreviewing = false;
        var hasFinished = false;

        /* 数値欄とスライダーの同期関数。ダイアログ構築より前に用意する
           Sync handlers per number field; must exist before the dialog is built */
        var fieldSyncHandlers = [];

        /* ダイアログのコントロール（buildDialog() で作る）/ Dialog controls, created by buildDialog() */
        var shadowDialog, presetDropdown;
        var offsetCheckbox, offsetInput, offsetUnitLabel;
        var joinRowGroup, joinMiterRadio, joinRoundRadio, joinBevelRadio;
        var distanceInput, angleInput, scaleInput, simplifyCheckbox;
        var previewCheckbox, btnCancel, btnOK;

        /**
         * 並びの向きと子要素の揃えを指定したグループを追加する
         * @param {Object} parentGroup - 追加先のコンテナ
         * @param {string} orientation - "row" または "column"
         * @param {string[]} childAlignment - 子要素の揃え（alignChildren）
         * @returns {Group} 追加したグループ
         */
        function addLayoutGroup(parentGroup, orientation, childAlignment) {
            var layoutGroup = parentGroup.add("group");
            layoutGroup.orientation = orientation;
            layoutGroup.alignChildren = childAlignment;
            return layoutGroup;
        }

        /**
         * 右揃え・固定幅の行ラベルを追加する
         * @param {Object} parentGroup - 追加先のコンテナ
         * @param {string} labelPath - ラベルキー
         * @returns {StaticText} 追加したラベル
         */
        function addRowLabel(parentGroup, labelPath) {
            var rowLabel = parentGroup.add("statictext", undefined, getLabel(labelPath));
            rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
            rowLabel.justify = "right";
            return rowLabel;
        }

        /**
         * プリセットの行を作る（左右中央に配置）
         * @returns {void}
         */
        function addPresetRow() {
            var presetRowGroup = addLayoutGroup(shadowDialog, "row", ["center", "center"]);
            presetRowGroup.alignment = "center";

            var presetGroup = addLayoutGroup(presetRowGroup, "row", ["left", "center"]);
            presetGroup.add("statictext", undefined, getLabel('fieldLabel.preset'));

            presetDropdown = presetGroup.add("dropdownlist", undefined, PRESET_ITEMS);
            presetDropdown.selection = 0;
            presetDropdown.helpTip = getLabel('tooltip.preset');
        }

        /**
         * オフセットパネルの中身を作る（左=量／右=角の処理の2カラム）
         * @param {Panel} offsetPanel - 追加先のパネル
         * @returns {void}
         */
        function addOffsetControls(offsetPanel) {
            var offsetRowGroup = addLayoutGroup(offsetPanel, "row", ["fill", "top"]);

            var offsetValueGroup = addLayoutGroup(offsetRowGroup, "column", ["left", "top"]);
            var offsetInputGroup = addLayoutGroup(offsetValueGroup, "row", ["left", "center"]);

            offsetCheckbox = offsetInputGroup.add("checkbox", undefined, "");
            offsetCheckbox.helpTip = getLabel('tooltip.offsetEnabled');
            offsetCheckbox.value = false;

            /* ∧∨と入力欄は隙間0で突き合わせる。オフセットはプレビューを更新しない（従来どおり）
               / butt the stepper against the field; the offset does not refresh the preview (as before) */
            var offsetStepperGroup = offsetInputGroup.add("group");
            offsetStepperGroup.orientation = "row";
            offsetStepperGroup.alignChildren = ["left", "center"];
            offsetStepperGroup.spacing = 0;
            offsetStepperGroup.margins = 0;
            var offsetStepper = addStepper(offsetStepperGroup, function () { return offsetInput; }, { min: 0 });
            offsetInput = offsetStepperGroup.add("edittext", undefined, String(getInitialOffsetPt()));
            offsetInput.characters = NUMBER_FIELD_CHARS;
            offsetInput.helpTip = getLabel('tooltip.offsetValue');
            offsetInput.stepperGroup = offsetStepper;
            bindSteppedArrowKeys(offsetInput, offsetStepper);

            offsetUnitLabel = offsetInputGroup.add("statictext", undefined, "pt");

            var joinColumnGroup = addLayoutGroup(offsetRowGroup, "column", ["fill", "top"]);
            joinRowGroup = addLayoutGroup(joinColumnGroup, "row", ["left", "top"]);
            joinRowGroup.margins = [0, 0, 0, 0];

            var joinLabelGroup = addLayoutGroup(joinRowGroup, "column", ["left", "top"]);
            addRowLabel(joinLabelGroup, 'fieldLabel.join');

            var joinRadioGroup = addLayoutGroup(joinRowGroup, "column", ["left", "center"]);
            joinMiterRadio = joinRadioGroup.add("radiobutton", undefined, getLabel('radio.joinMiter'));
            joinRoundRadio = joinRadioGroup.add("radiobutton", undefined, getLabel('radio.joinRound'));
            joinBevelRadio = joinRadioGroup.add("radiobutton", undefined, getLabel('radio.joinBevel'));
            joinMiterRadio.helpTip = joinRoundRadio.helpTip = joinBevelRadio.helpTip = getLabel('tooltip.join');
            joinRoundRadio.value = true;
        }

        /**
         * 設定パネルの中身（距離・角度・スケール・パスの単純化）を作る
         * @param {Panel} settingsPanel - 追加先のパネル
         * @returns {void}
         */
        function addSettingsControls(settingsPanel) {
            /* 距離の初期値は元のオブジェクトの「幅＋高さ」/ Default distance is width plus height */
            var sourceBoundsPt = sourceItem.geometricBounds; /* [left, top, right, bottom] */
            var defaultDistancePt = Math.round((sourceBoundsPt[2] - sourceBoundsPt[0]) + (sourceBoundsPt[1] - sourceBoundsPt[3]));
            var maxDistancePt = Math.max(DISTANCE_SLIDER_MIN_RANGE, defaultDistancePt * 3);

            distanceInput = addSliderRow(settingsPanel, 'fieldLabel.distance', 'tooltip.distance',
                String(defaultDistancePt), "pt", 0, maxDistancePt, defaultDistancePt, false);
            angleInput = addSliderRow(settingsPanel, 'fieldLabel.angle', 'tooltip.angle',
                "45", "°", ANGLE_MIN, ANGLE_MAX, 45, true);
            scaleInput = addSliderRow(settingsPanel, 'fieldLabel.scale', 'tooltip.scale',
                "100", "%", SCALE_MIN, SCALE_MAX, 100, false);

            var simplifyGroup = addLayoutGroup(settingsPanel, "row", ["center", "center"]);
            simplifyGroup.alignment = "center";
            simplifyGroup.margins = SIMPLIFY_ROW_MARGINS;

            simplifyCheckbox = simplifyGroup.add("checkbox", undefined, getLabel('checkbox.simplify'));
            simplifyCheckbox.helpTip = getLabel('tooltip.simplify');
            simplifyCheckbox.value = true;
            simplifyCheckbox.alignment = "center";
        }

        /**
         * ダイアログを組み立てる
         * @returns {void}
         */
        function buildDialog() {
            shadowDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            setupWindow(shadowDialog);

            addPresetRow();

            var panelColumnGroup = addLayoutGroup(shadowDialog, "column", ["fill", "top"]);
            panelColumnGroup.alignment = "fill";

            var settingsPanel = panelColumnGroup.add("panel", undefined, getLabel('panel.settings'));
            setupPanel(settingsPanel);
            settingsPanel.alignChildren = "left";
            var offsetPanel = panelColumnGroup.add("panel", undefined, getLabel('panel.offset'));
            setupPanel(offsetPanel);
            offsetPanel.alignChildren = "left";

            addOffsetControls(offsetPanel);
            addSettingsControls(settingsPanel);

            /* ボタン行（左：プレビュー／右：キャンセル・OK）/ Button row: preview on the left, Cancel/OK on the right */
            var buttonRow = addButtonRow(shadowDialog);
            previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel('checkbox.preview'));
            previewCheckbox.helpTip = getLabel('tooltip.preview');
            previewCheckbox.value = true;

            btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
            btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
            btnOK.active = true;
        }

        /**
         * 元のオブジェクトのサイズからオフセットの初期値を求める
         * @returns {number} オフセットの初期値（pt）
         */
        function getInitialOffsetPt() {
            var sourceBounds = null;
            /* 種類によって geometricBounds を読めないことがある / geometricBounds is not always readable */
            try { sourceBounds = sourceItem.geometricBounds; } catch (e) { sourceBounds = null; }
            if (!sourceBounds) {
                try { sourceBounds = sourceItem.visibleBounds; } catch (e) { sourceBounds = null; }
            }
            if (!sourceBounds || sourceBounds.length !== 4) return 0;

            var widthPt = Math.abs(sourceBounds[2] - sourceBounds[0]);
            var heightPt = Math.abs(sourceBounds[1] - sourceBounds[3]);
            var averageSizePt = (widthPt + heightPt) / 2;
            var offsetBasePt = averageSizePt / OFFSET_SIZE_DIVISOR;

            return isNaN(offsetBasePt) ? 0 : Math.round(offsetBasePt);
        }

        /**
         * 「ラベル＋数値欄＋単位＋スライダー」の1行を作る
         * @param {Panel} parentPanel - 追加先のパネル
         * @param {string} labelPath - 行ラベルのラベルキー
         * @param {string} tooltipPath - ツールチップのラベルキー
         * @param {string} initialText - 数値欄の初期値
         * @param {string} unitText - 単位の表示
         * @param {number} minValue - スライダーの最小値
         * @param {number} maxValue - スライダーの最大値
         * @param {number} initialValue - スライダーの初期値
         * @param {boolean} allowNegative - 負の値を許可するか
         * @returns {EditText} 作成した数値欄
         */
        function addSliderRow(parentPanel, labelPath, tooltipPath, initialText, unitText, minValue, maxValue, initialValue, allowNegative) {
            var rowGroup = parentPanel.add("group");
            addRowLabel(rowGroup, labelPath);

            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var stepperFieldGroup = rowGroup.add("group");
            stepperFieldGroup.orientation = "row";
            stepperFieldGroup.alignChildren = ["left", "center"];
            stepperFieldGroup.spacing = 0;
            stepperFieldGroup.margins = 0;
            var numberInput;
            var rowStepper = addStepper(stepperFieldGroup, function () { return numberInput; }, {
                min: allowNegative ? undefined : 0,
                /* スライダーを追従させ、プレビューを更新する / keep the slider in sync and refresh the preview */
                onStep: function (steppedInput) {
                    syncFieldToSlider(steppedInput);
                    refreshPreviewIfEnabled();
                }
            });
            numberInput = stepperFieldGroup.add("edittext", undefined, initialText);
            numberInput.characters = NUMBER_FIELD_CHARS;
            numberInput.helpTip = getLabel(tooltipPath);
            bindSteppedArrowKeys(numberInput, rowStepper);

            var unitLabel = rowGroup.add("statictext", undefined, unitText);
            unitLabel.preferredSize.width = UNIT_LABEL_WIDTH;

            var rowSlider = rowGroup.add("slider", undefined, initialValue, minValue, maxValue);
            rowSlider.preferredSize.width = SLIDER_WIDTH;
            rowSlider.helpTip = getLabel(tooltipPath);

            bindSliderToInput(numberInput, rowSlider, minValue, maxValue);

            return numberInput;
        }

        /**
         * オフセットのコントロールをチェックボックスに合わせて有効／無効にする
         * @returns {void}
         */
        function updateOffsetControlsEnabled() {
            var isEnabled = !!offsetCheckbox.value;
            offsetInput.enabled = isEnabled;
            offsetInput.stepperGroup.enabled = isEnabled;
            redrawSteppersIn(offsetInput.stepperGroup);
            offsetUnitLabel.enabled = isEnabled;
            joinRowGroup.enabled = isEnabled;
            joinMiterRadio.enabled = isEnabled;
            joinRoundRadio.enabled = isEnabled;
            joinBevelRadio.enabled = isEnabled;
        }

        // -----------------------------------------
        // 数値欄の操作 / Number field behavior
        // -----------------------------------------

        /**
         * 数値欄の内容をスライダーへ反映する（登録済みの欄のみ）
         * @param {EditText} editText - 対象の数値欄
         * @returns {void}
         */
        function syncFieldToSlider(editText) {
            for (var i = 0; i < fieldSyncHandlers.length; i++) {
                if (fieldSyncHandlers[i].input === editText) {
                    fieldSyncHandlers[i].sync();
                    return;
                }
            }
        }

        /**
         * 値を最小値と最大値の間に丸める
         * @param {number} value - 丸める値
         * @param {number} minValue - 最小値
         * @param {number} maxValue - 最大値
         * @returns {number} 範囲内に収めた値
         */
        function clamp(value, minValue, maxValue) {
            return Math.max(minValue, Math.min(maxValue, value));
        }

        /**
         * 数値欄とスライダーを双方向に同期させる
         * @param {EditText} editText - 数値欄
         * @param {Slider} rowSlider - スライダー
         * @param {number} minValue - 最小値
         * @param {number} maxValue - 最大値
         * @returns {void}
         */
        function bindSliderToInput(editText, rowSlider, minValue, maxValue) {
            var isSyncing = false;

            /**
             * スライダーの値を数値欄へ反映し、プレビューを更新する
             * @returns {void}
             */
            function setInputFromSlider() {
                if (isSyncing) return;
                isSyncing = true;
                editText.text = String(Math.round(rowSlider.value));
                isSyncing = false;
                refreshPreviewIfEnabled();
            }

            /**
             * 数値欄の値を範囲内に丸めてスライダーへ反映する
             * @returns {void}
             */
            function setSliderFromInput() {
                if (isSyncing) return;
                isSyncing = true;
                var value = Number(editText.text);
                if (isNaN(value)) value = 0;
                rowSlider.value = Math.round(clamp(value, minValue, maxValue));
                isSyncing = false;
            }

            setSliderFromInput();
            fieldSyncHandlers.push({ input: editText, sync: setSliderFromInput });

            rowSlider.onChanging = setInputFromSlider;
            rowSlider.onChange = setInputFromSlider;

            editText.onChanging = function () {
                setSliderFromInput();
                refreshPreviewIfEnabled();
            };
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * プレビューがONのときだけ再描画する
         * @returns {void}
         */
        function refreshPreviewIfEnabled() {
            if (previewCheckbox && previewCheckbox.value) updatePreview();
        }

        /**
         * 直前のプレビューを取り消して描き直す
         * @returns {void}
         */
        function updatePreview() {
            undoPreview();
            if (previewCheckbox.value) {
                buildLongShadow(false);
                isPreviewing = true;
            }
            app.redraw();
        }

        /**
         * 表示中のプレビューを取り消す
         * @returns {void}
         */
        function undoPreview() {
            if (!isPreviewing) return;
            app.undo();
            isPreviewing = false;
        }

        /**
         * プレビュー用の半透明コピーを1つ作る
         * @param {PageItem} previewSource - 複製元のアイテム
         * @param {number} ratio - 影の先端までの比率（0〜1）
         * @param {number} opacityPercent - 不透明度（%）
         * @param {number} dx - 影の先端までの水平移動量（pt）
         * @param {number} dy - 影の先端までの垂直移動量（pt）
         * @param {number} scalePercent - 影の先端の拡大率
         * @returns {PageItem|null} 作成したコピー
         */
        function addPreviewCopy(previewSource, ratio, opacityPercent, dx, dy, scalePercent) {
            var previewCopy = null;
            try { previewCopy = previewSource.duplicate(); } catch (e) { previewCopy = null; }
            if (!previewCopy) return null;

            /* 一時ベースの複製はタグを外さないと後始末で消えてしまう
               Clear the temp tag, otherwise the cleanup sweep deletes the preview */
            setTempBaseTag(previewCopy, "");

            resizeFromCenter(previewCopy, 100 + (scalePercent - 100) * ratio);
            try { previewCopy.translate(dx * ratio, dy * ratio); } catch (e) { }
            try { previewCopy.move(sourceItem, ElementPlacement.PLACEBEFORE); } catch (e) { }
            setPropertySafely(previewCopy, "opacity", opacityPercent);
            return previewCopy;
        }

        // -----------------------------------------
        // 実行 / Execution
        // -----------------------------------------

        /**
         * スケールを実際に使う倍率へ変換する
         * @param {number} rawScalePercent - 入力されたスケール（%）
         * @returns {number} 実際に使う倍率（%）
         */
        function normalizeScalePercent(rawScalePercent) {
            var value = Number(rawScalePercent);
            if (isNaN(value)) return 100;
            /* プリセット「1%」は極小の影を作るため0.01%として扱う / the 1% preset runs as 0.01% */
            if (value === 1) return 0.01;
            return value;
        }

        /**
         * 入力欄から距離・角度・スケールを読み取る
         * @returns {{dx: number, dy: number, scalePercent: number}} 影の先端までの移動量と拡大率
         */
        function readShadowParameters() {
            var distancePt = parseFloat(distanceInput.text) || 0;
            var scalePercent = normalizeScalePercent(parseFloat(scaleInput.text));
            if (isNaN(scalePercent) || scalePercent <= 0) scalePercent = 100;

            /* 入力角度を反転して画面座標に合わせる / flip the angle to match screen coordinates */
            var angleRadians = -(parseFloat(angleInput.text) || 0) * Math.PI / 180;

            return {
                dx: distancePt * Math.cos(angleRadians),
                dy: distancePt * Math.sin(angleRadians),
                scalePercent: scalePercent
            };
        }

        /**
         * 影の元になる形を用意する（グループとテキストは一時パスへ変換する）
         * @param {boolean} isFinalRun - 本実行なら true
         * @returns {{item: PageItem, cleanup: function, ok: boolean, message: string}} 影の元になる形
         */
        function prepareShadowBaseItem(isFinalRun) {
            var sourceBase = { item: sourceItem, cleanup: function () { }, ok: true, message: "" };

            var mergedBase = null;
            if (sourceItem.typename === "GroupItem") {
                mergedBase = buildMergedItemFromGroup(sourceItem);
                if (!mergedBase.ok || !mergedBase.item) {
                    mergedBase.message = mergedBase.message || getLabel('alert.groupBuildFailed');
                    return mergedBase;
                }
            } else if (isFinalRun && sourceItem.typename === "TextFrame") {
                mergedBase = buildMergedItemFromText(sourceItem);
                if (!mergedBase.ok || !mergedBase.item) {
                    mergedBase.message = mergedBase.message || getLabel('alert.selectClosedPath');
                    return mergedBase;
                }
            }

            if (!mergedBase) return sourceBase;

            if (!areAllPathsClosed(collectSelectionPathItems(mergedBase.item, { unique: false }))) {
                mergedBase.ok = false;
                mergedBase.message = (sourceItem.typename === "GroupItem")
                    ? getLabel('alert.selectClosedGroup')
                    : getLabel('alert.selectClosedPath');
            }
            return mergedBase;
        }

        /**
         * 生成した影を元のオブジェクトの背面へ置く
         * @param {PageItem} shadowItem - 生成した影
         * @param {PageItem} originalItem - 元のオブジェクト
         * @returns {void}
         */
        function placeShadowBehindOriginal(shadowItem, originalItem) {
            if (!shadowItem || !originalItem) return;

            /* まずは元オブジェクトの直後（背面側）へ / first try placing it right behind the original */
            try {
                shadowItem.move(originalItem, ElementPlacement.PLACEAFTER);
                return;
            } catch (e) { }

            /* TextFrame などで move が失敗する場合は同一レイヤーの最背面へ
               When move fails, fall back to the back of the same layer */
            try {
                var ownerLayer = originalItem.layer;
                if (ownerLayer) {
                    shadowItem.move(ownerLayer, ElementPlacement.PLACEATBEGINNING);
                    return;
                }
            } catch (e) { }

            /* 最後の手段。これで見えなくなることもあるので最後に試す / last resort */
            try { shadowItem.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
        }

        /**
         * 影の面を合体させ、オフセットと単純化を適用したグループを返す
         * @param {GroupItem} shadowGroup - 面を集めたグループ
         * @param {boolean} isFinalRun - 本実行なら true
         * @returns {GroupItem} 仕上げた影のグループ
         */
        function finishShadowGroup(shadowGroup, isFinalRun) {
            /* パスの穴を潰すため、Merge → Add → Expand の順で実行 / merge, add, then expand */
            try {
                doc.selection = null;
                shadowGroup.selected = true;
                app.executeMenuCommand('Live Pathfinder Merge');
                app.executeMenuCommand('Live Pathfinder Add');
                app.executeMenuCommand('expandStyle');
            } catch (e) { }

            var mergedShadowItems = doc.selection;
            doc.selection = null;

            var shadowResultGroup = shadowGroup;
            if (mergedShadowItems && mergedShadowItems.length) {
                /* アクティブレイヤーがロックされているとグループを追加できない
                   groupItems.add() fails when the active layer is locked */
                try { shadowResultGroup = createGroupFrom(mergedShadowItems); } catch (e) { }
            }

            simplifyPathsInItem(shadowResultGroup);
            applyOffsetEffect(shadowResultGroup, isFinalRun);
            return shadowResultGroup;
        }

        /**
         * 渡された順序のままアイテムを新しいグループにまとめる
         * @param {PageItem[]} memberItems - まとめるアイテム
         * @returns {GroupItem} 作成したグループ
         */
        function createGroupFrom(memberItems) {
            var newGroup = doc.groupItems.add();
            for (var i = 0; i < memberItems.length; i++) {
                try { memberItems[i].move(newGroup, ElementPlacement.PLACEATEND); } catch (e) { }
            }
            return newGroup;
        }

        /**
         * ロングシャドウを生成する
         * @param {boolean} isFinalRun - 本実行なら true、プレビューなら false
         * @returns {void}
         */
        function buildLongShadow(isFinalRun) {
            var shadowParams = readShadowParameters();
            var shadowBase = prepareShadowBaseItem(isFinalRun);

            /**
             * 一時ベースと、名前タグの付いた残骸を削除する
             * @returns {void}
             */
            function cleanupTempBase() {
                shadowBase.cleanup();
                removeTempBaseItemsByName();
            }

            if (!shadowBase.ok) {
                cleanupTempBase();
                alert(shadowBase.message);
                return;
            }

            /* プレビューでは影を合体させず、半透明のコピーを並べるだけにする
               The preview stacks translucent copies instead of building the merged shadow */
            if (!isFinalRun) {
                for (var i = 0; i < PREVIEW_STEPS.length; i++) {
                    addPreviewCopy(shadowBase.item, PREVIEW_STEPS[i].ratio, PREVIEW_STEPS[i].opacity,
                        shadowParams.dx, shadowParams.dy, shadowParams.scalePercent);
                }
                doc.selection = null;
                cleanupTempBase();
                app.redraw();
                return;
            }

            var shadowGroup = buildShadowFaces(shadowBase.item, shadowParams.dx, shadowParams.dy, shadowParams.scalePercent);
            if (!shadowGroup) {
                cleanupTempBase();
                alert(getLabel('alert.selectClosedPath'));
                return;
            }

            var shadowResultGroup = finishShadowGroup(shadowGroup, isFinalRun);
            placeShadowBehindOriginal(shadowResultGroup, sourceItem);
            cleanupTempBase();

            /* 単一パスから作ったときは、生成した影を選択状態にする
               Select the generated shadow when the source was a single path */
            doc.selection = null;
            try {
                if (sourceItem.typename === "PathItem" && shadowResultGroup.isValid) {
                    doc.selection = [shadowResultGroup];
                }
            } catch (e) { }

            app.redraw();
        }

        // -----------------------------------------
        // イベントリスナー / Event listeners
        // -----------------------------------------

        /**
         * ダイアログのイベントハンドラーを設定する
         * @returns {void}
         */
        function bindDialogEvents() {
            offsetCheckbox.onClick = updateOffsetControlsEnabled;

            /* プリセット選択時：スケールと角度に反映（距離は変更しない）
               Presets set the scale and the angle; the distance is left as it is */
            presetDropdown.onChange = function () {
                if (!presetDropdown.selection) return;

                var matched = String(presetDropdown.selection.text)
                    .match(/^\s*(-?\d+(?:\.\d+)?)\s*%\s*\/\s*(-?\d+(?:\.\d+)?)\s*°\s*$/);
                if (!matched) return;

                scaleInput.text = matched[1];
                angleInput.text = matched[2];
                syncFieldToSlider(scaleInput);
                syncFieldToSlider(angleInput);

                refreshPreviewIfEnabled();
            };

            simplifyCheckbox.onClick = refreshPreviewIfEnabled;
            previewCheckbox.onClick = updatePreview;

            btnOK.onClick = function () {
                /* プレビューを消してから確定実行 / drop the preview before the real run */
                undoPreview();
                hasFinished = true;
                buildLongShadow(true);
                shadowDialog.close();
            };

            btnCancel.onClick = function () {
                undoPreview();
                hasFinished = true;
                removeTempBaseItemsByName();
                shadowDialog.close();
            };

            /* 閉じるボタンで閉じられたときもプレビューを後片付けする
               Clean the preview up when the dialog is dismissed by its close box */
            shadowDialog.onClose = function () {
                if (hasFinished) return;
                undoPreview();
                removeTempBaseItemsByName();
                app.redraw();
            };
        }

        buildDialog();
        updateOffsetControlsEnabled();
        bindDialogEvents();

        /* プレビューで選択が変わる前に選択範囲を測る / Measure the selection before the preview changes it */
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(shadowDialog, SCRIPT_NAME);

        /* ダイアログを開いた時点でプレビューを表示 / show the preview as the dialog opens */
        refreshPreviewIfEnabled();
        shadowDialog.show();
    }

    main();

})();
