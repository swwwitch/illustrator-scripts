#target illustrator
#targetengine "AiAnchorPointMarkerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトのすべてのアンカーポイントに、マーカー（自動生成の正方形／最前面オブジェクトの複製／シンボルのインスタンス）を配置します。
ダイアログを閉じずに、ライブプレビューで仕上がりを確かめながら設定を調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAnchorPointMarker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n757f8802dc4b

### Overview

Places a marker on every anchor point of the selected objects — a generated square, a copy of the frontmost object, or a symbol instance.
The settings are adjusted with a live preview, without closing the dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAnchorPointMarker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAnchorPointMarker";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAnchorPointMarker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAnchorPointMarker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n757f8802dc4b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /**
     * 追加するオブジェクトの種別 / Kind of object to place
     * @type {Object}
     */
    var OBJECT_SOURCE = {
        autoGenerate: "autoGenerate", /* 正方形を自動生成 / Auto-generated square */
        frontObject: "frontObject",   /* 最前面のオブジェクト / Frontmost object */
        symbol: "symbol"              /* シンボル / Symbol */
    };

    var PREVIEW_LAYER_NAME = "__ANCHOR_MARKER_PREVIEW__"; /* プレビュー専用レイヤー名 / Preview-only layer name */
    var ANCHOR_LAYER_NAME = "_anchorpoint";               /* マーカー移動先レイヤー名 / Destination layer for markers */
    var SQUARE_SYMBOL_NAME = "アンカーポイント";           /* 自動生成シンボルの名前 / Name for the generated symbol */

    /**
     * ダイアログの初期値 / Initial dialog values
     * @type {Object}
     */
    var DEFAULTS = {
        objectSource: OBJECT_SOURCE.autoGenerate, /* 追加するオブジェクトの種類 / Kind of object to add */
        squareSize: 6,                            /* 正方形の一辺（pt） / Square edge size (pt) */
        squareColor: { r: 79, g: 128, b: 255 },   /* 塗り色（RGB） / Fill color (RGB) */
        symbolize: true,                          /* 自動生成の正方形をシンボル化して配置 / Place squares as symbol instances */
        moveToLayer: false,                       /* 配置後のマーカーを専用レイヤーへ移動 / Move markers to a dedicated layer */
        groupItems: true,                         /* 配置後のマーカーを1つのグループにまとめる / Group the placed markers */
        scalePercent: 100,                        /* 最前面オブジェクト／シンボルの拡大縮小率（%） / Scale for frontmost object / symbol (%) */
        registrationIndex: 4                      /* 基準点 0..8（行優先, 4=中央）/ Registration point 0..8 row-major (4=center) */
    };

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

    var WIDGET_SIZE    = 56;                 /* 9軸ウィジェットの一辺（px）/ edge size of the 9-axis widget */
    var SWATCH_SIZE    = 20;                 /* カラースウォッチの一辺（px）/ edge size of the color swatch */

    /**
     * 縦並びのカラムグループを追加します。
     * @param {Window|Group|Panel} parent - 追加先
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parent, spacing) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["left", "center"];
        columnGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        return columnGroup;
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

    /**
     * 日英ラベル定義（カテゴリ別） / Japanese-English label definitions (by category)
     * getLabel("dialog.title") のようにドット区切りで参照する。短い文言は1行で記述。
     * @type {Object}
     */
    var LABELS = {
        dialog: {
            title: { ja: "アンカーポイントに複製", en: "Duplicate to Anchor Points" }
        },
        colorPicker: {
            title: { ja: "カラーを選択", en: "Choose Color" }
        },
        panel: {
            objectSource: { ja: "追加するオブジェクト", en: "Object to Add" },
            anchorPoint: { ja: "アンカーポイント", en: "Anchor Point" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            autoGenerate: { ja: "アンカーポイントを自動生成", en: "Auto-generate square" },
            frontObject: { ja: "最前面のオブジェクト", en: "Frontmost object" },
            symbol: { ja: "シンボル", en: "Symbol" }
        },
        label: {
            squareSize: { ja: "大きさ", en: "Size" },
            squareColor: { ja: "カラー", en: "Color" },
            scale: { ja: "スケール", en: "Scale" },
            unit: { ja: "pt", en: "pt" },
            percent: { ja: "%", en: "%" }
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
            autoGenerate: { ja: "指定した大きさ・カラーの正方形を生成して各アンカーポイントに配置します。", en: "Generate a square of the given size/color and place one at each anchor point." },
            frontObject: { ja: "最前面のオブジェクトを複製して各アンカーポイントに配置します。", en: "Duplicate the frontmost object and place it at each anchor point." },
            symbol: { ja: "選択したシンボルのインスタンスを各アンカーポイントに配置します。", en: "Place an instance of the selected symbol at each anchor point." },
            symbolDropdown: { ja: "配置するシンボルを選びます。", en: "Choose the symbol to place." },
            squareSize: { ja: "正方形の一辺（pt）。∧∨・↑↓：次の整数へ　Shift：10の倍数へ　Option：±0.1", en: "Square size (pt). ∧∨ / ↑↓: to the next whole number, Shift: to multiples of 10, Option: ±0.1" },
            squareColor: { ja: "正方形の塗り色。［選択...］でカラーを変更します。", en: "Fill color of the square. Click Choose... to change it." },
            symbolize: { ja: "生成した正方形をシンボル（アンカーポイント）として登録し、インスタンスで配置します。", en: "Register the generated square as a symbol and place instances." },
            moveToLayer: { ja: "配置後のマーカーを「_anchorpoint」レイヤーへ移動します。", en: "Move the placed markers to the \"_anchorpoint\" layer." },
            group: { ja: "配置後のマーカーを1つのグループにまとめます。", en: "Group the placed markers into a single group." },
            scale: { ja: "最前面オブジェクト・シンボルの拡大縮小率（%）。∧∨・↑↓：次の整数へ　Shift：10の倍数へ　Option：±0.1", en: "Scale for the frontmost object / symbol (%). ∧∨ / ↑↓: to the next whole number, Shift: to multiples of 10, Option: ±0.1" },
            registration: { ja: "基準点。マーカーのどの位置をアンカーポイントに合わせるかを選びます（中央が既定）。", en: "Registration point: which part of the marker aligns to the anchor (center by default)." }
        },
        checkbox: {
            symbolize: { ja: "シンボル化", en: "Symbolize" },
            moveToLayer: { ja: "レイヤーに移動", en: "Move to layer" },
            group: { ja: "グループ化", en: "Group" }
        },
        button: {
            chooseColor: { ja: "選択...", en: "Choose..." },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noAnchorPoints: { ja: "選択オブジェクトにアンカーポイントがありません。\nパスを含むオブジェクトを選択してください。", en: "The selection has no anchor points.\nSelect an object that contains paths." },
            invalidSize: { ja: "大きさには 0 より大きい数値を入力してください。", en: "Enter a size greater than 0." },
            invalidScale: { ja: "スケールには 0 より大きい数値を入力してください。", en: "Enter a scale greater than 0." },
            noSymbolChosen: { ja: "配置するシンボルを選択してください。", en: "Please choose a symbol to place." }
        }
    };

    // =========================================
    // 状態・UI配色 / State & UI colors
    // =========================================
    var registrationIndex = DEFAULTS.registrationIndex; /* 基準点 0..8（行優先, 4=中央）/ Registration point 0..8 (row-major, 4=center) */
    var activePreviewLayer = null; /* 生成したプレビュー用レイヤーの参照（名前でなく参照で管理）/ Reference to the preview layer we created */
    var cachedFrontmostItem;       /* 複製元の走査結果。undefined＝未計算, null＝該当なし / Cached duplication source (undefined = not resolved yet) */

    var lightUI = isLightUI();
    var widgetForeColor = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1]; /* セル・ケイ線 / Cells and rules */
    var widgetDimColor = lightUI ? [0.70, 0.70, 0.70, 1] : [0.42, 0.42, 0.42, 1];  /* 無効時 / Disabled */

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
    // メイン / Main
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var selectedItems = doc.selection;

    if (!selectedItems || selectedItems.length === 0) {
        alert(getLabel("alert.noSelection"));
        return;
    }

    if (collectAllAnchorPoints(selectedItems).length === 0) {
        alert(getLabel("alert.noAnchorPoints"));
        return; /* パスが無い（テキスト・画像のみ等）/ No path anchors (e.g. text/image only) */
    }

    toggleCanvasHelpers(); /* 実行時：エッジ等を隠す / On run: hide edges & annotator */
    try {
        placeMarkers();
    } finally {
        // 途中で例外が出ても、プレビューとエッジ表示は必ず元に戻す
        // Always clean up the preview and restore edges, even if something throws
        removePreviewLayer();
        toggleCanvasHelpers();
    }

    /**
     * 設定ダイアログを表示し、確定した設定でマーカーを配置します。
     * キャンセル時・配置元を用意できなかったときは、何もせずに戻ります。
     * @returns {void}
     */
    function placeMarkers() {
        var userSettings = showSettingsDialog();
        if (!userSettings) return; /* キャンセル / Cancelled */

        var placeMarkerAtPoint = buildMarkerPlacer(userSettings, false);
        if (!placeMarkerAtPoint) return; /* 配置元が用意できなかった / No placement source available */

        var placementPoints = resolveAnchorPoints(userSettings.objectSource);
        var placedItems = [];
        for (var i = 0; i < placementPoints.length; i++) {
            placedItems.push(placeMarkerAtPoint(placementPoints[i][0], placementPoints[i][1]));
        }

        // グループ化：まとめてから（必要なら）レイヤーへ移す / Group first, then move to the layer if requested
        var resultItems = (userSettings.groupItems && placedItems.length > 0) ? [groupPlacedItems(placedItems)] : placedItems;

        if (userSettings.moveToLayer) {
            var targetLayer = getOrCreateLayer(ANCHOR_LAYER_NAME);
            for (var m = 0; m < resultItems.length; m++) {
                moveItemToLayer(resultItems[m], targetLayer);
            }
        }

        // 配置物を選択状態にして直後の移動・整列をしやすく / Select the placed items for easy move / align
        if (resultItems.length > 0) {
            doc.selection = resultItems;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /**
     * マーカー配置の設定 / Marker placement settings
     * @typedef {Object} MarkerSettings
     * @property {string} objectSource - 追加するオブジェクトの種別（OBJECT_SOURCE の値）
     * @property {number} squareSize - 正方形の一辺（pt）
     * @property {Object} squareColor - 塗り色 {r, g, b}
     * @property {boolean} symbolize - 自動生成の正方形をシンボル化するか
     * @property {number} symbolIndex - 配置するシンボルのインデックス（未選択は -1）
     * @property {boolean} moveToLayer - 専用レイヤーへ移動するか
     * @property {boolean} groupItems - 1つのグループにまとめるか
     * @property {number} scalePercent - 拡大縮小率（%）
     * @property {number} registrationIndex - 基準点 0..8（行優先, 4=中央）
     */

    /**
     * 設定ダイアログを表示し、確定した設定を返します。
     * @returns {MarkerSettings|null} 確定した設定（キャンセル時は null）
     */
    function showSettingsDialog() {
        var pickedColor = {
            r: DEFAULTS.squareColor.r,
            g: DEFAULTS.squareColor.g,
            b: DEFAULTS.squareColor.b
        };

        var symbolNames = getSymbolNames();
        var hasSymbols = symbolNames.length > 0;

        var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow);

        // --- 追加するオブジェクト パネル / Object to Add panel ---
        var objectSourcePanel = dialogWindow.add("panel", undefined, getLabel("panel.objectSource"));
        setupPanel(objectSourcePanel, 6);

        var autoGenerateRadio = objectSourcePanel.add("radiobutton", undefined, getLabel("radio.autoGenerate"));
        autoGenerateRadio.helpTip = getLabel("tooltip.autoGenerate");
        var frontObjectRadio = objectSourcePanel.add("radiobutton", undefined, getLabel("radio.frontObject"));
        frontObjectRadio.helpTip = getLabel("tooltip.frontObject");

        var symbolSourceGroup = objectSourcePanel.add("group");
        setupRow(symbolSourceGroup, "left", 8);
        var symbolRadio = symbolSourceGroup.add("radiobutton", undefined, labelText("radio.symbol"));
        symbolRadio.helpTip = getLabel("tooltip.symbol");
        var symbolDropdown = symbolSourceGroup.add("dropdownlist", undefined, symbolNames);
        symbolDropdown.helpTip = getLabel("tooltip.symbolDropdown");
        symbolDropdown.preferredSize.width = 130;
        if (hasSymbols) {
            symbolDropdown.selection = 0;
        }

        // 選択が1つ（単一オブジェクト／グループ1つ）だと複製元を除くと配置先が無いのでディム
        var canUseFrontObject = (selectedItems.length > 1);

        autoGenerateRadio.value = (DEFAULTS.objectSource === OBJECT_SOURCE.autoGenerate);
        frontObjectRadio.value = (DEFAULTS.objectSource === OBJECT_SOURCE.frontObject) && canUseFrontObject;
        symbolRadio.value = (DEFAULTS.objectSource === OBJECT_SOURCE.symbol) && hasSymbols;
        frontObjectRadio.enabled = canUseFrontObject;
        symbolRadio.enabled = hasSymbols;

        // --- アンカーポイント パネル（2カラム：左＝大きさ・シンボル化 / 右＝カラー）---
        var anchorPointPanel = dialogWindow.add("panel", undefined, getLabel("panel.anchorPoint"));
        setupPanel(anchorPointPanel);
        /* 2カラムを横に並べる / Lay the two columns side by side */
        anchorPointPanel.orientation = "row";
        anchorPointPanel.alignChildren = ["left", "top"];
        anchorPointPanel.spacing = COLUMN_SPACING;

        // 左カラム：大きさ・シンボル化 / Left column: size, symbolize
        var anchorLeft = addColumnGroup(anchorPointPanel);

        var sizeInput = addLabeledField(anchorLeft, labelText("label.squareSize"), DEFAULTS.squareSize, 2, getLabel("label.unit"), renderPreview, undefined, 0.1);
        sizeInput.helpTip = getLabel("tooltip.squareSize"); /* ∧∨・↑↓の操作説明 / How the steppers and arrow keys work */

        var symbolizeCheckbox = anchorLeft.add("checkbox", undefined, getLabel("checkbox.symbolize"));
        symbolizeCheckbox.value = DEFAULTS.symbolize;
        symbolizeCheckbox.helpTip = getLabel("tooltip.symbolize");

        // 右カラム：カラー（2行：カラー■ / 選択...）/ Right column: color (2 rows)
        // OSのカラーパレットはモーダルを壊すため、自前のRGBダイアログを使う
        var anchorRight = addColumnGroup(anchorPointPanel, 0);

        var colorRow = anchorRight.add("group");
        setupRow(colorRow, "left", 8);
        colorRow.add("statictext", undefined, labelText("label.squareColor"));
        var colorSwatch = colorRow.add("panel");
        colorSwatch.preferredSize = [SWATCH_SIZE, SWATCH_SIZE]; /* 正方形・高さは短いまま / Square, keep the short height */
        colorSwatch.helpTip = getLabel("tooltip.squareColor");
        colorSwatch.onDraw = makeSwatchDrawer(colorSwatch, pickedColor);

        // 選択ボタンは上マージンを持つグループで包む（コントロール直接の margins は効かない環境があるため）
        var chooseColorWrap = anchorRight.add("group");
        chooseColorWrap.margins = [0, 10, 0, 0]; /* 上にマージン10 / 10px top margin */
        var chooseColorButton = chooseColorWrap.add("button", undefined, getLabel("button.chooseColor"));
        chooseColorButton.helpTip = getLabel("tooltip.squareColor");
        chooseColorButton.onClick = function () {
            var chosenColor = chooseRgbColor(pickedColor);
            if (chosenColor) {
                pickedColor.r = chosenColor.r;
                pickedColor.g = chosenColor.g;
                pickedColor.b = chosenColor.b;
                redrawControl(colorSwatch);
                renderPreview();
            }
        };

        // --- オプション パネル（2カラム：左＝3設定 / 右＝9軸）/ Options panel (2 columns) ---
        var optionsPanel = dialogWindow.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel);
        /* 2カラムを横に並べる / Lay the two columns side by side */
        optionsPanel.orientation = "row";
        optionsPanel.alignChildren = ["left", "top"];
        optionsPanel.spacing = COLUMN_SPACING;

        // 左カラム：スケール・レイヤーに移動・グループ化 / Left column: scale, move-to-layer, group
        var optionsLeft = addColumnGroup(optionsPanel);

        var scaleInput = addLabeledField(optionsLeft, labelText("label.scale"), DEFAULTS.scalePercent, 3, getLabel("label.percent"), renderPreview);
        scaleInput.helpTip = getLabel("tooltip.scale");

        var moveToLayerCheckbox = optionsLeft.add("checkbox", undefined, getLabel("checkbox.moveToLayer"));
        moveToLayerCheckbox.value = DEFAULTS.moveToLayer;
        moveToLayerCheckbox.helpTip = getLabel("tooltip.moveToLayer");

        var groupCheckbox = optionsLeft.add("checkbox", undefined, getLabel("checkbox.group"));
        groupCheckbox.value = DEFAULTS.groupItems;
        groupCheckbox.helpTip = getLabel("tooltip.group");

        // 右カラム：9軸（基準点）を天地左右中央に / Right column: 9-axis widget, centered both ways
        var optionsRight = addColumnGroup(optionsPanel);
        optionsRight.alignChildren = ["center", "center"];
        optionsRight.alignment = ["center", "center"]; /* 左カラムの高さに対して天地中央 / Vertically center against the left column */
        var registrationWidget = addRegistrationWidget(optionsRight, renderPreview);

        /**
         * 現在選択されている「追加するオブジェクト」の種別を返します。
         * @returns {string} OBJECT_SOURCE の値
         */
        function getChosenSource() {
            if (frontObjectRadio.value) return OBJECT_SOURCE.frontObject;
            if (symbolRadio.value) return OBJECT_SOURCE.symbol;
            return OBJECT_SOURCE.autoGenerate;
        }

        /**
         * 追加するオブジェクトの選択に応じて、各コントロールの有効／無効を切り替えます。
         * @returns {void}
         */
        function syncControlState() {
            var isAutoGenerate = autoGenerateRadio.value;
            sizeInput.enabled = isAutoGenerate;
            sizeInput.stepperGroup.enabled = isAutoGenerate;
            chooseColorButton.enabled = isAutoGenerate;
            symbolizeCheckbox.enabled = isAutoGenerate;
            // 自動生成では基準点を中央へ戻す / Reset registration to center in auto-generate
            if (isAutoGenerate) {
                registrationIndex = DEFAULTS.registrationIndex;
            }
            // 自動生成時はスケールと9軸のみディム（レイヤー移動・グループ化は常時有効）
            scaleInput.fieldRow.enabled = !isAutoGenerate;
            redrawSteppersIn(sizeInput.stepperGroup); /* ∧∨のディム表示を切り替える / update the stepper dimming */
            redrawSteppersIn(scaleInput.fieldRow);
            registrationWidget.enabled = !isAutoGenerate;
            redrawControl(registrationWidget); /* 自作描画なので色を更新 / Redraw the custom widget */
            symbolDropdown.enabled = symbolRadio.value;
        }

        /**
         * ラジオが別コンテナに分かれているため、排他選択を手動で担保します。
         * @param {RadioButton} selectedRadio - 選択されたラジオボタン
         * @returns {void}
         */
        function selectObjectSource(selectedRadio) {
            autoGenerateRadio.value = (selectedRadio === autoGenerateRadio);
            frontObjectRadio.value = (selectedRadio === frontObjectRadio);
            symbolRadio.value = (selectedRadio === symbolRadio);
            syncControlState();
            renderPreview();
        }
        autoGenerateRadio.onClick = function () { selectObjectSource(autoGenerateRadio); };
        frontObjectRadio.onClick = function () { selectObjectSource(frontObjectRadio); };
        symbolRadio.onClick = function () { selectObjectSource(symbolRadio); };
        symbolizeCheckbox.onClick = renderPreview;
        symbolDropdown.onChange = renderPreview;
        syncControlState();

        /**
         * 現在のコントロール値から設定オブジェクトを読み取ります（プレビュー・本適用の共通ソース）。
         * @returns {MarkerSettings} 現在の設定
         */
        function readCurrentSettings() {
            var sizeValue = parseFloat(sizeInput.text);
            var scaleValue = parseFloat(scaleInput.text);
            return {
                objectSource: getChosenSource(),
                squareSize: (isNaN(sizeValue) || sizeValue <= 0) ? DEFAULTS.squareSize : sizeValue,
                squareColor: { r: pickedColor.r, g: pickedColor.g, b: pickedColor.b },
                symbolize: symbolizeCheckbox.value,
                symbolIndex: symbolDropdown.selection ? symbolDropdown.selection.index : -1,
                moveToLayer: moveToLayerCheckbox.value,
                groupItems: groupCheckbox.value,
                scalePercent: (isNaN(scaleValue) || scaleValue <= 0) ? DEFAULTS.scalePercent : scaleValue,
                registrationIndex: registrationIndex
            };
        }

        /**
         * プレビュー専用レイヤーに、現在の設定でマーカーを描画します。
         * @returns {void}
         */
        function renderPreview() {
            removePreviewLayer();

            var currentSettings = readCurrentSettings();
            var previewPlacer = buildMarkerPlacer(currentSettings, true);
            if (previewPlacer) {
                var previewPoints = resolveAnchorPoints(currentSettings.objectSource);
                activePreviewLayer = doc.layers.add();
                activePreviewLayer.name = PREVIEW_LAYER_NAME;
                for (var p = 0; p < previewPoints.length; p++) {
                    moveItemToLayer(previewPlacer(previewPoints[p][0], previewPoints[p][1]), activePreviewLayer);
                }
            }
            app.redraw();
        }

        // --- ボタン / Buttons（Mac 規約: Cancel → OK）---
        var buttonRow = addButtonRow(dialogWindow);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var dialogResult = null;
        btnOK.onClick = function () {
            var chosenSource = getChosenSource();
            if (chosenSource === OBJECT_SOURCE.autoGenerate) {
                var sizeValue = parseFloat(sizeInput.text);
                if (isNaN(sizeValue) || sizeValue <= 0) {
                    alert(getLabel("alert.invalidSize"));
                    return;
                }
            } else {
                var scaleValue = parseFloat(scaleInput.text);
                if (isNaN(scaleValue) || scaleValue <= 0) {
                    alert(getLabel("alert.invalidScale"));
                    return;
                }
                if (chosenSource === OBJECT_SOURCE.symbol && !symbolDropdown.selection) {
                    alert(getLabel("alert.noSymbolChosen"));
                    return;
                }
            }

            removePreviewLayer(); /* プレビューを片付けてから本適用 / Clear preview before the real placement */
            dialogResult = readCurrentSettings();
            dialogWindow.close();
        };
        btnCancel.onClick = function () {
            removePreviewLayer();
            app.redraw();
            dialogResult = null;
            dialogWindow.close();
        };

        // 表示直後に初回プレビュー / Render the first preview once shown
        dialogWindow.onShow = function () {
            renderPreview();
        };

        // レイアウト確定後、選択ボタンを通常ボタンより -2 に詰める / After layout, trim the choose button by 2px
        dialogWindow.layout.layout(true);
        trimButtonHeight(chooseColorButton, 2);

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialogWindow, SCRIPT_NAME);
        dialogWindow.show();
        return dialogResult;
    }

    // =========================================
    // ダイアログ部品 / Dialog Helpers
    // =========================================
    /**
     * 「ラベル [入力] 単位」の1行を作り、入力欄を返します。
     * @param {Group|Panel} parentPanel - 追加先
     * @param {string} labelText - ラベル文字列
     * @param {number} initialValue - 初期値
     * @param {number} charCount - 入力欄の文字数（最小幅の指定）
     * @param {string} unitText - 単位表記
     * @param {function} onChange - 値が変わったときに呼ぶ関数
     * @param {number} [labelWidth] - ラベルの固定幅（px）
     * @param {number} [minValue] - 下限値（省略時は 0）
     * @returns {EditText} 追加した入力欄（行は .fieldRow、∧∨は .stepperGroup で参照できる）
     */
    function addLabeledField(parentPanel, labelText, initialValue, charCount, unitText, onChange, labelWidth, minValue) {
        var fieldRow = parentPanel.add("group");
        setupRow(fieldRow, "left", 8);
        var fieldLabel = fieldRow.add("statictext", undefined, labelText);
        if (labelWidth) {
            fieldLabel.preferredSize.width = labelWidth; /* 指定時のみ固定幅 / Fixed width only when given */
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var stepperGroup = addStepper(stepperFieldGroup, function () { return fieldInput; }, {
            min: (typeof minValue === "number") ? minValue : 0,
            onStep: onChange
        });
        var fieldInput = stepperFieldGroup.add("edittext", undefined, String(initialValue));
        fieldInput.characters = charCount;
        fieldInput.onChanging = onChange;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(fieldInput, stepperGroup);
        fieldRow.add("statictext", undefined, unitText);
        fieldInput.fieldRow = fieldRow;
        fieldInput.stepperGroup = stepperGroup;
        return fieldInput;
    }

    /**
     * スウォッチ（panel）を現在の色で塗る onDraw ハンドラを生成します。
     * @param {Panel} swatch - 描画対象のパネル
     * @param {Object} color - 塗り色 {r, g, b}
     * @returns {function} onDraw ハンドラ
     */
    function makeSwatchDrawer(swatch, color) {
        return function () {
            var swatchGraphics = swatch.graphics;
            var fillBrush = swatchGraphics.newBrush(
                swatchGraphics.BrushType.SOLID_COLOR,
                [color.r / 255, color.g / 255, color.b / 255, 1]
            );
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatch.size[0], swatch.size[1]);
            swatchGraphics.fillPath(fillBrush);
        };
    }

    /**
     * 自前の RGB カラーダイアログを表示します（OS のカラーパレットはモーダルを壊すため）。
     * @param {Object} startColor - 初期色 {r, g, b}
     * @returns {Object|null} 選択した色 {r, g, b}（キャンセル時は null）
     */
    function chooseRgbColor(startColor) {
        var workingColor = { r: startColor.r, g: startColor.g, b: startColor.b };
        var confirmed = false;

        var pickerWindow = new Window("dialog", getLabel("colorPicker.title"));
        setupWindow(pickerWindow);

        /* 上段：色見本と RGB 欄を横に並べる。ボタン行はその下 / Top row: swatch and RGB fields; the button row goes below */
        var pickerBodyRow = pickerWindow.add("group");
        pickerBodyRow.orientation = "row";
        pickerBodyRow.alignChildren = ["fill", "fill"];
        pickerBodyRow.spacing = WINDOW_SPACING;

        var previewSwatch = pickerBodyRow.add("panel");
        previewSwatch.preferredSize = [64, 64];
        previewSwatch.onDraw = makeSwatchDrawer(previewSwatch, workingColor);

        var fieldsColumn = addColumnGroup(pickerBodyRow, 6);

        /**
         * 3つの数値欄から作業色を読み直し、プレビューを再描画します。
         * @returns {void}
         */
        function refreshFromFields() {
            workingColor.r = clampColorChannel(Number(redInput.text));
            workingColor.g = clampColorChannel(Number(greenInput.text));
            workingColor.b = clampColorChannel(Number(blueInput.text));
            redrawControl(previewSwatch);
        }
        var redInput = addColorChannelField(fieldsColumn, "R", workingColor.r, refreshFromFields);
        var greenInput = addColorChannelField(fieldsColumn, "G", workingColor.g, refreshFromFields);
        var blueInput = addColorChannelField(fieldsColumn, "B", workingColor.b, refreshFromFields);

        var pickerButtonRow = addButtonRow(pickerWindow);
        var btnPickerCancel = pickerButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnPickerOK = pickerButtonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        // show() の戻り値に依存せず、明示的な onClick で確定する（環境差で OK が 1 を返さない対策）
        btnPickerOK.onClick = function () {
            refreshFromFields();
            confirmed = true;
            pickerWindow.close();
        };
        btnPickerCancel.onClick = function () {
            confirmed = false;
            pickerWindow.close();
        };
        alignRightOnlyButtonRow(pickerButtonRow);

        prepareDialogWindow(pickerWindow, SCRIPT_NAME + "_colorPicker");
        pickerWindow.show();
        return confirmed ? { r: workingColor.r, g: workingColor.g, b: workingColor.b } : null;
    }

    /**
     * RGB チャンネル1つ（ラベル＋スライダー＋数値入力を相互同期）を作り、入力欄を返します。
     * @param {Group} parentGroup - 追加先
     * @param {string} channelLabel - チャンネル名（"R" / "G" / "B"）
     * @param {number} initialValue - 初期値（0〜255）
     * @param {function} onChange - 値が変わったときに呼ぶ関数
     * @returns {EditText} 追加した入力欄
     */
    function addColorChannelField(parentGroup, channelLabel, initialValue, onChange) {
        var channelRow = parentGroup.add("group");
        setupRow(channelRow, "left", 6);
        var channelLabelText = channelRow.add("statictext", undefined, channelLabel);
        channelLabelText.preferredSize.width = 14;
        var channelSlider = channelRow.add("slider", undefined, initialValue, 0, 255);
        channelSlider.preferredSize = [140, 18];

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = channelRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        /* 整数・0〜255。増減後はスライダーへ反映 / integer 0-255; sync the slider after stepping */
        var stepperGroup = addStepper(stepperFieldGroup, function () { return channelInput; }, {
            min: 0, max: 255, integer: true,
            onStep: function () { syncFromInput(); }
        });
        var channelInput = stepperFieldGroup.add("edittext", undefined, String(initialValue));
        channelInput.characters = 4;

        /**
         * スライダーの値を数値欄へ反映します。
         * @returns {void}
         */
        function syncFromSlider() {
            channelInput.text = Math.round(channelSlider.value);
            onChange();
        }

        /**
         * 数値欄の値をスライダーへ反映します。
         * @returns {void}
         */
        function syncFromInput() {
            channelSlider.value = clampColorChannel(Number(channelInput.text));
            onChange();
        }
        channelSlider.onChanging = syncFromSlider; /* ドラッグ中（発火する環境）/ During drag where supported */
        channelSlider.onChange = syncFromSlider;   /* ドラッグ後（リリース）で確実に反映 / On release, reliably */
        channelInput.onChanging = syncFromInput;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(channelInput, stepperGroup);
        return channelInput;
    }

    /**
     * 数値を 0〜255 に丸めてクランプします。
     * @param {number} value - 入力値
     * @returns {number} 0〜255 の整数
     */
    function clampColorChannel(value) {
        if (isNaN(value)) return 0;
        value = Math.round(value);
        if (value < 0) return 0;
        if (value > 255) return 255;
        return value;
    }

    /**
     * 生成したプレビューレイヤーだけを参照で削除します（同名の既存レイヤーは触らない）。
     * @returns {void}
     */
    function removePreviewLayer() {
        if (activePreviewLayer) {
            try {
                activePreviewLayer.remove();
            } catch (e) {
                /* 既に無ければ無視 / Ignore if already gone */
            }
            activePreviewLayer = null;
        }
    }

    /**
     * 指定名のレイヤーを取得し、無ければ作成します。
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 取得または作成したレイヤー
     */
    function getOrCreateLayer(layerName) {
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            var layer = doc.layers.add();
            layer.name = layerName;
            return layer;
        }
    }

    /**
     * エッジ表示とライブコーナー注釈をトグルします（実行中は隠す）。
     * @returns {void}
     */
    function toggleCanvasHelpers() {
        app.executeMenuCommand('edge');
        app.executeMenuCommand('Live Corner Annotator');
    }

    /**
     * 生成物を指定レイヤーの最前面へ移動します。
     * @param {PageItem} item - 移動するオブジェクト
     * @param {Layer|GroupItem} layer - 移動先
     * @returns {void}
     */
    function moveItemToLayer(item, layer) {
        item.move(layer, ElementPlacement.PLACEATBEGINNING);
    }

    /**
     * コントロールを再描画します（notify は環境により例外を投げ得るので保護）。
     * @param {Object} control - 再描画するコントロール
     * @returns {void}
     */
    function redrawControl(control) {
        try {
            control.notify("onDraw");
        } catch (e) {}
    }

    /**
     * 配置済みアイテムを1つのグループにまとめます。
     * @param {Array<PageItem>} items - まとめる対象
     * @returns {GroupItem} 作成したグループ
     */
    function groupPlacedItems(items) {
        var markerGroup = doc.groupItems.add();
        for (var i = 0; i < items.length; i++) {
            items[i].move(markerGroup, ElementPlacement.PLACEATEND);
        }
        return markerGroup;
    }

    // =========================================
    // アンカーポイント収集 / Anchor Point Collection
    // =========================================
    /**
     * 選択オブジェクト群からアンカー座標を収集します。
     * @param {Array<PageItem>} items - 走査対象
     * @param {PageItem} [excludeItem] - 除外するオブジェクト
     * @returns {Array<Array<number>>} アンカー座標 [x, y] の配列
     */
    function collectAllAnchorPoints(items, excludeItem) {
        var collectedPoints = [];
        for (var i = 0; i < items.length; i++) {
            collectAnchorPoints(items[i], collectedPoints, excludeItem);
        }
        return collectedPoints;
    }

    /**
     * オブジェクトの種別に応じてアンカー座標を集めます（excludeItem は対象外）。
     * @param {PageItem} item - 対象オブジェクト
     * @param {Array<Array<number>>} collectedPoints - 収集先の配列
     * @param {PageItem} [excludeItem] - 除外するオブジェクト
     * @returns {void}
     */
    function collectAnchorPoints(item, collectedPoints, excludeItem) {
        if (excludeItem && item === excludeItem) {
            return; /* このオブジェクト自身のアンカーは対象外 / Skip this object's own anchors */
        }
        if (item.typename === "PathItem") {
            for (var i = 0; i < item.pathPoints.length; i++) {
                collectedPoints.push(item.pathPoints[i].anchor);
            }
        } else if (item.typename === "CompoundPathItem") {
            for (var j = 0; j < item.pathItems.length; j++) {
                collectAnchorPoints(item.pathItems[j], collectedPoints, excludeItem);
            }
        } else if (item.typename === "GroupItem") {
            for (var k = 0; k < item.pageItems.length; k++) {
                collectAnchorPoints(item.pageItems[k], collectedPoints, excludeItem);
            }
        }
    }

    /**
     * 配置に使うアンカー座標を返します。
     * 「最前面のオブジェクト」モードでは、複製元（最前面オブジェクト）自身のアンカーを除外します。
     * @param {string} objectSource - 追加するオブジェクトの種別（OBJECT_SOURCE の値）
     * @returns {Array<Array<number>>} アンカー座標 [x, y] の配列
     */
    function resolveAnchorPoints(objectSource) {
        var excludeItem = (objectSource === OBJECT_SOURCE.frontObject) ? getFrontmostItem() : null;
        return collectAllAnchorPoints(selectedItems, excludeItem);
    }

    // =========================================
    // 配置処理 / Placement
    // =========================================
    /**
     * ユーザー設定から、アンカー座標に配置する処理（placer 関数）を組み立てます。
     * placer は生成した pageItem を返します（プレビュー時にレイヤーへ移すため）。
     * forPreview=true のときは自動生成の「シンボル化」を無視し、見た目が同じ正方形パスを描きます
     * （プレビューのたびに新規シンボルを登録してシンボルパネルを汚さないため）。
     * @param {MarkerSettings} settings - 現在の設定
     * @param {boolean} forPreview - プレビュー用なら true
     * @returns {function|null} (anchorX, anchorY) を受け取り PageItem を返す関数（用意できなければ null）
     */
    function buildMarkerPlacer(settings, forPreview) {
        var fractionX = (settings.registrationIndex % 3) / 2;         /* 0=左, 0.5=中央, 1=右 / 0=left, .5=center, 1=right */
        var fractionY = Math.floor(settings.registrationIndex / 3) / 2; /* 0=上, 0.5=中央, 1=下 / 0=top, .5=middle, 1=bottom */

        if (settings.objectSource === OBJECT_SOURCE.frontObject) {
            var frontmostItem = getFrontmostItem();
            if (!frontmostItem) return null;
            return function (anchorX, anchorY) {
                return duplicateItemAtPoint(frontmostItem, anchorX, anchorY, settings.scalePercent, fractionX, fractionY);
            };
        }

        if (settings.objectSource === OBJECT_SOURCE.symbol) {
            if (settings.symbolIndex < 0) return null;
            var chosenSymbol = doc.symbols[settings.symbolIndex];
            return function (anchorX, anchorY) {
                return placeSymbolInstance(chosenSymbol, anchorX, anchorY, settings.scalePercent, fractionX, fractionY);
            };
        }

        // OBJECT_SOURCE.autoGenerate（大きさは pt 指定なのでスケールは 100 固定）
        var squareColor = toRgbColor(settings.squareColor);
        if (settings.symbolize && !forPreview) {
            var squareSymbol = createSquareSymbol(settings.squareSize, squareColor);
            return function (anchorX, anchorY) {
                return placeSymbolInstance(squareSymbol, anchorX, anchorY, 100, fractionX, fractionY);
            };
        }
        return function (anchorX, anchorY) {
            return placeSquareRect(settings.squareSize, squareColor, anchorX, anchorY, fractionX, fractionY);
        };
    }

    /**
     * 基準点の割合（0〜1）に合わせて生成物を配置します。
     * @param {PageItem} item - 配置するオブジェクト
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} fractionX - 横方向の基準位置（0=左, 0.5=中央, 1=右）
     * @param {number} fractionY - 縦方向の基準位置（0=上, 0.5=中央, 1=下）
     * @returns {void}
     */
    function positionItemAtAnchor(item, anchorX, anchorY, fractionX, fractionY) {
        item.left = anchorX - fractionX * item.width;
        item.top = anchorY + fractionY * item.height;
    }

    /**
     * {r, g, b} を RGBColor へ変換します。
     * @param {Object} rgb - 色 {r, g, b}
     * @returns {RGBColor} 変換した色
     */
    function toRgbColor(rgb) {
        var color = new RGBColor();
        color.red = rgb.r;
        color.green = rgb.g;
        color.blue = rgb.b;
        return color;
    }

    /**
     * 現在の大きさ・カラーで正方形シンボルを毎回新規登録して返します（既存は再利用しない）。
     * プレビュー（正方形パス）と本適用（同設定のシンボルインスタンス）の見た目を一致させるためです。
     * @param {number} squareSize - 正方形の一辺（pt）
     * @param {RGBColor} squareColor - 塗り色
     * @returns {Symbol} 登録したシンボル
     */
    function createSquareSymbol(squareSize, squareColor) {
        var masterSquare = doc.pathItems.rectangle(squareSize / 2, -squareSize / 2, squareSize, squareSize);
        masterSquare.filled = true;
        masterSquare.fillColor = squareColor;
        masterSquare.stroked = false;

        var symbolDefinition = doc.symbols.add(masterSquare);
        try {
            symbolDefinition.name = SQUARE_SYMBOL_NAME;
        } catch (e) {
            /* 同名シンボルが既にある場合は既定名のまま（重複を許容）/ Keep the default name on collision */
        }
        masterSquare.remove();
        return symbolDefinition;
    }

    /**
     * 基準点に合わせてシンボルインスタンスを配置します。
     * @param {Symbol} symbolDefinition - 配置するシンボル
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} scalePercent - 拡大縮小率（%）
     * @param {number} fractionX - 横方向の基準位置
     * @param {number} fractionY - 縦方向の基準位置
     * @returns {SymbolItem} 配置したインスタンス
     */
    function placeSymbolInstance(symbolDefinition, anchorX, anchorY, scalePercent, fractionX, fractionY) {
        var symbolInstance = doc.symbolItems.add(symbolDefinition);
        if (scalePercent !== 100) {
            symbolInstance.resize(scalePercent, scalePercent);
        }
        positionItemAtAnchor(symbolInstance, anchorX, anchorY, fractionX, fractionY);
        return symbolInstance;
    }

    /**
     * 基準点に合わせて最前面オブジェクトを複製します。
     * @param {PageItem} sourceItem - 複製元
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} scalePercent - 拡大縮小率（%）
     * @param {number} fractionX - 横方向の基準位置
     * @param {number} fractionY - 縦方向の基準位置
     * @returns {PageItem} 複製したオブジェクト
     */
    function duplicateItemAtPoint(sourceItem, anchorX, anchorY, scalePercent, fractionX, fractionY) {
        var duplicatedItem = sourceItem.duplicate();
        if (scalePercent !== 100) {
            duplicatedItem.resize(scalePercent, scalePercent);
        }
        positionItemAtAnchor(duplicatedItem, anchorX, anchorY, fractionX, fractionY);
        return duplicatedItem;
    }

    /**
     * 基準点に合わせて正方形パスを配置します。
     * @param {number} squareSize - 正方形の一辺（pt）
     * @param {RGBColor} squareColor - 塗り色
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} fractionX - 横方向の基準位置
     * @param {number} fractionY - 縦方向の基準位置
     * @returns {PathItem} 配置した正方形
     */
    function placeSquareRect(squareSize, squareColor, anchorX, anchorY, fractionX, fractionY) {
        var squareRect = doc.pathItems.rectangle(0, 0, squareSize, squareSize);
        squareRect.filled = true;
        squareRect.fillColor = squareColor;
        squareRect.stroked = false;
        positionItemAtAnchor(squareRect, anchorX, anchorY, fractionX, fractionY);
        return squareRect;
    }

    /**
     * 「選択範囲内で最前面のオブジェクト」を返します（未選択の最前面は対象外）。
     * 複製元はスクリプト実行中ずっと同じもの（選択は起動時に確定し、モーダル表示中は変更できず、
     * プレビューの複製物は除外レイヤーへ逃がしている）なので、初回の走査結果を使い回します。
     * @returns {PageItem|null} 選択範囲内の最前面オブジェクト（無ければ null）
     */
    function getFrontmostItem() {
        if (cachedFrontmostItem === undefined) {
            cachedFrontmostItem = findFrontmostSelectedItem();
        }
        return cachedFrontmostItem;
    }

    /**
     * ドキュメントを前面から走査して、最初に selected なアイテムを返します。
     * 選択されているものだけを対象にすることで、未選択オブジェクトを複製元に選んでしまう事故を防ぎます。
     * @returns {PageItem|null} 最前面の選択アイテム（無ければ null）
     */
    function findFrontmostSelectedItem() {
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.visible || layer.locked || layer.name === PREVIEW_LAYER_NAME) {
                continue;
            }
            var foundItem = frontmostSelectedInContainer(layer);
            if (foundItem) return foundItem;
        }
        return null;
    }

    /**
     * コンテナ（レイヤー／グループ）を前面から走査し、最初に選択されているアイテムを返します。
     * @param {Layer|GroupItem} container - 走査するコンテナ
     * @returns {PageItem|null} 最初に見つかった選択アイテム（無ければ null）
     */
    function frontmostSelectedInContainer(container) {
        var childItems = container.pageItems;
        for (var i = 0; i < childItems.length; i++) {
            var item = childItems[i];
            if (item.selected) {
                return item; /* この選択アイテムが最前面 / This selected item is frontmost */
            }
            if (item.typename === "GroupItem") {
                var nestedItem = frontmostSelectedInContainer(item);
                if (nestedItem) return nestedItem;
            }
        }
        return null;
    }

    /**
     * ドキュメント内のシンボル名の一覧を返します。
     * @returns {Array<string>} シンボル名の配列
     */
    function getSymbolNames() {
        var names = [];
        for (var i = 0; i < doc.symbols.length; i++) {
            names.push(doc.symbols[i].name);
        }
        return names;
    }

    // =========================================
    // 基準点ウィジェット（9軸）/ Registration widget (9-axis)
    // =========================================
    /**
     * 3×3 の基準点ウィジェットを生成します。クリックで registrationIndex を更新し onChange を呼びます。
     * @param {Group} parentGroup - 追加先
     * @param {function} onChange - 基準点が変わったときに呼ぶ関数
     * @returns {Button} 追加したウィジェット
     */
    function addRegistrationWidget(parentGroup, onChange) {
        var widget = parentGroup.add("button", undefined, "");
        widget.helpTip = getLabel("tooltip.registration");
        widget.preferredSize = [WIDGET_SIZE, WIDGET_SIZE];
        widget.minimumSize = [WIDGET_SIZE, WIDGET_SIZE];
        widget.maximumSize = [WIDGET_SIZE, WIDGET_SIZE];
        widget.onDraw = function () {
            drawRegistrationWidget(this);
        };
        widget.addEventListener("mousedown", function (event) {
            var columnIndex = clampCell(Math.floor(event.clientX / (widget.size[0] / 3)));
            var rowIndex = clampCell(Math.floor(event.clientY / (widget.size[1] / 3)));
            registrationIndex = rowIndex * 3 + columnIndex;
            redrawControl(widget);
            if (typeof onChange === "function") {
                onChange();
            }
        });
        return widget;
    }

    /**
     * セル位置を 0〜2 にクランプします。
     * @param {number} value - 入力値
     * @returns {number} 0〜2 の値
     */
    function clampCell(value) {
        if (value < 0) return 0;
        if (value > 2) return 2;
        return value;
    }

    /**
     * 9軸ウィジェットを描画します（外周の□をケイ線でつなぎ、中央は独立）。
     * @param {Button} widget - 描画対象のウィジェット
     * @returns {void}
     */
    function drawRegistrationWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];
        var foreColor = widget.enabled ? widgetForeColor : widgetDimColor; /* 無効時はディム / Dim when disabled */

        // 背景は塗らない（透過・親ウィンドウ色）/ No background fill (transparent, parent window color)
        var cellSize = 8;   /* 四角のサイズ / Square size */
        var cellGap = 6;    /* 四角どうしの間隔 / Gap between squares */
        var cellStep = cellSize + cellGap;
        var gridSize = cellSize * 3 + cellGap * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        /**
         * セルの左上X座標を返します。
         * @param {number} index - セル番号 0..8
         * @returns {number} X座標
         */
        function cellX(index) { return originX + (index % 3) * cellStep; }

        /**
         * セルの左上Y座標を返します。
         * @param {number} index - セル番号 0..8
         * @returns {number} Y座標
         */
        function cellY(index) { return originY + Math.floor(index / 3) * cellStep; }

        // 中央(4)を除く外周の□どうしをケイ線でつなぐ / Join the outer squares (skipping center) with rules
        var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, foreColor, 1);
        for (var i = 0; i < connections.length; i++) {
            var startCell = connections[i][0];
            var endCell = connections[i][1];
            graphics.newPath();
            if (endCell - startCell === 1) {
                graphics.moveTo(cellX(startCell) + cellSize, cellY(startCell) + cellSize / 2);
                graphics.lineTo(cellX(endCell), cellY(endCell) + cellSize / 2);
            } else {
                graphics.moveTo(cellX(startCell) + cellSize / 2, cellY(startCell) + cellSize);
                graphics.lineTo(cellX(endCell) + cellSize / 2, cellY(endCell));
            }
            graphics.strokePath(linePen);
        }

        for (var index = 0; index < 9; index++) {
            drawRegistrationCell(graphics, cellX(index), cellY(index), cellSize, index === registrationIndex, foreColor);
        }
    }

    /**
     * 基準点セルの□を1つ描画します。
     * 選択中は「塗り＋ケイ線」で描き、非選択の枠線と外周サイズをそろえます。
     * @param {Object} graphics - ScriptUI の graphics オブジェクト
     * @param {number} x - セルの左上X座標
     * @param {number} y - セルの左上Y座標
     * @param {number} size - セルの一辺（px）
     * @param {boolean} selected - 選択中なら true
     * @param {Array<number>} foreColor - 描画色 [r, g, b, a]
     * @returns {void}
     */
    function drawRegistrationCell(graphics, x, y, size, selected, foreColor) {
        if (selected) {
            buildCellPath(graphics, x, y, size);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, foreColor));
        }
        buildCellPath(graphics, x, y, size);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, foreColor, 1));
    }

    /**
     * セルの正方形パスを組み立てます。
     * @param {Object} graphics - ScriptUI の graphics オブジェクト
     * @param {number} x - 左上X座標
     * @param {number} y - 左上Y座標
     * @param {number} size - 一辺（px）
     * @returns {void}
     */
    function buildCellPath(graphics, x, y, size) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
    }

    // =========================================
    // UI の明暗判定 / Light vs. dark UI
    // =========================================
    /**
     * UI 明度からライトテーマかどうかを判定します（取得に失敗したらダーク扱い）。
     * @returns {boolean} ライトテーマなら true
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }
})();
