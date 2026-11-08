#target illustrator
#targetengine "BlendLikeSpacingEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.BlendLikeSpacing
Name=並んだオブジェクトの間隔をだんだん広くする
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.6.0
Release-Date=2026-11-06
Description-BEGIN
並んだ線やオブジェクトの両端を動かさずに、間隔が一定の比率でだんだん広くなるように並べ直します。
比率はスライダーで調整し、結果をプレビューで確認できます。

Rearranges a row of lines or objects so the gaps widen by a constant ratio, keeping both ends in place.
Adjust the ratio with a slider and check the result in the live preview.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

並んだ線やオブジェクトの両端を動かさずに、間隔が一定の比率でだんだん広くなるように並べ直します。
比率はスライダーで調整し、結果をプレビューで確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendLikeSpacing.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n67c10b92e5fd

### Overview

Rearranges a row of lines or objects so the gaps widen by a constant ratio, keeping both ends in place.
Adjust the ratio with a slider and check the result in the live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BlendLikeSpacing.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "BlendLikeSpacing";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-09";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "並んだオブジェクトの間隔をだんだん広くする"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendLikeSpacing.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BlendLikeSpacing.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n67c10b92e5fd"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/BlendLikeSpacing.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_RATIO = 130;  /* 比率の初期値（%） / initial ratio (%) */
    var RATIO_MIN = 100;      /* 比率の下限（%）。100%で等間隔 / lowest ratio (%); 100% means even spacing */
    var RATIO_MAX = 300;      /* 比率の上限（%） / highest ratio (%) */
    var RATIO_STEP = 1;       /* ∧∨・↑↓の増減量（%） / stepper increment (%) */
    var MIN_ITEM_COUNT = 2;   /* 必要な数。2つでも［個数］で増やせば並べられる / items needed; two work once copies are added */
    var DEFAULT_ADD_COUNT = 0; /* ［個数］（増やす数）の初期値 / initial number of copies to add */
    var MAX_ADD_COUNT = 64;   /* ［個数］（増やす数）の上限 / most copies that can be added */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var SLIDER_WIDTH = 220;       /* ［比率］［個数］のスライダーの幅（px） / slider width for Ratio and Count (px) */
    var AXIS_LABEL_PADDING = 2;   /* ［上下］［左右］の項目名の幅に足す余白（px） / extra width for the axis labels (px) */

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

    // =========================================
    // 共通部品 / Shared parts
    // =========================================

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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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
            title: { ja: "間隔をだんだん広く", en: "Gradually Widen Spacing" }
        },
        panel: {
            direction: { ja: "広げる向き", en: "Direction" },
            spacing: { ja: "間隔の広がり", en: "Spacing Growth" },
            objects: { ja: "オブジェクトの数", en: "Number of Objects" },
            color: { ja: "カラーのブレンド", en: "Color Blend" }
        },
        radio: {
            down: { ja: "下へ", en: "Down" },
            up: { ja: "上へ", en: "Up" },
            centerY: { ja: "中央へ", en: "Center" },
            endsY: { ja: "両端へ", en: "Ends" },
            randomY: { ja: "ランダム", en: "Random" },
            right: { ja: "右へ", en: "Right" },
            left: { ja: "左へ", en: "Left" },
            centerX: { ja: "中央へ", en: "Center" },
            endsX: { ja: "両端へ", en: "Ends" },
            randomX: { ja: "ランダム", en: "Random" }
        },
        checkbox: {
            fill: { ja: "塗り", en: "Fill" },
            stroke: { ja: "線", en: "Stroke" },
            reverse: { ja: "反転", en: "Reverse" }
        },
        fieldLabel: {
            y: { ja: "上下", en: "Vertical" },
            x: { ja: "左右", en: "Horizontal" },
            ratio: { ja: "比率", en: "Ratio" },
            count: { ja: "個数", en: "Add" },
            total: { ja: "総数", en: "Total" }
        },
        tooltip: {
            down: { ja: "下へ行くほど間隔を広げる", en: "Gaps widen toward the bottom" },
            up: { ja: "上へ行くほど間隔を広げる", en: "Gaps widen toward the top" },
            centerY: { ja: "上下の中央へ行くほど間隔を広げる", en: "Gaps widen toward the vertical center" },
            endsY: { ja: "上下の両端へ行くほど間隔を広げる（中央が最も狭い）", en: "Gaps widen toward the top and bottom (narrowest in the middle)" },
            randomY: { ja: "上下の間隔をランダムにする（比率は最も広い間隔と最も狭い間隔の比）", en: "Random vertical gaps (the ratio is the widest gap over the narrowest)" },
            right: { ja: "右へ行くほど間隔を広げる", en: "Gaps widen toward the right" },
            left: { ja: "左へ行くほど間隔を広げる", en: "Gaps widen toward the left" },
            centerX: { ja: "左右の中央へ行くほど間隔を広げる", en: "Gaps widen toward the horizontal center" },
            endsX: { ja: "左右の両端へ行くほど間隔を広げる（中央が最も狭い）", en: "Gaps widen toward the left and right (narrowest in the middle)" },
            randomX: { ja: "左右の間隔をランダムにする（比率は最も広い間隔と最も狭い間隔の比）", en: "Random horizontal gaps (the ratio is the widest gap over the narrowest)" },
            ratio: {
                ja: "1つ前の間隔に対する、次の間隔の比率（100%で等間隔）。［ランダム］では最も広い間隔と最も狭い間隔の比",
                en: "Each gap as a percentage of the previous one (100% means even spacing). With Random, the widest gap over the narrowest"
            },
            count: {
                ja: "増やすオブジェクトの個数。上端（左右の並びでは左端）のオブジェクトを複製して足す",
                en: "How many objects to add, as copies of the top object (the leftmost in a horizontal row)"
            },
            total: {
                ja: "選択したオブジェクトと増やす分を合わせた数",
                en: "Selected objects plus the ones to add"
            },
            fill: {
                ja: "両端のオブジェクトの塗りをもとに、あいだのオブジェクトの塗りを段階的に変える（どちらかの端が塗りなし・グラデーション・パターンのときは変えない）",
                en: "Blend the fills of the objects between the two ends (left as is when either end has no fill, a gradient, or a pattern)"
            },
            stroke: {
                ja: "両端のオブジェクトの線をもとに、あいだのオブジェクトの線のカラーを段階的に変える（どちらかの端が線なし・グラデーション・パターンのときは変えない）",
                en: "Blend the stroke colors of the objects between the two ends (left as is when either end has no stroke, a gradient, or a pattern)"
            },
            reverse: {
                ja: "広い側と狭い側を入れ替える（［下へ］と［上へ］、［右へ］と［左へ］、［中央へ］と［両端へ］）。［ランダム］では使わない",
                en: "Swap the wide and narrow sides (Down and Up, Right and Left, Center and Ends). Not used with Random"
            },
            shuffle: { ja: "ランダムな間隔を作り直す", en: "Generate new random gaps" },
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
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" },
            shuffle: { ja: "シャッフル", en: "Shuffle" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            needItems: { ja: "線やオブジェクトを{count}つ以上選択してください。", en: "Select at least {count} lines or objects." },
            samePosition: {
                ja: "選択したオブジェクトがすべて同じ位置にあるため、間隔を変えられません。",
                en: "All the selected objects are in the same position, so their spacing cannot be changed."
            }
        }
    };

    // =========================================
    // 間隔の計算 / Spacing calculation
    // =========================================

    /* 広げる向き。sign は間隔が広がっていく向きの座標の符号（Illustrator の Y は上が正）。
       spread は広がり方で、"oneWay" は一方へ、"center" は両端から中央へ、"ends" は中央から両端へ広げ、"random" はランダム
       （"oneWay" 以外は sign が並べる順を決めるだけ）
       Directions; sign is the coordinate sign the gaps widen toward (Illustrator's Y grows upward).
       spread: "oneWay" widens one way, "center" toward the middle, "ends" toward both ends, "random" at random
       (for all but "oneWay", sign only sets the order) */
    var SPACING_DIRECTIONS = [
        { key: "down",    axis: "y", sign: -1, spread: "oneWay" },
        { key: "up",      axis: "y", sign: 1,  spread: "oneWay" },
        { key: "centerY", axis: "y", sign: -1, spread: "center" },
        { key: "endsY",   axis: "y", sign: -1, spread: "ends" },
        { key: "randomY", axis: "y", sign: -1, spread: "random" },
        { key: "right",   axis: "x", sign: 1,  spread: "oneWay" },
        { key: "left",    axis: "x", sign: -1, spread: "oneWay" },
        { key: "centerX", axis: "x", sign: 1,  spread: "center" },
        { key: "endsX",   axis: "x", sign: 1,  spread: "ends" },
        { key: "randomX", axis: "x", sign: 1,  spread: "random" }
    ];

    /* 並べ替えのキーに足す値（pt）。千倍したキーがどの座標でも10桁にそろい、文字列のまま比べられる
       Added to sort keys (pt) so the key times 1000 has 10 digits at any coordinate and compares as a string */
    var SORT_KEY_OFFSET = 2000000;

    /**
     * 選択したオブジェクトと、動かす前の境界を控える
     * @param {PageItem[]} items - 対象のオブジェクト
     * @param {boolean} usePreviewBounds - true で visibleBounds、false で geometricBounds
     * @returns {Array<{item: PageItem, bounds: number[]}>} 測れたオブジェクトと境界 [左, 上, 右, 下]
     */
    function captureItemStates(items, usePreviewBounds) {
        var itemStates = [];
        for (var i = 0; i < items.length; i++) {
            var itemBounds = getClipAwareBounds(items[i], usePreviewBounds);
            if (itemBounds) itemStates.push({ item: items[i], bounds: itemBounds });
        }
        return itemStates;
    }

    /**
     * 軸に沿った境界の、座標が小さい側と大きい側を返す
     * @param {number[]} itemBounds - 境界 [左, 上, 右, 下]
     * @param {string} axis - "x" または "y"
     * @returns {{low: number, high: number}} 小さい側と大きい側の座標
     */
    function getAxisRange(itemBounds, axis) {
        if (axis === "x") return { low: itemBounds[0], high: itemBounds[2] };
        return { low: itemBounds[3], high: itemBounds[1] };
    }

    /**
     * 中心の散らばり（最大と最小の差）を測る
     * @param {Array<{bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {string} axis - "x" または "y"
     * @returns {number} 中心の散らばり（pt）
     */
    function getCenterSpread(itemStates, axis) {
        var minCenter = Infinity;
        var maxCenter = -Infinity;
        for (var i = 0; i < itemStates.length; i++) {
            var axisRange = getAxisRange(itemStates[i].bounds, axis);
            var itemCenter = (axisRange.low + axisRange.high) / 2;
            if (itemCenter < minCenter) minCenter = itemCenter;
            if (itemCenter > maxCenter) maxCenter = itemCenter;
        }
        return maxCenter - minCenter;
    }

    /**
     * 広げる向きに進む順でオブジェクトを並べる。
     * 座標は向きの符号を掛け、どの向きでも「値が増える方へ進む」形に読み替える。
     * 複製（sourceIndex を持つもの）は位置で並べず、複製元の端のすぐ内側に入れる
     * @param {Array<{bounds: number[], sourceIndex: number}>} itemStates - addItemCopies() の戻り値
     * @param {{axis: string, sign: number}} spacingDirection - 広げる向き
     * @returns {Array<{stateIndex: number, start: number, end: number}>} 進む順に並べた、手前の辺と奥の辺
     */
    function buildOrderedSpans(itemStates, spacingDirection) {
        var itemSpans = [];
        var sortKeys = [];
        var copyStateIndexes = [];
        for (var i = 0; i < itemStates.length; i++) {
            if (itemStates[i].sourceIndex != null) {
                copyStateIndexes.push(i);
                continue;
            }
            var axisRange = getAxisRange(itemStates[i].bounds, spacingDirection.axis);
            var spanStart = (spacingDirection.sign > 0) ? axisRange.low : -axisRange.high;
            var spanEnd = (spacingDirection.sign > 0) ? axisRange.high : -axisRange.low;
            itemSpans.push({ stateIndex: i, start: spanStart, end: spanEnd });
            sortKeys.push(Math.round(((spanStart + spanEnd) / 2 + SORT_KEY_OFFSET) * 1000) + "\u0001" + i);
        }
        /* 比較関数つきの sort() は並びが狂うので、文字列キーで並べる
           A comparator sort misorders in ExtendScript, so sort by string keys */
        sortKeys.sort();
        var orderedSpans = [];
        for (var k = 0; k < sortKeys.length; k++) {
            orderedSpans.push(itemSpans[parseInt(sortKeys[k].split("\u0001")[1], 10)]);
        }
        if (copyStateIndexes.length === 0) return orderedSpans;

        /* 複製元は端にある。端のすぐ内側に、複製元と同じ辺で入れる / the source is at an end; insert the copies just inside it */
        var sourceIndex = itemStates[copyStateIndexes[0]].sourceIndex;
        var sourcePosition = 0;
        for (var s = 0; s < orderedSpans.length; s++) {
            if (orderedSpans[s].stateIndex === sourceIndex) sourcePosition = s;
        }
        var sourceSpan = orderedSpans[sourcePosition];
        var insertAt = (sourcePosition === 0) ? 1 : sourcePosition;
        for (var c = 0; c < copyStateIndexes.length; c++) {
            orderedSpans.splice(insertAt, 0, { stateIndex: copyStateIndexes[c], start: sourceSpan.start, end: sourceSpan.end });
        }
        return orderedSpans;
    }

    /**
     * 間隔1つ分の重み（比率の何乗か）を返す。
     * 一方へ広げるときは手前から順に、中央へ広げるときは両端から中央へ、両端へ広げるときは中央から両端へ1段ずつ比率を掛ける
     * @param {number} gapIndex - 手前から数えた間隔の番号（0始まり）
     * @param {number} gapCount - 間隔の数
     * @param {number} growthRate - 1段ごとに掛ける比率（1.3 など）
     * @param {string} spread - "oneWay" / "center" / "ends" / "random"
     * @param {number[]} [randomExponents] - "random" のときの間隔ごとの指数（0〜1）
     * @returns {number} 重み
     */
    function getGapWeight(gapIndex, gapCount, growthRate, spread, randomExponents) {
        if (spread === "oneWay") return Math.pow(growthRate, gapIndex);
        if (spread === "random") return Math.pow(growthRate, randomExponents[gapIndex]);
        var stepsFromEnd = Math.min(gapIndex, gapCount - 1 - gapIndex);
        if (spread === "center") return Math.pow(growthRate, stepsFromEnd);
        return Math.pow(growthRate, Math.floor((gapCount - 1) / 2) - stepsFromEnd);
    }

    /**
     * ランダムの元になる値を、間隔ごとの指数（最も狭い間隔が0、最も広い間隔が1）にする。
     * 元の値は足りない分だけ足し、プレビューを描き直しても並びが変わらないようにする
     * @param {number[]} randomPool - 0〜1の乱数の控え（足りなければ足す）
     * @param {number} gapCount - 間隔の数
     * @returns {number[]} 間隔ごとの指数
     */
    function getRandomExponents(randomPool, gapCount) {
        while (randomPool.length < gapCount) randomPool.push(Math.random());
        var lowest = Infinity;
        var highest = -Infinity;
        for (var i = 0; i < gapCount; i++) {
            lowest = Math.min(lowest, randomPool[i]);
            highest = Math.max(highest, randomPool[i]);
        }
        var randomExponents = [];
        for (var j = 0; j < gapCount; j++) {
            randomExponents.push((highest > lowest) ? (randomPool[j] - lowest) / (highest - lowest) : 0);
        }
        return randomExponents;
    }

    /**
     * 両端を動かさず、間隔が一定の比率で広がる移動量を求める。
     * オブジェクトどうしのアキを比率で配分する。両端のあいだに収まらないほど重なっているときは、
     * 手前の辺どうしの距離を配分する（大きさのそろった線ならどちらも同じ結果）
     * @param {Array<{bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {{axis: string, sign: number, spread: string}} spacingDirection - 広げる向き
     * @param {number} ratioPercent - 1つ前の間隔に対する比率（%）
     * @param {number[]} randomPool - ランダムの元になる値（getRandomExponents() に渡す）
     * @returns {Array<{stateIndex: number, dx: number, dy: number}>} 動かす量（両端は含まない）
     */
    function computeSpacingMoves(itemStates, spacingDirection, ratioPercent, randomPool) {
        var orderedSpans = buildOrderedSpans(itemStates, spacingDirection);
        var lastIndex = orderedSpans.length - 1;
        var firstSpan = orderedSpans[0];
        var lastSpan = orderedSpans[lastIndex];

        var middleSizeTotal = 0;
        for (var i = 1; i < lastIndex; i++) {
            middleSizeTotal += orderedSpans[i].end - orderedSpans[i].start;
        }
        var freeSpace = lastSpan.start - firstSpan.end - middleSizeTotal;
        var usesItemSizes = freeSpace > SELECTION_ITEMS_TOLERANCE;
        if (!usesItemSizes) freeSpace = lastSpan.start - firstSpan.start;

        /* 比率の重みの合計 / Sum of the gap weights */
        var growthRate = ratioPercent / 100;
        var spread = spacingDirection.spread;
        var randomExponents = (spread === "random") ? getRandomExponents(randomPool, lastIndex) : null;
        var weightTotal = 0;
        for (var gapIndex = 0; gapIndex < lastIndex; gapIndex++) {
            weightTotal += getGapWeight(gapIndex, lastIndex, growthRate, spread, randomExponents);
        }

        var spacingMoves = [];
        var nextEdge = usesItemSizes ? firstSpan.end : firstSpan.start;
        for (var j = 1; j < lastIndex; j++) {
            var movingSpan = orderedSpans[j];
            nextEdge += freeSpace * getGapWeight(j - 1, lastIndex, growthRate, spread, randomExponents) / weightTotal;
            var moveAmount = (nextEdge - movingSpan.start) * spacingDirection.sign;
            spacingMoves.push({
                stateIndex: movingSpan.stateIndex,
                dx: (spacingDirection.axis === "x") ? moveAmount : 0,
                dy: (spacingDirection.axis === "y") ? moveAmount : 0
            });
            if (usesItemSizes) nextEdge += movingSpan.end - movingSpan.start;
        }
        return spacingMoves;
    }

    // =========================================
    // 移動と復元 / Moving and restoring
    // =========================================

    /**
     * 複製元にする端のオブジェクトを選ぶ。上下の並びなら上端、左右の並びなら左端
     * @param {Array<{bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {string} axis - "x" または "y"
     * @returns {number} itemStates の番号
     */
    function findCopySourceIndex(itemStates, axis) {
        var sourceIndex = 0;
        var sourceKey = -Infinity;
        for (var i = 0; i < itemStates.length; i++) {
            var axisRange = getAxisRange(itemStates[i].bounds, axis);
            var centerKey = (axisRange.low + axisRange.high) / 2 * ((axis === "y") ? 1 : -1);
            if (centerKey > sourceKey) {
                sourceKey = centerKey;
                sourceIndex = i;
            }
        }
        return sourceIndex;
    }

    /**
     * 端のオブジェクトを複製して数を増やす。複製は元のすぐ背面に、同じ位置で作る
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {number} copyCount - 足す数
     * @param {string} axis - 並びの軸（"x" または "y"）
     * @returns {Array<{item: PageItem, bounds: number[], sourceIndex: number}>} 元のあとに複製を続けた配列（複製は sourceIndex を持つ）
     */
    function addItemCopies(itemStates, copyCount, axis) {
        var allStates = itemStates.slice(0);
        if (copyCount <= 0) return allStates;
        var sourceIndex = findCopySourceIndex(itemStates, axis);
        var sourceItem = itemStates[sourceIndex].item;
        for (var i = 0; i < copyCount; i++) {
            allStates.push({
                item: sourceItem.duplicate(sourceItem, ElementPlacement.PLACEAFTER),
                bounds: itemStates[sourceIndex].bounds,
                sourceIndex: sourceIndex
            });
        }
        return allStates;
    }

    /**
     * 複製を削除する。app.undo() で消えていれば参照が切れているので何もしない
     * @param {Array<{item: PageItem}>} copyStates - 複製の一覧
     * @returns {void}
     */
    function removeItemCopies(copyStates) {
        for (var i = 0; i < copyStates.length; i++) {
            try {
                if (copyStates[i].item.typename) copyStates[i].item.remove();
            } catch (e) {
                /* 取り消しで消えた複製は読めない / copies removed by the undo can no longer be read */
            }
        }
    }

    /**
     * 求めた量だけオブジェクトを動かす
     * @param {Array<{item: PageItem}>} itemStates - captureItemStates() の戻り値
     * @param {Array<{stateIndex: number, dx: number, dy: number}>} spacingMoves - computeSpacingMoves() の戻り値
     * @returns {boolean} 1つでも動かしたら true
     */
    function applySpacingMoves(itemStates, spacingMoves) {
        var hasMoved = false;
        for (var i = 0; i < spacingMoves.length; i++) {
            var spacingMove = spacingMoves[i];
            if (isNearlySameCoordinate(spacingMove.dx, 0) && isNearlySameCoordinate(spacingMove.dy, 0)) continue;
            itemStates[spacingMove.stateIndex].item.translate(spacingMove.dx, spacingMove.dy);
            hasMoved = true;
        }
        return hasMoved;
    }

    /**
     * 控えた位置からずれているオブジェクトだけを戻す（app.undo() で戻りきらなかった分の保険）
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {boolean} usePreviewBounds - 控えたときと同じ境界の種類
     * @returns {void}
     */
    function restoreItemPositions(itemStates, usePreviewBounds) {
        for (var i = 0; i < itemStates.length; i++) {
            var currentBounds = getClipAwareBounds(itemStates[i].item, usePreviewBounds);
            if (!currentBounds) continue;
            var dx = itemStates[i].bounds[0] - currentBounds[0];
            var dy = itemStates[i].bounds[1] - currentBounds[1];
            /* undo が効いていれば差は0。ここで動かすと履歴が伸びる / Zero after a full undo; moving would add history */
            if (isNearlySameCoordinate(dx, 0) && isNearlySameCoordinate(dy, 0)) continue;
            itemStates[i].item.translate(dx, dy);
        }
    }

    // =========================================
    // カラーのブレンド / Color blending
    // =========================================

    /* カラーの種類ごとの値の上限（下限はどれも0。Lab の a・b は -128〜127）
       Upper bound of each value per color space (lower bound 0; Lab a and b run -128 to 127) */
    var COLOR_SPACE_RANGES = {
        rgb:  { min: [0, 0, 0], max: [255, 255, 255] },
        cmyk: { min: [0, 0, 0, 0], max: [100, 100, 100, 100] },
        gray: { min: [0], max: [100] },
        lab:  { min: [0, -128, -128], max: [100, 127, 127] }
    };

    /**
     * カラーを塗り替える対象（パスと文字）を前面から順に集める。グループは中へたどり、クリッピングマスクとガイドは除く。
     * 複合パスは先頭のパスに入れると全体に効く
     * @param {PageItem} item - 対象のオブジェクト
     * @param {Array<{kind: string, target: Object}>} [colorTargets] - 追加先（省略時は新しい配列）
     * @returns {Array<{kind: string, target: Object}>} kind は "path" か "text"。target は PathItem か TextRange
     */
    function collectColorTargets(item, colorTargets) {
        if (!colorTargets) colorTargets = [];
        if (isClipMaskItem(item)) return colorTargets;
        var typeName = item.typename;
        if (typeName === "PathItem") {
            if (!item.guides) colorTargets.push({ kind: "path", target: item });
        } else if (typeName === "CompoundPathItem") {
            if (item.pathItems.length > 0) colorTargets.push({ kind: "path", target: item.pathItems[0] });
        } else if (typeName === "TextFrame") {
            colorTargets.push({ kind: "text", target: item.textRange });
        } else if (typeName === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) collectColorTargets(item.pageItems[i], colorTargets);
        }
        return colorTargets;
    }

    /**
     * 塗りか線のカラーを読む。塗り・線なしなら null
     * @param {{kind: string, target: Object}} colorTarget - collectColorTargets() の要素
     * @param {string} attribute - "fill" または "stroke"
     * @returns {Color|null} カラー
     */
    function readTargetColor(colorTarget, attribute) {
        try {
            if (colorTarget.kind === "text") {
                var textAttributes = colorTarget.target.characterAttributes;
                return (attribute === "fill") ? textAttributes.fillColor : textAttributes.strokeColor;
            }
            var pathItem = colorTarget.target;
            if (attribute === "fill") return pathItem.filled ? pathItem.fillColor : null;
            return pathItem.stroked ? pathItem.strokeColor : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 塗りか線にカラーを入れる。パスは塗り・線なしでも入れたほうを有効にする
     * @param {{kind: string, target: Object}} colorTarget - collectColorTargets() の要素
     * @param {string} attribute - "fill" または "stroke"
     * @param {Color} color - 入れるカラー
     * @returns {boolean} 入れられたら true（ロックされた子などは false）
     */
    function writeTargetColor(colorTarget, attribute, color) {
        try {
            var colorHolder = (colorTarget.kind === "text") ? colorTarget.target.characterAttributes : colorTarget.target;
            if (attribute === "fill") {
                if (colorTarget.kind === "path") colorHolder.filled = true;
                colorHolder.fillColor = color;
            } else {
                if (colorTarget.kind === "path") colorHolder.stroked = true;
                colorHolder.strokeColor = color;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * カラーを、混ぜられる形（カラーの種類と値）にする。なし・グラデーション・パターンは null
     * @param {Color|null} color - 元のカラー
     * @returns {{space: string, values: number[]}|{space: string, spot: Spot, tint: number}|null} 混ぜられる形
     */
    function toColorModel(color) {
        if (!color) return null;
        switch (color.typename) {
            case "RGBColor": return { space: "rgb", values: [color.red, color.green, color.blue] };
            case "CMYKColor": return { space: "cmyk", values: [color.cyan, color.magenta, color.yellow, color.black] };
            case "GrayColor": return { space: "gray", values: [color.gray] };
            case "LabColor": return { space: "lab", values: [color.l, color.a, color.b] };
            case "SpotColor": return { space: "spot", spot: color.spot, tint: color.tint };
        }
        return null;
    }

    /**
     * 濃度100%の値に濃度を掛ける（紙の白へ近づける）
     * @param {number[]} colorValues - 濃度100%の値
     * @param {string} space - "rgb" / "cmyk" / "gray" / "lab"
     * @param {number} tint - 濃度（%）
     * @returns {number[]} 濃度を掛けた値
     */
    function applyTintToValues(colorValues, space, tint) {
        var tintRate = tint / 100;
        var tinted = [];
        for (var i = 0; i < colorValues.length; i++) {
            if (space === "rgb") tinted.push(255 - (255 - colorValues[i]) * tintRate);
            else if (space === "lab") tinted.push((i === 0) ? 100 - (100 - colorValues[i]) * tintRate : colorValues[i] * tintRate);
            else tinted.push(colorValues[i] * tintRate);
        }
        return tinted;
    }

    /**
     * 混ぜられる形のカラーを、指定の種類の値にする。スポットカラーは元のカラーに濃度を掛ける
     * @param {Object} colorModel - toColorModel() の戻り値
     * @param {string} space - "rgb" / "cmyk" / "gray" / "lab"
     * @returns {number[]|null} 値（換算できなければ null）
     */
    function convertColorModel(colorModel, space) {
        if (colorModel.space === "spot") {
            var baseModel = toColorModel(colorModel.spot.color);
            if (!baseModel || baseModel.space === "spot") return null;
            var baseValues = convertColorModel(baseModel, space);
            return baseValues ? applyTintToValues(baseValues, space, colorModel.tint) : null;
        }
        if (colorModel.space === space) return colorModel.values.slice(0);
        if (colorModel.space === "gray") {
            var grayLevel = colorModel.values[0];
            if (space === "cmyk") return [0, 0, 0, grayLevel];
            var rgbLevel = 255 * (1 - grayLevel / 100);
            if (space === "rgb") return [rgbLevel, rgbLevel, rgbLevel];
            colorModel = { space: "rgb", values: [rgbLevel, rgbLevel, rgbLevel] };
        }
        var imageSpaces = { rgb: ImageColorSpace.RGB, cmyk: ImageColorSpace.CMYK, lab: ImageColorSpace.LAB };
        if (!imageSpaces[colorModel.space] || !imageSpaces[space]) return null;
        try {
            return app.convertSampleColor(imageSpaces[colorModel.space], colorModel.values, imageSpaces[space], ColorConvertPurpose.defaultpurpose);
        } catch (e) {
            return null;
        }
    }

    /**
     * 値からカラーを作る（範囲の外は端で止める）
     * @param {string} space - "rgb" / "cmyk" / "gray" / "lab"
     * @param {number[]} colorValues - 値
     * @returns {Color} カラー
     */
    function createColor(space, colorValues) {
        var spaceRange = COLOR_SPACE_RANGES[space];
        var v = [];
        for (var i = 0; i < spaceRange.max.length; i++) {
            v.push(Math.min(spaceRange.max[i], Math.max(spaceRange.min[i], colorValues[i])));
        }
        var color;
        if (space === "rgb") {
            color = new RGBColor();
            color.red = v[0]; color.green = v[1]; color.blue = v[2];
        } else if (space === "cmyk") {
            color = new CMYKColor();
            color.cyan = v[0]; color.magenta = v[1]; color.yellow = v[2]; color.black = v[3];
        } else if (space === "gray") {
            color = new GrayColor();
            color.gray = v[0];
        } else {
            color = new LabColor();
            color.l = v[0]; color.a = v[1]; color.b = v[2];
        }
        return color;
    }

    /**
     * 2つのカラーを、指定の割合で混ぜる。
     * 同じスポットカラーどうしは濃度を、同じ種類どうしはその種類のまま、それ以外はドキュメントのカラーモードで混ぜる
     * @param {Object|null} startModel - 手前の端のカラー（toColorModel() の戻り値）
     * @param {Object|null} endModel - 奥の端のカラー
     * @param {number} mixRate - 奥の端の割合（0〜1）
     * @param {string} documentSpace - "rgb" または "cmyk"
     * @returns {Color|null} 混ぜたカラー（どちらかが混ぜられなければ null）
     */
    function blendColorModels(startModel, endModel, mixRate, documentSpace) {
        if (!startModel || !endModel) return null;
        if (startModel.space === "spot" && endModel.space === "spot" && startModel.spot.name === endModel.spot.name) {
            var spotColor = new SpotColor();
            spotColor.spot = startModel.spot;
            spotColor.tint = startModel.tint + (endModel.tint - startModel.tint) * mixRate;
            return spotColor;
        }
        var space = (startModel.space === endModel.space && startModel.space !== "spot") ? startModel.space : documentSpace;
        var startValues = convertColorModel(startModel, space);
        var endValues = convertColorModel(endModel, space);
        if (!startValues || !endValues) return null;
        var mixedValues = [];
        for (var i = 0; i < startValues.length; i++) {
            mixedValues.push(startValues[i] + (endValues[i] - startValues[i]) * mixRate);
        }
        return createColor(space, mixedValues);
    }

    /**
     * 両端のオブジェクトの塗り・線から、あいだのオブジェクトのカラーを段階的に決める。
     * グループなどの中身は前面から順に対応させ、数が足りない側は最後のものを使う
     * @param {Array<{item: PageItem}>} itemStates - addItemCopies() の戻り値
     * @param {Array<{stateIndex: number}>} orderedSpans - buildOrderedSpans() の戻り値
     * @param {{fill: boolean, stroke: boolean}} colorOptions - ブレンドする属性
     * @returns {boolean} 1つでも塗り替えたら true
     */
    function applyColorBlend(itemStates, orderedSpans, colorOptions) {
        var lastIndex = orderedSpans.length - 1;
        var attributes = [];
        if (colorOptions.fill) attributes.push("fill");
        if (colorOptions.stroke) attributes.push("stroke");
        if (lastIndex < 2 || attributes.length === 0) return false;

        var startTargets = collectColorTargets(itemStates[orderedSpans[0].stateIndex].item);
        var endTargets = collectColorTargets(itemStates[orderedSpans[lastIndex].stateIndex].item);
        if (startTargets.length === 0 || endTargets.length === 0) return false;
        var endModels = {};
        for (var a = 0; a < attributes.length; a++) {
            endModels[attributes[a]] = { start: [], end: [] };
            for (var s = 0; s < startTargets.length; s++) endModels[attributes[a]].start.push(toColorModel(readTargetColor(startTargets[s], attributes[a])));
            for (var e = 0; e < endTargets.length; e++) endModels[attributes[a]].end.push(toColorModel(readTargetColor(endTargets[e], attributes[a])));
        }
        var documentSpace = (app.activeDocument.documentColorSpace === DocumentColorSpace.CMYK) ? "cmyk" : "rgb";

        var hasChanged = false;
        for (var j = 1; j < lastIndex; j++) {
            var middleTargets = collectColorTargets(itemStates[orderedSpans[j].stateIndex].item);
            for (var k = 0; k < middleTargets.length; k++) {
                for (var b = 0; b < attributes.length; b++) {
                    var attributeModels = endModels[attributes[b]];
                    var mixedColor = blendColorModels(
                        attributeModels.start[Math.min(k, attributeModels.start.length - 1)],
                        attributeModels.end[Math.min(k, attributeModels.end.length - 1)],
                        j / lastIndex,
                        documentSpace
                    );
                    if (mixedColor && writeTargetColor(middleTargets[k], attributes[b], mixedColor)) hasChanged = true;
                }
            }
        }
        return hasChanged;
    }

    /**
     * カラーを比べるための文字列にする（なしは "none"）
     * @param {Color|null} color - カラー
     * @returns {string} 比較用の文字列
     */
    function describeColor(color) {
        var colorModel = toColorModel(color);
        if (!color) return "none";
        if (!colorModel) return color.typename;
        if (colorModel.space === "spot") return "spot:" + colorModel.spot.name + ":" + formatStepperNumber(colorModel.tint);
        var parts = [];
        for (var i = 0; i < colorModel.values.length; i++) parts.push(formatStepperNumber(colorModel.values[i]));
        return colorModel.space + ":" + parts.join(",");
    }

    /**
     * 元のオブジェクトの塗り・線を控える
     * @param {Array<{item: PageItem}>} itemStates - captureItemStates() の戻り値
     * @returns {Array<{colorTarget: Object, fill: Color|null, stroke: Color|null}>} 控え
     */
    function captureColorStates(itemStates) {
        var colorStates = [];
        for (var i = 0; i < itemStates.length; i++) {
            var colorTargets = collectColorTargets(itemStates[i].item);
            for (var k = 0; k < colorTargets.length; k++) {
                colorStates.push({
                    colorTarget: colorTargets[k],
                    fill: readTargetColor(colorTargets[k], "fill"),
                    stroke: readTargetColor(colorTargets[k], "stroke")
                });
            }
        }
        return colorStates;
    }

    /**
     * 控えた塗り・線から変わっているものだけを戻す（app.undo() で戻りきらなかった分の保険）
     * @param {Array<{colorTarget: Object, fill: Color|null, stroke: Color|null}>} colorStates - captureColorStates() の戻り値
     * @returns {void}
     */
    function restoreItemColors(colorStates) {
        var attributes = ["fill", "stroke"];
        for (var i = 0; i < colorStates.length; i++) {
            for (var a = 0; a < attributes.length; a++) {
                var savedColor = colorStates[i][attributes[a]];
                var colorTarget = colorStates[i].colorTarget;
                if (describeColor(readTargetColor(colorTarget, attributes[a])) === describeColor(savedColor)) continue;
                try {
                    if (savedColor) writeTargetColor(colorTarget, attributes[a], savedColor);
                    else if (colorTarget.kind === "path") colorTarget.target[(attributes[a] === "fill") ? "filled" : "stroked"] = false;
                    else writeTargetColor(colorTarget, attributes[a], new NoColor());
                } catch (e) {
                    /* 消えた対象は戻せない / a removed target cannot be restored */
                }
            }
        }
    }

    /**
     * プレビューの描き直しと取り消し、本番の適用をまとめる
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {boolean} usePreviewBounds - 控えたときと同じ境界の種類
     * @returns {{update: function(Object): void, clear: function(): void, commit: function(Object): void}} 描き直し・取り消し・本番
     */
    function createSpacingPreview(itemStates, usePreviewBounds) {
        /* 何も変えなかったときは false のまま。app.undo() がユーザーの操作を取り消さないようにする
           Stays false when nothing changed, so app.undo() never reverts the user's own step */
        var hasUncommittedPreview = false;
        /** @type {Array<{item: PageItem}>} プレビューで足した複製 */
        var previewCopyStates = [];
        /** @type {Array|null} 元の塗り・線の控え（初めてカラーをブレンドするときに取る） */
        var originalColorStates = null;
        /* プレビューでカラーを塗り替えたら true / true while the preview has blended colors */
        var hasColorPreview = false;

        /**
         * プレビューを取り除き、元の位置とカラーへ戻す。app.undo() で履歴ごと取り除き、残った複製は消し、
         * 戻りきらなかった分は座標と控えたカラーで戻す
         * @returns {void}
         */
        function clearPreview() {
            if (hasUncommittedPreview) app.undo();
            hasUncommittedPreview = false;
            removeItemCopies(previewCopyStates);
            previewCopyStates = [];
            restoreItemPositions(itemStates, usePreviewBounds);
            if (hasColorPreview) restoreItemColors(originalColorStates);
            hasColorPreview = false;
        }

        /**
         * 設定どおりに複製を足して並べ、選んでいればカラーをブレンドする
         * @param {{direction: Object, ratio: number, count: number, randomPool: number[], colorBlend: Object}} spacingSettings - 今の設定
         * @returns {boolean} 何か変えたら true
         */
        function applySettings(spacingSettings) {
            var colorBlend = spacingSettings.colorBlend;
            var blendsColor = colorBlend.fill || colorBlend.stroke;
            if (blendsColor && originalColorStates === null) originalColorStates = captureColorStates(itemStates);

            var allStates = addItemCopies(itemStates, spacingSettings.count - itemStates.length, spacingSettings.direction.axis);
            previewCopyStates = allStates.slice(itemStates.length);
            var spacingMoves = computeSpacingMoves(allStates, spacingSettings.direction, spacingSettings.ratio, spacingSettings.randomPool);
            var hasMoved = applySpacingMoves(allStates, spacingMoves);
            if (blendsColor) hasColorPreview = applyColorBlend(allStates, buildOrderedSpans(allStates, spacingSettings.direction), colorBlend);
            return hasMoved || previewCopyStates.length > 0 || hasColorPreview;
        }

        /**
         * 指定の設定でプレビューを描き直す
         * @param {{direction: Object, ratio: number, count: number, randomPool: number[], colorBlend: Object}} spacingSettings - 今の設定
         * @returns {void}
         */
        function updatePreview(spacingSettings) {
            clearPreview();
            hasUncommittedPreview = applySettings(spacingSettings);
            app.redraw();
        }

        /**
         * プレビューを消してから本番を1回だけ流す（取り消しが1回で済む）。複製は残す
         * @param {{direction: Object, ratio: number, count: number, randomPool: number[], colorBlend: Object}} spacingSettings - 確定した設定
         * @returns {void}
         */
        function commitSettings(spacingSettings) {
            clearPreview();
            applySettings(spacingSettings);
            previewCopyStates = [];
            hasColorPreview = false;
        }

        return { update: updatePreview, clear: clearPreview, commit: commitSettings };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ［広げる向き］のラジオボタン。reversed は［反転］をオンにしたときの向き（null は反転しない）
       Direction radios; reversed is the direction used while Reverse is on (null means it does not reverse) */
    var DIRECTION_CHOICES = [
        { axis: "y", normal: "down",    reversed: "up" },
        { axis: "y", normal: "centerY", reversed: "endsY" },
        { axis: "y", normal: "randomY", reversed: null },
        { axis: "x", normal: "right",   reversed: "left" },
        { axis: "x", normal: "centerX", reversed: "endsX" },
        { axis: "x", normal: "randomX", reversed: null }
    ];

    /**
     * ［広げる向き］のパネルを足す。［上下］［左右］の2行にそれぞれ一方へ・中央へ・ランダムを並べ、
     * 下に［反転］を置く。［反転］をオンにすると、ラジオボタンの文言も反転後の向きに入れ替える。
     * 上下の並びなら［下へ］、左右の並びなら［右へ］を選んでおく
     * @param {Window} spacingDialog - 追加先のダイアログ
     * @param {{x: number, y: number}} axisSpreads - 軸ごとの中心の散らばり
     * @param {function(Object): void} onDirectionChange - 向きを変えたときの処理（反転を反映した向きを渡す）
     * @returns {{key: string, axis: string, sign: number, spread: string}} 最初に選んだ向き
     */
    function addDirectionPanel(spacingDialog, axisSpreads, onDirectionChange) {
        var directionPanel = spacingDialog.add("panel", undefined, getLabel(LABELS.panel.direction));
        setupPanel(directionPanel);

        var directionsByKey = {};
        for (var d = 0; d < SPACING_DIRECTIONS.length; d++) directionsByKey[SPACING_DIRECTIONS[d].key] = SPACING_DIRECTIONS[d];

        /* 表示前のラジオ・チェックボックスの value は読み戻せないので、選んでいるものは変数で持つ
           Values cannot be read back before show(), so the choice and Reverse live in variables */
        var selectedChoice = DIRECTION_CHOICES[(axisSpreads.y >= axisSpreads.x) ? 0 : 3];
        var isReversed = false;
        /* ラジオは同じ親の中でしか排他にならないので、行をまたいだ解除は自分でする
           Radios are exclusive only within one parent, so clear the other row by hand */
        var directionRadios = [];
        var rowLabels = [];
        var axisRows = {};

        /**
         * ラジオボタンの選択と［反転］から、使う向きを返す
         * @param {{normal: string, reversed: string|null}} directionChoice - DIRECTION_CHOICES の要素
         * @returns {{key: string, axis: string, sign: number, spread: string}} 広げる向き
         */
        function resolveDirection(directionChoice) {
            var directionKey = (isReversed && directionChoice.reversed) ? directionChoice.reversed : directionChoice.normal;
            return directionsByKey[directionKey];
        }

        /**
         * 項目名つきの行を足す。項目名は2行目以降と幅をそろえる
         * @param {string} labelString - 項目名（空なら字下げだけ）
         * @param {boolean} isEnabled - 項目名を有効にするなら true
         * @returns {Group} 行
         */
        function addLabeledRow(labelString, isEnabled) {
            var labeledRow = directionPanel.add("group");
            setupRow(labeledRow, "left");
            var rowLabel = labeledRow.add("statictext", undefined, labelString);
            rowLabel.justify = "right";
            rowLabel.enabled = isEnabled;
            rowLabels.push(rowLabel);
            return labeledRow;
        }

        /**
         * 軸の行を返す。無ければ作る
         * @param {string} axis - "x" または "y"
         * @returns {Group} 行
         */
        function getAxisRow(axis) {
            if (!axisRows[axis]) axisRows[axis] = addLabeledRow(labelText(LABELS.fieldLabel[axis]), !isNearlySameCoordinate(axisSpreads[axis], 0));
            return axisRows[axis];
        }

        /**
         * ラジオボタンの文言とツールチップを、今の［反転］に合わせる
         * @param {RadioButton} directionRadio - 対象のラジオボタン
         * @returns {void}
         */
        function updateRadioLabel(directionRadio) {
            var directionKey = resolveDirection(directionRadio.directionChoice).key;
            directionRadio.text = getLabel(LABELS.radio[directionKey]);
            directionRadio.helpTip = getLabel(LABELS.tooltip[directionKey]);
        }

        /**
         * 向きのラジオボタンを1つ足す
         * @param {{axis: string, normal: string, reversed: string|null}} directionChoice - DIRECTION_CHOICES の要素
         * @returns {void}
         */
        function addDirectionRadio(directionChoice) {
            var directionRadio = getAxisRow(directionChoice.axis).add("radiobutton", undefined, "");
            directionRadio.directionChoice = directionChoice;
            updateRadioLabel(directionRadio);
            /* 散らばりのない軸では動かせない / An axis with no spread has nothing to move */
            directionRadio.enabled = !isNearlySameCoordinate(axisSpreads[directionChoice.axis], 0);
            directionRadio.value = (directionChoice === selectedChoice);
            directionRadio.onClick = function () {
                for (var r = 0; r < directionRadios.length; r++) {
                    if (directionRadios[r] !== directionRadio) directionRadios[r].value = false;
                }
                selectedChoice = directionChoice;
                reverseCheckbox.enabled = (selectedChoice.reversed !== null);
                onDirectionChange(resolveDirection(selectedChoice));
            };
            directionRadios.push(directionRadio);
        }

        for (var i = 0; i < DIRECTION_CHOICES.length; i++) {
            addDirectionRadio(DIRECTION_CHOICES[i]);
        }

        /* ［反転］はラジオの頭にそろえる。［ランダム］では使わない / align Reverse with the radios; unused with Random */
        var reverseCheckbox = addLabeledRow("", true).add("checkbox", undefined, getLabel(LABELS.checkbox.reverse));
        reverseCheckbox.helpTip = getLabel(LABELS.tooltip.reverse);
        reverseCheckbox.value = isReversed;
        reverseCheckbox.enabled = (selectedChoice.reversed !== null);
        reverseCheckbox.onClick = function () {
            isReversed = reverseCheckbox.value;
            for (var r = 0; r < directionRadios.length; r++) updateRadioLabel(directionRadios[r]);
            onDirectionChange(resolveDirection(selectedChoice));
        };

        /* 項目名を最も広い幅にそろえ、ラジオとチェックボックスの頭をそろえる / match the label widths so the controls line up */
        var labelWidth = 0;
        for (var k = 0; k < rowLabels.length; k++) {
            labelWidth = Math.max(labelWidth, rowLabels[k].graphics.measureString(rowLabels[k].text)[0]);
        }
        for (var m = 0; m < rowLabels.length; m++) {
            rowLabels[m].preferredSize.width = Math.ceil(labelWidth) + AXIS_LABEL_PADDING;
        }
        return resolveDirection(selectedChoice);
    }

    /**
     * ［間隔の広がり］のパネルを足す。入力欄とスライダーを連動させる
     * @param {Window} spacingDialog - 追加先のダイアログ
     * @param {function(): void} onRatioChange - 比率を変えたときの処理
     * @returns {function(): number} 比率（%）を読む関数（確定前の入力もそのまま読む）
     */
    function addRatioPanel(spacingDialog, onRatioChange) {
        var spacingPanel = spacingDialog.add("panel", undefined, getLabel(LABELS.panel.spacing));
        setupPanel(spacingPanel);

        var ratioFieldOptions = {
            label: labelText(LABELS.fieldLabel.ratio),
            text: DEFAULT_RATIO + "%",
            characters: 6,
            step: RATIO_STEP,
            min: RATIO_MIN,
            max: RATIO_MAX,
            unit: "%",
            onStep: applyFieldRatio
        };
        var ratioField = addSteppedField(spacingPanel, ratioFieldOptions);
        ratioField.helpTip = getLabel(LABELS.tooltip.ratio);

        var ratioSlider = spacingPanel.add("slider", undefined, DEFAULT_RATIO, RATIO_MIN, RATIO_MAX);
        ratioSlider.preferredSize.width = SLIDER_WIDTH;
        ratioSlider.helpTip = getLabel(LABELS.tooltip.ratio);

        /**
         * 入力欄の比率を読む
         * @returns {number} 比率（%）
         */
        function readRatio() {
            var ratioValue = parseFloat(ratioField.text);
            if (isNaN(ratioValue)) ratioValue = ratioSlider.value;
            return clampSteppedValue(ratioValue, ratioFieldOptions);
        }

        /**
         * 入力欄の値をスライダーへ写して知らせる
         * @returns {void}
         */
        function applyFieldRatio() {
            ratioSlider.value = readRatio();
            onRatioChange();
        }

        /* 値が同じなら何もしない。スライダーへの代入で呼ばれたときもここで止まる
           Nothing to do at the same value; this also stops calls caused by assigning to the slider */
        ratioSlider.onChanging = function () {
            var sliderRatio = Math.round(ratioSlider.value);
            if (sliderRatio === Math.round(readRatio())) return;
            writeSteppedValue(ratioField, sliderRatio, ratioFieldOptions);
            onRatioChange();
        };
        ratioSlider.onChange = ratioSlider.onChanging;

        var normalizeRatioField = ratioField.onChange;
        ratioField.onChange = function () {
            normalizeRatioField();
            applyFieldRatio();
        };
        return readRatio;
    }

    /**
     * ［オブジェクトの数］のパネルを足す。［個数］は増やす数で、入力欄とスライダーを連動させる。
     * 右に、選択と合わせた総数を表示する（編集不可）
     * @param {Window} spacingDialog - 追加先のダイアログ
     * @param {number} selectedCount - 選択したオブジェクトの数
     * @param {function(): void} onCountChange - 個数を変えたときの処理
     * @returns {function(): number} 総数を読む関数（確定前の入力もそのまま読む）
     */
    function addCountPanel(spacingDialog, selectedCount, onCountChange) {
        var objectsPanel = spacingDialog.add("panel", undefined, getLabel(LABELS.panel.objects));
        setupPanel(objectsPanel);

        var maxAddCount = MAX_ADD_COUNT;
        var initialAddCount = Math.min(DEFAULT_ADD_COUNT, maxAddCount);
        var countFieldOptions = {
            label: labelText(LABELS.fieldLabel.count),
            text: String(initialAddCount),
            characters: 6,
            step: 1,
            min: 0,
            max: maxAddCount,
            integer: true,
            onStep: applyFieldCount
        };
        var countField = addSteppedField(objectsPanel, countFieldOptions);
        countField.helpTip = getLabel(LABELS.tooltip.count);

        /* 入力欄の行の右に総数を出す。幅は最も大きい総数に合わせ、数が変わっても行が動かないようにする
           Show the total right of the field, sized for the largest total so the row never shifts */
        var totalText = countField.parent.parent.add("statictext", undefined, labelValueText(LABELS.fieldLabel.total, selectedCount + maxAddCount));
        totalText.preferredSize.width = Math.ceil(totalText.graphics.measureString(totalText.text)[0]) + AXIS_LABEL_PADDING;
        totalText.helpTip = getLabel(LABELS.tooltip.total);

        var countSlider = objectsPanel.add("slider", undefined, initialAddCount, 0, maxAddCount);
        countSlider.preferredSize.width = SLIDER_WIDTH;
        countSlider.helpTip = getLabel(LABELS.tooltip.count);

        /**
         * 入力欄の増やす数を読む
         * @returns {number} 増やす数
         */
        function readAddCount() {
            var addCount = parseInt(countField.text, 10);
            if (isNaN(addCount)) addCount = initialAddCount;
            return Math.round(clampSteppedValue(addCount, countFieldOptions));
        }

        /**
         * 総数の表示を今の値にする
         * @returns {void}
         */
        function updateTotalText() {
            totalText.text = labelValueText(LABELS.fieldLabel.total, selectedCount + readAddCount());
        }

        /**
         * 入力欄の値をスライダーと総数へ写して知らせる
         * @returns {void}
         */
        function applyFieldCount() {
            countSlider.value = readAddCount();
            updateTotalText();
            onCountChange();
        }

        /* 値が同じなら何もしない。スライダーへの代入で呼ばれたときもここで止まる
           Nothing to do at the same value; this also stops calls caused by assigning to the slider */
        countSlider.onChanging = function () {
            var sliderCount = Math.round(countSlider.value);
            if (sliderCount === readAddCount()) return;
            writeSteppedValue(countField, sliderCount, countFieldOptions);
            updateTotalText();
            onCountChange();
        };
        countSlider.onChange = countSlider.onChanging;

        var normalizeCountField = countField.onChange;
        countField.onChange = function () {
            normalizeCountField();
            applyFieldCount();
        };
        updateTotalText();

        return function readTotalCount() {
            return selectedCount + readAddCount();
        };
    }

    /**
     * ［カラーのブレンド］のパネルを足す。塗りと線を別々に選べる
     * @param {Window} spacingDialog - 追加先のダイアログ
     * @param {function(): void} onColorChange - 選択を変えたときの処理
     * @returns {function(): {fill: boolean, stroke: boolean}} ブレンドする属性を読む関数
     */
    function addColorPanel(spacingDialog, onColorChange) {
        var colorPanel = spacingDialog.add("panel", undefined, getLabel(LABELS.panel.color));
        setupPanel(colorPanel);
        var checkboxRow = colorPanel.add("group");
        setupRow(checkboxRow, "left");

        /* 表示前の checkbox の value は読み戻せないので、変数で持つ
           A checkbox's value cannot be read back before show(), so keep it in a variable */
        var colorOptions = { fill: false, stroke: false };

        /**
         * 属性のチェックボックスを1つ足す
         * @param {string} attribute - "fill" または "stroke"
         * @returns {void}
         */
        function addAttributeCheckbox(attribute) {
            var attributeCheckbox = checkboxRow.add("checkbox", undefined, getLabel(LABELS.checkbox[attribute]));
            attributeCheckbox.helpTip = getLabel(LABELS.tooltip[attribute]);
            attributeCheckbox.value = colorOptions[attribute];
            attributeCheckbox.onClick = function () {
                colorOptions[attribute] = attributeCheckbox.value;
                onColorChange();
            };
        }
        addAttributeCheckbox("fill");
        addAttributeCheckbox("stroke");

        return function readColorOptions() {
            return { fill: colorOptions.fill, stroke: colorOptions.stroke };
        };
    }

    /**
     * ダイアログを開き、OK なら間隔を確定する。キャンセルなら元の位置に戻す
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {boolean} usePreviewBounds - 境界の種類
     * @param {{x: number, y: number}} axisSpreads - 軸ごとの中心の散らばり
     * @returns {void}
     */
    function showSpacingDialog(itemStates, usePreviewBounds, axisSpreads) {
        var spacingPreview = createSpacingPreview(itemStates, usePreviewBounds);
        /** @type {Object|null} OK を押したときの設定（キャンセルなら null） */
        var committedSettings = null;
        /* ランダムの元になる値。シャッフルするまで同じ値を使い、比率や数を変えても並びを保つ
           Random source values; kept until Shuffle so changing the ratio or count keeps the pattern */
        var randomPool = [];

        var spacingDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(spacingDialog);

        /* 表示前のラジオの value は読み戻せないので、選んでいる向きは変数で持つ
           A radio's value cannot be read back before show(), so the chosen direction lives in a variable */
        var selectedDirection = addDirectionPanel(spacingDialog, axisSpreads, function (spacingDirection) {
            selectedDirection = spacingDirection;
            btnShuffle.enabled = (selectedDirection.spread === "random");
            refreshPreview();
        });
        var readRatio = addRatioPanel(spacingDialog, refreshPreview);
        var readTotalCount = addCountPanel(spacingDialog, itemStates.length, refreshPreview);
        var readColorOptions = addColorPanel(spacingDialog, refreshPreview);

        var buttonRow = addButtonRow(spacingDialog);
        var btnShuffle = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.shuffle));
        btnShuffle.helpTip = getLabel(LABELS.tooltip.shuffle);
        btnShuffle.enabled = (selectedDirection.spread === "random");
        btnShuffle.onClick = function () {
            randomPool.length = 0;
            refreshPreview();
        };
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * 今の設定を読む
         * @returns {{direction: Object, ratio: number, count: number, randomPool: number[], colorBlend: Object}} 今の設定
         */
        function readSettings() {
            return {
                direction: selectedDirection,
                ratio: readRatio(),
                count: readTotalCount(),
                randomPool: randomPool,
                colorBlend: readColorOptions()
            };
        }

        /**
         * 今の設定でプレビューを描き直す
         * @returns {void}
         */
        function refreshPreview() {
            spacingPreview.update(readSettings());
        }

        btnOK.onClick = function () {
            committedSettings = readSettings();
            spacingDialog.close(1);
        };

        addVersionFooter(spacingDialog);
        prepareDialogWindow(spacingDialog, SCRIPT_NAME);
        refreshPreview();
        spacingDialog.show();

        /* プレビューを消してから本番を1回だけ流す（取り消しが1回で済む）
           Drop the preview, then run once for real so a single undo reverts it */
        if (committedSettings !== null) spacingPreview.commit(committedSettings);
        else spacingPreview.clear();
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var selectedItems = collectSelectionItems(app.activeDocument.selection, {
            accept: function () { return true; }, /* グループは中へ入らず1つとして扱う / keep groups whole */
            skipLocked: true,
            skipHidden: true
        });
        var usePreviewBounds = readUsePreviewBoundsPreference();
        var itemStates = captureItemStates(selectedItems, usePreviewBounds);
        if (itemStates.length < MIN_ITEM_COUNT) {
            alert(getLabel(LABELS.alert.needItems, { count: MIN_ITEM_COUNT }));
            return;
        }

        var axisSpreads = { x: getCenterSpread(itemStates, "x"), y: getCenterSpread(itemStates, "y") };
        if (isNearlySameCoordinate(axisSpreads.x, 0) && isNearlySameCoordinate(axisSpreads.y, 0)) {
            alert(getLabel(LABELS.alert.samePosition));
            return;
        }
        showSpacingDialog(itemStates, usePreviewBounds, axisSpreads);
    }

    main();

})();
