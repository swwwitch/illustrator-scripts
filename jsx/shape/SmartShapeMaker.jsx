#target illustrator
#targetengine "MyScriptEngine"
#include "../stroke-table/ColorPicker.jsx"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

1つのダイアログから、正円・多角形・星形・スーパー楕円・ルーローの三角形などのカスタム形状を作成します。
リアルタイムプレビューで、辺の数・幅・回転・詳細オプションを調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartShapeMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n005a7087f9c3

### Overview

Creates custom shapes — circle, polygon, star, superellipse, Reuleaux-style — from a single dialog.
A real-time preview lets you adjust the number of sides, the width, the rotation and the advanced options.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartShapeMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartShapeMaker";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.4.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartShapeMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartShapeMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n005a7087f9c3"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 各入力の初期値 / Default values of the inputs */
    var SHAPE_DEFAULTS = {
        size: "100",             /* 幅 / width */
        rotation: "90",          /* 回転角（度） / rotation angle in degrees */
        customSides: 12,         /* ［それ以外］の辺の数 / custom side count */
        circleAnchors: 4,        /* 円のアンカーポイント数 / anchor count of a circle */
        innerRatio: 30,          /* スターの第2半径（%） / star inner radius in percent */
        superExponent: 2.5,      /* スーパー楕円の指数 / superellipse exponent */
        strokeWidth: "1",        /* 線幅 / stroke width */
        opacity: 100,            /* 不透明度（%） / opacity in percent */
        cornerRadius: "15",      /* 角丸の半径 / corner radius */
        cornerRadiusRatio: 0.15, /* 角丸半径の既定値＝幅×この比率 / corner radius default ratio */
        smoothing: 60,           /* スムージング（%） / smoothing in percent */
        reuleauxAmount: 100,     /* ルーローの度合い（%） / Reuleaux amount in percent */
        roughenDetail: "1",      /* ラフ効果の詳細 / roughen detail */
        fitViewPercent: 65,      /* ［画面にフィット］で図形が占める割合（%） / share of the window the fitted shape fills */
        segmentStrokeWidth: 0.3  /* 分割時に線がないときの既定線幅（pt）
                                    fallback stroke width of split segments, in pt */
    };

    /* 入力の下限と上限 [min, max] / Lower and upper bounds of the inputs */
    var SHAPE_RANGES = {
        customSides: [3, 36],     /* 辺の数 / side count */
        innerRatio: [0, 100],     /* 第2半径（%） / inner radius */
        superExponent: [1.5, 6],  /* スーパー楕円の指数 / superellipse exponent */
        smoothing: [0, 150],      /* スムージング（%） / smoothing */
        reuleauxAmount: [0, 200], /* ルーローの度合い（%） / Reuleaux amount */
        opacity: [0, 100],        /* 不透明度（%） / opacity */
        fitViewPercent: [10, 100] /* ［画面にフィット］の割合（%） / fit-to-window share */
    };

    /* Illustratorが受け付ける表示倍率の範囲（3.125%〜6400%） / Zoom range Illustrator accepts */
    var VIEW_ZOOM_RANGE = [0.03125, 64];

    /* 図形生成の内部パラメーター / Internal parameters of the shape generation */
    var SHAPE_GEOMETRY = {
        superEllipsePoints: 8,    /* スーパー楕円のサンプル点数 / sample point count of a superellipse */
        superEllipseHandle: 0.35, /* スーパー楕円のハンドル長比率 / handle length ratio of a superellipse */
        smoothingArmFactor: 0.8   /* 角丸のアーム長係数 / arm length factor of the corner smoothing */
    };

    /* ［辺の数］ラジオの選択肢（0は円）。最後に［それ以外］の手動入力が続く
       Choices of the side-count radios, 0 being a circle; the custom field follows the last one */
    var SIDE_CHOICES = [0, 3, 4, 5, 6, 8];
    var CUSTOM_SIDES_INDEX = SIDE_CHOICES.length;

    /* 円のアンカーポイント数の選択肢 / Choices of the circle anchor count */
    var CIRCLE_ANCHOR_CHOICES = [2, 3, 4, 5, 6];

    /* 開いたときの辺の数（4＝正方形） / Side count selected when the dialog opens, 4 being a square */
    var DEFAULT_SIDES = 4;

    /* 回転をONにしたときの円の既定角度（アンカー数ごと） / Default angle of a circle per anchor count, applied when the rotation is turned on */
    var CIRCLE_ROTATION_DEFAULTS = { 2: 90, 3: 180, 4: 45, 5: 180, 6: 30 };

    /* 三角形の向きごとの回転角 / Rotation angle per triangle direction */
    var TRIANGLE_ANGLES = { right: -90, left: 90, down: 60 };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window and panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* パネルが多い密なダイアログなので、パネル内と行内は既定より詰める
       This dialog stacks many panels, so panels and rows are tighter than the defaults */
    var DENSE_PANEL_SPACING = 6;   /* パネル内の要素間隔 / spacing inside a panel */
    var ROW_SPACING         = 6;   /* 行内の標準間隔 / default spacing inside a row */
    var TIGHT_ROW_SPACING   = 4;   /* 要素の多い行の間隔 / spacing of a crowded row */
    var WIDE_ROW_SPACING    = 10;  /* ラジオを横に並べる行の間隔 / spacing of a row of radios */

    /* コントロールの寸法 / Control sizes */
    var SWATCH_SIZE        = 16;   /* カラースウォッチの一辺（px） / color swatch size in px */
    var SLIDER_WIDTH       = 200;  /* 標準スライダー幅 / default slider width */
    var SHORT_SLIDER_WIDTH = 150;  /* 短いスライダー幅 / short slider width */
    var INLINE_SLIDER_WIDTH = 100; /* 項目と同じ行に置くスライダー幅（［辺の数］［スーパー楕円］） / slider width when it shares a row with its control */

    /**
     * ウィンドウの共通レイアウトを適用する。
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時はWINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通レイアウトを適用する。
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時はPANEL_SPACING）
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
     * 行グループの共通レイアウトを適用する（ボタン列など）。
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|Array} [alignment] - 整列指定（省略時は"left"）
     * @param {number} [spacing] - 要素間隔（省略時はPANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 見出し付きパネルを追加し、共通レイアウトを適用する。
     * @param {object} parent - 追加先のコンテナ
     * @param {string} labelText - パネルの見出し
     * @param {number} [spacing] - 要素間隔（省略時はDENSE_PANEL_SPACING）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelText, spacing) {
        var panel = parent.add("panel", undefined, labelText);
        setupPanel(panel, (typeof spacing === "number") ? spacing : DENSE_PANEL_SPACING);
        return panel;
    }

    /**
     * 左揃え・上下中央の行グループを追加する。
     * @param {object} parent - 追加先のコンテナ
     * @param {number} [spacing] - 要素間隔（省略時はROW_SPACING）
     * @returns {Group} 追加したグループ
     */
    function addControlRow(parent, spacing) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ["left", "center"];
        row.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
        return row;
    }

    /**
     * 左に∧∨を付けた数値入力欄を追加する（∧∨と入力欄は隙間0で突き合わせる）。
     * ↑↓キーも∧∨と同じ処理で増減する。増減は確定した編集として扱い、onChange があれば丸めまで、
     * なければ onChanging でプレビューだけ反映する
     * @param {object} parent - 追加先のコンテナ
     * @param {string|number} value - 初期値
     * @param {number} charCount - 入力欄の文字幅
     * @param {object} [stepOptions] - min / max / integer（addStepper() の stepOptions）
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addNumberField(parent, value, charCount, stepOptions) {
        var stepperFieldGroup = parent.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        stepOptions = stepOptions || {};
        stepOptions.onStep = function (numberInput) {
            var editHandler = (typeof numberInput.onChange === "function") ? numberInput.onChange : numberInput.onChanging;
            if (typeof editHandler === "function") editHandler();
        };
        var editText;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return editText; }, stepOptions);
        editText = stepperFieldGroup.add("edittext", undefined, String(value));
        editText.characters = charCount;
        editText.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(editText, stepperGroup);
        return editText;
    }

    /**
     * 幅を指定したスライダーを追加する。
     * @param {object} parent - 追加先のコンテナ
     * @param {number} value - 初期値
     * @param {Array<number>} range - [下限, 上限]
     * @param {number} width - スライダー幅（px）
     * @returns {Slider} 追加したスライダー
     */
    function addSlider(parent, value, range, width) {
        var slider = parent.add("slider", undefined, value, range[0], range[1]);
        slider.preferredSize.width = width;
        return slider;
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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
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
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "基本図形の作成", en: "Create Basic Shapes" },
            colorPicker: { ja: "カラーピッカー", en: "Color Picker" }
        },
        panel: {
            sides: { ja: "辺の数", en: "Sides" },
            rotation: { ja: "回転", en: "Rotate" },
            triangle: { ja: "三角形", en: "Triangle" },
            fillAndStroke: { ja: "塗りと線", en: "Fill & Stroke" },
            width: { ja: "幅", en: "Width" },
            star: { ja: "スター", en: "Star" },
            circle: { ja: "円", en: "Circle" },
            anchorCount: { ja: "アンカーポイント数", en: "Anchor Count" },
            anchorOps: { ja: "アンカーポイントの操作", en: "Anchor Point Operations" },
            cornerSmoothing: { ja: "角丸", en: "Rounded Corners" },
            option: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            superEllipse: { ja: "スーパー楕円", en: "Superellipse" },
            star: { ja: "スター", en: "Star" },
            pentagram: { ja: "五芒星", en: "Pentagram" },
            fill: { ja: "塗り", en: "Fill" },
            stroke: { ja: "線", en: "Stroke" },
            cornerRadius: { ja: "半径", en: "Radius" },
            liveShape: { ja: "ライブシェイプに変換", en: "Convert to Live Shape" },
            reuleaux: { ja: "ルーロー（定幅図形）", en: "Reuleaux (Constant-Width)" },
            splitAtAnchors: { ja: "アンカーポイントで分割", en: "Split at Anchor Points" },
            roughenAnchors: { ja: "ラフ効果で追加", en: "Add Anchors (Roughen)" },
            fitView: { ja: "画面にフィット", en: "Fit to Window" }
        },
        radio: {
            circleWithZero: { ja: "0（円）", en: "0 (Circle)" },
            squareWithFour: { ja: "4（正方形）", en: "4 (Square)" },
            triangleRight: { ja: "右", en: "Right" },
            triangleLeft: { ja: "左", en: "Left" },
            triangleDown: { ja: "下", en: "Down" },
            capButt: { ja: "なし", en: "Butt" },
            capRound: { ja: "丸型", en: "Round" },
            capProjecting: { ja: "突出", en: "Projecting" }
        },
        fieldLabel: {
            innerRadius: { ja: "第2半径", en: "Inner Radius" },
            opacity: { ja: "不透明度", en: "Opacity" },
            smoothing: { ja: "スムージング", en: "Smoothing" },
            strokeCap: { ja: "線端", en: "Cap" }
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
            circle: { ja: "円を作成します（E）", en: "Creates a circle (E)" },
            customSides: { ja: "辺の数を3〜36で指定します", en: "Sets a side count from 3 to 36" },
            rotate: { ja: "オンにすると、入力した角度だけ図形を回転します（A）", en: "When on, rotates the shape by the entered angle (A)" },
            triangleRight: { ja: "右向きの三角形にします（R）", en: "Points the triangle to the right (R)" },
            triangleLeft: { ja: "左向きの三角形にします（L）", en: "Points the triangle to the left (L)" },
            triangleDown: { ja: "下向きの三角形にします（B）", en: "Points the triangle down (B)" },
            sideChoice: { ja: "option（Alt）＋数字キーでも選べます", en: "Also selectable with Option (Alt) + the number key" },
            width: {
                ja: "円は直径、正方形は1辺の長さ、そのほかの多角形とスターは外接円の直径です。",
                en: "The diameter of a circle, the edge length of a square, and the circumscribed diameter of other polygons and stars."
            },
            fitView: {
                ja: "［幅］を変えたときに、作成する図形が収まるよう表示倍率を合わせます。",
                en: "Refits the view to the shape being created whenever the width changes."
            },
            fitViewPercent: {
                ja: "ウィンドウに対する図形の大きさ（100%でいっぱい）",
                en: "Size of the shape relative to the window; 100% fills it"
            },
            star: { ja: "スターにします（S）", en: "Makes a star (S)" },
            pentagram: {
                ja: "第2半径を黄金比で決めた、5辺のスターにします（P）",
                en: "Makes a five-pointed star whose inner radius follows the golden ratio (P)"
            },
            innerRadius: { ja: "外側の半径に対する内側の半径の比率", en: "Inner radius as a percentage of the outer radius" },
            superEllipse: { ja: "円を、四角形に近づけた丸い形にします", en: "Turns the circle into a squarish round shape" },
            superExponent: { ja: "指数。大きいほど四角形に近づきます", en: "Exponent; higher values look more like a square" },
            cornerRadius: {
                ja: "正方形の角を丸くします。スムージングが0のときは［角を丸くする］効果を使います。",
                en: "Rounds the corners of the square. With zero smoothing, the Round Corners effect is used."
            },
            smoothing: { ja: "角丸を直線部へなめらかにつなぐ量", en: "How smoothly the rounded corners blend into the straight edges" },
            roughenAnchors: {
                ja: "変形量0の［ラフ］効果でアンカーポイントを追加します。数値は［詳細］です。\n多角形で1のときは［アンカーポイントの追加］を使うため、プレビューには出ません。",
                en: "Adds anchor points with a Roughen effect of size 0. The number is its Detail.\nWith a polygon, 1 uses Add Anchor Points instead, which the preview does not show."
            },
            splitAtAnchors: { ja: "アンカーポイントごとに開いたパスへ分割します（D）", en: "Splits the shape into open paths at each anchor point (D)" },
            liveShape: { ja: "確定後に［シェイプに変換］でライブシェイプにします", en: "Converts the result into a live shape with Convert to Shape after OK" },
            reuleaux: {
                ja: "奇数辺の多角形の各辺を円弧にして、定幅図形にします。",
                en: "Turns each edge of an odd-sided polygon into an arc to make a constant-width shape."
            },
            reuleauxAmount: { ja: "円弧のふくらみ（100%で定幅図形）", en: "Arc bulge; 100% gives a constant-width shape" },
            preview: { ja: "プレビュー表示とアウトライン表示を切り替えます", en: "Switches the document between Preview and Outline view" },
            colorSwatch: { ja: "クリックしてカラーを選びます", en: "Click to choose a color" },
            strokeWidth: { ja: "線幅", en: "Stroke weight" },
            opacity: { ja: "Shiftキーを押しながらドラッグすると10%刻み", en: "Shift-drag to snap to 10% steps" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        alert: {
            previewError: { ja: "プレビューエラー：", en: "Preview error: " },
            finalError: { ja: "図形の作成エラー：", en: "Error while creating the shape: " },
            noDocument: {
                ja: "ドキュメントを開いてから実行してください。",
                en: "Open a document before running this script."
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

    // =========================================
    // セッション状態 / Session state
    // =========================================

    /* Illustratorの起動中だけダイアログの値を保持する（#targetengineの常駐エンジンを利用）。旧版の $.global のキーを1度だけ読み継ぐ
       Dialog values are kept only while Illustrator is running, on the engine named by #targetengine; the old $.global key is read once */
    var LEGACY_SESSION_STATE_KEY = "__SmartShapeMaker_State__";
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_STATE_KEY] || null; }
    });

    /**
     * 前回のセッション状態を読み込む。
     * 項目ごとの型は applyStateToUI() が確かめるため、既定値は {}（中身を問わず受け取る）
     * @returns {object} 保存されていた状態。未保存なら空のオブジェクト
     */
    function getSessionState() {
        return settingsStore.load({});
    }

    // =========================================
    // 確定時に参照する状態 / State referenced on finalize
    // =========================================

    var previewShape = null;                   /* プレビュー中の図形 / the shape currently previewed */
    var applyLiveShape = true;                 /* ライブシェイプ化するか / whether to convert to a live shape */
    var roughenAnchorsDetail = 0;              /* ラフ効果の詳細（0で無効） / roughen detail, 0 disables it */
    var roughenAnchorsUseMenuFallback = false; /* メニューコマンドで代替するか / whether to fall back to the menu command */

    // =========================================
    // 数学ヘルパー / Math helpers
    // =========================================

    /**
     * 数値を有効範囲に収める。空欄や数値でない入力は既定値に戻す。
     * @param {string|number} value - 入力値
     * @param {Array<number>} range - [下限, 上限]
     * @param {number} fallbackValue - 数値として読めないときの既定値
     * @param {boolean} [asInteger] - 整数に丸めるかどうか
     * @returns {number} 範囲内に収めた数値
     */
    function clampNumber(value, range, fallbackValue, asInteger) {
        /* 入力途中の空欄はNumber()が0になってしまうので既定値に戻す
           An empty field would become 0 through Number(), so fall back to the default */
        if (typeof value === "string" && !/\S/.test(value)) value = fallbackValue;
        value = Number(value);
        if (isNaN(value)) value = fallbackValue;
        if (asInteger) value = Math.round(value);
        if (value < range[0]) value = range[0];
        if (value > range[1]) value = range[1];
        return value;
    }

    /**
     * SHAPE_RANGES と SHAPE_DEFAULTS の同じキーを使い、整数に丸めて有効範囲に収める関数を作る。
     * @param {string} settingKey - SHAPE_RANGES と SHAPE_DEFAULTS のキー
     * @returns {function} 入力値を受け取り、範囲内の整数を返す関数
     */
    function makeRangeClamp(settingKey) {
        return function (value) {
            return clampNumber(value, SHAPE_RANGES[settingKey], SHAPE_DEFAULTS[settingKey], true);
        };
    }

    /**
     * 角度表示を整える（小数第3位まで、整数なら整数表記）。
     * @param {number} angle - 角度（度）
     * @returns {string} 表示用の角度文字列
     */
    function formatAngle(angle) {
        var rounded = Math.round(angle * 1000) / 1000;
        return (rounded % 1 === 0) ? String(Math.round(rounded)) : String(rounded);
    }

    /**
     * 数値の符号を返す（0と-0はそのまま返す）。
     * @param {number} value - 対象の数値
     * @returns {number} 1、-1、または0
     */
    function signOf(value) {
        return ((value > 0) - (value < 0)) || +value;
    }

    /**
     * 円のアンカーポイント数を2以上の整数にそろえる。
     * @param {number} anchorCount - 指定されたアンカーポイント数
     * @returns {number} 2以上の整数（数値でなければ既定値）
     */
    function normalizeCircleAnchorCount(anchorCount) {
        var normalizedCount = (typeof anchorCount === 'number') ? Math.round(anchorCount) : SHAPE_DEFAULTS.circleAnchors;
        return (normalizedCount < 2) ? 2 : normalizedCount;
    }

    /**
     * 曲率半径を保つベジェハンドルの長さを求める。
     * @param {number} arm - コーナーから直線部までのアーム長
     * @param {number} radius - 角丸の半径
     * @returns {number} ハンドルの長さ
     */
    function computeCornerHandleLength(arm, radius) {
        if (arm <= 0 || radius <= 0) return 0;
        var kappa = 16 / (3 * Math.sqrt(2));
        var discriminant = radius * (8 * kappa * arm + kappa * kappa * radius);
        return ((4 * arm + kappa * radius) - Math.sqrt(discriminant)) / 2;
    }

    // =========================================
    // プレビュー管理 / Preview manager
    // =========================================

    /**
     * プレビューの適用とUndoによる巻き戻しを管理する。
     * rollback()でプレビューをすべて取り消し、confirm()で巻き戻したうえで確定処理を1回だけ実行する。
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;

        /**
         * プレビュー処理を1ステップ実行し、Undo段数を記録する。
         * @param {function} stepAction - プレビューを描画する処理
         * @returns {void}
         */
        this.addStep = function (stepAction) {
            try {
                stepAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(getLabel(LABELS.alert.previewError) + e);
            }
        };

        /**
         * 記録済みのプレビュー処理をすべて取り消す。
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try { app.undo(); } catch (e) { break; }
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * プレビューを巻き戻したうえで確定処理を1回だけ実行する。
         * @param {function} finalAction - 確定時に実行する処理
         * @returns {void}
         */
        this.confirm = function (finalAction) {
            if (finalAction) {
                this.rollback();
                try { finalAction(); } catch (e) { alert(getLabel(LABELS.alert.finalError) + e); }
            }
            this.undoDepth = 0;
        };
    }

    // =========================================
    // UIヘルパー / UI helpers
    // =========================================

    /**
     * 入力途中で書き戻すと編集できなくなる文字列かどうかを判定する。
     * 空欄・符号だけ・小数点で終わる値は、まだ確定していないものとして扱う。
     * @param {string} text - 入力欄の文字列
     * @returns {boolean} 入力途中ならtrue
     */
    function isPartialNumberInput(text) {
        return /^\s*[+-]?\s*$/.test(text) || /\.\s*$/.test(text);
    }

    /**
     * 選択中のラジオボタンのインデックスを返す。
     * @param {Array} radios - ラジオボタンの配列
     * @returns {number} 選択中のインデックス（未選択なら-1）
     */
    function getSelectedRadioIndex(radios) {
        for (var i = 0; i < radios.length; i++) {
            if (radios[i].value) return i;
        }
        return -1;
    }

    /**
     * 指定したインデックスのラジオボタンだけをONにする。
     * ［それ以外］は別グループにあるため排他が効かず、全件を明示的に解除する必要がある。
     * @param {Array} radios - ラジオボタンの配列
     * @param {number} selectedIndex - ONにするインデックス
     * @returns {void}
     */
    function selectRadio(radios, selectedIndex) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (i === selectedIndex);
        }
    }

    /**
     * 選択肢の配列から値の位置を探す。
     * @param {Array<number>} choices - 選択肢の配列
     * @param {number} value - 探す値
     * @returns {number} 見つかった位置（無ければ-1）
     */
    function findChoiceIndex(choices, value) {
        for (var i = 0; i < choices.length; i++) {
            if (choices[i] === value) return i;
        }
        return -1;
    }

    /**
     * 複数のコントロールの有効・無効をまとめて切り替える。
     * @param {Array<object>} controls - 対象のコントロール
     * @param {boolean} isEnabled - 有効にするかどうか
     * @returns {void}
     */
    function setControlsEnabled(controls, isEnabled) {
        for (var i = 0; i < controls.length; i++) {
            controls[i].enabled = isEnabled;
            /* 数値欄の∧∨も一緒に切り替えて描き直す / toggle and redraw the field's stepper too */
            if (controls[i].stepperGroup) {
                controls[i].stepperGroup.enabled = isEnabled;
                redrawSteppersIn(controls[i].stepperGroup);
            }
        }
    }

    // =========================================
    // パス生成 / Path builders
    // =========================================

    /**
     * 図形を作る前にアクティブレイヤーの編集を許可し、ドキュメントウィンドウの中心を得る。
     * @param {Document} doc - 対象ドキュメント
     * @returns {{layer: Layer, centerX: number, centerY: number}} レイヤーと中心座標
     */
    function prepareActiveLayer(doc) {
        var layer = doc.activeLayer;
        layer.locked = false;
        layer.visible = true;
        var viewCenter = doc.activeView.centerPoint;
        return { layer: layer, centerX: viewCenter[0], centerY: viewCenter[1] };
    }

    /**
     * 自前で組んだパスの見た目をcreateShapeの既定に合わせる。
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} pathItem - 対象のパス
     * @returns {PathItem} 見た目を適用したパス
     */
    function applyDefaultAppearance(doc, pathItem) {
        pathItem.filled = true;
        pathItem.fillColor = doc.defaultFillColor;
        pathItem.stroked = false;
        return pathItem;
    }

    /**
     * 座標の配列から閉じたパスを作る。
     * @param {Layer} layer - 追加先のレイヤー
     * @param {Array} anchors - アンカー座標の配列
     * @returns {PathItem} 作成したパス
     */
    function createClosedPath(layer, anchors) {
        var pathItem = layer.pathItems.add();
        pathItem.setEntirePath(anchors);
        pathItem.closed = true;
        return pathItem;
    }

    /**
     * スーパー楕円のパスを作成する（サンプル点とスムーズハンドルで構成）。
     * @param {Document} doc - 対象ドキュメント
     * @param {number} sizePt - 幅（pt）
     * @param {number} exponent - スーパー楕円の指数
     * @param {number} [pointCount] - サンプル点数
     * @returns {PathItem} 作成したパス
     */
    function createSuperellipsePath(doc, sizePt, exponent, pointCount) {
        exponent = (typeof exponent === 'number' && exponent > 0) ? exponent : SHAPE_DEFAULTS.superExponent;
        pointCount = (typeof pointCount === 'number' && pointCount >= SHAPE_GEOMETRY.superEllipsePoints)
            ? Math.round(pointCount)
            : SHAPE_GEOMETRY.superEllipsePoints;

        var placement = prepareActiveLayer(doc);
        var radius = sizePt / 2;

        var anchors = [];
        for (var i = 0; i < pointCount; i++) {
            var theta = (Math.PI * 2 * i) / pointCount;
            var cosTheta = Math.cos(theta);
            var sinTheta = Math.sin(theta);
            var x = Math.pow(Math.abs(cosTheta), 2 / exponent) * radius * signOf(cosTheta);
            var y = Math.pow(Math.abs(sinTheta), 2 / exponent) * radius * signOf(sinTheta);
            anchors.push([x + placement.centerX, y + placement.centerY]);
        }

        var pathItem = createClosedPath(placement.layer, anchors);

        /* 接線方向（次点−前点）にスムーズハンドルを置く / Place smooth handles along the tangents (next - prev) */
        var pathPoints = pathItem.pathPoints;
        var anchorCount = pathPoints.length;
        for (var k = 0; anchorCount >= 4 && k < anchorCount; k++) {
            var prevAnchor = anchors[(k - 1 + anchorCount) % anchorCount];
            var currentAnchor = anchors[k];
            var nextAnchor = anchors[(k + 1) % anchorCount];

            var tangentX = nextAnchor[0] - prevAnchor[0];
            var tangentY = nextAnchor[1] - prevAnchor[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY);
            if (tangentLength === 0) continue;
            tangentX /= tangentLength;
            tangentY /= tangentLength;

            /* 前後のセグメント長のうち短いほうに合わせる / Follow the shorter of the neighbouring segments */
            var prevDeltaX = currentAnchor[0] - prevAnchor[0];
            var prevDeltaY = currentAnchor[1] - prevAnchor[1];
            var nextDeltaX = nextAnchor[0] - currentAnchor[0];
            var nextDeltaY = nextAnchor[1] - currentAnchor[1];
            var prevLength = Math.sqrt(prevDeltaX * prevDeltaX + prevDeltaY * prevDeltaY);
            var nextLength = Math.sqrt(nextDeltaX * nextDeltaX + nextDeltaY * nextDeltaY);
            var handleLength = Math.min(prevLength, nextLength) * SHAPE_GEOMETRY.superEllipseHandle;

            pathPoints[k].anchor = currentAnchor;
            pathPoints[k].leftDirection = [currentAnchor[0] - tangentX * handleLength, currentAnchor[1] - tangentY * handleLength];
            pathPoints[k].rightDirection = [currentAnchor[0] + tangentX * handleLength, currentAnchor[1] + tangentY * handleLength];
            pathPoints[k].pointType = PointType.SMOOTH;
        }

        return applyDefaultAppearance(doc, pathItem);
    }

    /**
     * 指定したアンカー数（2以上）で円形の閉じたパスを作成する。
     * ハンドル長は k = 4/3 * tan(pi/(2N)) を用いる。
     * @param {Document} doc - 対象ドキュメント
     * @param {number} sizePt - 直径（pt）
     * @param {number} anchorCount - アンカーポイント数
     * @returns {PathItem} 作成したパス
     */
    function createCirclePathWithNAnchors(doc, sizePt, anchorCount) {
        var placement = prepareActiveLayer(doc);
        var radius = sizePt / 2;
        anchorCount = normalizeCircleAnchorCount(anchorCount);

        var handleLength = radius * (4 / 3) * Math.tan(Math.PI / (2 * anchorCount));

        /* 頂点が真上に来るよう-90°から並べる / Start at -90 degrees so one anchor sits on top */
        var angles = [];
        var anchors = [];
        for (var i = 0; i < anchorCount; i++) {
            var angle = (-Math.PI / 2) + (2 * Math.PI * i) / anchorCount;
            angles.push(angle);
            anchors.push([placement.centerX + radius * Math.cos(angle), placement.centerY + radius * Math.sin(angle)]);
        }

        var pathItem = createClosedPath(placement.layer, anchors);

        /* 接線は半径に直交する [-sin, cos] / The tangent is perpendicular to the radius */
        var pathPoints = pathItem.pathPoints;
        for (var j = 0; j < anchorCount; j++) {
            var tangentX = -Math.sin(angles[j]);
            var tangentY = Math.cos(angles[j]);

            pathPoints[j].anchor = anchors[j];
            pathPoints[j].leftDirection = [anchors[j][0] - tangentX * handleLength, anchors[j][1] - tangentY * handleLength];
            pathPoints[j].rightDirection = [anchors[j][0] + tangentX * handleLength, anchors[j][1] + tangentY * handleLength];
            pathPoints[j].pointType = PointType.SMOOTH;
        }

        return applyDefaultAppearance(doc, pathItem);
    }

    /**
     * 奇数辺の正多角形の各辺を円弧に置き換えてルーロー図形にする。
     * 参考スクリプト reuleaux_polygon.jsx のロジックを移植。
     * @param {PathItem} pathItem - 対象のパス（奇数個のアンカーを持つ多角形）
     * @param {number} amount - 度合い（1.0が標準、0.0〜2.0）
     * @returns {PathItem} 変換後のパス
     */
    function applyReuleauxToPolygon(pathItem, amount) {
        if (!pathItem || pathItem.typename !== "PathItem") return pathItem;
        if (!pathItem.pathPoints || pathItem.pathPoints.length < 3) return pathItem;
        var pathPoints = pathItem.pathPoints;
        var pointCount = pathPoints.length;
        if (pointCount % 2 === 0) return pathItem; /* 奇数辺のみ / odd side counts only */

        amount = clampNumber(amount, [0, 2], 1);

        /* アンカー座標をキャッシュ / Cache the anchor coordinates */
        var anchorCoords = [];
        for (var i = 0; i < pointCount; i++) {
            anchorCoords.push([pathPoints[i].anchor[0], pathPoints[i].anchor[1]]);
        }

        for (i = 0; i < pointCount; i++) {
            var startIndex = i;
            var endIndex = (i + 1) % pointCount;
            var centerIndex = (i + Math.floor((pointCount + 1) / 2)) % pointCount;

            var arcStart = anchorCoords[startIndex];
            var arcEnd = anchorCoords[endIndex];
            var arcCenter = anchorCoords[centerIndex];

            var vectorToStart = [arcStart[0] - arcCenter[0], arcStart[1] - arcCenter[1]];
            var vectorToEnd = [arcEnd[0] - arcCenter[0], arcEnd[1] - arcCenter[1]];

            var radiusToStart = Math.sqrt(vectorToStart[0] * vectorToStart[0] + vectorToStart[1] * vectorToStart[1]);
            var radiusToEnd = Math.sqrt(vectorToEnd[0] * vectorToEnd[0] + vectorToEnd[1] * vectorToEnd[1]);
            if (radiusToStart === 0 || radiusToEnd === 0) continue;
            var arcRadius = (radiusToStart + radiusToEnd) / 2;

            var dotProduct = vectorToStart[0] * vectorToEnd[0] + vectorToStart[1] * vectorToEnd[1];
            var cosTheta = clampNumber(dotProduct / (radiusToStart * radiusToEnd), [-1, 1], 0);
            var deltaTheta = Math.acos(cosTheta);

            /* 円弧をベジェで近似したハンドル長に度合いを掛ける / Bezier approximation of the arc, scaled by the amount */
            var handleLength = arcRadius * (4 / 3) * Math.tan(deltaTheta / 4) * amount;

            /* 中心から見た回り方向に合わせて接線の向きを決める / Pick the tangent side from the winding around the arc center */
            var crossZ = vectorToStart[0] * vectorToEnd[1] - vectorToStart[1] * vectorToEnd[0];
            var tangentSign = (crossZ > 0) ? 1 : -1;
            var tangentStart = [-vectorToStart[1] * tangentSign, vectorToStart[0] * tangentSign];
            var tangentEnd = [vectorToEnd[1] * tangentSign, -vectorToEnd[0] * tangentSign];

            pathPoints[startIndex].rightDirection = [
                arcStart[0] + handleLength * tangentStart[0] / radiusToStart,
                arcStart[1] + handleLength * tangentStart[1] / radiusToStart
            ];
            pathPoints[endIndex].leftDirection = [
                arcEnd[0] + handleLength * tangentEnd[0] / radiusToEnd,
                arcEnd[1] + handleLength * tangentEnd[1] / radiusToEnd
            ];

            pathPoints[startIndex].pointType = PointType.CORNER;
            pathPoints[endIndex].pointType = PointType.CORNER;
        }

        pathItem.closed = true;
        return pathItem;
    }

    /**
     * スムージングを効かせた角丸長方形のパスを作成する。
     * 黒野真吾さんの corner_smoothing.jsx をもとにしている。
     * @param {Document} doc - 対象ドキュメント
     * @param {number} left - 左端のX座標
     * @param {number} top - 上端のY座標
     * @param {number} rectWidth - 幅
     * @param {number} rectHeight - 高さ
     * @param {number} radius - 角丸の半径
     * @param {number} smoothing - スムージング量（0以上、1.0＝100%）
     * @returns {PathItem} 作成したパス
     */
    function buildSmoothedRect(doc, left, top, rectWidth, rectHeight, radius, smoothing) {
        radius = Math.min(Math.abs(radius), rectWidth / 2, rectHeight / 2);
        smoothing = Math.max(0, smoothing);

        var armLength = radius * (1 + SHAPE_GEOMETRY.smoothingArmFactor * smoothing);
        var armX = Math.min(armLength, rectWidth / 2);
        var armY = Math.min(armLength, rectHeight / 2);

        var handleX = computeCornerHandleLength(armX, radius);
        var handleY = computeCornerHandleLength(armY, radius);

        var edgeLeft = left;
        var edgeTop = top;
        var edgeRight = left + rectWidth;
        var edgeBottom = top - rectHeight;
        var midX = (edgeLeft + edgeRight) / 2;
        var midY = (edgeTop + edgeBottom) / 2;

        var mergeTopBottom = (armLength >= rectWidth / 2 - 0.01);
        var mergeSides = (armLength >= rectHeight / 2 - 0.01);

        var pointSpecs = [];

        /**
         * アンカーと左右ハンドルの組を記録する。
         * @param {Array} anchor - アンカー座標
         * @param {Array} leftHandle - 左方向ハンドルの座標
         * @param {Array} rightHandle - 右方向ハンドルの座標
         * @returns {void}
         */
        function addPoint(anchor, leftHandle, rightHandle) {
            pointSpecs.push({ anchor: anchor, leftHandle: leftHandle, rightHandle: rightHandle });
        }

        /* 上辺 / Top edge */
        if (mergeTopBottom) {
            addPoint([midX, edgeTop], [midX - handleX, edgeTop], [midX + handleX, edgeTop]);
        } else {
            addPoint([edgeLeft + armX, edgeTop], [edgeLeft + armX - handleX, edgeTop], [edgeLeft + armX, edgeTop]);
            addPoint([edgeRight - armX, edgeTop], [edgeRight - armX, edgeTop], [edgeRight - armX + handleX, edgeTop]);
        }
        /* 右辺 / Right edge */
        if (mergeSides) {
            addPoint([edgeRight, midY], [edgeRight, midY + handleY], [edgeRight, midY - handleY]);
        } else {
            addPoint([edgeRight, edgeTop - armY], [edgeRight, edgeTop - armY + handleY], [edgeRight, edgeTop - armY]);
            addPoint([edgeRight, edgeBottom + armY], [edgeRight, edgeBottom + armY], [edgeRight, edgeBottom + armY - handleY]);
        }
        /* 下辺 / Bottom edge */
        if (mergeTopBottom) {
            addPoint([midX, edgeBottom], [midX + handleX, edgeBottom], [midX - handleX, edgeBottom]);
        } else {
            addPoint([edgeRight - armX, edgeBottom], [edgeRight - armX + handleX, edgeBottom], [edgeRight - armX, edgeBottom]);
            addPoint([edgeLeft + armX, edgeBottom], [edgeLeft + armX, edgeBottom], [edgeLeft + armX - handleX, edgeBottom]);
        }
        /* 左辺 / Left edge */
        if (mergeSides) {
            addPoint([edgeLeft, midY], [edgeLeft, midY - handleY], [edgeLeft, midY + handleY]);
        } else {
            addPoint([edgeLeft, edgeBottom + armY], [edgeLeft, edgeBottom + armY - handleY], [edgeLeft, edgeBottom + armY]);
            addPoint([edgeLeft, edgeTop - armY], [edgeLeft, edgeTop - armY], [edgeLeft, edgeTop - armY + handleY]);
        }

        var layer = doc.activeLayer;
        var pathItem = layer.pathItems.add();
        pathItem.closed = true;

        for (var i = 0; i < pointSpecs.length; i++) {
            var pathPoint = pathItem.pathPoints.add();
            pathPoint.anchor = pointSpecs[i].anchor;
            pathPoint.leftDirection = pointSpecs[i].leftHandle;
            pathPoint.rightDirection = pointSpecs[i].rightHandle;
            pathPoint.pointType = PointType.CORNER;
        }

        return pathItem;
    }

    /**
     * CMYKカラーを作る。
     * @param {number} cyan - C（0〜100）
     * @param {number} magenta - M（0〜100）
     * @param {number} yellow - Y（0〜100）
     * @param {number} black - K（0〜100）
     * @returns {CMYKColor} 作成したカラー
     */
    function makeCmykColor(cyan, magenta, yellow, black) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cyan;
        cmykColor.magenta = magenta;
        cmykColor.yellow = yellow;
        cmykColor.black = black;
        return cmykColor;
    }

    /**
     * RGBカラーを作る。
     * @param {number} red - R（0〜255）
     * @param {number} green - G（0〜255）
     * @param {number} blue - B（0〜255）
     * @returns {RGBColor} 作成したカラー
     */
    function makeRgbColor(red, green, blue) {
        var rgbColor = new RGBColor();
        rgbColor.red = red;
        rgbColor.green = green;
        rgbColor.blue = blue;
        return rgbColor;
    }

    /**
     * ドキュメントのカラーモードに合わせた黒を作る。
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラーオブジェクト
     */
    function createBlackColor(doc) {
        if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK) return makeCmykColor(0, 0, 0, 100);
        return makeRgbColor(0, 0, 0);
    }

    /**
     * 閉じたパスをアンカーポイントごとに分割し、開いたパスのグループにする。
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} pathItem - 分割元のパス
     * @param {object} strokeOptions - 線の設定 {enabled, color, widthPt}
     * @param {StrokeCap} strokeCap - 線端の種類
     * @returns {GroupItem|PathItem} 分割後のグループ（分割できない場合は元のパス）
     */
    function splitPathAtAnchors(doc, pathItem, strokeOptions, strokeCap) {
        if (!pathItem || !pathItem.pathPoints || pathItem.pathPoints.length < 2) return pathItem;

        var layer = doc.activeLayer;
        var segmentGroup = layer.groupItems.add();

        var pathPoints = pathItem.pathPoints;
        var pointCount = pathPoints.length;
        var isClosed = pathItem.closed;

        for (var i = 0; i < pointCount; i++) {
            var j = i + 1;
            if (j >= pointCount) {
                if (!isClosed) break;
                j = 0;
            }

            var startPoint = pathPoints[i];
            var endPoint = pathPoints[j];

            /* このセグメント用に開いたパスを作る / Create an open path for this segment */
            var segmentPath = segmentGroup.pathItems.add();
            segmentPath.closed = false;

            /* 先にアンカーを設定 / Set the anchors first */
            segmentPath.setEntirePath([startPoint.anchor, endPoint.anchor]);

            /* ハンドルを引き継ぐ。2点の開いたパスでは始点はright、終点はleftを使う
               Copy the handles: the start point uses rightDirection, the end point uses leftDirection */
            segmentPath.pathPoints[0].leftDirection = startPoint.anchor;
            segmentPath.pathPoints[0].rightDirection = startPoint.rightDirection;
            segmentPath.pathPoints[0].pointType = startPoint.pointType;

            segmentPath.pathPoints[1].leftDirection = endPoint.leftDirection;
            segmentPath.pathPoints[1].rightDirection = endPoint.anchor;
            segmentPath.pathPoints[1].pointType = endPoint.pointType;

            /* 開いたパスに塗りは合わないので線だけにする。線がOFFなら黒の細線で見えるようにする
               Open paths get a stroke, not a fill; without a stroke option they fall back to a thin black line */
            segmentPath.filled = false;
            segmentPath.stroked = true;
            var hasStroke = !!(strokeOptions && strokeOptions.enabled);
            /* 線幅欄が数値でないとstrokeWidthへの代入が例外になる。その線だけ既定の見た目で残す
               A non-numeric stroke width throws on assignment; that segment then keeps the default look */
            try {
                segmentPath.strokeColor = hasStroke ? strokeOptions.color : createBlackColor(doc);
                segmentPath.strokeWidth = hasStroke ? strokeOptions.widthPt : SHAPE_DEFAULTS.segmentStrokeWidth;
                if (strokeCap) segmentPath.strokeCap = strokeCap;
            } catch (e) { }
        }

        /* 元のパスを削除 / Remove the original path */
        pathItem.remove();

        return segmentGroup;
    }

    /**
     * 図形生成のパラメーター一式。ダイアログのgetCurrentShapeParams()が組み立てる。
     * @typedef {object} ShapeParams
     * @property {number} size - 幅（pt）
     * @property {number} sides - 辺の数（0は円）
     * @property {boolean} isStar - スターにするか
     * @property {number} innerRatio - スターの第2半径（%）
     * @property {boolean} rotateEnabled - 回転を適用するか
     * @property {number} angle - 回転角（度）
     * @property {boolean} splitAtAnchors - アンカーポイントで分割するか
     * @property {StrokeCap} strokeCap - 分割時の線端の種類
     * @property {boolean} useSuperEllipse - スーパー楕円にするか
     * @property {number} superExponent - スーパー楕円の指数
     * @property {number} circleAnchorCount - 円のアンカーポイント数
     * @property {boolean} useReuleaux - ルーロー図形にするか
     * @property {number} reuleauxAmount - ルーローの度合い（1.0が標準）
     * @property {object} fillOptions - 塗りの設定 {enabled, color}
     * @property {object} strokeOptions - 線の設定 {enabled, color, widthPt}
     * @property {object} cornerSmoothing - 角丸の設定 {radius, smoothing}（不要ならnull）
     * @property {number} opacity - 不透明度（%）
     * @property {number} roughenDetail - ラフ効果の詳細（0で無効）
     */

    /**
     * 図形を作れるパラメーターかどうかを判定する（幅と第2半径が数値であること）。
     * @param {ShapeParams} shapeParams - 図形生成のパラメーター
     * @returns {boolean} 作成できるならtrue
     */
    function isDrawableShapeParams(shapeParams) {
        return !!shapeParams && !isNaN(shapeParams.size) && !isNaN(shapeParams.innerRatio);
    }

    /**
     * 円のもとになるパスを作成する。
     * @param {Document} doc - 対象ドキュメント
     * @param {ShapeParams} shapeParams - 図形生成のパラメーター
     * @param {Array} viewCenter - ドキュメントウィンドウの中心座標
     * @returns {PathItem} 作成したパス
     */
    function createCircleBasePath(doc, shapeParams, viewCenter) {
        if (shapeParams.useSuperEllipse) {
            return createSuperellipsePath(doc, shapeParams.size, shapeParams.superExponent);
        }
        /* 既定の4アンカーはIllustratorの楕円、それ以外は独自のスムーズパス
           Four anchors use Illustrator's ellipse; other counts build a custom smooth path */
        var anchorCount = normalizeCircleAnchorCount(shapeParams.circleAnchorCount);
        if (anchorCount !== 4) return createCirclePathWithNAnchors(doc, shapeParams.size, anchorCount);

        var radius = shapeParams.size / 2;
        return doc.activeLayer.pathItems.ellipse(viewCenter[1] + radius, viewCenter[0] - radius, shapeParams.size, shapeParams.size);
    }

    /**
     * 正方形のもとになるパスを作成する（角丸の指定に応じて作り方を変える）。
     * @param {Document} doc - 対象ドキュメント
     * @param {ShapeParams} shapeParams - 図形生成のパラメーター
     * @param {Array} viewCenter - ドキュメントウィンドウの中心座標
     * @returns {PathItem} 作成したパス
     */
    function createSquareBasePath(doc, shapeParams, viewCenter) {
        var cornerSmoothing = shapeParams.cornerSmoothing;
        var hasCornerRadius = !!(cornerSmoothing && cornerSmoothing.radius > 0);

        if (hasCornerRadius && cornerSmoothing.smoothing > 0) {
            /* スムージングありの角丸は独自のベジェパス / Corner smoothing above zero builds a custom bezier path */
            return buildSmoothedRect(doc, viewCenter[0] - shapeParams.size / 2, viewCenter[1] + shapeParams.size / 2,
                shapeParams.size, shapeParams.size, cornerSmoothing.radius, cornerSmoothing.smoothing / 100);
        }

        /* 正方形は1辺の長さを幅として扱うので外接円の半径に換算する（既定の45°回転が前提）
           A square is sized by its edge, so convert to the circumscribed radius; assumes the default 45 degree rotation */
        var squarePath = doc.pathItems.polygon(viewCenter[0], viewCenter[1], shapeParams.size / Math.sqrt(2), 4);

        if (hasCornerRadius) {
            /* スムージング0の角丸は通常の正方形＋［角を丸くする］効果
               A zero smoothing value uses a plain square plus the Round Corners live effect */
            applyLiveEffect(squarePath, '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerSmoothing.radius + ' "/></LiveEffect>');
        }
        return squarePath;
    }

    /**
     * 辺の数と各オプションから、変形前のもとになるパスを作成する。
     * @param {Document} doc - 対象ドキュメント
     * @param {ShapeParams} shapeParams - 図形生成のパラメーター
     * @param {Array} viewCenter - ドキュメントウィンドウの中心座標
     * @returns {PathItem} 作成したパス
     */
    function createBasePath(doc, shapeParams, viewCenter) {
        var radius = shapeParams.size / 2;
        if (shapeParams.sides === 0) return createCircleBasePath(doc, shapeParams, viewCenter);
        if (shapeParams.isStar) {
            return doc.pathItems.star(viewCenter[0], viewCenter[1], radius, radius * (shapeParams.innerRatio / 100), shapeParams.sides);
        }
        if (shapeParams.sides === 4) return createSquareBasePath(doc, shapeParams, viewCenter);

        /* 正方形以外はsizeを外接円の直径として扱う（バウンディングボックスの幅とは一致しない）
           Other polygons treat the size as the circumscribed diameter, which is not the bounding box width */
        return doc.pathItems.polygon(viewCenter[0], viewCenter[1], radius, shapeParams.sides);
    }

    /**
     * 塗りと線の設定をオブジェクトに適用する。
     * @param {PathItem} shape - 対象の図形
     * @param {object} fillOptions - 塗りの設定 {enabled, color}
     * @param {object} strokeOptions - 線の設定 {enabled, color, widthPt}
     * @returns {void}
     */
    function applyFillAndStroke(shape, fillOptions, strokeOptions) {
        shape.filled = !!(fillOptions && fillOptions.enabled);
        if (shape.filled) shape.fillColor = fillOptions.color;

        shape.stroked = !!(strokeOptions && strokeOptions.enabled);
        if (shape.stroked) {
            shape.strokeColor = strokeOptions.color;
            shape.strokeWidth = strokeOptions.widthPt;
        }
    }

    /**
     * 指定したパラメーターから図形を作成し、選択状態にする。
     * @param {Document} doc - 対象ドキュメント
     * @param {ShapeParams} shapeParams - 図形生成のパラメーター
     * @returns {PathItem|GroupItem} 作成した図形
     */
    function createShape(doc, shapeParams) {
        var placement = prepareActiveLayer(doc);
        var viewCenter = [placement.centerX, placement.centerY];

        var shape = createBasePath(doc, shapeParams, viewCenter);
        applyFillAndStroke(shape, shapeParams.fillOptions, shapeParams.strokeOptions);

        /* ドキュメントウィンドウの中央にそろえる / Center the shape in the document window */
        var bounds = shape.geometricBounds;
        shape.translate(viewCenter[0] - (bounds[0] + bounds[2]) / 2, viewCenter[1] - (bounds[1] + bounds[3]) / 2);

        /* 奇数辺の多角形をルーロー（定幅図形）に変換 / Convert odd-sided polygons into constant-width shapes */
        if (shapeParams.useReuleaux && !shapeParams.isStar && shapeParams.sides > 0 && (shapeParams.sides % 2 === 1)) {
            shape = applyReuleauxToPolygon(shape, shapeParams.reuleauxAmount);
        }

        if (shapeParams.rotateEnabled && !isNaN(shapeParams.angle)) {
            shape.rotate(shapeParams.angle, true, true, true, true, Transformation.CENTER);
        }
        if (shapeParams.splitAtAnchors) {
            shape = splitPathAtAnchors(doc, shape, shapeParams.strokeOptions, shapeParams.strokeCap);
        }
        if (typeof shapeParams.opacity === "number" && shapeParams.opacity < 100) {
            shape.opacity = shapeParams.opacity;
        }
        doc.selection = [shape];
        return shape;
    }

    /**
     * ライブ効果をXMLで適用する。
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {string} effectXml - LiveEffectのXML
     * @returns {void}
     */
    function applyLiveEffect(targetItem, effectXml) {
        /* 効果を受け付けないオブジェクトでも図形の作成は続ける
           A rejected effect must not abort creating the shape itself */
        try { targetItem.applyEffect(effectXml); } catch (e) { }
    }

    /**
     * ラフ効果でアンカーポイントを追加する（変形量0なので位置は動かない）。
     * @param {PathItem|GroupItem} targetItem - 対象のオブジェクト
     * @param {number} detail - ラフ効果の詳細（0以下なら何もしない）
     * @returns {void}
     */
    function applyRoughenEffect(targetItem, detail) {
        if (!targetItem || !(detail > 0)) return;
        applyLiveEffect(targetItem, '<LiveEffect name="Adobe Roughen"><Dict data="R asiz 0 R size 0 R absoluteness 0 R dtal ' + detail + ' R roundness 0 "/></LiveEffect>');
    }

    /**
     * プレビュー中の図形を選択状態にし、プレビュー参照をクリアする。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function finalizeShape(doc) {
        if (!previewShape) return;
        doc.selection = [previewShape];
        previewShape = null;
    }

    /**
     * 指定したオブジェクトが収まるようにドキュメントウィンドウの表示倍率を合わせる。
     * ColorPaletteFromImage.jsx の fitViewToItems() をもとにしている。
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {number} fillRatio - ウィンドウに対して図形が占める割合（1.0でいっぱい）
     * @returns {void}
     */
    function fitViewToItem(doc, targetItem, fillRatio) {
        if (!targetItem) return;

        var bounds = targetItem.geometricBounds;
        var itemWidth = bounds[2] - bounds[0];
        var itemHeight = bounds[1] - bounds[3];

        var activeView = doc.activeView;
        activeView.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        if (itemWidth <= 0 || itemHeight <= 0) return;

        /* 中心をそろえたあとの表示範囲を基準に倍率を求める / Scale from the view bounds after the center has moved */
        var viewBounds = activeView.bounds;
        var scale = Math.min(
            (viewBounds[2] - viewBounds[0]) / itemWidth,
            (viewBounds[1] - viewBounds[3]) / itemHeight
        ) * fillRatio;
        activeView.zoom = clampNumber(activeView.zoom * scale, VIEW_ZOOM_RANGE, 1);
    }

    // =========================================
    // カラーの変換 / Color conversion
    // =========================================

    /**
     * Illustratorのカラーオブジェクトを ColorPicker 用の文字列に変換する。
     * @param {object} aiColor - Illustratorのカラーオブジェクト
     * @returns {string} ColorPickerが受け取る色文字列
     */
    function aiColorToPickerString(aiColor) {
        if (aiColor.typename === "RGBColor") {
            return ColorPicker.rgbToHex(aiColor.red, aiColor.green, aiColor.blue);
        } else if (aiColor.typename === "CMYKColor") {
            return "cmyk:" + Math.round(aiColor.cyan) + "," + Math.round(aiColor.magenta) + "," + Math.round(aiColor.yellow) + "," + Math.round(aiColor.black);
        } else if (aiColor.typename === "GrayColor") {
            return "cmyk:0,0,0," + Math.round(aiColor.gray);
        }
        return "000000";
    }

    /**
     * ColorPickerの戻り値をIllustratorのカラーオブジェクトに変換する。
     * @param {string} pickerString - ColorPickerが返した色文字列
     * @returns {CMYKColor|RGBColor} Illustratorのカラーオブジェクト
     */
    function pickerStringToAiColor(pickerString) {
        if (ColorPicker.isCmykString(pickerString)) {
            var cmykValues = ColorPicker.parseCmykString(pickerString);
            return makeCmykColor(cmykValues.c, cmykValues.m, cmykValues.y, cmykValues.k);
        }
        var rgbValues = ColorPicker.hexToRGB(pickerString);
        return makeRgbColor(rgbValues.r, rgbValues.g, rgbValues.b);
    }

    /**
     * Illustratorのカラーから、スウォッチ描画用のブラシを作る。
     * @param {object} graphics - ScriptUIのgraphicsオブジェクト
     * @param {object} aiColor - Illustratorのカラーオブジェクト
     * @returns {object} ブラシ（NoColorのときnull）
     */
    function aiColorToScriptUIBrush(graphics, aiColor) {
        var rgba = [1, 1, 1, 1]; /* 表示できない色は白 / Colors that cannot be shown are drawn white */
        if (aiColor.typename === "RGBColor") {
            rgba = [aiColor.red / 255, aiColor.green / 255, aiColor.blue / 255, 1];
        } else if (aiColor.typename === "CMYKColor") {
            /* 表示用にCMYKをRGBへ近似 / Approximate CMYK as RGB for display */
            rgba = [
                1 - Math.min(1, aiColor.cyan / 100 + aiColor.black / 100),
                1 - Math.min(1, aiColor.magenta / 100 + aiColor.black / 100),
                1 - Math.min(1, aiColor.yellow / 100 + aiColor.black / 100),
                1
            ];
        } else if (aiColor.typename === "GrayColor") {
            var grayValue = 1 - (aiColor.gray / 100);
            rgba = [grayValue, grayValue, grayValue, 1];
        } else if (aiColor.typename === "NoColor") {
            return null;
        }
        return graphics.newBrush(graphics.BrushType.SOLID_COLOR, rgba);
    }

    /**
     * クリックでカラーピッカーを開けるカラースウォッチを作る。
     * @param {Group} parent - 追加先のコンテナ
     * @param {object} aiColor - 初期表示に使うIllustratorのカラー
     * @returns {Group} スウォッチのグループ
     */
    function addColorSwatch(parent, aiColor) {
        var swatchGroup = parent.add("group");
        swatchGroup.preferredSize = [SWATCH_SIZE, SWATCH_SIZE];
        swatchGroup.minimumSize = [SWATCH_SIZE, SWATCH_SIZE];
        swatchGroup._aiColor = aiColor;
        swatchGroup.onDraw = function () {
            var graphics = this.graphics;
            var brush = aiColorToScriptUIBrush(graphics, this._aiColor);
            if (brush) {
                graphics.rectPath(0, 0, this.size[0], this.size[1]);
                graphics.fillPath(brush);
            }
            var borderPen = graphics.newPen(graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1);
            graphics.rectPath(0, 0, this.size[0], this.size[1]);
            graphics.strokePath(borderPen);
        };
        return swatchGroup;
    }

    /**
     * スウォッチを描き直す（色を変えたあとに呼ぶ）。
     * @param {Group} swatchGroup - 対象のスウォッチ
     * @returns {void}
     */
    function redrawSwatch(swatchGroup) {
        swatchGroup.hide();
        swatchGroup.show();
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを表示し、OKが押されたら確定用の状態を整える。
     * @param {object} rulerUnitInfo - 定規単位の情報 {label, pointsPerUnit}
     * @param {object} strokeUnitInfo - 線の単位情報 {label, pointsPerUnit}
     * @returns {boolean} OKで確定できたときtrue、キャンセルや失敗時はfalse
     */
    function showInputDialog(rulerUnitInfo, strokeUnitInfo) {
        var shapeDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        var previewManager = new PreviewManager();
        var doc = app.activeDocument;

        /* 確定に関わるダイアログの状態 / Dialog state referenced when finalizing */
        var isConfirmed = false;
        var finalParams = null;
        /* キャンセル時に戻す表示位置と倍率 / View center and zoom restored on cancel */
        var initialViewCenter = doc.activeView.centerPoint;
        var initialZoom = doc.activeView.zoom;

        /* 辺の数 / Side count */
        var sideRadios = [], customSidesInput, customSidesSlider;
        /* 回転と三角形 / Rotation and triangle */
        var rotateCheck, rotateInput, rotateUnitLabel;
        var trianglePanel, triangleRightRadio, triangleLeftRadio, triangleDownRadio;
        /* 塗りと線 / Fill and stroke */
        var fillCheck, fillSwatch, strokeCheck, strokeSwatch, strokeWidthInput, strokeWidthUnitLabel;
        var opacityInput, opacitySlider;
        /* 幅と表示 / Width and view */
        var sizeInput, fitViewCheck, fitViewPercentInput, fitViewPercentUnitLabel;
        /* スター / Star */
        var starPanel, starCheck, pentagramCheck;
        var innerRadiusLabel, innerRatioInput, innerPercentLabel, innerRatioSlider;
        /* 円 / Circle */
        var circlePanel, superEllipseCheck, superExponentSlider;
        var circleAnchorPanel, circleAnchorRadios = [];
        /* 角丸 / Corner smoothing */
        var cornerSmoothingPanel, cornerRadiusCheck, cornerRadiusInput, smoothingValueLabel, smoothingSlider;
        /* アンカーポイントの操作 / Anchor point operations */
        var roughenAnchorsCheck, roughenAnchorsInput, splitAtAnchorsCheck;
        var strokeCapLabel, capButtRadio, capRoundRadio, capProjectingRadio;
        /* オプション / Options */
        var liveShapeCheck, reuleauxCheck, reuleauxAmountInput, reuleauxAmountSlider;
        /* ボタン / Buttons */
        var btnPreview, btnCancel, btnOK;

        // -----------------------------------------
        // 現在の入力の読み取り / Reading the current input
        // -----------------------------------------

        /**
         * ラジオボタンまたは手動入力から辺の数を取得する。
         * @returns {number} 辺の数（0は円）
         */
        function getCurrentSides() {
            var selectedIndex = getSelectedRadioIndex(sideRadios);
            if (selectedIndex < 0) return DEFAULT_SIDES;
            return (selectedIndex === CUSTOM_SIDES_INDEX) ? parseInt(customSidesInput.text, 10) : SIDE_CHOICES[selectedIndex];
        }

        /**
         * 円のアンカーポイント数をラジオボタンから取得する。
         * @returns {number} アンカーポイント数
         */
        function getCircleAnchorCount() {
            var selectedIndex = getSelectedRadioIndex(circleAnchorRadios);
            return (selectedIndex < 0) ? SHAPE_DEFAULTS.circleAnchors : CIRCLE_ANCHOR_CHOICES[selectedIndex];
        }

        /**
         * スーパー楕円が実際に効く状態かどうかを判定する。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {boolean} 円かつスーパー楕円がONならtrue
         */
        function isSuperEllipseActive(sidesValue) {
            return !!(superEllipseCheck.value && sidesValue === 0);
        }

        /**
         * 選択されている線端の種類を取得する。
         * @returns {StrokeCap} 線端の種類
         */
        function getSelectedStrokeCap() {
            if (capRoundRadio.value) return StrokeCap.ROUNDENDCAP;
            if (capProjectingRadio.value) return StrokeCap.PROJECTINGENDCAP;
            return StrokeCap.BUTTENDCAP;
        }

        /**
         * スーパー楕円の指数を有効範囲に収める（小数第1位まで）。
         * @param {string|number} value - 入力値
         * @returns {number} 丸めた指数
         */
        function clampSuperExponent(value) {
            return Math.round(clampNumber(value, SHAPE_RANGES.superExponent, SHAPE_DEFAULTS.superExponent) * 10) / 10;
        }

        /* 整数に丸めて有効範囲に収める関数 / Clamp functions that round to integers */
        var clampReuleauxAmount = makeRangeClamp("reuleauxAmount");
        var clampCustomSides = makeRangeClamp("customSides");
        var clampInnerRatio = makeRangeClamp("innerRatio");
        var clampOpacity = makeRangeClamp("opacity");
        var clampSmoothing = makeRangeClamp("smoothing");
        var clampFitViewPercent = makeRangeClamp("fitViewPercent");

        // -----------------------------------------
        // パネルの組み立て / Panel construction
        // -----------------------------------------

        /**
         * ［辺の数］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildSidesPanel(parent) {
            var sidesPanel = addPanel(parent, getLabel(LABELS.panel.sides));

            sideRadios[0] = sidesPanel.add("radiobutton", undefined, getLabel(LABELS.radio.circleWithZero));
            sideRadios[0].helpTip = getLabel(LABELS.tooltip.circle);
            for (var i = 1; i < SIDE_CHOICES.length; i++) {
                var sideLabel = (SIDE_CHOICES[i] === 4) ? getLabel(LABELS.radio.squareWithFour) : String(SIDE_CHOICES[i]);
                sideRadios[i] = sidesPanel.add("radiobutton", undefined, sideLabel);
                sideRadios[i].helpTip = getLabel(LABELS.tooltip.sideChoice);
            }

            /* ［それ以外］はラベルを持たず、ラジオ・入力欄・スライダーを1行に並べる
               The custom side count has no label; its radio, field and slider share one row */
            var customSidesRow = addControlRow(sidesPanel);
            sideRadios[CUSTOM_SIDES_INDEX] = customSidesRow.add("radiobutton", undefined, "");
            sideRadios[CUSTOM_SIDES_INDEX].helpTip = getLabel(LABELS.tooltip.customSides);
            customSidesInput = addNumberField(customSidesRow, SHAPE_DEFAULTS.customSides, 3, { min: SHAPE_RANGES.customSides[0], max: SHAPE_RANGES.customSides[1], integer: true });
            customSidesSlider = addSlider(customSidesRow, SHAPE_DEFAULTS.customSides, SHAPE_RANGES.customSides, INLINE_SLIDER_WIDTH);
            customSidesInput.helpTip = getLabel(LABELS.tooltip.customSides);
            customSidesSlider.helpTip = getLabel(LABELS.tooltip.customSides);

            selectRadio(sideRadios, findChoiceIndex(SIDE_CHOICES, DEFAULT_SIDES));
            setCustomSidesEnabled(false);
        }

        /**
         * ［回転］パネルと、その中の［三角形］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildRotatePanel(parent) {
            var rotatePanel = addPanel(parent, getLabel(LABELS.panel.rotation));

            var rotateRow = addControlRow(rotatePanel);
            rotateCheck = rotateRow.add("checkbox", undefined, "");
            rotateInput = addNumberField(rotateRow, SHAPE_DEFAULTS.rotation, 4);
            rotateUnitLabel = rotateRow.add("statictext", undefined, "°");
            rotateCheck.helpTip = getLabel(LABELS.tooltip.rotate);
            /* 手動入力はチェック時だけ有効 / Manual entry is enabled only while checked */
            setRotateInputEnabled(rotateCheck.value);

            /* 三角形パネルは回転パネルの中に置く / The triangle panel lives inside the rotation panel */
            trianglePanel = addPanel(rotatePanel, getLabel(LABELS.panel.triangle));
            var triangleRow = addControlRow(trianglePanel, WIDE_ROW_SPACING);
            triangleRightRadio = triangleRow.add("radiobutton", undefined, getLabel(LABELS.radio.triangleRight));
            triangleLeftRadio = triangleRow.add("radiobutton", undefined, getLabel(LABELS.radio.triangleLeft));
            triangleDownRadio = triangleRow.add("radiobutton", undefined, getLabel(LABELS.radio.triangleDown));
            triangleRightRadio.helpTip = getLabel(LABELS.tooltip.triangleRight);
            triangleLeftRadio.helpTip = getLabel(LABELS.tooltip.triangleLeft);
            triangleDownRadio.helpTip = getLabel(LABELS.tooltip.triangleDown);
            triangleRightRadio.value = true;
        }

        /**
         * ［塗りと線］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildFillStrokePanel(parent) {
            var fillStrokePanel = addPanel(parent, getLabel(LABELS.panel.fillAndStroke));
            /* スウォッチの位置をそろえるため、［塗り］［線］の幅を固定する / Fixed width so the swatches line up */
            var checkboxWidth = (uiLang === 'ja') ? 66 : 70;

            var fillRow = addControlRow(fillStrokePanel);
            fillCheck = fillRow.add("checkbox", undefined, labelText(LABELS.checkbox.fill));
            fillCheck.preferredSize.width = checkboxWidth;
            fillCheck.value = true;
            fillSwatch = addColorSwatch(fillRow, doc.defaultFillColor);
            fillSwatch.helpTip = getLabel(LABELS.tooltip.colorSwatch);

            var strokeRow = addControlRow(fillStrokePanel);
            strokeCheck = strokeRow.add("checkbox", undefined, labelText(LABELS.checkbox.stroke));
            strokeCheck.preferredSize.width = checkboxWidth;
            strokeCheck.value = false;
            strokeSwatch = addColorSwatch(strokeRow, doc.defaultStrokeColor);
            strokeSwatch.helpTip = getLabel(LABELS.tooltip.colorSwatch);

            /* 線幅は［線］と同じ行に置く（ラベルなし） / The stroke width sits on the stroke row, without a label */
            strokeWidthInput = addNumberField(strokeRow, SHAPE_DEFAULTS.strokeWidth, 4, { min: 0 });
            strokeWidthInput.helpTip = getLabel(LABELS.tooltip.strokeWidth);
            strokeWidthUnitLabel = strokeRow.add("statictext", undefined, strokeUnitInfo.label);

            var opacityRow = addControlRow(fillStrokePanel);
            opacityRow.add("statictext", undefined, labelText(LABELS.fieldLabel.opacity));
            opacityInput = addNumberField(opacityRow, SHAPE_DEFAULTS.opacity, 4, { min: SHAPE_RANGES.opacity[0], max: SHAPE_RANGES.opacity[1], integer: true });
            opacityRow.add("statictext", undefined, "%");
            opacitySlider = addSlider(fillStrokePanel, SHAPE_DEFAULTS.opacity, SHAPE_RANGES.opacity, SHORT_SLIDER_WIDTH);
            opacitySlider.helpTip = getLabel(LABELS.tooltip.opacity);
        }

        /**
         * ［幅］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildWidthPanel(parent) {
            var widthPanel = addPanel(parent, getLabel(LABELS.panel.width));
            var widthRow = addControlRow(widthPanel);
            sizeInput = addNumberField(widthRow, SHAPE_DEFAULTS.size, 5, { min: 0 });
            sizeInput.helpTip = getLabel(LABELS.tooltip.width);
            widthRow.add("statictext", undefined, rulerUnitInfo.label);
        }

        /**
         * ［画面にフィット］の行を組み立てる（幅パネルの下に置く）。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildFitViewRow(parent) {
            var fitViewRow = addControlRow(parent);
            fitViewRow.margins = [20, 0, 0, 0];
            fitViewCheck = fitViewRow.add("checkbox", undefined, getLabel(LABELS.checkbox.fitView));
            fitViewCheck.value = false;
            fitViewCheck.helpTip = getLabel(LABELS.tooltip.fitView);
            fitViewPercentInput = addNumberField(fitViewRow, SHAPE_DEFAULTS.fitViewPercent, 3, { min: SHAPE_RANGES.fitViewPercent[0], max: SHAPE_RANGES.fitViewPercent[1], integer: true });
            fitViewPercentInput.helpTip = getLabel(LABELS.tooltip.fitViewPercent);
            fitViewPercentUnitLabel = fitViewRow.add("statictext", undefined, "%");
        }

        /**
         * ［スター］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildStarPanel(parent) {
            starPanel = addPanel(parent, getLabel(LABELS.panel.star));

            var starRow = addControlRow(starPanel, WIDE_ROW_SPACING);
            starCheck = starRow.add("checkbox", undefined, getLabel(LABELS.checkbox.star));
            pentagramCheck = starRow.add("checkbox", undefined, getLabel(LABELS.checkbox.pentagram));
            pentagramCheck.value = false;
            starCheck.helpTip = getLabel(LABELS.tooltip.star);
            pentagramCheck.helpTip = getLabel(LABELS.tooltip.pentagram);

            var innerRadiusRow = addControlRow(starPanel);
            innerRadiusLabel = innerRadiusRow.add("statictext", undefined, labelText(LABELS.fieldLabel.innerRadius));
            innerRatioInput = addNumberField(innerRadiusRow, SHAPE_DEFAULTS.innerRatio, 4, { min: SHAPE_RANGES.innerRatio[0], max: SHAPE_RANGES.innerRatio[1], integer: true });
            innerPercentLabel = innerRadiusRow.add("statictext", undefined, "%");

            innerRatioSlider = addSlider(starPanel, SHAPE_DEFAULTS.innerRatio, SHAPE_RANGES.innerRatio, SLIDER_WIDTH);
            innerRatioInput.helpTip = getLabel(LABELS.tooltip.innerRadius);
            innerRatioSlider.helpTip = getLabel(LABELS.tooltip.innerRadius);
        }

        /**
         * ［円］パネルと、その中の［アンカーポイント］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildCirclePanel(parent) {
            circlePanel = addPanel(parent, getLabel(LABELS.panel.circle));

            /* チェックボックスと指数のスライダーを1行に並べる / The checkbox and the exponent slider share one row */
            var superEllipseRow = addControlRow(circlePanel);
            superEllipseCheck = superEllipseRow.add("checkbox", undefined, getLabel(LABELS.checkbox.superEllipse));
            superEllipseCheck.value = false;
            superEllipseCheck.helpTip = getLabel(LABELS.tooltip.superEllipse);

            superExponentSlider = addSlider(superEllipseRow, SHAPE_DEFAULTS.superExponent, SHAPE_RANGES.superExponent, INLINE_SLIDER_WIDTH);
            superExponentSlider.helpTip = getLabel(LABELS.tooltip.superExponent);

            circleAnchorPanel = addPanel(circlePanel, getLabel(LABELS.panel.anchorCount));
            var circleAnchorRow = addControlRow(circleAnchorPanel, WIDE_ROW_SPACING);
            for (var i = 0; i < CIRCLE_ANCHOR_CHOICES.length; i++) {
                circleAnchorRadios[i] = circleAnchorRow.add("radiobutton", undefined, String(CIRCLE_ANCHOR_CHOICES[i]));
            }
            selectDefaultCircleAnchors();
        }

        /**
         * ［角丸］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildCornerSmoothingPanel(parent) {
            cornerSmoothingPanel = addPanel(parent, getLabel(LABELS.panel.cornerSmoothing));

            var cornerRadiusRow = addControlRow(cornerSmoothingPanel);
            cornerRadiusCheck = cornerRadiusRow.add("checkbox", undefined, labelText(LABELS.checkbox.cornerRadius));
            cornerRadiusCheck.value = false;
            cornerRadiusCheck.helpTip = getLabel(LABELS.tooltip.cornerRadius);
            cornerRadiusInput = addNumberField(cornerRadiusRow, SHAPE_DEFAULTS.cornerRadius, 5, { min: 0 });
            cornerRadiusRow.add("statictext", undefined, rulerUnitInfo.label);

            var smoothingLabelRow = addControlRow(cornerSmoothingPanel);
            smoothingLabelRow.add("statictext", undefined, labelText(LABELS.fieldLabel.smoothing));
            smoothingValueLabel = smoothingLabelRow.add("statictext", undefined, SHAPE_DEFAULTS.smoothing + "%");
            smoothingValueLabel.characters = 4;

            smoothingSlider = addSlider(cornerSmoothingPanel, SHAPE_DEFAULTS.smoothing, SHAPE_RANGES.smoothing, SLIDER_WIDTH);
            smoothingSlider.helpTip = getLabel(LABELS.tooltip.smoothing);
        }

        /**
         * ［アンカーポイントの操作］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildAnchorOpsPanel(parent) {
            var anchorOpsPanel = addPanel(parent, getLabel(LABELS.panel.anchorOps));

            var roughenAnchorsRow = addControlRow(anchorOpsPanel, TIGHT_ROW_SPACING);
            roughenAnchorsCheck = roughenAnchorsRow.add("checkbox", undefined, labelText(LABELS.checkbox.roughenAnchors));
            roughenAnchorsCheck.value = false;
            roughenAnchorsCheck.helpTip = getLabel(LABELS.tooltip.roughenAnchors);
            roughenAnchorsInput = addNumberField(roughenAnchorsRow, SHAPE_DEFAULTS.roughenDetail, 3, { min: 0, integer: true });
            setControlsEnabled([roughenAnchorsInput], false);

            splitAtAnchorsCheck = anchorOpsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.splitAtAnchors));
            splitAtAnchorsCheck.value = false;
            splitAtAnchorsCheck.helpTip = getLabel(LABELS.tooltip.splitAtAnchors);

            var strokeCapRow = addControlRow(anchorOpsPanel, TIGHT_ROW_SPACING);
            strokeCapLabel = strokeCapRow.add("statictext", undefined, labelText(LABELS.fieldLabel.strokeCap));
            capButtRadio = strokeCapRow.add("radiobutton", undefined, getLabel(LABELS.radio.capButt));
            capRoundRadio = strokeCapRow.add("radiobutton", undefined, getLabel(LABELS.radio.capRound));
            capProjectingRadio = strokeCapRow.add("radiobutton", undefined, getLabel(LABELS.radio.capProjecting));
            capButtRadio.value = true;
        }

        /**
         * ［オプション］パネルを組み立てる。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {void}
         */
        function buildOptionsPanel(parent) {
            var optionPanel = addPanel(parent, getLabel(LABELS.panel.option));

            liveShapeCheck = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.liveShape));
            liveShapeCheck.value = true;
            liveShapeCheck.helpTip = getLabel(LABELS.tooltip.liveShape);

            reuleauxCheck = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.reuleaux));
            reuleauxCheck.value = false;
            reuleauxCheck.helpTip = getLabel(LABELS.tooltip.reuleaux);

            var reuleauxAmountRow = addControlRow(optionPanel);
            reuleauxAmountInput = addNumberField(reuleauxAmountRow, SHAPE_DEFAULTS.reuleauxAmount, 4, { min: SHAPE_RANGES.reuleauxAmount[0], max: SHAPE_RANGES.reuleauxAmount[1], integer: true });
            reuleauxAmountSlider = addSlider(reuleauxAmountRow, SHAPE_DEFAULTS.reuleauxAmount, SHAPE_RANGES.reuleauxAmount, SHORT_SLIDER_WIDTH);
            reuleauxAmountInput.helpTip = getLabel(LABELS.tooltip.reuleauxAmount);
            reuleauxAmountSlider.helpTip = getLabel(LABELS.tooltip.reuleauxAmount);
        }

        // -----------------------------------------
        // 有効・無効の制御 / Enabled state
        // -----------------------------------------

        /**
         * ［それ以外］の入力欄とスライダーの有効・無効をまとめて切り替える。
         * @param {boolean} isEnabled - 有効にするかどうか
         * @returns {void}
         */
        function setCustomSidesEnabled(isEnabled) {
            setControlsEnabled([customSidesInput, customSidesSlider], isEnabled);
        }

        /**
         * 回転角の入力欄と単位の有効・無効を切り替える。
         * @param {boolean} isEnabled - 有効にするかどうか
         * @returns {void}
         */
        function setRotateInputEnabled(isEnabled) {
            setControlsEnabled([rotateInput, rotateUnitLabel], isEnabled);
        }

        /**
         * 円のアンカーポイント数を既定値（4）に戻す。
         * @returns {void}
         */
        function selectDefaultCircleAnchors() {
            selectRadio(circleAnchorRadios, findChoiceIndex(CIRCLE_ANCHOR_CHOICES, SHAPE_DEFAULTS.circleAnchors));
        }

        /**
         * ［画面にフィット］の割合の入力欄を、チェックの状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateFitViewPercentEnabled() {
            setControlsEnabled([fitViewPercentInput, fitViewPercentUnitLabel], fitViewCheck.value);
        }

        /**
         * 線幅の入力欄を［線］チェックの状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateStrokeWidthEnabled() {
            setControlsEnabled([strokeWidthInput, strokeWidthUnitLabel], strokeCheck.value);
        }

        /**
         * 円パネルの有効・無効を辺の数に応じて切り替える。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateCirclePanelEnabled(sidesValue) {
            var isEnabled = (sidesValue === 0);
            circlePanel.enabled = isEnabled;
            circleAnchorPanel.enabled = isEnabled;
            if (isEnabled) return;

            /* 円以外に切り替えたら円用の設定を既定へ戻す / Reset the circle options when leaving the circle */
            superEllipseCheck.value = false;
            selectDefaultCircleAnchors();
            setControlsEnabled(circleAnchorRadios, true);
        }

        /**
         * スターパネルの有効・無効を辺の数に応じて切り替える。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateStarPanelEnabled(sidesValue) {
            /* 円（0）にはスターの設定を適用しない / Star options do not apply to a circle */
            var isEnabled = (sidesValue !== 0);
            starPanel.enabled = isEnabled;
            if (!isEnabled) clearStarOptions();
        }

        /**
         * スターと五芒星をOFFにし、五芒星を選べなくする。
         * @returns {void}
         */
        function clearStarOptions() {
            starCheck.value = false;
            pentagramCheck.value = false;
            pentagramCheck.enabled = false;
        }

        /**
         * 第2半径の入力群を［スター］チェックの状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateInnerRadiusEnabled() {
            setControlsEnabled([innerRadiusLabel, innerRatioInput, innerPercentLabel, innerRatioSlider], !!starCheck.value);
        }

        /**
         * スーパー楕円の指数入力と、円のアンカーポイントパネルの有効状態を更新する。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateSuperEllipseControlsEnabled(sidesValue) {
            var isActive = isSuperEllipseActive(sidesValue);
            superExponentSlider.enabled = isActive;
            /* スーパー楕円がONのときはアンカーポイント数を選べない
               The anchor count cannot be chosen while the superellipse is on */
            circleAnchorPanel.enabled = !isActive;
            setControlsEnabled(circleAnchorRadios, !isActive);
        }

        /**
         * ルーローの度合いの入力群を有効・無効にする。
         * @returns {void}
         */
        function updateReuleauxAmountEnabled() {
            setControlsEnabled([reuleauxAmountInput, reuleauxAmountSlider], reuleauxCheck.enabled && reuleauxCheck.value);
        }

        /**
         * ルーローのチェックボックスを、辺の数とスターの状態に応じて有効・無効にする。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateReuleauxAvailability(sidesValue) {
            /* スターがONのときは奇数判定を行わない（スター側の制御を優先）
               While the star is on, the odd-side rule is skipped and the star logic wins */
            if (!starCheck.value) {
                /* ルーローは奇数辺（3、5、7…）のみ。円（0）と偶数辺は対象外
                   Reuleaux applies to odd side counts only, never to a circle or an even count */
                var isEnabled = (typeof sidesValue === "number") && (sidesValue > 0) && (sidesValue % 2 === 1);
                reuleauxCheck.enabled = isEnabled;
                if (!isEnabled) reuleauxCheck.value = false;
            }
            updateReuleauxAmountEnabled();
        }

        /**
         * 角丸の入力群を［半径］チェックの状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateCornerRadiusInputEnabled() {
            setControlsEnabled([cornerRadiusInput, smoothingSlider, smoothingValueLabel], cornerRadiusCheck.value);
        }

        /**
         * 角丸パネルの有効・無効を辺の数に応じて切り替える。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateCornerSmoothingEnabled(sidesValue) {
            var isEnabled = (sidesValue === 4);
            cornerSmoothingPanel.enabled = isEnabled;

            /* 正方形以外ではパネルを無効にするだけで入力値は破棄しない。
               角丸は sides === 4 のときしか参照されないため、値を残しても影響はない
               Outside a square the panel is only disabled; the values are kept, and they are
               read only when sides === 4, so keeping them is harmless */
            if (isEnabled && !(parseFloat(cornerRadiusInput.text) > 0)) {
                /* 有効化したときは幅に対する比率で既定値を入れる / Restore the ratio-based default when enabled */
                var currentWidth = parseFloat(sizeInput.text);
                if (!isNaN(currentWidth) && currentWidth > 0) {
                    cornerRadiusInput.text = String(Math.round(currentWidth * SHAPE_DEFAULTS.cornerRadiusRatio * 10) / 10);
                }
            }
        }

        /**
         * 線端の選択肢を［アンカーポイントで分割］の状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateStrokeCapEnabled() {
            setControlsEnabled([strokeCapLabel, capButtRadio, capRoundRadio, capProjectingRadio], splitAtAnchorsCheck.value);
        }

        /**
         * 現在のUIの状態からライブシェイプ化の可否を再計算する。
         * 分割・スーパー楕円・4以外のアンカー数・ルーロー・ラフ効果・角丸のいずれかが有効なら使えない。
         * @returns {void}
         */
        function refreshLiveShapeAvailability() {
            var sidesValue = getCurrentSides();
            var isSuperEllipse = isSuperEllipseActive(sidesValue);
            /* ライブシェイプ化できるのは円のアンカーが4のときだけ（スーパー楕円を除く）
               A live shape is possible only with four circle anchors and no superellipse */
            var isCustomCircleAnchors = (sidesValue === 0 && !isSuperEllipse && getCircleAnchorCount() !== 4);
            var isCornerSmoothing = (sidesValue === 4 && cornerRadiusCheck.value && parseFloat(cornerRadiusInput.text) > 0);

            var isBlocked = splitAtAnchorsCheck.value || isSuperEllipse || isCustomCircleAnchors ||
                reuleauxCheck.value || roughenAnchorsCheck.value || isCornerSmoothing;
            if (isBlocked) liveShapeCheck.value = false;
            liveShapeCheck.enabled = !isBlocked;
        }

        /**
         * 辺の数に依存する各パネルの有効状態をまとめて更新する。
         * @returns {number} 現在の辺の数
         */
        function refreshPanelStates() {
            var sidesValue = getCurrentSides();
            trianglePanel.enabled = (sidesValue === 3);
            updateCirclePanelEnabled(sidesValue);
            updateCornerSmoothingEnabled(sidesValue);
            updateSuperEllipseControlsEnabled(sidesValue);
            updateStarPanelEnabled(sidesValue);
            updateReuleauxAvailability(sidesValue);
            updateInnerRadiusEnabled();
            refreshLiveShapeAvailability();
            return sidesValue;
        }

        // -----------------------------------------
        // 入力どうしの整合 / Keeping the inputs consistent
        // -----------------------------------------

        /**
         * 回転のチェックを切り替え、角度の入力欄の有効状態もそろえる。
         * @param {boolean} isChecked - 回転をONにするかどうか
         * @returns {void}
         */
        function setRotateChecked(isChecked) {
            rotateCheck.value = isChecked;
            setRotateInputEnabled(isChecked);
        }

        /**
         * 回転がOFFのときに使う自動角度を、回転の入力欄に反映する。
         * 辺の数が変わったときは回転がONでも呼び出す。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {number} 反映した角度（対象外の辺の数ならNaN）
         */
        function applyAutoRotationForSides(sidesValue) {
            var angle = NaN;
            if (sidesValue === 0) angle = 45;
            else if (sidesValue >= 3) angle = 360 / (sidesValue * 2);
            if (!isNaN(angle)) rotateInput.text = formatAngle(angle);
            return angle;
        }

        /**
         * 回転をONにしたときの既定角度を決めて入力欄に反映する。
         * 円（辺の数0）はアンカー数ごとに 2→90、3→180、4→45、5→180、6→30 を使う。
         * @returns {void}
         */
        function applyDefaultRotationWhenEnablingRotate() {
            var sidesValue = getCurrentSides();
            if (sidesValue === 0 && !isSuperEllipseActive(sidesValue)) {
                var angle = CIRCLE_ROTATION_DEFAULTS[getCircleAnchorCount()];
                rotateInput.text = formatAngle((typeof angle === "number") ? angle : 45);
                return;
            }
            applyAutoRotationForSides(sidesValue);
            /* 三角形のときは［下］（60°）を既定にする / A triangle defaults to "down" (60 degrees) */
            if (sidesValue === 3) triangleDownRadio.value = true;
        }

        /**
         * スターと五芒星、およびルーローとの排他関係を整える。
         * @returns {void}
         */
        function validateStarAndPentagram() {
            /* スターパネルが無効（円）ならスター関連を強制的にOFF / Force the star options off while the panel is disabled */
            if (!starPanel.enabled) {
                starCheck.enabled = false;
                clearStarOptions();
                return;
            }

            starCheck.enabled = true;
            if (!starCheck.value) pentagramCheck.value = false;
            pentagramCheck.enabled = starCheck.value;

            /* ルーローはスターと併用できない / Reuleaux is not compatible with a star */
            if (starCheck.value) {
                reuleauxCheck.value = false;
                reuleauxCheck.enabled = false;
                updateReuleauxAmountEnabled();
            }

            if (pentagramCheck.value) {
                selectRadio(sideRadios, findChoiceIndex(SIDE_CHOICES, 5)); /* 五芒星は5辺 / A pentagram has five sides */
                setCustomSidesEnabled(false);
                applyAutoRotationForSides(5);
                setRotateChecked(false);
            }
            /* ルーローと第2半径の復帰は、続くrefreshPanelStates()が行う
               refreshPanelStates(), which runs next, restores Reuleaux and the inner radius */
        }

        /**
         * 丸めた値を入力欄とスライダーの両方に反映する。
         * @param {EditText} inputField - 入力欄
         * @param {Slider} slider - スライダー
         * @param {number} value - 反映する値（範囲内に収めたもの）
         * @returns {void}
         */
        function syncFieldAndSlider(inputField, slider, value) {
            inputField.text = String(value);
            slider.value = value;
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * 回転角を決め、自動で決まる角度は入力欄にも反映する。
         * @param {number} sides - 現在の辺の数
         * @returns {number} 回転角（度）
         */
        function resolveRotationAngle(sides) {
            /* 回転がOFFのときだけ自動角度を使う（ONなら入力値を尊重）
               The auto angle is used only while the rotation is off; manual mode keeps the entered angle */
            var angle = parseFloat(rotateInput.text);
            if (!rotateCheck.value) {
                var autoAngle = applyAutoRotationForSides(sides);
                if (!isNaN(autoAngle)) angle = autoAngle;
            }

            /* 三角形は向きごとの回転角で上書きする / A triangle overrides the angle with its direction */
            if (sides === 3 && trianglePanel.enabled) {
                if (triangleRightRadio.value) angle = TRIANGLE_ANGLES.right;
                else if (triangleLeftRadio.value) angle = TRIANGLE_ANGLES.left;
                else if (triangleDownRadio.value) angle = TRIANGLE_ANGLES.down;
                rotateInput.text = formatAngle(angle);
            }
            return angle;
        }

        /**
         * スターの第2半径を決める。五芒星のときは黄金比の値を入力欄にも反映する。
         * @param {number} sides - 現在の辺の数
         * @returns {number} 第2半径（%）
         */
        function resolveInnerRatio(sides) {
            if (starCheck.value && pentagramCheck.value && sides === 5) {
                var pentagramRatio = (3 - Math.sqrt(5)) / 2 * 100;
                innerRatioInput.text = pentagramRatio.toFixed(2);
                return pentagramRatio;
            }
            return parseFloat(innerRatioInput.text);
        }

        /**
         * プレビューに適用するラフ効果の詳細を決め、メニューコマンドで代替するかを記録する。
         * @param {number} sides - 現在の辺の数
         * @returns {number} ラフ効果の詳細（0で無効）
         */
        function resolveRoughenDetail(sides) {
            /* 詳細が1の多角形はメニューコマンドでアンカーを追加する。この経路はプレビューできない
               A detail of 1 on a polygon adds anchors through a menu command, which cannot be previewed */
            roughenAnchorsUseMenuFallback = (sides !== 0 && Math.round(Number(roughenAnchorsInput.text)) === 1);
            if (!roughenAnchorsCheck.value || roughenAnchorsUseMenuFallback) return 0;
            var roughenDetail = parseFloat(roughenAnchorsInput.text);
            return (isNaN(roughenDetail) || roughenDetail < 0) ? 0 : roughenDetail;
        }

        /**
         * 現在のUIから図形生成のパラメーターを組み立てる（プレビューと確定の両方で使う）。
         * @returns {ShapeParams} createShapeに渡すパラメーター一式
         */
        function getCurrentShapeParams() {
            validateStarAndPentagram();
            var sides = refreshPanelStates();

            var useSuperEllipse = isSuperEllipseActive(sides);
            /* スーパー楕円は回転を強制的にOFFにする / The superellipse forces the rotation off */
            if (useSuperEllipse) setRotateChecked(false);

            var angle = resolveRotationAngle(sides);
            var innerRatio = resolveInnerRatio(sides);
            var roughenDetail = resolveRoughenDetail(sides);

            /* 円のアンカーポイント数は、円かつスーパー楕円OFFのときだけ有効
               The circle anchor count only matters for a circle without the superellipse */
            var circleAnchorCount = (sides === 0 && !useSuperEllipse) ? getCircleAnchorCount() : SHAPE_DEFAULTS.circleAnchors;

            return {
                size: parseFloat(sizeInput.text) * rulerUnitInfo.pointsPerUnit,
                sides: sides,
                isStar: starCheck.value,
                innerRatio: innerRatio,
                rotateEnabled: rotateCheck.value,
                angle: angle,
                splitAtAnchors: splitAtAnchorsCheck.value,
                strokeCap: splitAtAnchorsCheck.value ? getSelectedStrokeCap() : null,
                useSuperEllipse: useSuperEllipse,
                superExponent: clampSuperExponent(superExponentSlider.value),
                circleAnchorCount: circleAnchorCount,
                useReuleaux: reuleauxCheck.value,
                reuleauxAmount: clampReuleauxAmount(reuleauxAmountInput.text) / 100,
                fillOptions: {
                    enabled: fillCheck.value,
                    color: fillSwatch._aiColor
                },
                strokeOptions: {
                    enabled: strokeCheck.value,
                    color: strokeSwatch._aiColor,
                    widthPt: parseFloat(strokeWidthInput.text) * strokeUnitInfo.pointsPerUnit
                },
                cornerSmoothing: (sides === 4 && cornerRadiusCheck.value) ? {
                    radius: parseFloat(cornerRadiusInput.text) * rulerUnitInfo.pointsPerUnit,
                    smoothing: Math.round(smoothingSlider.value)
                } : null,
                opacity: clampOpacity(opacityInput.text),
                roughenDetail: roughenDetail
            };
        }

        /**
         * ［画面にフィット］がONのとき、プレビューが収まるよう表示倍率を合わせる。
         * 呼ぶのは［幅］を変えたときとダイアログを開いたときだけ。ほかの設定でも呼ぶと
         * 1キーごとに倍率が動いて落ち着かないため。
         * 倍率の変更はUndo履歴に残らないので、キャンセル時はdiscardPreview()で戻す。
         * @returns {void}
         */
        function applyFitView() {
            if (!fitViewCheck.value || !previewShape) return;
            fitViewToItem(doc, previewShape, clampFitViewPercent(fitViewPercentInput.text) / 100);
            app.redraw();
        }

        /**
         * プレビューをすべて巻き戻し、プレビュー参照を外す。
         * @returns {void}
         */
        function resetPreview() {
            previewManager.rollback();
            previewShape = null;
        }

        /**
         * 入力内容からプレビューを描き直す（Undo履歴を汚さない）。
         * @returns {void}
         */
        function updatePreview() {
            /* 新しいプレビューの前に必ず前回分を巻き戻す / Always roll back the previous preview first */
            resetPreview();

            var shapeParams = getCurrentShapeParams();
            if (!isDrawableShapeParams(shapeParams)) return;

            previewManager.addStep(function () {
                previewShape = createShape(app.activeDocument, shapeParams);
                /* ラフ効果も同じステップ内で適用してプレビューに反映する
                   The roughen effect runs inside the same step so the preview shows it */
                applyRoughenEffect(previewShape, shapeParams.roughenDetail);
            });
        }

        /**
         * ［幅］を変えたときの処理。プレビューを更新してから画面にフィットさせる。
         * @returns {void}
         */
        function onSizeChange() {
            updatePreview();
            applyFitView();
        }

        // -----------------------------------------
        // セッション状態 / Session state
        // -----------------------------------------

        /**
         * 辺の数・幅・［画面にフィット］を復元する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function restoreSidesAndSize(sessionState) {
            var sideIndex = sessionState.selectedSideIndex;
            if (typeof sideIndex === "number" && sideIndex >= 0 && sideIndex < sideRadios.length) {
                selectRadio(sideRadios, sideIndex);
                var isCustomSides = (sideIndex === CUSTOM_SIDES_INDEX);
                if (isCustomSides && typeof sessionState.customSidesText === "string") {
                    customSidesInput.text = sessionState.customSidesText;
                    /* スライダーのつまみも復元した辺の数に合わせる / Move the slider thumb to the restored side count */
                    customSidesSlider.value = clampCustomSides(customSidesInput.text);
                }
                setCustomSidesEnabled(isCustomSides);
            }
            if (typeof sessionState.sizeText === "string") sizeInput.text = sessionState.sizeText;
            if (typeof sessionState.fitViewCheck === "boolean") fitViewCheck.value = sessionState.fitViewCheck;
            if (typeof sessionState.fitViewPercentText === "string") fitViewPercentInput.text = sessionState.fitViewPercentText;
            updateFitViewPercentEnabled();
        }

        /**
         * 塗りと線・不透明度を復元する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function restoreAppearance(sessionState) {
            if (typeof sessionState.fillCheck === "boolean") fillCheck.value = sessionState.fillCheck;
            if (typeof sessionState.strokeCheck === "boolean") strokeCheck.value = sessionState.strokeCheck;
            if (typeof sessionState.strokeWidthText === "string") strokeWidthInput.text = sessionState.strokeWidthText;
            if (sessionState.fillColorString) fillSwatch._aiColor = pickerStringToAiColor(sessionState.fillColorString);
            if (sessionState.strokeColorString) strokeSwatch._aiColor = pickerStringToAiColor(sessionState.strokeColorString);
            updateStrokeWidthEnabled();

            if (typeof sessionState.opacityText === "string") {
                opacityInput.text = sessionState.opacityText;
                opacitySlider.value = clampOpacity(sessionState.opacityText);
            }
        }

        /**
         * 角丸を復元する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function restoreCornerSmoothing(sessionState) {
            if (typeof sessionState.cornerRadiusCheck === "boolean") cornerRadiusCheck.value = sessionState.cornerRadiusCheck;
            if (typeof sessionState.cornerRadiusText === "string") cornerRadiusInput.text = sessionState.cornerRadiusText;
            if (typeof sessionState.smoothingValue === "number") {
                var restoredSmoothing = clampSmoothing(sessionState.smoothingValue);
                smoothingSlider.value = restoredSmoothing;
                smoothingValueLabel.text = restoredSmoothing + "%";
            }
            updateCornerRadiusInputEnabled();
        }

        /**
         * 回転と三角形の向きを復元する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function restoreRotation(sessionState) {
            if (typeof sessionState.rotateText === "string") rotateInput.text = sessionState.rotateText;
            setRotateChecked((typeof sessionState.rotateCheck === "boolean") ? sessionState.rotateCheck : rotateCheck.value);
            if (sessionState.triangleDir === "left") triangleLeftRadio.value = true;
            else if (sessionState.triangleDir === "down") triangleDownRadio.value = true;
            else if (sessionState.triangleDir === "right") triangleRightRadio.value = true;
        }

        /**
         * スター・五芒星・円の設定を復元する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function restoreStarAndCircle(sessionState) {
            if (typeof sessionState.starCheck === "boolean") starCheck.value = sessionState.starCheck;
            if (typeof sessionState.pentagramCheck === "boolean") pentagramCheck.value = sessionState.pentagramCheck;
            if (!starCheck.value) pentagramCheck.value = false;
            if (typeof sessionState.innerRatioText === "string") innerRatioInput.text = sessionState.innerRatioText;
            innerRatioSlider.value = clampInnerRatio(innerRatioInput.text);

            if (typeof sessionState.superEllipseCheck === "boolean") superEllipseCheck.value = sessionState.superEllipseCheck;
            if (typeof sessionState.superExponentValue === "number") superExponentSlider.value = clampSuperExponent(sessionState.superExponentValue);
            if (typeof sessionState.circleAnchorsValue === "number") {
                var anchorIndex = findChoiceIndex(CIRCLE_ANCHOR_CHOICES, sessionState.circleAnchorsValue);
                if (anchorIndex < 0) selectDefaultCircleAnchors();
                else selectRadio(circleAnchorRadios, anchorIndex);
            }
        }

        /**
         * アンカーポイントの操作とオプションを復元する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function restoreAnchorOpsAndOptions(sessionState) {
            /* ［アンカーポイントで分割］は毎回OFFで開く（復元しない）
               "Split at anchor points" always opens off and is never restored */
            splitAtAnchorsCheck.value = false;
            if (sessionState.strokeCapValue === "round") capRoundRadio.value = true;
            else if (sessionState.strokeCapValue === "projecting") capProjectingRadio.value = true;
            else capButtRadio.value = true;
            updateStrokeCapEnabled();

            if (typeof sessionState.reuleauxCheck === "boolean") reuleauxCheck.value = sessionState.reuleauxCheck;
            if (typeof sessionState.liveShapeCheck === "boolean") liveShapeCheck.value = sessionState.liveShapeCheck;
            if (typeof sessionState.roughenAnchorsCheck === "boolean") roughenAnchorsCheck.value = sessionState.roughenAnchorsCheck;
            if (typeof sessionState.roughenAnchorsText === "string") roughenAnchorsInput.text = sessionState.roughenAnchorsText;
            applyRoughenExclusions();
        }

        /**
         * セッション状態をUIに復元する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function applyStateToUI(sessionState) {
            if (!sessionState) return;

            restoreSidesAndSize(sessionState);
            restoreAppearance(sessionState);
            restoreCornerSmoothing(sessionState);
            restoreRotation(sessionState);
            restoreStarAndCircle(sessionState);
            restoreAnchorOpsAndOptions(sessionState);

            /* 五芒星やスーパー楕円が有効なら回転を強制的にOFF / Force the rotation off for a pentagram or a superellipse */
            if (pentagramCheck.value || isSuperEllipseActive(getCurrentSides())) setRotateChecked(false);
            refreshPanelStates();
        }

        /**
         * 現在のUIの状態をセッション状態に保存する。
         * @param {object} sessionState - セッション状態
         * @returns {void}
         */
        function saveStateFromUI(sessionState) {
            if (!sessionState) return;

            var selectedSideIndex = getSelectedRadioIndex(sideRadios);
            sessionState.selectedSideIndex = (selectedSideIndex < 0) ? 0 : selectedSideIndex;
            sessionState.customSidesText = customSidesInput.text;
            sessionState.sizeText = sizeInput.text;
            sessionState.fitViewCheck = fitViewCheck.value;
            sessionState.fitViewPercentText = fitViewPercentInput.text;

            /* 塗りと線・不透明度 / Fill and stroke, opacity */
            sessionState.fillCheck = fillCheck.value;
            sessionState.strokeCheck = strokeCheck.value;
            sessionState.strokeWidthText = strokeWidthInput.text;
            /* カラーは常駐エンジンに残るオブジェクト参照ではなく文字列で保存する
               Colors are stored as strings, not as object references kept alive by the engine */
            sessionState.fillColorString = aiColorToPickerString(fillSwatch._aiColor);
            sessionState.strokeColorString = aiColorToPickerString(strokeSwatch._aiColor);
            sessionState.opacityText = opacityInput.text;

            /* 角丸 / Corner smoothing */
            sessionState.cornerRadiusCheck = cornerRadiusCheck.value;
            sessionState.cornerRadiusText = cornerRadiusInput.text;
            sessionState.smoothingValue = Math.round(smoothingSlider.value);

            /* 回転と三角形 / Rotation and triangle */
            sessionState.rotateCheck = rotateCheck.value;
            sessionState.rotateText = rotateInput.text;
            sessionState.triangleDir = triangleRightRadio.value ? "right" : (triangleLeftRadio.value ? "left" : "down");

            /* スターと円 / Star and circle */
            sessionState.starCheck = starCheck.value;
            sessionState.pentagramCheck = pentagramCheck.value;
            sessionState.innerRatioText = innerRatioInput.text;
            sessionState.superEllipseCheck = superEllipseCheck.value;
            sessionState.superExponentValue = clampSuperExponent(superExponentSlider.value);
            sessionState.circleAnchorsValue = getCircleAnchorCount();

            /* アンカーポイントの操作とオプション / Anchor point operations and options */
            sessionState.strokeCapValue = capRoundRadio.value ? "round" : (capProjectingRadio.value ? "projecting" : "butt");
            sessionState.liveShapeCheck = liveShapeCheck.value;
            sessionState.roughenAnchorsCheck = roughenAnchorsCheck.value;
            sessionState.roughenAnchorsText = roughenAnchorsInput.text;
            sessionState.reuleauxCheck = reuleauxCheck.value;
        }

        // -----------------------------------------
        // イベントの割り当て / Event bindings
        // -----------------------------------------

        /**
         * ラフ効果と併用できない項目を、ラフ効果の状態に合わせて整える。
         * @returns {void}
         */
        function applyRoughenExclusions() {
            var isRoughenOn = roughenAnchorsCheck.value;
            setControlsEnabled([roughenAnchorsInput], isRoughenOn);
            if (isRoughenOn) splitAtAnchorsCheck.value = false;
            splitAtAnchorsCheck.enabled = !isRoughenOn;
            updateStrokeCapEnabled();
        }

        /**
         * 三角形の向きを変えたときの処理。回転を必ずONにしてプレビューを更新する。
         * @returns {void}
         */
        function onTriangleDirectionChange() {
            setRotateChecked(true);
            updatePreview();
        }

        /**
         * 辺の数のラジオを選び直したときの処理。
         * @param {number} selectedIndex - 選択したインデックス
         * @returns {void}
         */
        function selectSides(selectedIndex) {
            if (pentagramCheck.value) pentagramCheck.value = false;
            selectRadio(sideRadios, selectedIndex);
            setCustomSidesEnabled(selectedIndex === CUSTOM_SIDES_INDEX);
            /* 辺の数が変わったら回転がONでも角度を更新する / Update the angle on a side change, even while rotation is on */
            applyAutoRotationForSides(getCurrentSides());
        }

        /**
         * カラースウォッチをクリックしたときにカラーピッカーを開く。
         * @param {Group} swatchGroup - 対象のスウォッチ
         * @returns {void}
         */
        function pickSwatchColor(swatchGroup) {
            var pickedColor = ColorPicker.show({
                value: aiColorToPickerString(swatchGroup._aiColor),
                title: getLabel(LABELS.dialog.colorPicker),
                lang: uiLang
            });
            if (pickedColor === null) return;
            swatchGroup._aiColor = pickerStringToAiColor(pickedColor);
            redrawSwatch(swatchGroup);
            updatePreview();
        }

        /**
         * 入力欄とスライダーを連動させ、変更のたびにプレビューを更新する。
         * 入力中は入力欄へ書き戻さず（小数点の途中でカーソルが飛ぶため）、
         * 確定（Enter・フォーカス移動）とスライダー操作のときだけ丸めた値を書き戻す。
         * @param {EditText} inputField - 入力欄
         * @param {Slider} slider - スライダー
         * @param {function} clampValue - 値を有効範囲に収める関数
         * @param {function} [afterChange] - 値を反映したあとに呼ぶ処理
         * @returns {void}
         */
        function bindValueAndSlider(inputField, slider, clampValue, afterChange) {
            /**
             * 値を反映してプレビューを更新する。
             * @param {string|number} value - 反映する値
             * @param {boolean} writeBackText - 丸めた値を入力欄へ書き戻すかどうか
             * @returns {void}
             */
            function applyValue(value, writeBackText) {
                value = clampValue(value);
                if (writeBackText) inputField.text = String(value);
                slider.value = value;
                if (afterChange) afterChange();
                updatePreview();
            }
            inputField.onChanging = function () {
                /* 入力途中は確定を待つ / Wait for the rest of a value that is still being typed */
                if (isPartialNumberInput(inputField.text)) return;
                applyValue(inputField.text, false);
            };
            inputField.onChange = function () { applyValue(inputField.text, true); };
            slider.onChanging = function () { applyValue(slider.value, true); };
        }

        /**
         * 辺の数・回転・三角形のハンドラーを割り当てる。
         * @returns {void}
         */
        function bindShapeHandlers() {
            for (var i = 0; i < sideRadios.length; i++) {
                (function (radioIndex) {
                    sideRadios[radioIndex].onClick = function () {
                        selectSides(radioIndex);
                        updatePreview();
                    };
                })(i);
            }
            bindValueAndSlider(customSidesInput, customSidesSlider, clampCustomSides);

            sizeInput.onChanging = onSizeChange;
            fitViewCheck.onClick = function () {
                updateFitViewPercentEnabled();
                applyFitView();
            };
            fitViewPercentInput.onChanging = function () {
                if (!isPartialNumberInput(fitViewPercentInput.text)) applyFitView();
            };
            fitViewPercentInput.onChange = function () {
                fitViewPercentInput.text = String(clampFitViewPercent(fitViewPercentInput.text));
                applyFitView();
            };
            rotateInput.onChanging = updatePreview;
            rotateCheck.onClick = function () {
                setRotateInputEnabled(rotateCheck.value);
                if (rotateCheck.value) applyDefaultRotationWhenEnablingRotate();
                updatePreview();
            };
            triangleRightRadio.onClick = onTriangleDirectionChange;
            triangleLeftRadio.onClick = onTriangleDirectionChange;
            triangleDownRadio.onClick = onTriangleDirectionChange;
        }

        /**
         * 塗りと線、不透明度のハンドラーを割り当てる。
         * @returns {void}
         */
        function bindAppearanceHandlers() {
            fillCheck.onClick = updatePreview;
            fillSwatch.addEventListener("click", function () { pickSwatchColor(fillSwatch); });
            strokeCheck.onClick = function () {
                updateStrokeWidthEnabled();
                updatePreview();
            };
            strokeSwatch.addEventListener("click", function () { pickSwatchColor(strokeSwatch); });
            strokeWidthInput.onChanging = updatePreview;

            /* Shiftドラッグは10%刻み / Shift-dragging snaps to steps of ten percent */
            bindValueAndSlider(opacityInput, opacitySlider, function (value) {
                if (typeof value === "number" && ScriptUI.environment.keyboardState.shiftKey) {
                    value = Math.round(value / 10) * 10;
                }
                return clampOpacity(value);
            });
        }

        /* ここから下のハンドラーは有効状態の更新をupdatePreview()に任せる。
           updatePreview() → getCurrentShapeParams() → refreshPanelStates() が毎回まとめて整える
           The handlers below leave the enabled states to updatePreview(), whose
           getCurrentShapeParams() runs refreshPanelStates() every time */

        /**
         * スター・円のハンドラーを割り当てる。
         * @returns {void}
         */
        function bindStarAndCircleHandlers() {
            starCheck.onClick = updatePreview;
            pentagramCheck.onClick = updatePreview;
            bindValueAndSlider(innerRatioInput, innerRatioSlider, clampInnerRatio, function () {
                /* 第2半径を手で決めたら五芒星の固定値から外れる / A hand-picked inner radius leaves the pentagram preset */
                pentagramCheck.value = false;
            });

            /* スーパー楕円は円（辺の数0）でだけ効き、回転はgetCurrentShapeParams()がOFFにする
               The superellipse only applies to a circle; getCurrentShapeParams() turns the rotation off */
            superEllipseCheck.onClick = updatePreview;
            superExponentSlider.onChanging = updatePreview;
            for (var i = 0; i < circleAnchorRadios.length; i++) {
                circleAnchorRadios[i].onClick = updatePreview;
            }
        }

        /**
         * 角丸のハンドラーを割り当てる。
         * @returns {void}
         */
        function bindCornerSmoothingHandlers() {
            cornerRadiusCheck.onClick = function () {
                updateCornerRadiusInputEnabled();
                updatePreview();
            };
            cornerRadiusInput.onChanging = updatePreview;
            smoothingSlider.onChanging = function () {
                smoothingValueLabel.text = Math.round(smoothingSlider.value) + "%";
                updatePreview();
            };
        }

        /**
         * アンカーポイントの操作とオプションのハンドラーを割り当てる。
         * @returns {void}
         */
        function bindAnchorOpsAndOptionHandlers() {
            roughenAnchorsCheck.onClick = function () {
                applyRoughenExclusions();
                updatePreview();
            };
            roughenAnchorsInput.onChanging = updatePreview;
            splitAtAnchorsCheck.onClick = function () {
                /* 分割後は開いたパスになるので、塗りではなく線で見せる / Split segments are open paths, so show them with a stroke */
                if (splitAtAnchorsCheck.value) {
                    fillCheck.value = false;
                    strokeCheck.value = true;
                    updateStrokeWidthEnabled();
                }
                updateStrokeCapEnabled();
                updatePreview();
            };
            capButtRadio.onClick = updatePreview;
            capRoundRadio.onClick = updatePreview;
            capProjectingRadio.onClick = updatePreview;

            reuleauxCheck.onClick = function () {
                /* 有効にするたび既定値（100%）に戻す / Reset to the default amount whenever it is enabled */
                if (reuleauxCheck.value) syncFieldAndSlider(reuleauxAmountInput, reuleauxAmountSlider, SHAPE_DEFAULTS.reuleauxAmount);
                updatePreview();
            };
            bindValueAndSlider(reuleauxAmountInput, reuleauxAmountSlider, clampReuleauxAmount);
        }

        /**
         * キーボードショートカットを割り当てる。
         * E：円（0）／A：回転／S：スター／P：五芒星／D：アンカーポイントで分割／L・R・B：三角形の向き／
         * option（Alt）＋3・4・5・6・8：辺の数。
         * @returns {void}
         */
        function bindKeyboardShortcuts() {
            /**
             * 三角形のショートカット。辺の数を3にして向きを決める。
             * @param {RadioButton} directionRadio - 向きのラジオボタン
             * @returns {void}
             */
            function applyTriangleShortcut(directionRadio) {
                selectSides(findChoiceIndex(SIDE_CHOICES, 3));
                directionRadio.value = true;
                onTriangleDirectionChange();
            }

            /**
             * option（Alt）＋数字のショートカット。辺の数を選ぶ
             * @param {number} sideIndex - SIDE_CHOICES の添字
             * @returns {Function} ショートカットの処理
             */
            function makeSidesShortcut(sideIndex) {
                return function () {
                    selectSides(sideIndex);
                    updatePreview();
                };
            }

            var shortcutMap = {
                "E": function () {
                    selectSides(findChoiceIndex(SIDE_CHOICES, 0));
                    updatePreview();
                },
                "L": function () { applyTriangleShortcut(triangleLeftRadio); },
                "R": function () { applyTriangleShortcut(triangleRightRadio); },
                "B": function () { applyTriangleShortcut(triangleDownRadio); },
                "D": splitAtAnchorsCheck,
                "A": rotateCheck,
                "S": starCheck,
                "P": function () {
                    /* 五芒星はスターがONのときだけ。スターを使えない形では何もしない
                       The pentagram needs the star to be on; nothing happens when the star is unavailable */
                    if (!starCheck.enabled) return true;
                    starCheck.value = true;
                    pentagramCheck.value = !pentagramCheck.value;
                    pentagramCheck.onClick();
                    return true;
                }
            };
            /* option（Alt）＋数字で辺の数を選ぶ。入力欄の編集中でも効かせる（文字は入力させない）
               Option (Alt) + digit picks a side count, even inside a text field, without typing the character */
            for (var sideIndex = 1; sideIndex < SIDE_CHOICES.length; sideIndex++) {
                shortcutMap["Alt+" + SIDE_CHOICES[sideIndex]] = { target: makeSidesShortcut(sideIndex), inFields: true };
            }

            addKeyShortcuts(shapeDialog, shortcutMap, {
                numericFields: [
                    customSidesInput, rotateInput, strokeWidthInput, opacityInput, sizeInput,
                    fitViewPercentInput, innerRatioInput, cornerRadiusInput, roughenAnchorsInput, reuleauxAmountInput
                ]
            });
        }

        // -----------------------------------------
        // ダイアログの組み立てと表示 / Building and showing the dialog
        // -----------------------------------------

        /**
         * パネルを縦に積むカラムを追加する。
         * @param {Group} parent - 追加先のコンテナ
         * @returns {Group} 追加したカラム
         */
        function addColumn(parent) {
            var column = parent.add("group");
            column.orientation = "column";
            column.alignChildren = "fill";
            column.alignment = "top";
            column.spacing = DENSE_PANEL_SPACING;
            return column;
        }

        /**
         * 2カラムのパネル群、画面ズーム、ボタン列を組み立てる。
         * @returns {void}
         */
        function buildDialogLayout() {
            setupWindow(shapeDialog);

            var columnsGroup = shapeDialog.add("group");
            columnsGroup.orientation = "row";
            columnsGroup.alignChildren = ["fill", "top"];
            columnsGroup.spacing = COLUMN_SPACING;

            /* 左カラム（辺の数・回転・塗りと線・幅） / Left column: sides, rotation, fill and stroke, width */
            var leftColumn = addColumn(columnsGroup);
            buildSidesPanel(leftColumn);
            buildRotatePanel(leftColumn);
            buildFillStrokePanel(leftColumn);
            buildWidthPanel(leftColumn);
            buildFitViewRow(leftColumn);

            /* 右カラム（スター・円・角丸・アンカー・オプション） / Right column: star, circle, corners, anchors, options */
            var rightColumn = addColumn(columnsGroup);
            buildStarPanel(rightColumn);
            buildCirclePanel(rightColumn);
            buildCornerSmoothingPanel(rightColumn);
            buildAnchorOpsPanel(rightColumn);
            buildOptionsPanel(rightColumn);

            /* ボタン行（左：プレビュー、右：キャンセル・OK） / Button row (left: Preview, right: Cancel and OK) */
            var buttonRow = addButtonRow(shapeDialog);
            btnPreview = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.preview));
            btnPreview.helpTip = getLabel(LABELS.tooltip.preview);
            btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
            btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        }

        /**
         * ［プレビュー］ボタン（プレビュー表示とアウトライン表示の切り替え）のハンドラーを割り当てる。
         * @returns {void}
         */
        function bindPreviewButtonHandler() {
            /* ドキュメントは通常プレビュー表示で開くので、ONから始めて表示と実態を合わせる
               A document normally opens in preview mode, so start on to keep the label truthful */
            var isViewPreviewOn = true;
            btnPreview.text = "● " + getLabel(LABELS.button.preview);
            btnPreview.onClick = function () {
                isViewPreviewOn = !isViewPreviewOn;
                btnPreview.text = (isViewPreviewOn ? "● " : "") + getLabel(LABELS.button.preview);
                app.executeMenuCommand('preview');
            };
        }

        /**
         * プレビューを巻き戻し、表示位置と倍率を開いたときの状態へ戻す。
         * @returns {void}
         */
        function discardPreview() {
            resetPreview();
            doc.activeView.centerPoint = initialViewCenter;
            doc.activeView.zoom = initialZoom;
        }

        /**
         * ［キャンセル］［OK］と、ダイアログの表示・終了時の処理を割り当てる。
         * @returns {void}
         */
        function bindDialogHandlers() {
            /* 巻き戻しはonCloseでまとめて行う / onClose does the rollback */
            btnCancel.onClick = function () { shapeDialog.close(); };
            btnOK.onClick = function () {
                /* ウィジェットが生きているうちにパラメーターを確定させる
                   Capture the parameters while the dialog widgets are still alive */
                finalParams = getCurrentShapeParams();
                applyLiveShape = liveShapeCheck.value;
                roughenAnchorsDetail = roughenAnchorsCheck.value ? parseFloat(roughenAnchorsInput.text) : 0;
                isConfirmed = true;
                shapeDialog.close();
            };

            shapeDialog.onShow = function () {
                /* 開いた時点でも一度フィットさせる / Fit once when the dialog opens, too */
                onSizeChange();
            };

            shapeDialog.onClose = function () {
                /* 保存に失敗しても閉じる処理は続ける（OK時の確定を巻き添えにしない）
                   A failed save must not abort the close, which would also abort the confirmed shape */
                try {
                    var closingState = {};
                    saveStateFromUI(closingState);
                    settingsStore.save(closingState);
                } catch (e) { }

                /* キャンセル時はプレビューを巻き戻してUndo履歴を汚さない / Roll the preview back on cancel */
                if (!isConfirmed) discardPreview();
            };
        }

        buildDialogLayout();

        bindShapeHandlers();
        bindAppearanceHandlers();
        bindStarAndCircleHandlers();
        bindCornerSmoothingHandlers();
        bindAnchorOpsAndOptionHandlers();
        bindKeyboardShortcuts();
        bindPreviewButtonHandler();
        bindDialogHandlers();

        /* レイアウトが決まる前に復元しておく。壊れた保存値でも既定値で開けるようにする
           Restore before the layout is measured; a broken stored value must still let the dialog open */
        try { applyStateToUI(getSessionState()); } catch (e) { }
        updateStrokeWidthEnabled();
        updateCornerRadiusInputEnabled();
        updateStrokeCapEnabled();
        updateFitViewPercentEnabled();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(shapeDialog, SCRIPT_NAME);
        shapeDialog.show();

        /* キャンセル時のプレビューは onClose で巻き戻し済み / A cancelled preview was already rolled back in onClose */
        if (!isConfirmed) return false;

        /* OKなら1回のUndoで取り消せる形で確定する / Finalize as a single undoable action when OK was pressed */
        previewManager.confirm(function () {
            if (!isDrawableShapeParams(finalParams)) return;
            previewShape = createShape(app.activeDocument, finalParams);
        });
        return !!previewShape;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 確定後にラフ効果でアンカーポイントを追加する。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function addAnchorsAfterConfirm(doc) {
        if (!(roughenAnchorsDetail > 0)) return;
        if (roughenAnchorsUseMenuFallback && Math.round(roughenAnchorsDetail) === 1) {
            /* この経路だけはメニューコマンドなのでプレビューには出ない
               Only this path uses a menu command, so it cannot appear in the preview */
            app.executeMenuCommand('Add Anchor Points2');
            return;
        }
        for (var i = 0; i < doc.selection.length; i++) {
            applyRoughenEffect(doc.selection[i], roughenAnchorsDetail);
        }
    }

    /**
     * ダイアログを開き、確定した図形を作成して後処理を行う。
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var rulerUnitInfo = getUnitInfo("rulerType");
        var strokeUnitInfo = getUnitInfo("strokeUnits");
        if (!showInputDialog(rulerUnitInfo, strokeUnitInfo)) return;

        /* ダイアログ表示中にドキュメントが切り替わった場合に備えて取得し直す
           Re-acquire the active document in case it changed while the dialog was open */
        var doc = app.activeDocument;
        finalizeShape(doc);
        addAnchorsAfterConfirm(doc);
        if (applyLiveShape) app.executeMenuCommand('Convert to Shape');
    }

    main();

})();
