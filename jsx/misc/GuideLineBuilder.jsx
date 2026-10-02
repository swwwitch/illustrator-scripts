#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "ExtendLinesEngine"

/*

### 概要

選択オブジェクト（グループ／複合パス／テキストを含む）から直線セグメントを拾い、描画範囲いっぱいに延長した補助線を描画します。
Bezier曲線から円を推定する［円弧から円］、アンカーポイントに円・正方形を置く機能もあります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GuideLineBuilder.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nd801b9b0367f

### Overview

Collects the straight segments of the selection — groups, compound paths and text included — and draws each one as a construction line extended across the drawing area.
It can also estimate circles from Bézier curves and place circles or squares on the anchor points.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GuideLineBuilder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GuideLineBuilder";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GuideLineBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GuideLineBuilder.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd801b9b0367f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 生成物の目印（再実行時に削除してよいものを見分ける）/ Marker that flags generated items */
    var SCRIPT_MARKER = "__ExtendLines__";

    /* 補助線を置くレイヤー名 / Layer name for construction lines */
    var LINE_LAYER_NAME = "_construction_guide";

    /* アンカー図形を置くレイヤー名 / Layer name for anchor shapes */
    var ANCHOR_LAYER_NAME = "_construction_anchorpoint";

    /* プレビュー用レイヤー名のもと / Base name of the preview layer */
    var PREVIEW_LAYER_BASE_NAME = "__ExtendLines_Preview";

    /* 線幅の初期値（mm）/ Default stroke width (mm) */
    var DEFAULT_STROKE_WIDTH_MM = 0.1;

    /* アンカー図形の大きさの初期値（mm）/ Default size of anchor shapes (mm) */
    var DEFAULT_ANCHOR_SIZE_MM = 1;

    /* 選択がアートボードの外にあるときの描画範囲の倍率 / Draw area scale when the selection is off the artboard */
    var OFF_ARTBOARD_SCALE = 4;

    /* アンカー図形のブルー [R, G, B] / Blue used for anchor shapes */
    var ANCHOR_BLUE_RGB = [78, 128, 255];

    /* 座標比較の許容値（pt）/ Tolerance for comparing coordinates (pt) */
    var POSITION_TOLERANCE = 0.001;

    /* ズームスライダーの範囲 / Range of the zoom slider */
    var ZOOM_MIN = 0.1;
    var ZOOM_MAX = 16;

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

    var NESTED_PANEL_MARGINS    = [15, 10, 15, 10];   /* 入れ子パネルの余白 [左,上,右,下] / nested panel margins */
    var ROW_SPACING             = 6;                  /* 行内の要素間隔 / spacing within a row */
    var RADIO_SPACING           = 10;                 /* 横並びラジオの間隔 / spacing between radios */
    var NUMBER_INPUT_CHARACTERS = 6;                  /* 数値入力欄の文字数 / numeric field width */
    var ZOOM_SLIDER_WIDTH       = 240;                /* ズームスライダーの幅 / zoom slider width */
    var ZOOM_ROW_MARGINS        = [0, 0, 0, 10];      /* ズームの行の余白 / zoom row margins */
    var STROKE_ROW_MARGINS      = [0, 6, 0, 0];       /* 線幅の行の余白 / stroke-width row margins */

    /**
     * パネルを追加し、共通レイアウトを設定する
     * @param {object} parentContainer - 追加先のウィンドウまたはグループ
     * @param {object} [titleSet] - ja/en を持つタイトル（省略時はタイトルなしの入れ子パネル）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentContainer, titleSet) {
        var newPanel = parentContainer.add("panel", undefined, titleSet ? getLabel(titleSet) : "");
        setupPanel(newPanel, 6);
        if (!titleSet) newPanel.margins = NESTED_PANEL_MARGINS; /* タイトルなしの入れ子は上を詰める / untitled nested panels get a smaller top margin */
        return newPanel;
    }

    /**
     * 行グループを追加し、共通レイアウトを設定する（横位置と天地を対で指定し、子は伸ばさない）
     * @param {object} parentContainer - 追加先のウィンドウまたはグループ
     * @param {number} [spacing] - 要素間隔（省略時は ROW_SPACING）
     * @returns {Group} 追加した行グループ
     */
    function addRow(parentContainer, spacing) {
        var rowGroup = parentContainer.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = ["left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
        return rowGroup;
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

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () {
            numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
            numberInput.active = true;
        });

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "補助線の描画", en: "Construction Lines" }
        },
        panel: {
            lines: { ja: "補助線を描画", en: "Draw Construction Lines" },
            anchorShapes: { ja: "アンカーポイントに図形", en: "Shapes on Anchor Points" },
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            straight: { ja: "直線", en: "Straight segments" },
            horizontal: { ja: "水平線", en: "Horizontal lines" },
            vertical: { ja: "垂直線", en: "Vertical lines" },
            diagonal: { ja: "斜線", en: "Diagonal lines" },
            arcToCircle: { ja: "円弧から円", en: "Create circles from arcs" },
            group: { ja: "グループ化", en: "Group output" },
            separateLayer: { ja: "別レイヤーに", en: "Use separate layer" },
            guide: { ja: "ガイド化", en: "Convert to guides" },
            dedup: { ja: "線のダブりを削除", en: "Remove duplicate lines" },
            preview: { ja: "プレビュー", en: "Preview" },
            lightMode: { ja: "軽量モード", en: "Light Preview" }
        },
        radio: {
            anchorNone: { ja: "なし", en: "None" },
            anchorCircle: { ja: "円", en: "Circle" },
            anchorSquare: { ja: "正方形", en: "Square" },
            colorBlack: { ja: "黒", en: "Black" },
            colorBlue: { ja: "ブルー", en: "Blue" },
            arcIgnore: { ja: "無視", en: "Ignore" },
            arcChord: { ja: "弦", en: "Chord" },
            arcExtendChord: { ja: "弦を延長", en: "Extend chord" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Stroke Width" },
            anchorSize: { ja: "大きさ", en: "Size" },
            zoom: { ja: "ズーム", en: "Zoom" }
        },
        tooltip: {
            arcFallback: { ja: "完全な円弧以外の場合", en: "For segments that are not true circular arcs" },
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
            straight: {
                ja: "選択したパスの直線セグメントを、描画範囲いっぱいまで延長して描きます",
                en: "Extends each straight segment of the selection across the drawing area"
            },
            soloDirection: {
                ja: "option キーを押しながらクリックすると、この向きだけを残します",
                en: "Option-click to keep only this direction"
            },
            arcToCircle: {
                ja: "円弧とみなせる曲線セグメントから、その円を描きます",
                en: "Draws the full circle of each curved segment that is a circular arc"
            },
            arcIgnore: { ja: "完全な円弧でないセグメントには何も描きません", en: "Draws nothing for segments that are not true arcs" },
            arcChord: {
                ja: "完全な円弧でないセグメントは、両端を結ぶ線分を描きます",
                en: "Draws the chord between the end points of segments that are not true arcs"
            },
            arcExtendChord: {
                ja: "完全な円弧でないセグメントは、両端を結ぶ直線を描画範囲いっぱいまで延長して描きます",
                en: "Extends the chord of segments that are not true arcs across the drawing area"
            },
            group: {
                ja: "補助線とアンカーポイントの図形を、それぞれグループにまとめます",
                en: "Groups the construction lines and the anchor shapes separately"
            },
            separateLayer: {
                ja: "補助線は「_construction_guide」、図形は「_construction_anchorpoint」レイヤーに描き、前回このスクリプトで描いたものは置き換えます",
                en: "Draws lines on \"_construction_guide\" and shapes on \"_construction_anchorpoint\", replacing what this script drew there before"
            },
            guide: {
                ja: "補助線と円をガイドにします（アンカーポイントの図形はガイドにしません）",
                en: "Turns the lines and circles into guides (anchor shapes stay as paths)"
            },
            dedup: {
                ja: "延長した補助線が同じ位置に重なるときは1本だけ描きます",
                en: "Draws only one line when extended lines fall in the same place"
            },
            zoom: {
                ja: "作業中の表示倍率を変えます。option キーでアートボードの中央、shift キーで選択の中央に合わせます",
                en: "Changes the view zoom. Hold Option to center on the artboard, or Shift to center on the selection"
            },
            lightMode: {
                ja: "ズームのスライダーを離したときだけ表示倍率を変えます",
                en: "Applies the zoom only when you release the slider"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        itemName: {
            history: { ja: "補助線の描画", en: "Construction Lines" },
            lineGroup: { ja: "補助線", en: "Construction Lines" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            noValidPath: { ja: "有効なパスが見つかりません。", en: "No valid paths were found." },
            error: { ja: "エラー: ", en: "Error: " }
        }
    };

    // =========================================
    // 定数 / Constants
    // =========================================

    /* アンカーポイントに置く図形 / Shapes drawn on anchor points */
    var ANCHOR_SHAPE = {
        NONE: "NONE",
        CIRCLE: "CIRCLE",
        SQUARE: "SQUARE"
    };

    /* アンカー図形の色 / Color of anchor shapes */
    var ANCHOR_COLOR = {
        BLACK: "BLACK",
        BLUE: "BLUE"
    };

    /* 完全な円弧として扱えないカーブの処理 / Fallback for curves that are not true arcs */
    var ARC_FALLBACK = {
        IGNORE: "IGNORE",
        CHORD: "CHORD",
        EXTEND: "EXTEND"
    };

    /* 直線セグメントの向き / Direction of a straight segment */
    var SEGMENT_DIRECTION = {
        HORIZONTAL: "HORIZONTAL",
        VERTICAL: "VERTICAL",
        DIAGONAL: "DIAGONAL"
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

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * エラーを ExtendScript コンソールへ書き出す
     * @param {Error} err - 捕まえた例外
     * @param {string} context - 発生箇所を示す文字列
     * @returns {void}
     */
    function logError(err, context) {
        $.writeln("[" + SCRIPT_NAME + "] " + context + ": " + err);
    }

    /**
     * mm を pt に変換する
     * @param {number} millimeters - ミリメートル値
     * @returns {number} ポイント値
     */
    function mmToPt(millimeters) {
        return millimeters * 72.0 / 25.4;
    }

    /**
     * 2点がほぼ同じ位置かどうかを判定する
     * @param {Array<number>} pointA - 座標 [x, y]
     * @param {Array<number>} pointB - 座標 [x, y]
     * @returns {boolean} ほぼ同じ位置なら true
     */
    function isSamePosition(pointA, pointB) {
        return Math.abs(pointA[0] - pointB[0]) < POSITION_TOLERANCE &&
            Math.abs(pointA[1] - pointB[1]) < POSITION_TOLERANCE;
    }

    /**
     * 座標を丸めて重複判定用のキーにする
     * @param {Array<number>} point - 座標 [x, y]
     * @returns {string} 重複判定キー
     */
    function makePointKey(point) {
        return Number(point[0]).toFixed(3) + "," + Number(point[1]).toFixed(3);
    }

    /**
     * 2点から線分の重複判定キーを作る（向きは無視）
     * @param {Array<number>} pointA - 座標 [x, y]
     * @param {Array<number>} pointB - 座標 [x, y]
     * @returns {string} 重複判定キー
     */
    function makeLineKey(pointA, pointB) {
        var keyA = makePointKey(pointA);
        var keyB = makePointKey(pointB);
        return (keyA < keyB) ? (keyA + "|" + keyB) : (keyB + "|" + keyA);
    }

    /**
     * 入力文字列を数値に直す（全角小数点・カンマ区切りを吸収する）
     * @param {string} text - 入力欄の文字列
     * @returns {number} 数値（読めない場合は NaN）
     */
    function parseNumberInput(text) {
        var normalized = String(text).replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        /* "1,5" は小数点、"1,000" は桁区切りとみなす / Treat "1,5" as a decimal and "1,000" as a separator */
        normalized = (normalized.indexOf(",") >= 0 && normalized.indexOf(".") < 0) ?
            normalized.replace(/,/g, ".") :
            normalized.replace(/,/g, "");
        return Number(normalized);
    }

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
     * 入力欄の文字列を pt に直す（読めない場合は直前の有効値を返す）
     * @param {EditText} editText - 対象の入力欄
     * @param {string} prefKey - 単位を決める環境設定キー（strokeUnits / rulerType）
     * @param {object} lastValidCache - 直前の有効値を持つ `{ value: number }`
     * @returns {number} ポイント値
     */
    function readLengthAsPt(editText, prefKey, lastValidCache) {
        var pointsPerUnit = getUnitInfo(prefKey).pointsPerUnit;
        var inputValue = parseNumberInput(editText.text);
        if (!(inputValue > 0)) return lastValidCache.value;

        var lengthPt = inputValue * pointsPerUnit;
        if (!(lengthPt > 0)) return lastValidCache.value;

        lastValidCache.value = lengthPt;
        return lengthPt;
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * 名前でレイヤーを探す
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー（無ければ null）
     */
    function findLayerByName(doc, layerName) {
        try { return doc.layers.getByName(layerName); } catch (e) { }
        return null;
    }

    /**
     * 既存と重ならないレイヤー名を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {string} baseName - もとにする名前
     * @returns {string} 重複しないレイヤー名
     */
    function createUniqueLayerName(doc, baseName) {
        var candidateName = baseName;
        var suffixNumber = 2;
        while (findLayerByName(doc, candidateName)) {
            candidateName = baseName + "_" + suffixNumber;
            suffixNumber++;
        }
        return candidateName;
    }

    /**
     * 名前が一致するレイヤーがあれば削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function removeLayerIfExists(doc, layerName) {
        var targetLayer = findLayerByName(doc, layerName);
        if (!targetLayer) return;
        /* ロックなどで削除できないことがある / Removal can fail, e.g. on a locked layer */
        try { targetLayer.remove(); } catch (e) { logError(e, "removeLayerIfExists"); }
    }

    /**
     * アイテムがこのスクリプトの生成物かどうかを判定する
     * @param {PageItem} item - 対象のアイテム
     * @returns {boolean} 生成物なら true
     */
    function isGeneratedItem(item) {
        /* note / name を読めないアイテムは生成物ではないとみなす / Treat unreadable items as not generated */
        try {
            if (item.note === SCRIPT_MARKER) return true;
            return String(item.name || "").indexOf(SCRIPT_MARKER) === 0;
        } catch (e) { }
        return false;
    }

    /**
     * このスクリプトが生成したオブジェクトだけをレイヤーから削除する
     * @param {Layer} layer - 対象のレイヤー
     * @returns {void}
     */
    function clearGeneratedItemsInLayer(layer) {
        if (!layer) return;
        /* ロックされたアイテムなどは削除できない / Locked items cannot be removed */
        try {
            for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                var layerItem = layer.pageItems[i];
                if (layerItem && isGeneratedItem(layerItem)) layerItem.remove();
            }
        } catch (e) { logError(e, "clearGeneratedItemsInLayer"); }
    }

    /**
     * 選択がすべて指定レイヤー上にあるかどうかを判定する
     * @param {Array} selectedItems - 選択アイテム
     * @param {string} layerName - レイヤー名
     * @returns {boolean} すべてそのレイヤー上なら true
     */
    function isSelectionOnLayer(selectedItems, layerName) {
        if (!selectedItems || selectedItems.length === 0) return false;
        for (var i = 0; i < selectedItems.length; i++) {
            var itemLayer = null;
            /* layer を持たないアイテムがある / Some items have no layer */
            try { itemLayer = selectedItems[i].layer; } catch (e) { }
            if (!itemLayer || itemLayer.name !== layerName) return false;
        }
        return true;
    }

    /**
     * 補助線の描画先レイヤーを用意する（再実行時は自分の生成物だけ消す）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @returns {Layer} 描画先のレイヤー
     */
    function prepareLineLayer(doc, selectedItems) {
        var lineLayer;

        if (isSelectionOnLayer(selectedItems, LINE_LAYER_NAME)) {
            /* 選択自体が対象レイヤー上にあるので、消さずに退避して新しく作る / Keep the selection by backing up the layer */
            var existingLayer = findLayerByName(doc, LINE_LAYER_NAME);
            if (existingLayer) existingLayer.name = createUniqueLayerName(doc, LINE_LAYER_NAME + "_backup");
            lineLayer = doc.layers.add();
            lineLayer.name = LINE_LAYER_NAME;
        } else {
            lineLayer = findLayerByName(doc, LINE_LAYER_NAME);
            if (lineLayer) {
                clearGeneratedItemsInLayer(lineLayer);
            } else {
                lineLayer = doc.layers.add();
                lineLayer.name = LINE_LAYER_NAME;
            }
        }

        lineLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        return lineLayer;
    }

    /**
     * アンカー図形の描画先レイヤーを用意する（図形を作らないときも前回の生成物は消す）
     * @param {Document} doc - 対象のドキュメント
     * @param {boolean} needsLayer - 今回アンカー図形を作るかどうか
     * @returns {Layer|null} 描画先のレイヤー（作らない場合は null）
     */
    function prepareAnchorLayer(doc, needsLayer) {
        var anchorLayer = findLayerByName(doc, ANCHOR_LAYER_NAME);
        if (anchorLayer) clearGeneratedItemsInLayer(anchorLayer);

        if (!needsLayer) return null;

        if (!anchorLayer) {
            anchorLayer = doc.layers.add();
            anchorLayer.name = ANCHOR_LAYER_NAME;
        }
        anchorLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        return anchorLayer;
    }

    /**
     * 目印付きのグループを追加する
     * @param {Layer|GroupItem} parentContainer - 追加先のレイヤーまたはグループ
     * @param {string} groupName - グループ名
     * @returns {GroupItem} 追加したグループ
     */
    function addMarkedGroup(parentContainer, groupName) {
        var markedGroup = parentContainer.groupItems.add();
        markedGroup.name = groupName;
        markedGroup.note = SCRIPT_MARKER;
        return markedGroup;
    }

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
    // パスの収集 / Collecting paths
    // =========================================

    /**
     * 選択内のテキストを複製してアウトライン化する（元のテキストは触らない）
     * @param {Array} items - 選択アイテム
     * @returns {Array<PageItem>} 一時的に作ったアウトラインの入れ物
     */
    function outlineTextFromSelection(items) {
        var outlineRoots = [];

        /* グループの中のテキストも拾う / Text inside groups is included */
        var textFrames = collectSelectionTextFrames(items, { unique: false });
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            try {
                /* 同じレイヤーの末尾に複製し、複製だけをアウトライン化する / Duplicate first, outline the copy only */
                var textCopy = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATEND);
                var outlineGroup = textCopy.createOutline();
                /* createOutline() は複製を消費するので、ふつうは例外になる / createOutline() consumes the copy, so this usually throws */
                try { textCopy.remove(); } catch (e) { }
                if (outlineGroup) outlineRoots.push(outlineGroup);
            } catch (e) {
                /* 1つ失敗しても全体は止めない / Keep going even if one frame fails */
                logError(e, "outlineTextFromSelection/createOutline");
            }
        }

        return outlineRoots;
    }

    /**
     * 一時的に作ったアウトラインを削除する
     * @param {Array<PageItem>} outlineRoots - 削除するアイテム
     * @returns {void}
     */
    function cleanupTempOutlines(outlineRoots) {
        if (!outlineRoots) return;
        for (var i = outlineRoots.length - 1; i >= 0; i--) {
            /* すでに消えているものは飛ばす / Skip items that are already gone */
            try { outlineRoots[i].remove(); } catch (e) { }
        }
    }

    // =========================================
    // ジオメトリ / Geometry
    // =========================================

    /**
     * @typedef {object} DrawArea
     * @property {number} left - 左端
     * @property {number} top - 上端
     * @property {number} right - 右端
     * @property {number} bottom - 下端
     */

    /**
     * 補助線を伸ばす範囲を求める（通常はアートボード、選択がアートボードの外なら選択中心の矩形）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @returns {DrawArea} 描画範囲
     */
    function getDrawArea(doc, selectedItems) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var artboardArea = {
            left: artboardRect[0],
            top: artboardRect[1],
            right: artboardRect[2],
            bottom: artboardRect[3]
        };

        var selectionBounds = getClipAwareUnionBounds(selectedItems, false);
        if (!selectionBounds) return artboardArea;

        var selectionLeft = selectionBounds[0], selectionTop = selectionBounds[1];
        var selectionRight = selectionBounds[2], selectionBottom = selectionBounds[3];

        /* Illustrator 座標は上が大きい / In Illustrator coordinates, top is the larger Y */
        var intersects = !(selectionRight < artboardArea.left || selectionLeft > artboardArea.right ||
            selectionTop < artboardArea.bottom || selectionBottom > artboardArea.top);
        if (intersects) return artboardArea;

        /* 選択がアートボードと全く重ならないときは、選択を中心にした矩形を使う / Use a rect around the selection instead */
        var areaWidth = Math.max(selectionRight - selectionLeft, 1) * OFF_ARTBOARD_SCALE;
        var areaHeight = Math.max(selectionTop - selectionBottom, 1) * OFF_ARTBOARD_SCALE;
        var centerX = (selectionLeft + selectionRight) / 2;
        var centerY = (selectionTop + selectionBottom) / 2;

        return {
            left: centerX - areaWidth / 2,
            top: centerY + areaHeight / 2,
            right: centerX + areaWidth / 2,
            bottom: centerY - areaHeight / 2
        };
    }

    /**
     * 2つのアンカーポイント間が直線（ハンドルが出ていない）かどうかを判定する
     * @param {PathPoint} startPoint - 始点のアンカーポイント
     * @param {PathPoint} endPoint - 終点のアンカーポイント
     * @returns {boolean} 直線なら true
     */
    function isStraightSegment(startPoint, endPoint) {
        return isSamePosition(startPoint.rightDirection, startPoint.anchor) &&
            isSamePosition(endPoint.leftDirection, endPoint.anchor);
    }

    /**
     * 直線セグメントの向きを分類する
     * @param {Array<number>} start - 始点の座標 [x, y]
     * @param {Array<number>} end - 終点の座標 [x, y]
     * @returns {string} SEGMENT_DIRECTION のいずれか
     */
    function classifySegmentDirection(start, end) {
        if (Math.abs(start[1] - end[1]) < POSITION_TOLERANCE) return SEGMENT_DIRECTION.HORIZONTAL;
        if (Math.abs(start[0] - end[0]) < POSITION_TOLERANCE) return SEGMENT_DIRECTION.VERTICAL;
        return SEGMENT_DIRECTION.DIAGONAL;
    }

    /**
     * ベクトルの長さを求める
     * @param {Array<number>} vector - ベクトル [x, y]
     * @returns {number} 長さ
     */
    function vectorLength(vector) {
        return Math.sqrt(vector[0] * vector[0] + vector[1] * vector[1]);
    }

    /**
     * ベクトルの差を求める
     * @param {Array<number>} pointA - 座標 [x, y]
     * @param {Array<number>} pointB - 座標 [x, y]
     * @returns {Array<number>} pointA - pointB
     */
    function subtractVector(pointA, pointB) {
        return [pointA[0] - pointB[0], pointA[1] - pointB[1]];
    }

    /**
     * ベクトルの内積を求める
     * @param {Array<number>} vectorA - ベクトル [x, y]
     * @param {Array<number>} vectorB - ベクトル [x, y]
     * @returns {number} 内積
     */
    function dotProduct(vectorA, vectorB) {
        return vectorA[0] * vectorB[0] + vectorA[1] * vectorB[1];
    }

    /**
     * ベクトルを90度回す
     * @param {Array<number>} vector - ベクトル [x, y]
     * @returns {Array<number>} 直交するベクトル
     */
    function rotate90(vector) {
        return [-vector[1], vector[0]];
    }

    /**
     * 点と方向で表した2直線の交点を求める
     * @param {Array<number>} pointA - 直線1が通る点 [x, y]
     * @param {Array<number>} directionA - 直線1の方向 [x, y]
     * @param {Array<number>} pointB - 直線2が通る点 [x, y]
     * @param {Array<number>} directionB - 直線2の方向 [x, y]
     * @returns {Array<number>|null} 交点の座標（平行な場合は null）
     */
    function intersectLines(pointA, directionA, pointB, directionB) {
        var denominator = directionA[0] * directionB[1] - directionA[1] * directionB[0];
        if (Math.abs(denominator) < 1e-9) return null;

        var deltaX = pointB[0] - pointA[0];
        var deltaY = pointB[1] - pointA[1];
        var ratio = (deltaX * directionB[1] - deltaY * directionB[0]) / denominator;

        return [pointA[0] + directionA[0] * ratio, pointA[1] + directionA[1] * ratio];
    }

    /**
     * 3次ベジェ曲線上の点を求める
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} startHandle - 始点のハンドル [x, y]
     * @param {Array<number>} endHandle - 終点のハンドル [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {number} t - パラメーター（0〜1）
     * @returns {Array<number>} 曲線上の座標 [x, y]
     */
    function cubicBezierPoint(start, startHandle, endHandle, end, t) {
        var u = 1 - t;
        var uuu = u * u * u;
        var uut = 3 * u * u * t;
        var utt = 3 * u * t * t;
        var ttt = t * t * t;

        return [
            uuu * start[0] + uut * startHandle[0] + utt * endHandle[0] + ttt * end[0],
            uuu * start[1] + uut * startHandle[1] + utt * endHandle[1] + ttt * end[1]
        ];
    }

    /**
     * 円弧の両端の接線から円の中心を求める
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} startHandle - 始点のハンドル [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {Array<number>} endHandle - 終点のハンドル [x, y]
     * @returns {Array<number>|null} 中心の座標（求められない場合は null）
     */
    function getArcCenter(start, startHandle, end, endHandle) {
        var startTangent = subtractVector(startHandle, start);
        var endTangent = subtractVector(end, endHandle);
        return intersectLines(start, rotate90(startTangent), end, rotate90(endTangent));
    }

    /**
     * 曲線セグメントを円弧とみなしてよいかを判定する（両端と中点が同じ円上にあり、接線が半径と直交する）
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} startHandle - 始点のハンドル [x, y]
     * @param {Array<number>} endHandle - 終点のハンドル [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {Array<number>} center - 推定した中心 [x, y]
     * @param {number} radius - 推定した半径
     * @returns {boolean} 円弧とみなせるなら true
     */
    function isApproxCircularArc(start, startHandle, endHandle, end, center, radius) {
        if (!center || !(radius > 0)) return false;

        /* 半径の1%、ただし最低 0.2pt を許容値にする / Tolerance: 1% of the radius, at least 0.2pt */
        var tolerance = Math.max(0.2, radius * 0.01);

        if (Math.abs(vectorLength(subtractVector(start, center)) - radius) > tolerance) return false;
        if (Math.abs(vectorLength(subtractVector(end, center)) - radius) > tolerance) return false;

        var midPoint = cubicBezierPoint(start, startHandle, endHandle, end, 0.5);
        if (Math.abs(vectorLength(subtractVector(midPoint, center)) - radius) > tolerance) return false;

        var startTangent = subtractVector(startHandle, start);
        var endTangent = subtractVector(end, endHandle);
        if (vectorLength(startTangent) < 1e-6 || vectorLength(endTangent) < 1e-6) return false;

        if (Math.abs(dotProduct(subtractVector(start, center), startTangent)) > tolerance * vectorLength(startTangent)) return false;
        if (Math.abs(dotProduct(subtractVector(end, center), endTangent)) > tolerance * vectorLength(endTangent)) return false;

        return true;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * K100 のカラーを作る
     * @returns {CMYKColor} 黒100%のカラー
     */
    function createBlackColor() {
        var blackColor = new CMYKColor();
        blackColor.cyan = 0;
        blackColor.magenta = 0;
        blackColor.yellow = 0;
        blackColor.black = 100;
        return blackColor;
    }

    /**
     * 補助線のスタイルを適用する（塗りなし、ガイド化またはK100の線）
     * @param {PathItem} pathItem - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyLineStyle(pathItem, settings) {
        pathItem.note = SCRIPT_MARKER;
        pathItem.filled = false;
        pathItem.fillColor = new NoColor();

        if (settings.guide) {
            pathItem.stroked = false;
            pathItem.guides = true;
            return;
        }

        pathItem.stroked = true;
        pathItem.strokeColor = createBlackColor();
        pathItem.strokeWidth = settings.strokeWidthPt;
    }

    /**
     * アンカー図形のスタイルを適用する（塗りのみ、線なし。ガイド化はしない）
     * @param {PathItem} pathItem - 対象のパス
     * @param {string} anchorColor - ANCHOR_COLOR のいずれか
     * @returns {void}
     */
    function applyAnchorShapeStyle(pathItem, anchorColor) {
        pathItem.note = SCRIPT_MARKER;
        pathItem.filled = true;

        if (anchorColor === ANCHOR_COLOR.BLUE) {
            var blueColor = new RGBColor();
            blueColor.red = ANCHOR_BLUE_RGB[0];
            blueColor.green = ANCHOR_BLUE_RGB[1];
            blueColor.blue = ANCHOR_BLUE_RGB[2];
            pathItem.fillColor = blueColor;
        } else {
            pathItem.fillColor = createBlackColor();
        }

        pathItem.stroked = false;
        pathItem.strokeColor = new NoColor();
        pathItem.guides = false;
    }

    /**
     * アンカーポイントの位置に図形（円または正方形）を作る
     * @param {Array<PathItem>} pathItems - 対象のパス
     * @param {Layer|GroupItem} container - 図形の追加先
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function createAnchorShapes(pathItems, container, settings) {
        var shapeSize = settings.anchorSizePt;
        var halfSize = shapeSize / 2;
        var placedPointKeys = {};

        for (var i = 0; i < pathItems.length; i++) {
            var pathPoints = pathItems[i].pathPoints;
            if (!pathPoints) continue;

            for (var j = 0; j < pathPoints.length; j++) {
                var anchorPosition = pathPoints[j].anchor;
                if (!anchorPosition) continue;

                /* 同じ位置に重ねて作らない / Do not stack shapes on the same position */
                var pointKey = makePointKey(anchorPosition);
                if (placedPointKeys[pointKey]) continue;
                placedPointKeys[pointKey] = true;

                var shapeTop = anchorPosition[1] + halfSize;
                var shapeLeft = anchorPosition[0] - halfSize;
                var anchorShapeItem = (settings.anchorShape === ANCHOR_SHAPE.CIRCLE) ?
                    container.pathItems.ellipse(shapeTop, shapeLeft, shapeSize, shapeSize) :
                    container.pathItems.rectangle(shapeTop, shapeLeft, shapeSize, shapeSize);

                anchorShapeItem.closed = true;
                applyAnchorShapeStyle(anchorShapeItem, settings.anchorColor);
            }
        }
    }

    /**
     * 2点を結ぶ開いたパス（弦など）を描いて補助線のスタイルを適用する
     * @param {Layer|GroupItem} container - 線の追加先
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function drawOpenLine(container, start, end, settings) {
        var openLine = container.pathItems.add();
        openLine.setEntirePath([start, end]);
        openLine.closed = false;
        applyLineStyle(openLine, settings);
    }

    /**
     * 2点を通る直線と描画範囲との交点を求める
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {DrawArea} area - 描画範囲
     * @returns {Array<Array<number>>} 交点の配列
     */
    function getDrawAreaIntersections(start, end, area) {
        var intersections = [];

        if (Math.abs(start[0] - end[0]) < POSITION_TOLERANCE) {
            /* 垂直線 / Vertical line */
            intersections.push([start[0], area.top]);
            intersections.push([start[0], area.bottom]);
            return intersections;
        }

        if (Math.abs(start[1] - end[1]) < POSITION_TOLERANCE) {
            /* 水平線 / Horizontal line */
            intersections.push([area.left, start[1]]);
            intersections.push([area.right, start[1]]);
            return intersections;
        }

        var slope = (end[1] - start[1]) / (end[0] - start[0]);
        var intercept = start[1] - slope * start[0];

        var yAtLeft = slope * area.left + intercept;
        if (yAtLeft <= area.top + POSITION_TOLERANCE && yAtLeft >= area.bottom - POSITION_TOLERANCE) {
            intersections.push([area.left, yAtLeft]);
        }

        var yAtRight = slope * area.right + intercept;
        if (yAtRight <= area.top + POSITION_TOLERANCE && yAtRight >= area.bottom - POSITION_TOLERANCE) {
            intersections.push([area.right, yAtRight]);
        }

        var xAtTop = (area.top - intercept) / slope;
        if (xAtTop >= area.left - POSITION_TOLERANCE && xAtTop <= area.right + POSITION_TOLERANCE) {
            intersections.push([xAtTop, area.top]);
        }

        var xAtBottom = (area.bottom - intercept) / slope;
        if (xAtBottom >= area.left - POSITION_TOLERANCE && xAtBottom <= area.right + POSITION_TOLERANCE) {
            intersections.push([xAtBottom, area.bottom]);
        }

        return intersections;
    }

    /**
     * 2点を通る直線を描画範囲いっぱいまで延長して描く
     * @param {Layer|GroupItem} container - 線の追加先
     * @param {Array<number>} start - 始点 [x, y]
     * @param {Array<number>} end - 終点 [x, y]
     * @param {DrawArea} area - 描画範囲
     * @param {object} settings - ダイアログの設定
     * @param {object} dedupMap - 重複判定用のキー置き場
     * @returns {void}
     */
    function drawLineAcrossDrawArea(container, start, end, area, settings, dedupMap) {
        var intersections = getDrawAreaIntersections(start, end, area);
        if (intersections.length < 2) return;

        /* 同じ位置の交点（角をかすめた場合など）は1つにまとめる / Collapse duplicate hits, e.g. at a corner */
        var lineStart = intersections[0];
        var lineEnd = null;
        for (var i = 1; i < intersections.length; i++) {
            if (!isSamePosition(intersections[i], lineStart)) {
                lineEnd = intersections[i];
                break;
            }
        }
        if (!lineEnd) return;

        if (settings.dedup) {
            var lineKey = makeLineKey(lineStart, lineEnd);
            if (dedupMap[lineKey]) return;
            dedupMap[lineKey] = true;
        }

        drawOpenLine(container, lineStart, lineEnd, settings);
    }

    /**
     * パスのセグメント数を返す（3点以上の閉じたパスは最後の点から最初の点へのセグメントも数える）
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} セグメント数
     */
    function getSegmentCount(pathItem) {
        var pointCount = pathItem.pathPoints.length;
        return (pathItem.closed && pointCount >= 3) ? pointCount : pointCount - 1;
    }

    /**
     * パスの中から曲線セグメント（どちらかにハンドルが出ているセグメント）を集める
     * @param {PathItem} pathItem - 対象のパス
     * @returns {Array<object>} `{ start, end }` の配列
     */
    function getCurvedSegments(pathItem) {
        var pathPoints = pathItem.pathPoints;
        var segmentCount = getSegmentCount(pathItem);
        var segments = [];

        for (var i = 0; i < segmentCount; i++) {
            var startPoint = pathPoints[i];
            var endPoint = pathPoints[(i + 1) % pathPoints.length];
            if (!isStraightSegment(startPoint, endPoint)) {
                segments.push({ start: startPoint, end: endPoint });
            }
        }

        return segments;
    }

    /**
     * 曲線セグメントから円を推定して描く（円弧とみなせない場合は設定に従う）
     * @param {PathItem} pathItem - 対象のパス
     * @param {Layer|GroupItem} container - 図形の追加先
     * @param {object} settings - ダイアログの設定
     * @param {DrawArea} area - 描画範囲
     * @param {object} dedupMap - 重複判定用のキー置き場
     * @returns {void}
     */
    function createCirclesFromArcPath(pathItem, container, settings, area, dedupMap) {
        var segments = getCurvedSegments(pathItem);

        for (var i = 0; i < segments.length; i++) {
            var start = segments[i].start.anchor;
            var end = segments[i].end.anchor;
            var startHandle = segments[i].start.rightDirection;
            var endHandle = segments[i].end.leftDirection;

            var center = getArcCenter(start, startHandle, end, endHandle);
            if (!center) continue;

            var radius = vectorLength(subtractVector(start, center));
            if (!(radius > 0)) continue;

            if (!isApproxCircularArc(start, startHandle, endHandle, end, center, radius)) {
                if (settings.arcFallback === ARC_FALLBACK.CHORD) {
                    drawOpenLine(container, start, end, settings);
                } else if (settings.arcFallback === ARC_FALLBACK.EXTEND) {
                    drawLineAcrossDrawArea(container, start, end, area, settings, dedupMap);
                }
                /* IGNORE は何もしない / IGNORE draws nothing */
                continue;
            }

            var circle = container.pathItems.ellipse(
                center[1] + radius,
                center[0] - radius,
                radius * 2,
                radius * 2
            );
            applyLineStyle(circle, settings);
        }
    }

    /**
     * 直線セグメントの向きが描画対象に選ばれているかを判定する
     * @param {string} direction - SEGMENT_DIRECTION のいずれか
     * @param {object} settings - ダイアログの設定
     * @returns {boolean} 描画する向きなら true
     */
    function isDirectionEnabled(direction, settings) {
        if (direction === SEGMENT_DIRECTION.HORIZONTAL) return settings.horizontal;
        if (direction === SEGMENT_DIRECTION.VERTICAL) return settings.vertical;
        return settings.diagonal;
    }

    /**
     * パスの直線セグメントを描画範囲いっぱいまで延長して描く
     * @param {PathItem} pathItem - 対象のパス
     * @param {Layer|GroupItem} container - 線の追加先
     * @param {object} settings - ダイアログの設定
     * @param {DrawArea} area - 描画範囲
     * @param {object} dedupMap - 重複判定用のキー置き場
     * @returns {void}
     */
    function drawExtensionsFromPath(pathItem, container, settings, area, dedupMap) {
        var pathPoints = pathItem.pathPoints;
        if (!pathPoints || pathPoints.length < 2) return;

        var segmentCount = getSegmentCount(pathItem);

        for (var i = 0; i < segmentCount; i++) {
            var startPoint = pathPoints[i];
            var endPoint = pathPoints[(i + 1) % pathPoints.length];
            if (!isStraightSegment(startPoint, endPoint)) continue;

            var start = startPoint.anchor;
            var end = endPoint.anchor;

            /* 2点が重なっているゴミパスは飛ばす / Skip degenerate segments */
            if (isSamePosition(start, end)) continue;
            if (!isDirectionEnabled(classifySegmentDirection(start, end), settings)) continue;

            drawLineAcrossDrawArea(container, start, end, area, settings, dedupMap);
        }
    }

    /**
     * 設定に従って、アンカー図形・円・補助線をまとめて描く
     * @param {Array<PathItem>} targetPaths - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @param {DrawArea} area - 描画範囲
     * @param {Layer|GroupItem} lineContainer - 線と円の追加先
     * @param {Layer|GroupItem} anchorContainer - アンカー図形の追加先
     * @returns {void}
     */
    function drawConstructionItems(targetPaths, settings, area, lineContainer, anchorContainer) {
        var dedupMap = {};
        var i;

        if (settings.anchorShape !== ANCHOR_SHAPE.NONE) {
            createAnchorShapes(targetPaths, anchorContainer, settings);
        }

        if (settings.arcToCircle) {
            for (i = 0; i < targetPaths.length; i++) {
                try {
                    createCirclesFromArcPath(targetPaths[i], lineContainer, settings, area, dedupMap);
                } catch (e) { logError(e, "createCirclesFromArcPath"); }
            }
        }

        if (settings.straightLines) {
            for (i = 0; i < targetPaths.length; i++) {
                drawExtensionsFromPath(targetPaths[i], lineContainer, settings, area, dedupMap);
            }
        }
    }

    /**
     * @typedef {object} OutputContainers
     * @property {Layer|GroupItem} lineContainer - 線と円の追加先
     * @property {Layer|GroupItem} anchorContainer - アンカー図形の追加先
     */

    /**
     * 設定に従って描画先のレイヤーとグループを用意する
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @param {Layer} activeLayer - もとのアクティブレイヤー
     * @param {object} settings - ダイアログの設定
     * @returns {OutputContainers} 描画先
     */
    function prepareOutputContainers(doc, selectedItems, activeLayer, settings) {
        var lineLayer = activeLayer;
        var anchorLayer = activeLayer;
        var needsAnchors = (settings.anchorShape !== ANCHOR_SHAPE.NONE);

        if (settings.separateLayer) {
            lineLayer = prepareLineLayer(doc, selectedItems);
            anchorLayer = prepareAnchorLayer(doc, needsAnchors) || lineLayer;
        }

        var outputContainers = { lineContainer: lineLayer, anchorContainer: anchorLayer };
        if (!settings.group) return outputContainers;

        /* 空のグループを残さないよう、描くものがある側だけグループを作る / Only group what will actually be drawn */
        if (settings.straightLines || settings.arcToCircle) {
            outputContainers.lineContainer = addMarkedGroup(lineLayer, SCRIPT_MARKER + "_" + getLabel(LABELS.itemName.lineGroup));
        }
        if (needsAnchors) {
            outputContainers.anchorContainer = addMarkedGroup(anchorLayer, SCRIPT_MARKER + "_AnchorShapes");
        }

        return outputContainers;
    }

    // =========================================
    // 画面ズーム / Zoom controls
    // =========================================

    /**
     * @typedef {object} ViewState
     * @property {View} view - 対象のビュー
     * @property {number} zoom - 倍率
     * @property {Array<number>} center - 中心の座標 [x, y]
     */

    /**
     * 現在の表示状態を控える
     * @param {Document} doc - 対象のドキュメント
     * @returns {ViewState} 表示状態
     */
    function captureViewState(doc) {
        var viewState = { view: null, zoom: null, center: null };
        /* ビューが取れないことがある / The view may be unavailable */
        try {
            viewState.view = doc.activeView;
            viewState.zoom = viewState.view.zoom;
            viewState.center = viewState.view.centerPoint;
        } catch (e) { logError(e, "captureViewState"); }
        return viewState;
    }

    /**
     * 控えておいた表示状態に戻す
     * @param {ViewState} viewState - 表示状態
     * @returns {void}
     */
    function restoreViewState(viewState) {
        if (!viewState || !viewState.view) return;
        /* ビューが閉じられていることがある / The view may be gone */
        try {
            if (viewState.zoom != null) viewState.view.zoom = viewState.zoom;
            if (viewState.center != null) viewState.view.centerPoint = viewState.center;
        } catch (e) { logError(e, "restoreViewState"); }
    }

    /**
     * ズームの中心を修飾キーに応じて移す（option でアートボード中心、shift で選択中心）
     * @param {Document} doc - 対象のドキュメント
     * @param {View} targetView - 対象のビュー
     * @returns {boolean} メニューコマンドでズームまで処理した場合は true
     */
    function moveZoomCenterByModifier(doc, targetView) {
        var keyboardState = ScriptUI.environment.keyboardState;
        if (!keyboardState) return false;

        var hasSelection = !!(doc.selection && doc.selection.length > 0);

        /* option + shift：選択をウィンドウにフィット / Fit the selection in the window */
        if (keyboardState.altKey && keyboardState.shiftKey && hasSelection) {
            try {
                app.executeMenuCommand("fitinwindow");
                return true;
            } catch (e) { logError(e, "moveZoomCenterByModifier/fitinwindow"); }
            return false;
        }

        /* option：アートボード中心 / Center on the artboard */
        if (keyboardState.altKey) {
            var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            targetView.centerPoint = [(artboardRect[0] + artboardRect[2]) / 2, (artboardRect[1] + artboardRect[3]) / 2];
            return false;
        }

        /* shift：選択中心 / Center on the selection */
        if (keyboardState.shiftKey && hasSelection) {
            var selectionBounds = getClipAwareUnionBounds(doc.selection, false);
            if (selectionBounds) {
                targetView.centerPoint = [(selectionBounds[0] + selectionBounds[2]) / 2, (selectionBounds[1] + selectionBounds[3]) / 2];
            }
        }

        return false;
    }

    /**
     * ズームのスライダーと軽量モードのチェックボックスを追加する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @param {Document} doc - 対象のドキュメント
     * @param {ViewState} initialState - 開いた時点の表示状態
     * @returns {{restoreInitial: Function}} 開いた時点の表示に戻す関数
     */
    function addZoomControls(parentWindow, doc, initialState) {
        var zoomRow = parentWindow.add("group");
        zoomRow.orientation = "row";
        zoomRow.alignChildren = ["center", "center"];
        zoomRow.alignment = "center";
        zoomRow.margins = ZOOM_ROW_MARGINS;

        zoomRow.add("statictext", undefined, labelText(LABELS.fieldLabel.zoom));

        var initialZoom = Number(initialState && initialState.zoom);
        if (!initialZoom || isNaN(initialZoom)) initialZoom = 1;

        var zoomSlider = zoomRow.add("slider", undefined, initialZoom, ZOOM_MIN, ZOOM_MAX);
        zoomSlider.preferredSize.width = ZOOM_SLIDER_WIDTH;
        zoomSlider.helpTip = getLabel(LABELS.tooltip.zoom);

        var lightModeCheckbox = addCheckbox(zoomRow, LABELS.checkbox.lightMode, false, LABELS.tooltip.lightMode);

        /**
         * スライダーの値を画面に反映する
         * @returns {void}
         */
        function applyZoom() {
            var targetView = (initialState && initialState.view) ? initialState.view : doc.activeView;
            if (!targetView) return;

            /* ビューが閉じられている・メニューコマンドが失敗するときは何もしない / Ignore a missing view or a failed menu command */
            try {
                if (moveZoomCenterByModifier(doc, targetView)) return;
                targetView.zoom = Number(zoomSlider.value);
                app.redraw();
            } catch (e) { logError(e, "addZoomControls/applyZoom"); }
        }

        /* 軽量モードではドラッグ中に再描画せず、離したときだけ反映する / Light mode applies on release only */
        zoomSlider.onChanging = function() {
            if (lightModeCheckbox.value) return;
            applyZoom();
        };
        zoomSlider.onChange = applyZoom;
        lightModeCheckbox.onClick = applyZoom;

        return {
            restoreInitial: function() { restoreViewState(initialState); }
        };
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * チェックボックスを追加する
     * @param {object} parentContainer - 追加先のウィンドウ・パネル・グループ
     * @param {object} labelSet - ja/en を持つラベル
     * @param {boolean} initialValue - 初期値
     * @param {object} [tooltipSet] - ja/en を持つ tooltip
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, labelSet, initialValue, tooltipSet) {
        var newCheckbox = parentContainer.add("checkbox", undefined, getLabel(labelSet));
        newCheckbox.value = initialValue;
        if (tooltipSet) newCheckbox.helpTip = getLabel(tooltipSet);
        return newCheckbox;
    }

    /**
     * ラジオボタンを追加する
     * @param {object} parentContainer - 追加先のパネル・グループ
     * @param {object} labelSet - ja/en を持つラベル
     * @param {object} [tooltipSet] - ja/en を持つ tooltip
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parentContainer, labelSet, tooltipSet) {
        var newRadio = parentContainer.add("radiobutton", undefined, getLabel(labelSet));
        if (tooltipSet) newRadio.helpTip = getLabel(tooltipSet);
        return newRadio;
    }

    /**
     * 長さの入力行（項目名・入力欄・単位）を追加する
     * @param {object} parentContainer - 追加先のパネル
     * @param {object} labelSet - 項目名のラベル
     * @param {string} prefKey - 単位を決める環境設定キー（strokeUnits / rulerType）
     * @param {number} initialPt - 初期値（pt）
     * @returns {{row: Group, input: EditText, stepper: Group}} 追加した行と入力欄、∧∨
     */
    function addLengthRow(parentContainer, labelSet, prefKey, initialPt) {
        var lengthRow = addRow(parentContainer);
        lengthRow.add("statictext", undefined, labelText(labelSet));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = lengthRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var lengthUnit = getUnitInfo(prefKey);
        var lengthInput;
        /* 長さはマイナスにしない。値を変えたら入力欄の onChanging（プレビュー更新）を呼ぶ
           Lengths stay non-negative; run the field's onChanging (preview refresh) after each step */
        var lengthStepper = addStepper(stepperInputGroup, function () { return lengthInput; }, {
            min: 0,
            onStep: function (numberInput) {
                if (typeof numberInput.onChanging === "function") numberInput.onChanging();
            }
        });
        lengthInput = stepperInputGroup.add("edittext", undefined, (initialPt / lengthUnit.pointsPerUnit).toFixed(3));
        lengthInput.characters = NUMBER_INPUT_CHARACTERS;
        bindSteppedArrowKeys(lengthInput, lengthStepper);
        lengthRow.add("statictext", undefined, lengthUnit.label);
        return { row: lengthRow, input: lengthInput, stepper: lengthStepper };
    }

    /**
     * 設定ダイアログを組み立てる（イベントは showDialog() で結線する）
     * @param {Document} doc - 対象のドキュメント
     * @returns {object} ダイアログと各コントロール
     */
    function buildDialog(doc) {
        var mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        // --- 2カラム / Two columns ---
        var columnsGroup = mainDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.alignment = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        // --- 左カラム：補助線を描画 / Left column: construction lines ---
        var linesPanel = addPanel(columnsGroup, LABELS.panel.lines);
        var straightCheckbox = addCheckbox(linesPanel, LABELS.checkbox.straight, true, LABELS.tooltip.straight);

        var directionPanel = addPanel(linesPanel);
        var horizontalCheckbox = addCheckbox(directionPanel, LABELS.checkbox.horizontal, true, LABELS.tooltip.soloDirection);
        var verticalCheckbox = addCheckbox(directionPanel, LABELS.checkbox.vertical, true, LABELS.tooltip.soloDirection);
        var diagonalCheckbox = addCheckbox(directionPanel, LABELS.checkbox.diagonal, true, LABELS.tooltip.soloDirection);

        var arcToCircleCheckbox = addCheckbox(linesPanel, LABELS.checkbox.arcToCircle, true, LABELS.tooltip.arcToCircle);

        var arcFallbackPanel = addPanel(linesPanel);
        arcFallbackPanel.helpTip = getLabel(LABELS.tooltip.arcFallback);
        var arcIgnoreRadio = addRadio(arcFallbackPanel, LABELS.radio.arcIgnore, LABELS.tooltip.arcIgnore);
        var arcChordRadio = addRadio(arcFallbackPanel, LABELS.radio.arcChord, LABELS.tooltip.arcChord);
        var arcExtendRadio = addRadio(arcFallbackPanel, LABELS.radio.arcExtendChord, LABELS.tooltip.arcExtendChord);
        arcIgnoreRadio.value = true;

        var strokeWidthCache = { value: mmToPt(DEFAULT_STROKE_WIDTH_MM) };
        var strokeWidthField = addLengthRow(linesPanel, LABELS.fieldLabel.strokeWidth, "strokeUnits", strokeWidthCache.value);
        strokeWidthField.row.margins = STROKE_ROW_MARGINS;

        // --- 右カラム / Right column ---
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.alignment = ["fill", "top"];
        rightColumn.spacing = COLUMN_SPACING;

        // --- アンカーポイントに図形 / Shapes on anchor points ---
        var anchorPanel = addPanel(rightColumn, LABELS.panel.anchorShapes);

        var anchorShapeRow = addRow(anchorPanel, RADIO_SPACING);
        var anchorNoneRadio = addRadio(anchorShapeRow, LABELS.radio.anchorNone);
        var anchorCircleRadio = addRadio(anchorShapeRow, LABELS.radio.anchorCircle);
        var anchorSquareRadio = addRadio(anchorShapeRow, LABELS.radio.anchorSquare);
        anchorNoneRadio.value = true;

        var anchorSizeCache = { value: mmToPt(DEFAULT_ANCHOR_SIZE_MM) };
        var anchorSizeField = addLengthRow(anchorPanel, LABELS.fieldLabel.anchorSize, "rulerType", anchorSizeCache.value);

        var anchorColorRow = addRow(anchorPanel, RADIO_SPACING);
        var anchorBlackRadio = addRadio(anchorColorRow, LABELS.radio.colorBlack);
        var anchorBlueRadio = addRadio(anchorColorRow, LABELS.radio.colorBlue);
        anchorBlackRadio.value = true;

        // --- オプション / Options ---
        var optionsPanel = addPanel(rightColumn, LABELS.panel.options);
        var groupCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.group, true, LABELS.tooltip.group);
        var separateLayerCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.separateLayer, true, LABELS.tooltip.separateLayer);
        var guideCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.guide, false, LABELS.tooltip.guide);
        var dedupCheckbox = addCheckbox(optionsPanel, LABELS.checkbox.dedup, true, LABELS.tooltip.dedup);

        // --- 画面ズーム / Zoom controls ---
        var zoomControls = addZoomControls(mainDialog, doc, captureViewState(doc));

        mainDialog.add("panel", undefined, undefined);

        // --- ボタンエリア / Button row ---
        var buttonRow = addButtonRow(mainDialog);
        var previewCheckbox = addCheckbox(buttonRow.leftGroup, LABELS.checkbox.preview, true);

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            dialog: mainDialog,
            straightCheckbox: straightCheckbox,
            directionPanel: directionPanel,
            horizontalCheckbox: horizontalCheckbox,
            verticalCheckbox: verticalCheckbox,
            diagonalCheckbox: diagonalCheckbox,
            arcToCircleCheckbox: arcToCircleCheckbox,
            arcFallbackPanel: arcFallbackPanel,
            arcIgnoreRadio: arcIgnoreRadio,
            arcChordRadio: arcChordRadio,
            arcExtendRadio: arcExtendRadio,
            strokeWidthInput: strokeWidthField.input,
            strokeWidthCache: strokeWidthCache,
            anchorNoneRadio: anchorNoneRadio,
            anchorCircleRadio: anchorCircleRadio,
            anchorSquareRadio: anchorSquareRadio,
            anchorSizeRow: anchorSizeField.row,
            anchorSizeInput: anchorSizeField.input,
            anchorSizeCache: anchorSizeCache,
            anchorColorRow: anchorColorRow,
            anchorBlackRadio: anchorBlackRadio,
            anchorBlueRadio: anchorBlueRadio,
            groupCheckbox: groupCheckbox,
            separateLayerCheckbox: separateLayerCheckbox,
            guideCheckbox: guideCheckbox,
            dedupCheckbox: dedupCheckbox,
            zoomControls: zoomControls,
            previewCheckbox: previewCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 設定ダイアログを表示して設定を取得する
     * @param {Document} doc - 対象のドキュメント
     * @param {Array} selectedItems - 選択アイテム
     * @param {Array<PathItem>} targetPaths - 対象のパス
     * @returns {object|null} ダイアログの設定（キャンセル時は null）
     */
    function showDialog(doc, selectedItems, targetPaths) {
        var dialogControls = buildDialog(doc);

        var mainDialog = dialogControls.dialog;
        var straightCheckbox = dialogControls.straightCheckbox;
        var horizontalCheckbox = dialogControls.horizontalCheckbox;
        var verticalCheckbox = dialogControls.verticalCheckbox;
        var diagonalCheckbox = dialogControls.diagonalCheckbox;
        var arcToCircleCheckbox = dialogControls.arcToCircleCheckbox;
        var arcChordRadio = dialogControls.arcChordRadio;
        var arcExtendRadio = dialogControls.arcExtendRadio;
        var anchorNoneRadio = dialogControls.anchorNoneRadio;
        var anchorCircleRadio = dialogControls.anchorCircleRadio;
        var anchorSquareRadio = dialogControls.anchorSquareRadio;
        var anchorBlueRadio = dialogControls.anchorBlueRadio;
        var groupCheckbox = dialogControls.groupCheckbox;
        var separateLayerCheckbox = dialogControls.separateLayerCheckbox;
        var guideCheckbox = dialogControls.guideCheckbox;
        var dedupCheckbox = dialogControls.dedupCheckbox;
        var previewCheckbox = dialogControls.previewCheckbox;
        var zoomControls = dialogControls.zoomControls;

        // =========================================
        // ダイアログの状態 / Dialog behavior
        // =========================================

        var previewLayerName = createUniqueLayerName(doc, PREVIEW_LAYER_BASE_NAME);
        var isAccepted = false;

        /**
         * ダイアログの入力内容を設定オブジェクトにまとめる
         * @returns {object} ダイアログの設定
         */
        function getUISettings() {
            return {
                straightLines: straightCheckbox.value,
                horizontal: horizontalCheckbox.value,
                vertical: verticalCheckbox.value,
                diagonal: diagonalCheckbox.value,
                arcToCircle: arcToCircleCheckbox.value,
                arcFallback: arcChordRadio.value ? ARC_FALLBACK.CHORD :
                    (arcExtendRadio.value ? ARC_FALLBACK.EXTEND : ARC_FALLBACK.IGNORE),
                group: groupCheckbox.value,
                separateLayer: separateLayerCheckbox.value,
                guide: guideCheckbox.value,
                dedup: dedupCheckbox.value,
                anchorShape: anchorCircleRadio.value ? ANCHOR_SHAPE.CIRCLE :
                    (anchorSquareRadio.value ? ANCHOR_SHAPE.SQUARE : ANCHOR_SHAPE.NONE),
                anchorColor: anchorBlueRadio.value ? ANCHOR_COLOR.BLUE : ANCHOR_COLOR.BLACK,
                anchorSizePt: readLengthAsPt(dialogControls.anchorSizeInput, "rulerType", dialogControls.anchorSizeCache),
                strokeWidthPt: readLengthAsPt(dialogControls.strokeWidthInput, "strokeUnits", dialogControls.strokeWidthCache)
            };
        }

        /**
         * 選べない組み合わせのコントロールをディムする
         * @returns {void}
         */
        function updateEnabledState() {
            dialogControls.directionPanel.enabled = straightCheckbox.value;
            dialogControls.arcFallbackPanel.enabled = arcToCircleCheckbox.value;

            /* 線を1本も描かないなら、ガイド化とダブり削除は効かない / Both options are moot without any line */
            var drawsLines = (straightCheckbox.value || arcToCircleCheckbox.value);
            guideCheckbox.enabled = drawsLines;
            dedupCheckbox.enabled = drawsLines;

            var drawsAnchors = !anchorNoneRadio.value;
            dialogControls.anchorSizeRow.enabled = drawsAnchors;
            dialogControls.anchorColorRow.enabled = drawsAnchors;
            redrawSteppersIn(dialogControls.anchorSizeRow); /* 自作描画の∧∨を描き直す / redraw the custom-drawn stepper */
        }

        /**
         * プレビューを消す
         * @returns {void}
         */
        function clearPreview() {
            removeLayerIfExists(doc, previewLayerName);
            app.redraw();
        }

        /**
         * 現在の設定でプレビューを描き直す（ユーザーのオブジェクトには触らない）
         * @returns {void}
         */
        function refreshPreview() {
            if (!previewCheckbox.value) return;

            removeLayerIfExists(doc, previewLayerName);

            var previewLayer = doc.layers.add();
            previewLayer.name = previewLayerName;
            previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT);

            var previewGroup = addMarkedGroup(previewLayer, SCRIPT_MARKER + "_PREVIEW");
            var settings = getUISettings();

            /* 本処理と同じグループ構造でプレビューする / Mirror the grouping used by the real run */
            var lineContainer = previewGroup;
            var anchorContainer = previewGroup;
            if (settings.group) {
                lineContainer = addMarkedGroup(previewGroup, SCRIPT_MARKER + "_" + getLabel(LABELS.itemName.lineGroup));
                anchorContainer = addMarkedGroup(previewGroup, SCRIPT_MARKER + "_AnchorShapes");
            }

            drawConstructionItems(targetPaths, settings, getDrawArea(doc, selectedItems), lineContainer, anchorContainer);
            app.redraw();
        }

        /**
         * 設定変更のハンドラーを作る（ディム状態を更新してプレビューを描き直す）
         * @param {function} [beforeRefresh] - プレビュー更新の前に行う処理
         * @returns {function} onClick に割り当てるハンドラー
         */
        function onSettingChanged(beforeRefresh) {
            return function() {
                if (typeof beforeRefresh === "function") beforeRefresh();
                updateEnabledState();
                refreshPreview();
            };
        }

        /**
         * option + クリックで、その向きだけを残す
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @param {Checkbox} otherA - もう一方のチェックボックス
         * @param {Checkbox} otherB - さらにもう一方のチェックボックス
         * @returns {void}
         */
        function soloDirectionOnAltClick(clickedCheckbox, otherA, otherB) {
            var keyboardState = ScriptUI.environment.keyboardState;
            if (!keyboardState || !keyboardState.altKey) return;
            if (!clickedCheckbox.value) return;

            otherA.value = false;
            otherB.value = false;
        }

        var settingControls = [
            straightCheckbox, arcToCircleCheckbox, groupCheckbox, separateLayerCheckbox, guideCheckbox, dedupCheckbox,
            dialogControls.arcIgnoreRadio, arcChordRadio, arcExtendRadio,
            anchorNoneRadio, anchorCircleRadio, anchorSquareRadio, dialogControls.anchorBlackRadio, anchorBlueRadio
        ];
        for (var i = 0; i < settingControls.length; i++) {
            settingControls[i].onClick = onSettingChanged();
        }

        horizontalCheckbox.onClick = onSettingChanged(function() {
            soloDirectionOnAltClick(horizontalCheckbox, verticalCheckbox, diagonalCheckbox);
        });
        verticalCheckbox.onClick = onSettingChanged(function() {
            soloDirectionOnAltClick(verticalCheckbox, horizontalCheckbox, diagonalCheckbox);
        });
        diagonalCheckbox.onClick = onSettingChanged(function() {
            soloDirectionOnAltClick(diagonalCheckbox, horizontalCheckbox, verticalCheckbox);
        });

        previewCheckbox.onClick = function() {
            if (previewCheckbox.value) refreshPreview();
            else clearPreview();
        };

        dialogControls.strokeWidthInput.onChanging = refreshPreview;
        dialogControls.anchorSizeInput.onChanging = refreshPreview;

        dialogControls.btnOK.onClick = function() {
            isAccepted = true;
            mainDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function() {
            zoomControls.restoreInitial();
            mainDialog.close(0);
        };

        /* タイトルバーの×もキャンセルと同じ扱いにする / Closing from the title bar cancels too */
        mainDialog.onClose = function() {
            if (!isAccepted) zoomControls.restoreInitial();
            return true;
        };

        updateEnabledState();
        /* 既定でONなので、ダイアログを開く前に描いておく / Preview is on by default, so draw it up front */
        refreshPreview();
        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        var dialogResult = mainDialog.show();

        clearPreview();

        return (dialogResult === 1) ? getUISettings() : null;
    }

    // =========================================
    // エントリーポイント / Entry point
    // =========================================

    /**
     * 補助線を描く本体の処理（suspendHistory から呼ばれる）
     * @returns {void}
     */
    function mainImpl() {
        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var targetPaths = collectSelectionPathItems(selectedItems, { unique: false });

        /* テキストは複製をアウトライン化してからパスを拾う / Outline a copy of the text to read its paths */
        var outlineRoots = outlineTextFromSelection(selectedItems);
        if (outlineRoots.length > 0) targetPaths = targetPaths.concat(collectSelectionPathItems(outlineRoots, { unique: false }));

        try {
            if (targetPaths.length === 0) {
                alert(getLabel(LABELS.alert.noValidPath));
                return;
            }

            /* プレビューがレイヤーを増減するので、先にアクティブレイヤーを控える / Preview adds and removes layers */
            var activeLayer = doc.activeLayer;

            var settings = showDialog(doc, selectedItems, targetPaths);
            if (!settings) return;

            var containers = prepareOutputContainers(doc, selectedItems, activeLayer, settings);
            drawConstructionItems(
                targetPaths,
                settings,
                getDrawArea(doc, selectedItems),
                containers.lineContainer,
                containers.anchorContainer
            );
        } finally {
            /* 失敗しても一時アウトラインは必ず片付ける / Always clean up the temporary outlines */
            cleanupTempOutlines(outlineRoots);
        }
    }

    /**
     * ドキュメントを確認して本体を実行する（suspendHistory が使える環境では取り消し1ステップにまとめる）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        try {
            var doc = app.activeDocument;
            if (typeof doc.suspendHistory === "function") {
                doc.suspendHistory(getLabel(LABELS.itemName.history), "mainImpl()");
            } else {
                mainImpl();
            }
        } catch (e) {
            logError(e, "main");
            alert(getLabel(LABELS.alert.error) + e);
        }
    }

    main();

})();
