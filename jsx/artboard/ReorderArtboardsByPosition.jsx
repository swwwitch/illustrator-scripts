#target illustrator
#targetengine "ReorderArtboardsByPositionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

カンバス上の見た目の並び（左上基準・Y優先）をもとに、［アートボード］パネルの並び順を整理します。
列数指定やアートボード名にもとづくカンバス上の再配置、「行-列」形式へのリネームも実行できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReorderArtboardsByPosition.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb416cb01728a

### Overview

Reorders the Artboards panel to match the visual arrangement on the canvas, working from the top-left with Y taking priority.
It can also rearrange the artboards on the canvas by column count or by name, and rename them into a "row-column" form.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReorderArtboardsByPosition.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReorderArtboardsByPosition";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-11-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReorderArtboardsByPosition.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReorderArtboardsByPosition.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb416cb01728a"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 座標比較に使う丸め桁数 / Rounding digits used when comparing coordinates */
    var COORDINATE_PRECISION_DIGITS = 3;

    /* 許容差の自動計算に使う係数と下限 / Factors used to auto-calculate the row tolerance */
    var TOLERANCE_AUTO_MARGIN_RATIO  = 1.1; /* 自動計算値に掛ける倍率 / multiplier applied to the auto value */
    var TOLERANCE_AUTO_MARGIN_POINTS = 2;   /* 最小差に加えるマージン（pt） / margin added to the smallest gap (pt) */
    var TOLERANCE_FALLBACK_POINTS    = 5;   /* 上辺に差がないときの既定値（pt） / fallback when all top edges match (pt) */
    var TOLERANCE_MAX_HEIGHT_RATIO   = 0.5; /* 初期値の上限（最も低いアートボードの高さに対する比） / cap relative to the shortest artboard height */

    /* 再配置ダイアログの初期値と下限 / Initial values and minimums for the rearrange settings */
    var DEFAULT_COLUMN_COUNT = 4;   /* 列数 / column count */
    var MIN_COLUMN_COUNT     = 1;   /* 列数の下限 / minimum column count */
    var DEFAULT_GAP_VALUE    = 100; /* 列間・行間（表示単位） / column and row gap in the display unit */
    var MIN_GAP_VALUE        = 0;   /* 列間・行間の下限 / minimum gap value */

    /* byName 再配置で例外領域（未指定／重複）の境界に適用するギャップ倍率
     * Gap multiplier applied at the boundary into the unspecified/duplicate exception area in byName mode. */
    var EXCEPTION_BOUNDARY_GAP_MULTIPLIER = 3;

    /* カンバスの一辺（pt）。再配置後の位置をカンバス内に収める判定に使う / Canvas side length (pt), used to keep the layout on the canvas */
    var CANVAS_SIZE_POINTS = 16383;

    /* 複製対象のアートボードプロパティ / Artboard properties to copy when duplicating */
    var ARTBOARD_COPYABLE_PROPS = ["name", "rulerOrigin", "rulerPAR", "showCenter", "showCrossHairs", "showSafeAreas"];

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

    /* コントロールの寸法 / Control metrics */
    var FIELD_LABEL_WIDTH   = 50;  /* 入力欄ラベルの幅 / labeled field label width */
    var FIELD_INPUT_CHARS   = 5;   /* 入力欄の文字数 / edittext width in characters */
    var SLIDER_WIDTH        = 250; /* 許容差スライダーの幅 / tolerance slider width */
    var SLIDER_HEIGHT       = 20;  /* 許容差スライダーの高さ / tolerance slider height */
    var PREVIEW_LIST_WIDTH  = 160; /* 並び順リストの幅 / reorder list width */
    var PREVIEW_LIST_HEIGHT = 120; /* 並び順リストの最小高さ / reorder list minimum height */

    /**
     * 列グループの共通設定を適用する
     * @param {Group} columnGroup - 対象グループ
     * @param {string|Array} [alignChildren] - 子要素の配置（省略時は ["fill", "top"]）
     * @param {string|Array} [alignment] - グループ自身の配置（省略時は ["fill", "top"]）
     * @returns {void}
     */
    function setupColumn(columnGroup, alignChildren, alignment) {
        columnGroup.orientation = "column";
        columnGroup.alignChildren = alignChildren || ["fill", "top"];
        columnGroup.alignment = alignment || ["fill", "top"];
    }

    /**
     * 2カラムを横に並べる行グループの共通設定を適用する
     * @param {Group} columnsRowGroup - 対象グループ
     * @returns {void}
     */
    function setupColumnsRow(columnsRowGroup) {
        columnsRowGroup.orientation = "row";
        columnsRowGroup.alignChildren = ["fill", "top"];
        columnsRowGroup.alignment = ["fill", "top"];
        columnsRowGroup.spacing = COLUMN_SPACING;
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
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle) {
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = LINK_ICON_SIZE;
        linkToggle.minimumSize = LINK_ICON_SIZE;
        linkToggle.maximumSize = LINK_ICON_SIZE;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = LINK_ICON_SIZE[0];
            var iconHeight = LINK_ICON_SIZE[1];
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
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
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
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22;
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードの並び順を整理", en: "Organize Artboard Order" }
        },
        panel: {
            reorder:          { ja: "パネル上の並び順", en: "Artboards Panel Order" },
            rearrange:        { ja: "カンバス上のアートボードを再配置", en: "Rearrange Canvas Artboards" },
            duplicateHandling: { ja: "未指定／重複", en: "Unspecified / Duplicate" },
            naming:           { ja: "アートボード名の更新", en: "Update names as row-column" },
            namingSeparator:  { ja: "区切り文字", en: "Separator" },
            namingPadWidth:   { ja: "桁数", en: "Digits" }
        },
        radio: {
            sortByName:               { ja: "名前順", en: "By name" },
            sortByPosition:           { ja: "カンバス上の並び順に", en: "Match canvas order" },
            sortKeepAsIs:             { ja: "変更しない", en: "Keep as is" },
            duplicateAppendToRowEnd:  { ja: "各行の末尾に配置", en: "Append to each row end" },
            duplicateGroupInLastRow:  { ja: "最終行の次の行にまとめる", en: "Group in row after last row" },
            namingFromPosition:       { ja: "配置位置から作成", en: "Create from position" },
            namingFromExisting:       { ja: "既存名を整形", en: "Reformat existing names" }
        },
        checkbox: {
            rearrangeByColumns: { ja: "列数を指定して再配置", en: "Rearrange by column count" },
            rearrangeByName:    { ja: "アートボード名から行列に再配置", en: "Rearrange by row-column from names" },
            namingEnable:       { ja: "「行-列」形式に更新", en: "Update" }
        },
        fieldLabel: {
            columns:   { ja: "列数", en: "Columns" },
            columnGap: { ja: "列間", en: "Column gap" },
            rowGap:    { ja: "行間", en: "Row gap" }
        },
        tooltip: {
            sortByName:     { ja: "アートボードパネルの並び順を、アートボード名の昇順に整えます。", en: "Sorts the Artboards panel by artboard name." },
            sortByPosition: { ja: "アートボードパネルの並び順を、カンバス上の配置（左上から右下）に合わせます。", en: "Sorts the Artboards panel to match the canvas layout, top-left to bottom-right." },
            sortKeepAsIs:   { ja: "アートボードパネルの並び順は変更しません。", en: "Leaves the Artboards panel order untouched." },
            tolerance: {
                ja: "上下のずれがこの値以内なら同じ行とみなします。行の切れ目が合わないときに調整します。",
                en: "Artboards within this vertical distance count as one row. Adjust it when the rows come out wrong."
            },
            rearrangeByColumns: { ja: "指定した列数で折り返しながら、カンバス上のアートボードを並べ直します。", en: "Rearranges the artboards on the canvas, wrapping at the given column count." },
            rearrangeByName:    { ja: "アートボード名に含まれる「行-列」を読み取って、その位置へ並べ直します。", en: "Reads the row-column part of each artboard name and places it accordingly." },
            gapLink:            { ja: "列間と同じ値を行間にも使います。", en: "Uses the column gap for the row gap too." },
            duplicateAppendToRowEnd: { ja: "行列を読み取れなかったアートボードや重複したものを、それぞれの行の末尾に置きます。", en: "Puts artboards with no readable row-column, or duplicates, at the end of each row." },
            duplicateGroupInLastRow: { ja: "行列を読み取れなかったアートボードや重複したものを、最終行の次の行にまとめます。", en: "Collects artboards with no readable row-column, or duplicates, into a row after the last one." },
            namingEnable:       { ja: "処理のあと、アートボード名を「行-列」形式に付け直します。", en: "Renames the artboards as row-column once the rearranging is done." },
            namingFromPosition: { ja: "並べ直したあとの位置から、新しい名前を作ります。", en: "Builds the new names from the positions after rearranging." },
            namingFromExisting: { ja: "既存の名前に含まれる行列を読み取り、区切り文字と桁数だけ整えます。", en: "Keeps the row-column found in the existing names and only fixes the separator and digits." },
            namingSeparator:    { ja: "行番号と列番号のあいだに入れる文字です。", en: "Character placed between the row number and the column number." },
            namingPadWidth:     { ja: "行番号・列番号がこの桁数になるまで、先頭に 0 を補います（例: 00 → 01-02）。", en: "Pads the row and column numbers with leading zeros to this many digits (e.g. 00 gives 01-02)." },
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
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noArtboards: { ja: "アートボードが存在しません。", en: "No artboards found." },
            errorPrefix: { ja: "エラー: ", en: "Error: " },
            noMatch: {
                ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）形式のアートボード名が見つかりませんでした。",
                en: "No artboard names in '<row><sep><column>' (e.g., 1-2) or '<prefix><sep><number>' (e.g., banner-1) format were found."
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

    var currentUnitLabel = getUnitInfo().label;

    /**
     * 入力欄の表示単位値を読み取り、ポイントに変換する
     * @param {EditText} valueInput - 対象の入力欄
     * @param {number} defaultValue - 数値として読めない場合の既定値（表示単位）
     * @returns {number} ポイント値
     */
    function readDisplayUnitInputAsPoints(valueInput, defaultValue) {
        var displayValue = parseFloat(valueInput.text);
        if (isNaN(displayValue)) displayValue = defaultValue;
        return displayValue * getUnitInfo().pointsPerUnit;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * @typedef {Object} ArtboardEntry
     * @property {string} name - アートボード名
     * @property {number[]} artboardRect - [左, 上, 右, 下]
     * @property {number} [sourceIndex] - 元のアートボードインデックス
     * @property {number} [rowBand] - 上辺の近さで割り当てた 0 始まりの行番号
     */

    /**
     * @typedef {Object} PreviewContext
     * @property {ArtboardEntry[]} artboardEntries - プレビュー用のアートボード情報
     * @property {number} sliderMax - 許容差スライダーの最大値
     * @property {number} defaultTolerance - 許容差スライダーの初期値
     */

    /**
     * ドキュメントを検証してダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (doc.artboards.length === 0) {
            alert(getLabel("alert.noArtboards"));
            return;
        }

        /* 座標はドキュメント基準で読み書きする。アートボード基準だと原点が作業中のアートボードに付いて動き、
         * 再配置の前後で座標が食い違う（カンバスの外への代入で 'CoOA' エラー）。設定はアプリ全体に残るので最後に戻す
         * Work in document coordinates: with artboard coordinates the origin follows the active artboard, so
         * rects read before and after the rearrange disagree. The setting is app-wide, so restore it at the end */
        var savedCoordinateSystem = app.coordinateSystem;
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        try {
            var previewContext = buildPreviewContext(doc);
            var dialogUI = buildDialogUI(previewContext.defaultTolerance, previewContext.sliderMax);

            bindEvents(doc, previewContext, dialogUI);

            prepareDialogWindow(dialogUI.reorderDialog, SCRIPT_NAME);
            dialogUI.reorderDialog.show();
        } finally {
            app.coordinateSystem = savedCoordinateSystem;
        }
    }

    /**
     * プレビューと許容差の初期値を組み立てる
     * @param {Document} doc - 対象ドキュメント
     * @returns {PreviewContext} プレビュー用のコンテキスト
     */
    function buildPreviewContext(doc) {
        /* アートボード情報を配列に格納（プレビュー・自動計算用） / Snapshot artboards for preview and auto tolerance */
        var artboardEntries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            artboardEntries.push({
                name: doc.artboards[artboardIndex].name,
                artboardRect: doc.artboards[artboardIndex].artboardRect
            });
        }

        /* スライダー最大値をアートボードの最大高さに設定 / Use the tallest artboard as the slider maximum */
        var maxHeight = 0;
        var minHeight = Infinity;
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            var artboardRect = artboardEntries[entryIndex].artboardRect;
            var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);
            if (artboardHeight > maxHeight) maxHeight = artboardHeight;
            if (artboardHeight < minHeight) minHeight = artboardHeight;
        }
        var sliderMax = Math.round(maxHeight);

        /* スライダー初期値を自動計算値にマージンを掛けて設定 / Seed the slider with the auto value plus a margin */
        var autoTolerance = calculateAutoTolerance(artboardEntries);
        var defaultTolerance = Math.round(autoTolerance * TOLERANCE_AUTO_MARGIN_RATIO);
        /* 上辺がそろった並びでは最小の差が行の間隔そのものになり、行がまとまってしまう。
         * 重ならない行の間隔は上の行の高さ以上なので、最も低いアートボードの高さの半分で止める
         * With aligned rows the smallest gap is the row pitch itself; rows that do not overlap are at least
         * one artboard height apart, so cap at half the shortest height to keep them separate */
        defaultTolerance = Math.min(defaultTolerance, Math.floor(minHeight * TOLERANCE_MAX_HEIGHT_RATIO), sliderMax);

        return {
            artboardEntries: artboardEntries,
            sliderMax: sliderMax,
            defaultTolerance: defaultTolerance
        };
    }

    /**
     * 許容差を自動計算する（上辺の最小差 + マージン）
     * @param {ArtboardEntry[]} artboardEntries - アートボード情報
     * @returns {number} 許容差（pt）
     */
    function calculateAutoTolerance(artboardEntries) {
        var topEdges = [];
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            topEdges.push(artboardEntries[entryIndex].artboardRect[1]);
        }
        /* 上から下に並べる / Sort top to bottom */
        topEdges.sort(function (firstTop, secondTop) {
            return secondTop - firstTop;
        });

        var topGaps = [];
        for (var topIndex = 1; topIndex < topEdges.length; topIndex++) {
            var topGap = Math.abs(topEdges[topIndex] - topEdges[topIndex - 1]);
            if (topGap > 0) topGaps.push(topGap);
        }

        /* 差が無ければデフォルト / Default if no difference */
        if (topGaps.length === 0) return TOLERANCE_FALLBACK_POINTS;

        /* 少しマージンを加える / Add some margin */
        return Math.min.apply(null, topGaps) + TOLERANCE_AUTO_MARGIN_POINTS;
    }

    /**
     * ダイアログのイベントを設定する
     * @param {Document} doc - 対象ドキュメント
     * @param {PreviewContext} previewContext - プレビュー用のコンテキスト
     * @param {Object} dialogUI - buildDialogUI() が返すUI参照
     * @returns {void}
     */
    function bindEvents(doc, previewContext, dialogUI) {
        var sortRadios = dialogUI.preview.sortModeRadios;

        /* ラジオは親グループが異なるため、ScriptUI の自動排他が効かない。
         * 手動で他のラジオを OFF にして相互排他を成立させる。
         * Radios live in separate parent groups, so ScriptUI's built-in exclusion
         * does not apply. Toggle the others off manually to enforce exclusivity. */
        /**
         * 並び順モードのラジオボタンを排他的に切り替え、関連するUIを同期する
         * @param {RadioButton} targetRadio - 選択状態にするラジオボタン
         * @returns {void}
         */
        function selectSortMode(targetRadio) {
            sortRadios.byName.value = (targetRadio === sortRadios.byName);
            sortRadios.byPosition.value = (targetRadio === sortRadios.byPosition);
            sortRadios.keepAsIs.value = (targetRadio === sortRadios.keepAsIs);
            syncSortMode();
        }

        /* 許容差はパネル並び順と「配置位置から作成」の両方で使うため、どちらかが有効なら操作できるようにする
         * The tolerance drives both the panel order and position-based naming, so enable it for either */
        /**
         * 許容差スライダーの有効／無効を、パネル並び順と「配置位置から作成」の状態に合わせる
         * @returns {void}
         */
        function syncToleranceEnabled() {
            var usedByPanelOrder = sortRadios.byPosition.value;
            var usedByNaming = dialogUI.naming.enableCheckbox.value && dialogUI.naming.source.fromPosition.value;
            dialogUI.preview.toleranceSlider.enabled = usedByPanelOrder || usedByNaming;
        }

        /**
         * 並び順モードの変更をUIとプレビューに反映する
         * @returns {void}
         */
        function syncSortMode() {
            syncToleranceEnabled();
            updateReorderPreview(previewContext, dialogUI, Math.round(dialogUI.preview.toleranceSlider.value));
        }

        /**
         * リネーム設定グループの有効／無効を、リネームのON/OFFに合わせる
         * @returns {void}
         */
        function syncNamingState() {
            dialogUI.naming.settingsGroup.enabled = dialogUI.naming.enableCheckbox.value;
            syncToleranceEnabled();
        }

        syncNamingState();
        syncSortMode();

        dialogUI.preview.toleranceSlider.onChanging = function () {
            updateReorderPreview(previewContext, dialogUI, Math.round(dialogUI.preview.toleranceSlider.value));
        };

        sortRadios.byPosition.onClick = function () { selectSortMode(sortRadios.byPosition); };
        sortRadios.byName.onClick = function () { selectSortMode(sortRadios.byName); };
        sortRadios.keepAsIs.onClick = function () { selectSortMode(sortRadios.keepAsIs); };

        dialogUI.naming.enableCheckbox.onClick = syncNamingState;
        dialogUI.naming.source.fromPosition.onClick = syncToleranceEnabled;
        dialogUI.naming.source.fromExisting.onClick = syncToleranceEnabled;

        /* 再配置の設定が変わったら、再配置後の並びでプレビューを出し直す（既存のハンドラーのあとに呼ぶ）
         * Refresh the preview when the rearrange settings change, after the existing handlers */
        var rearrangeUI = dialogUI.rearrange;
        appendHandler(rearrangeUI.modeChecks.byColumns, "onClick", syncSortMode);
        appendHandler(rearrangeUI.modeChecks.byName, "onClick", syncSortMode);
        appendHandler(rearrangeUI.columnsInput, "onChange", syncSortMode);
        appendHandler(rearrangeUI.duplicateRadios.append, "onClick", syncSortMode);
        appendHandler(rearrangeUI.duplicateRadios.groupLast, "onClick", syncSortMode);

        dialogUI.buttons.btnOK.onClick = function () {
            executeReorder(doc, dialogUI);
        };

        dialogUI.buttons.btnCancel.onClick = function () {
            dialogUI.reorderDialog.close(-1);
        };
    }

    /**
     * コントロールの既存のハンドラーのあとに、別の処理を足す
     * @param {Object} control - 対象のコントロール
     * @param {string} handlerName - "onClick" / "onChange" など
     * @param {Function} extraHandler - 足す処理
     * @returns {void}
     */
    function appendHandler(control, handlerName, extraHandler) {
        var previousHandler = control[handlerName];
        control[handlerName] = function () {
            if (previousHandler) previousHandler.apply(this, arguments);
            extraHandler();
        };
    }

    /**
     * 並び順リストのプレビューを更新する
     * @param {PreviewContext} previewContext - プレビュー用のコンテキスト
     * @param {Object} dialogUI - buildDialogUI() が返すUI参照
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
    function updateReorderPreview(previewContext, dialogUI, tolerance) {
        var reorderList = dialogUI.preview.reorderList;
        reorderList.removeAll();

        /* 変更しないモード: 現在の並びをそのまま表示 / Keep-as-is mode: show current order untouched */
        if (dialogUI.preview.sortModeRadios.keepAsIs.value) {
            for (var keepIndex = 0; keepIndex < previewContext.artboardEntries.length; keepIndex++) {
                reorderList.add("item", previewContext.artboardEntries[keepIndex].name);
            }
            return;
        }

        /* 名前順モード: アートボード名で並べてリスト表示 / By-name mode: list names in name-sort order */
        if (dialogUI.preview.sortModeRadios.byName.value) {
            var byNameEntries = previewContext.artboardEntries.slice();
            sortArtboardsByName(byNameEntries);
            for (var nameIndex = 0; nameIndex < byNameEntries.length; nameIndex++) {
                reorderList.add("item", byNameEntries[nameIndex].name);
            }
            return;
        }

        /* 位置順モード: 行ごとにまとめて表示。再配置がオンなら再配置後の並びを出す
         * Position mode: show one row per line, as it will be after the rearrange when that is on */
        var rowGroups = buildRearrangedPreviewRows(previewContext, dialogUI);
        if (!rowGroups) {
            var decimalPlaces = Math.pow(10, COORDINATE_PRECISION_DIGITS);
            var sortedEntries = previewContext.artboardEntries.slice();
            sortArtboardsTopLeftWithTolerance(sortedEntries, decimalPlaces, tolerance);
            rowGroups = groupSortedIntoRows(sortedEntries);
        }
        for (var rowIndex = 0; rowIndex < rowGroups.length; rowIndex++) {
            var rowNames = [];
            for (var columnIndex = 0; columnIndex < rowGroups[rowIndex].length; columnIndex++) {
                rowNames.push(rowGroups[rowIndex][columnIndex].name);
            }
            reorderList.add("item", rowNames.join(" | "));
        }
    }

    /**
     * 再配置後のカンバス上の行を、ドキュメントに触れずに組み立てる
     * @param {PreviewContext} previewContext - プレビュー用のコンテキスト
     * @param {Object} dialogUI - buildDialogUI() が返すUI参照
     * @returns {Array<ArtboardEntry[]>|null} 行ごとにまとめた配列。再配置しない・名前が解析できないときは null
     */
    function buildRearrangedPreviewRows(previewContext, dialogUI) {
        var modeChecks = dialogUI.rearrange.modeChecks;
        var artboardEntries = previewContext.artboardEntries;
        var rowGroups = [];

        /* 列数指定: パネル順に列数ずつ折り返す / By columns: wrap the panel order at the column count */
        if (modeChecks.byColumns.value) {
            var columns = parseColumnCount(dialogUI.rearrange.columnsInput.text);
            for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
                if (entryIndex % columns === 0) rowGroups.push([]);
                rowGroups[rowGroups.length - 1].push(artboardEntries[entryIndex]);
            }
            return rowGroups;
        }
        if (!modeChecks.byName.value) return null;

        /* アートボード名: 実行時と同じ割り当てで行と列を決める / By name: same slot assignment as the real run */
        var placementContext = parseArtboardNamePlacements(artboardEntries);
        if (!placementContext.hasMatchedArtboard) return null;
        var exceptionMode = dialogUI.rearrange.duplicateRadios.groupLast.value ? 'lastRow' : 'rowEnd';
        assignArtboardPlacementSlots(placementContext, exceptionMode);

        var placementsByRow = {};
        var placementItems = placementContext.placementItems;
        for (var itemIndex = 0; itemIndex < placementItems.length; itemIndex++) {
            var rowKey = placementItems[itemIndex].assignedRow;
            if (!placementsByRow[rowKey]) placementsByRow[rowKey] = [];
            placementsByRow[rowKey].push(placementItems[itemIndex]);
        }
        var rowNumbers = collectSortedNumericKeys(placementsByRow);
        for (var rowIndex = 0; rowIndex < rowNumbers.length; rowIndex++) {
            var rowPlacements = placementsByRow[rowNumbers[rowIndex]];
            rowPlacements.sort(function (firstItem, secondItem) {
                return firstItem.assignedColumn - secondItem.assignedColumn;
            });
            var rowEntries = [];
            for (var columnIndex = 0; columnIndex < rowPlacements.length; columnIndex++) {
                rowEntries.push(rowPlacements[columnIndex].artboard);
            }
            rowGroups.push(rowEntries);
        }
        return rowGroups;
    }

    /**
     * ダイアログの設定を読み取り、再配置・並べ替え・リネームを実行する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} dialogUI - buildDialogUI() が返すUI参照
     * @returns {void}
     */
    function executeReorder(doc, dialogUI) {
        var tolerance = Math.round(dialogUI.preview.toleranceSlider.value) || 0;

        /* 再配置に失敗したらアラート済みなので、ダイアログを開いたまま中断する
         * Abort with the dialog still open when the rearrange failed (already alerted) */
        var rearrangeResult = applyCanvasRearrange(doc, dialogUI);
        if (!rearrangeResult.ok) return;

        /* 再配置は済んでいるので、ここで失敗してもダイアログは閉じる（開いたままだと再実行で二重に動かしてしまう）
         * The rearrange is already applied, so close even on failure; a second OK would move everything again */
        try {
            applyPanelReorder(doc, dialogUI, tolerance);
            applyArtboardRenaming(doc, dialogUI, tolerance);
        } catch (reorderError) {
            alert(getLabel("alert.errorPrefix") + reorderError.message);
        }

        dialogUI.reorderDialog.close(1);

        /* ダイアログを閉じてからビューを合わせる（モーダル表示中のメニュー実行を避ける）
         * Fit the view after closing, so no menu command runs while the modal dialog is up */
        if (rearrangeResult.rearranged) app.executeMenuCommand('fitall');
    }

    /**
     * カンバス上の再配置を実行する
     * 再配置を先に実行すると、パネル順が「再配置後の見た目」と揃う
     * Run rearrange first so the panel order reflects post-rearrange positions
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} dialogUI - buildDialogUI() が返すUI参照
     * @returns {{ok: boolean, rearranged: boolean}} ok は後続処理を続けてよいか、rearranged は実際に動かしたか
     */
    function applyCanvasRearrange(doc, dialogUI) {
        var modeChecks = dialogUI.rearrange.modeChecks;
        if (!modeChecks.byColumns.value && !modeChecks.byName.value) {
            return { ok: true, rearranged: false };
        }

        /* 入力値は表示単位として受け取り、ptに変換して渡す / Read inputs in display units and convert to points */
        var columnGapPoints = readDisplayUnitInputAsPoints(dialogUI.rearrange.columnGapInput, DEFAULT_GAP_VALUE);
        var rowGapPoints = dialogUI.rearrange.gapLinkToggle.value
            ? columnGapPoints
            : readDisplayUnitInputAsPoints(dialogUI.rearrange.rowGapInput, DEFAULT_GAP_VALUE);

        try {
            if (modeChecks.byColumns.value) {
                rearrangeArtboardsWithGaps(doc, readColumnCount(dialogUI.rearrange.columnsInput), columnGapPoints, rowGapPoints);
                return { ok: true, rearranged: true };
            }
            var exceptionMode = dialogUI.rearrange.duplicateRadios.groupLast.value ? 'lastRow' : 'rowEnd';
            var matched = rearrangeArtboardsByRowColumnName(doc, columnGapPoints, rowGapPoints, exceptionMode);
            return { ok: matched, rearranged: matched };
        } catch (rearrangeError) {
            alert(getLabel("alert.errorPrefix") + rearrangeError.message);
            return { ok: false, rearranged: false };
        }
    }

    /**
     * 列数の文字列を数値にし、下限でクランプする
     * @param {string} columnsText - 列数の入力値
     * @returns {number} 1 以上の列数
     */
    function parseColumnCount(columnsText) {
        var columns = parseInt(columnsText, 10);
        if (isNaN(columns)) columns = DEFAULT_COLUMN_COUNT;
        if (columns < MIN_COLUMN_COUNT) columns = MIN_COLUMN_COUNT;
        return columns;
    }

    /**
     * 列数入力を読み取り、下限でクランプする
     * @param {EditText} columnsInput - 列数の入力欄
     * @returns {number} 1 以上の列数
     */
    function readColumnCount(columnsInput) {
        var columns = parseColumnCount(columnsInput.text);
        columnsInput.text = columns;
        return columns;
    }

    /**
     * ［アートボード］パネルの並び順を更新する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} dialogUI - buildDialogUI() が返すUI参照
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
    function applyPanelReorder(doc, dialogUI, tolerance) {
        /* 並び順モードに応じてソーター切替（変更しない場合はスキップ）
         * Pick a sorter based on the selected sort mode; skip when "Keep as is" */
        if (dialogUI.preview.sortModeRadios.keepAsIs.value) return;

        var sortFunction;
        if (dialogUI.preview.sortModeRadios.byName.value) {
            sortFunction = function (artboardEntries) {
                sortArtboardsByName(artboardEntries);
            };
        } else {
            sortFunction = function (artboardEntries, decimalPlaces) {
                sortArtboardsTopLeftWithTolerance(artboardEntries, decimalPlaces, tolerance);
            };
        }
        rebuildArtboardsInSortedOrder(doc, sortFunction, COORDINATE_PRECISION_DIGITS);
    }

    /**
     * アートボード名を「行-列」形式に更新する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} dialogUI - buildDialogUI() が返すUI参照
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
    function applyArtboardRenaming(doc, dialogUI, tolerance) {
        if (!dialogUI.naming.enableCheckbox.value) return;

        var separators = dialogUI.naming.separators;
        var separator = separators.underscore.value ? "_" : (separators.x.value ? "x" : "-");
        var padRadios = dialogUI.naming.padRadios;
        var padWidth = padRadios.w3.value ? 3 : (padRadios.w2.value ? 2 : 1);

        /* 配置位置から作成 / 既存名を整形 / Create from position, or reformat existing names */
        if (dialogUI.naming.source.fromPosition.value) {
            renameArtboardsFromPositions(doc, separator, padWidth, tolerance);
        } else {
            renameArtboardsFromExistingNames(doc, separator, padWidth);
        }
    }

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * ダイアログ全体を組み立てる
     * @param {number} defaultTolerance - 許容差スライダーの初期値
     * @param {number} sliderMax - 許容差スライダーの最大値
     * @returns {Object} ダイアログ（reorderDialog）とUI参照をまとめたオブジェクト
     */
    function buildDialogUI(defaultTolerance, sliderMax) {
        var reorderDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(reorderDialog);

        /* 並び順は上段にフル幅 / Order panel spans full width on top */
        var previewControls = buildPreviewPanel(reorderDialog, defaultTolerance, sliderMax);

        /* 下段は 2 カラム（左: 再配置 / 右: 命名） / Two-column row: rearrange (left) / naming (right) */
        var twoColumnsRow = reorderDialog.add("group");
        setupColumnsRow(twoColumnsRow);

        var leftColumn = twoColumnsRow.add("group");
        setupColumn(leftColumn);

        var rightColumn = twoColumnsRow.add("group");
        setupColumn(rightColumn);

        /* プロパティの評価順がそのままUIの並び順になる / Property evaluation order is the on-screen order */
        return {
            reorderDialog: reorderDialog,
            preview: previewControls,
            rearrange: buildRearrangePanel(leftColumn),
            naming: buildNamingPanel(rightColumn),
            buttons: buildDialogButtons(reorderDialog)
        };
    }

    /**
     * プレビューパネル（並び順モード＋許容差スライダー＋並び順リスト）を作成する
     * @param {Group|Window} parentContainer - 追加先のコンテナ
     * @param {number} defaultTolerance - 許容差スライダーの初期値
     * @param {number} sliderMax - 許容差スライダーの最大値
     * @returns {Object} 並び順ラジオ・スライダー・リストの参照
     */
    function buildPreviewPanel(parentContainer, defaultTolerance, sliderMax) {
        var reorderPanel = parentContainer.add("panel", undefined, getLabel("panel.reorder"));
        setupPanel(reorderPanel, 6);

        /* 名前順（ラジオ） / Radio "By name" */
        var byNameRow = reorderPanel.add("group");
        setupRow(byNameRow);
        var sortByNameRadio = byNameRow.add("radiobutton", undefined, getLabel("radio.sortByName"));
        sortByNameRadio.helpTip = getLabel("tooltip.sortByName");

        /* カンバス上の並び順に（ラジオ）＋ 許容差スライダーを同じ行に配置
         * Radio "Match canvas order" and tolerance slider on the same row */
        var byPositionRow = reorderPanel.add("group");
        setupRow(byPositionRow);

        var sortByPositionRadio = byPositionRow.add("radiobutton", undefined, getLabel("radio.sortByPosition"));
        sortByPositionRadio.helpTip = getLabel("tooltip.sortByPosition");
        sortByPositionRadio.value = true;

        var toleranceSlider = byPositionRow.add("slider", undefined, defaultTolerance, 0, sliderMax);
        toleranceSlider.helpTip = getLabel("tooltip.tolerance");
        toleranceSlider.preferredSize = [SLIDER_WIDTH, SLIDER_HEIGHT];

        /* 変更しない（ラジオ） / Radio "Keep as is" */
        var keepAsIsRow = reorderPanel.add("group");
        setupRow(keepAsIsRow);
        var sortKeepAsIsRadio = keepAsIsRow.add("radiobutton", undefined, getLabel("radio.sortKeepAsIs"));
        sortKeepAsIsRadio.helpTip = getLabel("tooltip.sortKeepAsIs");

        var reorderList = reorderPanel.add("listbox", undefined, [], { multiselect: false });
        reorderList.preferredSize.width = PREVIEW_LIST_WIDTH;
        reorderList.minimumSize.height = PREVIEW_LIST_HEIGHT;
        reorderList.alignment = ["fill", "fill"];

        return {
            sortModeRadios: {
                byPosition: sortByPositionRadio,
                byName: sortByNameRadio,
                keepAsIs: sortKeepAsIsRadio
            },
            toleranceSlider: toleranceSlider,
            reorderList: reorderList
        };
    }

    /**
     * 再配置パネルを作成する
     * @param {Group} parentContainer - 追加先のコンテナ
     * @returns {Object} 再配置設定コントロールの参照
     */
    function buildRearrangePanel(parentContainer) {
        var rearrangePanel = parentContainer.add("panel", undefined, getLabel("panel.rearrange"));
        setupPanel(rearrangePanel);

        var modeChecks = buildRearrangeModeCheckboxes(rearrangePanel);

        var rearrangeSettingsGroup = rearrangePanel.add("group");
        setupColumn(rearrangeSettingsGroup);

        var columnsRow = addLabeledInput(rearrangeSettingsGroup, labelText("fieldLabel.columns"), String(DEFAULT_COLUMN_COUNT), { min: MIN_COLUMN_COUNT, integer: true });
        var spacingControls = buildRearrangeSpacingControls(rearrangeSettingsGroup);
        var duplicateRadios = buildDuplicateHandlingPanel(rearrangePanel);

        /**
         * 再配置パネルの各コントロールの有効／無効を、選択中のモードに合わせる
         * @returns {void}
         */
        function syncEnabled() {
            syncRearrangePanelState(modeChecks, rearrangeSettingsGroup, columnsRow, duplicateRadios.panel);
            /* 親の有効／無効は自作描画に出ないため、連動アイコンも描き直す / Custom drawing does not follow the parent's state, so redraw the link icon */
            redrawLinkToggle(spacingControls.gapLinkToggle);
        }

        modeChecks.byColumns.onClick = function () {
            if (modeChecks.byColumns.value) modeChecks.byName.value = false;
            syncEnabled();
        };
        modeChecks.byName.onClick = function () {
            if (modeChecks.byName.value) modeChecks.byColumns.value = false;
            syncEnabled();
        };
        syncEnabled();

        return {
            modeChecks: modeChecks,
            columnsInput: columnsRow.input,
            columnGapInput: spacingControls.columnGapInput,
            rowGapInput: spacingControls.rowGapInput,
            gapLinkToggle: spacingControls.gapLinkToggle,
            duplicateRadios: {
                append: duplicateRadios.append,
                groupLast: duplicateRadios.groupLast
            }
        };
    }

    /**
     * 再配置モードのチェックボックスを作成する（列数指定 / アートボード名、排他的に選択）
     * @param {Panel} parentContainer - 追加先のパネル
     * @returns {Object} 各モードのチェックボックス参照
     */
    function buildRearrangeModeCheckboxes(parentContainer) {
        var modeGroup = parentContainer.add("group");
        setupColumn(modeGroup, ["left", "top"]);

        var byColumnsCheckbox = modeGroup.add("checkbox", undefined, getLabel("checkbox.rearrangeByColumns"));
        byColumnsCheckbox.helpTip = getLabel("tooltip.rearrangeByColumns");
        var byNameCheckbox = modeGroup.add("checkbox", undefined, getLabel("checkbox.rearrangeByName"));
        byNameCheckbox.helpTip = getLabel("tooltip.rearrangeByName");
        byColumnsCheckbox.value = false;
        byNameCheckbox.value = false;

        return {
            byColumns: byColumnsCheckbox,
            byName: byNameCheckbox
        };
    }

    /**
     * 再配置の列間／行間入力と連動アイコンを作成する
     * @param {Group} parentContainer - 追加先のコンテナ
     * @returns {Object} 列間・行間入力欄と連動アイコンの参照
     */
    function buildRearrangeSpacingControls(parentContainer) {
        /* 列間／行間の2行の右に連動アイコンを置く / Gap inputs (left) + link icon (right), vertically centred */
        var gapsRow = parentContainer.add("group");
        setupRow(gapsRow, "left", COLUMN_SPACING);

        var gapsLeft = gapsRow.add("group");
        setupColumn(gapsLeft, ["left", "top"], ["left", "top"]);

        var columnGapRow = addLabeledInput(gapsLeft, labelText("fieldLabel.columnGap"), String(DEFAULT_GAP_VALUE), { min: MIN_GAP_VALUE });
        var columnGapInput = columnGapRow.input;
        columnGapRow.row.add("statictext", undefined, currentUnitLabel);

        var rowGapRow = addLabeledInput(gapsLeft, labelText("fieldLabel.rowGap"), String(DEFAULT_GAP_VALUE), { min: MIN_GAP_VALUE });
        var rowGapInput = rowGapRow.input;
        rowGapRow.row.add("statictext", undefined, currentUnitLabel);

        var gapLinkToggle = bindGapLinkControls(gapsRow, columnGapInput, rowGapInput, rowGapRow.row);

        return {
            columnGapInput: columnGapInput,
            rowGapInput: rowGapInput,
            gapLinkToggle: gapLinkToggle
        };
    }

    /**
     * 列間／行間の連動アイコンを追加し、連動挙動を設定する
     * @param {Group} parentRow - アイコンの追加先（入力欄の列の右）
     * @param {EditText} columnGapInput - 列間の入力欄
     * @param {EditText} rowGapInput - 行間の入力欄
     * @param {Group} rowGapControlRow - 行間の入力行（連動時にディムする）
     * @returns {Group} 連動アイコン（.value で連動中かを読む）
     */
    function bindGapLinkControls(parentRow, columnGapInput, rowGapInput, rowGapControlRow) {
        var gapLinkToggle = addLinkToggle(parentRow, true, function () {
            if (gapLinkToggle.value) copyGapText(columnGapInput, rowGapInput);
            syncRowGapEnabled();
        });
        gapLinkToggle.helpTip = getLabel("tooltip.gapLink");

        /**
         * 行間の入力欄の有効／無効を、間隔の連動アイコンに合わせる
         * @returns {void}
         */
        function syncRowGapEnabled() {
            rowGapControlRow.enabled = !gapLinkToggle.value;
            redrawSteppersIn(rowGapControlRow);
        }

        /**
         * 入力欄の値をもう一方へ写す（直前の正しい値としても控える）
         * @param {EditText} sourceInput - 写す元
         * @param {EditText} targetInput - 写す先
         * @returns {void}
         */
        function copyGapText(sourceInput, targetInput) {
            targetInput.text = sourceInput.text;
            targetInput.lastValidText = sourceInput.text;
        }

        columnGapInput.onChange = function () {
            normalizeLabeledInput(columnGapInput);
            if (gapLinkToggle.value) copyGapText(columnGapInput, rowGapInput);
        };
        rowGapInput.onChange = function () {
            normalizeLabeledInput(rowGapInput);
            if (gapLinkToggle.value) copyGapText(rowGapInput, columnGapInput);
        };

        syncRowGapEnabled();
        return gapLinkToggle;
    }

    /**
     * 未指定／重複の扱いパネルを作成する
     * @param {Panel} parentContainer - 追加先のパネル
     * @returns {Object} パネルと各ラジオの参照
     */
    function buildDuplicateHandlingPanel(parentContainer) {
        var duplicatePanel = parentContainer.add("panel", undefined, getLabel("panel.duplicateHandling"));
        setupPanel(duplicatePanel, 6);

        var duplicateAppendRadio = duplicatePanel.add("radiobutton", undefined, getLabel("radio.duplicateAppendToRowEnd"));
        duplicateAppendRadio.helpTip = getLabel("tooltip.duplicateAppendToRowEnd");
        var duplicateGroupRadio = duplicatePanel.add("radiobutton", undefined, getLabel("radio.duplicateGroupInLastRow"));
        duplicateGroupRadio.helpTip = getLabel("tooltip.duplicateGroupInLastRow");
        duplicateAppendRadio.value = true;

        return {
            panel: duplicatePanel,
            append: duplicateAppendRadio,
            groupLast: duplicateGroupRadio
        };
    }

    /**
     * 再配置パネルの有効／無効状態を同期する
     * @param {Object} modeChecks - 再配置モードのチェックボックス参照
     * @param {Group} rearrangeSettingsGroup - 再配置設定のコンテナ
     * @param {Object} columnsRow - 列数入力行（addLabeledInput の戻り値）
     * @param {Panel} duplicatePanel - 未指定／重複パネル
     * @returns {void}
     */
    function syncRearrangePanelState(modeChecks, rearrangeSettingsGroup, columnsRow, duplicatePanel) {
        var hasRearrangeMode = modeChecks.byColumns.value || modeChecks.byName.value;
        rearrangeSettingsGroup.enabled = hasRearrangeMode;
        /* 列数は「列数を指定して再配置」のときだけアクティブ
         * (byName では名前から行列を取るので列数は不要) / Columns input is enabled only for byColumns */
        columnsRow.row.enabled = modeChecks.byColumns.value;
        /* 未指定／重複は byName のときだけアクティブ / Duplicate panel is active only for byName */
        duplicatePanel.enabled = modeChecks.byName.value;
        redrawSteppersIn(rearrangeSettingsGroup); /* ∧∨のディム表示を描き直す / redraw the steppers' dimming */
    }

    /**
     * 命名パネルを作成する
     * @param {Group} parentContainer - 追加先のコンテナ
     * @returns {Object} 命名設定コントロールの参照
     */
    function buildNamingPanel(parentContainer) {
        var namingPanel = parentContainer.add("panel", undefined, getLabel("panel.naming"));
        setupPanel(namingPanel);

        var namingEnableCheckbox = namingPanel.add("checkbox", undefined, getLabel("checkbox.namingEnable"));
        namingEnableCheckbox.helpTip = getLabel("tooltip.namingEnable");
        namingEnableCheckbox.value = false;

        var namingSettingsGroup = namingPanel.add("group");
        setupColumn(namingSettingsGroup);

        /* 命名ソース / Naming source */
        var sourceGroup = namingSettingsGroup.add("group");
        setupColumn(sourceGroup, ["left", "top"]);
        var fromPositionRadio = sourceGroup.add("radiobutton", undefined, getLabel("radio.namingFromPosition"));
        fromPositionRadio.helpTip = getLabel("tooltip.namingFromPosition");
        var fromExistingRadio = sourceGroup.add("radiobutton", undefined, getLabel("radio.namingFromExisting"));
        fromExistingRadio.helpTip = getLabel("tooltip.namingFromExisting");
        fromPositionRadio.value = true;

        /* 区切り文字パネルと桁数パネルを横並び / Separator and digits panels side by side */
        var separatorPadRow = namingSettingsGroup.add("group");
        setupColumnsRow(separatorPadRow);

        var separatorRadios = buildOptionRadioPanel(separatorPadRow, getLabel("panel.namingSeparator"), ["-", "_", "x"], getLabel("tooltip.namingSeparator"));
        var padRadios = buildOptionRadioPanel(separatorPadRow, getLabel("panel.namingPadWidth"), ["0", "00", "000"], getLabel("tooltip.namingPadWidth"));

        /* 有効／無効の同期は bindEvents() が担当（許容差スライダーの状態と連動するため）
         * bindEvents() owns the enable sync, since it also drives the tolerance slider */
        return {
            enableCheckbox: namingEnableCheckbox,
            settingsGroup: namingSettingsGroup,
            source: {
                fromPosition: fromPositionRadio,
                fromExisting: fromExistingRadio
            },
            separators: {
                hyphen: separatorRadios[0],
                underscore: separatorRadios[1],
                x: separatorRadios[2]
            },
            padRadios: {
                w1: padRadios[0],
                w2: padRadios[1],
                w3: padRadios[2]
            }
        };
    }

    /**
     * 選択肢ラジオを縦に並べたパネルを作成する（先頭を選択状態にする）
     * @param {Group} parentContainer - 追加先のコンテナ
     * @param {string} panelLabel - パネルのラベル
     * @param {string[]} optionLabels - ラジオのラベル
     * @param {string} optionTooltip - 各ラジオのツールチップ
     * @returns {RadioButton[]} 作成したラジオボタン
     */
    function buildOptionRadioPanel(parentContainer, panelLabel, optionLabels, optionTooltip) {
        var optionPanel = parentContainer.add("panel", undefined, panelLabel);
        setupPanel(optionPanel, 6);

        var optionGroup = optionPanel.add("group");
        setupColumn(optionGroup, ["left", "top"]);

        var optionRadios = [];
        for (var optionIndex = 0; optionIndex < optionLabels.length; optionIndex++) {
            var optionRadio = optionGroup.add("radiobutton", undefined, optionLabels[optionIndex]);
            optionRadio.helpTip = optionTooltip;
            optionRadios.push(optionRadio);
        }
        if (optionRadios.length > 0) optionRadios[0].value = true;
        return optionRadios;
    }

    /**
     * キャンセル / OK ボタンを作成する
     * @param {Window} parentWindow - 対象ダイアログ
     * @returns {{btnCancel: Button, btnOK: Button}} 各ボタンの参照
     */
    function buildDialogButtons(parentWindow) {
        var buttonRow = addButtonRow(parentWindow);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var buttonWidth = Math.max(btnOK.preferredSize.width, btnCancel.preferredSize.width);
        btnOK.preferredSize.width = buttonWidth;
        btnCancel.preferredSize.width = buttonWidth;
        alignRightOnlyButtonRow(buttonRow);

        return {
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * ラベル付き edittext 行を追加する（入力欄の左に∧∨を置き、↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} parentContainer - 追加先のコンテナ
     * @param {string} fieldLabelText - ラベル文字列（コロン付き）
     * @param {string} defaultValue - 入力欄の初期値
     * @param {Object} stepOptions - ∧∨の設定（min / integer など。addStepper() に渡す）
     * @returns {{row: Group, input: EditText}} 行グループと入力欄
     */
    function addLabeledInput(parentContainer, fieldLabelText, defaultValue, stepOptions) {
        var labeledRow = parentContainer.add("group");
        setupRow(labeledRow);
        var fieldLabel = labeledRow.add("statictext", undefined, fieldLabelText);
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = labeledRow.add("group");
        setupRow(stepperInputGroup, "left", 0);
        stepperInputGroup.margins = 0;

        var fieldInput;
        /* ↑↓・∧∨のあとは onChange を呼んで連動を反映する / fire onChange so linked fields follow */
        stepOptions.onStep = function (numberInput) { numberInput.notify("onChange"); };
        var stepperGroup = addStepper(stepperInputGroup, function () { return fieldInput; }, stepOptions);
        fieldInput = stepperInputGroup.add("edittext", undefined, defaultValue);
        fieldInput.characters = FIELD_INPUT_CHARS;
        bindSteppedArrowKeys(fieldInput, stepperGroup);

        /* 直接入力も整数化・下限でそろえる（onChange を差し替える側も normalizeLabeledInput() を呼ぶ）
         * Normalize typed values too; handlers that replace onChange call normalizeLabeledInput() themselves */
        fieldInput.stepOptions = stepOptions;
        fieldInput.lastValidText = fieldInput.text;
        fieldInput.onChange = function () { normalizeLabeledInput(fieldInput); };
        return { row: labeledRow, input: fieldInput };
    }

    /**
     * 入力欄の値を整数化・下限でそろえる。数値でなければ直前の値に戻す
     * @param {EditText} fieldInput - addLabeledInput() で作った入力欄
     * @returns {void}
     */
    function normalizeLabeledInput(fieldInput) {
        var value = parseFloat(fieldInput.text);
        if (isNaN(value)) {
            fieldInput.text = fieldInput.lastValidText;
            return;
        }
        writeSteppedValue(fieldInput, value, fieldInput.stepOptions);
    }

    // =========================================
    // ソート処理 / Sorting
    // =========================================

    /**
     * アートボードのプロパティをコピーする
     * @param {Artboard|Object} sourceArtboard - コピー元
     * @param {Artboard|Object} targetArtboard - コピー先
     * @returns {void}
     */
    function copyArtboardProps(sourceArtboard, targetArtboard) {
        for (var propertyIndex = 0; propertyIndex < ARTBOARD_COPYABLE_PROPS.length; propertyIndex++) {
            var propertyName = ARTBOARD_COPYABLE_PROPS[propertyIndex];
            targetArtboard[propertyName] = sourceArtboard[propertyName];
        }
    }

    /**
     * アートボードを並べ替えて［アートボード］パネル順を再構築する
     * @param {Document} doc - 対象ドキュメント
     * @param {function(ArtboardEntry[], number): void} sorterFunction - 並べ替え関数
     * @param {number} precisionDigits - 座標比較の丸め桁数
     * @returns {void}
     */
    function rebuildArtboardsInSortedOrder(doc, sorterFunction, precisionDigits) {
        var decimalPlaces = Math.pow(10, precisionDigits || COORDINATE_PRECISION_DIGITS);

        /* 元アートボードのスナップショット（プロパティ＋rect）を取得 / Snapshot original artboards */
        var artboardSnapshots = [];
        var sortableEntries = [];
        var liveOriginalIndexes = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var sourceArtboard = doc.artboards[artboardIndex];
            var snapshot = { artboardRect: sourceArtboard.artboardRect };
            copyArtboardProps(sourceArtboard, snapshot);
            artboardSnapshots.push(snapshot);
            sortableEntries.push({
                artboardRect: sourceArtboard.artboardRect,
                sourceIndex: artboardIndex,
                name: sourceArtboard.name
            });
            liveOriginalIndexes.push(artboardIndex);
        }

        sorterFunction(sortableEntries, decimalPlaces);

        /* 1件複製するたびに元を削除する。まとめて複製すると一時的にアートボード数が倍になり、
         * Illustrator のアートボード上限に当たるため。未削除の元は常に先頭側に残る。
         * Copy-then-delete one at a time: duplicating them all at once would temporarily
         * double the artboard count and hit Illustrator's limit. Pending originals stay at the front. */
        for (var sortedIndex = 0; sortedIndex < sortableEntries.length; sortedIndex++) {
            var sourceIndex = sortableEntries[sortedIndex].sourceIndex;
            var sourceSnapshot = artboardSnapshots[sourceIndex];
            var newArtboard = doc.artboards.add(sourceSnapshot.artboardRect);
            copyArtboardProps(sourceSnapshot, newArtboard);

            var livePosition = indexOfValue(liveOriginalIndexes, sourceIndex);
            doc.artboards[livePosition].remove();
            liveOriginalIndexes.splice(livePosition, 1);
        }
    }

    /**
     * 配列から値の位置を探す（ES3 に Array#indexOf がないため）
     * @param {Array} list - 検索対象の配列
     * @param {*} value - 探す値
     * @returns {number} 見つかった位置。見つからなければ -1
     */
    function indexOfValue(list, value) {
        for (var searchIndex = 0; searchIndex < list.length; searchIndex++) {
            if (list[searchIndex] === value) return searchIndex;
        }
        return -1;
    }

    /**
     * 座標を指定の係数で丸める
     * @param {number} value - 座標値
     * @param {number} decimalPlaces - 丸め係数（10 の冪）
     * @returns {number} 丸めた座標値
     */
    function roundCoordinate(value, decimalPlaces) {
        return Math.round(value * decimalPlaces) / decimalPlaces;
    }

    /**
     * 上辺の近さで行バンド（0 始まりの行番号）を割り当てる
     * 行の基準は「その行でいちばん上のアートボードの上辺」なので、
     * 許容差ずつ階段状にずれた並びが1行に連鎖して吸収されることはない
     * Each row is anchored to its own topmost edge, so a staircase of
     * within-tolerance steps no longer collapses into a single row.
     * @param {ArtboardEntry[]} artboardEntries - 対象の配列（上辺降順に並べ替え、rowBand を書き込む）
     * @param {number} decimalPlaces - 座標の丸め係数
     * @param {number} tolerance - 同一行とみなす上辺の差（pt）
     * @returns {void}
     */
    function assignRowBands(artboardEntries, decimalPlaces, tolerance) {
        artboardEntries.sort(function (firstEntry, secondEntry) {
            return roundCoordinate(secondEntry.artboardRect[1], decimalPlaces) -
                roundCoordinate(firstEntry.artboardRect[1], decimalPlaces);
        });

        var bandIndex = 0;
        var bandTop = null;
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            var currentTop = roundCoordinate(artboardEntries[entryIndex].artboardRect[1], decimalPlaces);
            if (bandTop === null) {
                bandTop = currentTop;
            } else if (bandTop - currentTop > tolerance) {
                bandIndex++;
                bandTop = currentTop;
            }
            artboardEntries[entryIndex].rowBand = bandIndex;
        }
    }

    /**
     * 許容差付きの左上基準ソート（行バンド昇順 → 左辺昇順）
     * 行バンドを整数化してから比較するため、比較関数は推移律を満たす
     * @param {ArtboardEntry[]} artboardEntries - 並べ替える配列（破壊的に並べ替える）
     * @param {number} decimalPlaces - 座標の丸め係数
     * @param {number} tolerance - 同一行とみなす上辺の差（pt）
     * @returns {void}
     */
    function sortArtboardsTopLeftWithTolerance(artboardEntries, decimalPlaces, tolerance) {
        decimalPlaces = decimalPlaces || Math.pow(10, COORDINATE_PRECISION_DIGITS);
        assignRowBands(artboardEntries, decimalPlaces, tolerance);
        artboardEntries.sort(function (firstEntry, secondEntry) {
            if (firstEntry.rowBand !== secondEntry.rowBand) {
                return firstEntry.rowBand - secondEntry.rowBand;
            }
            return roundCoordinate(firstEntry.artboardRect[0], decimalPlaces) -
                roundCoordinate(secondEntry.artboardRect[0], decimalPlaces);
        });
    }

    /**
     * 名前順（自然順）にソートする。数字列を10桁ゼロ埋めして比較し、大文字小文字は無視する
     * @param {ArtboardEntry[]} artboardEntries - 並べ替える配列（破壊的に並べ替える）
     * @returns {void}
     */
    function sortArtboardsByName(artboardEntries) {
        /* 比較関数つき sort() は遅いので、「整えた名前＋区切り＋元の位置」の文字列を引数なしの sort() で並べる。
         * 区切りの \u0001 は名前に使う文字より小さいので、前方一致の短い名前が先に来る。同名は元の順を保つ
         * sort() with a comparator is slow, so sort plain "key + \u0001 + index" strings instead;
         * \u0001 sorts before any name character, and equal names keep their original order */
        var indexDigits = 6;
        var sortKeys = [];
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            var naturalKey = padNumbersForNaturalSort((artboardEntries[entryIndex].name || "").toLowerCase());
            sortKeys.push(naturalKey + "\u0001" + padNumber(entryIndex, indexDigits));
        }
        sortKeys.sort();

        var originalEntries = artboardEntries.slice();
        for (var sortedIndex = 0; sortedIndex < sortKeys.length; sortedIndex++) {
            artboardEntries[sortedIndex] = originalEntries[parseInt(sortKeys[sortedIndex].slice(-indexDigits), 10)];
        }
    }

    /**
     * 文字列中の数字列を10桁ゼロ埋めする
     * @param {string} text - 対象文字列
     * @returns {string} ゼロ埋め後の文字列
     */
    function padNumbersForNaturalSort(text) {
        return text.replace(/\d+/g, function (digitRun) {
            var zeroPadding = "0000000000";
            return (zeroPadding + digitRun).slice(-zeroPadding.length);
        });
    }

    /**
     * 行バンドごとにエントリをまとめる
     * 事前に sortArtboardsTopLeftWithTolerance() を通し、rowBand が入っている必要がある
     * @param {ArtboardEntry[]} sortedEntries - rowBand 付きでソート済みのアートボード情報
     * @returns {Array<ArtboardEntry[]>} 行ごとにまとめた配列
     */
    function groupSortedIntoRows(sortedEntries) {
        var rowGroups = [];
        for (var entryIndex = 0; entryIndex < sortedEntries.length; entryIndex++) {
            if (entryIndex === 0 || sortedEntries[entryIndex].rowBand !== sortedEntries[entryIndex - 1].rowBand) {
                rowGroups.push([]);
            }
            rowGroups[rowGroups.length - 1].push(sortedEntries[entryIndex]);
        }
        return rowGroups;
    }

    // =========================================
    // 再配置処理（列数指定） / Rearrange by column count
    // =========================================

    /**
     * 列間を spacing として GridByRow で初回レイアウトし、行ごとに追加オフセットを適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {number} columns - 列数
     * @param {number} columnGapPoints - 列間（pt）
     * @param {number} rowGapPoints - 行間（pt）
     * @returns {void}
     */
    function rearrangeArtboardsWithGaps(doc, columns, columnGapPoints, rowGapPoints) {
        /* rearrangeArtboards() はカンバスの左上から並べるので、元の並びの中心を控えて最後に戻す
         * rearrangeArtboards() lays out from the canvas top-left, so remember the original center and move back to it */
        var originalRects = [];
        for (var originalIndex = 0; originalIndex < doc.artboards.length; originalIndex++) {
            originalRects.push(doc.artboards[originalIndex].artboardRect);
        }
        var originalCenter = getRectsBoundsCenter(originalRects);

        doc.rearrangeArtboards(DocumentArtboardLayout.GridByRow, columns, columnGapPoints, true);

        /* 各アートボードの追加シフト量を placements として組み立てる
         * Build placements: each artboard gets deltaY = -row * extraRowGap (Y up). */
        var extraRowGap = rowGapPoints - columnGapPoints;
        var placements = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var rect = doc.artboards[artboardIndex].artboardRect;
            var rowShift = Math.floor(artboardIndex / columns) * extraRowGap;
            placements.push({
                artboard: doc.artboards[artboardIndex],
                oldRect: [rect[0], rect[1], rect[2], rect[3]],
                newRect: [rect[0], rect[1] - rowShift, rect[2], rect[3] - rowShift],
                deltaX: 0,
                deltaY: -rowShift
            });
        }
        /* rearrangeArtboards() はカンバスの左上端に並べるので、その位置からカンバスの範囲を割り出す。
         * 元の並びがカンバスの端にあると、中心に合わせたときにはみ出して 'CoOA' エラーになるため、範囲内に収める
         * rearrangeArtboards() puts the layout at the canvas top-left, which tells us the canvas extent; clamp to it,
         * since centering on artboards that sit at the canvas edge would push the layout off the canvas ('CoOA') */
        var nativeRects = [];
        for (var nativeIndex = 0; nativeIndex < placements.length; nativeIndex++) {
            nativeRects.push(placements[nativeIndex].oldRect);
        }
        var nativeBounds = getRectsBounds(nativeRects);
        var canvasBounds = [nativeBounds[0], nativeBounds[1], nativeBounds[0] + CANVAS_SIZE_POINTS, nativeBounds[1] - CANVAS_SIZE_POINTS];
        centerPlacementsOn(placements, originalCenter, canvasBounds);

        /* アートボードを先に動かし、成功してからアイテムを移動する / Move the artboards first, then the items */
        applyPlacementRects(placements);
        applyItemTranslations(doc, placements);
    }

    /**
     * 矩形をすべて囲む範囲を返す
     * @param {Array<number[]>} rects - [左, 上, 右, 下] の配列
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getRectsBounds(rects) {
        var bounds = [Infinity, -Infinity, -Infinity, Infinity];
        for (var rectIndex = 0; rectIndex < rects.length; rectIndex++) {
            var rect = rects[rectIndex];
            if (rect[0] < bounds[0]) bounds[0] = rect[0];
            if (rect[1] > bounds[1]) bounds[1] = rect[1];
            if (rect[2] > bounds[2]) bounds[2] = rect[2];
            if (rect[3] < bounds[3]) bounds[3] = rect[3];
        }
        return bounds;
    }

    /**
     * 矩形をすべて囲む範囲の中心を返す
     * @param {Array<number[]>} rects - [左, 上, 右, 下] の配列
     * @returns {number[]} [x, y]
     */
    function getRectsBoundsCenter(rects) {
        var bounds = getRectsBounds(rects);
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /**
     * 移動量を下限・上限の範囲に収める。範囲が成り立たない（収まりきらない）ときは動かさない
     * @param {number} shift - 移動量
     * @param {number} minShift - 下限
     * @param {number} maxShift - 上限
     * @returns {number} 収めた移動量
     */
    function clampShift(shift, minShift, maxShift) {
        if (minShift > maxShift) return 0;
        return Math.min(Math.max(shift, minShift), maxShift);
    }

    /**
     * 並べ終えた配置全体を、中心が targetCenter に来るよう平行移動する（移動量は整数に丸める）
     * @param {Array<{newRect: number[], deltaX: number, deltaY: number}>} placements - 配置情報（newRect と移動量を書き換える）
     * @param {number[]} targetCenter - 合わせる中心 [x, y]
     * @param {number[]} [limitBounds] - はみ出させない範囲 [左, 上, 右, 下]（省略時は制限しない）
     * @returns {void}
     */
    function centerPlacementsOn(placements, targetCenter, limitBounds) {
        var newRects = [];
        for (var placementIndex = 0; placementIndex < placements.length; placementIndex++) {
            newRects.push(placements[placementIndex].newRect);
        }
        var layoutBounds = getRectsBounds(newRects);
        var shiftX = Math.round(targetCenter[0] - (layoutBounds[0] + layoutBounds[2]) / 2);
        var shiftY = Math.round(targetCenter[1] - (layoutBounds[1] + layoutBounds[3]) / 2);
        if (limitBounds) {
            /* Y は上が正なので、上端は上限・下端は下限になる / Y points up: the top edge caps the shift, the bottom floors it */
            shiftX = clampShift(shiftX, Math.ceil(limitBounds[0] - layoutBounds[0]), Math.floor(limitBounds[2] - layoutBounds[2]));
            shiftY = clampShift(shiftY, Math.ceil(limitBounds[3] - layoutBounds[3]), Math.floor(limitBounds[1] - layoutBounds[1]));
        }
        for (var shiftIndex = 0; shiftIndex < placements.length; shiftIndex++) {
            var placement = placements[shiftIndex];
            var newRect = placement.newRect;
            placement.newRect = [newRect[0] + shiftX, newRect[1] + shiftY, newRect[2] + shiftX, newRect[3] + shiftY];
            placement.deltaX += shiftX;
            placement.deltaY += shiftY;
        }
    }

    /**
     * 各アートボードに newRect を代入する。途中で失敗したら代入済みの分を oldRect に戻して例外を投げ直す
     * @param {Array<{artboard: Artboard, oldRect: number[], newRect: number[]}>} placements - 配置情報
     * @returns {void}
     */
    function applyPlacementRects(placements) {
        var appliedCount = 0;
        try {
            for (; appliedCount < placements.length; appliedCount++) {
                placements[appliedCount].artboard.artboardRect = placements[appliedCount].newRect;
            }
        } catch (rectError) {
            /* カンバスの外に出るときなど。中途半端な配置を残さない / e.g. beyond the canvas; leave no half-applied layout */
            for (var rollbackIndex = appliedCount - 1; rollbackIndex >= 0; rollbackIndex--) {
                placements[rollbackIndex].artboard.artboardRect = placements[rollbackIndex].oldRect;
            }
            throw rectError;
        }
    }

    // =========================================
    // 再配置処理（アートボード名） / Rearrange by row-column names
    // =========================================

    /**
     * @typedef {Object} PlacementItem
     * @property {Artboard} artboard - 対象アートボード
     * @property {number} width - 幅（pt）
     * @property {number} height - 高さ（pt）
     * @property {boolean} matched - 名前が解析できたか
     * @property {string} matchType - "rowColumn" または "prefixNumber"
     * @property {string} prefixName - 接頭辞（小文字）
     * @property {number} rowNumber - 名前から得た行番号
     * @property {number} columnNumber - 名前から得た列番号
     * @property {number} assignedRow - 実際に割り当てた行番号
     * @property {number} assignedColumn - 実際に割り当てた列番号
     */

    /**
     * アートボード名（「行-列」または「接頭辞-番号」）に従ってグリッド状に再配置する
     * 区切り文字は -, _, x（大小文字無視）/ Separators: -, _, x (case-insensitive)
     * @param {Document} doc - 対象ドキュメント
     * @param {number} columnGapPoints - 列間（pt）
     * @param {number} rowGapPoints - 行間（pt）
     * @param {string} exceptionMode - "rowEnd"（各行の末尾に配置）または "lastRow"（最終行の次の行にまとめる）
     * @returns {boolean} 解析できる名前が1件以上あり、再配置した場合は true
     */
    function rearrangeArtboardsByRowColumnName(doc, columnGapPoints, rowGapPoints, exceptionMode) {
        var placementContext = parseArtboardNamePlacements(doc.artboards);
        if (!placementContext.hasMatchedArtboard) {
            alert(getLabel("alert.noMatch"));
            return false;
        }

        assignArtboardPlacementSlots(placementContext, exceptionMode);
        var placementMetrics = collectPlacementMetrics(placementContext.placementItems);
        computePlacementOffsets(placementContext, placementMetrics, columnGapPoints, rowGapPoints);

        /* 並べ終えた全体の中心を、元の並びの中心に合わせる / Keep the layout centered where the artboards were */
        var originalRects = [];
        for (var itemIndex = 0; itemIndex < placementContext.placementItems.length; itemIndex++) {
            originalRects.push(placementContext.placementItems[itemIndex].oldRect);
        }
        centerPlacementsOn(placementContext.placementItems, getRectsBoundsCenter(originalRects));

        applyComputedArtboardPositions(doc, placementContext.placementItems);
        return true;
    }

    /**
     * アートボード名を解析して配置候補を作成する
     * @param {Artboards} artboards - 対象のアートボードコレクション
     * @returns {Object} 配置候補と集計値をまとめたコンテキスト
     */
    function parseArtboardNamePlacements(artboards) {
        var maxColumnNumber = 0;
        var maxRowNumber = 0;
        var placementItems = [];
        var hasMatchedArtboard = false;
        var prefixOrder = [];
        var prefixRowOffsetByName = {};

        for (var artboardIndex = 0; artboardIndex < artboards.length; artboardIndex++) {
            var artboard = artboards[artboardIndex];
            var artboardName = artboard.name;
            var artboardRect = artboard.artboardRect;
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];
            var placementItem = createPlacementItem(artboard, artboardWidth, artboardHeight);

            if (applyRowColumnNameMatch(placementItem, artboardName)) {
                if (placementItem.columnNumber > maxColumnNumber) maxColumnNumber = placementItem.columnNumber;
                if (placementItem.rowNumber > maxRowNumber) maxRowNumber = placementItem.rowNumber;
                hasMatchedArtboard = true;
            } else if (applyPrefixNumberNameMatch(placementItem, artboardName, prefixOrder, prefixRowOffsetByName)) {
                if (placementItem.columnNumber > maxColumnNumber) maxColumnNumber = placementItem.columnNumber;
                hasMatchedArtboard = true;
            }
            placementItems.push(placementItem);
        }

        return {
            maxColumnNumber: maxColumnNumber,
            maxRowNumber: maxRowNumber,
            maxMatchedRowNumber: maxRowNumber + prefixOrder.length,
            placementItems: placementItems,
            hasMatchedArtboard: hasMatchedArtboard,
            prefixOrder: prefixOrder,
            prefixRowOffsetByName: prefixRowOffsetByName
        };
    }

    /**
     * 配置候補の初期オブジェクトを作成する
     * @param {Artboard} artboard - 対象アートボード
     * @param {number} artboardWidth - 幅（pt）
     * @param {number} artboardHeight - 高さ（pt）
     * @returns {PlacementItem} 配置候補
     */
    function createPlacementItem(artboard, artboardWidth, artboardHeight) {
        return {
            artboard: artboard,
            width: artboardWidth,
            height: artboardHeight,
            matched: false,
            matchType: '',
            prefixName: '',
            rowNumber: 0,
            columnNumber: 0,
            assignedRow: 0,
            assignedColumn: 0
        };
    }

    /**
     * 「行-列」形式の名前を配置候補に反映する
     * @param {PlacementItem} placementItem - 配置候補
     * @param {string} artboardName - アートボード名
     * @returns {boolean} 反映できたら true
     */
    function applyRowColumnNameMatch(placementItem, artboardName) {
        var rowColumnMatch = artboardName.match(/^(\d+)([-_x])(\d+)$/i);
        if (!rowColumnMatch) return false;

        placementItem.rowNumber = parseInt(rowColumnMatch[1], 10);
        placementItem.columnNumber = parseInt(rowColumnMatch[3], 10);
        if (placementItem.rowNumber < 1 || placementItem.columnNumber < 1) return false;
        /* 「1920x1080」のようなサイズ名は行列として読まない / Size names such as 1920x1080 are not row-column */
        if (isArtboardSizeName(rowColumnMatch[2], placementItem.rowNumber, placementItem.columnNumber, placementItem.width, placementItem.height)) return false;

        placementItem.matched = true;
        placementItem.matchType = 'rowColumn';
        return true;
    }

    /**
     * 「接頭辞-番号」形式の名前を配置候補に反映する
     * @param {PlacementItem} placementItem - 配置候補
     * @param {string} artboardName - アートボード名
     * @param {string[]} prefixOrder - 出現順の接頭辞リスト（破壊的に追加する）
     * @param {Object} prefixRowOffsetByName - 接頭辞ごとの行オフセット（破壊的に追加する）
     * @returns {boolean} 反映できたら true
     */
    function applyPrefixNumberNameMatch(placementItem, artboardName, prefixOrder, prefixRowOffsetByName) {
        /* x は英字の直後では区切りとみなさない（「index2」を接頭辞「inde」と読まない）
         * x counts as a separator only after a non-letter, so "index2" is not read as prefix "inde" */
        var prefixNumberMatch = artboardName.match(/^(.+?)[-_](\d+)$/) || artboardName.match(/^(.*[^a-z])x(\d+)$/i);
        if (!prefixNumberMatch || /^\d+$/.test(prefixNumberMatch[1])) return false;

        placementItem.prefixName = prefixNumberMatch[1].toLowerCase();
        placementItem.columnNumber = parseInt(prefixNumberMatch[2], 10);
        if (placementItem.columnNumber < 1) return false;

        placementItem.matched = true;
        placementItem.matchType = 'prefixNumber';
        if (prefixRowOffsetByName[placementItem.prefixName] === undefined) {
            prefixRowOffsetByName[placementItem.prefixName] = prefixOrder.length + 1;
            prefixOrder.push(placementItem.prefixName);
        }
        return true;
    }

    /**
     * 「数値x数値」の名前が、そのアートボードの幅×高さを表しているかを判定する（pt／px と定規の単位で比べる）
     * @param {string} separator - 名前の区切り文字（x のときだけ判定する）
     * @param {number} firstNumber - 区切りの前の数値
     * @param {number} secondNumber - 区切りの後の数値
     * @param {number} artboardWidth - アートボードの幅（pt）
     * @param {number} artboardHeight - アートボードの高さ（pt）
     * @returns {boolean} サイズ名なら true
     */
    function isArtboardSizeName(separator, firstNumber, secondNumber, artboardWidth, artboardHeight) {
        if (separator.toLowerCase() !== "x") return false;
        var pointsPerUnitList = [1, getUnitInfo().pointsPerUnit];
        for (var unitIndex = 0; unitIndex < pointsPerUnitList.length; unitIndex++) {
            var pointsPerUnit = pointsPerUnitList[unitIndex];
            if (Math.round(artboardWidth / pointsPerUnit) === firstNumber &&
                Math.round(artboardHeight / pointsPerUnit) === secondNumber) return true;
        }
        return false;
    }

    /**
     * 各配置候補に行・列番号を割り当てる
     * @param {Object} placementContext - parseArtboardNamePlacements() の戻り値
     * @param {string} exceptionMode - "rowEnd" または "lastRow"
     * @returns {void}
     */
    function assignArtboardPlacementSlots(placementContext, exceptionMode) {
        var occupiedSlots = {};
        var nextFreeColumnByRow = {};
        var rowEndFirstColumn = placementContext.maxColumnNumber + 1;
        var currentRowNumber = 1;
        var exceptionRowNumber = placementContext.maxMatchedRowNumber + 1;
        var placementItems = placementContext.placementItems;

        for (var assignIndex = 0; assignIndex < placementItems.length; assignIndex++) {
            var assignTarget = placementItems[assignIndex];

            if (assignTarget.matched) {
                currentRowNumber = assignMatchedPlacementSlot(assignTarget, placementContext, occupiedSlots, nextFreeColumnByRow);
            } else if (exceptionMode === 'lastRow') {
                /* 最終行の次の行に、列 1 から出現順に並べる / Row after the last one, from column 1 in order */
                assignTarget.assignedRow = exceptionRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(exceptionRowNumber, 1, occupiedSlots, nextFreeColumnByRow);
            } else {
                /* 直前に割り当てた行の末尾に置く / End of the row assigned just before */
                assignTarget.assignedRow = currentRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(currentRowNumber, rowEndFirstColumn, occupiedSlots, nextFreeColumnByRow);
            }
        }
    }

    /**
     * 一致した名前の配置スロットを割り当てる
     * @param {PlacementItem} assignTarget - 配置候補
     * @param {Object} placementContext - parseArtboardNamePlacements() の戻り値
     * @param {Object} occupiedSlots - 使用済みスロット（"行,列" をキーにする）
     * @param {Object} nextFreeColumnByRow - 行ごとの次に試す列番号
     * @returns {number} 割り当てた行番号
     */
    function assignMatchedPlacementSlot(assignTarget, placementContext, occupiedSlots, nextFreeColumnByRow) {
        var targetRowNumber = assignTarget.rowNumber;
        if (assignTarget.matchType === 'prefixNumber') {
            targetRowNumber = placementContext.maxRowNumber + placementContext.prefixRowOffsetByName[assignTarget.prefixName];
        }

        assignTarget.assignedRow = targetRowNumber;
        var requestedKey = targetRowNumber + ',' + assignTarget.columnNumber;
        if (!occupiedSlots[requestedKey]) {
            occupiedSlots[requestedKey] = true;
            assignTarget.assignedColumn = assignTarget.columnNumber;
        } else {
            /* 同じスロットが埋まっている場合は行末に逃がす / Fall back to the row end when the slot is taken */
            assignTarget.assignedColumn = reserveNextFreeColumn(
                targetRowNumber, placementContext.maxColumnNumber + 1, occupiedSlots, nextFreeColumnByRow
            );
        }
        return targetRowNumber;
    }

    /**
     * 指定行の firstColumn 以降で空きスロットを予約する
     * @param {number} rowNumber - 対象の行番号
     * @param {number} firstColumn - その行で最初に試す列番号
     * @param {Object} occupiedSlots - 使用済みスロット（"行,列" をキーにする）
     * @param {Object} nextFreeColumnByRow - 行ごとの次に試す列番号
     * @returns {number} 予約した列番号
     */
    function reserveNextFreeColumn(rowNumber, firstColumn, occupiedSlots, nextFreeColumnByRow) {
        if (nextFreeColumnByRow[rowNumber] === undefined) {
            nextFreeColumnByRow[rowNumber] = firstColumn;
        }
        while (true) {
            var candidateColumn = nextFreeColumnByRow[rowNumber]++;
            var candidateSlotKey = rowNumber + ',' + candidateColumn;
            if (!occupiedSlots[candidateSlotKey]) {
                occupiedSlots[candidateSlotKey] = true;
                return candidateColumn;
            }
        }
    }

    /**
     * 列ごとの最大幅・行ごとの最大高さを集計する
     * @param {PlacementItem[]} placementItems - 配置候補
     * @returns {Object} 列幅・行高さと、使用されている列番号・行番号
     */
    function collectPlacementMetrics(placementItems) {
        var columnWidthByNumber = {};
        var rowHeightByNumber = {};
        for (var aggregateIndex = 0; aggregateIndex < placementItems.length; aggregateIndex++) {
            var aggregateItem = placementItems[aggregateIndex];
            var columnKey = aggregateItem.assignedColumn;
            var rowKey = aggregateItem.assignedRow;
            if (columnWidthByNumber[columnKey] === undefined || aggregateItem.width > columnWidthByNumber[columnKey]) {
                columnWidthByNumber[columnKey] = aggregateItem.width;
            }
            if (rowHeightByNumber[rowKey] === undefined || aggregateItem.height > rowHeightByNumber[rowKey]) {
                rowHeightByNumber[rowKey] = aggregateItem.height;
            }
        }

        return {
            columnWidthByNumber: columnWidthByNumber,
            rowHeightByNumber: rowHeightByNumber,
            usedColumnNumbers: collectSortedNumericKeys(columnWidthByNumber),
            usedRowNumbers: collectSortedNumericKeys(rowHeightByNumber)
        };
    }

    /**
     * 数値キーを昇順で取り出す
     * @param {Object} numericKeyedObject - 数値をキーにしたオブジェクト
     * @returns {number[]} 昇順の数値キー
     */
    function collectSortedNumericKeys(numericKeyedObject) {
        var sortedKeys = [];
        for (var numericKey in numericKeyedObject) {
            if (numericKeyedObject.hasOwnProperty(numericKey)) {
                sortedKeys.push(parseInt(numericKey, 10));
            }
        }
        sortedKeys.sort(function (firstKey, secondKey) { return firstKey - secondKey; });
        return sortedKeys;
    }

    /**
     * 累積オフセットと移動量を計算する
     * @param {Object} placementContext - parseArtboardNamePlacements() の戻り値
     * @param {Object} placementMetrics - collectPlacementMetrics() の戻り値
     * @param {number} columnGapPoints - 列間（pt）
     * @param {number} rowGapPoints - 行間（pt）
     * @returns {void}
     */
    function computePlacementOffsets(placementContext, placementMetrics, columnGapPoints, rowGapPoints) {
        var columnOffsetByNumber = computeCumulativeOffsets(
            placementMetrics.usedColumnNumbers,
            placementMetrics.columnWidthByNumber,
            columnGapPoints,
            placementContext.maxColumnNumber + 1
        );
        var rowOffsetByNumber = computeCumulativeOffsets(
            placementMetrics.usedRowNumbers,
            placementMetrics.rowHeightByNumber,
            rowGapPoints,
            placementContext.maxMatchedRowNumber + 1
        );

        var originX = 0;
        var originY = 0;
        var placementItems = placementContext.placementItems;
        for (var placeIndex = 0; placeIndex < placementItems.length; placeIndex++) {
            var placeItem = placementItems[placeIndex];
            var oldRect = placeItem.artboard.artboardRect;
            var newLeft = originX + columnOffsetByNumber[placeItem.assignedColumn];
            var newTop = originY - rowOffsetByNumber[placeItem.assignedRow];
            placeItem.oldRect = oldRect;
            placeItem.newRect = [newLeft, newTop, newLeft + placeItem.width, newTop - placeItem.height];
            placeItem.deltaX = newLeft - oldRect[0];
            placeItem.deltaY = newTop - oldRect[1];
        }
    }

    /**
     * 累積オフセットを計算する（boundaryNumber と一致するキーの直前の間隔は EXCEPTION_BOUNDARY_GAP_MULTIPLIER 倍）
     * @param {number[]} sortedNumbers - 昇順の行番号または列番号
     * @param {Object} sizeByNumber - 行高さまたは列幅
     * @param {number} gap - 基本の間隔（pt）
     * @param {number} boundaryNumber - 例外領域が始まる番号
     * @returns {Object} 番号ごとの累積オフセット
     */
    function computeCumulativeOffsets(sortedNumbers, sizeByNumber, gap, boundaryNumber) {
        var offsetByNumber = {};
        var cumulative = 0;
        for (var orderIndex = 0; orderIndex < sortedNumbers.length; orderIndex++) {
            var currentNumber = sortedNumbers[orderIndex];
            if (orderIndex > 0) {
                var currentGap = gap;
                if (currentNumber === boundaryNumber) currentGap *= EXCEPTION_BOUNDARY_GAP_MULTIPLIER;
                cumulative += currentGap;
            }
            offsetByNumber[currentNumber] = cumulative;
            cumulative += sizeByNumber[currentNumber];
        }
        return offsetByNumber;
    }

    /**
     * オブジェクト移動とアートボード矩形更新を適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacementItem[]} placementItems - 移動量を計算済みの配置候補
     * @returns {void}
     */
    function applyComputedArtboardPositions(doc, placementItems) {
        /* アートボードを先に動かし、成功してからアイテムを移動する / Move the artboards first, then the items */
        applyPlacementRects(placementItems);
        applyItemTranslations(doc, placementItems);
    }

    // =========================================
    // アイテム移動 / Item translation
    // =========================================

    /**
     * placements に従ってアートボード上のアイテムを移動する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<{artboard: Artboard, oldRect: number[], deltaX: number, deltaY: number}>} placements - 移動量
     * @returns {void}
     */
    function applyItemTranslations(doc, placements) {
        /* ロック・非表示のレイヤーとアイテムは動かせないので、一時的に解除して最後に戻す。
         * 飛ばすとアートボードだけが動き、アートワークが置き去りになる
         * Locked or hidden layers and items cannot be moved, so release them temporarily and restore afterwards;
         * skipping them would leave the artwork behind while the artboard moves */
        var releasedStates = [];
        try {
            releaseLayerLocks(doc.layers, releasedStates);
            var itemTranslations = [];
            appendItemTranslations(doc.layers, placements, itemTranslations);
            for (var translateIndex = 0; translateIndex < itemTranslations.length; translateIndex++) {
                var pendingMove = itemTranslations[translateIndex];
                if (pendingMove.deltaX === 0 && pendingMove.deltaY === 0) continue;
                releaseItemLock(pendingMove.item, releasedStates);
                try {
                    pendingMove.item.translate(pendingMove.deltaX, pendingMove.deltaY);
                } catch (translateError) { /* 念のためのフォールバック / Defensive fallback */ }
            }
        } finally {
            restoreReleasedStates(releasedStates);
        }
    }

    /**
     * レイヤーとサブレイヤーのロック・非表示を解除し、元の状態を控える
     * @param {Layers} layerCollection - 対象のレイヤーコレクション
     * @param {Array<{target: Object, locked: boolean, hidden: boolean}>} releasedStates - 控えの追加先
     * @returns {void}
     */
    function releaseLayerLocks(layerCollection, releasedStates) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            /* 親を先に解除しないと、子レイヤーの状態を変えられない / Release the parent before its sublayers */
            if (layer.locked || !layer.visible) {
                releasedStates.push({ target: layer, locked: layer.locked, hidden: !layer.visible });
                layer.visible = true;
                layer.locked = false;
            }
            if (layer.layers.length > 0) releaseLayerLocks(layer.layers, releasedStates);
        }
    }

    /**
     * アイテムのロック・非表示を解除し、元の状態を控える
     * @param {PageItem} pageItem - 対象アイテム
     * @param {Array<{target: Object, locked: boolean, hidden: boolean}>} releasedStates - 控えの追加先
     * @returns {void}
     */
    function releaseItemLock(pageItem, releasedStates) {
        try {
            if (!pageItem.locked && !pageItem.hidden) return;
            releasedStates.push({ target: pageItem, locked: pageItem.locked, hidden: pageItem.hidden });
            pageItem.hidden = false;
            pageItem.locked = false;
        } catch (lockError) { /* 読み書きできないアイテムはそのまま移動を試す / Try moving it as is */ }
    }

    /**
     * 解除したロック・非表示を元に戻す（子から親の順に戻すため、控えた順の逆にたどる）
     * @param {Array<{target: Object, locked: boolean, hidden: boolean}>} releasedStates - 控え
     * @returns {void}
     */
    function restoreReleasedStates(releasedStates) {
        for (var stateIndex = releasedStates.length - 1; stateIndex >= 0; stateIndex--) {
            var releasedState = releasedStates[stateIndex];
            var target = releasedState.target;
            try {
                /* ロックを先に戻すと非表示にできないので、非表示を先に戻す / Hide before locking */
                if (target.typename === "Layer") {
                    if (releasedState.hidden) target.visible = false;
                } else if (releasedState.hidden) {
                    target.hidden = true;
                }
                if (releasedState.locked) target.locked = true;
            } catch (restoreError) { /* 戻せなくても処理は続ける / Keep going */ }
        }
    }

    /**
     * レイヤーとサブレイヤーを走査して移動予定を追加する
     * @param {Layers} layerCollection - 走査するレイヤーコレクション
     * @param {Array<Object>} placementItems - oldRect と移動量を持つ配置情報
     * @param {Array<Object>} itemTranslations - 追加先の配列（{ item, deltaX, deltaY }）
     * @returns {void}
     */
    function appendItemTranslations(layerCollection, placementItems, itemTranslations) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
                var pageItem = layer.pageItems[itemIndex];
                /* サブレイヤーやグループの中身は親側で移動するのでスキップ / Sublayer and group children move with their parent */
                if (pageItem.parent !== layer) continue;
                appendTranslationForItem(pageItem, placementItems, itemTranslations);
            }
            if (layer.layers && layer.layers.length > 0) {
                appendItemTranslations(layer.layers, placementItems, itemTranslations);
            }
        }
    }

    /**
     * オブジェクトの幾何中心が含まれる元アートボードを特定し、移動量を1件追加する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {Array<Object>} placementItems - oldRect と移動量を持つ配置情報
     * @param {Array<Object>} itemTranslations - 追加先の配列（{ item, deltaX, deltaY }）
     * @returns {void}
     */
    function appendTranslationForItem(pageItem, placementItems, itemTranslations) {
        var center = getItemGeometricCenter(pageItem);
        if (!center) return;
        for (var placementIndex = 0; placementIndex < placementItems.length; placementIndex++) {
            var placement = placementItems[placementIndex];
            if (isCenterInsideRect(center, placement.oldRect)) {
                itemTranslations.push({ item: pageItem, deltaX: placement.deltaX, deltaY: placement.deltaY });
                break;
            }
        }
    }

    /**
     * オブジェクトの幾何中心を取得する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {number[]|null} [x, y]、取得できない場合は null
     */
    function getItemGeometricCenter(pageItem) {
        try {
            var bounds = pageItem.geometricBounds; /* [左, 上, 右, 下] / [left, top, right, bottom] */
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (geometricBoundsError) {
            return null;
        }
    }

    /**
     * 中心座標が矩形内かを判定する（Illustrator の Y は上方向が正）
     * @param {number[]} center - [x, y]
     * @param {number[]} rect - [左, 上, 右, 下]
     * @returns {boolean} 矩形内なら true
     */
    function isCenterInsideRect(center, rect) {
        return center[0] >= rect[0] && center[0] <= rect[2] &&
            center[1] <= rect[1] && center[1] >= rect[3];
    }

    // =========================================
    // リネーム処理 / Renaming
    // =========================================

    /**
     * 数値をゼロ埋めする
     * @param {number} value - 対象の数値
     * @param {number} width - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function padNumber(value, width) {
        var padded = String(value);
        while (padded.length < width) padded = "0" + padded;
        return padded;
    }

    /**
     * 「行-列」形式の名前を組み立てる
     * @param {number} rowNumber - 行番号
     * @param {number} columnNumber - 列番号
     * @param {string} separator - 区切り文字
     * @param {number} padWidth - ゼロ埋め桁数
     * @returns {string} 組み立てた名前
     */
    function formatRowColumnName(rowNumber, columnNumber, separator, padWidth) {
        return padNumber(rowNumber, padWidth) + separator + padNumber(columnNumber, padWidth);
    }

    /**
     * 物理的な配置から「行-列」名を割り当てる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} separator - 区切り文字
     * @param {number} padWidth - ゼロ埋め桁数
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
    function renameArtboardsFromPositions(doc, separator, padWidth, tolerance) {
        var decimalPlaces = Math.pow(10, COORDINATE_PRECISION_DIGITS);

        var positionEntries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            positionEntries.push({
                sourceIndex: artboardIndex,
                artboardRect: doc.artboards[artboardIndex].artboardRect
            });
        }
        sortArtboardsTopLeftWithTolerance(positionEntries, decimalPlaces, tolerance);

        var rowGroups = groupSortedIntoRows(positionEntries);
        for (var rowIndex = 0; rowIndex < rowGroups.length; rowIndex++) {
            for (var columnIndex = 0; columnIndex < rowGroups[rowIndex].length; columnIndex++) {
                var targetArtboard = doc.artboards[rowGroups[rowIndex][columnIndex].sourceIndex];
                targetArtboard.name = formatRowColumnName(rowIndex + 1, columnIndex + 1, separator, padWidth);
            }
        }
    }

    /**
     * 既存の「行-列」名を再フォーマットする
     * @param {Document} doc - 対象ドキュメント
     * @param {string} separator - 区切り文字
     * @param {number} padWidth - ゼロ埋め桁数
     * @returns {void}
     */
    function renameArtboardsFromExistingNames(doc, separator, padWidth) {
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboard = doc.artboards[artboardIndex];
            var nameMatch = artboard.name.match(/^(\d+)([-_x])(\d+)(.*)$/i);
            if (!nameMatch) continue;
            var rowNumber = parseInt(nameMatch[1], 10);
            var columnNumber = parseInt(nameMatch[3], 10);
            var trailingText = nameMatch[4] || "";
            /* 「1920x1080」のようなサイズ名は書き換えない / Leave size names such as 1920x1080 alone */
            var artboardRect = artboard.artboardRect;
            if (isArtboardSizeName(nameMatch[2], rowNumber, columnNumber, artboardRect[2] - artboardRect[0], artboardRect[1] - artboardRect[3])) continue;
            artboard.name = formatRowColumnName(rowNumber, columnNumber, separator, padWidth) + trailingText;
        }
    }

    main();
})();
