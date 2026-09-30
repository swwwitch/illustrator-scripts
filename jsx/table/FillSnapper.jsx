#target illustrator
#targetengine "FillSnapperEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトを「動かす対象」と「スナップ基準」に分類し、対象のバウンディングボックスを最寄りの基準線へ合わせます。
パスはアンカーポイントを直接変形するため、クリップグループ内の子パスにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillSnapper.md

### Overview

Classifies the current selection into items to move and snap references, then snaps each item's bounding box to the nearest reference line.
Paths are transformed at the anchor level, so child paths inside clipping groups are handled too.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillSnapper.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FillSnapper";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillSnapper.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillSnapper.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ダイアログの初期値（tolerance / maxDistance は pt、maxDistance 0 は距離制限なし）
       Initial dialog values (tolerance / maxDistance in pt; maxDistance 0 means no limit) */
    var DEFAULT_SNAP_OPTIONS = {
        filled: true,
        strokedOnly: true,
        blank: true,
        group: true,
        clipGroup: false,
        unrotate: true,
        tolerance: 0.5,
        maxDistance: 0,
        includeGuides: true,
        includeArtboard: true,
        preview: true
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

    var OPTION_LABEL_WIDTH = 120;           /* 数値欄の項目名の幅 / Width of the field labels */

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
     * 入力欄に表示するため、数値を小数第3位で丸めて文字列にする
     * @param {number} value - 表示する数値
     * @returns {string} 整形した文字列
     */
    function formatUnitValue(value) {
        return String(Math.round(value * 1000) / 1000);
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "塗りを線にスナップ", en: "Snap Fill to Lines" }
        },
        panel: {
            target: { ja: "動かす対象", en: "Items to move" },
            snapBasis: { ja: "スナップ基準", en: "Snap references" },
            option: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            filled: { ja: "塗りのあるクローズパス", en: "Filled closed paths" },
            group: { ja: "グループ内のパスも対象にする", en: "Include paths inside groups" },
            clipGroup: { ja: "クリップグループ内のパスも対象にする", en: "Include paths inside clipping groups" },
            strokedOnly: { ja: "線だけのパス", en: "Stroke-only paths" },
            blank: { ja: "塗り／線のないオープンパス", en: "Open paths without fill or stroke" },
            includeGuides: { ja: "ガイドライン", en: "Guide lines" },
            includeArtboard: { ja: "アートボードのエッジ", en: "Artboard edges" },
            unrotate: { ja: "回転補正", en: "Rotation correction" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            tolerance: { ja: "線判定の許容差", en: "Line detection tolerance" },
            maxDistance: { ja: "最大スナップ距離", en: "Max snap distance" }
        },
        tooltip: {
            filled: { ja: "塗りのあるオブジェクトを対象にします。", en: "Targets objects that have a fill." },
            group: { ja: "グループも対象にします。", en: "Targets groups as well." },
            clipGroup: { ja: "クリップグループも対象にします。", en: "Targets clipping groups as well." },
            strokedOnly: {
                ja: "線だけのオブジェクトを、吸着先の基準として使います。",
                en: "Uses stroke-only objects as the edges to snap to."
            },
            blank: {
                ja: "塗りも線もないオブジェクトも基準に含めます。",
                en: "Includes objects with neither fill nor stroke as edges."
            },
            includeGuides: { ja: "ガイドも吸着先の基準に含めます。", en: "Includes guides as edges to snap to." },
            includeArtboard: {
                ja: "アートボードの端も吸着先の基準に含めます。",
                en: "Includes the artboard edges as edges to snap to."
            },
            unrotate: {
                ja: "回転しているオブジェクトを、いったん角度0に戻してから合わせます。",
                en: "Straightens rotated objects before snapping them."
            },
            tolerance: {
                ja: "同じ位置とみなす許容差です。",
                en: "How far apart two edges can be and still count as aligned."
            },
            maxDistance: { ja: "この距離までの基準にだけ吸着します。", en: "Only snaps to edges within this distance." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            },
            /* ステップボタン用 / for the stepper */
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
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectTarget: { ja: "スナップ対象のオブジェクトを選択してください。", en: "Select a snap target object." },
            noTarget: { ja: "変形対象が見つかりません。", en: "No transformable target found." },
            noReferenceLines: {
                ja: "基準となる罫線や枠（塗りなしパス）が見つかりません。",
                en: "No reference lines or frames (unfilled paths) found."
            }
        }
    };

    // =========================================
    // 共通ユーティリティ / Common utilities
    // =========================================

    /**
     * 配列の末尾に別の配列（またはコレクション）の要素をすべて追加する
     * @param {Array} targetList - 追加先の配列
     * @param {Array} sourceList - 追加する要素
     * @returns {void}
     */
    function appendAll(targetList, sourceList) {
        for (var i = 0; i < sourceList.length; i++) {
            targetList.push(sourceList[i]);
        }
    }

    /**
     * 選択からスナップの処理対象候補を再帰的に集める（グループ・複合パスは中身を展開）
     * @param {PageItem[]} pageItems - 走査するアイテム
     * @param {{includeNormalGroups: boolean, includeClipGroups: boolean}} groupOptions - グループを展開するかどうか
     * @returns {PageItem[]} 処理対象候補
     */
    function collectProcessableItems(pageItems, groupOptions) {
        var includeNormalGroups = (groupOptions.includeNormalGroups !== false);
        var includeClipGroups = (groupOptions.includeClipGroups === true);
        var processableItems = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "GroupItem") {
                if (pageItem.clipped === true) {
                    /* クリップグループ：マスクパスは基準線化しやすいため除外し、中身だけを再帰 / Clip group: skip the clipping mask path and recurse into contents */
                    if (!includeClipGroups) continue;
                    var clipGroupChildren = pageItem.pageItems;
                    for (var j = 0; j < clipGroupChildren.length; j++) {
                        var childItem = clipGroupChildren[j];
                        if (childItem.typename === "PathItem" && childItem.clipping === true) continue;
                        appendAll(processableItems, collectProcessableItems([childItem], groupOptions));
                    }
                } else if (includeNormalGroups) {
                    appendAll(processableItems, collectProcessableItems(pageItem.pageItems, groupOptions));
                }
            } else if (pageItem.typename === "CompoundPathItem") {
                appendAll(processableItems, collectProcessableItems(pageItem.pathItems, groupOptions));
            } else {
                processableItems.push(pageItem);
            }
        }
        return processableItems;
    }

    /**
     * アイテムが編集できるかどうかを、親グループとレイヤーのロック・表示までたどって判定する
     * @param {PageItem} pageItem - 判定するアイテム
     * @returns {boolean} 編集できるとき true
     */
    function isEditableItem(pageItem) {
        /* 親をたどる途中でプロパティを持たない型に当たることがある / some parents may lack these properties */
        try {
            if (!pageItem) return false;
            if (pageItem.locked === true || pageItem.hidden === true) return false;

            var parentItem = pageItem.parent;
            while (parentItem && parentItem.typename !== "Document") {
                if (parentItem.locked === true || parentItem.hidden === true) return false;
                parentItem = parentItem.parent;
            }

            if (pageItem.layer) {
                if (pageItem.layer.locked === true || pageItem.layer.visible === false) return false;
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    /**
     * プレビューの復元用に、編集できるパスの元の形状を保存する
     * @param {PageItem[]} pageItems - 保存する候補
     * @returns {Object[]} パスとアンカーポイントの座標の組
     */
    function captureOriginalGeometry(pageItems) {
        var snapshotData = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename !== "PathItem") continue;
            if (!isEditableItem(pageItem)) continue;

            /* 読み取れないパスは保存せずに飛ばす / skip paths whose points cannot be read */
            var savedPoints = [];
            try {
                var pathPointList = pageItem.pathPoints;
                for (var j = 0; j < pathPointList.length; j++) {
                    var pathPoint = pathPointList[j];
                    savedPoints.push({
                        anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                        leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                        rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]]
                    });
                }
            } catch (e) {
                savedPoints = [];
            }
            if (savedPoints.length > 0) {
                snapshotData.push({ pathItem: pageItem, savedPoints: savedPoints });
            }
        }
        return snapshotData;
    }

    /**
     * 保存した元の形状へ戻す
     * @param {Object[]} snapshotData - captureOriginalGeometry() の戻り値
     * @returns {void}
     */
    function restoreOriginalGeometry(snapshotData) {
        for (var i = 0; i < snapshotData.length; i++) {
            var snapshotEntry = snapshotData[i];
            if (!snapshotEntry || !snapshotEntry.pathItem || !isEditableItem(snapshotEntry.pathItem)) continue;

            /* 書き戻せないパスはそこで打ち切り、次のパスへ / stop at a path that cannot be written and move on */
            try {
                var pathPointList = snapshotEntry.pathItem.pathPoints;
                var pointCount = Math.min(snapshotEntry.savedPoints.length, pathPointList.length);
                for (var j = 0; j < pointCount; j++) {
                    var savedPoint = snapshotEntry.savedPoints[j];
                    pathPointList[j].anchor = savedPoint.anchor;
                    pathPointList[j].leftDirection = savedPoint.leftDirection;
                    pathPointList[j].rightDirection = savedPoint.rightDirection;
                }
            } catch (e) {
                /* 続行 / continue with the next path */
            }
        }
    }

    /**
     * 候補値の中から最も近い値を返す（最大距離を超えるときは元の値のまま）
     * @param {number} targetValue - 基準の値
     * @param {number[]} candidates - 候補の値
     * @param {number} maxDistance - 最大距離。0 なら制限なし
     * @returns {number} 最も近い候補、または targetValue
     */
    function findClosestCoordinate(targetValue, candidates, maxDistance) {
        if (!candidates || candidates.length === 0) return targetValue;
        var closestValue = candidates[0];
        var closestDistance = Math.abs(targetValue - closestValue);
        for (var i = 1; i < candidates.length; i++) {
            var candidateDistance = Math.abs(targetValue - candidates[i]);
            if (candidateDistance < closestDistance) {
                closestDistance = candidateDistance;
                closestValue = candidates[i];
            }
        }
        if (maxDistance && maxDistance > 0 && closestDistance > maxDistance) return targetValue;
        return closestValue;
    }

    /**
     * パスのアンカーポイントとハンドルを、指定の範囲に収まるよう直接拡大・縮小して移動する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} left - 範囲の左端
     * @param {number} top - 範囲の上端
     * @param {number} right - 範囲の右端
     * @param {number} bottom - 範囲の下端
     * @returns {boolean} 変形できたとき true
     */
    function fitPathItemToTargetBounds(pathItem, left, top, right, bottom) {
        var geometricBounds = pathItem.geometricBounds; /* [left, top, right, bottom] */
        var sourceLeft = geometricBounds[0];
        var sourceTop = geometricBounds[1];
        var sourceWidth = geometricBounds[2] - sourceLeft;
        var sourceHeight = sourceTop - geometricBounds[3];
        var targetWidthPt = right - left;
        var targetHeightPt = top - bottom;

        if (Math.abs(sourceWidth) <= 0.0001 || Math.abs(sourceHeight) <= 0.0001) return false;
        if (Math.abs(targetWidthPt) <= 0.0001 || Math.abs(targetHeightPt) <= 0.0001) return false;

        var scaleX = targetWidthPt / sourceWidth;
        var scaleY = targetHeightPt / sourceHeight;
        var pathPointList = pathItem.pathPoints;

        /**
         * 元の範囲の座標を、指定の範囲の座標に写す
         * @param {number[]} pointArray - [x, y]
         * @returns {number[]} 写した [x, y]
         */
        function mapPoint(pointArray) {
            return [
                left + (pointArray[0] - sourceLeft) * scaleX,
                top - (sourceTop - pointArray[1]) * scaleY
            ];
        }

        /* アンカーの書き込みに失敗したら失敗扱い / treat a failed anchor write as failure */
        try {
            for (var i = 0; i < pathPointList.length; i++) {
                var pathPoint = pathPointList[i];
                pathPoint.anchor = mapPoint(pathPoint.anchor);
                pathPoint.leftDirection = mapPoint(pathPoint.leftDirection);
                pathPoint.rightDirection = mapPoint(pathPoint.rightDirection);
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    /**
     * 外接矩形から基準線を追加する（細い縦長は左端、細い横長は上端、それ以外は4辺）
     * @param {number[]} geometricBounds - [left, top, right, bottom]
     * @param {number} tolerance - 線とみなす幅（pt）
     * @param {{horizontalLines: number[], verticalLines: number[]}} referenceLines - 追加先
     * @returns {void}
     */
    function addReferenceEdges(geometricBounds, tolerance, referenceLines) {
        var boundsWidthPt = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var boundsHeightPt = Math.abs(geometricBounds[1] - geometricBounds[3]);
        if (boundsWidthPt < tolerance) {
            referenceLines.verticalLines.push(geometricBounds[0]);
        } else if (boundsHeightPt < tolerance) {
            referenceLines.horizontalLines.push(geometricBounds[1]);
        } else {
            /* 矩形などは4辺すべてを採用 / Rectangles: use all four sides */
            referenceLines.horizontalLines.push(geometricBounds[1]);
            referenceLines.horizontalLines.push(geometricBounds[3]);
            referenceLines.verticalLines.push(geometricBounds[0]);
            referenceLines.verticalLines.push(geometricBounds[2]);
        }
    }

    /**
     * アクティブなアートボードの4辺を基準線として返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {{horizontalLines: number[], verticalLines: number[]}} 基準線
     */
    function gatherArtboardEdges(doc) {
        var artboardIndex;
        /* 取得できない環境では先頭のアートボード / fall back to the first artboard */
        try { artboardIndex = doc.artboards.getActiveArtboardIndex(); } catch (e) { artboardIndex = 0; }
        var artboardRect = doc.artboards[artboardIndex].artboardRect; /* [left, top, right, bottom] */
        return {
            horizontalLines: [artboardRect[1], artboardRect[3]],
            verticalLines: [artboardRect[0], artboardRect[2]]
        };
    }

    /**
     * ドキュメント内のガイドから水平線・垂直線を集める
     * @param {Document} doc - 対象のドキュメント
     * @param {number} tolerance - 線とみなす幅（pt）
     * @returns {{horizontalLines: number[], verticalLines: number[]}} 基準線
     */
    function gatherDocumentGuides(doc, tolerance) {
        var guideLines = { horizontalLines: [], verticalLines: [] };
        var documentPathItems = doc.pathItems;
        for (var i = 0; i < documentPathItems.length; i++) {
            var guidePathItem = documentPathItems[i];
            if (guidePathItem.guides !== true) continue;
            addReferenceEdges(guidePathItem.geometricBounds, tolerance, guideLines);
        }
        return guideLines;
    }

    // =========================================
    // 角度・回転補正 / Angle & Rotation correction
    // =========================================

    /**
     * パスの最初の辺の角度を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 角度（度）。点が2つ未満なら 0
     */
    function getEdgeAngleDeg(pathItem) {
        var pathPoints = pathItem.pathPoints;
        if (!pathPoints || pathPoints.length < 2) return 0;
        var firstAnchor = pathPoints[0].anchor;
        var secondAnchor = pathPoints[1].anchor;
        return Math.atan2(secondAnchor[1] - firstAnchor[1], secondAnchor[0] - firstAnchor[0]) * 180 / Math.PI;
    }

    /**
     * 最寄りの 90 度の倍数からのずれを返す
     * @param {number} angleDeg - 角度（度）
     * @returns {number} -45〜45 のずれ（度）
     */
    function getNearestRightAngleOffset(angleDeg) {
        var modulus = angleDeg % 90;
        if (modulus > 45) modulus -= 90;
        if (modulus < -45) modulus += 90;
        return modulus;
    }

    // =========================================
    // スナップ判定・実行 / Snap classification & execution
    // =========================================

    /**
     * 処理対象候補を、動かす対象と基準線に振り分ける
     * @param {PageItem[]} pageItems - 処理対象候補
     * @param {Object} snapOptions - ダイアログの設定（readSnapOptions() の戻り値）
     * @returns {{targets: PathItem[], horizontalLines: number[], verticalLines: number[]}} 振り分けの結果
     */
    function classifyTargetsAndReferences(pageItems, snapOptions) {
        var tolerance = snapOptions.tolerance;
        var classification = { targets: [], horizontalLines: [], verticalLines: [] };

        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename !== "PathItem") continue;

            var geometricBounds = pageItem.geometricBounds; /* [left, top, right, bottom] */
            var itemWidthPt = Math.abs(geometricBounds[2] - geometricBounds[0]);
            var itemHeightPt = Math.abs(geometricBounds[1] - geometricBounds[3]);

            var isFilledClosedPath = pageItem.filled && pageItem.closed;
            var isStrokedOpenPath = !pageItem.filled && pageItem.stroked && !pageItem.closed;
            var isBlankOpenPath = !pageItem.filled && !pageItem.stroked && !pageItem.closed;

            /* 変形対象：塗りのあるクローズパスのみ / Targets: filled closed paths only */
            if (isFilledClosedPath) {
                if (!snapOptions.filled) continue;
                if (itemWidthPt > tolerance && itemHeightPt > tolerance) {
                    classification.targets.push(pageItem);
                    continue;
                }
                /* 極端に細い filled closed は基準として扱う / Extremely thin filled closed: treat as reference */
            }

            /* スナップ基準（補助）：トグル OFF 時はスキップ / Snap references (auxiliary): skip if toggle is OFF */
            if (isStrokedOpenPath && !snapOptions.strokedOnly) continue;
            if (isBlankOpenPath && !snapOptions.blank) continue;

            /* 残りはすべて基準線として追加 / Remaining items contribute as reference lines */
            addReferenceEdges(geometricBounds, tolerance, classification);
        }
        return classification;
    }

    /**
     * 1つの対象を、最寄りの基準線に合わせて変形する
     * @param {PathItem} pathItem - 動かす対象
     * @param {number[]} horizontalLines - 水平の基準線（Y 座標）
     * @param {number[]} verticalLines - 垂直の基準線（X 座標）
     * @param {Object} snapOptions - ダイアログの設定
     * @returns {boolean} 変形できたとき true
     */
    function applySnapToSingleTarget(pathItem, horizontalLines, verticalLines, snapOptions) {
        if (!isEditableItem(pathItem)) return false;
        if (snapOptions.unrotate) {
            var rotationOffsetDeg = getNearestRightAngleOffset(getEdgeAngleDeg(pathItem));
            if (Math.abs(rotationOffsetDeg) > 0.001) {
                /* 回転できないアイテムはスキップ / skip items that cannot be rotated */
                try {
                    pathItem.rotate(-rotationOffsetDeg, true, true, true, true, Transformation.CENTER);
                } catch (e) {
                    return false;
                }
            }
        }

        var maxDistance = snapOptions.maxDistance;
        var geometricBounds = pathItem.geometricBounds;
        var nearestTop = findClosestCoordinate(geometricBounds[1], horizontalLines, maxDistance);
        var nearestBottom = findClosestCoordinate(geometricBounds[3], horizontalLines, maxDistance);
        var nearestLeft = findClosestCoordinate(geometricBounds[0], verticalLines, maxDistance);
        var nearestRight = findClosestCoordinate(geometricBounds[2], verticalLines, maxDistance);

        var top = Math.max(nearestTop, nearestBottom);
        var bottom = Math.min(nearestTop, nearestBottom);
        var left = Math.min(nearestLeft, nearestRight);
        var right = Math.max(nearestLeft, nearestRight);

        if (Math.abs(right - left) <= 0 || Math.abs(top - bottom) <= 0) return false;

        return fitPathItemToTargetBounds(pathItem, left, top, right, bottom);
    }

    /**
     * すべての対象にスナップを適用する
     * @param {PageItem[]} pageItems - 処理対象候補
     * @param {Object} snapOptions - ダイアログの設定
     * @returns {{snapped: number, targets: number, horizontalLines: number, verticalLines: number}} 適用結果の件数
     */
    function applySnapToTargets(pageItems, snapOptions) {
        var classification = classifyTargetsAndReferences(pageItems, snapOptions);
        var doc = app.activeDocument;

        /* ガイドラインを基準線として合算 / Merge document guides as additional reference lines */
        if (snapOptions.includeGuides) {
            var guideLines = gatherDocumentGuides(doc, snapOptions.tolerance);
            appendAll(classification.horizontalLines, guideLines.horizontalLines);
            appendAll(classification.verticalLines, guideLines.verticalLines);
        }

        /* アートボードのエッジを基準線として合算 / Merge artboard edges as additional reference lines */
        if (snapOptions.includeArtboard) {
            var artboardEdges = gatherArtboardEdges(doc);
            appendAll(classification.horizontalLines, artboardEdges.horizontalLines);
            appendAll(classification.verticalLines, artboardEdges.verticalLines);
        }

        var snappedTargetCount = 0;
        for (var i = 0; i < classification.targets.length; i++) {
            if (applySnapToSingleTarget(classification.targets[i], classification.horizontalLines, classification.verticalLines, snapOptions)) {
                snappedTargetCount++;
            }
        }
        return {
            snapped: snappedTargetCount,
            targets: classification.targets.length,
            horizontalLines: classification.horizontalLines.length,
            verticalLines: classification.verticalLines.length
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * tooltip 付きのチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * 「項目名＋∧∨＋数値欄＋単位」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 項目名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {number} valuePt - 初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 表示する単位
     * @param {number} minValue - ∧∨で下げられる下限（表示単位）
     * @returns {EditText} 数値欄
     */
    function addUnitField(parentPanel, labelPath, tooltipPath, valuePt, rulerUnit, minValue) {
        var fieldGroup = parentPanel.add("group");
        var fieldLabel = fieldGroup.add("statictext", undefined, labelText(labelPath));
        fieldLabel.preferredSize = [OPTION_LABEL_WIDTH, -1];

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var valueInput;
        /* 増減後は手入力と同じ onChange（プレビュー更新。showSnapDialog() で結線）を通す
           after stepping, run the same onChange as typing (preview) */
        var stepperGroup = addStepper(stepperInputGroup, function () { return valueInput; }, {
            min: minValue,
            onStep: function (numberInput) { if (numberInput.onChange) numberInput.onChange(); }
        });
        valueInput = stepperInputGroup.add("edittext", undefined, formatUnitValue(valuePt / rulerUnit.pointsPerUnit));
        valueInput.helpTip = getLabel(tooltipPath);
        valueInput.characters = 3;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(valueInput, stepperGroup);
        fieldGroup.add("statictext", undefined, rulerUnit.label);
        return valueInput;
    }

    /**
     * ダイアログを組み立てる
     * @param {Object} initialOptions - 初期値
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 数値欄の単位
     * @returns {Object} ダイアログ（snapDialog）と各コントロール
     */
    function buildSnapDialog(initialOptions, rulerUnit) {
        var snapDialog = new Window("dialog", getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        setupWindow(snapDialog);

        /* 「変形するもの」パネル：動かされる側 / Targets panel: items being moved */
        var targetPanel = snapDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);
        var filledCheckbox = addOptionCheckbox(targetPanel, "checkbox.filled", "tooltip.filled", initialOptions.filled);
        var groupCheckbox = addOptionCheckbox(targetPanel, "checkbox.group", "tooltip.group", initialOptions.group);
        var clipGroupCheckbox = addOptionCheckbox(targetPanel, "checkbox.clipGroup", "tooltip.clipGroup", initialOptions.clipGroup);

        /* 「スナップ基準」パネル：基準として扱うパス・追加基準 / Snap references panel: paths and extra references */
        var snapBasisPanel = snapDialog.add("panel", undefined, getLabel("panel.snapBasis"));
        setupPanel(snapBasisPanel, 6);
        var strokedOnlyCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.strokedOnly", "tooltip.strokedOnly", initialOptions.strokedOnly);
        var blankCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.blank", "tooltip.blank", initialOptions.blank);
        var includeGuidesCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.includeGuides", "tooltip.includeGuides", initialOptions.includeGuides);
        var includeArtboardCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.includeArtboard", "tooltip.includeArtboard", initialOptions.includeArtboard);

        /* 「オプション」パネル / Options panel */
        var optionPanel = snapDialog.add("panel", undefined, getLabel("panel.option"));
        setupPanel(optionPanel, 6);
        optionPanel.alignChildren = ["left", "top"];  /* 数値欄の行は広げない / keep the field rows at their natural width */
        var unrotateCheckbox = addOptionCheckbox(optionPanel, "checkbox.unrotate", "tooltip.unrotate", initialOptions.unrotate);
        var toleranceInput = addUnitField(optionPanel, "fieldLabel.tolerance", "tooltip.tolerance", initialOptions.tolerance, rulerUnit, 0.1);
        var maxDistanceInput = addUnitField(optionPanel, "fieldLabel.maxDistance", "tooltip.maxDistance", initialOptions.maxDistance, rulerUnit, 0);

        /* 下段：左＝プレビュー、中央＝余白、右＝ボタン / Bottom row: left=preview, center=spacer, right=buttons */
        var buttonRow = addButtonRow(snapDialog);
        var previewCheckbox = addOptionCheckbox(buttonRow.leftGroup, "checkbox.preview", "tooltip.preview", initialOptions.preview);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            snapDialog: snapDialog,
            filledCheckbox: filledCheckbox,
            groupCheckbox: groupCheckbox,
            clipGroupCheckbox: clipGroupCheckbox,
            strokedOnlyCheckbox: strokedOnlyCheckbox,
            blankCheckbox: blankCheckbox,
            includeGuidesCheckbox: includeGuidesCheckbox,
            includeArtboardCheckbox: includeArtboardCheckbox,
            unrotateCheckbox: unrotateCheckbox,
            toleranceInput: toleranceInput,
            maxDistanceInput: maxDistanceInput,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * ダイアログの現在の入力値を読み取る（数値は pt に換算）
     * @param {Object} dialogControls - buildSnapDialog() の戻り値
     * @param {{pointsPerUnit: number}} rulerUnit - 数値欄の単位
     * @returns {Object} スナップの設定
     */
    function readSnapOptions(dialogControls, rulerUnit) {
        var parsedTolerance = parseFloat(dialogControls.toleranceInput.text) * rulerUnit.pointsPerUnit;
        if (isNaN(parsedTolerance) || parsedTolerance <= 0) parsedTolerance = 0.5;
        var parsedMaxDistance = parseFloat(dialogControls.maxDistanceInput.text) * rulerUnit.pointsPerUnit;
        if (isNaN(parsedMaxDistance) || parsedMaxDistance < 0) parsedMaxDistance = 0;
        return {
            filled: dialogControls.filledCheckbox.value,
            strokedOnly: dialogControls.strokedOnlyCheckbox.value,
            blank: dialogControls.blankCheckbox.value,
            group: dialogControls.groupCheckbox.value,
            clipGroup: dialogControls.clipGroupCheckbox.value,
            unrotate: dialogControls.unrotateCheckbox.value,
            tolerance: parsedTolerance,
            maxDistance: parsedMaxDistance,
            includeGuides: dialogControls.includeGuidesCheckbox.value,
            includeArtboard: dialogControls.includeArtboardCheckbox.value,
            preview: dialogControls.previewCheckbox.value
        };
    }

    /**
     * ダイアログを表示し、確定した設定を返す
     * @param {Object} initialOptions - 初期値
     * @param {Function} onPreview - 設定が変わるたびに呼ぶ関数（引数は設定）
     * @returns {Object|null} 確定した設定。キャンセルなら null
     */
    function showSnapDialog(initialOptions, onPreview) {
        var rulerUnit = getUnitInfo();
        var dialogControls = buildSnapDialog(initialOptions, rulerUnit);

        /**
         * 現在の設定でプレビューを呼ぶ
         * @returns {void}
         */
        function notifyPreview() {
            onPreview(readSnapOptions(dialogControls, rulerUnit));
        }

        dialogControls.filledCheckbox.onClick = notifyPreview;
        dialogControls.strokedOnlyCheckbox.onClick = notifyPreview;
        dialogControls.blankCheckbox.onClick = notifyPreview;
        dialogControls.groupCheckbox.onClick = notifyPreview;
        dialogControls.clipGroupCheckbox.onClick = notifyPreview;
        dialogControls.unrotateCheckbox.onClick = notifyPreview;
        dialogControls.includeGuidesCheckbox.onClick = notifyPreview;
        dialogControls.includeArtboardCheckbox.onClick = notifyPreview;
        dialogControls.previewCheckbox.onClick = notifyPreview;
        dialogControls.toleranceInput.onChange = notifyPreview;
        dialogControls.maxDistanceInput.onChange = notifyPreview;

        prepareDialogWindow(dialogControls.snapDialog, SCRIPT_NAME);
        if (dialogControls.snapDialog.show() !== 1) return null;
        return readSnapOptions(dialogControls, rulerUnit);
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * プレビュー付きのダイアログで設定を決め、選択中の対象をスナップする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var selectedPageItems = app.activeDocument.selection;
        if (selectedPageItems.length < 1) {
            alert(getLabel("alert.selectTarget"));
            return;
        }

        /* スナップショットは（プレビュー切替で復元できるように）通常グループもクリップグループも展開して取得 */
        /* Snapshot covers everything — both normal and clip-group descendants — so preview toggles can reset cleanly */
        var allProcessableItemsForPreviewReset = collectProcessableItems(selectedPageItems, { includeNormalGroups: true, includeClipGroups: true });
        var originalGeometrySnapshot = captureOriginalGeometry(allProcessableItemsForPreviewReset);
        var shouldRestoreOriginalGeometry = true;

        /**
         * 現在の設定に応じて処理対象候補を集める
         * @param {Object} snapOptions - スナップの設定
         * @returns {PageItem[]} 処理対象候補
         */
        function collectItemsForOptions(snapOptions) {
            return collectProcessableItems(selectedPageItems, {
                includeNormalGroups: snapOptions.group === true,
                includeClipGroups: snapOptions.clipGroup === true
            });
        }

        /**
         * 元の形状に戻してから、プレビューが ON ならスナップを適用する
         * @param {Object} snapOptions - スナップの設定
         * @returns {void}
         */
        function applyPreview(snapOptions) {
            restoreOriginalGeometry(originalGeometrySnapshot);
            if (snapOptions.preview) {
                applySnapToTargets(collectItemsForOptions(snapOptions), snapOptions);
            }
            app.redraw();
        }

        /* 途中で抜けたとき（キャンセル・エラー）は元の形状に戻す / restore on cancel or error */
        try {
            /* 初回プレビュー / Initial preview */
            applyPreview(DEFAULT_SNAP_OPTIONS);

            var confirmedOptions = showSnapDialog(DEFAULT_SNAP_OPTIONS, applyPreview);
            if (!confirmedOptions) {
                /* キャンセル：状態を巻き戻し / Cancel: roll back to original state */
                return;
            }

            /* OK：スナップショットから本適用 / OK: re-apply cleanly from snapshot */
            restoreOriginalGeometry(originalGeometrySnapshot);
            var snapApplyResult = applySnapToTargets(collectItemsForOptions(confirmedOptions), confirmedOptions);

            if (snapApplyResult.targets === 0) {
                alert(getLabel("alert.noTarget"));
                return;
            }
            if (snapApplyResult.horizontalLines === 0 && snapApplyResult.verticalLines === 0) {
                alert(getLabel("alert.noReferenceLines"));
                return;
            }

            shouldRestoreOriginalGeometry = false;
        } finally {
            if (shouldRestoreOriginalGeometry) {
                restoreOriginalGeometry(originalGeometrySnapshot);
                app.redraw();
            }
        }
    }

    main();

})();
