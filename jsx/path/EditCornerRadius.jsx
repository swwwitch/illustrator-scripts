#target illustrator
#targetengine "EditCornerRadiusEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択・現在のアートボード・ドキュメント全体のいずれかにある角丸長方形の角の半径を、ダイアログで変更します。
吹き出し形状と、グループ・複合パス・複合シェイプの中の長方形も対象です。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditCornerRadius.md

### 注意

水平・垂直に置かれた長方形だけが対象です。回転した長方形、ロック・非表示のオブジェクトは変更しません。
複合シェイプは「グループ＋［パスファインダー：合体］」に変換するため、各パスのモードは合体になります。

### Overview

Edits the corner radius of rounded rectangles in the selection, on the current artboard, or in the entire document, using a dialog.
Callout shapes and rectangles inside groups, compound paths and compound shapes are included.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/EditCornerRadius.md

### Notes

Only rectangles aligned to the horizontal and vertical axes are changed. Rotated rectangles and locked or hidden objects are left as they are.
Compound shapes are converted to a group with the Pathfinder Add effect, so every shape mode becomes Add.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "EditCornerRadius";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditCornerRadius.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/EditCornerRadius.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビューの初期状態 / Initial state of the preview checkbox */
    var PREVIEW_DEFAULT = true;

    /* ［0の半径は0のままに］の初期状態 / Initial state of the "keep zero radii" checkbox */
    var KEEP_ZERO_RADII_DEFAULT = true;

    /* ［「角を丸くする」効果を含む］の初期状態 / Initial state of the "include Round Corners effect" checkbox */
    var INCLUDE_EFFECT_DEFAULT = false;

    /* ［「効果」に変換］の初期状態 / Initial state of the "convert to effect" checkbox */
    var CONVERT_TO_EFFECT_DEFAULT = false;

    // =========================================
    // 計測設定 / Measurement settings
    // =========================================

    /* 長さ・座標を 0 とみなす許容値（pt）/ Tolerance treated as zero for lengths and coordinates (pt) */
    var GEOMETRY_TOLERANCE = 0.01;

    /* 直線とみなす接線の角度（ラジアン）/ Tangent angle treated as straight (radians) */
    var MIN_ARC_ANGLE = 0.001;

    /* 向きの比較の許容値（単位ベクトルの成分）/ Tolerance for comparing directions (unit vector components) */
    var DIRECTION_TOLERANCE = 0.001;

    /* 対象として調べるパスのアンカー数の上限 / Maximum anchor count of paths to examine */
    var MAX_POINT_COUNT = 32;

    /* 90° の円弧をベジェで近似するときのハンドル長の係数 / Handle length ratio for a 90-degree Bezier arc */
    var ARC_HANDLE_RATIO = 0.5522847498;

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

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS        = 16;                /* ウィンドウ外周の余白 */
    var WINDOW_SPACING        = 12;                /* ウィンドウ内の要素間隔 */
    var PANEL_MARGINS         = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING         = 6;                 /* パネル内の要素間隔 */
    var FIELD_SPACING         = 6;                 /* 名前・入力欄・単位の間隔 */
    var FIELD_CHARACTERS      = 4;                 /* 半径の入力欄の文字数 */

    /**
     * パネルの共通レイアウトを設定する
     * @param {Panel} targetPanel - 対象パネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
    }

    /**
     * 横並びグループの共通レイアウトを設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = FIELD_SPACING;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    var LABELS = {
        dialog: {
            title: { ja: "角丸の半径を変更", en: "Edit Corner Radius" }
        },
        panel: {
            targetScope: { ja: "対象", en: "Target" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            selection: { ja: "選択したオブジェクトのみ", en: "Selected objects only" },
            artboard: { ja: "現在のアートボードのみ", en: "Current artboard only" },
            document: { ja: "ドキュメント全体", en: "Entire document" }
        },
        fieldLabel: {
            radius: { ja: "半径", en: "Radius" }
        },
        checkbox: {
            keepZeroRadii: { ja: "0の半径は0のままに", en: "Keep zero radii at zero" },
            includeEffect: { ja: "「角を丸くする」効果を含む", en: "Include the Round Corners effect" },
            convertToEffect: { ja: "「角を丸くする」効果に変換", en: "Convert to Round Corners effect" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            artboard: {
                ja: "現在のアートボードに一部でも重なる長方形（吹き出し形状を含む）が対象です。グループ・複合パス・複合シェイプの中も含みます（複合シェイプは合体に変換。ロック・非表示は除く）",
                en: "Rectangles (including callout shapes) that overlap the current artboard, including those inside groups, compound paths and compound shapes (compound shapes are converted to Add; locked or hidden ones are skipped)"
            },
            document: {
                ja: "ドキュメント内のすべての長方形（吹き出し形状を含む）が対象です。グループ・複合パス・複合シェイプの中も含みます（複合シェイプは合体に変換。ロック・非表示は除く）",
                en: "All rectangles (including callout shapes) in the document, including those inside groups, compound paths and compound shapes (compound shapes are converted to Add; locked or hidden ones are skipped)"
            },
            radiusField: {
                ja: "短辺の半分（吹き出しは口の付け根まで）を超える値は、そこまでに制限されます。↑↓で増減（Shift：10、Option：0.1）",
                en: "Values over half the shorter side (or the distance to a callout tail) are limited to it. Up/Down to change (Shift: 10, Option: 0.1)"
            },
            keepZeroRadii: {
                ja: "オンのときは角丸の無い角を角のまま残します",
                en: "When on, corners without rounding stay square"
            },
            includeEffect: {
                ja: "オンのときは「角を丸くする」効果で角丸になった長方形も対象にし、効果を付け直します。付け直すとほかの効果は外れます（塗り・線・不透明度は残ります）。吹き出し形状・複合パス・複合シェイプの中は対象外",
                en: "When on, rectangles rounded by the Round Corners effect are included and the effect is reapplied. Other effects are removed (fill, stroke and opacity are kept). Not for callout shapes or paths in compound paths or compound shapes"
            },
            convertToEffect: {
                ja: "オンのときは、4つの角がすべて角丸の長方形を角のない長方形に戻し、「角を丸くする」効果で角丸を付けます。吹き出し形状・複合パス・複合シェイプの中は対象外",
                en: "When on, rectangles with all four corners rounded are made square and rounded with the Round Corners effect. Not for callout shapes or paths in compound paths or compound shapes"
            },
            skippedCount: {
                ja: "水平・垂直の長方形（吹き出し形状、複合パス・複合シェイプの中を含む）のみ変更します",
                en: "Only axis-aligned rectangles (including callout shapes and those in compound paths and compound shapes) are changed"
            },
            noTargetInSelection: {
                ja: "選択の中に、水平・垂直の長方形がありません",
                en: "The selection contains no axis-aligned rectangles"
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
        status: {
            skippedCount: {
                ja: "対象外のオブジェクト：{count} 個",
                en: "Skipped objects: {count}"
            },
            noTargetInSelection: { ja: "選択に対象がありません", en: "Nothing selected can be changed" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    // =========================================
    // 半径の計測 / Radius measurement
    // =========================================

    /**
     * 2 点間の距離を返す
     * @param {number[]} fromPoint - [x, y]
     * @param {number[]} toPoint - [x, y]
     * @returns {number} 距離
     */
    function getDistance(fromPoint, toPoint) {
        var dx = toPoint[0] - fromPoint[0];
        var dy = toPoint[1] - fromPoint[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 正規化した方向ベクトルを返す
     * @param {number[]} fromPoint - 始点 [x, y]
     * @param {number[]} toPoint - 終点 [x, y]
     * @returns {number[]|null} 単位ベクトル（長さが 0 なら null）
     */
    function getUnitVector(fromPoint, toPoint) {
        var vectorLength = getDistance(fromPoint, toPoint);
        if (vectorLength < GEOMETRY_TOLERANCE) return null;
        return [(toPoint[0] - fromPoint[0]) / vectorLength, (toPoint[1] - fromPoint[1]) / vectorLength];
    }

    /**
     * 曲線セグメントを円弧とみなして半径を求める
     * 両端の接線がなす角 θ と弦の長さ c から r = c / (2 sin(θ/2))
     * @param {PathPoint} startPoint - セグメントの始点
     * @param {PathPoint} endPoint - セグメントの終点
     * @returns {number|null} 半径（直線・算出できないときは null）
     */
    function getSegmentRadius(startPoint, endPoint) {
        var startAnchor = startPoint.anchor;
        var startHandle = startPoint.rightDirection;
        var endAnchor = endPoint.anchor;
        var endHandle = endPoint.leftDirection;
        if (getDistance(startAnchor, startHandle) < GEOMETRY_TOLERANCE &&
            getDistance(endAnchor, endHandle) < GEOMETRY_TOLERANCE) return null;

        /* ハンドルが片側だけのときは、もう一方のハンドルへ向かう線を接線に使う
           When only one handle exists, use the line toward the other handle as the tangent */
        var startTangent = getUnitVector(startAnchor, startHandle) || getUnitVector(startAnchor, endHandle);
        var endTangent = getUnitVector(endHandle, endAnchor) || getUnitVector(startHandle, endAnchor);
        if (!startTangent || !endTangent) return null;

        var dotProduct = startTangent[0] * endTangent[0] + startTangent[1] * endTangent[1];
        var arcAngle = Math.acos(Math.max(-1, Math.min(1, dotProduct)));
        if (arcAngle < MIN_ARC_ANGLE) return null;

        return getDistance(startAnchor, endAnchor) / (2 * Math.sin(arcAngle / 2));
    }

    /**
     * 角丸のある角の半径の平均を返す（角丸が 1 つも無ければ 0）
     * @param {Object[]} measurements - measureCorners() の戻り値の配列
     * @returns {number} 平均の半径（pt）
     */
    function getAverageRadius(measurements) {
        var radiusSum = 0;
        var roundedCount = 0;
        for (var i = 0; i < measurements.length; i++) {
            var cornerRadii = measurements[i].effectRadii || measurements[i].pathRadii;
            for (var j = 0; j < cornerRadii.length; j++) {
                if (cornerRadii[j] < GEOMETRY_TOLERANCE) continue;
                radiusSum += cornerRadii[j];
                roundedCount++;
            }
        }
        return (roundedCount > 0) ? radiusSum / roundedCount : 0;
    }

    /**
     * 角丸のある角（半径が 0 より大きい角）の数を返す
     * @param {number[]} cornerRadii - [左上, 右上, 右下, 左下] の半径（pt）
     * @returns {number} 角丸の数（0〜4）
     */
    function countRoundedCorners(cornerRadii) {
        var roundedCount = 0;
        for (var i = 0; i < cornerRadii.length; i++) {
            if (cornerRadii[i] >= GEOMETRY_TOLERANCE) roundedCount++;
        }
        return roundedCount;
    }

    /**
     * 範囲の短辺の半分（角丸の半径の上限）を返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @returns {number} 半径の上限（pt）
     */
    function getMaxRadius(bounds) {
        return Math.min(bounds[2] - bounds[0], bounds[1] - bounds[3]) / 2;
    }

    /**
     * パスのポイントを { anchor, leftDirection, rightDirection } の配列に読み出す（DOM の読み取りを 1 回で済ませる）
     * @param {PathPoints} pathPoints - パスのポイント
     * @returns {Object[]} ポイントの配列
     */
    function readPointSpecs(pathPoints) {
        var pointSpecs = [];
        for (var i = 0; i < pathPoints.length; i++) {
            pointSpecs.push({
                anchor: pathPoints[i].anchor,
                leftDirection: pathPoints[i].leftDirection,
                rightDirection: pathPoints[i].rightDirection
            });
        }
        return pointSpecs;
    }

    /**
     * 2 つのベクトルの内積を返す
     * @param {number[]} vectorA - [x, y]
     * @param {number[]} vectorB - [x, y]
     * @returns {number} 内積
     */
    function dotProduct(vectorA, vectorB) {
        return vectorA[0] * vectorB[0] + vectorA[1] * vectorB[1];
    }

    /**
     * 2 つのベクトルの外積（z 成分）を返す
     * @param {number[]} vectorA - [x, y]
     * @param {number[]} vectorB - [x, y]
     * @returns {number} 外積
     */
    function crossProduct(vectorA, vectorB) {
        return vectorA[0] * vectorB[1] - vectorA[1] * vectorB[0];
    }

    /**
     * 単位ベクトルが水平または垂直かを返す
     * @param {number[]} direction - 単位ベクトル
     * @returns {boolean} 軸に沿っていれば true
     */
    function isAxisDirection(direction) {
        return Math.abs(direction[0]) < DIRECTION_TOLERANCE || Math.abs(direction[1]) < DIRECTION_TOLERANCE;
    }

    /**
     * セグメントが指定の向きの直線か（終点とハンドルが、始点を通る向きの直線上にあるか）を返す
     * 角度ではなく直線からの距離で判定する（長い辺でわずかに外れたアンカーを取りこぼさない）
     * @param {Object} startSpec - 始点のポイント
     * @param {Object} endSpec - 終点のポイント
     * @param {number[]} direction - 向きの単位ベクトル
     * @returns {boolean} 直線なら true
     */
    function isStraightAlong(startSpec, endSpec, direction) {
        var startAnchor = startSpec.anchor;
        var chord = [endSpec.anchor[0] - startAnchor[0], endSpec.anchor[1] - startAnchor[1]];
        if (dotProduct(chord, direction) < GEOMETRY_TOLERANCE) return false;
        var checkedPoints = [endSpec.anchor, startSpec.rightDirection, endSpec.leftDirection];
        for (var i = 0; i < checkedPoints.length; i++) {
            var offset = [checkedPoints[i][0] - startAnchor[0], checkedPoints[i][1] - startAnchor[1]];
            if (Math.abs(crossProduct(offset, direction)) > GEOMETRY_TOLERANCE) return false;
        }
        return true;
    }

    /**
     * 曲線セグメントが水平・垂直の辺をつなぐ 90° の角丸かを調べる
     * @param {Object} startSpec - 始点のポイント
     * @param {Object} endSpec - 終点のポイント
     * @returns {Object|null} 角の情報（角丸でなければ null）
     */
    function findArcCorner(startSpec, endSpec) {
        var radius = getSegmentRadius(startSpec, endSpec);
        if (radius === null) return null;
        var startAnchor = startSpec.anchor;
        var endAnchor = endSpec.anchor;
        /* 接線の求め方は getSegmentRadius() と同じ / Tangents are taken as in getSegmentRadius() */
        var inDirection = getUnitVector(startAnchor, startSpec.rightDirection) || getUnitVector(startAnchor, endSpec.leftDirection);
        var outDirection = getUnitVector(endSpec.leftDirection, endAnchor) || getUnitVector(startSpec.rightDirection, endAnchor);
        if (!isAxisDirection(inDirection) || !isAxisDirection(outDirection)) return null;
        if (Math.abs(dotProduct(inDirection, outDirection)) > DIRECTION_TOLERANCE) return null;
        /* 2 本の接線の交点が元の角 / The tangent intersection is the original corner */
        var isHorizontalIn = Math.abs(inDirection[1]) < DIRECTION_TOLERANCE;
        var vertex = isHorizontalIn ? [endAnchor[0], startAnchor[1]] : [startAnchor[0], endAnchor[1]];
        return { vertex: vertex, inDirection: inDirection, outDirection: outDirection, radius: radius };
    }

    /**
     * アンカーが水平・垂直の辺どうしの直角の角かを調べる
     * @param {Object} previousSpec - 前のポイント
     * @param {Object} pointSpec - 調べるポイント
     * @param {Object} nextSpec - 次のポイント
     * @returns {Object|null} 角の情報（直角の角でなければ null）
     */
    function findSquareCorner(previousSpec, pointSpec, nextSpec) {
        var inDirection = getUnitVector(previousSpec.anchor, pointSpec.anchor);
        var outDirection = getUnitVector(pointSpec.anchor, nextSpec.anchor);
        if (!inDirection || !outDirection) return null;
        if (!isAxisDirection(inDirection) || !isAxisDirection(outDirection)) return null;
        if (Math.abs(dotProduct(inDirection, outDirection)) > DIRECTION_TOLERANCE) return null;
        if (!isStraightAlong(previousSpec, pointSpec, inDirection) || !isStraightAlong(pointSpec, nextSpec, outDirection)) return null;
        return { vertex: pointSpec.anchor, inDirection: inDirection, outDirection: outDirection, radius: 0 };
    }

    /**
     * アンカーの並びが反時計回り（y 上向き）かを返す
     * @param {Object[]} pointSpecs - readPointSpecs() の戻り値
     * @returns {boolean} 反時計回りなら true
     */
    function isCounterClockwise(pointSpecs) {
        var signedArea = 0;
        for (var i = 0; i < pointSpecs.length; i++) {
            signedArea += crossProduct(pointSpecs[i].anchor, pointSpecs[(i + 1) % pointSpecs.length].anchor);
        }
        return signedArea > 0;
    }

    /**
     * パスの外向きの角（角丸と直角）をパスの並び順に集める
     * 吹き出しの口の付け根のように内側へ曲がる角は含めない
     * @param {Object[]} pointSpecs - readPointSpecs() の戻り値
     * @returns {Object[]} 角の情報（startIndex・endIndex は角を作るアンカーの番号）
     */
    function findCorners(pointSpecs) {
        var pointCount = pointSpecs.length;
        var convexTurn = isCounterClockwise(pointSpecs);
        var corners = [];
        for (var i = 0; i < pointCount; i++) {
            var previousSpec = pointSpecs[(i - 1 + pointCount) % pointCount];
            var nextIndex = (i + 1) % pointCount;
            var squareCorner = findSquareCorner(previousSpec, pointSpecs[i], pointSpecs[nextIndex]);
            if (squareCorner && isConvexCorner(squareCorner, convexTurn)) {
                squareCorner.startIndex = i;
                squareCorner.endIndex = i;
                corners.push(squareCorner);
                continue;
            }
            var arcCorner = findArcCorner(pointSpecs[i], pointSpecs[nextIndex]);
            if (arcCorner && isConvexCorner(arcCorner, convexTurn)) {
                arcCorner.startIndex = i;
                arcCorner.endIndex = nextIndex;
                corners.push(arcCorner);
            }
        }
        return corners;
    }

    /**
     * 角がパスの回転と同じ向き（外向き）に曲がっているかを返す
     * @param {Object} corner - 角の情報
     * @param {boolean} convexTurn - パスが反時計回りなら true
     * @returns {boolean} 外向きなら true
     */
    function isConvexCorner(corner, convexTurn) {
        return (crossProduct(corner.inDirection, corner.outDirection) > 0) === convexTurn;
    }

    /**
     * 4 つの角に位置（0：左上、1：右上、2：右下、3：左下）を割り当てる
     * 角が長方形の頂点に並んでいなければ false
     * @param {Object[]} corners - findCorners() の戻り値（4 つ）
     * @returns {boolean} 割り当てられたら true
     */
    function assignCornerPositions(corners) {
        var centerX = 0, centerY = 0;
        for (var i = 0; i < corners.length; i++) {
            centerX += corners[i].vertex[0] / corners.length;
            centerY += corners[i].vertex[1] / corners.length;
        }
        var vertexByPosition = [];
        for (var j = 0; j < corners.length; j++) {
            var vertex = corners[j].vertex;
            var isTop = vertex[1] > centerY;
            var isLeft = vertex[0] < centerX;
            var position = isTop ? (isLeft ? 0 : 1) : (isLeft ? 3 : 2);
            if (vertexByPosition[position]) return false;
            vertexByPosition[position] = vertex;
            corners[j].position = position;
        }
        /* 左右の辺が縦に、上下の辺が横にそろうこと / Left and right sides vertical, top and bottom horizontal */
        return Math.abs(vertexByPosition[0][0] - vertexByPosition[3][0]) < GEOMETRY_TOLERANCE &&
            Math.abs(vertexByPosition[1][0] - vertexByPosition[2][0]) < GEOMETRY_TOLERANCE &&
            Math.abs(vertexByPosition[0][1] - vertexByPosition[1][1]) < GEOMETRY_TOLERANCE &&
            Math.abs(vertexByPosition[3][1] - vertexByPosition[2][1]) < GEOMETRY_TOLERANCE;
    }

    /**
     * 隣り合う 2 つの角のあいだの辺を調べ、角丸の上限と残すポイントを求める
     * 辺の上に一直線に並ぶだけのアンカーは作り直しで除き、吹き出しの口のように辺から外れる部分は残す
     * @param {Object[]} pointSpecs - readPointSpecs() の戻り値
     * @param {Object} fromCorner - 辺の始まりの角（outLimit を設定する）
     * @param {Object} toCorner - 辺の終わりの角（inLimit を設定する）
     * @returns {Object[]|null} 残すポイント（辺が直線でつながっていなければ null）
     */
    function measureEdge(pointSpecs, fromCorner, toCorner) {
        var pointCount = pointSpecs.length;
        var edgeDirection = fromCorner.outDirection;
        if (dotProduct(edgeDirection, toCorner.inDirection) < 1 - DIRECTION_TOLERANCE) return null;
        var edgeLength = dotProduct([toCorner.vertex[0] - fromCorner.vertex[0], toCorner.vertex[1] - fromCorner.vertex[1]], edgeDirection);

        /* 隣の角丸とアンカーを共有している / Sharing an anchor with the next rounded corner */
        if (fromCorner.endIndex === toCorner.startIndex) {
            fromCorner.outLimit = edgeLength / 2;
            toCorner.inLimit = edgeLength / 2;
            return [];
        }

        /* 角から角までのアンカー（両端を含む）/ Anchors from corner to corner, both ends included */
        var edgeSpecs = [];
        for (var i = fromCorner.endIndex; ; i = (i + 1) % pointCount) {
            edgeSpecs.push(pointSpecs[i]);
            if (i === toCorner.startIndex) break;
        }
        /* 両端から辺に沿う直線を進め、外れる手前のアンカーまでを残す範囲にする
           Walk the straight run from both ends; what lies between is kept */
        var firstKept = 0;
        while (firstKept < edgeSpecs.length - 1 && isStraightAlong(edgeSpecs[firstKept], edgeSpecs[firstKept + 1], edgeDirection)) firstKept++;
        if (firstKept === edgeSpecs.length - 1) {
            fromCorner.outLimit = edgeLength / 2;
            toCorner.inLimit = edgeLength / 2;
            return [];
        }
        var lastKept = edgeSpecs.length - 1;
        while (lastKept > firstKept && isStraightAlong(edgeSpecs[lastKept - 1], edgeSpecs[lastKept], edgeDirection)) lastKept--;
        /* 角のすぐ先で辺から外れるものは対象外 / Shapes leaving the edge right at a corner are not supported */
        if (firstKept === 0 || lastKept === edgeSpecs.length - 1) return null;

        var firstAnchor = edgeSpecs[firstKept].anchor;
        var lastAnchor = edgeSpecs[lastKept].anchor;
        fromCorner.outLimit = Math.max(0, dotProduct([firstAnchor[0] - fromCorner.vertex[0], firstAnchor[1] - fromCorner.vertex[1]], edgeDirection));
        toCorner.inLimit = Math.max(0, dotProduct([toCorner.vertex[0] - lastAnchor[0], toCorner.vertex[1] - lastAnchor[1]], edgeDirection));
        return edgeSpecs.slice(firstKept, lastKept + 1);
    }

    /**
     * 角丸を変更できる形（水平・垂直の長方形、または辺に吹き出しの口などが付いた長方形）かを調べる
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {Object|null} { corners, keptSpecs, hasOffEdgePoints }（対象外なら null）
     *   corners はパスの並び順、keptSpecs[i] は corners[i] と次の角のあいだに残すポイント
     */
    function analyzeCornerShape(pageItem) {
        if (pageItem.typename !== "PathItem" || !pageItem.closed) return null;
        var pathPoints = pageItem.pathPoints;
        if (pathPoints.length < 4 || pathPoints.length > MAX_POINT_COUNT) return null;

        /* 重なったアンカーは長さ 0 の辺になるので、先にまとめる / Merge overlapping anchors first (they make zero-length sides) */
        var pointSpecs = mergeCoincidentPoints(readPointSpecs(pathPoints));
        var corners = findCorners(pointSpecs);
        if (corners.length !== 4 || !assignCornerPositions(corners)) return null;

        var keptSpecs = [];
        var hasOffEdgePoints = false;
        for (var i = 0; i < corners.length; i++) {
            var edgeSpecs = measureEdge(pointSpecs, corners[i], corners[(i + 1) % corners.length]);
            if (!edgeSpecs) return null;
            if (edgeSpecs.length > 0) hasOffEdgePoints = true;
            keptSpecs.push(edgeSpecs);
        }
        return { corners: corners, keptSpecs: keptSpecs, hasOffEdgePoints: hasOffEdgePoints };
    }

    /**
     * 角丸を変更できる形かを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isRectangularShape(pageItem) {
        return analyzeCornerShape(pageItem) !== null;
    }

    /**
     * 形の 4 隅の角丸半径を返す
     * @param {Object} cornerShape - analyzeCornerShape() の戻り値
     * @returns {number[]} [左上, 右上, 右下, 左下] の半径（pt）
     */
    function getShapeRadii(cornerShape) {
        var cornerRadii = [0, 0, 0, 0];
        for (var i = 0; i < cornerShape.corners.length; i++) {
            cornerRadii[cornerShape.corners[i].position] = cornerShape.corners[i].radius;
        }
        return cornerRadii;
    }

    /**
     * パスの 4 隅の角丸半径を求める（曲線の無い隅は 0）
     * @param {PathItem} pathItem - 対象パス（isRectangularShape() が true のもの）
     * @returns {number[]} [左上, 右上, 右下, 左下] の半径（pt）
     */
    function getCornerRadii(pathItem) {
        var cornerShape = analyzeCornerShape(pathItem);
        return cornerShape ? getShapeRadii(cornerShape) : [0, 0, 0, 0];
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * オブジェクトと親（グループ・レイヤー）がすべてロック解除・表示中かを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 編集できれば true
     */
    function isEditable(pageItem) {
        var ancestorItem = pageItem;
        while (ancestorItem && ancestorItem.typename !== "Document") {
            if (ancestorItem.typename === "Layer") {
                if (ancestorItem.locked || !ancestorItem.visible) return false;
            } else if (ancestorItem.locked || ancestorItem.hidden) {
                return false;
            }
            ancestorItem = ancestorItem.parent;
        }
        return true;
    }

    /**
     * 2 つの範囲が重なるかを返す
     * @param {number[]} itemBounds - [左, 上, 右, 下]
     * @param {number[]} areaBounds - [左, 上, 右, 下]
     * @returns {boolean} 重なっていれば true
     */
    function boundsOverlap(itemBounds, areaBounds) {
        return itemBounds[0] < areaBounds[2] && itemBounds[2] > areaBounds[0] &&
            itemBounds[1] > areaBounds[3] && itemBounds[3] < areaBounds[1];
    }

    /**
     * 複合パスの一部かを返す
     * @param {PathItem} pathItem - 判定するパス
     * @returns {boolean} 複合パスの一部なら true
     */
    function isCompoundMember(pathItem) {
        return pathItem.parent.typename === "CompoundPathItem";
    }

    /**
     * 複合シェイプの中のパスかを返す（ダイレクト選択したときだけ選択に入る）
     * 中のパスは PluginItem 直下の見えないグループに属し、さらにサブグループや入れ子の複合シェイプの中にも置ける
     * @param {PathItem} pathItem - 判定するパス
     * @returns {boolean} 複合シェイプの中のパスなら true
     */
    function isCompoundShapeMember(pathItem) {
        for (var ancestorItem = pathItem.parent; ancestorItem.typename === "GroupItem"; ancestorItem = ancestorItem.parent) {
            if (ancestorItem.parent.typename === "PluginItem") return true;
        }
        return false;
    }

    /**
     * 選択したオブジェクト 1 つから対象の長方形を集める
     * グループは中を再帰でたどり（ロック・非表示の子は除く）、複合パスは中の長方形、複合シェイプは変換結果の中の長方形を対象にする
     * @param {PageItem} pageItem - 選択したオブジェクト、またはグループの子
     * @param {PathItem[]} targetPaths - 見つけたパスを追加する配列
     * @returns {void}
     */
    function collectItemRectangles(pageItem, targetPaths) {
        switch (pageItem.typename) {
            case "GroupItem":
                for (var i = 0; i < pageItem.pageItems.length; i++) {
                    var childItem = pageItem.pageItems[i];
                    if (childItem.locked || childItem.hidden) continue;
                    collectItemRectangles(childItem, targetPaths);
                }
                break;
            case "CompoundPathItem":
                for (var j = 0; j < pageItem.pathItems.length; j++) {
                    if (isRectangularShape(pageItem.pathItems[j])) targetPaths.push(pageItem.pathItems[j]);
                }
                break;
            case "PluginItem":
                var shapePaths = collectShapeRectangles(pageItem);
                for (var k = 0; k < shapePaths.length; k++) targetPaths.push(shapePaths[k]);
                break;
            default:
                if (isRectangularShape(pageItem)) targetPaths.push(pageItem);
        }
    }

    /**
     * 選択から対象の長方形を集める（グループ・複合パス・複合シェイプの中も含む）
     * @param {PageItem[]} selectedItems - 選択
     * @returns {{targetPaths: PathItem[], skippedCount: number}} 対象パスと、対象を含まない選択の数
     */
    function collectSelectedRectangles(selectedItems) {
        var targetPaths = [];
        var skippedCount = 0;
        for (var i = 0; i < selectedItems.length; i++) {
            var foundCount = targetPaths.length;
            collectItemRectangles(selectedItems[i], targetPaths);
            if (targetPaths.length === foundCount) skippedCount++;
        }
        return { targetPaths: targetPaths, skippedCount: skippedCount };
    }

    /**
     * ドキュメント内の編集できる長方形を集める（グループ・複合パスの中と、変換した複合シェイプの中も含む。ガイドは除く）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} [areaBounds] - 指定したときは、この範囲に重なるものだけ
     * @returns {PathItem[]} 対象パス
     */
    function collectDocumentRectangles(doc, areaBounds) {
        var targetPaths = [];
        var pathItems = doc.pathItems;
        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            if (pathItem.guides) continue;
            if (!isRectangularShape(pathItem) || !isEditable(pathItem)) continue;
            if (areaBounds && !boundsOverlap(pathItem.geometricBounds, areaBounds)) continue;
            targetPaths.push(pathItem);
        }
        /* 複合シェイプは変換すると PluginItem が増減するので、先に一覧を写す
           Converting adds and removes PluginItems, so snapshot the list first */
        var pluginItems = [];
        for (var k = 0; k < doc.pluginItems.length; k++) pluginItems.push(doc.pluginItems[k]);
        for (var m = 0; m < pluginItems.length; m++) {
            var pluginItem = pluginItems[m];
            if (!isEditable(pluginItem)) continue;
            if (areaBounds && !boundsOverlap(pluginItem.geometricBounds, areaBounds)) continue;
            targetPaths = targetPaths.concat(collectShapeRectangles(pluginItem));
        }
        return targetPaths;
    }

    // =========================================
    // パスの作り直し / Path rebuild
    // =========================================

    /**
     * 角を指定の半径にしたポイントを返す（半径は隣の角・吹き出しの口までの距離で制限する）
     * @param {Object} corner - analyzeCornerShape() の角
     * @param {number} requestedRadius - 半径（pt）
     * @returns {Object[]} { anchor, leftDirection, rightDirection } の配列（1 つか 2 つ）
     */
    function buildCornerPoints(corner, requestedRadius) {
        var vertex = corner.vertex;
        var incoming = corner.inDirection;
        var outgoing = corner.outDirection;
        var radius = Math.min(Math.max(requestedRadius, 0), corner.inLimit, corner.outLimit);
        if (radius < GEOMETRY_TOLERANCE) {
            return [{ anchor: vertex, leftDirection: vertex, rightDirection: vertex }];
        }
        var handleLength = radius * ARC_HANDLE_RATIO;
        var arcStart = [vertex[0] - incoming[0] * radius, vertex[1] - incoming[1] * radius];
        var arcEnd = [vertex[0] + outgoing[0] * radius, vertex[1] + outgoing[1] * radius];
        return [{
            anchor: arcStart,
            leftDirection: arcStart,
            rightDirection: [arcStart[0] + incoming[0] * handleLength, arcStart[1] + incoming[1] * handleLength]
        }, {
            anchor: arcEnd,
            leftDirection: [arcEnd[0] - outgoing[0] * handleLength, arcEnd[1] - outgoing[1] * handleLength],
            rightDirection: arcEnd
        }];
    }

    /**
     * 同じ位置に重なった隣り合うアンカーを 1 つにまとめる（半径が上限に達したときに生じる）
     * まとめたアンカーは、前のポイントの左ハンドルと後ろのポイントの右ハンドルを持つ
     * @param {Object[]} pointSpecs - { anchor, leftDirection, rightDirection } の配列
     * @returns {Object[]} 重なりを除いたポイント
     */
    function mergeCoincidentPoints(pointSpecs) {
        var mergedSpecs = [];
        for (var i = 0; i < pointSpecs.length; i++) {
            var previousSpec = mergedSpecs[mergedSpecs.length - 1];
            if (previousSpec && getDistance(previousSpec.anchor, pointSpecs[i].anchor) < GEOMETRY_TOLERANCE) {
                previousSpec.rightDirection = pointSpecs[i].rightDirection;
            } else {
                mergedSpecs.push(pointSpecs[i]);
            }
        }
        /* 末尾と先頭の重なり / Overlap between the last and first points */
        var lastSpec = mergedSpecs[mergedSpecs.length - 1];
        if (mergedSpecs.length > 1 && getDistance(lastSpec.anchor, mergedSpecs[0].anchor) < GEOMETRY_TOLERANCE) {
            mergedSpecs[0].leftDirection = lastSpec.leftDirection;
            mergedSpecs.pop();
        }
        return mergedSpecs;
    }

    /**
     * パスの角を指定の半径に作り直す（アンカーの並びと、吹き出しの口など辺から外れる部分は元のまま）
     * @param {PathItem} pathItem - 対象パス（isRectangularShape() が true のもの）
     * @param {number[]} cornerRadii - [左上, 右上, 右下, 左下] の半径（pt）
     * @returns {void}
     */
    function rebuildPath(pathItem, cornerRadii) {
        var cornerShape = analyzeCornerShape(pathItem);
        if (!cornerShape) return;
        var pointSpecs = [];
        for (var c = 0; c < cornerShape.corners.length; c++) {
            var corner = cornerShape.corners[c];
            pointSpecs = pointSpecs.concat(buildCornerPoints(corner, cornerRadii[corner.position]), cornerShape.keptSpecs[c]);
        }
        pointSpecs = mergeCoincidentPoints(pointSpecs);

        var anchors = [];
        for (var i = 0; i < pointSpecs.length; i++) anchors.push(pointSpecs[i].anchor);
        pathItem.setEntirePath(anchors);

        for (var j = 0; j < pointSpecs.length; j++) {
            var pathPoint = pathItem.pathPoints[j];
            pathPoint.leftDirection = pointSpecs[j].leftDirection;
            pathPoint.rightDirection = pointSpecs[j].rightDirection;
        }
    }

    /**
     * 角丸を指定の半径にする
     * - 効果で角丸になっている長方形：アピアランスを消去して効果を付け直す
     * - ［「効果」に変換］がオンで 4 つの角が角丸の長方形：角のない長方形に戻して効果を付ける
     * - それ以外（複合パス・複合シェイプの一部、吹き出しなどは常にこちら）：パスを作り直す（［0の半径は0のままに］なら角丸の無い角は残す）
     * @param {PathItem} pathItem - 対象パス
     * @param {Object} measurement - measureCorners() の戻り値
     * @param {number} cornerRadius - 半径（pt）
     * @param {Object} cornerOptions - { keepZeroRadii, convertToEffect }
     * @returns {PathItem} 処理後のパス（アピアランスの消去で作り直されたときは新しい参照）
     */
    function applyCornerRadius(pathItem, measurement, cornerRadius, cornerOptions) {
        if (measurement.effectRadii) {
            /* 消去でオブジェクトが作り直され、元の参照が無効になることがある
               Clearing may recreate the object and invalidate the old reference */
            var clearedPath = clearAppearance(pathItem);
            applyRoundCornersEffect(clearedPath, cornerRadius);
            return clearedPath;
        }
        /* ［0の半径は0のままに］がオンのときは変換しない（ダイアログでもディム）
           No conversion while "keep zero radii" is on (dimmed in the dialog) */
        if (cornerOptions.convertToEffect && !cornerOptions.keepZeroRadii && measurement.canUseEffect &&
            countRoundedCorners(measurement.pathRadii) === measurement.pathRadii.length) {
            rebuildPath(pathItem, [0, 0, 0, 0]);
            applyRoundCornersEffect(pathItem, cornerRadius);
            return pathItem;
        }
        var cornerRadii = [];
        for (var i = 0; i < measurement.pathRadii.length; i++) {
            var isKept = cornerOptions.keepZeroRadii && measurement.pathRadii[i] < GEOMETRY_TOLERANCE;
            cornerRadii.push(isKept ? 0 : cornerRadius);
        }
        rebuildPath(pathItem, cornerRadii);
        return pathItem;
    }

    // =========================================
    // 「角を丸くする」効果 / Round Corners effect
    // =========================================

    /**
     * 「角を丸くする」効果の LiveEffect XML を返す
     * @param {number} radius - 半径（pt）
     * @returns {string} applyEffect() に渡す XML
     */
    function buildRoundCornersXml(radius) {
        return '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>';
    }

    /**
     * 「角を丸くする」効果を付ける（半径は短辺の半分まで、0 なら付けない）
     * @param {PathItem} pathItem - 対象パス
     * @param {number} cornerRadius - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(pathItem, cornerRadius) {
        if (cornerRadius < GEOMETRY_TOLERANCE) return;
        pathItem.applyEffect(buildRoundCornersXml(Math.min(cornerRadius, getMaxRadius(pathItem.geometricBounds))));
    }

    /**
     * カラー値を複製する
     * @param {Object} sourceColor - 複製元のカラー
     * @returns {Object|null} 複製したカラー
     */
    function cloneColorValue(sourceColor) {
        if (!sourceColor) return null;
        var clonedColor;
        switch (sourceColor.typename) {
            case "RGBColor":
                clonedColor = new RGBColor();
                clonedColor.red = sourceColor.red;
                clonedColor.green = sourceColor.green;
                clonedColor.blue = sourceColor.blue;
                return clonedColor;
            case "CMYKColor":
                clonedColor = new CMYKColor();
                clonedColor.cyan = sourceColor.cyan;
                clonedColor.magenta = sourceColor.magenta;
                clonedColor.yellow = sourceColor.yellow;
                clonedColor.black = sourceColor.black;
                return clonedColor;
            case "GrayColor":
                clonedColor = new GrayColor();
                clonedColor.gray = sourceColor.gray;
                return clonedColor;
            case "SpotColor":
                clonedColor = new SpotColor();
                clonedColor.spot = sourceColor.spot;
                clonedColor.tint = sourceColor.tint;
                return clonedColor;
            case "PatternColor":
                clonedColor = new PatternColor();
                clonedColor.pattern = sourceColor.pattern;
                return clonedColor;
            case "GradientColor":
                clonedColor = new GradientColor();
                clonedColor.gradient = sourceColor.gradient;
                clonedColor.angle = sourceColor.angle;
                clonedColor.length = sourceColor.length;
                clonedColor.matrix = sourceColor.matrix;
                clonedColor.origin = sourceColor.origin;
                clonedColor.hiliteAngle = sourceColor.hiliteAngle;
                clonedColor.hiliteLength = sourceColor.hiliteLength;
                return clonedColor;
            case "NoColor":
                return new NoColor();
            default:
                return sourceColor;
        }
    }

    /**
     * パスの基本の塗り・線・不透明度を控える
     * @param {PathItem} pathItem - 対象パス
     * @returns {Object} 控えたスタイル
     */
    function capturePathStyle(pathItem) {
        return {
            filled: pathItem.filled,
            fillColor: pathItem.filled ? cloneColorValue(pathItem.fillColor) : null,
            stroked: pathItem.stroked,
            strokeColor: pathItem.stroked ? cloneColorValue(pathItem.strokeColor) : null,
            strokeWidth: pathItem.strokeWidth,
            strokeDashes: pathItem.strokeDashes,
            strokeDashOffset: pathItem.strokeDashOffset,
            strokeCap: pathItem.strokeCap,
            strokeJoin: pathItem.strokeJoin,
            strokeMiterLimit: pathItem.strokeMiterLimit,
            opacity: pathItem.opacity,
            blendingMode: pathItem.blendingMode
        };
    }

    /**
     * 控えたスタイルをパスに戻す
     * @param {PathItem} pathItem - 対象パス
     * @param {Object} pathStyle - capturePathStyle() の戻り値
     * @returns {void}
     */
    function restorePathStyle(pathItem, pathStyle) {
        pathItem.opacity = pathStyle.opacity;
        pathItem.blendingMode = pathStyle.blendingMode;
        pathItem.filled = pathStyle.filled;
        if (pathStyle.filled && pathStyle.fillColor) pathItem.fillColor = cloneColorValue(pathStyle.fillColor);
        pathItem.stroked = pathStyle.stroked;
        if (pathStyle.stroked && pathStyle.strokeColor) {
            pathItem.strokeColor = cloneColorValue(pathStyle.strokeColor);
            pathItem.strokeWidth = pathStyle.strokeWidth;
            pathItem.strokeDashes = pathStyle.strokeDashes;
            pathItem.strokeDashOffset = pathStyle.strokeDashOffset;
            pathItem.strokeCap = pathStyle.strokeCap;
            pathItem.strokeJoin = pathStyle.strokeJoin;
            pathItem.strokeMiterLimit = pathStyle.strokeMiterLimit;
        }
    }

    /* ［アピアランスを消去］のダイナミックアクション（セット「Appearance」／アクション「clear」）
       Dynamic action for Clear Appearance (set "Appearance", action "clear") */
    var CLEAR_APPEARANCE_ACTION = [
        "/version 3",
        "/name [ 10",
        " 417070656172616e6365",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        " /name [ 5",
        " 636c656172",
        " ]",
        " /keyIndex 0",
        " /colorIndex 0",
        " /isOpen 1",
        " /eventCount 1",
        " /event-1 {",
        " /useRulersIn1stQuadrant 0",
        " /internalName (ai_plugin_appearance)",
        " /localizedName [ 18",
        " e382a2e38394e382a2e383a9e383b3e382b9",
        " ]",
        " /isOpen 1",
        " /isOn 1",
        " /hasDialog 0",
        " /parameterCount 1",
        " /parameter-1 {",
        " /key 1835363957",
        " /showInPalette 4294967295",
        " /type (enumerated)",
        " /name [ 27",
        " e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb",
        " ]",
        " /value 6",
        " }",
        " }",
        "}"
    ].join("\n");

    /* ［複合シェイプを解除］のダイナミックアクション（セット「CompoundShape」／アクション「Release」）
       複合シェイプにだけ効き、ブレンドやエンベロープは解除しない
       Dynamic action for Release Compound Shape; it affects compound shapes only, not blends or envelopes */
    var RELEASE_COMPOUND_SHAPE_ACTION = [
        "/version 3",
        "/name [ 13",
        " 436f6d706f756e645368617065",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        " /name [ 7",
        " 52656c65617365",
        " ]",
        " /keyIndex 0",
        " /colorIndex 0",
        " /isOpen 1",
        " /eventCount 1",
        " /event-1 {",
        " /useRulersIn1stQuadrant 0",
        " /internalName (ai_release_compound_shape)",
        " /localizedName [ 27",
        " e8a487e59088e382b7e382a7e382a4e38397e38292e8a7a3e999a4",
        " ]",
        " /isOpen 0",
        " /isOn 1",
        " /hasDialog 0",
        " /parameterCount 1",
        " /parameter-1 {",
        " /key 1919710053",
        " /showInPalette 4294967295",
        " /type (integer)",
        " /value 0",
        " }",
        " }",
        "}"
    ].join("\n");

    /**
     * オブジェクトだけを選択する（メニューコマンドやアクションの対象にする）
     * @param {PageItem} pageItem - 対象オブジェクト（ロック・非表示でないこと）
     * @returns {void}
     */
    function selectOnly(pageItem) {
        var doc = app.activeDocument;
        doc.selection = null;
        /* グループへ移したばかりのオブジェクトなどは、null の代入で外れずに残ることがある
           Items just moved into a group can stay selected after assigning null */
        var remainingItems = doc.selection;
        for (var i = remainingItems.length - 1; i >= 0; i--) remainingItems[i].selected = false;
        pageItem.selected = true;
    }

    /**
     * 選択の先頭を返す（何も選択されていなければ null）
     * @returns {PageItem|null} 選択の先頭
     */
    function getFirstSelectedItem() {
        var currentSelection = app.activeDocument.selection;
        return (currentSelection.length > 0) ? currentSelection[0] : null;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は runTemporaryAction / loadTemporaryActionSet / unloadTemporaryActionSet / toActionHex / buildActionNameLines
    // 2. アクション定義は配列＋join("\n") で組み立てる（''' は ES3 の構文エラー）。
    //    セット名・アクション名は英数字にする。/name [ n 16進 ] は buildActionNameLines で作るとバイト数がずれない
    //      var actionSource = [
    //          "/version 3"
    //      ].concat(buildActionNameLines("", "MySet"), [
    //          "/isOpen 1", "/actionCount 1", "/action-1 {"
    //      ], buildActionNameLines("\t", "myAction"), [ … ]).join("\n");
    // 3. 1回だけ実行するとき:
    //      if (!runTemporaryAction(actionSource, "MySet", "myAction")) alert(getLabel("alert.actionFailed"));
    //    何度も実行するとき（オブジェクトごとなど）は、読み込み・解除を1回ずつにする:
    //      if (!loadTemporaryActionSet(actionSource, "MySet")) { alert(…); return; }
    //      try { for (…) app.doScript("myAction", "MySet"); } finally { unloadTemporaryActionSet("MySet"); }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。警告を出すかはコピー先で決める
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * オブジェクトだけを選択して、ダイナミックアクションを読み込んで実行する（実行後は読み込みを外す）
     * @param {string} actionSource - アクションの定義（.aia の内容）
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {void}
     */
    function runDynamicAction(actionSource, setName, actionName, pageItem) {
        selectOnly(pageItem);
        /* 失敗は従来どおり例外で呼び出し元へ伝える / Report a failure to the caller as an exception, as before */
        if (!runTemporaryAction(actionSource, setName, actionName)) {
            throw new Error("Could not run the action: " + setName + " / " + actionName);
        }
    }

    /**
     * ダイナミックアクションで［アピアランスを消去］を実行し、基本の塗り・線・不透明度を戻す
     * （効果を外すためのメニューコマンドは無いのでアクションで実行する）
     * @param {PathItem} pathItem - 対象パス
     * @returns {PathItem} 処理後のパス（作り直されたときは新しい参照）
     */
    function clearAppearance(pathItem) {
        var pathStyle = capturePathStyle(pathItem);
        runDynamicAction(CLEAR_APPEARANCE_ACTION, "Appearance", "clear", pathItem);
        var clearedPath = findFirstPathItem(getFirstSelectedItem()) || pathItem;
        restorePathStyle(clearedPath, pathStyle);
        return clearedPath;
    }

    /**
     * オブジェクトだけを選択してメニューコマンドを実行し、実行後の選択の先頭を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {string} commandName - executeMenuCommand のコマンド名
     * @returns {PageItem|null} 実行後に選択されているオブジェクト
     */
    function runMenuCommand(pageItem, commandName) {
        selectOnly(pageItem);
        app.redraw();
        app.executeMenuCommand(commandName);
        return getFirstSelectedItem();
    }

    /**
     * グループの中から最初のパスを探す
     * @param {PageItem} pageItem - 探す対象
     * @returns {PathItem|null} 見つかったパス
     */
    function findFirstPathItem(pageItem) {
        if (!pageItem) return null;
        if (pageItem.typename === "PathItem") return pageItem;
        if (pageItem.typename !== "GroupItem") return null;
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            var foundPath = findFirstPathItem(pageItem.pageItems[i]);
            if (foundPath) return foundPath;
        }
        return null;
    }

    /**
     * 効果を含めた 4 隅の角丸半径を求める（複製のアピアランスを分割して計測する）
     * @param {PathItem} pathItem - 対象パス（非表示でないこと）
     * @returns {number[]} [左上, 右上, 右下, 左下] の半径（pt）
     */
    function getEffectiveCornerRadii(pathItem) {
        var measureCopy = pathItem.duplicate();
        var expandedItem = runMenuCommand(measureCopy, "expandStyle");
        var measuredPath = findFirstPathItem(expandedItem);
        var cornerRadii = (measuredPath && isRectangularShape(measuredPath))
            ? getCornerRadii(measuredPath) : getCornerRadii(pathItem);
        if (expandedItem) {
            expandedItem.remove();
        } else {
            measureCopy.remove();
        }
        return cornerRadii;
    }

    /**
     * パスの角丸を計測する
     * 効果を含めるときは、パスに角丸が無いものだけ複製を分割して効果の角丸を調べる（複合パス・複合シェイプの一部、吹き出しなどは除く）
     * @param {PathItem} pathItem - 対象パス（非表示でないこと）
     * @param {boolean} includeEffect - true なら「角を丸くする」効果も調べる
     * @returns {{pathRadii: number[], effectRadii: number[]|null, canUseEffect: boolean}}
     *   パスの半径、効果による半径（無ければ null）、「角を丸くする」効果を扱えるか
     */
    function measureCorners(pathItem, includeEffect) {
        var cornerShape = analyzeCornerShape(pathItem);
        var pathRadii = getShapeRadii(cornerShape);
        /* 効果は複合パス・複合シェイプ全体に付き、吹き出しの口まで丸めるので、いずれも効果では扱わない
           Effects apply to the whole compound path or shape and would round a callout tail too */
        var canUseEffect = !isCompoundMember(pathItem) && !isCompoundShapeMember(pathItem) &&
            !findShapeConversion(pathItem) && !cornerShape.hasOffEdgePoints;
        var effectRadii = null;
        if (includeEffect && canUseEffect && countRoundedCorners(pathRadii) === 0) {
            var expandedRadii = getEffectiveCornerRadii(pathItem);
            if (countRoundedCorners(expandedRadii) > 0) effectRadii = expandedRadii;
        }
        return { pathRadii: pathRadii, effectRadii: effectRadii, canUseEffect: canUseEffect };
    }

    // =========================================
    // 複合シェイプの変換 / Compound shape conversion
    // =========================================

    /* 複合シェイプの中のパスは PluginItem からたどれないので、複製を解除して「グループ＋［パスファインダー：合体］」に
       変換し、そのグループの中の長方形を対象にする。変換結果は隠しておき、OK で元と差し替える（キャンセルなら捨てる）
       Members of a compound shape cannot be reached from the PluginItem, so a copy is released and rebuilt as
       a group with the Pathfinder Add effect. The result stays hidden and replaces the original on OK */

    /* { original: PluginItem, converted: GroupItem|null, isShown: boolean, isUsed: boolean } の配列
       converted が null のものは複合シェイプではなかった / converted is null when it was not a compound shape */
    var shapeConversions = [];

    /* プレビュー中に表示している変換結果 / Conversions shown while previewing */
    var shownConversions = [];

    /**
     * 現在の選択を配列に写す（選択は後の操作で変わるため）
     * @returns {PageItem[]} 選択の写し
     */
    function copySelection() {
        var currentSelection = app.activeDocument.selection;
        var selectionCopy = [];
        for (var i = 0; i < currentSelection.length; i++) selectionCopy.push(currentSelection[i]);
        return selectionCopy;
    }

    /**
     * オブジェクトをグループにまとめ、［パスファインダー：合体］効果を付ける（中の複合シェイプも変換する）
     * @param {PageItem[]} pageItems - まとめるオブジェクト（前面から背面の順）
     * @returns {GroupItem} 作ったグループ
     */
    function groupWithPathfinderAdd(pageItems) {
        var shapeGroup = pageItems[0].parent.groupItems.add();
        shapeGroup.move(pageItems[0], ElementPlacement.PLACEBEFORE);
        /* 前面から順に末尾へ入れて重なり順を保つ / Append front to back to keep the stacking order */
        for (var i = 0; i < pageItems.length; i++) pageItems[i].move(shapeGroup, ElementPlacement.PLACEATEND);

        /* 入れ子の複合シェイプは 1 段ずつしか解除されないので、中でも変換する / Nested compound shapes release one level at a time */
        var nestedItems = [];
        for (var j = 0; j < shapeGroup.pageItems.length; j++) {
            if (shapeGroup.pageItems[j].typename === "PluginItem") nestedItems.push(shapeGroup.pageItems[j]);
        }
        for (var k = 0; k < nestedItems.length; k++) {
            var nestedGroup = convertCompoundShapeCopy(nestedItems[k]);
            if (!nestedGroup) continue;
            nestedGroup.move(nestedItems[k], ElementPlacement.PLACEBEFORE);
            nestedItems[k].remove();
        }

        bakeRoundCornersEffects(shapeGroup);

        /* applyEffect() だとアピアランスの「内容」の下に入って図形が消えるので、メニューコマンドで付ける
           applyEffect() puts the effect below Contents and the art vanishes, so use the menu command */
        runMenuCommand(shapeGroup, "Live Pathfinder Add");
        return shapeGroup;
    }

    /**
     * グループの中の長方形に「角を丸くする」効果が付いていれば外し、同じ半径をパスの角丸にする
     * 効果の有無は、複製のアピアランスを分割した半径とパスの半径の違いで判定する
     * （外すのは［アピアランスを消去］なので、そのパスのほかの効果も外れる。塗り・線・不透明度は戻す）
     * @param {GroupItem} shapeGroup - 解除した複合シェイプをまとめたグループ
     * @returns {void}
     */
    function bakeRoundCornersEffects(shapeGroup) {
        var rectanglePaths = [];
        collectGroupRectangles(shapeGroup, rectanglePaths);
        for (var i = 0; i < rectanglePaths.length; i++) {
            var pathRadii = getCornerRadii(rectanglePaths[i]);
            var effectRadii = getEffectiveCornerRadii(rectanglePaths[i]);
            var hasRoundCornersEffect = false;
            for (var j = 0; j < pathRadii.length; j++) {
                if (Math.abs(effectRadii[j] - pathRadii[j]) >= GEOMETRY_TOLERANCE) hasRoundCornersEffect = true;
            }
            if (!hasRoundCornersEffect) continue;
            /* 消去でパスが作り直されることがあるので、戻り値に作り直す / Clearing may recreate the path, so rebuild the returned one */
            rebuildPath(clearAppearance(rectanglePaths[i]), effectRadii);
        }
    }

    /**
     * 複合シェイプなら、その複製を「グループ＋［パスファインダー：合体］」に変換して返す（元は変更しない）
     * @param {PluginItem} pluginItem - 調べるオブジェクト
     * @returns {GroupItem|null} 変換したグループ（複合シェイプでなければ null）
     */
    function convertCompoundShapeCopy(pluginItem) {
        var shapeCopy = pluginItem.duplicate();
        runDynamicAction(RELEASE_COMPOUND_SHAPE_ACTION, "CompoundShape", "Release", shapeCopy);
        var releasedItems = copySelection();
        /* 解除されなければ複製が選択に残る。エンベロープは MeshItem と 2 つで選択に入るので、数では判定しない
           An unreleased copy stays selected; envelopes select as two items, so do not judge by count */
        for (var i = 0; i < releasedItems.length; i++) {
            if (releasedItems[i] === shapeCopy) {
                shapeCopy.remove();
                return null;
            }
        }
        if (releasedItems.length === 0) return null;
        return groupWithPathfinderAdd(releasedItems);
    }

    /**
     * PluginItem の変換結果を返す（初回だけ変換し、結果は隠しておく）
     * @param {PluginItem} pluginItem - 対象オブジェクト
     * @returns {Object} shapeConversions の要素
     */
    function getShapeConversion(pluginItem) {
        for (var i = 0; i < shapeConversions.length; i++) {
            if (shapeConversions[i].original === pluginItem) return shapeConversions[i];
        }
        var convertedGroup = convertCompoundShapeCopy(pluginItem);
        if (convertedGroup) convertedGroup.hidden = true;
        var shapeConversion = { original: pluginItem, converted: convertedGroup, isShown: false, isUsed: false };
        shapeConversions.push(shapeConversion);
        return shapeConversion;
    }

    /**
     * パスが属する変換結果を返す
     * @param {PathItem} pathItem - 調べるパス
     * @returns {Object|null} shapeConversions の要素（変換結果の中でなければ null）
     */
    function findShapeConversion(pathItem) {
        if (shapeConversions.length === 0) return null;
        for (var ancestorItem = pathItem.parent; ancestorItem.typename === "GroupItem" ||
            ancestorItem.typename === "CompoundPathItem"; ancestorItem = ancestorItem.parent) {
            for (var i = 0; i < shapeConversions.length; i++) {
                if (shapeConversions[i].converted === ancestorItem) return shapeConversions[i];
            }
        }
        return null;
    }

    /**
     * グループの中の長方形を集める（サブグループ・複合パスの中も含む）
     * @param {GroupItem} containerGroup - 探すグループ
     * @param {PathItem[]} targetPaths - 見つけたパスを追加する配列
     * @returns {void}
     */
    function collectGroupRectangles(containerGroup, targetPaths) {
        for (var i = 0; i < containerGroup.pageItems.length; i++) {
            var childItem = containerGroup.pageItems[i];
            if (childItem.typename === "GroupItem") {
                collectGroupRectangles(childItem, targetPaths);
            } else if (childItem.typename === "CompoundPathItem") {
                for (var j = 0; j < childItem.pathItems.length; j++) {
                    if (isRectangularShape(childItem.pathItems[j])) targetPaths.push(childItem.pathItems[j]);
                }
            } else if (isRectangularShape(childItem)) {
                targetPaths.push(childItem);
            }
        }
    }

    /**
     * 複合シェイプを変換し、中の長方形を返す（複合シェイプでなければ空）
     * @param {PluginItem} pluginItem - 対象オブジェクト
     * @returns {PathItem[]} 変換結果の中の長方形
     */
    function collectShapeRectangles(pluginItem) {
        var shapeConversion = getShapeConversion(pluginItem);
        var targetPaths = [];
        if (shapeConversion.converted) collectGroupRectangles(shapeConversion.converted, targetPaths);
        return targetPaths;
    }

    /**
     * プレビュー用に、元の複合シェイプを隠して変換結果を表示する
     * @param {Object} shapeConversion - shapeConversions の要素
     * @returns {void}
     */
    function showShapeConversion(shapeConversion) {
        if (shapeConversion.isShown) return;
        shapeConversion.original.hidden = true;
        hiddenOriginals.push(shapeConversion.original);
        shapeConversion.converted.hidden = false;
        shapeConversion.isShown = true;
        shownConversions.push(shapeConversion);
    }

    /**
     * 使った変換結果で元の複合シェイプを差し替え、使わなかったものは捨てる
     * @param {PathItem[]} targetPaths - 変更したパス
     * @param {PageItem[]} restoredSelection - 戻す選択（差し替えた元は変換結果に置き換える）
     * @returns {void}
     */
    function finishShapeConversions(targetPaths, restoredSelection) {
        for (var i = 0; i < targetPaths.length; i++) {
            var shapeConversion = findShapeConversion(targetPaths[i]);
            if (shapeConversion) shapeConversion.isUsed = true;
        }
        for (var j = 0; j < shapeConversions.length; j++) {
            var original = shapeConversions[j].original;
            var convertedGroup = shapeConversions[j].converted;
            if (!convertedGroup) continue;
            if (!shapeConversions[j].isUsed) {
                convertedGroup.remove();
                continue;
            }
            convertedGroup.hidden = false;
            convertedGroup.move(original, ElementPlacement.PLACEBEFORE);
            convertedGroup.name = original.name;
            convertedGroup.opacity = original.opacity;
            convertedGroup.blendingMode = original.blendingMode;
            for (var k = 0; k < restoredSelection.length; k++) {
                if (restoredSelection[k] === original) restoredSelection[k] = convertedGroup;
            }
            original.remove();
        }
        shapeConversions = [];
    }

    /**
     * 変換結果をすべて捨てる（キャンセル時）
     * @returns {void}
     */
    function discardShapeConversions() {
        for (var i = 0; i < shapeConversions.length; i++) {
            if (shapeConversions[i].converted) shapeConversions[i].converted.remove();
        }
        shapeConversions = [];
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* プレビュー用の複製と、一時的に隠した元のパス（複合パスの一部は複合パスごと）
       Preview copies and temporarily hidden originals (whole compound paths for their members) */
    var previewCopies = [];
    var hiddenOriginals = [];

    /* 直接書き換えた複合シェイプの中のパスと、書き換え前のポイント
       Compound-shape members edited in place, with their points before the edit */
    var editedOriginals = [];

    /**
     * パスのポイントを種類ごと控える
     * @param {PathItem} pathItem - 対象パス
     * @returns {Object[]} { anchor, leftDirection, rightDirection, pointType } の配列
     */
    function capturePathPoints(pathItem) {
        var savedPoints = [];
        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            var pathPoint = pathItem.pathPoints[i];
            savedPoints.push({
                anchor: pathPoint.anchor,
                leftDirection: pathPoint.leftDirection,
                rightDirection: pathPoint.rightDirection,
                pointType: pathPoint.pointType
            });
        }
        return savedPoints;
    }

    /**
     * 控えたポイントをパスに書き戻す
     * @param {PathItem} pathItem - 対象パス
     * @param {Object[]} savedPoints - capturePathPoints() の戻り値
     * @returns {void}
     */
    function restorePathPoints(pathItem, savedPoints) {
        var anchors = [];
        for (var i = 0; i < savedPoints.length; i++) anchors.push(savedPoints[i].anchor);
        pathItem.setEntirePath(anchors);
        for (var j = 0; j < savedPoints.length; j++) {
            var pathPoint = pathItem.pathPoints[j];
            /* 種類を先に入れる（後から入れるとハンドルが動くことがある）/ Set the type first; setting it later may move handles */
            pathPoint.pointType = savedPoints[j].pointType;
            pathPoint.leftDirection = savedPoints[j].leftDirection;
            pathPoint.rightDirection = savedPoints[j].rightDirection;
        }
    }

    /**
     * プレビューを消して、隠した元のパスを表示に戻し、直接書き換えたパスを元に戻す
     * @returns {void}
     */
    function clearPreview() {
        /* 複製の削除より先に元を表示に戻す（途中で止まっても元が隠れたままにならない）
           Unhide the originals first so they never stay hidden if a removal fails */
        for (var j = 0; j < hiddenOriginals.length; j++) hiddenOriginals[j].hidden = false;
        /* 同じパスを重ねて控えたときも最初の状態に戻るよう、後から順に戻す
           Restore in reverse so a path captured twice ends at its first state */
        for (var k = editedOriginals.length - 1; k >= 0; k--) {
            restorePathPoints(editedOriginals[k].pathItem, editedOriginals[k].savedPoints);
        }
        for (var m = 0; m < shownConversions.length; m++) {
            shownConversions[m].converted.hidden = true;
            shownConversions[m].isShown = false;
        }
        shownConversions = [];
        for (var i = 0; i < previewCopies.length; i++) previewCopies[i].remove();
        previewCopies = [];
        hiddenOriginals = [];
        editedOriginals = [];
    }

    /**
     * 複合パスの中でのパスの番号を返す
     * @param {CompoundPathItem} compoundPath - 複合パス
     * @param {PathItem} pathItem - 中のパス
     * @returns {number} 番号（見つからなければ -1）
     */
    function getMemberIndex(compoundPath, pathItem) {
        for (var i = 0; i < compoundPath.pathItems.length; i++) {
            if (compoundPath.pathItems[i] === pathItem) return i;
        }
        return -1;
    }

    /**
     * 複製に半径を適用してプレビューを表示する（元のパスは一時的に隠す。前のプレビューは消してから呼ぶ）
     * 複合パスの一部は、複合パスごと複製して中の同じ番号のパスに適用する
     * 複合シェイプの中のパスは、複製すると形の一部になるので、元を控えて直接書き換える
     * 変換した複合シェイプは、元を隠して変換結果を表示し、中のパスを控えて直接書き換える
     * @param {PathItem[]} targetPaths - 対象パス
     * @param {Object[]} measurements - パスごとの計測結果（targetPaths と同じ並び）
     * @param {number} cornerRadius - 半径（pt）
     * @param {Object} cornerOptions - { keepZeroRadii, convertToEffect }
     * @returns {void}
     */
    function showPreview(targetPaths, measurements, cornerRadius, cornerOptions) {
        /* 複製済みの複合パスと、その複製 / Compound paths already copied, and their copies */
        var copiedCompounds = [];
        var compoundCopies = [];
        for (var i = 0; i < targetPaths.length; i++) {
            var targetPath = targetPaths[i];
            var shapeConversion = findShapeConversion(targetPath);
            if (shapeConversion) showShapeConversion(shapeConversion);
            if (shapeConversion || isCompoundShapeMember(targetPath)) {
                editedOriginals.push({ pathItem: targetPath, savedPoints: capturePathPoints(targetPath) });
                applyCornerRadius(targetPath, measurements[i], cornerRadius, cornerOptions);
                continue;
            }
            if (!isCompoundMember(targetPath)) {
                /* duplicate() は hidden を引き継ぐので、隠す前に複製する / duplicate() inherits hidden, so copy first */
                var previewCopy = targetPath.duplicate();
                targetPath.hidden = true;
                hiddenOriginals.push(targetPath);
                previewCopies.push(applyCornerRadius(previewCopy, measurements[i], cornerRadius, cornerOptions));
                continue;
            }
            var compoundPath = targetPath.parent;
            var compoundCopy = null;
            for (var j = 0; j < copiedCompounds.length; j++) {
                if (copiedCompounds[j] === compoundPath) compoundCopy = compoundCopies[j];
            }
            if (!compoundCopy) {
                compoundCopy = compoundPath.duplicate();
                compoundPath.hidden = true;
                hiddenOriginals.push(compoundPath);
                previewCopies.push(compoundCopy);
                copiedCompounds.push(compoundPath);
                compoundCopies.push(compoundCopy);
            }
            var memberIndex = getMemberIndex(compoundPath, targetPath);
            if (memberIndex >= 0) {
                applyCornerRadius(compoundCopy.pathItems[memberIndex], measurements[i], cornerRadius, cornerOptions);
            }
        }
        app.redraw();
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 入力欄の値を pt で読む（数値でない・負の値は 0）
     * @param {EditText} inputField - 半径の入力欄
     * @param {number} pointsPerUnit - 1 単位あたりの pt
     * @returns {number} 半径（pt）
     */
    function readRadiusField(inputField, pointsPerUnit) {
        var fieldValue = parseFloat(inputField.text);
        if (isNaN(fieldValue) || fieldValue < 0) return 0;
        return fieldValue * pointsPerUnit;
    }

    /**
     * pt の値を定規単位の表示用文字列にする（小数第 2 位まで）
     * @param {number} points - 値（pt）
     * @param {number} pointsPerUnit - 1 単位あたりの pt
     * @returns {string} 表示用の文字列
     */
    function formatRadius(points, pointsPerUnit) {
        return String(Math.round(points / pointsPerUnit * 100) / 100);
    }

    /**
     * 「半径：［入力欄］単位」の行を左右中央に追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} initialText - 入力欄の初期値
     * @param {string} unitLabel - 単位の表示
     * @param {function} onValueStepped - ∧∨・↑↓キーで値を変えたあとに呼ぶコールバック
     * @returns {EditText} 半径の入力欄
     */
    function addRadiusRow(parentWindow, initialText, unitLabel, onValueStepped) {
        var radiusRow = parentWindow.add("group");
        setupRow(radiusRow, "center");
        radiusRow.add("statictext", undefined, labelText("fieldLabel.radius"));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = radiusRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var radiusField;
        var radiusStepper = addStepper(stepperInputGroup, function () { return radiusField; }, {
            min: 0,
            onStep: function () { onValueStepped(); }
        });
        radiusField = stepperInputGroup.add("edittext", undefined, initialText);
        radiusField.characters = FIELD_CHARACTERS;
        bindSteppedArrowKeys(radiusField, radiusStepper);
        radiusField.helpTip = getLabel("tooltip.radiusField");
        radiusRow.add("statictext", undefined, unitLabel);
        return radiusField;
    }

    /**
     * ［対象］パネルとラジオボタンを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} initialScope - 最初に選ぶ対象のキー
     * @param {boolean} hasSelection - 選択に対象の長方形があるか（無ければ［選択したオブジェクトのみ］をディム）
     * @returns {Object} { selection, artboard, document } のラジオボタン
     */
    function addScopePanel(parentWindow, initialScope, hasSelection) {
        var scopePanel = parentWindow.add("panel", undefined, getLabel("panel.targetScope"));
        setupPanel(scopePanel);
        var scopeRadios = {
            selection: scopePanel.add("radiobutton", undefined, getLabel("radio.selection")),
            artboard: scopePanel.add("radiobutton", undefined, getLabel("radio.artboard")),
            document: scopePanel.add("radiobutton", undefined, getLabel("radio.document"))
        };
        scopeRadios.artboard.helpTip = getLabel("tooltip.artboard");
        scopeRadios.document.helpTip = getLabel("tooltip.document");
        scopeRadios.selection.enabled = hasSelection;
        scopeRadios[initialScope].value = true;
        return scopeRadios;
    }

    /**
     * ラベルと tooltip の付いたチェックボックスを追加する
     * @param {Object} parentContainer - 追加先（パネル・ダイアログ）
     * @param {string} labelKey - LABELS.checkbox のキー
     * @param {boolean} initialValue - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentContainer, labelKey, initialValue) {
        var optionCheckbox = parentContainer.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        var tooltipText = getLabel("tooltip." + labelKey);
        if (tooltipText !== "tooltip." + labelKey) optionCheckbox.helpTip = tooltipText;
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 角丸の半径を入力するダイアログを表示する
     * @param {Object} scopeTargets - 対象ごとのパス { selection, artboard, document }（artboard・document は未収集なら null）
     * @param {function} collectScopeTargets - 対象のキーを受け取り、パスを集めて返す関数
     * @param {number} skippedCount - 選択のうち対象外のオブジェクト数
     * @returns {Object|null} OK なら { targetPaths, measurements, radius, cornerOptions }、キャンセルなら null
     */
    function showRadiusDialog(scopeTargets, collectScopeTargets, skippedCount) {
        var unitInfo = getUnitInfo();
        var hasSelection = scopeTargets.selection.length > 0;
        var currentScope = hasSelection ? "selection" : "artboard";
        var targetPaths = collectScopeTargets(currentScope);

        /* チェックボックスの状態は変数で持つ。表示前に代入した value は読み戻すと false になることがある
           Track checkbox states in variables: a value assigned before show() can read back as false */
        var cornerOptions = {
            keepZeroRadii: KEEP_ZERO_RADII_DEFAULT,
            includeEffect: INCLUDE_EFFECT_DEFAULT,
            convertToEffect: CONVERT_TO_EFFECT_DEFAULT
        };
        var isPreviewOn = PREVIEW_DEFAULT;

        /* 効果を含めた計測は重いので対象ごとに控える / Cache the effect-inclusive measurement per scope */
        var effectMeasurementsByScope = {};

        /**
         * 対象パスごとの計測結果を返す
         * プレビューの複製を消し、元のパスが表示に戻ってから呼ぶ
         * @returns {Object[]} measureCorners() の戻り値の配列
         */
        function getMeasurements() {
            var includeEffect = cornerOptions.includeEffect;
            if (includeEffect && effectMeasurementsByScope[currentScope]) return effectMeasurementsByScope[currentScope];
            var measurements = [];
            for (var i = 0; i < targetPaths.length; i++) measurements.push(measureCorners(targetPaths[i], includeEffect));
            if (includeEffect) effectMeasurementsByScope[currentScope] = measurements;
            return measurements;
        }

        var radiusDialog = new Window("dialog", getLabel("dialog.title"));
        radiusDialog.orientation = "column";
        radiusDialog.alignChildren = ["fill", "top"];
        radiusDialog.margins = WINDOW_MARGINS;
        radiusDialog.spacing = WINDOW_SPACING;

        /* 初期値は対象の角丸の平均 / Start from the average radius of the targets */
        var radiusField = addRadiusRow(radiusDialog,
            formatRadius(getAverageRadius(getMeasurements()), unitInfo.pointsPerUnit), unitInfo.label,
            function () { refreshPreview(); });
        var scopeRadios = addScopePanel(radiusDialog, currentScope, hasSelection);

        var optionsPanel = radiusDialog.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel);
        var keepZeroCheckbox = addOptionCheckbox(optionsPanel, "keepZeroRadii", KEEP_ZERO_RADII_DEFAULT);
        var includeEffectCheckbox = addOptionCheckbox(optionsPanel, "includeEffect", INCLUDE_EFFECT_DEFAULT);
        var convertToEffectCheckbox = addOptionCheckbox(optionsPanel, "convertToEffect", CONVERT_TO_EFFECT_DEFAULT);
        /* ［0の半径は0のままに］がオンのあいだは変換できない / Conversion is unavailable while keeping zero radii */
        convertToEffectCheckbox.enabled = !KEEP_ZERO_RADII_DEFAULT;

        /* 対象外の数は「選択したオブジェクトのみ」のときだけ表示。選べないときは行ごと作らない（空欄でも幅と高さが残る）
           Skipped count shows only for the selection scope; omit the row when it is unavailable (an empty text keeps its size) */
        var skippedText = null;
        var skippedLabel = getLabel("status.skippedCount").replace("{count}", skippedCount);
        if (hasSelection && skippedCount > 0) {
            skippedText = radiusDialog.add("statictext", undefined, skippedLabel);
            skippedText.helpTip = getLabel("tooltip.skippedCount");
        } else if (!hasSelection && skippedCount > 0) {
            /* 選択はあるが対象が無い（複合シェイプ全体を選んだときなど）。対象を切り替えても出したままにする
               Something is selected but nothing qualifies (e.g. a whole compound shape); keep it across scopes */
            var noTargetText = radiusDialog.add("statictext", undefined, getLabel("status.noTargetInSelection"));
            noTargetText.helpTip = getLabel("tooltip.noTargetInSelection");
            noTargetText.alignment = ["center", "top"];
        }

        var previewCheckbox = addOptionCheckbox(radiusDialog, "preview", PREVIEW_DEFAULT);
        previewCheckbox.alignment = ["center", "top"];

        var buttonRow = addButtonRow(radiusDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * プレビューの表示状態を入力に合わせる
         * @returns {void}
         */
        function refreshPreview() {
            /* 計測は元のパスが表示に戻ってから / Measure after the originals are visible again */
            clearPreview();
            if (isPreviewOn) {
                showPreview(targetPaths, getMeasurements(), readRadiusField(radiusField, unitInfo.pointsPerUnit),
                    cornerOptions);
            } else {
                app.redraw();
            }
        }

        /**
         * 対象を切り替えて、対象外の表示・OK の可否・プレビューを更新する
         * @param {string} scopeKey - "selection" / "artboard" / "document"
         * @returns {void}
         */
        function changeScope(scopeKey) {
            /* プレビューで隠した元のパスを取りこぼさないよう、集める前に消す
               Clear the preview first so hidden originals are not skipped while collecting */
            clearPreview();
            currentScope = scopeKey;
            targetPaths = collectScopeTargets(scopeKey);
            if (skippedText) skippedText.text = (scopeKey === "selection") ? skippedLabel : "";
            btnOK.enabled = targetPaths.length > 0;
            refreshPreview();
        }

        for (var scopeKey in scopeRadios) {
            scopeRadios[scopeKey].onClick = (function (clickedScope) {
                return function () { changeScope(clickedScope); };
            })(scopeKey);
        }
        radiusField.onChanging = refreshPreview;
        keepZeroCheckbox.onClick = function () {
            cornerOptions.keepZeroRadii = keepZeroCheckbox.value;
            convertToEffectCheckbox.enabled = !cornerOptions.keepZeroRadii;
            refreshPreview();
        };
        includeEffectCheckbox.onClick = function () {
            cornerOptions.includeEffect = includeEffectCheckbox.value;
            refreshPreview();
        };
        convertToEffectCheckbox.onClick = function () {
            cornerOptions.convertToEffect = convertToEffectCheckbox.value;
            refreshPreview();
        };
        previewCheckbox.onClick = function () {
            isPreviewOn = previewCheckbox.value;
            refreshPreview();
        };

        radiusDialog.onShow = function () {
            radiusField.active = true;
            changeScope(currentScope);
        };

        prepareDialogWindow(radiusDialog, SCRIPT_NAME);
        var dialogResult = radiusDialog.show();
        clearPreview();
        if (dialogResult !== 1) return null;
        return {
            targetPaths: targetPaths,
            measurements: getMeasurements(),
            radius: readRadiusField(radiusField, unitInfo.pointsPerUnit),
            cornerOptions: cornerOptions
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログの入力値を対象のパスに適用する
     * @param {Object} dialogValues - showRadiusDialog() の戻り値
     * @param {PageItem[]} initialSelection - 実行前の選択
     * @returns {PageItem[]} 戻す選択（アピアランスの消去で作り直されたパスは新しい参照に置き換える）
     */
    function applyDialogValues(dialogValues, initialSelection) {
        var restoredSelection = initialSelection.slice(0);
        for (var i = 0; i < dialogValues.targetPaths.length; i++) {
            var targetPath = dialogValues.targetPaths[i];
            var resultPath = applyCornerRadius(targetPath, dialogValues.measurements[i],
                dialogValues.radius, dialogValues.cornerOptions);
            for (var j = 0; j < restoredSelection.length; j++) {
                if (restoredSelection[j] === targetPath) restoredSelection[j] = resultPath;
            }
        }
        return restoredSelection;
    }

    /**
     * 選択・アートボード・ドキュメントのいずれかにある長方形の角丸半径をダイアログで変更する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var initialSelection = doc.selection || [];

        var scopeTargets = { selection: [], artboard: null, document: null };

        /**
         * 対象のパスを返す（アートボード・ドキュメントは初回だけ集める）
         * @param {string} scopeKey - "selection" / "artboard" / "document"
         * @returns {PathItem[]} 対象パス
         */
        function collectScopeTargets(scopeKey) {
            if (scopeTargets[scopeKey] === null) {
                var artboardRect = null;
                if (scopeKey === "artboard") {
                    artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
                }
                scopeTargets[scopeKey] = collectDocumentRectangles(doc, artboardRect);
            }
            return scopeTargets[scopeKey];
        }

        var restoredSelection = initialSelection;
        /* 途中で止まっても、隠した変換結果をドキュメントに残さない / Never leave hidden conversions behind */
        try {
            /* 選択の複合シェイプはここで変換する / Compound shapes in the selection are converted here */
            var selectedRectangles = collectSelectedRectangles(initialSelection);
            scopeTargets.selection = selectedRectangles.targetPaths;
            var dialogValues = showRadiusDialog(scopeTargets, collectScopeTargets, selectedRectangles.skippedCount);
            if (dialogValues) {
                restoredSelection = applyDialogValues(dialogValues, initialSelection);
                finishShapeConversions(dialogValues.targetPaths, restoredSelection);
            }
        } finally {
            discardShapeConversions();
        }
        /* 計測・変換・効果の付け直しで変わる選択を元に戻す / Restore the selection changed by measuring, converting and reapplying */
        doc.selection = (restoredSelection.length > 0) ? restoredSelection : null;
        app.redraw();
    }

    main();

})();
