#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### 概要

作業アートボード、すべてのアートボード、または番号で指定したアートボードを、プレビューしながら指定の幅・高さに変更します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeArtboardsAll.md

### Overview

Resizes the active artboard, every artboard, or the artboards you list by number to a given width and height, with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResizeArtboardsAll.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResizeArtboardsAll";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResizeArtboardsAll.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResizeArtboardsAll.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 15;                 /* ウィンドウ外周の余白 / window margin */
    var PANEL_MARGINS  = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 15;                 /* 2カラムの間隔 / gap between columns */
    var COLUMN_PANEL_SPACING = 10;           /* カラム内のパネル間隔 / gap between panels in a column */
    var SIZE_ROW_SPACING = 6;                /* 幅・高さの行間 / gap between the width and height rows */
    var ANCHOR_RADIO_SPACING = 12;           /* 基準点のラジオの間隔 / gap between the reference point radios */
    var BUTTON_ROW_TOP_MARGIN = 5;           /* ボタンエリアの上余白 / top margin of the button row */
    var SIZE_FIELD_CHARACTERS = 5;           /* 幅・高さ欄の桁数 / width & height field characters */
    var SPECIFY_FIELD_CHARACTERS = 12;       /* 番号指定欄の桁数 / artboard number field characters */
    var LABEL_WIDTH_PADDING = 6;             /* 項目名の幅に足す余白 / padding added to the measured label width */

    /* ダイアログの不透明度と、初回表示時の画面中央からの横オフセット / Dialog opacity and first-run offset from screen center */
    var DIALOG_OPACITY = 0.95;
    var DIALOG_FIRST_RUN_OFFSET_X = 300;

    /* プレビューの再描画の最短間隔（ms） / Minimum interval between preview redraws (ms) */
    var REDRAW_INTERVAL_MS = 40;

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
     * pt・px のように整数で表示する単位か判定する
     * @param {{label: string}} unitInfo - 単位の情報
     * @returns {boolean} 整数で表示するなら true
     */
    function isIntegerUnit(unitInfo) {
        return unitInfo.label === "pt" || unitInfo.label === "px";
    }

    /**
     * pt の値を、単位に合わせた表示用の文字列にする（pt・px は整数、それ以外は小数2桁まで）
     * @param {number} valuePt - pt の値
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 単位の情報
     * @returns {string} 表示用の文字列
     */
    function formatSizeValue(valuePt, unitInfo) {
        var unitValue = valuePt / unitInfo.pointsPerUnit;
        if (isIntegerUnit(unitInfo)) return String(Math.round(unitValue));
        return String(Math.round(unitValue * 100) / 100);
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
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
    // 6. この欄には別の↑↓キー増減処理を付けない（↑↓キーが二重に効く）
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
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkStepperUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var STEPPER_UI_DARK           = isDarkStepperUI();
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

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "アートボードのサイズを変更", en: "Resize Artboards" }
        },
        panel: {
            size: { ja: "サイズ", en: "Size" },
            anchor: { ja: "基準点", en: "Reference Point" },
            target: { ja: "対象のアートボード", en: "Target Artboards" }
        },
        fieldLabel: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        radio: {
            anchorTopLeft: { ja: "左上", en: "Top-Left" },
            anchorCenter: { ja: "中央", en: "Center" },
            activeArtboard: { ja: "作業アートボードのみ", en: "Active artboard only" },
            allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
            specify: { ja: "指定", en: "Specify" }
        },
        tooltip: {
            sizeField: {
                ja: "↑↓で次の整数へ、Shift+↑↓で次の10の倍数へ増減します。",
                en: "Up/Down: step to the next whole number. Shift+Up/Down: snap to the next multiple of 10."
            },
            anchorTopLeft: { ja: "左上の位置を保ったまま、右と下に伸び縮みさせます。", en: "Keep the top-left corner and resize to the right and down." },
            anchorCenter: { ja: "中心の位置を保ったまま伸び縮みさせます。", en: "Keep the center and resize around it." },
            specifyField: {
                ja: "アートボードの番号（1始まり）を範囲やカンマ区切りで指定します。\n例：1-3 / 1,3 / 2-4,7",
                en: "Artboard numbers (starting at 1), as ranges or comma-separated.\ne.g. 1-3 / 1,3 / 2-4,7"
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            errorOccurred: { ja: "エラーが発生しました：", en: "An error occurred: " }
        }
    };

    /**
     * ローカライズ文字列を取得する（キー漏れ時は英語へフォールバック）
     * @param {Object} labelSet - { ja, en } のラベル
     * @returns {string} 表示言語の文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        if (labelSet[uiLang] != null) return labelSet[uiLang];
        return (labelSet.en != null) ? labelSet.en : "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - ラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 単位を括弧で添えたパネル名を返す（日本語は全角括弧、英語は半角）
     * @param {Object} labelSet - パネル名のラベル
     * @param {string} unitLabel - 単位の表示ラベル
     * @returns {string} 単位付きのパネル名
     */
    function panelTitleWithUnit(labelSet, unitLabel) {
        return (uiLang === "ja")
            ? getLabel(labelSet) + "（" + unitLabel + "）"
            : getLabel(labelSet) + " (" + unitLabel + ")";
    }

    // =========================================
    // エラー処理 / Error handling
    // =========================================

    /**
     * Error を行番号・ファイル名付きで読みやすく整形する
     * @param {Error} error - 例外
     * @returns {string} 整形した文字列
     */
    function formatError(error) {
        var messageText = (error && error.message) ? String(error.message) : String(error);
        var lineText = (error && error.line) ? (" line " + error.line) : "";
        var fileText = (error && error.fileName) ? (" (" + error.fileName + ")") : "";
        return messageText + lineText + fileText;
    }

    // =========================================
    // ダイアログ位置の記憶 / Dialog position persistence
    // =========================================
    // #targetengine の $.global に置くので、Illustrator を終了するまで位置が残る。
    // Kept in $.global of the named engine, so it lasts until Illustrator quits.

    var DIALOG_POSITION_KEY = "__ResizeArtboardsAll_Dialog";

    /**
     * 保存済みのダイアログ位置を取得する
     * @param {string} storageKey - $.global のキー
     * @returns {number[]|null} [x, y]。無ければ null
     */
    function getStoredLocation(storageKey) {
        return $.global[storageKey] && $.global[storageKey].length === 2 ? $.global[storageKey] : null;
    }

    /**
     * ダイアログ位置をセッションに保存する
     * @param {string} storageKey - $.global のキー
     * @param {number[]} location - [x, y]
     * @returns {void}
     */
    function storeLocation(storageKey, location) {
        $.global[storageKey] = [location[0], location[1]];
    }

    /**
     * 位置を画面内に収める
     * @param {number[]} location - [x, y]
     * @returns {number[]} 画面内に収めた [x, y]
     */
    function clampLocationToScreen(location) {
        /* 画面情報が取れない環境では元の位置のまま / keep the location when screen info is unavailable */
        try {
            var visibleBounds = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
            var clampedX = Math.max(visibleBounds[0] + 10, Math.min(location[0], visibleBounds[2] - 10));
            var clampedY = Math.max(visibleBounds[1] + 10, Math.min(location[1], visibleBounds[3] - 10));
            return [clampedX, clampedY];
        } catch (e) {
            return location;
        }
    }

    /**
     * ダイアログ位置の記憶を設定し、保存関数を返す
     * 保存位置があれば表示時に復元し、無ければ初回は画面中央から横にずらして表示する
     * @param {Window} dialogWindow - 対象のダイアログ
     * @param {string} positionKey - $.global のキー
     * @param {number} firstRunOffsetX - 初回表示時の中央からの横オフセット
     * @returns {function} 現在位置を保存する関数
     */
    function attachPositionPersistence(dialogWindow, positionKey, firstRunOffsetX) {
        var savedLocation = getStoredLocation(positionKey);

        var persist = function () {
            storeLocation(positionKey, [dialogWindow.location[0], dialogWindow.location[1]]);
        };

        if (savedLocation) {
            dialogWindow.onShow = function () {
                dialogWindow.location = clampLocationToScreen(savedLocation);
            };
        } else {
            dialogWindow.onShow = function () {
                dialogWindow.layout.layout(true);
                var screenWidth = $.screens[0].right - $.screens[0].left;
                var screenHeight = $.screens[0].bottom - $.screens[0].top;
                var centerX = screenWidth / 2 - dialogWindow.bounds.width / 2;
                var centerY = screenHeight / 2 - dialogWindow.bounds.height / 2;
                dialogWindow.location = [centerX + firstRunOffsetX, centerY];
            };
        }

        dialogWindow.onMove = persist;
        return persist;
    }

    // =========================================
    // 入力の解釈 / Input parsing
    // =========================================

    /**
     * 入力欄の文字列から数値を取り出す（数字・小数点・マイナス以外は無視）
     * @param {string} inputText - 入力欄の文字列
     * @returns {number} 数値（読み取れなければ NaN）
     */
    function parseNumberInput(inputText) {
        return parseFloat(String(inputText).replace(/[^0-9.\-]/g, ""));
    }

    /**
     * 「1-3,5,7-9」形式の番号指定を、0始まりのインデックス配列にする
     * 範囲外の番号は無視し、重複を除いて昇順に並べる
     * @param {string} specifyText - 入力された番号指定（1始まり）
     * @param {number} artboardCount - アートボードの数
     * @returns {number[]} 0始まりのインデックス
     */
    function parseArtboardNumbers(specifyText, artboardCount) {
        var compactText = String(specifyText || "").replace(/\s+/g, "");
        if (!compactText) return [];

        var isIndexListed = {};

        /**
         * 1始まりの番号を範囲内なら登録する
         * @param {number} artboardNumber - 1始まりの番号
         * @returns {void}
         */
        function addArtboardNumber(artboardNumber) {
            var artboardIndex = artboardNumber - 1;
            if (artboardIndex >= 0 && artboardIndex < artboardCount) isIndexListed[artboardIndex] = true;
        }

        var specifyParts = compactText.split(",");
        for (var i = 0; i < specifyParts.length; i++) {
            if (!specifyParts[i]) continue;
            var rangeMatch = specifyParts[i].match(/^(\d+)-(\d+)$/);
            if (rangeMatch) {
                var rangeStart = parseInt(rangeMatch[1], 10);
                var rangeEnd = parseInt(rangeMatch[2], 10);
                if (rangeStart > rangeEnd) {
                    var swapValue = rangeStart;
                    rangeStart = rangeEnd;
                    rangeEnd = swapValue;
                }
                for (var artboardNumber = rangeStart; artboardNumber <= rangeEnd; artboardNumber++) {
                    addArtboardNumber(artboardNumber);
                }
            } else {
                var singleNumber = parseInt(specifyParts[i], 10);
                if (!isNaN(singleNumber)) addArtboardNumber(singleNumber);
            }
        }

        var artboardIndexes = [];
        for (var indexKey in isIndexListed) {
            if (isIndexListed.hasOwnProperty(indexKey)) artboardIndexes.push(parseInt(indexKey, 10));
        }
        artboardIndexes.sort(function (a, b) { return a - b; });
        return artboardIndexes;
    }

    // =========================================
    // アートボードの変更 / Artboard resizing
    // =========================================

    /**
     * すべてのアートボードの矩形を控える
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {Array<number[]>} [左, 上, 右, 下] の配列
     */
    function captureArtboardRects(targetDocument) {
        var artboardRects = [];
        for (var i = 0; i < targetDocument.artboards.length; i++) {
            artboardRects.push(targetDocument.artboards[i].artboardRect.slice());
        }
        return artboardRects;
    }

    /**
     * 控えた矩形に、すべてのアートボードを戻す
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {Array<number[]>} artboardRects - 控えた [左, 上, 右, 下] の配列
     * @returns {void}
     */
    function restoreArtboardRects(targetDocument, artboardRects) {
        for (var i = 0; i < artboardRects.length; i++) {
            targetDocument.artboards[i].artboardRect = artboardRects[i].slice();
        }
    }

    /**
     * 元の矩形と基準点から、新しいサイズの矩形を求める
     * px 単位のときは、左上を整数座標にそろえる
     * @param {number[]} sourceRect - 元の [左, 上, 右, 下]
     * @param {number} widthPt - 新しい幅（pt）
     * @param {number} heightPt - 新しい高さ（pt）
     * @param {boolean} isTopLeftAnchor - 左上基準なら true、中央基準なら false
     * @param {boolean} snapToPixel - 左上を整数座標にそろえるなら true
     * @returns {number[]} 新しい [左, 上, 右, 下]
     */
    function computeResizedRect(sourceRect, widthPt, heightPt, isTopLeftAnchor, snapToPixel) {
        var leftEdge, topEdge;
        if (isTopLeftAnchor) {
            leftEdge = sourceRect[0];
            topEdge = sourceRect[1];
        } else {
            leftEdge = (sourceRect[0] + sourceRect[2]) / 2 - widthPt / 2;
            topEdge = (sourceRect[1] + sourceRect[3]) / 2 + heightPt / 2;
        }
        if (snapToPixel) {
            leftEdge = Math.round(leftEdge);
            topEdge = Math.round(topEdge);
        }
        return [leftEdge, topEdge, leftEdge + widthPt, topEdge - heightPt];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名と数値入力欄の行を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} initialText - 入力欄の初期値
     * @param {Function} onStep - ∧∨・↑↓キーで増減したあとに呼ぶ関数
     * @returns {{label: StaticText, input: EditText}} 追加した項目名と入力欄
     */
    function addSizeRow(parentGroup, labelSet, initialText, onStep) {
        var sizeRowGroup = parentGroup.add("group");
        sizeRowGroup.orientation = "row";
        var sizeLabel = sizeRowGroup.add("statictext", undefined, labelText(labelSet));
        sizeLabel.justify = "right";

        /* ∧∨と入力欄は隙間0で突き合わせる。0未満にはしない / butt the stepper against the field; never below 0 */
        var stepperInputGroup = sizeRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var sizeInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return sizeInput; }, {
            step: 1, min: 0,
            onStep: function () { onStep(); }
        });
        sizeInput = stepperInputGroup.add("edittext", undefined, initialText);
        sizeInput.characters = SIZE_FIELD_CHARACTERS;
        sizeInput.helpTip = getLabel(LABELS.tooltip.sizeField);
        bindSteppedArrowKeys(sizeInput, stepperGroup);
        return { label: sizeLabel, input: sizeInput };
    }

    /**
     * 項目名の幅を、長いほうに合わせてそろえる（右揃えで入力欄の左端が縦にそろう）
     * @param {Window} dialogWindow - 文字幅を測るウィンドウ
     * @param {StaticText[]} labelControls - そろえる項目名
     * @returns {void}
     */
    function equalizeLabelWidths(dialogWindow, labelControls) {
        var maxWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            var labelWidth;
            /* 表示前は measureString が使えない環境があるため、文字数から見積もる / estimate when measuring fails */
            try {
                labelWidth = Math.ceil(dialogWindow.graphics.measureString(labelControls[i].text)[0]);
            } catch (e) {
                labelWidth = labelControls[i].text.length * 7;
            }
            if (labelWidth > maxWidth) maxWidth = labelWidth;
        }
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxWidth + LABEL_WIDTH_PADDING;
        }
    }

    /**
     * タイトル付きのパネルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} panelTitle - パネル名
     * @param {string} childOrientation - 中のグループの並び（"row" / "column"）
     * @returns {Group} コントロールを入れるグループ
     */
    function addTitledPanel(parentGroup, panelTitle, childOrientation) {
        var titledPanel = parentGroup.add("panel", undefined, panelTitle);
        titledPanel.orientation = "row";
        titledPanel.alignChildren = ["left", "top"];
        titledPanel.margins = PANEL_MARGINS;
        var contentGroup = titledPanel.add("group");
        contentGroup.orientation = childOrientation;
        contentGroup.alignChildren = ["left", "center"];
        return contentGroup;
    }

    /**
     * 列のグループを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Group} 追加した列
     */
    function addColumn(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "fill";
        columnGroup.spacing = COLUMN_PANEL_SPACING;
        return columnGroup;
    }

    /**
     * サイズ変更のダイアログを表示し、プレビューしながらアートボードを変更する
     * OK で変更を確定し、キャンセルで開いた時点のサイズに戻す
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 定規の単位
     * @returns {boolean} OK で閉じたら true
     */
    function showResizeDialog(targetDocument, unitInfo) {
        var artboardCount = targetDocument.artboards.length;
        /* プレビューのたびに開いた時点の矩形へ戻してから適用する / Every preview starts from the original rects */
        var originalRects = captureArtboardRects(targetDocument);
        var activeRect = originalRects[targetDocument.artboards.getActiveArtboardIndex()];

        var resizeDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        resizeDialog.orientation = "column";
        resizeDialog.alignChildren = "fill";
        resizeDialog.margins = WINDOW_MARGINS;
        resizeDialog.opacity = DIALOG_OPACITY;
        var persistLocation = attachPositionPersistence(resizeDialog, DIALOG_POSITION_KEY, DIALOG_FIRST_RUN_OFFSET_X);

        var columnsGroup = resizeDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = addColumn(columnsGroup);
        var rightColumn = addColumn(columnsGroup);

        /* サイズパネル：作業アートボードの今のサイズを初期値にする / Size panel, seeded with the active artboard */
        var sizeGroup = addTitledPanel(leftColumn, panelTitleWithUnit(LABELS.panel.size, unitInfo.label), "column");
        sizeGroup.spacing = SIZE_ROW_SPACING;
        var widthRow = addSizeRow(sizeGroup, LABELS.fieldLabel.width, formatSizeValue(Math.abs(activeRect[2] - activeRect[0]), unitInfo), applyResizePreview);
        var heightRow = addSizeRow(sizeGroup, LABELS.fieldLabel.height, formatSizeValue(Math.abs(activeRect[1] - activeRect[3]), unitInfo), applyResizePreview);
        var widthInput = widthRow.input;
        var heightInput = heightRow.input;
        equalizeLabelWidths(resizeDialog, [widthRow.label, heightRow.label]);

        /* 基準点パネル / Reference point panel */
        var anchorGroup = addTitledPanel(leftColumn, getLabel(LABELS.panel.anchor), "row");
        anchorGroup.spacing = ANCHOR_RADIO_SPACING;
        var anchorTopLeftRadio = anchorGroup.add("radiobutton", undefined, getLabel(LABELS.radio.anchorTopLeft));
        var anchorCenterRadio = anchorGroup.add("radiobutton", undefined, getLabel(LABELS.radio.anchorCenter));
        anchorTopLeftRadio.helpTip = getLabel(LABELS.tooltip.anchorTopLeft);
        anchorCenterRadio.helpTip = getLabel(LABELS.tooltip.anchorCenter);
        anchorTopLeftRadio.value = true;

        /* 対象パネル / Target panel */
        var targetGroup = addTitledPanel(rightColumn, getLabel(LABELS.panel.target), "column");
        targetGroup.alignChildren = ["left", "top"];
        var activeArtboardRadio = targetGroup.add("radiobutton", undefined, getLabel(LABELS.radio.activeArtboard));
        var allArtboardsRadio = targetGroup.add("radiobutton", undefined, getLabel(LABELS.radio.allArtboards));
        var specifyRadio = targetGroup.add("radiobutton", undefined, getLabel(LABELS.radio.specify));
        var specifyInput = targetGroup.add("edittext", undefined, "");
        specifyInput.characters = SPECIFY_FIELD_CHARACTERS;
        specifyInput.helpTip = getLabel(LABELS.tooltip.specifyField);
        activeArtboardRadio.value = true;

        /* アートボードが1つなら「すべて」「指定」は選べない / Only one artboard: nothing else to pick */
        if (artboardCount <= 1) {
            allArtboardsRadio.enabled = false;
            specifyRadio.enabled = false;
        }

        /* ボタンエリア / Button row */
        var btnRowGroup = resizeDialog.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        var lastRedrawTime = 0;

        /**
         * 再描画を間引く（連続入力で画面がもたつかないように）
         * @returns {void}
         */
        function throttledRedraw() {
            var currentTime = (new Date()).getTime();
            if (currentTime - lastRedrawTime >= REDRAW_INTERVAL_MS) {
                app.redraw();
                lastRedrawTime = currentTime;
            }
        }

        /**
         * 番号指定の入力欄を、［指定］を選んでいるときだけ使えるようにする
         * @returns {void}
         */
        function updateSpecifyEnabled() {
            specifyInput.enabled = (specifyRadio.value === true);
        }

        /**
         * 選択中の対象から、変更するアートボードのインデックスを決める
         * ［指定］で番号が1つも拾えないときは作業アートボードにする
         * @returns {number[]} 0始まりのインデックス
         */
        function resolveTargetIndexes() {
            if (specifyRadio.value) {
                var specifiedIndexes = parseArtboardNumbers(specifyInput.text, artboardCount);
                if (specifiedIndexes.length) return specifiedIndexes;
            } else if (allArtboardsRadio.value) {
                var allIndexes = [];
                for (var i = 0; i < artboardCount; i++) allIndexes.push(i);
                return allIndexes;
            }
            return [targetDocument.artboards.getActiveArtboardIndex()];
        }

        /**
         * 入力中の幅・高さで、対象のアートボードを変更する（プレビュー）
         * 幅・高さが数値として読めないあいだは何もしない
         * @returns {void}
         */
        function applyResizePreview() {
            var widthValue = parseNumberInput(widthInput.text);
            var heightValue = parseNumberInput(heightInput.text);
            if (isNaN(widthValue) || isNaN(heightValue) || widthValue <= 0 || heightValue <= 0) return;

            var widthPt = widthValue * unitInfo.pointsPerUnit;
            var heightPt = heightValue * unitInfo.pointsPerUnit;
            var isTopLeftAnchor = (anchorTopLeftRadio.value === true);
            var snapToPixel = (unitInfo.label === "px");

            restoreArtboardRects(targetDocument, originalRects);
            var targetIndexes = resolveTargetIndexes();
            for (var i = 0; i < targetIndexes.length; i++) {
                var artboardIndex = targetIndexes[i];
                targetDocument.artboards[artboardIndex].artboardRect =
                    computeResizedRect(originalRects[artboardIndex], widthPt, heightPt, isTopLeftAnchor, snapToPixel);
            }
            throttledRedraw();
        }

        /**
         * 対象の選び直しを反映する
         * @returns {void}
         */
        function onTargetChanged() {
            updateSpecifyEnabled();
            applyResizePreview();
        }

        widthInput.onChanging = applyResizePreview;
        heightInput.onChanging = applyResizePreview;
        anchorTopLeftRadio.onClick = applyResizePreview;
        anchorCenterRadio.onClick = applyResizePreview;
        activeArtboardRadio.onClick = onTargetChanged;
        allArtboardsRadio.onClick = onTargetChanged;
        specifyRadio.onClick = onTargetChanged;
        /* 番号は［指定］を選んでいるときだけ効く / The numbers only matter in Specify mode */
        specifyInput.onChanging = function () { if (specifyRadio.value) applyResizePreview(); };
        specifyInput.onChange = specifyInput.onChanging;

        var isConfirmed = false;
        btnOK.onClick = function () {
            persistLocation();
            applyResizePreview();
            isConfirmed = true;
            resizeDialog.close(1);
        };
        btnCancel.onClick = function () {
            persistLocation();
            restoreArtboardRects(targetDocument, originalRects);
            app.redraw();
            resizeDialog.close(0);
        };

        /* 表示時の位置合わせに続けて、幅の欄にフォーカスを置く / Focus the width field after positioning */
        var positionOnShow = resizeDialog.onShow;
        resizeDialog.onShow = function () {
            positionOnShow();
            widthInput.active = true;
        };

        updateSpecifyEnabled();
        applyResizePreview();
        resizeDialog.show();
        return isConfirmed;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを確かめてダイアログを開く（変更はダイアログ内のプレビューで確定する）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        showResizeDialog(app.activeDocument, getUnitInfo());
    }

    try {
        main();
    } catch (e) {
        $.writeln("[" + SCRIPT_NAME + "] ERROR: " + formatError(e));
        alert(getLabel(LABELS.alert.errorOccurred) + formatError(e));
    }

    app.selectTool("Adobe Select Tool");

})();
