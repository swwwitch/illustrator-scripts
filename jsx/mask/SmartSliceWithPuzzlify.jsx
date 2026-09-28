#target illustrator
#targetengine "SmartSliceWithPuzzlifyEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した画像やオブジェクトを、グリッド（格子）またはジグソーパズル形状のピースに分割し、各ピースでマスクします。
行数・列数・ピース数のほか、オフセット、オーバーラップ、バラけ、ケイ線、角丸を指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSliceWithPuzzlify.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n89f63325c0bc

### Overview

Slices the selected image or artwork into grid cells or jigsaw pieces and masks each piece.
Besides rows, columns and piece count, you can set offset, overlap, scatter, stroke and rounded corners.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSliceWithPuzzlify.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartSliceWithPuzzlify";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSliceWithPuzzlify.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSliceWithPuzzlify.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n89f63325c0bc"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値 / Initial dialog values */
    var DEFAULT_PIECES_PUZZLE = "25";  /* パズル時のピース数 / piece count in puzzle mode */
    var DEFAULT_PIECES_GRID   = "2";   /* グリッド時のピース数 / piece count in grid mode */
    var DEFAULT_COLUMNS       = "6";   /* 選択の寸法が取れないときの列数 / columns when the selection size is unknown */
    var DEFAULT_ROWS          = "4";   /* 選択の寸法が取れないときの行数 / rows when the selection size is unknown */
    var DEFAULT_OFFSET        = "-2";  /* オフセット / offset */
    var DEFAULT_OVERLAP       = "10";  /* オーバーラップ / overlap */
    var DEFAULT_SCATTER       = "30";  /* バラけの最大移動量 / maximum scatter distance */
    var DEFAULT_ROUND_RADIUS  = "3";   /* 角丸の半径 / round corner radius */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var MODE_ROW_MARGINS     = [10, 5, 10, 5];   /* 分割方法の行の余白 / margins of the mode row */
    var PANEL_MARGINS        = [15, 20, 15, 10]; /* パネル余白 [左,上,右,下] / panel margins */
    var SHAPE_ROW_MARGINS    = [0, 10, 0, 10];   /* 形状の行の余白 / margins of the shape row */
    var PROGRESS_ROW_MARGINS = [10, 0, 10, 0];   /* プログレスバーの行の余白 / margins of the progress row */
    var PROGRESS_BAR_SIZE    = [200, 7];         /* プログレスバーの寸法 / progress bar size */

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グリッド／パズルに分割", en: "Slice into Grid or Puzzle" }
        },
        panel: {
            slice:   { ja: "分割", en: "Slice" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            mode:        { ja: "分割方法", en: "Method" },
            totalPieces: { ja: "ピース数", en: "Pieces" },
            columns:     { ja: "列数", en: "Columns" },
            rows:        { ja: "行数", en: "Rows" },
            shape:       { ja: "形状", en: "Shape" }
        },
        radio: {
            grid:        { ja: "グリッド", en: "Grid" },
            puzzle:      { ja: "パズル", en: "Puzzle" },
            traditional: { ja: "トラディショナル", en: "Traditional" },
            random:      { ja: "ランダム", en: "Random" }
        },
        checkbox: {
            offset:       { ja: "オフセット", en: "Offset" },
            overlap:      { ja: "オーバーラップ", en: "Overlap" },
            scatter:      { ja: "バラけさせる", en: "Scatter" },
            stroke:       { ja: "ケイ線を追加", en: "Add stroke" },
            roundCorners: { ja: "角丸", en: "Round corners" }
        },
        tooltip: {
            modeGrid:         { ja: "画像を格子状に切り分けます。", en: "Cuts the image into a plain grid." },
            modePuzzle:       { ja: "画像をジグソーパズルのピース状に切り分けます。", en: "Cuts the image into jigsaw puzzle pieces." },
            totalPieces: {
                ja: "作るピースのおおよその数です。オブジェクトを1つ選択しているときは、その縦横比から行数・列数を決めます。",
                en: "Approximate number of pieces. With one object selected, the rows and columns follow from its aspect ratio."
            },
            columns:          { ja: "横に並べるピースの数です。0 にすると行数と縦横比から決めます。", en: "Pieces across. Enter 0 to derive it from the rows and aspect ratio." },
            rows:             { ja: "縦に並べるピースの数です。0 にすると列数と縦横比から決めます。", en: "Pieces down. Enter 0 to derive it from the columns and aspect ratio." },
            shapeTraditional: { ja: "はめ込みの突起を規則的に並べた、よくあるパズル形状にします。", en: "Uses the familiar puzzle shape with regularly placed tabs." },
            shapeRandom:      { ja: "突起の向きをランダムにします。", en: "Randomizes the direction of the tabs." },
            offset:           { ja: "ピースの輪郭をずらします。マイナスで内側に縮みます。", en: "Offsets the outline of each piece. A negative value shrinks it inward." },
            overlap:          { ja: "隣り合うピースを重ねる幅です。継ぎ目のすき間を防ぎます。", en: "How far neighbouring pieces overlap. Use it to hide the seams." },
            scatter:          { ja: "切り分けたピースを少しずつずらして散らします。", en: "Nudges the finished pieces apart so they scatter." },
            scatterDistance:  { ja: "ピースをずらす最大距離です。", en: "Maximum distance a piece moves." },
            stroke:           { ja: "各ピースに線を追加します。", en: "Adds a stroke to each piece." },
            roundCorners:     { ja: "各ピースに効果［角を丸くする］を適用します。", en: "Applies the Round Corners effect to each piece." },
            roundRadius:      { ja: "角丸の半径です。", en: "Corner radius." },
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
            noSelection:       { ja: "分割するオブジェクトを選択してください。", en: "Select the artwork to slice." },
            symbolizeMultiple: { ja: "複数オブジェクトのシンボル化に失敗しました：", en: "Failed to symbolize multiple objects: " },
            symbolizeRaster:   { ja: "埋め込み画像のシンボル化に失敗しました：", en: "Failed to symbolize embedded artwork: " },
            symbolizeVector:   { ja: "ベクターオブジェクトのシンボル化に失敗しました：", en: "Failed to symbolize vector artwork: " },
            maskNotPath: {
                ja: "マスク用のパスを作れなかったため、このピースはマスクせずに残します。",
                en: "No mask path was available, so this piece is left unmasked."
            },
            offsetNoPath: {
                ja: "オフセット後にパスが見つからないため、このピースのマスクをスキップします。",
                en: "No path was found after the offset, so the mask for this piece is skipped."
            },
            offsetFailed:      { ja: "オフセットの適用中にエラーが発生しました：", en: "An error occurred while applying the offset: " },
            scriptError:       { ja: "スクリプトの実行中にエラーが発生しました：", en: "An error occurred while running the script: " }
        }
    };

    // =========================================
    // 入力補助 / Input helpers
    // =========================================

    /**
     * ∧∨と数値欄を隙間なく並べて追加する。↑↓キーも∧∨と同じ処理で増減し、増減後は欄の onChanging を呼ぶ
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} defaultText - 初期値
     * @param {number} characters - 欄の文字数
     * @param {Object} stepOptions - min / max / integer（addStepper() に渡す）
     * @returns {EditText} 追加した数値欄（∧∨は .stepperGroup で参照できる）
     */
    function addStepperInput(parentContainer, defaultText, characters, stepOptions) {
        var stepperInputGroup = parentContainer.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        /* text の代入では onChanging が発火しないので、増減後に呼んで連動を保つ / assigning text does not fire onChanging */
        stepOptions.onStep = function (numberInput) {
            if (typeof numberInput.onChanging === "function") numberInput.onChanging();
        };
        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, defaultText);
        numberInput.characters = characters;
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 数値欄と∧∨の有効／無効をまとめて切り替える
     * @param {EditText} numberInput - addStepperInput() で作った欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup); /* ∧∨は自作描画なので描き直す / redraw the custom-drawn stepper */
    }

    /**
     * 入力欄の数値を pt に換算する（数値でなければ 0）
     * @param {EditText} editText - 定規の単位で入力された欄
     * @returns {number} pt 値
     */
    function readLengthInPoints(editText) {
        var inputValue = parseFloat(editText.text);
        return (isNaN(inputValue) ? 0 : inputValue) * getUnitInfo().pointsPerUnit;
    }

    /**
     * ピース数と縦横比から行数・列数を計算する
     * @param {number} artworkWidth - 対象の幅
     * @param {number} artworkHeight - 対象の高さ
     * @param {number} pieceCount - 作りたいピース数
     * @returns {{rows: number, columns: number}} 行数と列数
     */
    function calcGridSizeFromPieceCount(artworkWidth, artworkHeight, pieceCount) {
        var aspectRatio = artworkWidth / artworkHeight;
        var columns = Math.round(Math.sqrt(pieceCount * aspectRatio));
        if (columns < 1) columns = 1;
        var rows = Math.round(pieceCount / columns);
        if (rows < 1) rows = 1;
        return { rows: rows, columns: columns };
    }

    /**
     * 1つだけ選択している対象の寸法を返す（対象外なら null）
     * @param {Document} doc - 対象ドキュメント
     * @returns {{width: number, height: number}|null} 幅と高さ
     */
    function getSelectedArtworkSize(doc) {
        if (doc.selection.length != 1) return null;
        var selectedItem = doc.selection[0];
        var sizableTypes = { RasterItem: 1, PlacedItem: 1, SymbolItem: 1, PathItem: 1, GroupItem: 1, CompoundPathItem: 1 };
        if (!sizableTypes[selectedItem.typename]) return null;
        var bounds = selectedItem.geometricBounds;
        return { width: bounds[2] - bounds[0], height: Math.abs(bounds[1] - bounds[3]) };
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名＋数値欄の組を追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @param {string} defaultText - 初期値
     * @param {number} characters - 欄の文字数
     * @param {string} tooltipKey - LABELS.tooltip のキー
     * @param {Object} stepOptions - ∧∨の min / max / integer
     * @returns {{label: StaticText, input: EditText}} 追加したコントロール
     */
    function addNumberField(parentContainer, labelKey, defaultText, characters, tooltipKey, stepOptions) {
        var fieldGroup = parentContainer.add("group");
        fieldGroup.orientation = "row";
        var fieldLabel = fieldGroup.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        var fieldInput = addStepperInput(fieldGroup, defaultText, characters, stepOptions);
        fieldInput.helpTip = getLabel("tooltip." + tooltipKey);
        return { label: fieldLabel, input: fieldInput };
    }

    /**
     * チェックボックス＋数値欄＋単位の行を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} checkboxKey - LABELS.checkbox のキー
     * @param {string} defaultText - 初期値
     * @param {number} characters - 欄の文字数
     * @param {string} inputTooltipKey - 数値欄の LABELS.tooltip のキー
     * @param {Object} stepOptions - ∧∨の min / max / integer
     * @returns {{checkbox: Checkbox, input: EditText, unitLabel: StaticText}} 追加したコントロール
     */
    function addCheckboxValueRow(parentPanel, checkboxKey, defaultText, characters, inputTooltipKey, stepOptions) {
        var valueRow = parentPanel.add("group");
        valueRow.orientation = "row";
        valueRow.alignChildren = "left";
        var rowCheckbox = valueRow.add("checkbox", undefined, getLabel("checkbox." + checkboxKey));
        rowCheckbox.helpTip = getLabel("tooltip." + checkboxKey);
        var rowInput = addStepperInput(valueRow, defaultText, characters, stepOptions);
        rowInput.helpTip = getLabel("tooltip." + inputTooltipKey);
        var rowUnitLabel = valueRow.add("statictext", undefined, getUnitInfo().label);
        return { checkbox: rowCheckbox, input: rowInput, unitLabel: rowUnitLabel };
    }

    /**
     * 分割方法（グリッド／パズル）の行を作る
     * @param {Window} dlg - ダイアログ
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildModeRow(dlg, controls) {
        var modeRow = dlg.add("group");
        modeRow.orientation = "row";
        modeRow.alignChildren = "left";
        modeRow.margins = MODE_ROW_MARGINS;
        modeRow.add("statictext", undefined, labelText("fieldLabel.mode"));
        controls.modeGridRadio = modeRow.add("radiobutton", undefined, getLabel("radio.grid"));
        controls.modeGridRadio.helpTip = getLabel("tooltip.modeGrid");
        controls.modePuzzleRadio = modeRow.add("radiobutton", undefined, getLabel("radio.puzzle"));
        controls.modePuzzleRadio.helpTip = getLabel("tooltip.modePuzzle");
        controls.modeGridRadio.value = true;
        controls.modeRow = modeRow;
    }

    /**
     * 分割パネル（ピース数／列数・行数／形状／オフセット／オーバーラップ）を作る
     * @param {Group} parentGroup - 追加先
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildSlicePanel(parentGroup, controls) {
        var slicePanel = parentGroup.add("panel", undefined, getLabel("panel.slice"));
        slicePanel.orientation = "column";
        slicePanel.alignChildren = "left";
        slicePanel.margins = PANEL_MARGINS;

        var totalPiecesField = addNumberField(slicePanel, "totalPieces", DEFAULT_PIECES_PUZZLE, 4, "totalPieces", { integer: true, min: 1 });
        controls.totalPiecesLabel = totalPiecesField.label;
        controls.totalPiecesInput = totalPiecesField.input;

        var gridSizeRow = slicePanel.add("group");
        gridSizeRow.orientation = "row";
        gridSizeRow.alignChildren = "left";
        controls.columnsInput = addNumberField(gridSizeRow, "columns", DEFAULT_COLUMNS, 3, "columns", { integer: true, min: 0 }).input;
        controls.rowsInput = addNumberField(gridSizeRow, "rows", DEFAULT_ROWS, 3, "rows", { integer: true, min: 0 }).input;

        var shapeRow = slicePanel.add("group");
        shapeRow.orientation = "row";
        shapeRow.alignChildren = ["left", "top"];
        shapeRow.margins = SHAPE_ROW_MARGINS;
        shapeRow.add("statictext", undefined, labelText("fieldLabel.shape"));
        var shapeRadioColumn = shapeRow.add("group");
        shapeRadioColumn.orientation = "column";
        shapeRadioColumn.alignChildren = "left";
        controls.shapeTraditionalRadio = shapeRadioColumn.add("radiobutton", undefined, getLabel("radio.traditional"));
        controls.shapeTraditionalRadio.helpTip = getLabel("tooltip.shapeTraditional");
        controls.shapeRandomRadio = shapeRadioColumn.add("radiobutton", undefined, getLabel("radio.random"));
        controls.shapeRandomRadio.helpTip = getLabel("tooltip.shapeRandom");
        controls.shapeTraditionalRadio.value = true;
        controls.shapeRow = shapeRow;

        controls.offsetRow = addCheckboxValueRow(slicePanel, "offset", DEFAULT_OFFSET, 4, "offset", {});
        controls.overlapRow = addCheckboxValueRow(slicePanel, "overlap", DEFAULT_OVERLAP, 4, "overlap", { min: 0 });
        controls.slicePanel = slicePanel;
    }

    /**
     * オプションパネル（バラけ／ケイ線／角丸）を作る
     * @param {Group} parentGroup - 追加先
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildOptionsPanel(parentGroup, controls) {
        var optionsPanel = parentGroup.add("panel", undefined, getLabel("panel.options"));
        optionsPanel.orientation = "column";
        optionsPanel.alignChildren = "left";
        optionsPanel.margins = PANEL_MARGINS;

        controls.scatterRow = addCheckboxValueRow(optionsPanel, "scatter", DEFAULT_SCATTER, 4, "scatterDistance", { min: 0 });

        var strokeRow = optionsPanel.add("group");
        strokeRow.orientation = "row";
        strokeRow.alignChildren = "left";
        controls.strokeCheckbox = strokeRow.add("checkbox", undefined, getLabel("checkbox.stroke"));
        controls.strokeCheckbox.helpTip = getLabel("tooltip.stroke");

        controls.roundCornerRow = addCheckboxValueRow(optionsPanel, "roundCorners", DEFAULT_ROUND_RADIUS, 5, "roundRadius", { min: 0 });
        controls.optionsPanel = optionsPanel;
    }

    /**
     * プログレスバー（処理中のみ表示）とボタンエリアを作る
     * @param {Window} dlg - ダイアログ
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildProgressAndButtons(dlg, controls) {
        /* プログレスバーとボタンを同じ位置に重ね、非表示の行でボタンの上に余白ができないようにする
           Stack the progress bar and buttons so the hidden row adds no gap above the buttons */
        var bottomStack = dlg.add("group");
        bottomStack.orientation = "stack";
        bottomStack.alignment = ["fill", "bottom"];

        var progressRow = bottomStack.add("group");
        progressRow.orientation = "column";
        progressRow.alignment = ["fill", "center"];
        progressRow.alignChildren = "fill";
        progressRow.margins = PROGRESS_ROW_MARGINS;
        controls.progressBar = progressRow.add("progressbar", undefined, 0, 100);
        controls.progressBar.preferredSize = PROGRESS_BAR_SIZE;
        progressRow.visible = false;
        controls.progressRow = progressRow;

        var buttonRow = addButtonRow(bottomStack, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnOK.active = true;
        controls.btnCancel = btnCancel;
        controls.btnOK = btnOK;
        controls.btnRowGroup = buttonRow.rowGroup;
    }

    /**
     * チェックボックス付きの行を有効／無効にする（数値欄はチェック時のみ有効）
     * @param {Object} valueRow - addCheckboxValueRow() の戻り値
     * @param {boolean} rowEnabled - 行を有効にするなら true
     * @returns {void}
     */
    function setValueRowEnabled(valueRow, rowEnabled) {
        valueRow.checkbox.enabled = rowEnabled;
        setStepperInputEnabled(valueRow.input, rowEnabled && valueRow.checkbox.value);
        valueRow.unitLabel.enabled = rowEnabled && valueRow.checkbox.value;
    }

    /**
     * 分割方法とチェック状態に合わせて各コントロールを有効／無効にする
     * @param {Object} controls - ダイアログのコントロール
     * @returns {void}
     */
    function syncEnabledStates(controls) {
        var isPuzzle = controls.modePuzzleRadio.value;
        /* パズル時のみ有効 / Puzzle only */
        controls.totalPiecesLabel.enabled = isPuzzle;
        setStepperInputEnabled(controls.totalPiecesInput, isPuzzle);
        controls.shapeRow.enabled = isPuzzle;
        setValueRowEnabled(controls.offsetRow, isPuzzle);
        setValueRowEnabled(controls.scatterRow, isPuzzle);
        /* グリッド時のみ有効 / Grid only */
        setValueRowEnabled(controls.overlapRow, !isPuzzle);
        setValueRowEnabled(controls.roundCornerRow, !isPuzzle);
    }

    /**
     * 分割方法ごとの初期値に戻す
     * @param {Object} controls - ダイアログのコントロール
     * @returns {void}
     */
    function applyModeDefaults(controls) {
        controls.totalPiecesInput.text = controls.modeGridRadio.value ? DEFAULT_PIECES_GRID : DEFAULT_PIECES_PUZZLE;
        controls.offsetRow.checkbox.value = false;
        controls.offsetRow.input.text = DEFAULT_OFFSET;
        controls.overlapRow.checkbox.value = false;
        controls.overlapRow.input.text = DEFAULT_OVERLAP;
        controls.scatterRow.checkbox.value = false;
        controls.scatterRow.input.text = DEFAULT_SCATTER;
        controls.strokeCheckbox.value = false;
        controls.roundCornerRow.checkbox.value = false;
        controls.roundCornerRow.input.text = DEFAULT_ROUND_RADIUS;
    }

    /**
     * ダイアログのイベントを結び付け、初期状態を整える
     * @param {Object} controls - ダイアログのコントロール
     * @param {{width: number, height: number}|null} artworkSize - 選択対象の寸法
     * @returns {void}
     */
    function bindDialogEvents(controls, artworkSize) {
        /* ピース数から行数・列数を決める / Derive rows and columns from the piece count */
        function updateGridSizeFromPieces() {
            var pieceCount = parseInt(controls.totalPiecesInput.text, 10);
            if (isNaN(pieceCount) || pieceCount < 1 || !artworkSize) return;
            var gridSize = calcGridSizeFromPieceCount(artworkSize.width, artworkSize.height, pieceCount);
            controls.rowsInput.text = String(gridSize.rows);
            controls.columnsInput.text = String(gridSize.columns);
        }

        function onSyncEnabled() {
            syncEnabledStates(controls);
        }

        function onModeChange() {
            applyModeDefaults(controls);
            updateGridSizeFromPieces();
            syncEnabledStates(controls);
        }

        controls.totalPiecesInput.onChanging = updateGridSizeFromPieces;
        controls.modeGridRadio.onClick = onModeChange;
        controls.modePuzzleRadio.onClick = onModeChange;
        controls.offsetRow.checkbox.onClick = onSyncEnabled;
        controls.overlapRow.checkbox.onClick = onSyncEnabled;
        controls.scatterRow.checkbox.onClick = onSyncEnabled;
        controls.roundCornerRow.checkbox.onClick = onSyncEnabled;

        onModeChange();
    }

    /**
     * ダイアログを作る
     * @param {{width: number, height: number}|null} artworkSize - 選択対象の寸法
     * @returns {Object} ダイアログ（dialog）と各コントロール
     */
    function buildDialog(artworkSize) {
        var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        var controls = { dialog: dlg };

        buildModeRow(dlg, controls);

        var panelStack = dlg.add("group");
        panelStack.orientation = "column";
        panelStack.alignChildren = "fill";
        buildSlicePanel(panelStack, controls);
        buildOptionsPanel(panelStack, controls);

        buildProgressAndButtons(dlg, controls);
        bindDialogEvents(controls, artworkSize);
        return controls;
    }

    /**
     * ダイアログの値を分割の設定として読み取る（分割方法で使わない項目は無効扱い）
     * @param {Object} controls - ダイアログのコントロール
     * @returns {Object} 分割の設定
     */
    function readSliceSettings(controls) {
        var isGridMode = controls.modeGridRadio.value;
        return {
            isGridMode: isGridMode,
            isRandomShape: controls.shapeRandomRadio.value,
            columnCount: Math.round(Number(controls.columnsInput.text)),
            rowCount: Math.round(Number(controls.rowsInput.text)),
            shouldApplyOffset: !isGridMode && controls.offsetRow.checkbox.value,
            offsetInPoints: readLengthInPoints(controls.offsetRow.input),
            overlapInPoints: (isGridMode && controls.overlapRow.checkbox.value) ? readLengthInPoints(controls.overlapRow.input) : 0,
            shouldScatter: !isGridMode && controls.scatterRow.checkbox.value,
            scatterDistance: readLengthInPoints(controls.scatterRow.input),
            shouldAddStroke: controls.strokeCheckbox.value,
            shouldApplyRoundCorners: isGridMode && controls.roundCornerRow.checkbox.value,
            roundRadiusInPoints: readLengthInPoints(controls.roundCornerRow.input)
        };
    }

    /**
     * 処理中の表示（入力を無効化してプログレスバーを出す）に切り替える
     * @param {Object} controls - ダイアログのコントロール
     * @returns {void}
     */
    function showProgressState(controls) {
        controls.modeRow.enabled = false;
        controls.slicePanel.enabled = false;
        controls.optionsPanel.enabled = false;
        redrawSteppersIn(controls.slicePanel);
        redrawSteppersIn(controls.optionsPanel);
        controls.btnRowGroup.visible = false;
        controls.progressRow.visible = true;
        controls.progressBar.value = 0;
        controls.dialog.layout.layout(true);
        controls.dialog.update();
    }

    // =========================================
    // 元オブジェクトの準備 / Source preparation
    // =========================================

    /**
     * オブジェクトをシンボル化し、同じ位置にインスタンスを置いて元を削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} sourceItem - シンボル化する対象
     * @returns {SymbolItem} 置き換えたインスタンス
     */
    function convertToSymbolItem(doc, sourceItem) {
        var sourceLeft = sourceItem.left;
        var sourceTop = sourceItem.top;
        var sourceParent = sourceItem.parent;
        var createdSymbol = doc.symbols.add(sourceItem);
        var symbolInstance = sourceParent.symbolItems.add(createdSymbol);
        symbolInstance.left = sourceLeft;
        symbolInstance.top = sourceTop;
        sourceItem.remove();
        return symbolInstance;
    }

    /**
     * 選択中のオブジェクトを重ね順を保ったまま1つのグループにまとめる
     * @param {Document} doc - 対象ドキュメント
     * @returns {GroupItem} まとめたグループ
     */
    function groupSelectedItems(doc) {
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }
        var selectionGroup = doc.groupItems.add();
        for (var j = selectedItems.length - 1; j >= 0; j--) {
            selectedItems[j].move(selectionGroup, ElementPlacement.PLACEATBEGINNING);
        }
        return selectionGroup;
    }

    /**
     * 選択を分割できる形に整える（複数選択・埋め込み画像・ベクターはシンボル化、
     * 画像やシンボルはマスク用の矩形を用意）
     * @param {Document} doc - 対象ドキュメント
     * @returns {{contentSourceItem: PageItem, maskSourceItem: PageItem, isTemporaryBoundsRect: boolean}|null} 準備した対象（失敗時は null）
     */
    function prepareSourceItems(doc) {
        var vectorTypes = { PathItem: 1, GroupItem: 1, CompoundPathItem: 1 };
        var failureAlertKey = "symbolizeMultiple";
        var workingItem;

        try {
            if (doc.selection.length > 1) {
                workingItem = convertToSymbolItem(doc, groupSelectedItems(doc));
            } else {
                workingItem = doc.selection[0];
                if (workingItem.typename === "RasterItem") {
                    failureAlertKey = "symbolizeRaster";
                    workingItem = convertToSymbolItem(doc, workingItem);
                } else if (vectorTypes[workingItem.typename]) {
                    failureAlertKey = "symbolizeVector";
                    workingItem = convertToSymbolItem(doc, workingItem);
                }
            }
        } catch (e) {
            alert(getLabel("alert." + failureAlertKey) + e);
            return null;
        }

        /* 配置画像とシンボルは外接矩形をマスクの元にする / Placed images and symbols get a bounding rectangle as the mask source */
        if (workingItem.typename === "PlacedItem" || workingItem.typename === "SymbolItem") {
            var sourceBounds = workingItem.geometricBounds;
            var rectWidth = sourceBounds[2] - sourceBounds[0];
            var rectHeight = Math.abs(sourceBounds[1] - sourceBounds[3]);
            return {
                contentSourceItem: workingItem,
                maskSourceItem: doc.pathItems.rectangle(sourceBounds[1], sourceBounds[0], rectWidth, rectHeight),
                isTemporaryBoundsRect: true
            };
        }
        return { contentSourceItem: null, maskSourceItem: workingItem, isTemporaryBoundsRect: false };
    }

    /**
     * 分割の元にしたオブジェクトと一時矩形を削除する
     * @param {Object} preparedItems - prepareSourceItems() の戻り値
     * @returns {void}
     */
    function cleanupSourceItems(preparedItems) {
        if (preparedItems.contentSourceItem) preparedItems.contentSourceItem.remove();
        if (preparedItems.isTemporaryBoundsRect) preparedItems.maskSourceItem.remove();
    }

    /**
     * 0 の列数／行数を、もう一方と縦横比から補う
     * @param {number} columnCount - 列数（0 なら自動）
     * @param {number} rowCount - 行数（0 なら自動）
     * @param {number[]} bounds - マスク元の geometricBounds
     * @returns {{columnCount: number, rowCount: number}} 補った列数と行数
     */
    function resolveGridCounts(columnCount, rowCount, bounds) {
        var boundsWidth = bounds[2] - bounds[0];
        var boundsHeight = bounds[3] - bounds[1];
        if (columnCount == 0) {
            columnCount = Math.max(1, Math.round(Math.abs(boundsWidth / boundsHeight) * rowCount));
        }
        if (rowCount == 0) {
            rowCount = Math.max(1, Math.round(Math.abs(boundsHeight / boundsWidth) * columnCount));
        }
        return { columnCount: columnCount, rowCount: rowCount };
    }

    /**
     * 列数・行数の入力が分割できる組み合わせか（どちらかが1以上、もう一方は0以上）
     * @param {number} columnCount - 列数
     * @param {number} rowCount - 行数
     * @returns {boolean} 分割できるなら true
     */
    function isValidGridCount(columnCount, rowCount) {
        return columnCount >= 0 && rowCount >= 0 && (columnCount >= 1 || rowCount >= 1);
    }

    // =========================================
    // マスク形状 / Mask shapes
    // =========================================

    /**
     * 分割の基準となる格子を作る（パズル時は各ピースの突起の向きとずれも決める）
     * @param {number[]} bounds - マスク元の geometricBounds
     * @param {number} columnCount - 列数
     * @param {number} rowCount - 行数
     * @param {boolean} isPuzzle - パズル形状なら true
     * @param {boolean} isRandomShape - 突起の向きをランダムにするなら true
     * @returns {Object} 格子の情報
     */
    function buildSliceGrid(bounds, columnCount, rowCount, isPuzzle, isRandomShape) {
        var sliceGrid = {
            originX: bounds[0],
            originY: bounds[1],
            right: bounds[2],
            bottom: bounds[3],
            columnCount: columnCount,
            rowCount: rowCount,
            pieceWidth: (bounds[2] - bounds[0]) / columnCount,
            pieceHeight: (bounds[1] - bounds[3]) / rowCount,
            edgeData: null
        };
        if (!isPuzzle) return sliceGrid;

        var edgeData = new Array(rowCount);
        for (var y = 0; y < rowCount; y++) {
            edgeData[y] = new Array(columnCount);
            for (var x = 0; x < columnCount; x++) {
                var isTopOut = isRandomShape ? Math.random() < 0.5 : (x & 1) ^ (y & 1);
                var isRightOut = isRandomShape ? Math.random() < 0.5 : !((x & 1) ^ (y & 1));
                edgeData[y][x] = {
                    topOut: isTopOut,
                    rightOut: isRightOut,
                    verticalOffset: sliceGrid.pieceHeight * (Math.random() - 0.5) / 10,
                    horizontalOffset: sliceGrid.pieceWidth * (Math.random() - 0.5) / 10
                };
            }
        }
        sliceGrid.edgeData = edgeData;
        return sliceGrid;
    }

    /**
     * グリッド1マス分の矩形マスクを作る（オーバーラップ分広げ、元の範囲内に収める）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} columnIndex - 列番号
     * @param {number} rowIndex - 行番号
     * @param {number} overlapInPoints - オーバーラップ（pt）
     * @returns {PathItem} 矩形マスク
     */
    function createGridMask(doc, sliceGrid, columnIndex, rowIndex, overlapInPoints) {
        var rectLeft = sliceGrid.originX + columnIndex * sliceGrid.pieceWidth - overlapInPoints / 2;
        var rectWidth = sliceGrid.pieceWidth + overlapInPoints;
        if (rectLeft < sliceGrid.originX) {
            rectWidth -= (sliceGrid.originX - rectLeft);
            rectLeft = sliceGrid.originX;
        }
        if (rectLeft + rectWidth > sliceGrid.right) {
            rectWidth = sliceGrid.right - rectLeft;
        }

        var rectTop = sliceGrid.originY - rowIndex * sliceGrid.pieceHeight + overlapInPoints / 2;
        var rectHeight = sliceGrid.pieceHeight + overlapInPoints;
        if (rectTop > sliceGrid.originY) {
            rectHeight -= (rectTop - sliceGrid.originY);
            rectTop = sliceGrid.originY;
        }
        if (rectTop - rectHeight < sliceGrid.bottom) {
            rectHeight = rectTop - sliceGrid.bottom;
        }

        var gridMask = doc.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
        gridMask.closed = true;
        gridMask.filled = false;
        gridMask.stroked = false;
        return gridMask;
    }

    /**
     * コーナーポイントを追加する
     * @param {PathItem} pathItem - 対象パス
     * @param {number} anchorX - X座標
     * @param {number} anchorY - Y座標
     * @returns {void}
     */
    function addCornerPoint(pathItem, anchorX, anchorY) {
        var cornerPoint = pathItem.pathPoints.add();
        cornerPoint.anchor = [anchorX, anchorY];
        cornerPoint.leftDirection = [anchorX, anchorY];
        cornerPoint.rightDirection = [anchorX, anchorY];
        cornerPoint.pointType = PointType.CORNER;
    }

    /**
     * スムーズポイントを追加する
     * @param {PathItem} pathItem - 対象パス
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} leftX - 前側ハンドルのX座標
     * @param {number} leftY - 前側ハンドルのY座標
     * @param {number} rightX - 後側ハンドルのX座標
     * @param {number} rightY - 後側ハンドルのY座標
     * @returns {void}
     */
    function addCurvePoint(pathItem, anchorX, anchorY, leftX, leftY, rightX, rightY) {
        var curvePoint = pathItem.pathPoints.add();
        curvePoint.anchor = [anchorX, anchorY];
        curvePoint.leftDirection = [leftX, leftY];
        curvePoint.rightDirection = [rightX, rightY];
        curvePoint.pointType = PointType.SMOOTH;
    }

    /**
     * パズルピース1つ分のマスクパスを作る（下辺→右辺→上辺→左辺の順に突起を描く）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {PathItem} マスクパス
     */
    function createPuzzleMaskPath(doc, sliceGrid, x, y) {
        var leftX = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rightX = sliceGrid.originX + (x + 1) * sliceGrid.pieceWidth;
        var topY = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var bottomY = sliceGrid.originY - (y + 1) * sliceGrid.pieceHeight;

        var maskPath = doc.pathItems.add();
        addCornerPoint(maskPath, leftX, bottomY);
        appendLowerTab(maskPath, sliceGrid, x, y);
        addCornerPoint(maskPath, rightX, bottomY);
        appendRightTab(maskPath, sliceGrid, x, y);
        addCornerPoint(maskPath, rightX, topY);
        appendUpperTab(maskPath, sliceGrid, x, y);
        addCornerPoint(maskPath, leftX, topY);
        appendLeftTab(maskPath, sliceGrid, x, y);
        maskPath.closed = true;
        return maskPath;
    }

    /**
     * 下隣のピースとの境界（y+1 行目との辺）の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendLowerTab(maskPath, sliceGrid, x, y) {
        if (y >= sliceGrid.rowCount - 1) return;
        var edge = sliceGrid.edgeData[y + 1][x];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var rowBottom = sliceGrid.originY - (y + 1) * sliceGrid.pieceHeight;
        var thirdWidth = sliceGrid.pieceWidth / 3;
        var quarterHeight = sliceGrid.pieceHeight / 4;
        var shift = edge.verticalOffset;

        if (edge.topOut) {
            addCurvePoint(maskPath,
                left + thirdWidth, rowBottom - shift,
                left + 0.67 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowBottom - quarterHeight - shift,
                left + 0.67 * thirdWidth, rowBottom - quarterHeight + 0.33 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowBottom - quarterHeight - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowBottom - quarterHeight - shift,
                left + 1.67 * thirdWidth, rowBottom - quarterHeight - 0.33 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowBottom - quarterHeight + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowBottom - shift,
                left + 1.67 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift);
        } else {
            addCurvePoint(maskPath,
                left + thirdWidth, rowBottom - shift,
                left + 0.67 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - 3 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop - 3.5 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowTop - 2.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - 3 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop - 2.5 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowTop - 3.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowBottom - shift,
                left + 1.67 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift);
        }
    }

    /**
     * 右隣のピースとの境界の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendRightTab(maskPath, sliceGrid, x, y) {
        if (x >= sliceGrid.columnCount - 1) return;
        var edge = sliceGrid.edgeData[y][x + 1];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var thirdHeight = sliceGrid.pieceHeight / 3;
        var quarterWidth = sliceGrid.pieceWidth / 4;
        var shift = edge.horizontalOffset;

        if (edge.rightOut) {
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 5 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 5.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 5 * quarterWidth - shift, rowTop - thirdHeight,
                left + 5.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
        } else {
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 3 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 2.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 3 * quarterWidth - shift, rowTop - thirdHeight,
                left + 2.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
        }
    }

    /**
     * 上隣のピースとの境界（y 行目の上辺）の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendUpperTab(maskPath, sliceGrid, x, y) {
        if (y <= 0) return;
        var edge = sliceGrid.edgeData[y][x];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var thirdWidth = sliceGrid.pieceWidth / 3;
        var quarterHeight = sliceGrid.pieceHeight / 4;
        var shift = edge.verticalOffset;

        if (edge.topOut) {
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - shift,
                left + 2.33 * thirdWidth, rowTop + 0.33 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - quarterHeight - shift,
                left + 2.33 * thirdWidth, rowTop - quarterHeight + 0.33 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop - quarterHeight - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - quarterHeight - shift,
                left + 1.33 * thirdWidth, rowTop - quarterHeight - 0.33 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop - quarterHeight + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - shift,
                left + 1.33 * thirdWidth, rowTop - 0.33 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop + 0.33 * quarterHeight - shift);
        } else {
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - shift,
                left + 2.33 * thirdWidth, rowTop - 0.33 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop + quarterHeight - shift,
                left + 2.33 * thirdWidth, rowTop + 0.5 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop + 1.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop + quarterHeight - shift,
                left + 1.33 * thirdWidth, rowTop + 1.5 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop + 0.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - shift,
                left + 1.33 * thirdWidth, rowTop + 0.33 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop - 0.33 * quarterHeight - shift);
        }
    }

    /**
     * 左隣のピースとの境界の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendLeftTab(maskPath, sliceGrid, x, y) {
        if (x <= 0) return;
        var edge = sliceGrid.edgeData[y][x];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var thirdHeight = sliceGrid.pieceHeight / 3;
        var quarterWidth = sliceGrid.pieceWidth / 4;
        var shift = edge.horizontalOffset;

        if (edge.rightOut) {
            addCurvePoint(maskPath,
                left - shift, rowTop - thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left + quarterWidth - shift, rowTop - thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left + 1.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left + quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 1.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - shift, rowTop - 2 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
        } else {
            addCurvePoint(maskPath,
                left - shift, rowTop - thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - quarterWidth - shift, rowTop - thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left - 1.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - quarterWidth - shift, rowTop - 2 * thirdHeight,
                left - 1.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - shift, rowTop - 2 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
        }
    }

    // =========================================
    // オフセット / Offset
    // =========================================

    /**
     * 効果［パスのオフセット］を適用して分割し、分割後のオブジェクトを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} maskPath - 対象のマスクパス（処理後は削除される）
     * @param {number} offsetInPoints - オフセット量（pt）
     * @returns {PageItem|null} 分割後のオブジェクト
     */
    function applyOffsetEffect(doc, maskPath, offsetInPoints) {
        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        try {
            doc.selection = null;
            var offsetTarget = maskPath.duplicate(maskPath, ElementPlacement.PLACEAFTER);
            maskPath.remove();
            offsetTarget.selected = true;
            offsetTarget.applyEffect('<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetInPoints + ' I jntp 2 "/></LiveEffect>');
            app.redraw();
            app.executeMenuCommand("expandStyle");
            var expandedItem = (doc.selection.length > 0) ? doc.selection[0] : null;
            doc.selection = null;
            return expandedItem;
        } finally {
            /* 例外時も警告の表示設定を戻す / Restore the alert level even on failure */
            app.userInteractionLevel = previousInteractionLevel;
        }
    }

    /**
     * オブジェクトからマスクに使えるパスを取り出す（グループなら最初のパス）
     * @param {PageItem|null} targetItem - 対象
     * @returns {PathItem|null} 見つかったパス
     */
    function findMaskPath(targetItem) {
        if (!targetItem) return null;
        if (targetItem.typename === "PathItem") return targetItem;
        if (targetItem.typename === "GroupItem") {
            for (var i = 0; i < targetItem.pageItems.length; i++) {
                if (targetItem.pageItems[i].typename === "PathItem") return targetItem.pageItems[i];
            }
        }
        return null;
    }

    /**
     * マスクパスにオフセットを掛けたパスを返す（パスが得られなければ警告して元の結果を返す）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} maskPath - 対象のマスクパス
     * @param {number} offsetInPoints - オフセット量（pt）
     * @returns {PageItem|null} マスクに使うオブジェクト
     */
    function offsetMaskPath(doc, maskPath, offsetInPoints) {
        try {
            var expandedItem = applyOffsetEffect(doc, maskPath, offsetInPoints);
            var offsetPath = findMaskPath(expandedItem);
            if (offsetPath) return offsetPath;
            alert(getLabel("alert.offsetNoPath"));
            return expandedItem;
        } catch (e) {
            alert(getLabel("alert.offsetFailed") + e.message);
            return maskPath;
        }
    }

    // =========================================
    // ピースの仕上げ / Piece finishing
    // =========================================

    /**
     * 中央寄りの正規乱数を返す（Box-Muller 法）
     * @returns {number} 平均0・標準偏差1の乱数
     */
    function randomGaussian() {
        var u = 0, v = 0;
        while (u === 0) u = Math.random();
        while (v === 0) v = Math.random();
        return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
    }

    /**
     * 値を範囲内に収める
     * @param {number} value - 対象の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clampValue(value, minValue, maxValue) {
        return Math.max(minValue, Math.min(maxValue, value));
    }

    /**
     * 元オブジェクトの複製をマスクパスでクリップしたグループを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} contentSourceItem - 中身にする配置画像／シンボル
     * @param {PageItem|null} maskPath - マスクにするパス
     * @returns {GroupItem} クリップグループ
     */
    function createClippedPiece(doc, contentSourceItem, maskPath) {
        var contentCopy = contentSourceItem.duplicate();
        var clippingGroup = doc.groupItems.add();
        contentCopy.moveToBeginning(clippingGroup);
        if (maskPath && maskPath.typename === "PathItem") {
            maskPath.moveToBeginning(clippingGroup);
            maskPath.clipping = true;
            clippingGroup.clipped = true;
        } else {
            alert(getLabel("alert.maskNotPath"));
        }
        return clippingGroup;
    }

    /**
     * ピースにバラけ・ケイ線・角丸を適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem|null} pieceItem - 対象のピース
     * @param {Object} sliceSettings - readSliceSettings() の戻り値
     * @returns {void}
     */
    function finishPiece(doc, pieceItem, sliceSettings) {
        if (!pieceItem) return;
        if (sliceSettings.shouldScatter && sliceSettings.scatterDistance > 0) {
            /* ガウス分布（中央寄り）で動かし、最大移動量で抑える / Gaussian nudge clamped to the maximum distance */
            var maxDistance = sliceSettings.scatterDistance;
            var dx = clampValue(randomGaussian() * maxDistance * 0.5, -maxDistance, maxDistance);
            var dy = clampValue(randomGaussian() * maxDistance * 0.5, -maxDistance, maxDistance);
            pieceItem.transform(app.getTranslationMatrix(dx, dy));
        }
        pieceItem.selected = true;

        if (sliceSettings.shouldAddStroke) {
            doc.selection = null;
            pieceItem.selected = true;
            app.executeMenuCommand("Adobe New Stroke Shortcut");
            app.executeMenuCommand("Live Pathfinder Add");
        }
        if (sliceSettings.shouldApplyRoundCorners && sliceSettings.roundRadiusInPoints > 0) {
            pieceItem.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + sliceSettings.roundRadiusInPoints + ' "/></LiveEffect>');
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを分割してピースを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} sliceSettings - readSliceSettings() の戻り値
     * @param {Function} onProgress - 進み具合を受け取る関数 (完了数, 総数)
     * @returns {void}
     */
    function executeSlice(doc, sliceSettings, onProgress) {
        var preparedItems = prepareSourceItems(doc);
        if (!preparedItems) return;

        var maskSourceItem = preparedItems.maskSourceItem;
        var contentSourceItem = preparedItems.contentSourceItem;
        maskSourceItem.selected = false;

        var bounds = maskSourceItem.geometricBounds;
        var gridCounts = resolveGridCounts(sliceSettings.columnCount, sliceSettings.rowCount, bounds);
        var sliceGrid = buildSliceGrid(bounds, gridCounts.columnCount, gridCounts.rowCount,
            !sliceSettings.isGridMode, sliceSettings.isRandomShape);

        var totalPieces = sliceGrid.rowCount * sliceGrid.columnCount;
        var finishedPieces = 0;
        onProgress(0, totalPieces);

        for (var y = 0; y < sliceGrid.rowCount; y++) {
            for (var x = 0; x < sliceGrid.columnCount; x++) {
                var maskPath;
                if (sliceSettings.isGridMode) {
                    maskPath = createGridMask(doc, sliceGrid, x, y, sliceSettings.overlapInPoints);
                } else {
                    maskPath = createPuzzleMaskPath(doc, sliceGrid, x, y);
                    if (sliceSettings.shouldApplyOffset) {
                        maskPath = offsetMaskPath(doc, maskPath, sliceSettings.offsetInPoints);
                    }
                }

                var pieceItem = contentSourceItem ? createClippedPiece(doc, contentSourceItem, maskPath) : maskPath;
                finishPiece(doc, pieceItem, sliceSettings);

                finishedPieces++;
                onProgress(finishedPieces, totalPieces);
            }
        }

        cleanupSourceItems(preparedItems);
    }

    /**
     * エントリーポイント：ダイアログを表示し、OK で分割を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var controls = buildDialog(getSelectedArtworkSize(doc));
        var sliceDialog = controls.dialog;

        controls.btnOK.onClick = function () {
            var sliceSettings = readSliceSettings(controls);
            /* 分割できない列数・行数なら閉じずに待つ / Stay open when the counts cannot be sliced */
            if (!isValidGridCount(sliceSettings.columnCount, sliceSettings.rowCount)) return;

            showProgressState(controls);
            try {
                executeSlice(doc, sliceSettings, function (current, total) {
                    controls.progressBar.value = total > 0 ? (current / total) * 100 : 0;
                    sliceDialog.update();
                });
            } catch (e) {
                alert(getLabel("alert.scriptError") + e);
            }
            sliceDialog.close(1);
        };

        prepareDialogWindow(sliceDialog, SCRIPT_NAME);
        sliceDialog.show();
    }

    main();

})();
