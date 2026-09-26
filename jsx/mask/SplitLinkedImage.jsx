#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したリンク画像を、指定した列数・行数のグリッドに分割します。
分割後の各画像はクリッピングマスクで切り出され、元の配置を保ったまま並びます。

詳細は README を参照してください。

### Overview

Splits the selected linked image into a grid of the given number of columns and rows.
Each piece is cut out with a clipping mask and keeps the original placement.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitLinkedImage";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "リンク画像を分割",
            en: "Split Linked Image"
        },
        countLR: {
            ja: "左右",
            en: "Columns"
        },
        countTB: {
            ja: "上下",
            en: "Rows"
        },
        overlap: {
            ja: "オーバーラップ",
            en: "Overlap"
        },
        groupCheck: {
            ja: "グループ化",
            en: "Create Group"
        },
        tipColumns: { ja: "横に何枚へ分割するかです。", en: "How many pieces to cut the image into across." },
        tipRows: { ja: "縦に何枚へ分割するかです。", en: "How many pieces to cut the image into down." },
        tipOverlap: { ja: "隣り合う分割片を重ねる幅です。継ぎ目を目立たせたくないときに使います。", en: "How far neighbouring pieces overlap. Use it to hide the seams." },
        tipGroupCheck: { ja: "分割してできた画像を1つのグループにまとめます。", en: "Groups the resulting pieces into a single group." },
        tipRuleCheck: { ja: "分割した境目にケイ線を引きます。", en: "Draws a rule along each cut." },
        tipRoundCheck: { ja: "分割片の角を丸めます。半径は右の欄で指定します。", en: "Rounds the corners of each piece. The field on the right sets the radius." },
        tipRoundRadius: { ja: "角丸の半径です。", en: "Radius of the rounded corners." },
        tipStepUp: {
            ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        tipStepDown: {
            ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        tipStepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        tipStepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
        optionsPanel: {
            ja: "オプション",
            en: "Options"
        },
        splitPanel: {
            ja: "分割数",
            en: "Split Count"
        },
        ruleCheck: {
            ja: "ケイ",
            en: "Add Stroke"
        },
        roundCheck: {
            ja: "角丸",
            en: "Apply Round Corners"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        errNoDoc: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        errNoSelection: {
            ja: "リンク画像を1つ以上選択してください。",
            en: "Please select one or more linked images."
        },
        errLR: {
            ja: "左右の分割数は1以上の整数を入力してください。",
            en: "Columns must be an integer of 1 or more."
        },
        errTB: {
            ja: "上下の分割数は1以上の整数を入力してください。",
            en: "Rows must be an integer of 1 or more."
        },
        errBoth: {
            ja: "左右または上下のいずれかを2以上にしてください。",
            en: "Either columns or rows must be 2 or more."
        },
        errOverlap: {
            ja: "オーバーラップは0以上の数値を入力してください。",
            en: "Overlap must be a number of 0 or more."
        },
        processed: {
            ja: "分割したリンク画像",
            en: "Processed"
        },
        skipped: {
            ja: "スキップ",
            en: "Skipped"
        },
        unit: {
            ja: "個",
            en: " item(s)"
        }
    };

    function getLabel(key) {
        return LABELS[key][uiLang];
    }

    function labelText(key) {
        return getLabel(key) + (uiLang === 'ja' ? '：' : ':');
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    // 2. コピー先の LABELS に tipStepUp / tipStepDown / tipStepUpInteger / tipStepDownInteger を足す（このファイルの LABELS から写す）。
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
    // 6. この欄に↑↓キーの増減処理を別に付けない（↑↓キーが二重に効く）
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
        var upTooltip = stepOptions.integer ? "tipStepUpInteger" : "tipStepUp";
        var downTooltip = stepOptions.integer ? "tipStepDownInteger" : "tipStepDown";
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
        if (direction > 0) return Math.floor(value / multiple) * multiple + multiple;
        return Math.ceil(value / multiple) * multiple - multiple;
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

    /**
     * ∧∨と数値入力欄を隙間0で突き合わせて追加し、↑↓キーも∧∨と同じ処理で増減させる
     * @param {Group} parentGroup - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {number} characters - 入力欄の文字数
     * @param {Object} stepOptions - ∧∨の設定（min / integer）
     * @returns {EditText} 追加した入力欄
     */
    function addStepperInput(parentGroup, initialText, characters, stepOptions) {
        var stepperFieldGroup = parentGroup.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperFieldGroup.add("edittext", undefined, initialText);
        numberInput.characters = characters;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    // =========================================
    // 事前チェック / Pre-check
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel('errNoDoc'));
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {
        alert(getLabel('errNoSelection'));
        return;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    var dlg = new Window("dialog", getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var splitPanel = dlg.add("panel", undefined, getLabel('splitPanel'));
    splitPanel.orientation = "column";
    splitPanel.alignChildren = "left";
    splitPanel.margins = [15, 20, 15, 10];

    var countLRGroup = splitPanel.add("group");
    countLRGroup.orientation = "row";
    countLRGroup.add("statictext", undefined, labelText('countLR'));
    /* 分割数は1以上の整数 / split counts are integers of 1 or more */
    var columnsInput = addStepperInput(countLRGroup, "2", 4, { integer: true, min: 1 });
    columnsInput.helpTip = getLabel('tipColumns');

    var countTBGroup = splitPanel.add("group");
    countTBGroup.orientation = "row";
    countTBGroup.add("statictext", undefined, labelText('countTB'));
    var rowsInput = addStepperInput(countTBGroup, "1", 4, { integer: true, min: 1 });
    rowsInput.helpTip = getLabel('tipRows');

    var optionsPanel = dlg.add("panel", undefined, getLabel('optionsPanel'));
    optionsPanel.orientation = "column";
    optionsPanel.alignChildren = "left";
    optionsPanel.margins = [15, 20, 15, 10];

    var overlapGroup = optionsPanel.add("group");
    overlapGroup.orientation = "row";
    overlapGroup.add("statictext", undefined, labelText('overlap'));
    var overlapInput = addStepperInput(overlapGroup, "0", 4, { min: 0 });
    overlapInput.helpTip = getLabel('tipOverlap');
    overlapGroup.add("statictext", undefined, getUnitInfo().label);

    var groupCheck = optionsPanel.add("checkbox", undefined, getLabel('groupCheck'));
    groupCheck.helpTip = getLabel('tipGroupCheck');
    groupCheck.value = false;

    var ruleCheck = optionsPanel.add("checkbox", undefined, getLabel('ruleCheck'));
    ruleCheck.helpTip = getLabel('tipRuleCheck');
    ruleCheck.value = false;

    var roundGroup = optionsPanel.add("group");
    roundGroup.orientation = "row";
    var roundCheck = roundGroup.add("checkbox", undefined, getLabel('roundCheck'));
    roundCheck.helpTip = getLabel('tipRoundCheck');
    roundCheck.value = false;
    var roundRadiusInput = addStepperInput(roundGroup, "3", 5, { min: 0 });
    roundRadiusInput.helpTip = getLabel('tipRoundRadius');
    roundGroup.add("statictext", undefined, getUnitInfo().label);

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    btnGroup.add("button", undefined, getLabel('cancel'), { name: "cancel" });
    btnGroup.add("button", undefined, getLabel('ok'), { name: "ok" });

    if (dlg.show() != 1) {
        return;
    }

    var columnCount = parseInt(columnsInput.text, 10);
    var rowCount = parseInt(rowsInput.text, 10);
    /* オーバーラップは rulerType 入力、内部では pt 換算 / Overlap input in ruler unit, used in pt */
    var overlapInputValue = parseFloat(overlapInput.text);
    var overlapInPoints = (isNaN(overlapInputValue) ? NaN : overlapInputValue * getUnitInfo().pointsPerUnit);
    var shouldCreateGroup = groupCheck.value;
    var shouldAddStroke = ruleCheck.value;
    var shouldApplyRoundCorners = roundCheck.value;
    var roundRadiusInputValue = parseFloat(roundRadiusInput.text);
    /* rulerType → pt 変換 / Convert ruler unit to pt */
    var roundRadiusInPoints = (isNaN(roundRadiusInputValue) ? 0 : roundRadiusInputValue) * getUnitInfo().pointsPerUnit;

    if (isNaN(columnCount) || columnCount < 1) {
        alert(getLabel('errLR'));
        return;
    }

    if (isNaN(rowCount) || rowCount < 1) {
        alert(getLabel('errTB'));
        return;
    }

    if (columnCount < 2 && rowCount < 2) {
        alert(getLabel('errBoth'));
        return;
    }

    if (isNaN(overlapInPoints) || overlapInPoints < 0) {
        alert(getLabel('errOverlap'));
        return;
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================
    /* 選択項目を退避 / Save selection */
    var selectedItems = [];
    for (var selectedIndex = 0; selectedIndex < doc.selection.length; selectedIndex++) {
        selectedItems.push(doc.selection[selectedIndex]);
    }

    var processedCount = 0;
    var skippedCount = 0;
    /* 生成アイテムを保持 / Keep created items for reselection */
    var createdItems = [];

    for (var selectedIndex = 0; selectedIndex < selectedItems.length; selectedIndex++) {
        var selectedItem = selectedItems[selectedIndex];

        if (selectedItem.typename !== "PlacedItem") {
            skippedCount++;
            continue;
        }

        try {
            var left = selectedItem.left;
            var top = selectedItem.top;
            var width = selectedItem.width;
            var height = selectedItem.height;

            var partWidth = width / columnCount;
            var partHeight = height / rowCount;

            /* グループ化ON時は親グループを作成 / Create parent group when Group is ON */
            var parentGroup = shouldCreateGroup ? doc.groupItems.add() : null;

            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    var clippingGroup = parentGroup
                        ? parentGroup.groupItems.add()
                        : doc.groupItems.add();

                    /* 左右方向の矩形計算 / Horizontal rectangle calculation */
                    var rectLeft = left + (partWidth * columnIndex) - overlapInPoints / 2;
                    var rectWidth = partWidth + overlapInPoints;

                    if (rectLeft < left) {
                        rectWidth -= (left - rectLeft);
                        rectLeft = left;
                    }
                    if (rectLeft + rectWidth > left + width) {
                        rectWidth = (left + width) - rectLeft;
                    }

                    /* 上下方向の矩形計算 / Vertical rectangle calculation */
                    var rectTop = top - (partHeight * rowIndex) + overlapInPoints / 2;
                    var rectHeight = partHeight + overlapInPoints;

                    if (rectTop > top) {
                        rectHeight -= (rectTop - top);
                        rectTop = top;
                    }
                    if (rectTop - rectHeight < top - height) {
                        rectHeight = rectTop - (top - height);
                    }

                    var clippingRect = clippingGroup.pathItems.rectangle(
                        rectTop,
                        rectLeft,
                        rectWidth,
                        rectHeight
                    );
                    clippingRect.stroked = false;
                    clippingRect.filled = false;

                    selectedItem.duplicate(clippingGroup, ElementPlacement.PLACEATEND);

                    clippingRect.zOrder(ZOrderMethod.BRINGTOFRONT);
                    clippingGroup.clipped = true;

                    /* ケイ設定の適用 / Apply rule (stroke + pathfinder) */
                    if (shouldAddStroke) {
                        doc.selection = null;
                        clippingGroup.selected = true;
                        app.executeMenuCommand('Adobe New Stroke Shortcut');
                        app.executeMenuCommand('Live Pathfinder Add');
                    }

                    /* 角丸ライブエフェクトの適用 / Apply round corners live effect */
                    if (shouldApplyRoundCorners && roundRadiusInPoints > 0) {
                        var roundXML = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + roundRadiusInPoints + ' "/></LiveEffect>';
                        clippingGroup.applyEffect(roundXML);
                    }

                    if (!parentGroup) {
                        createdItems.push(clippingGroup);
                    }
                }
            }

            if (parentGroup) {
                createdItems.push(parentGroup);
            }

            selectedItem.remove();
            processedCount++;

        } catch (e) {
            skippedCount++;
        }
    }

    /* 実行後に生成アイテムを選択し直し / Reselect created items after execution */
    doc.selection = null;
    for (var createdIndex = 0; createdIndex < createdItems.length; createdIndex++) {
        createdItems[createdIndex].selected = true;
    }

    /* スキップがある場合のみ通知 / Notify only when some items were skipped */
    if (skippedCount > 0) {
        if (uiLang === 'ja') {
            alert(
                labelText('processed') + processedCount + getLabel('unit') + "\n" +
                labelText('skipped') + skippedCount + getLabel('unit')
            );
        } else {
            alert(
                getLabel('processed') + ': ' + processedCount + getLabel('unit') + "\n" +
                getLabel('skipped') + ': ' + skippedCount + getLabel('unit')
            );
        }
    }
})();