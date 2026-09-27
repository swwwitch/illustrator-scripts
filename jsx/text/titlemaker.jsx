#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のテキストフレームを、指定した文字（、。〜など）の直後で改行します。
改行の対象にする文字はチェックボックスで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/titlemaker.md

### Overview

Breaks the selected text frame onto a new line right after the characters you choose, such as 、 。 or 〜.
Which characters trigger a break is set with checkboxes.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/titlemaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "titlemaker";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/titlemaker.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/titlemaker.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 改行対象文字と、開いたときのチェック状態 / Line-break characters and whether each starts checked */
    var TARGET_CHARS = [
        { mark: "、", defaultOn: true },
        { mark: "。", defaultOn: true },
        { mark: "〜", defaultOn: false }
    ];

    /* サイズ調整のデフォルト倍率（%）/ Default scale for size adjustment (%) */
    var DEFAULT_SIZE_PERCENT = 80;

    // =========================================
    // レイアウト / Layout
    // =========================================
    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;
    var PANEL_ITEM_SPACING = 6;   /* チェックボックスが並ぶパネルの間隔 / spacing inside the checkbox panels */
    var SIZE_ROW_SPACING = 4;     /* サイズ入力行の間隔 / spacing of the size row */
    var SIZE_FIELD_CHARS = 4;     /* サイズ欄の文字数 / size field width */

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は PANEL_SPACING）
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
     * グループに共通のレイアウト設定を適用する（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // 文字種の判定 / Character Matching
    // =========================================
    /* 格助詞本体の正規表現ソース（長い候補を先に）/ Regex source for the case particle itself (longer alternatives first) */
    var CASE_PARTICLE_SOURCE = "(?:から|より|が|を|に|へ|と|で|の)";

    /* 助詞の直前に許される文字クラス（漢字・ひらがな・カタカナ・英数字）。ExtendScript は lookbehind 非対応のため手動判定に使う / Allowed preceding-char class for a particle (kanji / hiragana / katakana / alphanumeric); used manually because ExtendScript lacks lookbehind */
    var PARTICLE_PRECURSOR_SOURCE = "[一-龯ぁ-んァ-ヶA-Za-z0-9]";

    /**
     * ひらがなかどうか判定する（U+3041〜U+309F）
     * @param {string} oneChar - 判定する1文字
     * @returns {boolean} ひらがななら true
     */
    function isHiragana(oneChar) {
        var charCode = oneChar.charCodeAt(0);
        return charCode >= 0x3041 && charCode <= 0x309F;
    }

    /**
     * 格助詞として縮小する文字のインデックスを集める（直前が漢字・かな・英数字のものだけ）
     * @param {string} sourceText - 対象の文字列
     * @returns {Object} インデックスをキーにした集合（値は true）
     */
    function findCaseParticleIndices(sourceText) {
        var markedIndices = {};
        var particlePattern = new RegExp(CASE_PARTICLE_SOURCE, "g");
        var precursorPattern = new RegExp(PARTICLE_PRECURSOR_SOURCE);
        var particleMatch;

        while ((particleMatch = particlePattern.exec(sourceText)) !== null) {
            var matchStart = particleMatch.index;
            /* lookbehind の代わりに直前の1文字を判定 / Emulate lookbehind by testing the single preceding char */
            var prevChar = (matchStart > 0) ? sourceText.charAt(matchStart - 1) : "";
            if (prevChar !== "" && precursorPattern.test(prevChar)) {
                for (var k = 0; k < particleMatch[0].length; k++) markedIndices[matchStart + k] = true;
            }
            /* 空マッチによる無限ループを防ぐ / Guard against infinite loops on zero-length matches */
            if (particleMatch.index === particlePattern.lastIndex) particlePattern.lastIndex++;
        }
        return markedIndices;
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

    /**
     * UI の表示言語を判定する（ja で始まれば日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "タイトルメーカー", en: "Title Maker" }
        },
        panel: {
            targets: { ja: "改行対象文字", en: "Target Characters" },
            sizeAdjust: { ja: "フォントサイズのサイズ調整", en: "Font Size Adjustment" }
        },
        checkbox: {
            caseParticle: { ja: "格助詞", en: "Case particles" },
            hiragana: { ja: "ひらがな", en: "Hiragana" }
        },
        fieldLabel: {
            size: { ja: "サイズ", en: "Size" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            targetChar: { ja: "この文字の後ろで改行します。", en: "Inserts a line break after this character." },
            caseParticle: {
                ja: "「が」「を」「に」などの格助詞を小さくします。",
                en: "Shrinks case particles such as \u0022\u304c\u0022, \u0022\u3092\u0022, and \u0022\u306b\u0022."
            },
            hiragana: { ja: "ひらがなをまとめて小さくします。", en: "Shrinks all hiragana." },
            size: {
                ja: "小さくするときの大きさです。元のフォントサイズに対する割合で指定します。",
                en: "Size applied when shrinking, as a percentage of the original font size."
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
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            noTarget: {
                ja: "改行対象文字またはサイズ調整の対象を1つ以上選択してください。",
                en: "Please select at least one line-break character or size-adjustment target."
            },
            invalidSize: {
                ja: "サイズには正の数値を入力してください。",
                en: "Please enter a positive number for the size."
            }
        }
    };

    /**
     * ドット区切りのキーで表示言語の文言を返す（ステップボタンからは { ja, en } のオブジェクトでも呼ばれる）
     * @param {string|Object} keyPath - "dialog.title" のようなキー、または { ja, en }
     * @returns {string} 表示言語の文言
     */
    function getLabel(keyPath) {
        if (keyPath && typeof keyPath === "object") return keyPath[uiLang] || keyPath.ja;
        var keyParts = keyPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            labelNode = labelNode[keyParts[i]];
        }
        return labelNode[uiLang];
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} keyPath - ラベルのキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(keyPath) {
        return getLabel(keyPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // テキストの加工 / Text processing
    // =========================================

    /**
     * 対象文字の後ろで改行し、欧文ベースラインを適用する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string[]} breakMarks - 改行する文字
     * @returns {void}
     */
    function insertLineBreaksAfterMarks(textFrame, breakMarks) {
        var frameContents = textFrame.contents;

        /* 各対象文字の後ろに改行を挿入 / Insert a line break after each target character */
        for (var i = 0; i < breakMarks.length; i++) {
            var breakMark = breakMarks[i];
            frameContents = frameContents.split(breakMark).join(breakMark + "\r");
        }

        /* 連続改行を整理 / Collapse consecutive line breaks */
        frameContents = frameContents.replace(/\r\r+/g, "\r");

        textFrame.contents = frameContents;

        /* 文字揃え：欧文ベースライン / Character alignment: Roman baseline */
        textFrame.textRange.characterAttributes.baselinePosition =
            FontBaselineOption.NORMALBASELINE;
    }

    /**
     * 格助詞・ひらがなのフォントサイズを指定の割合に縮小する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {{adjustCaseParticle: boolean, adjustHiragana: boolean, sizePercent: number}} sizeOptions - サイズ調整の設定
     * @returns {void}
     */
    function adjustCharacterSizes(textFrame, sizeOptions) {
        var sizeScale = sizeOptions.sizePercent / 100;
        var frameText = textFrame.contents;
        var frameCharacters = textFrame.textRange.characters;

        /* 縮小対象の文字のインデックス（文字列のインデックスと characters は 1:1 対応）/ Indices to scale (string index maps 1:1 to characters) */
        var markedIndices = sizeOptions.adjustCaseParticle ? findCaseParticleIndices(frameText) : {};

        /* ひらがなは1文字ずつ判定して追加 / Add hiragana per character */
        if (sizeOptions.adjustHiragana) {
            for (var i = 0; i < frameText.length; i++) {
                if (isHiragana(frameText.charAt(i))) markedIndices[i] = true;
            }
        }

        /* マーク済みの文字を縮小 / Scale the marked characters */
        for (var j = 0; j < frameCharacters.length; j++) {
            if (markedIndices[j]) {
                var charAttributes = frameCharacters[j].characterAttributes;
                charAttributes.size = charAttributes.size * sizeScale;
            }
        }
    }

    /**
     * 選択中の各テキストフレームに改行挿入とサイズ調整を適用する
     * @param {string[]} breakMarks - 改行する文字（空なら改行しない）
     * @param {Object|null} sizeOptions - サイズ調整の設定（null なら調整しない）
     * @returns {void}
     */
    function processSelection(breakMarks, sizeOptions) {
        var docSelection = app.activeDocument.selection;
        for (var i = 0; i < docSelection.length; i++) {
            var selectedItem = docSelection[i];

            if (selectedItem.typename !== "TextFrame") continue;

            if (breakMarks.length > 0) insertLineBreaksAfterMarks(selectedItem, breakMarks);
            if (sizeOptions) adjustCharacterSizes(selectedItem, sizeOptions);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる（イベントの配線は main() で行う）
     * @returns {Object} 作成したダイアログとコントロール
     */
    function buildTitleMakerDialog() {
        /* タイトルバーにバージョンを表示 / Version shown in the title bar */
        var lineBreakDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        lineBreakDialog.orientation = "column";
        lineBreakDialog.alignChildren = "fill";

        /* 改行対象文字のチェックボックスパネル / Checkbox panel for target characters */
        var targetPanel = lineBreakDialog.add("panel", undefined, getLabel("panel.targets"));
        setupPanel(targetPanel, PANEL_ITEM_SPACING);

        var targetCheckboxes = [];
        for (var i = 0; i < TARGET_CHARS.length; i++) {
            var targetCheckbox = targetPanel.add("checkbox", undefined, TARGET_CHARS[i].mark);
            targetCheckbox.helpTip = getLabel("tooltip.targetChar");
            targetCheckbox.value = TARGET_CHARS[i].defaultOn;
            targetCheckboxes.push(targetCheckbox);
        }

        /* フォントサイズ調整のパネル / Panel for font size adjustment */
        var sizePanel = lineBreakDialog.add("panel", undefined, getLabel("panel.sizeAdjust"));
        setupPanel(sizePanel, PANEL_ITEM_SPACING);

        var caseParticleCheckbox = sizePanel.add("checkbox", undefined, getLabel("checkbox.caseParticle"));
        caseParticleCheckbox.helpTip = getLabel("tooltip.caseParticle");
        var hiraganaCheckbox = sizePanel.add("checkbox", undefined, getLabel("checkbox.hiragana"));
        hiraganaCheckbox.helpTip = getLabel("tooltip.hiragana");

        /* サイズ：［　］% の入力行 / Size: [ ] % input row */
        var sizeRow = sizePanel.add("group");
        setupGroup(sizeRow, "row", SIZE_ROW_SPACING);
        sizeRow.add("statictext", undefined, labelText("fieldLabel.size"));
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var sizeStepperInputGroup = sizeRow.add("group");
        sizeStepperInputGroup.orientation = "row";
        sizeStepperInputGroup.alignChildren = ["left", "center"];
        sizeStepperInputGroup.spacing = 0;
        sizeStepperInputGroup.margins = 0;
        var sizeField;
        /* 正の値だけ受け付ける（OK 時の検証と同じ）ので下限は1 / positive values only, as validated on OK */
        var sizeStepperGroup = addStepper(sizeStepperInputGroup, function () { return sizeField; }, { min: 1 });
        sizeField = sizeStepperInputGroup.add("edittext", undefined, String(DEFAULT_SIZE_PERCENT));
        sizeField.characters = SIZE_FIELD_CHARS;
        sizeField.helpTip = getLabel("tooltip.size");
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(sizeField, sizeStepperGroup);
        sizeRow.add("statictext", undefined, "%");

        /* ボタン（Mac 規約: キャンセル → OK、OK は右）/ Buttons (Mac convention: Cancel → OK, OK on the right) */
        var btnRowGroup = lineBreakDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));

        return {
            lineBreakDialog: lineBreakDialog,
            targetCheckboxes: targetCheckboxes,
            caseParticleCheckbox: caseParticleCheckbox,
            hiraganaCheckbox: hiraganaCheckbox,
            sizeField: sizeField,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを表示し、OK で改行とサイズ調整を行う
     * @returns {void}
     */
    function main() {
        /* 事前チェック / Pre-flight checks */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        if (app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var dialogControls = buildTitleMakerDialog();
        var lineBreakDialog = dialogControls.lineBreakDialog;

        dialogControls.btnCancel.onClick = function () {
            lineBreakDialog.close();
        };

        dialogControls.btnOK.onClick = function () {
            /* チェックされた改行対象文字を収集 / Collect the checked line-break characters */
            var selectedMarks = [];
            for (var i = 0; i < dialogControls.targetCheckboxes.length; i++) {
                if (dialogControls.targetCheckboxes[i].value) selectedMarks.push(TARGET_CHARS[i].mark);
            }

            /* サイズ調整の設定を収集 / Collect the size-adjustment settings */
            var sizeOptions = {
                adjustCaseParticle: dialogControls.caseParticleCheckbox.value,
                adjustHiragana: dialogControls.hiraganaCheckbox.value,
                sizePercent: parseFloat(dialogControls.sizeField.text)
            };
            var wantsSizeAdjust = sizeOptions.adjustCaseParticle || sizeOptions.adjustHiragana;

            if (selectedMarks.length === 0 && !wantsSizeAdjust) {
                alert(getLabel("alert.noTarget"));
                return;
            }
            if (wantsSizeAdjust && (isNaN(sizeOptions.sizePercent) || sizeOptions.sizePercent <= 0)) {
                alert(getLabel("alert.invalidSize"));
                return;
            }

            processSelection(selectedMarks, wantsSizeAdjust ? sizeOptions : null);
            lineBreakDialog.close();
        };

        lineBreakDialog.center();
        lineBreakDialog.show();
    }

    main();

})();
