#target illustrator
#targetengine "ArcTextGeneratorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

円弧に沿ったパス上文字を生成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArcTextGenerator.md

note記事も参照してください。
https://note.com/gautt/n/n92f6faeda048

### Overview

Generates text on an arc-shaped path.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArcTextGenerator.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArcTextGenerator";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArcTextGenerator.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArcTextGenerator.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/gautt/n/n92f6faeda048"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
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
        dialogTitle: { ja: "アーチ文字に変換", en: "Convert to Arc Text" },
        arcRoundness: { ja: "カーブ", en: "Curve" },
        arcDirection: { ja: "向き", en: "Direction" },
        arcDirectionUp: { ja: "上向き", en: "Up" },
        arcDirectionDown: { ja: "下向き", en: "Down" },
        fit: { ja: "パス幅に合わせる", en: "Fit to path width" },
        fitNone: { ja: "しない", en: "None" },
        fitByFontSize: { ja: "文字サイズ", en: "Font size" },
        fitByTracking: { ja: "トラッキング", en: "Tracking" },
        effect: { ja: "効果", en: "Effect" },
        effectRainbow: { ja: "虹", en: "Rainbow" },
        effectDistort: { ja: "歪み", en: "Skew" },
        effectRibbon: { ja: "3D リボン", en: "3D Ribbon" },
        effectStep: { ja: "階段", en: "Stair Step" },
        effectGravity: { ja: "引力", en: "Gravity" },
        tracking: { ja: "トラッキング", en: "Tracking" },
        preview: { ja: "プレビュー", en: "Preview" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        alertNoDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
        alertNoText: { ja: "対象のテキストが見つかりません", en: "No target text found." },
        alertArcFail: { ja: "アーチ状のパス生成に失敗しました", en: "Failed to generate an arc path." },
        tipRoundness: { ja: "0で直線に近く、100で最も丸くなります。", en: "0 is almost straight; 100 gives the roundest curve." },
        tipDirection: { ja: "アーチの膨らむ方向を指定します。", en: "Sets the direction in which the arc bulges." },
        tipFit: { ja: "パスの端まで文字を収める方法を選びます。", en: "Chooses how the text is fitted to the path endpoints." },
        tipFitNone: { ja: "パス幅には合わせず、アーチ変換のみ行います。", en: "Does not fit to the path width; only converts to arc text." },
        tipFitMethod: { ja: "文字サイズを変えて合わせるか、文字サイズを保ったままトラッキングで合わせるかを選びます。", en: "Choose whether to fit by changing the font size, or by keeping the font size and adjusting tracking." },
        tipEffect: { ja: "Illustratorの「パス上文字オプション」の効果を適用します。", en: "Applies Illustrator's Type on a Path effect." },
        tipTracking: { ja: "既存のトラッキング値に加算します。", en: "Adds this value to the existing tracking." },
        tipTrackingToggle: { ja: "ONでトラッキングを調整できます。OFFにすると0に戻ります。", en: "Enable to adjust tracking. Turning it off resets it to 0." },
        tipPreview: { ja: "ONの間は仮の結果を表示します。OFFまたはキャンセルで元に戻ります。", en: "Shows a temporary result while enabled. Turning it off or cancelling restores the original." },
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        }
    };

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

    /* ===== ユーティリティ / Utilities ===== */

    // Parse a number from a string; return fallback when not numeric
    function parseNumber(str, fallback) {
        var n = Number(str);
        if (isNaN(n)) return fallback;
        return n;
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

    /* 対象のテキスト：ポイント文字・パス上文字をグループの中まで（文字の選択は対象外）
       Target text: point and path text, including inside groups (text selections are not used) */
    var TARGET_TEXT_OPTIONS = { kinds: ["point", "path"], textRangeToFrame: false, unique: false };

    /* 一緒に選択したパス：パスと複合パス（複合パスはまとめて1つ）をグループの中まで
       Paths selected alongside: paths and compound paths (as a whole), including inside groups */
    var TARGET_PATH_OPTIONS = { compoundPaths: "whole", textRangeToFrame: false, unique: false };

    /* ===== 選択の取得 / Selection ===== */
    if (app.documents.length === 0) {
        alert(getLabel('alertNoDoc'));
        return;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;

    // Base selection snapshot (used for stable preview while dialog is open)
    var baseSelection = [];
    try { baseSelection = sel.slice(0); } catch (e) { baseSelection = []; }

    var targetItems = collectSelectionTextFrames(sel, TARGET_TEXT_OPTIONS);
    var selectedPaths = collectSelectionPathItems(sel, TARGET_PATH_OPTIONS);

    if (targetItems.length === 0) {
        alert(getLabel('alertNoText'));
        return;
    }

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

    /* ===== ダイアログ / Dialog ===== */
    var dlg = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = [15, 20, 15, 15];

    // Common width for the leading label of each row (keeps controls aligned)
    var LABEL_COLUMN_WIDTH = 120;

    /* まるみ / Roundness */
    var grpRoundness = dlg.add('group');
    grpRoundness.orientation = 'row';
    grpRoundness.alignChildren = ['left', 'center'];
    // grpRoundness.margins = [0, 5, 0, 10];

    var stArcRoundness = grpRoundness.add('statictext', undefined, labelText('arcRoundness'));
    stArcRoundness.preferredSize.width = LABEL_COLUMN_WIDTH;
    stArcRoundness.helpTip = getLabel('tipRoundness');
    // Slider: 0 = flat, 50 = default arch, 100 = roundest
    var slArcRoundness = grpRoundness.add('slider', undefined, 50, 0, 100);
    slArcRoundness.preferredSize.width = 200;
    slArcRoundness.helpTip = getLabel('tipRoundness');

    /* アーチ方向 / Arc direction */
    var grpArcDirection = dlg.add('group');
    grpArcDirection.orientation = 'row';
    grpArcDirection.alignChildren = ['left', 'center'];
    // grpArcDirection.margins = [0, 0, 0, 10];

    var stArcDirection = grpArcDirection.add('statictext', undefined, labelText('arcDirection'));
    stArcDirection.preferredSize.width = LABEL_COLUMN_WIDTH;
    stArcDirection.helpTip = getLabel('tipDirection');
    var rbArcDirectionUp = grpArcDirection.add('radiobutton', undefined, getLabel('arcDirectionUp'));
    rbArcDirectionUp.helpTip = getLabel('tipDirection');
    var rbArcDirectionDown = grpArcDirection.add('radiobutton', undefined, getLabel('arcDirectionDown'));
    rbArcDirectionDown.helpTip = getLabel('tipDirection');
    rbArcDirectionUp.value = true;

    /* フィット / Fit */
    var grpFit = dlg.add('group');
    grpFit.orientation = 'row';
    grpFit.alignChildren = ['left', 'center'];

    // 先頭ラベル「パス幅に合わせる：」 / Leading label "Fit to path width:"
    var stFit = grpFit.add('statictext', undefined, labelText('fit'));
    stFit.preferredSize.width = LABEL_COLUMN_WIDTH;
    stFit.helpTip = getLabel('tipFit');
    // フィット方法：しない（初期値）／文字サイズ＝サイズ変更／トラッキング＝サイズ維持で字間調整
    var rbFitNone = grpFit.add('radiobutton', undefined, getLabel('fitNone'));
    rbFitNone.helpTip = getLabel('tipFitNone');
    var rbFitFontSize = grpFit.add('radiobutton', undefined, getLabel('fitByFontSize'));
    rbFitFontSize.helpTip = getLabel('tipFitMethod');
    var rbFitTracking = grpFit.add('radiobutton', undefined, getLabel('fitByTracking'));
    rbFitTracking.helpTip = getLabel('tipFitMethod');
    // 既定は従来どおりパス幅に合わせない / Default = no fit (legacy default)
    rbFitNone.value = true;

    /* 効果 / Effect */
    var grpEffect = dlg.add('group');
    grpEffect.orientation = 'row';
    grpEffect.alignChildren = ['left', 'center'];
    // grpEffect.margins = [0, 0, 0, 10];

    var stEffect = grpEffect.add('statictext', undefined, labelText('effect'));
    stEffect.preferredSize.width = LABEL_COLUMN_WIDTH;
    stEffect.helpTip = getLabel('tipEffect');
    var rbEffectRainbow = grpEffect.add('radiobutton', undefined, getLabel('effectRainbow'));
    rbEffectRainbow.helpTip = getLabel('tipEffect');
    var rbEffectDistort = grpEffect.add('radiobutton', undefined, getLabel('effectDistort'));
    rbEffectDistort.helpTip = getLabel('tipEffect');
    var rbEffectRibbon = grpEffect.add('radiobutton', undefined, getLabel('effectRibbon'));
    rbEffectRibbon.helpTip = getLabel('tipEffect');
    var rbEffectStep = grpEffect.add('radiobutton', undefined, getLabel('effectStep'));
    rbEffectStep.helpTip = getLabel('tipEffect');
    var rbEffectGravity = grpEffect.add('radiobutton', undefined, getLabel('effectGravity'));
    rbEffectGravity.helpTip = getLabel('tipEffect');
    // 既定はパス上文字の標準スタイルと同じ「虹」 / Default = Rainbow (Illustrator's own default)
    rbEffectRainbow.value = true;

    /* トラッキング / Tracking */
    var grpTracking = dlg.add('group');
    grpTracking.orientation = 'row';
    grpTracking.alignChildren = ['left', 'center'];
    // grpTracking.margins = [0, 0, 0, 10];

    var stTracking = grpTracking.add('statictext', undefined, labelText('tracking'));
    stTracking.preferredSize.width = LABEL_COLUMN_WIDTH;
    stTracking.helpTip = getLabel('tipTracking');
    // チェックOFFでトラッキング加算を無効化（値は0に固定）/ Checkbox OFF disables tracking (forced to 0)
    var cbTracking = grpTracking.add('checkbox', undefined, '');
    cbTracking.helpTip = getLabel('tipTrackingToggle');
    cbTracking.value = true;
    /* ∧∨と入力欄は隙間0で突き合わせる。∧∨・↑↓キーで増減したらスライダーとプレビューを追従させる
       Butt the stepper against the field; stepping syncs the slider and the preview */
    var grpTrackingStepper = grpTracking.add('group');
    grpTrackingStepper.orientation = 'row';
    grpTrackingStepper.alignChildren = ['left', 'center'];
    grpTrackingStepper.spacing = 0;
    grpTrackingStepper.margins = 0;
    var etTracking;
    var trackingStepper = addStepper(grpTrackingStepper, function () { return etTracking; }, {
        min: -100, max: 500, integer: true,
        onStep: function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); }
    });
    etTracking = grpTrackingStepper.add('edittext', undefined, '0');
    etTracking.characters = 6;
    etTracking.helpTip = getLabel('tipTracking');
    var slTracking = grpTracking.add('slider', undefined, 0, -100, 500);
    slTracking.preferredSize.width = 150;
    slTracking.helpTip = getLabel('tipTracking');
    // 矢印キーも∧∨と同じ処理で増減 / Arrow keys share the stepper's logic
    bindSteppedArrowKeys(etTracking, trackingStepper);

    /* フッター / Footer */
    var buttonRow = addButtonRow(dlg);
    var cbPreview = buttonRow.leftGroup.add('checkbox', undefined, getLabel('preview'));
    cbPreview.helpTip = getLabel('tipPreview');
    cbPreview.value = true;

    var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('cancel'));
    var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('ok'), { name: 'ok' });

    /* ===== プレビュー（Undoなし） / Preview (no undo) ===== */
    var previewTempItems = [];        // items created during preview
    var previewHiddenOriginals = [];  // originals hidden during preview

    function clearPreview() {
        // Remove temp items
        for (var i = previewTempItems.length - 1; i >= 0; i--) {
            try { previewTempItems[i].remove(); } catch (e) { }
        }
        previewTempItems = [];

        // Restore originals visibility
        for (var j = previewHiddenOriginals.length - 1; j >= 0; j--) {
            try { previewHiddenOriginals[j].hidden = false; } catch (e) { }
        }
        previewHiddenOriginals = [];
    }

    function hideOriginalForPreview(item) {
        try {
            if (!item) return;
            for (var k = 0; k < previewHiddenOriginals.length; k++) {
                if (previewHiddenOriginals[k] === item) return;
            }
            item.hidden = true;
            previewHiddenOriginals.push(item);
        } catch (e) { }
    }

    function applyPreview() {
        clearPreview();

        // Restore base selection so preview stays stable even after selection changes
        try { doc.selection = baseSelection; } catch (e) { }

        var currentSelection = [];
        try { currentSelection = doc.selection; } catch (e) { currentSelection = []; }
        if (!currentSelection || currentSelection.length === 0) {
            currentSelection = baseSelection;
        }
        targetItems = collectSelectionTextFrames(currentSelection, TARGET_TEXT_OPTIONS);
        selectedPaths = collectSelectionPathItems(currentSelection, TARGET_PATH_OPTIONS);

        if (!targetItems || targetItems.length === 0) {
            cbPreview.value = false;
            return;
        }

        generateArcText(false, true);
        app.redraw();
    }

    function refreshPreviewIfNeeded() {
        if (cbPreview.value) applyPreview();
    }

    /* ===== ハンドラ / Handlers ===== */
    slArcRoundness.onChange = refreshPreviewIfNeeded;
    rbArcDirectionUp.onClick = refreshPreviewIfNeeded;
    rbArcDirectionDown.onClick = refreshPreviewIfNeeded;
    // 「トラッキング」フィット選択中は手動トラッキング行をディム表示
    function updateTrackingEnabled() {
        var fitByTrackingActive = rbFitTracking.value;
        stTracking.enabled = !fitByTrackingActive;
        cbTracking.enabled = !fitByTrackingActive;
        var manualTrackingActive = cbTracking.value && !fitByTrackingActive;
        etTracking.enabled = manualTrackingActive;
        trackingStepper.enabled = manualTrackingActive;
        redrawSteppersIn(trackingStepper); // ∧∨のディム表示を描き直す / redraw the stepper's dimming
        slTracking.enabled = manualTrackingActive;
    }
    function onFitMethodChanged() {
        updateTrackingEnabled();
        refreshPreviewIfNeeded();
    }
    rbFitNone.onClick = onFitMethodChanged;
    rbFitFontSize.onClick = onFitMethodChanged;
    rbFitTracking.onClick = onFitMethodChanged;
    rbEffectRainbow.onClick = refreshPreviewIfNeeded;
    rbEffectDistort.onClick = refreshPreviewIfNeeded;
    rbEffectRibbon.onClick = refreshPreviewIfNeeded;
    rbEffectStep.onClick = refreshPreviewIfNeeded;
    rbEffectGravity.onClick = refreshPreviewIfNeeded;

    /* トラッキング UI 同期 / Tracking UI sync (edittext <-> slider) */
    var trackingSyncLock = false;

    function syncTrackingFromEdit() {
        if (trackingSyncLock) return;
        trackingSyncLock = true;
        try {
            var v = Math.max(-100, Math.min(500, Math.round(parseNumber(etTracking.text, 0))));
            etTracking.text = String(v);
            slTracking.value = v;
        } catch (e) { }
        trackingSyncLock = false;
    }

    function syncTrackingFromSlider() {
        if (trackingSyncLock) return;
        trackingSyncLock = true;
        try {
            var v = Math.round(slTracking.value);
            etTracking.text = String(v);
        } catch (e) { }
        trackingSyncLock = false;
    }

    etTracking.onChanging = function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); };
    slTracking.onChanging = function () { syncTrackingFromSlider(); };
    slTracking.onChange = function () { syncTrackingFromSlider(); refreshPreviewIfNeeded(); };
    cbTracking.onClick = function () {
        // OFFにしたらトラッキングを0に戻す / Reset tracking to 0 when turned off
        if (!cbTracking.value) {
            etTracking.text = '0';
            syncTrackingFromEdit();
        }
        updateTrackingEnabled();
        refreshPreviewIfNeeded();
    };
    syncTrackingFromEdit();
    updateTrackingEnabled();

    cbPreview.onClick = function () {
        if (cbPreview.value) {
            applyPreview();
        } else {
            clearPreview();
            app.redraw();
        }
    };

    btnCancel.onClick = function () {
        clearPreview();
        dlg.close(0);
    };

    btnOK.onClick = function () {
        // If preview is ON, clear it first so we don't stack temporary objects.
        if (cbPreview.value) clearPreview();
        generateArcText(true, false);
        dlg.close(1);
    };

    /* 起動時に一度プレビュー / Auto-apply preview once on open */
    try {
        if (cbPreview.value) {
            applyPreview();
            app.redraw();
        }
    } catch (e) { }

    alignRightOnlyButtonRow(buttonRow);
    prepareDialogWindow(dlg, SCRIPT_NAME);
    var dialogResult = dlg.show();
    if (dialogResult !== 1) return;

    /* ===== パススタイル / Path style ===== */

    // Run styleFn for each underlying PathItem (handles CompoundPathItem too).
    function forEachPathItem(pathItem, styleFn) {
        try {
            if (!pathItem || !styleFn) return;
            if (pathItem.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < pathItem.pathItems.length; cp++) {
                    try { styleFn(pathItem.pathItems[cp]); } catch (e) { }
                }
            } else {
                try { styleFn(pathItem); } catch (e) { }
            }
        } catch (e) { }
    }

    // Per-PathItem style: invisible (no fill, stroke color/weight 0).
    function styleInvisiblePath(pi) {
        pi.filled = false;
        pi.stroked = false;
        pi.strokeWidth = 0;
    }

    // Generated arc path: stroke color/weight 0 (invisible guide).
    function applyInvisiblePathStyle(pathItem) {
        forEachPathItem(pathItem, styleInvisiblePath);
    }

    /* ===== アーチ生成 / Arc generation ===== */

    // If a path is selected together with text, it should not remain.
    // Preview: hide it (restored by clearPreview). Execute: delete it.
    function handleSelectedPaths(previewMode) {
        try {
            if (!selectedPaths || selectedPaths.length === 0) return;
            for (var i = selectedPaths.length - 1; i >= 0; i--) {
                var path = selectedPaths[i];
                if (!path) continue;
                if (previewMode) {
                    hideOriginalForPreview(path);
                } else {
                    try { path.remove(); } catch (e) { }
                }
            }
        } catch (e) { }
    }

    // Apply center justification to all paragraphs of a text frame
    function applyCenterJustification(textFrame) {
        try {
            if (!textFrame) return;
            try {
                if (textFrame.paragraphs && textFrame.paragraphs.length > 0) {
                    for (var i = 0; i < textFrame.paragraphs.length; i++) {
                        try {
                            textFrame.paragraphs[i].paragraphAttributes.justification = Justification.CENTER;
                        } catch (e) { }
                    }
                }
            } catch (e) { }
            try { textFrame.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
        } catch (e) { }
    }

    /* ===== 効果・トラッキング / Effect & tracking ===== */

    // Resolve the menu command for the selected effect (null = none)
    function getSelectedEffectCommand() {
        if (rbEffectRainbow.value) return 'Rainbow';
        if (rbEffectDistort.value) return 'Skew';
        if (rbEffectRibbon.value) return '3D ribbon';
        if (rbEffectStep.value) return 'Stair Step';
        if (rbEffectGravity.value) return 'Gravity';
        return null;
    }

    // Apply the selected path-text effect via menu command (requires selection)
    function applyPathTextEffect(textFrame) {
        var cmd = getSelectedEffectCommand();
        if (!cmd) return;

        var prevSelection = null;
        try { prevSelection = doc.selection; } catch (e) { prevSelection = null; }

        try {
            try { doc.selection = []; } catch (e) { }
            try { textFrame.selected = true; } catch (e) { }
            app.executeMenuCommand(cmd);
        } catch (e) { }

        // Always restore the previous selection (safer)
        try { doc.selection = prevSelection; } catch (e) { }
    }

    // Collect every textRange of a frame (falls back to its single textRange)
    function collectTextRanges(textFrame) {
        var ranges = [];
        try {
            if (textFrame.textRanges && textFrame.textRanges.length > 0) {
                for (var i = 0; i < textFrame.textRanges.length; i++) ranges.push(textFrame.textRanges[i]);
            }
        } catch (e) { }
        if (ranges.length === 0) {
            try { if (textFrame.textRange) ranges = [textFrame.textRange]; } catch (e) { ranges = []; }
        }
        return ranges;
    }

    // Add the tracking delta to the existing tracking of each textRange
    function applyTrackingDelta(textFrame) {
        try {
            if (!textFrame) return;
            var delta = Math.round(parseNumber(etTracking.text, 0));
            if (!delta) return; // 0 -> do nothing

            var ranges = collectTextRanges(textFrame);
            for (var r = 0; r < ranges.length; r++) {
                var textRange = ranges[r];
                try {
                    var currentTracking = textRange.characterAttributes.tracking;
                    textRange.characterAttributes.tracking = currentTracking + delta;
                } catch (e) { }
            }
        } catch (e) { }
    }

    // Measure rendered text bounds via temporary outlines: [L, T, R, B] or null
    function measureTextBounds(originalText) {
        try {
            // [0] holds only the first line (used for the baseline),
            // [1] holds everything (used for the left/top/right extents).
            var measureTexts = [originalText.duplicate(), originalText.duplicate()];
            measureTexts[0].contents = '';
            for (var i = 0; i < originalText.lines[0].length; i++) {
                originalText.textRanges[i].duplicate(measureTexts[0]);
            }

            for (var k = 0; k < measureTexts.length; k++) {
                measureTexts[k] = measureTexts[k].createOutline();
            }

            var bounds = measureTexts[1].geometricBounds; // [L, T, R, B]
            bounds[3] = measureTexts[0].geometricBounds[3]; // baseline from the first line

            for (var r = 0; r < measureTexts.length; r++) {
                try { measureTexts[r].remove(); } catch (e) { }
            }
            return bounds;
        } catch (e) {
            return null;
        }
    }

    // Roundness from the slider, clamped to 0-100 (50 = default arch)
    function readRoundnessPercent() {
        var percent = 50;
        try { percent = Number(slArcRoundness.value); } catch (e) { percent = 50; }
        if (isNaN(percent)) return 50;
        return Math.max(0, Math.min(100, percent));
    }

    // Bend a 2-point straight path into an arc by moving its Bézier handles.
    // directionSign: +1 = bulge upward, -1 = bulge downward.
    // 水平ハンドルは固定、垂直ハンドル＝カーブの深さ（50% で pathLength/4）。
    function applyArcHandles(arcPath, roundnessPercent, directionSign) {
        var pathLength = arcPath.length;
        var horizontalHandle = pathLength / 3.5;
        var verticalHandle = pathLength * (roundnessPercent / 100) * 0.5 * directionSign;

        var pathPoints = arcPath.pathPoints;
        pathPoints[0].rightDirection = [
            pathPoints[0].rightDirection[0] + horizontalHandle,
            pathPoints[0].rightDirection[1] + verticalHandle
        ];
        pathPoints[1].leftDirection = [
            pathPoints[1].leftDirection[0] - horizontalHandle,
            pathPoints[1].leftDirection[1] + verticalHandle
        ];
    }

    // Create an arc-shaped path sized to the given point text
    function createArcPathFromText(originalText, layer) {
        var baselineYMultiplier = 1.02;

        // Guard: empty / invalid text
        try {
            if (!originalText || originalText.typename !== 'TextFrame') return null;
            if (!originalText.lines || originalText.lines.length === 0) return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;
        } catch (e) {
            return null;
        }

        try {
            var textBounds = measureTextBounds(originalText);
            if (!textBounds) return null;

            // Base straight path along the baseline
            var baselineY = textBounds[3] * baselineYMultiplier;
            var arcPath = layer.pathItems.add();
            arcPath.setEntirePath([
                [textBounds[0], baselineY],
                [textBounds[2], baselineY]
            ]);
            try {
                arcPath.stroked = false;
                arcPath.filled = false;
            } catch (e) { }

            // Bend the straight path into an arc（上＝＋ / 下＝−）
            var directionSign = (rbArcDirectionDown && rbArcDirectionDown.value) ? -1 : 1;
            applyArcHandles(arcPath, readRoundnessPercent(), directionSign);

            return arcPath;
        } catch (e) {
            return null;
        }
    }

    // Main process: generate an arc path and place the text on it
    function generateArcText(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var createdTexts = [];

        // A path selected together with text should not remain in arc mode
        handleSelectedPaths(previewMode);

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentLayer = originalText.layer;

            // Create an arc-like path from the text bounds
            var arcPath = createArcPathFromText(originalText, currentLayer);
            if (!arcPath) {
                if (showAlerts) alert(getLabel('alertArcFail'));
                continue;
            }

            // Generated arc path: stroke color/weight 0 (invisible guide)
            applyInvisiblePathStyle(arcPath);
            if (previewMode) previewTempItems.push(arcPath);

            // Create text on the path
            var textOnAPath = currentLayer.textFrames.pathText(arcPath);
            // Keep stacking position (avoid appearing to disappear behind other objects)
            try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (e) { }
            if (previewMode) previewTempItems.push(textOnAPath);

            // Keep the path used by the PathText invisible (AI may override style on conversion)
            try {
                var pathTextPath = textOnAPath.textPath;
                if (pathTextPath) applyInvisiblePathStyle(pathTextPath);
            } catch (e) { }

            // Duplicate textRanges from the original text frame
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

            // 行揃え：常に中央 ※ duplicate 後に適用しないと上書きされる
            applyCenterJustification(textOnAPath);

            // トラッキング（既存値 + 指定値）※ フィット前に適用してオーバーセット判定へ反映
            // 「トラッキング」フィット時は手動値を加算しない（ディム表示と挙動を一致させる）
            if (!rbFitTracking.value) applyTrackingDelta(textOnAPath);

            // 効果（パス上文字の効果をメニューコマンドで適用）
            applyPathTextEffect(textOnAPath);

            // Collect created PathText for fit (even in preview)
            createdTexts.push(textOnAPath);

            // Remove or hide the original text frame
            if (previewMode) {
                hideOriginalForPreview(originalText);
            } else {
                originalText.remove();
            }

            // Select the created text on a path
            if (!previewMode) {
                try { textOnAPath.selected = true; } catch (e) { }
            }
        }

        // フィット：「しない」以外を選んだとき（ループ後にまとめて適用）
        if (rbFitTracking.value) {
            // 文字サイズを保ったまま、トラッキングでパス幅に合わせる
            try { fitTextToOpenPathByTracking(createdTexts); } catch (e) { }
        } else if (rbFitFontSize.value) {
            // 文字サイズを変更してパス幅に合わせる（従来）
            try { fitTextToOpenPath(createdTexts); } catch (e) { }
        }
    }

    /* ===== フィット / Fit ===== */

    // True if the path (or first sub-path of a compound path) is closed
    function isClosedPathItem(item) {
        try {
            if (!item) return false;
            if (item.typename === 'CompoundPathItem') {
                if (item.pathItems && item.pathItems.length > 0) return !!item.pathItems[0].closed;
                return false;
            }
            if (item.typename === 'PathItem') return !!item.closed;
        } catch (e) { }
        return false;
    }

    // True if the frame is editable PathText placed on an OPEN path (a fit target)
    function isTargetPathText(textFrame) {
        try {
            if (!textFrame || textFrame.typename !== 'TextFrame') return false;
            if (textFrame.kind !== TextType.PATHTEXT) return false;
            if (!textFrame.editable || textFrame.locked || textFrame.hidden) return false;
            var path = textFrame.textPath;
            if (!path) return false;
            // Only OPEN paths
            if (isClosedPathItem(path)) return false;
            return true;
        } catch (e) { }
        return false;
    }

    // Overset detection: some characters are pushed past the visible lines
    function isOverset(textFrame, lineAmount) {
        try {
            if (!textFrame) return false;

            if (textFrame.lines.length > 0) {
                var charactersOnVisibleLines = 0;

                if (typeof (lineAmount) === 'undefined' || lineAmount === null) {
                    lineAmount = 1;
                } else {
                    lineAmount = Math.floor(lineAmount);
                    if (lineAmount < 1) lineAmount = 1;
                    if (lineAmount > textFrame.lines.length) lineAmount = textFrame.lines.length;
                }

                for (var i = 0; i < lineAmount; i++) {
                    charactersOnVisibleLines += textFrame.lines[i].characters.length;
                }
                return (charactersOnVisibleLines < textFrame.characters.length);
            } else if (textFrame.characters.length > 0) {
                return true;
            }
        } catch (e) { }
        return false;
    }

    // Visible line count of a frame (always >= 1)
    function getLineAmount(textFrame) {
        try {
            if (textFrame.lines && textFrame.lines.length > 0) return textFrame.lines.length;
        } catch (e) { }
        return 1;
    }

    // Fit PathText to OPEN path endpoints by adjusting font size only.
    // - If overset: shrink by small steps until it fits.
    // - If NOT overset: grow to intentionally create overset, then shrink to fit.
    function fitTextToOpenPath(frames) {
        if (!frames || frames.length === 0) return false;

        var opt = {
            increment: 0.1,
            minFontSize: 0.1,
            maxShrinkIter: 2000,
            maxGrowIter: 10,
            maxFontSize: 2000
        };

        function shrinkFont(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);

                // If it is NOT overset, grow first (doubling) until it becomes overset
                if (!isOverset(textFrame, lineAmount)) {
                    var growIter = 0;
                    while (!isOverset(textFrame, lineAmount) && growIter < opt.maxGrowIter) {
                        var growSize = textFrame.textRange.characterAttributes.size;
                        if (growSize >= opt.maxFontSize) break;
                        textFrame.textRange.characterAttributes.size = Math.min(opt.maxFontSize, growSize * 2);
                        growIter++;
                    }
                }

                // Then shrink in small steps until it fits
                var shrinkIter = 0;
                while (isOverset(textFrame, lineAmount)) {
                    var currentSize = textFrame.textRange.characterAttributes.size;
                    if (currentSize <= opt.minFontSize) break;

                    textFrame.textRange.characterAttributes.size = Math.max(opt.minFontSize, currentSize - opt.increment);

                    shrinkIter++;
                    if (shrinkIter >= opt.maxShrinkIter) break;
                }
            } catch (e) { }
        }

        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            if (!isTargetPathText(textFrame)) continue;
            shrinkFont(textFrame);
        }

        return true;
    }

    // Fit PathText to OPEN path endpoints by adjusting tracking only (font size kept).
    // - A coarse pass drives the text across the overset boundary,
    // - then a fine pass settles on the widest tracking that still fits.
    function fitTextToOpenPathByTracking(frames) {
        if (!frames || frames.length === 0) return false;

        var opt = {
            coarseStep: 50,     // tracking units per coarse step
            fineStep: 1,        // tracking units per fine step
            minTracking: -1000, // tightest allowed cumulative delta
            maxTracking: 20000, // loosest allowed cumulative delta
            maxIter: 4000
        };

        // Add a tracking delta to every textRange of the frame
        function addTrackingToFrame(textFrame, delta) {
            if (!delta) return;
            var ranges = collectTextRanges(textFrame);
            for (var r = 0; r < ranges.length; r++) {
                try {
                    var attributes = ranges[r].characterAttributes;
                    attributes.tracking = attributes.tracking + delta;
                } catch (e) { }
            }
        }

        function fitByTracking(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);
                var applied = 0; // cumulative tracking delta applied so far
                var iterations;

                if (isOverset(textFrame, lineAmount)) {
                    // Too wide: tighten (coarse) until it fits
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied - opt.coarseStep < opt.minTracking) break;
                        addTrackingToFrame(textFrame, -opt.coarseStep);
                        applied -= opt.coarseStep;
                        iterations++;
                    }
                    // Loosen back (fine) until it overflows again
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied + opt.fineStep > opt.maxTracking) break;
                        addTrackingToFrame(textFrame, opt.fineStep);
                        applied += opt.fineStep;
                        iterations++;
                    }
                    // Stepped one fineStep too far: pull back once so it fits
                    if (isOverset(textFrame, lineAmount)) {
                        addTrackingToFrame(textFrame, -opt.fineStep);
                        applied -= opt.fineStep;
                    }
                } else {
                    // Fits with room: loosen (coarse) until it overflows
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied + opt.coarseStep > opt.maxTracking) break;
                        addTrackingToFrame(textFrame, opt.coarseStep);
                        applied += opt.coarseStep;
                        iterations++;
                    }
                    // Tighten back (fine) until it fits
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied - opt.fineStep < opt.minTracking) break;
                        addTrackingToFrame(textFrame, -opt.fineStep);
                        applied -= opt.fineStep;
                        iterations++;
                    }
                }
            } catch (e) { }
        }

        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            if (!isTargetPathText(textFrame)) continue;
            fitByTracking(textFrame);
        }

        return true;
    }
}());
