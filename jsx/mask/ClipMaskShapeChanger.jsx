#target illustrator
#targetengine "ClipMaskShapeChangerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した画像（配置／埋め込み）または既存のクリップグループを、正方形・正円・六角形のクリップ形状へ置き換えます。
ケイ線の追加、角丸、複数オブジェクトの大きさ揃えにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipMaskShapeChanger.md

### Overview

Replaces the clipping shape of the selected images, or of an existing clipping group, with a square, circle or hexagon.
Adding a stroke, rounding corners, and matching sizes across several objects are supported too.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipMaskShapeChanger.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipMaskShapeChanger";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipMaskShapeChanger.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipMaskShapeChanger.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 形状モード / Shape mode */
    // 'square' | 'circle' | 'hexA' | 'hexB' | 'octagon'
    var __shapeMode = 'square'; // 'square' | 'circle' | 'hexA' | 'hexB'
    /* アピアランス設定 / Appearance settings */
    var __appearanceMode = 'strokeOnly'; // 'strokeOnly' | 'clipGroup'
    var __roundCornersEnabled = false;
    var __roundCornerRadius = 0;
    var __sameSizeEnabled = false; // 「大きさを揃える」
    var __sameSizeMode = 'max'; // 'max' | 'min'  （最大/最小）
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

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 }, /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 }, /* 歯 / H */
        "p": { unit: "pc", amount: 1 }     /* パイカ / pica */
    };

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
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
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
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
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
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
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) return;
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
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
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
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
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
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

    /*
    正六角形パスを作成 / Create a regular hexagon path
    - centerX/centerY を中心に、半径 r の外接円上に6点を配置
    - rotationDeg で回転（0=フラットトップ寄り、30=ポイントトップ寄り）
    */
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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "クリップグループの形状変更", en: "Change Clipping Group Shape" }
        },
        panel: {
            shape:      { ja: "形状", en: "Shape" },
            appearance: { ja: "アピアランス", en: "Appearance" },
            option:     { ja: "複数オブジェクト", en: "Multiple Objects" }
        },
        radio: {
            noChange:    { ja: "変更なし", en: "No change" },
            square:      { ja: "正方形", en: "Square" },
            circle:      { ja: "正円", en: "Circle" },
            hexA:        { ja: "六角形A", en: "Hexagon A" },
            hexB:        { ja: "六角形B", en: "Hexagon B" },
            octagon:     { ja: "八角形", en: "Octagon" },
            sameSizeMax: { ja: "最大", en: "Largest" },
            sameSizeMin: { ja: "最小", en: "Smallest" }
        },
        checkbox: {
            sameSize:      { ja: "大きさを揃える", en: "Match sizes" },
            addStrokeOnly: { ja: "ケイ線を追加", en: "Add stroke" },
            roundCorners:  { ja: "角丸", en: "Rounded corners" }
        },
        tooltip: {
            noChange:      { ja: "マスクの形はそのままにして、アピアランスや大きさだけを変えます。", en: "Keeps the mask shape and only changes the appearance or the size." },
            square:        { ja: "マスクを正方形にします。", en: "Makes the mask a square." },
            circle:        { ja: "マスクを正円にします。", en: "Makes the mask a circle." },
            hexA:          { ja: "マスクを六角形（頂点が上下）にします。", en: "Makes the mask a hexagon with points at the top and bottom." },
            hexB:          { ja: "マスクを六角形（辺が上下）にします。", en: "Makes the mask a hexagon with flat top and bottom." },
            octagon:       { ja: "マスクを八角形にします。", en: "Makes the mask an octagon." },
            sameSize:      { ja: "複数のクリップグループの大きさをそろえます。", en: "Gives every clipping group the same size." },
            sameSizeMax:   { ja: "いちばん大きいものに合わせます。", en: "Matches the largest one." },
            sameSizeMin:   { ja: "いちばん小さいものに合わせます。", en: "Matches the smallest one." },
            addStrokeOnly: { ja: "マスクパスに線を追加します。塗りは変えません。", en: "Adds a stroke to the mask path, leaving the fill alone." },
            roundCorners:  { ja: "マスクの角を丸めます。右の欄で半径を指定します。", en: "Rounds the corners of the mask. The field on the right sets the radius." },
            roundRadius:   { ja: "角丸の半径（pt）です。", en: "Radius of the rounded corners, in points." },
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
        }
    };

    function createHexagonPath(targetLayer, centerX, centerY, r, rotationDeg) {
        var pts = [];
        var base = rotationDeg * Math.PI / 180;
        for (var i = 0; i < 6; i++) {
            var a = base + (Math.PI / 3) * i; // 60deg step
            var x = centerX + r * Math.cos(a);
            var y = centerY + r * Math.sin(a);
            pts.push([x, y]);
        }
        var p = targetLayer.pathItems.add();
        p.setEntirePath(pts);
        p.closed = true;
        p.stroked = false;
        p.filled = false;
        return p;
    }

    // 正八角形パスを作成 / Create a regular octagon path
    function createOctagonPath(targetLayer, centerX, centerY, r, rotationDeg) {
        var pts = [];
        var base = rotationDeg * Math.PI / 180;
        for (var i = 0; i < 8; i++) {
            var a = base + (Math.PI / 4) * i; // 45deg step
            var x = centerX + r * Math.cos(a);
            var y = centerY + r * Math.sin(a);
            pts.push([x, y]);
        }
        var p = targetLayer.pathItems.add();
        p.setEntirePath(pts);
        p.closed = true;
        p.stroked = false;
        p.filled = false;
        return p;
    }

    /* 角丸 LiveEffect / Round corners LiveEffect */
    function createRoundCornersEffectXML(radius) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radius);
    }

    function applyRoundCornersLiveEffect(targetItem, radius) {
        if (!targetItem) return false;
        var r = Number(radius);
        if (isNaN(r) || r < 0) return false;

        var xml = createRoundCornersEffectXML(r);

        try {
            targetItem.applyEffect(xml);
            return true;
        } catch (e) {
            return false;
        }
    }

    /*
    作業用レイヤーを取得/作成 / Get or create a reusable work layer
    - 名称 / Name: _clip_work
    - 既存がロック/テンプレでもこのレイヤーは常に編集可能に設定
    */
    function getOrCreateWorkLayer(doc) {
        var name = "_clip_work";
        var lyr = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                lyr = doc.layers[i];
                break;
            }
        }
        if (!lyr) {
            lyr = doc.layers.add();
            lyr.name = name;
        }
        // ensure editable
        lyr.locked = false;
        lyr.visible = true;
        lyr.isTemplate = false;
        return lyr;
    }

    /*
    配置/埋め込み画像を「元サイズの長方形マスク」でクリップグループ化 / Wrap image in a clipping group (rect mask = image bounds)
    - 画像自体はスケールしない
    - ロック/テンプレレイヤーの場合は作業用レイヤーに作成
    */
    function wrapImageToClipGroup(doc, imageItem) {
        if (!doc || !imageItem) return null;
        if (!(imageItem.typename === 'PlacedItem' || imageItem.typename === 'RasterItem')) return null;

        // すでにクリップグループ内なら何もしない
        try {
            if (imageItem.parent && imageItem.parent.typename === 'GroupItem' && imageItem.parent.clipped) {
                return imageItem.parent;
            }
        } catch (e0) { }

        var b = null;
        try { b = imageItem.visibleBounds; } catch (e1) { b = null; }
        if (!b || b.length !== 4) return null;

        // Illustrator bounds: [L, T, R, B]
        var left = b[0];
        var top = b[1];
        var right = b[2];
        var bottom = b[3];
        var w = right - left;
        var h = top - bottom;
        if (w === 0 || h === 0) return null;

        var parentLayer = imageItem.layer;
        var isLockedOrTemplate = parentLayer.locked || parentLayer.isTemplate;
        var targetLayer = isLockedOrTemplate ? getOrCreateWorkLayer(doc) : parentLayer;

        // 長方形マスク（元画像の外接）/ Rect mask equals image bounds
        var rect = targetLayer.pathItems.rectangle(top, left, w, h);
        rect.stroked = false;
        rect.filled = false;

        var g = targetLayer.groupItems.add();
        // move image & rect into group (rect should be on top for clipping path)
        imageItem.moveToBeginning(g);
        rect.moveToBeginning(g);

        g.clipped = true;
        g.selected = true;
        return g;
    }

    function __getBoundsSize(b) {
        if (!b || b.length !== 4) return null;
        var w = b[2] - b[0];
        var h = b[1] - b[3];
        return { w: Math.abs(w), h: Math.abs(h) };
    }

    // クリップグループ内のクリップパス（マスク形状）の bounds を取得（基準算出/サイズ計測に使用）
    function __getClipPathBoundsFromGroup(groupItem) {
        if (!groupItem || groupItem.typename !== 'GroupItem') return null;

        var clipPath = null;
        try {
            for (var i = 0; i < groupItem.pageItems.length; i++) {
                var pi = groupItem.pageItems[i];
                if (pi && pi.typename === 'PathItem' && pi.clipping === true) {
                    clipPath = pi;
                    break;
                }
            }
        } catch (e0) { clipPath = null; }

        if (!clipPath) return null;

        var b = null;
        try { b = clipPath.geometricBounds; } catch (e1) { b = null; }
        if (!b || b.length !== 4) {
            try { b = clipPath.visibleBounds; } catch (e2) { b = null; }
        }
        return (b && b.length === 4) ? b : null;
    }

    // 選択内の「クリップグループ」から基準サイズ（最大/最小: 面積）を取得
    function __getTargetSizeFromSelection(sel, mode) {
        if (!sel || sel.length < 2) return null;
        var useMin = (mode === 'min');

        var targetW = 0, targetH = 0;
        var bestArea = useMin ? Infinity : -1;

        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            // 基準は「クリップグループ」のみ（clipped==true）/ Use clipping groups only
            if (!(it.typename === 'GroupItem' && it.clipped === true)) continue;

            var b = __getClipPathBoundsFromGroup(it);
            var sz = __getBoundsSize(b);
            if (!sz) continue;

            var area = sz.w * sz.h;
            if (useMin) {
                if (area > 0 && area < bestArea) {
                    bestArea = area;
                    targetW = sz.w;
                    targetH = sz.h;
                }
            } else {
                if (area > bestArea) {
                    bestArea = area;
                    targetW = sz.w;
                    targetH = sz.h;
                }
            }
        }

        if ((useMin && !isFinite(bestArea)) || (!useMin && bestArea <= 0) || targetW <= 0 || targetH <= 0) return null;
        return { w: targetW, h: targetH };
    }

    // 生成したクリップグループ全体を基準サイズに合わせてリサイズ（縦横比は維持：長辺合わせ）
    function __resizeGroupsToTargetSize(groups, targetSize) {
        if (!groups || !groups.length || !targetSize) return;

        var targetW = targetSize.w;
        var targetH = targetSize.h;
        if (!targetW || !targetH) return;

        // 基準は「長辺」/ Use long side as reference to avoid distortion
        var targetLong = Math.max(targetW, targetH);
        if (targetLong <= 0) return;

        for (var i = 0; i < groups.length; i++) {
            var g = groups[i];
            if (!g) continue;

            var b = __getClipPathBoundsFromGroup(g);
            if (!b || b.length !== 4) continue;

            var w = Math.abs(b[2] - b[0]);
            var h = Math.abs(b[1] - b[3]);
            if (w === 0 || h === 0) continue;

            var longSide = Math.max(w, h);
            if (longSide === 0) continue;

            // すでに長辺が同じならスキップ
            if (Math.abs(longSide - targetLong) < 0.001) continue;

            var scale = (targetLong / longSide) * 100;

            try {
                // sx == sy で縦横比維持 / Keep aspect ratio
                g.resize(scale, scale, true, true, true, true, true, Transformation.CENTER);
            } catch (e3) { }
        }
    }

    /*
    画像に最小正方形を追加し、クリッピンググループを作成 / Build a clipping group with the minimal square
    - 入力 / Input: doc (Document), image (PlacedItem|RasterItem), groups (Array)
    - 動作 / Behavior: visibleBounds から正方形を作成→画像と同グループに配置→グループをクリップ化
    */
    function processImage(doc, image) {
        var bounds = image.visibleBounds;
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];

        // 基準サイズ / Base sizes
        var sideLength = Math.min(width, height); // 内接正方形の一辺 / inscribed square side
        var squareSize = sideLength; // 内接正方形の一辺 / inscribed square side
        var centerX = bounds[0] + width / 2;
        var centerY = bounds[1] - height / 2;

        // 現在のレイヤーが編集可能か確認 / Check if current layer is editable
        var parentLayer = image.layer;
        var isLockedOrTemplate = parentLayer.locked || parentLayer.isTemplate;

        // 編集可能なレイヤーを使用 / Choose a writable layer
        var targetLayer = isLockedOrTemplate ? getOrCreateWorkLayer(doc) : parentLayer;

        // 形状作成 / Create shape
        var shapePath = null;

        if (__shapeMode === 'circle') {
            // 正円（円） / Circle (perfect circle)
            shapePath = targetLayer.pathItems.ellipse(centerY + sideLength / 2, centerX - sideLength / 2, sideLength, sideLength);
            shapePath.stroked = false;
            shapePath.filled = false;
        } else if (__shapeMode === 'hexA') {
            // 六角形A / Hexagon A (rotation 0deg)
            shapePath = createHexagonPath(targetLayer, centerX, centerY, sideLength / 2, 0);
        } else if (__shapeMode === 'hexB') {
            // 六角形B / Hexagon B (rotation 30deg)
            shapePath = createHexagonPath(targetLayer, centerX, centerY, sideLength / 2, 30);
        } else if (__shapeMode === 'octagon') {
            // 八角形 / Octagon
            shapePath = createOctagonPath(targetLayer, centerX, centerY, sideLength / 2, 22.5);
        } else {
            // 正方形 / Square
            shapePath = targetLayer.pathItems.rectangle(centerY + squareSize / 2, centerX - squareSize / 2, squareSize, squareSize);
            shapePath.stroked = false;
            shapePath.filled = false;
        }

        var group = targetLayer.groupItems.add();
        image.moveToBeginning(group);
        shapePath.moveToBeginning(group);
        group.clipped = true;
        group.selected = true; // 生成直後に即選択 / Select immediately after creation
        return group;
    }

    function main() {
        // 安全ガード / Safety guard
        if (!app.documents.length || !app.selection.length) {
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = app.selection;

        // 単体の配置/埋め込み画像は、事前に「元サイズ長方形マスク」でクリップグループ化してから処理
        var normalizedSelection = [];
        for (var ni = 0; ni < selectedItems.length; ni++) {
            var it0 = selectedItems[ni];
            if (it0 && (it0.typename === 'PlacedItem' || it0.typename === 'RasterItem')) {
                var cg = wrapImageToClipGroup(doc, it0);
                if (cg) {
                    normalizedSelection.push(cg);
                } else {
                    normalizedSelection.push(it0);
                }
            } else {
                normalizedSelection.push(it0);
            }
        }
        selectedItems = normalizedSelection;

        // 「大きさを揃える」の基準は、形状変更後に生成されたクリップグループから取得する
        var __targetSizeForGroups = null;
        doc.selection = null; // 先に選択をクリア / Clear selection first
        var createdGroups = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var item = selectedItems[i];

            // クリッピングマスクの処理 / Handle clipping mask groups
            if (item.typename === 'GroupItem' && item.clipped) {
                item.clipped = false;

                var itemsToProcess = [];
                for (var j = item.pageItems.length - 1; j >= 0; j--) {
                    var pageItem = item.pageItems[j];
                    if (pageItem.typename === 'PlacedItem' || pageItem.typename === 'RasterItem') {
                        itemsToProcess.push(pageItem);
                    } else {
                        pageItem.remove();
                    }
                }

                // 解除後のグループ内にある画像を処理 / Process images extracted from the released group
                for (var k = 0; k < itemsToProcess.length; k++) {
                    var imageItem = itemsToProcess[k];
                    var g1 = processImage(doc, imageItem);
                    if (g1) createdGroups.push(g1);
                }
            }
            // 単体の配置/埋め込み画像を処理 / Handle standalone placed/embedded images
            else if (item.typename === 'PlacedItem' || item.typename === 'RasterItem') {
                var g2 = processImage(doc, item);
                if (g2) createdGroups.push(g2);
            }
        }

        // クリップグループ後の処理 / Post process for clipping groups
        if (createdGroups.length > 0) {
            var prevSel = createdGroups.slice(0);

            // 「大きさを揃える」ONなら、形状変更後のクリップグループから基準サイズ（最大/最小）を決めて揃える
            if (__sameSizeEnabled) {
                __targetSizeForGroups = __getTargetSizeFromSelection(createdGroups, __sameSizeMode);
                if (__targetSizeForGroups) {
                    __resizeGroupsToTargetSize(createdGroups, __targetSizeForGroups);
                }
            }

            // 角丸（チェック時のみ）
            if (__roundCornersEnabled) {
                for (var r = 0; r < createdGroups.length; r++) {
                    applyRoundCornersLiveEffect(createdGroups[r], __roundCornerRadius);
                }
            }

            if (__appearanceMode === 'clipGroup') {
                // 「ケイ線を追加」OFF の場合は、クリップグループ作成のみ（ケイ線は付けない）
            } else {
                // 「ケイ線を追加」ON：ケイ線を追加 + Exclude
                for (var s = 0; s < createdGroups.length; s++) {
                    try {
                        doc.selection = [createdGroups[s]];
                        app.executeMenuCommand('Adobe New Stroke Shortcut');
                        app.executeMenuCommand('Live Pathfinder Exclude');
                    } catch (e) { }
                }
            }

            doc.selection = prevSel;
        }
    }

    /* ダイアログボックス / Dialog */
    function showDialog() {
        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dialog);

        // 2カラム / Two columns
        var cols = dialog.add('group');
        cols.orientation = 'row';
        cols.alignChildren = ['fill', 'top'];
        cols.spacing = COLUMN_SPACING;

        // 左：形状 / Left: Shape
        var shapePanel = cols.add('panel', undefined, getLabel('panel.shape'));
        setupPanel(shapePanel, 6);

        var rbNoChange = shapePanel.add('radiobutton', undefined, getLabel('radio.noChange'));
        rbNoChange.helpTip = getLabel('tooltip.noChange');
        rbNoChange.value = true;

        var rbSquare = shapePanel.add('radiobutton', undefined, getLabel('radio.square'));
        rbSquare.helpTip = getLabel('tooltip.square');
        var rbCircle = shapePanel.add('radiobutton', undefined, getLabel('radio.circle'));
        rbCircle.helpTip = getLabel('tooltip.circle');
        var rbHexA = shapePanel.add('radiobutton', undefined, getLabel('radio.hexA'));
        rbHexA.helpTip = getLabel('tooltip.hexA');
        var rbHexB = shapePanel.add('radiobutton', undefined, getLabel('radio.hexB'));
        rbHexB.helpTip = getLabel('tooltip.hexB');
        var rbOctagon = shapePanel.add('radiobutton', undefined, getLabel('radio.octagon'));
        rbOctagon.helpTip = getLabel('tooltip.octagon');

        // 右カラム / Right column
        var rightCol = cols.add('group');
        rightCol.orientation = 'column';
        rightCol.alignChildren = ['fill', 'top'];
        rightCol.spacing = 10;

        // 右：アピアランス / Right: Appearance
        var appearancePanel = rightCol.add('panel', undefined, getLabel('panel.appearance'));
        setupPanel(appearancePanel, 6);

        // オプションパネル / Options panel（アピアランスの外）
        var optionPanel = rightCol.add('panel', undefined, getLabel('panel.option'));
        setupPanel(optionPanel, 6);

        // 選択が1つ以下なら「複数オブジェクト」パネルはディム表示
        var selCount = 0;
        try {
            if (app.documents.length && app.activeDocument.selection) {
                selCount = app.activeDocument.selection.length;
            }
        } catch (e) { selCount = 0; }
        optionPanel.enabled = (selCount > 1);

        var cbSameSize = optionPanel.add('checkbox', undefined, getLabel('checkbox.sameSize'));
        cbSameSize.helpTip = getLabel('tooltip.sameSize');
        cbSameSize.value = false;

        var sameSizeModeGroup = optionPanel.add('group');
        sameSizeModeGroup.orientation = 'column';
        sameSizeModeGroup.alignChildren = 'left';
        sameSizeModeGroup.margins = [20, 0, 0, 0];

        var rbSameSizeMax = sameSizeModeGroup.add('radiobutton', undefined, getLabel('radio.sameSizeMax'));
        rbSameSizeMax.helpTip = getLabel('tooltip.sameSizeMax');
        var rbSameSizeMin = sameSizeModeGroup.add('radiobutton', undefined, getLabel('radio.sameSizeMin'));
        rbSameSizeMin.helpTip = getLabel('tooltip.sameSizeMin');
        rbSameSizeMax.value = true;

        function updateSameSizeModeUI() {
            sameSizeModeGroup.enabled = (cbSameSize.value === true);
        }
        cbSameSize.onClick = updateSameSizeModeUI;
        updateSameSizeModeUI();

        var cbAddStrokeOnly = appearancePanel.add('checkbox', undefined, getLabel('checkbox.addStrokeOnly'));
        cbAddStrokeOnly.helpTip = getLabel('tooltip.addStrokeOnly');

        // 選択オブジェクト全体の外接矩形から角丸のデフォルト値を算出（PlacedImageStroke.jsx）
        function calcDefaultRoundRadiusFromSelection(sel) {
            try {
                if (!sel || sel.length === 0) return 10;

                var top = -Infinity, left = Infinity, bottom = Infinity, right = -Infinity;
                for (var i = 0, n = sel.length; i < n; i++) {
                    var b = sel[i].geometricBounds; // [top, left, bottom, right]
                    if (!b || b.length !== 4) continue;
                    if (b[0] > top) top = b[0];
                    if (b[1] < left) left = b[1];
                    if (b[2] < bottom) bottom = b[2];
                    if (b[3] > right) right = b[3];
                }

                if (!isFinite(top) || !isFinite(left) || !isFinite(bottom) || !isFinite(right)) return 10;

                var height = Math.abs(top - bottom);
                var width = Math.abs(right - left);
                var A = height + width;
                var B = A / 2;
                var r = Math.max(0, B / 20);
                return r;
            } catch (e) {
                return 10;
            }
        }

        var defaultRoundRadius = 10;
        try {
            if (app.documents.length) {
                defaultRoundRadius = calcDefaultRoundRadiusFromSelection(app.activeDocument.selection);
            }
        } catch (e) { }
        defaultRoundRadius = Math.round(defaultRoundRadius);

        // 角丸（クリップグループ用）
        var roundRow = appearancePanel.add('group');
        roundRow.orientation = 'row';
        roundRow.alignChildren = ['left', 'center'];
        roundRow.margins = [0, 0, 0, 0];

        var cbRoundCorners = roundRow.add('checkbox', undefined, getLabel('checkbox.roundCorners'));
        cbRoundCorners.helpTip = getLabel('tooltip.roundCorners');
        cbRoundCorners.value = false;

        // ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field
        var roundRadiusGroup = roundRow.add('group');
        roundRadiusGroup.orientation = 'row';
        roundRadiusGroup.alignChildren = ['left', 'center'];
        roundRadiusGroup.spacing = 0;
        roundRadiusGroup.margins = 0;
        var editRoundRadius;
        var roundRadiusStepper = addStepper(roundRadiusGroup, function () { return editRoundRadius; }, { min: 0 });
        editRoundRadius = roundRadiusGroup.add('edittext', undefined, String(defaultRoundRadius));
        editRoundRadius.helpTip = getLabel('tooltip.roundRadius');
        editRoundRadius.characters = 6;
        bindSteppedArrowKeys(editRoundRadius, roundRadiusStepper); // ↑↓キーも∧∨と同じ処理で増減
        var roundUnit = roundRow.add('statictext', undefined, 'pt');

        // default
        cbAddStrokeOnly.value = false;

        /**
         * 角丸の半径欄（∧∨・単位を含む）の有効／無効を切り替える
         * @param {boolean} isEnabled - 有効にするなら true
         * @returns {void}
         */
        function setRoundRadiusEnabled(isEnabled) {
            editRoundRadius.enabled = isEnabled;
            roundUnit.enabled = isEnabled;
            roundRadiusStepper.enabled = isEnabled;
            redrawSteppersIn(roundRadiusStepper); // ∧∨は自作描画なので描き直す
        }

        function updateRoundRadiusUI() {
            // 現在の形状選択をUIから判定（「変更なし」の場合は現在値を参照）
            var effectiveShapeMode = __shapeMode;
            if (rbCircle.value) effectiveShapeMode = 'circle';
            else if (rbHexA.value) effectiveShapeMode = 'hexA';
            else if (rbHexB.value) effectiveShapeMode = 'hexB';
            else if (rbSquare.value) effectiveShapeMode = 'square';
            // rbNoChange の場合は __shapeMode のまま

            // 正円の場合は角丸は無効（ディム表示）にする
            if (effectiveShapeMode === 'circle') {
                cbRoundCorners.enabled = false;
                setRoundRadiusEnabled(false);
                return;
            }

            // 角丸チェックは操作可能（ケイ線ON/OFFではディムにしない）
            cbRoundCorners.enabled = true;

            // 角丸チェックに連動して数値だけ enable/disable
            setRoundRadiusEnabled(cbRoundCorners.value === true);
        }

        // UI変更時にUI状態を更新
        cbAddStrokeOnly.onClick = updateRoundRadiusUI;
        rbNoChange.onClick = updateRoundRadiusUI;
        rbSquare.onClick = updateRoundRadiusUI;
        rbCircle.onClick = updateRoundRadiusUI;
        rbHexA.onClick = updateRoundRadiusUI;
        rbHexB.onClick = updateRoundRadiusUI;
        cbRoundCorners.onClick = updateRoundRadiusUI;

        updateRoundRadiusUI();

        editRoundRadius.onChange = function () {
            var v = Number(editRoundRadius.text);
            if (isNaN(v)) return;
            editRoundRadius.text = String(v);
        };

        // ボタン / Buttons
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        btnOK.onClick = function () {
            if (rbNoChange.value) {
                // keep current __shapeMode
            } else if (rbCircle.value) {
                __shapeMode = 'circle';
            } else if (rbHexA.value) {
                __shapeMode = 'hexA';
            } else if (rbHexB.value) {
                __shapeMode = 'hexB';
            } else if (rbOctagon.value) {
                __shapeMode = 'octagon';
            } else {
                __shapeMode = 'square';
            }

            // アピアランス設定の確定 / Confirm appearance settings
            __appearanceMode = cbAddStrokeOnly.value ? 'strokeOnly' : 'clipGroup';
            __sameSizeEnabled = (cbSameSize.value === true);
            __sameSizeMode = (rbSameSizeMin.value === true) ? 'min' : 'max';

            // 正円は角丸無効 + ケイ線ON時は角丸適用しない
            __roundCornersEnabled = (__shapeMode !== 'circle') && (cbRoundCorners.value === true);
            __roundCornerRadius = __roundCornersEnabled ? Number(editRoundRadius.text) : 0;
            if (__roundCornersEnabled && (isNaN(__roundCornerRadius) || __roundCornerRadius < 0)) {
                __roundCornerRadius = 0;
            }

            dialog.close(1);
        };

        btnCancel.onClick = function () {
            dialog.close(0);
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        return dialog.show() === 1;
    }

    if (showDialog()) {
        main();
    }

})();
