#target illustrator
#targetengine "SmartBatchImporterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

開いているドキュメントやフォルダー内の Illustrator ファイル（.ai / .svg / .eps）をまとめて読み込み、1つのドキュメントに、全体が正方形に近いグリッドで並べます。
［アートボード単位］で読み込むと、元のアートボードの大きさと内容の位置を保ったアートボードを作ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBatchImporter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8180588e5630

### Overview

Imports open documents or Illustrator files (.ai / .svg / .eps) from a folder into one document, arranged in a grid that comes out close to square.
With Import per artboard, each source artboard becomes an artboard of the same size with its content in the same position.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBatchImporter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartBatchImporter";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-05";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBatchImporter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBatchImporter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8180588e5630"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var IMPORT_SETTINGS = {
        itemSpacing: 100,              // アイテムの間隔の初期値（pt）/ Default gap between items (pt)
        splitFileCount: 20,            // 分割するときの1ドキュメントあたりのファイル数の初期値 / Default files per document when splitting
        artboardMargin: 8.5,           // アートボードと内容の余白（pt）/ Margin around content within an artboard (pt)
        labelFont: "HiraginoSans-W3",  // ラベルのフォント / Font used for labels
        labelSize: 9,                  // ラベルの文字サイズ（pt）/ Label font size (pt)
        labelLayerName: "_label"       // ラベルを置くレイヤー名 / Name of the layer that holds labels
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

    var LABELS = {
        dialog: {
            title: { ja: "ファイル一括読み込み", en: "Batch Import Files" },
            newDocSettings: { ja: "新規ドキュメントの設定", en: "New Document Settings" }
        },
        panel: {
            source: { ja: "読み込み対象", en: "Source" },
            filter: { ja: "フィルター", en: "Filter" },
            destination: { ja: "読み込み先", en: "Destination" },
            colorMode: { ja: "カラーモード", en: "Color Mode" },
            resolution: { ja: "解像度", en: "Resolution" },
            docSize: { ja: "サイズ", en: "Size" },
            importOptions: { ja: "読み込みオプション", en: "Import Options" },
            targetArtboards: { ja: "対象アートボード", en: "Target Artboards" }
        },
        radio: {
            openDocs: { ja: "開いているドキュメント", en: "Open documents" },
            specifyFolder: { ja: "フォルダーを指定", en: "Specify folder" },
            currentDoc: { ja: "現在のドキュメント", en: "Current document" },
            newDoc: { ja: "新規ドキュメント", en: "New document" },
            artboardFirst: { ja: "1のみ", en: "Artboard 1 only" },
            artboardAll: { ja: "すべて", en: "All" },
            artboardSpecify: { ja: "指定", en: "Specify" },
            closeDoc: { ja: "閉じる", en: "Close" },
            keepOpen: { ja: "開いたまま", en: "Keep open" }
        },
        checkbox: {
            byArtboard: { ja: "アートボード単位", en: "Import per artboard" },
            attachLabel: { ja: "ファイル名をラベルとして追加", en: "Add file names as labels" },
            includeGuides: { ja: "ガイドを含める", en: "Include guides" },
            scale: { ja: "拡大・縮小", en: "Scale" },
            includeSubfolders: { ja: "サブフォルダーを含める", en: "Include subfolders" },
            splitDocs: { ja: "分割", en: "Split every" }
        },
        dropdown: {
            presetCustom: { ja: "カスタム", en: "Custom" },
            presetA4: { ja: "A4（210 × 297 mm）", en: "A4 (210 × 297 mm)" },
            presetFullHD: { ja: "フルHD（1920 × 1080 px）", en: "Full HD (1920 × 1080 px)" },
            presetLargeCanvas: { ja: "ラージカンバス", en: "Large Canvas" }
        },
        fieldLabel: {
            fileType: { ja: "ファイル形式", en: "File format" },
            fileName: { ja: "ファイル名", en: "File name" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            unit: { ja: "単位", en: "Unit" },
            afterImport: { ja: "読み込み後", en: "After import" },
            itemSpacing: { ja: "間隔", en: "Spacing" },
            profile: { ja: "プロファイル", en: "Profile" },
            splitUnit: { ja: "ファイルごと", en: "files" }
        },
        button: {
            chooseFolder: { ja: "選択...", en: "Choose..." },
            newDocSettings: { ja: "設定...", en: "Settings..." },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        progress: {
            title: { ja: "ファイルを読み込み中...", en: "Importing files..." },
            count: { ja: "読み込み", en: "Imported" }
        },
        prompt: {
            selectFolder: {
                ja: "読み込むファイルが入ったフォルダーを選択してください",
                en: "Select a folder that contains files to import"
            }
        },
        alert: {
            noValidFile: { ja: "読み込むファイルが見つかりませんでした。", en: "No files to import were found." },
            noCurrentDoc: {
                ja: "「現在のドキュメント」に読み込むには、ドキュメントを開いておいてください。",
                en: "Open a document first to import into the current document."
            },
            invalidArtboardSpec: {
                ja: "対象アートボードの指定が正しくありません。番号で指定してください。例: 1, 3-5",
                en: "The target artboard specification is invalid. Specify by number, e.g. 1, 3-5"
            },
            noArtboardImported: {
                ja: "取り込める内容が見つかりませんでした。対象アートボードの番号やファイルの内容を確認してください。",
                en: "Nothing could be imported. Check the target artboard numbers and the file contents."
            },
            invalidNumber: { ja: "幅・高さの値が正しくありません。", en: "The width or height is invalid." },
            invalidSpacing: { ja: "間隔の値が正しくありません。0以上の数値を入力してください。", en: "The spacing is invalid. Enter a number of 0 or more." },
            pasteFail: { ja: "ペーストに失敗しました", en: "Paste failed" },
            cancelled: {
                ja: "読み込みを中断しました。ここまでに読み込んだ内容はドキュメントに残っています。",
                en: "Import was stopped. Items imported so far remain in the document."
            },
            invalidScale: {
                ja: "拡大・縮小の値が正しくありません。0より大きい数値を入力してください。",
                en: "The scale value is invalid. Enter a number greater than 0."
            },
            invalidFilter: {
                ja: "ファイル名フィルターの正規表現が正しくありません。式を見直すか、空欄にしてください。",
                en: "The file-name filter is not a valid regular expression. Fix it or clear the field."
            }
        },
        confirm: {
            discardUnsaved: {
                ja: "未保存の変更があるドキュメントが含まれています。［読み込み後］が［閉じる］なので、保存せずに閉じます。続行しますか？",
                en: "Some documents have unsaved changes. \"After import\" is set to \"Close\", so they will be closed without saving. Continue?"
            }
        },
        tooltip: {
            openDocs: {
                ja: "Illustrator で開いているドキュメントを読み込みます。",
                en: "Imports the documents currently open in Illustrator."
            },
            specifyFolder: {
                ja: "指定したフォルダー内の .ai / .svg / .eps ファイルを読み込みます。",
                en: "Imports .ai / .svg / .eps files in the chosen folder."
            },
            chooseFolder: {
                ja: "読み込むファイル（.ai / .svg / .eps）が入ったフォルダーを選びます。",
                en: "Choose a folder that contains the files to import (.ai / .svg / .eps)."
            },
            fileType: {
                ja: "フォルダー指定のときに読み込むファイル形式を選びます。",
                en: "Choose which file formats to import from the folder."
            },
            fileNameFilter: {
                ja: "正規表現でファイル名を絞り込みます（大文字・小文字は区別しません）。",
                en: "Filters file names with a regular expression (case-insensitive)."
            },
            currentDoc: {
                ja: "アクティブなドキュメントに読み込みます。新規ドキュメントの設定は使いません。",
                en: "Imports into the active document. The new document settings are not used."
            },
            newDoc: {
                ja: "［設定...］の内容で新規ドキュメントを作成し、そこに読み込みます。",
                en: "Creates a new document with the Settings... values and imports into it."
            },
            newDocSettings: {
                ja: "新規ドキュメントのプロファイル・カラーモード・解像度・サイズを設定します。",
                en: "Sets the profile, color mode, resolution and size of the new document."
            },
            colorMode: {
                ja: "新規ドキュメントのカラーモードです。読み込み元の色を変換するものではありません。",
                en: "Color mode of the new document. It does not convert the colors of the source files."
            },
            profile: {
                ja: "［新規ドキュメント］のプロファイルです。スウォッチやブラシなどの初期内容が決まります。",
                en: "New document profile. It sets the starting swatches, brushes and so on."
            },
            resolution: {
                ja: "新規ドキュメントのラスタライズ効果の解像度です。",
                en: "Raster effects resolution of the new document."
            },
            sizePreset: {
                ja: "プリセットを選ぶと幅・高さ・単位が入ります。A4 はカラーモードも CMYK にします。",
                en: "Choosing a preset fills in the width, height and unit. A4 also switches the color mode to CMYK."
            },
            unit: {
                ja: "幅・高さの単位です。切り替えると値を換算します。",
                en: "Unit of the width and height. Switching converts the values."
            },
            byArtboard: {
                ja: "アートボードごとに読み込み、元のアートボードの大きさと内容の位置を保ちます。ロック・非表示のオブジェクトも含めます。",
                en: "Imports each artboard separately and keeps the original artboard size and content position. Locked/hidden objects are included."
            },
            attachLabel: {
                ja: "読み込んだ内容の下に、元のファイル名のラベルを追加します。",
                en: "Adds a source file-name label below each imported item."
            },
            includeGuides: {
                ja: "ガイドも読み込みます（［ガイドをロック］がオンでも読み込みます）。横はアートボードの短辺以下、縦は長辺以下の長さのガイドを含め、それより長いルーラーガイドは除きます。",
                en: "Imports guides too, even when Lock Guides is on. Horizontal guides up to the artboard's shorter side and vertical guides up to its longer side are included; longer ruler guides are excluded."
            },
            scale: {
                ja: "読み込む内容とアートボードを指定の％で拡大・縮小します。線幅も同じ比率で変わります。",
                en: "Scales imported content and artboards by the specified percentage. Stroke widths scale by the same ratio."
            },
            includeSubfolders: {
                ja: "選んだフォルダーの中のサブフォルダーにあるファイルも読み込みます。",
                en: "Also imports files in subfolders of the chosen folder."
            },
            splitDocs: {
                ja: "指定したファイル数ごとに、別の新規ドキュメントに分けて読み込みます。",
                en: "Imports into a separate new document for every specified number of files."
            },
            itemSpacing: {
                ja: "並べるアイテム（アートボード）どうしの間隔です。",
                en: "Gap between the arranged items (artboards)."
            },
            targetArtboards: {
                ja: "アートボード単位で読み込むときに、取り込むアートボードを選びます。",
                en: "Choose which artboards to import when importing per artboard."
            },
            artboardSpecify: {
                ja: "取り込むアートボードを番号で指定します。例: 1, 3-5（カンマ区切り・範囲指定可）",
                en: "Specify artboards to import by number, e.g. 1, 3-5 (comma-separated, ranges allowed)."
            },
            closeDoc: {
                ja: "読み込み後に元のドキュメントを保存せずに閉じます。未保存の変更は失われます。",
                en: "Closes the source documents after import without saving. Unsaved changes will be lost."
            },
            keepOpen: {
                ja: "読み込み後も元のドキュメントを開いたままにします。一時的に解除したロック・非表示は元に戻します。",
                en: "Keeps the source documents open after import. Lock/hidden states changed temporarily are restored."
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
        }
    };

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
    // 単位とサイズ / Units and sizes
    // =========================================

    /* 新規ドキュメントの幅・高さの単位と、1単位あたりのポイント数（Illustrator の px は 1pt）
       Units for the new document's size and points per unit (a px is 1pt in Illustrator) */
    var SIZE_UNITS = ["mm", "px"];
    var SIZE_UNIT_POINTS = { mm: 72 / 25.4, px: 1, inch: 72 };

    /* 幅・高さの単位に合わせた、新規ドキュメントの定規の単位 / ruler units of the new document per size unit */
    var SIZE_UNIT_RULER_UNITS = { mm: RulerUnits.Millimeters, px: RulerUnits.Pixels };

    /* サイズのプリセット（並びはドロップダウンと同じ）。width / height は unit での値、displayUnit は幅・高さ欄の単位。
       カスタムは値を入れず単位だけ切り替える。isPrint ならカラーモードを CMYK にする
       Size presets in dropdown order; values are in unit, shown in displayUnit. Custom only switches the unit */
    var SIZE_PRESETS = [
        { labelKey: "dropdown.presetCustom", displayUnit: "px" },
        { labelKey: "dropdown.presetA4", unit: "mm", width: 210, height: 297, displayUnit: "mm", isPrint: true },
        { labelKey: "dropdown.presetFullHD", unit: "px", width: 1920, height: 1080, displayUnit: "px" },
        { labelKey: "dropdown.presetLargeCanvas", unit: "inch", width: 2270, height: 2270, displayUnit: "px" }
    ];

    /* 定規の単位コードに対応する表示ラベルと、1単位あたりのポイント数
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

    var GROUP_SPACING = 8;      /* グループ内の要素間隔 / spacing inside groups */
    var FIELD_LABEL_WIDTH = 40; /* 幅・高さ・単位の項目名の幅 / label width for width / height / unit */
    var FILTER_LABEL_WIDTH = 100; /* ファイル形式・ファイル名の項目名の幅 / label width for file format / file name */

    /**
     * グループの共通設定（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - 要素間隔（省略時は GROUP_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : GROUP_SPACING;
    }

    /**
     * 右揃えで固定幅の項目名（コロン付き）を行に追加する
     * @param {Group} parentRow - 追加先の行
     * @param {string} labelKey - LABELS のパス
     * @param {number} labelWidth - 項目名の幅（px）
     * @returns {StaticText} 項目名
     */
    function addRowLabel(parentRow, labelKey, labelWidth) {
        var rowLabel = parentRow.add("statictext", undefined, labelText(labelKey));
        rowLabel.preferredSize = [labelWidth, 20];
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 行に「∧∨・入力欄」を隙間0で突き合わせて追加する。↑↓キーも∧∨と同じ処理で増減する。
     * 直接入力した値も確定時に計算・整数化・下限・上限・単位へそろえ、数値でなければ直前の値に戻す
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - addStepper() に渡す step / min / max / integer / unit / onStep
     * @returns {EditText} 入力欄（∧∨は .stepperGroup、包む group は .parent）
     */
    function addSteppedInput(parentRow, initialText, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);

        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, stepOptions);
        };
        return numberInput;
    }

    // =========================================
    // 読み込みの下請け / Import helpers
    // =========================================

    /**
     * ラベル用のレイヤーを取得し、無ければ作成する
     * @param {Document} destDoc - 読み込み先のドキュメント
     * @returns {Layer} ラベル用のレイヤー
     */
    function getOrCreateLabelLayer(destDoc) {
        try {
            return destDoc.layers.getByName(IMPORT_SETTINGS.labelLayerName);
        } catch (e) {
            /* getByName は見つからないと例外 / getByName throws when the layer is missing */
            var labelLayer = destDoc.layers.add();
            labelLayer.name = IMPORT_SETTINGS.labelLayerName;
            return labelLayer;
        }
    }

    /**
     * 2つの矩形（[左, 上, 右, 下]、上 > 下）が重なるか判定する
     * @param {number[]} rectA - 矩形A
     * @param {number[]} rectB - 矩形B
     * @returns {boolean} 重なれば true
     */
    function rectsIntersect(rectA, rectB) {
        return !(rectA[2] < rectB[0] || rectA[0] > rectB[2] || rectA[3] > rectB[1] || rectA[1] < rectB[3]);
    }

    /**
     * 例外を出さずにプロパティへ代入する（削除済み・読み取り専用で失敗しても続ける）
     * @param {Object} target - 代入先
     * @param {string} propertyName - プロパティ名
     * @param {*} value - 値
     * @returns {void}
     */
    function setPropertySafely(target, propertyName, value) {
        try {
            target[propertyName] = value;
        } catch (e) { /* 1件の失敗で復元全体を止めない / one failure must not stop the whole restore */ }
    }

    /**
     * 取り込めるように、レイヤーをすべて、オブジェクトは対象アートボードに重なるものだけロック・非表示を解除する。
     * 変えた分を lockState に控える（restoreLockHiddenState で戻せる）
     * @param {Document} sourceDoc - 読み込み元のドキュメント
     * @param {number[][]|null} artboardRects - 対象アートボードの矩形の配列。null ならすべてのオブジェクトを解除
     * @param {{layers: Array, items: Array}} lockState - 元の状態を控える先（途中で失敗しても控えた分は戻せる）
     * @returns {void}
     */
    function unlockItemsForImport(sourceDoc, artboardRects, lockState) {
        function unlockLayers(layerList) {
            for (var i = 0; i < layerList.length; i++) {
                var layer = layerList[i];
                lockState.layers.push({ ref: layer, locked: layer.locked, visible: layer.visible });
                layer.locked = false;
                layer.visible = true;
                unlockLayers(layer.layers);
            }
        }
        unlockLayers(sourceDoc.layers);

        for (var i = 0; i < sourceDoc.pageItems.length; i++) {
            var pageItem = sourceDoc.pageItems[i];
            if (!pageItem.locked && !pageItem.hidden) continue; /* 既に選択できるものは触らない / leave already-selectable items alone */
            if (artboardRects && !overlapsAnyRect(pageItem, artboardRects)) continue; /* 他のアートボードのものには触れない / leave other artboards alone */
            lockState.items.push({ ref: pageItem, locked: pageItem.locked, hidden: pageItem.hidden });
            pageItem.locked = false;
            pageItem.hidden = false;
        }
    }

    /**
     * オブジェクトが矩形のどれかに重なるか判定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {number[][]} rects - 矩形の配列
     * @returns {boolean} 重なれば true（境界を測れなければ false）
     */
    function overlapsAnyRect(pageItem, rects) {
        var itemBounds;
        try {
            itemBounds = pageItem.geometricBounds;
        } catch (e) {
            return false; /* 境界を持たない種類 / kinds without bounds */
        }
        for (var i = 0; i < rects.length; i++) {
            if (rectsIntersect(itemBounds, rects[i])) return true;
        }
        return false;
    }

    /**
     * 控えたロック・非表示の状態を元に戻す。1件失敗しても残りは続ける
     * @param {{layers: Array, items: Array}} lockState - unlockItemsForImport() で控えた状態
     * @returns {void}
     */
    function restoreLockHiddenState(lockState) {
        for (var i = 0; i < lockState.layers.length; i++) {
            setPropertySafely(lockState.layers[i].ref, "locked", lockState.layers[i].locked);
            setPropertySafely(lockState.layers[i].ref, "visible", lockState.layers[i].visible);
        }
        for (var j = 0; j < lockState.items.length; j++) {
            setPropertySafely(lockState.items[j].ref, "locked", lockState.items[j].locked);
            setPropertySafely(lockState.items[j].ref, "hidden", lockState.items[j].hidden);
        }
    }

    /**
     * "1, 3-5" のような指定を、1始まりのアートボード番号の配列にする（カンマ・空白区切り、範囲指定可）
     * @param {string} specText - 指定の文字列
     * @returns {number[]} アートボード番号（読めない部分は無視）
     */
    function parseArtboardNumbers(specText) {
        var artboardNumbers = [];
        if (!specText) return artboardNumbers;
        var specParts = specText.split(/[,，\s]+/);
        for (var i = 0; i < specParts.length; i++) {
            var specPart = specParts[i];
            var rangeMatch = specPart.match(/^(\d+)\s*[-–~]\s*(\d+)$/);
            if (rangeMatch) {
                var rangeStart = parseInt(rangeMatch[1], 10);
                var rangeEnd = parseInt(rangeMatch[2], 10);
                for (var number = Math.min(rangeStart, rangeEnd); number <= Math.max(rangeStart, rangeEnd); number++) {
                    artboardNumbers.push(number);
                }
            } else if (/^\d+$/.test(specPart)) {
                artboardNumbers.push(parseInt(specPart, 10));
            }
        }
        return artboardNumbers;
    }

    /**
     * 対象の指定から、取り込むアートボードの0始まりの番号を若い順に返す
     * @param {Document} sourceDoc - 読み込み元のドキュメント
     * @param {string} targetMode - "first"（1のみ）/ "all"（すべて）/ "specify"（指定）
     * @param {string} specText - 「指定」のときの番号の文字列
     * @returns {number[]} アートボードの番号（0始まり、重複なし）
     */
    function resolveTargetArtboardIndices(sourceDoc, targetMode, specText) {
        var artboardCount = sourceDoc.artboards.length;
        var isTarget = [];
        var i;
        if (targetMode === "all") {
            for (i = 0; i < artboardCount; i++) isTarget[i] = true;
        } else if (targetMode === "specify") {
            var artboardNumbers = parseArtboardNumbers(specText);
            for (i = 0; i < artboardNumbers.length; i++) isTarget[artboardNumbers[i] - 1] = true; /* 1始まり → 0始まり / 1-based → 0-based */
        } else {
            isTarget[0] = true;
        }
        /* 番号順に拾うので並べ替えは要らない / picking in index order keeps them sorted */
        var targetIndices = [];
        for (i = 0; i < artboardCount; i++) {
            if (isTarget[i]) targetIndices.push(i);
        }
        return targetIndices;
    }

    /**
     * 取り込むガイドを集める。横は基準のアートボードの短辺以下、縦は長辺以下の長さのものだけ含める（それより長いルーラーガイドは除く）。
     * ロック・非表示のガイドは選択できないので除く。withinRect を渡すと、その矩形に重なるものだけにする
     * @param {Document} sourceDoc - 読み込み元のドキュメント
     * @param {number[]} baseArtboardRect - 長さの基準にするアートボードの矩形 [左, 上, 右, 下]
     * @param {number[]|null} withinRect - 絞り込む矩形（null なら絞り込まない）
     * @returns {PageItem[]} ガイド
     */
    function collectImportableGuides(sourceDoc, baseArtboardRect, withinRect) {
        var artboardWidth = baseArtboardRect[2] - baseArtboardRect[0];
        var artboardHeight = baseArtboardRect[1] - baseArtboardRect[3];
        var shortSide = Math.min(artboardWidth, artboardHeight);
        var longSide = Math.max(artboardWidth, artboardHeight);
        var guideItems = [];
        for (var i = 0; i < sourceDoc.pageItems.length; i++) {
            var pageItem = sourceDoc.pageItems[i];
            if (!pageItem.guides || pageItem.locked || pageItem.hidden) continue;
            var guideBounds = pageItem.geometricBounds;
            var guideWidth = guideBounds[2] - guideBounds[0];
            var guideHeight = guideBounds[1] - guideBounds[3];
            /* 横長なら横のガイドとして短辺以下、縦長なら縦のガイドとして長辺以下なら含める。辺にぴったりの長さは誤差を許容して含める
               Include horizontal guides up to the short side and vertical ones up to the long side, with a tolerance for exact fits */
            var guideLimit = (guideWidth >= guideHeight) ? shortSide : longSide;
            if (Math.max(guideWidth, guideHeight) > guideLimit + SELECTION_ITEMS_TOLERANCE) continue;
            if (withinRect && !rectsIntersect(guideBounds, withinRect)) continue;
            guideItems.push(pageItem);
        }
        return guideItems;
    }

    /**
     * ドキュメントのすべてのアートボードを囲む矩形を返す
     * @param {Document} targetDoc - 対象のドキュメント
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getArtboardsUnionRect(targetDoc) {
        var unionRect = targetDoc.artboards[0].artboardRect.slice(0);
        for (var i = 1; i < targetDoc.artboards.length; i++) {
            var artboardRect = targetDoc.artboards[i].artboardRect;
            if (artboardRect[0] < unionRect[0]) unionRect[0] = artboardRect[0];
            if (artboardRect[1] > unionRect[1]) unionRect[1] = artboardRect[1];
            if (artboardRect[2] > unionRect[2]) unionRect[2] = artboardRect[2];
            if (artboardRect[3] < unionRect[3]) unionRect[3] = artboardRect[3];
        }
        return unionRect;
    }

    /**
     * 読み込み元の選択をコピーし、読み込み先へペーストする。
     * ペースト前に読み込み先の選択を解除する（選択中の文字などを置き換えないため）
     * @param {Document} sourceDoc - 読み込み元（選択済み）
     * @param {Document} destDoc - 読み込み先
     * @returns {void}
     */
    function copySelectionToDocument(sourceDoc, destDoc) {
        app.activeDocument = sourceDoc;
        app.copy();
        app.activeDocument = destDoc;
        destDoc.selection = null;
        app.paste();
    }

    /**
     * ペーストした直後のオブジェクトをグループにまとめてセルとして記録し、必要ならラベルを付ける。
     * ［コピー元のレイヤーにペースト］がオンで複数のレイヤーに入ったときは、レイヤーごとに1つずつグループにする。
     * placementContext.createArtboards が true のときだけ、セルに合わせたアートボードを追加する。
     * artboardCell を渡すと、元のアートボードの大きさと内容の相対位置を保つ
     * @param {Object} placementContext - 配置の状態（destDoc / cells / scalePercent / showLabel など）
     * @param {string} labelName - ラベル・アートボードの名前
     * @param {{width: number, height: number, offsetX: number, offsetY: number}} [artboardCell] - 元のアートボードの大きさと内容の位置
     * @returns {boolean} 配置できたら true
     */
    function placePastedGroup(placementContext, labelName, artboardCell) {
        var destDoc = placementContext.destDoc;
        var pastedItems = destDoc.selection;
        if (pastedItems.length === 0) {
            alert(labelValueText('alert.pasteFail', labelName));
            return false;
        }

        var activeLayer = destDoc.activeLayer;
        var pastedGroups = groupItemsByLayer(pastedItems);
        destDoc.activeLayer = activeLayer; /* ペースト直後にアクティブレイヤーを戻す / restore the active layer right after pasting */

        /* 内容とアートボード枠を同率で拡大・縮小する。線幅・パターン・グラデーションも同率。
           グループが複数なら、互いの位置がずれないよう同じ基準点（ドキュメントの原点）で拡大・縮小する
           Scale the content and the cell at the same rate, including strokes, patterns and gradients.
           Several groups share one reference point (the document origin) so their relative positions hold */
        var scalePercent = placementContext.scalePercent;
        if (scalePercent !== 100) {
            for (var i = 0; i < pastedGroups.length; i++) {
                if (pastedGroups.length > 1) {
                    pastedGroups[i].resize(scalePercent, scalePercent, true, true, true, true, scalePercent, Transformation.DOCUMENTORIGIN);
                } else {
                    pastedGroups[i].resize(scalePercent, scalePercent, true, true, true, true, scalePercent);
                }
            }
        }
        var scaleFactor = scalePercent / 100;

        /* クリップグループはマスクで測る（読み込み元と同じ測り方） / clip groups by their mask, the same way as the source side */
        var contentBounds = getClipAwareUnionBounds(pastedGroups, true) || pastedGroups[0].visibleBounds;

        /* セル（＝追加するアートボード）の大きさと、その中での内容の位置 / Cell size and the content's position within it */
        var cellPadding = placementContext.artboardPadding;
        var cellWidth, cellHeight, offsetX, offsetY;
        if (artboardCell) {
            cellWidth = artboardCell.width * scaleFactor;
            cellHeight = artboardCell.height * scaleFactor;
            offsetX = artboardCell.offsetX * scaleFactor;
            offsetY = artboardCell.offsetY * scaleFactor;
        } else {
            cellWidth = (contentBounds[2] - contentBounds[0]) + cellPadding * 2;
            cellHeight = (contentBounds[1] - contentBounds[3]) + cellPadding * 2;
            offsetX = cellPadding;
            offsetY = cellPadding;
        }

        /* 内容の今の位置から求めたセルの左上（最終位置はあとでグリッドに並べる） / Cell top-left from the content's current spot */
        var cellLeft = contentBounds[0] - offsetX;
        var cellTop = contentBounds[1] + offsetY;

        var newArtboard = null;
        if (placementContext.createArtboards) {
            newArtboard = destDoc.artboards.add([cellLeft, cellTop, cellLeft + cellWidth, cellTop - cellHeight]);
            newArtboard.name = labelName;
        }

        var labelFrame = placementContext.showLabel ? addFileNameLabel(destDoc, labelName, cellLeft, cellTop - cellHeight) : null;

        /* グループ・アートボード・ラベルをグリッド配置で一緒に動かすため記録する。アートボードが無いときのために今の左上も控える
           Record the cell so its parts move together; keep the current top-left for the no-artboard case */
        placementContext.cells.push({
            groups: pastedGroups,
            artboard: newArtboard,
            label: labelFrame,
            width: cellWidth,
            height: cellHeight,
            currentLeft: cellLeft,
            currentTop: cellTop
        });
        placementContext.placedCount++;
        return true;
    }

    /**
     * オブジェクトを、入っているレイヤーごとに1つのグループにまとめる（重なり順は保つ）。
     * ［コピー元のレイヤーにペースト］がオンだと、ペーストしたものは元と同じ名前のレイヤー（サブレイヤーも）に分かれて入る
     * @param {PageItem[]} items - ペーストしたオブジェクト（前面→背面の順）
     * @returns {GroupItem[]} レイヤーごとのグループ
     */
    function groupItemsByLayer(items) {
        var layerGroups = [];
        for (var i = items.length - 1; i >= 0; i--) {
            var itemLayer = items[i].layer;
            var layerGroup = null;
            for (var j = 0; j < layerGroups.length; j++) {
                if (layerGroups[j].layer === itemLayer) layerGroup = layerGroups[j];
            }
            if (!layerGroup) {
                layerGroup = itemLayer.groupItems.add();
                layerGroups.push(layerGroup);
            }
            /* 背面から順に先頭へ入れるので、元の重なり順になる / moving back-to-front to the beginning keeps the stacking order */
            items[i].moveToBeginning(layerGroup);
        }
        return layerGroups;
    }

    /**
     * ファイル名のラベルを、ラベル用のレイヤーのセル左下に追加する
     * @param {Document} destDoc - 読み込み先のドキュメント
     * @param {string} labelName - ラベルの文字列
     * @param {number} cellLeft - セルの左端
     * @param {number} cellBottom - セルの下端
     * @returns {TextFrame} ラベル
     */
    function addFileNameLabel(destDoc, labelName, cellLeft, cellBottom) {
        var labelFrame = getOrCreateLabelLayer(destDoc).textFrames.add();
        labelFrame.contents = labelName;
        try {
            labelFrame.textRange.characterAttributes.textFont = app.textFonts.getByName(IMPORT_SETTINGS.labelFont);
        } catch (e) { /* フォントが無い環境では既定のフォントのまま / keep the default font when it is missing */ }
        labelFrame.textRange.characterAttributes.size = IMPORT_SETTINGS.labelSize;
        labelFrame.left = cellLeft;
        labelFrame.top = cellBottom - 4; /* アートボードの下端のすぐ下 / just below the artboard */
        return labelFrame;
    }

    /**
     * 全体が正方形に近くなる列数を選ぶ
     * @param {number} cellCount - セルの数
     * @param {number} slotWidth - 1マスの幅
     * @param {number} slotHeight - 1マスの高さ
     * @param {number} gapX - 横の間隔
     * @param {number} gapY - 縦の間隔
     * @returns {number} 列数
     */
    function chooseSquarestColumnCount(cellCount, slotWidth, slotHeight, gapX, gapY) {
        var bestColumns = 1;
        var bestDiff = -1;
        for (var columnCount = 1; columnCount <= cellCount; columnCount++) {
            var rowCount = Math.ceil(cellCount / columnCount);
            var gridWidth = columnCount * slotWidth + (columnCount - 1) * gapX;
            var gridHeight = rowCount * slotHeight + (rowCount - 1) * gapY;
            var diff = Math.abs(gridWidth - gridHeight);
            if (bestDiff < 0 || diff < bestDiff) {
                bestDiff = diff;
                bestColumns = columnCount;
            }
        }
        return bestColumns;
    }

    /**
     * 記録したセルを、全体が正方形に近いグリッドに並べる。
     * 新規ドキュメントはカンバス中央、現在のドキュメントは既存のアートボードの下に置く。
     * セルごとにグループ・アートボード・ラベルを同じ量だけ動かす
     * @param {Object} placementContext - 配置の状態
     * @returns {void}
     */
    function layoutCellsAsCenteredGrid(placementContext) {
        var cells = placementContext.cells;
        var cellCount = cells.length;
        if (cellCount === 0) return;

        /* マスの大きさ＝最大のセル（大きさが混ざっても重ならない） / slot = the largest cell, so mixed sizes never overlap */
        var slotWidth = 0;
        var slotHeight = 0;
        for (var i = 0; i < cellCount; i++) {
            if (cells[i].width > slotWidth) slotWidth = cells[i].width;
            if (cells[i].height > slotHeight) slotHeight = cells[i].height;
        }

        /* ［間隔］の値を縦横に使う / the Spacing value is used both ways */
        var gapX = placementContext.itemSpacing;
        var gapY = placementContext.itemSpacing;

        var columnCount = chooseSquarestColumnCount(cellCount, slotWidth, slotHeight, gapX, gapY);
        var rowCount = Math.ceil(cellCount / columnCount);
        var gridWidth = columnCount * slotWidth + (columnCount - 1) * gapX;
        var gridHeight = rowCount * slotHeight + (rowCount - 1) * gapY;

        var startLeft, startTop;
        var avoidRect = placementContext.avoidRect;
        if (avoidRect) {
            /* 既存のアートボードの下にひと間隔あけて置き、横は既存の中央にそろえる / below the existing artboards, centered on them */
            startLeft = (avoidRect[0] + avoidRect[2]) / 2 - gridWidth / 2;
            startTop = avoidRect[3] - gapY;
        } else {
            startLeft = placementContext.canvasCenterX - gridWidth / 2;
            startTop = placementContext.canvasCenterY + gridHeight / 2;
        }

        for (var k = 0; k < cellCount; k++) {
            var cell = cells[k];
            var slotLeft = startLeft + (k % columnCount) * (slotWidth + gapX);
            var slotTop = startTop - Math.floor(k / columnCount) * (slotHeight + gapY);
            /* セルをマスの中央に / center the cell in its slot */
            var targetLeft = slotLeft + (slotWidth - cell.width) / 2;
            var targetTop = slotTop - (slotHeight - cell.height) / 2;

            /* アートボードがあればその枠、無ければ控えたセルの左上を基準に動かす / measure from the artboard, or the recorded top-left */
            var cellRect = cell.artboard ?
                cell.artboard.artboardRect :
                [cell.currentLeft, cell.currentTop, cell.currentLeft + cell.width, cell.currentTop - cell.height];
            var dx = targetLeft - cellRect[0];
            var dy = targetTop - cellRect[1];

            for (var j = 0; j < cell.groups.length; j++) cell.groups[j].translate(dx, dy);
            if (cell.artboard) cell.artboard.artboardRect = [cellRect[0] + dx, cellRect[1] + dy, cellRect[2] + dx, cellRect[3] + dy];
            if (cell.label) cell.label.translate(dx, dy);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 読み込み対象パネル（開いているドキュメント／フォルダー、フィルター）を作る
     * @param {Window} dialog - ダイアログ
     * @param {Document[]} openDocs - 開いているドキュメント
     * @returns {Object} sourceUi（isFolderMode() / getSources() / getNameFilterRegExp() / filterInput と、モード切替時に呼ぶ onModeChange）
     */
    function buildSourcePanel(dialog, openDocs) {
        var sourceUi = { onModeChange: null };
        var selectedFolder = null;
        var folderFiles = [];      /* 種類と名前で絞り込んだファイル / files after the type and name filters */
        var folderFilesTotal = 0;  /* 種類に合うファイルの総数（名前で絞る前） / total matching the types, before the name filter */

        var sourcePanel = dialog.add("panel", undefined, getLabel('panel.source'));
        setupPanel(sourcePanel, 6);
        var openDocsRadio = sourcePanel.add("radiobutton", undefined, getLabel('radio.openDocs'));
        openDocsRadio.helpTip = getLabel('tooltip.openDocs');
        reserveCountWidth(openDocsRadio);

        var folderRow = sourcePanel.add("group");
        setupGroup(folderRow, "row");
        var folderRadio = folderRow.add("radiobutton", undefined, getLabel('radio.specifyFolder'));
        folderRadio.helpTip = getLabel('tooltip.specifyFolder');
        reserveCountWidth(folderRadio);
        var btnChooseFolder = folderRow.add("button", undefined, getLabel('button.chooseFolder'));
        btnChooseFolder.helpTip = getLabel('tooltip.chooseFolder');

        /* 選んだフォルダーのパスは別の行に出す / the chosen folder path goes on its own row */
        var folderPathRow = sourcePanel.add("group");
        setupGroup(folderPathRow, "row");
        var folderPathText = folderPathRow.add("statictext", undefined, "", { truncate: "middle" });
        folderPathText.preferredSize = [360, 20];
        folderPathText.minimumSize = [360, 20];

        var subfoldersCheckbox = sourcePanel.add("checkbox", undefined, getLabel('checkbox.includeSubfolders'));
        subfoldersCheckbox.helpTip = getLabel('tooltip.includeSubfolders');
        subfoldersCheckbox.onClick = refreshFolderFiles;

        /* フィルター（種類＋ファイル名の正規表現）。読み込み対象パネルの中に入れる / Filter panel nested in the source panel */
        var filterPanel = sourcePanel.add("panel", undefined, getLabel('panel.filter'));
        setupPanel(filterPanel, 6);

        var fileTypeRow = filterPanel.add("group");
        setupGroup(fileTypeRow, "row");
        var fileTypeLabel = addRowLabel(fileTypeRow, 'fieldLabel.fileType', FILTER_LABEL_WIDTH);
        var fileTypeCheckboxes = [];
        var fileTypeExtensions = ["ai", "svg", "eps"];
        for (var i = 0; i < fileTypeExtensions.length; i++) {
            var fileTypeCheckbox = fileTypeRow.add("checkbox", undefined, fileTypeExtensions[i].toUpperCase());
            fileTypeCheckbox.helpTip = getLabel('tooltip.fileType');
            fileTypeCheckbox.value = (i === 0); /* 既定は AI だけ / AI only by default */
            fileTypeCheckbox.onClick = refreshFolderFiles;
            fileTypeCheckboxes.push(fileTypeCheckbox);
        }

        var fileNameRow = filterPanel.add("group");
        setupGroup(fileNameRow, "row");
        var fileNameLabel = addRowLabel(fileNameRow, 'fieldLabel.fileName', FILTER_LABEL_WIDTH);
        fileNameLabel.helpTip = getLabel('tooltip.fileNameFilter');
        var filterInput = fileNameRow.add("edittext", undefined, "");
        filterInput.characters = 20;
        filterInput.helpTip = getLabel('tooltip.fileNameFilter');
        filterInput.onChanging = refreshFolderFiles; /* 件数も更新する / also refreshes the count */

        /**
         * ファイル名フィルターの正規表現を返す
         * @returns {RegExp|null} 空欄や正しくない式（入力途中など）なら null
         */
        function getNameFilterRegExp() {
            if (filterInput.text === "") return null;
            try {
                return new RegExp(filterInput.text, "i");
            } catch (e) {
                return null; /* 入力途中の正しくない式では絞り込まない / do not filter on an incomplete pattern */
            }
        }

        /**
         * 名前がフィルターに合うものだけを残す
         * @param {Array} candidates - name を持つ Document か File
         * @returns {Array} 残したもの
         */
        function filterByName(candidates) {
            var nameRegExp = getNameFilterRegExp();
            if (!nameRegExp) return candidates.slice(0);
            var matched = [];
            for (var i = 0; i < candidates.length; i++) {
                if (nameRegExp.test(candidates[i].name)) matched.push(candidates[i]);
            }
            return matched;
        }

        /**
         * ファイル数を各ラジオボタンの後ろに出す。フィルターが効いていれば「開いているドキュメント（3/5）」、無ければ「（5）」。
         * フォルダーは選んだあとだけ出す
         * @returns {void}
         */
        function updateFileCount() {
            var isFiltering = (getNameFilterRegExp() !== null);
            openDocsRadio.text = getLabel('radio.openDocs') + formatFileCount(isFiltering ? filterByName(openDocs).length : null, openDocs.length);
            folderRadio.text = getLabel('radio.specifyFolder') +
                (selectedFolder ? formatFileCount(isFiltering ? folderFiles.length : null, folderFilesTotal) : "");
        }

        /**
         * 選んだフォルダーのファイルを、今の種類とファイル名で絞り直し、件数を更新する
         * @returns {void}
         */
        function refreshFolderFiles() {
            folderFiles = [];
            folderFilesTotal = 0;
            var extensions = [];
            for (var i = 0; i < fileTypeCheckboxes.length; i++) {
                if (fileTypeCheckboxes[i].value) extensions.push(fileTypeExtensions[i]);
            }
            if (selectedFolder && extensions.length > 0) {
                var extensionRegExp = new RegExp("\\.(" + extensions.join("|") + ")$", "i");
                var typeMatchedFiles = collectFolderFiles(selectedFolder, extensionRegExp, subfoldersCheckbox.value);
                folderFilesTotal = typeMatchedFiles.length;
                folderFiles = sortFilesByPath(filterByName(typeMatchedFiles));
            }
            updateFileCount();
        }

        /**
         * 読み込み対象のモードを切り替える（2つのラジオは別の親にあるので排他を手で保つ）
         * @param {boolean} useFolder - フォルダーを指定するなら true
         * @returns {void}
         */
        function setFolderMode(useFolder) {
            folderRadio.value = useFolder;
            openDocsRadio.value = !useFolder;
            fileTypeLabel.enabled = useFolder; /* 種類とサブフォルダーはフォルダー指定のときだけ / file types and subfolders only apply to folders */
            subfoldersCheckbox.enabled = useFolder;
            for (var i = 0; i < fileTypeCheckboxes.length; i++) fileTypeCheckboxes[i].enabled = useFolder;
            updateFileCount();
            if (sourceUi.onModeChange) sourceUi.onModeChange(useFolder);
        }

        openDocsRadio.onClick = function () { setFolderMode(false); };
        folderRadio.onClick = function () { setFolderMode(true); };
        btnChooseFolder.onClick = function () {
            var chosenFolder = Folder.selectDialog(getLabel('prompt.selectFolder'));
            if (!chosenFolder) return;
            selectedFolder = chosenFolder;
            var folderPath = decodeURI(chosenFolder.fsName);
            folderPathText.text = folderPath;
            folderPathText.helpTip = folderPath;
            refreshFolderFiles();
            setFolderMode(true);
        };

        /* 開いているドキュメントが無ければフォルダー指定にする / fall back to folder mode when nothing is open */
        openDocsRadio.enabled = openDocs.length > 0;

        sourceUi.filterInput = filterInput;
        sourceUi.getNameFilterRegExp = getNameFilterRegExp;
        sourceUi.setFolderMode = setFolderMode;
        sourceUi.isFolderMode = function () { return folderRadio.value; };
        sourceUi.getSources = function () { return folderRadio.value ? folderFiles : filterByName(openDocs); };
        return sourceUi;
    }

    var SUBFOLDER_MAX_DEPTH = 20; /* サブフォルダーをたどる深さの上限（エイリアスの循環よけ） / depth limit, guards against alias loops */

    /**
     * フォルダー内の、拡張子が合うファイルを集める（includeSubfolders ならサブフォルダーもたどる。名前が「.」で始まるものは除く）
     * @param {Folder} rootFolder - フォルダー
     * @param {RegExp} extensionRegExp - 拡張子の条件
     * @param {boolean} includeSubfolders - サブフォルダーも含めるなら true
     * @returns {File[]} ファイル
     */
    function collectFolderFiles(rootFolder, extensionRegExp, includeSubfolders) {
        var matchedFiles = [];
        function visitFolder(currentFolder, depth) {
            var entries = currentFolder.getFiles();
            for (var i = 0; i < entries.length; i++) {
                var entry = entries[i];
                if (entry.name.charAt(0) === ".") continue;
                if (entry instanceof File) {
                    if (extensionRegExp.test(entry.name)) matchedFiles.push(entry);
                } else if (includeSubfolders && depth < SUBFOLDER_MAX_DEPTH) {
                    visitFolder(entry, depth + 1);
                }
            }
        }
        visitFolder(rootFolder, 0);
        return matchedFiles;
    }

    /**
     * 件数の表示を返す（「（3/5）」または「（5）」。英語は半角括弧）
     * @param {number|null} filteredCount - 絞り込んだ件数（絞り込んでいなければ null）
     * @param {number} totalCount - 総数
     * @returns {string} 件数の表示
     */
    function formatFileCount(filteredCount, totalCount) {
        var countText = (filteredCount === null) ? String(totalCount) : filteredCount + "/" + totalCount;
        return (uiLang === "ja") ? "（" + countText + "）" : " (" + countText + ")";
    }

    /**
     * 件数を後ろに付けても切れないよう、ラジオボタンの幅を先に取っておく（表示後は広がらない）
     * @param {RadioButton} radioButton - 対象のラジオボタン
     * @returns {void}
     */
    function reserveCountWidth(radioButton) {
        var RADIO_INDICATOR_WIDTH = 24; /* ラジオの丸と間の幅 / width of the radio circle and gap */
        var widestText = radioButton.text + formatFileCount(9999, 9999);
        radioButton.preferredSize.width = Math.ceil(radioButton.graphics.measureString(widestText)[0]) + RADIO_INDICATOR_WIDTH;
    }

    /**
     * ファイルをパス（大文字・小文字を区別しない）の順に並べる。同じフォルダーの中では名前順になる。
     * 比較関数つきの sort() は遅く並びも狂うので、文字列のキーを引数なしの sort() で並べる
     * @param {File[]} files - ファイル
     * @returns {File[]} 並べた新しい配列
     */
    function sortFilesByPath(files) {
        var sortKeys = [];
        for (var i = 0; i < files.length; i++) sortKeys.push(files[i].fsName.toLowerCase() + "\u0000" + i);
        sortKeys.sort();
        var sortedFiles = [];
        for (var j = 0; j < sortKeys.length; j++) {
            sortedFiles.push(files[parseInt(sortKeys[j].substring(sortKeys[j].lastIndexOf("\u0000") + 1), 10)]);
        }
        return sortedFiles;
    }

    /* ラスタライズ効果の解像度の選択肢と、DocumentPreset に渡す値 / raster effects resolution choices and their DocumentPreset values */
    var RESOLUTION_PPI_LIST = [72, 150, 300];
    var RESOLUTION_PRESET_VALUES = [
        DocumentRasterResolution.ScreenResolution,
        DocumentRasterResolution.MediumResolution,
        DocumentRasterResolution.HighResolution
    ];

    /**
     * 新規ドキュメントのプロファイル名の一覧を返す（［新規ドキュメント］のプリント・Web など。ユーザー定義も含む）
     * @returns {string[]} プロファイル名
     */
    function getDocumentProfileNames() {
        var profileNames = [];
        for (var i = 0; i < app.startupPresetsList.length; i++) profileNames.push(app.startupPresetsList[i]);
        return profileNames;
    }

    /**
     * 既定のプロファイル名を返す（プリントがあればそれ、無ければ先頭）
     * @returns {string} プロファイル名
     */
    function getDefaultProfileName() {
        var profileNames = getDocumentProfileNames();
        for (var i = 0; i < profileNames.length; i++) {
            if (profileNames[i] === "プリント" || profileNames[i] === "Print") return profileNames[i];
        }
        return profileNames[0];
    }

    /**
     * 新規ドキュメントの設定の初期値を返す
     * @returns {{profileName: string, presetIndex: number, isRgb: boolean, ppiIndex: number, widthText: string, heightText: string, unit: string}} 設定
     */
    function createDefaultNewDocSettings() {
        return { profileName: getDefaultProfileName(), presetIndex: 0, isRgb: true, ppiIndex: 2, widthText: "1000", heightText: "1000", unit: "px" };
    }

    /**
     * 新規ドキュメントの設定を1行の要約にする（例「プリント / RGB / 300 ppi / 1000 × 1000 px」）
     * @param {Object} newDocSettings - 新規ドキュメントの設定
     * @returns {string} 要約
     */
    function formatNewDocSummary(newDocSettings) {
        return newDocSettings.profileName + " / " + (newDocSettings.isRgb ? "RGB" : "CMYK") + " / " +
            RESOLUTION_PPI_LIST[newDocSettings.ppiIndex] + " ppi / " +
            newDocSettings.widthText + " × " + newDocSettings.heightText + " " + newDocSettings.unit;
    }

    /**
     * 読み込み先パネル（現在／新規と、新規ドキュメントの設定を開くボタン）を作る
     * @param {Window} dialog - ダイアログ
     * @returns {Object} destinationUi（currentDocRadio と、新規ドキュメントの設定を返す getNewDocSettings()）
     */
    function buildDestinationPanel(dialog) {
        var newDocSettings = createDefaultNewDocSettings();

        var destinationPanel = dialog.add("panel", undefined, getLabel('panel.destination'));
        setupPanel(destinationPanel, 6);

        /* 縦に並べ、［新規ドキュメント］の右に［設定...］を置く / stacked, with Settings... next to New document */
        var currentDocRadio = destinationPanel.add("radiobutton", undefined, getLabel('radio.currentDoc'));
        currentDocRadio.helpTip = getLabel('tooltip.currentDoc');
        var newDocRow = destinationPanel.add("group");
        setupGroup(newDocRow, "row");
        var newDocRadio = newDocRow.add("radiobutton", undefined, getLabel('radio.newDoc'));
        newDocRadio.helpTip = getLabel('tooltip.newDoc');
        newDocRadio.value = true;
        var btnNewDocSettings = newDocRow.add("button", undefined, getLabel('button.newDocSettings'));
        btnNewDocSettings.helpTip = getLabel('tooltip.newDocSettings');

        /* 新規ドキュメントの設定の要約。幅は表示後に広がらないので先に取っておく / summary; reserve the width up front */
        var newDocSummaryText = destinationPanel.add("statictext", undefined, formatNewDocSummary(newDocSettings));
        newDocSummaryText.preferredSize.width = 300;

        /* 「□分割［20］ファイルごと」：新規ドキュメントを指定のファイル数ごとに分ける / split new documents every N files */
        var splitRow = destinationPanel.add("group");
        setupGroup(splitRow, "row");
        var splitCheckbox = splitRow.add("checkbox", undefined, getLabel('checkbox.splitDocs'));
        splitCheckbox.helpTip = getLabel('tooltip.splitDocs');
        var splitCountInput = addSteppedInput(splitRow, String(IMPORT_SETTINGS.splitFileCount), { min: 1, integer: true });
        splitCountInput.characters = 4;
        splitCountInput.helpTip = getLabel('tooltip.splitDocs');
        splitRow.add("statictext", undefined, getLabel('fieldLabel.splitUnit'));
        function updateSplitState() {
            /* ∧∨は入力欄の兄弟なので、包む group ごと切り替える / toggle the wrapper so the stepper dims too */
            splitCountInput.parent.enabled = splitCheckbox.value;
            splitCountInput.enabled = splitCheckbox.value;
            redrawSteppersIn(splitRow);
        }
        splitCheckbox.onClick = updateSplitState;
        updateSplitState();

        /* 現在のドキュメントに読み込むときは新規ドキュメントの設定をディムにする / dim the new-document settings for the current document */
        function updateDestinationState() {
            btnNewDocSettings.enabled = newDocRadio.value;
            newDocSummaryText.enabled = newDocRadio.value;
            splitRow.enabled = newDocRadio.value;
            redrawSteppersIn(splitRow);
        }
        /* 2つのラジオは別の親にあるので、排他を手で保つ / the radios have different parents, so keep them exclusive by hand */
        currentDocRadio.onClick = function () {
            newDocRadio.value = false;
            updateDestinationState();
        };
        newDocRadio.onClick = function () {
            currentDocRadio.value = false;
            updateDestinationState();
        };

        btnNewDocSettings.onClick = function () {
            var editedSettings = showNewDocSettingsDialog(newDocSettings);
            if (!editedSettings) return;
            newDocSettings = editedSettings;
            newDocSummaryText.text = formatNewDocSummary(newDocSettings);
        };

        return {
            currentDocRadio: currentDocRadio,
            getNewDocSettings: function () { return newDocSettings; },
            /* 分割するファイル数（分けないなら 0） / files per document, 0 when not splitting */
            getSplitFileCount: function () {
                return splitCheckbox.value ? parseInt(splitCountInput.text, 10) : 0;
            }
        };
    }

    /**
     * 新規ドキュメントの設定（カラーモード・解像度・サイズ）のダイアログボックスを表示する
     * @param {Object} initialSettings - 今の設定（書き換えない）
     * @returns {Object|null} OK なら新しい設定、キャンセルなら null
     */
    function showNewDocSettingsDialog(initialSettings) {
        var settingsDialog = new Window("dialog", getLabel('dialog.newDocSettings'));
        setupWindow(settingsDialog);

        /* プロファイル（スウォッチ・ブラシなどの初期内容が決まる） / profile, which sets the starting swatches, brushes, etc. */
        var profileRow = settingsDialog.add("group");
        setupGroup(profileRow, "row");
        profileRow.alignment = ["center", "top"]; /* 左右中央に置く / center horizontally */
        profileRow.add("statictext", undefined, labelText('fieldLabel.profile'));
        var profileNames = getDocumentProfileNames();
        var profileDropdown = profileRow.add("dropdownlist", undefined, profileNames);
        profileDropdown.helpTip = getLabel('tooltip.profile');
        for (var k = 0; k < profileNames.length; k++) {
            if (profileNames[k] === initialSettings.profileName) profileDropdown.selection = k;
        }
        if (!profileDropdown.selection) profileDropdown.selection = 0;

        /* カラーモード＋解像度とサイズを2カラムで並べる / color mode + resolution and size in two columns */
        var settingsColumns = settingsDialog.add("group");
        setupGroup(settingsColumns, "row", 15);
        settingsColumns.alignChildren = ["left", "fill"]; /* 2カラムの高さをそろえる / match the columns' heights */

        var colorAndResolutionColumn = settingsColumns.add("group");
        setupGroup(colorAndResolutionColumn, "column");
        colorAndResolutionColumn.alignChildren = ["fill", "top"];

        var colorModePanel = colorAndResolutionColumn.add("panel", undefined, getLabel('panel.colorMode'));
        setupPanel(colorModePanel, 6);
        var rgbRadio = colorModePanel.add("radiobutton", undefined, "RGB");
        rgbRadio.helpTip = getLabel('tooltip.colorMode');
        var cmykRadio = colorModePanel.add("radiobutton", undefined, "CMYK");
        cmykRadio.helpTip = getLabel('tooltip.colorMode');
        rgbRadio.value = initialSettings.isRgb;
        cmykRadio.value = !initialSettings.isRgb;

        /* ラスタライズ効果の解像度 / raster effects resolution */
        var resolutionPanel = colorAndResolutionColumn.add("panel", undefined, getLabel('panel.resolution'));
        setupPanel(resolutionPanel, 6);
        var resolutionNames = [];
        for (var i = 0; i < RESOLUTION_PPI_LIST.length; i++) resolutionNames.push(RESOLUTION_PPI_LIST[i] + " ppi");
        var resolutionDropdown = resolutionPanel.add("dropdownlist", undefined, resolutionNames);
        resolutionDropdown.selection = initialSettings.ppiIndex;
        resolutionDropdown.helpTip = getLabel('tooltip.resolution');

        var sizePanel = settingsColumns.add("panel", undefined, getLabel('panel.docSize'));
        setupPanel(sizePanel, 6);
        var presetNames = [];
        for (var j = 0; j < SIZE_PRESETS.length; j++) presetNames.push(getLabel(SIZE_PRESETS[j].labelKey));
        var presetDropdown = sizePanel.add("dropdownlist", undefined, presetNames);
        presetDropdown.selection = initialSettings.presetIndex; /* onChange を付ける前に選ぶ / select before onChange is attached */
        presetDropdown.helpTip = getLabel('tooltip.sizePreset');

        var currentUnit = initialSettings.unit;

        var widthRow = sizePanel.add("group");
        setupGroup(widthRow, "row");
        addRowLabel(widthRow, 'fieldLabel.width', FIELD_LABEL_WIDTH);
        var widthInput = addSteppedInput(widthRow, initialSettings.widthText, { min: 1 });
        widthInput.characters = 5;

        var heightRow = sizePanel.add("group");
        setupGroup(heightRow, "row");
        addRowLabel(heightRow, 'fieldLabel.height', FIELD_LABEL_WIDTH);
        var heightInput = addSteppedInput(heightRow, initialSettings.heightText, { min: 1 });
        heightInput.characters = 5;

        var unitRow = sizePanel.add("group");
        setupGroup(unitRow, "row");
        addRowLabel(unitRow, 'fieldLabel.unit', FIELD_LABEL_WIDTH);
        var unitDropdown = unitRow.add("dropdownlist", undefined, SIZE_UNITS);
        unitDropdown.helpTip = getLabel('tooltip.unit');

        /**
         * 幅・高さを別の単位へ換算して書き換える（物理的な大きさを保つ）
         * @param {number|string} widthValue - 幅
         * @param {number|string} heightValue - 高さ
         * @param {string} fromUnit - 元の単位
         * @param {string} toUnit - 換算先の単位
         * @returns {void}
         */
        function writeConvertedSize(widthValue, heightValue, fromUnit, toUnit) {
            widthInput.text = convertSizeText(widthValue, fromUnit, toUnit);
            heightInput.text = convertSizeText(heightValue, fromUnit, toUnit);
            currentUnit = toUnit;
        }

        /* プログラムからの選択で onChange を動かさないための目印 / suppresses onChange for programmatic selection */
        var isSelectingUnit = false;
        function selectUnit(unitName) {
            isSelectingUnit = true;
            for (var i = 0; i < SIZE_UNITS.length; i++) {
                if (SIZE_UNITS[i] === unitName) unitDropdown.selection = i;
            }
            isSelectingUnit = false;
        }
        selectUnit(currentUnit);

        unitDropdown.onChange = function () {
            if (isSelectingUnit || !unitDropdown.selection) return;
            var newUnit = unitDropdown.selection.text;
            if (newUnit === currentUnit) return;
            writeConvertedSize(widthInput.text, heightInput.text, currentUnit, newUnit);
        };

        presetDropdown.onChange = function () {
            var sizePreset = SIZE_PRESETS[presetDropdown.selection.index];
            if (sizePreset.width !== undefined) {
                /* プリセットの値を表示の単位へ換算 / convert the preset's native values */
                writeConvertedSize(sizePreset.width, sizePreset.height, sizePreset.unit, sizePreset.displayUnit);
            } else {
                /* カスタム：今の値を換算して大きさを保つ / Custom: keep the physical size */
                writeConvertedSize(widthInput.text, heightInput.text, currentUnit, sizePreset.displayUnit);
            }
            selectUnit(currentUnit);
            /* 印刷向けは CMYK、画面向けは RGB / CMYK for print presets, RGB for screen */
            cmykRadio.value = !!sizePreset.isPrint;
            rgbRadio.value = !sizePreset.isPrint;
        };

        var buttonRow = addButtonRow(settingsDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(settingsDialog, SCRIPT_NAME + "_NewDocSettings");
        if (settingsDialog.show() !== 1) return null;

        return {
            profileName: profileDropdown.selection.text,
            presetIndex: presetDropdown.selection.index,
            isRgb: rgbRadio.value,
            ppiIndex: resolutionDropdown.selection.index,
            widthText: widthInput.text,
            heightText: heightInput.text,
            unit: currentUnit
        };
    }

    /**
     * 大きさの値を別の単位へ換算した文字列にする（整数に丸める）
     * @param {string|number} sizeValue - 値
     * @param {string} fromUnit - 元の単位（SIZE_UNIT_POINTS のキー）
     * @param {string} toUnit - 換算先の単位
     * @returns {string} 換算した値（数値でなければそのまま）
     */
    function convertSizeText(sizeValue, fromUnit, toUnit) {
        var numericValue = parseFloat(sizeValue);
        if (isNaN(numericValue)) return String(sizeValue);
        return String(Math.round(numericValue * SIZE_UNIT_POINTS[fromUnit] / SIZE_UNIT_POINTS[toUnit]));
    }

    /**
     * 読み込みオプションパネルを作る
     * @param {Window} dialog - ダイアログ
     * @returns {Object} optionsUi（各コントロールと、読み込み後の動作を有効にする setAfterImportEnabled()）
     */
    function buildOptionsPanel(dialog) {
        var optionsPanel = dialog.add("panel", undefined, getLabel('panel.importOptions'));
        setupPanel(optionsPanel, 6);
        /* 左：アートボード単位・ラベル・ガイド・拡大・縮小、右：対象アートボード / left: options, right: target artboards */
        var optionsColumns = optionsPanel.add("group");
        setupGroup(optionsColumns, "row", 20);
        optionsColumns.alignChildren = ["left", "top"];

        var optionsLeftColumn = optionsColumns.add("group");
        setupGroup(optionsLeftColumn, "column");
        var byArtboardCheckbox = optionsLeftColumn.add("checkbox", undefined, getLabel('checkbox.byArtboard'));
        byArtboardCheckbox.value = false;
        byArtboardCheckbox.helpTip = getLabel('tooltip.byArtboard');

        var fileNameLabelCheckbox = optionsLeftColumn.add("checkbox", undefined, getLabel('checkbox.attachLabel'));
        fileNameLabelCheckbox.helpTip = getLabel('tooltip.attachLabel');

        var includeGuidesCheckbox = optionsLeftColumn.add("checkbox", undefined, getLabel('checkbox.includeGuides'));
        includeGuidesCheckbox.value = true;
        includeGuidesCheckbox.helpTip = getLabel('tooltip.includeGuides');

        var scaleRow = optionsLeftColumn.add("group");
        setupGroup(scaleRow, "row");
        var scaleCheckbox = scaleRow.add("checkbox", undefined, getLabel('checkbox.scale'));
        scaleCheckbox.value = false;
        scaleCheckbox.helpTip = getLabel('tooltip.scale');
        var scaleInput = addSteppedInput(scaleRow, "100", { min: 1 });
        scaleInput.helpTip = getLabel('tooltip.scale');
        scaleInput.characters = 4;
        scaleRow.add("statictext", undefined, "%");
        /* ∧∨は入力欄の兄弟なので、包む group ごと切り替えてディム表示にする / toggle the wrapper so the stepper dims too */
        function updateScaleState() {
            scaleInput.parent.enabled = scaleCheckbox.value;
            scaleInput.enabled = scaleCheckbox.value;
            redrawSteppersIn(scaleInput.parent);
        }
        scaleCheckbox.onClick = updateScaleState;
        updateScaleState();

        /* アイテムの間隔（定規の単位） / gap between items, in ruler units */
        var rulerUnit = getUnitInfo();
        var spacingRow = optionsLeftColumn.add("group");
        setupGroup(spacingRow, "row");
        var spacingLabel = spacingRow.add("statictext", undefined, labelText('fieldLabel.itemSpacing'));
        spacingLabel.helpTip = getLabel('tooltip.itemSpacing');
        var spacingStepOptions = { min: 0, unit: " " + rulerUnit.label };
        var spacingInput = addSteppedInput(spacingRow, formatSteppedValue(IMPORT_SETTINGS.itemSpacing / rulerUnit.pointsPerUnit, spacingStepOptions), spacingStepOptions);
        spacingInput.characters = 7;
        spacingInput.helpTip = getLabel('tooltip.itemSpacing');

        var targetArtboardPanel = optionsColumns.add("panel", undefined, getLabel('panel.targetArtboards'));
        setupPanel(targetArtboardPanel, 6);
        var artboardFirstRadio = targetArtboardPanel.add("radiobutton", undefined, getLabel('radio.artboardFirst'));
        artboardFirstRadio.helpTip = getLabel('tooltip.targetArtboards');
        var artboardAllRadio = targetArtboardPanel.add("radiobutton", undefined, getLabel('radio.artboardAll'));
        artboardAllRadio.helpTip = getLabel('tooltip.targetArtboards');
        var artboardSpecRow = targetArtboardPanel.add("group");
        setupGroup(artboardSpecRow, "row");
        var artboardSpecRadio = artboardSpecRow.add("radiobutton", undefined, getLabel('radio.artboardSpecify'));
        artboardSpecRadio.helpTip = getLabel('tooltip.artboardSpecify');
        var artboardSpecInput = artboardSpecRow.add("edittext", undefined, "");
        artboardSpecInput.characters = 7;
        artboardSpecInput.helpTip = getLabel('tooltip.artboardSpecify');

        /* 3つのラジオは別の親にまたがるので、排他を手で保つ / the radios span containers, so keep them exclusive by hand */
        function selectArtboardTarget(targetMode) {
            artboardFirstRadio.value = (targetMode === "first");
            artboardAllRadio.value = (targetMode === "all");
            artboardSpecRadio.value = (targetMode === "specify");
            artboardSpecInput.enabled = artboardSpecRadio.value;
        }
        artboardFirstRadio.onClick = function () { selectArtboardTarget("first"); };
        artboardAllRadio.onClick = function () { selectArtboardTarget("all"); };
        artboardSpecRadio.onClick = function () { selectArtboardTarget("specify"); };
        selectArtboardTarget("first");

        /* 対象アートボードはアートボード単位のときだけ / target artboards apply only per artboard */
        byArtboardCheckbox.onClick = function () { targetArtboardPanel.enabled = byArtboardCheckbox.value; };
        byArtboardCheckbox.onClick();

        /* 読み込み後の元ドキュメント（開いているドキュメントのときだけ有効） / after-import action, open documents only */
        var afterImportRow = optionsPanel.add("group");
        setupGroup(afterImportRow, "row");
        afterImportRow.add("statictext", undefined, labelText('fieldLabel.afterImport'));
        var closeDocRadio = afterImportRow.add("radiobutton", undefined, getLabel('radio.closeDoc'));
        closeDocRadio.helpTip = getLabel('tooltip.closeDoc');
        var keepOpenRadio = afterImportRow.add("radiobutton", undefined, getLabel('radio.keepOpen'));
        keepOpenRadio.helpTip = getLabel('tooltip.keepOpen');
        keepOpenRadio.value = true; /* 既定は開いたまま（未保存の変更を失わない） / keep open by default so unsaved changes are never lost */

        return {
            byArtboardCheckbox: byArtboardCheckbox,
            fileNameLabelCheckbox: fileNameLabelCheckbox,
            includeGuidesCheckbox: includeGuidesCheckbox,
            scaleCheckbox: scaleCheckbox,
            scaleInput: scaleInput,
            /* 間隔を pt で返す（読めなければ NaN） / spacing in points, NaN when unreadable */
            getItemSpacingPt: function () {
                return evaluateArithmetic(spacingInput.text, spacingStepOptions.unit) * rulerUnit.pointsPerUnit;
            },
            closeDocRadio: closeDocRadio,
            getArtboardTargetMode: function () {
                return artboardFirstRadio.value ? "first" : (artboardAllRadio.value ? "all" : "specify");
            },
            artboardSpecInput: artboardSpecInput,
            setAfterImportEnabled: function (isEnabled) { afterImportRow.enabled = isEnabled; }
        };
    }

    /**
     * ダイアログを作る
     * @param {Document[]} openDocs - 開いているドキュメント
     * @returns {{dialog: Window, sourceUi: Object, destinationUi: Object, optionsUi: Object}} ダイアログと各パネル
     */
    function buildImportDialog(openDocs) {
        var dialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var sourceUi = buildSourcePanel(dialog, openDocs);
        var destinationUi = buildDestinationPanel(dialog);
        var optionsUi = buildOptionsPanel(dialog);

        /* 読み込み対象を切り替えたら、読み込み後の動作とラベルの既定を合わせる
           Follow the source mode: after-import action for open documents only; file-name labels on by default for open documents */
        sourceUi.onModeChange = function (useFolder) {
            optionsUi.setAfterImportEnabled(!useFolder);
            optionsUi.fileNameLabelCheckbox.value = !useFolder;
        };
        sourceUi.setFolderMode(openDocs.length === 0);

        var buttonRow = addButtonRow(dialog);
        /* name を "cancel" / "ok" にすると Esc / Enter でも閉じる / the names bind Esc / Enter */
        buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);

        return { dialog: dialog, sourceUi: sourceUi, destinationUi: destinationUi, optionsUi: optionsUi };
    }

    // =========================================
    // 設定の確定 / Settings
    // =========================================

    /**
     * 読み込み元から、読み込み先のドキュメント自身を外す（開いて閉じてしまう事故を防ぐ）
     * @param {Array} sources - Document か File の配列
     * @param {Document} targetDoc - 読み込み先のドキュメント
     * @param {boolean} fromFolder - sources が File なら true
     * @returns {Array} 外した残り
     */
    function excludeTargetDocument(sources, targetDoc, fromFolder) {
        var targetPath = null;
        if (fromFolder) {
            /* 保存済みのドキュメントだけパスで比べられる。未保存は fullName を読むと例外のことがある
               Only saved documents can be compared by path; fullName may throw for unsaved ones */
            try {
                if (targetDoc.fullName.exists) targetPath = targetDoc.fullName.fsName;
            } catch (e) {
                return sources;
            }
            if (targetPath === null) return sources;
        }
        var remainingSources = [];
        for (var i = 0; i < sources.length; i++) {
            var isTarget = fromFolder ? (sources[i].fsName === targetPath) : (sources[i] === targetDoc);
            if (!isTarget) remainingSources.push(sources[i]);
        }
        return remainingSources;
    }

    /**
     * ダイアログの値を確かめ、読み込みの設定にまとめる。正しくなければ知らせて null を返す
     * @param {{sourceUi: Object, destinationUi: Object, optionsUi: Object}} dialogUi - buildImportDialog() の戻り値
     * @returns {Object|null} 読み込みの設定（中止なら null）
     */
    function readImportSettings(dialogUi) {
        var sourceUi = dialogUi.sourceUi;
        var destinationUi = dialogUi.destinationUi;
        var optionsUi = dialogUi.optionsUi;

        /* 入力途中は無視してよいが、OK のあとに正しくない式を「すべて対象」にはしない（意図しないファイルを読み込まない）
           After OK, an invalid pattern must not fall back to "match all" */
        if (sourceUi.filterInput.text !== "" && !sourceUi.getNameFilterRegExp()) {
            alert(getLabel('alert.invalidFilter'));
            return null;
        }

        var importFromFolder = sourceUi.isFolderMode();
        var useCurrentDoc = destinationUi.currentDocRadio.value;
        var sources = sourceUi.getSources();
        var targetDoc = null;
        if (useCurrentDoc) {
            if (app.documents.length === 0) {
                alert(getLabel('alert.noCurrentDoc'));
                return null;
            }
            targetDoc = app.activeDocument;
            sources = excludeTargetDocument(sources, targetDoc, importFromFolder);
        }
        if (sources.length < 1) {
            alert(getLabel('alert.noValidFile'));
            return null;
        }

        var importByArtboard = optionsUi.byArtboardCheckbox.value;
        var artboardTargetMode = optionsUi.getArtboardTargetMode();
        var artboardSpecText = optionsUi.artboardSpecInput.text;
        /* 番号を1つも読めない指定は、黙って何もせずに終わらないよう止める / stop instead of silently importing nothing */
        if (importByArtboard && artboardTargetMode === "specify" && parseArtboardNumbers(artboardSpecText).length === 0) {
            alert(getLabel('alert.invalidArtboardSpec'));
            return null;
        }

        /* 拡大・縮小がオンで正しくない値（数値でない・0以下）は、等倍にせず止める / stop on an invalid scale instead of using 100% */
        var scalePercent = 100;
        if (optionsUi.scaleCheckbox.value) {
            scalePercent = parseFloat(optionsUi.scaleInput.text);
            if (isNaN(scalePercent) || scalePercent <= 0) {
                alert(getLabel('alert.invalidScale'));
                return null;
            }
        }

        var itemSpacingPt = optionsUi.getItemSpacingPt();
        if (isNaN(itemSpacingPt) || itemSpacingPt < 0) {
            alert(getLabel('alert.invalidSpacing'));
            return null;
        }

        /* 新規ドキュメントの大きさは、進捗パレットを出す前に確かめる / validate the size before the progress palette appears */
        var docWidthPt = 0;
        var docHeightPt = 0;
        var newDocSettings = destinationUi.getNewDocSettings();
        if (!useCurrentDoc) {
            var pointsPerUnit = SIZE_UNIT_POINTS[newDocSettings.unit];
            docWidthPt = parseFloat(newDocSettings.widthText) * pointsPerUnit;
            docHeightPt = parseFloat(newDocSettings.heightText) * pointsPerUnit;
            if (isNaN(docWidthPt) || isNaN(docHeightPt)) {
                alert(getLabel('alert.invalidNumber'));
                return null;
            }
        }

        /* 開いているドキュメントを閉じるとき、未保存の変更があれば失われるので確かめる / confirm before discarding unsaved changes */
        var closeOpenDocs = !importFromFolder && optionsUi.closeDocRadio.value;
        if (closeOpenDocs && hasUnsavedDocument(sources) && !confirm(getLabel('confirm.discardUnsaved'))) {
            return null;
        }

        return {
            importFromFolder: importFromFolder,
            sources: sources,
            targetDoc: targetDoc,
            byArtboard: importByArtboard,
            artboardTargetMode: artboardTargetMode,
            artboardSpecText: artboardSpecText,
            includeGuides: optionsUi.includeGuidesCheckbox.value,
            showLabel: optionsUi.fileNameLabelCheckbox.value,
            scalePercent: scalePercent,
            itemSpacingPt: itemSpacingPt,
            /* 分割は新規ドキュメントのときだけ / splitting only applies to new documents */
            splitFileCount: useCurrentDoc ? 0 : (destinationUi.getSplitFileCount() || 0),
            docWidthPt: docWidthPt,
            docHeightPt: docHeightPt,
            profileName: newDocSettings.profileName,
            colorSpace: newDocSettings.isRgb ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK,
            rasterResolution: RESOLUTION_PRESET_VALUES[newDocSettings.ppiIndex],
            rulerUnits: SIZE_UNIT_RULER_UNITS[newDocSettings.unit],
            /* フォルダーから開いたファイルは常に閉じる。開いているドキュメントは［読み込み後］に従う
               Files opened from a folder are always closed; open documents follow the After Import choice */
            shouldCloseSource: importFromFolder || closeOpenDocs
        };
    }

    /**
     * 未保存の変更があるドキュメントが含まれるか判定する
     * @param {Document[]} documents - ドキュメント
     * @returns {boolean} 1つでもあれば true
     */
    function hasUnsavedDocument(documents) {
        for (var i = 0; i < documents.length; i++) {
            if (!documents[i].saved) return true;
        }
        return false;
    }

    // =========================================
    // 読み込みの実行 / Import
    // =========================================

    /**
     * 進捗パレットを表示する
     * @param {number} totalCount - 読み込む件数
     * @returns {{update: Function, isCancelled: Function, close: Function}} 進捗の更新・中止の確認・閉じる
     */
    function showProgressPalette(totalCount) {
        var progressPalette = new Window("palette", getLabel('progress.title'));
        setupWindow(progressPalette);
        var progressTextGroup = progressPalette.add("group");
        progressTextGroup.alignment = ["center", "top"];
        var progressCountText = progressTextGroup.add("statictext", undefined, labelValueText('progress.count', "0/" + totalCount));
        progressCountText.preferredSize = [100, 30];

        var progressBar = progressPalette.add("progressbar", undefined, 0, totalCount);
        progressBar.preferredSize = [300, 6];

        var progressCancelGroup = progressPalette.add("group");
        progressCancelGroup.alignment = "right";
        var btnStop = progressCancelGroup.add("button", undefined, getLabel('button.cancel'));

        var isCancelled = false;
        btnStop.onClick = function () { isCancelled = true; };
        progressPalette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") isCancelled = true;
        });
        progressPalette.show();

        return {
            update: function (doneCount) {
                progressBar.value = doneCount;
                progressCountText.text = labelValueText('progress.count', doneCount + "/" + totalCount);
                progressPalette.update();
            },
            isCancelled: function () { return isCancelled; },
            close: function () {
                /* パレットの close() はまれに MRAP エラーを出す（処理は完了済み）。本体の例外を上書きしないよう握りつぶす
                   close() occasionally throws a spurious MRAP error; swallow it so it never masks a real error */
                try { progressPalette.hide(); } catch (e) { }
                try { progressPalette.close(); } catch (err) { }
            }
        };
    }

    /**
     * 新規ドキュメントを、選んだプロファイルで作る（通常の上限を超える大きさは自動でラージカンバスになる）。
     * DocumentPreset の width / height は units にかかわらず pt で、units はドキュメントの定規の単位になる
     * @param {Object} importSettings - readImportSettings() の戻り値
     * @returns {Document} 作ったドキュメント
     */
    function createDestinationDocument(importSettings) {
        var docPreset = new DocumentPreset();
        docPreset.units = importSettings.rulerUnits;
        docPreset.width = importSettings.docWidthPt;
        docPreset.height = importSettings.docHeightPt;
        docPreset.colorMode = importSettings.colorSpace;
        docPreset.rasterResolution = importSettings.rasterResolution;
        docPreset.numArtboards = 1;
        var destDoc = app.documents.addDocument(importSettings.profileName, docPreset);
        app.activeDocument = destDoc;
        return destDoc;
    }

    /**
     * 配置の状態を作る。基準は新規なら最初の（仮の）アートボード、現在のドキュメントならアクティブなアートボード
     * @param {Document} destDoc - 読み込み先
     * @param {Object} importSettings - 読み込みの設定
     * @returns {Object} placementContext
     */
    function createPlacementContext(destDoc, importSettings) {
        var useCurrentDoc = (importSettings.targetDoc !== null);
        var baseArtboardIndex = useCurrentDoc ? destDoc.artboards.getActiveArtboardIndex() : 0;
        var baseRect = destDoc.artboards[baseArtboardIndex].artboardRect;
        return {
            destDoc: destDoc,
            cells: [],
            canvasCenterX: (baseRect[0] + baseRect[2]) / 2,
            canvasCenterY: (baseRect[1] + baseRect[3]) / 2,
            /* 現在のドキュメントでは、読み込み前のアートボードの下に重ならないよう並べる / place below the existing artboards */
            avoidRect: useCurrentDoc ? getArtboardsUnionRect(destDoc) : null,
            byArtboard: importSettings.byArtboard,
            createArtboards: importSettings.byArtboard, /* アートボードを作るのはアートボード単位のときだけ / artboards only per artboard */
            placedCount: 0,
            showLabel: importSettings.showLabel,
            artboardPadding: IMPORT_SETTINGS.artboardMargin,
            itemSpacing: importSettings.itemSpacingPt,
            scalePercent: importSettings.scalePercent
        };
    }

    /**
     * 1つの読み込み元を取り込む。失敗しても一時ファイルは必ず閉じ、ロック・非表示も戻す
     * @param {Document|File} source - 開いているドキュメントか、フォルダーのファイル
     * @param {Object} importSettings - 読み込みの設定
     * @param {Object} placementContext - 配置の状態
     * @returns {void}
     */
    function importSource(source, importSettings, placementContext) {
        var sourceDoc = importSettings.importFromFolder ? app.open(source) : source;
        var lockState = { layers: [], items: [] };
        var guidesWereLocked = false;
        try {
            app.activeDocument = sourceDoc;
            var labelName = sourceDoc.name.replace(/\.[^\.]+$/, "");
            if (importSettings.byArtboard) {
                var targetIndices = resolveTargetArtboardIndices(sourceDoc, importSettings.artboardTargetMode, importSettings.artboardSpecText);
                /* すべて：ドキュメント全体を解除。1のみ・指定：対象アートボードに重なるものだけ解除し、ほかは触れない
                   All: unlock everything. First / Specify: unlock only what overlaps the target artboards */
                var targetRects = null;
                if (importSettings.artboardTargetMode !== "all") {
                    targetRects = [];
                    for (var i = 0; i < targetIndices.length; i++) targetRects.push(sourceDoc.artboards[targetIndices[i]].artboardRect);
                }
                unlockItemsForImport(sourceDoc, targetRects, lockState);
                if (importSettings.includeGuides) guidesWereLocked = unlockGuidesInView(sourceDoc);
                importArtboards(sourceDoc, targetIndices, labelName, importSettings, placementContext);
            } else {
                if (importSettings.includeGuides) guidesWereLocked = unlockGuidesInView(sourceDoc);
                importWholeDocument(sourceDoc, labelName, importSettings, placementContext);
            }
        } finally {
            if (importSettings.shouldCloseSource) {
                sourceDoc.close(SaveOptions.DONOTSAVECHANGES); /* 閉じるので元に戻さなくてよい / closing makes the restore unnecessary */
            } else {
                restoreLockHiddenState(lockState);
                if (guidesWereLocked) {
                    app.activeDocument = sourceDoc;
                    app.executeMenuCommand("lockguide"); /* ［ガイドをロック］を戻す / turn Lock Guides back on */
                }
            }
        }
    }

    /**
     * ［表示］→［ガイド］→［ガイドをロック］がオンなら、オフに切り替える。
     * オンのあいだはガイドを選択できず（例外も出ず選択数が0のまま）、ガイドの locked にも現れないので、1本選んでみて判定する
     * @param {Document} sourceDoc - 読み込み元（アクティブにしておく）
     * @returns {boolean} オフに切り替えたら true（あとで戻す）
     */
    function unlockGuidesInView(sourceDoc) {
        var probeGuide = null;
        for (var i = 0; i < sourceDoc.pageItems.length && !probeGuide; i++) {
            var pageItem = sourceDoc.pageItems[i];
            /* 個別のロック・非表示や、ロックされたレイヤーのガイドでは判定できない / skip guides that are unselectable for other reasons */
            if (pageItem.guides && !pageItem.locked && !pageItem.hidden && !pageItem.layer.locked && pageItem.layer.visible) probeGuide = pageItem;
        }
        if (!probeGuide) return false;
        sourceDoc.selection = [probeGuide];
        var guidesLocked = (sourceDoc.selection.length === 0);
        sourceDoc.selection = null;
        if (guidesLocked) app.executeMenuCommand("lockguide");
        return guidesLocked;
    }

    /**
     * アートボードごとに、ロック・非表示も含めて取り込む（元のアートボードの大きさと内容の位置を保つ）
     * @param {Document} sourceDoc - 読み込み元
     * @param {number[]} targetIndices - 取り込むアートボードの番号（0始まり）
     * @param {string} labelName - ラベル名
     * @param {Object} importSettings - 読み込みの設定
     * @param {Object} placementContext - 配置の状態
     * @returns {void}
     */
    function importArtboards(sourceDoc, targetIndices, labelName, importSettings, placementContext) {
        for (var i = 0; i < targetIndices.length; i++) {
            var artboardIndex = targetIndices[i];
            sourceDoc.artboards.setActiveArtboardIndex(artboardIndex);
            sourceDoc.selection = null;
            sourceDoc.selectObjectsOnActiveArtboard();
            if (sourceDoc.selection.length === 0) continue;

            var artboardRect = sourceDoc.artboards[artboardIndex].artboardRect;
            /* ガイドは位置を測る前に選択へ足す / add the guides before measuring */
            if (importSettings.includeGuides) {
                var artboardGuides = collectImportableGuides(sourceDoc, artboardRect, artboardRect);
                if (artboardGuides.length > 0) sourceDoc.selection = normalizeSelectionItems(sourceDoc.selection).concat(artboardGuides);
            }

            /* 元のアートボードの大きさと、その中の内容の位置。クリップグループはマスクで測る（ペースト側と同じ）
               Original artboard size and the content's position; clip groups by their mask, as on the pasted side */
            var selectionBounds = getClipAwareUnionBounds(sourceDoc.selection, true);
            var artboardCell = {
                width: artboardRect[2] - artboardRect[0],
                height: artboardRect[1] - artboardRect[3],
                offsetX: selectionBounds[0] - artboardRect[0],
                offsetY: artboardRect[1] - selectionBounds[1]
            };
            copySelectionToDocument(sourceDoc, placementContext.destDoc);
            placePastedGroup(placementContext, labelName, artboardCell); /* 失敗は placePastedGroup が知らせる / failures are alerted there */
            app.activeDocument = sourceDoc;
        }
    }

    /**
     * ファイル全体の表示中・選択できるオブジェクトを1つにまとめて取り込む（対象アートボードは使わない）
     * @param {Document} sourceDoc - 読み込み元
     * @param {string} labelName - ラベル名
     * @param {Object} importSettings - 読み込みの設定
     * @param {Object} placementContext - 配置の状態
     * @returns {void}
     */
    function importWholeDocument(sourceDoc, labelName, importSettings, placementContext) {
        app.executeMenuCommand("selectall");
        var importItems = [];
        for (var i = 0; i < sourceDoc.selection.length; i++) {
            var selectedItem = sourceDoc.selection[i];
            if (!selectedItem.locked && !selectedItem.hidden) importItems.push(selectedItem);
        }
        /* ガイドの長さは最初のアートボードを基準にする / guide length is judged against the first artboard */
        if (importSettings.includeGuides) {
            importItems = importItems.concat(collectImportableGuides(sourceDoc, sourceDoc.artboards[0].artboardRect, null));
        }
        if (importItems.length === 0) return;
        sourceDoc.selection = importItems;
        copySelectionToDocument(sourceDoc, placementContext.destDoc);
        placePastedGroup(placementContext, labelName);
    }

    /**
     * 読み込み元を順に取り込み、グリッドに並べる。分割するときは指定のファイル数ごとに新規ドキュメントを作る
     * @param {Object} importSettings - readImportSettings() の戻り値
     * @returns {void}
     */
    function runImport(importSettings) {
        var sources = importSettings.sources;
        var filesPerDoc = importSettings.splitFileCount > 0 ? importSettings.splitFileCount : sources.length;
        var progress = showProgressPalette(sources.length);
        var placedTotal = 0;
        var lastDestDoc = null;

        /* エラーが出ても進捗パレットは必ず閉じる / always close the progress palette, even on error */
        try {
            for (var chunkStart = 0; chunkStart < sources.length && !progress.isCancelled(); chunkStart += filesPerDoc) {
                lastDestDoc = importSettings.targetDoc || createDestinationDocument(importSettings);
                app.activeDocument = lastDestDoc;
                var placementContext = createPlacementContext(lastDestDoc, importSettings);
                var chunkEnd = Math.min(chunkStart + filesPerDoc, sources.length);
                for (var i = chunkStart; i < chunkEnd; i++) {
                    $.sleep(0);
                    app.redraw();
                    if (progress.isCancelled()) break; /* ここまでの分で後始末する / finish with what was placed so far */
                    importSource(sources[i], importSettings, placementContext);
                    progress.update(i + 1);
                }
                finishDestinationDocument(placementContext, importSettings);
                placedTotal += placementContext.placedCount;
            }
        } finally {
            progress.close();
        }

        if (progress.isCancelled()) {
            alert(getLabel('alert.cancelled'));
        } else if (placedTotal === 0) {
            /* 1件も置けなかったときも黙って終わらない / do not finish silently when nothing was placed */
            alert(getLabel('alert.noArtboardImported'));
        }
    }

    /**
     * 読み込み先の仕上げ：グリッドに並べ、新規ドキュメントの仮のアートボードを消し、全体を表示する
     * @param {Object} placementContext - 配置の状態
     * @param {Object} importSettings - 読み込みの設定
     * @returns {void}
     */
    function finishDestinationDocument(placementContext, importSettings) {
        var destDoc = placementContext.destDoc;
        layoutCellsAsCenteredGrid(placementContext);
        /* 新規ドキュメントは最初の仮のアートボードを消す（現在のドキュメントの既存のものは残す）
           Remove the placeholder artboard of a new document; keep the user's own in the current document */
        if (!importSettings.targetDoc && placementContext.placedCount > 0 && destDoc.artboards.length > placementContext.placedCount) {
            destDoc.artboards.remove(0);
        }
        app.activeDocument = destDoc;
        app.executeMenuCommand("fitall");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、選んだ読み込み元を取り込んで並べる
     * @returns {void}
     */
    function main() {
        var openDocs = [];
        for (var i = 0; i < app.documents.length; i++) openDocs.push(app.documents[i]);
        var dialogUi = buildImportDialog(openDocs);
        if (dialogUi.dialog.show() !== 1) return;
        var importSettings = readImportSettings(dialogUi);
        if (importSettings) runImport(importSettings);
    }

    main();

})();
