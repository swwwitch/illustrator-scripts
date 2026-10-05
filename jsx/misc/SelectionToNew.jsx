#targetengine "SelectionToNewEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトから、新規レイヤー・新規アートボード・新規ドキュメントを作成します。
作成対象はラジオボタン（`L` / `A` / `D` キー）で切り替え、元のオブジェクトを残すかどうかも指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionToNew.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n0f02f73a748d

### Overview

Creates a new layer, a new artboard, or a new document from the selected objects.
The target is picked with radio buttons (or the L / A / D keys), and you can choose whether the originals stay put.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectionToNew.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectionToNew";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-06";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionToNew.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectionToNew.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0f02f73a748d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 前回のダイアログ設定。#targetengine を指定しているのでエンジンはIllustratorの起動中は残るが、
       関数スコープの var は巻き上げで必ず undefined から始まり、実行のたびに作り直されるため
       前回の値を持ち越せない。$.global に載せてエンジン側に残す。
       名前はドキュメントごとに変わるので保持しない
       The last dialog settings. The #targetengine directive keeps the engine alive between runs,
       but a function-scoped var is hoisted as undefined and rebuilt every time, so it cannot
       carry anything over; keep it on $.global instead. Names are per-document, so they are
       deliberately not kept */
    var SESSION_OPTIONS_GLOBAL_KEY = "__SelectionToNewCreateOptions";

    /**
     * 前回の実行で残したダイアログ設定を取得します。
     *
     * @returns {object|null} 前回の設定。まだ無ければ null。
     */
    function getSessionCreateOptions() {
        return $.global[SESSION_OPTIONS_GLOBAL_KEY] || null;
    }

    /**
     * 今回のダイアログ設定を次回の実行のために残します。
     *
     * @param {object} createOptions - 残すダイアログ設定。
     * @returns {void}
     */
    function setSessionCreateOptions(createOptions) {
        $.global[SESSION_OPTIONS_GLOBAL_KEY] = createOptions;
    }

    (function () {

        // =========================================
        // ユーザー設定 / User settings
        // =========================================

        /* 新規アートボードの挿入位置 / Insert position of the new artboard */
        /* true = 現在のアートボードの次 / false = 末尾 */
        /* true = after the current artboard, false = at the end */
        var ARTBOARD_INSERT_AFTER_CURRENT = true;

        /* ダイアログの［方向］の初期選択 / Initial selection of the dialog's Direction */
        /* 0 = 右（横並び） / 1 = 下（縦並び） */
        /* 0 = right (horizontal), 1 = down (vertical) */
        var ARTBOARD_DIRECTION_AXIS = 0;

        /* 新規アートボード名の初期値に付ける接尾辞 / Suffix for the new artboard's default name */
        var ARTBOARD_NAME_SUFFIX = "_new";

        /* 複製ドキュメントのファイル名の初期値に付ける接尾辞
           Suffix for the duplicated document's default file name */
        var DOCUMENT_NAME_SUFFIX = "_selection";

        /* 作成対象を切り替えるキーボードショートカット
           Keyboard shortcuts that switch the create target */
        var SHORTCUT_TARGETS = { L: "layer", A: "artboard", D: "document" };

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

        // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

        /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
        var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

        /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
        var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

        /* 修飾キーの別名 / Aliases accepted for the modifiers */
        var KEY_SHORTCUT_MODIFIER_ALIASES = {
            SHIFT: "SHIFT",
            ALT: "ALT", OPTION: "ALT", OPT: "ALT",
            CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
        };

        /**
         * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
         * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
         * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
         */
        function normalizeKeyShortcutSpec(keySpec) {
            var specParts = String(keySpec).split("+");
            var baseKey = specParts.pop().toUpperCase();
            var modifierFlags = {};
            for (var i = 0; i < specParts.length; i++) {
                var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
                if (modifierName) modifierFlags[modifierName] = true;
            }
            return buildKeyShortcutSpec(modifierFlags, baseKey);
        }

        /**
         * 修飾キーの状態とキー名から照合用の表記を組み立てる
         * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
         * @param {string} baseKey - 大文字のキー名
         * @returns {string} 照合用の表記
         */
        function buildKeyShortcutSpec(modifierFlags, baseKey) {
            var specText = "";
            for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
                if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
            }
            return specText + baseKey;
        }

        /**
         * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
         * @param {Object} keyEvent - keydown イベント
         * @returns {string} 照合用の表記。キー名が無いときは空文字
         */
        function readKeyShortcutSpec(keyEvent) {
            if (!keyEvent || !keyEvent.keyName) return "";
            var keyboardState = {};
            try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
            var modifierFlags = {
                SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
                ALT: !!(keyEvent.altKey || keyboardState.altKey),
                CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
            };
            return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
        }

        /**
         * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
         * @param {Object} control - コントロール
         * @returns {boolean} 押せるなら true
         */
        function isKeyShortcutControlUsable(control) {
            for (var node = control; node; node = node.parent) {
                if (node.enabled === false || node.visible === false) return false;
            }
            return true;
        }

        /**
         * キーを受けたコントロールが、文字を入力する欄か
         * @param {Object} focusedControl - イベントの発生元
         * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
         * @returns {boolean} 入力中としてショートカットを止めるなら true
         */
        function isKeyShortcutTypingTarget(focusedControl, numericFields) {
            if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
            for (var i = 0; i < numericFields.length; i++) {
                if (numericFields[i] === focusedControl) return false;
            }
            return true;
        }

        /**
         * コントロールをクリックしたときと同じ動作をする
         * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
         * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
         * @returns {void}
         */
        function pressKeyShortcutControl(control) {
            if (control.type === "radiobutton") {
                /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
                var siblings = control.parent ? control.parent.children : [];
                for (var i = 0; i < siblings.length; i++) {
                    if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
                }
                control.value = true;
            } else if (control.type === "checkbox") {
                control.value = !control.value;
            }
            if (typeof control.onClick === "function") {
                control.onClick.call(control);
            } else if (control.type === "button" && typeof control.notify === "function") {
                /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
                control.notify("onClick");
            }
        }

        /**
         * 1つのショートカットを実行する
         * @param {Object|Function} shortcutTarget - コントロール、または関数
         * @param {Object} keyEvent - keydown イベント
         * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
         */
        function runKeyShortcutTarget(shortcutTarget, keyEvent) {
            var targetControl = shortcutTarget;
            if (typeof shortcutTarget === "function") {
                var runResult = shortcutTarget(keyEvent);
                if (runResult === false || runResult === null) return false;
                if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
                targetControl = runResult;
            }
            /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
            if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
            return true;
        }

        /**
         * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
         * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
         * @returns {string} 表示用の表記（"Shift+R" など）
         */
        function formatKeyShortcutLabel(normalizedSpec) {
            var isMac = ($.os.indexOf("Mac") === 0);
            var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
            var specParts = normalizedSpec.split("+");
            var baseKey = specParts.pop();
            var labelText = "";
            for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
            if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
            return labelText + baseKey;
        }

        /**
         * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
         * @param {Object} control - コントロール
         * @param {string} normalizedSpec - 照合用の表記
         * @returns {void}
         */
        function appendKeyShortcutToTip(control, normalizedSpec) {
            var keyLabel = formatKeyShortcutLabel(normalizedSpec);
            var currentTip = control.helpTip ? String(control.helpTip) : "";
            if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
            var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
            control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
        }

        /**
         * ダイアログ・パレットに文字キーのショートカットを付ける
         * @param {Window} targetWindow - キーを受けるダイアログ・パレット
         * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
         * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
         * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
         */
        function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
            var shortcutSettings = shortcutOptions || {};
            var numericFields = shortcutSettings.numericFields || [];
            var bindingTable = {};

            for (var keySpec in shortcutMap) {
                if (!shortcutMap.hasOwnProperty(keySpec)) continue;
                var mapEntry = shortcutMap[keySpec];
                if (!mapEntry) continue;
                var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
                var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
                bindingTable[normalizedSpec] = {
                    target: isWrapped ? mapEntry.target : mapEntry,
                    inFields: !!(isWrapped && mapEntry.inFields)
                };
                var tipControl = bindingTable[normalizedSpec].target;
                if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                    appendKeyShortcutToTip(tipControl, normalizedSpec);
                }
            }

            /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
            targetWindow.addEventListener("keydown", function (keyEvent) {
                var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
                if (!binding) return;
                if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
                if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
                if (keyEvent.preventDefault) keyEvent.preventDefault();
                if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
            }, true);

            return bindingTable;
        }

        // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialog: {
                title: { ja: "選択オブジェクトから新規作成", en: "New from Selection" }
            },
            panel: {
                createTarget: { ja: "作成するものと名前", en: "Create" },
                artboardLayout: { ja: "アートボードの配置", en: "Artboard placement" },
                options: { ja: "オプション", en: "Options" }
            },
            label: {
                direction: { ja: "方向", en: "Direction" },
                spacing: { ja: "間隔", en: "Spacing" }
            },
            radio: {
                targetLayer: { ja: "レイヤー", en: "Layer" },
                targetArtboard: { ja: "アートボード", en: "Artboard" },
                targetDocument: { ja: "ドキュメント", en: "Document" },
                directionRight: { ja: "右", en: "Right" },
                directionDown: { ja: "下", en: "Down" }
            },
            checkbox: {
                duplicate: { ja: "元のオブジェクトを残す", en: "Keep original objects" },
                includeLocked: { ja: "ロックされたオブジェクトを含める", en: "Include locked objects" },
                includeHidden: { ja: "非表示オブジェクトを含める", en: "Include hidden objects" }
            },
            button: {
                ok: { ja: "OK", en: "OK" },
                cancel: { ja: "キャンセル", en: "Cancel" }
            },
            artwork: {
                newLayerName: { ja: "新規レイヤー", en: "New Layer" }
            },
            alert: {
                noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
                noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
                artboardLimit: {
                    ja: "アートボードの最大数を超えるため、作成できません。",
                    en: "Cannot create: the maximum number of artboards would be exceeded."
                },
                noSpace: {
                    ja: "十分なスペースがないため、アートボードを作成できません。",
                    en: "Cannot create the artboard: there is not enough space."
                },
                needsSave: {
                    ja: "ドキュメントを複製するため、先に保存してください。",
                    en: "Save the document first so it can be duplicated."
                },
                sameAsSource: {
                    ja: "元のファイルと同じ名前は指定できません。別の名前を入力してください。",
                    en: "The name must differ from the source file. Enter another name."
                },
                overwrite: {
                    ja: "同名のファイルがすでにあります。上書きしますか？",
                    en: "A file with that name already exists. Overwrite it?"
                },
                unexpected: {
                    ja: "エラーが発生したため、処理を実行できませんでした。\nエラー内容：",
                    en: "Processing could not be completed because an error occurred.\nError:"
                }
            },
            tip: {
                targetLayer: {
                    ja: "選択オブジェクトを収めた新規レイヤーを、最前面に作成します。",
                    en: "Creates a new layer at the top holding the selected objects."
                },
                targetArtboard: {
                    ja: "アクティブアートボードと同じサイズの新規アートボードを作成し、相対位置を保って配置します。",
                    en: "Creates a new artboard matching the active one and keeps the relative position."
                },
                targetDocument: {
                    ja: "選択オブジェクトだけを残した複製ドキュメントを作成します。保存済みのドキュメントでのみ実行できます。",
                    en: "Creates a duplicate document containing only the selection. Saved documents only."
                },
                extension: {
                    ja: "元のファイルと同じ拡張子で保存します。",
                    en: "Saved with the same extension as the source file."
                },
                ok: {
                    ja: "選択した内容で作成します。",
                    en: "Create with the selected settings."
                },
                cancel: {
                    ja: "何もせずに閉じます。",
                    en: "Close without doing anything."
                },
                layerName: {
                    ja: "同名のレイヤーがある場合は連番が付きます。",
                    en: "A numeric suffix is added when a layer with that name exists."
                },
                artboardName: {
                    ja: "現在のアートボードの次に、同じサイズで作成します。",
                    en: "Created next to the current artboard, at the same size."
                },
                documentName: {
                    ja: "元のファイルと同じ場所に作成します。同名のファイルがある場合は確認します。",
                    en: "Created next to the source file. An existing file prompts for confirmation."
                },
                direction: {
                    ja: "新規アートボードを現在のアートボードのどちら側に並べるかを選びます。",
                    en: "Choose which side of the current artboard the new one goes."
                },
                spacing: {
                    ja: "アートボードどうしの間隔。初期値は既存の並びから推定した値で、定規の単位で入力します。",
                    en: "Gap between artboards, in the ruler unit. The default is inferred from the existing layout."
                },
                duplicate: {
                    ja: "ドキュメントは元のドキュメントを残すため、常に複製になります。",
                    en: "Document always keeps the source document, so it is always a copy."
                },
                includeObjects: {
                    ja: "ドキュメント作成時のみ。現在のアートボード上にある最上位のオブジェクトが対象です。",
                    en: "Document only. Applies to top-level objects on the current artboard."
                }
            },
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
            },
            progress: {
                title: { ja: "処理中", en: "Working" },
                duplicating: { ja: "ドキュメントを複製しています...", en: "Duplicating the document..." },
                deleting: { ja: "選択オブジェクト以外を削除しています...", en: "Removing everything but the selection..." },
                reopening: { ja: "元のドキュメントを開き直しています...", en: "Reopening the source document..." }
            }
        };

        // =========================================
        // 矩形の共通処理 / Rect helpers
        // =========================================

        /**
         * 矩形の中心座標を求めます。
         *
         * @param {Array<number>} rect - [左, 上, 右, 下]。
         * @returns {Array<number>} [X, Y] の中心座標。
         */
        function getRectCenter(rect) {
            return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
        }

        /**
         * 座標が矩形の内側にあるかを判定します。
         *
         * @param {Array<number>} point - [X, Y] の座標。
         * @param {Array<number>} rect - [左, 上, 右, 下]。
         * @returns {boolean} 内側にある場合は true。
         */
        function isPointInsideRect(point, rect) {
            /* Y軸は上→下なので、上端以下・下端以上で判定する
               The Y axis is top-down, so compare against top as max and bottom as min */
            return point[0] >= rect[0] && point[0] <= rect[2] &&
                   point[1] <= rect[1] && point[1] >= rect[3];
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

        // =========================================
        // 選択オブジェクトの取得 / Collecting the selection
        // =========================================

        /**
         * 選択オブジェクトを配列にスナップショットします。
         *
         * `doc.selection` は参照するたびに現在の選択状態から作り直されるため、
         * move() で選択が変化するループ内で直接参照すると対象を取りこぼします。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {Array<PageItem>} 選択オブジェクトの配列。
         */
        function snapshotSelection(doc) {
            /* ライブな selection を配列へ写し取る / Copy the live selection into a plain array */
            var selectedItems = [];
            var liveSelection = doc.selection;

            for (var i = 0; i < liveSelection.length; i++) {
                selectedItems.push(liveSelection[i]);
            }
            return selectedItems;
        }

        /**
         * 重ね順の比較に使うソートキーを取得します。
         *
         * レイヤーとコンテナそれぞれの zOrderPosition を組み合わせ、
         * 値が大きいほど前面（上）になるようにします。
         *
         * @param {PageItem} item - 対象のオブジェクト。
         * @returns {{layerOrder: number, itemOrder: number}} 並べ替え用のキー。
         */
        function getStackOrderKey(item) {
            /* zOrderPosition を持たないアイテムもあるため、まとめて保護する
               Some items expose no zOrderPosition, so guard the whole lookup */
            try {
                return { layerOrder: item.layer.zOrderPosition, itemOrder: item.zOrderPosition };
            } catch (e) {
                return { layerOrder: 0, itemOrder: 0 };
            }
        }

        /**
         * オブジェクトを前面から背面の順（元の重ね順）に並べ替えます。
         *
         * @param {Array<PageItem>} items - 並べ替えるオブジェクトの配列。
         * @returns {Array<PageItem>} 前面が先頭になるよう並べ替えた配列。
         */
        function sortByStackOrder(items) {
            /* 並べ替えキーを先に作る / Build the sort keys up front */
            var sortEntries = [];

            for (var i = 0; i < items.length; i++) {
                sortEntries.push({ item: items[i], key: getStackOrderKey(items[i]), index: i });
            }

            sortEntries.sort(function (entryA, entryB) {
                if (entryA.key.layerOrder !== entryB.key.layerOrder) {
                    return entryB.key.layerOrder - entryA.key.layerOrder;
                }
                if (entryA.key.itemOrder !== entryB.key.itemOrder) {
                    return entryB.key.itemOrder - entryA.key.itemOrder;
                }
                /* キーが同じときは元の並び順を保つ / Keep the original order for ties */
                return entryA.index - entryB.index;
            });

            var sortedItems = [];
            for (var j = 0; j < sortEntries.length; j++) {
                sortedItems.push(sortEntries[j].item);
            }
            return sortedItems;
        }

        /**
         * オブジェクト群の中心を含むアートボードのインデックスを求めます。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {Array<PageItem>} items - 対象のオブジェクト。
         * @param {number} fallbackIndex - 見つからないときに返すインデックス。
         * @returns {number} アートボードのインデックス。
         */
        function findArtboardIndexForItems(doc, items, fallbackIndex) {
            var unionBounds = getClipAwareUnionBounds(items, false);
            if (unionBounds === null) return fallbackIndex;

            /* 中心が載っているアートボードを探す / Find the artboard holding the center */
            var center = getRectCenter(unionBounds);
            for (var i = 0; i < doc.artboards.length; i++) {
                if (isPointInsideRect(center, doc.artboards[i].artboardRect)) return i;
            }
            return fallbackIndex;
        }

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

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        /**
         * ダイアログの設定値。
         *
         * @typedef {object} CreateOptions
         * @property {string} createTarget - 作成対象（"layer" / "artboard" / "document"）。
         * @property {string} createName - 新規レイヤー名／アートボード名／ファイル名（拡張子なし）。
         * @property {number} directionAxis - アートボードを並べる方向（0=右 / 1=下）。
         * @property {number} spacingPt - アートボードどうしの間隔（pt）。
         * @property {boolean} useDuplicate - 元のオブジェクトを残して複製する場合は true。
         * @property {boolean} includeLocked - ロックされたオブジェクトも残す場合は true。
         * @property {boolean} includeHidden - 非表示オブジェクトも残す場合は true。
         */

        // =========================================
        // 単位 / Unit
        // =========================================

        /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
         * 設定キーごとの単位情報を取得する
         * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
         * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
         */
        function getUnitInfo(prefKey) {
            var unitKey = prefKey || "rulerType";
            var unitCode = app.preferences.getIntegerPreference(unitKey);
            var unit = UNITS[unitCode] || UNITS[2];
            var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
            return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
        }

        /**
         * 表示用に小数を整えます。
         *
         * @param {number} value - 整形する数値。
         * @returns {string} 小数第2位までに丸めた文字列。
         */
        function formatSpacingValue(value) {
            return String(Math.round(value * 100) / 100);
        }

        /**
         * 作成対象ごとの名前の初期値を組み立てます。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {{layer: string, artboard: string, document: string}} 対象ごとの初期値。
         */
        function buildDefaultNames(doc) {
            /* アートボードは基準になるアートボード名、ドキュメントは元のファイル名から作る
               The artboard follows the active artboard, the document follows the file name */
            var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];

            return {
                layer: getLabel(LABELS.artwork.newLayerName),
                artboard: activeArtboard.name + ARTBOARD_NAME_SUFFIX,
                document: getFileBaseName(doc.name) + DOCUMENT_NAME_SUFFIX
            };
        }

        /* ラジオボタンの幅。日英どちらの文言も収まる値にそろえて、
           3行の入力欄が縦に並ぶようにする
           Radio width: fixed so the three input fields line up, and wide enough
           for both languages */
        var TARGET_RADIO_WIDTH = 120;

        /**
         * ラジオボタンと名前の入力欄を横一列に並べた、作成対象1つ分の行を作ります。
         *
         * 何の名前かはラジオボタンのラベルで分かるので、入力欄に見出しは付けません。
         *
         * @param {Panel} parentPanel - 追加先のパネル。
         * @param {object} radioEntry - ラジオボタンのラベル定義。
         * @param {string} defaultName - 入力欄の初期値。
         * @param {object} inputTipEntry - 入力欄のツールチップ定義。
         * @param {object} radioTipEntry - ラジオボタンのツールチップ定義。
         * @returns {{radio: RadioButton, input: EditText, group: Group}} 作った部品。
         */
        function addTargetRow(parentPanel, radioEntry, defaultName, inputTipEntry, radioTipEntry) {
            var rowGroup = parentPanel.add("group");
            rowGroup.orientation = "row";
            rowGroup.alignChildren = ["left", "center"];
            rowGroup.spacing = 8;

            var targetRadio = rowGroup.add("radiobutton", undefined, getLabel(radioEntry));
            targetRadio.preferredSize.width = TARGET_RADIO_WIDTH;
            targetRadio.helpTip = getLabel(radioTipEntry);

            var nameInput = rowGroup.add("edittext", undefined, defaultName);
            nameInput.characters = 18;
            nameInput.helpTip = getLabel(inputTipEntry);

            return { radio: targetRadio, input: nameInput, group: rowGroup };
        }

        /**
         * 何を作成するかを選択するダイアログを表示します。
         *
         * @param {Document} doc - 対象のドキュメント。名前の初期値を作るのに使います。
         * @returns {CreateOptions|null} 選択された設定。キャンセルされた場合は null。
         */
        function showCreateTargetDialog(doc) {
            var defaultNames = buildDefaultNames(doc);
            var documentExtension = getFileExtension(doc.name) || ".ai";

            /* タイトルにバージョンを添える / Show the version next to the title */
            var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
            setupWindow(dialog);

            /* 作成対象パネル。1行につき「ラジオボタン＋名前の入力欄」を並べる
               Create target panel: one row per target, radio + name input */
            var createTargetPanel = dialog.add("panel", undefined, getLabel(LABELS.panel.createTarget));
            setupPanel(createTargetPanel, 6);

            var layerRow = addTargetRow(createTargetPanel, LABELS.radio.targetLayer,
                defaultNames.layer, LABELS.tip.layerName, LABELS.tip.targetLayer);
            var artboardRow = addTargetRow(createTargetPanel, LABELS.radio.targetArtboard,
                defaultNames.artboard, LABELS.tip.artboardName, LABELS.tip.targetArtboard);
            var documentRow = addTargetRow(createTargetPanel, LABELS.radio.targetDocument,
                defaultNames.document, LABELS.tip.documentName, LABELS.tip.targetDocument);

            /* ドキュメントだけは拡張子を添えて示す / Show the extension next to the document name */
            var extensionLabel = documentRow.group.add("statictext", undefined, documentExtension);
            extensionLabel.helpTip = getLabel(LABELS.tip.extension);

            /* アートボードの配置パネル / Artboard placement panel */
            var rulerUnit = getUnitInfo("rulerType");
            var autoSpacingPt = computeAutoSpacingPt(doc,
                app.preferences.getRealPreference('plugin/ArtboardRearrange/ArtboardSpacing'));

            var artboardLayoutPanel = dialog.add("panel", undefined, getLabel(LABELS.panel.artboardLayout));
            setupPanel(artboardLayoutPanel, 6);
            /* 項目を横一列に並べる / Lay the controls out in one row */
            artboardLayoutPanel.orientation = "row";
            artboardLayoutPanel.alignChildren = ["left", "center"];

            var directionLabel = artboardLayoutPanel.add("statictext", undefined, labelText(LABELS.label.direction));

            /* ラジオボタンは同じグループに入れて排他にする
               Keep both radios in one group so ScriptUI makes them exclusive */
            var directionGroup = artboardLayoutPanel.add("group");
            directionGroup.orientation = "row";
            directionGroup.spacing = 8;
            var directionRightRadio = directionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.directionRight));
            var directionDownRadio = directionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.directionDown));

            var spacingLabel = artboardLayoutPanel.add("statictext", undefined, labelText(LABELS.label.spacing));
            /* ∧∨と入力欄は隙間0で突き合わせる。間隔は0未満にしない
               Butt the stepper against the field; the spacing never goes below 0 */
            var spacingStepperGroup = artboardLayoutPanel.add("group");
            spacingStepperGroup.orientation = "row";
            spacingStepperGroup.alignChildren = ["left", "center"];
            spacingStepperGroup.spacing = 0;
            spacingStepperGroup.margins = 0;
            var spacingInput;
            var spacingStepper = addStepper(spacingStepperGroup, function () { return spacingInput; }, { min: 0 });
            spacingInput = spacingStepperGroup.add("edittext", undefined,
                formatSpacingValue(autoSpacingPt / rulerUnit.pointsPerUnit));
            spacingInput.characters = 5;
            /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
            bindSteppedArrowKeys(spacingInput, spacingStepper);
            var spacingUnitLabel = artboardLayoutPanel.add("statictext", undefined, rulerUnit.label);

            directionRightRadio.helpTip = getLabel(LABELS.tip.direction);
            directionDownRadio.helpTip = getLabel(LABELS.tip.direction);
            spacingInput.helpTip = getLabel(LABELS.tip.spacing);

            /* オプションパネル / Options panel */
            var optionPanel = dialog.add("panel", undefined, getLabel(LABELS.panel.options));
            setupPanel(optionPanel, 6);

            var duplicateCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.duplicate));
            var includeLockedCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.includeLocked));
            var includeHiddenCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.includeHidden));

            /* ディム表示になる理由はUIから読み取れないので、ツールチップで補う
               Nothing on screen explains why a control is dimmed, so tooltips fill that in */
            duplicateCheckbox.helpTip = getLabel(LABELS.tip.duplicate);
            includeLockedCheckbox.helpTip = getLabel(LABELS.tip.includeObjects);
            includeHiddenCheckbox.helpTip = getLabel(LABELS.tip.includeObjects);

            /* 前回の設定があれば復元する / Restore the previous settings when there are any */
            var savedOptions = getSessionCreateOptions();
            duplicateCheckbox.value = (savedOptions !== null) && savedOptions.useDuplicate;
            includeLockedCheckbox.value = (savedOptions !== null) && savedOptions.includeLocked;
            includeHiddenCheckbox.value = (savedOptions !== null) && savedOptions.includeHidden;

            var savedDirectionAxis = (savedOptions !== null) ? savedOptions.directionAxis : ARTBOARD_DIRECTION_AXIS;
            directionDownRadio.value = (savedDirectionAxis === 1);
            directionRightRadio.value = !directionDownRadio.value;

            /**
             * 選ばれている作成対象を読み取ります。
             *
             * @returns {string} "layer" / "artboard" / "document"。
             */
            function readCreateTarget() {
                if (artboardRow.radio.value) return "artboard";
                if (documentRow.radio.value) return "document";
                return "layer";
            }

            /**
             * 名前欄とオプションの有効・無効を作成対象に合わせて切り替えます。
             *
             * 名前欄は3つとも並べたまま、対象の行だけを有効にします。
             * ［元のオブジェクトを残す］はドキュメント作成では常に複製になるため、
             * ［アートボードの配置］はアートボード作成のときだけ意味があるので、
             * それぞれ対象外のときはディム表示にします。
             * ［ロック／非表示オブジェクト］はレイヤーとアートボードでは対象範囲を
             * 定義できないため、ドキュメント作成のときだけ有効にします。
             *
             * @returns {void}
             */
            function syncOptionStates() {
                var createTarget = readCreateTarget();
                var isDocumentTarget = (createTarget === "document");

                /* 名前欄は行ごとに、対象のものだけ有効にする / Enable the matching row only */
                layerRow.input.enabled = (createTarget === "layer");
                artboardRow.input.enabled = (createTarget === "artboard");
                documentRow.input.enabled = isDocumentTarget;

                /* 配置はアートボード作成のときだけ意味がある
                   Placement only applies when creating an artboard */
                var isArtboardTarget = (createTarget === "artboard");
                directionLabel.enabled = isArtboardTarget;
                directionRightRadio.enabled = isArtboardTarget;
                directionDownRadio.enabled = isArtboardTarget;
                spacingLabel.enabled = isArtboardTarget;
                spacingInput.enabled = isArtboardTarget;
                spacingStepper.enabled = isArtboardTarget;
                redrawSteppersIn(spacingStepper);
                spacingUnitLabel.enabled = isArtboardTarget;

                duplicateCheckbox.enabled = !isDocumentTarget;
                includeLockedCheckbox.enabled = isDocumentTarget;
                includeHiddenCheckbox.enabled = isDocumentTarget;
            }

            /**
             * 作成対象を1つだけ選ばれた状態にします。
             *
             * ScriptUIのラジオボタンは同じ親コンテナ内でしか排他になりません。
             * ここでは行ごとにグループを分けているため、自前で他をオフにします。
             *
             * @param {object} activeRow - 選択された行（addTargetRow の戻り値）。
             * @returns {void}
             */
            function selectCreateTarget(activeRow) {
                layerRow.radio.value = (activeRow === layerRow);
                artboardRow.radio.value = (activeRow === artboardRow);
                documentRow.radio.value = (activeRow === documentRow);
                syncOptionStates();
            }

            layerRow.radio.onClick = function () { selectCreateTarget(layerRow); };
            artboardRow.radio.onClick = function () { selectCreateTarget(artboardRow); };
            documentRow.radio.onClick = function () { selectCreateTarget(documentRow); };

            var rowsByTarget = { layer: layerRow, artboard: artboardRow, document: documentRow };
            var savedTarget = (savedOptions !== null) ? savedOptions.createTarget : "layer";
            var initialRow = rowsByTarget[savedTarget] || layerRow;
            selectCreateTarget(initialRow);

            /* L / A / D で作成対象を切り替える。名前欄の入力中は文字として入れ、間隔の数値欄では効かせる
               L / A / D switch the target. They type normally in the name fields but still work in the spacing field */
            var targetShortcutMap = {};
            for (var shortcutKey in SHORTCUT_TARGETS) {
                if (SHORTCUT_TARGETS.hasOwnProperty(shortcutKey)) {
                    targetShortcutMap[shortcutKey] = rowsByTarget[SHORTCUT_TARGETS[shortcutKey]].radio;
                }
            }
            addKeyShortcuts(dialog, targetShortcutMap, { numericFields: [spacingInput] });

            /* ショートカットが最初から効くよう、フォーカスはラジオボタンに置く
               Focus a radio so the shortcuts work without clicking first */
            dialog.onShow = function () {
                initialRow.radio.active = true;
            };

            /* ボタン行 / Button row */
            var buttonRow = addButtonRow(dialog);
            var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
            var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
            btnCancel.helpTip = getLabel(LABELS.tip.cancel);
            btnOK.helpTip = getLabel(LABELS.tip.ok);

            prepareDialogWindow(dialog, SCRIPT_NAME);
            if (dialog.show() !== 1) return null;

            var createTarget = readCreateTarget();

            /* 空欄のときは初期値に戻す / Fall back to the default when left blank */
            var createName = rowsByTarget[createTarget].input.text.replace(/^\s+|\s+$/g, "");
            if (createName === "") createName = defaultNames[createTarget];

            /* 次回のために保持する。ディム中の値もそのまま覚えておかないと、
               作成対象を切り替えて戻したときにチェックが消えてしまう
               Remember the raw values: zeroing the dimmed ones would clear the
               checkboxes when the user switches targets and comes back */
            var directionAxis = directionDownRadio.value ? 1 : 0;

            /* 入力が空・不正なら推定値に戻す / Fall back to the inferred value on blank or bad input */
            var spacingValue = parseFloat(spacingInput.text);
            var spacingPt = isNaN(spacingValue) ? autoSpacingPt : (spacingValue * rulerUnit.pointsPerUnit);

            /* 間隔はドキュメントごとに変わるので保持しない
               The spacing is document-specific, so it is not remembered */
            setSessionCreateOptions({
                createTarget: createTarget,
                directionAxis: directionAxis,
                useDuplicate: duplicateCheckbox.value,
                includeLocked: includeLockedCheckbox.value,
                includeHidden: includeHiddenCheckbox.value
            });

            /* ディム中の値は無視する / Ignore the values while the controls are dimmed */
            return {
                createTarget: createTarget,
                createName: createName,
                directionAxis: directionAxis,
                spacingPt: spacingPt,
                useDuplicate: (duplicateCheckbox.enabled && duplicateCheckbox.value),
                includeLocked: (includeLockedCheckbox.enabled && includeLockedCheckbox.value),
                includeHidden: (includeHiddenCheckbox.enabled && includeHiddenCheckbox.value)
            };
        }

        // =========================================
        // レイヤーを作成 / Create a layer
        // =========================================

        /**
         * 同じ名前のレイヤーがすでに存在するかを判定します。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {string} layerName - 探すレイヤー名。
         * @returns {boolean} 存在する場合は true。
         */
        function layerNameExists(doc, layerName) {
            /* 最上位レイヤーだけを見る / Only top-level layers are checked */
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === layerName) return true;
            }
            return false;
        }

        /**
         * 重複しないレイヤー名を作ります（「新規レイヤー 2」のように連番を付与）。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {string} baseName - 基準となるレイヤー名。
         * @returns {string} 重複しないレイヤー名。
         */
        function makeUniqueLayerName(doc, baseName) {
            if (!layerNameExists(doc, baseName)) return baseName;

            /* 空いている連番を探す / Look for a free sequence number */
            var suffixNumber = 2;
            while (layerNameExists(doc, baseName + " " + suffixNumber)) {
                suffixNumber++;
            }
            return baseName + " " + suffixNumber;
        }

        /**
         * 選択オブジェクトを収めた新規レイヤーを作成します。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
         * @param {string} layerName - 新規レイヤーの名前。
         * @returns {void}
         */
        function createLayerFromSelection(doc, useDuplicate, layerName) {
            /* ループ前にスナップショットを取る（move で選択が変化するため）
               Snapshot before the loop: move() changes the live selection */
            var selectedItems = snapshotSelection(doc);
            if (selectedItems.length === 0) {
                alert(getLabel(LABELS.alert.noSelection));
                return;
            }

            selectedItems = sortByStackOrder(selectedItems);

            /* 最前面に新規レイヤーを作る / Add the new layer at the top */
            var newLayer = doc.layers.add();
            newLayer.name = makeUniqueLayerName(doc, layerName);

            /* 前面のものから PLACEATEND で送ると元の重ね順が保たれる
               Sending front-to-back with PLACEATEND preserves the original order */
            var placedItems = [];
            for (var i = 0; i < selectedItems.length; i++) {
                if (useDuplicate) {
                    placedItems.push(selectedItems[i].duplicate(newLayer, ElementPlacement.PLACEATEND));
                } else {
                    selectedItems[i].move(newLayer, ElementPlacement.PLACEATEND);
                    placedItems.push(selectedItems[i]);
                }
            }

            /* 新規レイヤー側のオブジェクトを選択状態にする
               Select the items that ended up on the new layer */
            doc.selection = placedItems;
        }

        // =========================================
        // アートボードの配置計算 / Artboard layout math
        // =========================================

        /**
         * 軸方向のアートボードサイズを取得します。
         *
         * @param {Array<number>} artboardRect - artboardRect（[左, 上, 右, 下]）。
         * @param {number} axisIndex - 0 なら幅、1 なら高さ。
         * @returns {number} 指定軸のサイズ。
         */
        function getArtboardAxisSize(artboardRect, axisIndex) {
            /* Y軸は上→下なので絶対値をとる / The Y axis is top-down, so take the absolute value */
            return (axisIndex === 0)
                ? (artboardRect[2] - artboardRect[0])
                : Math.abs(artboardRect[3] - artboardRect[1]);
        }

        /**
         * 隣り合う2枚のアートボードから主軸を判定します。
         *
         * @param {Array<number>} firstRect - 1枚目の artboardRect。
         * @param {Array<number>} secondRect - 2枚目の artboardRect。
         * @returns {number} 左端が同じなら 1（縦並び）、違えば 0（横並び）。
         */
        function detectPrimaryAxisIndex(firstRect, secondRect) {
            /* 左端が揃っていれば縦に積まれている / A shared left edge means a vertical stack */
            return (firstRect[0] === secondRect[0]) ? 1 : 0;
        }

        /**
         * 既存アートボードの並びから現在の間隔（pt）を推定します。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {number} fallbackSpacingPt - 2枚未満のときに使う値（pt）。
         * @returns {number} 推定した間隔（pt）。
         */
        function computeAutoSpacingPt(doc, fallbackSpacingPt) {
            var artboardList = doc.artboards;
            if (artboardList.length < 2) return fallbackSpacingPt;

            /* ピッチからアートボード自身のサイズを引いた残りが隙間
               The gap is the pitch minus the artboard's own size */
            var firstRect = artboardList[0].artboardRect;
            var secondRect = artboardList[1].artboardRect;
            var primaryAxisIndex = detectPrimaryAxisIndex(firstRect, secondRect);
            var pitch = Math.abs(secondRect[primaryAxisIndex] - firstRect[primaryAxisIndex]);
            var spacing = pitch - getArtboardAxisSize(firstRect, primaryAxisIndex);
            return (spacing < 0) ? 0 : spacing;
        }

        /**
         * 最大カンバス範囲を取得します。
         *
         * Original idea by OMOTI
         * https://forums.adobe.com/thread/2459293
         *
         * @param {Document} doc - 対象のドキュメント。
         * @returns {Array<number>} [左, 上, 右, 下]。
         */
        function getLargestCanvasBounds(doc) {
            var LARGEST_SIZE = 16383;

            /* 一時テキストの変換行列からカンバス左上を得る
               Read the canvas origin from a temporary text frame's matrix */
            var tempLayer = doc.layers.add();
            var tempText = tempLayer.textFrames.add();
            var canvasLeft = tempText.matrix.mValueTX;
            var canvasTop = tempText.matrix.mValueTY;
            tempLayer.remove();

            return [canvasLeft, canvasTop, canvasLeft + LARGEST_SIZE, canvasTop - LARGEST_SIZE];
        }

        /**
         * 既存の並びから引き継げるグリッドを検出します。
         *
         * 既存の並びが指定方向と一致する場合だけ、そのピッチ（gridStep）と
         * 折り返し位置（columns）を読み取ります。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {Array<number>} gridStep - 既定のグリッド移動量 [X, Y]。破壊的に更新します。
         * @param {number} primaryAxisIndex - 主軸（0=横 / 1=縦）。
         * @param {number} secondaryAxisIndex - 副軸。
         * @returns {{canInherit: boolean, columns: number}} 検出結果。
         */
        function detectInheritedGrid(doc, gridStep, primaryAxisIndex, secondaryAxisIndex) {
            var artboards = doc.artboards;
            if (artboards.length < 2) return { canInherit: false, columns: 0 };

            var firstRect = artboards[0].artboardRect;
            var secondRect = artboards[1].artboardRect;
            if (detectPrimaryAxisIndex(firstRect, secondRect) !== primaryAxisIndex) {
                return { canInherit: false, columns: 0 };
            }

            /* 実際のピッチを主軸の移動量として採用 / Adopt the real pitch as the primary step */
            gridStep[primaryAxisIndex] = secondRect[primaryAxisIndex] - firstRect[primaryAxisIndex];

            /* 副軸の座標が変わる位置が折り返し＝列数 / The wrap point gives the column count */
            for (var i = 2; i < artboards.length; i++) {
                var scannedRect = artboards[i].artboardRect;
                if (firstRect[secondaryAxisIndex] !== scannedRect[secondaryAxisIndex]) {
                    gridStep[secondaryAxisIndex] = scannedRect[secondaryAxisIndex] - firstRect[secondaryAxisIndex];
                    return { canInherit: true, columns: i };
                }
            }
            return { canInherit: true, columns: 0 };
        }

        /**
         * カンバスに収まるグリッドの行数・列数を求めます。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {Array<number>} firstRect - 先頭アートボードの artboardRect。
         * @param {Array<number>} gridStep - グリッド1セルあたりの移動量 [X, Y]。
         * @param {number} primaryAxisIndex - 主軸（0=横 / 1=縦）。
         * @param {number} secondaryAxisIndex - 副軸。
         * @returns {{columns: number, rows: number}} 収まる列数と行数。
         */
        function countGridCapacity(doc, firstRect, gridStep, primaryAxisIndex, secondaryAxisIndex) {
            var canvasRect = getLargestCanvasBounds(doc);

            /* 1セル分の枠をカンバスの端と比べて、何個並ぶかを数える
               Count how many cells fit between one cell and the canvas edge */
            var gridUnitRect = [
                firstRect[0] + Math.abs(gridStep[0]),
                firstRect[1] - Math.abs(gridStep[1]),
                firstRect[0],
                firstRect[1]
            ];

            var primaryEdgeIndex =
                (primaryAxisIndex ^ +(gridStep[primaryAxisIndex] < 0)) ? primaryAxisIndex : primaryAxisIndex + 2;
            var secondaryEdgeIndex =
                (secondaryAxisIndex ^ +(gridStep[secondaryAxisIndex] < 0)) ? secondaryAxisIndex : secondaryAxisIndex + 2;

            return {
                columns: Math.abs(Math.floor(
                    (canvasRect[primaryEdgeIndex] - gridUnitRect[primaryEdgeIndex]) / gridStep[primaryAxisIndex])),
                rows: Math.abs(Math.floor(
                    (canvasRect[secondaryEdgeIndex] - gridUnitRect[secondaryEdgeIndex]) / gridStep[secondaryAxisIndex]))
            };
        }

        /**
         * アートボードの配置計画。
         *
         * @typedef {object} ArtboardLayout
         * @property {boolean} canInherit - 既存の並びのグリッドを引き継げる場合は true。
         * @property {Array<number>} gridStep - グリッド1セルあたりの移動量 [X, Y]。
         * @property {number} columns - グリッドの列数。
         * @property {number} primaryAxisIndex - 主軸（0=横 / 1=縦）。
         * @property {number} secondaryAxisIndex - 副軸。
         * @property {number} primarySign - 主軸方向の符号（+1 / -1）。
         * @property {number} spacing - アートボード間の間隔（pt）。
         * @property {Array<number>} firstRect - 先頭アートボードの artboardRect。
         * @property {Array<number>} referenceRect - サイズの基準にする artboardRect。
         */

        /**
         * 既存の並びを解析し、新規アートボードの配置計画を作ります。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {number} referenceIndex - サイズの基準にするアートボードのインデックス。
         * @param {number} directionAxis - 並べる方向（0=右 / 1=下）。ダイアログで指定します。
         * @param {number} spacing - アートボードどうしの間隔（pt）。ダイアログで指定します。
         * @returns {ArtboardLayout|null} 配置計画。スペースが足りない場合は null。
         */
        function planArtboardLayout(doc, referenceIndex, directionAxis, spacing) {
            var artboards = doc.artboards;
            var firstRect = artboards[0].artboardRect;
            var primaryAxisIndex = directionAxis;
            var secondaryAxisIndex = 1 - primaryAxisIndex;

            /* グリッド1セルあたりの移動量（Y軸は上→下なので間隔を引く）
               Grid step per cell (the Y axis is top-down, so the spacing is subtracted) */
            var gridStep = [
                (firstRect[2] - firstRect[0]) + spacing,
                (firstRect[3] - firstRect[1]) - spacing
            ];

            var inheritedGrid = detectInheritedGrid(doc, gridStep, primaryAxisIndex, secondaryAxisIndex);
            var capacity = countGridCapacity(doc, firstRect, gridStep, primaryAxisIndex, secondaryAxisIndex);

            /* 既存の並びから列数が読めたときはそちらを優先
               Prefer the column count read from the existing layout */
            var columns = inheritedGrid.columns || capacity.columns;
            if (artboards.length + 1 > columns * capacity.rows) return null;

            return {
                canInherit: inheritedGrid.canInherit,
                gridStep: gridStep,
                columns: columns,
                primaryAxisIndex: primaryAxisIndex,
                secondaryAxisIndex: secondaryAxisIndex,
                primarySign: (gridStep[primaryAxisIndex] < 0) ? -1 : 1,
                spacing: spacing,
                firstRect: firstRect,
                referenceRect: artboards[referenceIndex].artboardRect
            };
        }

        /**
         * グリッド上のインデックスに対応する位置を求めます。
         *
         * @param {ArtboardLayout} artboardLayout - 配置計画。
         * @param {number} gridIndex - グリッド上のインデックス。
         * @returns {Array<number>} [左, 上] の座標。
         */
        function getArtboardGridPosition(artboardLayout, gridIndex) {
            /* 主軸は列内の位置、副軸は何行目かで決まる
               The primary axis gives the column, the secondary axis the row */
            var offset = [];
            offset[artboardLayout.primaryAxisIndex] =
                (gridIndex % artboardLayout.columns) * artboardLayout.gridStep[artboardLayout.primaryAxisIndex];
            offset[artboardLayout.secondaryAxisIndex] =
                Math.floor(gridIndex / artboardLayout.columns) * artboardLayout.gridStep[artboardLayout.secondaryAxisIndex];

            return [artboardLayout.firstRect[0] + offset[0], artboardLayout.firstRect[1] + offset[1]];
        }

        /**
         * グリッドを引き継げないときの配置位置を求めます。
         *
         * 先頭ではなく、挿入位置の直前のアートボードを基準に指定方向へ1枚分進めます。
         * グリッドのインデックスを歩数に使うと、既存の並びと軸が違う場合に
         * 枚数分だけ離れた位置へ飛んでしまうため、実位置から積み上げます。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {ArtboardLayout} artboardLayout - 配置計画。
         * @param {number} anchorIndex - 基準にするアートボードのインデックス。
         * @returns {Array<number>} [左, 上] の座標。
         */
        function getAnchoredArtboardPosition(doc, artboardLayout, anchorIndex) {
            /* 基準アートボードの実位置から1枚分だけ進める / Step one artboard from the anchor */
            var anchorRect = doc.artboards[anchorIndex].artboardRect;
            var advance = getArtboardAxisSize(anchorRect, artboardLayout.primaryAxisIndex) + artboardLayout.spacing;
            var position = [anchorRect[0], anchorRect[1]];
            position[artboardLayout.primaryAxisIndex] += artboardLayout.primarySign * advance;
            return position;
        }

        // =========================================
        // アートボードを作成 / Create an artboard
        // =========================================

        /**
         * アートボード上のアイテムを収集します。
         *
         * doc.pageItems はグループ・複合パスの子まで再帰的に含むため、最上位
         * （親がレイヤー）のアイテムのみを対象にします。子は親と一緒に動くので、
         * ここで拾うと二重に処理されてしまいます。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {Array<number>} artboardRect - 対象アートボードの矩形。
         * @returns {Array<PageItem>} アートボードに属するアイテム。
         */
        function getItemsAssignedToArtboard(doc, artboardRect) {
            /* 重心がアートボード内にあるものを所属とみなす
               An item belongs to the artboard when its center falls inside */
            var assignedItems = [];

            for (var i = 0; i < doc.pageItems.length; i++) {
                var item = doc.pageItems[i];
                if (item.parent.typename !== 'Layer') continue;

                if (isPointInsideRect(getRectCenter(item.geometricBounds), artboardRect)) {
                    assignedItems.push(item);
                }
            }
            return assignedItems;
        }

        /**
         * アイテムと祖先レイヤーのロック／表示状態を一時的に解除します。
         *
         * ロック／非表示のアイテムは translate() が例外になります。途中で例外になると
         * そこまで動かしたアートボードだけが残って崩れるため、一時解除してから処理します。
         *
         * @param {PageItem} item - 対象のオブジェクト。
         * @returns {Array<object>} 復元用の情報リスト。
         */
        function unlockItemTemporarily(item) {
            /* 祖先レイヤーから順に解除し、戻すための記録を残す
               Clear the ancestors first, recording what to restore */
            var restoreList = [];
            var ancestorLayer = item.parent;

            while (ancestorLayer && ancestorLayer.typename === 'Layer') {
                if (ancestorLayer.locked) {
                    ancestorLayer.locked = false;
                    restoreList.push({ target: ancestorLayer, property: 'locked', value: true });
                }
                if (!ancestorLayer.visible) {
                    ancestorLayer.visible = true;
                    restoreList.push({ target: ancestorLayer, property: 'visible', value: false });
                }
                ancestorLayer = ancestorLayer.parent;
            }
            if (item.locked) {
                item.locked = false;
                restoreList.push({ target: item, property: 'locked', value: true });
            }
            if (item.hidden) {
                item.hidden = false;
                restoreList.push({ target: item, property: 'hidden', value: true });
            }
            return restoreList;
        }

        /**
         * 一時解除したロック／表示状態を元に戻します。
         *
         * @param {Array<object>} restoreList - unlockItemTemporarily の戻り値。
         * @returns {void}
         */
        function restoreLockAndVisibility(restoreList) {
            /* 解除と逆順に戻す（内側→外側）/ Restore in reverse order (inner → outer) */
            for (var i = restoreList.length - 1; i >= 0; i--) {
                restoreList[i].target[restoreList[i].property] = restoreList[i].value;
            }
        }

        /**
         * 挿入位置以降のアートボードを1枚分だけ後ろへずらし、アートワークも一緒に運びます。
         *
         * アイテムの帰属は移動前の位置でまとめて取得（スナップショット）してから動かすため、
         * 移動順による取り違えが起きません。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {ArtboardLayout} artboardLayout - 配置計画。
         * @param {number} fromIndex - ずらし始めるインデックス。
         * @returns {void}
         */
        function relayoutExistingArtboards(doc, artboardLayout, fromIndex) {
            var artboards = doc.artboards;

            /* 先に移動計画を立てる（動かしながら判定すると帰属がずれる）
               Plan every move first: judging while moving would misassign items */
            var plannedMoves = [];
            for (var i = fromIndex; i < artboards.length; i++) {
                var currentRect = artboards[i].artboardRect;
                var targetPosition = getArtboardGridPosition(artboardLayout, i + 1);
                plannedMoves.push({
                    index: i,
                    dx: targetPosition[0] - currentRect[0],
                    dy: targetPosition[1] - currentRect[1],
                    items: getItemsAssignedToArtboard(doc, currentRect)
                });
            }

            /* アートワークを動かしてから、アートボード自体を動かす
               Move the artwork, then the artboard itself */
            for (var j = 0; j < plannedMoves.length; j++) {
                var plannedMove = plannedMoves[j];

                for (var k = 0; k < plannedMove.items.length; k++) {
                    var restoreList = unlockItemTemporarily(plannedMove.items[k]);
                    try {
                        plannedMove.items[k].translate(plannedMove.dx, plannedMove.dy);
                    } finally {
                        restoreLockAndVisibility(restoreList);
                    }
                }

                var movedRect = artboards[plannedMove.index].artboardRect;
                artboards[plannedMove.index].artboardRect = [
                    movedRect[0] + plannedMove.dx, movedRect[1] + plannedMove.dy,
                    movedRect[2] + plannedMove.dx, movedRect[3] + plannedMove.dy
                ];
            }
        }

        /**
         * 末尾に追加されたアートボードを、パネル上の挿入位置へ並べ替えます。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {number} insertIndex - 挿入位置のインデックス。
         * @param {number} originalCount - 追加前のアートボード数。
         * @returns {void}
         */
        function reorderAppendedArtboard(doc, insertIndex, originalCount) {
            var artboards = doc.artboards;

            /* 追加分の枠と名前を退避（後続シフトで上書きされる前に）
               Save the appended rect and name before the shift clobbers them */
            var appendedRect = artboards[originalCount].artboardRect;
            var appendedName = artboards[originalCount].name;

            /* 後続を1枚分だけ後ろへ（高位から処理して上書き衝突を回避）
               Shift trailing artboards back by one (high → low to avoid clobbering) */
            for (var i = originalCount - 1; i >= insertIndex; i--) {
                artboards[i + 1].artboardRect = artboards[i].artboardRect;
                artboards[i + 1].name = artboards[i].name;
            }

            artboards[insertIndex].artboardRect = appendedRect;
            artboards[insertIndex].name = appendedName;
        }

        /**
         * 選択オブジェクトを指定量だけ移動します。複製が指定された場合は複製側を移動します。
         *
         * @param {Array<PageItem>} items - 対象のオブジェクト。
         * @param {number} dx - X方向の移動量。
         * @param {number} dy - Y方向の移動量。
         * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
         * @returns {Array<PageItem>} 移動先に置かれたオブジェクト。
         */
        function moveOrDuplicateItems(items, dx, dy, useDuplicate) {
            /* 複製時は元を残し、複製した側だけを動かす
               When duplicating, keep the originals and move only the copies */
            var placedItems = [];

            for (var i = 0; i < items.length; i++) {
                var targetItem = useDuplicate ? items[i].duplicate() : items[i];
                targetItem.translate(dx, dy);
                placedItems.push(targetItem);
            }
            return placedItems;
        }

        /**
         * 選択オブジェクトから新規アートボードを作成します。
         *
         * 新規アートボードのサイズはアクティブアートボードと同じで、選択オブジェクトは
         * 元のアートボード内での相対位置を保ったまま移動（または複製）されます。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
         * @param {string} artboardName - 新規アートボードの名前。
         * @param {number} directionAxis - 並べる方向（0=右 / 1=下）。
         * @param {number} spacingPt - アートボードどうしの間隔（pt）。
         * @returns {void}
         */
        function createArtboardFromSelection(doc, useDuplicate, artboardName, directionAxis, spacingPt) {
            var selectedItems = snapshotSelection(doc);
            if (selectedItems.length === 0) {
                alert(getLabel(LABELS.alert.noSelection));
                return;
            }

            /* 上限を超えるなら何も変更せずに終える / Bail out before touching anything */
            var artboards = doc.artboards;
            var originalArtboardCount = artboards.length;
            var artboardLimit = (parseFloat(app.version) >= 22) ? 1000 : 100;

            if (originalArtboardCount + 1 > artboardLimit) {
                alert(getLabel(LABELS.alert.artboardLimit));
                return;
            }

            var activeArtboardIndex = artboards.getActiveArtboardIndex();
            var insertArtboardIndex = ARTBOARD_INSERT_AFTER_CURRENT
                ? (activeArtboardIndex + 1)
                : originalArtboardCount;

            var artboardLayout = planArtboardLayout(doc, activeArtboardIndex, directionAxis, spacingPt);
            if (artboardLayout === null) {
                alert(getLabel(LABELS.alert.noSpace));
                return;
            }

            /* 基準になるアートボードと選択位置は、ずらす前に記録しておく
               Record the source artboard and the selection's position before anything shifts */
            var sourceArtboardIndex = findArtboardIndexForItems(doc, selectedItems, activeArtboardIndex);
            var sourceRect = artboards[sourceArtboardIndex].artboardRect;
            var boundsBeforeShift = getClipAwareUnionBounds(selectedItems, false);

            /* 既存アートボードを後ろへずらして挿入スペースを空ける。
               グリッドを引き継げないときは既存の並びを動かさない（別軸のグリッドへ
               流し込むと並びが崩れるため）
               Free the insert space. When the grid can't be inherited, leave the existing
               artboards alone — re-flowing onto the other axis would break the arrangement */
            if (artboardLayout.canInherit) {
                relayoutExistingArtboards(doc, artboardLayout, insertArtboardIndex);
            }

            /* 選択オブジェクトが実際にずれた量を測る。
               アートボード内に重心があるものだけがずれるため、ずれ幅は0のこともある。
               この実測値を引くことで、重心がアートボード外にあるオブジェクトでも
               元のアートボードに対する相対位置が正しく保たれる
               Measure how far the selection actually moved: only art whose center sits inside
               the artboard is shifted, so this can be zero. Subtracting the measured amount
               keeps the relative position correct even for art centered outside its artboard */
            var boundsAfterShift = getClipAwareUnionBounds(selectedItems, false);
            var shiftX = boundsAfterShift[0] - boundsBeforeShift[0];
            var shiftY = boundsAfterShift[1] - boundsBeforeShift[1];

            var newPosition = artboardLayout.canInherit
                ? getArtboardGridPosition(artboardLayout, insertArtboardIndex)
                : getAnchoredArtboardPosition(doc, artboardLayout, insertArtboardIndex - 1);

            /* アクティブアートボードと同じサイズで作る / Match the active artboard's size */
            var referenceRect = artboardLayout.referenceRect;
            artboards.add([
                newPosition[0],
                newPosition[1],
                newPosition[0] + (referenceRect[2] - referenceRect[0]),
                newPosition[1] + (referenceRect[3] - referenceRect[1])
            ]);

            /* パネル上の順序を挿入位置へ並べ替え / Reorder in the panel */
            reorderAppendedArtboard(doc, insertArtboardIndex, originalArtboardCount);
            artboards[insertArtboardIndex].name = artboardName;

            /* 相対位置を保ったまま新規アートボードへ（すでにずれた分は差し引く）
               Carry the relative position over, minus the shift already applied */
            var placedItems = moveOrDuplicateItems(
                selectedItems,
                newPosition[0] - sourceRect[0] - shiftX,
                newPosition[1] - sourceRect[1] - shiftY,
                useDuplicate);

            artboards.setActiveArtboardIndex(insertArtboardIndex);
            doc.selection = placedItems;
            app.redraw();
        }

        // =========================================
        // ドキュメントの走査 / Walking a document
        // =========================================

        /**
         * 進捗を示すパレットを開きます。
         *
         * ドキュメントの複製は保存・削除・再オープンと時間のかかる処理が続くため、
         * 無反応に見えないよう進捗を出します。
         *
         * @returns {{setMessage: function, close: function}} 進捗パレットの操作口。
         */
        function openProgressPalette() {
            var palette = new Window("palette", getLabel(LABELS.progress.title));
            setupWindow(palette);

            var messageText = palette.add("statictext", undefined, "");
            messageText.preferredSize.width = 280;
            palette.show();

            return {
                /**
                 * 表示するメッセージを差し替えます。
                 *
                 * @param {object} labelEntry - 表示する文言のラベル定義。
                 * @returns {void}
                 */
                setMessage: function (labelEntry) {
                    messageText.text = getLabel(labelEntry);
                    palette.update();
                },

                /**
                 * パレットを閉じます。
                 *
                 * @returns {void}
                 */
                close: function () {
                    palette.close();
                }
            };
        }

        /**
         * ファイル名から拡張子を除いた部分を取り出します。
         *
         * @param {string} fileName - ファイル名（拡張子込み）。
         * @returns {string} 拡張子を除いた部分。
         */
        function getFileBaseName(fileName) {
            /* 「my.logo.ai」のような名前でも最後のドットだけで切る
               Split at the last dot only, even for names like "my.logo.ai" */
            var dotIndex = fileName.lastIndexOf('.');
            return (dotIndex > 0) ? fileName.substring(0, dotIndex) : fileName;
        }

        /**
         * ファイル名から拡張子を取り出します。
         *
         * @param {string} fileName - ファイル名（拡張子込み）。
         * @returns {string} ドットを含む拡張子。拡張子がない場合は空文字。
         */
        function getFileExtension(fileName) {
            var dotIndex = fileName.lastIndexOf('.');
            return (dotIndex > 0) ? fileName.substring(dotIndex) : '';
        }

        /**
         * すべてのレイヤーのロック・非表示を一時的に解除します。
         *
         * 反転選択はロック・非表示のレイヤーを拾わないため、削除の前に解除します。
         * レイヤーの表示状態は成果物に残るので、あとで元に戻せるよう記録します。
         *
         * @param {Layers} layers - 対象のレイヤーコレクション。
         * @param {Array<object>} restoreList - 復元用の情報を追加する配列。
         * @returns {void}
         */
        function unlockLayersTemporarily(layers, restoreList) {
            /* サブレイヤーまで再帰的に解除する / Clear sub-layers recursively too */
            for (var i = 0; i < layers.length; i++) {
                restoreList.push({ layer: layers[i], locked: layers[i].locked, visible: layers[i].visible });
                layers[i].locked = false;
                layers[i].visible = true;
                unlockLayersTemporarily(layers[i].layers, restoreList);
            }
        }

        /**
         * 一時解除したレイヤーのロック・非表示を元に戻します。
         *
         * @param {Array<object>} restoreList - unlockLayersTemporarily で記録した情報。
         * @returns {void}
         */
        function restoreLayerStates(restoreList) {
            /* 解除と逆順に戻す（内側→外側）/ Restore in reverse order (inner → outer) */
            for (var i = restoreList.length - 1; i >= 0; i--) {
                restoreList[i].layer.locked = restoreList[i].locked;
                restoreList[i].layer.visible = restoreList[i].visible;
            }
        }

        /**
         * 記録済みのレイヤー状態から、指定レイヤーの分を探します。
         *
         * @param {Array<object>} layerStates - unlockLayersTemporarily で記録した情報。
         * @param {Layer} layer - 探すレイヤー。
         * @returns {object|null} 見つかった記録。見つからない場合は null。
         */
        function findLayerState(layerStates, layer) {
            /* レイヤー数はたかが知れているので線形探索で足りる
               Layer counts stay small, so a linear scan is enough */
            for (var i = 0; i < layerStates.length; i++) {
                if (layerStates[i].layer === layer) return layerStates[i];
            }
            return null;
        }

        /**
         * 反転削除の前に、残したいオブジェクトだけをロック／非表示のままにします。
         *
         * 反転選択はロック・非表示のオブジェクトを拾わないため、ロック／非表示のままに
         * したものは削除されずに残ります。逆に残さないものはここで解除しておきます。
         * レイヤー単位のロック・非表示も、そのオブジェクト自身の状態として扱います。
         *
         * 対象は最上位（親がレイヤー）のオブジェクトだけです。グループ内のオブジェクトは
         * ロック／非表示にしても、親グループが反転で選択されればまとめて削除されるため、
         * 残す保証ができません。状態には触れず、親の扱いに従わせます。
         *
         * @param {Document} targetDoc - 対象のドキュメント。
         * @param {Array<object>} layerStates - 解除前に記録したレイヤー状態。
         * @param {boolean} includeLocked - ロックされたオブジェクトを残す場合は true。
         * @param {boolean} includeHidden - 非表示オブジェクトを残す場合は true。
         * @returns {Array<object>} 残す印を付けたオブジェクトと、その元の状態。
         */
        function protectItemsFromDelete(targetDoc, layerStates, includeLocked, includeHidden) {
            /* コレクションはループ外に退避する（毎回DOMを経由するため）
               Hoist the collection: each access crosses the DOM bridge */
            var pageItems = targetDoc.pageItems;
            var protectedEntries = [];

            for (var i = 0; i < pageItems.length; i++) {
                var item = pageItems[i];

                /* グループ・複合パスの中身は親の扱いに任せる
                   Leave the contents of groups and compound paths to their parent */
                if (item.parent.typename !== 'Layer') continue;

                var layerState = findLayerState(layerStates, item.layer);
                var ownLocked = item.locked;
                var ownHidden = item.hidden;
                var wasLocked = ownLocked || (layerState !== null && layerState.locked);
                var wasHidden = ownHidden || (layerState !== null && !layerState.visible);

                var keepLocked = includeLocked && wasLocked;
                var keepHidden = includeHidden && wasHidden;

                if (keepLocked || keepHidden) {
                    protectedEntries.push({ item: item, locked: ownLocked, hidden: ownHidden });
                }

                /* 変化するときだけ書き込む / Write only when the value actually changes */
                if (ownLocked !== keepLocked) item.locked = keepLocked;
                if (ownHidden !== keepHidden) item.hidden = keepHidden;
            }
            return protectedEntries;
        }

        /**
         * 残したオブジェクトを、アートボードの内外で仕分けます。
         *
         * アートボード内のものは自身のロック・非表示状態を元に戻し、
         * 外にあるものは削除します。
         *
         * @param {Array<object>} protectedEntries - protectItemsFromDelete の戻り値。
         * @param {Array<number>} artboardRect - 残すアートボードの矩形。
         * @returns {void}
         */
        function finalizeProtectedItems(protectedEntries, artboardRect) {
            for (var i = 0; i < protectedEntries.length; i++) {
                var entry = protectedEntries[i];

                /* 親グループごと削除されている場合があるため保護
                   The parent group may already be gone, so guard the access */
                try {
                    if (isPointInsideRect(getRectCenter(entry.item.geometricBounds), artboardRect)) {
                        entry.item.locked = entry.locked;
                        entry.item.hidden = entry.hidden;
                    } else {
                        entry.item.locked = false;
                        entry.item.hidden = false;
                        entry.item.remove();
                    }
                } catch (e) {}
            }
        }

        // =========================================
        // ドキュメントを作成 / Create a document
        // =========================================

        /**
         * 指定した1枚を除いて、すべてのアートボードを削除します。
         *
         * @param {Document} targetDoc - 対象のドキュメント。
         * @param {number} keepIndex - 残すアートボードのインデックス。
         * @returns {void}
         */
        function keepOnlyArtboard(targetDoc, keepIndex) {
            /* 降順に処理するので、削除しても未処理側のインデックスはずれない
               A descending loop keeps the pending indexes valid */
            var artboards = targetDoc.artboards;

            for (var i = artboards.length - 1; i >= 0; i--) {
                if (i !== keepIndex) artboards.remove(i);
            }
        }

        /**
         * 選択オブジェクトから新規ドキュメントを作成します。
         *
         * 別名保存で複製ドキュメントを作り、複製側で選択範囲を反転して削除し、
         * 現在のアートボード以外を削除します。スウォッチ・シンボル・ドキュメント設定は
         * そのまま引き継がれます。
         *
         * `saveAs()` は開いているドキュメント自体を保存先に紐づけ直すため、複製側に
         * 選択状態がそのまま残ります。これを使うと「選択以外」をIllustrator自身の
         * 反転選択で求められるので、オブジェクトの突き合わせが不要になります。
         * 元ファイルはディスク上に保存時のまま残るので、最後に開き直します。
         *
         * @param {Document} doc - 対象のドキュメント。
         * @param {boolean} includeLocked - ロックされたオブジェクトも残す場合は true。
         * @param {boolean} includeHidden - 非表示オブジェクトも残す場合は true。
         * @param {string} fileBaseName - 複製ドキュメントのファイル名（拡張子なし）。
         * @returns {void}
         */
        function createDocumentFromSelection(doc, includeLocked, includeHidden, fileBaseName) {
            if (snapshotSelection(doc).length === 0) {
                alert(getLabel(LABELS.alert.noSelection));
                return;
            }

            /* 未保存・変更ありだと、開き直したときに変更が失われる
               Unsaved changes would be lost when the source document is reopened */
            var originalFile = doc.saved ? doc.fullName : null;
            if (originalFile === null || !originalFile.exists) {
                alert(getLabel(LABELS.alert.needsSave));
                return;
            }

            /* 保存先は元ファイルと同じ場所、拡張子も元のまま
               The duplicate sits next to the source, keeping its extension */
            var duplicateFile = new File(
                originalFile.parent.fsName + "/" + fileBaseName + getFileExtension(originalFile.name));

            /* 元ファイルに上書きすると、開き直す先が複製になってしまう
               Overwriting the source would leave nothing to reopen */
            if (duplicateFile.fsName === originalFile.fsName) {
                alert(getLabel(LABELS.alert.sameAsSource));
                return;
            }
            if (duplicateFile.exists && !confirm(getLabel(LABELS.alert.overwrite) + "\n" + duplicateFile.name, true)) {
                return;
            }

            var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();

            /* 保存・削除・再オープンと時間がかかるので、進捗を出しておく
               Saving, deleting and reopening all take time, so show the progress */
            var progressPalette = openProgressPalette();

            /* 何があってもパレットは閉じ、元ドキュメントは必ず開き直す
               Always close the palette and reopen the source document */
            var layerRestoreList = [];
            try {
                progressPalette.setMessage(LABELS.progress.duplicating);

                /* 別名保存すると、このドキュメント自体が複製ファイルに紐づく（選択状態はそのまま）
                   Saving as rebinds this very document to the duplicate, selection intact */
                doc.saveAs(duplicateFile);

                progressPalette.setMessage(LABELS.progress.deleting);

                /* オブジェクト単位で状態を扱えるよう、レイヤーはいったんすべて解除する
                   Clear every layer first so each item can be addressed individually */
                unlockLayersTemporarily(doc.layers, layerRestoreList);

                /* 残すものだけロック／非表示のままにする（反転が拾わない＝削除されない）
                   Keep only the survivors locked or hidden: Inverse skips them */
                var protectedEntries = protectItemsFromDelete(doc, layerRestoreList, includeLocked, includeHidden);

                /* メニューコマンドはアクティブドキュメントに働くので、対象を明示しておく
                   Menu commands act on the active document, so make the target explicit */
                app.activeDocument = doc;

                /* 選択範囲を反転して削除 / Invert the selection and delete it */
                app.executeMenuCommand('Inverse menu item');
                app.executeMenuCommand('clear');

                keepOnlyArtboard(doc, activeArtboardIndex);

                /* 残したもののうち、アートボード外にあるものは削除する
                   Drop the survivors that ended up outside the artboard */
                finalizeProtectedItems(protectedEntries, doc.artboards[0].artboardRect);

                /* レイヤーの表示・ロック状態は元ドキュメントのまま残す
                   Leave the layers' visibility and lock state as they were */
                restoreLayerStates(layerRestoreList);
            } finally {
                /* saveAs より前で失敗した場合は複製に紐づいていないので開き直さない
                   Nothing to reopen when the failure happened before saveAs */
                if (doc.fullName.fsName === duplicateFile.fsName) {
                    progressPalette.setMessage(LABELS.progress.reopening);
                    app.open(originalFile);
                }
                progressPalette.close();
            }

            app.activeDocument = doc;
            app.redraw();
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        /**
         * ダイアログで作成対象を選び、対応する処理を実行します。
         *
         * @returns {void}
         */
        function main() {
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }

            var doc = app.activeDocument;

            /* キャンセルされたら何もしない / Do nothing when cancelled */
            var createOptions = showCreateTargetDialog(doc);
            if (createOptions === null) return;

            /* 選ばれた作成対象へ振り分ける / Dispatch to the chosen target */
            if (createOptions.createTarget === "artboard") {
                createArtboardFromSelection(doc, createOptions.useDuplicate, createOptions.createName,
                    createOptions.directionAxis, createOptions.spacingPt);
            } else if (createOptions.createTarget === "document") {
                createDocumentFromSelection(doc, createOptions.includeLocked, createOptions.includeHidden,
                    createOptions.createName);
            } else {
                createLayerFromSelection(doc, createOptions.useDuplicate, createOptions.createName);
            }
        }

        /* 想定外のエラーはここで受け止める。個々の処理には try を置かない方針なので、
           捕まえ損ねるとExtendScriptの生のエラーダイアログが出てしまう
           Catch the unexpected here: the individual steps deliberately avoid try blocks,
           so anything uncaught would surface as a raw ExtendScript error dialog */
        try {
            main();
        } catch (err) {
            alert(getLabel(LABELS.alert.unexpected) + err);
        }

    })();

})();
