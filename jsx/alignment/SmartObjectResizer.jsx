#targetengine "SOR_Engine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、最大／最小／キーオブジェクト／指定サイズ／基準辺／面積／アートボード／裁ち落としのいずれかの基準でリサイズし、あわせて横位置・縦位置の整列も行えます。
縦横比保持と片辺のみを切り替えでき、操作はリアルタイムにプレビューされます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectResizer.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n6f35bd4000ec

### Overview

Resizes the selected objects to the largest, the smallest, the key object, a given size, a reference edge, an area, the artboard, or the bleed, and can align them horizontally and vertically at the same time.
You can switch between keeping the aspect ratio and constraining a single edge, with a real-time preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectResizer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectResizer";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectResizer.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectResizer.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6f35bd4000ec"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考 / Reference
 * キーオブジェクトの取得方法
 * 自分用メモ (@mute_racoon3631)
 * https://note.com/mute_racoon3631/n/n5dfae854988a
 *
 * ロゴなどの大きさ調整（面積を使うアイデア）
 * Gorolib Design
 * https://gorolib.blog.jp/archives/75031515.html
 */

(function () {
    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var BLEED_OFFSET_MM = 6;                 /* 裁ち落とし: 片側3mm＝幅/高さそれぞれ+6mm / bleed 3mm per side */

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

    var LABEL_WIDTH    = 122;                /* 行ラベルの共通幅（右揃え）/ shared row-label width (right-aligned) */

    /**
     * 縦方向のスペーサー（区切り用の空グループ）を追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {number} height - 高さ（px）
     * @returns {Group} 追加したスペーサー
     */
    function addSpacer(parentContainer, height) {
        var spacerGroup = parentContainer.add("group");
        spacerGroup.minimumSize.height = height;
        spacerGroup.maximumSize.height = height;
        return spacerGroup;
    }

    /**
     * helpTip（ツールチップ）を1コントロールまたはコントロール配列へ設定する
     * @param {Object|Object[]} targetControls - コントロール、またはその配列
     * @param {string} tipText - ツールチップの文言
     * @returns {void}
     */
    function setHelpTip(targetControls, tipText) {
        if (!tipText || !targetControls) return;
        if (typeof targetControls !== "string" && typeof targetControls.length === "number") {
            for (var i = 0; i < targetControls.length; i++) {
                if (targetControls[i]) targetControls[i].helpTip = tipText;
            }
        } else {
            targetControls.helpTip = tipText;
        }
    }

    /**
     * 共通幅で右揃えの行ラベルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} rowLabelText - ラベル文字列（空文字なら幅だけのスペーサーになる）
     * @returns {StaticText} 追加したラベル
     */
    function addRowLabel(parentGroup, rowLabelText) {
        var rowLabel = parentGroup.add("statictext", undefined, rowLabelText);
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
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

    /* 入力された単位を欄の単位へ換算するための、1単位あたりのポイント数（値は UNITS 表と同じ。キーは小文字）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Points per unit for converting typed units into the field's unit (same values as the UNITS table; lowercase keys) */
    var STEPPER_POINTS_PER_UNIT = {
        "in": 72, "inch": 72, "mm": 72 / 25.4, "cm": 72 / 2.54, "m": 72 / 25.4 * 1000,
        "pt": 1, "px": 1, "p": 12, "pc": 12, "pica": 12,
        "q": 72 / 25.4 * 0.25, "h": 72 / 25.4 * 0.25, "ft": 72 * 12, "yd": 72 * 36
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
     * 数値の後ろの単位は欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
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
        var fieldPointsPerUnit = STEPPER_POINTS_PER_UNIT[fieldUnitKey];
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
            var pointsPerUnit = STEPPER_POINTS_PER_UNIT[unitKey];
            if (pointsPerUnit === undefined || fieldPointsPerUnit === undefined) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = value * pointsPerUnit;
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldPointsPerUnit;
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
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

    /*
    LABELS のカテゴリ規約 / Category rules
      dialog     : ダイアログタイトル / dialog title
      panel      : パネル見出し / panel headers
      fieldLabel : 行ラベル（コロンは含めず、描画時に labelText() が付与） / row labels (colon added by labelText())
      radio      : ラジオボタン / radio buttons
      checkbox   : チェックボックス / checkboxes
      button     : ボタン（Cancel / Reset。OK は非ローカライズの "OK" 直書き）
      tooltip    : 意味が自明でないコントロールのツールチップ / tooltips for non-obvious controls
      alert      : 警告メッセージ / alerts
    括弧などの記号は言語ごとに直接書く（日本語は全角、英語は半角） / Write symbols such as parentheses per language (full-width JA, half-width EN)
    */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトのリサイズ", en: "SmartObjectResizer" }
        },
        panel: {
            base: { ja: "リサイズ基準", en: "Resize base" },
            /* 整列（横）: 左/中央/右 ＋ 縦方向の分配 / horizontal alignment */
            hAlign: { ja: "整列（横）", en: "Align (H)" },
            /* 整列（縦）: 上/中央/下 ＋ 横方向の分配 / vertical alignment */
            vAlign: { ja: "整列（縦）", en: "Align (V)" }
        },
        fieldLabel: {
            max: { ja: "最大", en: "Max" },
            min: { ja: "最小", en: "Min" },
            key: { ja: "キーオブジェクト", en: "Key object" },
            fixed: { ja: "指定サイズ", en: "Fixed Size" },
            base: { ja: "基準辺", en: "Ref. side" },
            area: { ja: "面積", en: "Area" },
            artboard: { ja: "アートボード", en: "Artboard" },
            bleed: { ja: "裁ち落とし", en: "Bleed" }
        },
        radio: {
            keepAspect: { ja: "縦横比保持", en: "Keep aspect" },
            oneSideOnly: { ja: "片辺のみ", en: "One side only" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            longSide: { ja: "長辺", en: "Long side" },
            shortSide: { ja: "短辺", en: "Short side" },
            areaMax: { ja: "最大", en: "Max" },
            areaMin: { ja: "最小", en: "Min" }
        },
        checkbox: {
            textOutlineBounds: { ja: "テキストをアウトライン境界で計測", en: "Measure text by outline bounds" },
            previewBounds: { ja: "プレビュー境界で計測", en: "Measure by preview bounds" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignEven: { ja: "均等", en: "Distribute evenly" },
            alignZero: { ja: "0間隔", en: "Zero gap" },
            alignTop: { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" }
        },
        tooltip: {
            keepAspect: { ja: "縦横比を保ったまま拡大／縮小します。", en: "Scale while keeping the aspect ratio." },
            oneSideOnly: {
                ja: "幅または高さの片辺だけを変更します（縦横比は保持しません）。",
                en: "Change only one side (width or height); aspect ratio is not preserved."
            },
            key: {
                ja: "キーオブジェクト（整列の基準に指定したオブジェクト）の幅／高さに、各オブジェクトをそろえます。",
                en: "Match every object to the width / height of the key object (the one set as the align target)."
            },
            keyNone: {
                ja: "キーオブジェクトが設定されていません。選択したうえで、基準にしたいオブジェクトをもう一度クリックしてください。",
                en: "No key object is set. With the objects selected, click the one you want as the key again."
            },
            base: {
                ja: "基準となる辺（長辺／短辺）の長さに合わせて、各オブジェクトをリサイズします。",
                en: "Resize each object to match the chosen reference side (long / short)."
            },
            area: {
                ja: "選択オブジェクトの面積を、最大／最小のものにそろえます。",
                en: "Match object areas to the largest / smallest in the selection."
            },
            artboard: {
                ja: "選択全体をアートボードの幅／高さに合わせ、中央に配置します。",
                en: "Fit the whole selection to the artboard width / height and center it."
            },
            bleed: {
                ja: "アートボード＋裁ち落とし（片側3mm）の幅／高さに合わせ、中央に配置します。",
                en: "Fit to the artboard plus bleed (3mm per side) and center it."
            },
            alignEven: {
                ja: "オブジェクトの間隔が均等になるように分配します（3つ以上で有効）。",
                en: "Distribute objects with equal gaps (needs 3+ objects)."
            },
            alignZero: {
                ja: "オブジェクトを間隔0で隙間なく並べます（2つ以上で有効）。",
                en: "Place objects with zero gap, no spacing (needs 2+ objects)."
            },
            textOutline: {
                ja: "テキストをアウトライン化した実際の字形の境界で計測します。",
                en: "Measure text by the actual outlined glyph bounds."
            },
            preview: {
                ja: "線幅や効果を含むプレビュー境界で計測します（オフは幾何境界）。",
                en: "Measure by preview bounds incl. strokes / effects (off = geometric bounds)."
            },
            reset: {
                ja: "サイズ・位置・整列をすべて元の状態に戻します。",
                en: "Revert size, position, and alignment to the original state."
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
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectObject: { ja: "オブジェクトを選択してください。", en: "Please select an object." }
        }
    };

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
    // 境界の計測 / Bounds measuring
    // =========================================

    /**
     * オブジェクトの境界を { width, height, left, top } で返す（クリップグループはマスクで測る）
     * @param {PageItem} pageItem - 対象
     * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
     * @returns {{width: number, height: number, left: number, top: number}} 境界
     */
    function getPageItemBoundsObject(pageItem, useVisibleBounds) {
        return boundsArrayToObject(getClipAwareBounds(pageItem, useVisibleBounds));
    }

    /**
     * [左, 上, 右, 下] を { width, height, left, top } にする
     * @param {number[]|null} boundsArray - 境界（null なら 0 の境界）
     * @returns {{width: number, height: number, left: number, top: number}} 境界
     */
    function boundsArrayToObject(boundsArray) {
        if (!boundsArray) return { width: 0, height: 0, left: 0, top: 0 };
        return {
            width: boundsArray[2] - boundsArray[0],
            height: boundsArray[1] - boundsArray[3],
            left: boundsArray[0],
            top: boundsArray[1]
        };
    }

    /**
     * 複数のオブジェクトを包む境界を { width, height, left, top } で返す（クリップグループはマスクで測る）
     * @param {PageItem[]} pageItems - 対象
     * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
     * @returns {{width: number, height: number, left: number, top: number}} 境界（空なら 0）
     */
    function getBoundsFromItems(pageItems, useVisibleBounds) {
        return boundsArrayToObject(getClipAwareUnionBounds(pageItems, useVisibleBounds));
    }

    /**
     * アウトライン境界で計測すべきテキストを含むか（テキスト、またはテキストを含むグループ）
     * @param {PageItem} pageItem - 対象
     * @returns {boolean} 含むなら true
     */
    function containsTextForOutlineBounds(pageItem) {
        if (!pageItem) return false;
        return collectSelectionTextFrames(pageItem, { textRangeToFrame: false, unique: false }).length > 0;
    }

    /**
     * 計測用に複製したグループの中のテキストをアウトライン化する（グループ自体は残る）
     * @param {GroupItem} groupItem - 複製したグループ
     * @returns {void}
     */
    function outlineTextFramesInGroupDuplicate(groupItem) {
        var textFrames = collectSelectionTextFrames(groupItem.pageItems, { unique: false });
        for (var i = textFrames.length - 1; i >= 0; i--) {
            if (textFrames[i] && textFrames[i].isValid) {
                textFrames[i].createOutline();
            }
        }
    }

    /**
     * 計測用の一時オブジェクトを削除する（ロック・非表示を解いてから）
     * @param {PageItem[]} pageItems - 削除するオブジェクト
     * @returns {void}
     */
    function removeItemsSafe(pageItems) {
        if (!pageItems || pageItems.length === 0) return;
        for (var i = pageItems.length - 1; i >= 0; i--) {
            /* 無効になった参照もあるので1操作ずつ守る / Guard each step; some references may be stale */
            try {
                pageItems[i].locked = false;
            } catch (e) { }
            try {
                pageItems[i].hidden = false;
            } catch (e) { }
            try {
                pageItems[i].remove();
            } catch (e) { }
        }
    }

    /**
     * キャッシュキー用に座標を丸める
     * @param {number} coordinate - 座標
     * @returns {number} 小数第3位までに丸めた値
     */
    function roundCacheCoord(coordinate) {
        return Math.round(coordinate * 1000) / 1000;
    }

    /**
     * ドキュメントの選択を配列にコピーする（doc.selection はライブ参照になりうるため）
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択のコピー（選択なしなら空配列）
     */
    function copySelectionItems(doc) {
        var copiedItems = [];
        var currentSelection = doc.selection;
        if (!currentSelection) return copiedItems;
        for (var i = 0; i < currentSelection.length; i++) copiedItems.push(currentSelection[i]);
        return copiedItems;
    }

    /**
     * 各オブジェクトのサイズと位置を控える
     * @param {PageItem[]} pageItems - 対象
     * @returns {Array<{item: PageItem, width: number, height: number, left: number, top: number}>} 控え
     */
    function captureItemStates(pageItems) {
        var itemStates = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            itemStates.push({
                item: pageItem,
                width: pageItem.width,
                height: pageItem.height,
                left: pageItem.left,
                top: pageItem.top
            });
        }
        return itemStates;
    }

    // =========================================
    // リサイズの基本操作 / Resize primitives
    // =========================================

    /**
     * 現在の長さを目標の長さにする倍率（%）を返す
     * @param {number} currentLength - 現在の長さ
     * @param {number} targetLength - 目標の長さ
     * @returns {number} 倍率（%）
     */
    function getScalePercent(currentLength, targetLength) {
        return (targetLength / currentLength) * 100;
    }

    /**
     * モードに応じて基準とする1辺の長さを返す（長辺／短辺／幅／高さ）
     * @param {{width: number, height: number}} bounds - 境界
     * @param {Object} mode - getSelectedResizeMode() の戻り値
     * @returns {number} 辺の長さ
     */
    function measureSide(bounds, mode) {
        if (mode.isLong) return Math.max(bounds.width, bounds.height);
        if (mode.isShort) return Math.min(bounds.width, bounds.height);
        return mode.isWidth ? bounds.width : bounds.height;
    }

    /**
     * 片辺のみをスケールする（基準点は左上、線幅は変えない）
     * 引数を省略した 2 引数版は基準点が中心になり、他モード（TOPLEFT）と挙動が食い違うため、常に明示指定する。
     * @param {PageItem} pageItem - 対象
     * @param {boolean} isWidth - 幅なら true、高さなら false
     * @param {number} scalePct - 倍率（%）
     * @returns {void}
     */
    function resizeOneSide(pageItem, isWidth, scalePct) {
        var scaleX = isWidth ? scalePct : 100;
        var scaleY = isWidth ? 100 : scalePct;
        pageItem.resize(scaleX, scaleY, true, true, true, true, 100, Transformation.TOPLEFT);
    }

    /**
     * 目標サイズへリサイズする。除数 0（線のみ・空テキスト等）は Infinity 回避のため 100%＝現状維持とする
     * 線幅倍率は呼び出し側が明示する（省略時は 100%＝据え置き）。scaleW を流用すると、
     * 線幅を変えていない片辺のみのリサイズを戻すときに線幅だけが縮んでいく。
     * ※ 幅/高さが 0 の要素は resize では拡大できないため「厳密復元」ではなく現状スキップに近い（実用上は許容）
     * @param {PageItem} pageItem - 対象
     * @param {number} targetWidth - 目標の幅
     * @param {number} targetHeight - 目標の高さ
     * @param {number} [lineScalePct] - 線幅の倍率（%）
     * @returns {void}
     */
    function resizeItemToSize(pageItem, targetWidth, targetHeight, lineScalePct) {
        var scaleW = (targetWidth === 0 || pageItem.width === 0) ? 100 : (targetWidth / pageItem.width) * 100;
        var scaleH = (targetHeight === 0 || pageItem.height === 0) ? 100 : (targetHeight / pageItem.height) * 100;
        var lineScale = (typeof lineScalePct === "number") ? lineScalePct : 100;
        pageItem.resize(scaleW, scaleH, true, true, true, true, lineScale, Transformation.TOPLEFT);
    }

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================
    /* Illustrator の DOM にキーオブジェクトを示すプロパティは無いため、整列コマンドで実測して特定する。
       キーオブジェクトが設定されていると、どの向きに整列してもそのオブジェクトだけは動かない。
       左右上下の4方向すべてで動かなかったものだけを採用する。1方向だけだと「たまたま端にいた
       オブジェクト」を拾ってしまうが、4方向すべての端を兼ねることは（同一バウンズでない限り）無い。
       候補が0個または2個以上のときは判定不能として null を返し、UI 側でこの基準をディムする。
       Illustrator exposes no key-object property, so probe it: run the align commands and see
       which item stays put in all four directions. Ambiguous results yield null. */

    /* 整列後の位置差をどこまで「動いていない」とみなすか（pt） / Move tolerance in points */
    var KEY_DETECT_TOLERANCE_PT = 0.001;

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * @param {PageItem[]} pageItems - 判定対象のオブジェクト
     * @returns {PageItem|null} キーオブジェクト。判定できないときは null
     */
    function detectKeyObject(pageItems) {
        if (!pageItems || pageItems.length < 2) return null;
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        for (var i = 0; i < pageItems.length; i++) stayedPut.push(true);

        for (var commandIndex = 0; commandIndex < alignCommands.length; commandIndex++) {
            var savedPositions = [];
            for (var j = 0; j < pageItems.length; j++) savedPositions.push([pageItems[j].left, pageItems[j].top]);
            app.redraw(); /* executeMenuCommand は直前の DOM 変更が反映されていないと空振りする / the command misfires without a redraw */
            app.executeMenuCommand(alignCommands[commandIndex]);
            for (var k = 0; k < pageItems.length; k++) {
                if (Math.abs(pageItems[k].left - savedPositions[k][0]) > KEY_DETECT_TOLERANCE_PT ||
                    Math.abs(pageItems[k].top - savedPositions[k][1]) > KEY_DETECT_TOLERANCE_PT) {
                    stayedPut[k] = false;
                }
                /* 整列は検出のための試行なので、その場で元の位置へ戻す / Undo the probe move right away */
                pageItems[k].left = savedPositions[k][0];
                pageItems[k].top = savedPositions[k][1];
            }
        }
        app.redraw();

        var keyCandidate = null;
        for (var m = 0; m < pageItems.length; m++) {
            if (!stayedPut[m]) continue;
            if (keyCandidate !== null) return null; /* 複数残った＝判定不能 / more than one left: ambiguous */
            keyCandidate = pageItems[m];
        }
        return keyCandidate;
    }

    // =========================================
    // 整列の座標 / Alignment edges
    // =========================================

    /* 境界から各辺・中心の座標を取り出す関数 / Accessors for each edge and center of a bounds object */
    var EDGE_OF = {
        left: function (bounds) { return bounds.left; },
        right: function (bounds) { return bounds.left + bounds.width; },
        centerX: function (bounds) { return bounds.left + bounds.width / 2; },
        top: function (bounds) { return bounds.top; },
        bottom: function (bounds) { return bounds.top - bounds.height; },
        centerY: function (bounds) { return bounds.top - bounds.height / 2; }
    };

    /* X 座標を動かす辺 / Edges that move along X */
    var HORIZONTAL_EDGES = { left: true, right: true, centerX: true };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、選択オブジェクトのリサイズと整列を行う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var rulerUnit = getUnitInfo();
        var unitLabel = rulerUnit.label;

        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("alert.selectObject"));
            return;
        }
        /* doc.selection はライブ参照になりうるため、配列にコピーして固定する / Freeze the selection into an array */
        var targetItems = copySelectionItems(doc);
        var originalStates = captureItemStates(targetItems);
        var resizeBaseStates = [];

        /* 直近の変形で線幅に掛けた倍率（%）。並びは originalStates と同じ。
           片辺のみは線幅を変えない（100%）ので、復元時に一律 scaleW を掛けると線幅だけがずれていく。
           Line-width percentage applied by the last transform, per item (same order as originalStates). */
        var appliedLineScales = [];
        resetAppliedLineScales();

        var keyObject = null;
        try {
            keyObject = detectKeyObject(targetItems);
        } catch (e) {
            /* 検出に失敗してもスクリプト自体は続行する（この基準がディムされるだけ）/ Keep going; the key row just dims */
            $.writeln("SmartObjectResizer: キーオブジェクトの検出に失敗 / key object detection failed — " + e);
            keyObject = null;
        }

        /* アウトライン計測のキャッシュ / Outline measurement cache */
        var outlineBoundsCache = {};
        var outlineBoundsCacheSeq = 1;
        var outlineIdMap = [];

        /* 確定／破棄の判定は dialog.show() の戻り値に一本化する。
           ボタンの onClick だけで復元すると、ESC キーやウィンドウの閉じるボタンでは
           onClick が発火しないため、プレビュー中の変形がそのまま確定してしまう。
           Deciding commit vs. discard from the return value of show() (not from the button
           handlers) is what makes ESC / the window close box behave as a real cancel. */
        var DIALOG_RESULT_OK = 1;
        var DIALOG_RESULT_CANCEL = 2;

        /* ダイアログのコントロール（buildDialog で作る）/ Dialog controls, created in buildDialog() */
        var keepRatioRadio, oneSideOnlyRadio;
        var allRadioButtons = [];       /* 基準ラジオのみ（縦横比保持／片辺のみは含まない）/ base radios only */
        var resizeBaseRadioGroups = []; /* [最大, 最小, キーオブジェクト, 指定サイズ, 基準辺, 面積, アートボード, 裁ち落とし] */
        var baseRadios, areaRadios;
        var fixedWidthRadio, fixedHeightRadio, fixedSizeInput;
        var textOutlineBoundsCheck, previewBoundsCheck;
        var alignLeftCheck, alignCenterCheck, alignRightCheck, verticalEvenCheck, verticalZeroGapCheck;
        var alignTopCheck, alignMiddleCheck, alignBottomCheck, horizontalEvenCheck, horizontalZeroGapCheck;
        var alignAxes = [];

        var resizeDialog = buildDialog();
        prepareDialogWindow(resizeDialog, SCRIPT_NAME);
        var dialogResult = resizeDialog.show();
        if (dialogResult !== DIALOG_RESULT_OK) {
            /* キャンセル／ESC／ウィンドウを閉じる: プレビュー中の変形をすべて破棄 / Discard every preview transform */
            restoreOriginalState();
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        /**
         * ダイアログを組み立てる
         * @returns {Window} ダイアログ
         */
        function buildDialog() {
            var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
            setupWindow(dialogWindow);

            dialogWindow.onShow = onDialogShow;

            addAspectModeGroup(dialogWindow);

            var columnsGroup = dialogWindow.add("group");
            columnsGroup.orientation = "row";
            columnsGroup.alignChildren = ["top", "top"]; /* 左右ペインを上揃えに / top-align both panes */
            columnsGroup.spacing = COLUMN_SPACING;

            /* 左ペイン（リサイズ基準と計測オプション）。余白は dialogWindow.margins と columnsGroup.spacing で管理 */
            var leftPane = columnsGroup.add("group");
            leftPane.orientation = "column";
            leftPane.alignChildren = ["left", "top"];
            addResizeBasePanel(leftPane);
            addMeasureOptions(leftPane);

            /* 右ペイン（整列）/ Right pane: alignment */
            var rightPane = columnsGroup.add("group");
            rightPane.orientation = "column";
            rightPane.alignChildren = ["fill", "top"];
            addAlignPanels(rightPane);
            bindAlignChecks();

            /* フッター（左=リセット / 右=キャンセル・OK）/ Footer (left: Reset, right: Cancel / OK) */
            var buttonRow = addButtonRow(dialogWindow);
            var btnReset = buttonRow.leftGroup.add("button", undefined, getLabel("button.reset"));
            setHelpTip(btnReset, getLabel("tooltip.reset"));
            btnReset.onClick = resetToOriginal;

            /* Mac 規約で Cancel → OK の順 / Cancel then OK, per macOS convention */
            var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            btnCancel.onClick = function () {
                dialogWindow.close(DIALOG_RESULT_CANCEL);
            };

            var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });
            btnOK.onClick = function () {
                /* 確定（一時グループを使わないので親階層の復元処理は不要）/ Commit; no temporary groups to unwind */
                dialogWindow.close(DIALOG_RESULT_OK);
            };
            return dialogWindow;
        }

        /**
         * 表示時：UI 状態を初期化し、元の状態へ戻す（基準が選択済みならそのモードを1回だけ反映）
         * 基準は初期未選択なので applyResizeBySelection() は何もしない（＝開いた直後は変形しない）。
         * @returns {void}
         */
        function onDialogShow() {
            /* UI 状態の初期化（DOM に触れないので保護不要。失敗したら顕在化させる）/ UI only; failures should surface */
            resetAlignChecks();
            updateInputState();
            updateRadioGroupStates();
            updateTextOutlineOptionState();

            /* 初期プレビュー（DOM 操作）のみ保護。失敗しても表示は継続するが、黙殺せずログに残す / Guard only the DOM work and log failures */
            try {
                resizeFromOriginal();
            } catch (e) {
                $.writeln("SmartObjectResizer: 初期プレビューに失敗 / initial preview failed — " + e);
            }
        }

        /**
         * 「縦横比保持」「片辺のみ」の行を作る
         * @param {Window} dialogWindow - ダイアログ
         * @returns {void}
         */
        function addAspectModeGroup(dialogWindow) {
            var aspectModeGroup = dialogWindow.add("group");
            aspectModeGroup.orientation = "row";
            aspectModeGroup.alignChildren = ["left", "center"];
            aspectModeGroup.margins = [20, 0, 0, 0];
            /* コントロール単体の margins は ScriptUI では無視されるため、間隔は親の spacing で取る
               Per-control margins are ignored in ScriptUI; use the parent group's spacing instead. */
            aspectModeGroup.spacing = 20;
            aspectModeGroup.alignment = ["center", "top"];

            keepRatioRadio = aspectModeGroup.add("radiobutton", undefined, getLabel("radio.keepAspect"));
            oneSideOnlyRadio = aspectModeGroup.add("radiobutton", undefined, getLabel("radio.oneSideOnly"));

            keepRatioRadio.value = true;
            setHelpTip(keepRatioRadio, getLabel("tooltip.keepAspect"));
            setHelpTip(oneSideOnlyRadio, getLabel("tooltip.oneSideOnly"));

            oneSideOnlyRadio.onClick = function () { onRatioModeChanged(true); };
            keepRatioRadio.onClick = function () { onRatioModeChanged(false); };
        }

        /**
         * 「リサイズ基準」パネルを作る
         * @param {Group} leftPane - 追加先の左ペイン
         * @returns {void}
         */
        function addResizeBasePanel(leftPane) {
            var resizeBasePanel = leftPane.add("panel", undefined, getLabel("panel.base"));
            setupPanel(resizeBasePanel, 6);

            var widthHeightLabels = [getLabel("radio.width"), getLabel("radio.height")];
            var maxRadios = addRadioRow(labelText("fieldLabel.max"), widthHeightLabels, resizeBasePanel);
            var minRadios = addRadioRow(labelText("fieldLabel.min"), widthHeightLabels, resizeBasePanel);
            var keyRadios = addRadioRow(labelText("fieldLabel.key"), widthHeightLabels, resizeBasePanel);
            var fixedRadios = addFixedSizeRows(labelText("fieldLabel.fixed"), widthHeightLabels, resizeBasePanel);
            baseRadios = addRadioRow(labelText("fieldLabel.base"), [getLabel("radio.longSide"), getLabel("radio.shortSide")], resizeBasePanel);
            areaRadios = addRadioRow(labelText("fieldLabel.area"), [getLabel("radio.areaMax"), getLabel("radio.areaMin")], resizeBasePanel);
            resizeBasePanel.add("statictext", undefined, "  ───────────────  ");
            var artboardRadios = addRadioRow(labelText("fieldLabel.artboard"), widthHeightLabels, resizeBasePanel);
            var bleedRadios = addRadioRow(labelText("fieldLabel.bleed"), widthHeightLabels, resizeBasePanel);

            /* キーオブジェクトが特定できなかったときは、この行だけディムする / Dim the key row when no key object was found */
            keyRadios[0].parent.enabled = !!keyObject;

            resizeBaseRadioGroups.push(maxRadios, minRadios, keyRadios, fixedRadios, baseRadios, areaRadios, artboardRadios, bleedRadios);

            /* 意味が自明でない基準にツールチップを設定（最大／最小／指定サイズは自明なため付けない）/ Tooltips for non-obvious bases */
            setHelpTip(keyRadios, keyObject ? getLabel("tooltip.key") : getLabel("tooltip.keyNone"));
            setHelpTip(baseRadios, getLabel("tooltip.base"));
            setHelpTip(areaRadios, getLabel("tooltip.area"));
            setHelpTip(artboardRadios, getLabel("tooltip.artboard"));
            setHelpTip(bleedRadios, getLabel("tooltip.bleed"));

            /* 初期状態ではどの基準も選択しない（開いた直後は変形しない）/ No base is selected initially, so opening performs no transform */
            for (var i = 0; i < allRadioButtons.length; i++) {
                allRadioButtons[i].onClick = onResizeBaseRadioClick;
            }
        }

        /**
         * 「ラベル＋ラジオ複数」の1行を作る
         * @param {string} rowLabelText - 行ラベル
         * @param {string[]} optionLabels - ラジオの文言
         * @param {Panel} parentPanel - 追加先
         * @returns {RadioButton[]} 作ったラジオ
         */
        function addRadioRow(rowLabelText, optionLabels, parentPanel) {
            var rowGroup = parentPanel.add("group");
            rowGroup.orientation = "row";
            rowGroup.alignChildren = ["left", "center"];
            addRowLabel(rowGroup, rowLabelText);
            var rowRadios = [];
            for (var i = 0; i < optionLabels.length; i++) {
                var radioButton = rowGroup.add("radiobutton", undefined, optionLabels[i]);
                rowRadios.push(radioButton);
                allRadioButtons.push(radioButton);
            }
            return rowRadios;
        }

        /**
         * 指定サイズの2行（「ラベル＋幅/高さラジオ」と「数値入力＋単位」）を作る
         * @param {string} rowLabelText - 行ラベル
         * @param {string[]} optionLabels - 幅・高さの文言
         * @param {Panel} parentPanel - 追加先
         * @returns {RadioButton[]} 幅・高さのラジオ
         */
        function addFixedSizeRows(rowLabelText, optionLabels, parentPanel) {
            /* 1行目：ラベルとラジオボタン / Row 1: label and radios */
            var sizeHeaderGroup = parentPanel.add("group");
            sizeHeaderGroup.orientation = "row";
            sizeHeaderGroup.alignChildren = ["left", "center"];
            addRowLabel(sizeHeaderGroup, rowLabelText);
            fixedWidthRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[0]);
            fixedHeightRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[1]);
            allRadioButtons.push(fixedWidthRadio, fixedHeightRadio);

            /* 2行目：空ラベル（入力欄の左端をそろえる）と入力欄 / Row 2: blank label and the field */
            var sizeInputGroup = parentPanel.add("group");
            sizeInputGroup.orientation = "row";
            sizeInputGroup.alignChildren = ["left", "center"];
            addRowLabel(sizeInputGroup, "");

            /* 選択オブジェクトの平均幅を初期値に。getReferenceBounds() は pt を返すが、入力欄は定規単位で扱うため必ず換算する。
               換算を忘れると mm 定規で「283」と表示され、そのまま 283mm にリサイズされる。
               Bounds are in points but this field is in ruler units — always convert. */
            var totalWidth = 0;
            for (var i = 0; i < targetItems.length; i++) {
                totalWidth += getReferenceBounds(targetItems[i], true).width;
            }
            var avgWidthPt = targetItems.length > 0 ? (totalWidth / targetItems.length) : 100;
            var avgWidth = avgWidthPt / rulerUnit.pointsPerUnit;
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var sizeStepperGroup = sizeInputGroup.add("group");
            sizeStepperGroup.orientation = "row";
            sizeStepperGroup.alignChildren = ["left", "center"];
            sizeStepperGroup.spacing = 0;
            sizeStepperGroup.margins = 0;
            var sizeStepper = addStepper(sizeStepperGroup, function () { return fixedSizeInput; }, {
                step: 1, min: 0,
                /* ∧∨・↑↓では onChange が発火しないので明示的に呼ぶ。中の DOM 操作が失敗しても操作は止めない
                   Fire onChange explicitly; keep stepping alive even if its DOM work fails */
                onStep: function () {
                    try {
                        if (typeof fixedSizeInput.onChange === "function") fixedSizeInput.onChange();
                    } catch (e) { }
                }
            });
            fixedSizeInput = sizeStepperGroup.add("edittext", undefined, avgWidth.toFixed(0));
            fixedSizeInput.characters = 5;
            fixedSizeInput.stepperGroup = sizeStepper;
            bindSteppedArrowKeys(fixedSizeInput, sizeStepper);
            sizeInputGroup.add("statictext", undefined, unitLabel);

            fixedSizeInput.onChange = function () {
                if ((fixedWidthRadio.value || fixedHeightRadio.value) && !isNaN(parseFloat(fixedSizeInput.text))) {
                    /* 片辺のみ／縦横比保持のどちらも applyResizeBySelection() に集約 / Both modes go through applyResizeBySelection() */
                    resizeFromOriginal();
                }
            };
            return [fixedWidthRadio, fixedHeightRadio];
        }

        /**
         * 計測オプション（アウトライン境界／プレビュー境界）のチェックボックスを作る
         * @param {Group} leftPane - 追加先の左ペイン
         * @returns {void}
         */
        function addMeasureOptions(leftPane) {
            var measureOptionsGroup = leftPane.add("group");
            measureOptionsGroup.orientation = "column";
            measureOptionsGroup.alignChildren = ["left", "top"];
            measureOptionsGroup.margins = [0, 5, 0, 0];

            textOutlineBoundsCheck = measureOptionsGroup.add("checkbox", undefined, getLabel("checkbox.textOutlineBounds"));
            textOutlineBoundsCheck.value = false;
            textOutlineBoundsCheck.onClick = onMeasureOptionChanged;
            setHelpTip(textOutlineBoundsCheck, getLabel("tooltip.textOutline"));

            previewBoundsCheck = measureOptionsGroup.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
            previewBoundsCheck.value = true;
            previewBoundsCheck.onClick = onMeasureOptionChanged;
            setHelpTip(previewBoundsCheck, getLabel("tooltip.preview"));

            updateTextOutlineOptionState();
        }

        /**
         * 整列（横）・整列（縦）のパネルを作る
         * 横位置パネルは左/中央/右＋縦方向の分配、縦位置パネルは上/中央/下＋横方向の分配
         * @param {Group} rightPane - 追加先の右ペイン
         * @returns {void}
         */
        function addAlignPanels(rightPane) {
            var hAlignPanel = rightPane.add("panel", undefined, getLabel("panel.hAlign"));
            setupPanel(hAlignPanel, 6);
            alignLeftCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignLeft"));
            alignCenterCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignCenter"));
            alignRightCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignRight"));
            addSpacer(hAlignPanel, 5); /* 「均等」の上の余白（整列⇔分配の区切り）/ gap between align and distribute */
            verticalEvenCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignEven"));
            verticalZeroGapCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignZero"));
            setHelpTip(verticalEvenCheck, getLabel("tooltip.alignEven"));
            setHelpTip(verticalZeroGapCheck, getLabel("tooltip.alignZero"));

            var vAlignPanel = rightPane.add("panel", undefined, getLabel("panel.vAlign"));
            setupPanel(vAlignPanel, 6);
            alignTopCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignTop"));
            alignMiddleCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignMiddle"));
            alignBottomCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignBottom"));
            addSpacer(vAlignPanel, 5); /* 「均等」の上の余白（整列⇔分配の区切り）/ gap between align and distribute */
            horizontalEvenCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignEven"));
            horizontalZeroGapCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignZero"));
            setHelpTip(horizontalEvenCheck, getLabel("tooltip.alignEven"));
            setHelpTip(horizontalZeroGapCheck, getLabel("tooltip.alignZero"));
        }

        /**
         * 整列の定義テーブルを作り、クリック時の処理と有効・無効を設定する
         * テーブルは onClick の割り当てと再適用の唯一のソース。minItems はその整列が成立する最小オブジェクト数
         * （「均等」は両端を固定して間を分けるので3個以上、「0間隔」は2個以上）。
         * minItems を有効・無効にも反映しないと、2個選択で「均等」をチェックできるのに何も起きない。
         * @returns {void}
         */
        function bindAlignChecks() {
            alignAxes = [
                /* 横位置（X座標を変更）: 左 / 中央 / 右 / 横均等 / 横0 / Horizontal (moves X) */
                [
                    { check: alignLeftCheck,         minItems: 1, apply: function () { alignToExtremeEdge("left", false); } },
                    { check: alignCenterCheck,       minItems: 1, apply: function () { alignToSelectionCenter("centerX"); } },
                    { check: alignRightCheck,        minItems: 1, apply: function () { alignToExtremeEdge("right", true); } },
                    { check: horizontalEvenCheck,    minItems: 3, apply: function () { distributeHorizontal(true); } },
                    { check: horizontalZeroGapCheck, minItems: 2, apply: function () { distributeHorizontal(false); } }
                ],
                /* 縦位置（Y座標を変更）: 上 / 中央 / 下 / 均等 / 0 / Vertical (moves Y) */
                [
                    { check: alignTopCheck,        minItems: 1, apply: function () { alignToExtremeEdge("top", true); } },
                    { check: alignMiddleCheck,     minItems: 1, apply: function () { alignToSelectionCenter("centerY"); } },
                    { check: alignBottomCheck,     minItems: 1, apply: function () { alignToExtremeEdge("bottom", false); } },
                    { check: verticalEvenCheck,    minItems: 3, apply: function () { distributeVertical(true); } },
                    { check: verticalZeroGapCheck, minItems: 2, apply: function () { distributeVertical(false); } }
                ]
            ];

            for (var i = 0; i < alignAxes.length; i++) {
                var axisEntries = alignAxes[i];
                for (var j = 0; j < axisEntries.length; j++) {
                    var siblingChecks = [];
                    for (var k = 0; k < axisEntries.length; k++) {
                        if (k !== j) siblingChecks.push(axisEntries[k].check);
                    }
                    axisEntries[j].check.onClick = makeAlignHandler(axisEntries[j].check, siblingChecks);
                    axisEntries[j].check.enabled = targetItems.length >= axisEntries[j].minItems;
                }
            }
        }

        // -----------------------------------------
        // UI の状態 / UI state
        // -----------------------------------------

        /**
         * 整列チェックボックスをすべてオフにする
         * @returns {void}
         */
        function resetAlignChecks() {
            for (var i = 0; i < alignAxes.length; i++) {
                for (var j = 0; j < alignAxes[i].length; j++) alignAxes[i][j].check.value = false;
            }
        }

        /**
         * 「片辺のみ」のとき基準辺（長辺／短辺）と面積をディムする
         * アートボード／裁ち落としは片辺スケールにも対応するため常に有効のままにする。
         * @returns {void}
         */
        function updateRadioGroupStates() {
            var keepAspectOnly = !oneSideOnlyRadio.value;
            baseRadios[0].parent.enabled = keepAspectOnly;
            areaRadios[0].parent.enabled = keepAspectOnly;
        }

        /**
         * 指定サイズ（幅/高さ）が選択されているときだけ入力欄を有効にする
         * @returns {void}
         */
        function updateInputState() {
            var isFixedSize = !!(fixedWidthRadio.value || fixedHeightRadio.value);
            fixedSizeInput.enabled = isFixedSize;
            fixedSizeInput.stepperGroup.enabled = isFixedSize;
            redrawSteppersIn(fixedSizeInput.stepperGroup);
        }

        /**
         * 選択にテキストがあるときだけ「テキストをアウトライン境界で計測」を有効にしてオンにする
         * @returns {void}
         */
        function updateTextOutlineOptionState() {
            var hasText = false;
            for (var i = 0; i < targetItems.length; i++) {
                if (containsTextForOutlineBounds(targetItems[i])) {
                    hasText = true;
                    break;
                }
            }

            if (hasText) {
                textOutlineBoundsCheck.enabled = true;
                textOutlineBoundsCheck.value = true;
            } else {
                textOutlineBoundsCheck.value = false;
                textOutlineBoundsCheck.enabled = false;
            }
        }

        /**
         * 「片辺のみ」でディムされる基準（基準辺／面積）の選択を解除する
         * ディムするだけだと value が残り、getSelectedResizeMode() が enabled を見ないため
         * 「片辺のみのはずが縦横比保持でリサイズされる」状態になる。
         * @returns {void}
         */
        function clearDimmedBaseSelections() {
            var dimmedGroups = [baseRadios, areaRadios];
            for (var groupIndex = 0; groupIndex < dimmedGroups.length; groupIndex++) {
                for (var radioIndex = 0; radioIndex < dimmedGroups[groupIndex].length; radioIndex++) {
                    dimmedGroups[groupIndex][radioIndex].value = false;
                }
            }
        }

        /**
         * 選択中のラジオからリサイズモードをフラグ付きで取得する
         * @returns {Object|null} モード。基準が未選択なら null
         */
        function getSelectedResizeMode() {
            for (var groupIndex = 0; groupIndex < resizeBaseRadioGroups.length; groupIndex++) {
                var sideIndex = -1;
                if (resizeBaseRadioGroups[groupIndex][0].value) sideIndex = 0;
                else if (resizeBaseRadioGroups[groupIndex][1].value) sideIndex = 1;
                if (sideIndex < 0) continue;
                return {
                    isWidth: sideIndex === 0,
                    isMax: groupIndex === 0,
                    isMin: groupIndex === 1,
                    isKey: groupIndex === 2,
                    isFixed: groupIndex === 3,
                    isLong: groupIndex === 4 && sideIndex === 0,
                    isShort: groupIndex === 4 && sideIndex === 1,
                    isArea: groupIndex === 5,
                    isAreaMax: groupIndex === 5 && sideIndex === 0,
                    isArtboard: groupIndex === 6,
                    isBleed: groupIndex === 7
                };
            }
            return null;
        }

        // -----------------------------------------
        // イベント / Events
        // -----------------------------------------

        /**
         * 縦横比保持／片辺のみを切り替える（有効・無効の切り替えは updateRadioGroupStates() が担当）
         * @param {boolean} useOneSideOnly - 片辺のみなら true
         * @returns {void}
         */
        function onRatioModeChanged(useOneSideOnly) {
            keepRatioRadio.value = !useOneSideOnly;
            oneSideOnlyRadio.value = useOneSideOnly;

            if (useOneSideOnly) clearDimmedBaseSelections();
            reapplyCurrentSelection();
        }

        /**
         * 基準ラジオのクリック：他の基準ラジオをオフにしてから再適用する（this はクリックしたラジオ）
         * @returns {void}
         */
        function onResizeBaseRadioClick() {
            for (var i = 0; i < allRadioButtons.length; i++) {
                if (allRadioButtons[i] !== this) {
                    allRadioButtons[i].value = false;
                }
            }
            reapplyCurrentSelection();
        }

        /**
         * 計測オプション（アウトライン境界／プレビュー境界）の切り替え
         * @returns {void}
         */
        function onMeasureOptionChanged() {
            resetAlignChecks();
            resizeFromOriginal();
        }

        /**
         * 整列チェックのクリック処理を作る
         * ON なら同一軸の兄弟を OFF にし、基準状態へ戻してから、チェック中の整列を両軸まとめて再適用する。
         * restoreResizeBaseState() は left/top を両方戻すため、自分の軸だけ再適用するともう一方の軸の整列が消える。
         * @param {Checkbox} alignCheck - クリックされたチェックボックス
         * @param {Checkbox[]} siblingChecks - 同じ軸の他のチェックボックス
         * @returns {Function} onClick に設定する関数
         */
        function makeAlignHandler(alignCheck, siblingChecks) {
            return function () {
                if (alignCheck.value) {
                    for (var i = 0; i < siblingChecks.length; i++) siblingChecks[i].value = false;
                }
                restoreResizeBaseState();
                reapplyActiveAlignments();
            };
        }

        /**
         * リセット：基準ラジオ・整列チェック・基準状態をすべて解除し、UI と実状態を初期状態にそろえる
         * resizeBaseStates を消さないと、この後に整列をクリックしたとき
         * restoreResizeBaseState() がリセット前のリサイズ結果を復元してしまう。
         * @returns {void}
         */
        function resetToOriginal() {
            for (var i = 0; i < allRadioButtons.length; i++) {
                allRadioButtons[i].value = false;
            }
            resizeBaseStates = [];
            resetAlignChecks();
            updateInputState();
            restoreOriginalState();
        }

        // -----------------------------------------
        // 復元と再適用 / Restore and reapply
        // -----------------------------------------

        /**
         * 線幅倍率の記録を 100% に戻す
         * @returns {void}
         */
        function resetAppliedLineScales() {
            appliedLineScales = [];
            for (var i = 0; i < targetItems.length; i++) appliedLineScales.push(100);
        }

        /**
         * 元のサイズへ戻す（掛けた線幅倍率の逆数で線幅も戻す）
         * @returns {void}
         */
        function restoreOriginalGeometry() {
            for (var i = 0; i < originalStates.length; i++) {
                var appliedLineScale = appliedLineScales[i] || 100;
                resizeItemToSize(originalStates[i].item, originalStates[i].width, originalStates[i].height, 10000 / appliedLineScale);
            }
            resetAppliedLineScales();
        }

        /**
         * 元の位置へ戻す
         * @returns {void}
         */
        function restoreOriginalPosition() {
            for (var i = 0; i < originalStates.length; i++) {
                originalStates[i].item.left = originalStates[i].left;
                originalStates[i].item.top = originalStates[i].top;
            }
        }

        /**
         * キャッシュを捨て、サイズと位置を元に戻して再描画する
         * @returns {void}
         */
        function restoreOriginalState() {
            clearOutlineBoundsCache();
            restoreOriginalGeometry();
            restoreOriginalPosition();
            app.redraw();
        }

        /**
         * 元の状態に戻してから、選択中の基準でリサイズし直す
         * @returns {void}
         */
        function resizeFromOriginal() {
            restoreOriginalState();
            applyResizeBySelection();
        }

        /**
         * 現在の選択モードを再適用する（ラジオはクリアしない）。チェック中の整列も新しいサイズに対して再適用する
         * @returns {void}
         */
        function reapplyCurrentSelection() {
            updateInputState();
            resizeFromOriginal();
            reapplyActiveAlignments();
            updateRadioGroupStates();
        }

        /**
         * リサイズ直後の状態を整列の基準として控える
         * @returns {void}
         */
        function captureResizeBaseState() {
            resizeBaseStates = captureItemStates(targetItems);
        }

        /**
         * 控えたリサイズ直後の状態へ戻す
         * @returns {void}
         */
        function restoreResizeBaseState() {
            if (!resizeBaseStates || resizeBaseStates.length === 0) return;
            for (var i = 0; i < resizeBaseStates.length; i++) {
                var itemState = resizeBaseStates[i];
                resizeItemToSize(itemState.item, itemState.width, itemState.height);
                itemState.item.left = itemState.left;
                itemState.item.top = itemState.top;
            }
            app.redraw();
        }

        /**
         * チェック中の整列を、両軸それぞれ最大1つずつ再適用する
         * @returns {void}
         */
        function reapplyActiveAlignments() {
            var changed = false;
            for (var i = 0; i < alignAxes.length; i++) {
                var axisEntries = alignAxes[i];
                for (var j = 0; j < axisEntries.length; j++) {
                    var alignEntry = axisEntries[j];
                    if (alignEntry.check.value && targetItems.length >= alignEntry.minItems) {
                        alignEntry.apply();
                        changed = true;
                        break; /* 同一軸は1つだけ / one per axis */
                    }
                }
            }
            if (changed) app.redraw();
        }

        // -----------------------------------------
        // 境界 / Bounds
        // -----------------------------------------

        /**
         * リサイズと整列に使う境界を返す
         * 「プレビュー境界で計測」ON のときは visibleBounds（線幅・効果込み＝見た目の端）、OFF のときは geometricBounds。
         * テキストは「アウトライン境界で計測」ON なら、アウトライン化した複製で計測する。
         * @param {PageItem} pageItem - 対象
         * @param {boolean} [forceVisible] - 設定に関係なくプレビュー境界を使うなら true
         * @returns {{width: number, height: number, left: number, top: number}} 境界
         */
        function getReferenceBounds(pageItem, forceVisible) {
            var useVisibleBounds = !!(forceVisible || (previewBoundsCheck && previewBoundsCheck.value));
            if (textOutlineBoundsCheck && textOutlineBoundsCheck.value && containsTextForOutlineBounds(pageItem)) {
                return getOutlinedBoundsCached(pageItem, useVisibleBounds);
            }
            return getPageItemBoundsObject(pageItem, useVisibleBounds);
        }

        /**
         * 選択全体（クラスタ）の合成境界を返す
         * @param {PageItem[]} pageItems - 対象
         * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 境界
         */
        function getCombinedReferenceBounds(pageItems) {
            var left = null, top = null, right = null, bottom = null;
            for (var i = 0; i < pageItems.length; i++) {
                var itemBounds = getReferenceBounds(pageItems[i]);
                if (left === null || itemBounds.left < left) left = itemBounds.left;
                if (top === null || itemBounds.top > top) top = itemBounds.top;
                if (right === null || (itemBounds.left + itemBounds.width) > right) right = itemBounds.left + itemBounds.width;
                if (bottom === null || (itemBounds.top - itemBounds.height) < bottom) bottom = itemBounds.top - itemBounds.height;
            }
            return { left: left, top: top, right: right, bottom: bottom, width: right - left, height: top - bottom };
        }

        /**
         * アウトライン計測のキャッシュを全消去する（幾何変化はキーでも吸収するが、モード切り替え時は明示的に消す）
         * @returns {void}
         */
        function clearOutlineBoundsCache() {
            outlineBoundsCache = {};
            outlineIdMap = [];
        }

        /**
         * アウトライン計測のキャッシュキーを返す（オブジェクトの ID ＋境界）
         * @param {PageItem} pageItem - 対象
         * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
         * @returns {string} キー
         */
        function getOutlineCacheKey(pageItem, useVisibleBounds) {
            var boundsArray = useVisibleBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
            return getOutlineId(pageItem) + "_" + (useVisibleBounds ? "v_" : "g_") +
                roundCacheCoord(boundsArray[0]) + "_" +
                roundCacheCoord(boundsArray[1]) + "_" +
                roundCacheCoord(boundsArray[2]) + "_" +
                roundCacheCoord(boundsArray[3]);
        }

        /**
         * キャッシュ用にオブジェクトへ振った ID を返す（無ければ振る）
         * @param {PageItem} pageItem - 対象
         * @returns {string} ID
         */
        function getOutlineId(pageItem) {
            for (var i = 0; i < outlineIdMap.length; i++) {
                if (outlineIdMap[i].item === pageItem) {
                    return outlineIdMap[i].id;
                }
            }
            var newId = "sor_" + (outlineBoundsCacheSeq++);
            outlineIdMap.push({ item: pageItem, id: newId });
            return newId;
        }

        /**
         * アウトライン境界をキャッシュ経由で返す
         * @param {PageItem} pageItem - 対象
         * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
         * @returns {{width: number, height: number, left: number, top: number}} 境界
         */
        function getOutlinedBoundsCached(pageItem, useVisibleBounds) {
            var cacheKey = getOutlineCacheKey(pageItem, useVisibleBounds);
            if (outlineBoundsCache.hasOwnProperty(cacheKey)) {
                return outlineBoundsCache[cacheKey];
            }
            var measuredBounds = measureOutlinedBoundsByDuplicate(pageItem, useVisibleBounds);
            outlineBoundsCache[cacheKey] = measuredBounds;
            return measuredBounds;
        }

        /**
         * 複製をアウトライン化して境界を計測する（一時オブジェクトは必ず削除）
         * 計測用の一時オブジェクトは、ドキュメント全体の差分ではなく生成物を直接ためて追跡する
         * （全件差分は参照の線形探索と組み合わさって O(N^2) になる）。
         * @param {PageItem} pageItem - 対象
         * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
         * @returns {{width: number, height: number, left: number, top: number}} 境界（失敗時は元の境界）
         */
        function measureOutlinedBoundsByDuplicate(pageItem, useVisibleBounds) {
            var createdItems = [];
            try {
                var duplicateItem = pageItem.duplicate();
                createdItems.push(duplicateItem);

                /* テキスト計測用の複製では、アウトライン化の前に分割を適用する / Expand appearance before outlining */
                if (containsTextForOutlineBounds(duplicateItem)) {
                    var expandedItems = expandDuplicateAppearance(duplicateItem);
                    if (expandedItems.length > 0) {
                        /* expandStyle は元の複製を作り直すので、追跡対象ごと差し替える / expandStyle rebuilds the duplicate */
                        createdItems = expandedItems;
                        duplicateItem = expandedItems[0];
                    }
                }

                if (duplicateItem.typename === "TextFrame") {
                    /* createOutline() は元のテキストフレームを置き換える / createOutline() replaces the frame */
                    duplicateItem = duplicateItem.createOutline();
                    createdItems = [duplicateItem];
                } else if (duplicateItem.typename === "GroupItem") {
                    /* グループ自体は残り、中のテキストフレームだけが置き換わる / The group stays; only its text is replaced */
                    outlineTextFramesInGroupDuplicate(duplicateItem);
                }

                if (createdItems.length === 0) return getPageItemBoundsObject(pageItem, useVisibleBounds);
                return getBoundsFromItems(createdItems, useVisibleBounds);
            } catch (e) {
                return getPageItemBoundsObject(pageItem, useVisibleBounds);
            } finally {
                removeItemsSafe(createdItems);
            }
        }

        /**
         * 複製に「アピアランスを分割」を適用し、生成された新しいアイテムの配列を返す
         * @param {PageItem} duplicateItem - 計測用の複製
         * @returns {PageItem[]} 生成されたアイテム（失敗時は空配列。呼び出し側は元の複製をそのまま使う）
         */
        function expandDuplicateAppearance(duplicateItem) {
            var expandedItems = [];
            var previousSelection = copySelectionItems(doc);
            try {
                doc.selection = null;
                duplicateItem.selected = true;
                app.executeMenuCommand('expandStyle');
                expandedItems = copySelectionItems(doc);
            } catch (e) {
            } finally {
                restoreSelectionItems(previousSelection);
            }
            return expandedItems;
        }

        /**
         * 選択を復元する（配列を doc.selection に直接代入せず、1件ずつ selected を立てる）
         * @param {PageItem[]} pageItems - 選択し直すオブジェクト
         * @returns {void}
         */
        function restoreSelectionItems(pageItems) {
            try {
                doc.selection = null;
                for (var i = 0; i < pageItems.length; i++) {
                    if (pageItems[i] && pageItems[i].isValid) pageItems[i].selected = true;
                }
            } catch (e) { }
        }

        // -----------------------------------------
        // リサイズ / Resize
        // -----------------------------------------

        /**
         * アクティブアートボードの矩形を返す
         * @returns {number[]} [左, 上, 右, 下]
         */
        function getActiveArtboardRect() {
            return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        }

        /**
         * 目標寸法（pt）を求める（面積モードは applyAreaResize で別処理）
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @returns {number|null} 目標寸法。求められなければ null
         */
        function computeReferenceValue(mode) {
            if (mode.isKey) {
                if (!keyObject) return null;
                var keyBounds = getReferenceBounds(keyObject);
                return mode.isWidth ? keyBounds.width : keyBounds.height;
            }
            if (mode.isFixed) {
                var parsedSize = parseFloat(fixedSizeInput.text);
                if (isNaN(parsedSize) || parsedSize <= 0) return null;
                return parsedSize * rulerUnit.pointsPerUnit;
            }
            if (mode.isArtboard) {
                var artboardRect = getActiveArtboardRect();
                return mode.isWidth ? (artboardRect[2] - artboardRect[0]) : (artboardRect[1] - artboardRect[3]);
            }
            if (mode.isBleed) {
                var bleedBase = getActiveArtboardRect();
                var sideLength = mode.isWidth ? (bleedBase[2] - bleedBase[0]) : (bleedBase[1] - bleedBase[3]);
                return sideLength + BLEED_OFFSET_MM * UNITS[1].pointsPerUnit;
            }
            /* 最大／最小／長辺／短辺: 全アイテムから基準値を集計 / Max, min, long, short: aggregate over all items */
            var referenceValue = null;
            for (var i = 0; i < targetItems.length; i++) {
                var measuredLength = measureSide(getReferenceBounds(targetItems[i]), mode);
                if (mode.isMin && measuredLength === 0) continue;
                if (referenceValue === null) referenceValue = measuredLength;
                else if (mode.isMax && measuredLength > referenceValue) referenceValue = measuredLength;
                else if (mode.isMin && measuredLength < referenceValue) referenceValue = measuredLength;
            }
            return referenceValue;
        }

        /**
         * 選択全体をひとまとまりとしてスケールする（アートボード／裁ち落とし基準。グループ化しない）
         * クラスタ左上を原点に、各アイテムのサイズと相対位置を同じ倍率でスケールする。
         * 「片辺のみ」では基準辺の軸だけ倍率を掛け、もう一方の軸は等倍のまま残す。
         * 親階層・重ね順を一切変更しないので、一時グループ方式の復元リスクを構造的に回避する。
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @param {number} referenceValue - 目標寸法（pt）
         * @returns {void}
         */
        function applyClusterResize(mode, referenceValue) {
            var clusterBounds = getCombinedReferenceBounds(targetItems);
            var currentLength = mode.isWidth ? clusterBounds.width : clusterBounds.height;
            if (!currentLength) return;
            var scaleFactor = referenceValue / currentLength;
            var isOneSideOnly = oneSideOnlyRadio.value;
            var factorX = (!isOneSideOnly || mode.isWidth) ? scaleFactor : 1;
            var factorY = (!isOneSideOnly || !mode.isWidth) ? scaleFactor : 1;

            /* リサイズで位置がずれる前に、各アイテムの左上と原点（クラスタ左上）を記録 / Record positions before resizing */
            var originLeft = null, originTop = null;
            var originalPositions = [];
            for (var i = 0; i < targetItems.length; i++) {
                var itemLeft = targetItems[i].left, itemTop = targetItems[i].top;
                originalPositions.push({ left: itemLeft, top: itemTop });
                if (originLeft === null || itemLeft < originLeft) originLeft = itemLeft;
                if (originTop === null || itemTop > originTop) originTop = itemTop;
            }

            /* 線幅は等倍のときだけ追従させる（非等倍では resizeOneSide と同じく 100%）/ Line widths follow uniform scaling only */
            var lineScalePct = isOneSideOnly ? 100 : scaleFactor * 100;
            for (var j = 0; j < targetItems.length; j++) {
                var pageItem = targetItems[j];
                pageItem.resize(factorX * 100, factorY * 100, true, true, true, true, lineScalePct, Transformation.TOPLEFT);
                appliedLineScales[j] = lineScalePct;
                /* 原点からの相対位置も同倍率でスケール（クラスタとして拡大縮小）/ Scale the offsets from the origin too */
                pageItem.left = originLeft + (originalPositions[j].left - originLeft) * factorX;
                pageItem.top = originTop - (originTop - originalPositions[j].top) * factorY;
            }
        }

        /**
         * 各アイテムを基準値に合わせてスケールする（縦横比保持／片辺のみ）
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @param {number} referenceValue - 目標寸法（pt）
         * @returns {void}
         */
        function resizeItemsToReference(mode, referenceValue) {
            var keepOneSideOnly = oneSideOnlyRadio.value;
            for (var i = 0; i < targetItems.length; i++) {
                var currentLength = measureSide(getReferenceBounds(targetItems[i]), mode);
                if (currentLength === 0) continue;
                var scalePct = getScalePercent(currentLength, referenceValue);

                if (!keepOneSideOnly) {
                    targetItems[i].resize(scalePct, scalePct, true, true, true, true, scalePct, Transformation.TOPLEFT);
                    appliedLineScales[i] = scalePct;
                } else {
                    /* 指定サイズも measureSide() が幅／高さを返すので、ここは共通で扱える / Fixed size shares this path */
                    resizeOneSide(targetItems[i], mode.isWidth, scalePct);
                    appliedLineScales[i] = 100;
                }
            }
        }

        /**
         * 選択全体をアートボードの中心へ移動する（裁ち落としのオフセットは上下左右対称なので、中心はアートボードと一致する）
         * @returns {void}
         */
        function centerItemsOnArtboard() {
            var artboardRect = getActiveArtboardRect();
            var centerX = (artboardRect[0] + artboardRect[2]) / 2;
            var centerY = (artboardRect[1] + artboardRect[3]) / 2;
            var clusterBounds = getCombinedReferenceBounds(targetItems);
            var dx = centerX - (clusterBounds.left + clusterBounds.width / 2);
            var dy = centerY - (clusterBounds.top - clusterBounds.height / 2);
            for (var i = 0; i < targetItems.length; i++) {
                targetItems[i].left += dx;
                targetItems[i].top += dy;
            }
        }

        /**
         * 面積を目標の面積に合わせる
         * @param {PageItem} pageItem - 対象
         * @param {number} targetArea - 目標の面積
         * @returns {number} 適用した線幅倍率（%）。復元時の逆算に使う
         */
        function resizeToMatchArea(pageItem, targetArea) {
            var bounds = getReferenceBounds(pageItem);
            var area = bounds.width * bounds.height;
            if (area === 0) return 100;
            var scalePct = Math.sqrt(targetArea / area) * 100;
            pageItem.resize(scalePct, scalePct, true, true, true, true, scalePct, Transformation.TOPLEFT);
            return scalePct;
        }

        /**
         * 面積基準：全アイテムの面積を最大／最小の面積に合わせる
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @returns {void}
         */
        function applyAreaResize(mode) {
            var areas = [];
            for (var i = 0; i < targetItems.length; i++) {
                var itemBounds = getReferenceBounds(targetItems[i]);
                areas.push(itemBounds.width * itemBounds.height);
            }
            var baseArea = mode.isAreaMax ? Math.max.apply(null, areas) : Math.min.apply(null, areas);
            if (!baseArea || baseArea <= 0) return;
            for (var j = 0; j < targetItems.length; j++) {
                appliedLineScales[j] = resizeToMatchArea(targetItems[j], baseArea);
            }
            captureResizeBaseState();
            app.redraw();
        }

        /**
         * 選択中の基準でリサイズする（呼び出し側は元のジオメトリへ戻してから呼ぶ）
         * @returns {void}
         */
        function applyResizeBySelection() {
            /* ここで基準状態も捨てておく。捨てないと、以降の早期 return（基準未選択・指定サイズが 0 など）で古いリサイズ結果が
               resizeBaseStates に残り、整列クリック時の restoreResizeBaseState() がそれを復元してしまう。
               Drop the stale base state first, or an early return would leave a previous resize to be restored. */
            resizeBaseStates = [];

            var mode = getSelectedResizeMode();
            if (!mode) return;
            clearOutlineBoundsCache();

            if (mode.isArea) {
                applyAreaResize(mode);
                return;
            }

            var referenceValue = computeReferenceValue(mode);
            if (referenceValue === null || referenceValue <= 0) return;

            if (mode.isArtboard || mode.isBleed) {
                /* 選択全体をアートボード（＋裁ち落とし）に合わせてスケールし、中心へ配置 / Fit the cluster and center it */
                applyClusterResize(mode, referenceValue);
                centerItemsOnArtboard();
            } else {
                /* 最大／最小／指定サイズ／長辺／短辺: 各アイテムを個別にスケール / Scale each item on its own */
                resizeItemsToReference(mode, referenceValue);
            }
            captureResizeBaseState();
            app.redraw();
        }

        // -----------------------------------------
        // 整列と分配 / Align and distribute
        // -----------------------------------------

        /**
         * 各アイテムの指定の辺（または中心）を targetValue へ動かす（位置は差分で動かす）
         * item.top / item.left は visibleBounds 基準なので、境界の値を直接代入すると
         * 「プレビュー境界で計測」OFF（geometricBounds）のとき線幅の半分ずれる。
         * @param {number} targetValue - 揃える座標
         * @param {string} edgeName - EDGE_OF のキー
         * @returns {void}
         */
        function shiftItemsToEdge(targetValue, edgeName) {
            var isHorizontal = HORIZONTAL_EDGES[edgeName] === true;
            for (var i = 0; i < targetItems.length; i++) {
                var delta = targetValue - EDGE_OF[edgeName](getReferenceBounds(targetItems[i]));
                if (isHorizontal) {
                    targetItems[i].left += delta;
                } else {
                    targetItems[i].top += delta;
                }
            }
        }

        /**
         * いちばん外側の辺に揃える（左・下は最小、右・上は最大）
         * @param {string} edgeName - "left" / "right" / "top" / "bottom"
         * @param {boolean} useMaximum - 最大値に揃えるなら true
         * @returns {void}
         */
        function alignToExtremeEdge(edgeName, useMaximum) {
            var extremeValue = null;
            for (var i = 0; i < targetItems.length; i++) {
                var edgeValue = EDGE_OF[edgeName](getReferenceBounds(targetItems[i]));
                if (extremeValue === null || (useMaximum ? edgeValue > extremeValue : edgeValue < extremeValue)) extremeValue = edgeValue;
            }
            if (extremeValue === null) return;
            shiftItemsToEdge(extremeValue, edgeName);
        }

        /**
         * 選択全体の境界の中心に揃える（各中心の平均ではなく、Illustrator 標準の整列と同じ基準）
         * @param {string} centerName - "centerX" / "centerY"
         * @returns {void}
         */
        function alignToSelectionCenter(centerName) {
            shiftItemsToEdge(EDGE_OF[centerName](getCombinedReferenceBounds(targetItems)), centerName);
        }

        /**
         * 縦方向に分配する（上から順に並べる。useGap=false で 0 間隔）
         * @param {boolean} useGap - 間隔を均等にするなら true、0 間隔なら false
         * @returns {void}
         */
        function distributeVertical(useGap) {
            var sortedItems = targetItems.slice(0).sort(function (itemA, itemB) {
                return getReferenceBounds(itemB).top - getReferenceBounds(itemA).top;
            });
            var topMost = getReferenceBounds(sortedItems[0]).top;
            var gap = 0;
            if (useGap) {
                var lastBounds = getReferenceBounds(sortedItems[sortedItems.length - 1]);
                var bottomMost = lastBounds.top - lastBounds.height;
                var totalHeight = 0;
                for (var i = 0; i < sortedItems.length; i++) totalHeight += getReferenceBounds(sortedItems[i]).height;
                gap = (topMost - bottomMost - totalHeight) / (sortedItems.length - 1);
            }
            var currentY = topMost;
            for (var j = 0; j < sortedItems.length; j++) {
                var bounds = getReferenceBounds(sortedItems[j]);
                sortedItems[j].top += currentY - bounds.top;
                currentY -= (bounds.height + gap);
            }
        }

        /**
         * 横方向に分配する（左から順に並べる。useGap=false で 0 間隔）
         * @param {boolean} useGap - 間隔を均等にするなら true、0 間隔なら false
         * @returns {void}
         */
        function distributeHorizontal(useGap) {
            var sortedItems = targetItems.slice(0).sort(function (itemA, itemB) {
                return getReferenceBounds(itemA).left - getReferenceBounds(itemB).left;
            });
            var leftMost = getReferenceBounds(sortedItems[0]).left;
            var gap = 0;
            if (useGap) {
                var lastBounds = getReferenceBounds(sortedItems[sortedItems.length - 1]);
                var rightMost = lastBounds.left + lastBounds.width;
                var totalWidth = 0;
                for (var i = 0; i < sortedItems.length; i++) totalWidth += getReferenceBounds(sortedItems[i]).width;
                gap = (rightMost - leftMost - totalWidth) / (sortedItems.length - 1);
            }
            var currentX = leftMost;
            for (var j = 0; j < sortedItems.length; j++) {
                var bounds = getReferenceBounds(sortedItems[j]);
                sortedItems[j].left += currentX - bounds.left;
                currentX += (bounds.width + gap);
            }
        }
    }

    main();
})();
