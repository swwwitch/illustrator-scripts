#target illustrator
#targetengine "CouponTicketMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した長方形パスから、ミシン目・ギザギザ・コーナー処理・スリット／ホールを組み合わせたチケット風の形状を生成します。
専用のプレビューレイヤーで結果を確認しながら設定でき、［OK］したときだけ元のオブジェクトへ適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CouponTicketMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n2e949946228a

### Overview

Turns a selected rectangle into a ticket-like shape combining perforations, zigzag edges, corner treatments and slits or holes.
The result is set up on a dedicated preview layer and applied to the original only when you confirm with OK.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CouponTicketMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CouponTicketMaker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CouponTicketMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CouponTicketMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2e949946228a"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 塗りも線も持たないオブジェクトに使う代替色の濃度（%）/ Tint used when an object has neither fill nor stroke */
    var FALLBACK_GRAY_TINT = 60;

    /* 逆角丸で四隅だけを残すための破線間隔（pt）/ Dash gap that leaves only the corners for inverse rounding */
    var INVERSE_CORNER_DASH_GAP = 1000;

    /* 長方形判定に使う座標の許容誤差（pt）/ Coordinate tolerance for the rectangle check */
    var RECTANGLE_TOLERANCE = 0.01;

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

    var PRESET_ROW_MARGINS   = [20, 0, 20, 0];     /* プリセット行の余白 */
    var PRESET_ROW_SPACING   = 12;                 /* プリセット行の要素間隔 */
    var SAVE_BUTTON_SIZE     = [70, 24];           /* ［保存］ボタンのサイズ */
    var OFFSET_SLIDER_WIDTH  = 200;                /* 分割位置スライダーの幅 */
    var NUMBER_FIELD_CHARS   = 3;                  /* 数値入力欄の標準幅（文字数） */
    var OFFSET_FIELD_CHARS   = 5;                  /* 分割位置入力欄の幅（文字数） */
    var INSET_FIELD_CHARS    = 4;                  /* 負値を入れる入力欄の幅（文字数） */
    var ZIGZAG_LABEL_WIDTH   = { ja: 60, en: 55 }; /* ギザギザパネルのラベル幅 */

    /**
     * ラベル付きパネルを生成し、共通レイアウトを適用する
     * @param {Window|Panel|Group} parent - 追加先のコンテナ
     * @param {string} titleText - パネルのタイトル
     * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "top"]）
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parent, titleText, alignChildren) {
        var panel = parent.add('panel', undefined, titleText);
        setupPanel(panel, 6);
        if (alignChildren) panel.alignChildren = alignChildren;
        return panel;
    }

    /**
     * 横並びのグループを生成する
     * @param {Window|Panel|Group} parent - 追加先のコンテナ
     * @returns {Group} 生成したグループ
     */
    function addRow(parent) {
        var row = parent.add('group');
        row.orientation = 'row';
        row.alignChildren = ['left', 'center'];
        return row;
    }

    /**
     * 縦積みのグループを生成する
     * @param {Window|Panel|Group} parent - 追加先のコンテナ
     * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "top"]）
     * @returns {Group} 生成したグループ
     */
    function addColumn(parent, alignChildren) {
        var column = parent.add('group');
        column.orientation = 'column';
        column.alignChildren = alignChildren || ['fill', 'top'];
        return column;
    }

    /**
     * 「ラベル＋∧∨＋数値入力欄＋単位」の1行を生成する
     * @param {Panel|Group} parent - 追加先のコンテナ
     * @param {Object} options - 生成オプション
     * @param {string} options.label - ラベル文言
     * @param {string} options.value - 入力欄の初期値
     * @param {string} options.unit - 単位表記（不要なら省略）
     * @param {number} options.labelWidth - ラベル幅（省略時は指定しない）
     * @param {number} options.characters - 入力欄の幅（文字数、省略時は NUMBER_FIELD_CHARS）
     * @param {Object} options.stepOptions - ∧∨と↑↓キーの増減指定（min / max / integer。省略時は min: 0）
     * @param {function} options.onChange - ∧∨・↑↓キーで値が変わったときに呼ぶ関数
     * @returns {Object} { row: Group, input: EditText }
     */
    function addNumberField(parent, options) {
        var row = addRow(parent);
        var label = row.add('statictext', undefined, options.label);
        if (options.labelWidth) label.preferredSize.width = options.labelWidth;

        var input = addSteppedInput(row, options.value, options.characters || NUMBER_FIELD_CHARS,
            options.stepOptions || { min: 0 }, options.onChange);

        if (options.unit) row.add('statictext', undefined, options.unit);

        return { row: row, input: input };
    }

    /**
     * ∧∨と入力欄を隙間なく並べて追加し、↑↓キーも∧∨と同じ処理で増減させる
     * @param {Group} parent - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {number} characters - 入力欄の幅（文字数）
     * @param {Object} stepOptions - min / max / integer（どれも省略可）
     * @param {function} onStep - 増減したときに呼ぶ関数（省略可）
     * @returns {EditText} 入力欄
     */
    function addSteppedInput(parent, initialText, characters, stepOptions, onStep) {
        var stepperInputGroup = parent.add('group');
        stepperInputGroup.orientation = 'row';
        stepperInputGroup.alignChildren = ['left', 'center'];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var input;
        stepOptions.onStep = function () { if (onStep) onStep(); };
        var stepperGroup = addStepper(stepperInputGroup, function () { return input; }, stepOptions);
        input = stepperInputGroup.add('edittext', undefined, initialText);
        input.characters = characters;
        bindSteppedArrowKeys(input, stepperGroup);
        return input;
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

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "チケットメーカー", en: "Ticket Maker" }
        },
        panel: {
            sides:           { ja: "左右", en: "L/R" },
            sidePerforation: { ja: "ミシン目", en: "Perforation" },
            zigzag:          { ja: "ギザギザ", en: "Zigzag" },
            centerSplit:     { ja: "左右分割", en: "Center Divider" },
            divider:         { ja: "分割線", en: "Divider Line" },
            edge:            { ja: "エッジ", en: "Edge" },
            corner:          { ja: "コーナー", en: "Corner" },
            hole:            { ja: "スリット／ホール", en: "Slit / Hole" }
        },
        tooltip: {
            preset:        { ja: "保存したチケットの形を読み込みます。", en: "Loads a saved ticket shape." },
            cornerNone:    { ja: "角は加工しません。", en: "Leaves the corners square." },
            cornerRound:   { ja: "角を丸めます。", en: "Rounds the corners." },
            cornerInverse: { ja: "角を内側にえぐった形にします。", en: "Scoops the corners inward." },
            cornerChamfer: { ja: "角を面取りします。", en: "Chamfers the corners." },
            zigzagNone:    { ja: "辺はまっすぐのままにします。", en: "Leaves the edges straight." },
            zigzagLeftRight: { ja: "左右の辺をギザギザにします。", en: "Makes the left and right edges jagged." },
            zigzagTopBottom: { ja: "上下の辺をギザギザにします。", en: "Makes the top and bottom edges jagged." },
            sidePerforation: { ja: "左右の辺にミシン目を入れます。", en: "Adds a perforation along the left and right edges." },
            linkToCenter:  { ja: "ミシン目の設定を中央の切り取り線と連動させます。", en: "Links the perforation settings to the centre tear line." },
            holeNone:      { ja: "穴は開けません。", en: "Punches no hole." },
            holeCircle:    { ja: "丸い穴を開けます。", en: "Punches a round hole." },
            holeTriangle:  { ja: "三角の穴を開けます。", en: "Punches a triangular hole." },
            holeSide:      { ja: "この側に穴を開けます。", en: "Punches the hole on this side." },
            centerSplit:   { ja: "中央に切り取り線を入れて、券面を分けます。", en: "Adds a tear line down the middle to split the ticket." },
            centerOffset:  { ja: "切り取り線の位置を中央からずらす量です。", en: "How far the tear line sits from the centre." },
            dividerDot:    { ja: "切り取り線を点線にします。", en: "Draws the tear line as dots." },
            dividerDash:   { ja: "切り取り線を破線にします。", en: "Draws the tear line as dashes." },
            edgeNone:      { ja: "端の飾りを付けません。", en: "Adds no edge notches." },
            edgeCircle:    { ja: "端に半円の切り欠きを入れます。", en: "Cuts semicircular notches into the edge." },
            edgeTriangle:  { ja: "端に三角の切り欠きを入れます。", en: "Cuts triangular notches into the edge." },
            edgeDoubleRound: { ja: "切り欠きを二重の円にします。", en: "Uses a double circle for the notch." },
            edgesOnly:     { ja: "切り欠きだけを作り、券面の枠は描きません。", en: "Creates only the notches, without the ticket outline." },
            preview:       { ja: "結果を画面で確認します。キャンセルすると元に戻ります。", en: "Shows the result on the canvas. Cancel restores the original state." },
            expandAppearance: { ja: "作った形のアピアランスを分割・拡張して、実体のあるパスにします。", en: "Expands the appearance so the shape becomes real paths." },
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
        radio: {
            none:      { ja: "なし", en: "None" },
            dot:       { ja: "ドット", en: "Dot" },
            dash:      { ja: "破線", en: "Dash" },
            circle:    { ja: "円", en: "Circle" },
            triangle:  { ja: "三角", en: "Triangle" },
            round:     { ja: "角丸", en: "Rounded" },
            inverse:   { ja: "逆角丸", en: "Inverse Round" },
            chamfer:   { ja: "面取り", en: "Chamfer" },
            leftRight: { ja: "左右", en: "L/R" },
            topBottom: { ja: "上下", en: "T/B" }
        },
        checkbox: {
            enable:           { ja: "有効", en: "Enable" },
            linkToCenter:     { ja: "分割線に連動", en: "Link to Split Line" },
            left:             { ja: "左", en: "Left" },
            right:            { ja: "右", en: "Right" },
            doubleRound:      { ja: "ダブル角丸", en: "Double Rounded" },
            edgesOnly:        { ja: "エッジのみ", en: "Edges Only" },
            preview:          { ja: "プレビュー", en: "Preview" },
            expandAppearance: { ja: "アピアランスを分割", en: "Expand Appearance" }
        },
        fieldLabel: {
            lineWidth:    { ja: "線幅", en: "Weight" },
            gap:          { ja: "間隔", en: "Gap" },
            inset:        { ja: "長さ", en: "Inset Length" },
            size:         { ja: "サイズ", en: "Size" },
            zigzagSize:   { ja: "大きさ", en: "Size" },
            zigzagRepeat: { ja: "繰り返し", en: "Repeat" }
        },
        button: {
            save:       { ja: "保存", en: "Save" },
            cancel:     { ja: "キャンセル", en: "Cancel" },
            ok:         { ja: "OK", en: "OK" },
            outlineOn:  { ja: "アウトライン表示", en: "Outline View" },
            outlineOff: { ja: "プレビュー表示", en: "Preview View" }
        },
        alert: {
            openDocument:      { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectRectangle:   { ja: "長方形を選択してください。", en: "Please select a rectangle." },
            rectangleOnly:     { ja: "長方形のパスを1つだけ選択してください。", en: "Please select exactly one rectangular path." },
            presetName:        { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
            singleOnly: {
                ja: "複数選択時は実行できません。オブジェクトを1つだけ選択してください。",
                en: "This script cannot run with multiple selections. Please select only one object."
            },
            groupNotAllowed: {
                ja: "グループを選択しているときは実行できません。単体のオブジェクトを選択してください。",
                en: "This script cannot run when a group is selected. Please select a single object."
            },
            enterValidNumbers: {
                ja: "数値欄に正しい数値を入力してください。",
                en: "Please enter valid numeric values in the numeric fields."
            },
            presetSaved: {
                ja: "プリセットを保存しました。\n\nコード組み込み用:\n",
                en: "Preset saved.\n\nCode snippet for embedding:\n"
            }
        },
        layerName: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fallbackName: {
            customPreset: { ja: "カスタム", en: "Custom" }
        }
    };

    // =========================================
    // 単位ユーティリティ / Unit utilities
    // =========================================

    /* 環境設定の単位コードと表記の対応 / Unit code to label */
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
     * 指定単位の値をポイントへ変換する
     * @param {number} value - 変換前の値
     * @param {number} factor - 1単位あたりのポイント数
     * @returns {number} ポイント値
     */
    function toPt(value, factor) {
        return value * factor;
    }

    /**
     * ポイント値を指定単位へ変換する
     * @param {number} value - ポイント値
     * @param {number} factor - 1単位あたりのポイント数
     * @returns {number} 変換後の値
     */
    function fromPt(value, factor) {
        return value / factor;
    }

    var rulerUnit = getUnitInfo('rulerType');
    var rulerUnitLabel = rulerUnit.label;
    var rulerPtFactor = rulerUnit.pointsPerUnit;

    /* 分割線の線幅だけは線の単位に従う / Stroke units apply to the divider line weight only */
    var strokeUnit = getUnitInfo('strokeUnits');
    var strokeUnitLabel = strokeUnit.label;
    var strokePtFactor = strokeUnit.pointsPerUnit;

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

    /* 破線を「線の位置：中央」で描くための一時アクション定義 / Action data that draws dashes with centered alignment */
    var STROKE_DOT_ACTION = '/version 3 /name [ 9 5374726f6b65446f74 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 3 646f74 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_setStroke) /localizedName [ 12 e7b79ae38292e8a8ade5ae9a ] /isOpen 0 /isOn 1 /hasDialog 0 /parameterCount 12 /parameter-1 { /key 2003072104 /showInPalette 4294967295 /type (unit real) /value 4.0 /unit 592476268 } /parameter-2 { /key 1667330094 /showInPalette 4294967295 /type (enumerated) /name [ 12 e4b8b8e59e8be7b79ae7abaf ] /value 1 } /parameter-3 { /key 1836344690 /showInPalette 4294967295 /type (real) /value 10.0 } /parameter-4 { /key 1785686382 /showInPalette 4294967295 /type (enumerated) /name [ 18 e3839ee382a4e382bfe383bce7b590e59088 ] /value 0 } /parameter-5 { /key 1684825454 /showInPalette 4294967295 /type (integer) /value 2 } /parameter-6 { /key 1685284913 /showInPalette 4294967295 /type (unit real) /value 0.0 /unit 592476268 } /parameter-7 { /key 1685284914 /showInPalette 4294967295 /type (unit real) /value 6.0 /unit 592476268 } /parameter-8 { /key 1684104298 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-9 { /key 1634231345 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-10 { /key 1634231346 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-11 { /key 1634230636 /showInPalette 4294967295 /type (enumerated) /name [ 24 e38391e382b9e381aee7b582e782b9e381abe9858de7bdae ] /value 0 } /parameter-12 { /key 1634494318 /showInPalette 4294967295 /type (enumerated) /name [ 6 e4b8ade5a4ae ] /value 0 } } }';

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

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

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    var isStrokeDotActionLoaded = false;

    /**
     * StrokeDot アクションを一時ファイル経由で読み込む（読み込み済みなら何もしない）。
     * プレビューのたびに読み込まないよう、解除は main() の終わりで1回だけ行う
     * @returns {void}
     */
    function loadStrokeDotAction() {
        if (isStrokeDotActionLoaded) return;
        /* 失敗は従来どおり例外で伝える / Report a failure as an exception, as before */
        if (!loadTemporaryActionSet(STROKE_DOT_ACTION, 'StrokeDot')) {
            throw new Error('Could not load the action: StrokeDot');
        }
        isStrokeDotActionLoaded = true;
    }

    /**
     * StrokeDot アクションを破棄する
     * @returns {void}
     */
    function unloadStrokeDotAction() {
        if (!isStrokeDotActionLoaded) return;
        unloadTemporaryActionSet('StrokeDot');
        isStrokeDotActionLoaded = false;
    }

    /**
     * 対象パスへ StrokeDot アクションを適用する（選択状態は元に戻す）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} item - 対象パス
     * @returns {void}
     */
    function applyStrokeDotAction(doc, item) {
        loadStrokeDotAction();
        var previousSelection = saveSelection(doc);
        try {
            selectItems(doc, [item]);
            app.doScript('dot', 'StrokeDot', false);
        } finally {
            restoreSelection(doc, previousSelection);
        }
    }

    // =========================================
    // 選択・オブジェクト操作 / Selection and object helpers
    // =========================================

    /**
     * 現在の選択を安全に取得する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択オブジェクトの配列（取得できないときは空配列）
     */
    function getSafeSelection(doc) {
        try {
            return doc.selection || [];
        } catch (e) {
            return [];
        }
    }

    /**
     * 現在の選択を配列として控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択オブジェクトの配列
     */
    function saveSelection(doc) {
        var items = [];
        var selection = getSafeSelection(doc);
        for (var i = 0; i < selection.length; i++) items.push(selection[i]);
        return items;
    }

    /**
     * 控えておいた選択を復元する（削除済みのオブジェクトは読み飛ばす）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - 復元するオブジェクトの配列
     * @returns {void}
     */
    function restoreSelection(doc, items) {
        doc.selection = null;
        if (!items) return;
        for (var i = 0; i < items.length; i++) {
            try {
                items[i].selected = true;
            } catch (e) {
                /* 生成前に消えたオブジェクトは無視 / Skip items that no longer exist */
            }
        }
    }

    /**
     * 指定オブジェクトだけを選択する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - 選択するオブジェクトの配列
     * @returns {void}
     */
    function selectItems(doc, items) {
        doc.selection = null;
        for (var i = 0; i < items.length; i++) {
            items[i].selected = true;
        }
    }

    /**
     * 選択が1つだけのときにそのオブジェクトを返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} 先頭の選択オブジェクト（なければ null）
     */
    function getSingleSelection(doc) {
        var selection = getSafeSelection(doc);
        return selection.length > 0 ? selection[0] : null;
    }

    /**
     * 線をアウトライン化する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} item - 対象パス
     * @returns {Object} アウトライン化後のオブジェクト
     */
    function outlineStrokeItem(doc, item) {
        var previousSelection = saveSelection(doc);
        try {
            selectItems(doc, [item]);
            app.executeMenuCommand('Live Outline Stroke');
            var outlined = getSingleSelection(doc);
            if (!outlined) throw new Error('Live Outline Stroke failed.');
            return outlined;
        } finally {
            restoreSelection(doc, previousSelection);
        }
    }

    /**
     * 複数のオブジェクトをグループ化する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - グループ化するオブジェクトの配列
     * @returns {Object} グループ（1つだけならそのオブジェクト、空なら null）
     */
    function groupItems(doc, items) {
        if (!items || items.length === 0) return null;
        if (items.length === 1) return items[0];

        var previousSelection = saveSelection(doc);
        try {
            selectItems(doc, items);
            app.executeMenuCommand('group');
            var grouped = getSingleSelection(doc);
            if (!grouped) throw new Error('Group command failed.');
            return grouped;
        } finally {
            restoreSelection(doc, previousSelection);
        }
    }

    /**
     * ドキュメントのカラースペースに応じたグレーを生成する
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|GrayColor} 生成した色
     */
    function makeFallbackGray(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = FALLBACK_GRAY_TINT;
            return cmyk;
        }
        var gray = new GrayColor();
        gray.gray = FALLBACK_GRAY_TINT;
        return gray;
    }

    /**
     * パスファインダーで抜けるように、対象を「塗りのみ」の状態へそろえる
     * @param {Array} items - 対象オブジェクトの配列
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function normalizeInitialAppearance(items, doc) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];

            if (item.filled) {
                item.stroked = false;
                continue;
            }

            if (item.stroked) {
                try {
                    item.fillColor = item.strokeColor;
                } catch (e) {
                    item.fillColor = makeFallbackGray(doc);
                }
            } else {
                item.fillColor = makeFallbackGray(doc);
            }
            item.filled = true;
            item.stroked = false;
        }
    }

    /**
     * 長方形パスかどうかを判定する
     * @param {Object} item - 判定するオブジェクト
     * @returns {boolean} 4隅がバウンディングボックスに一致する長方形なら true
     */
    function isRectanglePath(item) {
        if (!item || item.typename !== 'PathItem') return false;
        if (item.guides || item.clipping) return false;
        if (item.pathPoints.length !== 4) return false;

        var geometricBounds = item.geometricBounds;
        var left = geometricBounds[0];
        var top = geometricBounds[1];
        var right = geometricBounds[2];
        var bottom = geometricBounds[3];

        var hasLeftTop = false, hasRightTop = false, hasRightBottom = false, hasLeftBottom = false;

        for (var i = 0; i < 4; i++) {
            var anchor = item.pathPoints[i].anchor;
            var isLeft = Math.abs(anchor[0] - left) <= RECTANGLE_TOLERANCE;
            var isRight = Math.abs(anchor[0] - right) <= RECTANGLE_TOLERANCE;
            var isTop = Math.abs(anchor[1] - top) <= RECTANGLE_TOLERANCE;
            var isBottom = Math.abs(anchor[1] - bottom) <= RECTANGLE_TOLERANCE;

            if (isLeft && isTop) hasLeftTop = true;
            else if (isRight && isTop) hasRightTop = true;
            else if (isRight && isBottom) hasRightBottom = true;
            else if (isLeft && isBottom) hasLeftBottom = true;
            else return false;
        }

        return hasLeftTop && hasRightTop && hasRightBottom && hasLeftBottom;
    }

    // =========================================
    // 形状の生成 / Shape builders
    // =========================================

    /**
     * 角丸のライブエフェクトを適用する
     * @param {Array} items - 対象オブジェクトの配列
     * @param {number} radiusPt - 角丸の半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(items, radiusPt) {
        if (!(radiusPt > 0)) return;
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>';
        for (var i = 0; i < items.length; i++) {
            items[i].applyEffect(xml);
        }
    }

    /**
     * 塗りだけを持つ円を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {number} centerX - 中心のX座標（pt）
     * @param {number} centerY - 中心のY座標（pt）
     * @param {number} sizePt - 直径（pt）
     * @returns {PathItem} 生成した円
     */
    function addFilledCircle(doc, centerX, centerY, sizePt) {
        var circle = doc.pathItems.ellipse(centerY + sizePt / 2, centerX - sizePt / 2, sizePt, sizePt);
        circle.filled = true;
        circle.stroked = false;
        return circle;
    }

    /**
     * 塗りだけを持つ正方形を45°回転させて追加する（三角スリット・ギザギザ用）
     * @param {Document} doc - 対象ドキュメント
     * @param {number} centerX - 中心のX座標（pt）
     * @param {number} centerY - 中心のY座標（pt）
     * @param {number} sizePt - 一辺の長さ（pt）
     * @returns {PathItem} 生成したひし形
     */
    function addFilledDiamond(doc, centerX, centerY, sizePt) {
        var diamond = doc.pathItems.rectangle(centerY + sizePt / 2, centerX - sizePt / 2, sizePt, sizePt);
        diamond.filled = true;
        diamond.stroked = false;
        diamond.rotate(45);
        return diamond;
    }

    /**
     * 線だけを持つ直線を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<Array<number>>} points - 始点と終点の座標
     * @returns {PathItem} 生成した直線
     */
    function addStrokedLine(doc, points) {
        var line = doc.pathItems.add();
        line.setEntirePath(points);
        line.filled = false;
        line.stroked = true;
        return line;
    }

    /**
     * 破線の長さが間隔にできるだけ近づくよう、破線パターンを求める
     * @param {number} lineLength - 破線を敷く全体長（pt）
     * @param {number} gapPt - 間隔（pt）
     * @returns {Array<number>} [破線長, 間隔] の配列（求まらなければ null）
     */
    function computeDashPattern(lineLength, gapPt) {
        if (!(gapPt > 0) || !(lineLength > 0)) return null;

        var targetGapCount = (lineLength - gapPt) / (gapPt * 2);
        var candidates = [
            Math.floor(targetGapCount),
            Math.ceil(targetGapCount),
            Math.floor(targetGapCount) - 1,
            Math.ceil(targetGapCount) + 1,
            1
        ];
        var bestDashLength = 0;
        var bestDifference = Number.MAX_VALUE;

        for (var i = 0; i < candidates.length; i++) {
            var gapCount = Math.max(1, candidates[i]);
            var dashLength = (lineLength - gapCount * gapPt) / (gapCount + 1);
            if (dashLength <= 0) continue;

            var difference = Math.abs(dashLength - gapPt);
            if (difference < bestDifference) {
                bestDifference = difference;
                bestDashLength = dashLength;
            }
        }

        return bestDashLength > 0 ? [bestDashLength, gapPt] : null;
    }

    /**
     * 中央の分割線（ミシン目／破線）を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - パスファインダー対象を追加する配列
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function addCenterPerforation(doc, items, bounds, settings, geometry) {
        if (!settings.center.enabled) return;
        if (settings.edge.shapeOnly && settings.edge.mode !== 'none') return;

        var insetPt = geometry.centerInsetPt;
        var line = addStrokedLine(doc, [
            [bounds.centerLineX, bounds.top - insetPt],
            [bounds.centerLineX, bounds.top - bounds.height + insetPt]
        ]);

        var useDotStyle = (settings.center.mode !== 'dash');
        if (useDotStyle || insetPt === 0) {
            applyStrokeDotAction(doc, line);
        }

        line.strokeWidth = geometry.centerWidthPt;
        if (useDotStyle) {
            line.strokeDashes = [0, geometry.centerGapPt];
        } else {
            line.strokeCap = StrokeCap.BUTTENDCAP;
            var dashPattern = (insetPt !== 0)
                ? computeDashPattern(Math.max(0, bounds.height - insetPt * 2), geometry.centerGapPt)
                : null;
            line.strokeDashes = dashPattern || [geometry.centerGapPt, geometry.centerGapPt];
        }

        items.push(outlineStrokeItem(doc, line));
    }

    /**
     * 左右のミシン目を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - パスファインダー対象を追加する配列
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function addSidePerforation(doc, items, bounds, settings, geometry) {
        if (!settings.lr.enabled) return;

        var insetPt = geometry.sideInsetPt;
        var positions = [bounds.left, bounds.left + bounds.width];

        for (var i = 0; i < positions.length; i++) {
            var line = addStrokedLine(doc, [
                [positions[i], bounds.top - insetPt],
                [positions[i], bounds.top - bounds.height + insetPt]
            ]);
            applyStrokeDotAction(doc, line);
            line.strokeWidth = geometry.sideWidthPt;
            line.strokeDashes = [0, geometry.sideGapPt];
            items.push(outlineStrokeItem(doc, line));
        }
    }

    /**
     * 辺のギザギザ（45°回転した正方形の連続）を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - パスファインダー対象を追加する配列
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function addZigzag(doc, items, bounds, settings, geometry) {
        if (settings.zigzag.mode === 'none' || !(geometry.zigzagSizePt > 0)) return;

        /* 指定サイズは対角線の長さなので、正方形の辺に変換する / The size is the diagonal, convert it to a side */
        var side = geometry.zigzagSizePt / Math.sqrt(2);
        var step = geometry.zigzagSizePt + geometry.zigzagGapPt;
        var total = step * geometry.zigzagRepeat - geometry.zigzagGapPt;
        var zigzagItems = [];
        var lineIndex, repeatIndex;

        if (settings.zigzag.mode === 'lr') {
            /* 左右：辺の垂直中央を基準に配置 / Left and right: centered vertically on each side */
            var startY = (bounds.top - bounds.height / 2) + total / 2 - geometry.zigzagSizePt / 2;
            var xPositions = [bounds.left, bounds.left + bounds.width];
            for (lineIndex = 0; lineIndex < xPositions.length; lineIndex++) {
                for (repeatIndex = 0; repeatIndex < geometry.zigzagRepeat; repeatIndex++) {
                    zigzagItems.push(addFilledDiamond(doc, xPositions[lineIndex], startY - step * repeatIndex, side));
                }
            }
        } else {
            /* 上下：辺の水平中央を基準に配置 / Top and bottom: centered horizontally on each side */
            var startX = (bounds.left + bounds.width / 2) - total / 2 + geometry.zigzagSizePt / 2;
            var yPositions = [bounds.top, bounds.top - bounds.height];
            for (lineIndex = 0; lineIndex < yPositions.length; lineIndex++) {
                for (repeatIndex = 0; repeatIndex < geometry.zigzagRepeat; repeatIndex++) {
                    zigzagItems.push(addFilledDiamond(doc, startX + step * repeatIndex, yPositions[lineIndex], side));
                }
            }
        }

        if (zigzagItems.length > 0) items.push(groupItems(doc, zigzagItems));
    }

    /**
     * 分割線の両端に置くエッジ形状を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - パスファインダー対象を追加する配列
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function addCenterEdges(doc, items, bounds, settings, geometry) {
        if (!settings.center.enabled) return;
        if (settings.edge.mode !== 'circle' && settings.edge.mode !== 'triangle') return;
        if (!(geometry.edgeSizePt > 0)) return;

        var yPositions = [bounds.top, bounds.top - bounds.height];
        for (var i = 0; i < yPositions.length; i++) {
            items.push(settings.edge.mode === 'circle'
                ? addFilledCircle(doc, bounds.centerLineX, yPositions[i], geometry.edgeSizePt)
                : addFilledDiamond(doc, bounds.centerLineX, yPositions[i], geometry.edgeSizePt));
        }
    }

    /**
     * 左右のスリット／ホールを追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - パスファインダー対象を追加する配列
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function addHoles(doc, items, bounds, settings, geometry) {
        if (settings.hole.mode === 'none' || !(geometry.holeSizePt > 0)) return;

        var centerY = bounds.top - bounds.height / 2;
        var xPositions = [];
        if (settings.hole.left) xPositions.push(bounds.left);
        if (settings.hole.right) xPositions.push(bounds.left + bounds.width);

        var holeItems = [];
        for (var i = 0; i < xPositions.length; i++) {
            holeItems.push(settings.hole.mode === 'circle'
                ? addFilledCircle(doc, xPositions[i], centerY, geometry.holeSizePt)
                : addFilledDiamond(doc, xPositions[i], centerY, geometry.holeSizePt));
        }

        if (holeItems.length > 0) items.push(groupItems(doc, holeItems));
    }

    /**
     * 四隅の逆角丸を追加する（上下辺に太い破線を敷いて角だけを削る）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - パスファインダー対象を追加する配列
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function addInverseCorners(doc, items, bounds, settings, geometry) {
        if (settings.corner.mode !== 'inverse' || !(geometry.cornerSizePt > 0)) return;

        var right = bounds.left + bounds.width;
        var yPositions = [bounds.top, bounds.top - bounds.height];

        for (var i = 0; i < yPositions.length; i++) {
            var line = addStrokedLine(doc, [[bounds.left, yPositions[i]], [right, yPositions[i]]]);
            applyStrokeDotAction(doc, line);
            line.strokeWidth = geometry.cornerSizePt;
            line.strokeDashes = [0, INVERSE_CORNER_DASH_GAP];
            items.push(outlineStrokeItem(doc, line));
        }
    }

    /**
     * 四隅の面取りを追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - パスファインダー対象を追加する配列
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function addChamferCorners(doc, items, bounds, settings, geometry) {
        if (settings.corner.mode !== 'chamfer' || !(geometry.cornerSizePt > 0)) return;

        var right = bounds.left + bounds.width;
        var bottom = bounds.top - bounds.height;
        var corners = [
            [bounds.left, bounds.top],
            [right, bounds.top],
            [bounds.left, bottom],
            [right, bottom]
        ];

        for (var i = 0; i < corners.length; i++) {
            items.push(addFilledDiamond(doc, corners[i][0], corners[i][1], geometry.cornerSizePt));
        }
    }

    /**
     * ダブル角丸を適用する（分割線で2分割し、それぞれに角丸をかけて合体する）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} rect - 元の長方形
     * @param {Array} items - パスファインダー対象の配列（先頭を差し替える）
     * @param {Object} bounds - 長方形の位置とサイズ
     * @param {Object} settings - UIから読み取った設定
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {void}
     */
    function applyDoubleRoundEdge(doc, rect, items, bounds, settings, geometry) {
        if (!(geometry.edgeSizePt > 0)) return;

        var leftRect = doc.pathItems.rectangle(bounds.top, bounds.left, bounds.centerLineX - bounds.left, bounds.height);
        var rightRect = doc.pathItems.rectangle(bounds.top, bounds.centerLineX, bounds.left + bounds.width - bounds.centerLineX, bounds.height);
        var halves = [leftRect, rightRect];

        for (var i = 0; i < halves.length; i++) {
            halves[i].filled = true;
            halves[i].stroked = false;
        }
        try {
            leftRect.fillColor = rect.fillColor;
            rightRect.fillColor = rect.fillColor;
        } catch (e) {
            /* 元の塗りを引き継げないときは既定色のまま / Keep the default fill when it cannot be copied */
        }

        /* エッジサイズの角丸に、コーナーパネルの角丸を重ねる / Stack the corner radius on top of the edge radius */
        applyRoundCornersEffect(halves, geometry.edgeSizePt);
        if (settings.corner.mode === 'round') {
            applyRoundCornersEffect(halves, geometry.cornerSizePt);
        }

        rect.remove();
        selectItems(doc, [groupItems(doc, halves)]);
        app.executeMenuCommand('Live Pathfinder Add');
        items[0] = getSingleSelection(doc);
    }

    /**
     * 生成した形状をまとめて前面オブジェクトで型抜きする
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - 型抜き対象の配列（先頭がベースの長方形）
     * @param {Array} results - 結果を追加する配列
     * @returns {void}
     */
    function finalizeSubtract(doc, items, results) {
        selectItems(doc, [groupItems(doc, items)]);
        app.executeMenuCommand('Live Pathfinder Subtract');
        var subtracted = getSingleSelection(doc);
        if (subtracted) results.push(subtracted);
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    var PRESETS = [
        {
            name:    { ja: "0: クリア", en: "0: Clear" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: false, offset: 30, mode: "dot", width: 3, gap: 6, length: -10 },
            zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
            corner:  { mode: "none", size: 5 },
            hole:    { mode: "none", left: true, right: true, size: 10 },
            edge:    { mode: "wround", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "1: W角丸＋分割線", en: "1: Double Rounded + Divider" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: true, offset: 30, mode: "dot", width: 3, gap: 3, length: -10 },
            zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
            corner:  { mode: "none", size: 5 },
            hole:    { mode: "none", left: true, right: true, size: 10 },
            edge:    { mode: "wround", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "2: 左右に2ホール", en: "2: Two Side Holes" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: false, offset: 30, mode: "dot", width: 3, gap: 6, length: -6 },
            zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
            corner:  { mode: "none", size: 5 },
            hole:    { mode: "circle", left: true, right: true, size: 15 },
            edge:    { mode: "none", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "3: ミシン目", en: "3: Perforation" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: true, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
            zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
            corner:  { mode: "round", size: 6 },
            hole:    { mode: "none", left: true, right: true, size: 10 },
            edge:    { mode: "none", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "4: 四隅に逆角丸", en: "4: Inverse Round Corners" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: false, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
            zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
            corner:  { mode: "inverse", size: 25 },
            hole:    { mode: "none", left: true, right: true, size: 10 },
            edge:    { mode: "none", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "5: 左右に三角スリット", en: "5: Triangular Side Slits" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: false, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
            zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
            corner:  { mode: "round", size: 3 },
            hole:    { mode: "triangle", left: true, right: true, size: 15 },
            edge:    { mode: "none", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "6: ミシン目＋ギザギザ上下", en: "6: Perforation + T/B Zigzag" },
            lr:      { enabled: true, linkCenter: false, width: 4, gap: 7, length: 0 },
            center:  { enabled: false, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
            zigzag:  { mode: "tb", size: 10, gap: 0, repeat: 7 },
            corner:  { mode: "none", size: 3 },
            hole:    { mode: "none", left: true, right: true, size: 15 },
            edge:    { mode: "none", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "7: 面取り＋分割線（破線）", en: "7: Chamfer + Dashed Divider" },
            lr:      { enabled: false, linkCenter: false, width: 4, gap: 7, length: 0 },
            center:  { enabled: true, offset: 35, mode: "dash", width: 1, gap: 3, length: 0 },
            zigzag:  { mode: "none", size: 10, gap: 0, repeat: 7 },
            corner:  { mode: "chamfer", size: 13 },
            hole:    { mode: "none", left: true, right: true, size: 15 },
            edge:    { mode: "none", size: 5, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "8: 分割線＋エッジ（円）＋ギザギザ", en: "8: Divider + Circle Edges + Zigzag" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: true, offset: 30, mode: "dot", width: 3, gap: 6, length: -13 },
            zigzag:  { mode: "lr", size: 7, gap: 0, repeat: 7 },
            corner:  { mode: "none", size: 5 },
            hole:    { mode: "none", left: true, right: true, size: 10 },
            edge:    { mode: "circle", size: 10, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        },
        {
            name:    { ja: "9: 分割線（破線）＋三角エッジ＋左ホール", en: "9: Dashed Divider + Triangle Edges + Left Hole" },
            lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
            center:  { enabled: true, offset: 40, mode: "dash", width: 1, gap: 3, length: -7.2 },
            zigzag:  { mode: "none", size: 7, gap: 0, repeat: 7 },
            corner:  { mode: "round", size: 3 },
            hole:    { mode: "circle", left: true, right: false, size: 20 },
            edge:    { mode: "triangle", size: 6, shapeOnly: false },
            preview: { enabled: true, expandAppearance: false }
        }
    ];

    /**
     * プリセットや設定オブジェクトを再帰的に複製する
     * @param {*} value - 複製する値
     * @returns {*} 複製した値
     */
    function cloneSimpleValue(value) {
        if (value === null || typeof value !== 'object') return value;
        if (value instanceof Array) {
            var cloned = [];
            for (var i = 0; i < value.length; i++) cloned.push(cloneSimpleValue(value[i]));
            return cloned;
        }
        var out = {};
        for (var key in value) {
            if (value.hasOwnProperty(key)) out[key] = cloneSimpleValue(value[key]);
        }
        return out;
    }

    /**
     * 有限な数値かどうかを判定する
     * @param {*} value - 判定する値
     * @returns {boolean} 有限な数値なら true
     */
    function isFiniteNumber(value) {
        return typeof value === 'number' && !isNaN(value) && isFinite(value);
    }

    /**
     * 値をJavaScriptのソース表記へ変換する
     * @param {*} value - 変換する値
     * @param {number} indentLevel - 字下げの深さ
     * @returns {string} ソース表記
     */
    function toCodeValue(value, indentLevel) {
        var indent = new Array(indentLevel + 1).join('    ');
        var nextIndent = new Array(indentLevel + 2).join('    ');
        var parts = [];
        var i, key;

        if (typeof value === 'string') {
            return '"' + value.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r/g, '\\r').replace(/\n/g, '\\n') + '"';
        }
        if (typeof value === 'number') {
            return isFiniteNumber(value) ? String(value) : '0';
        }
        if (typeof value === 'boolean') {
            return value ? 'true' : 'false';
        }
        if (value === null) {
            return 'null';
        }
        if (value instanceof Array) {
            for (i = 0; i < value.length; i++) parts.push(toCodeValue(value[i], indentLevel + 1));
            return '[' + parts.join(', ') + ']';
        }
        if (typeof value === 'object') {
            for (key in value) {
                if (value.hasOwnProperty(key)) {
                    parts.push(nextIndent + key + ': ' + toCodeValue(value[key], indentLevel + 1));
                }
            }
            return '{\n' + parts.join(',\n') + '\n' + indent + '}';
        }
        return 'null';
    }

    /**
     * プリセットを PRESETS へ貼り付けられるコードに変換する
     * @param {Object} preset - 対象プリセット
     * @returns {string} 組み込み用のコード
     */
    function presetToCode(preset) {
        return ',\n' + toCodeValue(preset, 0);
    }

    /**
     * プリセットの表示名を取り出す
     * @param {Object} preset - 対象プリセット
     * @returns {string} 表示名
     */
    function getPresetDisplayName(preset) {
        if (!preset) return '';
        if (typeof preset.name === 'string') return preset.name;
        if (preset.name && typeof preset.name === 'object') return getLabel(preset.name);
        return '';
    }

    /**
     * ドロップダウンの選択項目の文言を取り出す
     * @param {DropDownList} dropdown - 対象ドロップダウン
     * @returns {string} 選択中の文言（未選択なら空文字）
     */
    function getDropdownSelectionText(dropdown) {
        return (dropdown && dropdown.selection) ? dropdown.selection.text : '';
    }

    /**
     * プリセット一覧をドロップダウンへ流し込む
     * @param {DropDownList} dropdown - 対象ドロップダウン
     * @param {Array<Object>} presets - プリセットの配列
     * @returns {void}
     */
    function setDropdownItemsFromPresets(dropdown, presets) {
        dropdown.removeAll();
        for (var i = 0; i < presets.length; i++) {
            dropdown.add('item', getPresetDisplayName(presets[i]));
        }
        if (dropdown.items.length > 0) dropdown.selection = 0;
    }

    // =========================================
    // 設定の解決 / Settings resolution
    // =========================================

    /**
     * 設定をポイント単位の寸法へ変換する
     * @param {Object} settings - UIから読み取った設定
     * @returns {Object} ポイント換算済みの寸法
     */
    function toGeometry(settings) {
        /* ドットのミシン目に連動しているときは、左右も分割線と同じ値・同じ単位系を使う / Linked dot perforation shares the divider values and units */
        var linked = settings.lr.enabled && settings.lr.linkCenter && settings.center.mode === 'dot';
        var side = linked ? settings.center : settings.lr;

        return {
            linked: linked,
            sideWidthPt: toPt(side.width, linked ? strokePtFactor : rulerPtFactor),
            sideGapPt: toPt(side.gap, rulerPtFactor),
            sideInsetPt: toPt(Math.abs(Math.min(0, side.length)), rulerPtFactor),

            centerWidthPt: toPt(settings.center.width, strokePtFactor),
            centerGapPt: toPt(settings.center.gap, rulerPtFactor),
            centerInsetPt: toPt(Math.abs(Math.min(0, settings.center.length)), rulerPtFactor),
            offsetPt: toPt(settings.center.offset, rulerPtFactor),

            zigzagSizePt: toPt(settings.zigzag.size, rulerPtFactor),
            zigzagGapPt: toPt(settings.zigzag.gap || 0, rulerPtFactor),
            zigzagRepeat: Math.max(1, Math.round(settings.zigzag.repeat || 1)),

            cornerSizePt: toPt(settings.corner.size, rulerPtFactor),
            holeSizePt: toPt(settings.hole.size, rulerPtFactor),
            edgeSizePt: toPt(settings.edge.size, rulerPtFactor)
        };
    }

    /**
     * 寸法に数値以外が混じっていないかを確認する
     * @param {Object} geometry - ポイント換算済みの寸法
     * @returns {boolean} すべて数値なら true
     */
    function isValidGeometry(geometry) {
        var keys = [
            'sideWidthPt', 'sideGapPt', 'sideInsetPt',
            'centerWidthPt', 'centerGapPt', 'centerInsetPt', 'offsetPt',
            'zigzagSizePt', 'zigzagGapPt', 'zigzagRepeat',
            'cornerSizePt', 'holeSizePt', 'edgeSizePt'
        ];
        for (var i = 0; i < keys.length; i++) {
            if (isNaN(geometry[keys[i]])) return false;
        }
        return true;
    }

    // =========================================
    // 起動時チェック / Startup validation
    // =========================================

    /**
     * 実行できる選択状態かどうかを確認する
     * @returns {Object} { doc: Document, selectedItems: Array }（実行できないときは null）
     */
    function validateStartupState() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.openDocument));
            return null;
        }

        var doc = app.activeDocument;
        var selectedItems = getSafeSelection(doc);

        if (selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.selectRectangle));
            return null;
        }
        if (selectedItems.length > 1) {
            alert(getLabel(LABELS.alert.singleOnly));
            return null;
        }
        if (selectedItems[0].typename === 'GroupItem') {
            alert(getLabel(LABELS.alert.groupNotAllowed));
            return null;
        }
        if (!isRectanglePath(selectedItems[0])) {
            alert(getLabel(LABELS.alert.rectangleOnly));
            return null;
        }

        return { doc: doc, selectedItems: selectedItems };
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

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        var startup = validateStartupState();
        if (!startup) return;

        var doc = startup.doc;
        var selectedItems = startup.selectedItems;

        app.executeMenuCommand('edge');
        app.executeMenuCommand('AI Bounding Box Toggle');

        try {
            var originalLayer = selectedItems[0].layer;

            /* プレビュー専用レイヤーは必要になるまで作らない / The preview layer is created lazily */
            var previewLayer = null;

            /**
             * レイヤーがまだ生きているかを判定する
             * @param {Layer} layer - 対象レイヤー
             * @returns {boolean} 参照できれば true
             */
            function hasUsableLayer(layer) {
                if (!layer) return false;
                try {
                    return !!layer.name;
                } catch (e) {
                    return false;
                }
            }

            /**
             * プレビュー専用レイヤーを用意する
             * @returns {Layer} プレビューレイヤー
             */
            function ensurePreviewLayer() {
                if (hasUsableLayer(previewLayer)) return previewLayer;
                previewLayer = doc.layers.add();
                previewLayer.name = getLabel(LABELS.layerName.preview);
                return previewLayer;
            }

            /**
             * プレビューレイヤーの中身を空にする
             * @returns {void}
             */
            function clearPreviewLayerItems() {
                if (!hasUsableLayer(previewLayer)) return;
                try {
                    while (previewLayer.pageItems.length > 0) {
                        previewLayer.pageItems[0].remove();
                    }
                } catch (e) {
                    /* 途中で参照が切れても続行 / Continue even if the reference breaks */
                }
            }

            /**
             * プレビューを消して元オブジェクトを表示に戻す
             * @param {boolean} skipRedraw - 再描画を省略するか
             * @returns {void}
             */
            function removePreview(skipRedraw) {
                clearPreviewLayerItems();
                for (var i = 0; i < selectedItems.length; i++) {
                    selectedItems[i].hidden = false;
                }
                if (!skipRedraw) app.redraw();
            }

            /**
             * プレビューレイヤーそのものを削除する
             * @returns {void}
             */
            function removePreviewLayer() {
                if (!hasUsableLayer(previewLayer)) return;
                clearPreviewLayerItems();
                try {
                    previewLayer.remove();
                } catch (e) {
                    /* すでに削除済みでも続行 / Continue even if it is already removed */
                }
                previewLayer = null;
            }

            /**
             * UIの状態をプリセットと同じ構造の設定として読み取る
             * @param {string} presetName - 設定に付ける名前（省略時は選択中のプリセット名）
             * @returns {Object} 設定オブジェクト
             */
            function readSettingsFromUI(presetName) {
                return {
                    name: presetName || getDropdownSelectionText(ddPreset) || getLabel(LABELS.fallbackName.customPreset),
                    lr: {
                        enabled: chkSidePerforation.value,
                        linkCenter: chkLinkToCenter.value,
                        width: parseFloat(txtSideWidth.text),
                        gap: parseFloat(txtSideGap.text),
                        length: parseFloat(txtSideInset.text)
                    },
                    center: {
                        enabled: chkCenterSplit.value,
                        offset: parseFloat(txtCenterOffset.text),
                        mode: rdDividerDash.value ? 'dash' : 'dot',
                        width: parseFloat(txtDividerWidth.text),
                        gap: parseFloat(txtDividerGap.text),
                        length: parseFloat(txtDividerInset.text)
                    },
                    zigzag: {
                        mode: rdZigzagLeftRight.value ? 'lr' : (rdZigzagTopBottom.value ? 'tb' : 'none'),
                        size: parseFloat(txtZigzagSize.text),
                        gap: parseFloat(txtZigzagGap.text),
                        repeat: parseFloat(txtZigzagRepeat.text)
                    },
                    corner: {
                        mode: rdCornerRound.value ? 'round' : (rdCornerInverse.value ? 'inverse' : (rdCornerChamfer.value ? 'chamfer' : 'none')),
                        size: parseFloat(txtCornerSize.text)
                    },
                    hole: {
                        mode: rdHoleCircle.value ? 'circle' : (rdHoleTriangle.value ? 'triangle' : 'none'),
                        left: chkHoleLeft.value,
                        right: chkHoleRight.value,
                        size: parseFloat(txtHoleSize.text)
                    },
                    edge: {
                        mode: chkEdgeDoubleRound.value ? 'wround' : (rdEdgeCircle.value ? 'circle' : (rdEdgeTriangle.value ? 'triangle' : 'none')),
                        size: parseFloat(txtEdgeSize.text),
                        shapeOnly: chkEdgeOnly.value
                    },
                    preview: {
                        enabled: chkPreview.value,
                        expandAppearance: chkExpandAppearance.value
                    }
                };
            }

            /**
             * 設定をUIへ反映する
             * @param {Object} settings - 反映する設定
             * @returns {void}
             */
            function applySettingsToUI(settings) {
                if (!settings) return;

                chkSidePerforation.value = !!settings.lr.enabled;
                chkLinkToCenter.value = !!settings.lr.linkCenter;
                txtSideWidth.text = String(settings.lr.width);
                txtSideGap.text = String(settings.lr.gap);
                txtSideInset.text = String(settings.lr.length);

                chkCenterSplit.value = !!settings.center.enabled;
                txtCenterOffset.text = String(settings.center.offset);
                rdDividerDot.value = settings.center.mode !== 'dash';
                rdDividerDash.value = settings.center.mode === 'dash';
                txtDividerWidth.text = String(settings.center.width);
                txtDividerGap.text = String(settings.center.gap);
                txtDividerInset.text = String(settings.center.length);

                rdZigzagNone.value = settings.zigzag.mode === 'none';
                rdZigzagLeftRight.value = settings.zigzag.mode === 'lr';
                rdZigzagTopBottom.value = settings.zigzag.mode === 'tb';
                txtZigzagSize.text = String(settings.zigzag.size);
                txtZigzagGap.text = String(settings.zigzag.gap);
                txtZigzagRepeat.text = String(settings.zigzag.repeat);

                rdCornerNone.value = settings.corner.mode === 'none';
                rdCornerRound.value = settings.corner.mode === 'round';
                rdCornerInverse.value = settings.corner.mode === 'inverse';
                rdCornerChamfer.value = settings.corner.mode === 'chamfer';
                txtCornerSize.text = String(settings.corner.size);

                rdHoleNone.value = settings.hole.mode === 'none';
                rdHoleCircle.value = settings.hole.mode === 'circle';
                rdHoleTriangle.value = settings.hole.mode === 'triangle';
                chkHoleLeft.value = !!settings.hole.left;
                chkHoleRight.value = !!settings.hole.right;
                txtHoleSize.text = String(settings.hole.size);

                chkEdgeDoubleRound.value = settings.edge.mode === 'wround';
                /* ダブル角丸はエッジ形状と排他なので、ラジオは「なし」に寄せる / Double rounding excludes the edge shapes */
                rdEdgeNone.value = settings.edge.mode === 'none' || settings.edge.mode === 'wround';
                rdEdgeCircle.value = settings.edge.mode === 'circle';
                rdEdgeTriangle.value = settings.edge.mode === 'triangle';
                txtEdgeSize.text = String(settings.edge.size);
                chkEdgeOnly.value = !!settings.edge.shapeOnly;

                chkPreview.value = !!settings.preview.enabled;
                chkExpandAppearance.value = !!settings.preview.expandAppearance;

                syncOffsetSliderFromInput();

                updateHoleState();
                updateCornerState();
                updateCenterState();
                updateZigzagState();
                updateSideState();
                updateEdgeState();
            }

            /**
             * 現在の設定でチケット形状を生成する
             * @param {boolean} isPreview - true ならプレビューレイヤーへ複製して適用、false なら元オブジェクトへ直接適用
             * @returns {void}
             */
            function applyEffect(isPreview) {
                var settings = readSettingsFromUI();
                var geometry = toGeometry(settings);
                /* 入力途中で数値になっていないときは描かない（確定時は呼び出し側でアラート）/ Skip while a field is still being typed */
                if (!isValidGeometry(geometry)) return;

                var targets = [];
                var results = [];
                var i;

                for (i = 0; i < selectedItems.length; i++) {
                    if (isPreview) {
                        var targetLayer = ensurePreviewLayer();
                        doc.activeLayer = targetLayer;
                        var duplicated = selectedItems[i].duplicate(targetLayer, ElementPlacement.PLACEATEND);
                        normalizeInitialAppearance([duplicated], doc);
                        selectedItems[i].hidden = true;
                        targets.push(duplicated);
                    } else {
                        normalizeInitialAppearance([selectedItems[i]], doc);
                        targets.push(selectedItems[i]);
                    }
                }

                /* ダブル角丸のときは分割後の半片に角丸をかけるので、ここでは適用しない / Double rounding applies the radius to each half instead */
                var useDoubleRound = settings.center.enabled && settings.edge.mode === 'wround';
                if (!useDoubleRound && settings.corner.mode === 'round') {
                    applyRoundCornersEffect(targets, geometry.cornerSizePt);
                }

                for (i = 0; i < targets.length; i++) {
                    var rect = targets[i];
                    var bounds = {
                        left: rect.left,
                        top: rect.top,
                        width: rect.width,
                        height: rect.height,
                        centerLineX: rect.left + rect.width / 2 + geometry.offsetPt
                    };
                    var items = [rect];

                    if (useDoubleRound) {
                        applyDoubleRoundEdge(doc, rect, items, bounds, settings, geometry);
                    }
                    addCenterPerforation(doc, items, bounds, settings, geometry);
                    addSidePerforation(doc, items, bounds, settings, geometry);
                    addZigzag(doc, items, bounds, settings, geometry);
                    addCenterEdges(doc, items, bounds, settings, geometry);
                    addHoles(doc, items, bounds, settings, geometry);
                    addInverseCorners(doc, items, bounds, settings, geometry);
                    addChamferCorners(doc, items, bounds, settings, geometry);

                    finalizeSubtract(doc, items, results);
                }

                if (!isPreview) {
                    selectItems(doc, results.length > 0 ? results : targets);
                }
                app.redraw();
            }

            /**
             * プレビューを描き直す
             * @returns {void}
             */
            function updatePreview() {
                updateZigzagState();
                removePreview(chkPreview.value);
                if (!chkPreview.value) return;
                applyEffect(true);
            }

            // -----------------------------------------
            // ダイアログ / Dialog
            // -----------------------------------------

            var dialog = new Window('dialog', getLabel(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
            setupWindow(dialog);

            /* プリセット行 / Preset row */
            var presetRow = addRow(dialog);
            presetRow.alignChildren = ['fill', 'center'];
            presetRow.alignment = ['fill', 'top'];
            presetRow.margins = PRESET_ROW_MARGINS;
            presetRow.spacing = PRESET_ROW_SPACING;

            var ddPreset = presetRow.add('dropdownlist', undefined, []);
            ddPreset.helpTip = getLabel(LABELS.tooltip.preset);
            ddPreset.alignment = ['fill', 'center'];

            var btnSavePreset = presetRow.add('button', undefined, getLabel(LABELS.button.save));
            btnSavePreset.preferredSize.width = SAVE_BUTTON_SIZE[0];
            btnSavePreset.preferredSize.height = SAVE_BUTTON_SIZE[1];
            btnSavePreset.alignment = ['right', 'center'];

            /* 1行目：コーナー／ギザギザ ＋ 左右 / Row 1: corner & zigzag, then the sides */
            var topRow = addRow(dialog);
            topRow.alignChildren = ['fill', 'top'];

            var leftColumn = addColumn(topRow);

            var panelCorner = addPanel(leftColumn, getLabel(LABELS.panel.corner), ['left', 'top']);
            var rdCornerNone = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.none));
            rdCornerNone.helpTip = getLabel(LABELS.tooltip.cornerNone);
            var rdCornerRound = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.round));
            rdCornerRound.helpTip = getLabel(LABELS.tooltip.cornerRound);
            var rdCornerInverse = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.inverse));
            rdCornerInverse.helpTip = getLabel(LABELS.tooltip.cornerInverse);
            var rdCornerChamfer = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.chamfer));
            rdCornerChamfer.helpTip = getLabel(LABELS.tooltip.cornerChamfer);
            rdCornerNone.value = true;

            var cornerSizeField = addNumberField(panelCorner, {
                label: labelText(LABELS.fieldLabel.size),
                value: '5',
                unit: rulerUnitLabel,
                onChange: function () { updatePreview(); }
            });
            var txtCornerSize = cornerSizeField.input;

            var panelZigzag = addPanel(leftColumn, getLabel(LABELS.panel.zigzag));

            var zigzagModeRow = addRow(panelZigzag);
            var rdZigzagNone = zigzagModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.none));
            rdZigzagNone.helpTip = getLabel(LABELS.tooltip.zigzagNone);
            var rdZigzagLeftRight = zigzagModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.leftRight));
            rdZigzagLeftRight.helpTip = getLabel(LABELS.tooltip.zigzagLeftRight);
            var rdZigzagTopBottom = zigzagModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.topBottom));
            rdZigzagTopBottom.helpTip = getLabel(LABELS.tooltip.zigzagTopBottom);
            rdZigzagNone.value = true;

            var zigzagLabelWidth = ZIGZAG_LABEL_WIDTH[uiLang];

            var zigzagSizeField = addNumberField(panelZigzag, {
                label: labelText(LABELS.fieldLabel.zigzagSize),
                value: '10',
                unit: rulerUnitLabel,
                labelWidth: zigzagLabelWidth,
                onChange: function () { updatePreview(); }
            });
            var txtZigzagSize = zigzagSizeField.input;

            var zigzagRepeatField = addNumberField(panelZigzag, {
                label: labelText(LABELS.fieldLabel.zigzagRepeat),
                value: '3',
                labelWidth: zigzagLabelWidth,
                stepOptions: { integer: true, min: 1 },
                onChange: function () { updateZigzagState(); updatePreview(); }
            });
            var txtZigzagRepeat = zigzagRepeatField.input;

            var zigzagGapField = addNumberField(panelZigzag, {
                label: labelText(LABELS.fieldLabel.gap),
                value: '0',
                unit: rulerUnitLabel,
                labelWidth: zigzagLabelWidth,
                stepOptions: {},
                onChange: function () { updatePreview(); }
            });
            var txtZigzagGap = zigzagGapField.input;

            /* 左右（ミシン目＋スリット／ホール）/ Sides: perforation and slit / hole */
            var panelSides = addPanel(topRow, getLabel(LABELS.panel.sides));

            var panelSidePerforation = addPanel(panelSides, getLabel(LABELS.panel.sidePerforation));

            var sideEnableRow = addRow(panelSidePerforation);
            var chkSidePerforation = sideEnableRow.add('checkbox', undefined, getLabel(LABELS.checkbox.enable));
            chkSidePerforation.helpTip = getLabel(LABELS.tooltip.sidePerforation);
            var chkLinkToCenter = sideEnableRow.add('checkbox', undefined, getLabel(LABELS.checkbox.linkToCenter));
            chkLinkToCenter.helpTip = getLabel(LABELS.tooltip.linkToCenter);
            chkLinkToCenter.value = true;

            var sideWidthField = addNumberField(panelSidePerforation, {
                label: labelText(LABELS.fieldLabel.lineWidth),
                value: '3',
                unit: rulerUnitLabel,
                onChange: function () { updatePreview(); }
            });
            var txtSideWidth = sideWidthField.input;

            var sideGapField = addNumberField(panelSidePerforation, {
                label: labelText(LABELS.fieldLabel.gap),
                value: '6',
                unit: rulerUnitLabel,
                onChange: function () { updatePreview(); }
            });
            var txtSideGap = sideGapField.input;

            var sideInsetField = addNumberField(panelSidePerforation, {
                label: labelText(LABELS.fieldLabel.inset),
                value: '0',
                unit: rulerUnitLabel,
                characters: INSET_FIELD_CHARS,
                stepOptions: { max: 0 },
                onChange: function () { updatePreview(); }
            });
            var txtSideInset = sideInsetField.input;

            var panelHole = addPanel(panelSides, getLabel(LABELS.panel.hole), ['left', 'top']);

            var holeModeRow = addRow(panelHole);
            var rdHoleNone = holeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.none));
            rdHoleNone.helpTip = getLabel(LABELS.tooltip.holeNone);
            var rdHoleCircle = holeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.circle));
            rdHoleCircle.helpTip = getLabel(LABELS.tooltip.holeCircle);
            var rdHoleTriangle = holeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.triangle));
            rdHoleTriangle.helpTip = getLabel(LABELS.tooltip.holeTriangle);
            rdHoleNone.value = true;

            var holeSideRow = addRow(panelHole);
            var chkHoleLeft = holeSideRow.add('checkbox', undefined, getLabel(LABELS.checkbox.left));
            chkHoleLeft.helpTip = getLabel(LABELS.tooltip.holeSide);
            var chkHoleRight = holeSideRow.add('checkbox', undefined, getLabel(LABELS.checkbox.right));
            chkHoleRight.helpTip = getLabel(LABELS.tooltip.holeSide);
            chkHoleLeft.value = true;
            chkHoleRight.value = true;

            var holeSizeField = addNumberField(panelHole, {
                label: labelText(LABELS.fieldLabel.size),
                value: '10',
                unit: rulerUnitLabel,
                onChange: function () { updatePreview(); }
            });
            var txtHoleSize = holeSizeField.input;

            /* 左右分割（分割線＋エッジ）/ Center split: divider line and edges */
            var panelCenter = addPanel(dialog, getLabel(LABELS.panel.centerSplit));

            var centerTopRow = addRow(panelCenter);
            centerTopRow.alignChildren = ['center', 'center'];

            var chkCenterSplit = centerTopRow.add('checkbox', undefined, getLabel(LABELS.checkbox.enable));
            chkCenterSplit.helpTip = getLabel(LABELS.tooltip.centerSplit);
            chkCenterSplit.value = true;
            /* 複数選択時は最小幅の半分を上限にして、すべての対象で安全な範囲に制限 / Clamp to the narrowest item so every target stays valid */
            var minHalfWidthPt = null;
            for (var itemIndex = 0; itemIndex < selectedItems.length; itemIndex++) {
                var halfWidthPt = Math.abs(selectedItems[itemIndex].width) / 2;
                if (minHalfWidthPt === null || halfWidthPt < minHalfWidthPt) minHalfWidthPt = halfWidthPt;
            }
            var maxOffset = Math.round(fromPt(minHalfWidthPt || 0, rulerPtFactor) * 10) / 10;

            var txtCenterOffset = addSteppedInput(centerTopRow, '0', OFFSET_FIELD_CHARS, { min: -maxOffset, max: maxOffset },
                function () { syncOffsetSliderFromInput(); updatePreview(); });
            txtCenterOffset.helpTip = getLabel(LABELS.tooltip.centerOffset);
            var lblCenterOffsetUnit = centerTopRow.add('statictext', undefined, rulerUnitLabel);

            var sliderCenterOffset = centerTopRow.add('slider', undefined, 0, -maxOffset, maxOffset);
            sliderCenterOffset.helpTip = getLabel(LABELS.tooltip.centerOffset);
            sliderCenterOffset.preferredSize.width = OFFSET_SLIDER_WIDTH;

            var centerBodyRow = addRow(panelCenter);
            centerBodyRow.alignChildren = ['fill', 'top'];

            var panelDivider = addPanel(centerBodyRow, getLabel(LABELS.panel.divider));

            var dividerModeRow = addRow(panelDivider);
            var rdDividerDot = dividerModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.dot));
            rdDividerDot.helpTip = getLabel(LABELS.tooltip.dividerDot);
            var rdDividerDash = dividerModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.dash));
            rdDividerDash.helpTip = getLabel(LABELS.tooltip.dividerDash);
            rdDividerDot.value = true;

            var dividerWidthField = addNumberField(panelDivider, {
                label: labelText(LABELS.fieldLabel.lineWidth),
                value: '3',
                unit: strokeUnitLabel,
                onChange: function () { syncSideFromCenter(); updatePreview(); }
            });
            var txtDividerWidth = dividerWidthField.input;

            var dividerGapField = addNumberField(panelDivider, {
                label: labelText(LABELS.fieldLabel.gap),
                value: '6',
                unit: rulerUnitLabel,
                onChange: function () { syncSideFromCenter(); updatePreview(); }
            });
            var txtDividerGap = dividerGapField.input;

            var dividerInsetField = addNumberField(panelDivider, {
                label: labelText(LABELS.fieldLabel.inset),
                value: '0',
                unit: rulerUnitLabel,
                characters: INSET_FIELD_CHARS,
                stepOptions: { max: 0 },
                onChange: function () { syncSideFromCenter(); updatePreview(); }
            });
            var txtDividerInset = dividerInsetField.input;

            var panelEdge = addPanel(centerBodyRow, getLabel(LABELS.panel.edge), ['left', 'top']);

            var edgeModeRow = addRow(panelEdge);
            var rdEdgeNone = edgeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.none));
            rdEdgeNone.helpTip = getLabel(LABELS.tooltip.edgeNone);
            var rdEdgeCircle = edgeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.circle));
            rdEdgeCircle.helpTip = getLabel(LABELS.tooltip.edgeCircle);
            var rdEdgeTriangle = edgeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.triangle));
            rdEdgeTriangle.helpTip = getLabel(LABELS.tooltip.edgeTriangle);
            rdEdgeNone.value = true;

            var chkEdgeDoubleRound = panelEdge.add('checkbox', undefined, getLabel(LABELS.checkbox.doubleRound));
            chkEdgeDoubleRound.helpTip = getLabel(LABELS.tooltip.edgeDoubleRound);

            var edgeSizeField = addNumberField(panelEdge, {
                label: labelText(LABELS.fieldLabel.size),
                value: '10',
                unit: rulerUnitLabel,
                onChange: function () { syncDividerInsetFromEdgeSize(); updatePreview(); }
            });
            var txtEdgeSize = edgeSizeField.input;

            var chkEdgeOnly = panelEdge.add('checkbox', undefined, getLabel(LABELS.checkbox.edgesOnly));
            chkEdgeOnly.helpTip = getLabel(LABELS.tooltip.edgesOnly);

            /* プレビュー行 / Preview row */
            var previewRow = addRow(dialog);
            previewRow.alignment = ['center', 'top'];

            var chkPreview = previewRow.add('checkbox', undefined, getLabel(LABELS.checkbox.preview));
            chkPreview.helpTip = getLabel(LABELS.tooltip.preview);
            chkPreview.value = true;

            var chkExpandAppearance = previewRow.add('checkbox', undefined, getLabel(LABELS.checkbox.expandAppearance));
            chkExpandAppearance.helpTip = getLabel(LABELS.tooltip.expandAppearance);
            chkExpandAppearance.value = false;

            /* ボタン行 / Button row */
            var buttonRow = addButtonRow(dialog);

            var isOutlineMode = false;
            var btnOutlineToggle = buttonRow.leftGroup.add('button', undefined, getLabel(LABELS.button.outlineOn));

            var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel(LABELS.button.cancel), { name: 'cancel' });
            var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel(LABELS.button.ok), { name: 'ok' });

            // -----------------------------------------
            // ディム制御と連動 / Enable state and syncing
            // -----------------------------------------

            /**
             * 入力欄の値からスライダー位置を合わせる
             * @returns {void}
             */
            function syncOffsetSliderFromInput() {
                var offsetValue = parseFloat(txtCenterOffset.text);
                if (isNaN(offsetValue)) return;
                sliderCenterOffset.value = Math.max(-maxOffset, Math.min(maxOffset, offsetValue));
            }

            /**
             * 連動が有効なとき、左右のミシン目に分割線の値をコピーする
             * @returns {void}
             */
            function syncSideFromCenter() {
                if (!chkLinkToCenter.value || !rdDividerDot.value) return;
                txtSideWidth.text = txtDividerWidth.text;
                txtSideGap.text = txtDividerGap.text;
                txtSideInset.text = txtDividerInset.text;
            }

            /**
             * エッジのサイズに合わせて分割線の長さ（食い込み量）を決める
             * @returns {void}
             */
            function syncDividerInsetFromEdgeSize() {
                if (!chkCenterSplit.value) return;
                if (rdEdgeNone.value && !chkEdgeDoubleRound.value) return;

                var edgeSizeValue = parseFloat(txtEdgeSize.text);
                if (isNaN(edgeSizeValue)) return;

                var ratio = 1.2;
                if (rdEdgeCircle.value) ratio = 1.0;
                else if (chkEdgeDoubleRound.value) ratio = 2.0;

                txtDividerInset.text = String(-(Math.round(edgeSizeValue * ratio * 1000) / 1000));
            }

            /**
             * 左右のミシン目まわりのディムを更新する
             * @returns {void}
             */
            function updateSideState() {
                var isOn = chkSidePerforation.value;
                var isLinked = isOn && chkLinkToCenter.value && rdDividerDot.value;
                chkLinkToCenter.enabled = isOn;
                sideWidthField.row.enabled = isOn && !isLinked;
                sideGapField.row.enabled = isOn && !isLinked;
                sideInsetField.row.enabled = isOn && !isLinked;
                redrawSteppersIn(panelSidePerforation);
                if (isLinked) syncSideFromCenter();
            }

            /**
             * ギザギザまわりのディムを更新する
             * @returns {void}
             */
            function updateZigzagState() {
                var isOn = !rdZigzagNone.value;
                zigzagSizeField.row.enabled = isOn;
                zigzagRepeatField.row.enabled = isOn;
                zigzagGapField.row.enabled = isOn && Math.round(parseFloat(txtZigzagRepeat.text) || 0) > 1;
                redrawSteppersIn(panelZigzag);
            }

            /**
             * スリット／ホールまわりのディムを更新する
             * @returns {void}
             */
            function updateHoleState() {
                var isOn = !rdHoleNone.value;
                holeSideRow.enabled = isOn;
                holeSizeField.row.enabled = isOn;
                redrawSteppersIn(holeSizeField.row);
            }

            /**
             * コーナーまわりのディムを更新する
             * @returns {void}
             */
            function updateCornerState() {
                var isDoubleRound = chkEdgeDoubleRound.value;
                panelCorner.enabled = !isDoubleRound;
                cornerSizeField.row.enabled = !isDoubleRound && !rdCornerNone.value;
                redrawSteppersIn(cornerSizeField.row);
            }

            /**
             * 左右分割まわりのディムを更新する
             * @returns {void}
             */
            function updateCenterState() {
                var isOn = chkCenterSplit.value;
                txtCenterOffset.parent.enabled = isOn; /* ∧∨と入力欄をまとめて / the stepper and the field together */
                lblCenterOffsetUnit.enabled = isOn;
                sliderCenterOffset.enabled = isOn;
                centerBodyRow.enabled = isOn;
                redrawSteppersIn(panelCenter);
            }

            /**
             * エッジまわりのディムを更新する
             * @returns {void}
             */
            function updateEdgeState() {
                var isDoubleRound = chkCenterSplit.value && chkEdgeDoubleRound.value;
                var isOn = (chkCenterSplit.value && !rdEdgeNone.value) || isDoubleRound;
                edgeModeRow.enabled = !isDoubleRound;
                edgeSizeField.row.enabled = isOn;
                chkEdgeOnly.enabled = isOn;
                panelDivider.enabled = chkCenterSplit.value && !(isOn && chkEdgeOnly.value);
                redrawSteppersIn(centerBodyRow);
                if (isOn) syncDividerInsetFromEdgeSize();
            }

            // -----------------------------------------
            // イベント / Events
            // -----------------------------------------

            ddPreset.onChange = function () {
                var index = ddPreset.selection ? ddPreset.selection.index : -1;
                if (index < 0 || index >= PRESETS.length) return;
                applySettingsToUI(cloneSimpleValue(PRESETS[index]));
                updatePreview();
            };

            btnSavePreset.onClick = function () {
                var baseName = getDropdownSelectionText(ddPreset) || getLabel(LABELS.fallbackName.customPreset);
                var newName = prompt(getLabel(LABELS.alert.presetName), baseName);
                if (newName === null) return;
                newName = String(newName).replace(/^\s+|\s+$/g, '');
                if (!newName) return;

                var preset = readSettingsFromUI(newName);
                if (!isValidGeometry(toGeometry(preset))) {
                    alert(getLabel(LABELS.alert.enterValidNumbers));
                    return;
                }

                PRESETS.push(cloneSimpleValue(preset));
                setDropdownItemsFromPresets(ddPreset, PRESETS);
                ddPreset.selection = ddPreset.items.length - 1;

                alert(getLabel(LABELS.alert.presetSaved) + presetToCode(preset));
            };

            chkSidePerforation.onClick = function () {
                /* 左右のミシン目とギザギザ（左右）は同じ辺を使うので併用しない / Side perforation and L/R zigzag share the same edges */
                if (chkSidePerforation.value && rdZigzagLeftRight.value) {
                    rdZigzagNone.value = true;
                    rdZigzagLeftRight.value = false;
                    updateZigzagState();
                }
                updateSideState();
                updatePreview();
            };
            chkLinkToCenter.onClick = function () { updateSideState(); updatePreview(); };

            rdZigzagNone.onClick = function () { updateZigzagState(); updatePreview(); };
            rdZigzagLeftRight.onClick = function () { updateZigzagState(); updatePreview(); };
            rdZigzagTopBottom.onClick = function () { updateZigzagState(); updatePreview(); };

            rdHoleNone.onClick = function () { updateHoleState(); updatePreview(); };
            rdHoleCircle.onClick = function () { updateHoleState(); updatePreview(); };
            rdHoleTriangle.onClick = function () { updateHoleState(); updatePreview(); };
            chkHoleLeft.onClick = function () { updatePreview(); };
            chkHoleRight.onClick = function () { updatePreview(); };

            rdCornerNone.onClick = function () { updateCornerState(); updatePreview(); };
            rdCornerRound.onClick = function () { updateCornerState(); updatePreview(); };
            rdCornerInverse.onClick = function () { updateCornerState(); updatePreview(); };
            rdCornerChamfer.onClick = function () { updateCornerState(); updatePreview(); };

            chkCenterSplit.onClick = function () { updateCenterState(); updatePreview(); };
            rdDividerDot.onClick = function () { updateSideState(); updatePreview(); };
            rdDividerDash.onClick = function () { updateSideState(); updatePreview(); };

            rdEdgeNone.onClick = function () { updateEdgeState(); updatePreview(); };
            rdEdgeCircle.onClick = function () { updateEdgeState(); updatePreview(); };
            rdEdgeTriangle.onClick = function () { updateEdgeState(); updatePreview(); };
            chkEdgeOnly.onClick = function () { updateEdgeState(); updatePreview(); };

            chkEdgeDoubleRound.onClick = function () {
                if (chkEdgeDoubleRound.value) {
                    /* ダブル角丸はコーナー処理・エッジ形状と併用しない / Double rounding replaces the corner and edge shapes */
                    rdCornerNone.value = true;
                    rdCornerRound.value = false;
                    rdCornerInverse.value = false;
                    rdCornerChamfer.value = false;
                    rdEdgeNone.value = true;
                    rdEdgeCircle.value = false;
                    rdEdgeTriangle.value = false;
                    txtEdgeSize.text = '5';
                } else if (rdCornerNone.value) {
                    /* OFFに戻したときは、強制的に「なし」にしていたコーナーを角丸へ戻す / Restore the corner mode forced to "none" */
                    rdCornerNone.value = false;
                    rdCornerRound.value = true;
                }
                updateCornerState();
                syncDividerInsetFromEdgeSize();
                updateEdgeState();
                updatePreview();
            };

            chkPreview.onClick = function () { updatePreview(); };

            txtCenterOffset.onChanging = function () {
                syncOffsetSliderFromInput();
                updatePreview();
            };
            sliderCenterOffset.onChanging = function () {
                txtCenterOffset.text = String(Math.round(sliderCenterOffset.value * 10) / 10);
                updatePreview();
            };

            txtSideWidth.onChanging = updatePreview;
            txtSideGap.onChanging = updatePreview;
            txtSideInset.onChanging = function () { clampToNegative(txtSideInset); updatePreview(); };
            txtZigzagSize.onChanging = updatePreview;
            txtZigzagGap.onChanging = updatePreview;
            txtZigzagRepeat.onChanging = function () { updateZigzagState(); updatePreview(); };
            txtCornerSize.onChanging = updatePreview;
            txtHoleSize.onChanging = updatePreview;
            txtDividerWidth.onChanging = function () { syncSideFromCenter(); updatePreview(); };
            txtDividerGap.onChanging = function () { syncSideFromCenter(); updatePreview(); };
            txtDividerInset.onChanging = function () {
                clampToNegative(txtDividerInset);
                syncSideFromCenter();
                updatePreview();
            };
            txtEdgeSize.onChanging = function () { syncDividerInsetFromEdgeSize(); updatePreview(); };

            btnOutlineToggle.onClick = function () {
                try {
                    app.executeMenuCommand('preview');
                    isOutlineMode = !isOutlineMode;
                    btnOutlineToggle.text = isOutlineMode ? getLabel(LABELS.button.outlineOff) : getLabel(LABELS.button.outlineOn);
                } catch (e) {
                    /* 表示モードを切り替えられなくても続行 / Continue even if the view cannot be toggled */
                }
            };

            /**
             * 入力欄の値が正なら0に丸める（食い込み量は0以下のみ）
             * @param {EditText} input - 対象の入力欄
             * @returns {void}
             */
            function clampToNegative(input) {
                var value = parseFloat(input.text);
                if (!isNaN(value) && value > 0) input.text = '0';
            }

            // -----------------------------------------
            // 表示と確定 / Show and apply
            // -----------------------------------------

            setDropdownItemsFromPresets(ddPreset, PRESETS);
            if (PRESETS.length > 0) {
                applySettingsToUI(cloneSimpleValue(PRESETS[0]));
            }
            syncDividerInsetFromEdgeSize();
            updatePreview();

            alignRightOnlyButtonRow(buttonRow);
            prepareDialogWindow(dialog, SCRIPT_NAME);
            var isConfirmed = (dialog.show() === 1);
            removePreview();
            removePreviewLayer();

            if (!isConfirmed) return;

            if (!isValidGeometry(toGeometry(readSettingsFromUI()))) {
                alert(getLabel(LABELS.alert.enterValidNumbers));
                return;
            }

            doc.activeLayer = originalLayer;
            applyEffect(false);
            if (chkExpandAppearance.value) {
                app.executeMenuCommand('expandStyle');
            }
        } finally {
            unloadStrokeDotAction();
            app.executeMenuCommand('edge');
            app.executeMenuCommand('AI Bounding Box Toggle');
        }
    }

    main();

})();
