#target illustrator
#targetengine "SymbolListBuilderEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメントに登録されたシンボルを一覧表示する専用アートボード「シンボル一覧」を自動生成します。
ダイアログでパラメーターを操作しながらライブプレビューで確認でき、［OK］で確定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolListBuilder.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ncac687d0a3a0

### Overview

Generates a dedicated "Symbol List" artboard that lays out every symbol registered in the document.
Parameters are adjusted with a live preview in the dialog, and OK commits the result.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolListBuilder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SymbolListBuilder";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolListBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolListBuilder.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ncac687d0a3a0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /**
     * @typedef {object} ListSettings
     * @property {string} position - 作成方向（"right" / "below"）
     * @property {string} baseMode - 基準（"last" = 最終アートボード / "specified" = 番号指定）
     * @property {number} baseArtboardNumber - 基準にするアートボード番号（1 始まり）
     * @property {number} artboardGap - 基準アートボードとの間隔（pt）
     * @property {number} margin - アートボード内側の余白（pt）
     * @property {boolean} update - 既存のシンボル一覧を削除して作り直すか
     * @property {boolean} showCaption - シンボル名を表示するか
     * @property {string} filter - 収集対象（"all" / "used"）
     * @property {number} symbolGap - シンボル同士の間隔（pt）
     * @property {number} maxRowWidth - 1 行の最大幅（pt）
     * @property {string} bgColor - 背景（BACKGROUND_CHOICES のいずれか）
     * @property {string} captionPosition - キャプションの位置（"above" / "below"）
     * @property {number} fontSize - キャプションのフォントサイズ（pt）
     * @property {?number} [widthOverridePt] - 固定する幅（pt、null なら自動）
     * @property {?number} [heightOverridePt] - 固定する高さ（pt、null なら自動）
     */

    /* 設定の既定値（寸法は pt）。［更新］と［収集対象］は保存値を使わず、起動時は常に ON／すべて
     * Default settings (sizes in pt). Update and Collect ignore saved values and always start as on / all */
    var DEFAULT_SETTINGS = {
        position: "below",
        baseMode: "last",
        baseArtboardNumber: 1,
        artboardGap: 100,
        margin: 50,
        update: true,
        showCaption: false,
        filter: "all",
        symbolGap: 20,
        maxRowWidth: 800,
        bgColor: "none",
        captionPosition: "below",
        fontSize: 9
    };

    /* シンボルとキャプションの間隔（pt）/ Gap between a symbol and its caption (pt) */
    var CAPTION_GAP_PT = 6;

    /* キャプションのフォント（UI 言語別。見つからなければ既定のまま）/ Caption font per UI language (left as is if missing) */
    var CAPTION_FONT_NAMES = { ja: "HiraginoSans-W3", en: "MyriadPro-Regular" };

    /* 起動直後のプレビューのズーム倍率 / Zoom factor for the initial preview */
    var INITIAL_PREVIEW_ZOOM = 0.6;

    /* ウィンドウに合わせた後に掛けるズーム倍率 / Zoom ratio applied after fitting to the window */
    var FIT_ZOOM_RATIO = 0.9;

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

    var ROW_SPACING             = 6;                  /* 行内の要素間隔 / spacing within a row */
    var RADIO_SPACING           = 10;                 /* 横並びラジオの間隔 / spacing between radios */
    var FIELD_LABEL_WIDTH       = 70;                 /* 項目名の幅 / field label width */
    var NUMBER_INPUT_CHARACTERS = 4;                  /* 数値入力欄の文字数 / numeric field width */
    var BACKGROUND_RADIO_WIDTH  = 60;                 /* 背景ラジオの幅（2×2 の列揃え）/ background radio width */

    /**
     * タイトル付きのパネルを追加し、共通レイアウトを設定する
     * @param {object} parent - 追加先のパネルまたはグループ
     * @param {object} titleSet - パネル名のラベル（ja/en）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleSet) {
        var newPanel = parent.add("panel", undefined, getLabel(titleSet));
        setupPanel(newPanel, 6);
        return newPanel;
    }

    /**
     * 2カラムの列にする縦並びのグループを追加する
     * @param {Group} parent - 追加先のグループ
     * @returns {Group} 追加した列グループ
     */
    function addColumnGroup(parent) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "fill";
        return columnGroup;
    }

    /**
     * 横並びの行グループを追加する（子は左詰め・天地中央）
     * @param {object} parent - 追加先のパネルまたはグループ
     * @param {number} [spacing] - 要素間隔（省略時は ROW_SPACING）
     * @returns {Group} 追加した行グループ
     */
    function addRowGroup(parent, spacing) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
        return rowGroup;
    }

    /**
     * 行の先頭に右揃えの項目名を追加する
     * @param {Group} rowGroup - 追加先の行グループ
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @returns {StaticText} 追加した項目名
     */
    function addFieldLabel(rowGroup, labelSet) {
        var fieldLabel = rowGroup.add("statictext", undefined, labelText(labelSet));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return fieldLabel;
    }

    /**
     * ラベルとツールチップ付きのラジオボタン・チェックボックス・ボタンを追加する
     * @param {object} parent - 追加先のパネルまたはグループ
     * @param {string} controlType - "radiobutton" / "checkbox" / "button"
     * @param {object} labelSet - 表示するラベル（ja/en）
     * @param {object} tooltipSet - ツールチップ（ja/en）
     * @returns {object} 追加したコントロール
     */
    function addLabeledControl(parent, controlType, labelSet, tooltipSet) {
        var control = parent.add(controlType, undefined, getLabel(labelSet));
        control.helpTip = getLabel(tooltipSet);
        return control;
    }

    /**
     * ツールチップ付きの数値入力欄を、左に∧∨を付けて追加する（↑↓キーも∧∨と同じ処理で増減する）
     * 増減後の処理と下限は、あとで bindNumberInput() が .stepOptions に入れる
     * @param {Group} parent - 追加先の行グループ
     * @param {string} initialText - 初期表示する値
     * @param {object} tooltipSet - ツールチップ（ja/en）
     * @param {object} [stepOptions] - ∧∨の step / min / integer（省略時は 0 以上の小数）
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup、増減の設定は .stepOptions で参照できる）
     */
    function addNumberInput(parent, initialText, tooltipSet, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = parent.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var inputStepOptions = stepOptions || { step: 1, min: 0 };
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, inputStepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = NUMBER_INPUT_CHARACTERS;
        numberInput.helpTip = getLabel(tooltipSet);
        numberInput.stepperGroup = stepperGroup;
        numberInput.stepOptions = inputStepOptions;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 親の異なるラジオボタンを手動で排他選択にする
     * @param {RadioButton[]} radios - 排他にするラジオボタン
     * @param {RadioButton} selectedRadio - 選択するラジオボタン
     * @returns {void}
     */
    function selectExclusiveRadio(radios, selectedRadio) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (radios[i] === selectedRadio);
        }
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

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 以前の設定保存用キー（Illustrator環境設定。読み継ぎ用）/ Former preference key for saved settings (migration) */
    var LEGACY_PREF_KEY = "swwwitch.listupallsymbol.settings";

    /* 保存する設定のキー（保存文字列の並び順）/ Keys of saved settings (serialized order) */
    var SETTINGS_KEYS = [
        "position", "baseMode", "baseArtboardNumber", "artboardGap", "margin",
        "update", "showCaption", "filter", "symbolGap", "maxRowWidth",
        "bgColor", "captionPosition", "fontSize"
    ];

    /* 背景の選択肢（ラジオの並び順）と、塗りに使う墨の濃度（%）/ Background choices (radio order) and black ink percentage */
    var BACKGROUND_CHOICES = ["none", "black", "white", "gray"];
    var BACKGROUND_BLACK_PERCENT = { black: 100, white: 0, gray: 50 };

    /* 選択肢で持つ設定と、取りうる値 / Settings stored as one of fixed choices */
    var SETTING_CHOICES = {
        position: ["right", "below"],
        baseMode: ["last", "specified"],
        filter: ["all", "used"],
        bgColor: BACKGROUND_CHOICES,
        captionPosition: ["above", "below"]
    };

    /* 旧アートボード上のアイテムを拾う許容幅（pt）/ Tolerance for picking items on an old artboard (pt) */
    var ARTBOARD_HIT_TOLERANCE_PT = 1.0;

    /* プレビュー用アートボードの矩形を照合する許容誤差（pt）/ Tolerance for matching the preview artboard rect (pt) */
    var RECT_MATCH_TOLERANCE_PT = 0.01;

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "シンボル一覧を作成", en: "Create Symbol List" }
        },
        panel: {
            artboard: { ja: "作成するアートボード", en: "New artboard" },
            location: { ja: "作成位置", en: "Location" },
            sizeAndPadding: { ja: "サイズと余白", en: "Size & padding" },
            background: { ja: "背景", en: "Background" },
            symbols: { ja: "収集するシンボル", en: "Symbols" },
            collectTarget: { ja: "収集対象", en: "Collect" },
            arrangement: { ja: "並べ方", en: "Layout" },
            caption: { ja: "キャプション", en: "Caption" }
        },
        radio: {
            baseLast: { ja: "最終アートボード", en: "Last artboard" },
            baseSpecified: { ja: "指定", en: "Specified" },
            directionRight: { ja: "右側", en: "Right" },
            directionBelow: { ja: "下側", en: "Below" },
            background: {
                none: { ja: "なし", en: "None" },
                black: { ja: "黒", en: "Black" },
                white: { ja: "白", en: "White" },
                gray: { ja: "グレー", en: "Gray" }
            },
            captionAbove: { ja: "上", en: "Top" },
            captionBelow: { ja: "下", en: "Bottom" },
            filterAll: { ja: "すべて", en: "All" },
            filterUsed: { ja: "使用中のみ", en: "Used only" }
        },
        checkbox: {
            update: { ja: "更新", en: "Update" },
            showCaption: { ja: "シンボル名を表示", en: "Show symbol name" }
        },
        fieldLabel: {
            direction: { ja: "方向", en: "Direction" },
            gap: { ja: "間隔", en: "Gap" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            margin: { ja: "余白", en: "Padding" },
            maxRowWidth: { ja: "最大幅", en: "Max row width" },
            captionPosition: { ja: "位置", en: "Position" },
            fontSize: { ja: "フォントサイズ", en: "Font size" }
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            baseLast: {
                ja: "アートボード一覧の末尾を基準にする。更新 ON の場合は既存のシンボル一覧を除外",
                en: "Use the last artboard in the list as the base; with Update on, the existing symbol list is excluded"
            },
            baseSpecified: { ja: "指定した番号のアートボードを基準にする（1 始まり）", en: "Use the artboard with the given number as the base (1-based)" },
            baseArtboardNumber: { ja: "基準にするアートボード番号（1 始まり）", en: "Base artboard number (1-based)" },
            directionRight: { ja: "基準アートボードの右に新規アートボードを作成", en: "Place the new artboard to the right of the base" },
            directionBelow: { ja: "基準アートボードの下に新規アートボードを作成", en: "Place the new artboard below the base" },
            artboardGap: { ja: "基準アートボードと新規アートボードの間隔", en: "Distance between the base and new artboards" },
            margin: { ja: "新規アートボードの内側に設ける余白", en: "Padding inside the new artboard around the symbols" },
            symbolGap: { ja: "シンボル同士の間隔", en: "Spacing between adjacent symbols" },
            maxRowWidth: { ja: "1行に並べる最大幅。超えると次の行に折り返し", en: "Max width per row; symbols wrap to the next row when exceeded" },
            width: {
                ja: "アートボードの幅。空欄または 0 で自動計算、入力すると指定値で固定",
                en: "Artboard width. Empty/0 = auto; a number forces that exact size"
            },
            height: {
                ja: "アートボードの高さ。空欄または 0 で自動計算、入力すると指定値で固定",
                en: "Artboard height. Empty/0 = auto; a number forces that exact size"
            },
            update: {
                ja: "既存の「シンボル一覧」アートボードと対象オブジェクトを削除して作り直す",
                en: "Delete the existing Symbol List artboard and its objects, then rebuild"
            },
            showCaption: { ja: "各シンボルの近くにシンボル名をキャプションとして表示", en: "Show the symbol name as a caption near each symbol" },
            captionAbove: { ja: "シンボルの上にシンボル名を表示", en: "Place the name above the symbol" },
            captionBelow: { ja: "シンボルの下にシンボル名を表示", en: "Place the name below the symbol" },
            fontSize: {
                ja: "キャプションのフォントサイズ。単位は Illustrator の文字設定に従う",
                en: "Caption font size; unit follows Illustrator's type preferences"
            },
            filterAll: { ja: "ドキュメントに登録されているすべてのシンボルを並べる", en: "List every symbol registered in the document" },
            filterUsed: { ja: "ドキュメント内に配置されているシンボルだけを並べる", en: "List only symbols placed in the document" },
            background: {
                none: { ja: "背景の塗りを作成しない", en: "Do not create a background fill" },
                black: { ja: "アートボード背面に黒（K100）の塗りを敷く", en: "Place a solid black (K100) fill behind the artboard" },
                white: { ja: "アートボード背面に白の塗りを敷く", en: "Place a solid white fill behind the artboard" },
                gray: { ja: "アートボード背面にグレー（K50）の塗りを敷く", en: "Place a 50% gray (K50) fill behind the artboard" }
            },
            fitSymbolList: { ja: "作成したアートボードのみをウィンドウに合わせて表示", en: "Fit the created artboard to the window" },
            fitAll: { ja: "すべてのアートボードをウィンドウに合わせて表示", en: "Fit all artboards to the window" }
        },
        button: {
            fitSymbolList: { ja: "シンボル一覧", en: "Symbol list" },
            fitAll: { ja: "全体表示", en: "Fit all" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSymbols: { ja: "登録されているシンボルがありません。", en: "No symbols are registered." },
            noUsedSymbols: { ja: "ドキュメント内で使用中のシンボルがありません。", en: "No symbols are currently used in the document." }
        },
        itemName: {
            symbolList: { ja: "シンボル一覧", en: "Symbol List" }
        }
    };

    /* 作成するレイヤーとアートボードの名前 / Name of the created layer and artboard */
    var SYMBOL_LIST_NAME = getLabel(LABELS.itemName.symbolList);

    /**
     * シンボル一覧のレイヤー名・アートボード名か（別の言語で作成したものも含む）
     * @param {string} name - 調べる名前
     * @returns {boolean} シンボル一覧の名前なら true
     */
    function isSymbolListName(name) {
        return name === LABELS.itemName.symbolList.ja || name === LABELS.itemName.symbolList.en;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * @typedef {object} UnitInfo
     * @property {string} label - 表示する単位名
     * @property {number} factor - 1 単位あたりの pt
     */

    /**
     * 単位のプリファレンスから表示単位と pt への換算係数を取得する
     * （0=inch / 1=mm / 3=pica / 4=cm / 5=Q / 6=px / その他=pt）
     * @param {string} preferenceKey - 整数プリファレンスのキー（"rulerType" / "text/units"）
     * @returns {UnitInfo} 単位情報
     */
    function getUnitInfo(preferenceKey) {
        switch (app.preferences.getIntegerPreference(preferenceKey)) {
            case 0: return { label: "inch", factor: 72.0 };
            case 1: return { label: "mm", factor: 72.0 / 25.4 };
            case 3: return { label: "pica", factor: 12.0 };
            case 4: return { label: "cm", factor: 72.0 / 2.54 };
            case 5: return { label: "Q", factor: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", factor: 1.0 };
            default: return { label: "pt", factor: 1.0 };
        }
    }

    /* ルーラー単位（寸法・余白・間隔）/ Ruler unit for sizes, padding and gaps */
    var RULER_UNIT = getUnitInfo("rulerType");

    /* 文字の単位（フォントサイズ）/ Type unit for font size */
    var TYPE_UNIT = getUnitInfo("text/units");

    /**
     * pt の値を表示単位の整数の文字列にする
     * @param {number} valuePt - 値（pt）
     * @param {UnitInfo} unitInfo - 表示単位
     * @returns {string} 表示用の文字列
     */
    function formatUnitValue(valuePt, unitInfo) {
        return String(Math.round(valuePt / unitInfo.factor));
    }

    /**
     * ルーラー単位の入力欄を pt で読む
     * @param {EditText} valueInput - 入力欄
     * @param {number} defaultPt - 数値でないときの値（pt）
     * @returns {number} 値（pt）
     */
    function readUnitValuePt(valueInput, defaultPt) {
        var parsedValue = parseFloat(valueInput.text);
        return isNaN(parsedValue) ? defaultPt : parsedValue * RULER_UNIT.factor;
    }

    /**
     * 文字の単位のフォントサイズ欄を pt で読む（0 以下や数値以外は既定値）
     * @param {EditText} fontSizeInput - フォントサイズの入力欄
     * @returns {number} フォントサイズ（pt）
     */
    function readFontSizePt(fontSizeInput) {
        var fontSize = parseFloat(fontSizeInput.text);
        return (isNaN(fontSize) || fontSize <= 0) ? DEFAULT_SETTINGS.fontSize : fontSize * TYPE_UNIT.factor;
    }

    /**
     * アートボード番号の文字列を 1 以上の整数にする
     * @param {string} text - 入力された文字列
     * @returns {number} アートボード番号（1 始まり）
     */
    function parseArtboardNumber(text) {
        var artboardNumber = parseInt(text, 10);
        return (isNaN(artboardNumber) || artboardNumber < 1) ? 1 : artboardNumber;
    }

    // =========================================
    // 設定の保存・復元 / Save & restore settings
    // =========================================

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^\uFEFF/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    /* 設定の保存先（Folder.userData/illustrator-scripts/SymbolListBuilder.json）。再起動しても残る。
       以前は Illustrator の環境設定に JSON 風の文字列で保存していたので、新しい保存が無いときだけそこから読み継ぐ
       Settings store (survives a restart); the JSON-like string formerly in Illustrator's preferences is read only until the first save */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent", {
        legacy: function () {
            return readSettingsLegacyPreference(LEGACY_PREF_KEY);
        }
    });

    /**
     * 保存値として使えるか（型が既定値と同じで、選択肢や下限を満たす）
     * @param {string} key - 設定のキー
     * @param {*} value - 保存値
     * @returns {boolean} 使えるなら true
     */
    function isValidSettingValue(key, value) {
        if (typeof value !== typeof DEFAULT_SETTINGS[key]) return false;
        if (SETTING_CHOICES.hasOwnProperty(key)) {
            for (var i = 0; i < SETTING_CHOICES[key].length; i++) {
                if (value === SETTING_CHOICES[key][i]) return true;
            }
            return false;
        }
        if (key === "fontSize") return value > 0;
        if (key === "baseArtboardNumber") return value >= 1;
        return true;
    }

    /**
     * 設定を保存する
     * @param {ListSettings} settings - 保存する設定
     * @returns {void}
     */
    function saveSettings(settings) {
        var savedValues = {};
        for (var i = 0; i < SETTINGS_KEYS.length; i++) {
            savedValues[SETTINGS_KEYS[i]] = settings[SETTINGS_KEYS[i]];
        }
        settingsStore.save(savedValues);
    }

    /**
     * 設定を読み出す（保存値が無い・使えない値は既定値）
     * @returns {ListSettings} 設定
     */
    function loadSettings() {
        var storedSettings = settingsStore.load(DEFAULT_SETTINGS);
        var settings = {};
        for (var i = 0; i < SETTINGS_KEYS.length; i++) {
            var key = SETTINGS_KEYS[i];
            settings[key] = isValidSettingValue(key, storedSettings[key]) ? storedSettings[key] : DEFAULT_SETTINGS[key];
        }
        return settings;
    }

    // =========================================
    // シンボルの収集 / Collect symbols
    // =========================================

    /**
     * 配置されているシンボル名の集合を返す（シンボル一覧レイヤー上のインスタンスは数えない）
     * @param {Document} doc - 対象のドキュメント
     * @returns {object} シンボル名をキー、true を値にしたオブジェクト
     */
    function getUsedSymbolNames(doc) {
        var usedNames = {};
        var symbolItems = doc.symbolItems;
        for (var i = 0; i < symbolItems.length; i++) {
            if (isSymbolListName(symbolItems[i].layer.name)) continue;
            usedNames[symbolItems[i].symbol.name] = true;
        }
        return usedNames;
    }

    /**
     * 収集対象に応じて並べるシンボルを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {string} filter - 収集対象（"all" / "used"）
     * @returns {Symbol[]} 並べるシンボル
     */
    function getTargetSymbols(doc, filter) {
        var usedNames = (filter === "used") ? getUsedSymbolNames(doc) : null;
        var targetSymbols = [];
        for (var i = 0; i < doc.symbols.length; i++) {
            if (!usedNames || usedNames[doc.symbols[i].name]) targetSymbols.push(doc.symbols[i]);
        }
        return targetSymbols;
    }

    // =========================================
    // シンボル一覧の作成 / Build the symbol list
    // =========================================

    /**
     * @typedef {object} SymbolEntry
     * @property {SymbolItem} symbolItem - 配置したシンボルインスタンス
     * @property {number} symbolWidth - シンボルの幅（pt）
     * @property {number} symbolHeight - シンボルの高さ（pt）
     * @property {TextFrame} [caption] - シンボル名のキャプション
     * @property {number} [captionWidth] - キャプションの幅（pt）
     * @property {number} [captionHeight] - キャプションの高さ（pt）
     * @property {number} width - シンボルとキャプションを合わせた枠の幅（pt）
     * @property {number} height - シンボルとキャプションを合わせた枠の高さ（pt）
     * @property {number} x - 枠の左上の相対位置（pt、パッキングで決まる）
     * @property {number} y - 枠の左上の相対位置（pt、下方向が負）
     */

    /**
     * @typedef {object} ArtboardInfo
     * @property {number} index - 作成時のアートボードの index
     * @property {number[]} rect - 作成時の矩形 [left, top, right, bottom]
     */

    /**
     * @typedef {object} SymbolListLayout
     * @property {Layer} layer - シンボル一覧のレイヤー
     * @property {ArtboardInfo} artboardInfo - シンボル一覧のアートボード
     */

    /**
     * シンボル一覧用のレイヤーを作成する
     * @param {Document} doc - 対象のドキュメント
     * @returns {?Layer} 作成したレイヤー（作成できなければ null）
     */
    function createSymbolListLayer(doc) {
        try {
            var listLayer = doc.layers.add();
            listLayer.name = SYMBOL_LIST_NAME;
            return listLayer;
        } catch (e) {
            return null;
        }
    }

    /**
     * ページアイテムの表示上の幅と高さを返す
     * @param {PageItem} pageItem - 対象のアイテム
     * @returns {{width: number, height: number}} 幅と高さ（pt）
     */
    function getVisibleSize(pageItem) {
        var bounds = pageItem.visibleBounds; // [left, top, right, bottom]
        return { width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
    }

    /**
     * 墨の濃度から、ドキュメントのカラーモードに合ったグレーの色を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {number} blackPercent - 墨の濃度（0〜100）
     * @returns {Color} CMYKColor または RGBColor
     */
    function createGrayColor(doc, blackPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0; cmykColor.magenta = 0; cmykColor.yellow = 0; cmykColor.black = blackPercent;
            return cmykColor;
        }
        var channelValue = Math.round(255 * (1 - blackPercent / 100));
        var rgbColor = new RGBColor();
        rgbColor.red = channelValue; rgbColor.green = channelValue; rgbColor.blue = channelValue;
        return rgbColor;
    }

    /**
     * シンボル名のキャプションを作成する
     * @param {Layer} listLayer - 作成先のレイヤー
     * @param {string} symbolName - シンボル名
     * @param {number} fontSizePt - フォントサイズ（pt）
     * @param {?Color} fillColor - 文字の色（null なら既定のまま）
     * @returns {TextFrame} 作成したキャプション
     */
    function createCaption(listLayer, symbolName, fontSizePt, fillColor) {
        var caption = listLayer.textFrames.add();
        caption.contents = symbolName;
        var textRange = caption.textRange;
        textRange.paragraphAttributes.justification = Justification.CENTER;
        if (fillColor) textRange.characterAttributes.fillColor = fillColor;
        /* 範囲外のサイズは既定のまま / Keep the default size when out of range */
        try { textRange.characterAttributes.size = fontSizePt; } catch (e) { }
        /* フォントが無ければ既定のまま / Keep the default font when not installed */
        try { textRange.characterAttributes.textFont = app.textFonts.getByName(CAPTION_FONT_NAMES[uiLang]); } catch (e) { }
        return caption;
    }

    /**
     * シンボルを配置し（必要ならキャプションも作成して）、大きさを測ったエントリを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} listLayer - 配置先のレイヤー
     * @param {ListSettings} settings - 設定
     * @param {Symbol[]} targetSymbols - 並べるシンボル
     * @returns {SymbolEntry[]} エントリ
     */
    function createSymbolEntries(doc, listLayer, settings, targetSymbols) {
        /* 背景が黒のときだけキャプションを白に / White captions only on a black background */
        var captionFillColor = (settings.bgColor === "black") ? createGrayColor(doc, 0) : null;
        var entries = [];
        for (var i = 0; i < targetSymbols.length; i++) {
            var symbolItem = doc.symbolItems.add(targetSymbols[i]);
            symbolItem.moveToBeginning(listLayer);
            var symbolSize = getVisibleSize(symbolItem);
            var entry = {
                symbolItem: symbolItem,
                symbolWidth: symbolSize.width,
                symbolHeight: symbolSize.height,
                width: symbolSize.width,
                height: symbolSize.height
            };

            if (settings.showCaption) {
                entry.caption = createCaption(listLayer, targetSymbols[i].name, settings.fontSize, captionFillColor);
                var captionSize = getVisibleSize(entry.caption);
                entry.captionWidth = captionSize.width;
                entry.captionHeight = captionSize.height;
                entry.width = Math.max(symbolSize.width, captionSize.width);
                entry.height = symbolSize.height + CAPTION_GAP_PT + captionSize.height;
            }
            entries.push(entry);
        }
        return entries;
    }

    /**
     * エントリを左から詰め、最大幅を超えたら次の行に折り返して位置を決める（シェルフパッキング）
     * @param {SymbolEntry[]} entries - エントリ（x, y を書き込む）
     * @param {number} maxRowWidth - 1 行の最大幅（pt）
     * @param {number} gap - エントリ同士の間隔（pt）
     * @returns {{width: number, height: number}} 並べた全体の幅と高さ（pt）
     */
    function packEntriesIntoRows(entries, maxRowWidth, gap) {
        var rowX = 0, rowY = 0, rowHeight = 0, totalWidth = 0;
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            if (rowX > 0 && rowX + entry.width > maxRowWidth) {
                rowY -= rowHeight + gap;
                rowX = 0;
                rowHeight = 0;
            }
            entry.x = rowX;
            entry.y = rowY;
            rowX += entry.width + gap;
            if (rowX - gap > totalWidth) totalWidth = rowX - gap;
            if (entry.height > rowHeight) rowHeight = entry.height;
        }
        return { width: totalWidth, height: -rowY + rowHeight };
    }

    /**
     * キャンバス上で最も右下にあるアートボードの番号を返す（シンボル一覧は除く）
     * 右端が大きく下端が小さいほど右下とみなし、right − bottom で比べる
     * @param {Document} doc - 対象のドキュメント
     * @returns {number} アートボード番号（1 始まり）
     */
    function findBottomRightArtboardNumber(doc) {
        var bestNumber = 1;
        var bestScore = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isSymbolListName(doc.artboards[i].name)) continue;
            var artboardRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
            var score = artboardRect[2] - artboardRect[3];
            if (bestScore === null || score > bestScore) {
                bestScore = score;
                bestNumber = i + 1;
            }
        }
        return bestNumber;
    }

    /**
     * 基準アートボードの矩形を返す
     * - 指定：その番号のアートボード（範囲外は最後のアートボード）
     * - 最終・更新 ON：OK 時に消える既存のシンボル一覧を除いた、最後のアートボード
     * - 最終・更新 OFF：最後のアートボード
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} settings - 設定
     * @returns {number[]} 矩形 [left, top, right, bottom]
     */
    function resolveBaseArtboardRect(doc, settings) {
        var artboards = doc.artboards;
        var lastIndex = artboards.length - 1;
        if (settings.baseMode === "specified") {
            return artboards[Math.min(settings.baseArtboardNumber - 1, lastIndex)].artboardRect;
        }
        if (settings.update) {
            for (var i = lastIndex; i >= 0; i--) {
                if (!isSymbolListName(artboards[i].name)) return artboards[i].artboardRect;
            }
        }
        return artboards[lastIndex].artboardRect;
    }

    /**
     * 基準アートボードの右または下に、新しいアートボードを置く左上の座標を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} settings - 設定
     * @returns {{left: number, top: number}} 左上の座標
     */
    function computeArtboardOrigin(doc, settings) {
        var baseRect = resolveBaseArtboardRect(doc, settings);
        if (settings.position === "right") {
            return { left: baseRect[2] + settings.artboardGap, top: baseRect[1] };
        }
        return { left: baseRect[0], top: baseRect[3] - settings.artboardGap };
    }

    /**
     * シンボル一覧のアートボードを追加する
     * @param {Document} doc - 対象のドキュメント
     * @param {{left: number, top: number}} origin - 左上の座標
     * @param {number} width - 幅（pt）
     * @param {number} height - 高さ（pt）
     * @returns {ArtboardInfo} 追加したアートボードの情報
     */
    function addSymbolListArtboard(doc, origin, width, height) {
        var artboardIndex = doc.artboards.length;
        var rect = [origin.left, origin.top, origin.left + width, origin.top - height];
        doc.artboards.add(rect).name = SYMBOL_LIST_NAME;
        return { index: artboardIndex, rect: rect };
    }

    /**
     * パッキングの結果に従ってシンボルとキャプションを配置する
     * @param {SymbolEntry[]} entries - エントリ
     * @param {{left: number, top: number}} origin - アートボードの左上
     * @param {number} margin - アートボード内側の余白（pt）
     * @param {string} captionPosition - キャプションの位置（"above" / "below"）
     * @returns {void}
     */
    function placeEntries(entries, origin, margin, captionPosition) {
        var isCaptionAbove = (captionPosition === "above");
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            var slotLeft = origin.left + margin + entry.x;
            var slotTop = origin.top - margin + entry.y;

            /* 枠の中で水平中央に揃える / Center horizontally within the slot */
            var symbolTop = (entry.caption && isCaptionAbove) ? slotTop - entry.captionHeight - CAPTION_GAP_PT : slotTop;
            entry.symbolItem.position = [slotLeft + (entry.width - entry.symbolWidth) / 2, symbolTop];

            if (entry.caption) {
                var captionTop = isCaptionAbove ? slotTop : slotTop - entry.symbolHeight - CAPTION_GAP_PT;
                entry.caption.position = [slotLeft + (entry.width - entry.captionWidth) / 2, captionTop];
            }
        }
    }

    /**
     * アートボードいっぱいの背景の塗りを作り、レイヤーの最背面に送る
     * @param {Layer} listLayer - 作成先のレイヤー
     * @param {number[]} rect - アートボードの矩形 [left, top, right, bottom]
     * @param {Color} fillColor - 塗りの色
     * @returns {void}
     */
    function createBackgroundFill(listLayer, rect, fillColor) {
        var backgroundRect = listLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
        backgroundRect.filled = true;
        backgroundRect.stroked = false;
        backgroundRect.fillColor = fillColor;
        backgroundRect.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * シンボル一覧（レイヤー・シンボル・キャプション・アートボード・背景）を作成する
     * 既存のシンボル一覧はここでは消さない（削除は OK 時の removeOtherSymbolLists）
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} settings - 設定
     * @returns {?SymbolListLayout} 作成した一覧（並べるシンボルが無ければ null）
     */
    function buildLayout(doc, settings) {
        var targetSymbols = getTargetSymbols(doc, settings.filter);
        if (targetSymbols.length === 0) return null;
        var listLayer = createSymbolListLayer(doc);
        if (!listLayer) return null;

        var entries = createSymbolEntries(doc, listLayer, settings, targetSymbols);
        var contentSize = packEntriesIntoRows(entries, settings.maxRowWidth, settings.symbolGap);
        var origin = computeArtboardOrigin(doc, settings);
        var artboardWidth = (settings.widthOverridePt > 0) ? settings.widthOverridePt : contentSize.width + settings.margin * 2;
        var artboardHeight = (settings.heightOverridePt > 0) ? settings.heightOverridePt : contentSize.height + settings.margin * 2;
        var artboardInfo = addSymbolListArtboard(doc, origin, artboardWidth, artboardHeight);
        placeEntries(entries, origin, settings.margin, settings.captionPosition);

        if (settings.bgColor !== "none") {
            createBackgroundFill(listLayer, artboardInfo.rect, createGrayColor(doc, BACKGROUND_BLACK_PERCENT[settings.bgColor]));
        }
        return { layer: listLayer, artboardInfo: artboardInfo };
    }

    // =========================================
    // シンボル一覧の削除 / Remove symbol lists
    // =========================================

    /**
     * 2 つの矩形が許容誤差の範囲で一致するか
     * @param {number[]} rectA - 矩形 [left, top, right, bottom]
     * @param {number[]} rectB - 矩形 [left, top, right, bottom]
     * @returns {boolean} 一致すれば true
     */
    function isSameRect(rectA, rectB) {
        for (var i = 0; i < 4; i++) {
            if (Math.abs(rectA[i] - rectB[i]) >= RECT_MATCH_TOLERANCE_PT) return false;
        }
        return true;
    }

    /**
     * 作成したシンボル一覧のアートボードの現在の index を探す（後ろから探す）
     * @param {Document} doc - 対象のドキュメント
     * @param {ArtboardInfo} artboardInfo - 作成時のアートボード情報
     * @returns {number} index（見つからなければ -1）
     */
    function findArtboardIndex(doc, artboardInfo) {
        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            var artboard = doc.artboards[i];
            if (artboard.name === SYMBOL_LIST_NAME && isSameRect(artboard.artboardRect, artboardInfo.rect)) return i;
        }
        return -1;
    }

    /**
     * プレビューで作成したシンボル一覧（アートボードとレイヤー）を削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {SymbolListLayout} layout - 削除する一覧
     * @returns {void}
     */
    function removeLayout(doc, layout) {
        var artboardIndex = findArtboardIndex(doc, layout.artboardInfo);
        if (artboardIndex >= 0 && doc.artboards.length > 1) doc.artboards.remove(artboardIndex);
        layout.layer.remove();
    }

    /**
     * 中心が矩形の内側にあるか（許容幅つき。y は上が正）
     * @param {number[]} itemBounds - アイテムの境界 [left, top, right, bottom]
     * @param {number[]} rect - 矩形 [left, top, right, bottom]
     * @param {number} tolerance - 許容幅（pt）
     * @returns {boolean} 内側なら true
     */
    function isCenterInRect(itemBounds, rect, tolerance) {
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        return centerX >= rect[0] - tolerance &&
            centerX <= rect[2] + tolerance &&
            centerY <= rect[1] + tolerance &&
            centerY >= rect[3] - tolerance;
    }

    /**
     * 旧アートボードの矩形に中心が入るページアイテムを全レイヤーから削除する
     * 保護レイヤー上のアイテムは残す（同じ位置に作った新しい一覧を巻き込まないため）
     * @param {Document} doc - 対象のドキュメント
     * @param {number[]} rect - 旧アートボードの矩形
     * @param {Layer} protectedLayer - 削除しないレイヤー
     * @returns {void}
     */
    function removeItemsInRect(doc, rect, protectedLayer) {
        var pageItems = doc.pageItems;
        var itemsToRemove = [];
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].layer === protectedLayer) continue;
            if (isCenterInRect(pageItems[i].geometricBounds, rect, ARTBOARD_HIT_TOLERANCE_PT)) itemsToRemove.push(pageItems[i]);
        }
        for (var j = 0; j < itemsToRemove.length; j++) {
            /* 親グループごと消えた子やロックされたアイテムは例外になるので飛ばす / Skip children of removed groups and locked items */
            try { itemsToRemove[j].remove(); } catch (e) { }
        }
    }

    /**
     * 確定した一覧を残して、ほかのシンボル一覧（アートボード・その上のアイテム・レイヤー）を削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {SymbolListLayout} keptLayout - 残す一覧
     * @returns {void}
     */
    function removeOtherSymbolLists(doc, keptLayout) {
        var keptIndex = findArtboardIndex(doc, keptLayout.artboardInfo);
        if (keptIndex < 0) keptIndex = keptLayout.artboardInfo.index;

        /* 後ろから消すので、keptIndex より前を消しても比較はずれない / Removing backwards keeps keptIndex valid for comparison */
        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            if (i === keptIndex || doc.artboards.length <= 1 || !isSymbolListName(doc.artboards[i].name)) continue;
            removeItemsInRect(doc, doc.artboards[i].artboardRect, keptLayout.layer);
            doc.artboards.remove(i);
        }

        for (var j = doc.layers.length - 1; j >= 0; j--) {
            var existingLayer = doc.layers[j];
            if (existingLayer === keptLayout.layer || doc.layers.length <= 1 || !isSymbolListName(existingLayer.name)) continue;
            /* ロックされたレイヤーは削除できないので残す / Locked layers cannot be removed */
            try { existingLayer.remove(); } catch (e) { }
        }
    }

    // =========================================
    // 表示 / View
    // =========================================

    /**
     * @typedef {object} ViewState
     * @property {View} view - 対象のビュー
     * @property {number} zoom - ズーム倍率
     * @property {number[]} centerPoint - 表示の中心
     */

    /**
     * 現在のズームと表示の中心を控える
     * @param {Document} doc - 対象のドキュメント
     * @returns {ViewState} 控えた表示状態
     */
    function captureViewState(doc) {
        var activeView = doc.activeView;
        return { view: activeView, zoom: activeView.zoom, centerPoint: activeView.centerPoint };
    }

    /**
     * 控えたズームと表示の中心に戻す
     * @param {ViewState} viewState - 控えた表示状態
     * @returns {void}
     */
    function restoreViewState(viewState) {
        viewState.view.zoom = viewState.zoom;
        viewState.view.centerPoint = viewState.centerPoint;
    }

    /**
     * アートボードをアクティブにし、その中心を指定倍率で表示する
     * @param {Document} doc - 対象のドキュメント
     * @param {number} artboardIndex - アートボードの index
     * @param {number} zoomFactor - ズーム倍率
     * @returns {void}
     */
    function zoomToArtboard(doc, artboardIndex, zoomFactor) {
        doc.artboards.setActiveArtboardIndex(artboardIndex);
        var rect = doc.artboards[artboardIndex].artboardRect; // [left, top, right, bottom]
        var activeView = doc.activeView;
        activeView.zoom = zoomFactor;
        activeView.centerPoint = [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    /**
     * メニューコマンドでウィンドウに合わせた後、少し縮小して表示する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} menuCommand - "fitin"（アクティブなアートボード）/ "fitall"（すべて）
     * @returns {void}
     */
    function fitWindowThenZoomOut(doc, menuCommand) {
        app.executeMenuCommand(menuCommand);
        /* 最小倍率を下回ると例外になるので、そのときはフィットのまま / Stay fitted if below the minimum zoom */
        try { doc.activeView.zoom = doc.activeView.zoom * FIT_ZOOM_RATIO; } catch (e) { }
        app.redraw();
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名・数値入力欄・単位の行を追加する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @param {number} valuePt - 初期値（pt）
     * @param {UnitInfo} unitInfo - 表示単位
     * @param {object} tooltipSet - 項目名と入力欄のツールチップ（ja/en）
     * @returns {EditText} 追加した入力欄
     */
    function addUnitValueRow(parent, labelSet, valuePt, unitInfo, tooltipSet) {
        var valueRow = addRowGroup(parent);
        addFieldLabel(valueRow, labelSet).helpTip = getLabel(tooltipSet);
        var valueInput = addNumberInput(valueRow, formatUnitValue(valuePt, unitInfo), tooltipSet);
        valueRow.add("statictext", undefined, unitInfo.label);
        return valueInput;
    }

    /**
     * 「作成位置」パネル（基準・間隔・方向）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function buildLocationPanel(parent, controls, initialSettings, doc) {
        var locationPanel = addPanel(parent, LABELS.panel.location);

        /* 基準：最終アートボード／指定［番号］（親が異なるので排他はイベント側で処理）
         * Base: last artboard or specified [number] (different parents; exclusivity is handled in events) */
        var baseModeGroup = locationPanel.add("group");
        baseModeGroup.orientation = "column";
        baseModeGroup.alignChildren = "left";
        baseModeGroup.spacing = ROW_SPACING;
        controls.baseLastRadio = addLabeledControl(baseModeGroup, "radiobutton", LABELS.radio.baseLast, LABELS.tooltip.baseLast);
        var specifiedBaseRow = addRowGroup(baseModeGroup);
        controls.baseSpecifiedRadio = addLabeledControl(specifiedBaseRow, "radiobutton", LABELS.radio.baseSpecified, LABELS.tooltip.baseSpecified);
        /* 番号の初期値はキャンバスの最も右下にあるアートボード / Default number: bottom-right-most artboard */
        controls.baseArtboardNumberInput = addNumberInput(specifiedBaseRow, String(findBottomRightArtboardNumber(doc)), LABELS.tooltip.baseArtboardNumber,
            { step: 1, min: 1, integer: true });
        controls.baseLastRadio.value = (initialSettings.baseMode === "last");
        controls.baseSpecifiedRadio.value = (initialSettings.baseMode === "specified");

        controls.artboardGapInput = addUnitValueRow(locationPanel, LABELS.fieldLabel.gap, initialSettings.artboardGap, RULER_UNIT, LABELS.tooltip.artboardGap);

        /* 方向：右側／下側 / Direction: right or below */
        var directionRow = addRowGroup(locationPanel);
        addFieldLabel(directionRow, LABELS.fieldLabel.direction);
        var directionRadioGroup = addRowGroup(directionRow, RADIO_SPACING);
        controls.directionRightRadio = addLabeledControl(directionRadioGroup, "radiobutton", LABELS.radio.directionRight, LABELS.tooltip.directionRight);
        controls.directionBelowRadio = addLabeledControl(directionRadioGroup, "radiobutton", LABELS.radio.directionBelow, LABELS.tooltip.directionBelow);
        controls.directionRightRadio.value = (initialSettings.position === "right");
        controls.directionBelowRadio.value = (initialSettings.position === "below");
    }

    /**
     * 「サイズと余白」パネル（幅・高さ・余白）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildSizePanel(parent, controls, initialSettings) {
        var sizePanel = addPanel(parent, LABELS.panel.sizeAndPadding);
        controls.widthInput = addUnitValueRow(sizePanel, LABELS.fieldLabel.width, 0, RULER_UNIT, LABELS.tooltip.width);
        controls.heightInput = addUnitValueRow(sizePanel, LABELS.fieldLabel.height, 0, RULER_UNIT, LABELS.tooltip.height);
        controls.marginInput = addUnitValueRow(sizePanel, LABELS.fieldLabel.margin, initialSettings.margin, RULER_UNIT, LABELS.tooltip.margin);
    }

    /**
     * 「背景」パネル（なし／黒／白／グレーを 2 行 2 列）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildBackgroundPanel(parent, controls, initialSettings) {
        var backgroundPanel = addPanel(parent, LABELS.panel.background);
        var radioRow;
        controls.backgroundRadios = [];
        for (var i = 0; i < BACKGROUND_CHOICES.length; i++) {
            if (i % 2 === 0) radioRow = addRowGroup(backgroundPanel, RADIO_SPACING);
            var choice = BACKGROUND_CHOICES[i];
            var backgroundRadio = addLabeledControl(radioRow, "radiobutton", LABELS.radio.background[choice], LABELS.tooltip.background[choice]);
            backgroundRadio.preferredSize.width = BACKGROUND_RADIO_WIDTH;
            backgroundRadio.value = (choice === initialSettings.bgColor);
            controls.backgroundRadios.push(backgroundRadio);
        }
    }

    /**
     * 「収集対象」パネル（すべて／使用中のみ）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @returns {void}
     */
    function buildCollectTargetPanel(parent, controls) {
        var collectTargetPanel = addPanel(parent, LABELS.panel.collectTarget);
        var filterRow = addRowGroup(collectTargetPanel, RADIO_SPACING);
        controls.filterAllRadio = addLabeledControl(filterRow, "radiobutton", LABELS.radio.filterAll, LABELS.tooltip.filterAll);
        controls.filterUsedRadio = addLabeledControl(filterRow, "radiobutton", LABELS.radio.filterUsed, LABELS.tooltip.filterUsed);
        /* 保存値は使わず、起動時は常に「すべて」/ Always start with "all" regardless of saved settings */
        controls.filterAllRadio.value = true;
    }

    /**
     * 「並べ方」パネル（間隔・最大幅）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildArrangementPanel(parent, controls, initialSettings) {
        var arrangementPanel = addPanel(parent, LABELS.panel.arrangement);
        controls.symbolGapInput = addUnitValueRow(arrangementPanel, LABELS.fieldLabel.gap, initialSettings.symbolGap, RULER_UNIT, LABELS.tooltip.symbolGap);
        controls.maxRowWidthInput = addUnitValueRow(arrangementPanel, LABELS.fieldLabel.maxRowWidth, initialSettings.maxRowWidth, RULER_UNIT, LABELS.tooltip.maxRowWidth);
    }

    /**
     * 「キャプション」パネル（表示・位置・フォントサイズ）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildCaptionPanel(parent, controls, initialSettings) {
        var captionPanel = addPanel(parent, LABELS.panel.caption);
        controls.showCaptionCheckbox = addLabeledControl(captionPanel, "checkbox", LABELS.checkbox.showCaption, LABELS.tooltip.showCaption);
        controls.showCaptionCheckbox.value = initialSettings.showCaption;

        /* 位置：上／下 / Position: above or below */
        controls.captionPositionRow = addRowGroup(captionPanel);
        controls.captionPositionRow.add("statictext", undefined, labelText(LABELS.fieldLabel.captionPosition));
        var captionPositionRadioGroup = addRowGroup(controls.captionPositionRow, RADIO_SPACING);
        controls.captionAboveRadio = addLabeledControl(captionPositionRadioGroup, "radiobutton", LABELS.radio.captionAbove, LABELS.tooltip.captionAbove);
        controls.captionBelowRadio = addLabeledControl(captionPositionRadioGroup, "radiobutton", LABELS.radio.captionBelow, LABELS.tooltip.captionBelow);
        controls.captionAboveRadio.value = (initialSettings.captionPosition === "above");
        controls.captionBelowRadio.value = (initialSettings.captionPosition === "below");

        /* フォントサイズ：項目名が項目名の幅に収まらないので、入力欄の上に左揃えで置く（単位は環境設定の文字の単位）
         * Font size: the label does not fit the label width, so it sits above the field, left-aligned (type unit preference) */
        controls.fontSizeGroup = captionPanel.add("group");
        controls.fontSizeGroup.orientation = "column";
        controls.fontSizeGroup.alignChildren = "left";
        controls.fontSizeGroup.spacing = ROW_SPACING;
        controls.fontSizeGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.fontSize)).helpTip = getLabel(LABELS.tooltip.fontSize);
        var fontSizeInputRow = addRowGroup(controls.fontSizeGroup);
        controls.fontSizeInput = addNumberInput(fontSizeInputRow, formatUnitValue(initialSettings.fontSize, TYPE_UNIT), LABELS.tooltip.fontSize);
        fontSizeInputRow.add("statictext", undefined, TYPE_UNIT.label);
    }

    /**
     * ダイアログを構築し、コントロールの参照を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} initialSettings - 初期値
     * @returns {object} dialog と各コントロールの参照
     */
    function buildDialog(doc, initialSettings) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dialog);
        var controls = { dialog: dialog };

        /* 2 カラム構成 / Two-column layout */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 左列：作成するアートボード / Left column: new artboard */
        var artboardPanel = addPanel(addColumnGroup(columnsGroup), LABELS.panel.artboard);
        buildLocationPanel(artboardPanel, controls, initialSettings, doc);
        buildSizePanel(artboardPanel, controls, initialSettings);
        buildBackgroundPanel(artboardPanel, controls, initialSettings);
        /* 保存値は使わず、起動時は常に ON / Always start checked regardless of saved settings */
        controls.updateCheckbox = addLabeledControl(artboardPanel, "checkbox", LABELS.checkbox.update, LABELS.tooltip.update);
        controls.updateCheckbox.value = true;

        /* 右列：収集するシンボル / Right column: symbols */
        var symbolsPanel = addPanel(addColumnGroup(columnsGroup), LABELS.panel.symbols);
        buildCollectTargetPanel(symbolsPanel, controls);
        buildArrangementPanel(symbolsPanel, controls, initialSettings);
        buildCaptionPanel(symbolsPanel, controls, initialSettings);

        /* 最下段のボタン行（左：表示合わせ／右：キャンセル・OK）/ Bottom button row (left: fit view / right: Cancel, OK) */
        var buttonRow = addButtonRow(dialog);
        controls.btnFitSymbolList = addLabeledControl(buttonRow.leftGroup, "button", LABELS.button.fitSymbolList, LABELS.tooltip.fitSymbolList);
        controls.btnFitAll = addLabeledControl(buttonRow.leftGroup, "button", LABELS.button.fitAll, LABELS.tooltip.fitAll);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        return controls;
    }

    /**
     * 選択中の背景ラジオの値を返す
     * @param {RadioButton[]} backgroundRadios - 背景ラジオ（BACKGROUND_CHOICES と同じ並び）
     * @returns {string} 背景の選択肢
     */
    function readBackgroundChoice(backgroundRadios) {
        for (var i = 0; i < backgroundRadios.length; i++) {
            if (backgroundRadios[i].value) return BACKGROUND_CHOICES[i];
        }
        return "none";
    }

    /**
     * ダイアログの入力から設定を読む（幅・高さの固定は含まない）
     * @param {object} controls - コントロールの参照
     * @returns {ListSettings} 設定
     */
    function readDialogSettings(controls) {
        return {
            position: controls.directionRightRadio.value ? "right" : "below",
            baseMode: controls.baseSpecifiedRadio.value ? "specified" : "last",
            baseArtboardNumber: parseArtboardNumber(controls.baseArtboardNumberInput.text),
            artboardGap: readUnitValuePt(controls.artboardGapInput, DEFAULT_SETTINGS.artboardGap),
            margin: readUnitValuePt(controls.marginInput, DEFAULT_SETTINGS.margin),
            update: controls.updateCheckbox.value,
            showCaption: controls.showCaptionCheckbox.value,
            filter: controls.filterUsedRadio.value ? "used" : "all",
            symbolGap: readUnitValuePt(controls.symbolGapInput, DEFAULT_SETTINGS.symbolGap),
            maxRowWidth: readUnitValuePt(controls.maxRowWidthInput, DEFAULT_SETTINGS.maxRowWidth),
            bgColor: readBackgroundChoice(controls.backgroundRadios),
            captionPosition: controls.captionBelowRadio.value ? "below" : "above",
            fontSize: readFontSizePt(controls.fontSizeInput)
        };
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * @typedef {object} PreviewSession
     * @property {?SymbolListLayout} currentLayout - 表示中のプレビュー
     * @property {?number} widthOverridePt - 入力で固定した幅（pt、null なら自動）
     * @property {?number} heightOverridePt - 入力で固定した高さ（pt、null なら自動）
     * @property {function(): ListSettings} readSettings - 入力と固定サイズから設定を読む
     * @property {function(): void} clear - プレビューを削除する
     * @property {function(): void} refresh - プレビューを作り直す
     * @property {function(): void} refreshWithAutoSize - 幅・高さの固定を解除して作り直す
     */

    /**
     * 自動サイズのとき、作成したアートボードの幅・高さを入力欄に書き戻す（onChange は発火しない）
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビューの状態
     * @returns {void}
     */
    function writeBackAutoSize(controls, previewSession) {
        var rect = previewSession.currentLayout.artboardInfo.rect;
        if (previewSession.widthOverridePt === null) controls.widthInput.text = formatUnitValue(rect[2] - rect[0], RULER_UNIT);
        if (previewSession.heightOverridePt === null) controls.heightInput.text = formatUnitValue(rect[1] - rect[3], RULER_UNIT);
    }

    /**
     * ダイアログの入力に追従するプレビューを作る（前回のプレビューを消してから作り直す）
     * @param {Document} doc - 対象のドキュメント
     * @param {object} controls - コントロールの参照
     * @returns {PreviewSession} プレビューの状態と操作
     */
    function createPreviewSession(doc, controls) {
        var previewSession = { currentLayout: null, widthOverridePt: null, heightOverridePt: null };

        previewSession.readSettings = function () {
            var settings = readDialogSettings(controls);
            settings.widthOverridePt = previewSession.widthOverridePt;
            settings.heightOverridePt = previewSession.heightOverridePt;
            return settings;
        };
        previewSession.clear = function () {
            if (!previewSession.currentLayout) return;
            removeLayout(doc, previewSession.currentLayout);
            previewSession.currentLayout = null;
        };
        previewSession.refresh = function () {
            previewSession.clear();
            previewSession.currentLayout = buildLayout(doc, previewSession.readSettings());
            if (previewSession.currentLayout) writeBackAutoSize(controls, previewSession);
            app.redraw();
        };
        previewSession.refreshWithAutoSize = function () {
            previewSession.widthOverridePt = null;
            previewSession.heightOverridePt = null;
            previewSession.refresh();
        };
        return previewSession;
    }

    // =========================================
    // イベント / Events
    // =========================================

    /**
     * 複数のコントロールに同じ onClick を設定する
     * @param {object[]} controlList - 対象のコントロール
     * @param {function(): void} handler - クリック時の処理
     * @returns {void}
     */
    function setClickHandler(controlList, handler) {
        for (var i = 0; i < controlList.length; i++) controlList[i].onClick = handler;
    }

    /**
     * 親の異なるラジオボタンを排他にし、選択後の処理を設定する
     * @param {RadioButton[]} radios - 排他にするラジオボタン
     * @param {function(): void} onSelect - 選択後の処理
     * @returns {void}
     */
    function bindExclusiveRadios(radios, onSelect) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].onClick = function () {
                selectExclusiveRadio(radios, this);
                onSelect();
            };
        }
    }

    /**
     * 数値入力欄の確定と∧∨・↑↓キーに同じ処理を設定する
     * @param {EditText} numberInput - addNumberInput() で作った入力欄
     * @param {function(): void} onValueChange - 値が変わったときの処理
     * @param {number} [minValue] - ∧∨・↑↓キーでの下限値（省略時は 0）
     * @returns {void}
     */
    function bindNumberInput(numberInput, onValueChange, minValue) {
        numberInput.onChange = onValueChange;
        if (typeof minValue === "number") numberInput.stepOptions.min = minValue;
        numberInput.stepOptions.onStep = function () { onValueChange(); };
    }

    /**
     * 入力欄と∧∨の有効／無効をまとめて切り替える
     * @param {EditText} numberInput - addNumberInput() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setNumberInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * 「作成位置」パネルのイベントを設定する
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindLocationEvents(controls, previewSession) {
        var baseNumberInput = controls.baseArtboardNumberInput;

        /* 基準が「指定」のときだけ番号を入力できる / Enable the number only when the base is "specified" */
        function updateBaseNumberEnabled() {
            setNumberInputEnabled(baseNumberInput, controls.baseSpecifiedRadio.value);
        }

        bindExclusiveRadios([controls.baseLastRadio, controls.baseSpecifiedRadio], function () {
            updateBaseNumberEnabled();
            previewSession.refreshWithAutoSize();
        });
        setClickHandler([controls.directionRightRadio, controls.directionBelowRadio], previewSession.refreshWithAutoSize);

        /* 番号は 1 未満を 1 に補正してから作り直す / Clamp the number to 1 or more, then rebuild */
        bindNumberInput(baseNumberInput, function () {
            baseNumberInput.text = String(parseArtboardNumber(baseNumberInput.text));
            previewSession.refreshWithAutoSize();
        }, 1);
        updateBaseNumberEnabled();
    }

    /**
     * 幅・高さの入力を設定する（正の値で固定、0 や空欄で自動）
     * 幅を固定したときは［最大幅］も「幅 − 余白 × 2」に合わせる
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindSizeEvents(controls, previewSession) {
        bindNumberInput(controls.widthInput, function () {
            var widthPt = readUnitValuePt(controls.widthInput, 0);
            previewSession.widthOverridePt = (widthPt > 0) ? widthPt : null;
            var maxRowWidthPt = widthPt - 2 * readUnitValuePt(controls.marginInput, DEFAULT_SETTINGS.margin);
            if (widthPt > 0 && maxRowWidthPt > 0) controls.maxRowWidthInput.text = formatUnitValue(maxRowWidthPt, RULER_UNIT);
            previewSession.refresh();
        });
        bindNumberInput(controls.heightInput, function () {
            var heightPt = readUnitValuePt(controls.heightInput, 0);
            previewSession.heightOverridePt = (heightPt > 0) ? heightPt : null;
            previewSession.refresh();
        });
    }

    /**
     * 「キャプション」パネルのイベントを設定する
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindCaptionEvents(controls, previewSession) {
        /* 「シンボル名を表示」OFF のときは位置とフォントサイズをディム / Dim position and font size when captions are off */
        function updateCaptionRowsEnabled() {
            var isCaptionShown = controls.showCaptionCheckbox.value;
            controls.captionPositionRow.enabled = isCaptionShown;
            controls.fontSizeGroup.enabled = isCaptionShown;
            redrawSteppersIn(controls.fontSizeGroup);
        }

        controls.showCaptionCheckbox.onClick = function () {
            updateCaptionRowsEnabled();
            previewSession.refreshWithAutoSize();
        };
        setClickHandler([controls.captionAboveRadio, controls.captionBelowRadio], previewSession.refreshWithAutoSize);
        updateCaptionRowsEnabled();
    }

    /**
     * 表示合わせボタンのイベントを設定する
     * @param {Document} doc - 対象のドキュメント
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindViewButtons(doc, controls, previewSession) {
        controls.btnFitSymbolList.onClick = function () {
            if (!previewSession.currentLayout) return;
            doc.artboards.setActiveArtboardIndex(previewSession.currentLayout.artboardInfo.index);
            fitWindowThenZoomOut(doc, "fitin");
        };
        controls.btnFitAll.onClick = function () {
            fitWindowThenZoomOut(doc, "fitall");
        };
    }

    /**
     * ダイアログのすべてのイベントを設定する
     * @param {Document} doc - 対象のドキュメント
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindDialogEvents(doc, controls, previewSession) {
        bindLocationEvents(controls, previewSession);
        bindSizeEvents(controls, previewSession);
        bindCaptionEvents(controls, previewSession);
        bindViewButtons(doc, controls, previewSession);

        /* 寸法と関係のない操作は幅・高さの固定を解除して作り直す / Other changes release the fixed size and rebuild */
        bindExclusiveRadios(controls.backgroundRadios, previewSession.refreshWithAutoSize);
        setClickHandler([controls.updateCheckbox, controls.filterAllRadio, controls.filterUsedRadio], previewSession.refreshWithAutoSize);
        var numberInputs = [
            controls.artboardGapInput,
            controls.marginInput,
            controls.symbolGapInput,
            controls.maxRowWidthInput,
            controls.fontSizeInput
        ];
        for (var i = 0; i < numberInputs.length; i++) {
            bindNumberInput(numberInputs[i], previewSession.refreshWithAutoSize);
        }
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * OK 時にプレビューを消して一覧を作り直し、更新 ON なら既存の一覧を削除して設定を保存する
     * @param {Document} doc - 対象のドキュメント
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function commitLayout(doc, previewSession) {
        var finalSettings = previewSession.readSettings();
        previewSession.clear();
        var finalLayout = buildLayout(doc, finalSettings);
        if (!finalLayout) {
            alert(getLabel(LABELS.alert.noUsedSymbols));
        } else {
            if (finalSettings.update) removeOtherSymbolLists(doc, finalLayout);
            saveSettings(finalSettings);
        }
        app.redraw();
    }

    /**
     * ダイアログを表示し、OK なら確定、キャンセルならプレビューを消して表示を戻す
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function showDialog(doc) {
        var viewState = captureViewState(doc);
        var controls = buildDialog(doc, loadSettings());
        var previewSession = createPreviewSession(doc, controls);
        bindDialogEvents(doc, controls, previewSession);

        /* 起動直後にプレビューを作り、新しいアートボードの中心を表示 / Initial preview centered on the new artboard */
        previewSession.refresh();
        if (previewSession.currentLayout) {
            zoomToArtboard(doc, previewSession.currentLayout.artboardInfo.index, INITIAL_PREVIEW_ZOOM);
            app.redraw();
        }

        prepareDialogWindow(controls.dialog, SCRIPT_NAME);
        if (controls.dialog.show() === 1) {
            commitLayout(doc, previewSession);
            return;
        }
        previewSession.clear();
        restoreViewState(viewState);
        app.redraw();
    }

    /**
     * ドキュメントとシンボルの有無を確認してダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        if (doc.symbols.length === 0) {
            alert(getLabel(LABELS.alert.noSymbols));
            return;
        }
        showDialog(doc);
    }

    main();

})();
