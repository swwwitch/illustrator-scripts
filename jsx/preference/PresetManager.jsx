#target illustrator
#targetengine "PresetManagerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorの主要な環境設定を、カテゴリ一覧から選んで「デフォルト／現在の値」を並べて確認・変更します。
変更した項目には印と件数が出て、［OK］でまとめて書き込まれます。従来の2カラム表示にも切り替えられます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManager.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n3b33862538f6

### Overview

Reviews and changes the main Illustrator preferences by category, listing each default next to the current value.
Changed items are marked and counted, and everything is written at once when you click OK. You can also switch to the classic two-column view.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManager";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.10.3";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-10";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManager.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManager.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n3b33862538f6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 数値入力欄の許容範囲と既定値 / Allowed range and default for numeric fields */
    /* 「0＝非表示」はチェックボックスOFFで表現するため、入力欄の最小値は1 / "0 = hidden" is expressed by unchecking the box, so the field itself starts at 1 */
    var NUMERIC_INPUT_RULES = {
        recentFonts: { min: 1, max: 30, defaultValue: 15 },     /* 表示数 1〜30 / Visible count 1-30 */
        historyStates: { min: 1, max: 1000, defaultValue: 100 } /* 想定範囲 1〜1000 / Expected range 1-1000 */
    };

    /* 明るさ（4段階スウォッチ）のUIを表示するか。［OK］後に環境設定を開く必要があるため既定は非表示 */
    /* Whether to show the brightness swatches; off by default because it forces Preferences to open after [OK] */
    var SHOW_BRIGHTNESS_UI = false;

    /* アートボードのストローク幅の選択肢 / Selectable artboard stroke widths */
    var ARTBOARD_STROKE_WIDTHS = [1, 2, 3, 4];

    /* アンカーポイントのサイズ（スライダー4段階）が anchorSizePref に書き込む値。既定は先頭の5 */
    /* Values written to anchorSizePref by the four-step anchor point size slider; 5 is the default */
    var ANCHOR_SIZE_LEVELS = [5, 7, 9, 11];

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

    /* ダイアログ固有の寸法 / Dialog-specific sizes */
    var ROW_SPACING = 6;                         /* 行内の要素間隔 / spacing inside a row */
    var PAGE_ROW_SPACING = 4;                    /* カテゴリ内の行の間隔 / spacing between rows in a category */
    var PRESET_ROW_MARGINS = [0, 0, 0, 2];       /* プリセット行の余白 / preset row margins */
    var SIDEBAR_ITEM_HEIGHT = 20;                /* カテゴリ一覧の1行の高さ / height of one category list row */
    var SIDEBAR_PADDING = 12;                    /* カテゴリ一覧の上下の余白の合計 / total vertical padding of the category list */
    var SIDEBAR_TEXT_PADDING = 24;               /* カテゴリ一覧の幅に足す余白（左右の余白とスクロールバー）/ extra list width for side padding and the scrollbar */
    var COLUMN_TEXT_PADDING = 8;                 /* 実測した列幅に足す余白 / extra width added to a measured column */
    var CHANGE_MARKER_WIDTH = 8;                 /* 変更した行の印の列幅 / changed-row marker column width */
    var CHANGE_COUNT_WIDTH = 140;                /* 変更件数の表示幅 / width of the change count */
    var ANCHOR_SLIDER_WIDTH = 90;                /* アンカーサイズのスライダー幅 / anchor size slider width */
    var RECENT_FONTS_INPUT_CHARS = 3;            /* 最近使用したフォントの入力欄の桁数 / recent-fonts field width (chars) */
    var HISTORY_INPUT_CHARS = 4;                 /* ヒストリー数の入力欄の桁数 / history field width (chars) */
    var BRIGHTNESS_SWATCH_SIZE = 23;             /* 明るさスウォッチの一辺(px) / brightness swatch side (px) */
    var BRIGHTNESS_SWATCH_SPACING = 8;           /* スウォッチの間隔 / gap between swatches */

    /* 2カラム表示の寸法（従来と同じ）/ Two-column view sizes (same as the classic layout) */
    var CLASSIC_ROW_SPACING = 10;                  /* 行内の要素間隔 / spacing inside a row */
    var CLASSIC_PRESET_ROW_MARGINS = [0, 10, 20, 20]; /* プリセット行の余白 / preset row margins */
    var CLASSIC_ANCHOR_SLIDER_WIDTH = 110;         /* アンカーサイズのスライダー幅 / anchor size slider width */
    var CLASSIC_GUIDE_LABEL_WIDTH = 80;            /* ガイド内ラベルの共通幅 / unified label width inside Guides */

    /**
     * カテゴリのページを追加し、先頭に列見出し（項目・デフォルト・現在の値）を置く。
     * ページはスタックに重ね、カテゴリ一覧で選んだものだけを表示する
     * Add a category page with the column header; pages are stacked and only the selected one is shown
     * @param {object} dialogControls - コントロールをまとめたオブジェクト（pageStack にページを足す）
     * @param {string} labelPath - カテゴリ名のパス
     * @returns {Group} 追加したページ
     */
    function addCategoryPage(dialogControls, labelPath) {
        var categoryPage = dialogControls.pageStack.add("group");
        categoryPage.orientation = "column";
        categoryPage.alignChildren = ["fill", "top"];
        categoryPage.alignment = ["fill", "top"];
        categoryPage.spacing = PAGE_ROW_SPACING;

        var headerRow = categoryPage.add("group");
        setupRow(headerRow, "fill", ROW_SPACING);
        headerRow.add("statictext", undefined, "").preferredSize = [CHANGE_MARKER_WIDTH, -1];
        var headerTexts = ["column.setting", "column.defaultValue", "column.currentValue"];
        var headerLabels = [];
        for (var i = 0; i < headerTexts.length; i++) {
            var headerLabel = headerRow.add("statictext", undefined, getLabel(headerTexts[i]));
            headerLabel.enabled = false; /* 薄く表示して項目と見分ける / dimmed to set it apart from the rows */
            headerLabels.push(headerLabel);
        }

        /* 列幅は fitColumnWidths() で文字に合わせる / column widths are fitted to the text by fitColumnWidths() */
        dialogControls.categoryPages.push({ page: categoryPage, labelPath: labelPath, rows: [], headerLabels: headerLabels });
        return categoryPage;
    }

    /**
     * 「印・項目名・デフォルト・現在の値」の1行を追加する / Add a "marker / setting / default / current" row
     * @param {Group} categoryPage - 追加先のページ
     * @param {string} labelPath - 項目名のパス
     * @param {string} [tooltipPath] - 項目名のツールチップのパス（省略可）
     * @returns {Group} 値を変えるコントロールを入れるグループ（.settingLabel / .defaultLabel / .changeMarker で参照できる）
     */
    function addSettingRow(categoryPage, labelPath, tooltipPath) {
        var settingRow = categoryPage.add("group");
        setupRow(settingRow, "fill", ROW_SPACING);
        var changeMarker = settingRow.add("statictext", undefined, "•");
        changeMarker.preferredSize = [CHANGE_MARKER_WIDTH, -1];
        changeMarker.visible = false;
        var settingLabel = settingRow.add("statictext", undefined, getLabel(labelPath));
        if (tooltipPath) settingLabel.helpTip = getLabel(tooltipPath);
        var defaultLabel = settingRow.add("statictext", undefined, "");

        var valueGroup = settingRow.add("group");
        setupRow(valueGroup, "left", ROW_SPACING);
        valueGroup.settingLabel = settingLabel;
        valueGroup.defaultLabel = defaultLabel;
        valueGroup.changeMarker = changeMarker;
        return valueGroup;
    }

    /**
     * 行を変更の追跡に登録する。watchedControls を操作するたびに変更件数を数え直す
     * Register a row for change tracking; the count is refreshed whenever a watched control is used
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {Group} valueGroup - addSettingRow() の戻り値
     * @param {function} describeValue - コントロールの状態を表示用の文字列で返す関数
     * @param {Array} watchedControls - click / change を見張るコントロール
     * @returns {void}
     */
    function registerSettingRow(dialogControls, valueGroup, describeValue, watchedControls) {
        var settingRowState = { valueGroup: valueGroup, describeValue: describeValue, baselineText: null };
        dialogControls.settingRows.push(settingRowState);
        dialogControls.categoryPages[dialogControls.categoryPages.length - 1].rows.push(settingRowState);
        var refreshChangeState = function () { updateChangeState(dialogControls); };
        for (var i = 0; i < watchedControls.length; i++) {
            watchedControls[i].addEventListener("click", refreshChangeState);
            watchedControls[i].addEventListener("change", refreshChangeState);
        }
    }

    /**
     * 項目名のクリックでチェックボックスを切り替える（チェックボックス自体は文字なし）
     * Toggle the checkbox when its setting name is clicked (the checkbox itself has no text)
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {StaticText} settingLabel - 項目名
     * @param {Checkbox} targetCheckbox - 切り替えるチェックボックス
     * @returns {void}
     */
    function linkLabelToCheckbox(dialogControls, settingLabel, targetCheckbox) {
        settingLabel.addEventListener("click", function () {
            if (!targetCheckbox.enabled) return;
            targetCheckbox.value = !targetCheckbox.value;
            if (targetCheckbox.onClick) targetCheckbox.onClick();
            updateChangeState(dialogControls);
        });
    }

    /**
     * オン／オフの表示 / On or Off text
     * @param {boolean} isOn - オンなら true
     * @returns {string} 表示文字列
     */
    function formatOnOff(isOn) {
        return getLabel(isOn ? "value.on" : "value.off");
    }

    /**
     * ∧∨付きの整数入力欄を追加する（∧∨と入力欄は隙間0で突き合わせ、↑↓キーも∧∨と同じ処理で増減する）
     * Add an integer field with a stepper butted against it
     * @param {Panel|Group} parentContainer - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {number} fieldCharacters - 入力欄の桁数
     * @param {NumericRule} numericRule - 下限・上限（min/max）
     * @param {function} [onStep] - ∧∨・↑↓キーで増減したあとに呼ぶ関数
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addIntStepperInput(parentContainer, initialText, fieldCharacters, numericRule, onStep) {
        var stepperInputGroup = parentContainer.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, {
            integer: true, min: numericRule.min, max: numericRule.max, onStep: onStep
        });
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = fieldCharacters;
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * ラジオの並びから選択中のラベルを返す / Text of the selected radio in a list
     * @param {RadioButton[]} radioButtons - ラジオボタンの配列
     * @returns {string} 選択中のラベル（未選択なら空文字）
     */
    function getSelectedRadioText(radioButtons) {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) return radioButtons[i].text;
        }
        return "";
    }

    /**
     * 項目の行にラジオボタンを並べ、変更の追跡に登録する / Add a setting row of radio buttons and track it
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {Group} categoryPage - 追加先のページ
     * @param {string} labelPath - 項目名のパス
     * @param {string[]} radioLabelPaths - ラジオのラベルのパス（並び順）
     * @returns {RadioButton[]} 追加したラジオボタン
     */
    function addRadioRow(dialogControls, categoryPage, labelPath, radioLabelPaths) {
        var radioRow = addSettingRow(categoryPage, labelPath);
        var radioButtons = [];
        for (var i = 0; i < radioLabelPaths.length; i++) {
            radioButtons.push(radioRow.add("radiobutton", undefined, getLabel(radioLabelPaths[i])));
        }
        registerSettingRow(dialogControls, radioRow, function () { return getSelectedRadioText(radioButtons); }, radioButtons);
        return radioButtons;
    }

    /**
     * 排他の2択ラジオをまとめて設定 / Set a mutually exclusive pair of radio buttons
     * @param {RadioButton} radioWhenTrue - 条件が真のときに選ぶラジオ
     * @param {RadioButton} radioWhenFalse - 条件が偽のときに選ぶラジオ
     * @param {boolean} isConditionMet - 条件の判定結果
     * @returns {void}
     */
    function setRadioPair(radioWhenTrue, radioWhenFalse, isConditionMet) {
        radioWhenTrue.value = !!isConditionMet;
        radioWhenFalse.value = !isConditionMet;
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

    /* 日英ラベル定義（カテゴリ別に構造化）/ Japanese-English label definitions (grouped by category) */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "環境設定をまとめて変更", en: "Illustrator Preferences Utility" }
        },
        /* 2カラム表示のパネル見出し / Panel titles in the two-column view */
        classicPanel: {
            general: { ja: "［一般］カテゴリ", en: "[General] Category" },
            selectionAnchor: { ja: "［選択範囲・アンカー表示］カテゴリ", en: "[Selection & Anchor Display] Category" },
            artboard: { ja: "アートボード", en: "Artboard" },
            text: { ja: "［テキスト］カテゴリ", en: "[Text] Category" },
            guides: { ja: "ガイド", en: "Guides" },
            smartGuides: { ja: "スマートガイド", en: "Smart Guides" },
            userInterface: { ja: "［ユーザーインターフェイス］カテゴリ", en: "[User Interface] Category" },
            performance: { ja: "［パフォーマンス］カテゴリ", en: "[Performance] Category" },
            fileManagement: { ja: "［ファイル管理］カテゴリ", en: "[File Management] Category" },
            clipboard: { ja: "クリップボードの処理", en: "Clipboard Handling" },
            limitToPath: { ja: "パスに制限", en: "Limit to Path" }
        },
        /* カテゴリ名 / Category names */
        category: {
            general: { ja: "一般", en: "General" },
            selectionAnchor: { ja: "選択範囲・アンカー表示", en: "Selection & Anchor Display" },
            artboard: { ja: "アートボード", en: "Artboard" },
            text: { ja: "テキスト", en: "Type" },
            guides: { ja: "ガイド", en: "Guides" },
            smartGuides: { ja: "スマートガイド", en: "Smart Guides" },
            userInterface: { ja: "ユーザーインターフェイス", en: "User Interface" },
            performance: { ja: "パフォーマンス", en: "Performance" },
            fileManagement: { ja: "ファイル管理", en: "File Handling" },
            clipboard: { ja: "クリップボードの処理", en: "Clipboard Handling" },
            withChangeCount: { ja: "%1（%2）", en: "%1 (%2)" }
        },
        /* 列見出し / Column headers */
        column: {
            setting: { ja: "項目", en: "Setting" },
            defaultValue: { ja: "デフォルト", en: "Default" },
            currentValue: { ja: "現在の値", en: "Current" }
        },
        /* 値の表示 / Value texts */
        value: {
            on: { ja: "オン", en: "On" },
            off: { ja: "オフ", en: "Off" },
            anchorStep: { ja: "%1 / %2", en: "%1 / %2" }
        },
        /* 変更件数 / Change count */
        status: {
            changeCount: { ja: "変更した項目 %1件", en: "%1 changed" }
        },
        /* チェックボックス / Checkboxes */
        checkbox: {
            richToolTips: { ja: "詳細なツールヒントを表示", en: "Show Rich Tool Tips" },
            homeScreen: { ja: "「ホーム画面」を表示", en: "Show the Home Screen" },
            legacyNewDoc: { ja: "以前の「新規ドキュメント」インターフェイス", en: "Legacy \"File > New\" Interface" },
            printBleedWidget: { ja: "「裁ち落としを印刷」生成AIボタンを表示", en: "Show 'Print Bleed' generative AI buttons on Bleed" },
            moveLockedArt: { ja: "ロックまたは非表示オブジェクトを一緒に移動", en: "Move Locked and Hidden Artwork" },
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            zoomToSelection: { ja: "選択範囲へズーム", en: "Zoom to Selection" },
            unlockOnCanvas: { ja: "カンバス上でロック解除", en: "Unlock on Canvas" },
            objectPathOnly: { ja: "オブジェクトの選択範囲をパスに制限", en: "Object Selection by Path Only" },
            textPathOnly: { ja: "テキストオブジェクトの選択範囲をパスに制限", en: "Type Object Selection by Path Only" },
            autoSizeAreaText: { ja: "新規エリア内文字の自動サイズ調整", en: "Auto Size New Area Type" },
            recentFonts: { ja: "最近使用したフォントの表示数", en: "Number of Recent Fonts" },
            missingGlyphProtection: { ja: "見つからない字形の保護を有効にする", en: "Enable Missing Glyph Protection" },
            alternateGlyph: { ja: "選択された文字の異体字を表示", en: "Show Character Alternates" },
            objectHighlighting: { ja: "オブジェクトのハイライト表示", en: "Object Highlighting" },
            animatedZoom: { ja: "アニメーションズーム", en: "Animated Zoom" },
            realTimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-Time Drawing and Editing" },
            editOriginalSystemDefault: { ja: "「オリジナルの編集」にシステムデフォルトを使用", en: "Use System Defaults for ‘Edit Original’" },
            autoActivateFonts: { ja: "Adobe Fonts を自動アクティベート", en: "Auto-activate Adobe Fonts" },
            includeSvgCode: { ja: "SVGコードを含める", en: "Include SVG Code" }
        },
        /* 行ラベル / Field labels */
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            artboardColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            artboardStrokeWidth: { ja: "ストロークの幅", en: "Stroke Width" },
            anchorSize: { ja: "アンカーポイントのサイズ", en: "Anchor Point Size" },
            guideColor: { ja: "カラー", en: "Color" },
            guideStyle: { ja: "スタイル", en: "Style" },
            brightness: { ja: "明るさ", en: "Brightness" },
            canvasColor: { ja: "カンバスカラー", en: "Canvas Color" },
            historyStates: { ja: "ヒストリー数", en: "History States" },
            saveLocation: { ja: "ファイルの保存先", en: "Save Location" },
            updateLinks: { ja: "リンクを更新", en: "Update Links" }
        },
        /* ドロップダウンの項目 / Dropdown items */
        dropdown: {
            presetCurrent: { ja: "現在の設定", en: "Current Settings" },
            presetDefault: { ja: "デフォルト", en: "Default" },
            preset1: { ja: "プリセット1", en: "Preset 1" },
            colorLightBlue: { ja: "ライトブルー", en: "Light Blue" },
            colorLightRed: { ja: "サーモンピンク", en: "Light Red" },
            colorGreen: { ja: "グリーン", en: "Green" },
            colorMediumBlue: { ja: "ミディアムブルー", en: "Medium Blue" },
            colorMagenta: { ja: "マゼンタ", en: "Magenta" },
            colorCyan: { ja: "シアン", en: "Cyan" },
            colorWhite: { ja: "ホワイト", en: "White" },
            colorLightGray: { ja: "ライトグレー", en: "Light Gray" },
            colorBlack: { ja: "ブラック", en: "Black" },
            colorYellow: { ja: "イエロー", en: "Yellow" }
        },
        /* ラジオボタン / Radio buttons */
        radio: {
            guideColorCyan: { ja: "シアン", en: "Cyan" },
            guideColorLightBlue: { ja: "ライトブルー", en: "Light Blue" },
            guideStyleLines: { ja: "ライン", en: "Lines" },
            guideStyleDots: { ja: "点線", en: "Dots" },
            canvasMatch: { ja: "UIに合わせる", en: "Match Brightness" },
            canvasWhite: { ja: "ホワイト", en: "White" },
            saveToComputer: { ja: "コンピューター", en: "Computer" },
            saveToCloud: { ja: "クラウド", en: "Cloud" },
            updateLinksAuto: { ja: "自動", en: "Automatic" },
            updateLinksManual: { ja: "手動", en: "Manual" },
            updateLinksAsk: { ja: "確認", en: "Ask When Modified" }
        },
        /* 明るさスウォッチ / Brightness swatches */
        swatch: {
            dark: { ja: "暗", en: "Dark" },
            mediumDark: { ja: "やや暗", en: "Medium Dark" },
            mediumLight: { ja: "やや明", en: "Medium Light" },
            light: { ja: "明", en: "Light" }
        },
        /* ボタン / Buttons（OKは日英同一なのでリテラル）/ ("OK" is identical in both languages, so it stays a literal) */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            revertAll: { ja: "すべての変更を取り消す", en: "Revert All Changes" },
            toClassicLayout: { ja: "2カラム表示", en: "Two-Column View" },
            toSidebarLayout: { ja: "サイドバー表示", en: "Sidebar View" }
        },
        /* ツールチップ / Tooltips */
        tooltip: {
            preset: {
                ja: "選んだ設定一式をダイアログに反映します。環境設定への保存は［OK］を押したときです",
                en: "Fills the dialog with the chosen set. Nothing is saved until you click OK."
            },
            brightness: {
                ja: "インターフェイスカラーは直接反映できないため、［OK］後に環境設定（ユーザーインターフェイス）が開きます。矢印キー＋Return で確定してください",
                en: "Interface color can’t be applied directly, so Preferences opens after [OK]. Confirm with the arrow keys + Return."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            switchLayout: {
                ja: "レイアウトを切り替えて開き直します。途中の変更は引き継がれます",
                en: "Reopens the dialog in the other layout. Your unsaved changes carry over."
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            historyStates: { ja: "ヒストリー数を設定", en: "Set history states" },
            recentFonts: {
                ja: "OFFにすると一覧を表示しません（0として保存）。表示数は1〜30",
                en: "Turn off to hide the list (saved as 0). The count ranges from 1 to 30."
            },
            unlockOnCanvas: { ja: "カンバス上のオブジェクトとアートボードを選択してロック解除", en: "Select and Unlock Objects and Artboards on Canvas" },
            anchorSize: {
                ja: "アンカーポイント・ハンドル・バウンディングボックスの表示サイズ（4段階）",
                en: "Display size of anchor points, handles and the bounding box (four steps)"
            },
            homeScreen: { ja: "ドキュメントを開いていないときに「ホーム画面」を表示", en: "Show the Home Screen When No Documents Are Open" },
            legacyNewDoc: { ja: "以前の「新規ドキュメント」インターフェイスを使用", en: "Use Legacy \"File > New\" Interface" }
        }
    };
    // =========================================
    // 環境設定の読み書き / Preference accessors
    // =========================================

    /* 注意：Illustrator は未登録のキーを読んでも例外を投げない（返る値は 0 / false とは限らず true / 1 のこともある）ため、
       下記の fallbackValue は「例外が出た場合」にしか効かない。0 が正当な値でないキーは
       safeGetPositiveInt() を使うこと */
    /* Note: Illustrator does not throw on unknown keys (the value may be 0 / false or even true / 1), so the
       fallbacks below only apply on exceptions. Use safeGetPositiveInt() for keys where 0
       is not a valid value */

    /**
     * 例外を握りつぶして処理を続行する共通ラッパー（このスクリプト唯一の try）
     * Shared guard that swallows exceptions so the dialog keeps working (the only try in the script)
     * @param {function} operation - 実行する処理
     * @param {string} actionName - ログに出す処理名
     * @param {string} targetDetail - ログに出す対象（キー名など）
     * @param {*} [fallbackValue] - 失敗したときに返す値
     * @returns {*} operation の戻り値、失敗時は fallbackValue
     */
    function runSafely(operation, actionName, targetDetail, fallbackValue) {
        try {
            return operation();
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] " + actionName + ": " + targetDetail + " / " + e);
            return fallbackValue;
        }
    }

    /**
     * 整数の環境設定を書き込む（失敗しても処理を止めない）/ Write an integer preference (never throws)
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} preferenceValue - 書き込む値
     * @returns {void}
     */
    function safeSetInt(preferenceKey, preferenceValue) {
        runSafely(function () {
            app.preferences.setIntegerPreference(preferenceKey, preferenceValue);
        }, "safeSetInt failed", preferenceKey + " = " + preferenceValue);
    }

    /**
     * 真偽値の環境設定を書き込む / Write a boolean preference (never throws)
     * @param {string} preferenceKey - 環境設定キー
     * @param {boolean} preferenceValue - 書き込む値
     * @returns {void}
     */
    function safeSetBool(preferenceKey, preferenceValue) {
        runSafely(function () {
            app.preferences.setBooleanPreference(preferenceKey, preferenceValue);
        }, "safeSetBool failed", preferenceKey + " = " + preferenceValue);
    }

    /**
     * 実数の環境設定を書き込む / Write a real-number preference (never throws)
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} preferenceValue - 書き込む値
     * @returns {void}
     */
    function safeSetReal(preferenceKey, preferenceValue) {
        runSafely(function () {
            app.preferences.setRealPreference(preferenceKey, preferenceValue);
        }, "safeSetReal failed", preferenceKey + " = " + preferenceValue);
    }

    /**
     * 整数の環境設定を読む。読み取りに失敗したら fallbackValue / Read an integer preference, falling back on failure
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} fallbackValue - 読み取りに失敗したときの値
     * @returns {number} 環境設定の値または fallbackValue
     */
    function safeGetInt(preferenceKey, fallbackValue) {
        return runSafely(function () {
            return app.preferences.getIntegerPreference(preferenceKey);
        }, "safeGetInt fallback", preferenceKey, fallbackValue);
    }

    /**
     * 正の整数として環境設定を読む。0以下は「キーが未登録」とみなして fallbackValue を返す
     * Read an integer preference as a positive number; zero or less means "key not present"
     * @param {string} preferenceKey - 環境設定キー（0以下が正当な値ではないキーに限る）
     * @param {number} fallbackValue - キーが未登録・読み取り失敗のときの値
     * @returns {number} 正の整数、または fallbackValue
     */
    function safeGetPositiveInt(preferenceKey, fallbackValue) {
        var storedValue = safeGetInt(preferenceKey, fallbackValue);
        return (storedValue > 0) ? storedValue : fallbackValue;
    }

    /**
     * 真偽値の環境設定を読む / Read a boolean preference, falling back on failure
     * @param {string} preferenceKey - 環境設定キー
     * @param {boolean} fallbackValue - 読み取りに失敗したときの値
     * @returns {boolean} 環境設定の値または fallbackValue
     */
    function safeGetBool(preferenceKey, fallbackValue) {
        return runSafely(function () {
            return app.preferences.getBooleanPreference(preferenceKey);
        }, "safeGetBool fallback", preferenceKey, fallbackValue);
    }

    /**
     * 実数の環境設定を読む / Read a real-number preference, falling back on failure
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} fallbackValue - 読み取りに失敗したときの値
     * @returns {number} 環境設定の値または fallbackValue
     */
    function safeGetReal(preferenceKey, fallbackValue) {
        return runSafely(function () {
            return app.preferences.getRealPreference(preferenceKey);
        }, "safeGetReal fallback", preferenceKey, fallbackValue);
    }

    // =========================================
    // 数値ユーティリティ / Numeric utilities
    // =========================================

    /**
     * @typedef {object} NumericRule
     * @property {number} min - 許容する最小値
     * @property {number} max - 許容する最大値
     * @property {number} defaultValue - 数値として解釈できないときの既定値
     */

    /**
     * 値を min〜max に収める / Clamp a value into the min-max range
     * @param {number} targetValue - 対象の値
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @returns {number} 範囲内に収めた値
     */
    function clamp(targetValue, minValue, maxValue) {
        return Math.max(minValue, Math.min(maxValue, targetValue));
    }

    /**
     * 整数に変換。数値でなければ null / Parse as an integer, or null when not numeric
     * @param {string|number} rawValue - 変換する値
     * @returns {number|null} 整数、または変換できない場合は null
     */
    function parseIntOrNull(rawValue) {
        var parsedInt = parseInt(rawValue, 10);
        return isNaN(parsedInt) ? null : parsedInt;
    }

    /**
     * ルール（min/max/defaultValue）に沿って整数へ補正 / Normalize a value against a min/max/default rule
     * @param {string|number} rawValue - 対象の値
     * @param {NumericRule} numericRule - 適用するルール
     * @returns {number} 補正後の整数
     */
    function clampIntToRule(rawValue, numericRule) {
        var parsedInt = parseIntOrNull(rawValue);
        if (parsedInt === null) return numericRule.defaultValue;
        return clamp(parsedInt, numericRule.min, numericRule.max);
    }

    /**
     * 入力欄の値を補正して書き戻す / Normalize an edittext value in place and return it
     * @param {EditText} inputField - 対象の入力欄
     * @param {NumericRule} numericRule - 適用するルール
     * @returns {number} 補正後の整数
     */
    function normalizeIntInput(inputField, numericRule) {
        var normalizedValue = clampIntToRule(inputField.text, numericRule);
        inputField.text = String(normalizedValue);
        return normalizedValue;
    }

    /**
     * 距離が最も小さい要素の index を返す（離散的な選択肢へのスナップ用）
     * Index of the item with the smallest distance (snaps to discrete choices)
     * @param {Array} candidateItems - 候補の配列
     * @param {function} measureDistance - 要素を受け取り距離（数値）を返す関数
     * @returns {number} 最も近い要素の index（配列が空なら 0）
     */
    function findNearestIndexBy(candidateItems, measureDistance) {
        var nearestIndex = 0;
        var nearestDistance = Infinity;
        for (var i = 0; i < candidateItems.length; i++) {
            var distance = measureDistance(candidateItems[i]);
            if (distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /**
     * 数値配列のうち目標値に最も近い要素の index / Index of the number closest to a target value
     * @param {number[]} candidateValues - 選択肢の数値配列
     * @param {number} targetValue - 目標値
     * @returns {number} 最も近い要素の index
     */
    function findNearestIndex(candidateValues, targetValue) {
        return findNearestIndexBy(candidateValues, function (candidateValue) {
            return Math.abs(candidateValue - targetValue);
        });
    }

    /**
     * 許容誤差つきで比較 / Compare two numbers with a tolerance
     * @param {number} firstValue - 比較する値
     * @param {number} secondValue - 比較する値
     * @param {number} [tolerance] - 許容誤差（省略時は 0.02）
     * @returns {boolean} 許容誤差の範囲内なら true
     */
    function almostEqual(firstValue, secondValue, tolerance) {
        if (typeof tolerance !== "number") tolerance = 0.02;
        return Math.abs(firstValue - secondValue) < tolerance;
    }

    // =========================================
    // 設定値の定義 / Preference value definitions
    // =========================================

    /* アートボードのハイライトカラープリセット（RGB 0..1）/ Artboard highlight color presets (RGB 0..1) */
    var ARTBOARD_COLOR_PRESETS = [
        { label: getLabel("dropdown.colorLightBlue"), red: 0.29, green: 0.52, blue: 1.0 },
        { label: getLabel("dropdown.colorLightRed"), red: 1.0, green: 0.29, blue: 0.29 },
        { label: getLabel("dropdown.colorGreen"), red: 0.0, green: 0.65, blue: 0.31 },
        { label: getLabel("dropdown.colorMediumBlue"), red: 0.0, green: 0.45, blue: 0.78 },
        { label: getLabel("dropdown.colorMagenta"), red: 1.0, green: 0.0, blue: 1.0 },
        { label: getLabel("dropdown.colorCyan"), red: 0.0, green: 1.0, blue: 1.0 },
        { label: getLabel("dropdown.colorWhite"), red: 1.0, green: 1.0, blue: 1.0 },
        { label: getLabel("dropdown.colorLightGray"), red: 0.65, green: 0.65, blue: 0.65 },
        { label: getLabel("dropdown.colorBlack"), red: 0.0, green: 0.0, blue: 0.0 },
        { label: getLabel("dropdown.colorYellow"), red: 1.0, green: 1.0, blue: 0.0 }
    ];

    /* ガイドカラーの2択（RGB 0..1）/ The two selectable guide colors (RGB 0..1) */
    var GUIDE_COLOR_CYAN = { red: 0.0, green: 1.0, blue: 1.0 };
    var GUIDE_COLOR_LIGHT_BLUE = { red: 0.29, green: 0.52, blue: 1.0 };

    /* ガイドスタイルの値 / Guide style values */
    var GUIDE_STYLE_LINES = 0;
    var GUIDE_STYLE_DOTS = 1;

    /* 「リンクを更新」の値（ラジオの並び順と同じ）/ "Update Links" values (same order as the radios) */
    var UPDATE_LINKS_AUTO = 0;
    var UPDATE_LINKS_MANUAL = 1;
    var UPDATE_LINKS_ASK = 2;

    /* UI明るさの4プリセット：uiBrightness 値・スウォッチのシェード（RGB 0..1）・ラベル / Four UI-brightness presets: uiBrightness value, swatch shade (RGB 0..1), label */
    /* 値は連続値ではなく離散プリセット（0.5 と 0.50999999 が別段階）/ Values are discrete presets, not a continuous scale (0.5 vs 0.50999999 are distinct steps) */
    var BRIGHTNESS_LEVELS = [
        { value: 0.0, shade: [0.22, 0.22, 0.22], labelPath: "swatch.dark" },
        { value: 0.5, shade: [0.33, 0.33, 0.33], labelPath: "swatch.mediumDark" },
        { value: 0.50999999046326, shade: [0.70, 0.70, 0.70], labelPath: "swatch.mediumLight" },
        { value: 1.0, shade: [0.94, 0.94, 0.94], labelPath: "swatch.light" }
    ];
    var DEFAULT_BRIGHTNESS_INDEX = 1;                       /* デフォルトの列に出す段階（やや暗）/ Step shown in the default column (Medium Dark) */
    var BRIGHTNESS_SELECTED_BORDER = [0.15, 0.5, 0.92];     /* 選択枠の青 / Blue selection border */
    var BRIGHTNESS_SWATCH_OUTLINE = [0.5, 0.5, 0.5];        /* 通常時の細枠 / Thin outline when not selected */
    var BRIGHTNESS_TOLERANCE = 0.001;                       /* プリセット同士の判定用（0.5 と 0.50999999 を区別）/ Preset comparison tolerance (keeps 0.5 and 0.50999999 distinct) */

    /* ［デフォルト］の設定一式。全項目を必ず定義する / [Default] preset; every field must be defined */
    var PRESET_STATE_DEFAULT = {
        richToolTips: true,
        homeScreen: true,
        legacyNewDoc: false,
        printBleedWidget: true,
        moveLockedArt: false,
        objectPathOnly: false,
        zoomToSelection: true,
        unlockOnCanvas: false,
        anchorSize: 5, /* anchorSizePref の値（5/7/9/11）/ anchorSizePref value (5/7/9/11) */
        textPathOnly: false,
        showArtboardName: true,
        artboardColorIndex: 0,
        artboardStrokeWidth: 1,
        autoSizeAreaText: false,
        recentFontsEnabled: true,
        recentFontsCount: 10,
        missingGlyphProtection: true,
        alternateGlyph: true,
        objectHighlighting: true,
        animatedZoom: true,
        historyStates: 100,
        realTimeDrawing: true,
        editOriginalSystemDefault: false,
        autoActivateFonts: false,
        canvasWhite: false,
        guideColorIsLightBlue: false,
        guideStyleIsDots: false,
        saveToCloud: true,
        updateLinks: UPDATE_LINKS_ASK,
        includeSvgCode: false
    };

    /* ［プリセット1］の設定一式。全項目を必ず定義する / [Preset 1]; every field must be defined */
    var PRESET_STATE_1 = {
        richToolTips: false,
        homeScreen: false,
        legacyNewDoc: true,
        printBleedWidget: false,
        moveLockedArt: true,
        objectPathOnly: false,
        zoomToSelection: false,
        unlockOnCanvas: false,
        anchorSize: 7, /* anchorSizePref の値（5/7/9/11）/ anchorSizePref value (5/7/9/11) */
        textPathOnly: false,
        showArtboardName: false,
        artboardColorIndex: 7, /* ライトグレー（アートボード表示パレットの「ライト」と同じ）/ Light Gray (same as "Light" in the artboard display palette) */
        artboardStrokeWidth: 1,
        autoSizeAreaText: true,
        recentFontsEnabled: true,
        recentFontsCount: 15,
        missingGlyphProtection: false,
        alternateGlyph: false,
        objectHighlighting: false,
        animatedZoom: false,
        historyStates: 50,
        realTimeDrawing: false,
        editOriginalSystemDefault: true,
        autoActivateFonts: true,
        canvasWhite: true,
        guideColorIsLightBlue: true,
        guideStyleIsDots: false,
        saveToCloud: false,
        updateLinks: UPDATE_LINKS_AUTO,
        includeSvgCode: true
    };

    /* プリセットのドロップダウン項目（並び順どおり）。state が null の項目は現在の環境設定を読み込む */
    /* Preset dropdown entries in display order; a null state reloads the current preferences */
    var PRESET_CHOICES = [
        { labelPath: "dropdown.presetCurrent", state: null },
        { labelPath: "dropdown.presetDefault", state: PRESET_STATE_DEFAULT },
        { labelPath: "dropdown.preset1", state: PRESET_STATE_1 }
    ];

    // =========================================
    // 設定値のヘルパー / Preference value helpers
    // =========================================

    /**
     * 指定RGBに最も近いハイライトカラーの index / Index of the highlight color nearest to the given RGB
     * @param {number} red - 赤成分（0〜1）
     * @param {number} green - 緑成分（0〜1）
     * @param {number} blue - 青成分（0〜1）
     * @returns {number} ARTBOARD_COLOR_PRESETS の index
     */
    function findNearestArtboardColorIndex(red, green, blue) {
        return findNearestIndexBy(ARTBOARD_COLOR_PRESETS, function (colorPreset) {
            return Math.abs(colorPreset.red - red) +
                Math.abs(colorPreset.green - green) +
                Math.abs(colorPreset.blue - blue);
        });
    }

    /**
     * uiBrightness 値に最も近いプリセットの index / Index of the brightness preset closest to a uiBrightness value
     * @param {number} brightnessValue - uiBrightness の値
     * @returns {number} BRIGHTNESS_LEVELS の index
     */
    function findNearestBrightnessIndex(brightnessValue) {
        return findNearestIndexBy(BRIGHTNESS_LEVELS, function (brightnessLevel) {
            return Math.abs(brightnessLevel.value - brightnessValue);
        });
    }

    /**
     * 現在のガイドカラーがライトブルーか / Whether the current guide color is Light Blue
     * @returns {boolean} ライトブルーなら true
     */
    function isGuideColorLightBlue() {
        return almostEqual(safeGetReal("Guide/Color/red", GUIDE_COLOR_CYAN.red), GUIDE_COLOR_LIGHT_BLUE.red) &&
            almostEqual(safeGetReal("Guide/Color/green", GUIDE_COLOR_CYAN.green), GUIDE_COLOR_LIGHT_BLUE.green) &&
            almostEqual(safeGetReal("Guide/Color/blue", GUIDE_COLOR_CYAN.blue), GUIDE_COLOR_LIGHT_BLUE.blue);
    }

    /**
     * ガイドカラーを書き込む / Write the guide color
     * @param {{red: number, green: number, blue: number}} guideColor - 書き込むRGB（0〜1）
     * @returns {void}
     */
    function writeGuideColor(guideColor) {
        safeSetReal("Guide/Color/red", guideColor.red);
        safeSetReal("Guide/Color/green", guideColor.green);
        safeSetReal("Guide/Color/blue", guideColor.blue);
    }

    // =========================================
    // 明るさスウォッチ / Brightness swatches
    // =========================================

    /**
     * スウォッチの描画ハンドラを作る：塗り＋（選択時のみ）青い枠
     * Build a swatch draw handler: fill + (only when selected) a blue border
     * @param {object} brightnessPicker - addBrightnessPicker() の戻り値
     * @param {number} levelIndex - BRIGHTNESS_LEVELS の index
     * @returns {function} onDraw に割り当てる関数
     */
    function createBrightnessSwatchDrawHandler(brightnessPicker, levelIndex) {
        return function () {
            var swatchGraphics = this.graphics;
            var swatchWidth = this.size[0];
            var swatchHeight = this.size[1];
            var brightnessLevel = BRIGHTNESS_LEVELS[levelIndex];

            /* シェードで塗りつぶし / Fill with the shade */
            var fillBrush = swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, brightnessLevel.shade.concat(1));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(fillBrush);

            if (levelIndex === brightnessPicker.selectedIndex) {
                /* 選択：太めの青枠 / Selected: thicker blue border */
                var selectedPen = swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, BRIGHTNESS_SELECTED_BORDER.concat(1), 2);
                swatchGraphics.newPath();
                swatchGraphics.rectPath(1, 1, swatchWidth - 2, swatchHeight - 2);
                swatchGraphics.strokePath(selectedPen);
            } else {
                /* 非選択：細いグレー枠（明シェードを明背景でも視認）/ Not selected: thin gray outline (keep light swatches visible) */
                var outlinePen = swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, BRIGHTNESS_SWATCH_OUTLINE.concat(1), 1);
                swatchGraphics.newPath();
                swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
                swatchGraphics.strokePath(outlinePen);
            }
        };
    }

    /**
     * 全スウォッチを再描画（hide/show で onDraw を確実に再実行し、旧選択枠を残さない＝排他表示）
     * Force all swatches to repaint (hide/show reliably re-runs onDraw so the old selection border never lingers = exclusive)
     * @param {object} brightnessPicker - addBrightnessPicker() の戻り値
     * @returns {void}
     */
    function refreshBrightnessSwatches(brightnessPicker) {
        for (var i = 0; i < brightnessPicker.swatchButtons.length; i++) {
            brightnessPicker.swatchButtons[i].hide();
            brightnessPicker.swatchButtons[i].show();
        }
    }

    /**
     * uiBrightness 値から選択状態だけ更新（書き込みは［OK］時）
     * Update the selection only, from a uiBrightness value (writes happen on [OK])
     * @param {object} brightnessPicker - addBrightnessPicker() の戻り値
     * @param {number} brightnessValue - uiBrightness の値
     * @returns {void}
     */
    function selectBrightnessByValue(brightnessPicker, brightnessValue) {
        var nearestIndex = findNearestBrightnessIndex(brightnessValue);
        if (nearestIndex === brightnessPicker.selectedIndex) return;
        brightnessPicker.selectedIndex = nearestIndex;
        refreshBrightnessSwatches(brightnessPicker);
    }

    /**
     * 明るさの行（ラベル＋4段階のスウォッチ）を追加。SHOW_BRIGHTNESS_UI が false なら状態だけ持つ
     * Add the brightness row (label + four swatches); only keeps state when SHOW_BRIGHTNESS_UI is false
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {Group|Panel} parentContainer - 追加先（サイドバー表示はページ、2カラム表示はパネル）
     * @param {boolean} [isClassicLayout] - 2カラム表示なら true（項目名の行にし、変更の追跡には載せない）
     * @returns {{swatchButtons: Array, selectedIndex: number, isTouched: boolean}} 明るさの選択状態
     */
    function addBrightnessPicker(dialogControls, parentContainer, isClassicLayout) {
        /* isTouched：ユーザーがスウォッチを操作したか（未操作なら書き込まない）/ Whether the user clicked a swatch (no write when untouched) */
        var brightnessPicker = { swatchButtons: [], selectedIndex: -1, isTouched: false };
        if (!SHOW_BRIGHTNESS_UI) return brightnessPicker;

        var brightnessRow = isClassicLayout
            ? addClassicLabeledRow(parentContainer, "fieldLabel.brightness")
            : addSettingRow(parentContainer, "fieldLabel.brightness");
        var swatchGroup = brightnessRow.add("group");
        setupRow(swatchGroup, "left", BRIGHTNESS_SWATCH_SPACING);

        for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
            var swatchButton = swatchGroup.add("iconbutton", undefined, undefined, { style: "toolbutton" });
            swatchButton.preferredSize = [BRIGHTNESS_SWATCH_SIZE, BRIGHTNESS_SWATCH_SIZE];
            swatchButton.alignment = ["left", "center"];
            swatchButton.helpTip = getLabel(BRIGHTNESS_LEVELS[i].labelPath) + " / " + getLabel("tooltip.brightness");
            swatchButton.onDraw = createBrightnessSwatchDrawHandler(brightnessPicker, i);
            swatchButton.onClick = (function (levelIndex) {
                return function () {
                    /* UIのみ更新し、保存は［OK］時 / UI only; persist on [OK] */
                    brightnessPicker.selectedIndex = levelIndex;
                    brightnessPicker.isTouched = true;
                    refreshBrightnessSwatches(brightnessPicker);
                    updateChangeState(dialogControls);
                };
            })(i);
            brightnessPicker.swatchButtons.push(swatchButton);
        }
        if (isClassicLayout) return brightnessPicker;
        /* 選択は onClick で変わるので、見張りは onClick に任せる / the selection changes in onClick, which refreshes the count itself */
        registerSettingRow(dialogControls, brightnessRow, function () {
            var selectedLevel = BRIGHTNESS_LEVELS[brightnessPicker.selectedIndex];
            return selectedLevel ? getLabel(selectedLevel.labelPath) : "";
        }, []);
        return brightnessPicker;
    }

    /**
     * UIの明るさを書き込む。値の書き込みだけでは反映されないため、変更したときだけ書き込む
     * Write the UI brightness; writing the value alone does not apply it, so write only when it changed
     * @param {object} brightnessPicker - addBrightnessPicker() の戻り値
     * @returns {boolean} 環境設定（ユーザーインターフェイス）を開く必要があるか
     */
    function saveBrightness(brightnessPicker) {
        var selectedBrightnessLevel = BRIGHTNESS_LEVELS[brightnessPicker.selectedIndex];
        if (!brightnessPicker.isTouched || !selectedBrightnessLevel) return false;
        if (almostEqual(safeGetReal("uiBrightness", 0.0), selectedBrightnessLevel.value, BRIGHTNESS_TOLERANCE)) return false;
        safeSetReal("uiBrightness", selectedBrightnessLevel.value);
        return true;
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
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * チェックボックスの行を追加し、対応表と変更の追跡に登録する / Add a checkbox row, register it in the binding table and track it
     * 対応表（dialogControls.checkboxBindings）は読み込み・保存・プリセット適用の3か所が共通で参照する
     * The binding table is shared by load, save and preset apply
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {Group} categoryPage - 追加先のページ
     * @param {{label: string, key: string, preset: string, type: string, tooltip: string}} checkboxSpec
     *        label=ラベルのパス / key=環境設定キー / preset=プリセット項目名 /
     *        type=省略時 "bool"、0=ON・1=OFF の反転キーは "invertedInt" / tooltip=ツールチップのパス（省略可）
     * @returns {Checkbox} 追加したチェックボックス
     */
    function bindCheckbox(dialogControls, categoryPage, checkboxSpec) {
        var valueGroup = addSettingRow(categoryPage, checkboxSpec.label, checkboxSpec.tooltip);
        var newCheckbox = valueGroup.add("checkbox", undefined, "");
        newCheckbox.helpTip = getLabel(checkboxSpec.tooltip || checkboxSpec.label);
        linkLabelToCheckbox(dialogControls, valueGroup.settingLabel, newCheckbox);
        registerSettingRow(dialogControls, valueGroup, function () { return formatOnOff(newCheckbox.value); }, [newCheckbox]);
        dialogControls.checkboxBindings.push({
            key: checkboxSpec.key,
            valueType: checkboxSpec.type || "bool",
            control: newCheckbox,
            presetField: checkboxSpec.preset
        });
        return newCheckbox;
    }

    /**
     * ［一般］［選択範囲・アンカー表示］［アートボード］［テキスト］のページを組み立てる
     * Build the General, Selection & Anchor Display, Artboard and Type pages
     * @param {object} dialogControls - コントロールを登録するオブジェクト
     * @returns {void}
     */
    function buildDisplayPages(dialogControls) {
        var refreshChangeState = function () { updateChangeState(dialogControls); };

        /* ［一般］/ General */
        var generalPage = addCategoryPage(dialogControls, "category.general");
        bindCheckbox(dialogControls, generalPage, { label: "checkbox.richToolTips", key: "showRichToolTips", preset: "richToolTips" });
        bindCheckbox(dialogControls, generalPage, { label: "checkbox.homeScreen", key: "Hello/ShowHomeScreenWS", preset: "homeScreen", tooltip: "tooltip.homeScreen" });
        bindCheckbox(dialogControls, generalPage, { label: "checkbox.legacyNewDoc", key: "Hello/NewDoc", preset: "legacyNewDoc", tooltip: "tooltip.legacyNewDoc" });
        bindCheckbox(dialogControls, generalPage, { label: "checkbox.printBleedWidget", key: "enablePrintBleedWidget", preset: "printBleedWidget" });

        /* ［選択範囲・アンカー表示］。パスに制限は 0=ON / 1=OFF の反転キー / Selection & Anchor Display; the Limit to Path keys store 0 = ON */
        var selectionAnchorPage = addCategoryPage(dialogControls, "category.selectionAnchor");
        bindCheckbox(dialogControls, selectionAnchorPage, { label: "checkbox.zoomToSelection", key: "zoomToSelection", preset: "zoomToSelection" });
        bindCheckbox(dialogControls, selectionAnchorPage, { label: "checkbox.unlockOnCanvas", key: "showLockIcon", preset: "unlockOnCanvas", tooltip: "tooltip.unlockOnCanvas" });
        bindCheckbox(dialogControls, selectionAnchorPage, { label: "checkbox.objectPathOnly", key: "hitShapeOnPreview", preset: "objectPathOnly", type: "invertedInt" });
        bindCheckbox(dialogControls, selectionAnchorPage, { label: "checkbox.textPathOnly", key: "hitTypeShapeOnPreview", preset: "textPathOnly", type: "invertedInt" });

        /* アンカーポイントのサイズ（4段階スライダー）/ Anchor point size (four-step slider) */
        var anchorSizeRow = addSettingRow(selectionAnchorPage, "fieldLabel.anchorSize", "tooltip.anchorSize");
        var anchorSizeSlider = anchorSizeRow.add("slider", undefined, 1, 1, ANCHOR_SIZE_LEVELS.length);
        anchorSizeSlider.preferredSize = [ANCHOR_SLIDER_WIDTH, -1];
        anchorSizeSlider.helpTip = getLabel("tooltip.anchorSize");
        anchorSizeSlider.addEventListener("changing", refreshChangeState);
        dialogControls.anchorSizeSlider = anchorSizeSlider;
        registerSettingRow(dialogControls, anchorSizeRow, function () {
            var anchorStep = clamp(Math.round(anchorSizeSlider.value), 1, ANCHOR_SIZE_LEVELS.length);
            return getLabel("value.anchorStep", [anchorStep, ANCHOR_SIZE_LEVELS.length]);
        }, [anchorSizeSlider]);

        /* アートボード / Artboard */
        var artboardPage = addCategoryPage(dialogControls, "category.artboard");
        bindCheckbox(dialogControls, artboardPage, { label: "checkbox.moveLockedArt", key: "moveLockedAndHiddenArt", preset: "moveLockedArt" });
        bindCheckbox(dialogControls, artboardPage, { label: "checkbox.showArtboardName", key: "showArtboardLabelOnCanvas", preset: "showArtboardName" });

        /* ハイライトのカラー（9色のプリセットから選択）/ Highlight color (nine presets) */
        var artboardColorRow = addSettingRow(artboardPage, "fieldLabel.artboardColor");
        var artboardColorLabels = [];
        for (var i = 0; i < ARTBOARD_COLOR_PRESETS.length; i++) {
            artboardColorLabels.push(ARTBOARD_COLOR_PRESETS[i].label);
        }
        var artboardColorDropdown = artboardColorRow.add("dropdownlist", undefined, artboardColorLabels);
        dialogControls.artboardColorDropdown = artboardColorDropdown;
        registerSettingRow(dialogControls, artboardColorRow, function () {
            return artboardColorDropdown.selection ? artboardColorDropdown.selection.text : "";
        }, [artboardColorDropdown]);

        /* ストロークの幅（1〜4）/ Stroke width (1-4) */
        var artboardStrokeWidthRow = addSettingRow(artboardPage, "fieldLabel.artboardStrokeWidth");
        var artboardStrokeWidthRadios = [];
        for (var j = 0; j < ARTBOARD_STROKE_WIDTHS.length; j++) {
            artboardStrokeWidthRadios.push(
                artboardStrokeWidthRow.add("radiobutton", undefined, String(ARTBOARD_STROKE_WIDTHS[j]))
            );
        }
        dialogControls.artboardStrokeWidthRadios = artboardStrokeWidthRadios;
        registerSettingRow(dialogControls, artboardStrokeWidthRow, function () {
            return getSelectedRadioText(artboardStrokeWidthRadios);
        }, artboardStrokeWidthRadios);

        /* ［テキスト］/ Type */
        var textPage = addCategoryPage(dialogControls, "category.text");
        bindCheckbox(dialogControls, textPage, { label: "checkbox.autoSizeAreaText", key: "text/autoSizing", preset: "autoSizeAreaText" });

        /* 最近使用したフォントの表示数（チェックOFFで0＝非表示）/ Recent font count (unchecked means 0 = hidden) */
        var recentFontsRow = addSettingRow(textPage, "checkbox.recentFonts", "tooltip.recentFonts");
        var recentFontsCheckbox = recentFontsRow.add("checkbox", undefined, "");
        recentFontsCheckbox.helpTip = getLabel("tooltip.recentFonts");
        linkLabelToCheckbox(dialogControls, recentFontsRow.settingLabel, recentFontsCheckbox);
        dialogControls.recentFontsCheckbox = recentFontsCheckbox;
        var recentFontsInput = addIntStepperInput(recentFontsRow, "0", RECENT_FONTS_INPUT_CHARS,
            NUMERIC_INPUT_RULES.recentFonts, refreshChangeState);
        recentFontsInput.helpTip = getLabel("tooltip.recentFonts");
        dialogControls.recentFontsInput = recentFontsInput;
        /* チェックボックスの切り替えは onClick で入力欄を書き換えるので、数え直しは main の onClick で行う */
        /* The checkbox's onClick rewrites the field, so main's onClick refreshes the count */
        registerSettingRow(dialogControls, recentFontsRow, function () {
            if (!recentFontsCheckbox.value) return formatOnOff(false);
            return String(clampIntToRule(recentFontsInput.text, NUMERIC_INPUT_RULES.recentFonts));
        }, [recentFontsInput]);

        bindCheckbox(dialogControls, textPage, { label: "checkbox.missingGlyphProtection", key: "text/doFontLocking", preset: "missingGlyphProtection" });
        bindCheckbox(dialogControls, textPage, { label: "checkbox.alternateGlyph", key: "text/enableAlternateGlyph", preset: "alternateGlyph" });
    }

    /**
     * ［ガイド］［スマートガイド］［ユーザーインターフェイス］［パフォーマンス］［ファイル管理］［クリップボードの処理］のページを組み立てる
     * Build the Guides, Smart Guides, User Interface, Performance, File Handling and Clipboard pages
     * @param {object} dialogControls - コントロールを登録するオブジェクト
     * @returns {void}
     */
    function buildBehaviorPages(dialogControls) {
        var refreshChangeState = function () { updateChangeState(dialogControls); };

        /* ガイド（カラー：シアン／ライトブルー、スタイル：ライン／点線）/ Guides (color: Cyan / Light Blue, style: Lines / Dots) */
        var guidesPage = addCategoryPage(dialogControls, "category.guides");
        var guideColorRadios = addRadioRow(dialogControls, guidesPage, "fieldLabel.guideColor",
            ["radio.guideColorCyan", "radio.guideColorLightBlue"]);
        dialogControls.guideColorCyanRadio = guideColorRadios[0];
        dialogControls.guideColorLightBlueRadio = guideColorRadios[1];
        var guideStyleRadios = addRadioRow(dialogControls, guidesPage, "fieldLabel.guideStyle",
            ["radio.guideStyleLines", "radio.guideStyleDots"]);
        dialogControls.guideStyleLinesRadio = guideStyleRadios[0];
        dialogControls.guideStyleDotsRadio = guideStyleRadios[1];

        /* スマートガイド / Smart Guides */
        var smartGuidesPage = addCategoryPage(dialogControls, "category.smartGuides");
        bindCheckbox(dialogControls, smartGuidesPage, { label: "checkbox.objectHighlighting", key: "smartGuides/showObjectHighlighting", preset: "objectHighlighting" });

        /* ［ユーザーインターフェイス］/ User Interface */
        var userInterfacePage = addCategoryPage(dialogControls, "category.userInterface");
        dialogControls.brightnessPicker = addBrightnessPicker(dialogControls, userInterfacePage);
        var canvasColorRadios = addRadioRow(dialogControls, userInterfacePage, "fieldLabel.canvasColor",
            ["radio.canvasMatch", "radio.canvasWhite"]);
        dialogControls.canvasMatchRadio = canvasColorRadios[0];
        dialogControls.canvasWhiteRadio = canvasColorRadios[1];

        /* ［パフォーマンス］/ Performance */
        var performancePage = addCategoryPage(dialogControls, "category.performance");
        bindCheckbox(dialogControls, performancePage, { label: "checkbox.animatedZoom", key: "Performance/AnimZoom", preset: "animatedZoom" });

        var historyStatesRow = addSettingRow(performancePage, "fieldLabel.historyStates", "tooltip.historyStates");
        var historyStatesInput = addIntStepperInput(historyStatesRow, String(NUMERIC_INPUT_RULES.historyStates.defaultValue),
            HISTORY_INPUT_CHARS, NUMERIC_INPUT_RULES.historyStates, refreshChangeState);
        historyStatesInput.helpTip = getLabel("tooltip.historyStates");
        dialogControls.historyStatesInput = historyStatesInput;
        registerSettingRow(dialogControls, historyStatesRow, function () {
            return String(clampIntToRule(historyStatesInput.text, NUMERIC_INPUT_RULES.historyStates));
        }, [historyStatesInput]);

        bindCheckbox(dialogControls, performancePage, { label: "checkbox.realTimeDrawing", key: "LiveEdit_State_Machine", preset: "realTimeDrawing" });

        /* ［ファイル管理］/ File Handling */
        var fileManagementPage = addCategoryPage(dialogControls, "category.fileManagement");
        bindCheckbox(dialogControls, fileManagementPage, { label: "checkbox.editOriginalSystemDefault", key: "useSysDefEdit", preset: "editOriginalSystemDefault" });
        bindCheckbox(dialogControls, fileManagementPage, { label: "checkbox.autoActivateFonts", key: "AutoActivateMissingFont", preset: "autoActivateFonts" });

        /* ファイルの保存先（コンピューター／クラウド）/ Save location (Computer / Cloud) */
        var saveLocationRadios = addRadioRow(dialogControls, fileManagementPage, "fieldLabel.saveLocation",
            ["radio.saveToComputer", "radio.saveToCloud"]);
        dialogControls.saveToComputerRadio = saveLocationRadios[0];
        dialogControls.saveToCloudRadio = saveLocationRadios[1];

        /* リンクを更新（自動／手動／確認）。並びは UPDATE_LINKS_* の値と同じ / Update Links; order matches the UPDATE_LINKS_* values */
        dialogControls.updateLinksRadios = addRadioRow(dialogControls, fileManagementPage, "fieldLabel.updateLinks",
            ["radio.updateLinksAuto", "radio.updateLinksManual", "radio.updateLinksAsk"]);

        /* クリップボードの処理 / Clipboard Handling */
        var clipboardPage = addCategoryPage(dialogControls, "category.clipboard");
        bindCheckbox(dialogControls, clipboardPage, { label: "checkbox.includeSvgCode", key: "plugin/FileClipboard/copySVGCode", preset: "includeSvgCode" });
    }

    // =========================================
    // 2カラム表示（従来のレイアウト） / Two-column view (the classic layout)
    // =========================================

    /**
     * ラベル付きチェックボックスを追加 / Add a labeled checkbox
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelPath - LABELS のドット区切りパス
     * @param {string} [tooltipPath] - ツールチップ用のパス。省略時はラベルをそのまま使う
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addClassicCheckbox(parentContainer, labelPath, tooltipPath) {
        var newCheckbox = parentContainer.add("checkbox", undefined, getLabel(labelPath));
        newCheckbox.helpTip = getLabel(tooltipPath || labelPath);
        return newCheckbox;
    }

    /**
     * 見出し付きパネルを追加 / Add a titled panel with the shared layout
     * @param {Group} parentColumn - 追加先のカラム
     * @param {string} labelPath - LABELS のドット区切りパス
     * @returns {Panel} 追加したパネル
     */
    function addClassicPanel(parentColumn, labelPath) {
        var newPanel = parentColumn.add("panel", undefined, getLabel(labelPath));
        setupPanel(newPanel, 6);
        return newPanel;
    }

    /**
     * パネルを縦に積むカラムを追加 / Add a column that stacks panels vertically
     * @param {Group} parentContainer - 追加先のグループ
     * @returns {Group} 追加したカラムグループ
     */
    function addClassicColumn(parentContainer) {
        var columnGroup = parentContainer.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        return columnGroup;
    }

    /**
     * ラベル＋コントロールを並べる行を追加 / Add a row that starts with a static label
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelPath - LABELS のドット区切りパス
     * @param {number} [labelWidth] - ラベル幅。指定すると右揃えで幅を固定
     * @returns {Group} 追加した行グループ（続けてコントロールを add する）
     */
    function addClassicLabeledRow(parentContainer, labelPath, labelWidth) {
        var rowGroup = parentContainer.add("group");
        setupRow(rowGroup, "left", CLASSIC_ROW_SPACING);
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelPath));
        if (typeof labelWidth === "number") {
            rowLabel.preferredSize = [labelWidth, -1]; /* 高さは自動 / height stays automatic */
            rowLabel.justify = "right";
        }
        return rowGroup;
    }

    /**
     * ラベル付きの行にラジオボタンを並べる / Add a labeled row of radio buttons
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelPath - 行ラベルのパス
     * @param {string[]} radioLabelPaths - ラジオのラベルのパス（並び順）
     * @param {number} [labelWidth] - ラベル幅（addClassicLabeledRow と同じ）
     * @returns {RadioButton[]} 追加したラジオボタン
     */
    function addClassicRadioRow(parentContainer, labelPath, radioLabelPaths, labelWidth) {
        var radioRow = addClassicLabeledRow(parentContainer, labelPath, labelWidth);
        var radioButtons = [];
        for (var i = 0; i < radioLabelPaths.length; i++) {
            radioButtons.push(radioRow.add("radiobutton", undefined, getLabel(radioLabelPaths[i])));
        }
        return radioButtons;
    }

    /**
     * チェックボックスを追加し、対応表に登録する / Add a checkbox and register it in the binding table
     * 対応表（dialogControls.checkboxBindings）は読み込み・保存・プリセット適用の3か所が共通で参照する
     * The binding table is shared by load, save and preset apply
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {Panel|Group} parentPanel - 追加先
     * @param {{label: string, key: string, preset: string, type: string, tooltip: string}} checkboxSpec
     *        label=ラベルのパス / key=環境設定キー / preset=プリセット項目名 /
     *        type=省略時 "bool"、0=ON・1=OFF の反転キーは "invertedInt" / tooltip=ツールチップのパス（省略可）
     * @returns {Checkbox} 追加したチェックボックス
     */
    function bindClassicCheckbox(dialogControls, parentPanel, checkboxSpec) {
        var newCheckbox = addClassicCheckbox(parentPanel, checkboxSpec.label, checkboxSpec.tooltip);
        dialogControls.checkboxBindings.push({
            key: checkboxSpec.key,
            valueType: checkboxSpec.type || "bool",
            control: newCheckbox,
            presetField: checkboxSpec.preset
        });
        return newCheckbox;
    }

    /**
     * 左カラムのパネル（一般・選択範囲・アートボード・テキスト・UI）を組み立てる
     * Build the left-column panels (General, Selection, Artboard, Text, User Interface)
     * @param {object} dialogControls - コントロールを登録するオブジェクト
     * @param {Group} leftColumnGroup - 左カラム
     * @returns {void}
     */
    function buildClassicLeftColumn(dialogControls, leftColumnGroup) {
        /* ［一般］/ General */
        var generalPanel = addClassicPanel(leftColumnGroup, "classicPanel.general");
        bindClassicCheckbox(dialogControls, generalPanel, { label: "checkbox.richToolTips", key: "showRichToolTips", preset: "richToolTips" });
        bindClassicCheckbox(dialogControls, generalPanel, { label: "checkbox.homeScreen", key: "Hello/ShowHomeScreenWS", preset: "homeScreen", tooltip: "tooltip.homeScreen" });
        bindClassicCheckbox(dialogControls, generalPanel, { label: "checkbox.legacyNewDoc", key: "Hello/NewDoc", preset: "legacyNewDoc", tooltip: "tooltip.legacyNewDoc" });
        bindClassicCheckbox(dialogControls, generalPanel, { label: "checkbox.printBleedWidget", key: "enablePrintBleedWidget", preset: "printBleedWidget" });

        /* ［選択範囲・アンカー表示］/ Selection & Anchor Display */
        var selectionAnchorPanel = addClassicPanel(leftColumnGroup, "classicPanel.selectionAnchor");
        bindClassicCheckbox(dialogControls, selectionAnchorPanel, { label: "checkbox.zoomToSelection", key: "zoomToSelection", preset: "zoomToSelection" });
        bindClassicCheckbox(dialogControls, selectionAnchorPanel, { label: "checkbox.unlockOnCanvas", key: "showLockIcon", preset: "unlockOnCanvas", tooltip: "tooltip.unlockOnCanvas" });

        /* アンカーポイントのサイズ（4段階スライダー）/ Anchor point size (four-step slider) */
        var anchorSizeRow = addClassicLabeledRow(selectionAnchorPanel, "fieldLabel.anchorSize");
        var anchorSizeSlider = anchorSizeRow.add("slider", undefined, 1, 1, ANCHOR_SIZE_LEVELS.length);
        anchorSizeSlider.preferredSize = [CLASSIC_ANCHOR_SLIDER_WIDTH, -1];
        anchorSizeSlider.helpTip = getLabel("tooltip.anchorSize");
        dialogControls.anchorSizeSlider = anchorSizeSlider;

        /* アートボード / Artboard */
        var artboardPanel = addClassicPanel(leftColumnGroup, "classicPanel.artboard");
        bindClassicCheckbox(dialogControls, artboardPanel, { label: "checkbox.moveLockedArt", key: "moveLockedAndHiddenArt", preset: "moveLockedArt" });
        bindClassicCheckbox(dialogControls, artboardPanel, { label: "checkbox.showArtboardName", key: "showArtboardLabelOnCanvas", preset: "showArtboardName" });

        /* ハイライトのカラー（9色のプリセットから選択）/ Highlight color (nine presets) */
        var artboardColorRow = addClassicLabeledRow(artboardPanel, "fieldLabel.artboardColor");
        var artboardColorLabels = [];
        for (var i = 0; i < ARTBOARD_COLOR_PRESETS.length; i++) {
            artboardColorLabels.push(ARTBOARD_COLOR_PRESETS[i].label);
        }
        dialogControls.artboardColorDropdown = artboardColorRow.add("dropdownlist", undefined, artboardColorLabels);

        /* ストロークの幅（1〜4）/ Stroke width (1-4) */
        var artboardStrokeWidthRow = addClassicLabeledRow(artboardPanel, "fieldLabel.artboardStrokeWidth");
        dialogControls.artboardStrokeWidthRadios = [];
        for (var j = 0; j < ARTBOARD_STROKE_WIDTHS.length; j++) {
            dialogControls.artboardStrokeWidthRadios.push(
                artboardStrokeWidthRow.add("radiobutton", undefined, String(ARTBOARD_STROKE_WIDTHS[j]))
            );
        }

        /* ［テキスト］/ Text */
        var textPanel = addClassicPanel(leftColumnGroup, "classicPanel.text");
        bindClassicCheckbox(dialogControls, textPanel, { label: "checkbox.autoSizeAreaText", key: "text/autoSizing", preset: "autoSizeAreaText" });

        /* 最近使用したフォントの表示数（チェックOFFで0＝非表示）/ Recent font count (unchecked means 0 = hidden) */
        var recentFontsRow = textPanel.add("group");
        setupRow(recentFontsRow, "left", CLASSIC_ROW_SPACING);
        dialogControls.recentFontsCheckbox = addClassicCheckbox(recentFontsRow, "checkbox.recentFonts", "tooltip.recentFonts");
        var recentFontsInput = addIntStepperInput(recentFontsRow, "0", RECENT_FONTS_INPUT_CHARS, NUMERIC_INPUT_RULES.recentFonts);
        recentFontsInput.helpTip = getLabel("tooltip.recentFonts");
        dialogControls.recentFontsInput = recentFontsInput;

        bindClassicCheckbox(dialogControls, textPanel, { label: "checkbox.missingGlyphProtection", key: "text/doFontLocking", preset: "missingGlyphProtection" });
        bindClassicCheckbox(dialogControls, textPanel, { label: "checkbox.alternateGlyph", key: "text/enableAlternateGlyph", preset: "alternateGlyph" });

        /* ［ユーザーインターフェイス］/ User Interface */
        var userInterfacePanel = addClassicPanel(leftColumnGroup, "classicPanel.userInterface");
        dialogControls.brightnessPicker = addBrightnessPicker(dialogControls, userInterfacePanel, true);
        var canvasColorRadios = addClassicRadioRow(userInterfacePanel, "fieldLabel.canvasColor", ["radio.canvasMatch", "radio.canvasWhite"]);
        dialogControls.canvasMatchRadio = canvasColorRadios[0];
        dialogControls.canvasWhiteRadio = canvasColorRadios[1];
    }

    /**
     * 右カラムのパネル（ガイド・スマートガイド・パフォーマンス・ファイル管理・クリップボード・パスに制限）を組み立てる
     * Build the right-column panels (Guides, Smart Guides, Performance, File Management, Clipboard, Limit to Path)
     * @param {object} dialogControls - コントロールを登録するオブジェクト
     * @param {Group} rightColumnGroup - 右カラム
     * @returns {void}
     */
    function buildClassicRightColumn(dialogControls, rightColumnGroup) {
        /* ガイド（カラー：シアン／ライトブルー、スタイル：ライン／点線）/ Guides (color: Cyan / Light Blue, style: Lines / Dots) */
        var guidesPanel = addClassicPanel(rightColumnGroup, "classicPanel.guides");
        var guideColorRadios = addClassicRadioRow(guidesPanel, "fieldLabel.guideColor",
            ["radio.guideColorCyan", "radio.guideColorLightBlue"], CLASSIC_GUIDE_LABEL_WIDTH);
        dialogControls.guideColorCyanRadio = guideColorRadios[0];
        dialogControls.guideColorLightBlueRadio = guideColorRadios[1];
        var guideStyleRadios = addClassicRadioRow(guidesPanel, "fieldLabel.guideStyle",
            ["radio.guideStyleLines", "radio.guideStyleDots"], CLASSIC_GUIDE_LABEL_WIDTH);
        dialogControls.guideStyleLinesRadio = guideStyleRadios[0];
        dialogControls.guideStyleDotsRadio = guideStyleRadios[1];

        /* スマートガイド / Smart Guides */
        var smartGuidesPanel = addClassicPanel(rightColumnGroup, "classicPanel.smartGuides");
        bindClassicCheckbox(dialogControls, smartGuidesPanel, { label: "checkbox.objectHighlighting", key: "smartGuides/showObjectHighlighting", preset: "objectHighlighting" });

        /* ［パフォーマンス］/ Performance */
        var performancePanel = addClassicPanel(rightColumnGroup, "classicPanel.performance");
        bindClassicCheckbox(dialogControls, performancePanel, { label: "checkbox.animatedZoom", key: "Performance/AnimZoom", preset: "animatedZoom" });

        var historyStatesRow = addClassicLabeledRow(performancePanel, "fieldLabel.historyStates");
        var historyStatesInput = addIntStepperInput(historyStatesRow, String(NUMERIC_INPUT_RULES.historyStates.defaultValue),
            HISTORY_INPUT_CHARS, NUMERIC_INPUT_RULES.historyStates);
        historyStatesInput.helpTip = getLabel("tooltip.historyStates");
        dialogControls.historyStatesInput = historyStatesInput;

        bindClassicCheckbox(dialogControls, performancePanel, { label: "checkbox.realTimeDrawing", key: "LiveEdit_State_Machine", preset: "realTimeDrawing" });

        /* ［ファイル管理］/ File Management */
        var fileManagementPanel = addClassicPanel(rightColumnGroup, "classicPanel.fileManagement");
        bindClassicCheckbox(dialogControls, fileManagementPanel, { label: "checkbox.editOriginalSystemDefault", key: "useSysDefEdit", preset: "editOriginalSystemDefault" });
        bindClassicCheckbox(dialogControls, fileManagementPanel, { label: "checkbox.autoActivateFonts", key: "AutoActivateMissingFont", preset: "autoActivateFonts" });

        /* ファイルの保存先（コンピューター／クラウド）/ Save location (Computer / Cloud) */
        var saveLocationRadios = addClassicRadioRow(fileManagementPanel, "fieldLabel.saveLocation", ["radio.saveToComputer", "radio.saveToCloud"]);
        dialogControls.saveToComputerRadio = saveLocationRadios[0];
        dialogControls.saveToCloudRadio = saveLocationRadios[1];

        /* リンクを更新（自動／手動／確認）。並びは UPDATE_LINKS_* の値と同じ / Update Links; order matches the UPDATE_LINKS_* values */
        dialogControls.updateLinksRadios = addClassicRadioRow(fileManagementPanel, "fieldLabel.updateLinks",
            ["radio.updateLinksAuto", "radio.updateLinksManual", "radio.updateLinksAsk"]);

        /* クリップボードの処理 / Clipboard Handling */
        var clipboardPanel = addClassicPanel(rightColumnGroup, "classicPanel.clipboard");
        bindClassicCheckbox(dialogControls, clipboardPanel, { label: "checkbox.includeSvgCode", key: "plugin/FileClipboard/copySVGCode", preset: "includeSvgCode" });

        /* パスに制限。0=ON / 1=OFF の反転キー / Limit to Path; these keys store 0 = ON and 1 = OFF */
        var limitToPathPanel = addClassicPanel(rightColumnGroup, "classicPanel.limitToPath");
        bindClassicCheckbox(dialogControls, limitToPathPanel, { label: "checkbox.objectPathOnly", key: "hitShapeOnPreview", preset: "objectPathOnly", type: "invertedInt" });
        bindClassicCheckbox(dialogControls, limitToPathPanel, { label: "checkbox.textPathOnly", key: "hitTypeShapeOnPreview", preset: "textPathOnly", type: "invertedInt" });
    }

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
    // UIの値の読み書き / UI value accessors
    // =========================================

    /**
     * ラジオの並びから選択中の index を返す / Index of the selected radio in a list
     * @param {RadioButton[]} radioButtons - ラジオボタンの配列
     * @param {number} fallbackIndex - 未選択のときの index
     * @returns {number} 選択中の index
     */
    function getSelectedRadioIndex(radioButtons, fallbackIndex) {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) return i;
        }
        return fallbackIndex;
    }

    /**
     * 指定 index のラジオだけを選ぶ / Select only the radio at the given index
     * @param {RadioButton[]} radioButtons - ラジオボタンの配列
     * @param {number} selectedIndex - 選ぶ index
     * @returns {void}
     */
    function selectRadioByIndex(radioButtons, selectedIndex) {
        for (var i = 0; i < radioButtons.length; i++) {
            radioButtons[i].value = (i === selectedIndex);
        }
    }

    /**
     * ストローク幅のラジオを選択（最も近い選択肢に丸める）
     * Select the stroke-width radio, snapping to the nearest available width
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {number} strokeWidth - ストローク幅
     * @returns {void}
     */
    function selectArtboardStrokeWidth(dialogControls, strokeWidth) {
        selectRadioByIndex(dialogControls.artboardStrokeWidthRadios, findNearestIndex(ARTBOARD_STROKE_WIDTHS, strokeWidth));
    }

    /**
     * アンカーポイントサイズのスライダーを指定段階に合わせる / Move the anchor size slider to a step
     * @param {Slider} anchorSizeSlider - 対象のスライダー
     * @param {number} anchorStep - 1〜ANCHOR_SIZE_LEVELS.length の段階番号
     * @returns {void}
     */
    function applyAnchorSizeStep(anchorSizeSlider, anchorStep) {
        anchorSizeSlider.value = clamp(Math.round(anchorStep), 1, ANCHOR_SIZE_LEVELS.length);
    }

    /**
     * anchorSizePref の値から最も近い段階にスナップ / Snap the slider to the step nearest a stored value
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {number} anchorSizeValue - anchorSizePref の値
     * @returns {void}
     */
    function selectAnchorSize(dialogControls, anchorSizeValue) {
        applyAnchorSizeStep(dialogControls.anchorSizeSlider, findNearestIndex(ANCHOR_SIZE_LEVELS, anchorSizeValue) + 1);
    }

    /**
     * 選択中の段階に対応する anchorSizePref の値を返す
     * Return the anchorSizePref value for the selected step
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {number} ANCHOR_SIZE_LEVELS のいずれかの値
     */
    function getSelectedAnchorSize(dialogControls) {
        var selectedStep = clamp(Math.round(dialogControls.anchorSizeSlider.value), 1, ANCHOR_SIZE_LEVELS.length);
        return ANCHOR_SIZE_LEVELS[selectedStep - 1];
    }

    /**
     * 「リンクを更新」のラジオを選択。未知の値は「確認」/ Select the "Update Links" radio; unknown values fall back to Ask
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {number} updateLinksValue - UPDATE_LINKS_* のいずれか
     * @returns {void}
     */
    function selectUpdateLinks(dialogControls, updateLinksValue) {
        var isKnownValue = (updateLinksValue === UPDATE_LINKS_AUTO || updateLinksValue === UPDATE_LINKS_MANUAL);
        selectRadioByIndex(dialogControls.updateLinksRadios, isKnownValue ? updateLinksValue : UPDATE_LINKS_ASK);
    }

    /**
     * 「最近使用したフォント」のチェックと入力欄を同期
     * Sync the recent-fonts checkbox and its input
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {number} recentFontCount - 表示数（0以下はチェックOFF＝非表示）
     * @returns {void}
     */
    function applyRecentFontsCount(dialogControls, recentFontCount) {
        var isListVisible = (recentFontCount > 0);
        dialogControls.recentFontsCheckbox.value = isListVisible;
        /* ONのときだけルール（1〜30）で補正。OFFは0を表示 / Normalize only when enabled; show 0 when off */
        dialogControls.recentFontsInput.text = String(isListVisible ? clampIntToRule(recentFontCount, NUMERIC_INPUT_RULES.recentFonts) : 0);
        dialogControls.recentFontsInput.enabled = isListVisible;
        /* ∧∨も合わせて切り替え、自作描画のディム表示を描き直す / toggle and redraw the custom-drawn stepper too */
        dialogControls.recentFontsInput.stepperGroup.enabled = isListVisible;
        redrawSteppersIn(dialogControls.recentFontsInput.stepperGroup);
    }

    // =========================================
    // 変更の追跡とカテゴリの切り替え / Change tracking & category switching
    // =========================================

    /**
     * ［デフォルト］の設定一式をいったんUIに入れ、各行の表示をデフォルトの列に書き込む。
     * UIは書き換わるので、このあと環境設定を読み込み直す
     * Put the [Default] set into the UI, copy each row's text into the default column; reload the preferences afterwards
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function fillDefaultColumn(dialogControls) {
        applyPresetStateToUI(dialogControls, PRESET_STATE_DEFAULT);
        /* 明るさはプリセットに無いので、Illustrator の初期値（やや暗）を入れる / brightness is not in the presets; use Illustrator's initial Medium Dark */
        dialogControls.brightnessPicker.selectedIndex = DEFAULT_BRIGHTNESS_INDEX;
        for (var i = 0; i < dialogControls.settingRows.length; i++) {
            var settingRowState = dialogControls.settingRows[i];
            settingRowState.valueGroup.defaultLabel.text = settingRowState.describeValue();
        }
    }

    /**
     * 文字の表示幅を測る / Measure the display width of a text
     * @param {StaticText|ListBox} control - 測るのに使うコントロール（そのフォントで測る）
     * @param {string} text - 測る文字列
     * @returns {number} 幅（px）
     */
    function measureTextWidth(control, text) {
        return control.graphics.measureString(text)[0];
    }

    /**
     * 項目名とデフォルトの列を、全カテゴリでいちばん長い文字に合わせた幅にそろえる（デフォルトの列を埋めたあとに呼ぶ）
     * Fit the setting and default columns to the longest text across every category (call after filling the default column)
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function fitColumnWidths(dialogControls) {
        var settingColumn = [];
        var defaultColumn = [];
        for (var i = 0; i < dialogControls.categoryPages.length; i++) {
            var categoryPageState = dialogControls.categoryPages[i];
            settingColumn.push(categoryPageState.headerLabels[0]);
            defaultColumn.push(categoryPageState.headerLabels[1]);
            for (var j = 0; j < categoryPageState.rows.length; j++) {
                settingColumn.push(categoryPageState.rows[j].valueGroup.settingLabel);
                defaultColumn.push(categoryPageState.rows[j].valueGroup.defaultLabel);
            }
        }
        var columns = [settingColumn, defaultColumn];
        for (var k = 0; k < columns.length; k++) {
            var columnWidth = 0;
            for (var m = 0; m < columns[k].length; m++) {
                columnWidth = Math.max(columnWidth, measureTextWidth(columns[k][m], columns[k][m].text));
            }
            for (var n = 0; n < columns[k].length; n++) {
                columns[k][n].preferredSize = [columnWidth + COLUMN_TEXT_PADDING, -1]; /* 高さは自動 / height stays automatic */
            }
        }
    }

    /**
     * カテゴリ一覧の幅を、件数の付いた名前がいちばん長くなる場合に合わせる / Fit the category list to its longest name with a change count
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {number} 幅（px）
     */
    function measureCategoryListWidth(dialogControls) {
        var listWidth = 0;
        for (var i = 0; i < dialogControls.categoryPages.length; i++) {
            var longestName = formatCategoryName(dialogControls.categoryPages[i].labelPath, 99);
            listWidth = Math.max(listWidth, measureTextWidth(dialogControls.categoryList, longestName));
        }
        return listWidth + SIDEBAR_TEXT_PADDING;
    }

    /**
     * 今のUIの状態を、変更を数える基準として控える（環境設定を読み込んだ直後に呼ぶ）
     * Record the UI state as the baseline for counting changes (call right after loading the preferences)
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function captureBaseline(dialogControls) {
        for (var i = 0; i < dialogControls.settingRows.length; i++) {
            var settingRowState = dialogControls.settingRows[i];
            settingRowState.baselineText = settingRowState.describeValue();
        }
        updateChangeState(dialogControls);
    }

    /**
     * 変更した行に印を付け、カテゴリごとの件数・合計件数・［すべての変更を取り消す］の有効／無効を更新する
     * Mark changed rows and refresh the per-category counts, the total and the Revert button
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function updateChangeState(dialogControls) {
        /* 現在の値を控える前（組み立て中）は数えない / skip until the baseline exists (during construction) */
        if (!dialogControls.changeCountLabel) return;
        var totalChangeCount = 0;
        for (var i = 0; i < dialogControls.categoryPages.length; i++) {
            var categoryPageState = dialogControls.categoryPages[i];
            var pageChangeCount = 0;
            for (var j = 0; j < categoryPageState.rows.length; j++) {
                var settingRowState = categoryPageState.rows[j];
                var isChanged = (settingRowState.baselineText !== null && settingRowState.describeValue() !== settingRowState.baselineText);
                settingRowState.valueGroup.changeMarker.visible = isChanged;
                if (isChanged) pageChangeCount++;
            }
            dialogControls.categoryList.items[i].text = formatCategoryName(categoryPageState.labelPath, pageChangeCount);
            totalChangeCount += pageChangeCount;
        }
        dialogControls.changeCountLabel.text = getLabel("status.changeCount", [totalChangeCount]);
        dialogControls.btnRevertAll.enabled = (totalChangeCount > 0);
    }

    /**
     * カテゴリ一覧の表示名（変更があれば件数を添える）/ Category list text, with the change count when there are changes
     * @param {string} labelPath - カテゴリ名のパス
     * @param {number} changeCount - そのカテゴリで変更した項目の数
     * @returns {string} 表示名
     */
    function formatCategoryName(labelPath, changeCount) {
        if (changeCount === 0) return getLabel(labelPath);
        return getLabel("category.withChangeCount", [getLabel(labelPath), changeCount]);
    }

    /**
     * 指定したカテゴリのページだけを表示する / Show only the page of the given category
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {number} pageIndex - categoryPages の index
     * @returns {void}
     */
    function showCategoryPage(dialogControls, pageIndex) {
        for (var i = 0; i < dialogControls.categoryPages.length; i++) {
            dialogControls.categoryPages[i].page.visible = (i === pageIndex);
        }
    }

    // =========================================
    // 読み込みとプリセット適用 / Load & preset apply
    // =========================================

    /**
     * 対応表のチェックボックスを環境設定から復元 / Restore the table-driven checkboxes from the preferences
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function loadBindingsIntoUI(dialogControls) {
        for (var i = 0; i < dialogControls.checkboxBindings.length; i++) {
            var checkboxBinding = dialogControls.checkboxBindings[i];
            if (checkboxBinding.valueType === "bool") {
                checkboxBinding.control.value = !!safeGetBool(checkboxBinding.key, false);
            } else if (checkboxBinding.valueType === "invertedInt") {
                checkboxBinding.control.value = (safeGetInt(checkboxBinding.key, 1) === 0); /* 0がON / 0 means ON */
            }
        }
    }

    /**
     * 数値入力欄を環境設定から復元 / Restore the numeric fields from the preferences
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function loadNumericInputsIntoUI(dialogControls) {
        /* ヒストリー数は 0 を取り得ないので、0＝キー未登録とみなして既定値に戻す */
        /* History states can never be 0, so treat 0 as "key not present" and fall back to the default */
        dialogControls.historyStatesInput.text = String(clampIntToRule(
            safeGetPositiveInt("maximumUndoDepth", NUMERIC_INPUT_RULES.historyStates.defaultValue),
            NUMERIC_INPUT_RULES.historyStates
        ));
        /* 0 は「一覧を非表示」なのでそのまま渡す / 0 means "hide the list", so pass it through */
        applyRecentFontsCount(dialogControls, safeGetInt("text/recentFontMenu/showNEntries", 0));
    }

    /**
     * ドロップダウン・ラジオを環境設定から復元 / Restore the dropdowns and radio groups from the preferences
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function loadChoicesIntoUI(dialogControls) {
        /* アートボードのハイライト（Real値なので個別に読み込み）/ Artboard highlight (real values, read individually) */
        dialogControls.artboardColorDropdown.selection = findNearestArtboardColorIndex(
            safeGetReal("ArtboardBBColorRed", 0.0),
            safeGetReal("ArtboardBBColorGreen", 0.0),
            safeGetReal("ArtboardBBColorBlue", 0.0)
        );
        selectArtboardStrokeWidth(dialogControls, safeGetReal("ArtboardBBWidth", 1.0));

        /* 選択範囲・アンカー表示（キー未登録の 0 は最小値の段階に落ちる）/ Selection & anchor display (a missing key reads 0 and snaps to the smallest step) */
        selectAnchorSize(dialogControls, safeGetInt("anchorSizePref", ANCHOR_SIZE_LEVELS[0]));

        /* ユーザーインターフェイス / User interface */
        setRadioPair(dialogControls.canvasWhiteRadio, dialogControls.canvasMatchRadio, safeGetInt("uiCanvasIsWhite", 0) === 1);
        selectBrightnessByValue(dialogControls.brightnessPicker, safeGetReal("uiBrightness", 0.0));

        /* ガイド / Guides */
        setRadioPair(dialogControls.guideColorLightBlueRadio, dialogControls.guideColorCyanRadio, isGuideColorLightBlue());
        setRadioPair(dialogControls.guideStyleDotsRadio, dialogControls.guideStyleLinesRadio,
            safeGetInt("Guide/Style", GUIDE_STYLE_LINES) === GUIDE_STYLE_DOTS);

        /* ファイル管理 / File management */
        setRadioPair(dialogControls.saveToCloudRadio, dialogControls.saveToComputerRadio,
            safeGetBool("AdobeSaveAsCloudDocumentPreference", false));
        selectUpdateLinks(dialogControls, safeGetInt("plugin/FileClipboard/linkoptions", UPDATE_LINKS_ASK));
    }

    /**
     * 現在の環境設定を読み込んでUIに反映 / Load the current preferences into the UI
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function loadPreferencesIntoUI(dialogControls) {
        loadBindingsIntoUI(dialogControls);
        loadNumericInputsIntoUI(dialogControls);
        loadChoicesIntoUI(dialogControls);
    }

    /**
     * プリセットの設定一式をUIに反映（環境設定には書き込まない）
     * Load a preset into the UI (nothing is written to preferences)
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {object} presetState - PRESET_STATE_DEFAULT / PRESET_STATE_1 と同じ形の設定一式
     * @returns {void}
     */
    function applyPresetStateToUI(dialogControls, presetState) {
        /* チェックボックスは対応表の presetField 経由 / Checkboxes come from the table's presetField */
        for (var i = 0; i < dialogControls.checkboxBindings.length; i++) {
            var checkboxBinding = dialogControls.checkboxBindings[i];
            checkboxBinding.control.value = !!presetState[checkboxBinding.presetField];
        }

        /* 対応表に載らない項目 / Fields the table doesn't cover */
        dialogControls.artboardColorDropdown.selection = presetState.artboardColorIndex;
        selectArtboardStrokeWidth(dialogControls, presetState.artboardStrokeWidth);
        selectAnchorSize(dialogControls, presetState.anchorSize);
        applyRecentFontsCount(dialogControls, presetState.recentFontsEnabled ? presetState.recentFontsCount : 0);
        dialogControls.historyStatesInput.text = String(presetState.historyStates);
        setRadioPair(dialogControls.canvasWhiteRadio, dialogControls.canvasMatchRadio, presetState.canvasWhite);
        setRadioPair(dialogControls.guideColorLightBlueRadio, dialogControls.guideColorCyanRadio, presetState.guideColorIsLightBlue);
        setRadioPair(dialogControls.guideStyleDotsRadio, dialogControls.guideStyleLinesRadio, presetState.guideStyleIsDots);
        setRadioPair(dialogControls.saveToCloudRadio, dialogControls.saveToComputerRadio, presetState.saveToCloud);
        selectUpdateLinks(dialogControls, presetState.updateLinks);
    }

    /**
     * ドロップダウンで選んだ項目をUIに反映 / Apply the chosen preset dropdown entry to the UI
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @param {number} choiceIndex - PRESET_CHOICES の index
     * @returns {void}
     */
    function applyPresetChoice(dialogControls, choiceIndex) {
        var presetChoice = PRESET_CHOICES[choiceIndex] || PRESET_CHOICES[0];
        if (presetChoice.state) {
            applyPresetStateToUI(dialogControls, presetChoice.state);
        } else {
            loadPreferencesIntoUI(dialogControls);
        }
    }

    /**
     * UIの状態を、プリセットと同じ形の設定一式として読み取る（レイアウトを切り替えるときに引き継ぐ）
     * Read the UI as a set shaped like the presets (carried over when switching layouts)
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {object} PRESET_STATE_DEFAULT と同じ形の設定一式
     */
    function readPresetStateFromUI(dialogControls) {
        var presetState = {};
        for (var i = 0; i < dialogControls.checkboxBindings.length; i++) {
            var checkboxBinding = dialogControls.checkboxBindings[i];
            presetState[checkboxBinding.presetField] = !!checkboxBinding.control.value;
        }
        var artboardColorSelection = dialogControls.artboardColorDropdown.selection;
        presetState.artboardColorIndex = artboardColorSelection ? artboardColorSelection.index : 0;
        presetState.artboardStrokeWidth = ARTBOARD_STROKE_WIDTHS[getSelectedRadioIndex(dialogControls.artboardStrokeWidthRadios, 0)];
        presetState.anchorSize = getSelectedAnchorSize(dialogControls);
        presetState.recentFontsEnabled = !!dialogControls.recentFontsCheckbox.value;
        presetState.recentFontsCount = clampIntToRule(dialogControls.recentFontsInput.text, NUMERIC_INPUT_RULES.recentFonts);
        presetState.historyStates = clampIntToRule(dialogControls.historyStatesInput.text, NUMERIC_INPUT_RULES.historyStates);
        presetState.canvasWhite = !!dialogControls.canvasWhiteRadio.value;
        presetState.guideColorIsLightBlue = !!dialogControls.guideColorLightBlueRadio.value;
        presetState.guideStyleIsDots = !!dialogControls.guideStyleDotsRadio.value;
        presetState.saveToCloud = !!dialogControls.saveToCloudRadio.value;
        presetState.updateLinks = getSelectedRadioIndex(dialogControls.updateLinksRadios, UPDATE_LINKS_ASK);
        return presetState;
    }

    // =========================================
    // 保存 / Save
    // =========================================

    /**
     * 対応表（checkboxBindings）のキーをまとめて書き込む
     * Write every key listed in the checkbox binding table
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function savePreferenceBindings(dialogControls) {
        for (var i = 0; i < dialogControls.checkboxBindings.length; i++) {
            var checkboxBinding = dialogControls.checkboxBindings[i];
            if (checkboxBinding.valueType === "bool") {
                safeSetBool(checkboxBinding.key, !!checkboxBinding.control.value);
            } else if (checkboxBinding.valueType === "invertedInt") {
                safeSetInt(checkboxBinding.key, checkboxBinding.control.value ? 0 : 1); /* 0がON / 0 means ON */
            }
        }
    }

    /**
     * ［OK］時に全項目を環境設定へ書き込む / Write every setting to the preferences on [OK]
     * @param {object} dialogControls - コントロールをまとめたオブジェクト
     * @returns {void}
     */
    function savePreferences(dialogControls) {
        /* 対応表で扱えるキー（チェックボックス）/ Keys covered by the binding table (checkboxes) */
        savePreferenceBindings(dialogControls);

        /* テキスト / Text */
        safeSetInt("text/recentFontMenu/showNEntries", dialogControls.recentFontsCheckbox.value
            ? normalizeIntInput(dialogControls.recentFontsInput, NUMERIC_INPUT_RULES.recentFonts) : 0);

        /* 選択範囲・アンカー表示 / Selection & anchor display */
        safeSetInt("anchorSizePref", getSelectedAnchorSize(dialogControls));

        /* ユーザーインターフェイス / User interface */
        safeSetInt("uiCanvasIsWhite", dialogControls.canvasWhiteRadio.value ? 1 : 0);

        /* ガイド / Guides */
        writeGuideColor(dialogControls.guideColorLightBlueRadio.value ? GUIDE_COLOR_LIGHT_BLUE : GUIDE_COLOR_CYAN);
        safeSetInt("Guide/Style", dialogControls.guideStyleDotsRadio.value ? GUIDE_STYLE_DOTS : GUIDE_STYLE_LINES);

        /* パフォーマンス / Performance */
        safeSetInt("maximumUndoDepth", normalizeIntInput(dialogControls.historyStatesInput, NUMERIC_INPUT_RULES.historyStates));

        /* ファイル管理・クリップボード / File management & clipboard */
        safeSetBool("AdobeSaveAsCloudDocumentPreference", !!dialogControls.saveToCloudRadio.value);
        safeSetInt("plugin/FileClipboard/linkoptions", getSelectedRadioIndex(dialogControls.updateLinksRadios, UPDATE_LINKS_ASK));

        /* アートボード / Artboard */
        var artboardColorSelection = dialogControls.artboardColorDropdown.selection;
        var selectedArtboardColor = ARTBOARD_COLOR_PRESETS[artboardColorSelection ? artboardColorSelection.index : 0];
        safeSetReal("ArtboardBBColorRed", selectedArtboardColor.red);
        safeSetReal("ArtboardBBColorGreen", selectedArtboardColor.green);
        safeSetReal("ArtboardBBColorBlue", selectedArtboardColor.blue);
        safeSetReal("ArtboardBBWidth", ARTBOARD_STROKE_WIDTHS[getSelectedRadioIndex(dialogControls.artboardStrokeWidthRadios, 0)]);
    }

    /**
     * 環境設定の変更後に画面を強制再描画（redraw だけでは反映されないため）
     * Force a redraw after preference changes (redraw alone is unreliable)
     * ドキュメントが開いていないときは何もしない / Does nothing when no document is open
     * @returns {void}
     */
    function forceScreenRefresh() {
        if (app.documents.length === 0) return;
        runSafely(function () {
            app.executeMenuCommand("zoomout");
            app.executeMenuCommand("zoomin");
        }, "forceScreenRefresh failed", "zoomout/zoomin");
    }

    /**
     * 環境設定パネルを開く（モーダルを閉じてから呼ぶ）
     * Open a Preferences panel (call after the modal dialog is closed)
     * @param {string} menuCommand - メニューコマンド名（例: "UIPref"）
     * @returns {void}
     */
    function openPreferencePanel(menuCommand) {
        runSafely(function () {
            app.executeMenuCommand(menuCommand);
        }, "openPreferencePanel failed", menuCommand);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* レイアウトの種類と、切り替えで閉じたときに show() が返す値 / Layout names and the show() result when closed to switch */
    var LAYOUT_SIDEBAR = "sidebar";
    var LAYOUT_CLASSIC = "classic";
    var LAYOUT_SWITCH_RESULT = 3;
    var LAYOUT_STORAGE_KEY = "__" + SCRIPT_NAME + "_Layout"; /* 最後に使ったレイアウト / last used layout */

    /**
     * サイドバー表示の本体（左にカテゴリ一覧、右にページを重ねたスタック）を組み立てる
     * Build the sidebar view body (category list on the left, stacked pages on the right)
     * @param {object} dialogControls - コントロールを登録するオブジェクト
     * @param {Group} dialogContentGroup - 追加先
     * @returns {void}
     */
    function buildSidebarBody(dialogControls, dialogContentGroup) {
        var dialogBodyGroup = dialogContentGroup.add("group");
        dialogBodyGroup.orientation = "row";
        dialogBodyGroup.alignChildren = ["fill", "fill"];
        dialogBodyGroup.spacing = COLUMN_SPACING;
        var categoryList = dialogBodyGroup.add("listbox");
        dialogControls.categoryList = categoryList;

        var pagePanel = dialogBodyGroup.add("panel");
        setupPanel(pagePanel, PAGE_ROW_SPACING);
        pagePanel.orientation = "stack";
        pagePanel.alignChildren = ["fill", "top"];
        dialogControls.pageStack = pagePanel;
        buildDisplayPages(dialogControls);
        buildBehaviorPages(dialogControls);
        for (var i = 0; i < dialogControls.categoryPages.length; i++) {
            categoryList.add("item", getLabel(dialogControls.categoryPages[i].labelPath));
        }
        /* 全カテゴリが見える高さにする（ページ側もこの高さまで伸びる）/ Tall enough to show every category; the page panel stretches to match */
        categoryList.preferredSize = [measureCategoryListWidth(dialogControls),
            dialogControls.categoryPages.length * SIDEBAR_ITEM_HEIGHT + SIDEBAR_PADDING];

        /* 前回開いていたカテゴリで開く。一覧の空きをクリックして選択が外れたら選び直す */
        /* Reopen the last category; reselect when a click on empty space clears the selection */
        var lastPageKey = "__" + SCRIPT_NAME + "_LastPage";
        var lastPageIndex = $.global[lastPageKey];
        if (typeof lastPageIndex !== "number" || lastPageIndex >= dialogControls.categoryPages.length) lastPageIndex = 0;
        categoryList.selection = lastPageIndex;
        showCategoryPage(dialogControls, lastPageIndex);
        categoryList.onChange = function () {
            if (!categoryList.selection) {
                categoryList.selection = $.global[lastPageKey] || 0;
                return;
            }
            $.global[lastPageKey] = categoryList.selection.index;
            showCategoryPage(dialogControls, categoryList.selection.index);
        };
    }

    /**
     * 2カラム表示の本体（従来と同じパネルの並び）を組み立てる
     * Build the two-column view body (the same panels as the classic layout)
     * @param {object} dialogControls - コントロールを登録するオブジェクト
     * @param {Group} dialogContentGroup - 追加先
     * @returns {void}
     */
    function buildClassicBody(dialogControls, dialogContentGroup) {
        var panelColumnsGroup = dialogContentGroup.add("group");
        panelColumnsGroup.orientation = "row";
        panelColumnsGroup.alignChildren = "top";
        panelColumnsGroup.spacing = COLUMN_SPACING;
        buildClassicLeftColumn(dialogControls, addClassicColumn(panelColumnsGroup));
        buildClassicRightColumn(dialogControls, addClassicColumn(panelColumnsGroup));
    }

    /**
     * ダイアログを組み立てて表示し、［OK］で環境設定を保存する。
     * レイアウトの切り替えで閉じたときは、引き継ぐ状態を返す
     * Build and show the dialog, saving on [OK]; when closed to switch layouts, return the state to carry over
     * @param {string} layoutName - LAYOUT_SIDEBAR / LAYOUT_CLASSIC
     * @param {object|null} carriedSession - 前のレイアウトから引き継ぐ状態（presetState / presetIndex / brightnessIndex / brightnessTouched）
     * @returns {object|null} 切り替えで閉じたときは引き継ぐ状態、それ以外は null
     */
    function showPreferencesDialog(layoutName, carriedSession) {
        var isSidebarLayout = (layoutName === LAYOUT_SIDEBAR);
        /* コントロールと、チェックボックスの対応表をまとめて持つ / Holds the controls and the checkbox binding table */
        /* 行の一覧（settingRows）とカテゴリのページ（categoryPages）はサイドバー表示の変更の追跡に使う / settingRows and categoryPages drive the sidebar view's change tracking */
        var dialogControls = { checkboxBindings: [], settingRows: [], categoryPages: [] };
        var switchSession = null;

        /* ダイアログ本体 / Dialog window */
        var preferencesDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(preferencesDialog);

        /* 全体を縦に積むコンテナ / Vertical container for the whole dialog */
        var dialogContentGroup = preferencesDialog.add("group");
        dialogContentGroup.orientation = "column";
        dialogContentGroup.alignChildren = "fill";

        /* プリセット選択行（現在の設定 / デフォルト / プリセット1）/ Preset selector row (Current / Default / Preset 1) */
        var presetRow = dialogContentGroup.add("group");
        presetRow.margins = isSidebarLayout ? PRESET_ROW_MARGINS : CLASSIC_PRESET_ROW_MARGINS;
        setupRow(presetRow, "center", ROW_SPACING);
        var presetSelectorGroup = presetRow.add("group");
        setupRow(presetSelectorGroup, "left", ROW_SPACING);
        presetSelectorGroup.add("statictext", undefined, labelText("fieldLabel.preset"));
        var presetLabels = [];
        for (var i = 0; i < PRESET_CHOICES.length; i++) {
            presetLabels.push(getLabel(PRESET_CHOICES[i].labelPath));
        }
        var presetDropdown = presetSelectorGroup.add("dropdownlist", undefined, presetLabels);
        /* onChange を付ける前に選ぶので、引き継いだ値を上書きしない / chosen before onChange is attached, so the carried state is not overwritten */
        presetDropdown.selection = carriedSession ? carriedSession.presetIndex : 0; /* 初期選択は「現在の設定」/ Default selection is "Current Settings" */
        presetDropdown.helpTip = getLabel("tooltip.preset");

        if (isSidebarLayout) {
            buildSidebarBody(dialogControls, dialogContentGroup);
        } else {
            buildClassicBody(dialogControls, dialogContentGroup);
        }

        /* 下段：左に［レイアウト切り替え］（サイドバー表示は変更件数も）、右に［すべての変更を取り消す］［キャンセル］［OK］ */
        /* Bottom row: layout switch (and the change count in the sidebar view) on the left, buttons on the right */
        var buttonRow = addButtonRow(dialogContentGroup);
        var btnSwitchLayout = buttonRow.leftGroup.add("button", undefined,
            getLabel(isSidebarLayout ? "button.toClassicLayout" : "button.toSidebarLayout"));
        btnSwitchLayout.helpTip = getLabel("tooltip.switchLayout");
        var btnRevertAll = null;
        if (isSidebarLayout) {
            var changeCountLabel = buttonRow.leftGroup.add("statictext", undefined, "");
            changeCountLabel.preferredSize = [CHANGE_COUNT_WIDTH, -1];
            btnRevertAll = buttonRow.rightGroup.add("button", undefined, getLabel("button.revertAll"));
            dialogControls.changeCountLabel = changeCountLabel;
            dialogControls.btnRevertAll = btnRevertAll;
        }
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /* プリセット選択でUIを差し替え / Swap the UI when a preset is selected */
        presetDropdown.onChange = function () {
            if (!presetDropdown.selection) return;
            applyPresetChoice(dialogControls, presetDropdown.selection.index);
            updateChangeState(dialogControls);
            app.redraw();
        };

        /* 環境設定を読み直して変更をすべて捨てる / Reload the preferences and discard every change */
        if (btnRevertAll) {
            btnRevertAll.onClick = function () {
                presetDropdown.selection = 0; /* 「現在の設定」/ "Current Settings" */
                loadPreferencesIntoUI(dialogControls);
                updateChangeState(dialogControls);
            };
        }

        /* 今の状態を控えて閉じ、もう一方のレイアウトで開き直す / Keep the current state, close, and reopen in the other layout */
        btnSwitchLayout.onClick = function () {
            switchSession = {
                presetState: readPresetStateFromUI(dialogControls),
                presetIndex: presetDropdown.selection ? presetDropdown.selection.index : 0,
                brightnessIndex: dialogControls.brightnessPicker.selectedIndex,
                brightnessTouched: dialogControls.brightnessPicker.isTouched
            };
            preferencesDialog.close(LAYOUT_SWITCH_RESULT);
        };

        /* チェックOFFで入力欄を無効化（0＝非表示として保存）/ Disable the field when unchecked (saved as 0 = hidden) */
        dialogControls.recentFontsCheckbox.onClick = function () {
            var currentFontCount = 0;
            if (dialogControls.recentFontsCheckbox.value) {
                /* ONに戻したとき 0 のままでは矛盾するので既定値を入れる / Restore the default instead of leaving 0 when re-enabled */
                currentFontCount = parseIntOrNull(dialogControls.recentFontsInput.text);
                if (currentFontCount === null || currentFontCount < NUMERIC_INPUT_RULES.recentFonts.min) {
                    currentFontCount = NUMERIC_INPUT_RULES.recentFonts.defaultValue;
                }
            }
            applyRecentFontsCount(dialogControls, currentFontCount);
            updateChangeState(dialogControls);
        };

        dialogControls.recentFontsInput.onChange = function () {
            normalizeIntInput(dialogControls.recentFontsInput, NUMERIC_INPUT_RULES.recentFonts);
        };

        /* ドラッグを離したときに段階へスナップ / Snap to a step when the drag is released */
        dialogControls.anchorSizeSlider.onChange = function () {
            applyAnchorSizeStep(dialogControls.anchorSizeSlider, dialogControls.anchorSizeSlider.value);
        };

        dialogControls.historyStatesInput.onChange = function () {
            normalizeIntInput(dialogControls.historyStatesInput, NUMERIC_INPUT_RULES.historyStates);
        };

        btnOK.onClick = function () {
            /* 書き込みはすべて runSafely 経由なので、失敗しても例外は上がらない / Every write goes through runSafely, so nothing throws here */
            savePreferences(dialogControls);
            var willOpenUserInterface = saveBrightness(dialogControls.brightnessPicker);
            forceScreenRefresh();
            preferencesDialog.close();
            /* 環境設定パネルはモーダルダイアログを閉じてから開く（表示中は開けないため）*/
            /* Open the Preferences panel only after this modal dialog is closed (it cannot open while it is up) */
            if (willOpenUserInterface) openPreferencePanel("UIPref");
        };

        /* サイドバー表示はデフォルトの列を埋めてから、現在の環境設定を読み込んで変更を数える基準にする */
        /* The sidebar view fills the default column first, then loads the current preferences as the baseline */
        if (isSidebarLayout) {
            fillDefaultColumn(dialogControls);
            fitColumnWidths(dialogControls);
        }
        loadPreferencesIntoUI(dialogControls);
        if (isSidebarLayout) captureBaseline(dialogControls);

        /* 前のレイアウトでの変更を戻す / Restore the changes made in the previous layout */
        if (carriedSession) {
            applyPresetStateToUI(dialogControls, carriedSession.presetState);
            dialogControls.brightnessPicker.selectedIndex = carriedSession.brightnessIndex;
            dialogControls.brightnessPicker.isTouched = carriedSession.brightnessTouched;
            updateChangeState(dialogControls);
        }

        /* レイアウトを切り替えたときは前回の位置を使わず、既定の位置（画面の中央）で開く
           After a layout switch, skip the remembered location and open at the default (centered) position */
        if (carriedSession) $.global["__" + SCRIPT_NAME + "_DialogLocation"] = null;
        prepareDialogWindow(preferencesDialog, SCRIPT_NAME);
        var dialogResult = preferencesDialog.show();
        return (dialogResult === LAYOUT_SWITCH_RESULT) ? switchSession : null;
    }

    /**
     * 最後に使ったレイアウト（初回は2カラム表示）で開き、切り替えのたびに状態を引き継いで開き直す
     * Open in the last used layout (the two-column view at first) and reopen with the carried state on every switch
     * @returns {void}
     */
    function main() {
        var layoutName = ($.global[LAYOUT_STORAGE_KEY] === LAYOUT_SIDEBAR) ? LAYOUT_SIDEBAR : LAYOUT_CLASSIC;
        var carriedSession = null;
        do {
            carriedSession = showPreferencesDialog(layoutName, carriedSession);
            if (carriedSession) {
                layoutName = (layoutName === LAYOUT_SIDEBAR) ? LAYOUT_CLASSIC : LAYOUT_SIDEBAR;
                $.global[LAYOUT_STORAGE_KEY] = layoutName;
            }
        } while (carriedSession);
    }

    main();
})();
