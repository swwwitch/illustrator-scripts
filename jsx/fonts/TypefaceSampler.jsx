#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorで利用できるフォントをウェイト・スタイル順に並べ、ファミリー単位でアートボードへ整列描画します。
キーワード、ウェイト（5段階）、種類で対象を絞り込み、フォント名・PostScript名・サンプル・カスタムテキストから出力内容を選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypefaceSampler.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n103ac6622657

### Overview

Lays out the fonts available in Illustrator on the artboard, grouped by family and ordered by weight and style.
Narrow the list by keyword, weight (5 ranks) or style category, and output the font name, the PostScript name, a sample string, or your own custom text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypefaceSampler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TypefaceSampler";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypefaceSampler.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypefaceSampler.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n103ac6622657"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_COLUMN_COUNT = 3;  /* 列数の初期値 / default column count */
    var SAMPLE_FONT_SIZE     = 10; /* 描画するサンプルの文字サイズ（pt）/ sample font size in points */

    /* ラジオの見出しと描画内容を兼ねるサンプル文字列 / Sample strings used both as radio captions and as drawn text */
    var SAMPLE_ALPHABET_TEXT = "The quick brown fox jumps over the lazy dog.";
    var SAMPLE_NUMBERS_TEXT  = "1234567890";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS      = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING      = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS       = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING       = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING      = 12;               /* 2カラムの間隔 / gap between columns */
    var BUTTON_BAR_MARGINS  = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */
    var BUTTON_BAR_SPACING  = 10;               /* ボタンバー内の要素間隔 / spacing inside the button bar */
    var KEYWORD_FIELD_CHARS = 30;               /* キーワード欄の最小幅（文字数）/ minimum width of the keyword field */
    var COLUMN_FIELD_CHARS  = 3;                /* 列数欄の最小幅（文字数）/ minimum width of the column field */

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
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
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 左寄せの縦並びグループを生成する（ラジオ列など）
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addLeftAlignedColumn(parentContainer) {
        var createdGroup = parentContainer.add("group");
        createdGroup.orientation = "column";
        createdGroup.alignChildren = ["left", "center"];
        return createdGroup;
    }

    /**
     * 数値入力欄に上下キーでの増減を割り当てる
     * @param {EditText} editText - 対象の入力欄
     * @param {function} [onChanged] - 値が変わったときに呼ぶ関数
     * @param {number} [minValue] - 下限値（指定時はこれ未満にしない）
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onChanged, minValue) {
        editText.addEventListener("keydown", function(event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            event.preventDefault();

            if (keyboard.shiftKey) {
                /* Shift：10 単位にスナップ / Shift snaps to multiples of 10 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                /* Option：0.1 刻み / Option steps by 0.1 */
                value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                value = Math.round(value + (isUp ? 1 : -1));
            }

            if (typeof minValue === "number" && value < minValue) value = minValue;

            editText.text = value;
            if (typeof onChanged === "function") onChanged();
        });
    }

    /**
     * ラジオボタン群に上下キーでの選択移動を割り当てる
     * @param {RadioButton[]} radioButtons - 対象のラジオボタン
     * @returns {void}
     */
    function enableArrowKeyNavigation(radioButtons) {
        if (!radioButtons || radioButtons.length === 0) return;

        /* addEventListener を持つ祖先までさかのぼる / Walk up to an ancestor that accepts listeners */
        var eventTarget = radioButtons[0].parent;
        while (eventTarget && typeof eventTarget.addEventListener !== "function" && eventTarget.parent) {
            eventTarget = eventTarget.parent;
        }
        if (!eventTarget || typeof eventTarget.addEventListener !== "function") return;

        eventTarget.addEventListener("keydown", function(event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var currentIndex = -1;
            for (var i = 0; i < radioButtons.length; i++) {
                if (radioButtons[i].value) {
                    currentIndex = i;
                    break;
                }
            }
            if (currentIndex === -1) return;

            var lastIndex = radioButtons.length - 1;
            var nextIndex;
            if (event.keyName === "Up") {
                nextIndex = (currentIndex === 0) ? lastIndex : currentIndex - 1;
            } else {
                nextIndex = (currentIndex === lastIndex) ? 0 : currentIndex + 1;
            }

            radioButtons[nextIndex].value = true;
            radioButtons[nextIndex].active = true;
            if (typeof radioButtons[nextIndex].onClick === "function") {
                radioButtons[nextIndex].onClick();
            }
            event.preventDefault && event.preventDefault();
        });
    }

    // =========================================
    // 描画寸法 / Drawing metrics
    // =========================================
    var ARTBOARD_PADDING     = 20; /* アートボード端からの余白（pt）/ padding from the artboard edge */
    var SAMPLE_COLUMN_GAP    = 20; /* ファミリー列どうしのすき間（pt）/ gap between family columns */
    var SAMPLE_ROW_GAP       = 30; /* ファミリー行どうしのすき間（pt）/ gap between family rows */
    var CATEGORY_LINE_HEIGHT = 16; /* カテゴリー一覧の行送り（pt）/ line height of the category list */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        /* "ja" で始まるロケール（ja, ja_JP など）は日本語扱い / Treat "ja*" locales as Japanese */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title:        { ja: "フォントを一覧表示", en: "Typeface Sampler" },
            confirmTitle: { ja: "確認", en: "Confirmation" }
        },
        panel: {
            output: { ja: "出力内容", en: "Output Content" },
            option: { ja: "表示オプション", en: "Display Options" },
            weight: { ja: "ウェイト", en: "Weight" },
            type:   { ja: "種類", en: "Style" }
        },
        radio: {
            fontNameWeightStyle: { ja: "フォント名＋ウェイト／スタイル", en: "Font Name + Weight/Style" },
            postscriptName:      { ja: "PostScript名", en: "PostScript Name" },
            custom:              { ja: "カスタム", en: "Custom" }
        },
        checkbox: {
            showWeightCount: { ja: "ウェイト数", en: "Weight Count" },
            showWeightList:  { ja: "ウェイト一覧", en: "Weight List" },
            showScore:       { ja: "スコア（検証用）", en: "Debug Score" },
            weightVeryThin:  { ja: "超極細・極細", en: "Hairline / Thin" },
            weightLight:     { ja: "細め", en: "Light" },
            weightRegular:   { ja: "標準", en: "Regular" },
            weightSemiBold:  { ja: "中太", en: "Medium / SemiBold" },
            weightBold:      { ja: "太字・極太", en: "Bold / Black" },
            typeBasic:       { ja: "基本", en: "Basic" },
            typeNarrow:      { ja: "狭める系", en: "Condensed" },
            typeWide:        { ja: "広げる系", en: "Expanded" },
            typeDecor:       { ja: "装飾・特殊用途", en: "Display / Special" },
            typeSizeProp:    { ja: "サイズ・プロポーション系", en: "Size / Proportion" }
        },
        tooltip: {
            keyword:           { ja: "フォント名に含まれる文字で絞り込みます。空欄ならすべて表示します。", en: "Filters the list by text in the font name. Leave it empty to show everything." },
            displayFontName:   { ja: "各行に「フォント名＋ウェイト／スタイル」を表示します。", en: "Shows the font name with its weight and style on each line." },
            displayPostScript: { ja: "各行に PostScript 名を表示します。", en: "Shows the PostScript name on each line." },
            displaySample:     { ja: "各行にこのサンプル文字を表示します。", en: "Shows this sample text on each line." },
            displayCustom:     { ja: "各行に、下の欄に入れた文字を表示します。", en: "Shows the text you type below on each line." },
            customText:        { ja: "「カスタム」で表示する文字です。", en: "The text shown when Custom is selected." },
            showWeightCount:   { ja: "フォントファミリーごとのウェイト数を添えます。", en: "Adds the number of weights in each family." }
        },
        fieldLabel: {
            keyword: { ja: "フォント名に含まれるキーワード（空欄→全対象）", en: "Keyword in font name (leave blank for all)" },
            columns: { ja: "列数", en: "Columns" }
        },
        defaultText: {
            customSample: { ja: "愛のあるユニークで豊かな書体ABCabcGg349", en: "Lorem ipsum dolor sit amet, consectetur adipiscing elit" }
        },
        button: {
            cancel:  { ja: "キャンセル", en: "Cancel" },
            ok:      { ja: "OK", en: "OK" },
            stop:    { ja: "中止する", en: "Cancel" },
            proceed: { ja: "続行する", en: "Proceed" }
        },
        alert: {
            noDocument:          { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noMatchingFont:      { ja: "条件に該当するフォントが見つかりませんでした。", en: "No font matched the given conditions." },
            errorOccurred:       { ja: "エラーが発生しました：", en: "An error occurred:" },
            confirmAllFonts:     { ja: "すべてのフォントを対象に実行しますか？", en: "Do you want to process all fonts?" },
            confirmAllFontsNote: { ja: "非常に時間がかかることがあります。", en: "This may take a long time." }
        }
    };

    /**
     * 現在のUI言語に対応するラベル文字列を返す
     * @param {object} labelSet - ja / en を持つラベル定義
     * @returns {string} 表示用の文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * コロン付きラベルを返す（日本語は全角、英語は半角）
     * @param {object} labelSet - ja / en を持つラベル定義
     * @returns {string} コロンを付けた表示用の文字列
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialogs
    // =========================================

    /* ラジオボタンの並びに対応する出力モード / Output modes matching the radio button order */
    var DISPLAY_MODES = ["family+style", "postscript", "alphabet", "numbers", "custom"];

    /**
     * 列数欄の文字列を1以上の整数に直す
     * @param {string} columnText - 列数欄の文字列
     * @returns {number} 列数。数値として読めなければ DEFAULT_COLUMN_COUNT
     */
    function parseColumnCount(columnText) {
        var columnCount = parseInt(columnText, 10);
        if (isNaN(columnCount)) return DEFAULT_COLUMN_COUNT;
        return Math.max(1, columnCount);
    }

    /**
     * ダイアログを表示してユーザー入力を取得する
     * @returns {object} 入力内容。キャンセル時は null
     */
    function showFontListDialog() {
        var dialogWindow = new Window("dialog", getLabel(LABELS.dialog.title));
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.margins = WINDOW_MARGINS;
        dialogWindow.spacing = WINDOW_SPACING;

        dialogWindow.add("statictext", undefined, labelText(LABELS.fieldLabel.keyword));
        var keywordField = dialogWindow.add("edittext", undefined, "");
        keywordField.helpTip = getLabel(LABELS.tooltip.keyword);
        keywordField.characters = KEYWORD_FIELD_CHARS;
        keywordField.active = true;

        /* 出力内容 / Output content */
        var outputPanel = addPanel(dialogWindow, getLabel(LABELS.panel.output));
        var displayModeColumn = addLeftAlignedColumn(outputPanel);

        var displayModeRadios = [];
        displayModeRadios[0] = displayModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.fontNameWeightStyle));
        displayModeRadios[0].helpTip = getLabel(LABELS.tooltip.displayFontName);
        displayModeRadios[1] = displayModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.postscriptName));
        displayModeRadios[1].helpTip = getLabel(LABELS.tooltip.displayPostScript);
        displayModeRadios[2] = displayModeColumn.add("radiobutton", undefined, SAMPLE_ALPHABET_TEXT);
        displayModeRadios[2].helpTip = getLabel(LABELS.tooltip.displaySample);
        displayModeRadios[3] = displayModeColumn.add("radiobutton", undefined, SAMPLE_NUMBERS_TEXT);
        displayModeRadios[3].helpTip = getLabel(LABELS.tooltip.displaySample);
        displayModeRadios[4] = displayModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.custom));
        displayModeRadios[4].helpTip = getLabel(LABELS.tooltip.displayCustom);
        displayModeRadios[0].value = true;

        var customTextField = outputPanel.add("edittext", undefined, getLabel(LABELS.defaultText.customSample));
        customTextField.helpTip = getLabel(LABELS.tooltip.customText);
        customTextField.characters = KEYWORD_FIELD_CHARS;
        customTextField.enabled = false;

        enableArrowKeyNavigation(displayModeRadios);

        /* 表示オプション / Display options */
        var optionPanel = addPanel(dialogWindow, getLabel(LABELS.panel.option));

        var weightOptionRow = optionPanel.add("group");
        setupRow(weightOptionRow);
        var showWeightCountCheckbox = weightOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.showWeightCount));
        showWeightCountCheckbox.helpTip = getLabel(LABELS.tooltip.showWeightCount);
        var showWeightListCheckbox = weightOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.showWeightList));
        showWeightListCheckbox.value = true;

        var columnAndScoreRow = optionPanel.add("group");
        setupRow(columnAndScoreRow);
        columnAndScoreRow.add("statictext", undefined, labelText(LABELS.fieldLabel.columns));
        var columnField = columnAndScoreRow.add("edittext", undefined, DEFAULT_COLUMN_COUNT + "");
        columnField.characters = COLUMN_FIELD_CHARS;

        /**
         * 列数欄を1以上の整数に書き直す（option＋↑↓の小数刻みも丸める）
         * @returns {void}
         */
        function normalizeColumnField() {
            columnField.text = parseColumnCount(columnField.text) + "";
        }
        columnField.onChange = normalizeColumnField;
        changeValueByArrowKey(columnField, normalizeColumnField, 1);

        var showScoreCheckbox = columnAndScoreRow.add("checkbox", undefined, getLabel(LABELS.checkbox.showScore));

        /* ウェイト・種類による絞り込み / Weight and style filters */
        var filterRow = dialogWindow.add("group");
        filterRow.orientation = "row";
        /* 2つのパネルを同じ幅で並べ、上端をそろえる / Lay both panels out at equal width, aligned at the top */
        filterRow.alignment = ["fill", "top"];
        filterRow.alignChildren = ["fill", "top"];
        filterRow.spacing = COLUMN_SPACING;

        var weightPanel = addPanel(filterRow, getLabel(LABELS.panel.weight));
        var weightVeryThinCheckbox = weightPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.weightVeryThin));
        var weightLightCheckbox    = weightPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.weightLight));
        var weightRegularCheckbox  = weightPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.weightRegular));
        var weightSemiBoldCheckbox = weightPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.weightSemiBold));
        var weightBoldCheckbox     = weightPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.weightBold));

        var typePanel = addPanel(filterRow, getLabel(LABELS.panel.type));
        var typeBasicCheckbox    = typePanel.add("checkbox", undefined, getLabel(LABELS.checkbox.typeBasic));
        var typeNarrowCheckbox   = typePanel.add("checkbox", undefined, getLabel(LABELS.checkbox.typeNarrow));
        var typeWideCheckbox     = typePanel.add("checkbox", undefined, getLabel(LABELS.checkbox.typeWide));
        var typeDecorCheckbox    = typePanel.add("checkbox", undefined, getLabel(LABELS.checkbox.typeDecor));
        var typeSizePropCheckbox = typePanel.add("checkbox", undefined, getLabel(LABELS.checkbox.typeSizeProp));

        /**
         * 「ウェイト一覧」の状態に合わせて各コントロールの有効・無効を更新する
         * @returns {void}
         */
        function updateDialogState() {
            outputPanel.enabled = showWeightListCheckbox.value;
            columnField.enabled = showWeightListCheckbox.value;

            /* 一覧をやめたときは1列・フォント名表示に戻す / Fall back to a single column of font names */
            if (!showWeightListCheckbox.value) {
                for (var i = 0; i < displayModeRadios.length; i++) {
                    displayModeRadios[i].value = false;
                }
                displayModeRadios[0].value = true;
                columnField.text = "1";
            }

            /* スコアはフォント名＋ウェイト／スタイル表示のときだけ有効 / Score applies to the family+style mode only */
            if (!showWeightListCheckbox.value || !displayModeRadios[0].value) {
                showScoreCheckbox.value = false;
                showScoreCheckbox.enabled = false;
            } else {
                showScoreCheckbox.enabled = true;
            }
        }

        displayModeRadios[4].onClick = function() {
            customTextField.enabled = true;
            updateDialogState();
        };
        for (var i = 0; i < displayModeRadios.length - 1; i++) {
            displayModeRadios[i].onClick = function() {
                customTextField.enabled = false;
                updateDialogState();
            };
        }
        showWeightListCheckbox.onClick = updateDialogState;
        updateDialogState();

        /* メイングループ（横並び）/ Main group (horizontal layout) */
        var btnRowGroup = dialogWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.spacing = BUTTON_BAR_SPACING;

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ / Right-side button group */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        if (dialogWindow.show() !== 1) return null;

        var displayMode = DISPLAY_MODES[0];
        for (i = 0; i < displayModeRadios.length; i++) {
            if (displayModeRadios[i].value) {
                displayMode = DISPLAY_MODES[i];
                break;
            }
        }

        return {
            keyword: keywordField.text.replace(/^\s+|\s+$/g, ""),
            displayMode: displayMode,
            customText: customTextField.text,
            columns: parseColumnCount(columnField.text),
            showWeight: showWeightListCheckbox.value,
            showWeightCount: showWeightCountCheckbox.value,
            showScore: showScoreCheckbox.value,
            weightFilters: {
                veryThin: weightVeryThinCheckbox.value,
                light: weightLightCheckbox.value,
                regular: weightRegularCheckbox.value,
                semiBold: weightSemiBoldCheckbox.value,
                bold: weightBoldCheckbox.value
            },
            typeFilters: {
                basic: typeBasicCheckbox.value,
                narrow: typeNarrowCheckbox.value,
                wide: typeWideCheckbox.value,
                decor: typeDecorCheckbox.value,
                sizeProp: typeSizePropCheckbox.value
            }
        };
    }

    /**
     * 全フォントを対象にしてよいか確認する
     * @returns {boolean} 続行する場合は true
     */
    function confirmShowAllFonts() {
        var confirmDialog = new Window("dialog", getLabel(LABELS.dialog.confirmTitle));
        confirmDialog.orientation = "column";
        confirmDialog.alignChildren = ["left", "top"];
        confirmDialog.margins = WINDOW_MARGINS;
        confirmDialog.spacing = WINDOW_SPACING;

        confirmDialog.add("statictext", undefined, getLabel(LABELS.alert.confirmAllFonts));
        confirmDialog.add("statictext", undefined, getLabel(LABELS.alert.confirmAllFontsNote));

        /* メイングループ（横並び）/ Main group (horizontal layout) */
        var btnRowGroup = confirmDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.spacing = BUTTON_BAR_SPACING;

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ / Right-side button group */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.stop), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.proceed), { name: "ok" });

        return confirmDialog.show() === 1;
    }

    // =========================================
    // ウェイト語句の定義 / Weight term definitions
    // =========================================

    /* ウェイト語句の並び（インデックスが大きいほど太い）/ Weight terms ordered from thin to bold */
    var WEIGHT_GROUPS = [
        ["hairline"], // +0
        ["ultra thin", "ultrathin", "ut"], // +1
        ["thin", "th"], // +2
        ["default"], // +3
        ["ultralight", "ultra light", "ultlt", "ul"], // +4
        ["extralight", "extra light", "el", "xlight", "xl"], // +5
        ["lightsemi"], // +6
        ["light", "lt", "lite", "l"], // +7
        ["lb"], // +8
        ["book", "bk"], // +9
        ["n", "normal"], // +10
        ["middle"], // +11
        ["regular", "roman", "レギュラー", "r"], // +12
        ["rb"], // +13
        ["medium", "md", "ミディアム", "m"], // +14
        ["semibold", "semi bold", "sb"], // +15
        ["demibold", "demi bold", "db", "デミボールド", "demi", "d", "demixtra"], // +16
        ["bold", "bd", "ボールド", "b"], // +17
        ["extrabold", "extra bold", "xbold", "エクストラボールド", "e", "eb", "xb"], // +18
        ["heavy", "h"], // +19
        ["black"], // +20
        ["xblack", "extra black", "extrablack"], // +21
        ["ultra", "u", "ub", "ultra black", "ultrablack"] // +22
    ];

    /* 単独で使われたら Regular 扱いする装飾語句 / Decoration-only styles treated as Regular */
    var DECORATION_ONLY_STYLES = [
        "display", "compressed", "comp", "compact", "expanded", "extended", "semiextended",
        "ultracondensed", "extracondensed", "semicondensed", "cond", "condensed", "wide",
        "headline", "text", "low", "micro", "extra compressed",
        "semi expanded", "semiexpanded"
    ];

    /* 幅を表す複合語（Ultra Condensed など）。ultra / extra をウェイト語と取り違えないよう照合前に除く
       Width compounds such as "ultra condensed"; removed first so "ultra" is not read as a weight */
    var WIDTH_COMPOUND_PATTERN = /(^|\s)(ultra|extra|semi)\s+(condensed|cond|compressed|comp|expanded|extended)(?=\s|$)/g;

    /* WEIGHT_GROUPS における Regular のインデックス / Index of "regular" in WEIGHT_GROUPS */
    var REGULAR_GROUP_INDEX = (function() {
        for (var i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (var j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                if (WEIGHT_GROUPS[i][j] === "regular") return i;
            }
        }
        return 12; /* fallback */
    })();

    /* 複合語一致用に、長い語から順に並べた照合テーブル / Match table sorted by term length */
    var WEIGHT_TERM_PATTERNS = (function() {
        var termPatterns = [];
        for (var i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (var j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                var weightTerm = WEIGHT_GROUPS[i][j];
                /* \b は和文の前後で効かないため、英数字以外を境界とみなす / \b fails next to Japanese, so treat any non-alphanumeric as a boundary */
                termPatterns.push({
                    term: weightTerm,
                    groupIndex: i,
                    pattern: new RegExp("(?:^|[^a-z0-9])" + weightTerm.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + "(?=[^a-z0-9]|$)")
                });
            }
        }
        /* 同じ長さなら細い方を先に（並べ替えの結果を環境で変えない）/ Break ties by group so the order is deterministic */
        termPatterns.sort(function(a, b) {
            return (b.term.length - a.term.length) || (a.groupIndex - b.groupIndex);
        });
        return termPatterns;
    })();

    /**
     * スタイル文字列を照合用に正規化する
     * @param {string} rawStyle - font.style の値
     * @returns {string} 小文字化し、区切り記号を空白に置き換えた文字列
     */
    function normalizeStyle(rawStyle) {
        return (rawStyle || "").toLowerCase().replace(/[_\-]+/g, " ").replace(/^\s+|\s+$/g, "");
    }

    /**
     * スタイル文字列に一致する WEIGHT_GROUPS のインデックスを返す
     * 幅の複合語を除いてから、完全一致を優先し、なければ長い語から順に語の境界つきで照合する
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {number} 一致したインデックス。見つからない場合は -1
     */
    function getWeightGroupIndex(normalizedStyle) {
        var weightStyle = normalizedStyle.replace(WIDTH_COMPOUND_PATTERN, " ").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
        if (weightStyle === "") return -1;

        var i, j;
        for (i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                if (weightStyle === WEIGHT_GROUPS[i][j]) return i;
            }
        }

        for (i = 0; i < WEIGHT_TERM_PATTERNS.length; i++) {
            if (WEIGHT_TERM_PATTERNS[i].pattern.test(weightStyle)) return WEIGHT_TERM_PATTERNS[i].groupIndex;
        }

        return -1;
    }

    /**
     * W3・W600・25 Ultra Light のような数値スタイルを読み取る
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {object|null} value（数値）と digitCount（桁数）。数値スタイルでなければ null
     */
    function getNumericWeight(normalizedStyle) {
        var numericMatch = normalizedStyle.match(/^w?(\d{1,3})(?=\D|$)/);
        if (!numericMatch) return null;
        return { value: parseInt(numericMatch[1], 10), digitCount: numericMatch[1].length };
    }

    // =========================================
    // ウェイト評価 / Weight scoring
    // =========================================

    /**
     * スタイル文字列に対する基本ウェイトスコアを取得する
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @param {string} postscriptName - 小文字化した PostScript 名
     * @param {string} familyName - 小文字化したファミリー名
     * @returns {number} ウェイトの評価値（小さいほど細い）
     */
    function getBaseWeightScore(normalizedStyle, postscriptName, familyName) {
        var styleWords = normalizedStyle.split(/\s+/);
        var i;

        var applyFrutigerCorrection = (/frutiger/i.test(familyName) && /ultralight/.test(normalizedStyle));

        /* W0〜W9、W000〜W999、先頭数値（例：25 Ultra Light）/ Numeric styles */
        var numericWeight = getNumericWeight(normalizedStyle);
        if (numericWeight) return numericWeight.value;

        /* 特例：HelveticaNeue, Tazugane, UniversNextPro + Ultra Light → 999 */
        if (
            (
                /helveticaneue/i.test(postscriptName) ||
                /tazugane/i.test(postscriptName) ||
                /universnextpro/i.test(postscriptName)
            ) &&
            /ultralight|ultra light|ultlt/i.test(normalizedStyle)
        ) {
            return 999;
        }

        /* 単独語が italic / oblique / wide → Regular 扱い */
        if (styleWords.length === 1 && /^(italic|oblique|it|wide)$/.test(styleWords[0])) {
            return 1000 + REGULAR_GROUP_INDEX;
        }

        /* 装飾語だけなら Regular 扱い */
        if (styleWords.length === 1) {
            for (i = 0; i < DECORATION_ONLY_STYLES.length; i++) {
                if (styleWords[0] === DECORATION_ONLY_STYLES[i]) return 1000 + REGULAR_GROUP_INDEX;
            }
        }

        /* 完全一致・複合語一致（長い語優先）/ Exact match, then longest-term match */
        var groupIndex = getWeightGroupIndex(normalizedStyle);
        if (groupIndex !== -1) {
            var weightScore = 1000 + groupIndex;
            if (applyFrutigerCorrection && groupIndex === 4) weightScore -= 5;
            return weightScore;
        }

        /* fallbackScore：Regular 扱い */
        var fallbackScore = 1000 + REGULAR_GROUP_INDEX;
        if (applyFrutigerCorrection) fallbackScore -= 5;
        return fallbackScore;
    }

    /**
     * ウェイトと装飾語をあわせた並べ替え用の評価値を取得する
     * @param {object} fontInfo - readFontInfo() が返したフォント情報
     * @returns {number} 並べ替えに使う評価値
     */
    function getFontSortScore(fontInfo) {
        var styleName = normalizeStyle(fontInfo.style);
        var postscriptName = fontInfo.name.toLowerCase();
        var familyName = fontInfo.family.toLowerCase();

        /* 特例：PostScript名が「FuturaPT-Heavy」なら 1015 固定（加点処理なし）/ Fixed rank, no offsets */
        if (postscriptName === "futurapt-heavy") return 1015;

        var baseScore = getBaseWeightScore(styleName, postscriptName, familyName);
        var decorationOffset = 0;
        var styleWords = styleName.split(/\s+/);

        /* 装飾フラグ初期化 / Initialize decoration flags */
        var decorationFlags = {
            hasText: false,
            hasHeadline: false,
            hasCondensed: false,
            hasCn: false,
            hasExpanded: false,
            hasExtended: false,
            hasUltraCondensed: false,
            hasExtraCondensed: false,
            hasSemiCondensed: false,
            hasCompressed: false,
            hasExtraCompressed: false,
            hasCompact: false,
            hasDisplay: false,
            hasMicro: false,
            hasLow: false,
            hasWide: false
        };

        /* 装飾キーワードに応じたフラグ設定 / Set a flag for each decoration keyword */
        for (var i = 0; i < styleWords.length; i++) {
            var styleWord = styleWords[i];
            if (styleWord === "text") decorationFlags.hasText = true;
            if (styleWord === "headline") decorationFlags.hasHeadline = true;
            if (styleWord === "cond" || styleWord === "condensed") decorationFlags.hasCondensed = true;
            if (styleWord === "cn") decorationFlags.hasCn = true;
            if (styleWord === "expanded") decorationFlags.hasExpanded = true;
            if (styleWord === "extended") decorationFlags.hasExtended = true;
            if (styleWord === "semiextended" || (styleWord === "semi" && styleWords[i + 1] === "extended")) decorationFlags.hasExtended = true;
            if (styleWord === "semiexpanded" || (styleWord === "semi" && styleWords[i + 1] === "expanded")) decorationFlags.hasExpanded = true;
            if (styleWord === "ultracondensed" || (styleWord === "ultra" && styleWords[i + 1] === "condensed")) decorationFlags.hasUltraCondensed = true;
            if (styleWord === "extracondensed" || (styleWord === "extra" && styleWords[i + 1] === "condensed")) decorationFlags.hasExtraCondensed = true;
            if (styleWord === "semicondensed" || (styleWord === "semi" && styleWords[i + 1] === "condensed")) decorationFlags.hasSemiCondensed = true;
            if (styleWord === "compressed" || styleWord === "comp") decorationFlags.hasCompressed = true;
            if (styleWord === "extra" && styleWords[i + 1] === "compressed") decorationFlags.hasExtraCompressed = true;
            if (styleWord === "compact") decorationFlags.hasCompact = true;
            if (styleWord === "display") decorationFlags.hasDisplay = true;
            if (styleWord === "micro") decorationFlags.hasMicro = true;
            if (styleWord === "low") decorationFlags.hasLow = true;
            if (styleWord === "wide") decorationFlags.hasWide = true;
        }

        /* Italic 判定（全体 styleName に対して）/ Detect italic across the whole styleName */
        var isItalic = /italic|oblique|slanted|inclined|kursiv|\bit\b/.test(styleName);

        /* 加点処理（100刻み + 特例あり）/ Offsets in steps of 100, with exceptions */
        if (decorationFlags.hasDisplay) decorationOffset += 100;
        if (decorationFlags.hasCompressed) decorationOffset += 200;
        if (decorationFlags.hasCompact) decorationOffset += 300;
        if (decorationFlags.hasExpanded) decorationOffset += 400;
        if (decorationFlags.hasExtended) decorationOffset += 500;
        if (decorationFlags.hasUltraCondensed) decorationOffset += 600;
        if (decorationFlags.hasExtraCondensed) decorationOffset += 700;
        if (decorationFlags.hasSemiCondensed) decorationOffset += 850;

        /* Condensed系代表加点（複数条件一致でも一度のみ）/ Applied once even on multiple matches */
        if (
            decorationFlags.hasCondensed ||
            decorationFlags.hasCn ||
            decorationFlags.hasWide ||
            decorationFlags.hasSemiCondensed ||
            decorationFlags.hasExtraCompressed
        ) {
            decorationOffset += 900;
        }

        if (decorationFlags.hasHeadline) decorationOffset += 1000;
        if (decorationFlags.hasText) decorationOffset += 1100;
        if (decorationFlags.hasLow) decorationOffset += 1200;
        if (decorationFlags.hasMicro) decorationOffset += 1250;
        if (decorationFlags.hasWide) decorationOffset += 1275;
        if (decorationFlags.hasExtraCompressed) decorationOffset += 150; /* 特別加点 / Extra offset */
        if (isItalic) decorationOffset += 1300;

        return baseScore + decorationOffset;
    }

    /**
     * 数値スタイルのウェイトを5段階カテゴリーに変換する
     * @param {number} numericWeight - 数値ウェイト
     * @param {number} digitCount - 桁数（3=CSS相当、2=Adobe系、1=和文のW0〜W9）
     * @returns {string} 5段階カテゴリーのキー
     */
    function getWeightCategoryFromNumber(numericWeight, digitCount) {
        /* 100〜900（CSS相当）/ CSS-style numeric weight */
        if (digitCount >= 3) {
            if (numericWeight < 200) return "veryThin";
            if (numericWeight < 400) return "light";
            if (numericWeight < 500) return "regular";
            if (numericWeight < 700) return "semiBold";
            return "bold";
        }

        /* 25 Ultra Light 〜 95 Black（Adobe系）/ Adobe-style two-digit weight */
        if (digitCount === 2) {
            if (numericWeight < 40) return "veryThin";
            if (numericWeight < 50) return "light";
            if (numericWeight < 60) return "regular";
            if (numericWeight < 70) return "semiBold";
            return "bold";
        }

        /* W0〜W9（和文）/ Single-digit weight used by Japanese fonts */
        if (numericWeight <= 1) return "veryThin";
        if (numericWeight <= 3) return "light";
        if (numericWeight === 4) return "regular";
        if (numericWeight <= 6) return "semiBold";
        return "bold";
    }

    /**
     * WEIGHT_GROUPS のインデックスを5段階カテゴリーに変換する
     * @param {number} groupIndex - WEIGHT_GROUPS のインデックス
     * @returns {string} 5段階カテゴリーのキー
     */
    function getWeightCategoryFromGroupIndex(groupIndex) {
        if (groupIndex <= 2) return "veryThin";  /* hairline 〜 thin */
        if (groupIndex === 3) return "regular";  /* default（既定ウェイト）/ default weight */
        if (groupIndex <= 8) return "light";     /* ultralight 〜 light */
        if (groupIndex <= 13) return "regular";  /* book 〜 regular */
        if (groupIndex <= 16) return "semiBold"; /* medium 〜 demibold */
        return "bold";                           /* bold 以上 */
    }

    /**
     * フォントのウェイトを5段階カテゴリーに分類する
     * @param {object} fontInfo - readFontInfo() が返したフォント情報
     * @returns {string} veryThin / light / regular / semiBold / bold のいずれか
     */
    function getWeightCategory(fontInfo) {
        var normalizedStyle = normalizeStyle(fontInfo.style);

        var numericWeight = getNumericWeight(normalizedStyle);
        if (numericWeight) return getWeightCategoryFromNumber(numericWeight.value, numericWeight.digitCount);

        var groupIndex = getWeightGroupIndex(normalizedStyle);
        if (groupIndex === -1) return "regular";
        return getWeightCategoryFromGroupIndex(groupIndex);
    }

    // =========================================
    // フォント収集 / Font collection
    // =========================================

    /* 種類フィルターの判定に使う語（部分一致）。Compact は「装飾・特殊用途」なので comp から外す
       Words for the style filters (substring match); "compact" belongs to decor, not to "comp" */
    var TYPE_FILTER_PATTERNS = {
        basic:    /text|headline/,
        narrow:   /cond|cn|comp(?!act)/,
        wide:     /expanded|extended/,
        decor:    /compact|display/,
        sizeProp: /micro|low|wide/
    };

    /**
     * フォントの名前情報を1回だけ読み取って控える（DOMの読み直しを避ける）
     * @param {TextFont} textFont - 対象のフォント
     * @returns {object} textFont と name / family / style を持つフォント情報
     */
    function readFontInfo(textFont) {
        return {
            textFont: textFont,
            name: textFont.name || "",
            family: textFont.family || "",
            style: textFont.style || ""
        };
    }

    /**
     * 環境にないフォントの置き換え用の仮エントリか判定する
     * @param {object} fontInfo - readFontInfo() が返したフォント情報
     * @returns {boolean} 仮エントリなら true
     */
    function isPlaceholderFont(fontInfo) {
        /* 仮エントリはスタイルが空で、ファミリー名に PostScript 名が入る / Placeholders have no style and reuse the PostScript name as family */
        return fontInfo.style === "" && fontInfo.family === fontInfo.name;
    }

    /**
     * 引用符つき検索用に文字列を正規化する（空白・ハイフンを除去）
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 正規化した文字列
     */
    function normalizeForSearch(sourceText) {
        return sourceText.toLowerCase().replace(/[\s\-　]/g, "");
    }

    /**
     * 検索語1つを解釈する（^ で先頭一致、引用符で空白・ハイフンを無視した一致）
     * @param {string} tokenText - 検索語の文字列
     * @returns {object|null} keyword / isPrefix / isQuoted を持つ検索語。空なら null
     */
    function parseSearchTerm(tokenText) {
        var termText = tokenText.toLowerCase();

        var isPrefix = (termText.charAt(0) === "^");
        if (isPrefix) termText = termText.substring(1);

        var isQuoted = (termText.length >= 2 && termText.charAt(0) === '"' && termText.charAt(termText.length - 1) === '"');
        if (isQuoted) termText = normalizeForSearch(termText.slice(1, -1));

        if (termText.length === 0) return null;
        return { keyword: termText, isPrefix: isPrefix, isQuoted: isQuoted };
    }

    /**
     * キーワード入力を AND / OR / NOT の検索条件に分解する
     * @param {string} rawKeyword - ダイアログで入力されたキーワード
     * @returns {object} andGroups（ORの配列をANDで並べたもの）と notTerms（除外する検索語）を持つ検索条件
     */
    function parseKeywordQuery(rawKeyword) {
        var keywordQuery = { andGroups: [], notTerms: [] };
        if (!rawKeyword) return keywordQuery;

        /* 全角の空白・カンマを半角に / Full-width space and comma to half-width */
        var normalizedKeyword = rawKeyword.replace(/　/g, " ").replace(/，/g, ",");

        /* 引用符で囲んだ語句は空白を含めて1語、+ は AND の区切り、空白・カンマは OR の区切り
           A quoted phrase is one token; "+" separates AND groups; spaces and commas separate OR terms */
        var tokens = normalizedKeyword.match(/[\-\^]*"[^"]*"|\+|[^\s,+]+/g) || [];

        var orGroup = [];
        for (var i = 0; i < tokens.length; i++) {
            if (tokens[i] === "+") {
                if (orGroup.length > 0) keywordQuery.andGroups.push(orGroup);
                orGroup = [];
                continue;
            }

            /* - 付きは除外。- だけの語は無視する / "-" marks a NOT term; a bare "-" is ignored */
            var isNot = (tokens[i].charAt(0) === "-");
            var searchTerm = parseSearchTerm(isNot ? tokens[i].substring(1) : tokens[i]);
            if (!searchTerm) continue;

            if (isNot) {
                keywordQuery.notTerms.push(searchTerm);
            } else {
                orGroup.push(searchTerm);
            }
        }
        if (orGroup.length > 0) keywordQuery.andGroups.push(orGroup);

        return keywordQuery;
    }

    /**
     * 検索語1つがフォント名・ファミリー名・スタイル名のいずれかに合致するか判定する
     * @param {object} searchTerm - parseSearchTerm() が返した検索語
     * @param {string} postscriptName - 小文字化した PostScript 名
     * @param {string} familyName - 小文字化したファミリー名
     * @param {string} styleName - 小文字化したスタイル名
     * @returns {boolean} 合致すれば true
     */
    function matchesKeywordTerm(searchTerm, postscriptName, familyName, styleName) {
        var candidateNames = [postscriptName, familyName, styleName];
        for (var i = 0; i < candidateNames.length; i++) {
            var candidateText = searchTerm.isQuoted ? normalizeForSearch(candidateNames[i]) : candidateNames[i];
            var matchIndex = candidateText.indexOf(searchTerm.keyword);
            if (searchTerm.isPrefix ? matchIndex === 0 : matchIndex !== -1) return true;
        }
        return false;
    }

    /**
     * フォントが検索条件に合致するか判定する
     * @param {object} keywordQuery - parseKeywordQuery() が返した検索条件
     * @param {string} postscriptName - 小文字化した PostScript 名
     * @param {string} familyName - 小文字化したファミリー名
     * @param {string} styleName - 小文字化したスタイル名
     * @returns {boolean} 合致すれば true
     */
    function matchesKeywordQuery(keywordQuery, postscriptName, familyName, styleName) {
        var i, j;

        /* NOT条件：含まれていたら除外 / Exclude when a NOT term matches */
        for (i = 0; i < keywordQuery.notTerms.length; i++) {
            if (matchesKeywordTerm(keywordQuery.notTerms[i], postscriptName, familyName, styleName)) return false;
        }

        /* AND × OR 条件 / AND groups of OR terms */
        for (i = 0; i < keywordQuery.andGroups.length; i++) {
            var orGroup = keywordQuery.andGroups[i];
            var orGroupMatched = false;

            for (j = 0; j < orGroup.length; j++) {
                if (matchesKeywordTerm(orGroup[j], postscriptName, familyName, styleName)) {
                    orGroupMatched = true;
                    break;
                }
            }

            if (!orGroupMatched) return false;
        }

        return true;
    }

    /**
     * フィルターのいずれかが選択されているか判定する
     * @param {object} filters - チェックボックスの選択状態
     * @returns {boolean} 1つでも選択されていれば true
     */
    function hasAnyFilterSelected(filters) {
        if (!filters) return false;
        for (var filterKey in filters) {
            if (filters.hasOwnProperty(filterKey) && filters[filterKey]) return true;
        }
        return false;
    }

    /**
     * 種類フィルターに合致するか判定する（複数カテゴリーへの該当を許容）
     * @param {string} styleName - 小文字化したスタイル文字列
     * @param {object} typeFilters - 種類フィルターの選択状態
     * @returns {boolean} 選択されたカテゴリーのいずれかに該当すれば true
     */
    function matchesTypeFilters(styleName, typeFilters) {
        for (var filterKey in TYPE_FILTER_PATTERNS) {
            if (!TYPE_FILTER_PATTERNS.hasOwnProperty(filterKey)) continue;
            if (typeFilters[filterKey] && TYPE_FILTER_PATTERNS[filterKey].test(styleName)) return true;
        }
        return false;
    }

    /**
     * 条件に合致するフォントを収集し、ファミリー単位にまとめる
     * @param {object} keywordQuery - parseKeywordQuery() が返した検索条件
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @returns {object} ファミリー名をキーにしたフォント情報の配列
     */
    function collectFonts(keywordQuery, userInput) {
        var groupedFonts = {};
        var seenNames = {};
        var weightFilters = userInput.weightFilters;
        var typeFilters = userInput.typeFilters;
        var useWeightFilter = hasAnyFilterSelected(weightFilters);
        var useTypeFilter = hasAnyFilterSelected(typeFilters);

        var installedFonts = app.textFonts;
        var fontCount = installedFonts.length;
        for (var i = 0; i < fontCount; i++) {
            var fontInfo = readFontInfo(installedFonts[i]);
            if (isPlaceholderFont(fontInfo)) continue;

            var postscriptName = fontInfo.name.toLowerCase();
            var familyName = fontInfo.family.toLowerCase();
            var styleName = fontInfo.style.toLowerCase();

            if (!matchesKeywordQuery(keywordQuery, postscriptName, familyName, styleName)) continue;

            /* ウェイト・種類フィルター：選択のある項目だけ絞り込む / Apply only the filters in use */
            if (useWeightFilter && !weightFilters[getWeightCategory(fontInfo)]) continue;
            if (useTypeFilter && !matchesTypeFilters(styleName, typeFilters)) continue;

            /* 同じ PostScript 名は1回だけ / Keep each PostScript name once */
            if (seenNames.hasOwnProperty(fontInfo.name)) continue;
            seenNames[fontInfo.name] = true;

            if (!groupedFonts.hasOwnProperty(fontInfo.family)) groupedFonts[fontInfo.family] = [];
            groupedFonts[fontInfo.family].push(fontInfo);
        }

        return groupedFonts;
    }

    /**
     * 各ファミリーをウェイト＋スタイル順に並べ替える（評価値は sortScore に控える）
     * @param {object} groupedFonts - ファミリー名をキーにしたフォント情報の配列
     * @returns {void}
     */
    function sortFontGroups(groupedFonts) {
        for (var familyName in groupedFonts) {
            if (!groupedFonts.hasOwnProperty(familyName)) continue;

            /* 評価値を先に1回だけ求めてから並べ替える / Score each font once, then sort */
            var familyFonts = groupedFonts[familyName];
            for (var i = 0; i < familyFonts.length; i++) {
                familyFonts[i].sortScore = getFontSortScore(familyFonts[i]);
            }
            /* 同点は PostScript 名順（並びを環境で変えない）/ Break ties by PostScript name */
            familyFonts.sort(function(a, b) {
                if (a.sortScore !== b.sortScore) return a.sortScore - b.sortScore;
                return (a.name < b.name) ? -1 : (a.name > b.name) ? 1 : 0;
            });
        }
    }

    /**
     * グループの合計フォント数を数える
     * @param {object} groupedFonts - ファミリー名をキーにしたフォント情報の配列
     * @returns {number} フォントの総数
     */
    function countFonts(groupedFonts) {
        var totalCount = 0;
        for (var familyName in groupedFonts) {
            if (!groupedFonts.hasOwnProperty(familyName)) continue;
            totalCount += groupedFonts[familyName].length;
        }
        return totalCount;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 表示するテキスト内容を決定する
     * @param {object} fontInfo - readFontInfo() が返したフォント情報
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @returns {string} 描画する文字列
     */
    function getDisplayText(fontInfo, userInput) {
        var displayMode = userInput.displayMode;
        if (displayMode === "postscript") return fontInfo.name;
        if (displayMode === "alphabet") return SAMPLE_ALPHABET_TEXT;
        if (displayMode === "numbers") return SAMPLE_NUMBERS_TEXT;
        if (displayMode === "custom") return userInput.customText || "";

        /* family+style（既定）/ family+style (default) */
        var displayText = fontInfo.family + (fontInfo.style ? " " + fontInfo.style : "");
        if (userInput.showScore) displayText += " (" + fontInfo.sortScore + ")";
        return displayText;
    }

    /**
     * サンプル用のテキストフレームを作成して配置する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} contents - 流し込む文字列
     * @param {number} left - 左端の座標
     * @param {number} top - 上端の座標
     * @param {TextFont} [textFont] - 適用するフォント（省略時は既定フォント）
     * @returns {TextFrame} 作成したテキストフレーム
     */
    function addSampleFrame(doc, contents, left, top, textFont) {
        var textFrame = doc.textFrames.add();
        textFrame.contents = contents;
        textFrame.textRange.characterAttributes.size = SAMPLE_FONT_SIZE;
        if (textFont) {
            try {
                textFrame.textRange.characterAttributes.textFont = textFont;
            } catch (e) {
                /* 適用できなかった枠は残さない / Do not leave the failed frame behind */
                textFrame.remove();
                throw e;
            }
        }
        textFrame.left = left;
        textFrame.top = top;
        textFrame.selected = true;
        return textFrame;
    }

    /**
     * ファミリー名を並べ替えて取得する（空のファミリーは除く）
     * @param {object} groupedFonts - ファミリー名をキーにしたフォント情報の配列
     * @returns {string[]} 並べ替え済みのファミリー名
     */
    function getSortedGroupLabels(groupedFonts) {
        var groupLabels = [];
        for (var familyName in groupedFonts) {
            if (!groupedFonts.hasOwnProperty(familyName)) continue;
            if (groupedFonts[familyName].length > 0) groupLabels.push(familyName);
        }
        groupLabels.sort();
        return groupLabels;
    }

    /**
     * 1ファミリー分の見出しとウェイト一覧を縦に描画し、占める範囲を測る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} familyName - ファミリー名
     * @param {object[]} familyFonts - 並べ替え済みのフォント情報
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @param {number} left - 左端の座標
     * @param {number} top - 上端の座標
     * @returns {object} frames（作成した枠）と width / height を持つブロック
     */
    function drawFamilyBlock(doc, familyName, familyFonts, userInput, left, top) {
        var blockFrames = [];
        var blockWidth = 0;
        var currentTop = top;

        var headingText = "[" + familyName + "]" + (userInput.showWeightCount ? " (" + familyFonts.length + ")" : "");
        var headingFrame = addSampleFrame(doc, headingText, left, currentTop);
        blockFrames.push(headingFrame);
        blockWidth = headingFrame.width;
        currentTop -= headingFrame.height + SAMPLE_FONT_SIZE * 0.5;

        for (var i = 0; i < familyFonts.length; i++) {
            var fontInfo = familyFonts[i];
            try {
                var sampleFrame = addSampleFrame(doc, getDisplayText(fontInfo, userInput), left, currentTop, fontInfo.textFont);
                blockFrames.push(sampleFrame);
                blockWidth = Math.max(blockWidth, sampleFrame.width);
                currentTop -= sampleFrame.height;
            } catch (e) {
                /* 適用できないフォントは飛ばして続行 / Skip fonts that cannot be applied */
                $.writeln("描画失敗：" + fontInfo.name + " → " + e);
            }
        }

        return { frames: blockFrames, width: blockWidth, height: top - currentTop };
    }

    /**
     * 同じ位置に描いたファミリーのブロックを、列ごとの最大幅・行ごとの最大高さで格子状に並べ直す
     * @param {object[]} familyBlocks - drawFamilyBlock() が返したブロック
     * @param {number} columnCount - 列数
     * @returns {void}
     */
    function arrangeFamilyBlocks(familyBlocks, columnCount) {
        var columnWidths = [];
        var rowHeights = [];
        var i, j;

        for (i = 0; i < familyBlocks.length; i++) {
            var columnIndex = i % columnCount;
            var rowIndex = Math.floor(i / columnCount);
            columnWidths[columnIndex] = Math.max(columnWidths[columnIndex] || 0, familyBlocks[i].width);
            rowHeights[rowIndex] = Math.max(rowHeights[rowIndex] || 0, familyBlocks[i].height);
        }

        /* 各列・各行の開始位置（描画位置からのずれ）/ Offset of each column and row from the drawing origin */
        var columnOffsets = [0];
        for (i = 1; i < columnWidths.length; i++) {
            columnOffsets[i] = columnOffsets[i - 1] + columnWidths[i - 1] + SAMPLE_COLUMN_GAP;
        }
        var rowOffsets = [0];
        for (i = 1; i < rowHeights.length; i++) {
            rowOffsets[i] = rowOffsets[i - 1] + rowHeights[i - 1] + SAMPLE_ROW_GAP;
        }

        for (i = 0; i < familyBlocks.length; i++) {
            var dx = columnOffsets[i % columnCount];
            var dy = -rowOffsets[Math.floor(i / columnCount)];
            if (dx === 0 && dy === 0) continue;

            var blockFrames = familyBlocks[i].frames;
            for (j = 0; j < blockFrames.length; j++) {
                blockFrames[j].translate(dx, dy);
            }
        }
    }

    /**
     * ファミリーごとにウェイトを一覧描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} groupedFonts - ファミリー名をキーにしたフォント情報の配列
     * @param {string[]} groupLabels - 並べ替え済みのファミリー名
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @param {number} startX - 描画開始位置の左端
     * @param {number} startY - 描画開始位置の上端
     * @returns {void}
     */
    function drawWeightSamples(doc, groupedFonts, groupLabels, userInput, startX, startY) {
        /* いったん同じ位置に描いて寸法を測り、あとで格子に並べる / Draw everything at the origin, measure, then lay out */
        var familyBlocks = [];
        for (var i = 0; i < groupLabels.length; i++) {
            familyBlocks.push(drawFamilyBlock(doc, groupLabels[i], groupedFonts[groupLabels[i]], userInput, startX, startY));
        }
        arrangeFamilyBlocks(familyBlocks, userInput.columns);
    }

    /**
     * ファミリーごとに代表フォント1つだけを1行で描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} groupedFonts - ファミリー名をキーにしたフォント情報の配列
     * @param {string[]} groupLabels - 並べ替え済みのファミリー名
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @param {number} startX - 描画開始位置の左端
     * @param {number} startY - 描画開始位置の上端
     * @returns {void}
     */
    function drawCategorySamples(doc, groupedFonts, groupLabels, userInput, startX, startY) {
        for (var i = 0; i < groupLabels.length; i++) {
            var familyFonts = groupedFonts[groupLabels[i]];

            var headingText = groupLabels[i] + (userInput.showWeightCount ? " (" + familyFonts.length + ")" : "");
            var top = startY - i * CATEGORY_LINE_HEIGHT;

            /* 見出しは最も細いウェイトで組む（並べ替え済みなので先頭が最小）/ Heading uses the lightest weight; groups are pre-sorted */
            try {
                addSampleFrame(doc, headingText, startX, top, familyFonts[0].textFont);
            } catch (e) {
                /* フォントを適用できなくても見出しは残す / Keep the heading even if the font cannot be applied */
                $.writeln("カテゴリフォント適用失敗：" + familyFonts[0].name + " → " + e);
                addSampleFrame(doc, headingText, startX, top);
            }
        }
    }

    /**
     * アートボード上にフォントサンプルを描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} groupedFonts - ファミリー名をキーにしたフォント情報の配列
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @returns {void}
     */
    function drawFontSamples(doc, groupedFonts, userInput) {
        var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var artboardRect = activeArtboard.artboardRect;
        var startX = artboardRect[0] + ARTBOARD_PADDING;
        var startY = artboardRect[1] - ARTBOARD_PADDING;

        app.executeMenuCommand("deselectall");

        var groupLabels = getSortedGroupLabels(groupedFonts);

        if (userInput.showWeight) {
            drawWeightSamples(doc, groupedFonts, groupLabels, userInput, startX, startY);
        } else {
            drawCategorySamples(doc, groupedFonts, groupLabels, userInput, startX, startY);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトの入口。ダイアログを表示し、フォントを収集して描画する
     * @returns {void}
     */
    function main() {
        try {
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }

            var doc = app.activeDocument;
            var userInput = showFontListDialog();
            if (!userInput) return;

            var keywordQuery = parseKeywordQuery(userInput.keyword);

            /* キーワード・ウェイト・種類のどれでも絞り込まないときは全フォントが対象になるため確認する
               Confirm when neither the keyword nor any filter narrows the list */
            var isUnfiltered = keywordQuery.andGroups.length === 0 &&
                !hasAnyFilterSelected(userInput.weightFilters) &&
                !hasAnyFilterSelected(userInput.typeFilters);
            if (isUnfiltered && !confirmShowAllFonts()) return;

            var groupedFonts = collectFonts(keywordQuery, userInput);
            sortFontGroups(groupedFonts);

            if (countFonts(groupedFonts) === 0) {
                alert(getLabel(LABELS.alert.noMatchingFont));
                return;
            }

            drawFontSamples(doc, groupedFonts, userInput);
        } catch (e) {
            alert(getLabel(LABELS.alert.errorOccurred) + e.message);
        }
    }

    main(); /* スクリプト実行開始 / Start */

})();
