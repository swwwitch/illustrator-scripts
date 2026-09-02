#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorで利用できるフォントをウェイト・スタイル順に並べ、ファミリー単位でアートボードへ整列描画します。
キーワード、ウェイト（5段階）、種類で対象を絞り込み、フォント名・PostScript名・サンプル・カスタムテキストから出力内容を選べます。

詳細は README を参照してください。

### Overview

Lays out the fonts available in Illustrator on the artboard, grouped by family and ordered by weight and style.
Narrow the list by keyword, weight (5 ranks) or style category, and output the font name, the PostScript name, a sample string, or your own custom text.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TypefaceSampler";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypefaceSampler.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypefaceSampler.md
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
     * @param {Array} radioButtons - 対象のラジオボタン
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
    var ARTBOARD_PADDING      = 20;  /* アートボード端からの余白（pt）/ padding from the artboard edge */
    var SAMPLE_COLUMN_SPACING = 220; /* カテゴリー列の間隔（pt）/ gap between category columns */
    var SAMPLE_ROW_SPACING    = 300; /* カテゴリー行の間隔（pt）/ gap between category rows */
    var CATEGORY_LINE_HEIGHT  = 16;  /* カテゴリー一覧の行送り（pt）/ line height of the category list */

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
        fieldLabel: {
            keyword: { ja: "フォント名に含まれるキーワード（空欄→全対象）", en: "Keyword in font name (leave blank for all)" },
            columns: { ja: "列数", en: "Columns" }
        },
        button: {
            cancel:  { ja: "キャンセル", en: "Cancel" },
            ok:      { ja: "OK", en: "OK" },
            stop:    { ja: "中止する", en: "Cancel" },
            proceed: { ja: "続行する", en: "Proceed" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noMatchingFont: { ja: "条件に該当するフォントが見つかりませんでした。", en: "No font matched the given conditions." },
            errorOccurred: { ja: "エラーが発生しました：", en: "An error occurred:" },
            confirmAllFonts: {
                ja: "すべてのフォントを対象に実行しますか？",
                en: "Do you want to process all fonts?"
            },
            confirmAllFontsNote: {
                ja: "非常に時間がかかることがあります。",
                en: "This may take a long time."
            }
        },
        sampleText: {
            ja: "愛のあるユニークで豊かな書体ABCabcGg349",
            en: "Lorem ipsum dolor sit amet, consectetur adipiscing elit"
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
        keywordField.characters = KEYWORD_FIELD_CHARS;
        keywordField.active = true;

        /* 出力内容 / Output content */
        var outputPanel = addPanel(dialogWindow, getLabel(LABELS.panel.output));
        var displayModeColumn = addLeftAlignedColumn(outputPanel);

        var displayModeRadios = [];
        displayModeRadios[0] = displayModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.fontNameWeightStyle));
        displayModeRadios[1] = displayModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.postscriptName));
        displayModeRadios[2] = displayModeColumn.add("radiobutton", undefined, SAMPLE_ALPHABET_TEXT);
        displayModeRadios[3] = displayModeColumn.add("radiobutton", undefined, SAMPLE_NUMBERS_TEXT);
        displayModeRadios[4] = displayModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.custom));
        displayModeRadios[0].value = true;

        var customTextField = outputPanel.add("edittext", undefined, getLabel(LABELS.sampleText));
        customTextField.characters = KEYWORD_FIELD_CHARS;
        customTextField.enabled = false;

        enableArrowKeyNavigation(displayModeRadios);

        /* 表示オプション / Display options */
        var optionPanel = addPanel(dialogWindow, getLabel(LABELS.panel.option));

        var weightOptionRow = optionPanel.add("group");
        setupRow(weightOptionRow);
        var showWeightCountCheckbox = weightOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.showWeightCount));
        var showWeightListCheckbox = weightOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.showWeightList));
        showWeightListCheckbox.value = true;

        var columnAndScoreRow = optionPanel.add("group");
        setupRow(columnAndScoreRow);
        columnAndScoreRow.add("statictext", undefined, labelText(LABELS.fieldLabel.columns));
        var columnField = columnAndScoreRow.add("edittext", undefined, DEFAULT_COLUMN_COUNT + "");
        columnField.characters = COLUMN_FIELD_CHARS;
        changeValueByArrowKey(columnField, null, 1);
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
            keyword: keywordField.text.toLowerCase().replace(/^\s+|\s+$/g, ""),
            displayMode: displayMode,
            customText: customTextField.text,
            columns: parseInt(columnField.text, 10) || DEFAULT_COLUMN_COUNT,
            useCategory: true,
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
        ["regular", "roman", "normal", "レギュラー", "r"], // +12
        ["rb"], // +13
        ["medium", "md", "ミディアム", "m"], // +14
        ["semibold", "semi bold", "sb"], // +15
        ["demibold", "demi bold", "db", "デミボールド", "demi", "d", "demixtra"], // +16
        ["bold", "bd", "ボールド", "b"], // +17
        ["extrabold", "extra bold", "xbold", "エクストラボールド", "e", "eb", "xb"], // +18
        ["heavy", "h"], // +19
        ["black"], // +20
        ["xblack", "extra black", "extrablack", "xb"], // +21
        ["ultra", "u", "ub", "ultra black", "ultrablack"] // +22
    ];

    /* 単独で使われたら Regular 扱いする装飾語句 / Decoration-only styles treated as Regular */
    var DECORATION_ONLY_STYLES = [
        "display", "compressed", "comp", "compact", "expanded", "extended", "semiextended",
        "ultracondensed", "extracondensed", "semicondensed", "cond", "condensed", "wide",
        "headline", "text", "low", "micro", "extra compressed",
        "semi expanded", "semiexpanded"
    ];

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
                termPatterns.push({
                    term: weightTerm,
                    groupIndex: i,
                    pattern: new RegExp("\\b" + weightTerm.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + "\\b")
                });
            }
        }
        termPatterns.sort(function(a, b) {
            return b.term.length - a.term.length;
        });
        return termPatterns;
    })();

    /**
     * スタイル文字列に一致する WEIGHT_GROUPS のインデックスを返す
     * 完全一致を優先し、なければ長い語から順に単語境界つきで照合する
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {number} 一致したインデックス。見つからない場合は -1
     */
    function getWeightGroupIndex(normalizedStyle) {
        var i, j;

        for (i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                if (normalizedStyle === WEIGHT_GROUPS[i][j]) return i;
            }
        }

        for (i = 0; i < WEIGHT_TERM_PATTERNS.length; i++) {
            if (WEIGHT_TERM_PATTERNS[i].pattern.test(normalizedStyle)) return WEIGHT_TERM_PATTERNS[i].groupIndex;
        }

        return -1;
    }

    /**
     * スタイル文字列を照合用に正規化する
     * @param {string} rawStyle - font.style の値
     * @returns {string} 小文字化し、区切り記号を空白に置き換えた文字列
     */
    function normalizeStyle(rawStyle) {
        return (rawStyle || "").toLowerCase().replace(/[_\-]+/g, " ").replace(/^\s+|\s+$/g, "");
    }

    // =========================================
    // ウェイト評価 / Weight scoring
    // =========================================

    /**
     * スタイル文字列に対する基本ウェイトスコアを取得する
     * @param {string} rawStyle - font.style の値
     * @param {string} postscriptName - 小文字化した PostScript 名
     * @param {string} familyName - 小文字化したファミリー名
     * @returns {number} ウェイトの評価値（小さいほど細い）
     */
    function getBaseWeightScore(rawStyle, postscriptName, familyName) {
        var normalizedStyle = normalizeStyle(rawStyle);
        var styleWords = normalizedStyle.split(/\s+/);
        var i;

        var applyFrutigerCorrection = (/frutiger/i.test(familyName) && /ultralight/.test(normalizedStyle));

        /* W0〜W9 */
        var singleDigitMatch = normalizedStyle.match(/^w(\d)$/);
        if (singleDigitMatch !== null) return parseInt(singleDigitMatch[1], 10);

        /* W000〜W999 */
        var tripleDigitMatch = normalizedStyle.match(/^w(\d{3})$/);
        if (tripleDigitMatch !== null) return parseInt(tripleDigitMatch[1], 10);

        /* 先頭数値（例：25 Ultra Light）/ Leading number */
        var leadingNumberMatch = normalizedStyle.match(/^(\d{1,3})(?=\D|$)/);
        if (leadingNumberMatch) return parseInt(leadingNumberMatch[1], 10);

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
     * @param {TextFont} font - 対象のフォント
     * @returns {number} 並べ替えに使う評価値
     */
    function getFontSortScore(font) {
        var styleName = (font.style || "").toLowerCase();
        var postscriptName = (font.name || "").toLowerCase();
        var familyName = (font.family || "").toLowerCase();

        /* 特例：PostScript名が「FuturaPT-Heavy」なら 1015 固定（加点処理なし）/ Fixed rank, no offsets */
        if (postscriptName === "futurapt-heavy") return 1015;

        var baseScore = getBaseWeightScore(styleName, postscriptName, familyName);
        var decorationOffset = 0;
        var styleWords = styleName.split(/\s+/);

        /* 装飾フラグ初期化 / Initialize decoration decorationFlags */
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
        if (decorationFlags.hasExtraCompressed) decorationOffset += 150; /* 特別加点 / Extra decorationOffset */
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
     * @param {TextFont} font - 判定対象のフォント
     * @returns {string} veryThin / light / regular / semiBold / bold のいずれか
     */
    function getWeightCategory(font) {
        var normalizedStyle = normalizeStyle(font.style);

        /* W3、W600、25 Ultra Light のような数値スタイル / Numeric styles */
        var numericMatch = normalizedStyle.match(/^w?(\d{1,3})(?=\D|$)/);
        if (numericMatch) {
            return getWeightCategoryFromNumber(parseInt(numericMatch[1], 10), numericMatch[1].length);
        }

        var groupIndex = getWeightGroupIndex(normalizedStyle);
        if (groupIndex === -1) return "regular";
        return getWeightCategoryFromGroupIndex(groupIndex);
    }

    // =========================================
    // フォント収集 / Font collection
    // =========================================

    /**
     * 引用符つき検索用に文字列を正規化する（空白・ハイフンを除去）
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 正規化した文字列
     */
    function normalizeForSearch(sourceText) {
        return sourceText.toLowerCase().replace(/[\s\-　]/g, "");
    }

    /**
     * キーワード入力を AND / OR / NOT の検索条件に分解する
     * @param {string} rawKeyword - ダイアログで入力されたキーワード
     * @returns {object} andGroups（ORの配列をANDで並べたもの）と notKeywords を持つ検索条件
     */
    function parseKeywordQuery(rawKeyword) {
        var keywordQuery = { andGroups: [], notKeywords: [] };
        if (!rawKeyword) return keywordQuery;

        /* 入力の正規化 / Normalize the input */
        var normalizedKeyword = rawKeyword
            .replace(/　/g, " ")        /* 全角スペース→半角 / Full-width space to half-width */
            .replace(/\s*,\s*/g, ",")   /* カンマの前後スペース除去 / Trim around commas */
            .replace(/\s+/g, " ");      /* 連続スペース→1つ / Collapse spaces */

        /* NOTキーワード（-付き）を抜き出す / Extract NOT keywords */
        var spaceSeparatedParts = normalizedKeyword.split(" ");
        var includeParts = [];
        var i, j;
        for (i = 0; i < spaceSeparatedParts.length; i++) {
            if (spaceSeparatedParts[i].charAt(0) === "-") {
                keywordQuery.notKeywords.push(spaceSeparatedParts[i].substring(1).toLowerCase());
            } else {
                includeParts.push(spaceSeparatedParts[i]);
            }
        }

        /* AND（+区切り）と OR（スペース・カンマ）を組み立てる / Build AND groups of OR terms */
        var andSegments = includeParts.join(" ").split("+");
        for (i = 0; i < andSegments.length; i++) {
            var orGroup = [];
            var orTerms = andSegments[i].split(/[\s,]+/);

            for (j = 0; j < orTerms.length; j++) {
                var termText = orTerms[j].toLowerCase();
                if (termText.length === 0) continue;

                var isPrefix = false;
                if (termText.charAt(0) === "^") {
                    isPrefix = true;
                    termText = termText.substring(1);
                }

                var isQuoted = false;
                if (termText.charAt(0) === '"' && termText.charAt(termText.length - 1) === '"') {
                    termText = termText.slice(1, -1);
                    isQuoted = true;
                }

                orGroup.push({ keyword: termText, isPrefix: isPrefix, isQuoted: isQuoted });
            }

            if (orGroup.length > 0) keywordQuery.andGroups.push(orGroup);
        }

        return keywordQuery;
    }

    /**
     * 検索語1つがフォント名・ファミリー名・スタイル名のいずれかに合致するか判定する
     * @param {object} searchTerm - parseKeywordQuery() が組み立てた検索語
     * @param {string} postscriptName - 小文字化した PostScript 名
     * @param {string} familyName - 小文字化したファミリー名
     * @param {string} styleName - 小文字化したスタイル名
     * @returns {boolean} 合致すれば true
     */
    function matchesKeywordTerm(searchTerm, postscriptName, familyName, styleName) {
        var termText = searchTerm.keyword;

        if (searchTerm.isPrefix) {
            return postscriptName.substr(0, termText.length) === termText ||
                familyName.substr(0, termText.length) === termText ||
                styleName.substr(0, termText.length) === termText;
        }

        if (searchTerm.isQuoted) {
            var normalizedTerm = normalizeForSearch(termText);
            return normalizeForSearch(postscriptName).indexOf(normalizedTerm) !== -1 ||
                normalizeForSearch(familyName).indexOf(normalizedTerm) !== -1 ||
                normalizeForSearch(styleName).indexOf(normalizedTerm) !== -1;
        }

        return postscriptName.indexOf(termText) !== -1 ||
            familyName.indexOf(termText) !== -1 ||
            styleName.indexOf(termText) !== -1;
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

        /* NOT条件：含まれていたら除外 / Exclude when a NOT keyword matches */
        for (i = 0; i < keywordQuery.notKeywords.length; i++) {
            var excludeTerm = keywordQuery.notKeywords[i];
            if (postscriptName.indexOf(excludeTerm) !== -1 || familyName.indexOf(excludeTerm) !== -1 || styleName.indexOf(excludeTerm) !== -1) {
                return false;
            }
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
        /* 基本 / Basic */
        if (typeFilters.basic &&
            (styleName.indexOf("text") !== -1 || styleName.indexOf("headline") !== -1)) return true;

        /* 狭める系 / Condensed */
        if (typeFilters.narrow &&
            (styleName.indexOf("cond") !== -1 || styleName.indexOf("cn") !== -1 ||
                styleName.indexOf("compressed") !== -1 || styleName.indexOf("comp") !== -1)) return true;

        /* 広げる系 / Expanded */
        if (typeFilters.wide &&
            (styleName.indexOf("expanded") !== -1 || styleName.indexOf("extended") !== -1)) return true;

        /* 装飾・特殊用途 / Display */
        if (typeFilters.decor &&
            (styleName.indexOf("compact") !== -1 || styleName.indexOf("display") !== -1)) return true;

        /* サイズ・プロポーション系 / Size and proportion */
        if (typeFilters.sizeProp &&
            (styleName.indexOf("micro") !== -1 || styleName.indexOf("low") !== -1 ||
                styleName.indexOf("wide") !== -1)) return true;

        return false;
    }

    /**
     * 重複を避けてグループにフォントを追加する
     * @param {object} groupedFonts - カテゴリー名をキーにしたフォントの入れ物
     * @param {string} groupKey - 追加先のカテゴリー名
     * @param {TextFont} font - 追加するフォント
     * @returns {void}
     */
    function addFontToGroup(groupedFonts, groupKey, font) {
        if (!groupedFonts[groupKey]) groupedFonts[groupKey] = [];

        var groupFonts = groupedFonts[groupKey];
        for (var i = 0; i < groupFonts.length; i++) {
            if (groupFonts[i].name === font.name) return;
        }
        groupFonts.push(font);
    }

    /**
     * 条件に合致するフォントを収集し、カテゴリー単位にまとめる
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @returns {object} カテゴリー名をキーにしたフォントの配列
     */
    function collectFonts(userInput) {
        var groupedFonts = {};
        var keywordQuery = parseKeywordQuery(userInput.keyword);
        var weightFilters = userInput.weightFilters;
        var typeFilters = userInput.typeFilters;
        var useWeightFilter = hasAnyFilterSelected(weightFilters);
        var useTypeFilter = hasAnyFilterSelected(typeFilters);

        for (var i = 0; i < textFonts.length; i++) {
            var font = textFonts[i];
            var postscriptName = font.name.toLowerCase();
            var familyName = font.family.toLowerCase();
            var styleName = (font.style || "").toLowerCase();

            if (!matchesKeywordQuery(keywordQuery, postscriptName, familyName, styleName)) continue;

            /* ウェイト・種類フィルター：選択のある項目だけ絞り込む / Apply only the filters in use */
            if (useWeightFilter && !weightFilters[getWeightCategory(font)]) continue;
            if (useTypeFilter && !matchesTypeFilters(styleName, typeFilters)) continue;

            addFontToGroup(groupedFonts, userInput.useCategory ? font.family : "Uncategorized", font);
        }

        return groupedFonts;
    }

    /**
     * 各グループをウェイト＋スタイル順に並べ替える
     * @param {object} groupedFonts - カテゴリー名をキーにしたフォントの配列
     * @returns {void}
     */
    function sortFontGroups(groupedFonts) {
        for (var groupLabel in groupedFonts) {
            if (!groupedFonts.hasOwnProperty(groupLabel)) continue;

            /* 評価値を先に1回だけ求めてから並べ替える / Score each font once, then sort */
            var groupFonts = groupedFonts[groupLabel];
            var scoredFonts = [];
            var i;
            for (i = 0; i < groupFonts.length; i++) {
                scoredFonts.push({ font: groupFonts[i], rank: getFontSortScore(groupFonts[i]) });
            }
            scoredFonts.sort(function(a, b) {
                return a.rank - b.rank;
            });
            for (i = 0; i < scoredFonts.length; i++) {
                groupFonts[i] = scoredFonts[i].font;
            }
        }
    }

    /**
     * グループの合計フォント数を数える
     * @param {object} groupedFonts - カテゴリー名をキーにしたフォントの配列
     * @returns {number} フォントの総数
     */
    function countFonts(groupedFonts) {
        var totalCount = 0;
        for (var groupLabel in groupedFonts) {
            if (!groupedFonts.hasOwnProperty(groupLabel)) continue;
            totalCount += groupedFonts[groupLabel].length;
        }
        return totalCount;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 表示するテキスト内容を決定する
     * @param {TextFont} font - 対象のフォント
     * @param {string} displayMode - 出力モード（DISPLAY_MODES のいずれか）
     * @param {string} customText - カスタムテキスト
     * @param {boolean} showScore - スコアを併記するか
     * @returns {string} 描画する文字列
     */
    function getDisplayText(font, displayMode, customText, showScore) {
        if (displayMode === "postscript") return font.name;
        if (displayMode === "alphabet") return SAMPLE_ALPHABET_TEXT;
        if (displayMode === "numbers") return SAMPLE_NUMBERS_TEXT;
        if (displayMode === "custom") return customText || "";

        /* family+style（既定）/ family+style (default) */
        var displayText = font.family + (font.style ? " " + font.style : "");
        if (showScore) displayText += " (" + getFontSortScore(font) + ")";
        return displayText;
    }

    /**
     * サンプル用のテキストフレームを作成して配置する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} contents - 流し込む文字列
     * @param {number} left - 左端の座標
     * @param {number} top - 上端の座標
     * @param {TextFont} [font] - 適用するフォント（省略時は既定フォント）
     * @returns {TextFrame} 作成したテキストフレーム
     */
    function addSampleFrame(doc, contents, left, top, font) {
        var textFrame = doc.textFrames.add();
        textFrame.contents = contents;
        textFrame.textRange.characterAttributes.size = SAMPLE_FONT_SIZE;
        if (font) textFrame.textRange.characterAttributes.textFont = font;
        textFrame.left = left;
        textFrame.top = top;
        textFrame.selected = true;
        return textFrame;
    }

    /**
     * カテゴリー名を並べ替えて取得する（空のカテゴリーは除く）
     * @param {object} groupedFonts - カテゴリー名をキーにしたフォントの配列
     * @returns {Array<string>} 並べ替え済みのカテゴリー名
     */
    function getSortedGroupLabels(groupedFonts) {
        var groupLabels = [];
        for (var groupLabel in groupedFonts) {
            if (!groupedFonts.hasOwnProperty(groupLabel)) continue;
            if (groupedFonts[groupLabel].length > 0) groupLabels.push(groupLabel);
        }
        groupLabels.sort();
        return groupLabels;
    }

    /**
     * カテゴリーごとにウェイトを一覧描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} groupedFonts - カテゴリー名をキーにしたフォントの配列
     * @param {Array<string>} groupLabels - 並べ替え済みのカテゴリー名
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @param {number} startX - 描画開始位置の左端
     * @param {number} startY - 描画開始位置の上端
     * @returns {void}
     */
    function drawWeightSamples(doc, groupedFonts, groupLabels, userInput, startX, startY) {
        var columnIndex = 0;
        var rowIndex = 0;

        for (var i = 0; i < groupLabels.length; i++) {
            var groupLabel = groupLabels[i];
            var groupFonts = groupedFonts[groupLabel];

            var left = startX + columnIndex * SAMPLE_COLUMN_SPACING;
            var top = startY - rowIndex * SAMPLE_ROW_SPACING;

            if (++columnIndex >= userInput.columns) {
                columnIndex = 0;
                rowIndex++;
            }

            if (userInput.useCategory) {
                var headingText = "[" + groupLabel + "]" + (userInput.showWeightCount ? " (" + groupFonts.length + ")" : "");
                var headingFrame = addSampleFrame(doc, headingText, left, top);
                top -= headingFrame.height + SAMPLE_FONT_SIZE * 0.5;
            }

            for (var j = 0; j < groupFonts.length; j++) {
                var font = groupFonts[j];
                var sampleText = getDisplayText(font, userInput.displayMode, userInput.customText, userInput.showScore);
                try {
                    var sampleFrame = addSampleFrame(doc, sampleText, left, top, font);
                    top -= sampleFrame.height;
                } catch (e) {
                    /* 適用できないフォントは飛ばして続行 / Skip groupFonts that cannot be applied */
                    $.writeln("描画失敗：" + font.name + " → " + e);
                }
            }
        }
    }

    /**
     * カテゴリーごとに代表フォント1つだけを1行で描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} groupedFonts - カテゴリー名をキーにしたフォントの配列
     * @param {Array<string>} groupLabels - 並べ替え済みのカテゴリー名
     * @param {object} userInput - ダイアログで取得したユーザー入力
     * @param {number} startX - 描画開始位置の左端
     * @param {number} startY - 描画開始位置の上端
     * @returns {void}
     */
    function drawCategorySamples(doc, groupedFonts, groupLabels, userInput, startX, startY) {
        for (var i = 0; i < groupLabels.length; i++) {
            var groupLabel = groupLabels[i];
            var groupFonts = groupedFonts[groupLabel];

            var headingText = groupLabel + (userInput.showWeightCount ? " (" + groupFonts.length + ")" : "");
            var top = startY - i * CATEGORY_LINE_HEIGHT;

            /* 見出しは最も細いウェイトで組む（並べ替え済みなので先頭が最小）/ Heading uses the lightest weight; groups are pre-sorted */
            try {
                addSampleFrame(doc, headingText, startX, top, groupFonts[0]);
            } catch (e) {
                /* フォントを適用できなくても見出しは残す / Keep the heading even if the font cannot be applied */
                $.writeln("カテゴリフォント適用失敗：" + groupFonts[0].name + " → " + e);
                addSampleFrame(doc, headingText, startX, top);
            }
        }
    }

    /**
     * アートボード上にフォントサンプルを描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} groupedFonts - カテゴリー名をキーにしたフォントの配列
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

            /* キーワード未入力なら全フォントが対象になるため確認する / Confirm when no keyword narrows the list */
            if (!userInput.keyword && !confirmShowAllFonts()) return;

            var groupedFonts = collectFonts(userInput);
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