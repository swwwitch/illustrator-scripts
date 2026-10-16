#target illustrator
#targetengine "VerticalTextFormatterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストを縦組みにし、欧文回転・縦中横・数字の書き換え（全角・漢数字・和暦）・禁則・文字組みなどをまとめて適用します。
ダイアログボックスで項目を選び、変換の結果を一覧で確かめてから実行できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VerticalTextFormatter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ndb55c5516fd6

### Overview

Turns the selected text vertical and applies vertical-text settings in one go: rotated Roman characters,
Tate-Chu-Yoko, number rewriting (full-width digits, kanji numerals, Japanese era years), kinsoku and mojikumi.
Choose the options in a dialog box and check the result list before running.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/VerticalTextFormatter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "VerticalTextFormatter";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.18.0";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VerticalTextFormatter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/VerticalTextFormatter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ndb55c5516fd6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* ダイアログの初期値（前回の値は SettingsStore に残る） / Dialog defaults (the last values are kept by SettingsStore) */
    var DEFAULT_SETTINGS = {
        rotateRoman: true,
        tcyNumbers: true,
        tcyDigits2: true,
        tcyDigits3: true,
        tcyDigits4: false,
        tcyExcludeAdjacent: true,
        tcyCapitals: false,
        tcyExclamation: false,
        tcyGlyphByLength: true,
        singleDigitFullwidth: false,
        japaneseEra: false,
        kanji: false,
        kanjiMinDigits: 4,
        kanjiMaxDigits: 20,
        kanjiStyle: "digit",
        kanjiZeroRei: false,
        kanjiIchiSen: false,
        halfwidthAlnum: false,
        commaStyle: "touten",
        phoneHyphenStyle: "hyphen",
        colorCheck: false,
        areaAutoSize: true,
        areaHeightToSize: true,
        periodToNakaguro: true,
        symbolsFullwidth: true,
        quotesToChonchon: true,
        rangeToWave: true,
        exclamFullwidth: true,
        singleLetterFullwidth: false,
        unitsKumimoji: false,
        proportionalFullwidthLatin: true,
        rotateSkipLongLatin: false,
        rotateSkipMinChars: 5,
        shortWordFullwidth: false,
        shortWordMaxChars: 4,
        kinsokuType: "Soft_v2",
        mojikumiIndex: 6,
        mojikumiCustomName: "",
        leadingPercent: 175,
        dialogMode: "simple"
    };

    /* ダイアログの真偽値の項目（設定のキーとコントロール名が同じ） / Boolean dialog options (setting key = control name) */
    var CHECKBOX_SETTING_KEYS = ["rotateRoman", "tcyNumbers", "tcyDigits2", "tcyDigits3", "tcyDigits4", "tcyExcludeAdjacent", "tcyCapitals",
        "tcyExclamation", "tcyGlyphByLength", "singleDigitFullwidth", "japaneseEra", "kanji", "kanjiZeroRei", "kanjiIchiSen", "halfwidthAlnum",
        "colorCheck", "areaAutoSize", "areaHeightToSize", "periodToNakaguro", "symbolsFullwidth", "quotesToChonchon", "rangeToWave", "exclamFullwidth", "singleLetterFullwidth",
        "unitsKumimoji", "proportionalFullwidthLatin", "rotateSkipLongLatin", "shortWordFullwidth"];

    /* 簡易モードで表示していなくてもオフにしない項目（常にかける） / Options kept on even when hidden in the simple mode (always applied) */
    var ALWAYS_APPLIED_KEYS = ["areaHeightToSize"];

    var MODE_SWITCH_RESULT = 3;   /* モードを切り替えて開き直すときの show() の戻り値 / show() result for reopening in another mode */

    /* 禁則の選択肢（paragraphAttributes.kinsoku の値） / Kinsoku choices (values of paragraphAttributes.kinsoku) */
    var KINSOKU_TYPES = ["None", "Hard", "Soft", "Soft_v2"];
    /* 組み込みの文字組みアキ量設定の選択肢（doc.mojikumiSet の番号、-1 は「なし」）。組み込みの要素は内部 ID の文字列なので、表示名は LABELS に持つ。
       8番目（番号7）以降はカスタムのセットで、要素の文字列がそのまま名前（mojikumiCustomName に控える）
       Built-in mojikumi choices (doc.mojikumiSet index, -1 = none). Built-in entries are internal ID strings, so LABELS holds the names;
       entries from index 7 on are custom sets whose string is the name (kept in mojikumiCustomName) */
    var MOJIKUMI_INDEXES = [-1, 0, 1, 2, 3, 4, 5, 6];
    var MOJIKUMI_BUILTIN_COUNT = 7;
    /* 桁区切りの選択肢 / Thousands separator choices */
    var COMMA_STYLES = ["touten", "remove", "keep"];
    /* 番号のハイフンの選択肢 / Phone hyphen choices */
    var PHONE_HYPHEN_STYLES = ["hyphen", "choon", "keep"];

    /* 組み込みのプリセット。書いていない項目は DEFAULT_SETTINGS の値
       Built-in presets; omitted items use DEFAULT_SETTINGS.
       よりどころ：モリサワ「数字を含んだ文章はどのように組めばいいの？」、東洋美術印刷「縦組みの基本！」、
       校正視点「縦書き数字の表記方法」（A：漢数字が原則、B：算用数字が原則） */
    var KANJI_ALL_DIGITS = { tcyNumbers: false, kanji: true, kanjiMinDigits: 1, kanjiMaxDigits: 20 };   /* すべて漢数字にする共通の値 / shared values for all-kanji presets */
    var BUILTIN_PRESETS = [
        /* 初期値のまま：2〜3桁を縦中横、字数に合わせた等幅字形、「、」「・」 / Defaults: 2–3 digit TCY, fixed-width forms, 、 and ・ */
        { labelKey: "standard", data: {} },
        /* 校正視点 B：2桁は縦中横、1桁と3桁以上は1字ずつ縦に / Kousei B: TCY for 2 digits only, others upright one by one */
        { labelKey: "arabicTcy2", data: { tcyDigits3: false, singleDigitFullwidth: true } },
        /* 東洋美術印刷・モリサワ：2〜3桁を縦中横（3桁は三分）、単位は組文字、短い略称は全角で縦・長い欧文は横（縦中横は数字だけ）
           Toyobijutsu / Morisawa: 2–3 digit TCY, unit squares, short abbreviations full-width upright, long Roman text sideways (TCY for numbers only) */
        {
            labelKey: "arabicMagazine",
            data: {
                singleDigitFullwidth: true, unitsKumimoji: true,
                rotateSkipLongLatin: true, shortWordFullwidth: true
            }
        },
        /* 校正視点 A①：万・億・兆だけ（一般的な方法） / Kousei A-1: only 万・億・兆 (the common style) */
        { labelKey: "kanjiLargeUnits", data: mergeObjects(KANJI_ALL_DIGITS, { kanjiStyle: "largeUnits", commaStyle: "remove" }) },
        /* 校正視点 A②：十・百・千も入れる / Kousei A-2: every unit word */
        { labelKey: "kanjiPositional", data: mergeObjects(KANJI_ALL_DIGITS, { kanjiStyle: "positional", commaStyle: "remove" }) },
        /* 校正視点 A③：単位語なしで1字ずつ、3桁ごとに読点 / Kousei A-3: digit by digit with 、 every three digits */
        { labelKey: "kanjiDigitTouten", data: mergeObjects(KANJI_ALL_DIGITS, { kanjiStyle: "digit", commaStyle: "touten" }) },
        {
            labelKey: "rotateOnly",
            data: {
                tcyNumbers: false, tcyGlyphByLength: false, commaStyle: "keep", phoneHyphenStyle: "keep", periodToNakaguro: false,
                symbolsFullwidth: false, quotesToChonchon: false, rangeToWave: false, exclamFullwidth: false,
                proportionalFullwidthLatin: false
            }
        }
    ];

    /**
     * 2つのオブジェクトを重ねた新しいオブジェクトを返す（後ろが優先）
     * @param {Object} baseObject - 元の値
     * @param {Object} overrideObject - 上書きする値
     * @returns {Object} 重ねた値
     */
    function mergeObjects(baseObject, overrideObject) {
        var mergedObject = {};
        var key;
        for (key in baseObject) {
            if (baseObject.hasOwnProperty(key)) mergedObject[key] = baseObject[key];
        }
        for (key in overrideObject) {
            if (overrideObject.hasOwnProperty(key)) mergedObject[key] = overrideObject[key];
        }
        return mergedObject;
    }

    var PREVIEW_MAX_LINES = 300;   /* 確認の一覧に出す最大の行数 / max lines in the check list */

    /* 確認用の色分けの初期色（ColorNumberByWidth から移植） / Initial colors for the check coloring (ported from ColorNumberByWidth) */
    var CHECK_BASE_RGB         = [0, 0, 0];         /* ベース（ほかの文字） / base (other characters) */
    var CHECK_BASE_CMYK        = [0, 0, 0, 100];
    var CHECK_FULL_WIDTH_RGB   = [255, 0, 0];       /* 全角 / full-width */
    var CHECK_FULL_WIDTH_CMYK  = [0, 100, 100, 0];
    var CHECK_HALF_WIDTH_RGB   = [0, 153, 255];     /* 半角（#0099ff） / half-width */
    var CHECK_HALF_WIDTH_CMYK  = [100, 40, 0, 0];
    var CHECK_GLYPH_WIDTH_RGB  = [0, 170, 0];       /* 等幅字形 / fixed-width glyphs */
    var CHECK_GLYPH_WIDTH_CMYK = [80, 0, 100, 0];

    /* 全角：０-９ Ａ-Ｚ ａ-ｚ ！？、半角：0-9 A-Z a-z !? / full-width and half-width alphanumerics, ! and ? */
    var FULL_WIDTH_PATTERN = /[０-９Ａ-Ｚａ-ｚ！？]/;
    var HALF_WIDTH_PATTERN = /[0-9A-Za-z!?]/;

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

    var BUTTON_WIDTH        = 90;    /* OK／キャンセルの幅 / OK and Cancel width */
    var OPTION_INDENT       = 20;    /* 親のチェックボックスに従う項目の字下げ / indent for dependent options */
    var PREVIEW_LIST_HEIGHT = 142;   /* 確認の一覧の高さ / height of the check list */
    var PROGRESS_WIDTH      = 300;   /* プログレスバーの幅 / progress bar width */
    var PRESET_ICON_SIZE    = [20, 18];   /* プリセットの保存・削除アイコンの大きさ / preset save and delete icon size */
    var PRESET_DROPDOWN_WIDTH = 220; /* プリセットのドロップダウンの幅 / preset dropdown width */
    var PRESET_NAME_CHARS   = 20;    /* プリセット名の入力欄の文字数 / preset name field width */
    var COLOR_LABEL_WIDTH   = 60;         /* 確認用の色分けの項目名の幅（右揃え） / check-coloring label width (right-aligned) */
    var COLOR_SWATCH_SIZE   = [40, 18];   /* 確認用の色分けの色見本の大きさ / check-coloring swatch size */
    var PARAGRAPH_LABEL_WIDTH = 60;
    var PARAGRAPH_OPTION_TOP_MARGIN = 5;   /* 「文字・段落」のチェックボックスの上の余白 / gap above the checkboxes in Character and Paragraph */  /* 段落のドロップダウンの項目名の幅（右揃え） / paragraph dropdown label width (right-aligned) */

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
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
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
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
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
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
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
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
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

    // アイコンのボタン（再利用パーツ） / Icon buttons (reusable)

    var ICON_BUTTON_SIZE = [24, 22]; /* 既定の大きさ / default size */
    var ICON_BUTTON_UI_DARK = isDarkUI();
    var ICON_BUTTON_COLOR     = ICON_BUTTON_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* 絵の色 / icon color */
    var ICON_BUTTON_DIM_COLOR = ICON_BUTTON_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の色 / color when disabled */

    /**
     * onDraw で絵を描くアイコンのボタンを追加する。押すと onClick を呼ぶ（無効の間は押せず、薄く描く）
     * @param {Group} parent - 追加先
     * @param {number[]} iconSize - [幅, 高さ]
     * @param {Function} drawIcon - 絵を描く関数 (iconGraphics, iconWidth, iconHeight, iconColor)
     * @returns {Group} アイコン
     */
    function addIconButton(parent, iconSize, drawIcon) {
        var iconButton = parent.add("group");
        iconButton.preferredSize = iconSize;
        iconButton.minimumSize = iconSize;
        iconButton.maximumSize = iconSize;
        iconButton.onDraw = function () {
            var iconColor = isIconButtonEnabledInTree(iconButton) ? ICON_BUTTON_COLOR : ICON_BUTTON_DIM_COLOR;
            drawIcon(iconButton.graphics, iconSize[0], iconSize[1], iconColor);
        };
        iconButton.addEventListener("mousedown", function () {
            if (!isIconButtonEnabledInTree(iconButton)) return;
            if (typeof iconButton.onClick === "function") iconButton.onClick();
        });
        return iconButton;
    }

    /**
     * アイコンのボタンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} iconButton - addIconButton() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setIconButtonEnabled(iconButton, isEnabled) {
        if (iconButton.enabled === isEnabled) return;
        iconButton.enabled = isEnabled;
        /* group には notify() が無いため、隠して再表示して描き直させる / groups have no notify(), so hide and show to repaint */
        iconButton.hide();
        iconButton.show();
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isIconButtonEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * 絵の座標の長方形の並びを、縦横比を保って中央に置いて塗る（ScriptUI は多角形を塗れないため、斜めも長方形の並びで描く）
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} designSize - 絵の [幅, 高さ]
     * @param {number[][]} designRects - 長方形 [左, 上, 右, 下] の並び
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function fillIconButtonRects(iconGraphics, iconWidth, iconHeight, designSize, designRects, iconColor) {
        var iconScale = Math.min(iconWidth / designSize[0], iconHeight / designSize[1]);
        var originX = (iconWidth - designSize[0] * iconScale) / 2;
        var originY = (iconHeight - designSize[1] * iconScale) / 2;
        iconGraphics.newPath();
        for (var i = 0; i < designRects.length; i++) {
            var designRect = designRects[i];
            if (designRect[2] <= designRect[0] || designRect[3] <= designRect[1]) continue;
            iconGraphics.rectPath(originX + designRect[0] * iconScale, originY + designRect[1] * iconScale,
                (designRect[2] - designRect[0]) * iconScale, (designRect[3] - designRect[1]) * iconScale);
        }
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
    }

    /**
     * 保存アイコン（トレイに下向きの矢印）を描く。1223×993 の絵。
     * 矢じりとトレイの V 字の切り欠きは細い長方形を並べ、トレイの四角い穴は塗らずに残す
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawSaveIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var designWidth = 1223;
        var trayHole = [78, 688, 230, 840]; /* トレイの四角い穴 [左, 上, 右, 下] / the tray's square hole */
        var sliceCount = 16;
        var designRects = [[535, 0, 688, 383]]; /* 矢印の軸 / arrow shaft */
        var k;

        /** 穴と重なる部分を除いて長方形を足す / add a rectangle minus the tray hole */
        function addTrayRect(left, top, right, bottom) {
            if (right <= trayHole[0] || left >= trayHole[2] || bottom <= trayHole[1] || top >= trayHole[3]) {
                designRects.push([left, top, right, bottom]);
                return;
            }
            designRects.push([left, top, right, trayHole[1]], [left, trayHole[3], right, bottom],
                [left, Math.max(top, trayHole[1]), trayHole[0], Math.min(bottom, trayHole[3])],
                [trayHole[2], Math.max(top, trayHole[1]), right, Math.min(bottom, trayHole[3])]);
        }

        /* 下向きの矢じり（上の付け根から先端へ細っていく）/ the downward head */
        for (k = 0; k < sliceCount; k++) {
            var headHalf = 251 * (1 - (k + 0.5) / sliceCount);
            var headTop = 383 + (703 - 383) * k / sliceCount;
            designRects.push([612 - headHalf, headTop, 612 + headHalf, headTop + (703 - 383) / sliceCount]);
        }
        /* トレイ：上辺に V 字の切り欠き（幅 458 から下へ細り、856 で閉じる）/ tray with a V notch along the top */
        for (k = 0; k < sliceCount; k++) {
            var notchTop = 612 + (856 - 612) * k / sliceCount;
            var notchBottom = notchTop + (856 - 612) / sliceCount;
            var notchHalf = 229 * (1 - (k + 0.5) / sliceCount);
            addTrayRect(0, notchTop, 612 - notchHalf, notchBottom);
            addTrayRect(612 + notchHalf, notchTop, designWidth, notchBottom);
        }
        addTrayRect(0, 856, designWidth, 993);
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [designWidth, 993], designRects, iconColor);
    }

    /**
     * 削除アイコン（ゴミ箱）を描く。756×825 の絵
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawTrashIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [756, 825], [
            [206, 0, 550, 70], [206, 70, 275, 137], [481, 70, 550, 137],     /* 取っ手 / handle */
            [0, 137, 756, 207],                                               /* ふた / lid */
            [69, 207, 138, 825], [618, 207, 688, 825], [138, 756, 618, 825], /* 本体 / body */
            [206, 275, 275, 687], [343, 275, 413, 687], [481, 275, 550, 687] /* 縦の線 / ribs */
        ], iconColor);
    }

    // アイコンのボタン（再利用パーツ）ここまで / End of the reusable icon buttons

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
            title: { ja: "縦組みの整形", en: "Vertical Text Formatter" }
        },
        panel: {
            rotateTcy: { ja: "欧文回転・縦中横", en: "Rotation and Tate-Chu-Yoko" },
            replace: { ja: "文字の変換", en: "Character Conversion" },
            numbers: { ja: "数字", en: "Numbers" },
            paragraph: { ja: "文字・段落", en: "Character and Paragraph" },
            colorCheck: { ja: "確認用の色分け", en: "Check Coloring" },
            preview: { ja: "変換の確認", en: "Conversion Check" }
        },
        checkbox: {
            rotateRoman: { ja: "縦組み中の欧文回転", en: "Rotate Roman characters" },
            rotateSkipLongLatin: { ja: "長い欧文は除く", en: "Skip long Roman text" },
            shortWordFullwidth: { ja: "短い英単語を全角に", en: "Short words to full-width" },
            tcyNumbers: { ja: "数字を縦中横", en: "TCY numbers" },
            tcyDigits2: { ja: "2桁", en: "2 digits" },
            tcyDigits3: { ja: "3桁", en: "3 digits" },
            tcyDigits4: { ja: "4桁", en: "4 digits" },
            tcyGlyphByLength: { ja: "字数に合わせた等幅字形", en: "Fixed-width forms by length" },
            colorCheck: { ja: "処理のあとに文字を色分けする", en: "Color characters after processing" },
            areaAutoSize: { ja: "エリア内文字を自動サイズ調整", en: "Auto-size area type" },
            areaHeightToSize: { ja: "高さを文字サイズの整数倍に", en: "Height to a multiple of the type size" },
            tcyExcludeAdjacent: { ja: "桁区切り・小数点の隣は除く", en: "Skip numbers next to separators" },
            tcyCapitals: { ja: "大文字の英字（2〜3字）を縦中横", en: "TCY 2–3 capitals" },
            tcyExclamation: { ja: "「!!」「!?」を縦中横", en: "TCY \"!!\" and \"!?\"" },
            halfwidthAlnum: { ja: "全角英数字を半角に", en: "Full-width alnum to half-width" },
            periodToNakaguro: { ja: "小数点「.」→「・」", en: "Decimal \".\" → \"・\"" },
            symbolsFullwidth: { ja: "記号を全角に", en: "Symbols to full-width" },
            quotesToChonchon: { ja: "引用符を〝〟に", en: "Quotes to 〝〟" },
            rangeToWave: { ja: "範囲の「-」「~」→「〜」", en: "Range \"-\" \"~\" → \"〜\"" },
            exclamFullwidth: { ja: "「!」「?」を全角に", en: "\"!\" \"?\" to full-width" },
            singleLetterFullwidth: { ja: "1文字の英字を全角に", en: "Single letters to full-width" },
            proportionalFullwidthLatin: { ja: "全角の欧文にプロポーショナルメトリクス", en: "Proportional metrics for full-width Latin" },
            singleDigitFullwidth: { ja: "1桁の数字を全角に", en: "Single digits to full-width" },
            japaneseEra: { ja: "西暦を和暦に", en: "Western years to Japanese era" },
            kanjiZeroRei: { ja: "0を「零」に", en: "0 as 零" },
            kanjiIchiSen: { ja: "千を「一千」に", en: "千 as 一千" },
            unitsKumimoji: { ja: "単位を組文字に（cm → ㌢）", en: "Units to squares (cm → ㌢)" },
            kanji: { ja: "漢数字に変換", en: "Convert to kanji numerals" }
        },
        glyphShort: {
            half: { ja: "半角", en: "half" },
            third: { ja: "三分", en: "third" },
            quarter: { ja: "四分", en: "quarter" }
        },
        dropdown: {
            kinsoku: [
                { ja: "なし", en: "None" },
                { ja: "強い禁則", en: "Hard" },
                { ja: "弱い禁則", en: "Soft" },
                { ja: "弱い禁則v2", en: "Soft v2" }
            ],
            comma: [
                { ja: "「、」にする", en: "Use 、" },
                { ja: "取る", en: "Remove" },
                { ja: "そのまま", en: "Keep" }
            ],
            phoneHyphen: [
                { ja: "「‐」（ハイフン）", en: "‐ (hyphen)" },
                { ja: "「ー」（音引き）", en: "ー (long vowel mark)" },
                { ja: "そのまま", en: "Keep" }
            ],
            mojikumi: [
                { ja: "なし", en: "None" },
                { ja: "行末約物全角/半角", en: "Line-end punctuation full/half" },
                { ja: "約物半角", en: "Half-width punctuation" },
                { ja: "行末約物半角", en: "Line-end punctuation half" },
                { ja: "行末約物全角", en: "Line-end punctuation full" },
                { ja: "約物全角", en: "Full-width punctuation" },
                { ja: "ツメ組み", en: "Tight" },
                { ja: "ベタ組み", en: "Solid" }
            ]
        },
        radio: {
            modeSimple: { ja: "簡易", en: "Simple" },
            modeDetail: { ja: "詳細", en: "Detailed" },
            kanjiDigit: { ja: "二〇二六", en: "二〇二六" },
            kanjiPositional: { ja: "二千二十六", en: "二千二十六" },
            kanjiLargeUnits: { ja: "二万五〇〇〇", en: "二万五〇〇〇" }
        },
        fieldLabel: {
            digits: { ja: "桁数", en: "Digits" },
            rangeTo: { ja: "〜", en: "to" },
            charsOrMore: { ja: "文字以上", en: "or more characters" },
            charsOrLess: { ja: "文字以下", en: "or fewer characters" },
            kinsoku: { ja: "禁則", en: "Kinsoku" },
            comma: { ja: "桁区切り", en: "Separators" },
            colorBase: { ja: "ベース", en: "Base" },
            colorFullWidth: { ja: "全角", en: "Full-width" },
            colorHalfWidth: { ja: "半角", en: "Half-width" },
            colorGlyphWidth: { ja: "等幅字形", en: "Fixed-width" },
            phoneHyphen: { ja: "番号のハイフン", en: "Number hyphens" },
            mojikumi: { ja: "文字組み", en: "Mojikumi" },
            leading: { ja: "行送り", en: "Leading" }
        },
        preview: {
            tcyMark: { ja: "［縦中横］", en: " [TCY]" },
            noRotateMark: { ja: "［回転なし］", en: " [not rotated]" },
            tcyGlyphMark: { ja: "［縦中横・{glyph}］", en: " [TCY, {glyph}]" },
            none: { ja: "（変換する文字はありません）", en: "(Nothing to convert)" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        preset: {
            label: { ja: "プリセット", en: "Preset" },
            placeholder: { ja: "---", en: "---" },
            name: { ja: "プリセット名", en: "Preset name" },
            saveTitle: { ja: "プリセットを保存", en: "Save Preset" },
            builtin: {
                standard: { ja: "標準", en: "Standard" },
                arabicTcy2: { ja: "算用数字・2桁だけ縦中横", en: "Arabic, TCY for 2 digits only" },
                arabicMagazine: { ja: "算用数字・2〜3桁縦中横（雑誌・広告）", en: "Arabic, TCY for 2–3 digits (magazines, ads)" },
                kanjiLargeUnits: { ja: "漢数字・万億兆（新聞・一般書）", en: "Kanji, 万・億・兆 only (newspapers, books)" },
                kanjiPositional: { ja: "漢数字・十百千（文芸）", en: "Kanji with 十・百・千 (literature)" },
                kanjiDigitTouten: { ja: "漢数字・1字ずつ＋読点", en: "Kanji digit by digit with 、" },
                rotateOnly: { ja: "欧文回転だけ", en: "Rotation only" }
            }
        },
        confirm: {
            presetOverwrite: { ja: "「{name}」を上書きしますか？", en: "Overwrite \"{name}\"?" },
            presetDelete: { ja: "「{name}」を削除しますか？", en: "Delete \"{name}\"?" }
        },
        progress: {
            title: { ja: "縦組みの整形", en: "Vertical Text Formatter" },
            rotate: { ja: "欧文回転を適用中...", en: "Rotating Roman characters..." },
            edit: { ja: "文字を変換中... {current}/{total}", en: "Converting text... {current}/{total}" },
            autoSize: { ja: "自動サイズ調整を適用中...", en: "Applying auto size..." },
            tcy: { ja: "縦中横を適用中... {current}/{total}", en: "Applying Tate-Chu-Yoko... {current}/{total}" }
        },
        alert: {
            noSelection: { ja: "テキストを選択してください。", en: "Select text first." },
            presetSaveFailed: { ja: "プリセットを保存できませんでした。", en: "Could not save the preset." },
            presetNameTaken: { ja: "組み込みのプリセットと同じ名前は使えません。", en: "That name is used by a built-in preset." },
            actionFailed: { ja: "縦組み中の欧文回転・縦中横を適用できませんでした。", en: "Could not apply the vertical text settings." }
        },
        /* ステップボタン用（stepUp〜）と、項目の説明 / for the stepper (stepUp…) and option descriptions */
        tooltip: {
            leading: {
                ja: "行送りを、文字サイズに対する自動行送りの％で指定します（50〜500%）。行送りは常に「自動」になります",
                en: "Leading as an auto-leading percentage of the type size (50–500%)"
            },
            kinsoku: { ja: "段落の禁則処理を選びます", en: "Kinsoku set for the paragraphs" },
            mojikumi: {
                ja: "段落の文字組みアキ量設定を選びます（行揃えは常に行頭行末揃え（最終行上））",
                en: "Mojikumi set for the paragraphs (justification is always justified, last line top)"
            },
            modeSimple: {
                ja: "よく使う項目だけを表示して処理します。表示していない項目はオフにし、禁則・文字組みは変えません",
                en: "Shows and applies only the common options; hidden options are off and kinsoku / mojikumi stay unchanged"
            },
            modeDetail: { ja: "すべての項目を表示します", en: "Shows all options" },
            preset: { ja: "組み込みのプリセットか、保存した設定を読み込みます。", en: "Loads a built-in preset or saved settings." },
            presetSave: { ja: "今の設定に名前を付けて保存します。", en: "Saves the current settings under a name." },
            presetDelete: { ja: "選んでいる保存したプリセットを削除します。", en: "Deletes the selected saved preset." },
            presetName: {
                ja: "保存する設定の名前です。同じ名前は上書きします。",
                en: "Name the settings are saved under. The same name is overwritten."
            },
            shortWordFullwidth: {
                ja: "指定した文字数以下の英単語（英字だけの並び）を全角にして正立させます（OK → ＯＫ）。文字数は［長い欧文は除く］の1つ下に連動します",
                en: "Makes words of at most this many letters full-width so they stand upright (OK → ＯＫ); linked to one below \"Skip long Roman text\""
            },
            rotateSkipLongLatin: {
                ja: "指定した文字数以上の欧文（スペースなどでつながった単語のまとまり）は回転せず、横倒しのままにします。［短い英単語を全角に］の文字数はこの1つ下に連動します",
                en: "Leaves Roman text of at least this many characters (words joined by spaces etc.) sideways"
            },
            tcyNumbers: { ja: "下で選んだ桁数の半角数字を縦中横にします", en: "Tate-Chu-Yoko for half-width numbers of the lengths checked below" },
            kanji: {
                ja: "桁数が範囲内の数字を漢数字にします。縦中横の桁数と重なるときは漢数字を優先します",
                en: "Converts numbers within the digit range to kanji numerals; this wins over Tate-Chu-Yoko"
            },
            tcyGlyphByLength: {
                ja: "縦中横の文字を、2字は等幅半角・3字は等幅三分・4字は等幅四分の字形にします",
                en: "Uses half-width forms for 2 characters, third-width for 3 and quarter-width for 4"
            },
            tcyExcludeAdjacent: {
                ja: "「3,000」の「000」や「12.5」の「12」を縦中横にしません",
                en: "Leaves \"000\" in \"3,000\" and \"12\" in \"12.5\" alone"
            },
            tcyCapitals: {
                ja: "「PDF」「OK」など、すべて大文字の2〜3字の英字を縦中横にします",
                en: "Tate-Chu-Yoko for all-capital words of 2–3 letters such as PDF"
            },
            tcyExclamation: {
                ja: "感嘆符・疑問符の2字の組を縦中横にします（全角は半角にそろえます）",
                en: "Tate-Chu-Yoko for pairs of ! and ? (full-width marks become half-width)"
            },
            halfwidthAlnum: {
                ja: "「２０」「ＰＤＦ」なども縦中横・欧文回転の対象にします",
                en: "Makes full-width digits and letters targets for Tate-Chu-Yoko and rotation"
            },
            areaHeightToSize: {
                ja: "エリア内文字の枠の高さ（縦組みでは1行の長さ）を、文字サイズのいちばん近い整数倍にそろえます（内側の余白は除いて測ります）。簡易モードでも常にかけます",
                en: "Sets the area-type frame height (the line length in vertical text) to the nearest whole multiple of the type size (inset excluded); always applied, in Simple mode too"
            },
            areaAutoSize: {
                ja: "処理のあと、エリア内文字に［エリア内文字オプション］の［自動サイズ調整］を適用し、枠を文字に合わせます（連結したテキストは対象外）",
                en: "After processing, turns on Auto Size in Area Type Options so the frame fits the text (threaded text is skipped)"
            },
            colorCheck: {
                ja: "処理のあと、全角の英数字と「！？」を赤、半角を青、等幅字形の文字を緑、ほかをベースの色で塗ります（仕上がりの確認用。元の色には戻りません）",
                en: "After processing, colors full-width alphanumerics and ！？ red, half-width blue, fixed-width glyphs green and the rest in the base color (for checking; the original colors are not restored)"
            },
            colorBase: { ja: "全角・半角・等幅字形のどれにも当たらない文字の色。クリックでカラーピッカーを開きます", en: "Color for all other characters. Click to open the color picker" },
            colorFullWidth: { ja: "全角の英数字と「！？」の色。クリックでカラーピッカーを開きます", en: "Color for full-width alphanumerics and ！？. Click to open the color picker" },
            colorHalfWidth: { ja: "半角の英数字と「!?」の色。クリックでカラーピッカーを開きます", en: "Color for half-width alphanumerics and !?. Click to open the color picker" },
            colorGlyphWidth: { ja: "等幅全角・半角・三分・四分字形の文字の色（全角・半角より優先）。クリックでカラーピッカーを開きます", en: "Color for fixed-width glyphs (takes priority). Click to open the color picker" },
            comma: {
                ja: "桁区切りの「,」「，」の扱いです（「、」にする：3,000 → 3、000／取る：3,000 → 3000）",
                en: "How to handle thousands separators (Use 、: 3,000 → 3、000 / Remove: 3,000 → 3000)"
            },
            phoneHyphen: {
                ja: "電話番号・郵便番号（数字-数字-数字、3桁-4桁）のハイフンの扱いです。番号は範囲の「〜」・縦中横・漢数字の対象にしません",
                en: "Hyphens in phone and postal numbers; these numbers are never ranges, Tate-Chu-Yoko or kanji"
            },
            periodToNakaguro: { ja: "小数点の「.」「．」を「・」にします（12.5 → 12・5）", en: "Replaces decimal points \".\" \"．\" with \"・\"" },
            symbolsFullwidth: {
                ja: "半角の ( ) [ ] { } : ; / % & # @ * + = < > を全角にして正立させます",
                en: "Makes half-width ( ) [ ] { } : ; / % & # @ * + = < > full-width so they stand upright"
            },
            quotesToChonchon: { ja: "“ ” と \" \" をチョンチョン（〝 〟）にします", en: "Replaces “ ” and \" \" with 〝 〟" },
            rangeToWave: {
                ja: "数字の間の「-」「~」「－」「～」を「〜」にします（10-20 → 10〜20）",
                en: "Replaces - ~ between digits with 〜 (10-20 → 10〜20)"
            },
            exclamFullwidth: { ja: "縦中横にしない「!」「?」を全角にして正立させます", en: "Makes ! and ? that are not in Tate-Chu-Yoko full-width" },
            proportionalFullwidthLatin: {
                ja: "全角の英字（Ａ〜Ｚ・ａ〜ｚ）に OpenType 機能の［プロポーショナルメトリクス］を適用します（文字の書き換えのあと）",
                en: "Applies the OpenType Proportional Metrics feature to full-width letters (after the rewriting)"
            },
            singleLetterFullwidth: {
                ja: "1字だけの英字（Aランクの A など）を全角にして正立させます",
                en: "Makes a lone letter (the A in Aランク) full-width"
            },
            singleDigitFullwidth: { ja: "1桁の数字を全角にして正立させます", en: "Makes single digits full-width so they stand upright" },
            japaneseEra: {
                ja: "「2026年」を「令和8年」にします。月日があれば改元の前後を月日で判定します",
                en: "2026年 → 令和8年; a following month and day decide the era at a change"
            },
            kanjiZeroRei: { ja: "0を「〇」ではなく「零」にします（二〇二六 → 二零二六）", en: "Writes 0 as 零 instead of 〇" },
            kanjiIchiSen: { ja: "十・百・千を入れるとき、千を「一千」と書きます（1500 → 一千五百）", en: "Writes 1000 as 一千 when using 十・百・千" },
            unitsKumimoji: {
                ja: "数字の直後の mm・cm・km・m・mg・g・kg・t・ml・L を ㍉・㌢・㌖・㍍・㍉㌘・㌘・㌕・㌧・㍉㍑・㍑ にします。字形の無いフォントもあります",
                en: "Turns mm, cm, km, m, mg, g, kg, t, ml, L after digits into katakana squares; some fonts lack these glyphs"
            },
            kanjiDigit: { ja: "1字ずつ漢数字にします（2026 → 二〇二六）", en: "Digit by digit (2026 → 二〇二六)" },
            kanjiPositional: { ja: "十・百・千を入れます（2000 → 二千）。直後が「年」の西暦は1字ずつ", en: "With 十・百・千 (2000 → 二千); years followed by 年 stay digit by digit" },
            kanjiLargeUnits: {
                ja: "万・億・兆だけを入れ、間は1字ずつ書きます（25000 → 二万五〇〇〇）。直後が「年」の西暦は1字ずつ",
                en: "Only 万・億・兆, digit by digit in between (25000 → 二万五〇〇〇); years followed by 年 stay digit by digit"
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
    // 一時アクション定義 / Action definition
    // =========================================
    var ACTION_SET_NAME = "VerticalTextFormatter";       /* アクションセット名 / action set name */
    var ROTATE_ACTION_NAME = "rotateRoman";              /* 欧文回転のアクション名 / rotate action name */
    var TCY_ACTION_NAME = "tateChuYoko";                 /* 縦中横のアクション名 / Tate-Chu-Yoko action name */
    var ROTATE_OFF_ACTION_NAME = "rotateRomanOff";       /* 欧文回転を外すアクション名 / action name that turns rotation off */
    var AREA_AUTO_SIZE_ACTION_NAME = "areaAutoSize";     /* エリア内文字の自動サイズ調整のアクション名 / area-type auto-size action name */

    /**
     * 文字パネルの真偽値の項目を1つだけ設定するアクション1件ぶんの行を返す
     * （文字パネルのイベントは、書いたパラメーターだけを適用する）
     * @param {number} actionIndex - アクションの番号（1始まり）
     * @param {string} actionName - アクション名
     * @param {number} paramKey - パラメーターのキー
     * @param {number} paramValue - 値（1 でオン、0 でオフ）
     * @returns {string[]} アクション定義の行
     */
    function buildCharacterToggleActionLines(actionIndex, actionName, paramKey, paramValue) {
        return ["/action-" + actionIndex + " {"].concat(
            buildActionNameLines("\t", actionName),
            [
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (adobe_SLOCharacterPalette)"
            ],
            buildActionNameLines("\t\t", "文字", "localizedName"),
            [
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key " + paramKey,
                "\t\t\t/showInPalette 4294967295",
                "\t\t\t/type (boolean)",
                "\t\t\t/value " + paramValue,
                "\t\t}",
                "\t}",
                "}"
            ]
        );
    }

    /**
     * ［エリア内文字オプション］の［自動サイズ調整］をオンにするアクション1件ぶんの行を返す（AutoFitTextFrame と同じ定義）
     * @param {number} actionIndex - アクションの番号（1始まり）
     * @param {string} actionName - アクション名
     * @returns {string[]} アクション定義の行
     */
    function buildAreaAutoSizeActionLines(actionIndex, actionName) {
        return ["/action-" + actionIndex + " {"].concat(
            buildActionNameLines("\t", actionName),
            [
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (adobe_SLOAreaTextDialog)"
            ],
            buildActionNameLines("\t\t", "エリア内文字オプション", "localizedName"),
            [
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key 1952539754",
                "\t\t\t/showInPalette 4294967295",
                "\t\t\t/type (integer)",
                "\t\t\t/value 1",
                "\t\t}",
                "\t}",
                "}"
            ]
        );
    }

    /**
     * 欧文回転（'rriv' オン・オフ）・縦中横（'tyen'）・エリア内文字の自動サイズ調整の4件を持つアクション定義を返す
     * @returns {string} アクション定義のテキスト
     */
    function buildVerticalTextActions() {
        return ["/version 3"].concat(
            buildActionNameLines("", ACTION_SET_NAME),
            ["/isOpen 1", "/actionCount 4"],
            buildCharacterToggleActionLines(1, ROTATE_ACTION_NAME, 1920100726, 1),
            buildCharacterToggleActionLines(2, TCY_ACTION_NAME, 1954112878, 1),
            buildCharacterToggleActionLines(3, ROTATE_OFF_ACTION_NAME, 1920100726, 0),
            buildAreaAutoSizeActionLines(4, AREA_AUTO_SIZE_ACTION_NAME),
            [""]
        ).join("\n");
    }

    // =========================================
    // 変換の計画 / Conversion plan
    // =========================================
    var KANJI_DIGITS = ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
    var KANJI_SMALL_UNITS = ["", "十", "百", "千"];
    var KANJI_LARGE_UNITS = ["", "万", "億", "兆", "京"];

    /**
     * 半角数字か
     * @param {string} currentChar - 1文字
     * @returns {boolean} 0〜9 なら true
     */
    function isDigitChar(currentChar) {
        return currentChar >= "0" && currentChar <= "9";
    }

    /**
     * 半角英字か
     * @param {string} currentChar - 1文字
     * @returns {boolean} A〜Z・a〜z なら true
     */
    function isLatinLetter(currentChar) {
        return (currentChar >= "A" && currentChar <= "Z") || (currentChar >= "a" && currentChar <= "z");
    }

    /**
     * 感嘆符・疑問符か（半角・全角）
     * @param {string} currentChar - 1文字
     * @returns {boolean} ! ? ！ ？ なら true
     */
    function isExclamOrQuestion(currentChar) {
        return currentChar === "!" || currentChar === "?" || currentChar === "！" || currentChar === "？";
    }

    /* ［単位を組文字に］の対応（数字の直後の英字だけ） / Units turned into katakana squares (only right after digits) */
    var UNIT_KUMIMOJI = {
        mm: "㍉", cm: "㌢", km: "㌖", m: "㍍",
        mg: "㍉㌘", g: "㌘", kg: "㌕", t: "㌧",
        ml: "㍉㍑", mL: "㍉㍑", l: "㍑", L: "㍑"
    };

    /* 欧文のまとまり：単語（英字で始まるか、数字のあとに英字が続くもの。. ' - を含む。全角の英数字も可）が、スペースまたは「&」（全角も可）でつながったもの。
       スペースは半角・全角のほか、ノーブレークスペースや細いスペースなども含める。
       ExtendScript の正規表現は、入れ子の量指定と選択を含むパターンを誤って評価する（Design Week の「Week」が外れる）ので、1文字ずつ走査する
       A Roman run: words (starting with a letter, or digits followed by a letter; may contain . ' -; full-width alphanumerics allowed) joined by spaces or & (full-width too).
       ExtendScript's regex engine mis-evaluates nested quantifiers with alternation (dropping "Week" from Design Week), so scan by hand */
    var LATIN_SPACES = "               　";
    var LATIN_AMPERSANDS = "&＆";

    /**
     * 半角・全角の英字か（欧文のまとまりの判定用。1語ずつの変換は半角の isLatinLetter だけを見る）
     * @param {string} currentChar - 1文字
     * @returns {boolean} A〜Z・a〜z・Ａ〜Ｚ・ａ〜ｚ なら true
     */
    function isAnyLatinLetter(currentChar) {
        return isLatinLetter(currentChar) || (currentChar >= "Ａ" && currentChar <= "Ｚ") || (currentChar >= "ａ" && currentChar <= "ｚ");
    }

    /**
     * 単語を構成する文字か（英数字と . ' -）
     * @param {string} currentChar - 1文字
     * @returns {boolean} 単語の文字なら true
     */
    function isLatinWordPart(currentChar) {
        return currentChar !== "" && (isAnyLatinLetter(currentChar) || isAnyDigitChar(currentChar) || ".'-".indexOf(currentChar) >= 0);
    }

    /**
     * 欧文のまとまりのつなぎに使うスペースか
     * @param {string} currentChar - 1文字
     * @returns {boolean} スペースなら true
     */
    function isLatinSpace(currentChar) {
        return currentChar !== "" && LATIN_SPACES.indexOf(currentChar) >= 0;
    }

    /**
     * その位置から始まる1語を読む。単語の途中・数字＋単位（5mm など。［単位を組文字に］の対象）は語として扱わない
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 調べる位置
     * @returns {{end: number, isSentenceEnd: boolean, isCapsTcy: boolean}|null} 終わり（含まない）、文末のピリオドで終わるか、
     *     縦中横にする大文字（半角2〜3字）か。語でなければ null
     */
    function readLatinWord(text, startIndex) {
        var previousChar = text.charAt(startIndex - 1);
        if (startIndex > 0 && (isAnyLatinLetter(previousChar) || isAnyDigitChar(previousChar))) return null;
        var index = startIndex;
        while (index < text.length && isAnyDigitChar(text.charAt(index))) index++;
        var letterStart = index;
        if (!isAnyLatinLetter(text.charAt(letterStart))) return null;
        var wordEnd = letterStart + 1;
        while (wordEnd < text.length && isLatinWordPart(text.charAt(wordEnd))) wordEnd++;
        var wordText = text.substring(startIndex, wordEnd);
        /* 数字のあとが単位だけなら、数字と単位として扱う / Digits followed only by a unit stay a number and a unit */
        if (letterStart > startIndex && UNIT_KUMIMOJI.hasOwnProperty(text.substring(letterStart, wordEnd))) return null;
        /* 文末：4字以上の語がピリオドで終わる（Mr. や U.S. などの略語は除く） / Sentence end: a word of 4+ chars ending in a period (not Mr. or U.S.) */
        var coreText = wordText.replace(/\.+$/, "");
        var isSentenceEnd = coreText !== wordText && coreText.length >= 4 && coreText.indexOf(".") < 0;
        var isCapsTcy = /^[A-Z]{2,3}$/.test(wordText);
        return { end: wordEnd, isSentenceEnd: isSentenceEnd, isCapsTcy: isCapsTcy };
    }

    /**
     * その位置から始まる欧文のまとまりの終わりと語数を返す
     * ［大文字を縦中横に］がオンで、全部の語が縦中横にする大文字（NHK BS など）なら、最初の1語だけを返す
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 調べる位置（どこでもよい）
     * @param {Object} formatOptions - 変換の設定（tcyCapitals を見る）
     * @returns {{end: number, wordCount: number}|null} 終わり（含まない）と語数。語で始まらなければ null
     */
    function scanLatinRun(text, startIndex, formatOptions) {
        var word = readLatinWord(text, startIndex);
        if (!word) return null;
        var firstWordEnd = word.end;
        var wordEnd = word.end;
        var wordCount = 1;
        var isAllCapsTcy = word.isCapsTcy;
        while (true) {
            /* つなぎ：スペースの並び、または前後にスペースを挟んでよい「&」 / joiner: spaces, or & with optional spaces around it */
            var nextIndex = wordEnd;
            while (nextIndex < text.length && isLatinSpace(text.charAt(nextIndex))) nextIndex++;
            var joinerChar = text.charAt(nextIndex);
            if (joinerChar !== "" && LATIN_AMPERSANDS.indexOf(joinerChar) >= 0) {
                nextIndex++;
                while (nextIndex < text.length && isLatinSpace(text.charAt(nextIndex))) nextIndex++;
            }
            if (nextIndex === wordEnd) break;
            var nextWord = readLatinWord(text, nextIndex);
            if (!nextWord) break;
            /* 文末のピリオドのあと、大文字で始まる語は別の文 / After a sentence-ending period, a capitalized word starts a new sentence */
            var nextChar = text.charAt(nextIndex);
            if (word.isSentenceEnd && ((nextChar >= "A" && nextChar <= "Z") || (nextChar >= "Ａ" && nextChar <= "Ｚ"))) break;
            word = nextWord;
            wordEnd = word.end;
            wordCount++;
            if (!word.isCapsTcy) isAllCapsTcy = false;
        }
        if (wordCount >= 2 && isAllCapsTcy && formatOptions && formatOptions.tcyCapitals) return { end: firstWordEnd, wordCount: 1 };
        return { end: wordEnd, wordCount: wordCount };
    }

    var FULLWIDTH_SYMBOLS = "()[]{}:;/%&#@*+=<>";   /* ［記号を全角に］の対象 / symbols made full-width */
    var RANGE_MARKS = "-~－～";                      /* ［範囲の「-」「~」→「〜」］の対象 / range marks */

    /**
     * 半角の英数字・記号（! 〜 ~）を全角にする
     * @param {string} asciiText - 半角の文字列
     * @returns {string} 全角の文字列
     */
    function toFullwidthAscii(asciiText) {
        var resultText = "";
        for (var i = 0; i < asciiText.length; i++) {
            var charCode = asciiText.charCodeAt(i);
            resultText += (charCode >= 0x21 && charCode <= 0x7E) ? String.fromCharCode(charCode + 0xFEE0) : asciiText.charAt(i);
        }
        return resultText;
    }

    /**
     * 半角・全角の数字か
     * @param {string} currentChar - 1文字
     * @returns {boolean} 0〜9・０〜９ なら true
     */
    function isAnyDigitChar(currentChar) {
        return isDigitChar(currentChar) || (currentChar >= "０" && currentChar <= "９");
    }

    /**
     * 数字・英字・感嘆符以外の1文字を、記号・引用符・範囲の設定に合わせて置き換える
     * @param {string} text - 対象の文字列
     * @param {number} index - 文字の位置
     * @param {Object} formatOptions - 変換の設定
     * @param {Object} quoteState - 半角の「"」の開き・閉じを数える状態（{ isOpen }）
     * @returns {string} 置き換えた文字
     */
    function convertOtherChar(text, index, formatOptions, quoteState) {
        var currentChar = text.charAt(index);
        if (formatOptions.rangeToWave && RANGE_MARKS.indexOf(currentChar) >= 0 && isAnyDigitChar(text.charAt(index - 1)) && isAnyDigitChar(text.charAt(index + 1))) {
            return "〜";
        }
        if (formatOptions.quotesToChonchon) {
            if (currentChar === "“") return "〝";
            if (currentChar === "”") return "〟";
            if (currentChar === "\"") {
                /* 半角の「"」は開きと閉じを交互に / Straight quotes alternate between opening and closing */
                quoteState.isOpen = !quoteState.isOpen;
                return quoteState.isOpen ? "〝" : "〟";
            }
        }
        if (formatOptions.symbolsFullwidth && FULLWIDTH_SYMBOLS.indexOf(currentChar) >= 0) return toFullwidthAscii(currentChar);
        return currentChar;
    }

    /**
     * その位置から始まる英単語の連続（スペースや「&」でつながった2語以上）を返す
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 調べる位置（どこでもよい。語で始まらなければ null）
     * @param {Object} formatOptions - 変換の設定
     * @returns {string|null} まとまりの文字列（語で始まらない・1語だけなら null）
     */
    function matchLatinPhrase(text, startIndex, formatOptions) {
        var latinRun = scanLatinRun(text, startIndex, formatOptions);
        return (latinRun && latinRun.wordCount >= 2) ? text.substring(startIndex, latinRun.end) : null;
    }

    /**
     * 全角にする短い英単語か（［1文字の英字を全角に］［短い英単語を全角に］）
     * @param {string} letterText - 英字だけの並び
     * @param {Object} formatOptions - 変換の設定
     * @returns {boolean} 全角にするなら true
     */
    function isShortWordForFullwidth(letterText, formatOptions) {
        if (formatOptions.singleLetterFullwidth && letterText.length === 1) return true;
        return formatOptions.shortWordFullwidth && letterText.length <= formatOptions.shortWordMaxChars;
    }

    /**
     * 全角の英数字を半角にする（1文字ずつ置き換えるので文字数は変わらない）
     * @param {string} text - 対象の文字列
     * @returns {string} 置き換えた文字列
     */
    function toHalfwidthAlnum(text) {
        var resultText = "";
        for (var i = 0; i < text.length; i++) {
            var charCode = text.charCodeAt(i);
            var isFullAlnum = (charCode >= 0xFF10 && charCode <= 0xFF19) || (charCode >= 0xFF21 && charCode <= 0xFF3A) || (charCode >= 0xFF41 && charCode <= 0xFF5A);
            resultText += isFullAlnum ? String.fromCharCode(charCode - 0xFEE0) : text.charAt(i);
        }
        return resultText;
    }

    /**
     * 数字1字の漢数字を返す（0 は設定で「〇」か「零」）
     * @param {number} digitValue - 0〜9
     * @param {Object} formatOptions - 変換の設定
     * @returns {string} 漢数字1字
     */
    function getKanjiDigit(digitValue, formatOptions) {
        return (digitValue === 0 && formatOptions.kanjiZeroRei) ? "零" : KANJI_DIGITS[digitValue];
    }

    /**
     * 数字を1字ずつ漢数字にする（2026 → 二〇二六）
     * @param {string} digitText - 半角数字の文字列
     * @param {Object} formatOptions - 変換の設定
     * @returns {string} 漢数字の文字列
     */
    function toKanjiDigitByDigit(digitText, formatOptions) {
        var resultText = "";
        for (var i = 0; i < digitText.length; i++) {
            resultText += getKanjiDigit(digitText.charCodeAt(i) - 48, formatOptions);
        }
        return resultText;
    }

    /**
     * 整数を十・百・千と万・億…を入れた漢数字にする（2026 → 二千二十六、10000 → 一万）
     * 十・百の前の「一」は省く。千の前は設定で「一千」にする。京の位を超える桁は1字ずつにする
     * @param {string} digitText - 半角数字の文字列
     * @param {Object} formatOptions - 変換の設定
     * @returns {string} 漢数字の文字列
     */
    function toKanjiPositional(digitText, formatOptions) {
        var trimmedText = digitText.replace(/^0+/, "");
        if (trimmedText === "") return getKanjiDigit(0, formatOptions);
        if (trimmedText.length > KANJI_LARGE_UNITS.length * 4) return toKanjiDigitByDigit(digitText, formatOptions);
        var resultText = "";
        var groupCount = Math.ceil(trimmedText.length / 4);
        for (var g = groupCount - 1; g >= 0; g--) {
            var groupEnd = trimmedText.length - g * 4;
            var groupText = trimmedText.substring(Math.max(0, groupEnd - 4), groupEnd);
            var groupKanji = "";
            for (var i = 0; i < groupText.length; i++) {
                var digitValue = groupText.charCodeAt(i) - 48;
                if (digitValue === 0) continue;
                var place = groupText.length - 1 - i;
                /* 十・百の前の「一」は書かない。千は設定次第（一千） / Omit "one" before ten and hundred; before thousand it depends on the setting */
                var omitsOne = (digitValue === 1 && place > 0 && !(place === 3 && formatOptions.kanjiIchiSen));
                if (!omitsOne) groupKanji += KANJI_DIGITS[digitValue];
                groupKanji += KANJI_SMALL_UNITS[place];
            }
            if (groupKanji !== "") resultText += groupKanji + KANJI_LARGE_UNITS[g];
        }
        return resultText;
    }

    /**
     * 整数を万・億・兆…の単位語だけ入れた漢数字にする（25000 → 二万五〇〇〇、123456789 → 一億二三四五万六七八九）
     * 単位語の間は1字ずつ書き、すべて0の4桁は単位語ごと省く
     * @param {string} digitText - 半角数字の文字列
     * @param {Object} formatOptions - 変換の設定
     * @returns {string} 漢数字の文字列
     */
    function toKanjiLargeUnits(digitText, formatOptions) {
        var trimmedText = digitText.replace(/^0+/, "");
        if (trimmedText === "") return getKanjiDigit(0, formatOptions);
        if (trimmedText.length > KANJI_LARGE_UNITS.length * 4) return toKanjiDigitByDigit(digitText, formatOptions);
        var resultText = "";
        var groupCount = Math.ceil(trimmedText.length / 4);
        for (var g = groupCount - 1; g >= 0; g--) {
            var groupEnd = trimmedText.length - g * 4;
            var groupText = trimmedText.substring(Math.max(0, groupEnd - 4), groupEnd);
            if (/^0+$/.test(groupText)) continue;
            resultText += toKanjiDigitByDigit(groupText, formatOptions) + KANJI_LARGE_UNITS[g];
        }
        return resultText;
    }

    /**
     * 文字列のその位置から数字のまとまり（桁区切り・小数を含む）を読む
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 数字で始まる位置
     * @returns {{length: number, parts: Array<{type: string, text: string}>, integerDigits: number, hasSeparator: boolean, hasComma: boolean, hasDecimal: boolean}}
     *   parts の type は "digits" / "separator"（, ，）/ "decimal"（. ．）
     */
    function readNumberToken(text, startIndex) {
        var parts = [];
        var i = startIndex;
        var integerDigits = 0;
        var hasSeparator = false;
        var hasComma = false;
        var hasDecimal = false;

        /**
         * その位置から続く数字を読み、parts に足す
         * @returns {number} 読んだ桁数
         */
        function readDigits() {
            var digitStart = i;
            while (i < text.length && isDigitChar(text.charAt(i))) i++;
            parts.push({ type: "digits", text: text.substring(digitStart, i) });
            return i - digitStart;
        }

        integerDigits += readDigits();
        /* 桁区切り：「,」「，」のあとにちょうど3桁 / Group separator: "," or "，" followed by exactly three digits */
        while ((text.charAt(i) === "," || text.charAt(i) === "，") && isDigitChar(text.charAt(i + 1)) && isDigitChar(text.charAt(i + 2)) &&
                isDigitChar(text.charAt(i + 3)) && !isDigitChar(text.charAt(i + 4))) {
            parts.push({ type: "separator", text: text.charAt(i) });
            i++;
            integerDigits += readDigits();
            hasSeparator = true;
            hasComma = true;
        }
        /* 小数点：「.」「．」のあとに数字 / Decimal point: "." or "．" followed by a digit */
        if ((text.charAt(i) === "." || text.charAt(i) === "．") && isDigitChar(text.charAt(i + 1))) {
            parts.push({ type: "decimal", text: text.charAt(i) });
            i++;
            readDigits();
            hasSeparator = true;
            hasDecimal = true;
        }
        return { length: i - startIndex, parts: parts, integerDigits: integerDigits, hasSeparator: hasSeparator, hasComma: hasComma, hasDecimal: hasDecimal };
    }

    /**
     * 桁区切り・小数点を設定に合わせて置き換える（置き換えないときは元の文字のまま）
     * @param {Object} part - readNumberToken の parts の要素（separator / decimal）
     * @param {Object} formatOptions - 変換の設定
     * @returns {string} 置き換えた文字
     */
    function convertNumberMark(part, formatOptions) {
        if (part.type === "separator") {
            if (formatOptions.commaStyle === "touten") return "、";
            if (formatOptions.commaStyle === "remove") return "";
            return part.text;
        }
        return formatOptions.periodToNakaguro ? "・" : part.text;
    }

    /**
     * 数字のまとまりを漢数字にする（小数部は書き方によらず1字ずつ）
     * @param {Object} numberToken - readNumberToken の戻り値
     * @param {Object} formatOptions - 変換の設定
     * @param {boolean} [forceDigitByDigit] - 書き方によらず1字ずつにする（西暦の年など）
     * @returns {string} 漢数字の文字列
     */
    function convertNumberToKanji(numberToken, formatOptions, forceDigitByDigit) {
        var parts = numberToken.parts;
        var resultText = "";
        if (!forceDigitByDigit && (formatOptions.kanjiStyle === "positional" || formatOptions.kanjiStyle === "largeUnits")) {
            /* 整数部は桁区切りを外して位取りで読む / Read the integer part positionally, without separators */
            var integerText = "";
            var partIndex = 0;
            while (partIndex < parts.length && parts[partIndex].type !== "decimal") {
                if (parts[partIndex].type === "digits") integerText += parts[partIndex].text;
                partIndex++;
            }
            resultText = (formatOptions.kanjiStyle === "positional") ? toKanjiPositional(integerText, formatOptions) : toKanjiLargeUnits(integerText, formatOptions);
            if (partIndex < parts.length) {
                resultText += convertNumberMark(parts[partIndex], formatOptions) + toKanjiDigitByDigit(parts[partIndex + 1].text, formatOptions);
            }
            return resultText;
        }
        for (var i = 0; i < parts.length; i++) {
            resultText += (parts[i].type === "digits") ? toKanjiDigitByDigit(parts[i].text, formatOptions) : convertNumberMark(parts[i], formatOptions);
        }
        return resultText;
    }

    /**
     * 数字だけの並び（1つの digits）を、縦中横・全角・そのままのどれかにして出力に足す
     * @param {string} digitText - 半角数字の文字列
     * @param {Object} formatOptions - 変換の設定
     * @param {function(string, boolean): void} emit - 出力に足す関数（文字列、縦中横にするか）
     * @returns {void}
     */
    function emitDigitRun(digitText, formatOptions, emit) {
        var digitCount = digitText.length;
        if (formatOptions.tcyNumbers && formatOptions["tcyDigits" + digitCount] === true) {
            emit(digitText, true);
        } else if (formatOptions.singleDigitFullwidth && digitCount === 1) {
            emit(toFullwidthAscii(digitText), false);
        } else {
            emit(digitText, false);
        }
    }

    /**
     * 数字だけの並びを、範囲内なら漢数字、そうでなければ縦中横・全角・そのままにして出力に足す
     * @param {string} digitText - 半角数字の文字列
     * @param {Object} formatOptions - 変換の設定
     * @param {function(string, boolean): void} emit - 出力に足す関数
     * @returns {void}
     */
    function emitPlainDigits(digitText, formatOptions, emit) {
        if (formatOptions.kanji && digitText.length >= formatOptions.kanjiMinDigits && digitText.length <= formatOptions.kanjiMaxDigits) {
            emit(convertNumberToKanji({ parts: [{ type: "digits", text: digitText }] }, formatOptions), false);
        } else {
            emitDigitRun(digitText, formatOptions, emit);
        }
    }

    /* 元号と開始日（YYYYMMDD）。新しい順 / Japanese eras and their start dates (YYYYMMDD), newest first */
    var JAPANESE_ERAS = [
        { name: "令和", start: 20190501 },
        { name: "平成", start: 19890108 },
        { name: "昭和", start: 19261225 },
        { name: "大正", start: 19120730 },
        { name: "明治", start: 18681023 }
    ];

    /**
     * その位置の「西暦4桁＋年」を読み、元号と年を返す。月日が続けば改元の境目も月日で判定する
     * （月日が無い改元の年は新しい元号にする：2019年 → 令和元年）
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 数字で始まる位置
     * @returns {{eraName: string, eraYear: number}|null} 西暦の年でなければ null
     */
    function matchWesternYear(text, startIndex) {
        /* 入れ子の量指定の正規表現は ExtendScript で誤評価されるので、年・月・日を順に読む
           Nested quantified groups mis-evaluate in ExtendScript, so read year, month and day in turn */
        var yearMatch = /^(\d{4})年/.exec(text.substr(startIndex, 5));
        if (!yearMatch) return null;
        var year = parseInt(yearMatch[1], 10);
        var month = 12;
        var day = 31;
        var monthMatch = /^(\d{1,2})月/.exec(text.substr(startIndex + 5, 3));
        if (monthMatch) {
            month = parseInt(monthMatch[1], 10);
            var dayMatch = /^(\d{1,2})日/.exec(text.substr(startIndex + 5 + monthMatch[0].length, 3));
            if (dayMatch) day = parseInt(dayMatch[1], 10);
        }
        var dateKey = year * 10000 + month * 100 + day;
        for (var i = 0; i < JAPANESE_ERAS.length; i++) {
            if (dateKey >= JAPANESE_ERAS[i].start) {
                return { eraName: JAPANESE_ERAS[i].name, eraYear: year - Math.floor(JAPANESE_ERAS[i].start / 10000) + 1 };
            }
        }
        return null;
    }

    /* 番号のハイフンとみなす文字 / Characters treated as hyphens in phone and postal numbers */
    var PHONE_HYPHENS = "-－‐−";

    /**
     * その位置の電話番号・郵便番号（数字-数字-数字…、または 3桁-4桁）を読む
     * @param {string} text - 対象の文字列
     * @param {number} startIndex - 数字で始まる位置
     * @returns {string|null} 番号の文字列（番号でなければ null）
     */
    function matchPhoneNumber(text, startIndex) {
        var restText = text.substr(startIndex, 40);
        var phoneMatch = /^\d+(?:[-－‐−]\d+){2,}/.exec(restText) || /^\d{3}[-－‐−]\d{4}(?!\d)/.exec(restText);
        return phoneMatch ? phoneMatch[0] : null;
    }

    /**
     * 番号のハイフンを設定に合わせて置き換える
     * @param {string} phoneText - 番号の文字列
     * @param {Object} formatOptions - 変換の設定（phoneHyphenStyle）
     * @returns {string} 置き換えた文字列
     */
    function convertPhoneHyphens(phoneText, formatOptions) {
        if (formatOptions.phoneHyphenStyle === "keep") return phoneText;
        /* 縦組みで縦になる文字にする。全角の「－」（U+FF0D）は縦組み用の字形が無く横棒のまま
           Use characters that turn vertical; full-width － (U+FF0D) has no vertical form and stays horizontal */
        var hyphenText = (formatOptions.phoneHyphenStyle === "choon") ? "ー" : "\u2010";
        var resultText = "";
        for (var i = 0; i < phoneText.length; i++) {
            resultText += (PHONE_HYPHENS.indexOf(phoneText.charAt(i)) >= 0) ? hyphenText : phoneText.charAt(i);
        }
        return resultText;
    }

    /**
     * テキストの変換内容を決める。書き換え（後ろから順）と、書き換え後の縦中横の位置を返す
     * @param {string} sourceText - 元の文字列
     * @param {Object} formatOptions - 変換の設定
     * @returns {{edits: Array<{start: number, length: number, text: string}>, tcySpans: Array<{start: number, length: number}>,
     *     changes: Array<{before: string, after: string, tcy: boolean, tcyLength: number}>, outputLength: number}}
     *     edits は後ろから順、changes は確認の一覧用（前から順）
     */
    function planVerticalText(sourceText, formatOptions) {
        var text = formatOptions.halfwidthAlnum ? toHalfwidthAlnum(sourceText) : sourceText;
        var edits = [];
        var tcySpans = [];
        var changes = [];
        var outputLength = 0;
        var segmentText = "";
        var segmentHasTcy = false;
        var segmentTcyLength = 0;
        var quoteState = { isOpen: false };

        /**
         * 出力に文字列を足す（縦中横なら位置を控える）
         * @param {string} pieceText - 足す文字列
         * @param {boolean} isTcy - 縦中横にするか
         * @returns {void}
         */
        function emit(pieceText, isTcy) {
            if (isTcy) {
                tcySpans.push({ start: outputLength, length: pieceText.length });
                segmentHasTcy = true;
                segmentTcyLength = pieceText.length;
            }
            outputLength += pieceText.length;
            segmentText += pieceText;
        }

        /**
         * 元の [start, start+length) を出力 segmentText に置き換えたことを控える（変わらなければ控えない）
         * @param {number} start - 元の位置
         * @param {number} length - 元の文字数
         * @returns {void}
         */
        function closeSegment(start, length) {
            var originalText = sourceText.substr(start, length);
            if (segmentText !== originalText) {
                edits.push({ start: start, length: length, text: segmentText });
            }
            if (segmentText !== originalText || segmentHasTcy) {
                changes.push({ before: originalText, after: segmentText, tcy: segmentHasTcy, tcyLength: segmentTcyLength });
            }
            segmentText = "";
            segmentHasTcy = false;
            segmentTcyLength = 0;
        }

        var i = 0;
        while (i < text.length) {
            var currentChar = text.charAt(i);
            var runStart = i;
            var westernYear = (formatOptions.japaneseEra && isDigitChar(currentChar)) ? matchWesternYear(text, i) : null;
            var phraseText = westernYear ? null : matchLatinPhrase(text, i, formatOptions);
            if (westernYear) {
                emit(westernYear.eraName, false);
                /* 1年は「元年」 / The first year is written 元 */
                if (westernYear.eraYear === 1) emit("元", false);
                else emitPlainDigits(String(westernYear.eraYear), formatOptions, emit);
                i += 4;
            } else if (phraseText) {
                /* 英単語の連続（2語以上。3D Printing のように数字で始まる語も）は横倒しの欧文として残す（回転しない）。
                   中の全角英数字・幅の広いスペース・「＆」は半角に戻す
                   A run of two or more words (including ones like 3D Printing) stays sideways (not rotated);
                   full-width alphanumerics, wide spaces and ＆ inside become half-width */
                emit(toHalfwidthAlnum(phraseText).replace(/[\u2000-\u2003\u3000]/g, " ").replace(/＆/g, "&"), false);
                i += phraseText.length;
            } else if (isDigitChar(currentChar) && matchPhoneNumber(text, i)) {
                /* 電話番号・郵便番号は範囲の「〜」にも縦中横・漢数字にもしない / Phone and postal numbers are not ranges, TCY or kanji */
                var phoneText = matchPhoneNumber(text, i);
                emit(convertPhoneHyphens(phoneText, formatOptions), false);
                i += phoneText.length;
            } else if (isDigitChar(currentChar)) {
                var numberToken = readNumberToken(text, i);
                var tokenEnd = i + numberToken.length;
                var isMergedComma = numberToken.hasComma && !numberToken.hasDecimal && formatOptions.commaStyle === "remove";
                if (formatOptions.kanji && numberToken.integerDigits >= formatOptions.kanjiMinDigits && numberToken.integerDigits <= formatOptions.kanjiMaxDigits) {
                    /* 西暦の年（直後が「年」）は書き方によらず1字ずつ / A year (followed by 年) is always digit by digit */
                    emit(convertNumberToKanji(numberToken, formatOptions, text.charAt(tokenEnd) === "年"), false);
                } else if (!numberToken.hasSeparator || isMergedComma) {
                    /* 桁区切りを取ったときは、つないだ数字を1つの数として扱う / With separators removed, the joined digits are one number */
                    var plainDigits = numberToken.hasSeparator ? text.substring(i, tokenEnd).replace(/[,，]/g, "") : text.substring(i, tokenEnd);
                    emitDigitRun(plainDigits, formatOptions, emit);
                } else if (formatOptions.tcyExcludeAdjacent) {
                    /* 桁区切り・小数点の隣の数字は縦中横・全角にしない / Leave numbers next to separators as they are */
                    for (var p = 0; p < numberToken.parts.length; p++) {
                        var part = numberToken.parts[p];
                        emit(part.type === "digits" ? part.text : convertNumberMark(part, formatOptions), false);
                    }
                } else {
                    for (var q = 0; q < numberToken.parts.length; q++) {
                        var numberPart = numberToken.parts[q];
                        if (numberPart.type === "digits") emitDigitRun(numberPart.text, formatOptions, emit);
                        else emit(convertNumberMark(numberPart, formatOptions), false);
                    }
                }
                i = tokenEnd;
            } else if (isLatinLetter(currentChar)) {
                /* 1語は半角だけを変換する（全角の英字はそのままでよい） / Single words: only half-width letters need converting */
                while (i < text.length && isLatinLetter(text.charAt(i))) i++;
                var letterText = text.substring(runStart, i);
                var isCapsTcy = formatOptions.tcyCapitals && letterText.length >= 2 && letterText.length <= 3 && letterText === letterText.toUpperCase();
                var isAfterDigit = isAnyDigitChar(text.charAt(runStart - 1));
                if (formatOptions.unitsKumimoji && isAfterDigit && UNIT_KUMIMOJI.hasOwnProperty(letterText)) emit(UNIT_KUMIMOJI[letterText], false);
                else if (isCapsTcy) emit(letterText, true);
                else if (isShortWordForFullwidth(letterText, formatOptions)) emit(toFullwidthAscii(letterText), false);
                else emit(letterText, false);
            } else if (isExclamOrQuestion(currentChar)) {
                while (i < text.length && isExclamOrQuestion(text.charAt(i))) i++;
                var markText = text.substring(runStart, i);
                if (formatOptions.tcyExclamation && markText.length === 2) {
                    /* 縦中横にする組は半角にそろえる / Use half-width marks for the Tate-Chu-Yoko pair */
                    emit(markText.replace(/！/g, "!").replace(/？/g, "?"), true);
                } else if (formatOptions.exclamFullwidth) {
                    emit(toFullwidthAscii(markText), false);
                } else {
                    emit(markText, false);
                }
            } else {
                emit(convertOtherChar(text, i, formatOptions, quoteState), false);
                i++;
            }
            closeSegment(runStart, i - runStart);
        }
        edits.reverse();
        return { edits: edits, tcySpans: tcySpans, changes: changes, outputLength: outputLength };
    }

    // =========================================
    // 対象のテキスト / Text targets
    // =========================================

    /**
     * テキストフレームが縦組みか
     * @param {TextFrame} textFrame - 判定するフレーム
     * @returns {boolean} 縦組みなら true
     */
    function isVerticalFrame(textFrame) {
        return !!textFrame && textFrame.orientation === TextOrientation.VERTICAL;
    }

    /**
     * 選択から処理の対象を集める。文字選択中はその範囲、オブジェクト選択ならテキストフレームごと
     * @param {Document} doc - 対象のドキュメント
     * @returns {Array<Object>} { frame（連結していないフレーム）, story, ownerFrames（属するテキストフレーム）, offset（ストーリーの文字の番号のずれ）, text, isTextSelection }
     */
    function collectTextTargets(doc) {
        var selection = doc.selection;
        var textTargets = [];
        /* 文字選択中は selection が TextRange（[0] が無い） / While text is selected, selection is a TextRange without [0] */
        if (selection && selection.typename === "TextRange") {
            if (selection.length > 0) {
                textTargets.push({
                    story: selection.story,
                    offset: selection.characterOffset - 1,
                    text: selection.contents,
                    ownerFrames: getStoryFrames(selection.story),
                    isTextSelection: true
                });
            }
            return textTargets;
        }
        var textFrames = collectSelectionTextFrames(selection, { skipLocked: true, skipHidden: true });
        var storyHeadFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            var storyFrames = getStoryFrames(textFrames[i].story);
            if (storyFrames.length <= 1) {
                textTargets.push({ frame: textFrames[i], story: textFrames[i].story, offset: 0, text: textFrames[i].contents, ownerFrames: [textFrames[i]], isTextSelection: false });
                continue;
            }
            /* 連結テキストはストーリー全体を1つの対象にする（書き換えで文字が次のフレームへ流れても位置がずれない）
               Threaded text is one target for the whole story, so reflow between frames does not shift positions */
            if (indexOfValue(storyHeadFrames, storyFrames[0]) >= 0) continue;
            storyHeadFrames.push(storyFrames[0]);
            textTargets.push({ story: textFrames[i].story, offset: 0, text: textFrames[i].story.textRange.contents, ownerFrames: storyFrames, isTextSelection: false });
        }
        return textTargets;
    }

    /**
     * ストーリーのテキストフレームを配列で返す
     * @param {Story} textStory - ストーリー
     * @returns {TextFrame[]} 連結しているテキストフレーム（連結していなければ1つ）
     */
    function getStoryFrames(textStory) {
        var storyFrames = [];
        for (var i = 0; i < textStory.textFrames.length; i++) storyFrames.push(textStory.textFrames[i]);
        return storyFrames;
    }

    /**
     * 対象の文字のコレクションを返す（文字選択はストーリー全体の番号で扱う）
     * @param {Object} textTarget - 対象
     * @returns {Characters} 文字のコレクション
     */
    function getTargetCharacters(textTarget) {
        return textTarget.frame ? textTarget.frame.characters : textTarget.story.characters;
    }

    /**
     * 書き換えを適用する。文字数が同じなら1文字ずつ、異なるなら余分を消して先頭の文字に入れる（書式は先頭の文字のもの）
     * @param {Object} textTarget - 対象
     * @param {Array<Object>} edits - planVerticalText() の edits（後ろから順）
     * @returns {void}
     */
    function applyTextEdits(textTarget, edits) {
        for (var i = 0; i < edits.length; i++) {
            var characters = getTargetCharacters(textTarget);
            var baseIndex = textTarget.offset + edits[i].start;
            var editText = edits[i].text;
            if (editText.length === edits[i].length) {
                for (var k = 0; k < editText.length; k++) {
                    if (characters[baseIndex + k].contents !== editText.charAt(k)) characters[baseIndex + k].contents = editText.charAt(k);
                }
            } else {
                for (var r = edits[i].length - 1; r >= 1; r--) characters[baseIndex + r].remove();
                characters[baseIndex].contents = editText;
            }
        }
    }

    /**
     * 書き換えを元の文字列に当てて、書き換え後の文字列を返す
     * @param {string} sourceText - 元の文字列
     * @param {Array<Object>} edits - planVerticalText() の edits（後ろから順）
     * @returns {string} 書き換え後の文字列
     */
    function buildEditedText(sourceText, edits) {
        var editedText = sourceText;
        for (var i = 0; i < edits.length; i++) {
            editedText = editedText.substring(0, edits[i].start) + edits[i].text + editedText.substring(edits[i].start + edits[i].length);
        }
        return editedText;
    }

    /**
     * 全角の英字（Ａ〜Ｚ・ａ〜ｚ）の並びを探す
     * @param {string} targetText - 対象の文字列
     * @returns {Array<{start: number, length: number}>} 見つかった位置
     */
    function findFullwidthLatinRuns(targetText) {
        var latinRuns = [];
        var latinPattern = /[Ａ-Ｚａ-ｚ]+/g;
        var latinMatch;
        while ((latinMatch = latinPattern.exec(targetText)) !== null) {
            latinRuns.push({ start: latinMatch.index, length: latinMatch[0].length });
        }
        return latinRuns;
    }

    /**
     * 回転させない欧文を探す。英単語がスペースや「&」でつながった2語以上のまとまりは常に、
     * 1語は minChars 文字以上のとき（minChars が 0 なら1語は対象外）
     * @param {string} targetText - 対象の文字列
     * @param {number} minChars - 1語を対象にする文字数（0 で対象にしない）
     * @param {Object} formatOptions - 変換の設定
     * @returns {Array<{start: number, length: number}>} 見つかった位置
     */
    function findNoRotateLatinRuns(targetText, minChars, formatOptions) {
        var latinRuns = [];
        var i = 0;
        while (i < targetText.length) {
            var latinRun = scanLatinRun(targetText, i, formatOptions);
            if (!latinRun) {
                i++;
                continue;
            }
            var runLength = latinRun.end - i;
            if (latinRun.wordCount >= 2 || (minChars > 0 && runLength >= minChars)) latinRuns.push({ start: i, length: runLength });
            i = latinRun.end;
        }
        return latinRuns;
    }

    /**
     * 回転させない欧文の位置を、設定に合わせて返す（欧文回転がオフなら空）
     * @param {string} targetText - 対象の文字列（書き換え前）
     * @param {Object} formatOptions - 変換の設定
     * @returns {Array<{start: number, length: number}>} 見つかった位置
     */
    function getNoRotateRuns(targetText, formatOptions) {
        if (!formatOptions.rotateRoman) return [];
        return findNoRotateLatinRuns(targetText, formatOptions.rotateSkipLongLatin ? formatOptions.rotateSkipMinChars : 0, formatOptions);
    }

    /**
     * 書き換え後のテキストの全角の英字に、プロポーショナルメトリクスを適用する（選択はしない）
     * @param {Object} textTarget - 対象
     * @param {Object} plan - planVerticalText() の戻り値
     * @returns {void}
     */
    function applyProportionalMetricsToFullwidthLatin(textTarget, plan) {
        var latinRuns = findFullwidthLatinRuns(buildEditedText(textTarget.text, plan.edits));
        for (var i = 0; i < latinRuns.length; i++) {
            var latinRange = getTargetCharacters(textTarget)[textTarget.offset + latinRuns[i].start];
            latinRange.length = latinRuns[i].length;
            latinRange.characterAttributes.proportionalMetrics = true;
        }
    }

    /**
     * 対象の中の範囲だけを選択する
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} textTarget - 対象
     * @param {number} start - 範囲の先頭（対象の中の位置）
     * @param {number} length - 文字数
     * @returns {TextRange} 選択した範囲
     */
    function selectTargetSpan(doc, textTarget, start, length) {
        /* select() は選択に足されるので、先に外す / select() adds to the selection, so clear it first */
        doc.selection = null;
        var spanRange = getTargetCharacters(textTarget)[textTarget.offset + start];
        spanRange.length = length;
        spanRange.select();
        return spanRange;
    }

    /* 縦中横の字数ごとの等幅字形（確認の一覧の字形名のキーも兼ねる） / Fixed-width form per Tate-Chu-Yoko length (also the label key) */
    var TCY_GLYPH_KEYS_BY_LENGTH = { 2: "half", 3: "third", 4: "quarter" };

    /**
     * 縦中横の字数に合わせた等幅字形を返す（2字は半角・3字は三分・4字は四分で、ほぼ全角1字分に収める）
     * @param {number} charCount - 縦中横の字数
     * @returns {AlternateGlyphsForm|null} 字形（2〜4字以外は null）
     */
    function getTcyGlyphForm(charCount) {
        if (charCount === 2) return AlternateGlyphsForm.HALFWIDTH;
        if (charCount === 3) return AlternateGlyphsForm.THIRDWIDTH;
        if (charCount === 4) return AlternateGlyphsForm.QUARTERWIDTH;
        return null;
    }

    /**
     * 範囲に等幅字形を適用する。直接の代入が効かなければ、一時的な文字スタイル経由で入れる
     * （文字スタイルの方法は uske-s「Illustratorで字形を置換するスクリプト」による）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextRange} spanRange - 対象の範囲
     * @param {AlternateGlyphsForm} glyphForm - 字形
     * @returns {void}
     */
    function applyGlyphForm(doc, spanRange, glyphForm) {
        try {
            spanRange.characterAttributes.alternateGlyphs = glyphForm;
            if (spanRange.characterAttributes.alternateGlyphs === glyphForm) return;
        } catch (e) {
            $.writeln("applyGlyphForm (direct): " + e);
        }
        var tempStyle = null;
        try {
            tempStyle = doc.characterStyles.add(SCRIPT_NAME + "_" + new Date().getTime());
            tempStyle.characterAttributes.alternateGlyphs = glyphForm;
            /* false：手動の書式は残す / false keeps the local formatting */
            tempStyle.applyTo(spanRange, false);
        } catch (e) {
            $.writeln("applyGlyphForm (style): " + e);
        } finally {
            if (tempStyle) {
                try { tempStyle.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
            }
        }
    }

    /**
     * ドキュメントにあるカスタムの文字組みアキ量設定の名前を返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {Array<string>} 名前（doc.mojikumiSet の番号7以降の順）
     */
    function listCustomMojikumiNames(doc) {
        var customNames = [];
        var mojikumiSets = doc.mojikumiSet;
        for (var i = MOJIKUMI_BUILTIN_COUNT; i < mojikumiSets.length; i++) customNames.push(String(mojikumiSets[i]));
        return customNames;
    }

    /**
     * 文字組みアキ量設定の設定から、段落に代入する値を返す。カスタムの名前がドキュメントにあればそれを、無ければ組み込みの番号を使う
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} formatOptions - 変換の設定（mojikumiIndex：-1 は「なし」、null は変えない／mojikumiCustomName）
     * @returns {*} 代入する値（変えないとき・ドキュメントに無いときは null）
     */
    function resolveMojikumiValue(doc, formatOptions) {
        var mojikumiIndex = formatOptions.mojikumiIndex;
        if (mojikumiIndex === null) return null;
        if (formatOptions.mojikumiCustomName) {
            var customIndex = indexOfValue(listCustomMojikumiNames(doc), formatOptions.mojikumiCustomName);
            if (customIndex >= 0) return doc.mojikumiSet[MOJIKUMI_BUILTIN_COUNT + customIndex];
        }
        if (mojikumiIndex === -1) return "なし";
        return (mojikumiIndex >= 0 && mojikumiIndex < doc.mojikumiSet.length) ? doc.mojikumiSet[mojikumiIndex] : null;
    }

    /**
     * 対象の段落に、禁則・文字組みアキ量設定・行送り（自動行送りの％）と行頭行末揃え（最終行上）を設定する
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} textTarget - 対象
     * @param {Object} formatOptions - 変換の設定（kinsokuType / mojikumiIndex（mojikumiCustomName） / leadingPercent。null は変えない）
     * @returns {void}
     */
    function applyParagraphSettings(doc, textTarget, formatOptions) {
        var mojikumiValue = resolveMojikumiValue(doc, formatOptions);
        var paragraphs;
        if (textTarget.frame) {
            paragraphs = textTarget.frame.paragraphs;
        } else if (!textTarget.isTextSelection) {
            paragraphs = textTarget.story.textRange.paragraphs;
        } else {
            /* 文字選択は、いまのストーリーから範囲を取り直す / Re-take a text selection from the current story */
            var selectionRange = textTarget.story.characters[textTarget.offset];
            selectionRange.length = textTarget.text.length;
            paragraphs = selectionRange.paragraphs;
        }
        for (var i = 0; i < paragraphs.length; i++) {
            /* 末尾の空の段落は length に数えられるが取り出すと No such element になる / A trailing empty paragraph is counted but throws "No such element" */
            var paragraphAttributes;
            try {
                paragraphAttributes = paragraphs[i].paragraphAttributes;
            } catch (e) {
                continue;
            }
            if (mojikumiValue !== null) paragraphAttributes.mojikumi = mojikumiValue;
            if (formatOptions.kinsokuType !== null) paragraphAttributes.kinsoku = formatOptions.kinsokuType;
            /* 縦組みでは「最終行左揃え」が最終行上揃え / In vertical text, "last line left" means last line top */
            paragraphAttributes.justification = Justification.FULLJUSTIFYLASTLINELEFT;
            /* 行送りの％は段落の属性。文字の自動行送りは処理の最後に applyAutoLeadingToTarget() でかける
               The leading percentage is a paragraph attribute; characters get auto leading at the end (applyAutoLeadingToTarget) */
            if (formatOptions.leadingPercent !== null) paragraphAttributes.autoLeadingAmount = formatOptions.leadingPercent;
        }
    }

    /**
     * 処理したテキスト全体の文字を自動行送りにする（行送りが常に「自動」になり、文字サイズに追従する。UnifiedTypePalette と同じ考え方）
     * 書き換え・縦中横・字形の変更のあとにかけるので、途中で書式が変わった文字も自動になる
     * @param {Object} textTarget - 対象
     * @param {Object} plan - planVerticalText() の戻り値
     * @returns {void}
     */
    function applyAutoLeadingToTarget(textTarget, plan) {
        if (plan.outputLength === 0) return;
        var outputRange = getTargetCharacters(textTarget)[textTarget.offset];
        outputRange.length = plan.outputLength;
        outputRange.characterAttributes.autoLeading = true;
    }

    /**
     * 処理の対象をまとめて選択する（文字選択はその範囲、オブジェクト選択はテキストフレーム）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {number} rangeLength - 文字選択のときの範囲の文字数
     * @returns {void}
     */
    function selectTextTargets(doc, textTargets, rangeLength) {
        if (textTargets[0].isTextSelection) {
            selectTargetSpan(doc, textTargets[0], 0, rangeLength);
            return;
        }
        /* 連結テキストは連結しているフレームをすべて選ぶ / Select every frame of threaded text */
        doc.selection = null;
        for (var i = 0; i < textTargets.length; i++) {
            for (var j = 0; j < textTargets[i].ownerFrames.length; j++) textTargets[i].ownerFrames[j].selected = true;
        }
    }

    /**
     * 処理のあと、選択を元の対象に戻す。自動サイズ調整したフレームは、アクションのあとに取り直したものを使う
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {number} rangeLength - 文字選択のときの範囲の文字数
     * @param {TextFrame[]|null} autoSizedFrames - applyAreaAutoSize() の戻り値
     * @returns {void}
     */
    function restoreFinalSelection(doc, textTargets, rangeLength, autoSizedFrames) {
        if (textTargets[0].isTextSelection) {
            if (autoSizedFrames && autoSizedFrames.length > 0) textTargets[0].story = autoSizedFrames[0].story;
            selectTargetSpan(doc, textTargets[0], 0, rangeLength);
            return;
        }
        doc.selection = null;
        if (autoSizedFrames) {
            for (var i = 0; i < autoSizedFrames.length; i++) autoSizedFrames[i].selected = true;
        }
        for (var j = 0; j < textTargets.length; j++) {
            if (textTargets[j].isAutoSized) continue;
            for (var k = 0; k < textTargets[j].ownerFrames.length; k++) textTargets[j].ownerFrames[k].selected = true;
        }
    }

    /**
     * 対象の段落の設定をする（文字の書式には触れない。縦組みへの変換は convertTargetsToVertical() で先に行う）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {Object} formatOptions - 変換の設定
     * @returns {void}
     */
    function applyFrameSettings(doc, textTargets, formatOptions) {
        for (var i = 0; i < textTargets.length; i++) applyParagraphSettings(doc, textTargets[i], formatOptions);
    }

    /**
     * 横組みのテキストフレームを縦組みにし、選択から対象を集め直す。
     * 向きを変えるとテキストフレームが作り直され、変える前の参照は使えなくなる（there is no document）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - collectTextTargets() の戻り値
     * @returns {Array<Object>} 集め直した対象（横組みが無ければ渡したまま）
     */
    function convertTargetsToVertical(doc, textTargets) {
        var hasHorizontal = false;
        for (var i = 0; i < textTargets.length && !hasHorizontal; i++) {
            for (var j = 0; j < textTargets[i].ownerFrames.length; j++) {
                if (!isVerticalFrame(textTargets[i].ownerFrames[j])) hasHorizontal = true;
            }
        }
        if (!hasHorizontal) return textTargets;

        var textSelectionTarget = textTargets[0].isTextSelection ? textTargets[0] : null;
        doc.selection = null;
        for (var k = 0; k < textTargets.length; k++) {
            for (var m = 0; m < textTargets[k].ownerFrames.length; m++) textTargets[k].ownerFrames[m].selected = true;
        }
        /* 向きを変えるたびに選択を読み直し、横組みが残っていれば1つずつ変える（作り直されたフレームの参照を使わない）
           Re-read the selection after each change and convert one horizontal frame at a time (never reuse a recreated frame) */
        var frameCount = collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true }).length;
        for (var guard = 0; guard < frameCount; guard++) {
            var currentFrames = collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true });
            var horizontalFrame = null;
            for (var n = 0; n < currentFrames.length && !horizontalFrame; n++) {
                if (!isVerticalFrame(currentFrames[n])) horizontalFrame = currentFrames[n];
            }
            if (!horizontalFrame) break;
            horizontalFrame.orientation = TextOrientation.VERTICAL;
        }

        if (!textSelectionTarget) return collectTextTargets(doc);
        /* 文字選択は、作り直したフレームのストーリーで同じ位置を指す / A text selection points at the same offset in the new story */
        var newFrames = collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true });
        var newStory = newFrames[0].story;
        return [{
            story: newStory,
            offset: textSelectionTarget.offset,
            text: textSelectionTarget.text,
            ownerFrames: getStoryFrames(newStory),
            isTextSelection: true
        }];
    }

    /**
     * 縦中横の箇所を1つずつ選んで縦中横（と字数に合わせた等幅字形）を適用する
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {Array<Object>} plans - planVerticalText() の戻り値（対象と同じ順）
     * @param {Object} formatOptions - 変換の設定
     * @param {function(number): void} onSpanDone - 1か所終わるごとに呼ぶ（引数は終わった数）
     * @returns {void}
     */
    function applyTcySpans(doc, textTargets, plans, formatOptions, onSpanDone) {
        var tcyDone = 0;
        for (var i = 0; i < textTargets.length; i++) {
            var tcySpans = plans[i].tcySpans;
            for (var j = 0; j < tcySpans.length; j++) {
                var tcyRange = selectTargetSpan(doc, textTargets[i], tcySpans[j].start, tcySpans[j].length);
                app.doScript(TCY_ACTION_NAME, ACTION_SET_NAME, false);
                var glyphForm = formatOptions.tcyGlyphByLength ? getTcyGlyphForm(tcySpans[j].length) : null;
                if (glyphForm !== null) applyGlyphForm(doc, tcyRange, glyphForm);
                tcyDone++;
                onSpanDone(tcyDone);
            }
        }
    }

    /**
     * エリア内文字の枠の高さ（縦組みでは1行の長さ）を、文字サイズの整数倍にそろえる（いちばん近い倍数、最小1文字分）
     * 内側の余白（spacing）は除いて測り、足し戻す
     * @param {Array<Object>} textTargets - 対象
     * @returns {void}
     */
    function adjustAreaHeightToTextSize(textTargets) {
        for (var i = 0; i < textTargets.length; i++) {
            var ownerFrames = textTargets[i].ownerFrames;
            for (var j = 0; j < ownerFrames.length; j++) {
                var areaFrame = ownerFrames[j];
                if (areaFrame.kind !== TextType.AREATEXT || !areaFrame.editable || areaFrame.characters.length === 0) continue;
                var textSize = areaFrame.characters[0].characterAttributes.size;
                if (!(textSize > 0)) continue;
                var insetSpacing = areaFrame.spacing || 0;
                var contentHeight = areaFrame.textPath.height - insetSpacing * 2;
                var charCount = Math.max(1, Math.round(contentHeight / textSize));
                areaFrame.textPath.height = charCount * textSize + insetSpacing * 2;
            }
        }
    }

    /**
     * 対象のうち、連結していないエリア内文字に自動サイズ調整を適用する（連結したテキストでは使えない）
     * 適用した対象には isAutoSized を付ける
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @returns {TextFrame[]|null} アクションのあとに選択から取り直したフレーム（適用しなかったら null）
     */
    function applyAreaAutoSize(doc, textTargets) {
        var areaFrames = [];
        for (var i = 0; i < textTargets.length; i++) {
            var ownerFrames = textTargets[i].ownerFrames;
            if (ownerFrames.length === 1 && ownerFrames[0].kind === TextType.AREATEXT && ownerFrames[0].editable) areaFrames.push(ownerFrames[0]);
        }
        if (areaFrames.length === 0) return null;
        for (var t = 0; t < textTargets.length; t++) {
            if (indexOfValue(areaFrames, textTargets[t].ownerFrames[0]) >= 0) textTargets[t].isAutoSized = true;
        }
        doc.selection = null;
        for (var j = 0; j < areaFrames.length; j++) areaFrames[j].selected = true;
        app.doScript(AREA_AUTO_SIZE_ACTION_NAME, ACTION_SET_NAME, false);
        /* アクションのあとのフレームは選択から取り直す / Re-take the frames from the selection after the action */
        return collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true });
    }

    /**
     * 段落の設定・欧文回転・文字の書き換え・縦中横を順に適用する
     * （欧文回転は書き換えの前に全体へ掛け、書き換えた文字は前の文字の書式として回転を引き継ぐ）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {Object} formatOptions - 変換の設定
     * @returns {string} 失敗したときの理由（成功は空文字）
     */
    function runVerticalTextConversion(doc, textTargets, formatOptions) {
        /* 縦組みにするとフレームが作り直されるので、最初に変えて対象を集め直す / Converting recreates frames, so do it first and re-collect */
        try {
            textTargets = convertTargetsToVertical(doc, textTargets);
        } catch (convertError) {
            return convertError + " (line " + convertError.line + ")";
        }
        var plans = [];
        var tcyTotal = 0;
        for (var i = 0; i < textTargets.length; i++) {
            plans.push(planVerticalText(textTargets[i].text, formatOptions));
            tcyTotal += plans[i].tcySpans.length;
        }
        var noRotateRunsList = [];
        for (var r = 0; r < textTargets.length; r++) {
            noRotateRunsList.push(getNoRotateRuns(textTargets[r].text, formatOptions));
        }
        var needsAction = formatOptions.rotateRoman || tcyTotal > 0 || formatOptions.areaAutoSize;
        var autoSizedFrames = null;
        var stepTotal = (formatOptions.rotateRoman ? 1 : 0) + textTargets.length + tcyTotal + (formatOptions.areaAutoSize ? 1 : 0);
        var stepDone = 0;
        var progress = createProgressWindow(getLabel("progress.title"));

        /**
         * 進み具合を1つ進めて表示する
         * @param {string} message - 表示する文言
         * @returns {void}
         */
        function advanceProgress(message) {
            stepDone++;
            progress.set(stepDone / stepTotal, message);
        }

        if (needsAction && !loadTemporaryActionSet(buildVerticalTextActions(), ACTION_SET_NAME)) {
            progress.close();
            return "loadTemporaryActionSet";
        }
        try {
            applyFrameSettings(doc, textTargets, formatOptions);

            /* まず欧文回転を、書き換え前のテキスト全体に / Rotate first, over the whole text before any edits */
            if (formatOptions.rotateRoman) {
                progress.set(0, getLabel("progress.rotate"));
                selectTextTargets(doc, textTargets, textTargets[0].text.length);
                app.doScript(ROTATE_ACTION_NAME, ACTION_SET_NAME, false);
                /* 長い欧文は回転を外す（書き換え前の位置で） / Turn rotation off for long Roman text (positions before edits) */
                for (var k = 0; k < textTargets.length; k++) {
                    var noRotateRuns = noRotateRunsList[k];
                    for (var n = 0; n < noRotateRuns.length; n++) {
                        selectTargetSpan(doc, textTargets[k], noRotateRuns[n].start, noRotateRuns[n].length);
                        app.doScript(ROTATE_OFF_ACTION_NAME, ACTION_SET_NAME, false);
                    }
                }
                advanceProgress(getLabel("progress.rotate"));
            }

            for (var j = 0; j < textTargets.length; j++) {
                applyTextEdits(textTargets[j], plans[j].edits);
                if (formatOptions.proportionalFullwidthLatin) applyProportionalMetricsToFullwidthLatin(textTargets[j], plans[j]);
                advanceProgress(getLabel("progress.edit", { current: j + 1, total: textTargets.length }));
            }

            applyTcySpans(doc, textTargets, plans, formatOptions, function (tcyDone) {
                advanceProgress(getLabel("progress.tcy", { current: tcyDone, total: tcyTotal }));
            });

            /* 行送りは常に自動に（簡易・詳細とも） / Leading is always auto (both modes) */
            for (var a = 0; a < textTargets.length; a++) applyAutoLeadingToTarget(textTargets[a], plans[a]);

            /* 枠の高さ（1行の長さ）を文字サイズの整数倍に。自動サイズ調整より前に / Height to a multiple of the type size, before auto size */
            if (formatOptions.areaHeightToSize) adjustAreaHeightToTextSize(textTargets);

            /* 文字数が変わったあとの大きさに合わせる / Fit the frames after the text has changed */
            if (formatOptions.areaAutoSize) {
                autoSizedFrames = applyAreaAutoSize(doc, textTargets);
                advanceProgress(getLabel("progress.autoSize"));
            }

            return "";
        } catch (e) {
            $.writeln("runVerticalTextConversion: " + e);
            return e + " (line " + e.line + ")";
        } finally {
            if (needsAction) unloadTemporaryActionSet(ACTION_SET_NAME);
            progress.close();
            /* 元の選択に戻す / Restore the original selection */
            try {
                restoreFinalSelection(doc, textTargets, plans[0].outputLength, autoSizedFrames);
            } catch (restoreError) { /* 戻せなくても続ける / keep going */ }
        }
    }

    // =========================================
    // 確認用の色分け（ColorNumberByWidth から移植） / Check coloring (ported from ColorNumberByWidth)
    // =========================================

    /* このダイアログで選んだ色（モードを切り替えて作り直しても引き継ぐ） / Colors chosen in this run (kept across mode switches) */
    var checkColors = null;

    /**
     * 確認用の色分けの色を用意する（まだ無ければ初期色で作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {{baseColor: Color, fullWidthColor: Color, halfWidthColor: Color, glyphWidthColor: Color}} 色
     */
    function ensureCheckColors(doc) {
        if (!checkColors) {
            checkColors = {
                baseColor: createDocumentColor(doc, CHECK_BASE_RGB, CHECK_BASE_CMYK),
                fullWidthColor: createDocumentColor(doc, CHECK_FULL_WIDTH_RGB, CHECK_FULL_WIDTH_CMYK),
                halfWidthColor: createDocumentColor(doc, CHECK_HALF_WIDTH_RGB, CHECK_HALF_WIDTH_CMYK),
                glyphWidthColor: createDocumentColor(doc, CHECK_GLYPH_WIDTH_RGB, CHECK_GLYPH_WIDTH_CMYK)
            };
        }
        return checkColors;
    }

    /**
     * ドキュメントのカラーモードに合わせて色を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} rgbValues - R, G, B
     * @param {number[]} cmykValues - C, M, Y, K
     * @returns {Color} RGBColor または CMYKColor
     */
    function createDocumentColor(doc, rgbValues, cmykValues) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = cmykValues[0];
            cmykColor.magenta = cmykValues[1];
            cmykColor.yellow = cmykValues[2];
            cmykColor.black = cmykValues[3];
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = rgbValues[0];
        rgbColor.green = rgbValues[1];
        rgbColor.blue = rgbValues[2];
        return rgbColor;
    }

    /**
     * 色を色見本に描くための RGB（0〜1）にする（CMYK・グレーは簡易換算）
     * @param {Color} swatchColor - 色
     * @returns {number[]} [r, g, b]（0〜1）
     */
    function colorToScreenRgb(swatchColor) {
        if (swatchColor.typename === "RGBColor") return [swatchColor.red / 255, swatchColor.green / 255, swatchColor.blue / 255];
        if (swatchColor.typename === "GrayColor") return [1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100];
        if (swatchColor.typename === "CMYKColor") {
            var blackRatio = 1 - swatchColor.black / 100;
            return [
                (1 - swatchColor.cyan / 100) * blackRatio,
                (1 - swatchColor.magenta / 100) * blackRatio,
                (1 - swatchColor.yellow / 100) * blackRatio
            ];
        }
        /* スポットカラーなどは灰色で示す / show other color types in gray */
        return [0.5, 0.5, 0.5];
    }

    /**
     * 等幅字形（等幅全角・半角・1/3・1/4字形）の AlternateGlyphsForm を集める
     * @returns {Array} AlternateGlyphsForm の値
     */
    function getFixedWidthGlyphForms() {
        return [
            AlternateGlyphsForm.FULLWIDTH,
            AlternateGlyphsForm.HALFWIDTH,
            AlternateGlyphsForm.THIRDWIDTH,
            AlternateGlyphsForm.QUARTERWIDTH
        ];
    }

    /**
     * 文字が等幅字形になっているかを返す
     * @param {TextRange} targetCharacter - 1文字の範囲
     * @param {Array} glyphForms - getFixedWidthGlyphForms() の戻り値
     * @returns {boolean} 等幅字形なら true
     */
    function isFixedWidthGlyph(targetCharacter, glyphForms) {
        var glyphForm;
        /* 字形を読めない文字は等幅字形ではないとみなす / treat characters whose glyph form cannot be read as not fixed-width */
        try {
            glyphForm = targetCharacter.characterAttributes.alternateGlyphs;
        } catch (e) {
            return false;
        }
        for (var i = 0; i < glyphForms.length; i++) {
            if (glyphForm === glyphForms[i]) return true;
        }
        return false;
    }

    /**
     * テキスト範囲を塗り分ける（全体をベースの色にし、等幅字形の文字を緑、英数字と「!?」を全角は赤・半角は青）
     * @param {TextRange} targetRange - 対象のテキスト範囲
     * @param {{baseColor: Color, fullWidthColor: Color, halfWidthColor: Color, glyphWidthColor: Color}} chosenColors - 塗る色
     * @param {Array} glyphForms - getFixedWidthGlyphForms() の戻り値
     * @returns {void}
     */
    function colorCharactersInRange(targetRange, chosenColors, glyphForms) {
        var rangeCharacters = targetRange.characters;
        var characterCount = rangeCharacters.length;

        /* 先に全体をベースの色にし、当たる文字だけ塗り直す / paint the whole range with the base color first, then recolor matches */
        targetRange.characterAttributes.fillColor = chosenColors.baseColor;

        for (var i = 0; i < characterCount; i++) {
            /* 文字は1字ずつ読む（まとめて読んだ contents は書き換え前のことがある） / Read each character; a bulk contents read can be stale */
            var characterText = rangeCharacters[i].contents;
            var characterColor = null;
            /* 等幅字形は全角・半角より優先する / fixed-width glyphs take priority over full- and half-width */
            if (isFixedWidthGlyph(rangeCharacters[i], glyphForms)) {
                characterColor = chosenColors.glyphWidthColor;
            } else if (FULL_WIDTH_PATTERN.test(characterText)) {
                characterColor = chosenColors.fullWidthColor;
            } else if (HALF_WIDTH_PATTERN.test(characterText)) {
                characterColor = chosenColors.halfWidthColor;
            }
            if (!characterColor) continue;
            rangeCharacters[i].characterAttributes.fillColor = characterColor;
        }
    }

    /**
     * 処理が終わって元に戻した選択から、テキストを集め直して色分けする（ColorNumberByWidth と同じ手順）
     * 処理中のフレームの参照は、一時アクションのあとで古い内容を返すことがあるので使わない
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function colorCheckSelection(doc) {
        app.redraw();
        ensureCheckColors(doc);
        var glyphForms = getFixedWidthGlyphForms();
        var currentSelection = doc.selection;
        var targetRanges = [];
        if (currentSelection && currentSelection.typename === "TextRange") {
            targetRanges.push(currentSelection);
        } else {
            var textFrames = collectSelectionTextFrames(currentSelection, { skipLocked: true, skipHidden: true });
            for (var i = 0; i < textFrames.length; i++) targetRanges.push(textFrames[i].textRange);
        }
        for (var j = 0; j < targetRanges.length; j++) colorCharactersInRange(targetRanges[j], checkColors, glyphForms);
    }

    /**
     * 項目名と、クリックでカラーピッカーを開く色見本の1行を作る
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.fieldLabel / LABELS.tooltip のキー
     * @param {Color} initialColor - 最初の色
     * @param {function(Color): void} [onColorChange] - 色を変えたときに呼ぶ
     * @returns {{label: StaticText, getColor: Function}} 項目名と、今の色を返す getColor を持つオブジェクト
     */
    function addColorSwatchRow(parentPanel, labelKey, initialColor, onColorChange) {
        var swatchColor = initialColor;
        var swatchRow = parentPanel.add("group");
        setupRow(swatchRow, "left", 8);

        var rowLabel = swatchRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));

        var colorSwatch = swatchRow.add("group");
        colorSwatch.preferredSize = COLOR_SWATCH_SIZE;
        colorSwatch.helpTip = getLabel("tooltip." + labelKey);
        colorSwatch.onDraw = function () {
            var swatchGraphics = colorSwatch.graphics;
            var swatchWidth = colorSwatch.size[0];
            var swatchHeight = colorSwatch.size[1];
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, colorToScreenRgb(swatchColor)));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1));
        };
        /* クリックで標準のカラーピッカーを開く。キャンセルは渡した色が返る
           Clicking opens the standard color picker; Cancel returns the color passed in */
        colorSwatch.addEventListener("click", function () {
            swatchColor = app.showColorPicker(swatchColor);
            if (onColorChange) onColorChange(swatchColor);
            colorSwatch.hide(); /* group には notify() が無いので、隠して再表示して描き直す / redraw the group */
            colorSwatch.show();
        });

        return {
            label: rowLabel,
            getColor: function () { return swatchColor; }
        };
    }

    // =========================================
    // プログレスバー / Progress bar
    // =========================================

    /**
     * 進捗表示用のパレットウィンドウを生成して表示する（非モーダルなので処理中に更新できる）
     * 戻り値の set(fraction 0..1, message) で進捗を更新、close() で閉じる
     * @param {string} title - ウィンドウのタイトル
     * @returns {{set: function, close: function}} 進捗の更新と終了の操作
     */
    function createProgressWindow(title) {
        var progressWindow = new Window("palette", title);
        setupWindow(progressWindow, 8);
        var progressBar = progressWindow.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize = [PROGRESS_WIDTH, 8];
        var progressLabel = progressWindow.add("statictext", undefined, "");
        progressLabel.preferredSize.width = PROGRESS_WIDTH;
        progressWindow.show();
        progressWindow.update();
        return {
            set: function (fraction, message) {
                var percent = Math.round(fraction * 100);
                progressBar.value = (percent < 0) ? 0 : (percent > 100 ? 100 : percent);
                if (message != null) { progressLabel.text = message; }
                progressWindow.update(); /* 同期ループ中でも再描画させる / force a repaint even inside a synchronous loop */
            },
            close: function () { progressWindow.close(); }
        };
    }

    /**
     * 確認の一覧で縦中横の箇所に付ける印を返す（字形を変えるときは字形名も入れる）
     * @param {boolean} glyphByLength - 字数に合わせた等幅字形にするか
     * @param {number} charCount - 縦中横の字数
     * @returns {string} 印
     */
    function getTcyPreviewMark(glyphByLength, charCount) {
        var glyphLabelKey = glyphByLength ? TCY_GLYPH_KEYS_BY_LENGTH[charCount] : undefined;
        if (glyphLabelKey === undefined) return getLabel("preview.tcyMark");
        return getLabel("preview.tcyGlyphMark", { glyph: getLabel("glyphShort." + glyphLabelKey) });
    }

    /**
     * 確認の一覧に出す行を作る（同じ変換はまとめて件数を付ける）
     * @param {Array<Object>} textTargets - 対象
     * @param {Object} formatOptions - 変換の設定
     * @returns {string[]} 一覧の行
     */
    function buildPreviewLines(textTargets, formatOptions) {
        var lineKeys = [];
        var lineCounts = {};

        /**
         * 一覧に1行足す（同じ行は件数を数える）
         * @param {string} lineText - 行の文字列
         * @returns {void}
         */
        function addPreviewLine(lineText) {
            if (lineCounts.hasOwnProperty(lineText)) {
                lineCounts[lineText]++;
            } else {
                lineCounts[lineText] = 1;
                lineKeys.push(lineText);
            }
        }

        for (var i = 0; i < textTargets.length; i++) {
            var noRotateRuns = getNoRotateRuns(textTargets[i].text, formatOptions);
            for (var r = 0; r < noRotateRuns.length; r++) {
                addPreviewLine(textTargets[i].text.substr(noRotateRuns[r].start, noRotateRuns[r].length) + getLabel("preview.noRotateMark"));
            }
            var changes = planVerticalText(textTargets[i].text, formatOptions).changes;
            for (var c = 0; c < changes.length; c++) {
                var change = changes[c];
                var lineText = (change.before === change.after) ? change.before : change.before + " → " + change.after;
                if (change.tcy) lineText += getTcyPreviewMark(formatOptions.tcyGlyphByLength, change.tcyLength);
                addPreviewLine(lineText);
            }
        }
        var previewLines = [];
        for (var k = 0; k < lineKeys.length && k < PREVIEW_MAX_LINES; k++) {
            previewLines.push(lineKeys[k] + (lineCounts[lineKeys[k]] > 1 ? "  ×" + lineCounts[lineKeys[k]] : ""));
        }
        return previewLines;
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
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
    // プリセット / Presets
    // =========================================
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "Presets", "persistent");

    /**
     * 保存したプリセットを返す
     * @returns {Object} プリセット名をキーにした設定の集まり（書き換えても保存されない）
     */
    function loadPresetMap() {
        return presetSettingsStore.load({});
    }

    /**
     * 保存したプリセットの名前を、名前順で返す
     * @returns {string[]} プリセット名
     */
    function getPresetNames() {
        var presetMap = loadPresetMap();
        var presetNames = [];
        for (var presetName in presetMap) {
            if (presetMap.hasOwnProperty(presetName)) presetNames.push(presetName);
        }
        presetNames.sort();
        return presetNames;
    }

    /**
     * 名前が組み込みのプリセットと同じか
     * @param {string} presetName - プリセット名
     * @returns {boolean} 同じなら true
     */
    function isBuiltinPresetName(presetName) {
        for (var i = 0; i < BUILTIN_PRESETS.length; i++) {
            if (getLabel("preset.builtin." + BUILTIN_PRESETS[i].labelKey) === presetName) return true;
        }
        return false;
    }

    /**
     * 既定値にプリセットの値を重ねた設定を返す
     * @param {Object} presetData - プリセットの値（一部だけでもよい）
     * @returns {Object} DEFAULT_SETTINGS と同じ形
     */
    function mergePresetData(presetData) {
        var mergedSettings = {};
        for (var settingKey in DEFAULT_SETTINGS) {
            if (!DEFAULT_SETTINGS.hasOwnProperty(settingKey)) continue;
            var hasValue = presetData.hasOwnProperty(settingKey) && typeof presetData[settingKey] === typeof DEFAULT_SETTINGS[settingKey];
            mergedSettings[settingKey] = hasValue ? presetData[settingKey] : DEFAULT_SETTINGS[settingKey];
        }
        return mergedSettings;
    }

    /**
     * プリセット名を尋ねる
     * @param {string} initialName - 入力欄に入れておく名前
     * @returns {string|null} プリセット名。キャンセルか空なら null
     */
    function showPresetNameDialog(initialName) {
        var nameDialog = new Window("dialog", getLabel("preset.saveTitle"));
        setupWindow(nameDialog);
        var nameRow = nameDialog.add("group");
        setupRow(nameRow, ["left", "center"], 8);
        nameRow.add("statictext", undefined, labelText("preset.name"));
        var nameInput = nameRow.add("edittext", undefined, initialName);
        nameInput.helpTip = getLabel("tooltip.presetName");
        nameInput.characters = PRESET_NAME_CHARS;
        nameInput.active = true;
        var buttonRow = addButtonRow(nameDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;
        prepareDialogWindow(nameDialog, SCRIPT_NAME + "_presetName");
        if (nameDialog.show() !== 1) return null;
        var presetName = nameInput.text.replace(/^\s+|\s+$/g, "");
        if (!presetName || presetName === getLabel("preset.placeholder")) return null;
        if (isBuiltinPresetName(presetName)) {
            alert(getLabel("alert.presetNameTaken"));
            return null;
        }
        return presetName;
    }

    /**
     * 今の設定を名前を付けて保存する（同じ名前は確認してから上書き）
     * @param {Object} currentSettings - readDialogOptions() の戻り値
     * @param {string} initialName - 入力欄に入れておく名前
     * @returns {string|null} 保存したプリセット名。保存しなかったら null
     */
    function saveCurrentPreset(currentSettings, initialName) {
        var presetName = showPresetNameDialog(initialName);
        if (!presetName) return null;
        var presetMap = loadPresetMap();
        if (presetMap.hasOwnProperty(presetName) && !confirm(getLabel("confirm.presetOverwrite", { name: presetName }), true)) return null;
        presetMap[presetName] = currentSettings;
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel("alert.presetSaveFailed"));
            return null;
        }
        return presetName;
    }

    /**
     * 保存したプリセットを確認してから削除する
     * @param {string} presetName - プリセット名
     * @returns {boolean} 削除したら true
     */
    function deletePreset(presetName) {
        if (!confirm(getLabel("confirm.presetDelete", { name: presetName }), true)) return false;
        var presetMap = loadPresetMap();
        delete presetMap[presetName];
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel("alert.presetSaveFailed"));
            return false;
        }
        return true;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * チェックボックスを追加する
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelPath - LABELS のパス
     * @param {boolean} value - 初期値
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parent, labelPath, value) {
        var optionCheckbox = parent.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.value = value;
        /* 説明があればツールチップに / Show the description as a tooltip when there is one */
        var tooltipEntry = LABELS.tooltip[labelPath.split(".")[1]];
        if (tooltipEntry) optionCheckbox.helpTip = getLabel(tooltipEntry);
        return optionCheckbox;
    }

    /**
     * 字下げしたグループ（親のチェックボックスに従う項目用）を追加する
     * @param {Group|Panel} parent - 追加先
     * @returns {Group} グループ
     */
    function addIndentedGroup(parent) {
        var indentedGroup = parent.add("group");
        indentedGroup.orientation = "column";
        indentedGroup.alignChildren = ["left", "top"];
        indentedGroup.margins = [OPTION_INDENT, 0, 0, 0];
        indentedGroup.spacing = 6;
        return indentedGroup;
    }

    /* 数値欄の∧∨・↑↓で呼ぶ処理（bindDialogEvents で入れる） / Handler for the steppers and arrow keys (set in bindDialogEvents) */
    var digitStepHandler = null;

    /**
     * 数値欄の∧∨・↑↓で増減したことを知らせる
     * @param {EditText} numberInput - 増減した入力欄
     * @returns {void}
     */
    function notifyDigitStep(numberInput) {
        if (digitStepHandler) digitStepHandler(numberInput);
    }

    /**
     * 「項目名 : [ドロップダウン]」の行を追加する
     * @param {Panel} parent - 追加先
     * @param {string} labelPath - 項目名の LABELS のパス
     * @param {Object[]} itemLabels - 選択肢の { ja, en } の配列
     * @param {string} tooltipPath - ツールチップの LABELS のパス
     * @param {number} [labelWidth] - 項目名の幅（右揃え）。省くと文字に合わせる
     * @returns {DropDownList} ドロップダウン
     */
    function addLabeledDropdown(parent, labelPath, itemLabels, tooltipPath, labelWidth) {
        var dropdownRow = parent.add("group");
        setupRow(dropdownRow, ["left", "center"], 8);
        var rowLabel = dropdownRow.add("statictext", undefined, labelText(labelPath));
        if (labelWidth) {
            rowLabel.preferredSize.width = labelWidth;
            rowLabel.justify = "right";
        }
        var optionDropdown = dropdownRow.add("dropdownlist", undefined, []);
        for (var i = 0; i < itemLabels.length; i++) optionDropdown.add("item", getLabel(itemLabels[i]));
        optionDropdown.helpTip = getLabel(tooltipPath);
        return optionDropdown;
    }

    /**
     * 配列の中で値が一致する位置を返す（ES3 に Array.indexOf は無い）
     * @param {Array} values - 配列
     * @param {*} targetValue - 探す値
     * @returns {number} 位置（無ければ -1）
     */
    function indexOfValue(values, targetValue) {
        for (var i = 0; i < values.length; i++) {
            if (values[i] === targetValue) return i;
        }
        return -1;
    }

    /**
     * 「桁数 : [最小] 〜 [最大]」の行を追加する
     * @param {Group|Panel} parent - 追加先
     * @returns {{minInput: EditText, maxInput: EditText}} 入力欄
     */
    function addDigitRangeRow(parent) {
        var rangeRow = parent.add("group");
        setupRow(rangeRow, ["left", "center"], 4);
        var minInput = addSteppedField(rangeRow, { label: labelText("fieldLabel.digits"), text: "", characters: 3, step: 1, min: 1, max: 99, integer: true, onStep: notifyDigitStep });
        var maxInput = addSteppedField(rangeRow, { label: getLabel("fieldLabel.rangeTo"), text: "", characters: 3, step: 1, min: 1, max: 99, integer: true, onStep: notifyDigitStep });
        return { minInput: minInput, maxInput: maxInput };
    }

    /**
     * 縦に並べる列のグループを追加する
     * @param {Group|Window} parent - 追加先
     * @returns {Group} 列のグループ
     */
    function addLayoutColumn(parent) {
        var layoutColumn = parent.add("group");
        layoutColumn.orientation = "column";
        layoutColumn.alignChildren = ["fill", "top"];
        layoutColumn.spacing = WINDOW_SPACING;
        return layoutColumn;
    }

    /**
     * プリセットの行（ドロップダウン・保存・削除）を追加する
     * @param {Window} dialog - ダイアログ
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildPresetRow(dialog, dialogControls) {
        var presetRow = dialog.add("group");
        setupRow(presetRow, "center", 8); /* ダイアログボックスの左右中央に置く / centered in the dialog */
        presetRow.add("statictext", undefined, labelText("preset.label"));
        dialogControls.presetDropdown = presetRow.add("dropdownlist", undefined, []);
        dialogControls.presetDropdown.helpTip = getLabel("tooltip.preset");
        dialogControls.presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
        dialogControls.presetSaveIcon = addIconButton(presetRow, PRESET_ICON_SIZE, drawSaveIcon);
        dialogControls.presetSaveIcon.helpTip = getLabel("tooltip.presetSave");
        dialogControls.presetDeleteIcon = addIconButton(presetRow, PRESET_ICON_SIZE, drawTrashIcon);
        dialogControls.presetDeleteIcon.helpTip = getLabel("tooltip.presetDelete");
    }

    /**
     * 「欧文回転・縦中横」パネルを追加する（簡易モードは数字の縦中横と等幅字形まで）
     * @param {Group} parent - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @param {Object} settings - 初期値
     * @param {boolean} isDetail - 詳細モードか
     * @returns {void}
     */
    function buildTcyPanel(parent, dialogControls, settings, isDetail) {
        var tcyPanel = parent.add("panel", undefined, getLabel("panel.rotateTcy"));
        setupPanel(tcyPanel, 6);
        dialogControls.rotateRoman = addOptionCheckbox(tcyPanel, "checkbox.rotateRoman", settings.rotateRoman);
        /* 長い欧文は除く ［5］文字以上 / Skip long Roman text [5] or more characters */
        var rotateSkipRow = addIndentedGroup(tcyPanel).add("group");
        setupRow(rotateSkipRow, ["left", "center"], 4);
        dialogControls.rotateSkipLongLatin = addOptionCheckbox(rotateSkipRow, "checkbox.rotateSkipLongLatin", settings.rotateSkipLongLatin);
        dialogControls.rotateSkipMinInput = addSteppedField(rotateSkipRow, { label: "", text: "", characters: 3, step: 1, min: 2, max: 99, integer: true, onStep: notifyDigitStep });
        rotateSkipRow.add("statictext", undefined, getLabel("fieldLabel.charsOrMore"));
        dialogControls.tcyNumbers = addOptionCheckbox(tcyPanel, "checkbox.tcyNumbers", settings.tcyNumbers);
        /* 桁数は次の行に字下げして並べる / Digit counts go on the next, indented row */
        var tcyNumberGroup = addIndentedGroup(tcyPanel);
        var tcyDigitRow = tcyNumberGroup.add("group");
        setupRow(tcyDigitRow, ["left", "center"], 8);
        dialogControls.tcyDigits2 = addOptionCheckbox(tcyDigitRow, "checkbox.tcyDigits2", settings.tcyDigits2);
        dialogControls.tcyDigits3 = addOptionCheckbox(tcyDigitRow, "checkbox.tcyDigits3", settings.tcyDigits3);
        dialogControls.tcyDigits4 = addOptionCheckbox(tcyDigitRow, "checkbox.tcyDigits4", settings.tcyDigits4);
        if (isDetail) {
            dialogControls.tcyExcludeAdjacent = addOptionCheckbox(tcyNumberGroup, "checkbox.tcyExcludeAdjacent", settings.tcyExcludeAdjacent);
            dialogControls.tcyCapitals = addOptionCheckbox(tcyPanel, "checkbox.tcyCapitals", settings.tcyCapitals);
            dialogControls.tcyExclamation = addOptionCheckbox(tcyPanel, "checkbox.tcyExclamation", settings.tcyExclamation);
        }
        dialogControls.tcyGlyphByLength = addOptionCheckbox(tcyPanel, "checkbox.tcyGlyphByLength", settings.tcyGlyphByLength);
    }

    /**
     * 「文字の変換」パネルを追加する（パネルの中は1列）。
     * 簡易モードは桁区切り・小数点・1文字の英字・1桁の数字だけ（「数字」パネルを作らない）
     * @param {Group} parent - 追加先の列
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @param {Object} settings - 初期値
     * @param {boolean} isDetail - 詳細モードか
     * @returns {void}
     */
    function buildReplacePanel(parent, dialogControls, settings, isDetail) {
        var replacePanel = parent.add("panel", undefined, getLabel("panel.replace"));
        setupPanel(replacePanel, 6);
        dialogControls.commaDropdown = addLabeledDropdown(replacePanel, "fieldLabel.comma", LABELS.dropdown.comma, "tooltip.comma");
        dialogControls.periodToNakaguro = addOptionCheckbox(replacePanel, "checkbox.periodToNakaguro", settings.periodToNakaguro);
        if (isDetail) {
            dialogControls.rangeToWave = addOptionCheckbox(replacePanel, "checkbox.rangeToWave", settings.rangeToWave);
            dialogControls.halfwidthAlnum = addOptionCheckbox(replacePanel, "checkbox.halfwidthAlnum", settings.halfwidthAlnum);
            dialogControls.symbolsFullwidth = addOptionCheckbox(replacePanel, "checkbox.symbolsFullwidth", settings.symbolsFullwidth);
            dialogControls.quotesToChonchon = addOptionCheckbox(replacePanel, "checkbox.quotesToChonchon", settings.quotesToChonchon);
            dialogControls.exclamFullwidth = addOptionCheckbox(replacePanel, "checkbox.exclamFullwidth", settings.exclamFullwidth);
        }
        dialogControls.singleLetterFullwidth = addOptionCheckbox(replacePanel, "checkbox.singleLetterFullwidth", settings.singleLetterFullwidth);
        if (isDetail) {
            /* 短い英単語を全角に ［2］文字以下 / Short words to full-width [2] or fewer characters */
            var shortWordRow = replacePanel.add("group");
            setupRow(shortWordRow, ["left", "center"], 4);
            dialogControls.shortWordFullwidth = addOptionCheckbox(shortWordRow, "checkbox.shortWordFullwidth", settings.shortWordFullwidth);
            dialogControls.shortWordMaxInput = addSteppedField(shortWordRow, { label: "", text: "", characters: 3, step: 1, min: 1, max: 98, integer: true, onStep: notifyDigitStep });
            shortWordRow.add("statictext", undefined, getLabel("fieldLabel.charsOrLess"));
        }
        /* 簡易モードは「数字」パネルを作らず、1桁の全角もここに入れる / The simple mode has no Numbers panel; single digits go here */
        if (!isDetail) dialogControls.singleDigitFullwidth = addOptionCheckbox(replacePanel, "checkbox.singleDigitFullwidth", settings.singleDigitFullwidth);
    }

    /**
     * 「数字」パネル（詳細モードだけ）を追加する
     * @param {Group} parent - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @param {Object} settings - 初期値
     * @returns {void}
     */
    function buildNumberPanel(parent, dialogControls, settings) {
        var numberPanel = parent.add("panel", undefined, getLabel("panel.numbers"));
        setupPanel(numberPanel, 6);
        dialogControls.singleDigitFullwidth = addOptionCheckbox(numberPanel, "checkbox.singleDigitFullwidth", settings.singleDigitFullwidth);
        dialogControls.unitsKumimoji = addOptionCheckbox(numberPanel, "checkbox.unitsKumimoji", settings.unitsKumimoji);
        dialogControls.japaneseEra = addOptionCheckbox(numberPanel, "checkbox.japaneseEra", settings.japaneseEra);
        dialogControls.kanji = addOptionCheckbox(numberPanel, "checkbox.kanji", settings.kanji);
        var kanjiGroup = addIndentedGroup(numberPanel);
        var kanjiRange = addDigitRangeRow(kanjiGroup);
        dialogControls.kanjiMinInput = kanjiRange.minInput;
        dialogControls.kanjiMaxInput = kanjiRange.maxInput;
        /* ラジオは同じ親の中だけ排他になるので、1つのグループに並べる / Radios are exclusive only within one parent */
        var kanjiStyleRow = kanjiGroup.add("group");
        setupRow(kanjiStyleRow, ["left", "center"], 8);
        dialogControls.kanjiDigitRadio = kanjiStyleRow.add("radiobutton", undefined, getLabel("radio.kanjiDigit"));
        dialogControls.kanjiPositionalRadio = kanjiStyleRow.add("radiobutton", undefined, getLabel("radio.kanjiPositional"));
        dialogControls.kanjiDigitRadio.helpTip = getLabel("tooltip.kanjiDigit");
        dialogControls.kanjiPositionalRadio.helpTip = getLabel("tooltip.kanjiPositional");
        dialogControls.kanjiLargeUnitsRadio = kanjiStyleRow.add("radiobutton", undefined, getLabel("radio.kanjiLargeUnits"));
        dialogControls.kanjiLargeUnitsRadio.helpTip = getLabel("tooltip.kanjiLargeUnits");
        var kanjiDetailRow = kanjiGroup.add("group");
        setupRow(kanjiDetailRow, ["left", "center"], 8);
        dialogControls.kanjiZeroRei = addOptionCheckbox(kanjiDetailRow, "checkbox.kanjiZeroRei", settings.kanjiZeroRei);
        dialogControls.kanjiIchiSen = addOptionCheckbox(kanjiDetailRow, "checkbox.kanjiIchiSen", settings.kanjiIchiSen);
        dialogControls.phoneHyphenDropdown = addLabeledDropdown(numberPanel, "fieldLabel.phoneHyphen", LABELS.dropdown.phoneHyphen, "tooltip.phoneHyphen");
    }

    /**
     * 「文字・段落」パネル（詳細モードだけ）を追加する
     * @param {Group} parent - 追加先
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @param {Object} settings - 初期値
     * @returns {void}
     */
    function buildParagraphPanel(parent, dialogControls, settings) {
        var paragraphPanel = parent.add("panel", undefined, getLabel("panel.paragraph"));
        setupPanel(paragraphPanel, 6);
        dialogControls.kinsokuDropdown = addLabeledDropdown(paragraphPanel, "fieldLabel.kinsoku", LABELS.dropdown.kinsoku, "tooltip.kinsoku", PARAGRAPH_LABEL_WIDTH);
        dialogControls.mojikumiDropdown = addLabeledDropdown(paragraphPanel, "fieldLabel.mojikumi", LABELS.dropdown.mojikumi, "tooltip.mojikumi", PARAGRAPH_LABEL_WIDTH);
        /* 組み込みのあとにドキュメントのカスタムのセットを並べる / Custom sets of the document follow the built-in ones */
        dialogControls.customMojikumiNames = listCustomMojikumiNames(app.activeDocument);
        for (var m = 0; m < dialogControls.customMojikumiNames.length; m++) dialogControls.mojikumiDropdown.add("item", dialogControls.customMojikumiNames[m]);
        /* 行送り : [150%]（自動行送りの％） / Leading: [150%] (auto-leading percentage) */
        dialogControls.leadingInput = addSteppedField(paragraphPanel, {
            label: labelText("fieldLabel.leading"), labelWidth: PARAGRAPH_LABEL_WIDTH,
            text: "", characters: 5, step: 1, min: 50, max: 500, integer: true, unit: "%"
        });
        dialogControls.leadingInput.helpTip = getLabel("tooltip.leading");
        /* チェックボックスはドロップダウン・行送りの欄と間を空ける / Keep a gap between the fields and the checkboxes */
        var paragraphOptionGroup = paragraphPanel.add("group");
        paragraphOptionGroup.orientation = "column";
        paragraphOptionGroup.alignChildren = ["left", "top"];
        paragraphOptionGroup.spacing = 6;
        paragraphOptionGroup.margins = [0, PARAGRAPH_OPTION_TOP_MARGIN, 0, 0];
        /* 文字の書き換えではなく OpenType の書式なので、ここに置く / An OpenType setting, not a rewrite, so it lives here */
        dialogControls.proportionalFullwidthLatin = addOptionCheckbox(paragraphOptionGroup, "checkbox.proportionalFullwidthLatin", settings.proportionalFullwidthLatin);
        dialogControls.areaHeightToSize = addOptionCheckbox(paragraphOptionGroup, "checkbox.areaHeightToSize", settings.areaHeightToSize);
        dialogControls.areaAutoSize = addOptionCheckbox(paragraphOptionGroup, "checkbox.areaAutoSize", settings.areaAutoSize);
    }

    /**
     * 「確認用の色分け」パネル（詳細モードだけ）を追加する。色見本は2段組（2つずつ2行）、項目名は右揃え
     * @param {Group} parent - 追加先の列
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @param {Object} settings - 初期値
     * @returns {void}
     */
    function buildColorCheckPanel(parent, dialogControls, settings) {
        var colorCheckPanel = parent.add("panel", undefined, getLabel("panel.colorCheck"));
        setupPanel(colorCheckPanel, 6);
        dialogControls.colorCheck = addOptionCheckbox(colorCheckPanel, "checkbox.colorCheck", settings.colorCheck);
        var chosenColors = ensureCheckColors(app.activeDocument);
        var swatchSpecs = [
            ["colorBase", "baseColor"], ["colorFullWidth", "fullWidthColor"],
            ["colorHalfWidth", "halfWidthColor"], ["colorGlyphWidth", "glyphWidthColor"]
        ];
        /* 2段組：2つずつ2行に並べ、項目名は右揃え / Two columns: two swatches per row, labels right-aligned */
        for (var i = 0; i < swatchSpecs.length; i += 2) {
            var swatchPairRow = colorCheckPanel.add("group");
            setupRow(swatchPairRow, ["left", "center"], 12);
            for (var j = i; j < i + 2; j++) {
                var swatchRowParts = addColorSwatchRow(swatchPairRow, swatchSpecs[j][0], chosenColors[swatchSpecs[j][1]], createColorSetter(swatchSpecs[j][1]));
                swatchRowParts.label.preferredSize.width = COLOR_LABEL_WIDTH;
                swatchRowParts.label.justify = "right";
            }
        }
    }

    /**
     * 色見本で選んだ色を checkColors に書き込む関数を返す
     * @param {string} colorKey - checkColors のキー
     * @returns {function(Color): void} 書き込む関数
     */
    function createColorSetter(colorKey) {
        return function (chosenColor) {
            checkColors[colorKey] = chosenColor;
        };
    }

    /**
     * ボタン行（左：モードの切り替え、右：キャンセル・OK）を追加する
     * @param {Window} dialog - ダイアログ
     * @param {Object} dialogControls - コントロールを入れるオブジェクト
     * @param {boolean} isDetail - 詳細モードか
     * @returns {void}
     */
    function buildButtonRow(dialog, dialogControls, isDetail) {
        var buttonRow = addButtonRow(dialog);
        dialogControls.modeSimpleRadio = buttonRow.leftGroup.add("radiobutton", undefined, getLabel("radio.modeSimple"));
        dialogControls.modeDetailRadio = buttonRow.leftGroup.add("radiobutton", undefined, getLabel("radio.modeDetail"));
        dialogControls.modeSimpleRadio.helpTip = getLabel("tooltip.modeSimple");
        dialogControls.modeDetailRadio.helpTip = getLabel("tooltip.modeDetail");
        dialogControls.modeSimpleRadio.value = !isDetail;
        dialogControls.modeDetailRadio.value = isDetail;
        dialogControls.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogControls.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        dialogControls.btnCancel.preferredSize.width = BUTTON_WIDTH;
        dialogControls.btnOK.preferredSize.width = BUTTON_WIDTH;
    }

    /**
     * ダイアログを組み立てる。簡易モードはよく使う項目だけを並べ、ほかの項目はコントロールを作らない
     * @param {Object} settings - 初期値（DEFAULT_SETTINGS の形）
     * @returns {Object} dialog と各コントロール（作らなかった項目は settings の値を baseSettings に持つ）
     */
    function buildDialog(settings) {
        var isDetail = (settings.dialogMode === "detail");
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);
        var dialogControls = { dialog: dialog, baseSettings: mergePresetData(settings), dialogMode: isDetail ? "detail" : "simple" };

        buildPresetRow(dialog, dialogControls);

        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = addLayoutColumn(columnsGroup);
        buildTcyPanel(leftColumn, dialogControls, settings, isDetail);
        if (isDetail) {
            /* 左：欧文回転・縦中横、文字・段落、確認用の色分け／右：数字、文字の変換
               Left: rotation, TCY, character & paragraph, check coloring; right: numbers, character conversion */
            buildParagraphPanel(leftColumn, dialogControls, settings);
            buildColorCheckPanel(leftColumn, dialogControls, settings);
            var rightColumn = addLayoutColumn(columnsGroup);
            buildNumberPanel(rightColumn, dialogControls, settings);
            buildReplacePanel(rightColumn, dialogControls, settings, isDetail);
        } else {
            buildReplacePanel(leftColumn, dialogControls, settings, isDetail);
        }

        /* 変換の確認 / Conversion check */
        var previewPanel = dialog.add("panel", undefined, getLabel("panel.preview"));
        setupPanel(previewPanel);
        dialogControls.previewList = previewPanel.add("listbox", undefined, []);
        dialogControls.previewList.preferredSize.height = PREVIEW_LIST_HEIGHT;

        buildButtonRow(dialog, dialogControls, isDetail);
        /* 数値欄・ラジオ・ドロップダウンの初期値はまとめて入れる / Fill fields, radios and dropdowns in one place */
        writeDialogSettings(dialogControls, settings);
        return dialogControls;
    }

    /**
     * 数値欄の整数を読む（読めなければ既定値）
     * @param {EditText} numberInput - 入力欄
     * @param {number} fallbackValue - 読めないときの値
     * @returns {number} 整数
     */
    function readDigitInput(numberInput, fallbackValue) {
        var value = parseInt(numberInput.text, 10);
        return isNaN(value) ? fallbackValue : value;
    }

    /**
     * ダイアログの値を設定のオブジェクトにする（最小と最大が逆なら入れ替える。作らなかった項目は baseSettings の値）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {Object} DEFAULT_SETTINGS と同じ形
     */
    function readDialogOptions(dialogControls) {
        var formatOptions = mergePresetData(dialogControls.baseSettings);
        for (var i = 0; i < CHECKBOX_SETTING_KEYS.length; i++) {
            var settingKey = CHECKBOX_SETTING_KEYS[i];
            if (dialogControls[settingKey]) formatOptions[settingKey] = dialogControls[settingKey].value;
        }
        if (dialogControls.kanjiMinInput) {
            var kanjiMin = readDigitInput(dialogControls.kanjiMinInput, DEFAULT_SETTINGS.kanjiMinDigits);
            var kanjiMax = readDigitInput(dialogControls.kanjiMaxInput, DEFAULT_SETTINGS.kanjiMaxDigits);
            formatOptions.kanjiMinDigits = Math.min(kanjiMin, kanjiMax);
            formatOptions.kanjiMaxDigits = Math.max(kanjiMin, kanjiMax);
        }
        formatOptions.rotateSkipMinChars = Math.max(2, readDigitInput(dialogControls.rotateSkipMinInput, DEFAULT_SETTINGS.rotateSkipMinChars));
        /* 短い英単語は「長い欧文は除く」の1つ下に連動 / Short words are linked to one below "Skip long Roman text" */
        formatOptions.shortWordMaxChars = formatOptions.rotateSkipMinChars - 1;
        if (dialogControls.kanjiPositionalRadio) formatOptions.kanjiStyle = readKanjiStyle(dialogControls);
        if (dialogControls.commaDropdown.selection) formatOptions.commaStyle = COMMA_STYLES[dialogControls.commaDropdown.selection.index];
        if (dialogControls.phoneHyphenDropdown && dialogControls.phoneHyphenDropdown.selection) formatOptions.phoneHyphenStyle = PHONE_HYPHEN_STYLES[dialogControls.phoneHyphenDropdown.selection.index];
        if (dialogControls.kinsokuDropdown && dialogControls.kinsokuDropdown.selection) formatOptions.kinsokuType = KINSOKU_TYPES[dialogControls.kinsokuDropdown.selection.index];
        if (dialogControls.leadingInput) formatOptions.leadingPercent = Math.min(500, Math.max(50, readDigitInput(dialogControls.leadingInput, DEFAULT_SETTINGS.leadingPercent)));
        if (dialogControls.mojikumiDropdown && dialogControls.mojikumiDropdown.selection) {
            var mojikumiItemIndex = dialogControls.mojikumiDropdown.selection.index;
            if (mojikumiItemIndex < MOJIKUMI_INDEXES.length) {
                formatOptions.mojikumiIndex = MOJIKUMI_INDEXES[mojikumiItemIndex];
                formatOptions.mojikumiCustomName = "";
            } else {
                /* カスタムは名前で控える（番号はドキュメントごとに異なる）。番号には組み込みの既定を残す
                   Custom sets are kept by name (indexes differ per document); the index keeps the built-in default */
                formatOptions.mojikumiIndex = DEFAULT_SETTINGS.mojikumiIndex;
                formatOptions.mojikumiCustomName = dialogControls.customMojikumiNames[mojikumiItemIndex - MOJIKUMI_INDEXES.length];
            }
        }
        formatOptions.dialogMode = dialogControls.dialogMode;
        return formatOptions;
    }

    /**
     * 漢数字の書き方のラジオから値を返す
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {string} "digit" / "positional" / "largeUnits"
     */
    function readKanjiStyle(dialogControls) {
        if (dialogControls.kanjiPositionalRadio.value) return "positional";
        if (dialogControls.kanjiLargeUnitsRadio.value) return "largeUnits";
        return "digit";
    }

    /**
     * 実行と確認の一覧に使う設定を返す。簡易モードでは、表示していない項目をオフにし、禁則・文字組みは変えない
     * （保存する設定は readDialogOptions() のまま、詳細モードの値を残す）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {Object} 実行に使う設定
     */
    function readEffectiveOptions(dialogControls) {
        var formatOptions = readDialogOptions(dialogControls);
        if (dialogControls.dialogMode !== "simple") return formatOptions;
        for (var i = 0; i < CHECKBOX_SETTING_KEYS.length; i++) {
            var settingKey = CHECKBOX_SETTING_KEYS[i];
            if (dialogControls[settingKey] || indexOfValue(ALWAYS_APPLIED_KEYS, settingKey) >= 0) continue;
            formatOptions[settingKey] = false;
        }
        formatOptions.kinsokuType = null;
        formatOptions.leadingPercent = null;
        formatOptions.phoneHyphenStyle = "keep";
        formatOptions.mojikumiIndex = null;
        return formatOptions;
    }

    /**
     * 設定の値をダイアログの各項目に書き込む（作っていない項目は baseSettings に持つ）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Object} settings - DEFAULT_SETTINGS と同じ形
     * @returns {void}
     */
    function writeDialogSettings(dialogControls, settings) {
        dialogControls.baseSettings = mergePresetData(settings);
        for (var i = 0; i < CHECKBOX_SETTING_KEYS.length; i++) {
            var settingKey = CHECKBOX_SETTING_KEYS[i];
            if (dialogControls[settingKey]) dialogControls[settingKey].value = settings[settingKey];
        }
        if (dialogControls.kanjiMinInput) {
            dialogControls.kanjiMinInput.text = String(settings.kanjiMinDigits);
            dialogControls.kanjiMaxInput.text = String(settings.kanjiMaxDigits);
        }
        dialogControls.rotateSkipMinInput.text = String(settings.rotateSkipMinChars);
        if (dialogControls.shortWordMaxInput) dialogControls.shortWordMaxInput.text = String(settings.rotateSkipMinChars - 1);
        if (dialogControls.kanjiPositionalRadio) {
            dialogControls.kanjiPositionalRadio.value = (settings.kanjiStyle === "positional");
            dialogControls.kanjiLargeUnitsRadio.value = (settings.kanjiStyle === "largeUnits");
            dialogControls.kanjiDigitRadio.value = !dialogControls.kanjiPositionalRadio.value && !dialogControls.kanjiLargeUnitsRadio.value;
        }
        dialogControls.commaDropdown.selection = Math.max(0, indexOfValue(COMMA_STYLES, settings.commaStyle));
        if (dialogControls.phoneHyphenDropdown) dialogControls.phoneHyphenDropdown.selection = Math.max(0, indexOfValue(PHONE_HYPHEN_STYLES, settings.phoneHyphenStyle));
        if (dialogControls.leadingInput) dialogControls.leadingInput.text = settings.leadingPercent + "%";
        if (dialogControls.kinsokuDropdown) dialogControls.kinsokuDropdown.selection = Math.max(0, indexOfValue(KINSOKU_TYPES, settings.kinsokuType));
        if (dialogControls.mojikumiDropdown) {
            var customItemIndex = settings.mojikumiCustomName ? indexOfValue(dialogControls.customMojikumiNames, settings.mojikumiCustomName) : -1;
            dialogControls.mojikumiDropdown.selection = (customItemIndex >= 0)
                ? MOJIKUMI_INDEXES.length + customItemIndex
                : Math.max(0, indexOfValue(MOJIKUMI_INDEXES, settings.mojikumiIndex));
        }
    }

    /**
     * ダイアログのイベントをつなぐ（有効／無効の切り替えと確認の一覧の更新）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Array<Object>} textTargets - collectTextTargets() の戻り値
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, textTargets) {
        /**
         * 親のチェックボックスに合わせて子の項目を有効／無効にする
         * @returns {void}
         */
        function updateEnabledStates() {
            /**
             * 作ってある項目だけ有効／無効を切り替える（簡易モードでは作らない項目がある）
             * @param {string} controlName - dialogControls のキー
             * @param {boolean} isEnabled - 有効にするか
             * @returns {void}
             */
            function setEnabledIfExists(controlName, isEnabled) {
                if (dialogControls[controlName]) dialogControls[controlName].enabled = isEnabled;
            }
            setEnabledIfExists("rotateSkipLongLatin", dialogControls.rotateRoman.value);
            setSteppedFieldEnabled(dialogControls.rotateSkipMinInput, dialogControls.rotateRoman.value && dialogControls.rotateSkipLongLatin.value);
            if (dialogControls.shortWordMaxInput) setSteppedFieldEnabled(dialogControls.shortWordMaxInput, dialogControls.shortWordFullwidth.value);
            setEnabledIfExists("tcyDigits2", dialogControls.tcyNumbers.value);
            setEnabledIfExists("tcyDigits3", dialogControls.tcyNumbers.value);
            setEnabledIfExists("tcyDigits4", dialogControls.tcyNumbers.value);
            setEnabledIfExists("tcyExcludeAdjacent", dialogControls.tcyNumbers.value);
            /* 漢数字の項目は詳細モードだけ / Kanji options exist only in the detailed mode */
            if (dialogControls.kanji) {
                setSteppedFieldEnabled(dialogControls.kanjiMinInput, dialogControls.kanji.value);
                setSteppedFieldEnabled(dialogControls.kanjiMaxInput, dialogControls.kanji.value);
                setEnabledIfExists("kanjiDigitRadio", dialogControls.kanji.value);
                setEnabledIfExists("kanjiPositionalRadio", dialogControls.kanji.value);
                setEnabledIfExists("kanjiLargeUnitsRadio", dialogControls.kanji.value);
                setEnabledIfExists("kanjiZeroRei", dialogControls.kanji.value);
                /* 一千は十・百・千を入れるときだけ / 一千 applies only to the positional style */
                setEnabledIfExists("kanjiIchiSen", dialogControls.kanji.value && dialogControls.kanjiPositionalRadio.value);
            }
        }

        /**
         * 確認の一覧を今の設定で作り直す
         * @returns {void}
         */
        function updatePreview() {
            var formatOptions = readEffectiveOptions(dialogControls);
            var previewLines = buildPreviewLines(textTargets, formatOptions);
            dialogControls.previewList.removeAll();
            if (previewLines.length === 0) previewLines.push(getLabel("preview.none"));
            for (var i = 0; i < previewLines.length; i++) dialogControls.previewList.add("item", previewLines[i]);
        }

        var clickControlNames = CHECKBOX_SETTING_KEYS.concat(["kanjiDigitRadio", "kanjiPositionalRadio", "kanjiLargeUnitsRadio"]);
        for (var i = 0; i < clickControlNames.length; i++) {
            if (!dialogControls[clickControlNames[i]]) continue;
            dialogControls[clickControlNames[i]].onClick = function () {
                updateEnabledStates();
                updatePreview();
            };
        }

        /**
         * 数値欄が変わったら、連動する欄をそろえてから確認の一覧を更新する
         * @param {EditText} [changedInput] - 変わった入力欄
         * @returns {void}
         */
        function onDigitInputChanged(changedInput) {
            syncWordLengthFields(dialogControls, changedInput);
            updatePreview();
        }

        /* 確認の一覧に効くドロップダウン / Dropdowns that change the check list */
        dialogControls.commaDropdown.onChange = updatePreview;
        if (dialogControls.phoneHyphenDropdown) dialogControls.phoneHyphenDropdown.onChange = updatePreview;

        digitStepHandler = onDigitInputChanged;
        wrapDigitInputUpdate(dialogControls.rotateSkipMinInput, onDigitInputChanged);
        if (dialogControls.shortWordMaxInput) wrapDigitInputUpdate(dialogControls.shortWordMaxInput, onDigitInputChanged);
        if (dialogControls.kanjiMinInput) {
            wrapDigitInputUpdate(dialogControls.kanjiMinInput, updatePreview);
            wrapDigitInputUpdate(dialogControls.kanjiMaxInput, updatePreview);
        }

        bindModeSwitch(dialogControls, dialogControls.modeSimpleRadio, "simple");
        bindModeSwitch(dialogControls, dialogControls.modeDetailRadio, "detail");
        bindPresetEvents(dialogControls, function () {
            updateEnabledStates();
            updatePreview();
        });
        updateEnabledStates();
        updatePreview();
    }

    /**
     * モードのラジオボタンを押したら、今の値のまま開き直す（作り直しは main で行う）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {RadioButton} modeRadio - モードのラジオボタン
     * @param {string} dialogMode - "simple" または "detail"
     * @returns {void}
     */
    function bindModeSwitch(dialogControls, modeRadio, dialogMode) {
        modeRadio.onClick = function () {
            if (dialogControls.dialogMode === dialogMode) return;
            dialogControls.dialogMode = dialogMode;
            dialogControls.switchCenter = getWindowCenter(dialogControls.dialog);
            dialogControls.dialog.close(MODE_SWITCH_RESULT);
        };
    }

    /**
     * プリセットの行のイベントをつなぐ（選ぶと読み込み、保存・削除でドロップダウンを作り直す）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {function(): void} onPresetApplied - プリセットを読み込んだあとに呼ぶ
     * @returns {void}
     */
    function bindPresetEvents(dialogControls, onPresetApplied) {
        /* プリセットの並び：--- / 組み込み / 区切り / 保存したもの / Preset list: placeholder, built-ins, separator, saved ones */
        var presetEntries = [];
        var isFillingPresets = false;

        /**
         * プリセットのドロップダウンを作り直す（作り直しでも onChange が来るので、その間は読まない）
         * @param {string|null} selectedName - 選んでおく保存したプリセット名（null なら「---」）
         * @returns {void}
         */
        function fillPresetDropdown(selectedName) {
            isFillingPresets = true;
            var presetDropdown = dialogControls.presetDropdown;
            presetDropdown.removeAll();
            presetEntries = [{ type: "none" }];
            presetDropdown.add("item", getLabel("preset.placeholder"));
            for (var b = 0; b < BUILTIN_PRESETS.length; b++) {
                presetEntries.push({ type: "builtin", data: BUILTIN_PRESETS[b].data });
                presetDropdown.add("item", getLabel("preset.builtin." + BUILTIN_PRESETS[b].labelKey));
            }
            var savedNames = getPresetNames();
            var selectedIndex = 0;
            if (savedNames.length > 0) {
                presetEntries.push({ type: "separator" });
                presetDropdown.add("separator");
                for (var n = 0; n < savedNames.length; n++) {
                    if (savedNames[n] === selectedName) selectedIndex = presetEntries.length;
                    presetEntries.push({ type: "saved", name: savedNames[n] });
                    presetDropdown.add("item", savedNames[n]);
                }
            }
            presetDropdown.selection = selectedIndex;
            setIconButtonEnabled(dialogControls.presetDeleteIcon, presetEntries[selectedIndex].type === "saved");
            isFillingPresets = false;
        }

        /**
         * 選んでいるプリセットの項目を返す
         * @returns {Object} presetEntries の要素
         */
        function getSelectedPresetEntry() {
            var presetSelection = dialogControls.presetDropdown.selection;
            return presetSelection ? presetEntries[presetSelection.index] : presetEntries[0];
        }

        dialogControls.presetDropdown.onChange = function () {
            if (isFillingPresets) return;
            var presetEntry = getSelectedPresetEntry();
            setIconButtonEnabled(dialogControls.presetDeleteIcon, presetEntry.type === "saved");
            var presetData = null;
            if (presetEntry.type === "builtin") presetData = presetEntry.data;
            else if (presetEntry.type === "saved") presetData = loadPresetMap()[presetEntry.name];
            if (!presetData) return;
            writeDialogSettings(dialogControls, mergePresetData(presetData));
            onPresetApplied();
        };
        dialogControls.presetSaveIcon.onClick = function () {
            var presetEntry = getSelectedPresetEntry();
            var savedName = saveCurrentPreset(readDialogOptions(dialogControls), presetEntry.type === "saved" ? presetEntry.name : "");
            if (savedName) fillPresetDropdown(savedName);
        };
        dialogControls.presetDeleteIcon.onClick = function () {
            var presetEntry = getSelectedPresetEntry();
            if (presetEntry.type !== "saved") return;
            if (deletePreset(presetEntry.name)) fillPresetDropdown(null);
        };

        fillPresetDropdown(null);
    }

    /**
     * ウィンドウの中心の座標を返す
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {number[]} [x, y]
     */
    function getWindowCenter(targetWindow) {
        var windowBounds = targetWindow.bounds;
        return [(windowBounds.left + windowBounds.right) / 2, (windowBounds.top + windowBounds.bottom) / 2];
    }

    /**
     * 表示したときに、ウィンドウの中心を指定の座標に合わせる（位置の復元のあとに動かす）
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number[]} centerPoint - 中心の座標 [x, y]
     * @returns {void}
     */
    function keepWindowCenterOnShow(targetWindow, centerPoint) {
        var previousOnShow = targetWindow.onShow;
        targetWindow.onShow = function () {
            if (previousOnShow) previousOnShow.apply(this, arguments);
            var windowBounds = targetWindow.bounds;
            var windowWidth = windowBounds.right - windowBounds.left;
            var windowHeight = windowBounds.bottom - windowBounds.top;
            targetWindow.location = [Math.round(centerPoint[0] - windowWidth / 2), Math.round(centerPoint[1] - windowHeight / 2)];
        };
    }

    /**
     * 数値欄を直接入力したときも確認の一覧を更新する
     * （部品の onChange は値をそろえるだけなので、包んで後ろに足す）
     * @param {EditText} numberInput - addSteppedField() の入力欄
     * @param {function(EditText): void} onUpdate - 更新の処理（引数は入力欄）
     * @returns {void}
     */
    function wrapDigitInputUpdate(numberInput, onUpdate) {
        var normalizeOnChange = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeOnChange();
            onUpdate(numberInput);
        };
    }

    /**
     * ［長い欧文は除く］と［短い英単語を全角に］の文字数をそろえる（全角は「除く」の1つ下）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {EditText} [changedInput] - 変わった入力欄
     * @returns {void}
     */
    function syncWordLengthFields(dialogControls, changedInput) {
        var shortWordInput = dialogControls.shortWordMaxInput;
        if (!shortWordInput) return;
        var rotateSkipInput = dialogControls.rotateSkipMinInput;
        if (changedInput === shortWordInput) {
            rotateSkipInput.text = String(Math.min(99, readDigitInput(shortWordInput, DEFAULT_SETTINGS.shortWordMaxChars) + 1));
        } else if (changedInput === rotateSkipInput) {
            shortWordInput.text = String(Math.max(1, readDigitInput(rotateSkipInput, DEFAULT_SETTINGS.rotateSkipMinChars) - 1));
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、ダイアログで設定を選んで実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        var doc = app.activeDocument;
        var textTargets = collectTextTargets(doc);
        if (textTargets.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
        var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
        var dialogControls;
        var dialogResult;
        var switchCenter = null;
        /* モードの切り替えは、値を引き継いで同じ中心の位置に作り直す / A mode switch rebuilds the dialog at the same center, keeping the values */
        do {
            dialogControls = buildDialog(dialogSettings);
            bindDialogEvents(dialogControls, textTargets);
            prepareDialogWindow(dialogControls.dialog, SCRIPT_NAME);
            if (switchCenter) keepWindowCenterOnShow(dialogControls.dialog, switchCenter);
            dialogResult = dialogControls.dialog.show();
            if (dialogResult === MODE_SWITCH_RESULT) {
                dialogSettings = readDialogOptions(dialogControls);
                switchCenter = dialogControls.switchCenter;
            }
        } while (dialogResult === MODE_SWITCH_RESULT);
        if (dialogResult !== 1) return;

        settingsStore.save(readDialogOptions(dialogControls));
        var formatOptions = readEffectiveOptions(dialogControls);
        var failureReason = runVerticalTextConversion(doc, textTargets, formatOptions);
        if (failureReason !== "") {
            alert(getLabel("alert.actionFailed") + "\n\n" + failureReason);
            return;
        }
        /* 確認用の色分けは、処理を終えて選択を戻したあとに、選択から集め直して塗る / Check coloring runs after processing, on ranges re-collected from the restored selection */
        if (formatOptions.colorCheck) colorCheckSelection(doc);
    }

    main();

})();
