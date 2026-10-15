#target illustrator
#targetengine "ShimbunTitleMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した見出しのまわりに罫線と段組みの新聞ダミーを敷いて、新聞の紙面を切り取ったような画像を作ります。
新聞ダミーは縦組み・横組み、和文・欧文（英字新聞モード）で作れ、全体を傾けて欠けない長方形でマスクします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShimbunTitleMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ndb9bee6b7a2e

### Overview

Surrounds the selected headline with rules and newspaper-style dummy columns, like a clipping from a newspaper page.
The dummy can be vertical or horizontal, Japanese or Western (English newspaper mode), and the piece is tilted and masked so no corner is missing.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShimbunTitleMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ShimbunTitleMaker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-12";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShimbunTitleMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShimbunTitleMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ndb9bee6b7a2e"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var DEFAULT_VERTICAL_MARGIN_PT = 5;    /* 上下マージンの既定値（pt）/ Default vertical margin */
    var DEFAULT_HORIZONTAL_MARGIN_PT = 8;  /* 左右マージンの既定値（pt）/ Default horizontal margin */
    var DEFAULT_LINK_MARGINS = false;      /* 上下・左右マージンの連動の既定値 / Default for linking the margins */
    var DEFAULT_RULE_STYLE = "double";   /* 罫線の種類（"double" 二重 / "single" 1本 / "none" なし）/ Rule style */
    var DEFAULT_LINE_GAP_PT = 5;         /* 二重線の間隔の既定値（pt）/ Default gap between the two rules */
    var DEFAULT_RULE_STROKE_WIDTH_PT = 0.6; /* 罫線の線幅の既定値（pt）/ Default stroke width of the rules */
    var DEFAULT_ROTATION_DEG = 3;        /* 回転の既定値（°）/ Default rotation angle */
    var DEFAULT_FIT_VIEW = true;         /* ［画面にフィット］の既定値 / Default for Fit to Window */
    var DEFAULT_MASK_VERTICAL_PT = 28.35; /* マスクの［上下］：罫線の上端／下端から見せる長さ（pt、10mm）/ Mask reach beyond the rules */
    var DEFAULT_USE_MASK = true;         /* マスクの既定値 / Default for the mask option */
    var DEFAULT_CHARS_PER_LINE = 11;     /* 新聞ダミーの1行あたりの文字数 / Dummy characters per line */
    var HORIZONTAL_HEADLINE_TIER_COUNT = 1; /* 見出しが和文の横組みのときの段数の初期値 / Initial tiers for a horizontal Japanese headline */
    var DEFAULT_TIER_COUNT = 3;          /* 新聞ダミーの段数（0で作らない）/ Dummy tiers (0 = none) */
    var DEFAULT_TIER_GAP_PT = 8.5;       /* 段間の既定値（pt、3mm）/ Default gutter between tiers */
    var DEFAULT_EXTRA_TIER_COUNT = 1;    /* 新聞ダミーの上下に足す段数 / Dummy tiers added above and below */
    var DEFAULT_DUMMY_WIDTH_PT = 85.04;  /* 新聞ダミーの左右それぞれの幅（pt、30mm）/ Dummy width on each side */
    var DUMMY_LEADING_RATIO = 1.4;       /* 新聞ダミーの行送り（文字サイズ比）/ Dummy leading as a ratio of type size */
    var DUMMY_VERTICAL_SCALE = 85;       /* 新聞ダミーの垂直比率（%）/ Dummy vertical scale */
    /* 新聞ダミーのフォント（PostScript 名、前から順に探す）/ Dummy font PostScript names, tried in order */
    var DUMMY_FONT_NAMES = ["HiraMinProN-W3", "HiraMinPro-W3"];
    var DEFAULT_USE_TIER_RULES = true;   /* 段間の罫線の既定値 / Default for the rules between tiers */
    var DEFAULT_TIER_RULE_WIDTH_PT = 0.3; /* 段間の罫線の線幅の既定値（pt）/ Default stroke width of the tier rules */
    var DUMMY_BACKGROUND_CMYK = [0, 0, 10, 25]; /* 新聞ダミーの背景色 [C,M,Y,K]（%）/ Dummy background color */
    var RULE_CMYK = [0, 0, 0, 100];      /* 二重線・区切り罫の色 [C,M,Y,K]（%）/ Color of the rules */
    var DEFAULT_USE_BLUR = true;         /* ぼかしの既定値 / Default for the blur option */
    var DEFAULT_BLUR_RADIUS = 8;         /* 新聞ダミーのぼかし（ガウス）の半径（px、0でかけない）/ Dummy blur radius in px (0 = none) */
    var EFFECT_RESOLUTION = 300;         /* 効果のプレビュー解像度（ppi）/ effect preview resolution */
    /* 新聞ダミーの文字列。句点は禁則を「なし」にして行頭にも置く / Dummy text; kinsoku is turned off so periods may start a line */
    var DUMMY_TEXT = "吾輩は猫である。名前はまだ無い。どこで生れたかとんと見当がつかぬ。何でも薄暗い所で泣いていた事だけは記憶している。";
    var DEFAULT_DUMMY_SCRIPT = "japanese"; /* 新聞ダミーの文章（"japanese" 和文 / "western" 欧文）。見出しがすべて欧文なら欧文で始める / Dummy script; all-Western headlines start as Western */
    var DEFAULT_DUMMY_DIRECTION = "vertical"; /* 新聞ダミーの組み方向（"vertical" 縦組み / "horizontal" 横組み）。欧文で始めるときは横組み / Dummy direction; horizontal when starting as Western */
    var DEFAULT_COLUMN_COUNT = 3;        /* 欧文：見出しの横幅に入る段組数 / Western: columns spanned by the headline */
    var WESTERN_CHARS_PER_LINE = 32;     /* 欧文の横組み：1行あたりのおよその文字数 / Western horizontal: approximate characters per line */
    var WESTERN_TEXT_RATIO = 2.5;        /* 欧文の縦組み：字詰に対して流す文字数の倍率（字幅が狭いぶん）/ Western vertical: extra characters for the narrower glyphs */
    var WESTERN_CHAR_WIDTH_RATIO = 0.5;  /* 欧文：1文字の平均幅（文字サイズ比）/ Western: average character width as a ratio of type size */
    var WESTERN_LEADING_RATIO = 1.2;     /* 欧文：行送り（文字サイズ比）/ Western: leading as a ratio of type size */
    /* 欧文の新聞ダミーのフォント（PostScript 名、前から順に探す）/ Western dummy font PostScript names, tried in order */
    var WESTERN_FONT_NAMES = ["TimesNewRomanPSMT", "Times-Roman", "Georgia"];
    /* 欧文の新聞ダミーの文字列 / Western dummy text */
    var WESTERN_DUMMY_TEXT = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

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

    var DUMMY_TEXT_INPUT_SIZE = [320, 72]; /* 文言の入力欄の大きさ / size of the dummy text fields */
    var RADIO_COLUMN_SPACING = 6;         /* 縦に積んだラジオボタンの間隔 / spacing of stacked radio buttons */
    var FIELD_CHARS = 5;                  /* 数値入力欄の幅（文字数）/ width of the numeric fields */
    var UNIT_FIELD_CHARS = 7;             /* 単位を中に入れる長さの欄の幅（文字数）/ width of the length fields that hold their unit */
    var FIELD_ROW_SPACING = 6;            /* 入力行の間隔 / spacing inside a field row */

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

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    var LINK_UI_DARK = isDarkUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @param {number[]} [iconSize] - アイコンの [幅, 高さ]（省略時は LINK_ICON_SIZE。絵は 22px 基準から拡大縮小する）
     * @param {number} [chainRatio] - 鎖の絵の大きさの比率（省略時は 1。枠・地の大きさは変えず、鎖だけ縮める）
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle, iconSize, chainRatio) {
        var toggleSize = iconSize || LINK_ICON_SIZE;
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = toggleSize;
        linkToggle.minimumSize = toggleSize;
        linkToggle.maximumSize = toggleSize;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = toggleSize[0];
            var iconHeight = toggleSize[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR, chainRatio);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @param {number} [chainRatio] - 鎖の大きさの比率（省略時は 1）。中央に置いたまま縮める
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

    // 画面にフィット（再利用パーツ、_templates/FitViewToItems.jsx） / Fit view to items (reusable)

    var FitViewToItems = (function () {

        // =========================================
        // ユーザー設定 / User Settings
        // =========================================

        /* ウィンドウに対して対象が占める割合の既定値（%）と範囲。呼び出し側で上書きできる
           Default share of the window the items fill, in percent, and its range; callers can override the default */
        var DEFAULT_FIT_PERCENT = 65;
        var FIT_PERCENT_RANGE = [10, 100];

        /* Illustratorが受け付ける表示倍率の範囲（3.125%〜6400%） / Zoom range Illustrator accepts */
        var VIEW_ZOOM_RANGE = [0.03125, 64];

        // =========================================
        // ローカライズ / Localization
        // =========================================
        var LABELS = {
            checkbox: {
                fitView: { ja: "画面にフィット", en: "Fit to Window" }
            },
            tooltip: {
                fitView: {
                    ja: "作成するオブジェクトが収まるよう表示倍率を合わせます。",
                    en: "Refits the view to the objects being created."
                },
                fitViewPercent: {
                    ja: "ウィンドウに対するオブジェクトの大きさ（100%でいっぱい）",
                    en: "Size of the objects relative to the window; 100% fills it"
                }
            }
        };

        /**
         * UI言語を返す
         * @returns {string} "ja" または "en"
         */
        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }

        /**
         * LABELS の組から指定言語の文言を返す
         * @param {object} labelSet - { ja: string, en: string } の組
         * @param {string} uiLang - "ja" または "en"
         * @returns {string} 文言（無ければ英語）
         */
        function getLabel(labelSet, uiLang) {
            return (labelSet[uiLang] != null) ? labelSet[uiLang] : labelSet.en;
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        /**
         * 数値を範囲に収める。数値として読めないときは既定値を返す
         * @param {string|number} value - 入力値
         * @param {number[]} range - [下限, 上限]
         * @param {number} fallbackValue - 読めないときの既定値
         * @returns {number} 範囲内の数値
         */
        function clampNumber(value, range, fallbackValue) {
            var numberValue = Number(value);
            if (isNaN(numberValue) || (typeof value === "string" && !/\S/.test(value))) numberValue = fallbackValue;
            return Math.min(range[1], Math.max(range[0], numberValue));
        }

        /**
         * 「□画面にフィット［65］%」の行を追加する（ラベルとツールチップは内蔵）
         * @param {Group|Panel|Window} parentContainer - 追加先のコンテナ
         * @param {object} [rowOptions] - value: チェックの初期値（既定 false）／percent: 割合の初期値／lang: 表示言語／
         *     stepper: true で入力欄の左に∧∨を付け、↑↓キーでも増減（StepperButtons 部品の addStepper / bindSteppedArrowKeys /
         *     redrawSteppersIn を使う。増減したら percentInput.onChange を呼ぶ）
         * @returns {{row: Group, checkbox: Checkbox, percentInput: EditText, getFillRatio: function, updateEnabled: function}} 作成したコントロール一式
         */
        function addControls(parentContainer, rowOptions) {
            if (!rowOptions) rowOptions = {};
            var uiLang = rowOptions.lang || getCurrentLang();

            var fitViewRow = parentContainer.add("group");
            fitViewRow.orientation = "row";
            fitViewRow.alignChildren = ["left", "center"];
            fitViewRow.spacing = 6;

            var fitViewCheck = fitViewRow.add("checkbox", undefined, getLabel(LABELS.checkbox.fitView, uiLang));
            fitViewCheck.helpTip = getLabel(LABELS.tooltip.fitView, uiLang);
            /* 明示的に true を渡したときだけONで始める / only an explicit true starts it checked */
            fitViewCheck.value = (rowOptions.value === true);

            var fitPercent = (rowOptions.percent > 0) ? rowOptions.percent : DEFAULT_FIT_PERCENT;
            var percentInput;
            var percentParent = fitViewRow;
            var stepperGroup = null;
            if (rowOptions.stepper === true) {
                /* ∧∨と入力欄は隙間0で突き合わせる / Butt the stepper against the field */
                percentParent = fitViewRow.add("group");
                percentParent.orientation = "row";
                percentParent.alignChildren = ["left", "center"];
                percentParent.spacing = 0;
                percentParent.margins = 0;
                stepperGroup = addStepper(percentParent, function () { return percentInput; }, {
                    integer: true,
                    min: FIT_PERCENT_RANGE[0],
                    max: FIT_PERCENT_RANGE[1],
                    /* text の代入ではイベントが来ないので、増減後に onChange を呼ぶ / assigning text fires no event */
                    onStep: function () { if (percentInput.onChange) percentInput.onChange(); }
                });
            }
            percentInput = percentParent.add("edittext", undefined, String(fitPercent));
            percentInput.characters = 3;
            percentInput.helpTip = getLabel(LABELS.tooltip.fitViewPercent, uiLang);
            if (stepperGroup) bindSteppedArrowKeys(percentInput, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */
            var percentUnitLabel = fitViewRow.add("statictext", undefined, "%");

            var controls = {
                row: fitViewRow,
                checkbox: fitViewCheck,
                percentInput: percentInput,

                /**
                 * 入力欄の割合を 0〜1 の比率で返す
                 * @returns {number} ウィンドウに対して占める割合（1でいっぱい）
                 */
                getFillRatio: function () {
                    return clampNumber(percentInput.text, FIT_PERCENT_RANGE, fitPercent) / 100;
                },

                /**
                 * 割合の入力欄をチェックの状態に合わせて有効・無効にする
                 * @returns {void}
                 */
                updateEnabled: function () {
                    percentInput.enabled = fitViewCheck.value;
                    percentUnitLabel.enabled = fitViewCheck.value;
                    if (!stepperGroup) return;
                    stepperGroup.enabled = fitViewCheck.value;
                    /* ∧∨は自作描画なので、表示中なら描き直してディム表示を切り替える / redraw the custom-drawn stepper once shown */
                    if (fitViewRow.window.visible) redrawSteppersIn(stepperGroup);
                }
            };
            controls.updateEnabled();
            return controls;
        }

        /**
         * 複数アイテムを囲む外接範囲を求める（効果を含まない geometricBounds）
         * @param {PageItem[]} targetItems - 対象アイテム
         * @returns {number[]|null} [left, top, right, bottom]（求められない場合は null）
         */
        function getItemsBounds(targetItems) {
            var unionBounds = null;
            for (var i = 0; i < targetItems.length; i++) {
                var itemBounds = targetItems[i].geometricBounds;
                if (unionBounds === null) {
                    unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
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
         * 対象が指定の割合でウィンドウに収まるよう、中心を合わせて表示倍率を変える
         * 拡大・縮小のどちらも行う（KeepInView と違い、常に同じ大きさに見せる）
         * @param {PageItem[]} targetItems - 対象アイテム
         * @param {object} [fitOptions] - doc: 対象ドキュメント（省略時は最前面）／fillRatio: 占める割合（1でいっぱい）
         * @returns {boolean} 表示を動かしたら true
         */
        function fit(targetItems, fitOptions) {
            if (!targetItems || targetItems.length === 0) return false;
            if (!fitOptions) fitOptions = {};

            var targetDoc = fitOptions.doc || app.activeDocument;
            var bounds = getItemsBounds(targetItems);
            if (bounds === null) return false;

            var itemWidth = bounds[2] - bounds[0];
            var itemHeight = bounds[1] - bounds[3];
            var activeView = targetDoc.activeView;
            activeView.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
            if (itemWidth <= 0 || itemHeight <= 0) return true;

            /* 中心をそろえたあとの表示範囲を基準に倍率を求める / Scale from the view bounds after the center has moved */
            var fillRatio = (fitOptions.fillRatio > 0) ? fitOptions.fillRatio : DEFAULT_FIT_PERCENT / 100;
            var viewBounds = activeView.bounds;
            var scale = Math.min(
                (viewBounds[2] - viewBounds[0]) / itemWidth,
                (viewBounds[1] - viewBounds[3]) / itemHeight
            ) * fillRatio;
            activeView.zoom = clampNumber(activeView.zoom * scale, VIEW_ZOOM_RANGE, 1);
            return true;
        }

        /**
         * 現在の表示位置と倍率を控える（キャンセル時に restoreView で戻す）
         * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
         * @returns {{centerPoint: number[], zoom: number}} 控えた表示状態
         */
        function captureView(targetDoc) {
            var activeView = (targetDoc || app.activeDocument).activeView;
            return { centerPoint: activeView.centerPoint, zoom: activeView.zoom };
        }

        /**
         * captureView で控えた表示位置と倍率に戻す
         * @param {{centerPoint: number[], zoom: number}} viewState - 控えた表示状態
         * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
         * @returns {void}
         */
        function restoreView(viewState, targetDoc) {
            if (!viewState) return;
            var activeView = (targetDoc || app.activeDocument).activeView;
            activeView.centerPoint = viewState.centerPoint;
            activeView.zoom = viewState.zoom;
        }

        return {
            addControls: addControls,
            fit: fit,
            captureView: captureView,
            restoreView: restoreView,
            getItemsBounds: getItemsBounds
        };

    })();

    // 画面にフィット（再利用パーツ）ここまで / End of the reusable fit view

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "新聞見出しメーカー", en: "Newspaper Headline Maker" },
            dummyTextTitle: { ja: "ダミー文章", en: "Dummy Text" }
        },
        panel: {
            margin: { ja: "マージン", en: "Margins" },
            sideRules: { ja: "罫線", en: "Rules" },
            typesetting: { ja: "組版設定", en: "Typesetting" },
            dummyText: { ja: "ダミー文章", en: "Dummy Text" },
            maskRotation: { ja: "マスクと回転", en: "Mask and Rotation" }
        },
        fieldLabel: {
            marginVertical: { ja: "上下", en: "Vertical" },
            marginHorizontal: { ja: "左右", en: "Horizontal" },
            ruleStrokeWidth: { ja: "線幅", en: "Weight" },
            tierRuleWidth: { ja: "線幅", en: "Weight" },
            lineGap: { ja: "線の間隔", en: "Gap" },
            dummyDirection: { ja: "組み方向", en: "Direction" },
            charsPerLine: { ja: "字詰", en: "Chars per line" },
            tierCount: { ja: "段数", en: "Tiers" },
            tierGap: { ja: "段間", en: "Gutter" },
            extraTierCount: { ja: "段数追加", en: "Extra tiers" },
            dummyWidth: { ja: "左右", en: "Horizontal" },
            maskVertical: { ja: "上下", en: "Vertical" },
            blurRadius: { ja: "ぼかし", en: "Blur" },
            rotation: { ja: "回転", en: "Rotation" }
        },
        radio: {
            ruleDouble: { ja: "二重", en: "Double" },
            ruleSingle: { ja: "1本（上下）", en: "Single (both)" },
            ruleSingleTop: { ja: "1本（上のみ）", en: "Single (top only)" },
            ruleNone: { ja: "なし", en: "None" },
            vertical: { ja: "縦組み", en: "Vertical" },
            horizontal: { ja: "横組み", en: "Horizontal" },
            japanese: { ja: "和文", en: "Japanese" },
            western: { ja: "欧文", en: "Western" }
        },
        checkbox: {
            useMask: { ja: "マスク", en: "Mask" },
            useBlur: { ja: "ぼかし", en: "Blur" },
            useTierRules: { ja: "段間の罫線", en: "Tier rules" },
            englishNewspaper: { ja: "英字新聞モード", en: "English newspaper mode" }
        },
        unit: {
            chars: { ja: "文字", en: "chars" },
            tiers: { ja: "段", en: "tiers" },
            columns: { ja: "段", en: "columns" }
        },
        tooltip: {
            marginVertical: {
                ja: "見出しの上端／下端から、罫線と新聞ダミーの天地を伸ばす量。横組みの欧文では見出しと横棒の距離",
                en: "How far the rules and the dummy height extend beyond the top and bottom of the headline. Horizontal Western: distance between the headline and the rules"
            },
            marginHorizontal: {
                ja: "見出しと内側の線との距離。新聞ダミーの行をそろえるため、少し広がることがある。横組みの欧文では横棒を見出しの左右に伸ばす量",
                en: "Distance between the headline and the inner rule. It may widen slightly so the dummy lines align. Horizontal Western: how far the rules extend past the headline"
            },
            linkMargins: { ja: "上下と左右を同じ値にする", en: "Keeps the vertical and horizontal margins equal" },
            ruleStrokeWidth: { ja: "左右の罫線の線幅。［なし］では無効", en: "Stroke weight of the side rules. Disabled with None" },
            tierRuleWidth: {
                ja: "段間の罫線の線幅。［段間の罫線］がOFFのときは無効",
                en: "Stroke weight of the tier rules. Disabled while Tier rules is off"
            },
            lineGap: { ja: "内側の線と外側の線の間隔。［二重］のときだけ有効", en: "Distance between the inner and outer rule. Double only" },
            ruleDouble: {
                ja: "見出しの左右に二重線を引く。横組みの欧文では上下に横棒で引く",
                en: "Draws double rules on both sides of the headline. Horizontal Western draws them above and below"
            },
            ruleSingle: {
                ja: "見出しの左右に1本ずつ線を引く。横組みの欧文では上下に横棒で引く",
                en: "Draws a single rule on each side of the headline. Horizontal Western draws them above and below"
            },
            ruleSingleTop: {
                ja: "見出しの上にだけ横棒を1本引く。横組みの欧文のときだけ選べる",
                en: "Draws one horizontal rule above the headline only. Available for horizontal Western only"
            },
            ruleNone: {
                ja: "罫線を引かない。新聞ダミーの位置は［マージン］で決まる",
                en: "Draws no rules. The dummy text is placed by Margins"
            },
            charsPerLine: {
                ja: "1段の1行に入れる文字数。欧文の横組みでは無効",
                en: "Characters in one line of a tier or column. Disabled for Western horizontal"
            },
            tierCount: {
                ja: "縦組み：罫線の天地を分ける段数。0で新聞ダミーを作らない",
                en: "Vertical: number of tiers the rule height is split into. 0 adds no dummy text"
            },
            tierGap: {
                ja: "段と段の間隔。段間の罫線はこの中央に引く",
                en: "Space between tiers. Tier rules run through its middle"
            },
            extraTierCount: {
                ja: "上と下にそれぞれ足す、左右の幅いっぱいの段の数。0で足さない",
                en: "Full-width tiers added above and below each. 0 adds none"
            },
            dummyWidth: {
                ja: "左右それぞれの新聞ダミーの幅。マスクがOFFでも新聞ダミーの幅として使う",
                en: "Width of the dummy text on each side. Still used as the dummy width while Mask is off"
            },
            columnCount: {
                ja: "横組み：見出しの横幅に入る段の数。段の幅と文字サイズはこれから決まる。0で新聞ダミーを作らない",
                en: "Horizontal: how many columns the headline spans. Sets the column width and type size. 0 adds no dummy text"
            },
            dummyText: {
                ja: "新聞ダミーに流す文章を和文・欧文それぞれ設定する",
                en: "Sets the text poured into the dummy, for Japanese and Western"
            },
            dummyTextInput: {
                ja: "足りない分は繰り返す。空欄なら既定の文章。改行は詰める",
                en: "Repeated as needed. Empty uses the default text. Line breaks are removed"
            },
            vertical: { ja: "罫線の天地を段に分けた縦組みで新聞ダミーを作る", en: "Builds the dummy as vertical tiers that split the rule height" },
            horizontal: { ja: "見出しの横幅を段組みに分けた横組みで新聞ダミーを作る", en: "Builds the dummy as horizontal columns across the headline width" },
            japanese: { ja: "和文の文章とフォント（ヒラギノ明朝 W3）で新聞ダミーを作る", en: "Uses Japanese text and font (Hiragino Mincho W3) for the dummy" },
            western: { ja: "欧文の文章とフォント（Times New Roman）で新聞ダミーを作る", en: "Uses Western text and font (Times New Roman) for the dummy" },
            westernVertical: {
                ja: "欧文の文章とフォント（Times New Roman）で新聞ダミーを作る。縦組みの見出しでは選べない",
                en: "Uses Western text and font (Times New Roman) for the dummy. Not available for vertical headlines"
            },
            blurRadius: { ja: "新聞ダミーにかけるぼかし（ガウス）の半径（px）", en: "Radius of the Gaussian blur on the dummy text, in pixels" },
            useBlur: { ja: "新聞ダミーにぼかし（ガウス）をかける", en: "Applies a Gaussian blur to the dummy text" },
            maskVertical: {
                ja: "マスクの天地を、罫線の上端／下端からこの長さまでにする。回転後も角が欠けない大きさが上限",
                en: "Limits the mask to this distance beyond the top and bottom of the rules. Capped at the largest size that leaves no corner missing"
            },
            englishNewspaper: {
                ja: "横組み・欧文に固定し、罫線を［1本（上のみ）］にする。OFFで元の組み方向と和文／欧文に戻す",
                en: "Locks Horizontal and Western and sets the rules to Single (top only). Turning it off restores the previous direction and script"
            },
            englishNewspaperVertical: {
                ja: "横組み・欧文に固定し、罫線を［1本（上のみ）］にする。縦組みの見出しでは選べない",
                en: "Locks Horizontal and Western and sets the rules to Single (top only). Not available for vertical headlines"
            },
            useTierRules: {
                ja: "段と段の間に罫線を引く。横組みの和文は見出しと上下の段組みの間、横組みの欧文は段と段の間の縦棒",
                en: "Draws rules between the tiers. Horizontal Japanese: between the headline and the bands above and below. Horizontal Western: vertical rules between columns"
            },
            rotation: {
                ja: "全体のグループに［変形］効果で適用する回転角度。0で効果を付けない",
                en: "Rotation applied to the whole group with the Transform effect. 0 adds no effect"
            },
            useMask: {
                ja: "回転後も角が欠けない範囲で、天地を［上下］で切った長方形で全体をマスクする。［段数追加］を0にするとOFFになる",
                en: "Masks the whole piece with a rectangle that leaves no corner missing after the rotation, its height cut by Vertical. Turns off when Extra tiers is set to 0"
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
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" },
            dummyText: { ja: "文章...", en: "Text..." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTarget: { ja: "テキストオブジェクトまたはグループを1つ選択してください。", en: "Select one text object or one group." }
        },
        itemName: {
            ruleGroup: { ja: "左右の罫線", en: "Side rules" },
            dummyGroup: { ja: "新聞ダミー", en: "Newspaper dummy" }
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

    // =========================================
    // 入力の読み取り / Input handling
    // =========================================

    /**
     * 数値の入力文字列を読む。空白を除き、全角のカンマ・小数点を半角にし、
     * 「1,5」のような小数カンマは小数点に、「1,000」のような桁区切りは取り除く
     * @param {string} inputText - 入力された文字列
     * @returns {number} 読み取った数値（読めなければ NaN）
     */
    function parseLocaleNumber(inputText) {
        var normalizedText = String(inputText);
        normalizedText = normalizedText.replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        if (normalizedText.indexOf(",") >= 0 && normalizedText.indexOf(".") < 0) {
            normalizedText = normalizedText.replace(/,/g, ".");
        } else {
            normalizedText = normalizedText.replace(/,/g, "");
        }
        return Number(normalizedText);
    }

    /**
     * 入力欄の値を数値で読む（読めないときは既定値）
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} fallbackValue - 読めなかったときに使う値
     * @param {boolean} allowNegative - 負の値を許すか（許さないときは0に丸める）
     * @returns {number} 読み取った値
     */
    function readFieldAsNumber(inputField, fallbackValue, allowNegative) {
        var enteredValue = parseLocaleNumber(inputField.text);
        /* 単位付き（"5 mm"）や計算式はステップボタン部品の計算で読み、欄の単位に換算する / units and expressions go through the stepper's evaluator */
        if (isNaN(enteredValue)) enteredValue = evaluateArithmetic(inputField.text, inputField.fieldUnit);
        if (isNaN(enteredValue)) return fallbackValue;
        if (!allowNegative && enteredValue < 0) return 0;
        return enteredValue;
    }

    /**
     * 入力欄の値（現在の単位）を pt に換算して読む。負の値は0に丸める
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} fallbackPt - 読めなかったときに使う値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {number} 換算した値（pt）
     */
    function readFieldAsPoints(inputField, fallbackPt, pointsPerUnit) {
        var enteredValue = readFieldAsNumber(inputField, NaN, false);
        return isNaN(enteredValue) ? fallbackPt : enteredValue * pointsPerUnit;
    }

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

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確認し、ダイアログでプレビューしながら二重線を描画してグループ化する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var targetItem = getTargetItem(doc.selection);
        if (!targetItem) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        showDialog(doc, targetItem);
    }

    /**
     * 基準にできる選択かを判定して返す（テキスト1つ、またはグループ1つ）
     * @param {Object} currentSelection - ドキュメントの選択
     * @returns {PageItem|null} 基準にするオブジェクト（対象外なら null）
     */
    function getTargetItem(currentSelection) {
        if (!currentSelection || currentSelection.length !== 1) return null;

        var selectedItem = currentSelection[0];
        var isSupported = (
            selectedItem.typename === "TextFrame" ||
            selectedItem.typename === "GroupItem"
        );
        return isSupported ? selectedItem : null;
    }

    /**
     * 対象の文字がすべて欧文なら "western"、和文（かな・漢字・全角）を含むか文字が無ければ既定の組み方を返す。
     * 縦組みのテキストを含むときは欧文を選べないので "japanese"
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @returns {string} "japanese" または "western"
     */
    function detectDummyScript(targetItem) {
        if (hasVerticalText(targetItem)) return "japanese";
        var targetText = collectTextContents(targetItem);
        if (!/\S/.test(targetText)) return DEFAULT_DUMMY_SCRIPT;
        /* 全角スペース・句読点・かな・漢字・全角英数を和文と見なす / CJK punctuation, kana, kanji and full-width forms count as Japanese */
        return /[\u3000-\u30FF\u3400-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/.test(targetText) ? "japanese" : "western";
    }

    /**
     * 縦組みのテキストを含むかを返す（グループは入れ子もたどる）
     * @param {PageItem} pageItem - 対象
     * @returns {boolean} 縦組みのテキストがあれば true
     */
    function hasVerticalText(pageItem) {
        if (pageItem.typename === "TextFrame") return pageItem.orientation === TextOrientation.VERTICAL;
        if (pageItem.typename !== "GroupItem") return false;
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            if (hasVerticalText(pageItem.pageItems[i])) return true;
        }
        return false;
    }

    /**
     * テキストまたはグループの中のテキストの内容をすべてつないで返す（グループは入れ子もたどる）
     * @param {PageItem} pageItem - 対象
     * @returns {string} つないだ文字列
     */
    function collectTextContents(pageItem) {
        if (pageItem.typename === "TextFrame") return pageItem.contents;
        if (pageItem.typename !== "GroupItem") return "";
        var collectedText = "";
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            collectedText += collectTextContents(pageItem.pageItems[i]);
        }
        return collectedText;
    }

    /**
     * 罫線の枠（見出しをマージンぶん広げた範囲と、内側・外側の線の位置）を求める。［二重］以外は外側の線を内側の線に重ねる。
     * 横棒（見出しの上下に引く）のときは、外側の線のぶん topY / bottomY を広げ、新聞ダミーはその外に置く
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {{marginVerticalPt: number, marginHorizontalPt: number, lineGapPt: number, ruleStyle: string}} headlineSettings - マージンと間隔（pt）、罫線の種類
     * @param {boolean} usesHorizontalRules - 罫線を見出しの上下に横棒で引くなら true
     * @returns {{isHorizontal: boolean, topY: number, bottomY: number, innerTopY: number, innerBottomY: number, innerLeftX: number, innerRightX: number, outerLeftX: number, outerRightX: number}} 罫線の枠
     */
    function getRuleFrame(targetItem, headlineSettings, usesHorizontalRules) {
        /* visibleBounds: [左, 上, 右, 下] / visibleBounds: [left, top, right, bottom] */
        var targetBounds = targetItem.visibleBounds;
        var innerLeftX = targetBounds[0] - headlineSettings.marginHorizontalPt;
        var innerRightX = targetBounds[2] + headlineSettings.marginHorizontalPt;
        var innerTopY = targetBounds[1] + headlineSettings.marginVerticalPt;
        var innerBottomY = targetBounds[3] - headlineSettings.marginVerticalPt;
        var lineGapPt = (headlineSettings.ruleStyle === "double") ? headlineSettings.lineGapPt : 0;
        var sideGapPt = usesHorizontalRules ? 0 : lineGapPt;
        var endGapPt = usesHorizontalRules ? lineGapPt : 0;
        return {
            isHorizontal: usesHorizontalRules,
            topY: innerTopY + endGapPt,
            bottomY: innerBottomY - endGapPt,
            innerTopY: innerTopY,
            innerBottomY: innerBottomY,
            innerLeftX: innerLeftX,
            innerRightX: innerRightX,
            outerLeftX: innerLeftX - sideGapPt,
            outerRightX: innerRightX + sideGapPt
        };
    }

    /**
     * 対象の左右（横棒のときは上下）に罫線を描画する（［二重］は2本ずつ、［1本］は内側の線だけ、［1本（上のみ）］は上の線だけ）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {Object} ruleFrame - getRuleFrame() で求めた罫線の枠
     * @param {{ruleStyle: string, ruleStrokeWidthPt: number}} headlineSettings - 罫線の種類と線幅（pt）
     * @returns {GroupItem|null} 生成した罫線のグループ。［なし］のときは null
     */
    function drawSideRules(doc, targetItem, ruleFrame, headlineSettings) {
        if (headlineSettings.ruleStyle === "none") return null;
        var ruleGroup = targetItem.layer.groupItems.add();
        ruleGroup.name = getLabel("itemName.ruleGroup");

        var strokeColor = createProcessColor(doc, RULE_CMYK);
        var linePoints = [];
        if (ruleFrame.isHorizontal) {
            /* 見出しの上下に横棒。外側の線が内側に重なる（［1本］か間隔0）なら描かない / horizontal rules above and below; skip stacked outer lines */
            /* ［1本（上のみ）］は上の線だけ / Single (top) draws only the upper rule */
            var lineYPositions = (headlineSettings.ruleStyle === "singleTop") ? [ruleFrame.innerTopY] : [ruleFrame.innerTopY, ruleFrame.innerBottomY];
            if (ruleFrame.topY > ruleFrame.innerTopY) lineYPositions.push(ruleFrame.topY, ruleFrame.bottomY);
            for (var j = 0; j < lineYPositions.length; j++) {
                linePoints.push([[ruleFrame.innerLeftX, lineYPositions[j]], [ruleFrame.innerRightX, lineYPositions[j]]]);
            }
        } else {
            /* 見出しの左右に縦の線。外側の線が内側に重なるなら描かない / vertical rules left and right; skip stacked outer lines */
            var lineXPositions = [ruleFrame.innerLeftX, ruleFrame.innerRightX];
            if (ruleFrame.outerRightX > ruleFrame.innerRightX) lineXPositions.push(ruleFrame.outerLeftX, ruleFrame.outerRightX);
            for (var k = 0; k < lineXPositions.length; k++) {
                linePoints.push([[lineXPositions[k], ruleFrame.topY], [lineXPositions[k], ruleFrame.bottomY]]);
            }
        }
        for (var i = 0; i < linePoints.length; i++) {
            createRuleLine(ruleGroup, linePoints[i][0], linePoints[i][1], strokeColor, headlineSettings.ruleStrokeWidthPt);
        }

        /* 対象のすぐ背面へ / Move the rules just behind the target */
        ruleGroup.move(targetItem, ElementPlacement.PLACEAFTER);
        return ruleGroup;
    }

    /**
     * 縦組みの新聞ダミーの文字サイズ・行送り・左右の行数を求める。
     * 罫線の天地（A）を段数で分け、段の間は段間（pt）だけ空ける。
     * 1文字の送り（文字サイズ × 垂直比率）は（A −（段数 − 1）× 段間）÷（段数 × 字詰）。
     * 左右の幅は行の整数倍に切り詰める
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠
     * @param {{charsPerLine: number, tierCount: number, tierGap: number, dummyWidthPt: number}} headlineSettings - 新聞ダミーの設定（tierGap は pt）
     * @param {Object} typeStyle - getDummyTypeStyle() で求めた書体の設定
     * @returns {Object|null} 新聞ダミーの寸法（作らないときは null）
     */
    function computeTierLayout(ruleFrame, headlineSettings, typeStyle) {
        var tierCount = headlineSettings.tierCount;
        var charsPerLine = headlineSettings.charsPerLine;
        var ruleHeight = ruleFrame.topY - ruleFrame.bottomY;
        if (tierCount < 1 || charsPerLine < 1 || headlineSettings.dummyWidthPt <= 0 || ruleHeight <= 0) return null;

        /* 縦組みでは垂直比率のぶん字送りが詰まる / in vertical text the vertical scale shortens each character's advance */
        var tierGap = headlineSettings.tierGap;
        var charAdvance = (ruleHeight - (tierCount - 1) * tierGap) / (tierCount * charsPerLine);
        if (charAdvance <= 0) return null; /* 段間が天地を食い尽くしたら作らない / the gutters leave no room for text */
        var fontSize = charAdvance / (typeStyle.verticalScale / 100);
        var dummyLayout = {
            tierCount: tierCount,
            charsPerLine: charsPerLine,
            fontSize: fontSize,
            charAdvance: charAdvance,
            linePitch: fontSize * DUMMY_LEADING_RATIO,
            tierGap: tierGap,
            tierPitch: charsPerLine * charAdvance + tierGap
        };
        dummyLayout.sideLineCount = countFittingLines(dummyLayout, headlineSettings.dummyWidthPt);
        dummyLayout.sideTextWidth = getTextWidth(dummyLayout, dummyLayout.sideLineCount);
        return dummyLayout;
    }

    /**
     * 上下の段の行が左右の段の行と揃うように、二重線を左右へ均等に広げる。
     * 左の本文の右端から右の本文の右端までが行送りの整数倍になればよい
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠（書き換える）
     * @param {Object} dummyLayout - computeTierLayout() で求めた寸法
     * @returns {void}
     */
    function alignRuleFrameToDummyLines(ruleFrame, dummyLayout) {
        var linePitch = dummyLayout.linePitch;
        /* 外側の線と本文の間は1文字分 / one character of space between the outer rule and the text */
        var lineSpan = (ruleFrame.outerRightX - ruleFrame.outerLeftX) + 2 * dummyLayout.fontSize + dummyLayout.sideTextWidth;
        /* 浮動小数の誤差で1行ぶん余計に広げない / keep float error from adding a whole extra line */
        var alignedSpan = Math.ceil(lineSpan / linePitch - 1e-6) * linePitch;
        var shiftX = (alignedSpan - lineSpan) / 2;

        ruleFrame.innerLeftX -= shiftX;
        ruleFrame.outerLeftX -= shiftX;
        ruleFrame.innerRightX += shiftX;
        ruleFrame.outerRightX += shiftX;
    }

    /**
     * 罫線の外側の左右に縦組みの段の新聞ダミーを描画し、上下には左右の幅いっぱいの段を足す。
     * 段の間は段間だけ空けて、その中央に区切り罫を引く
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @param {Object} ruleFrame - 二重線の枠
     * @param {Object} dummyLayout - computeTierLayout() で求めた寸法
     * @param {number} extraTierCount - 上下それぞれに足す段数
     * @param {Object} typeStyle - getDummyTypeStyle() で求めた書体の設定
     * @param {number} tierRuleWidthPt - 段間の罫線の線幅（pt、0なら引かない）
     * @returns {GroupItem} 生成した新聞ダミーのグループ
     */
    function drawTierDummy(doc, targetLayer, ruleFrame, dummyLayout, extraTierCount, typeStyle, tierRuleWidthPt) {
        var fontSize = dummyLayout.fontSize;
        var charAdvance = dummyLayout.charAdvance;
        var tierPitch = dummyLayout.tierPitch;
        var tierGap = dummyLayout.tierGap;
        var sideLineCount = dummyLayout.sideLineCount;
        var sideTextWidth = dummyLayout.sideTextWidth;

        var dummyGroup = targetLayer.groupItems.add();
        dummyGroup.name = getLabel("itemName.dummyGroup");
        var tierContext = {
            targetLayer: targetLayer,
            dummyGroup: dummyGroup,
            charsPerLine: dummyLayout.charsPerLine,
            fontSize: fontSize,
            charAdvance: charAdvance,
            linePitch: dummyLayout.linePitch,
            textFont: findDummyFont(typeStyle.fontNames),
            typeStyle: typeStyle,
            tierRuleWidthPt: tierRuleWidthPt,
            strokeColor: createProcessColor(doc, RULE_CMYK)
        };

        /* 外側の線と本文の間は1文字分 / one character of space between the outer rule and the text */
        var sideTextLefts = [
            ruleFrame.outerLeftX - fontSize - sideTextWidth,
            ruleFrame.outerRightX + fontSize
        ];

        /* 左右の段。2段目以降は上の段との間に罫 / side tiers; a rule above every tier but the first */
        for (var i = 0; i < sideTextLefts.length; i++) {
            for (var k = 0; k < dummyLayout.tierCount; k++) {
                var sideTierTop = ruleFrame.topY - k * tierPitch;
                addDummyTier(tierContext, sideTextLefts[i], sideTierTop, sideLineCount);
                if (k > 0) addTierRule(tierContext, sideTextLefts[i], sideTierTop + tierGap / 2, sideTextWidth);
            }
        }

        /* 上下の段は左の本文の左端から右の本文の右端まで。alignRuleFrameToDummyLines() で行送りの整数倍になっている
           Tiers above and below span both side regions, already a whole number of lines wide */
        var fullLeft = sideTextLefts[0];
        var fullWidth = sideTextLefts[1] + sideTextWidth - fullLeft;
        var fullLineCount = countFittingLines(tierContext, fullWidth);

        /* どれも内側（ボックス側）の段との間に罫 / each has a rule on the side facing the box */
        for (var j = 1; j <= extraTierCount; j++) {
            var upperTop = ruleFrame.topY + j * tierPitch;
            addDummyTier(tierContext, fullLeft, upperTop, fullLineCount);
            addTierRule(tierContext, fullLeft, upperTop - tierPitch + tierGap / 2, fullWidth);

            var lowerTop = ruleFrame.bottomY - tierGap - (j - 1) * tierPitch;
            addDummyTier(tierContext, fullLeft, lowerTop, fullLineCount);
            addTierRule(tierContext, fullLeft, lowerTop + tierGap / 2, fullWidth);
        }
        return dummyGroup;
    }

    /**
     * 新聞ダミーのフォントを探す（見つからなければ null で既定のフォントのまま）
     * @param {string[]} fontNames - PostScript 名の候補（前から順に探す）
     * @returns {TextFont|null} フォント
     */
    function findDummyFont(fontNames) {
        for (var i = 0; i < fontNames.length; i++) {
            /* getByName は見つからないと例外 / getByName throws when the font is missing */
            try {
                return app.textFonts.getByName(fontNames[i]);
            } catch (e) {}
        }
        return null;
    }

    /**
     * 指定の幅に収まる行数を返す（最低1行）
     * @param {{fontSize: number, linePitch: number}} tierContext - 段の共通設定（または新聞ダミーの寸法）
     * @param {number} availableWidth - 使える幅（pt）
     * @returns {number} 行数
     */
    function countFittingLines(tierContext, availableWidth) {
        /* ちょうど割り切れる幅が誤差で1行減らないよう、わずかに足す / a tiny epsilon keeps exact fits from losing a line */
        var lineCount = Math.floor((availableWidth - tierContext.fontSize) / tierContext.linePitch + 1e-6) + 1;
        return Math.max(1, lineCount);
    }

    /**
     * 行数ぶんの本文の幅（最初の行の右端から最後の行の左端まで）を返す
     * @param {{fontSize: number, linePitch: number}} tierContext - 段の共通設定
     * @param {number} lineCount - 行数
     * @returns {number} 幅（pt）
     */
    function getTextWidth(tierContext, lineCount) {
        return (lineCount - 1) * tierContext.linePitch + tierContext.fontSize;
    }

    /**
     * 段の区切り罫を1本引く（線幅が0、つまり段間の罫線がOFFなら何もしない）
     * @param {{dummyGroup: GroupItem, strokeColor: Object, tierRuleWidthPt: number}} dummyContext - 段の共通設定
     * @param {number[]} startPoint - 始点 [x, y]
     * @param {number[]} endPoint - 終点 [x, y]
     * @returns {void}
     */
    function addDummyRule(dummyContext, startPoint, endPoint) {
        if (!(dummyContext.tierRuleWidthPt > 0)) return;
        createRuleLine(dummyContext.dummyGroup, startPoint, endPoint, dummyContext.strokeColor, dummyContext.tierRuleWidthPt);
    }

    /**
     * 段の間に横の区切り罫を1本引く
     * @param {Object} dummyContext - 段の共通設定
     * @param {number} ruleLeft - 左端の X 座標
     * @param {number} ruleY - Y 座標
     * @param {number} ruleWidth - 長さ（pt）
     * @returns {void}
     */
    function addTierRule(dummyContext, ruleLeft, ruleY, ruleWidth) {
        addDummyRule(dummyContext, [ruleLeft, ruleY], [ruleLeft + ruleWidth, ruleY]);
    }

    /**
     * 段と段の間に縦の区切り罫を1本引く
     * @param {Object} dummyContext - 段の共通設定
     * @param {number} ruleX - X 座標
     * @param {number} ruleTop - 上端の Y 座標
     * @param {number} ruleHeight - 長さ（pt）
     * @returns {void}
     */
    function addColumnRule(dummyContext, ruleX, ruleTop, ruleHeight) {
        addDummyRule(dummyContext, [ruleX, ruleTop], [ruleX, ruleTop - ruleHeight]);
    }

    /**
     * 新聞ダミー1段ぶんの縦組みエリア内文字を作り、新聞ダミーのグループに入れる。
     * 本文がちょうど指定の行数で終わるように、幅と文字数を合わせる
     * @param {{targetLayer: Layer, dummyGroup: GroupItem, charsPerLine: number, fontSize: number, charAdvance: number, linePitch: number, textFont: TextFont}} tierContext - 段の共通設定
     * @param {number} textLeft - 本文の左端の X 座標
     * @param {number} tierTop - 上端の Y 座標
     * @param {number} lineCount - 行数
     * @returns {TextFrame} 生成したテキストフレーム
     */
    function addDummyTier(tierContext, textLeft, tierTop, lineCount) {
        var charsPerLine = tierContext.charsPerLine;
        var charAdvance = tierContext.charAdvance;
        /* 丸めで行や文字がこぼれないよう、幅は左に、高さは下に半端な余裕を足す
           Add slack to the left and bottom so rounding never drops a line or a character */
        var widthSlack = (tierContext.linePitch - tierContext.fontSize) / 2;
        var frameWidth = getTextWidth(tierContext, lineCount) + widthSlack;
        var frameHeight = charsPerLine * charAdvance + charAdvance / 2;

        var targetLayer = tierContext.targetLayer;
        var framePath = targetLayer.pathItems.rectangle(tierTop, textLeft - widthSlack, frameWidth, frameHeight);
        var dummyFrame = targetLayer.textFrames.areaText(framePath, TextOrientation.VERTICAL);
        var typeStyle = tierContext.typeStyle;
        dummyFrame.contents = buildDummyText(typeStyle.sourceText, Math.ceil(lineCount * charsPerLine * typeStyle.textRatio), 0);

        /* 句点が行頭に来ても字詰どおりに流す / keep every line at the set length even when a period starts it */
        dummyFrame.textRange.paragraphAttributes.kinsoku = "None";

        var dummyAttributes = dummyFrame.textRange.characterAttributes;
        if (tierContext.textFont) dummyAttributes.textFont = tierContext.textFont;
        dummyAttributes.size = tierContext.fontSize;
        dummyAttributes.verticalScale = typeStyle.verticalScale;
        dummyAttributes.autoLeading = false;
        dummyAttributes.leading = tierContext.linePitch;

        dummyFrame.move(tierContext.dummyGroup, ElementPlacement.PLACEATEND);
        return dummyFrame;
    }

    /**
     * 和文／欧文に合わせた新聞ダミーの書体の設定を返す（組み方向とは別に決まる）
     * @param {boolean} isWestern - 欧文なら true
     * @param {string} sourceText - 流す文章（normalizeDummyText() で整えたもの）
     * @param {number} charsPerLine - 字詰（和文の横組みでは1行の文字数に使う）
     * @returns {{fontNames: string[], sourceText: string, textRatio: number, columnCharsPerLine: number, charWidthRatio: number, columnLeadingRatio: number, verticalScale: number, hyphenation: boolean, verticalColumnRules: boolean}} 書体の設定
     */
    function getDummyTypeStyle(isWestern, sourceText, charsPerLine) {
        if (isWestern) {
            return {
                fontNames: WESTERN_FONT_NAMES,
                sourceText: sourceText,
                textRatio: WESTERN_TEXT_RATIO,
                columnCharsPerLine: WESTERN_CHARS_PER_LINE,
                charWidthRatio: WESTERN_CHAR_WIDTH_RATIO,
                columnLeadingRatio: WESTERN_LEADING_RATIO,
                verticalScale: 100,
                hyphenation: true,
                verticalColumnRules: true /* 横組みの段間の罫線は段と段の間の縦棒 / horizontal layout: vertical rules between columns */
            };
        }
        return {
            fontNames: DUMMY_FONT_NAMES,
            sourceText: sourceText,
            textRatio: 1,
            columnCharsPerLine: charsPerLine,
            charWidthRatio: 1, /* 全角 / full-width */
            columnLeadingRatio: DUMMY_LEADING_RATIO,
            verticalScale: DUMMY_VERTICAL_SCALE,
            hyphenation: false,
            verticalColumnRules: false
        };
    }

    /**
     * 新聞ダミーに流す文言を整える。改行は詰め（欧文は空白に）、空なら既定の文言にする。
     * 欧文は繰り返しの継ぎ目で語がくっつかないよう、末尾に空白を足す
     * @param {string} inputText - 入力された文言
     * @param {boolean} isWestern - 欧文なら true
     * @returns {string} 整えた文言
     */
    function normalizeDummyText(inputText, isWestern) {
        var dummyText = String(inputText || "").replace(/[\r\n]+/g, isWestern ? " " : "");
        if (/^\s*$/.test(dummyText)) return isWestern ? WESTERN_DUMMY_TEXT : DUMMY_TEXT;
        if (isWestern && !/\s$/.test(dummyText)) dummyText += " ";
        return dummyText;
    }

    /**
     * ダミー文字列を指定の文字数まで繰り返す
     * @param {string} sourceText - 繰り返す元の文字列
     * @param {number} charCount - 文字数
     * @param {number} startOffset - 元の文字列のどこから始めるか（文字数）
     * @returns {string} ダミー文字列
     */
    function buildDummyText(sourceText, charCount, startOffset) {
        var dummyText = sourceText.substr(startOffset % sourceText.length);
        while (dummyText.length < charCount) {
            dummyText += sourceText;
        }
        return dummyText.substr(0, charCount);
    }

    /**
     * 横組みの新聞ダミーの寸法を求める。
     * 罫線の外側の線から外側の線までを段組数ぶんの段（段間は指定の pt）とし、
     * 段の幅（およその1行の文字数 × 平均字幅）から文字サイズを決める。左右の段数は幅に収まる数
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠
     * @param {{columnCount: number, tierGap: number, dummyWidthPt: number}} headlineSettings - 新聞ダミーの設定（tierGap は pt）
     * @param {Object} typeStyle - getDummyTypeStyle() で求めた書体の設定
     * @returns {Object|null} 新聞ダミーの寸法（作らないときは null）
     */
    function computeColumnLayout(ruleFrame, headlineSettings, typeStyle) {
        var columnCount = headlineSettings.columnCount;
        var headlineSpan = ruleFrame.outerRightX - ruleFrame.outerLeftX;
        var ruleHeight = ruleFrame.topY - ruleFrame.bottomY;
        if (columnCount < 1 || headlineSettings.dummyWidthPt <= 0 || headlineSpan <= 0 || ruleHeight <= 0) return null;

        /* 段の幅を em で表すと、見出しの幅 = 段組数 × 段の幅 ＋（段組数 − 1）× 段間 / the span holds the columns plus the gutters between them */
        var columnWidthEm = typeStyle.columnCharsPerLine * typeStyle.charWidthRatio;
        var gutter = headlineSettings.tierGap;
        var fontSize = (headlineSpan - (columnCount - 1) * gutter) / (columnCount * columnWidthEm);
        if (fontSize <= 0) return null; /* 段間が見出しの幅を食い尽くしたら作らない / the gutters leave no room for columns */
        var columnWidth = columnWidthEm * fontSize;
        var columnPitch = columnWidth + gutter;
        var sideColumnCount = Math.max(1, Math.floor((headlineSettings.dummyWidthPt + gutter) / columnPitch + 1e-6));
        return {
            fontSize: fontSize,
            linePitch: fontSize * typeStyle.columnLeadingRatio,
            gutter: gutter,
            columnWidth: columnWidth,
            columnPitch: columnPitch,
            sideColumnCount: sideColumnCount,
            sideWidth: sideColumnCount * columnPitch - gutter,
            bandHeight: ruleHeight
        };
    }

    /**
     * 罫線の外側の左右に横組みの段組みを描画し、上下には左右の幅いっぱいの段組みを足す。
     * 上下の段組みは罫線と同じ高さで、見出しとの間に横罫を引く
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @param {Object} ruleFrame - 二重線の枠
     * @param {Object} dummyLayout - computeColumnLayout() で求めた寸法
     * @param {number} extraTierCount - 上下それぞれに足す段組みの数
     * @param {Object} typeStyle - getDummyTypeStyle() で求めた書体の設定
     * @param {number} tierRuleWidthPt - 段間の罫線の線幅（pt、0なら引かない）
     * @returns {GroupItem} 生成した新聞ダミーのグループ
     */
    function drawColumnDummy(doc, targetLayer, ruleFrame, dummyLayout, extraTierCount, typeStyle, tierRuleWidthPt) {
        var gutter = dummyLayout.gutter;
        var bandHeight = dummyLayout.bandHeight;

        var dummyGroup = targetLayer.groupItems.add();
        dummyGroup.name = getLabel("itemName.dummyGroup");
        var columnContext = {
            targetLayer: targetLayer,
            dummyGroup: dummyGroup,
            fontSize: dummyLayout.fontSize,
            linePitch: dummyLayout.linePitch,
            columnWidth: dummyLayout.columnWidth,
            columnPitch: dummyLayout.columnPitch,
            gutter: gutter,
            textFont: findDummyFont(typeStyle.fontNames),
            typeStyle: typeStyle,
            tierRuleWidthPt: tierRuleWidthPt,
            strokeColor: createProcessColor(doc, RULE_CMYK),
            columnIndex: 0
        };

        /* 外側の線と本文の間も段間と同じ。見出しの上下で段の並びがそろう / one gutter between the outer rule and the text, so the grid lines up across the headline */
        var leftSideLeft = ruleFrame.outerLeftX - gutter - dummyLayout.sideWidth;
        var rightSideLeft = ruleFrame.outerRightX + gutter;
        var sideColumnCount = dummyLayout.sideColumnCount;
        var columnPitch = dummyLayout.columnPitch;
        var bandPitch = bandHeight + gutter;
        /* 上下は左の段から右の段まで、見出しの上も同じ段の並びで / above and below, the same column grid runs across the headline */
        var fullColumnCount = sideColumnCount * 2 + Math.round((ruleFrame.outerRightX - ruleFrame.outerLeftX + gutter) / columnPitch);
        var fullWidth = fullColumnCount * columnPitch - gutter;

        if (columnContext.typeStyle.joinsSideColumns) {
            /* 英字新聞モード：左右の段は上下の段組みの分までつないで1本の長い段にし、上下は見出しの上の段だけ足す
               English newspaper mode: the side columns run through the bands as single tall columns; bands only over the headline */
            var columnsTop = ruleFrame.topY + extraTierCount * bandPitch;
            var columnsHeight = (ruleFrame.topY - ruleFrame.bottomY) + extraTierCount * bandPitch * 2;
            addDummyColumns(columnContext, leftSideLeft, columnsTop, sideColumnCount, columnsHeight);
            addDummyColumns(columnContext, rightSideLeft, columnsTop, sideColumnCount, columnsHeight);

            var middleLeft = leftSideLeft + sideColumnCount * columnPitch;
            var middleCount = fullColumnCount - sideColumnCount * 2;
            for (var j = 0; j < extraTierCount; j++) {
                var upperBandTop = ruleFrame.topY + (j + 1) * bandPitch;
                var lowerBandTop = ruleFrame.bottomY - gutter - j * bandPitch;
                addDummyColumns(columnContext, middleLeft, upperBandTop, middleCount, bandHeight);
                addDummyColumns(columnContext, middleLeft, lowerBandTop, middleCount, bandHeight);
                /* 左右の長い段との境にも縦の罫 / vertical rules where the bands meet the tall side columns */
                var boundaryXs = [middleLeft - gutter / 2, middleLeft + middleCount * columnPitch - gutter / 2];
                for (var k = 0; k < boundaryXs.length; k++) {
                    addColumnRule(columnContext, boundaryXs[k], upperBandTop, bandHeight);
                    addColumnRule(columnContext, boundaryXs[k], lowerBandTop, bandHeight);
                }
            }
            return dummyGroup;
        }

        addDummyColumns(columnContext, leftSideLeft, ruleFrame.topY, sideColumnCount, bandHeight);
        addDummyColumns(columnContext, rightSideLeft, ruleFrame.topY, sideColumnCount, bandHeight);
        for (var j = 0; j < extraTierCount; j++) {
            var upperRuleY = ruleFrame.topY + gutter / 2 + j * bandPitch;
            if (!columnContext.typeStyle.verticalColumnRules) addTierRule(columnContext, leftSideLeft, upperRuleY, fullWidth);
            addDummyColumns(columnContext, leftSideLeft, upperRuleY + gutter / 2 + bandHeight, fullColumnCount, bandHeight);

            var lowerRuleY = ruleFrame.bottomY - gutter / 2 - j * bandPitch;
            if (!columnContext.typeStyle.verticalColumnRules) addTierRule(columnContext, leftSideLeft, lowerRuleY, fullWidth);
            addDummyColumns(columnContext, leftSideLeft, lowerRuleY - gutter / 2, fullColumnCount, bandHeight);
        }
        return dummyGroup;
    }

    /**
     * 横組みの段を横に並べて作る。段ごとに本文の書き出しをずらす。
     * 欧文では段と段の間の中央に縦の罫を引く
     * @param {Object} columnContext - 段の共通設定（columnIndex を進める）
     * @param {number} columnsLeft - 最初の段の左端の X 座標
     * @param {number} columnsTop - 上端の Y 座標
     * @param {number} columnCount - 段の数
     * @param {number} columnHeight - 段の高さ（pt）
     * @returns {void}
     */
    function addDummyColumns(columnContext, columnsLeft, columnsTop, columnCount, columnHeight) {
        for (var i = 0; i < columnCount; i++) {
            var columnLeft = columnsLeft + i * columnContext.columnPitch;
            addDummyColumn(columnContext, columnLeft, columnsTop, columnHeight);
            columnContext.columnIndex++;
            if (i > 0 && columnContext.typeStyle.verticalColumnRules) {
                addColumnRule(columnContext, columnLeft - columnContext.gutter / 2, columnsTop, columnHeight);
            }
        }
    }

    /**
     * 横組みの段を1つ、両端揃えのエリア内文字で作り、新聞ダミーのグループに入れる。
     * あふれるだけの本文を流し込み、あふれた分は見せない
     * @param {Object} columnContext - 段の共通設定
     * @param {number} columnLeft - 左端の X 座標
     * @param {number} columnTop - 上端の Y 座標
     * @param {number} columnHeight - 高さ（pt）
     * @returns {TextFrame} 生成したテキストフレーム
     */
    function addDummyColumn(columnContext, columnLeft, columnTop, columnHeight) {
        var targetLayer = columnContext.targetLayer;
        var framePath = targetLayer.pathItems.rectangle(columnTop, columnLeft, columnContext.columnWidth, columnHeight);
        var dummyFrame = targetLayer.textFrames.areaText(framePath);

        var lineCount = Math.ceil(columnHeight / columnContext.linePitch) + 1;
        var typeStyle = columnContext.typeStyle;
        var charCount = Math.ceil(lineCount * typeStyle.columnCharsPerLine * 1.5);
        /* 段ごとに書き出しをずらして、同じ文が横に並ばないようにする / start each column elsewhere in the text */
        var columnText = buildDummyText(typeStyle.sourceText, charCount, columnContext.columnIndex * 37);
        /* 欧文は語の途中や空白から始めない / Western columns start at a word, not mid-word or on a space */
        if (typeStyle.hyphenation && columnContext.columnIndex > 0) columnText = columnText.replace(/^\S*\s+/, "");
        dummyFrame.contents = columnText;

        var dummyParagraph = dummyFrame.textRange.paragraphAttributes;
        dummyParagraph.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        dummyParagraph.hyphenation = typeStyle.hyphenation;
        /* 和文は句点が行頭に来ても詰めて流す / let Japanese periods start a line */
        if (!typeStyle.hyphenation) dummyParagraph.kinsoku = "None";

        var dummyAttributes = dummyFrame.textRange.characterAttributes;
        if (columnContext.textFont) dummyAttributes.textFont = columnContext.textFont;
        dummyAttributes.size = columnContext.fontSize;
        dummyAttributes.verticalScale = typeStyle.verticalScale;
        dummyAttributes.autoLeading = false;
        dummyAttributes.leading = columnContext.linePitch;

        dummyFrame.move(columnContext.dummyGroup, ElementPlacement.PLACEATEND);
        return dummyFrame;
    }

    /**
     * 対象・二重線・新聞ダミーをグループ化し、新聞ダミーをぼかし、回転があれば［変形］効果で回転する。
     * 最後に、マスクがONなら回転後も欠けない長方形で全体をマスクする
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {Object} headlineSettings - readHeadlineSettings() でまとめた設定
     * @returns {GroupItem} 全体をまとめたグループ（マスクしたときはクリップグループ）
     */
    function buildHeadlineGroup(doc, targetItem, headlineSettings) {
        var isWestern = headlineSettings.dummyScript === "western";
        var isHorizontal = headlineSettings.dummyDirection === "horizontal";
        /* 横組みの欧文では、見出しの罫線を上下の横棒にする / horizontal Western puts the headline rules above and below */
        var ruleFrame = getRuleFrame(targetItem, headlineSettings, isWestern && isHorizontal);
        var typeStyle = getDummyTypeStyle(isWestern, normalizeDummyText(headlineSettings.dummyText, isWestern), headlineSettings.charsPerLine);
        /* 英字新聞モードでは左右の段を上下までつなぐ / English newspaper mode joins the side columns top to bottom */
        typeStyle.joinsSideColumns = headlineSettings.englishNewspaper === true;
        var dummyLayout = isHorizontal ?
            computeColumnLayout(ruleFrame, headlineSettings, typeStyle) :
            computeTierLayout(ruleFrame, headlineSettings, typeStyle);
        /* 罫線を描く前に、縦組みの新聞ダミーの行に合わせて左右の間隔を調整 / adjust the side spacing to the vertical dummy lines before drawing the rules */
        if (dummyLayout && !isHorizontal) alignRuleFrameToDummyLines(ruleFrame, dummyLayout);

        var ruleGroup = drawSideRules(doc, targetItem, ruleFrame, headlineSettings);
        var dummyGroup = null;
        if (dummyLayout) {
            var drawDummy = isHorizontal ? drawColumnDummy : drawTierDummy;
            dummyGroup = drawDummy(doc, targetItem.layer, ruleFrame, dummyLayout, headlineSettings.extraTierCount,
                typeStyle, headlineSettings.useTierRules ? headlineSettings.tierRuleWidthPt : 0);
        }
        var headlineGroup = groupHeadlineParts(targetItem, ruleGroup, dummyGroup);

        var blursDummy = dummyGroup && headlineSettings.blurRadius > 0;
        if (dummyGroup) {
            /* ぼかす前の範囲で背景を敷く / lay the background over the unblurred dummy bounds */
            addDummyBackground(doc, headlineGroup, dummyGroup.geometricBounds);
            /* ボックス（対象と二重線）以外の新聞ダミーだけをぼかす / blur only the dummy text, not the title box */
            if (blursDummy) applyGaussianBlur(dummyGroup, headlineSettings.blurRadius);
        }

        /* 効果をかける前の範囲。ぼかしでにじむ縁はマスクの外に出す
           Bounds before the effects; keep the blurred fringe outside the mask */
        var contentBounds = headlineGroup.geometricBounds;
        var blurInset = blursDummy ? headlineSettings.blurRadius : 0;

        if (headlineSettings.rotationDeg !== 0) {
            headlineGroup.applyEffect(buildTransformEffectXml(headlineSettings.rotationDeg));
        }
        if (!headlineSettings.useMask) return headlineGroup;
        /* マスクの高さは罫線の天地＋［上下］の長さまで / cap the mask height at the rules plus the Vertical reach */
        var maxMaskHeight = (ruleFrame.topY - ruleFrame.bottomY) + headlineSettings.maskVerticalPt * 2;
        return maskWithInscribedRect(headlineGroup, contentBounds, blurInset, headlineSettings.rotationDeg, maxMaskHeight);
    }

    /**
     * 対象の位置にグループを置き、前面から対象・二重線・新聞ダミーの順に入れる
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {GroupItem|null} ruleGroup - 罫線のグループ（無ければ null）
     * @param {GroupItem|null} dummyGroup - 新聞ダミーのグループ（無ければ null）
     * @returns {GroupItem} まとめたグループ
     */
    function groupHeadlineParts(targetItem, ruleGroup, dummyGroup) {
        var headlineGroup = targetItem.layer.groupItems.add();
        headlineGroup.move(targetItem, ElementPlacement.PLACEBEFORE);
        targetItem.move(headlineGroup, ElementPlacement.PLACEATEND);
        if (ruleGroup) ruleGroup.move(headlineGroup, ElementPlacement.PLACEATEND);
        if (dummyGroup) dummyGroup.move(headlineGroup, ElementPlacement.PLACEATEND);
        return headlineGroup;
    }

    /**
     * 新聞ダミー全体を覆う背景色の長方形を、グループの最背面に置く
     * @param {Document} doc - 対象のドキュメント
     * @param {GroupItem} headlineGroup - 追加先のグループ
     * @param {number[]} dummyBounds - 新聞ダミーの範囲 [左, 上, 右, 下]
     * @returns {PathItem} 生成した長方形
     */
    function addDummyBackground(doc, headlineGroup, dummyBounds) {
        var backgroundRect = headlineGroup.pathItems.rectangle(
            dummyBounds[1],
            dummyBounds[0],
            dummyBounds[2] - dummyBounds[0],
            dummyBounds[1] - dummyBounds[3]
        );
        backgroundRect.filled = true;
        backgroundRect.fillColor = createProcessColor(doc, DUMMY_BACKGROUND_CMYK);
        backgroundRect.stroked = false;
        backgroundRect.move(headlineGroup, ElementPlacement.PLACEATEND);
        return backgroundRect;
    }

    /**
     * 回転後の内容に収まる最大の長方形で、グループ全体をマスクする
     * @param {GroupItem} headlineGroup - マスクするグループ
     * @param {number[]} contentBounds - 回転前の範囲 [左, 上, 右, 下]
     * @param {number} edgeInset - 範囲の四辺から内側に詰める量（pt）
     * @param {number} rotationDeg - 回転角度（°）
     * @param {number} maxMaskHeight - マスクの高さの上限（pt）
     * @returns {GroupItem} クリップグループ（範囲が無いときは headlineGroup のまま）
     */
    function maskWithInscribedRect(headlineGroup, contentBounds, edgeInset, rotationDeg, maxMaskHeight) {
        var contentWidth = contentBounds[2] - contentBounds[0] - edgeInset * 2;
        var contentHeight = contentBounds[1] - contentBounds[3] - edgeInset * 2;
        if (contentWidth <= 0 || contentHeight <= 0) return headlineGroup;

        var maskSize = getInscribedRectSize(contentWidth, contentHeight, rotationDeg);
        maskSize.height = Math.min(maskSize.height, maxMaskHeight);
        /* ［変形］効果は中心を基準に回すので、マスクも中心をそろえる / the Transform effect rotates around the center */
        var centerX = (contentBounds[0] + contentBounds[2]) / 2;
        var centerY = (contentBounds[1] + contentBounds[3]) / 2;

        var clipGroup = headlineGroup.layer.groupItems.add();
        clipGroup.move(headlineGroup, ElementPlacement.PLACEBEFORE);
        headlineGroup.move(clipGroup, ElementPlacement.PLACEATEND);

        var maskRect = clipGroup.pathItems.rectangle(
            centerY + maskSize.height / 2,
            centerX - maskSize.width / 2,
            maskSize.width,
            maskSize.height
        );
        /* クリップグループは最前面のパスでマスクする / a clip group masks with its frontmost path */
        maskRect.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        clipGroup.clipped = true;
        return clipGroup;
    }

    /**
     * 幅 × 高さの長方形を回転させたとき、その内側に収まる面積最大の水平な長方形の寸法を返す
     * @param {number} rectWidth - 回転前の幅（pt）
     * @param {number} rectHeight - 回転前の高さ（pt）
     * @param {number} rotationDeg - 回転角度（°）
     * @returns {{width: number, height: number}} 内接する長方形の寸法（pt）
     */
    function getInscribedRectSize(rectWidth, rectHeight, rotationDeg) {
        var angleRad = rotationDeg * Math.PI / 180;
        var sinA = Math.abs(Math.sin(angleRad));
        var cosA = Math.abs(Math.cos(angleRad));
        if (sinA < 1e-10) return { width: rectWidth, height: rectHeight };

        var widthIsLonger = rectWidth >= rectHeight;
        var longSide = widthIsLonger ? rectWidth : rectHeight;
        var shortSide = widthIsLonger ? rectHeight : rectWidth;

        /* 細長い場合や45°付近は、短辺の中点2つで決まる / thin rectangles and angles near 45° are bound by the short side */
        if (shortSide <= 2 * sinA * cosA * longSide || Math.abs(sinA - cosA) < 1e-10) {
            var halfShort = shortSide / 2;
            return widthIsLonger ?
                { width: halfShort / sinA, height: halfShort / cosA } :
                { width: halfShort / cosA, height: halfShort / sinA };
        }
        /* それ以外は4辺すべてが回転した長方形に接する / otherwise all four corners touch the rotated rectangle */
        var cos2A = cosA * cosA - sinA * sinA;
        return {
            width: (rectWidth * cosA - rectHeight * sinA) / cos2A,
            height: (rectHeight * cosA - rectWidth * sinA) / cos2A
        };
    }

    /**
     * ぼかし（ガウス）をライブエフェクトとして適用する
     * @param {PageItem} targetItem - 対象のページアイテム
     * @param {number} blurRadius - ぼかしの半径（px）
     * @returns {void}
     */
    function applyGaussianBlur(targetItem, blurRadius) {
        var effectXml = '<LiveEffect name="Adobe PSL Gaussian Blur">' +
            '<Dict data="R blur ' + blurRadius + ' R PrevDocScale 1 I PrevDres ' + EFFECT_RESOLUTION + ' "/>' +
            '</LiveEffect>';
        targetItem.applyEffect(effectXml);
    }

    /**
     * 中心を基準に回転だけを行う［変形］効果の XML を作る
     * @param {number} rotateDegrees - 回転角度（°、反時計回りが正）
     * @returns {string} LiveEffect の XML
     */
    function buildTransformEffectXml(rotateDegrees) {
        return '<LiveEffect name="Adobe Transform"><Dict data="' +
            'R scaleH_Percent 100' +
            ' R scaleV_Percent 100' +
            ' R scaleH_Factor 1' +
            ' R scaleV_Factor 1' +
            ' R moveH_Pts 0' +
            ' R moveV_Pts 0' +
            ' R rotate_Degrees ' + rotateDegrees +
            ' R rotate_Radians ' + (rotateDegrees * Math.PI / 180) +
            ' I numCopies 0' +
            ' I pinPoint 4' +
            ' B scaleLines 1' +
            ' B transformPatterns 1' +
            ' B transformObjects 1' +
            ' B reflectX 0' +
            ' B reflectY 0' +
            ' B randomize 0' +
            ' "/></LiveEffect>';
    }

    /**
     * 直線を1本描く
     * @param {GroupItem} parentGroup - 追加先のグループ
     * @param {number[]} startPoint - 始点 [X, Y]
     * @param {number[]} endPoint - 終点 [X, Y]
     * @param {Object} strokeColor - 線の色
     * @param {number} strokeWidth - 線幅（pt）
     * @returns {PathItem} 生成したパス
     */
    function createRuleLine(parentGroup, startPoint, endPoint, strokeColor, strokeWidth) {
        var linePath = parentGroup.pathItems.add();
        linePath.setEntirePath([startPoint, endPoint]);
        linePath.stroked = true;
        linePath.filled = false;
        linePath.strokeWidth = strokeWidth;
        linePath.strokeColor = strokeColor;
        linePath.strokeCap = StrokeCap.BUTTENDCAP;
        return linePath;
    }

    /**
     * ドキュメントのカラーモードに合わせた色を CMYK の値から作る。RGB では単純換算した近似色
     * @param {Document} doc - 対象のドキュメント
     * @param {number[]} cmykValues - [C, M, Y, K]（%）
     * @returns {Object} RGBColor または CMYKColor
     */
    function createProcessColor(doc, cmykValues) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var keyRatio = 1 - cmykValues[3] / 100;
            var rgbColor = new RGBColor();
            rgbColor.red = Math.round(255 * (1 - cmykValues[0] / 100) * keyRatio);
            rgbColor.green = Math.round(255 * (1 - cmykValues[1] / 100) * keyRatio);
            rgbColor.blue = Math.round(255 * (1 - cmykValues[2] / 100) * keyRatio);
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cmykValues[0];
        cmykColor.magenta = cmykValues[1];
        cmykColor.yellow = cmykValues[2];
        cmykColor.black = cmykValues[3];
        return cmykColor;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログと各コントロールを作る。パネルはそれぞれの関数で作り、選択の連動は bindDummyChoices() でつなぐ
     * @param {string} unitLabel - 入力欄に添える単位の表示名
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @param {string} initialDummyScript - 和文／欧文の初期値（"japanese" / "western"）
     * @param {string} initialDummyDirection - 組み方向の初期値（"vertical" / "horizontal"）
     * @param {number} initialTierCount - 縦組みの［段数］の初期値
     * @param {boolean} allowsWestern - ［欧文］を選べるか（縦組みの見出しでは false）
     * @returns {{dialog: Window, fieldInputs: Object, maskCheckbox: Checkbox, tierRuleCheckbox: Checkbox, englishNewspaperCheckbox: Checkbox, fitViewControls: Object, getDummyScript: Function, getDummyDirection: Function, getRuleStyle: Function, getDummyText: Function}} ダイアログ、キーごとの入力欄、マスク・段間の罫線・英字新聞モードのチェックボックス、［画面にフィット］の部品、和文／欧文・組み方向・罫線の種類・文章の取得
     */
    function buildDialog(unitLabel, pointsPerUnit, onValueChange, initialDummyScript, initialDummyDirection, initialTierCount, allowsWestern) {
        var headlineDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(headlineDialog);

        /* パネルの関数で共有する入力欄・単位・状態。表示前のラジオの value は読み戻せないので、選択は state に持つ
           Shared by the panel builders; radio values cannot be read back before show(), so choices live in state */
        var dialogContext = {
            fieldInputs: {},
            pointsPerUnit: pointsPerUnit,
            lengthUnit: " " + unitLabel, /* 長さの欄は単位を中に入れる / length fields hold their unit */
            strokeUnitInfo: getUnitInfo("strokeUnits"), /* 線幅は線の単位 / stroke weight uses the stroke unit */
            allowsWestern: allowsWestern,
            onValueChange: onValueChange,
            onFieldChange: null,
            state: {
                ruleStyle: DEFAULT_RULE_STYLE,
                dummyScript: initialDummyScript,
                dummyDirection: initialDummyDirection,
                dummyTexts: { japanese: DUMMY_TEXT, western: WESTERN_DUMMY_TEXT },
                tierCountTexts: { vertical: String(initialTierCount), horizontal: String(DEFAULT_COLUMN_COUNT) },
                lastExtraTierCount: DEFAULT_EXTRA_TIER_COUNT
            },
            controls: {}
        };
        var fieldInputs = dialogContext.fieldInputs;
        var controls = dialogContext.controls;

        /**
         * 数値欄の変更を伝える。［段数追加］が0になったらマスクをOFFにする（0のままでも手動でONに戻せる）
         * @returns {void}
         */
        dialogContext.onFieldChange = function () {
            /* 空欄は入力途中なので0と見なさない（NaN）/ an empty field is mid-edit, not 0 */
            var extraTierCount = Math.floor(readFieldAsNumber(fieldInputs.extraTierCount, NaN, false));
            if (extraTierCount === 0 && dialogContext.state.lastExtraTierCount !== 0) {
                controls.maskCheckbox.value = false;
                controls.updateMaskFields(true);
            }
            dialogContext.state.lastExtraTierCount = extraTierCount;
            onValueChange();
        };

        /* 2カラム：左にマージン・罫線・マスクと回転、右に組版設定・ダミー文章 / two columns: margins, rules, and mask and rotation on the left; typesetting and dummy text on the right */
        var columnsGroup = headlineDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = addDialogColumn(columnsGroup);
        var rightColumn = addDialogColumn(columnsGroup);

        buildMarginPanel(leftColumn, dialogContext);
        buildRulePanel(leftColumn, dialogContext);
        buildTypesettingPanel(rightColumn, dialogContext);
        buildDummyTextPanel(rightColumn, dialogContext);
        buildMaskRotationPanel(leftColumn, dialogContext);
        bindDummyChoices(dialogContext);
        alignDialogLabels(dialogContext);

        var buttonRow = addButtonRow(headlineDialog);
        /* ボタン行の左に［画面にフィット］/ Fit to Window on the left of the button row */
        var fitViewControls = FitViewToItems.addControls(buttonRow.leftGroup, { value: DEFAULT_FIT_VIEW, stepper: true });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /* option＋B で［ぼかし］を切り替える。数値欄にフォーカスがあっても効かせる / Option+B toggles Blur, even from the numeric fields */
        var numericFields = [fitViewControls.percentInput];
        for (var fieldKey in fieldInputs) {
            if (fieldInputs.hasOwnProperty(fieldKey)) numericFields.push(fieldInputs[fieldKey]);
        }
        addKeyShortcuts(headlineDialog, {
            "Alt+B": fieldInputs.blurRadius.optionCheckbox
        }, { numericFields: numericFields, showInTip: true });

        var state = dialogContext.state;
        return {
            dialog: headlineDialog,
            fieldInputs: fieldInputs,
            maskCheckbox: controls.maskCheckbox,
            tierRuleCheckbox: controls.tierRuleCheckbox,
            englishNewspaperCheckbox: controls.englishNewspaperCheckbox,
            fitViewControls: fitViewControls,
            getDummyScript: function () { return state.dummyScript; },
            getDummyDirection: function () { return state.dummyDirection; },
            getRuleStyle: function () { return state.ruleStyle; },
            getDummyText: function () { return state.dummyTexts[state.dummyScript]; }
        };
    }

    /**
     * ダイアログの縦のカラム（パネルを縦に積む group）を追加する
     * @param {Group} columnsGroup - 追加先の横並びの group
     * @returns {Group} 追加したカラム
     */
    function addDialogColumn(columnsGroup) {
        var dialogColumn = columnsGroup.add("group");
        dialogColumn.orientation = "column";
        dialogColumn.alignChildren = ["fill", "top"];
        dialogColumn.spacing = WINDOW_SPACING;
        return dialogColumn;
    }

    /**
     * ［マージン］パネルを作る。上下・左右の欄を縦に積み、右に連動のリンクアイコンを置く
     * @param {Group} parentColumn - 追加先のカラム
     * @param {Object} dialogContext - buildDialog() の共有データ
     * @returns {void}
     */
    function buildMarginPanel(parentColumn, dialogContext) {
        var fieldInputs = dialogContext.fieldInputs;
        var marginPanel = parentColumn.add("panel", undefined, getLabel("panel.margin"));
        setupPanel(marginPanel);
        var marginRow = marginPanel.add("group");
        marginRow.orientation = "row";
        marginRow.alignChildren = ["left", "center"];
        marginRow.spacing = FIELD_ROW_SPACING;
        var marginFieldsGroup = marginRow.add("group");
        marginFieldsGroup.orientation = "column";
        marginFieldsGroup.alignChildren = ["left", "center"];
        marginFieldsGroup.spacing = marginPanel.spacing;
        marginFieldsGroup.margins = 0;
        var marginLinkToggle = null;

        /**
         * 連動中なら、変更した側の値をもう一方へ写してから変更を通知する
         * @param {string} sourceKey - 変更した欄のキー
         * @returns {void}
         */
        function handleMarginChange(sourceKey) {
            if (marginLinkToggle && marginLinkToggle.value) {
                var targetKey = (sourceKey === "marginVertical") ? "marginHorizontal" : "marginVertical";
                fieldInputs[targetKey].text = fieldInputs[sourceKey].text;
            }
            dialogContext.onFieldChange();
        }

        fieldInputs.marginVertical = addNumericFieldRow(marginFieldsGroup,
            { key: "marginVertical", initialText: formatUnitValue(DEFAULT_VERTICAL_MARGIN_PT, dialogContext.pointsPerUnit), fieldUnit: dialogContext.lengthUnit },
            function () { handleMarginChange("marginVertical"); });
        fieldInputs.marginHorizontal = addNumericFieldRow(marginFieldsGroup,
            { key: "marginHorizontal", initialText: formatUnitValue(DEFAULT_HORIZONTAL_MARGIN_PT, dialogContext.pointsPerUnit), fieldUnit: dialogContext.lengthUnit },
            function () { handleMarginChange("marginHorizontal"); });
        /* 連動をONにしたら上下の値に左右をそろえる / turning the link on copies the vertical value to the horizontal one */
        marginLinkToggle = addLinkToggle(marginRow, DEFAULT_LINK_MARGINS, function () { handleMarginChange("marginVertical"); });
        marginLinkToggle.helpTip = getLabel("tooltip.linkMargins");
        if (DEFAULT_LINK_MARGINS) fieldInputs.marginHorizontal.text = fieldInputs.marginVertical.text;
    }

    /**
     * ［罫線］パネルを作る。種類のラジオボタン（4つなので縦に積む）と、線幅・線の間隔
     * @param {Group} parentColumn - 追加先のカラム
     * @param {Object} dialogContext - buildDialog() の共有データ（controls.ruleStyleRadios と controls.updateRuleFields を足す）
     * @returns {void}
     */
    function buildRulePanel(parentColumn, dialogContext) {
        var fieldInputs = dialogContext.fieldInputs;
        var state = dialogContext.state;
        var strokeUnitInfo = dialogContext.strokeUnitInfo;
        var rulePanel = parentColumn.add("panel", undefined, getLabel("panel.sideRules"));
        setupPanel(rulePanel);

        /* 4つあるので縦に積み、ダイアログの幅を広げない / stacked so the four choices do not widen the dialog */
        var ruleStyleGroup = rulePanel.add("group");
        ruleStyleGroup.orientation = "column";
        ruleStyleGroup.alignChildren = ["left", "center"];
        ruleStyleGroup.spacing = RADIO_COLUMN_SPACING;
        var ruleStyleOptions = [
            { name: "double", labelKey: "ruleDouble" },
            { name: "single", labelKey: "ruleSingle" },
            { name: "singleTop", labelKey: "ruleSingleTop" },
            { name: "none", labelKey: "ruleNone" }
        ];
        var ruleStyleRadios = {};
        for (var i = 0; i < ruleStyleOptions.length; i++) {
            var ruleStyleRadio = ruleStyleGroup.add("radiobutton", undefined, getLabel("radio." + ruleStyleOptions[i].labelKey));
            ruleStyleRadios[ruleStyleOptions[i].name] = ruleStyleRadio;
            ruleStyleRadio.helpTip = getLabel("tooltip." + ruleStyleOptions[i].labelKey);
            ruleStyleRadio.value = (ruleStyleOptions[i].name === state.ruleStyle);
            ruleStyleRadio.ruleStyle = ruleStyleOptions[i].name;
            ruleStyleRadio.onClick = function () {
                state.ruleStyle = this.ruleStyle;
                updateRuleFields(true);
                dialogContext.onValueChange();
            };
        }

        addNumericFieldRows(rulePanel, [
            {
                key: "ruleStrokeWidth", fieldUnit: " " + strokeUnitInfo.label,
                initialText: formatUnitValue(DEFAULT_RULE_STROKE_WIDTH_PT, strokeUnitInfo.pointsPerUnit)
            },
            { key: "lineGap", initialText: formatUnitValue(DEFAULT_LINE_GAP_PT, dialogContext.pointsPerUnit), fieldUnit: dialogContext.lengthUnit }
        ], fieldInputs, dialogContext.onFieldChange);

        /**
         * ［線幅］は［なし］以外、［線の間隔］は［二重］のときだけ有効にする
         * @param {boolean} redrawsSteppers - ∧∨も描き直すか（表示前は false）
         * @returns {void}
         */
        function updateRuleFields(redrawsSteppers) {
            setFieldRowEnabled(fieldInputs.ruleStrokeWidth, state.ruleStyle !== "none", redrawsSteppers);
            setFieldRowEnabled(fieldInputs.lineGap, state.ruleStyle === "double", redrawsSteppers);
        }
        updateRuleFields(false);

        dialogContext.controls.ruleStyleRadios = ruleStyleRadios;
        dialogContext.controls.updateRuleFields = updateRuleFields;
    }

    /**
     * ［組版設定］パネルを作る。組み方向・英字新聞モード・字詰・段数・段間・段数追加・段間の罫線と線幅・ぼかし
     * @param {Group} parentColumn - 追加先のカラム
     * @param {Object} dialogContext - buildDialog() の共有データ（組み方向のラジオと各チェックボックスを controls に足す）
     * @returns {void}
     */
    function buildTypesettingPanel(parentColumn, dialogContext) {
        var fieldInputs = dialogContext.fieldInputs;
        var controls = dialogContext.controls;
        var state = dialogContext.state;
        var strokeUnitInfo = dialogContext.strokeUnitInfo;
        var typesettingPanel = parentColumn.add("panel", undefined, getLabel("panel.typesetting"));
        setupPanel(typesettingPanel);

        /* 組み方向：縦組み（段）／横組み（段組み） / direction: vertical tiers or horizontal columns */
        var directionRow = typesettingPanel.add("group");
        setupRow(directionRow, undefined, FIELD_ROW_SPACING);
        controls.directionLabel = directionRow.add("statictext", undefined, labelText("fieldLabel.dummyDirection"));
        controls.verticalRadio = directionRow.add("radiobutton", undefined, getLabel("radio.vertical"));
        controls.horizontalRadio = directionRow.add("radiobutton", undefined, getLabel("radio.horizontal"));
        controls.verticalRadio.helpTip = getLabel("tooltip.vertical");
        controls.horizontalRadio.helpTip = getLabel("tooltip.horizontal");
        controls.verticalRadio.value = (state.dummyDirection !== "horizontal");
        controls.horizontalRadio.value = (state.dummyDirection === "horizontal");

        /* 英字新聞モード：ON の間は横組み・欧文に固定し、罫線を［1本（上のみ）］にする / English newspaper mode */
        var englishNewspaperCheckbox = addCheckboxRow(typesettingPanel, "englishNewspaper", false, dialogContext.onValueChange);
        englishNewspaperCheckbox.enabled = dialogContext.allowsWestern; /* 縦組みの見出しでは欧文を使わない / not for vertical headlines */
        if (!dialogContext.allowsWestern) englishNewspaperCheckbox.helpTip = getLabel("tooltip.englishNewspaperVertical");
        controls.englishNewspaperCheckbox = englishNewspaperCheckbox;

        addNumericFieldRows(typesettingPanel, [
            { key: "charsPerLine", initialText: String(DEFAULT_CHARS_PER_LINE), unitLabel: getLabel("unit.chars"), integer: true, min: 1 },
            /* ［段数］は縦組みと横組みで意味を切り替える。単位は長い方で幅を取っておく / one Tiers field whose meaning follows the direction; the unit is sized for the longer label */
            { key: "tierCount", initialText: state.tierCountTexts[state.dummyDirection], unitLabel: getLabel("unit.columns"), integer: true },
            { key: "tierGap", initialText: formatUnitValue(DEFAULT_TIER_GAP_PT, dialogContext.pointsPerUnit), fieldUnit: dialogContext.lengthUnit },
            { key: "extraTierCount", initialText: String(DEFAULT_EXTRA_TIER_COUNT), unitLabel: getLabel("unit.tiers"), integer: true }
        ], fieldInputs, dialogContext.onFieldChange);
        /* 単位は長い方（columns）で作ってあるので、短い単位に替えても幅を保つ / the unit was created with the longer label; keep that width */
        fieldInputs.tierCount.unitText.preferredSize.width = fieldInputs.tierCount.unitText.preferredSize.width;

        var tierRuleCheckbox = addCheckboxRow(typesettingPanel, "useTierRules", DEFAULT_USE_TIER_RULES, dialogContext.onValueChange);
        addNumericFieldRows(typesettingPanel, [
            {
                key: "tierRuleWidth", fieldUnit: " " + strokeUnitInfo.label,
                initialText: formatUnitValue(DEFAULT_TIER_RULE_WIDTH_PT, strokeUnitInfo.pointsPerUnit)
            },
            { key: "blurRadius", initialText: String(DEFAULT_BLUR_RADIUS), unitLabel: "px", checkboxKey: "useBlur", checked: DEFAULT_USE_BLUR }
        ], fieldInputs, dialogContext.onFieldChange);

        /* ［段間の罫線］がOFFなら［線幅］をディムにする / dim Weight while Tier rules is off */
        tierRuleCheckbox.onClick = function () {
            setFieldRowEnabled(fieldInputs.tierRuleWidth, tierRuleCheckbox.value, true);
            dialogContext.onValueChange();
        };
        setFieldRowEnabled(fieldInputs.tierRuleWidth, DEFAULT_USE_TIER_RULES, false);
        controls.tierRuleCheckbox = tierRuleCheckbox;
    }

    /**
     * ［ダミー文章］パネルを作る。和文／欧文のラジオボタンと、文章を設定する［文章...］ボタン
     * @param {Group} parentColumn - 追加先のカラム
     * @param {Object} dialogContext - buildDialog() の共有データ（和文／欧文のラジオを controls に足す）
     * @returns {void}
     */
    function buildDummyTextPanel(parentColumn, dialogContext) {
        var controls = dialogContext.controls;
        var state = dialogContext.state;
        var dummyTextPanel = parentColumn.add("panel", undefined, getLabel("panel.dummyText"));
        setupPanel(dummyTextPanel);

        var scriptRow = dummyTextPanel.add("group");
        setupRow(scriptRow);
        controls.japaneseRadio = scriptRow.add("radiobutton", undefined, getLabel("radio.japanese"));
        controls.westernRadio = scriptRow.add("radiobutton", undefined, getLabel("radio.western"));
        controls.japaneseRadio.helpTip = getLabel("tooltip.japanese");
        controls.westernRadio.helpTip = getLabel(dialogContext.allowsWestern ? "tooltip.western" : "tooltip.westernVertical");
        controls.westernRadio.enabled = dialogContext.allowsWestern; /* 縦組みの見出しには欧文を組み合わせない / no Western text for vertical headlines */
        controls.japaneseRadio.value = (state.dummyScript !== "western");
        controls.westernRadio.value = (state.dummyScript === "western");

        /* 文章は和文・欧文で別々に持ち、別ダイアログボックスで設定する / the text is kept per script and edited in a separate dialog */
        var dummyTextRow = dummyTextPanel.add("group");
        setupRow(dummyTextRow);
        var btnDummyText = dummyTextRow.add("button", undefined, getLabel("button.dummyText"));
        btnDummyText.helpTip = getLabel("tooltip.dummyText");
        btnDummyText.onClick = function () {
            if (showDummyTextDialog(state.dummyTexts)) dialogContext.onValueChange();
        };
    }

    /**
     * ［マスクと回転］パネルを作る。上から回転・マスク・左右・上下
     * @param {Group} parentColumn - 追加先のカラム
     * @param {Object} dialogContext - buildDialog() の共有データ（controls.maskCheckbox と controls.updateMaskFields を足す）
     * @returns {void}
     */
    function buildMaskRotationPanel(parentColumn, dialogContext) {
        var fieldInputs = dialogContext.fieldInputs;
        var maskRotationPanel = addNumericFieldPanel(parentColumn, "panel.maskRotation", [
            { key: "rotation", initialText: String(DEFAULT_ROTATION_DEG), unitLabel: "°", allowNegative: true }
        ], fieldInputs, dialogContext.onFieldChange);
        var maskCheckbox = addCheckboxRow(maskRotationPanel, "useMask", DEFAULT_USE_MASK && DEFAULT_EXTRA_TIER_COUNT > 0, dialogContext.onValueChange);
        addNumericFieldRows(maskRotationPanel, [
            { key: "dummyWidth", initialText: formatUnitValue(DEFAULT_DUMMY_WIDTH_PT, dialogContext.pointsPerUnit), fieldUnit: dialogContext.lengthUnit },
            { key: "maskVertical", initialText: formatUnitValue(DEFAULT_MASK_VERTICAL_PT, dialogContext.pointsPerUnit), fieldUnit: dialogContext.lengthUnit }
        ], fieldInputs, dialogContext.onFieldChange);

        /**
         * マスクがOFFなら［左右］［上下］をディムにする
         * @param {boolean} redrawsSteppers - ∧∨も描き直すか（表示前は false）
         * @returns {void}
         */
        function updateMaskFields(redrawsSteppers) {
            setFieldRowEnabled(fieldInputs.dummyWidth, maskCheckbox.value, redrawsSteppers);
            setFieldRowEnabled(fieldInputs.maskVertical, maskCheckbox.value, redrawsSteppers);
        }
        maskCheckbox.onClick = function () {
            updateMaskFields(true);
            dialogContext.onValueChange();
        };
        updateMaskFields(false);

        dialogContext.controls.maskCheckbox = maskCheckbox;
        dialogContext.controls.updateMaskFields = updateMaskFields;
    }

    /**
     * 和文／欧文・組み方向・英字新聞モード・罫線の種類の連動をつなぐ。
     * ［段数］の値は組み方向ごとに持ち、［1本（上のみ）］は横組みの欧文のときだけ選べる
     * @param {Object} dialogContext - buildDialog() の共有データ（パネルをすべて作ったあとに呼ぶ）
     * @returns {void}
     */
    function bindDummyChoices(dialogContext) {
        var fieldInputs = dialogContext.fieldInputs;
        var controls = dialogContext.controls;
        var state = dialogContext.state;
        var ruleStyleRadios = controls.ruleStyleRadios;
        /* 英字新聞モードを ON にする前の選択（OFF で戻す）/ choices before English newspaper mode, restored on OFF */
        var choicesBeforeEnglishMode = null;

        /**
         * 和文／欧文を選び、ラジオボタンもそろえる（プレビューはしない）
         * @param {string} nextScript - "japanese" または "western"
         * @returns {void}
         */
        function setDummyScript(nextScript) {
            state.dummyScript = nextScript;
            controls.japaneseRadio.value = (nextScript !== "western");
            controls.westernRadio.value = (nextScript === "western");
        }

        /**
         * 縦組み／横組みを選び、ラジオボタンと［段数］もそろえる（プレビューはしない）。［段数］の値は組み方向ごとに持つ
         * @param {string} nextDirection - "vertical" または "horizontal"
         * @returns {void}
         */
        function setDummyDirection(nextDirection) {
            state.tierCountTexts[state.dummyDirection] = fieldInputs.tierCount.text;
            state.dummyDirection = nextDirection;
            fieldInputs.tierCount.text = state.tierCountTexts[nextDirection];
            controls.verticalRadio.value = (nextDirection !== "horizontal");
            controls.horizontalRadio.value = (nextDirection === "horizontal");
        }

        /**
         * 字詰を有効／無効にし（欧文の横組みだけ無効）、［段数］の説明と単位を組み方向に合わせる。
         * ［1本（上のみ）］は横組みの欧文のときだけ選べ、選べなくなったら［1本（上下）］に戻す
         * @param {boolean} redrawsSteppers - ∧∨も描き直すか（表示前は false）
         * @returns {void}
         */
        function updateDummyFields(redrawsSteppers) {
            var isHorizontal = state.dummyDirection === "horizontal";
            var isHorizontalWestern = isHorizontal && state.dummyScript === "western";
            setFieldRowEnabled(fieldInputs.charsPerLine, !isHorizontalWestern, redrawsSteppers);
            ruleStyleRadios.singleTop.enabled = isHorizontalWestern;
            if (!isHorizontalWestern && state.ruleStyle === "singleTop") {
                state.ruleStyle = "single";
                ruleStyleRadios.single.value = true;
                ruleStyleRadios.singleTop.value = false;
                controls.updateRuleFields(redrawsSteppers);
            }
            fieldInputs.tierCount.helpTip = getLabel(isHorizontal ? "tooltip.columnCount" : "tooltip.tierCount");
            fieldInputs.tierCount.unitText.text = getLabel(isHorizontal ? "unit.columns" : "unit.tiers");
        }

        /**
         * 選択を反映してプレビューする
         * @returns {void}
         */
        function refreshAfterChoice() {
            updateDummyFields(true);
            dialogContext.onValueChange();
        }

        /**
         * 英字新聞モードの ON／OFF を反映する。ON で横組み・欧文・［1本（上のみ）］にして組み方向と和文／欧文を固定、
         * OFF で元の組み方向と和文／欧文に戻す
         * @returns {void}
         */
        function applyEnglishNewspaperMode() {
            var isOn = controls.englishNewspaperCheckbox.value;
            if (isOn) {
                choicesBeforeEnglishMode = { script: state.dummyScript, direction: state.dummyDirection };
                setDummyScript("western");
                setDummyDirection("horizontal");
                state.ruleStyle = "singleTop";
                for (var styleName in ruleStyleRadios) {
                    if (ruleStyleRadios.hasOwnProperty(styleName)) ruleStyleRadios[styleName].value = (styleName === "singleTop");
                }
            } else if (choicesBeforeEnglishMode) {
                setDummyScript(choicesBeforeEnglishMode.script);
                setDummyDirection(choicesBeforeEnglishMode.direction);
                choicesBeforeEnglishMode = null;
            }
            controls.verticalRadio.enabled = !isOn;
            controls.horizontalRadio.enabled = !isOn;
            controls.japaneseRadio.enabled = !isOn;
            controls.westernRadio.enabled = !isOn && dialogContext.allowsWestern;
            controls.updateRuleFields(true);
            refreshAfterChoice();
        }

        controls.japaneseRadio.onClick = function () { state.dummyScript = "japanese"; refreshAfterChoice(); };
        controls.westernRadio.onClick = function () { state.dummyScript = "western"; refreshAfterChoice(); };
        controls.verticalRadio.onClick = function () { setDummyDirection("vertical"); refreshAfterChoice(); };
        controls.horizontalRadio.onClick = function () { setDummyDirection("horizontal"); refreshAfterChoice(); };
        controls.englishNewspaperCheckbox.onClick = applyEnglishNewspaperMode;
        updateDummyFields(false);
    }

    /**
     * 左右のカラムそれぞれで、項目名をいちばん長いものの実測幅にそろえて右揃えにする。
     * チェックボックスの行の字下げと、［ぼかし］のチェックボックスの幅も同じにして入力欄の列を合わせる
     * @param {Object} dialogContext - buildDialog() の共有データ（パネルをすべて作ったあとに呼ぶ）
     * @returns {void}
     */
    function alignDialogLabels(dialogContext) {
        var fieldInputs = dialogContext.fieldInputs;
        var controls = dialogContext.controls;

        var leftLabelWidth = alignLabelWidths([
            fieldInputs.marginVertical.fieldLabel,
            fieldInputs.marginHorizontal.fieldLabel,
            fieldInputs.ruleStrokeWidth.fieldLabel,
            fieldInputs.lineGap.fieldLabel,
            fieldInputs.rotation.fieldLabel,
            fieldInputs.dummyWidth.fieldLabel,
            fieldInputs.maskVertical.fieldLabel
        ], "right");
        controls.maskCheckbox.labelSpacer.preferredSize.width = leftLabelWidth;

        var typesettingLabels = [
            controls.directionLabel,
            fieldInputs.charsPerLine.fieldLabel,
            fieldInputs.tierCount.fieldLabel,
            fieldInputs.tierGap.fieldLabel,
            fieldInputs.extraTierCount.fieldLabel,
            fieldInputs.tierRuleWidth.fieldLabel
        ];
        var blurCheckbox = fieldInputs.blurRadius.optionCheckbox;
        var typesettingLabelWidth = Math.max(alignLabelWidths(typesettingLabels, "right"), blurCheckbox.preferredSize.width);
        for (var i = 0; i < typesettingLabels.length; i++) {
            typesettingLabels[i].preferredSize.width = typesettingLabelWidth;
        }
        blurCheckbox.preferredSize.width = typesettingLabelWidth;
        controls.tierRuleCheckbox.labelSpacer.preferredSize.width = typesettingLabelWidth;
        controls.englishNewspaperCheckbox.labelSpacer.preferredSize.width = typesettingLabelWidth;
    }

    /**
     * 新聞ダミーの文言を和文・欧文それぞれ編集するダイアログボックスを表示する。OK なら dummyTexts を書き換える
     * @param {{japanese: string, western: string}} dummyTexts - 和文・欧文の文言（OK で書き換える）
     * @returns {boolean} OK で閉じたら true
     */
    function showDummyTextDialog(dummyTexts) {
        var textDialog = new Window("dialog", getLabel("dialog.dummyTextTitle"));
        setupWindow(textDialog);

        var scriptKeys = ["japanese", "western"];
        var textInputs = {};
        for (var i = 0; i < scriptKeys.length; i++) {
            var textPanel = textDialog.add("panel", undefined, getLabel("radio." + scriptKeys[i]));
            setupPanel(textPanel);
            var textInput = textPanel.add("edittext", undefined, dummyTexts[scriptKeys[i]], { multiline: true, scrolling: true });
            textInput.preferredSize = DUMMY_TEXT_INPUT_SIZE;
            textInput.helpTip = getLabel("tooltip.dummyTextInput");
            textInputs[scriptKeys[i]] = textInput;
        }

        var buttonRow = addButtonRow(textDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(textDialog, SCRIPT_NAME + "DummyText");

        if (textDialog.show() !== 1) return false;
        for (var j = 0; j < scriptKeys.length; j++) {
            dummyTexts[scriptKeys[j]] = textInputs[scriptKeys[j]].text;
        }
        return true;
    }

    /**
     * パネルを作り、数値入力の行を並べる
     * @param {Group} parentColumn - 追加先のカラム
     * @param {string} panelLabelPath - パネル見出しのラベルパス
     * @param {Object[]} fieldSpecs - 行の定義（key, initialText, unitLabel, allowNegative, integer, min）
     * @param {Object} fieldInputs - 作った入力欄を key ごとに入れる先
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {Panel} 作ったパネル
     */
    function addNumericFieldPanel(parentColumn, panelLabelPath, fieldSpecs, fieldInputs, onValueChange) {
        var fieldPanel = parentColumn.add("panel", undefined, getLabel(panelLabelPath));
        setupPanel(fieldPanel);
        addNumericFieldRows(fieldPanel, fieldSpecs, fieldInputs, onValueChange);
        return fieldPanel;
    }

    /**
     * 数値入力の行を順に並べる
     * @param {Panel|Group} parentContainer - 追加先
     * @param {Object[]} fieldSpecs - 行の定義（addNumericFieldRow() を参照）
     * @param {Object} fieldInputs - 作った入力欄を key ごとに入れる先
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {void}
     */
    function addNumericFieldRows(parentContainer, fieldSpecs, fieldInputs, onValueChange) {
        for (var i = 0; i < fieldSpecs.length; i++) {
            fieldInputs[fieldSpecs[i].key] = addNumericFieldRow(parentContainer, fieldSpecs[i], onValueChange);
        }
    }

    /**
     * 数値入力の行（項目名・∧∨・入力欄・単位）をまとめて有効／無効にする
     * @param {EditText} inputField - addNumericFieldRow() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @param {boolean} redrawsSteppers - ∧∨も描き直すか（表示前は false）
     * @returns {void}
     */
    function setFieldRowEnabled(inputField, isEnabled, redrawsSteppers) {
        inputField.fieldRow.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        if (redrawsSteppers) redrawSteppersIn(inputField.fieldRow);
    }

    /**
     * 「項目名＋数値入力欄＋単位」の行を生成する。項目名と helpTip は fieldLabel.<key> / tooltip.<key> から引く
     * @param {Panel|Group} parentContainer - 追加先
     * @param {{key: string, initialText: string, unitLabel: string, fieldUnit: string, allowNegative: boolean, integer: boolean, min: number, checkboxKey: string, checked: boolean}} fieldSpec - 行の定義
     *     （integer は整数の欄、min は下限。負の値を許さない欄の下限は省略時0。
     *     checkboxKey を指定すると項目名をチェックボックスにし、OFFのあいだ入力欄を無効にする。checked はその初期状態。
     *     fieldUnit（" mm" など）を指定すると単位を入力欄の中に入れ、横の単位は置かない。別の単位で入れた値も換算する）
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {EditText} 生成した入力欄
     */
    function addNumericFieldRow(parentContainer, fieldSpec, onValueChange) {
        var fieldRow = parentContainer.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        fieldRow.spacing = FIELD_ROW_SPACING;

        var optionCheckbox = null;
        var fieldLabel;
        if (fieldSpec.checkboxKey) {
            optionCheckbox = fieldRow.add("checkbox", undefined, getLabel("checkbox." + fieldSpec.checkboxKey));
            optionCheckbox.helpTip = getLabel("tooltip." + fieldSpec.checkboxKey);
            fieldLabel = optionCheckbox;
        } else {
            fieldLabel = fieldRow.add("statictext", undefined, labelText("fieldLabel." + fieldSpec.key));
            fieldLabel.justify = "right";
        }

        /* チェックボックスの行は、入力欄と単位をまとめて有効／無効にする / the checkbox toggles the field and unit together */
        var valueGroup = fieldRow;
        if (optionCheckbox) {
            valueGroup = fieldRow.add("group");
            valueGroup.orientation = "row";
            valueGroup.alignChildren = ["left", "center"];
            valueGroup.spacing = FIELD_ROW_SPACING;
            valueGroup.margins = 0;
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = valueGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var inputField;
        var stepOptions = {
            integer: fieldSpec.integer === true,
            unit: fieldSpec.fieldUnit, /* 単位入りの欄は増減後も単位を付ける / keep the unit in the field after stepping */
            /* text の代入では onChanging が発火しないので、増減後に変更を通知する / assigning text does not fire onChanging */
            onStep: function () { onValueChange(); }
        };
        if (fieldSpec.allowNegative !== true) stepOptions.min = (fieldSpec.min !== undefined) ? fieldSpec.min : 0;
        var stepperGroup = addStepper(stepperInputGroup, function () { return inputField; }, stepOptions);

        inputField = stepperInputGroup.add("edittext", undefined, fieldSpec.initialText + (fieldSpec.fieldUnit || ""));
        inputField.characters = fieldSpec.fieldUnit ? UNIT_FIELD_CHARS : FIELD_CHARS;
        inputField.fieldUnit = fieldSpec.fieldUnit;
        inputField.helpTip = getLabel("tooltip." + fieldSpec.key);
        inputField.onChanging = onValueChange;
        bindSteppedArrowKeys(inputField, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */

        if (!fieldSpec.fieldUnit) inputField.unitText = valueGroup.add("statictext", undefined, fieldSpec.unitLabel);
        inputField.fieldRow = fieldRow;
        inputField.fieldLabel = fieldLabel; /* 項目名の幅をあとで実測してそろえる / label widths are aligned after measuring */

        if (optionCheckbox) {
            /* 表示前の value は読み戻せないので、初期状態は fieldSpec.checked から決める / value cannot be read back before show() */
            optionCheckbox.value = fieldSpec.checked === true;
            valueGroup.enabled = fieldSpec.checked === true;
            optionCheckbox.onClick = function () {
                valueGroup.enabled = optionCheckbox.value;
                redrawSteppersIn(valueGroup);
                onValueChange();
            };
            inputField.optionCheckbox = optionCheckbox;
        }
        return inputField;
    }

    /**
     * 複数のラベル（statictext）の幅を最長のものへ揃え、指定方向に揃えます。
     * 実際の描画幅（preferredSize.width）を測るため、文言長やロケールに依存しません。
     * @param {StaticText[]} labelControls - 幅を揃える statictext の配列。
     * @param {string} [justify] - 揃え方向 "left" | "center" | "right"（省略時は "right"）。
     * @returns {number} 揃えた後の共通ラベル幅（px）。
     */
    function alignLabelWidths(labelControls, justify) {
        if (!labelControls || !labelControls.length) return 0;
        justify = justify || 'right';

        /* 1パス目：各ラベルの自然幅を測り、最大値を求める / Pass 1: find the widest natural width */
        var maxLabelWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            /* preferredSize を返さない環境では実サイズで代用する / fall back to the laid-out size */
            var naturalWidth = labelControls[i].preferredSize.width || labelControls[i].size.width;
            if (naturalWidth > maxLabelWidth) maxLabelWidth = naturalWidth;
        }

        /* 2パス目：全ラベルを最大幅に固定し、揃え方向を適用 / Pass 2: apply common width and justification */
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxLabelWidth;
            labelControls[j].justify = justify;
        }

        return maxLabelWidth;
    }

    /**
     * 入力欄の列にそろえてチェックボックスの行を作る。表示名と helpTip は checkbox.<key> / tooltip.<key> から引く
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} checkboxKey - ラベルのキー
     * @param {boolean} initialValue - 初期状態
     * @param {Function} onValueChange - 状態が変わったときに呼ぶ関数
     * @returns {Checkbox} 生成したチェックボックス
     */
    function addCheckboxRow(parentContainer, checkboxKey, initialValue, onValueChange) {
        var checkboxRow = parentContainer.add("group");
        checkboxRow.orientation = "row";
        checkboxRow.alignChildren = ["left", "center"];
        checkboxRow.spacing = FIELD_ROW_SPACING;

        /* 項目名の欄を空けて入力欄の左端にそろえる / leave the label column empty to line up with the fields */
        var labelSpacer = checkboxRow.add("statictext", undefined, "");

        var optionCheckbox = checkboxRow.add("checkbox", undefined, getLabel("checkbox." + checkboxKey));
        optionCheckbox.value = initialValue;
        optionCheckbox.helpTip = getLabel("tooltip." + checkboxKey);
        optionCheckbox.onClick = onValueChange;
        optionCheckbox.labelSpacer = labelSpacer; /* 幅は項目名の実測後に入れる / sized after the labels are measured */
        return optionCheckbox;
    }

    /**
     * pt の値を現在の単位の表示文字列にする（小数第2位まで）
     * @param {number} valuePt - 値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {string} 入力欄に入れる文字列
     */
    function formatUnitValue(valuePt, pointsPerUnit) {
        var unitValue = valuePt / pointsPerUnit;
        return String(Math.round(unitValue * 100) / 100);
    }

    /**
     * ダイアログの入力欄と選択を、buildHeadlineGroup() に渡す設定値にまとめる
     * @param {Object} dialogUI - buildDialog() の戻り値
     * @param {number} pointsPerUnit - 定規の単位の1単位あたりのポイント数
     * @param {number} strokePointsPerUnit - 線の単位の1単位あたりのポイント数
     * @returns {{marginVerticalPt: number, marginHorizontalPt: number, lineGapPt: number, ruleStyle: string, ruleStrokeWidthPt: number, charsPerLine: number, tierCount: number, tierGap: number, extraTierCount: number, dummyScript: string, dummyDirection: string, dummyText: string, columnCount: number, dummyWidthPt: number, blurRadius: number, useTierRules: boolean, tierRuleWidthPt: number, rotationDeg: number, useMask: boolean, englishNewspaper: boolean, maskVerticalPt: number}} マージン・間隔・幅（pt）、罫線の種類と線幅（pt）、和文／欧文と組み方向と文章、文字数・段数・段間（pt）、ぼかし（px）、段間の罫線の有無と線幅（pt）・回転（°）・マスクの有無とマスクの上下（pt）
     */
    function readHeadlineSettings(dialogUI, pointsPerUnit, strokePointsPerUnit) {
        var fieldInputs = dialogUI.fieldInputs;
        return {
            marginVerticalPt: readFieldAsPoints(fieldInputs.marginVertical, DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit),
            marginHorizontalPt: readFieldAsPoints(fieldInputs.marginHorizontal, DEFAULT_HORIZONTAL_MARGIN_PT, pointsPerUnit),
            lineGapPt: readFieldAsPoints(fieldInputs.lineGap, DEFAULT_LINE_GAP_PT, pointsPerUnit),
            ruleStyle: dialogUI.getRuleStyle(),
            ruleStrokeWidthPt: readFieldAsPoints(fieldInputs.ruleStrokeWidth, DEFAULT_RULE_STROKE_WIDTH_PT, strokePointsPerUnit),
            /* 文字数・段数は整数に切り捨て / character and tier counts are truncated to integers */
            charsPerLine: Math.floor(readFieldAsNumber(fieldInputs.charsPerLine, DEFAULT_CHARS_PER_LINE, false)),
            tierCount: Math.floor(readFieldAsNumber(fieldInputs.tierCount, DEFAULT_TIER_COUNT, false)),
            tierGap: readFieldAsPoints(fieldInputs.tierGap, DEFAULT_TIER_GAP_PT, pointsPerUnit),
            extraTierCount: Math.floor(readFieldAsNumber(fieldInputs.extraTierCount, DEFAULT_EXTRA_TIER_COUNT, false)),
            dummyScript: dialogUI.getDummyScript(),
            dummyDirection: dialogUI.getDummyDirection(),
            dummyText: dialogUI.getDummyText(),
            /* ［段数］は横組みでは段組数として使う / the Tiers field doubles as the column count for horizontal */
            columnCount: Math.floor(readFieldAsNumber(fieldInputs.tierCount, DEFAULT_COLUMN_COUNT, false)),
            dummyWidthPt: readFieldAsPoints(fieldInputs.dummyWidth, DEFAULT_DUMMY_WIDTH_PT, pointsPerUnit),
            /* ぼかしがOFFなら半径0（かけない）/ radius 0 (no blur) when the blur option is off */
            blurRadius: fieldInputs.blurRadius.optionCheckbox.value ? readFieldAsNumber(fieldInputs.blurRadius, DEFAULT_BLUR_RADIUS, false) : 0,
            rotationDeg: readFieldAsNumber(fieldInputs.rotation, DEFAULT_ROTATION_DEG, true),
            useTierRules: dialogUI.tierRuleCheckbox.value,
            tierRuleWidthPt: readFieldAsPoints(fieldInputs.tierRuleWidth, DEFAULT_TIER_RULE_WIDTH_PT, strokePointsPerUnit),
            useMask: dialogUI.maskCheckbox.value,
            englishNewspaper: dialogUI.englishNewspaperCheckbox.value,
            maskVerticalPt: readFieldAsPoints(fieldInputs.maskVertical, DEFAULT_MASK_VERTICAL_PT, pointsPerUnit)
        };
    }

    /**
     * 選択した見出しから、ダイアログの初期の選択を決める。
     * すべて欧文なら［欧文］・［横組み］で始め、和文の横組みなら［段数］を1段にする。縦組みなら［欧文］は選べない
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @returns {{dummyScript: string, dummyDirection: string, tierCount: number, allowsWestern: boolean}} 初期の選択
     */
    function getInitialDummyChoices(targetItem) {
        var dummyScript = detectDummyScript(targetItem);
        var isVerticalHeadline = hasVerticalText(targetItem);
        var isHorizontalJapaneseHeadline = dummyScript === "japanese" && !isVerticalHeadline &&
            /\S/.test(collectTextContents(targetItem));
        return {
            dummyScript: dummyScript,
            dummyDirection: (dummyScript === "western") ? "horizontal" : DEFAULT_DUMMY_DIRECTION,
            tierCount: isHorizontalJapaneseHeadline ? HORIZONTAL_HEADLINE_TIER_COUNT : DEFAULT_TIER_COUNT,
            allowsWestern: !isVerticalHeadline
        };
    }

    /**
     * ダイアログを表示し、入力に合わせてプレビューする。
     * プレビューは対象の複製で作り、その間は元の対象を隠しておく。
     * OK で元の対象を使って確定し、キャンセルで元に戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @returns {GroupItem|null} 確定したグループ。キャンセル時は null
     */
    function showDialog(doc, targetItem) {
        var unitInfo = getUnitInfo();
        var pointsPerUnit = unitInfo.pointsPerUnit;
        var strokePointsPerUnit = getUnitInfo("strokeUnits").pointsPerUnit;
        var previewGroup = null;

        /**
         * 前回のプレビューを削除する
         * @returns {void}
         */
        function removePreview() {
            if (previewGroup) {
                previewGroup.remove();
                previewGroup = null;
            }
        }

        /**
         * 前回のプレビューを消して、対象の複製と現在の入力値で作り直す
         * @returns {void}
         */
        function updatePreview() {
            removePreview();
            /* 複製は隠した元の hidden を引き継ぐので表示に戻す / the duplicate inherits hidden from the original */
            var previewSource = targetItem.duplicate();
            previewSource.hidden = false;
            previewGroup = buildHeadlineGroup(doc, previewSource, readHeadlineSettings(dialogUI, pointsPerUnit, strokePointsPerUnit));
            fitViewToPreview();
            app.redraw();
        }

        /**
         * ［画面にフィット］がオンなら、プレビューが指定の割合でウィンドウに収まるよう表示を合わせる
         * @returns {void}
         */
        function fitViewToPreview() {
            var fitViewControls = dialogUI.fitViewControls;
            if (!previewGroup || !fitViewControls.checkbox.value) return;
            FitViewToItems.fit([previewGroup], { doc: doc, fillRatio: fitViewControls.getFillRatio() });
        }

        var initialChoices = getInitialDummyChoices(targetItem);
        var dialogUI = buildDialog(unitInfo.label, pointsPerUnit, updatePreview, initialChoices.dummyScript,
            initialChoices.dummyDirection, initialChoices.tierCount, initialChoices.allowsWestern);
        dialogUI.dialog.onShow = updatePreview;
        dialogUI.fitViewControls.checkbox.onClick = function () {
            dialogUI.fitViewControls.updateEnabled();
            fitViewToPreview();
            app.redraw();
        };
        dialogUI.fitViewControls.percentInput.onChange = function () {
            fitViewToPreview();
            app.redraw();
        };

        prepareDialogWindow(dialogUI.dialog, SCRIPT_NAME);
        var initialViewState = FitViewToItems.captureView(doc); /* キャンセルで戻す表示 / view restored on Cancel */
        targetItem.hidden = true;
        var dialogResult = dialogUI.dialog.show();
        removePreview();
        targetItem.hidden = false;

        if (dialogResult !== 1) {
            FitViewToItems.restoreView(initialViewState, doc);
            targetItem.selected = true;
            app.redraw();
            return null;
        }

        var headlineGroup = buildHeadlineGroup(doc, targetItem, readHeadlineSettings(dialogUI, pointsPerUnit, strokePointsPerUnit));
        doc.selection = [headlineGroup];
        return headlineGroup;
    }

    main();
})();
