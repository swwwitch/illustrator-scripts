#target illustrator
#targetengine "SankeyFlowMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを左端のノードにして、そこから枝分かれする流れ図（サンキー図）を作ります。
ノードはキーオブジェクト（無ければ一番左のオブジェクト）。並べてあるラベルのテキストは動かさず、その位置と文字サイズに帯を合わせます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SankeyFlowMaker.md

### Overview

Uses the selected object as the left node and draws a branching flow diagram (Sankey diagram) from it.
The node is the key object (or the leftmost object); selected label texts stay put, and the bands are fitted to their positions and size.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SankeyFlowMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SankeyFlowMaker";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-08";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SankeyFlowMaker.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SankeyFlowMaker.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログボックスの初期値（2回目からは前回の設定）。長さは pt / Initial dialog values (later runs use the last ones); lengths in pt */
    var DEFAULT_SETTINGS = {
        flowText: [
            "9280 inside sandbox",
            "720 auto-reviewed",
            "  713 approved and continue",
            "  7 denied",
            "    4 continue via safer alternative",
            "    3 stop and ask user"
        ].join("\n"),
        widthScale: "sqrt",      /* 太さの決め方："linear"（値に比例）/ "sqrt"（平方根に比例。小さい値も見える）/ how widths follow values */
        curveLengthPt: 100,      /* カーブ部分の長さ / length of the curved part */
        straightLengthPt: 120,   /* カーブのあとの直線部分の長さ / length of the straight part after the curve */
        branchGapPt: 16,         /* 枝どうしの間隔 / gap between branches */
        minWidthPt: 1.5,         /* 帯の最小の太さ / minimum band width */
        fontSizePt: 10,          /* ラベルの文字サイズ / label font size */
        exitRatio: 70,           /* ノードの高さに対する帯の出口の高さ（%）/ band exit height as a percentage of the node height */
        showValues: true,        /* ラベルに数値を付ける / prefix labels with the value */
        arrowTips: true,         /* 末端を矢印にする / arrow tips at the ends */
        strokeBands: false,      /* 帯を線（太さ＝線幅）で描く / draw bands as strokes (width = stroke weight) */
        moveLabels: false,       /* ラベルのテキストを動かす（オフはテキストに帯を合わせる）/ move label texts (off fits the bands to them) */
        /* 分岐点1（1段目の枝の分岐点）/ Junction 1 (on first-level branches) */
        junction1Shape: "circle",  /* 形："circle"（円）/ "rectangle"（長方形）/ "none"（なし）/ shape */
        junction1SizeMode: "size", /* 幅・高さの意味："size"（形の大きさ）/ "margin"（テキストまわりの余白）/ meaning of width and height */
        junction1WidthPt: 0,       /* 幅（0 は自動）/ width (0 = auto) */
        junction1HeightPt: 0,      /* 高さ（0 は自動）/ height (0 = auto) */
        junction1Linked: true,     /* 幅と高さを連動 / link width and height */
        /* 分岐点2（2段目以降の枝の分岐点）/ Junction 2 (on second-level and deeper branches) */
        junction2Shape: "circle",
        junction2SizeMode: "size",
        junction2WidthPt: 0,
        junction2HeightPt: 0,
        junction2Linked: true,
        /* カラー / Color */
        colorMode: "single",       /* 色の付け方："single"（単色）/ "branch"（1段目の枝ごと）/ coloring mode */
        bandColorHex: "#D9D9D9",   /* 単色のときの帯の色（グレーは CMYK ドキュメントで K だけ）/ band color in single mode (grays become K-only in CMYK) */
        bandOpacity: 60            /* 帯1本ずつの不透明度（%）/ per-band opacity (%) */
    };

    /* 1段目の枝ごとに順に振る帯の色 / Band colors assigned to the first-level branches in turn */
    var BAND_PALETTE = ["#2B8AE6", "#E53935", "#F5B04A", "#1DB9B6", "#BE45C1", "#3E8E3F", "#FF8A1F", "#8FA142", "#FF5F8F", "#4C5BE0"];

    /* 色（K の濃度 %。RGB ドキュメントでは同じ濃さのグレー）/ Colors (K tint %; gray of the same darkness in RGB documents) */
    var LABEL_TINT         = 100;  /* ラベル / labels */
    var OUTLINE_TINT       = 30;   /* 分岐点の円・ノードの角丸長方形の線 / stroke of junction circles and the node box */
    var OUTLINE_STROKE_WIDTH = 0.75; /* その線幅（pt）/ its stroke width (pt) */

    var NODE_HEIGHT_PER_FONT_SIZE = 16;   /* テキストのノードの角丸長方形の高さの初期値（文字サイズの倍数）/ initial node box height for a text node (multiple of the font size) */
    var NODE_BOX_SIDE_PADDING     = 2;    /* ノードの角丸長方形の左右の余白（文字サイズの倍数）/ side padding of the node box (multiple of the font size) */
    var NODE_BOX_CORNER_RADIUS    = 1.5;  /* ノードの角丸の半径（文字サイズの倍数）/ corner radius of the node box (multiple of the font size) */
    var JUNCTION_PADDING          = 0.6;  /* 分岐点の形と分岐点名のテキストの間の余白（文字サイズの倍数）/ padding between the junction shape and its caption (multiple of the font size) */
    var JUNCTION_POSITION_RATIO   = 0.6;  /* 位置の決まらない分岐点を、出発点から子の右端までのどこに置くか / where an unplaced junction sits between the start and its children's ends */

    var GROUP_NAME          = "Sankey Flow"; /* 作ったグループの名前 / name of the created group */
    var BAND_GROUP_NAME     = "Bands";       /* 帯のサブグループ（効果を掛けるグループ）/ band subgroup (carries the effects) */
    var JUNCTION_GROUP_NAME = "Junctions";   /* 分岐点の形のサブグループ / junction shape subgroup */
    var LABEL_GROUP_NAME    = "Labels";      /* 新しく作るラベルのサブグループ / subgroup for new labels */

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

    var FLOW_TEXT_SIZE = [320, 150]; /* フローの入力欄の大きさ / size of the flow text field */
    var LABEL_WIDTH    = 90;         /* 数値欄の項目名の幅 / label width of the numeric fields */
    var FIELD_CHARS    = 6;          /* 数値欄の文字数 / numeric field characters */
    var COLOR_SWATCH_SIZE = [40, 20]; /* 色見本の大きさ / color swatch size */

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
         * @param {object} [rowOptions] - value: チェックの初期値（既定 false）／percent: 割合の初期値／lang: 表示言語
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
            var percentInput = fitViewRow.add("edittext", undefined, String(fitPercent));
            percentInput.characters = 3;
            percentInput.helpTip = getLabel(LABELS.tooltip.fitViewPercent, uiLang);
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

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================

    /* 整列後に「動いていない」とみなす差（pt） / Tolerance for "did not move" after aligning */
    var KEY_DETECT_TOLERANCE_PT = 0.001;

    /**
     * スマートガイドの表示を切り替える（検出の前後で同じ状態に戻す）
     * @returns {void}
     */
    function toggleSmartGuides() {
        /* メニューコマンドが使えない状態がある / The menu command may be unavailable */
        try {
            app.executeMenuCommand("edge");
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] toggleSmartGuides error: " + e);
        }
    }

    /**
     * 控えておいた位置へ戻す（1つ失敗しても残りは戻す）
     * @param {PageItem[]} items - 対象のオブジェクト
     * @param {number[][]} positions - [[left, top], ...]
     * @returns {void}
     */
    function restorePositions(items, positions) {
        for (var i = 0; i < items.length; i++) {
            /* ロック中などで動かせないことがある / Some items may refuse to move (e.g. locked) */
            try {
                items[i].left = positions[i][0];
                items[i].top = positions[i][1];
            } catch (e) {
                $.writeln("[" + SCRIPT_NAME + "] restorePositions error: " + e);
            }
        }
    }

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * DOM にキーオブジェクトを示すプロパティは無いため、整列コマンドを実行して
     * 「どの向きに整列しても動かないもの」を実測で特定する。
     * 判定中は app.redraw() を呼ばない（描画すると整列がそのつど取り消し履歴に積まれる）
     * @param {PageItem[]} items - 選択中のオブジェクト
     * @returns {number} キーオブジェクトの番号。判定できないときは -1
     */
    function detectKeyObjectIndex(items) {
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        var originPositions = [];
        var i;
        for (i = 0; i < items.length; i++) {
            stayedPut.push(true);
            originPositions.push([items[i].left, items[i].top]);
        }

        try {
            for (var commandIndex = 0; commandIndex < alignCommands.length; commandIndex++) {
                /* 2回目以降だけ元の位置へ戻す / Restore only from the second pass on */
                if (commandIndex > 0) restorePositions(items, originPositions);
                app.executeMenuCommand(alignCommands[commandIndex]);
                for (i = 0; i < items.length; i++) {
                    if (!stayedPut[i]) continue;
                    if (Math.abs(items[i].left - originPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                        Math.abs(items[i].top - originPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                        stayedPut[i] = false;
                    }
                }
            }
        } finally {
            /* 例外で抜けるときも整列結果を残さない / Never leave the aligned positions behind */
            restorePositions(items, originPositions);
        }

        var foundIndex = -1;
        for (i = 0; i < items.length; i++) {
            if (!stayedPut[i]) continue;
            if (foundIndex !== -1) return -1; /* 複数残った＝判定不能 / More than one stayed: undecidable */
            foundIndex = i;
        }
        return foundIndex;
    }

    /**
     * キーオブジェクトの番号を返す（2つ以上選択しているときだけ判定する）
     * @param {PageItem[]} items - 選択中のオブジェクト
     * @returns {number} キーオブジェクトの番号。無いときは -1
     */
    function findKeyObjectIndex(items) {
        if (items.length < 2) return -1;
        toggleSmartGuides();
        /* 検出できなくても処理は続ける / Carry on even if detection fails */
        try {
            return detectKeyObjectIndex(items);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] detectKeyObjectIndex error: " + e);
            return -1;
        } finally {
            toggleSmartGuides();
        }
    }

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "サンキー図を作成", en: "Make Sankey Flow" }
        },
        radio: {
            linear:    { ja: "値に比例", en: "Proportional" },
            sqrt:      { ja: "平方根に比例", en: "Square Root" },
            bySize:    { ja: "大きさで指定", en: "By Size" },
            byMargin:  { ja: "余白で指定", en: "By Margin" },
            single:    { ja: "単色", en: "Single Color" },
            branch:    { ja: "1段目の枝ごと", en: "By First-Level Branch" },
            circle:    { ja: "円", en: "Circle" },
            rectangle: { ja: "長方形", en: "Rectangle" },
            none:      { ja: "なし", en: "None" }
        },
        panel: {
            flow:      { ja: "フロー", en: "Flow" },
            size:      { ja: "サイズ", en: "Size" },
            options:   { ja: "オプション", en: "Options" },
            junction1: { ja: "分岐点1", en: "Junction 1" },
            junction2: { ja: "分岐点2", en: "Junction 2" },
            color:     { ja: "カラー", en: "Color" }
        },
        fieldLabel: {
            widthScale:     { ja: "太さ", en: "Width" },
            nodeHeight:     { ja: "ノードの高さ", en: "Node Height" },
            exitRatio:      { ja: "出口の高さ", en: "Exit Height" },
            junctionWidth:  { ja: "幅", en: "Width" },
            junctionHeight: { ja: "高さ", en: "Height" },
            curveLength:    { ja: "カーブの長さ", en: "Curve Length" },
            straightLength: { ja: "直線の長さ", en: "Straight Length" },
            branchGap:      { ja: "枝の間隔", en: "Branch Gap" },
            minWidth:       { ja: "最小の太さ", en: "Min Width" },
            fontSize:       { ja: "文字サイズ", en: "Font Size" },
            bandColor:      { ja: "色", en: "Color" },
            bandOpacity:    { ja: "不透明度", en: "Opacity" }
        },
        note: {
            flowFormat: { ja: "1行に「数値 名前」。字下げで枝分かれ", en: "One \"value name\" per line. Indent to branch" }
        },
        checkbox: {
            showValues: { ja: "ラベルに数値を付ける", en: "Show Values in Labels" },
            arrowTips:  { ja: "末端を矢印に", en: "Arrow Tips" },
            strokeBands: { ja: "帯を線で描く", en: "Draw Bands as Strokes" },
            moveLabels: { ja: "ラベルを動かす", en: "Move Labels" }
        },
        button: {
            reset:  { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "作成", en: "Create" }
        },
        tooltip: {
            reset:          { ja: "フロー以外の設定を初期値に戻します", en: "Resets every setting except the flows" },
            fitView:        { ja: "サンキー図が収まるよう表示倍率を合わせます。キャンセルすると元の表示に戻ります", en: "Zooms so the Sankey diagram fits in the window; Cancel restores the original view" },
            bandColor:      { ja: "クリックしてカラーピッカーで帯の色を選びます", en: "Click to choose the band color in the Color Picker" },
            branchColor:    { ja: "1段目の枝ごとにパレットの色を順に振り、その先の枝は同じ色を引き継ぎます（パレットはスクリプト冒頭の BAND_PALETTE）", en: "Assigns palette colors to the first-level branches in turn; deeper branches inherit them (palette: BAND_PALETTE at the top of the script)" },
            bandOpacity:    { ja: "帯1本ずつの不透明度。重なった帯が透けて見えます", en: "Opacity of each band; overlapping bands show through" },
            flowText: {
                ja: "数値は行頭か行末に（「9280 inside sandbox」「inside sandbox, 9280」）。\n字下げ（スペース）を深くした行が、直前の浅い行の枝になります。\n# で始まる行は無視します",
                en: "Put the value at the start or end of the line (\"9280 inside sandbox\", \"inside sandbox, 9280\").\nA line indented deeper (with spaces) branches from the shallower line above it.\nLines starting with # are ignored"
            },
            widthScale:     { ja: "平方根に比例にすると、小さい値の帯も太めに見えます（分岐の前後で太さの合計は一致しません）", en: "Square Root keeps small values visible (widths no longer add up at junctions)" },
            curveLength:    { ja: "枝分かれのカーブ部分の横の長さ", en: "Horizontal length of the curved part" },
            straightLength: { ja: "カーブのあとの直線部分の長さ", en: "Length of the straight part after the curve" },
            branchGap:      { ja: "隣り合う枝の間隔", en: "Gap between neighboring branches" },
            fontSizeFromText: { ja: "選択中のラベルの文字サイズを使います", en: "Uses the size of the selected labels" },
            nodeHeight:     { ja: "テキストのノードの背面に作る角丸長方形の高さ（図形のノードはその高さで固定）", en: "Height of the rounded box behind a text node (fixed to the shape's height for a shape node)" },
            junctionWidth:  { ja: "大きさで指定：形の幅。余白で指定：テキストの左右の余白（どちらも 0 は自動）", en: "By Size: shape width. By Margin: margin left and right of the text (0 = auto for both)" },
            junctionHeight: { ja: "大きさで指定：形の高さ。余白で指定：テキストの上下の余白（どちらも 0 は自動）", en: "By Size: shape height. By Margin: margin above and below the text (0 = auto for both)" },
            sizeMode:       { ja: "幅・高さを、形の大きさで指定するか、分岐点名のテキストまわりの余白で指定するか", en: "Whether width and height set the shape size or the margin around the junction caption" },
            strokeBands:    { ja: "帯を塗りの形ではなく、中心を通る線（線幅＝帯の太さ）で描きます。矢印は別の三角形になります", en: "Draws each band as a stroke along its center (stroke weight = band width) instead of a filled shape; arrow tips become separate triangles" },
            moveLabels:     { ja: "オン：帯を設定の長さと間隔で配置し、選択中のラベルのテキストをその位置へ移します。オフ：テキストは動かさず、帯をテキストに合わせます", en: "On: lays out the bands by the set lengths and gaps and moves the selected label texts to fit. Off: the texts stay put and the bands are fitted to them" },
            linkJunctionSize: { ja: "幅と高さを連動", en: "Link width and height" },
            junctionShape:  { ja: "分岐点名のテキストがあれば、テキストは動かさず、それが収まる大きさで囲みます", en: "With a junction caption, the text stays put and the shape is sized to hold it" },
            junction1:      { ja: "1段目の枝の先の分岐点", en: "Junctions at the end of first-level branches" },
            junction2:      { ja: "2段目以降の枝の先の分岐点", en: "Junctions at the end of second-level and deeper branches" },
            exitRatio:      { ja: "ノードの高さに対する、帯の出口の高さ（1段目の帯の太さの合計）。ノードの大きさは変わりません", en: "Band exit height (total width of the first-level bands) relative to the node height; the node itself keeps its size" },
            minWidth:       { ja: "これより細くなる帯は、この太さにします", en: "Bands thinner than this are drawn at this width" },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "ノードにするオブジェクトを選択してください。", en: "Select the object to use as the node." },
            noFlows:     { ja: "フローを入力してください。", en: "Enter the flows." },
            noValue:     { ja: "%1 行目に数値がありません :\n%2", en: "Line %1 has no value:\n%2" },
            zeroTotal:   { ja: "1段目の数値の合計が 0 です。", en: "The first-level values add up to 0." }
        }
    };

    // =========================================
    // 本体 / Main
    // =========================================

    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");

    /**
     * 選択中のオブジェクトをノードにして、ダイアログボックスの設定でサンキー図を作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters returns a TextRange, which has no [0] */
        if (!selectedItems || selectedItems.typename === "TextRange" || selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var nodeInfo = getNodeInfo(selectedItems);
        var savedSettings = settingsStore.load(DEFAULT_SETTINGS);
        var settings = showFlowDialog(doc, nodeInfo, savedSettings);
        if (!settings) return;
        /* ラベルから取った文字サイズは覚えない（ノードの高さは既定値に無いので保存されない）
           Do not remember a size taken from the labels (the node height is not in the defaults, so it is not saved) */
        if (nodeInfo.labelTexts.length > 0) settings.fontSizePt = savedSettings.fontSizePt;
        settingsStore.save(settings);
    }

    /**
     * ノードとラベルを選択から振り分ける。ノードはキーオブジェクト、無ければ一番左のオブジェクト。
     * ノード以外の選択中のテキストはラベルの候補にする
     * @param {Array} selectedItems - 選択中のオブジェクト（前面から順）
     * @returns {{left: number, top: number, right: number, bottom: number, isTextNode: boolean, nodeItem: PageItem, labelTexts: TextFrame[], labelPositions: number[][], labelSize: number, backItem: PageItem}} ノードの外形・ノードがテキストか・ラベルの候補と元の位置・文字サイズ（pt）・最背面のオブジェクト
     */
    function getNodeInfo(selectedItems) {
        var nodeItem = null;
        var keyIndex = findKeyObjectIndex(selectedItems);
        if (keyIndex >= 0) {
            nodeItem = selectedItems[keyIndex];
        } else {
            nodeItem = selectedItems[0];
            for (var i = 1; i < selectedItems.length; i++) {
                if (selectedItems[i].geometricBounds[0] < nodeItem.geometricBounds[0]) nodeItem = selectedItems[i];
            }
        }

        var labelTexts = [];
        var labelPositions = [];
        for (var j = 0; j < selectedItems.length; j++) {
            var selectedItem = selectedItems[j];
            if (selectedItem === nodeItem || selectedItem.typename !== "TextFrame" || selectedItem.contents === "") continue;
            labelTexts.push(selectedItem);
            labelPositions.push([selectedItem.position[0], selectedItem.position[1]]);
        }

        var nodeBounds = nodeItem.geometricBounds;
        return {
            left: nodeBounds[0], top: nodeBounds[1], right: nodeBounds[2], bottom: nodeBounds[3],
            isTextNode: nodeItem.typename === "TextFrame",
            nodeItem: nodeItem,
            labelTexts: labelTexts,
            labelPositions: labelPositions,
            labelSize: (labelTexts.length > 0) ? labelTexts[0].textRange.characters[0].characterAttributes.size : 0,
            /* 選択は前面から順なので、最後が最背面 / The selection runs front to back, so the last one is rearmost */
            backItem: selectedItems[selectedItems.length - 1]
        };
    }

    // =========================================
    // フローの読み取り / Flow parsing
    // =========================================

    /**
     * フローの文字列を木構造にする。字下げが深い行は、直前の浅い行の子になる
     * @param {string} flowText - 入力された文字列
     * @returns {{children: Object[], errorLine: number, errorText: string}} 1段目の枝（{ name, value, children }）。数値の無い行があれば errorLine にその行番号（1始まり、無ければ 0）
     */
    function parseFlowText(flowText) {
        var rootNode = { children: [] };
        var parentStack = [{ indent: -1, node: rootNode }];
        var lines = String(flowText).split(/\r\n|\r|\n/);
        for (var i = 0; i < lines.length; i++) {
            var lineText = lines[i];
            var trimmedText = lineText.replace(/^[\s　]+|[\s　]+$/g, "");
            if (trimmedText === "" || trimmedText.charAt(0) === "#") continue;

            var flowEntry = parseFlowLine(trimmedText);
            if (!flowEntry) return { children: [], errorLine: i + 1, errorText: trimmedText };

            var indent = measureIndent(lineText);
            while (parentStack[parentStack.length - 1].indent >= indent) parentStack.pop();
            flowEntry.children = [];
            parentStack[parentStack.length - 1].node.children.push(flowEntry);
            parentStack.push({ indent: indent, node: flowEntry });
        }
        return { children: rootNode.children, errorLine: 0, errorText: "" };
    }

    /**
     * 1行から数値と名前を取り出す（数値は行頭か行末。桁区切りのカンマ可）
     * @param {string} trimmedText - 前後の空白を除いた1行
     * @returns {{name: string, value: number}|null} 数値が無ければ null
     */
    function parseFlowLine(trimmedText) {
        var leadingMatch = trimmedText.match(/^(\d[\d,]*(?:\.\d+)?)(?:[\s　,:：]+(.*))?$/);
        var trailingMatch = trimmedText.match(/^(.*?)[\s　,:：]+(\d[\d,]*(?:\.\d+)?)$/);
        var valueText;
        var name;
        if (leadingMatch) {
            valueText = leadingMatch[1];
            name = leadingMatch[2] || "";
        } else if (trailingMatch) {
            valueText = trailingMatch[2];
            name = trailingMatch[1];
        } else {
            return null;
        }
        var value = Number(valueText.replace(/,/g, ""));
        if (isNaN(value)) return null;
        return { name: name, value: value };
    }

    /**
     * 行頭の字下げの深さを返す（タブは4、全角スペースは2として数える）
     * @param {string} lineText - 1行
     * @returns {number} 字下げの深さ
     */
    function measureIndent(lineText) {
        var indent = 0;
        for (var i = 0; i < lineText.length; i++) {
            var indentChar = lineText.charAt(i);
            if (indentChar === " ") indent += 1;
            else if (indentChar === "\t") indent += 4;
            else if (indentChar === "　") indent += 2;
            else break;
        }
        return indent;
    }

    /**
     * 数値を桁区切りのカンマ付きで返す（9280 → 9,280）
     * @param {number} value - 数値
     * @returns {string} 桁区切りの文字列
     */
    function formatNumber(value) {
        var parts = String(value).split(".");
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        return parts.join(".");
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * サンキー図を描いてグループにまとめ、選択の背面に置く。
     * 選択中のテキストがフローの項目と一致すれば、そのテキストは動かさず、位置に帯を合わせる
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} nodeInfo - getNodeInfo() の結果
     * @param {Object[]} flowRoots - parseFlowText() の children
     * @param {Object} settings - 描画の設定（DEFAULT_SETTINGS に nodeHeightPt を足した形）
     * @returns {GroupItem} 作ったグループ
     */
    function drawSankeyFlow(doc, nodeInfo, flowRoots, settings) {
        /* ［ラベルを動かす］で動かしたテキストは、描くたびに元の位置から始める / Texts moved by "Move Labels" start from their original positions on every draw */
        restoreLabelPositions(nodeInfo);
        var flowGroup = nodeInfo.backItem.layer.groupItems.add();
        flowGroup.name = GROUP_NAME;
        flowGroup.move(nodeInfo.backItem, ElementPlacement.PLACEAFTER);
        /* 帯のグループはレイヤー直下に作り、効果を掛けてから図のグループへ移す（入れ子のグループを選んでメニューコマンドを実行すると、
           効果は一番外側のグループに付くため）。円とラベルはあとで足して前面に置く
           The band group starts at layer level and moves into the diagram group after its effects are applied
           (a menu command run on a nested group applies to the outermost group). Circles and labels go in front */
        var bandGroup = nodeInfo.backItem.layer.groupItems.add();
        bandGroup.name = BAND_GROUP_NAME;
        var overlayGroup = flowGroup.groupItems.add();
        overlayGroup.name = JUNCTION_GROUP_NAME;
        var labelGroup = flowGroup.groupItems.add(); /* 分岐点の形より前面 / in front of the junction shapes */
        labelGroup.name = LABEL_GROUP_NAME;

        var scaledTotal = 0;
        for (var i = 0; i < flowRoots.length; i++) scaledTotal += scaleFlowValue(flowRoots[i].value, settings.widthScale);
        /* ノードの高さ（テキストのノードは角丸長方形の高さ）の「出口の高さ」％を帯の出口にする
           The band exit is the set percentage of the node height (the box height for a text node) */
        var nodeHeight = getNodeHeight(nodeInfo, settings);
        var exitRatio = Math.min(Math.max(settings.exitRatio, 1), 100) / 100;

        var drawContext = {
            doc: doc,
            settings: settings,
            bandGroup: bandGroup,
            overlayGroup: overlayGroup,
            labelGroup: labelGroup,
            moveLabels: settings.moveLabels,                /* ラベルのテキストを動かす / move the label texts */
            /* 1段目の帯の合計が出口の高さになる倍率 / scale that makes the first-level bands add up to the exit height */
            scale: nodeHeight * exitRatio / scaledTotal,
            insideLabelMinWidth: settings.fontSizePt * 1.6, /* ラベルを帯の中に入れる最小の太さ / min band width for an inside label */
            headSize: settings.fontSizePt * 0.7,            /* 細い帯の矢印の長さの目安 / reference arrow length for thin bands */
            templateText: null,                             /* 新しく作るラベルの手本 / template for new labels */
            insideEndX: null,                               /* 中にラベルがある末端の帯の右端 / right end of leaf bands with inside labels */
            labelColor: makeGray(doc, LABEL_TINT),
            whiteColor: makeGray(doc, 0),
            outlineColor: makeGray(doc, OUTLINE_TINT)
        };
        measureFlowTree(flowRoots, drawContext, 1);
        assignBandColors(flowRoots, doc, settings);
        assignLabelTexts(flowRoots, nodeInfo.labelTexts, drawContext);
        if (!settings.moveLabels) drawContext.insideEndX = findInsideEndX(flowRoots, settings.fontSizePt);

        var nodeCenterX = (nodeInfo.left + nodeInfo.right) / 2;
        var nodeCenterY = (nodeInfo.top + nodeInfo.bottom) / 2;
        var startX = nodeInfo.right;
        /* テキストのノードは背面に角丸長方形を作り、帯はその右端から出す / A text node gets a rounded box behind it; bands leave from its right edge */
        if (nodeInfo.isTextNode) startX = drawNodeBox(flowGroup, nodeInfo, nodeHeight, drawContext);
        drawBranches(flowRoots, startX, nodeCenterY, nodeCenterX, drawContext);
        applyBandEffects(bandGroup, doc);
        bandGroup.move(flowGroup, ElementPlacement.PLACEATEND); /* 図のグループの最背面へ / to the back of the diagram group */
        return flowGroup;
    }

    /**
     * 帯のグループに［パスのアウトライン］と［パスファインダー（合流）］の効果を掛ける。
     * パスのアウトラインはメニューコマンドで、合流は XML で掛ける（XML は「内容」の下に入る。メニューコマンドの合流は上に入る）。
     * 後から掛けた効果ほどアピアランスパネルの下に並ぶ。メニューコマンドは選択に効くので、
     * 入れ子でない（レイヤー直下の）帯のグループを選択したまま掛け、最後に選択を元に戻す
     * @param {GroupItem} bandGroup - 帯のグループ
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function applyBandEffects(bandGroup, doc) {
        var savedSelection = [];
        for (var i = 0; i < doc.selection.length; i++) savedSelection.push(doc.selection[i]);
        /* isPre を付けないと「内容」の下に入る / Without isPre it lands below "Contents" */
        var mergeParams = [
            "I Command 8",              /* 8 = 合流 / Merge */
            "B ConvertCustom 1",
            "R Precision 10",           /* 精度（マイクロメートル）/ precision (micrometers) */
            "B RemovePoints 1"          /* 余分なポイントを削除 / remove redundant points */
        ].join(" ");
        try {
            doc.selection = null;
            bandGroup.selected = true;
            app.redraw(); /* executeMenuCommand は直前の DOM 変更が反映されていないと空振りする / the command misfires without a redraw */
            app.executeMenuCommand("Live Outline Stroke");
            app.redraw(); /* パスのアウトラインを確定させてから合流を足す / commit Outline Stroke before adding Merge */
            bandGroup.applyEffect('<LiveEffect name="Adobe Pathfinder"><Dict data="' + mergeParams + ' "/></LiveEffect>');
            app.redraw();
        } finally {
            doc.selection = null;
            for (var j = 0; j < savedSelection.length; j++) savedSelection[j].selected = true;
        }
    }

    /**
     * ノードの高さを返す。テキストのノードは設定の高さ（テキストが収まらなければ広げる）、図形のノードはその高さ
     * @param {Object} nodeInfo - getNodeInfo() の結果
     * @param {Object} settings - 描画の設定
     * @returns {number} ノードの高さ（pt）
     */
    function getNodeHeight(nodeInfo, settings) {
        var itemHeight = nodeInfo.top - nodeInfo.bottom;
        if (!nodeInfo.isTextNode) return itemHeight;
        return Math.max(settings.nodeHeightPt, itemHeight + settings.fontSizePt * 2);
    }

    /**
     * ラベルの候補を元の位置へ戻す（［ラベルを動かす］のプレビューのたびと、キャンセルのときに呼ぶ）
     * @param {Object} nodeInfo - getNodeInfo() の結果
     * @returns {void}
     */
    function restoreLabelPositions(nodeInfo) {
        for (var i = 0; i < nodeInfo.labelTexts.length; i++) {
            var originalPosition = nodeInfo.labelPositions[i];
            var currentPosition = nodeInfo.labelTexts[i].position;
            if (currentPosition[0] === originalPosition[0] && currentPosition[1] === originalPosition[1]) continue;
            nodeInfo.labelTexts[i].position = originalPosition;
        }
    }

    /**
     * テキストのノードの背面に角丸長方形を置く（幅はテキストの左右に余白）
     * @param {GroupItem} flowGroup - 図のグループ
     * @param {Object} nodeInfo - getNodeInfo() の結果
     * @param {number} boxHeight - 角丸長方形の高さ（getNodeHeight() の結果）
     * @param {Object} drawContext - 描画の共通情報
     * @returns {number} 角丸長方形の右端の X
     */
    function drawNodeBox(flowGroup, nodeInfo, boxHeight, drawContext) {
        var fontSize = drawContext.settings.fontSizePt;
        var boxWidth = (nodeInfo.right - nodeInfo.left) + fontSize * NODE_BOX_SIDE_PADDING * 2;
        var boxLeft = (nodeInfo.left + nodeInfo.right) / 2 - boxWidth / 2;
        var boxTop = (nodeInfo.top + nodeInfo.bottom) / 2 + boxHeight / 2;
        var cornerRadius = fontSize * NODE_BOX_CORNER_RADIUS;
        /* あとから足すので帯より前面になる / Added last, so it sits in front of the bands */
        var nodeBox = flowGroup.pathItems.roundedRectangle(boxTop, boxLeft, boxWidth, boxHeight, cornerRadius, cornerRadius);
        nodeBox.filled = true;
        nodeBox.fillColor = drawContext.whiteColor;
        nodeBox.stroked = true;
        nodeBox.strokeColor = drawContext.outlineColor;
        nodeBox.strokeWidth = OUTLINE_STROKE_WIDTH;
        return boxLeft + boxWidth;
    }

    /**
     * 各枝の深さ・太さと、子孫を含めて縦に占める高さを求めて書き込む（depth / width / hasInsideLabel / extent）
     * @param {Object[]} flowNodes - 枝の配列
     * @param {Object} drawContext - 描画の共通情報
     * @param {number} depth - 枝の段（1段目が 1）
     * @returns {void}
     */
    function measureFlowTree(flowNodes, drawContext, depth) {
        var settings = drawContext.settings;
        for (var i = 0; i < flowNodes.length; i++) {
            var flowNode = flowNodes[i];
            flowNode.depth = depth;
            var scaledWidth = scaleFlowValue(flowNode.value, settings.widthScale) * drawContext.scale;
            flowNode.width = Math.max(scaledWidth, settings.minWidthPt);
            flowNode.hasInsideLabel = flowNode.width >= drawContext.insideLabelMinWidth;
            /* 外に出すラベルの分の高さも取る / Reserve room for a label drawn outside the band */
            flowNode.extent = flowNode.hasInsideLabel ? flowNode.width : Math.max(flowNode.width, settings.fontSizePt * 1.6);
            if (flowNode.children.length === 0) continue;

            measureFlowTree(flowNode.children, drawContext, depth + 1);
            var childrenExtent = settings.branchGapPt * (flowNode.children.length - 1);
            for (var j = 0; j < flowNode.children.length; j++) childrenExtent += flowNode.children[j].extent;
            flowNode.extent = Math.max(flowNode.extent, childrenExtent);
        }
    }

    /**
     * 各枝の帯の色を決めて書き込む（bandColor）。単色ならすべて同じ色、
     * 1段目の枝ごとならパレットの色を順に振り、その先の枝は同じ色を引き継ぐ
     * @param {Object[]} flowRoots - 1段目の枝
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} settings - 描画の設定
     * @returns {void}
     */
    function assignBandColors(flowRoots, doc, settings) {
        var singleColor = hexToDocumentColor(doc, settings.bandColorHex);
        for (var i = 0; i < flowRoots.length; i++) {
            var branchColor = (settings.colorMode === "branch") ? hexToDocumentColor(doc, BAND_PALETTE[i % BAND_PALETTE.length]) : singleColor;
            var branchNodes = flattenFlowTree([flowRoots[i]], []);
            for (var j = 0; j < branchNodes.length; j++) branchNodes[j].bandColor = branchColor;
        }
    }

    /**
     * #RRGGBB をドキュメントのカラーモードの色にする。グレー（R=G=B）は K だけのグレー、
     * CMYK ドキュメントのそれ以外の色はカラー設定で変換する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} hexText - #RRGGBB
     * @returns {RGBColor|CMYKColor} 色（形式が違えば黒）
     */
    function hexToDocumentColor(doc, hexText) {
        var rgbValues = parseHexColor(hexText) || [0, 0, 0];
        if (rgbValues[0] === rgbValues[1] && rgbValues[1] === rgbValues[2]) return makeGray(doc, Math.round((1 - rgbValues[0] / 255) * 100));
        if (doc.documentColorSpace !== DocumentColorSpace.CMYK) {
            var rgbColor = new RGBColor();
            rgbColor.red = rgbValues[0];
            rgbColor.green = rgbValues[1];
            rgbColor.blue = rgbValues[2];
            return rgbColor;
        }
        var cmykValues = app.convertSampleColor(ImageColorSpace.RGB, rgbValues, ImageColorSpace.CMYK, ColorConvertPurpose.defaultpurpose);
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cmykValues[0];
        cmykColor.magenta = cmykValues[1];
        cmykColor.yellow = cmykValues[2];
        cmykColor.black = cmykValues[3];
        return cmykColor;
    }

    /**
     * #RRGGBB（# は省略可）を [R, G, B] にする
     * @param {string} hexText - HEX
     * @returns {number[]|null} 0〜255 の配列。形式が違えば null
     */
    function parseHexColor(hexText) {
        var hexMatch = String(hexText || "").replace(/^\s+|\s+$/g, "").match(/^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/);
        if (!hexMatch) return null;
        return [parseInt(hexMatch[1], 16), parseInt(hexMatch[2], 16), parseInt(hexMatch[3], 16)];
    }

    /**
     * カラーピッカーが返した色を #RRGGBB にする（CMYK・グレーは RGB に変換）
     * @param {Color} pickedColor - 色
     * @returns {string|null} HEX。変換できない種類なら null
     */
    function colorToHex(pickedColor) {
        var rgbValues;
        if (pickedColor.typename === "RGBColor") {
            rgbValues = [pickedColor.red, pickedColor.green, pickedColor.blue];
        } else if (pickedColor.typename === "CMYKColor") {
            rgbValues = app.convertSampleColor(ImageColorSpace.CMYK,
                [pickedColor.cyan, pickedColor.magenta, pickedColor.yellow, pickedColor.black],
                ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
        } else if (pickedColor.typename === "GrayColor") {
            rgbValues = app.convertSampleColor(ImageColorSpace.GrayScale, [pickedColor.gray],
                ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
        } else {
            return null;
        }
        var hexText = "#";
        for (var i = 0; i < 3; i++) {
            var channelHex = Math.round(rgbValues[i]).toString(16);
            hexText += (channelHex.length === 1 ? "0" : "") + channelHex;
        }
        return hexText.toUpperCase();
    }

    /**
     * 値を太さの決め方に合わせて変換する（倍率を掛ける前の値）
     * @param {number} value - 枝の値
     * @param {string} widthScale - "linear" / "sqrt"
     * @returns {number} 変換した値
     */
    function scaleFlowValue(value, widthScale) {
        return (widthScale === "sqrt") ? Math.sqrt(Math.max(value, 0)) : value;
    }

    // -----------------------------------------
    // 選択中のテキストの割り当て / Assigning the selected texts
    // -----------------------------------------

    /**
     * 選択中のテキストを枝に割り当てる。中身が「数値 名前」か「名前」と一致すればその枝のラベル（labelFrame）、
     * 数値を含まない残りは、左から順に分岐点の名前（captionFrame）にする
     * @param {Object[]} flowRoots - 1段目の枝
     * @param {TextFrame[]} labelTexts - ラベルの候補
     * @param {Object} drawContext - 描画の共通情報（templateText を書き込む）
     * @returns {void}
     */
    function assignLabelTexts(flowRoots, labelTexts, drawContext) {
        var flowNodes = flattenFlowTree(flowRoots, []);
        var isUsed = [];
        var i;
        var j;
        for (i = 0; i < labelTexts.length; i++) isUsed.push(false);

        for (i = 0; i < flowNodes.length; i++) {
            for (j = 0; j < labelTexts.length; j++) {
                if (isUsed[j] || !isLabelForFlow(labelTexts[j].contents, flowNodes[i])) continue;
                flowNodes[i].labelFrame = labelTexts[j];
                if (!drawContext.templateText) drawContext.templateText = labelTexts[j];
                isUsed[j] = true;
                break;
            }
        }
        if (!drawContext.templateText && labelTexts.length > 0) drawContext.templateText = labelTexts[0];

        /* 自分で置いた途中の枝のラベルは、帯の太さにかかわらず帯のラベルとして扱う（太さの設定で分岐点名に変わらない）。
           ［ラベルを動かす］のときは新しく作るラベルと同じ扱い
           A placed label on a middle branch always stays a band label, so changing the widths never turns it into a caption;
           with "Move Labels" it is treated like a new label */
        for (i = 0; i < flowNodes.length && !drawContext.moveLabels; i++) {
            if (flowNodes[i].labelFrame && flowNodes[i].children.length > 0) flowNodes[i].hasInsideLabel = true;
        }

        /* 数値を含まない残りのテキストを、左にあるものから分岐点へ順に当てる / Remaining texts without digits go to the junctions, leftmost first */
        for (i = 0; i < flowNodes.length; i++) {
            if (flowNodes[i].children.length === 0 || flowNodes[i].captionFrame) continue;
            var captionIndex = -1;
            for (j = 0; j < labelTexts.length; j++) {
                if (isUsed[j] || /\d/.test(labelTexts[j].contents)) continue;
                if (captionIndex < 0 || labelTexts[j].geometricBounds[0] < labelTexts[captionIndex].geometricBounds[0]) captionIndex = j;
            }
            if (captionIndex < 0) break;
            flowNodes[i].captionFrame = labelTexts[captionIndex];
            isUsed[captionIndex] = true;
        }
    }

    /**
     * 枝の先の分岐点の設定を返す（1段目の枝は分岐点1、2段目以降は分岐点2）
     * @param {Object} flowNode - 枝（measureFlowTree() 済み）
     * @param {Object} drawContext - 描画の共通情報
     * @returns {{shape: string, sizeMode: string, widthPt: number, heightPt: number}} 形・幅と高さの意味（"size" / "margin"）・幅・高さ（0 で自動）
     */
    function getJunctionStyle(flowNode, drawContext) {
        var settings = drawContext.settings;
        if (flowNode.depth <= 1) return { shape: settings.junction1Shape, sizeMode: settings.junction1SizeMode, widthPt: settings.junction1WidthPt, heightPt: settings.junction1HeightPt };
        return { shape: settings.junction2Shape, sizeMode: settings.junction2SizeMode, widthPt: settings.junction2WidthPt, heightPt: settings.junction2HeightPt };
    }

    /**
     * 枝のラベルを分岐点の形の中に入れるかどうか（帯の中に入らない途中の枝で、分岐点に形を置くとき）
     * @param {Object} flowNode - 枝
     * @param {Object} drawContext - 描画の共通情報
     * @returns {boolean} 入れるなら true
     */
    function isLabelInJunction(flowNode, drawContext) {
        return flowNode.children.length > 0 && !flowNode.hasInsideLabel && getJunctionStyle(flowNode, drawContext).shape !== "none";
    }

    /**
     * 枝を上から順（深さ優先）に並べた配列を返す
     * @param {Object[]} flowNodes - 枝の配列
     * @param {Object[]} flatNodes - 書き足す配列
     * @returns {Object[]} flatNodes
     */
    function flattenFlowTree(flowNodes, flatNodes) {
        for (var i = 0; i < flowNodes.length; i++) {
            flatNodes.push(flowNodes[i]);
            flattenFlowTree(flowNodes[i].children, flatNodes);
        }
        return flatNodes;
    }

    /**
     * テキストの中身が枝のラベルかどうか（大文字・小文字、空白、桁区切り、記号の違いは無視）
     * @param {string} textContents - テキストの中身
     * @param {Object} flowNode - 枝
     * @returns {boolean} 「数値 名前」か「名前」と一致すれば true
     */
    function isLabelForFlow(textContents, flowNode) {
        var normalizedText = normalizeLabel(textContents);
        if (normalizedText === normalizeLabel(flowNode.value + flowNode.name)) return true;
        return flowNode.name !== "" && normalizedText === normalizeLabel(flowNode.name);
    }

    /**
     * 照合用に文字列をそろえる（小文字にし、空白・改行・桁区切り・記号を除く）
     * @param {string} labelString - 文字列
     * @returns {string} そろえた文字列
     */
    function normalizeLabel(labelString) {
        return String(labelString).toLowerCase().replace(/[\s　,，、.．:：\-－_]/g, "");
    }

    /**
     * 中にラベルがある末端の帯の右端をそろえる位置を返す（一番長いラベルの右に余白）
     * @param {Object[]} flowRoots - 1段目の枝
     * @param {number} fontSize - 文字サイズ（pt）
     * @returns {number|null} 右端の X。該当する帯が無ければ null
     */
    function findInsideEndX(flowRoots, fontSize) {
        var flowNodes = flattenFlowTree(flowRoots, []);
        var endX = null;
        for (var i = 0; i < flowNodes.length; i++) {
            var flowNode = flowNodes[i];
            if (!flowNode.labelFrame || !flowNode.hasInsideLabel || flowNode.children.length > 0) continue;
            var labelRight = flowNode.labelFrame.geometricBounds[2] + fontSize * 1.5;
            if (endX === null || labelRight > endX) endX = labelRight;
        }
        return endX;
    }

    // -----------------------------------------
    // 帯の位置 / Band positions
    // -----------------------------------------

    /**
     * テキストの外形と中心を返す
     * @param {PageItem} pageItem - 対象
     * @returns {{left: number, top: number, right: number, bottom: number, centerX: number, centerY: number}} 外形と中心
     */
    function getItemBox(pageItem) {
        var itemBounds = pageItem.geometricBounds;
        return {
            left: itemBounds[0], top: itemBounds[1], right: itemBounds[2], bottom: itemBounds[3],
            centerX: (itemBounds[0] + itemBounds[2]) / 2,
            centerY: (itemBounds[1] + itemBounds[3]) / 2
        };
    }

    /**
     * 割り当てたテキストから決まる、帯の中心の Y を返す（ラベルの高さの中央）。
     * ラベルが無い分岐する枝は、子の中心の平均にする。［ラベルを動かす］のときは決めない
     * @param {Object} flowNode - 枝
     * @param {Object} drawContext - 描画の共通情報
     * @returns {number|null} 中心の Y。決まらなければ null
     */
    function findAnchoredCenterY(flowNode, drawContext) {
        if (drawContext.moveLabels) return null;
        if (flowNode.labelFrame) return getItemBox(flowNode.labelFrame).centerY;
        var centerSum = 0;
        var centerCount = 0;
        for (var i = 0; i < flowNode.children.length; i++) {
            var childCenterY = findAnchoredCenterY(flowNode.children[i], drawContext);
            if (childCenterY === null) continue;
            centerSum += childCenterY;
            centerCount++;
        }
        return (centerCount > 0) ? centerSum / centerCount : null;
    }

    /**
     * 割り当てたテキストだけで決まる、帯の右端（分岐点）の X を返す（［ラベルを動かす］のときは決めない）
     * @param {Object} flowNode - 枝
     * @param {Object} drawContext - 描画の共通情報
     * @returns {number|null} 右端の X。決まらなければ null
     */
    function findFixedEndX(flowNode, drawContext) {
        if (drawContext.moveLabels) return null;
        if (flowNode.children.length > 0) return flowNode.captionFrame ? getItemBox(flowNode.captionFrame).centerX : null;
        if (!flowNode.labelFrame) return null;
        if (flowNode.hasInsideLabel) return drawContext.insideEndX;
        /* 外のラベルは矢印の先端の右にある / An outside label sits right of the arrow tip */
        return getItemBox(flowNode.labelFrame).left - drawContext.settings.fontSizePt * 0.4 - getTipLength(flowNode, drawContext);
    }

    /**
     * 子孫の帯の右端のうち、一番左のものを返す
     * @param {Object} flowNode - 枝
     * @param {Object} drawContext - 描画の共通情報
     * @returns {number|null} 右端の X。決まらなければ null
     */
    function findDescendantEndX(flowNode, drawContext) {
        var minEndX = null;
        for (var i = 0; i < flowNode.children.length; i++) {
            var childEndX = findFixedEndX(flowNode.children[i], drawContext);
            if (childEndX === null) childEndX = findDescendantEndX(flowNode.children[i], drawContext);
            if (childEndX !== null && (minEndX === null || childEndX < minEndX)) minEndX = childEndX;
        }
        return minEndX;
    }

    /**
     * 末端の矢印の長さを返す
     * @param {Object} flowNode - 枝
     * @param {Object} drawContext - 描画の共通情報
     * @returns {number} 矢印の長さ（矢印にしない枝は 0）
     */
    function getTipLength(flowNode, drawContext) {
        if (!drawContext.settings.arrowTips || flowNode.children.length > 0) return 0;
        /* 細い帯でも先端が見える長さにする（帯の太さを超えない）/ Keep the tip visible on thin bands, but no longer than the band is wide */
        return Math.max(flowNode.width * 0.15, Math.min(flowNode.width, drawContext.headSize) * 0.8);
    }

    /**
     * 1つの分岐点から出る枝を描き、子のある枝はその先でさらに枝分かれさせる。
     * 位置は割り当てたテキストに合わせ、決まらない分は設定の長さと間隔で補う
     * @param {Object[]} flowNodes - 枝の配列（measureFlowTree() 済み）
     * @param {number} startX - 枝の出発点の X
     * @param {number} centerY - 分岐点の中心の Y
     * @param {number} tailX - 帯の根元を延ばす先の X（startX と同じなら延ばさない）
     * @param {Object} drawContext - 描画の共通情報
     * @returns {void}
     */
    function drawBranches(flowNodes, startX, centerY, tailX, drawContext) {
        var settings = drawContext.settings;
        var fontSize = settings.fontSizePt;
        var sourceTotal = 0;
        var targetTotal = settings.branchGapPt * (flowNodes.length - 1);
        for (var i = 0; i < flowNodes.length; i++) {
            sourceTotal += flowNodes[i].width;
            targetTotal += flowNodes[i].extent;
        }

        /* 出発点では帯を隙間なく積み、行き先では高さ＋間隔で振り分ける（どちらも中心をそろえる）
           Bands are stacked tightly at the start and spread by extent + gap at the end, both centered */
        var sourceTop = centerY + sourceTotal / 2;
        var targetTop = centerY + targetTotal / 2;

        for (var j = 0; j < flowNodes.length; j++) {
            var flowNode = flowNodes[j];

            /* 右端：テキストで決まればそこ、子孫の位置が分かれば手前に、どちらも無ければ設定の長さ
               End: from the texts if set, short of the descendants if known, else the set lengths */
            var endX = findFixedEndX(flowNode, drawContext);
            if (endX === null && flowNode.children.length > 0) {
                var descendantEndX = findDescendantEndX(flowNode, drawContext);
                if (descendantEndX !== null) endX = startX + (descendantEndX - startX) * JUNCTION_POSITION_RATIO;
            }
            var isEndFixed = (endX !== null);
            if (!isEndFixed) endX = startX + settings.curveLengthPt + settings.straightLengthPt;
            endX = Math.max(endX, startX);

            /* カーブの終わり：中のラベルの手前。右端が決まっているときは半分までに収める / Curve end: before an inside label; at most halfway when the end is fixed */
            var curveEndX;
            if (flowNode.labelFrame && flowNode.hasInsideLabel && !drawContext.moveLabels) curveEndX = getItemBox(flowNode.labelFrame).left - fontSize;
            else if (isEndFixed) curveEndX = startX + Math.min(settings.curveLengthPt, (endX - startX) / 2);
            else curveEndX = startX + settings.curveLengthPt;
            curveEndX = Math.min(Math.max(curveEndX, startX), endX);

            var anchoredCenterY = findAnchoredCenterY(flowNode, drawContext);
            var targetCenterY = (anchoredCenterY !== null) ? anchoredCenterY : targetTop - flowNode.extent / 2;
            var bandGeometry = {
                tailX: tailX,
                startX: startX,
                curveEndX: curveEndX,
                endX: endX,
                sourceTop: sourceTop,
                sourceBottom: sourceTop - flowNode.width,
                targetTop: targetCenterY + flowNode.width / 2,
                targetBottom: targetCenterY - flowNode.width / 2,
                targetCenterY: targetCenterY
            };
            var tipEndX = drawBand(flowNode, bandGeometry, drawContext);
            placeBranchLabel(flowNode, bandGeometry, tipEndX, drawContext);

            if (flowNode.children.length > 0) {
                /* ［ラベルを動かす］のときは分岐点名を分岐点の中心へ / With "Move Labels", center the caption on the junction */
                if (drawContext.moveLabels && flowNode.captionFrame) centerItemAt(flowNode.captionFrame, endX, targetCenterY);
                if (getJunctionStyle(flowNode, drawContext).shape !== "none") drawJunctionShape(flowNode, endX, targetCenterY, drawContext);
                drawBranches(flowNode.children, endX, targetCenterY, endX, drawContext);
            }
            sourceTop -= flowNode.width;
            targetTop -= flowNode.extent + settings.branchGapPt;
        }
    }

    // -----------------------------------------
    // 図形とラベル / Shapes and labels
    // -----------------------------------------

    /**
     * 帯を1本描く（根元 → S字のカーブ → 直線 → 末端。末端は子が無ければ矢印）
     * @param {Object} flowNode - 枝
     * @param {Object} bandGeometry - 帯の座標（drawBranches() で作る）
     * @param {Object} drawContext - 描画の共通情報
     * @returns {number} 帯の右端の X（矢印なら先端）
     */
    function drawBand(flowNode, bandGeometry, drawContext) {
        if (drawContext.settings.strokeBands) return drawStrokeBand(flowNode, bandGeometry, drawContext);
        var geometry = bandGeometry;
        var handleLength = (geometry.curveEndX - geometry.startX) / 2;
        var bandPoints = [];
        var tipLength = getTipLength(flowNode, drawContext);

        /* 上の辺を左から右へ / Top edge, left to right */
        if (geometry.tailX < geometry.startX) bandPoints.push(cornerPoint(geometry.tailX, geometry.sourceTop));
        bandPoints.push({ anchor: [geometry.startX, geometry.sourceTop], left: [geometry.startX, geometry.sourceTop], right: [geometry.startX + handleLength, geometry.sourceTop] });
        bandPoints.push({ anchor: [geometry.curveEndX, geometry.targetTop], left: [geometry.curveEndX - handleLength, geometry.targetTop], right: [geometry.curveEndX, geometry.targetTop] });
        if (geometry.endX > geometry.curveEndX) bandPoints.push(cornerPoint(geometry.endX, geometry.targetTop));

        /* 末端の矢印（帯の幅のまま尖らせる）/ Arrow tip, pointed at the band's own width */
        if (tipLength > 0) bandPoints.push(cornerPoint(geometry.endX + tipLength, geometry.targetCenterY));

        /* 下の辺を右から左へ / Bottom edge, right to left */
        if (geometry.endX > geometry.curveEndX) bandPoints.push(cornerPoint(geometry.endX, geometry.targetBottom));
        bandPoints.push({ anchor: [geometry.curveEndX, geometry.targetBottom], left: [geometry.curveEndX, geometry.targetBottom], right: [geometry.curveEndX - handleLength, geometry.targetBottom] });
        bandPoints.push({ anchor: [geometry.startX, geometry.sourceBottom], left: [geometry.startX + handleLength, geometry.sourceBottom], right: [geometry.startX, geometry.sourceBottom] });
        if (geometry.tailX < geometry.startX) bandPoints.push(cornerPoint(geometry.tailX, geometry.sourceBottom));

        addFilledBandPath(bandPoints, flowNode, drawContext);
        return geometry.endX + tipLength;
    }

    /**
     * 帯を1本、中心を通る線で描く（線幅＝帯の太さ、線端なし）。矢印は線の右端に続く塗りの三角形で描く
     * @param {Object} flowNode - 枝
     * @param {Object} bandGeometry - 帯の座標（drawBranches() で作る）
     * @param {Object} drawContext - 描画の共通情報
     * @returns {number} 帯の右端の X（矢印なら先端）
     */
    function drawStrokeBand(flowNode, bandGeometry, drawContext) {
        var geometry = bandGeometry;
        var handleLength = (geometry.curveEndX - geometry.startX) / 2;
        var sourceCenterY = geometry.sourceTop - flowNode.width / 2;
        var targetCenterY = geometry.targetCenterY;
        var linePoints = [];
        if (geometry.tailX < geometry.startX) linePoints.push(cornerPoint(geometry.tailX, sourceCenterY));
        linePoints.push({ anchor: [geometry.startX, sourceCenterY], left: [geometry.startX, sourceCenterY], right: [geometry.startX + handleLength, sourceCenterY] });
        linePoints.push({ anchor: [geometry.curveEndX, targetCenterY], left: [geometry.curveEndX - handleLength, targetCenterY], right: [geometry.curveEndX, targetCenterY] });
        if (geometry.endX > geometry.curveEndX) linePoints.push(cornerPoint(geometry.endX, targetCenterY));

        var linePath = addBandPath(linePoints, drawContext);
        linePath.closed = false;
        linePath.filled = false;
        linePath.stroked = true;
        linePath.strokeColor = flowNode.bandColor;
        linePath.strokeWidth = flowNode.width;
        linePath.strokeCap = StrokeCap.BUTTENDCAP;
        linePath.strokeJoin = StrokeJoin.MITERENDJOIN;

        var tipLength = getTipLength(flowNode, drawContext);
        if (tipLength > 0) {
            addFilledBandPath([
                cornerPoint(geometry.endX, geometry.targetTop),
                cornerPoint(geometry.endX + tipLength, targetCenterY),
                cornerPoint(geometry.endX, geometry.targetBottom)
            ], flowNode, drawContext);
        }
        return geometry.endX + tipLength;
    }

    /**
     * 帯のグループにパスを足し、ポイントを並べて不透明度を設定する（塗り・線は呼び出し側で決める）
     * @param {Object[]} bandPoints - ポイントの座標（cornerPoint() と同じ形）
     * @param {Object} drawContext - 描画の共通情報
     * @returns {PathItem} 作ったパス
     */
    function addBandPath(bandPoints, drawContext) {
        var bandPath = drawContext.bandGroup.pathItems.add();
        for (var i = 0; i < bandPoints.length; i++) {
            var pathPoint = bandPath.pathPoints.add();
            pathPoint.anchor = bandPoints[i].anchor;
            pathPoint.leftDirection = bandPoints[i].left;
            pathPoint.rightDirection = bandPoints[i].right;
            pathPoint.pointType = PointType.CORNER;
        }
        bandPath.opacity = Math.min(Math.max(drawContext.settings.bandOpacity, 0), 100);
        return bandPath;
    }

    /**
     * 帯の色で塗った閉じたパスを足す
     * @param {Object[]} bandPoints - ポイントの座標
     * @param {Object} flowNode - 枝
     * @param {Object} drawContext - 描画の共通情報
     * @returns {PathItem} 作ったパス
     */
    function addFilledBandPath(bandPoints, flowNode, drawContext) {
        var bandPath = addBandPath(bandPoints, drawContext);
        bandPath.closed = true;
        bandPath.stroked = false;
        bandPath.filled = true;
        bandPath.fillColor = flowNode.bandColor;
        return bandPath;
    }

    /**
     * 方向線の無いアンカーポイントの座標を返す
     * @param {number} x - X 座標
     * @param {number} y - Y 座標
     * @returns {{anchor: number[], left: number[], right: number[]}} ポイントの座標
     */
    function cornerPoint(x, y) {
        return { anchor: [x, y], left: [x, y], right: [x, y] };
    }

    /**
     * 枝のラベルを置く。動かさないテキストのラベルはそのまま。新しく作るラベルと、［ラベルを動かす］のときのテキストは、
     * 帯の中に入らない途中の枝なら分岐点の中心（分岐点名として形で囲む）、それ以外は placeFlowLabel() の位置へ
     * @param {Object} flowNode - 枝
     * @param {Object} bandGeometry - 帯の座標
     * @param {number} tipEndX - 帯の右端の X
     * @param {Object} drawContext - 描画の共通情報
     * @returns {void}
     */
    function placeBranchLabel(flowNode, bandGeometry, tipEndX, drawContext) {
        if (flowNode.labelFrame && !drawContext.moveLabels) return;
        var labelFrame = flowNode.labelFrame;
        if (!labelFrame) {
            var labelString = buildLabelString(flowNode, drawContext.settings);
            if (labelString === "") return;
            labelFrame = createLabelFrame(labelString, drawContext);
        }
        if (isLabelInJunction(flowNode, drawContext) && !flowNode.captionFrame) {
            centerItemAt(labelFrame, bandGeometry.endX, bandGeometry.targetCenterY);
            flowNode.captionFrame = labelFrame;
            return;
        }
        placeFlowLabel(labelFrame, flowNode, bandGeometry, tipEndX, drawContext.settings.fontSizePt);
    }

    /**
     * 帯に対してラベルを置く。太い帯は帯の中、細い末端は矢印の右、細い途中の帯は帯の下の中央
     * @param {TextFrame} labelFrame - ラベルのテキスト
     * @param {Object} flowNode - 枝
     * @param {Object} bandGeometry - 帯の座標
     * @param {number} tipEndX - 帯の右端の X
     * @param {number} fontSize - 文字サイズ（pt）
     * @returns {void}
     */
    function placeFlowLabel(labelFrame, flowNode, bandGeometry, tipEndX, fontSize) {
        var labelBox = getItemBox(labelFrame);
        var frameWidth = labelBox.right - labelBox.left;
        var frameHeight = labelBox.top - labelBox.bottom;
        var geometry = bandGeometry;
        if (flowNode.hasInsideLabel) {
            moveBoundsTo(labelFrame, geometry.curveEndX + fontSize, geometry.targetCenterY + frameHeight / 2);
        } else if (flowNode.children.length === 0) {
            moveBoundsTo(labelFrame, tipEndX + fontSize * 0.4, geometry.targetCenterY + frameHeight / 2);
        } else {
            moveBoundsTo(labelFrame, (geometry.curveEndX + geometry.endX) / 2 - frameWidth / 2, geometry.targetBottom - fontSize * 0.3);
        }
    }

    /**
     * 外形の左上が指定の位置に来るように動かす（position は外形の左上とずれることがあるので、差分で動かす）
     * @param {PageItem} pageItem - 対象
     * @param {number} left - 外形の左端の X
     * @param {number} top - 外形の上端の Y
     * @returns {void}
     */
    function moveBoundsTo(pageItem, left, top) {
        var itemBounds = pageItem.geometricBounds;
        pageItem.position = [pageItem.position[0] + left - itemBounds[0], pageItem.position[1] + top - itemBounds[1]];
    }

    /**
     * 外形の中心が指定の位置に来るように動かす
     * @param {PageItem} pageItem - 対象
     * @param {number} centerX - 中心の X
     * @param {number} centerY - 中心の Y
     * @returns {void}
     */
    function centerItemAt(pageItem, centerX, centerY) {
        var itemBox = getItemBox(pageItem);
        moveBoundsTo(pageItem, itemBox.left + centerX - itemBox.centerX, itemBox.top + centerY - itemBox.centerY);
    }

    /**
     * 新しく作るラベルの文字列を返す（「9,280 inside sandbox」の形。数値を付けない設定なら名前だけ）
     * @param {Object} flowNode - 枝
     * @param {Object} settings - 描画の設定
     * @returns {string} ラベルの文字列（空のこともある）
     */
    function buildLabelString(flowNode, settings) {
        if (!settings.showValues) return flowNode.name;
        return formatNumber(flowNode.value) + (flowNode.name ? " " + flowNode.name : "");
    }

    /**
     * ラベルのテキストを作る。手本のテキストがあればその書式に合わせる
     * （ポイント文字は複製して中身を差し替え、エリア内・パス上文字はフォント・サイズ・色だけ写す）
     * @param {string} labelString - ラベルの文字列
     * @param {Object} drawContext - 描画の共通情報
     * @returns {TextFrame} 作ったテキスト
     */
    function createLabelFrame(labelString, drawContext) {
        var templateText = drawContext.templateText;
        var labelFrame;
        if (templateText && templateText.kind === TextType.POINTTEXT) {
            labelFrame = templateText.duplicate(drawContext.labelGroup, ElementPlacement.PLACEATEND);
            labelFrame.contents = labelString;
            return labelFrame;
        }

        labelFrame = drawContext.labelGroup.textFrames.add();
        labelFrame.contents = labelString;
        var labelAttributes = labelFrame.textRange.characterAttributes;
        if (templateText) {
            var templateAttributes = templateText.textRange.characters[0].characterAttributes;
            labelAttributes.textFont = templateAttributes.textFont;
            labelAttributes.size = templateAttributes.size;
            labelAttributes.fillColor = templateAttributes.fillColor;
        } else {
            labelAttributes.size = drawContext.settings.fontSizePt;
            labelAttributes.fillColor = drawContext.labelColor;
        }
        return labelFrame;
    }

    /**
     * 分岐点に円か長方形を置く。分岐点名のテキストがあれば、テキストは動かさず、その中心にそろえて収まる大きさで囲む。
     * 無ければ分岐点の中心に、帯の太さより一回り大きく置く
     * @param {Object} flowNode - 分岐する枝
     * @param {number} centerX - 分岐点の X
     * @param {number} centerY - 分岐点の Y
     * @param {Object} drawContext - 描画の共通情報
     * @returns {void}
     */
    function drawJunctionShape(flowNode, centerX, centerY, drawContext) {
        var shapeBox = getJunctionShapeBox(flowNode, centerX, centerY, drawContext);
        var shapeTop = shapeBox.centerY + shapeBox.height / 2;
        var shapeLeft = shapeBox.centerX - shapeBox.width / 2;
        var junctionShape = (getJunctionStyle(flowNode, drawContext).shape === "circle")
            ? drawContext.overlayGroup.pathItems.ellipse(shapeTop, shapeLeft, shapeBox.width, shapeBox.height)
            : drawContext.overlayGroup.pathItems.rectangle(shapeTop, shapeLeft, shapeBox.width, shapeBox.height);
        junctionShape.filled = true;
        junctionShape.fillColor = drawContext.whiteColor;
        junctionShape.stroked = true;
        junctionShape.strokeColor = drawContext.outlineColor;
        junctionShape.strokeWidth = OUTLINE_STROKE_WIDTH;
    }

    /**
     * 分岐点の形の中心と大きさを返す
     * @param {Object} flowNode - 分岐する枝
     * @param {number} centerX - 分岐点の X
     * @param {number} centerY - 分岐点の Y
     * @param {Object} drawContext - 描画の共通情報
     * @returns {{centerX: number, centerY: number, width: number, height: number}} 中心と幅・高さ
     */
    function getJunctionShapeBox(flowNode, centerX, centerY, drawContext) {
        var fontSize = drawContext.settings.fontSizePt;
        var junctionStyle = getJunctionStyle(flowNode, drawContext);
        var isCircle = (junctionStyle.shape === "circle");
        var shapeWidth;
        var shapeHeight;
        var isMarginMode = (junctionStyle.sizeMode === "margin");
        if (flowNode.captionFrame) {
            var captionBox = getItemBox(flowNode.captionFrame);
            centerX = captionBox.centerX;
            centerY = captionBox.centerY;
            /* 余白：余白で指定なら入力値（0 は自動）、大きさで指定なら自動 / Margins: the entered values in margin mode (0 = auto), else auto */
            var autoMargin = fontSize * JUNCTION_PADDING;
            var marginX = (isMarginMode && junctionStyle.widthPt > 0) ? junctionStyle.widthPt : autoMargin;
            var marginY = (isMarginMode && junctionStyle.heightPt > 0) ? junctionStyle.heightPt : autoMargin;
            var boxWidth = (captionBox.right - captionBox.left) + marginX * 2;
            var boxHeight = (captionBox.top - captionBox.bottom) + marginY * 2;
            if (!isCircle) {
                shapeWidth = boxWidth;
                shapeHeight = boxHeight;
            } else if (isMarginMode) {
                /* 余白を足した長方形に外接する楕円 / The ellipse circumscribing the padded rectangle */
                shapeWidth = boxWidth * Math.SQRT2;
                shapeHeight = boxHeight * Math.SQRT2;
            } else {
                /* テキストの外形の対角線に余白を足した直径 / The caption's diagonal plus the margins */
                var captionWidth = captionBox.right - captionBox.left;
                var captionHeight = captionBox.top - captionBox.bottom;
                shapeWidth = shapeHeight = Math.sqrt(captionWidth * captionWidth + captionHeight * captionHeight) + autoMargin * 2;
            }
        } else {
            shapeWidth = shapeHeight = Math.max(flowNode.width + fontSize, fontSize * 2.4);
        }
        /* 大きさで指定なら入力値を使う（0 は自動）/ In size mode use the entered values (0 = auto) */
        if (!isMarginMode && junctionStyle.widthPt > 0) shapeWidth = junctionStyle.widthPt;
        if (!isMarginMode && junctionStyle.heightPt > 0) shapeHeight = junctionStyle.heightPt;
        return { centerX: centerX, centerY: centerY, width: shapeWidth, height: shapeHeight };
    }

    /**
     * ドキュメントのカラーモードに合わせたグレーを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {number} tint - K の濃度（0〜100）
     * @returns {CMYKColor|RGBColor} 色
     */
    function makeGray(doc, tint) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = tint;
            return cmykColor;
        }
        var level = Math.round(255 * (1 - tint / 100));
        var rgbColor = new RGBColor();
        rgbColor.red = level;
        rgbColor.green = level;
        rgbColor.blue = level;
        return rgbColor;
    }

    // =========================================
    // ダイアログボックス / Dialog
    // =========================================

    /**
     * 設定のダイアログボックスを表示し、プレビューしながら図を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} nodeInfo - getNodeInfo() の結果
     * @param {Object} initialSettings - 設定の初期値（DEFAULT_SETTINGS と同じ形）
     * @returns {Object|null} 決定した設定（キャンセル時は null）
     */
    function showFlowDialog(doc, nodeInfo, initialSettings) {
        var lengthUnit = getUnitInfo();
        var textUnit = getUnitInfo("text/units");
        var previewGroup = null;

        var flowDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(flowDialog);

        /* フロー / Flow */
        var flowPanel = flowDialog.add("panel", undefined, getLabel("panel.flow"));
        setupPanel(flowPanel, 6);
        var flowTextField = flowPanel.add("edittext", undefined, initialSettings.flowText, { multiline: true, scrolling: true, wantReturn: true });
        flowTextField.preferredSize = FLOW_TEXT_SIZE;
        flowTextField.helpTip = getLabel("tooltip.flowText");
        var flowFormatNote = flowPanel.add("statictext", undefined, getLabel("note.flowFormat"));
        flowFormatNote.helpTip = getLabel("tooltip.flowText");

        /* サイズ / Size */
        /* 左のカラムにサイズとオプション、右のカラムに分岐点1・2 / Size and Options on the left, Junctions 1 and 2 on the right */
        var settingsColumns = flowDialog.add("group");
        settingsColumns.orientation = "row";
        settingsColumns.alignChildren = ["fill", "fill"];
        settingsColumns.spacing = COLUMN_SPACING;

        var leftColumn = settingsColumns.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = WINDOW_SPACING;
        var sizePanel = leftColumn.add("panel", undefined, getLabel("panel.size"));
        setupPanel(sizePanel, 6);
        var widthScaleKeys = ["linear", "sqrt"];
        var widthScaleRow = sizePanel.add("group");
        setupRow(widthScaleRow, "left", 3);
        widthScaleRow.alignChildren = ["left", "top"]; /* ラベルを1つ目のラジオボタンの高さに / label level with the first radio */
        var widthScaleLabel = widthScaleRow.add("statictext", undefined, labelText("fieldLabel.widthScale"));
        widthScaleLabel.preferredSize.width = LABEL_WIDTH;
        widthScaleLabel.justify = "right";
        widthScaleLabel.helpTip = getLabel("tooltip.widthScale");
        /* ラジオボタンは同じグループに入れて排他にする（縦に並べてパネルの幅を広げない）
           Keep the radios in one group so they are exclusive (stacked so the panel does not widen) */
        var widthScaleColumn = widthScaleRow.add("group");
        widthScaleColumn.orientation = "column";
        widthScaleColumn.alignChildren = ["left", "top"];
        widthScaleColumn.spacing = 4;
        var widthScaleRadios = [];
        for (var w = 0; w < widthScaleKeys.length; w++) {
            var widthScaleRadio = widthScaleColumn.add("radiobutton", undefined, getLabel("radio." + widthScaleKeys[w]));
            widthScaleRadio.helpTip = getLabel("tooltip.widthScale");
            widthScaleRadio.onClick = function () { updatePreview(); };
            widthScaleRadios.push(widthScaleRadio);
        }
        setWidthScale(initialSettings.widthScale);

        /**
         * 太さの決め方のラジオボタンを選ぶ
         * @param {string} scaleKey - "linear" / "sqrt"
         * @returns {void}
         */
        function setWidthScale(scaleKey) {
            for (var i = 0; i < widthScaleKeys.length; i++) widthScaleRadios[i].value = (widthScaleKeys[i] === scaleKey);
            if (indexOfKey(widthScaleKeys, scaleKey) < 0) widthScaleRadios[0].value = true;
        }

        /**
         * 選んでいる太さの決め方を返す
         * @returns {string} "linear" / "sqrt"
         */
        function getWidthScale() {
            for (var i = 0; i < widthScaleKeys.length; i++) {
                if (widthScaleRadios[i].value) return widthScaleKeys[i];
            }
            return widthScaleKeys[0];
        }
        /* ノードの高さは毎回ノードから決める（保存しない）。図形のノードはその高さで固定
           The node height comes from the node every time (not saved); fixed for a shape node */
        var initialFontSize = (nodeInfo.labelTexts.length > 0) ? nodeInfo.labelSize : initialSettings.fontSizePt;
        var initialNodeHeight = nodeInfo.isTextNode ? initialFontSize * NODE_HEIGHT_PER_FONT_SIZE : nodeInfo.top - nodeInfo.bottom;
        var nodeHeightInput = addLengthField(sizePanel, "nodeHeight", initialNodeHeight, lengthUnit, 0.1);
        if (!nodeInfo.isTextNode) setSteppedFieldEnabled(nodeHeightInput, false);
        var exitRatioInput = addSteppedField(sizePanel, {
            label: labelText("fieldLabel.exitRatio"), labelWidth: LABEL_WIDTH,
            text: initialSettings.exitRatio + "%", characters: FIELD_CHARS,
            step: 1, min: 1, max: 100, unit: "%",
            onStep: function () { updatePreview(); }
        });
        exitRatioInput.helpTip = exitRatioInput.fieldLabel.helpTip = getLabel("tooltip.exitRatio");
        var normalizeExitRatio = exitRatioInput.onChange;
        exitRatioInput.onChange = function () {
            normalizeExitRatio();
            updatePreview();
        };

        var curveLengthInput = addLengthField(sizePanel, "curveLength", initialSettings.curveLengthPt, lengthUnit, 0);
        var straightLengthInput = addLengthField(sizePanel, "straightLength", initialSettings.straightLengthPt, lengthUnit, 0);
        var branchGapInput = addLengthField(sizePanel, "branchGap", initialSettings.branchGapPt, lengthUnit, 0);
        var minWidthInput = addLengthField(sizePanel, "minWidth", initialSettings.minWidthPt, lengthUnit, 0);
        /* ラベルのテキストがあれば、その文字サイズに合わせて欄は固定する / With label texts, use their size and lock the field */
        var fontSizeInput = addLengthField(sizePanel, "fontSize", initialFontSize, textUnit, 0.1);
        if (nodeInfo.labelTexts.length > 0) {
            setSteppedFieldEnabled(fontSizeInput, false);
            fontSizeInput.helpTip = fontSizeInput.fieldLabel.helpTip = getLabel("tooltip.fontSizeFromText");
        }

        /* オプション（左のカラムでサイズの下）/ Options (below Size in the left column) */
        var optionsPanel = leftColumn.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 6);
        var chkShowValues = optionsPanel.add("checkbox", undefined, getLabel("checkbox.showValues"));
        var chkArrowTips = optionsPanel.add("checkbox", undefined, getLabel("checkbox.arrowTips"));
        var chkStrokeBands = optionsPanel.add("checkbox", undefined, getLabel("checkbox.strokeBands"));
        chkStrokeBands.helpTip = getLabel("tooltip.strokeBands");
        chkStrokeBands.value = initialSettings.strokeBands;
        var chkMoveLabels = optionsPanel.add("checkbox", undefined, getLabel("checkbox.moveLabels"));
        chkMoveLabels.helpTip = getLabel("tooltip.moveLabels");
        chkMoveLabels.value = initialSettings.moveLabels;
        chkShowValues.value = initialSettings.showValues;
        chkArrowTips.value = initialSettings.arrowTips;

        /* 分岐点1・2（右のカラム）/ Junctions 1 and 2 (right column) */
        var junctionShapeKeys = ["circle", "rectangle", "none"];
        var rightColumn = settingsColumns.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];
        rightColumn.spacing = WINDOW_SPACING;
        var junction1Controls = addJunctionPanel(rightColumn, "junction1", initialSettings.junction1Shape, initialSettings.junction1SizeMode, initialSettings.junction1WidthPt, initialSettings.junction1HeightPt, initialSettings.junction1Linked);
        var junction2Controls = addJunctionPanel(rightColumn, "junction2", initialSettings.junction2Shape, initialSettings.junction2SizeMode, initialSettings.junction2WidthPt, initialSettings.junction2HeightPt, initialSettings.junction2Linked);
        var colorControls = addColorPanel(rightColumn, initialSettings);

        /**
         * カラーのパネル（色の付け方のラジオボタン・色見本・不透明度）を追加する
         * @param {Group} parent - 追加先
         * @param {Object} colorSettings - 初期値（colorMode / bandColorHex / bandOpacity）
         * @returns {{getMode: Function, setMode: Function, getHex: Function, setHex: Function, opacityInput: EditText}} パネルの操作
         */
        function addColorPanel(parent, colorSettings) {
            var colorModeKeys = ["single", "branch"];
            var bandColorHex = colorSettings.bandColorHex;
            var colorPanel = parent.add("panel", undefined, getLabel("panel.color"));
            setupPanel(colorPanel, 6);

            /* ラジオボタンは同じグループに入れて排他にする / Keep the radios in one group so they are exclusive */
            var modeRow = colorPanel.add("group");
            setupRow(modeRow);
            var modeRadios = [];
            for (var r = 0; r < colorModeKeys.length; r++) {
                var modeRadio = modeRow.add("radiobutton", undefined, getLabel("radio." + colorModeKeys[r]));
                modeRadio.onClick = function () {
                    updateEnabled();
                    updatePreview();
                };
                modeRadios.push(modeRadio);
            }
            modeRadios[1].helpTip = getLabel("tooltip.branchColor");

            /* 色見本（クリックで標準のカラーピッカー）/ Swatch that opens the standard Color Picker */
            var swatchRow = colorPanel.add("group");
            setupRow(swatchRow, "left", 3);
            var swatchLabel = swatchRow.add("statictext", undefined, labelText("fieldLabel.bandColor"));
            swatchLabel.preferredSize.width = LABEL_WIDTH;
            swatchLabel.justify = "right";
            var colorSwatch = swatchRow.add("group");
            colorSwatch.preferredSize = COLOR_SWATCH_SIZE;
            colorSwatch.helpTip = swatchLabel.helpTip = getLabel("tooltip.bandColor");
            colorSwatch.onDraw = function () {
                var swatchGraphics = colorSwatch.graphics;
                var swatchRgb = parseHexColor(bandColorHex) || [0, 0, 0];
                var alpha = colorSwatch.enabled ? 1 : 0.3;
                /* rectPath の前に newPath() しないとパスが累積する / Call newPath() before rectPath or paths accumulate */
                swatchGraphics.newPath();
                swatchGraphics.rectPath(0, 0, COLOR_SWATCH_SIZE[0], COLOR_SWATCH_SIZE[1]);
                swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, [swatchRgb[0] / 255, swatchRgb[1] / 255, swatchRgb[2] / 255, alpha]));
                swatchGraphics.newPath();
                swatchGraphics.rectPath(0.5, 0.5, COLOR_SWATCH_SIZE[0] - 1, COLOR_SWATCH_SIZE[1] - 1);
                swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, alpha], 1));
            };
            colorSwatch.addEventListener("click", function () {
                if (!colorSwatch.enabled) return;
                /* 選んだ色は戻り値で返る。キャンセルは渡した色がそのまま返る / The pick is returned; cancel returns the color passed in */
                var pickedHex = colorToHex(app.showColorPicker(hexToDocumentColor(doc, bandColorHex)));
                if (!pickedHex || pickedHex === bandColorHex) return;
                setHex(pickedHex);
                updatePreview();
            });

            var opacityInput = addSteppedField(colorPanel, {
                label: labelText("fieldLabel.bandOpacity"), labelWidth: LABEL_WIDTH,
                text: colorSettings.bandOpacity + "%", characters: FIELD_CHARS,
                step: 1, min: 0, max: 100, unit: "%",
                onStep: function () { updatePreview(); }
            });
            opacityInput.helpTip = opacityInput.fieldLabel.helpTip = getLabel("tooltip.bandOpacity");
            var normalizeOpacity = opacityInput.onChange;
            opacityInput.onChange = function () {
                normalizeOpacity();
                updatePreview();
            };

            /**
             * 選んでいる色の付け方のキーを返す
             * @returns {string} "single" / "branch"
             */
            function getMode() {
                return modeRadios[1].value ? "branch" : "single";
            }

            /**
             * 色の付け方を選ぶ
             * @param {string} modeKey - "single" / "branch"
             * @returns {void}
             */
            function setMode(modeKey) {
                modeRadios[0].value = (modeKey !== "branch");
                modeRadios[1].value = (modeKey === "branch");
                updateEnabled();
            }

            /**
             * 単色の色を返す
             * @returns {string} #RRGGBB
             */
            function getHex() {
                return bandColorHex;
            }

            /**
             * 単色の色を変えて色見本を描き直す（group に notify は無いので hide → show）
             * @param {string} hexText - #RRGGBB
             * @returns {void}
             */
            function setHex(hexText) {
                bandColorHex = hexText;
                colorSwatch.hide();
                colorSwatch.show();
            }

            /**
             * 単色のときだけ色見本を有効にする
             * @returns {void}
             */
            function updateEnabled() {
                var isSingle = getMode() === "single";
                swatchLabel.enabled = isSingle;
                colorSwatch.enabled = isSingle;
                colorSwatch.hide();
                colorSwatch.show();
            }

            setMode(colorSettings.colorMode);
            return { getMode: getMode, setMode: setMode, getHex: getHex, setHex: setHex, opacityInput: opacityInput };
        }

        /**
         * 分岐点のパネル（形のラジオボタン・幅・高さ）を追加する
         * @param {Group} parent - 追加先
         * @param {string} panelKey - LABELS.panel のキー（"junction1" / "junction2"）
         * @param {string} initialShape - 形の初期値
         * @param {string} initialSizeMode - 幅・高さの意味の初期値（"size" / "margin"）
         * @param {number} initialWidthPt - 幅の初期値（pt）
         * @param {number} initialHeightPt - 高さの初期値（pt）
         * @param {boolean} initialLinked - 幅と高さの連動の初期値
         * @returns {{getShape: Function, setShape: Function, getSizeMode: Function, setSizeMode: Function, widthInput: EditText, heightInput: EditText, linkToggle: Group, setLinked: Function}} パネルの操作
         */
        function addJunctionPanel(parent, panelKey, initialShape, initialSizeMode, initialWidthPt, initialHeightPt, initialLinked) {
            var junctionPanel = parent.add("panel", undefined, getLabel("panel." + panelKey));
            setupPanel(junctionPanel, 6);
            junctionPanel.helpTip = getLabel("tooltip." + panelKey);
            /* 左に形のラジオボタンを縦に、右に幅・高さ / Shape radios stacked on the left, width and height on the right */
            var junctionBody = junctionPanel.add("group");
            junctionBody.orientation = "row";
            junctionBody.alignChildren = ["left", "center"];
            junctionBody.spacing = COLUMN_SPACING;
            /* ラジオボタンは同じグループに入れて排他にする / Keep the radios in one group so they are exclusive */
            var shapeColumn = junctionBody.add("group");
            shapeColumn.orientation = "column";
            shapeColumn.alignChildren = ["left", "center"];
            shapeColumn.spacing = 6;
            var shapeRadios = [];
            for (var r = 0; r < junctionShapeKeys.length; r++) {
                var shapeRadio = shapeColumn.add("radiobutton", undefined, getLabel("radio." + junctionShapeKeys[r]));
                shapeRadio.helpTip = getLabel("tooltip.junctionShape");
                shapeRadio.onClick = function () {
                    updateEnabled();
                    updatePreview();
                };
                shapeRadios.push(shapeRadio);
            }
            /* 右側：大きさ／余白のラジオボタンの下に、幅・高さと連動のアイコン / Right side: size/margin radios above width, height and the link toggle */
            var sizeColumn = junctionBody.add("group");
            sizeColumn.orientation = "column";
            sizeColumn.alignChildren = ["left", "center"];
            sizeColumn.spacing = 6;
            var sizeModeRow = sizeColumn.add("group");
            setupRow(sizeModeRow);
            var rbBySize = sizeModeRow.add("radiobutton", undefined, getLabel("radio.bySize"));
            var rbByMargin = sizeModeRow.add("radiobutton", undefined, getLabel("radio.byMargin"));
            rbBySize.helpTip = rbByMargin.helpTip = getLabel("tooltip.sizeMode");
            rbBySize.value = (initialSizeMode !== "margin");
            rbByMargin.value = (initialSizeMode === "margin");
            rbBySize.onClick = updatePreview;
            rbByMargin.onClick = updatePreview;
            /* 幅・高さを縦に積み、右に連動のアイコン / Width and height stacked, with the link toggle to their right */
            var sizeRow = sizeColumn.add("group");
            sizeRow.orientation = "row";
            sizeRow.alignChildren = ["left", "center"];
            sizeRow.spacing = 6;
            var sizeFields = sizeRow.add("group");
            sizeFields.orientation = "column";
            sizeFields.alignChildren = ["right", "center"]; /* 項目名の幅が違っても入力欄の右端をそろえる / align the fields' right edges despite label widths */
            sizeFields.spacing = 6;
            /* 項目名はナリユキの幅（パネルの中央が空かないように）/ Natural-width labels so the panel has no gap in the middle */
            var widthInput = addLengthField(sizeFields, "junctionWidth", initialWidthPt, lengthUnit, 0, syncLinkedSize, true);
            var heightInput = addLengthField(sizeFields, "junctionHeight", initialHeightPt, lengthUnit, 0, syncLinkedSize, true);
            var linkToggle = addLinkToggle(sizeRow, initialLinked, function () {
                if (linkToggle.value) syncLinkedSize(widthInput);
                updatePreview();
            });
            linkToggle.helpTip = getLabel("tooltip.linkJunctionSize");
            /* 連動中に開いたら高さを幅にそろえる / When opened linked, match the height to the width */
            if (linkToggle.value) setFieldText(heightInput, widthInput.text);

            /**
             * 連動中なら、変えた欄の値をもう一方の欄へ写す
             * @param {EditText} changedInput - 値を変えた欄
             * @returns {void}
             */
            function syncLinkedSize(changedInput) {
                if (!linkToggle || !linkToggle.value) return;
                setFieldText((changedInput === widthInput) ? heightInput : widthInput, changedInput.text);
            }

            /**
             * 連動の状態をコードから変える（連動にしたら高さを幅にそろえる）
             * @param {boolean} isLinked - 連動にするなら true
             * @returns {void}
             */
            function setLinked(isLinked) {
                setLinkToggleValue(linkToggle, isLinked);
                syncLinkedSize(widthInput);
            }

            /**
             * 幅・高さの意味を返す
             * @returns {string} "size" / "margin"
             */
            function getSizeMode() {
                return rbByMargin.value ? "margin" : "size";
            }

            /**
             * 幅・高さの意味を選ぶ
             * @param {string} sizeMode - "size" / "margin"
             * @returns {void}
             */
            function setSizeMode(sizeMode) {
                rbBySize.value = (sizeMode !== "margin");
                rbByMargin.value = (sizeMode === "margin");
            }

            /**
             * 選んでいる形のキーを返す
             * @returns {string} "circle" / "rectangle" / "none"
             */
            function getShape() {
                for (var k = 0; k < shapeRadios.length; k++) {
                    if (shapeRadios[k].value) return junctionShapeKeys[k];
                }
                return junctionShapeKeys[0];
            }
            /**
             * 形を選ぶ
             * @param {string} shapeKey - "circle" / "rectangle" / "none"
             * @returns {void}
             */
            function setShape(shapeKey) {
                var shapeIndex = Math.max(0, indexOfKey(junctionShapeKeys, shapeKey));
                for (var k = 0; k < shapeRadios.length; k++) shapeRadios[k].value = (k === shapeIndex);
                updateEnabled();
            }
            /**
             * 形を置かないときは幅・高さの欄を無効にする
             * @returns {void}
             */
            function updateEnabled() {
                var hasShape = getShape() !== "none";
                setSteppedFieldEnabled(widthInput, hasShape);
                setSteppedFieldEnabled(heightInput, hasShape);
                setLinkToggleEnabled(linkToggle, hasShape);
                rbBySize.enabled = hasShape;
                rbByMargin.enabled = hasShape;
            }

            setShape(initialShape);
            return { getShape: getShape, setShape: setShape, getSizeMode: getSizeMode, setSizeMode: setSizeMode, widthInput: widthInput, heightInput: heightInput, linkToggle: linkToggle, setLinked: setLinked };
        }

        var buttonRow = addButtonRow(flowDialog);
        var btnReset = buttonRow.leftGroup.add("button", undefined, getLabel("button.reset"));
        btnReset.helpTip = getLabel("tooltip.reset");
        var fitViewControls = FitViewToItems.addControls(buttonRow.leftGroup);
        fitViewControls.checkbox.helpTip = getLabel("tooltip.fitView");
        var initialViewState = FitViewToItems.captureView(doc); /* キャンセルで戻す表示 / view restored on Cancel */
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 長さの数値欄を追加する
         * @param {Panel} parent - 追加先
         * @param {string} labelKey - LABELS.fieldLabel のキー
         * @param {number} valuePt - 初期値（pt）
         * @param {{label: string, pointsPerUnit: number}} unitInfo - 表示の単位
         * @param {number} minValue - 下限
         * @param {Function} [beforePreview] - 値が変わったとき、プレビューの前に呼ぶ関数（入力欄を渡す）
         * @param {boolean} [isNaturalLabel] - true なら項目名をナリユキの幅にする（省略時は LABEL_WIDTH で右揃え）
         * @returns {EditText} 入力欄
         */
        function addLengthField(parent, labelKey, valuePt, unitInfo, minValue, beforePreview, isNaturalLabel) {
            var lengthInput = addSteppedField(parent, {
                label: labelText("fieldLabel." + labelKey), labelWidth: isNaturalLabel ? 0 : LABEL_WIDTH,
                text: formatLength(valuePt, unitInfo), characters: FIELD_CHARS,
                step: 1, min: minValue, unit: " " + unitInfo.label,
                onStep: function () {
                    if (beforePreview) beforePreview(lengthInput);
                    updatePreview();
                }
            });
            var tooltipEntry = LABELS.tooltip[labelKey];
            if (tooltipEntry) lengthInput.helpTip = lengthInput.fieldLabel.helpTip = getLabel(tooltipEntry);
            /* 部品の onChange（値の正規化）のあとでプレビューする / Preview after the part's onChange normalizes the value */
            var normalizeOnChange = lengthInput.onChange;
            lengthInput.onChange = function () {
                normalizeOnChange();
                if (beforePreview) beforePreview(lengthInput);
                updatePreview();
            };
            return lengthInput;
        }

        /**
         * フロー以外の設定を初期値に戻す（ノードの高さと文字サイズは開いたときの値）
         * @returns {void}
         */
        function resetSettings() {
            setWidthScale(DEFAULT_SETTINGS.widthScale);
            setFieldText(nodeHeightInput, formatLength(initialNodeHeight, lengthUnit));
            setFieldText(exitRatioInput, DEFAULT_SETTINGS.exitRatio + "%");
            setFieldText(curveLengthInput, formatLength(DEFAULT_SETTINGS.curveLengthPt, lengthUnit));
            setFieldText(straightLengthInput, formatLength(DEFAULT_SETTINGS.straightLengthPt, lengthUnit));
            setFieldText(branchGapInput, formatLength(DEFAULT_SETTINGS.branchGapPt, lengthUnit));
            setFieldText(minWidthInput, formatLength(DEFAULT_SETTINGS.minWidthPt, lengthUnit));
            /* ラベルから取った文字サイズはそのまま / Keep a size taken from the labels */
            if (nodeInfo.labelTexts.length === 0) setFieldText(fontSizeInput, formatLength(DEFAULT_SETTINGS.fontSizePt, textUnit));
            chkShowValues.value = DEFAULT_SETTINGS.showValues;
            chkArrowTips.value = DEFAULT_SETTINGS.arrowTips;
            chkStrokeBands.value = DEFAULT_SETTINGS.strokeBands;
            chkMoveLabels.value = DEFAULT_SETTINGS.moveLabels;
            junction1Controls.setShape(DEFAULT_SETTINGS.junction1Shape);
            junction1Controls.setSizeMode(DEFAULT_SETTINGS.junction1SizeMode);
            setFieldText(junction1Controls.widthInput, formatLength(DEFAULT_SETTINGS.junction1WidthPt, lengthUnit));
            setFieldText(junction1Controls.heightInput, formatLength(DEFAULT_SETTINGS.junction1HeightPt, lengthUnit));
            junction1Controls.setLinked(DEFAULT_SETTINGS.junction1Linked);
            junction2Controls.setShape(DEFAULT_SETTINGS.junction2Shape);
            junction2Controls.setSizeMode(DEFAULT_SETTINGS.junction2SizeMode);
            setFieldText(junction2Controls.widthInput, formatLength(DEFAULT_SETTINGS.junction2WidthPt, lengthUnit));
            setFieldText(junction2Controls.heightInput, formatLength(DEFAULT_SETTINGS.junction2HeightPt, lengthUnit));
            junction2Controls.setLinked(DEFAULT_SETTINGS.junction2Linked);
            colorControls.setMode(DEFAULT_SETTINGS.colorMode);
            colorControls.setHex(DEFAULT_SETTINGS.bandColorHex);
            setFieldText(colorControls.opacityInput, DEFAULT_SETTINGS.bandOpacity + "%");
            updatePreview();
        }

        /**
         * ダイアログボックスの今の設定を返す
         * @returns {Object} 設定（DEFAULT_SETTINGS に nodeHeightPt を足した形）
         */
        function collectSettings() {
            return {
                flowText: flowTextField.text,
                nodeHeightPt: readPt(nodeHeightInput, lengthUnit),
                exitRatio: parseFloat(exitRatioInput.text),
                widthScale: getWidthScale(),
                curveLengthPt: readPt(curveLengthInput, lengthUnit),
                straightLengthPt: readPt(straightLengthInput, lengthUnit),
                branchGapPt: readPt(branchGapInput, lengthUnit),
                minWidthPt: readPt(minWidthInput, lengthUnit),
                fontSizePt: readPt(fontSizeInput, textUnit),
                showValues: chkShowValues.value,
                arrowTips: chkArrowTips.value,
                strokeBands: chkStrokeBands.value,
                moveLabels: chkMoveLabels.value,
                junction1Shape: junction1Controls.getShape(),
                junction1SizeMode: junction1Controls.getSizeMode(),
                junction1WidthPt: readPt(junction1Controls.widthInput, lengthUnit),
                junction1HeightPt: readPt(junction1Controls.heightInput, lengthUnit),
                junction1Linked: junction1Controls.linkToggle.value,
                junction2Shape: junction2Controls.getShape(),
                junction2SizeMode: junction2Controls.getSizeMode(),
                junction2WidthPt: readPt(junction2Controls.widthInput, lengthUnit),
                junction2HeightPt: readPt(junction2Controls.heightInput, lengthUnit),
                junction2Linked: junction2Controls.linkToggle.value,
                colorMode: colorControls.getMode(),
                bandColorHex: colorControls.getHex(),
                bandOpacity: parseFloat(colorControls.opacityInput.text)
            };
        }

        /**
         * プレビューを消す
         * @returns {void}
         */
        function removePreview() {
            if (!previewGroup) return;
            previewGroup.remove();
            previewGroup = null;
        }

        /**
         * プレビューを描き直す（常に表示。入力に誤りがあるときは描かない）
         * @returns {void}
         */
        function updatePreview() {
            removePreview();
            restoreLabelPositions(nodeInfo);
            var flowSettings = collectSettings();
            var flowRead = readFlows(flowSettings.flowText);
            if (!flowRead.errorMessage) previewGroup = drawSankeyFlow(doc, nodeInfo, flowRead.flowRoots, flowSettings);
            fitViewToPreview();
            app.redraw();
        }

        /**
         * ［画面にフィット］がオンなら、プレビューとノード・ラベルが指定の割合で収まるよう表示を合わせる
         * @returns {void}
         */
        function fitViewToPreview() {
            if (!fitViewControls.checkbox.value || !previewGroup) return;
            FitViewToItems.fit([previewGroup, nodeInfo.nodeItem].concat(nodeInfo.labelTexts), { doc: doc, fillRatio: fitViewControls.getFillRatio() });
        }

        flowTextField.onChange = updatePreview;
        chkShowValues.onClick = updatePreview;
        chkArrowTips.onClick = updatePreview;
        chkStrokeBands.onClick = updatePreview;
        chkMoveLabels.onClick = updatePreview;
        btnReset.onClick = resetSettings;
        fitViewControls.checkbox.onClick = function () {
            fitViewControls.updateEnabled();
            fitViewToPreview();
            app.redraw();
        };
        fitViewControls.percentInput.onChange = function () {
            fitViewToPreview();
            app.redraw();
        };
        /* 部品の％欄には∧∨が無いので、↑↓キーだけほかの数値欄と同じ増減をつなぐ（10〜100%、整数）/ Arrow keys for the fit percentage */
        var fitPercentOptions = { step: 1, min: 10, max: 100, integer: true };
        bindSteppedArrowKeys(fitViewControls.percentInput, {
            stepBy: function (direction) {
                var percentInput = fitViewControls.percentInput;
                if (!percentInput.enabled) return;
                var percentValue = parseFloat(percentInput.text);
                if (isNaN(percentValue)) percentValue = 0;
                writeSteppedValue(percentInput, computeSteppedValue(percentValue, direction, fitPercentOptions), fitPercentOptions);
                fitViewToPreview();
                app.redraw();
            }
        });

        /* 入力を確かめてから閉じる / Check the input before closing */
        btnOK.onClick = function () {
            var flowRead = readFlows(flowTextField.text);
            if (flowRead.errorMessage) {
                alert(flowRead.errorMessage);
                return;
            }
            flowDialog.close(1);
        };
        btnCancel.onClick = function () {
            flowDialog.close(2);
        };

        flowDialog.onShow = function () {
            updatePreview();
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(flowDialog, SCRIPT_NAME);
        var dialogResult = flowDialog.show();
        removePreview();
        if (dialogResult !== 1) {
            restoreLabelPositions(nodeInfo);
            FitViewToItems.restoreView(initialViewState, doc);
            app.redraw();
            return null;
        }

        /* 確定はプレビューと別に描き直す / Draw the final diagram separately from the preview */
        var finalSettings = collectSettings();
        var flowGroup = drawSankeyFlow(doc, nodeInfo, readFlows(finalSettings.flowText).flowRoots, finalSettings);
        doc.selection = null;
        flowGroup.selected = true;
        return finalSettings;
    }

    // -----------------------------------------
    // ダイアログボックスの下請け / Dialog helpers
    // -----------------------------------------

    /**
     * pt の値を欄の単位の表示（「20 mm」の形、小数2桁まで）にする
     * @param {number} valuePt - 値（pt）
     * @param {{label: string, pointsPerUnit: number}} unitInfo - 欄の単位
     * @returns {string} 表示の文字列
     */
    function formatLength(valuePt, unitInfo) {
        return (Math.round(valuePt / unitInfo.pointsPerUnit * 100) / 100) + " " + unitInfo.label;
    }

    /**
     * 数値欄に値を書き込む（部品が戻す先の値もそろえる）
     * @param {EditText} numberInput - 入力欄
     * @param {string} fieldText - 書き込む文字列
     * @returns {void}
     */
    function setFieldText(numberInput, fieldText) {
        numberInput.text = fieldText;
        numberInput.lastValidText = fieldText;
    }

    /**
     * 入力欄の値を pt で返す
     * @param {EditText} lengthInput - 入力欄
     * @param {{pointsPerUnit: number}} unitInfo - 欄の単位
     * @returns {number} pt の値
     */
    function readPt(lengthInput, unitInfo) {
        return parseFloat(lengthInput.text) * unitInfo.pointsPerUnit;
    }

    /**
     * フローを読み取り、描けない入力ならその理由を返す
     * @param {string} flowText - 入力された文字列
     * @returns {{flowRoots: Object[], errorMessage: string}} 1段目の枝と、エラーの文言（無ければ空文字）
     */
    function readFlows(flowText) {
        var parsedFlow = parseFlowText(flowText);
        if (parsedFlow.errorLine > 0) return { flowRoots: [], errorMessage: getLabel("alert.noValue", [parsedFlow.errorLine, parsedFlow.errorText]) };
        if (parsedFlow.children.length === 0) return { flowRoots: [], errorMessage: getLabel("alert.noFlows") };
        var valueTotal = 0;
        for (var i = 0; i < parsedFlow.children.length; i++) valueTotal += parsedFlow.children[i].value;
        if (valueTotal <= 0) return { flowRoots: [], errorMessage: getLabel("alert.zeroTotal") };
        return { flowRoots: parsedFlow.children, errorMessage: "" };
    }

    /**
     * キーの配列から位置を返す
     * @param {string[]} keys - キーの配列
     * @param {string} key - 探すキー
     * @returns {number} 位置（無ければ -1）
     */
    function indexOfKey(keys, key) {
        for (var i = 0; i < keys.length; i++) {
            if (keys[i] === key) return i;
        }
        return -1;
    }

    main();

})();
