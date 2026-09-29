#target illustrator
#targetengine "AiCurlyBracketMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

カーリーブラケット（波括弧）のパスを、2つの半径・中央の移動・全体の長さ・線の設定（ブラシを含む）から作成します。
値を変えるたびにアートボード上のプレビューが更新され、オブジェクトを選択して実行すると、その大きさと向きに合わせて配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCurlyBracketMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nd6b3e36ff79d

### Overview

Creates a curly bracket path from two radii, the position of its middle point, the overall length and the stroke settings, brush included.
The artboard preview updates as the values change, and running it with a selection fits the bracket to that selection's size and direction.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCurlyBracketMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiCurlyBracketMaker";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCurlyBracketMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCurlyBracketMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd6b3e36ff79d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_TOTAL_LENGTH_MM  = 100;         /* 全体の長さの初期値（mm）/ initial overall length (mm) */
    var DEFAULT_RADIUS_RATIO     = 1 / 15;      /* 半径の初期値＝全体の長さ×この比率 / initial radius = overall length x this ratio */
    var DEFAULT_LINK_RADIUS      = true;        /* 2つの半径を連動させて始めるか / whether the two radii start linked */
    var DEFAULT_CHAMFER          = false;       /* 面取りを初期状態でONにするか / whether the chamfer starts on */

    var DEFAULT_CENTER_OFFSET_MM = 0;           /* 中央の位置の初期値（mm、0で中央）/ initial center offset (mm, 0 = centered) */
    var DEFAULT_EXTENSION_PT     = 0;           /* 両端の延長の初期値（pt）/ initial extension at both ends (pt) */
    var DEFAULT_OBJECT_GAP_MM        = 2;           /* 選択オブジェクトとの余白の初期値（mm）/ initial gap from the selection (mm) */
    var DEFAULT_STROKE_WIDTH_PT  = 2;           /* 線の太さの初期値（pt）/ initial stroke width (pt) */
    var DEFAULT_DIRECTION        = "right";     /* 初期の向き（DIRECTION_KEYS のいずれか）/ initial direction (one of DIRECTION_KEYS) */
    var DEFAULT_STROKE_CAP       = "buttCap";   /* 初期の線端 / initial stroke cap */
    var DEFAULT_CORNER_JOIN      = "miterJoin"; /* 初期の角の形状 / initial corner shape */
    var DEFAULT_BRUSH            = "";          /* 初期のブラシ名（空でブラシなし）/ initial brush name ("" applies no brush) */

    /* ブラシライブラリー（このスクリプトと同じ場所に置く）。ドロップダウンにはこのファイルのブラシが並ぶ
       Brush library, kept beside this script; its brushes are what the dropdown lists */
    var BRUSH_LIBRARY_NAME = "brushlibrary.ai";

    /* ドロップダウンに出さないブラシ名。どの書類にも最初から入っている既定のブラシを、日本語UI・英語UIの名前で外す
       Brush names left out of the dropdown: the default brush every document already carries, in Japanese and English */
    var EXCLUDED_BRUSH_NAMES = ["カリグラフィブラシをタッチ", "Touch Calligraphic Brush"];

    /* 生成したブラケットに付ける目印。値には向きを入れ、選び直したときの復元に使う / Marker added to the bracket we create; its value holds the direction, used when it is selected again */
    var BRACKET_TAG_NAME = "AiCurlyBracketMaker";

    /* 面取りに使うジグザグ効果（大きさ0・折り返し0・roundness 0＝直線的に）/ Zig Zag effect for the chamfer: amount 0, ridges 0, roundness 0 (corner points) */
    var CHAMFER_EFFECT_XML = '<LiveEffect name="Adobe Zigzag"><Dict data="R amount 0 R relAmount 0 R absoluteness 1 R ridges 0 R roundness 0 "/></LiveEffect>';

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 8;                /* パネル内の要素間隔 / panel spacing */
    var FIELD_ROW_SPACING  = 6;                /* ラベル・入力欄・単位表記の間隔 / gap inside a labeled row */
    var LABEL_WIDTH        = 98;               /* 行ラベルの共通幅 / shared width of row labels */
    var FIELD_CHARACTERS   = 3;                /* 数値欄の文字数（＝最小幅）/ characters of a numeric field */
    var DROPDOWN_WIDTH     = 150;              /* ドロップダウンの幅（長い項目でダイアログを広げないよう固定）/ fixed width, so a long item cannot widen the dialog */

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
            title: { ja: "カーリーブラケットの作成", en: "Create Curly Bracket" }
        },
        panel: {
            shapeAndSize: { ja: "形状と大きさ", en: "Shape & Size" },
            stroke:       { ja: "線（線幅と形状）", en: "Stroke (Width & Shape)" }
        },
        fieldLabel: {
            centerRadius: { ja: "半径（中央）", en: "Center Radius" },
            endRadius:    { ja: "半径（両端）", en: "End Radius" },
            centerOffset: { ja: "中央の移動", en: "Center Shift" },
            totalLength:  { ja: "全体の長さ", en: "Length" },
            extension:    { ja: "両端の延長", en: "End Extension" },
            objectGap:    { ja: "オフセット", en: "Offset" },
            strokeWidth:  { ja: "線の太さ", en: "Stroke Width" },
            strokeCap:    { ja: "線端", en: "Cap" },
            cornerJoin:   { ja: "角の形状", en: "Corner Shape" },
            brush:        { ja: "ブラシ", en: "Brush" },
            direction:    { ja: "向き", en: "Direction" }
        },
        checkbox: {
            chamfer:    { ja: "面取り", en: "Chamfer" }
        },
        radio: {
            buttCap:   { ja: "線端なし", en: "Butt Cap" },
            roundCap:  { ja: "丸型線端", en: "Round Cap" },
            miterJoin: { ja: "マイター結合", en: "Miter Join" },
            roundJoin: { ja: "ラウンド結合", en: "Round Join" },
            up:        { ja: "上", en: "Up" },
            down:      { ja: "下", en: "Down" },
            left:      { ja: "左", en: "Left" },
            right:     { ja: "右", en: "Right" }
        },
        dropdown: {
            noBrush: { ja: "なし", en: "None" }
        },
        tooltip: {
            arrowKeys:    { ja: "↑↓で増減（Shiftで10単位、Optionで0.1刻み）", en: "Up/Down to step (Shift for 10s, Option for 0.1)" },
            centerRadius: { ja: "中央の突起を作る円弧の半径。0にすると突起がなくなり直線になります", en: "Radius of the arcs that form the point in the middle. 0 drops the point and leaves a straight line" },
            endRadius:    { ja: "両端で外へ折れ返る円弧の半径。0にすると円弧なしの直角になります", en: "Radius of the arcs that curl outward at both ends. 0 leaves a right angle with no arc" },
            linkRadius:   { ja: "両端の半径を中央に合わせる（ONのあいだ両端は編集できません）", en: "Keep the end radius equal to the center radius (the end field is disabled while on)" },
            chamfer:      { ja: "ジグザグ効果（大きさ0・折り返し0）を適用し、円弧を直線でつないだ面取りにします", en: "Applies a Zig Zag effect (size 0, ridges 0) so the arcs become straight chamfers" },
            centerOffset: { ja: "中央の突起を長さ方向にずらす量。0で中央、左右の向きでは上へ、上下の向きでは右へ動きます（マイナスで逆）", en: "Moves the middle point along the length. 0 keeps it centered; up for a left/right bracket, right for an up/down one (negative reverses)" },
            totalLength:  { ja: "ブラケット全体の長さ。半径を変えても総長は変わらず、直線部分が自動で伸縮します", en: "Overall length of the bracket. Changing the radius keeps this length and resizes the straight sections instead" },
            extension:    { ja: "両端から外側へ伸ばす直線の長さ（0で延長なし）", en: "Straight run added at both ends, away from the middle (0 adds none)" },
            objectGap:    { ja: "選択オブジェクトとブラケットのあいだの間隔（選択して実行したときのみ有効）", en: "Gap between the selection and the bracket (only when run with a selection)" },
            strokeWidth:  { ja: "ブラケットの線幅（pt）", en: "Stroke width of the bracket (pt)" },
            strokeCap:    { ja: "両端の線の先を丸めるかどうか", en: "Whether both ends of the stroke are rounded" },
            cornerJoin:   { ja: "中央の角を尖らせるか丸めるか", en: "Whether the corner in the middle is pointed or rounded" },
            brush:        { ja: "ブラシライブラリーのブラシを適用します（「なし」で通常の線）", en: "Applies a brush from the brush library (\"None\" leaves a plain stroke)" },
            direction:    { ja: "中央の突起を向ける方向。選択オブジェクトがあれば、その辺に沿って配置されます", en: "The direction the middle point faces. With a selection, the bracket hugs the matching edge" },
            create:       { ja: "プレビューの状態で確定します（大きさの参照にしたパスは削除されます）", en: "Commits exactly what the preview shows (a path used as the size reference is removed)" },
            cancel:       { ja: "作成せずに閉じます。プレビューは削除し、隠していた参照のパスは元に戻します", en: "Closes without creating: the preview is removed and a hidden reference path is restored" },
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
            create: { ja: "作成", en: "Create" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            lockedLayer:  { ja: "アクティブレイヤーがロックまたは非表示です。", en: "The active layer is locked or hidden." },
            invalidValue: { ja: "数値が正しくありません。線の太さは0より大きい値、ほかの項目は0以上を入力してください。", en: "Some values are not valid. The stroke width must be greater than 0, and the other fields 0 or more." }
        }
    };

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ダイアログ全体の並びと余白を設定する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルの並びと余白を設定する
     * @param {Panel} targetPanel - 対象パネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
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
     * 共通幅で右揃えの行ラベルを追加する
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} rowLabelText - 表示するラベル
     * @returns {StaticText} 生成したラベル
     */
    function addRowLabel(parentRow, rowLabelText) {
        var rowLabel = parentRow.add("statictext", undefined, rowLabelText);
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 「ラベル＋数値欄＋単位」の1行を生成する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} rowLabelText - 行ラベル（コロン付き）
     * @param {string} initialValue - 数値欄の初期値
     * @param {string} unitSuffix - 数値欄に添える単位表記
     * @param {string} helpTipText - 数値欄のツールチップ
     * @returns {EditText} 生成した数値欄（増減の設定は .stepOptions、∧∨は .stepperGroup、行は .fieldRow で参照できる）
     */
    function addNumberFieldRow(parentContainer, rowLabelText, initialValue, unitSuffix, helpTipText) {
        var numberFieldRow = parentContainer.add("group");
        setupRow(numberFieldRow, "left", FIELD_ROW_SPACING);
        addRowLabel(numberFieldRow, rowLabelText);

        /* 下限と増減後の処理は bindPreviewField() があとから stepOptions に入れる
           / bindPreviewField() fills in the minimum and the step callback later */
        var stepOptions = {};
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = numberFieldRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;
        var numberField;
        var numberStepper = addStepper(stepperFieldGroup, function () { return numberField; }, stepOptions);
        numberField = stepperFieldGroup.add("edittext", undefined, initialValue);
        numberField.characters = FIELD_CHARACTERS;
        numberField.stepOptions = stepOptions;
        numberField.fieldRow = numberFieldRow; /* 行ごとディム表示するときの参照（親は∧∨と組んだグループ） / the whole row, for dimming */
        numberField.stepperGroup = numberStepper;
        bindSteppedArrowKeys(numberField, numberStepper);
        /* ↑↓キーの説明は全数値欄に共通で添える / Every numeric field carries the shared arrow-key hint */
        numberField.helpTip = helpTipText + "\n" + getLabel('tooltip.arrowKeys');

        numberFieldRow.add("statictext", undefined, unitSuffix);
        return numberField;
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

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
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
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle) {
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = LINK_ICON_SIZE;
        linkToggle.minimumSize = LINK_ICON_SIZE;
        linkToggle.maximumSize = LINK_ICON_SIZE;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = LINK_ICON_SIZE[0];
            var iconHeight = LINK_ICON_SIZE[1];
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
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
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
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22;
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

    // =========================================
    // 入力欄の補助 / Input field helpers
    // =========================================

    /**
     * 入力欄に入れる数値を小数2桁までに丸めて文字列にする
     * @param {number} value - 表示したい値
     * @returns {string} 入力欄に入れる文字列
     */
    function formatFieldValue(value) {
        return String(Math.round(value * 100) / 100);
    }

    // =========================================
    // ブラシ / Brush
    // =========================================

    /* ライブラリーから読んだブラシ名を同一セッション内だけ覚えておく $.global 上のキー / Key on $.global that remembers the library's brush names within this session */
    var BRUSH_LIBRARY_CACHE_KEY = "aiCurlyBracketMakerBrushLibrary";

    /**
     * ブラシライブラリーのファイルを返す（スクリプトと同じ場所）
     * @returns {File|null} ライブラリーのファイル（無ければ null）
     */
    function getBrushLibraryFile() {
        /* スクリプトの場所が取れない実行方法もある / Some ways of running a script leave no path to work from */
        var scriptFolder = File($.fileName).parent;
        if (!scriptFolder) return null;

        var libraryFile = new File(scriptFolder.fsName + "/" + BRUSH_LIBRARY_NAME);
        return libraryFile.exists ? libraryFile : null;
    }

    /**
     * 書類の［ブラシ］パネルに、その名前のブラシがあるか
     * @param {Document} doc - 対象ドキュメント
     * @param {string} brushName - 探すブラシ名
     * @returns {boolean} あれば true
     */
    function hasBrush(doc, brushName) {
        for (var i = 0; i < doc.brushes.length; i++) {
            if (doc.brushes[i].name === brushName) return true;
        }
        return false;
    }

    /**
     * ブラシライブラリーを開いて処理に渡し、開いたぶんは閉じて元の書類に戻る
     * @param {Document} doc - 元の書類（処理後にアクティブへ戻す）
     * @param {function} handleLibrary - ライブラリーのドキュメントを受け取る処理
     * @returns {*} handleLibrary の戻り値（ライブラリーが無ければ null）
     */
    function withBrushLibrary(doc, handleLibrary) {
        var libraryFile = getBrushLibraryFile();
        if (!libraryFile) return null;

        /* すでに開いているライブラリーを開き直すと、開くのではなく手前に来るだけで書類数が変わらない。
           これを目印にすれば、パス文字列の比較（正規化やケース違いで外しうる）に頼らず判定できる
           Reopening an already open library only brings it to front, leaving the document count unchanged;
           that is a safer signal than comparing path strings, which normalization or case can defeat */
        var documentCountBefore = app.documents.length;
        var libraryDoc = app.open(libraryFile);
        var wasLibraryOpen = (app.documents.length === documentCountBefore);

        try {
            return handleLibrary(libraryDoc);
        } finally {
            /* 開いたのがこのスクリプトなら閉じる。もともと開いていたものは触らない / Close only what this script opened; leave what was already open */
            if (!wasLibraryOpen) libraryDoc.close(SaveOptions.DONOTSAVECHANGES);
            app.activeDocument = doc;
        }
    }

    /**
     * ドロップダウンに出さないブラシか
     * @param {string} brushName - 調べるブラシ名
     * @returns {boolean} 除外するなら true
     */
    function isExcludedBrush(brushName) {
        for (var i = 0; i < EXCLUDED_BRUSH_NAMES.length; i++) {
            if (EXCLUDED_BRUSH_NAMES[i] === brushName) return true;
        }
        return false;
    }

    /**
     * ライブラリーのブラシ名を並び順で返す（除外するブラシは飛ばす）
     * @param {Document} doc - 対象ドキュメント
     * @returns {string[]} ブラシ名（ライブラリーが無ければ空）
     */
    function getLibraryBrushNames(doc) {
        var libraryFile = getBrushLibraryFile();
        if (!libraryFile) return [];

        /* 一度読んだ名前は覚えておき、ライブラリーを開き直さない（更新されていれば読み直す）
           Remembered names save reopening the library; a newer file is read again */
        var libraryCache = $.global[BRUSH_LIBRARY_CACHE_KEY];
        var libraryStamp = libraryFile.fsName + "\t" + libraryFile.modified.getTime();
        if (libraryCache && libraryCache.stamp === libraryStamp) return libraryCache.names;

        var libraryNames = withBrushLibrary(doc, function (libraryDoc) {
            var names = [];
            for (var i = 0; i < libraryDoc.brushes.length; i++) {
                var libraryBrushName = libraryDoc.brushes[i].name;
                if (!isExcludedBrush(libraryBrushName)) names.push(libraryBrushName);
            }
            return names;
        }) || [];

        $.global[BRUSH_LIBRARY_CACHE_KEY] = { stamp: libraryStamp, names: libraryNames };
        return libraryNames;
    }

    /* 取り込めなかったブラシ名（実行中だけ覚える）/ Brushes that could not be imported, remembered for this run */
    var failedBrushNames = {};

    /**
     * ブラシを書類の［ブラシ］パネルに用意する（無ければライブラリーから取り込む）
     * ブラシを適用したパスを複製すると、そのパスを消してもブラシはパネルに残る
     * @param {Document} doc - 取り込み先のドキュメント
     * @param {string} brushName - 用意したいブラシ名
     * @returns {boolean} 使える状態になったら true
     */
    function ensureBrush(doc, brushName) {
        if (hasBrush(doc, brushName)) return true;
        /* 一度取り込めなかったブラシは、プレビューのたびにライブラリーを開き直さない / A brush that failed once must not reopen the library on every preview */
        if (failedBrushNames[brushName] === true) return false;

        var carrierCopy = withBrushLibrary(doc, function (libraryDoc) {
            var carrierPath = null;
            var copiedPath = null;
            try {
                var libraryBrush = libraryDoc.brushes.getByName(brushName);
                carrierPath = libraryDoc.pathItems.add();
                carrierPath.setEntirePath([[0, 0], [10, 0]]);
                carrierPath.filled = false;
                carrierPath.stroked = true;
                libraryBrush.applyTo(carrierPath);
                copiedPath = carrierPath.duplicate(doc.activeLayer, ElementPlacement.PLACEATEND);
            } catch (eImport) {
                copiedPath = null;
            }
            if (carrierPath) carrierPath.remove();
            return copiedPath;
        });

        /* 取り込みに使った複製は、ライブラリーを閉じて戻ってから消す / The carrier copy goes once the library is closed and we are back */
        if (carrierCopy) carrierCopy.remove();

        var isBrushReady = hasBrush(doc, brushName);
        if (!isBrushReady) failedBrushNames[brushName] = true;
        return isBrushReady;
    }

    /**
     * 書類のブラシをパスに適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} targetPath - 適用先のパス
     * @param {string} brushName - 適用するブラシの名前
     * @returns {void}
     */
    function applyBrush(doc, targetPath, brushName) {
        try {
            if (!ensureBrush(doc, brushName)) return;
            doc.brushes.getByName(brushName).applyTo(targetPath);
        } catch (eBrush) {
            /* ブラシを適用できなくても、描いたパスはそのまま残す / Keep the path as drawn even when the brush cannot be applied */
        }
    }

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

    // =========================================
    // ブラケットの形状 / Bracket geometry
    // =========================================

    /* ブラケットの寸法は常に mm で扱う / bracket dimensions are always handled in mm */
    var POINTS_PER_MM = UNITS[1].pointsPerUnit;

    /* ベジェ曲線で90度円弧を描くための制御点ハンドル係数 / Handle factor for a 90-degree Bezier arc */
    var KAPPA = 0.55228474983;

    /* 線端ごとの端点の形状 / Stroke cap per end shape */
    var STROKE_CAPS = {
        buttCap:  StrokeCap.BUTTENDCAP,
        roundCap: StrokeCap.ROUNDENDCAP
    };

    /* 角の形状ごとの線の結合 / Stroke join per corner shape */
    var CORNER_JOINS = {
        miterJoin: StrokeJoin.MITERENDJOIN,
        roundJoin: StrokeJoin.ROUNDENDJOIN
    };

    /* 選択できる向き（ラジオの並び順）/ Selectable directions, in the order the radios appear */
    var DIRECTION_KEYS = ["up", "down", "left", "right"];

    /* 右向きを基準にした向きごとの回転量（cos/sin）/ Rotation per direction, based on the right-facing bracket */
    var DIRECTION_ROTATIONS = {
        right: { cos:  1, sin:  0 },
        up:    { cos:  0, sin:  1 },
        left:  { cos: -1, sin:  0 },
        down:  { cos:  0, sin: -1 }
    };

    /**
     * 向きに対応する回転量を返す
     * @param {string} direction - 向きのキー（DIRECTION_KEYS のいずれか）
     * @returns {{cos: number, sin: number}} 回転量（未知のキーは既定の向き）
     */
    function getDirectionRotation(direction) {
        return DIRECTION_ROTATIONS[direction] || DIRECTION_ROTATIONS[DEFAULT_DIRECTION];
    }

    /**
     * 座標を中心まわりに回転する
     * @param {number[]} basePoint - 回転前の座標 [x, y]
     * @param {{cos: number, sin: number}} rotation - 回転量（cos/sin）
     * @param {number} centerX - 回転中心のX座標（pt）
     * @param {number} centerY - 回転中心のY座標（pt）
     * @returns {number[]} 回転後の座標 [x, y]
     */
    function rotateAroundCenter(basePoint, rotation, centerX, centerY) {
        var dx = basePoint[0] - centerX;
        var dy = basePoint[1] - centerY;
        return [
            centerX + dx * rotation.cos - dy * rotation.sin,
            centerY + dx * rotation.sin + dy * rotation.cos
        ];
    }

    /**
     * ハンドルを持たない（前後と直線でつながる）パスポイントを返す
     * @param {number} anchorX - アンカーのX座標（pt）
     * @param {number} anchorY - アンカーのY座標（pt）
     * @returns {{anchor: number[], left: number[], right: number[]}} パスポイント
     */
    function createStraightPoint(anchorX, anchorY) {
        return {
            anchor: [anchorX, anchorY],
            left:   [anchorX, anchorY],
            right:  [anchorX, anchorY]
        };
    }

    /**
     * @typedef {object} BracketSettings
     * @property {number} centerRadiusPt - 中央の突起を作る円弧の半径（pt）
     * @property {number} endRadiusPt - 両端で外へ折れ返る円弧の半径（pt）
     * @property {number} totalLengthPt - 全体の長さ（pt）
     * @property {number} centerOffsetPt - 中央の突起を長さ方向にずらす量（pt、0で中央）
     * @property {number} extensionPt - 両端から外側へ伸ばす長さ（pt）
     * @property {number} objectGapPt - 選択オブジェクトとの余白（pt）
     * @property {number} strokeWidthPt - 線の太さ（pt）
     * @property {string} strokeCap - 線端（STROKE_CAPS のキー）
     * @property {string} cornerJoin - 角の形状（CORNER_JOINS のキー）
     * @property {string|null} brushName - 適用するブラシ名（適用しないなら null）
     * @property {string} direction - 向き（DIRECTION_KEYS のいずれか）
     * @property {boolean} isChamfer - 面取り（ジグザグ効果）を適用するか
     */

    /**
     * 中央をずらす向きを返す（左右の向きでは上が＋、上下の向きでは右が＋）
     * @param {string} direction - 向きのキー（DIRECTION_KEYS のいずれか）
     * @returns {number} 右向き基準の座標系での符号（+1 または -1）
     */
    function getCenterOffsetSign(direction) {
        return (direction === "left" || direction === "up") ? -1 : 1;
    }

    /**
     * 座標を水平の中心線で反転する
     * @param {number[]} basePoint - 反転前の座標 [x, y]
     * @param {number} centerY - 中心線のY座標（pt）
     * @returns {number[]} 反転後の座標 [x, y]
     */
    function mirrorAcrossCenterY(basePoint, centerY) {
        return [basePoint[0], 2 * centerY - basePoint[1]];
    }

    /**
     * 上半分のパスポイントを反転して下半分のパスポイントにする
     * 進行方向が逆になるので、左右のハンドルも入れ替える
     * @param {{anchor: number[], left: number[], right: number[]}} halfPoint - 上半分のパスポイント
     * @param {number} centerY - 中心線のY座標（pt）
     * @returns {{anchor: number[], left: number[], right: number[]}} 下半分のパスポイント
     */
    function mirrorPointAcrossCenterY(halfPoint, centerY) {
        return {
            anchor: mirrorAcrossCenterY(halfPoint.anchor, centerY),
            left:   mirrorAcrossCenterY(halfPoint.right, centerY),
            right:  mirrorAcrossCenterY(halfPoint.left, centerY)
        };
    }

    /**
     * 同じ位置に来た2つのパスポイントを1つにまとめる（入りのハンドルは前から、出のハンドルは後ろから）
     * @param {{anchor: number[], left: number[], right: number[]}} firstPoint - 前のパスポイント
     * @param {{anchor: number[], left: number[], right: number[]}} secondPoint - 後ろのパスポイント
     * @returns {{anchor: number[], left: number[], right: number[]}} まとめたパスポイント
     */
    function mergePathPoints(firstPoint, secondPoint) {
        return {
            anchor: secondPoint.anchor,
            left:   firstPoint.left,
            right:  secondPoint.right
        };
    }

    /**
     * 片側のパスポイントを、腕の先から中央の突起の手前まで順に返す（右向き基準）
     * 下半分はこの結果を中央で反転して使う
     * @param {number} endRadius - 両端で外へ折れ返る円弧の半径（pt）
     * @param {number} centerRadius - 中央の突起を作る円弧の半径（pt）
     * @param {number} straightLength - この側の直線部分の長さ（pt）
     * @param {number} extension - 腕の先に足す延長（pt）
     * @param {number} centerX - 中心軸のX座標（pt）
     * @param {number} centerY - 中央の突起のY座標（pt）
     * @returns {Array<{anchor: number[], left: number[], right: number[]}>} 片側のパスポイント
     */
    function buildHalfPoints(endRadius, centerRadius, straightLength, extension, centerX, centerY) {
        var endHandle = endRadius * KAPPA;
        var centerHandle = centerRadius * KAPPA;
        var straightTopY = centerY + centerRadius + straightLength; /* 直線の上端 / Top of the straight section */
        var straightBottomY = centerY + centerRadius;               /* 直線の下端 / Bottom of the straight section */
        var armEndY = straightTopY + endRadius;                  /* 腕の先のY座標 / Y of the arm end */
        /* 腕の先（両端の円弧の外端）/ Arm end (outer end of the end arc) */
        var armEndPoint = {
            anchor: [centerX - endRadius, armEndY],
            left:   [centerX - endRadius, armEndY],
            right:  [centerX - (endRadius - endHandle), armEndY]
        };
        /* 直線の開始点（両端の円弧の内端）/ Start of the straight section */
        var straightTopPoint = {
            anchor: [centerX, straightTopY],
            left:   [centerX, straightTopY + endHandle],
            right:  [centerX, straightTopY]
        };
        /* 直線の終了点（中央の円弧の外端）/ End of the straight section */
        var straightBottomPoint = {
            anchor: [centerX, straightBottomY],
            left:   [centerX, straightBottomY],
            right:  [centerX, straightBottomY - centerHandle]
        };

        var halfPoints = [];

        /* 延長線（中央とは反対の外側へ伸ばす）/ Extension, running outward, away from the middle */
        if (extension > 0) {
            halfPoints.push(createStraightPoint(centerX - endRadius - extension, armEndY));
        }

        /* 半径0のときは腕の先と直線の開始点が重なり、円弧のない直角になる / A zero radius merges the arm end into the straight start, leaving a right angle */
        var armPoint = (endRadius > 0) ? armEndPoint : mergePathPoints(armEndPoint, straightTopPoint);
        halfPoints.push(armPoint);

        if (straightLength > 0) {
            if (endRadius > 0) halfPoints.push(straightTopPoint);
            halfPoints.push(straightBottomPoint);
        } else if (endRadius > 0) {
            /* 直線がないときは変曲点ひとつにまとめる / Without a straight section the two points merge into an inflection point */
            halfPoints.push(mergePathPoints(straightTopPoint, straightBottomPoint));
        } else {
            /* 直線も円弧もないときは、腕の先まで含めてひとつにまとめる / With neither, the arm end merges in as well */
            halfPoints[halfPoints.length - 1] = mergePathPoints(armPoint, straightBottomPoint);
        }
        return halfPoints;
    }

    /**
     * ブラケットのアンカーポイントと方向線を計算する
     * 上半分だけを組み立て、中央の突起をはさんで鏡像を並べ、最後に向きへ回転する
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {number} centerX - 配置位置の中心X（pt）
     * @param {number} centerY - 配置位置の中心Y（pt）
     * @returns {Array<{anchor: number[], left: number[], right: number[]}>} 向きを反映したパスポイント
     */
    function buildBracketPoints(bracketSettings, centerX, centerY) {
        var endRadius = bracketSettings.endRadiusPt;
        var centerRadius = bracketSettings.centerRadiusPt;
        var centerHandle = centerRadius * KAPPA;
        var extension = bracketSettings.extensionPt;

        /* 全体の長さから円弧ぶん（片側 両端＋中央）を差し引いた直線部分 / The straight run left after the arcs (end + center per side) */
        var straightLength = Math.max(0, bracketSettings.totalLengthPt / 2 - endRadius - centerRadius);

        /* 中央のずらし量。全長を保つため、直線が尽きるところで止める / Offset of the middle point, capped where the straight run runs out so the overall length holds */
        var centerOffset = getCenterOffsetSign(bracketSettings.direction) * bracketSettings.centerOffsetPt;
        centerOffset = Math.max(-straightLength, Math.min(straightLength, centerOffset));

        /* ずらしたぶん、上下で直線の長さが変わる / The offset makes the two halves differ */
        var beakY = centerY + centerOffset;
        var upperStraightLength = straightLength - centerOffset;
        var lowerStraightLength = straightLength + centerOffset;
        var upperPoints = buildHalfPoints(endRadius, centerRadius, upperStraightLength, extension, centerX, beakY);
        var lowerPoints = buildHalfPoints(endRadius, centerRadius, lowerStraightLength, extension, centerX, beakY);

        var bracketPoints = upperPoints.slice();
        var lowerCount = lowerPoints.length;
        var i;

        if (centerRadius > 0) {
            /* 中央の突起は上下の折り返し点なので反転しない / The middle point is where the halves meet, so it is not mirrored */
            bracketPoints.push({
                anchor: [centerX + centerRadius, beakY],
                left:   [centerX + (centerRadius - centerHandle), beakY],
                right:  [centerX + (centerRadius - centerHandle), beakY]
            });
        } else {
            /* 半径0のときは突起がなくなり、上下の内端が中央で重なる / A zero radius drops the point, leaving the two inner ends stacked in the middle */
            lowerCount--;
            var innerPoint = mergePathPoints(bracketPoints.pop(), mirrorPointAcrossCenterY(lowerPoints[lowerCount], beakY));
            /* 上下とも直線が残っていれば、まとめた点は直線の途中なので省く / With a straight run on both sides the merged point falls mid-line, so it is left out */
            if (upperStraightLength <= 0 || lowerStraightLength <= 0) bracketPoints.push(innerPoint);
        }

        /* 下半分は中央で折り返した鏡像を逆順に並べる / The lower half is mirrored across the middle point, in reverse order */
        for (i = lowerCount - 1; i >= 0; i--) {
            bracketPoints.push(mirrorPointAcrossCenterY(lowerPoints[i], beakY));
        }

        /* 右向きで組み立てた座標を、選んだ向きへまとめて回転 / Rotate the right-facing points into the chosen direction */
        var rotation = getDirectionRotation(bracketSettings.direction);
        for (i = 0; i < bracketPoints.length; i++) {
            bracketPoints[i].anchor = rotateAroundCenter(bracketPoints[i].anchor, rotation, centerX, centerY);
            bracketPoints[i].left = rotateAroundCenter(bracketPoints[i].left, rotation, centerX, centerY);
            bracketPoints[i].right = rotateAroundCenter(bracketPoints[i].right, rotation, centerX, centerY);
        }
        return bracketPoints;
    }

    /**
     * @typedef {object} SelectionReference
     * @property {number[]} bounds - 選択全体の外接矩形 [左, 上, 右, 下]
     * @property {PathItem|null} sizeReferenceItem - 大きさの参照にした開いたパス（囲む対象のときは null）
     * @property {string|null} bracketDirection - このスクリプトで作ったブラケットなら、その向き
     */

    /**
     * このスクリプトで作ったブラケットかどうかを目印のタグで判定する
     * @param {PathItem|null} pathItem - 調べるパス
     * @returns {string|null} 記録されていた向き（ブラケットでなければ null）
     */
    function readBracketDirection(pathItem) {
        if (!pathItem) return null;
        for (var i = 0; i < pathItem.tags.length; i++) {
            if (pathItem.tags[i].name === BRACKET_TAG_NAME) {
                return DIRECTION_ROTATIONS[pathItem.tags[i].value] ? pathItem.tags[i].value : null;
            }
        }
        return null;
    }

    /**
     * 選択オブジェクトを基準情報として読み取る
     * 開いたパス1本だけなら「大きさの参照」、それ以外は「囲む対象」として扱う
     * @param {Document} doc - 対象ドキュメント
     * @returns {SelectionReference|null} 基準の情報（境界を持つ選択がなければ null）
     */
    function readSelectionReference(doc) {
        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0) return null;

        var selectionBounds = null;
        var boundedItemCount = 0;
        var firstBoundedItem = null;

        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = null;
            /* 文字選択（TextRange）には geometricBounds がないので、取れないものは飛ばす / A text range has no geometricBounds, so skip what cannot be read */
            try {
                itemBounds = selectedItems[i].geometricBounds;
            } catch (eBounds) {
                itemBounds = null;
            }
            if (!itemBounds) continue;

            boundedItemCount++;
            if (!firstBoundedItem) firstBoundedItem = selectedItems[i];

            if (!selectionBounds) {
                selectionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
            } else {
                selectionBounds[0] = Math.min(selectionBounds[0], itemBounds[0]);
                selectionBounds[1] = Math.max(selectionBounds[1], itemBounds[1]);
                selectionBounds[2] = Math.max(selectionBounds[2], itemBounds[2]);
                selectionBounds[3] = Math.min(selectionBounds[3], itemBounds[3]);
            }
        }
        if (!selectionBounds) return null;

        /* 閉じたパスや図形・テキストは囲む対象なので参照にしない / Closed paths, shapes and text are things to bracket, not references */
        var isSizeReference = (boundedItemCount === 1 && firstBoundedItem.typename === "PathItem" && !firstBoundedItem.closed);
        var sizeReferenceItem = isSizeReference ? firstBoundedItem : null;
        return {
            bounds: selectionBounds,
            sizeReferenceItem: sizeReferenceItem,
            bracketDirection: readBracketDirection(sizeReferenceItem)
        };
    }

    /**
     * 大きさの参照にしたパスの表示・非表示を切り替える（囲む対象のときは何もしない）
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @param {boolean} isHidden - 隠すなら true、戻すなら false
     * @returns {void}
     */
    function setSizeReferenceHidden(selectionReference, isHidden) {
        if (!selectionReference || !selectionReference.sizeReferenceItem) return;
        selectionReference.sizeReferenceItem.hidden = isHidden;
    }

    /**
     * 大きさの参照にしたパスを削除する（囲む対象のときは何もしない）
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @returns {void}
     */
    function removeSizeReference(selectionReference) {
        if (!selectionReference || !selectionReference.sizeReferenceItem) return;
        selectionReference.sizeReferenceItem.remove();
    }

    /**
     * 向きに合わせて、基準の外接矩形から長さを取る
     * 左右の向きなら高さ、上下の向きなら幅
     * @param {SelectionReference} selectionReference - 基準の情報
     * @param {string} direction - 向きのキー
     * @returns {number} 全体の長さ（mm）
     */
    function getReferenceLengthMm(selectionReference, direction) {
        var referenceBounds = selectionReference.bounds;
        var isVertical = (direction === "left" || direction === "right");
        var referenceLengthPt = isVertical
            ? (referenceBounds[1] - referenceBounds[3])
            : (referenceBounds[2] - referenceBounds[0]);
        return referenceLengthPt / POINTS_PER_MM;
    }

    /**
     * ダイアログの初期値（長さ・半径・向き）を決める
     * 選択があれば選択から決め直し、選択がなければ前回ダイアログを閉じたときの値を復元する
     * 参照のパスは左（縦長）・上（横長）に、囲む対象は右（縦長）・下（横長）に置く
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @param {object} savedSettings - 同一セッション内に記憶していた値
     * @returns {{totalLengthMm: number, centerRadiusMm: number, endRadiusMm: number, direction: string}} 初期値
     */
    function getInitialValues(selectionReference, savedSettings) {
        if (!selectionReference) {
            /* 選択がないときは前回の値をそのまま復元する / With no selection, restore what was there last time */
            var savedLengthMm = readSavedNumber(savedSettings.totalLength, DEFAULT_TOTAL_LENGTH_MM);
            var savedRadiusMm = savedLengthMm * DEFAULT_RADIUS_RATIO;
            return {
                totalLengthMm: savedLengthMm,
                centerRadiusMm: readSavedNumber(savedSettings.centerRadius, savedRadiusMm),
                endRadiusMm: readSavedNumber(savedSettings.endRadius, savedRadiusMm),
                direction: DIRECTION_ROTATIONS[savedSettings.direction] ? savedSettings.direction : DEFAULT_DIRECTION
            };
        }

        var referenceBounds = selectionReference.bounds;
        var referenceWidth = referenceBounds[2] - referenceBounds[0];
        var referenceHeight = referenceBounds[1] - referenceBounds[3];
        var isSizeReference = (selectionReference.sizeReferenceItem !== null);
        var direction;

        if (selectionReference.bracketDirection) {
            /* このスクリプトで作ったブラケットは、その向きのまま描き直す / A bracket we made keeps the direction it was drawn with */
            direction = selectionReference.bracketDirection;
        } else if (referenceHeight >= referenceWidth) {
            direction = isSizeReference ? "left" : "right";
        } else {
            direction = isSizeReference ? "up" : "down";
        }

        var totalLengthMm = getReferenceLengthMm(selectionReference, direction);
        /* 半径（中央・両端とも）は全体の長さに追従させる（以後は個別に変更できる）/ Both radii follow the overall length, and can be changed on their own afterwards */
        var radiusMm = totalLengthMm * DEFAULT_RADIUS_RATIO;

        return {
            totalLengthMm: totalLengthMm,
            centerRadiusMm: radiusMm,
            endRadiusMm: radiusMm,
            direction: direction
        };
    }

    /**
     * アクティブアートボードの中心座標を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {{x: number, y: number}} 中心座標（pt）
     */
    function getActiveArtboardCenter(doc) {
        var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var artboardRect = activeArtboard.artboardRect; /* [左, 上, 右, 下] / [left, top, right, bottom] */
        return {
            x: (artboardRect[0] + artboardRect[2]) / 2,
            y: (artboardRect[1] + artboardRect[3]) / 2
        };
    }

    /**
     * ブラケットを置く中心座標を返す
     * このスクリプトで作ったブラケットを選んでいれば、その中央に重ねる
     * それ以外の選択では、腕の先が外接矩形の辺に触れる位置に置く
     * @param {Document} doc - 対象ドキュメント
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @returns {{x: number, y: number}} 中心座標（pt）
     */
    function getPlacementCenter(doc, bracketSettings, selectionReference) {
        if (!selectionReference) return getActiveArtboardCenter(doc);

        var referenceBounds = selectionReference.bounds;
        var rotation = getDirectionRotation(bracketSettings.direction);
        var referenceCenterX = (referenceBounds[0] + referenceBounds[2]) / 2;
        var referenceCenterY = (referenceBounds[1] + referenceBounds[3]) / 2;

        if (selectionReference.bracketDirection) {
            /* 突起と腕で伸び方が違うぶんを戻し、見た目の中央を選択に重ねる / Undo the asymmetry between the point and the arms so the visual center lands on the selection */
            var shapeOffset = (bracketSettings.centerRadiusPt - bracketSettings.endRadiusPt - bracketSettings.extensionPt) / 2;
            return {
                x: referenceCenterX - rotation.cos * shapeOffset,
                y: referenceCenterY - rotation.sin * shapeOffset
            };
        }

        /* 中央の突起の向きへ、外接矩形の半分＋腕の長さ＋余白ぶん寄せると腕の先が辺に触れる / Shifting along the facing direction by half the bounds plus the arm and the gap puts the arm ends on the edge */
        var halfWidth = (referenceBounds[2] - referenceBounds[0]) / 2;
        var halfHeight = (referenceBounds[1] - referenceBounds[3]) / 2;
        var halfSize = Math.abs(rotation.cos) * halfWidth + Math.abs(rotation.sin) * halfHeight;
        var placementOffset = halfSize + bracketSettings.endRadiusPt + bracketSettings.extensionPt + bracketSettings.objectGapPt;

        return {
            x: referenceCenterX + rotation.cos * placementOffset,
            y: referenceCenterY + rotation.sin * placementOffset
        };
    }

    /**
     * ブラケットのパスを生成して配置する
     * @param {Document} doc - 対象ドキュメント
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {SelectionReference|null} selectionReference - 基準の情報（null ならアートボード中央）
     * @returns {PathItem} 生成したパス
     */
    function createBracketPath(doc, bracketSettings, selectionReference) {
        var placementCenter = getPlacementCenter(doc, bracketSettings, selectionReference);
        var bracketPoints = buildBracketPoints(bracketSettings, placementCenter.x, placementCenter.y);

        var bracketPath = doc.pathItems.add();
        bracketPath.closed = false;
        bracketPath.filled = false;
        bracketPath.stroked = true;
        bracketPath.strokeWidth = bracketSettings.strokeWidthPt;
        bracketPath.strokeCap = STROKE_CAPS[bracketSettings.strokeCap] || STROKE_CAPS[DEFAULT_STROKE_CAP];
        bracketPath.strokeJoin = CORNER_JOINS[bracketSettings.cornerJoin] || CORNER_JOINS[DEFAULT_CORNER_JOIN];

        var bracketStrokeColor = new GrayColor();
        bracketStrokeColor.gray = 100;
        bracketPath.strokeColor = bracketStrokeColor;

        for (var i = 0; i < bracketPoints.length; i++) {
            var pathPoint = bracketPath.pathPoints.add();
            pathPoint.anchor = bracketPoints[i].anchor;
            pathPoint.leftDirection = bracketPoints[i].left;
            pathPoint.rightDirection = bracketPoints[i].right;
            pathPoint.pointType = PointType.CORNER;
        }

        /* 面取りは円弧を直線でつなぐジグザグ効果で表現する。形が決まってから適用する / The chamfer is a Zig Zag effect that straightens the arcs; apply it once the shape exists */
        if (bracketSettings.isChamfer) {
            bracketPath.applyEffect(CHAMFER_EFFECT_XML);
        }

        /* ブラシは形が決まってから適用する / Apply the brush once the shape exists */
        if (bracketSettings.brushName) {
            applyBrush(doc, bracketPath, bracketSettings.brushName);
        }

        /* 次に選び直したとき同じ位置・向きで描き直せるよう目印を残す / Leave a marker so a later run can redraw it in place */
        var bracketTag = bracketPath.tags.add();
        bracketTag.name = BRACKET_TAG_NAME;
        bracketTag.value = bracketSettings.direction;

        return bracketPath;
    }

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
    // セッション記憶 / Session memory
    // =========================================

    /* ダイアログの状態は Illustrator を終了すると消える（#targetengine が必須）。旧版の $.global のキーを1度だけ読み継ぐ
       Dialog state is gone once Illustrator quits (needs #targetengine); the old $.global key is read once */
    var LEGACY_SESSION_SETTINGS_KEY = "aiCurlyBracketMakerSettings";
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_SETTINGS_KEY] || null; }
    });

    /**
     * 前回ダイアログを閉じたときの状態を読み出す。
     * 項目ごとの型は readSaved*() が確かめるため、既定値は {}（中身を問わず受け取る）
     * @returns {object} 記憶していた値（無ければ空オブジェクト。毎回新しい写し）
     */
    function loadSessionSettings() {
        return settingsStore.load({});
    }

    /**
     * ダイアログの状態を同一セッション内だけ覚える
     * @param {object} settings - 記憶する値
     * @returns {void}
     */
    function saveSessionSettings(settings) {
        settingsStore.save(settings);
    }

    /**
     * 記憶していた値を数値として読む
     * @param {*} savedValue - 記憶していた値
     * @param {number} fallbackValue - 読めないときに使う値
     * @returns {number} 数値
     */
    function readSavedNumber(savedValue, fallbackValue) {
        var value = Number(savedValue);
        return (savedValue === undefined || savedValue === "" || isNaN(value)) ? fallbackValue : value;
    }

    /**
     * 記憶していた値を入力欄の文字列として読む
     * @param {*} savedValue - 記憶していた値
     * @param {string} fallbackText - 読めないときに使う文字列
     * @returns {string} 入力欄に入れる文字列
     */
    function readSavedText(savedValue, fallbackText) {
        return (savedValue === undefined || savedValue === "") ? fallbackText : String(savedValue);
    }

    /**
     * 記憶していた値をチェックボックスの状態として読む
     * @param {*} savedValue - 記憶していた値
     * @param {boolean} fallbackFlag - 読めないときに使う状態
     * @returns {boolean} チェック状態
     */
    function readSavedFlag(savedValue, fallbackFlag) {
        return (savedValue === undefined) ? fallbackFlag : (savedValue === true);
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
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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
     * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function centerButtonRowIfRightOnly(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ブラケット作成ダイアログを構築する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Window} 構築済みのダイアログ
     */
    function createBracketDialog(doc) {
        /* 現在描画中のプレビューパス / The preview path currently on the artboard */
        var previewPath = null;
        /* ［作成］で確定したか（確定後はプレビューを消さない）/ Whether Create committed the path */
        var bracketCommitted = false;

        /* 選択があれば、その大きさと向きを初期値にして辺に沿わせる / A selection seeds the size and direction, and the bracket hugs its edge */
        var selectionReference = readSelectionReference(doc);
        /* 同一セッション内に覚えていた前回の状態 / What the previous run left behind, within this session */
        var savedSettings = loadSessionSettings();
        var initialValues = getInitialValues(selectionReference, savedSettings);
        /* ライブラリーを開く処理は、参照パスを隠す前に済ませる（ここで失敗しても隠したままにならない）
           Read the library before anything is hidden, so a failure here cannot leave the reference path hidden */
        var libraryBrushNames = getLibraryBrushNames(doc);

        /* 参照にしたパスは開いた時点で隠す（確定で削除、キャンセルで元に戻す）/ Hide the reference path as the dialog opens: removed on Create, restored on Cancel */
        setSizeReferenceHidden(selectionReference, true);

        var bracketDialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        setupWindow(bracketDialog);

        var shapePanel = addPanel(bracketDialog, getLabel('panel.shapeAndSize'));

        /* 半径は2行を1つの列にまとめ、連動をその右に天地中央で添える / Both radius rows share a column, with the link centered beside them */
        var radiusRow = shapePanel.add("group");
        setupRow(radiusRow, "left", FIELD_ROW_SPACING);

        var radiusColumn = radiusRow.add("group");
        radiusColumn.orientation = "column";
        radiusColumn.alignChildren = ["left", "center"];
        radiusColumn.spacing = PANEL_SPACING;

        var centerRadiusInput = addNumberFieldRow(
            radiusColumn, labelText('fieldLabel.centerRadius'),
            formatFieldValue(initialValues.centerRadiusMm), "mm", getLabel('tooltip.centerRadius')
        );
        var endRadiusInput = addNumberFieldRow(
            radiusColumn, labelText('fieldLabel.endRadius'),
            formatFieldValue(initialValues.endRadiusMm), "mm", getLabel('tooltip.endRadius')
        );

        var linkRadiusToggle = addLinkToggle(radiusRow, readSavedFlag(savedSettings.linkRadius, DEFAULT_LINK_RADIUS), onLinkRadiusToggle);
        linkRadiusToggle.alignment = ["left", "center"];
        linkRadiusToggle.helpTip = getLabel('tooltip.linkRadius');

        /* 面取り：ラベルは空のまま、字下げして数値欄の列にそろえる / Chamfer: an empty label keeps it aligned with the fields */
        var chamferRow = shapePanel.add("group");
        setupRow(chamferRow, "left", FIELD_ROW_SPACING);
        addRowLabel(chamferRow, "");
        var chamferCheckbox = chamferRow.add("checkbox", undefined, getLabel('checkbox.chamfer'));
        chamferCheckbox.value = readSavedFlag(savedSettings.chamfer, DEFAULT_CHAMFER);
        chamferCheckbox.helpTip = getLabel('tooltip.chamfer');

        /* 向き：ラベル幅を数値欄の行にそろえ、上下・左右の2行に分けて並べる / Direction: share the label width, split into up-down and left-right rows */
        var directionRow = shapePanel.add("group");
        setupRow(directionRow, "left", FIELD_ROW_SPACING);
        /* ラベルを1行目のラジオに合わせる / Align the label with the first row of radios */
        directionRow.alignChildren = ["left", "top"];
        addRowLabel(directionRow, labelText('fieldLabel.direction'));

        var directionColumn = directionRow.add("group");
        directionColumn.orientation = "column";
        directionColumn.alignChildren = ["left", "center"];
        directionColumn.spacing = FIELD_ROW_SPACING;

        /* キー順（上・下／左・右）に2つずつ並べる / Two per row, in key order */
        var directionRadios = {};
        for (var i = 0; i < DIRECTION_KEYS.length; i += 2) {
            var directionGridRow = directionColumn.add("group");
            setupRow(directionGridRow, "left", FIELD_ROW_SPACING);
            for (var j = i; j < i + 2 && j < DIRECTION_KEYS.length; j++) {
                var directionKey = DIRECTION_KEYS[j];
                var directionRadio = directionGridRow.add("radiobutton", undefined, getLabel("radio." + directionKey));
                directionRadio.helpTip = getLabel('tooltip.direction');
                directionRadios[directionKey] = directionRadio;
            }
        }

        var centerOffsetInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel.centerOffset'),
            readSavedText(savedSettings.centerOffset, String(DEFAULT_CENTER_OFFSET_MM)), "mm", getLabel('tooltip.centerOffset')
        );

        var totalLengthInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel.totalLength'),
            formatFieldValue(initialValues.totalLengthMm), "mm", getLabel('tooltip.totalLength')
        );
        var extensionInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel.extension'),
            readSavedText(savedSettings.extension, String(DEFAULT_EXTENSION_PT)), "pt", getLabel('tooltip.extension')
        );
        var objectGapInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel.objectGap'),
            readSavedText(savedSettings.objectGap, String(DEFAULT_OBJECT_GAP_MM)), "mm", getLabel('tooltip.objectGap')
        );
        /* 余白が効くのは「囲む対象」があるときだけなので、それ以外は行ごとディム表示 / The gap only applies to something being bracketed, so dim the row otherwise */
        objectGapInput.fieldRow.enabled = (selectionReference !== null && !selectionReference.bracketDirection);
        redrawSteppersIn(objectGapInput.fieldRow);

        var strokePanel = addPanel(bracketDialog, getLabel('panel.stroke'));

        var strokeWidthInput = addNumberFieldRow(
            strokePanel, labelText('fieldLabel.strokeWidth'),
            readSavedText(savedSettings.strokeWidth, String(DEFAULT_STROKE_WIDTH_PT)), "pt", getLabel('tooltip.strokeWidth')
        );

        /* 線端：同じグループ内なのでScriptUIが排他にしてくれる / Stroke cap: one group, so ScriptUI keeps the radios exclusive */
        var strokeCapRow = strokePanel.add("group");
        setupRow(strokeCapRow, "left", FIELD_ROW_SPACING);
        addRowLabel(strokeCapRow, labelText('fieldLabel.strokeCap'));
        var buttCapRadio = strokeCapRow.add("radiobutton", undefined, getLabel('radio.buttCap'));
        var roundCapRadio = strokeCapRow.add("radiobutton", undefined, getLabel('radio.roundCap'));
        var initialStrokeCap = STROKE_CAPS[savedSettings.strokeCap] ? savedSettings.strokeCap : DEFAULT_STROKE_CAP;
        buttCapRadio.value = (initialStrokeCap === "buttCap");
        roundCapRadio.value = !buttCapRadio.value;
        buttCapRadio.helpTip = getLabel('tooltip.strokeCap');
        roundCapRadio.helpTip = getLabel('tooltip.strokeCap');

        /* 角の形状：ラジオは縦並び。同じグループ内なのでScriptUIが排他にしてくれる / Corner shape: radios stacked in one group, so ScriptUI keeps them exclusive */
        var cornerJoinRow = strokePanel.add("group");
        setupRow(cornerJoinRow, "left", FIELD_ROW_SPACING);
        /* ラベルを1行目のラジオに合わせる（天地中央だと2行ぶんの中央に落ちる）/ Align the label with the first radio instead of centering it over both rows */
        cornerJoinRow.alignChildren = ["left", "top"];
        addRowLabel(cornerJoinRow, labelText('fieldLabel.cornerJoin'));

        var cornerJoinColumn = cornerJoinRow.add("group");
        cornerJoinColumn.orientation = "column";
        cornerJoinColumn.alignChildren = ["left", "center"];
        cornerJoinColumn.spacing = FIELD_ROW_SPACING;
        var miterJoinRadio = cornerJoinColumn.add("radiobutton", undefined, getLabel('radio.miterJoin'));
        var roundJoinRadio = cornerJoinColumn.add("radiobutton", undefined, getLabel('radio.roundJoin'));
        var initialCornerJoin = CORNER_JOINS[savedSettings.cornerJoin] ? savedSettings.cornerJoin : DEFAULT_CORNER_JOIN;
        miterJoinRadio.value = (initialCornerJoin === "miterJoin");
        roundJoinRadio.value = !miterJoinRadio.value;
        miterJoinRadio.helpTip = getLabel('tooltip.cornerJoin');
        roundJoinRadio.helpTip = getLabel('tooltip.cornerJoin');

        /* ブラシ：ライブラリーに出せるブラシがなければ、行ごと作らない / Brush: without a library brush to offer, the row is not built at all */
        var brushDropdown = null;
        if (libraryBrushNames.length > 0) {
            var brushRow = strokePanel.add("group");
            setupRow(brushRow, "fill", FIELD_ROW_SPACING);
            addRowLabel(brushRow, labelText('fieldLabel.brush'));

            /* ライブラリーのブラシを「なし」に続けてそのまま並べる / "None" followed by the library's brushes, as it lists them */
            var brushItems = [getLabel('dropdown.noBrush')].concat(libraryBrushNames);
            brushDropdown = brushRow.add("dropdownlist", undefined, brushItems);
            brushDropdown.helpTip = getLabel('tooltip.brush');
            /* 名前が長くてもダイアログを広げない / A long brush name must not widen the dialog */
            brushDropdown.preferredSize.width = DROPDOWN_WIDTH;
            brushDropdown.alignment = ["fill", "center"];

            /* 覚えていたブラシがライブラリーから消えていることもあるので、無ければ「なし」に戻す / A remembered brush may be gone from the library, so fall back to "none" */
            var initialBrushName = readSavedText(savedSettings.brush, DEFAULT_BRUSH);
            brushDropdown.selection = 0;
            for (var i = 1; i < brushItems.length; i++) {
                if (brushItems[i] === initialBrushName) brushDropdown.selection = i;
            }
        }

        /**
         * 選択中のブラシ名を返す
         * @returns {string|null} ブラシ名（「なし」またはブラシの行がなければ null）
         */
        function readSelectedBrushName() {
            if (!brushDropdown) return null;
            var selectedItem = brushDropdown.selection;
            return (selectedItem && selectedItem.index > 0) ? selectedItem.text : null;
        }

        /**
         * 選択中の向きを返す
         * @returns {string} DIRECTION_KEYS のいずれか（未選択なら DEFAULT_DIRECTION）
         */
        function readSelectedDirection() {
            for (var i = 0; i < DIRECTION_KEYS.length; i++) {
                if (directionRadios[DIRECTION_KEYS[i]].value) return DIRECTION_KEYS[i];
            }
            return DEFAULT_DIRECTION;
        }

        /**
         * 入力欄からブラケットの設定値を読み取る
         * @returns {BracketSettings|null} 設定値（数値として読めない欄があれば null）
         */
        function readBracketSettings() {
            var endRadiusMm = Number(endRadiusInput.text);
            var centerRadiusMm = Number(centerRadiusInput.text);
            var totalLengthMm = Number(totalLengthInput.text);
            var extensionPt = Number(extensionInput.text);
            var centerOffsetMm = Number(centerOffsetInput.text);
            var objectGapMm = Number(objectGapInput.text);
            var strokeWidthPt = Number(strokeWidthInput.text);

            if (isNaN(endRadiusMm) || endRadiusMm < 0) return null;
            if (isNaN(centerRadiusMm) || centerRadiusMm < 0) return null;
            if (isNaN(totalLengthMm) || totalLengthMm < 0) return null;
            if (isNaN(extensionPt) || extensionPt < 0) return null;
            if (isNaN(centerOffsetMm)) return null;
            if (isNaN(objectGapMm) || objectGapMm < 0) return null;
            if (isNaN(strokeWidthPt) || strokeWidthPt <= 0) return null;

            return {
                endRadiusPt: endRadiusMm * POINTS_PER_MM,
                centerRadiusPt: centerRadiusMm * POINTS_PER_MM,
                totalLengthPt: totalLengthMm * POINTS_PER_MM,
                centerOffsetPt: centerOffsetMm * POINTS_PER_MM,
                extensionPt: extensionPt,
                objectGapPt: objectGapMm * POINTS_PER_MM,
                strokeWidthPt: strokeWidthPt,
                strokeCap: roundCapRadio.value ? "roundCap" : "buttCap",
                cornerJoin: roundJoinRadio.value ? "roundJoin" : "miterJoin",
                brushName: readSelectedBrushName(),
                direction: readSelectedDirection(),
                isChamfer: chamferCheckbox.value
            };
        }

        /**
         * プレビューのパスを削除する
         * @returns {void}
         */
        function removePreview() {
            if (!previewPath) return;
            previewPath.remove();
            previewPath = null;
        }

        /**
         * 入力値でプレビューを描き直す
         * @returns {void}
         */
        function drawPreview() {
            var bracketSettings = readBracketSettings();
            /* 入力途中で読めないときは直前のプレビューを残す / Keep the last preview while the input is unreadable */
            if (!bracketSettings) return;

            removePreview();
            previewPath = createBracketPath(doc, bracketSettings, selectionReference);
            app.redraw();
        }

        /**
         * 連動がONのとき、両端の半径を中央に合わせる
         * @returns {void}
         */
        function syncLinkedRadius() {
            if (!linkRadiusToggle.value) return;
            if (endRadiusInput.text === centerRadiusInput.text) return;
            endRadiusInput.text = centerRadiusInput.text;
        }

        /**
         * 連動の状態を両端の行に反映する（連動は行の外にあるので操作できるまま残る）
         * @returns {void}
         */
        function updateEndRadiusEnabled() {
            endRadiusInput.fieldRow.enabled = !linkRadiusToggle.value;
            redrawSteppersIn(endRadiusInput.fieldRow);
        }

        /**
         * 数値欄にプレビュー更新（キー入力・∧∨・↑↓キー）を割り当てる
         * @param {EditText} numberField - addNumberFieldRow() で作った入力欄
         * @param {function} onChanged - 値が変わったときの処理
         * @param {number} [minValue] - ∧∨・↑↓キーでの下限（省略時は制限なし）
         * @returns {void}
         */
        function bindPreviewField(numberField, onChanged, minValue) {
            numberField.addEventListener("changing", onChanged);
            numberField.stepOptions.min = minValue;
            numberField.stepOptions.onStep = function () { onChanged(); };
        }

        /* 中央は連動の書き写しをはさむ。ほかの数値欄はそのままプレビュー更新 / The center radius syncs first; the other fields just refresh the preview */
        bindPreviewField(centerRadiusInput, function () {
            syncLinkedRadius();
            drawPreview();
        }, 0);
        bindPreviewField(endRadiusInput, drawPreview, 0);

        /* 中央の位置だけはマイナスを許す / Only the center offset may go negative */
        bindPreviewField(centerOffsetInput, drawPreview);

        var previewNumberFields = [totalLengthInput, extensionInput, objectGapInput, strokeWidthInput];
        for (var i = 0; i < previewNumberFields.length; i++) {
            bindPreviewField(previewNumberFields[i], drawPreview, 0);
        }

        /* 連動をONにした時点で、両端を中央の値にそろえてディム表示にする / Turning the link on snaps the end radius to the center radius and dims it */
        /**
         * 連動アイコンを切り替えたときの処理（addLinkToggle() に渡す）
         * @returns {void}
         */
        function onLinkRadiusToggle() {
            updateEndRadiusEnabled();
            syncLinkedRadius();
            drawPreview();
        }
        /**
         * 線端に角の形状をそろえるクリック処理を作る（丸型→ラウンド結合、なし→マイター結合）
         * @param {boolean} isRoundCap - 丸型を選んだときの処理なら true
         * @returns {function} クリックハンドラ
         */
        function createStrokeCapClickHandler(isRoundCap) {
            return function () {
                roundJoinRadio.value = isRoundCap;
                miterJoinRadio.value = !isRoundCap;
                drawPreview();
            };
        }
        buttCapRadio.onClick = createStrokeCapClickHandler(false);
        roundCapRadio.onClick = createStrokeCapClickHandler(true);
        chamferCheckbox.onClick = drawPreview;
        miterJoinRadio.onClick = drawPreview;
        roundJoinRadio.onClick = drawPreview;
        if (brushDropdown) brushDropdown.onChange = drawPreview;
        /**
         * 指定した向きだけを選択状態にする（行をまたぐラジオはScriptUIでは排他にならない）
         * @param {string} selectedKey - 選択する向きのキー
         * @returns {void}
         */
        function selectDirection(selectedKey) {
            for (var i = 0; i < DIRECTION_KEYS.length; i++) {
                directionRadios[DIRECTION_KEYS[i]].value = (DIRECTION_KEYS[i] === selectedKey);
            }
        }

        /**
         * 向きラジオのクリック処理を作る（ループ変数を閉じ込める）
         * @param {string} directionKey - 対象の向きのキー
         * @returns {function} クリックハンドラ
         */
        function createDirectionClickHandler(directionKey) {
            return function () {
                selectDirection(directionKey);
                /* 選択があるときは、新しい向きに合わせて長さを取り直す / With a selection, re-read the length across the new direction */
                if (selectionReference) {
                    totalLengthInput.text = formatFieldValue(getReferenceLengthMm(selectionReference, directionKey));
                }
                drawPreview();
            };
        }
        for (var i = 0; i < DIRECTION_KEYS.length; i++) {
            directionRadios[DIRECTION_KEYS[i]].onClick = createDirectionClickHandler(DIRECTION_KEYS[i]);
        }

        /* ボタンエリア（右：キャンセル・作成）/ Button row: Cancel and Create on the right */
        var buttonRow = addButtonRow(bracketDialog);
        /* キャンセルは既定動作で閉じ、後片付けは bracketDialog.onClose が行う / Cancel closes by default; cleanup happens in bracketDialog.onClose */
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        btnCancel.helpTip = getLabel('tooltip.cancel');
        var btnCreate = buttonRow.rightGroup.add("button", undefined, getLabel('button.create'), { name: "ok" });
        btnCreate.helpTip = getLabel('tooltip.create');

        /* ［作成］：プレビューをそのまま成果物として残す / Create: keep the preview as the result */
        btnCreate.onClick = function () {
            var bracketSettings = readBracketSettings();
            if (!bracketSettings) {
                alert(getLabel('alert.invalidValue'));
                return;
            }
            /* 入力途中で止まったプレビューを最新の値にそろえる / Bring the preview up to date with the current values */
            removePreview();
            previewPath = createBracketPath(doc, bracketSettings, selectionReference);
            /* 隠しておいた参照のパスは役目を終えたので削除 / The hidden reference path has done its job, so remove it */
            removeSizeReference(selectionReference);

            /* 実行後は作成したブラケットだけを選択状態にする / Leave only the new bracket selected */
            doc.selection = null;
            previewPath.selected = true;

            bracketCommitted = true;
            previewPath = null;
            bracketDialog.close();
        };

        /* ダイアログを閉じたら後片付け（キャンセル・ESCも含む）/ Clean up on close (Cancel and ESC included) */
        bracketDialog.onClose = function () {
            if (!bracketCommitted) {
                removePreview();
                setSizeReferenceHidden(selectionReference, false);
                app.redraw();
            }

            /* 閉じたときの状態を、同一セッション内の次回実行のために覚えておく / Remember the closing state for the next run in this session */
            saveSessionSettings({
                centerRadius: centerRadiusInput.text,
                endRadius: endRadiusInput.text,
                linkRadius: linkRadiusToggle.value,
                chamfer: chamferCheckbox.value,
                direction: readSelectedDirection(),
                centerOffset: centerOffsetInput.text,
                totalLength: totalLengthInput.text,
                extension: extensionInput.text,
                objectGap: objectGapInput.text,
                strokeWidth: strokeWidthInput.text,
                strokeCap: roundCapRadio.value ? "roundCap" : "buttCap",
                cornerJoin: roundJoinRadio.value ? "roundJoin" : "miterJoin",
                /* ブラシの行が無い実行では、覚えていた値をそのまま持ち越す / Without the brush row, carry the remembered value over untouched */
                brush: brushDropdown ? (readSelectedBrushName() || "") : readSavedText(savedSettings.brush, "")
            });
        };

        selectDirection(initialValues.direction);
        updateEndRadiusEnabled();
        centerRadiusInput.active = true;
        drawPreview(); /* 初回プレビュー / First preview */
        return bracketDialog;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * ブラケット作成ダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            app.documents.add();
        }
        var doc = app.activeDocument;

        /* プレビューを描けないレイヤーでは開始しない / Do not start when the preview cannot be drawn */
        var activeLayer = doc.activeLayer;
        if (activeLayer.locked || !activeLayer.visible) {
            alert(getLabel('alert.lockedLayer'));
            return;
        }

        var bracketDialog = createBracketDialog(doc);
        centerButtonRowIfRightOnly(buttonRow);
        prepareDialogWindow(bracketDialog, SCRIPT_NAME);
        bracketDialog.show();
    }

    main();

})();
