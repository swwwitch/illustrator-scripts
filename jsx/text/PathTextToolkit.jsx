#target illustrator
#targetengine "PathTextToolkitEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

パス上文字の作成と調整をまとめて行うツールです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathTextToolkit.md

### Overview

A toolkit for creating and adjusting text on a path.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathTextToolkit.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathTextToolkit";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathTextToolkit.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathTextToolkit.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    var ARC_LABEL_WIDTH       = 70;                /* ［アーチ方向］の項目名の幅 / width of the arc direction label */
    var ARC_SLIDER_WIDTH      = 130;               /* ［まるみ］スライダーの幅 / width of the roundness slider */
    var FIT_PANEL_GAP         = 5;                 /* ［パス幅に合わせる］パネルの上の余白 / gap above the fit-to-path-width panel */
    var POSITION_LABEL_WIDTH  = 60;                /* 開始／終了位置の項目名の幅 / width of the start/end labels */
    var POSITION_FIELD_CHARS  = 5;                 /* 開始／終了位置の入力欄の幅 / width of the start/end fields */
    var ADJUST_LABEL_WIDTH    = 95;                /* テキスト調整の項目名の幅 / width of the text-adjust labels */
    var ADJUST_FIELD_CHARS    = 6;                 /* テキスト調整の入力欄の幅 / width of the text-adjust fields */
    var ROW_SLIDER_WIDTH      = 180;               /* 各行のスライダーの幅 / width of the row sliders */
    var EDIT_HINT_WIDTH       = 360;               /* テキスト編集ダイアログの説明の幅 / width of the text-edit hint */
    var EDIT_FIELD_SIZE       = [320, 80];         /* テキスト編集ダイアログの入力欄 / size of the text-edit field */

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
            title: { ja: "パス上文字（変換と調整）", en: "Text on a Path (Convert & Adjust)" },
            textEditTitle: { ja: "テキスト編集", en: "Edit Text" },
            textEditHint: {
                ja: "内容を編集してOKで反映します。複数選択の場合は全て置換します。",
                en: "Edit the content and press OK to apply. If multiple items are selected, it replaces all."
            }
        },
        panel: {
            process: { ja: "パス上文字にする", en: "Create Path Text" },
            split: { ja: "テキストを分離", en: "Split Text" },
            option: { ja: "オプション", en: "Options" },
            effect: { ja: "効果", en: "Effect" },
            position: { ja: "位置", en: "Position" },
            textAdjust: { ja: "テキスト調整", en: "Text Adjust" },
            fitWidth: { ja: "パス幅に合わせる", en: "Fit to Path Width" }
        },
        radio: {
            toPathText: { ja: "パス上文字にする", en: "Convert to Path Text" },
            genArcPath: { ja: "アーチ状のパスを生成", en: "Generate Arc Path" },
            genCircle: { ja: "正円を生成", en: "Generate Circle" },
            splitKeepFormat: { ja: "書式を保持", en: "Keep Formatting" },
            splitNoFormat: { ja: "書式を保持しない", en: "Do Not Keep Formatting" },
            alignLeft: { ja: "左揃え", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右揃え", en: "Right" },
            alignFullJustify: { ja: "両端揃え", en: "Justify" },
            arcUp: { ja: "上", en: "Up" },
            arcDown: { ja: "下", en: "Down" },
            fitWidthNone: { ja: "しない", en: "None" },
            fitWidthFontSize: { ja: "文字サイズを調整", en: "Adjust Font Size" },
            fitWidthTracking: { ja: "トラッキングを調整", en: "Adjust Tracking" },
            effectRainbow: { ja: "虹", en: "Rainbow" },
            effectDistort: { ja: "歪み", en: "Skew" },
            effectRibbon: { ja: "3D リボン", en: "3D Ribbon" },
            effectStep: { ja: "階段", en: "Stair Step" },
            effectGravity: { ja: "引力", en: "Gravity" }
        },
        checkbox: {
            splitDeletePath: { ja: "パスを削除", en: "Remove Path" },
            reverse: { ja: "内側配置", en: "Inside" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            align: { ja: "行揃え", en: "Alignment" },
            arcDirection: { ja: "アーチ方向", en: "Arc Direction" },
            arcRoundness: { ja: "まるみ", en: "Roundness" },
            startPos: { ja: "開始位置", en: "Start" },
            endPos: { ja: "終了位置", en: "End" },
            baseShift: { ja: "ベースライン", en: "Baseline Shift" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            fontSize: { ja: "文字サイズ", en: "Font Size" }
        },
        button: {
            textEdit: { ja: "テキスト編集", en: "Edit Text" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            textEdit: { ja: "選択したテキストの内容を、別ダイアログで書き換えます。", en: "Opens a separate dialog for rewriting the selected text." },
            toPathText: { ja: "選択したテキストを、一緒に選んだパスの上に流し込みます。", en: "Flows the selected text onto the path selected with it." },
            genArcPath: { ja: "テキストの幅に合わせたアーチ状のパスを作り、その上に流し込みます。", en: "Builds an arc sized to the text and flows the text onto it." },
            genCircle: {
                ja: "テキストの幅を円周とする正円を作り、その上に流し込みます。",
                en: "Builds a circle whose circumference matches the text width and flows the text onto it."
            },
            splitKeepFormat: { ja: "パス上文字を、書式を保ったままテキストとパスに分けます。", en: "Splits path text into text and path, keeping the formatting." },
            splitNoFormat: { ja: "パス上文字を、書式を落としてテキストとパスに分けます。", en: "Splits path text into text and path, dropping the formatting." },
            splitDeletePath: { ja: "分離したあと、元のパスを削除します。", en: "Deletes the original path after the split." },
            reverse: { ja: "文字をパスの反対側（内側）に配置します。", en: "Places the characters on the other side of the path." },
            arcUp: { ja: "上に膨らんだアーチにします。", en: "Bulges the arc upward." },
            arcDown: { ja: "下に膨らんだアーチにします。", en: "Bulges the arc downward." },
            arcRoundness: { ja: "アーチの曲がり具合です。0で直線に近づきます。", en: "How strongly the arc curves. 0 is nearly straight." },
            fitWidthNone: { ja: "文字の大きさも字間もそのままにします。", en: "Leaves both the size and the spacing as they are." },
            fitWidthFontSize: { ja: "文字サイズを変えて、パスの長さいっぱいに収めます。", en: "Changes the font size so the text fills the path." },
            fitWidthTracking: { ja: "字間を変えて、パスの長さいっぱいに収めます。", en: "Changes the tracking so the text fills the path." },
            effectRainbow: { ja: "1文字ずつ色を変えて虹色にします。", en: "Colours each character to make a rainbow." },
            effectDistort: { ja: "1文字ずつ斜めに傾けます。", en: "Skews each character." },
            effectRibbon: { ja: "奥行きのあるリボンのように、1文字ずつ変形します。", en: "Transforms each character like a ribbon with depth." },
            effectStep: { ja: "1文字ずつ高さをずらして階段状にします。", en: "Offsets each character vertically, like stairs." },
            effectGravity: { ja: "中央へ引き寄せられたように、1文字ずつ大きさと位置を変えます。", en: "Varies each character as if pulled toward the centre." },
            startPosEnabled: {
                ja: "開始位置を指定します。オフのときはパスの先頭から始めます。",
                en: "Sets the start position. When off, the text starts at the beginning of the path."
            },
            startPos: { ja: "文字の流し込みを始める位置です。", en: "Where the text begins along the path." },
            endPosEnabled: {
                ja: "終了位置を指定します。オフのときはパスの終端まで使います。",
                en: "Sets the end position. When off, the text runs to the end of the path."
            },
            endPos: { ja: "文字の流し込みを終える位置です。", en: "Where the text ends along the path." },
            align: { ja: "開始位置と終了位置の間での文字の揃え方です。", en: "How the characters are aligned between the start and end positions." },
            baseShift: { ja: "パスから文字を浮かせる量です。負の値で沈みます。", en: "How far the characters sit above the path. Negative values sink below it." },
            tracking: { ja: "文字と文字の間隔です。単位は1/1000em。", en: "Spacing between characters, in 1/1000 em." },
            fontSize: { ja: "元の文字サイズからの増減です。", en: "Change applied to the original font size." },
            preview: { ja: "結果を画面で確認します。キャンセルすると元に戻ります。", en: "Shows the result on the canvas. Cancel restores the original state." },
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
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noText: { ja: "対象のテキストが見つかりません", en: "No target text found." },
            needPath: { ja: "パスを一緒に選択してください", en: "Select a path together." },
            dupPathFail: { ja: "パスの複製に失敗しました", en: "Failed to duplicate the path." },
            arcFail: { ja: "アーチ状のパス生成に失敗しました", en: "Failed to generate an arc path." },
            needPathText: { ja: "パス上文字を選択してください", en: "Select path text." }
        }
    };

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
    // 共通ヘルパー / Shared helpers
    // =========================================

    /**
     * 増減後の処理を、プレビュー更新で DOM の例外が出ても∧∨・↑↓キーの操作が続くよう包む
     * @param {Function} onStepped - 値を変えたあとに呼ぶ処理
     * @returns {Function} stepOptions.onStep に渡す関数
     */
    function makeSafeStepHandler(onStepped) {
        return function () {
            try { onStepped(); } catch (e) { }
        };
    }

    /**
     * 左に∧∨を付けた入力欄を追加する
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - ∧∨の設定（min / max / integer / onStep）
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSteppedInput(parentRow, initialText, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = parentRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;
        var valueInput;
        var valueStepper = addStepper(stepperFieldGroup, function () { return valueInput; }, stepOptions);
        valueInput = stepperFieldGroup.add("edittext", undefined, initialText);
        valueInput.stepperGroup = valueStepper;
        bindSteppedArrowKeys(valueInput, valueStepper);
        return valueInput;
    }

    /**
     * 入力欄の有効／無効を∧∨ごと切り替える。状態が変わったときだけ∧∨を描き直す（同期のたびに呼ばれるため）
     * @param {EditText} valueInput - addSteppedInput() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedInputEnabled(valueInput, isEnabled) {
        if (valueInput.enabled === isEnabled && valueInput.stepperGroup.enabled === isEnabled) return;
        valueInput.enabled = isEnabled;
        valueInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(valueInput.stepperGroup);
    }

    /**
     * 値を min〜max の範囲に収める
     * @param {number} value - 対象の値
     * @param {number} min - 下限
     * @param {number} max - 上限
     * @returns {number} 範囲に収めた値
     */
    function clampNumber(value, min, max) {
        if (value < min) return min;
        if (value > max) return max;
        return value;
    }

    /**
     * 文字列を数値へ変換する（失敗したら既定値）
     * @param {string} text - 対象の文字列
     * @param {number} fallback - 既定値
     * @returns {number} 数値
     */
    function parseNumberOr(text, fallback) {
        var parsed = Number(text);
        if (isNaN(parsed)) return fallback;
        return parsed;
    }

    /**
     * 文字列を t 値（0〜5）として解釈する
     * @param {string} text - 対象の文字列
     * @param {number} fallback - 既定値
     * @returns {number} t 値
     */
    function parseTValue(text, fallback) {
        var parsed = Number(text);
        if (isNaN(parsed)) return fallback;
        return clampNumber(parsed, 0, 5);
    }

    /**
     * 数値を小数1桁の文字列にする（0.1pt 単位）
     * @param {number} value - 対象の値
     * @returns {string} 小数1桁の文字列
     */
    function formatOneDecimal(value) {
        return (Math.round(value * 10) / 10).toFixed(1);
    }

    /**
     * スミ100%の CMYKColor を作る
     * @returns {CMYKColor} 黒
     */
    function makeBlackCMYK() {
        var black = new CMYKColor();
        black.cyan = 0;
        black.magenta = 0;
        black.yellow = 0;
        black.black = 100;
        return black;
    }

    // =========================================
    // 選択の解析 / Selection analysis
    // =========================================

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

    /* 対象のテキスト：ポイント文字・パス上文字をグループの中まで（文字の選択は対象外）
       Target text: point and path text, including inside groups (text selections are not used) */
    var TARGET_TEXT_OPTIONS = { kinds: ["point", "path"], textRangeToFrame: false, unique: false };

    /* 一緒に選択したパス：パスと複合パス（複合パスはまとめて1つ）をグループの中まで
       Paths selected alongside: paths and compound paths (as a whole), including inside groups */
    var TARGET_PATH_OPTIONS = { compoundPaths: "whole", textRangeToFrame: false, unique: false };

    /**
     * 選択から、使える処理（パス上文字にする／分離）を判定する
     * @param {TextFrame[]} textItems - 対象のテキスト
     * @param {PageItem[]} pathItems - 一緒に選択されたパス
     * @returns {{hasPathTextTarget: boolean, canToPathText: boolean, canSplit: boolean}} 判定結果
     */
    function detectModeAvailability(textItems, pathItems) {
        var hasSelectedPath = (pathItems && pathItems.length > 0);
        var hasPathTextTarget = false;
        for (var i = 0; i < textItems.length; i++) {
            try {
                var textItem = textItems[i];
                if (textItem && textItem.typename === "TextFrame" && textItem.kind === TextType.PATHTEXT && textItem.textPath) {
                    hasPathTextTarget = true;
                    break;
                }
            } catch (e) { }
        }
        /* 「パス上文字にする」はパスを選んでいるか、パス上文字が自分のパスを持っていれば使える
           「分離」はパス上文字を選んでいるときだけ */
        return {
            hasPathTextTarget: hasPathTextTarget,
            canToPathText: hasSelectedPath || hasPathTextTarget,
            canSplit: hasPathTextTarget
        };
    }

    // =========================================
    // 前提チェック / Preconditions
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return false;
    }
    var doc = app.activeDocument;
    var initialSelection = doc.selection;

    /* ダイアログ表示中のプレビューを安定させるための、最初の選択の控え / Base selection snapshot for a stable preview */
    var baseSelection = [];
    /* テキスト編集中の selection は配列ではない（slice が無い） / the selection is not an array while editing text */
    try { baseSelection = initialSelection.slice(0); } catch (e) { baseSelection = []; }

    var targetItems = collectSelectionTextFrames(initialSelection, TARGET_TEXT_OPTIONS);
    var selectedPaths = collectSelectionPathItems(initialSelection, TARGET_PATH_OPTIONS);

    if (targetItems.length === 0) {
        alert(getLabel("alert.noText"));
        return false;
    }

    /**
     * 現在の選択を返す（プレビュー中に選択が空になったら最初の選択の控えを使う）
     * @returns {Object[]} 選択
     */
    function readCurrentSelection() {
        var currentSelection = [];
        try { currentSelection = doc.selection; } catch (e) { currentSelection = []; }
        if (!currentSelection || currentSelection.length === 0) currentSelection = baseSelection;
        return currentSelection;
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 横並び（左寄せ・上下中央）の行を追加する
     * @param {Object} parentContainer - 追加先
     * @returns {Group} 作成した行
     */
    function addRow(parentContainer) {
        var newRow = parentContainer.add("group");
        newRow.orientation = "row";
        newRow.alignChildren = ["left", "center"];
        return newRow;
    }

    /**
     * 上寄せ・幅いっぱいのグループを追加する
     * @param {Object} parentContainer - 追加先
     * @param {string} orientation - "row" または "column"
     * @returns {Group} 作成したグループ
     */
    function addFillGroup(parentContainer, orientation) {
        var newGroup = parentContainer.add("group");
        newGroup.orientation = orientation;
        newGroup.alignChildren = ["fill", "top"];
        newGroup.alignment = ["fill", "top"];
        return newGroup;
    }

    /**
     * tooltip 付きのラジオボタンを追加する
     * @param {Object} parentContainer - 追加先
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - tooltip のパス
     * @returns {RadioButton} 作成したラジオボタン
     */
    function addRadio(parentContainer, labelPath, tooltipPath) {
        var radio = parentContainer.add("radiobutton", undefined, getLabel(labelPath));
        radio.helpTip = getLabel(tooltipPath);
        return radio;
    }

    /**
     * 「項目名＋入力欄＋スライダー」の行を追加する（テキスト調整用）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 項目名のパス
     * @param {string} tooltipPath - 入力欄とスライダーの tooltip のパス
     * @param {string} initialText - 入力欄の初期値
     * @param {number} sliderMin - スライダーの最小値
     * @param {number} sliderMax - スライダーの最大値
     * @param {Object} stepOptions - ∧∨の設定（min / max / integer / onStep）
     * @returns {{valueInput: EditText, slider: Slider}} 作成した入力欄とスライダー
     */
    function addAdjustRow(parentPanel, labelPath, tooltipPath, initialText, sliderMin, sliderMax, stepOptions) {
        var adjustRow = addRow(parentPanel);
        var adjustLabel = adjustRow.add("statictext", undefined, labelText(labelPath));
        adjustLabel.preferredSize.width = ADJUST_LABEL_WIDTH;

        var valueInput = addSteppedInput(adjustRow, initialText, stepOptions);
        valueInput.helpTip = getLabel(tooltipPath);
        valueInput.characters = ADJUST_FIELD_CHARS;
        var slider = adjustRow.add("slider", undefined, 0, sliderMin, sliderMax);
        slider.helpTip = getLabel(tooltipPath);
        slider.preferredSize.width = ROW_SLIDER_WIDTH;
        return { valueInput: valueInput, slider: slider };
    }

    /**
     * 「チェック＋項目名＋入力欄＋スライダー」の行を追加する（開始／終了位置用）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 項目名のパス
     * @param {string} enabledTooltipPath - チェックボックスの tooltip のパス
     * @param {string} tooltipPath - 入力欄とスライダーの tooltip のパス
     * @param {string} initialText - 入力欄の初期値
     * @param {number} sliderValue - スライダーの初期値
     * @param {number} sliderMax - スライダーの最大値
     * @param {Object} stepOptions - ∧∨の設定（min / onStep）
     * @returns {{enabledCheckbox: Checkbox, valueInput: EditText, slider: Slider}} 作成したコントロール
     */
    function addPositionRow(parentPanel, labelPath, enabledTooltipPath, tooltipPath, initialText, sliderValue, sliderMax, stepOptions) {
        var positionRow = addRow(parentPanel);
        var enabledCheckbox = positionRow.add("checkbox", undefined, "");
        enabledCheckbox.helpTip = getLabel(enabledTooltipPath);
        enabledCheckbox.value = false;
        var positionLabel = positionRow.add("statictext", undefined, getLabel(labelPath));
        positionLabel.preferredSize.width = POSITION_LABEL_WIDTH;
        var valueInput = addSteppedInput(positionRow, initialText, stepOptions);
        valueInput.characters = POSITION_FIELD_CHARS;
        valueInput.helpTip = getLabel(tooltipPath);
        var slider = positionRow.add("slider", undefined, sliderValue, 0, sliderMax);
        slider.preferredSize.width = ROW_SLIDER_WIDTH;
        slider.helpTip = getLabel(tooltipPath);
        return { enabledCheckbox: enabledCheckbox, valueInput: valueInput, slider: slider };
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

    var pathTextDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(pathTextDialog);

    /* 2カラムレイアウト / Two-column layout */
    var columnsRow = pathTextDialog.add("group");
    columnsRow.orientation = "row";
    columnsRow.alignChildren = ["fill", "top"];
    columnsRow.alignment = ["fill", "top"];

    var leftCol = addFillGroup(columnsRow, "column");
    var rightCol = addFillGroup(columnsRow, "column");

    /* 処理パネル（パス上文字にする / アーチ生成 / 正円生成） / Process panel */
    var pnlProcess = leftCol.add("panel", undefined, getLabel("panel.process"));
    setupPanel(pnlProcess, 6);

    var rbToPathText = addRadio(pnlProcess, "radio.toPathText", "tooltip.toPathText");
    var rbGenArcPath = addRadio(pnlProcess, "radio.genArcPath", "tooltip.genArcPath");
    var rbGenCircle = addRadio(pnlProcess, "radio.genCircle", "tooltip.genCircle");
    rbToPathText.value = true;

    var btnTextEdit = pnlProcess.add("button", undefined, getLabel("button.textEdit"));
    btnTextEdit.helpTip = getLabel("tooltip.textEdit");
    btnTextEdit.alignment = ["left", "center"];

    /* 分離パネル（処理パネルの下） / Split panel (under Process) */
    var pnlSplit = leftCol.add("panel", undefined, getLabel("panel.split"));
    setupPanel(pnlSplit, 6);

    var rbSplitTextAndPath = addRadio(pnlSplit, "radio.splitKeepFormat", "tooltip.splitKeepFormat");
    var rbSplitTextAndPathNoFormat = addRadio(pnlSplit, "radio.splitNoFormat", "tooltip.splitNoFormat");
    rbSplitTextAndPath.value = false;
    rbSplitTextAndPathNoFormat.value = false;

    var cbSplitDeletePath = pnlSplit.add("checkbox", undefined, getLabel("checkbox.splitDeletePath"));
    cbSplitDeletePath.helpTip = getLabel("tooltip.splitDeletePath");
    cbSplitDeletePath.value = false; /* 既定はオフ / off by default */

    /* 選択に応じて使える処理を切り替える（アーチの既定値は UI をすべて作ってから入れる）
       Enable modes per selection (arc presets are applied once the whole UI exists) */
    var initialAvailability = detectModeAvailability(targetItems, selectedPaths);
    applyModeAvailability(initialAvailability);
    switchToArcIfUnavailable(initialAvailability);

    /* オプションパネル（右カラム） / Option panel (right column) */
    var pnlOption = rightCol.add("panel", undefined, getLabel("panel.option"));
    setupPanel(pnlOption, 6);

    var cbReverse = pnlOption.add("checkbox", undefined, getLabel("checkbox.reverse"));
    cbReverse.helpTip = getLabel("tooltip.reverse");
    cbReverse.value = false;

    /* アーチ方向 / Arc direction */
    var arcDirectionRow = addRow(pnlOption);
    var arcDirectionLabel = arcDirectionRow.add("statictext", undefined, labelText("fieldLabel.arcDirection"));
    arcDirectionLabel.preferredSize.width = ARC_LABEL_WIDTH;

    var rbArcUp = addRadio(arcDirectionRow, "radio.arcUp", "tooltip.arcUp");
    var rbArcDown = addRadio(arcDirectionRow, "radio.arcDown", "tooltip.arcDown");
    rbArcUp.value = true;

    /* まるみ（0=平ら、50=既定、100=最も丸い） / Arc roundness (0 = flat, 50 = default, 100 = roundest) */
    var arcRoundnessRow = addRow(pnlOption);
    arcRoundnessRow.add("statictext", undefined, getLabel("fieldLabel.arcRoundness"));
    var slArcRoundness = arcRoundnessRow.add("slider", undefined, 50, 0, 100);
    slArcRoundness.helpTip = getLabel("tooltip.arcRoundness");
    slArcRoundness.preferredSize.width = ARC_SLIDER_WIDTH;

    /* パス幅パネルの上の余白 / Spacer above the Fit-to-path-width panel */
    var fitWidthSpacer = pnlOption.add("group");
    fitWidthSpacer.preferredSize.height = FIT_PANEL_GAP;

    /* パス幅に合わせるパネル（オプションパネル内の最下部） / Fit-to-path-width panel (bottom of the Options panel) */
    var pnlFitWidth = pnlOption.add("panel", undefined, getLabel("panel.fitWidth"));
    setupPanel(pnlFitWidth, 6);

    var rbFitWidthNone = addRadio(pnlFitWidth, "radio.fitWidthNone", "tooltip.fitWidthNone");
    var rbFitWidthFontSize = addRadio(pnlFitWidth, "radio.fitWidthFontSize", "tooltip.fitWidthFontSize");
    var rbFitWidthTracking = addRadio(pnlFitWidth, "radio.fitWidthTracking", "tooltip.fitWidthTracking");
    rbFitWidthNone.value = true;

    /* 全幅（2カラムの下）：効果 → 位置 → テキスト調整 / Full width: effect -> position -> text adjust */
    var fullWidthColumn = addFillGroup(pathTextDialog, "column");

    /* 効果パネル：5項目を横並びにする / Effect panel (5 radios in a row) */
    var pnlEffect = fullWidthColumn.add("panel", undefined, getLabel("panel.effect"));
    setupPanel(pnlEffect, 6);
    pnlEffect.orientation = "row";
    pnlEffect.alignChildren = ["center", "center"];

    var rbEffectRainbow = addRadio(pnlEffect, "radio.effectRainbow", "tooltip.effectRainbow");
    var rbEffectDistort = addRadio(pnlEffect, "radio.effectDistort", "tooltip.effectDistort");
    var rbEffectRibbon = addRadio(pnlEffect, "radio.effectRibbon", "tooltip.effectRibbon");
    var rbEffectStep = addRadio(pnlEffect, "radio.effectStep", "tooltip.effectStep");
    var rbEffectGravity = addRadio(pnlEffect, "radio.effectGravity", "tooltip.effectGravity");
    var effectRadios = [rbEffectRainbow, rbEffectDistort, rbEffectRibbon, rbEffectStep, rbEffectGravity];

    /* 既定は効果なし。パス上文字を選んでいるときだけ「虹」 / No effect by default; Rainbow when path text is selected */
    for (var effectIndex = 0; effectIndex < effectRadios.length; effectIndex++) effectRadios[effectIndex].value = false;
    if (initialAvailability.hasPathTextTarget) rbEffectRainbow.value = true;

    /* 位置（開始／終了）パネル / Position (start/end) panel */
    var pnlPosition = fullWidthColumn.add("panel", undefined, getLabel("panel.position"));
    setupPanel(pnlPosition, 6);

    /* スライダーは t 値×100（開始 0.0〜4.0、終了 0.0〜5.0） / sliders hold t × 100 */
    /* ∧∨・↑↓キーで変えたらスライダーを追従させてプレビューを更新 / sync the slider and refresh the preview on each step */
    var onPositionStepped = makeSafeStepHandler(function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); });
    var startPosRow = addPositionRow(pnlPosition, "fieldLabel.startPos", "tooltip.startPosEnabled", "tooltip.startPos", "0.0", 0, 400,
        { min: 0, onStep: onPositionStepped });
    var cbStartT = startPosRow.enabledCheckbox;
    var etStartT = startPosRow.valueInput;
    var slStartT = startPosRow.slider;

    var endPosRow = addPositionRow(pnlPosition, "fieldLabel.endPos", "tooltip.endPosEnabled", "tooltip.endPos", "1.0", 100, 500,
        { min: 0, onStep: onPositionStepped });
    var cbEndT = endPosRow.enabledCheckbox;
    var etEndT = endPosRow.valueInput;
    var slEndT = endPosRow.slider;

    /* テキスト調整パネル / Text adjust panel */
    var pnlTextAdjust = fullWidthColumn.add("panel", undefined, getLabel("panel.textAdjust"));
    setupPanel(pnlTextAdjust, 6);

    /* 行揃え（パネル先頭の行） / Alignment (top row) */
    var alignRow = addRow(pnlTextAdjust);
    var alignLabel = alignRow.add("statictext", undefined, labelText("fieldLabel.align"));
    alignLabel.preferredSize.width = ADJUST_LABEL_WIDTH;

    var rbAlignLeft = addRadio(alignRow, "radio.alignLeft", "tooltip.align");
    var rbAlignCenter = addRadio(alignRow, "radio.alignCenter", "tooltip.align");
    var rbAlignRight = addRadio(alignRow, "radio.alignRight", "tooltip.align");
    var rbAlignFullJustify = addRadio(alignRow, "radio.alignFullJustify", "tooltip.align");
    rbAlignCenter.value = true;

    /* ベースライン・文字サイズのスライダーは 0.1pt 単位で、範囲は ±文字サイズに合わせて更新する
       Baseline and font-size sliders use 0.1pt steps; the range follows ±font size */
    /* ∧∨・↑↓キーで変えたらスライダーを追従させてプレビューを更新 / sync the slider and refresh the preview on each step */
    var baseShiftRow = addAdjustRow(pnlTextAdjust, "fieldLabel.baseShift", "tooltip.baseShift", "0.0", -1000, 1000,
        { onStep: makeSafeStepHandler(function () { baseShiftSync.fromEdit(); refreshPreviewIfNeeded(); }) });
    var etBaseShift = baseShiftRow.valueInput;
    var slBaseShift = baseShiftRow.slider;

    /* トラッキングは整数（-100〜500） / tracking is a whole number (-100 to 500) */
    var trackingRow = addAdjustRow(pnlTextAdjust, "fieldLabel.tracking", "tooltip.tracking", "0", -100, 500,
        { min: -100, max: 500, integer: true, onStep: makeSafeStepHandler(function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); }) });
    var etTracking = trackingRow.valueInput;
    var slTracking = trackingRow.slider;

    var fontSizeRow = addAdjustRow(pnlTextAdjust, "fieldLabel.fontSize", "tooltip.fontSize", "0.0", -1000, 1000,
        { onStep: makeSafeStepHandler(function () { fontSizeSync.fromEdit(); refreshPreviewIfNeeded(); }) });
    var etFontSize = fontSizeRow.valueInput;
    var slFontSize = fontSizeRow.slider;

    /* ボタンエリア（左＝プレビュー、右＝キャンセル／OK） / Buttons (preview on the left, Cancel/OK on the right) */
    var buttonRow = addButtonRow(pathTextDialog);

    var cbPreview = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
    cbPreview.helpTip = getLabel("tooltip.preview");
    cbPreview.value = true;

    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    // =========================================
    // ダイアログの状態 / Dialog state
    // =========================================

    /**
     * 「処理」ラジオをすべてオフにする（別パネルの「分離」と手動で排他にする）
     * @returns {void}
     */
    function clearProcessRadios() {
        rbToPathText.value = false;
        rbGenArcPath.value = false;
        rbGenCircle.value = false;
    }

    /**
     * 「分離」ラジオをすべてオフにする（別パネルの「処理」と手動で排他にする）
     * @returns {void}
     */
    function clearSplitRadios() {
        rbSplitTextAndPath.value = false;
        rbSplitTextAndPathNoFormat.value = false;
    }

    /**
     * 使える処理に合わせてラジオの有効／無効を切り替える
     * @param {{canToPathText: boolean, canSplit: boolean}} availability - detectModeAvailability() の結果
     * @returns {void}
     */
    function applyModeAvailability(availability) {
        rbToPathText.enabled = availability.canToPathText;
        rbSplitTextAndPath.enabled = availability.canSplit;
        rbSplitTextAndPathNoFormat.enabled = availability.canSplit;
        cbSplitDeletePath.enabled = availability.canSplit;
        rbGenCircle.enabled = true;
    }

    /**
     * 選んでいる処理が使えないときだけ「アーチ状のパスを生成」へ切り替える
     * @param {{canToPathText: boolean, canSplit: boolean}} availability - detectModeAvailability() の結果
     * @returns {boolean} 切り替えたら true
     */
    function switchToArcIfUnavailable(availability) {
        var switched = false;
        if (rbToPathText.value && !availability.canToPathText) {
            rbToPathText.value = false;
            rbGenArcPath.value = true;
            switched = true;
        }
        if ((rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value) && !availability.canSplit) {
            clearSplitRadios();
            rbGenArcPath.value = true;
            switched = true;
        }
        return switched;
    }

    /**
     * アーチモードの既定値（中央揃え・虹・文字サイズでフィット）を入れる
     * @returns {void}
     */
    function applyArcModePresets() {
        rbAlignCenter.value = true;
        rbEffectRainbow.value = true;
        rbFitWidthFontSize.value = true;
    }

    /**
     * 処理に応じてパネルの有効／無効を切り替える
     * （分離では処理・分離以外を無効に。正円ではオープンパス用のパス幅・アーチ方向・まるみも無効に）
     * @returns {void}
     */
    function updatePanelsByMode() {
        var isSplitMode = rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value;
        var isCircleMode = rbGenCircle.value;
        var isAdjustable = !isSplitMode;

        pnlProcess.enabled = true;
        pnlSplit.enabled = true;
        pnlOption.enabled = isAdjustable;
        pnlEffect.enabled = isAdjustable;
        pnlTextAdjust.enabled = isAdjustable;
        pnlPosition.enabled = isAdjustable;
        pnlFitWidth.enabled = isAdjustable && !isCircleMode;
        arcDirectionRow.enabled = isAdjustable && !isCircleMode;
        arcRoundnessRow.enabled = isAdjustable && !isCircleMode;
        /* 自作描画の∧∨はパネルの無効化に追従しないので描き直す / redraw the custom-drawn steppers */
        redrawSteppersIn(pnlTextAdjust);
        redrawSteppersIn(pnlPosition);
    }

    /**
     * 現在の選択から対象テキスト・パスと、処理ラジオの有効状態を取り直す
     * @returns {void}
     */
    function refreshTargetsFromSelection() {
        var currentSelection = readCurrentSelection();
        targetItems = collectSelectionTextFrames(currentSelection, TARGET_TEXT_OPTIONS);
        selectedPaths = collectSelectionPathItems(currentSelection, TARGET_PATH_OPTIONS);

        var availability = detectModeAvailability(targetItems, selectedPaths);
        applyModeAvailability(availability);
        if (switchToArcIfUnavailable(availability)) applyArcModePresets();
        updatePanelsByMode();
    }

    // =========================================
    // 数値欄とスライダーの同期 / Field-slider sync
    // =========================================

    /**
     * 基準の文字サイズ（先頭のテキスト）を pt で返す
     * @returns {number} 文字サイズ（取れなければ 100）
     */
    function getReferenceFontSizePt() {
        try {
            var firstTarget = (targetItems && targetItems.length > 0) ? targetItems[0] : null;
            if (firstTarget && firstTarget.typename === "TextFrame") {
                /* 先頭の textRange を優先し、無ければ全体 / Prefer the first textRange, fall back to the whole range */
                if (firstTarget.textRanges && firstTarget.textRanges.length > 0) {
                    var firstRangeSize = firstTarget.textRanges[0].characterAttributes.size;
                    if (firstRangeSize && !isNaN(firstRangeSize)) return firstRangeSize;
                }
                var fallbackSize = firstTarget.textRange.characterAttributes.size;
                if (fallbackSize && !isNaN(fallbackSize)) return fallbackSize;
            }
        } catch (e) { }
        return 100;
    }

    /**
     * 増減スライダーの範囲を ±文字サイズ（0.1pt 単位）にする
     * @param {Slider} deltaSlider - 対象のスライダー
     * @returns {void}
     */
    function updateDeltaSliderRange(deltaSlider) {
        var sliderLimit = Math.max(1, Math.round(getReferenceFontSizePt() * 10));
        deltaSlider.minvalue = -sliderLimit;
        deltaSlider.maxvalue = sliderLimit;
    }

    /**
     * 0.1pt 単位の増減欄とスライダーを相互に同期する関数の組を作る（ベースライン・文字サイズ用）
     * @param {EditText} deltaInput - 入力欄
     * @param {Slider} deltaSlider - スライダー（値＝増減×10）
     * @returns {{fromEdit: Function, fromSlider: Function}} 入力欄→スライダー、スライダー→入力欄の同期
     */
    function createDeltaSliderSync(deltaInput, deltaSlider) {
        var isSyncing = false;
        return {
            fromEdit: function () {
                if (isSyncing) return;
                isSyncing = true;
                updateDeltaSliderRange(deltaSlider);
                /* スライダーの範囲に収める / clamp to the slider range */
                var deltaValue = clampNumber(parseNumberOr(deltaInput.text, 0), deltaSlider.minvalue / 10.0, deltaSlider.maxvalue / 10.0);
                deltaInput.text = formatOneDecimal(deltaValue);
                deltaSlider.value = Math.round(deltaValue * 10);
                isSyncing = false;
            },
            fromSlider: function () {
                if (isSyncing) return;
                isSyncing = true;
                updateDeltaSliderRange(deltaSlider);
                deltaInput.text = formatOneDecimal(deltaSlider.value / 10.0);
                isSyncing = false;
            }
        };
    }

    var baseShiftSync = createDeltaSliderSync(etBaseShift, slBaseShift);
    var fontSizeSync = createDeltaSliderSync(etFontSize, slFontSize);

    var isTrackingSyncing = false;

    /**
     * 入力欄からトラッキングのスライダーへ同期する（-100〜500 の整数）
     * @returns {void}
     */
    function syncTrackingFromEdit() {
        if (isTrackingSyncing) return;
        isTrackingSyncing = true;
        var trackingValue = Math.round(parseNumberOr(etTracking.text, 0));
        if (trackingValue > 500) trackingValue = 500;
        if (trackingValue < -100) trackingValue = -100;
        etTracking.text = String(trackingValue);
        slTracking.value = trackingValue;
        isTrackingSyncing = false;
    }

    /**
     * スライダーからトラッキングの入力欄へ同期する
     * @returns {void}
     */
    function syncTrackingFromSlider() {
        if (isTrackingSyncing) return;
        isTrackingSyncing = true;
        etTracking.text = String(Math.round(slTracking.value));
        isTrackingSyncing = false;
    }

    var isPositionSyncing = false;

    /**
     * チェックボックスに応じて開始／終了位置の入力を有効にする
     * @returns {void}
     */
    function updatePositionInputsEnabled() {
        setSteppedInputEnabled(etStartT, cbStartT.value);
        slStartT.enabled = cbStartT.value;
        setSteppedInputEnabled(etEndT, cbEndT.value);
        slEndT.enabled = cbEndT.value;
    }

    /**
     * 開始／終了位置の値を表示欄の範囲に収めて書き戻す（オフの側は既定値 0.0 / 1.0）
     * @param {number} startValue - 開始位置
     * @param {number} endValue - 終了位置
     * @returns {{startValue: number, endValue: number}} 収めた値
     */
    function writePositionFields(startValue, endValue) {
        startValue = clampNumber(startValue, 0.0, 4.0);
        endValue = clampNumber(endValue, 0.0, 5.0);
        if (!cbStartT.value) startValue = 0.0;
        if (!cbEndT.value) endValue = 1.0;
        etStartT.text = formatOneDecimal(startValue);
        etEndT.text = formatOneDecimal(endValue);
        return { startValue: startValue, endValue: endValue };
    }

    /**
     * 入力欄から開始／終了位置のスライダーへ同期する
     * @returns {void}
     */
    function syncPositionFromEdits() {
        if (isPositionSyncing) return;
        isPositionSyncing = true;
        updatePositionInputsEnabled();
        var startValue = cbStartT.value ? parseTValue(etStartT.text, 0.0) : 0.0;
        var endValue = cbEndT.value ? parseTValue(etEndT.text, 1.0) : 1.0;
        var clampedValues = writePositionFields(startValue, endValue);
        slStartT.value = Math.round(clampedValues.startValue * 100);
        slEndT.value = Math.round(clampedValues.endValue * 100);
        isPositionSyncing = false;
    }

    /**
     * スライダーから開始／終了位置の入力欄へ同期する
     * @returns {void}
     */
    function syncPositionFromSliders() {
        if (isPositionSyncing) return;
        isPositionSyncing = true;
        updatePositionInputsEnabled();
        var startValue = cbStartT.value ? (slStartT.value / 100.0) : 0.0;
        var endValue = cbEndT.value ? (slEndT.value / 100.0) : 1.0;
        writePositionFields(startValue, endValue);
        isPositionSyncing = false;
    }

    // =========================================
    // プレビュー（取り消しを使わない） / Preview (no undo)
    // =========================================
    var previewTempItems = [];        /* プレビューで作ったもの / items created during preview */
    var previewHiddenOriginals = [];  /* プレビュー中に隠した元のオブジェクト / originals hidden during preview */
    var previewPathStates = [];       /* { item, stroked, filled, strokeWidth, strokeColor, opacity } */

    /**
     * プレビューで作った一時オブジェクトを削除し、元の状態へ戻す
     * @returns {void}
     */
    function clearPreview() {
        for (var i = previewTempItems.length - 1; i >= 0; i--) {
            try { previewTempItems[i].remove(); } catch (e) { }
        }
        previewTempItems = [];

        for (var j = previewHiddenOriginals.length - 1; j >= 0; j--) {
            try { previewHiddenOriginals[j].hidden = false; } catch (e) { }
        }
        /* プレビュー中に変えたパスの見た目を戻す / Restore path appearance changed during preview */
        for (var k = previewPathStates.length - 1; k >= 0; k--) {
            try {
                var pathState = previewPathStates[k];
                pathState.item.stroked = pathState.stroked;
                pathState.item.filled = pathState.filled;
                pathState.item.strokeWidth = pathState.strokeWidth;
                pathState.item.strokeColor = pathState.strokeColor;
                pathState.item.opacity = pathState.opacity;
            } catch (e) { }
        }
        previewPathStates = [];
        previewHiddenOriginals = [];
    }

    /**
     * プレビューの復元用に、パスの元の見た目を控える
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function recordPreviewPathState(pathItem) {
        try {
            if (!pathItem) return;
            for (var i = 0; i < previewPathStates.length; i++) {
                if (previewPathStates[i].item === pathItem) return;
            }
            previewPathStates.push({
                item: pathItem,
                stroked: pathItem.stroked,
                filled: pathItem.filled,
                strokeWidth: pathItem.strokeWidth,
                strokeColor: pathItem.strokeColor,
                opacity: pathItem.opacity
            });
        } catch (e) { }
    }

    /**
     * 複合パスも展開して、各 PathItem に処理を適用する
     * @param {PageItem} pathItem - PathItem または CompoundPathItem
     * @param {Function} styleFunction - 各 PathItem を受け取る関数
     * @returns {void}
     */
    function forEachPathItem(pathItem, styleFunction) {
        try {
            if (!pathItem || !styleFunction) return;
            if (pathItem.typename === "CompoundPathItem") {
                for (var i = 0; i < pathItem.pathItems.length; i++) {
                    try { styleFunction(pathItem.pathItems[i]); } catch (e) { }
                }
            } else {
                try { styleFunction(pathItem); } catch (e) { }
            }
        } catch (e) { }
    }

    /**
     * パスを見えるガイド（塗りなし・黒1pt・不透明度50%）にする
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function styleVisiblePath(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeWidth = 1;
        pathItem.strokeColor = makeBlackCMYK();
        pathItem.opacity = 50;
    }

    /**
     * パスを見えない状態（塗り・線なし）にする
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function styleInvisiblePath(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = false;
        pathItem.strokeWidth = 0;
    }

    /**
     * プレビュー：元の見た目を控えてから、見えるガイドにする
     * @param {PageItem} basePath - 対象のパス
     * @returns {void}
     */
    function applyPreviewPathStyle(basePath) {
        forEachPathItem(basePath, function (pathItem) {
            recordPreviewPathState(pathItem);
            styleVisiblePath(pathItem);
        });
    }

    /**
     * 実行：パスを見えるガイドのまま残す（戻さない）
     * @param {PageItem} pathItem - 対象のパス
     * @returns {void}
     */
    function applyExecutePathStyle(pathItem) {
        forEachPathItem(pathItem, styleVisiblePath);
    }

    /**
     * 生成したパスを見えないガイド（線幅0）にする
     * @param {PageItem} pathItem - 対象のパス
     * @returns {void}
     */
    function applyInvisiblePathStyle(pathItem) {
        forEachPathItem(pathItem, styleInvisiblePath);
    }

    /**
     * 元のオブジェクトをプレビュー中だけ隠す
     * @param {PageItem} originalItem - 対象のオブジェクト
     * @returns {void}
     */
    function hideOriginalForPreview(originalItem) {
        /* ロック中のオブジェクトは hidden の代入が例外になりうる / hiding a locked item can throw */
        try {
            if (!originalItem) return;
            for (var i = 0; i < previewHiddenOriginals.length; i++) {
                if (previewHiddenOriginals[i] === originalItem) return;
            }
            originalItem.hidden = true;
            previewHiddenOriginals.push(originalItem);
        } catch (e) { }
    }

    /**
     * 選んでいる処理を実行する
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューとして実行するか
     * @returns {void}
     */
    function runSelectedMode(showAlerts, previewMode) {
        if (rbSplitTextAndPath.value) {
            splitPathTextAndPath(showAlerts, previewMode, true);
        } else if (rbSplitTextAndPathNoFormat.value) {
            splitPathTextAndPath(showAlerts, previewMode, false);
        } else if (rbToPathText.value) {
            createPathTextOnSelectedPath(showAlerts, previewMode);
        } else if (rbGenArcPath.value) {
            createPathTextOnArc(showAlerts, previewMode);
        } else if (rbGenCircle.value) {
            createPathTextOnCircle(showAlerts, previewMode);
        }
    }

    /**
     * 現在の設定でプレビューを作り直す（取り消しは使わない）
     * @returns {void}
     */
    function rebuildPreview() {
        clearPreview();
        /* 選択が変わってもプレビューが安定するよう、最初の選択に戻す / Restore the base selection for a stable preview */
        try { doc.selection = baseSelection; } catch (e) { }
        refreshTargetsFromSelection();

        if (!targetItems || targetItems.length === 0) {
            cbPreview.value = false;
            return;
        }

        runSelectedMode(false, true);
        app.redraw();
    }

    /**
     * プレビューがオンのときだけ作り直す
     * @returns {void}
     */
    function refreshPreviewIfNeeded() {
        if (cbPreview.value) {
            rebuildPreview();
        }
    }

    // =========================================
    // テキスト編集ダイアログ / Text edit dialog
    // =========================================

    /**
     * 選択中のテキストの内容を書き換える小さなダイアログを開く
     * @returns {void}
     */
    function showTextEditDialog() {
        var textEditDialog = new Window("dialog", getLabel("dialog.textEditTitle"));
        setupWindow(textEditDialog);

        var hintLabel = textEditDialog.add("statictext", undefined, getLabel("dialog.textEditHint"), { multiline: true });
        hintLabel.preferredSize.width = EDIT_HINT_WIDTH;

        /* 対象はクリック時の選択から取り直す（古い控えに頼らない） / Resolve targets at click time */
        var editTargets = [];
        /* テキスト編集中の選択は配列でなく、たどると例外になる / a text-editing selection throws when walked */
        try { editTargets = collectSelectionTextFrames(readCurrentSelection(), TARGET_TEXT_OPTIONS); } catch (e) { editTargets = []; }

        var initialText = "";
        try {
            if (editTargets && editTargets.length > 0 && editTargets[0] && editTargets[0].typename === "TextFrame") {
                initialText = String(editTargets[0].contents);
            }
        } catch (e) { initialText = ""; }

        var textEditInput = textEditDialog.add("edittext", undefined, initialText, { multiline: true });
        textEditInput.preferredSize = EDIT_FIELD_SIZE;

        var editButtonRow = addButtonRow(textEditDialog);
        var btnEditCancel = editButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
        var btnEditOK = editButtonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(editButtonRow);

        btnEditCancel.onClick = function () {
            textEditDialog.close(0);
        };
        btnEditOK.onClick = function () {
            var newText = textEditInput.text;
            for (var i = 0; i < editTargets.length; i++) {
                var textFrame = editTargets[i];
                if (!textFrame || textFrame.typename !== "TextFrame") continue;
                try { textFrame.contents = newText; } catch (e) { }
            }
            refreshTargetsFromSelection();
            refreshPreviewIfNeeded();
            app.redraw();
            textEditDialog.close(1);
        };
        prepareDialogWindow(textEditDialog, SCRIPT_NAME + "_textEdit");
        textEditDialog.show();
    }

    // =========================================
    // イベント / Event handlers
    // =========================================

    btnTextEdit.onClick = showTextEditDialog;

    cbPreview.onClick = function () {
        if (cbPreview.value) {
            rebuildPreview();
        } else {
            clearPreview();
            app.redraw();
        }
    };

    /**
     * 処理・分離ラジオを押したときの共通処理
     * @param {Function} clearOtherGroup - もう一方のグループのラジオをオフにする関数
     * @param {boolean} isArcMode - アーチモードなら true（既定値を入れる。ほかはパス幅フィットを「しない」に戻す）
     * @returns {void}
     */
    function onModeRadioClick(clearOtherGroup, isArcMode) {
        clearOtherGroup();
        if (!isArcMode) rbFitWidthNone.value = true;
        updatePanelsByMode();
        if (isArcMode) applyArcModePresets();
        refreshPreviewIfNeeded();
    }

    rbToPathText.onClick = function () { onModeRadioClick(clearSplitRadios, false); };
    rbGenCircle.onClick = function () { onModeRadioClick(clearSplitRadios, false); };
    rbGenArcPath.onClick = function () { onModeRadioClick(clearSplitRadios, true); };
    rbSplitTextAndPath.onClick = function () { onModeRadioClick(clearProcessRadios, false); };
    rbSplitTextAndPathNoFormat.onClick = function () { onModeRadioClick(clearProcessRadios, false); };

    /* 押したらプレビューを作り直すだけのコントロール / Controls that only refresh the preview */
    var previewRefreshControls = [
        cbReverse,
        rbEffectRainbow, rbEffectDistort, rbEffectRibbon, rbEffectStep, rbEffectGravity,
        rbAlignLeft, rbAlignCenter, rbAlignRight, rbAlignFullJustify,
        rbArcUp, rbArcDown,
        rbFitWidthNone, rbFitWidthFontSize, rbFitWidthTracking
    ];
    for (var controlIndex = 0; controlIndex < previewRefreshControls.length; controlIndex++) {
        previewRefreshControls[controlIndex].onClick = refreshPreviewIfNeeded;
    }

    /* まるみ：スライダーを離したら更新 / Roundness: refresh on release */
    slArcRoundness.onChange = refreshPreviewIfNeeded;

    etBaseShift.onChanging = function () { baseShiftSync.fromEdit(); refreshPreviewIfNeeded(); };
    slBaseShift.onChanging = function () { baseShiftSync.fromSlider(); };
    slBaseShift.onChange = function () { baseShiftSync.fromSlider(); refreshPreviewIfNeeded(); };
    baseShiftSync.fromEdit();

    etTracking.onChanging = function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); };
    slTracking.onChanging = function () { syncTrackingFromSlider(); };
    slTracking.onChange = function () { syncTrackingFromSlider(); refreshPreviewIfNeeded(); };
    syncTrackingFromEdit();

    etFontSize.onChanging = function () { fontSizeSync.fromEdit(); refreshPreviewIfNeeded(); };
    slFontSize.onChanging = function () { fontSizeSync.fromSlider(); };
    slFontSize.onChange = function () { fontSizeSync.fromSlider(); refreshPreviewIfNeeded(); };
    fontSizeSync.fromEdit();

    etStartT.onChanging = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };
    etEndT.onChanging = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };
    slStartT.onChanging = function () { syncPositionFromSliders(); };
    slStartT.onChange = function () { syncPositionFromSliders(); refreshPreviewIfNeeded(); };
    slEndT.onChanging = function () { syncPositionFromSliders(); };
    slEndT.onChange = function () { syncPositionFromSliders(); refreshPreviewIfNeeded(); };

    cbStartT.onClick = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };
    cbEndT.onClick = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };

    syncPositionFromEdits();

    btnCancel.onClick = function () {
        clearPreview();
        pathTextDialog.close(0);
    };

    btnOK.onClick = function () {
        /* プレビューが出ていれば先に消し、一時オブジェクトを重ねない / Clear the preview first so temp items do not stack */
        if (cbPreview.value) {
            clearPreview();
        }
        runSelectedMode(true, false);
        pathTextDialog.close(1);
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 起動時にアーチモードへ切り替わっていたら、UI がそろったここで既定値を入れる
       If arc mode was chosen at startup, apply its presets now that the whole UI exists */
    if (rbGenArcPath.value) applyArcModePresets();

    updatePanelsByMode();

    /* 開いた直後に一度プレビュー（取り消しは使わない） / Preview once on open (no undo) */
    /* 変換処理の DOM 例外でダイアログが開かなくならないように / keep the dialog opening even if the preview throws */
    try {
        if (cbPreview.value) {
            rebuildPreview();
            app.redraw();
        }
    } catch (e) { }

    alignRightOnlyButtonRow(buttonRow);
    prepareDialogWindow(pathTextDialog, SCRIPT_NAME);
    var dialogResult = pathTextDialog.show();
    if (dialogResult !== 1) return false;
    return true;

    // =========================================
    // 変換処理 / Conversion
    // =========================================

    /**
     * 選んでいる効果をメニューコマンドでパス上文字に適用する
     * @param {TextFrame} textFrame - 対象のパス上文字
     * @returns {void}
     */
    function applyPathTextEffect(textFrame) {
        var effectCommand = getSelectedEffectCommand();
        if (!effectCommand) return;

        var previousSelection = null;
        try { previousSelection = doc.selection; } catch (e) { previousSelection = null; }

        try {
            try { doc.selection = []; } catch (e) { }
            try { textFrame.selected = true; } catch (e) { }
            app.executeMenuCommand(effectCommand);
        } catch (e) { }

        /* 選択は必ず戻す / Always restore the selection */
        try { doc.selection = previousSelection; } catch (e) { }
    }

    /**
     * 選んでいる効果に対応するメニューコマンド名を返す
     * @returns {string|null} メニューコマンド名（効果なしなら null）
     */
    function getSelectedEffectCommand() {
        if (rbEffectRainbow.value) return "Rainbow";
        if (rbEffectDistort.value) return "Skew";
        if (rbEffectRibbon.value) return "3D ribbon";
        if (rbEffectStep.value) return "Stair Step";
        if (rbEffectGravity.value) return "Gravity";
        return null;
    }

    /**
     * ［行揃え］の選択を Justification に変換する（左揃えは「最終行左揃え」）
     * @returns {Justification} 行揃え
     */
    function readSelectedJustification() {
        if (rbAlignFullJustify.value) return Justification.FULLJUSTIFY;
        if (rbAlignRight.value) return Justification.RIGHT;
        if (rbAlignCenter.value) return Justification.CENTER;
        return Justification.FULLJUSTIFYLASTLINELEFT;
    }

    /**
     * テキストフレームの全段落と全体に行揃えを設定する（内容を複製したあとに呼ぶ）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Justification} justification - 行揃え
     * @returns {void}
     */
    function setFrameJustification(textFrame, justification) {
        if (!textFrame) return;
        try {
            if (textFrame.paragraphs && textFrame.paragraphs.length > 0) {
                for (var i = 0; i < textFrame.paragraphs.length; i++) {
                    try { textFrame.paragraphs[i].paragraphAttributes.justification = justification; } catch (e) { }
                }
            }
        } catch (e) { }
        /* 念のため全体にも / Also apply to the whole range */
        try { textFrame.textRange.paragraphAttributes.justification = justification; } catch (e) { }
    }

    /**
     * UI の開始／終了位置をパス上文字へ適用する
     * @param {TextFrame} textFrame - 対象のパス上文字
     * @returns {void}
     */
    function applyStartEndTValue(textFrame) {
        try {
            if (!textFrame) return;
            if (cbStartT.value) textFrame.startTValue = parseTValue(etStartT.text, 0.0);
            if (cbEndT.value) textFrame.endTValue = parseTValue(etEndT.text, 1.0);
        } catch (e) { }
    }

    /**
     * テキストフレームの textRange を配列で返す（無ければ textRange 単体）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {TextRange[]} textRange の配列
     */
    function collectTextRanges(textFrame) {
        var textRanges = [];
        try {
            if (textFrame.textRanges && textFrame.textRanges.length > 0) {
                for (var i = 0; i < textFrame.textRanges.length; i++) textRanges.push(textFrame.textRanges[i]);
            }
        } catch (e) { }
        if (textRanges.length === 0) {
            try { if (textFrame.textRange) textRanges = [textFrame.textRange]; } catch (e) { textRanges = []; }
        }
        return textRanges;
    }

    /**
     * 各 textRange の文字属性に増減を加える（clampMin があれば下限を適用）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} attributeName - 文字属性の名前
     * @param {number} delta - 増減
     * @param {number} [clampMin] - 下限
     * @returns {void}
     */
    function addDeltaToTextRanges(textFrame, attributeName, delta, clampMin) {
        if (!textFrame || !delta) return;
        var textRanges = collectTextRanges(textFrame);
        for (var i = 0; i < textRanges.length; i++) {
            try {
                var attributes = textRanges[i].characterAttributes;
                var nextValue = attributes[attributeName] + delta;
                if (typeof clampMin === "number" && nextValue < clampMin) nextValue = clampMin;
                attributes[attributeName] = nextValue;
            } catch (e) { }
        }
    }

    /**
     * 各 textRange の文字サイズに［文字サイズ］の増減を加える（最小 0.1pt）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function applyFontSizeDelta(textFrame) {
        addDeltaToTextRanges(textFrame, "size", parseNumberOr(etFontSize.text, 0), 0.1);
    }

    /**
     * 各 textRange のベースラインとトラッキングに［テキスト調整］の増減を加える（行揃えのあとに呼ぶ）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function applyShiftAndTrackingDelta(textFrame) {
        addDeltaToTextRanges(textFrame, "baselineShift", parseNumberOr(etBaseShift.text, 0));
        addDeltaToTextRanges(textFrame, "tracking", Math.round(parseNumberOr(etTracking.text, 0)));
    }

    /**
     * 指定のパス上にパス上文字を作り、元のテキストの直前へ置く
     * @param {PathItem} textPath - 使うパス
     * @param {TextFrame} originalText - 元のテキスト
     * @param {Layer} currentLayer - 作成先のレイヤー
     * @param {boolean} previewMode - プレビューなら true
     * @returns {TextFrame} 作成したパス上文字
     */
    function createPathTextFrame(textPath, originalText, currentLayer, previewMode) {
        var textOnAPath = currentLayer.textFrames.pathText(textPath);
        /* 重ね順を保つ（ほかのオブジェクトの背面に回り込まないように） / Keep stacking order */
        try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (e) { }
        if (previewMode) previewTempItems.push(textOnAPath);
        return textOnAPath;
    }

    /**
     * パス上文字の共通の仕上げ：内側配置・効果・内容の複製・各種調整を適用し、元のテキストを隠す／削除する
     * @param {TextFrame} textOnAPath - 作成したパス上文字
     * @param {TextFrame} originalText - 元のテキスト
     * @param {boolean} previewMode - プレビューなら true
     * @param {{forceCenter: boolean, circleInsideShift: boolean}} [decorateOptions] - 正円用の指定
     * @returns {void}
     */
    function decoratePathText(textOnAPath, originalText, previewMode, decorateOptions) {
        decorateOptions = decorateOptions || {};

        /* 内側配置 / Inside placement */
        if (cbReverse.value) {
            try { textOnAPath.textPath.polarity = PolarityValues.NEGATIVE; } catch (e) { }
        }

        applyPathTextEffect(textOnAPath);
        applyStartEndTValue(textOnAPath);

        /* 正円＋内側：開始／終了位置を +0.5 ずらして見かけの開始位置を合わせる（上限は 5、開始は 1 を超えてよい）
           Circle + inside: shift start/end by +0.5 (clamped to the global cap of 5) */
        if (decorateOptions.circleInsideShift) {
            try {
                var shiftAmount = 0.5;
                textOnAPath.startTValue = clampNumber(textOnAPath.startTValue + shiftAmount, 0.0, 5.0);
                textOnAPath.endTValue = clampNumber(textOnAPath.endTValue + shiftAmount, 0.0, 5.0);
            } catch (e) { }
        }

        /* テキストの内容を元から複製 / Copy the text content from the original */
        for (var i = 0; i < originalText.textRanges.length; i++) {
            originalText.textRanges[i].duplicate(textOnAPath);
        }

        applyFontSizeDelta(textOnAPath);

        /* 行揃え（複製のあとでないと上書きされる） / Justification (apply after the content is copied) */
        setFrameJustification(textOnAPath, decorateOptions.forceCenter ? Justification.CENTER : readSelectedJustification());

        applyShiftAndTrackingDelta(textOnAPath);

        if (previewMode) {
            hideOriginalForPreview(originalText);
        } else {
            try { originalText.remove(); } catch (e) { }
            /* 実行時は作ったパス上文字を選択 / Select the created path text on execute */
            try { textOnAPath.selected = true; } catch (e) { }
        }
    }

    /**
     * 各テキストに使う元のパスを決める（選択したパスを優先し、無ければ自分の textPath）
     * @param {PageItem[]} pathItems - 選択したパス
     * @param {number} textCount - テキストの数
     * @param {number} textIndex - このテキストの位置
     * @param {TextFrame} originalText - 元のテキスト
     * @returns {PageItem|null} 使うパス
     */
    function resolveBasePathForText(pathItems, textCount, textIndex, originalText) {
        if (pathItems && pathItems.length > 0) {
            if (pathItems.length === textCount) return pathItems[textIndex];
            return pathItems[0];
        }
        try {
            if (originalText && originalText.typename === "TextFrame") {
                if (originalText.kind === TextType.PATHTEXT && originalText.textPath) {
                    return originalText.textPath;
                }
            }
        } catch (e) { }
        return null;
    }

    /**
     * 選択した元のパスを見えなくする（塗り・線なし）
     * @param {PageItem} basePath - 対象のパス
     * @returns {void}
     */
    function makeOriginalPathInvisible(basePath) {
        try {
            if (basePath.typename === "CompoundPathItem") {
                for (var i = 0; i < basePath.pathItems.length; i++) {
                    basePath.pathItems[i].stroked = false;
                    basePath.pathItems[i].filled = false;
                }
            } else {
                basePath.stroked = false;
                basePath.filled = false;
            }
        } catch (e) { }
    }

    /**
     * 複製に使う PathItem を返す（複合パスは先頭）
     * @param {PageItem} basePath - 対象のパス
     * @returns {PathItem|null} 複製に使うパス
     */
    function resolveSourcePath(basePath) {
        try {
            return (basePath.typename === "CompoundPathItem") ? basePath.pathItems[0] : basePath;
        } catch (e) {
            return null;
        }
    }

    /**
     * ハンドルとポイントの種類を保ったままパスを複製する
     * @param {PathItem} originalPath - 元のパス
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @returns {PathItem|null} 複製したパス
     */
    function duplicatePathWithHandles(originalPath, targetLayer) {
        try {
            if (!originalPath) return null;
            var pathContainer = (targetLayer && targetLayer.pathItems) ? targetLayer.pathItems : doc.pathItems;
            var newPath = pathContainer.add();

            for (var i = 0; i < originalPath.pathPoints.length; i++) {
                var originalPoint = originalPath.pathPoints[i];
                var newPoint = newPath.pathPoints.add();
                newPoint.anchor = originalPoint.anchor;
                newPoint.leftDirection = originalPoint.leftDirection;
                newPoint.rightDirection = originalPoint.rightDirection;
                newPoint.pointType = originalPoint.pointType;
            }
            newPath.closed = originalPath.closed;
            return newPath;
        } catch (e) {
            return null;
        }
    }

    /**
     * テキスト用にパスを複製し、同じレイヤーへ置く
     * @param {PageItem} basePath - 元のパス
     * @param {Layer} currentLayer - 置き先のレイヤー
     * @returns {PathItem|null} 複製したパス
     */
    function duplicatePathForText(basePath, currentLayer) {
        var sourcePath = resolveSourcePath(basePath);
        if (!sourcePath) return null;

        /* 1) まず標準の複製 / Try the native duplicate first */
        try {
            var duplicatedPath = sourcePath.duplicate();
            try { duplicatedPath.move(currentLayer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
            return duplicatedPath;
        } catch (e) {
            /* 2) だめならアンカーとハンドルを写して作る（失敗すると null） / Fall back to copying anchors and handles */
            return duplicatePathWithHandles(sourcePath, currentLayer);
        }
    }

    /**
     * 処理：選択したパスに沿ってパス上文字を作る
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function createPathTextOnSelectedPath(showAlerts, previewMode) {
        var createdTexts = [];

        for (var i = 0; i < targetItems.length; i++) {
            var originalText = targetItems[i];
            var currentLayer = originalText.layer;

            var basePath = resolveBasePathForText(selectedPaths, targetItems.length, i, originalText);
            if (!basePath) {
                /* 事前に確かめているので通常は来ない / Normally unreachable after the availability check */
                if (showAlerts) alert(getLabel("alert.needPath"));
                return;
            }

            if (!previewMode) {
                makeOriginalPathInvisible(basePath);
            } else {
                /* プレビュー：元のパスを見えるガイドにする（プレビューを消すと戻る） / Preview: show the base path as a guide */
                applyPreviewPathStyle(basePath);
            }

            var textPath = duplicatePathForText(basePath, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(getLabel("alert.dupPathFail"));
                continue;
            }
            if (previewMode) {
                previewTempItems.push(textPath);
            } else {
                applyExecutePathStyle(textPath);
            }

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);
            decoratePathText(textOnAPath, originalText, previewMode);
            createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(createdTexts);
    }

    /**
     * アーチ／正円モードで一緒に選択したパスを、プレビューでは隠し、実行では削除する
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function handleArcModeSelectedPaths(previewMode) {
        if (!selectedPaths || selectedPaths.length === 0) return;
        for (var i = selectedPaths.length - 1; i >= 0; i--) {
            var selectedPath = selectedPaths[i];
            if (!selectedPath) continue;
            if (previewMode) {
                hideOriginalForPreview(selectedPath);
            } else {
                try { selectedPath.remove(); } catch (e) { }
            }
        }
    }

    /**
     * 処理：アーチ状のパスを作ってパス上文字にする
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function createPathTextOnArc(showAlerts, previewMode) {
        var createdTexts = [];
        handleArcModeSelectedPaths(previewMode);

        for (var i = 0; i < targetItems.length; i++) {
            var originalText = targetItems[i];
            var currentLayer = originalText.layer;

            var textPath = createArcPathFromText(originalText, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(getLabel("alert.arcFail"));
                continue;
            }

            /* 生成したアーチは見えないガイド / The generated arc is an invisible guide */
            applyInvisiblePathStyle(textPath);
            if (previewMode) previewTempItems.push(textPath);

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);

            /* 変換時に Illustrator がスタイルを上書きするので、パス上文字のパスも見えなくする
               Illustrator restyles the path on conversion, so hide the path-text path too */
            try {
                if (textOnAPath.textPath) applyInvisiblePathStyle(textOnAPath.textPath);
            } catch (e) { }

            decoratePathText(textOnAPath, originalText, previewMode);
            createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(createdTexts);
    }

    /**
     * 処理：テキストの幅を直径とする正円を作ってパス上文字にする
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function createPathTextOnCircle(showAlerts, previewMode) {
        var createdTexts = [];
        handleArcModeSelectedPaths(previewMode);

        for (var i = 0; i < targetItems.length; i++) {
            var originalText = targetItems[i];
            var currentLayer = originalText.layer;

            var textPath = createCirclePathFromText(originalText, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(getLabel("alert.arcFail"));
                continue;
            }

            if (previewMode) {
                applyPreviewPathStyle(textPath);
                previewTempItems.push(textPath);
            } else {
                applyExecutePathStyle(textPath);
            }

            /* 正円を中心まわりに 90° 回す / Rotate the circle 90° around its center */
            try { textPath.rotate(90, true, true, true, true, Transformation.CENTER); } catch (e) { }

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);

            /* 変換時に Illustrator がスタイルを上書きするので、パス上文字のパスを整え直す
               Illustrator restyles the path on conversion, so restyle the path-text path */
            try {
                if (textOnAPath.textPath) {
                    if (previewMode) {
                        applyPreviewPathStyle(textOnAPath.textPath);
                    } else {
                        applyExecutePathStyle(textOnAPath.textPath);
                    }
                }
            } catch (e) { }

            /* 正円は常に中央揃え。内側配置のときは開始／終了位置を +0.5 補正 / Circles are always centered */
            decoratePathText(textOnAPath, originalText, previewMode, {
                forceCenter: true,
                circleInsideShift: cbReverse.value
            });
            createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(createdTexts);
    }

    /**
     * 1件のパス上文字を、パスとポイント文字に分ける
     * @param {TextFrame} pathText - 対象のパス上文字
     * @param {boolean} previewMode - プレビューなら true
     * @param {boolean} keepFormatting - 書式を保つなら true
     * @returns {boolean} 処理したら true
     */
    function splitOnePathText(pathText, previewMode, keepFormatting) {
        if (!pathText || pathText.typename !== "TextFrame") return false;

        var isPathText = false;
        try { isPathText = (pathText.kind === TextType.PATHTEXT); } catch (e) { isPathText = false; }
        if (!isPathText) return false;

        var originalPath = null;
        try { originalPath = pathText.textPath; } catch (e) { originalPath = null; }
        if (!originalPath) return false;

        var pathTextLayer = null;
        try { pathTextLayer = pathText.layer; } catch (e) { pathTextLayer = null; }

        /* 1) 「パスを削除」がオフなら、ハンドルを保ってパスを複製（塗りなし・スミ1pt） / Keep a copy of the path */
        if (!cbSplitDeletePath.value) {
            var newPath = duplicatePathWithHandles(originalPath, pathTextLayer);
            if (newPath) {
                try {
                    newPath.filled = false;
                    newPath.stroked = true;
                    newPath.strokeColor = makeBlackCMYK();
                    newPath.strokeWidth = 1;
                } catch (e) { }
                if (previewMode) previewTempItems.push(newPath);
            }
        }

        /* 2) 元の開始アンカーの位置にポイント文字を作る / Create point text at the first anchor */
        var newText = null;
        try {
            newText = (pathTextLayer ? pathTextLayer.textFrames.add() : doc.textFrames.add());
        } catch (e) {
            newText = null;
        }

        if (newText) {
            try {
                var anchorPoint = originalPath.pathPoints[0].anchor;
                newText.position = [anchorPoint[0], anchorPoint[1]];
            } catch (e) { }

            /* 内容のコピー（書式を保つ／保たない） / Copy the content (with or without formatting) */
            if (keepFormatting) {
                try { pathText.textRange.duplicate(newText); } catch (e) { }
            } else {
                try { newText.contents = pathText.contents; } catch (e) { }
            }

            applyFontSizeDelta(newText);
            applyShiftAndTrackingDelta(newText);

            /* 書式を保つときだけ塗り／線の色を引き継ぐ / Carry over fill/stroke colors only when keeping formatting */
            if (keepFormatting) {
                for (var i = 0; i < pathText.textRanges.length; i++) {
                    try {
                        var sourceAttributes = pathText.textRanges[i].characterAttributes;
                        var newAttributes = newText.textRanges[i].characterAttributes;
                        newAttributes.fillColor = sourceAttributes.fillColor;
                        newAttributes.strokeColor = sourceAttributes.strokeColor;
                    } catch (e) { }
                }
            }

            /* 行揃え（複製のあと）と重ね順の保持 / Justification after the copy, and keep the stacking order */
            setFrameJustification(newText, readSelectedJustification());
            try { newText.move(pathText, ElementPlacement.PLACEBEFORE); } catch (e) { }

            if (previewMode) previewTempItems.push(newText);
        }

        /* 3) ポイント文字を作れたときだけ元のパス上文字を隠す／削除（失敗したら元を残す） / Keep the original on failure */
        if (newText) {
            if (previewMode) {
                hideOriginalForPreview(pathText);
            } else {
                try { pathText.remove(); } catch (e) { }
            }
        }
        return true;
    }

    /**
     * パス上文字を、見えるパスとポイント文字に分ける
     * @param {boolean} showAlerts - 何も処理しなかったとき alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @param {boolean} keepFormatting - 書式を保つなら true
     * @returns {void}
     */
    function splitPathTextAndPath(showAlerts, previewMode, keepFormatting) {
        var didAny = false;
        for (var i = targetItems.length - 1; i >= 0; i--) {
            if (splitOnePathText(targetItems[i], previewMode, keepFormatting)) didAny = true;
        }

        if (!didAny && showAlerts) {
            alert(getLabel("alert.needPathText"));
        }
    }

    // =========================================
    // パスの生成 / Path generation
    // =========================================

    /**
     * テキストの範囲をアウトライン化して測る（ライブテキストの geometricBounds より安定する）
     * @param {TextFrame} originalText - 対象のテキスト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function measureArcTextBounds(originalText) {
        var measureTexts = [originalText.duplicate(), originalText.duplicate()];
        measureTexts[0].contents = "";

        /* 先頭行の textRange を measureTexts[0] に複製（ポイント文字＋1行以上が前提の旧来の手法）
           Copy the first line's textRanges into measureTexts[0] (legacy approach) */
        for (var i = 0; i < originalText.lines[0].length; i++) {
            originalText.textRanges[i].duplicate(measureTexts[0]);
        }
        for (var j = 0; j < measureTexts.length; j++) {
            measureTexts[j] = measureTexts[j].createOutline();
        }

        var textBounds = measureTexts[1].geometricBounds;
        textBounds[3] = measureTexts[0].geometricBounds[3];

        for (var k = 0; k < measureTexts.length; k++) {
            try { measureTexts[k].remove(); } catch (e) { }
        }
        return textBounds;
    }

    /**
     * ［まるみ］スライダーを読む（0〜100、既定 50）
     * @returns {number} まるみ（%）
     */
    function readArcRoundnessPercent() {
        var percent = Number(slArcRoundness.value);
        if (isNaN(percent)) percent = 50;
        if (percent < 0) percent = 0;
        if (percent > 100) percent = 100;
        return percent;
    }

    /**
     * アーチ方向の符号を返す（上=+1、下=-1）
     * @returns {number} 符号
     */
    function readArcDirectionSign() {
        return rbArcDown.value ? -1 : 1;
    }

    /**
     * 直線のパスのハンドルを動かしてアーチ状に曲げる
     * @param {PathItem} arcPath - 2点の直線パス
     * @param {number} roundnessPercent - まるみ（%）
     * @param {number} directionSign - 上=+1、下=-1
     * @returns {void}
     */
    function applyArcHandles(arcPath, roundnessPercent, directionSign) {
        var handleOffsetDivisors = [3.5, 4];
        var pathLength = arcPath.length;

        /* handleOffsets[0]: 水平ハンドル（固定）/ handleOffsets[1]: 垂直ハンドル＝カーブの深さ
           50% で垂直ハンドルは pathLength / handleOffsetDivisors[1]（旧来の既定） */
        var handleOffsets = [
            pathLength / handleOffsetDivisors[0],
            pathLength * (roundnessPercent / 100) * (2 / handleOffsetDivisors[1])
        ];
        var pathPoints = arcPath.pathPoints;

        pathPoints[0].rightDirection = [
            pathPoints[0].rightDirection[0] + handleOffsets[0],
            pathPoints[0].rightDirection[1] + (directionSign * handleOffsets[1])
        ];
        pathPoints[1].leftDirection = [
            pathPoints[1].leftDirection[0] - handleOffsets[0],
            pathPoints[1].leftDirection[1] + (directionSign * handleOffsets[1])
        ];
    }

    /**
     * テキストのアウトラインの範囲からアーチ状のパスを作る
     * @param {TextFrame} originalText - 対象のテキスト
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @returns {PathItem|null} 作ったパス
     */
    function createArcPathFromText(originalText, targetLayer) {
        try {
            if (!originalText || originalText.typename !== "TextFrame") return null;
            if (!originalText.lines || originalText.lines.length === 0) return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;

            var textBounds = measureArcTextBounds(originalText);

            /* ベースラインに沿った直線を作る（1.02 は旧来の既定） / Straight line along the baseline */
            var baselineY = textBounds[3] * 1.02;
            var arcPath = targetLayer.pathItems.add();
            arcPath.setEntirePath([
                [textBounds[0], baselineY],
                [textBounds[2], baselineY]
            ]);

            try {
                arcPath.stroked = false;
                arcPath.filled = false;
            } catch (e) { }

            applyArcHandles(arcPath, readArcRoundnessPercent(), readArcDirectionSign());
            return arcPath;
        } catch (e) {
            return null;
        }
    }

    /**
     * テキストの幅を直径とする正円のパスを作る
     * @param {TextFrame} originalText - 対象のテキスト
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @returns {PathItem|null} 作ったパス
     */
    function createCirclePathFromText(originalText, targetLayer) {
        try {
            if (!originalText || originalText.typename !== "TextFrame") return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;

            /* 安定させるためアウトラインの範囲で測る / Measure the outline for stability */
            var measureText = originalText.duplicate();
            var measureOutline = null;
            try { measureOutline = measureText.createOutline(); } catch (e) { measureOutline = null; }
            /* createOutline() は複製を消費するので通常ここは例外になる / createOutline() consumes the duplicate */
            try { measureText.remove(); } catch (e) { }
            if (!measureOutline) return null;

            var outlineBounds = measureOutline.geometricBounds;
            try { measureOutline.remove(); } catch (e) { }

            var textWidth = outlineBounds[2] - outlineBounds[0];
            if (!textWidth || isNaN(textWidth) || textWidth <= 0) return null;

            var centerX = (outlineBounds[0] + outlineBounds[2]) / 2;
            var centerY = (outlineBounds[1] + outlineBounds[3]) / 2;

            var diameter = textWidth;
            var circlePath = targetLayer.pathItems.ellipse(centerY + (diameter / 2), centerX - (diameter / 2), diameter, diameter);

            /* 既定は見えない状態（プレビューの見た目は呼び出し側で付ける） / Invisible by default */
            try {
                circlePath.stroked = false;
                circlePath.filled = false;
            } catch (e) { }

            return circlePath;
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // パス幅に合わせる / Fit to path width
    // =========================================

    /**
     * ［パス幅に合わせる］の選択に応じてフィットする
     * @param {TextFrame[]} textFrames - 作成したパス上文字
     * @returns {void}
     */
    function applyFitToPathWidth(textFrames) {
        if (rbFitWidthFontSize.value) {
            fitTextToOpenPath(textFrames);
        } else if (rbFitWidthTracking.value) {
            fitTextToOpenPathByTracking(textFrames);
        }
    }

    /**
     * パスが閉じているか判定する
     * @param {PageItem} pathItem - 対象のパス
     * @returns {boolean} 閉じていれば true
     */
    function isClosedPathItem(pathItem) {
        try {
            if (!pathItem) return false;
            if (pathItem.typename === "CompoundPathItem") {
                if (pathItem.pathItems && pathItem.pathItems.length > 0) return !!pathItem.pathItems[0].closed;
                return false;
            }
            if (pathItem.typename === "PathItem") return !!pathItem.closed;
        } catch (e) { }
        return false;
    }

    /**
     * フィットの対象（編集できるオープンパス上のパス上文字）か判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 対象なら true
     */
    function isTargetPathText(textFrame) {
        try {
            if (!textFrame || textFrame.typename !== "TextFrame") return false;
            if (textFrame.kind !== TextType.PATHTEXT) return false;
            if (!textFrame.editable || textFrame.locked || textFrame.hidden) return false;
            var textPath = textFrame.textPath;
            if (!textPath) return false;
            if (isClosedPathItem(textPath)) return false;
            return true;
        } catch (e) { }
        return false;
    }

    /**
     * 表示されている行数を返す（最低1）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} 行数
     */
    function getLineAmount(textFrame) {
        try {
            if (textFrame.lines && textFrame.lines.length > 0) return textFrame.lines.length;
        } catch (e) { }
        return 1;
    }

    /**
     * テキストがあふれているか判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} lineAmount - 表示する行数
     * @returns {boolean} あふれていれば true
     */
    function isOverset(textFrame, lineAmount) {
        try {
            if (!textFrame) return false;
            if (textFrame.lines.length > 0) {
                var charactersOnVisibleLines = 0;
                if (typeof (lineAmount) === "undefined" || lineAmount === null) {
                    lineAmount = 1;
                } else {
                    lineAmount = Math.floor(lineAmount);
                    if (lineAmount < 1) lineAmount = 1;
                    if (lineAmount > textFrame.lines.length) lineAmount = textFrame.lines.length;
                }
                for (var i = 0; i < lineAmount; i++) {
                    charactersOnVisibleLines += textFrame.lines[i].characters.length;
                }
                return (charactersOnVisibleLines < textFrame.characters.length);
            } else if (textFrame.characters.length > 0) {
                return true;
            }
        } catch (e) { }
        return false;
    }

    /**
     * オープンパスの端まで、文字サイズだけでパス上文字を合わせる
     * あふれていれば少しずつ縮め、収まっていれば一度あふれるまで広げてから縮める（ほぼ最大）
     * @param {TextFrame[]} textFrames - 対象のパス上文字
     * @returns {boolean} 対象があれば true
     */
    function fitTextToOpenPath(textFrames) {
        if (!textFrames || textFrames.length === 0) return false;

        var fitOptions = {
            increment: 0.1,
            minFontSize: 0.1,
            maxShrinkIter: 2000,
            maxGrowIter: 10,
            maxFontSize: 2000
        };

        /**
         * あふれるまで拡大してから、収まるまで縮小する
         * @param {TextFrame} textFrame - 対象のパス上文字
         * @returns {void}
         */
        function shrinkFont(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);

                /* 収まっていれば、あふれるまで倍にしていく（上限あり） / Grow by doubling until it overflows */
                if (!isOverset(textFrame, lineAmount)) {
                    var growIterations = 0;
                    while (!isOverset(textFrame, lineAmount) && growIterations < fitOptions.maxGrowIter) {
                        var growSize = textFrame.textRange.characterAttributes.size;
                        if (growSize >= fitOptions.maxFontSize) break;
                        textFrame.textRange.characterAttributes.size = Math.min(fitOptions.maxFontSize, growSize * 2);
                        growIterations++;
                    }
                }

                /* 収まるまで少しずつ縮める / Then shrink in small steps until it fits */
                var shrinkIterations = 0;
                while (isOverset(textFrame, lineAmount)) {
                    var currentSize = textFrame.textRange.characterAttributes.size;
                    if (currentSize <= fitOptions.minFontSize) break;

                    textFrame.textRange.characterAttributes.size = Math.max(fitOptions.minFontSize, currentSize - fitOptions.increment);

                    shrinkIterations++;
                    if (shrinkIterations >= fitOptions.maxShrinkIter) break;
                }
            } catch (e) { }
        }

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!isTargetPathText(textFrame)) continue;
            shrinkFont(textFrame);
        }

        return true;
    }

    /**
     * オープンパスの端まで、トラッキングだけで合わせる（文字サイズは保つ）
     * @param {TextFrame[]} textFrames - 対象のパス上文字
     * @returns {boolean} 対象があれば true
     */
    function fitTextToOpenPathByTracking(textFrames) {
        if (!textFrames || textFrames.length === 0) return false;

        var fitOptions = {
            coarseStep: 50,     /* 粗調整1ステップのトラッキング量 / tracking units per coarse step */
            fineStep: 1,        /* 微調整1ステップのトラッキング量 / tracking units per fine step */
            minTracking: -1000, /* 累積トラッキングの下限 / tightest allowed cumulative delta */
            maxTracking: 20000, /* 累積トラッキングの上限 / loosest allowed cumulative delta */
            maxIter: 4000
        };

        /**
         * 1つのパス上文字をトラッキングでパス幅に合わせる
         * @param {TextFrame} textFrame - 対象のパス上文字
         * @returns {void}
         */
        function fitByTracking(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);
                var appliedTracking = 0; /* これまでに加えた累積トラッキング / cumulative tracking delta applied so far */
                var iterations;

                if (isOverset(textFrame, lineAmount)) {
                    /* あふれている：収まるまで粗く詰める / too wide: tighten (coarse) until it fits */
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking - fitOptions.coarseStep < fitOptions.minTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", -fitOptions.coarseStep);
                        appliedTracking -= fitOptions.coarseStep;
                        iterations++;
                    }
                    /* 再びあふれるまで細かく広げる / loosen back (fine) until it overflows again */
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking + fitOptions.fineStep > fitOptions.maxTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", fitOptions.fineStep);
                        appliedTracking += fitOptions.fineStep;
                        iterations++;
                    }
                    /* 1ステップ行き過ぎたら戻して収める / stepped one step too far: pull back once so it fits */
                    if (isOverset(textFrame, lineAmount)) {
                        addDeltaToTextRanges(textFrame, "tracking", -fitOptions.fineStep);
                        appliedTracking -= fitOptions.fineStep;
                    }
                } else {
                    /* 余裕あり：あふれるまで粗く広げる / fits with room: loosen (coarse) until it overflows */
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking + fitOptions.coarseStep > fitOptions.maxTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", fitOptions.coarseStep);
                        appliedTracking += fitOptions.coarseStep;
                        iterations++;
                    }
                    /* 収まるまで細かく詰める / tighten back (fine) until it fits */
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking - fitOptions.fineStep < fitOptions.minTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", -fitOptions.fineStep);
                        appliedTracking -= fitOptions.fineStep;
                        iterations++;
                    }
                }
            } catch (e) { }
        }

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!isTargetPathText(textFrame)) continue;
            fitByTracking(textFrame);
        }

        return true;
    }
}());
