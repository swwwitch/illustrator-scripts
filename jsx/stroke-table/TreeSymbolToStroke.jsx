#target illustrator
#targetengine "TreeSymbolToStrokeEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストのツリー記号（├─・└─・│ や tree コマンドの ├── など）をパスの罫線に変換し、字下げと名前・説明のあいだの空きをタブにします。
各行を囲み罫で囲むこともでき、罫線・タブストップ・囲み罫はプレビューを見ながら調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TreeSymbolToStroke.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nc961754b7cad

### Overview

Converts the tree symbols (├─, └─, │, the ├── from the tree command and so on) in the selected text into stroked paths,
and turns indents and the gaps between names and descriptions into tabs. Each line can also be boxed; the lines, tab stops
and boxes are adjusted with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TreeSymbolToStroke.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TreeSymbolToStroke";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TreeSymbolToStroke.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TreeSymbolToStroke.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc961754b7cad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_STROKE_WIDTH_PT = 1;              /* 線幅の初期値（pt） / default line weight in points */
    var LINE_COLOR_CMYK         = [0, 0, 0, 100]; /* 罫線の色（CMYK） / line color */
    var ROUND_CAP               = false;          /* 線端の初期値。true で丸型（角の形状はラウンド）、false でなし（マイター） / default cap */
    var LINE_START_PLACEHOLDER  = " ";            /* 行頭の記号を消すときに1文字だけ残す文字（カーニングで元の幅にする） / character left for a symbol at a line start */
    var CONVERT_INDENTS         = true;           /* ［行頭の字下げをタブに変換］の初期値 / default of the indent option */
    var LINK_LEVELS_EVENLY      = false;          /* レベルの［均等に連動］の初期値 / default of the even level link */
    var LINK_STEMS_EVENLY       = false;          /* 縦罫の［均等に連動］の初期値 / default of the even stem link */
    var CONVERT_SPACE_RUNS      = true;           /* ［連続したスペースをタブに変換］の初期値 / default of the space-run option */
    var DRAW_BOX                = false;          /* ［囲み罫］の初期値 / default of the box option */
    var BOX_MARGIN_SIZE_RATIO   = 1 / 8;          /* 囲み罫のマージン（左右・上下）の初期値。文字サイズに掛ける割合 / default box margins as a share of the type size */
    var LINK_BOX_MARGINS        = true;           /* 囲み罫のマージンの［連動］の初期値 / default of the margin link */
    var ATTACH_ARMS_TO_BOX      = true;           /* ［横線を囲み罫につなげる］の初期値 / default of attaching the arms to the boxes */
    var ALIGN_BOX_RIGHT         = true;           /* ［右端を揃える］の初期値 / default of the right-edge alignment of the boxes */
    var SPLIT_BOXED_LINES       = false;          /* ［行ごとに分割］の初期値 / default of splitting the boxed lines */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var NUMBER_FIELD_CHARS         = 6;  /* 数値欄の幅（文字数） / width of the number fields in characters */
    var FIELD_LABEL_WIDTH          = 70; /* 数値欄の項目名の幅（右揃え） / width of the right-aligned field labels */
    var COLOR_SWATCH_SIZE          = 18; /* 色見本の大きさ（px） / swatch size */
    var COLOR_HEX_FIELD_CHARS      = 6;  /* 16進数の欄の幅（文字数） / width of the hex field in characters */
    var COLOR_HEX_LEFT_MARGIN      = 10; /* 色見本と「#」のあいだの余白（px） / gap between the swatch and "#" */
    var CAP_RADIO_SPACING          = 6;  /* 線端のラジオボタンの間隔 / spacing between the cap radio buttons */
    var CHECKBOX_SPACING           = 6;  /* チェックボックスを縦に並べる間隔 / spacing between stacked checkboxes */
    var COMPACT_BUTTON_FONT_SHRINK = 2;  /* パネル内の小さいボタン（リセット）の文字を小さくする量（pt） / font shrink for compact buttons */
    var COMPACT_BUTTON_TRIM_PX     = 4;  /* 小さいボタンの高さを詰める量（px） / height trim for compact buttons */
    var NEIGHBOR_PUSH_GAP          = 1;  /* タブストップ・縦罫が隣に追いついたとき、押し出して空ける差（定規の単位） / gap left when a value pushes its neighbor, in ruler units */

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

    // 日英ラベル定義 / Japanese-English label definitions
    var LABELS = {
        dialog: {
            title: { ja: "ツリー記号を罫線に変換", en: "Tree Symbols to Lines" }
        },
        panel: {
            lines: { ja: "罫線", en: "Lines" },
            text: { ja: "テキスト", en: "Text" },
            tabStops: { ja: "タブストップ", en: "Tab Stops" },
            stems: { ja: "縦罫", en: "Vertical Rules" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Weight" },
            lineColor: { ja: "カラー", en: "Color" },
            strokeCap: { ja: "線端", en: "Cap" },
            hexPrefix: { ja: "#", en: "#" },
            leading: { ja: "行送り", en: "Leading" },
            levelStop: { ja: "レベル %1", en: "Level %1" },
            descriptionStop: { ja: "説明", en: "Description" },
            stemPosition: { ja: "縦罫 %1", en: "Rule %1" },
            boxMarginX: { ja: "左右", en: "Left/Right" },
            boxMarginY: { ja: "上下", en: "Top/Bottom" }
        },
        radio: {
            roundCap: { ja: "丸型", en: "Round" },
            buttCap: { ja: "なし", en: "Butt" }
        },
        checkbox: {
            convertIndents: { ja: "行頭の字下げをタブに変換", en: "Convert indents to tabs" },
            convertSpaceRuns: { ja: "連続したスペースをタブに変換", en: "Convert space runs to tabs" },
            drawBox: { ja: "囲み罫", en: "Box" },
            attachArmsToBox: { ja: "横線を囲み罫につなげる", en: "Connect arms to boxes" },
            alignBoxRight: { ja: "右端を揃える", en: "Align right edges" },
            splitBoxedLines: { ja: "行ごとに分割", en: "Split into lines" }
        },
        tooltip: {
            leading: {
                ja: "テキスト全体の行送りです。変えると固定値になります。",
                en: "Leading for the whole text. Changing it sets a fixed value."
            },
            levelStop: {
                ja: "このレベルの名前をそろえるタブストップの位置です。変えるとその位置に固定します。",
                en: "Tab stop that aligns the names at this level. Changing it fixes the stop at that position."
            },
            descriptionStop: {
                ja: "説明をそろえるタブストップの位置です。変えるとその位置に固定します。",
                en: "Tab stop that aligns the descriptions. Changing it fixes the stop at that position."
            },
            stemPosition: {
                ja: "この縦罫の横位置です（タブ位置と同じ基準）。変えるとその位置に固定します。",
                en: "Horizontal position of this vertical rule, measured like the tab stops. Changing it fixes the rule there."
            },
            roundCap: {
                ja: "線の端を丸くします。角の形状もラウンドになります。",
                en: "Rounds the line ends; corners become round joins."
            },
            buttCap: {
                ja: "線の端を切りっぱなしにします。角の形状はマイターになります。",
                en: "Leaves the line ends flat; corners become miter joins."
            },
            lineColorSwatch: {
                ja: "クリックするとカラーピッカーで罫線の色を選べます。",
                en: "Click to choose the line color in the Color Picker."
            },
            lineColorHex: { ja: "罫線の色を16進数（RRGGBB）で入力します。", en: "Line color as a hex value (RRGGBB)." },
            resetTabStops: {
                ja: "タブストップを自動の位置に戻し、均等に連動を外します。",
                en: "Returns the tab stops to their automatic positions and turns off Link evenly."
            },
            resetStems: {
                ja: "縦罫を自動の位置に戻し、均等に連動を外します。",
                en: "Returns the rules to their automatic positions and turns off Link evenly."
            },
            linkStems: {
                ja: "均等に連動：オンにすると、縦罫1と2の差の間隔で、縦罫3以降を並べます。",
                en: "Link evenly: when on, rules 3 and later continue at the spacing between rules 1 and 2."
            },
            linkLevels: {
                ja: "均等に連動：オンにすると、レベル1と2の差の間隔で、レベル3以降を並べます。",
                en: "Link evenly: when on, levels 3 and later continue at the spacing between levels 1 and 2."
            },
            showHiddenChars: {
                ja: "タブや改行などの制御文字の表示を切り替えます。",
                en: "Toggles hidden characters such as tabs and returns."
            },
            convertIndents: {
                ja: "ツリー記号を含む行頭の字下げをタブにし、名前の元の位置にタブストップを置きます。オフのときは記号の跡をカーニングで埋めます。",
                en: "Turns indents that contain tree symbols into tabs, with tab stops at the original name positions. When off, the gaps are filled with kerning."
            },
            convertSpaceRuns: {
                ja: "文字のあとに続く2つ以上の半角スペースをタブにし、説明がそろう位置（いちばん右の説明）にタブストップを置きます。",
                en: "Turns two or more spaces after text into a tab, with a tab stop that lines up the descriptions at the rightmost one."
            },
            drawBox: {
                ja: "各行を長方形の罫線で囲みます。線幅・線端・カラーは罫線と同じです。横線は囲みの左辺で止めます。",
                en: "Surrounds each line with a rectangle in the same weight, cap and color as the lines. The arms stop at the box."
            },
            boxMarginX: {
                ja: "文字から囲み罫の左右の辺までの間隔です。",
                en: "Distance from the characters to the left and right sides of the box."
            },
            boxMarginY: {
                ja: "文字から囲み罫の上下の辺までの間隔です。",
                en: "Distance from the characters to the top and bottom of the box."
            },
            linkBoxMargins: {
                ja: "連動：オンにすると、左右と上下のマージンを同じ値にします。",
                en: "Link: when on, the left/right and top/bottom margins share one value."
            },
            attachArmsToBox: {
                ja: "オンにすると、横線を囲み罫の左辺まで伸ばしてつなげます。［テキスト］の［囲み罫］がオンのときに使えます。",
                en: "When on, the arms run all the way to the left side of the boxes. Available when Box is on in the Text panel."
            },
            alignBoxRight: {
                ja: "オンにすると、囲み罫の右端をいちばん長い行にそろえます。オフのときは行ごとの文字の右端で囲みます。",
                en: "When on, every box extends to the longest line. When off, each box ends at its own line."
            },
            splitBoxedLines: {
                ja: "テキストを行ごとに分割し、それぞれの囲み罫とグループ化します。",
                en: "Splits the text into lines and groups each line with its box."
            },
            strokeWidth: {
                ja: "描く罫線の線幅です。単位は環境設定の［単位］の［線］に従います。",
                en: "Weight of the lines to draw, in the Stroke units set in Preferences > Units."
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
            showHiddenChars: { ja: "制御文字", en: "Hidden Characters" },
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            noTextFrame: { ja: "テキストを選択してください。", en: "Select some text." },
            noSymbol: {
                ja: "選択したテキストに ├─・└─・│ も、連続したスペースもありません。",
                en: "The selected text has no ├─, └─, │ or runs of spaces."
            },
            measureFailed: {
                ja: "%1 個のテキストは文字の位置を測れなかったため、変換していません。",
                en: "%1 text object(s) were skipped because the character positions could not be measured."
            }
        },
        fallbackName: {
            lineGroup: { ja: "ツリー罫線", en: "Tree Lines" }
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
        fieldLabel.addEventListener("click", function () {
            numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
            numberInput.active = true;
        });

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
    // ツリー記号 / Tree symbols
    // =========================================

    /* 角の記号（├ は縦線が下まで続く、└ は横線で止まる） / corner symbols (├ continues down, └ stops at the arm) */
    var CORNER_SYMBOLS = { "├": "tee", "└": "elbow", "┣": "tee", "┗": "elbow" };
    /* 角の記号に続けて1組にする横線（字形が文字枠の両端まで届く罫線素片）。長音の「ー」は文字として残すので入れない
       Horizontal bars paired with a corner (box-drawing glyphs that span the whole cell). The long vowel ー stays as text */
    var HORIZONTAL_BARS = { "─": true, "━": true };
    /* 縦線の記号 / vertical bar symbols */
    var VERTICAL_BARS = { "│": true, "┃": true };

    /**
     * 文字列からツリー記号の位置を拾う（├─・└─ は角と続く横線をまとめて1個（tree コマンドの ├── も1個）、│ は1文字で1個）
     * @param {string} textContents - テキストフレームの contents
     * @returns {Object[]} { index, kind（"tee" / "elbow" / "vertical"）, length } の配列（前から順）
     */
    function findTreeSymbols(textContents) {
        var symbols = [];
        for (var i = 0; i < textContents.length; i++) {
            var currentChar = textContents.charAt(i);
            var cornerKind = CORNER_SYMBOLS[currentChar];
            var barCount = 0;
            if (cornerKind) {
                while (HORIZONTAL_BARS.hasOwnProperty(textContents.charAt(i + 1 + barCount))) barCount++;
            }
            if (barCount > 0) {
                symbols.push({ index: i, kind: cornerKind, length: 1 + barCount });
                i += barCount;
            } else if (VERTICAL_BARS[currentChar]) {
                symbols.push({ index: i, kind: "vertical", length: 1 });
            }
        }
        return symbols;
    }

    /**
     * いずれかのテキストフレームに、ツリー記号か連続したスペースがあるかを返す
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {boolean} 1個でもあれば true
     */
    function hasConvertibleText(textFrames) {
        for (var i = 0; i < textFrames.length; i++) {
            var textContents = textFrames[i].contents;
            if (findTreeSymbols(textContents).length > 0 || findSpaceRuns(textContents).length > 0) return true;
        }
        return false;
    }

    // =========================================
    // 文字位置の実測 / Glyph measurement
    // =========================================

    /**
     * 空白（字形を持たない文字）かどうかを返す
     * @param {string} oneChar - 1文字
     * @returns {boolean} 空白なら true
     */
    function isBlankChar(oneChar) {
        return /[\s\u00A0\u3000\u200B]/.test(oneChar);
    }

    /**
     * アウトライン化した結果から字形を集める。
     * 字形は1つのグループに CompoundPathItem が平らに並び、空白には何も作られない（2026-10-01 実測）。
     * 入れ子のグループがあっても末端までたどる
     * @param {PageItem} outlinedItem - createOutline() の戻り値、またはその子
     * @param {PageItem[]} glyphItems - 集めた字形を入れる配列
     * @returns {void}
     */
    function collectOutlinedGlyphs(outlinedItem, glyphItems) {
        if (outlinedItem.typename !== "GroupItem") {
            glyphItems.push(outlinedItem);
            return;
        }
        for (var i = 0; i < outlinedItem.pageItems.length; i++) collectOutlinedGlyphs(outlinedItem.pageItems[i], glyphItems);
    }

    /**
     * 字形の並びが読み順の逆（最後の文字が先頭）かを返す。
     * 実測では逆順に並ぶが、先頭と末尾の位置で確かめる（上の行が先、同じ行なら左が先）
     * @param {PageItem[]} glyphItems - collectOutlinedGlyphs() で集めた字形
     * @returns {boolean} 逆順なら true
     */
    function isGlyphOrderReversed(glyphItems) {
        if (glyphItems.length < 2) return false;
        var firstBounds = glyphItems[0].geometricBounds;
        var lastBounds = glyphItems[glyphItems.length - 1].geometricBounds;
        var firstMiddleY = (firstBounds[1] + firstBounds[3]) / 2;
        var lastMiddleY = (lastBounds[1] + lastBounds[3]) / 2;
        /* 同じ行と見なす高さの差は、字形の高さの半分まで / same line when the centers are within half a glyph height */
        if (Math.abs(firstMiddleY - lastMiddleY) > (firstBounds[1] - firstBounds[3]) / 2) return firstMiddleY < lastMiddleY;
        return firstBounds[0] > lastBounds[0];
    }

    /**
     * 字形を持つ文字の番号を、前から順に返す（改行・空白・サロゲートペアの後半を除く）
     * @param {string} textContents - テキストフレームの contents
     * @returns {number[]} 文字の番号
     */
    function listGlyphCharIndices(textContents) {
        var glyphCharIndices = [];
        for (var i = 0; i < textContents.length; i++) {
            var oneChar = textContents.charAt(i);
            var charCode = oneChar.charCodeAt(0);
            if (isLineBreakChar(oneChar) || isBlankChar(oneChar)) continue;
            if (charCode >= 0xDC00 && charCode <= 0xDFFF) continue; /* サロゲートペアの後半 / low surrogate */
            glyphCharIndices.push(i);
        }
        return glyphCharIndices;
    }

    /**
     * テキストの複製をアウトライン化して、指定した文字の字形の境界を測る
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object} wantedIndices - 測る文字の番号をキーにしたオブジェクト
     * @returns {Object|null} 文字の番号 → [左, 上, 右, 下]（pt）。位置を測れなければ null
     */
    function measureGlyphBounds(textFrame, wantedIndices) {
        var duplicatedFrame = textFrame.duplicate();
        var outlinedGroup;
        try {
            outlinedGroup = duplicatedFrame.createOutline();
        } catch (e) {
            /* createOutline() は成功すると複製を消費する。失敗して残った複製だけ片付ける
               createOutline() consumes the duplicate on success; remove the one left behind by a failure */
            try { duplicatedFrame.remove(); } catch (err) { /* すでに無い / already gone */ }
            return null;
        }

        /* 測る途中で例外になっても、アウトラインは必ず消す / always remove the outlines, even when measuring throws */
        try {
            var glyphItems = [];
            collectOutlinedGlyphs(outlinedGroup, glyphItems);
            var glyphCharIndices = listGlyphCharIndices(textFrame.contents);
            /* 字形と文字の数が合わなければ対応が取れない（合字など） / counts must match to pair them up (ligatures break it) */
            if (glyphItems.length !== glyphCharIndices.length) return null;
            if (isGlyphOrderReversed(glyphItems)) glyphItems.reverse();
            var boundsByIndex = {};
            for (var i = 0; i < glyphCharIndices.length; i++) {
                if (wantedIndices[glyphCharIndices[i]]) boundsByIndex[glyphCharIndices[i]] = glyphItems[i].geometricBounds;
            }
            return boundsByIndex;
        } finally {
            outlinedGroup.remove();
        }
    }

    // =========================================
    // 罫線の組み立て / Line geometry
    // =========================================
    var STEM_JOIN_TOLERANCE_PT = 1; /* 縦線どうし・横線との位置を同じと見なす差（pt） / tolerance for joining stems and arms */

    /**
     * ツリー記号1個ぶんの罫線を、点の並びの配列で返す
     * @param {Object} symbol - findTreeSymbols() の要素
     * @param {Object} boundsByIndex - measureGlyphBounds() の戻り値
     * @param {Object} cellMetrics - cellWidth（横線の字形で文字枠を測れないときの1文字の幅 pt）/ leading（行送り pt）
     * @returns {number[][][]} 罫線ごとの点 [[x, y], …] の配列
     */
    function buildSymbolLines(symbol, boundsByIndex, cellMetrics) {
        var glyphBounds = boundsByIndex[symbol.index];
        var glyphLeft = glyphBounds[0];
        var glyphTop = glyphBounds[1];
        var glyphRight = glyphBounds[2];
        var glyphBottom = glyphBounds[3];
        /* 下へ続く縦線は、行送りの分だけ伸ばして次の行の字形の上端につなぐ（行送りが文字より大きいと字形だけでは切れる）
           a bar that continues down spans the leading so it meets the next line's glyph top */
        var continuedBottom = Math.min(glyphBottom, glyphTop - cellMetrics.leading);

        /* 縦線は字形の中央に引く / the vertical runs through the glyph center */
        if (symbol.kind === "vertical") {
            var centerX = (glyphLeft + glyphRight) / 2;
            return [[[centerX, glyphTop], [centerX, continuedBottom]]];
        }

        /* 横線の字形は文字枠の両端まで届くので、その幅と中央の高さを使う。横線が続くときは最後の横線の右端まで
           the bar glyph spans the cell, so take the cell width and arm height from it; the arm runs to the last bar */
        var barCount = symbol.length - 1;
        var barBounds = boundsByIndex[symbol.index + 1];
        var lastBarBounds = boundsByIndex[symbol.index + barCount];
        var cellWidth, lineRight, armY;
        if (barBounds && lastBarBounds) {
            cellWidth = barBounds[2] - barBounds[0];
            lineRight = lastBarBounds[2];
            armY = (barBounds[1] + barBounds[3]) / 2;
        } else {
            cellWidth = cellMetrics.cellWidth;
            lineRight = glyphRight + cellWidth * barCount;
            armY = (symbol.kind === "tee") ? (glyphTop + glyphBottom) / 2 : glyphBottom;
        }
        /* 名前を深さごとにそろえて右へ動かした分だけ、横線を伸ばす / extend the arm to follow a name moved to its depth's position */
        if (symbol.armExtension) lineRight += symbol.armExtension;
        /* 囲み罫があるときは、その左辺で止める（つなげるときは左辺の位置そのもの） / stop at the box's left edge (or end exactly on it) */
        if (symbol.armTrim) lineRight -= symbol.armTrim;
        if (symbol.armEnd !== undefined) lineRight = symbol.armEnd;
        /* 角の記号の右端は文字枠の右端なので、縦線はそこから半文字戻った位置 / the stem sits half a cell left of the corner's right edge */
        var stemX = glyphRight - cellWidth / 2;

        if (symbol.kind === "tee") {
            return [
                [[stemX, glyphTop], [stemX, continuedBottom]],
                [[stemX, armY], [lineRight, armY]]
            ];
        }
        /* └ は1本の折れ線にして角をきれいにつなぐ / draw └ as one polyline so the corner joins cleanly */
        return [[[stemX, glyphTop], [stemX, armY], [lineRight, armY]]];
    }

    /**
     * 記号の縦線を別の横位置へ動かす（縦線の点と、そこから出る横線の左端を一緒に動かす）
     * @param {Object} symbolShape - { kind, lines }
     * @param {number} newStemX - 新しい縦線の横位置（pt）
     * @returns {void}
     */
    function moveStem(symbolShape, newStemX) {
        var oldStemX = symbolShape.lines[0][0][0];
        for (var i = 0; i < symbolShape.lines.length; i++) {
            for (var k = 0; k < symbolShape.lines[i].length; k++) {
                if (Math.abs(symbolShape.lines[i][k][0] - oldStemX) < 0.001) symbolShape.lines[i][k][0] = newStemX;
            }
        }
    }

    /**
     * 縦線を列（記号が元のテキストの何桁目にあるか）ごとに1つの横位置へそろえる。
     * 位置は列の平均。ダイアログで指定した列（stemOverrides）はその位置に固定する
     * @param {Object[]} symbolShapes - { kind, lines, stemKey } の配列
     * @param {number} originX - タブ位置 0 に当たる横位置（pt）
     * @param {Object} stemOverrides - 列のキー → 縦罫の位置（pt、originX から）
     * @param {Object} stemPositions - 列のキー → そろえた縦罫の位置（pt）。ここに足す
     * @returns {void}
     */
    function alignStemsByColumn(symbolShapes, originX, stemOverrides, stemPositions) {
        var stemSums = {};
        var i;
        for (i = 0; i < symbolShapes.length; i++) {
            var stemKey = symbolShapes[i].stemKey;
            if (!stemSums[stemKey]) stemSums[stemKey] = { total: 0, count: 0 };
            stemSums[stemKey].total += symbolShapes[i].lines[0][0][0];
            stemSums[stemKey].count++;
        }
        for (i = 0; i < symbolShapes.length; i++) {
            var shapeKey = symbolShapes[i].stemKey;
            var stemX = stemOverrides.hasOwnProperty(shapeKey)
                ? originX + stemOverrides[shapeKey]
                : stemSums[shapeKey].total / stemSums[shapeKey].count;
            moveStem(symbolShapes[i], stemX);
            if (!stemPositions.hasOwnProperty(shapeKey)) stemPositions[shapeKey] = stemX - originX;
        }
    }

    /**
     * 記号が行頭から何桁目にあるかで、縦罫の列のキーを返す
     * @param {string} textContents - テキストフレームの contents
     * @param {number} symbolIndex - 記号の文字の番号
     * @returns {string} 列のキー
     */
    function getStemKey(textContents, symbolIndex) {
        var lineStart = symbolIndex;
        while (lineStart > 0 && !isLineBreakChar(textContents.charAt(lineStart - 1))) lineStart--;
        return "stem" + countColumns(textContents, lineStart, symbolIndex);
    }

    /**
     * 縦線の上端が上の行の縦線に続いていなければ、上にある前のレベルの横線か、上の行（囲み罫）の下辺まで伸ばしてつなぐ
     * （子の └ が親の ├─ の横線から下りる形にする。いちばん上の縦線は最初の項目から下ろす）
     * @param {Object[]} symbolShapes - { kind, lines }（lines は buildSymbolLines() の戻り値）の配列。点をその場で書き換える
     * @param {number[][]} lineBoxes - 囲み罫（囲み罫が無いときは行）の範囲 [左, 上, 右, 下] の配列
     * @returns {void}
     */
    function connectStemsToParentArms(symbolShapes, lineBoxes) {
        var stems = [];
        var arms = [];
        var i, k;
        for (i = 0; i < symbolShapes.length; i++) {
            var shapeLines = symbolShapes[i].lines;
            /* どの記号も最初の線の先頭が縦線の上端 / the first point of the first line is always the stem top */
            var firstLine = shapeLines[0];
            stems.push({ topPoint: firstLine[0], bottomY: firstLine[1][1] });
            if (symbolShapes[i].kind === "tee") arms.push({ y: shapeLines[1][0][1], left: shapeLines[1][0][0], right: shapeLines[1][1][0] });
            if (symbolShapes[i].kind === "elbow") arms.push({ y: firstLine[1][1], left: firstLine[1][0], right: firstLine[2][0] });
        }
        for (i = 0; i < stems.length; i++) {
            var stemX = stems[i].topPoint[0];
            var stemTop = stems[i].topPoint[1];
            var hasStemAbove = false;
            for (k = 0; k < stems.length && !hasStemAbove; k++) {
                hasStemAbove = (k !== i && Math.abs(stems[k].topPoint[0] - stemX) < STEM_JOIN_TOLERANCE_PT && Math.abs(stems[k].bottomY - stemTop) < STEM_JOIN_TOLERANCE_PT);
            }
            if (hasStemAbove) continue;
            /* 上にあって、この縦線の位置を通る横線のうち、いちばん近いもの。横線の左端（角）に重なるときもつなぐ
               （縦罫を親と同じ位置にすると、└ の角から縦線が続いて ├ の形になる）
               the nearest arm above that passes this x, including its left end (the corner), so └ becomes ├ */
            var parentArmY = null;
            for (k = 0; k < arms.length; k++) {
                var crossesStem = arms[k].left - STEM_JOIN_TOLERANCE_PT < stemX && stemX < arms[k].right + STEM_JOIN_TOLERANCE_PT;
                if (crossesStem && arms[k].y > stemTop && (parentArmY === null || arms[k].y < parentArmY)) parentArmY = arms[k].y;
            }
            /* 横線より近くに上の行（囲み罫）があれば、その下辺で止める / a line or box nearer than any arm takes the stem instead */
            for (k = 0; k < lineBoxes.length; k++) {
                var boxCoversStem = lineBoxes[k][0] < stemX && stemX < lineBoxes[k][2];
                var boxBottom = lineBoxes[k][3];
                if (boxCoversStem && boxBottom > stemTop - STEM_JOIN_TOLERANCE_PT && (parentArmY === null || boxBottom < parentArmY)) parentArmY = boxBottom;
            }
            if (parentArmY !== null) stems[i].topPoint[1] = parentArmY;
        }
    }

    /**
     * 記号の文字から、1文字の幅と行送りを読む
     * @param {TextRange} symbolChar - 記号の1文字
     * @returns {{cellWidth: number, leading: number}} 1文字の幅と行送り（pt）
     */
    function readCellMetrics(symbolChar) {
        var symbolAttributes = symbolChar.characterAttributes;
        /* 自動行送りのときは、段落の自動行送り量（％）から出す / auto leading comes from the paragraph's percentage */
        var leading = symbolAttributes.autoLeading
            ? symbolAttributes.size * symbolChar.paragraphAttributes.autoLeadingAmount / 100
            : symbolAttributes.leading;
        return { cellWidth: symbolAttributes.size * symbolAttributes.horizontalScale / 100, leading: leading };
    }

    /**
     * 点の並びから線だけのパスを作る
     * @param {GroupItem} lineGroup - 追加先のグループ
     * @param {number[][]} points - [[x, y], …]
     * @param {Object} lineStyle - strokeWidth（pt）/ strokeColor / strokeCap / strokeJoin
     * @returns {PathItem} 作ったパス
     */
    function addLinePath(lineGroup, points, lineStyle) {
        var linePath = lineGroup.pathItems.add();
        linePath.setEntirePath(points);
        applyLineStyle(linePath, lineStyle);
        return linePath;
    }

    /**
     * 長方形の罫線（囲み罫）を作る
     * @param {GroupItem} lineGroup - 追加先のグループ
     * @param {number[]} boxBounds - 長方形の範囲 [左, 上, 右, 下]（pt）
     * @param {Object} lineStyle - strokeWidth（pt）/ strokeColor / strokeCap / strokeJoin
     * @returns {PathItem} 作ったパス
     */
    function addBoxPath(lineGroup, boxBounds, lineStyle) {
        var boxPath = lineGroup.pathItems.rectangle(boxBounds[1], boxBounds[0], boxBounds[2] - boxBounds[0], boxBounds[1] - boxBounds[3]);
        applyLineStyle(boxPath, lineStyle);
        return boxPath;
    }

    /**
     * 行ごとの字形の範囲を拾う（前から順。字形の無い行は firstGlyph が -1）
     * @param {string} textContents - テキストフレームの contents
     * @returns {Object[]} { start, firstGlyph, lastGlyph } の配列
     */
    function listTextLines(textContents) {
        var textLines = [];
        var currentLine = { start: 0, firstGlyph: -1, lastGlyph: -1 };
        for (var i = 0; i < textContents.length; i++) {
            var oneChar = textContents.charAt(i);
            if (isLineBreakChar(oneChar)) {
                textLines.push(currentLine);
                currentLine = { start: i + 1, firstGlyph: -1, lastGlyph: -1 };
                continue;
            }
            if (isBlankChar(oneChar)) continue;
            if (currentLine.firstGlyph < 0) currentLine.firstGlyph = i;
            currentLine.lastGlyph = i;
        }
        textLines.push(currentLine);
        return textLines;
    }

    /**
     * 範囲をマージンの分だけ広げる
     * @param {number[]} bounds - [左, 上, 右, 下]（pt）
     * @param {Object} margins - x（左右 pt）/ y（上下 pt）
     * @returns {number[]} 広げた範囲
     */
    function expandBounds(bounds, margins) {
        return [bounds[0] - margins.x, bounds[1] + margins.y, bounds[2] + margins.x, bounds[3] - margins.y];
    }

    /**
     * 変換したあとのテキストの各行を囲む、囲み罫の範囲を決める。
     * 横は行の最初と最後の字形、縦はポイント文字の外形を行送りで行ごとに分けた高さ（1行の仮想ボディ）で囲む。
     * ポイント文字以外は行の位置を出せないので、テキスト全体を1つで囲む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object} boxOptions - margins（x / y、pt）/ alignRight（右端をいちばん長い行にそろえる）
     * @returns {number[][]} 囲み罫の範囲 [左, 上, 右, 下]（マージン込み、pt）の配列（上の行から順）
     */
    function measureLineBoxes(textFrame, boxOptions) {
        var frameBounds = textFrame.geometricBounds;
        if (textFrame.kind !== TextType.POINTTEXT) return [expandBounds(frameBounds, boxOptions.margins)];
        var textContents = textFrame.contents;
        var textLines = listTextLines(textContents);
        var wantedIndices = {};
        var i;
        for (i = 0; i < textLines.length; i++) {
            if (textLines[i].firstGlyph < 0) continue;
            wantedIndices[textLines[i].firstGlyph] = true;
            wantedIndices[textLines[i].lastGlyph] = true;
        }
        var boundsByIndex = measureGlyphBounds(textFrame, wantedIndices);
        if (!boundsByIndex) return [expandBounds(frameBounds, boxOptions.margins)];

        /* 各行の上端は、外形の上端から前の行までの行送りを引いた位置。行送りは行頭の文字のもの
           each line's top is the frame top minus the leading of the lines above, read from each line's first character */
        var lineTops = [frameBounds[1]];
        var leadingTotal = 0;
        for (i = 1; i < textLines.length; i++) {
            var lineHeadIndex = Math.min(textLines[i].start, textContents.length - 1);
            leadingTotal += readCellMetrics(textFrame.characters[lineHeadIndex]).leading;
            lineTops.push(frameBounds[1] - leadingTotal);
        }
        /* 外形の高さから行送りの合計を引いた残りが1行の高さ / the frame height minus the leadings is one line's height */
        var lineHeight = (frameBounds[1] - frameBounds[3]) - leadingTotal;

        var alignedRight = -Infinity;
        for (i = 0; i < textLines.length; i++) {
            if (textLines[i].firstGlyph >= 0) alignedRight = Math.max(alignedRight, boundsByIndex[textLines[i].lastGlyph][2]);
        }
        var lineBoxes = [];
        for (i = 0; i < textLines.length; i++) {
            if (textLines[i].firstGlyph < 0) continue;
            var boxRight = boxOptions.alignRight ? alignedRight : boundsByIndex[textLines[i].lastGlyph][2];
            lineBoxes.push(expandBounds([boundsByIndex[textLines[i].firstGlyph][0], lineTops[i], boxRight, lineTops[i] - lineHeight], boxOptions.margins));
        }
        return lineBoxes;
    }

    /**
     * 名前が続く角の記号の横線を、囲み罫の左辺で止めるよう短くする量を決める。
     * つなげるときは、名前の位置を測れた記号の横線を囲み罫の左辺（名前の左端 − マージン）で終える
     * @param {Object[]} symbols - findTreeSymbols() の戻り値。armTrim か armEnd を足す
     * @param {string} textContents - 書き換える前の contents
     * @param {Object} symbolIndices - 記号の文字の番号をキーにしたオブジェクト
     * @param {number} marginPt - 囲み罫の左右のマージン（pt）
     * @param {boolean} attachArms - 横線を囲み罫につなげるなら true
     * @returns {void}
     */
    function trimArmsForBoxes(symbols, textContents, symbolIndices, marginPt, attachArms) {
        for (var i = 0; i < symbols.length; i++) {
            if (symbols[i].kind === "vertical") continue;
            /* 同じ行の後ろに名前（字形）があるときだけ、その行の囲みに届く / only when a name follows on the same line */
            for (var k = symbols[i].index + symbols[i].length; k < textContents.length; k++) {
                var nextChar = textContents.charAt(k);
                if (isLineBreakChar(nextChar)) break;
                if (symbolIndices[k] || isBlankChar(nextChar)) continue;
                if (attachArms && symbols[i].nameLeft !== undefined) symbols[i].armEnd = symbols[i].nameLeft - marginPt;
                else symbols[i].armTrim = marginPt;
                break;
            }
        }
    }

    /**
     * パスを塗りなし・罫線の線にする
     * @param {PathItem} pathItem - 対象のパス
     * @param {Object} lineStyle - strokeWidth（pt）/ strokeColor / strokeCap / strokeJoin
     * @returns {void}
     */
    function applyLineStyle(pathItem, lineStyle) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeWidth = lineStyle.strokeWidth;
        pathItem.strokeColor = lineStyle.strokeColor;
        pathItem.strokeCap = lineStyle.strokeCap;
        pathItem.strokeJoin = lineStyle.strokeJoin;
    }

    /**
     * 罫線の色を作る
     * @returns {CMYKColor} 罫線の色
     */
    function createLineColor() {
        var lineColor = new CMYKColor();
        lineColor.cyan = LINE_COLOR_CMYK[0];
        lineColor.magenta = LINE_COLOR_CMYK[1];
        lineColor.yellow = LINE_COLOR_CMYK[2];
        lineColor.black = LINE_COLOR_CMYK[3];
        return lineColor;
    }

    // =========================================
    // 色 / Color
    // =========================================

    /**
     * RGB の値から RGBColor を作る
     * @param {number} red - 0〜255
     * @param {number} green - 0〜255
     * @param {number} blue - 0〜255
     * @returns {RGBColor} 色
     */
    function makeRGBColor(red, green, blue) {
        var rgbColor = new RGBColor();
        rgbColor.red = red;
        rgbColor.green = green;
        rgbColor.blue = blue;
        return rgbColor;
    }

    /**
     * HEX（#RRGGBB、# は省略可）を RGBColor にする
     * @param {string} hexText - HEX
     * @returns {RGBColor|null} 色。形式が違えば null
     */
    function parseHexColor(hexText) {
        var hexMatch = String(hexText || "").replace(/^\s+|\s+$/g, "").match(/^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/);
        if (!hexMatch) return null;
        return makeRGBColor(parseInt(hexMatch[1], 16), parseInt(hexMatch[2], 16), parseInt(hexMatch[3], 16));
    }

    /**
     * カラーピッカーが返した色を RGBColor にする（CMYK・グレーは RGB に変換）
     * @param {Color} pickedColor - 色
     * @returns {RGBColor|null} 色。変換できない種類なら null
     */
    function toRGBColor(pickedColor) {
        if (pickedColor.typename === "RGBColor") return pickedColor;
        var rgbValues;
        if (pickedColor.typename === "CMYKColor") {
            rgbValues = app.convertSampleColor(ImageColorSpace.CMYK,
                [pickedColor.cyan, pickedColor.magenta, pickedColor.yellow, pickedColor.black],
                ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
        } else if (pickedColor.typename === "GrayColor") {
            rgbValues = app.convertSampleColor(ImageColorSpace.GrayScale, [pickedColor.gray],
                ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
        } else {
            return null;
        }
        return makeRGBColor(rgbValues[0], rgbValues[1], rgbValues[2]);
    }

    /**
     * RGBColor を #RRGGBB にする
     * @param {RGBColor} rgbColor - 色
     * @returns {string} HEX
     */
    function toHexColor(rgbColor) {
        var channels = [rgbColor.red, rgbColor.green, rgbColor.blue];
        var hexText = "#";
        for (var i = 0; i < channels.length; i++) {
            var channelHex = Math.round(channels[i]).toString(16);
            hexText += (channelHex.length === 1 ? "0" : "") + channelHex;
        }
        return hexText;
    }

    // =========================================
    // 連続したスペース / Space runs
    // =========================================

    /**
     * 文字が罫線素片（U+2500〜U+257F）かどうかを返す
     * @param {string} oneChar - 1文字
     * @returns {boolean} 罫線素片なら true
     */
    function isBoxDrawingChar(oneChar) {
        var charCode = oneChar.charCodeAt(0);
        return charCode >= 0x2500 && charCode <= 0x257F;
    }

    /**
     * タブに変える「2つ以上続く半角スペース」の範囲を拾う。
     * 行頭の字下げと、ツリー記号の前の空き（「│  └─」の字下げ）は対象にしない。行末の空白も残す
     * @param {string} textContents - テキストフレームの contents
     * @returns {Object[]} { start, end }（end は含まない）の配列（前から順）
     */
    function findSpaceRuns(textContents) {
        var spaceRuns = [];
        var runPattern = /[ \u00A0]{2,}/g;
        var match;
        while ((match = runPattern.exec(textContents)) !== null) {
            var runStart = match.index;
            var runEnd = runStart + match[0].length;
            var charBefore = textContents.charAt(runStart - 1);
            var charAfter = textContents.charAt(runEnd);
            if (runStart === 0 || isLineBreakChar(charBefore) || isBlankChar(charBefore)) continue;
            if (charAfter === "" || isLineBreakChar(charAfter) || isBoxDrawingChar(charAfter)) continue;
            spaceRuns.push({ start: runStart, end: runEnd });
        }
        return spaceRuns;
    }

    // =========================================
    // 記号の削除と位置合わせ / Symbol removal and position fix
    // =========================================
    var POSITION_TOLERANCE_PT   = 0.01; /* このずれ未満なら合わせ直さない（pt） / offsets below this are left alone */
    var POSITION_FIX_MAX_PASSES = 3;    /* 測り直して合わせる最大回数 / max measure-and-fix passes */

    /**
     * 文字が改行かどうかを返す
     * @param {string} oneChar - 1文字（範囲外は空文字）
     * @returns {boolean} 改行なら true
     */
    function isLineBreakChar(oneChar) {
        return oneChar === "\r" || oneChar === "\n" || oneChar === "\u0003";
    }

    /**
     * 文字ごとの段落の番号を返す（段落の区切りは \r）
     * @param {string} textContents - テキストフレームの contents
     * @returns {number[]} 文字の番号 → 段落の番号
     */
    function listParagraphNumbers(textContents) {
        var paragraphNumbers = [];
        var paragraphNumber = 0;
        for (var i = 0; i < textContents.length; i++) {
            paragraphNumbers.push(paragraphNumber);
            if (textContents.charAt(i) === "\r") paragraphNumber++;
        }
        return paragraphNumbers;
    }

    /**
     * ツリー記号を含む行頭の字下げ（行頭から続く空白と記号）の範囲を、行ごとに拾う
     * @param {string} textContents - テキストフレームの contents
     * @param {Object} symbolIndices - ツリー記号の文字の番号をキーにしたオブジェクト
     * @returns {Object[]} { start, end（字下げのあとの最初の文字）, reachesLineEnd } の配列（前から順）
     */
    function findIndentRegions(textContents, symbolIndices) {
        var indentRegions = [];
        var lineStart = 0;
        for (var i = 0; i <= textContents.length; i++) {
            if (i < textContents.length && !isLineBreakChar(textContents.charAt(i))) continue;
            var indentEnd = lineStart;
            var hasSymbol = false;
            for (; indentEnd < i; indentEnd++) {
                if (symbolIndices[indentEnd]) hasSymbol = true;
                else if (!isBlankChar(textContents.charAt(indentEnd))) break;
            }
            if (hasSymbol) indentRegions.push({ start: lineStart, end: indentEnd, reachesLineEnd: indentEnd === i });
            lineStart = i + 1;
        }
        return indentRegions;
    }

    /**
     * 字下げの外にある記号ごとに、削除のしかたと、位置を保つ基準の文字を決める。
     * 記号の直前に残る文字があれば記号を丸ごと消し、その文字のカーニングで幅を埋める。
     * 行頭や、記号が隙間なく続くときは1文字だけ LINE_START_PLACEHOLDER として残し、その文字のカーニングで幅を埋める
     * @param {string} textContents - テキストフレームの contents
     * @param {Object[]} kernedSymbols - カーニングで位置を保つ記号（前から順）
     * @param {Object} symbolIndices - すべての記号の文字の番号をキーにしたオブジェクト
     * @returns {void} 各記号に keepPlaceholder / anchorIndex（位置を保つ字形のある文字。無ければ -1）/ lineNumber を足す
     */
    function planKernedSymbols(textContents, kernedSymbols, symbolIndices) {
        var lineNumber = 0;
        var scanIndex = 0;
        for (var i = 0; i < kernedSymbols.length; i++) {
            var symbol = kernedSymbols[i];
            for (; scanIndex < symbol.index; scanIndex++) {
                if (isLineBreakChar(textContents.charAt(scanIndex))) lineNumber++;
            }
            symbol.lineNumber = lineNumber;
            /* 前の文字が行頭・改行・削除する記号なら、カーニングを入れる文字が無いので1文字残す
               keep one character when nothing usable precedes the symbol (line start, line break or another symbol) */
            symbol.keepPlaceholder = (symbol.index === 0 || isLineBreakChar(textContents.charAt(symbol.index - 1)) || symbolIndices[symbol.index - 1] === true);
            /* 同じ行で、記号の後ろにある最初の字形（空白・ほかの記号は飛ばす） / first glyph after the symbol on the same line */
            symbol.anchorIndex = -1;
            for (var k = symbol.index + symbol.length; k < textContents.length; k++) {
                var nextChar = textContents.charAt(k);
                if (isLineBreakChar(nextChar)) break;
                if (symbolIndices[k] || isBlankChar(nextChar)) continue;
                symbol.anchorIndex = k;
                break;
            }
        }
    }

    /**
     * 1つのテキストフレームで行う書き換えを決める（字下げ → タブ、連続したスペース → タブ、残りの記号 → 削除）
     * @param {string} textContents - テキストフレームの contents
     * @param {Object[]} symbols - findTreeSymbols() の戻り値
     * @param {Object} conversionSettings - convertIndents / convertSpaceRuns
     * @returns {{indentRegions: Object[], spaceRuns: Object[], kernedSymbols: Object[], symbolIndices: Object}} 書き換えの計画
     */
    function planFrameEdits(textContents, symbols, conversionSettings) {
        var symbolIndices = {};
        var i, k;
        for (i = 0; i < symbols.length; i++) {
            for (k = 0; k < symbols[i].length; k++) symbolIndices[symbols[i].index + k] = true;
        }
        var indentRegions = conversionSettings.convertIndents ? findIndentRegions(textContents, symbolIndices) : [];
        var allSpaceRuns = conversionSettings.convertSpaceRuns ? findSpaceRuns(textContents) : [];

        /**
         * 番号が字下げの中にあれば、その字下げを返す
         * @param {number} charIndex - 文字の番号
         * @returns {Object|null} 字下げ（無ければ null）
         */
        function findIndentRegionAt(charIndex) {
            for (var j = 0; j < indentRegions.length; j++) {
                if (charIndex >= indentRegions[j].start && charIndex < indentRegions[j].end) return indentRegions[j];
            }
            return null;
        }

        /* 字下げの中の連続したスペースは字下げのタブにまとめ、その行の続きは説明の列とみなす
           a space run inside an indent is absorbed by the indent tab, and the rest of the line belongs to the description column */
        var spaceRuns = [];
        for (i = 0; i < allSpaceRuns.length; i++) {
            var containingRegion = findIndentRegionAt(allSpaceRuns[i].start);
            if (containingRegion) containingRegion.isDescription = true;
            else spaceRuns.push(allSpaceRuns[i]);
        }
        var kernedSymbols = [];
        for (i = 0; i < symbols.length; i++) {
            if (!findIndentRegionAt(symbols[i].index)) kernedSymbols.push(symbols[i]);
        }
        planKernedSymbols(textContents, kernedSymbols, symbolIndices);
        return { indentRegions: indentRegions, spaceRuns: spaceRuns, kernedSymbols: kernedSymbols, symbolIndices: symbolIndices };
    }

    /**
     * 書き換え後の文字の番号を返す
     * @param {Object[]} textEdits - { start, end, replacement } の配列
     * @param {number} originalIndex - 書き換え前の番号
     * @returns {number} 書き換え後の番号
     */
    function shiftIndexAfterEdits(textEdits, originalIndex) {
        var shiftedIndex = originalIndex;
        for (var i = 0; i < textEdits.length; i++) {
            if (textEdits[i].end <= originalIndex) shiftedIndex -= (textEdits[i].end - textEdits[i].start) - textEdits[i].replacement.length;
        }
        return shiftedIndex;
    }

    /**
     * 範囲の書き換えを後ろから順に行う（前の文字の番号をずらさない。書式は先頭の文字のものが残る）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} textEdits - { start, end, replacement } の配列（重ならないこと）
     * @returns {void}
     */
    function applyTextEdits(textFrame, textEdits) {
        var sortedEdits = textEdits.slice(0);
        /* 比較関数つきの sort() は遅く並びも狂うので、番号の大きい順に選んで並べる / avoid sort() with a comparator */
        var orderedEdits = [];
        while (sortedEdits.length > 0) {
            var lastPosition = 0;
            for (var i = 1; i < sortedEdits.length; i++) {
                if (sortedEdits[i].start > sortedEdits[lastPosition].start) lastPosition = i;
            }
            orderedEdits.push(sortedEdits.splice(lastPosition, 1)[0]);
        }
        for (i = 0; i < orderedEdits.length; i++) {
            var textEdit = orderedEdits[i];
            var frameChars = textFrame.characters;
            for (var k = textEdit.end - 1; k > textEdit.start; k--) frameChars[k].remove();
            if (textEdit.replacement === "") frameChars[textEdit.start].remove();
            else frameChars[textEdit.start].contents = textEdit.replacement;
        }
    }

    /**
     * 段落のタブストップを、左揃えの位置の並びで置き換える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object} paragraphStops - charIndex（段落内の文字の番号）/ stops（{ position } の配列、左から順）
     * @returns {void}
     */
    function applyParagraphTabStops(textFrame, paragraphStops) {
        var tabStops = [];
        for (var i = 0; i < paragraphStops.stops.length; i++) {
            var tabStop = new TabStopInfo();
            tabStop.alignment = TabStopAlignment.Left;
            tabStop.position = paragraphStops.stops[i].position;
            tabStops.push(tabStop);
        }
        textFrame.characters[paragraphStops.charIndex].paragraphs[0].paragraphAttributes.tabStops = tabStops;
    }

    /**
     * 段落ごとのタブストップを決める。名前は深さごとの位置に、説明は1列にそろえた位置に置く。
     * ダイアログで位置を指定した深さ（tabStopOverrides）は、その位置に固定する
     * @param {Object} framePlan - planFrameEdits() の戻り値（assignTabTargets() 済み）
     * @param {number[]} paragraphNumbers - listParagraphNumbers() の戻り値
     * @param {number} originX - タブ位置 0 に当たる横位置の見込み（pt）
     * @param {Object} tabStopOverrides - 深さのキー → タブストップの位置（pt）
     * @returns {Object[]} 段落ごとの { stops, charIndex（書き換え前）}。各範囲には tabCount を足す
     */
    function planTabStops(framePlan, paragraphNumbers, originX, tabStopOverrides) {
        var stopsByParagraph = {};
        var paragraphList = [];

        /**
         * 段落のタブストップに深さを足し（同じ深さは1つにまとめる）、そのタブストップを返す
         * @param {Object} tabRange - 字下げかスペースの範囲（start / end / targetX / levelKey）
         * @returns {Object} タブストップ { levelKey, targetX, position, isFixed, anchorIndex }
         */
        function addStop(tabRange) {
            var paragraphNumber = paragraphNumbers[tabRange.start];
            var paragraphStops = stopsByParagraph[paragraphNumber];
            if (!paragraphStops) {
                paragraphStops = { stops: [], charIndex: tabRange.start };
                stopsByParagraph[paragraphNumber] = paragraphStops;
                paragraphList.push(paragraphStops);
            }
            for (var i = 0; i < paragraphStops.stops.length; i++) {
                if (paragraphStops.stops[i].levelKey === tabRange.levelKey) return paragraphStops.stops[i];
            }
            var isFixed = tabStopOverrides.hasOwnProperty(tabRange.levelKey);
            var newStop = {
                levelKey: tabRange.levelKey,
                targetX: tabRange.targetX,
                position: isFixed ? tabStopOverrides[tabRange.levelKey] : tabRange.targetX - originX,
                isFixed: isFixed,
                anchorIndex: tabRange.end
            };
            /* 左から順に差し込む / keep the stops in left-to-right order */
            var insertAt = 0;
            while (insertAt < paragraphStops.stops.length && paragraphStops.stops[insertAt].position < newStop.position) insertAt++;
            paragraphStops.stops.splice(insertAt, 0, newStop);
            return newStop;
        }

        var i;
        var regionsWithStop = [];
        for (i = 0; i < framePlan.indentRegions.length; i++) {
            var indentRegion = framePlan.indentRegions[i];
            if (indentRegion.targetX === undefined) continue;
            indentRegion.tabStop = addStop(indentRegion);
            regionsWithStop.push(indentRegion);
        }
        for (i = 0; i < framePlan.spaceRuns.length; i++) {
            if (framePlan.spaceRuns[i].targetX !== undefined) addStop(framePlan.spaceRuns[i]);
        }
        /* 字下げのタブの数は、段落のタブストップのうち何番目に合わせるか / an indent needs as many tabs as its stop's rank */
        for (i = 0; i < regionsWithStop.length; i++) {
            var regionStops = stopsByParagraph[paragraphNumbers[regionsWithStop[i].start]].stops;
            for (var k = 0; k < regionStops.length; k++) {
                if (regionStops[k] === regionsWithStop[i].tabStop) regionsWithStop[i].tabCount = k + 1;
            }
        }
        return paragraphList;
    }

    /**
     * タブストップを置き、固定していないものは字形を測り直してずれた分だけ直す（POSITION_FIX_MAX_PASSES 回まで）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} paragraphList - planTabStops() の戻り値（charIndex・anchorIndex は書き換え後の番号に直したもの）
     * @returns {void}
     */
    function placeTabStops(textFrame, paragraphList) {
        var wantedIndices = {};
        var hasAdjustableStop = false;
        var i, k;
        for (i = 0; i < paragraphList.length; i++) {
            applyParagraphTabStops(textFrame, paragraphList[i]);
            for (k = 0; k < paragraphList[i].stops.length; k++) {
                if (paragraphList[i].stops[k].isFixed) continue;
                wantedIndices[paragraphList[i].stops[k].anchorIndex] = true;
                hasAdjustableStop = true;
            }
        }
        for (var pass = 0; pass < POSITION_FIX_MAX_PASSES && hasAdjustableStop; pass++) {
            var boundsByIndex = measureGlyphBounds(textFrame, wantedIndices);
            if (!boundsByIndex) return;
            var hasOffset = false;
            for (i = 0; i < paragraphList.length; i++) {
                var isParagraphChanged = false;
                for (k = 0; k < paragraphList[i].stops.length; k++) {
                    var tabStop = paragraphList[i].stops[k];
                    var anchorBounds = boundsByIndex[tabStop.anchorIndex];
                    if (tabStop.isFixed || !anchorBounds) continue;
                    var offset = tabStop.targetX - anchorBounds[0];
                    if (Math.abs(offset) < POSITION_TOLERANCE_PT) continue;
                    tabStop.position += offset;
                    isParagraphChanged = true;
                }
                if (isParagraphChanged) {
                    applyParagraphTabStops(textFrame, paragraphList[i]);
                    hasOffset = true;
                }
            }
            if (!hasOffset) return;
        }
    }

    /**
     * 文字の手動カーニングを読む。手動の値が無い位置では例外になるので、そのときは 0
     * @param {TextRange} oneChar - 1文字の範囲
     * @returns {number} カーニング（1/1000em）
     */
    function readManualKerning(oneChar) {
        try {
            return oneChar.kerning;
        } catch (e) {
            return 0;
        }
    }

    /**
     * 字下げの外で記号を消して詰まった位置を、カーニングで元の位置に戻す。
     * アウトラインで測り直してずれを埋める、を POSITION_FIX_MAX_PASSES 回まで繰り返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} kernedSymbols - planKernedSymbols() 済みの記号（anchorLeft に削除前の左端を入れたもの）
     * @param {Object[]} textEdits - 行った書き換え（番号の読み替えに使う）
     * @returns {void}
     */
    function restoreKernedPositions(textFrame, kernedSymbols, textEdits) {
        var anchoredSymbols = [];
        var wantedIndices = {};
        for (var i = 0; i < kernedSymbols.length; i++) {
            var symbol = kernedSymbols[i];
            if (symbol.anchorLeft === undefined) continue;
            /* 残した1文字か、記号の直前の文字にカーニングを入れる / kern the kept character or the one before the symbol */
            symbol.kerningCharIndex = symbol.keepPlaceholder
                ? shiftIndexAfterEdits(textEdits, symbol.index)
                : shiftIndexAfterEdits(textEdits, symbol.index - 1);
            symbol.shiftedAnchorIndex = shiftIndexAfterEdits(textEdits, symbol.anchorIndex);
            wantedIndices[symbol.shiftedAnchorIndex] = true;
            anchoredSymbols.push(symbol);
        }

        for (var pass = 0; pass < POSITION_FIX_MAX_PASSES && anchoredSymbols.length > 0; pass++) {
            var boundsByIndex = measureGlyphBounds(textFrame, wantedIndices);
            if (!boundsByIndex) return;
            var hasOffset = false;
            var previousLine = -1;
            var previousOffset = 0;
            for (i = 0; i < anchoredSymbols.length; i++) {
                var anchoredSymbol = anchoredSymbols[i];
                var anchorBounds = boundsByIndex[anchoredSymbol.shiftedAnchorIndex];
                if (!anchorBounds) continue;
                /* ずれは同じ行の前の記号の分を含むので、差分だけこの記号で埋める
                   the offset includes the earlier symbols on the line, so fill only the difference here */
                var offset = anchoredSymbol.anchorLeft - anchorBounds[0];
                var ownOffset = (anchoredSymbol.lineNumber === previousLine) ? offset - previousOffset : offset;
                previousLine = anchoredSymbol.lineNumber;
                previousOffset = offset;
                if (Math.abs(ownOffset) < POSITION_TOLERANCE_PT) continue;
                hasOffset = true;
                var kerningChar = textFrame.characters[anchoredSymbol.kerningCharIndex];
                var charSize = kerningChar.characterAttributes.size;
                kerningChar.kerning = Math.round(readManualKerning(kerningChar) + ownOffset / charSize * 1000);
            }
            if (!hasOffset) return;
        }
    }

    // =========================================
    // 変換 / Conversion
    // =========================================
    var DESCRIPTION_LEVEL_KEY = "description"; /* 説明の列のタブストップのキー / key of the description column's tab stop */

    /**
     * テキストのタブ位置 0 に当たる横位置の見込みを返す（ずれは測り直して直す）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} 横位置（pt）
     */
    function estimateTabOriginX(textFrame) {
        return (textFrame.kind === TextType.POINTTEXT) ? textFrame.anchor[0] : textFrame.geometricBounds[0];
    }

    /**
     * 測る字形の番号をまとめる（記号とその横線、位置を保つ字形、字下げ・スペースのあとの字形）
     * @param {Object[]} symbols - findTreeSymbols() の戻り値
     * @param {Object} framePlan - planFrameEdits() の戻り値
     * @returns {Object} 文字の番号をキーにしたオブジェクト
     */
    function collectWantedIndices(symbols, framePlan) {
        var wantedIndices = {};
        var i;
        for (i = 0; i < symbols.length; i++) {
            wantedIndices[symbols[i].index] = true;
            /* 最初と最後の横線（1本なら同じ文字） / the first and last bars (the same character when there is one) */
            if (symbols[i].length >= 2) {
                wantedIndices[symbols[i].index + 1] = true;
                wantedIndices[symbols[i].index + symbols[i].length - 1] = true;
            }
        }
        for (i = 0; i < framePlan.kernedSymbols.length; i++) {
            if (framePlan.kernedSymbols[i].anchorIndex >= 0) wantedIndices[framePlan.kernedSymbols[i].anchorIndex] = true;
        }
        for (i = 0; i < framePlan.indentRegions.length; i++) {
            if (!framePlan.indentRegions[i].reachesLineEnd) wantedIndices[framePlan.indentRegions[i].end] = true;
        }
        for (i = 0; i < framePlan.spaceRuns.length; i++) wantedIndices[framePlan.spaceRuns[i].end] = true;
        return wantedIndices;
    }

    /**
     * 字下げ・スペースのあとの字形の位置から、タブで合わせる位置を決める。
     * 説明（スペースのあと、または字下げの中にスペースがあった行の続き）は、いちばん右の説明の位置で1列にそろえる。
     * 名前はツリーの深さごとにまとめる（横線は、タブストップを置いたあとの名前の位置まで伸ばす）
     * @param {Object} framePlan - planFrameEdits() の戻り値
     * @param {Object} boundsByIndex - measureGlyphBounds() の戻り値
     * @param {string} textContents - テキストフレームの contents
     * @param {Object[]} symbols - findTreeSymbols() の戻り値
     * @returns {void} 各範囲に targetX / levelKey を、名前の字下げに originalX / cornerSymbol を足す
     */
    function assignTabTargets(framePlan, boundsByIndex, textContents, symbols) {
        var descriptionX = null;
        var i, glyphBounds;
        var descriptionRanges = [];
        for (i = 0; i < framePlan.spaceRuns.length; i++) descriptionRanges.push(framePlan.spaceRuns[i]);
        for (i = 0; i < framePlan.indentRegions.length; i++) {
            if (framePlan.indentRegions[i].isDescription) descriptionRanges.push(framePlan.indentRegions[i]);
        }
        for (i = 0; i < descriptionRanges.length; i++) {
            glyphBounds = boundsByIndex[descriptionRanges[i].end];
            if (glyphBounds && (descriptionX === null || glyphBounds[0] > descriptionX)) descriptionX = glyphBounds[0];
        }
        for (i = 0; i < descriptionRanges.length; i++) {
            descriptionRanges[i].levelKey = DESCRIPTION_LEVEL_KEY;
            if (descriptionX !== null) descriptionRanges[i].targetX = descriptionX;
        }
        /* 名前はツリーの深さごとに1つの位置へまとめる。全角と半角が混ざると、同じ深さでも元の位置が
           ばらつくため（「│  └─」と「   ├─」）。位置はその深さでいちばん右の名前に合わせ、あとで横線を伸ばしてつなぐ
           names are grouped by tree depth: mixed full/half widths scatter the original positions of the same depth.
           Each depth goes to its rightmost name; the arms are extended later to meet it */
        var levelX = {};
        var nameRegions = [];
        for (i = 0; i < framePlan.indentRegions.length; i++) {
            var indentRegion = framePlan.indentRegions[i];
            glyphBounds = boundsByIndex[indentRegion.end];
            if (indentRegion.isDescription || !glyphBounds) continue;
            indentRegion.originalX = glyphBounds[0];
            /* 深さは角の記号（├・└）が何桁目にあるかで決まる。角が無い行は名前の桁で代える
               depth comes from the column of the corner (├, └); rows without one use the name's column */
            var cornerSymbol = findLastCornerSymbol(symbols, indentRegion);
            indentRegion.cornerSymbol = cornerSymbol;
            indentRegion.levelKey = cornerSymbol
                ? "corner" + countColumns(textContents, indentRegion.start, cornerSymbol.index)
                : "name" + countColumns(textContents, indentRegion.start, indentRegion.end);
            if (levelX[indentRegion.levelKey] === undefined || indentRegion.originalX > levelX[indentRegion.levelKey]) {
                levelX[indentRegion.levelKey] = indentRegion.originalX;
            }
            nameRegions.push(indentRegion);
        }
        for (i = 0; i < nameRegions.length; i++) nameRegions[i].targetX = levelX[nameRegions[i].levelKey];
    }

    /**
     * 範囲の桁数を返す（ツリーの出力に合わせて1文字1桁、全角スペースだけ2桁）
     * @param {string} textContents - テキストフレームの contents
     * @param {number} startIndex - 範囲の先頭（含む）
     * @param {number} endIndex - 範囲の末尾（含まない）
     * @returns {number} 桁数
     */
    function countColumns(textContents, startIndex, endIndex) {
        var columnCount = 0;
        for (var i = startIndex; i < endIndex; i++) {
            columnCount += (textContents.charAt(i) === "\u3000") ? 2 : 1;
        }
        return columnCount;
    }

    /**
     * 字下げの中でいちばん後ろにある角の記号（名前へ横線を伸ばす記号）を返す
     * @param {Object[]} symbols - findTreeSymbols() の戻り値
     * @param {Object} indentRegion - findIndentRegions() の要素
     * @returns {Object|null} 記号（無ければ null）
     */
    function findLastCornerSymbol(symbols, indentRegion) {
        var lastCorner = null;
        for (var i = 0; i < symbols.length; i++) {
            var symbol = symbols[i];
            if (symbol.index >= indentRegion.start && symbol.index < indentRegion.end && symbol.kind !== "vertical") lastCorner = symbol;
        }
        return lastCorner;
    }

    /**
     * 計画から、範囲の書き換えの一覧を作る
     * @param {Object} framePlan - planFrameEdits() の戻り値（planTabStops() 済み）
     * @returns {Object[]} { start, end, replacement } の配列
     */
    function buildTextEdits(framePlan) {
        var textEdits = [];
        var i;
        for (i = 0; i < framePlan.indentRegions.length; i++) {
            var indentRegion = framePlan.indentRegions[i];
            /* 字形の位置を測れなかった字下げ・記号だけの行は、字下げを消すだけ / no target: just drop the indent */
            var tabCount = indentRegion.tabCount || 0;
            textEdits.push({ start: indentRegion.start, end: indentRegion.end, replacement: new Array(tabCount + 1).join("\t") });
        }
        for (i = 0; i < framePlan.spaceRuns.length; i++) {
            textEdits.push({ start: framePlan.spaceRuns[i].start, end: framePlan.spaceRuns[i].end, replacement: "\t" });
        }
        for (i = 0; i < framePlan.kernedSymbols.length; i++) {
            var symbol = framePlan.kernedSymbols[i];
            textEdits.push({ start: symbol.index, end: symbol.index + symbol.length, replacement: symbol.keepPlaceholder ? LINE_START_PLACEHOLDER : "" });
        }
        return textEdits;
    }

    /**
     * テキスト全体の行送りを固定値にする
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} leadingPt - 行送り（pt）
     * @returns {void}
     */
    function applyLeading(textFrame, leadingPt) {
        var frameAttributes = textFrame.textRange.characterAttributes;
        frameAttributes.autoLeading = false;
        frameAttributes.leading = leadingPt;
    }

    /**
     * タブストップを置いたあとの名前の位置を測り、角の記号の横線を名前まで伸ばす量を決める
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object} framePlan - planFrameEdits() の戻り値（assignTabTargets() 済み）
     * @param {Object[]} textEdits - 行った書き換え（番号の読み替えに使う）
     * @returns {void} 角の記号に armExtension と nameLeft（名前の左端 pt）を足す
     */
    function measureArmExtensions(textFrame, framePlan, textEdits) {
        var wantedIndices = {};
        var armRegions = [];
        for (var i = 0; i < framePlan.indentRegions.length; i++) {
            var indentRegion = framePlan.indentRegions[i];
            if (!indentRegion.cornerSymbol || indentRegion.originalX === undefined) continue;
            indentRegion.shiftedNameIndex = shiftIndexAfterEdits(textEdits, indentRegion.end);
            wantedIndices[indentRegion.shiftedNameIndex] = true;
            armRegions.push(indentRegion);
        }
        if (armRegions.length === 0) return;
        var boundsByIndex = measureGlyphBounds(textFrame, wantedIndices);
        if (!boundsByIndex) return;
        for (i = 0; i < armRegions.length; i++) {
            var nameBounds = boundsByIndex[armRegions[i].shiftedNameIndex];
            if (!nameBounds) continue;
            armRegions[i].cornerSymbol.armExtension = nameBounds[0] - armRegions[i].originalX;
            armRegions[i].cornerSymbol.nameLeft = nameBounds[0];
        }
    }

    /**
     * 段落のタブストップの位置を、深さのキーごとに控える（最初に見つけた値）
     * @param {Object[]} paragraphList - placeTabStops() 済みの段落
     * @param {Object} levelPositions - 深さのキー → タブストップの位置（pt）。ここに足す
     * @returns {void}
     */
    function recordLevelPositions(paragraphList, levelPositions) {
        for (var i = 0; i < paragraphList.length; i++) {
            for (var k = 0; k < paragraphList[i].stops.length; k++) {
                var tabStop = paragraphList[i].stops[k];
                if (!levelPositions.hasOwnProperty(tabStop.levelKey)) levelPositions[tabStop.levelKey] = tabStop.position;
            }
        }
    }

    /**
     * 1つのテキストフレームのツリー記号を罫線のパスに変換し、記号を削除する。
     * 字下げと連続したスペースはタブにしてタブストップで位置を決め、そのほかの記号の跡はカーニングで元の位置へ戻す。
     * 横線は、名前が最後に落ち着いた位置まで伸ばす
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Function} getLineGroup - 罫線を入れるグループを返す関数（最初の罫線を描くときに作る）
     * @param {Object} lineStyle - strokeWidth（pt）/ strokeColor / strokeCap / strokeJoin
     * @param {Object} conversionSettings - leadingPt / convertIndents / convertSpaceRuns / drawBox / boxMarginXPt / boxMarginYPt / alignBoxRight / attachArmsToBox /
     *     tabStopOverrides / stemOverrides
     * @param {Object} levelPositions - 深さのキー → 置いたタブストップの位置（pt）。ここに足す
     * @param {Object} stemPositions - 列のキー → そろえた縦罫の位置（pt）。ここに足す
     * @returns {{symbolCount: number, indentCount: number, spaceRunCount: number, boxPaths: PathItem[]}|null} 変換した数と描いた囲み罫（位置を測れなければ null）
     */
    function convertTextFrame(textFrame, getLineGroup, lineStyle, conversionSettings, levelPositions, stemPositions) {
        /* 行送りは記号の位置を動かすので、測る前に当てる / leading moves the symbols, so apply it before measuring */
        if (conversionSettings.leadingPt !== null) applyLeading(textFrame, conversionSettings.leadingPt);

        var textContents = textFrame.contents;
        var symbols = findTreeSymbols(textContents);
        var framePlan = planFrameEdits(textContents, symbols, conversionSettings);
        var frameCounts = { symbolCount: symbols.length, indentCount: framePlan.indentRegions.length, spaceRunCount: framePlan.spaceRuns.length, boxPaths: [] };
        if (symbols.length === 0 && framePlan.spaceRuns.length === 0) {
            drawBoxes(textFrame, getLineGroup, lineStyle, conversionSettings, frameCounts.boxPaths);
            return frameCounts;
        }

        var boundsByIndex = measureGlyphBounds(textFrame, collectWantedIndices(symbols, framePlan));
        if (!boundsByIndex) return null;

        rememberGlyphInfo(textFrame, textContents, symbols, framePlan, boundsByIndex);
        assignTabTargets(framePlan, boundsByIndex, textContents, symbols);
        var originX = estimateTabOriginX(textFrame);
        var paragraphList = planTabStops(framePlan, listParagraphNumbers(textContents), originX, conversionSettings.tabStopOverrides);

        var textEdits = buildTextEdits(framePlan);
        applyTextEdits(textFrame, textEdits);

        /* 番号を書き換え後に読み替えてから、タブストップ → カーニングの順に合わせる / remap indices, then fix tab stops and kerning */
        remapParagraphStops(paragraphList, textEdits);
        placeTabStops(textFrame, paragraphList);
        restoreKernedPositions(textFrame, framePlan.kernedSymbols, textEdits);
        recordLevelPositions(paragraphList, levelPositions);

        measureArmExtensions(textFrame, framePlan, textEdits);
        if (conversionSettings.drawBox) trimArmsForBoxes(symbols, textContents, framePlan.symbolIndices, conversionSettings.boxMarginXPt, conversionSettings.attachArmsToBox);
        /* 囲み罫は変換したあとの文字の位置で決め、縦線をつなぐ先にも使う。囲み罫が無いときも、行の範囲（マージンなし）を
           つなぐ先にして、いちばん上の縦線を最初の項目まで伸ばす
           boxes follow the converted text and also anchor the stems; without boxes the bare line bounds anchor them,
           so the topmost stem still reaches the first item */
        var lineBoxes = conversionSettings.drawBox
            ? drawBoxes(textFrame, getLineGroup, lineStyle, conversionSettings, frameCounts.boxPaths)
            : measureLineBoxes(textFrame, { margins: { x: 0, y: 0 }, alignRight: true });
        drawTreeLines(symbols, boundsByIndex, textContents, {
            originX: originX,
            stemOverrides: conversionSettings.stemOverrides,
            stemPositions: stemPositions,
            lineBoxes: lineBoxes,
            getLineGroup: getLineGroup,
            lineStyle: lineStyle
        });
        return frameCounts;
    }

    /**
     * ［囲み罫］がオンなら、変換したあとのテキストの各行を囲む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Function} getLineGroup - 罫線を入れるグループを返す関数
     * @param {Object} lineStyle - strokeWidth（pt）/ strokeColor / strokeCap / strokeJoin
     * @param {Object} conversionSettings - drawBox / boxMarginXPt / boxMarginYPt / alignBoxRight
     * @param {PathItem[]} boxPaths - 描いた囲み罫のパス。ここに足す（上の行から順）
     * @returns {number[][]} 描いた囲み罫の範囲（オフなら空）
     */
    function drawBoxes(textFrame, getLineGroup, lineStyle, conversionSettings, boxPaths) {
        if (!conversionSettings.drawBox) return [];
        var lineBoxes = measureLineBoxes(textFrame, {
            margins: { x: conversionSettings.boxMarginXPt, y: conversionSettings.boxMarginYPt },
            alignRight: conversionSettings.alignBoxRight
        });
        for (var i = 0; i < lineBoxes.length; i++) boxPaths.push(addBoxPath(getLineGroup(), lineBoxes[i], lineStyle));
        return lineBoxes;
    }

    /**
     * 書き換える前に、罫線の形と位置合わせに要る情報を控える（1文字の幅と行送り・位置を保つ字形の左端）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} textContents - 書き換える前の contents
     * @param {Object[]} symbols - findTreeSymbols() の戻り値。cellMetrics を足す
     * @param {Object} framePlan - planFrameEdits() の戻り値。カーニングで保つ記号に anchorLeft を足す
     * @param {Object} boundsByIndex - measureGlyphBounds() の戻り値
     * @returns {void}
     */
    function rememberGlyphInfo(textFrame, textContents, symbols, framePlan, boundsByIndex) {
        var i;
        for (i = 0; i < symbols.length; i++) {
            symbols[i].cellMetrics = readCellMetrics(textFrame.characters[symbols[i].index]);
        }
        for (i = 0; i < framePlan.kernedSymbols.length; i++) {
            var anchorBounds = boundsByIndex[framePlan.kernedSymbols[i].anchorIndex];
            if (anchorBounds) framePlan.kernedSymbols[i].anchorLeft = anchorBounds[0];
        }
    }

    /**
     * 段落のタブストップが指す文字の番号を、書き換え後の番号に読み替える
     * @param {Object[]} paragraphList - planTabStops() の戻り値（その場で書き換える）
     * @param {Object[]} textEdits - 行った書き換え
     * @returns {void}
     */
    function remapParagraphStops(paragraphList, textEdits) {
        for (var i = 0; i < paragraphList.length; i++) {
            paragraphList[i].charIndex = shiftIndexAfterEdits(textEdits, paragraphList[i].charIndex);
            for (var k = 0; k < paragraphList[i].stops.length; k++) {
                paragraphList[i].stops[k].anchorIndex = shiftIndexAfterEdits(textEdits, paragraphList[i].stops[k].anchorIndex);
            }
        }
    }

    /**
     * ツリー記号の罫線を描く。縦線は列ごとにそろえ、上が途切れた縦線は前のレベルの横線か囲み罫までつなぐ
     * @param {Object[]} symbols - rememberGlyphInfo() 済みの記号
     * @param {Object} boundsByIndex - 書き換える前の字形の境界
     * @param {string} textContents - 書き換える前の contents
     * @param {Object} drawOptions - originX / stemOverrides / stemPositions（ここに足す）/ lineBoxes / getLineGroup / lineStyle
     * @returns {void}
     */
    function drawTreeLines(symbols, boundsByIndex, textContents, drawOptions) {
        var symbolShapes = [];
        for (var i = 0; i < symbols.length; i++) {
            symbolShapes.push({
                kind: symbols[i].kind,
                lines: buildSymbolLines(symbols[i], boundsByIndex, symbols[i].cellMetrics),
                stemKey: getStemKey(textContents, symbols[i].index)
            });
        }
        alignStemsByColumn(symbolShapes, drawOptions.originX, drawOptions.stemOverrides, drawOptions.stemPositions);
        connectStemsToParentArms(symbolShapes, drawOptions.lineBoxes);
        for (i = 0; i < symbolShapes.length; i++) {
            for (var j = 0; j < symbolShapes[i].lines.length; j++) addLinePath(drawOptions.getLineGroup(), symbolShapes[i].lines[j], drawOptions.lineStyle);
        }
    }

    /**
     * テキストフレームをすべて変換する
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {Object} conversionSettings - strokeWidthPt / strokeColor / roundCap / leadingPt / convertIndents / convertSpaceRuns /
     *     drawBox / boxMarginXPt / boxMarginYPt / alignBoxRight / attachArmsToBox / tabStopOverrides / stemOverrides
     * @param {PageItem[]} [createdItems] - 罫線のグループを作ったらすぐここに足す（途中で例外になっても片付けられるように）
     * @returns {Object} symbolCount / indentCount / spaceRunCount / failedFrameCount（位置を測れなかったテキスト）/
     *     levelPositions（深さのキー → タブストップの位置 pt）/ stemPositions（列のキー → 縦罫の位置 pt）/
     *     lineGroup（作った罫線のグループ。無ければ null）/ boxedFrames（囲み罫を描いたテキストごとの { textFrame, boxPaths }）
     */
    function convertTextFrames(doc, textFrames, conversionSettings, createdItems) {
        /* 角の形状は線端に連動させる（丸型 → ラウンド、なし → マイター） / the join follows the cap */
        var lineStyle = {
            strokeWidth: conversionSettings.strokeWidthPt,
            strokeColor: conversionSettings.strokeColor || createLineColor(),
            strokeCap: conversionSettings.roundCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP,
            strokeJoin: conversionSettings.roundCap ? StrokeJoin.ROUNDENDJOIN : StrokeJoin.MITERENDJOIN
        };
        var conversionResult = { symbolCount: 0, indentCount: 0, spaceRunCount: 0, failedFrameCount: 0, levelPositions: {}, stemPositions: {}, lineGroup: null, boxedFrames: [] };

        /**
         * 罫線のグループを返す（最初に呼ばれたときに作る）
         * @returns {GroupItem} 罫線のグループ
         */
        function getLineGroup() {
            if (!conversionResult.lineGroup) {
                conversionResult.lineGroup = doc.groupItems.add();
                conversionResult.lineGroup.name = getLabel("fallbackName.lineGroup");
                if (createdItems) createdItems.push(conversionResult.lineGroup);
            }
            return conversionResult.lineGroup;
        }

        for (var i = 0; i < textFrames.length; i++) {
            var frameCounts = convertTextFrame(textFrames[i], getLineGroup, lineStyle, conversionSettings, conversionResult.levelPositions, conversionResult.stemPositions);
            if (!frameCounts) {
                conversionResult.failedFrameCount++;
                continue;
            }
            conversionResult.symbolCount += frameCounts.symbolCount;
            conversionResult.indentCount += frameCounts.indentCount;
            conversionResult.spaceRunCount += frameCounts.spaceRunCount;
            if (frameCounts.boxPaths.length > 0) conversionResult.boxedFrames.push({ textFrame: textFrames[i], boxPaths: frameCounts.boxPaths });
        }
        return conversionResult;
    }

    // =========================================
    // 行ごとの分割 / Split into lines
    // =========================================

    /**
     * アイテムを1つのグループにまとめ、基準のアイテムの位置（重なり順）に置く
     * @param {PageItem} anchorItem - 置く位置の基準
     * @param {PageItem[]} items - まとめるアイテム（前面から順）
     * @returns {GroupItem} 作ったグループ
     */
    function groupItemsAt(anchorItem, items) {
        var itemGroup = anchorItem.parent.groupItems.add();
        itemGroup.move(anchorItem, ElementPlacement.PLACEBEFORE);
        for (var i = 0; i < items.length; i++) items[i].move(itemGroup, ElementPlacement.PLACEATEND);
        return itemGroup;
    }

    /**
     * 囲み罫を描いたテキストを行ごとのポイント文字に分け、行と囲み罫を1つずつグループにする。
     * 分けた行は、前の行を消して上がった分（行送りの合計）だけ下げて元の位置に戻す。
     * ポイント文字以外や、行と囲み罫の数が合わないときは分けずに、テキストと囲み罫をまとめる
     * @param {TextFrame} textFrame - 変換したテキスト（分けたら削除する）
     * @param {PathItem[]} boxPaths - そのテキストの囲み罫（上の行から順）
     * @returns {void}
     */
    function splitBoxedLines(textFrame, boxPaths) {
        var textContents = textFrame.contents;
        var textLines = listTextLines(textContents);
        var glyphLines = [];
        var leadingTotal = 0;
        for (var i = 0; i < textLines.length; i++) {
            var lineHead = textFrame.characters[Math.min(textLines[i].start, textContents.length - 1)];
            if (i > 0) leadingTotal += readCellMetrics(lineHead).leading;
            if (textLines[i].firstGlyph < 0) continue;
            /* 段落の書式は前の段落を消すと入れ替わりうるので、控えて当て直す / paragraph settings may change when earlier paragraphs go, so keep them */
            var lineParagraph = lineHead.paragraphAttributes;
            glyphLines.push({
                start: textLines[i].start,
                end: (i + 1 < textLines.length) ? textLines[i + 1].start - 1 : textContents.length,
                offsetY: leadingTotal,
                tabStops: lineParagraph.tabStops,
                justification: lineParagraph.justification
            });
        }
        if (textFrame.kind !== TextType.POINTTEXT || glyphLines.length < 2 || glyphLines.length !== boxPaths.length) {
            groupItemsAt(textFrame, [textFrame].concat(boxPaths));
            return;
        }
        for (i = 0; i < glyphLines.length; i++) {
            var glyphLine = glyphLines[i];
            var lineFrame = textFrame.duplicate();
            var lineEdits = [];
            if (glyphLine.end < textContents.length) lineEdits.push({ start: glyphLine.end, end: textContents.length, replacement: "" });
            if (glyphLine.start > 0) lineEdits.push({ start: 0, end: glyphLine.start, replacement: "" });
            applyTextEdits(lineFrame, lineEdits);
            var lineAttributes = lineFrame.textRange.paragraphAttributes;
            if (glyphLine.tabStops.length > 0) lineAttributes.tabStops = glyphLine.tabStops;
            lineAttributes.justification = glyphLine.justification;
            lineFrame.translate(0, -glyphLine.offsetY);
            groupItemsAt(textFrame, [lineFrame, boxPaths[i]]);
        }
        textFrame.remove();
    }

    /**
     * 変換結果の囲み罫を描いたテキストを、すべて行ごとに分けてグループにする。罫線のグループが空になったら消す
     * @param {Object} conversionResult - convertTextFrames() の戻り値
     * @returns {void}
     */
    function splitAllBoxedLines(conversionResult) {
        for (var i = 0; i < conversionResult.boxedFrames.length; i++) {
            splitBoxedLines(conversionResult.boxedFrames[i].textFrame, conversionResult.boxedFrames[i].boxPaths);
        }
        var lineGroup = conversionResult.lineGroup;
        if (lineGroup && lineGroup.pageItems.length === 0) lineGroup.remove();
    }

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

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューを作る。元のテキストを複製して隠し、複製のほうを変換する（確定のときは元をそのまま変換する）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 元のテキストフレーム
     * @returns {{show: Function, clear: Function, getItems: Function}} show(conversionSettings) で描き直して変換結果を返し、
     *     clear() で片付ける。getItems() はプレビューのテキストと罫線を返す
     */
    function createPreview(doc, textFrames) {
        var previewItems = [];
        var isShowing = false;

        /**
         * プレビューを片付け、元のテキストを表示に戻す
         * @returns {void}
         */
        function clear() {
            for (var i = 0; i < previewItems.length; i++) previewItems[i].remove();
            previewItems = [];
            if (isShowing) {
                for (i = 0; i < textFrames.length; i++) textFrames[i].hidden = false;
                isShowing = false;
            }
            app.redraw();
        }

        /**
         * 設定どおりのプレビューを描き直す
         * @param {Object} conversionSettings - convertTextFrames() と同じ
         * @returns {Object} convertTextFrames() の戻り値
         */
        function show(conversionSettings) {
            clear();
            /* 複製は hidden を引き継ぐので、元を隠す前に作る / duplicates inherit hidden, so copy before hiding */
            var previewFrames = [];
            for (var i = 0; i < textFrames.length; i++) previewFrames.push(textFrames[i].duplicate());
            for (i = 0; i < textFrames.length; i++) textFrames[i].hidden = true;
            isShowing = true;
            /* 変換の前に控え、途中で例外になっても複製と罫線を片付けて元を表示に戻す
               record the items before converting so a failure still cleans them up and shows the originals again */
            /* 罫線のグループは previewItems に足されるので、変換中の previewFrames とは別の配列にする
               line groups get pushed to previewItems, so keep it a separate array from the frames being converted */
            previewItems = previewFrames.slice();
            var conversionResult;
            try {
                conversionResult = convertTextFrames(doc, previewFrames, conversionSettings, previewItems);
            } catch (e) {
                clear();
                throw e;
            }
            app.redraw();
            return conversionResult;
        }

        /**
         * プレビューのテキストと罫線を返す
         * @returns {PageItem[]} プレビューのアイテム
         */
        function getItems() {
            return previewItems.slice(0);
        }

        return { show: show, clear: clear, getItems: getItems };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 位置のキーを、表示する順（位置の小さい順、説明は最後）に並べる
     * @param {Object} positions - キー → 位置（pt）
     * @returns {string[]} 並べたキー
     */
    function orderPositionKeys(positions) {
        var positionKeys = [];
        for (var positionKey in positions) {
            if (positions.hasOwnProperty(positionKey) && positionKey !== DESCRIPTION_LEVEL_KEY) positionKeys.push(positionKey);
        }
        /* 比較関数つきの sort() は使わず、位置の小さい順に選んで並べる / avoid sort() with a comparator */
        var orderedKeys = [];
        while (positionKeys.length > 0) {
            var leftmost = 0;
            for (var i = 1; i < positionKeys.length; i++) {
                if (positions[positionKeys[i]] < positions[positionKeys[leftmost]]) leftmost = i;
            }
            orderedKeys.push(positionKeys.splice(leftmost, 1)[0]);
        }
        if (positions.hasOwnProperty(DESCRIPTION_LEVEL_KEY)) orderedKeys.push(DESCRIPTION_LEVEL_KEY);
        return orderedKeys;
    }

    /**
     * 数値欄の値が変わったときの処理を、欄の既定の処理（値の正規化）のあとにつなぐ
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {Function} onValueChanged - 値が変わったときの処理
     * @returns {void}
     */
    function chainFieldChange(numberInput, onValueChanged) {
        var normalizeValue = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeValue();
            onValueChanged();
        };
    }

    /**
     * パネルを横に並べる行を追加する
     * @param {Window|Group} parent - 追加先
     * @returns {Group} パネルを入れる行
     */
    function addPanelColumns(parent) {
        var panelColumns = parent.add("group");
        panelColumns.orientation = "row";
        panelColumns.alignChildren = ["fill", "top"];
        panelColumns.spacing = COLUMN_SPACING;
        return panelColumns;
    }

    /**
     * 右揃えの項目名を先頭に置いた行を追加する（数値欄の項目名と幅をそろえる）
     * @param {Panel} parent - 追加先
     * @param {string} labelPath - 項目名の LABELS のパス
     * @param {number} rowSpacing - 行の中の間隔
     * @returns {Group} 行
     */
    function addLabeledRow(parent, labelPath, rowSpacing) {
        var labeledRow = parent.add("group");
        setupRow(labeledRow, "left", rowSpacing);
        var rowLabel = labeledRow.add("statictext", undefined, labelText(labelPath));
        rowLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        rowLabel.justify = "right";
        return labeledRow;
    }

    /**
     * パネルの右下に置く、ひとまわり小さいボタンを追加する（文字を小さくする。高さは表示後に詰める）
     * @param {Panel} parent - 追加先
     * @param {string} buttonText - ボタンの文言
     * @param {Button[]} compactButtons - 表示後に高さを詰めるボタンの一覧。ここに足す
     * @returns {Button} 作ったボタン
     */
    function addCompactButton(parent, buttonText, compactButtons) {
        var compactButton = parent.add("button", undefined, buttonText);
        compactButton.alignment = ["right", "bottom"];
        var baseFont = compactButton.graphics.font;
        compactButton.graphics.font = ScriptUI.newFont(baseFont.name, "REGULAR", baseFont.size - COMPACT_BUTTON_FONT_SHRINK);
        compactButtons.push(compactButton);
        return compactButton;
    }

    /**
     * 欄の値が次の欄に追いついたら次の欄を NEIGHBOR_PUSH_GAP だけ先へ、前の欄に追いついたら前の欄を同じだけ手前へ押し出す
     * （その先へも順に）。押し出した欄は onPushed で固定する
     * @param {EditText[]} orderedInputs - 左から順の欄
     * @param {EditText} changedInput - 値が変わった欄
     * @param {Function} onPushed - 押し出した欄を受け取る関数
     * @returns {void}
     */
    function pushNeighbors(orderedInputs, changedInput, onPushed) {
        var changedIndex = -1;
        for (var i = 0; i < orderedInputs.length; i++) {
            if (orderedInputs[i] === changedInput) changedIndex = i;
        }
        if (changedIndex < 0) return;
        var value = parseFloat(changedInput.text);
        for (i = changedIndex + 1; i < orderedInputs.length && parseFloat(orderedInputs[i].text) <= value; i++) {
            writeSteppedValue(orderedInputs[i], value + NEIGHBOR_PUSH_GAP, orderedInputs[i].stepOptions);
            value = parseFloat(orderedInputs[i].text);
            onPushed(orderedInputs[i]);
        }
        value = parseFloat(changedInput.text);
        for (i = changedIndex - 1; i >= 0 && parseFloat(orderedInputs[i].text) >= value; i--) {
            writeSteppedValue(orderedInputs[i], value - NEIGHBOR_PUSH_GAP, orderedInputs[i].stepOptions);
            value = parseFloat(orderedInputs[i].text);
            onPushed(orderedInputs[i]);
        }
    }

    /**
     * タブストップ・縦罫の位置の欄をまとめて作る。値を変えるとその位置に固定し、隣に追いついたら押し出す。
     * 連動の対象が3つ以上あれば、3つめの行にリンクアイコンを置く（オンのとき、1と2の差の間隔で3以降を並べる）。
     * 連動の対象より後ろにある欄（説明）は、連動で決まる値の上限になる
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} editorOptions - entries（{ key, label, tooltip, linkable } の配列、左から順）/ positions（キー → 自動の位置 pt）/
     *     overrides（キー → 固定した位置 pt。書き換える）/ unit（getUnitInfo() の戻り値）/ linkDefault / linkTooltip / resetTooltip /
     *     compactButtons / onChanged / isEntryEnabled(entry)（省略時は常に有効）/ isLinkEnabled()（省略時は常に有効）
     * @returns {{updateEnabled: Function, applyLink: Function, isLinkOn: Function}} 有効／無効の更新・連動の適用・連動中かの判定
     */
    function createPositionEditor(parentPanel, editorOptions) {
        var orderedInputs = [];
        var linkedInputs = [];
        var linkToggle = null;
        var i;
        for (i = 0; i < editorOptions.entries.length; i++) addPositionField(editorOptions.entries[i]);
        if (linkedInputs.length >= 3) {
            /* 連動は1と2の差で3以降を決めるので、アイコンは3つめの行に置く / the icon sits on the row of the third field */
            linkToggle = addLinkToggle(linkedInputs[2].fieldLabel.parent, editorOptions.linkDefault, onLinkToggled);
            linkToggle.helpTip = editorOptions.linkTooltip;
        }
        var btnReset = addCompactButton(parentPanel, getLabel("button.reset"), editorOptions.compactButtons);
        btnReset.helpTip = editorOptions.resetTooltip;
        btnReset.onClick = reset;

        /**
         * 位置の欄を1つ追加する
         * @param {Object} entry - { key, label, tooltip, linkable }
         * @returns {void}
         */
        function addPositionField(entry) {
            var fieldOptions = {
                label: entry.label,
                labelWidth: FIELD_LABEL_WIDTH,
                characters: NUMBER_FIELD_CHARS,
                step: 1,
                min: 0,
                unit: " " + editorOptions.unit.label
            };
            fieldOptions.text = formatSteppedValue(editorOptions.positions[entry.key] / editorOptions.unit.pointsPerUnit, fieldOptions);
            fieldOptions.onStep = onFieldChanged;
            var positionInput = addSteppedField(parentPanel, fieldOptions);
            positionInput.helpTip = entry.tooltip;
            positionInput.stepOptions = fieldOptions;
            positionInput.entry = entry;
            positionInput.linkIndex = entry.linkable ? linkedInputs.length : -1;
            chainFieldChange(positionInput, onFieldChanged);
            orderedInputs.push(positionInput);
            if (entry.linkable) linkedInputs.push(positionInput);

            /**
             * 入力値に固定し、隣を押し出す。連動中に1・2を変えたら3以降を合わせ直す
             * @returns {void}
             */
            function onFieldChanged() {
                pushNeighbors(orderedInputs, positionInput, fixInput);
                fixInput(positionInput);
                if (positionInput.linkIndex >= 0 && positionInput.linkIndex <= 1 && isLinkOn()) applyLink();
                editorOptions.onChanged();
            }
        }

        /**
         * 欄の値で、その位置に固定する
         * @param {EditText} positionInput - 位置の欄
         * @returns {void}
         */
        function fixInput(positionInput) {
            editorOptions.overrides[positionInput.entry.key] = parseFloat(positionInput.text) * editorOptions.unit.pointsPerUnit;
        }

        /**
         * 連動がオンかを返す
         * @returns {boolean} オンなら true
         */
        function isLinkOn() {
            return !!(linkToggle && linkToggle.value);
        }

        /**
         * 連動の対象より後ろにある最初の欄（説明）の値を返す。連動で決まる値はこれを超えない
         * @returns {number} 上限（無ければ Infinity）
         */
        function getUpperLimit() {
            for (var k = 0; k < orderedInputs.length; k++) {
                if (orderedInputs[k].linkIndex < 0 && k > 0) return parseFloat(orderedInputs[k].text);
            }
            return Infinity;
        }

        /**
         * 1と2の差（A）で、3以降を「ひとつ前 ＋ A」にそろえて固定する
         * @returns {void}
         */
        function applyLink() {
            var firstValue = parseFloat(linkedInputs[0].text);
            var secondValue = parseFloat(linkedInputs[1].text);
            if (isNaN(firstValue) || isNaN(secondValue)) return;
            var linkInterval = secondValue - firstValue;
            var upperLimit = getUpperLimit();
            fixInput(linkedInputs[0]);
            fixInput(linkedInputs[1]);
            var previousValue = secondValue;
            for (var k = 2; k < linkedInputs.length; k++) {
                writeSteppedValue(linkedInputs[k], Math.min(previousValue + linkInterval, upperLimit), linkedInputs[k].stepOptions);
                fixInput(linkedInputs[k]);
                previousValue = parseFloat(linkedInputs[k].text);
            }
        }

        /**
         * 欄の有効／無効を合わせる（連動中は3以降を1・2から決めるので入力できない）
         * @returns {void}
         */
        function updateEnabled() {
            for (var k = 0; k < orderedInputs.length; k++) {
                var isFollowing = isLinkOn() && orderedInputs[k].linkIndex >= 2;
                var isEntryEnabled = editorOptions.isEntryEnabled ? editorOptions.isEntryEnabled(orderedInputs[k].entry) : true;
                setSteppedFieldEnabled(orderedInputs[k], isEntryEnabled && !isFollowing);
            }
            if (linkToggle) setLinkToggleEnabled(linkToggle, editorOptions.isLinkEnabled ? editorOptions.isLinkEnabled() : true);
        }

        /**
         * 連動を切り替えたとき、オンなら並べ直して描き直す
         * @returns {void}
         */
        function onLinkToggled() {
            if (isLinkOn()) applyLink();
            updateEnabled();
            editorOptions.onChanged();
        }

        /**
         * 自動の位置に戻す（固定と連動を外す）
         * @returns {void}
         */
        function reset() {
            for (var k = 0; k < orderedInputs.length; k++) {
                var entryKey = orderedInputs[k].entry.key;
                writeSteppedValue(orderedInputs[k], editorOptions.positions[entryKey] / editorOptions.unit.pointsPerUnit, orderedInputs[k].stepOptions);
                delete editorOptions.overrides[entryKey];
            }
            if (linkToggle) setLinkToggleValue(linkToggle, false);
            updateEnabled();
            editorOptions.onChanged();
        }

        return { updateEnabled: updateEnabled, applyLink: applyLink, isLinkOn: isLinkOn };
    }

    /**
     * ［罫線］パネル（線幅・線端・カラー・横線を囲み罫につなげる）を作る
     * @param {Group} parent - 追加先
     * @param {Function} onChanged - 値が変わったときの処理
     * @returns {{focusInput: EditText, attachArmsCheckbox: Checkbox, readSettings: Function}} 最初に選ぶ欄・横線をつなげるチェックボックスと、
     *     設定を書き込む関数 readSettings(conversionSettings)
     */
    function buildLinesPanel(parent, onChanged) {
        var strokeUnit = getUnitInfo("strokeUnits");
        var linesPanel = parent.add("panel", undefined, getLabel("panel.lines"));
        setupPanel(linesPanel);

        var strokeWidthOptions = { label: labelText("fieldLabel.strokeWidth"), labelWidth: FIELD_LABEL_WIDTH, characters: NUMBER_FIELD_CHARS, step: 1, min: 0.01, unit: " " + strokeUnit.label };
        strokeWidthOptions.text = formatSteppedValue(DEFAULT_STROKE_WIDTH_PT / strokeUnit.pointsPerUnit, strokeWidthOptions);
        strokeWidthOptions.onStep = onChanged;
        var strokeWidthInput = addSteppedField(linesPanel, strokeWidthOptions);
        strokeWidthInput.helpTip = getLabel("tooltip.strokeWidth");
        chainFieldChange(strokeWidthInput, onChanged);

        /* 線端：丸型／なし（同じ親に入れて排他にする） / cap: round or butt, in one parent so they stay exclusive */
        var capRow = addLabeledRow(linesPanel, "fieldLabel.strokeCap", CAP_RADIO_SPACING);
        var roundCapRadio = capRow.add("radiobutton", undefined, getLabel("radio.roundCap"));
        roundCapRadio.helpTip = getLabel("tooltip.roundCap");
        var buttCapRadio = capRow.add("radiobutton", undefined, getLabel("radio.buttCap"));
        buttCapRadio.helpTip = getLabel("tooltip.buttCap");
        roundCapRadio.value = ROUND_CAP;
        buttCapRadio.value = !ROUND_CAP;
        roundCapRadio.onClick = onChanged;
        buttCapRadio.onClick = onChanged;

        /* カラー：■ #RRGGBB。■はクリックで標準のカラーピッカー / swatch opens the standard color picker */
        var lineColor = createLineColor();
        var colorRow = addLabeledRow(linesPanel, "fieldLabel.lineColor", 0);
        var colorSwatch = colorRow.add("panel");
        colorSwatch.preferredSize = [COLOR_SWATCH_SIZE, COLOR_SWATCH_SIZE];
        colorSwatch.helpTip = getLabel("tooltip.lineColorSwatch");
        /* 「#」と欄はまとめ、色見本とのあいだに余白を空ける / keep "#" with the field, spaced from the swatch */
        var colorHexGroup = colorRow.add("group");
        colorHexGroup.orientation = "row";
        colorHexGroup.alignChildren = ["left", "center"];
        colorHexGroup.spacing = 0;
        colorHexGroup.margins = [COLOR_HEX_LEFT_MARGIN, 0, 0, 0];
        colorHexGroup.add("statictext", undefined, getLabel("fieldLabel.hexPrefix"));
        var colorHexInput = colorHexGroup.add("edittext", undefined, "");
        colorHexInput.characters = COLOR_HEX_FIELD_CHARS;
        colorHexInput.helpTip = getLabel("tooltip.lineColorHex");
        showLineColor();

        /* 横線を囲み罫につなげる（項目名の幅だけ字下げ） / attach the arms to the boxes, indented past the labels */
        var attachArmsGroup = linesPanel.add("group");
        attachArmsGroup.margins = [FIELD_LABEL_WIDTH + STEPPER_FIELD_SPACING, 0, 0, 0];
        var attachArmsCheckbox = attachArmsGroup.add("checkbox", undefined, getLabel("checkbox.attachArmsToBox"));
        attachArmsCheckbox.value = ATTACH_ARMS_TO_BOX;
        attachArmsCheckbox.helpTip = getLabel("tooltip.attachArmsToBox");
        attachArmsCheckbox.onClick = onChanged;

        /**
         * 罫線の色を、色見本と16進数の欄に表示する
         * @returns {void}
         */
        function showLineColor() {
            var displayColor = toRGBColor(lineColor);
            if (!displayColor) return; /* 特色など RGB に直せない色は表示を変えない / keep the display for colors without an RGB form */
            colorHexInput.text = toHexColor(displayColor).substring(1);
            colorHexInput.lastValidText = colorHexInput.text;
            var swatchGraphics = colorSwatch.graphics;
            swatchGraphics.backgroundColor = swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR,
                [displayColor.red / 255, displayColor.green / 255, displayColor.blue / 255, 1]);
        }

        colorSwatch.addEventListener("mousedown", function () {
            /* 選んだ色は戻り値で返り、型はドキュメントのカラーモードに従う。キャンセルでは渡した色が返る
               the picked color is returned in the document's color model; Cancel returns the color passed in */
            lineColor = app.showColorPicker(lineColor);
            showLineColor();
            onChanged();
        });
        colorHexInput.onChange = function () {
            /* 形式が違えば直前の値に戻す / revert when the hex is malformed */
            var typedColor = parseHexColor(colorHexInput.text);
            if (!typedColor) {
                colorHexInput.text = colorHexInput.lastValidText;
                return;
            }
            lineColor = typedColor;
            showLineColor();
            onChanged();
        };

        return {
            focusInput: strokeWidthInput,
            attachArmsCheckbox: attachArmsCheckbox,
            /**
             * 線幅・色・線端・横線のつなげ方を設定に書き込む
             * @param {Object} conversionSettings - 書き込む先
             * @returns {void}
             */
            readSettings: function (conversionSettings) {
                var strokeWidthValue = parseFloat(strokeWidthInput.text);
                conversionSettings.strokeWidthPt = (strokeWidthValue > 0) ? strokeWidthValue * strokeUnit.pointsPerUnit : DEFAULT_STROKE_WIDTH_PT;
                conversionSettings.strokeColor = lineColor;
                conversionSettings.roundCap = roundCapRadio.value;
                conversionSettings.attachArmsToBox = attachArmsCheckbox.value;
            }
        };
    }

    /**
     * ［テキスト］パネル（行送り・タブ変換・囲み罫）を作る
     * @param {Group} parent - 追加先
     * @param {TextFrame} firstTextFrame - 行送りの初期値を読むテキスト
     * @param {Function} onChanged - 行送り・囲み罫が変わったときの処理
     * @param {Function} onTabOptionChanged - タブ変換のチェックを切り替えたときの処理
     * @returns {{indentCheckbox: Checkbox, spaceRunCheckbox: Checkbox, boxCheckbox: Checkbox, readSettings: Function}} チェックボックスと、
     *     設定を書き込む関数 readSettings(conversionSettings, enableAll)
     */
    function buildTextPanel(parent, firstTextFrame, onChanged, onTabOptionChanged) {
        var leadingUnit = getUnitInfo("text/units");
        var leadingOverridePt = null; /* 変えるまでは元の行送りのまま / keep the original leading until changed */
        var textPanel = parent.add("panel", undefined, getLabel("panel.text"));
        setupPanel(textPanel);

        var leadingOptions = { label: labelText("fieldLabel.leading"), labelWidth: FIELD_LABEL_WIDTH, characters: NUMBER_FIELD_CHARS, step: 1, min: 0.1, unit: " " + leadingUnit.label };
        leadingOptions.text = formatSteppedValue(readCellMetrics(firstTextFrame.characters[0]).leading / leadingUnit.pointsPerUnit, leadingOptions);
        leadingOptions.onStep = onLeadingChanged;
        var leadingInput = addSteppedField(textPanel, leadingOptions);
        leadingInput.helpTip = getLabel("tooltip.leading");
        chainFieldChange(leadingInput, onLeadingChanged);

        /**
         * 行送りを入力値に固定して描き直す
         * @returns {void}
         */
        function onLeadingChanged() {
            leadingOverridePt = parseFloat(leadingInput.text) * leadingUnit.pointsPerUnit;
            onChanged();
        }

        var tabOptionGroup = textPanel.add("group");
        tabOptionGroup.orientation = "column";
        tabOptionGroup.alignChildren = ["left", "top"];
        tabOptionGroup.spacing = CHECKBOX_SPACING;
        var indentCheckbox = tabOptionGroup.add("checkbox", undefined, getLabel("checkbox.convertIndents"));
        indentCheckbox.value = CONVERT_INDENTS;
        indentCheckbox.helpTip = getLabel("tooltip.convertIndents");
        var spaceRunCheckbox = tabOptionGroup.add("checkbox", undefined, getLabel("checkbox.convertSpaceRuns"));
        spaceRunCheckbox.value = CONVERT_SPACE_RUNS;
        spaceRunCheckbox.helpTip = getLabel("tooltip.convertSpaceRuns");
        indentCheckbox.onClick = onTabOptionChanged;
        spaceRunCheckbox.onClick = onTabOptionChanged;
        var boxCheckbox = tabOptionGroup.add("checkbox", undefined, getLabel("checkbox.drawBox"));
        boxCheckbox.value = DRAW_BOX;
        boxCheckbox.helpTip = getLabel("tooltip.drawBox");

        /* マージンは定規の単位で入れる。左右・上下の2欄を縦に積み、右に連動アイコンを置く
           the margins are in ruler units; the two fields are stacked with the link icon on their right */
        var boxMarginUnit = getUnitInfo("rulerType");
        /* 初期値は先頭の文字の文字サイズから決める / the default follows the size of the first character */
        var defaultBoxMarginPt = firstTextFrame.characters[0].characterAttributes.size * BOX_MARGIN_SIZE_RATIO;
        var boxMarginRow = textPanel.add("group");
        boxMarginRow.orientation = "row";
        boxMarginRow.alignChildren = ["left", "center"];
        boxMarginRow.spacing = STEPPER_FIELD_SPACING;
        var boxMarginColumn = boxMarginRow.add("group");
        boxMarginColumn.orientation = "column";
        boxMarginColumn.alignChildren = ["left", "top"];
        boxMarginColumn.spacing = CHECKBOX_SPACING;
        var boxMarginXInput = addBoxMarginField(boxMarginColumn, "fieldLabel.boxMarginX", "tooltip.boxMarginX");
        var boxMarginYInput = addBoxMarginField(boxMarginColumn, "fieldLabel.boxMarginY", "tooltip.boxMarginY");
        boxMarginXInput.linkedInput = boxMarginYInput;
        boxMarginYInput.linkedInput = boxMarginXInput;
        var boxMarginLink = addLinkToggle(boxMarginRow, LINK_BOX_MARGINS, function () {
            /* 連動にしたときは左右の値を上下へ写す / when linked, copy left/right to top/bottom */
            if (boxMarginLink.value) syncBoxMargin(boxMarginXInput);
            onChanged();
        });
        boxMarginLink.helpTip = getLabel("tooltip.linkBoxMargins");

        /**
         * マージンの欄を1つ追加する（連動中は、もう一方の欄にも同じ値を入れる）
         * @param {Group} parent - 追加先
         * @param {string} labelPath - 項目名の LABELS のパス
         * @param {string} tooltipPath - ツールチップの LABELS のパス
         * @returns {EditText} 入力欄
         */
        function addBoxMarginField(parent, labelPath, tooltipPath) {
            var marginOptions = { label: labelText(labelPath), labelWidth: FIELD_LABEL_WIDTH, characters: NUMBER_FIELD_CHARS, step: 1, min: 0, unit: " " + boxMarginUnit.label };
            marginOptions.text = formatSteppedValue(defaultBoxMarginPt / boxMarginUnit.pointsPerUnit, marginOptions);
            var marginInput;
            marginOptions.onStep = function () { onBoxMarginChanged(marginInput); };
            marginInput = addSteppedField(parent, marginOptions);
            marginInput.helpTip = getLabel(tooltipPath);
            marginInput.stepOptions = marginOptions;
            chainFieldChange(marginInput, function () { onBoxMarginChanged(marginInput); });
            return marginInput;
        }

        /**
         * マージンが変わったとき、連動中ならもう一方へ写して描き直す
         * @param {EditText} changedInput - 値が変わった欄
         * @returns {void}
         */
        function onBoxMarginChanged(changedInput) {
            if (boxMarginLink.value) syncBoxMargin(changedInput);
            onChanged();
        }

        /**
         * 欄の値を、もう一方の欄に写す
         * @param {EditText} sourceInput - 写す元の欄
         * @returns {void}
         */
        function syncBoxMargin(sourceInput) {
            writeSteppedValue(sourceInput.linkedInput, parseFloat(sourceInput.text), sourceInput.linkedInput.stepOptions);
        }

        /**
         * マージンの欄を pt で読む（読めなければ初期値）
         * @param {EditText} marginInput - マージンの欄
         * @returns {number} マージン（pt）
         */
        function readBoxMarginPt(marginInput) {
            var marginValue = parseFloat(marginInput.text);
            return (marginValue >= 0) ? marginValue * boxMarginUnit.pointsPerUnit : defaultBoxMarginPt;
        }
        /* マージンの欄の下に字下げして置く / indented under the margin fields */
        var boxOptionGroup = textPanel.add("group");
        boxOptionGroup.orientation = "column";
        boxOptionGroup.alignChildren = ["left", "top"];
        boxOptionGroup.spacing = CHECKBOX_SPACING;
        boxOptionGroup.margins = [FIELD_LABEL_WIDTH + STEPPER_FIELD_SPACING, 0, 0, 0];
        var alignBoxRightCheckbox = boxOptionGroup.add("checkbox", undefined, getLabel("checkbox.alignBoxRight"));
        alignBoxRightCheckbox.value = ALIGN_BOX_RIGHT;
        alignBoxRightCheckbox.helpTip = getLabel("tooltip.alignBoxRight");
        alignBoxRightCheckbox.onClick = onChanged;
        /* 分割は OK のあとに行うので、プレビューは描き直さない / the split happens after OK, so no redraw */
        var splitLinesCheckbox = boxOptionGroup.add("checkbox", undefined, getLabel("checkbox.splitBoxedLines"));
        splitLinesCheckbox.value = SPLIT_BOXED_LINES;
        splitLinesCheckbox.helpTip = getLabel("tooltip.splitBoxedLines");

        /**
         * マージン・右端の揃え・分割を、囲み罫のオン／オフに合わせて有効／無効にする
         * @returns {void}
         */
        function updateBoxEnabled() {
            setSteppedFieldEnabled(boxMarginXInput, boxCheckbox.value);
            setSteppedFieldEnabled(boxMarginYInput, boxCheckbox.value);
            setLinkToggleEnabled(boxMarginLink, boxCheckbox.value);
            alignBoxRightCheckbox.enabled = boxCheckbox.value;
            splitLinesCheckbox.enabled = boxCheckbox.value;
        }
        updateBoxEnabled();
        boxCheckbox.onClick = function () {
            updateBoxEnabled();
            onChanged();
        };

        return {
            indentCheckbox: indentCheckbox,
            spaceRunCheckbox: spaceRunCheckbox,
            boxCheckbox: boxCheckbox,
            /**
             * 行送り・タブ変換・囲み罫を設定に書き込む
             * @param {Object} conversionSettings - 書き込む先
             * @param {boolean} [enableAll] - true なら字下げ・スペースの変換をどちらも有効にする（最初のプレビュー用）
             * @returns {void}
             */
            readSettings: function (conversionSettings, enableAll) {
                conversionSettings.leadingPt = leadingOverridePt;
                conversionSettings.convertIndents = enableAll || indentCheckbox.value;
                conversionSettings.convertSpaceRuns = enableAll || spaceRunCheckbox.value;
                conversionSettings.drawBox = boxCheckbox.value;
                conversionSettings.alignBoxRight = alignBoxRightCheckbox.value;
                conversionSettings.splitBoxedLines = boxCheckbox.value && splitLinesCheckbox.value;
                conversionSettings.boxMarginXPt = readBoxMarginPt(boxMarginXInput);
                conversionSettings.boxMarginYPt = readBoxMarginPt(boxMarginYInput);
            }
        };
    }

    /**
     * ダイアログ下部のボタン行（制御文字・画面にフィット・キャンセル・OK）を作る
     * @param {Window} dialog - ダイアログ
     * @param {Document} doc - 対象のドキュメント
     * @param {Function} getPreviewItems - プレビューのアイテムを返す関数
     * @returns {{fitViewToPreview: Function, restoreView: Function}} 表示を合わせる関数と、開く前の表示に戻す関数
     */
    function buildDialogButtons(dialog, doc, getPreviewItems) {
        var buttonRow = addButtonRow(dialog);
        var btnShowHiddenChars = buttonRow.leftGroup.add("button", undefined, getLabel("button.showHiddenChars"));
        btnShowHiddenChars.helpTip = getLabel("tooltip.showHiddenChars");
        btnShowHiddenChars.onClick = function () {
            app.executeMenuCommand("showHiddenChar");
            app.redraw();
        };
        var fitViewControls = FitViewToItems.addControls(buttonRow.leftGroup, { lang: uiLang });
        var initialViewState = FitViewToItems.captureView(doc); /* キャンセルで戻す表示 / view restored on Cancel */
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * ［画面にフィット］がオンなら、プレビューが指定の割合で収まるよう表示を合わせる
         * @returns {void}
         */
        function fitViewToPreview() {
            if (!fitViewControls.checkbox.value) return;
            FitViewToItems.fit(getPreviewItems(), { doc: doc, fillRatio: fitViewControls.getFillRatio() });
            app.redraw();
        }

        fitViewControls.checkbox.onClick = function () {
            fitViewControls.updateEnabled();
            fitViewToPreview();
        };
        fitViewControls.percentInput.onChange = fitViewToPreview;
        /* 部品の％欄には↑↓が無いので、ほかの数値欄と同じ増減をつなぐ（10〜100%、整数） / arrow keys for the fit percentage */
        var fitPercentOptions = { step: 1, min: 10, max: 100, integer: true };
        bindSteppedArrowKeys(fitViewControls.percentInput, {
            stepBy: function (direction) {
                var percentInput = fitViewControls.percentInput;
                if (!percentInput.enabled) return;
                var percentValue = parseFloat(percentInput.text);
                if (isNaN(percentValue)) percentValue = 0;
                writeSteppedValue(percentInput, computeSteppedValue(percentValue, direction, fitPercentOptions), fitPercentOptions);
                fitViewToPreview();
            }
        });

        return {
            fitViewToPreview: fitViewToPreview,
            /**
             * ダイアログを開く前の表示位置と倍率に戻す
             * @returns {void}
             */
            restoreView: function () {
                FitViewToItems.restoreView(initialViewState, doc);
            }
        };
    }

    /**
     * 線幅・線端・カラー・行送り・タブ変換・囲み罫・タブストップ・縦罫を指定するダイアログを表示する（プレビューは常に表示）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {Object|null} convertTextFrames() に渡す設定。キャンセルなら null
     */
    function showSettingsDialog(doc, textFrames) {
        var tabUnit = getUnitInfo("rulerType");
        var tabStopOverrides = {};
        var stemOverrides = {};
        var compactButtons = []; /* 表示後に高さを詰めるボタン / buttons trimmed after the layout */
        var textPreview = createPreview(doc, textFrames);
        var tabStopEditor = null;
        var stemEditor = null;
        var dialogButtons = null;

        var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(settingsDialog);

        /* 罫線とテキストのパネルは横に並べる / the lines and text panels sit side by side */
        var settingsColumns = addPanelColumns(settingsDialog);
        var linesControls = buildLinesPanel(settingsColumns, refreshPreview);
        var textControls = buildTextPanel(settingsColumns, textFrames[0], refreshPreview, function () {
            if (tabStopEditor) tabStopEditor.updateEnabled();
            refreshPreview();
        });
        /* ［横線を囲み罫につなげる］は［囲み罫］がオンのときだけ使える / the arm option needs Box */
        var onBoxClicked = textControls.boxCheckbox.onClick;
        textControls.boxCheckbox.onClick = function () {
            linesControls.attachArmsCheckbox.enabled = textControls.boxCheckbox.value;
            onBoxClicked();
        };
        linesControls.attachArmsCheckbox.enabled = textControls.boxCheckbox.value;

        /* 自動で決めた位置を知るために、すべて有効にして一度プレビューする / preview once with everything on to learn the automatic positions */
        var initialPreviewResult = textPreview.show(readSettings(true));
        var levelKeys = orderPositionKeys(initialPreviewResult.levelPositions);
        var stemKeys = orderPositionKeys(initialPreviewResult.stemPositions);

        /* タブストップと縦罫のパネルは横に並べる / the tab stop and vertical rule panels sit side by side */
        var positionColumns = addPanelColumns(settingsDialog);
        if (levelKeys.length > 0) {
            var tabStopPanel = positionColumns.add("panel", undefined, getLabel("panel.tabStops"));
            setupPanel(tabStopPanel);
            var tabStopEntries = [];
            for (var i = 0; i < levelKeys.length; i++) {
                var isDescription = (levelKeys[i] === DESCRIPTION_LEVEL_KEY);
                tabStopEntries.push({
                    key: levelKeys[i],
                    label: isDescription ? labelText("fieldLabel.descriptionStop") : labelText("fieldLabel.levelStop", [i + 1]),
                    tooltip: getLabel(isDescription ? "tooltip.descriptionStop" : "tooltip.levelStop"),
                    linkable: !isDescription,
                    isDescription: isDescription
                });
            }
            tabStopEditor = createPositionEditor(tabStopPanel, {
                entries: tabStopEntries,
                positions: initialPreviewResult.levelPositions,
                overrides: tabStopOverrides,
                unit: tabUnit,
                linkDefault: LINK_LEVELS_EVENLY,
                linkTooltip: getLabel("tooltip.linkLevels"),
                resetTooltip: getLabel("tooltip.resetTabStops"),
                compactButtons: compactButtons,
                onChanged: refreshPreview,
                /* 名前のレベルは字下げの変換、説明はスペースの変換に従う / levels follow the indent option, the description the space-run option */
                isEntryEnabled: function (entry) {
                    return entry.isDescription ? textControls.spaceRunCheckbox.value : textControls.indentCheckbox.value;
                },
                isLinkEnabled: function () {
                    return textControls.indentCheckbox.value;
                }
            });
        }
        if (stemKeys.length > 0) {
            var stemPanel = positionColumns.add("panel", undefined, getLabel("panel.stems"));
            setupPanel(stemPanel);
            var stemEntries = [];
            for (i = 0; i < stemKeys.length; i++) {
                stemEntries.push({ key: stemKeys[i], label: labelText("fieldLabel.stemPosition", [i + 1]), tooltip: getLabel("tooltip.stemPosition"), linkable: true });
            }
            stemEditor = createPositionEditor(stemPanel, {
                entries: stemEntries,
                positions: initialPreviewResult.stemPositions,
                overrides: stemOverrides,
                unit: tabUnit,
                linkDefault: LINK_STEMS_EVENLY,
                linkTooltip: getLabel("tooltip.linkStems"),
                resetTooltip: getLabel("tooltip.resetStems"),
                compactButtons: compactButtons,
                onChanged: refreshPreview
            });
        }

        dialogButtons = buildDialogButtons(settingsDialog, doc, textPreview.getItems);

        /**
         * ダイアログの値から変換の設定を作る
         * @param {boolean} [enableAll] - true なら字下げ・スペースの変換をどちらも有効にする（最初のプレビュー用）
         * @returns {Object} convertTextFrames() に渡す設定
         */
        function readSettings(enableAll) {
            var conversionSettings = { tabStopOverrides: tabStopOverrides, stemOverrides: stemOverrides };
            linesControls.readSettings(conversionSettings);
            textControls.readSettings(conversionSettings, enableAll);
            return conversionSettings;
        }

        /**
         * プレビューを今の値で描き直し、［画面にフィット］がオンなら表示を合わせる
         * @returns {void}
         */
        function refreshPreview() {
            textPreview.show(readSettings());
            if (dialogButtons) dialogButtons.fitViewToPreview();
        }

        var positionEditors = [tabStopEditor, stemEditor];
        var needsRedraw = !textControls.indentCheckbox.value || !textControls.spaceRunCheckbox.value;
        for (i = 0; i < positionEditors.length; i++) {
            if (!positionEditors[i]) continue;
            positionEditors[i].updateEnabled();
            if (positionEditors[i].isLinkOn()) {
                positionEditors[i].applyLink();
                needsRedraw = true;
            }
        }
        /* 最初のプレビューは全部有効・連動なしで作ったので、違えば今の状態で描き直す / redraw when the first preview differs from the actual options */
        if (needsRedraw) refreshPreview();

        /* 高さはレイアウトが決まってから詰める（prepareDialogWindow より前に入れ、あとに位置の復元をつなぐ）
           trim heights once the layout exists; set before prepareDialogWindow so it chains after */
        settingsDialog.onShow = function () {
            for (var k = 0; k < compactButtons.length; k++) trimButtonHeight(compactButtons[k], COMPACT_BUTTON_TRIM_PX);
        };
        linesControls.focusInput.active = true;
        prepareDialogWindow(settingsDialog, SCRIPT_NAME);
        var dialogResult = settingsDialog.show();
        textPreview.clear();
        if (dialogResult !== 1) {
            dialogButtons.restoreView();
            return null;
        }
        return readSettings();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、ダイアログで設定を受け取って変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        /* 空のテキストは変換するものが無く、行送りの初期値も読めないので外す / drop empty text: nothing to convert, no leading to read */
        var selectedFrames = collectSelectionTextFrames(doc.selection);
        var textFrames = [];
        for (var i = 0; i < selectedFrames.length; i++) {
            if (selectedFrames[i].contents !== "") textFrames.push(selectedFrames[i]);
        }
        if (textFrames.length === 0) {
            alert(getLabel("alert.noTextFrame"));
            return;
        }

        /* 記号が無ければダイアログを出さない / skip the dialog when there is nothing to convert */
        if (!hasConvertibleText(textFrames)) {
            alert(getLabel("alert.noSymbol"));
            return;
        }

        var conversionSettings = showSettingsDialog(doc, textFrames);
        if (!conversionSettings) return;

        /* 成功したときは知らせない。位置を測れなかったテキストがあるときだけ伝える / only report the frames that could not be measured */
        var conversionResult = convertTextFrames(doc, textFrames, conversionSettings);
        if (conversionSettings.splitBoxedLines) splitAllBoxedLines(conversionResult);
        if (conversionResult.failedFrameCount > 0) alert(getLabel("alert.measureFailed", [conversionResult.failedFrameCount]));
    }

    main();

})();
