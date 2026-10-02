#target illustrator
#targetengine "SlideCollageEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選んだ .ai / .pdf のアートボード（PDFはページ）をグリッドに並べ、ポートフォリオ用のサムネイル一覧を作成します。
見開きを片ページに分けたり、回転・マスク・背景色を付けたりできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9f8c7370f4e5

### Overview

Lays out the artboards (or PDF pages) of the .ai / .pdf file you choose in a grid to build a portfolio-style thumbnail sheet.
Spreads can be split into single pages, and the layout can be rotated, masked and given a background color.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SlideCollage.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SlideCollage";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SlideCollage.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SlideCollage.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9f8c7370f4e5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ダイアログを開いたときの値。距離は pt、角丸の半径は定規の単位で持つ
       Values the dialog opens with. Distances in pt; corner radii in ruler units */
    var DEFAULT_SETTINGS = {
        cropIndex: 2,              /* 配置範囲（0 アート / 1 トリミング / 2 仕上がり / 3 裁ち落とし） / crop box */
        roundEnabled: false,       /* 角丸 / round corners */
        roundRadius: 10,           /* 角丸の半径（定規の単位） / corner radius (ruler units) */
        splitSpreads: false,       /* 見開きを左右に分割 / split spreads */
        flowMode: 1,               /* 方向（0 横 / 1 縦 / 2 ランダム） / flow */
        columns: 5,                /* 列数 / columns */
        spacingPt: 20,             /* 間隔 / gap */
        evenPlusSlot: false,       /* 偶数列に＋1スロット / one extra slot in even columns */
        evenShiftEnabled: true,    /* 偶数列のずらし / offset even columns */
        evenShift: 0,              /* ずらし量（定規の単位） / offset (ruler units) */
        scale: 100,                /* スケール（%） / scale */
        rotateEnabled: true,       /* 回転 / rotate */
        rotate: -12,               /* 回転角（°） / angle */
        backgroundEnabled: true,   /* 背景色 / background */
        backgroundHex: "#000000",  /* 背景色の HEX / background HEX */
        maskEnabled: true,         /* マスク / mask */
        marginPt: 20,              /* マージン / margin */
        maskRoundEnabled: false,   /* マスクの角丸 / round mask corners */
        maskRoundRadius: 20        /* マスクの角丸の半径（定規の単位） / mask corner radius (ruler units) */
    };

    /* ［リセット］で DEFAULT_SETTINGS から変える値。横方向の位置は左右の余白がそろうよう計算し直す
       Values [Reset] changes from DEFAULT_SETTINGS; the horizontal offset is recomputed to center the grid */
    var RESET_OVERRIDES = {
        columns: 4,
        evenShiftEnabled: false,
        rotateEnabled: false
    };

    /* 列数の上限（入力欄・スライダー共通） / maximum number of columns */
    var MAX_COLUMNS = 30;

    /* 見開きとみなす横長比（幅 > 高さ × この値） / aspect ratio that marks a page as a spread (width > height x this) */
    var SPREAD_ASPECT_RATIO = 1.2;

    /* 綴じ方向の判定で読む PDF の行数 / number of PDF lines scanned for the binding direction */
    var BINDING_SCAN_LINE_LIMIT = 200;

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

    var OFFSET_LABEL_WIDTH = 140;              /* 位置調整の項目名の幅 / offset label width */
    var UNIT_LABEL_WIDTH   = 24;               /* 単位の幅 / unit label width */
    var SLIDER_WIDTH       = 140;              /* スライダーの幅 / slider width */
    var SWATCH_SIZE        = 24;               /* 背景色の色見本の大きさ / swatch size */
    var MASK_ROW_SPACING   = 10;               /* マスクとマージンの間隔 / spacing between Mask and Margin */

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
            title: { ja: "Slide Collage for Illustrator", en: "Slide Collage for Illustrator" },
            chooseFile: { ja: "PDF/AIを選択してください", en: "Select a PDF or AI file" }
        },
        panel: {
            source: { ja: "読み込みファイル", en: "Source File" },
            artboards: { ja: "アートボードの読み込み", en: "Load Artboards" },
            item: { ja: "アイテム", en: "Items" },
            grid: { ja: "グリッド", en: "Grid" },
            evenColumns: { ja: "偶数列", en: "Even Columns" },
            layout: { ja: "レイアウト", en: "Layout" },
            artboardMask: { ja: "アートボードとマスク", en: "Artboard & Mask" }
        },
        fieldLabel: {
            range: { ja: "範囲", en: "Range" },
            count: { ja: "総数", en: "Total" },
            evenPage: { ja: "偶数ページ", en: "Even Pages" },
            direction: { ja: "方向", en: "Flow" },
            columns: { ja: "列数", en: "Columns" },
            spacing: { ja: "間隔", en: "Gap" },
            scale: { ja: "スケール", en: "Scale" },
            margin: { ja: "マージン", en: "Margin" }
        },
        radio: {
            horizontal: { ja: "横", en: "Horizontal" },
            vertical: { ja: "縦", en: "Vertical" },
            random: { ja: "ランダム", en: "Random" },
            evenPageRight: { ja: "右", en: "Right" },
            evenPageLeft: { ja: "左", en: "Left" }
        },
        checkbox: {
            roundCorners: { ja: "角丸", en: "Round Corners" },
            splitSpreads: { ja: "見開きを左右に分割", en: "Split Spreads" },
            evenPlusSlot: { ja: "＋1スロット", en: "+1 Slot" },
            evenShift: { ja: "ずらし", en: "Shift" },
            rotate: { ja: "回転", en: "Rotate" },
            offsetX: { ja: "横方向の位置調整", en: "Offset X" },
            offsetY: { ja: "縦方向の位置調整", en: "Offset Y" },
            background: { ja: "背景色", en: "Background" },
            mask: { ja: "マスク", en: "Mask" },
            maskRound: { ja: "マスク角丸", en: "Round Mask Corners" }
        },
        dropdown: {
            cropArt: { ja: "アート", en: "Art" },
            cropCrop: { ja: "トリミング", en: "Crop" },
            cropTrim: { ja: "仕上がり", en: "Trim" },
            cropBleed: { ja: "裁ち落とし", en: "Bleed" }
        },
        button: {
            chooseFile: { ja: "ファイルを選択...", en: "Choose File..." },
            load: { ja: "読み込み", en: "Load" },
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        fallbackName: {
            noFile: { ja: "未指定", en: "Not selected" }
        },
        alert: {
            needDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
            needFile: { ja: "先に［ファイルを選択...］で読み込みファイルを選んでください。", en: "Choose a source file first with [Choose File...]." },
            pickPdfAi: { ja: "PDFまたはAIファイルを選択してください。", en: "Select a PDF or AI file." },
            pageCountFailed: { ja: "ファイルのページ数を取得できませんでした。", en: "Could not read the page count of the file." }
        },
        tooltip: {
            range: {
                ja: "読み込むアートボード（PDFはページ）の番号。例：1-20、1,3,5",
                en: "Artboard (or PDF page) numbers to load, e.g. 1-20 or 1,3,5"
            },
            count: {
                ja: "配置する個数。範囲より多いときは範囲を繰り返します",
                en: "Number of items to place; the range repeats when this is larger"
            },
            load: {
                ja: "範囲のアートボードを配置してプレビューします",
                en: "Place the artboards in the range and preview the layout"
            },
            crop: {
                ja: "PDFを配置するときの範囲（AIファイルでは選べません）",
                en: "Box used when placing a PDF (not available for AI files)"
            },
            roundCorners: {
                ja: "各アイテムの角を丸めます。プレビューされず、OKのときに適用します",
                en: "Rounds the corners of each item. Not shown in the preview; applied on OK"
            },
            maskRound: {
                ja: "マスクの角を丸めます。プレビューされず、OKのときに適用します",
                en: "Rounds the corners of the mask. Not shown in the preview; applied on OK"
            },
            splitSpreads: {
                ja: "横長のページを見開きとみなし、左右2つに切り分けて並べます",
                en: "Treats landscape pages as spreads and lays them out as two halves"
            },
            evenPage: {
                ja: "見開きを分割したとき、偶数ページを置く側です。PDF の綴じ方向から自動で設定します",
                en: "Which side the even pages go to when a spread is split. Detected from the PDF binding direction"
            },
            evenPlusSlot: {
                ja: "偶数列にスロットを1つ追加します。空きが出ることがあります",
                en: "Adds one extra slot to even columns. Empty spaces may appear"
            },
            evenShift: { ja: "偶数列だけを上下にずらします（＋で下へ）", en: "Moves even columns only (positive values move down)" },
            scale: {
                ja: "アートボードに収まるよう自動で合わせた大きさに対する倍率",
                en: "Multiplier on top of the size fitted to the artboard"
            },
            rotate: {
                ja: "配置したアイテム全体を回転し、アートボードの中央に合わせます",
                en: "Rotates the whole layout and centers it on the artboard"
            },
            offsetX: { ja: "全体を左右に動かします（＋で右へ）", en: "Moves the whole layout sideways (positive values move right)" },
            offsetY: { ja: "全体を上下に動かします（＋で下へ）", en: "Moves the whole layout up or down (positive values move down)" },
            backgroundHex: {
                ja: "背景色のHEX（例 #000000）。左の色見本をクリックするとカラーピッカーが開きます",
                en: "Background color as HEX (e.g. #000000). Click the swatch to open the color picker"
            },
            mask: {
                ja: "OKのとき、マージンの内側でクリッピングします（背景は含めません）",
                en: "On OK, clips the items inside the margin (the background is not clipped)"
            },
            fitView: {
                ja: "アクティブなアートボードが収まるよう表示倍率を合わせます。キャンセルすると元の表示に戻ります",
                en: "Zooms so the active artboard fits in the window. Cancel restores the original view"
            },
            reset: {
                ja: "ファイルと範囲以外の設定を初期値に戻します",
                en: "Restores the settings other than the file and range"
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
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

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

    /**
     * ∧∨と入力欄を隙間なく並べて追加する。↑↓キーも∧∨と同じ処理で増減する
     * @param {Group} parentGroup - 追加先の行
     * @param {string} initialText - 初期値
     * @param {number} characters - 入力欄の文字数
     * @param {Object} stepOptions - addStepper() に渡す min / max / integer / onStep
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addStepperEditText(parentGroup, initialText, characters, stepOptions) {
        var stepperInputGroup = parentGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = characters;
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 入力欄の有効／無効を、左の∧∨ごと切り替える
     * @param {EditText} numberInput - addStepperEditText() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperEditEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup); /* ∧∨は自作描画なので描き直す / redraw the custom-drawn buttons */
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

    /**
     * 小数第2位で丸める
     * @param {number} value - 数値
     * @returns {number} 丸めた値
     */
    function roundTo2(value) {
        return Math.round(value * 100) / 100;
    }

    // =========================================
    // ページ番号 / Page numbers
    // =========================================

    /**
     * 全角数字を半角にする
     * @param {string} text - 文字列
     * @returns {string} 半角数字にした文字列
     */
    function toHalfWidthDigits(text) {
        return String(text || "").replace(/[０-９]/g, function (ch) {
            return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
        });
    }

    /**
     * 文字列の最初の正の整数を返す（全角数字・前後の空白・単位などの後ろの文字は許す）
     * @param {string} text - 文字列
     * @returns {number} 正の整数。見つからなければ 0
     */
    function parsePositiveInt(text) {
        var digitMatch = toHalfWidthDigits(text).match(/(\d+)/);
        if (!digitMatch) return 0;
        var value = parseInt(digitMatch[1], 10);
        return (value > 0) ? value : 0;
    }

    /**
     * 範囲の文字列（例 "1-20"、"1,3,5"）をページ番号の配列にする。全角数字・読点・前後の括弧も受け付ける
     * @param {string} rangeText - 範囲の文字列
     * @returns {number[]} ページ番号の配列（書かれた順、重複はそのまま）
     */
    function parsePageNumbers(rangeText) {
        var pageNumbers = [];
        var normalized = toHalfWidthDigits(rangeText).replace(/\s+/g, "");
        normalized = normalized.replace(/^[\[\(\{【［]/, "").replace(/[\]\)\}】］]$/, "");
        var rangeParts = normalized.split(/[,、]/);

        for (var i = 0; i < rangeParts.length; i++) {
            var bounds = rangeParts[i].split("-");
            var startPage = parseInt(bounds[0], 10);
            if (isNaN(startPage)) continue;
            var endPage = (bounds.length > 1) ? parseInt(bounds[1], 10) : startPage;
            if (isNaN(endPage)) endPage = startPage;
            for (var j = Math.min(startPage, endPage); j <= Math.max(startPage, endPage); j++) {
                pageNumbers.push(j);
            }
        }
        return pageNumbers;
    }

    /**
     * 配置するページ番号の並びを作る。ファイルにないページは除き、総数に届くまで範囲を繰り返す
     * @param {string} rangeText - 範囲の文字列
     * @param {number} totalCount - 総数（0 なら範囲のまま）
     * @param {number} sourcePageCount - ファイルのページ数（0 なら除外しない）
     * @returns {number[]} ページ番号の配列（最低1つ）
     */
    function buildTargetPages(rangeText, totalCount, sourcePageCount) {
        var rangePages = parsePageNumbers(rangeText);
        var validPages = [];
        for (var i = 0; i < rangePages.length; i++) {
            if (rangePages[i] < 1) continue;
            if (sourcePageCount > 0 && rangePages[i] > sourcePageCount) continue;
            validPages.push(rangePages[i]);
        }
        if (validPages.length === 0) validPages = [1];
        if (!(totalCount > 0)) return validPages;

        var targetPages = [];
        for (var j = 0; j < totalCount; j++) targetPages.push(validPages[j % validPages.length]);
        return targetPages;
    }

    // =========================================
    // 読み込みファイル / Source file
    // =========================================

    /**
     * 拡張子が .pdf / .ai かどうか
     * @param {File} sourceFile - ファイル
     * @returns {boolean} PDF または AI なら true
     */
    function isPdfOrAiFile(sourceFile) {
        return /\.(?:pdf|ai)$/i.test(decodeURIComponent(sourceFile.name));
    }

    /**
     * 拡張子が .pdf かどうか（配置範囲を選べるのは PDF だけ）
     * @param {File} sourceFile - ファイル（null 可）
     * @returns {boolean} PDF なら true
     */
    function isPdfFile(sourceFile) {
        return !!sourceFile && /\.pdf$/i.test(decodeURIComponent(sourceFile.name));
    }

    /**
     * 綴じ方向を判定する。先頭の BINDING_SCAN_LINE_LIMIT 行に /Direction /R2L があれば右綴じとみなす
     * @param {File} sourceFile - ファイル
     * @returns {boolean} 右綴じ（偶数ページが右）なら true。読めないときは false
     */
    function isRightBoundFile(sourceFile) {
        if (!sourceFile.open("r")) return false;
        try {
            for (var i = 0; i < BINDING_SCAN_LINE_LIMIT && !sourceFile.eof; i++) {
                if (/\/Direction\s*\/R2L/.test(sourceFile.readln())) return true;
            }
        } catch (e) {
            return false; /* 読めないファイルは左綴じとみなす / treat unreadable files as left-bound */
        } finally {
            sourceFile.close();
        }
        return false;
    }

    /**
     * 選択の中から最初の配置画像を探す（グループの中もたどる）
     * @param {Object} pageItems - 選択やグループの pageItems
     * @returns {PlacedItem|null} 見つかった配置画像
     */
    function findFirstPlacedItem(pageItems) {
        /* 文字の選択中は selection が TextRange になり、配列ではない / a text selection is a TextRange, not an array */
        if (!pageItems || pageItems.typename === "TextRange") return null;
        for (var i = 0; i < pageItems.length; i++) {
            if (!pageItems[i]) continue;
            if (pageItems[i].typename === "PlacedItem") return pageItems[i];
            if (pageItems[i].typename === "GroupItem") {
                var nestedItem = findFirstPlacedItem(pageItems[i].pageItems);
                if (nestedItem) return nestedItem;
            }
        }
        return null;
    }

    /**
     * PDF / AI ファイルの中身を読んで総ページ数を推定する（ドキュメントとして開かない）
     * @param {File} sourceFile - ファイル
     * @returns {number} ページ数。読めなければ 0
     */
    function readPageCountFromFile(sourceFile) {
        var countPattern = /<<\/Count\s(\d+)/;
        var linearizedPattern = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
        var pagePatterns = [/<<\/Type\/Page\/Parent/, /\/Type\s\/Page\s/, /\/StructParents\s\d+.*\/Type\/Page>>/];
        var pagesTreePattern = /\/Type\/Pages/;
        var pagesCountPattern = /\/Count\s(\d+)/;

        var pageCount = 0;
        var countedPages = 0;
        if (!sourceFile.open("r")) return 0;
        try {
            while (!sourceFile.eof) {
                var line = sourceFile.readln();
                if (countPattern.test(line) || linearizedPattern.test(line)) {
                    pageCount = Number(RegExp.$1);
                    break;
                }
                for (var i = 0; i < pagePatterns.length; i++) {
                    if (pagePatterns[i].test(line)) { countedPages++; break; }
                }
                if (pagesTreePattern.test(line)) {
                    line = sourceFile.readln();
                    if (pagesCountPattern.test(line)) countedPages = Math.max(countedPages, Number(RegExp.$1));
                }
            }
        } catch (e) {
            return 0; /* 読めないファイルはページ数不明として扱う / treat unreadable files as unknown */
        } finally {
            sourceFile.close();
        }
        if (countedPages > 0) pageCount = countedPages;
        return (pageCount > 0) ? pageCount : 0;
    }

    /**
     * キャッシュのキーにするファイルの識別子を返す。#targetengine でキャッシュが Illustrator の終了まで残るので、
     * 更新日時も含めて、ファイルを書き換えたら古い値を使わないようにする
     * @param {File} sourceFile - ファイル
     * @returns {string} パスと更新日時をつないだ文字列
     */
    function getSourceCacheKey(sourceFile) {
        return sourceFile.fsName + "|" + (sourceFile.modified ? sourceFile.modified.getTime() : "");
    }

    /**
     * ファイルをドキュメントとして開いてアートボード数を数える（中身から読めなかったときの予備。セッション中は覚えておく）。
     * すでに開いているファイルは閉じない。数え終えたら元のドキュメントを前面に戻す
     * @param {File} sourceFile - ファイル
     * @param {Document} returnDoc - 数え終えたら前面に戻すドキュメント
     * @returns {number} アートボード数。開けなければ 0
     */
    function countArtboardsByOpening(sourceFile, returnDoc) {
        if (!$.global.SlideCollage_artboardCountCache) $.global.SlideCollage_artboardCountCache = {};
        var countCache = $.global.SlideCollage_artboardCountCache;
        var cacheKey = getSourceCacheKey(sourceFile);
        if (countCache[cacheKey] > 0) return countCache[cacheKey];

        var artboardCount = 0;
        var openedDoc = null;
        /* 開いているかはパスでなく枚数の増減で判定する（日本語パスの比較は外れる） / detect by the document count, not the path */
        var documentCountBefore = app.documents.length;
        try {
            openedDoc = app.open(sourceFile);
            artboardCount = openedDoc.artboards.length;
        } catch (e) {
            /* 開けないファイルは 0 のまま / leave 0 when the file cannot be opened */
        } finally {
            /* 新しく開いたときだけ閉じる。すでに開いていたドキュメントは未保存の編集ごと残す / close only what this opened */
            if (openedDoc && app.documents.length > documentCountBefore) openedDoc.close(SaveOptions.DONOTSAVECHANGES);
            returnDoc.activate();
        }
        if (artboardCount > 0) countCache[cacheKey] = artboardCount;
        return artboardCount;
    }

    // =========================================
    // 配置 / Placement
    // =========================================

    /* plugin/PDFImport/CropTo の値（実測）。ドロップダウンの並び（アート／トリミング／仕上がり／裁ち落とし）と同じ
       Values of plugin/PDFImport/CropTo (measured), in the dropdown's order; 4 is the media box */
    var CROP_TO_VALUES = [0, 1, 2, 3];

    /* AI も PDF と同じ読み込み経路で配置されるので、ページ番号・配置範囲は PDFImport の設定で指定する
       AI files go through the PDF import pipeline too, so both use the PDFImport preferences */
    var PDF_PAGE_NUMBER_PREF = "plugin/PDFImport/PageNumber";
    var PDF_CROP_TO_PREF = "plugin/PDFImport/CropTo";

    /* プレビュー用グループの名前（前回の残りを見つけて消すための目印） / name marking the preview group */
    var PREVIEW_GROUP_NAME = "__SlideCollage_preview__";

    /**
     * ファイルの指定ページを配置する。ページ番号の設定は配置後に1へ戻す
     * @param {Document} doc - 配置先のドキュメント
     * @param {File} sourceFile - 配置するファイル
     * @param {number} pageNumber - ページ（アートボード）番号
     * @param {number} cropToValue - 配置範囲（CROP_TO_VALUES の値）
     * @returns {PlacedItem|null} 配置した画像。失敗したら null
     */
    function placeSourcePage(doc, sourceFile, pageNumber, cropToValue) {
        var placedItem = null;
        app.preferences.setIntegerPreference(PDF_CROP_TO_PREF, cropToValue);
        app.preferences.setIntegerPreference(PDF_PAGE_NUMBER_PREF, pageNumber);
        try {
            placedItem = doc.placedItems.add();
            placedItem.file = sourceFile;
        } catch (e) {
            /* 配置できないページは飛ばす / skip pages that cannot be placed */
            if (placedItem) placedItem.remove();
            placedItem = null;
        } finally {
            app.preferences.setIntegerPreference(PDF_PAGE_NUMBER_PREF, 1);
        }
        return placedItem;
    }

    /**
     * ページを配置したときの大きさを測る（一時的に配置して消す。セッション中は覚えておく）
     * @param {Document} doc - ドキュメント
     * @param {File} sourceFile - ファイル
     * @param {number} pageNumber - ページ番号
     * @param {number} cropToValue - 配置範囲
     * @returns {{width: number, height: number}|null} 大きさ（pt）。測れなければ null
     */
    function measureSourcePage(doc, sourceFile, pageNumber, cropToValue) {
        if (!$.global.SlideCollage_pageSizeCache) $.global.SlideCollage_pageSizeCache = {};
        var sizeCache = $.global.SlideCollage_pageSizeCache;
        var cacheKey = getSourceCacheKey(sourceFile) + "|" + cropToValue + "|" + pageNumber;
        if (sizeCache[cacheKey]) return sizeCache[cacheKey];

        var tempItem = placeSourcePage(doc, sourceFile, pageNumber, cropToValue);
        if (!tempItem) return null;
        var pageSize = { width: tempItem.width, height: tempItem.height };
        tempItem.remove();
        if (!(pageSize.width > 0 && pageSize.height > 0)) return null;
        sizeCache[cacheKey] = pageSize;
        return pageSize;
    }

    /**
     * 前回のプレビューが残っていたら消す（途中で止まったときの後始末）
     * @param {Document} doc - ドキュメント
     * @returns {void}
     */
    function removeLeftoverPreviewGroups(doc) {
        for (var i = doc.groupItems.length - 1; i >= 0; i--) {
            if (doc.groupItems[i].name === PREVIEW_GROUP_NAME) doc.groupItems[i].remove();
        }
    }

    // =========================================
    // グリッド計算 / Grid math
    // =========================================

    /* 方向 / Flow modes */
    var FLOW_HORIZONTAL = 0;
    var FLOW_VERTICAL = 1;
    var FLOW_RANDOM = 2;

    /**
     * ［＋1スロット］のときの列ごとのスロット数を返す（余りを左の列から配り、偶数列に1つ足す。合計が個数を超えた分は空きになる）
     * @param {number} itemCount - アイテム数
     * @param {number} columnCount - 列数
     * @returns {number[]|null} 列ごとのスロット数。1列なら null（通常の配置）
     */
    function getEvenPlusSlotCounts(itemCount, columnCount) {
        if (columnCount < 2) return null;
        var baseCount = Math.floor(itemCount / columnCount);
        var remainder = itemCount - baseCount * columnCount;
        var slotCounts = [];
        for (var i = 0; i < columnCount; i++) {
            /* 0始まりの奇数番目が、見た目の偶数列（2, 4, 6…） / 0-based odd indexes are the visible even columns */
            slotCounts.push(baseCount + (i < remainder ? 1 : 0) + (i % 2 === 1 ? 1 : 0));
        }
        return slotCounts;
    }

    /**
     * 並び順の番号から、グリッドの列・行を求める
     * @param {number} index - 並び順の番号
     * @param {number} itemCount - アイテム数
     * @param {number} columnCount - 列数
     * @param {number} flowMode - FLOW_HORIZONTAL は行優先、それ以外は列優先
     * @param {number[]|null} slotCounts - ［＋1スロット］の列ごとのスロット数（方向より優先）
     * @returns {{col: number, row: number}} 列と行（0始まり）
     */
    function getGridCell(index, itemCount, columnCount, flowMode, slotCounts) {
        if (slotCounts) {
            var columnStart = 0;
            for (var i = 0; i < slotCounts.length; i++) {
                if (index < columnStart + slotCounts[i]) return { col: i, row: index - columnStart };
                columnStart += slotCounts[i];
            }
            return { col: slotCounts.length - 1, row: 0 };
        }
        if (flowMode === FLOW_HORIZONTAL) {
            return { col: index % columnCount, row: Math.floor(index / columnCount) };
        }
        var rowCount = Math.ceil(itemCount / columnCount);
        return { col: Math.floor(index / rowCount), row: index % rowCount };
    }

    /**
     * グリッド全体がアートボードの内側に収まる倍率（%）を求める
     * @param {Object} fitParams - itemWidth / itemHeight（1つ分の大きさ）/ itemCount / columnCount / gapPt /
     *     innerWidth / innerHeight（マージンを除いた大きさ）/ slotCounts / evenShiftPt / rotateDeg
     * @returns {number} 倍率（整数の%）。求められなければ 100
     */
    function calcAutoFitPercent(fitParams) {
        if (!(fitParams.itemWidth > 0 && fitParams.itemHeight > 0)) return 100;
        var rowCount = Math.ceil(fitParams.itemCount / fitParams.columnCount);
        if (fitParams.slotCounts) rowCount = Math.max.apply(null, fitParams.slotCounts);

        var gridWidth = fitParams.columnCount * fitParams.itemWidth + (fitParams.columnCount - 1) * fitParams.gapPt;
        var gridHeight = rowCount * fitParams.itemHeight + (rowCount - 1) * fitParams.gapPt + Math.abs(fitParams.evenShiftPt);

        /* 回転するときは、グリッド全体を回した外接サイズで見積もる / use the bounding box of the rotated grid */
        if (fitParams.rotateDeg !== 0) {
            var radians = Math.abs(fitParams.rotateDeg) * Math.PI / 180;
            var sinValue = Math.sin(radians);
            var cosValue = Math.cos(radians);
            var rotatedWidth = gridWidth * cosValue + gridHeight * sinValue;
            gridHeight = gridWidth * sinValue + gridHeight * cosValue;
            gridWidth = rotatedWidth;
        }
        if (!(gridWidth > 0 && gridHeight > 0)) return 100;

        var fitPercent = Math.round(Math.min(fitParams.innerWidth / gridWidth, fitParams.innerHeight / gridHeight) * 100);
        return (fitPercent > 0) ? fitPercent : 100;
    }

    /**
     * 0〜count-1 の番号をシャッフルした配列を返す
     * @param {number} count - 個数
     * @returns {number[]} シャッフルした番号
     */
    function shuffledIndexes(count) {
        var indexes = [];
        for (var i = 0; i < count; i++) indexes.push(i);
        for (var j = indexes.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var swapValue = indexes[j];
            indexes[j] = indexes[k];
            indexes[k] = swapValue;
        }
        return indexes;
    }

    // =========================================
    // アートボードと図形 / Artboard and shapes
    // =========================================

    /**
     * アクティブなアートボードから、マージンを除いた内側の矩形を返す
     * @param {Document} doc - ドキュメント
     * @param {number} marginPt - マージン（pt）
     * @returns {{left: number, top: number, width: number, height: number}} 内側の矩形（幅・高さは最低1）
     */
    function getInnerArtboardRect(doc, marginPt) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; /* [left, top, right, bottom] */
        return {
            left: artboardRect[0] + marginPt,
            top: artboardRect[1] - marginPt,
            width: Math.max(1, Math.abs(artboardRect[2] - artboardRect[0]) - marginPt * 2),
            height: Math.max(1, Math.abs(artboardRect[1] - artboardRect[3]) - marginPt * 2)
        };
    }

    /**
     * アートボードと同じ大きさの背景を最背面に描く
     * @param {Document} doc - ドキュメント
     * @param {RGBColor} fillColor - 塗り
     * @returns {PathItem} 背景の長方形
     */
    function drawArtboardBackground(doc, fillColor) {
        var artboardRect = getInnerArtboardRect(doc, 0);
        var backgroundRect = doc.activeLayer.pathItems.rectangle(artboardRect.top, artboardRect.left, artboardRect.width, artboardRect.height);
        backgroundRect.stroked = false;
        backgroundRect.filled = true;
        backgroundRect.fillColor = fillColor;
        backgroundRect.zOrder(ZOrderMethod.SENDTOBACK);
        return backgroundRect;
    }

    /**
     * 矩形のマスクでアイテムをクリップグループにまとめる
     * @param {Document} doc - ドキュメント
     * @param {PageItem[]} itemsToClip - クリップするアイテム
     * @param {{left: number, top: number, width: number, height: number}} clipRect - マスクの矩形
     * @returns {GroupItem} クリップグループ
     */
    function clipItemsToRect(doc, itemsToClip, clipRect) {
        var maskPath = doc.activeLayer.pathItems.rectangle(clipRect.top, clipRect.left, clipRect.width, clipRect.height);
        maskPath.stroked = false;
        maskPath.filled = false;

        var clipGroup = doc.groupItems.add();
        /* 見た目の重なり順を保つよう、背面から順に末尾へ入れる / keep the stacking order */
        for (var i = itemsToClip.length - 1; i >= 0; i--) itemsToClip[i].moveToEnd(clipGroup);
        maskPath.moveToBeginning(clipGroup); /* マスクは最前面 / the mask goes on top */
        clipGroup.clipped = true;
        return clipGroup;
    }

    /**
     * アイテムを同じ大きさの矩形でクリップグループにする（角丸をクリップグループにかけるため）
     * @param {Document} doc - ドキュメント
     * @param {PageItem} item - アイテム
     * @returns {GroupItem} クリップグループ
     */
    function wrapWithClipGroup(doc, item) {
        var bounds = item.geometricBounds; /* [left, top, right, bottom] */
        return clipItemsToRect(doc, [item], {
            left: bounds[0],
            top: bounds[1],
            width: Math.max(1, Math.abs(bounds[2] - bounds[0])),
            height: Math.max(1, Math.abs(bounds[1] - bounds[3]))
        });
    }

    /**
     * 大きさが見開き（横長）かどうか
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @returns {boolean} 見開きなら true
     */
    function isSpreadSize(width, height) {
        return width > height * SPREAD_ASPECT_RATIO;
    }

    /**
     * 見開きの配置画像を左右の半ページ2つに分ける。複製した画像をそれぞれ半分の矩形でクリップする
     * @param {Document} doc - ドキュメント
     * @param {PlacedItem} spreadItem - 見開きの配置画像
     * @param {boolean} evenPageOnRight - 偶数ページが右なら true（右半分を先に並べる）
     * @returns {GroupItem[]} 並べる順の半ページ（クリップグループ）2つ
     */
    function splitSpreadItem(doc, spreadItem, evenPageOnRight) {
        var bounds = spreadItem.geometricBounds; /* [left, top, right, bottom] */
        var halfWidth = Math.abs(bounds[2] - bounds[0]) / 2;
        var height = Math.abs(bounds[1] - bounds[3]);
        var rightItem = spreadItem.duplicate();
        var leftHalf = clipItemsToRect(doc, [spreadItem], { left: bounds[0], top: bounds[1], width: halfWidth, height: height });
        var rightHalf = clipItemsToRect(doc, [rightItem], { left: bounds[0] + halfWidth, top: bounds[1], width: halfWidth, height: height });
        return evenPageOnRight ? [rightHalf, leftHalf] : [leftHalf, rightHalf];
    }

    /**
     * 角丸のライブエフェクトをかける
     * @param {PageItem} item - アイテム
     * @param {number} radiusPt - 半径（pt）
     * @returns {void}
     */
    function applyRoundCorners(item, radiusPt) {
        item.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>');
    }

    /**
     * 見えている範囲の境界を返す。クリップグループの width・geometricBounds はマスクの外に隠れた部分も含むので、
     * クリップグループはマスク（pageItems[0]）の範囲、通常のグループは中身の見えている範囲を合わせた範囲にする
     * A clip group's width and geometricBounds include the masked-out content, so measure its mask (pageItems[0])
     * @param {PageItem} item - アイテム
     * @returns {number[]} [left, top, right, bottom]
     */
    function getVisibleBounds(item) {
        if (item.typename !== "GroupItem") return item.geometricBounds;
        if (item.clipped) return item.pageItems[0].geometricBounds;
        var unionBounds = null;
        for (var i = 0; i < item.pageItems.length; i++) {
            var childBounds = getVisibleBounds(item.pageItems[i]);
            if (!unionBounds) {
                unionBounds = childBounds.slice(0);
                continue;
            }
            unionBounds[0] = Math.min(unionBounds[0], childBounds[0]);
            unionBounds[1] = Math.max(unionBounds[1], childBounds[1]);
            unionBounds[2] = Math.max(unionBounds[2], childBounds[2]);
            unionBounds[3] = Math.min(unionBounds[3], childBounds[3]);
        }
        return unionBounds || item.geometricBounds;
    }

    /**
     * アイテムの見えている範囲が、指定した枠にぴったり収まるよう拡大縮小して移動する
     * @param {PageItem} item - アイテム
     * @param {number} frameLeft - 枠の左端（pt）
     * @param {number} frameTop - 枠の上端（pt）
     * @param {number} frameWidth - 枠の幅（pt）
     * @param {number} frameHeight - 枠の高さ（pt）
     * @returns {void}
     */
    function fitItemToFrame(item, frameLeft, frameTop, frameWidth, frameHeight) {
        var currentBounds = getVisibleBounds(item);
        item.resize(frameWidth / (currentBounds[2] - currentBounds[0]) * 100, frameHeight / (currentBounds[1] - currentBounds[3]) * 100);
        var resizedBounds = getVisibleBounds(item);
        item.translate(frameLeft - resizedBounds[0], frameTop - resizedBounds[1]);
    }

    /**
     * アイテムの中心をアクティブなアートボードの中心に合わせる
     * @param {Document} doc - ドキュメント
     * @param {PageItem} item - アイテム
     * @returns {void}
     */
    function moveCenterToArtboardCenter(doc, item) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var itemBounds = getVisibleBounds(item);
        var dx = (artboardRect[0] + artboardRect[2]) / 2 - (itemBounds[0] + itemBounds[2]) / 2;
        var dy = (artboardRect[1] + artboardRect[3]) / 2 - (itemBounds[1] + itemBounds[3]) / 2;
        item.translate(dx, dy);
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
    // ダイアログの部品 / Dialog parts
    // =========================================

    /**
     * パネルを追加して、余白と並びをそろえる
     * @param {Group} parent - 追加先
     * @param {string} title - パネル名
     * @returns {Panel} パネル
     */
    function addPanel(parent, title) {
        var newPanel = parent.add("panel", undefined, title);
        setupPanel(newPanel);
        /* ボタンや入力欄は広げない / keep buttons and fields at their natural width */
        newPanel.alignChildren = ["left", "top"];
        return newPanel;
    }

    /**
     * 横並びの行を追加する
     * @param {Group|Panel} parent - 追加先
     * @returns {Group} 行
     */
    function addRow(parent) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * 余りの幅を吸うスペーサーを追加する（右端のスライダーの位置をそろえる）
     * @param {Group} rowGroup - 行
     * @returns {void}
     */
    function addRowSpacer(rowGroup) {
        var spacer = rowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
    }

    /**
     * 「項目名またはチェックボックス・∧∨・入力欄・単位・（スライダー）」の数値行を追加する。
     * 入力・∧∨・↑↓キー・スライダーの値を連動させ、値が変わったら onValueChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {Object} rowOptions - label（コロン込みの項目名）または checkboxLabel / tooltip /
     *     characters / min / max / integer / unit / fallback（数値でないときの値）/ slider（{ min, max }）
     * @returns {Object} 数値行（group / leadControl / checkbox / input / slider と getValue / setValue / isChecked / setChecked / onValueChange）
     */
    function addNumberRow(parent, rowOptions) {
        var numberRow = { onValueChange: null, checkbox: null, slider: null };
        var lastNotifiedState = null;
        numberRow.group = addRow(parent);

        var leadControl;
        if (rowOptions.checkboxLabel) {
            numberRow.checkbox = numberRow.group.add("checkbox", undefined, rowOptions.checkboxLabel);
            leadControl = numberRow.checkbox;
        } else {
            leadControl = numberRow.group.add("statictext", undefined, rowOptions.label);
        }
        numberRow.leadControl = leadControl;
        if (rowOptions.tooltip) leadControl.helpTip = rowOptions.tooltip;

        numberRow.input = addStepperEditText(numberRow.group, "", rowOptions.characters, {
            min: rowOptions.min, max: rowOptions.max, integer: rowOptions.integer,
            onStep: function () { writeValue(numberRow.getValue()); notifyChange(); }
        });
        var unitLabel = numberRow.group.add("statictext", undefined, rowOptions.unit || "");
        if (rowOptions.slider) {
            unitLabel.preferredSize.width = UNIT_LABEL_WIDTH;
            addRowSpacer(numberRow.group);
            numberRow.slider = numberRow.group.add("slider", undefined, rowOptions.slider.min, rowOptions.slider.min, rowOptions.slider.max);
            numberRow.slider.preferredSize.width = SLIDER_WIDTH;
        }

        /**
         * 値を下限・上限に収め、整数または小数第2位に丸める
         * @param {number} value - 数値
         * @returns {number} そろえた値
         */
        function normalizeValue(value) {
            if (isNaN(value)) value = rowOptions.fallback || 0;
            if (rowOptions.min !== undefined) value = Math.max(rowOptions.min, value);
            if (rowOptions.max !== undefined) value = Math.min(rowOptions.max, value);
            return rowOptions.integer ? Math.round(value) : roundTo2(value);
        }

        /**
         * 入力欄とスライダーに値を書き込む
         * @param {number} value - 数値
         * @returns {void}
         */
        function writeValue(value) {
            value = normalizeValue(value);
            numberRow.input.text = String(value);
            if (numberRow.slider) numberRow.slider.value = value;
        }

        /**
         * 今の値とチェックの状態を1つの文字列にする
         * @returns {string} 状態
         */
        function getRowState() {
            return numberRow.isChecked() + "|" + numberRow.getValue();
        }

        /**
         * 値かチェックの状態が前回から変わっていれば onValueChange を呼ぶ（同じ値での再描画を避ける）
         * @returns {void}
         */
        function notifyChange() {
            var currentState = getRowState();
            if (currentState === lastNotifiedState) return;
            lastNotifiedState = currentState;
            if (numberRow.onValueChange) numberRow.onValueChange();
        }

        numberRow.getValue = function () {
            return normalizeValue(parseFloat(numberRow.input.text));
        };
        numberRow.isChecked = function () {
            return numberRow.checkbox ? numberRow.checkbox.value : true;
        };
        /* 外から値を入れるときは呼び出し側がプレビューを更新するので、今の状態を通知済みとして控える
           Callers refresh the preview themselves, so record the state as already notified */
        numberRow.setValue = function (value) {
            writeValue(value);
            lastNotifiedState = getRowState();
        };
        numberRow.setChecked = function (isChecked) {
            numberRow.checkbox.value = isChecked;
            updateEnabledState();
            lastNotifiedState = getRowState();
        };

        /**
         * チェックに合わせて入力欄とスライダーの有効／無効を切り替える
         * @returns {void}
         */
        function updateEnabledState() {
            var isEnabled = numberRow.isChecked();
            setStepperEditEnabled(numberRow.input, isEnabled);
            if (numberRow.slider) numberRow.slider.enabled = isEnabled;
        }

        /* 入力中はテキストを書き換えず、スライダーとプレビューだけ追従させる / follow while typing without rewriting the text */
        numberRow.input.onChanging = function () {
            var typedValue = parseFloat(numberRow.input.text);
            if (isNaN(typedValue)) return;
            if (numberRow.slider) numberRow.slider.value = normalizeValue(typedValue);
            notifyChange();
        };
        numberRow.input.onChange = function () {
            writeValue(numberRow.getValue());
            notifyChange();
        };
        if (numberRow.slider) {
            numberRow.slider.onChanging = numberRow.slider.onChange = function () {
                writeValue(numberRow.slider.value);
                notifyChange();
            };
        }
        if (numberRow.checkbox) {
            numberRow.checkbox.onClick = function () {
                updateEnabledState();
                notifyChange();
            };
        }
        return numberRow;
    }

    /**
     * 数値行の頭（項目名・チェックボックス）の幅を、実際の描画幅でいちばん広いものにそろえる。
     * 固定幅だと「スケール：」のように長い項目名の末尾（コロン）が切れるため、実測してそろえる
     * @param {Object[]} numberRows - addNumberRow() の戻り値の配列
     * @returns {void}
     */
    function alignNumberRowLeads(numberRows) {
        var maxLeadWidth = 0;
        for (var i = 0; i < numberRows.length; i++) {
            /* 作成時に文字列から決まる自然な幅 / natural width set from the text on creation */
            maxLeadWidth = Math.max(maxLeadWidth, numberRows[i].leadControl.preferredSize.width || 0);
        }
        if (!(maxLeadWidth > 0)) return; /* 測れないときは自動の幅のまま / keep automatic widths if unmeasured */
        for (var j = 0; j < numberRows.length; j++) numberRows[j].leadControl.preferredSize.width = maxLeadWidth;
    }

    /**
     * 「チェックボックス・スライダー」の位置調整の行を追加する（数値は出さない）
     * @param {Panel} parent - 追加先
     * @param {string} checkboxLabel - チェックボックスの文言
     * @param {string} tooltip - ヒント
     * @param {number} range - スライダーの範囲（±、定規の単位）
     * @returns {Object} 位置調整の行（checkbox / slider と setChecked / onValueChange）
     */
    function addOffsetRow(parent, checkboxLabel, tooltip, range) {
        var offsetRow = { onValueChange: null };
        var rowGroup = addRow(parent);
        offsetRow.checkbox = rowGroup.add("checkbox", undefined, checkboxLabel);
        offsetRow.checkbox.preferredSize.width = OFFSET_LABEL_WIDTH;
        offsetRow.checkbox.helpTip = tooltip;
        addRowSpacer(rowGroup);
        offsetRow.slider = rowGroup.add("slider", undefined, 0, -range, range);
        offsetRow.slider.preferredSize.width = SLIDER_WIDTH;

        /**
         * 値が変わったことを知らせる
         * @returns {void}
         */
        function notifyChange() {
            if (offsetRow.onValueChange) offsetRow.onValueChange();
        }

        offsetRow.setChecked = function (isChecked) {
            offsetRow.checkbox.value = isChecked;
            offsetRow.slider.enabled = isChecked;
        };
        offsetRow.checkbox.onClick = function () {
            offsetRow.slider.enabled = offsetRow.checkbox.value;
            notifyChange();
        };
        offsetRow.slider.onChanging = offsetRow.slider.onChange = notifyChange;
        return offsetRow;
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
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを作る（イベントの配線は main 側）
     * @param {Document} doc - ドキュメント
     * @param {Object} rulerUnit - getUnitInfo() の戻り値
     * @returns {Object} ダイアログとコントロール一式
     */
    function buildDialog(doc, rulerUnit) {
        var dialogControls = {};
        dialogControls.dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogControls.dialog);

        var columnsGroup = dialogControls.dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "fill"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";
        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";

        /* 読み込みファイル / Source file */
        var sourcePanel = addPanel(leftColumn, getLabel("panel.source"));
        dialogControls.btnChooseFile = sourcePanel.add("button", undefined, getLabel("button.chooseFile"));
        dialogControls.sourceNameText = sourcePanel.add("statictext", undefined, getLabel("fallbackName.noFile"));
        dialogControls.sourceNameText.characters = 20;

        /* アートボードの読み込み（範囲・総数・読み込み） / Load artboards */
        var artboardsPanel = addPanel(leftColumn, getLabel("panel.artboards"));
        var rangeRowGroup = addRow(artboardsPanel);
        rangeRowGroup.add("statictext", undefined, labelText("fieldLabel.range")).helpTip = getLabel("tooltip.range");
        dialogControls.rangeInput = rangeRowGroup.add("edittext", undefined, "");
        dialogControls.rangeInput.characters = 10;
        dialogControls.rangeInput.helpTip = getLabel("tooltip.range");

        var countRowGroup = addRow(artboardsPanel);
        countRowGroup.add("statictext", undefined, labelText("fieldLabel.count")).helpTip = getLabel("tooltip.count");
        /* onStep は main 側で入れる（手で増減したら範囲との連動を切る） / onStep is set in main */
        dialogControls.countStepOptions = { integer: true, min: 1 };
        dialogControls.countInput = addStepperEditText(countRowGroup, "", 4, dialogControls.countStepOptions);
        dialogControls.countInput.helpTip = getLabel("tooltip.count");

        dialogControls.btnLoad = artboardsPanel.add("button", undefined, getLabel("button.load"));
        dialogControls.btnLoad.helpTip = getLabel("tooltip.load");

        /* アイテム / Items */
        var itemPanel = addPanel(leftColumn, getLabel("panel.item"));
        dialogControls.cropDropdown = itemPanel.add("dropdownlist", undefined, [
            getLabel("dropdown.cropArt"), getLabel("dropdown.cropCrop"), getLabel("dropdown.cropTrim"), getLabel("dropdown.cropBleed")
        ]);
        dialogControls.cropDropdown.minimumSize.width = 160;
        dialogControls.cropDropdown.helpTip = getLabel("tooltip.crop");
        dialogControls.roundRow = addNumberRow(itemPanel, {
            checkboxLabel: labelText("checkbox.roundCorners"), tooltip: getLabel("tooltip.roundCorners"),
            characters: 3, min: 0, unit: rulerUnit.label
        });

        dialogControls.splitSpreadsCheckbox = itemPanel.add("checkbox", undefined, getLabel("checkbox.splitSpreads"));
        dialogControls.splitSpreadsCheckbox.helpTip = getLabel("tooltip.splitSpreads");
        dialogControls.evenPageGroup = addRow(itemPanel);
        dialogControls.evenPageGroup.add("statictext", undefined, labelText("fieldLabel.evenPage")).helpTip = getLabel("tooltip.evenPage");
        dialogControls.evenPageRightRadio = dialogControls.evenPageGroup.add("radiobutton", undefined, getLabel("radio.evenPageRight"));
        dialogControls.evenPageLeftRadio = dialogControls.evenPageGroup.add("radiobutton", undefined, getLabel("radio.evenPageLeft"));
        dialogControls.evenPageRightRadio.helpTip = dialogControls.evenPageLeftRadio.helpTip = getLabel("tooltip.evenPage");
        dialogControls.evenPageRightRadio.value = true;

        /* グリッド / Grid */
        var gridPanel = addPanel(rightColumn, getLabel("panel.grid"));
        var directionRowGroup = addRow(gridPanel);
        directionRowGroup.add("statictext", undefined, labelText("fieldLabel.direction"));
        dialogControls.flowRadios = [
            directionRowGroup.add("radiobutton", undefined, getLabel("radio.horizontal")), /* FLOW_HORIZONTAL */
            directionRowGroup.add("radiobutton", undefined, getLabel("radio.vertical")),   /* FLOW_VERTICAL */
            directionRowGroup.add("radiobutton", undefined, getLabel("radio.random"))      /* FLOW_RANDOM */
        ];
        dialogControls.columnsRow = addNumberRow(gridPanel, {
            label: labelText("fieldLabel.columns"), characters: 4,
            min: 1, max: MAX_COLUMNS, integer: true, fallback: 1, unit: "", slider: { min: 1, max: MAX_COLUMNS }
        });
        dialogControls.spacingRow = addNumberRow(gridPanel, {
            label: labelText("fieldLabel.spacing"), characters: 4,
            min: 0, max: 100, unit: rulerUnit.label, slider: { min: 0, max: 100 }
        });

        /* 偶数列 / Even columns */
        var evenColumnsPanel = addPanel(rightColumn, getLabel("panel.evenColumns"));
        dialogControls.evenPlusCheckbox = evenColumnsPanel.add("checkbox", undefined, getLabel("checkbox.evenPlusSlot"));
        dialogControls.evenPlusCheckbox.helpTip = getLabel("tooltip.evenPlusSlot");
        dialogControls.evenShiftRow = addNumberRow(evenColumnsPanel, {
            checkboxLabel: labelText("checkbox.evenShift"), tooltip: getLabel("tooltip.evenShift"),
            characters: 4, min: -200, max: 200, unit: rulerUnit.label, slider: { min: -200, max: 200 }
        });

        /* レイアウト / Layout */
        var layoutPanel = addPanel(rightColumn, getLabel("panel.layout"));
        dialogControls.scaleRow = addNumberRow(layoutPanel, {
            label: labelText("fieldLabel.scale"), tooltip: getLabel("tooltip.scale"),
            characters: 4, min: 10, max: 250, integer: true, fallback: 100, unit: "%", slider: { min: 10, max: 250 }
        });
        dialogControls.rotateRow = addNumberRow(layoutPanel, {
            checkboxLabel: labelText("checkbox.rotate"), tooltip: getLabel("tooltip.rotate"),
            characters: 4, min: -30, max: 30, integer: true, unit: "°", slider: { min: -30, max: 30 }
        });
        /* スライダーのある行は、頭の幅をそろえて入力欄の位置を合わせる / align the leads of the slider rows */
        alignNumberRowLeads([dialogControls.columnsRow, dialogControls.spacingRow, dialogControls.evenShiftRow, dialogControls.scaleRow, dialogControls.rotateRow]);

        /* 位置調整のスライダーは、アートボードの幅・高さ（定規の単位）まで動かせる / offset range follows the artboard size */
        var artboardRect = getInnerArtboardRect(doc, 0);
        dialogControls.offsetXRow = addOffsetRow(layoutPanel, labelText("checkbox.offsetX"), getLabel("tooltip.offsetX"), artboardRect.width / rulerUnit.pointsPerUnit);
        dialogControls.offsetYRow = addOffsetRow(layoutPanel, labelText("checkbox.offsetY"), getLabel("tooltip.offsetY"), artboardRect.height / rulerUnit.pointsPerUnit);

        /* アートボードとマスク / Artboard & mask */
        var artboardMaskPanel = addPanel(rightColumn, getLabel("panel.artboardMask"));
        var backgroundRowGroup = addRow(artboardMaskPanel);
        dialogControls.backgroundCheckbox = backgroundRowGroup.add("checkbox", undefined, labelText("checkbox.background"));
        dialogControls.backgroundSwatch = backgroundRowGroup.add("panel");
        dialogControls.backgroundSwatch.preferredSize = [SWATCH_SIZE, SWATCH_SIZE];
        dialogControls.backgroundSwatch.helpTip = getLabel("tooltip.backgroundHex");
        dialogControls.backgroundHexInput = backgroundRowGroup.add("edittext", undefined, "");
        dialogControls.backgroundHexInput.characters = 7;
        dialogControls.backgroundHexInput.helpTip = getLabel("tooltip.backgroundHex");

        /* マスクとマージンは同じ行、マスク角丸はマージンの左端にそろえて次の行 / mask and margin share a row; the mask corner row starts under the margin */
        var maskRowGroup = addRow(artboardMaskPanel);
        maskRowGroup.spacing = MASK_ROW_SPACING; /* 字下げの計算に使うので明示する / set explicitly for the indent below */
        dialogControls.maskCheckbox = maskRowGroup.add("checkbox", undefined, getLabel("checkbox.mask"));
        dialogControls.maskCheckbox.helpTip = getLabel("tooltip.mask");
        dialogControls.marginRow = addNumberRow(maskRowGroup, {
            label: labelText("fieldLabel.margin"), characters: 5, min: 0, unit: rulerUnit.label
        });
        dialogControls.marginRow.group.margins = 0;
        dialogControls.maskRoundRow = addNumberRow(artboardMaskPanel, {
            checkboxLabel: labelText("checkbox.maskRound"), tooltip: getLabel("tooltip.maskRound"),
            characters: 3, min: 0, unit: rulerUnit.label
        });
        dialogControls.maskRoundRow.group.margins = [(dialogControls.maskCheckbox.preferredSize.width || 0) + maskRowGroup.spacing, 0, 0, 0];

        /* ボタン行（左：リセット・画面にフィット／右：キャンセル・OK） / Button row */
        var buttonRow = addButtonRow(dialogControls.dialog);
        dialogControls.btnReset = buttonRow.leftGroup.add("button", undefined, getLabel("button.reset"));
        dialogControls.btnReset.helpTip = getLabel("tooltip.reset");
        dialogControls.fitViewControls = FitViewToItems.addControls(buttonRow.leftGroup, { lang: uiLang });
        /* 部品の既定の説明は「作成するオブジェクト」向けなので、アートボードに合わせる説明に差し替える / describe the artboard fit */
        dialogControls.fitViewControls.checkbox.helpTip = getLabel("tooltip.fitView");
        dialogControls.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogControls.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        return dialogControls;
    }

    /**
     * ダイアログで設定し、選んだファイルのアートボードをグリッドに並べる
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.needDocument"));
            return;
        }
        var doc = app.activeDocument;
        var rulerUnit = getUnitInfo("rulerType");

        var sourceFile = null;          /* 読み込みファイル / source file */
        var sourcePageCount = 0;        /* ファイルのページ数（中身から読めたとき） / page count read from the file */
        var isCountLinkedToRange = true; /* 総数を範囲に合わせるか（手入力したら切る） / whether Total follows the range */
        var isApplyingSettings = false;  /* 設定をまとめて反映している最中か / true while applySettings() runs */

        /* 読み込んだプレビューの状態。items は並べる前の順で、baseWidths / baseHeights は配置したままの大きさ
           Loaded preview state; items keep their load order, base sizes are as placed */
        var previewState = {
            items: [], baseWidths: [], baseHeights: [],
            group: null, background: null, rotation: 0, randomOrder: null, roundRadiusPt: 0,
            cropIndex: -1, splitSpreads: false, evenPageOnRight: true
        };

        var dialogControls = buildDialog(doc, rulerUnit);
        var initialViewState = FitViewToItems.captureView(doc); /* キャンセルで戻す表示 / view restored on Cancel */

        /**
         * 定規の単位の値を pt にする
         * @param {number} value - 定規の単位の値
         * @returns {number} pt
         */
        function toPoints(value) {
            return value * rulerUnit.pointsPerUnit;
        }

        /**
         * pt を定規の単位の値にする
         * @param {number} valuePt - pt
         * @returns {number} 定規の単位の値
         */
        function fromPoints(valuePt) {
            return valuePt / rulerUnit.pointsPerUnit;
        }

        // -----------------------------------------
        // UI の値 / UI values
        // -----------------------------------------

        /**
         * 選んでいる方向を返す
         * @returns {number} FLOW_HORIZONTAL / FLOW_VERTICAL / FLOW_RANDOM
         */
        function getFlowMode() {
            for (var i = 0; i < dialogControls.flowRadios.length; i++) {
                if (dialogControls.flowRadios[i].value) return i;
            }
            return FLOW_VERTICAL;
        }

        /**
         * 選んでいる配置範囲の CropTo の値を返す
         * @returns {number} CROP_TO_VALUES の値
         */
        function getCropToValue() {
            return CROP_TO_VALUES[dialogControls.cropDropdown.selection ? dialogControls.cropDropdown.selection.index : DEFAULT_SETTINGS.cropIndex];
        }

        /**
         * マージンを pt で返す
         * @returns {number} マージン（pt）
         */
        function getMarginPt() {
            return toPoints(dialogControls.marginRow.getValue());
        }

        /**
         * ファイルのページ数を返す（中身から読めなければ、開いてアートボードを数える）
         * @returns {number} ページ数。分からなければ 0
         */
        function getSourcePageCount() {
            if (sourcePageCount > 0) return sourcePageCount;
            return sourceFile ? countArtboardsByOpening(sourceFile, doc) : 0;
        }

        /**
         * 範囲・総数・ファイルのページ数から、配置するページ番号の並びを返す
         * @returns {number[]} ページ番号の配列
         */
        function getTargetPages() {
            return buildTargetPages(dialogControls.rangeInput.text, parsePositiveInt(dialogControls.countInput.text), getSourcePageCount());
        }

        /**
         * グリッドの1つ分の大きさ（pt、拡大縮小前）を返す。読み込み済みならその1つ目、未読み込みなら一時的に配置して測る
         * @returns {{width: number, height: number}|null} 大きさ。測れなければ null
         */
        function getReferenceItemSize() {
            if (previewState.items.length > 0) return { width: previewState.baseWidths[0], height: previewState.baseHeights[0] };
            if (!sourceFile) return null;
            var pageSize = measureSourcePage(doc, sourceFile, getTargetPages()[0], getCropToValue());
            if (pageSize && dialogControls.splitSpreadsCheckbox.value && isSpreadSize(pageSize.width, pageSize.height)) {
                return { width: pageSize.width / 2, height: pageSize.height };
            }
            return pageSize;
        }

        /**
         * 今の設定でアートボードに収まる倍率（%）を返す（スケールの値は含まない）
         * @param {number} itemCount - 並べる数
         * @param {{width: number, height: number}} itemSize - 1つ分の大きさ（pt）
         * @param {number} [evenShiftPtOverride] - ずらし量（pt）をこの値として見積もる（省略時は UI の値）
         * @returns {number} 倍率（%）
         */
        function calcFitPercentForUI(itemCount, itemSize, evenShiftPtOverride) {
            var columnCount = dialogControls.columnsRow.getValue();
            var innerRect = getInnerArtboardRect(doc, getMarginPt());
            return calcAutoFitPercent({
                itemWidth: itemSize.width, itemHeight: itemSize.height,
                itemCount: itemCount, columnCount: columnCount,
                gapPt: toPoints(dialogControls.spacingRow.getValue()),
                innerWidth: innerRect.width, innerHeight: innerRect.height,
                slotCounts: dialogControls.evenPlusCheckbox.value ? getEvenPlusSlotCounts(itemCount, columnCount) : null,
                evenShiftPt: (evenShiftPtOverride !== undefined) ? evenShiftPtOverride :
                    (dialogControls.evenShiftRow.isChecked() ? toPoints(dialogControls.evenShiftRow.getValue()) : 0),
                rotateDeg: dialogControls.rotateRow.isChecked() ? dialogControls.rotateRow.getValue() : 0
            });
        }

        // -----------------------------------------
        // 読み込みファイル / Source file
        // -----------------------------------------

        /**
         * 読み込みファイルを設定し、ファイル名・配置範囲・綴じ方向・範囲・総数を更新する
         * @param {File} newSourceFile - ファイル
         * @returns {void}
         */
        function setSourceFile(newSourceFile) {
            sourceFile = newSourceFile;
            dialogControls.sourceNameText.text = decodeURIComponent(sourceFile.name);
            dialogControls.sourceNameText.helpTip = decodeURIComponent(sourceFile.fsName);
            dialogControls.cropDropdown.enabled = isPdfFile(sourceFile);
            /* 綴じ方向から偶数ページの位置を設定する / set the even-page side from the binding direction */
            dialogControls.evenPageRightRadio.value = isRightBoundFile(sourceFile);
            dialogControls.evenPageLeftRadio.value = !dialogControls.evenPageRightRadio.value;

            sourcePageCount = readPageCountFromFile(sourceFile);
            if (sourcePageCount > 0) {
                dialogControls.rangeInput.text = "1-" + sourcePageCount;
                isCountLinkedToRange = true;
                syncCountToRange();
            } else {
                alert(getLabel("alert.pageCountFailed"));
                dialogControls.rangeInput.text = "";
                dialogControls.countInput.text = "";
            }
        }

        /**
         * 総数を範囲のページ数に合わせる（総数を手入力したあとは合わせない）
         * @returns {void}
         */
        function syncCountToRange() {
            if (!isCountLinkedToRange) return;
            var pageCount = parsePageNumbers(dialogControls.rangeInput.text).length;
            dialogControls.countInput.text = (pageCount > 0) ? String(pageCount) : "";
        }

        /**
         * 選択している配置画像が PDF / AI なら、それを読み込みファイルにする
         * @returns {void}
         */
        function useSelectedPlacedFile() {
            var placedItem = findFirstPlacedItem(doc.selection);
            if (!placedItem) return;
            var linkedFile;
            try {
                linkedFile = placedItem.file;
            } catch (e) {
                return; /* リンク切れの画像は file の読み取りで例外になる / a missing link throws on .file */
            }
            if (linkedFile && isPdfOrAiFile(linkedFile)) setSourceFile(linkedFile);
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * プレビューのアイテムと背景を消して、状態を空に戻す
         * @returns {void}
         */
        function clearPreview() {
            if (previewState.group) previewState.group.remove(); /* 中のアイテムもまとめて消える / removes its children too */
            if (previewState.background) previewState.background.remove();
            previewState.items = [];
            previewState.baseWidths = [];
            previewState.baseHeights = [];
            previewState.group = null;
            previewState.background = null;
            previewState.rotation = 0;
            previewState.randomOrder = null;
            previewState.roundRadiusPt = 0;
        }

        /**
         * 読み込みファイルの対象ページを配置してプレビューを作り直す
         * @returns {boolean} 読み込めたら true
         */
        function loadPreview() {
            if (!sourceFile) {
                alert(getLabel("alert.needFile"));
                return false;
            }
            clearPreview();
            removeLeftoverPreviewGroups(doc);

            var targetPages = getTargetPages();
            if (!dialogControls.countInput.text) dialogControls.countInput.text = String(targetPages.length);
            var cropToValue = getCropToValue();
            previewState.cropIndex = dialogControls.cropDropdown.selection.index;
            previewState.splitSpreads = dialogControls.splitSpreadsCheckbox.value;
            previewState.evenPageOnRight = dialogControls.evenPageRightRadio.value;

            /* 回転・後始末をまとめて扱えるよう、1つのグループに入れる / keep everything in one group for rotation and cleanup */
            previewState.group = doc.groupItems.add();
            previewState.group.name = PREVIEW_GROUP_NAME;
            for (var i = 0; i < targetPages.length; i++) {
                var placedItem = placeSourcePage(doc, sourceFile, targetPages[i], cropToValue);
                if (!placedItem) continue;
                var placedPieces = [placedItem];
                if (previewState.splitSpreads && isSpreadSize(placedItem.width, placedItem.height)) {
                    placedPieces = splitSpreadItem(doc, placedItem, previewState.evenPageOnRight);
                }
                for (var j = 0; j < placedPieces.length; j++) {
                    placedPieces[j].moveToEnd(previewState.group);
                    previewState.items.push(placedPieces[j]);
                    var pieceBounds = getVisibleBounds(placedPieces[j]);
                    previewState.baseWidths.push(pieceBounds[2] - pieceBounds[0]);
                    previewState.baseHeights.push(pieceBounds[1] - pieceBounds[3]);
                }
            }
            updateBackgroundPreview();
            applyLayout(false);
            fitViewToArtboard();
            return true;
        }

        /**
         * ランダムのときの並び順を返す（読み込み直すか方向を選び直すまで同じ順を保つ）
         * @param {number} flowMode - 方向
         * @returns {number[]|null} 並び順。ランダムでなければ null
         */
        function getPlacementOrder(flowMode) {
            if (flowMode !== FLOW_RANDOM) {
                previewState.randomOrder = null;
                return null;
            }
            if (!previewState.randomOrder || previewState.randomOrder.length !== previewState.items.length) {
                previewState.randomOrder = shuffledIndexes(previewState.items.length);
            }
            return previewState.randomOrder;
        }

        /**
         * 今の設定でプレビューのアイテムを並べ直す（大きさ・位置・角丸・回転・位置調整）
         * @param {boolean} withRoundCorners - 角丸もかけるなら true（OK のとき。プレビューではかけない）
         * @returns {void}
         */
        function applyLayout(withRoundCorners) {
            var itemCount = previewState.items.length;
            if (itemCount === 0) return;

            var innerRect = getInnerArtboardRect(doc, getMarginPt());
            var columnCount = dialogControls.columnsRow.getValue();
            var gapPt = toPoints(dialogControls.spacingRow.getValue());
            var evenShiftPt = dialogControls.evenShiftRow.isChecked() ? toPoints(dialogControls.evenShiftRow.getValue()) : 0;
            var rotateDeg = dialogControls.rotateRow.isChecked() ? dialogControls.rotateRow.getValue() : 0;
            var offsetXPt = dialogControls.offsetXRow.checkbox.value ? toPoints(dialogControls.offsetXRow.slider.value) : 0;
            var offsetYPt = dialogControls.offsetYRow.checkbox.value ? toPoints(dialogControls.offsetYRow.slider.value) : 0; /* ＋で下へ / positive moves down */
            var flowMode = getFlowMode();
            var placementOrder = getPlacementOrder(flowMode);
            var slotCounts = dialogControls.evenPlusCheckbox.value ? getEvenPlusSlotCounts(itemCount, columnCount) : null;
            var finalPercent = calcFitPercentForUI(itemCount, { width: previewState.baseWidths[0], height: previewState.baseHeights[0] }) * dialogControls.scaleRow.getValue() / 100;

            /* 前回の回転を戻してから並べる / undo the previous rotation before laying out */
            if (previewState.rotation !== 0) {
                previewState.group.rotate(-previewState.rotation);
                previewState.rotation = 0;
            }

            /* 回転するときは回したあとで中央に合わせてから位置調整するので、ここでは足さない
               When rotating, the offset is applied after centering, not here */
            var startX = innerRect.left + (rotateDeg !== 0 ? 0 : offsetXPt);
            var startY = innerRect.top - (rotateDeg !== 0 ? 0 : offsetYPt);
            for (var i = 0; i < itemCount; i++) {
                var itemIndex = placementOrder ? placementOrder[i] : i;
                var layoutItem = previewState.items[itemIndex];
                /* 元の大きさから毎回計算し、拡大縮小が積み重ならないようにする / size from the base each time */
                var cellWidth = previewState.baseWidths[itemIndex] * finalPercent / 100;
                var cellHeight = previewState.baseHeights[itemIndex] * finalPercent / 100;
                var gridCell = getGridCell(i, itemCount, columnCount, flowMode, slotCounts);
                var cellTop = startY - gridCell.row * (cellHeight + gapPt);
                if (gridCell.col % 2 === 1) cellTop -= evenShiftPt; /* 偶数列だけずらす / offset even columns */
                fitItemToFrame(layoutItem, startX + gridCell.col * (cellWidth + gapPt), cellTop, cellWidth, cellHeight);
            }

            if (withRoundCorners) applyItemRoundCorners();

            if (rotateDeg !== 0) {
                previewState.group.rotate(rotateDeg);
                previewState.rotation = rotateDeg;
                moveCenterToArtboardCenter(doc, previewState.group);
                if (offsetXPt !== 0 || offsetYPt !== 0) previewState.group.translate(offsetXPt, -offsetYPt);
            }
            app.redraw();
        }

        /**
         * アイテムに角丸をかける。配置画像は同じ大きさのクリップグループに入れ、角丸はクリップグループにかける
         * @returns {void}
         */
        function applyItemRoundCorners() {
            if (!dialogControls.roundRow.isChecked()) return;
            var radiusPt = toPoints(dialogControls.roundRow.getValue());
            if (!(radiusPt > 0)) return;

            for (var i = 0; i < previewState.items.length; i++) {
                if (previewState.items[i].typename === "GroupItem") continue; /* 包み済み・見開きの半ページ / already a clip group */
                var clipGroup = wrapWithClipGroup(doc, previewState.items[i]);
                clipGroup.moveToEnd(previewState.group);
                previewState.items[i] = clipGroup; /* 番号を保ち、元の大きさ・並び順と対応させる / keep the index */
            }
            if (previewState.roundRadiusPt === radiusPt) return;
            for (var j = 0; j < previewState.items.length; j++) applyRoundCorners(previewState.items[j], radiusPt);
            previewState.roundRadiusPt = radiusPt;
        }

        /**
         * 読み込み済みならプレビューを並べ直す
         * @returns {void}
         */
        function refreshPreview() {
            if (previewState.items.length > 0) applyLayout(false);
        }

        /**
         * 配置範囲や見開きの分割など、配置し直しが要る設定が変わっていたら読み込み直し、それ以外は並べ直す
         * @returns {void}
         */
        function reloadOrRefreshPreview() {
            if (isApplyingSettings || previewState.items.length === 0) return;
            if (previewNeedsReload()) loadPreview();
            else applyLayout(false);
        }

        /**
         * 読み込んだときから、配置し直しが要る設定（配置範囲・見開きの分割・偶数ページの側）が変わったかを返す
         * @returns {boolean} 読み込み直しが要るなら true
         */
        function previewNeedsReload() {
            return previewState.cropIndex !== dialogControls.cropDropdown.selection.index ||
                previewState.splitSpreads !== dialogControls.splitSpreadsCheckbox.value ||
                (previewState.splitSpreads && previewState.evenPageOnRight !== dialogControls.evenPageRightRadio.value);
        }

        /**
         * ［画面にフィット］がオンなら、アクティブなアートボードが収まるよう表示を合わせる。
         * 並べたアイテムはアートボードに合わせて配置され、半ページのクリップグループは隠れた部分も境界に含むため、アートボードを基準にする
         * @returns {void}
         */
        function fitViewToArtboard() {
            if (!dialogControls.fitViewControls.checkbox.value) return;
            var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            /* fit() は geometricBounds だけを読むので、アートボードの矩形を渡す / fit() only reads geometricBounds */
            FitViewToItems.fit([{ geometricBounds: artboardRect }], { doc: doc, fillRatio: dialogControls.fitViewControls.getFillRatio() });
            app.redraw();
        }

        /**
         * 背景色の色見本を塗り直し、背景の有効／無効を切り替える
         * @returns {void}
         */
        function updateBackgroundControls() {
            var isEnabled = dialogControls.backgroundCheckbox.value;
            dialogControls.backgroundSwatch.enabled = isEnabled;
            dialogControls.backgroundHexInput.enabled = isEnabled;
            var swatchColor = getBackgroundColor();
            var swatchGraphics = dialogControls.backgroundSwatch.graphics;
            swatchGraphics.backgroundColor = swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR,
                [swatchColor.red / 255, swatchColor.green / 255, swatchColor.blue / 255, 1]);
        }

        /**
         * 背景色を返す（HEX が正しくなければ黒）
         * @returns {RGBColor} 背景色
         */
        function getBackgroundColor() {
            return parseHexColor(dialogControls.backgroundHexInput.text) || makeRGBColor(0, 0, 0);
        }

        /**
         * プレビューの背景だけを作る・塗り直す・消す（並べ直さない）
         * @returns {void}
         */
        function updateBackgroundPreview() {
            if (previewState.items.length === 0) return;
            if (!dialogControls.backgroundCheckbox.value) {
                if (previewState.background) previewState.background.remove();
                previewState.background = null;
            } else if (!previewState.background) {
                previewState.background = drawArtboardBackground(doc, getBackgroundColor());
            } else {
                previewState.background.fillColor = getBackgroundColor();
            }
            placeBackgroundBehind(previewState.group);
            app.redraw();
        }

        /**
         * 背景を、指定したアイテムのすぐ背面に置く（最背面送りだけでは、あとから作るグループとの前後が保証されない）
         * @param {PageItem} frontItem - 背景より前に見せるアイテム
         * @returns {void}
         */
        function placeBackgroundBehind(frontItem) {
            if (previewState.background && frontItem) previewState.background.move(frontItem, ElementPlacement.PLACEAFTER);
        }

        /**
         * カラーピッカーで背景色を選ぶ
         * @returns {void}
         */
        function pickBackgroundColor() {
            /* 選んだ色は戻り値で返り、型はドキュメントのカラーモードに従う（渡した色は書き換わらない）
               The picked color is returned in the document's color model; the argument is left unchanged */
            var pickedColor = toRGBColor(app.showColorPicker(getBackgroundColor()));
            if (!pickedColor) return;
            dialogControls.backgroundHexInput.text = toHexColor(pickedColor);
            updateBackgroundControls();
            updateBackgroundPreview();
        }

        /**
         * マスクに合わせて、マージンとマスク角丸の有効／無効を切り替える
         * @returns {void}
         */
        function updateMaskControls() {
            var isMaskEnabled = dialogControls.maskCheckbox.value;
            dialogControls.marginRow.group.enabled = isMaskEnabled;
            dialogControls.maskRoundRow.group.enabled = isMaskEnabled;
            redrawSteppersIn(dialogControls.marginRow.group);
            redrawSteppersIn(dialogControls.maskRoundRow.group);
        }

        /**
         * 見開きの分割に合わせて、偶数ページの位置の有効／無効を切り替える
         * @returns {void}
         */
        function updateEvenPageControls() {
            dialogControls.evenPageGroup.enabled = dialogControls.splitSpreadsCheckbox.value;
        }

        // -----------------------------------------
        // 設定の反映・リセット / Settings and reset
        // -----------------------------------------

        /**
         * 設定の値をダイアログに書き込む
         * @param {Object} settings - DEFAULT_SETTINGS と同じ形
         * @returns {void}
         */
        function applySettings(settings) {
            /* selection の代入で onChange が走るので、反映し終えるまで読み込み直しを止める / assigning selection fires onChange */
            isApplyingSettings = true;
            dialogControls.cropDropdown.selection = settings.cropIndex;
            dialogControls.roundRow.setChecked(settings.roundEnabled);
            dialogControls.roundRow.setValue(settings.roundRadius);
            dialogControls.splitSpreadsCheckbox.value = settings.splitSpreads;
            updateEvenPageControls();

            dialogControls.flowRadios[settings.flowMode].value = true;
            dialogControls.columnsRow.setValue(settings.columns);
            dialogControls.spacingRow.setValue(roundTo2(fromPoints(settings.spacingPt)));
            dialogControls.evenPlusCheckbox.value = settings.evenPlusSlot;
            dialogControls.evenShiftRow.setChecked(settings.evenShiftEnabled);
            dialogControls.evenShiftRow.setValue(settings.evenShift);

            dialogControls.scaleRow.setValue(settings.scale);
            dialogControls.rotateRow.setChecked(settings.rotateEnabled);
            dialogControls.rotateRow.setValue(settings.rotate);
            dialogControls.offsetXRow.setChecked(false);
            dialogControls.offsetXRow.slider.value = 0;
            dialogControls.offsetYRow.setChecked(false);
            dialogControls.offsetYRow.slider.value = 0;

            dialogControls.backgroundCheckbox.value = settings.backgroundEnabled;
            dialogControls.backgroundHexInput.text = settings.backgroundHex;
            updateBackgroundControls();
            dialogControls.maskCheckbox.value = settings.maskEnabled;
            dialogControls.marginRow.setValue(roundTo2(fromPoints(settings.marginPt)));
            dialogControls.maskRoundRow.setChecked(settings.maskRoundEnabled);
            dialogControls.maskRoundRow.setValue(settings.maskRoundRadius);
            updateMaskControls();
            isApplyingSettings = false;
        }

        /**
         * 左右の余白がそろうよう、横方向の位置調整をオンにして値を入れる（回転なしのグリッドを前提に計算）
         * @returns {void}
         */
        function centerGridHorizontally() {
            var itemSize = getReferenceItemSize();
            if (!itemSize) return;
            var itemCount = (previewState.items.length > 0) ? previewState.items.length : getTargetPages().length;
            var columnCount = dialogControls.columnsRow.getValue();
            var itemWidth = itemSize.width * calcFitPercentForUI(itemCount, itemSize) * dialogControls.scaleRow.getValue() / 10000;
            var gridWidth = columnCount * itemWidth + (columnCount - 1) * toPoints(dialogControls.spacingRow.getValue());
            var offsetValue = fromPoints((getInnerArtboardRect(doc, getMarginPt()).width - gridWidth) / 2);
            if (!isFinite(offsetValue)) return;
            dialogControls.offsetXRow.setChecked(true);
            dialogControls.offsetXRow.slider.value = offsetValue;
        }

        /**
         * ファイルと範囲以外の設定をリセットの値に戻し、プレビューに反映する
         * @returns {void}
         */
        function resetSettings() {
            var resetSettingsValues = {};
            for (var settingKey in DEFAULT_SETTINGS) resetSettingsValues[settingKey] = DEFAULT_SETTINGS[settingKey];
            for (var overrideKey in RESET_OVERRIDES) resetSettingsValues[overrideKey] = RESET_OVERRIDES[overrideKey];
            applySettings(resetSettingsValues);
            /* 中央寄せは読み込み直したあとのアイテムで測る / measure the centering on the reloaded items */
            if (previewState.items.length > 0 && previewNeedsReload()) loadPreview();
            centerGridHorizontally();
            refreshPreview();
        }

        /**
         * 偶数列のずらし量を自動で入れる（1つ分の高さと間隔の和の半分で、半マスずれる）。
         * ずらすと全体が縮んで1つ分の高さも変わるので、ずらし量と倍率が落ち着くまで数回計算し直す
         * @returns {void}
         */
        function setDefaultEvenShift() {
            var itemSize = getReferenceItemSize();
            if (!itemSize) return;
            var itemCount = (previewState.items.length > 0) ? previewState.items.length : getTargetPages().length;
            var gapPt = toPoints(dialogControls.spacingRow.getValue());
            var evenShiftPt = 0;
            for (var i = 0; i < 5; i++) {
                var itemHeight = itemSize.height * calcFitPercentForUI(itemCount, itemSize, evenShiftPt) * dialogControls.scaleRow.getValue() / 10000;
                evenShiftPt = (itemHeight + gapPt) / 2;
            }
            dialogControls.evenShiftRow.setValue(fromPoints(evenShiftPt));
        }

        // -----------------------------------------
        // 確定 / Finalize
        // -----------------------------------------

        /**
         * プレビューを確定する。角丸をかけて並べ直し、マスクがオンならマージンの内側でクリップする
         * @returns {void}
         */
        function finalizePreview() {
            if (previewState.items.length === 0) {
                clearPreview(); /* 1つも配置できなかったときは空のグループを残さない / leave no empty group */
                return;
            }
            applyLayout(true);
            previewState.group.name = "";
            if (!dialogControls.maskCheckbox.value) {
                placeBackgroundBehind(previewState.group);
                return;
            }
            var maskGroup = clipItemsToRect(doc, [previewState.group], getInnerArtboardRect(doc, getMarginPt()));
            var maskRadiusPt = toPoints(dialogControls.maskRoundRow.getValue());
            if (dialogControls.maskRoundRow.isChecked() && maskRadiusPt > 0) applyRoundCorners(maskGroup, maskRadiusPt);
            placeBackgroundBehind(maskGroup);
        }

        // -----------------------------------------
        // イベント / Events
        // -----------------------------------------

        dialogControls.btnChooseFile.onClick = function () {
            var chosenFile = File.openDialog(getLabel("dialog.chooseFile"), "PDF/AI:*.pdf;*.ai");
            if (!chosenFile) return;
            if (!isPdfOrAiFile(chosenFile)) {
                alert(getLabel("alert.pickPdfAi"));
                return;
            }
            setSourceFile(chosenFile);
        };
        dialogControls.rangeInput.onChanging = syncCountToRange;
        dialogControls.countInput.onChanging = dialogControls.countStepOptions.onStep = function () {
            isCountLinkedToRange = false;
        };
        dialogControls.btnLoad.onClick = loadPreview;

        dialogControls.cropDropdown.onChange = reloadOrRefreshPreview;
        dialogControls.splitSpreadsCheckbox.onClick = function () {
            updateEvenPageControls();
            reloadOrRefreshPreview();
        };
        dialogControls.evenPageRightRadio.onClick = dialogControls.evenPageLeftRadio.onClick = reloadOrRefreshPreview;

        var layoutRows = [dialogControls.columnsRow, dialogControls.spacingRow, dialogControls.evenShiftRow, dialogControls.scaleRow, dialogControls.rotateRow, dialogControls.marginRow, dialogControls.offsetXRow, dialogControls.offsetYRow];
        for (var i = 0; i < layoutRows.length; i++) layoutRows[i].onValueChange = refreshPreview;

        /* ずらしをオンにしたら、ずらし量を自動で入れてから並べ直す / auto-fill the offset when Shift is turned on */
        var toggleEvenShiftControls = dialogControls.evenShiftRow.checkbox.onClick;
        dialogControls.evenShiftRow.checkbox.onClick = function () {
            toggleEvenShiftControls();
            if (!dialogControls.evenShiftRow.isChecked()) return;
            setDefaultEvenShift();
            refreshPreview();
        };
        /* 角丸・マスク角丸は OK のときだけかけるので、プレビューは更新しない / round corners apply only on OK */

        for (var j = 0; j < dialogControls.flowRadios.length; j++) {
            dialogControls.flowRadios[j].onClick = function () {
                previewState.randomOrder = null; /* ランダムを選び直したら並びも引き直す / reshuffle on every click */
                refreshPreview();
            };
        }
        dialogControls.evenPlusCheckbox.onClick = refreshPreview;

        dialogControls.backgroundCheckbox.onClick = function () {
            updateBackgroundControls();
            updateBackgroundPreview();
        };
        dialogControls.backgroundHexInput.onChange = function () {
            updateBackgroundControls();
            updateBackgroundPreview();
        };
        dialogControls.backgroundSwatch.addEventListener("mousedown", pickBackgroundColor);
        dialogControls.maskCheckbox.onClick = updateMaskControls;

        dialogControls.btnReset.onClick = resetSettings;
        dialogControls.fitViewControls.checkbox.onClick = function () {
            dialogControls.fitViewControls.updateEnabled();
            fitViewToArtboard();
        };
        dialogControls.fitViewControls.percentInput.onChange = fitViewToArtboard;
        dialogControls.btnOK.onClick = function () {
            /* 読み込んでいなければ、ここで配置してから確定する / load first when nothing has been loaded */
            if (previewState.items.length === 0 && !loadPreview()) return;
            dialogControls.dialog.close(1);
        };

        // -----------------------------------------
        // 初期化と表示 / Initialize and show
        // -----------------------------------------

        applySettings(DEFAULT_SETTINGS);
        dialogControls.cropDropdown.enabled = false;
        useSelectedPlacedFile();
        if (sourceFile) setDefaultEvenShift();

        prepareDialogWindow(dialogControls.dialog, SCRIPT_NAME);
        if (dialogControls.dialog.show() === 1) {
            finalizePreview();
        } else {
            clearPreview();
            FitViewToItems.restoreView(initialViewState, doc);
            app.redraw();
        }
        doc.selection = null; /* ダイアログを閉じたら選択を解除 / clear the selection after closing */
    }

    main();

})();
