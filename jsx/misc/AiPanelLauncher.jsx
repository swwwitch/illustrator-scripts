#target illustrator
#targetengine "AiPanelLauncherEngine"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.AiPanelLauncher
Name=パネルを検索して表示
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-11-05
Description-BEGIN
［ウィンドウ］メニューのパネルを、サブメニュー（［書式］）の中まで展開して1つのリストに並べます。
検索欄で絞り込み、選んだパネルを表示します。

Lists the panels of the Window menu, including those in the Type submenu, in a single list.
Filter it with the search field and show the panels you pick.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

［ウィンドウ］メニューのパネルを、サブメニュー（［書式］）の中まで展開して1つのリストに並べます。
検索欄で絞り込み、選んだパネルを表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiPanelLauncher.md

### Overview

Lists the panels of the Window menu, including those in the Type submenu, in a single list.
Filter it with the search field and show the panels you pick.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiPanelLauncher.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiPanelLauncher";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-06";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "パネルを検索して表示"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiPanelLauncher.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiPanelLauncher.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiPanelLauncher.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［ウィンドウ］メニューのパネル（メニューの並び順）。Illustrator のスクリプトからはメニューを読めないので表に持つ。
       [日本語名, 英語名, 分類（サブメニュー。LABELS.category のキー、直下は空文字）, メニューコマンド ID（空文字は CC ライブラリ）, 対応する最小のバージョン]
       / Panels of the Window menu in menu order. Scripts cannot read Illustrator's menus, so they are listed here.
       [Japanese name, English name, category key, menu command ID, minimum version] */
    var PANEL_TABLE = [
        ["3D とマテリアル", "3D and Materials", "", "Adobe 3D Panel", 26],
        ["CC ライブラリ", "CC Libraries", "", "", 0],
        ["CSS プロパティ (非推奨)", "CSS Properties (Deprecated)", "", "CSS Menu Item", 0],
        ["Retype", "Retype", "", "ReTypeWindowMenu", 27.6],
        ["SVG インタラクティビティ", "SVG Interactivity", "", "Adobe SVG Interactivity Palette", 0],
        ["アクション", "Actions", "", "Adobe Action Palette", 0],
        ["アセットの書き出し", "Asset Export", "", "Adobe SmartExport Panel Menu Item", 0],
        ["アピアランス", "Appearance", "", "Style Palette", 0],
        ["アートボード", "Artboards", "", "Adobe Artboard Palette", 0],
        ["カラー", "Color", "", "Adobe Color Palette", 0],
        ["カラーガイド", "Color Guide", "", "Adobe Harmony Palette", 0],
        ["グラデーション", "Gradient", "", "Adobe Gradient Palette", 0],
        ["グラフィックスタイル", "Graphic Styles", "", "Adobe Style Palette", 0],
        ["コメント", "Comments", "", "Adobe Commenting Palette", 26],
        ["コンセプトからベクター生成", "Concept to Vector", "", "SketchToVectorUnified", 30.5],
        ["シンボル", "Symbols", "", "Adobe Symbol Palette", 0],
        ["スウォッチ", "Swatches", "", "Adobe Swatches Menu Item", 0],
        ["ドキュメント情報", "Document Info", "", "DocInfo1", 0],
        ["ナビゲーター", "Navigator", "", "AdobeNavigator", 0],
        ["バージョン履歴", "Version History", "", "Adobe Version History File Menu Item", 26],
        ["パスファインダー", "Pathfinder", "", "Adobe PathfinderUI", 0],
        ["パターンオプション", "Pattern Options", "", "Adobe Pattern Panel Toggle", 0],
        ["ヒストリー", "History", "", "Adobe History Panel Menu Item", 27],
        ["ブラシ", "Brushes", "", "Adobe BrushManager Menu Item", 0],
        ["プロパティ", "Properties", "", "Adobe Property Palette", 26],
        ["モックアップ", "Mockup", "", "Adobe Vector Edge Panel", 28],
        ["リンク", "Links", "", "Adobe LinkPalette Menu Item", 0],
        ["レイヤー", "Layers", "", "AdobeLayerPalette1", 0],
        ["分割・統合プレビュー", "Flattener Preview", "", "Adobe Flattening Preview", 0],
        ["分版プレビュー", "Separations Preview", "", "Adobe Separation Preview Panel", 0],
        ["変形", "Transform", "", "AdobeTransformObjects1", 0],
        ["変数", "Variables", "", "Adobe Variables Palette Menu Item", 0],
        ["属性", "Attributes", "", "internal palettes posing as plug-in menus-attributes", 0],
        ["情報", "Info", "", "internal palettes posing as plug-in menus-info", 0],
        ["整列", "Align", "", "AdobeAlignObjects2", 0],
        ["OpenType", "OpenType", "type", "internal palettes posing as plug-in menus-opentype", 0],
        ["タブ", "Tabs", "type", "internal palettes posing as plug-in menus-tab", 0],
        ["字形", "Glyphs", "type", "alternate glyph palette plugin 2", 0],
        ["文字", "Character", "type", "internal palettes posing as plug-in menus-character", 0],
        ["文字スタイル", "Character Styles", "type", "Character Styles", 0],
        ["文字組み更新", "Reflow Viewer", "type", "ReflowWindowMenu", 29],
        ["段落", "Paragraph", "type", "internal palettes posing as plug-in menus-paragraph", 0],
        ["段落スタイル", "Paragraph Styles", "type", "Adobe Paragraph Styles Palette", 0],
        ["生成履歴", "Generation History", "", "Adobe Generative Patterns Panel", 28.6],
        ["画像トレース", "Image Trace", "", "Adobe Vectorize Panel", 0],
        ["線", "Stroke", "", "Adobe Stroke Palette", 0],
        ["自動選択", "Magic Wand", "", "AI Magic Wand", 0],
        ["透明", "Transparency", "", "Adobe Transparency Palette Menu Item", 0]
    ];

    /* まとめて表示するパネルのセット（ボタンの順）。panels は PANEL_TABLE の日本語名で、上から順に表示する
       / Panel sets shown together (button order). panels are the Japanese names in PANEL_TABLE, shown top to bottom */
    var PANEL_SETS = [
        {
            label: { ja: "テキスト", en: "Text" },
            panels: ["文字", "段落", "文字スタイル", "段落スタイル", "字形"]
        },
        {
            label: { ja: "アピアランス", en: "Appearance" },
            panels: ["アピアランス", "グラフィックスタイル"]
        },
        {
            label: { ja: "カラー", en: "Color" },
            panels: ["カラー", "スウォッチ", "カラーガイド", "グラデーション"]
        }
    ];

    /* CC ライブラリのメニューコマンド ID は、末尾がメニュー名（UI の言語）になる
       / The CC Libraries command ID ends with the menu name in the UI language */
    var CC_LIBRARIES_COMMAND_PREFIX = "Adobe CSXS Extension com.adobe.DesignLibraries.angular";

    /* 並び順（ラジオボタンの順）。menu＝デフォルト（日本語版の［ウィンドウ］メニューの並び。英語版はメニューもアルファベット順なので名前順）、kana＝あいうえお順
       / Sort orders (radio button order): menu = default (Window menu order; the English menu is alphabetical), kana = alphabetical */
    var SORT_KEYS = ["menu", "kana"];
    var DEFAULT_SORT_KEY = "menu";

    /* あいうえお順と検索で使う、パネル名の漢字の読み（長い語から置き換える）
       / Readings of the kanji in Japanese panel names, used for the alphabetical order */
    var KANJI_READINGS = [
        ["文字組み更新", "モジグミコウシン"], ["画像", "ガゾウ"], ["自動選択", "ジドウセンタク"],
        ["分割", "ブンカツ"], ["統合", "トウゴウ"], ["分版", "ブンパン"], ["変形", "ヘンケイ"],
        ["変数", "ヘンスウ"], ["属性", "ゾクセイ"], ["情報", "ジョウホウ"], ["整列", "セイレツ"],
        ["字形", "ジケイ"], ["文字", "モジ"], ["段落", "ダンラク"], ["生成", "セイセイ"],
        ["履歴", "リレキ"], ["線", "セン"], ["透明", "トウメイ"], ["非推奨", "ヒスイショウ"],
        ["書式", "ショシキ"]
    ];

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

    var LIST_SIZE     = [440, 460];        /* リストの寸法 [幅,高さ] / list size */
    var COLUMN_WIDTHS = [250, 170];        /* 列幅（名前・分類）/ column widths (name, category) */
    var ROW_LABEL_WIDTH = 60;              /* 行ラベル（検索・並び順・セット）の幅 / width of the row labels (search, sort, sets) */

    /* 検索欄のクリアボタン（自前描画、AiScriptLauncher と同じ）/ Clear button next to the search field (custom drawn, same as AiScriptLauncher) */
    var CLEAR_BUTTON_SIZE  = 20;           /* クリアボタンの一辺 / clear button size */
    var CLEAR_CIRCLE_INSET = 2;            /* 円と外周の間隔 / inset of the circle */
    var CLEAR_GLYPH_INSET  = 6;            /* ×と外周の間隔 / inset of the × glyph */
    var CLEAR_STROKE_WIDTH = 1.5;          /* 円と×の線幅 / stroke width */

    // =========================================
    // ラベル定義 / Labels
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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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

    /**
     * クリアボタンの丸と×を描く
     * @param {Button} button - 対象ボタン
     * @returns {void}
     */
    function drawClearButton(button) {
        var graphics = button.graphics;
        var size = button.size[0];
        var useLightUI = !isDarkUI();

        /* 主役ではないので、通常時も少し薄めに描く。押している間は濃くして手応えを出す / Slightly muted, darker while pressed */
        var glyphColor = useLightUI ? [0.45, 0.45, 0.45, 1] : [0.72, 0.72, 0.72, 1];
        if (button.pressed) glyphColor = useLightUI ? [0.20, 0.20, 0.20, 1] : [0.95, 0.95, 0.95, 1];

        /* 検索欄が空のときは無効にしてあるので、さらに淡くする / Dim further while disabled */
        if (button.enabled === false) glyphColor = useLightUI ? [0.78, 0.78, 0.78, 1] : [0.42, 0.42, 0.42, 1];

        /* 前回の描画を消すため、まず親と同じ色で塗りつぶす。背景色のブラシを持たない環境では省く / Repaint the background first; skipped where no background brush is available */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, size, size);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, CLEAR_STROKE_WIDTH);
        var circleSize = size - CLEAR_CIRCLE_INSET * 2;
        graphics.newPath();
        graphics.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
        graphics.strokePath(pen);

        var glyphStart = CLEAR_GLYPH_INSET;
        var glyphEnd = size - CLEAR_GLYPH_INSET;
        graphics.newPath();
        graphics.moveTo(glyphStart, glyphStart);
        graphics.lineTo(glyphEnd, glyphEnd);
        graphics.strokePath(pen);
        graphics.newPath();
        graphics.moveTo(glyphEnd, glyphStart);
        graphics.lineTo(glyphStart, glyphEnd);
        graphics.strokePath(pen);
    }

    /**
     * 検索欄を消すクリアボタンを作る
     * @param {Group} parent - 追加先のグループ
     * @returns {Button} 丸に×を自前描画したボタン
     */
    function addClearButton(parent) {
        var button = parent.add("button", undefined, "");
        button.helpTip = getLabel(LABELS.tooltip.clearSearch);
        button.alignment = ["right", "center"];
        button.preferredSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.minimumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.maximumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.pressed = false;

        /* 実際の状態は呼び出し側が入力欄に合わせる / The caller syncs this with the field */
        button.enabled = false;

        button.onDraw = function () {
            drawClearButton(this);
        };
        button.addEventListener("mousedown", function () {
            this.pressed = true;
            this.notify("onDraw");
        });
        button.addEventListener("mouseup", function () {
            this.pressed = false;
            this.notify("onDraw");
        });
        /* ボタンの外でマウスを離したときも押下表示を戻す / Reset the pressed look on mouseout */
        button.addEventListener("mouseout", function () {
            if (!this.pressed) return;
            this.pressed = false;
            this.notify("onDraw");
        });

        return button;
    }

    var LABELS = {
        dialog: {
            title: { ja: "パネル", en: "Panels" }
        },
        column: {
            name:     { ja: "名前", en: "Name" },
            category: { ja: "分類", en: "Category" }
        },
        fieldLabel: {
            search:    { ja: "検索", en: "Search" },
            sortOrder: { ja: "並び順", en: "Sort by" },
            panelSet:  { ja: "セット", en: "Sets" }
        },
        sort: {
            menu: { ja: "デフォルト", en: "Default" },
            kana: { ja: "あいうえお順", en: "Alphabetical" }
        },
        button: {
            show:   { ja: "表示", en: "Show" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            search: {
                ja: "名前と分類（サブメニュー名）で絞り込みます。空白で区切ると、すべての語を含むものに絞ります。ひらがなはカタカナとして探し、漢字の名前は読みでも探します。↓でリストへ移ります。",
                en: "Filters by name and category (submenu name). Separate words with spaces to match all of them. Press Down to move to the list."
            },
            sortOrder: {
                ja: "リストの並び順を切り替えます。デフォルトは［ウィンドウ］メニューの並び（［書式］の中身は［書式］の位置）、あいうえお順はパネル名の読みの順です。",
                en: "Changes the order of the list. In the English UI both orders are alphabetical, like the Window menu."
            },
            clearSearch: { ja: "検索をクリア", en: "Clear search" },
            panelSet: {
                ja: "%1 をまとめて表示します。開いているパネルは閉じることがあります。",
                en: "Shows %1 together. Panels that are already open may close."
            },
            list: {
                ja: "ダブルクリックか Enter で表示します。⌘（Ctrl）／Shift＋クリックで複数選べます。",
                en: "Double-click or press Enter to show. Cmd (Ctrl) / Shift-click to pick more than one."
            }
        },
        counter: {
            itemUnit: { ja: "件", en: " panels" }
        },
        category: {
            type: { ja: "書式", en: "Type" }
        },
        alert: {
            notFound:   { ja: "この Illustrator で使えるパネルが見つかりませんでした。", en: "No panels are available in this version of Illustrator." },
            showFailed: { ja: "表示できなかったパネル", en: "Could not show" }
        }
    };

    // =========================================
    // データ / Data
    // =========================================

    /**
     * @typedef {Object} PanelEntry
     * @property {string} name - メニューに出る名前（UI の言語）
     * @property {string} category - 親のサブメニュー名（「書式」。直下の項目は空文字）
     * @property {string} commandId - executeMenuCommand に渡すメニューコマンド ID
     * @property {number} menuIndex - メニューの並び順（0 始まり。並び順の同順位もこれで決める）
     * @property {string} tableName - PANEL_TABLE の日本語名（PANEL_SETS から探す鍵）
     */

    /**
     * Illustrator のバージョンを数値で返す（"30.8.0" → 30.8）
     * @returns {number} メジャー.マイナーの数値
     */
    function getAppVersionNumber() {
        var versionParts = String(app.version).split(".");
        return parseFloat(versionParts[0] + "." + (versionParts[1] || "0"));
    }

    /**
     * パネルのメニューコマンド ID を返す（CC ライブラリだけは UI の言語のメニュー名を後ろに付ける）
     * @param {Array} panelRow - PANEL_TABLE の1行
     * @returns {string} メニューコマンド ID
     */
    function getPanelCommandId(panelRow) {
        if (panelRow[3] !== "") return panelRow[3];
        return CC_LIBRARIES_COMMAND_PREFIX + (uiLang === "ja" ? panelRow[0] : panelRow[1]);
    }

    /**
     * PANEL_TABLE から、この Illustrator で使えるパネルを PanelEntry にして集める
     * @returns {PanelEntry[]} パネルの一覧（メニューの並び順）
     */
    function collectPanelEntries() {
        var appVersion = getAppVersionNumber();
        var entries = [];
        for (var i = 0; i < PANEL_TABLE.length; i++) {
            var panelRow = PANEL_TABLE[i];
            if (appVersion < panelRow[4]) continue;
            entries.push({
                name: (uiLang === "ja") ? panelRow[0] : panelRow[1],
                category: (panelRow[2] === "") ? "" : getLabel(LABELS.category[panelRow[2]]),
                commandId: getPanelCommandId(panelRow),
                menuIndex: entries.length,
                tableName: panelRow[0]
            });
        }
        return entries;
    }

    /**
     * セットのパネルを、セットに書いた順で探す（この Illustrator で使えないパネルは飛ばす）
     * @param {PanelEntry[]} entries - すべてのパネル
     * @param {Object} panelSet - PANEL_SETS の1件
     * @returns {PanelEntry[]} 見つかったパネル
     */
    function findPanelSetEntries(entries, panelSet) {
        var found = [];
        for (var i = 0; i < panelSet.panels.length; i++) {
            for (var j = 0; j < entries.length; j++) {
                if (entries[j].tableName === panelSet.panels[i]) {
                    found.push(entries[j]);
                    break;
                }
            }
        }
        return found;
    }

    /**
     * セットのパネル名を、UI の言語で区切ってつなぐ（ツールチップ用）
     * @param {PanelEntry[]} entries - すべてのパネル
     * @param {Object} panelSet - PANEL_SETS の1件
     * @returns {string} 「文字・段落・…」のような文字列
     */
    function joinPanelSetNames(entries, panelSet) {
        var setEntries = findPanelSetEntries(entries, panelSet);
        var panelNames = [];
        for (var i = 0; i < setEntries.length; i++) {
            panelNames.push(setEntries[i].name);
        }
        return panelNames.join(uiLang === "ja" ? "・" : ", ");
    }

    /**
     * 検索用に文字列をそろえる（小文字にし、ひらがなをカタカナにする）
     * @param {string} sourceText - 元の文字列
     * @returns {string} そろえた文字列
     */
    function normalizeForSearch(sourceText) {
        return String(sourceText).toLowerCase().replace(/[ぁ-ゖ]/g, function (hiragana) {
            return String.fromCharCode(hiragana.charCodeAt(0) + 0x60);
        });
    }

    /**
     * 検索欄の文字列を語に分ける（空白は半角・全角とも区切り）
     * @param {string} searchText - 検索欄の文字列
     * @returns {string[]} そろえた語の一覧（空欄なら空配列）
     */
    function splitSearchWords(searchText) {
        var trimmedText = normalizeForSearch(searchText).replace(/^[\s　]+|[\s　]+$/g, "");
        return (trimmedText === "") ? [] : trimmedText.split(/[\s　]+/);
    }

    /**
     * 名前と分類（または漢字の読み）に、すべての語が含まれるパネルだけを返す
     * @param {PanelEntry[]} entries - パネルの一覧
     * @param {string} searchText - 検索欄の文字列
     * @returns {PanelEntry[]} 条件に合うパネル
     */
    function filterPanelEntries(entries, searchText) {
        var searchWords = splitSearchWords(searchText);
        if (searchWords.length === 0) return entries.slice(0);
        var matched = [];
        for (var i = 0; i < entries.length; i++) {
            /* 漢字の名前は読みでも探す（「しょしき」で［書式］）/ Kanji names also match their reading */
            var sourceText = entries[i].name + "\n" + entries[i].category;
            var haystack = normalizeForSearch(sourceText) + "\n" + normalizeForSearch(convertKanjiToReading(sourceText));
            var isMatch = true;
            for (var j = 0; j < searchWords.length; j++) {
                if (haystack.indexOf(searchWords[j]) < 0) {
                    isMatch = false;
                    break;
                }
            }
            if (isMatch) matched.push(entries[i]);
        }
        return matched;
    }

    /* 並べ替えキーの区切り（どの文字より前に並ぶ）/ Sort key separator (sorts before any character) */
    var SORT_KEY_SEPARATOR = "\u0001";

    /**
     * メニューの並び順を桁そろえした文字列にする（文字列の並べ替えで数値順になるように）
     * @param {number} menuIndex - メニューの並び順
     * @returns {string} 5桁にゼロ埋めした並び順
     */
    function padMenuIndex(menuIndex) {
        return ("00000" + menuIndex).slice(-5);
    }

    /**
     * 文字列の漢字を KANJI_READINGS の読み（カタカナ）に置き換える
     * @param {string} sourceText - 元の文字列
     * @returns {string} 漢字を読みにした文字列
     */
    function convertKanjiToReading(sourceText) {
        var readingText = String(sourceText);
        for (var i = 0; i < KANJI_READINGS.length; i++) {
            readingText = readingText.split(KANJI_READINGS[i][0]).join(KANJI_READINGS[i][1]);
        }
        return readingText;
    }

    /**
     * 並べ替え用に、名前を読みのカタカナにそろえる（英字は小文字）。
     * 漢字は KANJI_READINGS で読みに置き換え、長音（ー）と中黒（・）は読みの順に関係しないので除く
     * @param {string} panelName - パネル名
     * @returns {string} 並べ替え用の読み（英字は小文字）
     */
    function buildReadingKey(panelName) {
        return normalizeForSearch(convertKanjiToReading(panelName)).replace(/[ー・]/g, "");
    }

    /**
     * 並べ替え用の文字列キーを作る（同順位はメニューの並び順で決める）
     * @param {PanelEntry} panelEntry - 対象のパネル
     * @param {string} sortKey - SORT_KEYS のどれか
     * @returns {string} 並べ替えキー
     */
    function buildSortKey(panelEntry, sortKey) {
        var orderKey = padMenuIndex(panelEntry.menuIndex);
        /* 英語版の［ウィンドウ］メニューはアルファベット順なので、デフォルトも名前順にする / The English Window menu is alphabetical */
        if (sortKey === "kana" || uiLang !== "ja") return buildReadingKey(panelEntry.name) + SORT_KEY_SEPARATOR + orderKey;
        return orderKey;
    }

    /**
     * 指定した並び順に並べ替えた新しい配列を返す。
     * ExtendScript の比較関数つき sort() は遅く並びも狂うことがあるので、文字列キーを組み込みの sort() で並べる
     * @param {PanelEntry[]} entries - パネルの一覧（並べ替えない）
     * @param {string} sortKey - SORT_KEYS のどれか
     * @returns {PanelEntry[]} 並べ替えた一覧
     */
    function sortPanelEntries(entries, sortKey) {
        var sortKeys = [];
        for (var i = 0; i < entries.length; i++) {
            sortKeys.push(buildSortKey(entries[i], sortKey) + SORT_KEY_SEPARATOR + i);
        }
        sortKeys.sort();
        var sortedEntries = [];
        for (var j = 0; j < sortKeys.length; j++) {
            sortedEntries.push(entries[Number(sortKeys[j].substring(sortKeys[j].lastIndexOf(SORT_KEY_SEPARATOR) + 1))]);
        }
        return sortedEntries;
    }

    /**
     * パネルを表示する（メニューの項目を選ぶのと同じ）。
     * Illustrator のメニューは表示／非表示の切り替えで、スクリプトからは今の状態を読めないので、開いているパネルは閉じることがある
     * @param {PanelEntry[]} entries - 表示するパネル
     * @returns {string[]} 表示できなかったパネルの名前
     */
    function showPanels(entries) {
        var failedNames = [];
        for (var i = 0; i < entries.length; i++) {
            /* 古い ID などで例外になることがある / Outdated IDs may throw */
            try {
                app.executeMenuCommand(entries[i].commandId);
            } catch (e) {
                failedNames.push(entries[i].name);
            }
        }
        return failedNames;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 件数入りのダイアログタイトルを作る
     * @param {number} itemCount - 表示中の件数
     * @returns {string} ダイアログタイトル
     */
    function buildDialogTitle(itemCount) {
        var countText = itemCount + getLabel(LABELS.counter.itemUnit);
        return getLabel(LABELS.dialog.title) + (uiLang === "ja" ? "（" + countText + "）" : " (" + countText + ")");
    }

    /**
     * 名前・分類の2列のリストを作る
     * @param {Window} parent - 追加先
     * @returns {ListBox} パネルの一覧
     */
    function buildPanelList(parent) {
        var panelList = parent.add("listbox", [0, 0, LIST_SIZE[0], LIST_SIZE[1]], "", {
            multiselect: true,
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [
                getLabel(LABELS.column.name),
                getLabel(LABELS.column.category)
            ],
            columnWidths: COLUMN_WIDTHS
        });
        panelList.preferredSize = LIST_SIZE;
        panelList.alignment = ["fill", "fill"];
        panelList.helpTip = getLabel(LABELS.tooltip.list);
        return panelList;
    }

    /**
     * 右揃えのラベル付きの行を作る
     * @param {Window} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @returns {Group} 行グループ（入力部品はここへ追加する）
     */
    function addFieldRow(parent, labelSet) {
        var rowGroup = parent.add("group");
        setupRow(rowGroup, ["fill", "top"]);
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelSet));
        /* 幅を固定しないと右揃えが効かない / Right alignment needs a fixed width */
        rowLabel.minimumSize.width = ROW_LABEL_WIDTH;
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowGroup;
    }

    /**
     * 並び順のラジオボタンの行を作る（同じ行グループに入れて排他にする）
     * @param {Window} parent - 追加先
     * @param {string} selectedSortKey - 最初に選ぶ並び順
     * @returns {RadioButton[]} SORT_KEYS と同じ順のラジオボタン
     */
    function buildSortRow(parent, selectedSortKey) {
        var sortRow = addFieldRow(parent, LABELS.fieldLabel.sortOrder);
        var sortRadios = [];
        for (var i = 0; i < SORT_KEYS.length; i++) {
            var sortRadio = sortRow.add("radiobutton", undefined, getLabel(LABELS.sort[SORT_KEYS[i]]));
            sortRadio.helpTip = getLabel(LABELS.tooltip.sortOrder);
            sortRadio.value = (SORT_KEYS[i] === selectedSortKey);
            sortRadios.push(sortRadio);
        }
        return sortRadios;
    }

    /**
     * セットのボタンの行を作る
     * @param {Window} parent - 追加先
     * @param {PanelEntry[]} allEntries - すべてのパネル（ツールチップのパネル名に使う）
     * @returns {Button[]} PANEL_SETS と同じ順のボタン
     */
    function buildPanelSetRow(parent, allEntries) {
        var setRow = addFieldRow(parent, LABELS.fieldLabel.panelSet);
        var setButtons = [];
        for (var i = 0; i < PANEL_SETS.length; i++) {
            var setButton = setRow.add("button", undefined, getLabel(PANEL_SETS[i].label));
            setButton.helpTip = getLabel(LABELS.tooltip.panelSet, [joinPanelSetNames(allEntries, PANEL_SETS[i])]);
            setButtons.push(setButton);
        }
        return setButtons;
    }

    /**
     * 検索欄の行（検索欄とクリアボタン）を作る
     * @param {Window} parent - 追加先
     * @returns {{input: EditText, clearButton: Button}} 検索欄とクリアボタン
     */
    function buildSearchRow(parent) {
        var searchRow = addFieldRow(parent, LABELS.fieldLabel.search);
        var searchInput = searchRow.add("edittext", undefined, "");
        searchInput.alignment = ["fill", "center"];
        searchInput.helpTip = getLabel(LABELS.tooltip.search);
        searchInput.active = true;
        var clearButton = addClearButton(searchRow);
        return { input: searchInput, clearButton: clearButton };
    }

    /**
     * Enter／Return のキーイベントか
     * @param {KeyboardEvent} keyEvent - キーイベント
     * @returns {boolean} Enter／Return なら true
     */
    function isEnterKey(keyEvent) {
        return keyEvent.keyName === "Enter" || keyEvent.keyName === "Return";
    }

    /**
     * ダイアログを表示し、表示するパネルを選ばせる
     * @param {PanelEntry[]} allEntries - すべてのパネル
     * @returns {PanelEntry[]} 選ばれたパネル（キャンセルなら空配列）
     */
    function showPanelDialog(allEntries) {
        var visibleEntries = [];
        var pickedEntries = [];
        var lastSearchText = null;
        var currentSortKey = DEFAULT_SORT_KEY;

        var mainDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(mainDialog);

        var searchUI = buildSearchRow(mainDialog);
        var searchInput = searchUI.input;
        var clearButton = searchUI.clearButton;
        var sortRadios = buildSortRow(mainDialog, currentSortKey);
        var setButtons = buildPanelSetRow(mainDialog, allEntries);
        var panelList = buildPanelList(mainDialog);

        var buttonRow = addButtonRow(mainDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnShow = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.show), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * リストで選んでいるパネルを返す
         * @returns {PanelEntry[]} 選んでいるパネル（無ければ空配列）
         */
        function getSelectedEntries() {
            var selectedEntries = [];
            var selectedItems = panelList.selection;
            if (!selectedItems) return selectedEntries;
            for (var i = 0; i < selectedItems.length; i++) {
                selectedEntries.push(visibleEntries[selectedItems[i].index]);
            }
            return selectedEntries;
        }

        /**
         * 絞り込み・並べ替えの結果で一覧を作り直す。
         * 前に選んでいたパネルが残っていれば選び直し、無ければ先頭を選ぶ（Enter ですぐ表示できるように）
         * @param {PanelEntry[]} keepEntries - 選び直すパネル
         * @returns {void}
         */
        function rebuildList(keepEntries) {
            visibleEntries = sortPanelEntries(filterPanelEntries(allEntries, searchInput.text), currentSortKey);
            panelList.removeAll();
            var reselectIndexes = [];
            for (var i = 0; i < visibleEntries.length; i++) {
                var listItem = panelList.add("item", visibleEntries[i].name);
                listItem.subItems[0].text = visibleEntries[i].category;
                for (var j = 0; j < keepEntries.length; j++) {
                    if (keepEntries[j] === visibleEntries[i]) reselectIndexes.push(i);
                }
            }
            for (var k = 0; k < reselectIndexes.length; k++) {
                panelList.items[reselectIndexes[k]].selected = true;
            }
            if (reselectIndexes.length === 0 && visibleEntries.length > 0) panelList.selection = 0;
            mainDialog.text = buildDialogTitle(visibleEntries.length);
            btnShow.enabled = (visibleEntries.length > 0);
        }

        /**
         * 検索欄の文字が前回から変わっていれば一覧を作り直し、先頭を選ぶ
         * （onChanging と onChange の両方から呼ばれても1回で済ませる）
         * @returns {void}
         */
        function applySearchText() {
            /* 検索欄が空ならクリアボタンを無効にする / Disable the clear button while the field is empty */
            var hasText = (searchInput.text !== "");
            if (clearButton.enabled !== hasText) {
                clearButton.enabled = hasText;
                clearButton.notify("onDraw");
            }
            if (searchInput.text === lastSearchText) return;
            lastSearchText = searchInput.text;
            rebuildList([]);
        }

        /**
         * 並び順のラジオボタンで並べ替え、選んでいたパネルを選び直す
         * @param {string} sortKey - 選んだ並び順
         * @returns {void}
         */
        function applySortKey(sortKey) {
            if (sortKey === currentSortKey) return;
            currentSortKey = sortKey;
            rebuildList(getSelectedEntries());
        }

        /**
         * リストで選んだパネルを控えてダイアログを閉じる（選択が無ければ先頭のパネル）。
         * パネルの表示はダイアログを閉じたあとで行う
         * @returns {void}
         */
        function pickAndClose() {
            if (visibleEntries.length === 0) return;
            pickedEntries = getSelectedEntries();
            if (pickedEntries.length === 0) pickedEntries.push(visibleEntries[0]);
            mainDialog.close(1);
        }

        /**
         * セットのパネルを控えてダイアログを閉じる（検索・選択に関係なく、セットに書いたパネル）
         * @param {Object} panelSet - PANEL_SETS の1件
         * @returns {void}
         */
        function pickSetAndClose(panelSet) {
            pickedEntries = findPanelSetEntries(allEntries, panelSet);
            mainDialog.close(1);
        }

        searchInput.onChanging = applySearchText;
        searchInput.onChange = applySearchText;
        for (var i = 0; i < sortRadios.length; i++) {
            /* i をその場で閉じ込める / Capture i for each radio */
            sortRadios[i].onClick = (function (sortKey) {
                return function () { applySortKey(sortKey); };
            }(SORT_KEYS[i]));
        }
        for (var p = 0; p < setButtons.length; p++) {
            setButtons[p].onClick = (function (panelSet) {
                return function () { pickSetAndClose(panelSet); };
            }(PANEL_SETS[p]));
        }
        searchInput.addEventListener("keydown", function (keyEvent) {
            if (keyEvent.keyName === "Down" && panelList.items.length > 0) {
                panelList.active = true;
                if (!panelList.selection) panelList.selection = 0;
                keyEvent.preventDefault();
            } else if (isEnterKey(keyEvent)) {
                pickAndClose();
                keyEvent.preventDefault();
                if (keyEvent.stopPropagation) keyEvent.stopPropagation();
            }
        });
        clearButton.onClick = function () {
            searchInput.text = "";
            /* 続けて打ち直せるよう、フォーカスは入力欄へ戻す / Put the focus back in the field */
            searchInput.active = true;
            applySearchText();
        };
        panelList.onDoubleClick = pickAndClose;
        panelList.addEventListener("keydown", function (keyEvent) {
            if (isEnterKey(keyEvent)) {
                pickAndClose();
                keyEvent.preventDefault();
                if (keyEvent.stopPropagation) keyEvent.stopPropagation();
            }
        });
        btnShow.onClick = pickAndClose;

        applySearchText();
        addVersionFooter(mainDialog);
        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        if (mainDialog.show() !== 1) return [];
        return pickedEntries;
    }

    // =========================================
    // 実行 / Run
    // =========================================

    var allEntries = collectPanelEntries();
    if (allEntries.length === 0) {
        alert(getLabel(LABELS.alert.notFound));
        return;
    }
    var failedNames = showPanels(showPanelDialog(allEntries));
    if (failedNames.length > 0) {
        alert(labelText(LABELS.alert.showFailed) + "\n" + failedNames.join("\n"));
    }

}());
