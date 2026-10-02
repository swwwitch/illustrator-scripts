#target illustrator
#targetengine "FavoriteFontPickerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

よく使うフォントを分類と文字セットで絞り込んだ一覧から、クリックしたフォントを選択中のテキストに適用する常駐パレットです。
同じ書体の文字セット違い（Pr6N・Pr5 など）はチェックした文字セットの中で優先順位の高いものだけを残し、中国語・韓国語など日本語以外のフォントは外します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPickerPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ncf9ff6feebf0

### Overview

A persistent palette that lists the fonts you use often, narrowed by category and character set, and applies the one you click to the selected text.
Of the same typeface in several character sets (Pr6N, Pr5, etc.), only the highest-priority checked one is kept, and Chinese, Korean and other non-Japanese fonts are left out.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteFontPickerPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FavoriteFontPickerPalette";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "KOUJI & 相棒（Gem）";              /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2026-10-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-02";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPickerPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteFontPickerPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ncf9ff6feebf0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author KOUJI & 相棒（Gem）
 * @discussion 「よく使うフォントパネル」（favoriteFont_AI.jsx v1.1.2、Copyright (c) 2026 KOUJI & 相棒（Gem）、MIT License）をもとに改変
 */

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ［カスタムセット］1 の初期値（ファミリー名・PostScript 名の前方一致）。以後は一覧の下のボタンで出し入れし、設定ファイルに保存する
       Initial fonts of Custom set 1 (prefix match on the family or PostScript name); later changed in the dialog and saved */
    var CUSTOM_FONTS = ["Graphik", "DIN"];

    /* ［カスタムセット］の数 / Number of Custom sets */
    var CUSTOM_SET_COUNT = 5;

    /* 一覧から外すフォント（部分一致）。［すべて表示］では外さない
       Fonts left out of the list (substring match); ignored by Show All */
    var EXCLUDE_FONTS = ["NT"];

    /* EXCLUDE_FONTS に当たっても残すフォント（部分一致）
       Fonts kept even when EXCLUDE_FONTS matches them (substring match) */
    var RESCUE_FONTS = [];

    /* 同じ書体が複数あるとき残す接頭辞の優先順位（先頭ほど優先）
       Prefix priority when the same typeface exists more than once (first wins) */
    var PRIORITY_PREFIXES = ["A-OTF", "A P-OTF", "AP-OTF", "G-OTF", "U-OTF"];

    /* 同じ書体が複数あるとき残す文字セット（接尾辞）の優先順位（先頭ほど優先）。［文字セット］のチェックボックスにも並ぶ
       Character set (suffix) priority for the same typeface (first wins); also listed as Character set checkboxes */
    var PRIORITY_SUFFIXES = ["Pr6N", "Pr6", "Pr5N", "Pr5", "ProN", "Pro", "StdN", "Std"];

    /* メーカー別の分類（ファミリー名・PostScript 名の前方一致、大文字小文字は区別しない）
       Foundry categories (prefix match on the family or PostScript name, case-insensitive) */
    var FOUNDRY_FILTERS = [
        {
            key: "morisawa",
            label: { ja: "モリサワ", en: "Morisawa" },
            tooltip: { ja: "A-OTF・A P-OTF で始まるモリサワのフォント（リュウミン・新ゴなど）", en: "Morisawa fonts starting with A-OTF or A P-OTF (Ryumin, Shin Go, etc.)" },
            familyPrefixes: ["A-OTF", "A P-OTF", "AP-OTF"],
            psNamePrefixes: ["Ryumin", "ShinGo", "UDShinGo", "GothicMB101", "FutoGoB101", "MidashiGo"]
        },
        { key: "tb", label: { ja: "タイプバンク", en: "TypeBank" }, tooltip: { ja: "TB で始まるタイプバンクのフォント", en: "TypeBank fonts starting with TB" }, familyPrefixes: ["TB"], psNamePrefixes: ["TB"] },
        { key: "fot", label: { ja: "フォントワークス", en: "Fontworks" }, tooltip: { ja: "FOT- で始まるフォントワークスのフォント", en: "Fontworks fonts starting with FOT-" }, familyPrefixes: ["FOT-"], psNamePrefixes: ["FOT-"] },
        {
            key: "hiragino",
            label: { ja: "ヒラギノ", en: "Hiragino" },
            tooltip: { ja: "ヒラギノ角ゴ・明朝・丸ゴなど（中国語版は［日本語以外を除外］の「中国語」で外れる）", en: "Hiragino Kaku Gothic, Mincho, Maru Gothic, etc. (the Chinese versions follow Chinese under Exclude non-Japanese)" },
            familyPrefixes: ["ヒラギノ", "Hiragino"],
            psNamePrefixes: ["Hiragino", "HiraKaku", "HiraMin", "HiraMaru"]
        },
        {
            /* Adobe Fonts（配信サービス）とは別に、アドビ自社の和文書体をまとめる
               Adobe's own Japanese typefaces, apart from the Adobe Fonts service */
            key: "adobe",
            label: { ja: "アドビ", en: "Adobe" },
            tooltip: { ja: "小塚・源ノ角ゴシック／明朝・りょう・貂明朝・百千鳥・かづらきなど、アドビ自社の和文書体", en: "Adobe's own Japanese typefaces such as Kozuka, Source Han, Ryo, Ten Mincho, Momochidori and Kazuraki" },
            familyPrefixes: ["小塚", "源ノ", "りょう", "貂", "百千鳥", "かづらき", "Kozuka", "Source Han", "Ten Mincho", "Momochidori", "Kazuraki"],
            psNamePrefixes: ["Koz", "SourceHan", "RyoGothic", "RyoDisp", "RyoText", "TenMincho", "Momochidori", "Kazuraki"]
        },
        {
            /* 名前では見分けられないので、Creative Cloud の同期状態（entitlements.xml）でアクティベート中のファミリーを引く
               Not detectable by name, so active families are looked up in the Creative Cloud sync state */
            key: "adobeFonts",
            label: { ja: "Adobe Fonts", en: "Adobe Fonts" },
            tooltip: { ja: "Creative Cloud でアクティベートしている Adobe Fonts のフォント", en: "Adobe Fonts activated through Creative Cloud" },
            familyPrefixes: [],
            psNamePrefixes: [],
            usesAdobeFontsList: true
        }
    ];

    /* ［日本語以外を除外］で外す言語。名前（ファミリー名・PostScript 名）で判定し、上から順に当てる
       namePrefixes：前方一致（大文字小文字は区別しない）、familyPatterns：ファミリー名、psNamePatterns：PostScript 名
       Languages left out with Exclude non-Japanese, judged by name and tried from the top */
    var LANGUAGE_GROUPS = [
        {
            key: "chinese",
            label: { ja: "中国語", en: "Chinese" },
            tooltip: { ja: "PingFang・宋体・黑体・Source Han Sans SC など、中国語（簡体字・繁体字）のフォント", en: "Chinese (Simplified and Traditional) fonts such as PingFang, Songti, Heiti and Source Han Sans SC" },
            namePrefixes: ["PingFang", "Hiragino Sans GB", "Hiragino Sans CNS", "HiraginoSansCNS", "STHeiti", "STSong", "STKaiti", "STFangsong", "STXihei", "Songti", "Heiti", "Kaiti", "Lantinghei", "Hannotate", "Hanzipen", "Weibei", "Libian", "Xingkai", "Baoli", "Wawati", "Yuanti", "Yuppy", "BiauKai", "LiSong", "LiHei", "Apple LiGothic", "Apple LiSung", "GB18030", "SimSun", "NSimSun", "SimHei", "Microsoft YaHei", "Microsoft JhengHei", "MingLiU", "PMingLiU", "DFKai", "AdobeSong", "AdobeMing", "AdobeHeiti", "AdobeFangsong", "AdobeKaiti", "思源", "苹方", "蘋方", "华文", "華文", "黑体", "黑體", "宋体", "宋體", "楷体", "仿宋", "圆体", "圓體", "兰亭", "蘭亭", "翩翩", "魏碑", "隶变", "隸變", "行楷", "报隶", "報隸", "娃娃", "雅痞"],
            familyPatterns: [/ (SC|TC|HK|CN)$/],
            /* 地域の印は直前が小文字のときだけ（ZapfDingbatsITC の TC、RoNOWStd-GB のウェイト GB を拾わない）
               Region marks count only after a lowercase letter, so ITC or a GB weight name is not picked up */
            psNamePatterns: [/^[^-]*[a-z0-9](SC|TC|HK|CN|GB)(-|$)/, /CJK(sc|tc|hk)/i]
        },
        {
            key: "korean",
            label: { ja: "韓国語", en: "Korean" },
            tooltip: { ja: "Apple SD Gothic Neo・AppleMyungjo・Nanum など、ハングルのフォント", en: "Korean fonts such as Apple SD Gothic Neo, AppleMyungjo and Nanum" },
            namePrefixes: ["AppleGothic", "AppleMyungjo", "AppleSDGothicNeo", "Apple SD", "Nanum", "Malgun", "Gulim", "Batang", "Dotum", "Gungsuh", "PCMyungjo", "HeadLineA", "AdobeMyungjo", "AdobeGothicStd", "JCsmPC", "JCfg", "JCkg", "JCHEadA"],
            familyPatterns: [/[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF]/, / KR$/],
            psNamePatterns: [/^[^-]*[a-z0-9]KR(-|$)/, /CJKkr/i]
        },
        {
            key: "thai",
            label: { ja: "タイ語", en: "Thai" },
            tooltip: { ja: "Thonburi・Sukhumvit Set など、タイ語のフォント", en: "Thai fonts such as Thonburi and Sukhumvit Set" },
            namePrefixes: ["Thonburi", "Ayuthaya", "Krungthep", "Sathu", "Silom", "Sukhumvit"],
            familyPatterns: [/[\u0E00-\u0E7F]/, /\bThai\b/i],
            psNamePatterns: []
        },
        {
            key: "multilingual",
            label: { ja: "その他", en: "Other" },
            tooltip: {
                ja: "アラビア語・ヘブライ語・インドの諸文字など、そのほかの言語のフォント（Noto Sans の各言語版を含む）",
                en: "Fonts for other scripts such as Arabic, Hebrew and Indic (including the Noto Sans language versions)"
            },
            namePrefixes: ["Baghdad", "Damascus", "Farah", "Farisi", "Geeza", "Al Bayan", "Al Nile", "Al Tarikh", "Beirut", "Decotype Naskh", "Diwan", "Kufi", "Mishafi", "Muna", "Nadeem", "Sana", "Waseem", "Raanana", "Mshtakan", "Kefa", "Euphemia", "Kailasa", "InaiMathi", "Shree Devanagari"],
            familyPatterns: [
                /[\u0530-\u058F\u0590-\u05FF\u0600-\u06FF\u0700-\u074F\u0780-\u07BF\u0900-\u0DFF\u0E80-\u0FFF\u1000-\u109F\u10A0-\u10FF\u1200-\u137F\u1780-\u17FF\u1800-\u18AF]/,
                /\b(Arabic|Hebrew|Urdu|Persian|Devanagari|Bangla|Bengali|Gujarati|Gurmukhi|Kannada|Malayalam|Oriya|Odia|Sinhala|Tamil|Telugu|Myanmar|Khmer|Lao|Tibetan|Ethiopic|Armenian|Georgian|Mongolian|Cherokee|Syriac|Thaana|Grantha|Marathi|Nastaliq|Naskh|Sangam)\b/i,
                /* Noto Sans／Serif の言語版（JP・記号などは除く）/ Noto Sans/Serif language versions, except JP, symbols and the like */
                /^Noto (Sans|Serif)(?! (JP|CJK JP|Mono|Mono CJK JP|Display|Symbols|Symbols 2|Math|Music)$) \S/
            ],
            psNamePatterns: []
        }
    ];

    /* ドキュメントのフォントを書き出すファイル名（ドキュメントと同じフォルダー、InDesign 版と共通）
       File that holds the document's fonts (next to the document, shared with the InDesign version) */
    var PROJECT_FONTS_FILE_NAME = "_ProjectFonts.txt";

    /* 絞り込み欄に打つたびに一覧を更新する件数の上限。超えるときは Enter で更新する
       Max result count updated on every keystroke in the filter field; above it, press Enter */
    var LIVE_SEARCH_MAX_FONTS = 800;

    // =========================================
    // レイアウト / Layout
    // =========================================

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

    /* 自前描画のたびに環境設定を読まないよう、1回だけ控える / Read once so custom drawing does not query preferences every time */
    var UI_IS_DARK = isDarkUI();

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

    var FONT_LIST_SIZE = [320, 400];   /* フォント一覧の大きさ（高さは1行およそ 20px）/ font list size (about 20 px per row) */
    var FONT_LIST_TOP_MARGIN = 10;     /* フォント一覧の上に足す余白（px）/ extra space above the font list */
    var FONT_LIST_BOTTOM_MARGIN = 10;  /* フォント一覧の下に足す余白（px）/ extra space below the font list */

    /* 絞り込み欄のクリアボタン（自前描画）/ Clear button next to the filter field (custom drawn) */
    var CLEAR_BUTTON_SIZE = 20;        /* クリアボタンの一辺 / clear button size */
    var CLEAR_CIRCLE_INSET = 2;        /* 円と外周の間隔 / inset of the circle */
    var CLEAR_GLYPH_INSET = 6;         /* ×と外周の間隔 / inset of the × glyph */
    var CLEAR_STROKE_WIDTH = 1.5;      /* 円と×の線幅 / stroke width */

    /* ［カスタムセット］の出し入れボタン（重なった2枚のカードに番号としおりを自前描画）
       Custom set toggle buttons (two stacked cards with a number and a bookmark, custom drawn) */
    var SET_BUTTON_SIZE = 33;          /* ボタンの一辺 / button size */
    var SET_CARD_SIZE = 22;            /* カード1枚の一辺 / card size */
    var SET_CARD_OFFSET = 5;           /* 後ろのカードとのずれ / offset of the back card */
    var SET_CARD_STROKE_WIDTH = 2;     /* カードの線幅 / card stroke width */
    var SET_BOOKMARK_WIDTH = 7;        /* しおりの幅 / bookmark width */
    var SET_BOOKMARK_HEIGHT = 10;      /* しおりの高さ / bookmark height */
    var SET_BOOKMARK_NOTCH = 3;        /* しおりの切り込みの深さ / depth of the bookmark notch */
    var SET_BOOKMARK_INSET = 3;        /* しおりとカードの右端の間隔 / gap between the bookmark and the card's right edge */
    var SET_NUMBER_FONT_SIZE = 13;     /* 番号の文字サイズ / number font size */
    var STYLE_ROW_INDENT = "\u3000\u3000"; /* ウェイト（スタイル）行の字下げ / indent of style rows */
    var STANDARD_LEFT_WIDTH = 64;      /* ［文字セット］の左列（無印）の幅 / width of the left (plain) column under Character set */
    var LANGUAGE_LEFT_WIDTH = 96;      /* 環境設定の、外す言語の左列の幅 / left column width of the languages in Preferences */
    var CUSTOM_SET_LIST_SIZE = [260, 160]; /* 環境設定の、カスタムセットの中身の一覧の大きさ / size of the Custom set contents list in Preferences */

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
    // 設定の保存 / Settings
    // =========================================

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

    /* 設定とフォント情報のキャッシュのファイル名。改名前の名前のまま据え置く（保存済みの設定を読み継ぐ）
       Base name of the settings and font cache files, kept from before the rename so saved settings carry over */
    var SETTINGS_FILE_BASE_NAME = "FavoriteFontPicker";

    var settingsStore = createSettingsStore(SETTINGS_FILE_BASE_NAME, "persistent");

    /* 開いているパレットを $.global に載せるキー（二重起動を防ぐ）/ $.global key of the open palette (prevents a second instance) */
    var PALETTE_GLOBAL_KEY = "__FavoriteFontPickerPalette";

    /* ダイアログの初期値 / Dialog defaults */
    var DEFAULT_SETTINGS = {
        customSets: null,        /* ［カスタムセット］ごとのフォント名。null は未保存（resolveCustomSets() で作る）/ names per Custom set; null when unsaved */
        customSetsVisible: null, /* ［カスタムセット］ごとのチェック / checked state per Custom set */
        custom: null,            /* 旧形式の［カスタム］のチェック（セット1に読み継ぐ）/ legacy Custom checkbox, carried into set 1 */
        customFonts: null,       /* 旧形式の［カスタム］のフォント名（セット1に読み継ぐ）/ legacy Custom names, carried into set 1 */
        documentFonts: false,
        compositeFonts: false,
        foundries: [],          /* チェックしたメーカーの key / keys of the checked foundries */
        showAll: false,
        filterByStandard: true, /* ［文字セット］のチェックで絞り込む / filter by the Character set checkboxes */
        japaneseOnly: true,     /* ［和文フォントのみ］/ Japanese fonts only */
        excludedLanguages: ["chinese", "korean", "thai", "multilingual"], /* ［和文フォントのみ］で外す言語（環境設定）/ languages left out by Japanese fonts only (preferences) */
        hiddenSuffixes: [],     /* ［文字セット］で外した接尾辞（"" は［その他］）/ suffixes unchecked under Character set ("" is Other) */
        searchText: "",
        showPostScriptName: false   /* 一覧を PostScript 名で表示 / list PostScript names */
    };

    /* フォント情報のキャッシュ（1行に PostScript 名・ファミリー名・スタイル名をタブ区切り）
       Font cache: one font per line, PostScript name, family and style separated by tabs */
    var FONT_CACHE_FILE = new File(Folder.userData + "/illustrator-scripts/" + SETTINGS_FILE_BASE_NAME + "-fonts.txt");

    /* Creative Cloud が Adobe Fonts の同期状態を書くファイル（Mac は .c、Windows は c のフォルダー）
       Creative Cloud's Adobe Fonts sync state (in .c on Mac, c on Windows) */
    var ADOBE_FONTS_STATE_PATHS = [
        "/Adobe/CoreSync/plugins/livetype/.c/entitlements.xml",
        "/Adobe/CoreSync/plugins/livetype/c/entitlements.xml"
    ];

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

    var LABELS = {
        dialog: {
            title: { ja: "よく使うフォント", en: "Favorite Fonts" },
            preferences: { ja: "環境設定", en: "Preferences" },
            exportSet: { ja: "カスタムセット {number} を書き出す", en: "Export Custom set {number}" },
            importSet: { ja: "カスタムセット {number} に読み込む", en: "Import into Custom set {number}" },
            progress: { ja: "フォント情報を読み込み中", en: "Reading font information" }
        },
        panel: {
            category: { ja: "分類", en: "Category" },
            standard: { ja: "文字セット", en: "Standard" },
            excludedLanguages: { ja: "［和文フォントのみ］で外す言語", en: "Languages left out by Japanese fonts only" },
            fontInfo: { ja: "フォント情報", en: "Font information" },
            customSets: { ja: "カスタムセット", en: "Custom sets" }
        },
        checkbox: {
            documentFonts: { ja: "ドキュメントフォント", en: "Document fonts" },
            compositeFonts: { ja: "合成フォント", en: "Composite fonts" },
            showAll: { ja: "すべて表示", en: "Show all" },
            japaneseOnly: { ja: "和文フォントのみ", en: "Japanese fonts only" },
            filterByStandard: { ja: "文字セットで絞り込む", en: "Filter by standard" },
            noSuffix: { ja: "文字セットなし", en: "No standard" },
            showPostScriptName: { ja: "PostScript 名で表示", en: "Show PostScript names" }
        },
        fieldLabel: {
            search: { ja: "絞り込み", en: "Filter" },
            customSets: { ja: "カスタムセット", en: "Custom sets" }
        },
        button: {
            rescan: { ja: "再スキャン", en: "Rescan" },
            preferences: { ja: "環境設定...", en: "Preferences..." },
            exportSet: { ja: "書き出し...", en: "Export..." },
            importSet: { ja: "読み込み...", en: "Import..." },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            customSet: {
                ja: "カスタムセット {number}：一覧の下の {number} のボタンで加えたフォント。文字セット違いの間引きと除外はしない",
                en: "Custom set {number}: fonts added with button {number} below the list, shown without thinning out character sets or exclusions"
            },
            documentFonts: {
                ja: "ドキュメントと同じフォルダーの _ProjectFonts.txt に挙げたフォント。文字セット違いの間引きと除外はしない",
                en: "Fonts listed in _ProjectFonts.txt next to the document, shown without thinning out character sets or exclusions"
            },
            compositeFonts: {
                ja: "［書式］→［合成フォント］で作った合成フォント。文字セットの絞り込みと間引きはしない",
                en: "Composite fonts made with Type > Composite Fonts, not filtered or thinned out by standard"
            },
            showAll: {
                ja: "［分類］・文字セット違いの間引き・除外をやめてすべて表示（［文字セット］と［和文フォントのみ］は効く）",
                en: "Ignore Category, character set thinning and exclusions and list every font (Character set and Japanese fonts only still apply)"
            },
            standard: {
                ja: "チェックした文字セットだけを表示。同じ書体はチェックした中で優先順位がいちばん高い文字セットを残す",
                en: "Show only the checked character sets; for each typeface, keep the highest-priority checked standard"
            },
            filterByStandard: {
                ja: "オフにすると［文字セット］のチェックを無視する（同じ書体は、すべての文字セットの中で優先順位がいちばん高いものを残す）",
                en: "When off, the Character set checkboxes are ignored (each typeface keeps its highest-priority character set of all)"
            },
            noSuffix: { ja: "Pr6N などの文字セットが名前の末尾に付いていないフォント", en: "Fonts without a character set such as Pr6N at the end of the name" },
            search: { ja: "ファミリー名・スタイル名・PostScript 名の部分一致", en: "Substring match on the family, style or PostScript name" },
            clearSearch: { ja: "絞り込みをクリア", en: "Clear filter" },
            showPostScriptName: { ja: "一覧を PostScript 名（RyuminPr6N-Light など）で表示し、その順に並べる", en: "List PostScript names (such as RyuminPr6N-Light) and sort by them" },
            rescan: { ja: "インストールされているフォントを読み直す（フォントを追加・削除したあとに）", en: "Read the installed fonts again (after adding or removing fonts)" },
            preferences: { ja: "再スキャン、［和文フォントのみ］で外す言語、カスタムセットの書き出し・読み込み", en: "Rescan, languages left out by Japanese fonts only, and Custom set export and import" },
            japaneseOnly: { ja: "中国語・韓国語などのフォントを外す（外す言語は環境設定で選ぶ）。カスタムセット・ドキュメントフォントに挙げたものは残す", en: "Leave out Chinese, Korean and other fonts (choose the languages in Preferences); listed Custom set and Document fonts stay" },
            exportSet: { ja: "このセットの名前を1行1件のテキストファイルに書き出す", en: "Write this set's names to a text file, one per line" },
            importSet: { ja: "1行1件のテキストファイルから名前を読み込み、このセットに加える", en: "Read names from a text file, one per line, and add them to this set" },
            addCustomSet: { ja: "カスタムセット {number} に追加：一覧で選んだフォントのファミリーを加える", en: "Add to Custom set {number}: add the family of the chosen font" },
            removeCustomSet: { ja: "カスタムセット {number} から外す：一覧で選んだフォントに当たる名前を外す（同じ名前で始まるほかのフォントも外れる）", en: "Remove from Custom set {number}: remove the names matching the chosen font (other fonts starting with them are removed too)" },
            fontList: { ja: "クリックしたフォントを選択中のテキストに適用（選んだままの行はダブルクリックまたは Enter キー）", en: "Click a font to apply it to the selected text (double-click or press Enter for the row already chosen)" },
            soloClick: { ja: "option（Alt）＋クリック：これだけオン／すべてオン", en: "Option (Alt)-click: only this one / all on" }
        },
        message: {
            fontTotal: { ja: "{count} 件のフォントを読み込み済み", en: "{count} fonts loaded" },
            setCount: { ja: "{count} 件の名前", en: "{count} names" },
            fontCount: { ja: "{count} 件のフォント", en: "{count} fonts" },
            searchPending: { ja: "{count} 件のフォント（Enter キーで一覧を更新）", en: "{count} fonts (press Enter to update the list)" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextSelected: {
                ja: "テキストフレーム、またはテキストを含むグループを選択してから実行してください。",
                en: "Select a text frame, or a group that contains text, and try again."
            },
            fontNotFound: { ja: "フォントが見つかりません：{name}", en: "Font not found: {name}" },
            workerFailed: { ja: "ドキュメントを操作できませんでした（{action}）：{message}", en: "Could not work on the document ({action}): {message}" },
            rescanFailed: {
                ja: "パレットからはフォント情報を読み直せませんでした。\nパレットを閉じて開き直すと読み直します。",
                en: "Font information could not be reread from the palette.\nClose and reopen the palette to reread it."
            },
            rescanned: { ja: "フォント情報を更新しました（{count} 件）。", en: "Font information updated ({count} fonts)." },
            exported: { ja: "{count} 件の名前を書き出しました。", en: "Exported {count} names." },
            imported: { ja: "{count} 件の名前を加えました（重複は除く）。", en: "Added {count} names (duplicates skipped)." },
            writeFailed: { ja: "ファイルを書き出せませんでした：{message}", en: "Could not write the file: {message}" },
            confirmRemoveCustom: {
                ja: "カスタムセット {number} から「{names}」を外しますか？\nこの名前で始まるほかのフォントも外れます。",
                en: "Remove \"{names}\" from Custom set {number}?\nOther fonts starting with this name are removed too."
            }
        }
    };

    // =========================================
    // フォント情報 / Font information
    // =========================================

    /* キャッシュが今のフォントと同じかを確かめる見本の数（一覧から等間隔に取る PostScript 名）
       Number of PostScript names sampled at even intervals to check that the cache is current */
    var CACHE_SAMPLE_COUNT = 32;

    /* キャッシュの書式。中身の決まりを変えたら上げて、古いキャッシュを読み直させる
       Cache format; bump it when the contents change so old caches are rebuilt */
    var CACHE_FORMAT = "2";

    /* 名前の前後の区切り（半角・全角の空白、ハイフン、アンダースコア）/ Separators around name parts */
    var NAME_SEPARATOR_CHARS = " \t\u3000-_";

    /**
     * インストールされているフォントを読み、キャッシュに書く
     * 環境にないフォントの仮エントリは外す。合成フォントは PostScript 名が ATC-<名前の16進>、ファミリー名が合成フォント名
     * @returns {Object[]} フォント情報 { psName, family, style } の配列
     */
    function scanInstalledFonts() {
        var textFonts = app.textFonts;
        var fontTotal = textFonts.length;
        var progressWindow = createProgressWindow(fontTotal);
        var fontInfos = [];
        for (var i = 0; i < fontTotal; i++) {
            if (i % 200 === 0) updateProgressWindow(progressWindow, i);
            var psName, familyName, styleName;
            try {
                var textFont = textFonts[i];
                psName = String(textFont.name || "");
                familyName = String(textFont.family || "");
                styleName = String(textFont.style || "");
            } catch (e) {
                /* 壊れたフォントはプロパティを読むと例外になることがある / A broken font may throw on property access */
                continue;
            }
            if (psName === "") continue;
            /* 環境にないフォントは style が空で family に PostScript 名が入る / Missing fonts have an empty style and the PostScript name as family */
            if (styleName === "" && familyName === psName) continue;
            fontInfos.push({ psName: psName, family: familyName || psName, style: styleName });
        }
        progressWindow.close();
        writeFontCache(fontInfos, fontTotal, readFontSample(textFonts));
        return fontInfos;
    }

    /**
     * フォント一覧から等間隔に PostScript 名を取り、見本の文字列にする（総数が同じままの入れ替えを見分ける）
     * @param {TextFonts} textFonts - app.textFonts
     * @returns {string} 名前を "|" でつないだ文字列
     */
    function readFontSample(textFonts) {
        var fontTotal = textFonts.length;
        var sampleCount = Math.min(CACHE_SAMPLE_COUNT, fontTotal);
        var sampleNames = [];
        for (var i = 0; i < sampleCount; i++) {
            var fontIndex = (sampleCount > 1) ? Math.floor(i * (fontTotal - 1) / (sampleCount - 1)) : 0;
            try {
                sampleNames.push(textFonts[fontIndex].name);
            } catch (e) {
                /* 壊れたフォントはプロパティを読むと例外になることがある / A broken font may throw on property access */
                sampleNames.push("");
            }
        }
        return sampleNames.join("|");
    }

    /**
     * 読み込みの進み具合を出す小さなウィンドウを作る
     * @param {number} fontTotal - フォントの総数
     * @returns {Window} 進み具合のウィンドウ
     */
    function createProgressWindow(fontTotal) {
        var progressWindow = new Window("palette", getLabel("dialog.progress"));
        setupWindow(progressWindow);
        progressWindow.progressBar = progressWindow.add("progressbar", undefined, 0, Math.max(fontTotal, 1));
        progressWindow.progressBar.preferredSize.width = 300;
        progressWindow.show();
        return progressWindow;
    }

    /**
     * 進み具合のウィンドウを更新する
     * @param {Window} progressWindow - createProgressWindow() の戻り値
     * @param {number} doneCount - 読み終えた数
     * @returns {void}
     */
    function updateProgressWindow(progressWindow, doneCount) {
        progressWindow.progressBar.value = doneCount;
        progressWindow.update();
    }

    /**
     * フォント情報をキャッシュに書く。先頭3行は書式・フォントの総数・見本（変わったら読み直す目印）
     * @param {Object[]} fontInfos - フォント情報の配列
     * @param {number} fontTotal - app.textFonts.length
     * @param {string} fontSample - readFontSample() の戻り値
     * @returns {void}
     */
    function writeFontCache(fontInfos, fontTotal, fontSample) {
        var cacheLines = ["#format\t" + CACHE_FORMAT, "#total\t" + fontTotal, "#sample\t" + fontSample];
        for (var i = 0; i < fontInfos.length; i++) {
            cacheLines.push([fontInfos[i].psName, fontInfos[i].family, fontInfos[i].style].join("\t"));
        }
        if (!FONT_CACHE_FILE.parent.exists) FONT_CACHE_FILE.parent.create();
        writeTextFile(FONT_CACHE_FILE, cacheLines.join("\n"));
    }

    /**
     * キャッシュを読む。書式・フォントの総数・見本のどれかが今と違えば null（読み直しが要る）
     * @returns {Object[]|null} フォント情報の配列、使えないときは null
     */
    function readFontCache() {
        var cacheLines = splitTextLines(readTextFile(FONT_CACHE_FILE));
        var textFonts = app.textFonts;
        if (cacheLines[0] !== "#format\t" + CACHE_FORMAT) return null;
        if (cacheLines[1] !== "#total\t" + textFonts.length) return null;
        if (cacheLines[2] !== "#sample\t" + readFontSample(textFonts)) return null;
        var fontInfos = [];
        for (var i = 3; i < cacheLines.length; i++) {
            var cacheFields = cacheLines[i].split("\t");
            if (cacheFields.length < 3 || cacheFields[0] === "") continue;
            fontInfos.push({ psName: cacheFields[0], family: cacheFields[1], style: cacheFields[2] });
        }
        return fontInfos;
    }

    /**
     * フォント情報に絞り込み用の値（小文字の名前・名前の分解）を足す
     * @param {Object[]} fontInfos - フォント情報の配列（書き換える）
     * @returns {Object[]} 同じ配列
     */
    function createFontCatalog(fontInfos) {
        for (var i = 0; i < fontInfos.length; i++) {
            var fontInfo = fontInfos[i];
            fontInfo.order = i;
            fontInfo.psLower = fontInfo.psName.toLowerCase();
            fontInfo.familyLower = fontInfo.family.toLowerCase();
            fontInfo.styleLower = fontInfo.style.toLowerCase();
            fontInfo.isComposite = /^ATC-/i.test(fontInfo.psName);
            fontInfo.languageKey = fontInfo.isComposite ? "" : detectLanguageKey(fontInfo);
            fontInfo.parsed = parseFamilyName(fontInfo.family);
            /* 名前だけで決まる判定は、チェックのたびにくり返さないよう最初に済ませる / Name-only checks are done once, not on every click */
            fontInfo.isExcludedName = fontNameContainsAny(fontInfo, EXCLUDE_FONTS) && !fontNameContainsAny(fontInfo, RESCUE_FONTS);
            fontInfo.foundryMatches = {};
            for (var j = 0; j < FOUNDRY_FILTERS.length; j++) {
                if (!FOUNDRY_FILTERS[j].usesAdobeFontsList && matchesFoundryName(fontInfo, FOUNDRY_FILTERS[j])) fontInfo.foundryMatches[FOUNDRY_FILTERS[j].key] = true;
            }
        }
        return fontInfos;
    }

    /**
     * 名前から言語のグループを判定する（LANGUAGE_GROUPS の上から順に当てる）
     * @param {Object} fontInfo - psName・family・psLower・familyLower を持つフォント情報
     * @returns {string} LANGUAGE_GROUPS の key。どれにも当たらなければ ""
     */
    function detectLanguageKey(fontInfo) {
        for (var i = 0; i < LANGUAGE_GROUPS.length; i++) {
            var languageGroup = LANGUAGE_GROUPS[i];
            if (startsWithAny(fontInfo.familyLower, languageGroup.namePrefixes) || startsWithAny(fontInfo.psLower, languageGroup.namePrefixes)) return languageGroup.key;
            if (matchesAnyPattern(fontInfo.family, languageGroup.familyPatterns)) return languageGroup.key;
            if (matchesAnyPattern(fontInfo.psName, languageGroup.psNamePatterns)) return languageGroup.key;
        }
        return "";
    }

    /**
     * 文字列が正規表現のどれかに当たるか
     * @param {string} sourceText - 文字列
     * @param {RegExp[]} namePatterns - 正規表現の一覧
     * @returns {boolean} 当たれば true
     */
    function matchesAnyPattern(sourceText, namePatterns) {
        for (var i = 0; i < namePatterns.length; i++) {
            if (namePatterns[i].test(sourceText)) return true;
        }
        return false;
    }

    /**
     * 中心の名前ごとに、［文字セット］でチェックした中で優先順位がいちばん高い接頭辞・接尾辞の組を選ぶ
     * @param {Object[]} fontInfos - createFontCatalog() 済みのフォント情報
     * @param {Object} visibleSuffixes - 表示する接尾辞 → true（"" は文字セットなし）
     * @returns {Object} 中心の名前 → 残す組（parseFamilyName() の結果）
     */
    function buildBestVariantMap(fontInfos, visibleSuffixes) {
        var bestByCore = {};
        for (var i = 0; i < fontInfos.length; i++) {
            /* 合成フォントは名前が自由なので、文字セット違いの比較に入れない / Composite names are free-form, so they stay out of the comparison */
            if (fontInfos[i].isComposite) continue;
            var parsedName = fontInfos[i].parsed;
            if (visibleSuffixes[parsedName.suffix] !== true) continue;
            var bestVariant = bestByCore[parsedName.core];
            if (!bestVariant || parsedName.prefixRank < bestVariant.prefixRank ||
                (parsedName.prefixRank === bestVariant.prefixRank && parsedName.suffixRank < bestVariant.suffixRank)) {
                bestByCore[parsedName.core] = parsedName;
            }
        }
        return bestByCore;
    }

    /**
     * ファミリー名を接頭辞・中心の名前・接尾辞（文字セット）に分ける
     * @param {string} familyName - ファミリー名
     * @returns {{core: string, prefixRank: number, prefixText: string, suffixRank: number, suffixText: string, suffix: string}}
     *          中心の名前（小文字）、接頭辞・接尾辞の優先順位（無ければリストの長さ）と区切りを含む文字列、区切りを除いた接尾辞（無ければ ""）
     */
    function parseFamilyName(familyName) {
        var lowerName = familyName.toLowerCase();
        var prefixRank = PRIORITY_PREFIXES.length;
        var prefixText = "";
        for (var i = 0; i < PRIORITY_PREFIXES.length; i++) {
            var prefixLower = PRIORITY_PREFIXES[i].toLowerCase();
            if (lowerName.indexOf(prefixLower) !== 0) continue;
            var trailingSeparator = /^[ \t\u3000\-_]+/.exec(lowerName.substring(prefixLower.length));
            prefixRank = i;
            prefixText = lowerName.substring(0, prefixLower.length + (trailingSeparator ? trailingSeparator[0].length : 0));
            break;
        }
        var suffixRank = PRIORITY_SUFFIXES.length;
        var suffixText = "";
        var suffix = "";
        for (var j = 0; j < PRIORITY_SUFFIXES.length; j++) {
            var suffixLower = PRIORITY_SUFFIXES[j].toLowerCase();
            var suffixStart = lowerName.length - suffixLower.length;
            if (suffixStart <= 0 || lowerName.substring(suffixStart) !== suffixLower) continue;
            /* 直前が区切りのときだけ接尾辞とみなす（"Pro" で終わる単語を誤認しない）/ Only a separated suffix counts */
            if (NAME_SEPARATOR_CHARS.indexOf(lowerName.charAt(suffixStart - 1)) === -1) continue;
            var leadingSeparator = /[ \t\u3000\-_]+$/.exec(lowerName.substring(0, suffixStart));
            suffixRank = j;
            suffixText = lowerName.substring(suffixStart - leadingSeparator[0].length);
            suffix = PRIORITY_SUFFIXES[j];
            break;
        }
        var core = lowerName.substring(prefixText.length, lowerName.length - suffixText.length);
        core = core.replace(/^[ \t\u3000\-_]+|[ \t\u3000\-_]+$/g, "");
        return { core: core, prefixRank: prefixRank, prefixText: prefixText, suffixRank: suffixRank, suffixText: suffixText, suffix: suffix };
    }

    /**
     * 名前の一覧を小文字にする（照合のたびに小文字にしないよう、先に1回だけ）
     * @param {string[]} fontNames - 名前の一覧
     * @returns {string[]} 小文字にした新しい配列
     */
    function toLowerNames(fontNames) {
        var lowerNames = [];
        for (var i = 0; i < fontNames.length; i++) lowerNames.push(String(fontNames[i]).toLowerCase());
        return lowerNames;
    }

    /**
     * 小文字の文字列が、一覧のどれかで始まるか（大文字小文字は区別しない）
     * @param {string} lowerText - 小文字の文字列
     * @param {string[]} namePrefixes - 名前の一覧
     * @returns {boolean} 始まれば true
     */
    function startsWithAny(lowerText, namePrefixes) {
        for (var i = 0; i < namePrefixes.length; i++) {
            var prefixLower = String(namePrefixes[i]).toLowerCase();
            if (prefixLower !== "" && lowerText.indexOf(prefixLower) === 0) return true;
        }
        return false;
    }

    /**
     * 小文字の文字列が、一覧のどれかを語の単位で始まるか（直後が末尾か区切り。"Pr5" は "Pr5N" に当たらない）
     * @param {string} lowerText - 小文字の文字列
     * @param {string[]} fontNames - 名前の一覧
     * @returns {boolean} 始まれば true
     */
    function startsWithAnyWord(lowerText, fontNames) {
        for (var i = 0; i < fontNames.length; i++) {
            var nameLower = String(fontNames[i]).toLowerCase();
            if (nameLower === "" || lowerText.indexOf(nameLower) !== 0) continue;
            if (lowerText.length === nameLower.length || NAME_SEPARATOR_CHARS.indexOf(lowerText.charAt(nameLower.length)) !== -1) return true;
        }
        return false;
    }

    /**
     * 小文字の文字列が、一覧のどれかを含むか（大文字小文字は区別しない）
     * @param {string} lowerText - 小文字の文字列
     * @param {string[]} nameParts - 名前の一覧
     * @returns {boolean} 含めば true
     */
    function containsAny(lowerText, nameParts) {
        for (var i = 0; i < nameParts.length; i++) {
            var partLower = String(nameParts[i]).toLowerCase();
            if (partLower !== "" && lowerText.indexOf(partLower) !== -1) return true;
        }
        return false;
    }

    /**
     * フォントのファミリー名か PostScript 名が、一覧のどれかで始まるか（大文字小文字は区別しない）
     * @param {Object} fontInfo - createFontCatalog() 済みのフォント情報
     * @param {string[]} namePrefixes - 名前の一覧
     * @returns {boolean} 始まれば true
     */
    function fontNameStartsWithAny(fontInfo, namePrefixes) {
        return startsWithAny(fontInfo.familyLower, namePrefixes) || startsWithAny(fontInfo.psLower, namePrefixes);
    }

    /**
     * フォントのファミリー名か PostScript 名が、一覧のどれかを語の単位で始まるか
     * @param {Object} fontInfo - createFontCatalog() 済みのフォント情報
     * @param {string[]} fontNames - 名前の一覧
     * @returns {boolean} 始まれば true
     */
    function fontNameStartsWithAnyWord(fontInfo, fontNames) {
        return startsWithAnyWord(fontInfo.familyLower, fontNames) || startsWithAnyWord(fontInfo.psLower, fontNames);
    }

    /**
     * フォントのファミリー名か PostScript 名が、一覧のどれかを含むか
     * @param {Object} fontInfo - createFontCatalog() 済みのフォント情報
     * @param {string[]} nameParts - 名前の一覧
     * @returns {boolean} 含めば true
     */
    function fontNameContainsAny(fontInfo, nameParts) {
        return containsAny(fontInfo.familyLower, nameParts) || containsAny(fontInfo.psLower, nameParts);
    }

    /* アクティベート中の Adobe Fonts のファミリー名（小文字 → true）。初めて要るときに読む / Active Adobe Fonts families, read on first use */
    var adobeFontsFamilies = null;

    /**
     * Creative Cloud の同期状態から、アクティベート中の Adobe Fonts のファミリー名を読む。
     * 和文は Illustrator が日本語名で返すので、ロケール別の名前（<i18n> の中の familyName）も加える。
     * installState が OS のものは同じフォントが OS 側にあって使われていないので外す
     * @returns {Object} 小文字のファミリー名 → true（ファイルが無ければ空）
     */
    function readAdobeFontsFamilies() {
        var families = {};
        var stateText = "";
        for (var i = 0; i < ADOBE_FONTS_STATE_PATHS.length && stateText === ""; i++) {
            stateText = readTextFile(new File(Folder.userData + ADOBE_FONTS_STATE_PATHS[i]));
        }
        var fontBlocks = stateText.split("<font>");
        for (var j = 1; j < fontBlocks.length; j++) {
            var fontBlock = fontBlocks[j];
            if (fontBlock.indexOf("<owner>Adobe Fonts</owner>") === -1 || fontBlock.indexOf("<installState>CC</installState>") === -1) continue;
            var familyPattern = /<familyName>([^<]*)<\/familyName>/g;
            var familyMatch;
            while ((familyMatch = familyPattern.exec(fontBlock)) !== null) {
                families[decodeXmlText(familyMatch[1]).toLowerCase()] = true;
            }
        }
        return families;
    }

    /**
     * XML の文字参照を戻す（&amp; など）
     * @param {string} xmlText - XML の文字列
     * @returns {string} 戻した文字列
     */
    function decodeXmlText(xmlText) {
        return xmlText.replace(/&#x([0-9a-f]+);/gi, function (whole, hexCode) {
            return String.fromCharCode(parseInt(hexCode, 16));
        }).replace(/&#(\d+);/g, function (whole, decimalCode) {
            return String.fromCharCode(Number(decimalCode));
        }).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
    }

    /**
     * メーカー別の分類に当たるか
     * @param {Object} fontInfo - createFontCatalog() 済みのフォント情報
     * @param {Object} foundryFilter - FOUNDRY_FILTERS の1件
     * @returns {boolean} 当たれば true
     */
    function matchesFoundry(fontInfo, foundryFilter) {
        if (foundryFilter.usesAdobeFontsList) {
            if (!adobeFontsFamilies) adobeFontsFamilies = readAdobeFontsFamilies();
            return adobeFontsFamilies[fontInfo.familyLower] === true;
        }
        return fontInfo.foundryMatches[foundryFilter.key] === true;
    }

    /**
     * 名前の前方一致でメーカー別の分類に当たるか（createFontCatalog() で1回だけ呼ぶ）
     * @param {Object} fontInfo - psLower・familyLower を持つフォント情報
     * @param {Object} foundryFilter - FOUNDRY_FILTERS の1件
     * @returns {boolean} 当たれば true
     */
    function matchesFoundryName(fontInfo, foundryFilter) {
        return startsWithAny(fontInfo.familyLower, foundryFilter.familyPrefixes || []) ||
            startsWithAny(fontInfo.psLower, foundryFilter.psNamePrefixes || []);
    }

    /**
     * 保存した設定から、［カスタムセット］の中身とチェックを CUSTOM_SET_COUNT 個そろえて返す。
     * セットを保存していない旧形式なら、［カスタム］のフォント名とチェックをセット1に移す
     * @param {Object} filterSettings - DEFAULT_SETTINGS と同じ形
     * @returns {{fontSets: string[][], visible: boolean[]}} セットごとのフォント名とチェック
     */
    function resolveCustomSets(filterSettings) {
        var savedSets = (filterSettings.customSets instanceof Array) ? filterSettings.customSets : null;
        var savedVisible = (filterSettings.customSetsVisible instanceof Array) ? filterSettings.customSetsVisible : null;
        var fontSets = [];
        var visible = [];
        for (var i = 0; i < CUSTOM_SET_COUNT; i++) {
            if (savedSets) {
                fontSets.push((savedSets[i] instanceof Array) ? savedSets[i] : []);
                visible.push(!!(savedVisible && savedVisible[i] === true));
            } else if (i === 0) {
                fontSets.push((filterSettings.customFonts instanceof Array) ? filterSettings.customFonts : CUSTOM_FONTS);
                visible.push(filterSettings.custom !== false);
            } else {
                fontSets.push([]);
                visible.push(false);
            }
        }
        return { fontSets: fontSets, visible: visible };
    }

    /**
     * 保存形式の設定から、絞り込みの条件を作る
     * @param {Object} filterSettings - DEFAULT_SETTINGS と同じ形
     * @param {string[]} documentFontNames - _ProjectFonts.txt の名前
     * @returns {Object} isFontVisible() の filterState（foundries は FOUNDRY_FILTERS の要素、visibleSuffixes は接尾辞 → true）
     */
    function buildFilterState(filterSettings, documentFontNames) {
        var checkedFoundries = [];
        for (var i = 0; i < FOUNDRY_FILTERS.length; i++) {
            if (containsValue(filterSettings.foundries, FOUNDRY_FILTERS[i].key)) checkedFoundries.push(FOUNDRY_FILTERS[i]);
        }
        var visibleSuffixes = {};
        var suffixKeys = PRIORITY_SUFFIXES.concat([""]);
        for (var j = 0; j < suffixKeys.length; j++) {
            /* ［文字セットで絞り込む］がオフなら、どの文字セットも出す / With Filter by character set off, every character set is shown */
            if (!filterSettings.filterByStandard || !containsValue(filterSettings.hiddenSuffixes, suffixKeys[j])) visibleSuffixes[suffixKeys[j]] = true;
        }
        var excludedLanguages = {};
        if (filterSettings.japaneseOnly) {
            for (var k = 0; k < filterSettings.excludedLanguages.length; k++) excludedLanguages[filterSettings.excludedLanguages[k]] = true;
        }
        /* チェックしたセットの名前をまとめて照合する / Match against the names of all checked sets together */
        var customSets = resolveCustomSets(filterSettings);
        var customFontNames = [];
        for (var m = 0; m < CUSTOM_SET_COUNT; m++) {
            if (customSets.visible[m]) customFontNames = customFontNames.concat(customSets.fontSets[m]);
        }
        return {
            excludedLanguages: excludedLanguages,
            custom: customFontNames.length > 0,
            customFontNames: toLowerNames(customFontNames),
            documentFonts: filterSettings.documentFonts,
            compositeFonts: filterSettings.compositeFonts,
            documentFontNames: toLowerNames(documentFontNames),
            foundries: checkedFoundries,
            showAll: filterSettings.showAll,
            visibleSuffixes: visibleSuffixes,
            searchLower: filterSettings.searchText.replace(/^\s+|\s+$/g, "").toLowerCase(),
            sortByPostScriptName: filterSettings.showPostScriptName
        };
    }

    /**
     * 条件に合うフォントを、ファミリー名順（ファミリー内はインストール順）、PostScript 名で表示するときは PostScript 名順で返す
     * @param {Object[]} fontInfos - createFontCatalog() 済みのフォント情報
     * @param {Object} filterState - buildFilterState() の戻り値
     * @returns {Object[]} 表示するフォント情報
     */
    function filterFontInfos(fontInfos, filterState) {
        var bestByCore = buildBestVariantMap(fontInfos, filterState.visibleSuffixes);
        var sortKeys = [];
        for (var i = 0; i < fontInfos.length; i++) {
            var fontInfo = fontInfos[i];
            if (!isFontVisible(fontInfo, bestByCore, filterState)) continue;
            if (filterState.searchLower !== "" &&
                fontInfo.familyLower.indexOf(filterState.searchLower) === -1 &&
                fontInfo.psLower.indexOf(filterState.searchLower) === -1 &&
                fontInfo.styleLower.indexOf(filterState.searchLower) === -1) continue;
            /* 比較関数つき sort() は遅く並びも狂うので、文字列キーで並べる / Sort by string keys; comparator sorts are slow and unreliable */
            var primaryKey = (filterState.sortByPostScriptName && !fontInfo.isComposite) ? fontInfo.psName : fontInfo.family;
            sortKeys.push(primaryKey + "\u0001" + zeroPad(fontInfo.order, 6) + "\u0001" + i);
        }
        sortKeys.sort();
        var visibleFonts = [];
        for (var j = 0; j < sortKeys.length; j++) {
            visibleFonts.push(fontInfos[Number(sortKeys[j].split("\u0001").pop())]);
        }
        return visibleFonts;
    }

    /**
     * 1件のフォントを表示するか（絞り込みの文字列は除く）
     * @param {Object} fontInfo - createFontCatalog() 済みのフォント情報
     * @param {Object} bestByCore - 中心の名前 → 残す組
     * @param {Object} filterState - buildFilterState() の戻り値
     * @returns {boolean} 表示するなら true
     */
    function isFontVisible(fontInfo, bestByCore, filterState) {
        /* カスタム・ドキュメントに挙げたフォントは、文字セット違いの間引きと除外をせずに出す / Listed Custom and Document fonts skip thinning and exclusions */
        var isListed = (filterState.custom && fontNameStartsWithAny(fontInfo, filterState.customFontNames)) ||
            (filterState.documentFonts && fontNameStartsWithAnyWord(fontInfo, filterState.documentFontNames));
        /* ［日本語以外を除外］の言語は外す（カスタム・ドキュメントに挙げたものは残す）/ Leave out excluded languages, except listed Custom and Document fonts */
        if (!isListed && filterState.excludedLanguages[fontInfo.languageKey] === true) return false;
        /* 合成フォントは、［すべて表示］ではほかと同じく［文字セット］で絞り、それ以外は文字セットを問わない
           Composite fonts follow Character set under Show all like any other font, and ignore it otherwise */
        if (fontInfo.isComposite) {
            if (filterState.showAll) return filterState.visibleSuffixes[fontInfo.parsed.suffix] === true;
            return filterState.compositeFonts || isListed;
        }
        /* ［文字セット］の絞り込みは［すべて表示］でも効く / The Character set filter applies even with Show all */
        if (filterState.visibleSuffixes[fontInfo.parsed.suffix] !== true) return false;
        if (filterState.showAll || isListed) return true;

        var bestVariant = bestByCore[fontInfo.parsed.core];
        if (bestVariant.prefixText !== fontInfo.parsed.prefixText || bestVariant.suffixText !== fontInfo.parsed.suffixText) return false;
        if (fontInfo.isExcludedName) return false;

        for (var i = 0; i < filterState.foundries.length; i++) {
            if (matchesFoundry(fontInfo, filterState.foundries[i])) return true;
        }
        return false;
    }

    /**
     * 並んだフォントのうち、指定した位置と同じファミリーが続く数を返す
     * @param {Object[]} visibleFonts - ファミリーごとに続いたフォント情報
     * @param {number} fontIndex - 位置
     * @returns {number} 同じファミリーの数
     */
    function countFamilyRun(visibleFonts, fontIndex) {
        var familyName = visibleFonts[fontIndex].family;
        var startIndex = fontIndex;
        while (startIndex > 0 && visibleFonts[startIndex - 1].family === familyName) startIndex--;
        var endIndex = fontIndex;
        while (endIndex < visibleFonts.length - 1 && visibleFonts[endIndex + 1].family === familyName) endIndex++;
        return endIndex - startIndex + 1;
    }

    /**
     * 数を指定した桁までゼロで埋める
     * @param {number} value - 数
     * @param {number} digits - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function zeroPad(value, digits) {
        var paddedText = String(value);
        while (paddedText.length < digits) paddedText = "0" + paddedText;
        return paddedText;
    }

    /**
     * 配列に値が含まれるか
     * @param {Array} values - 配列
     * @param {*} targetValue - 探す値
     * @returns {boolean} 含まれれば true
     */
    function containsValue(values, targetValue) {
        for (var i = 0; i < values.length; i++) {
            if (values[i] === targetValue) return true;
        }
        return false;
    }

    // =========================================
    // ファイル / Files
    // =========================================

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string} 中身（無い・読めないときは ""）
     */
    function readTextFile(textFile) {
        if (!textFile.exists) return "";
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) return "";
        var fileText = textFile.read();
        textFile.close();
        return fileText;
    }

    /**
     * テキストファイルを UTF-8 で書く
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function writeTextFile(textFile, fileText) {
        textFile.encoding = "UTF-8";
        textFile.lineFeed = "Unix";
        if (!textFile.open("w")) return false;
        var isWritten = textFile.write(fileText);
        textFile.close();
        return isWritten;
    }

    /**
     * 文字列を行に分ける（改行コードは問わない）
     * @param {string} fileText - 文字列
     * @returns {string[]} 行の配列
     */
    function splitTextLines(fileText) {
        return fileText ? fileText.split(/\r\n|\r|\n/) : [];
    }

    /**
     * _ProjectFonts.txt の名前を読む
     * @param {string} documentFolderPath - ドキュメントのフォルダー（未保存・ドキュメントなしは ""）
     * @returns {string[]} フォント名の一覧（ファイルが無ければ空）
     */
    function readProjectFontNames(documentFolderPath) {
        if (!documentFolderPath) return [];
        var projectFontsFile = new File(documentFolderPath + "/" + PROJECT_FONTS_FILE_NAME);
        var fileLines = splitTextLines(readTextFile(projectFontsFile));
        var fontNames = [];
        for (var i = 0; i < fileLines.length; i++) {
            var fontName = fileLines[i].replace(/^\s+|\s+$/g, "");
            if (fontName !== "") fontNames.push(fontName);
        }
        return fontNames;
    }

    // =========================================
    // メインエンジンで実行する DOM 処理 / DOM work run on the main engine
    //
    // 常駐パレットからは app.activeDocument が読めないので、BridgeTalk でメインエンジンへ渡す。
    // favoriteFontWorker() は toString() で送るため JSDoc を付けず、中は ASCII だけでコメントも書かない。
    // 末尾の目印で切り取る（toString() は前後のコメントまで取り込むことがある）。
    // The palette cannot read app.activeDocument, so this runs on the main engine through BridgeTalk.
    // It is sent with toString(): no JSDoc, ASCII only, no comments inside; it is cut at the end marker.
    //
    // actionId: "state"（ドキュメントの場所と先頭の文字のフォント）/ "apply"（argText のフォントを適用）
    // 戻り値: "OK|…"、"NODOC"、"NOTEXT"、"NOFONT"（改行・タブは BridgeTalk で化けるので使わない）
    // =========================================

    function favoriteFontWorker(actionId, argText) {
        function collectFrames(item, frames) {
            if (!item) return;
            try {
                if (item.locked || item.hidden) return;
            } catch (e) { }
            if (item.typename === "TextFrame") {
                frames.push(item);
                return;
            }
            if (item.typename === "GroupItem") {
                for (var i = 0; i < item.pageItems.length; i++) collectFrames(item.pageItems[i], frames);
            }
        }
        function readTargets(doc) {
            var targets = { textRange: null, frames: [] };
            var docSelection = doc.selection;
            if (!docSelection) return targets;
            if (docSelection.typename === "TextRange") {
                targets.textRange = docSelection;
                return targets;
            }
            for (var i = 0; i < docSelection.length; i++) collectFrames(docSelection[i], targets.frames);
            return targets;
        }
        function readFontName(targets) {
            var targetRange = targets.textRange || (targets.frames.length > 0 ? targets.frames[0].textRange : null);
            if (!targetRange) return "";
            if (targetRange.characters.length > 0) targetRange = targetRange.characters[0];
            try {
                return targetRange.characterAttributes.textFont.name;
            } catch (e) {
                return "";
            }
        }
        if (app.documents.length === 0) return "NODOC";
        var doc = app.activeDocument;
        var targets = readTargets(doc);
        if (actionId === "state") {
            var docPath = "";
            try {
                if (doc.path && doc.path.fsName) docPath = doc.path.fsName;
            } catch (e) { }
            return "OK|" + encodeURIComponent(docPath) + "|" + encodeURIComponent(readFontName(targets));
        }
        if (actionId === "apply") {
            if (!targets.textRange && targets.frames.length === 0) return "NOTEXT";
            var textFont;
            try {
                textFont = app.textFonts.getByName(argText);
            } catch (e) {
                return "NOFONT";
            }
            if (targets.textRange) {
                targets.textRange.characterAttributes.textFont = textFont;
            } else {
                for (var a = 0; a < targets.frames.length; a++) targets.frames[a].textRange.characterAttributes.textFont = textFont;
            }
            app.redraw();
            return "OK";
        }
        return "OK";
        /* @@favorite-font-worker-end@@ */
    }

    /* favoriteFontWorker() の直後にはコメントを置かない（toString() に取り込まれる）。目印 / End marker of the worker source */
    var WORKER_END_MARKER = "/* @@favorite-font-worker-end@@ */";

    /**
     * BridgeTalk で送るワーカーのソースを作る（関数の先頭から目印までを切り出し、閉じ括弧を足す）
     * @returns {string} favoriteFontWorker() の宣言
     */
    function getWorkerSource() {
        var workerSource = favoriteFontWorker.toString().replace(/\r\n?/g, "\n");
        var startIndex = workerSource.indexOf("function favoriteFontWorker");
        var endIndex = workerSource.indexOf(WORKER_END_MARKER);
        return workerSource.substring(startIndex, endIndex) + "\n}";
    }

    /* 送るソースは起動時に1回だけ作る / Built once at startup */
    var WORKER_SOURCE = getWorkerSource();

    /**
     * メインエンジンで favoriteFontWorker() を動かす（非同期）。
     * BridgeTalk は本文の \ を二重にするので、コード全体を encodeURIComponent で包んで送る
     * @param {string} actionId - "state" / "apply"
     * @param {string} argText - ワーカーに渡す文字列（PostScript 名など）
     * @param {function(string): void} [onDone] - 戻り値を受け取る処理
     * @returns {void}
     */
    function runWorker(actionId, argText, onDone) {
        var workerCode = WORKER_SOURCE + "\nvar __favoriteFontResult = favoriteFontWorker(\"" + actionId + "\", decodeURIComponent(\"" + encodeURIComponent(argText || "") + "\"));\n__favoriteFontResult;";
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerCode) + "\"));";
        bridge.onResult = function (response) {
            if (onDone) onDone(String(response.body || ""));
        };
        bridge.onError = function (response) {
            alert(getLabel("alert.workerFailed", { action: actionId, message: (response && response.body) ? response.body : "BridgeTalk error" }));
        };
        bridge.send();
    }

    // =========================================
    // キーボードショートカット / Keyboard shortcuts
    // =========================================

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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 左の列（分類・すべて表示・文字セットのチェックボックス）を作る
     * @param {Group} parentGroup - 列を足す先
     * @returns {Object} コントロール一式
     */
    function buildFilterColumn(parentGroup) {
        var filterColumn = parentGroup.add("group");
        filterColumn.orientation = "column";
        filterColumn.alignChildren = ["fill", "top"];
        filterColumn.alignment = ["left", "fill"];
        filterColumn.spacing = WINDOW_SPACING;

        /* ［分類］全体を切り替える親スイッチなので、その上に置く / The master switch for Category, placed above it */
        var chkShowAll = addCheckbox(filterColumn, getLabel("checkbox.showAll"), "tooltip.showAll");
        var chkJapaneseOnly = addCheckbox(filterColumn, getLabel("checkbox.japaneseOnly"), "tooltip.japaneseOnly");

        var categoryPanel = filterColumn.add("panel", undefined, getLabel("panel.category"));
        setupPanel(categoryPanel, 6);
        /* カスタムセットは見出しの下に番号を1行に並べる / Custom sets: a heading, then the numbers on one row */
        categoryPanel.add("statictext", undefined, getLabel("fieldLabel.customSets"));
        var customSetRowGroup = categoryPanel.add("group");
        setupRow(customSetRowGroup, "left", 4);
        var customSetCheckboxes = [];
        for (var c = 0; c < CUSTOM_SET_COUNT; c++) {
            var chkCustomSet = customSetRowGroup.add("checkbox", undefined, String(c + 1));
            chkCustomSet.helpTip = getLabel("tooltip.customSet", { number: c + 1 });
            chkCustomSet.setIndex = c;
            customSetCheckboxes.push(chkCustomSet);
        }
        var customSetDivider = categoryPanel.add("panel");
        customSetDivider.alignment = ["fill", "top"];
        customSetDivider.minimumSize.height = customSetDivider.maximumSize.height = 1;
        var chkDocumentFonts = addCheckbox(categoryPanel, getLabel("checkbox.documentFonts"), "tooltip.documentFonts");
        var chkCompositeFonts = addCheckbox(categoryPanel, getLabel("checkbox.compositeFonts"), "tooltip.compositeFonts");
        var foundryCheckboxes = [];
        /* Adobe Fonts はメーカーではなく入手経路なので、線より上に置く / Adobe Fonts is a source, not a foundry, so it goes above the line */
        addFoundryCheckboxes(categoryPanel, foundryCheckboxes, true);
        /* 自分で決める分類と、メーカー別の分類を線で分ける / A line between your own categories and the foundry ones */
        var categoryDivider = categoryPanel.add("panel");
        categoryDivider.alignment = ["fill", "top"];
        categoryDivider.minimumSize.height = categoryDivider.maximumSize.height = 1;
        addFoundryCheckboxes(categoryPanel, foundryCheckboxes, false);

        var standardPanel = filterColumn.add("panel", undefined, getLabel("panel.standard"));
        setupPanel(standardPanel, 6);
        var chkFilterByStandard = addCheckbox(standardPanel, getLabel("checkbox.filterByStandard"), "tooltip.filterByStandard");
        var standardDivider = standardPanel.add("panel");
        standardDivider.alignment = ["fill", "top"];
        standardDivider.minimumSize.height = standardDivider.maximumSize.height = 1;
        /* ［文字セットで絞り込む］がオフのときに無効にする部分 / The part disabled while Filter by character set is off */
        var suffixGroup = standardPanel.add("group");
        suffixGroup.orientation = "column";
        suffixGroup.alignChildren = ["fill", "top"];
        suffixGroup.spacing = 6;
        suffixGroup.margins = 0;
        var suffixCheckboxes = [];
        var suffixRows = buildSuffixRows(PRIORITY_SUFFIXES);
        for (var j = 0; j < suffixRows.length; j++) {
            var suffixRowGroup = suffixGroup.add("group");
            setupRow(suffixRowGroup, "fill", 0);
            for (var k = 0; k < 2; k++) {
                var rowSuffix = suffixRows[j][k];
                if (!rowSuffix) {
                    /* 対になる文字セットが無いときは空けて列をそろえる / Leave a gap so the columns stay aligned */
                    if (k === 0) suffixRowGroup.add("group").preferredSize.width = STANDARD_LEFT_WIDTH;
                    continue;
                }
                var chkSuffix = addCheckbox(suffixRowGroup, rowSuffix, "tooltip.standard");
                if (k === 0) chkSuffix.preferredSize.width = STANDARD_LEFT_WIDTH;
                chkSuffix.suffix = rowSuffix;
                suffixCheckboxes.push(chkSuffix);
            }
        }
        var chkNoSuffix = addCheckbox(suffixGroup, getLabel("checkbox.noSuffix"), "tooltip.noSuffix");
        chkNoSuffix.suffix = "";
        suffixCheckboxes.push(chkNoSuffix);

        return {
            categoryPanel: categoryPanel,
            customSetCheckboxes: customSetCheckboxes,
            chkDocumentFonts: chkDocumentFonts,
            chkCompositeFonts: chkCompositeFonts,
            foundryCheckboxes: foundryCheckboxes,
            chkShowAll: chkShowAll,
            chkFilterByStandard: chkFilterByStandard,
            suffixGroup: suffixGroup,
            suffixCheckboxes: suffixCheckboxes,
            chkJapaneseOnly: chkJapaneseOnly
        };
    }

    /**
     * FOUNDRY_FILTERS のチェックボックスを足す（Adobe Fonts の一覧で判定するものか、それ以外かで分けて呼ぶ）
     * @param {Panel} parentPanel - 足す先
     * @param {Checkbox[]} foundryCheckboxes - 足したチェックボックスを加える配列
     * @param {boolean} usesAdobeFontsList - true なら Adobe Fonts の一覧で判定するものだけ、false ならそれ以外
     * @returns {void}
     */
    function addFoundryCheckboxes(parentPanel, foundryCheckboxes, usesAdobeFontsList) {
        for (var i = 0; i < FOUNDRY_FILTERS.length; i++) {
            if (!!FOUNDRY_FILTERS[i].usesAdobeFontsList !== usesAdobeFontsList) continue;
            var chkFoundry = addCheckbox(parentPanel, getLabel(FOUNDRY_FILTERS[i].label));
            chkFoundry.helpTip = getLabel(FOUNDRY_FILTERS[i].tooltip);
            chkFoundry.foundryKey = FOUNDRY_FILTERS[i].key;
            foundryCheckboxes.push(chkFoundry);
        }
    }

    /**
     * チェックボックスの組に、option（Alt）＋クリックの切り替えを付ける。
     * ほかがすべてオフ（クリックしたものだけオン）なら全部オン、それ以外ならクリックしたものだけオンにする
     * @param {Checkbox[]} checkboxGroup - 組にするチェックボックス
     * @param {function(): void} onGroupChange - クリックのたびに呼ぶ処理
     * @returns {void}
     */
    function addSoloClickGroup(checkboxGroup, onGroupChange) {
        /**
         * クリックされたチェックボックスを処理する
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @returns {void}
         */
        function handleClick(clickedCheckbox) {
            if (ScriptUI.environment.keyboardState.altKey) {
                /* onClick の時点で、クリックしたものの値はもう反転している / The clicked box is already toggled when onClick runs */
                var isOthersOff = true;
                for (var i = 0; i < checkboxGroup.length; i++) {
                    if (checkboxGroup[i] !== clickedCheckbox && checkboxGroup[i].value) isOthersOff = false;
                }
                for (var j = 0; j < checkboxGroup.length; j++) {
                    checkboxGroup[j].value = isOthersOff || checkboxGroup[j] === clickedCheckbox;
                }
            }
            onGroupChange();
        }

        for (var k = 0; k < checkboxGroup.length; k++) {
            var groupCheckbox = checkboxGroup[k];
            groupCheckbox.onClick = function () {
                handleClick(this);
            };
            groupCheckbox.helpTip = (groupCheckbox.helpTip ? groupCheckbox.helpTip + "\n" : "") + getLabel("tooltip.soloClick");
        }
    }

    /**
     * ［文字セット］の並びを作る。左に無印、右に N 付き（Pro と ProN など）を置き、上から優先順位の低い順に並べる
     * @param {string[]} prioritySuffixes - PRIORITY_SUFFIXES（優先順位の高い順）
     * @returns {Array} 行の配列。各行は [無印, N 付き]（無いほうは null）
     */
    function buildSuffixRows(prioritySuffixes) {
        var suffixRows = [];
        var rowByBase = {};
        for (var i = prioritySuffixes.length - 1; i >= 0; i--) {
            var suffix = prioritySuffixes[i];
            var isNVariant = /N$/.test(suffix);
            var baseName = isNVariant ? suffix.slice(0, -1) : suffix;
            if (!rowByBase["_" + baseName]) {
                rowByBase["_" + baseName] = [null, null];
                suffixRows.push(rowByBase["_" + baseName]);
            }
            rowByBase["_" + baseName][isNVariant ? 1 : 0] = suffix;
        }
        return suffixRows;
    }

    /**
     * 自前描画のボタンの色を返す（押している間は濃く、無効のときは淡く）
     * @param {Button} drawnButton - 対象のボタン
     * @returns {number[]} RGBA
     */
    function getDrawnButtonColor(drawnButton) {
        var useDarkUI = UI_IS_DARK;
        /* 無効にしてあるときは、さらに淡くする / Dim further while disabled */
        if (drawnButton.enabled === false) return useDarkUI ? [0.42, 0.42, 0.42, 1] : [0.78, 0.78, 0.78, 1];
        /* 押している間は濃くして手応えを出す / Stronger while pressed */
        if (drawnButton.pressed) return useDarkUI ? [0.95, 0.95, 0.95, 1] : [0.20, 0.20, 0.20, 1];
        /* 主役ではないので、通常時も少し薄めに描く / Slightly muted by default */
        return useDarkUI ? [0.72, 0.72, 0.72, 1] : [0.45, 0.45, 0.45, 1];
    }

    /**
     * 前回の描画を消すため、ボタンを親と同じ色で塗りつぶす（背景色のブラシを持たない環境では省く）
     * @param {Button} drawnButton - 対象のボタン
     * @returns {void}
     */
    function clearDrawnButton(drawnButton) {
        var buttonGraphics = drawnButton.graphics;
        try {
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, drawnButton.size[0], drawnButton.size[1]);
            buttonGraphics.fillPath(buttonGraphics.backgroundColor);
        } catch (e) {}
    }

    /**
     * 中身を自前で描く正方形のボタンを足す（押している間の表示も付ける）
     * @param {Group} parentGroup - 足す先
     * @param {number} buttonSize - 一辺（px）
     * @param {function(Button): void} drawButton - 描く処理
     * @returns {Button} 足したボタン
     */
    function addDrawnButton(parentGroup, buttonSize, drawButton) {
        var drawnButton = parentGroup.add("button", undefined, "");
        drawnButton.preferredSize = [buttonSize, buttonSize];
        drawnButton.minimumSize = [buttonSize, buttonSize];
        drawnButton.maximumSize = [buttonSize, buttonSize];
        drawnButton.pressed = false;

        drawnButton.onDraw = function () {
            drawButton(this);
        };
        drawnButton.addEventListener("mousedown", function () {
            this.pressed = true;
            this.notify("onDraw");
        });
        drawnButton.addEventListener("mouseup", function () {
            this.pressed = false;
            this.notify("onDraw");
        });
        /* ボタンの外でマウスを離したときも押下表示を戻す / Reset the pressed look on mouseout */
        drawnButton.addEventListener("mouseout", function () {
            if (!this.pressed) return;
            this.pressed = false;
            this.notify("onDraw");
        });
        return drawnButton;
    }

    /**
     * クリアボタンの丸と×を描く
     * @param {Button} clearButton - 対象のボタン
     * @returns {void}
     */
    function drawClearButton(clearButton) {
        var buttonGraphics = clearButton.graphics;
        var buttonSize = clearButton.size[0];
        clearDrawnButton(clearButton);

        var glyphPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, getDrawnButtonColor(clearButton), CLEAR_STROKE_WIDTH);
        var circleSize = buttonSize - CLEAR_CIRCLE_INSET * 2;
        buttonGraphics.newPath();
        buttonGraphics.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
        buttonGraphics.strokePath(glyphPen);

        var glyphStart = CLEAR_GLYPH_INSET;
        var glyphEnd = buttonSize - CLEAR_GLYPH_INSET;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(glyphStart, glyphStart);
        buttonGraphics.lineTo(glyphEnd, glyphEnd);
        buttonGraphics.strokePath(glyphPen);
        buttonGraphics.newPath();
        buttonGraphics.moveTo(glyphEnd, glyphStart);
        buttonGraphics.lineTo(glyphStart, glyphEnd);
        buttonGraphics.strokePath(glyphPen);
    }

    /**
     * 絞り込みの文字列を消すクリアボタンを足す（有効・無効は呼び出し側が入力欄に合わせる）
     * @param {Group} parentGroup - 足す先
     * @returns {Button} 丸に×を自前描画したボタン
     */
    function addClearButton(parentGroup) {
        var clearButton = addDrawnButton(parentGroup, CLEAR_BUTTON_SIZE, drawClearButton);
        clearButton.helpTip = getLabel("tooltip.clearSearch");
        clearButton.alignment = ["right", "center"];
        clearButton.enabled = false;
        return clearButton;
    }

    /**
     * しおりを描く。塗るときは onDraw では多角形を塗れないので、切り込みの部分は1px の横長の四角を積んで塗る
     * @param {ScriptUIGraphics} buttonGraphics - 描く先
     * @param {number[]} glyphColor - 色
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {boolean} isFilled - true なら塗り、false なら輪郭
     * @returns {void}
     */
    function drawBookmark(buttonGraphics, glyphColor, left, top, isFilled) {
        var right = left + SET_BOOKMARK_WIDTH;
        var bottom = top + SET_BOOKMARK_HEIGHT;
        var notchTop = bottom - SET_BOOKMARK_NOTCH;
        if (!isFilled) {
            var outlinePen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, glyphColor, 1.5);
            buttonGraphics.newPath();
            buttonGraphics.moveTo(left, top);
            buttonGraphics.lineTo(right, top);
            buttonGraphics.lineTo(right, bottom);
            buttonGraphics.lineTo(left + SET_BOOKMARK_WIDTH / 2, notchTop);
            buttonGraphics.lineTo(left, bottom);
            buttonGraphics.lineTo(left, top);
            buttonGraphics.strokePath(outlinePen);
            return;
        }
        var fillBrush = buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, glyphColor);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(left, top, SET_BOOKMARK_WIDTH, SET_BOOKMARK_HEIGHT - SET_BOOKMARK_NOTCH);
        buttonGraphics.fillPath(fillBrush);
        for (var y = notchTop; y < bottom; y++) {
            /* 切り込みは下へ行くほど広がる / The notch widens toward the bottom */
            var legWidth = SET_BOOKMARK_WIDTH / 2 - (y - notchTop + 0.5) * (SET_BOOKMARK_WIDTH / 2) / SET_BOOKMARK_NOTCH;
            if (legWidth <= 0) break;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(left, y, legWidth, 1);
            buttonGraphics.fillPath(fillBrush);
            buttonGraphics.newPath();
            buttonGraphics.rectPath(right - legWidth, y, legWidth, 1);
            buttonGraphics.fillPath(fillBrush);
        }
    }

    /**
     * ［カスタムセット］の出し入れボタンを描く。重なった2枚のカードの手前に番号としおり。
     * 選んだフォントがそのセットに入っていればアイコン全体を暗くしてしおりを塗り、外れていればしおりは輪郭だけ
     * @param {Button} setButton - 対象のボタン（setNumber に番号、isOn が true なら塗る）
     * @returns {void}
     */
    function drawCustomSetButton(setButton) {
        var buttonGraphics = setButton.graphics;
        clearDrawnButton(setButton);
        var glyphColor = getDrawnButtonColor(setButton);
        /* セットに入っているときは、アイコン全体を暗く（ダーク表示では明るく）してはっきり見せる / In the set, the whole icon gets the strongest contrast */
        if (setButton.isOn && setButton.enabled !== false) glyphColor = UI_IS_DARK ? [0.95, 0.95, 0.95, 1] : [0.10, 0.10, 0.10, 1];
        var cardPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, glyphColor, SET_CARD_STROKE_WIDTH);
        var backLeft = Math.round((setButton.size[0] - SET_CARD_SIZE - SET_CARD_OFFSET) / 2);
        var backTop = Math.round((setButton.size[1] - SET_CARD_SIZE - SET_CARD_OFFSET) / 2);
        var frontLeft = backLeft + SET_CARD_OFFSET;
        var frontTop = backTop + SET_CARD_OFFSET;

        /* 後ろのカードは、手前のカードに隠れない部分の線だけを描く（ボタンでは背景色で塗って隠せない）
           Only the visible part of the back card is drawn; buttons cannot hide it by painting the background */
        var backRight = backLeft + SET_CARD_SIZE;
        var backBottom = backTop + SET_CARD_SIZE;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backBottom);
        buttonGraphics.lineTo(backLeft, backBottom);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backRight, backTop);
        buttonGraphics.lineTo(backRight, frontTop);
        buttonGraphics.strokePath(cardPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, SET_CARD_SIZE, SET_CARD_SIZE);
        buttonGraphics.strokePath(cardPen);

        var bookmarkLeft = frontLeft + SET_CARD_SIZE - SET_BOOKMARK_INSET - SET_BOOKMARK_WIDTH;
        drawBookmark(buttonGraphics, glyphColor, bookmarkLeft, frontTop, setButton.isOn);

        var numberFont = ScriptUI.newFont("dialog", "BOLD", SET_NUMBER_FONT_SIZE);
        var numberText = String(setButton.setNumber);
        var numberSize = buttonGraphics.measureString(numberText, numberFont);
        var numberPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, glyphColor, 1);
        /* 番号はしおりの左の、カードの下寄りに置く / The number sits left of the bookmark, low on the card */
        var numberLeft = frontLeft + Math.round((bookmarkLeft - frontLeft - numberSize[0]) / 2);
        var numberTop = frontTop + SET_CARD_SIZE - numberSize[1] - 2;
        buttonGraphics.drawString(numberText, numberPen, numberLeft, numberTop, numberFont);
    }

    /**
     * ［カスタムセット］の出し入れボタンを足す（状態とツールチップは呼び出し側が選択に合わせる）
     * @param {Group} parentGroup - 足す先
     * @param {number} setNumber - セットの番号（1 から）
     * @returns {Button} 自前描画したボタン
     */
    function addCustomSetButton(parentGroup, setNumber) {
        var setButton = addDrawnButton(parentGroup, SET_BUTTON_SIZE, drawCustomSetButton);
        setButton.setNumber = setNumber;
        setButton.isOn = false;
        setButton.enabled = false;
        return setButton;
    }

    /**
     * チェックボックスを足す
     * @param {Panel|Group} parentContainer - 足す先
     * @param {string} checkboxText - 表示する文言
     * @param {string|null} tooltipPath - ツールチップの LABELS のパス（無ければ null）
     * @returns {Checkbox} 足したチェックボックス
     */
    function addCheckbox(parentContainer, checkboxText, tooltipPath) {
        var checkbox = parentContainer.add("checkbox", undefined, checkboxText);
        if (tooltipPath) checkbox.helpTip = getLabel(tooltipPath);
        return checkbox;
    }

    /**
     * 右の列（絞り込み欄・フォント一覧・件数）を作る
     * @param {Group} parentGroup - 列を足す先
     * @returns {Object} コントロール一式（searchInput・btnClearSearch・fontListBox・customSetButtons・chkShowPostScriptName・statusText）
     */
    function buildFontListColumn(parentGroup) {
        var fontListColumn = parentGroup.add("group");
        fontListColumn.orientation = "column";
        fontListColumn.alignChildren = ["fill", "top"];
        /* ダイアログを縦に広げたら一覧を伸ばす / The list grows when the dialog is made taller */
        fontListColumn.alignment = ["fill", "fill"];
        fontListColumn.spacing = 6;

        var searchRowGroup = fontListColumn.add("group");
        /* 右の列は縦に伸びるので、行は上詰めにする（center だと余った高さの中で下へずれる）/ The column stretches, so rows stick to the top */
        setupRow(searchRowGroup, ["fill", "top"], 6);
        searchRowGroup.margins = [0, 0, 0, FONT_LIST_TOP_MARGIN];
        searchRowGroup.add("statictext", undefined, labelText("fieldLabel.search"));
        var searchInput = searchRowGroup.add("edittext", undefined, "");
        searchInput.alignment = ["fill", "center"];
        searchInput.helpTip = getLabel("tooltip.search");
        var btnClearSearch = addClearButton(searchRowGroup);

        /* 一覧の見出しの行：左に件数、右に［PostScript 名で表示］/ List header: the count on the left, Show PostScript names on the right */
        var listHeaderRowGroup = fontListColumn.add("group");
        setupRow(listHeaderRowGroup, ["fill", "top"], 6);
        var statusText = listHeaderRowGroup.add("statictext", undefined, "");
        /* 件数は空いた幅をすべて使い、スペーサーを兼ねる（空のまま作ると幅が 0 になるので fill で広げる）
           The count takes the spare width and doubles as the spacer (fill keeps it from collapsing while empty) */
        statusText.alignment = ["fill", "center"];
        var chkShowPostScriptName = addCheckbox(listHeaderRowGroup, getLabel("checkbox.showPostScriptName"), "tooltip.showPostScriptName");
        chkShowPostScriptName.alignment = ["right", "center"];

        var fontListBox = fontListColumn.add("listbox", undefined, [], { multiselect: false });
        fontListBox.preferredSize = FONT_LIST_SIZE;
        fontListBox.alignment = ["fill", "fill"];
        fontListBox.helpTip = getLabel("tooltip.fontList");

        var customRowGroup = fontListColumn.add("group");
        /* 番号のボタンは列の左右中央に / The set buttons sit centered in the column */
        setupRow(customRowGroup, ["center", "top"], 6);
        customRowGroup.margins = [0, FONT_LIST_BOTTOM_MARGIN, 0, 0];
        var customSetButtons = [];
        for (var n = 0; n < CUSTOM_SET_COUNT; n++) customSetButtons.push(addCustomSetButton(customRowGroup, n + 1));

        return {
            searchInput: searchInput,
            btnClearSearch: btnClearSearch,
            fontListBox: fontListBox,
            customSetButtons: customSetButtons,
            chkShowPostScriptName: chkShowPostScriptName,
            statusText: statusText
        };
    }

    /**
     * 環境設定のダイアログを開く（再スキャン・［和文フォントのみ］で外す言語・カスタムセットの書き出しと読み込み）。
     * 再スキャンはその場で行い、キャンセルしても戻らない
     * @param {{fontTotal: number, excludedLanguages: string[], customFontSets: string[][], onRescan: function(): number}} preferenceState - 今の値と、再スキャンの処理（読み込んだ件数を返す。読めなかったら -1）
     * @returns {{excludedLanguages: string[], customFontSets: string[][]}|null} OK なら新しい値、キャンセルなら null
     */
    function showPreferencesDialog(preferenceState) {
        var customFontSets = [];
        for (var i = 0; i < preferenceState.customFontSets.length; i++) customFontSets.push(preferenceState.customFontSets[i].slice(0));

        var prefsDialog = new Window("dialog", getLabel("dialog.preferences"));
        setupWindow(prefsDialog);

        var fontInfoPanel = prefsDialog.add("panel", undefined, getLabel("panel.fontInfo"));
        setupPanel(fontInfoPanel, 6);
        var fontInfoRowGroup = fontInfoPanel.add("group");
        setupRow(fontInfoRowGroup, "fill", 6);
        var fontTotalText = fontInfoRowGroup.add("statictext", undefined, getLabel("message.fontTotal", { count: preferenceState.fontTotal }));
        fontTotalText.alignment = ["fill", "center"];
        var btnRescan = fontInfoRowGroup.add("button", undefined, getLabel("button.rescan"));
        btnRescan.helpTip = getLabel("tooltip.rescan");

        var languagePanel = prefsDialog.add("panel", undefined, getLabel("panel.excludedLanguages"));
        setupPanel(languagePanel, 6);
        var languageCheckboxes = [];
        /* 2列に並べる / Two columns */
        var languageRowGroup = null;
        for (var m = 0; m < LANGUAGE_GROUPS.length; m++) {
            if (m % 2 === 0) {
                languageRowGroup = languagePanel.add("group");
                setupRow(languageRowGroup, "fill", 0);
            }
            var chkLanguage = languageRowGroup.add("checkbox", undefined, getLabel(LANGUAGE_GROUPS[m].label));
            if (m % 2 === 0) chkLanguage.preferredSize.width = LANGUAGE_LEFT_WIDTH;
            chkLanguage.helpTip = getLabel(LANGUAGE_GROUPS[m].tooltip);
            chkLanguage.languageKey = LANGUAGE_GROUPS[m].key;
            chkLanguage.value = containsValue(preferenceState.excludedLanguages, chkLanguage.languageKey);
            languageCheckboxes.push(chkLanguage);
        }
        addSoloClickGroup(languageCheckboxes, function () {});

        var customSetPanel = prefsDialog.add("panel", undefined, getLabel("panel.customSets"));
        setupPanel(customSetPanel, 6);
        var setNumbers = [];
        for (var n = 0; n < customFontSets.length; n++) setNumbers.push(String(n + 1));
        var customSetRowGroup = customSetPanel.add("group");
        setupRow(customSetRowGroup, "fill", 6);
        var setDropdown = customSetRowGroup.add("dropdownlist", undefined, setNumbers);
        setDropdown.selection = 0;
        var setCountText = customSetRowGroup.add("statictext", undefined, "");
        setCountText.alignment = ["fill", "center"];
        var setListBox = customSetPanel.add("listbox", undefined, [], { multiselect: false });
        setListBox.preferredSize = CUSTOM_SET_LIST_SIZE;
        var setButtonRowGroup = customSetPanel.add("group");
        setupRow(setButtonRowGroup, "left", 6);
        var btnExportSet = setButtonRowGroup.add("button", undefined, getLabel("button.exportSet"));
        btnExportSet.helpTip = getLabel("tooltip.exportSet");
        var btnImportSet = setButtonRowGroup.add("button", undefined, getLabel("button.importSet"));
        btnImportSet.helpTip = getLabel("tooltip.importSet");

        var buttonRow = addButtonRow(prefsDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * ドロップダウンで選んだセットの中身を一覧に並べる
         * @returns {void}
         */
        function showSelectedSet() {
            var setFontNames = customFontSets[setDropdown.selection.index];
            setListBox.removeAll();
            for (var j = 0; j < setFontNames.length; j++) setListBox.add("item", setFontNames[j]);
            setCountText.text = getLabel("message.setCount", { count: setFontNames.length });
        }

        setDropdown.onChange = showSelectedSet;
        btnRescan.onClick = function () {
            var fontTotal = preferenceState.onRescan();
            if (fontTotal < 0) return;
            fontTotalText.text = getLabel("message.fontTotal", { count: fontTotal });
            alert(getLabel("alert.rescanned", { count: fontTotal }));
        };
        btnExportSet.onClick = function () {
            var setIndex = setDropdown.selection.index;
            var exportFile = File.saveDialog(getLabel("dialog.exportSet", { number: setIndex + 1 }), "*.txt");
            if (!exportFile) return;
            if (!/\.txt$/i.test(exportFile.name)) exportFile = new File(exportFile.fsName + ".txt");
            if (!writeTextFile(exportFile, customFontSets[setIndex].join("\n"))) {
                alert(getLabel("alert.writeFailed", { message: exportFile.error }));
                return;
            }
            alert(getLabel("alert.exported", { count: customFontSets[setIndex].length }));
        };
        btnImportSet.onClick = function () {
            var setIndex = setDropdown.selection.index;
            var importFile = File.openDialog(getLabel("dialog.importSet", { number: setIndex + 1 }));
            if (!importFile) return;
            var fileLines = splitTextLines(readTextFile(importFile));
            var addedCount = 0;
            for (var k = 0; k < fileLines.length; k++) {
                var fontName = fileLines[k].replace(/^\s+|\s+$/g, "");
                if (fontName === "" || containsValue(customFontSets[setIndex], fontName)) continue;
                customFontSets[setIndex].push(fontName);
                addedCount++;
            }
            showSelectedSet();
            alert(getLabel("alert.imported", { count: addedCount }));
        };

        showSelectedSet();
        prepareDialogWindow(prefsDialog, SCRIPT_NAME + "Preferences");
        if (prefsDialog.show() !== 1) return null;

        var excludedLanguages = [];
        for (var c = 0; c < languageCheckboxes.length; c++) {
            if (languageCheckboxes[c].value) excludedLanguages.push(languageCheckboxes[c].languageKey);
        }
        return { excludedLanguages: excludedLanguages, customFontSets: customFontSets };
    }

    /**
     * 保存形式の設定をダイアログに反映する
     * @param {Object} filterControls - buildFilterColumn() の戻り値に searchInput・chkShowPostScriptName・customFontSets・excludedLanguages を足したもの
     * @param {Object} filterSettings - DEFAULT_SETTINGS と同じ形
     * @returns {void}
     */
    function applySettingsToControls(filterControls, filterSettings) {
        var customSets = resolveCustomSets(filterSettings);
        filterControls.customFontSets = customSets.fontSets;
        for (var c = 0; c < filterControls.customSetCheckboxes.length; c++) {
            filterControls.customSetCheckboxes[c].value = customSets.visible[c];
        }
        filterControls.chkDocumentFonts.value = filterSettings.documentFonts;
        filterControls.chkCompositeFonts.value = filterSettings.compositeFonts;
        for (var i = 0; i < filterControls.foundryCheckboxes.length; i++) {
            var chkFoundry = filterControls.foundryCheckboxes[i];
            chkFoundry.value = containsValue(filterSettings.foundries, chkFoundry.foundryKey);
        }
        filterControls.chkShowAll.value = filterSettings.showAll;
        filterControls.chkFilterByStandard.value = filterSettings.filterByStandard;
        filterControls.chkJapaneseOnly.value = filterSettings.japaneseOnly;
        filterControls.excludedLanguages = filterSettings.excludedLanguages;
        for (var j = 0; j < filterControls.suffixCheckboxes.length; j++) {
            var chkSuffix = filterControls.suffixCheckboxes[j];
            chkSuffix.value = !containsValue(filterSettings.hiddenSuffixes, chkSuffix.suffix);
        }
        filterControls.searchInput.text = filterSettings.searchText;
        filterControls.chkShowPostScriptName.value = filterSettings.showPostScriptName;
    }

    /**
     * ダイアログの今の状態を保存形式の設定にする
     * @param {Object} filterControls - applySettingsToControls() と同じ
     * @returns {Object} DEFAULT_SETTINGS と同じ形
     */
    function readSettingsFromControls(filterControls) {
        var customSetsVisible = [];
        for (var c = 0; c < filterControls.customSetCheckboxes.length; c++) customSetsVisible.push(filterControls.customSetCheckboxes[c].value);
        var checkedFoundryKeys = [];
        for (var i = 0; i < filterControls.foundryCheckboxes.length; i++) {
            var chkFoundry = filterControls.foundryCheckboxes[i];
            if (chkFoundry.value) checkedFoundryKeys.push(chkFoundry.foundryKey);
        }
        var hiddenSuffixes = [];
        for (var j = 0; j < filterControls.suffixCheckboxes.length; j++) {
            var chkSuffix = filterControls.suffixCheckboxes[j];
            if (!chkSuffix.value) hiddenSuffixes.push(chkSuffix.suffix);
        }
        return {
            japaneseOnly: filterControls.chkJapaneseOnly.value,
            excludedLanguages: filterControls.excludedLanguages,
            customSets: filterControls.customFontSets,
            customSetsVisible: customSetsVisible,
            documentFonts: filterControls.chkDocumentFonts.value,
            compositeFonts: filterControls.chkCompositeFonts.value,
            foundries: checkedFoundryKeys,
            showAll: filterControls.chkShowAll.value,
            filterByStandard: filterControls.chkFilterByStandard.value,
            hiddenSuffixes: hiddenSuffixes,
            searchText: filterControls.searchInput.text,
            showPostScriptName: filterControls.chkShowPostScriptName.value
        };
    }

    /**
     * 常駐パレットを作って表示する
     * @param {Object[]} fontInfos - フォント情報
     * @returns {void}
     */
    function showFavoriteFontPalette(fontInfos) {
        var fontCatalog = createFontCatalog(fontInfos);
        /* ドキュメントの場所と先頭の文字のフォントは、表示とアクティブになるたびにメインエンジンから取り直す
           The document folder and the first character's font are re-read from the main engine on show and on activation */
        var documentFolderPath = "";
        var documentFontNames = [];
        var currentFontName = null;

        var mainPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: true });
        setupWindow(mainPalette);

        var columnsGroup = mainPalette.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.alignment = ["fill", "fill"];
        columnsGroup.spacing = COLUMN_SPACING;

        var filterControls = buildFilterColumn(columnsGroup);
        var fontListControls = buildFontListColumn(columnsGroup);
        var fontListBox = fontListControls.fontListBox;
        filterControls.searchInput = fontListControls.searchInput;
        filterControls.chkShowPostScriptName = fontListControls.chkShowPostScriptName;

        var buttonRow = addButtonRow(mainPalette);
        var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
        btnPreferences.helpTip = getLabel("tooltip.preferences");
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("button.close"));
        alignRightOnlyButtonRow(buttonRow);

        applySettingsToControls(filterControls, settingsStore.load(DEFAULT_SETTINGS));

        /**
         * ［すべて表示］のときは［分類］を、［文字セットで絞り込む］がオフのときは［文字セット］を無効にする
         * @returns {void}
         */
        function updatePanelStates() {
            filterControls.categoryPanel.enabled = !filterControls.chkShowAll.value;
            filterControls.suffixGroup.enabled = filterControls.chkFilterByStandard.value;
        }

        /* 絞り込み欄の内容をまだ一覧に反映していない / The filter text is not reflected in the list yet */
        var isSearchPending = false;
        /* 一覧を作り直している最中（onChange を無視する）/ The list is being rebuilt (ignore onChange) */
        var isUpdatingList = false;
        /* 直前に選んでいた行（見出しを飛ばす向きを決める）/ Previously selected row, for the direction of skipping headers */
        var lastSelectedIndex = -1;

        /**
         * 今の条件に合うフォントを返す
         * @returns {Object[]} 表示するフォント情報
         */
        function collectVisibleFonts() {
            return filterFontInfos(fontCatalog, buildFilterState(readSettingsFromControls(filterControls), documentFontNames));
        }

        /**
         * フォント一覧を作り直す
         * @returns {void}
         */
        function refreshFontList() {
            showFontList(collectVisibleFonts());
        }

        /**
         * フォント一覧に並べる（選んでいたフォントは残っていれば選び直す）。
         * スタイルが複数あるファミリーは、ファミリー名の見出し行の下にウェイトを字下げして並べる
         * @param {Object[]} visibleFonts - 表示するフォント情報（ファミリーごとに続いている）
         * @returns {void}
         */
        function showFontList(visibleFonts) {
            isSearchPending = false;
            var selectedPsName = fontListBox.selection ? fontListBox.selection.psName : null;
            var showsPostScriptName = filterControls.chkShowPostScriptName.value;
            isUpdatingList = true;
            fontListBox.removeAll();
            for (var i = 0; i < visibleFonts.length; i++) {
                var fontInfo = visibleFonts[i];
                /* 合成フォントの PostScript 名（ATC-…）は読めないので、いつも合成フォント名で出す / ATC-… names are unreadable, so composites always show their own name */
                var itemText;
                if (showsPostScriptName && !fontInfo.isComposite) {
                    itemText = fontInfo.psName;
                } else if (countFamilyRun(visibleFonts, i) === 1) {
                    /* スタイルが1つだけのファミリーは見出しを立てず1行で見せる / Show single-style families on one row */
                    itemText = fontInfo.family + (fontInfo.style ? " " + fontInfo.style : "");
                } else {
                    if (i === 0 || visibleFonts[i - 1].family !== fontInfo.family) {
                        fontListBox.add("item", fontInfo.family).isFamilyHeader = true;
                    }
                    itemText = STYLE_ROW_INDENT + fontInfo.style;
                }
                var fontListItem = fontListBox.add("item", itemText);
                fontListItem.psName = fontInfo.psName;
                fontListItem.familyName = fontInfo.family;
                if (fontInfo.psName === selectedPsName) fontListBox.selection = fontListItem;
            }
            isUpdatingList = false;
            lastSelectedIndex = fontListBox.selection ? fontListBox.selection.index : -1;
            fontListControls.statusText.text = getLabel("message.fontCount", { count: visibleFonts.length });
            updateCustomToggle();
        }

        /**
         * 見出し行が選ばれたら、進んでいた向きの隣のウェイトへ選択を送る（矢印キーで素通りできるように）
         * @returns {boolean} 選択を送ったときは true（このあと onChange がもう一度来る）
         */
        function skipFamilyHeader() {
            var selectedItem = fontListBox.selection;
            if (!selectedItem || !selectedItem.isFamilyHeader) {
                lastSelectedIndex = selectedItem ? selectedItem.index : -1;
                return false;
            }
            /* 真下の行から来たときだけ上向き（マウスで離れた見出しを押したときは、そのファミリーの先頭へ）
               Go up only when coming from the row right below; a clicked header goes to its own first style */
            var isMovingUp = lastSelectedIndex === selectedItem.index + 1 && selectedItem.index > 0;
            var nextIndex = isMovingUp ? selectedItem.index - 1 : selectedItem.index + 1;
            /* 上の行も見出しなら（1行ファミリーが無い並び）下へ / If the row above is a header too, go down instead */
            if (nextIndex >= fontListBox.items.length || fontListBox.items[nextIndex].isFamilyHeader) nextIndex = selectedItem.index + 1;
            fontListBox.selection = (nextIndex < fontListBox.items.length) ? fontListBox.items[nextIndex] : null;
            return true;
        }

        /**
         * 選択中のテキストのフォントを一覧で選び、見える位置までスクロールする（一覧に無ければ何もしない）
         * @returns {void}
         */
        function selectCurrentFont() {
            if (!currentFontName) return;
            for (var i = 0; i < fontListBox.items.length; i++) {
                var fontListItem = fontListBox.items[i];
                if (fontListItem.psName !== currentFontName) continue;
                /* 選び直しただけで適用しない（onChange を止める）/ Selecting it here must not apply it, so onChange is suppressed */
                isUpdatingList = true;
                fontListBox.selection = fontListItem;
                isUpdatingList = false;
                lastSelectedIndex = fontListItem.index;
                updateCustomToggle();
                fontListBox.revealItem(fontListItem);
                return;
            }
        }

        /**
         * ドキュメントの場所と、選択中のテキストの先頭の文字のフォントを読み直す（非同期）。
         * ドキュメントが替わっていたら［ドキュメントフォント］を読み直す
         * @param {boolean} selectsCurrentFont - 読んだフォントを一覧で選ぶなら true
         * @returns {void}
         */
        function readDocumentState(selectsCurrentFont) {
            runWorker("state", "", function (workerResult) {
                var resultParts = workerResult.split("|");
                var isOk = resultParts[0] === "OK";
                var folderPath = (isOk && resultParts[1]) ? decodeURIComponent(resultParts[1]) : "";
                currentFontName = (isOk && resultParts[2]) ? decodeURIComponent(resultParts[2]) : null;
                if (folderPath !== documentFolderPath) {
                    documentFolderPath = folderPath;
                    documentFontNames = readProjectFontNames(folderPath);
                    if (filterControls.chkDocumentFonts.value) refreshFontList();
                }
                if (selectsCurrentFont) selectCurrentFont();
            });
        }

        /**
         * 一覧で選んでいるフォントを、今選択しているテキストに適用する（パレットは開いたまま）。
         * テキストを選んでいないときは黙って何もしない（クリックのたびに警告を出さない）
         * @returns {void}
         */
        function applySelectedFont() {
            if (!fontListBox.selection) return;
            var psName = fontListBox.selection.psName;
            runWorker("apply", psName, function (workerResult) {
                if (workerResult === "NOFONT") alert(getLabel("alert.fontNotFound", { name: psName }));
                else if (workerResult === "OK") currentFontName = psName;
            });
        }

        /**
         * 一覧で選んだフォントのファミリーを［カスタムセット］に加え、そのセットのチェックを入れる
         * @param {number} setIndex - セットの位置（0 から）
         * @returns {void}
         */
        function addSelectedToSet(setIndex) {
            if (!fontListBox.selection) return;
            filterControls.customFontSets[setIndex] = filterControls.customFontSets[setIndex].concat([fontListBox.selection.familyName]);
            filterControls.customSetCheckboxes[setIndex].value = true;
            refreshFontList();
        }

        /**
         * 一覧で選んだフォントに当たる名前を［カスタムセット］から外す（ほかのフォントにも当たる名前なら確かめる）
         * @param {number} setIndex - セットの位置（0 から）
         * @returns {void}
         */
        function removeSelectedFromSet(setIndex) {
            if (!fontListBox.selection) return;
            var selectedItem = fontListBox.selection;
            var familyLower = selectedItem.familyName.toLowerCase();
            var psLower = selectedItem.psName.toLowerCase();
            var setFontNames = filterControls.customFontSets[setIndex];
            var keptNames = [];
            var removedNames = [];
            for (var i = 0; i < setFontNames.length; i++) {
                var isMatch = startsWithAny(familyLower, [setFontNames[i]]) || startsWithAny(psLower, [setFontNames[i]]);
                if (isMatch) removedNames.push(setFontNames[i]);
                else keptNames.push(setFontNames[i]);
            }
            if (removedNames.length === 0) return;
            var isOnlyThisFamily = (removedNames.length === 1 && removedNames[0].toLowerCase() === familyLower);
            if (!isOnlyThisFamily && !confirm(getLabel("alert.confirmRemoveCustom", { number: setIndex + 1, names: removedNames.join(uiLang === "ja" ? "」「" : "\", \"") }))) return;
            filterControls.customFontSets[setIndex] = keptNames;
            refreshFontList();
        }

        /**
         * 一覧で選んだフォントが［カスタムセット］に入っているか（名前の前方一致）
         * @param {number} setIndex - セットの位置（0 から）
         * @returns {boolean} 入っていれば true
         */
        function isSelectedInSet(setIndex) {
            var selectedItem = fontListBox.selection;
            if (!selectedItem || selectedItem.isFamilyHeader) return false;
            var setFontNames = filterControls.customFontSets[setIndex];
            return startsWithAny(selectedItem.familyName.toLowerCase(), setFontNames) ||
                startsWithAny(selectedItem.psName.toLowerCase(), setFontNames);
        }

        /**
         * ［カスタムセット］の出し入れボタンを、一覧の選択に合わせる（しおりの塗り・有効・ツールチップ）
         * @returns {void}
         */
        function updateCustomToggle() {
            var selectedItem = fontListBox.selection;
            var hasFont = !!selectedItem && !selectedItem.isFamilyHeader;
            for (var i = 0; i < fontListControls.customSetButtons.length; i++) {
                var setButton = fontListControls.customSetButtons[i];
                setButton.enabled = hasFont;
                setButton.isOn = isSelectedInSet(i);
                setButton.helpTip = getLabel(setButton.isOn ? "tooltip.removeCustomSet" : "tooltip.addCustomSet", { number: i + 1 });
                setButton.notify("onDraw");
            }
        }

        /**
         * 一覧で選んだフォントを、［カスタムセット］に入っていれば外し、入っていなければ加える
         * @param {number} setIndex - セットの位置（0 から）
         * @returns {void}
         */
        function toggleSelectedInSet(setIndex) {
            if (isSelectedInSet(setIndex)) removeSelectedFromSet(setIndex);
            else addSelectedToSet(setIndex);
            updateCustomToggle();
        }

        addSoloClickGroup(filterControls.customSetCheckboxes.concat([filterControls.chkDocumentFonts, filterControls.chkCompositeFonts], filterControls.foundryCheckboxes), refreshFontList);
        addSoloClickGroup(filterControls.suffixCheckboxes, refreshFontList);
        filterControls.chkJapaneseOnly.onClick = refreshFontList;
        filterControls.chkShowAll.onClick = filterControls.chkFilterByStandard.onClick = function () {
            updatePanelStates();
            refreshFontList();
        };
        filterControls.chkShowPostScriptName.onClick = refreshFontList;
        for (var b = 0; b < fontListControls.customSetButtons.length; b++) {
            fontListControls.customSetButtons[b].onClick = function () {
                toggleSelectedInSet(this.setNumber - 1);
            };
        }
        fontListBox.onChange = function () {
            if (isUpdatingList || skipFamilyHeader()) return;
            updateCustomToggle();
            applySelectedFont();
        };
        /**
         * クリアボタンの有効・無効を絞り込み欄に合わせる
         * @returns {void}
         */
        function updateClearButtonState() {
            var hasSearchText = filterControls.searchInput.text !== "";
            if (fontListControls.btnClearSearch.enabled === hasSearchText) return;
            fontListControls.btnClearSearch.enabled = hasSearchText;
            fontListControls.btnClearSearch.notify("onDraw");
        }

        /* 一覧に並べるのが重いので、件数が多いうちは Enter（onChange）まで待つ / Adding many rows is slow, so wait for Enter while the result is large */
        filterControls.searchInput.onChanging = function () {
            updateClearButtonState();
            var visibleFonts = collectVisibleFonts();
            if (visibleFonts.length <= LIVE_SEARCH_MAX_FONTS) {
                showFontList(visibleFonts);
                return;
            }
            isSearchPending = true;
            fontListControls.statusText.text = getLabel("message.searchPending", { count: visibleFonts.length });
        };
        filterControls.searchInput.onChange = function () {
            if (isSearchPending) refreshFontList();
        };
        fontListControls.btnClearSearch.onClick = function () {
            filterControls.searchInput.text = "";
            /* 続けて打ち直せるよう、フォーカスは入力欄へ戻す / Put the focus back in the field */
            filterControls.searchInput.active = true;
            updateClearButtonState();
            refreshFontList();
        };
        /* 前回の絞り込みを読み戻したときの状態に合わせる / Match the restored filter text */
        fontListControls.btnClearSearch.enabled = filterControls.searchInput.text !== "";

        btnPreferences.onClick = function () {
            var preferences = showPreferencesDialog({
                fontTotal: fontCatalog.length,
                excludedLanguages: filterControls.excludedLanguages,
                customFontSets: filterControls.customFontSets,
                onRescan: function () {
                    adobeFontsFamilies = null;
                    try {
                        fontCatalog = createFontCatalog(scanInstalledFonts());
                    } catch (e) {
                        /* パレットから app.textFonts を読めないときは、キャッシュを消して次の起動で読み直す
                           When the palette cannot read app.textFonts, drop the cache so the next launch rescans */
                        if (FONT_CACHE_FILE.exists) FONT_CACHE_FILE.remove();
                        alert(getLabel("alert.rescanFailed"));
                        return -1;
                    }
                    return fontCatalog.length;
                }
            });
            /* 再スキャンは環境設定をキャンセルしても戻らないので、いつも一覧を作り直す / A rescan stays even on Cancel, so always rebuild the list */
            if (preferences) {
                filterControls.excludedLanguages = preferences.excludedLanguages;
                filterControls.customFontSets = preferences.customFontSets;
            }
            refreshFontList();
        };

        /* 選んだままの行はクリックしても onChange が来ないので、別のテキストに同じフォントを当てるときはダブルクリック・Enter で
           Clicking the row already chosen fires no onChange, so double-click or Enter applies the same font to other text */
        fontListBox.onDoubleClick = applySelectedFont;
        fontListBox.addEventListener("keydown", function (keyEvent) {
            if (keyEvent.keyName !== "Enter") return;
            keyEvent.preventDefault();
            applySelectedFont();
        });
        btnClose.onClick = function () {
            mainPalette.close();
        };
        /* パレットは Esc で閉じないので割り当てる。keydown の中では BridgeTalk を待たず、片付けは onClose に任せる
           Palettes do not close on Esc; nothing waits on BridgeTalk here, onClose does the cleanup */
        addKeyShortcuts(mainPalette, {
            "Escape": { target: function () { mainPalette.close(); }, inFields: true }
        });

        mainPalette.onResizing = mainPalette.onResize = function () {
            this.layout.resize();
        };
        mainPalette.onShow = function () {
            filterControls.searchInput.active = true;
            readDocumentState(true);
        };
        /* 選択やドキュメントはパレットの外で替わるので、戻ってくるたびに読み直す / Selection and document change outside, so re-read on return */
        mainPalette.onActivate = function () {
            readDocumentState(false);
        };
        /* onClose では DOM に触れない（落ちる）/ Never touch the DOM in onClose (it crashes) */
        mainPalette.onClose = function () {
            /* 閉じ方にかかわらず、絞り込みの状態を覚える / Remember the filter state however the palette was closed */
            settingsStore.save(readSettingsFromControls(filterControls));
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        updatePanelStates();
        refreshFontList();

        prepareDialogWindow(mainPalette, SCRIPT_NAME);
        /* 参照を残しておかないと、スクリプトの終了とともにパレットが消える / Keep a reference, or the palette vanishes when the script ends */
        $.global[PALETTE_GLOBAL_KEY] = mainPalette;
        mainPalette.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 開いているパレットがあれば前面に出すだけ / If the palette is already open, just bring it to the front */
    var existingPalette = $.global[PALETTE_GLOBAL_KEY];
    if (existingPalette) {
        try {
            existingPalette.show();
            existingPalette.active = true;
            return;
        } catch (e) {
            /* 参照が切れていたら作り直す / A stale reference: build a new palette */
            $.global[PALETTE_GLOBAL_KEY] = null;
        }
    }

    showFavoriteFontPalette(readFontCache() || scanInstalledFonts());

})();
