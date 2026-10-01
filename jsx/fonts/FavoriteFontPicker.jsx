#target illustrator
#targetengine "FavoriteFontPickerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

よく使うフォントを分類と規格で絞り込んだ一覧から選び、プレビューで確かめて選択中のテキストに適用します。
同じ書体の規格違い（Pr6N・Pr5 など）はチェックした規格の中で優先順位の高いものだけを残し、中国語・韓国語など日本語以外のフォントは外します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPicker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ncf9ff6feebf0

### Overview

Lists the fonts you use often, narrowed by category and standard, and applies the one you pick to the selected text after a preview.
Of the same typeface in several standards (Pr6N, Pr5, etc.), only the highest-priority checked one is kept, and Chinese, Korean and other non-Japanese fonts are left out.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteFontPicker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FavoriteFontPicker";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "KOUJI & 相棒（Gem）";              /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2026-10-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-02";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPicker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteFontPicker.md"; /* README (English) */
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

    /* 同じ書体が複数あるとき残す規格（接尾辞）の優先順位（先頭ほど優先）。［規格］のチェックボックスにも並ぶ
       Standard (suffix) priority for the same typeface (first wins); also listed as Standard checkboxes */
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
        { key: "tb", label: { ja: "TB（タイプバンク）", en: "TB (TypeBank)" }, tooltip: { ja: "TB で始まるタイプバンクのフォント", en: "TypeBank fonts starting with TB" }, familyPrefixes: ["TB"], psNamePrefixes: ["TB"] },
        { key: "fot", label: { ja: "FOT（フォントワークス）", en: "FOT (Fontworks)" }, tooltip: { ja: "FOT- で始まるフォントワークスのフォント", en: "Fontworks fonts starting with FOT-" }, familyPrefixes: ["FOT-"], psNamePrefixes: ["FOT-"] },
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

    var FONT_LIST_SIZE = [320, 380];   /* フォント一覧の大きさ / font list size */
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
    var STANDARD_LEFT_WIDTH = 64;      /* ［規格］の左列（無印）の幅 / width of the left (plain) column under Standard */
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

    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");

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
        filterByStandard: true, /* ［規格］のチェックで絞り込む / filter by the Standard checkboxes */
        japaneseOnly: true,     /* ［和文フォントのみ］/ Japanese fonts only */
        excludedLanguages: ["chinese", "korean", "thai", "multilingual"], /* ［和文フォントのみ］で外す言語（環境設定）/ languages left out by Japanese fonts only (preferences) */
        hiddenSuffixes: [],     /* ［規格］で外した接尾辞（"" は［その他］）/ suffixes unchecked under Standard ("" is Other) */
        searchText: "",
        showPostScriptName: false,  /* 一覧を PostScript 名で表示 / list PostScript names */
        preview: true               /* 選んだフォントを仮に適用 / preview the chosen font */
    };

    /* フォント情報のキャッシュ（1行に PostScript 名・ファミリー名・スタイル名をタブ区切り）
       Font cache: one font per line, PostScript name, family and style separated by tabs */
    var FONT_CACHE_FILE = new File(Folder.userData + "/illustrator-scripts/" + SCRIPT_NAME + "-fonts.txt");

    /* Creative Cloud が Adobe Fonts の同期状態を書くファイル（Mac は .c、Windows は c のフォルダー）
       Creative Cloud's Adobe Fonts sync state (in .c on Mac, c on Windows) */
    var ADOBE_FONTS_STATE_PATHS = [
        "/Adobe/CoreSync/plugins/livetype/.c/entitlements.xml",
        "/Adobe/CoreSync/plugins/livetype/c/entitlements.xml"
    ];

    // =========================================
    // 選択の収集 / Selection
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
            standard: { ja: "規格", en: "Standard" },
            excludedLanguages: { ja: "［和文フォントのみ］で外す言語", en: "Languages left out by Japanese fonts only" },
            fontInfo: { ja: "フォント情報", en: "Font information" },
            customSets: { ja: "カスタムセット", en: "Custom sets" }
        },
        checkbox: {
            documentFonts: { ja: "ドキュメントフォント", en: "Document fonts" },
            compositeFonts: { ja: "合成フォント", en: "Composite fonts" },
            showAll: { ja: "すべて表示", en: "Show all" },
            japaneseOnly: { ja: "和文フォントのみ", en: "Japanese fonts only" },
            filterByStandard: { ja: "規格で絞り込む", en: "Filter by standard" },
            noSuffix: { ja: "規格なし", en: "No standard" },
            showPostScriptName: { ja: "PostScript 名で表示", en: "Show PostScript names" },
            preview: { ja: "選択中のテキストにプレビュー", en: "Preview on selected text" }
        },
        fieldLabel: {
            search: { ja: "絞り込み", en: "Filter" },
            customSets: { ja: "カスタムセット", en: "Custom sets" }
        },
        button: {
            rescan: { ja: "再スキャン", en: "Rescan" },
            preferences: { ja: "環境設定…", en: "Preferences…" },
            exportSet: { ja: "書き出し…", en: "Export…" },
            importSet: { ja: "読み込み…", en: "Import…" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" }
        },
        tooltip: {
            customSet: {
                ja: "カスタムセット {number}：一覧の下の {number} のボタンで加えたフォント。規格違いの間引きと除外はしない",
                en: "Custom set {number}: fonts added with button {number} below the list, shown without thinning out standards or exclusions"
            },
            documentFonts: {
                ja: "ドキュメントと同じフォルダーの _ProjectFonts.txt に挙げたフォント。規格違いの間引きと除外はしない",
                en: "Fonts listed in _ProjectFonts.txt next to the document, shown without thinning out standards or exclusions"
            },
            compositeFonts: {
                ja: "［書式］→［合成フォント］で作った合成フォント。規格の絞り込みと間引きはしない",
                en: "Composite fonts made with Type > Composite Fonts, not filtered or thinned out by standard"
            },
            showAll: {
                ja: "［分類］・規格違いの間引き・除外をやめてすべて表示（［規格］と［日本語以外を除外］は効く）",
                en: "Ignore Category, standard thinning and exclusions and list every font (Standard and Exclude non-Japanese still apply)"
            },
            standard: {
                ja: "チェックした規格だけを表示。同じ書体はチェックした中で優先順位がいちばん高い規格を残す",
                en: "Show only the checked standards; for each typeface, keep the highest-priority checked standard"
            },
            filterByStandard: {
                ja: "オフにすると［規格］のチェックを無視する（同じ書体は、すべての規格の中で優先順位がいちばん高いものを残す）",
                en: "When off, the Standard checkboxes are ignored (each typeface keeps its highest-priority standard of all)"
            },
            noSuffix: { ja: "Pr6N などの規格が名前の末尾に付いていないフォント", en: "Fonts without a standard such as Pr6N at the end of the name" },
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
            apply: { ja: "一覧で選んだフォントを選択中のテキストに適用して閉じる（ダブルクリック・Enter キーと同じ）", en: "Apply the chosen font to the selected text and close (same as double-click or Enter)" },
            preview: {
                ja: "一覧で選んだフォントを選択中のテキストに仮に適用する（文字単位の選択と、連結したテキストは対象外）",
                en: "Preview the chosen font on the selected text (not for selected characters or threaded text)"
            },
            fontList: { ja: "ダブルクリックまたは Enter キーで、選択中のテキストに適用", en: "Double-click or press Enter to apply to the selected text" },
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
     * 中心の名前ごとに、［規格］でチェックした中で優先順位がいちばん高い接頭辞・接尾辞の組を選ぶ
     * @param {Object[]} fontInfos - createFontCatalog() 済みのフォント情報
     * @param {Object} visibleSuffixes - 表示する接尾辞 → true（"" は規格なし）
     * @returns {Object} 中心の名前 → 残す組（parseFamilyName() の結果）
     */
    function buildBestVariantMap(fontInfos, visibleSuffixes) {
        var bestByCore = {};
        for (var i = 0; i < fontInfos.length; i++) {
            /* 合成フォントは名前が自由なので、規格違いの比較に入れない / Composite names are free-form, so they stay out of the comparison */
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
     * ファミリー名を接頭辞・中心の名前・接尾辞（規格）に分ける
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
            /* ［規格で絞り込む］がオフなら、どの規格も出す / With Filter by standard off, every standard is shown */
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
        /* カスタム・ドキュメントに挙げたフォントは、規格違いの間引きと除外をせずに出す / Listed Custom and Document fonts skip thinning and exclusions */
        var isListed = (filterState.custom && fontNameStartsWithAny(fontInfo, filterState.customFontNames)) ||
            (filterState.documentFonts && fontNameStartsWithAnyWord(fontInfo, filterState.documentFontNames));
        /* ［日本語以外を除外］の言語は外す（カスタム・ドキュメントに挙げたものは残す）/ Leave out excluded languages, except listed Custom and Document fonts */
        if (!isListed && filterState.excludedLanguages[fontInfo.languageKey] === true) return false;
        /* 合成フォントは、［すべて表示］ではほかと同じく［規格］で絞り、それ以外は規格を問わない
           Composite fonts follow Standard under Show all like any other font, and ignore it otherwise */
        if (fontInfo.isComposite) {
            if (filterState.showAll) return filterState.visibleSuffixes[fontInfo.parsed.suffix] === true;
            return filterState.compositeFonts || isListed;
        }
        /* ［規格］の絞り込みは［すべて表示］でも効く / The Standard filter applies even with Show all */
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
     * 作業中のドキュメントにある _ProjectFonts.txt を返す（未保存なら null）
     * @returns {File|null} ファイル（まだ無くてもよい）
     */
    function getProjectFontsFile() {
        if (app.documents.length === 0) return null;
        var documentFolder;
        try {
            documentFolder = app.activeDocument.path;
        } catch (e) {
            /* 未保存のドキュメントは path が読めないことがある / An unsaved document may have no readable path */
            return null;
        }
        if (!documentFolder || String(documentFolder.fsName || "") === "") return null;
        return new File(documentFolder.fsName + "/" + PROJECT_FONTS_FILE_NAME);
    }

    /**
     * _ProjectFonts.txt の名前を読む
     * @returns {string[]} フォント名の一覧（ファイルが無ければ空）
     */
    function readProjectFontNames() {
        var projectFontsFile = getProjectFontsFile();
        if (!projectFontsFile) return [];
        var fileLines = splitTextLines(readTextFile(projectFontsFile));
        var fontNames = [];
        for (var i = 0; i < fileLines.length; i++) {
            var fontName = fileLines[i].replace(/^\s+|\s+$/g, "");
            if (fontName !== "") fontNames.push(fontName);
        }
        return fontNames;
    }

    // =========================================
    // 適用 / Apply
    // =========================================

    /**
     * ダイアログを開く前に、適用先（文字の選択、またはテキストフレーム）を控える。
     * プレビューで元のテキストを隠すと選択が外れるので、閉じたあとはこの控えに適用する
     * @returns {{textRange: TextRange|null, textFrames: TextFrame[], selectedItems: Array}} 適用先と、開いたときの選択
     */
    function captureApplyTargets() {
        var applyTargets = { textRange: null, textFrames: [], selectedItems: [] };
        if (app.documents.length === 0) return applyTargets;
        var docSelection = app.activeDocument.selection;
        /* 文字ツールで選んだ文字（TextRange）はその範囲だけに / Characters selected with the Type tool get the font on their own */
        if (docSelection && docSelection.typename === "TextRange") {
            applyTargets.textRange = docSelection;
            return applyTargets;
        }
        applyTargets.selectedItems = normalizeSelectionItems(docSelection);
        applyTargets.textFrames = collectSelectionTextFrames(docSelection, { skipLocked: true, skipHidden: true });
        return applyTargets;
    }

    /**
     * 適用先があるか
     * @param {Object} applyTargets - captureApplyTargets() の戻り値
     * @returns {boolean} あれば true
     */
    function hasApplyTargets(applyTargets) {
        return applyTargets.textRange !== null || applyTargets.textFrames.length > 0;
    }

    /**
     * PostScript 名からフォントを引く
     * @param {string} psName - PostScript 名
     * @returns {TextFont|null} フォント。無ければ null
     */
    function findTextFont(psName) {
        try {
            return app.textFonts.getByName(psName);
        } catch (e) {
            /* キャッシュのあとでアンインストールされたフォント / A font uninstalled after the cache was written */
            return null;
        }
    }

    /**
     * 控えた適用先にフォントを適用する
     * @param {Object} applyTargets - captureApplyTargets() の戻り値
     * @param {string} psName - フォントの PostScript 名
     * @returns {boolean} 適用できたら true（できなければ理由を alert で出す）
     */
    function applyFontToTargets(applyTargets, psName) {
        var textFont = findTextFont(psName);
        if (!textFont) {
            alert(getLabel("alert.fontNotFound", { name: psName }));
            return false;
        }
        if (applyTargets.textRange) {
            applyTargets.textRange.characterAttributes.textFont = textFont;
        } else {
            for (var i = 0; i < applyTargets.textFrames.length; i++) {
                applyTargets.textFrames[i].textRange.characterAttributes.textFont = textFont;
            }
        }
        app.redraw();
        return true;
    }

    /**
     * 適用先の先頭の文字のフォント（PostScript 名）を返す
     * @param {Object} applyTargets - captureApplyTargets() の戻り値
     * @returns {string|null} PostScript 名。適用先が無い・読めないときは null
     */
    function getTargetFontName(applyTargets) {
        var targetRange = applyTargets.textRange || (applyTargets.textFrames.length > 0 ? applyTargets.textFrames[0].textRange : null);
        if (!targetRange) return null;
        /* 文字があれば先頭の文字、無ければ（文字カーソルだけなら）その位置の書式 / The first character, or the insertion point's own format */
        if (targetRange.characters.length > 0) targetRange = targetRange.characters[0];
        try {
            return targetRange.characterAttributes.textFont.name;
        } catch (e) {
            /* 環境にないフォントは textFont が読めない / textFont cannot be read for a missing font */
            return null;
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューの状態を作る。元のテキストを複製して隠し、複製にフォントを当てる
     * @param {Object} applyTargets - captureApplyTargets() の戻り値
     * @returns {{applyTargets: Object, previewFrames: TextFrame[], hiddenFrames: TextFrame[]}} プレビューの状態
     */
    function createFontPreview(applyTargets) {
        return { applyTargets: applyTargets, previewFrames: [], hiddenFrames: [] };
    }

    /**
     * 選んだフォントを複製に当てて見せる。まだ複製が無ければ作る
     * 文字単位の選択と、連結したテキスト（複製すると連結が切れる）はプレビューしない
     * @param {Object} fontPreview - createFontPreview() の戻り値
     * @param {string} psName - フォントの PostScript 名
     * @returns {void}
     */
    function updateFontPreview(fontPreview, psName) {
        var textFont = findTextFont(psName);
        if (!textFont) return;
        if (fontPreview.previewFrames.length === 0) {
            var textFrames = fontPreview.applyTargets.textFrames;
            for (var i = 0; i < textFrames.length; i++) {
                if (textFrames[i].story.textFrames.length > 1) continue;
                /* 複製は元を隠す前に作る（hidden を引き継ぐため）/ Duplicate before hiding, since duplicates inherit hidden */
                fontPreview.previewFrames.push(textFrames[i].duplicate(textFrames[i], ElementPlacement.PLACEBEFORE));
                textFrames[i].hidden = true;
                fontPreview.hiddenFrames.push(textFrames[i]);
            }
        }
        for (var j = 0; j < fontPreview.previewFrames.length; j++) {
            fontPreview.previewFrames[j].textRange.characterAttributes.textFont = textFont;
        }
        app.redraw();
    }

    /**
     * プレビューの複製を消し、隠した元のテキストを戻して選択も戻す
     * @param {Object} fontPreview - createFontPreview() の戻り値
     * @returns {void}
     */
    function clearFontPreview(fontPreview) {
        if (fontPreview.previewFrames.length === 0) return;
        for (var i = 0; i < fontPreview.previewFrames.length; i++) fontPreview.previewFrames[i].remove();
        for (var j = 0; j < fontPreview.hiddenFrames.length; j++) fontPreview.hiddenFrames[j].hidden = false;
        fontPreview.previewFrames = [];
        fontPreview.hiddenFrames = [];
        /* 隠すと選択から外れるので、開いたときの選択に戻す / Hiding deselects, so restore the original selection */
        app.activeDocument.selection = fontPreview.applyTargets.selectedItems;
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 左の列（分類・すべて表示・規格のチェックボックス）を作る
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
        /* ［規格で絞り込む］がオフのときに無効にする部分 / The part disabled while Filter by standard is off */
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
                    /* 対になる規格が無いときは空けて列をそろえる / Leave a gap so the columns stay aligned */
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
     * ［規格］の並びを作る。左に無印、右に N 付き（Pro と ProN など）を置き、上から優先順位の低い順に並べる
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
        setupRow(customRowGroup, ["left", "top"], 6);
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
     * @param {{fontTotal: number, excludedLanguages: string[], customFontSets: string[][], onRescan: function(): number}} preferenceState - 今の値と、再スキャンの処理（読み込んだ件数を返す）
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
     * @param {Object} filterControls - buildFilterColumn() の戻り値に searchInput・chkShowPostScriptName・chkPreview・customFontSets・excludedLanguages を足したもの
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
        filterControls.chkPreview.value = filterSettings.preview;
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
            showPostScriptName: filterControls.chkShowPostScriptName.value,
            preview: filterControls.chkPreview.value
        };
    }

    /**
     * ダイアログを作って表示する
     * @param {Object[]} fontInfos - フォント情報
     * @returns {void}
     */
    function showFavoriteFontDialog(fontInfos) {
        var fontCatalog = createFontCatalog(fontInfos);
        var documentFontNames = readProjectFontNames();
        var applyTargets = captureApplyTargets();
        var currentFontName = getTargetFontName(applyTargets);
        var fontPreview = createFontPreview(applyTargets);
        var chosenPsName = null;

        var mainDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: true });
        setupWindow(mainDialog);

        var columnsGroup = mainDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.alignment = ["fill", "fill"];
        columnsGroup.spacing = COLUMN_SPACING;

        var filterControls = buildFilterColumn(columnsGroup);
        var fontListControls = buildFontListColumn(columnsGroup);
        var fontListBox = fontListControls.fontListBox;
        filterControls.searchInput = fontListControls.searchInput;
        filterControls.chkShowPostScriptName = fontListControls.chkShowPostScriptName;

        var buttonRow = addButtonRow(mainDialog);
        var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
        btnPreferences.helpTip = getLabel("tooltip.preferences");
        /* 適用先（テキスト）に関わる設定なので、一覧ではなくボタン行に置く / It concerns the target text, so it sits in the button row */
        filterControls.chkPreview = addCheckbox(buttonRow.leftGroup, getLabel("checkbox.preview"), "tooltip.preview");
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLabel("button.apply"));
        btnApply.helpTip = getLabel("tooltip.apply");
        alignRightOnlyButtonRow(buttonRow);

        applySettingsToControls(filterControls, settingsStore.load(DEFAULT_SETTINGS));

        /**
         * ［すべて表示］のときは［分類］を、［規格で絞り込む］がオフのときは［規格］を無効にする
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
                fontListBox.selection = fontListItem;
                fontListBox.revealItem(fontListItem);
                return;
            }
        }

        /**
         * 一覧で選んでいるフォントに決めて閉じる（適用は閉じたあと）
         * @returns {void}
         */
        function applySelectedFont() {
            if (!fontListBox.selection) return;
            if (app.documents.length === 0) {
                alert(getLabel("alert.noDocument"));
                return;
            }
            if (!hasApplyTargets(applyTargets)) {
                alert(getLabel("alert.noTextSelected"));
                return;
            }
            chosenPsName = fontListBox.selection.psName;
            mainDialog.close(1);
        }

        /**
         * 一覧で選んだフォントをプレビューする（プレビューがオフなら消す）
         * @returns {void}
         */
        function refreshPreview() {
            if (!filterControls.chkPreview.value) {
                clearFontPreview(fontPreview);
                return;
            }
            /* 一覧を作り直す間は選択が外れるので、直前のプレビューを残す / The list loses its selection while rebuilding, so keep the last preview */
            if (!fontListBox.selection) return;
            var psName = fontListBox.selection.psName;
            /* 開いたときのフォントのままなら、複製を作らない / No duplicates while the font is still the original one */
            if (fontPreview.previewFrames.length === 0 && psName === currentFontName) return;
            updateFontPreview(fontPreview, psName);
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
        filterControls.chkPreview.onClick = refreshPreview;
        for (var b = 0; b < fontListControls.customSetButtons.length; b++) {
            fontListControls.customSetButtons[b].onClick = function () {
                toggleSelectedInSet(this.setNumber - 1);
            };
        }
        fontListBox.onChange = function () {
            if (isUpdatingList || skipFamilyHeader()) return;
            updateCustomToggle();
            refreshPreview();
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
                    fontCatalog = createFontCatalog(scanInstalledFonts());
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

        fontListBox.onDoubleClick = applySelectedFont;
        /* ［適用］は既定のボタンにしない（絞り込み欄の Enter で適用されないように）。一覧の Enter で適用する
           Apply is not the default button, so Enter in the filter field does not apply; Enter in the list does */
        fontListBox.addEventListener("keydown", function (keyEvent) {
            if (keyEvent.keyName !== "Enter") return;
            keyEvent.preventDefault();
            applySelectedFont();
        });
        btnApply.onClick = applySelectedFont;
        btnCancel.onClick = function () {
            mainDialog.close(2);
        };

        mainDialog.onResizing = mainDialog.onResize = function () {
            this.layout.resize();
        };
        mainDialog.onShow = function () {
            filterControls.searchInput.active = true;
            selectCurrentFont();
        };

        updatePanelStates();
        refreshFontList();

        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        var dialogResult = mainDialog.show();

        /* DOM の後始末と適用は閉じたあとに行う / Clean up and apply after the dialog has closed */
        clearFontPreview(fontPreview);
        if (dialogResult === 1 && chosenPsName) applyFontToTargets(applyTargets, chosenPsName);

        /* 閉じ方にかかわらず、絞り込みの状態を覚える / Remember the filter state however the dialog was closed */
        settingsStore.save(readSettingsFromControls(filterControls));
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    showFavoriteFontDialog(readFontCache() || scanInstalledFonts());

})();
