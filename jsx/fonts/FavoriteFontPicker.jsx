#target illustrator
#targetengine "FavoriteFontPickerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

よく使うフォントを分類と規格で絞り込んだ一覧から選び、プレビューで確かめて選択中のテキストに適用します。
同じ書体の規格違い（Pr6N・Pr5 など）はチェックした規格の中で優先順位の高いものだけを残し、中華・ハングルなどのフォントは外します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FavoriteFontPicker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ncf9ff6feebf0

### Overview

Lists the fonts you use often, narrowed by category and standard, and applies the one you pick to the selected text after a preview.
Of the same typeface in several standards (Pr6N, Pr5, etc.), only the highest-priority checked one is kept, and Chinese, Hangul and other fonts are left out.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FavoriteFontPicker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FavoriteFontPicker";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
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

    /* ［カスタム］の初期値（ファミリー名・PostScript 名の前方一致）。以後は［カスタムに追加］［カスタムから外す］で変え、設定ファイルに保存する
       Initial Custom fonts (prefix match on the family or PostScript name); later changed in the dialog and saved */
    var CUSTOM_FONTS = ["Graphik", "DIN"];

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
            familyPrefixes: ["A-OTF", "A P-OTF", "AP-OTF"],
            psNamePrefixes: ["Ryumin", "ShinGo", "UDShinGo", "GothicMB101", "FutoGoB101", "MidashiGo"]
        },
        { key: "tb", label: { ja: "TB（タイプバンク）", en: "TB (TypeBank)" }, familyPrefixes: ["TB"], psNamePrefixes: ["TB"] },
        { key: "fot", label: { ja: "FOT（フォントワークス）", en: "FOT (Fontworks)" }, familyPrefixes: ["FOT-"], psNamePrefixes: ["FOT-"] },
        {
            key: "hiragino",
            label: { ja: "ヒラギノ", en: "Hiragino" },
            familyPrefixes: ["ヒラギノ", "Hiragino"],
            psNamePrefixes: ["Hiragino", "HiraKaku", "HiraMin", "HiraMaru"]
        },
        {
            /* Adobe Fonts（配信サービス）は名前で見分けられないので、Adobe 自社の和文書体をまとめる
               Adobe Fonts (the service) cannot be told by name, so this groups Adobe's own Japanese typefaces */
            key: "adobe",
            label: { ja: "Adobe（小塚・源ノ・りょう）", en: "Adobe (Kozuka, Source Han, Ryo)" },
            familyPrefixes: ["小塚", "源ノ", "りょう", "かづらき", "Kozuka", "Source Han", "Kazuraki"],
            psNamePrefixes: ["Koz", "SourceHan", "RyoGothic", "RyoDisp", "RyoText", "Kazuraki"]
        }
    ];

    /* ［日本語以外を除外］で外す言語。名前（ファミリー名・PostScript 名）で判定し、上から順に当てる
       namePrefixes：前方一致（大文字小文字は区別しない）、familyPatterns：ファミリー名、psNamePatterns：PostScript 名
       Languages left out with Exclude non-Japanese, judged by name and tried from the top */
    var LANGUAGE_GROUPS = [
        {
            key: "chinese",
            label: { ja: "中華", en: "Chinese" },
            tooltip: { ja: "PingFang・宋体・黑体・Source Han Sans SC など、中国語（簡体字・繁体字）のフォント", en: "Chinese (Simplified and Traditional) fonts such as PingFang, Songti, Heiti and Source Han Sans SC" },
            namePrefixes: ["PingFang", "Hiragino Sans GB", "Hiragino Sans CNS", "HiraginoSansCNS", "STHeiti", "STSong", "STKaiti", "STFangsong", "STXihei", "Songti", "Heiti", "Kaiti", "Lantinghei", "Hannotate", "Hanzipen", "Weibei", "Libian", "Xingkai", "Baoli", "Wawati", "Yuanti", "Yuppy", "BiauKai", "LiSong", "LiHei", "Apple LiGothic", "Apple LiSung", "GB18030", "SimSun", "NSimSun", "SimHei", "Microsoft YaHei", "Microsoft JhengHei", "MingLiU", "PMingLiU", "DFKai", "AdobeSong", "AdobeMing", "AdobeHeiti", "AdobeFangsong", "AdobeKaiti", "思源", "苹方", "蘋方", "华文", "華文", "黑体", "黑體", "宋体", "宋體", "楷体", "仿宋", "圆体", "圓體", "兰亭", "蘭亭", "翩翩", "魏碑", "隶变", "隸變", "行楷", "报隶", "報隸", "娃娃", "雅痞"],
            familyPatterns: [/ (SC|TC|HK|CN)$/],
            /* 地域の印は直前が小文字のときだけ（ZapfDingbatsITC の TC、RoNOWStd-GB のウェイト GB を拾わない）
               Region marks count only after a lowercase letter, so ITC or a GB weight name is not picked up */
            psNamePatterns: [/^[^-]*[a-z0-9](SC|TC|HK|CN|GB)(-|$)/, /CJK(sc|tc|hk)/i]
        },
        {
            key: "korean",
            label: { ja: "ハングル", en: "Hangul" },
            tooltip: { ja: "Apple SD Gothic Neo・AppleMyungjo・Nanum など、ハングルのフォント", en: "Korean fonts such as Apple SD Gothic Neo, AppleMyungjo and Nanum" },
            namePrefixes: ["AppleGothic", "AppleMyungjo", "AppleSDGothicNeo", "Apple SD", "Nanum", "Malgun", "Gulim", "Batang", "Dotum", "Gungsuh", "PCMyungjo", "HeadLineA", "AdobeMyungjo", "AdobeGothicStd", "JCsmPC", "JCfg", "JCkg", "JCHEadA"],
            familyPatterns: [/[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF]/, / KR$/],
            psNamePatterns: [/^[^-]*[a-z0-9]KR(-|$)/, /CJKkr/i]
        },
        {
            key: "thai",
            label: { ja: "タイ", en: "Thai" },
            tooltip: { ja: "Thonburi・Sukhumvit Set など、タイ語のフォント", en: "Thai fonts such as Thonburi and Sukhumvit Set" },
            namePrefixes: ["Thonburi", "Ayuthaya", "Krungthep", "Sathu", "Silom", "Sukhumvit"],
            familyPatterns: [/[\u0E00-\u0E7F]/, /\bThai\b/i],
            psNamePatterns: []
        },
        {
            key: "multilingual",
            label: { ja: "その他マルチリンガル", en: "Other multilingual" },
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

    var FONT_LIST_SIZE = [320, 440];   /* フォント一覧の大きさ / font list size */
    var STATUS_TEXT_WIDTH = 320;       /* 件数表示の幅 / width of the count line */
    var COMPACT_BUTTON_FONT_SHRINK = 2; /* ボタン行以外のボタンの文字を小さくする量（pt）/ font shrink for buttons outside the button row */
    var COMPACT_BUTTON_TRIM_PX = 4;     /* そのボタンの高さを詰める量（px）/ height trim for those buttons */
    var STYLE_ROW_INDENT = "\u3000\u3000"; /* ウェイト（スタイル）行の字下げ / indent of style rows */
    var STANDARD_LEFT_WIDTH = 64;      /* ［規格］の左列（無印）の幅 / width of the left (plain) column under Standard */

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
        custom: true,
        customFonts: CUSTOM_FONTS, /* ［カスタム］のフォント名 / names under Custom */
        documentFonts: false,
        compositeFonts: false,
        foundries: [],          /* チェックしたメーカーの key / keys of the checked foundries */
        showAll: false,
        filterByStandard: true, /* ［規格］のチェックで絞り込む / filter by the Standard checkboxes */
        excludedLanguages: ["chinese", "korean", "thai", "multilingual"], /* ［日本語以外を除外］の言語 / excluded languages */
        hiddenSuffixes: [],     /* ［規格］で外した接尾辞（"" は［その他］）/ suffixes unchecked under Standard ("" is Other) */
        searchText: "",
        showPostScriptName: false,  /* 一覧を PostScript 名で表示 / list PostScript names */
        preview: true               /* 選んだフォントを仮に適用 / preview the chosen font */
    };

    /* フォント情報のキャッシュ（1行に PostScript 名・ファミリー名・スタイル名をタブ区切り）
       Font cache: one font per line, PostScript name, family and style separated by tabs */
    var FONT_CACHE_FILE = new File(Folder.userData + "/illustrator-scripts/" + SCRIPT_NAME + "-fonts.txt");

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
            progress: { ja: "フォント情報を読み込み中", en: "Reading font information" }
        },
        panel: {
            category: { ja: "分類", en: "Category" },
            standard: { ja: "規格", en: "Standard" },
            excludedLanguages: { ja: "日本語以外を除外", en: "Exclude non-Japanese" }
        },
        checkbox: {
            custom: { ja: "カスタム", en: "Custom" },
            documentFonts: { ja: "ドキュメント", en: "Document" },
            compositeFonts: { ja: "合成フォント", en: "Composite fonts" },
            showAll: { ja: "すべて表示", en: "Show all" },
            filterByStandard: { ja: "規格で絞り込む", en: "Filter by standard" },
            noSuffix: { ja: "規格なし", en: "No standard" },
            showPostScriptName: { ja: "PostScript名で表示", en: "Show PostScript names" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            search: { ja: "絞り込み", en: "Filter" }
        },
        button: {
            recordUsedFonts: { ja: "使用フォントをフォルダーに記録", en: "Record used fonts in folder" },
            rescan: { ja: "再スキャン", en: "Rescan" },
            addCustom: { ja: "カスタムに追加", en: "Add to Custom" },
            removeCustom: { ja: "カスタムから外す", en: "Remove from Custom" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" }
        },
        tooltip: {
            custom: {
                ja: "［カスタムに追加］で加えたフォント。規格違いの間引きと除外はしない",
                en: "Fonts added with Add to Custom, shown without thinning out standards or exclusions"
            },
            documentFonts: {
                ja: "ドキュメントと同じフォルダーの _ProjectFonts.txt に記録したフォント。規格違いも間引かずに表示",
                en: "Fonts recorded in _ProjectFonts.txt next to the document, shown without thinning out standards"
            },
            foundry: { ja: "ファミリー名・PostScript 名の先頭で判定", en: "Matched by the start of the family or PostScript name" },
            compositeFonts: {
                ja: "［書式］→［合成フォント］で作った合成フォント。規格の絞り込みと間引きはしない",
                en: "Composite fonts made with Type > Composite Fonts, not filtered or thinned out by standard"
            },
            showAll: {
                ja: "［分類］・規格違いの間引き・除外をやめてすべて表示（［規格］の絞り込みは効く）",
                en: "Ignore Category, standard thinning and exclusions and list every font (Standard still filters)"
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
            showPostScriptName: { ja: "一覧を PostScript 名（RyuminPr6N-Light など）で表示し、その順に並べる", en: "List PostScript names (such as RyuminPr6N-Light) and sort by them" },
            recordUsedFonts: {
                ja: "ドキュメントで使っているフォントを、同じフォルダーの _ProjectFonts.txt に追記する",
                en: "Add the fonts used in the document to _ProjectFonts.txt in the same folder"
            },
            rescan: { ja: "インストールされているフォントを読み直す", en: "Read the installed fonts again" },
            addCustom: { ja: "一覧で選んだフォントのファミリーを［カスタム］に加える", en: "Add the family of the chosen font to Custom" },
            removeCustom: { ja: "一覧で選んだフォントを［カスタム］から外す", en: "Remove the chosen font from Custom" },
            preview: {
                ja: "一覧で選んだフォントを選択中のテキストに仮に適用する（文字単位の選択と、連結したテキストは対象外）",
                en: "Preview the chosen font on the selected text (not for selected characters or threaded text)"
            },
            fontList: { ja: "ダブルクリックまたは Enter キーで、選択中のテキストに適用", en: "Double-click or press Enter to apply to the selected text" },
            soloClick: { ja: "option（Alt）＋クリック：これだけオン／すべてオン", en: "Option (Alt)-click: only this one / all on" }
        },
        message: {
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
            unsavedDocument: {
                ja: "ドキュメントが保存されていません。\n一度保存してから実行してください。",
                en: "The document has not been saved.\nSave it once and try again."
            },
            noNewFonts: { ja: "ドキュメントに新しいフォントは見つかりませんでした。", en: "No new fonts were found in the document." },
            recorded: { ja: "{count} 件のフォントを {file} に記録しました。", en: "Recorded {count} fonts in {file}." },
            writeFailed: { ja: "ファイルを書き出せませんでした：{message}", en: "Could not write the file: {message}" },
            rescanned: { ja: "フォント情報を更新しました（{count} 件）。", en: "Font information updated ({count} fonts)." },
            alreadyCustom: { ja: "「{name}」はすでに［カスタム］に入っています。", en: "\"{name}\" is already in Custom." },
            notInCustom: { ja: "「{name}」は［カスタム］に入っていません。", en: "\"{name}\" is not in Custom." },
            confirmRemoveCustom: {
                ja: "［カスタム］から「{names}」を外します。\nこの名前で始まるほかのフォントも外れます。",
                en: "Remove \"{names}\" from Custom?\nOther fonts starting with this name are removed too."
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

    /**
     * メーカー別の分類に当たるか
     * @param {Object} fontInfo - createFontCatalog() 済みのフォント情報
     * @param {Object} foundryFilter - FOUNDRY_FILTERS の1件
     * @returns {boolean} 当たれば true
     */
    function matchesFoundry(fontInfo, foundryFilter) {
        return startsWithAny(fontInfo.familyLower, foundryFilter.familyPrefixes || []) ||
            startsWithAny(fontInfo.psLower, foundryFilter.psNamePrefixes || []);
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
        for (var k = 0; k < filterSettings.excludedLanguages.length; k++) excludedLanguages[filterSettings.excludedLanguages[k]] = true;
        return {
            excludedLanguages: excludedLanguages,
            custom: filterSettings.custom,
            customFontNames: filterSettings.customFonts,
            documentFonts: filterSettings.documentFonts,
            compositeFonts: filterSettings.compositeFonts,
            documentFontNames: documentFontNames,
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
        if (fontNameContainsAny(fontInfo, EXCLUDE_FONTS) && !fontNameContainsAny(fontInfo, RESCUE_FONTS)) return false;

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

    /**
     * ドキュメントで使っているフォントのファミリー名と PostScript 名を集める
     * @param {Document} targetDoc - 対象のドキュメント
     * @returns {string[]} 名前の一覧（重複なし）
     */
    function collectDocumentFontNames(targetDoc) {
        var fontNames = [];

        /**
         * 名前を一度だけ加える
         * @param {string} fontName - 名前
         * @returns {void}
         */
        function addFontName(fontName) {
            if (fontName && !containsValue(fontNames, fontName)) fontNames.push(fontName);
        }

        var stories = targetDoc.stories;
        for (var i = 0; i < stories.length; i++) {
            var textRanges = stories[i].textRanges;
            for (var j = 0; j < textRanges.length; j++) {
                var textFont;
                try {
                    textFont = textRanges[j].characterAttributes.textFont;
                } catch (e) {
                    /* 環境にないフォントの範囲は textFont が読めない / textFont cannot be read for ranges in a missing font */
                    continue;
                }
                addFontName(textFont.family);
                /* 合成フォントの PostScript 名（ATC-…）は記録しない（合成フォント名で照合する）/ Skip ATC-… names; composites match by their own name */
                if (!/^ATC-/i.test(textFont.name)) addFontName(textFont.name);
            }
        }
        return fontNames;
    }

    /**
     * ドキュメントで使っているフォントを _ProjectFonts.txt に追記する（結果は alert で知らせる）
     * @returns {string[]|null} 追記後の名前の一覧。追記しなかったときは null
     */
    function recordUsedFonts() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return null;
        }
        var projectFontsFile = getProjectFontsFile();
        if (!projectFontsFile) {
            alert(getLabel("alert.unsavedDocument"));
            return null;
        }
        var documentFontNames = readProjectFontNames();
        var foundNames = collectDocumentFontNames(app.activeDocument);
        var addedCount = 0;
        for (var i = 0; i < foundNames.length; i++) {
            if (containsValue(documentFontNames, foundNames[i])) continue;
            documentFontNames.push(foundNames[i]);
            addedCount++;
        }
        if (addedCount === 0) {
            alert(getLabel("alert.noNewFonts"));
            return null;
        }
        if (!writeTextFile(projectFontsFile, documentFontNames.join("\n"))) {
            alert(getLabel("alert.writeFailed", { message: projectFontsFile.error }));
            return null;
        }
        alert(getLabel("alert.recorded", { count: addedCount, file: PROJECT_FONTS_FILE_NAME }));
        return documentFontNames;
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

        var categoryPanel = filterColumn.add("panel", undefined, getLabel("panel.category"));
        setupPanel(categoryPanel, 6);
        var chkCustom = addCheckbox(categoryPanel, getLabel("checkbox.custom"), "tooltip.custom");
        var chkDocumentFonts = addCheckbox(categoryPanel, getLabel("checkbox.documentFonts"), "tooltip.documentFonts");
        var foundryCheckboxes = [];
        for (var i = 0; i < FOUNDRY_FILTERS.length; i++) {
            var chkFoundry = addCheckbox(categoryPanel, getLabel(FOUNDRY_FILTERS[i].label), "tooltip.foundry");
            chkFoundry.foundryKey = FOUNDRY_FILTERS[i].key;
            foundryCheckboxes.push(chkFoundry);
        }
        var chkCompositeFonts = addCheckbox(categoryPanel, getLabel("checkbox.compositeFonts"), "tooltip.compositeFonts");
        var btnRecordUsedFonts = addCompactButton(categoryPanel, getLabel("button.recordUsedFonts"));
        btnRecordUsedFonts.alignment = "left";
        btnRecordUsedFonts.helpTip = getLabel("tooltip.recordUsedFonts");

        var chkShowAll = addCheckbox(filterColumn, getLabel("checkbox.showAll"), "tooltip.showAll");
        var chkFilterByStandard = addCheckbox(filterColumn, getLabel("checkbox.filterByStandard"), "tooltip.filterByStandard");

        var standardPanel = filterColumn.add("panel", undefined, getLabel("panel.standard"));
        setupPanel(standardPanel, 6);
        var suffixCheckboxes = [];
        var suffixRows = buildSuffixRows(PRIORITY_SUFFIXES);
        for (var j = 0; j < suffixRows.length; j++) {
            var suffixRowGroup = standardPanel.add("group");
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
        var chkNoSuffix = addCheckbox(standardPanel, getLabel("checkbox.noSuffix"), "tooltip.noSuffix");
        chkNoSuffix.suffix = "";
        suffixCheckboxes.push(chkNoSuffix);

        var languagePanel = filterColumn.add("panel", undefined, getLabel("panel.excludedLanguages"));
        setupPanel(languagePanel, 6);
        var languageCheckboxes = [];
        for (var m = 0; m < LANGUAGE_GROUPS.length; m++) {
            var chkLanguage = languagePanel.add("checkbox", undefined, getLabel(LANGUAGE_GROUPS[m].label));
            chkLanguage.helpTip = getLabel(LANGUAGE_GROUPS[m].tooltip);
            chkLanguage.languageKey = LANGUAGE_GROUPS[m].key;
            languageCheckboxes.push(chkLanguage);
        }

        return {
            categoryPanel: categoryPanel,
            chkCustom: chkCustom,
            chkDocumentFonts: chkDocumentFonts,
            chkCompositeFonts: chkCompositeFonts,
            foundryCheckboxes: foundryCheckboxes,
            btnRecordUsedFonts: btnRecordUsedFonts,
            chkShowAll: chkShowAll,
            chkFilterByStandard: chkFilterByStandard,
            standardPanel: standardPanel,
            suffixCheckboxes: suffixCheckboxes,
            languageCheckboxes: languageCheckboxes
        };
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
     * ボタン行以外に置く、ひとまわり小さいボタンを足す（文字を小さくする。高さは表示したときに詰める）
     * @param {Panel|Group} parentContainer - 足す先
     * @param {string} buttonText - ボタンの文言
     * @returns {Button} 足したボタン
     */
    function addCompactButton(parentContainer, buttonText) {
        var compactButton = parentContainer.add("button", undefined, buttonText);
        var baseFont = compactButton.graphics.font;
        compactButton.graphics.font = ScriptUI.newFont(baseFont.name, "REGULAR", baseFont.size - COMPACT_BUTTON_FONT_SHRINK);
        return compactButton;
    }

    /**
     * 小さいボタンの高さを詰め、その高さを preferredSize に入れて固定する（リサイズで元の高さに戻らないように）
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {Button[]} compactButtons - addCompactButton() で作ったボタン
     * @returns {void}
     */
    function fixCompactButtonHeights(targetWindow, compactButtons) {
        for (var i = 0; i < compactButtons.length; i++) {
            var compactButton = compactButtons[i];
            if (!compactButton.size) continue;
            compactButton.preferredSize = [compactButton.size.width, compactButton.size.height - COMPACT_BUTTON_TRIM_PX];
        }
        targetWindow.layout.layout(true);
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
     * @returns {Object} コントロール一式（searchInput・fontListBox・btnAddCustom・btnRemoveCustom・chkShowPostScriptName・chkPreview・statusText）
     */
    function buildFontListColumn(parentGroup) {
        var fontListColumn = parentGroup.add("group");
        fontListColumn.orientation = "column";
        fontListColumn.alignChildren = ["fill", "top"];
        fontListColumn.alignment = ["fill", "top"];
        fontListColumn.spacing = 6;

        var searchRowGroup = fontListColumn.add("group");
        setupRow(searchRowGroup, "fill", 6);
        searchRowGroup.add("statictext", undefined, labelText("fieldLabel.search"));
        var searchInput = searchRowGroup.add("edittext", undefined, "");
        searchInput.alignment = ["fill", "center"];
        searchInput.helpTip = getLabel("tooltip.search");

        var fontListBox = fontListColumn.add("listbox", undefined, [], { multiselect: false });
        fontListBox.preferredSize = FONT_LIST_SIZE;
        fontListBox.alignment = ["fill", "top"];
        fontListBox.helpTip = getLabel("tooltip.fontList");

        var customRowGroup = fontListColumn.add("group");
        setupRow(customRowGroup, "left", 6);
        var btnAddCustom = addCompactButton(customRowGroup, getLabel("button.addCustom"));
        btnAddCustom.helpTip = getLabel("tooltip.addCustom");
        var btnRemoveCustom = addCompactButton(customRowGroup, getLabel("button.removeCustom"));
        btnRemoveCustom.helpTip = getLabel("tooltip.removeCustom");

        var optionRowGroup = fontListColumn.add("group");
        setupRow(optionRowGroup, "left", PANEL_SPACING);
        var chkShowPostScriptName = addCheckbox(optionRowGroup, getLabel("checkbox.showPostScriptName"), "tooltip.showPostScriptName");
        var chkPreview = addCheckbox(optionRowGroup, getLabel("checkbox.preview"), "tooltip.preview");

        var statusText = fontListColumn.add("statictext", undefined, "");
        statusText.preferredSize.width = STATUS_TEXT_WIDTH;
        statusText.alignment = ["fill", "top"];

        return {
            searchInput: searchInput,
            fontListBox: fontListBox,
            btnAddCustom: btnAddCustom,
            btnRemoveCustom: btnRemoveCustom,
            chkShowPostScriptName: chkShowPostScriptName,
            chkPreview: chkPreview,
            statusText: statusText
        };
    }

    /**
     * 保存形式の設定をダイアログに反映する
     * @param {Object} filterControls - buildFilterColumn() の戻り値に searchInput・chkShowPostScriptName・chkPreview・customFontNames を足したもの
     * @param {Object} filterSettings - DEFAULT_SETTINGS と同じ形
     * @returns {void}
     */
    function applySettingsToControls(filterControls, filterSettings) {
        filterControls.chkCustom.value = filterSettings.custom;
        filterControls.customFontNames = filterSettings.customFonts;
        filterControls.chkDocumentFonts.value = filterSettings.documentFonts;
        filterControls.chkCompositeFonts.value = filterSettings.compositeFonts;
        for (var i = 0; i < filterControls.foundryCheckboxes.length; i++) {
            var chkFoundry = filterControls.foundryCheckboxes[i];
            chkFoundry.value = containsValue(filterSettings.foundries, chkFoundry.foundryKey);
        }
        filterControls.chkShowAll.value = filterSettings.showAll;
        filterControls.chkFilterByStandard.value = filterSettings.filterByStandard;
        for (var k = 0; k < filterControls.languageCheckboxes.length; k++) {
            var chkLanguage = filterControls.languageCheckboxes[k];
            chkLanguage.value = containsValue(filterSettings.excludedLanguages, chkLanguage.languageKey);
        }
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
        var excludedLanguageKeys = [];
        for (var k = 0; k < filterControls.languageCheckboxes.length; k++) {
            var chkLanguage = filterControls.languageCheckboxes[k];
            if (chkLanguage.value) excludedLanguageKeys.push(chkLanguage.languageKey);
        }
        return {
            excludedLanguages: excludedLanguageKeys,
            custom: filterControls.chkCustom.value,
            customFonts: filterControls.customFontNames,
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
        filterControls.chkPreview = fontListControls.chkPreview;

        var buttonRow = addButtonRow(mainDialog);
        var btnRescan = buttonRow.leftGroup.add("button", undefined, getLabel("button.rescan"));
        btnRescan.helpTip = getLabel("tooltip.rescan");
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLabel("button.apply"));
        alignRightOnlyButtonRow(buttonRow);

        applySettingsToControls(filterControls, settingsStore.load(DEFAULT_SETTINGS));

        /**
         * ［すべて表示］のときは［分類］を、［規格で絞り込む］がオフのときは［規格］を無効にする
         * @returns {void}
         */
        function updatePanelStates() {
            filterControls.categoryPanel.enabled = !filterControls.chkShowAll.value;
            filterControls.standardPanel.enabled = filterControls.chkFilterByStandard.value;
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
         * 一覧で選んだフォントのファミリーを［カスタム］に加える
         * @returns {void}
         */
        function addSelectedToCustom() {
            if (!fontListBox.selection) return;
            var familyName = fontListBox.selection.familyName;
            if (containsValue(filterControls.customFontNames, familyName)) {
                alert(getLabel("alert.alreadyCustom", { name: familyName }));
                return;
            }
            filterControls.customFontNames = filterControls.customFontNames.concat([familyName]);
            filterControls.chkCustom.value = true;
            refreshFontList();
        }

        /**
         * 一覧で選んだフォントに当たる名前を［カスタム］から外す（ほかのフォントにも当たる名前なら確かめる）
         * @returns {void}
         */
        function removeSelectedFromCustom() {
            if (!fontListBox.selection) return;
            var selectedItem = fontListBox.selection;
            var familyLower = selectedItem.familyName.toLowerCase();
            var psLower = selectedItem.psName.toLowerCase();
            var keptNames = [];
            var removedNames = [];
            for (var i = 0; i < filterControls.customFontNames.length; i++) {
                var customName = filterControls.customFontNames[i];
                var isMatch = startsWithAny(familyLower, [customName]) || startsWithAny(psLower, [customName]);
                if (isMatch) removedNames.push(customName);
                else keptNames.push(customName);
            }
            if (removedNames.length === 0) {
                alert(getLabel("alert.notInCustom", { name: selectedItem.familyName }));
                return;
            }
            var isOnlyThisFamily = (removedNames.length === 1 && removedNames[0].toLowerCase() === familyLower);
            if (!isOnlyThisFamily && !confirm(getLabel("alert.confirmRemoveCustom", { names: removedNames.join(uiLang === "ja" ? "」「" : "\", \"") }))) return;
            filterControls.customFontNames = keptNames;
            refreshFontList();
        }

        addSoloClickGroup([filterControls.chkCustom, filterControls.chkDocumentFonts].concat(filterControls.foundryCheckboxes, [filterControls.chkCompositeFonts]), refreshFontList);
        addSoloClickGroup(filterControls.suffixCheckboxes, refreshFontList);
        for (var n = 0; n < filterControls.languageCheckboxes.length; n++) filterControls.languageCheckboxes[n].onClick = refreshFontList;
        filterControls.chkShowAll.onClick = filterControls.chkFilterByStandard.onClick = function () {
            updatePanelStates();
            refreshFontList();
        };
        filterControls.chkShowPostScriptName.onClick = refreshFontList;
        filterControls.chkPreview.onClick = refreshPreview;
        fontListControls.btnAddCustom.onClick = addSelectedToCustom;
        fontListControls.btnRemoveCustom.onClick = removeSelectedFromCustom;
        fontListBox.onChange = function () {
            if (isUpdatingList || skipFamilyHeader()) return;
            refreshPreview();
        };
        /* 一覧に並べるのが重いので、件数が多いうちは Enter（onChange）まで待つ / Adding many rows is slow, so wait for Enter while the result is large */
        filterControls.searchInput.onChanging = function () {
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

        filterControls.btnRecordUsedFonts.onClick = function () {
            /* プレビューの複製を拾わないよう、記録の間は消しておく / Remove preview duplicates so they are not recorded */
            clearFontPreview(fontPreview);
            var updatedNames = recordUsedFonts();
            refreshPreview();
            if (!updatedNames) return;
            documentFontNames = updatedNames;
            filterControls.chkDocumentFonts.value = true;
            refreshFontList();
        };
        btnRescan.onClick = function () {
            fontCatalog = createFontCatalog(scanInstalledFonts());
            refreshFontList();
            alert(getLabel("alert.rescanned", { count: fontCatalog.length }));
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
            fixCompactButtonHeights(mainDialog, [filterControls.btnRecordUsedFonts, fontListControls.btnAddCustom, fontListControls.btnRemoveCustom]);
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
