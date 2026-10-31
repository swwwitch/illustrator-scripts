#target illustrator
#targetengine "AiFontConverterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.AiFontConverter
Name=フォントの種別をまとめて変更
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.8
Release-Date=2026-06-16
Description-BEGIN
フォントの種別（文字セット・P・UD・N・NT・ウェイト）をまとめて変更します。
対象は選択オブジェクト／ドキュメント全体／アクティブアートボードから選べます。

Changes font variants — character set, P, UD, N, NT and weight — across the artwork in one pass.
The scope can be the selection, the whole document, or the active artboard.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

フォントの種別（文字セット・P・UD・N・NT・ウェイト）をまとめて変更します。
対象は選択オブジェクト／ドキュメント全体／アクティブアートボードから選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFontConverter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n261c771b4b41

### Overview

Changes font variants — character set, P, UD, N, NT and weight — across the artwork in one pass.
The scope can be the selection, the whole document, or the active artboard.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiFontConverter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiFontConverter";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-24";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "フォントの種別をまとめて変更"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFontConverter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiFontConverter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n261c771b4b41"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/AiFontConverter.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 実行前に変更内容を確認するか（UI 非表示）/ Whether to confirm changes before running (hidden) */
    var CONFIRM_BEFORE_RUN = true;

    /* CID フォント（文字セット表記なし）を OTF へ変換するか / Whether to convert CID fonts to OTF */
    var CONVERT_CID_TO_OTF = true;

    /* G-OTF 学参書体を A-OTF の通常書体に統合してから変換するか（UI 非表示）/ Merge G-OTF Gakusan fonts into A-OTF before converting (hidden) */
    var INTEGRATE_GAKUSAN_TO_STANDARD = true;

    /* G-OTF 学参書体の PostScript 名プレフィックス（GJ=常改 / G=学参 / K=K書体）。長い順に試す / Gakusan PostScript prefixes, tried longest first */
    var GAKUSAN_PREFIXES = ["GJ", "G", "K"];

    /* 特殊シリーズ（太さが等価なメンバーを完全名で結ぶ）/ Special series: members of equal thickness linked by full PostScript name
       A1明朝は Std が Bold 1 ウェイトのみで、A P-OTF StdN の Regular と同じ太さ。これを同一シリーズとして対応付ける。
       members のキーは文字セット（+N）。汎用ロジックではなくこの表だけで変換する。 */
    var SPECIAL_SERIES = [
        {
            label: "A1明朝",
            members: {
                "Std": "A1MinchoStd-Bold",
                "StdN": "PA1MinchoStdN-Regular"
            }
        }
    ];

    /* CID 変換時、文字セットが「現状維持」のときに補完する文字セット / Charset filled in when converting CID with "keep" */
    var CID_FALLBACK_CHARSET = "Pr6";

    /* 文字セットを収録文字数の多い順に並べたランキング / Charsets ranked by glyph count, richest first
       Max  は N なしの種類から、MaxN は N あり込みから、最初に選べる文字セットを採用する */
    var CHARSET_RANK_NO_N = ["Pr6", "Pr5", "Pro", "Std", "Min2", "Min"];
    var CHARSET_RANK_WITH_N = ["Pr6N", "Pr6", "Pr5N", "Pr5", "ProN", "Pro", "StdN", "Std", "Min2", "Min"];

    /* 近いウェイトを探す並び順（軽い → 重い）/ Weight order for nearest-weight search (light to heavy) */
    var WEIGHT_ORDER = [
        "Light", "Regular", "Medium",
        "DemiBold", "Demi", "DeBold",
        "Bold", "ExtraBold", "ExBold",
        "Heavy", "ExtraHeavy", "ExHeavy", "Ultra"
    ];

    /* 変換対象のフォントファミリー定義（フォントデータベースから生成）/ Target font families (generated from the font database)
       regex の最後のグループをウェイト、それ以外を値（P / UD / Pro|Pr5|Pr6 / N）で自動判別する */
    var FONT_FAMILIES = [
        { label: "見出ミンMA31", baseName: "MidashiMinMA31", regex: /^(P)?MidashiMinMA31(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "見出ゴMB31", baseName: "MidashiGoMB31", regex: /^(P)?MidashiGoMB31(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "太ミンA101", baseName: "FutoMinA101", regex: /^(P)?FutoMinA101(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "ゴシックMB101", baseName: "GothicMB101", regex: /^(P)?GothicMB101(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD新ゴ コンデンス90", baseName: "ShinGoCOniz", regex: /^(P)?(UD)?ShinGoCOniz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD新ゴ コンデンス80", baseName: "ShinGoCOeiz", regex: /^(P)?(UD)?ShinGoCOeiz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD新ゴ コンデンス70", baseName: "ShinGoCOsez", regex: /^(P)?(UD)?ShinGoCOsez(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD新ゴ コンデンス60", baseName: "ShinGoCOsiz", regex: /^(P)?(UD)?ShinGoCOsiz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD新ゴ コンデンス50", baseName: "ShinGoCOfiz", regex: /^(P)?(UD)?ShinGoCOfiz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "見出ミン", baseName: "MidashiMin", regex: /^MidashiMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "太ゴB101", baseName: "FutoGoB101", regex: /^(P)?FutoGoB101(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "新正楷書CBSK1", baseName: "ShinseiKai", regex: /^ShinseiKai(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "凸版文久明朝", baseName: "BunkyuMin", regex: /^(P)?BunkyuMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "中ゴシックBBB", baseName: "GothicBBB", regex: /^(P)?GothicBBB(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "見出ゴ", baseName: "MidashiGo", regex: /^MidashiGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "正楷書CB1", baseName: "SeiKaiCB1", regex: /^(P)?SeiKaiCB1(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "黎ミンY10", baseName: "ReimYonz", regex: /^(P)?ReimYonz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "黎ミンY20", baseName: "ReimYtwz", regex: /^(P)?ReimYtwz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "黎ミンY30", baseName: "ReimYthz", regex: /^(P)?ReimYthz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "黎ミンY40", baseName: "ReimYfoz", regex: /^(P)?ReimYfoz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "秀英明朝", baseName: "ShueiMin", regex: /^(P)?ShueiMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "凸版文久ゴシック", baseName: "BunkyuGo", regex: /^(P)?BunkyuGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "新丸ゴ", baseName: "ShinMGo", regex: /^(P)?(UD)?ShinMGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "リュウミン", baseName: "Ryumin", regex: /^(P)?Ryumin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "新ゴ / 新ゴNT", baseName: "ShinGo", regex: /^(P)?(UD)?ShinGo(NT)?(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD黎ミン", baseName: "Reimin", regex: /^(P)?(UD)?Reimin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "游明朝体", baseName: "YuMin", regex: /^YuMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "黎ミン", baseName: "Reim", regex: /^(P)?Reim(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "游ゴシック体", baseName: "YuGo", regex: /^YuGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },

        // --- フォントワークス書体 ---
        { label: "セザンヌ", baseName: "Cezanne", regex: /^(P)?Cezanne(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "DNP 秀英明朝", baseName: "FOT_DNPShueiMin", regex: /^(P)?FOT_DNPShueiMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "ハミング", baseName: "Humming", regex: /^(P)?Humming(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "マティス", baseName: "Matisse", regex: /^(P)?Matisse(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "モード明朝Aラージ", baseName: "ModeMinALarge", regex: /^(P)?ModeMinALarge(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "モード明朝Bラージ", baseName: "ModeMinBLarge", regex: /^(P)?ModeMinBLarge(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "ニューセザンヌ", baseName: "NewCezanne", regex: /^(P)?NewCezanne(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "ニューグレコ", baseName: "NewGreco", regex: /^(P)?NewGreco(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "ニューロダン", baseName: "NewRodin", regex: /^(P)?NewRodin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "ロダン", baseName: "Rodin", regex: /^(P)?Rodin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "スーラ", baseName: "Seurat", regex: /^(P)?Seurat(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "スキップ", baseName: "Skip", regex: /^(P)?Skip(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "テロップ明朝", baseName: "TelopMin", regex: /^(P)?TelopMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫A見出ミン", baseName: "TsukuAMDMin", regex: /^(P)?TsukuAMDMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Aオールド明朝", baseName: "TsukuAOldMin", regex: /^(P)?TsukuAOldMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Aヴィンテージ明L", baseName: "TsukuAVintageMinL", regex: /^(P)?TsukuAVintageMinL(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Aヴィンテージ明S", baseName: "TsukuAVintageMinS", regex: /^(P)?TsukuAVintageMinS(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫B見出ミン", baseName: "TsukuBMDMin", regex: /^(P)?TsukuBMDMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫B明朝", baseName: "TsukuBMin", regex: /^(P)?TsukuBMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Bオールド明朝", baseName: "TsukuBOldMin", regex: /^(P)?TsukuBOldMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Bヴィンテージ明L", baseName: "TsukuBVintageMinL", regex: /^(P)?TsukuBVintageMinL(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Bヴィンテージ明S", baseName: "TsukuBVintageMinS", regex: /^(P)?TsukuBVintageMinS(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫C見出ミン", baseName: "TsukuCMDMin", regex: /^(P)?TsukuCMDMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Cオールド明朝", baseName: "TsukuCOldMin", regex: /^(P)?TsukuCOldMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Cヴィンテージ明L", baseName: "TsukuCVintageMinL", regex: /^(P)?TsukuCVintageMinL(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫Cヴィンテージ明S", baseName: "TsukuCVintageMinS", regex: /^(P)?TsukuCVintageMinS(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫ゴシック", baseName: "TsukuGo", regex: /^(P)?TsukuGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫明朝", baseName: "TsukuMin", regex: /^(P)?TsukuMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "筑紫新聞明朝", baseName: "TsukuNewsMin", regex: /^(P)?TsukuNewsMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD角ゴ_ラージ", baseName: "UDKakugo_Large", regex: /^(P)?UDKakugo_Large(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD角ゴ_スモール", baseName: "UDKakugo_Small", regex: /^(P)?UDKakugo_Small(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD丸ゴ_ラージ", baseName: "UDMarugo_Large", regex: /^(P)?UDMarugo_Large(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD丸ゴ_スモール", baseName: "UDMarugo_Small", regex: /^(P)?UDMarugo_Small(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
        { label: "UD明朝", baseName: "UDMincho", regex: /^(P)?UDMincho(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ }
    ];

    /* AXIS（Type Project）は文字セット体系が Morisawa 系と異なるため専用処理する / AXIS uses a different charset scheme; handled separately
       - 幅（Basic / Cond / Comp）と Joyo は保持する
       - N と Std⇄Pro だけ切り替える（AXIS に無い Pr5 / Pr6 を選んでも現状維持）
       - P / UD / NT は AXIS に無いので無視する
       - Max / MaxN プリセット時は設定に関わらず ProN（収録最多）へ寄せる
       グループ: 1=幅(Basic|Cond|Comp) / 2=文字セット(Std|Pro|Joyo) / 3=N / 4=ウェイト */
    var AXIS_FAMILY = {
        label: "AXIS",
        baseName: "Axis",
        regex: /^Axis(Basic|Cond|Comp)?(Std|Pro|Joyo)?(N)?-(.+)$/
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var CONFIRM_VISIBLE_ROWS = 16; /* 確認ダイアログに一度に並べる行数（見出しを含む）/ rows shown at once in the confirmation dialog, headings included */
    var CONFIRM_SCROLLBAR_WIDTH = 16; /* 確認ダイアログのスクロールバーの幅 / width of the confirmation dialog scrollbar */
    var CONFIRM_ROW_SPACING = 12; /* 確認ダイアログの行の間隔 / spacing between rows of the confirmation dialog */
    var CONFIRM_COLUMN_GAP = 4; /* 確認ダイアログの「変更前」と「変更後」の間隔 / gap between the before and after columns of the confirmation dialog */

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

    /**
     * グループの共通設定（row は縦中央、column は左揃えで子を並べる）
     * @param {Group} group - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

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

    // =========================================
    // ダイアログの位置 / Dialog position
    // =========================================

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

    var LABELS = {
        dialog: {
            title: { ja: "フォント種別変更", en: "Font Variant Switcher" }
        },
        panel: {
            target: { ja: "対象範囲", en: "Target Scope" },
            conversion: { ja: "変換設定", en: "Conversion Settings" },
            charset: { ja: "文字セット", en: "Character Set" },
            nSetting: { ja: "N（JIS2004対応）", en: "N (JIS2004)" },
            ntSetting: { ja: "新ゴ / 新ゴNT", en: "ShinGo / ShinGo NT" },
            udSetting: { ja: "UD", en: "UD" },
            pSetting: { ja: "AP版", en: "AP (Proportional)" },
            options: { ja: "処理オプション", en: "Processing Options" }
        },
        radio: {
            targetSelection: { ja: "選択オブジェクト", en: "Selected objects" },
            targetDocument: { ja: "ドキュメント全体", en: "Entire document" },
            targetArtboard: { ja: "アクティブアートボード", en: "Active artboard" },
            keep: { ja: "現状維持", en: "Keep current" },
            nOff: { ja: "Nなし", en: "Without N" },
            nOn: { ja: "Nあり", en: "With N" },
            ntOff: { ja: "NTなし", en: "Without NT" },
            ntOn: { ja: "NTあり", en: "With NT" },
            udOff: { ja: "UDなし", en: "Without UD" },
            udOn: { ja: "UDあり", en: "With UD" },
            pOff: { ja: "Pなし", en: "Without P" },
            pOn: { ja: "Pあり", en: "With P" }
        },
        checkbox: {
            integrateGakusan: {
                ja: "G-OTF学参書体をA-OTFに統合",
                en: "Merge G-OTF Gakusan fonts into A-OTF"
            },
            includeStyles: {
                ja: "段落/文字スタイルも変更",
                en: "Also update paragraph/character styles"
            },
            includeLocked: {
                ja: "ロックされたオブジェクトも対象にする",
                en: "Include locked objects"
            },
            includeHidden: {
                ja: "非表示オブジェクトも対象にする",
                en: "Include hidden objects"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" },
            run: { ja: "実行", en: "Run" }
        },
        confirm: {
            title: { ja: "変更内容の確認", en: "Confirm Changes" },
            willChange: {
                ja: "以下のフォントを変更します。",
                en: "The following fonts will be changed."
            },
            nearWeight: {
                ja: "同じウェイトのフォントが見つかりません。近いウェイトに置換します。",
                en: "Exact weight not found. The nearest weight will be substituted."
            },
            notInstalled: {
                ja: "次のフォントはインストールされていないため変更されません。",
                en: "The following fonts are not installed and will not be changed."
            },
            prompt: { ja: "実行しますか？", en: "Run now?" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: {
                ja: "テキストオブジェクトを選択してください。",
                en: "Please select text objects."
            },
            noTarget: {
                ja: "対象となるテキストが見つかりませんでした。",
                en: "No target text was found."
            },
            noChange: {
                ja: "変更対象のフォントが見つかりませんでした。",
                en: "No fonts to change were found."
            },
            done: { ja: "処理が完了しました。", en: "Done." },
            changedCount: { ja: "変更数", en: "Changed" }
        },
        tooltip: {
            target: {
                ja: "フォント種別を変更する範囲を選びます。選択オブジェクト、ドキュメント全体、アクティブアートボード内から選択できます。",
                en: "Choose the scope for changing font variants: selected objects, the entire document, or the active artboard."
            },
            charset: {
                ja: "文字セットを Std / Pro / Pr5 / Pr6 から選びます。N の有無は別の「N設定」で切り替えます。",
                en: "Choose the character set: Std, Pro, Pr5, or Pr6. Toggle the N variant separately under N Variant."
            },
            nSetting: {
                ja: "N あり/なしを切り替えます。例：Pr6 → Pr6N。現状維持では現在の N 状態を保ちます。",
                en: "Toggle the N variant on or off, for example Pr6 -> Pr6N. Keep current preserves the current N state."
            },
            udSetting: {
                ja: "UD 書体があるファミリーだけ、UD あり/なしを切り替えます。対応しないファミリーでは無視されます。",
                en: "Toggle the UD variant only for families that support it. Unsupported families are ignored."
            },
            pSetting: {
                ja: "AP版書体",
                en: "A P-OTF (proportional) version"
            },
            includeLocked: {
                ja: "ロック中のオブジェクトも一時的に解除して変更し、処理後に元のロック状態へ戻します。",
                en: "Temporarily unlock locked objects, update them, then restore the original lock state."
            },
            includeHidden: {
                ja: "非表示のオブジェクトも一時的に表示して変更し、処理後に元の表示状態へ戻します。",
                en: "Temporarily show hidden objects, update them, then restore the original visibility state."
            },
            nt: {
                ja: "新ゴと新ゴNTを切り替えます。対象は新ゴ系ファミリーのみです。",
                en: "Switch between ShinGo and ShinGo NT. This applies only to ShinGo families."
            },
            integrateGakusan: {
                ja: "G-OTF の学参・常改・K書体を、通常の A-OTF 書体として扱って変換します。",
                en: "Treat G-OTF Gakusan, revised, and K fonts as standard A-OTF fonts before conversion."
            },
            includeStyles: {
                ja: "段落スタイル・文字スタイルに設定されたフォントも、同じ条件で変更します。",
                en: "Also update fonts assigned in paragraph and character styles using the same conversion settings."
            },
            presetMax: {
                ja: "N なしで収録文字数が最も多い文字セットを選び、UD・P・NT をありにします。",
                en: "Select the richest available character set without N, and turn on UD, P, and NT."
            },
            presetMaxN: {
                ja: "N ありを含めて収録文字数が最も多い文字セットを選び、UD・P・NT をありにします。",
                en: "Select the richest available character set including N variants, and turn on UD, P, and NT."
            }
        }
    };

    // =========================================
    // メイン / Main
    // =========================================

    (function () {

        /* ドキュメントの有無を確認 / Ensure a document is open */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var documentRef = app.activeDocument;

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        var mainDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        /* 対象パネル（ラジオは縦並び）/ Target panel (radios in a column) */
        var targetPanel = addPanel(mainDialog, "panel.target", 6);
        targetPanel.helpTip = getLabel("tooltip.target");
        var rbTargetSelection = targetPanel.add("radiobutton", undefined, getLabel("radio.targetSelection"));
        var rbTargetDocument = targetPanel.add("radiobutton", undefined, getLabel("radio.targetDocument"));
        var rbTargetArtboard = targetPanel.add("radiobutton", undefined, getLabel("radio.targetArtboard"));
        rbTargetSelection.value = true;

        /* 変換設定パネル（文字セット/N/UD/P と Max/MaxN をまとめる）/ Conversion panel (variant columns + presets) */
        var conversionPanel = addPanel(mainDialog, "panel.conversion", 6);

        /* 変換設定: 3 カラム（左=文字セット / 中央=N・NT / 右=UD・P）/ Conversion: 3 columns */
        var variantColumns = conversionPanel.add("group");
        setupGroup(variantColumns, "row");
        variantColumns.alignChildren = ["fill", "top"];

        /* 左カラム: 文字セットパネル（Std/Pro/Pr5/Pr6 はローカライズ不要の固有名）/ Left: character set */
        var charsetPanel = addPanel(variantColumns, "panel.charset", 6);
        charsetPanel.helpTip = getLabel("tooltip.charset");
        var rbCharsetKeep = charsetPanel.add("radiobutton", undefined, getLabel("radio.keep"));
        var rbCharsetStd = charsetPanel.add("radiobutton", undefined, "Std");
        var rbCharsetPro = charsetPanel.add("radiobutton", undefined, "Pro");
        var rbCharsetPr5 = charsetPanel.add("radiobutton", undefined, "Pr5");
        var rbCharsetPr6 = charsetPanel.add("radiobutton", undefined, "Pr6");
        rbCharsetKeep.value = true;

        /* 中央カラム: N 設定・NT 設定 / Center: N and NT */
        var nNtColumn = variantColumns.add("group");
        setupGroup(nNtColumn, "column");

        var nVariantPanel = addPanel(nNtColumn, "panel.nSetting", 6);
        nVariantPanel.helpTip = getLabel("tooltip.nSetting");
        var rbNKeep = nVariantPanel.add("radiobutton", undefined, getLabel("radio.keep"));
        var rbNOff = nVariantPanel.add("radiobutton", undefined, getLabel("radio.nOff"));
        var rbNOn = nVariantPanel.add("radiobutton", undefined, getLabel("radio.nOn"));
        rbNKeep.value = true;

        /* NT 設定パネル（新ゴ ⇄ 新ゴNT）/ NT panel (ShinGo <-> ShinGo NT) */
        var ntPanel = addPanel(nNtColumn, "panel.ntSetting", 6);
        ntPanel.helpTip = getLabel("tooltip.nt");
        var rbNTKeep = ntPanel.add("radiobutton", undefined, getLabel("radio.keep"));
        var rbNTOff = ntPanel.add("radiobutton", undefined, getLabel("radio.ntOff"));
        var rbNTOn = ntPanel.add("radiobutton", undefined, getLabel("radio.ntOn"));
        rbNTKeep.value = true;

        /* 右カラム: UD 設定・P 設定 / Right: UD and P */
        var udProportionalColumn = variantColumns.add("group");
        setupGroup(udProportionalColumn, "column");

        var udVariantPanel = addPanel(udProportionalColumn, "panel.udSetting", 6);
        udVariantPanel.helpTip = getLabel("tooltip.udSetting");
        var rbUDKeep = udVariantPanel.add("radiobutton", undefined, getLabel("radio.keep"));
        var rbUDOff = udVariantPanel.add("radiobutton", undefined, getLabel("radio.udOff"));
        var rbUDOn = udVariantPanel.add("radiobutton", undefined, getLabel("radio.udOn"));
        rbUDKeep.value = true;

        var proportionalPanel = addPanel(udProportionalColumn, "panel.pSetting", 6);
        proportionalPanel.helpTip = getLabel("tooltip.pSetting");
        var rbPKeep = proportionalPanel.add("radiobutton", undefined, getLabel("radio.keep"));
        var rbPOff = proportionalPanel.add("radiobutton", undefined, getLabel("radio.pOff"));
        var rbPOn = proportionalPanel.add("radiobutton", undefined, getLabel("radio.pOn"));
        rbPKeep.value = true;

        /* 文字セット名 → ラジオボタンの対応 / Map charset name to its radio button */
        var charsetRadioByName = {
            Pro: rbCharsetPro,
            Pr5: rbCharsetPr5,
            Pr6: rbCharsetPr6,
            Std: rbCharsetStd
        };

        /**
         * ランキング順に、最初に選べる文字セット（＋N）をラジオに反映する
         * @param {string[]} rankList - 文字セットのランキング（CHARSET_RANK_NO_N / CHARSET_RANK_WITH_N）
         * @returns {void}
         */
        function applyRichestCharset(rankList) {
            for (var i = 0; i < rankList.length; i++) {
                var token = rankList[i];
                var withN = token.charAt(token.length - 1) === "N";
                var baseCharset = withN ? token.substring(0, token.length - 1) : token;
                var radio = charsetRadioByName[baseCharset];
                if (radio) {
                    radio.value = true;
                    if (withN) { rbNOn.value = true; } else { rbNOff.value = true; }
                    return;
                }
            }
        }

        /* Max/MaxN プリセットが押されたか / Whether a Max/MaxN preset is active
           AXIS は Std と ProN しか無いため、プリセット時は設定に関わらず ProN へ寄せる。
           押下後に文字セット/N を手動変更したら解除する。/ AXIS only has Std & ProN, so a preset forces ProN; cleared if charset/N changes manually. */
        var maxPresetActive = false;
        var presetResetRadios = [rbCharsetKeep, rbCharsetStd, rbCharsetPro, rbCharsetPr5, rbCharsetPr6, rbNKeep, rbNOff, rbNOn];
        for (var presetResetIndex = 0; presetResetIndex < presetResetRadios.length; presetResetIndex++) {
            presetResetRadios[presetResetIndex].onClick = function () { maxPresetActive = false; };
        }

        /* プリセット（Max=収録最多のNなし＋UD＋P、MaxN=収録最多のNあり込み＋UD＋P）/ Presets (Max / MaxN) */
        var presetRow = conversionPanel.add("group");
        setupRow(presetRow, "center"); // 左右中央 / horizontally centered
        presetRow.margins = [0, 5, 0, 0]; // 上に 5px の余白 / 5px top margin
        var presetMaxButton = presetRow.add("button", undefined, "Max");
        presetMaxButton.helpTip = getLabel("tooltip.presetMax");
        presetMaxButton.onClick = function () {
            applyRichestCharset(CHARSET_RANK_NO_N);
            rbUDOn.value = true;
            rbPOn.value = true;
            rbNTOn.value = true;
            maxPresetActive = true;
        };
        var presetMaxNButton = presetRow.add("button", undefined, "MaxN");
        presetMaxNButton.helpTip = getLabel("tooltip.presetMaxN");
        presetMaxNButton.onClick = function () {
            applyRichestCharset(CHARSET_RANK_WITH_N);
            rbUDOn.value = true;
            rbPOn.value = true;
            rbNTOn.value = true;
            maxPresetActive = true;
        };

        /* オプションパネル / Options panel */
        var optionsPanel = addPanel(mainDialog, "panel.options", 6);
        var cbIntegrateGakusan = optionsPanel.add("checkbox", undefined, getLabel("checkbox.integrateGakusan"));
        cbIntegrateGakusan.helpTip = getLabel("tooltip.integrateGakusan");
        cbIntegrateGakusan.value = INTEGRATE_GAKUSAN_TO_STANDARD;
        var cbIncludeStyles = optionsPanel.add("checkbox", undefined, getLabel("checkbox.includeStyles"));
        cbIncludeStyles.helpTip = getLabel("tooltip.includeStyles");
        cbIncludeStyles.value = true;
        var cbIncludeLocked = optionsPanel.add("checkbox", undefined, getLabel("checkbox.includeLocked"));
        cbIncludeLocked.helpTip = getLabel("tooltip.includeLocked");
        cbIncludeLocked.value = false;
        var cbIncludeHidden = optionsPanel.add("checkbox", undefined, getLabel("checkbox.includeHidden"));
        cbIncludeHidden.helpTip = getLabel("tooltip.includeHidden");
        cbIncludeHidden.value = false;

        /**
         * スタイルの変更は、対象がドキュメント全体のときだけ有効にする（範囲外で同じスタイルを使うテキストまで変わるため）
         * @returns {void}
         */
        function updateIncludeStylesEnabled() {
            cbIncludeStyles.enabled = rbTargetDocument.value;
        }
        rbTargetSelection.onClick = updateIncludeStylesEnabled;
        rbTargetDocument.onClick = updateIncludeStylesEnabled;
        rbTargetArtboard.onClick = updateIncludeStylesEnabled;
        /* 表示前は value を読み戻せないので、既定の対象（選択オブジェクト）に合わせて直接決める / value cannot be read back before show, so set it from the default scope (selection) */
        cbIncludeStyles.enabled = false;

        /* ボタン（Mac 規約: キャンセル → OK）/ Buttons (Mac order: Cancel then OK) */
        var buttonRow = addButtonRow(mainDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnCancel.onClick = function () { mainDialog.close(0); };
        btnOK.onClick = function () { mainDialog.close(1); };
        alignRightOnlyButtonRow(buttonRow);

        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        if (mainDialog.show() !== 1) {
            return;
        }

        // -----------------------------------------
        // 設定値の取得 / Read settings
        // -----------------------------------------

        var targetMode = "selection";
        if (rbTargetDocument.value) targetMode = "document";
        if (rbTargetArtboard.value) targetMode = "artboard";

        var charsetMode = "keep"; // keep / Pro / Pr5 / Pr6 / Std
        if (rbCharsetPro.value) charsetMode = "Pro";
        if (rbCharsetPr5.value) charsetMode = "Pr5";
        if (rbCharsetPr6.value) charsetMode = "Pr6";
        if (rbCharsetStd.value) charsetMode = "Std";

        var nMode = radioMode(rbNOn, rbNOff);   // keep / on / off
        var udMode = radioMode(rbUDOn, rbUDOff);
        var pMode = radioMode(rbPOn, rbPOff);
        var ntMode = radioMode(rbNTOn, rbNTOff);
        var maxPreset = maxPresetActive; // Max/MaxN プリセット中か（AXIS を ProN へ寄せる）/ whether a Max/MaxN preset is active (forces AXIS to ProN)

        var integrateGakusan = cbIntegrateGakusan.value;
        var includeStyles = (targetMode === "document") && cbIncludeStyles.value;
        var includeLocked = cbIncludeLocked.value;
        var includeHidden = cbIncludeHidden.value;
        var confirmBeforeRun = CONFIRM_BEFORE_RUN;

        /* 選択モードなのに未選択ならエラー / Error if selection mode but nothing selected */
        var selectedItems = getSelectedItems();
        if (targetMode === "selection" && selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        // -----------------------------------------
        // 対象テキストフレームの収集 / Collect target text frames
        // -----------------------------------------

        // カンバス上の位置（上→下、同じ高さは左→右）で並べ替え / Sort by canvas position (top→bottom, then left→right)
        var targetFrames = sortFramesByCanvasPosition(collectTargetTextFrames(targetMode));

        if (targetFrames.length === 0 && !includeStyles) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        // -----------------------------------------
        // 変更内容の事前計算 / Pre-compute changes
        // -----------------------------------------

        var fontExistsCache = {};
        var fontObjectCache = {};

        var processedFontNames = {};
        var directChanges = [];     // {oldName, newName}
        var weightSubChanges = [];  // {oldName, newName, oldWeight, newWeight}
        var missingChanges = [];    // {oldName, newName}

        /* 対象フレームをスキャンして変換候補を収集。区間は適用でも使い回す
           Scan frames to collect change candidates; the runs are reused when applying */
        var frameFontRuns = [];
        for (var fi = 0; fi < targetFrames.length; fi++) {
            var fontRuns = collectFontRuns(targetFrames[fi]);
            frameFontRuns.push(fontRuns);
            for (var ri = 0; ri < fontRuns.length; ri++) {
                if (fontRuns[ri].fontName) classifyFontChange(fontRuns[ri].fontName);
            }
        }

        /* 段落・文字スタイルもスキャン / Scan paragraph and character styles too */
        if (includeStyles) {
            scanStylesForChanges(documentRef.paragraphStyles);
            scanStylesForChanges(documentRef.characterStyles);
        }

        if (directChanges.length === 0 && weightSubChanges.length === 0 && missingChanges.length === 0) {
            alert(getLabel("alert.noChange"));
            return;
        }

        // -----------------------------------------
        // 確認ダイアログ / Confirmation dialog
        // -----------------------------------------

        var selectedOldNames = null; // null = すべて適用 / null = apply all
        if (confirmBeforeRun) {
            var confirmResult = showConfirmDialog();
            if (!confirmResult.ok) {
                return;
            }
            selectedOldNames = confirmResult.selected;
        }

        // -----------------------------------------
        // 変換マップの構築 / Build font-name map
        // -----------------------------------------

        var fontNameMap = {}; // oldName -> newName
        addSelectedChanges(directChanges, fontNameMap, selectedOldNames);
        addSelectedChanges(weightSubChanges, fontNameMap, selectedOldNames);

        // -----------------------------------------
        // 適用 / Apply
        // -----------------------------------------

        var changedCount = 0;
        for (var fj = 0; fj < targetFrames.length; fj++) {
            changedCount += applyChangesToFrame(targetFrames[fj], frameFontRuns[fj], fontNameMap);
        }
        // スタイルは適用のみ（テキストオブジェクト数には数えない）/ Styles are applied but not counted as text objects
        if (includeStyles) {
            applyChangesToStyles(documentRef.paragraphStyles, fontNameMap);
            applyChangesToStyles(documentRef.characterStyles, fontNameMap);
        }

        alert(getLabel("alert.done") + "\n\n" + labelValueText("alert.changedCount", changedCount));

        // =========================================
        // 対象収集 / Target collection
        // =========================================

        /**
         * 選択中のアイテムを配列で返す。文字ツールで文字を選択しているときは、そのストーリーのテキストフレームを返す
         * @returns {PageItem[]} 選択中のアイテム
         */
        function getSelectedItems() {
            var selection = documentRef.selection;
            var items = [];
            if (!selection) return items;
            /* 文字の選択中は TextRange になり [0] が無い（length は文字数）/ While characters are selected, selection is a TextRange with no [0] (length is the character count) */
            if (selection.typename === "TextRange") {
                var storyFrames = selection.story.textFrames;
                for (var i = 0; i < storyFrames.length; i++) {
                    items.push(storyFrames[i]);
                }
                return items;
            }
            for (var j = 0; j < selection.length; j++) {
                items.push(selection[j]);
            }
            return items;
        }

        /**
         * テキストフレームをカンバス上の位置順（上→下、同じ高さは左→右）に並べ替えた配列を返す。
         * 比較関数つきの sort() は遅く並びも狂うことがあるので、文字列キーを引数なしの sort() で並べる
         * @param {TextFrame[]} frames - 並べ替えるテキストフレーム
         * @returns {TextFrame[]} 並べ替えた配列
         */
        function sortFramesByCanvasPosition(frames) {
            var COORD_OFFSET = 100000; /* 座標を正の数にそろえる / shifts coordinates to positive numbers */
            var KEY_SEPARATOR = "\u0001";
            var sortKeys = [];
            for (var i = 0; i < frames.length; i++) {
                var frameBounds = frames[i].geometricBounds; // [left, top, right, bottom]
                /* top は 1pt 単位に丸め、大きい方（上）を先に。left は小さい方を先に / top rounded to 1pt, higher first; then smaller left first */
                sortKeys.push(
                    padNumber(Math.round(COORD_OFFSET - frameBounds[1]), 7) + KEY_SEPARATOR +
                    padNumber(Math.round((frameBounds[0] + COORD_OFFSET) * 100), 9) + KEY_SEPARATOR +
                    padNumber(i, 7)
                );
            }
            sortKeys.sort();
            var sortedFrames = [];
            for (var k = 0; k < sortKeys.length; k++) {
                sortedFrames.push(frames[parseInt(sortKeys[k].split(KEY_SEPARATOR)[2], 10)]);
            }
            return sortedFrames;
        }

        /**
         * 対象モードに応じてテキストフレームを集める
         * @param {string} mode - "selection" / "document" / "artboard"
         * @returns {TextFrame[]} 対象のテキストフレーム
         */
        function collectTargetTextFrames(mode) {
            if (mode === "document") {
                return collectDocumentFrames(null);
            }
            if (mode === "artboard") {
                var activeIndex = documentRef.artboards.getActiveArtboardIndex();
                return collectDocumentFrames(documentRef.artboards[activeIndex].artboardRect);
            }
            var selectedFrames = [];
            collectFramesFromSelection(selectedItems, selectedFrames);
            return selectedFrames;
        }

        /**
         * ドキュメント内の全テキストフレームを、ロック・非表示の設定に従って集める
         * @param {number[]|null} artboardRect - この範囲に重なるものだけを集める。null なら全体
         * @returns {TextFrame[]} 集めたテキストフレーム
         */
        function collectDocumentFrames(artboardRect) {
            var collected = [];
            for (var i = 0; i < documentRef.textFrames.length; i++) {
                var frame = documentRef.textFrames[i];
                if (!includeLocked && isLockedEffective(frame)) continue;
                if (!includeHidden && isHiddenEffective(frame)) continue;
                if (artboardRect && !boundsIntersect(frame.geometricBounds, artboardRect)) continue;
                collected.push(frame);
            }
            return collected;
        }

        /**
         * 選択アイテム（グループ内も再帰）からテキストフレームを集める
         * @param {PageItem[]|PageItems} items - 選択アイテム、またはグループの pageItems
         * @param {TextFrame[]} collected - 集めたフレームを足す配列
         * @returns {void}
         */
        function collectFramesFromSelection(items, collected) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (item.typename === "TextFrame") {
                    collected.push(item);
                } else if (item.typename === "GroupItem") {
                    collectFramesFromSelection(item.pageItems, collected);
                }
            }
        }

        /**
         * 2 つの矩形が交差するか
         * @param {number[]} boundsA - [left, top, right, bottom]
         * @param {number[]} boundsB - [left, top, right, bottom]
         * @returns {boolean} 交差すれば true
         */
        function boundsIntersect(boundsA, boundsB) {
            // bounds: [left, top, right, bottom]（Illustrator 座標は上が大きい）
            if (boundsA[2] < boundsB[0]) return false; // A.right  < B.left
            if (boundsA[0] > boundsB[2]) return false; // A.left   > B.right
            if (boundsA[3] > boundsB[1]) return false; // A.bottom > B.top
            if (boundsA[1] < boundsB[3]) return false; // A.top    < B.bottom
            return true;
        }

        // =========================================
        // ロック・非表示判定 / Lock & hidden checks
        // =========================================

        /**
         * 自身または祖先がロックされているか
         * @param {PageItem} item - 対象のアイテム
         * @returns {boolean} ロックされていれば true
         */
        function isLockedEffective(item) {
            if (item.locked) return true;
            var containers = collectAncestorContainers(item);
            for (var i = 0; i < containers.length; i++) {
                if (containers[i].locked) return true;
            }
            return false;
        }

        /**
         * 自身または祖先が非表示か
         * @param {PageItem} item - 対象のアイテム
         * @returns {boolean} 非表示なら true
         */
        function isHiddenEffective(item) {
            if (isContainerHidden(item)) return true;
            var containers = collectAncestorContainers(item);
            for (var i = 0; i < containers.length; i++) {
                if (isContainerHidden(containers[i])) return true;
            }
            return false;
        }

        /**
         * レイヤー・グループの祖先を内側から順に集める
         * @param {PageItem} item - 対象のアイテム
         * @returns {Array} 祖先のグループとレイヤー（内側が先頭）
         */
        function collectAncestorContainers(item) {
            var containers = [];
            var parent = item.parent;
            while (parent && (parent.typename === "Layer" || parent.typename === "GroupItem")) {
                containers.push(parent);
                parent = parent.parent;
            }
            return containers;
        }

        /**
         * 非表示かどうかを返す（レイヤーは visible の反転）
         * @param {Layer|PageItem} container - レイヤーまたはアイテム
         * @returns {boolean} 非表示なら true
         */
        function isContainerHidden(container) {
            return (container.typename === "Layer") ? !container.visible : container.hidden;
        }

        /**
         * ロック状態を設定する（失敗しても続ける）
         * @param {Layer|PageItem} container - レイヤーまたはアイテム
         * @param {boolean} locked - ロックするなら true
         * @returns {void}
         */
        function setContainerLocked(container, locked) {
            try { container.locked = locked; } catch (e) { }
        }

        /**
         * 非表示状態を設定する（レイヤーは visible の反転。失敗しても続ける）
         * @param {Layer|PageItem} container - レイヤーまたはアイテム
         * @param {boolean} hidden - 非表示にするなら true
         * @returns {void}
         */
        function setContainerHidden(container, hidden) {
            try {
                if (container.typename === "Layer") { container.visible = !hidden; }
                else { container.hidden = hidden; }
            } catch (e) { }
        }

        // =========================================
        // 変更内容のスキャン / Change scanning
        // =========================================

        /**
         * スタイルコレクションのフォントを変換候補に分類する
         * @param {ParagraphStyles|CharacterStyles} styles - スタイルのコレクション
         * @returns {void}
         */
        function scanStylesForChanges(styles) {
            for (var k = 0; k < styles.length; k++) {
                var fontName = styleFontName(styles[k]);
                if (fontName) classifyFontChange(fontName);
            }
        }

        /**
         * スタイルに設定されたフォント名を返す
         * @param {ParagraphStyle|CharacterStyle} style - 対象のスタイル
         * @returns {string|null} フォント名。無ければ null
         */
        function styleFontName(style) {
            try {
                var font = style.characterAttributes.textFont;
                return font ? font.name : null;
            } catch (e) {
                return null;
            }
        }

        /**
         * 旧フォント名を direct / weightSub / missing に分類する（同じ名前は 1 回だけ）
         * @param {string} oldName - 旧フォントの PostScript 名
         * @returns {void}
         */
        function classifyFontChange(oldName) {
            if (processedFontNames[oldName]) return;
            processedFontNames[oldName] = true;

            // 特殊シリーズ（A1明朝 など）を優先 / Special series first (e.g. A1 Mincho)
            var seriesTarget = resolveSpecialSeries(oldName);
            if (seriesTarget !== undefined) {
                if (seriesTarget === null) return; // 該当するが変更不要・対応なし
                if (fontExists(seriesTarget)) {
                    directChanges.push({ oldName: oldName, newName: seriesTarget });
                } else {
                    missingChanges.push({ oldName: oldName, newName: seriesTarget });
                }
                return;
            }

            // AXIS は文字セット体系が異なるため専用処理 / AXIS has its own charset scheme
            var axis = parseAxisName(oldName);
            if (axis) {
                classifyByDesired(oldName, buildAxisNameHead(axis), axis.weight);
                return;
            }

            var parsed = parseFontName(oldName);
            if (!parsed) return;

            classifyByDesired(oldName, buildConvertedNameHead(parsed), parsed.weight);
        }

        /**
         * 変換後の先頭（ウェイト除く）とウェイトから direct / weightSub / missing へ分類する
         * @param {string} oldName - 旧フォントの PostScript 名
         * @param {string|null} nameHead - 変換後の名前の先頭（末尾は「-」）。null なら変換しない
         * @param {string} weight - ウェイト
         * @returns {void}
         */
        function classifyByDesired(oldName, nameHead, weight) {
            if (nameHead === null) return; // 変換対象外（CID 非変換など）/ not convertible

            var desiredName = nameHead + weight;
            if (desiredName === oldName) return; // 変化なし / no change

            if (fontExists(desiredName)) {
                directChanges.push({ oldName: oldName, newName: desiredName });
                return;
            }

            var nearWeight = findNearestWeight(nameHead, weight);
            if (nearWeight) {
                weightSubChanges.push({
                    oldName: oldName,
                    newName: nameHead + nearWeight,
                    oldWeight: weight,
                    newWeight: nearWeight
                });
                return;
            }

            missingChanges.push({ oldName: oldName, newName: desiredName });
        }

        // =========================================
        // フォント名の分解・組み立て / Parse & build font names
        // =========================================

        /**
         * 特殊シリーズの対応を解決する
         * @param {string} oldName - 旧フォントの PostScript 名
         * @returns {string|null|undefined} 変換先の名前。該当するが対応なし・変更不要なら null、特殊シリーズでなければ undefined
         */
        function resolveSpecialSeries(oldName) {
            for (var seriesIndex = 0; seriesIndex < SPECIAL_SERIES.length; seriesIndex++) {
                var members = SPECIAL_SERIES[seriesIndex].members;

                // oldName がどのキー（文字セット）のメンバーか / Which charset key oldName belongs to
                var sourceKey = null;
                for (var key in members) {
                    if (members.hasOwnProperty(key) && members[key] === oldName) {
                        sourceKey = key;
                        break;
                    }
                }
                if (!sourceKey) continue;

                var targetKey = computeSeriesTargetKey(sourceKey);
                if (!targetKey || !members.hasOwnProperty(targetKey)) return null;

                var targetName = members[targetKey];
                if (!targetName || targetName === oldName) return null;
                return targetName;
            }
            return undefined;
        }

        /**
         * 設定から特殊シリーズの目標キー（文字セット[+N]）を求める
         * @param {string} sourceKey - 元のキー（"Std" / "StdN" など）
         * @returns {string} 目標のキー
         */
        function computeSeriesTargetKey(sourceKey) {
            var sourceHasN = sourceKey.charAt(sourceKey.length - 1) === "N";
            var sourceBaseCharset = sourceHasN ? sourceKey.substring(0, sourceKey.length - 1) : sourceKey;

            var targetBaseCharset = (charsetMode === "keep") ? sourceBaseCharset : charsetMode;

            var targetHasN = sourceHasN;
            if (nMode === "on") targetHasN = true;
            if (nMode === "off") targetHasN = false;

            return targetBaseCharset + (targetHasN ? "N" : "");
        }

        /**
         * AXIS の PostScript 名を分解する
         * @param {string} name - PostScript 名
         * @returns {{width: string, charset: string, hasN: boolean, weight: string}|null} 分解結果。AXIS でなければ null
         */
        function parseAxisName(name) {
            var matched = name.match(AXIS_FAMILY.regex);
            if (!matched) return null;
            return {
                width: matched[1] || "",    // Basic / Cond / Comp / ""
                charset: matched[2] || "",  // Std / Pro / Joyo / ""
                hasN: !!matched[3],
                weight: matched[4]
            };
        }

        /**
         * AXIS の変換後フォント名の先頭（ウェイト除く）を組み立てる。
         * 幅と Joyo は保持し、N と Std⇄Pro のみ反映。AXIS に無い Pr5/Pr6 は現状維持
         * @param {{width: string, charset: string, hasN: boolean, weight: string}} axis - parseAxisName() の結果
         * @returns {string} 名前の先頭（末尾は「-」）
         */
        function buildAxisNameHead(axis) {
            // Max/MaxN プリセット時は、設定に関わらず ProN（AXIS の収録最多）へ寄せる。AXIS は Std と ProN しか無いため。
            // Under a Max/MaxN preset, force ProN (the richest AXIS charset) regardless of settings; AXIS only has Std & ProN.
            if (maxPreset) {
                return AXIS_FAMILY.baseName + axis.width + "ProN" + "-";
            }

            // Joyo は別体系なので維持（文字セット・N の切り替え対象外）/ Joyo is a separate scheme; keep as-is
            if (axis.charset === "Joyo") {
                return AXIS_FAMILY.baseName + axis.width + "Joyo" + "-";
            }

            // 文字セット：AXIS は Std / Pro のみ。Pr5 / Pr6 や keep は現状維持 / AXIS only has Std / Pro
            var charset = axis.charset;
            if (charsetMode === "Std" || charsetMode === "Pro") charset = charsetMode;

            // N 切り替え（文字セットがある場合のみ意味を持つ）/ N toggle (only meaningful with a charset)
            var hasN = axis.hasN;
            if (nMode === "on") hasN = true;
            if (nMode === "off") hasN = false;

            var charsetCore = charset ? (charset + (hasN ? "N" : "")) : "";
            return AXIS_FAMILY.baseName + axis.width + charsetCore + "-";
        }

        /**
         * 旧フォント名を分解する（必要なら G-OTF 学参を A-OTF に読み替えて再試行）
         * @param {string} name - PostScript 名
         * @returns {Object|null} parseKnownFamily() の結果。既知のファミリーでなければ null
         */
        function parseFontName(name) {
            var parsed = parseKnownFamily(name);
            if (parsed) return parsed;

            if (integrateGakusan) {
                for (var prefixIndex = 0; prefixIndex < GAKUSAN_PREFIXES.length; prefixIndex++) {
                    var gakusanPrefix = GAKUSAN_PREFIXES[prefixIndex];
                    if (name.indexOf(gakusanPrefix) === 0) {
                        parsed = parseKnownFamily(name.substring(gakusanPrefix.length));
                        if (parsed) return parsed;
                    }
                }
            }
            return null;
        }

        /**
         * 既知ファミリーの正規表現で旧名を P / UD / NT / 文字セット / N / ウェイトに分解する
         * @param {string} name - PostScript 名
         * @returns {{family: Object, isProportional: boolean, isUD: boolean, isNT: boolean, charset: string, hasN: boolean, weight: string}|null} 分解結果。当てはまらなければ null
         */
        function parseKnownFamily(name) {
            for (var i = 0; i < FONT_FAMILIES.length; i++) {
                var family = FONT_FAMILIES[i];
                var matched = name.match(family.regex);
                if (!matched) continue;

                var parsed = {
                    family: family,
                    isProportional: false,
                    isUD: false,
                    isNT: false,
                    charset: "",
                    hasN: false,
                    weight: ""
                };

                // 最後のグループをウェイト、それ以外は値で判別 / Last group is the weight; others by value
                var lastIndex = matched.length - 1;
                parsed.weight = matched[lastIndex];

                for (var groupIndex = 1; groupIndex < lastIndex; groupIndex++) {
                    var groupValue = matched[groupIndex];
                    if (!groupValue) continue;
                    if (groupValue === "P") parsed.isProportional = true;
                    else if (groupValue === "UD") parsed.isUD = true;
                    else if (groupValue === "NT") parsed.isNT = true;
                    else if (groupValue === "N") parsed.hasN = true;
                    else if (groupValue === "Pro" || groupValue === "Pr5" || groupValue === "Pr6" || groupValue === "Std") parsed.charset = groupValue;
                }
                return parsed;
            }
            return null;
        }

        /**
         * 設定に基づき、変換後フォント名の先頭（ウェイト除く）を組み立てる
         * @param {Object} parsed - parseKnownFamily() の結果
         * @returns {string|null} 名前の先頭（末尾は「-」）。変換しないときは null
         */
        function buildConvertedNameHead(parsed) {
            var family = parsed.family;
            var supportsProportional = family.regex.source.indexOf("(P)") !== -1;
            var supportsUD = family.regex.source.indexOf("(UD)") !== -1;
            var isCID = (parsed.charset === "");

            // CID（文字セット表記なし）の扱い / Handle CID (no charset token)
            if (isCID && !CONVERT_CID_TO_OTF) {
                return null; // 変換しない
            }

            // 文字セット / Character set
            var charset;
            if (charsetMode !== "keep") {
                charset = charsetMode;
            } else if (isCID) {
                charset = CID_FALLBACK_CHARSET; // 現状維持かつ CID なら Pr6 を補完
            } else {
                charset = parsed.charset;
            }

            // P / Proportional（対応ファミリーのみ）/ Proportional (supported families only)
            var isProportional = parsed.isProportional;
            if (pMode === "on") isProportional = true;
            if (pMode === "off") isProportional = false;
            if (!supportsProportional) isProportional = false;

            // UD（対応ファミリーのみ）/ UD (supported families only)
            var isUD = parsed.isUD;
            if (udMode === "on") isUD = true;
            if (udMode === "off") isUD = false;
            if (!supportsUD) isUD = false;

            // NT（対応ファミリー＝新ゴ のみ）/ NT (only families that support it, i.e. ShinGo)
            var supportsNT = family.regex.source.indexOf("(NT)") !== -1;
            var isNT = parsed.isNT;
            if (ntMode === "on") isNT = true;
            if (ntMode === "off") isNT = false;
            if (!supportsNT) isNT = false;

            // N
            var hasN = parsed.hasN;
            if (nMode === "on") hasN = true;
            if (nMode === "off") hasN = false;

            var prefix = "";
            if (isProportional) prefix += "P";
            if (isUD) prefix += "UD";

            var charsetCore = "";
            if (charset) {
                charsetCore = charset + (hasN ? "N" : "");
            }

            // 基幹名（新ゴは NT 有無を付与）/ Base name (append NT for ShinGo when enabled)
            var baseName = family.baseName + (isNT ? "NT" : "");

            return prefix + baseName + charsetCore + "-";
        }

        /**
         * 同名のウェイトが無いときに、近いウェイトを探す（同じ距離なら重い方を先に）
         * @param {string} nameHead - 名前の先頭（末尾は「-」）
         * @param {string} weight - 元のウェイト
         * @returns {string|null} 見つかったウェイト。無ければ null
         */
        function findNearestWeight(nameHead, weight) {
            var weightIndex = indexOfArray(WEIGHT_ORDER, weight);
            if (weightIndex < 0) return null;

            for (var distance = 1; distance < WEIGHT_ORDER.length; distance++) {
                var heavierIndex = weightIndex + distance;
                var lighterIndex = weightIndex - distance;
                if (heavierIndex < WEIGHT_ORDER.length && fontExists(nameHead + WEIGHT_ORDER[heavierIndex])) {
                    return WEIGHT_ORDER[heavierIndex];
                }
                if (lighterIndex >= 0 && fontExists(nameHead + WEIGHT_ORDER[lighterIndex])) {
                    return WEIGHT_ORDER[lighterIndex];
                }
            }
            return null;
        }

        // =========================================
        // フォント存在判定 / Font availability
        // =========================================

        /**
         * 指定名のフォントがインストール済みか（キャッシュ付き）。
         * 環境にないフォントも app.textFonts に仮エントリ（style が空で family が PostScript 名）として並ぶので除く
         * @param {string} name - PostScript 名
         * @returns {boolean} インストール済みなら true
         */
        function fontExists(name) {
            if (fontExistsCache.hasOwnProperty(name)) return fontExistsCache[name];
            var exists = false;
            try {
                var fontObject = app.textFonts.getByName(name);
                exists = !(fontObject.style === "" && fontObject.family === fontObject.name);
                if (exists) fontObjectCache[name] = fontObject;
            } catch (e) {
                exists = false;
            }
            fontExistsCache[name] = exists;
            return exists;
        }

        /**
         * フォントオブジェクトを返す（キャッシュ付き）。見つからなければ例外
         * @param {string} name - PostScript 名
         * @returns {TextFont} フォント
         */
        function getFontObject(name) {
            if (fontObjectCache.hasOwnProperty(name)) return fontObjectCache[name];
            var fontObject = app.textFonts.getByName(name);
            fontObjectCache[name] = fontObject;
            return fontObject;
        }

        // =========================================
        // 確認ダイアログ / Confirmation dialog
        // =========================================

        /**
         * 変更内容を一覧にして、項目ごとに ON/OFF させる。
         * 行が CONFIRM_VISIBLE_ROWS を超えるときは、その行数だけを描いてスクロールバーで表示位置をずらす
         * @returns {{ok: boolean, selected: Object|null}} ok は［実行］で閉じたか、selected はチェックした旧フォント名の集合
         */
        function showConfirmDialog() {
            var confirmDialog = new Window("dialog", getLabel("confirm.title"));
            setupWindow(confirmDialog);
            // 左右マージンを +10 / Add 10 to left & right margins
            confirmDialog.margins = [WINDOW_MARGINS + 10, WINDOW_MARGINS, WINDOW_MARGINS + 10, WINDOW_MARGINS];

            var confirmEntries = buildConfirmEntries();

            // 「変更前」列の幅を全項目でそろえ、→ の位置を一定にする / Fix the "before" column width so arrows line up
            var beforeWidth = computeBeforeColumnWidth(confirmDialog, [directChanges, weightSubChanges]);
            var afterWidth = 0;
            var plainWidth = 0;
            for (var i = 0; i < confirmEntries.length; i++) {
                var entry = confirmEntries[i];
                if (entry.kind === "change") {
                    afterWidth = Math.max(afterWidth, measureTextWidth(confirmDialog, entry.afterText));
                } else {
                    plainWidth = Math.max(plainWidth, measureTextWidth(confirmDialog, entry.text));
                }
            }
            var rowWidth = Math.max(beforeWidth + CONFIRM_COLUMN_GAP + afterWidth, plainWidth);

            /* ScriptUI のグループはスクロールしないので、決まった数の行を使い回して表示位置をずらす
               ScriptUI groups do not scroll, so a fixed set of rows is reused and shifted */
            var visibleRowCount = Math.min(confirmEntries.length, CONFIRM_VISIBLE_ROWS);
            var listRow = confirmDialog.add("group");
            setupRow(listRow, ["fill", "top"], CONFIRM_COLUMN_GAP * 2);
            listRow.alignChildren = ["left", "fill"];
            var rowsHost = listRow.add("group");
            rowsHost.orientation = "column";
            rowsHost.alignChildren = ["left", "top"];
            rowsHost.spacing = CONFIRM_ROW_SPACING;

            var rowSlots = [];
            for (var s = 0; s < visibleRowCount; s++) {
                rowSlots.push(addConfirmRowSlot(rowsHost, beforeWidth, afterWidth, rowWidth));
            }

            var scrollOffset = 0;

            /**
             * 表示位置から行に項目を割り当てる
             * @returns {void}
             */
            function showConfirmEntries() {
                for (var slotIndex = 0; slotIndex < rowSlots.length; slotIndex++) {
                    var slot = rowSlots[slotIndex];
                    var slotEntry = confirmEntries[scrollOffset + slotIndex];
                    var isChange = (slotEntry.kind === "change");
                    slot.entry = slotEntry;
                    slot.changeRow.visible = isChange;
                    slot.plainText.visible = !isChange;
                    if (isChange) {
                        slot.checkbox.text = slotEntry.beforeText;
                        slot.checkbox.value = slotEntry.checked;
                        slot.afterText.text = slotEntry.afterText;
                    } else {
                        slot.plainText.text = slotEntry.text;
                    }
                }
            }

            if (confirmEntries.length > visibleRowCount) {
                var listScrollbar = listRow.add("scrollbar", undefined, 0, 0, confirmEntries.length - visibleRowCount);
                listScrollbar.preferredSize.width = CONFIRM_SCROLLBAR_WIDTH;
                listScrollbar.alignment = ["right", "fill"];
                listScrollbar.stepdelta = 1;
                listScrollbar.jumpdelta = visibleRowCount;
                listScrollbar.onChanging = listScrollbar.onChange = function () {
                    var newOffset = Math.round(listScrollbar.value);
                    if (newOffset === scrollOffset) return;
                    scrollOffset = newOffset;
                    showConfirmEntries();
                };
            }
            showConfirmEntries();

            confirmDialog.add("statictext", undefined, getLabel("confirm.prompt"));

            var confirmButtonRow = addButtonRow(confirmDialog);
            var btnConfirmCancel = confirmButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            var btnConfirmRun = confirmButtonRow.rightGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });
            btnConfirmCancel.onClick = function () { confirmDialog.close(0); };
            btnConfirmRun.onClick = function () { confirmDialog.close(1); };
            alignRightOnlyButtonRow(confirmButtonRow);

            prepareDialogWindow(confirmDialog, SCRIPT_NAME + "_confirm");
            if (confirmDialog.show() !== 1) {
                return { ok: false, selected: null };
            }

            // チェックが ON の項目だけ採用 / Keep only checked items
            var selected = {};
            for (var k = 0; k < confirmEntries.length; k++) {
                if (confirmEntries[k].kind === "change" && confirmEntries[k].checked) {
                    selected[confirmEntries[k].oldName] = true;
                }
            }
            return { ok: true, selected: selected };
        }

        /**
         * 確認ダイアログに並べる項目（見出し・変更・未インストールのフォント名）を作る
         * @returns {Object[]} 項目。kind は "change"（チェックボックス付き）または "text"
         */
        function buildConfirmEntries() {
            var entries = [];
            addChangeEntries(entries, directChanges, getLabel("confirm.willChange"));
            addChangeEntries(entries, weightSubChanges, getLabel("confirm.nearWeight"));
            if (missingChanges.length > 0) {
                entries.push({ kind: "text", text: getLabel("confirm.notInstalled") });
                var missingNames = uniqueArray(extractNewFontNames(missingChanges));
                for (var k = 0; k < missingNames.length; k++) {
                    entries.push({ kind: "text", text: "　" + missingNames[k] });
                }
            }
            return entries;
        }

        /**
         * 見出しと変更の項目（既定 ON）を足す
         * @param {Object[]} entries - 足す先の項目
         * @param {Object[]} changes - 変更の一覧
         * @param {string} headingText - 見出し
         * @returns {void}
         */
        function addChangeEntries(entries, changes, headingText) {
            if (changes.length === 0) return;
            entries.push({ kind: "text", text: headingText });
            for (var i = 0; i < changes.length; i++) {
                var change = changes[i];
                var afterLabel = "→ " + toDisplayFontName(change.newName);
                if (change.oldWeight) {
                    afterLabel += "（" + change.oldWeight + " → " + change.newWeight + "）";
                }
                entries.push({
                    kind: "change",
                    oldName: change.oldName,
                    beforeText: toDisplayFontName(change.oldName),
                    afterText: afterLabel,
                    checked: true
                });
            }
        }

        /**
         * 確認ダイアログの 1 行を作る。変更の行（チェックボックス＋変更後）と文字だけの行を重ねて持ち、表示を切り替える
         * @param {Group} parent - 行を足す先
         * @param {number} beforeWidth - 「変更前」列（チェックボックス）の幅
         * @param {number} afterWidth - 「変更後」列の幅
         * @param {number} rowWidth - 文字だけの行の幅
         * @returns {{changeRow: Group, checkbox: Checkbox, afterText: StaticText, plainText: StaticText, entry: Object|null}} 行の部品
         */
        function addConfirmRowSlot(parent, beforeWidth, afterWidth, rowWidth) {
            var slotGroup = parent.add("group");
            slotGroup.orientation = "stack";
            slotGroup.alignChildren = ["left", "center"];

            /* 高さが決まるよう仮の文字を入れておく（表示前に差し替える）/ Placeholder text so the height is set (replaced before show) */
            var changeRow = slotGroup.add("group");
            changeRow.orientation = "row";
            changeRow.alignChildren = ["left", "center"];
            changeRow.spacing = CONFIRM_COLUMN_GAP;
            var checkbox = changeRow.add("checkbox", undefined, "M");
            checkbox.preferredSize.width = beforeWidth;
            var afterText = changeRow.add("statictext", undefined, "M");
            afterText.preferredSize.width = afterWidth;
            var plainText = slotGroup.add("statictext", undefined, "M");
            plainText.preferredSize.width = rowWidth;

            var slot = { changeRow: changeRow, checkbox: checkbox, afterText: afterText, plainText: plainText, entry: null };
            checkbox.onClick = function () {
                if (slot.entry) slot.entry.checked = checkbox.value;
            };
            return slot;
        }

        /**
         * PostScript 名を和文表示名（family + style）に変換する
         * @param {string} psName - PostScript 名
         * @returns {string} 表示名。見つからなければ PostScript 名
         */
        function toDisplayFontName(psName) {
            try {
                var fontObject = getFontObject(psName);
                return fontObject.style ? (fontObject.family + " " + fontObject.style) : fontObject.family;
            } catch (e) {
                return psName;
            }
        }

        /**
         * 文字列の表示幅を測る（切れないよう少し余裕を足す）
         * @param {Window} dialog - 測るダイアログ
         * @param {string} text - 文字列
         * @returns {number} 幅（px）
         */
        function measureTextWidth(dialog, text) {
            var graphics = dialog.graphics;
            return graphics.measureString(text, graphics.font).width + 10;
        }

        /**
         * 「変更前」列に必要な最大幅（チェックボックスのボックス幅＋余白を加味）を返す
         * @param {Window} dialog - 測るダイアログ
         * @param {Object[][]} changeLists - 変更の一覧の配列
         * @returns {number} 幅（px）
         */
        function computeBeforeColumnWidth(dialog, changeLists) {
            var graphics = dialog.graphics;
            var maxTextWidth = 0;
            for (var a = 0; a < changeLists.length; a++) {
                var list = changeLists[a];
                for (var i = 0; i < list.length; i++) {
                    var displayName = toDisplayFontName(list[i].oldName);
                    var textWidth = graphics.measureString(displayName, graphics.font).width;
                    if (textWidth > maxTextWidth) maxTextWidth = textWidth;
                }
            }
            return maxTextWidth + 30; // チェックボックスのボックス＋すき間 / checkbox box + gap
        }

        /**
         * 選択された（または全件の）変更を変換マップに追加する
         * @param {Object[]} changes - 変更の一覧
         * @param {Object} nameMap - 旧名 → 新名のマップ
         * @param {Object|null} selectedOldNames - 採用する旧名の集合。null なら全件
         * @returns {void}
         */
        function addSelectedChanges(changes, nameMap, selectedOldNames) {
            for (var i = 0; i < changes.length; i++) {
                var change = changes[i];
                if (selectedOldNames === null || selectedOldNames[change.oldName]) {
                    nameMap[change.oldName] = change.newName;
                }
            }
        }

        /**
         * 変更の一覧から新フォント名だけを取り出す
         * @param {Object[]} changes - 変更の一覧
         * @returns {string[]} 新フォント名
         */
        function extractNewFontNames(changes) {
            var newNames = [];
            for (var i = 0; i < changes.length; i++) {
                newNames.push(changes[i].newName);
            }
            return newNames;
        }

        // =========================================
        // 適用 / Apply
        // =========================================

        /**
         * テキストフレームにフォント変更を適用する（ロック・非表示は一時的に解除）
         * @param {TextFrame} frame - 対象のフレーム
         * @param {Object[]} fontRuns - collectFontRuns() で集めた区間
         * @param {Object} nameMap - 旧名 → 新名のマップ
         * @returns {number} 変更したら 1、しなければ 0
         */
        function applyChangesToFrame(frame, fontRuns, nameMap) {
            /* 変える区間が無ければ、ロック・表示の切り替えもしない / Skip the lock/visibility toggling when nothing changes */
            var hasMappedRun = false;
            for (var r = 0; r < fontRuns.length; r++) {
                if (fontRuns[r].fontName && nameMap[fontRuns[r].fontName]) {
                    hasMappedRun = true;
                    break;
                }
            }
            if (!hasMappedRun) return 0;

            var containers = [frame].concat(collectAncestorContainers(frame)); /* 内側が先頭 / innermost first */
            var savedLocked = [];
            var savedHidden = [];

            /* 外側（レイヤー）から解除する。ロックされたレイヤーの中のアイテムは解除できないため
               Clear from the outside in, since items inside a locked layer cannot be unlocked */
            for (var i = containers.length - 1; i >= 0; i--) {
                savedLocked[i] = containers[i].locked;
                if (includeLocked && savedLocked[i]) setContainerLocked(containers[i], false);
                savedHidden[i] = isContainerHidden(containers[i]);
                if (includeHidden && savedHidden[i]) setContainerHidden(containers[i], false);
            }

            var frameChanged = false;
            var storyOffset = frame.textRange.start;

            // textRange 単位でまとめて適用 / Apply per run (textRange) for speed
            for (var runIndex = 0; runIndex < fontRuns.length; runIndex++) {
                var fontRun = fontRuns[runIndex];
                if (!fontRun.fontName) continue;
                var newName = nameMap[fontRun.fontName];
                if (!newName) continue;
                var runRange = frame.textRange;
                runRange.start = storyOffset + fontRun.start;
                runRange.end = storyOffset + fontRun.end;
                try {
                    runRange.characterAttributes.textFont = getFontObject(newName);
                    frameChanged = true;
                } catch (e) { }
            }

            /* 内側から戻し、外側のレイヤーを最後にロックする / Restore from the inside out so the outer layer is locked last */
            for (var restoreIndex = 0; restoreIndex < containers.length; restoreIndex++) {
                if (includeHidden && savedHidden[restoreIndex]) setContainerHidden(containers[restoreIndex], true);
                if (includeLocked && savedLocked[restoreIndex]) setContainerLocked(containers[restoreIndex], true);
            }

            // テキストオブジェクト単位でカウント / Count per text object
            return frameChanged ? 1 : 0;
        }

        /**
         * スタイルコレクションにフォント変更を適用する
         * @param {ParagraphStyles|CharacterStyles} styles - スタイルのコレクション
         * @param {Object} nameMap - 旧名 → 新名のマップ
         * @returns {number} 変更したスタイルの数
         */
        function applyChangesToStyles(styles, nameMap) {
            var changed = 0;
            for (var k = 0; k < styles.length; k++) {
                var fontName = styleFontName(styles[k]);
                if (!fontName) continue;
                var newName = nameMap[fontName];
                if (!newName) continue;
                try {
                    styles[k].characterAttributes.textFont = getFontObject(newName);
                    changed++;
                } catch (e) { }
            }
            return changed;
        }

        /**
         * 同じフォントが続く区間（run）を集める
         * @param {TextFrame} frame - 対象のフレーム
         * @returns {{fontName: string, start: number, end: number}[]} 区間。start と end はフレーム内の文字番号（end は含まない）
         */
        function collectFontRuns(frame) {
            var fontRuns = [];
            var frameChars = frame.characters;
            var charCount = frameChars.length;
            var i = 0;
            while (i < charCount) {
                var runFont = frameChars[i].characterAttributes.textFont;
                var runName = runFont ? runFont.name : "";
                var j = i + 1;
                while (j < charCount) {
                    var nextFont = frameChars[j].characterAttributes.textFont;
                    if ((nextFont ? nextFont.name : "") !== runName) break;
                    j++;
                }
                fontRuns.push({ fontName: runName, start: i, end: j });
                i = j;
            }
            return fontRuns;
        }

        // =========================================
        // ユーティリティ / Utilities
        // =========================================

        /**
         * ダイアログにパネルを追加する（共通設定）
         * @param {Window|Group|Panel} parent - 追加先
         * @param {string} labelPath - パネル名のラベルキー
         * @param {number} [spacing] - 要素間隔
         * @returns {Panel} 追加したパネル
         */
        function addPanel(parent, labelPath, spacing) {
            var panel = parent.add("panel", undefined, getLabel(labelPath));
            setupPanel(panel, spacing);
            return panel;
        }

        /**
         * ラジオ 3 択（あり / なし / 現状維持）の状態を返す
         * @param {RadioButton} onRadio - 「あり」のラジオ
         * @param {RadioButton} offRadio - 「なし」のラジオ
         * @returns {string} "on" / "off" / "keep"
         */
        function radioMode(onRadio, offRadio) {
            if (onRadio.value) return "on";
            if (offRadio.value) return "off";
            return "keep";
        }

        /**
         * 配列中の位置を返す（ES3 互換）
         * @param {Array} arr - 配列
         * @param {*} value - 探す値
         * @returns {number} 位置。無ければ -1
         */
        function indexOfArray(arr, value) {
            for (var i = 0; i < arr.length; i++) {
                if (arr[i] === value) return i;
            }
            return -1;
        }

        /**
         * 重複を除いた配列を返す
         * @param {string[]} arr - 配列
         * @returns {string[]} 重複を除いた配列
         */
        function uniqueArray(arr) {
            var seen = {};
            var unique = [];
            for (var i = 0; i < arr.length; i++) {
                if (!seen[arr[i]]) {
                    seen[arr[i]] = true;
                    unique.push(arr[i]);
                }
            }
            return unique;
        }

        /**
         * 0 埋めした数字の文字列を返す
         * @param {number} value - 0 以上の整数
         * @param {number} digits - 桁数
         * @returns {string} 0 埋めした文字列
         */
        function padNumber(value, digits) {
            var text = String(value);
            while (text.length < digits) text = "0" + text;
            return text;
        }

    })();

})();
