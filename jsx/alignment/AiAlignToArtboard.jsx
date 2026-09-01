#target illustrator
#targetengine "AiAlignToArtboard"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクトを、アートボードを対象に整列する常駐パレットです。
3×3のボタンで8方向へ寄せ、押すたびにガイド・アートボードの端・裁ち落としへと寄せ先が進みます。マージンや分割のガイドも引けます。

詳細は README を参照してください。

### Overview

A persistent palette that aligns the selected objects to the artboard.
A 3x3 grid of buttons moves the selection in eight directions, stepping the destination outwards on each
press: the guide, the artboard edge, then the bleed. It also draws margin and division guides.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAlignToArtboard";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAlignToArtboard.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAlignToArtboard.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n42952a7adcb6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 常駐エンジンに残すパレット参照（GC回避と多重起動防止を兼ねる）
       var の初期化は再実行のたびに走るため、既存の参照を消さないよう $.global から引き継ぐ
       The palette reference lives in the persistent engine; it is carried over from $.global so a
       re-run does not wipe it before closeExistingPalette() can close the old window */
    var paletteWindow = $.global.__aiAlignToArtboardWindow || null;

    (function() {

        // =========================================
        // 整列コマンド / Align commands
        // =========================================
        /* 「オブジェクト > 整列」のメニューコマンド名と、マージンぶん内側へ動かす向き
           整列コマンドはアートボードの辺にぴったり寄せるので、そこから offset 方向へマージンぶん動かす
           Y は上が正のため、上揃えは -1（下へ）・下揃えは +1（上へ）になる
           axis は整列する軸で、整列先の判定に使う仮移動の向きを決める
           mode はその軸のどこに寄せるか（start＝左・上／center＝中央／end＝右・下）で、
           字形の境界での補正（btApplyGlyphCorrection）が目標位置を計算するのに使う
           justification は水平方向の整列に合わせる行揃え（垂直方向は行揃えを変えないので null）
           Menu command names under Object > Align, with the direction to move by the margin;
           Y grows upward, so top align moves -1 (down) and bottom align +1 (up).
           axis is the axis being aligned, which sets the direction of the probe used to check the align target.
           justification is the paragraph justification to match; vertical aligns leave it alone (null) */
        var ALIGN_COMMANDS = {
            horizontalLeft:   { command: "Horizontal Align Left",   axis: "x", mode: "start",  offsetX:  1, offsetY:  0, justification: "LEFT" },
            horizontalCenter: { command: "Horizontal Align Center",  axis: "x", mode: "center", offsetX:  0, offsetY:  0, justification: "CENTER" },
            horizontalRight:  { command: "Horizontal Align Right",   axis: "x", mode: "end",    offsetX: -1, offsetY:  0, justification: "RIGHT" },
            verticalTop:      { command: "Vertical Align Top",       axis: "y", mode: "start",  offsetX:  0, offsetY: -1, justification: null },
            verticalCenter:   { command: "Vertical Align Center",    axis: "y", mode: "center", offsetX:  0, offsetY:  0, justification: null },
            verticalBottom:   { command: "Vertical Align Bottom",    axis: "y", mode: "end",    offsetX:  0, offsetY:  1, justification: null }
        };

        // =========================================
        // 定規の単位 / Ruler units
        // =========================================
        /* rulerType の単位コード→ラベルと pt 換算係数 / Unit code to label and points-per-unit */
        /* defaultMargin はマージン欄、defaultExtension は伸張欄の初期値
           どちらも換算値ではなく、その単位で扱いやすい丸めた数にする
           defaultMargin and defaultExtension are the initial margin and extension: round numbers that
           read well in that unit, not conversions */
        var UNIT_INFO = {
            "0":  { label: "in",    points: 72.0,                      defaultMargin: 0.25,  defaultExtension: 0.5 },
            "1":  { label: "mm",    points: 72.0 / 25.4,               defaultMargin: 5,     defaultExtension: 10 },
            "2":  { label: "pt",    points: 1.0,                       defaultMargin: 20,    defaultExtension: 20 },
            "3":  { label: "pica",  points: 12.0,                      defaultMargin: 1.5,   defaultExtension: 2 },
            "4":  { label: "cm",    points: 72.0 / 2.54,               defaultMargin: 0.5,   defaultExtension: 1 },
            "5":  { label: "Q/H",   points: (72.0 / 25.4) * 0.25,      defaultMargin: 20,    defaultExtension: 40 },
            "6":  { label: "px",    points: 1.0,                       defaultMargin: 20,    defaultExtension: 20 },
            "7":  { label: "ft/in", points: 864.0,                     defaultMargin: 0.02,  defaultExtension: 0.04 },
            "8":  { label: "m",     points: (72.0 / 25.4) * 1000.0,    defaultMargin: 0.005, defaultExtension: 0.01 },
            "9":  { label: "yd",    points: 2592.0,                    defaultMargin: 0.006, defaultExtension: 0.012 },
            "10": { label: "ft",    points: 864.0,                     defaultMargin: 0.02,  defaultExtension: 0.04 }
        };
        /* 単位が取れないときの既定 / Fallback when the ruler unit cannot be read */
        var FALLBACK_UNIT_INFO = { label: "pt", points: 1.0, defaultMargin: 20, defaultExtension: 20 };

        // =========================================
        // ユーザー設定 / User settings
        // =========================================
        /* チェックボックスの初期状態。［プレビュー境界］［字形の境界に整列］はパレットを開くときに
           環境設定の現在値で上書きするので、この値は読めなかったときの控えになる
           （整列のあいだだけワーカーが環境設定へ書き込み、終わったら元の設定に戻す）
           Initial checkbox states. Preview Bounds and Align to Glyph Bounds are overwritten with the current
           preferences when the palette opens, so these are only the fallback; the worker applies them for the
           duration of an align and restores the previous preferences afterwards */
        var DEFAULT_PREVIEW_BOUNDS       = false; /* プレビュー境界 / preview bounds */
        var DEFAULT_GLYPH_BOUNDS         = true;  /* 字形の境界に整列 / align to glyph bounds */
        var DEFAULT_CHANGE_JUSTIFICATION = true;  /* 行揃えを変更 / change justification */
        var DEFAULT_LINK_MARGINS         = true;  /* マージンの4値を連動させる / keep the four margins in sync */
        var DEFAULT_ALIGN_TO_BLEED       = false; /* 裁ち落としに整列 / align to the bleed */
        var DEFAULT_ALIGN_PER_ARTBOARD   = false; /* アートボードごとに整列 / align per artboard */
        var DEFAULT_BLEED                = 3;     /* 裁ち落とし欄の初期値（mm）/ initial bleed, in millimetres */
        /* 裁ち落としは印刷の値なので、定規の単位に追従させず mm で扱う
           The bleed is a print value, so it stays in millimetres instead of following the ruler */
        var BLEED_UNIT_LABEL  = "mm";
        var BLEED_UNIT_POINTS = 72.0 / 25.4;
        var DEFAULT_SHOW_GUIDE           = false; /* ガイドを追加 / add the margin guide */
        var DEFAULT_KEEP_GUIDE           = false; /* ガイドを保持（閉じても残す）/ keep the guide when the palette closes */
        /* 分割ガイドの初期状態。行・列はマージンの内側をいくつに分けるかで、1 のときはガイドを引かない
           行間・列間・伸張は定規の単位で扱う（伸張はアートボードの外へ伸ばす距離）
           Division guides: rows and columns split the area inside the margin, so 1 draws no guide;
           the gutters and the extension are in ruler units, the extension reaching outside the artboard */
        var DEFAULT_DIVIDE_MODE          = "none";
        var DEFAULT_ARTBOARD_EDGE        = false; /* アートボードのエッジにガイドを引く / draw guides on the artboard edges */
        /* 伸張の初期値は単位ごとに決まるので、ここには持たない（UNIT_INFO の defaultExtension を使う）
           The extension's default comes from the unit, so it is not listed here */
        var DEFAULT_DIVIDE_VALUES        = { rows: 1, columns: 2, rowGutter: 0, columnGutter: 0 };
        /* ［十字］の行数・列数（縦横とも2等分＝中央に十字のガイドが1本ずつ）
           "Cross" halves both directions, leaving one guide on each axis through the centre */
        var CROSS_DIVIDE_COUNTS          = { rows: 2, columns: 2 };
        /* 行数・列数の上限（大きな値でガイドを何千本も作らないための歯止め）/ Cap on the counts, so a large number cannot flood the document */
        var DIVIDE_COUNT_MAX             = 100;

        /* マージンのガイドを作るレイヤー名（他のガイド系スクリプトと共通）
           Layer that receives the margin guide, shared with the other guide scripts */
        var GUIDE_LAYER_NAME = "_guide";
        /* このスクリプトが作るガイドの名前。張り替えるときの目印にする
           The name given to the guide, used to find and replace it */
        var GUIDE_NAME = "AiAlignToArtboard-margin";
        /* 分割ガイドの名前。マージンのガイドと分けて、片方だけ張り替えられるようにする
           The name given to the division guides, kept apart from the margin guide */
        var DIVIDE_GUIDE_NAME = "AiAlignToArtboard-divide";

        /* 常駐エンジン（$.global）に控える値のキー
           ガイドの設定はパレットを開き直しても引き継ぐ（［ガイドを保持］で残したガイドを、
           次に閉じたときに消してしまわないため）
           Keys kept on $.global: the guide settings survive a close-and-reopen, so a guide left by
           "Keep Guides" is not deleted the next time the palette closes */
        var SETTINGS_KEY     = "__aiAlignToArtboardSettings";
        /* メインエンジンへ送り込んだワーカー定義の刻印を控えるキー / Key holding the stamp of the loaded worker source */
        var WORKER_STAMP_KEY = "__aiAlignToArtboardWorkerStamp";

        /* メインエンジンからの応答を待つ秒数 / seconds to wait for the main engine */
        var WORKER_TIMEOUT = 10;
        /* 整列先が［アートボード］かを判定するための仮移動量（pt）
           整列してもオブジェクトが動かなかったときだけ、このぶん内側へずらして整列し直し、戻ってくるかを見る
           Probe distance (pt): used only when an align moved nothing, to tell "already aligned" from a wrong target */
        var ALIGN_PROBE_PT = 4;
        /* ガイドが水平・垂直かを判定する許容値（pt）。これを超える幅・高さがあれば長方形とみなす
           Tolerance (pt) for calling a guide horizontal or vertical; anything thicker counts as a rectangle */
        var GUIDE_ORIENTATION_TOLERANCE = 0.01;
        /* 移動先がここまで近ければ「すでにその位置にいる」とみなす許容値（pt）
           手で吸着させたオブジェクトの辺とガイドは 1e-12 ほどずれることがあり、そのままでは行き先に選ばれて動かなくなる
           Tolerance (pt) for "already there"; a hand-snapped edge and its guide can differ by ~1e-12,
           which would otherwise be picked as the destination and move nothing */
        var MOVE_MIN_DELTA_PT = 0.001;
        /* 選択を取り直す最短間隔（mouseover は何度も発生するため間引く）/ Throttle for the mouseover refresh */
        var SELECTION_POLL_INTERVAL_MS = 400;

        // =========================================
        // レイアウト / Layout
        // =========================================
        var WINDOW_MARGINS  = 15;   /* パレット外周の余白 / window margin */
        var WINDOW_SPACING  = 12;   /* パレット内の要素間隔 / window spacing */
        var ICON_SIZE       = 30;   /* 整列アイコン1個の大きさ（px）/ size of each align icon (px) */
        var ICON_GAP        = 6;    /* 整列アイコンどうしの間隔 / gap between align icons */
        var RADIO_GAP       = 14;   /* ラジオボタンどうしの間隔 / gap between the radio buttons */
        var CROSS_GAP       = 2;    /* 移動ボタン（十字）どうしの間隔 / gap between the move buttons */
        var CENTER_ROW_TOP  = 8;    /* 十字と中央揃えボタンの間隔 / gap between the cross and the centre-align buttons */
        var PANEL_MARGINS   = [12, 16, 12, 10]; /* オプションパネルの余白 [左,上,右,下]（上はタイトルのぶん広め）/ options panel margins */
        var COLUMN_SPACING  = 18;   /* 方向ボタンの十字と計測オプションのパネルの間隔 / gap between the move-button cross and the options panel */
        var FIELD_CHARS     = 3;    /* マージン入力欄の文字数 / width of the margin field */
        var LABEL_FIELD_SPACING = 4; /* 入力欄と単位ラベルの間隔（既定は広すぎる）/ gap between the field and its unit label */
        /* マージン欄を3×3に並べるときの1セルの幅（日英で文字数が違うので分ける）
           Width of one cell in the 3x3 margin grid; the labels differ in length by language */
        var MARGIN_CELL_WIDTH = { ja: 70, en: 84 };
        /* 分割ガイドの項目名の幅。右そろえにして数値欄の頭をそろえる（日英で文字数が違うので分ける）
           Width of the division labels; right-aligned so the fields line up, and the labels differ by language */
        var DIVIDE_LABEL_WIDTH        = { ja: 40, en: 72 };
        var DIVIDE_GUTTER_LABEL_WIDTH = { ja: 40, en: 52 };
        /* 単位ラベルの幅（定規の単位が変わっても欄の位置が動かないよう固定）
           Fixed width for the unit labels, so the fields stay put when the ruler unit changes */
        var UNIT_LABEL_WIDTH          = 28;
        var STATUS_WIDTH    = 260;  /* 状況表示の幅（中身でパレット幅が変わらないよう固定）/ fixed width of the status line */
        var OPTION_SPACING  = 4;    /* オプションのチェックボックスどうしの間隔 / gap between the option checkboxes */

        // =========================================
        // アイコンの寸法 / Icon metrics
        // =========================================
        /* すべて ICON_SIZE に対する比率で持ち、アイコンサイズを変えても形が崩れないようにする
           All ratios of ICON_SIZE so the glyphs keep their shape when the icon size changes */
        var ICON_BAR_THICKNESS  = 0.25;  /* オブジェクトを表すバーの太さ / thickness of the bars standing for objects */
        var ICON_BAR_GAP        = 0.07;  /* バー2本の間隔 / gap between the two bars */
        var ICON_BAR_LONG       = 0.55;  /* 長いほうのバーの長さ / length of the longer bar */
        var ICON_BAR_SHORT      = 0.35;  /* 短いほうのバーの長さ / length of the shorter bar */
        var ICON_RULE_INSET     = 0.12;  /* 基準線の端の余白 / inset at both ends of the reference rule */
        var ICON_RULE_OFFSET    = 0.17;  /* 端に置く基準線の位置 / position of the rule when it sits at an edge */
        var ICON_RULE_CLEARANCE = 0.05;  /* 端の基準線とバーのすき間（中央の基準線はバーの下を通す）/ gap between an edge rule and the bars (a center rule runs behind them) */
        /* アイコンに置くオブジェクトの大きさ。どのアイコンも同じ横長の長方形にする
           The object block: every icon uses the same landscape rectangle */
        var ICON_BLOCK_WIDTH    = 0.38;  /* オブジェクトの幅 / width of the object block */
        var ICON_BLOCK_HEIGHT   = 0.30;  /* オブジェクトの高さ / height of the object block */

        // =========================================
        // ローカライズ / Localization
        // =========================================

        /**
         * 現在のUI言語を判定する
         * @returns {string} "ja" または "en"
         */
        function getCurrentLang() {
            var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
            return (localeText.indexOf("ja") === 0) ? "ja" : "en";
        }
        var uiLang = getCurrentLang();

        /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
        var LABELS = {
            dialog: {
                title: { ja: "ガイドやアートボードのエッジに整列", en: "Align to Guides or Artboard Edges" }
            },
            panel: {
                guide:       { ja: "マージンガイド", en: "Margin Guide" },
                divide:      { ja: "分割ガイド", en: "Division Guides" },
                options:     { ja: "計測オプション", en: "Measurement Options" },
                destination: { ja: "整列オプション", en: "Align Options" }
            },
            fieldLabel: {
                top:    { ja: "上", en: "Top" },
                bottom: { ja: "下", en: "Bottom" },
                left:   { ja: "左", en: "Left" },
                right:  { ja: "右", en: "Right" },
                rows:         { ja: "行", en: "Rows" },
                columns:      { ja: "列", en: "Columns" },
                rowGutter:    { ja: "行間", en: "Gutter" },
                columnGutter: { ja: "列間", en: "Gutter" },
                extension:    { ja: "伸張", en: "Extension" }
            },
            direction: {
                up:        { ja: "上", en: "Up" },
                left:      { ja: "左", en: "Left" },
                right:     { ja: "右", en: "Right" },
                down:      { ja: "下", en: "Down" },
                upLeft:    { ja: "左上", en: "Top left" },
                upRight:   { ja: "右上", en: "Top right" },
                downLeft:  { ja: "左下", en: "Bottom left" },
                downRight: { ja: "右下", en: "Bottom right" }
            },
            tooltip: {
                alignCenterH:   { ja: "水平方向中央に整列", en: "Horizontal Align Center" },
                alignCenterV:   { ja: "垂直方向中央に整列", en: "Vertical Align Center" },
                alignCenterAll: { ja: "水平・垂直方向中央に整列", en: "Align Center on Both Axes" },
                margin: { ja: "アートボードの端から空ける距離", en: "Distance to keep from the artboard edge" },
                linkMargins: {
                    ja: "上下左右のマージンを同じ値にする（どれかを変えると残りもそろえる）",
                    en: "Keep the four margins equal; changing one updates the rest"
                },
                alignToBleed: {
                    ja: "アートボードの端の次の寄せ先として、その外側の裁ち落としの位置を使う",
                    en: "Add the bleed outside the artboard as the stop after the artboard edge"
                },
                perArtboard: {
                    ja: "選択したオブジェクトを、それぞれが乗っているアートボードに整列（OFFのときは1つのアートボードにまとめて整列）",
                    en: "Align each object to the artboard it sits on (off: everything goes to a single artboard)"
                },
                moveToEdge: {
                    ja: "その方向のガイド → アートボードの端 → 裁ち落とし の順に寄せる（↑↓←→キーでも実行）",
                    en: "Step to the guide in that direction, then the artboard edge, then the bleed (the arrow keys do the same)"
                },
                keepGuide: {
                    ja: "パレットを閉じてもガイドを残す（OFFのときは閉じるときに削除）",
                    en: "Leave the guide in place when the palette closes (deleted on close when off)"
                },
                showGuide: {
                    ja: "マージンの位置に長方形のガイドを作る（「_guide」レイヤー、アクティブなアートボードに1つ）",
                    en: "Draw a rectangle guide at the margin (on the \"_guide\" layer, one on the active artboard)"
                },
                divideNone: {
                    ja: "分割のガイドを引かない",
                    en: "Draw no division guides"
                },
                divideCross: {
                    ja: "マージンの内側の中央に、縦横1本ずつの十字のガイドを引く",
                    en: "Draw a cross through the centre of the area inside the margin"
                },
                divideCustom: {
                    ja: "横方向・縦方向それぞれの分割数で、マージンの内側を等分するガイドを引く",
                    en: "Split the area inside the margin into the number of columns and rows below"
                },
                divideRows: {
                    ja: "マージンの内側を上下に何段に分けるか（1のときはガイドを引かない）",
                    en: "How many rows to split the area inside the margin into (1 draws no guide)"
                },
                divideColumns: {
                    ja: "マージンの内側を左右に何列に分けるか（1のときはガイドを引かない）",
                    en: "How many columns to split the area inside the margin into (1 draws no guide)"
                },
                divideRowGutter: {
                    ja: "行と行のあいだに空ける距離（0より大きいと、段の上下2本のガイドを引く）",
                    en: "Space between two rows; above zero, each split gets a guide on both sides"
                },
                divideColumnGutter: {
                    ja: "列と列のあいだに空ける距離（0より大きいと、列の左右2本のガイドを引く）",
                    en: "Space between two columns; above zero, each split gets a guide on both sides"
                },
                divideExtension: {
                    ja: "分割のガイドをアートボードの外へ伸ばす距離",
                    en: "Distance the division guides reach outside the artboard"
                },
                artboardEdge: {
                    ja: "アートボードの上下左右4辺にガイドを引く（伸張のぶんだけ外へ伸ばす）",
                    en: "Draw guides on the four edges of the artboard, reaching outside by the extension"
                },
                optionGlyphBounds: { ja: "Option＋クリックで字形の境界に整列", en: "Option-click to align to glyph bounds" },
                optionMarginCenter: {
                    ja: "Option＋クリックでマージンの内側の中央に整列（通常はアートボードの中央）",
                    en: "Option-click to centre inside the margin instead of the artboard"
                },
                optionNoMargin: {
                    ja: "Option＋クリックでマージンなし・字形の境界に整列",
                    en: "Option-click: ignore the margin, align to glyph bounds"
                },
                previewBounds: {
                    ja: "整列でプレビュー境界（線幅・効果を含む）を使用",
                    en: "Use preview bounds (incl. stroke & effects) when aligning"
                },
                glyphBounds: {
                    ja: "ポイント文字・エリア内文字を字形の境界で整列",
                    en: "Align point & area type to glyph bounds"
                },
                changeJustification: {
                    ja: "水平方向の整列に合わせて、1行だけのテキスト1つの行揃えも変える",
                    en: "Match the justification of a lone single-line text object to the horizontal alignment"
                }
            },
            radio: {
                divideNone:   { ja: "なし", en: "None" },
                divideCross:  { ja: "十字", en: "Cross" },
                divideCustom: { ja: "カスタム", en: "Custom" }
            },
            checkbox: {
                showGuide:     { ja: "ガイドを追加", en: "Add Guides" },
                keepGuide:     { ja: "ガイドを保持", en: "Keep Guides" },
                artboardEdge:  { ja: "アートボードのエッジ", en: "Artboard Edges" },
                linkMargins:   { ja: "連動", en: "Link" },
                previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
                glyphBounds:   { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
                alignToBleed:  { ja: "裁ち落としに整列", en: "Align to Bleed" },
                perArtboard:   { ja: "アートボードごとに整列", en: "Align per Artboard" },
                changeJustification: { ja: "行揃えを変更", en: "Change Justification" }
            },
            status: {
                done:           { ja: "整列しました。", en: "Aligned." },
                moved:          { ja: "移動しました。", en: "Moved." },
                noBounds:       { ja: "境界を取得できません。", en: "Could not measure the selection." },
                doneJustified:  { ja: "整列し、行揃えを{0}に変更しました。", en: "Aligned; justification set to {0}." },
                movedJustified: { ja: "移動し、行揃えを{0}に変更しました。", en: "Moved; justification set to {0}." },
                noDocument:     { ja: "ドキュメントが開かれていません。", en: "No document is open." },
                noSelection:    { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
                multipleLayers: { ja: "レイヤーをまたぐ選択は整列できません。", en: "Cannot align a selection spanning layers." },
                /* 状況表示は STATUS_WIDTH で切り詰められるため、全角20字ほどに収める
                   （切れても helpTip で全文を読める）
                   Keep it within the fixed status width; the full text is still available as a helpTip */
                alignTarget: {
                    ja: "整列先を［アートボード］にしてください。",
                    en: "Set Align To: Artboard."
                },
                noResponse:     { ja: "Illustrator から応答がありません。", en: "No response from Illustrator." },
                genericError:   { ja: "エラー：", en: "Error: " },
                /* doneJustified の {0} に入れる行揃えの名前 / Names substituted into doneJustified */
                justification: {
                    LEFT:   { ja: "左揃え",   en: "Left" },
                    CENTER: { ja: "中央揃え", en: "Center" },
                    RIGHT:  { ja: "右揃え",   en: "Right" }
                }
            }
        };

        /**
         * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('tooltip','alignLeft')）
         * @param {...string} keys - LABELS を辿るキー列
         * @returns {string} 該当するラベル（見つからない場合は空文字）
         */
        function getLabel() {
            var labelNode = LABELS;
            for (var i = 0; i < arguments.length; i++) {
                if (labelNode == null) break;
                labelNode = labelNode[arguments[i]];
            }
            return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
        }

        /**
         * 項目名の末尾にコロンを付けて取得する（日本語は全角、英語は半角）
         * @param {string} category - LABELS のカテゴリ
         * @param {string} key - ラベルのキー
         * @returns {string} コロン付きのラベル
         */
        function labelText(category, key) {
            return getLabel(category, key) + (uiLang === "ja" ? "：" : ":");
        }

        // =========================================
        // 配色 / Colors
        // =========================================
        /* アイコンの配色（initIconColors() で UI 明暗から設定）/ Icon colors (set from the light/dark UI in initIconColors()) */
        var iconColor, iconBaseBg, iconHoverBg, iconBorderColor;

        /**
         * UI 明度（0..1）を取得する
         * @returns {number} 0〜1 にクランプした明度（取得失敗時は 0＝暗い側）
         */
        function getUIBrightness() {
            try {
                var brightness = app.preferences.getRealPreference("uiBrightness");
                if (brightness < 0) { brightness = 0; }
                if (brightness > 1) { brightness = 1; }
                return brightness;
            } catch (e) {
                return 0;
            }
        }

        /**
         * グレーの RGBA を作る
         * @param {number} value - 明度（0..1 にクランプ）
         * @returns {number[]} [r, g, b, a] の配列
         */
        function grayColor(value) {
            if (value < 0) { value = 0; }
            if (value > 1) { value = 1; }
            return [value, value, value, 1];
        }

        /**
         * UI の明暗に合わせてアイコン色とマウスオーバー時の背景色を決める
         * @returns {void}
         */
        function initIconColors() {
            var uiBrightness = getUIBrightness();
            var lightUI = uiBrightness > 0.5;
            iconColor = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
            /* 通常時の背景はパレットの地色に近いグレー。graphics.backgroundColor は iconbutton などで取得できず
               fillPath() が例外を投げ、再描画のたびにボタンが消えるため、必ず明示色で塗る
               Always paint an explicit gray; graphics.backgroundColor is unavailable on some controls and makes
               fillPath() throw, which blanks the button on every redraw */
            iconBaseBg  = lightUI ? grayColor(uiBrightness)        : [0.28, 0.28, 0.28, 1];
            /* マウスオーバー時の背景（ライトは少し暗く、ダークは少し明るく）/ Hover background (slightly darker in light, lighter in dark) */
            iconHoverBg = lightUI ? grayColor(uiBrightness - 0.10) : [0.38, 0.38, 0.38, 1];
            /* マウスオーバー時の枠線（ライトは薄いグレー、ダークは背景より明るいグレー）
               Hover border: light gray in light UI, gray brighter than the background in dark UI */
            iconBorderColor = lightUI ? [0.65, 0.65, 0.65, 1] : [0.45, 0.45, 0.45, 1];
        }

        // =========================================
        // 描画ヘルパー / Drawing helpers
        // =========================================

        /**
         * 塗りつぶした矩形を描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} x - 左端
         * @param {number} y - 上端
         * @param {number} width - 幅
         * @param {number} height - 高さ
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function fillRect(graphics, x, y, width, height, color) {
            graphics.newPath();
            graphics.moveTo(x, y);
            graphics.lineTo(x + width, y);
            graphics.lineTo(x + width, y + height);
            graphics.lineTo(x, y + height);
            graphics.closePath();
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
        }

        /**
         * 整列の基準となる位置（左端・中央・右端）を求める
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} alignMode - "start" / "center" / "end"
         * @returns {number} 基準線の座標
         */
        function getRulePosition(size, alignMode) {
            if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
            if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
            return Math.round(size / 2) + 0.5;
        }

        /**
         * バー2本の並び方向の開始座標を求める（バー・間隔・バーの合計を中央に置く）
         * @param {number} size - アイコンの一辺の長さ
         * @returns {number} 1本目のバーの開始座標
         */
        function getBarStackOrigin(size) {
            var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
            return Math.round((size - stackLength) / 2);
        }

        /**
         * 基準線に対するバーの開始座標を求める（端の基準線からはすき間を空け、中央の基準線はバーの下を通す）
         * @param {number} rulePosition - 基準線の座標
         * @param {number} barLength - バーの長さ
         * @param {string} alignMode - "start" / "center" / "end"
         * @param {number} clearance - 端の基準線とバーのすき間
         * @returns {number} バーの開始座標
         */
        function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
            /* 基準線は rulePosition を中心とした太さ1なので、その両端から数えて左右・上下を対称にする
               The rule is 1 unit thick around rulePosition, so measure from its edges to keep both ends symmetrical */
            if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
            if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
            return Math.round(rulePosition - barLength / 2);
        }

        /**
         * バー2本の長さを並び順に返す（水平は上が短く、90°回した垂直は左が長い）
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} iconType - "horizontal" または "vertical"
         * @returns {number[]} 並び順のバーの長さ
         */
        function getBarLengths(size, iconType) {
            var longBar = Math.round(size * ICON_BAR_LONG);
            var shortBar = Math.round(size * ICON_BAR_SHORT);
            return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
        }

        /**
         * 向きに合わせて矩形を描く（垂直方向のアイコンは水平方向の座標を縦横入れ替えて描く）
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {boolean} isVertical - 垂直方向のアイコンなら true
         * @param {number} alignPos - 整列する軸の座標（水平は x、垂直は y）
         * @param {number} stackPos - バーが並ぶ軸の座標（水平は y、垂直は x）
         * @param {number} alignLen - 整列する軸方向の長さ
         * @param {number} stackLen - バーが並ぶ軸方向の長さ
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
            if (isVertical) {
                fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
            } else {
                fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
            }
        }

        /**
         * 整列アイコン（バー2本＋基準線）を描く
         * 水平方向と垂直方向は縦横が入れ替わるだけなので、同じ手順で描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} alignMode - "start"＝左・上 / "center"＝中央 / "end"＝右・下
         * @param {number[]} color - RGBA の配列
         * @param {string} iconType - "horizontal" または "vertical"
         * @returns {void}
         */
        function drawAlignIcon(graphics, size, alignMode, color, iconType) {
            var isVertical = (iconType === "vertical");
            var rulePosition = getRulePosition(size, alignMode);
            var barThickness = Math.round(size * ICON_BAR_THICKNESS);
            var barGap = Math.round(size * ICON_BAR_GAP);
            var barStackOrigin = getBarStackOrigin(size);
            var barLengths = getBarLengths(size, iconType);
            var clearance = Math.round(size * ICON_RULE_CLEARANCE);
            var ruleInset = Math.round(size * ICON_RULE_INSET);

            /* 基準線を先に描き、バーを上に重ねる（中央の基準線がバーの下を通って見える）
               Draw the rule first and the bars on top, so a center rule runs behind them */
            fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);

            for (var i = 0; i < barLengths.length; i++) {
                fillOrientedRect(graphics, isVertical,
                    getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                    barStackOrigin + i * (barThickness + barGap),
                    barLengths[i], barThickness, color);
            }
        }

        /**
         * アイコンの中心線の座標を返す（ケイ線もオブジェクトもここを中心に置く）
         * 太さ1のケイ線がピクセル境界に乗るよう、中心は .5 の位置に取る
         * @param {number} size - アイコンの一辺の長さ
         * @returns {number} 中心の座標
         */
        function getIconCenter(size) {
            return Math.round(size / 2) + 0.5;
        }

        /**
         * 辺の長さを奇数に丸める
         * 中心が .5 の位置にあるので、奇数にすると両側が同じ幅で割り振られ、中心がぴったり合う
         * @param {number} value - 丸める前の長さ
         * @returns {number} 奇数の長さ
         */
        function roundToOddLength(value) {
            var rounded = Math.round(value);
            if (rounded % 2 !== 0) { return rounded; }
            return (rounded > 1) ? rounded - 1 : 1;
        }

        /**
         * アイコンに置くオブジェクトの大きさを求める（どのアイコンも同じ長方形）
         * @param {number} size - アイコンの一辺の長さ
         * @returns {number[]} [幅, 高さ]
         */
        function getBlockSize(size) {
            return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
        }

        /**
         * 端のケイ線（寄せ先を表す線）を1本描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} side - "top" / "bottom" / "left" / "right"
         * @param {number[]} color - RGBA の配列
         * @returns {number} 引いたケイ線の座標
         */
        function drawEdgeRule(graphics, size, side, color) {
            var ruleInset = Math.round(size * ICON_RULE_INSET);
            var ruleLength = size - ruleInset * 2;
            var isTopOrLeft = (side === "top" || side === "left");
            var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
            if (side === "top" || side === "bottom") {
                fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
            } else {
                fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
            }
            return rulePosition;
        }

        /**
         * アイコンの中央を貫くケイ線を1本描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} ruleDirection - "vertical"＝縦のケイ線 / "horizontal"＝横のケイ線
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawCenterRule(graphics, size, ruleDirection, color) {
            var ruleInset = Math.round(size * ICON_RULE_INSET);
            var ruleLength = size - ruleInset * 2;
            var center = getIconCenter(size);
            if (ruleDirection === "vertical") {
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
            } else {
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
            }
        }

        /**
         * オブジェクトをアイコンの中央に置く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {number[]} blockSize - [幅, 高さ]
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function fillCenteredBlock(graphics, size, blockSize, color) {
            var center = getIconCenter(size);
            fillRect(graphics,
                center - blockSize[0] / 2, center - blockSize[1] / 2,
                blockSize[0], blockSize[1], color);
        }

        /**
         * 水平・垂直の中央に整列するアイコン（十字のケイ線＋中央に置いたオブジェクト）を描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawCenterBothIcon(graphics, size, color) {
            /* ケイ線を先に描き、オブジェクトを上に重ねる（ケイ線がオブジェクトの下を通って見える）
               Draw the rules first and the block on top, so they run behind it */
            drawCenterRule(graphics, size, "vertical", color);
            drawCenterRule(graphics, size, "horizontal", color);
            fillCenteredBlock(graphics, size, getBlockSize(size), color);
        }

        /**
         * 片方の軸だけ中央に整列するアイコン（中央を貫くケイ線1本＋オブジェクト）を描く
         * 水平方向中央は縦のケイ線、垂直方向中央は横のケイ線で見分ける
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} iconType - "horizontal" または "vertical"
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawCenterOneIcon(graphics, size, iconType, color) {
            drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
            fillCenteredBlock(graphics, size, getBlockSize(size), color);
        }

        /**
         * 方向ボタンのアイコン（寄せ先のケイ線と、そこへ寄せたオブジェクト）を描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {string} directionKey - MOVE_ICON_RULES のキー
         * @param {number} size - ボタンの一辺
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawMoveIcon(graphics, directionKey, size, color) {
            var rules, blockSize, clearance, center, blockX, blockY, side, rulePosition, i;
            rules = MOVE_ICON_RULES[directionKey];
            if (!rules) { return; }
            blockSize = getBlockSize(size);
            clearance = Math.round(size * ICON_RULE_CLEARANCE);
            center = getIconCenter(size);
            /* ケイ線のない軸では中央に置く / Centred on any axis without a rule */
            blockX = center - blockSize[0] / 2;
            blockY = center - blockSize[1] / 2;

            for (i = 0; i < rules.length; i++) {
                side = rules[i];
                rulePosition = drawEdgeRule(graphics, size, side, color);
                /* ケイ線は太さ1なので、その端から数えてすき間を取る / The rule is 1 unit thick, so measure from its edge */
                if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
            }

            fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
        }

        /**
         * ホバー状態に応じた背景色を返す
         * @param {Button} control - 対象のコントロール
         * @returns {number[]} 背景色の RGBA
         */
        function hoverBackground(control) {
            return (control.isHover === true) ? iconHoverBg : iconBaseBg;
        }

        /**
         * ボタンの下地（背景と、マウスオーバー中だけの枠線）を描く
         * @param {Button} button - 対象のボタン
         * @returns {void}
         */
        function drawButtonBase(button) {
            var graphics = button.graphics;
            var width = button.size[0];
            var height = button.size[1];

            try {
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, hoverBackground(button)));
            } catch (e) {
                try { graphics.drawOSControl(); } catch (osControlError) {}
            }

            /* マウスオーバー中だけ、ボタン領域（正方形）のエッジにグレーの枠を描く
               0.5 ずらすと1pxの線がピクセル境界に乗ってくっきり出る
               A gray border on the square button's edge while hovered; the 0.5 offset keeps the 1px line crisp */
            if (button.isHover === true) {
                try {
                    graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconBorderColor, 1));
                } catch (borderError) {}
            }
        }

        /**
         * 整列アイコンボタンを描画する
         * @param {Button} button - 対象のボタン（iconType と alignMode を持つ）
         * @returns {void}
         */
        function drawAlignButton(button) {
            drawButtonBase(button);
            if (button.iconType === "center") {
                drawCenterBothIcon(button.graphics, button.size[0], iconColor);
            } else if (button.alignMode === "center") {
                drawCenterOneIcon(button.graphics, button.size[0], button.iconType, iconColor);
            } else {
                drawAlignIcon(button.graphics, button.size[0], button.alignMode, iconColor, button.iconType);
            }
        }

        /**
         * 移動ボタン（寄せ先のケイ線とオブジェクト）を描画する
         * @param {Button} button - 対象のボタン（directionKey を持つ）
         * @returns {void}
         */
        function drawMoveButton(button) {
            drawButtonBase(button);
            drawMoveIcon(button.graphics, button.directionKey, button.size[0], iconColor);
        }

        // =========================================
        // UI 部品 / UI helpers
        // =========================================

        /**
         * コントロールのサイズを固定する（最小・推奨・最大を同じ値でそろえる）
         * @param {Object} control - 対象のコントロール
         * @param {number} width - 幅
         * @param {number} height - 高さ
         * @returns {void}
         */
        function fixControlSize(control, width, height) {
            control.minimumSize = [width, height];
            control.preferredSize = [width, height];
            control.maximumSize = [width, height];
        }

        /**
         * 行グループの共通設定を適用する（alignment と alignChildren を必ず対で指定する）
         * @param {Group} targetGroup - 対象のグループ
         * @param {string} [alignment] - グループ自身の横方向の配置（既定は "left"）
         * @param {number} [spacing] - グループ内の要素間隔
         * @returns {void}
         */
        function setupRow(targetGroup, alignment, spacing) {
            targetGroup.orientation = "row";
            targetGroup.alignment = [alignment || "left", "center"];
            targetGroup.alignChildren = ["left", "center"];
            targetGroup.spacing = (typeof spacing === "number") ? spacing : ICON_GAP;
        }

        /**
         * コントロールを再描画する（notify は環境により例外を投げ得るので保護）
         * @param {Object} control - 対象のコントロール
         * @returns {void}
         */
        function redrawControl(control) {
            try { control.notify("onDraw"); } catch (e) {}
        }

        /**
         * ↑↓キー操作後の次の値を求める（下限は0）
         * @param {number} currentValue - 現在の値
         * @param {number} direction - 1＝上 / -1＝下
         * @param {object} keyboard - ScriptUI.environment.keyboardState
         * @returns {number} 次の値
         */
        function computeArrowValue(currentValue, direction, keyboard) {
            if (keyboard.shiftKey) {
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
                if (direction > 0) { return Math.ceil((currentValue + 1) / 10) * 10; }
                return Math.max(0, Math.floor((currentValue - 1) / 10) * 10);
            }
            if (keyboard.altKey) {
                /* Optionキー押下時は0.1単位で増減し、小数第1位までに丸め / Step by 0.1 when Option is held */
                return Math.max(0, Math.round((currentValue + direction * 0.1) * 10) / 10);
            }
            /* 通常は1単位で増減し、整数に丸め / Step by 1 and round to an integer */
            return Math.max(0, Math.round(currentValue + direction));
        }

        /**
         * 数値入力欄を ↑↓ キーで増減できるようにする（Shift＝±10・Option＝±0.1）
         * @param {EditText} editText - 対象の入力欄
         * @returns {void}
         */
        function changeValueByArrowKey(editText) {
            editText.addEventListener("keydown", function(event) {
                /* 入れ子三項は括弧で右結合を明示（ExtendScriptは左結合に誤評価）
                   Parenthesize: ExtendScript mis-parses nested ternaries */
                var direction = (event.keyName === "Up") ? 1 : ((event.keyName === "Down") ? -1 : 0);
                /* ↑↓以外では欄に書き戻さない（書き戻すと入力途中の小数点が消える）
                   Never write back on other keys; doing so wipes a half-typed decimal point */
                if (direction === 0) return;

                var currentValue = Number(editText.text);
                if (isNaN(currentValue)) return;

                editText.text = computeArrowValue(currentValue, direction, ScriptUI.environment.keyboardState);
                /* ↑↓での増減は確定値なので、確定時と同じ処理を走らせる
                   プログラムからの変更では onChange が発火しないため明示的に呼ぶ
                   An arrow step is a committed value; programmatic changes do not fire onChange, so call it */
                if (typeof editText.onChange === "function") { editText.onChange(); }
                event.preventDefault();
            });
        }

        /**
         * Option（Alt）キーが押されているか判定する
         * @returns {boolean} 押されていれば true（取得できない環境では false）
         */
        function isAltPressed() {
            try {
                return ScriptUI.environment.keyboardState.altKey === true;
            } catch (e) {
                return false;
            }
        }

        /**
         * マウスオーバーの状態を button.isHover に反映して再描画する
         * @param {Button} button - 対象のボタン
         * @returns {void}
         */
        function attachHover(button) {
            try {
                button.addEventListener("mouseover", function() { button.isHover = true; redrawControl(button); });
                button.addEventListener("mouseout", function() { button.isHover = false; redrawControl(button); });
            } catch (e) {}
        }

        // =========================================
        // パレット構築 / Palette builder
        // =========================================

        /* チェックボックスの参照（アイコンのクリック時に読む）/ The checkboxes, read when an icon is clicked */
        var previewBoundsCheckbox = null;
        var glyphBoundsCheckbox = null;
        var changeJustificationCheckbox = null;
        /* マージン欄・単位ラベル・ガイド表示チェックボックスの参照
           The margin field, its unit label, and the guide checkbox */
        /* 辺の名前をキーにしたマージンの入力欄 / Margin fields, keyed by side */
        var marginFields = {};
        var marginPanel = null;
        var linkMarginsCheckbox = null;
        /* マージン欄がまだ既定値のままか。定規の単位が分かった時点で、その単位の既定値に入れ直す
           Whether the margin fields still hold defaults; they are refilled once the ruler unit is known */
        var marginsAreDefault = false;
        /* 伸張欄がまだ既定値のままか。マージン欄と同じく、定規の単位が分かった時点でその単位の既定値に入れ直す
           Whether the extension field still holds a default; like the margins, it is refilled once the ruler unit is known */
        var extensionIsDefault = false;
        var bleedField = null;
        var alignToBleedCheckbox = null;
        var alignPerArtboardCheckbox = null;
        var showGuideCheckbox = null;
        var keepGuideCheckbox = null;
        /* 分割ガイドの切り替えラジオと分割数の入力欄（どちらも軸・モードの名前をキーにする）
           The division mode radios and count fields, keyed by mode and axis */
        var divideRadios = {};
        var divideFields = {};
        /* 分割ガイドの間隔・伸張に添える単位ラベル（定規の単位が変わったら書き換える）
           The unit labels beside the gutter and extension fields, rewritten when the ruler unit changes */
        var divideUnitLabels = [];
        var artboardEdgeCheckbox = null;
        /* パレットを閉じてもガイドを残すか。閉じる処理でコントロールを触らずに済むよう値を控えておく
           Whether to keep the guide on close; mirrored so teardown never has to read a control */
        var keepGuideOnClose = DEFAULT_KEEP_GUIDE;
        /* パレットを組み立て終えたか。組み立てのあいだは控えを書かない
           （まだ作っていないコントロールを既定値として読み、引き継いだ設定を消してしまうため）
           Whether the palette is fully built; settings are not stored until it is, or controls that do not
           exist yet would be read as defaults and wipe the carried-over settings */
        var isPaletteReady = false;
        /* 現在の定規単位（パレットへフォーカスが来るたびに取り直す）/ The current ruler unit, re-read on focus */
        var currentUnitInfo = FALLBACK_UNIT_INFO;
        /* 状況表示の参照 / The status line */
        var statusText = null;

        /* 水平・垂直方向中央に整列するボタンの定義。移動ボタンの十字の中央に置く
           The centre-on-both-axes button; it sits in the middle of the move-button cross */
        var CENTER_ALIGN_BUTTON = { iconType: "center", alignMode: "center", tooltip: "alignCenterAll", alignKeys: ["horizontalCenter", "verticalCenter"] };

        /* 片方の軸だけ中央に整列するボタン2つ。十字の下に横に並べる
           The two single-axis centre buttons, in a row below the cross */
        var CENTER_ALIGN_BUTTONS = [
            { iconType: "horizontal", alignMode: "center", tooltip: "alignCenterH", alignKeys: ["horizontalCenter"] },
            { iconType: "vertical",   alignMode: "center", tooltip: "alignCenterV", alignKeys: ["verticalCenter"] }
        ];

        /* ショートカットから引くためのボタン定義の一覧
           端に寄せる整列は十字の外周ボタンが受け持つので、整列ボタンは中央揃えの3つだけになる
           Every align button definition in one list; the cross handles the edge alignments, so only the
           three centred buttons remain */
        var ALL_ALIGN_BUTTONS = CENTER_ALIGN_BUTTONS.concat([CENTER_ALIGN_BUTTON]);

        /* 移動ボタンの並び（中央は移動ではなく、水平・垂直方向中央に整列するボタン）
           Layout of the move buttons; the middle one aligns on both axes instead of moving */
        var MOVE_BUTTON_ROWS = [
            ["upLeft",   "up",     "upRight"],
            ["left",     "center", "right"],
            ["downLeft", "down",   "downRight"]
        ];

        /* 値を持つ辺（設定の読み書きはこの並びで回す）/ The sides that hold a value */
        var MARGIN_SIDES = ["top", "bottom", "left", "right"];

        /* 分割ガイドの入力欄（設定の読み書きはこの並びで回す。ラベルのキーも兼ねる）
           The division fields; the settings are read and written in this order, and the keys double as label keys */
        var DIVIDE_VALUE_KEYS = ["rows", "columns", "rowGutter", "columnGutter", "extension"];

        /* 分割数と、その間隔を横に並べる2行（伸張はこの下に単独で置く）
           The two rows pairing a count with its gutter; the extension sits below them on its own */
        var DIVIDE_FIELD_ROWS = [
            { count: "rows",    gutter: "rowGutter" },
            { count: "columns", gutter: "columnGutter" }
        ];

        /* 分割ガイドのモード（ラジオボタンの並び）/ The division modes, in the order the radio buttons appear */
        var DIVIDE_MODES = [
            { key: "none",   labelKey: "divideNone" },
            { key: "cross",  labelKey: "divideCross" },
            { key: "custom", labelKey: "divideCustom" }
        ];

        /* 数値欄のツールチップのキー（項目名のキーから引く）/ Tooltip keys, looked up by the field key */
        var DIVIDE_TOOLTIP_KEYS = {
            rows:         "divideRows",
            columns:      "divideColumns",
            rowGutter:    "divideRowGutter",
            columnGutter: "divideColumnGutter",
            extension:    "divideExtension"
        };

        /* マージン欄の並び（3×3、真ん中は連動のチェック、"" は位置合わせの空セル）
           The 3x3 margin grid: the link checkbox sits in the middle, "" is an empty spacer cell */
        var MARGIN_FIELD_ROWS = [
            ["",     "top",    ""],
            ["left", "link",   "right"],
            ["",     "bottom", ""]
        ];

        /* 方向キーと、寄せる辺の対応（斜めは2辺へ同時に寄せる）
           The edges each direction moves to; a diagonal moves to both of them at once */
        var MOVE_SIDES_BY_DIRECTION = {
            up:        ["top"],
            left:      ["left"],
            right:     ["right"],
            down:      ["bottom"],
            upLeft:    ["left", "top"],
            upRight:   ["right", "top"],
            downLeft:  ["left", "bottom"],
            downRight: ["right", "bottom"]
        };

        /* 文字キーで実行する整列（値は整列ボタンの定義の tooltip キー）
           C は左右中央（Center）、M は上下中央（Middle）、X は上下左右中央
           The letter keys that run an align: C centres horizontally, M vertically, X on both axes */
        var ALIGN_KEY_SHORTCUTS = { C: "alignCenterH", M: "alignCenterV", X: "alignCenterAll" };

        /* ［裁ち落としに整列］を切り替えるキー / The key that toggles Align to Bleed */
        var BLEED_TOGGLE_KEY = "B";

        /* 矢印キーで実行する方向ボタン（斜めはキーが無いので4方向だけ）
           The arrow keys that fire a move button; the diagonals have no key of their own */
        var MOVE_DIRECTION_BY_KEY_NAME = { Up: "up", Left: "left", Right: "right", Down: "down" };

        /* 方向ボタンのアイコンで、寄せ先を表すケイ線を引く辺（斜めは2辺）
           The edges that get a rule in each move-button icon; a diagonal gets two */
        var MOVE_ICON_RULES = {
            up:        ["top"],
            down:      ["bottom"],
            left:      ["left"],
            right:     ["right"],
            upLeft:    ["top", "left"],
            upRight:   ["top", "right"],
            downLeft:  ["bottom", "left"],
            downRight: ["bottom", "right"]
        };

        /**
         * 整列アイコンボタンを1つ生成する
         * @param {Group} parentRow - 追加先の行グループ
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {void}
         */
        function addAlignButton(parentRow, buttonDef) {
            /* iconbutton ではなく button を使う（画像なしの iconbutton はクリックが届かないことがある）
               Use button, not iconbutton: an image-less iconbutton does not always receive clicks */
            var button = parentRow.add("button", undefined, "");
            fixControlSize(button, ICON_SIZE, ICON_SIZE);
            /* キー操作はラベルに出さず helpTip に書く / Shortcuts belong in the tooltip, not the label */
            button.helpTip = getLabel("tooltip", buttonDef.tooltip) + shortcutSuffix(buttonDef) + "  —  " +
                getLabel("tooltip", optionTooltipKey(buttonDef));
            button.iconType = buttonDef.iconType;
            button.alignMode = buttonDef.alignMode;
            button.isHover = false;
            button.onDraw = function() { drawAlignButton(this); };
            button.onClick = function() { runAlignButton(buttonDef); };
            attachHover(button);
        }

        /**
         * そのボタンに割り当てたショートカットを「（C）」の形で返す（無ければ空文字）
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {string} helpTip に添える文字列
         */
        function shortcutSuffix(buttonDef) {
            for (var keyName in ALIGN_KEY_SHORTCUTS) {
                if (!ALIGN_KEY_SHORTCUTS.hasOwnProperty(keyName)) { continue; }
                if (ALIGN_KEY_SHORTCUTS[keyName] === buttonDef.tooltip) { return keyHint(keyName); }
            }
            return "";
        }

        /**
         * ショートカットのキーを「（B）」の形にする（日本語は全角括弧、英語は半角）
         * @param {string} keyName - キー名
         * @returns {string} helpTip に添える文字列
         */
        function keyHint(keyName) {
            return (uiLang === "ja") ? "（" + keyName + "）" : " (" + keyName + ")";
        }

        /**
         * パレットのイベント（選択の取り直し・キー操作・閉じるときの後始末）を結線する
         * @param {Window} win - 対象のパレット
         * @returns {void}
         */
        function attachPaletteEvents(win) {
            /* 選択の変化はパレットへ操作しに来た瞬間に拾う（Illustrator にタイマーAPIが無いため）
               Illustrator has no timer API, so the selection is re-read when the user comes to the palette */
            win.onActivate = function() { onPaletteFocus(true); };
            try {
                win.addEventListener("mouseover", function() { onPaletteFocus(false); });
            } catch (mouseoverError) {}

            win.addEventListener("keydown", function(event) { onPaletteKeyDown(win, event); });

            /* 閉じるとき：［ガイドを保持］がOFFならこのパレットが作ったガイドを消し、参照を解放する
               On close: delete the guide unless "Keep Guides" is on, then release the reference */
            win.onClose = function() {
                if (!keepGuideOnClose) { removeMarginGuide(); }
                paletteWindow = null;
                $.global.__aiAlignToArtboardWindow = null;
                return true;
            };
        }

        /**
         * パレット上のキー操作を振り分ける
         * Esc で閉じ、↑↓←→ で方向ボタン、C/M/X で中央揃え、B で裁ち落としの切り替えを行う
         * @param {Window} win - 対象のパレット
         * @param {object} event - keydown イベント
         * @returns {void}
         */
        function onPaletteKeyDown(win, event) {
            if (event.keyName === "Escape") { win.close(); return; }
            /* 入力欄の打鍵と、コピーなどのショートカットには割り込まない
               Never steal a keystroke meant for a field, or a Cmd/Ctrl shortcut */
            if (isTextFieldTarget(event) || isCommandKeyPressed()) { return; }

            var directionKey = MOVE_DIRECTION_BY_KEY_NAME[event.keyName];
            if (directionKey) {
                /* フォーカスの移動に使われないよう、ここで止める / Stop the key from moving the focus */
                event.preventDefault();
                runMoveButton(directionKey);
                return;
            }

            if (event.keyName === BLEED_TOGGLE_KEY && alignToBleedCheckbox !== null) {
                event.preventDefault();
                alignToBleedCheckbox.value = alignToBleedCheckbox.value !== true;
                syncBleedField();
                return;
            }

            var tooltipKey = ALIGN_KEY_SHORTCUTS[event.keyName];
            var buttonDef = tooltipKey ? findAlignButtonDef(tooltipKey) : null;
            if (buttonDef !== null) {
                event.preventDefault();
                runAlignButton(buttonDef);
            }
        }

        /**
         * 整列ボタンを1つ実行する（クリックとショートカットで共通）
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {void}
         */
        function runAlignButton(buttonDef) {
            var workerResult = null;
            var didRun = runExclusive(function() { workerResult = runAlign(buttonDef); });
            /* 実行後の選択に合わせてディムと単位を更新してから、今回の結果を表示する
               （先に表示すると、この更新で選択が変わったと見なされて消えてしまう）
               Refresh first, then show this run's result; showing it first would be wiped by the refresh */
            onPaletteFocus(true);
            if (didRun) { showWorkerResult(workerResult); }
        }

        /**
         * tooltip キーから整列ボタンの定義を探す
         * @param {string} tooltipKey - 整列ボタンの定義の tooltip
         * @returns {object} 見つかったボタン定義。無ければ null
         */
        function findAlignButtonDef(tooltipKey) {
            for (var i = 0; i < ALL_ALIGN_BUTTONS.length; i++) {
                if (ALL_ALIGN_BUTTONS[i].tooltip === tooltipKey) { return ALL_ALIGN_BUTTONS[i]; }
            }
            return null;
        }

        /**
         * 移動ボタンを1つ生成する（directionKey が "center" なら中央揃えの整列ボタン）
         * @param {Group} parentRow - 追加先の行グループ
         * @param {string} directionKey - MOVE_SIDES_BY_DIRECTION のキー（"center" は十字の中央）
         * @returns {void}
         */
        function addMoveButton(parentRow, directionKey) {
            /* 十字の中央は、整列アイコンの列と同じ「水平・垂直方向中央に整列」にする
               The middle of the cross is the same centre-on-both-axes align button as in the icon row */
            if (directionKey === "center") {
                addAlignButton(parentRow, CENTER_ALIGN_BUTTON);
                return;
            }
            /* 整列アイコンと同じ理由で iconbutton ではなく button を使う / Same reason as the align icons: button, not iconbutton */
            var button = parentRow.add("button", undefined, "");
            fixControlSize(button, ICON_SIZE, ICON_SIZE);
            button.helpTip = getLabel("direction", directionKey) + "  —  " + getLabel("tooltip", "moveToEdge");
            button.directionKey = directionKey;
            button.isHover = false;
            button.onDraw = function() { drawMoveButton(this); };
            button.onClick = function() { runMoveButton(directionKey); };
            attachHover(button);
        }

        /**
         * 方向ボタンを1つ実行する（クリックと矢印キーで共通）
         * @param {string} directionKey - MOVE_SIDES_BY_DIRECTION のキー
         * @returns {void}
         */
        function runMoveButton(directionKey) {
            var workerResult = null;
            var didRun = runExclusive(function() { workerResult = runMove(directionKey); });
            /* 整列ボタンと同じ順序：先に選択を取り直してから今回の結果を出す
               Same order as the align buttons: refresh first, then show this run's result */
            onPaletteFocus(true);
            if (didRun) { showWorkerResult(workerResult); }
        }

        /**
         * Command（Mac）または Ctrl（Windows）が押されているか判定する
         * コピーなどのショートカットをこのパレットで横取りしないために見る
         * @returns {boolean} 押されていれば true（取得できない環境では false）
         */
        function isCommandKeyPressed() {
            try {
                var keyboardState = ScriptUI.environment.keyboardState;
                return keyboardState.metaKey === true || keyboardState.ctrlKey === true;
            } catch (keyboardError) {
                return false;
            }
        }

        /**
         * キーイベントの発生元が数値の入力欄かを判定する
         * 入力欄の↑↓は値の増減なので、方向ボタンの実行には使わない
         * @param {object} event - keydown イベント
         * @returns {boolean} 入力欄なら true
         */
        function isTextFieldTarget(event) {
            try {
                return event.target != null && event.target.type === "edittext";
            } catch (targetError) {
                return false;
            }
        }

        /**
         * 移動ボタンを十字に並べる
         * @param {Group} parentColumn - 追加先の行グループ（2カラムの左側）
         * @returns {void}
         */
        function addMoveButtonCross(parentColumn) {
            var crossGroup = parentColumn.add("group");
            crossGroup.orientation = "column";
            /* 右カラムの高さに引き伸ばさず、真ん中に置く / Centered in the column instead of stretched to its height */
            crossGroup.alignment = ["center", "center"];
            crossGroup.alignChildren = ["center", "top"];
            crossGroup.spacing = CROSS_GAP;

            for (var i = 0; i < MOVE_BUTTON_ROWS.length; i++) {
                var crossRow = crossGroup.add("group");
                setupRow(crossRow, "center", CROSS_GAP);
                for (var j = 0; j < MOVE_BUTTON_ROWS[i].length; j++) {
                    addMoveButton(crossRow, MOVE_BUTTON_ROWS[i][j]);
                }
            }

            addCenterAlignRow(crossGroup);
        }

        /**
         * 十字の下に、片方の軸だけ中央に整列するボタン2つを並べる
         * @param {Group} crossGroup - 追加先の十字のグループ
         * @returns {void}
         */
        function addCenterAlignRow(crossGroup) {
            var centerRow = crossGroup.add("group");
            setupRow(centerRow, "center", CROSS_GAP);
            /* 十字と続けて見えないよう、上に少し余白を取る / A little room so it does not read as a fourth row of the cross */
            centerRow.margins = [0, CENTER_ROW_TOP, 0, 0];
            for (var i = 0; i < CENTER_ALIGN_BUTTONS.length; i++) {
                addAlignButton(centerRow, CENTER_ALIGN_BUTTONS[i]);
            }
        }

        /**
         * ボタンの下を2カラムに分け、左に移動ボタンの十字・右に計測オプションのパネルを置く
         * マージンガイドと整列オプションはここには入れず、パレットの下に幅いっぱいで並べる
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addColumnsRow(targetWindow) {
            var columnsRow = targetWindow.add("group");
            columnsRow.orientation = "row";
            columnsRow.alignment = ["fill", "top"];
            columnsRow.alignChildren = ["fill", "fill"];
            columnsRow.spacing = COLUMN_SPACING;

            addMoveButtonCross(columnsRow);
            addOptionsPanel(columnsRow);
        }

        /**
         * マージンガイドのパネル（上下左右の数値欄＋定規の単位ラベル）を組み立てる
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addMarginPanel(targetWindow) {
            marginPanel = targetWindow.add("panel", undefined, marginPanelTitle());
            marginPanel.orientation = "column";
            marginPanel.alignment = ["fill", "top"];
            marginPanel.alignChildren = ["left", "top"];
            marginPanel.margins = PANEL_MARGINS;
            marginPanel.spacing = OPTION_SPACING;

            /* 前回の設定を引き継ぐ（残したガイドと表示が食い違わないように）/ Carry over the previous settings */
            var paletteSettings = loadPaletteSettings();
            marginsAreDefault = paletteSettings.usesDefaultMargins;

            showGuideCheckbox = marginPanel.add("checkbox", undefined, getLabel("checkbox", "showGuide"));
            showGuideCheckbox.helpTip = getLabel("tooltip", "showGuide");
            showGuideCheckbox.value = paletteSettings.showGuide;
            showGuideCheckbox.onClick = function() {
                /* ガイドを出すときは四辺をそろえたいことが多いので、連動も入れて4つの値をそろえる
                   Turning the guide on usually means an even inset, so the link goes on with it */
                if (showGuideCheckbox.value === true && linkMarginsCheckbox !== null && linkMarginsCheckbox.value !== true) {
                    linkMarginsCheckbox.value = true;
                    copyMarginToLinkedFields(marginFields.top);
                }
                syncMarginFields();
                syncGuideControls();
                runExclusive(refreshMarginGuide);
            };

            addMarginFieldRows(marginPanel, paletteSettings);

            keepGuideCheckbox = marginPanel.add("checkbox", undefined, getLabel("checkbox", "keepGuide"));
            keepGuideCheckbox.helpTip = getLabel("tooltip", "keepGuide");
            keepGuideCheckbox.value = paletteSettings.keepGuide;
            keepGuideCheckbox.onClick = function() { syncGuideControls(); };

            syncGuideControls();
        }

        /**
         * マージンガイドのパネル名に、いまの定規の単位を添える
         * 欄が4つあるので、単位は欄ごとではなくパネル名にまとめて出す
         * @returns {string} パネル名
         */
        function marginPanelTitle() {
            var unitLabel = currentUnitInfo.label;
            return getLabel("panel", "guide") + (uiLang === "ja" ? "（" + unitLabel + "）" : " (" + unitLabel + ")");
        }

        /**
         * 上下左右のマージン欄を3×3で組み立てる
         * @param {Panel} marginPanel - 追加先のパネル
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addMarginFieldRows(marginPanel, paletteSettings) {
            for (var i = 0; i < MARGIN_FIELD_ROWS.length; i++) {
                var fieldRow = marginPanel.add("group");
                fieldRow.orientation = "row";
                /* パネルの幅いっぱいに広げず、真ん中に置く / Centred in the panel instead of stretched */
                fieldRow.alignment = ["center", "top"];
                fieldRow.alignChildren = ["center", "center"];
                fieldRow.spacing = LABEL_FIELD_SPACING;
                for (var j = 0; j < MARGIN_FIELD_ROWS[i].length; j++) {
                    addMarginCell(fieldRow, MARGIN_FIELD_ROWS[i][j], paletteSettings);
                }
            }
            syncMarginFields();
        }

        /**
         * ［ガイドを追加］と［連動］の状態に合わせてマージン欄の使える・使えないを切り替える
         * ガイドを出していないあいだは4辺まとめてディムにし、連動しているあいだは［上］の1つで4辺が決まるので、ほかをディムにする
         * @returns {void}
         */
        function syncMarginFields() {
            var addsGuide, isLinked, side, i;
            if (linkMarginsCheckbox === null) { return; }
            addsGuide = showGuideCheckbox !== null && showGuideCheckbox.value === true;
            isLinked = linkMarginsCheckbox.value === true;
            linkMarginsCheckbox.enabled = addsGuide;
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                if (!marginFields[side]) { continue; }
                marginFields[side].enabled = addsGuide && (isLinked ? (side === "top") : true);
            }
        }

        /**
         * 3×3の1セルを生成する（辺なら「ラベル＋入力欄」、"link" なら連動のチェック、"" なら空セル）
         * どのセルも同じ幅にして、上下左右が十字に並ぶようにする
         * @param {Group} fieldRow - 追加先の行グループ
         * @param {string} cellKey - "top" / "bottom" / "left" / "right" / "link" / ""
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addMarginCell(fieldRow, cellKey, paletteSettings) {
            var cellGroup = fieldRow.add("group");
            cellGroup.orientation = "row";
            cellGroup.alignment = ["center", "center"];
            cellGroup.alignChildren = ["center", "center"];
            cellGroup.spacing = LABEL_FIELD_SPACING;
            cellGroup.minimumSize.width = MARGIN_CELL_WIDTH[uiLang];
            if (cellKey === "") { return; }

            if (cellKey === "link") {
                linkMarginsCheckbox = cellGroup.add("checkbox", undefined, getLabel("checkbox", "linkMargins"));
                linkMarginsCheckbox.helpTip = getLabel("tooltip", "linkMargins");
                linkMarginsCheckbox.value = paletteSettings.linkMargins;
                /* 連動に切り替えた時点で、4つを上のマージンにそろえる
                   Turning the link back on levels the four values off against the top margin */
                linkMarginsCheckbox.onClick = function() {
                    copyMarginToLinkedFields(marginFields.top);
                    syncMarginFields();
                    runExclusive(refreshMarginGuide);
                    savePaletteSettings();
                };
                return;
            }

            cellGroup.add("statictext", undefined, labelText("fieldLabel", cellKey));
            var field = cellGroup.add("edittext", undefined, paletteSettings.margins[cellKey]);
            field.characters = FIELD_CHARS;
            field.helpTip = getLabel("tooltip", "margin");
            changeValueByArrowKey(field);
            /* 確定（Enter・フォーカス移動）でガイドを描き直す
               入力途中で毎回描き直すとそのつど委譲が走るため、描き直しは onChange だけにする
               Only redraw the guide when the field commits, not on every keystroke */
            field.onChange = function() {
                /* 一度でも触られたら、単位が変わっても既定値では上書きしない */
                marginsAreDefault = false;
                copyMarginToLinkedFields(field);
                runExclusive(refreshMarginGuide);
                savePaletteSettings();
            };
            marginFields[cellKey] = field;
        }

        /**
         * まだ触られていないマージン欄を、いまの定規の単位の既定値で埋め直す
         * パネルを組み立てる時点では定規の単位が分からないため、分かった時点で入れ直す
         * @returns {void}
         */
        function fillDefaultMargins() {
            var defaultText, side, i;
            if (!marginsAreDefault) { return; }
            defaultText = String(currentUnitInfo.defaultMargin);
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                if (!marginFields[side]) { continue; }
                marginFields[side].text = defaultText;
            }
            savePaletteSettings();
        }

        /**
         * 伸張欄がまだ既定値のままなら、いまの定規の単位の既定値に入れ直す
         * @returns {void}
         */
        function fillDefaultExtension() {
            if (!extensionIsDefault || !divideFields.extension) { return; }
            divideFields.extension.text = String(currentUnitInfo.defaultExtension);
            savePaletteSettings();
        }

        /**
         * ［連動］がONのとき、編集した欄の値を残りのマージン欄へ写す
         * @param {EditText} sourceField - 値の元にする入力欄
         * @returns {void}
         */
        function copyMarginToLinkedFields(sourceField) {
            if (linkMarginsCheckbox === null || linkMarginsCheckbox.value !== true) { return; }
            if (!sourceField) { return; }
            for (var side in marginFields) {
                if (!marginFields.hasOwnProperty(side)) { continue; }
                if (marginFields[side] === sourceField) { continue; }
                marginFields[side].text = sourceField.text;
            }
        }

        /**
         * 分割ガイドのパネル（なし・十字・カスタムの切り替えと、行・列・行間・列間・伸張）を組み立てる
         * マージンの内側を等分する位置にガイドを引き、伸張のぶんだけアートボードの外へ伸ばす
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addDividePanel(targetWindow) {
            var dividePanel, paletteSettings, extensionRow, i;

            /* 前回の設定を引き継ぐ（残したガイドと表示が食い違わないように）/ Carry over the previous settings */
            paletteSettings = loadPaletteSettings();
            extensionIsDefault = paletteSettings.usesDefaultExtension;

            dividePanel = targetWindow.add("panel", undefined, getLabel("panel", "divide"));
            dividePanel.orientation = "column";
            dividePanel.alignment = ["fill", "top"];
            dividePanel.alignChildren = ["left", "top"];
            dividePanel.margins = PANEL_MARGINS;
            dividePanel.spacing = OPTION_SPACING;

            addDivideModeRow(dividePanel, paletteSettings);
            for (i = 0; i < DIVIDE_FIELD_ROWS.length; i++) {
                addDivideFieldRow(dividePanel, DIVIDE_FIELD_ROWS[i], paletteSettings);
            }

            /* 伸張は行・列のどちらにも掛かるので、対にせず単独の行に置く
               The extension applies to both axes, so it sits on its own row instead of pairing off */
            extensionRow = dividePanel.add("group");
            setupRow(extensionRow, "left", LABEL_FIELD_SPACING);
            addDivideField(extensionRow, "extension", paletteSettings, DIVIDE_LABEL_WIDTH[uiLang], true);

            artboardEdgeCheckbox = dividePanel.add("checkbox", undefined, getLabel("checkbox", "artboardEdge"));
            artboardEdgeCheckbox.helpTip = getLabel("tooltip", "artboardEdge");
            artboardEdgeCheckbox.value = paletteSettings.artboardEdge;
            artboardEdgeCheckbox.onClick = function() {
                syncDivideFields();
                runExclusive(refreshMarginGuide);
            };

            syncDivideFields();
        }

        /**
         * 分割の仕方を選ぶラジオボタンの行（なし・十字・カスタム）を組み立てる
         * @param {Panel} dividePanel - 追加先のパネル
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addDivideModeRow(dividePanel, paletteSettings) {
            var modeRow, modeDef, radio, i;
            modeRow = dividePanel.add("group");
            setupRow(modeRow, "left", RADIO_GAP);
            for (i = 0; i < DIVIDE_MODES.length; i++) {
                modeDef = DIVIDE_MODES[i];
                radio = modeRow.add("radiobutton", undefined, getLabel("radio", modeDef.labelKey));
                radio.helpTip = getLabel("tooltip", modeDef.labelKey);
                radio.value = (modeDef.key === paletteSettings.divideMode);
                radio.onClick = function() {
                    syncDivideFields();
                    runExclusive(refreshMarginGuide);
                };
                divideRadios[modeDef.key] = radio;
            }
        }

        /**
         * 分割数と、その間隔を横に並べた1行を組み立てる
         * @param {Panel} dividePanel - 追加先のパネル
         * @param {object} rowDef - DIVIDE_FIELD_ROWS の1行
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addDivideFieldRow(dividePanel, rowDef, paletteSettings) {
            var fieldRow = dividePanel.add("group");
            setupRow(fieldRow, "left", LABEL_FIELD_SPACING);
            addDivideField(fieldRow, rowDef.count, paletteSettings, DIVIDE_LABEL_WIDTH[uiLang], false);
            addDivideField(fieldRow, rowDef.gutter, paletteSettings, DIVIDE_GUTTER_LABEL_WIDTH[uiLang], true);
        }

        /**
         * 分割ガイドの「項目名＋数値欄」を1組ぶん作る
         * 項目名は右そろえにして、日英で長さが違っても数値欄の頭がそろうようにする
         * @param {Group} fieldRow - 追加先の行グループ
         * @param {string} valueKey - DIVIDE_VALUE_KEYS のキー
         * @param {object} paletteSettings - 引き継いだ設定
         * @param {number} labelWidth - 項目名の幅
         * @param {boolean} hasUnit - 定規の単位ラベルを添えるなら true（分割数は単位を持たない）
         * @returns {void}
         */
        function addDivideField(fieldRow, valueKey, paletteSettings, labelWidth, hasUnit) {
            var label, field, unitLabel;

            label = fieldRow.add("statictext", undefined, labelText("fieldLabel", valueKey));
            label.preferredSize.width = labelWidth;
            label.justify = "right";

            field = fieldRow.add("edittext", undefined, paletteSettings.divideValues[valueKey]);
            field.characters = FIELD_CHARS;
            field.helpTip = getLabel("tooltip", DIVIDE_TOOLTIP_KEYS[valueKey]);
            changeValueByArrowKey(field);
            /* 確定（Enter・フォーカス移動）でガイドを描き直す / Only redraw the guide when the field commits */
            field.onChange = function() {
                /* 一度でも触られたら、単位が変わっても既定値では上書きしない */
                if (valueKey === "extension") { extensionIsDefault = false; }
                runExclusive(refreshMarginGuide);
                savePaletteSettings();
            };
            divideFields[valueKey] = field;

            if (!hasUnit) { return; }
            unitLabel = fieldRow.add("statictext", undefined, currentUnitInfo.label);
            unitLabel.preferredSize.width = UNIT_LABEL_WIDTH;
            divideUnitLabels.push(unitLabel);
        }

        /**
         * 分割の仕方に合わせて数値欄のディムを切り替える
         * ［なし］はすべて、［十字］は行・列と間隔をディムにする（伸張は十字にも効く）
         * @returns {void}
         */
        function syncDivideFields() {
            var mode, isCustom, usesExtension, valueKey, i;
            mode = readDivideMode();
            isCustom = (mode === "custom");
            /* 伸張は分割のガイドとアートボードのエッジの両方に掛かるので、どちらかがあれば使える
               The extension applies to the division guides and to the artboard edges alike */
            usesExtension = (mode !== "none") || readArtboardEdge();
            for (i = 0; i < DIVIDE_VALUE_KEYS.length; i++) {
                valueKey = DIVIDE_VALUE_KEYS[i];
                if (!divideFields[valueKey]) { continue; }
                divideFields[valueKey].enabled = (valueKey === "extension") ? usesExtension : isCustom;
            }
            syncGuideControls();
        }

        /**
         * ［アートボードのエッジ］が入っているかを返す
         * @returns {boolean} 入っていれば true
         */
        function readArtboardEdge() {
            return artboardEdgeCheckbox !== null && artboardEdgeCheckbox.value === true;
        }

        /**
         * 分割ガイドの単位ラベルを、いまの定規の単位に書き換える
         * @returns {void}
         */
        function updateDivideUnitLabels() {
            for (var i = 0; i < divideUnitLabels.length; i++) {
                divideUnitLabels[i].text = currentUnitInfo.label;
            }
        }

        /**
         * いま選ばれている分割の仕方を返す
         * @returns {string} "none" / "cross" / "custom"
         */
        function readDivideMode() {
            var mode, i;
            for (i = 0; i < DIVIDE_MODES.length; i++) {
                mode = DIVIDE_MODES[i].key;
                if (divideRadios[mode] && divideRadios[mode].value === true) { return mode; }
            }
            return DEFAULT_DIVIDE_MODE;
        }

        /**
         * ガイドを描くものがあるか（マージンのガイドか分割ガイドのどちらかが有効なら true）
         * @returns {boolean} 描くものがあれば true
         */
        function hasGuideToDraw() {
            if (showGuideCheckbox !== null && showGuideCheckbox.value === true) { return true; }
            if (readArtboardEdge()) { return true; }
            return readDivideMode() !== "none";
        }

        /**
         * ガイドの有無に合わせて［ガイドを保持］を更新する
         * ガイドを出していないときは保持しようがないのでディムする（設定は残したいので値は落とさない）
         * @returns {void}
         */
        function syncGuideControls() {
            if (showGuideCheckbox === null || keepGuideCheckbox === null) { return; }
            keepGuideCheckbox.enabled = hasGuideToDraw();
            keepGuideOnClose = keepGuideCheckbox.value === true;
            savePaletteSettings();
        }

        /**
         * 前回のパレット設定を常駐エンジンから読み出す（控えが無ければ既定値）
         * @returns {object} { showGuide, keepGuide, usesDefaultMargins, margins, divideMode, divideValues, usesDefaultExtension, artboardEdge, linkMargins, alignToBleed, bleed, perArtboard }
         */
        function loadPaletteSettings() {
            var stored = $.global[SETTINGS_KEY];
            if (!stored) {
                return {
                    showGuide:    DEFAULT_SHOW_GUIDE,
                    keepGuide:    DEFAULT_KEEP_GUIDE,
                    usesDefaultMargins: true,
                    margins:      marginStrings(null, String(currentUnitInfo.defaultMargin)),
                    divideMode:   DEFAULT_DIVIDE_MODE,
                    divideValues: divideValueStrings(null),
                    usesDefaultExtension: true,
                    artboardEdge: DEFAULT_ARTBOARD_EDGE,
                    linkMargins:  DEFAULT_LINK_MARGINS,
                    alignToBleed: DEFAULT_ALIGN_TO_BLEED,
                    bleed:        String(DEFAULT_BLEED),
                    perArtboard:  DEFAULT_ALIGN_PER_ARTBOARD
                };
            }
            return {
                showGuide:    stored.showGuide === true,
                keepGuide:    stored.keepGuide === true,
                usesDefaultMargins: stored.margins == null,
                margins:      marginStrings(stored.margins, String(currentUnitInfo.defaultMargin)),
                divideMode:   knownDivideMode(stored.divideMode),
                divideValues: divideValueStrings(stored.divideValues),
                usesDefaultExtension: !(stored.divideValues && stored.divideValues.extension != null),
                artboardEdge: stored.artboardEdge === true,
                linkMargins:  (stored.linkMargins != null) ? (stored.linkMargins === true) : DEFAULT_LINK_MARGINS,
                alignToBleed: stored.alignToBleed === true,
                bleed:        (stored.bleed != null) ? String(stored.bleed) : String(DEFAULT_BLEED),
                perArtboard:  stored.perArtboard === true
            };
        }

        /**
         * 控えのマージン4値を、欠けている辺を既定値で埋めた文字列の組にそろえる
         * @param {object} storedMargins - 控えの値（無ければ null）
         * @param {string} fallback - 欠けている辺に使う値
         * @returns {object} { top: string, bottom: string, left: string, right: string }
         */
        function marginStrings(storedMargins, fallback) {
            var margins, side, i;
            margins = {};
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                margins[side] = (storedMargins && storedMargins[side] != null) ? String(storedMargins[side]) : fallback;
            }
            return margins;
        }

        /**
         * 控えの分割の仕方が並びにあるものか確かめる（無ければ既定値）
         * @param {string} storedMode - 控えの値
         * @returns {string} "none" / "cross" / "custom"
         */
        function knownDivideMode(storedMode) {
            for (var i = 0; i < DIVIDE_MODES.length; i++) {
                if (DIVIDE_MODES[i].key === storedMode) { return storedMode; }
            }
            return DEFAULT_DIVIDE_MODE;
        }

        /**
         * 控えの分割ガイドの値を、欠けている項目を既定値で埋めた文字列の組にそろえる
         * @param {object} storedValues - 控えの値（無ければ null）
         * @returns {object} DIVIDE_VALUE_KEYS をキーにした文字列の組
         */
        function divideValueStrings(storedValues) {
            var values, valueKey, i;
            values = {};
            for (i = 0; i < DIVIDE_VALUE_KEYS.length; i++) {
                valueKey = DIVIDE_VALUE_KEYS[i];
                values[valueKey] = (storedValues && storedValues[valueKey] != null) ?
                    String(storedValues[valueKey]) : defaultDivideValue(valueKey);
            }
            return values;
        }

        /**
         * 分割ガイドの欄の初期値を返す（伸張だけは定規の単位ごとに決まる）
         * @param {string} valueKey - DIVIDE_VALUE_KEYS のキー
         * @returns {string} 初期値
         */
        function defaultDivideValue(valueKey) {
            if (valueKey === "extension") { return String(currentUnitInfo.defaultExtension); }
            return String(DEFAULT_DIVIDE_VALUES[valueKey]);
        }

        /**
         * 現在のパレット設定を常駐エンジンに控える
         * @returns {void}
         */
        function savePaletteSettings() {
            if (!isPaletteReady) { return; }
            $.global[SETTINGS_KEY] = {
                showGuide:    showGuideCheckbox !== null && showGuideCheckbox.value === true,
                keepGuide:    keepGuideCheckbox !== null && keepGuideCheckbox.value === true,
                margins:      readMarginTexts(),
                divideMode:   readDivideMode(),
                divideValues: readDivideValueTexts(),
                artboardEdge: readArtboardEdge(),
                linkMargins:  linkMarginsCheckbox !== null && linkMarginsCheckbox.value === true,
                alignToBleed: alignToBleedCheckbox !== null && alignToBleedCheckbox.value === true,
                bleed:        (bleedField !== null) ? String(bleedField.text) : String(DEFAULT_BLEED),
                perArtboard:  alignPerArtboardCheckbox !== null && alignPerArtboardCheckbox.value === true
            };
        }

        /**
         * 計測オプションパネル（プレビュー境界・字形の境界に整列・行揃えを変更）を組み立てる
         * 境界の測り方と、それに付随する行揃えの変更だけを置く（整列オプションは別パネル）
         * @param {Group} parentRow - 追加先の2カラムの行グループ
         * @returns {void}
         */
        function addOptionsPanel(parentRow) {
            var optionsPanel = parentRow.add("panel", undefined, getLabel("panel", "options"));
            optionsPanel.orientation = "column";
            /* パネル自身は右カラムの幅いっぱいに、中のチェックボックスは左そろえ（fill の継承を打ち消す）
               The panel fills the width of its column while its checkboxes stay left-aligned, cancelling the inherited fill */
            optionsPanel.alignment = ["fill", "top"];
            optionsPanel.alignChildren = ["left", "center"];
            optionsPanel.margins = PANEL_MARGINS;
            optionsPanel.spacing = OPTION_SPACING;

            /* ここでは控えの値を置くだけで、実際の初期値は loadBoundsPreferences() が入れ直す
               （常駐パレットの app は当てにならないため、環境設定はメインエンジンへ問い合わせる）
               Only the fallbacks are set here; loadBoundsPreferences() replaces them with the real preferences */
            previewBoundsCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox", "previewBounds"));
            previewBoundsCheckbox.helpTip = getLabel("tooltip", "previewBounds");
            previewBoundsCheckbox.value = DEFAULT_PREVIEW_BOUNDS;

            /* 字形の境界はポイント文字・エリア内文字をまとめてON/OFFする / Glyph bounds toggles point & area type together */
            glyphBoundsCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox", "glyphBounds"));
            glyphBoundsCheckbox.helpTip = getLabel("tooltip", "glyphBounds");
            glyphBoundsCheckbox.value = DEFAULT_GLYPH_BOUNDS;

            /* 行揃えはこのスクリプト内だけの設定 / Justification is script-local */
            changeJustificationCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox", "changeJustification"));
            changeJustificationCheckbox.helpTip = getLabel("tooltip", "changeJustification");
            changeJustificationCheckbox.value = DEFAULT_CHANGE_JUSTIFICATION;
        }

        /**
         * 整列オプションパネル（裁ち落としに整列・アートボードごとに整列）を組み立てる
         * 裁ち落としは印刷の値なので、定規の単位に追従させず mm で扱う
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addDestinationPanel(targetWindow) {
            var paletteSettings = loadPaletteSettings();

            var destinationPanel = targetWindow.add("panel", undefined, getLabel("panel", "destination"));
            destinationPanel.orientation = "column";
            destinationPanel.alignment = ["fill", "top"];
            destinationPanel.alignChildren = ["left", "center"];
            destinationPanel.margins = PANEL_MARGINS;
            destinationPanel.spacing = OPTION_SPACING;

            var bleedRow = destinationPanel.add("group");
            setupRow(bleedRow, "left", LABEL_FIELD_SPACING);

            alignToBleedCheckbox = bleedRow.add("checkbox", undefined, getLabel("checkbox", "alignToBleed"));
            alignToBleedCheckbox.helpTip = getLabel("tooltip", "alignToBleed") + keyHint(BLEED_TOGGLE_KEY);
            alignToBleedCheckbox.value = paletteSettings.alignToBleed;
            alignToBleedCheckbox.onClick = function() { syncBleedField(); };

            bleedField = bleedRow.add("edittext", undefined, paletteSettings.bleed);
            bleedField.characters = FIELD_CHARS;
            bleedField.helpTip = getLabel("tooltip", "alignToBleed");
            changeValueByArrowKey(bleedField);
            bleedField.onChange = function() { savePaletteSettings(); };

            bleedRow.add("statictext", undefined, BLEED_UNIT_LABEL);

            alignPerArtboardCheckbox = destinationPanel.add("checkbox", undefined, getLabel("checkbox", "perArtboard"));
            alignPerArtboardCheckbox.helpTip = getLabel("tooltip", "perArtboard");
            alignPerArtboardCheckbox.value = paletteSettings.perArtboard;
            alignPerArtboardCheckbox.onClick = function() { savePaletteSettings(); };

            syncBleedField();
        }

        /**
         * ［裁ち落としに整列］の状態に合わせて裁ち落としの数値欄を更新する
         * OFFのときは使わない値なのでディムする（設定は残したいので値は落とさない）
         * @returns {void}
         */
        function syncBleedField() {
            if (alignToBleedCheckbox === null || bleedField === null) { return; }
            bleedField.enabled = alignToBleedCheckbox.value === true;
            savePaletteSettings();
        }

        /**
         * ［プレビュー境界］［字形の境界に整列］の現在の環境設定をチェックボックスへ反映する
         * 常駐パレットの app は当てにならないためメインエンジンへ問い合わせ、取れなければ DEFAULT_* のままにする
         * @returns {void}
         */
        function loadBoundsPreferences() {
            if (previewBoundsCheckbox === null || glyphBoundsCheckbox === null) { return; }
            var workerResult = runWorker("btGetPreferenceFlags();");
            if (workerResult === null || workerResult.indexOf("ERR:") === 0) { return; }
            var parts = workerResult.split("|");
            if (parts.length < 2) { return; }
            previewBoundsCheckbox.value = (parts[0] === "1");
            glyphBoundsCheckbox.value = (parts[1] === "1");
        }

        /**
         * 状況表示を書き換える（固定幅で切り詰められるため、全文は helpTip に入れる）
         * @param {string} message - 表示する文言
         * @returns {void}
         */
        function setStatus(message) {
            if (statusText === null) { return; }
            statusText.text = message;
            statusText.helpTip = message;
        }

        /**
         * 状況表示の行を組み立てる（中身で幅が変わらないよう固定幅にする）
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addStatusLine(targetWindow) {
            statusText = targetWindow.add("statictext", undefined, "", { truncate: "end" });
            statusText.alignment = ["fill", "center"];
            statusText.preferredSize.width = STATUS_WIDTH;
            statusText.maximumSize.width = STATUS_WIDTH;
        }

        /**
         * すでに開いているパレットがあれば閉じる（多重起動防止と、修正後のコードで開き直すため）
         * @returns {void}
         */
        function closeExistingPalette() {
            try {
                if (paletteWindow) { paletteWindow.close(); }
            } catch (staleReferenceError) {} /* 参照が無効なら閉じる必要もない / A stale reference needs no closing */
            paletteWindow = null;
            $.global.__aiAlignToArtboardWindow = null;
        }

        /**
         * 整列パレットを組み立てて表示する
         * @returns {void}
         */
        function showPalette() {
            /* 多重起動防止：開いているパレットは必ず閉じてから作り直す / Close any open palette first */
            closeExistingPalette();
            isPaletteReady = false;
            initIconColors();

            var win = new Window("palette", getLabel("dialog", "title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
            win.orientation = "column";
            win.alignChildren = ["fill", "top"];
            win.margins = WINDOW_MARGINS;
            win.spacing = WINDOW_SPACING;

            addColumnsRow(win);
            addMarginPanel(win);
            addDividePanel(win);
            addDestinationPanel(win);
            addStatusLine(win);
            /* ここまでで全コントロールがそろうので、以降は操作のたびに控える
               Every control now exists, so from here on each change is stored */
            isPaletteReady = true;

            attachPaletteEvents(win);

            /* 常駐参照：GC 回避と多重起動の検出を兼ねる / Persistent reference: avoids GC and detects a second launch */
            paletteWindow = win;
            $.global.__aiAlignToArtboardWindow = win;
            /* 環境設定の現在値をチェックボックスに入れてから表示する（表示後だと目の前で切り替わる）
               Load the preferences before showing, so the checkboxes do not flip in front of the user */
            loadBoundsPreferences();
            win.layout.layout(true);
            win.show();
            refreshPaletteState();
            /* 引き継いだ設定でガイドを描き直す（前回残したガイドと入力欄の値をそろえる）
               Redraw the guide from the carried-over settings so it matches the field */
            if (hasGuideToDraw()) { runExclusive(refreshMarginGuide); }
        }

        // =========================================
        // BridgeTalk ワーカー / BridgeTalk workers
        // =========================================
        /* 以下の bt* 関数は toString() で連結してメインエンジンへ送るため、次の制約がある
           これらにJSDocを付けないのも同じ理由（出力が壊れて構文エラーになる）
           These bt* functions are stringified with toString() and shipped to the main engine, so:
           - 行コメント（//）は使わず、ブロックコメント（/* *\/）だけにする
           - 文は必ずセミコロンで終える（toString で改行が失われても壊れないように）
           - パレット側の変数は参照しない。必要な値は options で受け取る */

        function btAlignSelection(options) {
            var doc, selectedItems, result;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            selectedItems = btResolveSelection(doc);
            if (selectedItems === null) { return "NOSEL"; }
            result = btRunPerArtboard(doc, selectedItems, options, btAlignItems, "OK");
            if (result.indexOf("OK") !== 0) { return result; }
            /* ガイドは最後に1回だけ描く（アートボードごとに回すと、そのつど張り替えることになる）*/
            try {
                btDrawMarginGuide(doc, options);
            } catch (guideError) {
                return "ERR:" + guideError;
            }
            return result;
        }

        function btResolveSelection(doc) {
            var selectedItems;
            selectedItems = doc.selection;
            /* 文字を選択しているときは、そのストーリーのテキストオブジェクトに置き換える */
            if (selectedItems && !(selectedItems instanceof Array)) {
                selectedItems = btPromoteTextRange(doc, selectedItems);
            }
            if (!(selectedItems instanceof Array) || selectedItems.length === 0) { return null; }
            return selectedItems;
        }

        function btSelectItems(doc, items) {
            var i;
            doc.selection = null;
            for (i = 0; i < items.length; i++) {
                try { items[i].selected = true; } catch (selectError) {}
            }
        }

        function btRunPerArtboard(doc, selectedItems, options, runBucket, okMarker) {
            var buckets, result, lastResult, i;
            /* ［アートボードごとに整列］がOFF、またはアートボードが1つなら、まとめて1回で済ませる */
            if (options.perArtboard !== true || doc.artboards.length < 2) {
                return runBucket(doc, selectedItems, options);
            }
            buckets = btGroupItemsByArtboard(doc, selectedItems);
            lastResult = okMarker;
            for (i = 0; i < buckets.length; i++) {
                /* 整列も移動も選択を見て動くので、そのアートボードのぶんだけ選び直す */
                btSelectItems(doc, buckets[i]);
                result = runBucket(doc, buckets[i], options);
                /* 1つでも失敗したら、そこで止めて理由を返す */
                if (result.indexOf(okMarker) !== 0) { return result; }
                if (result !== okMarker) { lastResult = result; }
            }
            /* アートボードごとに選び直したので、最後に元の選択へ戻す */
            btSelectItems(doc, selectedItems);
            return lastResult;
        }

        function btGroupItemsByArtboard(doc, items) {
            var buckets, artboardIndexes, itemBounds, artboardIndex, position, i, j;
            buckets = [];
            artboardIndexes = [];
            for (i = 0; i < items.length; i++) {
                itemBounds = btUnionBounds([items[i]], true, false);
                artboardIndex = btFindOverlappingArtboardIndex(doc, itemBounds, doc.artboards.getActiveArtboardIndex());
                if (artboardIndex < 0) { artboardIndex = btFindNearestArtboardIndex(doc, itemBounds); }
                position = -1;
                for (j = 0; j < artboardIndexes.length; j++) {
                    if (artboardIndexes[j] === artboardIndex) { position = j; break; }
                }
                if (position < 0) {
                    artboardIndexes.push(artboardIndex);
                    buckets.push([]);
                    position = buckets.length - 1;
                }
                buckets[position].push(items[i]);
            }
            return buckets;
        }

        function btCopyOptions(options) {
            var copy, key;
            copy = {};
            for (key in options) {
                if (options.hasOwnProperty(key)) { copy[key] = options[key]; }
            }
            return copy;
        }

        function btAlignItems(doc, selectedItems, options) {
            var needsGroup, didGroup, previousPreferences, justification, failure;
            /* 寄せ先はこの組だけの判定で決まるので、呼び出し元の options は書き換えない */
            options = btCopyOptions(options);
            needsGroup = selectedItems.length > 1;
            if (needsGroup && btSpansMultipleLayers(selectedItems)) { return "MULTILAYER"; }
            btActivateArtboardForSelection(doc, selectedItems);
            /* すでにマージンの位置に寄っているなら、次はその外のアートボードの端へ寄せる（方向ボタンと同じ二段階）*/
            options.marginPt = btResolveStepMargin(doc, selectedItems, options);
            previousPreferences = btReadPreferences();
            justification = { previous: null, changed: null };
            /* 実際にグループ化できたかを控える。needsGroup で解除すると、グループ化の前で例外が出たときに
               選択していた既存のグループを解除してしまう
               Track whether the group actually happened; keying the ungroup off needsGroup would
               dissolve the user's own groups when something throws before the group runs */
            didGroup = false;
            /* 整列できなかった理由。finally で後始末をしてから返す / Why the align failed; returned after the teardown */
            failure = null;
            try {
                btWritePreferences(options.previewBounds === true, options.glyphBounds === true, options.glyphBounds === true);
                /* 書いた環境設定を整列コマンドに拾わせるため、いったん反映させる
                   ここを飛ばすと「字形の境界に整列」が効かないまま整列されることがある */
                app.redraw();
                justification = btApplyJustification(selectedItems, options);
                if (needsGroup) {
                    app.executeMenuCommand("group");
                    didGroup = true;
                }
                if (!btRunAlignCommands(doc, options)) {
                    failure = "NOTARGET";
                } else {
                    btNudgeSelection(doc,
                        options.offsetX * options.marginPt + options.centerOffsetX,
                        options.offsetY * options.marginPt + options.centerOffsetY);
                }
            } catch (alignError) {
                failure = "ERR:" + alignError;
            } finally {
                btFinishAlign(selectedItems, didGroup, previousPreferences, justification.previous, failure);
            }
            if (failure !== null) { return failure; }
            try {
                /* 整列コマンドが「字形の境界に整列」を拾えていないことがあるため、
                   字形を実測して目標位置との差を打ち消す（拾えていれば差は0で何も動かない）*/
                btApplyGlyphCorrection(doc, options);
            } catch (correctionError) {
                return "ERR:" + correctionError;
            }
            if (justification.changed !== null) { return "OK:" + justification.changed; }
            return "OK";
        }

        function btApplyJustification(selectedItems, options) {
            var previousJustification;
            if (!options.changeJustification || !options.justification) { return { previous: null, changed: null }; }
            if (!btIsSingleLineTextFrame(selectedItems)) { return { previous: null, changed: null }; }
            previousJustification = btSetJustification(selectedItems[0], options.justification);
            if (previousJustification === null) { return { previous: null, changed: null }; }
            /* もとから同じ行揃えだったときは「変更した」と言わない */
            if (previousJustification === btJustificationByName(options.justification)) {
                return { previous: previousJustification, changed: null };
            }
            return { previous: previousJustification, changed: options.justification };
        }

        function btFinishAlign(selectedItems, didGroup, previousPreferences, previousJustification, failure) {
            if (didGroup) { app.executeMenuCommand("ungroup"); }
            /* 整列が環境設定を使い終えてから戻す（先に戻すと反映前の値で整列されることがある）*/
            app.redraw();
            btWritePreferences(previousPreferences.previewBounds, previousPreferences.pointText, previousPreferences.areaText);
            /* 整列できなかったときは、先に変えた行揃えも元に戻す（例外で抜けた場合も同じ）
               Roll the justification back whenever the align did not go through, exceptions included */
            if (failure === null) { return; }
            btRestoreJustification(selectedItems, previousJustification);
        }

        function btRestoreJustification(selectedItems, previousJustification) {
            if (previousJustification === null) { return; }
            try {
                selectedItems[0].textRange.paragraphAttributes.justification = previousJustification;
            } catch (restoreError) {}
        }

        function btMoveToEdgeOrGuide(options) {
            var doc, selectedItems;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            selectedItems = btResolveSelection(doc);
            if (selectedItems === null) { return "NOSEL"; }
            if (doc.artboards.length === 0) { return "NODOC"; }
            return btRunPerArtboard(doc, selectedItems, options, btMoveItems, "MOVED");
        }

        function btMoveItems(doc, selectedItems, options) {
            var needsGroup, didGroup, justification, failure, bounds, artboardRect, deltaX, deltaY, side, delta, i;
            needsGroup = selectedItems.length > 1;
            if (needsGroup && btSpansMultipleLayers(selectedItems)) { return "MULTILAYER"; }
            btActivateArtboardForSelection(doc, selectedItems);
            justification = { previous: null, changed: null };
            /* 実際にグループ化できたかを控える（例外で抜けたときに、選択していた既存のグループを解除しないため）*/
            didGroup = false;
            failure = null;
            try {
                /* 行揃えを先に変える。ポイント文字は行揃えを変えると文字が動くので、
                   変えたあとの境界で移動量を出さないと、そのぶんズレる */
                justification = btApplyJustification(selectedItems, options);
                bounds = btUnionBounds(selectedItems, options.previewBounds === true, options.glyphBounds === true);
                if (bounds === null) {
                    failure = "NOBOUNDS";
                } else {
                    artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
                    deltaX = 0;
                    deltaY = 0;
                    /* 斜めは2辺ぶんの移動量を、どちらも動かす前の境界から出して1回で動かす
                       先に片方だけ動かすと、もう一方のガイド探しが動いたあとの位置で走ってしまう */
                    for (i = 0; i < options.sides.length; i++) {
                        side = options.sides[i];
                        delta = btEdgeDelta(doc, artboardRect, bounds, side, options);
                        if (side === "left" || side === "right") { deltaX = delta; } else { deltaY = delta; }
                    }
                    if (needsGroup) {
                        app.executeMenuCommand("group");
                        didGroup = true;
                    }
                    btNudgeSelection(doc, deltaX, deltaY);
                }
            } catch (moveError) {
                failure = "ERR:" + moveError;
            } finally {
                if (didGroup) { app.executeMenuCommand("ungroup"); }
                /* 動かせなかったときは、先に変えた行揃えも元に戻す */
                if (failure !== null) { btRestoreJustification(selectedItems, justification.previous); }
            }
            if (failure !== null) { return failure; }
            if (justification.changed !== null) { return "MOVED:" + justification.changed; }
            return "MOVED";
        }

        function btEdgeDelta(doc, artboardRect, bounds, side, options) {
            var selectionEdge, artboardEdge, bleedEdge, targetEdge, delta;
            selectionEdge = btEdgeValue(bounds, side);
            artboardEdge = btEdgeValue(artboardRect, side);
            /* 進む向きにガイドがあればそこで止め、無ければアートボードの端まで動かす */
            targetEdge = btFindGuideSnapValue(doc, artboardRect, bounds, side, options.guideTolerance, options.minDeltaPt);
            if (targetEdge === null) {
                targetEdge = artboardEdge;
                if (options.bleedPt > 0) {
                    bleedEdge = btOutsetValue(artboardEdge, side, options.bleedPt);
                    /* 裁ち落としまで出ていればそこに留め、アートボードの端に乗っていれば次は裁ち落としへ */
                    if (Math.abs(bleedEdge - selectionEdge) < options.minDeltaPt) { return 0; }
                    if (Math.abs(artboardEdge - selectionEdge) < options.minDeltaPt) { targetEdge = bleedEdge; }
                }
            }
            delta = targetEdge - selectionEdge;
            if (Math.abs(delta) < options.minDeltaPt) { delta = 0; }
            return delta;
        }

        function btOutsetValue(edgeValue, side, bleedPt) {
            if (side === "left") { return edgeValue - bleedPt; }
            if (side === "right") { return edgeValue + bleedPt; }
            if (side === "top") { return edgeValue + bleedPt; }
            return edgeValue - bleedPt;
        }

        function btEdgeValue(bounds, side) {
            if (side === "left") { return bounds[0]; }
            if (side === "top") { return bounds[1]; }
            if (side === "right") { return bounds[2]; }
            return bounds[3];
        }

        function btFindGuideSnapValue(doc, artboardRect, bounds, side, tolerance, epsilon) {
            var selectionEdge, oppositeEdge, ahead, aheadDistance, behindNear, behindNearDistance, behindFar, behindFarDistance,
                guideItems, i, guideValues, j, guideValue, distance;
            selectionEdge = btEdgeValue(bounds, side);
            oppositeEdge = btEdgeValue(bounds, btOppositeSide(side));
            ahead = null;
            aheadDistance = null;
            behindNear = null;
            behindNearDistance = null;
            behindFar = null;
            behindFarDistance = null;
            guideItems = doc.pathItems;
            for (i = 0; i < guideItems.length; i++) {
                if (guideItems[i].guides !== true) { continue; }
                guideValues = btGuideEdgeValues(guideItems[i].geometricBounds, bounds, side, tolerance, epsilon);
                for (j = 0; j < guideValues.length; j++) {
                    guideValue = guideValues[j];
                    if (!btIsInsideArtboard(guideValue, artboardRect, side)) { continue; }
                    distance = Math.abs(guideValue - selectionEdge);
                    /* すでにその辺が乗っているガイドは行き先にしない（手で吸着させたときの微小なずれを吸収する） */
                    if (distance <= epsilon) { continue; }
                    if (btIsAhead(guideValue, selectionEdge, side, epsilon)) {
                        if (ahead === null || distance < aheadDistance) { ahead = guideValue; aheadDistance = distance; }
                    } else {
                        if (behindNear === null || distance < behindNearDistance) { behindNear = guideValue; behindNearDistance = distance; }
                        if (behindFar === null || distance > behindFarDistance) { behindFar = guideValue; behindFarDistance = distance; }
                    }
                }
            }
            if (ahead !== null) { return ahead; }
            /* 進む向きにガイドが無く、なおかつオブジェクトがガイド全体をまたいでいる（ガイドの間隔より大きい）ときだけ、
               戻る向きの最も近いガイドに合わせる。またいでいなければ null を返し、呼び出し側がアートボードの端を使う */
            if (behindNear !== null && btIsAhead(behindFar, oppositeEdge, side, epsilon)) { return behindNear; }
            return null;
        }

        function btOppositeSide(side) {
            if (side === "left") { return "right"; }
            if (side === "right") { return "left"; }
            if (side === "top") { return "bottom"; }
            return "top";
        }

        function btGuideEdgeValues(guideBounds, bounds, side, tolerance, epsilon) {
            var isVertical, isHorizontal;
            isVertical = Math.abs(guideBounds[2] - guideBounds[0]) <= tolerance;
            isHorizontal = Math.abs(guideBounds[1] - guideBounds[3]) <= tolerance;
            /* 線のガイドはその1本、長方形のガイド（マージンガイドなど）は向かい合う2辺を候補にする
               動かす向きと直交する方向で選択範囲と重ならないガイド＝そのまま動かしてもぶつからないガイドは対象にしない
               （ルーラーのガイドはアートボードの外まで伸びているので、いつでも重なる） */
            if (side === "left" || side === "right") {
                if (isHorizontal) { return []; }
                if (guideBounds[1] - bounds[3] <= epsilon || bounds[1] - guideBounds[3] <= epsilon) { return []; }
                return isVertical ? [guideBounds[0]] : [guideBounds[0], guideBounds[2]];
            }
            if (isVertical) { return []; }
            if (guideBounds[2] - bounds[0] <= epsilon || bounds[2] - guideBounds[0] <= epsilon) { return []; }
            return isHorizontal ? [guideBounds[1]] : [guideBounds[1], guideBounds[3]];
        }

        function btIsInsideArtboard(value, artboardRect, side) {
            if (side === "left" || side === "right") { return value >= artboardRect[0] && value <= artboardRect[2]; }
            return value <= artboardRect[1] && value >= artboardRect[3];
        }

        function btIsAhead(value, edge, side, epsilon) {
            if (side === "left") { return value < edge - epsilon; }
            if (side === "right") { return value > edge + epsilon; }
            if (side === "top") { return value > edge + epsilon; }
            return value < edge - epsilon;
        }

        function btUpdateMarginGuide(options) {
            var doc;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            btDrawMarginGuide(doc, options);
            return "OK";
        }

        function btDrawMarginGuide(doc, options) {
            var layer, previousLayerState, artboardRect, area, drawsFrame, divideLines;
            btRemoveGuidesByName(doc, options.guideName, options.guideLayerName);
            btRemoveGuidesByName(doc, options.divideName, options.guideLayerName);
            if (doc.artboards.length === 0) { return; }
            artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            /* マージンが大きすぎて内側が残らないときは null。マージンに依らないエッジのガイドは引ける */
            area = btMarginArea(artboardRect, options.guideMargins);
            drawsFrame = area !== null && options.showGuide === true && btHasMargin(options.guideMargins);
            divideLines = btDivideLines(area, artboardRect, options.divisions);
            /* 描くものが無いときは、ガイド用レイヤーを作らずに戻る */
            if (!drawsFrame && divideLines.length === 0) { return; }
            layer = btGetGuideLayer(doc, options.guideLayerName);
            previousLayerState = btUnlockLayer(layer);
            try {
                if (drawsFrame) { btAddGuideRectangle(layer, area, options.guideName); }
                btAddGuideLines(layer, divideLines, options.divideName);
            } finally {
                btRestoreLayer(layer, previousLayerState);
            }
        }

        function btMarginArea(artboardRect, margins) {
            var area = {
                left:   artboardRect[0] + margins.left,
                top:    artboardRect[1] - margins.top,
                right:  artboardRect[2] - margins.right,
                bottom: artboardRect[3] + margins.bottom
            };
            if (area.right <= area.left || area.top <= area.bottom) { return null; }
            return area;
        }

        function btAddGuideRectangle(layer, area, guideName) {
            btMakeGuide(layer.pathItems.rectangle(area.top, area.left, area.right - area.left, area.top - area.bottom), guideName);
        }

        function btAddGuideLines(layer, lines, guideName) {
            var guideLine, i;
            for (i = 0; i < lines.length; i++) {
                guideLine = layer.pathItems.add();
                guideLine.setEntirePath(lines[i]);
                btMakeGuide(guideLine, guideName);
            }
        }

        function btMakeGuide(pathItem, guideName) {
            var isGuide = false;
            try {
                /* 名前を先に付ける。ガイドにできずに消し損ねても、次の描き直しで名前を頼りに消せる */
                pathItem.name = guideName;
                pathItem.stroked = false;
                pathItem.filled = false;
                pathItem.guides = true;
                isGuide = (pathItem.guides === true);
            } catch (guideError) {}
            if (isGuide) { return; }
            /* ガイドにできなかったときは、描画時の塗り・線を持ったオブジェクトとして残さずに消す */
            try { pathItem.remove(); } catch (removeError) {}
        }

        function btDivideLines(area, artboardRect, divisions) {
            var lines, extension, span;
            lines = [];
            if (!divisions) { return lines; }
            /* 分割の線はアートボードの幅・高さいっぱいに引き、伸張のぶんだけその外へ伸ばす */
            extension = divisions.extension > 0 ? divisions.extension : 0;
            span = {
                left:   artboardRect[0] - extension,
                top:    artboardRect[1] + extension,
                right:  artboardRect[2] + extension,
                bottom: artboardRect[3] - extension
            };
            btPushRowLines(lines, area, divisions, span);
            btPushColumnLines(lines, area, divisions, span);
            btPushArtboardEdgeLines(lines, artboardRect, divisions, span);
            return lines;
        }

        function btPushArtboardEdgeLines(lines, artboardRect, divisions, span) {
            if (divisions.artboardEdge !== true) { return; }
            lines.push([[span.left, artboardRect[1]], [span.right, artboardRect[1]]]);
            lines.push([[span.left, artboardRect[3]], [span.right, artboardRect[3]]]);
            lines.push([[artboardRect[0], span.top], [artboardRect[0], span.bottom]]);
            lines.push([[artboardRect[2], span.top], [artboardRect[2], span.bottom]]);
        }

        function btPushRowLines(lines, area, divisions, span) {
            var usableHeight, cellHeight, lineY, i;
            if (area === null || !(divisions.rows > 1)) { return; }
            usableHeight = (area.top - area.bottom) - (divisions.rows - 1) * divisions.rowGutter;
            /* 行間が大きすぎて段の高さが残らないときは引かない */
            if (usableHeight <= 0) { return; }
            cellHeight = usableHeight / divisions.rows;
            lineY = area.top;
            /* 最後の段は area.bottom にちょうど着地するので、内側の区切りだけを引く */
            for (i = 0; i < divisions.rows - 1; i++) {
                lineY -= cellHeight;
                lines.push([[span.left, lineY], [span.right, lineY]]);
                /* 行間が0のときは同じ位置に重なるので引かない */
                if (divisions.rowGutter > 0) {
                    lineY -= divisions.rowGutter;
                    lines.push([[span.left, lineY], [span.right, lineY]]);
                }
            }
        }

        function btPushColumnLines(lines, area, divisions, span) {
            var usableWidth, cellWidth, lineX, i;
            if (area === null || !(divisions.columns > 1)) { return; }
            usableWidth = (area.right - area.left) - (divisions.columns - 1) * divisions.columnGutter;
            if (usableWidth <= 0) { return; }
            cellWidth = usableWidth / divisions.columns;
            lineX = area.left;
            for (i = 0; i < divisions.columns - 1; i++) {
                lineX += cellWidth;
                lines.push([[lineX, span.top], [lineX, span.bottom]]);
                if (divisions.columnGutter > 0) {
                    lineX += divisions.columnGutter;
                    lines.push([[lineX, span.top], [lineX, span.bottom]]);
                }
            }
        }

        function btHasMargin(margins) {
            return margins.top > 0 || margins.bottom > 0 || margins.left > 0 || margins.right > 0;
        }

        function btGetGuideLayer(doc, layerName) {
            var layer;
            try {
                layer = doc.layers.getByName(layerName);
            } catch (missingLayerError) {
                layer = doc.layers.add();
                layer.name = layerName;
            }
            return layer;
        }

        function btUnlockLayer(layer) {
            var previousLayerState;
            /* ロック・非表示のままでは書き換えられないので一時的に外し、戻せるよう控える */
            previousLayerState = { locked: layer.locked, visible: layer.visible };
            layer.locked = false;
            layer.visible = true;
            return previousLayerState;
        }

        function btRestoreLayer(layer, previousLayerState) {
            try {
                layer.visible = previousLayerState.visible;
                layer.locked = previousLayerState.locked;
            } catch (restoreLayerError) {}
        }

        function btRemoveGuidesByName(doc, guideName, layerName) {
            var layer, previousLayerState, items, i, item;
            try {
                layer = doc.layers.getByName(layerName);
            } catch (missingLayerError) {
                return;
            }
            previousLayerState = btUnlockLayer(layer);
            try {
                items = layer.pathItems;
                for (i = items.length - 1; i >= 0; i--) {
                    item = items[i];
                    if (item.name !== guideName) { continue; }
                    try {
                        item.locked = false;
                        item.hidden = false;
                        item.remove();
                    } catch (removeError) {}
                }
            } finally {
                btRestoreLayer(layer, previousLayerState);
            }
        }

        function btAlignMovedSelection(doc, options) {
            var boundsBefore, boundsAfter, i;
            boundsBefore = btUnionBounds(doc.selection, true, false);
            for (i = 0; i < options.alignCommands.length; i++) {
                app.executeMenuCommand(options.alignCommands[i]);
            }
            boundsAfter = btUnionBounds(doc.selection, true, false);
            return !btSameBounds(boundsBefore, boundsAfter);
        }

        function btRunAlignCommands(doc, options) {
            if (btAlignMovedSelection(doc, options)) { return true; }
            return btProbeAlignTarget(doc, options);
        }

        function btProbeAlignTarget(doc, options) {
            var moved;
            /* 整列で動かなかったときだけ、いったんずらして整列し直し、
               戻ってくるかどうかで整列先がアートボードかを見る */
            btNudgeSelection(doc, options.probeX, options.probeY);
            try {
                moved = btAlignMovedSelection(doc, options);
            } catch (probeError) {
                btNudgeSelection(doc, -options.probeX, -options.probeY);
                throw probeError;
            }
            if (!moved) { btNudgeSelection(doc, -options.probeX, -options.probeY); }
            return moved;
        }

        function btSameBounds(boundsA, boundsB) {
            var i;
            for (i = 0; i < 4; i++) {
                if (Math.abs(boundsA[i] - boundsB[i]) > 0.0001) { return false; }
            }
            return true;
        }

        function btApplyGlyphCorrection(doc, options) {
            var items, bounds, artboardRect, deltaX, deltaY;
            if (options.glyphBounds !== true) { return; }
            items = doc.selection;
            if (!(items instanceof Array) || items.length === 0) { return; }
            if (!btHasGlyphBoundsTarget(items)) { return; }
            if (doc.artboards.length === 0) { return; }
            bounds = btUnionBounds(items, options.previewBounds === true, true);
            if (bounds === null) { return; }
            artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            /* artboardRect は [左, 上, 右, 下] で、X は右へ・Y は下へ向かって内側になるため符号が逆になる */
            deltaX = btAlignDelta(bounds, artboardRect, options.modeX, options.marginPt, options.centerOffsetX, 0, 2, 1);
            deltaY = btAlignDelta(bounds, artboardRect, options.modeY, options.marginPt, options.centerOffsetY, 1, 3, -1);
            if (Math.abs(deltaX) < 0.001) { deltaX = 0; }
            if (Math.abs(deltaY) < 0.001) { deltaY = 0; }
            btNudgeSelection(doc, deltaX, deltaY);
        }

        function btAlignDelta(bounds, artboardRect, mode, marginPt, centerOffset, startIndex, endIndex, inwardSign) {
            if (mode === "start") { return (artboardRect[startIndex] + inwardSign * marginPt) - bounds[startIndex]; }
            if (mode === "end") { return (artboardRect[endIndex] - inwardSign * marginPt) - bounds[endIndex]; }
            /* 中央はアートボードの中心。Option＋クリックのときは centerOffset のぶんマージンの内側の中心へずらす */
            if (mode === "center") {
                return (artboardRect[startIndex] + artboardRect[endIndex]) / 2 + centerOffset - (bounds[startIndex] + bounds[endIndex]) / 2;
            }
            return 0;
        }

        function btAlignStops(options) {
            var stops;
            /* 内側から順に、マージン → アートボードの端 → 裁ち落とし。
               値はアートボードの辺からの内向きの量なので、裁ち落としは負になる */
            stops = [];
            if (options.marginPt > 0) { stops.push(options.marginPt); }
            stops.push(0);
            if (options.bleedPt > 0) { stops.push(-options.bleedPt); }
            return stops;
        }

        function btResolveStepMargin(doc, selectedItems, options) {
            var stops, bounds, artboardRect, i;
            stops = btAlignStops(options);
            if (stops.length < 2) { return stops[0]; }
            if (doc.artboards.length === 0) { return stops[0]; }
            bounds = btUnionBounds(selectedItems, options.previewBounds === true, options.glyphBounds === true);
            if (bounds === null) { return stops[0]; }
            artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            /* いま乗っている停止位置の次へ進める。いちばん外まで来ていればそこに留める */
            for (i = 0; i < stops.length; i++) {
                if (btIsAtAlignTarget(bounds, artboardRect, options, stops[i])) {
                    return (i + 1 < stops.length) ? stops[i + 1] : stops[i];
                }
            }
            return stops[0];
        }

        function btIsAtAlignTarget(bounds, artboardRect, options, marginPt) {
            if (Math.abs(btAlignDelta(bounds, artboardRect, options.modeX, marginPt, options.centerOffsetX, 0, 2, 1)) > options.minDeltaPt) { return false; }
            if (Math.abs(btAlignDelta(bounds, artboardRect, options.modeY, marginPt, options.centerOffsetY, 1, 3, -1)) > options.minDeltaPt) { return false; }
            return true;
        }

        function btHasGlyphBoundsTarget(items) {
            var i, typeName;
            for (i = 0; i < items.length; i++) {
                typeName = items[i].typename;
                if (typeName === "TextFrame" || typeName === "SymbolItem") { return true; }
            }
            return false;
        }

        function btUnionBounds(items, usePreviewBounds, useGlyphBounds) {
            var bounds, itemBounds, i;
            bounds = null;
            for (i = 0; i < items.length; i++) {
                /* 削除済みの参照が混ざることがあるので、測れないものは飛ばす */
                try {
                    itemBounds = btGetMeasureBounds(items[i], usePreviewBounds, useGlyphBounds);
                } catch (boundsError) {
                    continue;
                }
                if (!itemBounds) { continue; }
                if (bounds === null) {
                    bounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                } else {
                    if (itemBounds[0] < bounds[0]) { bounds[0] = itemBounds[0]; }
                    if (itemBounds[1] > bounds[1]) { bounds[1] = itemBounds[1]; }
                    if (itemBounds[2] > bounds[2]) { bounds[2] = itemBounds[2]; }
                    if (itemBounds[3] < bounds[3]) { bounds[3] = itemBounds[3]; }
                }
            }
            return bounds;
        }

        function btGetMeasureBounds(item, usePreviewBounds, useGlyphBounds) {
            var bounds;
            bounds = null;
            if (useGlyphBounds === true) {
                if (item.typename === "TextFrame") { bounds = btGetOutlineBounds(item, usePreviewBounds); }
                else if (item.typename === "SymbolItem") { bounds = btGetSymbolGlyphBounds(item, usePreviewBounds); }
                if (bounds !== null) { return bounds; }
            }
            return btGetPlainBounds(item, usePreviewBounds);
        }

        function btGetPlainBounds(item, usePreviewBounds) {
            var bounds;
            /* クリップグループの境界はマスクを無視した中身全体を返すため、マスクのパスで測り直す */
            if (item.typename === "GroupItem" && item.clipped === true) {
                bounds = btGetClipPathBounds(item, usePreviewBounds);
                if (bounds !== null) { return bounds; }
            }
            return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
        }

        function btGetClipPathBounds(groupItem, usePreviewBounds) {
            var clipPath;
            clipPath = btFindClipPath(groupItem);
            if (clipPath === null) { return null; }
            return usePreviewBounds ? clipPath.visibleBounds : clipPath.geometricBounds;
        }

        function btFindClipPath(groupItem) {
            var child, typeName, i;
            for (i = 0; i < groupItem.pageItems.length; i++) {
                child = groupItem.pageItems[i];
                typeName = child.typename;
                if (typeName === "PathItem" && child.clipping === true) { return child; }
                /* 複合パスのマスクは、束ねているパスの clipping にだけ立つ */
                if (typeName === "CompoundPathItem" && child.pathItems.length > 0 && child.pathItems[0].clipping === true) { return child; }
            }
            return null;
        }

        function btGetOutlineBounds(textFrame, usePreviewBounds) {
            var duplicated, outlined, bounds;
            duplicated = null;
            outlined = null;
            bounds = null;
            try {
                duplicated = textFrame.duplicate();
                outlined = duplicated.createOutline();
                bounds = usePreviewBounds ? outlined.visibleBounds : outlined.geometricBounds;
            } catch (outlineError) {
                bounds = null;
            } finally {
                btSafeRemove(outlined);
                btSafeRemove(duplicated);
            }
            return bounds;
        }

        function btGetSymbolGlyphBounds(symbolItem, usePreviewBounds) {
            var doc, layer, selectionBefore, itemsBefore, layersBefore, brokenItems, hasText, bounds, i;
            layer = symbolItem.layer;
            if (!layer) { return null; }
            doc = app.activeDocument;
            selectionBefore = doc.selection;
            itemsBefore = btCollectionToArray(layer.pageItems);
            layersBefore = btCollectionToArray(layer.layers);
            bounds = null;
            try {
                /* 複製をレイヤー直下に置いてからリンクを解除する
                   グループの中で解除すると生成物の行き先が読めないため、いったん外へ出す */
                symbolItem.duplicate(layer, ElementPlacement.PLACEATBEGINNING).breakLink();
                /* breakLink は複製そのものを別のオブジェクトに置き換えるので、参照ではなく
                   レイヤー直下の差分で生成物を拾う（静的シンボルはサブレイヤーを作る）*/
                brokenItems = btCollectBrokenItems(layer, itemsBefore, layersBefore);
                hasText = false;
                for (i = 0; i < brokenItems.length; i++) {
                    if (btOutlineTextFramesIn(brokenItems[i])) { hasText = true; }
                }
                /* 文字が無ければシンボルの通常の境界と変わらないので、呼び出し元の既定に任せる */
                if (hasText) {
                    /* アウトライン化でテキストは別のオブジェクトに置き換わるため、拾い直してから測る */
                    bounds = btUnionBounds(btCollectBrokenItems(layer, itemsBefore, layersBefore), usePreviewBounds, false);
                }
            } catch (symbolError) {
                bounds = null;
            } finally {
                btRemoveBrokenItems(layer, itemsBefore, layersBefore);
                /* breakLink は生成物を選択状態にするため、元の選択に戻す */
                if (selectionBefore instanceof Array) {
                    try { doc.selection = selectionBefore; } catch (selectionError) {}
                }
            }
            return bounds;
        }

        function btCollectionToArray(collection) {
            var items, i;
            items = [];
            for (i = 0; i < collection.length; i++) { items.push(collection[i]); }
            return items;
        }

        function btCollectNewEntries(current, existing) {
            var newEntries, isExisting, i, j;
            newEntries = [];
            for (i = 0; i < current.length; i++) {
                isExisting = false;
                for (j = 0; j < existing.length; j++) {
                    if (existing[j] === current[i]) { isExisting = true; break; }
                }
                if (!isExisting) { newEntries.push(current[i]); }
            }
            return newEntries;
        }

        function btCollectBrokenItems(layer, itemsBefore, layersBefore) {
            var brokenItems, newLayers, i, j;
            brokenItems = btCollectNewEntries(layer.pageItems, itemsBefore);
            newLayers = btCollectNewEntries(layer.layers, layersBefore);
            for (i = 0; i < newLayers.length; i++) {
                for (j = 0; j < newLayers[i].pageItems.length; j++) { brokenItems.push(newLayers[i].pageItems[j]); }
            }
            return brokenItems;
        }

        function btRemoveBrokenItems(layer, itemsBefore, layersBefore) {
            var newItems, newLayers, i;
            try {
                newItems = btCollectNewEntries(layer.pageItems, itemsBefore);
                newLayers = btCollectNewEntries(layer.layers, layersBefore);
            } catch (collectError) {
                return;
            }
            for (i = 0; i < newItems.length; i++) { btSafeRemove(newItems[i]); }
            for (i = 0; i < newLayers.length; i++) { btSafeRemove(newLayers[i]); }
        }

        function btOutlineTextFramesIn(item) {
            var textFrames, outlined, i;
            textFrames = [];
            btCollectTextFrames(item, textFrames);
            /* 先に集めてからアウトライン化する。createOutline は元のテキストを置き換えるので、
               pageItems をたどりながらだと取りこぼす */
            outlined = false;
            for (i = 0; i < textFrames.length; i++) {
                try {
                    textFrames[i].createOutline();
                    outlined = true;
                } catch (outlineError) {}
            }
            return outlined;
        }

        function btCollectTextFrames(item, textFrames) {
            var i;
            if (!item) { return; }
            if (item.typename === "TextFrame") { textFrames.push(item); return; }
            if (item.typename !== "GroupItem") { return; }
            for (i = 0; i < item.pageItems.length; i++) { btCollectTextFrames(item.pageItems[i], textFrames); }
        }

        function btSafeRemove(item) {
            try {
                if (item) { item.remove(); }
            } catch (removeError) {}
        }

        function btGetSelectionKind() {
            var doc, selectedItems, i;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            selectedItems = doc.selection;
            if (selectedItems && !(selectedItems instanceof Array)) { return "TEXT"; }
            if (!(selectedItems instanceof Array) || selectedItems.length === 0) { return "NONE"; }
            for (i = 0; i < selectedItems.length; i++) {
                if (selectedItems[i].typename !== "TextFrame") { return "OTHER"; }
            }
            return "TEXT";
        }

        function btGetPaletteState() {
            var doc, count, artboardCount;
            count = 0;
            artboardCount = 0;
            if (app.documents.length > 0) {
                doc = app.activeDocument;
                if (doc.selection instanceof Array) { count = doc.selection.length; }
                else if (doc.selection) { count = 1; }
                artboardCount = doc.artboards.length;
            }
            return btGetSelectionKind() + "|" + app.preferences.getIntegerPreference("rulerType") +
                "|" + count + "|" + artboardCount;
        }

        function btNudgeSelection(doc, deltaX, deltaY) {
            var items, i;
            if (deltaX === 0 && deltaY === 0) { return; }
            items = doc.selection;
            if (!(items instanceof Array)) { return; }
            for (i = 0; i < items.length; i++) {
                items[i].translate(deltaX, deltaY);
            }
        }

        function btReadPreferences() {
            return {
                previewBounds: app.preferences.getBooleanPreference("includeStrokeInBounds") === true,
                pointText: app.preferences.getBooleanPreference("EnableActualPointTextSpaceAlign") === true,
                areaText: app.preferences.getBooleanPreference("EnableActualAreaTextSpaceAlign") === true
            };
        }

        function btGetPreferenceFlags() {
            var preferences;
            preferences = btReadPreferences();
            /* 字形の境界はポイント文字・エリア内文字のどちらかがONならON扱いにする */
            return (preferences.previewBounds ? "1" : "0") + "|" + ((preferences.pointText || preferences.areaText) ? "1" : "0");
        }

        function btWritePreferences(previewBounds, pointText, areaText) {
            app.preferences.setBooleanPreference("includeStrokeInBounds", previewBounds === true);
            app.preferences.setBooleanPreference("EnableActualPointTextSpaceAlign", pointText === true);
            app.preferences.setBooleanPreference("EnableActualAreaTextSpaceAlign", areaText === true);
        }

        function btPromoteTextRange(doc, textRange) {
            var storyFrames, targetFrames, i;
            storyFrames = textRange.story.textFrames;
            targetFrames = [];
            for (i = 0; i < storyFrames.length; i++) { targetFrames.push(storyFrames[i]); }
            app.executeMenuCommand("deselectall");
            for (i = 0; i < targetFrames.length; i++) { targetFrames[i].selected = true; }
            return doc.selection;
        }

        function btGetLayerKey(layer) {
            var keyParts, node;
            keyParts = [];
            node = layer;
            while (node && node.typename === "Layer") {
                keyParts.push(node.zOrderPosition + ":" + node.name);
                node = node.parent;
            }
            return keyParts.join("/");
        }

        function btSpansMultipleLayers(selectedItems) {
            var firstKey, i;
            firstKey = btGetLayerKey(selectedItems[0].layer);
            for (i = 1; i < selectedItems.length; i++) {
                if (btGetLayerKey(selectedItems[i].layer) !== firstKey) { return true; }
            }
            return false;
        }

        function btGetOverlapArea(boundsA, boundsB) {
            var overlapWidth, overlapHeight;
            overlapWidth = Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]);
            overlapHeight = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
            if (overlapWidth <= 0 || overlapHeight <= 0) { return 0; }
            return overlapWidth * overlapHeight;
        }

        function btFindOverlappingArtboardIndex(doc, selectionBounds, currentIndex) {
            var searchOrder, i, j, bestIndex, bestArea, area;
            searchOrder = [currentIndex];
            for (i = 0; i < doc.artboards.length; i++) {
                if (i !== currentIndex) { searchOrder.push(i); }
            }
            bestIndex = -1;
            bestArea = 0;
            for (j = 0; j < searchOrder.length; j++) {
                area = btGetOverlapArea(selectionBounds, doc.artboards[searchOrder[j]].artboardRect);
                if (area > bestArea) { bestArea = area; bestIndex = searchOrder[j]; }
            }
            return bestIndex;
        }

        function btFindNearestArtboardIndex(doc, selectionBounds) {
            var centerX, centerY, nearestIndex, nearestDistance, i, artboardRect, offsetX, offsetY, distance;
            centerX = (selectionBounds[0] + selectionBounds[2]) / 2;
            centerY = (selectionBounds[1] + selectionBounds[3]) / 2;
            nearestIndex = 0;
            nearestDistance = null;
            for (i = 0; i < doc.artboards.length; i++) {
                artboardRect = doc.artboards[i].artboardRect;
                offsetX = centerX - (artboardRect[0] + artboardRect[2]) / 2;
                offsetY = centerY - (artboardRect[1] + artboardRect[3]) / 2;
                distance = offsetX * offsetX + offsetY * offsetY;
                if (nearestDistance === null || distance < nearestDistance) {
                    nearestDistance = distance;
                    nearestIndex = i;
                }
            }
            return nearestIndex;
        }

        function btActivateArtboardForSelection(doc, selectedItems) {
            var selectionBounds, currentIndex, targetIndex;
            if (doc.artboards.length < 2) { return; }
            selectionBounds = btUnionBounds(selectedItems, true, false);
            currentIndex = doc.artboards.getActiveArtboardIndex();
            targetIndex = btFindOverlappingArtboardIndex(doc, selectionBounds, currentIndex);
            if (targetIndex < 0) { targetIndex = btFindNearestArtboardIndex(doc, selectionBounds); }
            if (targetIndex !== currentIndex) { doc.artboards.setActiveArtboardIndex(targetIndex); }
        }

        function btIsSingleLineTextFrame(selectedItems) {
            if (selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") { return false; }
            return selectedItems[0].lines.length === 1;
        }

        function btJustificationByName(justificationName) {
            if (justificationName === "LEFT") { return Justification.LEFT; }
            if (justificationName === "CENTER") { return Justification.CENTER; }
            if (justificationName === "RIGHT") { return Justification.RIGHT; }
            return null;
        }

        function btSetJustification(textFrame, justificationName) {
            var justification, previousJustification;
            justification = btJustificationByName(justificationName);
            if (justification === null) { return null; }
            previousJustification = textFrame.textRange.paragraphAttributes.justification;
            textFrame.textRange.paragraphAttributes.justification = justification;
            return previousJustification;
        }

        /* 送信するワーカー関数の一覧（追加したらここにも必ず登録する）/ Every worker function shipped to the main engine */
        var WORKER_FUNCS = [
            btAlignSelection, btResolveSelection, btSelectItems, btRunPerArtboard, btAlignItems,
            btGroupItemsByArtboard, btCopyOptions, btApplyJustification, btFinishAlign,
            btMoveItems, btAlignMovedSelection, btRunAlignCommands, btProbeAlignTarget, btSameBounds,
            btMoveToEdgeOrGuide, btRestoreJustification, btEdgeDelta, btOutsetValue, btEdgeValue, btOppositeSide, btFindGuideSnapValue, btGuideEdgeValues, btIsInsideArtboard, btIsAhead,
            btUpdateMarginGuide, btDrawMarginGuide, btMarginArea, btAddGuideRectangle, btAddGuideLines,
            btDivideLines, btPushRowLines, btPushColumnLines, btPushArtboardEdgeLines, btMakeGuide,
            btHasMargin, btGetGuideLayer, btUnlockLayer, btRestoreLayer, btRemoveGuidesByName,
            btApplyGlyphCorrection, btAlignDelta, btAlignStops, btResolveStepMargin, btIsAtAlignTarget, btHasGlyphBoundsTarget,
            btUnionBounds, btGetMeasureBounds, btGetPlainBounds, btGetClipPathBounds, btFindClipPath,
            btGetOutlineBounds, btSafeRemove,
            btGetSymbolGlyphBounds, btCollectionToArray, btCollectNewEntries, btCollectBrokenItems,
            btRemoveBrokenItems, btOutlineTextFramesIn, btCollectTextFrames,
            btGetPaletteState, btGetSelectionKind, btJustificationByName,
            btNudgeSelection, btReadPreferences, btGetPreferenceFlags, btWritePreferences, btPromoteTextRange,
            btGetLayerKey, btSpansMultipleLayers,
            btGetOverlapArea,
            btFindOverlappingArtboardIndex, btFindNearestArtboardIndex, btActivateArtboardForSelection,
            btIsSingleLineTextFrame, btSetJustification
        ];

        // =========================================
        // メインエンジンへの委譲 / Delegating to the main engine
        // =========================================
        /* 常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理は毎回メインエンジンへ送る
           A persistent palette loses its DOM connection while shown, so every DOM touch is delegated */

        /* 実行中フラグ（連打による多重実行を防ぐ）/ Guard against double execution from rapid clicks */
        var isBusy = false;

        /* ワーカーの戻り値マーカーと status ラベルの対応（「OK:CENTER」は showWorkerResult で組み立てる）
           Worker markers mapped to status labels; "OK:CENTER" is composed in showWorkerResult */
        var STATUS_BY_MARKER = {
            OK:         "done",
            MOVED:      "moved",
            NOBOUNDS:   "noBounds",
            NODOC:      "noDocument",
            NOSEL:      "noSelection",
            MULTILAYER: "multipleLayers",
            NOTARGET:   "alignTarget"
        };

        /**
         * 再入防止つきで処理を実行する（連打による多重実行を防ぐ）
         * @param {function} action - 実行する処理
         * @returns {boolean} 実行したら true（実行中で見送ったときは false）
         */
        function runExclusive(action) {
            if (isBusy) { return false; }
            isBusy = true;
            try {
                action();
            } finally {
                isBusy = false;
            }
            return true;
        }

        /**
         * 関数のソースから宣言行〜閉じ括弧行だけを切り出す
         * ExtendScript の toString() は改行を CR で返し、前後のコメント断片を閉じ「*」「/」を落として
         * 巻き込むことがあるため、行区切りを LF に正規化したうえで関数本体だけを取り出す
         * @param {function} targetFunction - 文字列化する関数
         * @returns {string} 関数宣言だけのソース文字列
         */
        function sliceFunctionSource(targetFunction) {
            var lines = String(targetFunction).replace(/\r\n?/g, "\n").split("\n");
            var firstIndex = -1;
            var lastIndex = -1;
            for (var i = 0; i < lines.length; i++) {
                if (firstIndex < 0 && /^\s*function\s/.test(lines[i])) { firstIndex = i; }
                if (firstIndex >= 0 && /^\s*\}[;\s]*$/.test(lines[i])) { lastIndex = i; }
            }
            if (firstIndex < 0) { return String(targetFunction); }
            if (lastIndex < firstIndex) {
                /* 1行で書かれた関数は、その行だけを取り出す / A function written on one line: keep just that line */
                return /\}[;\s]*$/.test(lines[firstIndex]) ? lines[firstIndex] : lines.slice(firstIndex).join("\n");
            }
            return lines.slice(firstIndex, lastIndex + 1).join("\n");
        }

        /* 連結済みのワーカーソースと、その刻印（1回だけ組み立てて使い回す）
           The assembled worker source and its stamp, built once and reused */
        var workerSourceCache = null;
        var workerStampCache = null;

        /**
         * ソースの取り違えを防ぐ刻印を作る（内容が1文字でも変われば別の値になる）
         * @param {string} source - 対象のソース
         * @returns {string} 刻印
         */
        function buildWorkerStamp(source) {
            var checksum = 0;
            for (var i = 0; i < source.length; i++) {
                checksum = (checksum * 31 + source.charCodeAt(i)) % 2147483647;
            }
            return SCRIPT_VERSION + "-" + source.length + "-" + checksum;
        }

        /**
         * ワーカー関数の定義をひとつの文字列にまとめる（2回目以降はキャッシュを返す）
         * @returns {string} 連結したワーカー関数のソース
         */
        function buildWorkerSource() {
            if (workerSourceCache !== null) { return workerSourceCache; }
            var sources = [];
            for (var i = 0; i < WORKER_FUNCS.length; i++) {
                sources.push(sliceFunctionSource(WORKER_FUNCS[i]));
            }
            workerSourceCache = sources.join("\n");
            workerStampCache = buildWorkerStamp(workerSourceCache);
            return workerSourceCache;
        }

        /* メインエンジンに送り込んだはずのワーカー定義の刻印（送り直しの要否を判断する）
           The stamp of the worker source believed to be loaded in the main engine */
        var loadedWorkerStamp = $.global[WORKER_STAMP_KEY] || null;

        /**
         * 本文をメインエンジンで同期実行し、結果を受け取る
         * @param {string} body - メインエンジンで評価する本文
         * @returns {string} 評価結果（応答がなければ null）
         */
        function sendToMainEngine(body) {
            /* 同期送信の結果は holder 経由で受け取る / The synchronous send hands its result back through holder */
            var holder = { result: null };
            var bridge = new BridgeTalk();
            bridge.target = "illustrator";
            bridge.body = body;
            bridge.onResult = function(message) { holder.result = String(message.body); };
            bridge.onError = function(message) { holder.result = "ERR:" + String(message.body); };
            bridge.send(WORKER_TIMEOUT);
            return holder.result;
        }

        /**
         * ワーカー定義ごと送る本文を組み立てる
         * 定義をメインエンジンのグローバルに残したいので、eval はトップレベルで実行する
         * （関数の中で eval するとその関数のローカルになり、次の呼び出しから見えない）
         * @param {string} workerSource - 連結したワーカー関数のソース
         * @param {string} stamp - そのソースの刻印
         * @param {string} functionCall - 評価する呼び出し式
         * @returns {string} 送信する本文
         */
        function buildFullBody(workerSource, stamp, functionCall) {
            /* バックスラッシュ・多バイト文字・改行が途中で壊れないよう、ソースはURIエンコードして送る
               URI-encode the source so backslashes, multi-byte characters and newlines survive the trip */
            /* 全体をひとつの式にして、最後に評価される呼び出しの値がそのまま結果として返るようにする
               One comma expression, so the value of the trailing call is what comes back */
            return "eval(decodeURIComponent(\"" + encodeURIComponent(workerSource) + "\")), " +
                "$.global." + WORKER_STAMP_KEY + " = \"" + stamp + "\", " +
                functionCall;
        }

        /**
         * 呼び出し式だけを送る本文を組み立てる（定義が残っていなければ "RELOAD" を返させる）
         * @param {string} stamp - 送り込んだはずのソースの刻印
         * @param {string} functionCall - 評価する呼び出し式
         * @returns {string} 送信する本文
         */
        function buildCallOnlyBody(stamp, functionCall) {
            return "(function() {" +
                "if ($.global." + WORKER_STAMP_KEY + " !== \"" + stamp + "\") { return \"RELOAD\"; }" +
                "if (typeof btAlignSelection !== \"function\") { return \"RELOAD\"; }" +
                "return " + functionCall +
                "})();";
        }

        /**
         * ワーカー関数の呼び出しをメインエンジンで同期実行し、戻り値のマーカーを受け取る
         * 定義はメインエンジンのグローバルに残るので、2回目以降は呼び出し式だけを送る
         * （毎回ソースを丸ごと送ると、選択を問い合わせるたびに10KB近いやり取りになる）
         * @param {string} functionCall - メインエンジンで評価する呼び出し式
         * @returns {string} ワーカーが返したマーカー（応答がなければ null）
         */
        function runWorker(functionCall) {
            var workerSource = buildWorkerSource();
            var stamp = workerStampCache;
            try {
                var workerResult = null;
                var needsFullBody = (loadedWorkerStamp !== stamp);
                if (!needsFullBody) {
                    workerResult = sendToMainEngine(buildCallOnlyBody(stamp, functionCall));
                    /* 定義が消えていた（別のエンジンが再起動した）ときは送り直す
                       Ship the source again when the definitions are gone */
                    needsFullBody = (workerResult === "RELOAD");
                }
                if (needsFullBody) {
                    workerResult = sendToMainEngine(buildFullBody(workerSource, stamp, functionCall));
                    if (workerResult !== null) {
                        loadedWorkerStamp = stamp;
                        $.global[WORKER_STAMP_KEY] = stamp;
                    }
                }
                return workerResult;
            } catch (bridgeError) {
                /* BridgeTalk が使えない環境では、このエンジンで直接実行する
                   Fallback: run in this engine when BridgeTalk is unavailable */
                try {
                    return String(eval(workerSource + "\n" + functionCall));
                } catch (evalError) {
                    return "ERR:" + evalError;
                }
            }
        }

        /**
         * 数値の入力欄を読んで pt に換算する（数値以外と負数は0に丸め、欄の表示もそろえる）
         * @param {EditText} field - 読み取る入力欄
         * @param {number} pointsPerUnit - 1単位あたりの pt
         * @returns {number} 入力値（pt）
         */
        function readFieldPt(field, pointsPerUnit) {
            if (field === null) { return 0; }
            var fieldValue = Number(field.text);
            if (isNaN(fieldValue) || fieldValue < 0) { fieldValue = 0; }
            /* 手入力が丸められたときは欄の表示も実際に使う値にそろえる / Show the value actually used */
            if (String(fieldValue) !== field.text) { field.text = fieldValue; }
            return fieldValue * pointsPerUnit;
        }

        /**
         * マージン4欄の表示中の文字列を辺ごとに読む（控え用）
         * @returns {object} { top: string, bottom: string, left: string, right: string }
         */
        function readMarginTexts() {
            var texts, side;
            texts = {};
            for (side in marginFields) {
                if (!marginFields.hasOwnProperty(side)) { continue; }
                texts[side] = String(marginFields[side].text);
            }
            return texts;
        }

        /**
         * 分割ガイドの数値欄を、入力されたままの文字列で読む（控え用）
         * @returns {object} DIVIDE_VALUE_KEYS をキーにした文字列の組
         */
        function readDivideValueTexts() {
            var texts, valueKey, i;
            texts = {};
            for (i = 0; i < DIVIDE_VALUE_KEYS.length; i++) {
                valueKey = DIVIDE_VALUE_KEYS[i];
                texts[valueKey] = (divideFields[valueKey]) ?
                    String(divideFields[valueKey].text) : defaultDivideValue(valueKey);
            }
            return texts;
        }

        /**
         * 分割数の欄を読む（読めない値と1未満は1＝分割なし、上限は DIVIDE_COUNT_MAX）
         * @param {string} valueKey - "rows" または "columns"
         * @returns {number} 分割数
         */
        function readDivideCount(valueKey) {
            var count = Math.floor(Number(divideFields[valueKey] ? divideFields[valueKey].text : ""));
            if (isNaN(count) || count < 1) { return 1; }
            return (count > DIVIDE_COUNT_MAX) ? DIVIDE_COUNT_MAX : count;
        }

        /**
         * 分割ガイドの設定を読んでワーカーに渡す形にする
         * ［なし］は分割なし、［十字］は縦横2等分（伸張はどちらにも効く）
         * @returns {object} { rows: number, columns: number, rowGutter: number, columnGutter: number, extension: number }（長さは pt）
         */
        function readDivisions() {
            var mode, divisions;
            mode = readDivideMode();
            divisions = {
                rows:         1,
                columns:      1,
                rowGutter:    0,
                columnGutter: 0,
                extension:    readFieldPt(divideFields.extension || null, currentUnitInfo.points),
                artboardEdge: readArtboardEdge()
            };
            if (mode === "cross") {
                divisions.rows    = CROSS_DIVIDE_COUNTS.rows;
                divisions.columns = CROSS_DIVIDE_COUNTS.columns;
                return divisions;
            }
            if (mode === "custom") {
                divisions.rows         = readDivideCount("rows");
                divisions.columns      = readDivideCount("columns");
                divisions.rowGutter    = readFieldPt(divideFields.rowGutter || null, currentUnitInfo.points);
                divisions.columnGutter = readFieldPt(divideFields.columnGutter || null, currentUnitInfo.points);
            }
            return divisions;
        }

        /**
         * マージン4欄を読んで pt に換算する
         * @returns {object} { top: number, bottom: number, left: number, right: number }（pt）
         */
        function readMarginsPt() {
            var margins, side, i;
            margins = {};
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                margins[side] = readFieldPt(marginFields[side] || null, currentUnitInfo.points);
            }
            return margins;
        }

        /**
         * その整列が使うマージンを、寄せる辺から1つ選ぶ（中央揃えは使わないので0）
         * @param {object} spec - readAlignSpec() の戻り値
         * @param {object} margins - readMarginsPt() の戻り値
         * @returns {number} マージン（pt）
         */
        function marginForAlign(spec, margins) {
            if (spec.modeX === "start") { return margins.left; }
            if (spec.modeX === "end") { return margins.right; }
            if (spec.modeY === "start") { return margins.top; }
            if (spec.modeY === "end") { return margins.bottom; }
            return 0;
        }

        /**
         * 裁ち落としの距離を pt で読む（［裁ち落としに整列］がOFFなら0）
         * @returns {number} 裁ち落とし（pt）
         */
        function readBleedPt() {
            if (alignToBleedCheckbox === null || alignToBleedCheckbox.value !== true) { return 0; }
            return readFieldPt(bleedField, BLEED_UNIT_POINTS);
        }

        /**
         * ボタン定義から、そのクリックに必要な値をまとめて求める
         * 実行するメニューコマンド、整列後にマージンぶん動かす向き（中央揃えは0）、整列先の判定に使う仮移動量、
         * 軸ごとの寄せ先（字形の境界での補正が目標位置の計算に使う）、合わせる行揃えを、ALIGN_COMMANDS の1周で得る
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {object} { commands: string[], offsetX: number, offsetY: number, probeX: number, probeY: number, modeX: string, modeY: string, justification: string }
         */
        function readAlignSpec(buttonDef) {
            var spec = {
                commands: [],
                offsetX: 0,
                offsetY: 0,
                probeX: 0,
                probeY: 0,
                modeX: null,
                modeY: null,
                justification: null
            };
            for (var i = 0; i < buttonDef.alignKeys.length; i++) {
                var alignCommand = ALIGN_COMMANDS[buttonDef.alignKeys[i]];
                spec.commands.push(alignCommand.command);
                spec.offsetX += alignCommand.offsetX;
                spec.offsetY += alignCommand.offsetY;
                /* 仮移動は端揃えなら内側へ、中央揃えは向きがないので＋方向へ（整列が空振りしたときだけ使う）
                   The probe moves inward for an edge align; a center align has no direction, so it uses the plus side */
                if (alignCommand.axis === "x") {
                    spec.probeX = (alignCommand.offsetX !== 0 ? alignCommand.offsetX : 1) * ALIGN_PROBE_PT;
                    spec.modeX = alignCommand.mode;
                } else {
                    spec.probeY = (alignCommand.offsetY !== 0 ? alignCommand.offsetY : 1) * ALIGN_PROBE_PT;
                    spec.modeY = alignCommand.mode;
                }
                /* 行揃えは水平方向の整列にだけ付いているので、最初に見つかったものを使う
                   Only horizontal aligns carry a justification, so the first one found wins */
                if (spec.justification === null && alignCommand.justification) {
                    spec.justification = alignCommand.justification;
                }
            }
            return spec;
        }

        /**
         * 中央に寄せる軸を持つ整列か判定する
         * @param {object} spec - readAlignSpec() の戻り値
         * @returns {boolean} 片方でも中央に寄せるなら true
         */
        function isCenterAlign(spec) {
            return spec.modeX === "center" || spec.modeY === "center";
        }

        /**
         * そのボタンの Option＋クリックの説明を選ぶ
         * 中央揃えはマージンの内側の中央へ、端に寄せる整列はマージンを無視して字形の境界に寄せる
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {string} tooltip のキー
         */
        function optionTooltipKey(buttonDef) {
            var spec = readAlignSpec(buttonDef);
            if (isCenterAlign(spec)) { return "optionMarginCenter"; }
            /* マージンの影響を受けるのは上下左右に寄せるボタンだけ / Only the edge alignments use a margin */
            return (spec.offsetX !== 0 || spec.offsetY !== 0) ? "optionNoMargin" : "optionGlyphBounds";
        }

        /**
         * Option＋クリックでマージンの内側の中央へ寄せるための、アートボードの中央からのずれを求める
         * 中央に寄せる軸だけ、向かい合うマージンの差の半分だけ内側へずらす
         * @param {object} spec - readAlignSpec() の戻り値
         * @param {object} margins - readMarginsPt() の戻り値
         * @returns {object} { x: number, y: number }（pt）
         */
        function centerOffsetInMargin(spec, margins) {
            return {
                x: (spec.modeX === "center") ? (margins.left - margins.right) / 2 : 0,
                y: (spec.modeY === "center") ? (margins.bottom - margins.top) / 2 : 0
            };
        }

        /**
         * ガイドの作成に必要な値を組み立てる
         * マージンは Option＋クリックの影響を受けない（ガイドは入力欄の値をそのまま表す）
         * @returns {object} ワーカーへ渡すガイドのオプション
         */
        function buildGuideOptions() {
            return {
                showGuide:      showGuideCheckbox !== null && showGuideCheckbox.value === true,
                guideMargins:   readMarginsPt(),
                divisions:      readDivisions(),
                guideName:      GUIDE_NAME,
                divideName:     DIVIDE_GUIDE_NAME,
                guideLayerName: GUIDE_LAYER_NAME
            };
        }

        /**
         * このスクリプトが作ったガイド（マージン・分割の両方）を削除する
         * パレットを閉じるときに呼ぶため、コントロールを参照せず値を直接組み立てる
         * （閉じる処理は何があっても止めないよう、失敗は握りつぶす）
         * @returns {void}
         */
        function removeMarginGuide() {
            try {
                var options = {
                    showGuide:      false,
                    guideMargins:   { top: 0, bottom: 0, left: 0, right: 0 },
                    divisions:      { rows: 1, columns: 1, rowGutter: 0, columnGutter: 0, extension: 0, artboardEdge: false },
                    guideName:      GUIDE_NAME,
                    divideName:     DIVIDE_GUIDE_NAME,
                    guideLayerName: GUIDE_LAYER_NAME
                };
                runWorker("btUpdateMarginGuide(" + options.toSource() + ");");
            } catch (removeGuideError) {}
        }

        /**
         * マージンのガイドを作り直す（チェックがOFFなら消すだけ）
         * 結果は状況表示に出さず、エラーのときだけ知らせる
         * @returns {void}
         */
        function refreshMarginGuide() {
            if (showGuideCheckbox === null) { return; }
            var workerResult = runWorker("btUpdateMarginGuide(" + buildGuideOptions().toSource() + ");");
            if (workerResult !== null && workerResult.indexOf("ERR:") === 0) { showWorkerResult(workerResult); }
        }

        /**
         * ワーカーに渡すオプションを組み立てる（パレット側の状態はすべてここで値にする）
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {object} ワーカーへ渡すオプション
         */
        function buildAlignOptions(buttonDef) {
            var spec = readAlignSpec(buttonDef);
            /* Option＋クリックの意味はボタンによって変わる
               中央揃え：アートボードではなくマージンの内側の中央へ寄せる
               上下左右：字形の境界をONにしたうえでマージンを無視し、アートボードの辺にぴったり寄せる
               Option-click means different things per button: centre inside the margin for the centred
               alignments, and flush against the artboard edge with glyph bounds for the edge ones */
            var altPressed = isAltPressed();
            var centersInMargin = altPressed && isCenterAlign(spec);
            var ignoresMargin = altPressed && !centersInMargin;
            var options = buildGuideOptions();
            var centerOffset = centersInMargin ?
                centerOffsetInMargin(spec, options.guideMargins) : { x: 0, y: 0 };
            options.alignCommands       = spec.commands;
            options.offsetX             = spec.offsetX;
            options.offsetY             = spec.offsetY;
            options.probeX              = spec.probeX;
            options.probeY              = spec.probeY;
            options.modeX               = spec.modeX;
            options.modeY               = spec.modeY;
            /* 中央からのずれ。Option＋クリックの中央揃えだけ 0 以外になる
               The shift from the artboard centre; only an Option-clicked centre alignment sets it */
            options.centerOffsetX       = centerOffset.x;
            options.centerOffsetY       = centerOffset.y;
            /* Option＋クリックは段階を踏まず、アートボードの端にぴったり寄せる
               Option-click skips the steps and sits flush against the artboard edge */
            options.marginPt            = ignoresMargin ? 0 : marginForAlign(spec, options.guideMargins);
            options.bleedPt             = ignoresMargin ? 0 : readBleedPt();
            options.perArtboard         = alignPerArtboardCheckbox !== null && alignPerArtboardCheckbox.value === true;
            options.minDeltaPt          = MOVE_MIN_DELTA_PT;
            options.previewBounds       = previewBoundsCheckbox !== null && previewBoundsCheckbox.value === true;
            options.glyphBounds         = ignoresMargin || (glyphBoundsCheckbox !== null && glyphBoundsCheckbox.value === true);
            options.changeJustification = changeJustificationCheckbox !== null && changeJustificationCheckbox.value === true;
            options.justification       = spec.justification;
            return options;
        }

        /**
         * 移動ボタンのオプションを組み立てる
         * マージンは使わず（ガイドかアートボードの端にぴったり寄せる）、境界の測り方だけオプションに従う
         * @param {string} directionKey - MOVE_SIDES_BY_DIRECTION のキー
         * @returns {object} ワーカーへ渡すオプション
         */
        function buildMoveOptions(directionKey) {
            var sides = MOVE_SIDES_BY_DIRECTION[directionKey];
            return {
                sides:          sides,
                bleedPt:        readBleedPt(),
                perArtboard:    alignPerArtboardCheckbox !== null && alignPerArtboardCheckbox.value === true,
                previewBounds:  previewBoundsCheckbox !== null && previewBoundsCheckbox.value === true,
                glyphBounds:    glyphBoundsCheckbox !== null && glyphBoundsCheckbox.value === true,
                changeJustification: changeJustificationCheckbox !== null && changeJustificationCheckbox.value === true,
                justification:  justificationForSides(sides),
                guideTolerance: GUIDE_ORIENTATION_TOLERANCE,
                minDeltaPt:     MOVE_MIN_DELTA_PT
            };
        }

        /**
         * その移動で合わせる行揃えを求める（左右へ寄せるときだけ。上下では変えない）
         * @param {string[]} sides - 寄せる辺
         * @returns {string} 行揃えの名前。合わせないときは null
         */
        function justificationForSides(sides) {
            for (var i = 0; i < sides.length; i++) {
                if (sides[i] === "left") { return ALIGN_COMMANDS.horizontalLeft.justification; }
                if (sides[i] === "right") { return ALIGN_COMMANDS.horizontalRight.justification; }
            }
            return null;
        }

        /**
         * ワーカーの戻り値を状況表示に反映する
         * @param {string} workerResult - ワーカーが返したマーカー
         * @returns {void}
         */
        function showWorkerResult(workerResult) {
            if (workerResult === null) {
                setStatus(getLabel("status", "noResponse"));
                return;
            }
            if (workerResult.indexOf("ERR:") === 0) {
                setStatus(getLabel("status", "genericError") + workerResult.substring(4));
                return;
            }
            /* 「OK:CENTER」のように行揃えを変えたときは、変更後の行揃えも知らせる
               An "OK:CENTER" marker also reports the justification that was applied */
            if (workerResult.indexOf("OK:") === 0) {
                showJustifiedResult(workerResult.substring(3), "doneJustified", "done");
                return;
            }
            if (workerResult.indexOf("MOVED:") === 0) {
                showJustifiedResult(workerResult.substring(6), "movedJustified", "moved");
                return;
            }
            var statusKey = STATUS_BY_MARKER[workerResult];
            setStatus(statusKey ? getLabel("status", statusKey) : workerResult);
        }

        /**
         * 行揃えを変えたときの状況表示を出す（名前が引けなければ、変えなかったときと同じ文言にする）
         * @param {string} justificationName - ワーカーが返した行揃えの名前
         * @param {string} justifiedKey - 行揃えを添える文言のキー
         * @param {string} plainKey - 添えないときの文言のキー
         * @returns {void}
         */
        function showJustifiedResult(justificationName, justifiedKey, plainKey) {
            var justificationLabel = getLabel("status", "justification", justificationName);
            setStatus(justificationLabel ?
                getLabel("status", justifiedKey).replace("{0}", justificationLabel) :
                getLabel("status", plainKey));
        }

        /* 取り直しの状態 / State of the refresh */
        var isRefreshingSelection = false;
        var lastSelectionRefreshTime = 0;
        /* 直前の選択（種類＋個数）。変わったら前回の実行結果の表示を消す
           The previous selection (kind and count); a change clears the last result from the status line */
        var lastSelectionSignature = null;

        /**
         * 選択の種類・定規の単位・選択数・アートボード数をメインエンジンに問い合わせ、ディムと単位ラベルを更新する
         * 1往復でまとめて受け取り（"TEXT|2|3|1" の形）、選択が変わっていれば状況表示も消す
         * @returns {void}
         */
        function refreshPaletteState() {
            /* 整列の実行中や取り直しの最中に割り込ませない（同期送信の待ち時間にイベントが入り得るため）
               Never nest inside a running align or another refresh; events can fire while the send waits */
            if (isBusy || isRefreshingSelection) { return; }
            isRefreshingSelection = true;
            try {
                var workerResult = runWorker("btGetPaletteState();");
                /* 応答なし・エラーのときは、当てにならない値で表示を書き換えない
                   Leave the palette as it is when there is no usable answer */
                if (workerResult === null || workerResult.indexOf("ERR:") === 0) { return; }
                var parts = workerResult.split("|");
                if (changeJustificationCheckbox !== null) {
                    changeJustificationCheckbox.enabled = (parts[0] === "TEXT");
                }
                /* アートボードが1つしかないなら束ね分ける先が無いので、［アートボードごとに整列］はディムにする */
                if (alignPerArtboardCheckbox !== null) {
                    alignPerArtboardCheckbox.enabled = (Number(parts[3]) > 1);
                }
                currentUnitInfo = UNIT_INFO[parts[1]] || FALLBACK_UNIT_INFO;
                if (marginPanel !== null) { marginPanel.text = marginPanelTitle(); }
                updateDivideUnitLabels();
                fillDefaultExtension();
                fillDefaultMargins();
                var selectionSignature = parts[0] + "|" + parts[2];
                if (lastSelectionSignature !== null && selectionSignature !== lastSelectionSignature) {
                    setStatus("");
                }
                lastSelectionSignature = selectionSignature;
            } finally {
                isRefreshingSelection = false;
            }
        }

        /**
         * パレットへフォーカスが来たときに選択と定規の単位を取り直す
         * Illustrator にタイマーAPIが無いため、変化はこの瞬間に拾う
         * @param {boolean} force - true なら間引きを無視して必ず取り直す
         * @returns {void}
         */
        function onPaletteFocus(force) {
            var now = (new Date()).getTime();
            if (!force && (now - lastSelectionRefreshTime) < SELECTION_POLL_INTERVAL_MS) { return; }
            lastSelectionRefreshTime = now;
            refreshPaletteState();
        }

        /**
         * 整列をメインエンジンへ委譲する
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {string} ワーカーが返したマーカー（応答がなければ null）
         */
        function runAlign(buttonDef) {
            var options = buildAlignOptions(buttonDef);
            return runWorker("btAlignSelection(" + options.toSource() + ");");
        }

        /**
         * 端またはガイドへの移動をメインエンジンへ委譲する
         * @param {string} directionKey - "up" / "left" / "right" / "down"
         * @returns {string} ワーカーが返したマーカー（応答がなければ null）
         */
        function runMove(directionKey) {
            var options = buildMoveOptions(directionKey);
            return runWorker("btMoveToEdgeOrGuide(" + options.toSource() + ");");
        }

        showPalette();

    })();

})();
