#target illustrator
#targetengine "AiAlignToArtboard"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクトを、アートボードを対象に整列する常駐パレットです。
アイコンのクリックで即時に整列し、アートボードの端から空けるマージンも指定できます。

詳細は README を参照してください。

### Overview

A persistent palette that aligns the selected objects to the artboard.
Clicking an icon aligns them immediately, and a margin from the artboard edge can be set.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAlignToArtboard";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-30";                   /* 更新日 / last updated */

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
        var UNIT_INFO = {
            "0":  { label: "in",    points: 72.0 },
            "1":  { label: "mm",    points: 72.0 / 25.4 },
            "2":  { label: "pt",    points: 1.0 },
            "3":  { label: "pica",  points: 12.0 },
            "4":  { label: "cm",    points: 72.0 / 2.54 },
            "5":  { label: "Q/H",   points: (72.0 / 25.4) * 0.25 },
            "6":  { label: "px",    points: 1.0 },
            "7":  { label: "ft/in", points: 864.0 },
            "8":  { label: "m",     points: (72.0 / 25.4) * 1000.0 },
            "9":  { label: "yd",    points: 2592.0 },
            "10": { label: "ft",    points: 864.0 }
        };
        /* 単位が取れないときの既定 / Fallback when the ruler unit cannot be read */
        var FALLBACK_UNIT_INFO = { label: "pt", points: 1.0 };

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
        var DEFAULT_MARGIN               = 0;     /* マージン欄の初期値（定規の単位）/ initial margin, in ruler units */
        var DEFAULT_SHOW_GUIDE           = false; /* ガイドを表示 / show the margin guide */
        var DEFAULT_KEEP_GUIDE           = false; /* ガイドを保持（閉じても残す）/ keep the guide when the palette closes */

        /* マージンのガイドを作るレイヤー名（他のガイド系スクリプトと共通）
           Layer that receives the margin guide, shared with the other guide scripts */
        var GUIDE_LAYER_NAME = "_guide";
        /* このスクリプトが作るガイドの名前。張り替えるときの目印にする
           The name given to the guide, used to find and replace it */
        var GUIDE_NAME = "AiAlignToArtboard-margin";

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
        /* 選択を取り直す最短間隔（mouseover は何度も発生するため間引く）/ Throttle for the mouseover refresh */
        var SELECTION_POLL_INTERVAL_MS = 400;

        // =========================================
        // レイアウト / Layout
        // =========================================
        var WINDOW_MARGINS  = 15;   /* パレット外周の余白 / window margin */
        var WINDOW_SPACING  = 12;   /* パレット内の要素間隔 / window spacing */
        var ICON_SIZE       = 30;   /* 整列アイコン1個の大きさ（px）/ size of each align icon (px) */
        var ICON_GAP        = 6;    /* 整列アイコンどうしの間隔 / gap between align icons */
        var ICON_GROUP_GAP  = 14;   /* 水平・中央・垂直の3グループの間隔 / gap between the three icon groups */
        var ICON_ROW_BOTTOM = 10;   /* ボタンエリアの下の余白 / margin below the icon row */
        var PANEL_MARGINS   = [12, 16, 12, 10]; /* オプションパネルの余白 [左,上,右,下]（上はタイトルのぶん広め）/ options panel margins */
        var COLUMN_SPACING  = 8;    /* マージンパネルとオプションパネルの間隔 / gap between the two columns */
        var FIELD_CHARS     = 3;    /* マージン入力欄の文字数 / width of the margin field */
        var LABEL_FIELD_SPACING = 4; /* 入力欄と単位ラベルの間隔（既定は広すぎる）/ gap between the field and its unit label */
        var FIELD_INDENT    = 18;   /* 入力欄の字下げ（チェックボックスのラベルに頭をそろえる）/ indent of the margin field */
        var FIELD_BOTTOM_GAP = 5;   /* 入力欄の下の余白（［ガイドを保持］との間を空ける）/ gap below the margin field */
        var UNIT_LABEL_WIDTH = 34;  /* 単位ラベルの幅（単位が変わっても幅が動かないよう固定）/ fixed width of the unit label */
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
        var ICON_BLOCK_WIDTH    = 0.38;  /* 中央揃えアイコンの中央に置くオブジェクトの幅 / width of the object block in the center icon */
        var ICON_BLOCK_HEIGHT   = 0.30;  /* 同・高さ / height of that block */

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
                title: { ja: "アートボードに整列", en: "Align to Artboard" }
            },
            panel: {
                guide:   { ja: "ガイド", en: "Guide" },
                options: { ja: "オプション", en: "Options" }
            },
            tooltip: {
                alignLeft:      { ja: "水平方向左に整列", en: "Horizontal Align Left" },
                alignCenterH:   { ja: "水平方向中央に整列", en: "Horizontal Align Center" },
                alignRight:     { ja: "水平方向右に整列", en: "Horizontal Align Right" },
                alignCenterAll: { ja: "水平・垂直方向中央に整列", en: "Align Center on Both Axes" },
                alignTop:       { ja: "垂直方向上に整列", en: "Vertical Align Top" },
                alignCenterV:   { ja: "垂直方向中央に整列", en: "Vertical Align Center" },
                alignBottom:    { ja: "垂直方向下に整列", en: "Vertical Align Bottom" },
                margin: { ja: "アートボードの端から空ける距離", en: "Distance to keep from the artboard edge" },
                keepGuide: {
                    ja: "パレットを閉じてもガイドを残す（OFFのときは閉じるときに削除）",
                    en: "Leave the guide in place when the palette closes (deleted on close when off)"
                },
                showGuide: {
                    ja: "マージンの位置に長方形のガイドを作る（「_guide」レイヤー、アクティブなアートボードに1つ）",
                    en: "Draw a rectangle guide at the margin (on the \"_guide\" layer, one on the active artboard)"
                },
                optionGlyphBounds: { ja: "Option＋クリックで字形の境界に整列", en: "Option-click to align to glyph bounds" },
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
            checkbox: {
                showGuide:     { ja: "ガイドを表示", en: "Show Guides" },
                keepGuide:     { ja: "ガイドを保持", en: "Keep Guides" },
                previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
                glyphBounds:   { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
                changeJustification: { ja: "行揃えを変更", en: "Change Justification" }
            },
            status: {
                done:           { ja: "整列しました。", en: "Aligned." },
                doneJustified:  { ja: "整列し、行揃えを{0}に変更しました。", en: "Aligned; justification set to {0}." },
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
         * 水平・垂直の中央に整列するアイコン（十字のケイ線＋中央に置いたオブジェクト）を描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawCenterBothIcon(graphics, size, color) {
            /* ケイ線は他のアイコンの基準線と同じ太さ・長さにそろえる / Match the rules in the other icons */
            var ruleInset = Math.round(size * ICON_RULE_INSET);
            var ruleLength = size - ruleInset * 2;
            var center = Math.round(size / 2) + 0.5;
            fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
            fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);

            /* 中央のオブジェクトはケイ線の上に重ねる / The block sits on top of the rules */
            var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
            var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
            fillRect(graphics,
                Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                blockWidth, blockHeight, color);
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
         * 整列アイコンボタンを描画する
         * @param {Button} button - 対象のボタン（iconType と alignMode を持つ）
         * @returns {void}
         */
        function drawAlignButton(button) {
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

            if (button.iconType === "center") {
                drawCenterBothIcon(graphics, width, iconColor);
            } else {
                drawAlignIcon(graphics, width, button.alignMode, iconColor, button.iconType);
            }
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
        var marginField = null;
        var marginUnitText = null;
        var showGuideCheckbox = null;
        var keepGuideCheckbox = null;
        /* パレットを閉じてもガイドを残すか。閉じる処理でコントロールを触らずに済むよう値を控えておく
           Whether to keep the guide on close; mirrored so teardown never has to read a control */
        var keepGuideOnClose = DEFAULT_KEEP_GUIDE;
        /* 現在の定規単位（パレットへフォーカスが来るたびに取り直す）/ The current ruler unit, re-read on focus */
        var currentUnitInfo = FALLBACK_UNIT_INFO;
        /* 状況表示の参照 / The status line */
        var statusText = null;

        /* 整列アイコンの定義（グループが変わるところで区切る）/ Align icon definitions, grouped as they appear in the row */
        var ALIGN_BUTTON_GROUPS = [
            [
                { iconType: "horizontal", alignMode: "start",  tooltip: "alignLeft",    alignKeys: ["horizontalLeft"] },
                { iconType: "horizontal", alignMode: "center", tooltip: "alignCenterH", alignKeys: ["horizontalCenter"] },
                { iconType: "horizontal", alignMode: "end",    tooltip: "alignRight",   alignKeys: ["horizontalRight"] }
            ],
            [
                { iconType: "center", alignMode: "center", tooltip: "alignCenterAll", alignKeys: ["horizontalCenter", "verticalCenter"] }
            ],
            [
                { iconType: "vertical", alignMode: "start",  tooltip: "alignTop",      alignKeys: ["verticalTop"] },
                { iconType: "vertical", alignMode: "center", tooltip: "alignCenterV",  alignKeys: ["verticalCenter"] },
                { iconType: "vertical", alignMode: "end",    tooltip: "alignBottom",   alignKeys: ["verticalBottom"] }
            ]
        ];

        /**
         * 整列アイコンボタンを1つ生成する
         * @param {Group} parentRow - 追加先の行グループ
         * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
         * @returns {void}
         */
        function addAlignButton(parentRow, buttonDef) {
            /* iconbutton ではなく button を使う（画像なしの iconbutton はクリックが届かないことがある）
               Use button, not iconbutton: an image-less iconbutton does not always receive clicks */
            var button = parentRow.add("button", undefined, "");
            fixControlSize(button, ICON_SIZE, ICON_SIZE);
            /* キー操作はラベルに出さず helpTip に書く / Shortcuts belong in the tooltip, not the label */
            button.helpTip = getLabel("tooltip", buttonDef.tooltip) + "  —  " +
                getLabel("tooltip", usesMargin(buttonDef) ? "optionNoMargin" : "optionGlyphBounds");
            button.iconType = buttonDef.iconType;
            button.alignMode = buttonDef.alignMode;
            button.isHover = false;
            button.onDraw = function() { drawAlignButton(this); };
            button.onClick = function() {
                var workerResult = null;
                var didRun = runExclusive(function() { workerResult = runAlign(buttonDef); });
                /* 実行後の選択に合わせてディムと単位を更新してから、今回の結果を表示する
                   （先に表示すると、この更新で選択が変わったと見なされて消えてしまう）
                   Refresh first, then show this run's result; showing it first would be wiped by the refresh */
                onPaletteFocus(true);
                if (didRun) { showWorkerResult(workerResult); }
            };
            attachHover(button);
        }

        /**
         * 整列アイコンの行を組み立てる
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addAlignButtonRow(targetWindow) {
            var alignRow = targetWindow.add("group");
            setupRow(alignRow, "center", ICON_GROUP_GAP);
            /* オプションパネルとの間を空ける（ウィンドウの spacing に上乗せ）/ Extra room before the options panel */
            alignRow.margins = [0, 0, 0, ICON_ROW_BOTTOM];
            for (var i = 0; i < ALIGN_BUTTON_GROUPS.length; i++) {
                var iconGroup = alignRow.add("group");
                setupRow(iconGroup, "left", ICON_GAP);
                for (var j = 0; j < ALIGN_BUTTON_GROUPS[i].length; j++) {
                    addAlignButton(iconGroup, ALIGN_BUTTON_GROUPS[i][j]);
                }
            }
        }

        /**
         * ボタンの下を2カラムに分け、左にマージン・右にオプションのパネルを並べる
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addColumnsRow(targetWindow) {
            var columnsRow = targetWindow.add("group");
            columnsRow.orientation = "row";
            /* 2枚のパネルを同じ高さでそろえる / Give both panels the same height */
            columnsRow.alignment = ["fill", "top"];
            columnsRow.alignChildren = ["fill", "fill"];
            columnsRow.spacing = COLUMN_SPACING;
            addMarginPanel(columnsRow);
            addOptionsPanel(columnsRow);
        }

        /**
         * マージンパネル（数値欄＋定規の単位ラベル）を組み立てる
         * @param {Group} parentRow - 追加先の行グループ
         * @returns {void}
         */
        function addMarginPanel(parentRow) {
            var marginPanel = parentRow.add("panel", undefined, getLabel("panel", "guide"));
            marginPanel.orientation = "column";
            marginPanel.alignment = ["fill", "fill"];
            marginPanel.alignChildren = ["left", "top"];
            marginPanel.margins = PANEL_MARGINS;
            marginPanel.spacing = OPTION_SPACING;

            /* 前回の設定を引き継ぐ（残したガイドと表示が食い違わないように）/ Carry over the previous settings */
            var guideSettings = loadGuideSettings();

            showGuideCheckbox = marginPanel.add("checkbox", undefined, getLabel("checkbox", "showGuide"));
            showGuideCheckbox.helpTip = getLabel("tooltip", "showGuide");
            showGuideCheckbox.value = guideSettings.showGuide;
            showGuideCheckbox.onClick = function() { syncGuideControls(); runExclusive(refreshMarginGuide); };

            var fieldRow = marginPanel.add("group");
            setupRow(fieldRow, "left", LABEL_FIELD_SPACING);
            /* 上下のチェックボックスのラベルに頭をそろえ、下は少し空ける
               Indented to line up with the checkbox labels, with a little room below */
            fieldRow.margins = [FIELD_INDENT, 0, 0, FIELD_BOTTOM_GAP];
            marginField = fieldRow.add("edittext", undefined, guideSettings.margin);
            marginField.characters = FIELD_CHARS;
            marginField.helpTip = getLabel("tooltip", "margin");
            changeValueByArrowKey(marginField);
            /* 単位ラベルは幅を固定する（単位が変わってもパレット幅が動かないように）
               The unit label has a fixed width so the palette does not resize when the unit changes */
            marginUnitText = fieldRow.add("statictext", undefined, currentUnitInfo.label);
            marginUnitText.preferredSize.width = UNIT_LABEL_WIDTH;
            marginUnitText.maximumSize.width = UNIT_LABEL_WIDTH;
            /* 確定（Enter・フォーカス移動）でガイドを描き直す
               入力途中で毎回描き直すとそのつど委譲が走るため、描き直しは onChange だけにする
               Only redraw the guide when the field commits, not on every keystroke */
            marginField.onChange = function() { runExclusive(refreshMarginGuide); saveGuideSettings(); };

            keepGuideCheckbox = marginPanel.add("checkbox", undefined, getLabel("checkbox", "keepGuide"));
            keepGuideCheckbox.helpTip = getLabel("tooltip", "keepGuide");
            keepGuideCheckbox.value = guideSettings.keepGuide;
            keepGuideCheckbox.onClick = function() { syncGuideControls(); };

            syncGuideControls();
        }

        /**
         * ［ガイドを表示］の状態に合わせて［ガイドを保持］を更新する
         * ガイドを出していないときは保持しようがないのでディムする（設定は残したいので値は落とさない）
         * @returns {void}
         */
        function syncGuideControls() {
            if (showGuideCheckbox === null || keepGuideCheckbox === null) { return; }
            keepGuideCheckbox.enabled = showGuideCheckbox.value === true;
            keepGuideOnClose = keepGuideCheckbox.value === true;
            saveGuideSettings();
        }

        /**
         * 前回のガイド設定を常駐エンジンから読み出す（控えが無ければ既定値）
         * @returns {object} { showGuide: boolean, keepGuide: boolean, margin: string }
         */
        function loadGuideSettings() {
            var stored = $.global[SETTINGS_KEY];
            if (!stored) {
                return { showGuide: DEFAULT_SHOW_GUIDE, keepGuide: DEFAULT_KEEP_GUIDE, margin: String(DEFAULT_MARGIN) };
            }
            return {
                showGuide: stored.showGuide === true,
                keepGuide: stored.keepGuide === true,
                margin: (stored.margin != null) ? String(stored.margin) : String(DEFAULT_MARGIN)
            };
        }

        /**
         * 現在のガイド設定を常駐エンジンに控える
         * @returns {void}
         */
        function saveGuideSettings() {
            $.global[SETTINGS_KEY] = {
                showGuide: showGuideCheckbox !== null && showGuideCheckbox.value === true,
                keepGuide: keepGuideCheckbox !== null && keepGuideCheckbox.value === true,
                margin:    (marginField !== null) ? String(marginField.text) : String(DEFAULT_MARGIN)
            };
        }

        /**
         * オプションパネル（プレビュー境界・字形の境界に整列・行揃えを変更）を組み立てる
         * @param {Group} parentRow - 追加先の行グループ
         * @returns {void}
         */
        function addOptionsPanel(parentRow) {
            var optionsPanel = parentRow.add("panel", undefined, getLabel("panel", "options"));
            optionsPanel.orientation = "column";
            /* パネル自身は行の高さに合わせ、中のチェックボックスは左そろえ（fill の継承を打ち消す）
               The panel matches the row height while its checkboxes stay left-aligned, cancelling the inherited fill */
            optionsPanel.alignment = ["fill", "fill"];
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
            initIconColors();

            var win = new Window("palette", getLabel("dialog", "title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
            win.orientation = "column";
            win.alignChildren = ["fill", "top"];
            win.margins = WINDOW_MARGINS;
            win.spacing = WINDOW_SPACING;

            addAlignButtonRow(win);
            addColumnsRow(win);
            addStatusLine(win);

            /* 選択の変化はパレットへ操作しに来た瞬間に拾う（Illustrator にタイマーAPIが無いため）
               Illustrator has no timer API, so the selection is re-read when the user comes to the palette */
            win.onActivate = function() { onPaletteFocus(true); };
            try {
                win.addEventListener("mouseover", function() { onPaletteFocus(false); });
            } catch (e) {}

            /* Esc で閉じる / Esc closes */
            win.addEventListener("keydown", function(event) {
                if (event.keyName === "Escape") { win.close(); }
            });
            /* 閉じるとき：［ガイドを保持］がOFFならこのパレットが作ったガイドを消し、参照を解放する
               On close: delete the guide unless "Keep Guides" is on, then release the reference */
            win.onClose = function() {
                if (!keepGuideOnClose) { removeMarginGuide(); }
                paletteWindow = null;
                $.global.__aiAlignToArtboardWindow = null;
                return true;
            };

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
            if (showGuideCheckbox !== null && showGuideCheckbox.value === true) { runExclusive(refreshMarginGuide); }
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
            var doc, selectedItems, needsGroup, didGroup, previousPreferences, previousJustification, changedJustification, failure;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            selectedItems = doc.selection;
            if (selectedItems && !(selectedItems instanceof Array)) {
                selectedItems = btPromoteTextRange(doc, selectedItems);
            }
            if (!(selectedItems instanceof Array) || selectedItems.length === 0) { return "NOSEL"; }
            needsGroup = selectedItems.length > 1;
            if (needsGroup && btSpansMultipleLayers(selectedItems)) { return "MULTILAYER"; }
            btActivateArtboardForSelection(doc, selectedItems);
            previousPreferences = btReadPreferences();
            previousJustification = null;
            changedJustification = null;
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
                if (options.changeJustification && options.justification && btIsSingleLineTextFrame(selectedItems)) {
                    previousJustification = btSetJustification(selectedItems[0], options.justification);
                    /* もとから同じ行揃えだったときは「変更した」と言わない */
                    if (previousJustification !== null && previousJustification !== btJustificationByName(options.justification)) {
                        changedJustification = options.justification;
                    }
                }
                if (needsGroup) {
                    app.executeMenuCommand("group");
                    didGroup = true;
                }
                if (!btRunAlignCommands(doc, options)) {
                    failure = "NOTARGET";
                } else {
                    btNudgeSelection(doc, options.offsetX * options.marginPt, options.offsetY * options.marginPt);
                }
            } catch (alignError) {
                failure = "ERR:" + alignError;
            } finally {
                if (didGroup) { app.executeMenuCommand("ungroup"); }
                /* 整列が環境設定を使い終えてから戻す（先に戻すと反映前の値で整列されることがある）*/
                app.redraw();
                btWritePreferences(previousPreferences.previewBounds, previousPreferences.pointText, previousPreferences.areaText);
                /* 整列できなかったときは、先に変えた行揃えも元に戻す（例外で抜けた場合も同じ）
                   Roll the justification back whenever the align did not go through, exceptions included */
                if (failure !== null && previousJustification !== null) {
                    try { selectedItems[0].textRange.paragraphAttributes.justification = previousJustification; } catch (restoreError) {}
                }
            }
            if (failure !== null) { return failure; }
            try {
                /* 整列コマンドが「字形の境界に整列」を拾えていないことがあるため、
                   字形を実測して目標位置との差を打ち消す（拾えていれば差は0で何も動かない）*/
                btApplyGlyphCorrection(doc, options);
                btDrawMarginGuide(doc, options);
            } catch (guideError) {
                return "ERR:" + guideError;
            }
            if (changedJustification !== null) { return "OK:" + changedJustification; }
            return "OK";
        }

        function btUpdateMarginGuide(options) {
            var doc;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            btDrawMarginGuide(doc, options);
            return "OK";
        }

        function btDrawMarginGuide(doc, options) {
            var layer, previousLayerState, artboardRect, left, top, right, bottom, guideRectangle;
            btRemoveGuidesByName(doc, options.guideName, options.guideLayerName);
            if (options.showGuide !== true || !(options.guideMarginPt > 0)) { return; }
            if (doc.artboards.length === 0) { return; }
            artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            left = artboardRect[0] + options.guideMarginPt;
            top = artboardRect[1] - options.guideMarginPt;
            right = artboardRect[2] - options.guideMarginPt;
            bottom = artboardRect[3] + options.guideMarginPt;
            /* マージンが大きすぎて内側が残らないときは何も描かない */
            if (right <= left || top <= bottom) { return; }
            layer = btGetGuideLayer(doc, options.guideLayerName);
            previousLayerState = btUnlockLayer(layer);
            try {
                guideRectangle = layer.pathItems.rectangle(top, left, right - left, top - bottom);
                guideRectangle.name = options.guideName;
                guideRectangle.stroked = false;
                guideRectangle.filled = false;
                guideRectangle.guides = true;
            } finally {
                btRestoreLayer(layer, previousLayerState);
            }
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

        function btExecuteAlignCommands(alignCommands) {
            var i;
            for (i = 0; i < alignCommands.length; i++) {
                app.executeMenuCommand(alignCommands[i]);
            }
        }

        function btRunAlignCommands(doc, options) {
            var boundsBefore, boundsAfter;
            boundsBefore = btGetSelectionBounds(doc.selection);
            btExecuteAlignCommands(options.alignCommands);
            boundsAfter = btGetSelectionBounds(doc.selection);
            if (!btSameBounds(boundsBefore, boundsAfter)) { return true; }
            return btProbeAlignTarget(doc, options);
        }

        function btProbeAlignTarget(doc, options) {
            var boundsBefore, boundsAfter;
            btNudgeSelection(doc, options.probeX, options.probeY);
            try {
                boundsBefore = btGetSelectionBounds(doc.selection);
                btExecuteAlignCommands(options.alignCommands);
                boundsAfter = btGetSelectionBounds(doc.selection);
            } catch (probeError) {
                btNudgeSelection(doc, -options.probeX, -options.probeY);
                throw probeError;
            }
            if (btSameBounds(boundsBefore, boundsAfter)) {
                btNudgeSelection(doc, -options.probeX, -options.probeY);
                return false;
            }
            return true;
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
            if (!btHasTextFrame(items)) { return; }
            if (doc.artboards.length === 0) { return; }
            bounds = btGetGlyphAwareBounds(items, options.previewBounds === true);
            if (bounds === null) { return; }
            artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            /* artboardRect は [左, 上, 右, 下] で、X は右へ・Y は下へ向かって内側になるため符号が逆になる */
            deltaX = btAlignDelta(bounds, artboardRect, options.modeX, options.marginPt, 0, 2, 1);
            deltaY = btAlignDelta(bounds, artboardRect, options.modeY, options.marginPt, 1, 3, -1);
            if (Math.abs(deltaX) < 0.001) { deltaX = 0; }
            if (Math.abs(deltaY) < 0.001) { deltaY = 0; }
            btNudgeSelection(doc, deltaX, deltaY);
        }

        function btAlignDelta(bounds, artboardRect, mode, marginPt, startIndex, endIndex, inwardSign) {
            if (mode === "start") { return (artboardRect[startIndex] + inwardSign * marginPt) - bounds[startIndex]; }
            if (mode === "end") { return (artboardRect[endIndex] - inwardSign * marginPt) - bounds[endIndex]; }
            if (mode === "center") { return (artboardRect[startIndex] + artboardRect[endIndex]) / 2 - (bounds[startIndex] + bounds[endIndex]) / 2; }
            return 0;
        }

        function btHasTextFrame(items) {
            var i;
            for (i = 0; i < items.length; i++) {
                if (items[i].typename === "TextFrame") { return true; }
            }
            return false;
        }

        function btGetGlyphAwareBounds(items, usePreviewBounds) {
            var bounds, i, itemBounds;
            bounds = null;
            for (i = 0; i < items.length; i++) {
                itemBounds = btGetMeasureBounds(items[i], usePreviewBounds);
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

        function btGetMeasureBounds(item, usePreviewBounds) {
            var bounds;
            if (item.typename === "TextFrame") {
                bounds = btGetOutlineBounds(item, usePreviewBounds);
                if (bounds !== null) { return bounds; }
            }
            return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
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
            var doc, count;
            count = 0;
            if (app.documents.length > 0) {
                doc = app.activeDocument;
                if (doc.selection instanceof Array) { count = doc.selection.length; }
                else if (doc.selection) { count = 1; }
            }
            return btGetSelectionKind() + "|" + app.preferences.getIntegerPreference("rulerType") + "|" + count;
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

        function btGetSelectionBounds(selectedItems) {
            var bounds, left, top, right, bottom, i, itemBounds;
            bounds = selectedItems[0].visibleBounds;
            left = bounds[0];
            top = bounds[1];
            right = bounds[2];
            bottom = bounds[3];
            for (i = 1; i < selectedItems.length; i++) {
                itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < left) { left = itemBounds[0]; }
                if (itemBounds[1] > top) { top = itemBounds[1]; }
                if (itemBounds[2] > right) { right = itemBounds[2]; }
                if (itemBounds[3] < bottom) { bottom = itemBounds[3]; }
            }
            return [left, top, right, bottom];
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
            selectionBounds = btGetSelectionBounds(selectedItems);
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
            btAlignSelection, btExecuteAlignCommands, btRunAlignCommands, btProbeAlignTarget, btSameBounds,
            btUpdateMarginGuide, btDrawMarginGuide, btGetGuideLayer, btUnlockLayer, btRestoreLayer, btRemoveGuidesByName,
            btApplyGlyphCorrection, btAlignDelta, btHasTextFrame,
            btGetGlyphAwareBounds, btGetMeasureBounds, btGetOutlineBounds, btSafeRemove,
            btGetPaletteState, btGetSelectionKind, btJustificationByName,
            btNudgeSelection, btReadPreferences, btGetPreferenceFlags, btWritePreferences, btPromoteTextRange,
            btGetLayerKey, btSpansMultipleLayers,
            btGetSelectionBounds, btGetOverlapArea,
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
         * マージン欄を読んで pt に換算する（数値以外と負数は0に丸め、欄の表示もそろえる）
         * @returns {number} マージン（pt）
         */
        function readMarginPt() {
            if (marginField === null) { return 0; }
            var marginValue = Number(marginField.text);
            if (isNaN(marginValue) || marginValue < 0) { marginValue = 0; }
            /* 手入力が丸められたときは欄の表示も実際に使う値にそろえる / Show the value actually used */
            if (String(marginValue) !== marginField.text) { marginField.text = marginValue; }
            return marginValue * currentUnitInfo.points;
        }

        /**
         * ボタン定義から、そのクリックに必要な値をまとめて求める
         * 実行するメニューコマンド、整列後にマージンぶん動かす向き（中央揃えは0）、整列先の判定に使う仮移動量、
         * 軸ごとの寄せ先（字形の境界での補正が目標位置の計算に使う）、合わせる行揃えを、ALIGN_COMMANDS の1周で得る
         * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
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
         * その整列がマージンの影響を受けるか判定する（上下左右のボタンだけ true）
         * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
         * @returns {boolean} 影響を受けるなら true
         */
        function usesMargin(buttonDef) {
            var spec = readAlignSpec(buttonDef);
            return spec.offsetX !== 0 || spec.offsetY !== 0;
        }

        /**
         * ガイドの作成に必要な値を組み立てる
         * マージンは Option＋クリックの影響を受けない（ガイドは入力欄の値をそのまま表す）
         * @returns {object} ワーカーへ渡すガイドのオプション
         */
        function buildGuideOptions() {
            return {
                showGuide:      showGuideCheckbox !== null && showGuideCheckbox.value === true,
                guideMarginPt:  readMarginPt(),
                guideName:      GUIDE_NAME,
                guideLayerName: GUIDE_LAYER_NAME
            };
        }

        /**
         * このスクリプトが作ったガイドを削除する
         * パレットを閉じるときに呼ぶため、コントロールを参照せず値を直接組み立てる
         * （閉じる処理は何があっても止めないよう、失敗は握りつぶす）
         * @returns {void}
         */
        function removeMarginGuide() {
            try {
                var options = {
                    showGuide:      false,
                    guideMarginPt:  0,
                    guideName:      GUIDE_NAME,
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
         * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
         * @returns {object} ワーカーへ渡すオプション
         */
        function buildAlignOptions(buttonDef) {
            var spec = readAlignSpec(buttonDef);
            /* Option＋クリックは、字形の境界をONにしたうえでマージンを無視し、アートボードの四辺にぴったり寄せる
               Option-click turns glyph bounds on and ignores the margin, sitting flush against the artboard edge */
            var altPressed = isAltPressed();
            var options = buildGuideOptions();
            options.alignCommands       = spec.commands;
            options.offsetX             = spec.offsetX;
            options.offsetY             = spec.offsetY;
            options.probeX              = spec.probeX;
            options.probeY              = spec.probeY;
            options.modeX               = spec.modeX;
            options.modeY               = spec.modeY;
            options.marginPt            = altPressed ? 0 : options.guideMarginPt;
            options.previewBounds       = previewBoundsCheckbox !== null && previewBoundsCheckbox.value === true;
            options.glyphBounds         = altPressed || (glyphBoundsCheckbox !== null && glyphBoundsCheckbox.value === true);
            options.changeJustification = changeJustificationCheckbox !== null && changeJustificationCheckbox.value === true;
            options.justification       = spec.justification;
            return options;
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
                var justificationLabel = getLabel("status", "justification", workerResult.substring(3));
                setStatus(justificationLabel ?
                    getLabel("status", "doneJustified").replace("{0}", justificationLabel) :
                    getLabel("status", "done"));
                return;
            }
            var statusKey = STATUS_BY_MARKER[workerResult];
            setStatus(statusKey ? getLabel("status", statusKey) : workerResult);
        }

        /* 取り直しの状態 / State of the refresh */
        var isRefreshingSelection = false;
        var lastSelectionRefreshTime = 0;
        /* 直前の選択（種類＋個数）。変わったら前回の実行結果の表示を消す
           The previous selection (kind and count); a change clears the last result from the status line */
        var lastSelectionSignature = null;

        /**
         * 選択の種類・定規の単位・選択数をメインエンジンに問い合わせ、ディムと単位ラベルを更新する
         * 1往復でまとめて受け取り（"TEXT|2|3" の形）、選択が変わっていれば状況表示も消す
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
                currentUnitInfo = UNIT_INFO[parts[1]] || FALLBACK_UNIT_INFO;
                if (marginUnitText !== null) { marginUnitText.text = currentUnitInfo.label; }
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
         * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
         * @returns {string} ワーカーが返したマーカー（応答がなければ null）
         */
        function runAlign(buttonDef) {
            var options = buildAlignOptions(buttonDef);
            return runWorker("btAlignSelection(" + options.toSource() + ");");
        }

        showPalette();

    })();

})();
