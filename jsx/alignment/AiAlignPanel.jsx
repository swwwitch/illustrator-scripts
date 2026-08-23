#target illustrator
#targetengine "AiAlignPanel"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクトを、アートボードの四辺から均一のマージンを取った位置に整列する常駐パレットです。水平方向（左・中央・右）、水平垂直の中央、垂直方向（上・中央・下）の7つを、アイコンのクリックで即時実行します。

詳細は README を参照してください。

### Overview

A persistent palette that aligns the selection to the artboard, inset by a uniform margin on all four sides.
Seven icons run horizontal (left / center / right), both axes at once, and vertical (top / center / bottom)
alignment immediately on click.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAlignPanel";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-23";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAlignPanel.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAlignPanel.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 常駐エンジンに残すパレット参照（GC回避と多重起動防止を兼ねる）
   var の初期化は再実行のたびに走るため、既存の参照を消さないよう $.global から引き継ぐ
   The palette reference lives in the persistent engine; it is carried over from $.global so a
   re-run does not wipe it before closeExistingPalette() can close the old window */
var paletteWindow = $.global.__aiAlignPanelWindow || null;

(function() {

    // =========================================
    // 整列コマンド / Align commands
    // =========================================
    /* 「オブジェクト > 整列」のメニューコマンド名と、マージンぶん内側へ動かす向き
       整列コマンドはアートボードの辺にぴったり寄せるので、そこから offset 方向へマージンぶん動かす
       Y は上が正のため、上揃えは -1（下へ）・下揃えは +1（上へ）になる
       Menu command names under Object > Align, with the direction to move by the margin;
       Y grows upward, so top align moves -1 (down) and bottom align +1 (up) */
    var ALIGN_COMMANDS = {
        horizontalLeft:   { command: "Horizontal Align Left",   offsetX:  1, offsetY:  0 },
        horizontalCenter: { command: "Horizontal Align Center",  offsetX:  0, offsetY:  0 },
        horizontalRight:  { command: "Horizontal Align Right",   offsetX: -1, offsetY:  0 },
        verticalTop:      { command: "Vertical Align Top",       offsetX:  0, offsetY: -1 },
        verticalCenter:   { command: "Vertical Align Center",    offsetX:  0, offsetY:  0 },
        verticalBottom:   { command: "Vertical Align Bottom",    offsetX:  0, offsetY:  1 }
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
    /* チェックボックスの初期状態。環境設定は読まず、この値を整列のたびにワーカーが書き込む
       Initial checkbox states; the preferences are not read back, the worker writes these on every align */
    var DEFAULT_PREVIEW_BOUNDS       = false; /* プレビュー境界 / preview bounds */
    var DEFAULT_GLYPH_BOUNDS         = true;  /* 字形の境界に整列 / align to glyph bounds */
    var DEFAULT_CHANGE_JUSTIFICATION = true;  /* 行揃えを変更 / change justification */
    var DEFAULT_MARGIN               = 0;     /* マージン欄の初期値（定規の単位）/ initial margin, in ruler units */

    /* メインエンジンからの応答を待つ秒数 / seconds to wait for the main engine */
    var WORKER_TIMEOUT = 10;
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
    var FIELD_CHARS     = 4;    /* マージン入力欄の文字数 / width of the margin field */
    var LABEL_FIELD_SPACING = 4; /* 入力欄と単位ラベルの間隔（既定は広すぎる）/ gap between the field and its unit label */
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
            margin:  { ja: "マージン", en: "Margin" },
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
                ja: "水平方向の中央揃えのとき、1行だけのテキスト1つの行揃えも中央にする",
                en: "Also center the justification of a lone single-line text object"
            }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
            glyphBounds:   { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            changeJustification: { ja: "行揃えを変更", en: "Change Justification" }
        },
        status: {
            done:           { ja: "整列しました。", en: "Aligned." },
            noDocument:     { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:    { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
            multipleLayers: { ja: "レイヤーをまたぐ選択は整列できません。", en: "Cannot align a selection spanning layers." },
            noResponse:     { ja: "Illustrator から応答がありません。", en: "No response from Illustrator." },
            genericError:   { ja: "エラー：", en: "Error: " }
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
     * UI が明るいテーマかを判定する
     * @returns {boolean} 明るいテーマなら true（取得失敗時は false＝暗い側）
     */
    function isLightUI() {
        return getUIBrightness() > 0.5;
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
        var lightUI = isLightUI();
        var uiBrightness = getUIBrightness();
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
     * 水平方向の整列アイコン（横バー2本＋縦の基準線）を描く
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number} size - アイコンの一辺の長さ
     * @param {string} alignMode - "start"＝左 / "center"＝中央 / "end"＝右
     * @param {number[]} color - RGBA の配列
     * @returns {void}
     */
    function drawHorizontalAlignIcon(graphics, size, alignMode, color) {
        var rulePosition = getRulePosition(size, alignMode);
        var barThickness = Math.round(size * ICON_BAR_THICKNESS);
        var barGap = Math.round(size * ICON_BAR_GAP);
        var barTop = getBarStackOrigin(size);
        var barLengths = getBarLengths(size, "horizontal");
        var clearance = Math.round(size * ICON_RULE_CLEARANCE);

        /* 基準線を先に描き、バーを上に重ねる（中央の基準線がバーの下を通って見える）
           Draw the rule first and the bars on top, so a center rule runs behind them */
        var ruleInset = Math.round(size * ICON_RULE_INSET);
        fillRect(graphics, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);

        for (var i = 0; i < barLengths.length; i++) {
            fillRect(graphics,
                getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                barTop + i * (barThickness + barGap),
                barLengths[i], barThickness, color);
        }
    }

    /**
     * 垂直方向の整列アイコン（縦バー2本＋横の基準線）を描く
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number} size - アイコンの一辺の長さ
     * @param {string} alignMode - "start"＝上 / "center"＝中央 / "end"＝下
     * @param {number[]} color - RGBA の配列
     * @returns {void}
     */
    function drawVerticalAlignIcon(graphics, size, alignMode, color) {
        var rulePosition = getRulePosition(size, alignMode);
        var barThickness = Math.round(size * ICON_BAR_THICKNESS);
        var barGap = Math.round(size * ICON_BAR_GAP);
        var barLeft = getBarStackOrigin(size);
        var barLengths = getBarLengths(size, "vertical");
        var clearance = Math.round(size * ICON_RULE_CLEARANCE);

        var ruleInset = Math.round(size * ICON_RULE_INSET);
        fillRect(graphics, ruleInset, rulePosition - 0.5, size - ruleInset * 2, 1, color);

        for (var i = 0; i < barLengths.length; i++) {
            fillRect(graphics,
                barLeft + i * (barThickness + barGap),
                getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                barThickness, barLengths[i], color);
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

        if (button.iconType === "horizontal") {
            drawHorizontalAlignIcon(graphics, width, button.alignMode, iconColor);
        } else if (button.iconType === "vertical") {
            drawVerticalAlignIcon(graphics, width, button.alignMode, iconColor);
        } else {
            drawCenterBothIcon(graphics, width, iconColor);
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
     * 数値入力欄を ↑↓ キーで増減できるようにする（Shift＝±10・Option＝±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                /* Optionキー押下時は0.1単位で増減 / Step by 0.1 when Option is held */
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                /* 小数第1位までに丸め / Round to one decimal place */
                value = Math.round(value * 10) / 10;
            } else {
                /* 整数に丸め / Round to an integer */
                value = Math.round(value);
            }

            editText.text = value;
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
    /* マージン欄と単位ラベルの参照 / The margin field and its unit label */
    var marginField = null;
    var marginUnitText = null;
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
            runExclusive(function() { runAlign(buttonDef); });
            /* 実行後の選択に合わせてディムを更新する（runExclusive を抜けてから呼ぶ）
               Refresh the dimming for the resulting selection, after runExclusive has released isBusy */
            onPaletteFocus(true);
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
        var marginPanel = parentRow.add("panel", undefined, getLabel("panel", "margin"));
        marginPanel.orientation = "column";
        marginPanel.alignment = ["fill", "fill"];
        marginPanel.alignChildren = ["left", "top"];
        marginPanel.margins = PANEL_MARGINS;
        marginPanel.spacing = OPTION_SPACING;

        var fieldRow = marginPanel.add("group");
        setupRow(fieldRow, "left", LABEL_FIELD_SPACING);
        marginField = fieldRow.add("edittext", undefined, String(DEFAULT_MARGIN));
        marginField.characters = FIELD_CHARS;
        marginField.helpTip = getLabel("tooltip", "margin");
        changeValueByArrowKey(marginField);
        /* 単位ラベルは幅を固定する（単位が変わってもパレット幅が動かないように）
           The unit label has a fixed width so the palette does not resize when the unit changes */
        marginUnitText = fieldRow.add("statictext", undefined, currentUnitInfo.label);
        marginUnitText.preferredSize.width = UNIT_LABEL_WIDTH;
        marginUnitText.maximumSize.width = UNIT_LABEL_WIDTH;
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

        /* 環境設定は読まない（常駐パレットからは信頼できないため）。ここでの値を整列のたびにワーカーが書き込む
           The preferences are not read here; the worker writes these values on every align */
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
     * 状況表示を書き換える
     * @param {string} message - 表示する文言
     * @returns {void}
     */
    function setStatus(message) {
        if (statusText !== null) { statusText.text = message; }
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
        $.global.__aiAlignPanelWindow = null;
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
        /* 閉じるとき：参照を解放（次回起動で作り直せるように）/ On close: release the reference so the next launch rebuilds */
        win.onClose = function() {
            paletteWindow = null;
            $.global.__aiAlignPanelWindow = null;
            return true;
        };

        /* 常駐参照：GC 回避と多重起動の検出を兼ねる / Persistent reference: avoids GC and detects a second launch */
        paletteWindow = win;
        $.global.__aiAlignPanelWindow = win;
        win.layout.layout(true);
        win.show();
        refreshPaletteState();
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
        var doc, selectedItems, needsGroup, i;
        if (app.documents.length === 0) { return "NODOC"; }
        doc = app.activeDocument;
        selectedItems = doc.selection;
        if (selectedItems && !(selectedItems instanceof Array)) {
            selectedItems = btPromoteTextRange(doc, selectedItems);
        }
        if (!(selectedItems instanceof Array) || selectedItems.length === 0) { return "NOSEL"; }
        needsGroup = selectedItems.length > 1;
        if (needsGroup && btSpansMultipleLayers(selectedItems)) { return "MULTILAYER"; }
        btApplyPreferences(options);
        btActivateArtboardForSelection(doc, selectedItems);
        try {
            if (options.changeJustification && options.centersHorizontally && btIsSingleLineTextFrame(selectedItems)) {
                btSetCenterJustification(selectedItems[0]);
            }
            if (needsGroup) { app.executeMenuCommand("group"); }
            for (i = 0; i < options.alignCommands.length; i++) {
                app.executeMenuCommand(options.alignCommands[i]);
            }
            btApplyMargin(doc, options);
        } catch (alignError) {
            return "ERR:" + alignError;
        } finally {
            if (needsGroup) { app.executeMenuCommand("ungroup"); }
        }
        return "OK";
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
        return btGetSelectionKind() + "|" + app.preferences.getIntegerPreference("rulerType");
    }

    function btApplyMargin(doc, options) {
        var items, i, deltaX, deltaY;
        deltaX = options.offsetX * options.marginPt;
        deltaY = options.offsetY * options.marginPt;
        if (deltaX === 0 && deltaY === 0) { return; }
        items = doc.selection;
        if (!(items instanceof Array)) { return; }
        for (i = 0; i < items.length; i++) {
            items[i].translate(deltaX, deltaY);
        }
    }

    function btApplyPreferences(options) {
        app.preferences.setBooleanPreference("includeStrokeInBounds", options.previewBounds === true);
        app.preferences.setBooleanPreference("EnableActualPointTextSpaceAlign", options.glyphBounds === true);
        app.preferences.setBooleanPreference("EnableActualAreaTextSpaceAlign", options.glyphBounds === true);
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

    function btSetCenterJustification(textFrame) {
        textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
    }

    /* 送信するワーカー関数の一覧（追加したらここにも必ず登録する）/ Every worker function shipped to the main engine */
    var WORKER_FUNCS = [
        btAlignSelection, btGetPaletteState, btGetSelectionKind,
        btApplyMargin, btApplyPreferences, btPromoteTextRange,
        btGetLayerKey, btSpansMultipleLayers,
        btGetSelectionBounds, btGetOverlapArea,
        btFindOverlappingArtboardIndex, btFindNearestArtboardIndex, btActivateArtboardForSelection,
        btIsSingleLineTextFrame, btSetCenterJustification
    ];

    // =========================================
    // メインエンジンへの委譲 / Delegating to the main engine
    // =========================================
    /* 常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理は毎回メインエンジンへ送る
       A persistent palette loses its DOM connection while shown, so every DOM touch is delegated */

    /* 実行中フラグ（連打による多重実行を防ぐ）/ Guard against double execution from rapid clicks */
    var isBusy = false;

    /* ワーカーの戻り値マーカーと status ラベルの対応 / Worker markers mapped to status labels */
    var STATUS_BY_MARKER = {
        OK:         "done",
        NODOC:      "noDocument",
        NOSEL:      "noSelection",
        MULTILAYER: "multipleLayers"
    };

    /**
     * 再入防止つきで処理を実行する（連打による多重実行を防ぐ）
     * @param {function} action - 実行する処理
     * @returns {void}
     */
    function runExclusive(action) {
        if (isBusy) { return; }
        isBusy = true;
        try {
            action();
        } finally {
            isBusy = false;
        }
    }

    /**
     * ワーカー関数の定義をひとつの文字列にまとめる
     * @returns {string} 連結したワーカー関数のソース
     */
    function buildWorkerSource() {
        var sources = [];
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            sources.push(WORKER_FUNCS[i].toString());
        }
        return sources.join("\n");
    }

    /**
     * ワーカー関数の呼び出しをメインエンジンで同期実行し、戻り値のマーカーを受け取る
     * @param {string} functionCall - メインエンジンで評価する呼び出し式
     * @returns {string} ワーカーが返したマーカー（応答がなければ null）
     */
    function runWorker(functionCall) {
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        /* バックスラッシュ・多バイト文字・改行が途中で壊れないよう、ソースはURIエンコードして送る
           URI-encode the source so backslashes, multi-byte characters and newlines survive the trip */
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(buildWorkerSource() + "\n" + functionCall) + "\"));";

        /* 同期送信の結果は holder 経由で受け取る / The synchronous send hands its result back through holder */
        var holder = { result: null };
        bridge.onResult = function(message) { holder.result = message.body; };
        bridge.onError = function(message) { holder.result = "ERR:" + message.body; };
        bridge.send(WORKER_TIMEOUT);
        return holder.result;
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
     * 整列後にマージンぶん動かす向きを求める（中央揃えは 0 になる）
     * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
     * @returns {{offsetX: number, offsetY: number}} 移動の向き
     */
    function getMarginOffsets(buttonDef) {
        var offsets = { offsetX: 0, offsetY: 0 };
        for (var i = 0; i < buttonDef.alignKeys.length; i++) {
            var alignCommand = ALIGN_COMMANDS[buttonDef.alignKeys[i]];
            offsets.offsetX += alignCommand.offsetX;
            offsets.offsetY += alignCommand.offsetY;
        }
        return offsets;
    }

    /**
     * その整列がマージンの影響を受けるか判定する（上下左右のボタンだけ true）
     * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
     * @returns {boolean} 影響を受けるなら true
     */
    function usesMargin(buttonDef) {
        var offsets = getMarginOffsets(buttonDef);
        return offsets.offsetX !== 0 || offsets.offsetY !== 0;
    }

    /**
     * ワーカーに渡すオプションを組み立てる（パレット側の状態はすべてここで値にする）
     * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
     * @returns {object} ワーカーへ渡すオプション
     */
    function buildAlignOptions(buttonDef) {
        var alignCommands = [];
        for (var i = 0; i < buttonDef.alignKeys.length; i++) {
            alignCommands.push(ALIGN_COMMANDS[buttonDef.alignKeys[i]].command);
        }
        var offsets = getMarginOffsets(buttonDef);
        /* Option＋クリックは、字形の境界をONにしたうえでマージンを無視し、アートボードの四辺にぴったり寄せる
           Option-click turns glyph bounds on and ignores the margin, sitting flush against the artboard edge */
        var altPressed = isAltPressed();
        var marginPt = readMarginPt();
        return {
            alignCommands:       alignCommands,
            offsetX:             offsets.offsetX,
            offsetY:             offsets.offsetY,
            marginPt:            altPressed ? 0 : marginPt,
            previewBounds:       previewBoundsCheckbox.value === true,
            glyphBounds:         glyphBoundsCheckbox.value === true || altPressed,
            changeJustification: changeJustificationCheckbox.value === true,
            centersHorizontally: centersHorizontally(buttonDef)
        };
    }

    /**
     * その整列が水平方向の中央揃えを含むか判定する
     * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
     * @returns {boolean} 含むなら true
     */
    function centersHorizontally(buttonDef) {
        for (var i = 0; i < buttonDef.alignKeys.length; i++) {
            if (buttonDef.alignKeys[i] === "horizontalCenter") { return true; }
        }
        return false;
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
        var statusKey = STATUS_BY_MARKER[workerResult];
        setStatus(statusKey ? getLabel("status", statusKey) : workerResult);
    }

    /* 取り直しの状態 / State of the refresh */
    var isRefreshingSelection = false;
    var lastSelectionRefreshTime = 0;

    /**
     * 選択の種類と定規の単位をメインエンジンに問い合わせ、ディムと単位ラベルを更新する
     * 1往復で両方受け取る（"TEXT|2" の形）
     * @returns {void}
     */
    function refreshPaletteState() {
        /* 整列の実行中や取り直しの最中に割り込ませない（同期送信の待ち時間にイベントが入り得るため）
           Never nest inside a running align or another refresh; events can fire while the send waits */
        if (isBusy || isRefreshingSelection) { return; }
        isRefreshingSelection = true;
        try {
            var parts = String(runWorker("btGetPaletteState();")).split("|");
            if (changeJustificationCheckbox !== null) {
                changeJustificationCheckbox.enabled = (parts[0] === "TEXT");
            }
            currentUnitInfo = UNIT_INFO[parts[1]] || FALLBACK_UNIT_INFO;
            if (marginUnitText !== null) { marginUnitText.text = currentUnitInfo.label; }
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
     * アイコンがクリックされたときに、整列をメインエンジンへ委譲して結果を表示する
     * @param {object} buttonDef - ALIGN_BUTTON_GROUPS のボタン定義
     * @returns {void}
     */
    function runAlign(buttonDef) {
        var options = buildAlignOptions(buttonDef);
        showWorkerResult(runWorker("btAlignSelection(" + options.toSource() + ");"));
    }

    showPalette();

})();
