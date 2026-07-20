#targetengine "SmartGridMakerEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択したパスアイテム（長方形など）、または現在のアートボードを基準に、囲み罫とグリッドをプレビュー付きで一括生成します。
- 外側エリア・タイトルエリア・内側エリア・フレームを、4つのタブに分けたひとつのダイアログで設定します。

### 主な機能

- **アートボードtab**：マージン（上下左右・連動）とフレーム。どちらもアートボードを基準に開始したときだけ使用でき、長方形を選択して開始した場合はタブごと非表示になります
- **フレーム**：アートボードの外周に作成。裁ち落とし（3mm）はフレームにのみ加算され、内側の罫線には影響しません
- **外側エリア**：外枠を残すかどうか、辺の伸縮（＋で伸ばす／−で短くする）、線端、角丸。辺の伸縮を0以外にすると外枠は4辺の線に分解されます
- **タイトルエリア**：位置（上下左右）・サイズ・塗り・境界線・境界線の伸縮。角丸は外側エリアの値を、配置位置に応じた2角にだけ適用します
- **内側エリア**：オフセット（上下左右・連動）、列／行の分割と間隔、セルの塗り、分割線、線種（実線／点線／ドット点線）
- **画面表示tab**：ズームとパン（Option＋ドラッグで1/10速）、および表示コマンド
- ダイアログの入力値は Illustrator の起動中は保持され、次回実行時に復元されます
- UIは日英ローカライズ対応。各コントロールにツールチップを用意しています

### 処理の流れ

1. 選択されたパスアイテム、または現在のアートボードから基準矩形を決める
2. ダイアログで各値を入力し、プレビューを随時再描画する
3. OK で最終生成、キャンセルでプレビュー生成物を破棄して元の状態へ戻す

---

### Overview

- Generates a frame rule and a grid from the selected path items (a rectangle, typically) or from the active artboard, with live preview.
- Outer area, title area, inner area and frame are configured across four tabs in a single dialog.

### Features

- **Artboard tab**: margins (top/bottom/left/right, linkable) and the frame. Both require the artboard as the base, so the whole tab is hidden when you start from a selected rectangle
- **Frame**: drawn around the artboard. The 3 mm bleed is added to the frame only and never affects the inner rules
- **Outer area**: keep-or-drop the outer frame, edge scale (positive extends, negative shortens), line caps and round corners. A non-zero edge scale splits the outer frame into four separate lines
- **Title area**: position (top/bottom/left/right), size, fill, rule and rule scaling. Round corners reuse the outer-area radius on just the two corners matching the position
- **Inner area**: offsets (four sides, linkable), column/row splitting with gutters, cell fill, dividers and line types (solid / dash / dots)
- **Display tab**: zoom and pan (hold Option while dragging for 1/10 speed), plus view commands
- Dialog values persist while Illustrator is running and are restored on the next run
- Japanese / English localization, with a tooltip on every control

### Flow

1. Determine the base rectangle from the selected path items, or from the active artboard
2. Adjust the values in the dialog; the preview is redrawn as you type
3. OK generates the final result; Cancel discards the preview and restores the original state

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartGridMaker";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-20";                   /* 更新日 / last updated */

/* 関連ドキュメント / Related documents */
var SCRIPT_ARTICLE_URL   = "https://note.com/dtp_tranist/n/n2b01f896c423";  /* 紹介記事（note） / article */
var SCRIPT_README_JA_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartGridMaker.md";  /* 日本語ドキュメント / Japanese readme */
var SCRIPT_README_EN_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartGridMaker.md";  /* 英語ドキュメント / English readme */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 裁ち落とし幅（mm固定） / Bleed width, always in millimeters */
var BLEED_MM = 3;

/* 内側エリアのセル塗りの濃度（CMYK K%） / Tint of the inner cell fill */
var INNER_FILL_BLACK_PCT = 15;

/* タイトル帯の塗りの濃度（CMYK K%） / Tint of the title band fill */
var TITLE_FILL_BLACK_PCT = 30;

/* フレームの塗りの濃度（CMYK K%） / Tint of the frame fill */
var FRAME_FILL_BLACK_PCT = 50;

/* ダイアログの表示位置オフセットと不透明度 / Dialog offset and opacity */
var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY  = 0.98;

/* セッション保持用のグローバルキー / Global key used to keep values within a session */
var SESSION_STATE_KEY = "__SmartGridMakerState__";

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

/* タブ内容が潰れないための最小サイズ / Minimum size so tab content is not collapsed */
var TAB_PANEL_SIZE = [320, 470];

/* 表示コマンドボタンの幅 / Width of the view command buttons */
var VIEW_BUTTON_WIDTH = 190;

/* 3行3列で上下左右を入力するときの1セル幅 / Cell width of the 3x3 directional input grid */
var FIELD_CELL_WIDTH = (getCurrentLang() === "ja") ? 92 : 116;

/**
 * ウィンドウ全体に共通のレイアウト（方向・余白・間隔）を適用します。
 *
 * @param {Window} targetWindow - 対象のウィンドウ。
 * @param {number} spacing - 要素間隔。数値以外なら既定値を使います。
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルに共通のレイアウト（方向・整列・余白・間隔）を適用します。
 *
 * @param {Panel} panel - 対象のパネル。
 * @param {number} spacing - 要素間隔。数値以外なら既定値を使います。
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = ["fill", "top"];
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * タブに共通のレイアウト（方向・整列・余白）を適用します。
 *
 * @param {Group} tab - 対象のタブ。
 * @param {number} spacing - 要素間隔。数値のときだけ適用します。
 * @returns {void}
 */
function setupTab(tab, spacing) {
    tab.orientation = "column";
    tab.alignChildren = "fill";
    tab.margins = TAB_MARGINS;
    if (typeof spacing === "number") tab.spacing = spacing;
}

/**
 * グループを横並びの行として設定します（ボタン列や入力行に使います）。
 *
 * @param {Group} group - 対象のグループ。
 * @param {string|Array<string>} alignment - グループ自身の配置指定。省略時は "left"。
 * @param {number} spacing - 要素間隔。数値以外なら既定値を使います。
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignChildren = ["left", "center"];
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * ボタンの高さを指定ピクセル分だけ詰めます。レイアウト確定後に呼びます。
 *
 * @param {Button} button - 対象のボタン。
 * @param {number} reductionPx - 詰める高さ（px）。
 * @returns {void}
 */
function trimButtonHeight(button, reductionPx) {
    try {
        button.size = [button.size.width, button.size.height - reductionPx];
    } catch (e) { }
}

// =========================================
// ローカライズ / Localization
// =========================================

/**
 * 実行環境のロケールから表示言語を判定します。
 *
 * @returns {string} "ja" または "en"。
 */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialog: {
        title: { ja: "囲み罫とグリッド", en: "Frame and Grid" }
    },

    tab: {
        artboard:  { ja: "アートボード", en: "Artboard" },
        outerArea: { ja: "外側エリア", en: "Outer Area" },
        innerArea: { ja: "内側エリア", en: "Inner Area" },
        display:   { ja: "画面表示", en: "Display" }
    },

    panel: {
        margin:    { ja: "マージン", en: "Margin" },
        outerArea: { ja: "外側エリア", en: "Outer Area" },
        lineCap:   { ja: "線端", en: "Line Caps" },
        titleArea: { ja: "タイトルエリア", en: "Title Area" },
        frame:     { ja: "フレーム", en: "Frame" },
        innerArea: { ja: "内側エリア", en: "Inner Area" },
        offset:    { ja: "オフセット", en: "Offset" },
        columns:   { ja: "列", en: "Columns" },
        rows:      { ja: "行", en: "Rows" },
        lineType:  { ja: "線の種類", en: "Line Type" },
        screen:    { ja: "表示コマンド", en: "View Commands" },
        zoomPan:   { ja: "ズームとパン", en: "Pan & Zoom" }
    },

    checkbox: {
        keepOuterFrame: { ja: "外枠を残す", en: "Keep outer frame" },
        edgeScale:      { ja: "辺の伸縮", en: "Edge scale" },
        roundCorner:    { ja: "角丸", en: "Round corners" },
        bleed:          { ja: "裁ち落とし", en: "Bleed" },
        link:           { ja: "連動", en: "Link" },
        fill:           { ja: "塗り", en: "Fill" },
        rule:           { ja: "境界線", en: "Rule" },
        divider:        { ja: "分割線", en: "Dividers" },
        preview:        { ja: "プレビュー", en: "Preview" }
    },

    radio: {
        capButt:    { ja: "なし", en: "Butt" },
        capRound:   { ja: "丸型", en: "Round" },
        capProject: { ja: "突出", en: "Project" },
        posTop:     { ja: "上", en: "Top" },
        posBottom:  { ja: "下", en: "Bottom" },
        posLeft:    { ja: "左", en: "Left" },
        posRight:   { ja: "右", en: "Right" },
        lineSolid:  { ja: "実線", en: "Solid" },
        lineDash:   { ja: "点線", en: "Dash" },
        lineDotted: { ja: "ドット点線", en: "Dots" }
    },

    field: {
        offsetTop:    { ja: "上", en: "Top" },
        offsetBottom: { ja: "下", en: "Bottom" },
        offsetLeft:   { ja: "左", en: "Left" },
        offsetRight:  { ja: "右", en: "Right" },
        width:        { ja: "幅", en: "Width" },
        titleSize:    { ja: "サイズ", en: "Size" },
        columnCount:  { ja: "列数", en: "Columns" },
        rowCount:     { ja: "行数", en: "Rows" },
        spacing:      { ja: "間隔", en: "Spacing" },
        zoom:         { ja: "ズーム", en: "Zoom" },
        panX:         { ja: "左右", en: "Pan L/R" },
        panY:         { ja: "上下", en: "Pan U/D" }
    },

    button: {
        cancel:     { ja: "キャンセル", en: "Cancel" },
        fitIn:      { ja: "アートボード全体表示", en: "Fit Artboard" },
        actualSize: { ja: "100%表示", en: "Actual Size" },
        fitAll:     { ja: "すべてを全体表示", en: "Fit All" }
    },

    tooltip: {
        keepOuterFrame: {
            ja: "基準となる長方形（またはアートボード枠）を外枠として残します。\nOFFにすると外枠は作成されません。",
            en: "Keep the base rectangle (or artboard frame) as the outer frame.\nTurn off to omit it."
        },
        outerEdgeScale: {
            ja: "外枠を4辺の線に分解し、各辺の長さを増減します。\n＋で伸ばし、−で短くします。0のときは分解しません。",
            en: "Split the outer frame into four lines and scale each edge.\nPositive extends, negative shortens. Zero keeps the rectangle."
        },
        outerRound: {
            ja: "外枠の角を丸めます。［辺の伸縮］とは併用できないため、\nONにすると辺の伸縮は0になりOFFに切り替わります。",
            en: "Round the corners of the outer frame.\nThis cannot be combined with Edge scale, so turning it on resets and disables it."
        },
        lineCap: {
            ja: "分解した4辺の線端の形状です。［辺の伸縮］が0のときは使用しません。",
            en: "Cap style of the four split edges. Unused while Edge scale is zero."
        },
        titleEnable: {
            ja: "タイトルエリアを作成します。",
            en: "Create a title area."
        },
        titleSize: {
            ja: "タイトルエリアの大きさ。上下配置では高さ、左右配置では幅になります。",
            en: "Size of the title area: height when placed top/bottom, width when placed left/right."
        },
        titleFill: {
            ja: "タイトルエリアに背景の塗りを作成します。",
            en: "Add a background fill behind the title area."
        },
        titleRule: {
            ja: "タイトルエリアと本文の境界に線を引きます。",
            en: "Draw a rule between the title area and the body."
        },
        titleEdgeScale: {
            ja: "境界線の長さを増減します。＋で伸ばし、−で短くします。",
            en: "Scale the length of the rule. Positive extends, negative shortens."
        },
        frameEnable: {
            ja: "アートボードの外周にフレームを作成します。\nアートボードを基準に開始したときのみ使用できます。",
            en: "Create a frame around the artboard.\nAvailable only when the artboard is used as the base."
        },
        frameWidth: {
            ja: "フレームの太さ。",
            en: "Thickness of the frame."
        },
        bleed: {
            ja: "フレームの外側に裁ち落とし3mmを加えます。裁ち落としはフレームにのみ適用されます。",
            en: "Extend the frame outward by a 3 mm bleed. Bleed applies to the frame only."
        },
        frameRound: {
            ja: "フレームの内側（穴側）の角を丸めます。",
            en: "Round the inner corners of the frame."
        },
        link: {
            ja: "［上］の値を［下］［左］［右］にも反映します。",
            en: "Apply the Top value to Bottom, Left and Right as well."
        },
        gutter: {
            ja: "セルどうしの間隔。列数（行数）が2以上のときに使用します。",
            en: "Gap between cells. Used when there are two or more columns (rows)."
        },
        innerFill: {
            ja: "分割した各セルに背景の塗りを作成します。",
            en: "Add a background fill to each split cell."
        },
        innerDivider: {
            ja: "列／行の境界に線を引きます。間隔があるときはその中心に引かれます。",
            en: "Draw a rule at each column/row boundary, centered in the gap when one is set."
        },
        preview: {
            ja: "設定内容をカンバス上に仮描画します。キャンセルすると破棄されます。",
            en: "Draw the current settings on the canvas. Discarded when you cancel."
        },
        viewSlider: {
            ja: "Option（Alt）を押しながらドラッグすると1/10の速度で動きます。",
            en: "Hold Option (Alt) while dragging to move at 1/10 speed."
        }
    },

    message: {
        viewControlUnavailable: {
            ja: "ズームとパンは利用できません。",
            en: "Zoom and pan are unavailable."
        }
    }
};

/**
 * ラベル（入力欄の見出し、および入力欄を伴うチェックボックス）の末尾に全角コロンを補います。すでに付いていればそのまま返します。
 *
 * @param {string} text - ラベル文字列。
 * @returns {string} コロンを補ったラベル文字列。
 */
function withColon(text) {
    if (!text) return "";
    var lastChar = text.charAt(text.length - 1);
    if (lastChar === "\uFF1A" || lastChar === ":") return text;
    return text + ((lang === "ja") ? "\uFF1A" : ":");
}

/**
 * コントロールにローカライズ済みのツールチップを設定します。
 *
 * @param {Object} control - 対象のUIコントロール（Checkbox / EditText / Button など）。
 * @param {Object} entry - LABELS 内のラベル定義（ja / en を持つオブジェクト）。
 * @returns {void}
 */
function setTooltip(control, entry) {
    try {
        var text = getLabel(entry);
        if (control && text) control.helpTip = text;
    } catch (e) { }
}

/**
 * ラベル定義から現在の表示言語に対応する文字列を取り出します。
 * 該当言語がなければ日本語、それもなければ空文字を返します。
 *
 * @param {Object} entry - LABELS 内のラベル定義（ja / en を持つオブジェクト）。
 * @returns {string} ローカライズ済みのラベル文字列。
 */
function getLabel(entry) {
    if (!entry) return "";
    if (entry[lang]) return entry[lang];
    if (entry.ja) return entry.ja;
    return "";
}

// =========================================
// 単位 / Units (rulerType)
// =========================================

/* 単位コード → ラベル / Unit code to label. Code 5 is Q/H; rulerType reports H */
var UNIT_LABEL_MAP = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/**
 * 単位コードに対応する「1単位あたりのポイント数」を返します。
 *
 * @param {number} unitCode - Illustrator の rulerType 単位コード。
 * @returns {number} 1単位あたりのポイント数。未知のコードでは 1。
 */
function getPointsPerUnit(unitCode) {
    switch (unitCode) {
        case 0: return 72.0;                  // in
        case 1: return 72.0 / 25.4;           // mm
        case 2: return 1.0;                   // pt
        case 3: return 12.0;                  // pica
        case 4: return 72.0 / 2.54;           // cm
        case 5: return 72.0 / 25.4 * 0.25;    // Q / H (0.25mm)
        case 6: return 1.0;                   // px（Illustratorの内部単位はpt基準）
        case 7: return 72.0 * 12.0;           // ft/in
        case 8: return 72.0 / 25.4 * 1000.0;  // m
        case 9: return 72.0 * 36.0;           // yd
        case 10: return 72.0 * 12.0;          // ft
        default: return 1.0;
    }
}

/**
 * 現在の定規単位コードを取得します。取得できない場合は pt を表すコードを返します。
 *
 * @returns {number} rulerType の単位コード。
 */
function getCurrentRulerUnitCode() {
    try {
        return app.preferences.getIntegerPreference("rulerType");
    } catch (e) {
        return 2; // pt
    }
}

/**
 * 現在の定規単位の表示ラベルを返します。
 *
 * @returns {string} 単位ラベル（"mm" など）。未知のコードでは "pt"。
 */
function getCurrentRulerUnitLabel() {
    return UNIT_LABEL_MAP[getCurrentRulerUnitCode()] || "pt";
}

/**
 * 現在の定規単位の1単位が何ポイントかを返します。
 *
 * @returns {number} 1単位あたりのポイント数（0以下にはなりません）。
 */
function getCurrentRulerPtFactor() {
    var factor = getPointsPerUnit(getCurrentRulerUnitCode());
    return (factor > 0) ? factor : 1;
}

/**
 * 裁ち落とし幅をポイント単位で返します。
 *
 * @returns {number} 裁ち落とし幅（pt）。
 */
function getBleedPt() {
    return (72.0 / 25.4) * BLEED_MM;
}

// =========================================
// セッション保持 / Session-persistent UI state
// =========================================
// Illustrator が起動している間だけ、前回のダイアログ値を保持する。
// #targetengine 指定により $.global がスクリプト実行をまたいで残る。

$.global[SESSION_STATE_KEY] = $.global[SESSION_STATE_KEY] || {};

/**
 * 前回のダイアログ値をセッション領域から読み出します。
 *
 * @returns {SessionState} 保存済みの設定値。未保存なら空のオブジェクト。
 */
function loadSessionState() {
    try { return $.global[SESSION_STATE_KEY] || {}; } catch (e) { return {}; }
}

/**
 * ダイアログ値をセッション領域へ書き込みます。
 *
 * @param {SessionState} state - 保存する設定値。
 * @returns {void}
 */
function saveSessionState(state) {
    try { $.global[SESSION_STATE_KEY] = state || {}; } catch (e) { }
}

/* =========================================
 * ViewControl / ズームとパンのユーティリティ（切り出し可能）
 *
 * UI（スライダーのみ）:
 *   ズーム <===== slider =====>
 *   左右   <===== slider =====>
 *   上下   <===== slider =====>
 *
 * - 基準の中心：アクティブなアートボードの中心
 * - パン量：基準の中心からの相対値
 * - ズームしてもパン量は維持される
 * - Option(Alt) を押しながらドラッグで 1/10 の速度
 * - パンの可動域：アートボード幅／高さの半分（安全側にクランプ）
 *
 * 使い方:
 *   var viewCtl = ViewControl.create(doc);
 *   viewCtl.buildUI(parentGroup, { labelWidth: 58, sliderWidth: 200 });
 *   viewCtl.restore();   // キャンセル時
 * ========================================= */

/**
 * ズームとパンをまとめて扱うモジュール。スライダーUIの生成と、
 * キャンセル時に元の表示へ戻す処理を提供します。
 *
 * @returns {Object} create(doc) を持つモジュールオブジェクト。
 */
var ViewControl = (function () {

    /**
     * 値を最小値と最大値の範囲に収めます。
     *
     * @param {number} value - 対象の値。
     * @param {number} minValue - 下限。
     * @param {number} maxValue - 上限。
     * @returns {number} 範囲内に収めた値。
     */
    function clamp(value, minValue, maxValue) {
        return (value < minValue) ? minValue : (value > maxValue) ? maxValue : value;
    }

    /**
     * Option（Alt）キーが押されているかを判定します。
     *
     * @returns {boolean} 押されていれば true。
     */
    function isAltKeyDown() {
        try {
            return !!(ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey);
        } catch (e) { }
        return false;
    }

    /**
     * スライダーの変化量を適用します。Option 押下中は移動量を1/10に落とします。
     *
     * @param {Slider} slider - 対象のスライダー。
     * @param {Object} sliderState - 直前の生値と実効値を保持する状態オブジェクト。
     * @param {Function} applyFn - 実効値を受け取って表示へ反映する関数。
     * @returns {void}
     */
    function applySliderAltFine(slider, sliderState, applyFn) {
        try {
            if (!slider || !sliderState || typeof applyFn !== "function") return;

            var rawValue = Number(slider.value);
            if (isNaN(rawValue)) rawValue = 0;

            if (sliderState.raw == null || sliderState.effective == null) {
                sliderState.raw = rawValue;
                sliderState.effective = rawValue;
            }

            if (isAltKeyDown()) {
                var delta = rawValue - Number(sliderState.raw);
                var effective = Number(sliderState.effective) + delta * 0.1;
                sliderState.raw = rawValue;
                sliderState.effective = effective;
                try { slider.value = effective; } catch (e) { }
                applyFn(effective);
            } else {
                sliderState.raw = rawValue;
                sliderState.effective = rawValue;
                applyFn(rawValue);
            }
        } catch (e) { }
    }

    /**
     * アクティブなアートボードの中心座標を返します。
     * 取得できない場合は現在の表示中心、それも不明なら原点を返します。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {Object} view - ドキュメントのビュー。
     * @returns {Array<number>} 中心座標 [x, y]。
     */
    function getActiveArtboardCenter(doc, view) {
        try {
            var index = doc.artboards.getActiveArtboardIndex();
            var rect = doc.artboards[index].artboardRect; // [L, T, R, B]
            return [
                rect[0] + (rect[2] - rect[0]) / 2,
                rect[1] + (rect[3] - rect[1]) / 2
            ];
        } catch (e) { }
        try { return (view && view.centerPoint) ? view.centerPoint : [0, 0]; } catch (e) { }
        return [0, 0];
    }

    /**
     * パンの可動域。
     *
     * @typedef {Object} PanRange
     * @property {number} maxX - 左右方向の可動域（pt）。
     * @property {number} maxY - 上下方向の可動域（pt）。
     */
    /**
     * パンの可動域を返します。可動域はアートボードの幅・高さの半分です。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {PanRange} 可動域（pt）。
     */
    function getPanRangePt(doc) {
        try {
            var index = doc.artboards.getActiveArtboardIndex();
            var rect = doc.artboards[index].artboardRect;
            var artboardWidth = Math.abs(rect[2] - rect[0]);
            var artboardHeight = Math.abs(rect[1] - rect[3]);
            var maxX = Math.round(artboardWidth / 2);
            var maxY = Math.round(artboardHeight / 2);
            if (!maxX || maxX < 100) maxX = 100;
            if (!maxY || maxY < 100) maxY = 100;
            if (maxX > 50000) maxX = 50000;
            if (maxY > 50000) maxY = 50000;
            return { maxX: maxX, maxY: maxY };
        } catch (e) { }
        return { maxX: 2000, maxY: 2000 };
    }

    /**
     * Illustrator が受け付けるズーム倍率の範囲に収めます。
     *
     * @param {number} zoomFactor - ズーム倍率（1.0 が100%）。
     * @returns {number} 範囲内に収めたズーム倍率。
     */
    function clampZoomFactor(zoomFactor) {
        if (zoomFactor < 0.0313) zoomFactor = 0.0313;
        if (zoomFactor > 640.0) zoomFactor = 640.0;
        return zoomFactor;
    }

    /**
     * ドキュメントに紐づくビュー操作オブジェクトを生成します。
     * 生成時点のズームと表示中心を控え、restore() で戻せるようにします。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {Object} ビュー操作オブジェクト。
     */
    function create(doc) {
        var viewCtl = {};

        viewCtl.doc = doc;
        viewCtl.view = null;

        viewCtl.originalZoom = null;
        viewCtl.originalCenter = null;

        viewCtl.panXPt = 0;
        viewCtl.panYPt = 0; // UIでは正の値が「下」
        viewCtl.panRange = { maxX: 2000, maxY: 2000 };

        viewCtl.zoomSlider = null;
        viewCtl.panXSlider = null;
        viewCtl.panYSlider = null;

        try {
            viewCtl.view = doc.views[0];
            viewCtl.originalZoom = viewCtl.view.zoom;
            viewCtl.originalCenter = viewCtl.view.centerPoint;
        } catch (e) { }

        /**
         * 現在のアートボードからパンの可動域を再計算して保持します。
         *
         * @returns {PanRange} 更新後の可動域（pt）。
         */
        viewCtl.refreshPanRange = function () {
            try { viewCtl.panRange = getPanRangePt(viewCtl.doc); } catch (e) { }
            return viewCtl.panRange;
        };

        /**
         * アートボード中心にパン量を加えた位置へ表示中心を移動します。
         *
         * @returns {void}
         */
        viewCtl.applyCenter = function () {
            try {
                if (!viewCtl.view) return;
                var center = getActiveArtboardCenter(viewCtl.doc, viewCtl.view);
                var centerX = center[0] + Number(viewCtl.panXPt || 0);
                // Illustrator は +Y が上、UI は正の値が下なので引く
                var centerY = center[1] - Number(viewCtl.panYPt || 0);
                viewCtl.view.centerPoint = [centerX, centerY];
                app.redraw();
            } catch (e) { }
        };

        /**
         * ズーム倍率をパーセント指定で設定します。パン量は維持されます。
         *
         * @param {number} percent - 設定するズーム率（%）。
         * @param {number} minPercent - 下限のズーム率（%）。
         * @param {number} maxPercent - 上限のズーム率（%）。
         * @returns {void}
         */
        viewCtl.setZoomPercent = function (percent, minPercent, maxPercent) {
            try {
                if (!viewCtl.view) return;
                var value = Number(percent);
                if (isNaN(value)) return;
                value = clamp(Math.round(value), minPercent, maxPercent);
                viewCtl.view.zoom = clampZoomFactor(value / 100.0);
                viewCtl.applyCenter(); // パン量は維持
            } catch (e) { }
        };

        /**
         * 左右方向のパン量を設定して表示を更新します。
         *
         * @param {number} value - パン量（pt）。可動域を超える分は丸められます。
         * @returns {void}
         */
        viewCtl.setPanX = function (value) {
            try {
                var amount = Number(value);
                if (isNaN(amount)) amount = 0;
                var range = viewCtl.refreshPanRange();
                viewCtl.panXPt = clamp(Math.round(amount), -range.maxX, range.maxX);
                viewCtl.applyCenter();
            } catch (e) { }
        };

        /**
         * 上下方向のパン量を設定して表示を更新します。
         *
         * @param {number} value - パン量（pt）。正の値が下方向です。
         * @returns {void}
         */
        viewCtl.setPanY = function (value) {
            try {
                var amount = Number(value);
                if (isNaN(amount)) amount = 0;
                var range = viewCtl.refreshPanRange();
                viewCtl.panYPt = clamp(Math.round(amount), -range.maxY, range.maxY);
                viewCtl.applyCenter();
            } catch (e) { }
        };

        /**
         * 生成時点のズームと表示中心へ戻し、パン量を0にします。
         *
         * @returns {void}
         */
        viewCtl.restore = function () {
            try {
                if (viewCtl.view && viewCtl.originalZoom != null && viewCtl.originalCenter != null) {
                    viewCtl.view.zoom = viewCtl.originalZoom;
                    viewCtl.view.centerPoint = viewCtl.originalCenter;
                    app.redraw();
                }
            } catch (e) { }
            viewCtl.panXPt = 0;
            viewCtl.panYPt = 0;
        };

        /**
         * ズーム／左右／上下の3本のスライダーを親グループに組み立てます。
         *
         * @param {Group} parent - スライダーを追加する親グループ。
         * @param {Object} options - ラベル幅・スライダー幅・ズーム上下限の指定。
         * @returns {Object} 自分自身のビュー操作オブジェクト。
         */
        viewCtl.buildUI = function (parent, options) {
            options = options || {};
            var labelWidth = (typeof options.labelWidth === "number") ? options.labelWidth : 58;
            var sliderWidth = (typeof options.sliderWidth === "number") ? options.sliderWidth : 200;
            var minZoomPercent = (typeof options.minZoomPercent === "number") ? options.minZoomPercent : 10;
            var maxZoomPercent = (typeof options.maxZoomPercent === "number") ? options.maxZoomPercent : 1600;

            var initialZoomPercent = 100;
            try {
                if (viewCtl.originalZoom != null) {
                    initialZoomPercent = Math.round(Number(viewCtl.originalZoom) * 100);
                }
            } catch (e) { }
            if (!initialZoomPercent || initialZoomPercent < minZoomPercent) initialZoomPercent = 100;

            var panRange = viewCtl.refreshPanRange();

            /**
             * ラベルとスライダーを並べるための1行を追加します。
             *
             * @param {string} labelText - 行の先頭に置くラベル文字列。
             * @returns {Group} 追加した行グループ。
             */
            function addSliderRow(labelText) {
                var row = parent.add("group");
                setupRow(row);
                var label = row.add("statictext", undefined, labelText);
                try { label.preferredSize.width = labelWidth; } catch (e) { }
                return row;
            }

            // ズーム
            var zoomRow = addSliderRow(withColon(getLabel(LABELS.field.zoom)));
            viewCtl.zoomSlider = zoomRow.add("slider", undefined, initialZoomPercent, minZoomPercent, maxZoomPercent);
            try { viewCtl.zoomSlider.preferredSize.width = sliderWidth; } catch (e) { }
            setTooltip(viewCtl.zoomSlider, LABELS.tooltip.viewSlider);
            var zoomSliderState = { raw: null, effective: null };
            /**
             * ズームスライダーの操作を表示へ反映します。
             *
             * @returns {void}
             */
            viewCtl.zoomSlider.onChanging = function () {
                applySliderAltFine(this, zoomSliderState, function (value) {
                    viewCtl.setZoomPercent(value, minZoomPercent, maxZoomPercent);
                });
            };

            // 左右
            var panXRow = addSliderRow(withColon(getLabel(LABELS.field.panX)));
            viewCtl.panXSlider = panXRow.add("slider", undefined, 0, -panRange.maxX, panRange.maxX);
            try { viewCtl.panXSlider.preferredSize.width = sliderWidth; } catch (e) { }
            setTooltip(viewCtl.panXSlider, LABELS.tooltip.viewSlider);
            var panXSliderState = { raw: null, effective: null };
            /**
             * 左右パンスライダーの操作を表示へ反映します。
             *
             * @returns {void}
             */
            viewCtl.panXSlider.onChanging = function () {
                applySliderAltFine(this, panXSliderState, function (value) { viewCtl.setPanX(value); });
            };

            // 上下
            var panYRow = addSliderRow(withColon(getLabel(LABELS.field.panY)));
            viewCtl.panYSlider = panYRow.add("slider", undefined, 0, -panRange.maxY, panRange.maxY);
            try { viewCtl.panYSlider.preferredSize.width = sliderWidth; } catch (e) { }
            setTooltip(viewCtl.panYSlider, LABELS.tooltip.viewSlider);
            var panYSliderState = { raw: null, effective: null };
            /**
             * 上下パンスライダーの操作を表示へ反映します。
             *
             * @returns {void}
             */
            viewCtl.panYSlider.onChanging = function () {
                applySliderAltFine(this, panYSliderState, function (value) { viewCtl.setPanY(value); });
            };

            return viewCtl;
        };

        return viewCtl;
    }

    return { create: create };
})();

// =========================================
// メイン処理 / Main
// =========================================
/**
 * メイン処理。基準矩形を決め、ダイアログを表示し、OK で囲み罫とグリッドを生成します。
 *
 * @returns {void}
 */
(function () {

    // --- 準備とチェック / Setup and guards ---
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;

    var viewCtl = null;
    try { viewCtl = ViewControl.create(doc); } catch (e) { viewCtl = null; }

    // 選択中のパスアイテムを基準矩形として集める
    var baseRectItems = [];
    /**
     * 選択中のパスアイテムを基準矩形の候補として集めます。
     *
     * @returns {void}
     */
    (function collectBaseRectItems() {
        var selection = doc.selection;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "PathItem") baseRectItems.push(selection[i]);
        }
    })();

    // 選択した長方形の初期状態を統一（塗りなし／黒1pt）
    for (var initIndex = 0; initIndex < baseRectItems.length; initIndex++) {
        try {
            var baseItem = baseRectItems[initIndex];
            baseItem.filled = false;
            baseItem.stroked = true;
            baseItem.strokeWidth = 1;
            baseItem.strokeColor = makeBlackTintFill(100);
        } catch (e) { }
    }

    // 選択がない場合は、現在のアートボードを基準にする
    var isArtboardBase = (baseRectItems.length === 0);
    var artboardBaseRect = null;
    var isBleedEnabled = false;

    // プレビュー用に生成した一時オブジェクト
    var previewItems = [];

    // --- 互換性：StrokeCap が未定義の環境対策 / Fallback when StrokeCap is missing ---
    if (typeof StrokeCap === "undefined") {
        StrokeCap = {
            BUTTENDCAP: 0,
            ROUNDENDCAP: 1,
            PROJECTINGENDCAP: 2
        };
    }

    // =========================================
    // 基準矩形（アートボード基準） / Artboard base rectangle
    // =========================================

    /**
     * アートボード基準で作成した一時矩形を破棄し、基準矩形の一覧からも外します。
     *
     * @returns {void}
     */
    function disposeArtboardBaseRect() {
        if (!isArtboardBase || !artboardBaseRect) return;

        // baseRectItems から外す（remove 後に触ると Error 45 になるため）
        try {
            for (var i = baseRectItems.length - 1; i >= 0; i--) {
                if (baseRectItems[i] === artboardBaseRect) baseRectItems.splice(i, 1);
            }
        } catch (e) { }

        try { artboardBaseRect.remove(); } catch (e) { }
        artboardBaseRect = null;
    }

    /**
     * アートボードをマージン分だけ内側に狭めた一時矩形を作り直します。
     * 裁ち落としはフレームにのみ適用するため、ここでは考慮しません。
     *
     * @param {number} marginTopPt - 上マージン（pt）。
     * @param {number} marginRightPt - 右マージン（pt）。
     * @param {number} marginBottomPt - 下マージン（pt）。
     * @param {number} marginLeftPt - 左マージン（pt）。
     * @returns {void}
     */
    function rebuildArtboardBaseRect(marginTopPt, marginRightPt, marginBottomPt, marginLeftPt) {
        if (!isArtboardBase) return;
        try {
            var artboardIndex = doc.artboards.getActiveArtboardIndex();
            var artboardRect = doc.artboards[artboardIndex].artboardRect; // [L, T, R, B]
            var outerLeft = artboardRect[0];
            var outerTop = artboardRect[1];
            var outerRight = artboardRect[2];
            var outerBottom = artboardRect[3];

            var marginTop = (marginTopPt > 0) ? marginTopPt : 0;
            var marginRight = (marginRightPt > 0) ? marginRightPt : 0;
            var marginBottom = (marginBottomPt > 0) ? marginBottomPt : 0;
            var marginLeft = (marginLeftPt > 0) ? marginLeftPt : 0;

            var insetLeft = outerLeft + marginLeft;
            var insetTop = outerTop - marginTop;
            var insetRight = outerRight - marginRight;
            var insetBottom = outerBottom + marginBottom;

            var rectWidth = insetRight - insetLeft;
            var rectHeight = insetTop - insetBottom;
            if (!(rectWidth > 0) || !(rectHeight > 0)) {
                // マージンが大きすぎる場合はマージン0として扱う
                insetLeft = outerLeft;
                insetTop = outerTop;
                insetRight = outerRight;
                insetBottom = outerBottom;
                rectWidth = insetRight - insetLeft;
                rectHeight = insetTop - insetBottom;
            }
            if (!(rectWidth > 0) || !(rectHeight > 0)) return;

            disposeArtboardBaseRect();

            artboardBaseRect = doc.activeLayer.pathItems.rectangle(insetTop, insetLeft, rectWidth, rectHeight);
            artboardBaseRect.stroked = true;
            artboardBaseRect.filled = false;
            try { artboardBaseRect.strokeColor = makeBlackTintFill(100); } catch (e) { }
            try { artboardBaseRect.strokeWidth = 1; } catch (e) { }

            baseRectItems.push(artboardBaseRect);
        } catch (e) { }
    }

    // =========================================
    // 色と角丸のヘルパー / Color and corner helpers
    // =========================================

    /**
     * K版のみを使ったCMYKのグレーを作成します。
     *
     * @param {number} blackPercent - K版の濃度（%）。
     * @returns {CMYKColor} 生成した色。
     */
    function makeBlackTintFill(blackPercent) {
        var color = new CMYKColor();
        color.cyan = 0;
        color.magenta = 0;
        color.yellow = 0;
        color.black = blackPercent;
        return color;
    }

    /**
     * 角丸ライブエフェクト（Adobe Round Corners）の定義XMLを組み立てます。
     *
     * @param {number} radiusPt - 角丸の半径（pt）。
     * @returns {string} ライブエフェクトのXML文字列。
     */
    function buildRoundCornersEffectXML(radiusPt) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radiusPt);
    }

    /**
     * アイテムに角丸ライブエフェクトを適用します。半径が0以下なら何もしません。
     *
     * @param {PathItem} item - 対象のアイテム。
     * @param {number} radiusPt - 角丸の半径（pt）。
     * @returns {void}
     */
    function applyRoundCornersEffect(item, radiusPt) {
        try {
            if (!item) return;
            if (!(radiusPt > 0)) return;
            item.applyEffect(buildRoundCornersEffectXML(radiusPt));
        } catch (e) { }
    }

    /**
     * 4点の長方形について、指定した2角だけをアンカー操作で角丸に置き換えます（その場で書き換え）。
     *
     * @param {PathItem} rect - 対象の閉じた長方形パス。
     * @param {string} cornerPair - 角丸にする2角の位置（"TOP" / "BOTTOM" / "LEFT" / "RIGHT"）。
     * @param {number} radiusPt - 角丸の半径（pt）。辺の長さに応じて自動的に抑えられます。
     * @returns {boolean} 角丸にできたら true。
     */
    function roundRectCornerPairInPlace(rect, cornerPair, radiusPt) {
        if (!rect || rect.typename !== "PathItem" || !rect.closed) return false;

        var requestedRadius = Number(radiusPt);
        if (isNaN(requestedRadius) || requestedRadius <= 0) return false;

        var bounds = rect.geometricBounds; // [L, T, R, B]
        var boundsLeft = bounds[0];
        var boundsTop = bounds[1];
        var boundsRight = bounds[2];
        var boundsBottom = bounds[3];
        var rectWidth = boundsRight - boundsLeft;
        var rectHeight = boundsTop - boundsBottom;
        if (!(rectWidth > 0) || !(rectHeight > 0)) return false;

        var radius = Math.min(requestedRadius, rectWidth / 2, rectHeight / 2);
        if (!(radius > 0)) return false;

        // 円弧を3次ベジェで近似するときのハンドル長 / Bezier handle length for a quarter circle
        var handleLength = radius * 0.5522847498307936;
        var pointTypes = (typeof PointType !== "undefined") ? PointType : { CORNER: 0, SMOOTH: 1 };

        var roundTopLeft = (cornerPair === "TOP" || cornerPair === "LEFT");
        var roundTopRight = (cornerPair === "TOP" || cornerPair === "RIGHT");
        var roundBottomRight = (cornerPair === "BOTTOM" || cornerPair === "RIGHT");
        var roundBottomLeft = (cornerPair === "BOTTOM" || cornerPair === "LEFT");

        var pointSpecs = [];
        /**
         * アンカーと方向線1組ぶんの情報を組み立てて一覧に追加します。
         *
         * @param {number} anchorX - アンカーのX座標。
         * @param {number} anchorY - アンカーのY座標。
         * @param {number} leftDirectionX - 進入方向ハンドルのX座標。null ならアンカーと同じ。
         * @param {number} leftDirectionY - 進入方向ハンドルのY座標。null ならアンカーと同じ。
         * @param {number} rightDirectionX - 進出方向ハンドルのX座標。null ならアンカーと同じ。
         * @param {number} rightDirectionY - 進出方向ハンドルのY座標。null ならアンカーと同じ。
         * @param {number} pointType - アンカーの種類（コーナー／スムーズ）。
         * @returns {void}
         */
        function addPoint(anchorX, anchorY, leftDirectionX, leftDirectionY, rightDirectionX, rightDirectionY, pointType) {
            pointSpecs.push({
                anchor: [anchorX, anchorY],
                leftDirection: (leftDirectionX == null) ? [anchorX, anchorY] : [leftDirectionX, leftDirectionY],
                rightDirection: (rightDirectionX == null) ? [anchorX, anchorY] : [rightDirectionX, rightDirectionY],
                pointType: pointType
            });
        }

        // 左上（上辺側）から開始
        if (roundTopLeft) {
            addPoint(boundsLeft + radius, boundsTop, boundsLeft + radius - handleLength, boundsTop, null, null, pointTypes.SMOOTH);
        } else {
            addPoint(boundsLeft, boundsTop, null, null, null, null, pointTypes.CORNER);
        }

        // 右上
        if (roundTopRight) {
            addPoint(boundsRight - radius, boundsTop, null, null, boundsRight - radius + handleLength, boundsTop, pointTypes.SMOOTH);
            addPoint(boundsRight, boundsTop - radius, boundsRight, boundsTop - radius + handleLength, null, null, pointTypes.SMOOTH);
        } else {
            addPoint(boundsRight, boundsTop, null, null, null, null, pointTypes.CORNER);
        }

        // 右下
        if (roundBottomRight) {
            addPoint(boundsRight, boundsBottom + radius, null, null, boundsRight, boundsBottom + radius - handleLength, pointTypes.SMOOTH);
            addPoint(boundsRight - radius, boundsBottom, boundsRight - radius + handleLength, boundsBottom, null, null, pointTypes.SMOOTH);
        } else {
            addPoint(boundsRight, boundsBottom, null, null, null, null, pointTypes.CORNER);
        }

        // 左下
        if (roundBottomLeft) {
            addPoint(boundsLeft + radius, boundsBottom, null, null, boundsLeft + radius - handleLength, boundsBottom, pointTypes.SMOOTH);
            addPoint(boundsLeft, boundsBottom + radius, boundsLeft, boundsBottom + radius - handleLength, null, null, pointTypes.SMOOTH);
        } else {
            addPoint(boundsLeft, boundsBottom, null, null, null, null, pointTypes.CORNER);
        }

        // 閉じる：左上（左辺側）
        if (roundTopLeft) {
            addPoint(boundsLeft, boundsTop - radius, null, null, boundsLeft, boundsTop - radius + handleLength, pointTypes.SMOOTH);
        }

        var anchorList = [];
        for (var specIndex = 0; specIndex < pointSpecs.length; specIndex++) {
            anchorList.push(pointSpecs[specIndex].anchor);
        }
        rect.setEntirePath(anchorList);
        rect.closed = true;

        for (var pointIndex = 0; pointIndex < rect.pathPoints.length; pointIndex++) {
            rect.pathPoints[pointIndex].leftDirection = pointSpecs[pointIndex].leftDirection;
            rect.pathPoints[pointIndex].rightDirection = pointSpecs[pointIndex].rightDirection;
            try { rect.pathPoints[pointIndex].pointType = pointSpecs[pointIndex].pointType; } catch (e) { }
        }
        return true;
    }

    /**
     * アイテムがタイトル帯の塗りとして生成されたものかを判定します。角丸の適用対象を限定するために使います。
     *
     * @param {PathItem} item - 判定するアイテム。
     * @returns {boolean} タイトル帯なら true。
     */
    function isTitleBandItem(item) {
        if (!item) return false;
        try { if (item.note === "__TitleFill__") return true; } catch (e) { }
        try { if (item.name === "__TitleFill__") return true; } catch (e) { }
        return false;
    }

    /**
     * タイトル帯に、配置位置に対応する2角だけ角丸を適用します。
     *
     * @param {PathItem} titleRect - タイトル帯の矩形。
     * @param {string} titlePositionKey - タイトルの配置（"top" / "bottom" / "left" / "right"）。
     * @param {number} radiusInUnit - 角丸の半径（定規単位）。
     * @returns {boolean} 角丸を適用できたら true。
     */
    function applyTitleAreaCornerRounding(titleRect, titlePositionKey, radiusInUnit) {
        if (!titleRect || titleRect.typename !== "PathItem") return false;
        if (!isTitleBandItem(titleRect)) return false;

        var radiusValue = Number(radiusInUnit);
        if (isNaN(radiusValue) || radiusValue <= 0) return false;

        var radiusPt = radiusValue * getCurrentRulerPtFactor();

        var cornerPair = null;
        if (titlePositionKey === "top") cornerPair = "TOP";
        else if (titlePositionKey === "right") cornerPair = "RIGHT";
        else if (titlePositionKey === "bottom") cornerPair = "BOTTOM";
        else if (titlePositionKey === "left") cornerPair = "LEFT";
        else return false;

        return roundRectCornerPairInPlace(titleRect, cornerPair, radiusPt);
    }

    // =========================================
    // 既定値の算出 / Default values
    // =========================================

    /**
     * 内側オフセットの初期値を求めます。基準矩形の(幅+高さ)/40を定規単位で丸めた値です。
     *
     * @returns {number} 内側オフセットの初期値（定規単位）。求められない場合は 0。
     */
    function calcDefaultInnerOffset() {
        try {
            if (!baseRectItems || baseRectItems.length === 0) return 0;
            var bounds = baseRectItems[0].geometricBounds; // [L, T, R, B]
            var widthPt = bounds[2] - bounds[0];
            var heightPt = bounds[1] - bounds[3];
            if (!(widthPt > 0) || !(heightPt > 0)) return 0;

            var offsetInUnit = ((widthPt + heightPt) / 40) / getCurrentRulerPtFactor();
            if (!(offsetInUnit > 0)) return 0;

            // mm/pt のような大きい値は10単位に丸め、cm/inch のような小さい値は小数第1位まで
            if (offsetInUnit >= 20) return Math.round(offsetInUnit / 10) * 10;
            return Math.round(offsetInUnit * 10) / 10;
        } catch (e) {
            return 0;
        }
    }

    /**
     * タイトルエリアの初期サイズを求めます。基準となる高さの1/5です。
     *
     * @returns {number} タイトルサイズの初期値（定規単位）。求められない場合は 10。
     */
    function calcDefaultTitleSize() {
        try {
            var bounds;
            if (baseRectItems && baseRectItems.length > 0) {
                bounds = baseRectItems[0].geometricBounds; // [L, T, R, B]
            } else {
                var artboardIndex = doc.artboards.getActiveArtboardIndex();
                bounds = doc.artboards[artboardIndex].artboardRect;
            }

            var heightPt = Math.abs(bounds[1] - bounds[3]);
            if (!(heightPt > 0)) return 10;

            var sizeInUnit = (heightPt / getCurrentRulerPtFactor()) / 5;
            if (!(sizeInUnit > 0)) return 10;

            var rounded = Math.round(sizeInUnit * 10) / 10;
            if (Math.abs(rounded - Math.round(rounded)) < 1e-6) rounded = Math.round(rounded);
            if (rounded < 0.1) rounded = 0.1;
            return rounded;
        } catch (e) { }
        return 10;
    }

    // アートボード基準ならここで基準矩形を作成
    if (isArtboardBase) rebuildArtboardBaseRect(0, 0, 0, 0);

    // =========================================
    // ダイアログの骨格 / Dialog skeleton
    // =========================================
    var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    setupWindow(dialog);

    try { dialog.opacity = DIALOG_OPACITY; } catch (e) { }
    /**
     * ダイアログ表示時に、既定位置から指定量だけずらして配置します。
     *
     * @returns {void}
     */
    dialog.onShow = function () {
        dialog.location = [dialog.location[0] + DIALOG_OFFSET_X, dialog.location[1] + DIALOG_OFFSET_Y];
    };

    // 4タブ：マージン／外側エリア／内側エリア／画面表示
    var tabPanel = dialog.add("tabbedpanel");
    tabPanel.alignChildren = ["fill", "top"];
    tabPanel.alignment = ["fill", "top"];
    tabPanel.margins = [5, 20, 0, 0];
    try {
        tabPanel.minimumSize = TAB_PANEL_SIZE;
        tabPanel.preferredSize = TAB_PANEL_SIZE;
    } catch (e) { }

    var artboardTab = tabPanel.add("tab", undefined, getLabel(LABELS.tab.artboard));
    var outerTab = tabPanel.add("tab", undefined, getLabel(LABELS.tab.outerArea));
    var innerTab = tabPanel.add("tab", undefined, getLabel(LABELS.tab.innerArea));
    var displayTab = tabPanel.add("tab", undefined, getLabel(LABELS.tab.display));

    setupTab(artboardTab, PANEL_SPACING);
    setupTab(outerTab, PANEL_SPACING);
    setupTab(innerTab, PANEL_SPACING);
    setupTab(displayTab, PANEL_SPACING);

    /**
     * タブの中身を縦積みするための列グループを追加します。
     *
     * @param {Group} tab - 追加先のタブ。
     * @returns {Group} 追加した列グループ。
     */
    function addTabColumn(tab) {
        var column = tab.add("group");
        column.orientation = "column";
        column.alignChildren = ["fill", "top"];
        column.spacing = PANEL_SPACING;
        return column;
    }

    var artboardColumn = addTabColumn(artboardTab);
    var outerColumn = addTabColumn(outerTab);
    var innerColumn = addTabColumn(innerTab);
    var displayColumn = addTabColumn(displayTab);

    // =========================================
    // UI部品のハンドル / UI control handles
    // =========================================

    // マージン
    var marginPanel;
    var editMarginTop, editMarginBottom, editMarginLeft, editMarginRight;
    var chkMarginLink;
    var isSyncingMargins = false;
    var applyMarginLinkState;

    // 外側エリア
    var outerPanel;
    var chkKeepOuterFrame;
    var chkOuterEdgeScale, editOuterEdgeScale, outerEdgeScaleGroup;
    var chkOuterRound, editOuterRound;
    var outerCapPanel;
    var rbCapButt, rbCapRound, rbCapProject;
    var applyEdgeScaleEnabledState, applyOuterCapEnabledState, applyOuterRoundEnabledState;

    // タイトルエリア
    var titlePanel;
    var chkTitleEnable, editTitleSize, titleSizeGroup;
    var titlePositionGroup, rbTitleTop, rbTitleBottom, rbTitleLeft, rbTitleRight;
    var titleOptionGroup, chkTitleFill, chkTitleRule;
    var titleEdgeScaleRow, chkTitleEdgeScale, editTitleEdgeScale;
    var prevTitleHasSize = false;

    // フレーム
    var framePanel;
    var chkFrameEnable, editFrameWidth;
    var chkBleed;
    var chkFrameRound, editFrameRound;

    // 内側エリア
    var innerPanel;
    var editInnerOffsetTop, editInnerOffsetBottom, editInnerOffsetLeft, editInnerOffsetRight;
    var chkInnerOffsetLink;
    var isSyncingInnerOffsets = false;
    var applyInnerOffsetLinkState;
    var editColumnCount, editRowCount;
    var editColumnGutter, editRowGutter;
    var chkInnerCellFill, chkInnerDivider;
    var innerLineTypePanel;
    var rbInnerLineSolid, rbInnerLineDash, rbInnerLineDotted;

    // 下段
    var chkPreview;

    // ［塗り］を手動で切り替えたら、ガター変更による自動ONを行わない
    var isInnerFillManuallySet = false;

    // 分割可能になった瞬間だけ［分割線］を自動ONにするためのフラグ
    var wasGridSplittable = false;

    // =========================================
    // UI構築：3行3列の方向入力 / A 3x3 directional input grid
    // =========================================
    // ScriptUI にグリッドがないため、幅を固定したセルを3つ並べた行を3段重ねて桁を揃える。

    /**
     * 方向入力グリッドの1行を追加します。
     *
     * @param {Panel|Group} parent - 追加先のパネルまたはグループ。
     * @returns {Group} 追加した行グループ。
     */
    function addFieldGridRow(parent) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["center", "top"];
        row.alignChildren = ["center", "center"];
        row.spacing = 0;
        return row;
    }

    /**
     * 方向入力グリッドの1セルを追加します。桁を揃えるため幅を固定します。
     *
     * @param {Group} row - 追加先の行グループ。
     * @returns {Group} 追加したセルグループ。
     */
    function addFieldGridCell(row) {
        var cell = row.add("group");
        cell.orientation = "row";
        cell.alignChildren = ["center", "center"];
        cell.spacing = 4;
        try {
            cell.preferredSize.width = FIELD_CELL_WIDTH;
            cell.minimumSize.width = FIELD_CELL_WIDTH;
            cell.maximumSize.width = FIELD_CELL_WIDTH;
        } catch (e) { }
        return cell;
    }

    /**
     * セルにラベルと数値入力欄を配置します。入力欄は上下キーでの増減に対応します。
     *
     * @param {Group} cell - 追加先のセルグループ。
     * @param {string} labelText - ラベル文字列。
     * @param {string} defaultText - 入力欄の初期値。
     * @returns {EditText} 追加した入力欄。
     */
    function addLabeledField(cell, labelText, defaultText) {
        cell.add("statictext", undefined, withColon(labelText));
        var field = cell.add("edittext", undefined, defaultText);
        field.characters = 4;
        changeValueByArrowKey(field, false);
        return field;
    }

    // =========================================
    // UI構築：マージン / Build the margin panel
    // =========================================
    /**
     * マージンパネル（上下左右の入力と連動チェック）を組み立てます。
     * 長方形選択で開始した場合は非表示にします。
     *
     * @param {Group} parent - 追加先の列グループ。
     * @returns {void}
     */
    function buildMarginUI(parent) {
        // 単位はパネル名に持たせ、各セルはラベル＋入力欄だけにして桁を揃える
        marginPanel = parent.add("panel", undefined,
            getLabel(LABELS.panel.margin) + " (" + getCurrentRulerUnitLabel() + ")");
        setupPanel(marginPanel, 6);

        // 3行3列：中央列に［上］［連動］［下］、左右列に［左］［右］
        //   [    ] [ 上 ] [    ]
        //   [ 左 ] [連動] [ 右 ]
        //   [    ] [ 下 ] [    ]
        var topRow = addFieldGridRow(marginPanel);
        addFieldGridCell(topRow);
        var topCell = addFieldGridCell(topRow);
        addFieldGridCell(topRow);

        var middleRow = addFieldGridRow(marginPanel);
        var leftCell = addFieldGridCell(middleRow);
        var linkCell = addFieldGridCell(middleRow);
        var rightCell = addFieldGridCell(middleRow);

        var bottomRow = addFieldGridRow(marginPanel);
        addFieldGridCell(bottomRow);
        var bottomCell = addFieldGridCell(bottomRow);
        addFieldGridCell(bottomRow);

        editMarginTop = addLabeledField(topCell, getLabel(LABELS.field.offsetTop), "15");
        editMarginLeft = addLabeledField(leftCell, getLabel(LABELS.field.offsetLeft), "15");
        editMarginRight = addLabeledField(rightCell, getLabel(LABELS.field.offsetRight), "15");
        editMarginBottom = addLabeledField(bottomCell, getLabel(LABELS.field.offsetBottom), "15");

        chkMarginLink = linkCell.add("checkbox", undefined, getLabel(LABELS.checkbox.link));
        chkMarginLink.value = true;
        setTooltip(chkMarginLink, LABELS.tooltip.link);

        // 連動ONのとき、下／左／右は上の値に追従してディム表示
        /**
         * 連動チェックの状態に合わせて、下／左／右の入力欄を上の値に追従させます。
         *
         * @returns {void}
         */
        applyMarginLinkState = function () {
            var isLinked = !!chkMarginLink.value;
            try { bottomCell.enabled = !isLinked; } catch (e) { }
            try { leftCell.enabled = !isLinked; } catch (e) { }
            try { rightCell.enabled = !isLinked; } catch (e) { }

            if (isLinked) {
                var value = editMarginTop.text;
                isSyncingMargins = true;
                try { editMarginBottom.text = value; } catch (e) { }
                try { editMarginLeft.text = value; } catch (e) { }
                try { editMarginRight.text = value; } catch (e) { }
                isSyncingMargins = false;
            }
        };

        // 長方形選択で開始した場合は非表示にし、レイアウト上のスペースも潰す
        marginPanel.enabled = isArtboardBase;
        try {
            marginPanel.visible = isArtboardBase;
            marginPanel.minimumSize.height = 0;
            marginPanel.maximumSize.height = isArtboardBase ? 10000 : 0;
        } catch (e) { }

        applyMarginLinkState();
    }

    // =========================================
    // UI構築：外側エリア / Build the outer area panel
    // =========================================
    /**
     * 外側エリアパネル（外枠の保持・辺の伸縮・角丸・線端）を組み立てます。
     *
     * @param {Group} parent - 追加先の列グループ。
     * @returns {void}
     */
    function buildOuterUI(parent) {
        outerPanel = parent.add("panel", undefined, getLabel(LABELS.panel.outerArea));
        setupPanel(outerPanel, 10);

        chkKeepOuterFrame = outerPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.keepOuterFrame));
        chkKeepOuterFrame.value = true;
        setTooltip(chkKeepOuterFrame, LABELS.tooltip.keepOuterFrame);

        // 辺の伸縮
        var edgeScaleRow = outerPanel.add("group");
        setupRow(edgeScaleRow);

        chkOuterEdgeScale = edgeScaleRow.add("checkbox", undefined, withColon(getLabel(LABELS.checkbox.edgeScale)));
        chkOuterEdgeScale.value = true;
        setTooltip(chkOuterEdgeScale, LABELS.tooltip.outerEdgeScale);

        outerEdgeScaleGroup = edgeScaleRow.add("group");
        setupRow(outerEdgeScaleGroup);
        editOuterEdgeScale = outerEdgeScaleGroup.add("edittext", undefined, "-5");
        editOuterEdgeScale.characters = 4;
        outerEdgeScaleGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
        editOuterEdgeScale.active = true;
        changeValueByArrowKey(editOuterEdgeScale, true);

        /**
         * 外枠の保持状態に合わせて、辺の伸縮の入力可否を切り替えます。
         *
         * @returns {void}
         */
        applyEdgeScaleEnabledState = function () {
            try {
                var isKeepingOuter = !!chkKeepOuterFrame.value;
                chkOuterEdgeScale.enabled = isKeepingOuter;
                outerEdgeScaleGroup.enabled = (isKeepingOuter && chkOuterEdgeScale.value);
            } catch (e) { }
        };
        applyEdgeScaleEnabledState();

        // 角丸
        var roundRow = outerPanel.add("group");
        setupRow(roundRow);

        chkOuterRound = roundRow.add("checkbox", undefined, withColon(getLabel(LABELS.checkbox.roundCorner)));
        chkOuterRound.value = false;
        setTooltip(chkOuterRound, LABELS.tooltip.outerRound);

        editOuterRound = roundRow.add("edittext", undefined, "0");
        editOuterRound.characters = 4;
        changeValueByArrowKey(editOuterRound, false);

        roundRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        /**
         * 角丸チェックの状態に合わせて、角丸の入力可否と値を整えます。
         *
         * @returns {void}
         */
        applyOuterRoundEnabledState = function () {
            try {
                // タイトル角丸の参照値としても使うため、辺の伸縮ONでも入力自体は許可する
                chkOuterRound.enabled = true;
                editOuterRound.enabled = !!chkOuterRound.value;
                if (!chkOuterRound.value) editOuterRound.text = "0";
            } catch (e) { }
        };

        /**
         * 角丸をONにしたとき、初期半径を補い、併用できない辺の伸縮をOFFにします。
         *
         * @returns {void}
         */
        chkOuterRound.onClick = function () {
            try {
                if (chkOuterRound.value) {
                    var radius = parseFloat(editOuterRound.text);
                    if (isNaN(radius) || radius === 0) editOuterRound.text = "2";

                    // 角丸は辺の伸縮OFFのときだけ効くため、辺の伸縮を0にしてOFFにする
                    editOuterEdgeScale.text = "0";
                    chkOuterEdgeScale.value = false;
                }
            } catch (e) { }
            applyEdgeScaleEnabledState();
            applyOuterCapEnabledState();
            applyOuterRoundEnabledState();
            requestPreviewUpdate();
        };

        /**
         * 角丸の値が変わったらプレビューを更新します。
         *
         * @returns {void}
         */
        editOuterRound.onChanging = function () {
            requestPreviewUpdate();
        };

        // 線端
        outerCapPanel = outerPanel.add("panel", undefined, getLabel(LABELS.panel.lineCap));
        setupPanel(outerCapPanel);
        outerCapPanel.orientation = "row";
        outerCapPanel.alignChildren = ["left", "center"];

        rbCapButt = outerCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capButt));
        rbCapRound = outerCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capRound));
        rbCapProject = outerCapPanel.add("radiobutton", undefined, getLabel(LABELS.radio.capProject));
        setTooltip(outerCapPanel, LABELS.tooltip.lineCap);

        // 初期値：選択オブジェクトの線端を優先
        /**
         * 線端のラジオボタンの初期選択を、選択オブジェクトの線端から決めます。
         *
         * @returns {void}
         */
        (function initCapSelection() {
            var currentCap = null;
            try {
                if (baseRectItems.length > 0 && baseRectItems[0].stroked) currentCap = baseRectItems[0].strokeCap;
            } catch (e) { }

            if (currentCap === StrokeCap.ROUNDENDCAP) rbCapRound.value = true;
            else if (currentCap === StrokeCap.PROJECTINGENDCAP) rbCapProject.value = true;
            else rbCapButt.value = true;
        })();

        /**
         * 外枠を残し、かつ辺の伸縮が有効なときだけ線端パネルを操作可能にします。
         *
         * @returns {void}
         */
        applyOuterCapEnabledState = function () {
            try {
                outerCapPanel.enabled = (!!chkKeepOuterFrame.value
                    && !!chkOuterEdgeScale.value
                    && getOuterEdgeScaleValue() !== 0);
            } catch (e) { }
        };
        applyOuterCapEnabledState();
    }

    // =========================================
    // UI構築：タイトルエリア / Build the title area panel
    // =========================================
    /**
     * タイトルエリアパネル（有効・サイズ・位置・塗り・線・辺の伸縮）を組み立てます。
     *
     * @param {Group} parent - 追加先の列グループ。
     * @returns {void}
     */
    function buildTitleUI(parent) {
        titlePanel = parent.add("panel", undefined, getLabel(LABELS.panel.titleArea));
        setupPanel(titlePanel, 10);

        // 有効 ＋ 幅／高さ
        var enableRow = titlePanel.add("group");
        setupRow(enableRow);

        chkTitleEnable = enableRow.add("checkbox", undefined, "");
        chkTitleEnable.value = false;
        setTooltip(chkTitleEnable, LABELS.tooltip.titleEnable);

        titleSizeGroup = enableRow.add("group");
        setupRow(titleSizeGroup);
        titleSizeGroup.add("statictext", undefined, withColon(getLabel(LABELS.field.titleSize)));
        editTitleSize = titleSizeGroup.add("edittext", undefined, "0");
        editTitleSize.characters = 4;
        setTooltip(editTitleSize, LABELS.tooltip.titleSize);
        titleSizeGroup.add("statictext", undefined, getCurrentRulerUnitLabel());
        changeValueByArrowKey(editTitleSize, false);

        // 位置
        titlePositionGroup = titlePanel.add("group");
        setupRow(titlePositionGroup);
        rbTitleTop = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.posTop));
        rbTitleBottom = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.posBottom));
        rbTitleLeft = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.posLeft));
        rbTitleRight = titlePositionGroup.add("radiobutton", undefined, getLabel(LABELS.radio.posRight));
        rbTitleTop.value = true;

        // 塗り／線
        titleOptionGroup = titlePanel.add("group");
        setupRow(titleOptionGroup);
        chkTitleFill = titleOptionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.fill));
        chkTitleFill.value = false;
        setTooltip(chkTitleFill, LABELS.tooltip.titleFill);
        chkTitleRule = titleOptionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.rule));
        chkTitleRule.value = true;
        setTooltip(chkTitleRule, LABELS.tooltip.titleRule);

        // 辺の伸縮（タイトル帯の線の長さに反映。＋で伸ばす／−で短くする）
        titleEdgeScaleRow = titlePanel.add("group");
        setupRow(titleEdgeScaleRow);
        chkTitleEdgeScale = titleEdgeScaleRow.add("checkbox", undefined, withColon(getLabel(LABELS.checkbox.edgeScale)));
        chkTitleEdgeScale.value = false;
        setTooltip(chkTitleEdgeScale, LABELS.tooltip.titleEdgeScale);
        editTitleEdgeScale = titleEdgeScaleRow.add("edittext", undefined, "0");
        editTitleEdgeScale.characters = 4;
        changeValueByArrowKey(editTitleEdgeScale, true);
        titleEdgeScaleRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        prevTitleHasSize = hasTitleSize();

        applyTitleEdgeScaleEnabledState();
        applyTitleAreaEnabledState();
    }

    /**
     * タイトルエリアのサイズが正の値として入力されているかを判定します。
     *
     * @returns {boolean} サイズが0より大きければ true。
     */
    function hasTitleSize() {
        var size = parseFloat(editTitleSize.text);
        return (!isNaN(size) && size > 0);
    }

    /**
     * タイトル境界線の辺の伸縮について、入力欄の有効／無効を切り替えます。
     *
     * @returns {void}
     */
    function applyTitleEdgeScaleEnabledState() {
        try {
            editTitleEdgeScale.enabled = !!chkTitleEdgeScale.value;
            if (!chkTitleEdgeScale.value) editTitleEdgeScale.text = "0";
        } catch (e) { }
    }

    /**
     * タイトルエリアの有効状態とサイズに応じて、関連コントロールの有効／無効と値を整えます。
     *
     * @returns {void}
     */
    function applyTitleAreaEnabledState() {
        var isAreaEnabled = !!chkTitleEnable.value;

        // 無効ならサイズ0扱い（＝タイトル生成なし）
        if (!isAreaEnabled) {
            if (editTitleSize.text !== "0") editTitleSize.text = "0";
            chkTitleFill.value = false;
            chkTitleRule.value = false;
        }

        titleSizeGroup.enabled = isAreaEnabled;
        titleOptionGroup.enabled = isAreaEnabled;

        var isSizedAndEnabled = (isAreaEnabled && hasTitleSize());

        titlePositionGroup.enabled = isSizedAndEnabled;

        chkTitleFill.enabled = isSizedAndEnabled;
        if (!isSizedAndEnabled) chkTitleFill.value = false;

        // 線：サイズが 0 → >0 になった瞬間だけ自動ON（以後はユーザー操作を尊重）
        if (!prevTitleHasSize && isSizedAndEnabled) chkTitleRule.value = true;
        chkTitleRule.enabled = isSizedAndEnabled;
        if (!isSizedAndEnabled) chkTitleRule.value = false;
        prevTitleHasSize = isSizedAndEnabled;

        // ［辺の伸縮］は「タイトル有効」かつ「線ON」のときのみ。
        // ※上で chkTitleRule を確定させたあとに判定しないと、1操作ぶん状態が遅れる
        var isEdgeRowEnabled = (isAreaEnabled && !!chkTitleRule.value);
        try { titleEdgeScaleRow.enabled = isEdgeRowEnabled; } catch (e) { }
        if (!isEdgeRowEnabled) {
            try { chkTitleEdgeScale.value = false; } catch (e) { }
            try { editTitleEdgeScale.text = "0"; } catch (e) { }
            applyTitleEdgeScaleEnabledState();
        }
    }

    /**
     * 選択中のタイトル配置をキー文字列で返します。
     *
     * @returns {string} "top" / "bottom" / "left" / "right" のいずれか。
     */
    function getTitlePositionKey() {
        try {
            if (rbTitleRight && rbTitleRight.value) return "right";
            if (rbTitleBottom && rbTitleBottom.value) return "bottom";
            if (rbTitleLeft && rbTitleLeft.value) return "left";
        } catch (e) { }
        return "top";
    }

    /**
     * 外側エリアの角丸値をタイトル帯にも流用して角丸を適用します。
     *
     * @param {PathItem} titleRect - タイトル帯の矩形。
     * @returns {boolean} 角丸を適用できたら true。
     */
    function maybeApplyTitleAreaRound(titleRect) {
        try {
            if (!titleRect) return false;
            if (!chkOuterRound || !chkOuterRound.value) return false;
            var radiusValue = parseFloat(editOuterRound.text);
            if (isNaN(radiusValue) || radiusValue <= 0) return false;
            return !!applyTitleAreaCornerRounding(titleRect, getTitlePositionKey(), radiusValue);
        } catch (e) { }
        return false;
    }

    // =========================================
    // UI構築：フレーム / Build the frame panel
    // =========================================
    /**
     * フレームパネル（有効・幅・裁ち落とし・角丸）を組み立てます。
     *
     * @param {Group} parent - 追加先の列グループ。
     * @returns {void}
     */
    function buildFrameUI(parent) {
        framePanel = parent.add("panel", undefined, getLabel(LABELS.panel.frame));
        setupPanel(framePanel, 10);

        var widthRow = framePanel.add("group");
        setupRow(widthRow);

        chkFrameEnable = widthRow.add("checkbox", undefined, "");
        chkFrameEnable.value = false;
        setTooltip(chkFrameEnable, LABELS.tooltip.frameEnable);

        widthRow.add("statictext", undefined, withColon(getLabel(LABELS.field.width)));
        editFrameWidth = widthRow.add("edittext", undefined, "0");
        editFrameWidth.characters = 4;
        setTooltip(editFrameWidth, LABELS.tooltip.frameWidth);
        widthRow.add("statictext", undefined, getCurrentRulerUnitLabel());
        changeValueByArrowKey(editFrameWidth, false);

        // 裁ち落とし
        var bleedRow = framePanel.add("group");
        setupRow(bleedRow);
        chkBleed = bleedRow.add("checkbox", undefined, getLabel(LABELS.checkbox.bleed));
        chkBleed.value = false;
        setTooltip(chkBleed, LABELS.tooltip.bleed);

        // 角丸
        var roundRow = framePanel.add("group");
        setupRow(roundRow);
        chkFrameRound = roundRow.add("checkbox", undefined, withColon(getLabel(LABELS.checkbox.roundCorner)));
        chkFrameRound.value = false;
        setTooltip(chkFrameRound, LABELS.tooltip.frameRound);
        editFrameRound = roundRow.add("edittext", undefined, "0");
        editFrameRound.characters = 4;
        changeValueByArrowKey(editFrameRound, false);
        roundRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        applyFrameEnabledState();
    }

    /**
     * フレーム関連コントロールの有効／無効を整えます。
     * アートボード基準でないときはパネルごと隠します。
     *
     * @returns {void}
     */
    function applyFrameEnabledState() {
        try {
            // フレームはアートボード基準のみ有効。長方形選択時はパネルごと隠す
            if (!isArtboardBase) {
                try {
                    framePanel.visible = false;
                    framePanel.minimumSize.height = 0;
                    framePanel.maximumSize.height = 0;
                } catch (e) { }

                chkFrameEnable.value = false;
                editFrameWidth.text = "0";
                editFrameWidth.enabled = false;
                chkBleed.value = false;
                chkBleed.enabled = false;
                chkFrameRound.value = false;
                chkFrameRound.enabled = false;
                editFrameRound.text = "0";
                editFrameRound.enabled = false;
                isBleedEnabled = false;
                return;
            }

            try {
                framePanel.visible = true;
                framePanel.minimumSize.height = 0;
                framePanel.maximumSize.height = 10000;
            } catch (e) { }

            var isEnabled = !!chkFrameEnable.value;

            editFrameWidth.enabled = isEnabled;
            if (!isEnabled) editFrameWidth.text = "0";

            var frameWidth = parseFloat(editFrameWidth.text);
            var hasWidth = (!isNaN(frameWidth) && frameWidth > 0);

            chkBleed.enabled = (isEnabled && hasWidth);
            if (!chkBleed.enabled) chkBleed.value = false;

            chkFrameRound.enabled = (isEnabled && hasWidth);
            if (!chkFrameRound.enabled) chkFrameRound.value = false;
            editFrameRound.enabled = (isEnabled && hasWidth && chkFrameRound.value);

            // チェックの取り消しに追従させ、フレーム無効時に裁ち落としだけ残らないようにする
            isBleedEnabled = !!chkBleed.value;
        } catch (e) { }
    }

    // =========================================
    // UI構築：内側エリア / Build the inner area panel
    // =========================================
    /**
     * 内側エリアパネル（オフセット・列／行・線の種類）を組み立てます。
     *
     * @param {Group} parent - 追加先の列グループ。
     * @returns {void}
     */
    function buildInnerUI(parent) {
        innerPanel = parent.add("panel", undefined, getLabel(LABELS.panel.innerArea));
        setupPanel(innerPanel);

        buildInnerOffsetSection(innerPanel);
        buildInnerGridSection(innerPanel);
        buildInnerLineTypeSection(innerPanel);
    }

    /**
     * 内側エリアのオフセット入力（上下左右と連動チェック）を組み立てます。
     *
     * @param {Panel} parent - 追加先のパネル。
     * @returns {void}
     */
    function buildInnerOffsetSection(parent) {
        // 単位はパネル名に持たせ、各セルはラベル＋入力欄だけにして桁を揃える
        var offsetPanel = parent.add("panel", undefined,
            getLabel(LABELS.panel.offset) + " (" + getCurrentRulerUnitLabel() + ")");
        setupPanel(offsetPanel, 6);

        var defaultOffset = String(calcDefaultInnerOffset());

        // 3行3列：中央列に［上］［連動］［下］、左右列に［左］［右］
        //   [    ] [ 上 ] [    ]
        //   [ 左 ] [連動] [ 右 ]
        //   [    ] [ 下 ] [    ]
        var topRow = addFieldGridRow(offsetPanel);
        addFieldGridCell(topRow);
        var topCell = addFieldGridCell(topRow);
        addFieldGridCell(topRow);

        var middleRow = addFieldGridRow(offsetPanel);
        var leftCell = addFieldGridCell(middleRow);
        var linkCell = addFieldGridCell(middleRow);
        var rightCell = addFieldGridCell(middleRow);

        var bottomRow = addFieldGridRow(offsetPanel);
        addFieldGridCell(bottomRow);
        var bottomCell = addFieldGridCell(bottomRow);
        addFieldGridCell(bottomRow);

        editInnerOffsetTop = addLabeledField(topCell, getLabel(LABELS.field.offsetTop), defaultOffset);
        editInnerOffsetLeft = addLabeledField(leftCell, getLabel(LABELS.field.offsetLeft), defaultOffset);
        editInnerOffsetRight = addLabeledField(rightCell, getLabel(LABELS.field.offsetRight), defaultOffset);
        editInnerOffsetBottom = addLabeledField(bottomCell, getLabel(LABELS.field.offsetBottom), defaultOffset);

        chkInnerOffsetLink = linkCell.add("checkbox", undefined, getLabel(LABELS.checkbox.link));
        chkInnerOffsetLink.value = true;
        setTooltip(chkInnerOffsetLink, LABELS.tooltip.link);

        // 連動ONのとき、下／左／右は上の値に追従してディム表示
        /**
         * 連動チェックの状態に合わせて、下／左／右のオフセットを上の値に追従させます。
         *
         * @returns {void}
         */
        applyInnerOffsetLinkState = function () {
            var isLinked = !!chkInnerOffsetLink.value;
            try { bottomCell.enabled = !isLinked; } catch (e) { }
            try { leftCell.enabled = !isLinked; } catch (e) { }
            try { rightCell.enabled = !isLinked; } catch (e) { }

            if (isLinked) {
                var value = editInnerOffsetTop.text;
                isSyncingInnerOffsets = true;
                try { editInnerOffsetBottom.text = value; } catch (e) { }
                try { editInnerOffsetLeft.text = value; } catch (e) { }
                try { editInnerOffsetRight.text = value; } catch (e) { }
                isSyncingInnerOffsets = false;
            }
        };
        applyInnerOffsetLinkState();
    }

    /**
     * 内側エリアの列・行の設定（分割数・間隔・塗り・分割線）を組み立てます。
     *
     * @param {Panel} parent - 追加先のパネル。
     * @returns {void}
     */
    function buildInnerGridSection(parent) {
        var gridWrap = parent.add("group");
        setupRow(gridWrap, ["fill", "top"]);
        gridWrap.alignChildren = ["left", "top"];

        var gridColumn = gridWrap.add("group");
        gridColumn.orientation = "column";
        gridColumn.alignChildren = ["left", "top"];
        gridColumn.alignment = ["left", "top"];
        gridColumn.spacing = COLUMN_SPACING;

        // 列
        var columnsPanel = gridColumn.add("panel", undefined, getLabel(LABELS.panel.columns));
        setupPanel(columnsPanel);

        var columnsRow = columnsPanel.add("group");
        setupRow(columnsRow);
        columnsRow.add("statictext", undefined, withColon(getLabel(LABELS.field.columnCount)));
        editColumnCount = columnsRow.add("edittext", undefined, "1");
        editColumnCount.characters = 3;
        changeValueByArrowKey(editColumnCount, false);

        columnsRow.add("statictext", undefined, withColon(getLabel(LABELS.field.spacing)));
        editColumnGutter = columnsRow.add("edittext", undefined, "0");
        editColumnGutter.characters = 4;
        setTooltip(editColumnGutter, LABELS.tooltip.gutter);
        changeValueByArrowKey(editColumnGutter, false);
        columnsRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        // 行
        var rowsPanel = gridColumn.add("panel", undefined, getLabel(LABELS.panel.rows));
        setupPanel(rowsPanel);

        var rowsRow = rowsPanel.add("group");
        setupRow(rowsRow);
        rowsRow.add("statictext", undefined, withColon(getLabel(LABELS.field.rowCount)));
        editRowCount = rowsRow.add("edittext", undefined, "1");
        editRowCount.characters = 3;
        changeValueByArrowKey(editRowCount, false);

        rowsRow.add("statictext", undefined, withColon(getLabel(LABELS.field.spacing)));
        editRowGutter = rowsRow.add("edittext", undefined, "0");
        editRowGutter.characters = 4;
        setTooltip(editRowGutter, LABELS.tooltip.gutter);
        changeValueByArrowKey(editRowGutter, false);
        rowsRow.add("statictext", undefined, getCurrentRulerUnitLabel());

        // 塗り／分割線
        var optionWrap = gridColumn.add("group");
        setupRow(optionWrap, ["fill", "top"]);
        optionWrap.alignChildren = ["center", "center"];

        var optionGroup = optionWrap.add("group");
        setupRow(optionGroup, ["center", "center"]);

        chkInnerCellFill = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.fill));
        chkInnerCellFill.value = false;
        setTooltip(chkInnerCellFill, LABELS.tooltip.innerFill);

        chkInnerDivider = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.divider));
        chkInnerDivider.value = false;
        setTooltip(chkInnerDivider, LABELS.tooltip.innerDivider);

        try { editColumnGutter.enabled = (parseInt(editColumnCount.text, 10) > 1); } catch (e) { }
        try { editRowGutter.enabled = (parseInt(editRowCount.text, 10) > 1); } catch (e) { }
    }

    /**
     * 分割線の線種を選ぶラジオボタン群を組み立てます。
     *
     * @param {Panel} parent - 追加先のパネル。
     * @returns {void}
     */
    function buildInnerLineTypeSection(parent) {
        innerLineTypePanel = parent.add("panel", undefined, getLabel(LABELS.panel.lineType));
        setupPanel(innerLineTypePanel);
        innerLineTypePanel.orientation = "row";
        innerLineTypePanel.alignChildren = ["left", "center"];

        rbInnerLineSolid = innerLineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineSolid));
        rbInnerLineDash = innerLineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineDash));
        rbInnerLineDotted = innerLineTypePanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineDotted));
        rbInnerLineSolid.value = true;
    }

    // =========================================
    // UI構築：画面表示 / Build the display panel
    // =========================================
    /**
     * 画面表示タブ（ズームとパンのスライダー、表示コマンドのボタン）を組み立てます。
     *
     * @param {Group} parent - 追加先の列グループ。
     * @returns {void}
     */
    function buildDisplayUI(parent) {
        // ズームとパン
        var zoomPanPanel = parent.add("panel", undefined, getLabel(LABELS.panel.zoomPan));
        setupPanel(zoomPanPanel, 10);

        var sliderGroup = zoomPanPanel.add("group");
        sliderGroup.orientation = "column";
        sliderGroup.alignChildren = "left";
        sliderGroup.spacing = PANEL_SPACING;

        if (viewCtl && typeof viewCtl.buildUI === "function") {
            viewCtl.buildUI(sliderGroup, { labelWidth: 58, sliderWidth: 200 });
        } else {
            sliderGroup.add("statictext", undefined, getLabel(LABELS.message.viewControlUnavailable));
        }

        // 表示コマンド
        var screenPanel = parent.add("panel", undefined, getLabel(LABELS.panel.screen));
        setupPanel(screenPanel, 10);

        // ボタンはパネル幅いっぱいに広げず左寄せにする
        var commandGroup = screenPanel.add("group");
        commandGroup.orientation = "column";
        commandGroup.alignChildren = ["left", "top"];
        commandGroup.alignment = ["left", "top"];
        commandGroup.spacing = 6;

        /**
         * 表示メニューコマンドを実行するボタンを追加します。
         *
         * @param {Object} labelEntry - ボタン名のラベル定義。
         * @param {string} menuCommand - 実行するメニューコマンド名。
         * @returns {Button} 追加したボタン。
         */
        function addViewCommandButton(labelEntry, menuCommand) {
            var button = commandGroup.add("button", undefined, getLabel(labelEntry));
            button.alignment = "left";
            try { button.preferredSize.width = VIEW_BUTTON_WIDTH; } catch (e) { }
            /**
             * ボタンに割り当てられた表示メニューコマンドを実行します。
             *
             * @returns {void}
             */
            button.onClick = function () {
                try { app.executeMenuCommand(menuCommand); } catch (e) { }
            };
            return button;
        }

        addViewCommandButton(LABELS.button.fitIn, 'fitin');
        addViewCommandButton(LABELS.button.actualSize, 'actualsize');
        addViewCommandButton(LABELS.button.fitAll, 'fitall');
    }

    // =========================================
    // UIの組み立て / Assemble the UI
    // =========================================
    buildMarginUI(artboardColumn);
    buildFrameUI(artboardColumn);
    buildOuterUI(outerColumn);
    buildTitleUI(outerColumn);
    buildInnerUI(innerColumn);
    buildDisplayUI(displayColumn);

    // 長方形選択で開始した場合はマージンタブごと隠す
    try {
        if (!isArtboardBase) {
            artboardTab.visible = false;
            artboardTab.enabled = false;
            tabPanel.selection = outerTab;
        }
    } catch (e) { }

    // 分割線の初期状態
    /**
     * 起動時の列数・行数から、分割線チェックの初期状態を決めます。
     *
     * @returns {void}
     */
    (function initGridSplittableState() {
        var columnCount = readPositiveInt(editColumnCount, 1);
        var rowCount = readPositiveInt(editRowCount, 1);
        wasGridSplittable = (columnCount > 1 || rowCount > 1);
        applyInnerDividerEnabledState(columnCount, rowCount, false);
    })();

    // =========================================
    // 下段（プレビュー／ボタン） / Bottom row
    // =========================================
    var footerRow = dialog.add("group");
    setupRow(footerRow, ["fill", "center"]);

    chkPreview = footerRow.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
    chkPreview.value = true;
    chkPreview.alignment = "left";
    setTooltip(chkPreview, LABELS.tooltip.preview);

    var footerSpacer = footerRow.add("group");
    footerSpacer.alignment = ["fill", "fill"];
    footerSpacer.minimumSize.width = 0;

    var buttonGroup = footerRow.add("group");
    setupRow(buttonGroup, ["right", "center"]);
    var cancelButton = buttonGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    /**
     * 可能であれば、ダイアログを開く前のズームと表示位置へ戻します。
     *
     * @returns {void}
     */
    function restoreViewIfPossible() {
        try {
            if (viewCtl && typeof viewCtl.restore === "function") viewCtl.restore();
        } catch (e) { }
    }

    /**
     * キャンセル時に、表示を元に戻しプレビュー生成物を破棄してダイアログを閉じます。
     *
     * @returns {void}
     */
    cancelButton.onClick = function () {
        restoreViewIfPossible();
        try { clearPreview(); } catch (e) { }
        try { dialog.close(0); } catch (e) { }
    };

    /**
     * ダイアログが閉じられるときに表示を元へ戻します。
     *
     * @returns {boolean} 常に true（閉じる操作を許可）。
     */
    dialog.onClose = function () {
        restoreViewIfPossible();
        return true;
    };

    // =========================================
    // セッション値の復元 / Restore the last session values
    // =========================================
    /**
     * セッションに保存された前回の入力値をダイアログへ復元します。
     *
     * @returns {void}
     */
    (function restoreSessionState() {
        var savedState = loadSessionState();
        if (!savedState) return;

        /**
         * 保存済み設定を階層キーでたどって値を取り出します。
         *
         * @param {Array<string>} pathParts - たどるキーの並び。
         * @returns {*} 見つかった値。存在しなければ undefined。
         */
        function readPath(pathParts) {
            try {
                var node = savedState;
                for (var i = 0; i < pathParts.length; i++) {
                    if (!node) return undefined;
                    node = node[pathParts[i]];
                }
                return node;
            } catch (e) { return undefined; }
        }

        /**
         * 保存済みの値を入力欄のテキストへ復元します。
         *
         * @param {EditText} field - 復元先の入力欄。
         * @param {Array<string>} pathParts - たどるキーの並び。
         * @returns {void}
         */
        function restoreText(field, pathParts) {
            try {
                var value = readPath(pathParts);
                if (typeof value === "undefined" || value === null) return;
                if (typeof value === "object") return;
                field.text = String(value);
            } catch (e) { }
        }

        /**
         * 保存済みの値をチェックボックスの状態へ復元します。
         *
         * @param {Checkbox} checkbox - 復元先のチェックボックス。
         * @param {Array<string>} pathParts - たどるキーの並び。
         * @returns {void}
         */
        function restoreCheck(checkbox, pathParts) {
            try {
                var value = readPath(pathParts);
                if (typeof value === "undefined" || value === null) return;
                checkbox.value = !!value;
            } catch (e) { }
        }

        /**
         * 保存済みのキーに対応するラジオボタンを選択状態にします。
         *
         * @param {string} savedKey - 保存されていた選択キー。
         * @param {Object} controlsByKey - キーとラジオボタンの対応表。
         * @returns {void}
         */
        function restoreChoice(savedKey, controlsByKey) {
            try {
                if (typeof savedKey === "undefined" || savedKey === null) return;
                if (controlsByKey[savedKey]) controlsByKey[savedKey].value = true;
            } catch (e) { }
        }

        restoreCheck(chkPreview, ["preview"]);

        // マージン
        restoreText(editMarginTop, ["margin", "top"]);
        restoreText(editMarginBottom, ["margin", "bottom"]);
        restoreText(editMarginLeft, ["margin", "left"]);
        restoreText(editMarginRight, ["margin", "right"]);
        restoreCheck(chkMarginLink, ["margin", "link"]);

        // 外側エリア
        restoreCheck(chkKeepOuterFrame, ["outer", "keepFrame"]);
        restoreCheck(chkOuterEdgeScale, ["outer", "edgeScaleEnabled"]);
        restoreText(editOuterEdgeScale, ["outer", "edgeScaleValue"]);
        restoreChoice(readPath(["outer", "cap"]),
            { round: rbCapRound, project: rbCapProject, butt: rbCapButt });
        restoreCheck(chkOuterRound, ["outer", "roundEnabled"]);
        restoreText(editOuterRound, ["outer", "roundValue"]);

        // タイトルエリア
        restoreCheck(chkTitleEnable, ["title", "enabled"]);
        restoreText(editTitleSize, ["title", "size"]);
        restoreChoice(readPath(["title", "position"]),
            { top: rbTitleTop, bottom: rbTitleBottom, left: rbTitleLeft, right: rbTitleRight });
        restoreCheck(chkTitleFill, ["title", "fill"]);
        restoreCheck(chkTitleRule, ["title", "rule"]);
        restoreCheck(chkTitleEdgeScale, ["title", "edgeScaleEnabled"]);
        restoreText(editTitleEdgeScale, ["title", "edgeScaleValue"]);

        // 復元したサイズを基準にしないと、［線］OFF が自動ONで上書きされてしまう
        prevTitleHasSize = (hasTitleSize() && !!chkTitleEnable.value);

        // フレーム
        restoreCheck(chkFrameEnable, ["frame", "enabled"]);
        restoreText(editFrameWidth, ["frame", "width"]);
        restoreCheck(chkBleed, ["frame", "bleed"]);
        restoreCheck(chkFrameRound, ["frame", "roundEnabled"]);
        restoreText(editFrameRound, ["frame", "roundValue"]);

        // 内側エリア
        restoreCheck(chkInnerOffsetLink, ["inner", "link"]);
        restoreText(editInnerOffsetTop, ["inner", "offset", "top"]);
        restoreText(editInnerOffsetBottom, ["inner", "offset", "bottom"]);
        restoreText(editInnerOffsetLeft, ["inner", "offset", "left"]);
        restoreText(editInnerOffsetRight, ["inner", "offset", "right"]);
        restoreText(editColumnCount, ["inner", "grid", "columns"]);
        restoreText(editRowCount, ["inner", "grid", "rows"]);
        restoreText(editColumnGutter, ["inner", "grid", "columnGutter"]);
        restoreText(editRowGutter, ["inner", "grid", "rowGutter"]);
        restoreCheck(chkInnerCellFill, ["inner", "grid", "fill"]);
        restoreCheck(chkInnerDivider, ["inner", "grid", "divider"]);
        restoreChoice(readPath(["inner", "grid", "lineType"]),
            { solid: rbInnerLineSolid, dash: rbInnerLineDash, dotted: rbInnerLineDotted });

        // 手動で切り替えた［塗り］は、復元後も自動ONで上書きしない
        var manualFill = readPath(["inner", "grid", "fillManuallySet"]);
        if (typeof manualFill !== "undefined" && manualFill !== null) {
            isInnerFillManuallySet = !!manualFill;
        }

        isBleedEnabled = !!chkBleed.value;

        // 他コントロールに依存する有効／無効状態を再適用
        applyMarginLinkState();
        applyEdgeScaleEnabledState();
        applyOuterCapEnabledState();
        applyOuterRoundEnabledState();
        applyTitleEdgeScaleEnabledState();
        applyTitleAreaEnabledState();
        applyFrameEnabledState();
        applyInnerOffsetLinkState();

        var restoredColumnCount = readPositiveInt(editColumnCount, 1);
        var restoredRowCount = readPositiveInt(editRowCount, 1);
        try { editColumnGutter.enabled = (restoredColumnCount > 1); } catch (e) { }
        try { editRowGutter.enabled = (restoredRowCount > 1); } catch (e) { }
        applyInnerDividerEnabledState(restoredColumnCount, restoredRowCount, false);
    })();

    // =========================================
    // イベント / Event handlers
    // =========================================

    /**
     * プレビューが有効なときだけ、カンバス上の仮描画を更新します。
     *
     * @returns {void}
     */
    function requestPreviewUpdate() {
        if (chkPreview && chkPreview.value) updatePreview(false);
    }

    // --- 外側エリア ---
    /**
     * 辺の伸縮の値が変わったとき、線端の有効状態とプレビューを更新します。
     *
     * @returns {void}
     */
    editOuterEdgeScale.onChanging = function () {
        if (!chkOuterEdgeScale.value) return;
        applyOuterCapEnabledState();
        requestPreviewUpdate();
    };

    /**
     * 辺の伸縮チェックの切り替えに合わせて、関連コントロールとプレビューを更新します。
     *
     * @returns {void}
     */
    chkOuterEdgeScale.onClick = function () {
        applyEdgeScaleEnabledState();
        applyOuterCapEnabledState();
        applyOuterRoundEnabledState();
        requestPreviewUpdate();
    };

    /**
     * 外枠を残すチェックの切り替えに合わせて、関連コントロールとプレビューを更新します。
     *
     * @returns {void}
     */
    chkKeepOuterFrame.onClick = function () {
        applyEdgeScaleEnabledState();
        applyOuterCapEnabledState();
        applyOuterRoundEnabledState();
        requestPreviewUpdate();
    };

    /**
     * 線端の選択が変わったらプレビューを更新します。
     *
     * @returns {void}
     */
    function onOuterCapChanged() { requestPreviewUpdate(); }
    rbCapButt.onClick = onOuterCapChanged;
    rbCapRound.onClick = onOuterCapChanged;
    rbCapProject.onClick = onOuterCapChanged;

    // --- マージン ---
    /**
     * 上マージンの変更を、連動が有効なら他の3辺へ反映し、プレビューを更新します。
     *
     * @returns {void}
     */
    editMarginTop.onChanging = function () {
        if (!isArtboardBase || isSyncingMargins) return;
        if (chkMarginLink && chkMarginLink.value) {
            var value = editMarginTop.text;
            isSyncingMargins = true;
            try { editMarginBottom.text = value; } catch (e) { }
            try { editMarginLeft.text = value; } catch (e) { }
            try { editMarginRight.text = value; } catch (e) { }
            isSyncingMargins = false;
        }
        requestPreviewUpdate();
    };

    /**
     * 下・左・右マージンの変更に合わせてプレビューを更新します。
     *
     * @returns {void}
     */
    function onMarginFieldChanged() {
        if (!isArtboardBase || isSyncingMargins) return;
        requestPreviewUpdate();
    }
    editMarginBottom.onChanging = onMarginFieldChanged;
    editMarginLeft.onChanging = onMarginFieldChanged;
    editMarginRight.onChanging = onMarginFieldChanged;

    /**
     * マージンの連動チェックの切り替えを反映し、プレビューを更新します。
     *
     * @returns {void}
     */
    chkMarginLink.onClick = function () {
        applyMarginLinkState();
        requestPreviewUpdate();
    };

    // --- タイトルエリア ---
    /**
     * タイトルエリアを有効にしたとき、サイズが未入力なら初期値を補います。
     *
     * @returns {void}
     */
    chkTitleEnable.onClick = function () {
        try {
            if (chkTitleEnable.value) {
                var size = parseFloat(editTitleSize.text);
                if (isNaN(size) || size === 0) editTitleSize.text = String(calcDefaultTitleSize());
            }
        } catch (e) { }
        applyTitleAreaEnabledState();
        requestPreviewUpdate();
    };

    /**
     * タイトルサイズの変更に合わせて、関連コントロールとプレビューを更新します。
     *
     * @returns {void}
     */
    editTitleSize.onChanging = function () {
        applyTitleAreaEnabledState();
        requestPreviewUpdate();
    };

    /**
     * タイトルの配置が変わったらプレビューを更新します。
     *
     * @returns {void}
     */
    function onTitlePositionChanged() { requestPreviewUpdate(); }
    rbTitleTop.onClick = onTitlePositionChanged;
    rbTitleBottom.onClick = onTitlePositionChanged;
    rbTitleLeft.onClick = onTitlePositionChanged;
    rbTitleRight.onClick = onTitlePositionChanged;

    /**
     * タイトル帯の塗りの切り替えに合わせてプレビューを更新します。
     *
     * @returns {void}
     */
    chkTitleFill.onClick = function () { requestPreviewUpdate(); };

    /**
     * タイトル境界線の切り替えに合わせて、関連コントロールとプレビューを更新します。
     *
     * @returns {void}
     */
    chkTitleRule.onClick = function () {
        applyTitleAreaEnabledState();
        requestPreviewUpdate();
    };

    /**
     * タイトル境界線の辺の伸縮チェックの切り替えを反映し、プレビューを更新します。
     *
     * @returns {void}
     */
    chkTitleEdgeScale.onClick = function () {
        applyTitleEdgeScaleEnabledState();
        requestPreviewUpdate();
    };

    /**
     * タイトル境界線の伸縮値の変更に合わせてプレビューを更新します。
     *
     * @returns {void}
     */
    editTitleEdgeScale.onChanging = function () {
        if (!chkTitleEdgeScale.value) return;
        requestPreviewUpdate();
    };

    // --- フレーム ---
    /**
     * フレームを有効にしたとき、幅の初期値と裁ち落としを補い、プレビューを更新します。
     *
     * @returns {void}
     */
    chkFrameEnable.onClick = function () {
        try {
            if (chkFrameEnable.value) {
                var width = parseFloat(editFrameWidth.text);
                if (isNaN(width) || width === 0) editFrameWidth.text = "10";
                try { if (chkBleed) chkBleed.value = true; } catch (e) { }
            }
            applyFrameEnabledState();
            isBleedEnabled = !!chkBleed.value;
        } catch (e) { }
        requestPreviewUpdate();
    };

    /**
     * フレーム幅の変更に合わせて、裁ち落としと角丸の有効状態およびプレビューを更新します。
     *
     * @returns {void}
     */
    editFrameWidth.onChanging = function () {
        if (!chkFrameEnable.value) {
            editFrameWidth.text = "0";
            applyFrameEnabledState();
            isBleedEnabled = !!chkBleed.value;
            requestPreviewUpdate();
            return;
        }
        try {
            var width = parseFloat(editFrameWidth.text);
            var hasWidth = (!isNaN(width) && width > 0);

            chkBleed.enabled = hasWidth;
            // 幅が 0 → >0 になったら裁ち落としを自動ON
            if (hasWidth) {
                if (!chkBleed.value) chkBleed.value = true;
            } else {
                chkBleed.value = false;
            }

            chkFrameRound.enabled = hasWidth;
            if (!hasWidth) {
                chkFrameRound.value = false;
                editFrameRound.enabled = false;
            }
            isBleedEnabled = !!chkBleed.value;
        } catch (e) { }
        requestPreviewUpdate();
    };

    /**
     * 裁ち落としの切り替えを反映し、プレビューを更新します。
     *
     * @returns {void}
     */
    chkBleed.onClick = function () {
        isBleedEnabled = !!chkBleed.value;
        requestPreviewUpdate();
    };

    /**
     * フレームの角丸チェックの切り替えを反映し、初期半径を補ってプレビューを更新します。
     *
     * @returns {void}
     */
    chkFrameRound.onClick = function () {
        try {
            editFrameRound.enabled = chkFrameRound.value;
            if (chkFrameRound.value) {
                var radius = parseFloat(editFrameRound.text);
                if (isNaN(radius) || radius === 0) editFrameRound.text = "2";
            }
        } catch (e) { }
        requestPreviewUpdate();
    };

    /**
     * フレームの角丸の値が変わったらプレビューを更新します。
     *
     * @returns {void}
     */
    editFrameRound.onChanging = function () { requestPreviewUpdate(); };

    // --- 内側エリア ---
    /**
     * 上オフセットの変更を、連動が有効なら他の3辺へ反映し、プレビューを更新します。
     *
     * @returns {void}
     */
    editInnerOffsetTop.onChanging = function () {
        if (isSyncingInnerOffsets) return;
        if (chkInnerOffsetLink.value) {
            var value = editInnerOffsetTop.text;
            isSyncingInnerOffsets = true;
            try { editInnerOffsetBottom.text = value; } catch (e) { }
            try { editInnerOffsetLeft.text = value; } catch (e) { }
            try { editInnerOffsetRight.text = value; } catch (e) { }
            isSyncingInnerOffsets = false;
        }
        requestPreviewUpdate();
    };

    /**
     * 下・左・右オフセットの変更に合わせてプレビューを更新します。
     *
     * @returns {void}
     */
    function onInnerOffsetChanged() {
        if (isSyncingInnerOffsets) return;
        requestPreviewUpdate();
    }
    editInnerOffsetBottom.onChanging = onInnerOffsetChanged;
    editInnerOffsetLeft.onChanging = onInnerOffsetChanged;
    editInnerOffsetRight.onChanging = onInnerOffsetChanged;

    /**
     * 内側オフセットの連動チェックの切り替えを反映し、プレビューを更新します。
     *
     * @returns {void}
     */
    chkInnerOffsetLink.onClick = function () {
        applyInnerOffsetLinkState();
        requestPreviewUpdate();
    };

    /**
     * 列数・行数の変更を正規化し、間隔と分割線の有効状態およびプレビューを更新します。
     *
     * @returns {void}
     */
    function onGridCountChanged() {
        var columnCount = readPositiveInt(editColumnCount, 1);
        var rowCount = readPositiveInt(editRowCount, 1);

        if (String(columnCount) !== editColumnCount.text) editColumnCount.text = String(columnCount);
        if (String(rowCount) !== editRowCount.text) editRowCount.text = String(rowCount);

        try { editColumnGutter.enabled = (columnCount > 1); } catch (e) { }
        try { editRowGutter.enabled = (rowCount > 1); } catch (e) { }

        applyInnerDividerEnabledState(columnCount, rowCount, true);
        requestPreviewUpdate();
    }
    editColumnCount.onChanging = onGridCountChanged;
    editRowCount.onChanging = onGridCountChanged;

    /**
     * 間隔の変更に合わせて、必要ならセル塗りを自動でONにしプレビューを更新します。
     *
     * @param {EditText} gutterField - 変更された間隔の入力欄。
     * @returns {void}
     */
    function onGutterChanged(gutterField) {
        var gutter = parseFloat(gutterField.text);
        // ガターが設定されたら塗りを自動ON（ただし手動操作があれば尊重）
        if (!isInnerFillManuallySet && !isNaN(gutter) && gutter !== 0) {
            try { if (!chkInnerCellFill.value) chkInnerCellFill.value = true; } catch (e) { }
        }
        requestPreviewUpdate();
    }
    /**
     * 列の間隔の変更を処理します。
     *
     * @returns {void}
     */
    editColumnGutter.onChanging = function () { onGutterChanged(editColumnGutter); };
    /**
     * 行の間隔の変更を処理します。
     *
     * @returns {void}
     */
    editRowGutter.onChanging = function () { onGutterChanged(editRowGutter); };

    /**
     * セル塗りが手動で操作されたことを記録し、プレビューを更新します。
     *
     * @returns {void}
     */
    chkInnerCellFill.onClick = function () {
        // ユーザーが明示的に操作したら以後は自動ONしない
        isInnerFillManuallySet = true;
        requestPreviewUpdate();
    };

    /**
     * 分割線の切り替えに合わせて線種パネルの有効状態とプレビューを更新します。
     *
     * @returns {void}
     */
    chkInnerDivider.onClick = function () {
        try { innerLineTypePanel.enabled = (chkInnerDivider.enabled && chkInnerDivider.value); } catch (e) { }
        requestPreviewUpdate();
    };

    /**
     * 分割線の線種が変わったらプレビューを更新します。
     *
     * @returns {void}
     */
    function onInnerLineTypeChanged() { requestPreviewUpdate(); }
    rbInnerLineSolid.onClick = onInnerLineTypeChanged;
    rbInnerLineDash.onClick = onInnerLineTypeChanged;
    rbInnerLineDotted.onClick = onInnerLineTypeChanged;

    // --- プレビュー ---
    /**
     * プレビューチェックの切り替えに応じて、仮描画を行うか破棄します。
     *
     * @returns {void}
     */
    chkPreview.onClick = function () {
        if (chkPreview.value) updatePreview(false);
        else clearPreview();
    };

    // =========================================
    // ダイアログ表示 / Show the dialog
    // =========================================

    // レイアウト確定（tabbedpanel の内容が潰れるのを防ぐ）
    try { dialog.layout.layout(true); } catch (e) { }
    try { dialog.layout.resize(); } catch (e) { }

    if (chkPreview.value) updatePreview(false);

    var dialogResult = dialog.show();

    persistSessionState();

    if (dialogResult != 1) {
        // キャンセル：プレビュー生成物を削除して終了
        restoreViewIfPossible();
        try { clearPreview(); } catch (e) { }
        return;
    }

    finalizeResult();
    return;

    // =========================================
    // 最終生成 / Finalize on OK
    // =========================================
    /**
     * OK時の最終生成を行い、不要になった外枠の後始末と選択解除まで済ませます。
     *
     * @returns {void}
     */
    function finalizeResult() {
        // 何も指定されていない場合は生成せず、元の状態へ戻して終了
        if (!hasAnythingToGenerate()) {
            clearPreview();
            return;
        }

        // 最終生成（OK結果はヒストリーに残す）
        updatePreview(true);

        disposeOuterFrameAsRequested();
        detachFinalItemsFromPreviewList();

        // ダイアログ終了時は何も選択しない状態にする
        try { doc.selection = null; } catch (e) { }
    }

    /**
     * 現在の設定で生成すべきものがあるかを判定します。
     *
     * @returns {boolean} 生成対象があれば true。
     */
    function hasAnythingToGenerate() {
        // ［外枠を残す］OFF は「元の長方形を消す」という明示的な指示
        if (!chkKeepOuterFrame.value) return true;
        if (getOuterEdgeScaleValue() !== 0) return true;
        if (readNonNegativeFloat(editInnerOffsetTop) > 0) return true;
        if (readNonNegativeFloat(editInnerOffsetBottom) > 0) return true;
        if (readNonNegativeFloat(editInnerOffsetLeft) > 0) return true;
        if (readNonNegativeFloat(editInnerOffsetRight) > 0) return true;
        if (readPositiveInt(editColumnCount, 1) > 1) return true;
        if (readPositiveInt(editRowCount, 1) > 1) return true;
        if (chkTitleEnable.value && hasTitleSize()) return true;
        if (chkFrameEnable.value && readNonNegativeFloat(editFrameWidth) > 0) return true;
        if (chkOuterRound.value && readNonNegativeFloat(editOuterRound) > 0) return true;
        if (chkInnerCellFill.value) return true;
        return false;
    }

    /**
     * 設定に応じて、元の長方形や生成した4辺線を削除または表示に戻します。
     *
     * @returns {void}
     */
    function disposeOuterFrameAsRequested() {
        // 辺の伸縮が有効なときは外枠を4辺線で表すため、元の長方形は不要
        var shouldRemoveBaseRects = (getOuterEdgeScaleValue() !== 0) || !chkKeepOuterFrame.value;

        if (shouldRemoveBaseRects) {
            for (var i = 0; i < baseRectItems.length; i++) {
                try { baseRectItems[i].remove(); } catch (e) { }
            }
            baseRectItems = [];
            artboardBaseRect = null;
        } else {
            for (var j = 0; j < baseRectItems.length; j++) {
                try { baseRectItems[j].hidden = false; } catch (e) { }
            }
        }

        // ［外枠を残す］OFF のときは、生成した4辺線も削除
        if (!chkKeepOuterFrame.value) removeTaggedPreviewItems("__OuterEdge__", true);
    }

    /**
     * 最終生成物をプレビュー一覧から外します。オブジェクト自体は残します。
     *
     * @returns {void}
     */
    function detachFinalItemsFromPreviewList() {
        removeTaggedPreviewItems("__InnerBoxFill__", false);
        removeTaggedPreviewItems("__TitleFill__", false);
        removeTaggedPreviewItems("__FrameFill__", false);
    }

    /**
     * 指定タグの付いた一時アイテムをプレビュー一覧から外します。必要なら削除もします。
     *
     * @param {string} tag - 対象アイテムに付けたタグ文字列。
     * @param {boolean} shouldDelete - true ならアイテム自体も削除します。
     * @returns {void}
     */
    function removeTaggedPreviewItems(tag, shouldDelete) {
        for (var i = previewItems.length - 1; i >= 0; i--) {
            try {
                if (!hasTag(previewItems[i], tag)) continue;
                if (shouldDelete) {
                    try { previewItems[i].remove(); } catch (e) { }
                }
                previewItems.splice(i, 1);
            } catch (e) { }
        }
    }

    /**
     * アイテムの note または name に指定のタグが付いているかを判定します。
     *
     * @param {PathItem|GroupItem} item - 判定するアイテム。
     * @param {string} tag - 探すタグ文字列。
     * @returns {boolean} タグが付いていれば true。
     */
    function hasTag(item, tag) {
        if (!item) return false;
        try { if (item.note === tag) return true; } catch (e) { }
        try { if (item.name === tag) return true; } catch (e) { }
        return false;
    }

    // =========================================
    // 入力値の読み取り / Reading input values
    // =========================================

    /**
     * 入力欄の値を1以上の整数として読み取ります。
     *
     * @param {EditText} field - 読み取る入力欄。
     * @param {number} fallback - 読み取れないときに返す値。
     * @returns {number} 1以上の整数、または fallback。
     */
    function readPositiveInt(field, fallback) {
        try {
            var value = parseInt(field.text, 10);
            if (isNaN(value) || value < 1) return fallback;
            return value;
        } catch (e) { return fallback; }
    }

    /**
     * 入力欄の値を0以上の数値として読み取ります。
     *
     * @param {EditText} field - 読み取る入力欄。
     * @returns {number} 0以上の数値。読み取れないときは 0。
     */
    function readNonNegativeFloat(field) {
        try {
            var value = parseFloat(field.text);
            if (isNaN(value) || value < 0) return 0;
            return value;
        } catch (e) { return 0; }
    }

    /**
     * 外側エリアの辺の伸縮の実効値を返します。外枠なしや無効時は0です。
     *
     * @returns {number} 辺の伸縮の値（定規単位）。
     */
    function getOuterEdgeScaleValue() {
        try {
            if (!chkKeepOuterFrame || !chkKeepOuterFrame.value) return 0;
            if (!chkOuterEdgeScale || !chkOuterEdgeScale.value) return 0;
            var value = parseFloat(editOuterEdgeScale.text);
            return isNaN(value) ? 0 : value;
        } catch (e) {
            return 0;
        }
    }

    /**
     * 入力欄に上下キーでの値の増減を割り当てます（Shiftで10刻み、Optionで0.1刻み）。
     *
     * @param {EditText} field - 対象の入力欄。
     * @param {boolean} allowNegative - 負の値を許可するなら true。
     * @returns {void}
     */
    function changeValueByArrowKey(field, allowNegative) {
        field.addEventListener("keydown", function (event) {
            if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

            var value = Number(field.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === "Up");

            if (keyboard.shiftKey) {
                // 10の倍数にスナップ
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                value = isUp ? value + 0.1 : value - 0.1;
            } else {
                value = isUp ? value + 1 : value - 1;
            }

            if (!allowNegative && value < 0) value = 0;

            value = keyboard.altKey ? (Math.round(value * 10) / 10) : Math.round(value);
            field.text = String(value);

            // keydown で text を書き換えた場合、onChanging が発火しないことがあるため明示的に呼ぶ
            try {
                if (typeof field.onChanging === "function") field.onChanging();
            } catch (e) { }

            try { event.preventDefault(); } catch (e) { }
        });
    }

    // =========================================
    // 有効／無効状態 / Enabled state helpers
    // =========================================

    /**
     * 列数・行数に応じて分割線チェックの有効／無効を切り替えます。
     * 1×1では分割線を作れないため無効になります。
     *
     * @param {number} columnCount - 列数。
     * @param {number} rowCount - 行数。
     * @param {boolean} allowAutoOn - 分割可能になった瞬間に自動ONしてよいなら true。
     * @returns {void}
     */
    function applyInnerDividerEnabledState(columnCount, rowCount, allowAutoOn) {
        var isSplittable = (columnCount > 1 || rowCount > 1);
        try { chkInnerDivider.enabled = isSplittable; } catch (e) { }

        if (!isSplittable) {
            try { chkInnerDivider.value = false; } catch (e) { }
        } else if (allowAutoOn && !wasGridSplittable) {
            // 1×1 から分割可能になった瞬間だけ自動ON
            try { chkInnerDivider.value = true; } catch (e) { }
        }

        wasGridSplittable = isSplittable;

        try { innerLineTypePanel.enabled = (chkInnerDivider.enabled && chkInnerDivider.value); } catch (e) { }
    }

    // =========================================
    // セッション保存 / Persist the session values
    // =========================================
    /**
     * セッションに保持するダイアログの入力値。
     *
     * @typedef {Object} SessionState
     * @property {boolean} preview - プレビューの有効状態。
     * @property {Object} margin - マージンの入力値（top / bottom / left / right / link）。
     * @property {Object} outer - 外側エリアの設定（keepFrame / edgeScaleEnabled / edgeScaleValue / cap / roundEnabled / roundValue）。
     * @property {Object} title - タイトルエリアの設定（enabled / size / position / fill / rule / edgeScaleEnabled / edgeScaleValue）。
     * @property {Object} frame - フレームの設定（enabled / width / bleed / roundEnabled / roundValue）。
     * @property {Object} inner - 内側エリアの設定（link / offset / grid）。
     */
    /**
     * ダイアログの入力値をセッション領域へ保存します。
     *
     * @returns {void}
     */
    function persistSessionState() {
        var state = {};

        // 1箇所が失敗しても他の値は保存されるよう、区分ごとに保護する
        /**
         * 区分ごとの保存処理を実行します。失敗しても他の区分の保存は続行されます。
         *
         * @param {Function} readSectionFn - 1区分ぶんの値を読み取って state に格納する関数。
         * @returns {void}
         */
        function captureSection(readSectionFn) {
            try { readSectionFn(); } catch (e) { }
        }

        captureSection(function () { state.preview = !!chkPreview.value; });

        captureSection(function () {
            state.margin = {
                top: editMarginTop.text,
                bottom: editMarginBottom.text,
                left: editMarginLeft.text,
                right: editMarginRight.text,
                link: !!chkMarginLink.value
            };
        });

        captureSection(function () {
            state.outer = {
                keepFrame: !!chkKeepOuterFrame.value,
                edgeScaleEnabled: !!chkOuterEdgeScale.value,
                edgeScaleValue: editOuterEdgeScale.text,
                cap: (rbCapRound.value ? "round" : (rbCapProject.value ? "project" : "butt")),
                roundEnabled: !!chkOuterRound.value,
                roundValue: editOuterRound.text
            };
        });

        captureSection(function () {
            state.title = {
                enabled: !!chkTitleEnable.value,
                size: editTitleSize.text,
                position: getTitlePositionKey(),
                fill: !!chkTitleFill.value,
                rule: !!chkTitleRule.value,
                edgeScaleEnabled: !!chkTitleEdgeScale.value,
                edgeScaleValue: editTitleEdgeScale.text
            };
        });

        captureSection(function () {
            state.frame = {
                enabled: !!chkFrameEnable.value,
                width: editFrameWidth.text,
                bleed: !!chkBleed.value,
                roundEnabled: !!chkFrameRound.value,
                roundValue: editFrameRound.text
            };
        });

        captureSection(function () {
            state.inner = {
                link: !!chkInnerOffsetLink.value,
                offset: {
                    top: editInnerOffsetTop.text,
                    bottom: editInnerOffsetBottom.text,
                    left: editInnerOffsetLeft.text,
                    right: editInnerOffsetRight.text
                },
                grid: {
                    columns: editColumnCount.text,
                    rows: editRowCount.text,
                    columnGutter: editColumnGutter.text,
                    rowGutter: editRowGutter.text,
                    fill: !!chkInnerCellFill.value,
                    fillManuallySet: !!isInnerFillManuallySet,
                    divider: !!chkInnerDivider.value,
                    lineType: (rbInnerLineDash.value ? "dash" : (rbInnerLineDotted.value ? "dotted" : "solid"))
                }
            };
        });

        saveSessionState(state);
    }

    // =========================================
    // 描画の下ごしらえ / Drawing helpers
    // =========================================

    /**
     * アクティブなアートボードの矩形を返します。裁ち落としが有効なら外側へ広げます。
     *
     * @returns {Array<number>|null} 矩形 [left, top, right, bottom]。取得できないときは null。
     */
    function getActiveArtboardBounds() {
        try {
            var artboardIndex = doc.artboards.getActiveArtboardIndex();
            var artboardRect = doc.artboards[artboardIndex].artboardRect; // [L, T, R, B]
            var boundsLeft = artboardRect[0];
            var boundsTop = artboardRect[1];
            var boundsRight = artboardRect[2];
            var boundsBottom = artboardRect[3];

            if (isBleedEnabled) {
                var bleedPt = getBleedPt();
                boundsLeft -= bleedPt;
                boundsTop += bleedPt;
                boundsRight += bleedPt;
                boundsBottom -= bleedPt;
            }
            return [boundsLeft, boundsTop, boundsRight, boundsBottom];
        } catch (e) {
            return null;
        }
    }

    /**
     * UIで選択されている線端の種類を返します。
     *
     * @returns {number} StrokeCap の値。
     */
    function getSelectedStrokeCap() {
        if (rbCapRound.value) return StrokeCap.ROUNDENDCAP;
        if (rbCapProject.value) return StrokeCap.PROJECTINGENDCAP;
        return StrokeCap.BUTTENDCAP;
    }

    /**
     * ドット点線の見た目（丸型線端＋長さ0のダッシュ）を線に適用します。
     *
     * @param {PathItem} pathItem - 対象の線。
     * @returns {void}
     */
    function applyDottedLineStyle(pathItem) {
        try {
            if (!pathItem.stroked) pathItem.stroked = true;
            pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
            try { pathItem.strokeJoin = StrokeJoin.ROUNDENDJOIN; } catch (e) { }
            pathItem.strokeDashes = [0, pathItem.strokeWidth * 2];
        } catch (e) { }
    }

    /**
     * 内側エリアの分割線に、UIで選択された線種（実線／点線／ドット点線）を適用します。
     *
     * @param {PathItem} pathItem - 対象の線。
     * @returns {void}
     */
    function applyInnerLineStyle(pathItem) {
        try {
            if (rbInnerLineSolid && rbInnerLineSolid.value) {
                try { pathItem.strokeWidth = 1; } catch (e) { }
                try { pathItem.strokeDashes = []; } catch (e) { }
                try { pathItem.strokeCap = getSelectedStrokeCap(); } catch (e) { }
                return;
            }

            if (rbInnerLineDash && rbInnerLineDash.value) {
                try { pathItem.strokeWidth = 1; } catch (e) { }
                try { pathItem.strokeDashes = [pathItem.strokeWidth * 4, pathItem.strokeWidth * 2]; } catch (e) { }
                try { pathItem.strokeCap = getSelectedStrokeCap(); } catch (e) { }
                return;
            }

            if (rbInnerLineDotted && rbInnerLineDotted.value) {
                try { pathItem.strokeWidth = 2; } catch (e) { }
                applyDottedLineStyle(pathItem);
            }
        } catch (e) { }
    }

    // =========================================
    // 生成：フレーム / Generate the frame
    // =========================================

    /**
     * アートボード外周のフレームを、内側が抜けた塗りとして生成します。
     *
     * @param {PathItem} ownerItem - 作成先レイヤーを決めるための基準アイテム。
     * @param {number} framePt - フレームの太さ（pt）。
     * @param {Array<number>} baseBounds - フレームの外形 [left, top, right, bottom]。
     * @returns {void}
     */
    function createFrameFill(ownerItem, framePt, baseBounds) {
        try {
            if (!framePt || framePt <= 0) return;

            var bounds = (baseBounds && baseBounds.length === 4) ? baseBounds : ownerItem.geometricBounds;
            var boundsLeft = bounds[0];
            var boundsTop = bounds[1];
            var boundsRight = bounds[2];
            var boundsBottom = bounds[3];
            var outerWidth = boundsRight - boundsLeft;
            var outerHeight = boundsTop - boundsBottom;
            if (!(outerWidth > 0) || !(outerHeight > 0)) return;

            var innerLeft = boundsLeft + framePt;
            var innerTop = boundsTop - framePt;
            var innerRight = boundsRight - framePt;
            var innerBottom = boundsBottom + framePt;
            var innerWidth = innerRight - innerLeft;
            var innerHeight = innerTop - innerBottom;
            if (!(innerWidth > 0) || !(innerHeight > 0)) return;

            var layer = ownerItem.layer;

            var outerRect = layer.pathItems.rectangle(boundsTop, boundsLeft, outerWidth, outerHeight);
            outerRect.stroked = false;
            outerRect.filled = true;
            outerRect.fillColor = makeBlackTintFill(FRAME_FILL_BLACK_PCT);

            var innerRect = layer.pathItems.rectangle(innerTop, innerLeft, innerWidth, innerHeight);
            innerRect.stroked = false;
            innerRect.filled = true;
            // 内側は一時的な塗り。最終的に Exclude の結果で穴になる
            innerRect.fillColor = makeBlackTintFill(0);

            // 角丸は穴側（内側の長方形）に適用する
            try {
                if (chkFrameRound && chkFrameRound.value) {
                    var radiusValue = parseFloat(editFrameRound.text);
                    if (!isNaN(radiusValue) && radiusValue > 0) {
                        applyRoundCornersEffect(innerRect, radiusValue * getCurrentRulerPtFactor());
                    }
                }
            } catch (e) { }

            var frameGroup = layer.groupItems.add();
            outerRect.move(frameGroup, ElementPlacement.PLACEATEND);
            innerRect.move(frameGroup, ElementPlacement.PLACEATEND);

            // 既存選択を退避し、グループだけ選択して Exclude を実行
            var previousSelection;
            try { previousSelection = doc.selection; } catch (e) { previousSelection = null; }
            try { doc.selection = null; } catch (e) { }
            try { frameGroup.selected = true; } catch (e) { }
            try { app.executeMenuCommand('Live Pathfinder Exclude'); } catch (e) { }

            var resultItem = null;
            try {
                if (doc.selection && doc.selection.length > 0) resultItem = doc.selection[0];
            } catch (e) { }

            try { doc.selection = previousSelection; } catch (e) { }

            if (!resultItem) resultItem = frameGroup;

            try { resultItem.name = "__FrameFill__"; } catch (e) { }
            try { resultItem.note = "__FrameFill__"; } catch (e) { }
            try { resultItem.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }

            previewItems.push(resultItem);
        } catch (e) { }
    }

    // =========================================
    // 生成：タイトルエリア / Generate the title area
    // =========================================

    /**
     * タイトル帯の背景となる塗り矩形を生成します。
     *
     * @param {PathItem} baseItem - 基準となる長方形。
     * @param {number} titleSizePt - タイトルエリアの大きさ（pt）。
     * @returns {void}
     */
    function createTitleFill(baseItem, titleSizePt) {
        try {
            if (!chkTitleFill || !chkTitleFill.value) return;
            if (!titleSizePt || titleSizePt <= 0) return;

            var bounds = baseItem.geometricBounds; // [L, T, R, B]
            var boundsLeft = bounds[0];
            var boundsTop = bounds[1];
            var boundsRight = bounds[2];
            var boundsBottom = bounds[3];
            var baseWidth = boundsRight - boundsLeft;
            var baseHeight = boundsTop - boundsBottom;
            if (!(baseWidth > 0) || !(baseHeight > 0)) return;

            var rectTop, rectLeft, rectWidth, rectHeight;
            var positionKey = getTitlePositionKey();

            if (positionKey === "top") {
                if (titleSizePt >= baseHeight) return;
                rectTop = boundsTop; rectLeft = boundsLeft; rectWidth = baseWidth; rectHeight = titleSizePt;
            } else if (positionKey === "bottom") {
                if (titleSizePt >= baseHeight) return;
                rectTop = boundsBottom + titleSizePt; rectLeft = boundsLeft; rectWidth = baseWidth; rectHeight = titleSizePt;
            } else if (positionKey === "left") {
                if (titleSizePt >= baseWidth) return;
                rectTop = boundsTop; rectLeft = boundsLeft; rectWidth = titleSizePt; rectHeight = baseHeight;
            } else {
                if (titleSizePt >= baseWidth) return;
                rectTop = boundsTop; rectLeft = boundsRight - titleSizePt; rectWidth = titleSizePt; rectHeight = baseHeight;
            }

            var titleRect = baseItem.layer.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
            titleRect.stroked = false;
            titleRect.filled = true;
            titleRect.fillColor = makeBlackTintFill(TITLE_FILL_BLACK_PCT);

            try { titleRect.note = "__TitleFill__"; } catch (e) { }
            try { titleRect.name = "__TitleFill__"; } catch (e) { }

            maybeApplyTitleAreaRound(titleRect);

            // 背面へ（他の罫線や要素の下に敷く）
            try { titleRect.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
            previewItems.push(titleRect);
        } catch (e) { }
    }

    /**
     * タイトル帯と本文を仕切る線を生成します。
     *
     * @param {PathItem} baseItem - 基準となる長方形。
     * @param {number} titleSizePt - タイトルエリアの大きさ（pt）。
     * @param {number} titleDistPt - 線の両端を詰める量（pt）。正で短く、負で長くなります。
     * @returns {void}
     */
    function createTitleDivider(baseItem, titleSizePt, titleDistPt) {
        try {
            if (!baseItem || baseItem.typename !== "PathItem") return;
            if (!titleSizePt || titleSizePt === 0) return;
            if (!(chkTitleRule && chkTitleRule.value)) return;

            var bounds = baseItem.geometricBounds; // [L, T, R, B]
            var boundsLeft = bounds[0];
            var boundsTop = bounds[1];
            var boundsRight = bounds[2];
            var boundsBottom = bounds[3];

            var startX, startY, endX, endY;
            var positionKey = getTitlePositionKey();

            if (positionKey === "left" || positionKey === "right") {
                startX = endX = (positionKey === "left")
                    ? (boundsLeft + titleSizePt)
                    : (boundsRight - titleSizePt);
                startY = boundsTop - titleDistPt;
                endY = boundsBottom + titleDistPt;
                // 短くしすぎて反転する場合は辺いっぱいに戻す
                if (startY <= endY) { startY = boundsTop; endY = boundsBottom; }
            } else {
                startY = endY = (positionKey === "bottom")
                    ? (boundsBottom + titleSizePt)
                    : (boundsTop - titleSizePt);
                startX = boundsLeft + titleDistPt;
                endX = boundsRight - titleDistPt;
                if (startX >= endX) { startX = boundsLeft; endX = boundsRight; }
            }

            var divider = baseItem.layer.pathItems.add();
            divider.stroked = true;
            divider.filled = false;
            try { divider.strokeColor = makeBlackTintFill(100); } catch (e) { }
            try { divider.strokeWidth = 1; } catch (e) { }

            divider.setEntirePath([[startX, startY], [endX, endY]]);

            try { divider.note = "__TitleDivider__"; } catch (e) { }
            try { divider.name = "__TitleDivider__"; } catch (e) { }

            previewItems.push(divider);
        } catch (e) { }
    }

    /**
     * タイトルエリアを除いた内側の計算領域を求めます。
     *
     * @param {PathItem} baseItem - 基準となる長方形。
     * @param {number} titleSizePt - タイトルエリアの大きさ（pt）。
     * @returns {Array<number>|null} 内側領域 [left, top, right, bottom]。求められないときは null。
     */
    function getInnerAreaBounds(baseItem, titleSizePt) {
        try {
            var bounds = baseItem.geometricBounds; // [L, T, R, B]
            var boundsLeft = bounds[0];
            var boundsTop = bounds[1];
            var boundsRight = bounds[2];
            var boundsBottom = bounds[3];
            var baseWidth = boundsRight - boundsLeft;
            var baseHeight = boundsTop - boundsBottom;
            if (!(baseWidth > 0) || !(baseHeight > 0)) return null;

            var titleSize = (titleSizePt > 0) ? titleSizePt : 0;
            if (titleSize <= 0) return [boundsLeft, boundsTop, boundsRight, boundsBottom];

            var positionKey = getTitlePositionKey();
            if (positionKey === "top") {
                if (titleSize >= baseHeight) return null;
                boundsTop -= titleSize;
            } else if (positionKey === "bottom") {
                if (titleSize >= baseHeight) return null;
                boundsBottom += titleSize;
            } else if (positionKey === "left") {
                if (titleSize >= baseWidth) return null;
                boundsLeft += titleSize;
            } else {
                if (titleSize >= baseWidth) return null;
                boundsRight -= titleSize;
            }

            if ((boundsRight - boundsLeft) <= 0 || (boundsTop - boundsBottom) <= 0) return null;
            return [boundsLeft, boundsTop, boundsRight, boundsBottom];
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // 生成：内側エリア / Generate the inner grid
    // =========================================

    /**
     * 内側エリアのセル塗りと分割線を生成します。
     *
     * @param {PathItem} baseItem - 基準となる長方形。
     * @param {Array<number>} baseBounds - 計算に使う領域 [left, top, right, bottom]。
     * @param {DrawSpec} drawSpec - pt に換算済みの描画設定。
     * @returns {void}
     */
    function createInnerGrid(baseItem, baseBounds, drawSpec) {
        var bounds = (baseBounds && baseBounds.length === 4) ? baseBounds : baseItem.geometricBounds;
        var boundsLeft = bounds[0];
        var boundsTop = bounds[1];
        var boundsRight = bounds[2];
        var boundsBottom = bounds[3];

        var offsetTopPt = (drawSpec.offsetTopPt > 0) ? drawSpec.offsetTopPt : 0;
        var offsetBottomPt = (drawSpec.offsetBottomPt > 0) ? drawSpec.offsetBottomPt : 0;
        var offsetLeftPt = (drawSpec.offsetLeftPt > 0) ? drawSpec.offsetLeftPt : 0;
        var offsetRightPt = (drawSpec.offsetRightPt > 0) ? drawSpec.offsetRightPt : 0;

        var areaWidth = (boundsRight - boundsLeft) - (offsetLeftPt + offsetRightPt);
        var areaHeight = (boundsTop - boundsBottom) - (offsetTopPt + offsetBottomPt);
        if (!(areaWidth > 0) || !(areaHeight > 0)) return;

        var areaLeft = boundsLeft + offsetLeftPt;
        var areaTop = boundsTop - offsetTopPt;
        var areaRight = areaLeft + areaWidth;
        var areaBottom = areaTop - areaHeight;

        var columnCount = (drawSpec.columnCount >= 1) ? drawSpec.columnCount : 1;
        var rowCount = (drawSpec.rowCount >= 1) ? drawSpec.rowCount : 1;
        var gutterX = (drawSpec.columnGutterPt > 0) ? drawSpec.columnGutterPt : 0;
        var gutterY = (drawSpec.rowGutterPt > 0) ? drawSpec.rowGutterPt : 0;

        var cellWidth = (areaWidth - gutterX * (columnCount - 1)) / columnCount;
        var cellHeight = (areaHeight - gutterY * (rowCount - 1)) / rowCount;
        if (!(cellWidth > 0) || !(cellHeight > 0)) return;

        // セル塗り（［塗り］ONのときだけ生成し、プレビューと結果を一致させる）
        if (chkInnerCellFill && chkInnerCellFill.value) {
            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                var cellTop = areaTop - (cellHeight + gutterY) * rowIndex;
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                    var cellLeft = areaLeft + (cellWidth + gutterX) * columnIndex;
                    addCellFillRect(baseItem, cellTop, cellLeft, cellWidth, cellHeight);
                }
            }
        }

        if (!(chkInnerDivider && chkInnerDivider.enabled && chkInnerDivider.value)) return;

        // 分割線の見た目：元オブジェクトの線を踏襲。なければ K100 / 1pt
        var lineColor = null;
        var lineWidth = 0;
        try {
            if (baseItem.stroked) {
                lineColor = baseItem.strokeColor;
                lineWidth = baseItem.strokeWidth;
            }
        } catch (e) { }
        if (!lineColor) lineColor = makeBlackTintFill(100);
        if (!lineWidth) lineWidth = 1;

        // 列の分割線：各ガターの中心に1本（ガター0なら境界そのもの）
        for (var dividerColumn = 1; dividerColumn <= columnCount - 1; dividerColumn++) {
            var boundaryX = areaLeft + (cellWidth * dividerColumn) + (gutterX * (dividerColumn - 1));
            var dividerX = boundaryX + (gutterX / 2);
            addDividerLine(baseItem, [[dividerX, areaTop], [dividerX, areaBottom]], lineColor, lineWidth);
        }

        // 行の分割線
        for (var dividerRow = 1; dividerRow <= rowCount - 1; dividerRow++) {
            var boundaryY = areaTop - (cellHeight * dividerRow) - (gutterY * (dividerRow - 1));
            var dividerY = boundaryY - (gutterY / 2);
            addDividerLine(baseItem, [[areaLeft, dividerY], [areaRight, dividerY]], lineColor, lineWidth);
        }
    }

    /**
     * セル1つぶんの背景塗り矩形を生成し、背面へ送ります。
     *
     * @param {PathItem} baseItem - 作成先レイヤーを決めるための基準アイテム。
     * @param {number} cellTop - セル上端のY座標。
     * @param {number} cellLeft - セル左端のX座標。
     * @param {number} cellWidth - セルの幅（pt）。
     * @param {number} cellHeight - セルの高さ（pt）。
     * @returns {void}
     */
    function addCellFillRect(baseItem, cellTop, cellLeft, cellWidth, cellHeight) {
        if (!(cellWidth > 0) || !(cellHeight > 0)) return;
        try {
            var cellRect = baseItem.layer.pathItems.rectangle(cellTop, cellLeft, cellWidth, cellHeight);
            cellRect.stroked = false;
            cellRect.filled = true;
            cellRect.fillColor = makeBlackTintFill(INNER_FILL_BLACK_PCT);

            try { cellRect.note = "__InnerBoxFill__"; } catch (e) { }
            try { cellRect.name = "__InnerBoxFill__"; } catch (e) { }

            // 背面へ（罫線の下に敷く）
            try { cellRect.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
            previewItems.push(cellRect);
        } catch (e) { }
    }

    /**
     * 内側エリアの分割線を1本生成します。
     *
     * @param {PathItem} baseItem - 作成先レイヤーを決めるための基準アイテム。
     * @param {Array<Array<number>>} pathPoints - 線の始点と終点の座標 [[x, y], [x, y]]。
     * @param {CMYKColor} lineColor - 線の色。
     * @param {number} lineWidth - 線の太さ（pt）。
     * @returns {void}
     */
    function addDividerLine(baseItem, pathPoints, lineColor, lineWidth) {
        try {
            var divider = baseItem.layer.pathItems.add();
            divider.setEntirePath(pathPoints);
            divider.stroked = true;
            divider.filled = false;
            divider.strokeColor = lineColor;
            divider.strokeWidth = lineWidth;
            applyInnerLineStyle(divider);
            previewItems.push(divider);
        } catch (e) { }
    }

    // =========================================
    // 生成：外側エリアの4辺線 / Generate the outer edge lines
    // =========================================

    /**
     * 外枠の各辺を、長さを伸縮させた線に置き換えて生成します。
     *
     * @param {PathItem} baseItem - 基準となる長方形。
     * @param {number} distPt - 伸縮量（pt）。正で伸ばし、負で短くします。
     * @returns {void}
     */
    function createOuterEdgeLines(baseItem, distPt) {
        var points = baseItem.pathPoints;
        var segmentCount = baseItem.closed ? points.length : points.length - 1;
        var absDist = Math.abs(distPt);

        for (var segmentIndex = 0; segmentIndex < segmentCount; segmentIndex++) {
            var startPoint = points[segmentIndex].anchor;
            var endPoint = points[(segmentIndex + 1) % points.length].anchor;

            var dx = endPoint[0] - startPoint[0];
            var dy = endPoint[1] - startPoint[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy);

            // 長さ0の辺は比率を計算できない
            if (!(segmentLength > 0)) continue;

            // 短くしすぎて消滅・反転する場合は作らない（伸ばす場合は制限しない）
            if (distPt < 0 && segmentLength <= absDist * 2) continue;

            var ratio = absDist / segmentLength;
            var startX, startY, endX, endY;

            if (distPt >= 0) {
                // 伸ばす
                startX = startPoint[0] - dx * ratio;
                startY = startPoint[1] - dy * ratio;
                endX = endPoint[0] + dx * ratio;
                endY = endPoint[1] + dy * ratio;
            } else {
                // 短くする
                startX = startPoint[0] + dx * ratio;
                startY = startPoint[1] + dy * ratio;
                endX = endPoint[0] - dx * ratio;
                endY = endPoint[1] - dy * ratio;
            }

            var edgeLine = baseItem.layer.pathItems.add();
            edgeLine.setEntirePath([[startX, startY], [endX, endY]]);
            edgeLine.stroked = true;
            edgeLine.filled = false;
            edgeLine.strokeColor = baseItem.strokeColor;
            edgeLine.strokeWidth = baseItem.strokeWidth;
            try { edgeLine.strokeCap = getSelectedStrokeCap(); } catch (e) { }

            try { edgeLine.note = "__OuterEdge__"; } catch (e) { }
            try { edgeLine.name = "__OuterEdge__"; } catch (e) { }

            previewItems.push(edgeLine);
        }
    }

    /**
     * 外側エリアの角丸プレビューを生成します。元オブジェクトは隠し、一時矩形に適用します。
     *
     * @param {number} radiusPt - 角丸の半径（pt）。
     * @returns {void}
     */
    function createOuterRoundPreview(radiusPt) {
        for (var i = 0; i < baseRectItems.length; i++) {
            try {
                var sourceItem = baseRectItems[i];
                if (!sourceItem || sourceItem.typename !== "PathItem") continue;

                var bounds = sourceItem.geometricBounds; // [L, T, R, B]
                var boundsLeft = bounds[0];
                var boundsTop = bounds[1];
                var rectWidth = bounds[2] - bounds[0];
                var rectHeight = bounds[1] - bounds[3];
                if (!(rectWidth > 0) || !(rectHeight > 0)) continue;

                try { sourceItem.hidden = true; } catch (e) { }

                var roundedRect = sourceItem.layer.pathItems.rectangle(boundsTop, boundsLeft, rectWidth, rectHeight);
                try { roundedRect.stroked = sourceItem.stroked; } catch (e) { }
                try { roundedRect.filled = sourceItem.filled; } catch (e) { }
                try { roundedRect.strokeColor = sourceItem.strokeColor; } catch (e) { }
                try { roundedRect.fillColor = sourceItem.fillColor; } catch (e) { }
                try { roundedRect.strokeWidth = sourceItem.strokeWidth; } catch (e) { }

                try { roundedRect.note = "__OuterRoundPreview__"; } catch (e) { }
                try { roundedRect.name = "__OuterRoundPreview__"; } catch (e) { }

                applyRoundCornersEffect(roundedRect, radiusPt);
                previewItems.push(roundedRect);
            } catch (e) { }
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * pt に換算済みの描画設定。
     *
     * @typedef {Object} DrawSpec
     * @property {number} factor - 定規単位1あたりのポイント数。
     * @property {number} edgeScalePt - 外側エリアの辺の伸縮量（pt）。
     * @property {number} titleSizePt - タイトルエリアの大きさ（pt）。
     * @property {number} titleDistPt - タイトル境界線の両端を詰める量（pt）。
     * @property {number} framePt - 裁ち落としを含むフレームの太さ（pt）。
     * @property {number} offsetTopPt - 内側エリアの上オフセット（pt）。
     * @property {number} offsetBottomPt - 内側エリアの下オフセット（pt）。
     * @property {number} offsetLeftPt - 内側エリアの左オフセット（pt）。
     * @property {number} offsetRightPt - 内側エリアの右オフセット（pt）。
     * @property {number} columnCount - 列数。
     * @property {number} rowCount - 行数。
     * @property {number} columnGutterPt - 列の間隔（pt）。
     * @property {number} rowGutterPt - 行の間隔（pt）。
     */
    /**
     * ダイアログの入力値をすべて読み取り、pt に換算した描画設定にまとめます。
     *
     * @returns {DrawSpec} pt に換算済みの描画設定。
     */
    function collectDrawSpec() {
        var factor = getCurrentRulerPtFactor();

        var columnCount = readPositiveInt(editColumnCount, 1);
        var rowCount = readPositiveInt(editRowCount, 1);

        try { editColumnGutter.enabled = (columnCount > 1); } catch (e) { }
        try { editRowGutter.enabled = (rowCount > 1); } catch (e) { }

        var columnGutter = parseFloat(editColumnGutter.text);
        var rowGutter = parseFloat(editRowGutter.text);
        var columnGutterPt = (columnCount > 1 && !isNaN(columnGutter) && columnGutter > 0) ? (columnGutter * factor) : 0;
        var rowGutterPt = (rowCount > 1 && !isNaN(rowGutter) && rowGutter > 0) ? (rowGutter * factor) : 0;

        // ガターが設定されたら塗りを自動ON（ただし手動操作があれば尊重）
        if (!isInnerFillManuallySet && (columnGutterPt > 0 || rowGutterPt > 0)) {
            try { if (!chkInnerCellFill.value) chkInnerCellFill.value = true; } catch (e) { }
        }

        applyInnerDividerEnabledState(columnCount, rowCount, false);
        applyOuterCapEnabledState();

        // タイトル帯の線の伸縮は正負を反転（＋で伸ばす／−で短くする）
        var titleEdgeScale = 0;
        try {
            if (chkTitleRule && chkTitleRule.value && chkTitleEdgeScale && chkTitleEdgeScale.value) {
                var scaleValue = parseFloat(editTitleEdgeScale.text);
                if (!isNaN(scaleValue)) titleEdgeScale = scaleValue;
            }
        } catch (e) { }

        var titleSize = (chkTitleEnable && !chkTitleEnable.value) ? 0 : parseFloat(editTitleSize.text);
        var titleSizePt = (!isNaN(titleSize) && titleSize > 0) ? (titleSize * factor) : 0;

        var frameWidth = (chkFrameEnable && chkFrameEnable.value) ? readNonNegativeFloat(editFrameWidth) : 0;
        // フレームを作らないときは裁ち落としも加算しない（幻のフレームが出るのを防ぐ）
        var framePt = frameWidth * factor;
        if (framePt > 0 && isBleedEnabled) framePt += getBleedPt();

        return {
            factor: factor,
            edgeScalePt: getOuterEdgeScaleValue() * factor,
            titleSizePt: titleSizePt,
            titleDistPt: (-titleEdgeScale) * factor,
            framePt: framePt,
            offsetTopPt: readNonNegativeFloat(editInnerOffsetTop) * factor,
            offsetBottomPt: readNonNegativeFloat(editInnerOffsetBottom) * factor,
            offsetLeftPt: readNonNegativeFloat(editInnerOffsetLeft) * factor,
            offsetRightPt: readNonNegativeFloat(editInnerOffsetRight) * factor,
            columnCount: columnCount,
            rowCount: rowCount,
            columnGutterPt: columnGutterPt,
            rowGutterPt: rowGutterPt
        };
    }

    /**
     * 現在のマージン入力値に合わせて、アートボード基準の一時矩形を作り直します。
     *
     * @param {number} factor - 定規単位1あたりのポイント数。
     * @returns {void}
     */
    function refreshArtboardBaseRect(factor) {
        if (!isArtboardBase) return;

        var marginTopPt = readNonNegativeFloat(editMarginTop) * factor;
        var marginBottomPt = readNonNegativeFloat(editMarginBottom) * factor;
        var marginLeftPt = readNonNegativeFloat(editMarginLeft) * factor;
        var marginRightPt = readNonNegativeFloat(editMarginRight) * factor;

        try { if (chkBleed) isBleedEnabled = !!chkBleed.value; } catch (e) { }
        rebuildArtboardBaseRect(marginTopPt, marginRightPt, marginBottomPt, marginLeftPt);
    }

    /**
     * 現在の設定でカンバス上に描画します。プレビューと最終生成の両方で使います。
     *
     * @param {boolean} isFinal - OK時の最終生成なら true、プレビューなら false。
     * @returns {void}
     */
    function updatePreview(isFinal) {
        removeTempItems();

        var drawSpec = collectDrawSpec();
        refreshArtboardBaseRect(drawSpec.factor);

        var i;
        var isSplittingEdges = (drawSpec.edgeScalePt !== 0);

        if (isSplittingEdges) {
            // 4辺線で外枠を表すため、元の長方形は常に隠す
            for (i = 0; i < baseRectItems.length; i++) {
                try { baseRectItems[i].hidden = true; } catch (e) { }
            }
            for (i = 0; i < baseRectItems.length; i++) {
                createOuterEdgeLines(baseRectItems[i], drawSpec.edgeScalePt);
            }
        } else {
            // 外枠の表示は［外枠を残す］に従う
            var shouldShowOuterRect = !!chkKeepOuterFrame.value;
            for (i = 0; i < baseRectItems.length; i++) {
                try { baseRectItems[i].hidden = !shouldShowOuterRect; } catch (e) { }
            }
            applyOuterRoundIfNeeded(shouldShowOuterRect, isFinal);
        }

        // フレーム（アートボードサイズ基準。代表として先頭アイテムのレイヤーに作成）
        if (drawSpec.framePt > 0 && baseRectItems.length > 0) {
            var artboardBounds = getActiveArtboardBounds();
            if (artboardBounds) createFrameFill(baseRectItems[0], drawSpec.framePt, artboardBounds);
        }

        // タイトルエリア
        if (drawSpec.titleSizePt > 0) {
            for (i = 0; i < baseRectItems.length; i++) {
                createTitleFill(baseRectItems[i], drawSpec.titleSizePt);
            }
            for (i = 0; i < baseRectItems.length; i++) {
                createTitleDivider(baseRectItems[i], drawSpec.titleSizePt, drawSpec.titleDistPt);
            }
        }

        // 内側エリア（オフセットが0でも描画）
        for (i = 0; i < baseRectItems.length; i++) {
            var innerBounds = getInnerAreaBounds(baseRectItems[i], drawSpec.titleSizePt);
            if (innerBounds) createInnerGrid(baseRectItems[i], innerBounds, drawSpec);
        }

        app.redraw();
    }

    /**
     * 外側エリアの角丸を適用します。辺の伸縮が有効なときは何もしません。
     *
     * @param {boolean} shouldShowOuterRect - 外枠を表示する設定なら true。
     * @param {boolean} isFinal - 最終生成なら true、プレビューなら false。
     * @returns {void}
     */
    function applyOuterRoundIfNeeded(shouldShowOuterRect, isFinal) {
        try {
            if (!shouldShowOuterRect) return;
            if (!chkOuterRound || !chkOuterRound.value) return;
            if (chkOuterEdgeScale && chkOuterEdgeScale.value) return;

            var radiusValue = parseFloat(editOuterRound.text);
            if (isNaN(radiusValue) || radiusValue <= 0) return;

            var radiusPt = radiusValue * getCurrentRulerPtFactor();
            if (!(radiusPt > 0)) return;

            if (isFinal) {
                // OK時：元の外枠にライブエフェクトとして適用
                for (var i = 0; i < baseRectItems.length; i++) {
                    try {
                        var item = baseRectItems[i];
                        if (!item || item.typename !== "PathItem") continue;
                        applyRoundCornersEffect(item, radiusPt);
                        item.hidden = false;
                    } catch (e) { }
                }
            } else {
                createOuterRoundPreview(radiusPt);
            }
        } catch (e) { }
    }

    /**
     * プレビュー生成物を消し、隠していた元オブジェクトを表示に戻します。
     *
     * @returns {void}
     */
    function clearPreview() {
        removeTempItems();

        // アートボード基準の一時矩形は先に破棄（baseRectItems に残ると無効参照になる）
        disposeArtboardBaseRect();

        for (var i = 0; i < baseRectItems.length; i++) {
            try { baseRectItems[i].hidden = false; } catch (e) { }
        }
        try { applyOuterCapEnabledState(); } catch (e) { }

        app.redraw();
    }

    /**
     * プレビュー用に生成した一時アイテムをすべて削除します。
     *
     * @returns {void}
     */
    function removeTempItems() {
        for (var i = 0; i < previewItems.length; i++) {
            try {
                previewItems[i].remove();
            } catch (e) {
                // すでに削除済みの場合など。無視してよい
            }
        }
        previewItems = [];
    }

})();
