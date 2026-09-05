#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

キーオブジェクトを起点に、選択した各図形へコネクターを引きます。
直線・ワープ・カギの3種類の線に、線の設定・矢印・白フチをプレビューしながら設定できます。

詳細は README を参照してください。

### Overview

Draws a connector from the key object to each of the selected objects.
Choose a straight, warped, or elbow line and set the stroke, arrowheads, and a white outline with a live preview.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiConnectorBuilder";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-05";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiConnectorBuilder.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiConnectorBuilder.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var CONNECTOR_LAYER_NAME = "コネクター";  /* コネクターの作成先レイヤー名 / target layer name */

    /* 線の初期値 / Line defaults */
    var DEFAULT_STROKE_WIDTH  = 1;    /* 線幅（pt） / stroke width */
    var DEFAULT_STROKE_JOIN   = 0;    /* 0=マイター 1=ラウンド 2=ベベル */
    var DEFAULT_DASH_STYLE    = 0;    /* 0=なし 1=破線 2=ドット */
    var DEFAULT_DASH_SEGMENTS = 8;    /* 分割数（線分・ドットの数） */
    var DEFAULT_DASH_GAP      = 3;    /* 破線の間隔（pt） */

    /* コネクターの初期値 / Connector defaults */
    var DEFAULT_LINE_SHAPE    = 1;    /* 0=なし（直線） 1=ワープ 2=カギ 3=分岐 */
    var DEFAULT_START_POINT   = 0;    /* 0=中心のみ 1=等分 */
    var DEFAULT_WARP_TYPE     = 0;    /* WARP_TYPE_CHOICES のインデックス（0=でこぼこ） */
    var DEFAULT_WARP_AMOUNT   = -80;  /* カーブ（%） */
    var WARP_AMOUNT_MIN       = -100; /* カーブの下限（%） */
    var WARP_AMOUNT_MAX       = 100;  /* カーブの上限（%） */
    var DEFAULT_WARP_AXIS     = 0;    /* 0=自動 1=水平 2=垂直 */
    var DEFAULT_CORNER_RADIUS = 10;   /* カギの角丸半径（pt）／0で角丸なし */
    var WARP_DEFORM_H         = 0;    /* 変形・水平方向（%）：XMLで必須のため固定値で渡す */
    var WARP_DEFORM_V         = 0;    /* 変形・垂直方向（%）：同上 */

    /* 矢印の初期値 / Arrowhead defaults */
    var DEFAULT_ARROW_INDEX  = 1;    /* ARROW_CHOICES のインデックス（0=なし） */
    var DEFAULT_ARROW_SCALE  = 100;  /* 矢印の倍率（%）：ARROW_CHOICES に指定がないときの値 */
    var DEFAULT_ARROW_POSITION = 0;  /* 0=終点のみ 1=両端 */
    /* 黒丸に使う矢印番号（［線］パネルの矢印リストに合わせて調整） */
    var ARROW_DOT_FILLED = 21;       /* 黒丸 */
    /* 白丸は黒丸の中心に白い●を重ねて作る */
    var WHITE_DOT_RATIO = 1.5;       /* 白い●の直径＝線幅の何倍か */

    /* 一時アクション / Temporary action（矢印はDOMから設定できないためアクションで適用） */
    var ACTION_SET_NAME  = "SwwwitchTempConnectorSet";
    var ACTION_NAME      = "SwwwitchTempConnector";
    var ACTION_FILE_NAME = File(Folder.temp).fsName + "/swwwitch_temp_connector.aia";

    /* パラメータキー / Parameter keys（記録した .aia から採取） */
    var KEY_STROKE_WIDTH  = 2003072104;      /* 線幅 / stroke width */
    var KEY_CAP           = 1667330094;      /* 線端 / cap */
    var KEY_JOIN          = 1785686382;      /* 角の形状 / join */
    var KEY_DASH_INT      = 1684825454;      /* 破線（整数）/ dash (integer) */
    var KEY_DASH_BOOL     = 1684104298;      /* 破線（真偽）/ dash (boolean) */
    var KEY_ARROW_HEAD_1  = 1634231345;      /* ahd1: 始点の形状 / start arrowhead */
    var KEY_ARROW_HEAD_2  = 1634231346;      /* ahd2: 終点の形状 / end arrowhead */
    var KEY_ARROW_SCALE_1 = 1634951985;      /* asc1: 始点の倍率 / start scale */
    var KEY_ARROW_SCALE_2 = 1634951986;      /* asc2: 終点の倍率 / end scale */
    var KEY_ARROW_ALIGN   = 1634230636;      /* ahal: 矢印の配置 / tip alignment */
    var KEY_ALIGN         = 1634494318;      /* algn: 線の位置 / stroke alignment */
    var UNIT_POINT        = 592476268;       /* ポイント / point（parameter /unit） */


    /* 判定の許容値 / Tolerances */
    var KEY_DETECT_TOLERANCE_PT = 0.001; /* 整列後に「動いていない」とみなす差（pt） */
    var COORD_TOLERANCE_PT      = 0.001; /* 座標が同じとみなす差（pt） */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 */
    var LABEL_WIDTH        = 84;             /* コネクターパネルの行ラベル幅（右揃え） */
    var COLUMN_LABEL_WIDTH = 58;             /* 線・矢印パネルの行ラベル幅（2カラムなので狭め） */
    var FIELD_CHARS    = 4;                  /* 数値欄の文字数 */
    var LIST_WIDTH     = 150;                /* ドロップダウンの幅 */
    var SLIDER_MIN_WIDTH = 60;               /* スライダーの最小幅（余白は fill で伸ばす） */
    var RADIO_COLUMN_SPACING = 4;            /* 縦並びラジオの間隔 */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 */
    var PRESET_BUTTON_WIDTH = 60;            /* プリセットの保存・削除ボタンの幅 */
    var BUTTON_ROW_TOP_MARGIN = 5;           /* ボタン行の上余白 */

    /* 行ラベルの幅。パネルごとに setLabelWidth() で切り替える */
    var currentLabelWidth = LABEL_WIDTH;

    /**
     * 以降に作る行ラベルの幅を切り替える
     * @param {number} width - 行ラベルの幅
     * @returns {void}
     */
    function setLabelWidth(width) {
        currentLabelWidth = width;
    }

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(win) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {object} panel - 対象のパネル
     * @param {number} spacing - 要素間隔（省略時は共通値）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 見出し付きのパネルを追加する
     * @param {object} parent - 追加先
     * @param {string} title - パネルの見出し
     * @returns {object} 追加したパネル
     */
    function addPanel(parent, title) {
        var panel = parent.add("panel", undefined, title);
        setupPanel(panel);
        return panel;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "コネクター", en: "Connector" }
        },
        panel: {
            connector: { ja: "コネクター", en: "Connector" },
            line:   { ja: "線", en: "Line" },
            arrow:  { ja: "矢印", en: "Arrowheads" }
        },
        fieldLabel: {
            preset:         { ja: "プリセット", en: "Preset" },
            strokeWidth:    { ja: "線幅", en: "Stroke width" },
            strokeJoin:     { ja: "角の形状", en: "Corner" },
            dashStyle:      { ja: "破線", en: "Dashes" },
            dashSegments:   { ja: "分割数", en: "Divisions" },
            dashGap:        { ja: "間隔", en: "Gap" },
            lineShape:      { ja: "形状", en: "Shape" },
            startPoint:     { ja: "開始ポイント", en: "Start point" },
            warpType:       { ja: "種類", en: "Style" },
            warpAmount:     { ja: "カーブ", en: "Bend" },
            warpAxis:       { ja: "方向", en: "Axis" },
            cornerRadius:   { ja: "角丸", en: "Round corners" },
            arrowShape:     { ja: "形状", en: "Shape" },
            arrowScale:     { ja: "倍率", en: "Scale" },
            arrowPosition:  { ja: "位置", en: "Position" },
            arrowTip:       { ja: "先端位置", en: "Tip" }
        },
        radio: {
            joinMiter:      { ja: "マイター", en: "Miter" },
            joinRound:      { ja: "ラウンド", en: "Round" },
            joinBevel:      { ja: "ベベル", en: "Bevel" },
            dashNone:       { ja: "なし", en: "None" },
            dashDashed:     { ja: "破線", en: "Dashed" },
            dashDotted:     { ja: "ドット", en: "Dotted" },
            startCenter:    { ja: "中心のみ", en: "Center only" },
            startDivided:   { ja: "等分", en: "Divided" },
            shapeNone:      { ja: "なし", en: "None" },
            shapeWarp:      { ja: "ワープ", en: "Warp" },
            shapeElbow:     { ja: "カギ", en: "Elbow" },
            shapeBranch:    { ja: "分岐", en: "Branch" },
            axisAuto:       { ja: "自動", en: "Auto" },
            axisHorizontal: { ja: "水平", en: "Horizontal" },
            axisVertical:   { ja: "垂直", en: "Vertical" },
            arrowNone:      { ja: "なし", en: "None" },
            arrow8:         { ja: "矢印8", en: "Arrow 8" },
            arrow11:        { ja: "矢印11", en: "Arrow 11" },
            dotFilled:      { ja: "黒丸", en: "Dot" },
            dotHollow:      { ja: "白丸", en: "Circle" },
            arrowEnd:       { ja: "終点", en: "End" },
            tipAtEnd:       { ja: "終点に", en: "At end" },
            tipBeyondEnd:   { ja: "終点から", en: "Beyond" },
            arrowBoth:      { ja: "両端", en: "Both ends" }
        },
        arrow: {
            none:   { ja: "[なし]", en: "[None]" },
            prefix: { ja: "矢印 ", en: "Arrow " },
            /* アクションに埋め込む Illustrator の表示名 / labels embedded in the action */
            tipAtEndName:     { ja: "パスの終点に配置", en: "Place Arrow Tip At End of Path" },
            tipBeyondEndName: { ja: "パスの終点から配置", en: "Extend Arrow Tip Beyond End of Path" }
        },
        unit: {
            blank:   { ja: "", en: "" },
            pt:      { ja: "pt", en: "pt" },
            percent: { ja: "%", en: "%" }
        },
        button: {
            save:   { ja: "保存", en: "Save" },
            remove: { ja: "削除", en: "Delete" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        preset: {
            custom: { ja: "（カスタム）", en: "(Custom)" }
        },
        tooltip: {
            strokeJoin:     { ja: "カギ・分岐の折れ角の見え方です（［線］パネルの角の形状）。", en: "How the elbow corners look (the Stroke panel's corner setting)." },
            dashSegments:   { ja: "線分（ドット）の数です。両端が線分で終わるように線分の長さを計算します。", en: "Number of dashes (dots); the dash length is solved so both ends finish with a dash." },
            dashGap:        { ja: "線分どうしのすき間です。ドットでは分割数から自動で決まります。", en: "Gap between dashes. For dots it is derived from the number of divisions." },
            startPoint:     { ja: "等分は、同じ辺から出るコネクターの本数＋1でその辺を等分し、起点をずらします。", en: "Divided spreads the start points along the key object's edge, splitting it into (connectors + 1) parts." },
            lineShape:      { ja: "なしは直線、ワープは直線にワープ効果、カギは直角に折れる線、分岐は折れ位置をそろえて幹を共有します。", en: "None = straight, Warp = straight plus a warp effect, Elbow = right-angled route, Branch = shared trunk with aligned bends." },
            warpAmount:     { ja: "マイナス値で曲がる向きが逆になります。", en: "A negative value bends the other way." },
            warpAxis:       { ja: "自動は線の傾きから水平／垂直を選びます。線に沿った向きのワープは曲がりません。", en: "Auto picks horizontal / vertical from the line's slope; warping along the line has no visible effect." },
            cornerRadius:   { ja: "カギの角を丸めます（0で角丸なし）。分岐では使いません。", en: "Round the elbow corners (0 = square corners). Not used by Branch." },
            preset:         { ja: "現在の設定に名前を付けて保存できます。保存先はユーザーの設定フォルダーです。", en: "Save the current settings under a name; presets are stored in your user settings folder." },
            arrowShape:     { ja: "［線］パネルの矢印を使います。黒丸も矢印の一種です。", en: "Uses the Stroke panel arrowheads; the dot is an arrowhead preset too." },
            arrowScale:     { ja: "矢印の大きさ（%）。線幅に対する比率です。", en: "Arrowhead size in percent, relative to the stroke width." },
            arrowTip:       { ja: "矢印の先端をパスの終点に配置するか、パスの終点から配置するかを選びます。", en: "Place the arrow tip at the end of the path, or extend it beyond the end." },
            arrowPosition:  { ja: "終点はキーオブジェクトと反対側、両端は起点にも付けます。", en: "End = the far side from the key object; Both ends also marks the start." }
        },
        alert: {
            presetName:    { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
            presetRemove:  { ja: "このプリセットを削除しますか？", en: "Delete this preset?" },
            presetFailed:  { ja: "プリセットを保存できませんでした。", en: "Could not save the presets." },
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectObjects: { ja: "2つ以上の図形を選択してください。", en: "Please select two or more objects." },
            actionFailed:  { ja: "一時アクションファイルを開けませんでした。", en: "Failed to open the temporary action file." },
            noKeyObject:   { ja: "キーオブジェクトが設定されていません。\n選択したうえで、基準にしたいオブジェクトをもう一度クリックしてください。", en: "No key object is set.\nWith the objects selected, click the one you want as the key again." }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {object} labelSet - { ja, en } のラベル定義
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        return labelSet[uiLang] || labelSet.en || "";
    }

    /**
     * コロン付きラベルを返す（日本語は全角、英語は半角）
     * @param {object} labelSet - { ja, en } のラベル定義
     * @returns {string} コロン付きラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    /* 使用する矢印。number=矢印番号、scale=倍率（%）、tip=先端位置（0=終点に 1=終点から） */
    /* scale と tip は、その矢印を選んだときに入れる既定値 */
    var ARROW_CHOICES = [
        { label: LABELS.radio.arrowNone, number: 0,                scale: 100, tip: 0 },
        { label: LABELS.radio.arrow8,    number: 8,                scale: 25,  tip: 0 },
        { label: LABELS.radio.arrow11,   number: 11,               scale: 100, tip: 0 },
        { label: LABELS.radio.dotFilled, number: ARROW_DOT_FILLED, scale: 50,  tip: 1 },
        { label: LABELS.radio.dotHollow, number: ARROW_DOT_FILLED, scale: 50,  tip: 1, innerDot: true }
    ];

    /* 角の形状 / Stroke join */
    var STROKE_JOIN_OPTIONS = [
        { label: LABELS.radio.joinMiter, value: StrokeJoin.MITERENDJOIN },
        { label: LABELS.radio.joinRound, value: StrokeJoin.ROUNDENDJOIN },
        { label: LABELS.radio.joinBevel, value: StrokeJoin.BEVELENDJOIN }
    ];

    /* 矢印の先端位置 / Arrow tip alignment（ahal の enumerated 値） */
    var ARROW_TIP_OPTIONS = [
        { label: LABELS.radio.tipAtEnd,     name: LABELS.arrow.tipAtEndName,     value: 0 },
        { label: LABELS.radio.tipBeyondEnd, name: LABELS.arrow.tipBeyondEndName, value: 1 }
    ];

    /* ダイアログで選べるワープの種類 / Warp styles offered in the dialog */
    /* style: ワープ効果の並び順（1始まり）／name: ライブエフェクトXMLで使う名前 */
    var WARP_TYPE_CHOICES = [
        { style: 5,  name: "Bulge",   ja: "でこぼこ", en: "Bulge" },
        { style: 14, name: "Squeeze", ja: "絞り込み", en: "Squeeze" }
    ];

    /**
     * ワープの種類の一覧を、現在の言語の表示名で返す
     * @returns {Array<string>} 表示名の配列
     */
    function getWarpTypeLabels() {
        var labels = [];
        for (var i = 0; i < WARP_TYPE_CHOICES.length; i++) {
            labels.push(getLabel(WARP_TYPE_CHOICES[i]));
        }
        return labels;
    }

    // =========================================
    // 選択オブジェクト / Selection
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length < 2) {
        alert(getLabel(LABELS.alert.selectObjects));
        return;
    }

    var selectedItems = [];
    for (var i = 0; i < doc.selection.length; i++) {
        selectedItems.push(doc.selection[i]);
    }

    // Illustratorの座標系（visibleBounds）：
    // 左 = bounds[0] / 上 = bounds[1] / 右 = bounds[2] / 下 = bounds[3]

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================

    /**
     * スマートガイドの表示を切り替える（実行前後で同じ状態に戻す）
     * @returns {void}
     */
    function toggleSmartGuides() {
        try {
            app.executeMenuCommand("edge");
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": スマートガイドの切り替えに失敗 / failed to toggle Smart Guides — " + e);
        }
    }

    /**
     * 控えておいた位置へ戻す（1つ失敗しても残りは戻す）
     * @param {Array<object>} items - 対象のオブジェクト配列
     * @param {Array<Array<number>>} positions - [[left, top], ...] の配列
     * @returns {void}
     */
    function restorePositions(items, positions) {
        for (var i = 0; i < items.length; i++) {
            try {
                items[i].left = positions[i][0];
                items[i].top = positions[i][1];
            } catch (e) {
                $.writeln(SCRIPT_NAME + ": 位置の復元に失敗 / failed to restore position — " + e);
            }
        }
    }

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * DOMにキーオブジェクトを示すプロパティは無いため、整列コマンドを実行して
     * 「どの向きに整列しても動かないもの」を実測で特定する。
     * @param {Array<object>} items - 判定対象のオブジェクト配列
     * @returns {number} キーオブジェクトのインデックス。判定できないときは -1
     */
    function detectKeyObjectIndex(items) {
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        var originPositions = [];
        var i;
        for (i = 0; i < items.length; i++) {
            stayedPut.push(true);
            originPositions.push([items[i].left, items[i].top]);
        }

        try {
            for (var c = 0; c < alignCommands.length; c++) {
                app.redraw(); // 直前のDOM変更が反映されていないと executeMenuCommand は空振りする
                app.executeMenuCommand(alignCommands[c]);
                for (i = 0; i < items.length; i++) {
                    if (Math.abs(items[i].left - originPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                        Math.abs(items[i].top - originPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                        stayedPut[i] = false;
                    }
                }
                // 検出のための試行なので、毎回その場で元の位置へ戻す
                restorePositions(items, originPositions);
            }
        } finally {
            // 例外で抜けるときも整列結果を残さない
            restorePositions(items, originPositions);
        }
        app.redraw();

        var foundIndex = -1;
        for (i = 0; i < items.length; i++) {
            if (!stayedPut[i]) continue;
            if (foundIndex !== -1) return -1; // 複数残った＝判定不能
            foundIndex = i;
        }
        return foundIndex;
    }

    /**
     * 中心Xがもっとも左にあるオブジェクトのインデックスを返す
     * @param {Array<object>} items - 対象のオブジェクト配列
     * @returns {number} インデックス
     */
    function getLeftmostIndex(items) {
        var leftmostIndex = 0;
        var minCenterX = null;
        for (var i = 0; i < items.length; i++) {
            var bounds = items[i].visibleBounds;
            var centerX = (bounds[0] + bounds[2]) / 2;
            if (minCenterX === null || centerX < minCenterX) {
                minCenterX = centerX;
                leftmostIndex = i;
            }
        }
        return leftmostIndex;
    }

    // =========================================
    // 経路の計算 / Connector geometry
    // =========================================

    /**
     * 2つの外接矩形から、向かい合う辺の中央どうしを結ぶ始点・終点を求める
     * @param {Array<number>} fromBounds - 始点側の visibleBounds
     * @param {Array<number>} toBounds - 終点側の visibleBounds
     * @returns {object} points（始点・終点）と horizontal（左右の辺どうしか）
     */
    function getConnectionPoints(fromBounds, toBounds) {
        var fromCenterX = (fromBounds[0] + fromBounds[2]) / 2;
        var fromCenterY = (fromBounds[1] + fromBounds[3]) / 2;
        var toCenterX = (toBounds[0] + toBounds[2]) / 2;
        var toCenterY = (toBounds[1] + toBounds[3]) / 2;

        var dx = toCenterX - fromCenterX;
        var dy = toCenterY - fromCenterY;

        // 横のずれが大きければ左右の辺、そうでなければ上下の辺でつなぐ
        var route;
        if (Math.abs(dx) >= Math.abs(dy)) {
            route = (dx >= 0)
                ? { points: [[fromBounds[2], fromCenterY], [toBounds[0], toCenterY]], horizontal: true, side: "right" }
                : { points: [[fromBounds[0], fromCenterY], [toBounds[2], toCenterY]], horizontal: true, side: "left" };
        } else {
            route = (dy >= 0)
                ? { points: [[fromCenterX, fromBounds[1]], [toCenterX, toBounds[3]]], horizontal: false, side: "top" }
                : { points: [[fromCenterX, fromBounds[3]], [toCenterX, toBounds[1]]], horizontal: false, side: "bottom" };
        }
        // 等分配置で辺に沿って並べ替えるための基準
        route.order = route.horizontal ? toCenterY : toCenterX;
        return route;
    }

    /**
     * 開始ポイントの指定を反映した経路の複製を返す
     * @param {number} startPoint - 0=中心のみ 1=等分
     * @returns {Array<object>} points（座標）と horizontal（左右接続か）の配列
     */
    function getRoutes(startPoint) {
        var routes = [];
        var i;
        for (i = 0; i < connectorPaths.length; i++) {
            var path = connectorPaths[i];
            routes.push({
                points: [[path.points[0][0], path.points[0][1]], [path.points[1][0], path.points[1][1]]],
                horizontal: path.horizontal,
                side: path.side,
                order: path.order
            });
        }
        if (startPoint !== 1) return routes;

        // 同じ辺から出るコネクターごとに、その辺を（本数＋1）等分して起点をずらす
        var groups = {};
        for (i = 0; i < routes.length; i++) {
            if (!groups[routes[i].side]) groups[routes[i].side] = [];
            groups[routes[i].side].push(routes[i]);
        }
        for (var side in groups) {
            if (!groups.hasOwnProperty(side)) continue;
            var group = groups[side];
            // 線が交差しないよう、相手の位置順に辺へ割り当てる
            group.sort(function (a, b) {
                return a.order - b.order;
            });
            for (var k = 0; k < group.length; k++) {
                var ratio = (k + 1) / (group.length + 1);
                if (group[k].horizontal) {
                    group[k].points[0][1] = keyBounds[3] + (keyBounds[1] - keyBounds[3]) * ratio;
                } else {
                    group[k].points[0][0] = keyBounds[0] + (keyBounds[2] - keyBounds[0]) * ratio;
                }
            }
        }
        return routes;
    }

    /**
     * カギ線（直角に折れる経路）の座標列を作る
     * @param {Array<Array<number>>} points - [[始点X, 始点Y], [終点X, 終点Y]]
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか
     * @param {number} bendPosition - 折れ位置の座標。省略時はすき間の中央
     * @returns {Array<Array<number>>} 座標の配列
     */
    function getElbowPoints(points, horizontal, bendPosition) {
        var start = points[0];
        var end = points[1];
        var hasBend = (typeof bendPosition === "number");
        if (horizontal) {
            // 折れ位置は2つの図形のすき間の中央
            if (Math.abs(end[1] - start[1]) < COORD_TOLERANCE_PT) return [start, end];
            var bendX = hasBend ? bendPosition : (start[0] + end[0]) / 2;
            return [start, [bendX, start[1]], [bendX, end[1]], end];
        }
        if (Math.abs(end[0] - start[0]) < COORD_TOLERANCE_PT) return [start, end];
        var bendY = hasBend ? bendPosition : (start[1] + end[1]) / 2;
        return [start, [start[0], bendY], [end[0], bendY], end];
    }

    /**
     * 分岐用に、同じ辺から出る経路で共有する折れ位置を求める
     * いちばん近い図形とのすき間の中央にそろえ、幹を1本にまとめる
     * @param {Array<object>} routes - getRoutes() の戻り値
     * @returns {object} 辺をキーにした折れ位置
     */
    function getSharedBendPositions(routes) {
        var bendPositions = {};
        var nearestDeltas = {};
        var i;
        for (i = 0; i < routes.length; i++) {
            var route = routes[i];
            var axis = route.horizontal ? 0 : 1;
            var delta = route.points[1][axis] - route.points[0][axis];
            if (nearestDeltas[route.side] === undefined || Math.abs(delta) < Math.abs(nearestDeltas[route.side])) {
                nearestDeltas[route.side] = delta;
                bendPositions[route.side] = route.points[0][axis] + delta / 2;
            }
        }
        return bendPositions;
    }

    // =========================================
    // 作図 / Drawing
    // =========================================

    /**
     * 指定名のレイヤーを取得する。無ければ作成する
     * @param {string} name - レイヤー名
     * @returns {object} layer（レイヤー）、existed（既存だったか）、locked／visible（変更前の状態）
     */
    function getOrCreateLayer(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                var existingLayer = doc.layers[i];
                var layerState = {
                    layer: existingLayer,
                    existed: true,
                    locked: existingLayer.locked,
                    visible: existingLayer.visible
                };
                existingLayer.locked = false;
                existingLayer.visible = true;
                return layerState;
            }
        }
        var newLayer = doc.layers.add();
        newLayer.name = name;
        return { layer: newLayer, existed: false, locked: false, visible: true };
    }

    /**
     * RGBColorを作る
     * @param {number} red - 赤（0〜255）
     * @param {number} green - 緑（0〜255）
     * @param {number} blue - 青（0〜255）
     * @returns {RGBColor} 生成した色
     */
    function createRGBColor(red, green, blue) {
        var color = new RGBColor();
        color.red = red;
        color.green = green;
        color.blue = blue;
        return color;
    }

    /**
     * ドキュメントのカラーモードに合わせた無彩色を作る
     * @param {number} blackPercent - 黒の割合（0=白、100=黒）
     * @returns {object} CMYKColor または RGBColor
     */
    function createGrayColor(blackPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = blackPercent;
            return cmykColor;
        }
        var level = Math.round(255 * (100 - blackPercent) / 100);
        return createRGBColor(level, level, level);
    }

    /**
     * ライブエフェクトを適用する（失敗しても続行）
     * @param {PathItem} item - 適用対象のパス
     * @param {string} xml - ライブエフェクトのXML
     * @returns {void}
     */
    function applyLiveEffect(item, xml) {
        try {
            item.applyEffect(xml);
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": ライブエフェクトの適用に失敗 / failed to apply effect — " + e);
        }
    }

    /**
     * ワープ効果を適用する
     * @param {PathItem} item - 適用対象のパス
     * @param {object} settings - ダイアログの設定
     * @param {boolean} isVertical - 垂直方向のワープにするか
     * @returns {void}
     */
    function applyWarpEffect(item, settings, isVertical) {
        applyLiveEffect(item, '<LiveEffect name="Adobe Deform"><Dict data="' +
            'S DisplayString Warp:' + settings.warpName +
            ' I DeformStyle ' + settings.warpStyle +
            ' B Rotate ' + (isVertical ? 1 : 0) +
            ' R DeformValue ' + (settings.warpAmount / 100) +
            ' R DeformHoriz ' + (WARP_DEFORM_H / 100) +
            ' R DeformVert ' + (WARP_DEFORM_V / 100) +
            ' "/></LiveEffect>');
    }

    /**
     * 角丸効果を適用する
     * @param {PathItem} item - 適用対象のパス
     * @param {number} radius - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(item, radius) {
        applyLiveEffect(item, '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>');
    }

    /**
     * パスの長さを返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 長さ（pt）。取得できないときは0
     */
    function getPathLength(pathItem) {
        var pathLength = pathItem.length;
        return (typeof pathLength === "number" && pathLength > 0) ? pathLength : 0;
    }

    /**
     * パスの長さに合わせた破線の線分・間隔を求める（両端を調整）
     * 線分の本数を n とすると n×線分＋(n−1)×間隔＝全長
     * @param {number} pathLength - パスの長さ（pt）
     * @param {number} segments - 分割数（線分の本数）
     * @param {number} gap - 間隔（pt）
     * @returns {Array<number>} strokeDashes に渡す配列
     */
    function calcFittedDashes(pathLength, segments, gap) {
        if (pathLength <= 0) return [];
        if (segments < 2) return []; // 1本＝実線

        var dash = (pathLength - gap * (segments - 1)) / segments;
        if (dash > 0) return [dash, gap];

        // 間隔が大きすぎて線分が残らないときは、線分と間隔を等分する
        var evenLength = pathLength / (segments * 2 - 1);
        return [evenLength, evenLength];
    }

    /**
     * パスの長さに合わせたドットの間隔を求める（両端にドットが乗る）
     * @param {number} pathLength - パスの長さ（pt）
     * @param {number} segments - 分割数（ドットの数）
     * @returns {Array<number>} strokeDashes に渡す配列
     */
    function calcFittedDots(pathLength, segments) {
        if (pathLength <= 0) return [];

        var dotCount = (segments < 2) ? 2 : segments;
        return [0, pathLength / (dotCount - 1)];
    }

    /**
     * 角の形状と破線を適用する（アクションで上書きされた後にも呼ぶ）
     * @param {PathItem} pathItem - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyStrokeStyle(pathItem, settings) {
        pathItem.strokeJoin = STROKE_JOIN_OPTIONS[settings.strokeJoin].value;
        applyDashStyle(pathItem, settings);
    }

    /**
     * 破線の設定を適用する（線分・間隔はパスの長さに合わせて計算する）
     * @param {PathItem} pathItem - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyDashStyle(pathItem, settings) {
        var pathLength = getPathLength(pathItem);
        if (settings.dashStyle === 1) {
            pathItem.strokeDashes = calcFittedDashes(pathLength, settings.dashSegments, settings.dashGap);
            pathItem.strokeCap = StrokeCap.BUTTENDCAP;
        } else if (settings.dashStyle === 2) {
            // 線分0＋丸い先端で点線にする
            pathItem.strokeDashes = calcFittedDots(pathLength, settings.dashSegments);
            pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
        } else {
            pathItem.strokeDashes = [];
            pathItem.strokeCap = StrokeCap.BUTTENDCAP;
        }
    }

    /**
     * コネクターの線を作成し、必要ならライブエフェクトを適用する
     * @param {object} container - 作成先のグループまたはレイヤー
     * @param {Array<Array<number>>} points - 座標の配列
     * @param {object} settings - ダイアログの設定
     * @returns {PathItem} 作成したパス
     */
    function drawConnectorLine(container, points, settings) {
        var connector = container.pathItems.add();
        connector.setEntirePath(points);
        connector.closed = false;
        connector.filled = false;
        connector.stroked = true;
        connector.strokeWidth = settings.strokeWidth;
        connector.strokeColor = createGrayColor(100);
        applyStrokeStyle(connector, settings);

        if (settings.lineShape === 1) {
            var last = points.length - 1;
            var isVertical;
            if (settings.warpAxis === 1) {
                isVertical = false;
            } else if (settings.warpAxis === 2) {
                isVertical = true;
            } else {
                // 線に沿った向きにワープしても曲がらないので、線の傾きから軸を選ぶ
                isVertical = Math.abs(points[last][1] - points[0][1]) > Math.abs(points[last][0] - points[0][0]);
            }
            applyWarpEffect(connector, settings, isVertical);
        }

        if (settings.cornerRadius > 0) {
            applyRoundCornersEffect(connector, settings.cornerRadius);
        }
        return connector;
    }

    /**
     * 白丸用に、パスの端点へ白い●を重ねる
     * @param {object} container - 作成先のグループまたはレイヤー
     * @param {Array<number>} point - [X, Y] 端点の座標（●の中心）
     * @param {object} settings - ダイアログの設定
     * @returns {PathItem} 作成したパス。作成できないときは null
     */
    function drawInnerDot(container, point, settings) {
        var innerDiameter = settings.strokeWidth * WHITE_DOT_RATIO;
        if (innerDiameter <= 0) return null;

        var innerRadius = innerDiameter / 2;
        var innerDot = container.pathItems.ellipse(point[1] + innerRadius, point[0] - innerRadius, innerDiameter, innerDiameter);
        innerDot.stroked = false;
        innerDot.filled = true;
        innerDot.fillColor = createGrayColor(0);
        return innerDot;
    }

    // =========================================
    // 矢印（一時アクション）/ Arrowheads via a temporary action
    // =========================================
    // 矢印はDOMから設定できないため、［線］パネルの設定を行うアクションを生成して実行する。
    // Arrowheads cannot be reached from the DOM, so a temporary action is generated and played.

    /**
     * 矢印番号から、アクションに渡す矢印名を作る
     * @param {number} number - ［線］パネルの矢印番号
     * @returns {string} 矢印名
     */
    function getArrowName(number) {
        return getLabel(LABELS.arrow.prefix) + number;
    }

    /**
     * 文字列をUTF-8バイト列の16進表現にする
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進表現
     */
    function stringToUtf8Hex(sourceText) {
        var hexText = "";
        for (var i = 0; i < sourceText.length; i++) {
            var code = sourceText.charCodeAt(i);
            var bytes;
            if (code < 0x80) {
                bytes = [code];
            } else if (code < 0x800) {
                bytes = [0xC0 | (code >> 6), 0x80 | (code & 0x3F)];
            } else {
                bytes = [0xE0 | (code >> 12), 0x80 | ((code >> 6) & 0x3F), 0x80 | (code & 0x3F)];
            }
            for (var j = 0; j < bytes.length; j++) {
                var byteHex = bytes[j].toString(16);
                if (byteHex.length < 2) byteHex = "0" + byteHex;
                hexText += byteHex;
            }
        }
        return hexText;
    }

    /**
     * 5 → "5.0" のように必ず小数点を含む文字列にする
     * @param {number} value - 変換する数値
     * @returns {string} 小数点を含む文字列
     */
    function toRealString(value) {
        var realText = String(Number(value));
        if (realText.indexOf(".") === -1 && realText.indexOf("e") === -1) realText += ".0";
        return realText;
    }

    /**
     * アクションセット名・アクション名の /name 行を作る
     * @param {string} name - 名前
     * @returns {string} /name 行
     */
    function buildNameLine(name) {
        var hexText = stringToUtf8Hex(name);
        return "/name [ " + (hexText.length / 2) + " \n\t" + hexText + "\n]\n";
    }

    /**
     * パラメータブロックの外枠を作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} body - ブロックの中身
     * @returns {string} パラメータブロック
     */
    function buildParamBlock(index, key, body) {
        return "\t\t/parameter-" + index + " {\n" +
            "\t\t\t/key " + key + "\n" +
            "\t\t\t/showInPalette 4294967295\n" +
            body +
            "\t\t}\n";
    }

    /**
     * 単位付き実数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @param {number} unitCode - 単位コード
     * @returns {string} パラメータブロック
     */
    function buildUnitRealParam(index, key, value, unitCode) {
        return buildParamBlock(index, key,
            "\t\t\t/type (unit real)\n" +
            "\t\t\t/value " + toRealString(value) + "\n" +
            "\t\t\t/unit " + unitCode + "\n");
    }

    /**
     * 列挙パラメータを作る（16進の表示名を直接指定）
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} nameHex - 表示名の16進表現
     * @param {number} byteLength - 表示名のバイト数
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildEnumParam(index, key, nameHex, byteLength, value) {
        return buildParamBlock(index, key,
            "\t\t\t/type (enumerated)\n" +
            "\t\t\t/name [ " + byteLength + " \n\t\t\t\t" + nameHex + "\n\t\t\t]\n" +
            "\t\t\t/value " + value + "\n");
    }

    /**
     * 列挙パラメータを表示名から作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} name - 表示名
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildEnumParamByName(index, key, name, value) {
        var hexText = stringToUtf8Hex(name);
        return buildEnumParam(index, key, hexText, hexText.length / 2, value);
    }

    /**
     * 整数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildIntParam(index, key, value) {
        return buildParamBlock(index, key, "\t\t\t/type (integer)\n\t\t\t/value " + value + "\n");
    }

    /**
     * 真偽値パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {boolean} value - 値
     * @returns {string} パラメータブロック
     */
    function buildBoolParam(index, key, value) {
        return buildParamBlock(index, key, "\t\t\t/type (boolean)\n\t\t\t/value " + (value ? 1 : 0) + "\n");
    }

    /**
     * 実数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildRealParam(index, key, value) {
        return buildParamBlock(index, key, "\t\t\t/type (real)\n\t\t\t/value " + toRealString(value) + "\n");
    }

    /**
     * Unicode文字列パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} value - 値
     * @returns {string} パラメータブロック
     */
    function buildUStrParam(index, key, value) {
        var hexText = stringToUtf8Hex(value);
        return buildParamBlock(index, key,
            "\t\t\t/type (ustring)\n" +
            "\t\t\t/value [ " + (hexText.length / 2) + " \n\t\t\t\t" + hexText + "\n\t\t\t]\n");
    }

    /**
     * ［線］の設定（矢印を含む）を行うアクションのソースを組み立てる
     * @param {object} settings - ダイアログの設定
     * @returns {string} アクションファイルの内容
     */
    function buildStrokeActionSource(settings) {
        return "/version 3\n" +
            buildNameLine(ACTION_SET_NAME) +
            "/isOpen 1\n" +
            "/actionCount 1\n" +
            "/action-1 {\n" +
            "\t" + buildNameLine(ACTION_NAME) +
            "\t/keyIndex 0\n" +
            "\t/colorIndex 0\n" +
            "\t/isOpen 1\n" +
            "\t/eventCount 1\n" +
            "\t/event-1 {\n" +
            "\t\t/useRulersIn1stQuadrant 0\n" +
            "\t\t/internalName (ai_plugin_setStroke)\n" +
            "\t\t/localizedName [ 0 \n\t\t]\n" +
            "\t\t/isOpen 1\n" +
            "\t\t/isOn 1\n" +
            "\t\t/hasDialog 0\n" +
            "\t\t/parameterCount 11\n" +
            /* 可変 / variable */
            buildUnitRealParam(1, KEY_STROKE_WIDTH, settings.strokeWidth, UNIT_POINT) +
            buildUStrParam(6, KEY_ARROW_HEAD_1, settings.startArrow) +
            buildUStrParam(7, KEY_ARROW_HEAD_2, settings.endArrow) +
            buildRealParam(8, KEY_ARROW_SCALE_1, settings.arrowScale) +
            buildRealParam(9, KEY_ARROW_SCALE_2, settings.arrowScale) +
            buildEnumParamByName(10, KEY_ARROW_ALIGN, getLabel(settings.arrowTip.name), settings.arrowTip.value) +
            /* 以下は記録した .aia のまま / recorded as-is */
            buildEnumParam(2, KEY_CAP, "e4b8b8e59e8be7b79ae7abaf", 12, 1) +             /* 線端: 丸型線端 */
            buildEnumParam(3, KEY_JOIN, "e383a9e382a6e383b3e38389e7b590e59088", 18, 1) + /* 角の形状: ラウンド結合 */
            buildIntParam(4, KEY_DASH_INT, 0) +
            buildBoolParam(5, KEY_DASH_BOOL, 0) +
            buildEnumParam(11, KEY_ALIGN, "e4b8ade5a4ae", 6, 0) +                        /* 線の位置: 中央 */
            "\t}\n" +
            "}\n";
    }

    /**
     * アクションを書き出して読み込み、実行後に破棄する
     * @param {string} actionSource - アクションファイルの内容
     * @returns {void}
     */
    function playTemporaryAction(actionSource) {
        var actionFile = new File(ACTION_FILE_NAME);
        var isActionLoaded = false;
        var isActionFileOpen = false;

        try { app.unloadAction(ACTION_SET_NAME, ""); } catch (e) {}

        try {
            actionFile.encoding = "BINARY";
            if (!actionFile.open("w")) throw new Error(getLabel(LABELS.alert.actionFailed));
            isActionFileOpen = true;

            actionFile.write(actionSource);
            actionFile.close();
            isActionFileOpen = false;

            app.loadAction(actionFile);
            isActionLoaded = true;
            app.doScript(ACTION_NAME, ACTION_SET_NAME, false);
        } finally {
            if (isActionFileOpen) {
                try { actionFile.close(); } catch (e) {}
            }
            if (actionFile.exists) {
                try { actionFile.remove(); } catch (e) {}
            }
            if (isActionLoaded) {
                try { app.unloadAction(ACTION_SET_NAME, ""); } catch (e) {}
            }
        }
    }

    /**
     * コネクターの線に矢印を設定する（選択を作ってアクションを1回だけ実行）
     * @param {Array<PathItem>} connectorLines - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyArrowheads(connectorLines, settings) {
        if (!connectorLines.length) return;

        doc.selection = null;
        for (var i = 0; i < connectorLines.length; i++) {
            connectorLines[i].selected = true;
        }
        app.redraw(); // 選択が反映されていないとアクションが空振りする
        try {
            playTemporaryAction(buildStrokeActionSource(settings));
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": 矢印の設定に失敗 / failed to set arrowheads — " + e);
        }
        doc.selection = null;

        // アクションは角の形状・線端・破線も上書きするので、線種を戻す
        for (var i = 0; i < connectorLines.length; i++) {
            applyStrokeStyle(connectorLines[i], settings);
        }
    }

    // =========================================
    // コネクターの作成 / Building connectors
    // =========================================

    toggleSmartGuides();

    var keyIndex = -1;
    try {
        keyIndex = detectKeyObjectIndex(selectedItems);
    } catch (e) {
        // 検出できなくてもスクリプトは続行する
        $.writeln(SCRIPT_NAME + ": キーオブジェクトの検出に失敗 / key object detection failed — " + e);
        keyIndex = -1;
    }

    if (keyIndex === -1) {
        if (selectedItems.length > 2) {
            alert(getLabel(LABELS.alert.noKeyObject));
            toggleSmartGuides();
            return;
        }
        // 2つのときはキーオブジェクトなしでも動くよう、左側を起点にする
        keyIndex = getLeftmostIndex(selectedItems);
    }

    var keyBounds = selectedItems[keyIndex].visibleBounds;
    var connectorPaths = [];
    for (var i = 0; i < selectedItems.length; i++) {
        if (i === keyIndex) continue;
        connectorPaths.push(getConnectionPoints(keyBounds, selectedItems[i].visibleBounds));
    }

    var connectorLayerState = getOrCreateLayer(CONNECTOR_LAYER_NAME);
    var connectorLayer = connectorLayerState.layer;
    var connectors = [];

    /**
     * 作成済みのコネクターを削除する
     * @returns {void}
     */
    function removeConnectors() {
        for (var i = 0; i < connectors.length; i++) {
            try {
                connectors[i].remove();
            } catch (e) {}
        }
        connectors = [];
    }

    /**
     * 1本ぶんの座標列を、形状の指定に合わせて作る
     * @param {object} route - getRoutes() の要素
     * @param {object} settings - ダイアログの設定
     * @param {object} sharedBends - 分岐で共有する折れ位置（分岐以外は null）
     * @returns {Array<Array<number>>} 座標の配列
     */
    function getLinePoints(route, settings, sharedBends) {
        if (settings.lineShape === 2) return getElbowPoints(route.points, route.horizontal);
        if (settings.lineShape === 3) return getElbowPoints(route.points, route.horizontal, sharedBends[route.side]);
        return route.points;
    }

    /**
     * 現在の設定でコネクターを作り直す（プレビュー兼本番）
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function buildConnectors(settings) {
        removeConnectors();
        var routes = getRoutes(settings.startPoint);
        var sharedBends = (settings.lineShape === 3) ? getSharedBendPositions(routes) : null;
        var connectorLines = [];
        var innerDotPoints = [];

        for (var i = 0; i < routes.length; i++) {
            var route = routes[i];
            var linePoints = getLinePoints(route, settings, sharedBends);
            var connectorLine = drawConnectorLine(connectorLayer, linePoints, settings);
            connectors.push(connectorLine);
            connectorLines.push(connectorLine);
            if (settings.arrowInnerDot) {
                innerDotPoints.push(linePoints[linePoints.length - 1]);
                if (settings.arrowPosition === 1) innerDotPoints.push(linePoints[0]);
            }
        }

        if (settings.hasArrow) applyArrowheads(connectorLines, settings);

        // 白丸は矢印の丸の上に重ねるので、線より後に作る
        for (var i = 0; i < innerDotPoints.length; i++) {
            var innerDot = drawInnerDot(connectorLayer, innerDotPoints[i], settings);
            if (innerDot) connectors.push(innerDot);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 文字列を数値に変換する。数値にならないときは既定値を返す
     * @param {string} inputText - 入力文字列
     * @param {number} fallback - 既定値
     * @returns {number} 数値
     */
    function toNumber(inputText, fallback) {
        // Number("") は 0 になるため、空欄は既定値に落とす
        var trimmedText = String(inputText).replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return fallback;
        var parsedValue = Number(trimmedText);
        return isNaN(parsedValue) ? fallback : parsedValue;
    }

    /**
     * 選択されているラジオのインデックスを返す
     * @param {Array<object>} radios - ラジオボタンの配列
     * @returns {number} インデックス
     */
    function getSelectedIndex(radios) {
        for (var i = 0; i < radios.length; i++) {
            if (radios[i].value) return i;
        }
        return 0;
    }

    /**
     * 右揃えラベル付きの1行を追加する
     * @param {object} parent - 追加先のパネル
     * @param {object} labelSet - 行ラベルの定義
     * @returns {object} 追加した行グループ
     */
    function addFieldRow(parent, labelSet) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["fill", "top"];
        row.alignChildren = ["left", "center"];
        var rowLabel = row.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = currentLabelWidth;
        rowLabel.justify = "right";
        row.label = rowLabel; // 縦並びのときに上揃えへ変えられるよう控える
        return row;
    }

    /**
     * ラベル付きのラジオボタン行を追加する
     * @param {object} parent - 追加先のパネル
     * @param {object} labelSet - 行ラベルの定義
     * @param {Array<object>} optionSets - 各ラジオのラベル定義
     * @param {number} defaultIndex - 初期選択のインデックス
     * @param {object} tipSet - helpTip の定義（省略可）
     * @param {boolean} vertical - ラジオを縦並びにするか（省略可）
     * @returns {object} row（行グループ）と radios（ラジオの配列）
     */
    function addRadioRow(parent, labelSet, optionSets, defaultIndex, tipSet, vertical) {
        var row = addFieldRow(parent, labelSet);
        var container = row;
        if (vertical) {
            // ラベルは1つ目のラジオに合わせて上揃えにする
            row.alignChildren = ["left", "top"];
            row.label.alignment = ["left", "top"];
            container = row.add("group");
            container.orientation = "column";
            container.alignChildren = ["left", "center"];
            container.spacing = RADIO_COLUMN_SPACING;
        }
        var radios = [];
        for (var i = 0; i < optionSets.length; i++) {
            var radio = container.add("radiobutton", undefined, getLabel(optionSets[i]));
            if (tipSet) radio.helpTip = getLabel(tipSet);
            radios.push(radio);
        }
        radios[defaultIndex].value = true;
        return { row: row, radios: radios };
    }

    /**
     * ラベル付きの数値入力行を追加する
     * @param {object} parent - 追加先のパネル
     * @param {object} labelSet - 行ラベルの定義
     * @param {number} value - 初期値
     * @param {object} unitSet - 単位表示の定義
     * @param {object} tipSet - helpTip の定義（省略可）
     * @param {Array<number>} range - [最小値, 最大値]。渡すとスライダーを付ける（省略可）
     * @returns {object} row（行グループ）、input（入力欄）、slider（スライダーまたは null）
     */
    function addNumberRow(parent, labelSet, value, unitSet, tipSet, range) {
        var row = addFieldRow(parent, labelSet);
        var input = row.add("edittext", undefined, String(value));
        input.characters = FIELD_CHARS;
        row.add("statictext", undefined, getLabel(unitSet));
        if (tipSet) input.helpTip = getLabel(tipSet);

        var slider = null;
        if (range) {
            slider = row.add("slider", undefined, value, range[0], range[1]);
            // 固定幅で広げるとダイアログが太るので、余りを吸わせる
            slider.minimumSize.width = SLIDER_MIN_WIDTH;
            slider.alignment = ["fill", "center"];
            if (tipSet) slider.helpTip = getLabel(tipSet);
        }
        return { row: row, input: input, slider: slider };
    }

    /**
     * 入力欄の値をスライダーへ反映する（範囲外は丸める）
     * @param {object} field - addNumberRow が返したフィールド
     * @returns {void}
     */
    function syncSliderToInput(field) {
        if (!field.slider) return;
        var value = Number(field.input.text);
        if (isNaN(value)) return;
        if (value < field.slider.minvalue) value = field.slider.minvalue;
        if (value > field.slider.maxvalue) value = field.slider.maxvalue;
        field.slider.value = value;
    }

    /**
     * ↑↓キーで数値を増減する（Shiftで±10、Optionで±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - マイナス値を許可するか
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative) {
        editText.addEventListener("keydown", function (event) {
            if (!(event && (event.keyName === "Up" || event.keyName === "Down"))) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            } else {
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }
            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            // 矢印キーでは onChange が発火しないことがあるため明示的に呼ぶ
            if (typeof editText.onChange === "function") editText.onChange();
        });
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /* ユーザーの設定フォルダーに保存する / Stored in the user settings folder */
    var PRESET_FILE = new File(Folder.userData + "/" + SCRIPT_NAME + "/presets.json");

    /* 直前の設定をIllustratorのセッション中だけ記憶する（終了でリセット） */
    /* Remember the last settings within this Illustrator session (resets on quit) */
    var SESSION_KEY = SCRIPT_NAME + "_lastSettings";
    if (typeof $.global[SESSION_KEY] === "undefined") {
        $.global[SESSION_KEY] = null;
    }

    /**
     * 保存済みのプリセットを読み込む
     * @returns {Array<object>} プリセットの配列。読み込めないときは空配列
     */
    function loadPresets() {
        if (!PRESET_FILE.exists) return [];
        try {
            PRESET_FILE.encoding = "UTF-8";
            if (!PRESET_FILE.open("r")) return [];
            var savedText = PRESET_FILE.read();
            PRESET_FILE.close();
            var parsedPresets = eval(savedText); // toSource() で書き出した内容
            return (parsedPresets && typeof parsedPresets.length === "number") ? parsedPresets : [];
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": プリセットの読み込みに失敗 / failed to load presets — " + e);
            return [];
        }
    }

    /**
     * プリセットを保存する
     * @param {Array<object>} presetList - 保存するプリセットの配列
     * @returns {boolean} 保存できたか
     */
    function writePresets(presetList) {
        try {
            var presetFolder = PRESET_FILE.parent;
            if (!presetFolder.exists) presetFolder.create();
            PRESET_FILE.encoding = "UTF-8";
            if (!PRESET_FILE.open("w")) return false;
            PRESET_FILE.write(presetList.toSource());
            PRESET_FILE.close();
            return true;
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": プリセットの保存に失敗 / failed to save presets — " + e);
            return false;
        }
    }

    /**
     * 名前でプリセットのインデックスを探す
     * @param {Array<object>} presetList - プリセットの配列
     * @param {string} name - プリセット名
     * @returns {number} インデックス。無ければ -1
     */
    function findPresetIndex(presetList, name) {
        for (var i = 0; i < presetList.length; i++) {
            if (presetList[i].name === name) return i;
        }
        return -1;
    }

    var presets = loadPresets();

    var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    setupWindow(dialog);

    /* プリセット / Presets */
    var presetRow = dialog.add("group");
    presetRow.orientation = "row";
    presetRow.alignment = ["fill", "top"];
    presetRow.alignChildren = ["left", "center"];
    var presetLabel = presetRow.add("statictext", undefined, labelText(LABELS.fieldLabel.preset));
    presetLabel.preferredSize.width = LABEL_WIDTH;
    presetLabel.justify = "right";
    var presetDropdown = presetRow.add("dropdownlist", undefined, []);
    presetDropdown.alignment = ["fill", "center"];
    presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);
    var btnSavePreset = presetRow.add("button", undefined, getLabel(LABELS.button.save));
    btnSavePreset.preferredSize.width = PRESET_BUTTON_WIDTH;
    btnSavePreset.alignment = ["right", "center"];
    var btnRemovePreset = presetRow.add("button", undefined, getLabel(LABELS.button.remove));
    btnRemovePreset.preferredSize.width = PRESET_BUTTON_WIDTH;
    btnRemovePreset.alignment = ["right", "center"];

    /* コネクター / Connector */
    var connectorPanel = addPanel(dialog, getLabel(LABELS.panel.connector));
    var lineShapeField = addRadioRow(connectorPanel, LABELS.fieldLabel.lineShape, [LABELS.radio.shapeNone, LABELS.radio.shapeWarp, LABELS.radio.shapeElbow, LABELS.radio.shapeBranch], DEFAULT_LINE_SHAPE, LABELS.tooltip.lineShape);

    var warpTypeRow = addFieldRow(connectorPanel, LABELS.fieldLabel.warpType);
    var warpTypeList = warpTypeRow.add("dropdownlist", undefined, getWarpTypeLabels());
    warpTypeList.preferredSize.width = LIST_WIDTH;
    warpTypeList.selection = DEFAULT_WARP_TYPE;

    var warpAmountField = addNumberRow(connectorPanel, LABELS.fieldLabel.warpAmount, DEFAULT_WARP_AMOUNT, LABELS.unit.percent, LABELS.tooltip.warpAmount, [WARP_AMOUNT_MIN, WARP_AMOUNT_MAX]);
    var warpAxisField = addRadioRow(connectorPanel, LABELS.fieldLabel.warpAxis, [LABELS.radio.axisAuto, LABELS.radio.axisHorizontal, LABELS.radio.axisVertical], DEFAULT_WARP_AXIS, LABELS.tooltip.warpAxis);
    var cornerField = addNumberRow(connectorPanel, LABELS.fieldLabel.cornerRadius, DEFAULT_CORNER_RADIUS, LABELS.unit.pt, LABELS.tooltip.cornerRadius);
    var startPointField = addRadioRow(connectorPanel, LABELS.fieldLabel.startPoint, [LABELS.radio.startCenter, LABELS.radio.startDivided], DEFAULT_START_POINT, LABELS.tooltip.startPoint);

    /* 線・矢印は2カラム。ラベル幅はコネクターパネルより狭くする */
    setLabelWidth(COLUMN_LABEL_WIDTH);

    var panelColumnsGroup = dialog.add("group");
    panelColumnsGroup.orientation = "row";
    panelColumnsGroup.alignChildren = ["fill", "fill"];
    panelColumnsGroup.alignment = ["fill", "top"];
    panelColumnsGroup.spacing = COLUMN_SPACING;

    /* 線 / Line */
    var linePanel = addPanel(panelColumnsGroup, getLabel(LABELS.panel.line));
    var strokeWidthField = addNumberRow(linePanel, LABELS.fieldLabel.strokeWidth, DEFAULT_STROKE_WIDTH, LABELS.unit.pt);
    var strokeJoinField = addRadioRow(linePanel, LABELS.fieldLabel.strokeJoin, [STROKE_JOIN_OPTIONS[0].label, STROKE_JOIN_OPTIONS[1].label, STROKE_JOIN_OPTIONS[2].label], DEFAULT_STROKE_JOIN, LABELS.tooltip.strokeJoin, true);
    var dashStyleField = addRadioRow(linePanel, LABELS.fieldLabel.dashStyle, [LABELS.radio.dashNone, LABELS.radio.dashDashed, LABELS.radio.dashDotted], DEFAULT_DASH_STYLE, null, true);
    var dashSegmentsField = addNumberRow(linePanel, LABELS.fieldLabel.dashSegments, DEFAULT_DASH_SEGMENTS, LABELS.unit.blank, LABELS.tooltip.dashSegments);
    var dashGapField = addNumberRow(linePanel, LABELS.fieldLabel.dashGap, DEFAULT_DASH_GAP, LABELS.unit.pt, LABELS.tooltip.dashGap);

    /* 矢印 / Arrowheads */
    var arrowPanel = addPanel(panelColumnsGroup, getLabel(LABELS.panel.arrow));
    var arrowShapeLabels = [];
    for (var i = 0; i < ARROW_CHOICES.length; i++) {
        arrowShapeLabels.push(ARROW_CHOICES[i].label);
    }
    var arrowShapeField = addRadioRow(arrowPanel, LABELS.fieldLabel.arrowShape, arrowShapeLabels, DEFAULT_ARROW_INDEX, LABELS.tooltip.arrowShape, true);
    // 初期値も選択中の矢印の既定倍率にそろえる
    var arrowScaleField = addNumberRow(arrowPanel, LABELS.fieldLabel.arrowScale, ARROW_CHOICES[DEFAULT_ARROW_INDEX].scale, LABELS.unit.percent, LABELS.tooltip.arrowScale);
    var arrowPositionField = addRadioRow(arrowPanel, LABELS.fieldLabel.arrowPosition, [LABELS.radio.arrowEnd, LABELS.radio.arrowBoth], DEFAULT_ARROW_POSITION, LABELS.tooltip.arrowPosition);
    var arrowTipField = addRadioRow(arrowPanel, LABELS.fieldLabel.arrowTip, [ARROW_TIP_OPTIONS[0].label, ARROW_TIP_OPTIONS[1].label], ARROW_CHOICES[DEFAULT_ARROW_INDEX].tip, LABELS.tooltip.arrowTip, true);


    /**
     * ダイアログの入力値を設定として取り出す
     * @returns {object} コネクターの設定
     */
    function getSettings() {
        var warpType = WARP_TYPE_CHOICES[warpTypeList.selection ? warpTypeList.selection.index : DEFAULT_WARP_TYPE];
        var strokeWidth = toNumber(strokeWidthField.input.text, DEFAULT_STROKE_WIDTH);
        if (strokeWidth <= 0) strokeWidth = DEFAULT_STROKE_WIDTH;

        // 矢印はパスの始点＝キーオブジェクト側、終点＝相手の図形側
        var noneName = getLabel(LABELS.arrow.none);
        var arrowChoice = ARROW_CHOICES[getSelectedIndex(arrowShapeField.radios)];
        var arrowName = (arrowChoice.number === 0) ? noneName : getArrowName(arrowChoice.number);
        var arrowPosition = getSelectedIndex(arrowPositionField.radios);
        var lineShape = getSelectedIndex(lineShapeField.radios);

        return {
            strokeWidth: strokeWidth,
            strokeJoin: getSelectedIndex(strokeJoinField.radios),
            dashStyle: getSelectedIndex(dashStyleField.radios),
            dashSegments: Math.round(toNumber(dashSegmentsField.input.text, DEFAULT_DASH_SEGMENTS)),
            dashGap: toNumber(dashGapField.input.text, DEFAULT_DASH_GAP),
            startPoint: getSelectedIndex(startPointField.radios),
            lineShape: lineShape,
            warpName: warpType.name,
            warpStyle: warpType.style,
            warpAmount: toNumber(warpAmountField.input.text, DEFAULT_WARP_AMOUNT),
            warpAxis: getSelectedIndex(warpAxisField.radios),
            // 角丸はカギのときだけ。分岐は角を丸めない / elbow only; Branch keeps square corners
            cornerRadius: (lineShape === 2) ? toNumber(cornerField.input.text, DEFAULT_CORNER_RADIUS) : 0,
            hasArrow: (arrowChoice.number !== 0),
            arrowInnerDot: (arrowChoice.innerDot === true),
            startArrow: (arrowPosition === 1) ? arrowName : noneName,
            endArrow: arrowName,
            arrowScale: toNumber(arrowScaleField.input.text, DEFAULT_ARROW_SCALE),
            arrowTip: ARROW_TIP_OPTIONS[getSelectedIndex(arrowTipField.radios)],
            arrowPosition: arrowPosition
        };
    }

    /* プリセット適用中は「（カスタム）」へ戻さない */
    var isApplyingPreset = false;

    /**
     * 現在の設定でプレビューを更新する
     * @returns {void}
     */
    function updatePreview() {
        // 手で変えたらプリセットの選択を外す
        if (!isApplyingPreset && presetDropdown.selection && presetDropdown.selection.index !== 0) {
            presetDropdown.selection = 0;
        }
        var settings = getSettings();
        warpTypeRow.enabled = (settings.lineShape === 1);
        warpAmountField.row.enabled = (settings.lineShape === 1);
        warpAxisField.row.enabled = (settings.lineShape === 1);
        cornerField.row.enabled = (settings.lineShape === 2);
        dashSegmentsField.row.enabled = (settings.dashStyle !== 0);
        dashGapField.row.enabled = (settings.dashStyle === 1);
        arrowScaleField.row.enabled = settings.hasArrow;
        arrowPositionField.row.enabled = settings.hasArrow;
        arrowTipField.row.enabled = settings.hasArrow;
        buildConnectors(settings);
        app.redraw();

        // 閉じた後はコントロールを読めないので、更新のたびに控えておく
        $.global[SESSION_KEY] = getPresetFromDialog("");
    }

    /**
     * ラジオボタンの選択を切り替える
     * @param {Array<object>} radios - ラジオボタンの配列
     * @param {number} index - 選択するインデックス
     * @returns {void}
     */
    function selectRadio(radios, index) {
        if (!(index >= 0) || index >= radios.length) return;
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (i === index);
        }
    }

    /**
     * プリセットに保存する項目の対応表を返す
     * field=数値入力の行、radios=ラジオの配列、list=ドロップダウン
     * @returns {Array<object>} 保存項目の配列
     */
    function getPresetEntries() {
        return [
            { key: "strokeWidth",   field: strokeWidthField },
            { key: "strokeJoin",    radios: strokeJoinField.radios },
            { key: "dashStyle",     radios: dashStyleField.radios },
            { key: "dashSegments",  field: dashSegmentsField },
            { key: "dashGap",       field: dashGapField },
            { key: "lineShape",     radios: lineShapeField.radios },
            { key: "warpType",      list: warpTypeList },
            { key: "warpAmount",    field: warpAmountField },
            { key: "warpAxis",      radios: warpAxisField.radios },
            { key: "cornerRadius",  field: cornerField },
            { key: "startPoint",    radios: startPointField.radios },
            { key: "arrowIndex",    radios: arrowShapeField.radios },
            { key: "arrowScale",    field: arrowScaleField },
            { key: "arrowPosition", radios: arrowPositionField.radios },
            { key: "arrowTip",      radios: arrowTipField.radios }
        ];
    }

    /**
     * 現在のダイアログの状態をプリセットとして取り出す
     * @param {string} name - プリセット名
     * @returns {object} プリセット
     */
    function getPresetFromDialog(name) {
        var entries = getPresetEntries();
        var preset = { name: name };
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            if (entry.field) {
                preset[entry.key] = entry.field.input.text;
            } else if (entry.radios) {
                preset[entry.key] = getSelectedIndex(entry.radios);
            } else {
                preset[entry.key] = entry.list.selection ? entry.list.selection.index : 0;
            }
        }
        return preset;
    }

    /**
     * プリセットをダイアログへ反映する
     * @param {object} preset - 反映するプリセット
     * @returns {void}
     */
    function applyPresetToDialog(preset) {
        if (!preset) return;

        var entries = getPresetEntries();
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            var value = preset[entry.key];
            if (value === undefined || value === null) continue;

            if (entry.field) {
                entry.field.input.text = value;
                syncSliderToInput(entry.field);
            } else if (entry.radios) {
                selectRadio(entry.radios, value);
            } else if (value >= 0 && value < entry.list.items.length) {
                entry.list.selection = value;
            }
        }
    }

    /**
     * プリセットのドロップダウンを作り直す
     * @param {string} selectName - 選択状態にするプリセット名（省略時は（カスタム））
     * @returns {void}
     */
    function refreshPresetDropdown(selectName) {
        isApplyingPreset = true;
        presetDropdown.removeAll();
        presetDropdown.add("item", getLabel(LABELS.preset.custom));
        for (var i = 0; i < presets.length; i++) {
            presetDropdown.add("item", presets[i].name);
        }
        var index = selectName ? (findPresetIndex(presets, selectName) + 1) : 0;
        presetDropdown.selection = (index > 0) ? index : 0;
        isApplyingPreset = false;
    }

    /**
     * ラジオボタンの配列にプレビュー更新を割り当てる
     * @param {Array<object>} radios - ラジオボタンの配列
     * @returns {void}
     */
    function bindPreviewToRadios(radios) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].onClick = updatePreview;
        }
    }

    /**
     * 数値フィールドにプレビュー更新・↑↓キー操作・スライダー連動を割り当てる
     * @param {object} field - addNumberRow が返したフィールド
     * @param {boolean} allowNegative - マイナス値を許可するか
     * @returns {void}
     */
    function bindPreviewToField(field, allowNegative) {
        field.input.onChange = function () {
            syncSliderToInput(field);
            updatePreview();
        };
        changeValueByArrowKey(field.input, allowNegative);

        if (!field.slider) return;
        // ドラッグ中は数値の表示だけ更新し、離したところで作り直す
        field.slider.onChanging = function () {
            field.input.text = Math.round(field.slider.value);
        };
        field.slider.onChange = function () {
            field.input.text = Math.round(field.slider.value);
            updatePreview();
        };
    }

    bindPreviewToRadios(lineShapeField.radios);
    bindPreviewToRadios(startPointField.radios);
    bindPreviewToRadios(strokeJoinField.radios);
    bindPreviewToRadios(dashStyleField.radios);
    bindPreviewToRadios(warpAxisField.radios);
    bindPreviewToRadios(arrowPositionField.radios);
    bindPreviewToRadios(arrowTipField.radios);
    bindPreviewToField(strokeWidthField, false);
    bindPreviewToField(dashSegmentsField, false);
    bindPreviewToField(dashGapField, false);
    bindPreviewToField(warpAmountField, true);
    bindPreviewToField(cornerField, false);
    bindPreviewToField(arrowScaleField, false);
    warpTypeList.onChange = updatePreview;
    for (var i = 0; i < arrowShapeField.radios.length; i++) {
        arrowShapeField.radios[i].onClick = function () {
            // 矢印ごとに見え方が違うので、選び直したらその矢印の既定値（倍率・先端位置）を入れる
            var selectedArrow = ARROW_CHOICES[getSelectedIndex(arrowShapeField.radios)];
            if (selectedArrow.number !== 0) {
                arrowScaleField.input.text = selectedArrow.scale;
                selectRadio(arrowTipField.radios, selectedArrow.tip);
            }
            updatePreview();
        };
    }

    /* ボタンエリア / Button row */
    var btnRowGroup = dialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.alignment = ["fill", "bottom"];

    /* スペーサー（伸縮）/ Spacer (stretchable) */
    var buttonSpacer = btnRowGroup.add("group");
    buttonSpacer.alignment = ["fill", "fill"];
    buttonSpacer.minimumSize.width = 0;

    /* 右側グループ / Right-side button group */
    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];

    // 確定／破棄の判定は show() の戻り値に一本化する
    // （ESCやウィンドウを閉じたときは onClick が発火しないため）
    var DIALOG_RESULT_OK = 1;
    var DIALOG_RESULT_CANCEL = 2;

    var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
    btnCancel.onClick = function () {
        dialog.close(DIALOG_RESULT_CANCEL);
    };

    var btnOK = btnRightGroup.add("button", undefined, "OK", { name: "ok" });
    btnOK.onClick = function () {
        dialog.close(DIALOG_RESULT_OK);
    };

    presetDropdown.onChange = function () {
        if (isApplyingPreset) return;
        if (!presetDropdown.selection || presetDropdown.selection.index === 0) return;

        isApplyingPreset = true;
        applyPresetToDialog(presets[presetDropdown.selection.index - 1]);
        isApplyingPreset = false;
        updatePreview();
    };

    btnSavePreset.onClick = function () {
        var selectedName = (presetDropdown.selection && presetDropdown.selection.index > 0) ? presetDropdown.selection.text : "";
        var inputName = prompt(getLabel(LABELS.alert.presetName), selectedName);
        if (inputName === null) return;
        var presetName = String(inputName).replace(/^\s+|\s+$/g, "");
        if (presetName === "") return;

        var preset = getPresetFromDialog(presetName);
        var existingIndex = findPresetIndex(presets, presetName);
        if (existingIndex === -1) {
            presets.push(preset);
        } else {
            presets[existingIndex] = preset; // 同名は上書き / overwrite when the name exists
        }
        if (!writePresets(presets)) {
            alert(getLabel(LABELS.alert.presetFailed));
            return;
        }
        refreshPresetDropdown(presetName);
    };

    btnRemovePreset.onClick = function () {
        if (!presetDropdown.selection || presetDropdown.selection.index === 0) return;
        if (!confirm(getLabel(LABELS.alert.presetRemove))) return;

        presets.splice(presetDropdown.selection.index - 1, 1);
        if (!writePresets(presets)) {
            alert(getLabel(LABELS.alert.presetFailed));
            return;
        }
        refreshPresetDropdown(null);
    };

    // 前回このセッションで閉じたときの設定に戻す
    if ($.global[SESSION_KEY]) {
        isApplyingPreset = true;
        applyPresetToDialog($.global[SESSION_KEY]);
        isApplyingPreset = false;
    }
    refreshPresetDropdown(null);
    updatePreview();

    if (dialog.show() !== DIALOG_RESULT_OK) {
        // キャンセル／ESC／ウィンドウを閉じる: プレビューを破棄
        removeConnectors();
        if (!connectorLayerState.existed) {
            try {
                connectorLayer.remove();
            } catch (e) {}
        } else {
            // 作成のために開いたロック・表示状態を元に戻す
            connectorLayer.locked = connectorLayerState.locked;
            connectorLayer.visible = connectorLayerState.visible;
        }
        app.redraw();
        toggleSmartGuides();
        return;
    }

    // コネクターを選択状態にする
    doc.selection = null;
    for (var i = 0; i < connectors.length; i++) {
        connectors[i].selected = true;
    }

    toggleSmartGuides();
})();
