#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

キーオブジェクトを起点に、選択した各図形へコネクターを引きます。
直線・ワープ・カギ・分岐・カーブの5種類の経路に、線・線端・矢印をプレビューしながら設定できます。

詳細は README を参照してください。

### Overview

Draws a connector from the key object to each of the selected objects.
Choose a straight, warped, elbow, branch, or curved route and set the stroke, caps, and arrowheads with a live preview.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiConnectorBuilder";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-09";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiConnectorBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiConnectorBuilder.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考、謝辞 / Reference and acknowledgements
 * カーブの作図（中点から振った1点を2次ベジェの制御点として扱う考え方）
 * Egor Chistyakov (@tchegr)
 * https://x.com/tchegr
 */

(function () {
    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var CONNECTOR_LAYER_NAME = { ja: "コネクター", en: "Connector" };  /* コネクターの作成先レイヤー名 / target layer name */

    /* 線の初期値 / Line defaults */
    var DEFAULT_STROKE_WIDTH  = 1;    /* 線幅（pt） / stroke width */
    var DEFAULT_STROKE_JOIN   = 0;    /* 0=マイター 1=ラウンド 2=ベベル */
    var DEFAULT_STROKE_CAP    = 0;    /* 0=なし 1=丸形 2=突出 */
    var DEFAULT_DASH_STYLE    = 0;    /* 0=なし 1=破線 2=ドット */
    var DEFAULT_DASH_SEGMENTS = 8;    /* 分割数（線分・ドットの数） */
    var DEFAULT_DASH_GAP      = 3;    /* 破線の間隔（pt） */

    /* コネクターの初期値 / Connector defaults */
    var DEFAULT_LINE_SHAPE    = 0;    /* 0=直線 1=ワープ 2=カギ 3=分岐 4=カーブ */
    var DEFAULT_START_POINT   = 0;    /* 0=各辺の中心 1=等分 2=中心 */
    var DEFAULT_UNIFY_ANCHOR  = 4;    /* 開始点をまとめるときの位置（0〜8。4=中央＝自動） */
    var DEFAULT_WARP_TYPE     = 0;    /* WARP_TYPE_CHOICES のインデックス（0=でこぼこ） */
    var DEFAULT_WARP_AMOUNT   = -80;  /* カーブ（%）：ワープの曲がり具合とカーブのふくらみに共用 */
    var WARP_AMOUNT_MIN       = -100; /* カーブの下限（%） */
    var WARP_AMOUNT_MAX       = 100;  /* カーブの上限（%） */
    var DEFAULT_WARP_AXIS     = 0;    /* 0=自動 1=水平 2=垂直 */
    var DEFAULT_CORNER_RADIUS = 0;    /* カギ・分岐の角丸半径（pt）／0で角丸なし */
    var WARP_DEFORM_H         = 0;    /* 変形・水平方向（%）：XMLで必須のため固定値で渡す */
    var WARP_DEFORM_V         = 0;    /* 変形・垂直方向（%）：同上 */

    /* 矢印の初期値 / Arrowhead defaults */
    var DEFAULT_ARROW_INDEX  = 1;    /* ARROW_CHOICES のインデックス（0=なし） */
    var DEFAULT_ARROW_SCALE  = 100;  /* 矢印の倍率（%）：ARROW_CHOICES に指定がないときの値 */
    var DEFAULT_END_GAP      = 0;    /* 終点と相手の図形とのすき間（pt） */
    var DEFAULT_ARROW_POSITION = 0;  /* 0=終点のみ 1=両端 */
    /* 黒丸に使う矢印番号（［線］パネルの矢印リストに合わせて調整） */
    var ARROW_DOT_FILLED = 21;       /* 黒丸 */
    /* 白丸は黒丸の中心に白い●を重ねて作る */
    var WHITE_DOT_RATIO = 1.5;       /* 白い●の直径＝線幅の何倍か（WHITE_DOT_BASE_SCALE のときの値） */
    var WHITE_DOT_BASE_SCALE = 50;   /* 白い●の基準になる矢印の倍率（%） */

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
    var BEND_CLEARANCE_PT       = 4;     /* 折れ位置を選択オブジェクトから離す量（pt） */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 */
    var LABEL_WIDTH        = 84;             /* コネクターパネルの行ラベル幅（右揃え） */
    var COLUMN_LABEL_WIDTH = 70;             /* 線パネルの行ラベル幅（2カラムなので狭め） */
    var ARROW_LABEL_WIDTH  = 30;             /* 矢印パネルの行ラベル幅（2文字ぶん） */
    var FIELD_CHARS    = 4;                  /* 数値欄の文字数 */
    var LIST_WIDTH     = 150;                /* ドロップダウンの幅 */
    var SLIDER_MIN_WIDTH = 60;               /* スライダーの最小幅（余白は fill で伸ばす） */
    var RADIO_COLUMN_SPACING = 4;            /* 縦並びラジオの間隔 */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 */
    var PRESET_BUTTON_WIDTH = 60;            /* プリセットの保存・削除ボタンの幅 */
    var BUTTON_ROW_TOP_MARGIN = 5;           /* ボタン行の上余白 */
    var ANCHOR_WIDGET_SIZE = 66;             /* 起点ウィジェット全体の大きさ */
    var ANCHOR_CELL_SIZE   = 9;              /* 起点ウィジェットの□1個の大きさ */
    var ANCHOR_CELL_GAP    = 7.5;            /* 起点ウィジェットの□どうしの間隔 */
    var ANCHOR_DEFAULT_INDEX = 4;            /* 起点ウィジェットの初期位置（4=中央） */
    var KEY_DIALOG_TEXT_WIDTH = 240;         /* 起点ダイアログの説明文の幅 */
    var PRESET_LIST_WIDTH = 120;             /* プリセットのドロップダウンの幅 */
    var ICON_BUTTON_SIZE = 24;               /* 矢印アイコン1個の大きさ */
    var ICON_BUTTON_SPACING = 2;             /* 矢印アイコンどうしの間隔 */
    var ICON_PADDING = 4;                    /* 矢印アイコンの内側の余白 */
    var ICON_ROW_BOTTOM_MARGIN = 5;          /* 矢印アイコン行の下余白 */

    /* 矢印アイコンの配色。UIの明暗に合わせて initIconColors() で入れ替える */
    var ICON_COLOR          = [0.25, 0.25, 0.25, 1];
    var ICON_SELECTED_COLOR = [1, 1, 1, 1];
    var ICON_BG             = [1, 1, 1, 1];
    var ICON_SELECTED_BG    = [0.4, 0.4, 0.4, 1];
    var ICON_BORDER_COLOR   = [0.65, 0.65, 0.65, 1];

    // 確定／破棄の判定は show() の戻り値に一本化する
    // （ESCやウィンドウを閉じたときは onClick が発火しないため）
    var DIALOG_RESULT_OK = 1;
    var DIALOG_RESULT_CANCEL = 2;

    /* 起点ウィジェットのケイ線・枠線（常時この色）／選択セルの塗りは明暗で入れ替える */
    var ANCHOR_LINE_COLOR    = [0.6, 0.6, 0.6, 1];
    var ANCHOR_SELECTED_FILL = [0.4, 0.4, 0.4, 1];
    var ANCHOR_DISABLED_LINE = [0.75, 0.75, 0.75, 1];
    var ANCHOR_DISABLED_FILL = [0.75, 0.75, 0.75, 1];
    /* 外周の□どうしをつなぐケイ線（中央は独立）*/
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

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
            title:     { ja: "コネクター", en: "Connector" },
            keyObject: { ja: "起点にするオブジェクト", en: "Start object" }
        },
        message: {
            noKeyObject: { ja: "キーオブジェクトが設定されていません。起点にするオブジェクトの位置を選んでください。", en: "No key object is set. Pick where the object you want to start from sits." }
        },
        panel: {
            connector: { ja: "コネクター", en: "Connector" },
            line:   { ja: "線", en: "Line" },
            arrow:  { ja: "線端と矢印", en: "Ends & arrowheads" }
        },
        fieldLabel: {
            preset:         { ja: "プリセット", en: "Preset" },
            strokeWidth:    { ja: "線幅", en: "Stroke width" },
            strokeJoin:     { ja: "角の形状", en: "Corner" },
            strokeCap:      { ja: "線端", en: "Cap" },
            dashStyle:      { ja: "破線", en: "Dashes" },
            dashSegments:   { ja: "分割数", en: "Divisions" },
            dashGap:        { ja: "間隔", en: "Gap" },
            lineShape:      { ja: "形状", en: "Shape" },
            startPoint:     { ja: "開始ポイント", en: "Start point" },
            warpType:       { ja: "種類", en: "Style" },
            warpAmount:     { ja: "カーブ", en: "Bend" },
            warpAxis:       { ja: "方向", en: "Axis" },
            cornerRadius:   { ja: "角丸", en: "Round corners" },
            arrowScale:     { ja: "倍率", en: "Scale" },
            endGap:         { ja: "余白", en: "Gap" },
            arrowPosition:  { ja: "先端", en: "Ends" },
            arrowTip:       { ja: "位置", en: "Position" }
        },
        radio: {
            joinMiter:      { ja: "マイター", en: "Miter" },
            joinRound:      { ja: "ラウンド", en: "Round" },
            joinBevel:      { ja: "ベベル", en: "Bevel" },
            capButt:        { ja: "なし", en: "Butt" },
            capRound:       { ja: "丸形", en: "Round" },
            capProjecting:  { ja: "突出", en: "Projecting" },
            dashNone:       { ja: "なし", en: "None" },
            dashDashed:     { ja: "破線", en: "Dashed" },
            dashDotted:     { ja: "ドット", en: "Dotted" },
            startCenter:    { ja: "各辺の中心", en: "Edge centers" },
            startDivided:   { ja: "等分", en: "Divided" },
            startKeyCenter: { ja: "中心", en: "Center" },
            shapeStraight:  { ja: "直線", en: "Straight" },
            shapeWarp:      { ja: "ワープ", en: "Warp" },
            shapeElbow:     { ja: "カギ", en: "Elbow" },
            shapeBranch:    { ja: "分岐", en: "Branch" },
            shapeCurve:     { ja: "カーブ", en: "Curve" },
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
        check: {
            manualKey:  { ja: "ダイアログを閉じて手動で選ぶ", en: "Close and pick it manually" },
            unifyStart: { ja: "開始点をまとめる", en: "Share one start point" }
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
            strokeCap:      { ja: "線の端の見え方です（［線］パネルの線端）。ドットは丸形にしないと点が出ません。", en: "How the line ends look (the Stroke panel's cap). Dots need the round cap to show up." },
            dashSegments:   { ja: "線分の数です。両端が線分で終わるように線分の長さを計算します。", en: "Number of dashes; the dash length is solved so both ends finish with a dash." },
            dashGap:        { ja: "線分（ドット）どうしのすき間です。ドットでは両端がドットで乗るよう、指定に近い間隔へそろえます。", en: "Gap between dashes (dots). For dots it is nudged to the nearest value that lands a dot on both ends." },
            keyObject:      { ja: "キーオブジェクトが未設定のため、起点にするオブジェクトを位置で選びます。選択範囲の左上〜中央〜右下のうち、押した位置にいちばん近いオブジェクトが起点になります。", en: "No key object was set, so pick the start object by position: the object nearest the corner or center you click becomes the start." },
            manualKey:      { ja: "ダイアログを閉じます。選択したうえで基準にしたいオブジェクトをもう一度クリックしてキーオブジェクトにしてから、スクリプトを実行し直してください。", en: "Closes the dialog. With the objects selected, click the one you want as the key object, then run the script again." },
            unifyStart:     { ja: "すべてのコネクターをキーオブジェクトの同じ位置から出します。位置は右の9分割で選びます。", en: "Runs every connector out of the same point on the key object; pick the point with the 3x3 grid on the right." },
            unifyAnchor:    { ja: "開始点をまとめるときの位置です。中央は本数のいちばん多い辺を自動で選びます。上下の中央は上辺・下辺、左右の中央は左辺・右辺、四隅はその高さの左辺・右辺から出します。", en: "Where the shared start point sits. Center picks the edge used by the most connectors; the top and bottom cells use those edges, the left and right cells use theirs, and the corners leave from the left or right edge at that height." },
            startPoint:     { ja: "等分は、同じ辺から出るコネクターの本数＋1でその辺を等分し、起点をずらします。中心は、キーオブジェクトの中心から相手へ向かう向きで、起点を辺の上に置きます。", en: "Divided spreads the start points along the key object's edge, splitting it into (connectors + 1) parts; Center aims each connector from the key object's center and starts it where that line meets the edge." },
            lineShape:      { ja: "直線はまっすぐ結び、ワープは直線にワープ効果、カギは直角に折れる線、分岐は折れ位置をそろえて幹を共有し、カーブは弧を描いて結びます。", en: "Straight connects directly, Warp adds a warp effect to the straight line, Elbow is a right-angled route, Branch shares a trunk with aligned bends, and Curve bows the line into an arc." },
            warpAmount:     { ja: "ワープの曲がり具合と、カーブのふくらみ（線の長さに対する割合）です。マイナス値で向きが逆になります。", en: "How much Warp bends, and how far Curve bows out relative to the line length. A negative value flips the direction." },
            warpAxis:       { ja: "自動は全コネクターをまとめて水平／垂直を選びます（線ごとに変えるとアピアランスが混在するため）。線に沿った向きのワープは曲がりません。", en: "Auto picks one axis for all the connectors together (a per-line axis would mix their appearances); warping along the line has no visible effect." },
            cornerRadius:   { ja: "カギ・分岐の角を丸めます（0で角丸なし）。", en: "Round the corners of Elbow and Branch routes (0 = square corners)." },
            preset:         { ja: "現在の設定に名前を付けて保存できます。保存先はユーザーの設定フォルダーです。", en: "Save the current settings under a name; presets are stored in your user settings folder." },
            arrowShape:     { ja: "［線］パネルの矢印を使います。黒丸も矢印の一種です。", en: "Uses the Stroke panel arrowheads; the dot is an arrowhead preset too." },
            arrowScale:     { ja: "矢印の大きさ（%）。線幅に対する比率です。", en: "Arrowhead size in percent, relative to the stroke width." },
            endGap:         { ja: "終点と相手の図形とのすき間です。最後の線分より大きい値は無視します。キーオブジェクト側は詰めません。", en: "Space left between the end of the connector and the object. Values longer than the final segment are ignored; the key-object end is not inset." },
            arrowTip:       { ja: "矢印の先端をパスの終点に配置するか、パスの終点から配置するかを選びます。", en: "Place the arrow tip at the end of the path, or extend it beyond the end." },
            arrowPosition:  { ja: "終点はキーオブジェクトと反対側、両端は起点にも付けます。", en: "End = the far side from the key object; Both ends also marks the start." }
        },
        alert: {
            presetName:    { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
            presetRemove:  { ja: "このプリセットを削除しますか？", en: "Delete this preset?" },
            presetFailed:  { ja: "プリセットを保存できませんでした。", en: "Could not save the presets." },
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectObjects: { ja: "2つ以上の図形を選択してください。", en: "Please select two or more objects." },
            actionFailed:  { ja: "一時アクションファイルを開けませんでした。", en: "Failed to open the temporary action file." }
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
        { label: LABELS.radio.arrow8,    number: 8,                scale: 25,  tip: 0, cap: 0 },
        { label: LABELS.radio.arrow11,   number: 11,               scale: 100, tip: 0, cap: 0 },
        { label: LABELS.radio.dotFilled, number: ARROW_DOT_FILLED, scale: 50,  tip: 1, cap: 1 },
        { label: LABELS.radio.dotHollow, number: ARROW_DOT_FILLED, scale: 50,  tip: 1, cap: 1, innerDot: true }
    ];

    /* 角の形状 / Stroke join */
    var STROKE_JOIN_OPTIONS = [
        { label: LABELS.radio.joinMiter, value: StrokeJoin.MITERENDJOIN },
        { label: LABELS.radio.joinRound, value: StrokeJoin.ROUNDENDJOIN },
        { label: LABELS.radio.joinBevel, value: StrokeJoin.BEVELENDJOIN }
    ];

    var STROKE_CAP_OPTIONS = [
        { label: LABELS.radio.capButt,       value: StrokeCap.BUTTENDCAP },
        { label: LABELS.radio.capRound,      value: StrokeCap.ROUNDENDCAP },
        { label: LABELS.radio.capProjecting, value: StrokeCap.PROJECTINGENDCAP }
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
     * 判定中は app.redraw() を呼ばない。描画するとスクリプトの操作がそこで確定し、
     * 整列がそのつど取り消し履歴に積まれてしまう（描画しなければ位置は同期的に読める）。
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
                // 2回目以降だけ元の位置へ戻す（1回目はまだ動かしていない）
                if (c > 0) restorePositions(items, originPositions);
                app.executeMenuCommand(alignCommands[c]);
                for (i = 0; i < items.length; i++) {
                    if (!stayedPut[i]) continue;
                    if (Math.abs(items[i].left - originPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                        Math.abs(items[i].top - originPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                        stayedPut[i] = false;
                    }
                }
            }
        } finally {
            // 判定で動かしたぶんを最後に一度だけ戻す（例外で抜けるときも整列結果を残さない）
            restorePositions(items, originPositions);
        }

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

    /**
     * 選択したオブジェクト全体を囲む外接矩形を返す
     * @param {Array<object>} items - 対象のオブジェクト
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getSelectionBounds(items) {
        var union = [items[0].visibleBounds[0], items[0].visibleBounds[1], items[0].visibleBounds[2], items[0].visibleBounds[3]];
        for (var i = 1; i < items.length; i++) {
            var bounds = items[i].visibleBounds;
            if (bounds[0] < union[0]) union[0] = bounds[0];
            if (bounds[1] > union[1]) union[1] = bounds[1];
            if (bounds[2] > union[2]) union[2] = bounds[2];
            if (bounds[3] < union[3]) union[3] = bounds[3];
        }
        return union;
    }

    /**
     * 3×3のどの位置かを指定して、そこにいちばん近いオブジェクトのインデックスを返す
     * @param {Array<object>} items - 対象のオブジェクト
     * @param {number} anchorIndex - 0〜8（左上から右下へ）
     * @returns {number} インデックス
     */
    function getIndexAtAnchor(items, anchorIndex) {
        var union = getSelectionBounds(items);
        var targetX = union[0] + (union[2] - union[0]) * ((anchorIndex % 3) / 2);
        var targetY = union[1] - (union[1] - union[3]) * (Math.floor(anchorIndex / 3) / 2);

        var nearestIndex = 0;
        var minDistance = null;
        for (var i = 0; i < items.length; i++) {
            var bounds = items[i].visibleBounds;
            var dx = (bounds[0] + bounds[2]) / 2 - targetX;
            var dy = (bounds[1] + bounds[3]) / 2 - targetY;
            var distance = dx * dx + dy * dy;
            if (minDistance === null || distance < minDistance) {
                minDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
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
        var side;
        if (Math.abs(dx) >= Math.abs(dy)) {
            side = (dx >= 0) ? "right" : "left";
        } else {
            side = (dy >= 0) ? "top" : "bottom";
        }
        return getConnectionPointsOnSide(fromBounds, toBounds, side);
    }

    /**
     * 辺を指定して、始点・終点を求める
     * @param {Array<number>} fromBounds - 始点側の visibleBounds
     * @param {Array<number>} toBounds - 終点側の visibleBounds
     * @param {string} side - 使う辺（"right" / "left" / "top" / "bottom"）
     * @returns {object} points（始点・終点）と horizontal（左右の辺どうしか）
     */
    function getConnectionPointsOnSide(fromBounds, toBounds, side) {
        var fromCenterX = (fromBounds[0] + fromBounds[2]) / 2;
        var fromCenterY = (fromBounds[1] + fromBounds[3]) / 2;
        var toCenterX = (toBounds[0] + toBounds[2]) / 2;
        var toCenterY = (toBounds[1] + toBounds[3]) / 2;

        var route;
        if (side === "right") {
            route = { points: [[fromBounds[2], fromCenterY], [toBounds[0], toCenterY]], horizontal: true, side: side };
        } else if (side === "left") {
            route = { points: [[fromBounds[0], fromCenterY], [toBounds[2], toCenterY]], horizontal: true, side: side };
        } else if (side === "top") {
            route = { points: [[fromCenterX, fromBounds[1]], [toCenterX, toBounds[3]]], horizontal: false, side: side };
        } else {
            route = { points: [[fromCenterX, fromBounds[3]], [toCenterX, toBounds[1]]], horizontal: false, side: side };
        }
        // 等分配置で辺に沿って並べ替えるための基準
        route.order = route.horizontal ? toCenterY : toCenterX;
        return route;
    }

    /**
     * いちばん本数の多い辺を返す
     * @param {Array<object>} paths - getConnectionPoints() の結果の配列
     * @returns {string} 辺の名前
     */
    function getMajoritySide(paths) {
        var counts = {};
        var majoritySide = paths[0].side;
        for (var i = 0; i < paths.length; i++) {
            var side = paths[i].side;
            counts[side] = (counts[side] || 0) + 1;
            if (counts[side] > counts[majoritySide]) majoritySide = side;
        }
        return majoritySide;
    }

    /**
     * 外接矩形の中心から指定した点へ向かう線が、矩形の辺と交わる位置を求める
     * @param {Array<number>} bounds - visibleBounds
     * @param {Array<number>} toPoint - 向かう先の座標 [X, Y]
     * @returns {Array<number>} 辺の上の座標 [X, Y]
     */
    function getEdgePoint(bounds, toPoint) {
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        var halfWidth = (bounds[2] - bounds[0]) / 2;
        var halfHeight = (bounds[1] - bounds[3]) / 2;
        var dx = toPoint[0] - centerX;
        var dy = toPoint[1] - centerY;
        if (dx === 0 && dy === 0) return [centerX, centerY];
        // 左右・上下それぞれの辺に届くまでの比率のうち、小さいほうが先に交わる辺
        var ratioX = (dx === 0) ? null : halfWidth / Math.abs(dx);
        var ratioY = (dy === 0) ? null : halfHeight / Math.abs(dy);
        var ratio = (ratioX === null) ? ratioY : ((ratioY === null) ? ratioX : Math.min(ratioX, ratioY));
        return [centerX + dx * ratio, centerY + dy * ratio];
    }

    /**
     * 9分割のどの位置かから、使う辺を求める
     * @param {number} anchorIndex - 0〜8（左上から右下へ）。4=中央は自動
     * @returns {string} 辺の名前。中央のときは null
     */
    function getSideForAnchor(anchorIndex) {
        if (anchorIndex === 1) return "top";
        if (anchorIndex === 7) return "bottom";
        if (anchorIndex === 3 || anchorIndex === 0 || anchorIndex === 6) return "left";
        if (anchorIndex === 5 || anchorIndex === 2 || anchorIndex === 8) return "right";
        return null; // 中央は本数のいちばん多い辺にまかせる
    }

    /**
     * すべてのコネクターを、同じ位置から出し直した経路を返す
     * @param {number} anchorIndex - 開始点の位置（0〜8）
     * @returns {Array<object>} 経路の配列
     */
    function getUnifiedPaths(anchorIndex) {
        var side = getSideForAnchor(anchorIndex);
        if (!side) side = getMajoritySide(connectorPaths);

        // 辺に沿った向きの位置は、9分割の行（左右の辺）または列（上下の辺）で決める
        var horizontal = (side === "left" || side === "right");
        var ratio = (horizontal ? Math.floor(anchorIndex / 3) : (anchorIndex % 3)) / 2;
        var startPerp = horizontal
            ? keyBounds[1] - (keyBounds[1] - keyBounds[3]) * ratio
            : keyBounds[0] + (keyBounds[2] - keyBounds[0]) * ratio;

        var paths = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (i === keyIndex) continue;
            var path = getConnectionPointsOnSide(keyBounds, selectedItems[i].visibleBounds, side);
            path.points[0][horizontal ? 1 : 0] = startPerp;
            paths.push(path);
        }
        return paths;
    }

    /**
     * 開始ポイントの指定を反映した経路の複製を返す
     * @param {number} startPoint - 0=各辺の中心 1=等分 2=中心
     * @param {boolean} unifyStart - すべてを同じ位置から出すか
     * @param {number} unifyAnchor - まとめるときの位置（0〜8）
     * @returns {Array<object>} points（座標）と horizontal（左右接続か）の配列
     */
    function getRoutes(startPoint, unifyStart, unifyAnchor) {
        var routes = [];
        var i;
        var paths = unifyStart ? getUnifiedPaths(unifyAnchor) : connectorPaths;
        for (i = 0; i < paths.length; i++) {
            var path = paths[i];
            routes.push({
                points: [[path.points[0][0], path.points[0][1]], [path.points[1][0], path.points[1][1]]],
                horizontal: path.horizontal,
                side: path.side,
                order: path.order
            });
        }
        if (startPoint === 2) {
            // キーオブジェクトの中心から相手へ向かう線を、キーオブジェクトの辺で止める
            for (i = 0; i < routes.length; i++) {
                var start = getEdgePoint(keyBounds, routes[i].points[1]);
                routes[i].points[0][0] = start[0];
                routes[i].points[0][1] = start[1];
            }
            return routes;
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
     * 折れ位置が選択オブジェクトにかからないよう、いちばん近い空き位置へずらす
     * @param {number} bendPosition - もとの折れ位置
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか（true なら折れ位置はX）
     * @param {number} lower - 折れ位置に使える下限
     * @param {number} upper - 折れ位置に使える上限
     * @param {number} perpLower - 折れ線が伸びる向きの下限
     * @param {number} perpUpper - 折れ線が伸びる向きの上限
     * @returns {number} ずらした折れ位置。よけられないときはもとの位置
     */
    function getClearBendPosition(bendPosition, horizontal, lower, upper, perpLower, perpUpper) {
        var clearPosition = bendPosition;
        // 重なった図形の外へ寄せる。図形が並んでいることもあるので選択数だけ繰り返す
        for (var pass = 0; pass < selectedItems.length; pass++) {
            var moved = false;
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = selectedItems[i].visibleBounds;
                var minAxis = horizontal ? bounds[0] : bounds[3];
                var maxAxis = horizontal ? bounds[2] : bounds[1];
                var minPerp = horizontal ? bounds[3] : bounds[0];
                var maxPerp = horizontal ? bounds[1] : bounds[2];

                var before = minAxis - BEND_CLEARANCE_PT;
                var after = maxAxis + BEND_CLEARANCE_PT;
                if (clearPosition <= before || clearPosition >= after) continue;
                if (maxPerp < perpLower || minPerp > perpUpper) continue; // 折れ線が届かない位置

                clearPosition = (clearPosition - before <= after - clearPosition) ? before : after;
                moved = true;
                break;
            }
            if (!moved) break;
        }
        // すき間の外まで押し出されるならよけられない
        if (clearPosition < lower || clearPosition > upper) return bendPosition;
        return clearPosition;
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
            var bendX = hasBend ? bendPosition : getClearBendPosition((start[0] + end[0]) / 2, true,
                Math.min(start[0], end[0]), Math.max(start[0], end[0]),
                Math.min(start[1], end[1]), Math.max(start[1], end[1]));
            return [start, [bendX, start[1]], [bendX, end[1]], end];
        }
        if (Math.abs(end[0] - start[0]) < COORD_TOLERANCE_PT) return [start, end];
        var bendY = hasBend ? bendPosition : getClearBendPosition((start[1] + end[1]) / 2, false,
            Math.min(start[1], end[1]), Math.max(start[1], end[1]),
            Math.min(start[0], end[0]), Math.max(start[0], end[0]));
        return [start, [start[0], bendY], [end[0], bendY], end];
    }

    /**
     * 分岐用に、同じ辺から出る経路で共有する折れ位置を求める
     * いちばん近い図形とのすき間の中央にそろえ、幹を1本にまとめる
     * @param {Array<object>} routes - getRoutes() の戻り値
     * @returns {object} 辺をキーにした折れ位置
     */
    function getSharedBendPositions(routes) {
        var nearestRoutes = {};
        var perpRanges = {};
        var i;
        for (i = 0; i < routes.length; i++) {
            var route = routes[i];
            var axis = route.horizontal ? 0 : 1;
            var perpAxis = route.horizontal ? 1 : 0;
            var routeSide = route.side;
            var delta = route.points[1][axis] - route.points[0][axis];
            if (!nearestRoutes[routeSide] || Math.abs(delta) < Math.abs(nearestRoutes[routeSide].delta)) {
                nearestRoutes[routeSide] = { route: route, delta: delta };
            }
            // 折れ線（背骨）が伸びる範囲。よける図形を絞り込むのに使う
            var perpStart = route.points[0][perpAxis];
            var perpEnd = route.points[1][perpAxis];
            if (!perpRanges[routeSide]) {
                perpRanges[routeSide] = [Math.min(perpStart, perpEnd), Math.max(perpStart, perpEnd)];
            } else {
                perpRanges[routeSide][0] = Math.min(perpRanges[routeSide][0], perpStart, perpEnd);
                perpRanges[routeSide][1] = Math.max(perpRanges[routeSide][1], perpStart, perpEnd);
            }
        }

        var bendPositions = {};
        for (var side in nearestRoutes) {
            if (!nearestRoutes.hasOwnProperty(side)) continue;
            var nearest = nearestRoutes[side];
            var bendAxis = nearest.route.horizontal ? 0 : 1;
            var startCoord = nearest.route.points[0][bendAxis];
            var endCoord = nearest.route.points[1][bendAxis];
            bendPositions[side] = getClearBendPosition(startCoord + nearest.delta / 2, nearest.route.horizontal,
                Math.min(startCoord, endCoord), Math.max(startCoord, endCoord),
                perpRanges[side][0], perpRanges[side][1]);
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
     * 2点のパスを弧にする（中点を線と垂直な向きへふくらませる）
     * 中点から振った1点を2次ベジェの制御点とみなし、両端のハンドルへ置き換える
     * @param {PathItem} pathItem - 対象のパス（アンカーが2点のもの）
     * @param {number} amountPercent - カーブ（%）。マイナスで反対側へふくらむ
     * @returns {void}
     */
    function bendIntoCurve(pathItem, amountPercent) {
        if (!amountPercent || pathItem.pathPoints.length !== 2) return;

        var startPoint = pathItem.pathPoints[0];
        var endPoint = pathItem.pathPoints[1];
        var from = startPoint.anchor;
        var to = endPoint.anchor;
        var dx = to[0] - from[0];
        var dy = to[1] - from[1];
        var length = Math.sqrt(dx * dx + dy * dy);
        if (length <= 0) return;

        // ふくらみは線の長さに比例させ、進行方向の左向き（法線）へ振る
        var sag = length / 2 * (amountPercent / 100);
        var control = [
            (from[0] + to[0]) / 2 - dy / length * sag,
            (from[1] + to[1]) / 2 + dx / length * sag
        ];
        // 2次ベジェの制御点を3次ベジェのハンドルに置き換える比率
        var handleRatio = 2 / 3;
        startPoint.rightDirection = [
            from[0] + (control[0] - from[0]) * handleRatio,
            from[1] + (control[1] - from[1]) * handleRatio
        ];
        endPoint.leftDirection = [
            to[0] + (control[0] - to[0]) * handleRatio,
            to[1] + (control[1] - to[1]) * handleRatio
        ];
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
     * 指定の間隔に近く、両端にドットが乗る間隔を求める
     * @param {number} pathLength - パスの長さ（pt）
     * @param {number} gap - ドットどうしの間隔（pt）
     * @returns {Array<number>} strokeDashes に渡す配列
     */
    function calcFittedDots(pathLength, gap) {
        if (pathLength <= 0) return [];

        var dotGaps = (gap > 0) ? Math.round(pathLength / gap) : 1;
        if (dotGaps < 1) dotGaps = 1;
        return [0, pathLength / dotGaps];
    }

    /**
     * 角の形状と破線を適用する（アクションで上書きされた後にも呼ぶ）
     * @param {PathItem} pathItem - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyStrokeStyle(pathItem, settings) {
        pathItem.strokeJoin = STROKE_JOIN_OPTIONS[settings.strokeJoin].value;
        pathItem.strokeCap = STROKE_CAP_OPTIONS[settings.strokeCap].value;
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
        } else if (settings.dashStyle === 2) {
            // 線分0＋丸い線端で点線にする
            pathItem.strokeDashes = calcFittedDots(pathLength, settings.dashGap);
        } else {
            pathItem.strokeDashes = [];
        }
    }

    /**
     * コネクターの線を作成し、必要ならライブエフェクトを適用する
     * @param {object} container - 作成先のグループまたはレイヤー
     * @param {Array<Array<number>>} points - 座標の配列
     * @param {object} settings - ダイアログの設定
     * @param {boolean} isVerticalWarp - ワープを垂直方向にするか
     * @returns {PathItem} 作成したパス
     */
    function drawConnectorLine(container, points, settings, isVerticalWarp) {
        var connector = container.pathItems.add();
        connector.setEntirePath(points);
        connector.closed = false;
        connector.filled = false;
        connector.stroked = true;
        connector.strokeWidth = settings.strokeWidth;
        connector.strokeColor = createGrayColor(100);
        // 破線はパスの長さに合わせるので、曲げてから線の設定を入れる
        if (settings.lineShape === 4) {
            bendIntoCurve(connector, settings.warpAmount);
        }
        applyStrokeStyle(connector, settings);

        if (settings.lineShape === 1) {
            applyWarpEffect(connector, settings, isVerticalWarp);
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
        // 黒丸の矢印は倍率で大きくなるので、白い●も同じ比率で合わせる
        var innerDiameter = settings.strokeWidth * WHITE_DOT_RATIO * (settings.arrowScale / WHITE_DOT_BASE_SCALE);
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
            // キーオブジェクトが無いときは、起点を別ダイアログで選んでもらう
            keyIndex = chooseKeyIndex(selectedItems);
            if (keyIndex === -1) {
                toggleSmartGuides();
                return;
            }
        } else {
            keyIndex = getLeftmostIndex(selectedItems); // 2つのときは左側を起点にする
        }
    }

    var keyBounds;
    var connectorPaths;

    /**
     * 起点にするオブジェクトを決め、経路を作り直す
     * @param {number} index - selectedItems のインデックス
     * @returns {void}
     */
    function setKeyIndex(index) {
        keyIndex = index;
        keyBounds = selectedItems[keyIndex].visibleBounds;
        connectorPaths = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (i === keyIndex) continue;
            connectorPaths.push(getConnectionPoints(keyBounds, selectedItems[i].visibleBounds));
        }
    }

    setKeyIndex(keyIndex);

    var connectorLayerState = getOrCreateLayer(getLabel(CONNECTOR_LAYER_NAME));
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
     * 1本ぶんの座標列を、形状の指定に合わせて作る（分岐は getBranchLines() で作る）
     * @param {object} route - getRoutes() の要素
     * @param {object} settings - ダイアログの設定
     * @returns {Array<Array<number>>} 座標の配列
     */
    function getLinePoints(route, settings) {
        if (settings.lineShape === 2) return getElbowPoints(route.points, route.horizontal);
        return route.points;
    }

    /**
     * 幹の片側にある枝を、幹に近い順に並べて返す
     * @param {Array<object>} group - 同じ辺から出る経路
     * @param {number} trunkPosition - 幹の位置
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか
     * @param {number} direction - 1=幹より先、-1=幹より手前（同じ位置は手前に含める）
     * @returns {Array<object>} 並べ替えた経路
     */
    function getBranchOrder(group, trunkPosition, horizontal, direction) {
        var axis = horizontal ? 1 : 0;
        var picked = [];
        for (var i = 0; i < group.length; i++) {
            var delta = group[i].points[1][axis] - trunkPosition;
            if ((direction > 0) ? (delta > 0) : (delta <= 0)) picked.push(group[i]);
        }
        picked.sort(function (a, b) {
            return Math.abs(a.points[1][axis] - trunkPosition) - Math.abs(b.points[1][axis] - trunkPosition);
        });
        return picked;
    }

    /**
     * 幹の片側の枝の座標列を作る。幹からいちばん遠い枝だけが幹から折れる長いパスになり、
     * 手前の枝はその背骨から横に出すだけにして重なりをなくす
     * @param {Array<object>} ordered - 幹に近い順に並べた経路
     * @param {number} trunkPosition - 幹の位置
     * @param {number} bendPosition - 折れ位置
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか
     * @returns {Array<Array<Array<number>>>} 枝の座標列
     */
    function getSideBranches(ordered, trunkPosition, bendPosition, horizontal) {
        var branches = [];
        var farthest = ordered.length - 1;
        for (var i = 0; i < ordered.length; i++) {
            var end = ordered[i].points[1];
            var position = horizontal ? end[1] : end[0];
            var isFarthest = (i === farthest);
            var startPosition = isFarthest ? trunkPosition : position;
            var start = horizontal ? [bendPosition, startPosition] : [startPosition, bendPosition];
            if (isFarthest && Math.abs(position - trunkPosition) >= COORD_TOLERANCE_PT) {
                branches.push([start, horizontal ? [bendPosition, position] : [position, bendPosition], end]);
            } else {
                branches.push([start, end]);
            }
        }
        return branches;
    }

    /**
     * 分岐の経路を、幹（キーから折れ位置まで）と枝（折れ位置から相手まで）に分ける
     * @param {Array<object>} routes - getRoutes() の結果
     * @param {object} sharedBends - 辺ごとの折れ位置
     * @returns {object} trunks（幹の座標列）と branches（枝の座標列）
     */
    function getBranchLines(routes, sharedBends) {
        var groups = {};
        var i;
        for (i = 0; i < routes.length; i++) {
            if (!groups[routes[i].side]) groups[routes[i].side] = [];
            groups[routes[i].side].push(routes[i]);
        }

        var trunks = [];
        var branches = [];
        for (var side in groups) {
            if (!groups.hasOwnProperty(side)) continue;
            var group = groups[side];
            var horizontal = group[0].horizontal;
            var bendPosition = sharedBends[side];

            // 幹は各起点の平均の位置（左右の辺なら高さ、上下の辺なら左右）に置く
            var total = 0;
            for (i = 0; i < group.length; i++) {
                total += group[i].points[0][horizontal ? 1 : 0];
            }
            var trunkPosition = total / group.length;
            var edgePosition = group[0].points[0][horizontal ? 0 : 1];

            // 矢印がキーオブジェクト側の端に付くよう、幹は折れ位置からキーへ向けて引く
            trunks.push(horizontal
                ? [[bendPosition, trunkPosition], [edgePosition, trunkPosition]]
                : [[trunkPosition, bendPosition], [trunkPosition, edgePosition]]);

            // 幹の前後それぞれで、いちばん遠い枝が背骨を兼ねる（手前の枝は横だけなので重ならない）
            branches = branches.concat(getSideBranches(getBranchOrder(group, trunkPosition, horizontal, 1), trunkPosition, bendPosition, horizontal));
            branches = branches.concat(getSideBranches(getBranchOrder(group, trunkPosition, horizontal, -1), trunkPosition, bendPosition, horizontal));
        }
        return { trunks: trunks, branches: branches };
    }

    /**
     * 終点を進行方向の手前へ戻す（相手の図形とのすき間を作る）
     * @param {Array<Array<number>>} points - 座標の配列
     * @param {number} gap - 空けるすき間（pt）
     * @returns {Array<Array<number>>} 終点をずらした座標の配列
     */
    function applyEndGap(points, gap) {
        if (!(gap > 0) || points.length < 2) return points;

        var last = points.length - 1;
        var from = points[last - 1];
        var to = points[last];
        var dx = to[0] - from[0];
        var dy = to[1] - from[1];
        var length = Math.sqrt(dx * dx + dy * dy);
        if (length <= gap) return points; // 最後の線分より大きいすき間は詰めない

        var shortened = [];
        for (var i = 0; i < last; i++) {
            shortened.push([points[i][0], points[i][1]]);
        }
        shortened.push([to[0] - dx / length * gap, to[1] - dy / length * gap]);
        return shortened;
    }

    /**
     * 始点側の矢印を外した設定の複製を返す（分岐で合流点に矢印を出さないため）
     * @param {object} settings - ダイアログの設定
     * @returns {object} 複製した設定
     */
    function getEndArrowOnlySettings(settings) {
        var copy = {};
        for (var key in settings) {
            if (settings.hasOwnProperty(key)) copy[key] = settings[key];
        }
        copy.startArrow = getLabel(LABELS.arrow.none);
        return copy;
    }

    /**
     * ワープの軸を決める。自動のときは全線をまとめて1つに決める
     * @param {Array<Array<Array<number>>>} allLinePoints - 全コネクターの座標列
     * @param {number} warpAxis - 0=自動 1=水平 2=垂直
     * @returns {boolean} 垂直方向のワープにするか
     */
    function getWarpAxis(allLinePoints, warpAxis) {
        if (warpAxis === 1) return false;
        if (warpAxis === 2) return true;

        // 線に沿った向きのワープは曲がらないので、いちばん曲がりにくい線でも
        // 幅（高さ）を確保できるほうの軸を選ぶ
        var minWidth = null;
        var minHeight = null;
        for (var i = 0; i < allLinePoints.length; i++) {
            var points = allLinePoints[i];
            var last = points.length - 1;
            var width = Math.abs(points[last][0] - points[0][0]);
            var height = Math.abs(points[last][1] - points[0][1]);
            if (minWidth === null || width < minWidth) minWidth = width;
            if (minHeight === null || height < minHeight) minHeight = height;
        }
        if (minWidth === null) return false;
        return minHeight > minWidth;
    }

    /**
     * 線と白丸をグループにまとめ、作成済みリストの項目をそのグループに差し替える
     * @param {PathItem} line - コネクターの線
     * @param {Array<Array<number>>} dotPoints - 白丸を置く座標
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function groupWithInnerDots(line, dotPoints, settings) {
        var group = connectorLayer.groupItems.add();
        line.move(group, ElementPlacement.PLACEATEND);
        // グループへ追加した順に前面へ入るので、白丸は線より後に作る
        for (var i = 0; i < dotPoints.length; i++) {
            drawInnerDot(group, dotPoints[i], settings);
        }
        for (var k = 0; k < connectors.length; k++) {
            if (connectors[k] === line) {
                connectors[k] = group;
                return;
            }
        }
        connectors.push(group);
    }

    /**
     * 現在の設定でコネクターを作り直す（プレビュー兼本番）
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function buildConnectors(settings) {
        removeConnectors();
        var routes = getRoutes(settings.startPoint, settings.unifyStart, settings.unifyAnchor);
        var sharedBends = (settings.lineShape === 3) ? getSharedBendPositions(routes) : null;
        var connectorLines = [];
        var lineJobs = [];

        // 分岐は幹（キーから折れ位置まで）と枝（折れ位置から相手まで）に分けて作る
        var isBranch = (settings.lineShape === 3);
        var trunkLinePoints = [];
        var allLinePoints = [];
        var i;
        if (isBranch) {
            var branchLines = getBranchLines(routes, sharedBends);
            trunkLinePoints = branchLines.trunks;
            allLinePoints = branchLines.branches;
        } else {
            for (i = 0; i < routes.length; i++) {
                allLinePoints.push(getLinePoints(routes[i], settings));
            }
        }

        // 幹はキーオブジェクト側で終わるので、すき間は相手側で終わる線だけに入れる
        for (i = 0; i < allLinePoints.length; i++) {
            allLinePoints[i] = applyEndGap(allLinePoints[i], settings.endGap);
        }

        // ワープの軸は全線で同じにする（線ごとに変えるとアピアランスが混在する）
        var isVerticalWarp = getWarpAxis(allLinePoints, settings.warpAxis);

        var trunkLines = [];
        for (i = 0; i < trunkLinePoints.length; i++) {
            var trunkLine = drawConnectorLine(connectorLayer, trunkLinePoints[i], settings, isVerticalWarp);
            connectors.push(trunkLine);
            trunkLines.push(trunkLine);
            // 両端のときは幹のキー側にも印を付ける
            if (settings.arrowInnerDot && settings.arrowPosition === 1) {
                lineJobs.push({ line: trunkLine, dotPoints: [trunkLinePoints[i][trunkLinePoints[i].length - 1]] });
            }
        }

        for (i = 0; i < allLinePoints.length; i++) {
            var linePoints = allLinePoints[i];
            var connectorLine = drawConnectorLine(connectorLayer, linePoints, settings, isVerticalWarp);
            connectors.push(connectorLine);
            connectorLines.push(connectorLine);
            if (settings.arrowInnerDot) {
                var dotPoints = [linePoints[linePoints.length - 1]];
                if (!isBranch && settings.arrowPosition === 1) dotPoints.push(linePoints[0]);
                lineJobs.push({ line: connectorLine, dotPoints: dotPoints });
            }
        }

        if (settings.hasArrow) {
            if (isBranch) {
                // 枝は終点だけ。両端のときは幹のキー側の端が始点側の矢印になる
                var arrowLines = (settings.arrowPosition === 1) ? connectorLines.concat(trunkLines) : connectorLines;
                applyArrowheads(arrowLines, getEndArrowOnlySettings(settings));
            } else {
                applyArrowheads(connectorLines, settings);
            }
        }

        // 白丸は矢印の丸の上に重ねるので、線より後に作り、その線とグループ化する
        for (i = 0; i < lineJobs.length; i++) {
            groupWithInnerDots(lineJobs[i].line, lineJobs[i].dotPoints, settings);
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
     * UIが明るいテーマかを判定する
     * @returns {boolean} 明るいテーマなら true（取得できないときは暗い側）
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * 正方形のパスを作る
     * @param {object} graphics - ScriptUIGraphics
     * @param {number} x - 左端
     * @param {number} y - 上端
     * @param {number} size - 一辺の長さ
     * @returns {void}
     */
    function squarePath(graphics, x, y, size) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
    }

    /**
     * 起点ウィジェットの□を1つ描く
     * @param {object} graphics - ScriptUIGraphics
     * @param {number} x - 左端
     * @param {number} y - 上端
     * @param {boolean} selected - 選択中か
     * @param {boolean} enabled - ウィジェットが有効か
     * @returns {void}
     */
    function drawAnchorCell(graphics, x, y, selected, enabled) {
        // 枠を上に描くので塗りを先に行う
        if (selected) {
            squarePath(graphics, x, y, ANCHOR_CELL_SIZE);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, enabled ? ANCHOR_SELECTED_FILL : ANCHOR_DISABLED_FILL));
        }
        squarePath(graphics, x, y, ANCHOR_CELL_SIZE);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, enabled ? ANCHOR_LINE_COLOR : ANCHOR_DISABLED_LINE, 1));
    }

    /**
     * 起点ウィジェットを描く（外周の□をケイ線でつなぐ・中央は独立）
     * @param {object} widget - 描画対象のボタン
     * @returns {void}
     */
    function drawAnchorWidget(widget) {
        var graphics = widget.graphics;

        // 背景はコントロールの地色で塗って、パネルに溶け込ませる
        try {
            graphics.rectPath(0, 0, widget.size[0], widget.size[1]);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((widget.size[0] - gridSize) / 2);
        var originY = Math.round((widget.size[1] - gridSize) / 2);

        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        // パネルのenabledは描画に伝わらないので、ウィジェット自身のenabledを見る
        var enabled = (widget.enabled !== false);
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, enabled ? ANCHOR_LINE_COLOR : ANCHOR_DISABLED_LINE, 1);
        for (i = 0; i < ANCHOR_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_CONNECTIONS[i][1] - ANCHOR_CONNECTIONS[i][0] === 1) {
                // 横方向：右隣の□へ
                graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE, cellA[1] + ANCHOR_CELL_SIZE / 2);
                graphics.lineTo(cellB[0], cellB[1] + ANCHOR_CELL_SIZE / 2);
            } else {
                // 縦方向：下の□へ
                graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE / 2, cellA[1] + ANCHOR_CELL_SIZE);
                graphics.lineTo(cellB[0] + ANCHOR_CELL_SIZE / 2, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < cellPositions.length; i++) {
            drawAnchorCell(graphics, cellPositions[i][0], cellPositions[i][1], i === widget.selectedAnchorIndex, enabled);
        }
    }

    /**
     * 起点を位置で選ぶ3×3ウィジェットを追加する
     * @param {object} parent - 追加先
     * @param {number} anchorIndex - 最初に選んでおくセル（0〜8）
     * @param {function} onSelect - セルを選んだときに呼ぶ処理。引数はセルのインデックス
     * @returns {object} 追加したウィジェット
     */
    function addAnchorWidget(parent, anchorIndex, onSelect) {
        var widget = parent.add("button", undefined, "");
        widget.minimumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        widget.preferredSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        widget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        widget.selectedAnchorIndex = anchorIndex;
        widget.onDraw = function () {
            drawAnchorWidget(this);
        };
        // クリック位置の判定は mousedown で行う（座標はコントロール基準）
        try {
            widget.addEventListener("mousedown", function (event) {
                if (widget.enabled === false) return;
                var col = Math.floor(event.clientX / (widget.size[0] / 3));
                var row = Math.floor(event.clientY / (widget.size[1] / 3));
                if (col < 0) col = 0;
                if (col > 2) col = 2;
                if (row < 0) row = 0;
                if (row > 2) row = 2;
                widget.selectedAnchorIndex = row * 3 + col;
                try { widget.notify("onDraw"); } catch (e) {}
                onSelect(widget.selectedAnchorIndex);
            });
        } catch (e) {}
        return widget;
    }

    /**
     * UIの明暗に合わせて矢印アイコンの配色を決める
     * @returns {void}
     */
    function initIconColors() {
        var lightUI = isLightUI();
        ICON_COLOR          = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
        ICON_SELECTED_COLOR = lightUI ? [1, 1, 1, 1]          : [0.15, 0.15, 0.15, 1];
        ICON_BG             = lightUI ? [1, 1, 1, 1]          : [0.22, 0.22, 0.22, 1];
        ICON_SELECTED_BG    = lightUI ? [0.4, 0.4, 0.4, 1]    : [0.8, 0.8, 0.8, 1];
        ICON_BORDER_COLOR   = lightUI ? [0.65, 0.65, 0.65, 1] : [0.45, 0.45, 0.45, 1];
    }

    /**
     * 矢印の形状アイコンを描く（線・実線矢印・線矢印・黒丸・白丸）
     * @param {object} button - 描画対象のボタン
     * @returns {void}
     */
    function drawArrowShapeIcon(button) {
        var graphics = button.graphics;
        var width = button.size[0];
        var height = button.size[1];
        var selected = button.isSelected;

        var backColor = selected ? ICON_SELECTED_BG : ICON_BG;
        graphics.rectPath(0, 0, width, height);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backColor));
        graphics.rectPath(0, 0, width, height);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));

        var color = selected ? ICON_SELECTED_COLOR : ICON_COLOR;
        var thickPen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, 3);
        var thinPen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1);
        var iconBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, color);
        var backBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, backColor);

        var left = ICON_PADDING;
        var right = width - ICON_PADDING;
        var centerY = Math.round(height / 2);
        var headSize = 5;
        var dotRadius = 4;

        /**
         * 軸線を引く
         * @param {number} endX - 線の右端
         * @param {object} pen - 使うペン
         * @returns {void}
         */
        function drawShaft(endX, pen) {
            graphics.newPath();
            graphics.moveTo(left, centerY);
            graphics.lineTo(endX, centerY);
            graphics.strokePath(pen);
        }

        if (button.arrowIndex === 0) {
            drawShaft(right, thickPen);
            return;
        }
        if (button.arrowIndex === 1) {
            // 実線の矢印：軸は太く、先端は塗りの三角
            drawShaft(right - headSize - 1, thickPen);
            graphics.newPath();
            graphics.moveTo(right - headSize - 1, centerY - headSize);
            graphics.lineTo(right, centerY);
            graphics.lineTo(right - headSize - 1, centerY + headSize);
            graphics.closePath();
            graphics.fillPath(iconBrush);
            return;
        }
        if (button.arrowIndex === 2) {
            // 線の矢印：軸も先端も細い線
            drawShaft(right, thinPen);
            graphics.newPath();
            graphics.moveTo(right - headSize, centerY - headSize);
            graphics.lineTo(right, centerY);
            graphics.lineTo(right - headSize, centerY + headSize);
            graphics.strokePath(thinPen);
            return;
        }

        // 黒丸・白丸：軸の先に円を置く
        drawShaft(right - dotRadius * 2 + 1, thickPen);
        graphics.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
        graphics.fillPath(button.arrowIndex === 3 ? iconBrush : backBrush);
        if (button.arrowIndex === 4) {
            graphics.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, 2));
        }
    }

    /**
     * アイコンボタンの選択状態を返す
     * @param {Array<object>} buttons - アイコンボタンの配列
     * @returns {number} 選択中のインデックス
     */
    function getSelectedIconIndex(buttons) {
        for (var i = 0; i < buttons.length; i++) {
            if (buttons[i].isSelected) return i;
        }
        return 0;
    }

    /**
     * アイコンボタンの選択状態を切り替える
     * @param {Array<object>} buttons - アイコンボタンの配列
     * @param {number} index - 選択するインデックス
     * @returns {void}
     */
    function selectIcon(buttons, index) {
        if (!(index >= 0) || index >= buttons.length) return;
        for (var i = 0; i < buttons.length; i++) {
            buttons[i].isSelected = (i === index);
            try { buttons[i].notify("onDraw"); } catch (e) {}
        }
    }

    /**
     * 矢印の形状をアイコンで選ぶ行を追加する（行ラベルなし）
     * @param {object} parent - 追加先のパネル
     * @param {number} defaultIndex - 初期選択のインデックス
     * @param {function} onSelect - 選び直したときに呼ぶ処理
     * @returns {object} row（行グループ）と buttons（アイコンボタンの配列）
     */
    function addArrowShapeRow(parent, defaultIndex, onSelect) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["center", "top"];
        row.alignChildren = ["left", "center"];
        row.spacing = ICON_BUTTON_SPACING;
        row.margins = [0, 0, 0, ICON_ROW_BOTTOM_MARGIN];

        var buttons = [];
        for (var i = 0; i < ARROW_CHOICES.length; i++) {
            var button = row.add("button", undefined, "");
            button.minimumSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
            button.preferredSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
            button.maximumSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
            button.arrowIndex = i;
            button.isSelected = false;
            button.helpTip = getLabel(ARROW_CHOICES[i].label) + "  —  " + getLabel(LABELS.tooltip.arrowShape);
            button.onDraw = function () {
                drawArrowShapeIcon(this);
            };
            button.onClick = function () {
                selectIcon(buttons, this.arrowIndex);
                onSelect();
            };
            buttons.push(button);
        }
        selectIcon(buttons, defaultIndex);
        return { row: row, buttons: buttons };
    }

    /**
     * 起点にするオブジェクトを位置で選ぶダイアログを出す
     * @param {Array<object>} items - 選択したオブジェクト
     * @returns {number} 起点にするインデックス。閉じたときは -1
     */
    function chooseKeyIndex(items) {
        var anchorIndex = ANCHOR_DEFAULT_INDEX;
        ANCHOR_SELECTED_FILL = isLightUI() ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];

        var keyDialog = new Window("dialog", getLabel(LABELS.dialog.keyObject));
        keyDialog.orientation = "column";
        keyDialog.alignChildren = ["fill", "top"];
        keyDialog.margins = WINDOW_MARGINS;
        keyDialog.spacing = WINDOW_SPACING;

        var messageText = keyDialog.add("statictext", undefined, getLabel(LABELS.message.noKeyObject), { multiline: true });
        messageText.preferredSize.width = KEY_DIALOG_TEXT_WIDTH;

        var widgetGroup = keyDialog.add("group");
        widgetGroup.alignment = ["center", "top"];
        var anchorWidget = addAnchorWidget(widgetGroup, anchorIndex, function (index) {
            anchorIndex = index;
        });
        anchorWidget.helpTip = getLabel(LABELS.tooltip.keyObject);

        var manualKeyCheck = keyDialog.add("checkbox", undefined, getLabel(LABELS.check.manualKey));
        manualKeyCheck.helpTip = getLabel(LABELS.tooltip.manualKey);
        manualKeyCheck.onClick = function () {
            // 手動で設定してもらうため、そのまま閉じる
            keyDialog.close(DIALOG_RESULT_CANCEL);
        };

        /* ボタンエリア / Button row */
        var btnRowGroup = keyDialog.add("group");
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

        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnCancel.onClick = function () {
            keyDialog.close(DIALOG_RESULT_CANCEL);
        };

        var btnOK = btnRightGroup.add("button", undefined, "OK", { name: "ok" });
        btnOK.onClick = function () {
            keyDialog.close(DIALOG_RESULT_OK);
        };

        if (keyDialog.show() !== DIALOG_RESULT_OK) return -1;
        return getIndexAtAnchor(items, anchorIndex);
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
    presetDropdown.preferredSize.width = PRESET_LIST_WIDTH;
    presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);

    /* スペーサー（伸縮）：保存・削除ボタンを右端に寄せる / Spacer so the buttons stay right */
    var presetSpacer = presetRow.add("group");
    presetSpacer.alignment = ["fill", "fill"];
    presetSpacer.minimumSize.width = 0;
    var btnSavePreset = presetRow.add("button", undefined, getLabel(LABELS.button.save));
    btnSavePreset.preferredSize.width = PRESET_BUTTON_WIDTH;
    btnSavePreset.alignment = ["right", "center"];
    var btnRemovePreset = presetRow.add("button", undefined, getLabel(LABELS.button.remove));
    btnRemovePreset.preferredSize.width = PRESET_BUTTON_WIDTH;
    btnRemovePreset.alignment = ["right", "center"];

    /* コネクター / Connector */
    var connectorPanel = addPanel(dialog, getLabel(LABELS.panel.connector));

    var lineShapeField = addRadioRow(connectorPanel, LABELS.fieldLabel.lineShape, [LABELS.radio.shapeStraight, LABELS.radio.shapeWarp, LABELS.radio.shapeElbow, LABELS.radio.shapeBranch, LABELS.radio.shapeCurve], DEFAULT_LINE_SHAPE, LABELS.tooltip.lineShape);

    var warpTypeRow = addFieldRow(connectorPanel, LABELS.fieldLabel.warpType);
    var warpTypeList = warpTypeRow.add("dropdownlist", undefined, getWarpTypeLabels());
    warpTypeList.preferredSize.width = LIST_WIDTH;
    warpTypeList.selection = DEFAULT_WARP_TYPE;

    var warpAmountField = addNumberRow(connectorPanel, LABELS.fieldLabel.warpAmount, DEFAULT_WARP_AMOUNT, LABELS.unit.percent, LABELS.tooltip.warpAmount, [WARP_AMOUNT_MIN, WARP_AMOUNT_MAX]);
    var warpAxisField = addRadioRow(connectorPanel, LABELS.fieldLabel.warpAxis, [LABELS.radio.axisAuto, LABELS.radio.axisHorizontal, LABELS.radio.axisVertical], DEFAULT_WARP_AXIS, LABELS.tooltip.warpAxis);
    var cornerField = addNumberRow(connectorPanel, LABELS.fieldLabel.cornerRadius, DEFAULT_CORNER_RADIUS, LABELS.unit.pt, LABELS.tooltip.cornerRadius);
    var startPointField = addRadioRow(connectorPanel, LABELS.fieldLabel.startPoint, [LABELS.radio.startCenter, LABELS.radio.startDivided, LABELS.radio.startKeyCenter], DEFAULT_START_POINT, LABELS.tooltip.startPoint, true);

    /* 開始ポイントの右に9分割のウィジェットを置く / 3x3 widget at the right of the row */
    var anchorSpacer = startPointField.row.add("group");
    anchorSpacer.alignment = ["fill", "fill"];
    anchorSpacer.minimumSize.width = 0;
    ANCHOR_SELECTED_FILL = isLightUI() ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
    var unifyAnchorWidget = addAnchorWidget(startPointField.row, DEFAULT_UNIFY_ANCHOR, function () {
        updatePreview();
    });
    unifyAnchorWidget.alignment = ["right", "center"];
    unifyAnchorWidget.helpTip = getLabel(LABELS.tooltip.unifyAnchor);
    unifyAnchorWidget.enabled = false;

    var unifyStartRow = connectorPanel.add("group");
    unifyStartRow.orientation = "row";
    unifyStartRow.alignment = ["fill", "top"];
    unifyStartRow.alignChildren = ["left", "center"];
    var unifyStartSpacer = unifyStartRow.add("statictext", undefined, "");
    unifyStartSpacer.preferredSize.width = LABEL_WIDTH;
    var unifyStartCheck = unifyStartRow.add("checkbox", undefined, getLabel(LABELS.check.unifyStart));
    unifyStartCheck.helpTip = getLabel(LABELS.tooltip.unifyStart);

    /* 線・矢印は2カラム。ラベル幅はコネクターパネルより狭くし、矢印はさらに狭くする */
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
    setLabelWidth(ARROW_LABEL_WIDTH);
    var arrowPanel = addPanel(panelColumnsGroup, getLabel(LABELS.panel.arrow));
    initIconColors();
    var arrowShapeField = addArrowShapeRow(arrowPanel, DEFAULT_ARROW_INDEX, function () {
        // 矢印ごとに見え方が違うので、選び直したらその矢印の既定値（倍率・位置・線端）を入れる
        var selectedArrow = ARROW_CHOICES[getSelectedIconIndex(arrowShapeField.buttons)];
        if (selectedArrow.number !== 0) {
            arrowScaleField.input.text = selectedArrow.scale;
            selectRadio(arrowTipField.radios, selectedArrow.tip);
        }
        // 丸は線端を丸形に、矢印はなしにそろえる
        if (selectedArrow.cap !== undefined) selectRadio(strokeCapField.radios, selectedArrow.cap);
        updatePreview();
    });
    // 初期値も選択中の矢印の既定倍率にそろえる
    var arrowScaleField = addNumberRow(arrowPanel, LABELS.fieldLabel.arrowScale, ARROW_CHOICES[DEFAULT_ARROW_INDEX].scale, LABELS.unit.percent, LABELS.tooltip.arrowScale);
    var arrowPositionField = addRadioRow(arrowPanel, LABELS.fieldLabel.arrowPosition, [LABELS.radio.arrowEnd, LABELS.radio.arrowBoth], DEFAULT_ARROW_POSITION, LABELS.tooltip.arrowPosition, true);
    var arrowTipField = addRadioRow(arrowPanel, LABELS.fieldLabel.arrowTip, [ARROW_TIP_OPTIONS[0].label, ARROW_TIP_OPTIONS[1].label], ARROW_CHOICES[DEFAULT_ARROW_INDEX].tip, LABELS.tooltip.arrowTip, true);
    var endGapField = addNumberRow(arrowPanel, LABELS.fieldLabel.endGap, DEFAULT_END_GAP, LABELS.unit.pt, LABELS.tooltip.endGap);
    var strokeCapField = addRadioRow(arrowPanel, LABELS.fieldLabel.strokeCap, [STROKE_CAP_OPTIONS[0].label, STROKE_CAP_OPTIONS[1].label, STROKE_CAP_OPTIONS[2].label], DEFAULT_STROKE_CAP, LABELS.tooltip.strokeCap, true);


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
        var arrowChoice = ARROW_CHOICES[getSelectedIconIndex(arrowShapeField.buttons)];
        var arrowName = (arrowChoice.number === 0) ? noneName : getArrowName(arrowChoice.number);
        var arrowPosition = getSelectedIndex(arrowPositionField.radios);
        var lineShape = getSelectedIndex(lineShapeField.radios);

        return {
            strokeWidth: strokeWidth,
            strokeJoin: getSelectedIndex(strokeJoinField.radios),
            strokeCap: getSelectedIndex(strokeCapField.radios),
            dashStyle: getSelectedIndex(dashStyleField.radios),
            dashSegments: Math.round(toNumber(dashSegmentsField.input.text, DEFAULT_DASH_SEGMENTS)),
            dashGap: toNumber(dashGapField.input.text, DEFAULT_DASH_GAP),
            startPoint: getSelectedIndex(startPointField.radios),
            unifyStart: unifyStartCheck.value,
            unifyAnchor: unifyAnchorWidget.selectedAnchorIndex,
            lineShape: lineShape,
            warpName: warpType.name,
            warpStyle: warpType.style,
            warpAmount: toNumber(warpAmountField.input.text, DEFAULT_WARP_AMOUNT),
            warpAxis: getSelectedIndex(warpAxisField.radios),
            // 角丸はカギのときだけ。分岐は角を丸めない / elbow only; Branch keeps square corners
            cornerRadius: (lineShape === 2 || lineShape === 3) ? toNumber(cornerField.input.text, DEFAULT_CORNER_RADIUS) : 0,
            hasArrow: (arrowChoice.number !== 0),
            arrowInnerDot: (arrowChoice.innerDot === true),
            startArrow: (arrowPosition === 1) ? arrowName : noneName,
            endArrow: arrowName,
            arrowScale: toNumber(arrowScaleField.input.text, DEFAULT_ARROW_SCALE),
            endGap: toNumber(endGapField.input.text, DEFAULT_END_GAP),
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
        // カーブのふくらみもこの欄で決める / Curve reuses this field for its bow
        warpAmountField.row.enabled = (settings.lineShape === 1 || settings.lineShape === 4);
        warpAxisField.row.enabled = (settings.lineShape === 1);
        cornerField.row.enabled = (settings.lineShape === 2 || settings.lineShape === 3);
        if (unifyAnchorWidget.enabled !== settings.unifyStart) {
            unifyAnchorWidget.enabled = settings.unifyStart;
            try { unifyAnchorWidget.notify("onDraw"); } catch (e) {}
        }
        dashSegmentsField.row.enabled = (settings.dashStyle === 1);
        dashGapField.row.enabled = (settings.dashStyle !== 0);
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
            { key: "strokeCap",     radios: strokeCapField.radios },
            { key: "endGap",        field: endGapField },
            { key: "dashStyle",     radios: dashStyleField.radios },
            { key: "dashSegments",  field: dashSegmentsField },
            { key: "dashGap",       field: dashGapField },
            { key: "lineShape",     radios: lineShapeField.radios },
            { key: "warpType",      list: warpTypeList },
            { key: "warpAmount",    field: warpAmountField },
            { key: "warpAxis",      radios: warpAxisField.radios },
            { key: "cornerRadius",  field: cornerField },
            { key: "startPoint",    radios: startPointField.radios },
            { key: "unifyStart",    check: unifyStartCheck },
            { key: "unifyAnchor",   widget: unifyAnchorWidget },
            { key: "arrowIndex",    icons: arrowShapeField.buttons },
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
            } else if (entry.icons) {
                preset[entry.key] = getSelectedIconIndex(entry.icons);
            } else if (entry.check) {
                preset[entry.key] = entry.check.value;
            } else if (entry.widget) {
                preset[entry.key] = entry.widget.selectedAnchorIndex;
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
            } else if (entry.icons) {
                selectIcon(entry.icons, value);
            } else if (entry.check) {
                entry.check.value = (value === true || value === "true");
            } else if (entry.widget) {
                entry.widget.selectedAnchorIndex = value;
                try { entry.widget.notify("onDraw"); } catch (e) {}
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
    unifyStartCheck.onClick = updatePreview;
    bindPreviewToRadios(strokeJoinField.radios);
    for (var i = 0; i < dashStyleField.radios.length; i++) {
        dashStyleField.radios[i].onClick = function () {
            // ドットは丸形の線端でないと点が出ないので、選んだときに丸形へそろえる
            if (getSelectedIndex(dashStyleField.radios) === 2) selectRadio(strokeCapField.radios, 1);
            updatePreview();
        };
    }
    bindPreviewToRadios(strokeCapField.radios);
    bindPreviewToRadios(warpAxisField.radios);
    bindPreviewToRadios(arrowPositionField.radios);
    bindPreviewToRadios(arrowTipField.radios);
    bindPreviewToField(strokeWidthField, false);
    bindPreviewToField(dashSegmentsField, false);
    bindPreviewToField(dashGapField, false);
    bindPreviewToField(warpAmountField, true);
    bindPreviewToField(cornerField, false);
    bindPreviewToField(arrowScaleField, false);
    bindPreviewToField(endGapField, false);
    warpTypeList.onChange = updatePreview;
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
