#target illustrator
#targetengine "SmartObjectDistributorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SmartObjectDistributor
Meta-Default-language=ja
Name=グリッドの各セル中央にオブジェクトを配置
Name[en]=Place objects at the center of each grid cell
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=3.9.1
Release-Date=2025-05-20
Tags=alignment,整列
X-com.swwwitch-Script-Name=SmartObjectDistributor
X-com.swwwitch-Updated=2026-11-22
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectDistributor.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectDistributor.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/na3c45cea09b7
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartObjectDistributor.jsx
Description-BEGIN
[ja]
選択したオブジェクトを、指定した行数・列数のグリッドの各セル中央へ配置します。
セルに合わせたリサイズやマスク、開始位置を指定でき、プレビューを見ながら調整できます。
[en]
Places the selected objects at the center of each cell of a grid with the rows and columns you specify.
Objects can be resized and masked to fit the cells, and the fill order can be set, all with a live preview.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var CONFIG = {
        defaultGutter: 10,           /* 間隔の初期値（定規単位） / default gutter */
        defaultMargin: 10,           /* マージンの初期値（定規単位） / default margin */
        defaultGutterPriority: false, /* ［セル間隔を優先］の初期値 / Prioritize Gutter starts unchecked */
        defaultResizeRatio: 100,     /* リサイズの比率の初期値（%） / default resize ratio */
        imageResizeMode: "area",     /* 画像を読み込んだときのリサイズの初期値（"none" / "width" / "height" / "area" / "scale"） / resize mode after importing images */
        maxResizeRatio: 1000,        /* リサイズの比率の上限（%） / max resize ratio */
        defaultMaskMode: "cell",     /* マスクの初期値（"none"：なし / "cell"：セル / "item"：画像） / default mask mode */
        defaultUseCornerRadius: false, /* ［角丸］の初期値 / Round Corners starts unchecked */
        defaultCornerRadius: 10,     /* 角丸の半径の初期値（定規単位） / default corner radius (ruler unit) */
        defaultUseFrame: false,      /* ［枠］の初期値 / Frame starts unchecked */
        defaultFrameWidth: 1,        /* 枠の線幅の初期値（線の単位） / default frame weight (stroke unit) */
        defaultItemAnchor: "center", /* アイテムの基準位置の初期値（開始位置と同じ名前） / where each object sits in its cell */
        defaultKeepOriginals: false, /* ［元のオブジェクトを残す］の初期値 / Keep Originals starts unchecked */
        defaultFitView: true,        /* ［画面にフィット］の初期値 / Fit to Window starts checked */
        defaultFitViewPercent: 65,   /* ［画面にフィット］の割合の初期値（%） / default share of the window */
        defaultSpreadEvenly: false,  /* ［均等にまばらに］の初期値 / Spread Evenly starts unchecked */
        defaultBrick: false,         /* ［レンガ］の初期値 / Brick starts unchecked */
        defaultFillStart: "topLeft", /* 開始位置の初期値（"topLeft" / "top" / "topRight" / "left" / "center" / "right" / "bottomLeft" / "bottom" / "bottomRight"） / where filling starts */
        blackCellOpacity: 15,        /* 黒セルの不透明度（%） / opacity for black cells */
        whiteCellOpacity: 100,       /* 白セルの不透明度（%） / opacity for white cells */
        fallbackDivision: 5,         /* 分割数を決められないときの既定値 / fallback division */
        maxDivision: 100,            /* 行・列の上限 / max rows or columns */
        maxCellCount: 1000,          /* セル総数の上限 / max total cells */
        previewThrottleCells: 200,   /* この数を超えるセルでは入力中のプレビューを間引く / throttle preview above this many cells */
        previewThrottleMs: 200,      /* 間引く間隔（ミリ秒） / throttle window in milliseconds */
        parkingStep: 200,            /* 退避時の移動量（pt） / step when parking objects */
        overlapTolerance: 1,         /* 重なり判定の余裕（pt） / overlap tolerance */
        artboardBuffer: 10,          /* 他アートボードとの余裕（pt） / buffer around artboards */
        cellLayerName: "cell-background",            /* セル描画レイヤー名 / cell layer name */
        previewCellLayerName: "_Preview_Background", /* プレビュー用レイヤー名 / preview layer name */
        legacyPreviewLayerName: "_Preview_Guides"    /* 旧版のプレビュー用レイヤー名 / legacy preview layer */
    };

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

    /* 間隔は共通部品 UILayout と同じにそろえる / Spacing matches the shared UILayout part */
    var RADIO_COLUMN_SPACING = 6;              /* 縦に並べたラジオボタンの間隔 / spacing of stacked radio buttons */
    var PANEL_SECTION_SPACING = PANEL_SPACING; /* パネル内のまとまりどうしの間隔 / spacing between sections in a panel */
    var COLUMN_PANEL_SPACING = WINDOW_SPACING; /* カラム内のパネルどうしの間隔 / spacing between panels in a column */
    var FILL_START_WIDGET_SIZE = 50; /* 開始位置のウィジェットの一辺 / size of the Start From widget */
    var COLOR_SWATCH_SIZE = [40, 20]; /* 枠の色見本の大きさ / size of the frame color swatch */
    var CHECKBOX_BOX_WIDTH = 24;  /* チェックボックスの□と文字の間の分の幅 / width of a checkbox's box and its gap */
    var PRESET_ICON_SIZE = [20, 18];     /* プリセットの保存・削除アイコンの大きさ / preset save and delete icon size */
    var SAVE_ICON_OPACITY = 0.75;        /* 保存アイコンの濃さ（ほかのアイコンに対する比率） / save icon opacity relative to the others */
    var PRESET_DROPDOWN_MIN_WIDTH = 60;  /* プリセットのドロップダウンの最小幅（あとはパネルの幅まで伸びる） / preset dropdown minimum width; it stretches to the panel */
    var PRESET_NAME_CHARS = 20;          /* プリセット名の入力欄の文字数 / preset name field width */

    /**
     * コントロールの文字の幅を測る
     * @param {Object} control - text を持つコントロール
     * @returns {number} 文字の幅（px）
     */
    function measureControlText(control) {
        return Math.ceil(control.graphics.measureString(control.text)[0]);
    }

    /**
     * ダイアログの1カラム分のグループを作る（パネルを縦に積み、幅をそろえる）
     * @param {Group} columnsGroup - カラムを並べる横並びのグループ
     * @returns {Group} カラムのグループ
     */
    function addDialogColumn(columnsGroup) {
        var dialogColumn = columnsGroup.add("group");
        dialogColumn.orientation = "column";
        dialogColumn.alignment = ["fill", "fill"];
        dialogColumn.alignChildren = ["fill", "top"];
        dialogColumn.spacing = COLUMN_PANEL_SPACING;
        return dialogColumn;
    }

    /**
     * ラジオボタンを縦に並べるグループを作る（排他は同じ親の中だけなので、1組につき1グループ）
     * @param {Panel|Group} parent - 追加先
     * @returns {Group} ラジオボタンを入れるグループ
     */
    function addRadioColumn(parent) {
        var radioColumn = parent.add("group");
        radioColumn.orientation = "column";
        radioColumn.alignment = ["left", "top"];
        radioColumn.alignChildren = ["left", "top"];
        radioColumn.spacing = RADIO_COLUMN_SPACING;
        return radioColumn;
    }

    // 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)

    // -----------------------------------------
    // 基準点ウィジェットの寸法 / Anchor widget metrics
    // -----------------------------------------
    var ANCHOR_WIDGET_SIZE      = 66;   /* ウィジェット全体の一辺 / overall size of the widget */
    var ANCHOR_WIDGET_CELL_SIZE = 9;    /* □1個の一辺 / size of one square */
    var ANCHOR_WIDGET_CELL_GAP  = 7.5;  /* □どうしの間隔 / gap between squares */
    var ANCHOR_WIDGET_NONE      = -1;   /* 未選択のインデックス / index while nothing is selected */

    /* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
       Cell names in row-major order, matching the Transformation enumeration */
    var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線 / Rules joining the outer squares (the center stands alone) */
    var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    // -----------------------------------------
    // 基準点ウィジェットの配色 / Anchor widget colors
    // -----------------------------------------
    var ANCHOR_WIDGET_UI_DARK = isDarkUI();
    /* 枠線・ケイ線はグレー、選択セルの塗りはライトで濃いグレー・ダークで明るいグレー（既存スクリプトの配色を踏襲）。
       無効時は同じ色を半透明にして背景へ沈める（不透明の薄いグレーだとダークUIで逆に明るく浮くため）
       Gray rules; the selected fill is dark gray on light UI and light gray on dark UI (as in the existing scripts).
       Disabled colors are translucent versions so they sink into any background */
    var ANCHOR_WIDGET_LINE_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 1]     : [0.42, 0.42, 0.42, 1];  /* 枠線・ケイ線 / rules */
    var ANCHOR_WIDGET_FILL_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 1]     : [0.27, 0.27, 0.27, 1];  /* 選択セルの塗り / selected fill */
    var ANCHOR_WIDGET_DIM_LINE_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 0.4]   : [0.42, 0.42, 0.42, 0.4];  /* 無効時の枠線 / rules when disabled */
    var ANCHOR_WIDGET_DIM_FILL_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 0.3]   : [0.27, 0.27, 0.27, 0.3];  /* 無効時の塗り / fill when disabled */

    // -----------------------------------------
    // ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 基準点（3×3）を選ぶウィジェットを追加する。クリックしたセルを選び、onChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
     * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
     * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）/ size（一辺。既定 66）
     * @returns {Button} ウィジェット（値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
     */
    function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
        var anchorOptions = widgetOptions || {};
        var widgetSize = anchorOptions.size || ANCHOR_WIDGET_SIZE;
        var anchorWidget = parent.add("button", undefined, "");
        anchorWidget.minimumSize = [widgetSize, widgetSize];
        anchorWidget.preferredSize = [widgetSize, widgetSize];
        anchorWidget.maximumSize = [widgetSize, widgetSize];
        anchorWidget.isAnchorWidget = true; /* redrawAnchorWidgetsIn() の目印 / marker for redrawAnchorWidgetsIn() */
        anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(anchorOptions.disabledCells);
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(initialValue, anchorWidget.anchorAllowNone);
        anchorWidget.onDraw = function () { drawAnchorWidget(anchorWidget); };
        anchorWidget.onClick = function () {}; /* セルの判定は mousedown で行う / hit-testing happens in mousedown */

        /* クリック座標（コントロール基準）を3分割してセルを判定する / split the control-relative click into thirds */
        anchorWidget.addEventListener("mousedown", function (event) {
            if (!isAnchorWidgetEnabledInTree(anchorWidget)) return;
            var cellIndex = getAnchorCellAt(event.clientX, event.clientY, anchorWidget.size[0], anchorWidget.size[1]);
            if (anchorWidget.anchorDisabledCells[cellIndex]) return;
            anchorWidget.anchorWidgetIndex = cellIndex;
            redrawAnchorWidget(anchorWidget);
            if (onChange) onChange(cellIndex, anchorWidget);
        });
        return anchorWidget;
    }

    /**
     * 選択中のセルのインデックスを返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {number} 0〜8（行優先）。未選択なら -1
     */
    function getAnchorWidgetIndex(anchorWidget) {
        return anchorWidget.anchorWidgetIndex;
    }

    /**
     * 選択中のセルの名前を返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {string} "topLeft" など。未選択なら ""
     */
    function getAnchorWidgetName(anchorWidget) {
        return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
    }

    /**
     * 選択するセルを変えて描き直す（onChange は呼ばない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * ウィジェットの有効／無効を切り替えて描き直す（無効の間は薄く描き、クリックも無視する）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
        anchorWidget.enabled = isEnabled;
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * 選べないセルを指定し直して描き直す（選択中のセルは変えない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
     * @returns {void}
     */
    function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(disabledCells);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * コンテナ以下にある基準点ウィジェットをすべて描き直す。パネルや行の enabled を切り替えたあとに呼ぶ
     * @param {Object} container - パネル・グループ・ウィンドウなど
     * @returns {void}
     */
    function redrawAnchorWidgetsIn(container) {
        if (container.isAnchorWidget) {
            redrawAnchorWidget(container);
            return;
        }
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            redrawAnchorWidgetsIn(container.children[i]);
        }
    }

    // -----------------------------------------
    // 値の変換 / Value helpers
    // -----------------------------------------
    /**
     * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
     * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
     * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
     * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
     */
    function resolveAnchorWidgetIndex(anchorValue, allowNone) {
        if (typeof anchorValue === "string") {
            for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
                if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
            }
            return 4;
        }
        if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
        if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
            return anchorValue;
        }
        return 4;
    }

    /**
     * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [横の割合, 縦の割合]
     */
    function getAnchorRatio(anchorValue) {
        var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    }

    /**
     * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
     * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [x, y]
     */
    function getAnchorPointOnBounds(bounds, anchorValue) {
        var anchorRatio = getAnchorRatio(anchorValue);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
            bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
        ];
    }

    /**
     * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
     * 基準は効果を含まない境界（geometricBounds）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {Transformation} Transformation.TOPLEFT など
     */
    function getAnchorTransformation(anchorValue) {
        var transformations = [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ];
        return transformations[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * symbols.add() に渡す登録点を返す（Illustrator 専用）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
     */
    function getAnchorSymbolRegistrationPoint(anchorValue) {
        var registrationPoints = [
            SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
        ];
        return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * クリック位置からセルのインデックスを求める（ウィジェットを縦横3等分し、外にはみ出した座標は端のセルに寄せる）
     * @param {number} clickX - コントロール基準の x
     * @param {number} clickY - コントロール基準の y
     * @param {number} widgetWidth - ウィジェットの幅
     * @param {number} widgetHeight - ウィジェットの高さ
     * @returns {number} 0〜8
     */
    function getAnchorCellAt(clickX, clickY, widgetWidth, widgetHeight) {
        var column = Math.min(2, Math.max(0, Math.floor(clickX / (widgetWidth / 3))));
        var row = Math.min(2, Math.max(0, Math.floor(clickY / (widgetHeight / 3))));
        return row * 3 + column;
    }

    /**
     * セルのインデックスの配列を、9個の真偽値に直す
     * @param {number[]} [cellIndexes] - セルのインデックスの配列
     * @returns {boolean[]} 含まれるセルだけ true
     */
    function toAnchorCellFlags(cellIndexes) {
        var cellFlags = [false, false, false, false, false, false, false, false, false];
        if (!cellIndexes) return cellFlags;
        for (var i = 0; i < cellIndexes.length; i++) {
            if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
        }
        return cellFlags;
    }

    // -----------------------------------------
    // 描画 / Drawing
    // -----------------------------------------
    /**
     * ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立。選択セルだけ塗る）
     * @param {Button} anchorWidget - 描くウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        var halfCell = cellSize / 2;
        /* 自作描画は自動でディムにならないので、親までたどって判定する / custom drawing is not dimmed automatically */
        var isEnabled = isAnchorWidgetEnabledInTree(anchorWidget);

        /* ボタンの地をコントロールの地色で塗り、パネルに溶け込ませる（backgroundColor が無い環境では例外）
           Paint the control's own background so the widget blends into the panel; throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
        var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);
        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1);
        for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
                graphics.lineTo(cellB[0], cellB[1] + halfCell);
            } else {
                /* 縦方向：下の□へ / vertical: to the square below */
                graphics.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
                graphics.lineTo(cellB[0] + halfCell, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < 9; i++) {
            var isCellEnabled = isEnabled && !anchorWidget.anchorDisabledCells[i];
            drawAnchorWidgetCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.anchorWidgetIndex, isCellEnabled);
        }
    }

    /**
     * □を1つ描く（選択中だけ塗り、枠は塗りの上に重ねる）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} cellX - 左端
     * @param {number} cellY - 上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {boolean} isEnabled - 選べるセルなら true（false なら薄く描く）
     * @returns {void}
     */
    function drawAnchorWidgetCell(graphics, cellX, cellY, isSelected, isEnabled) {
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        /* rectPath の前には毎回 newPath()（呼ばないとパスが累積して塗りが線画になる）
           Always call newPath() before rectPath(), or paths accumulate and fills turn into outlines */
        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(cellX, cellY, cellSize, cellSize);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_FILL_COLOR : ANCHOR_WIDGET_DIM_FILL_COLOR));
        }
        graphics.newPath();
        graphics.rectPath(cellX, cellY, cellSize, cellSize);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isAnchorWidgetEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false) return false;
        }
        return true;
    }

    /**
     * ウィジェットの onDraw を呼び直す。notify("onDraw") は環境によって例外や空振りになるため、隠して再表示して描き直させる
     * @param {Button} anchorWidget - 描き直すウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        anchorWidget.hide();
        anchorWidget.show();
    }

    // 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPTMETA の Version）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = getVersionInfoValue("version");
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA の項目を getScriptMetaValue() で直接読み（スクリプト情報の部品）、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPTMETA fields directly with getScriptMetaValue(),
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 「このスクリプトについて」とフッターに出す値を返す。SCRIPTMETA の項目を getScriptMetaValue() で直接読み、名前は SCRIPT_NAME。
     * getScriptMetaValue() が無いスクリプトは従来の SCRIPT_* を typeof で読む（未定義でも例外にならない）
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値（バージョンは「v1.2.0」の形）。無ければ空文字
     */
    function getVersionInfoValue(variableKey) {
        var infoValue = "";
        if (variableKey === "name") {
            infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        } else if (typeof getScriptMetaValue === "function") {
            var isJapanese = (VERSION_INFO_LANG === "ja");
            if (variableKey === "version") infoValue = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : "";
            else if (variableKey === "feature") infoValue = getScriptMetaValue("Name");
            else if (variableKey === "readme") infoValue = getScriptMetaValue(isJapanese ? "X-com.swwwitch-README-JA" : "X-com.swwwitch-README-EN");
            else if (variableKey === "source") infoValue = getScriptMetaValue("X-com.swwwitch-Source-URL");
            else if (variableKey === "article") infoValue = getScriptMetaValue("X-com.swwwitch-Article-URL");
        } else {
            if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
            else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
            else if (variableKey === "readme") {
                if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
                else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
            }
            else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
            else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        }
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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

    // アイコンのボタン（再利用パーツ） / Icon buttons (reusable)

    var ICON_BUTTON_SIZE = [24, 22]; /* 既定の大きさ / default size */
    var ICON_BUTTON_UI_DARK = isDarkUI();
    var ICON_BUTTON_COLOR     = ICON_BUTTON_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* 絵の色 / icon color */
    var ICON_BUTTON_DIM_COLOR = ICON_BUTTON_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の色 / color when disabled */

    /**
     * onDraw で絵を描くアイコンのボタンを追加する。押すと onClick を呼ぶ（無効の間は押せず、薄く描く）
     * @param {Group} parent - 追加先
     * @param {number[]} iconSize - [幅, 高さ]
     * @param {Function} drawIcon - 絵を描く関数 (iconGraphics, iconWidth, iconHeight, iconColor)
     * @returns {Group} アイコン
     */
    function addIconButton(parent, iconSize, drawIcon) {
        var iconButton = parent.add("group");
        iconButton.preferredSize = iconSize;
        iconButton.minimumSize = iconSize;
        iconButton.maximumSize = iconSize;
        iconButton.onDraw = function () {
            var iconColor = isIconButtonEnabledInTree(iconButton) ? ICON_BUTTON_COLOR : ICON_BUTTON_DIM_COLOR;
            drawIcon(iconButton.graphics, iconSize[0], iconSize[1], iconColor);
        };
        iconButton.addEventListener("mousedown", function () {
            if (!isIconButtonEnabledInTree(iconButton)) return;
            if (typeof iconButton.onClick === "function") iconButton.onClick();
        });
        return iconButton;
    }

    /**
     * アイコンのボタンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} iconButton - addIconButton() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setIconButtonEnabled(iconButton, isEnabled) {
        if (iconButton.enabled === isEnabled) return;
        iconButton.enabled = isEnabled;
        /* group には notify() が無いため、隠して再表示して描き直させる / groups have no notify(), so hide and show to repaint */
        iconButton.hide();
        iconButton.show();
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isIconButtonEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * 絵の座標の長方形の並びを、縦横比を保って中央に置いて塗る（ScriptUI は多角形を塗れないため、斜めも長方形の並びで描く）
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} designSize - 絵の [幅, 高さ]
     * @param {number[][]} designRects - 長方形 [左, 上, 右, 下] の並び
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function fillIconButtonRects(iconGraphics, iconWidth, iconHeight, designSize, designRects, iconColor) {
        var iconScale = Math.min(iconWidth / designSize[0], iconHeight / designSize[1]);
        var originX = (iconWidth - designSize[0] * iconScale) / 2;
        var originY = (iconHeight - designSize[1] * iconScale) / 2;
        iconGraphics.newPath();
        for (var i = 0; i < designRects.length; i++) {
            var designRect = designRects[i];
            if (designRect[2] <= designRect[0] || designRect[3] <= designRect[1]) continue;
            iconGraphics.rectPath(originX + designRect[0] * iconScale, originY + designRect[1] * iconScale,
                (designRect[2] - designRect[0]) * iconScale, (designRect[3] - designRect[1]) * iconScale);
        }
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
    }

    /**
     * 保存アイコン（トレイに下向きの矢印）を描く。1223×993 の絵。
     * 矢じりとトレイの V 字の切り欠きは細い長方形を並べ、トレイの四角い穴は塗らずに残す
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawSaveIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var designWidth = 1223;
        var trayHole = [78, 688, 230, 840]; /* トレイの四角い穴 [左, 上, 右, 下] / the tray's square hole */
        var sliceCount = 16;
        var designRects = [[535, 0, 688, 383]]; /* 矢印の軸 / arrow shaft */
        var k;

        /** 穴と重なる部分を除いて長方形を足す / add a rectangle minus the tray hole */
        function addTrayRect(left, top, right, bottom) {
            if (right <= trayHole[0] || left >= trayHole[2] || bottom <= trayHole[1] || top >= trayHole[3]) {
                designRects.push([left, top, right, bottom]);
                return;
            }
            designRects.push([left, top, right, trayHole[1]], [left, trayHole[3], right, bottom],
                [left, Math.max(top, trayHole[1]), trayHole[0], Math.min(bottom, trayHole[3])],
                [trayHole[2], Math.max(top, trayHole[1]), right, Math.min(bottom, trayHole[3])]);
        }

        /* 下向きの矢じり（上の付け根から先端へ細っていく）/ the downward head */
        for (k = 0; k < sliceCount; k++) {
            var headHalf = 251 * (1 - (k + 0.5) / sliceCount);
            var headTop = 383 + (703 - 383) * k / sliceCount;
            designRects.push([612 - headHalf, headTop, 612 + headHalf, headTop + (703 - 383) / sliceCount]);
        }
        /* トレイ：上辺に V 字の切り欠き（幅 458 から下へ細り、856 で閉じる）/ tray with a V notch along the top */
        for (k = 0; k < sliceCount; k++) {
            var notchTop = 612 + (856 - 612) * k / sliceCount;
            var notchBottom = notchTop + (856 - 612) / sliceCount;
            var notchHalf = 229 * (1 - (k + 0.5) / sliceCount);
            addTrayRect(0, notchTop, 612 - notchHalf, notchBottom);
            addTrayRect(612 + notchHalf, notchTop, designWidth, notchBottom);
        }
        addTrayRect(0, 856, designWidth, 993);
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [designWidth, 993], designRects, iconColor);
    }

    /**
     * 削除アイコン（ゴミ箱）を描く。756×825 の絵
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawTrashIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [756, 825], [
            [206, 0, 550, 70], [206, 70, 275, 137], [481, 70, 550, 137],     /* 取っ手 / handle */
            [0, 137, 756, 207],                                               /* ふた / lid */
            [69, 207, 138, 825], [618, 207, 688, 825], [138, 756, 618, 825], /* 本体 / body */
            [206, 275, 275, 687], [343, 275, 413, 687], [481, 275, 550, 687] /* 縦の線 / ribs */
        ], iconColor);
    }

    /**
     * 自動計算アイコン（魔法の杖と3つのきらめき）を描く。24×22 の絵を、縦横比を保って中央に置く。
     * 行数・列数や比率など、値を自動で求める［自動］ボタンの代わりに使う
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawAutoIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var iconScale = Math.min(iconWidth / 24, iconHeight / 22);
        var originX = (iconWidth - 24 * iconScale) / 2;
        var originY = (iconHeight - 22 * iconScale) / 2;

        /** 絵の座標で線を1本引く / stroke one line in design coordinates */
        function strokeLine(x1, y1, x2, y2, lineWidth) {
            iconGraphics.newPath();
            iconGraphics.moveTo(originX + x1 * iconScale, originY + y1 * iconScale);
            iconGraphics.lineTo(originX + x2 * iconScale, originY + y2 * iconScale);
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, lineWidth * iconScale));
        }

        /** 十字のきらめきを描く / draw a cross-shaped sparkle */
        function drawSparkle(centerX, centerY, radius) {
            strokeLine(centerX - radius, centerY, centerX + radius, centerY, 1.2);
            strokeLine(centerX, centerY - radius, centerX, centerY + radius, 1.2);
        }

        strokeLine(4, 19, 14, 9, 2.2);     /* 杖 / wand */
        strokeLine(14, 9, 16.5, 6.5, 3.2); /* 杖の先 / wand tip */
        drawSparkle(18.5, 12.5, 2.5);
        drawSparkle(9, 5, 2.5);
        drawSparkle(20, 3.5, 1.8);
    }

    // アイコンのボタン（再利用パーツ）ここまで / End of the reusable icon buttons

    // 画面にフィット（再利用パーツ、_templates/FitViewToItems.jsx） / Fit view to items (reusable)

    var FitViewToItems = (function () {

        // =========================================
        // ユーザー設定 / User Settings
        // =========================================

        /* ウィンドウに対して対象が占める割合の既定値（%）と範囲。呼び出し側で上書きできる
           Default share of the window the items fill, in percent, and its range; callers can override the default */
        var DEFAULT_FIT_PERCENT = 65;
        var FIT_PERCENT_RANGE = [10, 100];

        /* Illustratorが受け付ける表示倍率の範囲（3.125%〜6400%） / Zoom range Illustrator accepts */
        var VIEW_ZOOM_RANGE = [0.03125, 64];

        // =========================================
        // ローカライズ / Localization
        // =========================================
        var LABELS = {
            checkbox: {
                fitView: { ja: "画面にフィット", en: "Fit to Window" }
            },
            tooltip: {
                fitView: {
                    ja: "作成するオブジェクトが収まるよう表示倍率を合わせます。",
                    en: "Refits the view to the objects being created."
                },
                fitViewPercent: {
                    ja: "ウィンドウに対するオブジェクトの大きさ（100%でいっぱい）",
                    en: "Size of the objects relative to the window; 100% fills it"
                }
            }
        };

        /**
         * UI言語を返す
         * @returns {string} "ja" または "en"
         */
        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }

        /**
         * LABELS の組から指定言語の文言を返す
         * @param {object} labelSet - { ja: string, en: string } の組
         * @param {string} uiLang - "ja" または "en"
         * @returns {string} 文言（無ければ英語）
         */
        function getLabel(labelSet, uiLang) {
            return (labelSet[uiLang] != null) ? labelSet[uiLang] : labelSet.en;
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        /**
         * 数値を範囲に収める。数値として読めないときは既定値を返す
         * @param {string|number} value - 入力値
         * @param {number[]} range - [下限, 上限]
         * @param {number} fallbackValue - 読めないときの既定値
         * @returns {number} 範囲内の数値
         */
        function clampNumber(value, range, fallbackValue) {
            var numberValue = Number(value);
            if (isNaN(numberValue) || (typeof value === "string" && !/\S/.test(value))) numberValue = fallbackValue;
            return Math.min(range[1], Math.max(range[0], numberValue));
        }

        /**
         * 「□画面にフィット［65］%」の行を追加する（ラベルとツールチップは内蔵）
         * @param {Group|Panel|Window} parentContainer - 追加先のコンテナ
         * @param {object} [rowOptions] - value: チェックの初期値（既定 false）／percent: 割合の初期値／lang: 表示言語／
         *     stepper: true で入力欄の左に∧∨を付け、↑↓キーでも増減（StepperButtons 部品の addStepper / bindSteppedArrowKeys /
         *     redrawSteppersIn を使う。増減したら percentInput.onChange を呼ぶ）
         * @returns {{row: Group, checkbox: Checkbox, percentInput: EditText, getFillRatio: function, updateEnabled: function}} 作成したコントロール一式
         */
        function addControls(parentContainer, rowOptions) {
            if (!rowOptions) rowOptions = {};
            var uiLang = rowOptions.lang || getCurrentLang();

            var fitViewRow = parentContainer.add("group");
            fitViewRow.orientation = "row";
            fitViewRow.alignChildren = ["left", "center"];
            fitViewRow.spacing = 6;

            var fitViewCheck = fitViewRow.add("checkbox", undefined, getLabel(LABELS.checkbox.fitView, uiLang));
            fitViewCheck.helpTip = getLabel(LABELS.tooltip.fitView, uiLang);
            /* 明示的に true を渡したときだけONで始める / only an explicit true starts it checked */
            fitViewCheck.value = (rowOptions.value === true);

            var fitPercent = (rowOptions.percent > 0) ? rowOptions.percent : DEFAULT_FIT_PERCENT;
            var percentInput;
            var percentParent = fitViewRow;
            var stepperGroup = null;
            if (rowOptions.stepper === true) {
                /* ∧∨と入力欄は隙間0で突き合わせる / Butt the stepper against the field */
                percentParent = fitViewRow.add("group");
                percentParent.orientation = "row";
                percentParent.alignChildren = ["left", "center"];
                percentParent.spacing = 0;
                percentParent.margins = 0;
                stepperGroup = addStepper(percentParent, function () { return percentInput; }, {
                    integer: true,
                    min: FIT_PERCENT_RANGE[0],
                    max: FIT_PERCENT_RANGE[1],
                    /* text の代入ではイベントが来ないので、増減後に onChange を呼ぶ / assigning text fires no event */
                    onStep: function () { if (percentInput.onChange) percentInput.onChange(); }
                });
            }
            percentInput = percentParent.add("edittext", undefined, String(fitPercent));
            percentInput.characters = 3;
            percentInput.helpTip = getLabel(LABELS.tooltip.fitViewPercent, uiLang);
            if (stepperGroup) bindSteppedArrowKeys(percentInput, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */
            var percentUnitLabel = fitViewRow.add("statictext", undefined, "%");

            var controls = {
                row: fitViewRow,
                checkbox: fitViewCheck,
                percentInput: percentInput,

                /**
                 * 入力欄の割合を 0〜1 の比率で返す
                 * @returns {number} ウィンドウに対して占める割合（1でいっぱい）
                 */
                getFillRatio: function () {
                    return clampNumber(percentInput.text, FIT_PERCENT_RANGE, fitPercent) / 100;
                },

                /**
                 * 割合の入力欄をチェックの状態に合わせて有効・無効にする
                 * @returns {void}
                 */
                updateEnabled: function () {
                    percentInput.enabled = fitViewCheck.value;
                    percentUnitLabel.enabled = fitViewCheck.value;
                    if (!stepperGroup) return;
                    stepperGroup.enabled = fitViewCheck.value;
                    /* ∧∨は自作描画なので、表示中なら描き直してディム表示を切り替える / redraw the custom-drawn stepper once shown */
                    if (fitViewRow.window.visible) redrawSteppersIn(stepperGroup);
                }
            };
            controls.updateEnabled();
            return controls;
        }

        /**
         * 複数アイテムを囲む外接範囲を求める（効果を含まない geometricBounds）
         * @param {PageItem[]} targetItems - 対象アイテム
         * @returns {number[]|null} [left, top, right, bottom]（求められない場合は null）
         */
        function getItemsBounds(targetItems) {
            var unionBounds = null;
            for (var i = 0; i < targetItems.length; i++) {
                var itemBounds = targetItems[i].geometricBounds;
                if (unionBounds === null) {
                    unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
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
         * 対象が指定の割合でウィンドウに収まるよう、中心を合わせて表示倍率を変える
         * 拡大・縮小のどちらも行う（KeepInView と違い、常に同じ大きさに見せる）
         * @param {PageItem[]} targetItems - 対象アイテム
         * @param {object} [fitOptions] - doc: 対象ドキュメント（省略時は最前面）／fillRatio: 占める割合（1でいっぱい）
         * @returns {boolean} 表示を動かしたら true
         */
        function fit(targetItems, fitOptions) {
            if (!targetItems || targetItems.length === 0) return false;
            if (!fitOptions) fitOptions = {};

            var targetDoc = fitOptions.doc || app.activeDocument;
            var bounds = getItemsBounds(targetItems);
            if (bounds === null) return false;

            var itemWidth = bounds[2] - bounds[0];
            var itemHeight = bounds[1] - bounds[3];
            var activeView = targetDoc.activeView;
            activeView.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
            if (itemWidth <= 0 || itemHeight <= 0) return true;

            /* 中心をそろえたあとの表示範囲を基準に倍率を求める / Scale from the view bounds after the center has moved */
            var fillRatio = (fitOptions.fillRatio > 0) ? fitOptions.fillRatio : DEFAULT_FIT_PERCENT / 100;
            var viewBounds = activeView.bounds;
            var scale = Math.min(
                (viewBounds[2] - viewBounds[0]) / itemWidth,
                (viewBounds[1] - viewBounds[3]) / itemHeight
            ) * fillRatio;
            activeView.zoom = clampNumber(activeView.zoom * scale, VIEW_ZOOM_RANGE, 1);
            return true;
        }

        /**
         * 現在の表示位置と倍率を控える（キャンセル時に restoreView で戻す）
         * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
         * @returns {{centerPoint: number[], zoom: number}} 控えた表示状態
         */
        function captureView(targetDoc) {
            var activeView = (targetDoc || app.activeDocument).activeView;
            return { centerPoint: activeView.centerPoint, zoom: activeView.zoom };
        }

        /**
         * captureView で控えた表示位置と倍率に戻す
         * @param {{centerPoint: number[], zoom: number}} viewState - 控えた表示状態
         * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
         * @returns {void}
         */
        function restoreView(viewState, targetDoc) {
            if (!viewState) return;
            var activeView = (targetDoc || app.activeDocument).activeView;
            activeView.centerPoint = viewState.centerPoint;
            activeView.zoom = viewState.zoom;
        }

        return {
            addControls: addControls,
            fit: fit,
            captureView: captureView,
            restoreView: restoreView,
            getItemsBounds: getItemsBounds
        };

    })();

    // 画面にフィット（再利用パーツ）ここまで / End of the reusable fit view

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
    };

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

    /**
     * pt の値を単位の値に換算する
     * @param {number} ptValue - pt の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} 単位の値
     */
    function ptToUnit(ptValue, unit) {
        return ptValue / unit.pointsPerUnit;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 日英のラベル文言。
     *
     * @typedef {Object} LabelEntry
     * @property {string} ja - 日本語の文言。
     * @property {string} en - 英語の文言。
     */

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
            title: { ja: "グリッドに整列配置", en: "Arrange in Grid" },
            presetSave: { ja: "マスクのプリセットを保存", en: "Save Mask Preset" },
            chooseImages: { ja: "配置する画像を選択", en: "Choose images to place" },
            chooseSource: { ja: "画像の選択", en: "Choose Images" },
            chooseFolder: { ja: "画像のあるフォルダーを選択", en: "Choose a folder of images" }
        },
        dropdown: {
            presetPlaceholder: { ja: "---", en: "---" },
            presetDefault: { ja: "デフォルト", en: "Default" },
            presetBlackThin: { ja: "黒 0.3pt・角丸なし", en: "Black 0.3 pt, square" },
            presetWhiteRound2: { ja: "白 1pt・角丸 2pt", en: "White 1 pt, radius 2 pt" },
            presetWhiteRound50: { ja: "白 1pt・角丸 50pt", en: "White 1 pt, radius 50 pt" }
        },
        panel: {
            placement: { ja: "配置先", en: "Placement Area" },
            division: { ja: "分割とマージン", en: "Divisions & Margin" },
            resize: { ja: "リサイズ", en: "Resize" },
            mask: { ja: "マスク", en: "Mask" },
            cellHandling: { ja: "セルの扱い", en: "Cell Handling" },
            arrangement: { ja: "配置と整列", en: "Placement & Arrangement" },
            itemAnchor: { ja: "アイテムの基準位置", en: "Object Anchor" }
        },
        fieldLabel: {
            presetName: { ja: "プリセット名", en: "Preset name" },
            rowCount: { ja: "行数", en: "Rows" },
            columnCount: { ja: "列数", en: "Columns" },
            gutter: { ja: "セル間隔", en: "Gutter" },
            margin: { ja: "マージン", en: "Margin" },
            cellColor: { ja: "カラー", en: "Color" },
            cellOpacity: { ja: "不透明度", en: "Opacity" },
            resizeRatio: { ja: "比率", en: "Ratio" },
            fillStart: { ja: "開始位置", en: "Start From" }
        },
        checkbox: {
            gutterPriority: { ja: "セル間隔を優先", en: "Prioritize Gutter" },
            cornerRadius: { ja: "角丸", en: "Round Corners" },
            frameStroke: { ja: "枠", en: "Frame" },
            spreadEvenly: { ja: "均等にまばらに", en: "Spread Evenly" },
            brick: { ja: "レンガ", en: "Brick" },
            keepOriginals: { ja: "元のオブジェクトを残す", en: "Keep Originals" }
        },
        radio: {
            targetArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            targetBackmost: { ja: "最背面のオブジェクト", en: "Backmost Object" },
            noCell: { ja: "残さない", en: "Don't Keep" },
            keepCell: { ja: "長方形で残す", en: "Keep as Rectangles" },
            toGuide: { ja: "ガイド化", en: "Convert to Guides" },
            toArtboard: { ja: "アートボード化", en: "Convert to Artboards" },
            blackCell: { ja: "黒", en: "Black" },
            whiteCell: { ja: "白", en: "White" },
            transparentCell: { ja: "塗りなし", en: "No Fill" },
            resizeNone: { ja: "なし", en: "None" },
            resizeWidth: { ja: "幅に合わせる", en: "Fit Width" },
            resizeHeight: { ja: "高さに合わせる", en: "Fit Height" },
            resizeArea: { ja: "面積をそろえる", en: "Equal Area" },
            resizeScale: { ja: "スケール", en: "Scale" },
            maskNone: { ja: "なし", en: "None" },
            maskCell: { ja: "セル", en: "Cell" },
            maskItem: { ja: "画像", en: "Image" }
        },
        button: {
            transparencyGrid: { ja: "透明グリッド表示", en: "Transparency Grid" },
            randomize: { ja: "シャッフル", en: "Shuffle" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" },
            chooseFiles: { ja: "ファイルを選択...", en: "Choose Files..." },
            chooseFolder: { ja: "フォルダーを選択...", en: "Choose Folder..." }
        },
        tooltip: {
            preset: { ja: "保存したマスクの設定（種類・角丸・枠）を読み込みます。", en: "Loads a saved mask setup (mode, round corners, frame)." },
            presetSave: { ja: "今のマスクの設定に名前を付けて保存する", en: "Save the current mask settings under a name" },
            presetDelete: { ja: "選んでいるプリセットを削除する", en: "Delete the selected preset" },
            presetName: { ja: "保存する設定の名前。同じ名前は上書きする", en: "Name to save the settings under. The same name is overwritten" },
            targetUnavailable: { ja: "該当するオブジェクトがありません。", en: "No matching object was found." },
            targetUnavailableForImages: {
                ja: "画像を読み込んで整列するときは、現在のアートボードに並べます。",
                en: "Imported images are arranged on the current artboard."
            },
            targetArtboard: {
                ja: "ダイアログボックスを開いたときに作業中だったアートボードを配置先にします。",
                en: "Uses the artboard that was active when the dialog box opened as the area."
            },
            targetBackmost: {
                ja: "選択したオブジェクトのうち、いちばん背面にあるものを配置先（枠）にします。そのオブジェクト自体は配置しません。",
                en: "Uses the backmost of the selected objects as the area (frame). That object itself is not placed."
            },
            noCell: {
                ja: "セルの長方形を残しません（プレビューではセルの範囲をガイドで示します）。",
                en: "Leaves no cell rectangles (the preview shows the cells as guides)."
            },
            keepCell: {
                ja: "セルの長方形を「cell-background」レイヤーに残します。",
                en: "Keeps the cell rectangles on the \"cell-background\" layer."
            },
            toGuide: { ja: "セルの長方形をガイドに変換します。", en: "Converts the cell rectangles into guides." },
            toArtboard: {
                ja: "セルごとにアートボードを作成し、長方形は残しません。",
                en: "Creates one artboard per cell and keeps no rectangles."
            },
            division: {
                ja: "セル数より多いオブジェクトは、アートボードの外へ退避します（［元のオブジェクトを残す］では退避しません）。",
                en: "Objects beyond the number of cells are parked outside the artboard (not with Keep Originals)."
            },
            gutter: { ja: "行数・列数のいずれかが2以上のときに有効です。", en: "Available when the rows or columns are 2 or more." },
            margin: {
                ja: "配置先の四辺から内側に取る余白です（［セル間隔を優先］では最低限の余白）。",
                en: "Inset taken from each edge of the placement area (the minimum with Prioritize Gutter)."
            },
            gutterPriority: {
                ja: "セルの大きさをオブジェクトのいちばん大きい幅・高さに合わせ、セル間隔をオブジェクトどうしのアキとして取ります。余った分はマージンに回し、グリッドを配置先の中央に置きます（マージンは指定値以上）。",
                en: "Sizes the cells to the widest and tallest objects so the gutter becomes the gap between objects. The leftover space goes to the margin and the grid is centered in the area (the margin is at least the value set)."
            },
            resizeNone: { ja: "大きさを変えずに配置します。", en: "Places the objects without resizing them." },
            resizeWidth: {
                ja: "縦横比を保ったまま、幅をセルの幅×比率にそろえます。",
                en: "Keeps the aspect ratio and matches each width to the cell width times the ratio."
            },
            resizeHeight: {
                ja: "縦横比を保ったまま、高さをセルの高さ×比率にそろえます。",
                en: "Keeps the aspect ratio and matches each height to the cell height times the ratio."
            },
            resizeArea: {
                ja: "縦横比を保ったまま、幅×高さ（面積）をそろえます。どのオブジェクトもセル×比率に収まる大きさになります。",
                en: "Keeps the aspect ratio and equalizes width × height (area), sized so every object fits in the cell times the ratio."
            },
            resizeScale: {
                ja: "元の大きさに比率を掛けて拡大・縮小します。",
                en: "Scales each object by the ratio of its original size."
            },
            resizeRatio: {
                ja: "幅に合わせる・高さに合わせる・面積をそろえるではセルに対する比率、スケールでは元の大きさに対する比率です。",
                en: "Relative to the cell for Fit Width, Fit Height, and Equal Area; relative to the original size for Scale."
            },
            maskNone: { ja: "マスクしません。", en: "No mask." },
            maskCell: {
                ja: "セルに入れたオブジェクトを、セルの長方形でクリッピングマスクします（はみ出した分は切り抜かれます）。",
                en: "Clips each object placed in a cell with the cell rectangle (whatever extends past the cell is cropped)."
            },
            maskItem: {
                ja: "オブジェクトと同じ大きさの長方形でクリッピングマスクします。切り抜かれる部分はなく、角丸や枠を付けたいときに使います。",
                en: "Clips each object with a rectangle of its own size. Nothing is cropped; use it to add round corners or a frame."
            },
            keepOriginals: {
                ja: "元のオブジェクトはその場に残し、複製をセルに配置します。",
                en: "Leaves the original objects in place and places copies in the cells."
            },
            fillStart: {
                ja: "セルを埋め始める位置です。四隅からは横方向に、上・下の中央からは各行を中央から外へ、左・右の中央からは縦方向に各列を中央から外へ埋めます。中央は中心に近いセルから埋めます。option＋クリックで［均等にまばらに］をオフにします。",
                en: "Where filling starts. Corners fill row by row; top and bottom centers fill each row from the middle out; left and right centers fill column by column from the middle out; Center fills the cells nearest the middle first. Option-click also turns off Spread Evenly."
            },
            cornerRadius: {
                ja: "マスクの角を、指定した半径で丸めます（効果の［角を丸くする］）。",
                en: "Rounds the corners of each mask by the radius (the Round Corners effect)."
            },
            frameStroke: {
                ja: "マスクの長方形に線を付けます。色見本をクリックすると枠の色を選べます。",
                en: "Adds a stroke to each mask rectangle. Click the swatch to choose the frame color."
            },
            frameColor: { ja: "クリックで枠の色を選ぶ", en: "Click to choose the frame color" },
            itemAnchor: {
                ja: "オブジェクトのどこをセルのどこに合わせるかです。マスクの［セル］では、はみ出した反対側を切り抜きます。",
                en: "Which point of each object meets the same point of its cell. With the Cell mask, whatever extends past the opposite side is cropped."
            },
            spreadEvenly: {
                ja: "オブジェクトがセルより少ないとき、詰めずにグリッド全体へ均等に散らします（開始位置から始めます）。セルが余らないときはディム表示になります。",
                en: "When there are fewer objects than cells, spreads them evenly over the grid instead of packing them, starting from the start position. Dimmed when there are no spare cells."
            },
            brick: {
                ja: "偶数行のセルを1つ減らし、半セルずらしてレンガ積みにします（列数が2以上のとき）。",
                en: "Removes one cell from every even row and shifts it by half a cell, like a brick wall (with 2 or more columns)."
            },
            autoCornerRadius: {
                ja: "自動計算：角丸の半径を、セルの短い辺の半分にします（短い辺の両端が半円になります）。マスクの［画像］では、いちばん小さいマスクの短い辺の半分にします。",
                en: "Auto: sets the corner radius to half the shorter side of a cell (the short ends become semicircles). With the Image mask, half the shorter side of the smallest mask."
            },
            autoDivision: {
                ja: "行数・列数を自動計算します。配置するオブジェクトの数と配置先の縦横比から、セルが正方形に近くなる行数・列数を入れます（［レンガ］も考慮）。配置先や［レンガ］を切り替えたときも計算し直します。",
                en: "Calculates rows and columns automatically: enters the ones that make the cells closest to square for the number of objects and the shape of the area (Brick included). Also recalculated when you switch the area or Brick."
            },
            autoScale: {
                ja: "自動計算（スケールのとき）：すべてのオブジェクトがセルを覆い、上下左右にアキが出ない共通のスケールを計算して比率に入れます。はみ出した分はマスクの［セル］で切り抜けます。",
                en: "Auto (with Scale): calculates one scale at which every object covers its cell with no gap on any side, and enters it as the ratio. Use the Cell mask to crop what extends past the cell."
            },
            randomize: {
                ja: "セルへの割り当て順をシャッフルします。［均等にまばらに］では、距離が同じセルどうしの選び方も選び直します（押すたびに変わります）。",
                en: "Shuffles the order in which objects fill the cells, and with Spread Evenly also re-picks among cells at the same distance (changes on every click)."
            },
            transparencyGrid: {
                ja: "透明グリッドの表示を切り替えます。スクリプト終了時に元へ戻します。",
                en: "Toggles the transparency grid. It is restored when the script finishes."
            },
            /* ステップボタン用 / for the stepper */
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
            invalidGrid: {
                ja: "この設定ではグリッドを作れません。行数・列数・セル間隔・マージンを見直してください。",
                en: "These settings cannot form a grid. Check the rows, columns, gutter, and margin."
            },
            artboardError: { ja: "アートボードを作成できませんでした。", en: "Could not create the artboards." },
            artboardCreated: { ja: "%1 個のアートボードを作成しました。", en: "%1 artboards created." },
            presetSaveFailed: { ja: "プリセットを保存できませんでした。", en: "Could not save the preset." },
            presetNameBuiltin: { ja: "「{name}」は組み込みのプリセットの名前なので使えません。", en: "\"{name}\" is the name of a built-in preset and cannot be used." },
            noImagesInFolder: { ja: "「{name}」に配置できる画像がありません。", en: "\"{name}\" contains no images to place." },
            layerUnavailable: {
                ja: "作業中のレイヤー「{name}」がロックまたは非表示のため配置できません。",
                en: "Cannot place images because the active layer \"{name}\" is locked or hidden."
            },
            placeFailed: {
                ja: "次の画像を配置できませんでした。\n{list}",
                en: "The following images could not be placed.\n{list}"
            }
        },
        confirm: {
            presetOverwrite: { ja: "「{name}」を上書きしますか？", en: "Overwrite \"{name}\"?" },
            presetDelete: { ja: "「{name}」を削除しますか？", en: "Delete \"{name}\"?" }
        },
        message: {
            chooseSource: { ja: "配置する画像を、ファイルかフォルダーで選びます。", en: "Choose the images to place as files or as a folder." }
        }
    };

    // =========================================
    // 退避 / Parking
    // =========================================

    /**
     * 2つの矩形が重なるかを判定します（余裕を加味）。/ Test whether two rects overlap.
     *
     * @param {number[]} rectA - 矩形A [左, 上, 右, 下]。
     * @param {number[]} rectB - 矩形B [左, 上, 右, 下]。
     * @param {number} tolerance - 重なりと見なす余裕（pt）。
     * @returns {boolean} 重なっている場合は true。
     */
    function rectsOverlap(rectA, rectB, tolerance) {
        return !(rectA[2] < rectB[0] - tolerance || rectA[0] > rectB[2] + tolerance ||
            rectA[1] < rectB[3] - tolerance || rectA[3] > rectB[1] + tolerance);
    }

    /**
     * 対象領域や他アートボードと重ならない位置へオブジェクトを退避します。
     * / Park objects clear of the target area and other artboards.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {Array<PageItem>} pageItems - 退避するオブジェクト。
     * @param {number[]} avoidRect - 避けたい配置先の矩形 [左, 上, 右, 下]。
     * @returns {void}
     */
    function parkItemsOutside(doc, pageItems, avoidRect) {
        var activeIndex = doc.artboards.getActiveArtboardIndex();
        var activeRect = doc.artboards[activeIndex].artboardRect;

        /* アクティブアートボードと対象領域の和を「避けたい領域」とする / Avoid the union of the active artboard and the target area */
        var avoidArea = [
            Math.min(activeRect[0], avoidRect[0]),
            Math.max(activeRect[1], avoidRect[1]),
            Math.max(activeRect[2], avoidRect[2]),
            Math.min(activeRect[3], avoidRect[3])
        ];

        var parkedRects = [];
        var artboardTolerance = CONFIG.overlapTolerance + CONFIG.artboardBuffer;

        /**
         * 退避先が既存の退避先や他アートボードと衝突するかを判定します。
         * / Test a parking slot against parked items and artboards.
         *
         * @param {number[]} candidateRect - 退避先の候補矩形。
         * @returns {boolean} 使用できない場合は true。
         */
        function isSlotTaken(candidateRect) {
            for (var i = 0; i < parkedRects.length; i++) {
                if (rectsOverlap(candidateRect, parkedRects[i], CONFIG.overlapTolerance)) return true;
            }
            for (var j = 0; j < doc.artboards.length; j++) {
                if (j === activeIndex) continue;
                if (rectsOverlap(candidateRect, doc.artboards[j].artboardRect, artboardTolerance)) return true;
            }
            return false;
        }

        for (var i = 0; i < pageItems.length; i++) {
            var bounds = pageItems[i].visibleBounds;
            if (!rectsOverlap(bounds, avoidArea, 0)) continue;

            var dx = (avoidArea[2] - bounds[0]) + CONFIG.parkingStep;
            var slotRect;
            do {
                slotRect = [bounds[0] + dx, bounds[1], bounds[2] + dx, bounds[3]];
                if (!isSlotTaken(slotRect)) break;
                dx += CONFIG.parkingStep;
            } while (true);

            pageItems[i].translate(dx, 0);
            parkedRects.push(slotRect);
        }
    }

    // =========================================
    // 対象の検出 / Target detection
    // =========================================

    /**
     * 選択の中でいちばん背面のオブジェクトを返します（枠として配置先に使う）。/ Return the backmost object in the selection.
     * doc.selection は前面から背面の順に並ぶので、最後の要素が最背面です。
     * 枠を除いて整列するオブジェクトが残るよう、2つ以上選択しているときだけ返します。
     *
     * @param {Array<PageItem>} selectionItems - 選択中のオブジェクト（前面→背面の順）。
     * @returns {PageItem|null} 最背面のオブジェクト。1つ以下なら null。
     */
    function findBackmostSelectedItem(selectionItems) {
        if (!selectionItems || selectionItems.length < 2) return null;
        return selectionItems[selectionItems.length - 1];
    }

    /**
     * 指定のアイテムを除いた配列の写しを返します（順序は保持）。/ Copy the items without one item.
     * 除くものが無ければ null を渡し、選択（doc.selection）を配列に写すのにも使います。
     *
     * @param {Array<PageItem>|null} selectionItems - 元の配列（選択など）。
     * @param {PageItem|null} excludedItem - 除くアイテム。無ければ null。
     * @returns {Array<PageItem>} 残ったオブジェクト。
     */
    function copyItemsExcept(selectionItems, excludedItem) {
        var remainingItems = [];
        for (var i = 0; selectionItems && i < selectionItems.length; i++) {
            if (selectionItems[i] !== excludedItem) remainingItems.push(selectionItems[i]);
        }
        return remainingItems;
    }

    /**
     * セルが正方形に近くなる行数・列数を求めます。/ Find rows and columns that make cells nearly square.
     * 余白セルが行／列1本分以上になる構成は除外し、その範囲でセル比1に最も近いものを選びます。
     * ［レンガ］では偶数行が1つ少ない分を数え、足りるまで行を増やします。
     *
     * @param {number} itemCount - 配置するオブジェクトの数。
     * @param {number[]} targetRect - 配置先の矩形 [左, 上, 右, 下]。
     * @param {boolean} isBrick - ［レンガ］なら true。
     * @returns {{rowCount: number, columnCount: number}} 行数と列数。
     */
    function computeDefaultDivision(itemCount, targetRect, isBrick) {
        var fallbackDivision = { rowCount: CONFIG.fallbackDivision, columnCount: CONFIG.fallbackDivision };
        if (!itemCount || itemCount <= 0) return fallbackDivision;
        if (itemCount === 1) return { rowCount: 1, columnCount: 1 };

        var areaWidth = targetRect[2] - targetRect[0];
        var areaHeight = targetRect[1] - targetRect[3];
        if (areaWidth <= 0 || areaHeight <= 0) return fallbackDivision;

        /* セル総数の上限を超える分はどのみち溢れるので、上限の範囲で割り付ける。
           ［レンガ］で上限内に収まる構成が無いときは、入れる数を減らして探し直す（溢れた分は退避される）
           Anything beyond the cell cap overflows anyway; with Brick, shrink the count until a layout fits */
        for (var cellBudget = Math.min(itemCount, CONFIG.maxCellCount); cellBudget >= 1; cellBudget--) {
            var division = findSquarestDivision(cellBudget, areaWidth, areaHeight, isBrick);
            if (division) return division;
        }
        return fallbackDivision;
    }

    /**
     * 指定したセル数が入る行数・列数のうち、セルがいちばん正方形に近いものを返します。/ Find the squarest division for a cell count.
     *
     * @param {number} cellBudget - 入れるセルの数。
     * @param {number} areaWidth - 配置先の幅。
     * @param {number} areaHeight - 配置先の高さ。
     * @param {boolean} isBrick - ［レンガ］なら true。
     * @returns {{rowCount: number, columnCount: number}|null} 行数と列数。条件に合う構成が無ければ null。
     */
    function findSquarestDivision(cellBudget, areaWidth, areaHeight, isBrick) {
        var bestDivision = null;
        var bestAspectRatio = Infinity;
        for (var columnCount = 1; columnCount <= cellBudget; columnCount++) {
            var isBrickGrid = isBrick && columnCount >= 2; /* computeGridMetrics() と同じく1列ではレンガにしない / no brick with one column */
            var rowCount = Math.ceil(cellBudget / columnCount);
            while (countDivisionCells(rowCount, columnCount, isBrickGrid) < cellBudget) rowCount++;
            /* 余るセルが、いちばん短い行（列）1本分に届く構成は除く / Skip layouts that leave a whole shortest line empty */
            var emptyCellCount = countDivisionCells(rowCount, columnCount, isBrickGrid) - cellBudget;
            var shortestLine = Math.min(rowCount, isBrickGrid ? columnCount - 1 : columnCount);
            if (emptyCellCount > 0 && emptyCellCount >= shortestLine) continue;

            /* computeGridMetrics() が拒む組み合わせは初期値にしない / Never default to a grid the metrics reject */
            if (rowCount > CONFIG.maxDivision || columnCount > CONFIG.maxDivision) continue;
            if (rowCount * columnCount > CONFIG.maxCellCount) continue;

            var cellWidth = areaWidth / columnCount;
            var cellHeight = areaHeight / rowCount;
            var aspectRatio = (cellWidth > cellHeight) ? (cellWidth / cellHeight) : (cellHeight / cellWidth);
            if (aspectRatio < bestAspectRatio) {
                bestAspectRatio = aspectRatio;
                bestDivision = { rowCount: rowCount, columnCount: columnCount };
            }
        }
        return bestDivision;
    }

    // =========================================
    // レイヤー操作 / Layer helpers
    // =========================================

    /**
     * 指定名のレイヤーを探します。/ Find the layer with the given name.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - 探すレイヤー名。
     * @returns {Layer|null} 見つかったレイヤー。無ければ null。
     */
    function findLayerByName(doc, layerName) {
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            /* getByName は見つからないと例外 / getByName throws when the layer is missing */
            return null;
        }
    }

    /**
     * 指定名のレイヤーがあれば削除します。/ Remove the layer with the given name if present.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - 削除するレイヤー名。
     * @returns {void}
     */
    function removeLayerByName(doc, layerName) {
        var foundLayer = findLayerByName(doc, layerName);
        if (!foundLayer) return;
        try {
            foundLayer.locked = false;
            foundLayer.remove();
        } catch (e) {
            /* ドキュメントに残る最後のレイヤーなどは削除できない。そのときは残す / The last layer of a document cannot be removed; leave it */
        }
    }

    /**
     * レイヤーを取得し、無ければ作成します。/ Get the layer, creating it when missing.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - レイヤー名。
     * @returns {Layer} 取得または作成したレイヤー（ロック解除・表示済み）。
     */
    function getOrCreateLayer(doc, layerName) {
        var foundLayer = findLayerByName(doc, layerName);
        if (!foundLayer) {
            foundLayer = doc.layers.add();
            foundLayer.name = layerName;
        }
        foundLayer.locked = false;
        /* これから描き込むレイヤーは必ず表示にする / A layer we are about to draw into must be visible */
        foundLayer.visible = true;
        return foundLayer;
    }

    /**
     * ロック中でも失敗しないようにレイヤーの表示状態を切り替えます。/ Toggle a layer without failing on locks.
     *
     * @param {Layer|null} layer - 対象のレイヤー。null なら何もしません。
     * @param {boolean} visible - 表示する場合は true。
     * @returns {void}
     */
    function setLayerVisible(layer, visible) {
        if (!layer) return;
        try {
            layer.visible = visible;
        } catch (e) {
            /* 親レイヤーがロックされていると例外になる。表示は変えずに続ける / Throws under a locked parent layer; keep going */
        }
    }

    /**
     * ロック中でも失敗しないように表示状態を切り替えます。/ Toggle visibility without failing on locked items.
     *
     * @param {PageItem} pageItem - 対象のオブジェクト。
     * @param {boolean} hidden - 非表示にする場合は true。
     * @returns {void}
     */
    function setItemHidden(pageItem, hidden) {
        try {
            pageItem.hidden = hidden;
        } catch (e) {
            /* ロック中のオブジェクトやレイヤーでは例外になる。表示は変えずに続ける / Throws on locked items or layers; keep going */
        }
    }

    /**
     * セル描画レイヤーの長方形を選択します。/ Select the rectangles on the cell layer.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {boolean} 選択できた場合は true。レイヤーが無い場合は false。
     */
    function selectCellRectangles(doc) {
        var cellLayer = findLayerByName(doc, CONFIG.cellLayerName);
        if (!cellLayer) return false;
        var cellItems = [];
        for (var i = 0; i < cellLayer.pathItems.length; i++) cellItems.push(cellLayer.pathItems[i]);
        doc.selection = cellItems;
        return true;
    }

    // =========================================
    // グリッド / Grid
    // =========================================

    /**
     * グリッドの寸法（すべて pt、矩形は [左, 上, 右, 下]）。
     *
     * @typedef {Object} GridMetrics
     * @property {number} rowCount - 行数。
     * @property {number} columnCount - 列数（［レンガ］では奇数行の数。偶数行は1つ少ない）。
     * @property {boolean} brick - ［レンガ］なら true（偶数行を1つ減らし、半セル右へずらす）。
     * @property {number} originLeft - 1行1列目のセル左端。
     * @property {number} originTop - 1行1列目のセル上端。
     * @property {number} cellWidth - セルの幅。
     * @property {number} cellHeight - セルの高さ。
     * @property {number} gutter - セル間の間隔。
     * @property {number[]} targetRect - 配置先の矩形。
     */

    /**
     * 行数・列数・余白からグリッドの寸法を計算します。/ Compute the grid metrics.
     *
     * @param {number|null} rowCount - 行数。
     * @param {number|null} columnCount - 列数。
     * @param {number} margin - マージン（pt）。
     * @param {number} gutter - セル間隔（pt）。
     * @param {number[]} targetRect - 配置先の矩形 [左, 上, 右, 下]。
     * @param {boolean} isBrick - ［レンガ］なら true。列数が1のときは偶数行が空になるので無視します。
     * @returns {GridMetrics|null} グリッドの寸法。入力が無効な場合は null。
     */
    function computeGridMetrics(rowCount, columnCount, margin, gutter, targetRect, isBrick) {
        if (rowCount === null || columnCount === null) return null;
        if (rowCount < 1 || columnCount < 1) return null;
        if (rowCount > CONFIG.maxDivision || columnCount > CONFIG.maxDivision) return null;
        if (rowCount * columnCount > CONFIG.maxCellCount) return null;

        var usableWidth = (targetRect[2] - margin) - (targetRect[0] + margin);
        var usableHeight = (targetRect[1] - margin) - (targetRect[3] + margin);
        var cellWidth = (usableWidth - (columnCount - 1) * gutter) / columnCount;
        var cellHeight = (usableHeight - (rowCount - 1) * gutter) / rowCount;

        /* マージンや間隔が大きすぎてセルが成立しない場合は描画しない / No grid when margins or gutters leave no room */
        if (!isFinite(cellWidth) || !isFinite(cellHeight) || cellWidth <= 0 || cellHeight <= 0) return null;

        return {
            rowCount: rowCount,
            columnCount: columnCount,
            brick: !!isBrick && columnCount >= 2,
            originLeft: targetRect[0] + margin,
            originTop: targetRect[1] - margin,
            cellWidth: cellWidth,
            cellHeight: cellHeight,
            gutter: gutter,
            targetRect: targetRect
        };
    }

    /**
     * 全セルを行→列の順に走査します。/ Iterate cells in row-major order.
     *
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {function(number[]): (boolean|void)} handleCell - セル矩形を受け取る処理。false を返すと走査を中断します。
     * @returns {void}
     */
    function forEachCell(gridMetrics, handleCell) {
        for (var i = 0; i < gridMetrics.rowCount; i++) {
            var rowCellCount = countRowCells(gridMetrics, i);
            for (var j = 0; j < rowCellCount; j++) {
                if (handleCell(getCellRect(gridMetrics, i, j)) === false) return;
            }
        }
    }

    /**
     * 行のセル数を返します（［レンガ］では偶数行が1つ少ない）。/ Return the number of cells in a row.
     *
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {number} rowIndex - 行の番号（0 が上。1, 3, … が偶数行）。
     * @returns {number} その行のセル数。
     */
    function countRowCells(gridMetrics, rowIndex) {
        return isShortBrickRow(gridMetrics, rowIndex) ? gridMetrics.columnCount - 1 : gridMetrics.columnCount;
    }

    /**
     * ［レンガ］で1つ少なく半セルずれる行（偶数行）かを返します。/ Whether the row is a shortened brick row.
     *
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {number} rowIndex - 行の番号（0 が上）。
     * @returns {boolean} 偶数行で［レンガ］なら true。
     */
    function isShortBrickRow(gridMetrics, rowIndex) {
        return gridMetrics.brick && rowIndex % 2 === 1;
    }

    /**
     * グリッド全体のセル数を返します。/ Return the total number of cells.
     *
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @returns {number} セル数。
     */
    function countCells(gridMetrics) {
        return countDivisionCells(gridMetrics.rowCount, gridMetrics.columnCount, gridMetrics.brick);
    }

    /**
     * 行数・列数からセル数を返します（［レンガ］では偶数行が1つ少ない）。/ Count cells for rows and columns.
     *
     * @param {number} rowCount - 行数。
     * @param {number} columnCount - 列数（奇数行の数）。
     * @param {boolean} isBrick - ［レンガ］なら true（列数2以上のときだけ渡す）。
     * @returns {number} セル数。
     */
    function countDivisionCells(rowCount, columnCount, isBrick) {
        return rowCount * columnCount - (isBrick ? Math.floor(rowCount / 2) : 0);
    }

    /**
     * 行・列の番号からセルの矩形を求めます。/ Return the rect of one cell.
     *
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {number} rowIndex - 行の番号（0 が上）。
     * @param {number} columnIndex - 列の番号（0 が左）。［レンガ］の偶数行は半セル右へずらします。
     * @returns {number[]} セルの矩形 [左, 上, 右, 下]。
     */
    function getCellRect(gridMetrics, rowIndex, columnIndex) {
        var columnOffset = isShortBrickRow(gridMetrics, rowIndex) ? 0.5 : 0;
        var cellTop = gridMetrics.originTop - (gridMetrics.cellHeight + gridMetrics.gutter) * rowIndex;
        var cellLeft = gridMetrics.originLeft + (gridMetrics.cellWidth + gridMetrics.gutter) * (columnIndex + columnOffset);
        return [cellLeft, cellTop, cellLeft + gridMetrics.cellWidth, cellTop - gridMetrics.cellHeight];
    }

    /**
     * ［セル間隔を優先］のグリッドに作り直します。/ Rebuild the grid so the gutter is the gap between objects.
     * セルの大きさをオブジェクトのいちばん大きい幅・高さに合わせ、セル間隔はそのまま、グリッド全体を配置先の中央に置きます。
     * セルは元のグリッドより大きくしないので、マージンは指定値以上になります。
     *
     * @param {GridMetrics} gridMetrics - マージン・セル間隔どおりに作ったグリッド。
     * @param {Array<PageItem>} pageItems - セルに入れるオブジェクト（リサイズ後）。
     * @returns {GridMetrics} 作り直したグリッド（測れるオブジェクトが無ければ元のまま）。
     */
    function fitGridToItems(gridMetrics, pageItems) {
        var maxWidth = 0;
        var maxHeight = 0;
        for (var i = 0; i < pageItems.length; i++) {
            var bounds = pageItems[i].visibleBounds;
            maxWidth = Math.max(maxWidth, bounds[2] - bounds[0]);
            maxHeight = Math.max(maxHeight, bounds[1] - bounds[3]);
        }
        if (maxWidth <= 0 || maxHeight <= 0) return gridMetrics;

        var cellWidth = Math.min(maxWidth, gridMetrics.cellWidth);
        var cellHeight = Math.min(maxHeight, gridMetrics.cellHeight);
        var blockWidth = cellWidth * gridMetrics.columnCount + gridMetrics.gutter * (gridMetrics.columnCount - 1);
        var blockHeight = cellHeight * gridMetrics.rowCount + gridMetrics.gutter * (gridMetrics.rowCount - 1);
        var targetRect = gridMetrics.targetRect;
        return {
            rowCount: gridMetrics.rowCount,
            columnCount: gridMetrics.columnCount,
            brick: gridMetrics.brick,
            originLeft: (targetRect[0] + targetRect[2]) / 2 - blockWidth / 2,
            originTop: (targetRect[1] + targetRect[3]) / 2 + blockHeight / 2,
            cellWidth: cellWidth,
            cellHeight: cellHeight,
            gutter: gridMetrics.gutter,
            targetRect: targetRect
        };
    }

    /**
     * セルの矩形を、オブジェクトを入れる順に並べて返します。/ List the cell rects in fill order.
     * 四隅からは、その角の行から横方向に埋めます。上・下の中央からは、その辺の行から、各行を中央の列から外へ埋めます。
     * 左・右の中央からは、その辺の列から、各列を中央の行から外へ埋めます。中央は中心に近いセルから埋めます。
     * ［レンガ］では、行ごとのセル数で並べ、縦方向に埋めるときは半セル単位の横位置で列を数えます。
     *
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {number} fillStartIndex - 埋め始める位置（基準点ウィジェットの番号。0=左上, 1=上, 2=右上, 3=左, 4=中央, 5=右, 6=左下, 7=下, 8=右下）。
     * @returns {Array<number[]>} セルの矩形 [左, 上, 右, 下] の配列。
     */
    function listCellRectsInFillOrder(gridMetrics, fillStartIndex) {
        var rowCount = gridMetrics.rowCount;
        var anchorRow = Math.floor(fillStartIndex / 3);    /* 0=上, 1=中, 2=下 / 0 top, 1 middle, 2 bottom */
        var anchorColumn = fillStartIndex % 3;            /* 0=左, 1=中, 2=右 / 0 left, 1 middle, 2 right */
        var rowOrder = listLineOrder(rowCount, anchorRow);

        var cellPositions = [];
        var i, j;
        if (anchorRow === 1 && anchorColumn !== 1) {
            /* 左・右の中央は、列ごとに中央の行から外へ（列優先）。［レンガ］の偶数行は奇数行の列のあいだに入る
               Left and right centers: column by column, middle row outward; brick rows sit between the columns */
            var slotStep = gridMetrics.brick ? 2 : 1;
            var slotOrder = listLineOrder((gridMetrics.columnCount - 1) * slotStep + 1, anchorColumn);
            for (j = 0; j < slotOrder.length; j++) {
                for (i = 0; i < rowCount; i++) {
                    var rowIndex = rowOrder[i];
                    var rowSlot = slotOrder[j] - (isShortBrickRow(gridMetrics, rowIndex) ? 1 : 0);
                    if (rowSlot < 0 || rowSlot % slotStep !== 0) continue;
                    var columnIndex = rowSlot / slotStep;
                    if (columnIndex < countRowCells(gridMetrics, rowIndex)) cellPositions.push([rowIndex, columnIndex]);
                }
            }
        } else if (fillStartIndex === 4) {
            /* 中央は左上からの行優先に並べてから、中心に近い順へ（同じ距離は左上側が先）/ Center: row-major from the top left, then nearest the middle first */
            for (i = 0; i < rowCount; i++) {
                for (j = 0; j < countRowCells(gridMetrics, i); j++) cellPositions.push([i, j]);
            }
            cellPositions = orderCellsFromCenter(cellPositions, gridMetrics);
        } else {
            /* 四隅は、その角の行から横方向に。上・下の中央は、各行を中央の列から外へ（どちらも行優先）
               Corners: row by row from that corner; top and bottom centers: each row from the middle out */
            for (i = 0; i < rowCount; i++) {
                var columnOrder = listLineOrder(countRowCells(gridMetrics, rowOrder[i]), anchorColumn);
                for (j = 0; j < columnOrder.length; j++) cellPositions.push([rowOrder[i], columnOrder[j]]);
            }
        }

        var cellRects = [];
        for (var k = 0; k < cellPositions.length; k++) {
            cellRects.push(getCellRect(gridMetrics, cellPositions[k][0], cellPositions[k][1]));
        }
        return cellRects;
    }

    /**
     * 行（列）を埋める順に並べた番号を返します。/ Return line indexes in fill order.
     *
     * @param {number} lineCount - 行数（列数）。
     * @param {number} anchorSide - 0 なら先頭から、2 なら末尾から、1 なら中央から外へ（同じ距離なら先頭側が先）。
     * @returns {number[]} 行（列）の番号の配列。
     */
    function listLineOrder(lineCount, anchorSide) {
        var lineOrder = [];
        var i;
        if (anchorSide !== 1) {
            for (i = 0; i < lineCount; i++) lineOrder.push(anchorSide === 2 ? lineCount - 1 - i : i);
            return lineOrder;
        }
        /* 中央から外へ：中央の左（上）・右（下）を交互に足す / Middle outward, alternating toward the start and end */
        var lowIndex = Math.floor((lineCount - 1) / 2);
        var highIndex = lowIndex + 1;
        if (lineCount % 2 === 1) {
            lineOrder.push(lowIndex);
            lowIndex--;
        }
        while (lowIndex >= 0 || highIndex < lineCount) {
            if (lowIndex >= 0) lineOrder.push(lowIndex--);
            if (highIndex < lineCount) lineOrder.push(highIndex++);
        }
        return lineOrder;
    }

    /**
     * セルを、グリッドの中心に近い順に並べ直します（同じ距離どうしは元の順）。/ Order cells from the center outward.
     * 距離は行・列の番号で測ります（半セル単位の整数にして、比較関数つきの sort() を使わずに束ねます）。
     * ［レンガ］の偶数行は、半セルずらした位置で測ります。
     *
     * @param {Array<number[]>} cellPositions - [行の番号, 列の番号] の配列（左上から行優先）。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @returns {Array<number[]>} 並べ直した [行の番号, 列の番号] の配列。
     */
    function orderCellsFromCenter(cellPositions, gridMetrics) {
        var rowCount = gridMetrics.rowCount;
        var columnCount = gridMetrics.columnCount;
        var cellsByDistance = {};
        var distanceKeys = [];
        for (var i = 0; i < cellPositions.length; i++) {
            /* 中心からのずれを2倍して整数にする / Double the offsets so they stay integers */
            var dy = 2 * cellPositions[i][0] - (rowCount - 1);
            var dx = 2 * cellPositions[i][1] + (isShortBrickRow(gridMetrics, cellPositions[i][0]) ? 1 : 0) - (columnCount - 1);
            var distanceKey = dx * dx + dy * dy;
            if (!cellsByDistance.hasOwnProperty(distanceKey)) {
                cellsByDistance[distanceKey] = [];
                insertSortedNumber(distanceKeys, distanceKey);
            }
            cellsByDistance[distanceKey].push(cellPositions[i]);
        }

        var orderedPositions = [];
        for (var j = 0; j < distanceKeys.length; j++) {
            var sameDistanceCells = cellsByDistance[distanceKeys[j]];
            for (var k = 0; k < sameDistanceCells.length; k++) orderedPositions.push(sameDistanceCells[k]);
        }
        return orderedPositions;
    }

    /**
     * 使うセルを、グリッド全体に均等に散らして選びます（最遠点の順に選ぶ）。/ Pick cells spread evenly over the grid.
     * 埋める順の先頭のセルから始め、選んだセルからいちばん離れたセルを次々に選びます（同じ距離なら先に調べたもの）。
     * 乱数を渡すと、同じ距離どうしの選び方を乱数で決めます（シャッフル用）。
     * 選んだセルは埋める順のまま返すので、オブジェクトの並びは開始位置に従います。
     *
     * @param {Array<number[]>} cellRects - セルの矩形（埋める順）。
     * @param {number} pickCount - 選ぶ数（セル数より少ない）。
     * @param {function(): number|null} randomFn - 0 以上 1 未満を返す乱数。null なら決まった選び方。
     * @returns {Array<number[]>} 選んだセルの矩形（埋める順）。
     */
    function pickSpreadCells(cellRects, pickCount, randomFn) {
        if (pickCount <= 0) return [];
        var cellCount = cellRects.length;
        /* 調べる順。乱数があれば並べ替えて、同じ距離のときに選ばれるセルを散らす / Visiting order; shuffled to randomize ties */
        var visitOrder = [];
        for (var cellIndex = 0; cellIndex < cellCount; cellIndex++) visitOrder.push(cellIndex);
        if (randomFn) visitOrder = createShuffledCopy(visitOrder, randomFn);
        var centerXs = [];
        var centerYs = [];
        var nearestDistances = [];
        var isPicked = [];
        for (var i = 0; i < cellCount; i++) {
            centerXs.push((cellRects[i][0] + cellRects[i][2]) / 2);
            centerYs.push((cellRects[i][1] + cellRects[i][3]) / 2);
            nearestDistances.push(Infinity);
            isPicked.push(false);
        }

        /* 始めるセルは埋める順の先頭（開始位置）。乱数は同じ距離どうしの選び方だけに使う
           Start at the first cell in fill order (Start From); randomness only breaks ties */
        var pickIndex = 0;
        for (var pickedCount = 0; pickedCount < pickCount; pickedCount++) {
            isPicked[pickIndex] = true;
            /* 選んだセルまでの距離で、各セルの「最寄りの選択済みセルまでの距離」を更新しつつ、次に選ぶセルを探す
               Update each cell's distance to its nearest picked cell, and find the farthest one */
            var farthestIndex = -1;
            var farthestDistance = -1;
            for (var n = 0; n < cellCount; n++) {
                var j = visitOrder[n];
                if (isPicked[j]) continue;
                var dx = centerXs[j] - centerXs[pickIndex];
                var dy = centerYs[j] - centerYs[pickIndex];
                nearestDistances[j] = Math.min(nearestDistances[j], dx * dx + dy * dy);
                if (nearestDistances[j] > farthestDistance) {
                    farthestDistance = nearestDistances[j];
                    farthestIndex = j;
                }
            }
            if (farthestIndex < 0) break;
            pickIndex = farthestIndex;
        }

        var pickedRects = [];
        for (var k = 0; k < cellCount; k++) {
            if (isPicked[k]) pickedRects.push(cellRects[k]);
        }
        return pickedRects;
    }

    /**
     * 種から同じ並びを返す乱数を作ります（線形合同法）。/ Create a seeded random generator (LCG).
     * プレビューを描き直しても、シャッフルするまでは同じ結果にするために使います。
     *
     * @param {number} seed - 種（0 以上 2^32 未満の整数）。
     * @returns {function(): number} 0 以上 1 未満を返す関数。
     */
    function createSeededRandom(seed) {
        var randomState = seed;
        return function () {
            randomState = (randomState * 1664525 + 1013904223) % 4294967296;
            return randomState / 4294967296;
        };
    }

    /**
     * 昇順の配列に、順序を保って数値を差し込みます。/ Insert a number into an ascending array.
     *
     * @param {number[]} sortedNumbers - 昇順の配列（書き換える）。
     * @param {number} newNumber - 差し込む数値。
     * @returns {void}
     */
    function insertSortedNumber(sortedNumbers, newNumber) {
        var insertAt = sortedNumbers.length;
        while (insertAt > 0 && sortedNumbers[insertAt - 1] > newNumber) insertAt--;
        sortedNumbers.splice(insertAt, 0, newNumber);
    }

    /**
     * 全セルの長方形を描画します。/ Draw rectangles for every cell.
     *
     * @param {Layer} cellLayer - 描画先のレイヤー。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {boolean} asGuide - ガイドに変換する場合は true。
     * @param {CMYKColor|RGBColor|null} fillColor - セルの塗り色。塗りなしは null。
     * @param {number|null} opacity - セルの不透明度（%）。指定しない場合は null。
     * @returns {void}
     */
    function drawCells(cellLayer, gridMetrics, asGuide, fillColor, opacity) {
        /* 省略された場合も「塗りなし」として扱う / Treat a missing argument as no fill */
        var hasFill = (fillColor !== null && fillColor !== undefined);

        forEachCell(gridMetrics, function (cellRect) {
            var cellRectangle = cellLayer.pathItems.rectangle(
                cellRect[1], cellRect[0], gridMetrics.cellWidth, gridMetrics.cellHeight);
            cellRectangle.stroked = false;
            cellRectangle.filled = hasFill;
            if (hasFill) cellRectangle.fillColor = fillColor;
            if (opacity !== null) cellRectangle.opacity = opacity;
            if (asGuide) cellRectangle.guides = true;
        });
        cellLayer.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 各セルへオブジェクトを1つずつ配置します（オブジェクトの基準点をセルの同じ基準点に合わせる）。/ Place one object in each cell.
     *
     * @param {Array<PageItem>} pageItems - 配置するオブジェクト（配置順）。
     * @param {Array<number[]>} cellRects - セルの矩形（埋める順。listCellRectsInFillOrder() の戻り値）。
     * @param {number} anchorIndex - 合わせる基準点（基準点ウィジェットの番号。4 が中央）。
     * @returns {void}
     */
    function placeItemsInCells(pageItems, cellRects, anchorIndex) {
        var placeCount = Math.min(pageItems.length, cellRects.length);
        for (var i = 0; i < placeCount; i++) {
            var cellPoint = getAnchorPointOnBounds(cellRects[i], anchorIndex);
            var itemPoint = getAnchorPointOnBounds(pageItems[i].visibleBounds, anchorIndex);
            pageItems[i].translate(cellPoint[0] - itemPoint[0], cellPoint[1] - itemPoint[1]);
        }
    }

    /* マスクの種類 / Mask modes */
    var MASK_MODE_NONE = "none"; /* マスクしない / no mask */
    var MASK_MODE_CELL = "cell"; /* セルの長方形 / the cell rectangle */
    var MASK_MODE_ITEM = "item"; /* オブジェクトと同じ大きさの長方形 / a rectangle of the object's own size */

    /**
     * マスクに使う矩形を返します。/ Return the rectangles used as masks.
     * ［セル］ではセルの矩形、［画像］では配置後のオブジェクトの visibleBounds（切り抜かれる部分はない）。
     *
     * @param {Array<PageItem>} pageItems - セルに入れたオブジェクト（配置順。配置とリサイズのあと）。
     * @param {Array<number[]>} cellRects - セルの矩形（埋める順。placeItemsInCells() に渡したもの）。
     * @param {string} maskMode - MASK_MODE_CELL / MASK_MODE_ITEM。
     * @returns {Array<number[]>} マスクの矩形 [左, 上, 右, 下] の配列。
     */
    function listMaskRects(pageItems, cellRects, maskMode) {
        var maskCount = Math.min(pageItems.length, cellRects.length);
        var maskRects = [];
        for (var i = 0; i < maskCount; i++) {
            maskRects.push((maskMode === MASK_MODE_ITEM) ? pageItems[i].visibleBounds : cellRects[i]);
        }
        return maskRects;
    }

    /**
     * セルに入れたオブジェクトを、矩形でクリッピングマスクします。/ Clip each placed object with its rectangle.
     * オブジェクトのあった場所にグループを作り、オブジェクトとマスクの長方形（最前面）を入れて clipped にします。
     *
     * @param {Array<PageItem>} pageItems - セルに入れたオブジェクト（配置順）。
     * @param {Array<number[]>} maskRects - マスクの矩形（listMaskRects() の戻り値。pageItems と同じ順）。
     * @param {Array<GroupItem>} clipGroups - 作ったクリップグループを1つずつ足す配列（途中で例外になっても、作った分は残る）。
     * @param {MaskStyle} maskStyle - 角丸と枠の指定。
     * @returns {void}
     */
    function maskItemsToRects(pageItems, maskRects, clipGroups, maskStyle) {
        for (var i = 0; i < maskRects.length; i++) {
            clipGroups.push(clipItemToRect(pageItems[i], maskRects[i], maskStyle));
        }
    }

    /**
     * マスクの角丸と枠の指定。
     *
     * @typedef {Object} MaskStyle
     * @property {number} cornerRadius - 角丸の半径（pt）。0 なら角丸にしない。
     * @property {number} frameWidth - 枠の線幅（pt）。0 なら枠を付けない。
     * @property {CMYKColor|RGBColor|null} frameColor - 枠の色。
     */

    /**
     * 角丸の効果（Adobe Round Corners）の XML を返します。/ Return the Round Corners live effect XML.
     *
     * @param {number} radius - 半径（pt）。
     * @returns {string} applyEffect() に渡す XML。
     */
    function roundCornersEffectXml(radius) {
        return '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>';
    }

    /**
     * 1つのオブジェクトを矩形でクリッピングマスクします。/ Clip one object with a rectangle.
     *
     * @param {PageItem} pageItem - 対象のオブジェクト。
     * @param {number[]} clipRect - マスクの矩形 [左, 上, 右, 下]。
     * @param {MaskStyle} maskStyle - 角丸と枠の指定。
     * @returns {GroupItem} 作ったクリップグループ。
     */
    function clipItemToRect(pageItem, clipRect, maskStyle) {
        /* 元の重ね順の位置にグループを作る / Create the group where the object sits in the stacking order */
        var clipGroup = pageItem.parent.groupItems.add();
        clipGroup.move(pageItem, ElementPlacement.PLACEBEFORE);
        pageItem.move(clipGroup, ElementPlacement.PLACEATEND);

        var clipPath = clipGroup.pathItems.rectangle(clipRect[1], clipRect[0], clipRect[2] - clipRect[0], clipRect[1] - clipRect[3]);
        /* マスクにするパスはグループの最前面 / The clipping path must be frontmost in the group */
        clipPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        clipGroup.clipped = true;

        /* 枠はマスクにしたあとで付ける（マスクにすると塗り・線が消えるため）/ Stroke after clipping, which clears fill and stroke */
        if (maskStyle.frameWidth > 0 && maskStyle.frameColor) {
            clipPath.stroked = true;
            clipPath.strokeColor = maskStyle.frameColor;
            clipPath.strokeWidth = maskStyle.frameWidth;
        }
        /* 角丸はグループに効果で掛け、マスクと枠をいっしょに丸める / Round the group so the mask and frame round together */
        if (maskStyle.cornerRadius > 0) clipGroup.applyEffect(roundCornersEffectXml(maskStyle.cornerRadius));
        return clipGroup;
    }

    // =========================================
    // リサイズ / Resize
    // =========================================

    /**
     * リサイズの倍率（1 で等倍）をオブジェクトごとに求めます。/ Compute a uniform scale factor per object.
     * 大きさは配置と同じ visibleBounds で測ります。
     *
     * @param {Array<PageItem>} pageItems - セルに入れるオブジェクト。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {string} resizeMode - "width" / "height" / "area" / "scale"。
     * @param {number} ratioPct - 比率（%）。幅・高さ・面積はセルに対して、スケールは元の大きさに対して。
     * @returns {number[]} オブジェクトごとの倍率。測れないものは 1。
     */
    function computeResizeFactors(pageItems, gridMetrics, resizeMode, ratioPct) {
        var ratio = ratioPct / 100;
        var fitWidth = gridMetrics.cellWidth * ratio;
        var fitHeight = gridMetrics.cellHeight * ratio;

        var sizes = [];
        for (var i = 0; i < pageItems.length; i++) {
            var bounds = pageItems[i].visibleBounds;
            sizes.push({ width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] });
        }

        /* 面積：セル×比率に収めたときの面積のうち最小のものにそろえる。どれもセルからはみ出さない
           Area: equalize to the smallest fitted area so that every object still fits its cell */
        var targetArea = Infinity;
        if (resizeMode === "area") {
            for (var j = 0; j < sizes.length; j++) {
                if (sizes[j].width <= 0 || sizes[j].height <= 0) continue;
                var fitFactor = Math.min(fitWidth / sizes[j].width, fitHeight / sizes[j].height);
                targetArea = Math.min(targetArea, sizes[j].width * sizes[j].height * fitFactor * fitFactor);
            }
        }

        var factors = [];
        for (var k = 0; k < sizes.length; k++) {
            var width = sizes[k].width;
            var height = sizes[k].height;
            var factor = 1;
            if (resizeMode === "scale") factor = ratio;
            else if (resizeMode === "width" && width > 0) factor = fitWidth / width;
            else if (resizeMode === "height" && height > 0) factor = fitHeight / height;
            else if (resizeMode === "area" && width > 0 && height > 0 && isFinite(targetArea)) {
                factor = Math.sqrt(targetArea / (width * height));
            }
            factors.push(factor);
        }
        return factors;
    }

    /**
     * セルに入れるオブジェクトを、縦横比を保ってリサイズします（線幅も同じ倍率）。/ Resize objects for the cells.
     *
     * @param {Array<PageItem>} pageItems - セルに入れるオブジェクト（セル数ぶん）。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {string} resizeMode - "none" / "width" / "height" / "area" / "scale"。
     * @param {number|null} ratioPct - 比率（%）。読み取れない場合は null で、リサイズしない。
     * @returns {void}
     */
    function resizeItemsForCells(pageItems, gridMetrics, resizeMode, ratioPct) {
        if (resizeMode === "none" || ratioPct === null) return;

        var factors = computeResizeFactors(pageItems, gridMetrics, resizeMode, ratioPct);
        for (var i = 0; i < pageItems.length; i++) {
            var scalePct = factors[i] * 100;
            /* 等倍は履歴を増やさないよう触らない / Leave unchanged items alone to keep the history short */
            if (!isFinite(scalePct) || scalePct <= 0 || Math.abs(scalePct - 100) < 1e-6) continue;
            try {
                pageItems[i].resize(scalePct, scalePct, true, true, true, true, scalePct, Transformation.CENTER);
            } catch (e) {
                /* 極端な倍率で Illustrator が拒んだものは元の大きさのまま置く / Keep the original size when Illustrator rejects an extreme factor */
            }
        }
    }

    /**
     * すべてのオブジェクトがセルを覆う（上下左右にアキが出ない）、共通のスケール（%）を求めます。/ Compute one scale that makes every object cover its cell.
     * オブジェクトごとに、幅・高さの両方がセル以上になる倍率を求め、いちばん大きいものに合わせます。
     * アキが残らないよう 0.1% 単位で切り上げます（はみ出した分は［マスクする］で切り抜けます）。
     *
     * @param {Array<PageItem>} pageItems - セルに入れるオブジェクト。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @returns {number|null} スケール（%）。測れるオブジェクトが無ければ null。
     */
    function computeCoverScalePct(pageItems, gridMetrics) {
        var coverFactor = -Infinity;
        for (var i = 0; i < pageItems.length; i++) {
            var bounds = pageItems[i].visibleBounds;
            var width = bounds[2] - bounds[0];
            var height = bounds[1] - bounds[3];
            if (width <= 0 || height <= 0) continue;
            /* そのオブジェクトがセルを覆う倍率は、幅・高さの倍率の大きいほう / Covering takes the larger of the two factors */
            coverFactor = Math.max(coverFactor, gridMetrics.cellWidth / width, gridMetrics.cellHeight / height);
        }
        if (!isFinite(coverFactor)) return null;
        return Math.ceil(coverFactor * 1000) / 10;
    }

    /**
     * セルの寸法からアートボードを作成します。/ Create artboards from the cell metrics.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {GridMetrics|null} gridMetrics - グリッドの寸法。
     * @param {number} baseArtboardIndex - 作成後にアクティブへ戻すアートボードの番号。
     * @returns {number} 作成したアートボードの数。
     */
    function createArtboardsFromCells(doc, gridMetrics, baseArtboardIndex) {
        if (!gridMetrics) return 0;

        var createdCount = 0;
        forEachCell(gridMetrics, function (cellRect) {
            try {
                doc.artboards.add(cellRect);
                createdCount++;
            } catch (e) {
                /* アートボード数の上限などで失敗した場合は続行 / Keep going when a single artboard fails */
            }
        });
        doc.artboards.setActiveArtboardIndex(baseArtboardIndex);
        return createdCount;
    }

    /**
     * ドキュメントのカラーモードに合わせた無彩色を返します。/ Return a gray matching the document color mode.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {number} cmykBlack - CMYK のときのブラック値（0〜100）。
     * @param {number} rgbLevel - RGB のときの階調値（0〜255）。
     * @returns {CMYKColor|RGBColor} 生成したカラー。
     */
    function createGrayColor(doc, cmykBlack, rgbLevel) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = cmykBlack;
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = rgbLevel;
        rgbColor.green = rgbLevel;
        rgbColor.blue = rgbLevel;
        return rgbColor;
    }

    /**
     * 色見本に描く画面の RGB（0〜1）を返します。/ Return the screen RGB (0-1) for a swatch.
     *
     * @param {Color} swatchColor - 色。
     * @returns {number[]} [r, g, b]。
     */
    function colorToScreenRgb(swatchColor) {
        if (swatchColor.typename === "RGBColor") return [swatchColor.red / 255, swatchColor.green / 255, swatchColor.blue / 255];
        if (swatchColor.typename === "GrayColor") return [1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100];
        if (swatchColor.typename === "CMYKColor") {
            var blackRatio = 1 - swatchColor.black / 100;
            return [
                (1 - swatchColor.cyan / 100) * blackRatio,
                (1 - swatchColor.magenta / 100) * blackRatio,
                (1 - swatchColor.yellow / 100) * blackRatio
            ];
        }
        /* スポットカラーなどは灰色で示す / Show other color types in gray */
        return [0.5, 0.5, 0.5];
    }

    // =========================================
    // 配置順 / Placement order
    // =========================================

    /**
     * オブジェクトを複製します（元は隠しません）。/ Duplicate the objects, leaving the originals visible.
     *
     * @param {Array<PageItem>} sourceItems - 元のオブジェクト（配置順）。
     * @param {Array<PageItem>} copyStore - 作った複製を1つずつ足す配列（途中で例外になっても、作った分は残る）。
     * @returns {Array<PageItem>} 複製（元と同じ順）。
     */
    function duplicateItems(sourceItems, copyStore) {
        var copies = [];
        for (var i = 0; i < sourceItems.length; i++) {
            var copyItem = sourceItems[i].duplicate();
            copyStore.push(copyItem);
            copies.push(copyItem);
        }
        return copies;
    }

    /**
     * 配列に指定アイテムが含まれるかを判定します。/ Test whether the list contains the item.
     *
     * @param {Array<PageItem>} pageItems - 検索対象の配列。
     * @param {PageItem} pageItem - 探すオブジェクト。
     * @returns {boolean} 含まれる場合は true。
     */
    function containsItem(pageItems, pageItem) {
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i] === pageItem) return true;
        }
        return false;
    }

    /**
     * ランダム順を現在の配置対象に合わせ直します。/ Rebuild the random order for the current items.
     * 対象の切り替えでアイテムが増減しても破綻しないようにします。
     *
     * @param {Array<PageItem>} distributionItems - 現在の配置対象。
     * @param {Array<PageItem>|null} randomizedOrder - シャッフルで決めた順。未使用なら null。
     * @returns {Array<PageItem>} 配置順に並べたオブジェクト。
     */
    function orderByRandomizedOrder(distributionItems, randomizedOrder) {
        if (!randomizedOrder) return distributionItems;

        var orderedItems = [];
        for (var i = 0; i < randomizedOrder.length; i++) {
            if (containsItem(distributionItems, randomizedOrder[i])) orderedItems.push(randomizedOrder[i]);
        }
        for (var j = 0; j < distributionItems.length; j++) {
            if (!containsItem(orderedItems, distributionItems[j])) orderedItems.push(distributionItems[j]);
        }
        return orderedItems;
    }

    /**
     * マスクしたオブジェクトを、そのクリップグループに置き換えた配列を返します。/ Swap masked items for their clip groups.
     *
     * @param {Array<PageItem>} pageItems - 元のオブジェクト。
     * @param {Array<GroupItem>} clipGroups - maskItemsToRects() で作ったクリップグループ。
     * @returns {Array<PageItem>} 選択に使う配列。
     */
    function replaceMaskedItems(pageItems, clipGroups) {
        var resultItems = [];
        for (var i = 0; i < pageItems.length; i++) {
            resultItems.push(containsItem(clipGroups, pageItems[i].parent) ? pageItems[i].parent : pageItems[i]);
        }
        return resultItems;
    }

    /**
     * 順序をシャッフルした新しい配列を返します（Fisher-Yates）。/ Return a shuffled copy.
     *
     * @param {Array} sourceItems - 元の配列。
     * @param {function(): number} [randomFn] - 0 以上 1 未満を返す乱数。省略時は Math.random。
     * @returns {Array} シャッフルした配列。
     */
    function createShuffledCopy(sourceItems, randomFn) {
        var nextRandom = randomFn || Math.random;
        var shuffledItems = sourceItems.slice(0);
        for (var k = shuffledItems.length - 1; k > 0; k--) {
            var swapIndex = Math.floor(nextRandom() * (k + 1));
            var swapItem = shuffledItems[k];
            shuffledItems[k] = shuffledItems[swapIndex];
            shuffledItems[swapIndex] = swapItem;
        }
        return shuffledItems;
    }

    // =========================================
    // UIの組み立て / UI construction
    // =========================================

    /**
     * 対象パネルを作成します。/ Build the target panel.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @param {boolean} hasBackmostItem - 最背面のオブジェクトが存在する場合は true。
     * @returns {{artboardRadio: RadioButton, backmostRadio: RadioButton}} 配置先のラジオボタン。
     */
    function buildPlacementTargetPanel(parentColumn, hasBackmostItem) {
        var placementPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.placement));
        setupPanel(placementPanel, RADIO_COLUMN_SPACING);
        placementPanel.alignChildren = ["left", "top"];

        var placementControls = {
            artboardRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetArtboard)),
            backmostRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetBackmost))
        };

        placementControls.backmostRadio.enabled = hasBackmostItem;
        placementControls.artboardRadio.helpTip = getLabel(LABELS.tooltip.targetArtboard);

        /* 選べない理由、または選んだときの挙動をツールチップで補う / Tooltips explain why an option is unavailable or what it does */
        placementControls.backmostRadio.helpTip = hasBackmostItem
            ? getLabel(LABELS.tooltip.targetBackmost)
            : getLabel(LABELS.tooltip.targetUnavailable);
        placementControls.artboardRadio.value = true;
        return placementControls;
    }

    /**
     * ラベルと入力欄の1行を作成します。/ Build one label-and-field row.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @param {LabelEntry} labelEntry - 行ラベルの文言。
     * @param {number} initialValue - 入力欄の初期値。
     * @param {number} labelWidth - ラベルの幅（px）。
     * @param {string|null} unitSuffix - 入力欄の後ろに置く単位表記。不要なら null。
     * @param {LabelEntry|null} tooltipEntry - ラベルと入力欄に付けるツールチップ。不要なら null。
     * @param {Object} stepOptions - ∧∨の増減設定（addStepper() に渡す。個数の欄は integer: true）。
     * @returns {EditText} 作成した入力欄。
     */
    function addFieldRow(parentColumn, labelEntry, initialValue, labelWidth, unitSuffix, tooltipEntry, stepOptions) {
        var fieldRow = parentColumn.add("group");
        setupRow(fieldRow, "left", 4);
        fieldRow.alignChildren = "center";

        var fieldLabel = fieldRow.add("statictext", undefined, labelText(labelEntry));
        fieldLabel.preferredSize.width = labelWidth;
        fieldLabel.justify = "right";

        var fieldInput = addSteppedInput(fieldRow, String(initialValue), stepOptions);
        fieldInput.characters = unitSuffix ? 4 : 3;
        if (unitSuffix) fieldRow.add("statictext", undefined, unitSuffix);

        if (tooltipEntry) {
            fieldLabel.helpTip = getLabel(tooltipEntry);
            fieldInput.helpTip = getLabel(tooltipEntry);
        }
        return fieldInput;
    }

    /**
     * ∧∨と入力欄を隙間なく並べて行に追加し、↑↓キーも∧∨と同じ処理で増減させます。/ Add a field with a stepper.
     * 増減後は入力欄の onChange（enableNumericField() で結線）を呼び、範囲の正規化とプレビュー更新を手入力と同じにします。
     *
     * @param {Group} parentRow - 配置先の行。
     * @param {string} initialText - 入力欄の初期値。
     * @param {Object} stepOptions - ∧∨の増減設定（addStepper() に渡す。個数の欄は integer: true）。
     * @returns {EditText} 作成した入力欄（∧∨は .stepperGroup で参照できる）。
     */
    function addSteppedInput(parentRow, initialText, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var inputField;
        stepOptions.onStep = function (numberInput) { if (numberInput.onChange) numberInput.onChange(); };
        var stepperGroup = addStepper(stepperInputGroup, function () { return inputField; }, stepOptions);
        inputField = stepperInputGroup.add("edittext", undefined, initialText);
        inputField.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(inputField, stepperGroup);
        return inputField;
    }

    /**
     * 入力欄の有効／無効を、左の∧∨ごと切り替えます（∧∨は自作描画なので描き直してディム表示をそろえます）。
     * Enable or disable a field together with its stepper.
     *
     * @param {EditText} inputField - addSteppedInput() で作った入力欄。
     * @param {boolean} isEnabled - 有効にする場合は true。
     * @returns {void}
     */
    function setSteppedInputEnabled(inputField, isEnabled) {
        /* 入力のたびに呼ばれるので、変わらないときは描き直さない / called on every keystroke; skip when unchanged */
        if (inputField.enabled === isEnabled && inputField.stepperGroup.enabled === isEnabled) return;
        inputField.enabled = isEnabled;
        inputField.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(inputField.stepperGroup);
    }

    /**
     * 分割とマージンのパネルを作成します（行数・列数・セル間隔・マージンを縦に並べる）。/ Build the division panel.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @param {{rowCount: number, columnCount: number}} defaultDivision - 行数・列数の初期値。
     * @param {string} unitLabel - 定規単位のラベル。
     * @returns {{rowCountInput: EditText, columnCountInput: EditText, autoDivisionButton: Group, gutterInput: EditText, marginInput: EditText, gutterPriorityCheckbox: Checkbox}} 入力欄、行数・列数の自動計算のアイコン、［セル間隔を優先］。
     */
    function buildDivisionPanel(parentColumn, defaultDivision, unitLabel) {
        var divisionPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.division));
        setupPanel(divisionPanel);
        divisionPanel.alignChildren = ["left", "top"];

        /* 4項目の項目名を同じ幅で右揃えにする。幅は日本語環境と英語環境で個別に指定
           All four labels share one width, right-aligned; widths per language */
        var fieldLabelWidth = (uiLang === "ja") ? 70 : 75;

        /* 行数・列数の2行の右に、2行の天地中央で自動計算のアイコンを置く / The Auto icon sits to the right of the rows and columns, centered on both */
        var countGroup = divisionPanel.add("group");
        setupRow(countGroup, "left", 8);
        countGroup.alignChildren = ["left", "center"];
        var countFieldColumn = countGroup.add("group");
        countFieldColumn.orientation = "column";
        countFieldColumn.alignChildren = ["left", "top"];
        countFieldColumn.spacing = PANEL_SPACING;

        var divisionControls = {
            /* 範囲の正規化は enableNumericField() の onChange が受け持つ / range clamping is done in enableNumericField() */
            rowCountInput: addFieldRow(countFieldColumn, LABELS.fieldLabel.rowCount, defaultDivision.rowCount, fieldLabelWidth, null, LABELS.tooltip.division, { integer: true }),
            columnCountInput: addFieldRow(countFieldColumn, LABELS.fieldLabel.columnCount, defaultDivision.columnCount, fieldLabelWidth, null, LABELS.tooltip.division, { integer: true }),
            autoDivisionButton: addIconButton(countGroup, ICON_BUTTON_SIZE, drawAutoIcon),
            gutterInput: addFieldRow(divisionPanel, LABELS.fieldLabel.gutter, CONFIG.defaultGutter, fieldLabelWidth, unitLabel, LABELS.tooltip.gutter, {}),
            marginInput: addFieldRow(divisionPanel, LABELS.fieldLabel.margin, CONFIG.defaultMargin, fieldLabelWidth, unitLabel, LABELS.tooltip.margin, {})
        };
        divisionControls.autoDivisionButton.alignment = ["left", "center"];
        divisionControls.autoDivisionButton.helpTip = getLabel(LABELS.tooltip.autoDivision);
        divisionControls.gutterPriorityCheckbox = divisionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.gutterPriority));
        divisionControls.gutterPriorityCheckbox.helpTip = getLabel(LABELS.tooltip.gutterPriority);
        divisionControls.gutterPriorityCheckbox.value = CONFIG.defaultGutterPriority;
        return divisionControls;
    }

    /**
     * リサイズのパネルを作成します（基準のラジオボタンを縦に並べ、その下に比率と自動計算のアイコン）。/ Build the resize panel.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @returns {Object} 基準のラジオボタン、比率の項目名と入力欄、自動計算のアイコンをまとめたオブジェクト。
     */
    function buildResizePanel(parentColumn) {
        var resizePanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.resize));
        setupPanel(resizePanel, PANEL_SECTION_SPACING);
        resizePanel.alignChildren = ["left", "top"];

        var resizeModeGroup = addRadioColumn(resizePanel);

        var resizeRatioRow = resizePanel.add("group");
        setupRow(resizeRatioRow, "left", 4);
        resizeRatioRow.alignChildren = "center";

        var resizeControls = {
            noneRadio: resizeModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.resizeNone)),
            widthRadio: resizeModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.resizeWidth)),
            heightRadio: resizeModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.resizeHeight)),
            areaRadio: resizeModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.resizeArea)),
            scaleRadio: resizeModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.resizeScale)),
            ratioLabel: resizeRatioRow.add("statictext", undefined, labelText(LABELS.fieldLabel.resizeRatio))
        };
        resizeControls.ratioInput = addSteppedInput(resizeRatioRow, String(CONFIG.defaultResizeRatio), {});
        resizeControls.ratioInput.characters = 4;
        resizeControls.ratioUnitLabel = resizeRatioRow.add("statictext", undefined, "%");
        resizeControls.noneRadio.value = true;

        /* スケールのときだけ使う自動計算のアイコンを、比率の単位の右に置く / The Auto icon (Scale only) sits after the ratio's unit */
        resizeControls.autoScaleButton = addIconButton(resizeRatioRow, ICON_BUTTON_SIZE, drawAutoIcon);
        resizeControls.autoScaleButton.helpTip = getLabel(LABELS.tooltip.autoScale);

        /* 基準ごとの結果をツールチップで補う / Tooltips describe each mode */
        resizeControls.noneRadio.helpTip = getLabel(LABELS.tooltip.resizeNone);
        resizeControls.widthRadio.helpTip = getLabel(LABELS.tooltip.resizeWidth);
        resizeControls.heightRadio.helpTip = getLabel(LABELS.tooltip.resizeHeight);
        resizeControls.areaRadio.helpTip = getLabel(LABELS.tooltip.resizeArea);
        resizeControls.scaleRadio.helpTip = getLabel(LABELS.tooltip.resizeScale);
        resizeControls.ratioLabel.helpTip = getLabel(LABELS.tooltip.resizeRatio);
        resizeControls.ratioInput.helpTip = getLabel(LABELS.tooltip.resizeRatio);
        return resizeControls;
    }

    /**
     * マスクのパネルを作成します（先頭にマスクのプリセット、［なし］／［セル］／［画像］と、その下の角丸・枠）。/ Build the mask panel.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @param {string} rulerUnitLabel - 角丸の半径の単位（定規の単位）。
     * @param {string} strokeUnitLabel - 枠の線幅の単位（線の単位）。
     * @returns {Object} マスク・角丸・枠の項目をまとめたオブジェクト。
     */
    function buildMaskPanel(parentColumn, rulerUnitLabel, strokeUnitLabel) {
        var maskPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.mask));
        setupPanel(maskPanel, PANEL_SECTION_SPACING);
        maskPanel.alignChildren = ["left", "top"];

        var maskControls = {};
        maskControls.presetUI = buildPresetRow(maskPanel);
        /* ラジオは同じ親の中だけ排他になるので、3つを1行のグループにまとめる / Radios are exclusive only within one parent */
        var maskModeRow = maskPanel.add("group");
        setupRow(maskModeRow, "left");
        maskControls.maskNoneRadio = maskModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.maskNone));
        maskControls.maskCellRadio = maskModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.maskCell));
        maskControls.maskItemRadio = maskModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.maskItem));
        maskControls.maskNoneRadio.helpTip = getLabel(LABELS.tooltip.maskNone);
        maskControls.maskCellRadio.helpTip = getLabel(LABELS.tooltip.maskCell);
        maskControls.maskItemRadio.helpTip = getLabel(LABELS.tooltip.maskItem);
        selectMaskModeRadio(maskControls, CONFIG.defaultMaskMode);

        /* 角丸と枠はマスクの種類の下に、字下げせずに置く / Round Corners and Frame sit under the mask modes without an indent */
        var maskStyleGroup = maskPanel.add("group");
        maskStyleGroup.orientation = "column";
        maskStyleGroup.alignChildren = ["left", "top"];
        maskStyleGroup.spacing = RADIO_COLUMN_SPACING;

        var cornerRadiusRow = maskStyleGroup.add("group");
        setupRow(cornerRadiusRow, "left", 4);
        maskControls.cornerRadiusCheckbox = cornerRadiusRow.add("checkbox", undefined, labelText(LABELS.checkbox.cornerRadius));
        maskControls.cornerRadiusCheckbox.value = CONFIG.defaultUseCornerRadius;
        maskControls.cornerRadiusInput = addSteppedInput(cornerRadiusRow, String(CONFIG.defaultCornerRadius), {});
        maskControls.cornerRadiusInput.characters = 4;
        maskControls.cornerRadiusUnitLabel = cornerRadiusRow.add("statictext", undefined, rulerUnitLabel);
        /* 角丸の自動計算のアイコンを、半径の単位の右に置く / The Auto icon for the radius sits after its unit */
        maskControls.autoCornerRadiusButton = addIconButton(cornerRadiusRow, ICON_BUTTON_SIZE, drawAutoIcon);
        maskControls.autoCornerRadiusButton.helpTip = getLabel(LABELS.tooltip.autoCornerRadius);

        var frameRow = maskStyleGroup.add("group");
        setupRow(frameRow, "left", 4);
        maskControls.frameCheckbox = frameRow.add("checkbox", undefined, labelText(LABELS.checkbox.frameStroke));
        maskControls.frameCheckbox.value = CONFIG.defaultUseFrame;
        maskControls.frameWidthInput = addSteppedInput(frameRow, String(CONFIG.defaultFrameWidth), {});
        maskControls.frameWidthInput.characters = 4;
        maskControls.frameUnitLabel = frameRow.add("statictext", undefined, strokeUnitLabel);

        /* 角丸・枠のチェックボックスは項目名を兼ねるので、行の項目名の幅をそろえる / The checkboxes double as labels; align their widths */
        var maskStyleLabelWidth = Math.max(
            measureControlText(maskControls.cornerRadiusCheckbox),
            measureControlText(maskControls.frameCheckbox)
        ) + CHECKBOX_BOX_WIDTH;
        maskControls.cornerRadiusCheckbox.preferredSize.width = maskStyleLabelWidth;
        maskControls.frameCheckbox.preferredSize.width = maskStyleLabelWidth;

        /* 枠の色見本は次の行に、入力欄（∧∨）の左端にそろえて置く / The frame swatch sits on the next line, aligned with the field */
        var frameColorRow = maskStyleGroup.add("group");
        setupRow(frameColorRow, "left", 4);
        var frameColorSpacer = frameColorRow.add("group");
        frameColorSpacer.preferredSize = [maskStyleLabelWidth, 1];
        maskControls.frameColorSwatch = frameColorRow.add("group");
        maskControls.frameColorSwatch.preferredSize = COLOR_SWATCH_SIZE;
        maskControls.frameColorSwatch.helpTip = getLabel(LABELS.tooltip.frameColor);

        maskControls.cornerRadiusCheckbox.helpTip = getLabel(LABELS.tooltip.cornerRadius);
        maskControls.cornerRadiusInput.helpTip = getLabel(LABELS.tooltip.cornerRadius);
        maskControls.frameCheckbox.helpTip = getLabel(LABELS.tooltip.frameStroke);
        maskControls.frameWidthInput.helpTip = getLabel(LABELS.tooltip.frameStroke);
        return maskControls;
    }

    /**
     * アイテムの基準位置のパネルを作成します（セルの中でオブジェクトを寄せる位置、9か所）。/ Build the object anchor panel.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @param {function(): void} onItemAnchorChange - 基準位置を選んだときに呼ぶ処理。
     * @returns {{itemAnchorWidget: Button}} 基準点ウィジェット。
     */
    function buildItemAnchorPanel(parentColumn, onItemAnchorChange) {
        var itemAnchorPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.itemAnchor));
        setupPanel(itemAnchorPanel);
        itemAnchorPanel.alignChildren = ["left", "top"];
        var itemAnchorWidget = addAnchorWidget(itemAnchorPanel, CONFIG.defaultItemAnchor, onItemAnchorChange, { size: FILL_START_WIDGET_SIZE });
        itemAnchorWidget.alignment = ["center", "top"]; /* パネルの左右中央に置く / Center the widget in the panel */
        itemAnchorWidget.helpTip = getLabel(LABELS.tooltip.itemAnchor);
        return { itemAnchorWidget: itemAnchorWidget };
    }

    /**
     * セル描画パネルを作成します。/ Build the cell drawing panel.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @returns {Object} セル描画モード、カラー、不透明度、透明グリッドボタンをまとめたオブジェクト。
     */
    function buildCellHandlingPanel(parentColumn) {
        var cellHandlingPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.cellHandling));
        setupPanel(cellHandlingPanel, PANEL_SECTION_SPACING);
        cellHandlingPanel.alignChildren = ["left", "top"];

        var cellModeGroup = addRadioColumn(cellHandlingPanel);

        /* カラーは項目名の下にラジオボタンを縦に並べる / The color radios stack under their label */
        var cellColorGroup = addRadioColumn(cellHandlingPanel);
        cellColorGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.cellColor));

        /* 不透明度は項目名の下の行に入力欄を置く / The opacity field sits on the line below its label */
        var cellOpacityGroup = addRadioColumn(cellHandlingPanel);
        cellOpacityGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.cellOpacity));
        var cellOpacityRow = cellOpacityGroup.add("group");
        setupRow(cellOpacityRow, "left", 4);

        var cellControls = {
            noCellRadio: cellModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.noCell)),
            keepCellRadio: cellModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.keepCell)),
            toGuideRadio: cellModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.toGuide)),
            toArtboardRadio: cellModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.toArtboard)),
            blackCellRadio: cellColorGroup.add("radiobutton", undefined, getLabel(LABELS.radio.blackCell)),
            whiteCellRadio: cellColorGroup.add("radiobutton", undefined, getLabel(LABELS.radio.whiteCell)),
            transparentCellRadio: cellColorGroup.add("radiobutton", undefined, getLabel(LABELS.radio.transparentCell)),
            opacityInput: addSteppedInput(cellOpacityRow, String(CONFIG.blackCellOpacity), {})
        };
        /* マスクを掛けて始めるときは、セルの長方形を残さない / Start with no cell rectangles when a mask is on */
        selectCellModeRadio(cellControls, (CONFIG.defaultMaskMode === MASK_MODE_NONE) ? "keep" : "none");
        cellControls.blackCellRadio.value = true;
        cellControls.opacityInput.characters = 4;

        /* モードごとの結果をツールチップで補う / Tooltips describe each mode's result */
        cellControls.noCellRadio.helpTip = getLabel(LABELS.tooltip.noCell);
        cellControls.keepCellRadio.helpTip = getLabel(LABELS.tooltip.keepCell);
        cellControls.toGuideRadio.helpTip = getLabel(LABELS.tooltip.toGuide);
        cellControls.toArtboardRadio.helpTip = getLabel(LABELS.tooltip.toArtboard);

        cellOpacityRow.add("statictext", undefined, "%");

        /* ボタンはパネル幅いっぱいに広げない。カラムを細く保つよう、不透明度の下の行に置く
           Keep the button from stretching; it sits below the opacity to keep the column narrow */
        cellControls.transparencyGridButton = cellHandlingPanel.add("button", undefined, getLabel(LABELS.button.transparencyGrid));
        cellControls.transparencyGridButton.alignment = "left";
        cellControls.transparencyGridButton.helpTip = getLabel(LABELS.tooltip.transparencyGrid);
        return cellControls;
    }

    /**
     * 配置と整列のパネルを作成します（開始位置・均等にまばらに・レンガ・シャッフル）。/ Build the arrangement panel.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @param {function(): void} onFillStartChange - 開始位置を選んだときに呼ぶ処理。
     * @returns {{fillStartLabel: StaticText, fillStartWidget: Button, spreadEvenlyCheckbox: Checkbox, brickCheckbox: Checkbox, keepOriginalsCheckbox: Checkbox, btnShuffle: Button}} 開始位置の項目名・ウィジェット、［均等にまばらに］、［レンガ］、［元のオブジェクトを残す］、［シャッフル］。
     */
    function buildArrangementPanel(parentColumn, onFillStartChange) {
        var arrangementPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.arrangement));
        setupPanel(arrangementPanel, PANEL_SECTION_SPACING);
        arrangementPanel.alignChildren = ["left", "top"];

        /* 開始位置は基準点ウィジェットの9か所すべてから選べる / Start From can start at any of the nine anchor positions */
        var fillStartRow = arrangementPanel.add("group");
        setupRow(fillStartRow, "left", 8);
        fillStartRow.alignChildren = ["left", "top"];
        var fillStartLabel = fillStartRow.add("statictext", undefined, labelText(LABELS.fieldLabel.fillStart));
        var fillStartWidget = addAnchorWidget(fillStartRow, CONFIG.defaultFillStart, onFillStartChange, { size: FILL_START_WIDGET_SIZE });
        fillStartLabel.helpTip = getLabel(LABELS.tooltip.fillStart);
        fillStartWidget.helpTip = getLabel(LABELS.tooltip.fillStart);

        var spreadEvenlyCheckbox = arrangementPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.spreadEvenly));
        spreadEvenlyCheckbox.helpTip = getLabel(LABELS.tooltip.spreadEvenly);
        spreadEvenlyCheckbox.value = CONFIG.defaultSpreadEvenly;

        var brickCheckbox = arrangementPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.brick));
        brickCheckbox.helpTip = getLabel(LABELS.tooltip.brick);
        brickCheckbox.value = CONFIG.defaultBrick;

        var keepOriginalsCheckbox = arrangementPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.keepOriginals));
        keepOriginalsCheckbox.helpTip = getLabel(LABELS.tooltip.keepOriginals);
        keepOriginalsCheckbox.value = CONFIG.defaultKeepOriginals;

        var btnShuffle = arrangementPanel.add("button", undefined, getLabel(LABELS.button.randomize));
        btnShuffle.alignment = "left"; /* パネル幅いっぱいに広げない / Keep the button from stretching */
        btnShuffle.helpTip = getLabel(LABELS.tooltip.randomize);
        return {
            fillStartLabel: fillStartLabel,
            fillStartWidget: fillStartWidget,
            spreadEvenlyCheckbox: spreadEvenlyCheckbox,
            brickCheckbox: brickCheckbox,
            keepOriginalsCheckbox: keepOriginalsCheckbox,
            btnShuffle: btnShuffle
        };
    }

    /**
     * 下部のボタン列を作成します（左に［画面にフィット］、右に［キャンセル］［OK］）。/ Build the footer button row.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @returns {{fitViewControls: Object, btnCancel: Button, btnOK: Button}} ［画面にフィット］の一式（FitViewToItems.addControls() の戻り値）と各ボタン。
     */
    function buildFooterRow(parentWindow) {
        var buttonRow = addButtonRow(parentWindow);
        var fitViewControls = FitViewToItems.addControls(buttonRow.leftGroup, {
            value: CONFIG.defaultFitView,
            percent: CONFIG.defaultFitViewPercent,
            stepper: true
        });

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return { fitViewControls: fitViewControls, btnCancel: btnCancel, btnOK: btnOK };
    }

    /**
     * リサイズの基準のラジオを選びます（同じ親のほかのラジオは外す）。/ Select a resize-mode radio.
     *
     * @param {Object} resizeUI - buildResizePanel() の戻り値。
     * @param {string} resizeMode - "none" / "width" / "height" / "area" / "scale"。知らない値は "none"。
     * @returns {void}
     */
    function selectResizeRadio(resizeUI, resizeMode) {
        resizeUI.widthRadio.value = (resizeMode === "width");
        resizeUI.heightRadio.value = (resizeMode === "height");
        resizeUI.areaRadio.value = (resizeMode === "area");
        resizeUI.scaleRadio.value = (resizeMode === "scale");
        resizeUI.noneRadio.value = !(resizeUI.widthRadio.value || resizeUI.heightRadio.value || resizeUI.areaRadio.value || resizeUI.scaleRadio.value);
    }

    /**
     * マスクの種類のラジオを選びます（同じ行のほかのラジオは外す）。/ Select a mask-mode radio.
     *
     * @param {Object} maskUI - buildMaskPanel() の戻り値。
     * @param {string} maskMode - MASK_MODE_NONE / MASK_MODE_CELL / MASK_MODE_ITEM。
     * @returns {void}
     */
    function selectMaskModeRadio(maskUI, maskMode) {
        maskUI.maskNoneRadio.value = (maskMode === MASK_MODE_NONE);
        maskUI.maskCellRadio.value = (maskMode === MASK_MODE_CELL);
        maskUI.maskItemRadio.value = (maskMode === MASK_MODE_ITEM);
    }

    /**
     * セルの扱いのラジオを選びます（同じ列のほかのラジオは外す）。/ Select a cell-handling radio.
     *
     * @param {Object} cellUI - buildCellHandlingPanel() の戻り値。
     * @param {string} cellMode - "none" / "keep" / "guide" / "artboard"。
     * @returns {void}
     */
    function selectCellModeRadio(cellUI, cellMode) {
        cellUI.noCellRadio.value = (cellMode === "none");
        cellUI.keepCellRadio.value = (cellMode === "keep");
        cellUI.toGuideRadio.value = (cellMode === "guide");
        cellUI.toArtboardRadio.value = (cellMode === "artboard");
    }

    /**
     * 基準点ウィジェットのセルを選んで描き直します（onChange は呼びません）。/ Set an anchor widget's cell and redraw it.
     *
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット。
     * @param {number|string} anchorValue - 0〜8 か "topLeft" などの名前。
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * マスクのプリセットの行を作成します（［マスク］パネルの先頭。ドロップダウン・保存と削除のアイコン）。/ Build the mask preset row.
     *
     * @param {Panel} parentPanel - 追加先の［マスク］パネル。
     * @returns {{presetDropdown: DropDownList, presetSaveIcon: Group, presetDeleteIcon: Group}} プリセットのコントロール。
     */
    function buildPresetRow(parentPanel) {
        var presetRow = parentPanel.add("group");
        setupRow(presetRow, ["fill", "center"], 4);
        /* パネルの幅に合わせて伸ばす（幅を指定してダイアログボックスを広げない） / Stretch to the panel instead of widening the dialog */
        var presetDropdown = presetRow.add("dropdownlist", undefined, []);
        presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);
        presetDropdown.alignment = ["fill", "center"];
        presetDropdown.preferredSize.width = PRESET_DROPDOWN_MIN_WIDTH;
        /* 保存アイコンは塗りの面が大きく濃く見えるので、色を薄めて描く / The save icon is mostly solid and looks heavy, so draw it lighter */
        var presetSaveIcon = addIconButton(presetRow, PRESET_ICON_SIZE, function (iconGraphics, iconWidth, iconHeight, iconColor) {
            drawSaveIcon(iconGraphics, iconWidth, iconHeight, [iconColor[0], iconColor[1], iconColor[2], iconColor[3] * SAVE_ICON_OPACITY]);
        });
        presetSaveIcon.helpTip = getLabel(LABELS.tooltip.presetSave);
        var presetDeleteIcon = addIconButton(presetRow, PRESET_ICON_SIZE, drawTrashIcon);
        presetDeleteIcon.helpTip = getLabel(LABELS.tooltip.presetDelete);
        return { presetDropdown: presetDropdown, presetSaveIcon: presetSaveIcon, presetDeleteIcon: presetDeleteIcon };
    }

    /**
     * 3カラムのダイアログボックスを組み立てます（イベントは結び付けません）。/ Build the three-column dialog box.
     * 左に配置先・分割とマージン・配置と整列、中央にリサイズ・マスク、右にセルの扱いを置きます。
     *
     * @param {boolean} hasBackmostItem - 最背面のオブジェクトが存在する場合は true。
     * @param {{rowCount: number, columnCount: number}} defaultDivision - 行数・列数の初期値。
     * @param {string} rulerUnitLabel - 定規の単位のラベル。
     * @param {string} strokeUnitLabel - 線の単位のラベル。
     * @param {function(): void} onFillStartChange - 開始位置を選んだときに呼ぶ処理。
     * @param {function(): void} onItemAnchorChange - アイテムの基準位置を選んだときに呼ぶ処理。
     * @returns {Object} ダイアログボックス（distributeDialog）と、各パネル・ボタン行のコントロール。
     */
    function buildDistributeDialog(hasBackmostItem, defaultDivision, rulerUnitLabel, strokeUnitLabel, onFillStartChange, onItemAnchorChange) {
        var distributeDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(distributeDialog);

        var columnsGroup = distributeDialog.add("group");
        setupRow(columnsGroup, ["fill", "top"], COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "fill"];
        var leftColumn = addDialogColumn(columnsGroup);
        var centerColumn = addDialogColumn(columnsGroup);
        var rightColumn = addDialogColumn(columnsGroup);

        return {
            distributeDialog: distributeDialog,
            placementUI: buildPlacementTargetPanel(leftColumn, hasBackmostItem),
            divisionUI: buildDivisionPanel(leftColumn, defaultDivision, rulerUnitLabel),
            arrangementUI: buildArrangementPanel(leftColumn, onFillStartChange),
            resizeUI: buildResizePanel(centerColumn),
            maskUI: buildMaskPanel(centerColumn, rulerUnitLabel, strokeUnitLabel),
            cellUI: buildCellHandlingPanel(rightColumn),
            itemAnchorUI: buildItemAnchorPanel(rightColumn, onItemAnchorChange),
            footerUI: buildFooterRow(distributeDialog)
        };
    }

    // =========================================
    // 入力欄 / Input fields
    // =========================================

    /**
     * 入力欄の値を整数として読み取ります。/ Read an integer from a field.
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @returns {number|null} 読み取った整数。数値でない場合は null。
     */
    function readCount(inputField) {
        var parsedCount = parseInt(inputField.text, 10);
        return isFinite(parsedCount) ? parsedCount : null;
    }

    /**
     * 入力欄の値を数値として読み取り、下限・上限に収めます。/ Read a number from a field, clamped to a range.
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @param {number} minValue - 下限値。
     * @param {number} maxValue - 上限値。
     * @returns {number|null} 範囲に収めた値。数値でない場合は null。
     */
    function readClampedNumber(inputField, minValue, maxValue) {
        var parsedValue = parseFloat(inputField.text);
        if (!isFinite(parsedValue)) return null;
        return Math.max(minValue, Math.min(maxValue, parsedValue));
    }

    /**
     * 入力欄の値を pt 換算の長さとして読み取ります。/ Read a length in points from a field.
     * 負の値は配置先からはみ出したグリッドを生むため、0 で止めます。
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位（getUnitInfo() の戻り値）。
     * @returns {number} pt に換算した長さ。数値でない場合は 0。
     */
    function readLengthPt(inputField, rulerUnit) {
        var parsedLength = parseFloat(inputField.text);
        if (!isFinite(parsedLength)) return 0;
        return unitToPt(Math.max(0, parsedLength), rulerUnit);
    }

    /**
     * 数値入力欄に範囲の正規化を組み込みます（∧∨・↑↓キーの増減も、この onChange を通ります）。
     * Wire range clamping into a numeric field; the stepper and arrow keys go through this onChange too.
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @param {number} minValue - 下限値。
     * @param {number|null} maxValue - 上限値。上限なしは null。
     * @param {function(): void} onCommit - 値を変えたあとに呼ぶ処理。
     * @returns {void}
     */
    function enableNumericField(inputField, minValue, maxValue, onCommit) {
        /** 入力欄を書き換えている最中か（onChange の再入を防ぐ） / Guard against re-entering onChange. */
        var isNormalizing = false;

        /**
         * 値を下限・上限に収めます。/ Clamp a value into range.
         *
         * @param {number} value - 対象の値。
         * @returns {number} 範囲内に収めた値。
         */
        function clampValue(value) {
            var clamped = Math.max(minValue, value);
            return (maxValue === null) ? clamped : Math.min(maxValue, clamped);
        }

        /* 入力を確定した時点で範囲外の値を直し、表示と描画を一致させる
           Normalize out-of-range text on commit so the field matches what is drawn */
        inputField.onChange = function () {
            if (isNormalizing) return;

            var fieldValue = parseFloat(inputField.text);
            if (isFinite(fieldValue)) {
                var clamped = clampValue(fieldValue);
                if (clamped !== fieldValue) {
                    isNormalizing = true;
                    inputField.text = String(clamped);
                    isNormalizing = false;
                }
            }
            onCommit();
        };
    }

    // =========================================
    // 設定の保存 / Settings store
    // =========================================

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME を元にした名前）。ファイル名と $.global のキーに使う
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

    // =========================================
    // 画像の読み込み / Image import
    // =========================================

    /* 対象の拡張子 / Target extensions */
    var IMAGE_EXTENSIONS = ["jpg", "jpeg", "png", "tif", "tiff", "psd", "gif", "bmp", "pdf", "eps", "ai", "svg", "heic", "webp"];

    /* 画像の選び方 / How the images are chosen */
    var IMAGE_SOURCE_FILES = "files";
    var IMAGE_SOURCE_FOLDER = "folder";

    /* 前回選んだ方式（次回の既定のボタンにする）。再起動しても残る / The last choice, used as the default button next time; kept across restarts */
    var imageSourceStore = createSettingsStore(SCRIPT_NAME + "ImageSource", "persistent");

    /**
     * ファイル名の拡張子が対象の画像かどうかを返す。
     * @param {string} fileName - ファイル名
     * @returns {boolean} 対象の拡張子なら true
     */
    function isImageName(fileName) {
        var extMatch = /\.([^.]+)$/.exec(fileName);
        if (!extMatch) return false;
        var ext = extMatch[1].toLowerCase();
        for (var i = 0; i < IMAGE_EXTENSIONS.length; i++) {
            if (IMAGE_EXTENSIONS[i] === ext) return true;
        }
        return false;
    }

    /**
     * 画像をファイルで選ぶかフォルダーで選ぶかを尋ねる（File.openDialog ではフォルダーを選べないため）。前回選んだ方式を既定のボタンにする。
     * @returns {string|null} IMAGE_SOURCE_FILES / IMAGE_SOURCE_FOLDER。キャンセルなら null
     */
    function showImageSourceDialog() {
        var isFolderDefault = (imageSourceStore.load({ imageSource: IMAGE_SOURCE_FILES }).imageSource === IMAGE_SOURCE_FOLDER);
        var sourceDialog = new Window("dialog", getLabel(LABELS.dialog.chooseSource));
        setupWindow(sourceDialog);
        sourceDialog.add("statictext", undefined, getLabel(LABELS.message.chooseSource));

        /* ボタン行（右：キャンセル・フォルダー・ファイル）。前回選んだほうを既定（return）にする / Button row; the last choice is the default (Return) */
        var buttonRow = addButtonRow(sourceDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnFolder = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.chooseFolder), isFolderDefault ? { name: "ok" } : {});
        var btnFiles = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.chooseFiles), isFolderDefault ? {} : { name: "ok" });
        btnCancel.onClick = function () { sourceDialog.close(2); };
        btnFolder.onClick = function () { sourceDialog.close(3); };
        btnFiles.onClick = function () { sourceDialog.close(1); };
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(sourceDialog, SCRIPT_NAME + "_imageSource");

        var sourceDialogResult = sourceDialog.show();
        var imageSource = null;
        if (sourceDialogResult === 1) imageSource = IMAGE_SOURCE_FILES;
        if (sourceDialogResult === 3) imageSource = IMAGE_SOURCE_FOLDER;
        if (imageSource) imageSourceStore.save({ imageSource: imageSource });
        return imageSource;
    }

    /**
     * ファイル選択ダイアログボックスを開いて画像を選ばせる（複数選択可）。
     * @returns {File[]} 選んだファイル。キャンセル時は空の配列
     */
    function chooseImageFiles() {
        var fileFilter;
        if ($.os.indexOf("Windows") !== -1) {
            /* Windows はフィルター文字列 / Windows uses a filter string */
            fileFilter = "Images:*." + IMAGE_EXTENSIONS.join(";*.");
        } else {
            /* Mac は関数（フォルダーは辿れるように true） / Mac uses a filter function */
            fileFilter = function (f) {
                return (f instanceof Folder) || isImageName(f.name);
            };
        }
        var chosen = File.openDialog(getLabel(LABELS.dialog.chooseImages), fileFilter, true);
        if (!chosen) return [];
        return (chosen instanceof Array) ? chosen : [chosen];
    }

    /**
     * フォルダーを選ばせ、その中（直下）の画像ファイルを名前順で返す。画像が無ければ知らせる。
     * @returns {File[]} 画像ファイル。キャンセル時・画像が無いときは空の配列
     */
    function chooseImageFolderFiles() {
        var chosenFolder = Folder.selectDialog(getLabel(LABELS.dialog.chooseFolder));
        if (!chosenFolder) return [];
        var imageFiles = listImagesInFolder(chosenFolder);
        if (imageFiles.length === 0) alert(getLabel(LABELS.alert.noImagesInFolder, { name: decodeURI(chosenFolder.name) }));
        return imageFiles;
    }

    /**
     * フォルダーの直下にある画像ファイルを名前順で返す。サブフォルダーと不可視ファイル（.DS_Store など）は除く。
     * @param {Folder} imageFolder - 対象のフォルダー
     * @returns {File[]} 画像ファイル
     */
    function listImagesInFolder(imageFolder) {
        var folderFiles = imageFolder.getFiles(function (f) {
            var fileName = decodeURI(f.name);
            return (f instanceof File) && !f.hidden && fileName.charAt(0) !== "." && isImageName(fileName);
        });
        /* 比較関数つきの sort() は遅く並びも狂うので、「名前\n番号」の文字列を既定の sort() で並べる
           A comparator sort is slow and unreliable, so sort "name\nindex" strings with the default sort() */
        var sortKeys = [];
        for (var i = 0; i < folderFiles.length; i++) sortKeys.push(decodeURI(folderFiles[i].name).toLowerCase() + "\n" + i);
        sortKeys.sort();
        var imageFiles = [];
        for (var j = 0; j < sortKeys.length; j++) imageFiles.push(folderFiles[parseInt(sortKeys[j].split("\n")[1], 10)]);
        return imageFiles;
    }

    /**
     * 画像を配置する。失敗したファイルは failedNames に名前を積む。
     * @param {Document} doc - 配置先のドキュメント
     * @param {File[]} imageFiles - 配置する画像
     * @param {string[]} failedNames - 配置できなかったファイル名を積む配列
     * @returns {PlacedItem[]} 配置したアイテム
     */
    function placeImages(doc, imageFiles, failedNames) {
        var placedItems = [];
        for (var i = 0; i < imageFiles.length; i++) {
            var placedItem = doc.placedItems.add();
            try {
                placedItem.file = imageFiles[i];
                placedItems.push(placedItem);
            } catch (e) {
                placedItem.remove();
                failedNames.push(decodeURI(imageFiles[i].name));
            }
        }
        return placedItems;
    }

    /**
     * 画像を選んで配置し、整列のダイアログボックスへ渡します。/ Place chosen images, then open the distribution dialog.
     * 配置先は作業中のアートボードに固定します（［最背面のオブジェクト］はディム）。選択していた1つのオブジェクトは使いません。
     * キャンセルすると配置した画像を取り除き、元の選択に戻します。
     *
     * @returns {void}
     */
    function placeImagesThenDistribute() {
        var doc = app.activeDocument;
        var activeLayer = doc.activeLayer;
        if (activeLayer.locked || !activeLayer.visible) {
            alert(getLabel(LABELS.alert.layerUnavailable, { name: activeLayer.name }));
            return;
        }

        /* キャンセルで戻す選択（文字の編集中は戻さない） / Selection restored on Cancel (not while editing text) */
        var originalSelection = (doc.selection instanceof Array) ? copyItemsExcept(doc.selection, null) : null;

        var imageSource = showImageSourceDialog();
        if (!imageSource) return;
        var imageFiles = (imageSource === IMAGE_SOURCE_FOLDER) ? chooseImageFolderFiles() : chooseImageFiles();
        if (imageFiles.length === 0) return;

        var failedNames = [];
        var placedItems = placeImages(doc, imageFiles, failedNames);
        if (failedNames.length > 0) {
            alert(getLabel(LABELS.alert.placeFailed, { list: failedNames.join("\n") }));
        }
        if (placedItems.length === 0) return;

        showDistributeDialog({
            items: placedItems,
            artboardOnly: true,
            resizeMode: CONFIG.imageResizeMode,
            removeItemsOnCancel: true,
            selectionOnCancel: originalSelection
        });
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /* マスクのプリセット：名前をキーにしたマスクの設定の集まり（{} の既定値は中身を問わず受け取る）。再起動しても残る
       Mask presets: mask settings keyed by name (the {} default accepts any content), kept across restarts */
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "MaskPresets", "persistent");

    /* 組み込みのプリセット（data に無い項目は今の値のまま。削除できない）。「デフォルト」は起動したときのマスクの設定に戻す（data は null）。
       ほかは枠・角丸だけを変える。長さは pt、色の "black" / "white" はドキュメントのカラーモードで作る
       Built-in presets (keys missing from data keep their current values; cannot be deleted). "Default" restores the mask settings
       at launch (data is null); the others change only the frame and corner radius. Lengths in pt; "black" / "white" follow the document color mode */
    var BUILTIN_PRESETS = [
        { labelKey: "presetDefault", data: null },
        { labelKey: "presetBlackThin", data: { useFrame: true, frameWidthPt: 0.3, frameColor: { model: "black" }, useCornerRadius: false } },
        { labelKey: "presetWhiteRound2", data: { useFrame: true, frameWidthPt: 1, frameColor: { model: "white" }, useCornerRadius: true, cornerRadiusPt: 2 } },
        { labelKey: "presetWhiteRound50", data: { useFrame: true, frameWidthPt: 1, frameColor: { model: "white" }, useCornerRadius: true, cornerRadiusPt: 50 } }
    ];

    /**
     * 保存したプリセットを返す。
     * @returns {Object} プリセット名をキーにした設定の集まり（書き換えても保存されない）
     */
    function loadPresetMap() {
        return presetSettingsStore.load({});
    }

    /**
     * 保存したプリセットの名前を、名前順で返す。
     * @returns {string[]} プリセット名
     */
    function getPresetNames() {
        var presetMap = loadPresetMap();
        var presetNames = [];
        for (var presetName in presetMap) {
            if (presetMap.hasOwnProperty(presetName)) presetNames.push(presetName);
        }
        presetNames.sort();
        return presetNames;
    }

    /**
     * プリセット名を尋ねる。
     * @param {string} initialName - 入力欄に入れておく名前
     * @returns {string|null} プリセット名。キャンセルか空なら null
     */
    function showPresetNameDialog(initialName) {
        var nameDialog = new Window("dialog", getLabel(LABELS.dialog.presetSave));
        setupWindow(nameDialog);
        var nameRow = nameDialog.add("group");
        setupRow(nameRow, "left");
        nameRow.add("statictext", undefined, labelText(LABELS.fieldLabel.presetName));
        var nameInput = nameRow.add("edittext", undefined, initialName);
        nameInput.helpTip = getLabel(LABELS.tooltip.presetName);
        nameInput.characters = PRESET_NAME_CHARS;
        nameInput.active = true;

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(nameDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(nameDialog, SCRIPT_NAME + "_presetName");
        if (nameDialog.show() !== 1) return null;

        var presetName = nameInput.text.replace(/^\s+|\s+$/g, "");
        if (!presetName || presetName === getLabel(LABELS.dropdown.presetPlaceholder)) return null;
        return presetName;
    }

    /**
     * 設定を名前を付けて保存する（同じ名前は確認してから上書き）。
     * @param {string} initialName - 名前の入力欄に入れておく名前
     * @param {Object} presetData - 保存する設定
     * @returns {string|null} 保存したプリセット名。保存しなかったら null
     */
    function savePreset(initialName, presetData) {
        var presetName = showPresetNameDialog(initialName);
        if (!presetName) return null;
        for (var i = 0; i < BUILTIN_PRESETS.length; i++) {
            if (getLabel(LABELS.dropdown[BUILTIN_PRESETS[i].labelKey]) === presetName) {
                alert(getLabel(LABELS.alert.presetNameBuiltin, { name: presetName }));
                return null;
            }
        }
        var presetMap = loadPresetMap();
        if (presetMap.hasOwnProperty(presetName) && !confirm(getLabel(LABELS.confirm.presetOverwrite, { name: presetName }), true)) return null;
        presetMap[presetName] = presetData;
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel(LABELS.alert.presetSaveFailed));
            return null;
        }
        return presetName;
    }

    /**
     * プリセットを確認してから削除する。
     * @param {string} presetName - 削除するプリセット名
     * @returns {boolean} 削除したら true
     */
    function deletePreset(presetName) {
        if (!confirm(getLabel(LABELS.confirm.presetDelete, { name: presetName }), true)) return false;
        var presetMap = loadPresetMap();
        delete presetMap[presetName];
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel(LABELS.alert.presetSaveFailed));
            return false;
        }
        return true;
    }

    /**
     * 色をプリセットに保存できる形にします（RGB・CMYK・グレー。それ以外は null）。/ Serialize a color for a preset.
     *
     * @param {Color} sourceColor - 色。
     * @returns {Object|null} { model: "rgb"|"cmyk"|"gray", values: number[] }。
     */
    function colorToSettings(sourceColor) {
        if (sourceColor.typename === "RGBColor") return { model: "rgb", values: [sourceColor.red, sourceColor.green, sourceColor.blue] };
        if (sourceColor.typename === "CMYKColor") return { model: "cmyk", values: [sourceColor.cyan, sourceColor.magenta, sourceColor.yellow, sourceColor.black] };
        if (sourceColor.typename === "GrayColor") return { model: "gray", values: [sourceColor.gray] };
        return null;
    }

    /**
     * プリセットの色を Color に戻します。読めないときはドキュメントのカラーモードの黒。/ Restore a color from a preset.
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {Object|null} savedColor - colorToSettings() の戻り値、または { model: "black" | "white" }。
     * @returns {Color} 色。
     */
    function colorFromSettings(doc, savedColor) {
        if (savedColor && savedColor.model === "white") return createGrayColor(doc, 0, 255);
        if (!savedColor || !(savedColor.values instanceof Array)) return createGrayColor(doc, 100, 0);
        var values = savedColor.values;
        if (savedColor.model === "rgb" && values.length === 3) {
            var rgbColor = new RGBColor();
            rgbColor.red = values[0];
            rgbColor.green = values[1];
            rgbColor.blue = values[2];
            return rgbColor;
        }
        if (savedColor.model === "cmyk" && values.length === 4) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = values[0];
            cmykColor.magenta = values[1];
            cmykColor.yellow = values[2];
            cmykColor.black = values[3];
            return cmykColor;
        }
        if (savedColor.model === "gray" && values.length === 1) {
            var grayColor = new GrayColor();
            grayColor.gray = values[0];
            return grayColor;
        }
        return createGrayColor(doc, 100, 0);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 画像を読み込んだあとに渡す、整列の起動条件。/ Launch options passed after importing images.
     *
     * @typedef {Object} DistributeLaunchOptions
     * @property {Array<PageItem>} [items] - 整列するオブジェクト。省くと選択中のオブジェクト。
     * @property {boolean} [artboardOnly] - 配置先を現在のアートボードに固定し、［最背面のオブジェクト］をディムにする。
     * @property {string} [resizeMode] - リサイズの初期値（"none" / "width" / "height" / "area" / "scale"）。
     * @property {boolean} [removeItemsOnCancel] - キャンセルしたら items を削除する（読み込んだ画像）。
     * @property {Array<PageItem>} [selectionOnCancel] - キャンセルしたあとに選択するオブジェクト（removeItemsOnCancel のとき）。
     */

    /**
     * ダイアログを表示して配置処理を実行します。/ Show the dialog and run the distribution.
     *
     * @param {DistributeLaunchOptions} [launchOptions] - 画像を読み込んだあとの起動条件。省くと選択中のオブジェクトを整列。
     * @returns {void}
     */
    function showDistributeDialog(launchOptions) {
        var doc = app.activeDocument;
        var launchContext = launchOptions || {};

        /* 前回実行時のセル描画レイヤーは、OK を押すまで消さずに隠しておく。
           ここで削除するとキャンセルしても戻せない
           Hide the cell layer from a previous run instead of deleting it, so Cancel can restore it */
        var previousCellLayer = findLayerByName(doc, CONFIG.cellLayerName);
        var wasPreviousCellLayerVisible = previousCellLayer ? previousCellLayer.visible : false;
        setLayerVisible(previousCellLayer, false);

        /* 以後の計算は「ダイアログ起動時点のアクティブアートボード」を基準に固定 / Fix the base to the artboard active at launch */
        var baseArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var baseArtboardRect = doc.artboards[baseArtboardIndex].artboardRect;

        /* 選択内容はダイアログ起動時に固定する（順序も保持）/ Freeze the selection (and its order) at launch */
        var placeableItems = copyItemsExcept(launchContext.items || doc.selection, null);

        /* 枠の候補は選択の中の最背面。画像を読み込んだときは使わない / The frame candidate is the backmost selected object; never for imported images */
        var backmostItem = launchContext.artboardOnly ? null : findBackmostSelectedItem(placeableItems);

        var rulerUnit = getUnitInfo();
        var defaultDivision = computeDefaultDivision(placeableItems.length, baseArtboardRect, CONFIG.defaultBrick);
        var transparencyGridToggleCount = 0;

        var strokeUnit = getUnitInfo("strokeUnits");
        var dialogUI = buildDistributeDialog(backmostItem !== null, defaultDivision, rulerUnit.label, strokeUnit.label, handleFillStartChange, updatePreview);
        var distributeDialog = dialogUI.distributeDialog;
        var placementUI = dialogUI.placementUI;
        var divisionUI = dialogUI.divisionUI;
        var arrangementUI = dialogUI.arrangementUI;
        var resizeUI = dialogUI.resizeUI;
        var maskUI = dialogUI.maskUI;
        var cellUI = dialogUI.cellUI;
        var itemAnchorUI = dialogUI.itemAnchorUI;
        var footerUI = dialogUI.footerUI;
        var presetUI = maskUI.presetUI;

        /* 起動条件の初期値（画像を読み込んだときはディムの理由とリサイズ） / Initial state from the launch options */
        if (launchContext.artboardOnly) placementUI.backmostRadio.helpTip = getLabel(LABELS.tooltip.targetUnavailableForImages);
        if (launchContext.resizeMode) selectResizeRadio(resizeUI, launchContext.resizeMode);

        /** @type {Array<PageItem>|null} シャッフルで決めた配置順。未使用なら null。 */
        var randomizedOrder = null;
        /** @type {number|null} ［均等にまばらに］のセルの選び方を決める乱数の種。シャッフルするまでは null（決まった選び方）。 */
        var spreadSeed = null;
        /** @type {RadioButton} 「長方形で残す」に戻したときに復帰させるカラー選択。 */
        var lastCellColorRadio = cellUI.blackCellRadio;
        /** @type {Array<PageItem>} プレビュー用に作った複製。 */
        var previewCopies = [];
        /** @type {Array<PageItem>} プレビューの間だけ隠した元のオブジェクト。 */
        var previewHiddenItems = [];
        /** @type {boolean} OK／キャンセルで後始末済みか（onClose の二重実行を防ぐ）。 */
        var isCleanedUp = false;
        /** @type {number} 直近にプレビューを描き直した時刻（ミリ秒）。入力中の間引きに使います。 */
        var lastPreviewTime = 0;
        /** @type {Array<GroupItem>} 直近の描画で作ったクリップグループ。 */
        var maskGroups = [];
        /** @type {string} マスクの種類（MASK_MODE_*）。表示前に代入したラジオ・チェックボックスの value は読み戻せないことがあるので、変数で持ちます。 */
        var maskMode = CONFIG.defaultMaskMode;
        /** @type {Array<number[]>} 直近の描画で使ったマスクの矩形（角丸の自動計算に使う）。 */
        var lastMaskRects = [];
        /** @type {boolean} ［セル間隔を優先］の状態（理由は maskMode と同じ）。 */
        var isGutterPriorityEnabled = CONFIG.defaultGutterPriority;
        /** @type {GridMetrics|null} 直近の描画で使ったグリッド（［セル間隔を優先］で作り直したあとのもの）。 */
        var lastGridMetrics = null;
        /** @type {boolean} ［均等にまばらに］の状態（理由は maskMode と同じ）。 */
        var isSpreadEnabled = CONFIG.defaultSpreadEvenly;
        /** @type {boolean} 直近に行数・列数を計算したときの配置先が［最背面のオブジェクト］か（同じラジオを押し直したときに入力を上書きしない）。 */
        var wasBackmostTarget = false;
        /** @type {boolean} ［レンガ］の状態（理由は maskMode と同じ）。 */
        var isBrickEnabled = CONFIG.defaultBrick;
        /** @type {boolean} ［角丸］の状態（理由は maskMode と同じ）。 */
        var isCornerRadiusEnabled = CONFIG.defaultUseCornerRadius;
        /** @type {boolean} ［枠］の状態（理由は maskMode と同じ）。 */
        var isFrameEnabled = CONFIG.defaultUseFrame;
        /** @type {CMYKColor|RGBColor} 枠の色（標準のカラーピッカーで選ぶ）。 */
        var frameColor = createGrayColor(doc, 100, 0);
        /** @type {boolean} ［元のオブジェクトを残す］の状態（理由は maskMode と同じ）。 */
        var isKeepOriginalsEnabled = CONFIG.defaultKeepOriginals;
        /** @type {Array<PageItem>} ［元のオブジェクトを残す］で、本番の配置に作った複製。 */
        var keptCopies = [];

        bindDialogEvents();
        wireNumericFields();

        /* 初期状態を反映してプレビューを表示 / Apply the initial state and show the preview */
        syncAllControls();
        updatePreview();
        var fitViewControls = footerUI.fitViewControls;

        /* 画面にフィット：開いたときの表示を控えてから合わせる（ダイアログの位置はこのあとの表示で決める）
           Fit to Window: capture the view first, then fit (the dialog position is decided on the fitted view) */
        var initialViewState = FitViewToItems.captureView(doc);
        fitViewToTarget();

        /* 組み込みの［デフォルト］で戻す、起動したときのマスクの設定 / Mask settings at launch, restored by the built-in Default preset */
        var launchMaskSettings = collectMaskSettings();
        bindPresetEvents();
        fillPresetDropdown(null);

        /* バージョン表示は layout の前に足す。後から足すとウィンドウの大きさが決まったあとで、行の場所が無く表示されない
           Add the footer before layout; added afterwards it gets no room in the already-sized window */
        addVersionFooter(distributeDialog);
        distributeDialog.layout.layout(true);
        trimButtonHeight(cellUI.transparencyGridButton, 4);
        prepareDialogWindow(distributeDialog, SCRIPT_NAME);
        distributeDialog.show();

        /* 透明グリッドの表示状態を元へ戻す / Restore the transparency grid */
        if (transparencyGridToggleCount % 2 !== 0) {
            app.executeMenuCommand('TransparencyGrid Menu Item');
        }

        // -----------------------------------------
        // 入力値の取得 / Input readers
        // -----------------------------------------

        /**
         * 現在のセル描画モードを返します。/ Return the current cell drawing mode.
         *
         * @returns {string} "none"、"keep"、"guide"、"artboard" のいずれか。
         */
        function getCellMode() {
            if (cellUI.noCellRadio.value) return "none";
            if (cellUI.toArtboardRadio.value) return "artboard";
            if (cellUI.toGuideRadio.value) return "guide";
            return "keep";
        }

        /**
         * 現在のリサイズの基準を返します。/ Return the current resize mode.
         *
         * @returns {string} "none"、"width"、"height"、"area"、"scale" のいずれか。
         */
        function getResizeMode() {
            if (resizeUI.widthRadio.value) return "width";
            if (resizeUI.heightRadio.value) return "height";
            if (resizeUI.areaRadio.value) return "area";
            if (resizeUI.scaleRadio.value) return "scale";
            return "none";
        }

        /**
         * 角丸と枠の指定を読み取ります（オフ・読めない値は 0）。/ Read the corner radius and frame settings.
         *
         * @returns {MaskStyle} 角丸と枠の指定。
         */
        function readMaskStyle() {
            return {
                cornerRadius: isCornerRadiusEnabled ? readLengthPt(maskUI.cornerRadiusInput, rulerUnit) : 0,
                frameWidth: isFrameEnabled ? readLengthPt(maskUI.frameWidthInput, strokeUnit) : 0,
                frameColor: frameColor
            };
        }

        /**
         * 選択中の対象領域の矩形を返します。/ Return the rect of the selected target area.
         *
         * @returns {number[]} 配置先の矩形 [左, 上, 右, 下]。
         */
        function getTargetRect() {
            if (placementUI.backmostRadio.value && backmostItem) return backmostItem.geometricBounds;
            if (doc.artboards.length > baseArtboardIndex) return doc.artboards[baseArtboardIndex].artboardRect;
            return baseArtboardRect;
        }

        /**
         * 入力欄からグリッドの寸法を求めます。/ Compute the grid metrics from the fields.
         *
         * @returns {GridMetrics|null} グリッドの寸法。入力が無効な場合は null。
         */
        function readGridMetrics() {
            return computeGridMetrics(
                readCount(divisionUI.rowCountInput),
                readCount(divisionUI.columnCountInput),
                readLengthPt(divisionUI.marginInput, rulerUnit),
                readLengthPt(divisionUI.gutterInput, rulerUnit),
                getTargetRect(),
                isBrickEnabled
            );
        }

        /**
         * セルの塗り色を返します。/ Return the cell fill color.
         *
         * @returns {CMYKColor|RGBColor|null} 塗り色。透過を選んでいる場合は null。
         */
        function getCellFillColor() {
            if (cellUI.blackCellRadio.value) return createGrayColor(doc, 100, 0);
            if (cellUI.whiteCellRadio.value) return createGrayColor(doc, 0, 255);
            return null;
        }

        // -----------------------------------------
        // 配置対象 / Distribution items
        // -----------------------------------------

        /**
         * 枠として使うアイテムを除いた配置対象を返します。/ Return the items to place, excluding frame items.
         *
         * @returns {Array<PageItem>} 配置対象のオブジェクト。
         */
        function getDistributionItems() {
            if (!(placementUI.backmostRadio.value && backmostItem)) return placeableItems;
            return copyItemsExcept(placeableItems, backmostItem);
        }

        // -----------------------------------------
        // 描画と配置 / Drawing and placement
        // -----------------------------------------

        /**
         * セル描画と配置を実行します（プレビュー／本番共通）。/ Draw cells and place objects.
         * プレビューでは元のオブジェクトを複製して元を隠し、複製だけを動かします（元の位置・大きさには触れません）。
         * ［元のオブジェクトを残す］では、本番でも複製を動かし、元は隠さずその場に残します。
         *
         * @param {boolean} isPreview - プレビューとして描画する場合は true。
         * @returns {GridMetrics|null} 使ったグリッド（［セル間隔を優先］では作り直したもの）。入力が無効で何もしなかった場合は null。
         */
        function renderDistribution(isPreview) {
            var cellMode = getCellMode();
            var gridMetrics = readGridMetrics();

            /* グリッドが成立しない入力では、ドキュメントに一切手を加えない / Leave the document untouched without a valid grid */
            if (!gridMetrics) return null;

            var cellCount = countCells(gridMetrics);
            var distributionItems = prepareDistributionItems(cellCount, isPreview);

            /* セル数を超えたオブジェクトだけを対象領域の外へ退避する / Park only the objects beyond the cell count */
            var overflowItems = distributionItems.slice(cellCount);
            if (overflowItems.length > 0) {
                parkItemsOutside(doc, overflowItems, gridMetrics.targetRect);
            }

            /* セルに入るオブジェクトだけをリサイズしてから中央へ置く / Resize only the objects that go into cells, then center them */
            var cellItems = distributionItems.slice(0, cellCount);
            resizeItemsForCells(cellItems, gridMetrics, getResizeMode(), readClampedNumber(resizeUI.ratioInput, 1, CONFIG.maxResizeRatio));
            /* ［セル間隔を優先］では、リサイズ後のオブジェクトの大きさでセルを作り直す / Prioritize Gutter rebuilds the cells from the resized objects */
            if (isGutterPriorityEnabled) gridMetrics = fitGridToItems(gridMetrics, cellItems);
            lastGridMetrics = gridMetrics;

            drawCellLayer(gridMetrics, cellMode, isPreview);
            var cellRects = listTargetCellRects(gridMetrics, cellItems.length);
            placeItemsInCells(cellItems, cellRects, getAnchorWidgetIndex(itemAnchorUI.itemAnchorWidget));

            /* マスクはセルに入れたものだけ。プレビューのグループは複製と一緒に消す
               Mask only the objects in cells; preview groups are removed along with the copies */
            maskGroups = [];
            lastMaskRects = [];
            if (maskMode !== MASK_MODE_NONE) {
                lastMaskRects = listMaskRects(cellItems, cellRects, maskMode);
                maskItemsToRects(cellItems, lastMaskRects, isPreview ? previewCopies : maskGroups, readMaskStyle());
            }
            return gridMetrics;
        }

        /**
         * 動かすオブジェクトを用意します（元そのもの・プレビュー用の複製・残す場合の複製）。/ Prepare the objects to move.
         *
         * @param {number} cellCount - セルの数。
         * @param {boolean} isPreview - プレビューなら true。
         * @returns {Array<PageItem>} 動かすオブジェクト（配置順）。
         */
        function prepareDistributionItems(cellCount, isPreview) {
            var sourceItems = orderByRandomizedOrder(getDistributionItems(), randomizedOrder);
            if (!isKeepOriginalsEnabled) return isPreview ? createPreviewCopies(sourceItems) : sourceItems;

            /* 元はその場に残し、セルに入る分だけ複製する（あふれた分は複製も退避もしない）
               Leave the originals in place and copy only what fits in the cells (nothing to park) */
            if (!isPreview) keptCopies = [];
            return duplicateItems(sourceItems.slice(0, cellCount), isPreview ? previewCopies : keptCopies);
        }

        /**
         * セルの扱いに合わせてセルを描きます。/ Draw the cells for the current cell handling.
         * ［アートボード化］［残さない］は本番では描かず、プレビューでだけセルの範囲をガイドで示します。
         *
         * @param {GridMetrics} gridMetrics - グリッドの寸法。
         * @param {string} cellMode - getCellMode() の値。
         * @param {boolean} isPreview - プレビューなら true。
         * @returns {void}
         */
        function drawCellLayer(gridMetrics, cellMode, isPreview) {
            var drawsNoCells = (cellMode === "artboard" || cellMode === "none");
            if (drawsNoCells && !isPreview) return;
            var layerName = isPreview ? CONFIG.previewCellLayerName : CONFIG.cellLayerName;
            var asGuide = (cellMode === "guide") || drawsNoCells;
            drawCells(getOrCreateLayer(doc, layerName), gridMetrics, asGuide, getCellFillColor(), readClampedNumber(cellUI.opacityInput, 0, 100));
        }

        /**
         * オブジェクトを入れるセルを、入れる順に返します（開始位置と［均等にまばらに］を反映）。/ List the cells to fill, in order.
         *
         * @param {GridMetrics} gridMetrics - グリッドの寸法。
         * @param {number} itemCount - セルに入れるオブジェクトの数。
         * @returns {Array<number[]>} セルの矩形 [左, 上, 右, 下] の配列。
         */
        function listTargetCellRects(gridMetrics, itemCount) {
            var cellRects = listCellRectsInFillOrder(gridMetrics, getAnchorWidgetIndex(arrangementUI.fillStartWidget));
            /* セルが余るときは、使うセルをグリッド全体に散らす / With spare cells, spread the used cells over the grid */
            if (!isSpreadEnabled || itemCount >= cellRects.length) return cellRects;
            var spreadRandom = (spreadSeed === null) ? null : createSeededRandom(spreadSeed);
            return pickSpreadCells(cellRects, itemCount, spreadRandom);
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------
        /* プレビューは app.undo() で巻き戻さず、複製＋一時非表示で作る。
           元のオブジェクトを複製してから元を隠し、動かすのは複製だけ。消すときは複製を削除して元を表示に戻す。
           app.undo() は1プレビューが何段の履歴になるかで戻る量が狂い、スクリプト実行前の作業まで取り消してしまうため。
           Previews duplicate the objects and hide the originals instead of relying on app.undo(), whose step count
           drifts and can revert the user's own work from before the script ran. */

        /**
         * 元のオブジェクトを複製し、元を隠します。/ Duplicate the objects and hide the originals.
         * duplicate() は hidden を引き継ぐので、必ず複製してから隠します。
         *
         * @param {Array<PageItem>} sourceItems - 元のオブジェクト（配置順）。
         * @returns {Array<PageItem>} 複製（元と同じ順）。
         */
        function createPreviewCopies(sourceItems) {
            var copies = duplicateItems(sourceItems, previewCopies);
            for (var i = 0; i < sourceItems.length; i++) {
                sourceItems[i].hidden = true;
                previewHiddenItems.push(sourceItems[i]);
            }
            return copies;
        }

        /**
         * プレビューを消して元の状態に戻します。/ Discard the preview and restore the original state.
         *
         * @returns {void}
         */
        function clearPreview() {
            for (var i = 0; i < previewCopies.length; i++) {
                try {
                    previewCopies[i].remove();
                } catch (e) {
                    /* 参照が失われた複製は飛ばす / Skip copies whose reference went stale */
                }
            }
            for (var j = 0; j < previewHiddenItems.length; j++) {
                setItemHidden(previewHiddenItems[j], false);
            }
            previewCopies = [];
            previewHiddenItems = [];
            removeLayerByName(doc, CONFIG.previewCellLayerName);
            removeLayerByName(doc, CONFIG.legacyPreviewLayerName);
        }

        /**
         * 現在の設定でプレビューを描き直します。/ Redraw the preview with the current settings.
         *
         * @returns {void}
         */
        function updatePreview() {
            syncSpreadEvenlyEnabled();
            try {
                clearPreview();
                renderDistribution(true);
            } catch (e) {
                clearPreview();
            }
            app.redraw();
            lastPreviewTime = (new Date()).getTime();
        }

        /**
         * 入力中のプレビューを更新します（重いグリッドでは打鍵を間引きます）。
         * Refresh the preview while typing, throttling only when each redraw is slow.
         *
         * セル数が多いと1回の描き直しで数百〜数千の長方形を作り直すため、
         * 「100」と打つだけで同じ処理が3回走る。見送った分は入力欄の確定（onChange）で
         * 必ず描き直されるので、表示が取り残されたままにはならない。
         * A skipped redraw is always flushed when the field commits, so the preview never stays stale.
         *
         * @returns {void}
         */
        function updatePreviewWhileTyping() {
            var rowCount = readCount(divisionUI.rowCountInput);
            var columnCount = readCount(divisionUI.columnCountInput);
            var cellCount = (rowCount !== null && columnCount !== null) ? rowCount * columnCount : 0;

            /* 軽いグリッドは打鍵ごとに描き直す / Small grids stay fully live */
            if (cellCount <= CONFIG.previewThrottleCells) {
                updatePreview();
                return;
            }
            if ((new Date()).getTime() - lastPreviewTime < CONFIG.previewThrottleMs) return;
            updatePreview();
        }

        // -----------------------------------------
        // UI 状態の同期 / UI state
        // -----------------------------------------

        /**
         * 数値欄に範囲の正規化と、確定後の処理を結び付けます。/ Wire range clamping and commit handlers into the fields.
         *
         * @returns {void}
         */
        function wireNumericFields() {
            /**
             * 行数・列数を変えたら、セル間隔の有効・無効も合わせてプレビューを描き直す
             * @returns {void}
             */
            function handleDivisionChange() {
                syncGutterEnabled();
                updatePreview();
            }

            /* 行数・列数は 1 以上、上限はグリッドが成立する範囲まで / Rows and columns stay within a workable grid */
            enableNumericField(divisionUI.rowCountInput, 1, CONFIG.maxDivision, handleDivisionChange);
            enableNumericField(divisionUI.columnCountInput, 1, CONFIG.maxDivision, handleDivisionChange);
            /* 間隔・マージンは負にしない（配置先の外へはみ出す）/ Gutter and margin never go negative */
            enableNumericField(divisionUI.gutterInput, 0, null, updatePreview);
            enableNumericField(divisionUI.marginInput, 0, null, updatePreview);
            /* 不透明度は 0〜100%（描画側のクランプと表示を一致させる）/ Opacity matches what is actually drawn */
            enableNumericField(cellUI.opacityInput, 0, 100, updatePreview);
            /* リサイズの比率は 1〜上限% / The resize ratio stays within 1% and the upper limit */
            enableNumericField(resizeUI.ratioInput, 1, CONFIG.maxResizeRatio, updatePreview);
            /* 角丸の半径・枠の線幅は負にしない / Corner radius and frame weight never go negative */
            enableNumericField(maskUI.cornerRadiusInput, 0, null, updatePreview);
            enableNumericField(maskUI.frameWidthInput, 0, null, updatePreview);
        }

        /**
         * すべての項目の有効・無効を、今の設定に合わせます。/ Sync every control's enabled state.
         *
         * @returns {void}
         */
        function syncAllControls() {
            syncGutterEnabled();
            syncCellModeUI();
            syncResizeRatioEnabled();
            syncMaskStyleEnabled();
        }

        /**
         * 間隔欄の有効・無効を切り替えます。/ Enable or disable the gutter field.
         * 間隔は行・列いずれかが2以上のときだけ意味を持ちます。
         *
         * @returns {void}
         */
        function syncGutterEnabled() {
            var rowCount = readCount(divisionUI.rowCountInput);
            var columnCount = readCount(divisionUI.columnCountInput);

            /* 入力途中で両方が空のときは、現在の有効・無効を保つ / Keep the state while both fields are empty */
            if (rowCount === null && columnCount === null) return;
            setSteppedInputEnabled(divisionUI.gutterInput, rowCount > 1 || columnCount > 1);
        }

        /**
         * カラー選択を不透明度欄に反映します。/ Reflect the color choice in the opacity field.
         *
         * @param {boolean} resetsValue - カラーごとの既定値で上書きする場合は true。
         * @returns {void}
         */
        function syncOpacityEnabled(resetsValue) {
            setSteppedInputEnabled(cellUI.opacityInput, cellUI.blackCellRadio.value || cellUI.whiteCellRadio.value);
            if (!resetsValue) return;
            if (cellUI.blackCellRadio.value) {
                cellUI.opacityInput.text = String(CONFIG.blackCellOpacity);
            } else if (cellUI.whiteCellRadio.value) {
                cellUI.opacityInput.text = String(CONFIG.whiteCellOpacity);
            }
        }

        /**
         * セル描画モードに応じてUIを整えます。/ Update the UI for the current cell drawing mode.
         *
         * @returns {void}
         */
        function syncCellModeUI() {
            var usesFill = (getCellMode() === "keep");
            cellUI.blackCellRadio.enabled = usesFill;
            cellUI.whiteCellRadio.enabled = usesFill;
            cellUI.transparentCellRadio.enabled = usesFill;

            if (usesFill) {
                /* 「塗りなし」固定から、直前に選んでいたカラーへ戻す / Return to the color chosen before */
                lastCellColorRadio.value = true;
                syncOpacityEnabled(false);
                return;
            }
            /* ガイド化／アートボード化では塗りを持たないため塗りなしに固定 / Guides and artboards have no fill */
            cellUI.transparentCellRadio.value = true;
            setSteppedInputEnabled(cellUI.opacityInput, false);
        }

        /**
         * 比率欄の有効・無効を、リサイズの基準に合わせます。/ Enable the ratio field only when resizing.
         *
         * @returns {void}
         */
        function syncResizeRatioEnabled() {
            var isResizing = (getResizeMode() !== "none");
            resizeUI.ratioLabel.enabled = isResizing;
            resizeUI.ratioUnitLabel.enabled = isResizing;
            setSteppedInputEnabled(resizeUI.ratioInput, isResizing);
            setIconButtonEnabled(resizeUI.autoScaleButton, getResizeMode() === "scale");
        }

        /**
         * ［均等にまばらに］を、セルが余る（オブジェクトがセルより少ない）ときだけ有効にします。/ Enable Spread Evenly only with spare cells.
         * 描き直しのたびに呼ぶので、行数・列数・配置先・［レンガ］のどれを変えても追従します。
         *
         * @returns {void}
         */
        function syncSpreadEvenlyEnabled() {
            var gridMetrics = readGridMetrics();
            arrangementUI.spreadEvenlyCheckbox.enabled = !!gridMetrics && getDistributionItems().length < countCells(gridMetrics);
        }

        /**
         * 開始位置を選んだとき：option＋クリックなら［均等にまばらに］をオフにしてから描き直します。/ Handle a Start From click.
         * ウィジェットの onChange はクリックの中で呼ばれるので、そのときの修飾キーを読みます。
         *
         * @returns {void}
         */
        function handleFillStartChange() {
            if (isSpreadEnabled && ScriptUI.environment.keyboardState.altKey) {
                isSpreadEnabled = false;
                arrangementUI.spreadEvenlyCheckbox.value = false;
            }
            updatePreview();
        }

        /**
         * マスクの種類を切り替えて、角丸・枠の有効・無効とプレビューを更新します。/ Switch the mask mode.
         *
         * @param {string} nextMaskMode - MASK_MODE_NONE / MASK_MODE_CELL / MASK_MODE_ITEM。
         * @returns {void}
         */
        function selectMaskMode(nextMaskMode) {
            var previousMaskMode = maskMode;
            maskMode = nextMaskMode;
            syncMaskStyleEnabled();
            clearCellsWhenMaskTurnsOn(previousMaskMode);
            updatePreview();
        }

        /**
         * マスクを［なし］から［セル］／［画像］にしたとき、セルの扱いを［残さない］にします（マスクで見えるのでセルの長方形は要らない）。
         * When the mask goes from None to Cell or Image, switch Cell Handling to Don't Keep; the mask makes the cell rectangles unnecessary.
         * ［セル］と［画像］の切り替えでは変えないので、あとから選び直したセルの扱いはそのまま残ります。
         *
         * @param {string} previousMaskMode - 切り替える前のマスクの種類（MASK_MODE_*）。
         * @returns {void}
         */
        function clearCellsWhenMaskTurnsOn(previousMaskMode) {
            if (previousMaskMode !== MASK_MODE_NONE || maskMode === MASK_MODE_NONE) return;
            if (cellUI.noCellRadio.value) return;
            selectCellModeRadio(cellUI, "none");
            syncCellModeUI();
        }

        /**
         * 角丸・枠の項目を、［マスクする］と各チェックボックスに合わせてディムします。/ Dim the corner and frame controls.
         *
         * @returns {void}
         */
        function syncMaskStyleEnabled() {
            var isMaskEnabled = (maskMode !== MASK_MODE_NONE);
            maskUI.cornerRadiusCheckbox.enabled = isMaskEnabled;
            maskUI.frameCheckbox.enabled = isMaskEnabled;
            var usesCornerRadius = isMaskEnabled && isCornerRadiusEnabled;
            setSteppedInputEnabled(maskUI.cornerRadiusInput, usesCornerRadius);
            maskUI.cornerRadiusUnitLabel.enabled = usesCornerRadius;
            setIconButtonEnabled(maskUI.autoCornerRadiusButton, usesCornerRadius);
            var usesFrame = isMaskEnabled && isFrameEnabled;
            setSteppedInputEnabled(maskUI.frameWidthInput, usesFrame);
            maskUI.frameUnitLabel.enabled = usesFrame;
            redrawFrameColorSwatch();
        }

        /**
         * 枠の色見本を描き直します（group には notify() が無いので隠して再表示する）。/ Redraw the frame color swatch.
         *
         * @returns {void}
         */
        function redrawFrameColorSwatch() {
            maskUI.frameColorSwatch.hide();
            maskUI.frameColorSwatch.show();
        }

        /**
         * 枠の色見本を描きます。使えないときは薄く描きます（自作描画は enabled で淡色にならない）。/ Draw the frame color swatch.
         *
         * @returns {void}
         */
        function drawFrameColorSwatch() {
            var swatch = maskUI.frameColorSwatch;
            var swatchGraphics = swatch.graphics;
            var swatchWidth = swatch.size[0];
            var swatchHeight = swatch.size[1];
            var swatchAlpha = (maskMode !== MASK_MODE_NONE && isFrameEnabled) ? 1 : 0.25;
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, colorToScreenRgb(frameColor).concat(swatchAlpha)));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, swatchAlpha], 1));
        }

        /**
         * 行数・列数の自動計算のアイコン：今の配置先・オブジェクトの数・［レンガ］から行数・列数を計算し直して描き直します。/ Recalculate rows and columns.
         * 配置先・［レンガ］を切り替えたときにも呼びます。
         *
         * @returns {void}
         */
        function applyAutoDivision() {
            wasBackmostTarget = placementUI.backmostRadio.value;
            var division = computeDefaultDivision(getDistributionItems().length, getTargetRect(), isBrickEnabled);
            divisionUI.rowCountInput.text = String(division.rowCount);
            divisionUI.columnCountInput.text = String(division.columnCount);
            syncGutterEnabled();
            updatePreview();
        }

        /**
         * スケールの自動計算：セルに入るオブジェクトがすべてセルを覆う（アキが出ない）スケールを比率に入れます。/ Fill the ratio with a covering scale.
         *
         * @returns {void}
         */
        function applyAutoScale() {
            var gridMetrics = readGridMetrics();
            if (!gridMetrics) {
                alert(getLabel(LABELS.alert.invalidGrid));
                return;
            }
            var cellCount = countCells(gridMetrics);
            var cellItems = orderByRandomizedOrder(getDistributionItems(), randomizedOrder).slice(0, cellCount);
            var coverScalePct = computeCoverScalePct(cellItems, gridMetrics);
            if (coverScalePct === null) return;
            resizeUI.ratioInput.text = String(coverScalePct);
            /* 範囲の正規化とプレビューは onChange に任せる（代入では呼ばれない）/ Let onChange clamp and redraw; assignment does not fire it */
            resizeUI.ratioInput.onChange();
        }

        /**
         * 角丸の自動計算：半径をセルの短い辺の半分にします（［画像］ではいちばん小さいマスクの短い辺の半分）。/ Set the corner radius to half the shorter cell (or smallest mask) side.
         * 定規の単位に換算し、大きくならないよう 0.01 単位で切り捨てます。
         *
         * @returns {void}
         */
        function applyAutoCornerRadius() {
            /* ［セル間隔を優先］ではセルが作り直されるので、直近の描画のグリッドを使う / Prioritize Gutter rebuilds the cells, so use the last drawn grid */
            var gridMetrics = (isGutterPriorityEnabled && lastGridMetrics) ? lastGridMetrics : readGridMetrics();
            if (!gridMetrics) {
                alert(getLabel(LABELS.alert.invalidGrid));
                return;
            }
            var radiusPt = Math.min(gridMetrics.cellWidth, gridMetrics.cellHeight) / 2;
            /* ［画像］では、直近の描画のマスクのうちいちばん短い辺の半分 / With Image, half the shortest side among the last masks */
            if (maskMode === MASK_MODE_ITEM && lastMaskRects.length > 0) {
                radiusPt = Infinity;
                for (var i = 0; i < lastMaskRects.length; i++) {
                    var maskRect = lastMaskRects[i];
                    radiusPt = Math.min(radiusPt, (maskRect[2] - maskRect[0]) / 2, (maskRect[1] - maskRect[3]) / 2);
                }
            }
            maskUI.cornerRadiusInput.text = String(Math.floor(ptToUnit(radiusPt, rulerUnit) * 100) / 100);
            /* 範囲の正規化とプレビューは onChange に任せる（代入では呼ばれない）/ Let onChange clamp and redraw; assignment does not fire it */
            maskUI.cornerRadiusInput.onChange();
        }

        // -----------------------------------------
        // ボタン / Buttons
        // -----------------------------------------

        /**
         * 配置順をシャッフルしてプレビューを描き直します。/ Shuffle the placement order and redraw.
         *
         * @returns {void}
         */
        function shuffleDistributionOrder() {
            var distributionItems = getDistributionItems();
            if (distributionItems.length === 0) {
                alert(getLabel(LABELS.alert.noSelection));
                return;
            }
            randomizedOrder = createShuffledCopy(distributionItems);
            /* ［均等にまばらに］で使うセルの選び方も選び直す / Also re-pick the cells for Spread Evenly */
            spreadSeed = Math.floor(Math.random() * 4294967296);
            updatePreview();
        }

        /**
         * OK：プレビューを破棄して本番の配置を行い、ダイアログを閉じます。/ Commit the distribution and close.
         *
         * @returns {void}
         */
        function commitDistribution() {
            /* グリッドが成立しない入力では、閉じずに理由を知らせる / Explain instead of closing on an invalid grid */
            if (!readGridMetrics()) {
                alert(getLabel(LABELS.alert.invalidGrid));
                return;
            }

            var cellMode = getCellMode();
            var createdCount = -1;

            /* プレビューの複製を消し、元のオブジェクトを表示に戻してから本番処理へ / Drop the preview copies and show the originals before the real run */
            clearPreview();

            /* 前回実行時のセル描画レイヤーは、実際に描き直すこの時点で捨てる
               Discard the previous run's cell layer only now that we are really redrawing */
            removeLayerByName(doc, CONFIG.cellLayerName);

            var renderedGridMetrics = renderDistribution(false);
            /* アートボードは、実際に使ったグリッドのセルから作る（1枚ずつの失敗は createArtboardsFromCells() の中で飛ばす）
               Artboards come from the grid actually used; single failures are skipped inside createArtboardsFromCells() */
            if (cellMode === "artboard") createdCount = createArtboardsFromCells(doc, renderedGridMetrics, baseArtboardIndex);

            /* セル描画を残した場合はその長方形を、それ以外は配置したオブジェクト（元を残したときは複製）を選択状態にする
               Select the kept cells, or the placed objects (the copies when the originals are kept) */
            if (cellMode !== "keep" || !selectCellRectangles(doc)) {
                doc.selection = replaceMaskedItems(isKeepOriginalsEnabled ? keptCopies : placeableItems, maskGroups);
            }

            isCleanedUp = true;
            app.redraw();
            distributeDialog.close(1);

            /* 0 個は1枚も作れなかったとき / Zero means not a single artboard could be created */
            if (createdCount === 0) {
                alert(getLabel(LABELS.alert.artboardError));
            } else if (createdCount > 0) {
                alert(getLabel(LABELS.alert.artboardCreated, [createdCount]));
            }
        }

        /**
         * プレビューを消し、ダイアログ表示前の状態へ戻します。/ Discard the preview and restore the pre-dialog state.
         *
         * @returns {void}
         */
        function discardDialogChanges() {
            clearPreview();
            /* ［画面にフィット］で変えた表示を戻す / Undo Fit to Window */
            FitViewToItems.restoreView(initialViewState, doc);
            /* 起動時に隠した前回のセル描画レイヤーを元へ戻す / Restore the cell layer hidden at launch */
            setLayerVisible(previousCellLayer, wasPreviousCellLayerVisible);
            /* 読み込んだ画像は取り除く / Remove the imported images */
            if (launchContext.removeItemsOnCancel) {
                /* このスクリプトが配置したばかりで、ロックも非表示もない / Placed by this script just now; never locked or hidden */
                for (var i = placeableItems.length - 1; i >= 0; i--) placeableItems[i].remove();
            }
            /* プレビュー中に隠したことで外れた選択を戻す / Hiding the originals for the preview cleared the selection */
            try {
                doc.selection = launchContext.removeItemsOnCancel ? (launchContext.selectionOnCancel || null) : placeableItems;
            } catch (e) {
                /* 選択を戻せなくても閉じる処理は続ける / Keep closing even if the selection cannot be restored */
            }
            app.redraw();
        }

        // -----------------------------------------
        // プリセット / Presets
        // -----------------------------------------

        /**
         * 長さ（pt）を単位に換算して入力欄に入れる文字列にします。/ Format a length in points for a field.
         *
         * @param {number} lengthPt - 長さ（pt）。
         * @param {{label: string, pointsPerUnit: number}} unit - 単位（getUnitInfo() の戻り値）。
         * @returns {string} 小数第3位までの数値の文字列。
         */
        function formatLengthText(lengthPt, unit) {
            return String(Math.round(ptToUnit(lengthPt, unit) * 1000) / 1000);
        }

        /**
         * 今のマスクの設定をまとめます（プリセットの保存に使う。長さは pt）。/ Collect the current mask settings for a preset.
         *
         * @returns {Object} マスクの種類・角丸・枠の設定。
         */
        function collectMaskSettings() {
            return {
                maskMode: maskMode,
                useCornerRadius: isCornerRadiusEnabled,
                cornerRadiusPt: readLengthPt(maskUI.cornerRadiusInput, rulerUnit),
                useFrame: isFrameEnabled,
                frameWidthPt: readLengthPt(maskUI.frameWidthInput, strokeUnit),
                frameColor: colorToSettings(frameColor)
            };
        }

        /**
         * マスクの設定を各項目とプレビューに反映します（プリセットの読み込み用）。無い項目・型の合わない項目は今の値のままにします。
         * Apply mask settings to the controls and the preview; missing or mistyped keys keep their current values.
         *
         * @param {Object} presetData - collectMaskSettings() と同じ形の設定（一部だけでもよい）。
         * @param {string} previousMaskMode - 読み込む前のマスクの種類（［なし］からマスクを付けたらセルの扱いを［残さない］にする）。
         * @returns {void}
         */
        function applyMaskSettings(presetData, previousMaskMode) {
            var hasValue = function (key, typeName) {
                return presetData.hasOwnProperty(key) && typeof presetData[key] === typeName;
            };
            if (hasValue("maskMode", "string") && (presetData.maskMode === MASK_MODE_NONE || presetData.maskMode === MASK_MODE_CELL || presetData.maskMode === MASK_MODE_ITEM)) {
                maskMode = presetData.maskMode;
            }
            selectMaskModeRadio(maskUI, maskMode);
            if (hasValue("useCornerRadius", "boolean")) {
                isCornerRadiusEnabled = presetData.useCornerRadius;
                maskUI.cornerRadiusCheckbox.value = isCornerRadiusEnabled;
            }
            if (hasValue("cornerRadiusPt", "number")) maskUI.cornerRadiusInput.text = formatLengthText(presetData.cornerRadiusPt, rulerUnit);
            if (hasValue("useFrame", "boolean")) {
                isFrameEnabled = presetData.useFrame;
                maskUI.frameCheckbox.value = isFrameEnabled;
            }
            if (hasValue("frameWidthPt", "number")) maskUI.frameWidthInput.text = formatLengthText(presetData.frameWidthPt, strokeUnit);
            if (presetData.hasOwnProperty("frameColor")) frameColor = colorFromSettings(doc, presetData.frameColor);
            syncMaskStyleEnabled();
            clearCellsWhenMaskTurnsOn(previousMaskMode);
            updatePreview();
        }

        /**
         * 選んでいる保存したプリセットの名前を返します。/ Return the selected saved preset name.
         *
         * @returns {string|null} プリセット名。「---」・組み込みのプリセットなら null。
         */
        function getSelectedPresetName() {
            var presetSelection = presetUI.presetDropdown.selection;
            return (presetSelection && presetSelection.index > BUILTIN_PRESETS.length) ? presetSelection.text : null;
        }

        /**
         * 選んでいる組み込みのプリセットを返します。/ Return the selected built-in preset.
         *
         * @returns {Object|null} BUILTIN_PRESETS の要素。選んでいなければ null。
         */
        function getSelectedBuiltinPreset() {
            var presetSelection = presetUI.presetDropdown.selection;
            if (!presetSelection || presetSelection.index < 1 || presetSelection.index > BUILTIN_PRESETS.length) return null;
            return BUILTIN_PRESETS[presetSelection.index - 1];
        }

        /* ドロップダウンの作り直しでも onChange が来るので、作り直し中は読み込まない / Refilling also fires onChange, so ignore it meanwhile */
        var isFillingPresets = false;

        /**
         * プリセットのドロップダウンを作り直し、指定の名前を選びます。/ Refill the preset dropdown.
         *
         * @param {string|null} selectedName - 選んでおくプリセット名（null なら「---」）。
         * @returns {void}
         */
        function fillPresetDropdown(selectedName) {
            isFillingPresets = true;
            var presetDropdown = presetUI.presetDropdown;
            var presetNames = getPresetNames();
            var selectedIndex = 0;
            presetDropdown.removeAll();
            presetDropdown.add("item", getLabel(LABELS.dropdown.presetPlaceholder));
            for (var j = 0; j < BUILTIN_PRESETS.length; j++) {
                presetDropdown.add("item", getLabel(LABELS.dropdown[BUILTIN_PRESETS[j].labelKey]));
            }
            if (presetNames.length > 0) presetDropdown.add("separator");
            for (var i = 0; i < presetNames.length; i++) {
                presetDropdown.add("item", presetNames[i]);
                if (presetNames[i] === selectedName) selectedIndex = presetDropdown.items.length - 1;
            }
            presetDropdown.selection = selectedIndex;
            setIconButtonEnabled(presetUI.presetDeleteIcon, selectedIndex > BUILTIN_PRESETS.length);
            isFillingPresets = false;
        }

        /**
         * ［画面にフィット］がオンなら、配置先（アートボードか最背面のオブジェクト）が指定の割合で収まるよう表示を合わせます。
         * Fit the view to the placement area when Fit to Window is on.
         *
         * @returns {void}
         */
        function fitViewToTarget() {
            if (!fitViewControls.checkbox.value) return;
            /* 部品は geometricBounds だけを読むので、配置先の矩形をそのまま渡す / The part reads only geometricBounds */
            FitViewToItems.fit([{ geometricBounds: getTargetRect() }], { doc: doc, fillRatio: fitViewControls.getFillRatio() });
            app.redraw();
        }

        /**
         * プリセットの読み込み・保存・削除を結び付けます。/ Wire up loading, saving and deleting presets.
         *
         * @returns {void}
         */
        function bindPresetEvents() {
            presetUI.presetDropdown.onChange = function () {
                if (isFillingPresets) return;
                var previousMaskMode = maskMode;
                var presetName = getSelectedPresetName();
                setIconButtonEnabled(presetUI.presetDeleteIcon, presetName !== null);
                var builtinPreset = getSelectedBuiltinPreset();
                if (builtinPreset) {
                    if (!builtinPreset.data) {
                        applyMaskSettings(launchMaskSettings, previousMaskMode);
                        return;
                    }
                    /* 角丸・枠はマスクに付くので、マスクが［なし］なら［セル］にする / Corners and frames need a mask, so switch None to Cell */
                    if (!builtinPreset.data.hasOwnProperty("maskMode") && maskMode === MASK_MODE_NONE) maskMode = MASK_MODE_CELL;
                    applyMaskSettings(builtinPreset.data, previousMaskMode);
                    return;
                }
                if (!presetName) return;
                var presetData = loadPresetMap()[presetName];
                if (presetData) applyMaskSettings(presetData, previousMaskMode);
            };
            presetUI.presetSaveIcon.onClick = function () {
                var savedName = savePreset(getSelectedPresetName() || "", collectMaskSettings());
                if (savedName) fillPresetDropdown(savedName);
            };
            presetUI.presetDeleteIcon.onClick = function () {
                var presetName = getSelectedPresetName();
                if (presetName && deletePreset(presetName)) fillPresetDropdown(null);
            };
        }

        // -----------------------------------------
        // イベント / Event wiring
        // -----------------------------------------

        /**
         * ダイアログのイベントを、パネルのまとまりごとに結び付けます。/ Wire up the dialog events, panel group by panel group.
         *
         * @returns {void}
         */
        function bindDialogEvents() {
            bindLayoutEvents();
            bindResizeAndMaskEvents();
            bindCellHandlingEvents();
            bindFooterEvents();
        }

        /**
         * 左カラム（配置先・分割とマージン・配置と整列）のイベントを結び付けます。/ Wire up the left column.
         *
         * @returns {void}
         */
        function bindLayoutEvents() {
            /* 配置先を切り替えたら行数・列数を計算し直す。同じラジオの押し直しでは入力を残す
               Recalculate rows and columns when the area changes; a re-click keeps what was typed */
            placementUI.artboardRadio.onClick = placementUI.backmostRadio.onClick = function () {
                if (placementUI.backmostRadio.value === wasBackmostTarget) return;
                applyAutoDivision();
                fitViewToTarget();
            };

            divisionUI.autoDivisionButton.onClick = applyAutoDivision;
            divisionUI.rowCountInput.onChanging = divisionUI.columnCountInput.onChanging = function () {
                syncGutterEnabled();
                updatePreviewWhileTyping();
            };
            divisionUI.gutterInput.onChanging = updatePreviewWhileTyping;
            divisionUI.marginInput.onChanging = updatePreviewWhileTyping;
            divisionUI.gutterPriorityCheckbox.onClick = function () {
                isGutterPriorityEnabled = divisionUI.gutterPriorityCheckbox.value;
                updatePreview();
            };

            arrangementUI.spreadEvenlyCheckbox.onClick = function () {
                isSpreadEnabled = arrangementUI.spreadEvenlyCheckbox.value;
                updatePreview();
            };
            arrangementUI.brickCheckbox.onClick = function () {
                isBrickEnabled = arrangementUI.brickCheckbox.value;
                applyAutoDivision();
            };
            arrangementUI.keepOriginalsCheckbox.onClick = function () {
                isKeepOriginalsEnabled = arrangementUI.keepOriginalsCheckbox.value;
                updatePreview();
            };
            arrangementUI.btnShuffle.onClick = shuffleDistributionOrder;
        }

        /**
         * 中央カラム（リサイズ・マスク）のイベントを結び付けます。/ Wire up the center column.
         *
         * @returns {void}
         */
        function bindResizeAndMaskEvents() {
            resizeUI.noneRadio.onClick = resizeUI.widthRadio.onClick = resizeUI.heightRadio.onClick =
                resizeUI.areaRadio.onClick = resizeUI.scaleRadio.onClick = function () {
                    syncResizeRatioEnabled();
                    updatePreview();
                };
            resizeUI.ratioInput.onChanging = updatePreviewWhileTyping;
            resizeUI.autoScaleButton.onClick = applyAutoScale;

            maskUI.maskNoneRadio.onClick = function () { selectMaskMode(MASK_MODE_NONE); };
            maskUI.maskCellRadio.onClick = function () { selectMaskMode(MASK_MODE_CELL); };
            maskUI.maskItemRadio.onClick = function () { selectMaskMode(MASK_MODE_ITEM); };
            maskUI.cornerRadiusCheckbox.onClick = function () {
                isCornerRadiusEnabled = maskUI.cornerRadiusCheckbox.value;
                syncMaskStyleEnabled();
                updatePreview();
            };
            maskUI.frameCheckbox.onClick = function () {
                isFrameEnabled = maskUI.frameCheckbox.value;
                syncMaskStyleEnabled();
                updatePreview();
            };
            maskUI.cornerRadiusInput.onChanging = updatePreviewWhileTyping;
            maskUI.frameWidthInput.onChanging = updatePreviewWhileTyping;
            maskUI.autoCornerRadiusButton.onClick = applyAutoCornerRadius;
            maskUI.frameColorSwatch.onDraw = drawFrameColorSwatch;
            /* クリックで標準のカラーピッカーを開く。キャンセルは渡した色が返る / The standard color picker; Cancel returns the color passed in */
            maskUI.frameColorSwatch.addEventListener("click", function () {
                if (maskMode === MASK_MODE_NONE || !isFrameEnabled) return;
                frameColor = app.showColorPicker(frameColor);
                redrawFrameColorSwatch();
                updatePreview();
            });
        }

        /**
         * 右カラム（セルの扱い）のイベントを結び付けます。/ Wire up the Cell Handling panel.
         *
         * @returns {void}
         */
        function bindCellHandlingEvents() {
            cellUI.noCellRadio.onClick = cellUI.keepCellRadio.onClick = cellUI.toGuideRadio.onClick = cellUI.toArtboardRadio.onClick = function () {
                syncCellModeUI();
                updatePreview();
            };

            cellUI.blackCellRadio.onClick = cellUI.whiteCellRadio.onClick = cellUI.transparentCellRadio.onClick = function () {
                var selectedColorRadio = cellUI.whiteCellRadio.value ? cellUI.whiteCellRadio
                    : (cellUI.transparentCellRadio.value ? cellUI.transparentCellRadio : cellUI.blackCellRadio);

                /* 選択済みのラジオを押し直しても onClick は呼ばれる。カラーが変わっていなければ
                   入力済みの不透明度を既定値で上書きしない
                   A radio fires onClick even when re-clicked; keep a typed opacity when the color did not change */
                var hasColorChanged = (selectedColorRadio !== lastCellColorRadio);

                /* モードを往復してもカラー選択が失われないように覚えておく / Remember the color across mode switches */
                lastCellColorRadio = selectedColorRadio;
                syncOpacityEnabled(hasColorChanged);
                updatePreview();
            };
            cellUI.opacityInput.onChanging = updatePreviewWhileTyping;

            cellUI.transparencyGridButton.onClick = function () {
                app.executeMenuCommand('TransparencyGrid Menu Item');
                transparencyGridToggleCount++;
            };
        }

        /**
         * ボタン行（画面にフィット・キャンセル・OK）と、閉じたときの後始末を結び付けます。/ Wire up the button row and the close handler.
         *
         * @returns {void}
         */
        function bindFooterEvents() {
            footerUI.fitViewControls.checkbox.onClick = function () {
                footerUI.fitViewControls.updateEnabled();
                fitViewToTarget();
            };
            /* 割合の∧∨・↑↓キーは部品が受け持ち、増減したあと onChange を呼ぶ / The part's stepper calls onChange after stepping */
            footerUI.fitViewControls.percentInput.onChange = fitViewToTarget;

            footerUI.btnOK.onClick = commitDistribution;
            footerUI.btnCancel.onClick = function () {
                discardDialogChanges();
                isCleanedUp = true;
                distributeDialog.close(0);
            };

            /* OK／キャンセルを経由せずに閉じた場合の保険（ESC やウィンドウを閉じた操作でプレビューを残さない）
               Safety net when the dialog closes without OK or Cancel (ESC, closing the window) */
            distributeDialog.onClose = function () {
                if (!isCleanedUp) {
                    isCleanedUp = true;
                    try {
                        discardDialogChanges();
                    } catch (e) {
                        /* 後始末の失敗でクローズを妨げない / Never block the close */
                    }
                }
                /* falsy を返すとクローズが取り消される実装があるため、必ず true を返す / Always return true so the close goes through */
                return true;
            };
        }
    }

    // =========================================
    // 起動の振り分け / Launch routing
    // =========================================

    /**
     * 選択中のオブジェクトの数を返す（文字の編集中は1つとして数える）
     * Returns how many objects are selected (text editing counts as one)
     * @param {Document} doc - 対象のドキュメント / target document
     * @returns {number} 選択数 / number of selected objects
     */
    function countSelectedObjects(doc) {
        var selectedItems = doc.selection;
        if (selectedItems instanceof Array) return selectedItems.length;
        return selectedItems ? 1 : 0;
    }

    /* 2つ以上選択していれば整列、1つ以下なら画像を読み込んでから整列 / Distribute 2+ objects; otherwise import images first */
    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
    } else if (countSelectedObjects(app.activeDocument) < 2) {
        placeImagesThenDistribute();
    } else {
        showDistributeDialog();
    }

}());
