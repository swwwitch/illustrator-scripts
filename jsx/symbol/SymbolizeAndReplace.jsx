#target illustrator
#targetengine "SymbolizeAndReplaceEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトをシンボルに登録し、ドキュメント内の一致するオブジェクトをそのインスタンスにまとめて置き換えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeAndReplace.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n650a4b91329d

### Overview

Registers the selected object as a symbol and replaces every matching object in the document with an instance of it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeAndReplace.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SymbolizeAndReplace";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.10";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeAndReplace.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeAndReplace.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n650a4b91329d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* シンボル名の初期値。空欄のあいだは［OK］を押せない。テキスト選択時はその文字列が入る / Initial symbol name; OK stays disabled while it is empty. A selected TextFrame seeds it with its contents */
    var DEFAULT_SYMBOL_NAME = '';

    /* 基準点の初期値（"topLeft"〜"bottomRight" か 0〜8）/ Initial registration point ("topLeft" … "bottomRight" or 0-8) */
    var DEFAULT_REFERENCE_POINT = 'center';

    // =========================================
    // レイアウト / Layout
    // =========================================

    var SYMBOL_NAME_FIELD_WIDTH = 210; /* シンボル名欄の幅 / width of the symbol-name field */

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

    // =========================================
    // UI 部品 / UI parts
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

    /* 日英ラベル定義（UI の部品ごと。getLabel('dialog.title') のようにドット区切りで引く）/ Japanese-English labels grouped by UI part; look up with dotted keys such as getLabel('dialog.title') */
    var LABELS = {
        dialog: {
            title: { ja: 'シンボル化して置換', en: 'Symbolize and Replace' }
        },
        panel: {
            symbolName: { ja: 'シンボル名', en: 'Symbol Name' },
            referencePoint: { ja: '基準点', en: 'Registration Point' },
            text: { ja: 'テキスト', en: 'Text' }
        },
        checkbox: {
            allowSizeMismatch: { ja: 'フォントサイズ違いも対象にする', en: 'Include different font sizes' }
        },
        button: {
            cancel: { ja: 'キャンセル', en: 'Cancel' },
            ok: { ja: 'OK', en: 'OK' }
        },
        tooltip: {
            symbolName: {
                ja: '新しく作成するシンボルの名前です。空欄、および既存シンボルと同じ名前は使用できません。',
                en: 'Name of the new symbol. Empty names and names already used by existing symbols are not allowed.'
            },
            referencePoint: {
                ja: 'シンボル登録時の基準点です。置換時もこの点を使って元オブジェクトの位置に揃えます。',
                en: 'Sets the registration point for the new symbol. The same point is used to align each replacement instance to the original object.'
            },
            allowSizeMismatch: {
                ja: 'ON のときは、フォント・スタイル・文字列が一致していれば、フォントサイズが異なるテキストフレームも置換対象に含めます。',
                en: 'When enabled, text frames with matching font, style, and contents are included even if their font size differs.'
            }
        },
        alert: {
            skipped: {
                ja: '{count} 件はロック／非表示、または親レイヤーの状態により置換できませんでした。',
                en: '{count} item(s) could not be replaced because they or their parent layers were locked or hidden.'
            },
            duplicateSymbol: {
                ja: 'シンボル「{name}」は既に存在します。別の名前を指定してください。',
                en: 'A symbol named "{name}" already exists. Please choose a different name.'
            },
            noTargets: {
                ja: '置換対象が見つからなかったため、作成したシンボルを削除しました。',
                en: 'No replacement targets were found, so the created symbol was removed.'
            },
            multiSelectionNotGroups: {
                ja: '複数選択時は、すべてのアイテムがグループである必要があります。グループ以外を含む選択では実行できません。',
                en: 'When multiple items are selected, every item must be a group. The script cannot run if the selection includes non-group items.'
            }
        }
    };

    // =========================================
    // 一時アクション設定 / Temporary action settings
    // =========================================

    /* 一括選択用ダイナミックアクションのセット名・アクション名（任意の内部ラベル。doScript と /name で同じ定数を使う）/ Set and action names for the bulk-select dynamic action; arbitrary internal labels shared by doScript and the /name lines */
    var ACTION_SET_NAME = 'SymbolizeAndReplaceNote';
    var ACTION_NAME = 'AttachTempNote';

    /* 一括選択した対象に付ける note の値。アクション定義へは16進にして埋め込む / Note value tagged onto bulk-selected items; hex-encoded into the action definition */
    var TEMP_NOTE_VALUE = 'temp_memo';

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /**
     * 選択中のアイテムの note 属性に値を書き込むアクション定義を作る。
     * internalName（adobe_attributePalette）と parameter-1（/key 1852798053, /type ustring）は実際に記録した .aia から採取した値なので、勘で書き換えない
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} noteValue - 書き込む note の値
     * @returns {string} アクション定義のテキスト
     */
    function buildAttachNoteActionSource(setName, actionName, noteValue) {
        return ['/version 3']
            .concat(buildActionNameLines('', setName))
            .concat([
                '/isOpen 1',
                '/actionCount 1',
                '/action-1 {'
            ])
            .concat(buildActionNameLines('\t', actionName))
            .concat([
                '\t/keyIndex 0',
                '\t/colorIndex 0',
                '\t/isOpen 1',
                '\t/eventCount 1',
                '\t/event-1 {',
                '\t\t/useRulersIn1stQuadrant 0',
                '\t\t/internalName (adobe_attributePalette)',
                '\t\t/localizedName [ 0  ]',
                '\t\t/isOpen 1',
                '\t\t/isOn 1',
                '\t\t/hasDialog 0',
                '\t\t/parameterCount 1',
                '\t\t/parameter-1 {',
                '\t\t\t/key 1852798053',
                '\t\t\t/showInPalette 4294967295',
                '\t\t\t/type (ustring)'
            ])
            .concat(buildActionNameLines('\t\t\t', noteValue, 'value'))
            .concat([
                '\t\t}',
                '\t}',
                '}'
            ])
            .join('\n') + '\n';
    }

    // =========================================
    // 置換対象の収集 / Collecting replacement targets
    // =========================================

    /**
     * 指定した note の値を持つページアイテムをすべて集める
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {string} noteValue - 探す note の値
     * @returns {PageItem[]} 見つかったアイテム
     */
    function collectItemsByNote(targetDocument, noteValue) {
        var notedItems = [];
        var allPageItems = targetDocument.pageItems;
        for (var i = 0, len = allPageItems.length; i < len; i++) {
            if (allPageItems[i].note === noteValue) notedItems.push(allPageItems[i]);
        }
        return notedItems;
    }

    /**
     * アイテムの note を空にする
     * @param {PageItem[]} notedItems - 対象のアイテム
     * @returns {void}
     */
    function clearNoteOnItems(notedItems) {
        for (var i = 0; i < notedItems.length; i++) {
            /* ロック中のアイテムは書き込めないことがある。残っても次回の実行前に消す / Locked items may refuse the write; leftovers are cleared before the next run */
            try { notedItems[i].note = ''; } catch (e) { }
        }
    }

    /**
     * スマート編集（「同じ」の一括選択）で元オブジェクトと同じアイテムを選び、note を付けて集める。
     * 集めたら note は消す
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {PageItem[]} 元オブジェクトを含む一致アイテム
     */
    function collectSimilarItemsBySmartEdit(targetDocument) {
        /* 以前の実行が途中で止まって note が残っていれば先に消す / Clear notes left by an earlier run that stopped midway */
        clearNoteOnItems(collectItemsByNote(targetDocument, TEMP_NOTE_VALUE));

        app.executeMenuCommand('SmartEdit Menu Item');
        /* 失敗しても note が付かないだけで、後段の「対象なし」の警告で終わる / On failure no note is attached and the run ends with the no-targets alert */
        runTemporaryAction(buildAttachNoteActionSource(ACTION_SET_NAME, ACTION_NAME, TEMP_NOTE_VALUE), ACTION_SET_NAME, ACTION_NAME);
        /* doScript の実行でスマート編集は自動で OFF に戻る。もう一度 'SmartEdit Menu Item' を呼ぶと逆に ON になるので呼ばない / Running doScript turns SmartEdit off automatically; calling 'SmartEdit Menu Item' again would turn it back on */

        var similarItems = collectItemsByNote(targetDocument, TEMP_NOTE_VALUE);
        clearNoteOnItems(similarItems);
        return similarItems;
    }

    /**
     * 同じフォント・スタイル（allowSizeMismatch が false ならサイズも）で、文字列も同じテキストフレームを集める。
     * メニューコマンドで選択し直すので、document.selection が書き換わる
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {string} sourceTextContent - 元のテキストの文字列
     * @param {boolean} allowSizeMismatch - true ならフォントサイズ違いも含める
     * @returns {TextFrame[]} 一致したテキストフレーム
     */
    function findMatchingTextFrames(targetDocument, sourceTextContent, allowSizeMismatch) {
        app.executeMenuCommand(allowSizeMismatch
            ? 'Find Text Font Family Style menu item'
            : 'Find Text Font Family Style Size menu item');
        var matchedFrames = [];
        var currentSelection = targetDocument.selection;
        for (var i = 0; i < currentSelection.length; i++) {
            if (isTextFrame(currentSelection[i]) && currentSelection[i].contents === sourceTextContent) {
                matchedFrames.push(currentSelection[i]);
            }
        }
        return matchedFrames;
    }

    /**
     * 選択の種類に応じて置換対象を集める
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {Object} selectionInfo - analyzeSelection() の結果
     * @param {boolean} allowSizeMismatch - テキストのときフォントサイズ違いも含めるか
     * @returns {PageItem[]} 置換対象（元オブジェクトを含む）
     */
    function collectReplaceTargets(targetDocument, selectionInfo, allowSizeMismatch) {
        /* 複数グループ選択は選択そのもの / Multi-group selection uses the selection itself */
        if (selectionInfo.isMultiGroup) return selectionInfo.items;
        if (selectionInfo.isText) return findMatchingTextFrames(targetDocument, selectionInfo.textContent, allowSizeMismatch);
        return collectSimilarItemsBySmartEdit(targetDocument);
    }

    // =========================================
    // ヘルパー関数 / Helper functions
    // =========================================

    /**
     * テキストフレームかどうか
     * @param {PageItem} item - 調べるアイテム
     * @returns {boolean} TextFrame なら true
     */
    function isTextFrame(item) {
        return !!item && item.typename === 'TextFrame';
    }

    /**
     * 配列の要素がすべてグループかどうか（空の配列は false）
     * @param {PageItem[]} items - 調べるアイテム
     * @returns {boolean} すべて GroupItem なら true
     */
    function areAllGroups(items) {
        if (!items || items.length === 0) return false;
        for (var i = 0; i < items.length; i++) {
            if (!items[i] || items[i].typename !== 'GroupItem') return false;
        }
        return true;
    }

    /**
     * テキストをシンボル名に使える形にする（改行を半角スペースにまとめる）
     * @param {string} sourceText - 元の文字列
     * @returns {string} 整えた文字列
     */
    function sanitizeTextForSymbolName(sourceText) {
        return (sourceText || '').replace(/[\r\n]+/g, ' ');
    }

    /**
     * 前後の空白を取る（ExtendScript には String.prototype.trim が無い）
     * @param {string} sourceText - 元の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimWhitespace(sourceText) {
        return (sourceText || '').replace(/^\s+|\s+$/g, '');
    }

    /**
     * 指定名のシンボルが既にあるかどうか
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {string} symbolName - シンボル名
     * @returns {boolean} あれば true
     */
    function symbolNameExists(targetDocument, symbolName) {
        /* getByName は見つからないと例外を投げる / getByName throws when the name is missing */
        try {
            targetDocument.symbols.getByName(symbolName);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 選択を調べ、置換のモード（複数グループ／テキスト／それ以外）を決める。複数選択にグループ以外が混ざっていれば警告して null
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {{items: PageItem[], isMultiGroup: boolean, isText: boolean, textContent: string|null}|null} 選択の情報。実行できないときは null
     */
    function analyzeSelection(targetDocument) {
        var rawSelection = targetDocument.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
        if (!rawSelection || rawSelection.typename === 'TextRange' || !rawSelection.length) return null;

        var isMultiGroup = rawSelection.length > 1;
        if (isMultiGroup && !areAllGroups(rawSelection)) {
            alert(getLabel('alert.multiSelectionNotGroups'));
            return null;
        }

        /* createSymbolFromSelection が選択を1個に絞るので、参照を配列に控える / Keep references since createSymbolFromSelection narrows the selection to one item */
        var selectedItems = [];
        for (var i = 0; i < rawSelection.length; i++) {
            selectedItems.push(rawSelection[i]);
        }

        var isText = !isMultiGroup && isTextFrame(selectedItems[0]);
        return {
            items: selectedItems,
            isMultiGroup: isMultiGroup,
            isText: isText,
            textContent: isText ? selectedItems[0].contents : null
        };
    }

    // =========================================
    // シンボル化と置換 / Symbolize and replace
    // =========================================

    /**
     * 選択の先頭を複製してシンボルにし、複製は消す。そのあと元オブジェクトを選択し直す
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {string} symbolName - シンボル名
     * @param {number} referencePointIndex - 基準点（0〜8）
     * @returns {Symbol} 作成したシンボル
     */
    function createSymbolFromSelection(targetDocument, symbolName, referencePointIndex) {
        var sourceItem = targetDocument.selection[0];
        var duplicatedItem = sourceItem.duplicate();
        var createdSymbol = targetDocument.symbols.add(duplicatedItem, getAnchorSymbolRegistrationPoint(referencePointIndex));
        createdSymbol.name = symbolName;
        duplicatedItem.remove();

        /* 後続の検索・スマート編集が元選択を使えるよう、元オブジェクトを選択し直す / Reselect the original so the following search or SmartEdit can use it */
        targetDocument.selection = null;
        sourceItem.selected = true;
        return createdSymbol;
    }

    /**
     * 置換できるアイテムかどうか（本体と親レイヤーをたどって、ロック・非表示が無いか）
     * @param {PageItem} targetItem - 調べるアイテム
     * @returns {boolean} 置換できれば true
     */
    function isReplaceableItem(targetItem) {
        if (!targetItem || targetItem.locked || targetItem.hidden) return false;
        var currentLayer = targetItem.layer;
        while (currentLayer && currentLayer.typename === 'Layer') {
            if (currentLayer.locked || !currentLayer.visible) return false;
            currentLayer = currentLayer.parent;
        }
        return true;
    }

    /**
     * シンボルインスタンスを作り、元オブジェクトと基準点どうしが揃うよう移動する
     * @param {Layer} destinationLayer - 配置先のレイヤー
     * @param {Symbol} destinationSymbol - 配置するシンボル
     * @param {PageItem} targetItem - 位置を合わせる元オブジェクト
     * @param {number} referencePointIndex - 基準点（0〜8）
     * @returns {SymbolItem} 作成したインスタンス
     */
    function createAlignedSymbolItem(destinationLayer, destinationSymbol, targetItem, referencePointIndex) {
        var destinationPosition = getAnchorPointOnBounds(targetItem.visibleBounds, referencePointIndex);
        var newSymbolItem = destinationLayer.symbolItems.add(destinationSymbol);
        var sourcePosition = getAnchorPointOnBounds(newSymbolItem.visibleBounds, referencePointIndex);
        newSymbolItem.translate(destinationPosition[0] - sourcePosition[0], destinationPosition[1] - sourcePosition[1]);
        return newSymbolItem;
    }

    /**
     * 対象をひとつずつシンボルインスタンスに置き換え、置き換えたものを選択する。置換できなかった件数は警告で知らせる
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {PageItem[]} targetItems - 置換対象
     * @param {Symbol} destinationSymbol - 置き換えるシンボル
     * @param {number} referencePointIndex - 基準点（0〜8）
     * @returns {void}
     */
    function replaceItemsWithSymbol(targetDocument, targetItems, destinationSymbol, referencePointIndex) {
        var createdSymbolItems = [];
        var skippedCount = 0;

        for (var i = 0; i < targetItems.length; i++) {
            var currentItem = targetItems[i];
            var newSymbolItem = null;
            try {
                if (!isReplaceableItem(currentItem)) {
                    skippedCount++;
                    continue;
                }
                newSymbolItem = createAlignedSymbolItem(currentItem.layer, destinationSymbol, currentItem, referencePointIndex);
                currentItem.remove();
                createdSymbolItems.push(newSymbolItem);
            } catch (e) {
                /* 途中で失敗したら作りかけのインスタンスを消して数える / Remove the half-made instance and count the failure */
                if (newSymbolItem) {
                    try { newSymbolItem.remove(); } catch (removeError) { }
                }
                skippedCount++;
            }
        }

        /* 1個ずつ .selected = true にすると多数のとき固まるので、配列でまとめて代入する / Assign as an array; per-item .selected = true freezes on large counts */
        targetDocument.selection = createdSymbolItems;

        if (skippedCount > 0) {
            alert(getLabel('alert.skipped', { count: skippedCount }));
        }
    }

    // =========================================
    // ダイアログ / Dialog
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

    /**
     * シンボル名が既存のシンボルと重ならないか確かめる。重なれば警告して入力欄へ戻す
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {string} candidateName - 確かめるシンボル名
     * @param {EditText} symbolNameInput - シンボル名の入力欄
     * @returns {boolean} 使える名前なら true
     */
    function validateSymbolNameUnique(targetDocument, candidateName, symbolNameInput) {
        if (!symbolNameExists(targetDocument, candidateName)) return true;
        alert(getLabel('alert.duplicateSymbol', { name: candidateName }));
        symbolNameInput.active = true;
        return false;
    }

    /**
     * シンボル化の設定ダイアログを表示する
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {string} defaultName - シンボル名の初期値
     * @param {number|string} defaultReferencePoint - 基準点の初期値
     * @param {boolean} isTextSelection - テキストフレームを選択しているか（テキストのパネルを出す）
     * @returns {{symbolName: string, referencePoint: number, allowSizeMismatch: boolean}|null} 設定。キャンセルなら null
     */
    function showSymbolizeDialog(targetDocument, defaultName, defaultReferencePoint, isTextSelection) {
        var symbolizeDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(symbolizeDialog);

        /* シンボル名 / Symbol name */
        var symbolNamePanel = symbolizeDialog.add('panel', undefined, getLabel('panel.symbolName'));
        setupPanel(symbolNamePanel);
        var symbolNameInput = symbolNamePanel.add('edittext', undefined, defaultName);
        symbolNameInput.preferredSize.width = SYMBOL_NAME_FIELD_WIDTH;
        symbolNameInput.helpTip = getLabel('tooltip.symbolName');
        symbolNameInput.active = true;

        /* 基準点 / Registration point */
        var referencePointPanel = symbolizeDialog.add('panel', undefined, getLabel('panel.referencePoint'));
        setupPanel(referencePointPanel);
        var referencePointWidget = addAnchorWidget(referencePointPanel, defaultReferencePoint);
        referencePointWidget.alignment = 'center';
        referencePointWidget.helpTip = getLabel('tooltip.referencePoint');

        /* テキスト（テキスト選択時のみ）/ Text options (only when a TextFrame is selected) */
        var allowSizeMismatchCheckbox = null;
        if (isTextSelection) {
            var textOptionsPanel = symbolizeDialog.add('panel', undefined, getLabel('panel.text'));
            setupPanel(textOptionsPanel, 6);
            allowSizeMismatchCheckbox = textOptionsPanel.add('checkbox', undefined, getLabel('checkbox.allowSizeMismatch'));
            allowSizeMismatchCheckbox.helpTip = getLabel('tooltip.allowSizeMismatch');
        }

        /* ［OK］は重複チェックでダイアログを残せるよう name:'ok' を付けず、defaultElement で Return に結び付ける / OK omits name:'ok' so the duplicate check can keep the dialog open; defaultElement binds Return to it */
        var buttonRow = addButtonRow(symbolizeDialog);
        buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'));
        btnOK.onClick = function () {
            if (!validateSymbolNameUnique(targetDocument, trimWhitespace(symbolNameInput.text), symbolNameInput)) return;
            symbolizeDialog.close(1);
        };
        symbolizeDialog.defaultElement = btnOK;

        /* 空欄（空白だけを含む）のあいだは［OK］を押せない / OK stays disabled while the name is empty or blank */
        function updateOkButtonState() {
            btnOK.enabled = trimWhitespace(symbolNameInput.text).length > 0;
        }
        symbolNameInput.onChanging = updateOkButtonState;
        updateOkButtonState();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(symbolizeDialog, SCRIPT_NAME);
        if (symbolizeDialog.show() !== 1) return null;
        return {
            symbolName: trimWhitespace(symbolNameInput.text),
            referencePoint: getAnchorWidgetIndex(referencePointWidget),
            allowSizeMismatch: allowSizeMismatchCheckbox ? allowSizeMismatchCheckbox.value : false
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択をシンボル化し、一致するアイテムをそのインスタンスに置き換える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var activeDoc = app.activeDocument;
        var selectionInfo = analyzeSelection(activeDoc);
        if (!selectionInfo) return;

        /* テキストならその文字列をシンボル名の初期値にする / A TextFrame seeds the symbol name with its contents */
        var initialName = selectionInfo.textContent ? sanitizeTextForSymbolName(selectionInfo.textContent) : DEFAULT_SYMBOL_NAME;
        var dialogResult = showSymbolizeDialog(activeDoc, initialName, DEFAULT_REFERENCE_POINT, selectionInfo.isText);
        if (!dialogResult) return;

        var createdSymbol = createSymbolFromSelection(activeDoc, dialogResult.symbolName, dialogResult.referencePoint);
        var replaceTargets = collectReplaceTargets(activeDoc, selectionInfo, dialogResult.allowSizeMismatch);

        /* 元オブジェクトのほかに対象が無ければ、作ったシンボルを消して終える / Remove the created symbol and stop when nothing but the original was found */
        if (replaceTargets.length <= 1) {
            createdSymbol.remove();
            alert(getLabel('alert.noTargets'));
            return;
        }

        replaceItemsWithSymbol(activeDoc, replaceTargets, createdSymbol, dialogResult.referencePoint);
    }

    main();

})();
