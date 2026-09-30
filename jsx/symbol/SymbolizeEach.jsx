#target illustrator
#targetengine "SymbolizeEachEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを1つずつ、またはまとめてシンボルとして登録し、元のオブジェクトをシンボルインスタンスに置き換えます。
名前はテキスト内容・レイヤー名・メモ・連番から自動で付けるか、Illustrator 標準の［新規シンボル］ダイアログで確認しながら登録できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeEach.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nce9ec30232a0

### Overview

Registers the selected objects as symbols, one by one or as a single symbol, and replaces the originals with symbol instances.
Names are assigned automatically from the text contents, layer name, note, or a sequence number, or confirmed in Illustrator's native New Symbol dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeEach.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SymbolizeEach";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.10";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeEach.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeEach.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nce9ec30232a0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 名前を連番で付けるときの既定の接頭辞 / Default prefix for sequence-numbered names */
    var DEFAULT_PREFIX = 'Symbol_';

    /* 連番の既定の桁数（1〜3）/ Default number of sequence digits (1–3) */
    var DEFAULT_SEQUENCE_DIGITS = 3;

    /* ［テキスト内容をシンボル名に使う］の初期状態 / Initial state of "Use text contents as symbol name" */
    var DEFAULT_USE_TEXT_AS_NAME = true;

    /* 基準点の初期選択（0〜8 の行優先、4 = 中央）/ Initial registration point (0–8, row-major; 4 = center) */
    var DEFAULT_REFERENCE_POINT_INDEX = 4;

    /* 選択範囲の扱いの初期選択（'asGroup' = まとめて1つ / 'eachItem' = オブジェクトごと）/ Initial selection handling */
    var DEFAULT_GROUP_MODE = 'eachItem';

    /* 登録方法の初期選択（'individual' = 標準ダイアログで確認 / 'batch' = 自動登録）/ Initial registration method */
    var DEFAULT_SYMBOLIZE_MODE = 'batch';

    /* リンク画像の初期選択（'ignore' = 無視 / 'embed' = 埋め込んで登録）/ Initial linked-image policy */
    var DEFAULT_LINKED_IMAGE_POLICY = 'embed';

    /* シンボル名の最大文字数 / Maximum length of a symbol name */
    var SYMBOL_NAME_MAX_LENGTH = 80;

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

    var FIELD_LABEL_WIDTH  = 64;                 /* 項目名の幅 / field label width */
    var PREFIX_INPUT_WIDTH = 180;                /* 接頭辞の入力欄の幅 / prefix field width */

    /**
     * タイトル付きのパネルを追加し、共通レイアウトを設定する（子は左寄せ、パネルは天地にも伸ばす）
     * @param {object} parent - 追加先のウィンドウまたはグループ
     * @param {object} titleSet - パネル名のラベル（ja/en）
     * @param {number} [spacing] - パネル内の要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleSet, spacing) {
        var newPanel = parent.add('panel', undefined, getLabel(titleSet));
        setupPanel(newPanel, spacing);
        newPanel.alignChildren = ['left', 'top'];
        newPanel.alignment = ['fill', 'fill'];
        return newPanel;
    }

    /**
     * 行の先頭に右揃えの項目名を追加する
     * @param {Group} rowGroup - 追加先の行グループ
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @returns {StaticText} 追加した項目名
     */
    function addFieldLabel(rowGroup, labelSet) {
        var fieldLabel = rowGroup.add('statictext', undefined, labelText(labelSet));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = 'right';
        return fieldLabel;
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 選択範囲の扱い / Selection handling */
    var GROUP_MODE = { AS_GROUP: 'asGroup', EACH_ITEM: 'eachItem' };

    /* 登録方法 / Registration method */
    var SYMBOLIZE_MODE = { INDIVIDUAL: 'individual', BATCH: 'batch' };

    /* リンク画像の扱い / Linked-image policy */
    var LINKED_IMAGE_POLICY = { IGNORE: 'ignore', EMBED: 'embed' };

    /* 連番の桁数の選択肢（"0" / "00" / "000"）/ Sequence-digit choices */
    var SEQUENCE_DIGIT_OPTIONS = [1, 2, 3];

    /* 名前の候補にしない既定のレイヤー名 / Default layer names that are not used as symbol names */
    var DEFAULT_LAYER_NAME_PATTERN = /^(Layer|レイヤー)\s*\d+$/;

    /* 「新規シンボル…」のメニューコマンド / Menu command for "New Symbol..." */
    var NEW_SYMBOL_MENU_COMMAND = 'Adobe New Symbol Shortcut';

    /* 標準ダイアログ表示前に対象へ寄せるビュー（対象がビューの半分を占める倍率、上下限つき）/ View focus before the native dialog */
    var FOCUS_VIEW_RATIO = 0.5;
    var MIN_VIEW_ZOOM = 0.05;
    var MAX_VIEW_ZOOM = 64;

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: 'シンボル化', en: 'Symbolize' }
        },
        panel: {
            groupMode: { ja: '選択範囲の扱い', en: 'Selection handling' },
            symbolizeMode: { ja: '登録方法', en: 'Registration method' },
            symbolName: { ja: 'シンボル名', en: 'Symbol name' },
            referencePoint: { ja: '基準点', en: 'Registration point' },
            linkedImage: { ja: 'リンク画像', en: 'Linked images' }
        },
        radio: {
            asGroup: { ja: 'まとめて1つのシンボルにする', en: 'Create one symbol from selection' },
            eachItem: { ja: 'オブジェクトごとにシンボル化', en: 'Create symbols per object' },
            individual: { ja: '標準ダイアログで確認', en: 'Confirm with native dialog' },
            batch: { ja: '自動登録', en: 'Register automatically' },
            ignoreLinked: { ja: '無視', en: 'Ignore' },
            embedLinked: { ja: '埋め込んで登録', en: 'Embed and register' }
        },
        checkbox: {
            useTextAsName: { ja: 'テキスト内容をシンボル名に使う', en: 'Use text contents as symbol name' }
        },
        fieldLabel: {
            prefix: { ja: '接頭辞', en: 'Prefix' },
            sequence: { ja: '連番', en: 'Sequence' }
        },
        tooltip: {
            asGroup: {
                ja: '現在の選択全体を1つのシンボルとして登録します。この場合、登録方法は標準ダイアログでの確認になります。',
                en: 'Registers the current selection as one symbol. In this mode, the native dialog is used for confirmation.'
            },
            eachItem: { ja: '選択中の各オブジェクトを個別のシンボルとして登録します。', en: 'Registers each selected object as its own symbol.' },
            individual: {
                ja: '対象ごとに Illustrator 標準の「新規シンボル」ダイアログを開きます。名前や基準点を毎回確認したい場合に使います。',
                en: 'Opens Illustrator\'s native New Symbol dialog for each target. Use this when you want to confirm the name and registration point each time.'
            },
            batch: {
                ja: '接頭辞・連番・テキスト流用・基準点の設定に従って、確認なしで自動登録します。',
                en: 'Registers symbols automatically without confirmation, using the prefix, sequence, text-reuse, and registration-point settings.'
            },
            prefix: {
                ja: 'テキスト内容・レイヤー名・メモから名前を取得できない場合に使う接頭辞です。',
                en: 'Prefix used when the symbol name cannot be taken from text contents, the layer name, or the item note.'
            },
            sequence: {
                ja: '接頭辞に続ける連番の桁数です。例：0、00、000。',
                en: 'Number of digits for the sequence appended to the prefix, such as 0, 00, or 000.'
            },
            useTextAsName: {
                ja: 'TextFrame、またはグループ内で最初に見つかった TextFrame の内容をシンボル名に使います。空の場合は次の候補に進みます。',
                en: 'Uses the TextFrame contents, or the first TextFrame found inside a group, as the symbol name. If empty, the next naming source is used.'
            },
            referencePoint: {
                ja: '自動登録時に使うシンボルの基準点です。標準ダイアログで確認する場合は Illustrator 側で指定します。',
                en: 'Registration point used for automatic registration. When using the native dialog, set it in Illustrator.'
            },
            linkedImage: {
                ja: 'リンク画像（PlacedItem）の扱いです。［無視］は選択に残して登録対象から外します。［埋め込んで登録］はシンボル化前に埋め込みます。',
                en: 'Policy for linked images (PlacedItem). Ignore keeps them selected but excludes them. Embed and register embeds them before symbolization.'
            }
        },
        button: {
            cancel: { ja: 'キャンセル', en: 'Cancel' },
            ok: { ja: 'OK', en: 'OK' }
        },
        alert: {
            noDocument: { ja: 'ドキュメントが開かれていません。', en: 'No document is open.' },
            noSelection: { ja: 'シンボル化したいオブジェクトを選択してください。', en: 'Select the objects you want to symbolize.' },
            completed: { ja: '完了しました。', en: 'Completed.' },
            createdCount: { ja: '新規作成したシンボル数', en: 'Created symbols' },
            existingCount: { ja: 'スルーした既存シンボル数', en: 'Skipped existing symbols' },
            ignoredLinkedCount: { ja: '無視したリンク画像数', en: 'Ignored linked images' },
            failedCount: { ja: '失敗', en: 'Failed' },
            failureDetails: { ja: '失敗の詳細', en: 'Failure details' }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * @typedef {object} SymbolizeSettings
     * @property {string} groupMode - 選択範囲の扱い（GROUP_MODE）
     * @property {string} symbolizeMode - 登録方法（SYMBOLIZE_MODE）
     * @property {string} defaultPrefix - 連番で名前を付けるときの接頭辞
     * @property {number} sequenceDigits - 連番の桁数
     * @property {boolean} useTextAsName - テキスト内容をシンボル名に使うか
     * @property {SymbolRegistrationPoint} referencePoint - 自動登録時の基準点
     * @property {string} linkedImagePolicy - リンク画像の扱い（LINKED_IMAGE_POLICY）
     */

    /**
     * @typedef {object} ResultStats
     * @property {PageItem[]} finalSelection - 処理後に選択するアイテム
     * @property {number} createdCount - 新規作成したシンボル数
     * @property {number} existingSymbolCount - スルーした既存シンボルインスタンス数
     * @property {number} ignoredLinkedImageCount - 無視したリンク画像数
     * @property {string[]} failureMessages - 失敗の詳細（件数は length）
     */

    /**
     * 選択を確認し、ダイアログの設定に従ってシンボル化して結果を通知する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var originalSelection = snapshotSelection(doc);

        var symbolizeSettings = showSettingsDialog();
        if (!symbolizeSettings) return;

        doc.selection = null;

        var resultStats = (symbolizeSettings.groupMode === GROUP_MODE.AS_GROUP)
            ? processSelectionAsGroup(doc, originalSelection, symbolizeSettings)
            : processEachItem(doc, originalSelection, symbolizeSettings);

        applySelection(doc, resultStats.finalSelection);
        showResultSummary(resultStats);
    }

    // =========================================
    // シンボル化の流れ / Symbolize flow
    // =========================================

    /**
     * 現在の選択を配列としてコピーする
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択アイテムの配列
     */
    function snapshotSelection(doc) {
        var selectionItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectionItems.push(doc.selection[i]);
        }
        return selectionItems;
    }

    /**
     * 空の集計オブジェクトを作る
     * @returns {ResultStats} 集計オブジェクト
     */
    function createResultStats() {
        return {
            finalSelection: [],
            createdCount: 0,
            existingSymbolCount: 0,
            ignoredLinkedImageCount: 0,
            failureMessages: []
        };
    }

    /**
     * 既存のシンボルインスタンスと［無視］指定のリンク画像を、処理せず選択に残して数える
     * @param {PageItem} item - 判定するアイテム
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @param {ResultStats} resultStats - 集計先
     * @returns {boolean} 処理対象外なら true
     */
    function skipExcludedItem(item, symbolizeSettings, resultStats) {
        if (item.typename === 'SymbolItem') {
            resultStats.existingSymbolCount++;
        } else if (item.typename === 'PlacedItem' && symbolizeSettings.linkedImagePolicy === LINKED_IMAGE_POLICY.IGNORE) {
            resultStats.ignoredLinkedImageCount++;
        } else {
            return false;
        }
        resultStats.finalSelection.push(item);
        return true;
    }

    /**
     * 失敗したアイテムの詳細を集計に加える
     * @param {ResultStats} resultStats - 集計先
     * @param {PageItem} item - 失敗したアイテム
     * @param {Error} err - 発生したエラー
     * @returns {void}
     */
    function recordFailure(resultStats, item, err) {
        resultStats.failureMessages.push(formatFailureMessage(item, err));
    }

    /**
     * オブジェクトごとにシンボル化する（自動登録、または対象ごとに標準ダイアログで確認）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 処理するアイテム
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {ResultStats} 集計結果
     */
    function processEachItem(doc, items, symbolizeSettings) {
        var resultStats = createResultStats();
        var useNativeDialog = (symbolizeSettings.symbolizeMode === SYMBOLIZE_MODE.INDIVIDUAL);

        /* 標準ダイアログでは対象へビューを寄せるので、終了時に元のビューへ戻す / Restore the view after focusing on each target */
        var savedView = useNativeDialog ? captureViewState(doc) : null;

        try {
            for (var j = 0; j < items.length; j++) {
                var originalItem = items[j];
                if (skipExcludedItem(originalItem, symbolizeSettings, resultStats)) continue;

                try {
                    if (useNativeDialog) {
                        runNewSymbolDialog(doc, [ensureEmbedded(doc, originalItem)], resultStats);
                    } else {
                        resultStats.finalSelection.push(symbolizeOneItem(doc, originalItem, j, symbolizeSettings));
                        resultStats.createdCount++;
                    }
                } catch (err) {
                    recordFailure(resultStats, originalItem, err);
                }
            }
        } finally {
            restoreViewState(doc, savedView);
        }

        return resultStats;
    }

    /**
     * 選択全体をまとめて1つのシンボルとして、標準ダイアログで登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 処理するアイテム
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {ResultStats} 集計結果
     */
    function processSelectionAsGroup(doc, items, symbolizeSettings) {
        var resultStats = createResultStats();

        /* リンク画像の扱いを反映しながら、まとめて登録するアイテムを集める / Collect the targets, embedding linked images as needed */
        var targetItems = [];
        for (var j = 0; j < items.length; j++) {
            var originalItem = items[j];
            if (skipExcludedItem(originalItem, symbolizeSettings, resultStats)) continue;

            try {
                targetItems.push(ensureEmbedded(doc, originalItem));
            } catch (err) {
                recordFailure(resultStats, originalItem, err);
            }
        }
        if (targetItems.length === 0) return resultStats;

        var savedView = captureViewState(doc);
        try {
            runNewSymbolDialog(doc, targetItems, resultStats);
        } catch (err) {
            recordFailure(resultStats, targetItems[0], err);
        } finally {
            restoreViewState(doc, savedView);
        }

        return resultStats;
    }

    /**
     * 対象だけを選択してビューを寄せ、Illustrator 標準の「新規シンボル…」ダイアログを開く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - シンボルにするアイテム
     * @param {ResultStats} resultStats - 集計先
     * @returns {void}
     */
    function runNewSymbolDialog(doc, targetItems, resultStats) {
        applySelection(doc, targetItems);
        if (!doc.selection || doc.selection.length === 0) {
            throw new Error('Target items could not be selected.');
        }

        /* モーダル表示前に redraw して選択のハイライトを描く / Redraw so the highlight shows before the modal dialog */
        focusViewOnItems(doc, targetItems);
        app.redraw();

        var symbolCountBefore = doc.symbols.length;
        app.executeMenuCommand(NEW_SYMBOL_MENU_COMMAND);
        if (doc.symbols.length > symbolCountBefore) {
            resultStats.createdCount++;
        }

        /* 実行後の選択をそのまま結果に取り込む（キャンセル時は元のアイテムが残る）/ Keep whatever is selected afterward (cancel leaves the originals) */
        var selectionAfter = doc.selection;
        if (!selectionAfter) return;
        for (var k = 0; k < selectionAfter.length; k++) {
            resultStats.finalSelection.push(selectionAfter[k]);
        }
    }

    /**
     * 失敗したアイテムの種類・名前・エラー内容を1行にまとめる
     * @param {PageItem} item - 失敗したアイテム
     * @param {Error} err - 発生したエラー
     * @returns {string} 失敗の詳細
     */
    function formatFailureMessage(item, err) {
        var itemType = 'Unknown';
        var itemName = '';

        /* 削除済みなどで参照が無効なときは種類・名前を取れないまま続ける / Carry on without type/name when the reference is invalid */
        try {
            itemType = item.typename;
            itemName = item.name || '';
        } catch (e) { }

        var errorMessage = (err && err.message) ? err.message : String(err);
        return itemType + (itemName !== '' ? ' "' + itemName + '"' : '') + ': ' + errorMessage;
    }

    /**
     * 指定したアイテムだけを選択する（既存の選択は解除）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 選択するアイテム
     * @returns {void}
     */
    function applySelection(doc, items) {
        doc.selection = null;
        for (var k = 0; k < items.length; k++) {
            /* 参照が無効になったアイテムは飛ばす / Skip items whose reference is no longer valid */
            try {
                items[k].selected = true;
            } catch (e) { }
        }
    }

    /**
     * 処理結果の件数と失敗の詳細を通知する
     * @param {ResultStats} resultStats - 集計結果
     * @returns {void}
     */
    function showResultSummary(resultStats) {
        var countLines = [
            formatCountLine(LABELS.alert.createdCount, resultStats.createdCount),
            formatCountLine(LABELS.alert.existingCount, resultStats.existingSymbolCount),
            formatCountLine(LABELS.alert.ignoredLinkedCount, resultStats.ignoredLinkedImageCount),
            formatCountLine(LABELS.alert.failedCount, resultStats.failureMessages.length)
        ];
        var message = getLabel(LABELS.alert.completed) + '\n\n' + countLines.join('\n');

        if (resultStats.failureMessages.length > 0) {
            message += '\n\n' + getLabel(LABELS.alert.failureDetails) + '\n' + resultStats.failureMessages.join('\n');
        }

        alert(message);
    }

    /**
     * 件数の行を作る（英語はコロンの後に空白を入れる）
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @param {number} count - 件数
     * @returns {string} 「項目名：件数」の行
     */
    function formatCountLine(labelSet, count) {
        return labelText(labelSet) + (uiLang === 'ja' ? '' : ' ') + count;
    }

    // =========================================
    // 1オブジェクトのシンボル化 / Symbolize one item
    // =========================================

    /**
     * 1オブジェクトを新規シンボルとして登録し、元オブジェクトをそのインスタンスで置き換える
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} originalItem - シンボルにするアイテム
     * @param {number} index - 選択内での位置（連番に使う）
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {SymbolItem} 置き換えたシンボルインスタンス
     */
    function symbolizeOneItem(doc, originalItem, index, symbolizeSettings) {
        /* 埋め込みで境界が変わることがあるので、元の位置を先に控える / Capture bounds first because embedding may change them */
        var originalBounds = originalItem.geometricBounds;
        var workingItem = ensureEmbedded(doc, originalItem);

        var symbolName = resolveSymbolName(doc, workingItem, index, symbolizeSettings);
        var newSymbol = createSymbolFromItem(doc, workingItem, symbolName, symbolizeSettings.referencePoint);
        var newInstance = doc.symbolItems.add(newSymbol);

        /* 元オブジェクトの直前へ移して重なり順を近づける（失敗しても位置合わせは続ける）/ Keep the z-order close to the original */
        try {
            newInstance.move(workingItem, ElementPlacement.PLACEBEFORE);
        } catch (e) { }

        alignToBounds(newInstance, originalBounds);
        workingItem.remove();

        return newInstance;
    }

    /**
     * リンク画像なら埋め込み、置き換わったアイテムを返す（それ以外はそのまま返す）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のアイテム
     * @returns {PageItem} 以後の処理に使うアイテム
     */
    function ensureEmbedded(doc, item) {
        if (item.typename !== 'PlacedItem') return item;

        /* 埋め込みの前後で親の子アイテムを比べ、増えたものを置き換え後のアイテムとする / Diff the parent's children around embed() */
        var parentContainer = item.parent;
        var itemsBeforeEmbed = snapshotChildPageItems(parentContainer);

        item.embed();

        var embeddedItem = findNewPageItem(parentContainer, itemsBeforeEmbed);
        if (embeddedItem) return embeddedItem;

        /* 見つからないときは、埋め込み後に選択される置き換えアイテムを使う / Fall back to the replacement Illustrator selects after embedding */
        if (doc.selection && doc.selection.length > 0) {
            return doc.selection[0];
        }

        throw new Error('Embedded replacement item could not be located.');
    }

    /**
     * 親コンテナ直下のアイテムを配列として控える
     * @param {object} parentContainer - Layer または GroupItem
     * @returns {PageItem[]} 直下のアイテム
     */
    function snapshotChildPageItems(parentContainer) {
        var childItems = [];
        for (var i = 0; i < parentContainer.pageItems.length; i++) {
            childItems.push(parentContainer.pageItems[i]);
        }
        return childItems;
    }

    /**
     * 控えた一覧に無いアイテムを親コンテナ直下から探す
     * @param {object} parentContainer - Layer または GroupItem
     * @param {PageItem[]} itemsBefore - 埋め込み前に控えたアイテム
     * @returns {PageItem|null} 新しく増えたアイテム（見つからないときは null）
     */
    function findNewPageItem(parentContainer, itemsBefore) {
        /* 走査に失敗したときは null を返し、呼び出し側のフォールバックに任せる / Return null so the caller can fall back */
        try {
            for (var i = 0; i < parentContainer.pageItems.length; i++) {
                if (!containsPageItem(itemsBefore, parentContainer.pageItems[i])) {
                    return parentContainer.pageItems[i];
                }
            }
        } catch (e) { }
        return null;
    }

    /**
     * 配列に同じアイテムの参照が含まれるかを調べる
     * @param {PageItem[]} items - 調べる配列
     * @param {PageItem} targetItem - 探すアイテム
     * @returns {boolean} 含まれていれば true
     */
    function containsPageItem(items, targetItem) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === targetItem) return true;
        }
        return false;
    }

    /**
     * アイテムの複製をシンボルとして登録し、複製は破棄する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} sourceItem - 元のアイテム
     * @param {string} symbolName - シンボル名
     * @param {SymbolRegistrationPoint} referencePoint - 基準点
     * @returns {Symbol} 登録したシンボル
     */
    function createSymbolFromItem(doc, sourceItem, symbolName, referencePoint) {
        var duplicatedItem = sourceItem.duplicate();
        var newSymbol = doc.symbols.add(duplicatedItem, referencePoint);
        newSymbol.name = symbolName;

        /* 登録で複製が取り込まれて消えている場合もあるので、失敗は無視する / The duplicate may already be gone after registration */
        try {
            duplicatedItem.remove();
        } catch (e) { }

        return newSymbol;
    }

    /**
     * アイテムの左上が指定した境界の左上に合うよう移動する
     * @param {PageItem} item - 移動するアイテム
     * @param {number[]} targetBounds - 合わせる境界 [左, 上, 右, 下]
     * @returns {void}
     */
    function alignToBounds(item, targetBounds) {
        var currentBounds = item.geometricBounds;
        var dx = targetBounds[0] - currentBounds[0];
        var dy = targetBounds[1] - currentBounds[1];
        item.translate(dx, dy);
    }

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
    // ビュー / View
    // =========================================

    /**
     * 現在のビューのズーム倍率と中心を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {object|null} { zoom, centerPoint }（取得できないときは null）
     */
    function captureViewState(doc) {
        try {
            var view = doc.views[0];
            return {
                zoom: view.zoom,
                centerPoint: [view.centerPoint[0], view.centerPoint[1]]
            };
        } catch (e) {
            return null;
        }
    }

    /**
     * 控えたビューの状態に戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {object|null} viewState - captureViewState() の戻り値
     * @returns {void}
     */
    function restoreViewState(doc, viewState) {
        if (!viewState) return;
        try {
            var view = doc.views[0];
            view.zoom = viewState.zoom;
            view.centerPoint = viewState.centerPoint;
        } catch (e) { }
    }

    /**
     * アイテム全体がビューの中央でおよそ半分を占めるよう、ズームと中心を合わせる
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 対象のアイテム
     * @returns {void}
     */
    function focusViewOnItems(doc, items) {
        /* ビューを寄せられなくても登録は続ける / Registration continues even if the view cannot be focused */
        try {
            /* クリップグループはマスクの範囲で測る / Clip groups are measured by their mask */
            var itemBounds = getClipAwareUnionBounds(items, false);
            if (!itemBounds) return;
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];
            if (itemWidth <= 0 || itemHeight <= 0) return;

            var view = doc.views[0];
            var viewBounds = view.bounds;
            var viewWidth = viewBounds[2] - viewBounds[0];
            var viewHeight = viewBounds[1] - viewBounds[3];

            var targetZoom = Math.min(
                FOCUS_VIEW_RATIO * viewWidth * view.zoom / itemWidth,
                FOCUS_VIEW_RATIO * viewHeight * view.zoom / itemHeight
            );
            view.zoom = Math.max(MIN_VIEW_ZOOM, Math.min(MAX_VIEW_ZOOM, targetZoom));
            view.centerPoint = [(itemBounds[0] + itemBounds[2]) / 2, (itemBounds[1] + itemBounds[3]) / 2];
        } catch (e) { }
    }

    // =========================================
    // シンボル名 / Symbol name
    // =========================================

    /**
     * シンボル名を決める。優先順位はテキスト内容（オン時）→ レイヤー名 → メモ → 接頭辞＋連番
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のアイテム
     * @param {number} index - 選択内での位置（連番に使う）
     * @param {SymbolizeSettings} symbolizeSettings - ダイアログの設定
     * @returns {string} ドキュメント内で重複しないシンボル名
     */
    function resolveSymbolName(doc, item, index, symbolizeSettings) {
        var baseName = '';

        if (symbolizeSettings.useTextAsName) {
            baseName = sanitizeSymbolName(findTextContents(item));
        }
        if (baseName === '') {
            baseName = sanitizeSymbolName(getLayerNameFromItem(item));
        }
        if (baseName === '') {
            baseName = sanitizeSymbolName(getNoteFromItem(item));
        }
        if (baseName === '') {
            baseName = symbolizeSettings.defaultPrefix + zeroPadding(index + 1, symbolizeSettings.sequenceDigits);
        }

        return getUniqueSymbolName(doc, baseName);
    }

    /**
     * テキストフレームの内容、またはグループ内で最初に見つかったテキストフレームの内容を返す
     * @param {PageItem} item - 対象のアイテム
     * @returns {string|null} テキスト内容（テキストが無いときは null）
     */
    function findTextContents(item) {
        if (item.typename === 'TextFrame') return item.contents;
        if (item.typename !== 'GroupItem') return null;

        for (var i = 0; i < item.pageItems.length; i++) {
            var childItem = item.pageItems[i];
            if (childItem.typename === 'TextFrame') return childItem.contents;
            if (childItem.typename === 'GroupItem') {
                var nestedContents = findTextContents(childItem);
                if (nestedContents) return nestedContents;
            }
        }
        return null;
    }

    /**
     * アイテムが属するレイヤー名を返す（「レイヤー 1」などの既定名は除く）
     * @param {PageItem} item - 対象のアイテム
     * @returns {string|null} レイヤー名（使えないときは null）
     */
    function getLayerNameFromItem(item) {
        try {
            var layerName = item.layer.name;
            if (!layerName || DEFAULT_LAYER_NAME_PATTERN.test(layerName)) return null;
            return layerName;
        } catch (e) {
            return null;
        }
    }

    /**
     * アイテムのメモ（note）を返す（空白だけのものは除く）
     * @param {PageItem} item - 対象のアイテム
     * @returns {string|null} メモ（使えないときは null）
     */
    function getNoteFromItem(item) {
        try {
            var noteText = item.note;
            if (!noteText || trimString(String(noteText)) === '') return null;
            return noteText;
        } catch (e) {
            return null;
        }
    }

    /**
     * シンボル名として使えるよう整える（改行・タブを空白に、前後の空白を除き、最大文字数で切る）
     * @param {string|null} name - 元の文字列
     * @returns {string} 整えた名前（元が無いときは空文字）
     */
    function sanitizeSymbolName(name) {
        if (name === null || name === undefined) return '';
        var cleanName = trimString(String(name).replace(/[\r\n\t]+/g, ' '));
        return cleanName.substring(0, SYMBOL_NAME_MAX_LENGTH);
    }

    /**
     * 数値を指定した桁数までゼロで埋める
     * @param {number} num - 数値
     * @param {number} length - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function zeroPadding(num, length) {
        var paddedText = String(num);
        while (paddedText.length < length) paddedText = '0' + paddedText;
        return paddedText;
    }

    /**
     * 既存のシンボル名と重ならない名前を返す（重なるときは "_2"、"_3" … を付ける）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} baseName - 元の名前
     * @returns {string} 重複しない名前
     */
    function getUniqueSymbolName(doc, baseName) {
        var uniqueName = baseName;
        var suffixNumber = 2;
        while (symbolExists(doc, uniqueName)) {
            uniqueName = baseName + '_' + suffixNumber;
            suffixNumber++;
        }
        return uniqueName;
    }

    /**
     * 同じ名前のシンボルがあるかを調べる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} symbolName - シンボル名
     * @returns {boolean} あれば true
     */
    function symbolExists(doc, symbolName) {
        /* getByName は見つからないと例外を投げる / getByName throws when not found */
        try {
            doc.symbols.getByName(symbolName);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 前後の空白を除く（ES3 に String.trim が無いため）
     * @param {string} text - 元の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimString(text) {
        return text.replace(/^\s+|\s+$/g, '');
    }

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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを表示する
     * @returns {SymbolizeSettings|null} 設定（キャンセル時は null）
     */
    function showSettingsDialog() {
        var settingsDialog = new Window('dialog', getLabel(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        setupWindow(settingsDialog);

        var dialogControls = {
            groupModeRadios: buildGroupModePanel(settingsDialog),
            symbolizeModeRadios: buildSymbolizeModePanel(settingsDialog),
            symbolNameControls: buildSymbolNamePanel(settingsDialog)
        };

        /* 基準点とリンク画像を横に並べる / Place registration point and linked images side by side */
        var columnsGroup = settingsDialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignChildren = ['fill', 'fill'];
        columnsGroup.spacing = COLUMN_SPACING;
        dialogControls.anchorWidget = buildReferencePointPanel(columnsGroup);
        dialogControls.linkedImageRadios = buildLinkedImagePanel(columnsGroup);

        bindModeControls(dialogControls);

        var buttonRow = addButtonRow(settingsDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel(LABELS.button.cancel), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel(LABELS.button.ok), { name: 'ok' });
        settingsDialog.cancelElement = btnCancel;
        settingsDialog.defaultElement = btnOK;

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(settingsDialog, SCRIPT_NAME);
        if (settingsDialog.show() !== 1) return null;
        return readDialogSettings(dialogControls);
    }

    /**
     * ダイアログのコントロールから設定を読み取る
     * @param {object} dialogControls - showSettingsDialog() で集めたコントロール
     * @returns {SymbolizeSettings} 設定
     */
    function readDialogSettings(dialogControls) {
        var nameControls = dialogControls.symbolNameControls;
        var prefixValue = trimString(nameControls.prefixInput.text);

        return {
            groupMode: getSelectedRadioValue(dialogControls.groupModeRadios),
            symbolizeMode: getSelectedRadioValue(dialogControls.symbolizeModeRadios),
            defaultPrefix: (prefixValue === '') ? DEFAULT_PREFIX : prefixValue,
            sequenceDigits: getSelectedRadioValue(nameControls.sequenceRadios),
            useTextAsName: nameControls.useTextCheckbox.value,
            referencePoint: getAnchorSymbolRegistrationPoint(getAnchorWidgetIndex(dialogControls.anchorWidget)),
            linkedImagePolicy: getSelectedRadioValue(dialogControls.linkedImageRadios)
        };
    }

    /**
     * 選択範囲の扱いと登録方法の連動、自動登録専用の項目の有効／無効を設定する
     * @param {object} dialogControls - showSettingsDialog() で集めたコントロール
     * @returns {void}
     */
    function bindModeControls(dialogControls) {
        var groupModeRadios = dialogControls.groupModeRadios;
        var symbolizeModeRadios = dialogControls.symbolizeModeRadios;
        var nameControls = dialogControls.symbolNameControls;
        var anchorWidget = dialogControls.anchorWidget;
        var i;

        /**
         * 現在の選択に合わせて、登録方法と自動登録専用の項目を切り替える
         * @returns {void}
         */
        function refreshModeUI() {
            var asGroupMode = (getSelectedRadioValue(groupModeRadios) === GROUP_MODE.AS_GROUP);

            /* まとめて1つにするときは標準ダイアログでの確認に固定する / One symbol from selection always uses the native dialog */
            if (asGroupMode) selectRadioByValue(symbolizeModeRadios, SYMBOLIZE_MODE.INDIVIDUAL);
            for (var k = 0; k < symbolizeModeRadios.length; k++) {
                symbolizeModeRadios[k].enabled = !asGroupMode;
            }

            /* 接頭辞・連番・テキスト流用・基準点は自動登録のときだけ使う / Prefix, sequence, text reuse and registration point apply to automatic registration only */
            var batchEnabled = (getSelectedRadioValue(symbolizeModeRadios) === SYMBOLIZE_MODE.BATCH);
            nameControls.prefixRow.enabled = batchEnabled;
            nameControls.sequenceRow.enabled = batchEnabled;
            nameControls.useTextCheckbox.enabled = batchEnabled;
            /* パネルの enabled は子の onDraw に伝わらないので描き直す / The panel's enabled state does not redraw the widget by itself */
            anchorWidget.parent.enabled = batchEnabled;
            redrawAnchorWidgetsIn(anchorWidget.parent);
        }

        for (i = 0; i < groupModeRadios.length; i++) {
            groupModeRadios[i].onClick = function () {
                selectRadioByValue(groupModeRadios, this.optionValue);
                /* オブジェクトごとに戻したときは自動登録を選び直す / Switch back to automatic registration for per-object mode */
                if (this.optionValue === GROUP_MODE.EACH_ITEM) {
                    selectRadioByValue(symbolizeModeRadios, SYMBOLIZE_MODE.BATCH);
                }
                refreshModeUI();
            };
        }
        for (i = 0; i < symbolizeModeRadios.length; i++) {
            symbolizeModeRadios[i].onClick = function () {
                selectRadioByValue(symbolizeModeRadios, this.optionValue);
                refreshModeUI();
            };
        }

        refreshModeUI();
    }

    /**
     * 選択範囲の扱いパネル（まとめて1つ／オブジェクトごと）を追加する
     * @param {Window} parent - 追加先
     * @returns {RadioButton[]} ラジオボタン
     */
    function buildGroupModePanel(parent) {
        var groupModePanel = addPanel(parent, LABELS.panel.groupMode, 6);
        return addOptionRadios(groupModePanel, [
            { value: GROUP_MODE.AS_GROUP, label: LABELS.radio.asGroup, tooltip: LABELS.tooltip.asGroup },
            { value: GROUP_MODE.EACH_ITEM, label: LABELS.radio.eachItem, tooltip: LABELS.tooltip.eachItem }
        ], DEFAULT_GROUP_MODE);
    }

    /**
     * 登録方法パネル（標準ダイアログで確認／自動登録）を追加する
     * @param {Window} parent - 追加先
     * @returns {RadioButton[]} ラジオボタン
     */
    function buildSymbolizeModePanel(parent) {
        var symbolizeModePanel = addPanel(parent, LABELS.panel.symbolizeMode);
        var symbolizeModeRow = symbolizeModePanel.add('group');
        setupRow(symbolizeModeRow);
        return addOptionRadios(symbolizeModeRow, [
            { value: SYMBOLIZE_MODE.INDIVIDUAL, label: LABELS.radio.individual, tooltip: LABELS.tooltip.individual },
            { value: SYMBOLIZE_MODE.BATCH, label: LABELS.radio.batch, tooltip: LABELS.tooltip.batch }
        ], DEFAULT_SYMBOLIZE_MODE);
    }

    /**
     * シンボル名パネル（接頭辞・連番の桁数・テキスト流用）を追加する
     * @param {Window} parent - 追加先
     * @returns {object} { prefixRow, prefixInput, sequenceRow, sequenceRadios, useTextCheckbox }
     */
    function buildSymbolNamePanel(parent) {
        var symbolNamePanel = addPanel(parent, LABELS.panel.symbolName, 6);

        var prefixRow = symbolNamePanel.add('group');
        setupRow(prefixRow, 'left', 8);
        addFieldLabel(prefixRow, LABELS.fieldLabel.prefix).helpTip = getLabel(LABELS.tooltip.prefix);
        var prefixInput = prefixRow.add('edittext', undefined, DEFAULT_PREFIX);
        prefixInput.preferredSize.width = PREFIX_INPUT_WIDTH;
        prefixInput.helpTip = getLabel(LABELS.tooltip.prefix);
        prefixInput.active = true;

        var sequenceRow = symbolNamePanel.add('group');
        setupRow(sequenceRow, 'left', 8);
        addFieldLabel(sequenceRow, LABELS.fieldLabel.sequence).helpTip = getLabel(LABELS.tooltip.sequence);
        var sequenceOptions = [];
        for (var i = 0; i < SEQUENCE_DIGIT_OPTIONS.length; i++) {
            sequenceOptions.push({
                value: SEQUENCE_DIGIT_OPTIONS[i],
                labelString: zeroPadding(0, SEQUENCE_DIGIT_OPTIONS[i]),
                tooltip: LABELS.tooltip.sequence
            });
        }
        var sequenceRadios = addOptionRadios(sequenceRow, sequenceOptions, DEFAULT_SEQUENCE_DIGITS);

        var useTextCheckbox = symbolNamePanel.add('checkbox', undefined, getLabel(LABELS.checkbox.useTextAsName));
        useTextCheckbox.value = DEFAULT_USE_TEXT_AS_NAME;
        useTextCheckbox.helpTip = getLabel(LABELS.tooltip.useTextAsName);

        return {
            prefixRow: prefixRow,
            prefixInput: prefixInput,
            sequenceRow: sequenceRow,
            sequenceRadios: sequenceRadios,
            useTextCheckbox: useTextCheckbox
        };
    }

    /**
     * 基準点パネル（9軸ウィジェット）を追加する
     * @param {Group} parent - 追加先
     * @returns {Button} 基準点ウィジェット（値は getAnchorWidgetIndex() で読む）
     */
    function buildReferencePointPanel(parent) {
        var referencePointPanel = addPanel(parent, LABELS.panel.referencePoint);
        referencePointPanel.alignChildren = ['center', 'center'];
        referencePointPanel.helpTip = getLabel(LABELS.tooltip.referencePoint);

        var anchorWidget = addAnchorWidget(referencePointPanel, DEFAULT_REFERENCE_POINT_INDEX);
        anchorWidget.helpTip = getLabel(LABELS.tooltip.referencePoint);
        return anchorWidget;
    }

    /**
     * リンク画像パネル（無視／埋め込んで登録）を追加する
     * @param {Group} parent - 追加先
     * @returns {RadioButton[]} ラジオボタン
     */
    function buildLinkedImagePanel(parent) {
        var linkedImagePanel = addPanel(parent, LABELS.panel.linkedImage, 6);
        linkedImagePanel.helpTip = getLabel(LABELS.tooltip.linkedImage);
        return addOptionRadios(linkedImagePanel, [
            { value: LINKED_IMAGE_POLICY.IGNORE, label: LABELS.radio.ignoreLinked, tooltip: LABELS.tooltip.linkedImage },
            { value: LINKED_IMAGE_POLICY.EMBED, label: LABELS.radio.embedLinked, tooltip: LABELS.tooltip.linkedImage }
        ], DEFAULT_LINKED_IMAGE_POLICY);
    }

    // =========================================
    // ラジオボタン / Radio buttons
    // =========================================

    /**
     * 選択肢ごとにラジオボタンを追加し、クリックで排他選択になるようにする
     * @param {object} parent - 追加先
     * @param {object[]} optionDefs - { value, label（ja/en）または labelString, tooltip（ja/en） } の配列
     * @param {*} selectedValue - 最初に選ぶ値
     * @returns {RadioButton[]} ラジオボタン
     */
    function addOptionRadios(parent, optionDefs, selectedValue) {
        var radios = [];
        for (var i = 0; i < optionDefs.length; i++) {
            var optionDef = optionDefs[i];
            var radio = parent.add('radiobutton', undefined, optionDef.labelString || getLabel(optionDef.label));
            radio.optionValue = optionDef.value;
            radio.helpTip = getLabel(optionDef.tooltip);
            radio.onClick = function () {
                selectRadioByValue(radios, this.optionValue);
            };
            radios.push(radio);
        }
        selectRadioByValue(radios, selectedValue);
        return radios;
    }

    /**
     * 指定した値のラジオボタンだけを選択する
     * @param {RadioButton[]} radios - ラジオボタン
     * @param {*} optionValue - 選ぶ値
     * @returns {void}
     */
    function selectRadioByValue(radios, optionValue) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (radios[i].optionValue === optionValue);
        }
    }

    /**
     * 選択中のラジオボタンの値を返す
     * @param {RadioButton[]} radios - ラジオボタン
     * @returns {*} 選択中の値（未選択のときは先頭の値）
     */
    function getSelectedRadioValue(radios) {
        for (var i = 0; i < radios.length; i++) {
            if (radios[i].value) return radios[i].optionValue;
        }
        return radios[0].optionValue;
    }

    main();

})();
