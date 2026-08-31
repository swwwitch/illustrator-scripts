#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの端または中心を、アクティブアートボードの端・中央、または条件に合うガイドへ整列します。
整列先は3×3の9点から選べるほか、矢印キーでの1段階ずつの送りやファイル名からの自動判定にも対応します。

詳細は README を参照してください。

### Overview

Aligns the edges or the center of the selected objects to the edge or the center of the active artboard, or to a matching guide.
The target is picked from a 3x3 grid of nine points, stepped one target at a time with the arrow keys, or derived from the filename.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEdgeAlign";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-31";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlign.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlign.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4ae0e1e70481"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var SHOW_DIALOG                 = true;     /* ダイアログを表示する / show the dialog */
    var USE_GUIDES                  = true;     /* ガイドも整列先に含める / include guides as targets */
    var DEFAULT_ALIGNMENT_SIDE      = "right";  /* ファイル名から判定できないときの整列先 / fallback target */
    var GUIDE_SEARCH_MODE           = "inside"; /* "inside"＝進行方向の直近 / "nearest"＝最も近い */
    var GUIDE_ORIENTATION_TOLERANCE = 0.01;     /* ガイドの水平・垂直判定に使う許容値 / orientation tolerance */

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS        = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING        = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS         = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING         = 6;                /* パネル内の要素間隔 / panel spacing */
    var ANCHOR_PANEL_MARGINS  = [9, 13, 9, 4];    /* 9軸ウィジェット用の詰めた余白 / tighter padding for the anchor widget */
    var BUTTON_ROW_TOP_MARGIN = 10;               /* ボタンエリアの上余白 / top margin of the button row */

    /* 9軸ウィジェットの寸法（onDrawで描画） / 9-axis widget metrics (drawn via onDraw) */
    var ANCHOR_WIDGET_SIZE = 66;
    var ANCHOR_CELL_SIZE   = 9;
    var ANCHOR_CELL_GAP    = 7.5;

    // =========================================
    // 整列先の定義 / Alignment targets
    // =========================================

    /* 3×3の各セル（行優先）が担う水平・垂直の整列先とショートカットキー。
       各要素はそのまま整列軸（{ horizontal, vertical }）として使う
       Horizontal/vertical targets and shortcut key per cell of the 3x3 grid (row-major);
       each entry doubles as the { horizontal, vertical } pair used for aligning */
    var ANCHOR_DEFINITIONS = [
        { horizontal: "left",     vertical: "top",      shortcutKey: "W" },
        { horizontal: "CENTER_X", vertical: "top",      shortcutKey: "E" },
        { horizontal: "right",    vertical: "top",      shortcutKey: "R" },
        { horizontal: "left",     vertical: "CENTER_Y", shortcutKey: "S" },
        { horizontal: "CENTER_X", vertical: "CENTER_Y", shortcutKey: "D" },
        { horizontal: "right",    vertical: "CENTER_Y", shortcutKey: "F" },
        { horizontal: "left",     vertical: "bottom",   shortcutKey: "X" },
        { horizontal: "CENTER_X", vertical: "bottom",   shortcutKey: "C" },
        { horizontal: "right",    vertical: "bottom",   shortcutKey: "V" }
    ];

    /* 9軸ウィジェットが未選択のときのインデックス / Index used while no cell is selected */
    var NO_ANCHOR_INDEX = -1;

    /* 整列先ごとの座標の取り出し方。
       axis＝動かす軸、boundsIndex＝境界配列 [左,上,右,下] のインデックス（null は2辺の中点）、
       ahead＝揃える向き（座標が増える向きなら +1、中央揃えは 0＝ガイド吸着の対象外）
       axis is the axis to move, boundsIndex indexes [L,T,R,B] (null means the midpoint of two edges),
       ahead is the sign of the direction aligned toward (0 disables guide snapping) */
    var EDGE_RULES = {
        "left":     { axis: "x", boundsIndex: 0,    ahead: -1 },
        "right":    { axis: "x", boundsIndex: 2,    ahead:  1 },
        "top":      { axis: "y", boundsIndex: 1,    ahead:  1 },
        "bottom":   { axis: "y", boundsIndex: 3,    ahead: -1 },
        "CENTER_X": { axis: "x", boundsIndex: null, ahead:  0 },
        "CENTER_Y": { axis: "y", boundsIndex: null, ahead:  0 }
    };

    /* 整列先1つを水平・垂直の軸に展開した対応表。null はその軸を動かさない
       Each target expanded into horizontal/vertical axes; null leaves that axis alone */
    var AXES_BY_ALIGNMENT_SIDE = {
        "left":     { horizontal: "left",     vertical: null },
        "right":    { horizontal: "right",    vertical: null },
        "top":      { horizontal: null,       vertical: "top" },
        "bottom":   { horizontal: null,       vertical: "bottom" },
        "CENTER_X": { horizontal: "CENTER_X", vertical: null },
        "CENTER_Y": { horizontal: null,       vertical: "CENTER_Y" },
        "CENTER":   { horizontal: "CENTER_X", vertical: "CENTER_Y" }
    };

    /* ファイル名に含まれるキーワードと整列先。CENTERX / CENTERY を CENTER より先に判定する
       Filename keywords mapped to targets; CENTERX/CENTERY are matched before CENTER */
    var ALIGNMENT_SIDE_BY_FILENAME = [
        { keyword: "CENTERX", side: "CENTER_X" },
        { keyword: "CENTERY", side: "CENTER_Y" },
        { keyword: "CENTER",  side: "CENTER" },
        { keyword: "LEFT",    side: "left" },
        { keyword: "RIGHT",   side: "right" },
        { keyword: "TOP",     side: "top" },
        { keyword: "BOTTOM",  side: "bottom" }
    ];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードに整列", en: "Align to Artboard" }
        },
        panel: {
            alignment: { ja: "整列", en: "Alignment" },
            option:    { ja: "オプション", en: "Options" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            useGuides:     { ja: "ガイドを使用", en: "Use guides" },
            preview:       { ja: "プレビュー", en: "Preview" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noSelection:            { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
            generalError:           { ja: "エラーが発生しました", en: "An error occurred" },
            invalidGuideSearchMode: { ja: "GUIDE_SEARCH_MODE の指定が不正です", en: "Invalid GUIDE_SEARCH_MODE" }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.option" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var keyIndex = 0; keyIndex < pathKeys.length; keyIndex++) {
            labelNode = labelNode[pathKeys[keyIndex]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
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
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel", undefined, panelTitle);
        setupPanel(createdPanel);
        return createdPanel;
    }

    // =========================================
    // 9軸ウィジェットの描画 / Anchor widget drawing
    // =========================================

    /* 外周の□どうしをつなぐケイ線の組み合わせ（中央は独立） / Rules joining the outer squares (center stands alone) */
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    /* UIの明暗に合わせて initAnchorColors() で設定 / Set from the light/dark UI in initAnchorColors() */
    var anchorLineColor = [0.6, 0.6, 0.6, 1];
    var anchorSelectedFillColor = [0.4, 0.4, 0.4, 1];

    /**
     * 9軸ウィジェットの配色をIllustratorのUIの明暗に合わせて決める
     * @returns {void}
     */
    function initAnchorColors() {
        var lightUI = app.preferences.getRealPreference("uiBrightness") > 0.5;
        /* 選択セルの塗り：ライトは濃いグレー、ダークは明るいグレー / Selected-cell fill: dark gray in light UI, bright gray in dark UI */
        anchorLineColor = lightUI ? [0.6, 0.6, 0.6, 1] : [0.55, 0.55, 0.55, 1];
        anchorSelectedFillColor = lightUI ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
    }

    /**
     * 3×3のグリッド位置を 0〜2 に収める
     * @param {number} gridPosition - 計算した行または列
     * @returns {number} 0〜2 に丸めた値
     */
    function clampGridIndex(gridPosition) {
        if (gridPosition < 0) return 0;
        if (gridPosition > 2) return 2;
        return gridPosition;
    }

    /**
     * 9軸ウィジェットを再描画する
     * @param {Button} anchorWidget - 対象ウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        /* notify は環境により例外を投げ得るので保護 / notify can throw in some environments */
        try {
            anchorWidget.notify("onDraw");
        } catch (e) {}
    }

    /**
     * 9軸ウィジェットのセルを1つ描画する（選択時のみ塗りつぶす）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number} cellX - セルの左端
     * @param {number} cellY - セルの上端
     * @param {boolean} isSelected - 選択中なら true
     * @returns {void}
     */
    function drawAnchorCell(graphics, cellX, cellY, isSelected) {
        /**
         * セルの四角形パスを作る
         * @returns {void}
         */
        function addCellPath() {
            graphics.newPath();
            graphics.moveTo(cellX, cellY);
            graphics.lineTo(cellX + ANCHOR_CELL_SIZE, cellY);
            graphics.lineTo(cellX + ANCHOR_CELL_SIZE, cellY + ANCHOR_CELL_SIZE);
            graphics.lineTo(cellX, cellY + ANCHOR_CELL_SIZE);
            graphics.closePath();
        }

        /* 選択中のみ塗る（枠は塗りの上に描く） / Fill only when selected; the border is drawn over the fill */
        if (isSelected) {
            addCellPath();
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, anchorSelectedFillColor));
        }
        addCellPath();
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, anchorLineColor, 1));
    }

    /**
     * 9軸ウィジェットを描画する（外周の□をケイ線でつなぎ、選択セルを塗る）
     * @param {Button} anchorWidget - 対象ウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];

        /* 背景はコントロールの地色で塗り、パネルと同色に見せる / Paint the control background so the widget blends into the panel */
        try {
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);

        /**
         * セルの左上座標を返す
         * @param {number} cellIndex - 0〜8のセル番号（行優先）
         * @returns {number[]} [X座標, Y座標]
         */
        function getCellOrigin(cellIndex) {
            return [
                originX + (cellIndex % 3) * cellStep,
                originY + Math.floor(cellIndex / 3) * cellStep
            ];
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, anchorLineColor, 1);
        for (var connectionIndex = 0; connectionIndex < ANCHOR_CONNECTIONS.length; connectionIndex++) {
            var fromCell = getCellOrigin(ANCHOR_CONNECTIONS[connectionIndex][0]);
            var toCell = getCellOrigin(ANCHOR_CONNECTIONS[connectionIndex][1]);
            var isHorizontal = (ANCHOR_CONNECTIONS[connectionIndex][1] - ANCHOR_CONNECTIONS[connectionIndex][0] === 1);
            graphics.newPath();
            if (isHorizontal) {
                /* 横方向：右隣の□へ / Horizontal: to the square on the right */
                graphics.moveTo(fromCell[0] + ANCHOR_CELL_SIZE, fromCell[1] + ANCHOR_CELL_SIZE / 2);
                graphics.lineTo(toCell[0], toCell[1] + ANCHOR_CELL_SIZE / 2);
            } else {
                /* 縦方向：下の□へ / Vertical: to the square below */
                graphics.moveTo(fromCell[0] + ANCHOR_CELL_SIZE / 2, fromCell[1] + ANCHOR_CELL_SIZE);
                graphics.lineTo(toCell[0] + ANCHOR_CELL_SIZE / 2, toCell[1]);
            }
            graphics.strokePath(linePen);
        }

        for (var cellIndex = 0; cellIndex < ANCHOR_DEFINITIONS.length; cellIndex++) {
            var cellOrigin = getCellOrigin(cellIndex);
            drawAnchorCell(graphics, cellOrigin[0], cellOrigin[1], cellIndex === anchorWidget.selectedAnchorIndex);
        }
    }

    // =========================================
    // 境界と整列先の座標 / Bounds and target values
    // =========================================

    /**
     * オブジェクト1つの境界を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(pageItem, usePreviewBounds) {
        /* クリッピンググループはマスク形状（先頭アイテム）の幾何境界を使う
           A clipping group is measured by the geometric bounds of its mask (the first item) */
        if (pageItem.typename === "GroupItem" && pageItem.clipped === true) {
            return pageItem.pageItems[0].geometricBounds;
        }
        return usePreviewBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
    }

    /**
     * 選択オブジェクト群を包含する境界を返す
     * @param {Array} pageItems - 対象オブジェクトの配列
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {number[]} [左, 上, 右, 下]
     */
    function computeSelectionBounds(pageItems, usePreviewBounds) {
        var unionBounds = null;
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            var itemBounds = getItemBounds(pageItems[itemIndex], usePreviewBounds);
            if (!unionBounds) {
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
     * 整列先に対応する座標を境界から取り出す（中央揃えは2辺の中点）
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {number} 座標。対応しない整列先なら null
     */
    function getAlignmentValue(bounds, alignmentSide) {
        var edgeRule = EDGE_RULES[alignmentSide];
        if (!edgeRule) return null;
        if (edgeRule.boundsIndex !== null) return bounds[edgeRule.boundsIndex];
        return (edgeRule.axis === "x") ? (bounds[0] + bounds[2]) / 2 : (bounds[1] + bounds[3]) / 2;
    }

    /**
     * 基準の座標より揃える向きの先にあるかを返す
     * @param {number} value - 判定する座標
     * @param {number} referenceValue - 基準の座標
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {boolean} 先にあれば true
     */
    function isAheadOnSide(value, referenceValue, alignmentSide) {
        return (value - referenceValue) * EDGE_RULES[alignmentSide].ahead > 0;
    }

    // =========================================
    // ガイドの探索 / Guide lookup
    // =========================================

    /**
     * 整列方向に対応するガイド座標を返す（向きが合わないガイドは対象外）
     * @param {number[]} guideBounds - ガイドの幾何境界 [左, 上, 右, 下]
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {number} ガイドの座標。対象外なら null
     */
    function getGuideValue(guideBounds, alignmentSide) {
        if (EDGE_RULES[alignmentSide].axis === "x") {
            /* 左右の整列先になるのは縦ガイド（幅ゼロ）だけ / Only vertical guides (zero width) serve left/right */
            return (Math.abs(guideBounds[2] - guideBounds[0]) <= GUIDE_ORIENTATION_TOLERANCE) ? guideBounds[0] : null;
        }
        /* 上下の整列先になるのは横ガイド（高さゼロ）だけ / Only horizontal guides (zero height) serve top/bottom */
        return (Math.abs(guideBounds[1] - guideBounds[3]) <= GUIDE_ORIENTATION_TOLERANCE) ? guideBounds[1] : null;
    }

    /**
     * ガイド座標がアクティブアートボードの内側にあるかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {boolean} 内側なら true
     */
    function isGuideInsideArtboard(guideValue, artboardRect, alignmentSide) {
        var isHorizontalAxis = (EDGE_RULES[alignmentSide].axis === "x");
        var lowerBound = isHorizontalAxis ? artboardRect[0] : artboardRect[3];
        var upperBound = isHorizontalAxis ? artboardRect[2] : artboardRect[1];
        return guideValue >= lowerBound && guideValue <= upperBound;
    }

    /**
     * アートボード内側のガイドから、整列先として使う座標を探す
     * @param {number} selectionEdge - 選択範囲の境界値
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @param {Object} alignContext - 整列コンテキスト
     * @returns {number} 吸着先の座標。見つからなければ null
     */
    function findGuideSnapValue(selectionEdge, alignmentSide, alignContext) {
        var nearestGuideValue = null;
        var nearestGuideDistance = null;
        var guidePathItems = alignContext.documentRef.pathItems;
        var insideOnly = (GUIDE_SEARCH_MODE === "inside");

        for (var guidePathIndex = 0; guidePathIndex < guidePathItems.length; guidePathIndex++) {
            var guidePathItem = guidePathItems[guidePathIndex];
            if (guidePathItem.guides !== true) continue;

            var guideValue = getGuideValue(guidePathItem.geometricBounds, alignmentSide);
            if (guideValue === null) continue;
            if (!isGuideInsideArtboard(guideValue, alignContext.artboardRect, alignmentSide)) continue;
            /* "inside" は揃える向きの先にあるガイドだけを候補にする（"nearest" は向きを問わない）
               "inside" only accepts guides ahead of the selection; "nearest" takes either direction */
            if (insideOnly && !isAheadOnSide(guideValue, selectionEdge, alignmentSide)) continue;

            var guideDistance = Math.abs(guideValue - selectionEdge);
            if (nearestGuideDistance === null || guideDistance < nearestGuideDistance) {
                nearestGuideValue = guideValue;
                nearestGuideDistance = guideDistance;
            }
        }
        return nearestGuideValue;
    }

    // =========================================
    // 整列の適用 / Applying the alignment
    // =========================================

    /**
     * 整列コンテキストを作る
     * @param {Document} documentRef - 対象ドキュメント
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [左, 上, 右, 下]
     * @param {boolean} useGuides - ガイドを整列先に含めるなら true
     * @returns {Object} 整列コンテキスト
     */
    function createAlignContext(documentRef, artboardRect, useGuides) {
        return { documentRef: documentRef, artboardRect: artboardRect, useGuides: useGuides };
    }

    /**
     * 1軸分の整列オフセットを計算する
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @param {number[]} selectionBounds - 選択範囲の境界 [左, 上, 右, 下]
     * @param {Object} alignContext - 整列コンテキスト
     * @returns {number} 移動量
     */
    function computeAxisOffset(alignmentSide, selectionBounds, alignContext) {
        var selectionValue = getAlignmentValue(selectionBounds, alignmentSide);
        var targetValue = getAlignmentValue(alignContext.artboardRect, alignmentSide);
        if (selectionValue === null || targetValue === null) return 0;

        /* ガイドが吸着先になるのは端揃えのときだけ（中央揃えは ahead が 0）
           Guides only snap for edge alignment; center alignment has ahead 0 */
        if (alignContext.useGuides && EDGE_RULES[alignmentSide].ahead !== 0) {
            var guideValue = findGuideSnapValue(selectionValue, alignmentSide, alignContext);
            if (guideValue !== null) targetValue = guideValue;
        }
        return targetValue - selectionValue;
    }

    /**
     * 選択オブジェクトの現在位置を控える
     * @param {Array} pageItems - 対象オブジェクトの配列
     * @returns {Array<number[]>} [X, Y] の配列
     */
    function captureItemPositions(pageItems) {
        var capturedPositions = [];
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            var itemPosition = pageItems[itemIndex].position;
            capturedPositions.push([itemPosition[0], itemPosition[1]]);
        }
        return capturedPositions;
    }

    /**
     * 控えた位置へオブジェクトを戻す
     * @param {Array} pageItems - 対象オブジェクトの配列
     * @param {Array<number[]>} capturedPositions - captureItemPositions() の戻り値
     * @returns {void}
     */
    function restoreItemPositions(pageItems, capturedPositions) {
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            pageItems[itemIndex].position = capturedPositions[itemIndex];
        }
    }

    /**
     * 整列を1回分適用する（境界の計測からオフセットの適用まで）
     * @param {Array} pageItems - 対象オブジェクトの配列
     * @param {Object} alignmentAxes - 整列軸 { horizontal, vertical }
     * @param {Object} alignContext - 整列コンテキスト
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {void}
     */
    function applyAlignment(pageItems, alignmentAxes, alignContext, usePreviewBounds) {
        var selectionBounds = computeSelectionBounds(pageItems, usePreviewBounds);
        var offsetX = alignmentAxes.horizontal ? computeAxisOffset(alignmentAxes.horizontal, selectionBounds, alignContext) : 0;
        var offsetY = alignmentAxes.vertical ? computeAxisOffset(alignmentAxes.vertical, selectionBounds, alignContext) : 0;
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            pageItems[itemIndex].translate(offsetX, offsetY);
        }
    }

    /**
     * スクリプトのファイル名から整列先を判定する（例: GroupEdgeAlignRIGHT.jsx → "right"）
     * @returns {string} 整列先。判定できない場合は DEFAULT_ALIGNMENT_SIDE
     */
    function detectAlignmentSideFromFileName() {
        var fileNameUpper = File($.fileName).name.toUpperCase();
        for (var keywordIndex = 0; keywordIndex < ALIGNMENT_SIDE_BY_FILENAME.length; keywordIndex++) {
            if (fileNameUpper.indexOf(ALIGNMENT_SIDE_BY_FILENAME[keywordIndex].keyword) !== -1) {
                return ALIGNMENT_SIDE_BY_FILENAME[keywordIndex].side;
            }
        }
        return DEFAULT_ALIGNMENT_SIDE;
    }

    /**
     * プレビューと矢印キーのステップ移動をまとめた整列セッションを作る
     * @param {Array} pageItems - 対象オブジェクトの配列
     * @param {Document} documentRef - 対象ドキュメント
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [左, 上, 右, 下]
     * @returns {Object} preview / step / restoreOriginal をまとめたオブジェクト
     */
    function createAlignmentSession(pageItems, documentRef, artboardRect) {
        /* originalPositions はキャンセル時の完全復元用（不変）、
           basePositions はプレビュー復元の基準で、矢印キーのステップ移動のたびに進む
           originalPositions restores everything on Cancel; basePositions is the preview baseline
           and advances with each arrow-key step */
        var originalPositions = captureItemPositions(pageItems);
        var basePositions = captureItemPositions(pageItems);

        /**
         * 基準位置に戻してからプレビューの整列を適用する
         * @param {Object} settings - 整列軸・境界・ガイドの設定。null なら復元のみ
         * @returns {void}
         */
        function preview(settings) {
            restoreItemPositions(pageItems, basePositions);
            if (settings && settings.alignmentAxes) {
                var alignContext = createAlignContext(documentRef, artboardRect, settings.useGuides);
                applyAlignment(pageItems, settings.alignmentAxes, alignContext, settings.usePreviewBounds);
            }
            app.redraw();
        }

        /**
         * 矢印キーによる1段階の整列（スクリプトを1回実行したのと同じ挙動）
         * @param {string} alignmentSide - "left" / "right" / "top" / "bottom"
         * @param {Object} settings - 境界とガイドの設定
         * @returns {void}
         */
        function step(alignmentSide, settings) {
            preview({
                alignmentAxes: AXES_BY_ALIGNMENT_SIDE[alignmentSide],
                usePreviewBounds: settings.usePreviewBounds,
                useGuides: settings.useGuides
            });
            /* 動いた先を次の基準にして、押すたびにさらに先へ進めるようにする
               The new position becomes the baseline so each press advances further */
            basePositions = captureItemPositions(pageItems);
        }

        /**
         * 矢印キーでのステップ移動も含めて、実行前の位置へ戻す
         * @returns {void}
         */
        function restoreOriginal() {
            restoreItemPositions(pageItems, originalPositions);
            app.redraw();
        }

        return { preview: preview, step: step, restoreOriginal: restoreOriginal };
    }

    // =========================================
    // ダイアログ UI / Dialog UI
    // =========================================

    /**
     * 整列オプションのダイアログを表示し、選択内容を返す
     * @param {Object} initialSettings - 境界とガイドの初期値
     * @param {Object} alignmentSession - プレビューとステップ移動を担う整列セッション
     * @returns {Object} 整列軸・境界・ガイドの設定。キャンセル時は null
     */
    function showAlignmentDialog(initialSettings, alignmentSession) {
        initAnchorColors();

        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = WINDOW_MARGINS;
        dialog.spacing = WINDOW_SPACING;

        var alignmentPanel = addPanel(dialog, getLabel("panel.alignment"));
        alignmentPanel.margins = ANCHOR_PANEL_MARGINS;
        alignmentPanel.alignChildren = ["center", "top"];

        var anchorWidget = alignmentPanel.add("button", undefined, "");
        anchorWidget.preferredSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.minimumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.selectedAnchorIndex = NO_ANCHOR_INDEX;
        anchorWidget.onDraw = function () {
            drawAnchorWidget(this);
        };

        var optionPanel = addPanel(dialog, getLabel("panel.option"));
        optionPanel.alignChildren = ["left", "top"];

        var previewBoundsCheckbox = optionPanel.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.value = initialSettings.usePreviewBounds;

        var useGuidesCheckbox = optionPanel.add("checkbox", undefined, getLabel("checkbox.useGuides"));
        useGuidesCheckbox.value = initialSettings.useGuides;

        /* ボタンエリア：左にプレビュー、右にキャンセル・OK / Button row: preview on the left, Cancel/OK on the right */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.spacing = 0;

        var btnLeftGroup = btnRowGroup.add("group");
        setupRow(btnLeftGroup, "left");
        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = false;

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        setupRow(btnRightGroup, "right", 10);
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 現在のダイアログの状態を設定オブジェクトにまとめる
         * @returns {Object} 整列軸・境界・ガイドの設定（整列先が未選択なら alignmentAxes は null）
         */
        function getCurrentSettings() {
            var selectedIndex = anchorWidget.selectedAnchorIndex;
            var hasAnchor = (selectedIndex !== NO_ANCHOR_INDEX);
            return {
                /* ANCHOR_DEFINITIONS の要素はそのまま整列軸として使える / Each entry doubles as the axes pair */
                alignmentAxes: hasAnchor ? ANCHOR_DEFINITIONS[selectedIndex] : null,
                usePreviewBounds: previewBoundsCheckbox.value,
                /* 整列先を選んだときはアートボード基準に固定 / A chosen target always aligns to the artboard */
                useGuides: !hasAnchor && useGuidesCheckbox.value
            };
        }

        /**
         * プレビューを更新する（OFF・整列先未選択のときは基準位置へ戻す）
         * @returns {void}
         */
        function triggerPreview() {
            var currentSettings = getCurrentSettings();
            /* 整列先が未選択なら戻すだけ（ファイル名由来のフォールバックを抑止）
               With no target selected, only restore; the filename fallback must not kick in here */
            alignmentSession.preview((previewCheckbox.value && currentSettings.alignmentAxes) ? currentSettings : null);
        }

        /**
         * 整列先を選び直し、ウィジェットとガイドのチェックボックスを更新する
         * @param {number} selectedIndex - ANCHOR_DEFINITIONS 上のインデックス（未選択は NO_ANCHOR_INDEX）
         * @returns {void}
         */
        function selectAnchorAt(selectedIndex) {
            anchorWidget.selectedAnchorIndex = selectedIndex;
            redrawAnchorWidget(anchorWidget);
            /* 整列先を選ぶとガイドは使わないので、チェックボックスも無効にする
               A chosen target ignores guides, so the checkbox goes dim */
            useGuidesCheckbox.enabled = (selectedIndex === NO_ANCHOR_INDEX);
        }

        /**
         * ショートカットキーに対応する整列先を選ぶ
         * @param {string} pressedKey - 押されたキー名
         * @returns {boolean} 対応する整列先があれば true
         */
        function selectAnchorByShortcutKey(pressedKey) {
            for (var anchorIndex = 0; anchorIndex < ANCHOR_DEFINITIONS.length; anchorIndex++) {
                if (ANCHOR_DEFINITIONS[anchorIndex].shortcutKey !== pressedKey) continue;
                selectAnchorAt(anchorIndex);
                return true;
            }
            return false;
        }

        /* クリックした3×3のセルを整列先にする（座標はウィジェット基準）
           Set the target from the clicked 3x3 cell (coordinates are widget-relative) */
        anchorWidget.addEventListener("mousedown", function (event) {
            var column = clampGridIndex(Math.floor(event.clientX / (anchorWidget.size[0] / 3)));
            var row = clampGridIndex(Math.floor(event.clientY / (anchorWidget.size[1] / 3)));
            selectAnchorAt(row * 3 + column);
            triggerPreview();
        });

        previewBoundsCheckbox.onClick = triggerPreview;
        useGuidesCheckbox.onClick = triggerPreview;
        previewCheckbox.onClick = triggerPreview;

        /* 矢印キーの向きと、1段階の整列で使う整列先 / Arrow keys mapped to the target of one step */
        var STEP_SIDE_BY_KEY = { "Up": "top", "Down": "bottom", "Left": "left", "Right": "right" };

        dialog.addEventListener("keyup", function (keyEvent) {
            var stepSide = STEP_SIDE_BY_KEY[keyEvent.keyName];
            if (stepSide) {
                /* 矢印キー：押すたびに「スクリプトを1回実行」相当のステップ移動。
                   整列先の選択は解除する（OK 時に最終整列が二重適用されないようにするため）
                   Arrow keys step as if the script ran once; the target selection is cleared so
                   the final alignment on OK is not applied twice */
                selectAnchorAt(NO_ANCHOR_INDEX);
                alignmentSession.step(stepSide, getCurrentSettings());
                return;
            }
            if (selectAnchorByShortcutKey(keyEvent.keyName)) {
                triggerPreview();
                return;
            }
            if (keyEvent.keyName === "G" && useGuidesCheckbox.enabled) {
                useGuidesCheckbox.value = !useGuidesCheckbox.value;
                triggerPreview();
                return;
            }
            if (keyEvent.keyName === "B") {
                previewBoundsCheckbox.value = !previewBoundsCheckbox.value;
                triggerPreview();
            }
        });

        var dialogShowResult = dialog.show();

        /* OK / キャンセルどちらでも、閉じる際はいったん基準位置へ戻す（最終整列は main 側で改めて適用）
           On both OK and Cancel the items go back to the baseline; main re-applies the final alignment */
        alignmentSession.preview(null);
        return (dialogShowResult === 1) ? getCurrentSettings() : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトをアートボードまたはガイドへ整列する
     * @returns {void}
     */
    function main() {
        try {
            if (GUIDE_SEARCH_MODE !== "inside" && GUIDE_SEARCH_MODE !== "nearest") {
                alert(labelText("alert.invalidGuideSearchMode") + GUIDE_SEARCH_MODE);
                return;
            }

            var documentRef = app.activeDocument;
            var selectedItems = documentRef.selection;
            if (selectedItems.length === 0) {
                alert(getLabel("alert.noSelection"));
                return;
            }

            var artboards = documentRef.artboards;
            var artboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;

            var settings = {
                /* ダイアログを出さないときはファイル名から整列先を決める / Without the dialog the filename picks the target */
                alignmentAxes: AXES_BY_ALIGNMENT_SIDE[detectAlignmentSideFromFileName()],
                /* プレビュー境界使用の初期値は環境設定から取得 / Seed the preview-bounds option from the preferences */
                usePreviewBounds: app.preferences.getBooleanPreference("includeStrokeInBounds"),
                useGuides: USE_GUIDES
            };

            if (SHOW_DIALOG) {
                var alignmentSession = createAlignmentSession(selectedItems, documentRef, artboardRect);
                settings = showAlignmentDialog(settings, alignmentSession);
                if (settings === null) {
                    /* キャンセル：矢印キーでのステップ移動も含めて完全復元 / Cancel restores the arrow-key steps too */
                    alignmentSession.restoreOriginal();
                    return;
                }
                /* 整列先が未選択なら、矢印キーでの最終位置をそのまま確定する
                   With no target selected, the arrow-key result stands as-is */
                if (!settings.alignmentAxes) return;
            }

            var alignContext = createAlignContext(documentRef, artboardRect, settings.useGuides);
            applyAlignment(selectedItems, settings.alignmentAxes, alignContext, settings.usePreviewBounds);
        } catch (error) {
            alert(labelText("alert.generalError") + error.message);
        }
    }

    main();

})();
