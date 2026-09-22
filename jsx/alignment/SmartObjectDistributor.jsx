#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、指定した行数・列数のグリッドに沿って各セルの中央へ配置します。
配置先は「現在のアートボード」「最背面のオブジェクト」「_target レイヤーの長方形」から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectDistributor.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na3c45cea09b7

### Overview

Places the selected objects at the center of each cell of a grid with the rows and columns you specify.
The target area can be the current artboard, the backmost object, or a rectangle on the `_target` layer.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectDistributor.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectDistributor";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.9.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectDistributor.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectDistributor.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na3c45cea09b7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var CONFIG = {
        defaultGutter: 10,           /* 間隔の初期値（定規単位） / default gutter */
        defaultMargin: 10,           /* マージンの初期値（定規単位） / default margin */
        blackCellOpacity: 15,        /* 黒セルの不透明度（%） / opacity for black cells */
        whiteCellOpacity: 100,       /* 白セルの不透明度（%） / opacity for white cells */
        fallbackDivision: 5,         /* 分割数を決められないときの既定値 / fallback division */
        maxDivision: 100,            /* 行・列の上限 / max rows or columns */
        maxCellCount: 1000,          /* セル総数の上限 / max total cells */
        parkingStep: 200,            /* 退避時の移動量（pt） / step when parking objects */
        overlapTolerance: 1,         /* 重なり判定の余裕（pt） / overlap tolerance */
        artboardBuffer: 10,          /* 他アートボードとの余裕（pt） / buffer around artboards */
        targetLayerName: "_target",                  /* 配置先レイヤー名 / target layer name */
        cellLayerName: "cell-background",            /* セル描画レイヤー名 / cell layer name */
        previewCellLayerName: "_Preview_Background", /* プレビュー用レイヤー名 / preview layer name */
        legacyPreviewLayerName: "_Preview_Guides"    /* 旧版のプレビュー用レイヤー名 / legacy preview layer */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /**
     * ウィンドウに共通のレイアウト設定を適用します。/ Apply shared window layout.
     *
     * @param {Window} targetWindow - 対象のウィンドウ。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウト設定を適用します。/ Apply shared panel layout.
     *
     * @param {Panel} targetPanel - 対象のパネル。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
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
     * 横並びの行グループを設定します（ボタン列など）。/ Apply a horizontal row group.
     *
     * @param {Group} targetGroup - 対象のグループ。
     * @param {string} [alignment] - 親に対する揃え。省略時は "left"。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = alignment || "left";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定 px 詰めます（レイアウト確定後に呼びます）。/ Trim a button's height.
     *
     * @param {Button} targetButton - 対象のボタン。
     * @param {number} trimPx - 詰める高さ（px）。
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPx) {
        try {
            targetButton.size = [targetButton.size.width, targetButton.size.height - trimPx];
        } catch (e) { /* レイアウト前は size が無いことがある / size may be missing before layout */ }
    }

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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
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

    /**
     * 実行環境のロケールから表示言語を決めます。/ Pick the UI language from the locale.
     *
     * @returns {string} "ja" または "en"。
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "グリッドに整列配置", en: "Arrange in Grid" }
        },
        panel: {
            placement: { ja: "配置先", en: "Placement Area" },
            division: { ja: "分割とマージン", en: "Divisions & Margin" },
            cellDrawing: { ja: "セルの扱い", en: "Cell Handling" }
        },
        fieldLabel: {
            rowCount: { ja: "行数", en: "Rows" },
            columnCount: { ja: "列数", en: "Columns" },
            gutter: { ja: "セル間隔", en: "Gutter" },
            margin: { ja: "マージン", en: "Margin" },
            cellColor: { ja: "カラー", en: "Color" },
            cellOpacity: { ja: "不透明度", en: "Opacity" }
        },
        radio: {
            targetArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            targetBackmost: { ja: "最背面のオブジェクト", en: "Backmost Object" },
            targetRectLayer: { ja: "「_target」レイヤーの長方形", en: "Rectangle in '_target' Layer" },
            keepCell: { ja: "長方形で残す", en: "Keep as Rectangle" },
            toGuide: { ja: "ガイド化", en: "Convert to Guides" },
            toArtboard: { ja: "アートボード化", en: "Convert to Artboards" },
            blackCell: { ja: "黒", en: "Black" },
            whiteCell: { ja: "白", en: "White" },
            transparentCell: { ja: "塗りなし", en: "No Fill" }
        },
        button: {
            transparencyGrid: { ja: "透明グリッド表示", en: "Transparency Grid" },
            randomize: { ja: "シャッフル", en: "Shuffle" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            targetUnavailable: { ja: "該当するオブジェクトがありません。", en: "No matching object was found." },
            targetBackmost: {
                ja: "表示されていてロックされていないオブジェクトのうち、最も背面にあるものを配置先にします。そのオブジェクト自体は配置しません。",
                en: "Uses the backmost visible, unlocked object as the area. That object itself is not placed."
            },
            targetRectLayer: {
                ja: "この長方形は処理中だけ非表示になり、配置対象には含まれません。",
                en: "This rectangle is hidden while the script runs and is never placed into a cell."
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
                ja: "セル数より多いオブジェクトは、アートボードの外へ退避します。",
                en: "Objects beyond the number of cells are parked outside the artboard."
            },
            gutter: { ja: "行数・列数のいずれかが2以上のときに有効です。", en: "Available when the rows or columns are 2 or more." },
            margin: { ja: "配置先の四辺から内側に取る余白です。", en: "Inset taken from each edge of the placement area." },
            randomize: {
                ja: "セルへの割り当て順をシャッフルします（押すたびに変わります）。",
                en: "Shuffles the order in which objects fill the cells (changes on every click)."
            },
            transparencyGrid: {
                ja: "透明グリッドの表示を切り替えます。スクリプト終了時に元へ戻します。",
                en: "Toggles the transparency grid. It is restored when the script finishes."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
            artboardError: { ja: "アートボードの作成中にエラーが発生しました。", en: "Error occurred while creating artboards." },
            artboardCreated: { ja: " 個のアートボードを作成しました。", en: " artboards created." }
        }
    };

    /**
     * ラベルを現在の言語で取得します。/ Get a label in the current language.
     *
     * @param {LabelEntry} labelEntry - 日英の文言を持つラベル。
     * @returns {string} 表示言語の文言。
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang];
    }

    /**
     * コロン付きラベルを返します（日本語は全角、英語は半角）。/ Label with colon (full-width JA, half-width EN).
     *
     * @param {LabelEntry} labelEntry - 日英の文言を持つラベル。
     * @returns {string} コロンを付けた文言。
     */
    function labelText(labelEntry) {
        return getLabel(labelEntry) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 位置の記録と復元 / Position bookkeeping
    // =========================================
    /* プレビューの巻き戻しは app.undo() を主、ここでの座標復元を保険とする。
       app.undo() が効けば履歴は伸びないが、1プレビューが複数の履歴ステップに
       分かれると 1 回では戻りきらない。その差分だけをここで埋める。
       位置が変わっていなければ translate しないので、undo が完全に効いた場合は
       履歴を 1 ステップも増やさない。
       Preview rollback relies on app.undo(); restoring centers here only fills the gap when one
       preview spans several history steps, and never moves items that are already in place. */

    /** 位置が変化したと見なす最小差分（pt） / Minimum delta treated as a real move. */
    var RESTORE_TOLERANCE = 0.001;

    /** @type {Array<number[]>} 記録した中心座標 [x, y] の配列。 */
    var originalCenters = [];

    /**
     * 各オブジェクトの中心座標を記録します。/ Record the center of each object.
     *
     * @param {Array<PageItem>} pageItems - 記録対象のオブジェクト。
     * @returns {void}
     */
    function saveOriginalCenters(pageItems) {
        originalCenters = [];
        for (var i = 0; i < pageItems.length; i++) {
            var bounds = pageItems[i].visibleBounds;
            originalCenters.push([(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2]);
        }
    }

    /**
     * 記録した中心座標へオブジェクトを戻します。/ Move objects back to the recorded centers.
     * 既に元の位置にあるものは動かさないため、履歴を無駄に増やしません。
     *
     * @param {Array<PageItem>} pageItems - 復元対象のオブジェクト（記録時と同じ順序）。
     * @returns {void}
     */
    function restoreOriginalCenters(pageItems) {
        for (var i = 0; i < pageItems.length && i < originalCenters.length; i++) {
            try {
                var bounds = pageItems[i].visibleBounds;
                var dx = originalCenters[i][0] - (bounds[0] + bounds[2]) / 2;
                var dy = originalCenters[i][1] - (bounds[1] + bounds[3]) / 2;

                /* undo が効いていれば差分は 0。ここで translate すると履歴が伸びる / Zero after a full undo; moving would add history */
                if (Math.abs(dx) < RESTORE_TOLERANCE && Math.abs(dy) < RESTORE_TOLERANCE) continue;
                pageItems[i].translate(dx, dy);
            } catch (e) {
                /* 参照が失われたオブジェクトは飛ばす / Skip items whose reference went stale */
            }
        }
    }

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
     * 「_target」レイヤー内の最初の長方形を返します。/ Return the first rectangle in the "_target" layer.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {PathItem|null} 見つかった長方形。無ければ null。
     */
    function findTargetLayerRectangle(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            var candidateLayer = doc.layers[i];
            if (candidateLayer.name !== CONFIG.targetLayerName) continue;
            return (candidateLayer.pathItems.length > 0) ? candidateLayer.pathItems[0] : null;
        }
        return null;
    }

    /**
     * システムレイヤーを除く最背面のオブジェクトを返します。/ Return the backmost object, ignoring system layers.
     * 選択中のアイテムも候補に含めます（枠として使う場合の利便性を優先）。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {PageItem|null} 最背面のオブジェクト。無ければ null。
     */
    function findBackmostPageItem(doc) {
        var systemLayerNames = {
            "_target": 1,
            "_Preview_Guides": 1,
            "_Preview_Background": 1,
            "cell-background": 1,
            "placement_layer": 1
        };
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var candidateLayer = doc.layers[i];
            if (!candidateLayer.visible || candidateLayer.locked || systemLayerNames[candidateLayer.name]) continue;
            for (var j = candidateLayer.pageItems.length - 1; j >= 0; j--) {
                var pageItem = candidateLayer.pageItems[j];
                if (pageItem.hidden || pageItem.locked) continue;
                return pageItem;
            }
        }
        return null;
    }

    /**
     * 選択から、指定のアイテムを除いた配列を返します（順序は保持）。/ Copy the selection without one item.
     *
     * @param {Array<PageItem>|null} selectionItems - 選択。
     * @param {PageItem|null} excludedItem - 除くアイテム。
     * @returns {Array<PageItem>} 残ったオブジェクト。
     */
    function excludeItem(selectionItems, excludedItem) {
        var remainingItems = [];
        for (var i = 0; selectionItems && i < selectionItems.length; i++) {
            if (selectionItems[i] !== excludedItem) remainingItems.push(selectionItems[i]);
        }
        return remainingItems;
    }

    /**
     * セルが正方形に近くなる行数・列数を求めます。/ Find rows and columns that make cells nearly square.
     * 余白セルが行／列1本分以上になる構成は除外し、その範囲でセル比1に最も近いものを選びます。
     *
     * @param {number} itemCount - 配置するオブジェクトの数。
     * @param {number[]} targetRect - 配置先の矩形 [左, 上, 右, 下]。
     * @returns {{rowCount: number, columnCount: number}} 行数と列数。
     */
    function computeDefaultDivision(itemCount, targetRect) {
        var fallbackDivision = { rowCount: CONFIG.fallbackDivision, columnCount: CONFIG.fallbackDivision };
        if (!itemCount || itemCount <= 0) return fallbackDivision;
        if (itemCount === 1) return { rowCount: 1, columnCount: 1 };

        var areaWidth = targetRect[2] - targetRect[0];
        var areaHeight = targetRect[1] - targetRect[3];
        if (areaWidth <= 0 || areaHeight <= 0) return fallbackDivision;

        var bestDivision = { rowCount: 1, columnCount: itemCount };
        var bestAspectRatio = Infinity;
        for (var columnCount = 1; columnCount <= itemCount; columnCount++) {
            var rowCount = Math.ceil(itemCount / columnCount);
            var emptyCellCount = rowCount * columnCount - itemCount;
            if (emptyCellCount > 0 && emptyCellCount >= Math.min(rowCount, columnCount)) continue;

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
     * 指定名のレイヤーがあれば削除します。/ Remove the layer with the given name if present.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - 削除するレイヤー名。
     * @returns {void}
     */
    function removeLayerByName(doc, layerName) {
        try {
            var foundLayer = doc.layers.getByName(layerName);
            foundLayer.locked = false;
            foundLayer.remove();
        } catch (e) {
            /* 見つからない場合は何もしない / Nothing to do when the layer is missing */
        }
    }

    /**
     * レイヤーを取得し、無ければ作成します。/ Get the layer, creating it when missing.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - レイヤー名。
     * @returns {Layer} 取得または作成したレイヤー（ロック解除済み）。
     */
    function getOrCreateLayer(doc, layerName) {
        var foundLayer;
        try {
            foundLayer = doc.layers.getByName(layerName);
        } catch (e) {
            /* getByName は見つからないと例外 / getByName throws when the layer is missing */
            foundLayer = doc.layers.add();
            foundLayer.name = layerName;
        }
        foundLayer.locked = false;
        return foundLayer;
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
        } catch (e) { }
    }

    /**
     * セル描画レイヤーの長方形を選択します。/ Select the rectangles on the cell layer.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {boolean} 選択できた場合は true。レイヤーが無い場合は false。
     */
    function selectCellRectangles(doc) {
        var cellItems = [];
        try {
            var cellLayer = doc.layers.getByName(CONFIG.cellLayerName);
            for (var i = 0; i < cellLayer.pathItems.length; i++) cellItems.push(cellLayer.pathItems[i]);
        } catch (e) {
            return false;
        }
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
     * @property {number} columnCount - 列数。
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
     * @returns {GridMetrics|null} グリッドの寸法。入力が無効な場合は null。
     */
    function computeGridMetrics(rowCount, columnCount, margin, gutter, targetRect) {
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
            var cellTop = gridMetrics.originTop - (gridMetrics.cellHeight + gridMetrics.gutter) * i;
            for (var j = 0; j < gridMetrics.columnCount; j++) {
                var cellLeft = gridMetrics.originLeft + (gridMetrics.cellWidth + gridMetrics.gutter) * j;
                var cellRect = [cellLeft, cellTop, cellLeft + gridMetrics.cellWidth, cellTop - gridMetrics.cellHeight];
                if (handleCell(cellRect) === false) return;
            }
        }
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
        forEachCell(gridMetrics, function (cellRect) {
            var cellRectangle = cellLayer.pathItems.rectangle(
                cellRect[1], cellRect[0], gridMetrics.cellWidth, gridMetrics.cellHeight);
            cellRectangle.stroked = false;
            cellRectangle.filled = (fillColor !== null);
            if (fillColor) cellRectangle.fillColor = fillColor;
            if (opacity !== null) cellRectangle.opacity = opacity;
            if (asGuide) cellRectangle.guides = true;
        });
        cellLayer.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 各セルの中央へオブジェクトを1つずつ配置します。/ Place one object at the center of each cell.
     *
     * @param {Array<PageItem>} pageItems - 配置するオブジェクト（配置順）。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @returns {void}
     */
    function placeItemsInCells(pageItems, gridMetrics) {
        var itemIndex = 0;
        forEachCell(gridMetrics, function (cellRect) {
            if (itemIndex >= pageItems.length) return false;

            var bounds = pageItems[itemIndex].visibleBounds;
            var dx = (cellRect[0] + cellRect[2]) / 2 - (bounds[0] + bounds[2]) / 2;
            var dy = (cellRect[1] + cellRect[3]) / 2 - (bounds[1] + bounds[3]) / 2;
            pageItems[itemIndex].translate(dx, dy);
            itemIndex++;
        });
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

    // =========================================
    // 配置順 / Placement order
    // =========================================

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
     * 順序をシャッフルした新しい配列を返します（Fisher-Yates）。/ Return a shuffled copy.
     *
     * @param {Array<PageItem>} sourceItems - 元の配列。
     * @returns {Array<PageItem>} シャッフルした配列。
     */
    function createShuffledCopy(sourceItems) {
        var shuffledItems = [];
        for (var i = 0; i < sourceItems.length; i++) shuffledItems.push(sourceItems[i]);
        for (var k = shuffledItems.length - 1; k > 0; k--) {
            var swapIndex = Math.floor(Math.random() * (k + 1));
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
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @param {boolean} hasBackmostItem - 最背面のオブジェクトが存在する場合は true。
     * @param {boolean} hasTargetRect - 「_target」レイヤーの長方形が存在する場合は true。
     * @returns {{artboardRadio: RadioButton, backmostRadio: RadioButton, rectLayerRadio: RadioButton}} 配置先のラジオボタン。
     */
    function buildPlacementTargetPanel(parentWindow, hasBackmostItem, hasTargetRect) {
        var placementPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.placement));
        setupPanel(placementPanel, 6);
        placementPanel.alignChildren = ["left", "top"];

        var placementControls = {
            artboardRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetArtboard)),
            backmostRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetBackmost)),
            rectLayerRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetRectLayer))
        };

        placementControls.backmostRadio.enabled = hasBackmostItem;
        placementControls.rectLayerRadio.enabled = hasTargetRect;

        /* 選べない理由、または選んだときの挙動をツールチップで補う / Tooltips explain why an option is unavailable or what it does */
        placementControls.backmostRadio.helpTip = hasBackmostItem
            ? getLabel(LABELS.tooltip.targetBackmost)
            : getLabel(LABELS.tooltip.targetUnavailable);
        placementControls.rectLayerRadio.helpTip = hasTargetRect
            ? getLabel(LABELS.tooltip.targetRectLayer)
            : getLabel(LABELS.tooltip.targetUnavailable);

        if (hasTargetRect) {
            placementControls.rectLayerRadio.value = true;
        } else {
            placementControls.artboardRadio.value = true;
        }
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
     * @returns {EditText} 作成した入力欄。
     */
    function addFieldRow(parentColumn, labelEntry, initialValue, labelWidth, unitSuffix, tooltipEntry) {
        var fieldRow = parentColumn.add("group");
        setupRow(fieldRow, "left", 4);
        fieldRow.alignChildren = "center";

        var fieldLabel = fieldRow.add("statictext", undefined, labelText(labelEntry));
        fieldLabel.preferredSize.width = labelWidth;
        if (unitSuffix) fieldLabel.justify = "right";

        var fieldInput = fieldRow.add("edittext", undefined, String(initialValue));
        fieldInput.characters = unitSuffix ? 4 : 3;
        if (unitSuffix) fieldRow.add("statictext", undefined, unitSuffix);

        if (tooltipEntry) {
            fieldLabel.helpTip = getLabel(tooltipEntry);
            fieldInput.helpTip = getLabel(tooltipEntry);
        }
        return fieldInput;
    }

    /**
     * 分割とマージンのパネルを作成します（左：分割数／右：間隔）。/ Build the division panel.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @param {{rowCount: number, columnCount: number}} defaultDivision - 行数・列数の初期値。
     * @param {string} unitLabel - 定規単位のラベル。
     * @returns {{rowCountInput: EditText, columnCountInput: EditText, gutterInput: EditText, marginInput: EditText}} 入力欄。
     */
    function buildDivisionPanel(parentWindow, defaultDivision, unitLabel) {
        var divisionPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.division));
        setupPanel(divisionPanel, 6);
        divisionPanel.orientation = "row";
        divisionPanel.alignChildren = ["left", "top"];
        divisionPanel.spacing = COLUMN_SPACING;

        var divisionColumn = divisionPanel.add("group");
        divisionColumn.orientation = "column";
        divisionColumn.alignChildren = "left";
        divisionColumn.margins.right = COLUMN_SPACING;

        var spacingColumn = divisionPanel.add("group");
        spacingColumn.orientation = "column";
        spacingColumn.alignChildren = "left";

        /* ラベル幅は日本語環境と英語環境で個別に指定 / Label widths per language */
        var divisionLabelWidth = (uiLang === "ja") ? 45 : 60;
        var spacingLabelWidth = (uiLang === "ja") ? 70 : 75;

        return {
            rowCountInput: addFieldRow(divisionColumn, LABELS.fieldLabel.rowCount, defaultDivision.rowCount, divisionLabelWidth, null, LABELS.tooltip.division),
            columnCountInput: addFieldRow(divisionColumn, LABELS.fieldLabel.columnCount, defaultDivision.columnCount, divisionLabelWidth, null, LABELS.tooltip.division),
            gutterInput: addFieldRow(spacingColumn, LABELS.fieldLabel.gutter, CONFIG.defaultGutter, spacingLabelWidth, unitLabel, LABELS.tooltip.gutter),
            marginInput: addFieldRow(spacingColumn, LABELS.fieldLabel.margin, CONFIG.defaultMargin, spacingLabelWidth, unitLabel, LABELS.tooltip.margin)
        };
    }

    /**
     * セル描画パネルを作成します。/ Build the cell drawing panel.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @returns {Object} セル描画モード、カラー、不透明度、透明グリッドボタンをまとめたオブジェクト。
     */
    function buildCellDrawingPanel(parentWindow) {
        var cellDrawingPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.cellDrawing));
        setupPanel(cellDrawingPanel, 6);

        var cellModeRow = cellDrawingPanel.add("group");
        setupRow(cellModeRow, "left");

        var cellColorRow = cellDrawingPanel.add("group");
        setupRow(cellColorRow, "left");
        cellColorRow.add("statictext", undefined, labelText(LABELS.fieldLabel.cellColor));

        var cellOpacityRow = cellDrawingPanel.add("group");
        setupRow(cellOpacityRow, "left");
        cellOpacityRow.alignChildren = "center";
        cellOpacityRow.add("statictext", undefined, labelText(LABELS.fieldLabel.cellOpacity));

        var cellControls = {
            keepCellRadio: cellModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.keepCell)),
            toGuideRadio: cellModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.toGuide)),
            toArtboardRadio: cellModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.toArtboard)),
            blackCellRadio: cellColorRow.add("radiobutton", undefined, getLabel(LABELS.radio.blackCell)),
            whiteCellRadio: cellColorRow.add("radiobutton", undefined, getLabel(LABELS.radio.whiteCell)),
            transparentCellRadio: cellColorRow.add("radiobutton", undefined, getLabel(LABELS.radio.transparentCell)),
            opacityInput: cellOpacityRow.add("edittext", undefined, String(CONFIG.blackCellOpacity))
        };
        cellControls.keepCellRadio.value = true;
        cellControls.blackCellRadio.value = true;
        cellControls.opacityInput.characters = 4;

        /* モードごとの結果をツールチップで補う / Tooltips describe each mode's result */
        cellControls.keepCellRadio.helpTip = getLabel(LABELS.tooltip.keepCell);
        cellControls.toGuideRadio.helpTip = getLabel(LABELS.tooltip.toGuide);
        cellControls.toArtboardRadio.helpTip = getLabel(LABELS.tooltip.toArtboard);

        cellOpacityRow.add("statictext", undefined, "%");

        /* ボタンはパネル幅いっぱいに広げない / Keep the button from stretching across the panel */
        cellControls.transparencyGridButton = cellOpacityRow.add("button", undefined, getLabel(LABELS.button.transparencyGrid));
        cellControls.transparencyGridButton.alignment = "left";
        cellControls.transparencyGridButton.helpTip = getLabel(LABELS.tooltip.transparencyGrid);
        return cellControls;
    }

    /**
     * 下部のボタン列を作成します。/ Build the footer button row.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @returns {{btnShuffle: Button, btnCancel: Button, btnOK: Button}} 各ボタン。
     */
    function buildFooterRow(parentWindow) {
        var btnRowGroup = parentWindow.add("group");
        setupRow(btnRowGroup, "fill", 10);

        var footerButtons = {
            btnShuffle: btnRowGroup.add("button", undefined, getLabel(LABELS.button.randomize))
        };
        footerButtons.btnShuffle.alignment = "left";
        footerButtons.btnShuffle.helpTip = getLabel(LABELS.tooltip.randomize);

        /* 左右のボタンを引き離すためのスペーサー / Spacer that pushes the buttons apart */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = (uiLang === "ja") ? 40 : 60;
        spacer.maximumSize.height = 0;

        footerButtons.btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        footerButtons.btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        return footerButtons;
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
     * 入力欄の値を pt 換算の長さとして読み取ります。/ Read a length in points from a field.
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @param {number} pointsPerUnit - 定規単位 1 あたりの pt。
     * @returns {number} pt に換算した長さ。数値でない場合は 0。
     */
    function readLengthPt(inputField, pointsPerUnit) {
        var parsedLength = parseFloat(inputField.text);
        return isFinite(parsedLength) ? parsedLength * pointsPerUnit : 0;
    }

    /**
     * 上下キーで数値を増減します（Shift:10単位、Option:0.1単位）。/ Step a value with arrow keys.
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @param {boolean} allowsDecimal - 小数の増減を許可する場合は true。
     * @param {number} minValue - 下限値。
     * @param {function(): void} onStep - 値を変えたあとに呼ぶ処理。
     * @returns {void}
     */
    function enableArrowKeyStep(inputField, allowsDecimal, minValue, onStep) {
        inputField.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var fieldValue = Number(inputField.text);
            if (isNaN(fieldValue)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var direction = (event.keyName === "Up") ? 1 : -1;

            if (keyboard.shiftKey) {
                /* 10 の倍数へ丸めながら増減 / Step to the next multiple of 10 */
                fieldValue = (direction > 0) ? Math.ceil((fieldValue + 1) / 10) * 10 : Math.floor((fieldValue - 1) / 10) * 10;
            } else if (keyboard.altKey && allowsDecimal) {
                fieldValue = Math.round((fieldValue + direction * 0.1) * 10) / 10;
            } else {
                fieldValue = Math.round(fieldValue) + direction;
            }

            inputField.text = String(Math.max(minValue, fieldValue));
            event.preventDefault();

            onStep();
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示して配置処理を実行します。/ Show the dialog and run the distribution.
     *
     * @returns {void}
     */
    function showDistributeDialog() {
        var doc = app.activeDocument;

        /* 前回実行時のセル描画レイヤーが残っていれば削除 / Remove the cell layer left by a previous run */
        removeLayerByName(doc, CONFIG.cellLayerName);

        /* 以後の計算は「ダイアログ起動時点のアクティブアートボード」を基準に固定 / Fix the base to the artboard active at launch */
        var baseArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var baseArtboardRect = doc.artboards[baseArtboardIndex].artboardRect;

        /* 対象候補の検出（この時点では変更を加えない）/ Detect target candidates without changing anything */
        var targetRectItem = findTargetLayerRectangle(doc);
        var backmostItem = findBackmostPageItem(doc);

        /* 選択内容はダイアログ起動時に固定する（順序も保持）。枠として使う「_target」矩形は移動対象から除く
           Freeze the selection (and its order) at launch, leaving out the _target frame rectangle */
        var placeableItems = excludeItem(doc.selection, targetRectItem);
        saveOriginalCenters(placeableItems);

        /* 初期ターゲット矩形（_target 矩形があればそれ、なければ現在のアートボード）/ Initial target rect */
        var initialTargetRect = targetRectItem ? targetRectItem.geometricBounds : baseArtboardRect;

        /* 「_target」レイヤーの矩形は対象として選ばれている間は非表示にする / Hide the _target rectangle while it is the target */
        if (targetRectItem) setItemHidden(targetRectItem, true);

        var rulerUnit = getUnitInfo();
        var defaultDivision = computeDefaultDivision(placeableItems.length, initialTargetRect);
        var transparencyGridToggleCount = 0;

        var distributeDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(distributeDialog);

        var placementUI = buildPlacementTargetPanel(distributeDialog, backmostItem !== null, targetRectItem !== null);
        var divisionUI = buildDivisionPanel(distributeDialog, defaultDivision, rulerUnit.label);
        var cellUI = buildCellDrawingPanel(distributeDialog);
        var footerUI = buildFooterRow(distributeDialog);

        /** @type {Array<PageItem>|null} シャッフルで決めた配置順。未使用なら null。 */
        var randomizedOrder = null;
        /** @type {RadioButton} 「長方形で残す」に戻したときに復帰させるカラー選択。 */
        var lastCellColorRadio = cellUI.blackCellRadio;
        /** @type {boolean} app.undo() で剥がすべきプレビューが適用済みか。 */
        var hasUncommittedPreview = false;
        /** @type {boolean} OK／キャンセルで後始末済みか（onClose の二重実行を防ぐ）。 */
        var isCleanedUp = false;

        bindDialogEvents();

        /**
         * ↑↓キーで値を変えたあとに間隔欄とプレビューを更新します。/ Refresh after an arrow-key step.
         *
         * @returns {void}
         */
        function handleArrowKeyStep() {
            syncGutterEnabled();
            updatePreview();
        }

        /* 行数・列数は 1 未満にしない（0 ではグリッドが成立しない）/ Rows and columns never go below 1 */
        enableArrowKeyStep(divisionUI.rowCountInput, false, 1, handleArrowKeyStep);
        enableArrowKeyStep(divisionUI.columnCountInput, false, 1, handleArrowKeyStep);
        enableArrowKeyStep(divisionUI.gutterInput, true, 0, handleArrowKeyStep);
        enableArrowKeyStep(divisionUI.marginInput, true, 0, handleArrowKeyStep);
        enableArrowKeyStep(cellUI.opacityInput, true, 0, handleArrowKeyStep);

        /* 初期状態を反映してプレビューを表示 / Apply the initial state and show the preview */
        syncGutterEnabled();
        syncCellModeUI();
        updatePreview();

        distributeDialog.layout.layout(true);
        trimButtonHeight(cellUI.transparencyGridButton, 4);
        distributeDialog.show();

        /* 透明グリッドの表示状態を元へ戻す / Restore the transparency grid */
        if (transparencyGridToggleCount % 2 !== 0) {
            app.executeMenuCommand('TransparencyGrid Menu Item');
        }

        // -----------------------------------------
        // 入力値の取得 / Input readers
        // -----------------------------------------

        /**
         * 不透明度を 0〜100 に収めて読み取ります。/ Read the opacity clamped to 0-100.
         *
         * @returns {number|null} 不透明度（%）。数値でない場合は null。
         */
        function readOpacity() {
            var parsedOpacity = parseFloat(cellUI.opacityInput.text);
            if (!isFinite(parsedOpacity)) return null;
            return Math.max(0, Math.min(100, parsedOpacity));
        }

        /**
         * 現在のセル描画モードを返します。/ Return the current cell drawing mode.
         *
         * @returns {string} "keep"、"guide"、"artboard" のいずれか。
         */
        function getCellMode() {
            if (cellUI.toArtboardRadio.value) return "artboard";
            if (cellUI.toGuideRadio.value) return "guide";
            return "keep";
        }

        /**
         * 選択中の対象領域の矩形を返します。/ Return the rect of the selected target area.
         *
         * @returns {number[]} 配置先の矩形 [左, 上, 右, 下]。
         */
        function getTargetRect() {
            if (placementUI.backmostRadio.value && backmostItem) return backmostItem.geometricBounds;
            if (placementUI.rectLayerRadio.value && targetRectItem) return targetRectItem.geometricBounds;
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
                readLengthPt(divisionUI.marginInput, rulerUnit.pointsPerUnit),
                readLengthPt(divisionUI.gutterInput, rulerUnit.pointsPerUnit),
                getTargetRect()
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
            return excludeItem(placeableItems, backmostItem);
        }

        // -----------------------------------------
        // 描画と配置 / Drawing and placement
        // -----------------------------------------

        /**
         * セル描画と配置を実行します（プレビュー／本番共通）。/ Draw cells and place objects.
         *
         * @param {boolean} isPreview - プレビューとして描画する場合は true。
         * @param {function(): void} [onWillMutate] - ドキュメントを変更する直前に一度だけ呼ばれます。
         * @returns {boolean} ドキュメントを変更した場合は true。入力が無効で何もしなかった場合は false。
         */
        function renderDistribution(isPreview, onWillMutate) {
            /* 配置先の座標を読む前に、必ず元の位置へ戻しておく（通常は clearPreview() 済みで何も動かない）
               Restore positions before reading the target; normally a no-op after clearPreview() */
            restoreOriginalCenters(placeableItems);

            var cellMode = getCellMode();
            var distributionItems = orderByRandomizedOrder(getDistributionItems(), randomizedOrder);
            var gridMetrics = readGridMetrics();

            /* グリッドが成立しない入力では、ドキュメントに一切手を加えない / Leave the document untouched without a valid grid */
            if (!gridMetrics) return false;

            /* これ以降はドキュメントを変更する / Everything below mutates the document */
            if (onWillMutate) onWillMutate();

            /* セル数を超えたオブジェクトだけを対象領域の外へ退避する / Park only the objects beyond the cell count */
            var overflowItems = distributionItems.slice(gridMetrics.rowCount * gridMetrics.columnCount);
            if (overflowItems.length > 0) {
                parkItemsOutside(doc, overflowItems, gridMetrics.targetRect);
            }

            /* アートボード化のプレビューは、セルの範囲をガイドで示す / The artboard mode previews its cells as guides */
            if ((cellMode !== "artboard") || isPreview) {
                var layerName = isPreview ? CONFIG.previewCellLayerName : CONFIG.cellLayerName;
                var asGuide = (cellMode === "guide") || (cellMode === "artboard" && isPreview);
                drawCells(getOrCreateLayer(doc, layerName), gridMetrics, asGuide, getCellFillColor(), readOpacity());
            }

            placeItemsInCells(distributionItems, gridMetrics);
            return true;
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------
        /* app.undo() で直前のプレビューをヒストリごと取り除いてから描き直す。
           これをしないと、入力欄を 1 文字打つたびに数十〜数千ステップが積まれ、
           Illustrator の取り消し回数の上限を超えてユーザーの実行前履歴が失われる。

           ただし app.undo() は 1 回で 1 ステップしか戻さない。1 プレビューが
           複数ステップに分かれる場合は戻りきらないため、レイヤー削除と座標復元を
           保険として必ず併走させる（どちらも差分が無ければ何もしない）。

           undo の回数は「自分が積んだ 1 回分」に限る。ダイアログ表示前に
           cell-background レイヤーの削除と _target 矩形の非表示という 2 つの
           変更を済ませており、そこまで巻き戻すと状態が壊れるため。
           Each preview is peeled off with one app.undo() so typing does not flood the history;
           layer removal and center restoring back it up, and undo never reaches the two changes
           made before the dialog opened. */

        /**
         * プレビューを消して元の状態に戻します。/ Discard the preview and restore the original state.
         *
         * @returns {void}
         */
        function clearPreview() {
            /* 1. 直前のプレビューをヒストリから取り除く / Peel the last preview off the history */
            if (hasUncommittedPreview) {
                hasUncommittedPreview = false;
                try {
                    app.undo();
                } catch (e) {
                    /* undo できない状態なら、以下の後始末に任せる / Fall back to the cleanup below */
                }
            }

            /* 2. undo で戻りきらなかった分だけを片付ける / Clean up whatever undo left behind */
            removeLayerByName(doc, CONFIG.previewCellLayerName);
            removeLayerByName(doc, CONFIG.legacyPreviewLayerName);
            restoreOriginalCenters(placeableItems);
        }

        /**
         * 現在の設定でプレビューを描き直します。/ Redraw the preview with the current settings.
         *
         * @returns {void}
         */
        function updatePreview() {
            try {
                clearPreview();
                renderDistribution(true, function () {
                    /* 変更が始まった時点で印を付ける。途中で例外が出ても undo で剥がせる / Mark first so undo can peel it even on failure */
                    hasUncommittedPreview = true;
                });
            } catch (e) {
                clearPreview();
            }
            app.redraw();
        }

        // -----------------------------------------
        // UI 状態の同期 / UI state
        // -----------------------------------------

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
            divisionUI.gutterInput.enabled = (rowCount > 1 || columnCount > 1);
        }

        /**
         * カラー選択を不透明度欄に反映します。/ Reflect the color choice in the opacity field.
         *
         * @param {boolean} resetsValue - カラーごとの既定値で上書きする場合は true。
         * @returns {void}
         */
        function syncOpacityEnabled(resetsValue) {
            cellUI.opacityInput.enabled = cellUI.blackCellRadio.value || cellUI.whiteCellRadio.value;
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
            cellUI.opacityInput.enabled = false;
        }

        /**
         * 対象を切り替えます（「_target」矩形は選択中だけ非表示）。/ Switch the target area.
         *
         * @returns {void}
         */
        function handleTargetChange() {
            /* 表示状態を変える前にプレビューを剥がす。順序を逆にすると、
               app.undo() がプレビューではなく setItemHidden を取り消してしまう
               Peel the preview first; otherwise app.undo() would revert setItemHidden instead */
            clearPreview();
            if (targetRectItem) setItemHidden(targetRectItem, placementUI.rectLayerRadio.value === true);
            updatePreview();
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
            updatePreview();
        }

        /**
         * OK：プレビューを破棄して本番の配置を行い、ダイアログを閉じます。/ Commit the distribution and close.
         *
         * @returns {void}
         */
        function commitDistribution() {
            var cellMode = getCellMode();
            var createdCount = -1;
            var artboardFailed = false;

            /* プレビューを undo で破棄してから本番処理へ（プレビューの痕跡も履歴も残さない）/ Undo the preview before the real run */
            clearPreview();
            if (cellMode === "artboard") {
                try {
                    createdCount = createArtboardsFromCells(doc, readGridMetrics(), baseArtboardIndex);
                } catch (e) {
                    artboardFailed = true;
                }
            }
            /* 本番は undo の対象にしない（ユーザーの取り消し操作に委ねる）/ The real run is left to the user's own undo */
            renderDistribution(false);

            /* 「_target」レイヤーの矩形を再表示（プレビュー時に隠していた場合）/ Show the _target rectangle again */
            if (targetRectItem) setItemHidden(targetRectItem, false);

            /* セル描画を残した場合はその長方形を、それ以外は元の選択を選択状態にする / Select the kept cells, or the original selection */
            if (cellMode !== "keep" || !selectCellRectangles(doc)) {
                doc.selection = placeableItems;
            }

            isCleanedUp = true;
            app.redraw();
            distributeDialog.close(1);

            /* 0 個は入力値が無効でグリッドが成立しなかったケース / Zero means the grid was invalid */
            if (artboardFailed || createdCount === 0) {
                alert(getLabel(LABELS.alert.artboardError));
            } else if (createdCount > 0) {
                alert(createdCount + getLabel(LABELS.alert.artboardCreated));
            }
        }

        /**
         * プレビューを消し、「_target」矩形を再表示します。/ Discard the preview and show the _target rectangle again.
         *
         * @returns {void}
         */
        function discardDialogChanges() {
            clearPreview();
            if (targetRectItem) setItemHidden(targetRectItem, false);
            app.redraw();
        }

        // -----------------------------------------
        // イベント / Event wiring
        // -----------------------------------------

        /**
         * ダイアログのイベントを結び付けます。/ Wire up the dialog events.
         *
         * @returns {void}
         */
        function bindDialogEvents() {
            placementUI.artboardRadio.onClick = handleTargetChange;
            placementUI.backmostRadio.onClick = handleTargetChange;
            placementUI.rectLayerRadio.onClick = handleTargetChange;

            cellUI.keepCellRadio.onClick = cellUI.toGuideRadio.onClick = cellUI.toArtboardRadio.onClick = function () {
                syncCellModeUI();
                updatePreview();
            };

            cellUI.blackCellRadio.onClick = cellUI.whiteCellRadio.onClick = cellUI.transparentCellRadio.onClick = function () {
                /* モードを往復してもカラー選択が失われないように覚えておく / Remember the color across mode switches */
                lastCellColorRadio = cellUI.whiteCellRadio.value ? cellUI.whiteCellRadio
                    : (cellUI.transparentCellRadio.value ? cellUI.transparentCellRadio : cellUI.blackCellRadio);
                syncOpacityEnabled(true);
                updatePreview();
            };

            divisionUI.rowCountInput.onChanging = divisionUI.columnCountInput.onChanging = function () {
                syncGutterEnabled();
                updatePreview();
            };
            divisionUI.gutterInput.onChanging = updatePreview;
            divisionUI.marginInput.onChanging = updatePreview;
            cellUI.opacityInput.onChanging = updatePreview;

            cellUI.transparencyGridButton.onClick = function () {
                app.executeMenuCommand('TransparencyGrid Menu Item');
                transparencyGridToggleCount++;
            };

            footerUI.btnShuffle.onClick = shuffleDistributionOrder;
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

    if (app.documents.length === 0) {
        alert(LABELS.alert.noDocument.ja + "\n" + LABELS.alert.noDocument.en);
    } else {
        showDistributeDialog();
    }

}());
