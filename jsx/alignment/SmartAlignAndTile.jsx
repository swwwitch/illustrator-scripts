#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

重なって配置されたオブジェクトを、横方向または縦方向へ指定した間隔で並べ直します。
行数・列数を指定すればタイル状に、キーオブジェクトを設定すればその位置を基準に配置できます。

詳細は README を参照してください。

### Overview

Redistributes stacked objects along the horizontal or vertical axis at the spacing you specify.
Set a row or column count to tile them, or set a key object to anchor the layout to it.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignAndTile";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-06";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf426908d8bcd"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値 / Initial dialog values */
    var DEFAULT_DIRECTION          = "horizontal"; /* 並べる方向（"horizontal" / "vertical"）/ tiling direction */
    var DEFAULT_LANE_COUNT         = "1";   /* 行数（横）・列数（縦）/ row count (horizontal) or column count (vertical) */
    var DEFAULT_MARGIN             = "0";   /* 横・縦の間隔 / horizontal & vertical spacing */
    var DEFAULT_LINK_MARGINS       = true;  /* 横・縦の間隔を連動 / link both spacings */
    var DEFAULT_USE_PREVIEW_BOUNDS = true;  /* プレビュー境界を使用 / use preview bounds */
    var DEFAULT_USE_GRID           = false; /* グリッド配置 / grid layout */
    var DEFAULT_RANDOMIZE          = false; /* ランダム配置 / random order */

    /* 整列後の位置差をどこまで「動いていない」とみなすか（pt）/ Move tolerance when probing the key object (pt) */
    var KEY_DETECT_TOLERANCE_PT = 0.001;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between columns */
    var FIELD_LABEL_WIDTH  = 40;               /* 項目名の幅 / width of a field label */
    var FIELD_CHAR_WIDTH   = 3;                /* 数値欄の文字数 / character width of a numeric field */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */
    var DIALOG_OFFSET_X    = 300;              /* ダイアログの横位置オフセット / dialog offset X */
    var DIALOG_OFFSET_Y    = 0;                /* ダイアログの縦位置オフセット / dialog offset Y */
    var DIALOG_OPACITY     = 0.97;             /* ダイアログの不透明度 / dialog opacity */

    /**
     * ウィンドウに共通レイアウトを適用する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

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
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 右揃えの項目名を追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} fieldLabelText - 表示する項目名（コロン付き）
     * @returns {StaticText} 生成した項目名
     */
    function addFieldLabel(parentContainer, fieldLabelText) {
        var labelStatic = parentContainer.add("statictext", undefined, fieldLabelText);
        labelStatic.preferredSize.width = FIELD_LABEL_WIDTH;
        labelStatic.justify = "right";
        return labelStatic;
    }

    /**
     * オプションのチェックボックスを追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} checkboxLabel - 表示するラベル
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 生成したチェックボックス
     */
    function addOptionCheckbox(parentContainer, checkboxLabel, initialValue) {
        var createdCheckbox = parentContainer.add("checkbox", undefined, checkboxLabel);
        createdCheckbox.alignment = "left"; /* パネル幅いっぱいに広げない / Do not stretch to the panel width */
        createdCheckbox.value = initialValue;
        return createdCheckbox;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
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
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },
        panel: {
            direction: { ja: "方向", en: "Direction" },
            spacing:   { ja: "間隔", en: "Spacing" },
            alignment: { ja: "揃え", en: "Align" },
            options:   { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            rowCount:    { ja: "行数", en: "Rows" },
            columnCount: { ja: "列数", en: "Cols" },
            hMargin:     { ja: "横", en: "H" },
            vMargin:     { ja: "縦", en: "V" }
        },
        checkbox: {
            linkMargins:      { ja: "連動", en: "Link" },
            useKeyObject:     { ja: "キーオブジェクトを基準", en: "Anchor to key object" },
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            useGrid:          { ja: "グリッド", en: "Grid" },
            randomize:        { ja: "ランダム", en: "Random" }
        },
        radio: {
            directionHorizontal: { ja: "横", en: "Horizontal" },
            directionVertical:   { ja: "縦", en: "Vertical" },
            alignTop:    { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" },
            alignLeft:   { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight:  { ja: "右", en: "Right" },
            alignNone:   { ja: "なし", en: "None" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:     { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            previewError:    { ja: "プレビューでエラーが発生しました", en: "Preview error" },
            unexpectedError: { ja: "エラーが発生しました", en: "An error has occurred" }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('panel','spacing')）
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

    /**
     * コロン付きのラベルを返す（日本語は全角、英語は半角）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} コロン付きのラベル
     */
    function labelText() {
        return getLabel.apply(null, arguments) + ((uiLang === "ja") ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; array index equals the rulerType code */
    var UNITS = [
        { label: "in",    factor: 72.0 },                 /* 0 */
        { label: "mm",    factor: 72.0 / 25.4 },          /* 1 */
        { label: "pt",    factor: 1.0 },                  /* 2 */
        { label: "pica",  factor: 12.0 },                 /* 3 */
        { label: "cm",    factor: 72.0 / 2.54 },          /* 4 */
        { label: "Q/H",   factor: 72.0 / 25.4 * 0.25 },   /* 5 */
        { label: "px",    factor: 1.0 },                  /* 6 */
        { label: "ft/in", factor: 72.0 * 12.0 },          /* 7 */
        { label: "m",     factor: 72.0 / 25.4 * 1000.0 }, /* 8 */
        { label: "yd",    factor: 72.0 * 36.0 },          /* 9 */
        { label: "ft",    factor: 72.0 * 12.0 }           /* 10 */
    ];

    /* pt の添字（単位が特定できないときのフォールバック）/ Index of pt, used as the fallback unit */
    var POINT_UNIT_INDEX = 2;

    /**
     * 定規の単位設定から現在の単位を取得する
     * @returns {object} { label: string, factor: number }
     */
    function getCurrentUnit() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return UNITS[unitCode] || UNITS[POINT_UNIT_INDEX];
    }

    // =========================================
    // 位置とサイズ / Positions and sizes
    // =========================================

    /**
     * アイテムの境界を取得する（プレビュー境界ONならvisible、OFFならgeometric）
     * @param {PageItem} item - 対象オブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(item, usePreviewBounds) {
        return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
    }

    /**
     * 控えておいた位置へ戻す
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {Array} positions - [[left, top], ...] の配列
     * @returns {void}
     */
    function resetPositions(items, positions) {
        for (var i = 0; i < items.length; i++) {
            items[i].left = positions[i][0];
            items[i].top = positions[i][1];
        }
    }

    /**
     * 指定量だけまとめて移動する
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {number} dx - 横方向の移動量（pt）
     * @param {number} dy - 縦方向の移動量（pt）
     * @returns {void}
     */
    function shiftItems(items, dx, dy) {
        if (!dx && !dy) {
            return;
        }
        for (var i = 0; i < items.length; i++) {
            if (!items[i]) continue;
            items[i].left += dx;
            items[i].top += dy;
        }
    }

    /**
     * 左端の座標順に並べ替えた複製を返す
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortedCopyByLeft(items) {
        var copiedItems = items.slice();
        copiedItems.sort(function(a, b) {
            return a.left - b.left;
        });
        return copiedItems;
    }

    /**
     * 上端の座標順（上から下、同じなら左から右）に並べ替えた複製を返す
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortedCopyByTop(items) {
        var copiedItems = items.slice();
        copiedItems.sort(function(a, b) {
            if (a.top !== b.top) return b.top - a.top;
            return a.left - b.left;
        });
        return copiedItems;
    }

    /**
     * ランダムに並べ替えた複製を返す
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function shuffledCopy(items) {
        var copiedItems = items.slice();
        for (var i = copiedItems.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapped = copiedItems[i];
            copiedItems[i] = copiedItems[j];
            copiedItems[j] = swapped;
        }
        return copiedItems;
    }

    /**
     * 指定したオブジェクトを先頭へ移した複製を返す
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {object} targetItem - 先頭へ移すオブジェクト
     * @returns {PageItem[]} 並べ替えた配列（対象が見つからないときはそのままの複製）
     */
    function movedToFront(items, targetItem) {
        var reordered = items.slice();
        for (var i = 0; i < reordered.length; i++) {
            if (reordered[i] !== targetItem) continue;
            reordered.splice(i, 1);
            reordered.unshift(targetItem);
            break;
        }
        return reordered;
    }

    /**
     * 選択範囲全体の左上を取得する
     * @param {PageItem[]} items - 対象オブジェクト
     * @returns {number[]} [左端, 上端]
     */
    function getBlockOrigin(items) {
        var blockLeft = null;
        var blockTop = null;
        for (var i = 0; i < items.length; i++) {
            if (!items[i]) continue;
            if (blockLeft === null || items[i].left < blockLeft) blockLeft = items[i].left;
            if (blockTop === null || items[i].top > blockTop) blockTop = items[i].top;
        }
        return [blockLeft, blockTop];
    }

    /**
     * もっとも大きいアイテムの幅と高さを取得する（グリッドのセルサイズ）
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
     * @returns {object} { width: number, height: number }
     */
    function getMaxItemSize(items, usePreviewBounds) {
        var maxWidth = 0;
        var maxHeight = 0;
        for (var i = 0; i < items.length; i++) {
            if (!items[i]) continue;
            var itemBounds = getItemBounds(items[i], usePreviewBounds);
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];
            if (itemWidth > maxWidth) maxWidth = itemWidth;
            if (itemHeight > maxHeight) maxHeight = itemHeight;
        }
        return { width: maxWidth, height: maxHeight };
    }

    /**
     * セル内での揃え量を求める（X軸は左→右、Y軸は上→下を start→end とする）
     * @param {string} alignMode - "start" / "center" / "end" / "none"
     * @param {number} cellStart - セルの起点（左端または上端）
     * @param {number} cellEnd - セルの終点（右端または下端）
     * @param {number} itemStart - オブジェクトの起点
     * @param {number} itemEnd - オブジェクトの終点
     * @returns {number} 移動量（pt）
     */
    function getAlignDelta(alignMode, cellStart, cellEnd, itemStart, itemEnd) {
        if (alignMode === "none") {
            return 0;
        }
        if (alignMode === "center") {
            return (cellStart + cellEnd) / 2 - (itemStart + itemEnd) / 2;
        }
        if (alignMode === "end") {
            return cellEnd - itemEnd;
        }
        return cellStart - itemStart;
    }

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================
    // Illustrator の DOM にキーオブジェクトを示すプロパティは無いため、整列コマンドで実測して特定する。
    // キーオブジェクトが設定されていると、どの向きに整列してもそのオブジェクトだけは動かない。
    // 候補が0個または2個以上のときは判定不能として null を返し、UI側でこの基準をディムする。

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * @param {PageItem[]} items - 判定対象のオブジェクト
     * @returns {object} キーオブジェクト。判定できないときは null
     */
    function detectKeyObject(items) {
        if (!items || items.length < 2) {
            return null;
        }
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        var i;
        for (i = 0; i < items.length; i++) {
            stayedPut.push(true);
        }

        for (var c = 0; c < alignCommands.length; c++) {
            var savedPositions = [];
            for (i = 0; i < items.length; i++) {
                savedPositions.push([items[i].left, items[i].top]);
            }
            app.redraw(); /* 直前のDOM変更が反映されていないと executeMenuCommand は空振りする / executeMenuCommand misfires without a redraw */
            app.executeMenuCommand(alignCommands[c]);
            for (i = 0; i < items.length; i++) {
                if (Math.abs(items[i].left - savedPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                    Math.abs(items[i].top - savedPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                    stayedPut[i] = false;
                }
            }
            /* 整列は検出のための試行なので、その場で元の位置へ戻す / Undo the probe right away */
            resetPositions(items, savedPositions);
        }
        app.redraw();

        var foundItem = null;
        for (i = 0; i < items.length; i++) {
            if (!stayedPut[i]) continue;
            if (foundItem !== null) return null; /* 複数残った＝判定不能 / Ambiguous */
            foundItem = items[i];
        }
        return foundItem;
    }

    // =========================================
    // プレビュー管理 / Preview management
    // =========================================

    /**
     * プレビューの適用・巻き戻し・確定をまとめて管理する
     * @constructor
     */
    function PreviewManager() {
        /* プレビュー中に実行したアクションの回数 / Number of preview actions executed */
        this.undoDepth = 0;

        /**
         * 変更操作を実行し、履歴としてカウントする
         * @param {Function} previewAction - 実行したい処理
         * @returns {void}
         */
        this.addStep = function(previewAction) {
            try {
                previewAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(labelText('alert', 'previewError') + " " + e);
            }
        };

        /**
         * プレビューのために行った変更をすべて取り消す
         * @returns {void}
         */
        this.rollback = function() {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * 現在の状態を確定する（プレビューを巻き戻してから1回だけ本番処理を実行）
         * @param {Function} [finalAction] - 巻き戻したあとに実行する処理
         * @returns {void}
         */
        this.confirm = function(finalAction) {
            if (finalAction) {
                this.rollback();
                finalAction();
            }
            this.undoDepth = 0;
        };
    }

    // =========================================
    // 配置処理 / Arranging
    // =========================================

    /**
     * 1行ずつセルに割り当てて配置する
     * @param {PageItem[]} orderedItems - 配置順に並べたオブジェクト
     * @param {object} arrangeSettings - 配置設定
     * @returns {void}
     */
    function placeItems(orderedItems, arrangeSettings) {
        var usePreviewBounds = arrangeSettings.usePreviewBounds;
        var isHorizontal = (arrangeSettings.direction === "horizontal");
        var cellSize = getMaxItemSize(orderedItems, usePreviewBounds);

        /* 先頭オブジェクトの位置を配置の起点にする / The first item defines the origin of the layout */
        var startBounds = getItemBounds(orderedItems[0], usePreviewBounds);
        /* 主軸＝並べる方向、副軸＝行・列が積み重なる方向 / Main axis follows the tiling direction; lanes stack along the cross axis */
        var mainOrigin = isHorizontal ? startBounds[0] : startBounds[1];
        var crossOrigin = isHorizontal ? startBounds[1] : startBounds[0];
        var mainGap = isHorizontal ? arrangeSettings.hMarginPt : arrangeSettings.vMarginPt;
        var crossGap = isHorizontal ? arrangeSettings.vMarginPt : arrangeSettings.hMarginPt;
        var laneSize = isHorizontal ? cellSize.height : cellSize.width;
        /* 上方向がプラスのY軸に合わせ、横並びは下へ、縦並びは右へレーンを送る / Lanes go down (horizontal) or right (vertical) */
        var laneDirection = isHorizontal ? -1 : 1;

        /* 主軸の揃えはグリッド時のみ意味を持つ（セル＝オブジェクトの大きさでは差が出ない）/ Main-axis align only matters in grid mode */
        var mainAlign = arrangeSettings.useGrid ? (isHorizontal ? arrangeSettings.hAlign : arrangeSettings.vAlign) : "start";
        var crossAlign = isHorizontal ? arrangeSettings.vAlign : arrangeSettings.hAlign;

        var remainingItems = orderedItems.length;
        var index = 0;
        for (var lane = 0; lane < arrangeSettings.laneCount; lane++) {
            /* 残りを残りのレーン数で割り、指定した行数・列数を使い切る / Split the remainder so every lane is used */
            var perLane = Math.ceil(remainingItems / (arrangeSettings.laneCount - lane));
            remainingItems -= perLane;
            var laneOffset = lane * (laneSize + crossGap) * laneDirection;
            var crossStart = crossOrigin + laneOffset;
            var crossEnd = crossStart + laneSize * laneDirection;
            var mainStart = mainOrigin;

            for (var i = 0; i < perLane && index < orderedItems.length; i++, index++) {
                var item = orderedItems[index];
                if (!item) continue;

                var itemBounds = getItemBounds(item, usePreviewBounds);
                var itemMainSize = isHorizontal ? (itemBounds[2] - itemBounds[0]) : (itemBounds[1] - itemBounds[3]);
                var cellMainSize = arrangeSettings.useGrid ? (isHorizontal ? cellSize.width : cellSize.height) : itemMainSize;
                var mainEnd = mainStart + cellMainSize * (isHorizontal ? 1 : -1);

                /* 主軸：セル内での揃え（なし＝その軸は動かさない）/ Main axis: align inside the cell ("none" leaves it alone) */
                var mainDelta = isHorizontal
                    ? getAlignDelta(mainAlign, mainStart, mainEnd, itemBounds[0], itemBounds[2])
                    : getAlignDelta(mainAlign, mainStart, mainEnd, itemBounds[1], itemBounds[3]);
                /* 副軸：レーンの帯へ揃える（なし＝レーンの送り分だけ平行移動）/ Cross axis: align to the lane band ("none" only applies the lane offset) */
                var crossDelta = (crossAlign === "none")
                    ? laneOffset
                    : (isHorizontal
                        ? getAlignDelta(crossAlign, crossStart, crossEnd, itemBounds[1], itemBounds[3])
                        : getAlignDelta(crossAlign, crossStart, crossEnd, itemBounds[0], itemBounds[2]));

                item.left += isHorizontal ? mainDelta : crossDelta;
                item.top += isHorizontal ? crossDelta : mainDelta;

                /* 次のセルへ / Advance to the next cell */
                mainStart = mainEnd + mainGap * (isHorizontal ? 1 : -1);
            }
        }
    }

    /**
     * 設定に従ってオブジェクトを並べ直す（プレビュー・確定の共通処理）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} arrangeSettings - 配置設定
     * @returns {void}
     */
    function arrangeItems(targetItems, arrangeSettings) {
        if (!targetItems || targetItems.length === 0) {
            return;
        }
        var orderedItems;
        if (arrangeSettings.randomize) {
            orderedItems = shuffledCopy(targetItems);
        } else {
            orderedItems = (arrangeSettings.direction === "horizontal") ? sortedCopyByLeft(targetItems) : sortedCopyByTop(targetItems);
        }
        /* キーオブジェクトは配置の起点にする（そこから右／下へ並べる）/ The key object becomes the origin, so the rest follow to its right or below */
        var useKeyObject = arrangeSettings.useKeyObject && arrangeSettings.keyObject && arrangeSettings.keyOrigin;
        if (useKeyObject) {
            orderedItems = movedToFront(orderedItems, arrangeSettings.keyObject);
        }
        /* 並べ替える前の左上（ランダム時に位置を戻す基準）/ Top-left before the layout, used to keep a random block in place */
        var blockOrigin = getBlockOrigin(targetItems);

        placeItems(orderedItems, arrangeSettings);

        /* 基準の補正：キーオブジェクトを優先し、なければランダム時のみ左上を合わせる / Anchor correction: key object first, random block otherwise */
        if (useKeyObject) {
            shiftItems(orderedItems,
                arrangeSettings.keyOrigin[0] - arrangeSettings.keyObject.left,
                arrangeSettings.keyOrigin[1] - arrangeSettings.keyObject.top);
        } else if (arrangeSettings.randomize) {
            shiftItems(orderedItems,
                blockOrigin[0] - orderedItems[0].left,
                blockOrigin[1] - orderedItems[0].top);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 編集テキストで上下キーによる数値変更を有効にする
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負数を許可するかどうか
     * @param {Function} onUpdate - 値を変更したあとに呼ぶ処理
     * @param {number} [minimumValue] - 下限値（省略時は下限なし）
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onUpdate, minimumValue) {
        editText.addEventListener("keydown", function(event) {
            if (editText.text.length === 0) return;
            var value = Number(editText.text);
            if (isNaN(value)) return;
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            if (keyboard.shiftKey) {
                /* 押した向きの10の倍数へスナップ / Snap to the next multiple of 10 in that direction */
                value = isUp ? (Math.ceil((value + 1) / 10) * 10) : (Math.floor((value - 1) / 10) * 10);
            } else {
                value += isUp ? 1 : -1;
            }
            if (!allowNegative && value < 0) value = 0;
            if (typeof minimumValue === "number" && value < minimumValue) value = minimumValue;

            event.preventDefault();
            editText.text = value;
            if (typeof onUpdate === "function") {
                onUpdate();
            }
        });
    }

    /**
     * 配置ダイアログを表示し、プレビューしながら設定を決める
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} keyObject - キーオブジェクト（未検出のときは null）
     * @returns {object} 確定した配置設定。キャンセル時は null
     */
    function showArrangeDialog(targetItems, keyObject) {
        var dialogWindow = new Window("dialog", getLabel('dialog', 'title') + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow);
        dialogWindow.opacity = DIALOG_OPACITY;
        dialogWindow.onShow = function() {
            dialogWindow.location = [dialogWindow.location[0] + DIALOG_OFFSET_X, dialogWindow.location[1] + DIALOG_OFFSET_Y];
        };

        var previewManager = new PreviewManager();
        /* キャンセル時に戻せるよう、境界計算の環境設定を控える / Remember the bounds preference so Cancel can restore it */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");
        /* キーオブジェクトのプレビュー前の位置 / Key object position before any preview */
        var keyOrigin = keyObject ? [keyObject.left, keyObject.top] : null;

        /* 方向パネル（並べる方向と行数／列数）/ Direction panel: tiling direction and lane count */
        var directionPanel = addPanel(dialogWindow, getLabel('panel', 'direction'));

        var directionRow = directionPanel.add("group");
        setupRow(directionRow);
        var horizontalRadio = directionRow.add("radiobutton", undefined, getLabel('radio', 'directionHorizontal'));
        var verticalRadio = directionRow.add("radiobutton", undefined, getLabel('radio', 'directionVertical'));
        horizontalRadio.value = (DEFAULT_DIRECTION === "horizontal");
        verticalRadio.value = !horizontalRadio.value;

        /* 行数（横）／列数（縦）/ Row count (horizontal) or column count (vertical) */
        var laneCountRow = directionPanel.add("group");
        setupRow(laneCountRow);
        var laneCountLabel = addFieldLabel(laneCountRow, "");
        var laneCountInput = laneCountRow.add("edittext", undefined, DEFAULT_LANE_COUNT);
        laneCountInput.characters = FIELD_CHAR_WIDTH;

        /* 直近でプレビューへ反映した値（同じ値での二重更新を避ける）/ Value last pushed to the preview */
        var appliedLaneCountText = laneCountInput.text;

        /**
         * 行数・列数が変わったときだけプレビューを更新する
         * @returns {void}
         */
        function updatePreviewForLaneCount() {
            if (laneCountInput.text === appliedLaneCountText) return;
            appliedLaneCountText = laneCountInput.text;
            updatePreview();
        }

        /* 入力中は数値として読めるときだけ反映する（打っている途中で書き換えない）/ While typing, refresh only when the text parses */
        laneCountInput.onChanging = function() {
            var typedLaneCount = parseInt(laneCountInput.text, 10);
            if (isNaN(typedLaneCount) || typedLaneCount < 1) return;
            updatePreviewForLaneCount();
        };

        /* 確定時に1以上の整数へ丸める（表示と実際に使う値を一致させる）/ Snap to an integer of 1 or more on commit */
        laneCountInput.onChange = function() {
            var laneCountValue = parseInt(laneCountInput.text, 10);
            if (isNaN(laneCountValue) || laneCountValue < 1) laneCountValue = 1;
            laneCountInput.text = laneCountValue;
            updatePreviewForLaneCount();
        };
        changeValueByArrowKey(laneCountInput, false, updatePreviewForLaneCount, 1);

        var gridCheckbox = addOptionCheckbox(directionPanel, getLabel('checkbox', 'useGrid'), DEFAULT_USE_GRID);

        /* 間隔パネル（左＝横・縦の入力、右＝連動）/ Spacing panel: fields on the left, link on the right */
        var spacingPanel = addPanel(dialogWindow, getLabel('panel', 'spacing') + " (" + getCurrentUnit().label + ")");
        var spacingRow = spacingPanel.add("group");
        setupRow(spacingRow, "left", COLUMN_SPACING);

        var marginColumn = spacingRow.add("group");
        marginColumn.orientation = "column";
        marginColumn.alignChildren = ["left", "center"];
        marginColumn.spacing = PANEL_SPACING;

        var hMarginRow = marginColumn.add("group");
        setupRow(hMarginRow);
        addFieldLabel(hMarginRow, labelText('fieldLabel', 'hMargin'));
        var hMarginInput = hMarginRow.add("edittext", undefined, DEFAULT_MARGIN);
        hMarginInput.characters = FIELD_CHAR_WIDTH;
        changeValueByArrowKey(hMarginInput, true, syncMarginsAndPreview);

        var vMarginRow = marginColumn.add("group");
        setupRow(vMarginRow);
        addFieldLabel(vMarginRow, labelText('fieldLabel', 'vMargin'));
        var vMarginInput = vMarginRow.add("edittext", undefined, DEFAULT_MARGIN);
        vMarginInput.characters = FIELD_CHAR_WIDTH;
        changeValueByArrowKey(vMarginInput, true, updatePreview);

        var linkCheckbox = spacingRow.add("checkbox", undefined, getLabel('checkbox', 'linkMargins'));
        linkCheckbox.value = DEFAULT_LINK_MARGINS;
        /* 連動中は縦をディムして横の値に合わせる / While linked, dim V and mirror H */
        vMarginInput.enabled = !linkCheckbox.value;
        if (linkCheckbox.value) {
            vMarginInput.text = hMarginInput.text;
        }

        /* 揃えパネル（上段＝上下、下段＝左右）/ Align panel: vertical row on top, horizontal row below */
        var alignmentPanel = addPanel(dialogWindow, getLabel('panel', 'alignment'));

        var vAlignRow = alignmentPanel.add("group");
        setupRow(vAlignRow);
        var vAlignTopRadio = vAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignTop'));
        var vAlignMiddleRadio = vAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignMiddle'));
        var vAlignBottomRadio = vAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignBottom'));
        var vAlignNoneRadio = vAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignNone'));
        var vAlignRadios = [vAlignTopRadio, vAlignMiddleRadio, vAlignBottomRadio, vAlignNoneRadio];
        vAlignTopRadio.value = true;

        var hAlignRow = alignmentPanel.add("group");
        setupRow(hAlignRow);
        var hAlignLeftRadio = hAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignLeft'));
        var hAlignCenterRadio = hAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignCenter'));
        var hAlignRightRadio = hAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignRight'));
        var hAlignNoneRadio = hAlignRow.add("radiobutton", undefined, getLabel('radio', 'alignNone'));
        var hAlignRadios = [hAlignLeftRadio, hAlignCenterRadio, hAlignRightRadio, hAlignNoneRadio];
        hAlignLeftRadio.value = true;

        /* オプションパネル / Options panel */
        var optionsPanel = addPanel(dialogWindow, getLabel('panel', 'options'));

        /* キーオブジェクトが未検出のときはディム / Dimmed when no key object is detected */
        var keyObjectCheckbox = addOptionCheckbox(optionsPanel, getLabel('checkbox', 'useKeyObject'), !!keyObject);
        keyObjectCheckbox.enabled = !!keyObject;
        var previewBoundsCheckbox = addOptionCheckbox(optionsPanel, getLabel('checkbox', 'usePreviewBounds'), DEFAULT_USE_PREVIEW_BOUNDS);
        var randomizeCheckbox = addOptionCheckbox(optionsPanel, getLabel('checkbox', 'randomize'), DEFAULT_RANDOMIZE);

        /* ボタンエリア（左右中央）/ Button bar, centered */
        var btnRowGroup = dialogWindow.add("group");
        setupRow(btnRowGroup, "center");
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.add("button", undefined, getLabel('button', 'cancel'), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel('button', 'ok'), { name: "ok" });

        /**
         * ラジオボタンの一覧をまとめて有効・無効にする
         * @param {RadioButton[]} radioList - 対象のラジオボタン
         * @param {boolean} enabled - 有効にするかどうか
         * @returns {void}
         */
        function setRadiosEnabled(radioList, enabled) {
            for (var i = 0; i < radioList.length; i++) {
                radioList[i].enabled = enabled;
            }
        }

        /**
         * 方向とグリッドの状態に合わせてUIを整える（項目名と揃えの操作可否）
         * @returns {void}
         */
        function syncDirectionUI() {
            var isHorizontal = horizontalRadio.value;
            laneCountLabel.text = labelText('fieldLabel', isHorizontal ? 'rowCount' : 'columnCount');
            /* 主軸（並べる方向）の揃えはグリッド時のみ有効 / Main-axis align is available in grid mode only */
            setRadiosEnabled(hAlignRadios, isHorizontal ? gridCheckbox.value : true);
            setRadiosEnabled(vAlignRadios, isHorizontal ? true : gridCheckbox.value);
        }
        syncDirectionUI();

        /**
         * ダイアログの入力内容を配置設定として読み取る
         * @returns {object} 配置設定
         */
        function readArrangeSettings() {
            var unitFactor = getCurrentUnit().factor;

            var hMarginValue = parseFloat(hMarginInput.text);
            if (isNaN(hMarginValue)) hMarginValue = 0;
            var vMarginValue = parseFloat(vMarginInput.text);
            if (isNaN(vMarginValue)) vMarginValue = 0;
            var laneCount = parseInt(laneCountInput.text, 10);
            if (isNaN(laneCount) || laneCount < 1) laneCount = 1;

            /* 揃えは軸に依らない形（start / center / end / none）で持つ / Align values are axis-neutral */
            var vAlign = "start";
            if (vAlignMiddleRadio.value) vAlign = "center";
            else if (vAlignBottomRadio.value) vAlign = "end";
            else if (vAlignNoneRadio.value) vAlign = "none";

            var hAlign = "start";
            if (hAlignCenterRadio.value) hAlign = "center";
            else if (hAlignRightRadio.value) hAlign = "end";
            else if (hAlignNoneRadio.value) hAlign = "none";

            return {
                direction: horizontalRadio.value ? "horizontal" : "vertical",
                laneCount: laneCount,
                hMarginPt: hMarginValue * unitFactor,
                vMarginPt: vMarginValue * unitFactor,
                vAlign: vAlign,
                hAlign: hAlign,
                usePreviewBounds: previewBoundsCheckbox.value,
                useGrid: gridCheckbox.value,
                randomize: randomizeCheckbox.value,
                useKeyObject: keyObjectCheckbox.value,
                keyObject: keyObject,
                keyOrigin: keyOrigin
            };
        }

        /**
         * Undo履歴を汚さずにプレビューを更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            /* 境界計算に使う環境設定は変わったときだけ書き換える（切り替えた直後は再描画しないと古い境界のまま計算される）/ Write the bounds preference only when it changes; without a redraw the old bounds are used */
            if (app.preferences.getBooleanPreference("includeStrokeInBounds") !== previewBoundsCheckbox.value) {
                app.preferences.setBooleanPreference("includeStrokeInBounds", previewBoundsCheckbox.value);
                app.redraw();
            }
            previewManager.addStep(function() {
                arrangeItems(targetItems, readArrangeSettings());
            });
        }

        /**
         * 方向・グリッドに合わせてUIを整えてからプレビューを更新する
         * @returns {void}
         */
        function syncDirectionAndPreview() {
            syncDirectionUI();
            updatePreview();
        }

        /**
         * 連動がONなら横の値を縦へ反映してからプレビューを更新する
         * @returns {void}
         */
        function syncMarginsAndPreview() {
            if (linkCheckbox.value) {
                vMarginInput.text = hMarginInput.text;
            }
            updatePreview();
        }

        hMarginInput.onChanging = syncMarginsAndPreview;
        hMarginInput.onChange = syncMarginsAndPreview;
        vMarginInput.onChanging = updatePreview;
        linkCheckbox.onClick = function() {
            vMarginInput.enabled = !linkCheckbox.value;
            syncMarginsAndPreview();
        };
        horizontalRadio.onClick = syncDirectionAndPreview;
        verticalRadio.onClick = syncDirectionAndPreview;
        vAlignTopRadio.onClick = updatePreview;
        vAlignMiddleRadio.onClick = updatePreview;
        vAlignBottomRadio.onClick = updatePreview;
        vAlignNoneRadio.onClick = updatePreview;
        hAlignLeftRadio.onClick = updatePreview;
        hAlignCenterRadio.onClick = updatePreview;
        hAlignRightRadio.onClick = updatePreview;
        hAlignNoneRadio.onClick = updatePreview;
        keyObjectCheckbox.onClick = updatePreview;
        previewBoundsCheckbox.onClick = updatePreview;
        randomizeCheckbox.onClick = updatePreview;
        gridCheckbox.onClick = function() {
            if (gridCheckbox.value) {
                /* グリッドは天地・左右とも中央を既定にする / Grid defaults to centered on both axes */
                vAlignMiddleRadio.value = true;
                hAlignCenterRadio.value = true;
            }
            syncDirectionAndPreview();
        };

        updatePreview();
        laneCountInput.active = true;

        if (dialogWindow.show() !== 1) {
            /* キャンセル：プレビューを巻き戻し、環境設定も元に戻す / Cancel: roll back the preview and the preference */
            previewManager.rollback();
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return null;
        }

        var arrangeSettings = readArrangeSettings();
        /* 1回のUndoで取り消せるように、巻き戻してから一度だけ実行する / Confirm as a single undoable action */
        previewManager.confirm(function() {
            arrangeItems(targetItems, arrangeSettings);
            /* 環境設定はスクリプトの外へ影響を残さないよう元に戻す / Restore the preference so the script leaves no global side effect */
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
        });
        return arrangeSettings;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを取得し、キーオブジェクトを判定してダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert', 'noDocument'));
            return;
        }
        var doc = app.activeDocument;
        var targetItems = doc.selection;
        if (!targetItems || targetItems.length === 0) {
            alert(getLabel('alert', 'noSelection'));
            return;
        }

        /* キーオブジェクトを検出（整列コマンドで一時的に動かして元へ戻す）/ Detect the key object; it aligns temporarily, then restores */
        var keyObject = null;
        try {
            keyObject = detectKeyObject(targetItems);
        } catch (err) {
            /* 検出に失敗しても配置は続行する（チェックボックスがディムされるだけ）/ Keep going; only the checkbox is dimmed */
            $.writeln(SCRIPT_NAME + ": キーオブジェクトの検出に失敗 / key object detection failed — " + err);
        }

        showArrangeDialog(targetItems, keyObject);
    }

    try {
        main();
    } catch (e) {
        alert(labelText('alert', 'unexpectedError') + " " + e.message);
    }

})();
