#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

重なって配置されたオブジェクトを、横方向へ等間隔に並べ直します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile-simple.md

### Overview

Redistributes stacked objects evenly along the horizontal axis.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile-simple.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignAndTile-simple";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile-simple.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile-simple.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion https://github.com/johnwun/js4ai/blob/master/distributeStackedObjects.jsx
 * @discussion https://gorolib.blog.jp/archives/77282974.html
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［プレビュー境界を使用］の初期状態 / initial state of the preview-bounds checkbox */
    var DEFAULT_USE_PREVIEW_BOUNDS = true;
    /* ［順序をランダムに］の初期状態 / initial state of the randomize checkbox */
    var DEFAULT_RANDOMIZE_ORDER = false;
    /* 間隔の初期値（表示単位）/ initial spacing, in the ruler unit */
    var DEFAULT_SPACING = "0";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OPACITY = 0.97;
    var PANEL_MARGINS = [15, 20, 15, 10];
    var OPTION_GROUP_MARGINS = [15, 5, 15, 5];
    var SPACING_INPUT_CHARS = 3;

    /* プレビューを走らせる最短間隔（ミリ秒）。ドラッグ中の連続発火を抑える
       Minimum interval between previews, so dragging does not fire them back to back */
    var PREVIEW_MIN_INTERVAL_MS = 80;
    /* プレビューを遅らせる時間（ミリ秒）。ScriptUI のイベント内で重い処理を走らせない
       Delay before the preview runs, keeping heavy work out of the ScriptUI event handler */
    var PREVIEW_DELAY_MS = 60;

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* ダイアログ位置をセッション内で覚えておくためのキー / key holding the dialog position */
    var DIALOG_POSITION_KEY = "__SmartAlignAndTileSimple_DialogPosition__";
    /* scheduleTask から呼ぶためのプレビュー関数の置き場 / slot the scheduled task calls back into */
    var PREVIEW_CALLBACK_KEY = "__SmartAlignAndTileSimple_Preview__";

    // =========================================
    // 単位 / Units
    // =========================================

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
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },
        panel: {
            direction:      { ja: "方向", en: "Direction" },
            spacing:        { ja: "間隔", en: "Spacing" },
            alignHorizontal:{ ja: "揃え（左右）", en: "Align (H)" },
            alignVertical:  { ja: "揃え（上下）", en: "Align (V)" }
        },
        radio: {
            directionAuto:       { ja: "自動", en: "Auto" },
            directionVertical:   { ja: "縦", en: "Vertical" },
            directionHorizontal: { ja: "横", en: "Horizontal" },
            alignNone:   { ja: "なし", en: "None" },
            alignLeft:   { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight:  { ja: "右", en: "Right" },
            alignTop:    { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        checkbox: {
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            randomizeOrder:   { ja: "順序をランダムに", en: "Randomize order" }
        },
        tooltip: {
            directionAuto: {
                ja: "選択範囲が横長なら横並び、縦長なら縦並びとして扱います。",
                en: "Lays the objects out in a row when the selection is wider than tall, in a column otherwise."
            },
            directionVertical:   { ja: "上から下へ縦に並べます。", en: "Stacks the objects from top to bottom." },
            directionHorizontal: { ja: "左から右へ横に並べます。", en: "Lays the objects out from left to right." },
            spacing: {
                ja: "隣り合うオブジェクトのあいだにあける間隔です。↑↓キーで増減できます（Shiftで10刻み）。",
                en: "Gap left between neighbouring objects. The arrow keys step the value (Shift for 10)."
            },
            alignHorizontal: {
                ja: "縦に並べたときの左右の揃え方です。N／L／C／R キーでも切り替えられます。",
                en: "Horizontal alignment used when stacking vertically. The keys N / L / C / R switch it."
            },
            alignVertical: {
                ja: "横に並べたときの上下の揃え方です。N／T／M／B キーでも切り替えられます。",
                en: "Vertical alignment used when laying out horizontally. The keys N / T / M / B switch it."
            },
            usePreviewBounds: {
                ja: "線幅や効果を含めた見た目の端を基準にします。オフにするとパスの端が基準になります。",
                en: "Measures by the visible edges including strokes and effects. Off measures the path edges."
            },
            randomizeOrder: {
                ja: "並べ直す順序をシャッフルします。全体の左上の位置は変わりません。",
                en: "Shuffles the order the objects are laid out in. The top-left of the whole set stays put."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.spacing" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // 単位と境界 / Units and bounds
    // =========================================

    /**
     * オブジェクトの境界を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(pageItem, usePreviewBounds) {
        return usePreviewBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
    }

    /**
     * 選択範囲の横幅と高さを返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {{spanX: number, spanY: number}} 横幅と高さ
     */
    function getSelectionSpan(targetItems, usePreviewBounds) {
        if (targetItems.length === 0) return { spanX: 0, spanY: 0 };

        var unionBounds = getItemBounds(targetItems[0], usePreviewBounds).slice(0);
        for (var i = 1; i < targetItems.length; i++) {
            var itemBounds = getItemBounds(targetItems[i], usePreviewBounds);
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return { spanX: unionBounds[2] - unionBounds[0], spanY: unionBounds[1] - unionBounds[3] };
    }

    /**
     * 選択範囲の形から並べる向きを推定する
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {string} "horizontal" または "vertical"
     */
    function detectDirection(targetItems) {
        /* 判定はプレビュー境界の設定に左右されない geometricBounds で行う / judge on stable geometry */
        var selectionSpan = getSelectionSpan(targetItems, false);
        return (selectionSpan.spanX >= selectionSpan.spanY) ? "horizontal" : "vertical";
    }

    /**
     * 上から下（同じ高さなら左から右）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortTopToBottom(targetItems) {
        var sortedItems = targetItems.slice(0);
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.top !== itemB.top) return itemB.top - itemA.top;
            return itemA.left - itemB.left;
        });
        return sortedItems;
    }

    /**
     * 左から右（同じ位置なら上から下）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortLeftToRight(targetItems) {
        var sortedItems = targetItems.slice(0);
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.left !== itemB.left) return itemA.left - itemB.left;
            return itemB.top - itemA.top;
        });
        return sortedItems;
    }

    // =========================================
    // 位置の控えと復元 / Position snapshot
    // =========================================

    /**
     * 元の位置を控え、プレビューのたびにそこへ戻せるようにする
     * Undo に頼らず位置を直接書き戻すので、ヒストリーを汚さない。
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {object} restore メソッドを持つスナップショット
     */
    function createPositionSnapshot(targetItems) {
        var snapshotPositions = [];
        for (var i = 0; i < targetItems.length; i++) {
            /* 位置を持たないアイテムは復元の対象外にする / items without a position are skipped */
            try {
                snapshotPositions.push([targetItems[i].left, targetItems[i].top]);
            } catch (e) {
                snapshotPositions.push(null);
            }
        }

        return {
            /**
             * 控えた位置へ戻す
             * @returns {void}
             */
            restore: function () {
                for (var i = 0; i < targetItems.length; i++) {
                    if (!snapshotPositions[i]) continue;
                    targetItems[i].left = snapshotPositions[i][0];
                    targetItems[i].top = snapshotPositions[i][1];
                }
                app.redraw();
            }
        };
    }

    // =========================================
    // 並べ直しの計算 / Layout
    // =========================================

    /**
     * 並べ直す順序を決める（ランダム指定なら順序を保ったシャッフルを使う）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} layoutOptions - direction / randomOrder を持つ設定
     * @returns {PageItem[]} 並べる順に並んだ配列
     */
    function getLayoutOrder(targetItems, layoutOptions) {
        if (layoutOptions.randomOrder) {
            var shuffledItems = [];
            for (var i = 0; i < layoutOptions.randomOrder.length; i++) {
                shuffledItems.push(targetItems[layoutOptions.randomOrder[i]]);
            }
            return shuffledItems;
        }
        return (layoutOptions.direction === "horizontal") ? sortLeftToRight(targetItems) : sortTopToBottom(targetItems);
    }

    /**
     * 指定した長さのシャッフル済みインデックス列を作る
     * @param {number} itemCount - 要素数
     * @returns {number[]} シャッフルしたインデックス
     */
    function createShuffledOrder(itemCount) {
        var order = [];
        for (var i = 0; i < itemCount; i++) order.push(i);
        for (var j = order.length - 1; j > 0; j--) {
            var swapIndex = Math.floor(Math.random() * (j + 1));
            var swapped = order[j];
            order[j] = order[swapIndex];
            order[swapIndex] = swapped;
        }
        return order;
    }

    /**
     * 並べた列のうち、指定軸方向で最も大きい寸法を返す
     * @param {PageItem[]} orderedItems - 対象オブジェクト
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @param {boolean} isWidth - 幅を求めるなら true、高さなら false
     * @returns {number} 最大の幅または高さ
     */
    function getLargestItemSize(orderedItems, usePreviewBounds, isWidth) {
        var largestSize = 0;
        for (var i = 0; i < orderedItems.length; i++) {
            var itemBounds = getItemBounds(orderedItems[i], usePreviewBounds);
            var itemSize = isWidth ? (itemBounds[2] - itemBounds[0]) : (itemBounds[1] - itemBounds[3]);
            if (itemSize > largestSize) largestSize = itemSize;
        }
        return largestSize;
    }

    /**
     * 横一列に並べ、指定があれば上下方向にも揃える
     * @param {PageItem[]} orderedItems - 並べる順に並んだオブジェクト
     * @param {object} layoutOptions - spacingPt / usePreviewBounds / verticalAlign を持つ設定
     * @returns {void}
     */
    function layoutHorizontally(orderedItems, layoutOptions) {
        var startBounds = getItemBounds(orderedItems[0], layoutOptions.usePreviewBounds);
        var rowTop = startBounds[1];
        var rowHeight = getLargestItemSize(orderedItems, layoutOptions.usePreviewBounds, false);
        var currentLeft = startBounds[0];

        for (var i = 0; i < orderedItems.length; i++) {
            var itemBounds = getItemBounds(orderedItems[i], layoutOptions.usePreviewBounds);

            orderedItems[i].left += currentLeft - itemBounds[0];
            if (layoutOptions.verticalAlign !== "none") {
                orderedItems[i].top += getVerticalAlignOffset(itemBounds, rowTop, rowHeight, layoutOptions.verticalAlign);
            }

            currentLeft += (itemBounds[2] - itemBounds[0]) + layoutOptions.spacingPt;
        }
    }

    /**
     * 縦一列に並べ、指定があれば左右方向にも揃える
     * @param {PageItem[]} orderedItems - 並べる順に並んだオブジェクト
     * @param {object} layoutOptions - spacingPt / usePreviewBounds / horizontalAlign を持つ設定
     * @returns {void}
     */
    function layoutVertically(orderedItems, layoutOptions) {
        var startBounds = getItemBounds(orderedItems[0], layoutOptions.usePreviewBounds);
        var columnLeft = startBounds[0];
        var columnWidth = getLargestItemSize(orderedItems, layoutOptions.usePreviewBounds, true);
        var currentTop = startBounds[1];

        for (var i = 0; i < orderedItems.length; i++) {
            var itemBounds = getItemBounds(orderedItems[i], layoutOptions.usePreviewBounds);

            if (layoutOptions.horizontalAlign !== "none") {
                orderedItems[i].left += getHorizontalAlignOffset(itemBounds, columnLeft, columnWidth, layoutOptions.horizontalAlign);
            }
            orderedItems[i].top += currentTop - itemBounds[1];

            currentTop -= (itemBounds[1] - itemBounds[3]) + layoutOptions.spacingPt;
        }
    }

    /**
     * 行の高さに対する上下揃えの移動量を返す
     * @param {number[]} itemBounds - オブジェクトの境界
     * @param {number} rowTop - 行の上端
     * @param {number} rowHeight - 行の高さ
     * @param {string} verticalAlign - "top" / "middle" / "bottom"
     * @returns {number} Y方向の移動量
     */
    function getVerticalAlignOffset(itemBounds, rowTop, rowHeight, verticalAlign) {
        if (verticalAlign === "middle") {
            return (rowTop + (rowTop - rowHeight)) / 2 - (itemBounds[1] + itemBounds[3]) / 2;
        }
        if (verticalAlign === "bottom") {
            return (rowTop - rowHeight) - itemBounds[3];
        }
        return rowTop - itemBounds[1];
    }

    /**
     * 列の幅に対する左右揃えの移動量を返す
     * @param {number[]} itemBounds - オブジェクトの境界
     * @param {number} columnLeft - 列の左端
     * @param {number} columnWidth - 列の幅
     * @param {string} horizontalAlign - "left" / "center" / "right"
     * @returns {number} X方向の移動量
     */
    function getHorizontalAlignOffset(itemBounds, columnLeft, columnWidth, horizontalAlign) {
        if (horizontalAlign === "center") {
            return (columnLeft + (columnLeft + columnWidth)) / 2 - (itemBounds[0] + itemBounds[2]) / 2;
        }
        if (horizontalAlign === "right") {
            return (columnLeft + columnWidth) - itemBounds[2];
        }
        return columnLeft - itemBounds[0];
    }

    /**
     * 設定に従って選択オブジェクトを並べ直す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} layoutOptions - direction / spacingPt / usePreviewBounds / align / randomOrder
     * @returns {void}
     */
    function applyLayout(targetItems, layoutOptions) {
        if (targetItems.length === 0) return;

        /* ランダム時も全体の左上が動かないよう、並べる前の基準を控える
           Remember the top-left before laying out so a shuffle does not drift the whole set */
        var baseLeft = null;
        var baseTop = null;
        for (var i = 0; i < targetItems.length; i++) {
            if (baseLeft === null || targetItems[i].left < baseLeft) baseLeft = targetItems[i].left;
            if (baseTop === null || targetItems[i].top > baseTop) baseTop = targetItems[i].top;
        }

        var orderedItems = getLayoutOrder(targetItems, layoutOptions);
        if (layoutOptions.direction === "horizontal") {
            layoutHorizontally(orderedItems, layoutOptions);
        } else {
            layoutVertically(orderedItems, layoutOptions);
        }

        if (!layoutOptions.randomOrder) return;

        /* シャッフルで先頭が入れ替わった分を、全体でまとめて戻す / shift the whole set back */
        var offsetX = baseLeft - orderedItems[0].left;
        var offsetY = baseTop - orderedItems[0].top;
        for (var j = 0; j < orderedItems.length; j++) {
            orderedItems[j].left += offsetX;
            orderedItems[j].top += offsetY;
        }
    }

    // =========================================
    // 入力欄 / Numeric input
    // =========================================

    /**
     * ↑↓キーで数値欄の値を増減する（Shiftで10刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すなら true
     * @param {function} onValueChanged - 値が変わったときに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            if (editText.text.length === 0) return;

            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var isUp = (event.keyName === "Up");
            var step = 1;
            if (ScriptUI.environment.keyboardState.shiftKey) {
                /* 10の倍数にそろえてから10刻み / snap to a multiple of 10 first */
                currentValue = Math.floor(currentValue / 10) * 10;
                step = 10;
            }

            var newValue = currentValue + (isUp ? step : -step);
            if (!allowNegative && newValue < 0) newValue = 0;

            editText.text = newValue;
            event.preventDefault();

            if (onValueChanged) onValueChanged();
        });
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログ位置をセッション内で覚えておく
     * @param {number[]} dialogLocation - ダイアログの位置 [x, y]
     * @returns {void}
     */
    function saveDialogPosition(dialogLocation) {
        if (!dialogLocation || dialogLocation.length !== 2) return;
        $.global[DIALOG_POSITION_KEY] = [Math.round(dialogLocation[0]), Math.round(dialogLocation[1])];
    }

    /**
     * 覚えておいたダイアログ位置を返す
     * @returns {number[]|null} ダイアログの位置。記憶がなければ null
     */
    function loadDialogPosition() {
        var savedPosition = $.global[DIALOG_POSITION_KEY];
        return (savedPosition && savedPosition.length === 2) ? [savedPosition[0], savedPosition[1]] : null;
    }

    /**
     * 揃えのラジオボタンを1行ぶん追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string[]} labelPaths - ラジオのラベルキー（先頭が既定で選択される）
     * @param {string} tooltipPath - ツールチップのキー
     * @param {function} onValueChanged - 選択が変わったときに呼ぶ処理
     * @returns {Group} 追加した行グループ（children がラジオ）
     */
    function addAlignRadioRow(parentPanel, labelPaths, tooltipPath, onValueChanged) {
        var alignRowGroup = parentPanel.add("group");
        alignRowGroup.orientation = "row";
        alignRowGroup.alignChildren = ["left", "center"];

        for (var i = 0; i < labelPaths.length; i++) {
            var alignRadio = alignRowGroup.add("radiobutton", undefined, getLabel(labelPaths[i]));
            alignRadio.helpTip = getLabel(tooltipPath);
            alignRadio.value = (i === 0);
            alignRadio.onClick = onValueChanged;
        }
        return alignRowGroup;
    }

    /**
     * 行グループの中で選択されているラジオの添字を返す
     * @param {Group} radioRowGroup - ラジオを並べたグループ
     * @returns {number} 選択されている添字（無ければ 0）
     */
    function getSelectedRadioIndex(radioRowGroup) {
        for (var i = 0; i < radioRowGroup.children.length; i++) {
            if (radioRowGroup.children[i].value) return i;
        }
        return 0;
    }

    /**
     * 行グループのラジオを添字で選び直す
     * @param {Group} radioRowGroup - ラジオを並べたグループ
     * @param {number} selectedIndex - 選択する添字
     * @returns {void}
     */
    function selectRadioAt(radioRowGroup, selectedIndex) {
        for (var i = 0; i < radioRowGroup.children.length; i++) {
            radioRowGroup.children[i].value = (i === selectedIndex);
        }
    }

    /* 揃えのショートカットキーと、ラジオの添字（なし / 1 / 2 / 3）
       Shortcut keys mapped to the radio index in each align row */
    var HORIZONTAL_ALIGN_INDEX_BY_KEY = { "N": 0, "L": 1, "C": 2, "R": 3 };
    var VERTICAL_ALIGN_INDEX_BY_KEY = { "N": 0, "T": 1, "M": 2, "B": 3 };

    /* ラジオの添字に対応する揃えの指定 / Align value per radio index */
    var HORIZONTAL_ALIGN_VALUES = ["none", "left", "center", "right"];
    var VERTICAL_ALIGN_VALUES = ["none", "top", "middle", "bottom"];

    /**
     * 揃えのショートカットキーを登録する
     * @param {object} keyTarget - キーを受けるコントロール
     * @param {object} alignRows - horizontal / vertical の行グループと、現在の向きを返す getDirection
     * @param {function} onValueChanged - 切り替えたときに呼ぶ処理
     * @returns {void}
     */
    function addAlignKeyHandler(keyTarget, alignRows, onValueChanged) {
        keyTarget.addEventListener("keydown", function (event) {
            var isHorizontalLayout = (alignRows.getDirection() === "horizontal");
            /* 横並びのときは上下の揃え、縦並びのときは左右の揃えを操作する
               A horizontal layout adjusts the vertical align, and vice versa */
            var indexByKey = isHorizontalLayout ? VERTICAL_ALIGN_INDEX_BY_KEY : HORIZONTAL_ALIGN_INDEX_BY_KEY;
            var targetRow = isHorizontalLayout ? alignRows.vertical : alignRows.horizontal;

            var alignIndex = indexByKey[event.keyName];
            if (alignIndex === undefined) return;

            selectRadioAt(targetRow, alignIndex);
            event.preventDefault();
            if (onValueChanged) onValueChanged();
        });
    }

    /**
     * 整列と分布のダイアログを表示し、プレビューしながら設定させる
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {void}
     */
    function showArrangeDialog(targetItems) {
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");
        var positionSnapshot = createPositionSnapshot(targetItems);
        var detectedDirection = detectDirection(targetItems);

        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = "fill";
        alignDialog.opacity = DIALOG_OPACITY;

        var savedPosition = loadDialogPosition();
        if (savedPosition) alignDialog.location = savedPosition;

        /* 方向 / Direction */
        var directionPanel = alignDialog.add("panel", undefined, getLabel("panel.direction"));
        directionPanel.orientation = "row";
        directionPanel.alignChildren = ["center", "center"];
        directionPanel.margins = PANEL_MARGINS;

        var directionAutoRadio = directionPanel.add("radiobutton", undefined, getLabel("radio.directionAuto"));
        directionAutoRadio.helpTip = getLabel("tooltip.directionAuto");
        var directionVerticalRadio = directionPanel.add("radiobutton", undefined, getLabel("radio.directionVertical"));
        directionVerticalRadio.helpTip = getLabel("tooltip.directionVertical");
        var directionHorizontalRadio = directionPanel.add("radiobutton", undefined, getLabel("radio.directionHorizontal"));
        directionHorizontalRadio.helpTip = getLabel("tooltip.directionHorizontal");
        directionAutoRadio.value = true;

        /**
         * 現在選んでいる並べる向きを返す
         * @returns {string} "horizontal" または "vertical"
         */
        function getEffectiveDirection() {
            if (directionVerticalRadio.value) return "vertical";
            if (directionHorizontalRadio.value) return "horizontal";
            return detectedDirection;
        }

        /* 間隔 / Spacing */
        var spacingPanel = alignDialog.add("panel", undefined, getLabel("panel.spacing"));
        spacingPanel.orientation = "column";
        spacingPanel.alignChildren = ["center", "center"];
        spacingPanel.margins = PANEL_MARGINS;

        var spacingRowGroup = spacingPanel.add("group");
        spacingRowGroup.orientation = "row";
        spacingRowGroup.alignChildren = ["left", "center"];

        var spacingInput = spacingRowGroup.add("edittext", undefined, DEFAULT_SPACING);
        spacingInput.helpTip = getLabel("tooltip.spacing");
        spacingInput.characters = SPACING_INPUT_CHARS;
        spacingRowGroup.add("statictext", undefined, getUnitInfo().label);

        /* 揃え / Align */
        var alignPanel = alignDialog.add("panel", undefined, "");
        alignPanel.orientation = "column";
        alignPanel.alignChildren = ["left", "center"];
        alignPanel.margins = PANEL_MARGINS;

        var horizontalAlignRow = addAlignRadioRow(alignPanel,
            ["radio.alignNone", "radio.alignLeft", "radio.alignCenter", "radio.alignRight"],
            "tooltip.alignHorizontal", function () { requestPreviewUpdate(); });
        var verticalAlignRow = addAlignRadioRow(alignPanel,
            ["radio.alignNone", "radio.alignTop", "radio.alignMiddle", "radio.alignBottom"],
            "tooltip.alignVertical", function () { requestPreviewUpdate(); });

        /* オプション / Options */
        var optionGroup = alignDialog.add("group");
        optionGroup.orientation = "column";
        optionGroup.alignChildren = ["left", "center"];
        optionGroup.alignment = ["fill", "top"];
        optionGroup.margins = OPTION_GROUP_MARGINS;

        var previewBoundsCheckbox = optionGroup.add("checkbox", undefined, getLabel("checkbox.usePreviewBounds"));
        previewBoundsCheckbox.helpTip = getLabel("tooltip.usePreviewBounds");
        previewBoundsCheckbox.value = DEFAULT_USE_PREVIEW_BOUNDS;
        previewBoundsCheckbox.onClick = function () { requestPreviewUpdate(); };

        var randomizeCheckbox = optionGroup.add("checkbox", undefined, getLabel("checkbox.randomizeOrder"));
        randomizeCheckbox.helpTip = getLabel("tooltip.randomizeOrder");
        randomizeCheckbox.value = DEFAULT_RANDOMIZE_ORDER;

        /* ボタンエリア / Button row */
        var btnRowGroup = alignDialog.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.alignChildren = ["center", "center"];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /* シャッフル順は覚えておく（プレビューとOKの結果を一致させるため）
           The shuffle order is cached so OK produces exactly what the preview showed */
        var cachedRandomOrder = null;

        /**
         * 並べる向きに応じて、使わない側の揃えをディムし、パネル名を合わせる
         * @returns {void}
         */
        function syncAlignPanel() {
            var isHorizontalLayout = (getEffectiveDirection() === "horizontal");
            alignPanel.text = getLabel(isHorizontalLayout ? "panel.alignVertical" : "panel.alignHorizontal");
            horizontalAlignRow.enabled = !isHorizontalLayout;
            verticalAlignRow.enabled = isHorizontalLayout;
            alignDialog.layout.layout(true);
        }

        /**
         * ダイアログの現在値を並べ直しの設定にまとめる
         * @returns {object} applyLayout に渡す設定
         */
        function getLayoutOptions() {
            var spacingValue = parseFloat(spacingInput.text);
            if (isNaN(spacingValue)) spacingValue = 0;

            if (randomizeCheckbox.value) {
                if (!cachedRandomOrder || cachedRandomOrder.length !== targetItems.length) {
                    cachedRandomOrder = createShuffledOrder(targetItems.length);
                }
            } else {
                cachedRandomOrder = null;
            }

            return {
                direction: getEffectiveDirection(),
                spacingPt: spacingValue * getUnitInfo().pointsPerUnit,
                usePreviewBounds: previewBoundsCheckbox.value,
                horizontalAlign: HORIZONTAL_ALIGN_VALUES[getSelectedRadioIndex(horizontalAlignRow)],
                verticalAlign: VERTICAL_ALIGN_VALUES[getSelectedRadioIndex(verticalAlignRow)],
                randomOrder: cachedRandomOrder
            };
        }

        var isPreviewRunning = false;
        var lastPreviewTime = 0;
        var scheduledPreviewTaskId = 0;

        /**
         * 元の位置へ戻してから、現在の設定でプレビューを描き直す
         * @returns {void}
         */
        function updatePreview() {
            if (isPreviewRunning) return;

            var nowMs = new Date().getTime();
            if (nowMs - lastPreviewTime < PREVIEW_MIN_INTERVAL_MS) return;
            lastPreviewTime = nowMs;

            isPreviewRunning = true;
            try {
                positionSnapshot.restore();
                /* 境界の計算に使われる環境設定を、チェックボックスに合わせてから並べる
                   The bounds calculation follows this preference, so set it before laying out */
                app.preferences.setBooleanPreference("includeStrokeInBounds", previewBoundsCheckbox.value);
                applyLayout(targetItems, getLayoutOptions());
                app.redraw();
            } finally {
                isPreviewRunning = false;
            }
        }

        $.global[PREVIEW_CALLBACK_KEY] = updatePreview;

        /**
         * プレビューを少し遅らせて実行する（ScriptUI のイベント内で走らせない）
         * @returns {void}
         */
        function requestPreviewUpdate() {
            cancelScheduledPreview();
            $.global[PREVIEW_CALLBACK_KEY] = updatePreview;
            try {
                scheduledPreviewTaskId = app.scheduleTask(
                    '$.global["' + PREVIEW_CALLBACK_KEY + '"] && $.global["' + PREVIEW_CALLBACK_KEY + '"]();',
                    PREVIEW_DELAY_MS, false);
            } catch (e) {
                /* スケジュールできない環境ではその場で実行する / run inline when scheduling is unavailable */
                updatePreview();
            }
        }

        /**
         * 予約したプレビューを取り消す
         * @returns {void}
         */
        function cancelScheduledPreview() {
            if (!scheduledPreviewTaskId) return;
            /* すでに実行済みのIDを渡すと例外になるため受け流す / an already-run task id throws */
            try {
                app.cancelTask(scheduledPreviewTaskId);
            } catch (e) {}
            scheduledPreviewTaskId = 0;
        }

        var alignRows = {
            horizontal: horizontalAlignRow,
            vertical: verticalAlignRow,
            getDirection: getEffectiveDirection
        };
        addAlignKeyHandler(alignDialog, alignRows, requestPreviewUpdate);
        addAlignKeyHandler(spacingInput, alignRows, requestPreviewUpdate);
        changeValueByArrowKey(spacingInput, true, requestPreviewUpdate);

        directionAutoRadio.onClick = function () { syncAlignPanel(); requestPreviewUpdate(); };
        directionVerticalRadio.onClick = function () { syncAlignPanel(); requestPreviewUpdate(); };
        directionHorizontalRadio.onClick = function () { syncAlignPanel(); requestPreviewUpdate(); };
        randomizeCheckbox.onClick = function () {
            cachedRandomOrder = null;
            requestPreviewUpdate();
        };

        syncAlignPanel();
        requestPreviewUpdate();

        spacingInput.active = true;
        var dialogResult = alignDialog.show();

        cancelScheduledPreview();
        $.global[PREVIEW_CALLBACK_KEY] = null;
        saveDialogPosition(alignDialog.location);

        if (dialogResult !== 1) {
            /* キャンセル：位置も環境設定も元へ戻す / Cancel restores both the positions and the preference */
            positionSnapshot.restore();
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return;
        }

        /* OK：プレビューで適用済みの位置をそのまま確定する / OK keeps what the preview already applied */
        app.preferences.setBooleanPreference("includeStrokeInBounds", previewBoundsCheckbox.value);
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを整列・分布するダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var selectedObjects = app.activeDocument.selection;
        if (!selectedObjects || selectedObjects.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        showArrangeDialog(selectedObjects.slice(0));
    }

    main();

})();
