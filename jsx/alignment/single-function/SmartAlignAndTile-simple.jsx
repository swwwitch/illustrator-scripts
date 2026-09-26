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
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                             /* 更新日 / last updated */

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
                ja: "隣り合うオブジェクトのあいだにあける間隔です。∧∨や↑↓キーで増減できます（Shiftで10の倍数へ）。",
                en: "Gap left between neighbouring objects. The stepper and the arrow keys step the value (Shift to multiples of 10)."
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
            },
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkStepperUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var STEPPER_UI_DARK           = isDarkStepperUI();
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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
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
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp";
        var downTooltip = stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown";
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var spacingStepperGroup = spacingRowGroup.add("group");
        spacingStepperGroup.orientation = "row";
        spacingStepperGroup.alignChildren = ["left", "center"];
        spacingStepperGroup.spacing = 0;
        spacingStepperGroup.margins = 0;
        /* 間隔は負の値も許す（重ねて並べる）/ negative spacing is allowed (overlap) */
        var spacingStepper = addStepper(spacingStepperGroup, function () { return spacingInput; }, {
            step: 1,
            onStep: function () { requestPreviewUpdate(); }
        });
        var spacingInput = spacingStepperGroup.add("edittext", undefined, DEFAULT_SPACING);
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
        bindSteppedArrowKeys(spacingInput, spacingStepper);

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
