#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを重なりや距離、縦の列・横の行でまとまりに分け、まとまりごとにグループ化するか、クリッピングマスクを作成します（配置画像を1つずつクリップすることもできます）。
［OK］の前に、できるグループの数と範囲を件数表示とプレビューの枠で確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartClipAndGroup.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb23985473f80

### Overview

Splits the selection into clusters by overlap, distance, column or row, and groups or clips each cluster (placed images can also be clipped one by one).
Before you click OK, a count and preview frames show the groups to be created.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartClipAndGroup.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartClipAndGroup";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartClipAndGroup.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartClipAndGroup.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb23985473f80"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* しきい値スライダーの初期値と範囲（pt） / Default and range of the threshold slider (pt) */
    var DEFAULT_THRESHOLD = 10;
    var THRESHOLD_MIN     = 0;
    var THRESHOLD_MAX     = 100;

    /* 「間隔が空いたら分ける」の初期値 / Default of "Split at gaps" */
    var DEFAULT_SPLIT_AT_GAPS = false;

    /* 「プレビューを表示」の初期値 / Default of "Show preview" */
    var DEFAULT_SHOW_PREVIEW = true;

    /* 「重なり」や、列・行で揃っているとみなす距離（pt） / Gap treated as touching or aligned (pt) */
    var TOUCH_TOLERANCE = 0.01;

    /* スライダーのドラッグ中も件数とプレビューを更新する選択数の上限（超えると指を離したときだけ更新する）
       Max selection size for live updates while dragging; above it, the count and preview update on release only */
    var LIVE_COUNT_MAX_ITEMS = 300;

    /* プレビューの枠（赤・塗りなし・線 10 pt・不透明度 50%） / Preview frames: red, no fill, 10 pt stroke, 50% opacity */
    var PREVIEW_RGB          = [255, 0, 0];     /* RGB ドキュメントの線の色 / Stroke color in RGB documents */
    var PREVIEW_CMYK         = [0, 100, 100, 0]; /* CMYK ドキュメントの線の色 / Stroke color in CMYK documents */
    var PREVIEW_STROKE_WIDTH = 10;              /* 線幅（pt） / Stroke width (pt) */
    var PREVIEW_OPACITY      = 50;              /* 不透明度（%） / Opacity (%) */
    var PREVIEW_LAYER_NAME   = "SmartClipAndGroup Preview"; /* プレビュー用の一時レイヤー名 / Temporary preview layer name */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS       = [25, 20, 25, 20];  /* ダイアログ余白 [左,上,右,下] / Dialog margins */
    var PANEL_MARGINS        = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / Panel margins */
    var RADIO_COLUMN_SPACING = 20;                /* ラジオボタンの列の間隔 / Gap between radio button columns */
    var SLIDER_WIDTH         = 150;               /* しきい値スライダーの幅 / Threshold slider width */
    var VALUE_TEXT_CHARS     = 5;                 /* しきい値表示の文字数 / Width of the threshold readout in characters */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var uiLang = getCurrentLang();

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        dialog: {
            title: { ja: "クリップとグループ化", en: "Clip and Group" }
        },
        panel: {
            clip: { ja: "クリッピングマスク", en: "Clipping Mask" },
            grouping: { ja: "グループ化", en: "Grouping" }
        },
        radio: {
            clipWithFront: { ja: "最前面でクリップ", en: "Clip with Frontmost" },
            clipWithBack: { ja: "最背面でクリップ", en: "Clip with Backmost" },
            clipPlacedOnly: { ja: "配置画像のみをクリップ", en: "Clip Placed Images Only" },
            groupOverlap: { ja: "重なり", en: "Overlap" },
            groupProximity: { ja: "近接度", en: "Proximity" },
            groupVertical: { ja: "上下方向", en: "Vertical" },
            groupHorizontal: { ja: "左右方向", en: "Horizontal" }
        },
        fieldLabel: {
            threshold: { ja: "しきい値", en: "Threshold" },
            selectedCount: { ja: "選択しているオブジェクト", en: "Selected objects" },
            groupCount: { ja: "グループ化", en: "Groups" },
            clipCount: { ja: "クリップ", en: "Clips" }
        },
        checkbox: {
            splitAtGaps: { ja: "間隔が空いたら分ける", en: "Split at gaps" },
            perArtboard: { ja: "アートボードごと", en: "Within each artboard" },
            showPreview: { ja: "プレビューを表示", en: "Show preview" }
        },
        tooltip: {
            clipWithFront: {
                ja: "重なっているまとまりごとに、最前面のパスを型にしてクリップします。",
                en: "Clips each cluster of overlapping objects, using its frontmost path as the mask."
            },
            clipWithBack: {
                ja: "重なっているまとまりごとに、最背面のパスを型にしてクリップします。",
                en: "Clips each cluster of overlapping objects, using its backmost path as the mask."
            },
            clipPlacedOnly: {
                ja: "画像を1つずつ、同じ大きさの長方形でクリップします。クリップグループを選ぶと、元のマスクだけを削除して中の画像をクリップし直します。",
                en: "Clips each image with a rectangle of the same size. For a clipping group, only the old mask is removed and the images inside are clipped again."
            },
            groupOverlap: {
                ja: "外枠が重なっているか接しているオブジェクトをグループにします。",
                en: "Groups objects whose bounding boxes overlap or touch."
            },
            groupProximity: {
                ja: "しきい値以内の距離にあるオブジェクトを、向きを問わずグループにします。",
                en: "Groups objects that sit within the threshold distance, in any direction."
            },
            groupVertical: {
                ja: "上下に並んでいるオブジェクトを、縦の列ごとにグループにします。左右のすき間がしきい値以内なら同じ列とみなします。",
                en: "Groups objects that line up vertically, column by column. Objects whose horizontal gap is within the threshold count as the same column."
            },
            groupHorizontal: {
                ja: "左右に並んでいるオブジェクトを、横の行ごとにグループにします。上下のすき間がしきい値以内なら同じ行とみなします。",
                en: "Groups objects that line up horizontally, row by row. Objects whose vertical gap is within the threshold count as the same row."
            },
            threshold: {
                ja: "同じグループとみなす距離です。大きくするとまとまりが粗くなります。",
                en: "How close objects must be to land in the same group. Larger values group more loosely."
            },
            splitAtGaps: {
                ja: "上下方向・左右方向で使います。ONにすると、列（行）の途中ですき間がしきい値より空いたところでグループを分けます。つなぐのは揃って並んでいるもの（列なら左右が重なるもの）だけです。OFFのときは距離を問わず、列（行）全体をまとめます。",
                en: "Used with Vertical and Horizontal. When on, a column (row) is split wherever the gap exceeds the threshold, and only aligned objects (horizontally overlapping, for columns) are linked. When off, whole columns (rows) are grouped regardless of distance."
            },
            perArtboard: {
                ja: "ONにすると、別のアートボードにあるオブジェクトは別々のグループにします。オブジェクトの中心がどのアートボードに含まれるかで判定します。",
                en: "When on, objects on different artboards go into separate groups. The artboard is determined by each object's center point."
            },
            showPreview: {
                ja: "ONにすると、作成されるグループの範囲を赤い枠で表示します。枠はダイアログを閉じると消えます。",
                en: "When on, red frames show where the groups will be created. The frames disappear when the dialog closes."
            },
            resultCount: {
                ja: "［OK］で作成されるグループの数です。クリッピングマスクのモードでは、作成されるクリップグループの数です。",
                en: "The number of groups OK will create. In the clipping mask modes, the number of clipping groups."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} コロン付きのラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // モード / Modes
    // =========================================

    /* パネルごとのモードの並び（キーは LABELS.radio / LABELS.tooltip と共通）
       Mode order per panel; keys are shared with LABELS.radio and LABELS.tooltip */
    var CLIP_MODES  = ["clipWithFront", "clipWithBack", "clipPlacedOnly"];
    var GROUP_MODES = ["groupOverlap", "groupProximity", "groupVertical", "groupHorizontal"];
    var ALL_MODES   = CLIP_MODES.concat(GROUP_MODES);

    /* グループ化のモードごとに使う設定（クリップ系はどれも使わない）
       Settings each grouping mode uses; the clipping modes use none */
    var MODE_OPTIONS = {
        groupOverlap: { perArtboard: true },
        groupProximity: { threshold: true, perArtboard: true },
        groupVertical: { threshold: true, splitAtGaps: true, perArtboard: true },
        groupHorizontal: { threshold: true, splitAtGaps: true, perArtboard: true }
    };

    /* まとまりをクリップするモードと、型にするパスの側 / Cluster-clipping modes and which side supplies the mask */
    var MASK_SIDES = { clipWithFront: "front", clipWithBack: "back" };

    /* すき間の判定を通さず、面積を持つ重なりだけでつなぐ（すき間は0以上なので常に不成立）
       Link by area overlap only; gaps are never negative, so the gap test always fails */
    var OVERLAP_ONLY = -1;

    // =========================================
    // 選択 / Selection
    // =========================================

    /**
     * 処理対象の選択オブジェクトを返す
     * @returns {PageItem[]|null} 選択オブジェクト。ドキュメントがない・未選択・文字の選択中は null
     */
    function getValidSelection() {
        if (!app.documents.length) return null;
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る
           A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || !selectedItems.length || selectedItems.typename === "TextRange") return null;
        return selectedItems;
    }

    /**
     * 配置画像または埋め込み画像かどうかを返す
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @returns {boolean} PlacedItem か RasterItem なら true
     */
    function isImageItem(targetItem) {
        return targetItem.typename === "PlacedItem" || targetItem.typename === "RasterItem";
    }

    /**
     * すべてが画像かどうかを返す
     * @param {PageItem[]|null} targetItems - 判定するオブジェクト
     * @returns {boolean} 1つ以上あり、すべてが画像なら true
     */
    function isAllImages(targetItems) {
        if (!targetItems || !targetItems.length) return false;
        for (var i = 0; i < targetItems.length; i++) {
            if (!isImageItem(targetItems[i])) return false;
        }
        return true;
    }

    /**
     * 選択を解除して、指定のオブジェクトを選択する
     * @param {PageItem[]} itemsToSelect - 選択するオブジェクト
     * @returns {void}
     */
    function selectOnly(itemsToSelect) {
        app.activeDocument.selection = null;
        for (var i = 0; i < itemsToSelect.length; i++) {
            itemsToSelect[i].selected = true;
        }
    }

    // =========================================
    // まとまりの判定 / Cluster detection
    // =========================================

    /**
     * @typedef {Object} ClusterRule まとまりのつなぎ方
     * @property {number} maxGap - つなぐ最大のすき間（pt）。OVERLAP_ONLY なら重なりだけでつなぐ
     * @property {string} [direction] - "vertical" なら縦の列、"horizontal" なら横の行で判定
     * @property {boolean} [splitAtGaps] - 列・行の途中で、すき間がしきい値より空いたところで分ける（揃っているものだけをつなぐ）
     * @property {boolean} [perArtboard] - 別のアートボードにあるものはつながない
     */

    /**
     * 2つの矩形が面積を持って重なっているかを返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @returns {boolean} 重なっていれば true
     */
    function boundsOverlap(boundsA, boundsB) {
        var overlapWidth = Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]);
        var overlapHeight = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
        return overlapWidth > 0 && overlapHeight > 0;
    }

    /**
     * 2つの矩形の左右のすき間を返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @returns {number} すき間（pt）。左右の範囲が重なっていれば 0
     */
    function getHorizontalGap(boundsA, boundsB) {
        return Math.max(0, boundsB[0] - boundsA[2], boundsA[0] - boundsB[2]);
    }

    /**
     * 2つの矩形の上下のすき間を返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @returns {number} すき間（pt）。上下の範囲が重なっていれば 0
     */
    function getVerticalGap(boundsA, boundsB) {
        return Math.max(0, boundsA[3] - boundsB[1], boundsB[3] - boundsA[1]);
    }

    /**
     * 列・行の判定で2つの矩形をつなぐかどうかを返す
     * @param {number} crossGap - 並ぶ向きと直交する向きのすき間（列なら左右）
     * @param {number} alongGap - 並ぶ向きのすき間（列なら上下）
     * @param {ClusterRule} clusterRule - つなぎ方
     * @returns {boolean} つなぐなら true
     */
    function areInSameLine(crossGap, alongGap, clusterRule) {
        /* 間隔が空いたら分ける：揃っていて、並ぶ向きのすき間がしきい値以内 / Split at gaps: aligned, and the gap along the line is within the threshold */
        if (clusterRule.splitAtGaps) return crossGap <= TOUCH_TOLERANCE && alongGap <= clusterRule.maxGap;
        /* 距離は問わない：直交する向きのずれがしきい値以内 / Any distance along the line; the cross offset is within the threshold */
        return crossGap <= clusterRule.maxGap;
    }

    /**
     * 2つの矩形を同じまとまりとしてつなぐかどうかを返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @param {ClusterRule} clusterRule - つなぎ方
     * @returns {boolean} つなぐなら true
     */
    function areNeighbors(boundsA, boundsB, clusterRule) {
        var horizontalGap = getHorizontalGap(boundsA, boundsB);
        var verticalGap = getVerticalGap(boundsA, boundsB);
        if (clusterRule.direction === "vertical") return areInSameLine(horizontalGap, verticalGap, clusterRule);
        if (clusterRule.direction === "horizontal") return areInSameLine(verticalGap, horizontalGap, clusterRule);
        return boundsOverlap(boundsA, boundsB) || Math.max(horizontalGap, verticalGap) <= clusterRule.maxGap;
    }

    /**
     * 起点からつながっている矩形の番号をすべて集める（visited を更新する）
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @param {number} startIndex - 起点の番号
     * @param {boolean[]} visited - 集め済みの印
     * @param {ClusterRule} clusterRule - つなぎ方
     * @param {number[]|null} artboardIndexes - 各オブジェクトのアートボード番号（アートボードごとに分けないときは null）
     * @returns {number[]} つながっている番号（昇順）
     */
    function collectConnectedIndexes(boundsList, startIndex, visited, clusterRule, artboardIndexes) {
        var connectedIndexes = [];
        var pendingIndexes = [startIndex];
        visited[startIndex] = true;
        /* 再帰を使わず積み残しを順に処理する（大きなまとまりでスタックがあふれないように）
           Walk a work list instead of recursing so large clusters cannot overflow the stack */
        while (pendingIndexes.length) {
            var currentIndex = pendingIndexes.pop();
            connectedIndexes.push(currentIndex);
            for (var j = 0; j < boundsList.length; j++) {
                if (visited[j]) continue;
                if (artboardIndexes && artboardIndexes[j] !== artboardIndexes[currentIndex]) continue;
                if (!areNeighbors(boundsList[currentIndex], boundsList[j], clusterRule)) continue;
                visited[j] = true;
                pendingIndexes.push(j);
            }
        }
        connectedIndexes.sort(function (a, b) { return a - b; });
        return connectedIndexes;
    }

    /**
     * 矩形の並びを、重なりや距離でつながるまとまりに分ける
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @param {ClusterRule} clusterRule - つなぎ方
     * @param {number[]|null} artboardIndexes - 各オブジェクトのアートボード番号（アートボードごとに分けないときは null）
     * @returns {Array<number[]>} まとまりごとの番号の配列（各まとまりは昇順）
     */
    function collectClusterIndexes(boundsList, clusterRule, artboardIndexes) {
        var visited = [];
        for (var i = 0; i < boundsList.length; i++) {
            visited.push(false);
        }
        var clusterIndexes = [];
        for (var j = 0; j < boundsList.length; j++) {
            if (!visited[j]) clusterIndexes.push(collectConnectedIndexes(boundsList, j, visited, clusterRule, artboardIndexes));
        }
        return clusterIndexes;
    }

    /**
     * モードと設定に応じた、まとまりのつなぎ方を返す
     * @param {string} mode - モードのキー（「配置画像のみをクリップ」以外）
     * @param {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} groupingOptions - ダイアログの設定
     * @returns {ClusterRule|null} つなぎ方。該当しないモードは null
     */
    function getClusterRule(mode, groupingOptions) {
        var usedOptions = MODE_OPTIONS[mode] || {};
        var perArtboard = usedOptions.perArtboard === true && groupingOptions.perArtboard === true;
        switch (mode) {
            case "clipWithFront":
            case "clipWithBack":
                return { maxGap: OVERLAP_ONLY };
            case "groupOverlap":
                return { maxGap: TOUCH_TOLERANCE, perArtboard: perArtboard };
            case "groupProximity":
                return { maxGap: groupingOptions.threshold, perArtboard: perArtboard };
            case "groupVertical":
            case "groupHorizontal":
                return {
                    maxGap: groupingOptions.threshold,
                    direction: (mode === "groupVertical") ? "vertical" : "horizontal",
                    splitAtGaps: groupingOptions.splitAtGaps === true,
                    perArtboard: perArtboard
                };
        }
        return null;
    }

    // =========================================
    // アートボード / Artboards
    // =========================================

    /**
     * 矩形の中心が含まれるアートボードの番号を返す
     * @param {number[]} bounds - geometricBounds [左, 上, 右, 下]
     * @param {Array<number[]>} artboardRects - 各アートボードの artboardRect
     * @returns {number} アートボードの番号。どこにも含まれなければ -1
     */
    function findArtboardIndex(bounds, artboardRects) {
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        for (var i = 0; i < artboardRects.length; i++) {
            var artboardRect = artboardRects[i];
            if (centerX >= artboardRect[0] && centerX <= artboardRect[2] && centerY <= artboardRect[1] && centerY >= artboardRect[3]) return i;
        }
        return -1;
    }

    /**
     * 各矩形が属するアートボードの番号を返す
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @returns {number[]} アートボードの番号（boundsList と同じ並び）
     */
    function getArtboardIndexes(boundsList) {
        var artboards = app.activeDocument.artboards;
        var artboardRects = [];
        for (var i = 0; i < artboards.length; i++) {
            artboardRects.push(artboards[i].artboardRect);
        }
        var artboardIndexes = [];
        for (var j = 0; j < boundsList.length; j++) {
            artboardIndexes.push(findArtboardIndex(boundsList[j], artboardRects));
        }
        return artboardIndexes;
    }

    /**
     * 2つ以上のアートボードにまたがっているかを返す（どのアートボードにもないものは数えない）
     * @param {number[]} artboardIndexes - 各オブジェクトのアートボード番号
     * @returns {boolean} またがっていれば true
     */
    function spansMultipleArtboards(artboardIndexes) {
        var firstIndex = -1;
        for (var i = 0; i < artboardIndexes.length; i++) {
            if (artboardIndexes[i] < 0) continue;
            if (firstIndex < 0) {
                firstIndex = artboardIndexes[i];
            } else if (artboardIndexes[i] !== firstIndex) {
                return true;
            }
        }
        return false;
    }

    // =========================================
    // 選択の読み取りと結果の見積もり / Selection snapshot and result plan
    // =========================================

    /**
     * @typedef {Object} SelectionInfo ダイアログを開く前に読み取った選択の情報
     * @property {PageItem[]} items - 選択オブジェクト（前面から背面の順）
     * @property {number} itemCount - 選択オブジェクトの数
     * @property {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @property {boolean[]} isPath - 各オブジェクトがパスかどうか
     * @property {Array<number[]>} imageBoundsList - 「配置画像のみをクリップ」の対象になる画像の visibleBounds
     * @property {number[]} artboardIndexes - 各オブジェクトのアートボード番号
     * @property {boolean} hasMultipleArtboards - ドキュメントにアートボードが2つ以上あるか
     */

    /**
     * 選択の情報を一度だけ読み取る（件数・プレビュー・実行で共有する）
     * @param {PageItem[]|null} selectedItems - 選択オブジェクト
     * @returns {SelectionInfo} 選択の情報
     */
    function readSelection(selectedItems) {
        var selectionInfo = {
            items: selectedItems || [],
            itemCount: 0,
            boundsList: [],
            isPath: [],
            imageBoundsList: [],
            artboardIndexes: [],
            hasMultipleArtboards: false
        };
        if (!selectedItems) return selectionInfo;

        selectionInfo.itemCount = selectedItems.length;
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            selectionInfo.boundsList.push(selectedItem.geometricBounds);
            selectionInfo.isPath.push(selectedItem.typename === "PathItem");
            /* 「配置画像のみをクリップ」の対象：選択中の画像と、クリップグループ直下の画像
               Targets of Clip Placed Images Only: selected images and images directly inside clipping groups */
            if (isImageItem(selectedItem)) {
                selectionInfo.imageBoundsList.push(selectedItem.visibleBounds);
            } else if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
                var childImages = getChildImages(selectedItem);
                for (var j = 0; j < childImages.length; j++) {
                    selectionInfo.imageBoundsList.push(childImages[j].visibleBounds);
                }
            }
        }
        selectionInfo.artboardIndexes = getArtboardIndexes(selectionInfo.boundsList);
        selectionInfo.hasMultipleArtboards = app.activeDocument.artboards.length > 1;
        return selectionInfo;
    }

    /**
     * 選択を、つなぎ方に従ってまとまりの番号に分ける
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @param {ClusterRule} clusterRule - つなぎ方
     * @returns {Array<number[]>} まとまりごとの番号の配列
     */
    function collectSelectionClusters(selectionInfo, clusterRule) {
        var artboardIndexes = clusterRule.perArtboard ? selectionInfo.artboardIndexes : null;
        return collectClusterIndexes(selectionInfo.boundsList, clusterRule, artboardIndexes);
    }

    /**
     * まとまりにパスが含まれるかを返す
     * @param {boolean[]} isPath - 選択の各オブジェクトがパスかどうか
     * @param {number[]} memberIndexes - まとまりの番号
     * @returns {boolean} パスがあれば true
     */
    function hasPathMember(isPath, memberIndexes) {
        for (var i = 0; i < memberIndexes.length; i++) {
            if (isPath[memberIndexes[i]]) return true;
        }
        return false;
    }

    /**
     * まとまりの外枠（すべてを囲む矩形）を返す
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @param {number[]} memberIndexes - まとまりの番号
     * @returns {number[]} 外枠 [左, 上, 右, 下]
     */
    function getUnionBounds(boundsList, memberIndexes) {
        var unionBounds = boundsList[memberIndexes[0]].slice(0);
        for (var i = 1; i < memberIndexes.length; i++) {
            var bounds = boundsList[memberIndexes[i]];
            unionBounds[0] = Math.min(unionBounds[0], bounds[0]);
            unionBounds[1] = Math.max(unionBounds[1], bounds[1]);
            unionBounds[2] = Math.max(unionBounds[2], bounds[2]);
            unionBounds[3] = Math.min(unionBounds[3], bounds[3]);
        }
        return unionBounds;
    }

    /**
     * 選ばれたモードで作成されるグループの外枠を返す（実行時と同じ判定を使う）
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @param {string} mode - モードのキー
     * @param {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} groupingOptions - ダイアログの設定
     * @returns {Array<number[]>} 作成されるグループごとの外枠
     */
    function planResultBounds(selectionInfo, mode, groupingOptions) {
        if (mode === "clipPlacedOnly") return selectionInfo.imageBoundsList;
        var clusterRule = getClusterRule(mode, groupingOptions);
        if (!clusterRule) return [];

        var clusterIndexes = collectSelectionClusters(selectionInfo, clusterRule);
        var resultBounds = [];
        for (var i = 0; i < clusterIndexes.length; i++) {
            if (clusterIndexes[i].length < 2) continue;
            /* クリップはパスを含むまとまりだけが対象 / Clipping needs a path in the cluster */
            if (MASK_SIDES[mode] && !hasPathMember(selectionInfo.isPath, clusterIndexes[i])) continue;
            resultBounds.push(getUnionBounds(selectionInfo.boundsList, clusterIndexes[i]));
        }
        return resultBounds;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューの線の色を作る（ドキュメントのカラーモードに合わせる）
     * @param {Document} doc - 対象のドキュメント
     * @returns {RGBColor|CMYKColor} 線の色
     */
    function createPreviewColor(doc) {
        var previewColor;
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            previewColor = new CMYKColor();
            previewColor.cyan = PREVIEW_CMYK[0];
            previewColor.magenta = PREVIEW_CMYK[1];
            previewColor.yellow = PREVIEW_CMYK[2];
            previewColor.black = PREVIEW_CMYK[3];
        } else {
            previewColor = new RGBColor();
            previewColor.red = PREVIEW_RGB[0];
            previewColor.green = PREVIEW_RGB[1];
            previewColor.blue = PREVIEW_RGB[2];
        }
        return previewColor;
    }

    /**
     * 作成されるグループの範囲を枠で示すプレビューを用意する（最上位の専用レイヤーに描き、閉じるときにレイヤーごと消す）
     * @param {Document} doc - 対象のドキュメント
     * @returns {{show: function(Array<number[]>): void, remove: function(): void}} プレビューの操作
     */
    function createGroupPreview(doc) {
        var previousActiveLayer = doc.activeLayer;
        var previewColor = createPreviewColor(doc);
        var previewLayer = null;

        /* 外枠ごとに枠を描く / Draw one frame per result */
        function show(resultBounds) {
            if (!previewLayer) {
                previewLayer = doc.layers.add();
                previewLayer.name = PREVIEW_LAYER_NAME;
                /* 最上位レイヤーの一番上に置き、どのオブジェクトより手前に枠を出す
                   Keep the layer at the very top so the frames sit in front of every object */
                previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
            }
            for (var i = previewLayer.pageItems.length - 1; i >= 0; i--) {
                previewLayer.pageItems[i].remove();
            }
            for (var j = 0; j < resultBounds.length; j++) {
                var bounds = resultBounds[j];
                var previewFrame = previewLayer.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]);
                previewFrame.filled = false;
                previewFrame.stroked = true;
                previewFrame.strokeColor = previewColor;
                previewFrame.strokeWidth = PREVIEW_STROKE_WIDTH;
                previewFrame.opacity = PREVIEW_OPACITY;
            }
            app.redraw();
        }

        /* プレビューのレイヤーを消して、アクティブレイヤーを戻す / Remove the preview layer and restore the active layer */
        function remove() {
            if (!previewLayer) return;
            previewLayer.remove();
            previewLayer = null;
            doc.activeLayer = previousActiveLayer;
            app.redraw();
        }

        return { show: show, remove: remove };
    }

    // =========================================
    // クリッピングマスク / Clipping mask
    // =========================================

    /**
     * まとまりの中から型にするパスを探す
     * @param {PageItem[]} clusterItems - まとまり（前面から背面の順）
     * @param {string} maskSide - "front" なら最前面、"back" なら最背面のパス
     * @returns {PathItem|null} 型にするパス。パスがなければ null
     */
    function findMaskPath(clusterItems, maskSide) {
        var fromBack = (maskSide === "back");
        for (var i = 0; i < clusterItems.length; i++) {
            var candidateItem = clusterItems[fromBack ? clusterItems.length - 1 - i : i];
            if (candidateItem.typename === "PathItem") return candidateItem;
        }
        return null;
    }

    /**
     * まとまりを、指定のパスを型にしたクリップグループにする
     * @param {PageItem[]} clusterItems - まとまり（前面から背面の順）
     * @param {PathItem|null} maskPath - 型にするパス
     * @returns {void}
     */
    function clipCluster(clusterItems, maskPath) {
        if (clusterItems.length <= 1 || !maskPath) return;

        /* 型にするパスがあった位置にグループを置く / Put the group where the mask path was */
        var clipGroup = app.activeDocument.groupItems.add();
        clipGroup.move(maskPath, ElementPlacement.PLACEBEFORE);
        maskPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);

        /* 残りを前面から順に末尾へ足し、元の重ね順を保つ / Append the rest front to back to keep the stacking order */
        for (var i = 0; i < clusterItems.length; i++) {
            if (clusterItems[i] !== maskPath) {
                clusterItems[i].move(clipGroup, ElementPlacement.PLACEATEND);
            }
        }
        clipGroup.clipped = true;
    }

    /**
     * まとまりごとにクリッピングマスクを作成する
     * @param {Array<PageItem[]>} clusters - まとまりの配列
     * @param {string} maskSide - "front" なら最前面、"back" なら最背面のパスを型にする
     * @returns {void}
     */
    function clipEachCluster(clusters, maskSide) {
        for (var i = 0; i < clusters.length; i++) {
            clipCluster(clusters[i], findMaskPath(clusters[i], maskSide));
        }
    }

    // =========================================
    // 配置画像のクリップ / Placed image clipping
    // =========================================

    /**
     * 画像を、表示範囲と同じ大きさの長方形でクリップする（画像があった位置に置く）
     * @param {PlacedItem|RasterItem} image - クリップする画像
     * @returns {GroupItem} 作成したクリップグループ
     */
    function maskImageWithBounds(image) {
        var imageLayer = image.layer;
        var wasLocked = imageLayer.locked;
        var wasTemplate = imageLayer.isTemplate;
        if (wasLocked) imageLayer.locked = false;
        if (wasTemplate) imageLayer.isTemplate = false;

        var bounds = image.visibleBounds;
        var maskRect = imageLayer.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]);
        maskRect.stroked = false;
        maskRect.filled = false;

        /* 画像があった位置にグループを置き、画像→長方形の順に先頭へ入れて長方形を最前面にする
           Put the group where the image was; add the image, then the rectangle, so the rectangle ends up on top */
        var maskedGroup = imageLayer.groupItems.add();
        maskedGroup.move(image, ElementPlacement.PLACEBEFORE);
        image.moveToBeginning(maskedGroup);
        maskRect.moveToBeginning(maskedGroup);
        maskedGroup.clipped = true;

        if (wasLocked) imageLayer.locked = true;
        if (wasTemplate) imageLayer.isTemplate = true;
        return maskedGroup;
    }

    /**
     * グループの直下にある画像を返す
     * @param {GroupItem} parentGroup - 調べるグループ
     * @returns {Array<PlacedItem|RasterItem>} 画像（前面から背面の順）
     */
    function getChildImages(parentGroup) {
        var childImages = [];
        for (var i = 0; i < parentGroup.pageItems.length; i++) {
            if (isImageItem(parentGroup.pageItems[i])) childImages.push(parentGroup.pageItems[i]);
        }
        return childImages;
    }

    /**
     * クリップグループの元のマスクを削除し、中の画像を1つずつクリップし直す（画像以外はそのまま残す）
     * @param {GroupItem} clipGroup - クリップグループ
     * @returns {GroupItem[]} 作成したクリップグループ
     */
    function remaskClipGroup(clipGroup) {
        /* 元のマスクはクリップグループの最前面にある / The old mask is the topmost item of the clipping group */
        var oldMask = clipGroup.pageItems[0];
        clipGroup.clipped = false;
        if (!isImageItem(oldMask)) oldMask.remove();

        var childImages = getChildImages(clipGroup);
        var maskedGroups = [];
        for (var j = 0; j < childImages.length; j++) {
            maskedGroups.push(maskImageWithBounds(childImages[j]));
        }

        /* 画像以外が残っていなければ、作成したグループを前面から順に外へ出して元のグループを消す
           If nothing else is left, move the new groups out front to back and remove the original group */
        if (clipGroup.pageItems.length === maskedGroups.length) {
            for (var k = 0; k < maskedGroups.length; k++) {
                maskedGroups[k].move(clipGroup, ElementPlacement.PLACEBEFORE);
            }
            clipGroup.remove();
        }
        return maskedGroups;
    }

    /**
     * 選択中の画像とクリップグループ内の画像を、それぞれの大きさでクリップする
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {void}
     */
    function clipPlacedImages(selectedItems) {
        var maskedGroups = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
                maskedGroups = maskedGroups.concat(remaskClipGroup(selectedItem));
            } else if (isImageItem(selectedItem)) {
                maskedGroups.push(maskImageWithBounds(selectedItem));
            }
        }

        /* 作成したグループを選択する / Select the new groups */
        selectOnly(maskedGroups);
    }

    // =========================================
    // グループ化 / Grouping
    // =========================================

    /**
     * まとまりを1つのグループにする
     * @param {PageItem[]} clusterItems - まとまり（前面から背面の順）
     * @returns {GroupItem} 作成したグループ
     */
    function groupCluster(clusterItems) {
        /* 最前面のオブジェクトの位置にグループを置く / Put the group where the frontmost item was */
        var newGroup = app.activeDocument.groupItems.add();
        newGroup.move(clusterItems[0], ElementPlacement.PLACEBEFORE);

        /* 前面から順に末尾へ足し、元の重ね順を保つ / Append front to back to keep the stacking order */
        for (var i = 0; i < clusterItems.length; i++) {
            clusterItems[i].move(newGroup, ElementPlacement.PLACEATEND);
        }
        return newGroup;
    }

    /**
     * 2つ以上のオブジェクトを含むまとまりをグループ化し、作成したグループを選択する
     * @param {Array<PageItem[]>} clusters - まとまりの配列
     * @returns {void}
     */
    function groupEachCluster(clusters) {
        var newGroups = [];
        for (var i = 0; i < clusters.length; i++) {
            if (clusters[i].length > 1) newGroups.push(groupCluster(clusters[i]));
        }
        if (newGroups.length) selectOnly(newGroups);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * しきい値を表示用の文字列にする
     * @param {number} thresholdValue - しきい値（pt）
     * @returns {string} 「10 pt」の形の文字列
     */
    function formatThreshold(thresholdValue) {
        return Math.round(thresholdValue) + " pt";
    }

    /**
     * モードのラジオボタンを並べたパネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} panelTitle - パネルの見出し
     * @param {string[]} modeKeys - 並べるモードのキー
     * @param {Object} modeRadios - 作成したラジオボタンをモードのキーで登録する先
     * @param {number} [columnCount] - 列の数（省略時は1列）。左上から横へ順に並べる
     * @returns {Panel} 追加したパネル
     */
    function addModePanel(parentWindow, panelTitle, modeKeys, modeRadios, columnCount) {
        var modePanel = parentWindow.add("panel", undefined, panelTitle);
        modePanel.orientation = "column";
        modePanel.alignChildren = "left";
        modePanel.margins = PANEL_MARGINS;

        /* 複数列のときは列ごとのグループを横に並べる / For multiple columns, lay out one group per column side by side */
        var radioColumns = [modePanel];
        if (columnCount > 1) {
            var columnsRow = modePanel.add("group");
            columnsRow.orientation = "row";
            columnsRow.alignChildren = ["left", "top"];
            columnsRow.spacing = RADIO_COLUMN_SPACING;
            radioColumns = [];
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                var radioColumn = columnsRow.add("group");
                radioColumn.orientation = "column";
                radioColumn.alignChildren = "left";
                radioColumns.push(radioColumn);
            }
        }

        for (var i = 0; i < modeKeys.length; i++) {
            var modeKey = modeKeys[i];
            var modeRadio = radioColumns[i % radioColumns.length].add("radiobutton", undefined, getLabel(LABELS.radio[modeKey]));
            modeRadio.helpTip = getLabel(LABELS.tooltip[modeKey]);
            modeRadios[modeKey] = modeRadio;
        }
        return modePanel;
    }

    /**
     * しきい値の見出し・スライダー・値の表示を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @returns {{caption: StaticText, slider: Slider, valueText: StaticText}} 追加したコントロール
     */
    function addThresholdControls(parentPanel) {
        var thresholdTip = getLabel(LABELS.tooltip.threshold);

        var thresholdCaption = parentPanel.add("statictext", undefined, labelText(LABELS.fieldLabel.threshold));
        thresholdCaption.helpTip = thresholdTip;

        /* パネルの幅に合わせて伸ばし、値の表示をスライダーの中央に揃える / Stretch to the panel width so the readout stays centered under the slider */
        var thresholdSlider = parentPanel.add("slider", undefined, DEFAULT_THRESHOLD, THRESHOLD_MIN, THRESHOLD_MAX);
        thresholdSlider.preferredSize.width = SLIDER_WIDTH;
        thresholdSlider.alignment = "fill";
        thresholdSlider.helpTip = thresholdTip;

        var thresholdValueText = parentPanel.add("statictext", undefined, formatThreshold(DEFAULT_THRESHOLD));
        thresholdValueText.alignment = "center";
        thresholdValueText.characters = VALUE_TEXT_CHARS;
        thresholdValueText.helpTip = thresholdTip;

        return { caption: thresholdCaption, slider: thresholdSlider, valueText: thresholdValueText };
    }

    /**
     * チェックボックスを追加する
     * @param {Window|Panel} parentPanel - 追加先のダイアログかパネル
     * @param {string} optionKey - LABELS.checkbox / LABELS.tooltip のキー
     * @param {boolean} initialValue - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, optionKey, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(LABELS.checkbox[optionKey]));
        optionCheckbox.value = initialValue;
        optionCheckbox.helpTip = getLabel(LABELS.tooltip[optionKey]);
        return optionCheckbox;
    }

    /**
     * 選択数と、作成されるグループ数の表示を追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {number} itemCount - 選択しているオブジェクトの数
     * @returns {StaticText} 作成されるグループ数の表示
     */
    function addCountReadout(parentWindow, itemCount) {
        var countGroup = parentWindow.add("group");
        countGroup.orientation = "column";
        countGroup.alignment = "fill";
        countGroup.alignChildren = "fill";
        countGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.selectedCount) + itemCount);
        var resultCountText = countGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.groupCount) + 0);
        resultCountText.helpTip = getLabel(LABELS.tooltip.resultCount);
        return resultCountText;
    }

    /**
     * @typedef {Object} DialogControls ダイアログと、設定を読み書きするコントロール
     * @property {Window} dialogWindow - ダイアログ
     * @property {Object} modeRadios - モードのキーで引けるラジオボタン
     * @property {{caption: StaticText, slider: Slider, valueText: StaticText}} thresholdControls - しきい値の見出し・スライダー・値の表示
     * @property {Checkbox} splitAtGapsCheckbox - 「間隔が空いたら分ける」
     * @property {Checkbox} perArtboardCheckbox - 「アートボードごと」
     * @property {StaticText} resultCountText - 作成されるグループ数の表示
     * @property {Checkbox} previewCheckbox - 「プレビューを表示」
     */

    /**
     * ダイアログを組み立てる（表示はしない）
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @returns {DialogControls} ダイアログとコントロール
     */
    function buildDialog(selectionInfo) {
        var clipAndGroupDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        clipAndGroupDialog.orientation = "column";
        clipAndGroupDialog.alignChildren = "left";
        clipAndGroupDialog.margins = DIALOG_MARGINS;

        var modeRadios = {};
        addModePanel(clipAndGroupDialog, getLabel(LABELS.panel.clip), CLIP_MODES, modeRadios);
        var groupingPanel = addModePanel(clipAndGroupDialog, getLabel(LABELS.panel.grouping), GROUP_MODES, modeRadios, 2);
        groupingPanel.alignment = "fill";  /* ダイアログの幅いっぱいに広げる / Stretch to the dialog width */
        var thresholdControls = addThresholdControls(groupingPanel);
        var splitAtGapsCheckbox = addOptionCheckbox(groupingPanel, "splitAtGaps", DEFAULT_SPLIT_AT_GAPS);
        /* 選択が複数のアートボードにまたがるときだけ最初からON / On by default only when the selection spans artboards */
        var perArtboardCheckbox = addOptionCheckbox(groupingPanel, "perArtboard",
            selectionInfo.hasMultipleArtboards && spansMultipleArtboards(selectionInfo.artboardIndexes));
        var resultCountText = addCountReadout(clipAndGroupDialog, selectionInfo.itemCount);
        var previewCheckbox = addOptionCheckbox(clipAndGroupDialog, "showPreview", DEFAULT_SHOW_PREVIEW);

        /* ボタンエリア（左右中央） / Button row (centered) */
        var btnRowGroup = clipAndGroupDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return {
            dialogWindow: clipAndGroupDialog,
            modeRadios: modeRadios,
            thresholdControls: thresholdControls,
            splitAtGapsCheckbox: splitAtGapsCheckbox,
            perArtboardCheckbox: perArtboardCheckbox,
            resultCountText: resultCountText,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * 選択中のモードのキーを返す
     * @param {Object} modeRadios - モードのキーで引けるラジオボタン
     * @returns {string|null} モードのキー（どれも選ばれていなければ null）
     */
    function getSelectedMode(modeRadios) {
        for (var i = 0; i < ALL_MODES.length; i++) {
            if (modeRadios[ALL_MODES[i]].value) return ALL_MODES[i];
        }
        return null;
    }

    /**
     * ダイアログの設定を読み取る（しきい値は表示と実行で同じく整数に丸める）
     * @param {DialogControls} dialogControls - ダイアログとコントロール
     * @param {boolean} hasMultipleArtboards - ドキュメントにアートボードが2つ以上あるか
     * @returns {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} ダイアログの設定
     */
    function readGroupingOptions(dialogControls, hasMultipleArtboards) {
        return {
            threshold: Math.round(dialogControls.thresholdControls.slider.value),
            splitAtGaps: dialogControls.splitAtGapsCheckbox.value,
            perArtboard: dialogControls.perArtboardCheckbox.value && hasMultipleArtboards
        };
    }

    /**
     * 選んだモードで使う設定だけを有効にする
     * @param {DialogControls} dialogControls - ダイアログとコントロール
     * @param {boolean} hasMultipleArtboards - ドキュメントにアートボードが2つ以上あるか
     * @returns {void}
     */
    function updateOptionControls(dialogControls, hasMultipleArtboards) {
        var usedOptions = MODE_OPTIONS[getSelectedMode(dialogControls.modeRadios)] || {};
        var thresholdEnabled = usedOptions.threshold === true;
        dialogControls.thresholdControls.caption.enabled = thresholdEnabled;
        dialogControls.thresholdControls.slider.enabled = thresholdEnabled;
        dialogControls.thresholdControls.valueText.enabled = thresholdEnabled;
        dialogControls.splitAtGapsCheckbox.enabled = usedOptions.splitAtGaps === true;
        dialogControls.perArtboardCheckbox.enabled = usedOptions.perArtboard === true && hasMultipleArtboards;
    }

    /**
     * ダイアログを表示して、選ばれたモードと設定を返す（開いている間は結果の範囲をプレビューする）
     * @param {string} initialMode - 最初に選んでおくモードのキー
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @returns {{mode: string, groupingOptions: Object}|null} キャンセル時は null
     */
    function showDialog(initialMode, selectionInfo) {
        var dialogControls = buildDialog(selectionInfo);
        var modeRadios = dialogControls.modeRadios;
        var thresholdControls = dialogControls.thresholdControls;
        var previewCheckbox = dialogControls.previewCheckbox;
        var hasMultipleArtboards = selectionInfo.hasMultipleArtboards;

        var groupPreview = selectionInfo.itemCount ? createGroupPreview(app.activeDocument) : null;
        previewCheckbox.enabled = (groupPreview !== null);

        /* 作成されるグループの数と範囲を表示する（モードか設定が変わったときだけ数え直す）
           Show the number and extent of the groups to be created; recount only when the mode or settings change */
        var lastResultKey = null;
        var lastResultBounds = [];
        function refreshResult() {
            var mode = getSelectedMode(modeRadios);
            var groupingOptions = readGroupingOptions(dialogControls, hasMultipleArtboards);
            var resultKey = [mode, groupingOptions.threshold, groupingOptions.splitAtGaps, groupingOptions.perArtboard].join(":");
            if (resultKey === lastResultKey) return;
            lastResultKey = resultKey;

            lastResultBounds = planResultBounds(selectionInfo, mode, groupingOptions);
            var countLabel = (MASK_SIDES[mode] || mode === "clipPlacedOnly") ? LABELS.fieldLabel.clipCount : LABELS.fieldLabel.groupCount;
            dialogControls.resultCountText.text = labelText(countLabel) + lastResultBounds.length;
            updatePreview();
        }

        /* 「プレビューを表示」がONのときだけ枠を描き、OFFなら消す / Draw the frames only while "Show preview" is on; remove them otherwise */
        function updatePreview() {
            if (!groupPreview) return;
            if (previewCheckbox.value) {
                groupPreview.show(lastResultBounds);
            } else {
                groupPreview.remove();
            }
        }

        /* ラジオボタンは2つのパネルに分かれているため、排他は手動で行う
           The radios span two panels, so exclusivity is handled by hand */
        function selectMode() {
            for (var i = 0; i < ALL_MODES.length; i++) {
                modeRadios[ALL_MODES[i]].value = (modeRadios[ALL_MODES[i]] === this);
            }
            updateOptionControls(dialogControls, hasMultipleArtboards);
            refreshResult();
        }

        for (var i = 0; i < ALL_MODES.length; i++) {
            modeRadios[ALL_MODES[i]].onClick = selectMode;
        }
        dialogControls.splitAtGapsCheckbox.onClick = refreshResult;
        dialogControls.perArtboardCheckbox.onClick = refreshResult;
        previewCheckbox.onClick = updatePreview;
        /* 選択が多いときは、ドラッグ中は値の表示だけ更新し、指を離したときに数える
           With a large selection, update only the readout while dragging and recount on release */
        thresholdControls.slider.onChanging = function () {
            thresholdControls.valueText.text = formatThreshold(thresholdControls.slider.value);
            if (selectionInfo.itemCount <= LIVE_COUNT_MAX_ITEMS) refreshResult();
        };
        thresholdControls.slider.onChange = refreshResult;

        modeRadios[initialMode].value = true;
        updateOptionControls(dialogControls, hasMultipleArtboards);
        refreshResult();

        /* OK・キャンセルのどちらで閉じても、プレビューは実行前に消す / Remove the preview before running, however the dialog closes */
        var dialogReturnCode = dialogControls.dialogWindow.show();
        if (groupPreview) groupPreview.remove();
        if (dialogReturnCode !== 1) return null;
        return {
            mode: getSelectedMode(modeRadios),
            groupingOptions: readGroupingOptions(dialogControls, hasMultipleArtboards)
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選ばれたモードの処理を実行する
     * @param {string} mode - モードのキー
     * @param {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} groupingOptions - ダイアログの設定
     * @param {SelectionInfo} selectionInfo - ダイアログを開く前に読み取った選択の情報
     * @returns {void}
     */
    function runMode(mode, groupingOptions, selectionInfo) {
        if (!selectionInfo.itemCount) return;
        if (mode === "clipPlacedOnly") {
            clipPlacedImages(selectionInfo.items);
            return;
        }

        var clusterRule = getClusterRule(mode, groupingOptions);
        if (!clusterRule) return;

        /* プレビューと同じ判定でまとまりを作る / Build the clusters with the same rule as the preview */
        var clusterIndexes = collectSelectionClusters(selectionInfo, clusterRule);
        var clusters = [];
        for (var i = 0; i < clusterIndexes.length; i++) {
            var clusterItems = [];
            for (var j = 0; j < clusterIndexes[i].length; j++) {
                clusterItems.push(selectionInfo.items[clusterIndexes[i][j]]);
            }
            clusters.push(clusterItems);
        }

        if (MASK_SIDES[mode]) {
            clipEachCluster(clusters, MASK_SIDES[mode]);
        } else {
            groupEachCluster(clusters);
        }
    }

    /**
     * ダイアログを表示して、選ばれた処理を実行する
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        var selectionInfo = readSelection(selectedItems);
        /* すべて画像なら「配置画像のみをクリップ」を初期選択にする / Preselect "Clip Placed Images Only" when every item is an image */
        var initialMode = isAllImages(selectedItems) ? "clipPlacedOnly" : "groupProximity";
        var dialogSettings = showDialog(initialMode, selectionInfo);
        if (dialogSettings) runMode(dialogSettings.mode, dialogSettings.groupingOptions, selectionInfo);
    }

    main();

})();
