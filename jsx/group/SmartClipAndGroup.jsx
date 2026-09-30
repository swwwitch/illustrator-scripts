#target illustrator
#targetengine "SmartClipAndGroupEngine"
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
var SCRIPT_VERSION  = "v1.0.11";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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
     * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function centerButtonRowIfRightOnly(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

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
    // 選択の読み取りと結果の見積もり / Selection snapshot and result plan
    // =========================================

    /**
     * @typedef {Object} SelectionInfo ダイアログを開く前に読み取った選択の情報
     * @property {PageItem[]} items - 選択オブジェクト（前面から背面の順）
     * @property {number} itemCount - 選択オブジェクトの数
     * @property {Array<number[]>} boundsList - 各オブジェクトの geometricBounds（クリップグループはマスクの範囲）
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
            /* クリップグループはマスクの範囲で測る（隠れた部分でまとまりをつながない）
               Clip groups are measured by their mask, so hidden parts do not link clusters */
            selectionInfo.boundsList.push(getClipAwareBounds(selectedItem, false) || selectedItem.geometricBounds);
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
        var buttonRow = addButtonRow(clipAndGroupDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

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

        prepareDialogWindow(dialogControls.dialogWindow, SCRIPT_NAME);
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
