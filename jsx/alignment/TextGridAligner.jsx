#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストフレームを行・列単位で整列またはグループ化します。
行方向と列方向のしきい値を独立して調整でき、行・列のアキを均等に配置するオプションもあります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextGridAligner.md

### Overview

Aligns or groups text frames by row and by column.
The row and column thresholds are tuned independently, and an option evens out the gaps between them.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextGridAligner.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextGridAligner";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextGridAligner.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextGridAligner.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 行・列の判定に使う隙間のしきい値の初期値（pt）/ initial gap threshold used to group items */
    var DEFAULT_GAP_THRESHOLD = 10;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X = 300;   /* ダイアログの表示位置：右(+)／左(-) */
    var DIALOG_OFFSET_Y = 0;     /* ダイアログの表示位置：下(+)／上(-) */
    var DIALOG_OPACITY = 0.95;   /* ダイアログの不透明度 0.0 - 1.0 */
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネル余白 [左,上,右,下] / panel margins */
    var BUTTON_ROW_MARGINS = [0, 10, 0, 10];    /* ボタン行の余白 / button row margins */
    var SLIDER_WIDTH = 150;                     /* しきい値スライダーの幅 / threshold slider width */
    var THRESHOLD_LABEL_CHARS = 5;              /* しきい値表示の文字数 / width of the threshold readout */

    // =========================================
    // 方向の定義 / Direction constants
    // =========================================
    var DIRECTION_HORIZONTAL = "horizontal"; /* 横並び＝同じ行 / a row */
    var DIRECTION_VERTICAL = "vertical";     /* 縦並び＝同じ列 / a column */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "テキスト整列・グループ化", en: "Text Alignment & Grouping" }
        },
        panel: {
            rows:    { ja: "行", en: "Rows" },
            columns: { ja: "列", en: "Columns" }
        },
        checkbox: {
            alignRows:       { ja: "揃え", en: "Align" },
            groupRows:       { ja: "行をグループ化", en: "Group rows" },
            distributeRows:  { ja: "アキを均等に", en: "Distribute evenly" },
            alignColumns:    { ja: "揃え", en: "Align" },
            groupColumns:    { ja: "列をグループ化", en: "Group columns" },
            distributeColumns: { ja: "アキを均等に", en: "Distribute evenly" }
        },
        tooltip: {
            alignRows: {
                ja: "同じ行とみなしたテキストを、天地中央でそろえます。",
                en: "Vertically centers the text objects that were grouped into the same row."
            },
            groupRows: {
                ja: "同じ行とみなしたテキストを1つのグループにまとめます。行と列は同時にグループ化できません。",
                en: "Groups each detected row into one group. Rows and columns cannot be grouped at the same time."
            },
            distributeRows: {
                ja: "作成した行グループどうしの縦のアキを均等にします。",
                en: "Evens out the vertical gaps between the row groups."
            },
            rowThreshold: {
                ja: "左右の隙間がこの値以内なら、同じ行とみなします。スライダーを動かすと結果がすぐ反映されます。",
                en: "Text objects with a horizontal gap up to this value form one row. The canvas updates as you drag."
            },
            alignColumns: {
                ja: "同じ列とみなしたテキストを、左右中央でそろえます。",
                en: "Horizontally centers the text objects that were grouped into the same column."
            },
            groupColumns: {
                ja: "同じ列とみなしたテキストを1つのグループにまとめます。行と列は同時にグループ化できません。",
                en: "Groups each detected column into one group. Rows and columns cannot be grouped at the same time."
            },
            distributeColumns: {
                ja: "作成した列グループどうしの横のアキを均等にします。",
                en: "Evens out the horizontal gaps between the column groups."
            },
            columnThreshold: {
                ja: "上下の隙間がこの値以内なら、同じ列とみなします。スライダーを動かすと結果がすぐ反映されます。",
                en: "Text objects with a vertical gap up to this value form one column. The canvas updates as you drag."
            }
        },
        button: {
            run:    { ja: "実行", en: "Run" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrames: { ja: "テキストが選択されていません。", en: "No text object is selected." }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.rows" のようなドット区切りキー
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
    // 判定のしきい値 / Detection thresholds
    // =========================================

    /* スライダーで更新される、行・列の判定しきい値 / Updated as the sliders move */
    var rowGapThreshold = DEFAULT_GAP_THRESHOLD;
    var columnGapThreshold = DEFAULT_GAP_THRESHOLD;

    // =========================================
    // 境界の計測 / Bounds
    // =========================================

    /**
     * オブジェクト群の結合バウンディングボックスを返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getCombinedBounds(targetItems) {
        var combinedBounds = targetItems[0].geometricBounds.slice(0);
        for (var i = 1; i < targetItems.length; i++) {
            var itemBounds = targetItems[i].geometricBounds;
            if (itemBounds[0] < combinedBounds[0]) combinedBounds[0] = itemBounds[0];
            if (itemBounds[1] > combinedBounds[1]) combinedBounds[1] = itemBounds[1];
            if (itemBounds[2] > combinedBounds[2]) combinedBounds[2] = itemBounds[2];
            if (itemBounds[3] < combinedBounds[3]) combinedBounds[3] = itemBounds[3];
        }
        return combinedBounds;
    }

    /**
     * 探索方向における2つの境界の隙間を返す
     * 直交する方向にずれている組は同じ行／列とみなさないため、非常に大きい値を返す。
     * @param {number[]} boundsA - 一方の境界 [左, 上, 右, 下]
     * @param {number[]} boundsB - もう一方の境界
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @returns {number} 隙間（対象外なら Number.MAX_VALUE）
     */
    function getGapAlongDirection(boundsA, boundsB, direction) {
        var horizontalGap = Math.max(0, Math.max(boundsB[0] - boundsA[2], boundsA[0] - boundsB[2]));
        var verticalGap = Math.max(0, Math.max(boundsB[3] - boundsA[1], boundsA[3] - boundsB[1]));

        if (direction === DIRECTION_HORIZONTAL) {
            return (verticalGap > 0) ? Number.MAX_VALUE : horizontalGap;
        }
        return (horizontalGap > 0) ? Number.MAX_VALUE : verticalGap;
    }

    /**
     * 2つの境界の重なり率（面積が大きい方に対する割合）を返す
     * @param {number[]} boundsA - 一方の境界 [左, 上, 右, 下]
     * @param {number[]} boundsB - もう一方の境界
     * @returns {number} 重なり率（重ならない場合は 0）
     */
    function getOverlapRatio(boundsA, boundsB) {
        var overlapWidth = Math.max(0, Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]));
        var overlapHeight = Math.max(0, Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]));
        var overlapArea = overlapWidth * overlapHeight;
        if (overlapArea <= 0) return 0;

        var areaA = (boundsA[2] - boundsA[0]) * (boundsA[1] - boundsA[3]);
        var areaB = (boundsB[2] - boundsB[0]) * (boundsB[1] - boundsB[3]);
        return overlapArea / Math.max(areaA, areaB);
    }

    // =========================================
    // 行・列の抽出 / Detecting rows and columns
    // =========================================

    /**
     * 隣接または重なっているオブジェクトをたどって1グループ分を集める
     * @param {number} startIndex - 起点のインデックス
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean[]} visited - 走査済みフラグ
     * @param {PageItem[]} collectedItems - 集めたオブジェクトの受け皿
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @param {number} gapThreshold - 同じ行／列とみなす隙間の上限
     * @returns {void}
     */
    function collectConnectedItems(startIndex, targetItems, visited, collectedItems, direction, gapThreshold) {
        visited[startIndex] = true;
        collectedItems.push(targetItems[startIndex]);

        var boundsA = targetItems[startIndex].visibleBounds;
        for (var j = 0; j < targetItems.length; j++) {
            if (visited[j]) continue;
            var boundsB = targetItems[j].visibleBounds;
            if (getOverlapRatio(boundsA, boundsB) > 0 ||
                getGapAlongDirection(boundsA, boundsB, direction) <= gapThreshold) {
                collectConnectedItems(j, targetItems, visited, collectedItems, direction, gapThreshold);
            }
        }
    }

    /**
     * 指定方向で隣接・重なっているオブジェクトをグループにまとめる
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @returns {Array} PageItem[] の配列（1グループ＝1行または1列）
     */
    function getConnectedGroups(targetItems, direction) {
        var gapThreshold = (direction === DIRECTION_HORIZONTAL) ? rowGapThreshold : columnGapThreshold;
        var itemGroups = [];
        var visited = [];

        for (var i = 0; i < targetItems.length; i++) visited[i] = false;

        for (var j = 0; j < targetItems.length; j++) {
            if (visited[j]) continue;
            var collectedItems = [];
            collectConnectedItems(j, targetItems, visited, collectedItems, direction, gapThreshold);
            itemGroups.push(collectedItems);
        }
        return itemGroups;
    }

    // =========================================
    // 整列・分配 / Aligning and distributing
    // =========================================

    /**
     * 行または列ごとに、その方向と直交する軸の中央でそろえる（グループ化はしない）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} direction - DIRECTION_HORIZONTAL（行＝天地中央）または DIRECTION_VERTICAL（列＝左右中央）
     * @returns {void}
     */
    function alignGroupsToCenter(targetItems, direction) {
        if (!targetItems || targetItems.length === 0) return;

        var itemGroups = getConnectedGroups(targetItems, direction);
        for (var i = 0; i < itemGroups.length; i++) {
            if (itemGroups[i].length <= 1) continue;
            centerItemsInGroup(itemGroups[i], direction);
        }
    }

    /**
     * 1グループ分のオブジェクトを、方向と直交する軸の中央にそろえる
     * @param {PageItem[]} groupItems - 1グループ分のオブジェクト
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @returns {void}
     */
    function centerItemsInGroup(groupItems, direction) {
        var combinedBounds = getCombinedBounds(groupItems);
        var isRow = (direction === DIRECTION_HORIZONTAL);
        /* 行は天地中央、列は左右中央にそろえる / rows center vertically, columns horizontally */
        var groupCenter = isRow ?
            (combinedBounds[1] + combinedBounds[3]) / 2 :
            (combinedBounds[0] + combinedBounds[2]) / 2;

        for (var i = 0; i < groupItems.length; i++) {
            var itemBounds = groupItems[i].geometricBounds;
            if (isRow) {
                groupItems[i].top += groupCenter - (itemBounds[1] + itemBounds[3]) / 2;
            } else {
                groupItems[i].left += groupCenter - (itemBounds[0] + itemBounds[2]) / 2;
            }
        }
    }

    /**
     * 並んだオブジェクトのアキを均等にする（両端の位置は保つ）
     * @param {PageItem[]} orderedItems - 対象オブジェクト（この関数内で並べ替える）
     * @param {string} direction - DIRECTION_HORIZONTAL（横に均等）または DIRECTION_VERTICAL（縦に均等）
     * @returns {void}
     */
    function distributeSpacingEvenly(orderedItems, direction) {
        if (!orderedItems || orderedItems.length <= 1) return;

        var isHorizontal = (direction === DIRECTION_HORIZONTAL);
        orderedItems.sort(isHorizontal ?
            function (itemA, itemB) { return itemA.geometricBounds[0] - itemB.geometricBounds[0]; } :
            function (itemA, itemB) { return itemB.geometricBounds[1] - itemA.geometricBounds[1]; });

        var totalItemSize = 0;
        for (var i = 0; i < orderedItems.length; i++) {
            totalItemSize += getItemSizeAlong(orderedItems[i], isHorizontal);
        }

        var combinedBounds = getCombinedBounds(orderedItems);
        var availableSpace = isHorizontal ?
            (combinedBounds[2] - combinedBounds[0]) :
            (combinedBounds[1] - combinedBounds[3]);
        var spacing = (availableSpace - totalItemSize) / (orderedItems.length - 1);

        var cursor = isHorizontal ? orderedItems[0].geometricBounds[0] : orderedItems[0].geometricBounds[1];
        for (var j = 0; j < orderedItems.length; j++) {
            var itemSize = getItemSizeAlong(orderedItems[j], isHorizontal);
            if (isHorizontal) {
                orderedItems[j].left = cursor;
                cursor += itemSize + spacing;
            } else {
                orderedItems[j].top = cursor;
                cursor -= itemSize + spacing;
            }
        }
    }

    /**
     * 指定軸方向のオブジェクトの寸法を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {number} 幅または高さ
     */
    function getItemSizeAlong(pageItem, isHorizontal) {
        var itemBounds = pageItem.geometricBounds;
        return isHorizontal ? (itemBounds[2] - itemBounds[0]) : (itemBounds[1] - itemBounds[3]);
    }

    // =========================================
    // グループ化 / Grouping
    // =========================================

    /**
     * 指定方向で抽出した行または列を、それぞれ1つのグループにまとめる
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @param {boolean} centerBeforeGrouping - まとめる前に中央でそろえるなら true
     * @returns {GroupItem[]} 作成したグループ
     */
    function groupItemsByDirection(direction, centerBeforeGrouping) {
        if (app.documents.length === 0) return [];

        var selectedItems = app.activeDocument.selection;
        if (!selectedItems || selectedItems.length === 0) return [];

        var itemGroups = getConnectedGroups(selectedItems, direction);
        var createdGroups = [];

        for (var i = 0; i < itemGroups.length; i++) {
            var groupItems = itemGroups[i];
            if (groupItems.length <= 1) continue;

            /* グループ化でレイヤーが移るため、元のレイヤーを控えておく
               Grouping can move items between layers, so remember the original one */
            var originalLayer = groupItems[0].layer;

            if (centerBeforeGrouping) centerItemsInGroup(groupItems, direction);

            app.executeMenuCommand('deselectall');
            for (var j = 0; j < groupItems.length; j++) {
                groupItems[j].selected = true;
            }
            app.executeMenuCommand('group');

            var createdGroup = app.activeDocument.selection[0];
            createdGroup.layer = originalLayer;
            createdGroups.push(createdGroup);
        }

        app.redraw();

        app.activeDocument.selection = null;
        for (var k = 0; k < createdGroups.length; k++) {
            createdGroups[k].selected = true;
        }
        return createdGroups;
    }

    /**
     * グループ内のテキストだけを取り出す
     * @param {GroupItem} groupItem - 対象グループ
     * @returns {TextFrame[]} グループ直下のテキスト
     */
    function getTextFramesInGroup(groupItem) {
        var textFrames = [];
        var groupPageItems = groupItem.pageItems;
        for (var i = 0; i < groupPageItems.length; i++) {
            if (groupPageItems[i].typename === "TextFrame") textFrames.push(groupPageItems[i]);
        }
        return textFrames;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * しきい値の表示用文字列を返す
     * @param {number} thresholdValue - スライダーの値（pt）
     * @returns {string} "10 pt" の形の文字列
     */
    function formatThreshold(thresholdValue) {
        return Math.round(thresholdValue) + " pt";
    }

    /**
     * ラベルと tooltip を LABELS の同じキーから引いてチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - checkbox.* と tooltip.* に共通のキー
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelKey, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        optionCheckbox.helpTip = getLabel("tooltip." + labelKey);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * 行または列のパネル（揃え・グループ化・アキを均等に・しきい値）を組み立てる
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {object} axisKeys - LABELS のキー（panel / align / group / distribute / threshold）
     * @param {number} sliderMax - しきい値スライダーの最大値
     * @returns {object} alignCheckbox / groupCheckbox / distributeCheckbox / thresholdSlider / thresholdLabel
     */
    function addAxisPanel(parentWindow, axisKeys, sliderMax) {
        var axisPanel = parentWindow.add("panel", undefined, getLabel("panel." + axisKeys.panel));
        axisPanel.orientation = "column";
        axisPanel.alignChildren = "left";
        axisPanel.margins = PANEL_MARGINS;

        var axisControls = {
            alignCheckbox: addOptionCheckbox(axisPanel, axisKeys.align, true),
            groupCheckbox: addOptionCheckbox(axisPanel, axisKeys.group, false),
            distributeCheckbox: addOptionCheckbox(axisPanel, axisKeys.distribute, false)
        };
        /* 「アキを均等に」は「グループ化」をONにするまで使えない / enabled only while grouping */
        axisControls.distributeCheckbox.enabled = false;

        var thresholdSlider = axisPanel.add("slider", undefined, DEFAULT_GAP_THRESHOLD, 0, sliderMax);
        thresholdSlider.helpTip = getLabel("tooltip." + axisKeys.threshold);
        thresholdSlider.value = Math.min(DEFAULT_GAP_THRESHOLD, sliderMax);
        thresholdSlider.preferredSize.width = SLIDER_WIDTH;

        var thresholdLabel = axisPanel.add("statictext", undefined, formatThreshold(thresholdSlider.value));
        thresholdLabel.alignment = "center";
        thresholdLabel.characters = THRESHOLD_LABEL_CHARS;

        axisControls.thresholdSlider = thresholdSlider;
        axisControls.thresholdLabel = thresholdLabel;

        axisControls.alignCheckbox.onClick = function () {
            syncEnabledState(axisControls);
        };
        return axisControls;
    }

    /**
     * 「揃え」のON/OFFに合わせて、その行または列の他のコントロールをディムする
     * @param {object} axisControls - addAxisPanel() が返すコントロール
     * @returns {void}
     */
    function syncEnabledState(axisControls) {
        var isAligning = axisControls.alignCheckbox.value;
        axisControls.thresholdSlider.enabled = isAligning;
        axisControls.thresholdLabel.enabled = isAligning;
        axisControls.groupCheckbox.enabled = isAligning;
        if (!isAligning) {
            axisControls.groupCheckbox.value = false;
            axisControls.distributeCheckbox.enabled = false;
            axisControls.distributeCheckbox.value = false;
        }
    }

    /**
     * 「グループ化」をONにしたら他方の「グループ化」を外し、「アキを均等に」の有効状態を合わせる
     * 行と列を同時にグループ化はできないため / Rows and columns cannot both be grouped
     * @param {object} ownControls - クリックされた側のコントロール
     * @param {object} otherControls - もう一方（行なら列、列なら行）のコントロール
     * @returns {void}
     */
    function bindExclusiveGrouping(ownControls, otherControls) {
        ownControls.groupCheckbox.onClick = function () {
            var isGrouping = ownControls.groupCheckbox.value;
            if (isGrouping) otherControls.groupCheckbox.value = false;
            ownControls.distributeCheckbox.enabled = isGrouping;
            if (!isGrouping) ownControls.distributeCheckbox.value = false;
        };
    }

    /**
     * ダイアログを開く前の位置を控える（キャンセルで戻す）
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @returns {number[][]} 各テキストの geometricBounds の複製
     */
    function snapshotBounds(textFrames) {
        var originalBoundsList = [];
        for (var i = 0; i < textFrames.length; i++) {
            originalBoundsList.push(textFrames[i].geometricBounds.slice(0));
        }
        return originalBoundsList;
    }

    /**
     * 控えた位置へテキストを戻す
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @param {number[][]} originalBoundsList - snapshotBounds() の戻り値
     * @returns {void}
     */
    function restoreBounds(textFrames, originalBoundsList) {
        for (var i = 0; i < textFrames.length; i++) {
            textFrames[i].left = originalBoundsList[i][0];
            textFrames[i].top = originalBoundsList[i][1];
        }
    }

    /**
     * 整列・グループ化のダイアログを表示する
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @returns {object|null} ［実行］で確定した処理内容（キャンセルなら null）
     */
    function showTextGridDialog(textFrames) {
        /* スライダー操作で動いた分をキャンセルで戻すため / so Cancel can undo slider moves */
        var originalBoundsList = snapshotBounds(textFrames);
        var confirmedOptions = null;

        var combinedBounds = getCombinedBounds(textFrames);
        var totalWidth = combinedBounds[2] - combinedBounds[0];
        var totalHeight = combinedBounds[1] - combinedBounds[3];

        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = "fill";
        alignDialog.opacity = DIALOG_OPACITY;

        var rowControls = addAxisPanel(alignDialog, {
            panel: "rows", align: "alignRows", group: "groupRows", distribute: "distributeRows", threshold: "rowThreshold"
        }, totalWidth || 100);
        var columnControls = addAxisPanel(alignDialog, {
            panel: "columns", align: "alignColumns", group: "groupColumns", distribute: "distributeColumns", threshold: "columnThreshold"
        }, totalHeight || 100);

        bindExclusiveGrouping(rowControls, columnControls);
        bindExclusiveGrouping(columnControls, rowControls);

        rowControls.thresholdSlider.onChanging = function () {
            rowControls.thresholdLabel.text = formatThreshold(rowControls.thresholdSlider.value);
            rowGapThreshold = rowControls.thresholdSlider.value;
            if (!rowControls.alignCheckbox.value) return;

            if (rowControls.groupCheckbox.value) {
                groupItemsByDirection(DIRECTION_HORIZONTAL, true);
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_HORIZONTAL);
            }
            app.redraw();
        };

        columnControls.thresholdSlider.onChanging = function () {
            columnControls.thresholdLabel.text = formatThreshold(columnControls.thresholdSlider.value);
            columnGapThreshold = columnControls.thresholdSlider.value;
            if (!columnControls.alignCheckbox.value) return;

            if (columnControls.groupCheckbox.value) {
                var createdGroups = groupItemsByDirection(DIRECTION_VERTICAL, false);
                if (columnControls.distributeCheckbox.value) {
                    /* 各列グループの中で、テキストの左右のアキを均等にする
                       Even out the horizontal gaps inside each column group */
                    for (var i = 0; i < createdGroups.length; i++) {
                        distributeSpacingEvenly(getTextFramesInGroup(createdGroups[i]), DIRECTION_HORIZONTAL);
                    }
                }
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_VERTICAL);
            }
            app.redraw();
        };

        /* ボタンエリア / Button row */
        var btnRowGroup = alignDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        btnRowGroup.margins = BUTTON_ROW_MARGINS;

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnCancel.onClick = function () {
            restoreBounds(textFrames, originalBoundsList);
            alignDialog.close(0);
        };

        var btnRun = btnRowGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });
        btnRun.onClick = function () {
            confirmedOptions = {
                alignRows: rowControls.alignCheckbox.value,
                groupRows: rowControls.groupCheckbox.value,
                distributeRows: rowControls.distributeCheckbox.value,
                alignColumns: columnControls.alignCheckbox.value,
                groupColumns: columnControls.groupCheckbox.value,
                distributeColumns: columnControls.distributeCheckbox.value
            };
            alignDialog.close(1);
        };

        alignDialog.onShow = function () {
            alignDialog.location = [
                alignDialog.location[0] + DIALOG_OFFSET_X,
                alignDialog.location[1] + DIALOG_OFFSET_Y
            ];
        };

        return (alignDialog.show() === 1) ? confirmedOptions : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したテキストを行・列で整列し、必要ならグループ化する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var selectedObjects = app.activeDocument.selection;
        var textFrames = [];
        for (var i = 0; selectedObjects && i < selectedObjects.length; i++) {
            if (selectedObjects[i].typename === "TextFrame") textFrames.push(selectedObjects[i]);
        }
        if (textFrames.length === 0) {
            alert(getLabel("alert.noTextFrames"));
            return;
        }

        var confirmedOptions = showTextGridDialog(textFrames);
        if (!confirmedOptions) return;

        applyConfirmedOptions(textFrames, confirmedOptions);
    }

    /**
     * ダイアログで確定した内容を適用する
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @param {object} confirmedOptions - showTextGridDialog() が返す処理内容
     * @returns {void}
     */
    function applyConfirmedOptions(textFrames, confirmedOptions) {
        if (confirmedOptions.alignRows) {
            if (confirmedOptions.groupRows) {
                var rowGroups = groupItemsByDirection(DIRECTION_HORIZONTAL, true);
                if (confirmedOptions.distributeRows && rowGroups.length > 1) {
                    distributeSpacingEvenly(rowGroups, DIRECTION_VERTICAL);
                }
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_HORIZONTAL);
            }
        }

        if (confirmedOptions.alignColumns) {
            if (confirmedOptions.groupColumns) {
                var columnGroups = groupItemsByDirection(DIRECTION_VERTICAL, false);
                if (confirmedOptions.distributeColumns && columnGroups.length > 1) {
                    distributeSpacingEvenly(columnGroups, DIRECTION_HORIZONTAL);
                }
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_VERTICAL);
            }
        }
    }

    main();

})();
