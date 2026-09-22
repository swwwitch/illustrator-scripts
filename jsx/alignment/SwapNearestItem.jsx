#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを基準に、指定した方向（右／左／上／下）にある最も近いオブジェクトと位置を入れ替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItem.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n21a03e135423

### Overview

Swaps the selected object with the nearest object in a chosen direction — right, left, up or down.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItem.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapNearestItem";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItem.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItem.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n21a03e135423"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 入れ替え先を探す方向 "right" | "left" | "up" | "down"
       Direction to look for the object to swap with */
    var SEARCH_DIRECTION = "left";

    // =========================================
    // 基本設定 / Basic settings
    // =========================================

    /* 入れ替えの対象にできる typename / Typenames that can take part in a swap */
    var SWAPPABLE_TYPENAMES = [
        "PathItem", "CompoundPathItem", "GroupItem", "TextFrame",
        "PlacedItem", "RasterItem", "SymbolItem", "MeshItem",
        "PluginItem", "GraphItem"
    ];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のロケールに基づき言語コードを返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "1つ以上のオブジェクトを選択してください。", en: "Select at least one object." },
            noPosition:  { ja: "位置情報を持つオブジェクトを選択してください。", en: "Select an object that has a position." },
            noTarget:    { ja: "その方向に入れ替えられるオブジェクトが見つかりません。", en: "No object to swap with in that direction." }
        }
    };

    /**
     * ラベル定義からUI言語に応じた文字列を取得する
     * @param {string} labelPath - "alert.noDocument" のようなドット区切りのパス
     * @returns {string} 対応する文字列。見つからない場合はパスをそのまま返す
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // 判定・計測 / Checks and measurement
    // =========================================

    /**
     * 入れ替えの対象にできるオブジェクトか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象にできる場合はtrue
     */
    function isSwappableItem(pageItem) {
        /* ガイドは typename が "PathItem" のため guides で判定する / Guides are PathItems, so test .guides */
        if (pageItem.typename === "PathItem" && pageItem.guides === true) return false;

        for (var i = 0; i < SWAPPABLE_TYPENAMES.length; i++) {
            if (SWAPPABLE_TYPENAMES[i] === pageItem.typename) return true;
        }
        return false;
    }

    /**
     * ロック・非表示のレイヤーに属していないか判定する
     * @param {Layer} layer - 判定するレイヤー
     * @returns {boolean} 触れる場合はtrue
     */
    function isUsableLayer(layer) {
        /* サブレイヤーは自身がロックされていなくても親レイヤー側でロックされることがある
           A sublayer can be locked or hidden by an ancestor layer even when its own flags are clear */
        for (var layerNode = layer; layerNode && layerNode.typename === "Layer"; layerNode = layerNode.parent) {
            if (layerNode.locked || layerNode.visible === false) return false;
        }
        return true;
    }

    /**
     * 配列に指定したオブジェクトが含まれるか判定する
     * @param {PageItem[]} candidateItems - 探す対象の配列
     * @param {PageItem} targetItem - 探すオブジェクト
     * @returns {boolean} 含まれる場合はtrue
     */
    function containsItem(candidateItems, targetItem) {
        if (!candidateItems) return false;
        for (var i = 0; i < candidateItems.length; i++) {
            if (candidateItems[i] === targetItem) return true;
        }
        return false;
    }

    /**
     * オブジェクトの境界を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {number[]} [left, top, right, bottom]
     */
    function getBounds(pageItem) {
        return pageItem.visibleBounds;
    }

    /**
     * オブジェクトの中心座標を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {number[]} [x, y]
     */
    function getCenter(pageItem) {
        var itemBounds = getBounds(pageItem);
        return [(itemBounds[0] + itemBounds[2]) / 2, (itemBounds[1] + itemBounds[3]) / 2];
    }

    // =========================================
    // 探索と入れ替え / Search and swap
    // =========================================

    /**
     * 候補として見てよいオブジェクトか判定する
     * @param {PageItem} candidateItem - 判定するオブジェクト
     * @param {PageItem} referenceItem - 基準オブジェクト
     * @param {PageItem[]} excludeItems - 候補から除外するオブジェクト（省略可）
     * @returns {boolean} 候補にできる場合はtrue
     */
    function isSearchCandidate(candidateItem, referenceItem, excludeItems) {
        if (candidateItem === referenceItem) return false;
        if (containsItem(excludeItems, candidateItem)) return false;
        if (candidateItem.locked || candidateItem.hidden) return false;
        if (!isUsableLayer(candidateItem.layer)) return false;
        if (!isSwappableItem(candidateItem)) return false;
        /* グループや複合パス内の子オブジェクトは除外（親のみ処理対象）/ Skip children of groups and compound paths */
        if (candidateItem.parent &&
            (candidateItem.parent.typename === "GroupItem" || candidateItem.parent.typename === "CompoundPathItem")) return false;
        return true;
    }

    /**
     * 探索方向における2つの境界の隙間を返す
     * 探索軸と直交する方向で範囲が重なっていない場合や、方向が合わない場合は null を返す。
     * @param {number[]} referenceBounds - 基準オブジェクトの境界
     * @param {number[]} candidateBounds - 候補オブジェクトの境界
     * @param {number} dx - 中心X の差（候補 − 基準）
     * @param {number} dy - 中心Y の差（候補 − 基準）
     * @param {string} searchDirection - 探索方向（"right" / "left" / "up" / "down"）
     * @returns {number|null} 隙間（重なっている場合は0）。対象外なら null
     */
    function getDirectionalGap(referenceBounds, candidateBounds, dx, dy, searchDirection) {
        /* 探索軸と直交する方向で範囲が重なっているか / Overlap on the axis perpendicular to the search */
        var verticalOverlap = !(referenceBounds[3] > candidateBounds[1] || referenceBounds[1] < candidateBounds[3]);
        var horizontalOverlap = !(referenceBounds[0] > candidateBounds[2] || referenceBounds[2] < candidateBounds[0]);

        var gap;
        if (searchDirection === "right" && dx > 0 && verticalOverlap) {
            gap = candidateBounds[0] - referenceBounds[2];
        } else if (searchDirection === "left" && dx < 0 && verticalOverlap) {
            gap = referenceBounds[0] - candidateBounds[2];
        } else if (searchDirection === "up" && dy > 0 && horizontalOverlap) {
            gap = candidateBounds[3] - referenceBounds[1];
        } else if (searchDirection === "down" && dy < 0 && horizontalOverlap) {
            gap = referenceBounds[3] - candidateBounds[1];
        } else {
            return null;
        }

        /* 重なっている場合は隙間0として扱う / Treat overlapping items as a zero gap */
        return (gap < 0) ? 0 : gap;
    }

    /**
     * 指定方向にある最も近いオブジェクトを検索する
     * 近さは探索方向の隙間で測る（中心間の直線距離では斜めのオブジェクトが勝ってしまう）。
     * @param {PageItem} referenceItem - 基準オブジェクト
     * @param {string} searchDirection - 探索方向（"right" / "left" / "up" / "down"）
     * @param {PageItem[]} excludeItems - 候補から除外するオブジェクト（省略可）
     * @returns {PageItem|null} 見つかったオブジェクト。なければnull
     */
    function findNearestObjectInDirection(referenceItem, searchDirection, excludeItems) {
        var documentItems = app.activeDocument.pageItems;
        var referenceBounds = getBounds(referenceItem);
        var referenceCenter = getCenter(referenceItem);

        var nearestItem = null;
        var minGap = Number.MAX_VALUE;
        var minDistance = Number.MAX_VALUE;

        for (var i = 0; i < documentItems.length; i++) {
            var candidateItem = documentItems[i];
            if (!isSearchCandidate(candidateItem, referenceItem, excludeItems)) continue;

            var candidateCenter = getCenter(candidateItem);
            var dx = candidateCenter[0] - referenceCenter[0];
            var dy = candidateCenter[1] - referenceCenter[1];

            var gap = getDirectionalGap(referenceBounds, getBounds(candidateItem), dx, dy, searchDirection);
            if (gap === null) continue;

            var distance = Math.sqrt(dx * dx + dy * dy);
            if (gap < minGap || (gap === minGap && distance < minDistance)) {
                minGap = gap;
                minDistance = distance;
                nearestItem = candidateItem;
            }
        }

        return nearestItem;
    }

    /**
     * 2つのオブジェクトを左右に入れ替える（全体の占有範囲と隙間を保つ）
     * @param {PageItem} itemA - 入れ替える一方
     * @param {PageItem} itemB - 入れ替えるもう一方
     * @returns {void}
     */
    function swapHorizontally(itemA, itemB) {
        var boundsA = getBounds(itemA);
        var boundsB = getBounds(itemB);

        /* itemAが左、itemBが右になるように並び替え / Order so that itemA is the left one */
        if (boundsA[0] > boundsB[0]) {
            var tempItem = itemA;
            itemA = itemB;
            itemB = tempItem;

            var tempBounds = boundsA;
            boundsA = boundsB;
            boundsB = tempBounds;
        }

        var widthB = boundsB[2] - boundsB[0];
        var gap = boundsB[0] - boundsA[2];

        itemB.left = boundsA[0];
        itemA.left = boundsA[0] + widthB + gap;
    }

    /**
     * 2つのオブジェクトを上下に入れ替える（全体の占有範囲と隙間を保つ）
     * @param {PageItem} itemA - 入れ替える一方
     * @param {PageItem} itemB - 入れ替えるもう一方
     * @returns {void}
     */
    function swapVertically(itemA, itemB) {
        var boundsA = getBounds(itemA);
        var boundsB = getBounds(itemB);

        /* itemAが上、itemBが下になるように並び替え / Order so that itemA is the upper one */
        if (boundsA[1] < boundsB[1]) {
            var tempItem = itemA;
            itemA = itemB;
            itemB = tempItem;

            var tempBounds = boundsA;
            boundsA = boundsB;
            boundsB = tempBounds;
        }

        var heightB = boundsB[1] - boundsB[3];
        var gap = boundsA[3] - boundsB[1];

        itemB.top = boundsA[1];
        itemA.top = boundsA[1] - heightB - gap;
    }

    /**
     * 探索方向に応じて、2つのオブジェクトを入れ替える
     * @param {PageItem} itemA - 入れ替える一方
     * @param {PageItem} itemB - 入れ替えるもう一方
     * @returns {void}
     */
    function swapItems(itemA, itemB) {
        if (SEARCH_DIRECTION === "up" || SEARCH_DIRECTION === "down") {
            swapVertically(itemA, itemB);
        } else {
            swapHorizontally(itemA, itemB);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択内容に応じて入れ替えを実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedObjects = doc.selection;

        /* 文字カーソルが立っていると selection は TextRange で、length は文字数になる
           A text caret gives a TextRange whose length counts characters */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length < 1) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 2つ選択時は探索せずにその2つを入れ替える / With two objects selected, swap them directly */
        if (selectedObjects.length === 2) {
            swapItems(selectedObjects[0], selectedObjects[1]);
            return;
        }

        var baseItem = selectedObjects[0];
        if (!isSwappableItem(baseItem)) {
            alert(getLabel("alert.noPosition"));
            return;
        }

        /* 選択中のオブジェクト同士で入れ替わらないよう、選択全体を候補から外す / Keep the selection out of the candidates */
        var nearestItem = findNearestObjectInDirection(baseItem, SEARCH_DIRECTION, selectedObjects);
        if (!nearestItem) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        swapItems(baseItem, nearestItem);
    }

    main();

})();
