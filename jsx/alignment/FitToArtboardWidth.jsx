#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクト全体をひとまとまりとして、縦横比を保ったままアートボードの幅（既定は90%）に合わせてリサイズし、アートボードの中央に配置します。ダイアログはなく、実行するとそのまま処理します。

詳細は README を参照してください。

### Overview

Scales the whole selection as one cluster to a percentage of the artboard width (90% by default)
while keeping its aspect ratio, then centers it on the artboard. There is no dialog; running the script
applies the change directly.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitToArtboardWidth";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-21";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitToArtboardWidth.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitToArtboardWidth.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* アートボード幅に対する仕上がり幅の割合（%）。100 でアートボード幅ぴったり
       Target width as a percentage of the artboard width (100 = full artboard width) */
    var WIDTH_PERCENT = 90;

    /* 計測に使う境界 / Bounds used for measuring
       true : プレビュー境界（線幅・効果込みの見た目の端） / preview bounds (incl. strokes & effects)
       false: 幾何境界（パスの端） / geometric bounds (path edges) */
    var USE_PREVIEW_BOUNDS = true;

    /* リサイズ後に上下中央へもそろえるか（false なら縦位置はそのまま）
       Also center vertically after resizing (false keeps the vertical position) */
    var CENTER_VERTICALLY = true;

    // =========================================
    // 日英ラベル定義 / Japanese-English label definitions
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
            zeroWidth:   { ja: "選択範囲または仕上がり幅が0のため、リサイズできません。", en: "The selection or the target width is zero, so nothing can be resized." }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('alert','noDocument')）
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

    // =========================================
    // 矩形の計測 / Bounds
    // =========================================

    /**
     * オブジェクト1つの境界を返す（USE_PREVIEW_BOUNDS に従う）
     * @param {PageItem} item - 対象オブジェクト
     * @returns {number[]} [左, 上, 右, 下] の座標
     */
    function getItemBounds(item) {
        return USE_PREVIEW_BOUNDS ? item.visibleBounds : item.geometricBounds;
    }

    /**
     * 選択オブジェクト全体を囲む矩形を求める
     * @param {Array} items - 対象オブジェクト
     * @returns {number[]} [左, 上, 右, 下] の座標
     */
    function getCombinedBounds(items) {
        var bounds = getItemBounds(items[0]);
        var left = bounds[0];
        var top = bounds[1];
        var right = bounds[2];
        var bottom = bounds[3];
        for (var i = 1; i < items.length; i++) {
            var itemBounds = getItemBounds(items[i]);
            if (itemBounds[0] < left) left = itemBounds[0];
            if (itemBounds[1] > top) top = itemBounds[1];
            if (itemBounds[2] > right) right = itemBounds[2];
            if (itemBounds[3] < bottom) bottom = itemBounds[3];
        }
        return [left, top, right, bottom];
    }

    /**
     * 矩形の中心座標を求める
     * @param {number[]} rect - [左, 上, 右, 下] の座標
     * @returns {number[]} [中心X, 中心Y] の座標
     */
    function getRectCenter(rect) {
        return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    /**
     * 2つの矩形が重なっている面積を求める
     * @param {number[]} boundsA - [左, 上, 右, 下] の座標
     * @param {number[]} boundsB - [左, 上, 右, 下] の座標
     * @returns {number} 重なっている面積（重ならない場合は 0）
     */
    function getOverlapArea(boundsA, boundsB) {
        var overlapWidth = Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]);
        var overlapHeight = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
        if (overlapWidth <= 0 || overlapHeight <= 0) {
            return 0;
        }
        return overlapWidth * overlapHeight;
    }

    // =========================================
    // アートボード / Artboards
    // =========================================

    /**
     * 選択範囲と最も広く重なるアートボードを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} selectionBounds - [左, 上, 右, 下] の座標
     * @param {number} preferredIndex - 重なりが同じときに優先するアートボード番号
     * @returns {number} アートボード番号（どこにも重ならない場合は -1）
     */
    function findOverlappingArtboardIndex(doc, selectionBounds, preferredIndex) {
        /* 優先するアートボードを先に見て、重なりが同じなら切り替えない / Check the preferred artboard first so ties keep it */
        var searchOrder = [preferredIndex];
        for (var i = 0; i < doc.artboards.length; i++) {
            if (i !== preferredIndex) {
                searchOrder.push(i);
            }
        }
        var bestIndex = -1;
        var bestArea = 0;
        for (var j = 0; j < searchOrder.length; j++) {
            var area = getOverlapArea(selectionBounds, doc.artboards[searchOrder[j]].artboardRect);
            if (area > bestArea) {
                bestArea = area;
                bestIndex = searchOrder[j];
            }
        }
        return bestIndex;
    }

    /**
     * 選択範囲の中心に最も近いアートボードを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} selectionBounds - [左, 上, 右, 下] の座標
     * @returns {number} アートボード番号
     */
    function findNearestArtboardIndex(doc, selectionBounds) {
        var selectionCenter = getRectCenter(selectionBounds);
        var nearestIndex = 0;
        var nearestDistance = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardCenter = getRectCenter(doc.artboards[i].artboardRect);
            var dx = selectionCenter[0] - artboardCenter[0];
            var dy = selectionCenter[1] - artboardCenter[1];
            var distance = dx * dx + dy * dy;
            if (nearestDistance === null || distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /**
     * 基準にするアートボードを決める（選択と最も広く重なるもの、なければ中心が最も近いもの）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} selectionBounds - [左, 上, 右, 下] の座標
     * @returns {number[]} 基準にするアートボードの [左, 上, 右, 下] の座標
     */
    function findTargetArtboardRect(doc, selectionBounds) {
        var activeIndex = doc.artboards.getActiveArtboardIndex();
        if (doc.artboards.length < 2) {
            return doc.artboards[activeIndex].artboardRect;
        }
        var targetIndex = findOverlappingArtboardIndex(doc, selectionBounds, activeIndex);
        /* どのアートボードにも重ならないときは一番近いアートボードを使う / Fall back to the nearest artboard */
        if (targetIndex < 0) {
            targetIndex = findNearestArtboardIndex(doc, selectionBounds);
        }
        return doc.artboards[targetIndex].artboardRect;
    }

    // =========================================
    // リサイズと配置 / Resize & placement
    // =========================================

    /**
     * 選択全体をひとまとまりとして等倍スケールする（グループ化しない）
     * クラスタの左上を原点に、各オブジェクトのサイズと相対位置を同じ倍率で変形するため、
     * 親階層や重ね順を変えずに縦横比と配置の関係を保てる
     * @param {Array} items - 対象オブジェクト
     * @param {number} scaleFactor - 倍率（1 で等倍）
     * @returns {void}
     */
    function scaleItemsAsCluster(items, scaleFactor) {
        /* リサイズで位置がずれる前に、各オブジェクトの左上とクラスタの左上を控える
           Record each item's top-left and the cluster origin before anything moves */
        var originLeft = null;
        var originTop = null;
        var originalPositions = [];
        for (var i = 0; i < items.length; i++) {
            var itemLeft = items[i].left;
            var itemTop = items[i].top;
            originalPositions.push({ left: itemLeft, top: itemTop });
            if (originLeft === null || itemLeft < originLeft) originLeft = itemLeft;
            if (originTop === null || itemTop > originTop) originTop = itemTop;
        }

        var scalePercent = scaleFactor * 100;
        for (var j = 0; j < items.length; j++) {
            var item = items[j];
            /* 線幅・パターン・グラデーションも同じ倍率で変形する / Scale strokes, patterns and gradients alike */
            item.resize(scalePercent, scalePercent, true, true, true, true, scalePercent, Transformation.TOPLEFT);
            /* 原点からの相対位置も同倍率でスケール / Scale the offset from the origin by the same factor */
            item.left = originLeft + (originalPositions[j].left - originLeft) * scaleFactor;
            item.top = originTop - (originTop - originalPositions[j].top) * scaleFactor;
        }
    }

    /**
     * 選択全体をアートボードの中央へ移動する
     * @param {Array} items - 対象オブジェクト
     * @param {number[]} artboardRect - [左, 上, 右, 下] の座標
     * @param {boolean} centerVertically - 上下中央にもそろえる場合は true
     * @returns {void}
     */
    function centerItemsOnArtboard(items, artboardRect, centerVertically) {
        var artboardCenter = getRectCenter(artboardRect);
        var selectionCenter = getRectCenter(getCombinedBounds(items));
        var dx = artboardCenter[0] - selectionCenter[0];
        var dy = centerVertically ? (artboardCenter[1] - selectionCenter[1]) : 0;
        for (var i = 0; i < items.length; i++) {
            items[i].left += dx;
            items[i].top += dy;
        }
    }

    // =========================================
    // 選択の取得 / Selection
    // =========================================

    /**
     * 文字を部分選択している場合に、その文字を含むテキストオブジェクトを選択し直す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択し直したテキストオブジェクト
     */
    function selectTextFramesFromTextRange(doc) {
        var storyFrames = doc.selection.story.textFrames;
        var targetFrames = [];
        for (var i = 0; i < storyFrames.length; i++) {
            targetFrames.push(storyFrames[i]);
        }
        /* 文字編集を抜けてからテキストオブジェクトを選択 / Leave text editing, then select the frames */
        app.executeMenuCommand("deselectall");
        for (var j = 0; j < targetFrames.length; j++) {
            targetFrames[j].selected = true;
        }
        return targetFrames;
    }

    /**
     * 選択中のオブジェクトを固定した配列で取得する
     * doc.selection はライブ参照になりうるため、変形前にコピーして固定する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択中のオブジェクト（選択がない場合は空配列）
     */
    function getSelectedItems(doc) {
        var selection = doc.selection;
        /* 文字を部分選択しているときは selection が TextRange になるため、テキストオブジェクトに置き換える
           A partial text selection comes back as a TextRange; promote it to the text object */
        if (selection && !(selection instanceof Array)) {
            return selectTextFramesFromTextRange(doc);
        }
        var items = [];
        for (var i = 0; selection && i < selection.length; i++) {
            items.push(selection[i]);
        }
        return items;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確認し、選択全体をアートボードの幅に合わせてリサイズして中央に配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }
        var doc = app.activeDocument;

        var items = getSelectedItems(doc);
        if (items.length === 0) {
            alert(getLabel("alert", "noSelection"));
            return;
        }

        var selectionBounds = getCombinedBounds(items);
        var currentWidth = selectionBounds[2] - selectionBounds[0];
        var artboardRect = findTargetArtboardRect(doc, selectionBounds);
        var targetWidth = (artboardRect[2] - artboardRect[0]) * WIDTH_PERCENT / 100;
        if (currentWidth <= 0 || targetWidth <= 0) {
            alert(getLabel("alert", "zeroWidth"));
            return;
        }

        scaleItemsAsCluster(items, targetWidth / currentWidth);
        centerItemsOnArtboard(items, artboardRect, CENTER_VERTICALLY);
        app.redraw();
    }

    main();

})();
