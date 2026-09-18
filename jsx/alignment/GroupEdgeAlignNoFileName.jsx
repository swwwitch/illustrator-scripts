#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択されているオブジェクト群の端または中心を、`ALIGNMENT_SIDE` で指定した方向へ揃えます。
GroupEdgeAlign.jsx からファイル名による方向判定を外した版です。

詳細は README を参照してください。

### Overview

Aligns the edges or the center of the selected objects in the direction given by `ALIGNMENT_SIDE`.
This is GroupEdgeAlign.jsx without the filename-based direction detection.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEdgeAlignNoFileName";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-04-06";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignNoFileName.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlignNoFileName.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 揃える方向 "left" | "right" | "top" | "bottom" | "CENTER_X" | "CENTER_Y" | "CENTER"
       スクリプト引数で渡された場合はそちらが優先される */
    var ALIGNMENT_SIDE = "right";

    /* アートボード端の手前にガイドがあれば、そのガイドへ揃える / snap to a guide when one is in the way */
    var USE_GUIDES = true;

    /* 境界の取り方 "preference"（環境設定に従う）| "preview"（線を含む）| "geometric"（線を含まない）*/
    var BOUNDS_MODE = "preference";

    /* ガイドの探し方 "inside"（揃える向きで最も近いもの）| "nearest"（距離が最も近いもの）*/
    var GUIDE_SEARCH_MODE = "inside";

    /* ガイドの水平・垂直判定に使う許容値（pt）/ tolerance for classifying a guide as H or V */
    var GUIDE_ORIENTATION_TOLERANCE = 0.01;

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクト群の端または中心を、指定方向へ揃える
     * @param {string} [requestedAlignmentSide] - スクリプト引数で渡された方向（省略時は ALIGNMENT_SIDE）
     * @returns {void}
     */
    function main(requestedAlignmentSide) {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var documentRef = app.activeDocument;
        var selectedItems = documentRef.selection;
        if (selectedItems.length === 0) {
            alert("オブジェクトが選択されていません。");
            return;
        }

        var alignmentSide = requestedAlignmentSide || ALIGNMENT_SIDE;

        var includeStrokeInBounds = resolveIncludeStrokeInBounds(BOUNDS_MODE);
        if (includeStrokeInBounds === null) {
            alert("BOUNDS_MODE の指定が不正です：" + BOUNDS_MODE);
            return;
        }

        var artboards = documentRef.artboards;
        var activeArtboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
        var selectionBounds = getSelectionBounds(selectedItems, includeStrokeInBounds);

        var moveOffset = getAlignmentOffset(documentRef, selectionBounds, activeArtboardRect, alignmentSide);
        if (moveOffset === null) return; /* 方向の指定が不正。getAlignmentOffset が通知済み */

        for (var i = 0; i < selectedItems.length; i++) {
            selectedItems[i].translate(moveOffset.x, moveOffset.y);
        }
    }

    /**
     * BOUNDS_MODE から「線を境界に含めるか」を決める
     * @param {string} boundsMode - "preference" | "preview" | "geometric"
     * @returns {boolean|null} 線を含めるなら true。指定が不正なら null
     */
    function resolveIncludeStrokeInBounds(boundsMode) {
        if (boundsMode === "preference") return app.preferences.getBooleanPreference("includeStrokeInBounds");
        if (boundsMode === "preview") return true;
        if (boundsMode === "geometric") return false;
        return null;
    }

    /**
     * 選択オブジェクト全体を囲む境界を返す
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @param {boolean} includeStrokeInBounds - 線を境界に含めるか
     * @returns {number[]} [left, top, right, bottom]
     */
    function getSelectionBounds(selectedItems, includeStrokeInBounds) {
        var selectionBounds = getItemBounds(selectedItems[0], includeStrokeInBounds).slice(0);

        for (var i = 1; i < selectedItems.length; i++) {
            var itemBounds = getItemBounds(selectedItems[i], includeStrokeInBounds);
            if (itemBounds[0] < selectionBounds[0]) selectionBounds[0] = itemBounds[0];
            if (itemBounds[1] > selectionBounds[1]) selectionBounds[1] = itemBounds[1];
            if (itemBounds[2] > selectionBounds[2]) selectionBounds[2] = itemBounds[2];
            if (itemBounds[3] < selectionBounds[3]) selectionBounds[3] = itemBounds[3];
        }

        return selectionBounds;
    }

    /**
     * 揃える方向から移動量を求める
     * @param {Document} documentRef - 対象のドキュメント
     * @param {number[]} selectionBounds - 選択範囲の境界 [L, T, R, B]
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @returns {{x: number, y: number}|null} 移動量。方向の指定が不正なら null
     */
    function getAlignmentOffset(documentRef, selectionBounds, artboardRect, alignmentSide) {
        var centerOffsetX = getHorizontalCenterValue(artboardRect) - getHorizontalCenterValue(selectionBounds);
        var centerOffsetY = getVerticalCenterValue(artboardRect) - getVerticalCenterValue(selectionBounds);

        if (alignmentSide === "CENTER_X") return { x: centerOffsetX, y: 0 };
        if (alignmentSide === "CENTER_Y") return { x: 0, y: centerOffsetY };
        if (alignmentSide === "CENTER") return { x: centerOffsetX, y: centerOffsetY };

        var selectionEdge = getEdgeValueForAlignmentSide(selectionBounds, alignmentSide);
        var artboardEdge = getEdgeValueForAlignmentSide(artboardRect, alignmentSide);
        if (selectionEdge === null || artboardEdge === null) {
            alert("ALIGNMENT_SIDE の指定が不正です：" + alignmentSide);
            return null;
        }

        var targetEdge = artboardEdge;
        if (USE_GUIDES) {
            var snappedGuideValue = findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide,
                GUIDE_SEARCH_MODE, GUIDE_ORIENTATION_TOLERANCE);
            if (snappedGuideValue !== null) targetEdge = snappedGuideValue;
        }

        var edgeOffset = targetEdge - selectionEdge;
        var isHorizontal = (alignmentSide === "left" || alignmentSide === "right");
        return { x: isHorizontal ? edgeOffset : 0, y: isHorizontal ? 0 : edgeOffset };
    }

    /**
     * 指定した方向に対応する境界値を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @returns {number|null} 境界値。方向が不正なら null
     */
    function getEdgeValueForAlignmentSide(boundsRect, alignmentSide) {
        if (alignmentSide === "left") return boundsRect[0];
        if (alignmentSide === "top") return boundsRect[1];
        if (alignmentSide === "right") return boundsRect[2];
        if (alignmentSide === "bottom") return boundsRect[3];
        return null;
    }

    /**
     * 左右中央の座標を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @returns {number} 左右中央のX
     */
    function getHorizontalCenterValue(boundsRect) {
        return (boundsRect[0] + boundsRect[2]) / 2;
    }

    /**
     * 上下中央の座標を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @returns {number} 上下中央のY
     */
    function getVerticalCenterValue(boundsRect) {
        return (boundsRect[1] + boundsRect[3]) / 2;
    }

    /**
     * アートボード内側にあるガイドのうち、指定した方向と探索範囲に合う吸着先座標を返す
     * @param {Document} documentRef - 対象のドキュメント
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {number} selectionEdge - 選択範囲の該当する端の座標
     * @param {string} alignmentSide - 揃える方向
     * @param {string} guideSearchMode - "inside" | "nearest"
     * @param {number} guideOrientationTolerance - ガイドの水平・垂直判定に使う許容値
     * @returns {number|null} 吸着先の座標。該当するガイドがなければ null
     */
    function findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide, guideSearchMode, guideOrientationTolerance) {
        var nearestGuideValue = null;
        var nearestGuideDistance = null;
        var documentPathItems = documentRef.pathItems;

        for (var i = 0; i < documentPathItems.length; i++) {
            var guidePathItem = documentPathItems[i];
            if (guidePathItem.guides !== true) continue;

            var guideValue = getGuideValueForAlignmentSide(guidePathItem.geometricBounds, alignmentSide, guideOrientationTolerance);
            if (guideValue === null) continue;
            if (!isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide)) continue;

            if (guideSearchMode === "inside") {
                if (!isGuideOnAlignmentSide(guideValue, selectionEdge, alignmentSide)) continue;
                if (nearestGuideValue === null || isGuideCloserFromInside(guideValue, nearestGuideValue, alignmentSide)) {
                    nearestGuideValue = guideValue;
                }

            } else if (guideSearchMode === "nearest") {
                var guideDistance = Math.abs(guideValue - selectionEdge);
                if (nearestGuideDistance === null || guideDistance < nearestGuideDistance) {
                    nearestGuideValue = guideValue;
                    nearestGuideDistance = guideDistance;
                }

            } else {
                alert("GUIDE_SEARCH_MODE の指定が不正です：" + guideSearchMode);
                return null;
            }
        }

        return nearestGuideValue;
    }

    /**
     * 揃える方向に対応するガイド座標を返す
     * @param {number[]} guideBounds - ガイドの境界 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @param {number} guideOrientationTolerance - 水平・垂直判定に使う許容値
     * @returns {number|null} ガイドの座標。向きが対応しなければ null
     */
    function getGuideValueForAlignmentSide(guideBounds, alignmentSide, guideOrientationTolerance) {
        var isVerticalGuide = Math.abs(guideBounds[2] - guideBounds[0]) <= guideOrientationTolerance;
        var isHorizontalGuide = Math.abs(guideBounds[1] - guideBounds[3]) <= guideOrientationTolerance;

        if ((alignmentSide === "left" || alignmentSide === "right") && isVerticalGuide) return guideBounds[0];
        if ((alignmentSide === "top" || alignmentSide === "bottom") && isHorizontalGuide) return guideBounds[1];
        return null;
    }

    /**
     * ガイド座標がアクティブアートボード内にあるかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} 内側にあれば true
     */
    function isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide) {
        if (alignmentSide === "left" || alignmentSide === "right") {
            return guideValue >= artboardRect[0] && guideValue <= artboardRect[2];
        }
        if (alignmentSide === "top" || alignmentSide === "bottom") {
            return guideValue <= artboardRect[1] && guideValue >= artboardRect[3];
        }
        return false;
    }

    /**
     * GUIDE_SEARCH_MODE が inside のとき、揃える方向側にあるガイドかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number} selectionEdge - 選択範囲の該当する端の座標
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} 揃える方向側にあれば true
     */
    function isGuideOnAlignmentSide(guideValue, selectionEdge, alignmentSide) {
        if (alignmentSide === "left") return guideValue < selectionEdge;
        if (alignmentSide === "right") return guideValue > selectionEdge;
        if (alignmentSide === "top") return guideValue > selectionEdge;
        if (alignmentSide === "bottom") return guideValue < selectionEdge;
        return false;
    }

    /**
     * inside 側にある複数のガイドのうち、選択範囲に近い方かを判定する
     * @param {number} guideValue - 比較するガイドの座標
     * @param {number} currentBestGuideValue - 現時点で最も近いガイドの座標
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} より近ければ true
     */
    function isGuideCloserFromInside(guideValue, currentBestGuideValue, alignmentSide) {
        if (alignmentSide === "left") return guideValue > currentBestGuideValue;
        if (alignmentSide === "right") return guideValue < currentBestGuideValue;
        if (alignmentSide === "top") return guideValue < currentBestGuideValue;
        if (alignmentSide === "bottom") return guideValue > currentBestGuideValue;
        return false;
    }

    /**
     * 指定オブジェクトの境界を返す
     * クリッピンググループはマスクパス（先頭アイテム）の geometricBounds を使う。
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {boolean} includeStrokeInBounds - 線を境界に含めるか
     * @returns {number[]} 境界 [L, T, R, B]
     */
    function getItemBounds(pageItem, includeStrokeInBounds) {
        if (pageItem.typename === "GroupItem" && pageItem.clipped === true) {
            return pageItem.pageItems[0].geometricBounds;
        }
        return includeStrokeInBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
    }

    main(arguments[0]);

})();
