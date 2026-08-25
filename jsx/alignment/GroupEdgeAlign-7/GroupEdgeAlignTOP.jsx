#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択されているオブジェクト群の端または中心を取得し、上に揃えます。
揃え先は、アクティブアートボードの端、または条件に合うガイドです。

詳細は README を参照してください。

### Overview

Takes the edges or the center of the selected objects and aligns them to the top edge.
The target is the edge of the active artboard, or a matching guide.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEdgeAlignTOP";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-04-06";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignTOP.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlignTOP.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function main() {
        try {
            var documentRef = app.activeDocument;
            var selectedItems = documentRef.selection;
            var USE_GUIDES = true; // ガイドを使用するかどうかのスイッチ
            // ファイル名から揃え方向を自動判定（例: GroupEdgeAlignRIGHT.jsx → "right"）
            var scriptFileName = File($.fileName).name;
            var fileNameUpper = scriptFileName.toUpperCase();

            var ALIGNMENT_SIDE = "left"; // デフォルト

            if (fileNameUpper.indexOf("CENTERX") !== -1) {
                ALIGNMENT_SIDE = "CENTER_X";
            } else if (fileNameUpper.indexOf("CENTERY") !== -1) {
                ALIGNMENT_SIDE = "CENTER_Y";
            } else if (fileNameUpper.indexOf("CENTER") !== -1) {
                ALIGNMENT_SIDE = "CENTER";
            } else if (fileNameUpper.indexOf("LEFT") !== -1) {
                ALIGNMENT_SIDE = "left";
            } else if (fileNameUpper.indexOf("RIGHT") !== -1) {
                ALIGNMENT_SIDE = "right";
            } else if (fileNameUpper.indexOf("TOP") !== -1) {
                ALIGNMENT_SIDE = "top";
            } else if (fileNameUpper.indexOf("BOTTOM") !== -1) {
                ALIGNMENT_SIDE = "bottom";
            }
            var BOUNDS_MODE = "preference"; // "preference" | "preview" | "geometric"
            var GUIDE_SEARCH_MODE = "inside"; // "inside" | "nearest"
            var GUIDE_ORIENTATION_TOLERANCE = 0.01; // ガイドの水平・垂直判定に使う許容値

            // 選択されていない場合は処理中止
            if (selectedItems.length === 0) {
                alert("オブジェクトが選択されていません。");
                return;
            }

            var artboards = documentRef.artboards;
            var activeArtboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
            var includeStrokeInBounds;

            if (BOUNDS_MODE === "preference") {
                includeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

            } else if (BOUNDS_MODE === "preview") {
                includeStrokeInBounds = true;

            } else if (BOUNDS_MODE === "geometric") {
                includeStrokeInBounds = false;

            } else {
                alert("BOUNDS_MODE の指定が不正です：" + BOUNDS_MODE);
                return;
            }

            var selectionMinX, selectionMaxY, selectionMaxX, selectionMinY;

            // オブジェクト全体のバウンディングボックス範囲を取得
            for (var selectedItemIndex = 0; selectedItemIndex < selectedItems.length; selectedItemIndex++) {
                var selectedItemBounds = getItemBounds(selectedItems[selectedItemIndex], includeStrokeInBounds);

                if (selectedItemIndex === 0) {
                    selectionMinX = selectedItemBounds[0];
                    selectionMaxY = selectedItemBounds[1];
                    selectionMaxX = selectedItemBounds[2];
                    selectionMinY = selectedItemBounds[3];
                } else {
                    if (selectedItemBounds[0] < selectionMinX) selectionMinX = selectedItemBounds[0];
                    if (selectedItemBounds[1] > selectionMaxY) selectionMaxY = selectedItemBounds[1];
                    if (selectedItemBounds[2] > selectionMaxX) selectionMaxX = selectedItemBounds[2];
                    if (selectedItemBounds[3] < selectionMinY) selectionMinY = selectedItemBounds[3];
                }
            }

            var selectionBounds = [selectionMinX, selectionMaxY, selectionMaxX, selectionMinY];
            var offsetX = 0;
            var offsetY = 0;

            if (ALIGNMENT_SIDE === "CENTER_X") {
                offsetX = getHorizontalCenterValue(activeArtboardRect) - getHorizontalCenterValue(selectionBounds);

            } else if (ALIGNMENT_SIDE === "CENTER_Y") {
                offsetY = getVerticalCenterValue(activeArtboardRect) - getVerticalCenterValue(selectionBounds);

            } else if (ALIGNMENT_SIDE === "CENTER") {
                offsetX = getHorizontalCenterValue(activeArtboardRect) - getHorizontalCenterValue(selectionBounds);
                offsetY = getVerticalCenterValue(activeArtboardRect) - getVerticalCenterValue(selectionBounds);

            } else {
                var selectionEdge = getEdgeValueForAlignmentSide(selectionBounds, ALIGNMENT_SIDE);
                var artboardEdge = getEdgeValueForAlignmentSide(activeArtboardRect, ALIGNMENT_SIDE);

                if (selectionEdge === null || artboardEdge === null) {
                    alert("ALIGNMENT_SIDE の指定が不正です：" + ALIGNMENT_SIDE);
                    return;
                }

                var targetEdge;
                if (USE_GUIDES) {
                    var snappedGuideValue = findGuideSnapValue(documentRef, activeArtboardRect, selectionEdge, ALIGNMENT_SIDE, GUIDE_SEARCH_MODE, GUIDE_ORIENTATION_TOLERANCE);
                    targetEdge = (snappedGuideValue !== null) ? snappedGuideValue : artboardEdge;
                } else {
                    targetEdge = artboardEdge;
                }

                var offset = targetEdge - selectionEdge;
                offsetX = (ALIGNMENT_SIDE === "left" || ALIGNMENT_SIDE === "right") ? offset : 0;
                offsetY = (ALIGNMENT_SIDE === "top" || ALIGNMENT_SIDE === "bottom") ? offset : 0;
            }

            // オブジェクトを移動し、選択解除
            for (var moveItemIndex = 0; moveItemIndex < selectedItems.length; moveItemIndex++) {
                selectedItems[moveItemIndex].translate(offsetX, offsetY);
                // selectedItems[moveItemIndex].selected = false;
            }

        } catch (e) {
            alert("エラーが発生しました：" + e.message);
        }
    }

    // 指定した方向に対応する境界値を返す（不正な方向の場合は null）
    function getEdgeValueForAlignmentSide(bounds, alignmentSide) {
        if (alignmentSide === "left") return bounds[0];
        if (alignmentSide === "top") return bounds[1];
        if (alignmentSide === "right") return bounds[2];
        if (alignmentSide === "bottom") return bounds[3];
        return null;
    }

    // 左右中央の座標を返す
    function getHorizontalCenterValue(bounds) {
        return (bounds[0] + bounds[2]) / 2;
    }

    // 上下中央の座標を返す
    function getVerticalCenterValue(bounds) {
        return (bounds[1] + bounds[3]) / 2;
    }

    // アートボード内側にあるガイドのうち、指定した方向と探索範囲に合う吸着先座標を返す（なければ null）
    function findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide, guideSearchMode, guideOrientationTolerance) {
        var nearestGuideValue = null;
        var nearestGuideDistance = null;
        var guidePathItems = documentRef.pathItems;

        for (var guidePathIndex = 0; guidePathIndex < guidePathItems.length; guidePathIndex++) {
            var guidePathItem = guidePathItems[guidePathIndex];
            if (guidePathItem.guides !== true) continue;

            var guideBounds = guidePathItem.geometricBounds; // [L, T, R, B]
            var guideValue = getGuideValueForAlignmentSide(guideBounds, alignmentSide, guideOrientationTolerance);
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

    // ALIGNMENT_SIDE に対応するガイド座標を返す（対応しない向きのガイドは null）
    function getGuideValueForAlignmentSide(guideBounds, alignmentSide, guideOrientationTolerance) {
        var isVerticalGuide = Math.abs(guideBounds[2] - guideBounds[0]) <= guideOrientationTolerance;
        var isHorizontalGuide = Math.abs(guideBounds[1] - guideBounds[3]) <= guideOrientationTolerance;

        if ((alignmentSide === "left" || alignmentSide === "right") && isVerticalGuide) {
            return guideBounds[0];
        }

        if ((alignmentSide === "top" || alignmentSide === "bottom") && isHorizontalGuide) {
            return guideBounds[1];
        }

        return null;
    }

    // ガイド座標がアクティブアートボード内にあるかを返す
    function isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide) {
        if (alignmentSide === "left" || alignmentSide === "right") {
            return guideValue >= artboardRect[0] && guideValue <= artboardRect[2];
        }

        if (alignmentSide === "top" || alignmentSide === "bottom") {
            return guideValue <= artboardRect[1] && guideValue >= artboardRect[3];
        }

        return false;
    }

    // GUIDE_SEARCH_MODE が inside のとき、揃える方向側にあるガイドかを返す
    function isGuideOnAlignmentSide(guideValue, selectionEdge, alignmentSide) {
        if (alignmentSide === "left") return guideValue < selectionEdge;
        if (alignmentSide === "right") return guideValue > selectionEdge;
        if (alignmentSide === "top") return guideValue > selectionEdge;
        if (alignmentSide === "bottom") return guideValue < selectionEdge;
        return false;
    }

    // inside 側にある複数のガイドのうち、選択範囲に最も近いかを判定する
    function isGuideCloserFromInside(guideValue, currentBestGuideValue, alignmentSide) {
        if (alignmentSide === "left") return guideValue > currentBestGuideValue;
        if (alignmentSide === "right") return guideValue < currentBestGuideValue;
        if (alignmentSide === "top") return guideValue < currentBestGuideValue;
        if (alignmentSide === "bottom") return guideValue > currentBestGuideValue;
        return false;
    }

    // 指定オブジェクトの境界を取得する関数
    function getItemBounds(pageItem, includeStrokeInBounds) {
        var itemBounds;

        // クリッピンググループの場合、最初のアイテムのgeometricBoundsを使用
        if (pageItem.typename === "GroupItem" && pageItem.clipped === true) {
            itemBounds = pageItem.pageItems[0].geometricBounds;

        } else if (includeStrokeInBounds === true) {
            // 線を含める設定の場合、visibleBoundsを使用
            itemBounds = pageItem.visibleBounds;

        } else {
            // 線を含めない場合、geometricBoundsを使用
            itemBounds = pageItem.geometricBounds;
        }

        return itemBounds;
    }

    main();

})();
