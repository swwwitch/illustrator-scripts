#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した長方形を、左下隅・上辺中央・右下隅の3点を通る円弧に変換します。
長方形の幅と高さから円の半径・中心・開始角・終了角を求め、元の長方形は削除します。

詳細は README を参照してください。

### Overview

Converts the selected rectangle into an arc passing through its bottom-left corner, top-center and bottom-right corner.
The radius, center and start and end angles come from the rectangle's width and height, and the rectangle itself is deleted.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RectangleToArc";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-19";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArc.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangleToArc.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // バージョンと定数
    // =========================================

    var POSITION_TOLERANCE = 0.01;

    // =========================================
    // 判定ユーティリティ
    // =========================================
    /* 数値を許容誤差付きで比較 / Compare numbers with tolerance */
    function nearlyEqual(a, b) {
        return Math.abs(a - b) <= POSITION_TOLERANCE;
    }

    /* 水平セグメント判定 / Check horizontal segment */
    function isHorizontalSegment(pointA, pointB) {
        return nearlyEqual(pointA[1], pointB[1]) && !nearlyEqual(pointA[0], pointB[0]);
    }

    /* 垂直セグメント判定 / Check vertical segment */
    function isVerticalSegment(pointA, pointB) {
        return nearlyEqual(pointA[0], pointB[0]) && !nearlyEqual(pointA[1], pointB[1]);
    }

    /* 長方形パス判定 / Check rectangle path item */
    function isRectanglePathItem(pathItem) {
        if (!pathItem || pathItem.typename !== "PathItem") {
            return false;
        }
        if (!pathItem.closed || pathItem.pathPoints.length !== 4) {
            return false;
        }

        var points = [];
        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            points.push(pathItem.pathPoints[i].anchor);
        }

        var hasHorizontalSegment = false;
        var hasVerticalSegment = false;
        for (var j = 0; j < points.length; j++) {
            var currentPoint = points[j];
            var nextPoint = points[(j + 1) % points.length];

            if (isHorizontalSegment(currentPoint, nextPoint)) {
                hasHorizontalSegment = true;
                continue;
            }
            if (isVerticalSegment(currentPoint, nextPoint)) {
                hasVerticalSegment = true;
                continue;
            }
            return false;
        }

        return hasHorizontalSegment && hasVerticalSegment;
    }

    // =========================================
    // 外観設定
    // =========================================
    /* 円弧の外観設定 / Set arc appearance */
    function setArcAppearance(sourceRectangle, arcPath) {
        arcPath.filled = false;
        arcPath.stroked = true;

        if (!sourceRectangle.stroked) {
            return;
        }

        arcPath.strokeColor = sourceRectangle.strokeColor;
        arcPath.strokeWidth = sourceRectangle.strokeWidth;
    }

    // =========================================
    // データ取得
    // =========================================
    /* 選択配列をコピー / Copy selection items */
    function copySelectionItems(selectionItems) {
        var copiedItems = [];
        for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
            copiedItems.push(selectionItems[selectionIndex]);
        }
        return copiedItems;
    }

    /* 長方形情報を取得 / Get rectangle metrics */
    function getRectangleMetrics(sourceRectangle) {
        var bounds = sourceRectangle.geometricBounds;
        var width = Math.abs(bounds[2] - bounds[0]);
        var height = Math.abs(bounds[1] - bounds[3]);

        if (width === 0 || height === 0) {
            return null;
        }

        return {
            left: bounds[0],
            top: bounds[1],
            right: bounds[2],
            bottom: bounds[3],
            width: width,
            height: height
        };
    }

    // =========================================
    // 円弧計算
    // =========================================
    /* 円弧ジオメトリを計算 / Calculate arc geometry */
    function getArcGeometry(rectangleMetrics) {
        var arcRadius = (rectangleMetrics.width * rectangleMetrics.width) / (8 * rectangleMetrics.height) + (rectangleMetrics.height / 2);
        var centerX = rectangleMetrics.left + rectangleMetrics.width / 2;
        var centerY = rectangleMetrics.top - arcRadius;
        var bottomOffsetY = rectangleMetrics.bottom - centerY;
        var startAngle = Math.atan2(bottomOffsetY, rectangleMetrics.left - centerX);
        var endAngle = Math.atan2(bottomOffsetY, rectangleMetrics.right - centerX);

        while (startAngle < Math.PI / 2) {
            startAngle += 2 * Math.PI;
        }

        while (endAngle > Math.PI / 2) {
            endAngle -= 2 * Math.PI;
        }

        return {
            radius: arcRadius,
            centerX: centerX,
            centerY: centerY,
            startAngle: startAngle,
            endAngle: endAngle
        };
    }

    // =========================================
    // 円弧生成
    // =========================================
    /* 円弧ポイントを追加 / Add arc path points */
    function addArcPoints(arcPath, arcGeometry) {
        var numSegments = Math.ceil((arcGeometry.startAngle - arcGeometry.endAngle) / (Math.PI / 2));
        var angleStep = (arcGeometry.startAngle - arcGeometry.endAngle) / numSegments;

        for (var segmentIndex = 0; segmentIndex <= numSegments; segmentIndex++) {
            var currentAngle = arcGeometry.startAngle - segmentIndex * angleStep;
            var pathPoint = arcPath.pathPoints.add();
            var anchorX = arcGeometry.centerX + arcGeometry.radius * Math.cos(currentAngle);
            var anchorY = arcGeometry.centerY + arcGeometry.radius * Math.sin(currentAngle);
            var handleLength = (4 / 3) * arcGeometry.radius * Math.tan(angleStep / 4);
            var tangentX = Math.sin(currentAngle);
            var tangentY = -Math.cos(currentAngle);

            pathPoint.anchor = [anchorX, anchorY];
            pathPoint.leftDirection = [anchorX - handleLength * tangentX, anchorY - handleLength * tangentY];
            pathPoint.rightDirection = [anchorX + handleLength * tangentX, anchorY + handleLength * tangentY];
        }
    }

    /* 元オブジェクト前に円弧を作成 / Create arc before source item */
    function createArcPathBeforeSource(sourceRectangle, arcGeometry) {
        var arcPath = sourceRectangle.layer.pathItems.add();
        arcPath.move(sourceRectangle, ElementPlacement.PLACEBEFORE);
        addArcPoints(arcPath, arcGeometry);
        return arcPath;
    }

    /* 長方形を円弧へ変換 / Convert rectangle to arc */
    function convertRectangleToArc(sourceRectangle) {
        var rectangleMetrics = getRectangleMetrics(sourceRectangle);
        if (!rectangleMetrics) {
            return;
        }

        var arcGeometry = getArcGeometry(rectangleMetrics);
        var arcPath = createArcPathBeforeSource(sourceRectangle, arcGeometry);
        setArcAppearance(sourceRectangle, arcPath);
        sourceRectangle.remove();
    }

    // =========================================
    // メイン処理
    // =========================================
    /* メイン処理 / Main process */
    function main() {
        // ドキュメント確認 / Check document
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてください。");
            return;
        }

        var doc = app.activeDocument;
        var selectedPageItems = doc.selection;

        // 選択確認 / Check selection
        if (selectedPageItems.length === 0) {
            alert("長方形を選択してから実行してください。");
            return;
        }

        var sourceItems = copySelectionItems(selectedPageItems);
        for (var itemIndex = 0; itemIndex < sourceItems.length; itemIndex++) {
            var sourceRectangle = sourceItems[itemIndex];
            if (!isRectanglePathItem(sourceRectangle)) {
                continue;
            }
            convertRectangleToArc(sourceRectangle);
        }
    }

    main();

})();
