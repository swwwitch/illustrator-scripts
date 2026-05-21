#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

RectangleToArc.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArc.md

### 概要：

- 更新日：2026-05-19
- 選択した長方形を、左下隅・上辺中央・右下隅の 3 点を通る円弧に変換
- 長方形の幅と高さから円の半径・中心・開始角・終了角を求め、90 度以内のベジェセグメントに分割して円弧を生成
- 元の長方形は削除

### 主な機能：

- 閉じた 4 点パスかつ各辺が水平／垂直の長方形のみ処理（回転長方形・台形・菱形・不定形は対象外）
- 生成する円弧は塗りなし・線あり
- 元長方形に線がある場合は線の色・太さを引き継ぎ
- 元長方形が線なしでも変換は実行
- 複数選択に対応

### 更新履歴：

- v1.0.1 (2026-05-19) : 現行版

*/

/*

### Script Name:

RectangleToArc.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangleToArc.md

### Description:

- Last Updated: 2026-05-19
- Converts each selected rectangle into a circular arc passing through its bottom-left corner, top-edge midpoint, and bottom-right corner
- Derives the radius, center, start angle, and end angle from the rectangle's width and height, then splits the arc into bezier segments of 90 degrees or less
- The original rectangle is deleted

### Main Features:

- Only closed 4-point paths with horizontal/vertical edges are processed (rotated rectangles, trapezoids, rhombi, and other 4-point shapes are skipped)
- Generated arc is fill-less and stroked
- Inherits stroke color/width from the source rectangle when present
- Conversion runs even if the source rectangle has no stroke
- Multiple selection supported

### Changelog:

- v1.0.1 (2026-05-19) : Current version

*/

(function () {

    // =========================================
    // バージョンと定数
    // =========================================
    var SCRIPT_VERSION = "v1.0.1";

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