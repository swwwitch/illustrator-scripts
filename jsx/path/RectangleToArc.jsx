#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した長方形を、左下隅・上辺中央・右下隅の3点を通る円弧に変換します。
長方形の幅と高さから円の半径・中心・開始角・終了角を求め、元の長方形は削除します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArc.md

### Overview

Converts the selected rectangle into an arc passing through its bottom-left corner, top-center and bottom-right corner.
The radius, center and start and end angles come from the rectangle's width and height, and the rectangle itself is deleted.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangleToArc.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RectangleToArc";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArc.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangleToArc.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 座標比較の許容誤差（pt） / Tolerance for coordinate comparison (pt) */
    var POSITION_TOLERANCE = 0.01;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "長方形を選択してから実行してください。", en: "Select one or more rectangles and run the script again." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文字列を引く
     * @param {string} labelPath - "alert.noSelection" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    // =========================================
    // 判定ユーティリティ / Shape checks
    // =========================================

    /**
     * 数値を許容誤差付きで比較する
     * @param {number} a - 比較する値
     * @param {number} b - 比較する値
     * @returns {boolean} 差が許容誤差以内なら true
     */
    function nearlyEqual(a, b) {
        return Math.abs(a - b) <= POSITION_TOLERANCE;
    }

    /**
     * 2点を結ぶ線分が水平か判定する
     * @param {number[]} pointA - 始点 [x, y]
     * @param {number[]} pointB - 終点 [x, y]
     * @returns {boolean} 水平なら true
     */
    function isHorizontalSegment(pointA, pointB) {
        return nearlyEqual(pointA[1], pointB[1]) && !nearlyEqual(pointA[0], pointB[0]);
    }

    /**
     * 2点を結ぶ線分が垂直か判定する
     * @param {number[]} pointA - 始点 [x, y]
     * @param {number[]} pointB - 終点 [x, y]
     * @returns {boolean} 垂直なら true
     */
    function isVerticalSegment(pointA, pointB) {
        return nearlyEqual(pointA[0], pointB[0]) && !nearlyEqual(pointA[1], pointB[1]);
    }

    /**
     * 閉じた4点パスで、各辺が水平または垂直か（回転していない長方形か）を判定する
     * @param {PageItem} pathItem - 判定するオブジェクト
     * @returns {boolean} 長方形なら true
     */
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
    // 外観設定 / Appearance
    // =========================================

    /**
     * 円弧を塗りなし・線ありにし、元の長方形に線があれば色と太さを引き継ぐ
     * @param {PathItem} sourceRectangle - 元の長方形
     * @param {PathItem} arcPath - 生成した円弧
     * @returns {void}
     */
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
    // データ取得 / Data
    // =========================================

    /**
     * 選択の配列をコピーする（処理中に選択が変わっても影響を受けないように）
     * @param {PageItem[]} selectionItems - 選択中のオブジェクト
     * @returns {PageItem[]} コピーした配列
     */
    function copySelectionItems(selectionItems) {
        var copiedItems = [];
        for (var i = 0; i < selectionItems.length; i++) {
            copiedItems.push(selectionItems[i]);
        }
        return copiedItems;
    }

    /**
     * 長方形の座標と寸法を取得する
     * @param {PathItem} sourceRectangle - 元の長方形
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}|null} 寸法（幅か高さが0なら null）
     */
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
    // 円弧計算 / Arc geometry
    // =========================================

    /**
     * 左下隅・上辺中央・右下隅を通る円の半径・中心・開始角・終了角を求める
     * @param {{left: number, top: number, right: number, bottom: number, width: number, height: number}} rectangleMetrics - 長方形の寸法
     * @returns {{radius: number, centerX: number, centerY: number, startAngle: number, endAngle: number}} 円弧のジオメトリ
     */
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
    // 円弧生成 / Arc creation
    // =========================================

    /**
     * 円弧を90度以内のベジェセグメントに分けてアンカーポイントを追加する
     * @param {PathItem} arcPath - 空のパス
     * @param {{radius: number, centerX: number, centerY: number, startAngle: number, endAngle: number}} arcGeometry - 円弧のジオメトリ
     * @returns {void}
     */
    function addArcPoints(arcPath, arcGeometry) {
        var sweepAngle = arcGeometry.startAngle - arcGeometry.endAngle;
        var segmentCount = Math.ceil(sweepAngle / (Math.PI / 2));
        var angleStep = sweepAngle / segmentCount;
        var handleLength = (4 / 3) * arcGeometry.radius * Math.tan(angleStep / 4);

        for (var segmentIndex = 0; segmentIndex <= segmentCount; segmentIndex++) {
            var currentAngle = arcGeometry.startAngle - segmentIndex * angleStep;
            var pathPoint = arcPath.pathPoints.add();
            var anchorX = arcGeometry.centerX + arcGeometry.radius * Math.cos(currentAngle);
            var anchorY = arcGeometry.centerY + arcGeometry.radius * Math.sin(currentAngle);
            var tangentX = Math.sin(currentAngle);
            var tangentY = -Math.cos(currentAngle);

            pathPoint.anchor = [anchorX, anchorY];
            pathPoint.leftDirection = [anchorX - handleLength * tangentX, anchorY - handleLength * tangentY];
            pathPoint.rightDirection = [anchorX + handleLength * tangentX, anchorY + handleLength * tangentY];
        }
    }

    /**
     * 元の長方形の直前（前面）に円弧のパスを作る
     * @param {PathItem} sourceRectangle - 元の長方形
     * @param {{radius: number, centerX: number, centerY: number, startAngle: number, endAngle: number}} arcGeometry - 円弧のジオメトリ
     * @returns {PathItem} 生成した円弧
     */
    function createArcPathBeforeSource(sourceRectangle, arcGeometry) {
        var arcPath = sourceRectangle.layer.pathItems.add();
        arcPath.move(sourceRectangle, ElementPlacement.PLACEBEFORE);
        addArcPoints(arcPath, arcGeometry);
        return arcPath;
    }

    /**
     * 長方形を円弧に置き換える（元の長方形は削除）
     * @param {PathItem} sourceRectangle - 元の長方形
     * @returns {void}
     */
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
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中の長方形をすべて円弧に変換する
     * @returns {void}
     */
    function main() {
        /* ドキュメント確認 / Check document */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedPageItems = doc.selection;

        /* 選択確認 / Check selection */
        if (selectedPageItems.length === 0) {
            alert(getLabel("alert.noSelection"));
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
