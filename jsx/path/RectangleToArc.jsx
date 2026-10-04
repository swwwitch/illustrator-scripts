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
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "長方形を選択してから実行してください。", en: "Select one or more rectangles and run the script again." }
        }
    };

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
