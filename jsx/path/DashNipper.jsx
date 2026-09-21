#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

破線を選択して実行すると、アウトライン化して1つずつの線分に分け、それぞれを同じ見た目の線（中心線）に置き換えます。
アウトライン化済みの図形（アンカーポイントが4つの閉じたパス）を1つ選択したときは、その図形を中心線に置き換えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashNipper.md

### Overview

With a dashed line selected, outlines it, splits it into its dashes, and replaces each dash with a line that looks the same (its center line).
With one already-outlined shape selected (a closed path with four anchor points), replaces that shape with its center line.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DashNipper.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DashNipper";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashNipper.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DashNipper.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneItem: { ja: "長方形を1つだけ選択してください。", en: "Select only one rectangle." },
            notClosedPath: { ja: "閉じた長方形を1つ選択してください。", en: "Select one closed rectangle." },
            notFourAnchors: {
                ja: "4つのアンカーポイントを持つ長方形を選択してください。",
                en: "Select a rectangle with four anchor points."
            },
            dashOutlineFailed: { ja: "破線をアウトライン化できませんでした。", en: "Could not outline the dashed line." },
            skippedPieces: {
                ja: "中心線にできなかったパーツが {count} 個あります。アウトラインのまま残しました。",
                en: "{count} piece(s) could not be converted to a center line and were left as outlines."
            }
        }
    };

    /**
     * UI 言語に合った文言を返す
     * @param {object} labelEntry - ja / en を持つラベル定義
     * @returns {string} 表示する文言
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // 幾何計算 / Geometry
    // =========================================
    var LENGTH_TOLERANCE = 0.001;         /* 太さを同じとみなす差（pt）/ max difference treated as equal thickness (pt) */
    var STRAIGHT_TOLERANCE = 0.01;        /* 直線とみなすハンドルのずれ（pt）/ max handle offset treated as straight (pt) */
    var LENGTH_MISMATCH_RATIO = 0.01;     /* 長さを一致とみなす差（比率）/ allowed length mismatch (ratio) */
    var BEZIER_LENGTH_STEPS = 16;         /* 曲線の長さを測る分割数 / subdivisions used to measure a curve */
    var ARC_CHECK_STEPS = 4;              /* 半円かを確かめる、1セグメントあたりの点の数 / points checked per segment for a semicircle */
    var MAX_CAP_SEGMENTS = 4;             /* 丸型線端の半円を作るセグメント数の上限 / max segments forming a round cap */

    /**
     * @typedef {object} AnchorInfo
     * @property {number[]} anchor - アンカーポイントの座標 [x, y]
     * @property {number[]} leftDirection - 前のセグメント側のハンドル [x, y]
     * @property {number[]} rightDirection - 次のセグメント側のハンドル [x, y]
     */

    /**
     * @typedef {object} CenterLine
     * @property {AnchorInfo[]} points - 中心線のアンカー（始点から終点の順）
     * @property {number} thickness - 元の図形の太さ（両端の辺の長さ、または半円の直径の平均、pt）
     * @property {boolean} roundCap - 両端が半円（ピル形状）なら true。線端を丸型にする
     */

    /**
     * 2点の中点を返す
     * @param {number[]} pointA - 座標 [x, y]
     * @param {number[]} pointB - 座標 [x, y]
     * @returns {number[]} 中点 [x, y]
     */
    function getMidpoint(pointA, pointB) {
        return [(pointA[0] + pointB[0]) / 2, (pointA[1] + pointB[1]) / 2];
    }

    /**
     * 2点間の距離を返す
     * @param {number[]} pointA - 座標 [x, y]
     * @param {number[]} pointB - 座標 [x, y]
     * @returns {number} 距離（pt）
     */
    function getDistance(pointA, pointB) {
        var dx = pointB[0] - pointA[0];
        var dy = pointB[1] - pointA[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 点から直線までの距離を返す
     * @param {number[]} point - 測る点 [x, y]
     * @param {number[]} lineStart - 直線上の1点 [x, y]
     * @param {number[]} lineEnd - 直線上のもう1点 [x, y]
     * @returns {number} 距離（pt）。2点が重なるときは lineStart からの距離
     */
    function getDistanceToLine(point, lineStart, lineEnd) {
        var dx = lineEnd[0] - lineStart[0];
        var dy = lineEnd[1] - lineStart[1];
        var lineLength = Math.sqrt(dx * dx + dy * dy);
        var offsetX = point[0] - lineStart[0];
        var offsetY = point[1] - lineStart[1];
        if (lineLength === 0) return Math.sqrt(offsetX * offsetX + offsetY * offsetY);
        return Math.abs(dx * offsetY - dy * offsetX) / lineLength;
    }

    /**
     * 3次ベジェ曲線上の点を返す
     * @param {number[]} start - 始点 [x, y]
     * @param {number[]} startHandle - 始点のハンドル [x, y]
     * @param {number[]} endHandle - 終点のハンドル [x, y]
     * @param {number[]} end - 終点 [x, y]
     * @param {number} t - 曲線上の位置（0〜1）
     * @returns {number[]} 曲線上の点 [x, y]
     */
    function getPointOnBezier(start, startHandle, endHandle, end, t) {
        var inverseT = 1 - t;
        var startWeight = inverseT * inverseT * inverseT;
        var startHandleWeight = 3 * inverseT * inverseT * t;
        var endHandleWeight = 3 * inverseT * t * t;
        var endWeight = t * t * t;
        return [
            startWeight * start[0] + startHandleWeight * startHandle[0] + endHandleWeight * endHandle[0] + endWeight * end[0],
            startWeight * start[1] + startHandleWeight * startHandle[1] + endHandleWeight * endHandle[1] + endWeight * end[1]
        ];
    }

    /**
     * パスの1辺（アンカーから次のアンカーまで）の長さを、折れ線で近似して返す
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} sideIndex - 辺の始点になるアンカーの番号
     * @returns {number} 辺の長さ（pt）
     */
    function getSideLength(anchors, sideIndex) {
        var from = anchors[sideIndex % anchors.length];
        var to = anchors[(sideIndex + 1) % anchors.length];
        var length = 0;
        var previousPoint = from.anchor;
        for (var i = 1; i <= BEZIER_LENGTH_STEPS; i++) {
            var point = getPointOnBezier(from.anchor, from.rightDirection, to.leftDirection, to.anchor, i / BEZIER_LENGTH_STEPS);
            length += getDistance(previousPoint, point);
            previousPoint = point;
        }
        return length;
    }

    /**
     * パスの1辺が直線かどうかを返す（ハンドルが両端を結ぶ線上にあれば直線）
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} sideIndex - 辺の始点になるアンカーの番号
     * @returns {boolean} 直線なら true
     */
    function isStraightSide(anchors, sideIndex) {
        var from = anchors[sideIndex % anchors.length];
        var to = anchors[(sideIndex + 1) % anchors.length];
        return getDistanceToLine(from.rightDirection, from.anchor, to.anchor) <= STRAIGHT_TOLERANCE &&
            getDistanceToLine(to.leftDirection, from.anchor, to.anchor) <= STRAIGHT_TOLERANCE;
    }

    /**
     * 2つの長さが一致するとみなせるかを返す
     * @param {number} length - 比べる長さ（pt）
     * @param {number} referenceLength - 基準の長さ（pt）
     * @returns {boolean} 差が基準の LENGTH_MISMATCH_RATIO 以内（最小 LENGTH_TOLERANCE）なら true
     */
    function isMatchingLength(length, referenceLength) {
        return Math.abs(length - referenceLength) <= Math.max(LENGTH_TOLERANCE, referenceLength * LENGTH_MISMATCH_RATIO);
    }

    /**
     * 向かい合う2辺が、線分の両端（破線の切り口）として扱える形かどうかを返す
     * （どちらも直線で、長さがそろっていること）
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} endSideIndex - 一方の端の辺の番号（もう一方は半周先）
     * @returns {boolean} 両端として扱えれば true
     */
    function hasMatchingEnds(anchors, endSideIndex) {
        var oppositeSideIndex = endSideIndex + anchors.length / 2;
        if (!isStraightSide(anchors, endSideIndex) || !isStraightSide(anchors, oppositeSideIndex)) return false;
        return isMatchingLength(getSideLength(anchors, oppositeSideIndex), getSideLength(anchors, endSideIndex));
    }

    /**
     * 向かい合う2本の辺を、片方は順方向、もう片方は逆方向にたどり、同じ位置のアンカーの中点を結んだ中心線のアンカーを作る
     * （同心円の帯なら、ちょうど中間の半径の円弧になる）
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} forwardStartIndex - 順方向にたどる辺の始点のアンカー番号
     * @param {number} backwardStartIndex - 逆方向にたどる辺の始点のアンカー番号
     * @param {number} pointCount - 中心線のアンカーの数（辺のセグメント数 + 1）
     * @returns {AnchorInfo[]} 中心線のアンカー
     */
    function buildMidlinePoints(anchors, forwardStartIndex, backwardStartIndex, pointCount) {
        var anchorCount = anchors.length;
        var points = [];
        for (var k = 0; k < pointCount; k++) {
            var forwardAnchor = anchors[(forwardStartIndex + k) % anchorCount];
            var backwardAnchor = anchors[((backwardStartIndex - k) % anchorCount + anchorCount) % anchorCount];
            points.push({
                anchor: getMidpoint(forwardAnchor.anchor, backwardAnchor.anchor),
                leftDirection: getMidpoint(forwardAnchor.leftDirection, backwardAnchor.rightDirection),
                rightDirection: getMidpoint(forwardAnchor.rightDirection, backwardAnchor.leftDirection)
            });
        }
        /* 両端の外側のハンドルは畳む / Collapse the outer handles at both ends */
        points[0].leftDirection = points[0].anchor;
        points[pointCount - 1].rightDirection = points[pointCount - 1].anchor;
        return points;
    }

    /**
     * 向かい合う2辺を線端なしの両端とみなし、その間の2本の辺を平均した形の中心線を作る
     * @param {AnchorInfo[]} anchors - パスのアンカー情報（偶数個）
     * @param {number} endSideIndex - 一方の端の辺の番号（もう一方は半周先）
     * @returns {CenterLine} 中心線
     */
    function buildCenterLine(anchors, endSideIndex) {
        var halfCount = anchors.length / 2;
        return {
            points: buildMidlinePoints(anchors, endSideIndex + 1, endSideIndex, halfCount),
            thickness: (getSideLength(anchors, endSideIndex) + getSideLength(anchors, endSideIndex + halfCount)) / 2,
            roundCap: false
        };
    }

    /**
     * 連続するセグメントが半円（両端を直径とする円の上）になっているかを返す
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} startIndex - 半円の始点のアンカー番号
     * @param {number} segmentCount - 半円を作るセグメントの数
     * @returns {boolean} 半円なら true
     */
    function isSemicircleChain(anchors, startIndex, segmentCount) {
        var anchorCount = anchors.length;
        var chainStart = anchors[startIndex % anchorCount].anchor;
        var chainEnd = anchors[(startIndex + segmentCount) % anchorCount].anchor;
        var center = getMidpoint(chainStart, chainEnd);
        var radius = getDistance(chainStart, chainEnd) / 2;
        if (radius <= LENGTH_TOLERANCE) return false;
        for (var segment = 0; segment < segmentCount; segment++) {
            var from = anchors[(startIndex + segment) % anchorCount];
            var to = anchors[(startIndex + segment + 1) % anchorCount];
            for (var i = 1; i <= ARC_CHECK_STEPS; i++) {
                var point = getPointOnBezier(from.anchor, from.rightDirection, to.leftDirection, to.anchor, i / ARC_CHECK_STEPS);
                if (!isMatchingLength(getDistance(center, point), radius)) return false;
            }
        }
        return true;
    }

    /**
     * 両端が半円の図形（ピル形状）なら、両端の半円の中心を結ぶ中心線を作る
     * 半円どうしの間の2本の辺は、セグメントの数がそろっているものだけを受け付ける
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} [expectedThickness] - 太さがわかっていれば指定（pt）。半円の直径と一致するものだけを使う
     * @returns {CenterLine|null} 線端を丸型にする中心線。ピル形状でなければ null
     */
    function buildPillCenterLine(anchors, expectedThickness) {
        var anchorCount = anchors.length;
        for (var capStart = 0; capStart < anchorCount; capStart++) {
            for (var capSegments = 1; capSegments <= MAX_CAP_SEGMENTS; capSegments++) {
                if (!isSemicircleChain(anchors, capStart, capSegments)) continue;
                var capDiameter = getDistance(anchors[capStart].anchor, anchors[(capStart + capSegments) % anchorCount].anchor);
                if (expectedThickness !== undefined && !isMatchingLength(capDiameter, expectedThickness)) continue;

                for (var otherCapSegments = 1; otherCapSegments <= MAX_CAP_SEGMENTS; otherCapSegments++) {
                    var sideSegments = (anchorCount - capSegments - otherCapSegments) / 2;
                    if (sideSegments < 1 || sideSegments !== Math.floor(sideSegments)) continue;
                    var otherCapStart = capStart + capSegments + sideSegments;
                    if (!isSemicircleChain(anchors, otherCapStart, otherCapSegments)) continue;
                    var otherCapDiameter = getDistance(anchors[otherCapStart % anchorCount].anchor,
                        anchors[(otherCapStart + otherCapSegments) % anchorCount].anchor);
                    if (!isMatchingLength(otherCapDiameter, capDiameter)) continue;
                    return {
                        points: buildMidlinePoints(anchors, capStart + capSegments, capStart, sideSegments + 1),
                        thickness: (capDiameter + otherCapDiameter) / 2,
                        roundCap: true
                    };
                }
            }
        }
        return null;
    }

    /** 
     * 中心線が縦寄りかどうかを返す
     * @param {CenterLine} centerLine - 中心線
     * @returns {boolean} 始点と終点の差が横より縦に大きければ true
     */
    function isMostlyVertical(centerLine) {
        var start = centerLine.points[0].anchor;
        var end = centerLine.points[centerLine.points.length - 1].anchor;
        return Math.abs(end[1] - start[1]) > Math.abs(end[0] - start[0]);
    }

    /**
     * 中心線の向きを、横寄りなら左→右、縦寄りなら上→下にそろえる
     * @param {CenterLine} centerLine - 中心線
     * @returns {CenterLine} 向きをそろえた中心線
     */
    function orientCenterLine(centerLine) {
        var start = centerLine.points[0].anchor;
        var end = centerLine.points[centerLine.points.length - 1].anchor;
        var isReversed = isMostlyVertical(centerLine) ? start[1] < end[1] : start[0] > end[0];
        if (!isReversed) return centerLine;
        var reversedPoints = [];
        for (var i = centerLine.points.length - 1; i >= 0; i--) {
            reversedPoints.push({
                anchor: centerLine.points[i].anchor,
                leftDirection: centerLine.points[i].rightDirection,
                rightDirection: centerLine.points[i].leftDirection
            });
        }
        return { points: reversedPoints, thickness: centerLine.thickness, roundCap: centerLine.roundCap };
    }

    /**
     * 閉じたパスのアンカーから中心線を求める
     * 両端が半円（ピル形状）なら線端を丸型にする中心線を使う。
     * そうでなければ、両端の候補のうち太さ（両端の辺の長さ）が最も小さいものを使う。
     * アンカーが4つのときはどの形でも受け付け、それより多いときは両端が直線で長さのそろうものだけを受け付ける
     * @param {AnchorInfo[]} anchors - パスのアンカー情報
     * @param {number} [expectedThickness] - 太さがわかっていれば指定（pt）。一致する候補だけを使う
     *     （破線の線分は線幅より短いことがあり、細い方を選ぶと向きを取り違えるため）
     * @returns {CenterLine|null} 中心線。求められなければ null
     */
    function getCenterLine(anchors, expectedThickness) {
        var anchorCount = anchors.length;
        if (anchorCount < 4) return null;
        /* ピル形状の長辺は直線で長さがそろうため、線端なしの判定より先に見る
           Check pills first: their straight long sides would also pass as butt ends */
        var pillLine = buildPillCenterLine(anchors, expectedThickness);
        if (pillLine) return orientCenterLine(pillLine);
        if (anchorCount % 2 !== 0) return null;

        var bestLine = null;
        for (var endSideIndex = 0; endSideIndex < anchorCount / 2; endSideIndex++) {
            if (anchorCount > 4 && !hasMatchingEnds(anchors, endSideIndex)) continue;
            var candidate = buildCenterLine(anchors, endSideIndex);
            if (expectedThickness !== undefined && !isMatchingLength(candidate.thickness, expectedThickness)) continue;
            if (!bestLine || isBetterCenterLine(candidate, bestLine)) bestLine = candidate;
        }
        return bestLine ? orientCenterLine(bestLine) : null;
    }

    /**
     * 中心線の候補が、いまの候補より適しているかを返す
     * @param {CenterLine} candidate - 比べる候補
     * @param {CenterLine} currentBest - いまの候補
     * @returns {boolean} candidate の方が細ければ true。太さが同じ（正方形など）なら縦の線を優先
     */
    function isBetterCenterLine(candidate, currentBest) {
        if (Math.abs(candidate.thickness - currentBest.thickness) <= LENGTH_TOLERANCE) {
            return isMostlyVertical(candidate) && !isMostlyVertical(currentBest);
        }
        return candidate.thickness < currentBest.thickness;
    }

    // =========================================
    // 選択と描画 / Selection and drawing
    // =========================================

    /**
     * アイテムの中から指定の種類のものを集める（グループの中もたどる）
     * @param {PageItem[]} items - 対象のアイテム
     * @param {string} typename - 集める種類（"PathItem" など）
     * @returns {PageItem[]} 見つかったアイテム
     */
    function collectItemsByType(items, typename) {
        var foundItems = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            if (item.typename === typename) {
                foundItems.push(item);
            } else if (item.typename === "GroupItem") {
                foundItems = foundItems.concat(collectItemsByType(item.pageItems, typename));
            }
        }
        return foundItems;
    }

    /**
     * 選択の中から破線のパスを集める
     * @param {PageItem[]} selectedItems - 選択中のアイテム
     * @returns {PathItem[]} 破線のパス
     */
    function collectDashedPaths(selectedItems) {
        var dashedPaths = [];
        /* 文字の選択中は selection が TextRange になり、[i] は undefined
           While editing text, selection is a TextRange and [i] is undefined */
        for (var i = 0; i < selectedItems.length; i++) {
            var item = selectedItems[i];
            if (item && item.typename === "PathItem" && item.stroked && item.strokeDashes.length > 0) {
                dashedPaths.push(item);
            }
        }
        return dashedPaths;
    }

    /**
     * 破線をアウトライン化し、複合パスを解除して、線分ごとのパスを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} dashedPath - 破線のパス（処理後は参照できなくなる）
     * @returns {PathItem[]} 線分ごとのパス
     */
    function outlineDashedPath(doc, dashedPath) {
        doc.selection = [dashedPath];
        app.redraw();
        app.executeMenuCommand("Live Outline Stroke");
        /* 効果として適用されてパスのまま残ったときは、アピアランスを分割して実体にする
           If it stayed a path with a live effect, expand the appearance */
        if (collectItemsByType(doc.selection, "PathItem").length > 0) {
            app.executeMenuCommand("expandStyle");
        }

        var outlinedItems = doc.selection;
        /* 線分が1つだけのときは複合パスにならない / A single dash does not become a compound path */
        var dashPieces = collectItemsByType(outlinedItems, "PathItem");
        var compoundPaths = collectItemsByType(outlinedItems, "CompoundPathItem");
        if (compoundPaths.length > 0) {
            doc.selection = compoundPaths;
            app.executeMenuCommand("noCompoundPath");
            dashPieces = dashPieces.concat(collectItemsByType(doc.selection, "PathItem"));
        }
        return dashPieces;
    }

    /**
     * 選択中のパスが単体の変換の対象になるか確かめ、なれば返す（なれなければ警告を出して null）
     * @param {PageItem[]} selectedItems - 選択中のアイテム
     * @returns {PathItem|null} 対象のパス
     */
    function getSelectedSourcePath(selectedItems) {
        if (!selectedItems || selectedItems.length !== 1) {
            alert(getLabel(LABELS.alert.selectOneItem));
            return null;
        }
        var sourcePath = selectedItems[0];
        if (!sourcePath || sourcePath.typename !== "PathItem" || !sourcePath.closed) {
            alert(getLabel(LABELS.alert.notClosedPath));
            return null;
        }
        if (sourcePath.pathPoints.length !== 4) {
            alert(getLabel(LABELS.alert.notFourAnchors));
            return null;
        }
        return sourcePath;
    }

    /**
     * パスのアンカーとハンドルの座標を読み出す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {AnchorInfo[]} アンカー情報
     */
    function readAnchors(pathItem) {
        var anchors = [];
        var pathPoints = pathItem.pathPoints;
        for (var i = 0; i < pathPoints.length; i++) {
            anchors.push({
                anchor: pathPoints[i].anchor,
                leftDirection: pathPoints[i].leftDirection,
                rightDirection: pathPoints[i].rightDirection
            });
        }
        return anchors;
    }

    /**
     * 中心線を、元の図形のすぐ前面に、元の図形の太さと塗りの色を線にしたパスとして描く
     * （ピル形状なら、中心線は両端の半円の中心を結ぶ）
     * @param {PathItem} sourcePath - 元の図形
     * @param {CenterLine} centerLine - 中心線
     * @returns {PathItem} 描いたパス
     */
    function drawCenterLine(sourcePath, centerLine) {
        var linePath = sourcePath.parent.pathItems.add();
        linePath.move(sourcePath, ElementPlacement.PLACEBEFORE);

        var anchorPositions = [];
        for (var i = 0; i < centerLine.points.length; i++) {
            anchorPositions.push(centerLine.points[i].anchor);
        }
        linePath.setEntirePath(anchorPositions);
        /* 直線ならハンドルはアンカーと同じ位置になる / Handles sit on the anchors for a straight line */
        for (var j = 0; j < centerLine.points.length; j++) {
            linePath.pathPoints[j].leftDirection = centerLine.points[j].leftDirection;
            linePath.pathPoints[j].rightDirection = centerLine.points[j].rightDirection;
        }

        linePath.closed = false;
        linePath.filled = false;
        linePath.stroked = true;
        linePath.strokeWidth = centerLine.thickness;
        /* 線端（ピル形状は丸型、ほかはなし）と実線で、元の図形と同じ長さに描く
           Round caps for pills, butt caps otherwise, and a solid line keep the original length */
        linePath.strokeCap = centerLine.roundCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
        linePath.strokeDashes = [];
        if (sourcePath.filled) linePath.strokeColor = sourcePath.fillColor;
        return linePath;
    }

    /**
     * 閉じたパスを中心線に置き換える
     * @param {PathItem} sourcePath - 元の図形
     * @param {number} [expectedThickness] - 太さがわかっていれば指定（pt）
     * @returns {PathItem|null} 描いた中心線。置き換えられなければ null（元の図形はそのまま）
     */
    function replaceWithCenterLine(sourcePath, expectedThickness) {
        if (!sourcePath.closed) return null;
        var centerLine = getCenterLine(readAnchors(sourcePath), expectedThickness);
        if (!centerLine) return null;
        var linePath = drawCenterLine(sourcePath, centerLine);
        sourcePath.remove();
        return linePath;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 破線を線分ごとの中心線に置き換える
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} dashedPaths - 破線のパス
     * @returns {void}
     */
    function convertDashedPaths(doc, dashedPaths) {
        var linePaths = [];
        var skippedCount = 0;
        var hasOutlineFailure = false;
        /* 線幅は破線ごとに違いうるので、1本ずつアウトライン化する
           Outline one path at a time, since each may have its own stroke width */
        for (var i = 0; i < dashedPaths.length; i++) {
            var strokeWidth = dashedPaths[i].strokeWidth;
            var dashPieces = outlineDashedPath(doc, dashedPaths[i]);
            if (dashPieces.length === 0) hasOutlineFailure = true;
            for (var j = 0; j < dashPieces.length; j++) {
                var linePath = replaceWithCenterLine(dashPieces[j], strokeWidth);
                if (linePath) {
                    linePaths.push(linePath);
                } else {
                    skippedCount++;
                }
            }
        }

        doc.selection = linePaths;
        if (hasOutlineFailure) {
            alert(getLabel(LABELS.alert.dashOutlineFailed));
        }
        if (skippedCount > 0) {
            alert(getLabel(LABELS.alert.skippedPieces).replace("{count}", skippedCount));
        }
    }

    /**
     * 選択した1つの図形を中心線に置き換える
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function convertSelectedShape(doc) {
        var sourcePath = getSelectedSourcePath(doc.selection);
        if (!sourcePath) return;
        var linePath = replaceWithCenterLine(sourcePath);
        doc.selection = null;
        linePath.selected = true;
    }

    /**
     * 選択が破線なら線分ごとに、そうでなければ選択した図形を中心線に置き換える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var dashedPaths = doc.selection ? collectDashedPaths(doc.selection) : [];
        if (dashedPaths.length > 0) {
            convertDashedPaths(doc, dashedPaths);
        } else {
            convertSelectedShape(doc);
        }
    }

    main();

})();
